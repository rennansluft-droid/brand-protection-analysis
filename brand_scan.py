#!/usr/bin/env python3
"""
Varredura de proteção de marca — versão genérica (configurada por config.json).

ETAPA 1 (grátis): abre cada domínio, procura a marca (com contexto, para evitar homônimos),
os modelos-chave, logo/marca na página inicial e faz a busca interna da loja
(Shopify, WooCommerce, PrestaShop e URLs de busca comuns), guardando nome, preço,
preço anterior e URL de cada produto.

ETAPA 2 (opcional, --haiku): o Claude Haiku classifica só os suspeitos com os critérios do config.

Uso:
  python brand_scan.py lista.xlsx --limit 30      # teste
  python brand_scan.py lista.xlsx                 # etapa 1
  python brand_scan.py lista.xlsx --retry-blocked # nova tentativa dos bloqueados
  python brand_scan.py lista.xlsx --haiku         # etapa 2 (ANTHROPIC_API_KEY)
"""
import argparse, base64, concurrent.futures as cf, html as htmllib, io, json, os, re, sys, threading, time
from urllib.parse import urlparse, urljoin, quote_plus

import requests, openpyxl
from openpyxl.styles import Font, PatternFill

requests.packages.urllib3.disable_warnings()
try:
    import numpy as np
    from PIL import Image, ImageFilter, ImageStat
except ImportError:  # sem Pillow/numpy: a comparação visual de logos fica desligada
    Image = None
try:
    import cairosvg
except Exception:
    cairosvg = None

CFG = json.load(open(os.environ.get("BRAND_CONFIG", "config.json"), encoding="utf-8"))
BRAND = CFG["brand"]
import unicodedata as _ud
SLUG = re.sub(r"[^a-z0-9]+", "_", "".join(c for c in _ud.normalize("NFKD", BRAND.lower()) if not _ud.combining(c))).strip("_") or "marca"
KW = [k.lower() for k in CFG["keywords"]]
KW_RE = re.compile(r"(?<![a-z0-9])(" + "|".join(map(re.escape, KW)) + r")(?![a-z0-9])", re.I)
MODELS = [m.lower() for m in CFG.get("model_terms", [])]
CONTEXT = [c.lower() for c in CFG.get("context_terms", [])] + MODELS
OFFICIAL_DOMAINS = [d.lower() for d in CFG.get("official_domains", [])]
def _words(t):
    """minúsculas, sem acentos, só letras/números separados por espaço: 'Société des Produits Nestlé S.A.' → 'societe des produits nestle s a'"""
    t = unicodedata.normalize("NFKD", str(t or "").lower())
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    return " " + " ".join(re.findall(r"[a-z0-9]+", t)) + " "


import unicodedata
OFFICIAL_REG = [r for r in CFG.get("official_registrants", []) if str(r).strip()]
OFFICIAL_REG_W = [(r, _words(r)) for r in OFFICIAL_REG]
# Registradores/gestores de domínio usados pela marca (ex.: Com Laude): domínios deles contam como oficiais
OFFICIAL_RAR_W = [(r, _words(r)) for r in CFG.get("official_registrars", []) if str(r).strip()]


def _name_match(w, field_words):
    """palavra(s) inteira(s); aceita também a forma colada ('Com Laude' = 'ComLaude')"""
    w = w.strip()
    if not w:
        return False
    if f" {w} " in field_words:
        return True
    glued = w.replace(" ", "")
    return " " in w and len(glued) >= 6 and f" {glued} " in field_words
AUTHORIZED = [d.lower() for d in CFG.get("authorized_resellers", [])]
import unicodedata


def _plain(t):
    """minúsculas, sem acentos, sem hífens/pontos/espaços (para comparar com nomes de domínio)."""
    t = unicodedata.normalize("NFKD", str(t or "").lower())
    return re.sub(r"[^a-z0-9]", "", "".join(ch for ch in t if not unicodedata.combining(ch)))


KW_PLAIN = sorted({_plain(k) for k in KW if _plain(k)}, key=len, reverse=True)


def brand_in_domain(d):
    """'nestleoficial.com', 'nestle-shop.fr', 'loja-nestlé.com' → True (sem exigir separadores)."""
    base = _plain(re.sub(r"^www\.", "", str(d or "").lower()))
    return any(k in base for k in KW_PLAIN)


OFFICIAL_RE = re.compile(CFG["official_domain_regex"], re.I) if CFG.get("official_domain_regex") else None
OFFICIAL_WORDS = [w.lower() for w in CFG.get("official_words", [
    "official", "oficial", "officiel", "officielle", "offiziell", "offizieller", "ufficiale", "officieel",
    "oficjalny", "authorized", "autorizado", "autorisé", "flagship"])]
OFFICIAL_WORD_RE = re.compile(r"(?<![a-z])(" + "|".join(map(re.escape, OFFICIAL_WORDS)) + r")(?![a-z])", re.I)
CART_RE = re.compile(r"add[ -]to[ -]cart|add to bag|ajouter au panier|adicionar ao carrinho|comprar agora|"
                     r"in den warenkorb|aggiungi al carrello|añadir al carrito|a[ñn]adir a la cesta|"
                     r"woocommerce-cart|/cart\b|/checkout\b|/panier\b|/carrinho\b|/warenkorb\b|shopping[ -]cart|"
                     r"\"@type\"\s*:\s*\"product\"|itemprop=[\"']price[\"']|product-price|price-amount", re.I)
STORE_PLATFORMS = ("shopify", "woocommerce", "prestashop", "magento")

MIN_OFF = int(CFG.get("min_discount", 30))
PRICE_ALERT = float(CFG.get("price_alert", 0) or 0)
TERM = CFG.get("search_term") or KW[0]
PRICE_TERMS = [t.lower() for t in CFG.get("price_alert_terms", [])] or CONTEXT
PRODUCT_CTX = CONTEXT + [t.lower() for t in CFG.get("product_terms", [])]
SUSP_REG = [r.lower() for r in CFG.get("suspicious_registrars", [])]
RECENT = CFG.get("recent_from", "2025-01-01")
SALE_RE = re.compile(r"\b(" + "|".join(CFG.get("sale_words", ["sale", "outlet", "soldes", "promo"])) + r")\b", re.I)
HAIKU_MODEL = CFG.get("haiku_model", "claude-haiku-5-5")
LOGO_SIM = float(CFG.get("logo_similarity", 0.88))      # >= isto: logo oficial
LOGO_NEAR = float(CFG.get("logo_near", 0.75))           # >= isto: parecido (mostrar para revisão)

CACHE_FILE = f"cache_{SLUG}.jsonl"
OUT_FILE = f"{SLUG}_resultados.xlsx"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/130.0 Safari/537.36")
MAX_BYTES = 2_000_000
PARKED = ("domain is for sale", "this domain may be for sale", "buy this domain", "domain for sale",
          "sedoparking", "parkingcrew", "afternic", "hugedomains", "dan.com", "bodis.com",
          "this domain is parked", "domain parking", "le domaine est à vendre", "ce domaine est à vendre",
          "domain name is for sale", "godaddy.com/forsale", "parklogic")
CHALLENGE = ("cf-chl", "just a moment...", "attention required", "captcha", "access denied",
             "ddos-guard", "checking your browser")
NO_RESULT = ("no results", "no products found", "nothing found", "0 results", "aucun résultat",
             "aucun produit", "keine ergebnisse", "nessun risultato", "nenhum resultado",
             "no se encontraron", "sin resultados", "geen resultaten", "0 résultat")
RETRY = {"BLOCKED", "ERROR_403", "ERROR_429", "ERROR_503", "ERROR_999", "ERROR_520", "ERROR_521",
         "ERROR_401", "ERROR_405", "ERROR_409", "ERROR_418", "ERROR_402", "ERROR"}
HTML_SEARCH = ("/search?q={t}", "/?s={t}&post_type=product", "/recherche?controller=search&s={t}",
               "/fr/recherche?controller=search&s={t}", "/catalogsearch/result/?q={t}",
               "/search?query={t}")

# --------------------------------------------------------------------------- HTML helpers

def visible_text(h):
    h = re.sub(r"(?is)<(script|style|noscript|svg)\b.*?</\1>", " ", h)
    h = re.sub(r"(?s)<!--.*?-->", " ", h)
    h = re.sub(r"(?s)<[^>]+>", " ", h)
    return re.sub(r"\s+", " ", htmllib.unescape(h)).strip()


def title_of(h):
    m = re.search(r"(?is)<title[^>]*>(.*?)</title>", h)
    return re.sub(r"\s+", " ", htmllib.unescape(m.group(1))).strip()[:160] if m else ""


def lang_of(h):
    m = re.search(r'(?is)<html[^>]*\blang=["\']?([a-z]{2})', h)
    return m.group(1).lower() if m else ""


def has_context(text):
    low = text.lower()
    return any(c in low for c in CONTEXT)


def brand_mentions(text):
    """(total, com contexto): menções da marca e quantas têm termo de contexto por perto."""
    low, tot, ctx = text.lower(), 0, 0
    for m in KW_RE.finditer(low):
        tot += 1
        if any(c in low[max(0, m.start() - 250): m.end() + 250] for c in CONTEXT):
            ctx += 1
    return tot, ctx


def model_hits(text):
    low = text.lower()
    return sorted({m for m in MODELS if m in low})


def snippets(text, n=8, width=180):
    out, low = [], text.lower()
    for m in KW_RE.finditer(low):
        if len(out) >= n:
            break
        if not out or m.start() > out[-1][1]:
            out.append((text[max(0, m.start() - width): m.end() + width], m.end() + width))
    for t in MODELS:
        i = low.find(t)
        if i >= 0 and len(out) < n + 4:
            out.append((text[max(0, i - width): i + width], 0))
    return [s for s, _ in out]


def brand_tags(h):
    hits = []
    for tag in re.findall(r"(?is)<(?:img|source|link|meta)\b[^>]*>", h):
        if KW_RE.search(tag):
            hits.append(re.sub(r"\s+", " ", tag)[:200])
    hits.sort(key=lambda t: 0 if is_logo(t) else 1)
    return hits[:6]


def is_logo(tag):
    t = tag.lower()
    return ("logo" in t or "og:site_name" in t or "rel=\"icon\"" in t or "shortcut icon" in t
            or "apple-touch-icon" in t)


def product_links(h, base, limit=15):
    """Links cujo endereço ou texto cita a marca/modelos (útil em lojas sem API)."""
    out, seen = [], set()
    for m in re.finditer(r'(?is)<a\b[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', h):
        href, txt = m.group(1), m.group(2)
        t = visible_text(txt)[:100]
        hay = (href + " " + t).lower()
        if (KW_RE.search(hay) or any(m.replace(" ", "-") in hay or m in hay for m in MODELS)) \
                and not href.startswith(("#", "javascript", "mailto")):
            url = urljoin(base + "/", href).split("#")[0]
            if url not in seen and "search" not in url and "recherche" not in url and "?s=" not in url:
                seen.add(url)
                seg = h[m.end(): m.end() + 600]
                nxt = re.search(r"(?i)<a\b", seg)
                after = visible_text(seg[: nxt.start()] if nxt else seg)[:160]
                pm = [x or y for x, y in PRICE_RE.findall(after)]
                vals = [float(x.replace(" ", "").replace(".", "").replace(",", ".")) if re.search(r",\d{2}$", x)
                        else float(x.replace(",", "")) for x in pm[:2]]
                price, was = (min(vals), max(vals)) if len(vals) == 2 and vals[0] != vals[1] else ((vals[0], None) if vals else (None, None))
                out.append({"title": t or url.rsplit("/", 1)[-1], "price": price, "was": was,
                            "off": pct(was, price), "url": url, "cur": "EUR" if pm else ""})
        if len(out) >= limit:
            break
    return out


# --------------------------------------------------------------------------- logo (visual)

def load_image(data):
    if Image is None or not data:
        return None
    head = data[:600].lstrip().lower()
    if head.startswith(b"<svg") or (head.startswith(b"<?xml") and b"<svg" in data[:3000]):
        if not cairosvg:
            return None
        try:
            data = cairosvg.svg2png(bytestring=data, output_width=600)
        except Exception:
            return None
    try:
        im = Image.open(io.BytesIO(data))
        if im.width * im.height > 40_000_000:      # imagem gigante: ignora (protege a memória)
            return None
        im.seek(0)
        im.draft("RGB", (1200, 1200))             # JPEG: decodifica já reduzido
        if max(im.size) > 1200:
            im.thumbnail((1200, 1200))
        return im.convert("RGBA")
    except Exception:
        return None


def normalize_logo(im):
    """Fundo claro, traço escuro, bordas cortadas (funciona com logo branco/preto/colorido)."""
    if im is None or im.width < 12 or im.height < 8:
        return None
    best = None
    for bgc in ((255, 255, 255, 255), (0, 0, 0, 255)):
        bg = Image.new("RGBA", im.size, bgc)
        bg.alpha_composite(im)
        g = bg.convert("L")
        sd = ImageStat.Stat(g).stddev[0]
        if best is None or sd > best[0] + 1:
            best = (sd, g)
    if best[0] < 8:
        return None
    a = np.asarray(best[1], dtype=np.float32)
    border = np.concatenate([a[0], a[-1], a[:, 0], a[:, -1]])
    if np.median(border) < 128:
        a, border = 255 - a, 255 - border
    ys, xs = np.where(a < np.median(border) - 40)
    if len(xs) < 10:
        return None
    y0, y1 = max(ys.min() - 1, 0), min(ys.max() + 2, a.shape[0])
    x0, x1 = max(xs.min() - 1, 0), min(xs.max() + 2, a.shape[1])
    return Image.fromarray(a[y0:y1, x0:x1].astype(np.uint8), "L")


def _vec(g, size):
    v = np.asarray(g.resize(size, Image.LANCZOS).filter(ImageFilter.GaussianBlur(1.2)), dtype=np.float32).ravel()
    v = v - v.mean()
    n = np.linalg.norm(v)
    return v / n if n else v


def logo_similarity(off, g):
    asp_o, asp_c = off.width / off.height, g.width / g.height
    if not (1 / 1.35 <= asp_c / asp_o <= 1.35):
        return 0.0
    size = (128, max(8, min(128, round(128 / asp_o))))
    return max(0.0, float(np.dot(_vec(off, size), _vec(g, size))))


def thumb_b64(im, max_w=220, max_h=80):
    im = im.copy()
    im.thumbnail((max_w, max_h))
    bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
    bg.alpha_composite(im.convert("RGBA"))
    buf = io.BytesIO()
    bg.convert("RGB").save(buf, "PNG", optimize=True)
    return base64.b64encode(buf.getvalue()).decode()


def _load_official():
    out = []
    if Image is None:
        return out
    sources = [("b64", x) for x in CFG.get("official_logos_b64", [])] + \
              [("file", x) for x in CFG.get("official_logos", [])] + [("url", x) for x in CFG.get("official_logo_urls", [])]
    for kind, src in sources:
        try:
            data = (base64.b64decode(src) if kind == "b64" else
                    requests.get(src, timeout=15, headers={"User-Agent": UA}, verify=False).content if kind == "url"
                    else open(src, "rb").read())
        except Exception:
            continue
        im = load_image(data)
        g = normalize_logo(im)
        if g is not None:
            out.append({"src": src if kind != "b64" else "configuração", "g": g, "thumb": thumb_b64(im)})
    return out


OFFICIAL_LOGOS = _load_official()
LOGO_SEM = threading.BoundedSemaphore(4)   # no máximo 4 imagens sendo processadas ao mesmo tempo (memória)
LOGO_WORDS = ("logo", "brand", "site-title", "navbar-brand", "custom-logo", "site-logo", "header__heading")


def _img_src(tag):
    for attr in ("src", "data-src", "data-lazy-src", "data-original", "srcset", "data-srcset", "href", "content"):
        m = re.search(attr + r'\s*=\s*["\']([^"\']+)["\']', tag, re.I)
        if m:
            v = m.group(1).strip().split(",")[0].strip().split(" ")[0]
            if v and not v.startswith("data:"):
                return v
    return ""


def top_area(h):
    m = re.search(r"(?is)<header\b.*?</header>", h)
    body = re.search(r"(?is)<body\b", h)
    return m.group(0) if m else h[body.start(): body.start() + 25000] if body else h[:25000]


def store_signals(r, h):
    """Loja (vende produtos) + marca no cabeçalho/domínio + se diz 'oficial'."""
    top_txt = visible_text(top_area(h))[:4000]
    sn = re.search(r'(?is)<meta[^>]*property=["\']og:site_name["\'][^>]*content=["\']([^"\']*)', h)
    site_name = htmllib.unescape(sn.group(1)) if sn else ""
    title = r.get("title") or ""
    r["brand_header"] = bool(KW_RE.search(top_txt) or KW_RE.search(title) or KW_RE.search(site_name)
                             or any(KW_RE.search(c.get("alt", "") + " " + c.get("url", "").rsplit("/", 1)[-1])
                                    for c in r.get("logo_candidates", []) if c.get("where") == "cabeçalho"))
    claims = []
    dom = re.sub(r"[-_.]", " ", r["domain"])
    for where, txt in (("domínio", dom), ("título", title), ("nome do site", site_name), ("cabeçalho", top_txt)):
        m = OFFICIAL_WORD_RE.search(txt) or (where == "domínio" and any(w in _plain(r["domain"]) for w in
                                                                         ("official", "oficial", "officiel", "offiziell")) and True)
        if m:
            claims.append(f'"{m.group(0) if hasattr(m, "group") else "oficial"}" no {where}')
    r["official_claim"] = claims
    low = h.lower()
    r["ecommerce"] = bool(r.get("products") or r.get("platform") in STORE_PLATFORMS or CART_RE.search(low)
                          or len(PRICE_RE.findall(visible_text(h)[:20000])) >= 3)


def logo_candidates(h, base):
    """Imagens que podem ser o logo do próprio site, com a posição: cabeçalho, página, ícone."""
    top = top_area(h)
    out, seen = [], set()

    def add(tag, where):
        src = _img_src(tag)
        if not src:
            return
        url = urljoin(base + "/", htmllib.unescape(src))
        if url in seen or len(out) >= 6:
            return
        seen.add(url)
        alt = re.search(r'alt\s*=\s*["\']([^"\']*)', tag, re.I)
        out.append({"url": url, "where": where, "alt": htmllib.unescape(alt.group(1))[:80] if alt else "",
                    "tag": re.sub(r"\s+", " ", tag)[:200]})

    for tag in re.findall(r"(?is)<img\b[^>]*>", top):  # logo do cabeçalho primeiro
        if any(w in tag.lower() for w in LOGO_WORDS):
            add(tag, "cabeçalho")
    for tag in re.findall(r"(?is)<img\b[^>]*>", top)[:3]:
        add(tag, "cabeçalho")
    for tag in re.findall(r"(?is)<img\b[^>]*>", h):
        t = tag.lower()
        if "logo" in t and (KW_RE.search(t) or not out):
            add(tag, "página")
    for tag in re.findall(r"(?is)<link\b[^>]*rel=[\"'][^\"']*(?:apple-touch-icon|icon)[^\"']*[\"'][^>]*>", h)[:2]:
        add(tag, "ícone")
    for tag in re.findall(r"(?is)<meta\b[^>]*property=[\"']og:image[\"'][^>]*>", h)[:1]:
        add(tag, "og:image")
    return out


def identity(h, title):
    """Sinais de texto de que o site se apresenta como a marca."""
    sig = []
    sn = re.search(r'(?is)<meta[^>]*property=["\']og:site_name["\'][^>]*content=["\']([^"\']*)', h) or \
         re.search(r'(?is)<meta[^>]*name=["\']application-name["\'][^>]*content=["\']([^"\']*)', h)
    if sn and KW_RE.search(sn.group(1)):
        sig.append(f'nome do site: "{htmllib.unescape(sn.group(1))[:60]}"')
    if KW_RE.match(title.strip().lower()):
        sig.append(f'título começa com a marca: "{title[:60]}"')
    return sig


def check_logos(s, r, h, base, max_imgs=5):
    """Preenche r['logo_candidates'], r['identity'], r['logo_match'] (melhor semelhança com o logo oficial)."""
    cands = logo_candidates(h, base)
    r["identity"] = identity(h, r.get("title", ""))
    for c in cands:
        if c["where"] == "cabeçalho" and (KW_RE.search(c["alt"]) or KW_RE.search(c["url"].rsplit("/", 1)[-1])):
            r["identity"].append(f'logo do cabeçalho com o nome da marca: {c["url"].rsplit("/", 1)[-1][:60]}')
            break
    r["logo_candidates"] = [{k: c[k] for k in ("url", "where", "alt")} for c in cands]
    r["logo_match"] = {}
    if not OFFICIAL_LOGOS:
        return
    best = None
    for c in [c for c in cands if max_imgs >= 5 or c["where"] == "cabeçalho"][:max_imgs]:
        try:
            resp = s.get(c["url"], timeout=8, verify=False, headers={"User-Agent": UA}, stream=True)
            data = resp.raw.read(1_500_000, decode_content=True) if resp.status_code == 200 else b""
            resp.close()
        except Exception:
            continue
        with LOGO_SEM:
            im = load_image(data)
            del data
            g = normalize_logo(im)
            if g is None:
                continue
            sim = max(logo_similarity(o["g"], g) for o in OFFICIAL_LOGOS)
            if best is not None and sim <= best[0]:
                continue
            im.thumbnail((440, 160))  # guarda só uma versão pequena
        if best is None or sim > best[0]:
            best = (sim, c, im)
        if sim >= 0.97:
            break
    if best and best[0] >= LOGO_NEAR:
        r["logo_match"] = {"sim": round(best[0], 3), "url": best[1]["url"], "where": best[1]["where"],
                           "thumb": thumb_b64(best[2])}


def logo_pass(row, factory=None):
    """Só a parte de logos, para domínios já no cache (sem refazer a busca de produtos)."""
    r = dict(row)
    s = (factory or requests.Session)()
    url = r.get("final_url") or "https://" + r["domain"]
    code, url, body = fetch(s, url)
    if code == 200:
        base = f"{urlparse(url).scheme}://{urlparse(url).netloc}"
        brand_here = r.get("products") or brand_in_domain(r["domain"]) or KW_RE.search(body)
        check_logos(s, r, body, base, 5 if brand_here else 1)
        store_signals(r, body)
    r["logos_checked"] = True
    return r


PRICE_RE = re.compile(r"(?:€|eur|\$|£)\s?(\d{1,5}(?:[.,]\d{2})?)|(\d{1,5}(?:[ .]\d{3})*(?:[.,]\d{2})?)\s?(?:€|eur\b)", re.I)


def platform_of(low):
    if "cdn.shopify.com" in low or "shopify.theme" in low:
        return "shopify"
    if "woocommerce" in low:
        return "woocommerce"
    if "wp-content" in low:
        return "wordpress"
    if "prestashop" in low or "/modules/ps_" in low:
        return "prestashop"
    if "mage/cookies" in low or "magento" in low:
        return "magento"
    if "wixstatic.com" in low:
        return "wix"
    return ""

# --------------------------------------------------------------------------- network

def fetch(s, url, timeout=12):
    try:
        r = s.get(url, timeout=timeout, stream=True, allow_redirects=True, verify=False,
                  headers={"User-Agent": UA, "Accept-Language": "fr,en;q=0.8"})
        chunks, size = [], 0
        for c in r.iter_content(65536):
            chunks.append(c)
            size += len(c)
            if size > MAX_BYTES:
                break
        r.close()
        raw = b"".join(chunks)
        enc = r.encoding if "charset" in r.headers.get("content-type", "").lower() else None
        if not enc:
            m = re.search(rb'charset=["\']?([\w-]+)', raw[:4096], re.I)
            enc = m.group(1).decode() if m else "utf-8"
        try:
            text = raw.decode(enc, errors="replace")
        except LookupError:
            text = raw.decode("utf-8", errors="replace")
        return r.status_code, r.url, text
    except Exception as e:
        return None, url, f"{type(e).__name__}: {str(e)[:150]}"


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def pct(was, price):
    return round(100 * (was - price) / was) if was and price and was > price > 0 else 0


def parse_shopify(body, base):
    prods = json.loads(body).get("resources", {}).get("results", {}).get("products", [])
    out = []
    for p in prods:
        price = num(p.get("price") or p.get("price_min"))
        was = num(p.get("compare_at_price_max") or p.get("compare_at_price_min"))
        url = (p.get("url") or "").split("?")[0]
        out.append({"title": (p.get("title", "") + (" · " + p["vendor"] if p.get("vendor") else ""))[:110],
                    "price": price, "was": was if was and price and was > price else None,
                    "off": pct(was, price), "url": base + url if url.startswith("/") else url, "cur": ""})
    return out


def parse_woo(body):
    data = json.loads(body)
    if not isinstance(data, list):
        return []
    out = []
    for p in data:
        pr = p.get("prices", {})
        unit = 10 ** int(pr.get("currency_minor_unit", 2) or 0)
        price = (num(pr.get("price")) or 0) / unit or None
        was = (num(pr.get("regular_price")) or 0) / unit or None
        out.append({"title": htmllib.unescape(p.get("name", ""))[:110], "price": price,
                    "was": was if was and price and was > price else None, "off": pct(was, price),
                    "url": p.get("permalink", ""), "cur": pr.get("currency_code", "")})
    return out


def search_store(s, base, platform):
    t = quote_plus(TERM)
    tries = []
    if platform in ("shopify", ""):
        tries.append(("shopify", f"{base}/search/suggest.json?q={t}&resources[type]=product&resources[limit]=10",
                      f"{base}/search?q={t}&type=product"))
    if platform in ("woocommerce", "wordpress", ""):
        tries.append(("woo", f"{base}/wp-json/wc/store/v1/products?search={t}&per_page=50",
                      f"{base}/?s={t}&post_type=product"))
    for kind, url, page in tries:
        code, _, body = fetch(s, url)
        if code == 200 and body.lstrip()[:1] in "[{":
            try:
                prods = parse_shopify(body, base) if kind == "shopify" else parse_woo(body)
                return page, prods, 0, ""
            except (ValueError, AttributeError):
                pass
    for path in HTML_SEARCH:
        url = base + path.format(t=t)
        code, final, body = fetch(s, url)
        if code != 200 or not body:
            continue
        text = visible_text(body)
        low = text.lower()
        if any(m in low for m in NO_RESULT):
            continue
        tot, ctx = brand_mentions(text)
        hits = max(0, tot - 2) + len(model_hits(text))
        if hits >= 2:
            return url, product_links(body, base), hits, text
    return "", [], 0, ""


def relevant(p):
    """Produto da marca: modelo-chave no nome/URL, ou marca + termo de produto (evita 'capinha Gertrude',
    'livro de Gertrude Stein'), ou marca como fabricante (Shopify: '… · Gertrude')."""
    title = (p.get("title") or "").lower()
    t = title + " " + (p.get("url") or "").lower().replace("-", " ")
    if any(m in t for m in MODELS):
        return True
    vendor = title.rsplit(" · ", 1)[1] if " · " in title else ""
    return bool(KW_RE.search(t)) and (bool(KW_RE.search(vendor)) or any(c in t for c in PRODUCT_CTX))

# --------------------------------------------------------------------------- stage 1

def scan(row, factory=None):
    d = row["domain"]
    res = {"domain": d, "registrant": row.get("registrant", ""), "registrar": row.get("registrar", ""),
           "created": row.get("created", ""), "status": "", "final_url": "", "redirect_to": "",
           "title": "", "lang": "", "platform": "", "mentions": 0, "mentions_ctx": 0, "models": [],
           "brand_tags": [], "search_page": "", "search_hits": 0, "products": [], "snippets": [],
           "note": ""}
    s = (factory or requests.Session)()
    code, url, body = None, "", ""
    for scheme in ("https://", "http://"):
        code, url, body = fetch(s, scheme + d)
        if code:
            break
    if not code:
        res.update(status="DEAD", note=body)
        return res
    low = body.lower()
    host = (urlparse(url).hostname or "").lower()
    res.update(final_url=url, title=title_of(body), lang=lang_of(body))
    if host and not host.endswith(d):
        res["redirect_to"] = host
    if code in (401, 403, 429, 503) and any(m in low for m in CHALLENGE):
        res["status"] = "BLOCKED"
        return res
    if code >= 400:
        res["status"] = f"ERROR_{code}"
        return res
    text = visible_text(body)
    if any(m in low for m in PARKED) or len(text) < 80:
        res["status"] = "PARKED"
        return res
    res["status"] = "OK"
    res["platform"] = platform_of(low)
    res["mentions"], res["mentions_ctx"] = brand_mentions(text + " " + res["title"])
    res["models"] = model_hits(text)
    res["brand_tags"] = brand_tags(body)
    res["snippets"] = snippets(text)
    base = f"{urlparse(url).scheme}://{urlparse(url).netloc}"
    home_links = product_links(body, base)
    page, prods, hits, stext = search_store(s, base, res["platform"])
    res.update(search_page=page, search_hits=hits)
    if stext:
        res["snippets"] += snippets(stext, n=6)
        res["models"] = sorted(set(res["models"]) | set(model_hits(stext)))
    allp = prods + [p for p in home_links if p["url"] not in {q["url"] for q in prods}]
    res["products"] = [p for p in allp if relevant(p)][:25]
    # marca em algum lugar do HTML: analisa até 5 imagens; senão, só o logo do cabeçalho (cópia renomeada)
    check_logos(s, res, body, base, 5 if (res["products"] or brand_in_domain(d) or KW_RE.search(low)) else 1)
    store_signals(res, body)
    res["logos_checked"] = True
    res["models"] = sorted(set(res["models"]) | set(model_hits(" ".join(p["title"] + " " + p["url"] for p in res["products"]).replace("-", " "))))
    return res

# --------------------------------------------------------------------------- classification

def official_reason(r):
    """Por que o domínio é oficial ('' = não é): site oficial, padrão de domínio ou titular oficial."""
    host = (urlparse(r.get("final_url") or "").hostname or r["domain"]).lower()
    for o in OFFICIAL_DOMAINS:
        if host == o or host.endswith("." + o) or r["domain"] == o or r["domain"].endswith("." + o):
            return f"site oficial ({o})"
    if OFFICIAL_RE and (OFFICIAL_RE.search(host) or OFFICIAL_RE.search(r["domain"])):
        return "domínio oficial da marca"
    reg, rar = _words(r.get("registrant")), _words(r.get("registrar"))
    for name, w in OFFICIAL_REG_W:   # palavra inteira: 'Laude' não casa com 'Claude'
        if _name_match(w, reg):
            return f"titular oficial: {name}"
    for name, w in OFFICIAL_RAR_W:   # gestor de domínios da marca, no Registrar ou no Registrant
        if _name_match(w, rar) or _name_match(w, reg):
            return f"registrado via {name}"
    return ""


def is_official(r):
    return bool(official_reason(r))


def network(r):
    sig = []
    reg = (r.get("registrar") or "").lower()
    created = str(r.get("created") or "")[:10]
    if any(x in reg for x in SUSP_REG) and created >= RECENT:
        sig.append("registrador suspeito + domínio recente")
    elif created >= RECENT:
        sig.append(f"domínio criado em {created}")
    if SALE_RE.search(r.get("title") or "") or SALE_RE.search(r["domain"]):
        sig.append("promoção no título/domínio")
    if brand_in_domain(r["domain"]):
        sig.append("marca no domínio")
    if is_fake_official_store(r):
        sig.append("loja que se diz oficial (" + ", ".join(r["official_claim"][:2]) + ")")
    score = sum(2 if ("registrador" in s or "marca no" in s or "oficial" in s) else 1 for s in sig)
    return score, sig


def selling_brand(r):
    """Vende produtos E mostra a marca no cabeçalho ou no domínio."""
    return bool(r.get("ecommerce")) and (bool(r.get("brand_header")) or brand_in_domain(r["domain"]))


def is_fake_official_store(r):
    """Se diz 'oficial' + vende + marca no cabeçalho/domínio, e NÃO está na lista de domínios oficiais."""
    return selling_brand(r) and bool(r.get("official_claim")) and not is_official(r)


def low_price(p):
    path = urlparse(p.get("url") or "").path
    return bool(PRICE_ALERT and p["price"] and 0 < p["price"] < PRICE_ALERT
                and any(t in (p["title"] + " " + path).lower().replace("-", " ") for t in PRICE_TERMS))


def models_of(r):
    return [m for m in r.get("models", []) if m in MODELS]


def flags(r):
    prods = [p for p in r.get("products", []) if relevant(p)]
    A_items = [p for p in prods if (p["off"] or 0) >= MIN_OFF or low_price(p)]
    brand_title = bool(KW_RE.search(r.get("title") or ""))
    logo = any(is_logo(t) for t in r.get("brand_tags", []))
    vis = (r.get("logo_match") or {}).get("sim", 0) >= LOGO_SIM
    rel = r.get("mentions_ctx", 0) or prods or models_of(r) or vis or selling_brand(r)
    B = bool(rel) and (brand_title or logo or vis or bool(r.get("identity")) or brand_in_domain(r["domain"]))
    C = bool(models_of(r))
    return bool(A_items), B, C, A_items, bool(rel)


def classify(r):
    """(prioridade, rótulo). Prioridade 1–4 = suspeito; None = fora da lista de suspeitos."""
    if r.get("status") != "OK":
        return None, r.get("status", "?")
    if is_official(r):
        return None, "OFICIAL"
    if any(r["domain"] == a or r["domain"].endswith("." + a) for a in AUTHORIZED):
        return None, "REVENDEDOR AUTORIZADO"
    A, B, C, A_items, rel = flags(r)
    lm = r.get("logo_match") or {}
    vis = lm.get("sim", 0) >= LOGO_SIM
    own_logo = vis and lm.get("where") in ("cabeçalho", "ícone", "og:image")  # usa o logo oficial como o seu
    ident = bool(r.get("identity"))
    if not rel:
        return None, "HOMÔNIMO PROVÁVEL" if r.get("mentions") or brand_in_domain(r["domain"]) else "SEM MENÇÃO"
    net, _ = network(r)
    fake_off = is_fake_official_store(r)
    store_dom = selling_brand(r) and brand_in_domain(r["domain"])
    label = "+".join(x for x, f in (("A", A), ("B", B), ("C", C), ("OF", fake_off)) if f) or "MENÇÃO"
    if fake_off:  # "loja oficial" que não está na lista oficial: nunca descartar
        return 1, label
    A_price = any(low_price(p) for p in A_items)   # preço de falsificação
    if (A_price or (A and (B or net >= 2)) or (B and net >= 3) or own_logo
            or (ident and (C or A or net >= 1))):
        return 1, label
    if A or ((B or C) and net >= 2) or ident or store_dom:
        return 2, label
    if net >= 2:
        return 3, label
    if B or C or vis or selling_brand(r):
        return 4, label  # pode ser revendedor: verificar autorização
    return None, "MENÇÃO"

# --------------------------------------------------------------------------- stage 2 (Haiku)

HAIKU_SYSTEM = """You are a brand-protection analyst for {brand} (official sites: {official}).
You receive data already extracted from ONE website. Use ONLY this data. Criteria:
{criteria}
Authorized resellers or unrelated businesses that merely share the name are NOT suspects.
Calling itself "official"/"oficial"/"officiel"/"authorized" is NOT evidence of legitimacy: only the official sites listed above are official.
A site that sells products, shows the brand in its header or domain and claims to be official, but is not in the official list, is a strong suspect.
Do not guess; if the data does not show it, answer false/null. Do not conclude "fraud"; list observable red flags.
Return ONLY this JSON:
{{"flags":"e.g. A+C or none","evidence":"short, cite product names/prices/URLs","red_flags":[],"confidence":"high|medium|low"}}"""


def ask_haiku(client, r):
    ev = {k: r.get(k) for k in ("domain", "final_url", "redirect_to", "title", "lang", "platform",
                                "mentions", "mentions_ctx", "models", "brand_tags", "search_page",
                                "registrar", "created")}
    ev["products"] = r.get("products", [])[:12]
    ev["snippets"] = [s[:360] for s in r.get("snippets", [])[:10]]
    msg = client.messages.create(
        model=HAIKU_MODEL, max_tokens=400,
        system=HAIKU_SYSTEM.format(brand=BRAND, official=", ".join(OFFICIAL_DOMAINS), criteria=CFG.get("criteria", "")),
        messages=[{"role": "user", "content": json.dumps(ev, ensure_ascii=False)[:8000]}])
    txt = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
    m = re.search(r"\{.*\}", txt, re.S)
    out = json.loads(m.group(0)) if m else {"flags": "?", "evidence": txt[:200]}
    out["_tokens"] = [msg.usage.input_tokens, msg.usage.output_tokens]
    return out

# --------------------------------------------------------------------------- I/O

def read_domains(path):
    ws = openpyxl.load_workbook(path, read_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    head = [str(h or "").strip().lower() for h in rows[0]]

    def col(*names):
        for n in names:                      # nome exato primeiro, depois "contém"
            if n in head:
                return head.index(n)
        for i, h in enumerate(head):
            if any(n in h for n in names):
                return i
        return None
    di = col("domain", "domínio", "dominio", "domaine")
    di = 0 if di is None else di
    ci = {"registrant": col("registrant", "titular", "owner", "holder", "titulaire"),
          "registrar": col("registrar", "registrador", "bureau d'enregistrement"),
          "created": col("created", "creation", "criado", "criação", "création", "registered on")}
    if ci["registrant"] == ci["registrar"]:
        ci["registrant"] = None
    out, seen = [], set()
    for r in rows[1:]:
        d = str(r[di] or "").strip().lower()
        d = re.sub(r"^https?://", "", d).split("/")[0].removeprefix("www.")
        if d and d not in seen:
            seen.add(d)
            out.append({"domain": d, **{k: (str(r[i])[:10] if k == "created" else str(r[i]))
                                        if i is not None and r[i] else "" for k, i in ci.items()}})
    return out


def load_cache():
    cache = {}
    if os.path.exists(CACHE_FILE):
        for line in open(CACHE_FILE, encoding="utf-8"):
            try:
                r = json.loads(line)
                cache[r["domain"]] = r
            except (ValueError, KeyError):
                pass
    return cache


def link(url, label):
    return f'=HYPERLINK("{url[:250]}","{label}")' if url else ""


def write_excel(results):
    wb = openpyxl.Workbook()
    head_fill = PatternFill("solid", fgColor="1F3864")
    fills = {1: "F8CBAD", 2: "FCE4D6", 3: "FFF2CC", 4: "E2EFDA"}

    def style(ws, widths):
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = head_fill
        ws.freeze_panes = "D2"
        ws.auto_filter.ref = ws.dimensions
        for i, w in enumerate(widths):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w

    sus, prows, blocked, other, logos = [], [], [], [], []
    for r in results:
        pr, label = classify(r)
        r["_class"] = (pr, label)
        lm = r.get("logo_match") or {}
        if r.get("status") == "OK" and not is_official(r) and (lm or (pr and r.get("identity"))):
            logos.append({"prio": pr, "label": label, "domain": r["domain"], "site": r.get("final_url") or "https://" + r["domain"],
                          "where": lm.get("where", ""), "sim": lm.get("sim"), "thumb": lm.get("thumb", ""),
                          "img": lm.get("url", ""), "identity": r.get("identity", []), "title": r.get("title", "")[:120]})
        d = r["domain"]
        g = f"https://www.google.com/search?q=site%3A{d}+{quote_plus(TERM)}"
        if pr:
            A, B, C, A_items, _ = flags(r)
            net, sig = network(r)
            h = r.get("haiku") or {}
            ev_a = "; ".join(f'{p["title"][:50]} {p["price"]:g}' + (f' (antes {p["was"]:g}, -{p["off"]}%)' if p["off"] else "")
                             for p in sorted(A_items, key=lambda p: (-(p["off"] or 0), p["price"] or 0))[:3])
            prices = [p["price"] for p in r["products"] if p["price"] and relevant(p)]
            logo = next((t for t in r["brand_tags"] if is_logo(t)), "")
            lm = r.get("logo_match") or {}
            logo_txt = (f'logo oficial {lm["sim"]*100:.0f}% parecido ({lm["where"]})' if lm else "")
            logo_txt = "; ".join(x for x in [logo_txt] + r.get("identity", []) if x)
            sus.append([pr, label, d, link("https://" + d, "abrir"), link(r["search_page"], "ver produtos"),
                        link(g, "Google"), r["title"][:120], ev_a,
                        max([p["off"] for p in r["products"] if relevant(p)] or [0]) or None,
                        min(prices) if prices else None, ", ".join(models_of(r)),
                        logo_txt or logo[:150] or ("marca no título" if KW_RE.search(r["title"]) else ""),
                        ", ".join(sig), r.get("lang", ""), r.get("registrar", ""), r.get("created", ""),
                        h.get("flags", ""), h.get("evidence", ""), ", ".join(h.get("red_flags") or []), ""])
            for p in [p for p in r["products"] if relevant(p)]:
                prows.append([pr, label, d, p["title"], p["price"], p["was"], p["off"] or None, p.get("cur", ""),
                              link(p["url"], p["url"][:90])])
        elif r.get("status") in RETRY and not is_official(r):
            score, why = 0, []
            reg = (r.get("registrar") or "").lower()
            created = str(r.get("created") or "")[:10]
            if brand_in_domain(d): score += 3; why.append("marca no domínio")
            if any(w in _plain(d) for w in ("official", "oficial", "officiel", "offiziell", "store", "shop", "loja", "boutique")):
                score += 1; why.append("'oficial'/'loja' no domínio")
            if any(x in reg for x in SUSP_REG) and created >= RECENT: score += 2; why.append("registrador suspeito + recente")
            elif created >= RECENT: why.append("criado recentemente")
            if SALE_RE.search(d): score += 1; why.append("promoção no domínio")
            prio = "Alta" if score >= 3 else "Média" if score >= 1 else "Baixa"
            blocked.append([prio, d, link("https://" + d, "abrir"), link(f"https://{d}/search?q={quote_plus(TERM)}", "buscar"),
                            link(g, "Google"), ", ".join(why), r.get("registrar", ""), created, r["status"], ""])
        elif label in ("OFICIAL", "HOMÔNIMO PROVÁVEL", "REVENDEDOR AUTORIZADO") or is_official(r):
            why = f"OFICIAL — {official_reason(r)}" if is_official(r) else label
            other.append([d, why, r.get("title", "")[:120], (r.get("registrant") or "")[:80]])

    sus.sort(key=lambda x: (x[0], -(x[8] or 0), x[2]))
    prows.sort(key=lambda x: (x[0], x[2], -(x[6] or 0)))
    blocked.sort(key=lambda x: ({"Alta": 0, "Média": 1, "Baixa": 2}[x[0]], x[1]))

    ws = wb.active
    ws.title = "Suspeitos"
    ws.append(["prioridade", "critérios", "domínio", "site", "produtos", "Google", "título do site",
               "produtos com desconto/preço baixo (top 3)", "desconto máx %", "menor preço", "modelos-chave achados",
               "marca/logo na home", "sinais de loja falsa", "idioma", "registrador", "criado em",
               "haiku", "evidência haiku", "red flags haiku", "sua revisão"])
    for x in sus:
        ws.append(x)
    style(ws, [10, 10, 30, 7, 12, 8, 45, 60, 10, 10, 30, 40, 40, 7, 22, 11, 8, 40, 30, 14])
    for row in ws.iter_rows(min_row=2):
        for c in row[:3]:
            c.fill = PatternFill("solid", fgColor=fills[row[0].value])

    logos.sort(key=lambda x: (x["prio"] or 9, -(x["sim"] or 0), x["domain"]))
    ws = wb.create_sheet("Logo na home")
    ws.append(["prioridade", "domínio", "site", "onde está o logo", "semelhança com o oficial", "logo encontrado",
               "sinais de que se apresenta como a marca", "título do site", "URL da imagem"])
    from openpyxl.drawing.image import Image as XLImage
    for i, x in enumerate(logos, start=2):
        ws.append([x["prio"], x["domain"], link(x["site"], "abrir"), x["where"],
                   f'{x["sim"]*100:.0f}%' if x["sim"] else "", "", "; ".join(x["identity"]), x["title"],
                   link(x["img"], "imagem") if x["img"] else ""])
        if x["thumb"]:
            try:
                img = XLImage(io.BytesIO(base64.b64decode(x["thumb"])))
                img.width, img.height = min(img.width, 160), min(img.height, 50)
                ws.add_image(img, f"F{i}")
                ws.row_dimensions[i].height = 42
            except Exception:
                pass
    style(ws, [10, 30, 7, 14, 12, 26, 50, 45, 9])
    json.dump({"official": [o["thumb"] for o in OFFICIAL_LOGOS], "found": logos},
              open(re.sub(r"\.xlsx$", "", OUT_FILE) + "_logos.json", "w", encoding="utf-8"), ensure_ascii=False)

    ws = wb.create_sheet("Produtos (URLs)")
    ws.append(["prioridade", "critérios", "domínio", "produto", "preço", "preço antes", "desconto %", "moeda", "URL do produto"])
    for x in prows:
        ws.append(x)
    style(ws, [10, 10, 28, 60, 10, 11, 10, 7, 90])

    ws = wb.create_sheet("Bloqueados (priorizados)")
    ws.append(["prioridade", "domínio", "site", "busca", "Google", "motivo", "registrador", "criado em", "situação", "sua revisão"])
    for x in blocked:
        ws.append(x)
    style(ws, [10, 32, 7, 7, 8, 40, 22, 11, 14, 14])

    ws = wb.create_sheet("Descartados")
    ws.append(["domínio", "motivo", "título do site", "titular (Registrant)"])
    for x in sorted(other, key=lambda x: (x[1], x[0])):
        ws.append(x)
    style(ws, [32, 22, 60, 80])

    ws = wb.create_sheet("Resumo")
    from collections import Counter
    cnt = Counter(r["_class"][1] if not r["_class"][0] else f'P{r["_class"][0]} {r["_class"][1]}' for r in results)
    ws.append(["categoria", "domínios"])
    for k, v in sorted(cnt.items(), key=lambda x: (not x[0].startswith("P"), x[0] if x[0].startswith("P") else -x[1])):
        ws.append([k, v])
    ws.append([])
    ws.append(["Suspeitos (aba 1)", len(sus)])
    ws.append(["Logo na home (aba 2)", len(logos)])
    ws.append(["Produtos com URL (aba 3)", len(prows)])
    ws.append(["Bloqueados p/ revisão manual (aba 4)", len(blocked)])
    ws.append([])
    ws.append(["Legenda: A = desconto ≥ %d%% ou preço < %s | B = marca/logo na home | C = modelos-chave copiados"
               " | OF = loja que se diz oficial, vende e mostra a marca no cabeçalho/domínio, fora da lista oficial"
               % (MIN_OFF, PRICE_ALERT or "—")])
    ws.append(["Logo: comparação visual com o logo oficial — ≥ %d%% = logo oficial; %d–%d%% = parecido (revisar)%s"
               % (LOGO_SIM * 100, LOGO_NEAR * 100, LOGO_SIM * 100,
                  "" if OFFICIAL_LOGOS else " · ATENÇÃO: nenhum logo oficial carregado, só sinais de texto")])
    ws.column_dimensions["A"].width = 60
    wb.save(OUT_FILE)
    return len(sus), len(prows), len(blocked), cnt

# --------------------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--workers", type=int, default=24)
    ap.add_argument("--retry-blocked", action="store_true")
    ap.add_argument("--haiku", action="store_true")
    ap.add_argument("--logos", action="store_true", help="só refaz a análise de logos dos domínios já no cache")
    a = ap.parse_args()

    rows = read_domains(a.xlsx)
    if a.limit:
        rows = rows[: a.limit]
    cache = load_cache()
    factory, workers, previous = None, a.workers, {}
    if a.retry_blocked:
        for r in rows:
            c = cache.get(r["domain"])
            if c and c.get("status") in RETRY:
                previous[r["domain"]] = cache.pop(r["domain"])
        try:
            import cloudscraper
            factory = lambda: cloudscraper.create_scraper(browser={"browser": "chrome", "platform": "windows"})
        except ImportError:
            pass
        workers = min(workers, 6)
        print(f"Nova tentativa de {len(previous)} bloqueados" + (" com cloudscraper" if factory else ""))
    if a.logos:
        redo = [cache[r["domain"]] for r in rows if r["domain"] in cache and cache[r["domain"]].get("status") == "OK"]
        print(f"Reanalisando logos de {len(redo)} sites (logos oficiais carregados: {len(OFFICIAL_LOGOS)})")
        with cf.ThreadPoolExecutor(workers) as ex:
            for res in ex.map(logo_pass, redo):
                cache[res["domain"]] = res
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            for r in cache.values():
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
    todo = [r for r in rows if r["domain"] not in cache]
    print(f"{BRAND}: {len(rows)} domínios | {len(rows) - len(todo)} no cache | {len(todo)} a verificar")

    lock, done, t0 = threading.Lock(), 0, time.time()
    with open(CACHE_FILE, "a", encoding="utf-8") as f, cf.ThreadPoolExecutor(workers) as ex:
        futs = {ex.submit(scan, r, factory): r for r in todo}
        for fu in cf.as_completed(futs):
            row = futs[fu]
            try:
                res = fu.result()
            except Exception as e:
                res = {**row, "status": "ERROR", "note": str(e)[:200]}
            old = previous.get(row["domain"])
            if old and res.get("status") in ("DEAD", "ERROR"):  # nova tentativa falhou: mantém o status antigo
                res = {**old, "note": "2ª tentativa falhou"}
            with lock:
                cache[res["domain"]] = res
                f.write(json.dumps(res, ensure_ascii=False) + "\n")
                f.flush()
                done += 1
                if done % 50 == 0 or done == len(todo):
                    el = time.time() - t0
                    print(f"  {done}/{len(todo)}  ({el/60:.1f} min, faltam ~{el/done*(len(todo)-done)/60:.0f} min)")

    for r in cache.values():
        for k, v in (("products", []), ("brand_tags", []), ("models", []), ("snippets", []),
                     ("mentions", 0), ("mentions_ctx", 0), ("title", ""), ("final_url", ""), ("search_page", "")):
            r.setdefault(k, v)
    results = []
    for r in rows:  # dados do Excel atual (titular, registrador, data) prevalecem sobre o cache
        c = cache[r["domain"]]
        c.update({k: v for k, v in r.items() if v and k != "domain"})
        results.append(c)

    if a.haiku:
        import anthropic
        client = anthropic.Anthropic(max_retries=4)
        pend = [r for r in results if classify(r)[0] and (r.get("haiku") or {}).get("flags", "?") == "?"]
        print(f"Etapa 2: {len(pend)} suspeitos para o Haiku")
        if pend:
            try:
                pend[0]["haiku"] = ask_haiku(client, pend[0])
                pend = pend[1:]
            except Exception as e:
                print(f"\n*** O Haiku não respondeu: {str(e)[:300]}\n*** Verifique a chave e o crédito. Etapa 2 ignorada.\n")
                pend = []
        tin = tout = 0
        with cf.ThreadPoolExecutor(8) as ex:
            futs = {ex.submit(ask_haiku, client, r): r for r in pend}
            for fu in cf.as_completed(futs):
                try:
                    futs[fu]["haiku"] = fu.result()
                except Exception as e:
                    futs[fu]["haiku"] = {"flags": "?", "evidence": f"erro: {str(e)[:120]}"}
        for r in results:
            t = (r.get("haiku") or {}).get("_tokens")
            if t:
                tin += t[0]; tout += t[1]
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            for r in cache.values():
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
        print(f"Haiku: {tin:,} tokens entrada, {tout:,} saída ≈ US$ {tin/1e6*0.10 + tout/1e6*0.50:.2f}")

    ns, npr, nb, cnt = write_excel(results)
    print("\nResumo:", ", ".join(f"{k}: {v}" for k, v in cnt.most_common()))
    print(f"Suspeitos: {ns} | produtos com URL: {npr} | bloqueados p/ revisão: {nb}  →  {OUT_FILE}")


if __name__ == "__main__":
    sys.exit(main())
