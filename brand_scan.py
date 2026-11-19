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
import argparse, concurrent.futures as cf, html as htmllib, json, os, re, sys, threading, time
from urllib.parse import urlparse, urljoin, quote_plus

import requests, openpyxl
from openpyxl.styles import Font, PatternFill

requests.packages.urllib3.disable_warnings()

CFG = json.load(open(os.environ.get("BRAND_CONFIG", "config.json"), encoding="utf-8"))
BRAND = CFG["brand"]
SLUG = re.sub(r"\W+", "_", BRAND.lower()).strip("_")
KW = [k.lower() for k in CFG["keywords"]]
KW_RE = re.compile(r"(?<![a-z0-9])(" + "|".join(map(re.escape, KW)) + r")(?![a-z0-9])", re.I)
MODELS = [m.lower() for m in CFG.get("model_terms", [])]
CONTEXT = [c.lower() for c in CFG.get("context_terms", [])] + MODELS
OFFICIAL_DOMAINS = [d.lower() for d in CFG.get("official_domains", [])]
OFFICIAL_REG = [r.lower() for r in CFG.get("official_registrants", [])]
AUTHORIZED = [d.lower() for d in CFG.get("authorized_resellers", [])]
MIN_OFF = int(CFG.get("min_discount", 30))
PRICE_ALERT = float(CFG.get("price_alert", 0) or 0)
TERM = CFG.get("search_term") or KW[0]
PRICE_TERMS = [t.lower() for t in CFG.get("price_alert_terms", [])] or CONTEXT
SUSP_REG = [r.lower() for r in CFG.get("suspicious_registrars", [])]
RECENT = CFG.get("recent_from", "2025-01-01")
SALE_RE = re.compile(r"\b(" + "|".join(CFG.get("sale_words", ["sale", "outlet", "soldes", "promo"])) + r")\b", re.I)
HAIKU_MODEL = CFG.get("haiku_model", "claude-haiku-5-5")

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
    t = (p.get("title") or "").lower() + " " + (p.get("url") or "").lower()
    return bool(KW_RE.search(t)) or any(m in t or m.replace(" ", "-") in t for m in MODELS)

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
    res["models"] = sorted(set(res["models"]) | set(model_hits(" ".join(p["title"] + " " + p["url"] for p in res["products"]).replace("-", " "))))
    return res

# --------------------------------------------------------------------------- classification

def is_official(r):
    host = (urlparse(r.get("final_url") or "").hostname or r["domain"]).lower()
    reg = (r.get("registrant") or "").lower()
    return (any(host == o or host.endswith("." + o) for o in OFFICIAL_DOMAINS)
            or any(r["domain"] == o or r["domain"].endswith("." + o) for o in OFFICIAL_DOMAINS)
            or any(o and o in reg for o in OFFICIAL_REG))


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
    if KW_RE.search(r["domain"]):
        sig.append("marca no domínio")
    score = sum(2 if "registrador" in s or "marca no" in s else 1 for s in sig)
    return score, sig


def flags(r):
    prods = r.get("products", [])
    A_items = [p for p in prods if (p["off"] or 0) >= MIN_OFF
               or (PRICE_ALERT and p["price"] and 0 < p["price"] < PRICE_ALERT
                   and any(t in p["title"].lower() for t in PRICE_TERMS))]
    brand_title = bool(KW_RE.search(r.get("title") or ""))
    logo = any(is_logo(t) for t in r.get("brand_tags", []))
    rel = r.get("mentions_ctx", 0) or prods or r.get("models") or KW_RE.search(r["domain"])
    B = bool(rel) and (brand_title or logo or bool(KW_RE.search(r["domain"])))
    C = bool(r.get("models"))
    return bool(A_items), B, C, A_items, bool(rel)


def classify(r):
    """(prioridade, rótulo). Prioridade 1–4 = suspeito; None = fora da lista de suspeitos."""
    if r.get("status") != "OK":
        return None, r.get("status", "?")
    if is_official(r):
        return None, "OFICIAL"
    if any(r["domain"] == a or r["domain"].endswith("." + a) for a in AUTHORIZED):
        return None, "REVENDEDOR AUTORIZADO"
    A, B, C, _, rel = flags(r)
    if not rel:
        return None, "HOMÔNIMO PROVÁVEL" if r.get("mentions") else "SEM MENÇÃO"
    net, _ = network(r)
    label = "+".join(x for x, f in (("A", A), ("B", B), ("C", C)) if f) or "MENÇÃO"
    if (A and (B or C or net >= 2)) or (B and C and net >= 2):
        return 1, label
    if A or ((B or C) and net >= 2):
        return 2, label
    if net >= 2:
        return 3, label
    if B or C:
        return 4, label  # pode ser revendedor: verificar autorização
    return None, "MENÇÃO"

# --------------------------------------------------------------------------- stage 2 (Haiku)

HAIKU_SYSTEM = """You are a brand-protection analyst for {brand} (official sites: {official}).
You receive data already extracted from ONE website. Use ONLY this data. Criteria:
{criteria}
Authorized resellers or unrelated businesses that merely share the name are NOT suspects.
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
    col = lambda name: head.index(name) if name in head else None
    di = col("domain") if col("domain") is not None else 0
    ci = {k: col(k) for k in ("registrant", "registrar", "created")}
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

    sus, prows, blocked, other = [], [], [], []
    for r in results:
        pr, label = classify(r)
        r["_class"] = (pr, label)
        d = r["domain"]
        g = f"https://www.google.com/search?q=site%3A{d}+{quote_plus(TERM)}"
        if pr:
            A, B, C, A_items, _ = flags(r)
            net, sig = network(r)
            h = r.get("haiku") or {}
            ev_a = "; ".join(f'{p["title"][:50]} {p["price"]:g}' + (f' (antes {p["was"]:g}, -{p["off"]}%)' if p["off"] else "")
                             for p in sorted(A_items, key=lambda p: (-(p["off"] or 0), p["price"] or 0))[:3])
            prices = [p["price"] for p in r["products"] if p["price"]]
            logo = next((t for t in r["brand_tags"] if is_logo(t)), "")
            sus.append([pr, label, d, link("https://" + d, "abrir"), link(r["search_page"], "ver produtos"),
                        link(g, "Google"), r["title"][:120], ev_a, max([p["off"] for p in r["products"]] or [0]) or None,
                        min(prices) if prices else None, ", ".join(r["models"]),
                        logo[:150] or ("marca no título" if KW_RE.search(r["title"]) else ""),
                        ", ".join(sig), r.get("lang", ""), r.get("registrar", ""), r.get("created", ""),
                        h.get("flags", ""), h.get("evidence", ""), ", ".join(h.get("red_flags") or []), ""])
            for p in r["products"]:
                prows.append([pr, label, d, p["title"], p["price"], p["was"], p["off"] or None, p.get("cur", ""),
                              link(p["url"], p["url"][:90])])
        elif r.get("status") in RETRY:
            score, why = 0, []
            reg = (r.get("registrar") or "").lower()
            created = str(r.get("created") or "")[:10]
            if KW_RE.search(d): score += 3; why.append("marca no domínio")
            if any(x in reg for x in SUSP_REG) and created >= RECENT: score += 2; why.append("registrador suspeito + recente")
            elif created >= RECENT: score += 1; why.append("criado recentemente")
            if SALE_RE.search(d): score += 1; why.append("promoção no domínio")
            prio = "Alta" if score >= 3 else "Média" if score >= 1 else "Baixa"
            blocked.append([prio, d, link("https://" + d, "abrir"), link(f"https://{d}/search?q={quote_plus(TERM)}", "buscar"),
                            link(g, "Google"), ", ".join(why), r.get("registrar", ""), created, r["status"], ""])
        elif label in ("OFICIAL", "HOMÔNIMO PROVÁVEL", "REVENDEDOR AUTORIZADO"):
            other.append([d, label, r.get("title", "")[:120], ", ".join(r.get("snippets", [])[:1])[:200]])

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
    ws.append(["domínio", "motivo", "título do site", "trecho"])
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
    ws.append(["Produtos com URL (aba 2)", len(prows)])
    ws.append(["Bloqueados p/ revisão manual (aba 3)", len(blocked)])
    ws.append([])
    ws.append(["Legenda: A = desconto ≥ %d%% ou preço < %s | B = marca/logo na home | C = modelos-chave copiados"
               % (MIN_OFF, PRICE_ALERT or "—")])
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
    results = [cache[r["domain"]] for r in rows]

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
