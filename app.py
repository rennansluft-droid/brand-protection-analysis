"""Proteção de marca — app web (Streamlit).
Configure a marca e os critérios, envie o Excel com os domínios e veja os suspeitos."""
import base64
import concurrent.futures as cf
import importlib
import json
import os
import re
import threading
import time

import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Proteção de marca", page_icon="🛡️", layout="wide")
WORK = os.path.abspath("work")
os.makedirs(WORK, exist_ok=True)

COMMON = {
    "suspicious_registrars": ["west263", "west dimension", "eu technology", "bf internet", "cloud yuqu",
                              "spaceship", "gname", "web commerce", "dynadot", "realtime register"],
    "sale_words": ["sale", "outlet", "soldes", "promo\\w*", "pas cher", "petits? prix", "cheap", "discount",
                   "korting", "rabatt", "günstig\\w*", "d[ée]stockage", "liquidation", "black friday", "deals?"],
    "recent_from": "2025-01-01",
    "haiku_model": "claude-haiku-5-5",
}
PRESETS = {
    "Gertrude": {**COMMON,
        "brand": "Gertrude", "keywords": ["gertrude"], "official_domains": ["gertrude-store.com"],
        "official_logo_urls": [],
        "official_registrants": [], "authorized_resellers": [], "search_term": "gertrude",
        "context_terms": ["doudoune", "puffer", "down jacket", "veste", "manteau", "jacket", "coat", "parka",
                          "matelass", "duvet", "anorak"],
        "model_terms": ["azaelbright", "azael bright", "vianneybright", "vianney bright", "vianneybr",
                        "lilibright", "lili bright", "little lili", "francebright", "france bright"],
        "product_terms": ["casquette", "sweat", "t-shirt", "tee-shirt", "jogging", "hoodie", "bonnet", "pull",
                          "gilet", "sans manche", "short"],
        "min_discount": 30, "price_alert": 120,
        "price_alert_terms": ["doudoune", "puffer", "down jacket", "jacket", "parka", "azael", "vianney", "lili",
                              "bright", "sans manche"],
        "criteria": "A: doudoune ou produto Gertrude com desconto ≥ 30% ou preço abaixo de 120 € (oficiais: "
                    "Azael Bright 349 €, Lili Bright 349 €, Vianney Bright 329 €, Little Lili Bright 299 €). "
                    "Abaixo de 120 € = preço de falsificação; revendedores legítimos vendem o colete César a ~199 €.\n"
                    "B: nome ou logo Gertrude na página inicial de um site não oficial.\n"
                    "C: modelos-chave copiados (Azael, Vianney, Lili, Little Lili Bright), principalmente a doudoune preta.\n"
                    "'Gertrude' também é nome de pessoa/restaurante: só conta se houver relação com roupas."},
    "ASSOULINE": {**COMMON,
        "brand": "ASSOULINE", "keywords": ["assouline"], "official_domains": ["assouline.com"],
        "official_logo_urls": [],
        "official_registrants": ["assouline publishing", "editions assouline", "prosper assouline"],
        "authorized_resellers": [], "search_term": "assouline",
        "context_terms": ["book", "livre", "boek", "buch", "bildband", "libro", "livro", "coffee table",
                          "édition", "edition", "publishing", "bookstand"],
        "model_terms": ["impossible collection", "travel series", "ultimate collection", "bookstand"],
        "product_terms": [], "min_discount": 20, "price_alert": 0, "price_alert_terms": [],
        "criteria": "A: livros ASSOULINE com desconto ≥ 20%.\nB: logo/nome ASSOULINE na home de site não oficial, "
                    "apresentando-se como a marca.\nC: coleções-chave (Impossible Collection, Travel Series) em loja suspeita.\n"
                    "Assouline também é sobrenome: só conta se houver relação com livros."},
    "Nestlé": {**COMMON,
        "brand": "Nestlé", "keywords": ["nestlé", "nestle"], "search_term": "nestle",
        "official_domains": ["nestle.com"],
        "official_domain_regex": r"(^|\.)nestle\.(com|[a-z]{2}|com\.[a-z]{2}|co\.[a-z]{2})$",
        "official_logo_urls": [],
        "official_registrars": ["Com Laude", "Nom-IQ"],
        "official_registrants": ["société des produits nestlé", "societe des produits nestle", "nestlé s.a",
                                 "nestle s.a", "nestec", "nestlé brasil", "nestle brasil", "nestlé france",
                                 "nestle france", "nestlé usa", "nestle usa", "nestlé uk", "nestle uk",
                                 "nestlé españa", "nestle espana", "nestlé deutschland", "nestle deutschland",
                                 "nestlé portugal", "nestle portugal", "nestlé italiana", "nestle italiana",
                                 "nestle health science", "nestlé health science"],
        "authorized_resellers": [],
        "context_terms": ["café", "coffee", "chocolat", "chocolate", "cereal", "céréales", "cereais", "leite",
                          "lait", "milk", "baby", "bebê", "bébé", "infant formula", "fórmula infantil",
                          "lait infantile", "capsule", "cápsula", "food", "alimentos", "nutrition", "nutrição"],
        "model_terms": ["nescafé", "nescafe", "nespresso", "dolce gusto", "kitkat", "kit kat", "nesquik", "ninho",
                        "cerelac", "nan optipro", "nan supreme", "maggi", "purina", "leite moça", "galak", "nestea",
                        "smarties", "milo", "nido", "garoto", "neston", "nestogeno"],
        "product_terms": [], "min_discount": 50, "price_alert": 0, "price_alert_terms": [],
        "criteria": "A: produtos Nestlé com desconto ≥ 50%.\n"
                    "B: nome ou logo Nestlé no cabeçalho de um site que não é oficial.\n"
                    "C: produtos-chave (Nescafé, Dolce Gusto, KitKat, Nesquik, Ninho, NAN, Cerelac etc.).\n"
                    "OF: loja que vende produtos, mostra Nestlé no cabeçalho ou no domínio e se diz 'oficial' — "
                    "é SEMPRE suspeita se não estiver na lista de sites oficiais; dizer-se oficial não prova nada.\n"
                    "Atenção: 'nestle' também é um verbo em inglês ('to nestle'): só conta se houver relação com "
                    "alimentos/bebidas/marcas Nestlé."},
    "Nova marca": {**COMMON,
        "brand": "", "keywords": [], "official_domains": [], "official_logo_urls": [], "official_registrants": [],
        "authorized_resellers": [],
        "search_term": "", "context_terms": [], "model_terms": [], "product_terms": [], "min_discount": 30, "price_alert": 0,
        "price_alert_terms": [], "criteria": ""},
}
# Marcas salvas no repositório: cada arquivo marca_*.json (na raiz ou na pasta marcas/) vira uma opção
HERE = os.path.dirname(os.path.abspath(__file__))
_brand_files = [os.path.join(HERE, f) for f in os.listdir(HERE) if f.lower().startswith("marca_") and f.lower().endswith(".json")]
if os.path.isdir(os.path.join(HERE, "marcas")):
    _brand_files += [os.path.join(HERE, "marcas", f) for f in os.listdir(os.path.join(HERE, "marcas")) if f.lower().endswith(".json")]
for path in sorted(_brand_files):
            try:
                c = json.load(open(path, encoding="utf-8"))
                PRESETS[c.get("brand") or os.path.basename(path)[6:-5]] = {**COMMON, **PRESETS.get(c.get("brand"), {}), **c}
            except Exception:
                pass
PRESETS["Nova marca"] = PRESETS.pop("Nova marca")  # sempre por último

LIST_FIELDS = [
    ("keywords", "Como o nome da marca aparece nos sites"),
    ("official_domains", "Sites oficiais"),
    ("official_registrants", "Titulares oficiais — domínios registrados em nome destes (coluna Registrant) são oficiais"),
    ("official_registrars", "Registradores/gestores oficiais (ex.: Com Laude) — domínios registrados por eles são oficiais"),
    ("authorized_resellers", "Revendedores autorizados (domínios)"),
    ("context_terms", "Termos de contexto (provam que é a marca certa, evitam homônimos)"),
    ("model_terms", "Modelos/produtos-chave (critério C)"),
    ("product_terms", "Outros produtos da marca (ex.: casquette, sweat) — contam como produto da marca"),
    ("price_alert_terms", "Produtos aos quais o preço de alerta se aplica"),
]


def to_list(txt):
    return [x.strip() for x in re.split(r"[\n,;]+", txt or "") if x.strip()]


def link_of(v):
    m = re.match(r'=HYPERLINK\("([^"]+)"', str(v or ""))
    return m.group(1) if m else v


def slug(name):
    import unicodedata
    t = "".join(c for c in unicodedata.normalize("NFKD", (name or "marca").lower()) if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "_", t).strip("_") or "marca"


def load_engine(cfg):
    path = os.path.join(WORK, f"config_{slug(cfg['brand'])}.json")
    json.dump(cfg, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    os.environ["BRAND_CONFIG"] = path
    import brand_scan
    bs = importlib.reload(brand_scan)
    bs.CACHE_FILE = os.path.join(WORK, f"cache_{bs.SLUG}.jsonl")
    bs.OUT_FILE = os.path.join(WORK, f"{bs.SLUG}_resultados.xlsx")
    return bs


def run_scan(bs, xlsx_path, limit=0, workers=24, retry=False, on_progress=None, logos_only=False):
    """Mesma lógica do notebook: varre, guarda cache, gera o Excel. Devolve (resultados, resumo)."""
    rows = bs.read_domains(xlsx_path)
    if limit:
        rows = rows[:limit]
    cache = bs.load_cache()
    if logos_only:  # refaz só a análise de logos dos sites já varridos
        redo = [cache[r["domain"]] for r in rows if cache.get(r["domain"], {}).get("status") == "OK"]
        t0 = time.time()
        with cf.ThreadPoolExecutor(workers) as ex:
            futs = [ex.submit(bs.logo_pass, r) for r in redo]
            for i, fu in enumerate(cf.as_completed(futs), 1):
                try:
                    res = fu.result()
                    cache[res["domain"]] = res
                except Exception:
                    pass
                if on_progress:
                    on_progress(i, len(redo), time.time() - t0)
        with open(bs.CACHE_FILE, "w", encoding="utf-8") as f:
            for r in cache.values():
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
    factory, previous = None, {}
    if retry:
        for r in rows:
            c = cache.get(r["domain"])
            if c and c.get("status") in bs.RETRY:
                previous[r["domain"]] = cache.pop(r["domain"])
        try:
            import cloudscraper
            factory = lambda: cloudscraper.create_scraper(browser={"browser": "chrome", "platform": "windows"})
        except ImportError:
            pass
        workers = min(workers, 6)
    todo = [r for r in rows if r["domain"] not in cache]
    lock, done, t0 = threading.Lock(), 0, time.time()
    with open(bs.CACHE_FILE, "a", encoding="utf-8") as f, cf.ThreadPoolExecutor(workers) as ex:
        futs = {ex.submit(bs.scan, r, factory): r for r in todo}
        for fu in cf.as_completed(futs):
            row = futs[fu]
            try:
                res = fu.result()
            except Exception as e:
                res = {**row, "status": "ERROR", "note": str(e)[:200]}
            old = previous.get(row["domain"])
            if old and res.get("status") in ("DEAD", "ERROR"):
                res = {**old, "note": "2ª tentativa falhou"}
            with lock:
                cache[res["domain"]] = res
                f.write(json.dumps(res, ensure_ascii=False) + "\n")
                f.flush()
                done += 1
                if on_progress:
                    on_progress(done, len(todo), time.time() - t0)
    for r in cache.values():
        for k, v in (("products", []), ("brand_tags", []), ("models", []), ("snippets", []), ("mentions", 0),
                     ("mentions_ctx", 0), ("title", ""), ("final_url", ""), ("search_page", "")):
            r.setdefault(k, v)
    results = []
    for r in rows:  # dados do Excel atual (titular, registrador, data) prevalecem sobre o cache
        c = cache[r["domain"]]
        c.update({k: v for k, v in r.items() if v and k != "domain"})
        results.append(c)
    return results, bs.write_excel(results)


def run_haiku(bs, results, api_key, on_progress=None):
    import anthropic
    client = anthropic.Anthropic(api_key=api_key, max_retries=4)
    pend = [r for r in results if bs.classify(r)[0] and (r.get("haiku") or {}).get("flags", "?") == "?"]
    if not pend:
        return 0, 0.0, bs.write_excel(results)
    pend[0]["haiku"] = bs.ask_haiku(client, pend[0])  # testa chave/crédito; erro sobe para a interface
    n, tin, tout = 1, 0, 0
    with cf.ThreadPoolExecutor(8) as ex:
        futs = {ex.submit(bs.ask_haiku, client, r): r for r in pend[1:]}
        for fu in cf.as_completed(futs):
            try:
                futs[fu]["haiku"] = fu.result()
            except Exception as e:
                futs[fu]["haiku"] = {"flags": "?", "evidence": f"erro: {str(e)[:120]}"}
            n += 1
            if on_progress:
                on_progress(n, len(pend))
    for r in pend:
        t = (r.get("haiku") or {}).get("_tokens") or [0, 0]
        tin += t[0]; tout += t[1]
    with open(bs.CACHE_FILE, "w", encoding="utf-8") as f:
        for r in results:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    return len(pend), tin / 1e6 * 0.10 + tout / 1e6 * 0.50, bs.write_excel(results)


def sheet_df(path, name):
    ws = openpyxl.load_workbook(path)[name]
    data = list(ws.iter_rows(values_only=True))
    if not data:
        return pd.DataFrame()
    df = pd.DataFrame(data[1:], columns=data[0])
    for c in df.columns:
        if df[c].astype(str).str.startswith("=HYPERLINK").any():
            df[c] = df[c].map(link_of)
    return df.drop(columns=[c for c in ("sua revisão",) if c in df.columns])


def show_table(df, key):
    link_cols = {c: st.column_config.LinkColumn(c, display_text="abrir" if c in ("site", "busca") else
                 ("Google" if c == "Google" else ("ver produtos" if c == "produtos" else None)))
                 for c in df.columns if df[c].astype(str).str.match(r"https?://").any()}
    st.dataframe(df, column_config=link_cols, hide_index=True, width="stretch", key=key)

# =========================================================================== interface

st.title("🛡️ Proteção de marca")
st.caption("1. Configure a marca e os critérios  ·  2. Envie o Excel com os domínios  ·  3. Veja e baixe os suspeitos")

with st.sidebar:
    st.header("1 · Marca e critérios")
    preset = st.selectbox("Partir de", list(PRESETS), help="Escolha uma marca já configurada ou 'Nova marca'.")
    up_cfg = st.file_uploader("…ou carregar uma configuração salva (.json)", type=["json"])
    base = json.load(up_cfg) if up_cfg else PRESETS[preset]
    k = f"{preset}_{up_cfg.name if up_cfg else ''}"

    brand = st.text_input("Marca", base["brand"], key="brand" + k)
    search_term = st.text_input("Termo para a busca interna das lojas", base.get("search_term") or brand.lower(),
                                key="term" + k)
    saved_b64 = list(base.get("official_logos_b64", []))
    with st.expander("Endereço do logo oficial (opcional)"):
        logo_urls = to_list(st.text_area("URL(s) da imagem do logo", "\n".join(base.get("official_logo_urls", [])),
                                         height=68, key="lu" + k,
                                         help="Clique com o botão direito no logo do site oficial → Copiar endereço da imagem."))
    lists = {}
    for field, label in LIST_FIELDS:
        lists[field] = to_list(st.text_area(label, "\n".join(base.get(field, [])), height=90,
                                            key=field + k, help="Um item por linha (ou separados por vírgula)."))
    c1, c2 = st.columns(2)
    min_discount = c1.number_input("Desconto mínimo (%) – A", 0, 95, int(base.get("min_discount", 30)), key="md" + k)
    price_alert = c2.number_input("Preço de alerta – A", 0.0, 100000.0, float(base.get("price_alert", 0)),
                                  help="0 = desligado. Produtos abaixo deste preço contam como A.", key="pa" + k)
    with st.expander("Sinais de loja falsa (avançado)"):
        susp = to_list(st.text_area("Registradores suspeitos", "\n".join(base["suspicious_registrars"]), key="sr" + k))
        sale = to_list(st.text_area("Palavras de promoção (regex)", "\n".join(base["sale_words"]), key="sw" + k))
        recent = st.text_input("Domínio 'recente' a partir de (AAAA-MM-DD)", base.get("recent_from", "2025-01-01"),
                               key="rf" + k)
    criteria = st.text_area("Critérios para o Claude Haiku (opcional)", base.get("criteria", ""), height=160,
                            key="cr" + k)
    cfg = {**COMMON, "official_domain_regex": base.get("official_domain_regex", ""),
           "brand": brand, "search_term": search_term, **lists, "official_logos_b64": saved_b64,
           "official_logo_urls": logo_urls, "min_discount": int(min_discount),
           "price_alert": float(price_alert), "suspicious_registrars": susp, "sale_words": sale,
           "recent_from": recent, "criteria": criteria}

st.subheader("2 · Lista de domínios e logo oficial")
c1, c2, c3 = st.columns([5, 4, 3])
up_xlsx = c1.file_uploader("Excel com a coluna **Domain** (Registrar, Registrant e Created são aproveitados se existirem)",
                           type=["xlsx"])
up_logos = c2.file_uploader("🖼️ **Logo oficial da marca** (PNG, JPG, WebP ou SVG — pode enviar várias versões)",
                            type=["png", "jpg", "jpeg", "webp", "svg", "gif"], accept_multiple_files=True)
up_cache = c3.file_uploader("Cache de uma execução anterior (opcional)", type=["jsonl"])
if up_logos:
    cfg["official_logos_b64"] = [base64.b64encode(f.getvalue()).decode() for f in up_logos]
    c2.image([f.getvalue() for f in up_logos if not f.name.lower().endswith(".svg")], width=150)
elif saved_b64:
    c2.caption("✅ Usando o logo salvo na configuração desta marca:")
    c2.image([base64.b64decode(x) for x in saved_b64], width=150)
logo_b64 = cfg["official_logos_b64"]
st.download_button("💾 Salvar esta marca (critérios + logo) para usar depois", json.dumps(cfg, ensure_ascii=False, indent=1),
                   file_name=f"marca_{slug(brand)}.json", mime="application/json",
                   help="Na próxima vez, carregue este arquivo na barra lateral. Se colocar no GitHub, a marca aparece "
                        "na lista 'Partir de'.")
c1, c2, c3, c4 = st.columns(4)
test = c1.checkbox("Só testar com 30 domínios")
retry = c2.checkbox("Tentar de novo os bloqueados", help="Use depois de uma varredura completa.")
logos_only = c3.checkbox("Só reanalisar os logos", help="Usa o cache: refaz apenas a comparação de logos "
                         "(ex.: depois de enviar o logo oficial). Bem mais rápido.")
workers = c4.slider("Conexões simultâneas", 4, 24, 12, help="Menos conexões = mais lento, porém mais estável.")
if not (logo_b64 or logo_urls):
    st.warning("Envie o **logo oficial** (acima) para a comparação visual. Sem ele, o critério B usa só sinais de texto.")

ready = bool(up_xlsx and brand and lists["keywords"])
if not ready:
    st.info("Preencha pelo menos **Marca** e **Como o nome aparece** na barra lateral e envie o Excel.")

if st.button("▶ Iniciar varredura", type="primary", disabled=not ready):
    bs = load_engine(cfg)
    xlsx_path = os.path.join(WORK, "lista.xlsx")
    open(xlsx_path, "wb").write(up_xlsx.getvalue())
    if up_cache:
        open(bs.CACHE_FILE, "wb").write(up_cache.getvalue())
    bar, info = st.progress(0.0), st.empty()

    def prog(done, total, el):
        bar.progress(done / max(total, 1))
        info.write(f"{done}/{total} domínios · {el/60:.1f} min · faltam ~{el/done*(total-done)/60:.0f} min "
                   "— deixe esta aba aberta.")

    if logos_only and not os.path.exists(bs.CACHE_FILE):
        st.error("Para reanalisar só os logos, envie também o cache da varredura anterior.")
        st.stop()
    results, (ns, npr, nb, cnt) = run_scan(bs, xlsx_path, 30 if test else 0, workers, retry, prog, logos_only)
    bar.progress(1.0)
    info.success(f"Concluído: {ns} suspeitos, {npr} produtos com URL, {nb} bloqueados para revisão.")
    st.session_state.update(out=bs.OUT_FILE, cache=bs.CACHE_FILE, cfg=cfg, xlsx=xlsx_path,
                            counts=dict(cnt), totals=(ns, npr, nb))

if "out" in st.session_state and os.path.exists(st.session_state["out"]):
    st.subheader("3 · Resultado")
    ns, npr, nb = st.session_state["totals"]
    counts = st.session_state["counts"]
    out = st.session_state["out"]
    logo_json = re.sub(r"\.xlsx$", "", out) + "_logos.json"
    ldata = json.load(open(logo_json, encoding="utf-8")) if os.path.exists(logo_json) else {"official": [], "found": []}
    LSIM = float(st.session_state["cfg"].get("logo_similarity", 0.88))
    own = [x for x in ldata["found"] if (x.get("sim") or 0) >= LSIM and x.get("where") != "página"]
    m = st.columns(7)
    m[0].metric("🖼️ Logo oficial no cabeçalho", len(own))
    for i, (lab, key) in enumerate((("Prioridade 1", "P1"), ("Prioridade 2", "P2"), ("Prioridade 3", "P3"),
                                    ("Prioridade 4", "P4")), start=1):
        m[i].metric(lab, sum(v for kk, v in counts.items() if kk.startswith(key)))
    m[5].metric("Bloqueados p/ revisar", nb)
    m[6].metric("Homônimos descartados", counts.get("HOMÔNIMO PROVÁVEL", 0))
    st.caption("A = desconto/preço suspeito · B = marca/logo na home · C = modelos-chave · "
               "Prioridade 1 = mais grave. Prioridade 4 pode ser revendedor: verifique a autorização.")

    tabs = st.tabs(["🖼️ Logo na home", "Suspeitos", "Produtos (URLs)", "Bloqueados (priorizados)", "Descartados",
                    "Resumo"])
    with tabs[0]:
        if ldata["official"]:
            st.write("**Logo oficial usado na comparação:**")
            st.image([base64.b64decode(t) for t in ldata["official"]], width=180)
        else:
            st.info("Nenhum logo oficial foi carregado: aparecem só sites com sinais de texto (nome do site, título, "
                    "nome do arquivo do logo). Envie o logo oficial na seção 2 e use **Só reanalisar os logos**.")
        f1, f2 = st.columns(2)
        only_head = f1.toggle("Só logos usados como logo do próprio site (cabeçalho/ícone)", value=True)
        min_sim = f2.slider("Semelhança mínima", 0, 100, 75, step=5)
        items = [x for x in ldata["found"] if (not only_head or x.get("where") != "página")
                 and ((x.get("sim") or 0) * 100 >= min_sim or (not x.get("sim") and x.get("identity")))]
        st.caption(f"{len(items)} sites · ≥ {LSIM*100:.0f}% = logo oficial · abaixo disso = parecido, revise a imagem")
        for i in range(0, len(items), 3):
            cols = st.columns(3)
            for col, x in zip(cols, items[i:i + 3]):
                with col.container(border=True):
                    if x.get("thumb"):
                        st.image(base64.b64decode(x["thumb"]))
                    else:
                        st.caption("(sem imagem — sinal de texto)")
                    sim = f'{x["sim"]*100:.0f}% · {x["where"]}' if x.get("sim") else "texto"
                    tag = "🔴" if x.get("prio") == 1 else "🟠" if x.get("prio") == 2 else "🟡" if x.get("prio") else "⚪"
                    st.markdown(f'{tag} **{x["domain"]}**  \n{sim} · prioridade {x.get("prio") or "—"}')
                    if x.get("identity"):
                        st.caption("; ".join(x["identity"])[:200])
                    st.link_button("Abrir site", x["site"])
    for tab, name in zip(tabs[1:], ["Suspeitos", "Produtos (URLs)", "Bloqueados (priorizados)", "Descartados", "Resumo"]):
        with tab:
            df = sheet_df(out, name)
            if name == "Suspeitos" and not df.empty:
                pr = st.multiselect("Prioridades", [1, 2, 3, 4], [1, 2, 3, 4], key="prf")
                df = df[df["prioridade"].isin(pr)]
            show_table(df, key="t_" + name) if not df.empty else st.write("Nada aqui.")

    c1, c2 = st.columns(2)
    c1.download_button("⬇ Baixar Excel do resultado", open(out, "rb").read(), file_name=os.path.basename(out),
                       width="stretch", type="primary")
    c2.download_button("⬇ Baixar cache (para continuar depois)", open(st.session_state["cache"], "rb").read(),
                       file_name=os.path.basename(st.session_state["cache"]), width="stretch")

    with st.expander("Etapa opcional: classificar os suspeitos com Claude Haiku"):
        st.write("Precisa de uma chave de API **com crédito** (platform.claude.com). Custo típico: centavos de dólar. "
                 "A chave não é guardada.")
        key = st.text_input("Chave de API Anthropic", type="password")
        if st.button("Classificar com Haiku", disabled=not key):
            bs = load_engine(st.session_state["cfg"])
            cache = bs.load_cache()
            results = [cache[r["domain"]] for r in bs.read_domains(st.session_state["xlsx"]) if r["domain"] in cache]
            hb = st.progress(0.0)
            ok = False
            try:
                n, cost, (ns, npr, nb, cnt) = run_haiku(bs, results, key, lambda d, t: hb.progress(d / t))
                st.session_state.update(counts=dict(cnt), totals=(ns, npr, nb))
                st.toast(f"Haiku classificou {n} suspeitos · custo ≈ US$ {cost:.2f}")
                ok = True
            except Exception as e:
                st.error(f"O Haiku não respondeu: {str(e)[:300]}. Verifique a chave e o crédito.")
            if ok:
                st.rerun()
