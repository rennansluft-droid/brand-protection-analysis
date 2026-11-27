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
PEA = ['pea protein', 'protéine de pois', 'proteine de pois', 'erbsenprotein', 'proteína de guisante', 'proteína de ervilha', 'proteina di piselli', 'yellow pea', 'pois jaune', 'pea isolate', 'isolate', 'plant protein', 'plant-based', 'vegan', 'protein powder', 'supplement', 'ingredient']
CHIC = ['chicory', 'chicory root', 'chicory root fibre', 'chicory root fiber', 'chicory inulin', 'chicorée', 'racine de chicorée', 'fibre de chicorée', 'zichorie', 'zichorienwurzel', 'achicoria', 'raíz de achicoria', 'chicória', 'raiz de chicória', 'cicoria', 'inulin', 'inuline', 'inulina', 'oligofructose', 'fructooligosaccharide', 'fructo-oligosaccharide', 'fos', 'fructan', 'prebiotic', 'prébiotique', 'präbiotisch', 'prebiótico', 'soluble fibre', 'soluble fiber', 'dietary fibre', 'dietary fiber', 'fibre alimentaire', 'ballaststoff', 'fibra soluble', 'fibra alimentar', 'fibra solúvel', 'sugar reduction', 'fat replacer', 'ingredient', 'bulk', 'moq', 'kg', 'powder', 'supplement']

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
    "Quintet": {**COMMON,
        "brand": "Quintet", "brand_type": "finance", "keywords": ["quintet"], "search_term": "quintet",
        "official_domains": ["quintet.lu", "quintet.com", "myquintet.lu", "myquintetpro.lu", "puilaetco.be",
                             "quintet.dk", "merckfinck.de", "insingergilissen.nl", "brownshipley.com"],
        "official_logo_urls": [], "official_registrants": ["Quintet Private Bank", "KBL European Private Bankers"],
        "official_registrars": [], "authorized_resellers": [],
        "known_bad_domains": ["quintet-cp.com", "quintet-private.lu", "quintet-private.com", "quintetcapitalsa.com"],
        "context_terms": ["bank", "banque", "banking", "private bank", "banque privée", "privatbank", "bankhaus",
                          "wealth management", "gestion de patrimoine", "gestion de fortune", "vermögensverwaltung",
                          "asset management", "investment", "investissement", "e-banking", "cssf", "iban",
                          "family office", "custody", "dépositaire"],
        "model_terms": ["quintet private bank", "quintet luxembourg", "quintet europe", "quintet financial services",
                        "quintet capital", "myquintet", "kbl european private bankers", "kbl epb", "puilaetco",
                        "merck finck", "insingergilissen", "brown shipley"],
        "product_terms": [], "min_discount": 0, "price_alert": 0, "price_alert_terms": [],
        "criteria": "Banco privado (Quintet Private Bank (Europe) S.A., Luxemburgo). O risco é se passar pelo banco:\n"
                    "K: domínio já citado em alerta oficial (CSSF): quintet-private.com, quintetcapitalsa.com, quintet-cp.com, quintet-private.lu.\n"
                    "D: outro domínio que imita o banco (quintet + private/capital/bank/invest/login…), mesmo sem site.\n"
                    "L: página de login/senha com a marca fora dos sites oficiais (phishing).\n"
                    "F: oferta de investimento/crédito/recuperação de fundos com a marca (rendimento garantido, "
                    "cripto, WhatsApp, taxa de liberação, 'Quintet Capital S.A.', falsas filiais na Suécia/França).\n"
                    "B/OF: site que usa o nome/logo da Quintet ou se diz oficial.\n"
                    "Descartar: 'quintet' no sentido musical (string/jazz quintet), outras empresas chamadas Quintet "
                    "sem relação com banco, e notícias sobre o banco."},
    "COSUCRA": {**COMMON,
        "brand": "COSUCRA", "brand_type": "mixed",
        "keywords": ["cosucra", "pisane", "fibruline", "pizane", "pisanne", "fibrulin", "fibrulina"],
        "search_term": "pisane", "extra_search_terms": ["fibruline", "cosucra"],
        "keyword_context": {k: PEA for k in ("pisane", "pizane", "pisanne")} | {k: CHIC for k in ("fibruline", "fibrulin", "fibrulina")},
        "official_domains": ["cosucra.com", "chicoline.be"],
        "official_domain_regex": r"(^|\.)cosucra\.(com|[a-z]{2}|com\.[a-z]{2}|co\.[a-z]{2})$",
        "official_logo_urls": [], "official_registrants": ["Cosucra", "Groupe Warcoing"], "official_registrars": [],
        "authorized_resellers": ["dksh.com", "dkshdiscover.com", "knowde.com"],
        "ignore_domains": ["sciencedirect.com", "elsevier.com", "mdpi.com", "researchgate.net", "nih.gov", "europepmc.org",
                           "springer.com", "wiley.com", "tandfonline.com", "frontiersin.org", "nature.com", "acs.org",
                           "rsc.org", "cambridge.org", "oup.com", "academia.edu", "semanticscholar.org", "scholar.google.com",
                           "patents.google.com", "espacenet.com", "wipo.int", "wikipedia.org",
                           "foodnavigator.com", "foodnavigator-usa.com", "foodnavigator-asia.com", "foodnavigator-latam.com",
                           "nutraingredients.com", "nutraingredients-usa.com", "nutraingredients-asia.com",
                           "foodingredientsfirst.com", "nutritioninsight.com", "food.be", "fooddive.com", "just-food.com"],
        "known_bad_domains": [],
        "risk_domain_words": ["group", "groupe", "sa", "inc", "llc", "ltd", "order", "orders", "procurement", "purchase",
                              "purchasing", "invoice", "invoices", "payment", "billing", "finance", "sales", "export",
                              "trading", "supply", "supplier", "ingredients", "food", "jobs", "job", "career", "careers",
                              "recruitment", "recrutement", "hr", "rh", "login", "portal", "secure", "mail", "office",
                              "belgium", "europe", "usa", "asia", "china"],
        "strong_risk_words": ["group", "groupe", "sa", "inc", "order", "orders", "procurement", "purchase", "purchasing",
                              "invoice", "invoices", "payment", "billing", "finance", "jobs", "careers", "recruitment",
                              "recrutement", "login", "portal", "secure"],
        "scam_terms": ["purchase order", "bon de commande", "request for quotation", "rfq", "proforma", "pro forma invoice",
                       "bank details", "change of bank", "new bank account", "nouveau rib", "urgent payment",
                       "wire transfer", "recruitment fee", "frais de recrutement", "training fee", "visa fee",
                       "whatsapp", "telegram"],
        "context_terms": sorted(set(PEA + CHIC + ["food", "ingrédient", "zutat", "ingrediente"])),
        "model_terms": ["pisane c9", "pisane b9", "pisane m9", "pisane f9", "fibruline instant", "fibruline xl",
                        "cosucra groupe warcoing", "groupe warcoing", "warcoing"],
        "product_terms": [],
        "watch_terms": ["oem", "odm", "private label", "custom logo", "customized logo", "customised logo", "your logo",
                        "own brand", "factory price", "factory direct", "manufacturer of", "we produce", "replica",
                        "imitation", "equivalent to", "same as", "substitute for", "alternative to", "compare to"],
        "watch_terms_2": ["moq", "minimum order", "min. order", "min order", "per kg", "/kg", "usd/kg", "per ton",
                          "metric ton", "fob", "cif", "wholesale", "in stock", "add to cart", "ajouter au panier",
                          "in den warenkorb", "buy now", "1kg", "1 kg", "500g", "500 g", "25kg", "25 kg", "20 kg"],
        "watch_brand_terms": [],
        "min_discount": 50, "price_alert": 0, "price_alert_terms": [],
        "criteria": "Fornecedor B2B belga de ingredientes (COSUCRA Groupe Warcoing). Marcas no escopo: COSUCRA, PISANE™ "
                    "(proteína isolada de ervilha; grades C9, B9, M9…) e FIBRULINE™ (fibra/inulina de chicória; Instant, XL…). "
                    "Área: mundo todo. Infração buscada: produtos com as marcas ou logos da COSUCRA, ou com marcas semelhantes.\n"
                    "A: venda de PISANE/FIBRULINE/COSUCRA por vendedor que não é a COSUCRA nem distribuidor autorizado "
                    "(marketplaces B2B como Alibaba, Made-in-China, IndiaMART, TradeIndia; embalagens de varejo de 1 kg/500 g).\n"
                    "W: sinais de falsificação junto com a marca: vendedor que se diz fabricante de PISANE, OEM/private label/"
                    "custom logo, preço de fábrica, 'equivalent to/same as PISANE' (uso da marca para vender outro produto).\n"
                    "B/OF: logo ou nome COSUCRA usado como próprio, ou 'distribuidor oficial' fora da lista.\n"
                    "S: marca semelhante no mesmo setor (Pizane, Pisanne, Fibrulin, Fibrulina…).\n"
                    "D/L/F: domínio que imita a COSUCRA (cosucra-group, cosucra-order, cosucra-invoice…), páginas de login, "
                    "falsos pedidos de compra/faturas ou falsas vagas (BEC).\n"
                    "U: cliente que cita 'made with PISANE™/FIBRULINE™' num produto acabado: não é infração, marcar para "
                    "validação do cliente.\n"
                    "Descartar: 'pisane' = 'escrito' em polonês/croata/esloveno e 'de Pisa' em italiano; artigos científicos "
                    "que citam 'Pisane C9 (Cosucra)' nos métodos; imprensa especializada; feiras e catálogos com página "
                    "da própria COSUCRA."},
    "Air Liquide": {**COMMON,
        "brand": "Air Liquide", "brand_type": "mixed",
        "keywords": ["air liquide", "airliquide", "seppic", "vitalaire", "vital aire"],
        "search_term": "air liquide", "extra_search_terms": ["seppic", "vitalaire"],
        "watch_terms": [], "watch_terms_2": [], "watch_brand_terms": [],
        "official_domains": ["airliquide.com", "fondationairliquide.com", "airliquidehr.wd3.myworkdayjobs.com",
                             "airgas.com", "seppic.com", "schuelke.com"],
        "official_domain_regex": r"(^|\.)(airliquide|seppic|vitalaire)\.(com|[a-z]{2}|com\.[a-z]{2}|co\.[a-z]{2})$",
        "official_logo_urls": [], "official_registrants": ["L'Air Liquide", "Air Liquide"], "official_registrars": [],
        "authorized_resellers": [],
        "known_bad_domains": ["airliquide-finance.com"],
        "risk_domain_words": ["finance", "invest", "action", "actions", "shares", "share", "bourse", "capital",
                              "bank", "banque", "patrimoine", "dividende", "trading", "crypto", "recrutement",
                              "career", "careers", "job", "jobs", "emploi", "rh", "hr", "login", "secure", "support",
                              "portal", "my", "online", "client", "verify", "sa", "group", "groupe", "promo", "offre"],
        "scam_terms": ["action à prix réduit", "actions à prix réduit", "décote", "remise sur le cours",
                       "discounted shares", "shares at a discount", "rendement garanti", "guaranteed return",
                       "high return", "rendement élevé", "sans risque", "risk-free", "whatsapp", "telegram",
                       "frais de dossier", "frais de formation", "frais de visa", "processing fee", "training fee",
                       "visa fee", "recovery", "récupération de fonds", "verify your account"],
        "context_terms": ["gaz", "gas", "oxygène", "oxygen", "azote", "nitrogen", "hydrogène", "hydrogen", "argon",
                          "industriel", "industrial", "santé", "healthcare", "patient", "actionnaire", "shareholder",
                          "dividende", "dividend", "bourse", "euronext", "cac 40", "bouteille", "cylinder",
                          "soudage", "welding", "recrutement", "careers",
                          "cosmétique", "cosmetic", "ingrédient", "ingredient", "excipient", "émulsifiant", "emulsifier",
                          "adjuvant", "vaccin", "vaccine", "oxygénothérapie", "oxygen therapy", "apnée du sommeil",
                          "sleep apnea", "ppc", "cpap", "concentrateur", "concentrator", "diabète", "diabetes",
                          "pompe à insuline", "insulin pump", "santé à domicile", "home healthcare"],
        "model_terms": ["air liquide healthcare", "air liquide santé", "air liquide finance", "air liquide group",
                        "groupe air liquide", "seppic", "vitalaire", "orkyn", "montanide", "sepigel", "simulgel",
                        "sepimax", "sepinov", "montanov", "sepicalm", "aquaxyl", "lanol", "sepiplus", "sepifeel"],
        "product_terms": ["gel", "emulsifier", "émulsifiant", "polymer", "polymère", "thickener", "épaississant",
                          "cpap", "ppc", "masque", "mask", "concentrateur", "concentrator", "oxygen", "oxygène"],
        "min_discount": 50, "price_alert": 0, "price_alert_terms": [],
        "criteria": "Empresa industrial cotada em bolsa (Air Liquide), com as marcas Seppic (ingredientes cosméticos/"
                    "farmacêuticos, adjuvantes Montanide) e VitalAire (saúde a domicílio: oxigênio, CPAP, diabetes) no escopo.\n"
                    "A/B: lojas que vendem produtos Seppic ou VitalAire com desconto alto, ou usam o nome/logo como se fossem oficiais.\n"
                    "F: falsas ofertas de ações Air Liquide com desconto sobre a cotação ou rendimento garantido "
                    "(a Air Liquide nunca faz isso); falsas vagas que cobram taxas.\n"
                    "K: domínio já citado em alerta oficial (lista negra da AMF, ex.: airliquide-finance.com).\n"
                    "D: outro domínio que imita a empresa (airliquide + finance/invest/action/careers/login…), mesmo sem site.\n"
                    "L: página de login com a marca fora dos sites oficiais.\n"
                    "Descartar: 'air liquide' no sentido científico (ar liquefeito), distribuidores legítimos de gás, "
                    "notícias e análises da ação."},
    "Tiger Beer": {**COMMON,
        "brand": "Tiger Beer", "brand_type": "mixed",
        "watch_terms": ["vape", "vapes", "vaping", "vape juice", "e-cig", "e-cigs", "ecig", "e-cigarette", "e-cigarettes",
                        "e-liquid", "e-liquids", "eliquid", "e-juice", "ejuice", "vape pod", "pod system",
                        "disposable vape", "salt nic", "saltnic", "nic salt", "nicotine", "rokok elektrik", "liquid vape", "liquid",
                        "liquid saltnic", "freebase", "atomizer", "mod vape", "vapor", "vapour"],
        "watch_terms_2": ["t-shirt", "t-shirts", "tshirt", "tshirts", "t shirt", "tee", "tees", "kaos", "singlet", "singlets",
                          "tank top", "tank tops", "hoodie", "hoodies", "sweatshirt", "jersey", "polo shirt", "kemeja",
                          "baju", "jacket", "jaket", "snapback", "baseball cap", "trucker hat", "bucket hat", "cap",
                          "caps", "topi", "beanie", "apparel", "merchandise", "merch", "clothing", "shorts", "celana",
                          "sandal", "sandals", "bandana"],
        "watch_brand_terms": [],
        "keywords": ["tiger beer", "tigerbeer", "tiger crystal", "tiger soju", "tiger radler", "tiger lager"],
        "search_term": "tiger beer",
        "official_domains": ["tigerbeer.com", "theheinekencompany.com", "heineken.com", "heinekenmalaysia.com"],
        "official_domain_regex": r"(^|\.)tigerbeer\.(com|[a-z]{2}|com\.[a-z]{2}|co\.[a-z]{2})$",
        "official_logo_urls": [], "official_registrants": ["Heineken"], "official_registrars": [],
        "authorized_resellers": [], "known_bad_domains": [],
        "risk_domain_words": ["promo", "promosi", "giveaway", "hadiah", "undian", "gift", "free", "gratis",
                              "lucky", "prize", "win", "menang", "claim", "klaim", "voucher", "reward", "event",
                              "festival", "shop", "store", "official", "resmi", "karir", "career", "careers",
                              "loker", "rekrutmen", "jobs", "login", "my", "app"],
        "scam_terms": ["free beer", "bir gratis", "gratis bir", "giveaway", "lucky draw", "undian", "hadiah",
                       "menangkan", "klaim hadiah", "claim your prize", "claim now", "win a", "cooler", "mini fridge",
                       "kulkas mini", "survey", "quiz", "kuis", "biaya administrasi", "biaya pendaftaran",
                       "transfer biaya", "processing fee", "whatsapp", "telegram"],
        "product_terms": [], "min_discount": 50, "price_alert": 0, "price_alert_terms": [],
        "context_terms": ["beer", "bir", "lager", "brew", "brewery", "can", "cans", "bottle", "pint", "alcohol",
                          "bar", "pub", "minuman", "heineken", "asia pacific breweries"],
        "model_terms": ["tiger crystal", "tiger soju", "tiger radler", "tiger white", "tiger black",
                        "tiger original", "uncage", "tiger street food"],
        "criteria": "Cerveja (Tiger Beer, grupo Heineken). PRIORIDADE DO CLIENTE: roupas com o logo (camisetas, bonés, regatas) e cigarros eletrônicos/vapes/e-liquids com a marca — qualquer venda por terceiros é suspeita. Outros riscos: falsas promoções/sorteios de 'cerveja grátis' "
                    "(quiz, prêmio, mini frigobar) usados para phishing; lojas falsas com desconto alto; falsas vagas "
                    "que cobram taxas; sites que se dizem oficiais.\n"
                    "Descartar: 'tiger' sem relação com cerveja (animal, outras marcas Tiger), bares e lojas "
                    "legítimos que vendem Tiger a preço normal, notícias."},
    "Bir Bintang": {**COMMON,
        "brand": "Bir Bintang", "brand_type": "mixed",
        "watch_terms": ["vape", "vapes", "vaping", "vape juice", "e-cig", "e-cigs", "ecig", "e-cigarette", "e-cigarettes",
                        "e-liquid", "e-liquids", "eliquid", "e-juice", "ejuice", "vape pod", "pod system",
                        "disposable vape", "salt nic", "saltnic", "nic salt", "nicotine", "rokok elektrik", "liquid vape", "liquid",
                        "liquid saltnic", "freebase", "atomizer", "mod vape", "vapor", "vapour"],
        "watch_terms_2": ["t-shirt", "t-shirts", "tshirt", "tshirts", "t shirt", "tee", "tees", "kaos", "singlet", "singlets",
                          "tank top", "tank tops", "hoodie", "hoodies", "sweatshirt", "jersey", "polo shirt", "kemeja",
                          "baju", "jacket", "jaket", "snapback", "baseball cap", "trucker hat", "bucket hat", "cap",
                          "caps", "topi", "beanie", "apparel", "merchandise", "merch", "clothing", "shorts", "celana",
                          "sandal", "sandals", "bandana"],
        "watch_brand_terms": ["bintang"],
        "keywords": ["bir bintang", "birbintang", "bintang pilsener", "bintang radler", "bintang crystal",
                     "bintang 0.0", "bintang zero", "multi bintang", "multibintang", "bintang beer"],
        "search_term": "bintang",
        "official_domains": ["birbintang.co.id", "multibintang.co.id", "theheinekencompany.com", "heineken.com",
                             "moderationmatters.com"],
        "official_domain_regex": r"(^|\.)(birbintang|multibintang)\.(com|co\.id|id)$",
        "official_logo_urls": [], "official_registrants": ["Multi Bintang", "Heineken"], "official_registrars": [],
        "authorized_resellers": [], "known_bad_domains": [],
        "risk_domain_words": ["promo", "promosi", "giveaway", "hadiah", "undian", "gift", "free", "gratis",
                              "lucky", "prize", "win", "menang", "claim", "klaim", "voucher", "reward", "event",
                              "festival", "shop", "store", "official", "resmi", "karir", "career", "careers",
                              "loker", "rekrutmen", "jobs", "login", "my", "app"],
        "scam_terms": ["free beer", "bir gratis", "gratis bir", "giveaway", "lucky draw", "undian", "hadiah",
                       "menangkan", "klaim hadiah", "claim your prize", "claim now", "win a", "cooler", "mini fridge",
                       "kulkas mini", "survey", "quiz", "kuis", "biaya administrasi", "biaya pendaftaran",
                       "transfer biaya", "processing fee", "whatsapp", "telegram"],
        "product_terms": [], "min_discount": 50, "price_alert": 0, "price_alert_terms": [],
        "context_terms": ["bir", "beer", "lager", "pilsener", "radler", "minuman", "alkohol", "botol", "kaleng",
                          "krat", "heineken", "multi bintang"],
        "model_terms": ["bintang pilsener", "bintang radler", "bintang crystal", "bintang 0.0", "bintang zero",
                        "green sands", "heineken"],
        "criteria": "Cerveja (Bir Bintang, PT Multi Bintang Indonesia, grupo Heineken). PRIORIDADE DO CLIENTE: roupas com o logo (kaos, singlet, topi) e vapes/liquid com a marca — qualquer venda por terceiros é suspeita. Outros riscos: falsas promoções/"
                    "sorteios (hadiah, undian, 'bir gratis'); falsas vagas (rekrutmen/loker) que cobram taxas; lojas "
                    "falsas; sites que se dizem 'resmi'.\n"
                    "Descartar: 'bintang' no sentido comum (estrela) ou outras empresas/produtos Bintang, lojas "
                    "legítimas que vendem Bintang a preço normal, notícias."},
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
    ("ignore_domains", "Fontes a ignorar (imprensa, artigos científicos, patentes…)"),
    ("context_terms", "Termos de contexto (provam que é a marca certa, evitam homônimos)"),
    ("model_terms", "Modelos/produtos-chave (critério C)"),
    ("product_terms", "Outros produtos da marca (ex.: casquette, sweat) — contam como produto da marca"),
    ("price_alert_terms", "Produtos aos quais o preço de alerta se aplica"),
    ("watch_terms", "Produtos vigiados — PRIORIDADE 1 (ex.: vape, e-liquid): venda com a marca já é suspeita"),
    ("watch_terms_2", "Produtos vigiados — PRIORIDADE 2 (ex.: t-shirt, boné, singlet)"),
    ("watch_brand_terms", "Nome 'solto' da marca, só vale junto com produto vigiado (ex.: bintang)"),
    ("extra_search_terms", "Termos extras para a busca interna das lojas (até 2, ex.: seppic, vitalaire)"),
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


DEFAULTS = (("products", []), ("brand_tags", []), ("models", []), ("snippets", []), ("mentions", 0),
            ("mentions_ctx", 0), ("title", ""), ("final_url", ""), ("search_page", ""))


def collect(rows, cache):
    """Resultados dos domínios já processados (na ordem do Excel); dados do Excel atual prevalecem."""
    out = []
    for r in rows:
        c = cache.get(r["domain"])
        if not c:
            continue
        for k, v in DEFAULTS:
            c.setdefault(k, v)
        c.update({k: v for k, v in r.items() if v and k != "domain"})
        out.append(c)
    return out


def save_cache(bs, cache):
    tmp = bs.CACHE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        for r in cache.values():
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    os.replace(tmp, bs.CACHE_FILE)


def run_scan(bs, xlsx_path, limit=0, workers=12, retry=False, on_progress=None, logos_only=False,
             batch=500, on_batch=None, logo_sig="", haiku=None):
    """Varre em lotes. Depois de cada lote: salva o cache e gera o Excel parcial (on_batch mostra os downloads).
    Se algo der errado, basta enviar o mesmo Excel + o último cache: só o que falta é processado."""
    rows = bs.read_domains(xlsx_path)
    if limit:
        rows = rows[:limit]
    cache = bs.load_cache()
    batch = max(50, int(batch or 500))

    def checkpoint(phase, done, total, nbatch, nbatches, batch_domains=()):
        res = collect(rows, cache)
        if haiku:   # classifica só os suspeitos novos deste lote e grava no cache
            changed = haiku.run([cache[d] for d in batch_domains if d in cache])
            if changed:
                with open(bs.CACHE_FILE, "a", encoding="utf-8") as f:
                    for r in changed:
                        f.write(json.dumps(r, ensure_ascii=False) + "\n")
        counts = bs.write_excel(res)
        if on_batch:
            prev = []
            for r in res:
                pr, lab = bs.classify(r)
                if pr in (1, 2):
                    prev.append({"prio": pr, "critérios": lab, "domínio": r["domain"], "título": (r.get("title") or "")[:70],
                                 "haiku": (r.get("haiku") or {}).get("flags", "")})
            prev.sort(key=lambda x: (x["prio"], str(x["haiku"]).lower() == "none"))
            on_batch(phase, done, total, nbatch, nbatches, len(res), len(rows), counts, prev[:60])

    # 1) reanálise só dos logos (também em lotes; o que já foi refeito com este logo é pulado)
    if logos_only:
        redo = [cache[r["domain"]] for r in rows if cache.get(r["domain"], {}).get("status") == "OK"
                and cache[r["domain"]].get("logo_sig") != logo_sig]
        nb, t0, done = (len(redo) + batch - 1) // batch, time.time(), 0
        for bi in range(nb):
            chunk = redo[bi * batch:(bi + 1) * batch]
            with cf.ThreadPoolExecutor(workers) as ex:
                for fu in cf.as_completed([ex.submit(bs.logo_pass, r) for r in chunk]):
                    try:
                        res = fu.result()
                        res["logo_sig"] = logo_sig
                        cache[res["domain"]] = res
                    except Exception:
                        pass
                    done += 1
                    if on_progress:
                        on_progress(done, len(redo), time.time() - t0)
            save_cache(bs, cache)
            checkpoint("logos", done, len(redo), bi + 1, nb, [r["domain"] for r in chunk])

    # 2) varredura dos domínios que ainda não estão no cache
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
    nb, t0, done = (len(todo) + batch - 1) // batch, time.time(), 0
    for bi in range(nb):
        chunk = todo[bi * batch:(bi + 1) * batch]
        with open(bs.CACHE_FILE, "a", encoding="utf-8") as f, cf.ThreadPoolExecutor(workers) as ex:
            futs = {ex.submit(bs.scan, r, factory): r for r in chunk}
            for fu in cf.as_completed(futs):
                row = futs[fu]
                try:
                    res = fu.result()
                except Exception as e:
                    res = {**row, "status": "ERROR", "note": str(e)[:200]}
                old = previous.get(row["domain"])
                if old and res.get("status") in ("DEAD", "ERROR"):
                    res = {**old, "note": "2ª tentativa falhou"}
                if logo_sig:
                    res["logo_sig"] = logo_sig
                cache[res["domain"]] = res
                f.write(json.dumps(res, ensure_ascii=False) + "\n")
                f.flush()
                done += 1
                if on_progress:
                    on_progress(done, len(todo), time.time() - t0)
        checkpoint("varredura", done, len(todo), bi + 1, nb, [r["domain"] for r in chunk])

    results = collect(rows, cache)
    if haiku:   # suspeitos que ficaram sem Haiku (ex.: cache de uma execução anterior)
        changed = haiku.run(results)
        if changed:
            save_cache(bs, cache)
    return results, bs.write_excel(results)


class HaikuRunner:
    """Classifica com o Claude Haiku só os suspeitos novos de cada lote. Se a chave ou o crédito falharem,
    desliga o Haiku e a varredura continua normalmente."""

    def __init__(self, bs, api_key, prios=(1, 2, 3), workers=4, max_cost=None):
        import anthropic
        self.client = anthropic.Anthropic(api_key=api_key, max_retries=3)
        self.bs, self.prios, self.workers = bs, set(prios), workers
        self.max_cost = float(max_cost) if max_cost else None   # teto de gasto (US$) para esta varredura
        self.n = self.tin = self.tout = 0
        self.err, self.disabled = None, False

    def done(self, r):
        if r.get("haiku"):
            r["haiku"] = self.bs.haiku_fix(r["haiku"])
        return (r.get("haiku") or {}).get("flags") not in (None, "", "?")

    @property
    def cost(self):
        return self.tin / 1e6 * 0.10 + self.tout / 1e6 * 0.50

    def run(self, items):
        if self.disabled:
            return []
        pend = [r for r in items if self.bs.classify(r)[0] in self.prios and not self.done(r)]
        if not pend:
            return []
        try:  # 1ª chamada testa chave e crédito
            pend[0]["haiku"] = self.bs.ask_haiku(self.client, pend[0])
        except Exception as e:
            self.err, self.disabled = str(e)[:300], True
            return []
        self._count(pend[:1])
        rest, step = pend[1:], self.workers * 5
        for i in range(0, len(rest), step):      # em pequenos blocos, conferindo o teto de gasto
            if self.disabled:
                break
            if self.max_cost and self.cost >= self.max_cost:
                self.err, self.disabled = f"teto de gasto de US$ {self.max_cost:.2f} atingido", True
                break
            part = rest[i:i + step]
            with cf.ThreadPoolExecutor(self.workers) as ex:
                futs = {ex.submit(self.bs.ask_haiku, self.client, r): r for r in part}
                for fu in cf.as_completed(futs):
                    r = futs[fu]
                    try:
                        r["haiku"] = fu.result()
                    except Exception as e:
                        msg = str(e)
                        if any(x in msg.lower() for x in ("credit", "authentication", "api key", "api_key",
                                                          "billing", "permission")):
                            self.err, self.disabled = msg[:300], True
                        r["haiku"] = {"flags": "?", "evidence": f"erro: {msg[:120]}"}
            self._count(part)
        return pend

    def _count(self, rs):
        for r in rs:
            if self.done(r) and not r["haiku"].get("_counted"):
                t = r["haiku"].get("_tokens") or [0, 0]
                self.tin += t[0]
                self.tout += t[1]
                self.n += 1
                r["haiku"]["_counted"] = True


def sheet_df(path, name):
    """Lê uma aba do Excel uma vez só (fica em cache até o arquivo mudar)."""
    return _sheet_cached(path, name, os.path.getmtime(path)).copy()


@st.cache_data(show_spinner=False, max_entries=24)
def _sheet_cached(path, name, mtime):
    ws = openpyxl.load_workbook(path, read_only=True)[name]
    data = list(ws.iter_rows(values_only=True))
    if not data:
        return pd.DataFrame()
    df = pd.DataFrame(data[1:], columns=data[0])
    for c in df.columns:
        if df[c].astype(str).str.startswith("=HYPERLINK").any():
            df[c] = df[c].map(link_of)
    return df.drop(columns=[c for c in ("sua revisão",) if c in df.columns])


def show_table(df, key, max_rows=2000):
    if len(df) > max_rows:
        st.caption(f"Mostrando as primeiras {max_rows} de {len(df)} linhas — o Excel tem todas.")
        df = df.head(max_rows)
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
    TYPES = {"ecommerce": "Produtos (e-commerce)", "finance": "Banco / empresa (golpes e phishing)",
             "mixed": "Misto (produtos + golpes)"}
    btype_label = st.radio("Tipo de marca", list(TYPES.values()),
                           index=list(TYPES).index(base.get("brand_type", "ecommerce")), key="bt" + k,
                           help="Produtos: lojas falsas, descontos, logo. Banco/empresa: sites e e-mails que se passam "
                                "pela empresa (phishing, falso investimento, falsa vaga), inclusive domínios sem site. "
                                "Misto: os dois — ex.: cerveja (lojas falsas + falsas promoções/sorteios).")
    btype = {v: k_ for k_, v in TYPES.items()}[btype_label]
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
    kc_txt = "\n\n".join(f"[{k}]\n" + "\n".join(v) for k, v in (base.get("keyword_context") or {}).items())
    kc_txt = st.text_area("Contexto específico por marca (opcional) — a marca só conta perto destes termos",
                          kc_txt, height=110, key="kc" + k,
                          help="Formato: [marca] numa linha e os termos embaixo, um por linha. Ex.:\n[pisane]\npea protein\n"
                               "protéine de pois\n\n[fibruline]\nchicory\ninulin")
    keyword_context, cur = {}, None
    for line in kc_txt.splitlines():
        line = line.strip()
        mk = re.match(r"^\[(.+)\]$", line)
        if mk:
            cur = mk.group(1).strip().lower()
            keyword_context.setdefault(cur, [])
        elif line and cur:
            keyword_context[cur] += to_list(line)
    c1, c2 = st.columns(2)
    min_discount = c1.number_input("Desconto mínimo (%) – A", 0, 95, int(base.get("min_discount", 30)), key="md" + k)
    price_alert = c2.number_input("Preço de alerta – A", 0.0, 100000.0, float(base.get("price_alert", 0)),
                                  help="0 = desligado. Produtos abaixo deste preço contam como A.", key="pa" + k)
    fin_fields = {}
    if btype in ("finance", "mixed"):
        fin_fields["known_bad_domains"] = to_list(st.text_area(
            "Domínios fraudulentos conhecidos (ex.: avisos da CSSF/AMF) — sempre prioridade 1",
            "\n".join(base.get("known_bad_domains", [])), height=90, key="kb" + k))
        with st.expander("Palavras de risco (avançado)"):
            fin_fields["risk_domain_words"] = to_list(st.text_area(
                "Palavras que, junto com a marca no domínio, indicam imitação do banco",
                "\n".join(base.get("risk_domain_words", ["bank", "banque", "private", "privee", "capital", "invest",
                    "wealth", "patrimoine", "finance", "fund", "fonds", "asset", "login", "log", "secure", "account",
                    "compte", "client", "online", "ebanking", "banking", "support", "verify", "service", "europe",
                    "group", "sa", "cp", "credit", "loan", "pret", "trading", "crypto", "my", "app", "portal", "auth"])),
                height=110, key="rw" + k))
            fin_fields["scam_terms"] = to_list(st.text_area(
                "Termos de golpe financeiro procurados na página",
                "\n".join(base.get("scam_terms", ["rendement garanti", "guaranteed return", "sans risque", "risk-free",
                    "bitcoin", "crypto", "forex", "whatsapp", "telegram", "prêt entre particulier", "crédit rapide",
                    "frais de dossier", "frais de déblocage", "release fee", "unlock your funds", "compte suspendu",
                    "verify your account", "récupération de fonds", "fund recovery"])), height=110, key="st" + k))
    with st.expander("Sinais de loja falsa (avançado)"):
        susp = to_list(st.text_area("Registradores suspeitos", "\n".join(base["suspicious_registrars"]), key="sr" + k))
        sale = to_list(st.text_area("Palavras de promoção (regex)", "\n".join(base["sale_words"]), key="sw" + k))
        recent = st.text_input("Domínio 'recente' a partir de (AAAA-MM-DD)", base.get("recent_from", "2025-01-01"),
                               key="rf" + k)
    criteria = st.text_area("Critérios para o Claude Haiku (opcional)", base.get("criteria", ""), height=160,
                            key="cr" + k)
    cfg = {**COMMON, "official_domain_regex": base.get("official_domain_regex", ""),
           "brand_type": btype, **fin_fields,
           "brand": brand, "search_term": search_term, **lists, "official_logos_b64": saved_b64,
           "official_logo_urls": logo_urls, "min_discount": int(min_discount),
           "price_alert": float(price_alert), "suspicious_registrars": susp, "sale_words": sale,
           "recent_from": recent, "criteria": criteria, "keyword_context": keyword_context,
           **({"strong_risk_words": base["strong_risk_words"]} if base.get("strong_risk_words") else {})}

st.subheader("2 · Lista de domínios e logo oficial")
c1, c2, c3 = st.columns([5, 4, 3])
up_xlsx = c1.file_uploader("Excel com a coluna **Domain** ou **URL** (web content). Registrar, Registrant, Created, "
                           "Title e Snippet são aproveitados se existirem",
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
batch = st.number_input("Tamanho do lote (domínios)", 100, 2000, 500, step=100,
                        help="Depois de cada lote o app salva tudo e mostra o resultado parcial e o cache para baixar. "
                             "Se der erro, envie o mesmo Excel + o último cache e ele continua de onde parou.")
with st.expander("🤖 Classificar os suspeitos com Claude Haiku durante a varredura (opcional)"):
    st.caption("A cada lote, o Haiku revisa só os suspeitos novos e o resultado já entra no Excel parcial. "
               "Precisa de chave de API **com crédito** (platform.claude.com); custo típico: centavos por centenas de "
               "suspeitos. Se a chave ou o crédito falharem, o Haiku é desligado e a varredura continua. "
               "A chave não é guardada.")
    hk1, hk2, hk3 = st.columns([2, 1, 1])
    haiku_key = hk1.text_input("Chave de API Anthropic", type="password", key="haiku_key")
    haiku_prios = hk2.multiselect("Prioridades a revisar", [1, 2, 3, 4], [1, 2, 3], key="haiku_prios")
    haiku_max = hk3.number_input("Gastar no máximo (US$)", 0.10, 100.0, 2.00, step=0.50, key="haiku_max",
                                 help="Ao chegar neste valor, o Haiku para e a varredura continua sem ele.")
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
    bar, info, box = st.progress(0.0), st.empty(), st.empty()
    import hashlib
    logo_sig = hashlib.md5(json.dumps([cfg.get("official_logos_b64"), cfg.get("official_logo_urls")]).encode()).hexdigest()[:10]

    def prog(done, total, el):
        bar.progress(min(done / max(total, 1), 1.0))
        info.write(f"{done}/{total} domínios neste passo · {el/60:.1f} min · faltam ~{el/done*(total-done)/60:.0f} min "
                   "— deixe esta aba aberta.")

    hr = None
    if haiku_key:
        try:
            hr = HaikuRunner(bs, haiku_key, haiku_prios or (1, 2, 3), max_cost=haiku_max)
        except Exception as e:
            st.warning(f"Haiku não iniciado: {str(e)[:200]}")

    def on_batch(phase, done, total, nbatch, nbatches, nres, nrows, counts, preview=()):
        ns_, npr_, nb_, _ = counts
        with box.container(border=True):
            hmsg = ""
            if hr:
                hmsg = (f" · 🤖 Haiku: {hr.n} revisados (≈ US$ {hr.cost:.2f})" if not hr.disabled
                        else " · ⚠️ Haiku desligado: " + (hr.err or "")[:120])
            st.markdown(f"💾 **Lote {nbatch}/{nbatches} ({phase}) salvo** — {nres} de {nrows} domínios com resultado · "
                        f"{ns_} suspeitos até agora{hmsg}. Pode baixar sem interromper:")
            d1, d2 = st.columns(2)
            d1.download_button("⬇ Resultado parcial (Excel)", open(bs.OUT_FILE, "rb").read(),
                               file_name=f"{bs.SLUG}_resultados_parcial_{nres}de{nrows}.xlsx",
                               key=f"pe_{phase}_{nbatch}", on_click="ignore", width="stretch")
            d2.download_button("⬇ Cache para continuar depois", open(bs.CACHE_FILE, "rb").read(),
                               file_name=f"cache_{bs.SLUG}.jsonl", key=f"pc_{phase}_{nbatch}",
                               on_click="ignore", width="stretch")
            if preview:
                st.caption(f"Prévia: prioridades 1 e 2 encontradas até agora ({len(preview)} primeiras)")
                st.dataframe(pd.DataFrame(preview), hide_index=True, width="stretch", height=260,
                             key=f"pv_{phase}_{nbatch}")

    if logos_only and not os.path.exists(bs.CACHE_FILE):
        st.error("Para reanalisar só os logos, envie também o cache da varredura anterior.")
        st.stop()
    results, (ns, npr, nb, cnt) = run_scan(bs, xlsx_path, 30 if test else 0, workers, retry, prog, logos_only,
                                           batch=int(batch), on_batch=on_batch, logo_sig=logo_sig, haiku=hr)
    bar.progress(1.0)
    info.success(f"Concluído: {ns} suspeitos, {npr} produtos com URL, {nb} bloqueados para revisão."
                 + (f" Haiku: {hr.n} revisados, ≈ US$ {hr.cost:.2f}." if hr and hr.n else ""))
    if hr and hr.disabled:
        st.warning(f"O Haiku foi desligado no meio do caminho: {hr.err}. A varredura terminou normalmente; "
                   "os suspeitos sem revisão podem ser enviados ao Haiku depois (seção 3).")
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
        PER = 30
        pages = max(1, (len(items) + PER - 1) // PER)
        page = st.number_input(f"Página (de {pages})", 1, pages, 1, key="logo_page") if pages > 1 else 1
        page_items = items[(page - 1) * PER: page * PER]
        for i in range(0, len(page_items), 3):
            cols = st.columns(3)
            for col, x in zip(cols, page_items[i:i + 3]):
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
                f1, f2 = st.columns([3, 2])
                pr = f1.multiselect("Prioridades", [1, 2, 3, 4], [1, 2, 3, 4], key="prf")
                df = df[df["prioridade"].isin(pr)]
                if "haiku" in df.columns and (df["haiku"].astype(str).str.lower() == "none").any():
                    if f2.toggle("Esconder os descartados pelo Haiku", value=False, key="hidenone"):
                        df = df[df["haiku"].astype(str).str.lower() != "none"]
            show_table(df, key="t_" + name) if not df.empty else st.write("Nada aqui.")

    c1, c2 = st.columns(2)
    c1.download_button("⬇ Baixar Excel do resultado", open(out, "rb").read(), file_name=os.path.basename(out),
                       width="stretch", type="primary")
    c2.download_button("⬇ Baixar cache (para continuar depois)", open(st.session_state["cache"], "rb").read(),
                       file_name=os.path.basename(st.session_state["cache"]), width="stretch")

    with st.expander("🤖 Revisar com Claude Haiku os suspeitos que ficaram sem revisão"):
        st.write("Precisa de uma chave de API **com crédito** (platform.claude.com). Custo típico: centavos de dólar. "
                 "A chave não é guardada.")
        key = st.text_input("Chave de API Anthropic", type="password")
        max2 = st.number_input("Gastar no máximo (US$)", 0.10, 100.0, 2.00, step=0.50, key="haiku_max2")
        if st.button("Classificar com Haiku", disabled=not key):
            bs = load_engine(st.session_state["cfg"])
            cache = bs.load_cache()
            results = [cache[r["domain"]] for r in bs.read_domains(st.session_state["xlsx"]) if r["domain"] in cache]
            ok = False
            with st.spinner("Haiku revisando os suspeitos que ainda não foram revisados…"):
                try:
                    hr2 = HaikuRunner(bs, key, max_cost=max2)
                except Exception as e:
                    st.error(f"Haiku não iniciado: {str(e)[:200]}")
                    st.stop()
                for i in range(0, len(results), 300):   # em blocos, gravando a cada bloco
                    changed = hr2.run(results[i:i + 300])
                    if changed:
                        save_cache(bs, cache)   # results aponta para os mesmos registros do cache
                    if hr2.disabled:
                        break
            if hr2.n:
                ns, npr, nb, cnt = bs.write_excel(results)
                st.session_state.update(counts=dict(cnt), totals=(ns, npr, nb))
                st.toast(f"Haiku revisou {hr2.n} suspeitos · custo ≈ US$ {hr2.cost:.2f}")
                ok = True
            if hr2.disabled:
                st.error(f"O Haiku parou: {hr2.err}. Verifique a chave e o crédito.")
            elif not hr2.n:
                st.info("Nenhum suspeito pendente de revisão.")
            if ok:
                st.rerun()
