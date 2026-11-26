# Proteção de marca — app de varredura de domínios

App Streamlit que analisa listas de domínios e encontra sites que usam uma marca indevidamente:
lojas falsas, produtos não licenciados, logo copiado, sites que se dizem oficiais, phishing e golpes.

## Arquivos do repositório (só estes 4)

| Arquivo | Para que serve |
|---|---|
| `app.py` | A tela do app (marcas pré-configuradas, envio de arquivos, resultados) |
| `brand_scan.py` | O motor da varredura (abre os sites, aplica os filtros, gera o Excel) |
| `requirements.txt` | Componentes que o Streamlit instala |
| `README.md` | Este guia |

Opcional: arquivos `marca_<nome>.json` (gerados pelo botão **💾 Salvar esta marca**). Cada um que
estiver no repositório aparece na lista **"Partir de"**, já com critérios e logo.

### Atualizar o app

1. Apague da pasta Downloads as versões antigas dos arquivos **antes** de baixar as novas
   (senão o navegador salva como `app (1).py` e o app continua usando o antigo).
2. GitHub → **Add file → Upload files** → arraste os arquivos → **Commit changes**.
   Os nomes precisam ser exatamente os da tabela acima.
3. Streamlit → **Manage app → ⋮ → Reboot app** e aguarde 3–5 minutos.

Se aparecer erro: **Manage app** → role o registro até o fim e copie as últimas linhas.

## Como usar

1. **Barra lateral:** escolha a marca em "Partir de" (ou "Nova marca") e confira os critérios.
2. **Seção 2:** envie o Excel com a coluna `Domain` (as colunas Registrar, Registrant e Created
   são aproveitadas se existirem) e o **logo oficial** (pode enviar várias versões).
3. Opcional: abra **🤖 Classificar com Claude Haiku**, cole a chave de API e defina o
   **"Gastar no máximo (US$)"**.
4. Comece com **Só testar com 30 domínios**. Depois rode a lista completa.
5. A cada lote (500 domínios por padrão) o app salva tudo e mostra:
   prévia dos suspeitos, **Excel parcial** e **cache** para baixar — sem interromper.
6. Se o app cair: envie **o mesmo Excel + o último cache** e clique em Iniciar — ele continua
   de onde parou (sem repetir o Haiku nos já revisados).

Dicas: lotes de 300–500 e 12 conexões são o mais estável. Sites bloqueados vão para a aba
"Bloqueados", com links para revisão manual; a opção **Tentar de novo os bloqueados** ajuda em parte.

## Tipos de marca e critérios

| Tipo | Quando usar | Exemplos |
|---|---|---|
| **Produtos (e-commerce)** | Marcas de produto: lojas falsas, descontos, logo | Gertrude, ASSOULINE, Nestlé |
| **Banco / empresa (golpes e phishing)** | Bancos e empresas: phishing, falso investimento, falsa vaga, domínios sem site usados em e-mails | Quintet |
| **Misto** | Os dois juntos | Air Liquide (+ Seppic, VitalAire), Tiger Beer, Bir Bintang |

| Letra | Significado |
|---|---|
| A | Produto da marca com desconto ≥ X% ou abaixo do preço de alerta |
| B | Nome/logo da marca na página inicial (inclui comparação visual com o logo oficial) |
| C | Modelos/produtos-chave ou entidades do grupo citados |
| OF | Site que se diz "oficial" e não está na lista oficial (nunca é descartado por isso) |
| W | Produto vigiado/não licenciado com a marca (ex.: vape, camiseta) — suspeito mesmo a preço normal |
| K | Domínio na lista de fraudes conhecidas (ex.: avisos CSSF/AMF) |
| D | Domínio que imita a marca (marca + private/login/finance/promo…) |
| M | Domínio com e-mail ativo (MX), mesmo sem site |
| L | Página de login/senha com a marca fora dos sites oficiais |
| F | Termos de golpe (rendimento garantido, sorteio, taxa de liberação, WhatsApp…) |

Prioridade 1 = mais grave. Prioridade 4 = pode ser revendedor legítimo (verificar autorização).

## Checklist para um novo prospect

Antes de rodar, reúna (o site oficial e os avisos de fraude do próprio cliente ajudam muito):

1. **Tipo de marca** — produtos, banco/empresa ou misto (ver tabela acima). Pergunte ao cliente
   o que está no escopo: lojas falsas? phishing? falsas vagas? produtos não licenciados?
2. **Como o nome aparece** — variações com e sem acento/espaço (ex.: `nestlé`, `nestle`;
   `air liquide`, `airliquide`). Se o nome for uma palavra comum (Tiger, Bintang, Quintet),
   use só formas compostas (`tiger beer`, `bir bintang`) e o campo **Nome "solto"** apenas
   junto com produtos vigiados.
3. **Sites oficiais** — site principal, sites por país, portais de cliente/login, vagas,
   investidores, sites das marcas do grupo que estão no escopo.
4. **Titulares e registradores oficiais** — nome da empresa no Registrant (ex.: Société des
   Produits Nestlé S.A.) e o gestor de domínios da marca (ex.: Com Laude, MarkMonitor, CSC).
   Domínios registrados por eles são tratados como oficiais.
5. **Logo oficial** — PNG/JPG do topo do site oficial (botão direito → salvar imagem, ou print
   recortado). Inclua versões branca/colorida e logos das sub-marcas no escopo.
6. **Termos de contexto** — palavras que provam que a menção é da marca certa e eliminam
   homônimos (ex.: banco → `private bank`, `patrimoine`; cerveja → `beer`, `lager`).
7. **Produtos/modelos-chave** — nomes de produtos, coleções, entidades do grupo (critério C).
8. **Produtos vigiados** (se o cliente pedir) — prioridade 1 e 2 (ex.: vapes P1, roupas P2).
9. **Preço** — desconto mínimo que indica falsificação (30–60%) e, se houver faixa de preço
   clara, o preço de alerta e os produtos a que ele se aplica. Para bancos/empresas: 0 (desligado).
10. **Golpes** (banco/empresa/misto) — domínios fraudulentos já conhecidos (CSSF, AMF, avisos
    do cliente), palavras de risco no domínio e termos de golpe típicos do setor/idioma.
11. **Revendedores autorizados** — domínios de parceiros que não devem aparecer como suspeitos.
12. **Critérios para o Haiku** — descreva em português o que conta e o que deve ser descartado.
13. Rode um teste com 30 domínios, ajuste, e clique em **💾 Salvar esta marca**. Para deixar a
    marca fixa na lista, envie o `marca_<nome>.json` ao GitHub.

## Claude Haiku (opcional)

- Revisa só os suspeitos (prioridades escolhidas, por padrão 1–3), uma vez por domínio.
- Custo típico: ~US$ 0,0002–0,0003 por suspeito (centavos por centenas de suspeitos).
- Para nunca gastar além do previsto: use o campo **Gastar no máximo (US$)** (pode passar uns
  centavos, por causa das chamadas já em andamento), crédito pré-pago sem recarga automática e
  um limite de gasto no Console (platform.claude.com).
- Sem crédito ou com chave errada, o Haiku se desliga e a varredura continua.

## Marcas pré-configuradas

Gertrude · ASSOULINE · Nestlé · Quintet · Air Liquide · Tiger Beer · Bir Bintang
