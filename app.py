"""
TeraAg — Monitor do crédito privado agro (MVP)

Lê a planilha carteira_ativos.xlsx do repositório (ou uma planilha enviada pela barra lateral)
e mostra as carteiras dos fundos: volumes, taxas praticadas, concentração por devedor,
comparação entre fundos e a carteira completa para exportar.

Colunas obrigatórias: Fundo, Ativo, Valor.
Colunas usadas quando existirem: Codigo_Ativo, Indexador, Cupom, Segmento, Vencimento, Duration, Rating, Mes.
"""
from __future__ import annotations

import io
import re
import unicodedata

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Monitor de Crédito - Taxas e Volume Financeiro", page_icon="🌾", layout="wide")

ARQUIVO = "carteira_ativos.xlsx"


def _eh_celular() -> bool:
    try:
        ua = st.context.headers.get("User-Agent", "") or ""
    except Exception:
        ua = ""
    return any(k in ua for k in ("iPhone", "Android", "Mobile", "iPad"))


COMPACTO = _eh_celular()

GREEN, MID, GOLD, INK, MUTED, LINE, RED = "#123524", "#5F8B7A", "#A8721A", "#1E2B26", "#6B776F", "#D6DED9", "#9B1C1C"
PALETA = [GREEN, MID, GOLD, "#4F63A0", "#8C4F3A", "#9DBBA9", "#6E7A72", "#C9A24A", "#2F5A3E"]

st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=Source+Serif+4:opsz,wght@8..60,600&display=swap');
html, body, .stApp, .stMarkdown, button, input, select {{ font-family: 'IBM Plex Sans', 'Segoe UI', system-ui, sans-serif; }}
.stApp {{ background: #F6F8F6; }}
.block-container {{ padding-top: 2.6rem; max-width: 1400px; }}
.tg-head {{ display: flex; align-items: baseline; gap: 16px; flex-wrap: wrap; border-bottom: 1px solid {LINE}; padding-bottom: 12px; margin-bottom: 18px; }}
.tg-logo {{ font-family: 'Source Serif 4', Georgia, serif; font-weight: 600; font-size: 30px; color: {GREEN}; line-height: 1; }}
.tg-sub {{ color: {MUTED}; font-size: 15px; }}
.tg-ref {{ margin-left: auto; color: {MUTED}; font-size: 13px; }}
.kpis {{ display: grid; grid-template-columns: repeat(6, minmax(0, 1fr)); border-top: 2px solid {GREEN}; border-bottom: 1px solid {LINE}; margin: 4px 0 18px; }}
.kpi {{ padding: 12px 14px 12px 0; min-width: 0; }}
.kpi .l {{ font-size: 12px; color: {MUTED}; }}
.kpi .v {{ font-size: 22px; font-weight: 600; color: {INK}; font-variant-numeric: tabular-nums; line-height: 1.3; white-space: nowrap; }}
.kpi .n {{ font-size: 12px; color: {MUTED}; }}
.kpi .v.alert {{ color: {RED}; }}
@media (max-width: 1000px) {{ .kpis {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }} }}
.tg-t {{ font-size: 30px; font-weight: 600; color: {GREEN}; line-height: 1.2; }}
.tg-m {{ font-size: 14px; color: {MUTED}; margin-top: 6px; }}
@media (max-width: 640px) {{
  .block-container {{ padding: 2.6rem 0.9rem 2rem; }}
  .tg-t {{ font-size: 21px; }}
  .tg-m {{ font-size: 12.5px; }}
  .kpi {{ padding: 9px 8px 9px 0; }}
  .kpi .v {{ font-size: 17px; white-space: normal; }}
  .kpi .l, .kpi .n {{ font-size: 11px; }}
  .lead {{ font-size: 14.5px; line-height: 1.6; }}
  .figt {{ font-size: 14px; }}
  .figs, .note {{ font-size: 12px; }}
  .h2 {{ font-size: 17px; margin-top: 20px; }}
  div[data-testid="stTabs"] button p {{ font-size: 13.5px; }}
}}
.lead {{ font-size: 16px; line-height: 1.65; color: {INK}; max-width: 900px; margin: 0 0 6px; }}
.fig {{ font-size: 11.5px; font-weight: 700; color: {MID}; margin-top: 14px; }}
.figt {{ font-size: 15px; font-weight: 600; color: {INK}; }}
.figs {{ font-size: 12.5px; color: {MUTED}; margin-bottom: 2px; }}
.src {{ font-size: 11.5px; color: {MUTED}; margin: -4px 0 14px; }}
.note {{ font-size: 13px; color: {MUTED}; line-height: 1.55; }}
.h2 {{ font-size: 19px; font-weight: 600; color: {GREEN}; margin: 26px 0 6px; padding-top: 14px; border-top: 1px solid {LINE}; }}
div[data-testid="stTabs"] button p {{ font-size: 15px; font-weight: 500; }}
div[data-testid="stDataFrame"] {{ border-top: 2px solid {GREEN}; }}
</style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def norm(s) -> str:
    s = unicodedata.normalize("NFD", "" if s is None else str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", s.lower())


def num_br(v, dec=0) -> str:
    try:
        if v is None or pd.isna(v):
            return "—"
    except (TypeError, ValueError):
        return "—"
    return f"{float(v):,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def brl(v) -> str:
    if v is None or pd.isna(v):
        return "—"
    a = abs(v)
    if a >= 1e9:
        return f"R$ {num_br(v / 1e9, 2)} bi"
    if a >= 1e6:
        return f"R$ {num_br(v / 1e6, 1)} mi"
    if a >= 1e3:
        return f"R$ {num_br(v / 1e3, 0)} mil"
    return f"R$ {num_br(v, 0)}"


def taxa_txt(v, idx="CDI+") -> str:
    if v is None or pd.isna(v):
        return "—"
    return {"CDI+": f"CDI + {num_br(v, 2)}%", "IPCA+": f"IPCA + {num_br(v, 2)}%", "%CDI": f"{num_br(v, 1)}% do CDI",
            "Pré": f"{num_br(v, 2)}% a.a."}.get(idx, f"{num_br(v, 2)}%")


def para_numero(v) -> float:
    """Aceita 99055355.7, '99.055.355,70', 'R$ 1.234,56', '3,80%'."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return np.nan
    if isinstance(v, (int, float, np.number)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v))
    if not s:
        return np.nan
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return np.nan


def indexador_de(txt) -> str:
    n = norm(txt)
    if not n or n in ("na", "nan"):
        return "Sem indexador"
    if "ipca" in n:
        return "IPCA+"
    if "pct" in n or "%" in str(txt) and "cdi" in n:
        return "%CDI"
    if "cdi" in n or n == "di":
        return "CDI+"
    if "pre" in n:
        return "Pré"
    return str(txt).strip()


PREFIXOS = [("cprf", "CPR-F"), ("cpr", "CPR-F"), ("cra", "CRA"), ("cri", "CRI"), ("debenture", "Debênture"), ("fidc", "Cotas de FIDC"),
            ("fiagro", "Cotas de Fiagro"), ("cdca", "CDCA"), ("nc", "Nota comercial"), ("notacomercial", "Nota comercial")]


def instrumento_de(ativo) -> str:
    primeira = norm(str(ativo).split(" ")[0])
    for pre, nome in PREFIXOS:
        if primeira == pre:
            return nome
    return "Outros"


def devedor_de(ativo) -> str:
    """'CRA Languiru II 2S' -> 'Languiru'; 'FIDC Agroforte - Sênior' -> 'Agroforte'."""
    d = re.sub(r"^(CRA|CRI|CPR-F|CPRF|CPR|Debênture|Debenture|FIDC|FIAGRO|CDCA|NC)\s+", "", str(ativo).strip(), flags=re.I)
    d = re.split(r"\s+-\s+", d)[0]
    d = re.sub(r"(\s+(I{1,3}|IV|V|VI{0,3}|IX|X|\d+[ªa]?\s*S(érie)?|\d+S|S[eê]nior|Mezanino|Sub(ordinada)?))+$", "", d, flags=re.I)
    return d.strip() or str(ativo)


COLUNAS = {
    "fundo": ["fundo", "carteira", "veiculo", "ticker"],
    "ativo": ["ativo", "nomedoativo", "papel", "operacao"],
    "codigo": ["codigoativo", "codigo", "cetip", "codigob3", "isin"],
    "valor": ["valor", "valorrs", "posicao", "saldo", "financeiro"],
    "indexador": ["indexador", "index", "remuneracao"],
    "taxa": ["cupom", "coupon", "taxa", "spread"],
    "segmento": ["segmento", "setor"],
    "vencimento": ["vencimento", "datadevencimento"],
    "duration": ["duration", "durationanos", "prazo"],
    "rating": ["rating"],
    "mes": ["mes", "mesreferencia", "data", "database", "referencia"],
}


def mapear_colunas(cols) -> dict:
    out = {}
    for campo, chaves in COLUNAS.items():
        for c in cols:
            if norm(c) in chaves and c not in out.values():
                out[campo] = c
                break
    return out


@st.cache_data(show_spinner=False)
def ler_planilha(conteudo: bytes | None) -> pd.DataFrame:
    if conteudo is None:
        return pd.read_excel(ARQUIVO)
    return pd.read_excel(io.BytesIO(conteudo))


def preparar(bruto: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    mp = mapear_colunas(bruto.columns)
    faltam = [c for c in ["fundo", "ativo", "valor"] if c not in mp]
    if faltam:
        raise ValueError("A planilha precisa das colunas Fundo, Ativo e Valor. Não encontrei: " + ", ".join(faltam) + ".")
    df = pd.DataFrame({
        "Fundo": bruto[mp["fundo"]].astype(str).str.strip(),
        "Ativo": bruto[mp["ativo"]].astype(str).str.strip(),
        "Codigo": bruto[mp["codigo"]].astype(str).str.strip().replace({"nan": "", "-": "", "None": ""}) if "codigo" in mp else "",
        "Valor": bruto[mp["valor"]].map(para_numero),
    })
    df["Indexador"] = bruto[mp["indexador"]].map(indexador_de) if "indexador" in mp else "Sem indexador"
    if "taxa" in mp:
        t = bruto[mp["taxa"]].map(para_numero)
        # Excel guarda 3,80% como 0,038: se a mediana estiver abaixo de 1, a coluna está em fração
        if t.dropna().abs().median() < 1:
            t = t * 100
        df["Taxa"] = t
    else:
        df["Taxa"] = np.nan
    df.loc[df["Indexador"] == "Sem indexador", "Taxa"] = np.nan
    # "CDI" com taxa acima de 30 é percentual do CDI (ex.: 104% do CDI), não spread
    df.loc[(df["Indexador"] == "CDI+") & (df["Taxa"] > 30), "Indexador"] = "%CDI"
    df["Instrumento"] = df["Ativo"].map(instrumento_de)
    df["Devedor"] = df["Ativo"].map(devedor_de)
    df["Segmento"] = bruto[mp["segmento"]].astype(str).str.strip() if "segmento" in mp else ""
    df["Vencimento"] = pd.to_datetime(bruto[mp["vencimento"]], dayfirst=True, errors="coerce") if "vencimento" in mp else pd.NaT
    df["Duration"] = bruto[mp["duration"]].map(para_numero) if "duration" in mp else np.nan
    df["Rating"] = bruto[mp["rating"]].astype(str).replace({"nan": ""}) if "rating" in mp else ""
    df["Mes"] = bruto[mp["mes"]].astype(str).str[:7] if "mes" in mp else ""
    df = df[(df["Ativo"] != "") & (df["Ativo"].str.lower() != "nan") & df["Valor"].notna() & (df["Valor"] > 0)]
    df = df[~df["Ativo"].map(norm).str.startswith(("total", "caixa", "patrimonio", "rendimentos"))]
    return df.reset_index(drop=True), mp


def media_pond(d: pd.DataFrame, col="Taxa"):
    x = d[d[col].notna()]
    return float((x[col] * x["Valor"]).sum() / x["Valor"].sum()) if len(x) and x["Valor"].sum() > 0 else None


def dispersao(d: pd.DataFrame, por: str) -> pd.DataFrame:
    out = []
    for g, x in d[d["Taxa"].notna()].groupby(por):
        out.append({por: g, "Ativos": len(x), "Volume": x["Valor"].sum(), "Mínimo": x["Taxa"].min(), "Média ponderada": media_pond(x),
                    "Mediana": x["Taxa"].median(), "Máximo": x["Taxa"].max()})
    return pd.DataFrame(out).sort_values("Volume", ascending=False) if out else pd.DataFrame()


def operacoes(d: pd.DataFrame) -> pd.DataFrame:
    """Agrupa o mesmo ativo em todos os fundos: pelo código; sem código, pelo nome normalizado."""
    x = d.copy()
    x["Chave"] = np.where(x["Codigo"].astype(str).str.len() > 3, x["Codigo"].astype(str).str.upper(), x["Ativo"].map(norm))
    linhas = []
    for k, g in x.groupby("Chave"):
        idx = g.groupby("Indexador")["Valor"].sum().idxmax()
        gi = g[g["Indexador"] == idx]
        linhas.append({"Chave": k, "Ativo": g.loc[g["Valor"].idxmax(), "Ativo"], "Devedor": g["Devedor"].iloc[0], "Codigo": g["Codigo"].iloc[0],
                       "Instrumento": g["Instrumento"].iloc[0], "Indexador": idx, "Taxa": media_pond(gi),
                       "Volume": g["Valor"].sum(), "Nº fundos": g["Fundo"].nunique(), "Fundos": ", ".join(sorted(g["Fundo"].unique()))})
    return pd.DataFrame(linhas)


def rank_view(r: pd.DataFrame, total: float, curto=False, com_fundo=False) -> pd.DataFrame:
    v = pd.DataFrame({"#": r["#"], "Operação": r["Ativo"], "Instrumento": r["Instrumento"],
                      "Taxa": [taxa_txt(t, i) for t, i in zip(r["Taxa"], r["Indexador"])],
                      "Volume": (r["Volume"] / 1e6).round(1), "% do total": (r["Volume"] / total * 100).round(2), "Fundos": r["Fundos"]})
    if curto:
        v = v.drop(columns=["Instrumento", "% do total", "Fundos"])
        v["Operação"] = v["Operação"].str.replace(r"^(CRA|CRI|CPR-F|CPR|Debênture|FIDC|FIAGRO|CDCA|NC)\s+", "", regex=True)
    if com_fundo:
        # rankings de taxa: o fundo que detém a operação no lugar da numeração
        v = v.drop(columns=[c for c in ["#", "Instrumento", "% do total", "Fundos"] if c in v.columns])
        v.insert(0, "Fundo", r["Fundos"].values)
    return v


def RANK_CFG(total, vmax, idx=None):
    return {
        "#": st.column_config.NumberColumn("#", width=32 if COMPACTO else 40),
        "Operação": st.column_config.TextColumn("Operação", width=135 if COMPACTO else "medium"),
        "Taxa": st.column_config.TextColumn("Taxa", width=95 if COMPACTO else None),
        "Volume": st.column_config.NumberColumn("R$ mi" if COMPACTO else "Volume (R$ mi)", format="localized", width=60 if COMPACTO else None),
        "% do total": st.column_config.NumberColumn("% do total", format="%.2f%%"),
        "Fundos": st.column_config.TextColumn("Fundos", width="medium"),
        "Fundo": st.column_config.TextColumn("Fundo", width=80 if COMPACTO else "small"),
    }


def hhi(valores: pd.Series) -> float:
    s = valores / valores.sum()
    return float((s ** 2).sum() * 10000)


def estilo(fig, altura=380, legenda=True):
    if COMPACTO:
        altura = max(260, int(altura * 0.75))
    fig.update_layout(
        template="simple_white", height=altura, separators=",.", margin=dict(l=4, r=8, t=10, b=4) if COMPACTO else dict(l=8, r=16, t=10, b=8),
        font=dict(family="IBM Plex Sans, Segoe UI, sans-serif", size=11 if COMPACTO else 13, color=INK), dragmode=False,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=legenda,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title=None),
        hoverlabel=dict(bgcolor="white", font_size=13),
    )
    fig.update_xaxes(showgrid=False, linecolor=LINE, tickfont=dict(color=MUTED), fixedrange=True)
    fig.update_yaxes(showgrid=True, gridcolor="#E8EDEA", linecolor=LINE, tickfont=dict(color=MUTED), fixedrange=True)
    return fig


def mostrar(fig):
    try:
        st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})
    except TypeError:
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def legenda_fig(n, titulo, sub=""):
    st.markdown(f'<div class="fig">Figura {n}</div><div class="figt">{titulo}</div>' + (f'<div class="figs">{sub}</div>' if sub else ""), unsafe_allow_html=True)


def fonte(txt="carteiras informadas nos relatórios de gestão dos fundos."):
    st.markdown(f'<div class="src">Fonte: {txt}</div>', unsafe_allow_html=True)


def tabela(df, **kw):
    try:
        st.dataframe(df, hide_index=True, width="stretch", **kw)
    except TypeError:
        st.dataframe(df, hide_index=True, use_container_width=True, **kw)


# ---------------------------------------------------------------------------
# Dados
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### Base de dados")
    origem = st.radio("Origem", ["Planilha do repositório", "Enviar outra planilha"], label_visibility="collapsed")
    conteudo = None
    if origem == "Enviar outra planilha":
        up = st.file_uploader("Planilha (.xlsx)", type=["xlsx", "xls"])
        if up is None:
            st.info("Envie uma planilha com as colunas Fundo, Ativo e Valor.")
            st.stop()
        conteudo = up.getvalue()

try:
    df_all, mapa = preparar(ler_planilha(conteudo))
except FileNotFoundError:
    st.error(f"Não encontrei o arquivo {ARQUIVO} no repositório. Confira o nome do arquivo ou envie a planilha pela barra lateral.")
    st.stop()
except Exception as e:
    st.error(f"Não foi possível ler a planilha: {e}")
    st.stop()

if df_all.empty:
    st.warning("A planilha não tem linhas com Fundo, Ativo e Valor preenchidos.")
    st.stop()

# ---------------------------------------------------------------------------
# Filtros
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### Filtros")
    fundos_disp = sorted(df_all["Fundo"].unique())
    fundos = st.multiselect("Fundos", fundos_disp, default=fundos_disp)
    idx_disp = [i for i in ["CDI+", "IPCA+", "%CDI", "Pré"] if i in set(df_all["Indexador"])] + sorted(set(df_all["Indexador"]) - {"CDI+", "IPCA+", "%CDI", "Pré"})
    idxs = st.multiselect("Indexador", idx_disp, default=idx_disp)
    inst_disp = sorted(df_all["Instrumento"].unique())
    insts = st.multiselect("Instrumento", inst_disp, default=inst_disp)
    segs = []
    if (df_all["Segmento"] != "").any():
        seg_disp = sorted(s for s in df_all["Segmento"].unique() if s)
        segs = st.multiselect("Segmento", seg_disp, default=seg_disp)
    busca = st.text_input("Buscar ativo, devedor ou código")
    COMPACTO = st.toggle("Modo celular (tabelas e gráficos enxutos)", value=COMPACTO, help="Ligado automaticamente quando o app é aberto no celular.")
    st.markdown("---")
    with st.expander("Como enriquecer a planilha"):
        st.markdown(
            "Além de **Fundo, Ativo, Codigo_Ativo, Valor, Indexador e Cupom**, o painel usa, se existirem, as colunas "
            "**Segmento** (produtor, cooperativa, distribuidora…), **Vencimento**, **Duration**, **Rating** e **Mes** "
            "(mês de referência da carteira, AAAA-MM). Com elas aparecem as análises por setor, por prazo e a evolução mensal."
        )

df = df_all[df_all["Fundo"].isin(fundos) & df_all["Indexador"].isin(idxs) & df_all["Instrumento"].isin(insts)]
if segs:
    df = df[df["Segmento"].isin(segs) | (df["Segmento"] == "")]
if busca:
    nb = norm(busca)
    df = df[df["Ativo"].map(norm).str.contains(nb) | df["Devedor"].map(norm).str.contains(nb) | df["Codigo"].map(norm).str.contains(nb)]

tem_seg = (df["Segmento"] != "").any()
tem_venc = df["Vencimento"].notna().any() or df["Duration"].notna().any()

# ---------------------------------------------------------------------------
# Cabeçalho, indicadores e resumo
# ---------------------------------------------------------------------------
st.markdown(
    f'<div style="border-bottom:1px solid {LINE}; padding-bottom:14px; margin-bottom:20px;">'
    f'<div class="tg-t">Monitor de Crédito - Taxas e Volume Financeiro</div>'
    f'<div class="tg-m">TeraAg &nbsp;|&nbsp; {df_all["Fundo"].nunique()} fundos e {len(df_all)} ativos na base</div>'
    '</div>', unsafe_allow_html=True)

if df.empty:
    st.info("Nenhum ativo para os filtros escolhidos.")
    st.stop()

total = df["Valor"].sum()
cdi = df[df["Indexador"] == "CDI+"]
ipca = df[df["Indexador"] == "IPCA+"]
por_dev = df.groupby("Devedor")["Valor"].sum().sort_values(ascending=False)
top10 = por_dev.head(10).sum() / total * 100
maior_dev, maior_pct = por_dev.index[0], por_dev.iloc[0] / total * 100
cruzados = df.groupby("Devedor")["Fundo"].nunique()
n_cruz = int((cruzados > 1).sum())

kpis = [
    ("Volume em carteira", brl(total), f"{df['Fundo'].nunique()} fundo(s)"),
    ("Ativos", str(len(df)), f"{df['Devedor'].nunique()} devedores"),
    ("Taxa média em CDI+", taxa_txt(media_pond(cdi)), f"{len(cdi)} ativos, ponderada por valor"),
    ("Taxa média em IPCA+", taxa_txt(media_pond(ipca), "IPCA+"), f"{len(ipca)} ativos" if len(ipca) else "Nenhum ativo"),
    ("10 maiores devedores", f"{num_br(top10, 1)}%", "do volume filtrado"),
    ("Devedores em mais de um fundo", str(n_cruz), "mesmo risco em fundos diferentes" if n_cruz else "Nenhum"),
]
st.markdown('<div class="kpis">' + "".join(f'<div class="kpi"><div class="l">{l}</div><div class="v">{v}</div><div class="n">{n}</div></div>' for l, v, n in kpis) + "</div>",
            unsafe_allow_html=True)

disp_cdi = dispersao(cdi, "Fundo")
frases = [f"Os {df['Fundo'].nunique()} fundo(s) filtrados somam {brl(total)} em {len(df)} ativos de {df['Devedor'].nunique()} devedores."]
if len(cdi):
    frases.append(f"Em CDI+, as taxas vão de {taxa_txt(cdi['Taxa'].min())} a {taxa_txt(cdi['Taxa'].max())}, com média ponderada de {taxa_txt(media_pond(cdi))}.")
frases.append(f"O maior devedor é {maior_dev}, com {num_br(maior_pct, 1)}% do volume; os dez maiores somam {num_br(top10, 0)}%.")
if n_cruz:
    frases.append(f"{n_cruz} devedor(es) aparecem em mais de um fundo.")
st.markdown(f'<p class="lead">{" ".join(frases)}</p>', unsafe_allow_html=True)
if COMPACTO:
    st.caption("Filtros de fundo, indexador e instrumento ficam no menu (seta › no canto superior esquerdo).")

t1, t2, t3, t4, t5 = st.tabs(["Visão geral", "Taxas", "Concentração", "Comparar fundos", "Carteira"])

# ---------------------------------------------------------------------------
# Visão geral
# ---------------------------------------------------------------------------
with t1:
    c1, c2 = st.columns([1, 1], gap="large")
    with c1:
        legenda_fig(1, "Volume por fundo", "R$ milhões")
        vf = df.groupby("Fundo", as_index=False)["Valor"].sum().sort_values("Valor")
        vf["mi"] = vf["Valor"] / 1e6
        vf["rot"] = vf.apply(lambda r: f"{num_br(r['mi'], 1)}  ({num_br(r['Valor'] / total * 100, 0)}%)", axis=1)
        fig = px.bar(vf, x="mi", y="Fundo", orientation="h", text="rot", color_discrete_sequence=[GREEN])
        fig.update_traces(textposition="outside", cliponaxis=False, hovertemplate="%{y}<br>R$ %{x:,.1f} mi<extra></extra>")
        fig.update_xaxes(title=None, showticklabels=False, range=[0, vf["mi"].max() * 1.35])
        fig.update_yaxes(title=None, showgrid=False)
        mostrar(estilo(fig, max(220, 60 + 42 * len(vf)), legenda=False))
        fonte()
    with c2:
        legenda_fig(2, "Composição por instrumento", "Participação no volume de cada fundo")
        ci = df.groupby(["Fundo", "Instrumento"], as_index=False)["Valor"].sum()
        ci["pct"] = ci["Valor"] / ci.groupby("Fundo")["Valor"].transform("sum") * 100
        fig = px.bar(ci, x="pct", y="Fundo", color="Instrumento", orientation="h", color_discrete_sequence=PALETA,
                     custom_data=["Instrumento", "Valor"])
        fig.update_traces(hovertemplate="%{customdata[0]}: %{x:.1f}%<extra></extra>")
        fig.update_xaxes(title=None, ticksuffix="%", range=[0, 100])
        fig.update_yaxes(title=None, showgrid=False)
        fig.update_layout(barmode="stack")
        mostrar(estilo(fig, max(220, 90 + 42 * df["Fundo"].nunique())))
        fonte()

    # ---------------- Principais operações: rankings por volume e por taxa ----------------
    st.markdown('<div class="h2">Principais operações</div>', unsafe_allow_html=True)
    st.markdown('<p class="note">Cada operação reúne o mesmo ativo em todos os fundos filtrados (pelo código; sem código, pelo nome). '
                'Volume é a soma das posições; taxa é a média ponderada pelo volume.</p>', unsafe_allow_html=True)
    ops = operacoes(df)
    n_rank = st.select_slider("Tamanho dos rankings", options=[5, 10, 15, 20, 30], value=10 if COMPACTO else 15)

    st.markdown(f'<div class="figt">Maiores operações por volume</div><div class="figs">Top {n_rank}, todas as taxas e indexadores</div>', unsafe_allow_html=True)
    rv = ops.sort_values("Volume", ascending=False).head(n_rank).reset_index(drop=True)
    rv.insert(0, "#", range(1, len(rv) + 1))
    tabela(rank_view(rv, total, curto=COMPACTO), column_config=RANK_CFG(total, rv["Volume"].max()), height=38 + 35 * len(rv))
    fonte()

    idx_rank = [i for i in ["CDI+", "IPCA+", "%CDI", "Pré"] if (ops["Indexador"] == i).any()]
    if idx_rank:
        ir = st.radio("Indexador dos rankings de taxa", idx_rank, horizontal=True)
        base = ops[(ops["Indexador"] == ir) & ops["Taxa"].notna()]
        c1, c2 = st.columns(2, gap="large")
        for col, titulo, asc in [(c1, "Maiores taxas", False), (c2, "Menores taxas", True)]:
            with col:
                rt = base.sort_values(["Taxa", "Volume"], ascending=[asc, False]).head(n_rank).reset_index(drop=True)
                rt.insert(0, "#", range(1, len(rt) + 1))
                st.markdown(f'<div class="figt">{titulo} em {ir}</div><div class="figs">Top {min(n_rank, len(rt))} operações</div>', unsafe_allow_html=True)
                tabela(rank_view(rt, total, curto=True, com_fundo=True), column_config=RANK_CFG(total, max(1, base["Volume"].max()), idx=ir), height=38 + 35 * len(rt))
        fonte()

        legenda_fig(3, f"Volume e taxa das operações em {ir}", "Cada ponto é uma operação; as 10 maiores por volume estão identificadas. Linha tracejada: mediana da taxa")
        sc = base.copy()
        sc["mi"] = sc["Volume"] / 1e6
        top = set(sc.nlargest(5 if COMPACTO else 10, "Volume")["Chave"])
        sc["rotulo"] = np.where(sc["Chave"].isin(top), sc["Ativo"].str.replace(r"^(CRA|CRI|CPR-F|CPR|Debênture|FIDC|FIAGRO|CDCA|NC)\s+", "", regex=True), "")
        cores = {"CRA": GREEN, "CPR-F": GOLD, "Debênture": MID, "Nota comercial": "#4F63A0", "Cotas de Fiagro": "#9DBBA9", "Cotas de FIDC": "#8C4F3A", "Outros": "#6E7A72"}
        fig = px.scatter(sc, x="Taxa", y="mi", size="mi", text="rotulo", color="Instrumento", color_discrete_map=cores, size_max=22 if COMPACTO else 40,
                         custom_data=["Ativo", "Fundos", "Taxa", "mi"])
        fig.update_traces(textposition="top center", textfont=dict(size=9.5 if COMPACTO else 11, color=INK), marker=dict(opacity=0.8, line=dict(width=0.5, color="white")),
                          hovertemplate="<b>%{customdata[0]}</b><br>%{customdata[1]}<br>Taxa: %{customdata[2]:.2f}%<br>R$ %{customdata[3]:,.1f} mi<extra></extra>")
        fig.add_vline(x=float(sc["Taxa"].median()), line=dict(color=MUTED, width=1, dash="dash"))
        fig.update_xaxes(title=f"Taxa ({ir}, % a.a.)", ticksuffix="%", showgrid=True, gridcolor="#E8EDEA")
        fig.update_yaxes(title="Volume (R$ milhões)")
        mostrar(estilo(fig, 460))
        fonte()

    if tem_venc and df["Vencimento"].notna().any():
        legenda_fig(4, "Vencimentos por ano", "R$ milhões")
        vv = df[df["Vencimento"].notna()].assign(Ano=lambda x: x["Vencimento"].dt.year).groupby(["Ano", "Fundo"], as_index=False)["Valor"].sum()
        vv["mi"] = vv["Valor"] / 1e6
        fig = px.bar(vv, x="Ano", y="mi", color="Fundo", color_discrete_sequence=PALETA)
        fig.update_xaxes(title=None, dtick=1)
        fig.update_yaxes(title=None)
        mostrar(estilo(fig, 320))
        fonte()

# ---------------------------------------------------------------------------
# Taxas
# ---------------------------------------------------------------------------
with t2:
    idx_ok = [i for i in ["CDI+", "IPCA+", "%CDI", "Pré"] if (df["Indexador"] == i).any() and df.loc[df["Indexador"] == i, "Taxa"].notna().any()]
    if not idx_ok:
        st.info("Nenhum ativo com indexador e taxa preenchidos.")
    else:
        idx = st.radio("Indexador", idx_ok, horizontal=True, label_visibility="collapsed")
        d = df[(df["Indexador"] == idx) & df["Taxa"].notna()]
        st.markdown(f'<p class="note">{len(d)} ativos em {idx}, somando {brl(d["Valor"].sum())}. Média ponderada pelo valor em carteira.</p>', unsafe_allow_html=True)

        c1, c2 = st.columns([1.15, 1], gap="large")
        with c1:
            legenda_fig(5, f"Taxas por fundo, {idx}", "Cada ponto é um ativo")
            fig = px.strip(d, x="Taxa", y="Fundo", color="Fundo", color_discrete_sequence=PALETA, custom_data=["Ativo", "Valor"])
            fig.update_traces(marker=dict(size=10, opacity=0.75, line=dict(width=0.5, color="white")), jitter=0.45,
                              hovertemplate="%{customdata[0]}<br>%{x:.2f}%<br>R$ %{customdata[1]:,.0f}<extra></extra>")
            for _, r in dispersao(d, "Fundo").iterrows():
                fig.add_trace(go.Scatter(x=[r["Média ponderada"]], y=[r["Fundo"]], mode="markers", showlegend=False,
                                         marker=dict(symbol="line-ns", size=26, line=dict(width=3, color=INK)),
                                         hovertemplate=f"Média ponderada: {num_br(r['Média ponderada'], 2)}%<extra></extra>"))
            fig.update_xaxes(title=f"Taxa ({idx}, % a.a.)", ticksuffix="%", showgrid=True, gridcolor="#E8EDEA")
            fig.update_yaxes(title=None, showgrid=False)
            mostrar(estilo(fig, max(260, 90 + 70 * d["Fundo"].nunique()), legenda=False))
            st.markdown('<div class="src">Traço preto: média ponderada do fundo. Fonte: carteiras informadas nos relatórios de gestão.</div>', unsafe_allow_html=True)
        with c2:
            legenda_fig(6, "Volume por faixa de taxa", f"R$ milhões por faixa de 0,5 ponto, {idx}")
            lo, hi = np.floor(d["Taxa"].min() * 2) / 2, np.ceil(d["Taxa"].max() * 2) / 2 + 0.5
            faixas = pd.cut(d["Taxa"], bins=np.arange(lo, hi + 0.01, 0.5), right=False)
            hf = d.groupby(faixas, observed=True)["Valor"].sum().reset_index()
            hf["centro"] = hf["Taxa"].map(lambda b: (b.left + b.right) / 2)
            hf["faixa"] = hf["Taxa"].map(lambda b: f"{num_br(b.left, 1)}% a {num_br(b.right, 1)}%")
            hf["mi"] = hf["Valor"] / 1e6
            fig = px.bar(hf, x="centro", y="mi", color_discrete_sequence=[MID], text=hf["mi"].map(lambda v: num_br(v, 0)), custom_data=["faixa"])
            fig.update_traces(width=0.42, textposition="outside", cliponaxis=False, hovertemplate="%{customdata[0]}<br>R$ %{y:,.1f} mi<extra></extra>")
            fig.update_xaxes(title=f"Taxa ({idx}, % a.a.)", tickmode="linear", dtick=1 if COMPACTO else 0.5, ticksuffix="%", tickformat=".0f" if COMPACTO else ".1f")
            fig.update_yaxes(title=None, showticklabels=False, showgrid=False)
            mostrar(estilo(fig, 320, legenda=False))
            fonte()

        st.markdown('<div class="h2">Dispersão das taxas, por indexador</div>', unsafe_allow_html=True)
        st.markdown('<p class="note">Cada indexador tem sua própria tabela: spread sobre o CDI, juro real sobre o IPCA, percentual do CDI e taxa prefixada '
                    'não são comparáveis entre si. Médias ponderadas pelo valor em carteira.</p>', unsafe_allow_html=True)
        grupos = ["Fundo", "Instrumento"] + (["Segmento"] if tem_seg else [])
        por = st.radio("Agrupar por", grupos, horizontal=True)
        UNID = {"CDI+": ("Spread sobre o CDI", "CDI + %"), "IPCA+": ("Juro real sobre o IPCA", "IPCA + %"),
                "%CDI": ("Percentual do CDI", "% do CDI"), "Pré": ("Taxa prefixada", "% a.a.")}
        for ix in idx_ok:
            di = df[(df["Indexador"] == ix) & df["Taxa"].notna()]
            if di.empty:
                continue
            nome, un = UNID.get(ix, (ix, "%"))
            st.markdown(f'<div class="figt" style="margin-top:16px">{ix}: {nome.lower()}</div>'
                        f'<div class="figs">{len(di)} ativos, {brl(di["Valor"].sum())} em carteira. Valores em {un} ao ano</div>', unsafe_allow_html=True)
            tb = dispersao(di, por)
            tb["Volume"] = tb["Volume"] / 1e6
            tb = tb.round({"Volume": 1, "Mínimo": 2, "Média ponderada": 2, "Mediana": 2, "Máximo": 2})
            if COMPACTO:
                tb = tb[[por, "Ativos", "Média ponderada", "Mediana"]]
            tabela(tb, column_config={
                "Volume": st.column_config.NumberColumn("Volume (R$ mi)", format="localized"),
                "Mínimo": st.column_config.NumberColumn(f"Mínimo ({un})", format="localized"),
                "Média ponderada": st.column_config.NumberColumn(f"Média ponderada ({un})", format="localized"),
                "Mediana": st.column_config.NumberColumn(f"Mediana ({un})", format="localized"),
                "Máximo": st.column_config.NumberColumn(f"Máximo ({un})", format="localized"),
            })
        fonte()

        if d["Duration"].notna().any() or d["Vencimento"].notna().any():
            legenda_fig(7, "Taxa por prazo", "Cada ponto é um ativo")
            dd = d.copy()
            dd["Prazo (anos)"] = dd["Duration"].where(dd["Duration"].notna(), (dd["Vencimento"] - pd.Timestamp.today()).dt.days / 365.25)
            fig = px.scatter(dd, x="Prazo (anos)", y="Taxa", size="Valor", color="Fundo", color_discrete_sequence=PALETA, hover_name="Ativo", size_max=34)
            fig.update_yaxes(title=f"Taxa ({idx}, % a.a.)", ticksuffix="%")
            mostrar(estilo(fig, 380))
            fonte()

# ---------------------------------------------------------------------------
# Concentração
# ---------------------------------------------------------------------------
with t3:
    c1, c2 = st.columns([1.2, 1], gap="large")
    with c1:
        legenda_fig(8, "Maiores devedores", "Soma de todas as séries e fundos filtrados, R$ milhões")
        td = df.groupby("Devedor").agg(Valor=("Valor", "sum"), Fundos=("Fundo", "nunique")).sort_values("Valor", ascending=False).head(10 if COMPACTO else 15).reset_index()
        td["mi"] = td["Valor"] / 1e6
        td["rot"] = td.apply(lambda r: f"{num_br(r['Valor'] / total * 100, 1)}%" if COMPACTO else f"{num_br(r['mi'], 1)}  ({num_br(r['Valor'] / total * 100, 1)}%)", axis=1)
        td["cor"] = np.where(td["Fundos"] > 1, "Em mais de um fundo", "Em um fundo")
        fig = px.bar(td.iloc[::-1], x="mi", y="Devedor", orientation="h", text="rot", color="cor",
                     color_discrete_map={"Em um fundo": MID, "Em mais de um fundo": GREEN})
        fig.update_traces(textposition="outside", cliponaxis=False, hovertemplate="%{y}<br>R$ %{x:,.1f} mi<extra></extra>")
        fig.update_xaxes(title=None, showticklabels=False, range=[0, td["mi"].max() * 1.35])
        fig.update_yaxes(title=None, showgrid=False, categoryorder="total ascending")
        mostrar(estilo(fig, 520))
        fonte()
    with c2:
        legenda_fig(9, "Curva de concentração", "Volume acumulado pelos devedores, do maior para o menor")
        cum = (por_dev.cumsum() / total * 100).reset_index(drop=True)
        cc = pd.DataFrame({"Devedores": np.arange(1, len(cum) + 1), "Acumulado": cum.values})
        fig = px.area(cc, x="Devedores", y="Acumulado", color_discrete_sequence=[GREEN])
        fig.update_traces(line=dict(width=2), fillcolor="rgba(18,53,36,0.12)", hovertemplate="%{x} devedores: %{y:.1f}%<extra></extra>")
        for k in [5, 10]:
            if k <= len(cc):
                v = cc.loc[k - 1, "Acumulado"]
                fig.add_annotation(x=k, y=v, text=f"{k} maiores: {num_br(v, 0)}%", showarrow=True, arrowhead=0, ax=40, ay=30, font=dict(size=12, color=INK))
        fig.update_yaxes(title=None, ticksuffix="%", range=[0, 102])
        fig.update_xaxes(title="Número de devedores")
        mostrar(estilo(fig, 520, legenda=False))
        fonte()

    st.markdown('<div class="h2">Mesmo devedor em fundos diferentes</div>', unsafe_allow_html=True)
    mult = df.groupby("Devedor").filter(lambda x: x["Fundo"].nunique() > 1)
    if mult.empty:
        st.markdown('<p class="note">Nenhum devedor aparece em mais de um dos fundos filtrados.</p>', unsafe_allow_html=True)
    else:
        linhas = []
        for dev, x in mult.groupby("Devedor"):
            xc = x[x["Indexador"] == "CDI+"]
            linhas.append({"Devedor": dev, "Fundos": ", ".join(sorted(x["Fundo"].unique())), "Valor total": x["Valor"].sum(),
                           "Taxa mínima (CDI+)": xc["Taxa"].min() if len(xc) else np.nan, "Taxa máxima (CDI+)": xc["Taxa"].max() if len(xc) else np.nan,
                           "Diferença (bps)": (xc["Taxa"].max() - xc["Taxa"].min()) * 100 if len(xc) else np.nan})
        tm_ = pd.DataFrame(linhas).sort_values("Valor total", ascending=False)
        tm_["Nº de fundos"] = tm_["Fundos"].str.count(",") + 1
        tm_ = tm_.astype({"Taxa mínima (CDI+)": float, "Taxa máxima (CDI+)": float, "Diferença (bps)": float}).round({"Valor total": 0, "Taxa mínima (CDI+)": 2, "Taxa máxima (CDI+)": 2, "Diferença (bps)": 0})
        if COMPACTO:
            tm_ = tm_[["Devedor", "Nº de fundos", "Diferença (bps)"]]
        else:
            tm_ = tm_.drop(columns=["Nº de fundos"])
        tabela(tm_, column_config={
            "Valor total": st.column_config.NumberColumn("Valor total (R$)", format="localized"),
            "Taxa mínima (CDI+)": st.column_config.NumberColumn("Taxa mínima (CDI + %)", format="localized"),
            "Taxa máxima (CDI+)": st.column_config.NumberColumn("Taxa máxima (CDI + %)", format="localized"),
            "Diferença (bps)": st.column_config.NumberColumn(format="localized"),
        })
        st.markdown('<p class="note">Diferenças de taxa para o mesmo devedor costumam refletir séries, datas de aquisição ou garantias diferentes. Vale olhar caso a caso.</p>', unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Comparar fundos
# ---------------------------------------------------------------------------
with t4:
    linhas = []
    for f, x in df.groupby("Fundo"):
        dv = x.groupby("Devedor")["Valor"].sum().sort_values(ascending=False)
        tot = x["Valor"].sum()
        xc = x[x["Indexador"] == "CDI+"]
        linhas.append({"Fundo": f, "Volume": tot, "Ativos": len(x), "Devedores": len(dv), "Taxa média CDI+": media_pond(xc),
                       "Mediana CDI+": xc["Taxa"].median() if len(xc) else None, "Maior devedor": dv.index[0],
                       "Maior devedor (%)": dv.iloc[0] / tot * 100, "5 maiores (%)": dv.head(5).sum() / tot * 100, "HHI": hhi(dv),
                       "% em CRA": x.loc[x["Instrumento"] == "CRA", "Valor"].sum() / tot * 100,
                       "% em CPR-F": x.loc[x["Instrumento"] == "CPR-F", "Valor"].sum() / tot * 100})
    cf = pd.DataFrame(linhas).sort_values("Volume", ascending=False)
    cf = cf.astype({"Taxa média CDI+": float, "Mediana CDI+": float}).round({"Volume": 0, "Taxa média CDI+": 2, "Mediana CDI+": 2, "Maior devedor (%)": 1, "5 maiores (%)": 1, "HHI": 0, "% em CRA": 0, "% em CPR-F": 0})
    cf_v = cf.copy()
    if COMPACTO:
        cf_v["Volume"] = (cf_v["Volume"] / 1e6).round(1)
        cf_v = cf_v[["Fundo", "Volume", "Taxa média CDI+", "5 maiores (%)"]]
    tabela(cf_v, column_config={
        "Volume": st.column_config.NumberColumn("Volume (R$ mi)" if COMPACTO else "Volume (R$)", format="localized"),
        "Taxa média CDI+": st.column_config.NumberColumn("Taxa média (CDI + %)", format="localized"),
        "Mediana CDI+": st.column_config.NumberColumn("Mediana (CDI + %)", format="localized"),
        "Maior devedor (%)": st.column_config.NumberColumn(format="%.1f%%"),
        "5 maiores (%)": st.column_config.NumberColumn("5 maiores devedores", format="%.1f%%"),
        "HHI": st.column_config.NumberColumn("HHI", format="localized", help="Índice Herfindahl-Hirschman por devedor: acima de 1.500 indica concentração moderada; acima de 2.500, alta."),
        "% em CRA": st.column_config.NumberColumn(format="%.0f%%"), "% em CPR-F": st.column_config.NumberColumn(format="%.0f%%"),
    })
    if len(cf) > 1 and cf["Taxa média CDI+"].notna().any():
        legenda_fig(10, "Taxa e concentração", "Taxa média em CDI+ contra a participação dos 5 maiores devedores; tamanho pelo volume")
        fig = px.scatter(cf, x="5 maiores (%)", y="Taxa média CDI+", size="Volume", text="Fundo", color="Fundo", color_discrete_sequence=PALETA, size_max=30 if COMPACTO else 60)
        fig.update_traces(textposition="top center", hovertemplate="%{text}<br>5 maiores: %{x:.1f}%<br>CDI + %{y:.2f}%<extra></extra>")
        fig.update_xaxes(title="Participação dos 5 maiores devedores", ticksuffix="%", showgrid=True, gridcolor="#E8EDEA")
        fig.update_yaxes(title="Taxa média (CDI + % a.a.)", ticksuffix="%")
        mostrar(estilo(fig, 420, legenda=False))
        fonte()

# ---------------------------------------------------------------------------
# Carteira
# ---------------------------------------------------------------------------
with t5:
    cart = df.copy()
    cart["% do fundo"] = (cart["Valor"] / cart.groupby("Fundo")["Valor"].transform("sum") * 100).round(2)
    cart["Valor"] = cart["Valor"].round(2)
    cart["Taxa"] = [taxa_txt(t, i) for t, i in zip(cart["Taxa"], cart["Indexador"])]
    cols = ["Fundo", "Ativo", "Devedor", "Codigo", "Instrumento", "Valor", "% do fundo", "Indexador", "Taxa"]
    cols += [c for c in ["Segmento", "Vencimento", "Duration", "Rating"] if (cart[c].astype(str).str.strip().replace({"NaT": "", "nan": ""}) != "").any()]
    cart = cart.sort_values(["Fundo", "Valor"], ascending=[True, False])
    if COMPACTO:
        cart["Valor"] = (cart["Valor"] / 1e6).round(2)
        cols = ["Fundo", "Ativo", "Valor", "Taxa"]
    tabela(cart[cols], height=460 if COMPACTO else 560, column_config={
        "Codigo": st.column_config.TextColumn("Código"),
        "Valor": st.column_config.NumberColumn("Valor (R$ mi)" if COMPACTO else "Valor (R$)", format="localized"),
        "% do fundo": st.column_config.NumberColumn("% do fundo", format="%.2f%%"),
        "Vencimento": st.column_config.DateColumn(format="DD/MM/YYYY"),
    })
    exp = df.rename(columns={"Codigo": "codigo_ativo", "Taxa": "taxa_aa"}).copy()
    exp.columns = [norm(c) if c not in ("codigo_ativo", "taxa_aa") else c for c in exp.columns]
    c1, c2 = st.columns([1, 3])
    c1.download_button("Exportar CSV (pt-BR)", exp.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
                       file_name="carteiras_filtradas.csv", mime="text/csv", type="primary")
    c2.markdown('<p class="note">Exporta o que está filtrado, com separador ponto e vírgula e vírgula decimal, pronto para o Excel em português.</p>', unsafe_allow_html=True)
