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

# Identidade "revista": preto de tinta, um azul de destaque e cinzas suaves sobre branco
GREEN, MID, GOLD, INK, MUTED, LINE, RED = "#16181B", "#8C939C", "#2F4F9A", "#16181B", "#6A6F76", "#E6E8EB", "#A32020"
PALETA = ["#16181B", "#2F4F9A", "#9AA0A8", "#C08A2E", "#6F8FD0", "#C9CDD2", "#5A6068", "#1F3570", "#B5B9BF"]
# Faixa do topo em cor sólida. Azul escuro: "#0F2A4A". Laranja escuro (estilo Reuters): "#C2410C".
COR_CAPA = "#0F2A4A"

# Cores fixas: a mesma coisa tem sempre a mesma cor, em todos os gráficos
CORES_INDEXADOR = {"CDI+": "#1F4E8C", "IPCA+": "#C2410C", "%CDI": "#6B8CC2", "Pré": "#7A7F87", "USD": "#8C6D3F", "Sem indexador": "#C9CDD2"}
CORES_INSTRUMENTO = {"CRA": "#16181B", "CPR-F": "#1F4E8C", "CRI": "#6B8CC2", "Debênture": "#8C939C", "Nota comercial": "#C2410C",
                     "CDCA": "#5A6068", "Cotas de Fiagro": "#B5B9BF", "Cotas de FIDC": "#D9A47A", "Outros": "#DADDE1"}
PALETA_FUNDOS = ["#1F4E8C", "#C2410C", "#16181B", "#6B8CC2", "#8C939C", "#D9A47A", "#3F6E5A", "#7D5BA6", "#B5B9BF", "#A63D40"]

st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Schibsted+Grotesk:wght@400;500;600;700;800&family=Newsreader:ital,opsz,wght@0,6..72,400;0,6..72,500;1,6..72,400&display=swap');
html, body, .stApp, button, input, select, textarea {{ font-family: 'Schibsted Grotesk', 'Segoe UI', system-ui, sans-serif; }}
.stApp {{ background: #FFFFFF; }}
.block-container {{ padding-top: 2.4rem; max-width: 1240px; }}
.capa {{ background: {COR_CAPA}; color: #FFFFFF; padding: 52px 56px 44px; margin: 0 0 36px; }}
.kick {{ font-size: 14px; font-weight: 600; color: rgba(255,255,255,0.72); display: flex; gap: 10px; }}
.kick b {{ color: #FFFFFF; font-weight: 600; }}
.tg-t {{ font-size: 54px; font-weight: 800; color: #FFFFFF; line-height: 1.04; letter-spacing: -0.03em; margin: 14px 0 16px; max-width: 900px; }}
.tg-d {{ font-family: 'Newsreader', Georgia, serif; font-size: 22px; line-height: 1.5; color: rgba(255,255,255,0.88); max-width: 760px; margin: 0 0 18px; }}
.tg-m {{ font-size: 13px; color: rgba(255,255,255,0.68); }}
div[data-testid="stDownloadButton"] button {{ font-size: 12.5px; padding: 2px 12px; min-height: 30px; }}
.sc-id {{ font-size: 13px; color: {MUTED}; letter-spacing: 0.02em; }}
.sc-d {{ font-size: 21px; font-weight: 700; letter-spacing: -0.01em; color: {INK}; margin: 2px 0 2px; }}
.sc-alias {{ font-family: 'Newsreader', Georgia, serif; font-size: 15px; color: #4A5057; }}
.sc-meta {{ display: flex; flex-wrap: wrap; gap: 6px; margin: 10px 0 4px; }}
.bdg {{ display: inline-block; font-size: 12.5px; padding: 3px 10px; border: 1px solid {LINE}; border-radius: 999px; color: {INK}; background: #FFFFFF; }}
.bdg.hi {{ background: #FBEAEA; border-color: #F1C9C9; color: #8E1B1B; }}
.bdg.md {{ background: #FFF3D6; border-color: #F0DDA8; color: #7A4F00; }}
.bdg.lo {{ background: #E5F2EA; border-color: #C5E1CF; color: #1F5A3D; }}
.sumline {{ font-size: 14px; color: {MUTED}; margin: 6px 0 12px; }}
.kpis {{ display: grid; grid-template-columns: repeat(6, minmax(0, 1fr)); border-top: 1px solid {INK}; border-bottom: 1px solid {LINE}; margin: 4px 0 30px; }}
.kpi {{ padding: 16px 16px 16px 0; min-width: 0; }}
.kpi .l {{ font-size: 12.5px; color: {MUTED}; }}
.kpi .v {{ font-size: 24px; font-weight: 700; color: {INK}; font-variant-numeric: tabular-nums; line-height: 1.3; letter-spacing: -0.01em; white-space: nowrap; }}
.kpi .n {{ font-size: 12px; color: {MUTED}; }}
.kpi .v.alert {{ color: {RED}; }}
@media (max-width: 1000px) {{ .kpis {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }} }}
.lead {{ font-family: 'Newsreader', Georgia, serif; font-size: 20px; line-height: 1.7; color: #2C3036; max-width: 760px; margin: 0 0 28px; }}
.fig {{ font-size: 12px; font-weight: 600; color: {MUTED}; margin-top: 22px; }}
.figt {{ font-size: 19px; font-weight: 700; color: {INK}; letter-spacing: -0.01em; }}
.figs {{ font-family: 'Newsreader', Georgia, serif; font-size: 16px; color: #4A5057; margin-bottom: 4px; }}
.src {{ font-size: 11.5px; color: {MUTED}; margin: -4px 0 18px; }}
.note {{ font-family: 'Newsreader', Georgia, serif; font-size: 17px; color: #3A3E44; line-height: 1.6; max-width: 760px; }}
.h2 {{ font-size: 26px; font-weight: 700; color: {INK}; letter-spacing: -0.015em; margin: 40px 0 10px; padding-top: 28px; border-top: 1px solid {LINE}; }}
div[data-testid="stTabs"] button p {{ font-size: 15px; font-weight: 500; }}
div[data-testid="stDataFrame"] {{ border-top: 1px solid {INK}; }}
@media (max-width: 640px) {{
  .block-container {{ padding: 2.6rem 0.9rem 2rem; }}
  .capa {{ padding: 28px 22px 24px; margin-bottom: 24px; }}
  .kick {{ font-size: 12.5px; }}
  .tg-t {{ font-size: 30px; }}
  .tg-d {{ font-size: 17px; }}
  .kpi {{ padding: 10px 8px 10px 0; }}
  .kpi .v {{ font-size: 18px; white-space: normal; }}
  .kpi .l, .kpi .n {{ font-size: 11px; }}
  .lead {{ font-size: 16px; line-height: 1.6; }}
  .figt {{ font-size: 16px; }}
  .figs, .note {{ font-size: 14.5px; }}
  .h2 {{ font-size: 20px; margin-top: 26px; padding-top: 18px; }}
  div[data-testid="stTabs"] button p {{ font-size: 13.5px; }}
}}
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
    # textos e indicadores: escala automática (tabelas usam R$ mil)
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
    "gestora": ["gestora", "gestor"],
    "pct_pl": ["pl", "dopl", "pctpl", "percentualpl", "percentualdopl"],
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


COL_FUNDOS = {
    "fundo": ["fundo", "ticker", "codigo"], "nome": ["nome", "nomedofundo", "razaosocial"], "gestora": ["gestora", "gestor"],
    "base": ["database", "datadebase", "base", "mes", "referencia"], "pl": ["pl", "plrmilhoes", "plrsmilhoes", "patrimonioliquido", "patrimonioliquidors", "patrimonio"],
    "caixa": ["caixa", "caixadopl", "caixapl", "caixa%"],
}


@st.cache_data(show_spinner=False)
def ler_fundos(conteudo: bytes | None) -> pd.DataFrame:
    """Dados de cada fundo (PL, gestora, data-base, caixa), só da aba "Fundos" da própria planilha. Sem a aba, os campos ficam em branco."""
    bruto = None
    try:
        bruto = pd.read_excel(io.BytesIO(conteudo) if conteudo else ARQUIVO, sheet_name="Fundos")
    except Exception:
        bruto = None
    if bruto is None or bruto.empty:
        return pd.DataFrame()
    mp = {}
    for campo, chaves in COL_FUNDOS.items():
        for c in bruto.columns:
            if norm(c) in chaves and c not in mp.values():
                mp[campo] = c
                break
    if "fundo" not in mp or "pl" not in mp:
        return pd.DataFrame()
    f = pd.DataFrame({"Fundo": bruto[mp["fundo"]].astype(str).str.strip()})
    f["Nome"] = bruto[mp["nome"]].astype(str).replace({"nan": ""}) if "nome" in mp else ""
    f["Gestora"] = bruto[mp["gestora"]].astype(str).replace({"nan": ""}) if "gestora" in mp else ""
    f["Base"] = bruto[mp["base"]].astype(str).replace({"nan": ""}) if "base" in mp else ""
    pl = bruto[mp["pl"]].map(para_numero)
    # PL informado em R$ milhões (ex.: 2.164,5) ou em reais: abaixo de 1 milhão, considera milhões
    f["PL"] = np.where(pl < 1e6, pl * 1e6, pl)
    cx = bruto[mp["caixa"]].map(para_numero) if "caixa" in mp else pd.Series(np.nan, index=bruto.index)
    if cx.notna().any() and cx.dropna().abs().median() < 1:
        cx = cx * 100
    f["Caixa"] = cx
    return f[f["Fundo"].str.len() > 0].drop_duplicates("Fundo").reset_index(drop=True)


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
    df["Gestora"] = bruto[mp["gestora"]].astype(str).str.strip().replace({"nan": ""}) if "gestora" in mp else ""
    pp = bruto[mp["pct_pl"]].map(para_numero) if "pct_pl" in mp else pd.Series(np.nan, index=bruto.index)
    if pp.notna().any() and pp.dropna().abs().median() < 0.5:
        pp = pp * 100
    df["PctPL"] = pp
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
    cod = r["Codigo"].astype(str).str.strip().replace({"": "—", "nan": "—", "None": "—"})
    v = pd.DataFrame({"#": r["#"], "Operação": r["Ativo"], "Código": cod, "Instrumento": r["Instrumento"],
                      "Taxa": [taxa_txt(t, i) for t, i in zip(r["Taxa"], r["Indexador"])],
                      "Volume": (r["Volume"] / 1e3).round(0), "% do total": (r["Volume"] / total * 100).round(2), "Fundos": r["Fundos"]})
    if curto:
        v = v.drop(columns=["Instrumento", "% do total", "Fundos"])
        v["Operação"] = v["Operação"].str.replace(r"^(CRA|CRI|CPR-F|CPR|Debênture|FIDC|FIAGRO|CDCA|NC)\s+", "", regex=True)
    if com_fundo:
        # rankings de taxa: o fundo que detém a operação no lugar da numeração
        v = v.drop(columns=[c for c in ["#", "Fundos"] if c in v.columns])
        v.insert(2, "Fundo", r["Fundos"].values)
    return v


def RANK_CFG(total, vmax, idx=None):
    return {
        "#": st.column_config.NumberColumn("#", width=32 if COMPACTO else 40),
        "Operação": st.column_config.TextColumn("Operação", width=135 if COMPACTO else "medium"),
        "Taxa": st.column_config.TextColumn("Taxa", width=95 if COMPACTO else None),
        "Volume": st.column_config.NumberColumn("R$ mil" if COMPACTO else "Volume (R$ mil)", format="localized", width=75 if COMPACTO else None),
        "% do total": st.column_config.NumberColumn("% do total", format="%.2f%%"),
        "Fundos": st.column_config.TextColumn("Fundos", width="medium"),
        "Fundo": st.column_config.TextColumn("Fundo", width=80 if COMPACTO else "small"),
        "Código": st.column_config.TextColumn("Código", width=105 if COMPACTO else "small"),
    }


def hhi(valores: pd.Series) -> float:
    s = valores / valores.sum()
    return float((s ** 2).sum() * 10000)


def estilo(fig, altura=380, legenda=True):
    if COMPACTO:
        altura = max(260, int(altura * 0.75))
    fig.update_layout(
        template="simple_white", height=altura, separators=",.", margin=dict(l=4, r=8, t=10, b=4) if COMPACTO else dict(l=8, r=16, t=10, b=8),
        font=dict(family="Schibsted Grotesk, Segoe UI, sans-serif", size=11 if COMPACTO else 13, color=INK), dragmode=False,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=legenda,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title=None),
        hoverlabel=dict(bgcolor="white", font_size=13),
    )
    fig.update_xaxes(showgrid=False, linecolor=LINE, tickfont=dict(color=MUTED), fixedrange=True)
    fig.update_yaxes(showgrid=True, gridcolor="#EEF0F2", linecolor=LINE, tickfont=dict(color=MUTED), fixedrange=True)
    return fig


def mostrar(fig, dados: pd.DataFrame | None = None, nome: str = "grafico"):
    """Mostra o gráfico com o botão de baixar imagem (canto superior direito) e, se houver, os dados em Excel."""
    cfg = {"displaylogo": False, "modeBarButtons": [["toImage"]],
           "toImageButtonOptions": {"format": "png", "filename": f"teraag_{nome}", "scale": 2}}
    try:
        st.plotly_chart(fig, width="stretch", config=cfg, key=f"g_{nome}")
    except TypeError:
        st.plotly_chart(fig, use_container_width=True, config=cfg, key=f"g_{nome}")
    if dados is not None and len(dados):
        buf = io.BytesIO()
        dados.to_excel(buf, index=False, sheet_name="dados")
        st.download_button("Baixar dados (Excel)", buf.getvalue(), file_name=f"teraag_{nome}.xlsx", key=f"dl_{nome}",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


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
    fundos_info = ler_fundos(conteudo)
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
            "**Fundos:** para preencher patrimônio líquido, gestora e data-base, crie no **carteira_ativos.xlsx** uma aba chamada **Fundos** "
            "com as colunas **Fundo, Gestora, Data-base** e **PL** (em reais ou em R$ milhões). Enquanto a aba não existir, esses campos ficam em branco.\n\n"
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
_todos_fundos = sorted(df_all["Fundo"].unique())
CORES_FUNDO = {f: PALETA_FUNDOS[i % len(PALETA_FUNDOS)] for i, f in enumerate(_todos_fundos)}
GESTORA = dict(zip(fundos_info["Fundo"], fundos_info["Gestora"])) if len(fundos_info) else {}
for _f, _g in df_all[df_all["Gestora"] != ""].groupby("Fundo")["Gestora"].first().items():
    GESTORA.setdefault(_f, _g)
BASE = dict(zip(fundos_info["Fundo"], fundos_info["Base"])) if len(fundos_info) else {}
PL_FUNDO = dict(zip(fundos_info["Fundo"], fundos_info["PL"])) if len(fundos_info) else {}

# ---------------------------------------------------------------------------
# Cabeçalho, indicadores e resumo
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="capa"><div class="kick"><b>Crédito agro</b><span>|</span><span>TeraAg</span></div>'
    '<div class="tg-t">Monitor de Crédito - Taxas e Volume Financeiro</div>'
    '<p class="tg-d">Taxas, volumes e as maiores operações das carteiras de Fiagros e FIDCs, em um só lugar.</p>'
    f'<div class="tg-m">{df_all["Fundo"].nunique()} fundos e {len(df_all)} ativos na base</div></div>', unsafe_allow_html=True)

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

t1, t2, t3, t4, tpq, t5 = st.tabs(["Visão geral", "Taxas", "Concentração", "Fundos", "Pesquisa", "Carteira"])

# ---------------------------------------------------------------------------
# Visão geral
# ---------------------------------------------------------------------------
with t1:
    c1, c2 = st.columns([1, 1], gap="large")
    with c1:
        vf = df.groupby("Fundo", as_index=False)["Valor"].sum().sort_values("Valor")
        lider = vf.iloc[-1]
        legenda_fig(1, f"{lider['Fundo']} reúne {num_br(lider['Valor'] / total * 100, 0)}% do volume" if len(vf) > 1 else f"{lider['Fundo']}: {brl(lider['Valor'])}",
                    "Volume por fundo, R$ milhões")
        vf["mi"] = vf["Valor"] / 1e6
        vf["rot"] = vf.apply(lambda r: f"{num_br(r['mi'], 1)}  ({num_br(r['Valor'] / total * 100, 0)}%)", axis=1)
        fig = px.bar(vf, x="mi", y="Fundo", orientation="h", text="rot", color="Fundo", color_discrete_map=CORES_FUNDO)
        fig.update_traces(textposition="outside", cliponaxis=False, hovertemplate="%{y}<br>R$ %{x:,.1f} mi<extra></extra>")
        fig.update_xaxes(title=None, showticklabels=False, range=[0, vf["mi"].max() * 1.6])
        fig.update_yaxes(title=None, showgrid=False, categoryorder="total ascending")
        mostrar(estilo(fig, max(220, 60 + 42 * len(vf)), legenda=False), vf[["Fundo", "Valor"]].sort_values("Valor", ascending=False), "volume_por_fundo")
        fonte()
    with c2:
        ci = df.groupby(["Fundo", "Instrumento"], as_index=False)["Valor"].sum()
        inst_tot = ci.groupby("Instrumento")["Valor"].sum().sort_values(ascending=False)
        legenda_fig(2, f"{inst_tot.index[0]} responde por {num_br(inst_tot.iloc[0] / total * 100, 0)}% do volume",
                    "Composição por instrumento, participação no volume de cada fundo")
        ci["pct"] = ci["Valor"] / ci.groupby("Fundo")["Valor"].transform("sum") * 100
        fig = px.bar(ci, x="pct", y="Fundo", color="Instrumento", orientation="h", color_discrete_map=CORES_INSTRUMENTO,
                     category_orders={"Instrumento": list(inst_tot.index)},
                     custom_data=["Instrumento", "Valor"])
        fig.update_traces(hovertemplate="%{customdata[0]}: %{x:.1f}%<extra></extra>")
        fig.update_xaxes(title=None, ticksuffix="%", range=[0, 100])
        fig.update_yaxes(title=None, showgrid=False)
        fig.update_layout(barmode="stack")
        mostrar(estilo(fig, max(220, 90 + 42 * df["Fundo"].nunique())), ci.round({"pct": 2}).rename(columns={"pct": "% do fundo"}), "composicao_instrumento")
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
    tabela(rank_view(rv, total, curto=COMPACTO, com_fundo=True), column_config=RANK_CFG(total, rv["Volume"].max()), height=38 + 35 * len(rv))
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

        sc = base.copy()
        sc["mi"] = sc["Volume"] / 1e6
        k_top = 5 if COMPACTO else 10
        top = set(sc.nlargest(k_top, "Volume")["Chave"])
        acima = int((sc[sc["Chave"].isin(top)]["Taxa"] > sc["Taxa"].median()).sum())
        legenda_fig(3, f"{acima} das {min(k_top, len(sc))} maiores operações pagam acima da mediana em {ir}",
                    f"Volume e taxa das operações; as {k_top} maiores por volume estão identificadas. Linha tracejada: mediana da taxa")
        sc["rotulo"] = np.where(sc["Chave"].isin(top), sc["Ativo"].str.replace(r"^(CRA|CRI|CPR-F|CPR|Debênture|FIDC|FIAGRO|CDCA|NC)\s+", "", regex=True), "")
        fig = px.scatter(sc, x="Taxa", y="mi", size="mi", text="rotulo", color="Instrumento", color_discrete_map=CORES_INSTRUMENTO, size_max=22 if COMPACTO else 40,
                         custom_data=["Ativo", "Fundos", "Taxa", "mi"])
        fig.update_traces(textposition="top center", textfont=dict(size=9.5 if COMPACTO else 11, color=INK), marker=dict(opacity=0.8, line=dict(width=0.5, color="white")),
                          hovertemplate="<b>%{customdata[0]}</b><br>%{customdata[1]}<br>Taxa: %{customdata[2]:.2f}%<br>R$ %{customdata[3]:,.1f} mi<extra></extra>")
        fig.add_vline(x=float(sc["Taxa"].median()), line=dict(color=MUTED, width=1, dash="dash"))
        fig.update_xaxes(title=f"Taxa ({ir}, % a.a.)", ticksuffix="%", showgrid=True, gridcolor="#EEF0F2")
        fig.update_yaxes(title="Volume (R$ milhões)")
        mostrar(estilo(fig, 460), sc[["Ativo", "Codigo", "Instrumento", "Fundos", "Taxa", "Volume"]].sort_values("Volume", ascending=False), "volume_taxa_operacoes")
        fonte()

    if tem_venc and df["Vencimento"].notna().any():
        vv = df[df["Vencimento"].notna()].assign(Ano=lambda x: x["Vencimento"].dt.year).groupby(["Ano", "Fundo"], as_index=False)["Valor"].sum()
        ano_top = vv.groupby("Ano")["Valor"].sum().idxmax()
        legenda_fig(4, f"{ano_top} concentra os maiores vencimentos", "Vencimentos por ano, R$ milhões")
        vv["mi"] = vv["Valor"] / 1e6
        fig = px.bar(vv, x="Ano", y="mi", color="Fundo", color_discrete_map=CORES_FUNDO)
        fig.update_xaxes(title=None, dtick=1)
        fig.update_yaxes(title=None)
        mostrar(estilo(fig, 320), vv[["Ano", "Fundo", "Valor"]], "vencimentos_por_ano")
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
            dsp = dispersao(d, "Fundo")
            f_top = dsp.sort_values("Média ponderada", ascending=False).iloc[0]
            legenda_fig(5, f"{f_top['Fundo']} tem a maior taxa média: {taxa_txt(f_top['Média ponderada'], idx)}", f"Taxas por fundo em {idx}; cada ponto é um ativo")
            fig = px.strip(d, x="Taxa", y="Fundo", color="Fundo", color_discrete_map=CORES_FUNDO, custom_data=["Ativo", "Valor"])
            fig.update_traces(marker=dict(size=10, opacity=0.75, line=dict(width=0.5, color="white")), jitter=0.45,
                              hovertemplate="%{customdata[0]}<br>%{x:.2f}%<br>R$ %{customdata[1]:,.0f}<extra></extra>")
            for _, r in dsp.iterrows():
                fig.add_trace(go.Scatter(x=[r["Média ponderada"]], y=[r["Fundo"]], mode="markers", showlegend=False,
                                         marker=dict(symbol="line-ns", size=26, line=dict(width=3, color=INK)),
                                         hovertemplate=f"Média ponderada: {num_br(r['Média ponderada'], 2)}%<extra></extra>"))
            fig.update_xaxes(title=f"Taxa ({idx}, % a.a.)", ticksuffix="%", showgrid=True, gridcolor="#EEF0F2")
            fig.update_yaxes(title=None, showgrid=False)
            mostrar(estilo(fig, max(260, 90 + 70 * d["Fundo"].nunique()), legenda=False), d[["Fundo", "Ativo", "Codigo", "Taxa", "Valor"]], f"taxas_por_fundo_{norm(idx)}")
            st.markdown('<div class="src">Traço preto: média ponderada do fundo. Fonte: carteiras informadas nos relatórios de gestão.</div>', unsafe_allow_html=True)
        with c2:
            lo, hi = np.floor(d["Taxa"].min() * 2) / 2, np.ceil(d["Taxa"].max() * 2) / 2 + 0.5
            faixas = pd.cut(d["Taxa"], bins=np.arange(lo, hi + 0.01, 0.5), right=False)
            hf = d.groupby(faixas, observed=True)["Valor"].sum().reset_index()
            hf["centro"] = hf["Taxa"].map(lambda b: (b.left + b.right) / 2)
            hf["faixa"] = hf["Taxa"].map(lambda b: f"{num_br(b.left, 1)}% a {num_br(b.right, 1)}%")
            hf["mi"] = hf["Valor"] / 1e6
            legenda_fig(6, "O maior volume está entre " + hf.loc[hf["Valor"].idxmax(), "faixa"].replace("% a ", "% e "), f"Volume por faixa de taxa, R$ milhões por faixa de 0,5 ponto, {idx}")
            fig = px.bar(hf, x="centro", y="mi", color_discrete_sequence=[CORES_INDEXADOR.get(idx, MID)], text=hf["mi"].map(lambda v: num_br(v, 0)), custom_data=["faixa"])
            fig.update_traces(width=0.42, textposition="outside", cliponaxis=False, hovertemplate="%{customdata[0]}<br>R$ %{y:,.1f} mi<extra></extra>")
            fig.update_xaxes(title=f"Taxa ({idx}, % a.a.)", tickmode="linear", dtick=1 if COMPACTO else 0.5, ticksuffix="%", tickformat=".0f" if COMPACTO else ".1f")
            fig.update_yaxes(title=None, showticklabels=False, showgrid=False)
            mostrar(estilo(fig, 320, legenda=False), hf[["faixa", "Valor"]].rename(columns={"faixa": "Faixa de taxa"}), f"volume_por_faixa_{norm(idx)}")
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
            tb["Volume"] = tb["Volume"] / 1e3
            tb = tb.round({"Volume": 0, "Mínimo": 2, "Média ponderada": 2, "Mediana": 2, "Máximo": 2})
            if COMPACTO:
                tb = tb[[por, "Ativos", "Média ponderada", "Mediana"]]
            tabela(tb, column_config={
                "Volume": st.column_config.NumberColumn("Volume (R$ mil)", format="localized"),
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
            fig = px.scatter(dd, x="Prazo (anos)", y="Taxa", size="Valor", color="Fundo", color_discrete_map=CORES_FUNDO, hover_name="Ativo", size_max=34)
            fig.update_yaxes(title=f"Taxa ({idx}, % a.a.)", ticksuffix="%")
            mostrar(estilo(fig, 380), dd[["Fundo", "Ativo", "Prazo (anos)", "Taxa", "Valor"]], f"taxa_por_prazo_{norm(idx)}")
            fonte()

# ---------------------------------------------------------------------------
# Concentração
# ---------------------------------------------------------------------------
with t3:
    c1, c2 = st.columns([1.2, 1], gap="large")
    with c1:
        td = df.groupby("Devedor").agg(Valor=("Valor", "sum"), Fundos=("Fundo", "nunique")).sort_values("Valor", ascending=False).head(10 if COMPACTO else 15).reset_index()
        td["mi"] = td["Valor"] / 1e6
        legenda_fig(8, f"{td.loc[0, 'Devedor']} é o maior devedor, com {num_br(td.loc[0, 'Valor'] / total * 100, 1)}% do volume",
                    "Maiores devedores, soma de todas as séries e fundos filtrados, R$ milhões")
        td["rot"] = td.apply(lambda r: f"{num_br(r['Valor'] / total * 100, 1)}%" if COMPACTO else f"{num_br(r['mi'], 1)}  ({num_br(r['Valor'] / total * 100, 1)}%)", axis=1)
        td["cor"] = np.where(td["Fundos"] > 1, "Em mais de um fundo", "Em um fundo")
        fig = px.bar(td.iloc[::-1], x="mi", y="Devedor", orientation="h", text="rot", color="cor",
                     color_discrete_map={"Em um fundo": MID, "Em mais de um fundo": GREEN})
        fig.update_traces(textposition="outside", cliponaxis=False, hovertemplate="%{y}<br>R$ %{x:,.1f} mi<extra></extra>")
        fig.update_xaxes(title=None, showticklabels=False, range=[0, td["mi"].max() * 1.35])
        fig.update_yaxes(title=None, showgrid=False, categoryorder="total ascending")
        mostrar(estilo(fig, 520), td[["Devedor", "Valor", "Fundos"]].rename(columns={"Fundos": "Nº de fundos"}), "maiores_devedores")
        fonte()
    with c2:
        cum = (por_dev.cumsum() / total * 100).reset_index(drop=True)
        cc = pd.DataFrame({"Devedores": np.arange(1, len(cum) + 1), "Acumulado": cum.values})
        legenda_fig(9, f"Os 10 maiores devedores somam {num_br(top10, 0)}% do volume", "Curva de concentração: volume acumulado, do maior devedor para o menor")
        fig = px.area(cc, x="Devedores", y="Acumulado", color_discrete_sequence=[GREEN])
        fig.update_traces(line=dict(width=2), fillcolor="rgba(22,24,27,0.07)", hovertemplate="%{x} devedores: %{y:.1f}%<extra></extra>")
        for k in [5, 10]:
            if k <= len(cc):
                v = cc.loc[k - 1, "Acumulado"]
                fig.add_annotation(x=k, y=v, text=f"{k} maiores: {num_br(v, 0)}%", showarrow=True, arrowhead=0, ax=40, ay=30, font=dict(size=12, color=INK))
        fig.update_yaxes(title=None, ticksuffix="%", range=[0, 102])
        fig.update_xaxes(title="Número de devedores")
        mostrar(estilo(fig, 520, legenda=False), cc.round({"Acumulado": 2}).rename(columns={"Acumulado": "% acumulado do volume"}), "curva_concentracao")
        fonte()

    st.markdown('<div class="h2">Mesmo devedor em fundos diferentes</div>', unsafe_allow_html=True)
    mult = df.groupby("Devedor").filter(lambda x: x["Fundo"].nunique() > 1)
    if mult.empty:
        st.markdown('<p class="note">Nenhum devedor aparece em mais de um dos fundos filtrados.</p>', unsafe_allow_html=True)
    else:
        linhas = []
        for dev, x in mult.groupby("Devedor"):
            xc = x[x["Indexador"] == "CDI+"]
            cods = sorted(c for c in x["Codigo"].astype(str).unique() if c and c.lower() not in ("nan", "none", "-"))
            linhas.append({"Devedor": dev, "Códigos": ", ".join(cods) if cods else "—", "Fundos": ", ".join(sorted(x["Fundo"].unique())), "Valor total": x["Valor"].sum(),
                           "Taxa mínima (CDI+)": xc["Taxa"].min() if len(xc) else np.nan, "Taxa máxima (CDI+)": xc["Taxa"].max() if len(xc) else np.nan,
                           "Diferença (bps)": (xc["Taxa"].max() - xc["Taxa"].min()) * 100 if len(xc) else np.nan})
        tm_ = pd.DataFrame(linhas).sort_values("Valor total", ascending=False)
        tm_["Nº de fundos"] = tm_["Fundos"].str.count(",") + 1
        tm_["Valor total"] = tm_["Valor total"] / 1e3
        tm_ = tm_.astype({"Taxa mínima (CDI+)": float, "Taxa máxima (CDI+)": float, "Diferença (bps)": float}).round({"Valor total": 0, "Taxa mínima (CDI+)": 2, "Taxa máxima (CDI+)": 2, "Diferença (bps)": 0})
        if COMPACTO:
            tm_ = tm_[["Devedor", "Códigos", "Nº de fundos", "Diferença (bps)"]]
        else:
            tm_ = tm_.drop(columns=["Nº de fundos"])
        tabela(tm_, column_config={
            "Valor total": st.column_config.NumberColumn("Valor total (R$ mil)", format="localized"),
            "Taxa mínima (CDI+)": st.column_config.NumberColumn("Taxa mínima (CDI + %)", format="localized"),
            "Taxa máxima (CDI+)": st.column_config.NumberColumn("Taxa máxima (CDI + %)", format="localized"),
            "Diferença (bps)": st.column_config.NumberColumn(format="localized"),
        })
        st.markdown('<p class="note">Diferenças de taxa para o mesmo devedor costumam refletir séries, datas de aquisição ou garantias diferentes. Vale olhar caso a caso.</p>', unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Fundos (comparação, patrimônio líquido e gestoras)
# ---------------------------------------------------------------------------
with t4:
    tem_pl = bool(PL_FUNDO)
    tem_gestora = any(GESTORA.get(f) for f in df["Fundo"].unique())
    linhas = []
    for f, x in df.groupby("Fundo"):
        dv = x.groupby("Devedor")["Valor"].sum().sort_values(ascending=False)
        tot = x["Valor"].sum()
        xc = x[x["Indexador"] == "CDI+"]
        pl_f = PL_FUNDO.get(f)
        linhas.append({"Fundo": f, "Gestora": GESTORA.get(f, ""), "Data-base": BASE.get(f, "") or (x["Mes"].iloc[0] if x["Mes"].iloc[0] else ""),
                       "PL": pl_f if pl_f else np.nan, "Volume": tot, "% do PL em crédito": tot / pl_f * 100 if pl_f else np.nan,
                       "Ativos": len(x), "Devedores": len(dv), "Taxa média CDI+": media_pond(xc),
                       "Mediana CDI+": xc["Taxa"].median() if len(xc) else np.nan, "Maior devedor": dv.index[0],
                       "Maior devedor (%)": dv.iloc[0] / tot * 100, "5 maiores (%)": dv.head(5).sum() / tot * 100, "HHI": hhi(dv),
                       "% em CRA": x.loc[x["Instrumento"] == "CRA", "Valor"].sum() / tot * 100,
                       "% em CPR-F": x.loc[x["Instrumento"] == "CPR-F", "Valor"].sum() / tot * 100})
    cf = pd.DataFrame(linhas).astype({"PL": float, "% do PL em crédito": float, "Taxa média CDI+": float, "Mediana CDI+": float})
    cf = cf.sort_values("Volume", ascending=False).reset_index(drop=True)
    vol_tot = cf["Volume"].sum()
    cf["% do volume"] = cf["Volume"] / vol_tot * 100

    # indicadores: PL e gestoras só aparecem quando informados na aba Fundos
    pl_sum = cf["PL"].sum() if tem_pl else None
    kp = [("Volume em carteira", brl(vol_tot), f"{len(cf)} fundo(s)"),
          ("Maior carteira", cf.loc[0, "Fundo"], f"{brl(cf.loc[0, 'Volume'])}, {num_br(cf.loc[0, '% do volume'], 0)}% do volume"),
          ("5 maiores carteiras", f"{num_br(cf['Volume'].head(5).sum() / vol_tot * 100, 0)}%", "do volume filtrado"),
          ("PL somado", brl(pl_sum) if tem_pl else "—", "dos fundos com PL informado" if tem_pl else "a preencher na aba Fundos"),
          ("% do PL em crédito", f"{num_br(cf.loc[cf['PL'].notna(), 'Volume'].sum() / pl_sum * 100, 0)}%" if tem_pl and pl_sum else "—",
           "volume em carteira ÷ PL" if tem_pl else "depende do PL"),
          ("Gestoras", str(len({g for g in cf["Gestora"] if g})) if tem_gestora else "—", "na base filtrada" if tem_gestora else "a preencher na aba Fundos")]
    st.markdown('<div class="kpis">' + "".join(f'<div class="kpi"><div class="l">{l}</div><div class="v">{v}</div><div class="n">{n}</div></div>' for l, v, n in kp) + "</div>",
                unsafe_allow_html=True)
    if not tem_pl:
        st.markdown('<p class="note">Patrimônio líquido, gestora e data-base ficam em branco até serem preenchidos na planilha, '
                    'numa aba chamada <b>Fundos</b>. Veja como em "Como enriquecer a planilha", na barra lateral.</p>', unsafe_allow_html=True)

    cA, cB = st.columns([1, 2], vertical_alignment="bottom")
    opcoes = ["Por fundo", "Por gestora"] if tem_gestora else ["Por fundo"]
    visao = cA.radio("Agrupar", opcoes, horizontal=True, label_visibility="collapsed", key="f_visao")
    qf = norm(cB.text_input("Filtrar fundo ou gestora", placeholder="Filtrar fundo ou gestora", label_visibility="collapsed", key="f_q"))

    if visao == "Por fundo":
        tv = cf if not qf else cf[cf["Fundo"].map(norm).str.contains(qf) | cf["Gestora"].map(norm).str.contains(qf)]
        st.markdown(f'<div class="sumline">{len(tv)} fundo(s) · {brl(tv["Volume"].sum())} em carteira</div>', unsafe_allow_html=True)
        tv_v = tv.copy()
        tv_v["PL"] = (tv_v["PL"] / 1e3).round(0)
        tv_v["Volume"] = (tv_v["Volume"] / 1e3).round(0)
        tv_v = tv_v.round({"% do volume": 1, "% do PL em crédito": 1, "Taxa média CDI+": 2, "Mediana CDI+": 2, "Maior devedor (%)": 1, "5 maiores (%)": 1, "HHI": 0, "% em CRA": 0, "% em CPR-F": 0})
        tv_v["Gestora"] = tv_v["Gestora"].replace("", "—")
        tv_v["Data-base"] = tv_v["Data-base"].replace("", "—")
        if COMPACTO:
            cols = ["Fundo", "Volume", "Taxa média CDI+", "5 maiores (%)"]
        else:
            cols = ["Fundo", "Gestora", "Data-base", "PL", "Volume", "% do volume", "% do PL em crédito", "Ativos", "Devedores",
                    "Taxa média CDI+", "Mediana CDI+", "Maior devedor", "Maior devedor (%)", "5 maiores (%)", "HHI", "% em CRA", "% em CPR-F"]
        cfg = {"PL": st.column_config.NumberColumn("PL (R$ mil)", format="localized", help="Preenchido pela aba Fundos da planilha."),
               "Volume": st.column_config.NumberColumn("Volume em carteira (R$ mil)", format="localized"),
               "% do volume": st.column_config.NumberColumn(format="%.1f%%"),
               "% do PL em crédito": st.column_config.NumberColumn(format="%.1f%%", help="Volume em carteira dividido pelo PL."),
               "Taxa média CDI+": st.column_config.NumberColumn("Taxa média (CDI + %)", format="localized"),
               "Mediana CDI+": st.column_config.NumberColumn("Mediana (CDI + %)", format="localized"),
               "Maior devedor (%)": st.column_config.NumberColumn(format="%.1f%%"),
               "5 maiores (%)": st.column_config.NumberColumn("5 maiores devedores", format="%.1f%%"),
               "HHI": st.column_config.NumberColumn("HHI", format="localized", help="Índice Herfindahl-Hirschman por devedor: acima de 1.500 indica concentração moderada; acima de 2.500, alta."),
               "% em CRA": st.column_config.NumberColumn(format="%.0f%%"), "% em CPR-F": st.column_config.NumberColumn(format="%.0f%%")}
        sel = None
        try:
            ev = st.dataframe(tv_v[cols], hide_index=True, width="stretch", column_config=cfg, on_select="rerun",
                              selection_mode="single-row", key="f_tab", height=min(38 + 35 * len(tv_v), 560))
            linhas_sel = ev.selection.rows if ev and ev.selection else []
            sel = tv_v.iloc[linhas_sel[0]]["Fundo"] if linhas_sel else None
        except TypeError:
            tabela(tv_v[cols], column_config=cfg)
        st.markdown('<div class="src">Clique numa linha para ver a carteira do fundo. Clique no nome de uma coluna para ordenar.</div>', unsafe_allow_html=True)
        if sel:
            cx = df[df["Fundo"] == sel].sort_values("Valor", ascending=False).copy()
            sub = f'{len(cx)} ativos, {brl(cx["Valor"].sum())} em carteira' + (f', {num_br(cx["Valor"].sum() / PL_FUNDO[sel] * 100, 1)}% do PL' if PL_FUNDO.get(sel) else "")
            st.markdown(f'<div class="figt" style="margin-top:18px">Carteira do {sel}</div><div class="figs">{sub}</div>', unsafe_allow_html=True)
            cx["Taxa"] = [taxa_txt(t, i) for t, i in zip(cx["Taxa"], cx["Indexador"])]
            cx["% da carteira"] = (cx["Valor"] / cx["Valor"].sum() * 100).round(2)
            cx["Valor"] = (cx["Valor"] / 1e3).round(0)
            tabela(cx[["Ativo", "Codigo", "Instrumento", "Taxa", "Valor", "% da carteira"]], column_config={
                "Codigo": st.column_config.TextColumn("Código"), "Valor": st.column_config.NumberColumn("Valor (R$ mil)", format="localized"),
                "% da carteira": st.column_config.NumberColumn(format="%.2f%%")})
    else:
        gg = cf[cf["Gestora"] != ""].groupby("Gestora").agg(Fundos=("Fundo", "count"), Lista=("Fundo", lambda v: ", ".join(v)),
                                                             PL=("PL", "sum"), Volume=("Volume", "sum"), Ativos=("Ativos", "sum")).reset_index()
        gg["% do volume"] = gg["Volume"] / vol_tot * 100
        gg = gg.sort_values("Volume", ascending=False)
        if qf:
            gg = gg[gg["Gestora"].map(norm).str.contains(qf) | gg["Lista"].map(norm).str.contains(qf)]
        st.markdown(f'<div class="sumline">{len(gg)} gestora(s) · {brl(gg["Volume"].sum())} em carteira</div>', unsafe_allow_html=True)
        gg["PL"] = (gg["PL"].replace(0, np.nan) / 1e3).round(0)
        gg["Volume"] = (gg["Volume"] / 1e3).round(0)
        gg = gg.round({"% do volume": 1})
        cols = ["Gestora", "Fundos", "Volume"] if COMPACTO else ["Gestora", "Fundos", "Lista", "PL", "Volume", "% do volume", "Ativos"]
        tabela(gg[cols], column_config={"Lista": st.column_config.TextColumn("Fundos da gestora"),
                                        "PL": st.column_config.NumberColumn("PL (R$ mil)", format="localized"),
                                        "Volume": st.column_config.NumberColumn("Volume em carteira (R$ mil)", format="localized"),
                                        "% do volume": st.column_config.NumberColumn(format="%.1f%%")})

    if tem_pl and cf["PL"].notna().any():
        gp = cf[cf["PL"].notna()].sort_values("PL", ascending=False).head(10 if COMPACTO else 15).copy()
        gp["mi"] = gp["PL"] / 1e6
        legenda_fig(11, f"{gp.iloc[0]['Fundo']} tem o maior patrimônio líquido: {brl(gp.iloc[0]['PL'])}", "Patrimônio líquido por fundo, R$ milhões")
        fig = px.bar(gp, x="mi", y="Fundo", orientation="h", color="Fundo", color_discrete_map=CORES_FUNDO,
                     text=gp["mi"].map(lambda v: num_br(v, 0)))
        fig.update_traces(textposition="outside", cliponaxis=False, hovertemplate="%{y}<br>R$ %{x:,.0f} mi<extra></extra>")
        fig.update_xaxes(title=None, showticklabels=False, range=[0, gp["mi"].max() * 1.3])
        fig.update_yaxes(title=None, showgrid=False, categoryorder="total ascending")
        mostrar(estilo(fig, max(240, 60 + 34 * len(gp)), legenda=False), gp[["Fundo", "Gestora", "Data-base", "PL"]], "patrimonio_liquido")
        fonte("aba Fundos da planilha.")

    if len(cf) > 1 and cf["Taxa média CDI+"].notna().any():
        f_tx = cf.sort_values("Taxa média CDI+", ascending=False).iloc[0]
        f_cc = cf.sort_values("5 maiores (%)", ascending=False).iloc[0]
        legenda_fig(10, f"{f_tx['Fundo']} paga a maior taxa média e é o mais concentrado" if f_tx["Fundo"] == f_cc["Fundo"]
                    else f"{f_tx['Fundo']} paga a maior taxa média; {f_cc['Fundo']} é o mais concentrado",
                    "Taxa média em CDI+ contra a participação dos 5 maiores devedores; tamanho pelo volume em carteira")
        fig = px.scatter(cf, x="5 maiores (%)", y="Taxa média CDI+", size="Volume", text="Fundo", color="Fundo", color_discrete_map=CORES_FUNDO, size_max=30 if COMPACTO else 60)
        fig.update_traces(textposition="top center", hovertemplate="%{text}<br>5 maiores: %{x:.1f}%<br>CDI + %{y:.2f}%<extra></extra>")
        fig.update_xaxes(title="Participação dos 5 maiores devedores", ticksuffix="%", showgrid=True, gridcolor="#EEF0F2")
        fig.update_yaxes(title="Taxa média (CDI + % a.a.)", ticksuffix="%")
        mostrar(estilo(fig, 420, legenda=False), cf[["Fundo", "Volume", "Taxa média CDI+", "5 maiores (%)"]], "taxa_e_concentracao")
        fonte()

# ---------------------------------------------------------------------------
# Pesquisa de operações (mesmo modelo de "operação" dos rankings)
# ---------------------------------------------------------------------------
def _sugerir(v):
    st.session_state["busca_ativo"] = v


with tpq:
    st.markdown('<p class="note">Busque por código ou pelo nome do devedor. Cada resultado é uma operação, como nos rankings: '
                'o mesmo ativo reunido em todos os fundos filtrados que o detêm.</p>', unsafe_allow_html=True)
    q = st.text_input("Pesquisar", placeholder="Ex.: nome do devedor ou código do ativo", key="busca_ativo", label_visibility="collapsed")
    ops_all = operacoes(df)
    # exemplos tirados da própria base: operações presentes em mais fundos
    sug = ops_all[ops_all["Nº fundos"] > 1].sort_values(["Nº fundos", "Volume"], ascending=False)["Devedor"].drop_duplicates().head(4 if COMPACTO else 6).tolist()
    if sug:
        cs = st.columns(len(sug) + 1)
        cs[0].markdown('<div class="sumline" style="margin-top:8px">Em mais fundos:</div>', unsafe_allow_html=True)
        for c, v in zip(cs[1:], sug):
            try:
                c.button(v, key=f"sug_{norm(v)}", on_click=_sugerir, args=(v,), width="stretch")
            except TypeError:
                c.button(v, key=f"sug_{norm(v)}", on_click=_sugerir, args=(v,))

    qn = norm(q.strip())
    if len(qn) < 2:
        st.markdown('<p class="note">Digite pelo menos 2 caracteres.</p>', unsafe_allow_html=True)
    else:
        x = df.copy()
        x["Chave"] = np.where(x["Codigo"].astype(str).str.len() > 3, x["Codigo"].astype(str).str.upper(), x["Ativo"].map(norm))
        achou = x["Ativo"].map(norm).str.contains(qn, regex=False) | x["Devedor"].map(norm).str.contains(qn, regex=False) | x["Codigo"].map(norm).str.contains(qn, regex=False)
        chaves = set(x.loc[achou, "Chave"])
        res = ops_all[ops_all["Chave"].isin(chaves)].sort_values(["Nº fundos", "Volume"], ascending=False)
        if res.empty:
            st.markdown(f'<p class="note">Nenhuma operação encontrada para “{q.strip()}” nos fundos filtrados.</p>', unsafe_allow_html=True)
        else:
            linhas_q = x[x["Chave"].isin(chaves)]
            st.markdown(f'<div class="sumline">{len(res)} operação(ões) · em {linhas_q["Fundo"].nunique()} fundo(s) · {brl(res["Volume"].sum())} somados</div>', unsafe_allow_html=True)
            limite = 20
            for _, op in res.head(limite).iterrows():
                a = linhas_q[linhas_q["Chave"] == op["Chave"]].sort_values("Valor", ascending=False)
                nomes = list(dict.fromkeys(a["Ativo"]))
                nf = int(op["Nº fundos"])
                badges = [f'<span class="bdg">{op["Indexador"]}</span>', f'<span class="bdg">{op["Instrumento"]}</span>',
                          f'<span class="bdg">{nf} fundo{"s" if nf > 1 else ""}</span>',
                          f'<span class="bdg">{taxa_txt(op["Taxa"], op["Indexador"])} média</span>' if pd.notna(op["Taxa"]) else ""]
                tx = a.loc[a["Indexador"] == op["Indexador"], "Taxa"].dropna()
                if nf > 1 and len(tx) > 1:
                    bps = (tx.max() - tx.min()) * 100
                    cls = "hi" if bps >= 200 else ("md" if bps >= 50 else "lo")
                    badges.append(f'<span class="bdg {cls}">{num_br(bps, 0)} bps entre fundos</span>')
                with st.container(border=True):
                    st.markdown(f'<div class="sc-id">{op["Codigo"] if len(str(op["Codigo"])) > 3 else "Sem código na planilha"}</div>'
                                f'<div class="sc-d">{nomes[0]}</div>'
                                + (f'<div class="sc-alias">Também na planilha como: {", ".join(nomes[1:])}</div>' if len(nomes) > 1 else "")
                                + f'<div class="sc-meta">{"".join(badges)}</div>', unsafe_allow_html=True)
                    tb = pd.DataFrame({"Fundo": a["Fundo"].values})
                    if any(GESTORA.get(f) for f in a["Fundo"]):
                        tb["Gestora"] = [GESTORA.get(f, "") or "—" for f in a["Fundo"]]
                    tb["Taxa"] = [taxa_txt(t, i) for t, i in zip(a["Taxa"], a["Indexador"])]
                    if a["Vencimento"].notna().any():
                        tb["Vencimento"] = a["Vencimento"].dt.strftime("%d/%m/%Y").fillna("—").values
                    tb["Valor"] = (a["Valor"] / 1e3).round(0).values
                    tb["% da carteira do fundo"] = [round(v / df.loc[df["Fundo"] == f, "Valor"].sum() * 100, 2) for v, f in zip(a["Valor"], a["Fundo"])]
                    pct_pl = [(p if pd.notna(p) else (v / PL_FUNDO[f] * 100 if PL_FUNDO.get(f) else np.nan)) for p, v, f in zip(a["PctPL"], a["Valor"], a["Fundo"])]
                    if any(pd.notna(p) for p in pct_pl):
                        tb["% do PL"] = np.round(pct_pl, 2)
                    if a["Mes"].astype(str).str.len().gt(0).any():
                        tb["Mês"] = a["Mes"].values
                    if COMPACTO:
                        tb = tb[["Fundo", "Taxa", "Valor"]]
                    tabela(tb, column_config={"Valor": st.column_config.NumberColumn("Valor (R$ mil)", format="localized"),
                                              "% da carteira do fundo": st.column_config.NumberColumn(format="%.2f%%"),
                                              "% do PL": st.column_config.NumberColumn(format="%.2f%%")})
            if len(res) > limite:
                st.markdown(f'<p class="note">Mostrando as {limite} maiores de {len(res)} operações. Refine a busca para ver as demais.</p>', unsafe_allow_html=True)
            st.markdown('<div class="src">Diferença de taxa entre fundos, no mesmo indexador: 200 bps ou mais em vermelho, de 50 a 199 bps em amarelo, abaixo de 50 bps em verde. '
                        'Fonte: carteiras informadas nos relatórios de gestão dos fundos.</div>', unsafe_allow_html=True)

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
    cart["Valor"] = (cart["Valor"] / 1e3).round(0)
    if COMPACTO:
        cols = ["Fundo", "Ativo", "Codigo", "Valor", "Taxa"]
    tabela(cart[cols], height=460 if COMPACTO else 560, column_config={
        "Codigo": st.column_config.TextColumn("Código"),
        "Valor": st.column_config.NumberColumn("Valor (R$ mil)", format="localized"),
        "% do fundo": st.column_config.NumberColumn("% do fundo", format="%.2f%%"),
        "Vencimento": st.column_config.DateColumn(format="DD/MM/YYYY"),
    })
    exp = df.rename(columns={"Codigo": "codigo_ativo", "Taxa": "taxa_aa"}).copy()
    exp.columns = [norm(c) if c not in ("codigo_ativo", "taxa_aa") else c for c in exp.columns]
    c1, c2 = st.columns([1, 3])
    c1.download_button("Exportar CSV (pt-BR)", exp.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
                       file_name="carteiras_filtradas.csv", mime="text/csv", type="primary")
    c2.markdown('<p class="note">Exporta o que está filtrado, com separador ponto e vírgula e vírgula decimal, pronto para o Excel em português.</p>', unsafe_allow_html=True)
