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