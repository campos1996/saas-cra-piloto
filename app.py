import streamlit as st
import pandas as pd

st.set_page_config(page_title="Monitor de Crédito Privado - CRA", layout="wide")
st.title("📊 Monitor de Crédito - CRA ")
st.write("Painel conectado à base de dados em Excel.")

# Função para carregar os dados direto da planilha Excel no GitHub
@st.cache_data
def carregar_dados():
    # O Pandas lê o arquivo que está na mesma pasta do repositório
    df = pd.read_csv("carteira_ativos.csv")
    return df

# Tentamos carregar os dados
try:
    df = carregar_dados()

    # Métricas principais calculadas automaticamente com os dados reais
    col1, col2, col3 = st.columns(3)
    col1.metric("Valor Total Alocado", f"R$ {df['Valor'].sum():,.2f}")
    col2.metric("Quantidade de Ativos", len(df))
    col3.metric("Cupom Médio", f"{df['Cupom'].mean():,.2f}% a.a.")

    # Tabela interativa com os dados reais da planilha
    st.dataframe(df, use_container_width=True)

except FileNotFoundError:
    st.error("⚠️ O arquivo 'carteira_ativos.csv' não foi encontrado no repositório do GitHub. Por favor, faça o upload dele.")
