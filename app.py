import streamlit as st
import pandas as pd

st.set_page_config(page_title="Monitor de Crédito Privado - CRA", layout="wide")
st.title("📊 Monitor de Ativos de Crédito Privado (CRA)")
st.write("Painel protótipo para acompanhamento de carteira de ativos.")

# Exemplo de dados manuais baseado nos relatórios
dados_exemplo = {
    'Fundo': ['Fundo Exemplo', 'Fundo Exemplo', 'Fundo Exemplo'],
    'Ativo': ['CRA Fiagril VI', 'CRA Languiru', 'CRA Irmãos Walker'],
    'Emissor': ['Província', 'Opea', 'Opea'],
    'Segmento': ['Distribuidora', 'Cooperativa', 'Produtores'],
    'Valor': [99076894.41, 63837265.00, 53305877.52],
    'Coupon': [3.80, 6.75, 5.00],
    'Duration': [1.5, 4.6, 2.9]
}
df = pd.DataFrame(dados_exemplo)

# Métricas principais (KPIs)
col1, col2, col3 = st.columns(3)
col1.metric("Valor Total Alocado", f"R$ {df['Valor'].sum():,.2f}")
col2.metric("Quantidade de Ativos", len(df))
col3.metric("Coupon Médio", f"{df['Coupon'].mean():,.2f}% a.a.")

# Tabela interativa na tela
st.dataframe(df, use_container_width=True)
