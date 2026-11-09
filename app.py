import pandas as pd
import streamlit as st

st.set_page_config(page_title="Monitor de Crédito Privado - CRA", layout="wide")
st.title("📊 Monitor de Ativos de Crédito")
st.write("Painel conectado à base de dados em Excel.")

# Função para carregar os dados do arquivo Excel
@st.cache_data
def carregar_dados():
    # Agora o Pandas vai ler o arquivo Excel
    df = pd.read_excel("carteira_ativos.xlsx")
    return df

try:
    df = carregar_dados()

    # Verifica se o arquivo Excel está vazio ou sem as colunas principais
    if df.empty or "Valor" not in df.columns:
        st.warning(
            "⚠️ O seu arquivo 'carteira_ativos.xlsx' está vazio ou sem as colunas "
            "corretas. Preencha a planilha e suba novamente no GitHub."
        )
    else:
        # Métricas principais calculadas automaticamente
        col1, col2, col3 = st.columns(3)

        total_valor = df["Valor"].sum() if "Valor" in df.columns else 0
        qtd_ativos = len(df)
        
        # Tratamento especial para a coluna Cupom (para lidar com o '%' e vírgulas do Excel)
        if "Cupom" in df.columns:
            # Transforma tudo em texto, tira o '%' e troca vírgula por ponto para o Python calcular
            cupom_limpo = pd.to_numeric(
                df['Cupom'].astype(str).str.replace('%', '').str.replace(',', '.'), 
                errors='coerce'
            )
            media_cupom = cupom_limpo.mean()
            
            # Se a média der menor que 1 (ex: 0.05), multiplicamos por 100 para virar 5.0%
            if media_cupom < 1:
                media_cupom = media_cupom * 100
        else:
            media_cupom = 0

        # Exibição na tela
        col1.metric("Valor Total Alocado", f"R$ {total_valor:,.2f}")
        col2.metric("Quantidade de Ativos", qtd_ativos)
        col3.metric("Cupom Médio", f"{media_cupom:,.2f}% a.a.")

        # Tabela interativa com os dados reais da planilha
        st.dataframe(df, use_container_width=True)

except Exception as e:
    st.error(
        "⚠️ O arquivo 'carteira_ativos.xlsx' não foi encontrado ou está com "
        f"formato inválido. Por favor, verifique o upload no GitHub. Erro: {e}"
    )
