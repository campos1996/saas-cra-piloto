import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Monitor de Crédito Privado - CRA", layout="wide"
)
st.title("📊 Monitor de Ativos de Crédito Privado (CRA)")
st.write("Painel analítico conectado à base de dados em Excel.")


# Função para carregar os dados do arquivo Excel
@st.cache_data
def carregar_dados():
  df = pd.read_excel("carteira_ativos.xlsx")
  return df


try:
  df = carregar_dados()

  # Verifica se o arquivo Excel está vazio ou sem as colunas principais
  if df.empty or "Valor" not in df.columns or "Fundo" not in df.columns:
    st.warning(
        "⚠️ O seu arquivo 'carteira_ativos.xlsx' está vazio ou sem as colunas"
        " corretas ('Fundo', 'Ativo', 'Valor', etc.). Verifique a planilha."
    )
  else:
    # Tratamento de dados numéricos para garantir precisão nos cálculos e gráficos
    df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)

    # Tratamento da coluna Cupom (garantindo formato numérico percentual)
    if "Cupom" in df.columns:
      df["Cupom_Num"] = pd.to_numeric(
          df["Cupom"].astype(str).str.replace("%", "").str.replace(",", "."),
          errors="coerce",
      ).fillna(0)
      # Se os valores vierem em formato decimal (ex: 0.038), converte para percentual (3.8)
      if df["Cupom_Num"].max() <= 1.0 and df["Cupom_Num"].max() > 0:
        df["Cupom_Num"] = df["Cupom_Num"] * 100
    else:
      df["Cupom_Num"] = 0

    # --- BARRA LATERAL: SELEÇÃO DE FUNDO ---
    st.sidebar.header("🔍 Filtros de Navegação")
    lista_fundos = ["Todos os Fundos"] + sorted(
        df["Fundo"].dropna().unique().tolist()
    )
    fundo_selecionado = st.sidebar.selectbox(
        "Selecione o Fundo:", lista_fundos
    )

    # Filtragem do DataFrame com base na escolha do usuário
    if fundo_selecionado != "Todos os Fundos":
      df_filtrado = df[df["Fundo"] == fundo_selecionado]
      st.sidebar.info(
          f"Exibindo dados detalhados para: **{fundo_selecionado}**"
      )
    else:
      df_filtrado = df
      st.sidebar.info("Exibindo consolidação de **todos os fundos**.")

    # --- MÉTRICAS PRINCIPAIS ---
    col1, col2, col3 = st.columns(3)

    total_valor = df_filtrado["Valor"].sum()
    qtd_ativos = len(df_filtrado)
    cupom_medio = (
        df_filtrado["Cupom_Num"].mean() if not df_filtrado.empty else 0
    )

    col1.metric("Valor Total Alocado", f"R$ {total_valor:,.2f}")
    col2.metric("Quantidade de Ativos", qtd_ativos)
    col3.metric("Cupom Médio Ponderado", f"{cupom_medio:,.2f}% a.a.")

    st.markdown("---")

    # --- SEÇÃO DE GRÁFICOS ---
    if not df_filtrado.empty:
      gcol1, gcol2 = st.columns(2)

      with gcol1:
        st.subheader("📈 Alocação por Ativo (R$)")
        fig_pie = px.pie(
            df_filtrado,
            names="Ativo",
            values="Valor",
            hole=0.4,
            color_discrete_sequence=px.colors.sequential.Teal,
        )
        fig_pie.update_layout(margin=dict(t=20, b=20, l=20, r=20))
        st.plotly_chart(fig_pie, use_container_width=True)

      with gcol2:
        st.subheader("📊 Taxa de Cupom por Ativo (% a.a.)")
        fig_bar = px.bar(
            df_filtrado,
            x="Ativo",
            y="Cupom_Num",
            text="Cupom_Num",
            color="Cupom_Num",
            color_continuous_scale="Teal",
        )
        fig_bar.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
        fig_bar.update_layout(
            xaxis_title="Ativo",
            yaxis_title="Cupom (% a.a.)",
            showlegend=False,
            margin=dict(t=20, b=20, l=20, r=20),
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("---")

    # --- TABELA INTERATIVA ---
    st.subheader("📋 Detalhamento da Carteira")
    st.dataframe(df_filtrado, use_container_width=True)

except Exception as e:
  st.error(
      "⚠️ O arquivo 'carteira_ativos.xlsx' não foi encontrado ou apresenta"
      f" erro de estrutura. Detalhes do erro: {e}"
  )
