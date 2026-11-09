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

  if df.empty or "Valor" not in df.columns or "Fundo" not in df.columns:
    st.warning(
        "⚠️ O seu arquivo 'carteira_ativos.xlsx' está vazio ou sem as colunas"
        " corretas ('Fundo', 'Ativo', 'Valor', etc.)."
    )
  else:
    # Tratamento de dados numéricos
    df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)

    # Tratamento da coluna Cupom para exibição percentual limpa
    if "Cupom" in df.columns:
      df["Cupom_Num"] = pd.to_numeric(
          df["Cupom"].astype(str).str.replace("%", "").str.replace(",", "."),
          errors="coerce",
      ).fillna(0)
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

    # Filtragem do DataFrame
    if fundo_selecionado != "Todos os Fundos":
      df_filtrado = df[df["Fundo"] == fundo_selecionado].copy()
      st.sidebar.info(
          f"Exibindo dados detalhados para: **{fundo_selecionado}**"
      )
    else:
      df_filtrado = df.copy()
      st.sidebar.info("Exibindo consolidação de **todos os fundos**.")

    # Cálculo do percentual do ativo em relação ao total exibido (Fundo ou Consolidado)
    total_valor = df_filtrado["Valor"].sum()
    if total_valor > 0:
      df_filtrado["% do Fundo"] = (
          df_filtrado["Valor"] / total_valor
      ) * 100
    else:
      df_filtrado["% do Fundo"] = 0.0

    # Ordenação decrescente por Valor
    df_filtrado = df_filtrado.sort_values(by="Valor", ascending=False).reset_index(
        drop=True
    )

    # --- MÉTRICAS PRINCIPAIS ---
    col1, col2 = st.columns(2)
    col1.metric("Valor Total Alocado", f"R$ {total_valor:,.2f}")
    col2.metric("Quantidade de Ativos", len(df_filtrado))

    st.markdown("---")

    # --- SEÇÃO PRINCIPAL ---
    if fundo_selecionado == "Todos os Fundos":
      # PAINEL GERAL: RANKINGS DOS 10 MAIORES E 10 MENORES ATIVOS
      st.subheader("🏆 Ranking de Ativos (Visão Consolidada)")

      rcol1, rcol2 = st.columns(2)

      with rcol1:
        st.markdown("### Top 10 Maiores Ativos (R$)")
        top_10_maiores = df_filtrado.head(10)
        fig_maiores = px.bar(
            top_10_maiores.sort_values("Valor", ascending=True),
            x="Valor",
            y="Ativo",
            orientation="h",
            text_auto=".2s",
            color="Valor",
            color_continuous_scale="Teal",
        )
        fig_maiores.update_layout(
            showlegend=False,
            xaxis_title="Valor (R$)",
            yaxis_title="",
            margin=dict(t=10, b=10, l=10, r=10),
        )
        st.plotly_chart(fig_maiores, use_container_width=True)

      with rcol2:
        st.markdown("### Top 10 Menores Ativos (R$)")
        # Filtra valores maiores que zero para evitar lixo irrelevante no menor ranking
        top_10_menores = (
            df_filtrado[df_filtrado["Valor"] > 0].tail(10).sort_values("Valor", ascending=True)
        )
        fig_menores = px.bar(
            top_10_menores,
            x="Valor",
            y="Ativo",
            orientation="h",
            text_auto=".2s",
            color="Valor",
            color_continuous_scale="Blues",
        )
        fig_menores.update_layout(
            showlegend=False,
            xaxis_title="Valor (R$)",
            yaxis_title="",
            margin=dict(t=10, b=10, l=10, r=10),
        )
        st.plotly_chart(fig_menores, use_container_width=True)

    else:
      # PAINEL ESPECÍFICO DO FUNDO: GRÁFICO DE BARRAS LIMPO DOS ATIVOS
      st.subheader(f"📈 Alocação por Ativo - {fundo_selecionado}")
      fig_bar = px.bar(
          df_filtrado.sort_values("Valor", ascending=True),
          x="Valor",
          y="Ativo",
          orientation="h",
          text_auto=".2s",
          color="Valor",
          color_continuous_scale="Teal",
      )
      fig_bar.update_layout(
          xaxis_title="Valor Alocado (R$)",
          yaxis_title="",
          showlegend=False,
          margin=dict(t=20, b=20, l=20, r=20),
          height=max(400, len(df_filtrado) * 25),  # Ajusta altura dinamicamente
      )
      st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("---")

    # --- TABELA INTERATIVA ORDENADA COM % DO FUNDO ---
    st.
