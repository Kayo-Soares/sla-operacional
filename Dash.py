import streamlit as st
import pandas as pd
from Sla_Engine import calcular_sla

st.set_page_config(page_title="SLA Operacional", layout="wide")
st.title("📊 SLA Operacional – Análise Automática")

arquivo = st.file_uploader("📄 Envie o arquivo de SLA", type=["xlsx"])

if arquivo:
    try:
        df_sla = pd.read_excel(arquivo, engine="openpyxl")
        df = calcular_sla(df_sla)

        responsaveis = sorted(df["Responsavel da base"].unique())
        resp_sel = st.selectbox("Filtrar por Responsável da base", ["Todos"] + responsaveis)

        if resp_sel != "Todos":
            df_filtrado = df[df["Responsavel da base"] == resp_sel]
        else:
            df_filtrado = df.copy()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total a entregar", int(df_filtrado["Qtd a entregar"].sum()))
        c2.metric("Entregues no prazo", int(df_filtrado["Qtd Entregues no prazo"].sum()))
        c3.metric("Pacotes problemáticos", int(df_filtrado["Pacotes Problematicos"].sum()))
        c4.metric("SLA Geral (%)", f"{df_filtrado['SLA (%)'].mean():.2f}%")

        st.divider()

        st.dataframe(
            df_filtrado[
                [
                    "Nome da base de entrega",
                    "Sua cidade",
                    "Responsavel da base",
                    "Coordenador",
                    "Qtd a entregar",
                    "Qtd Entregues no prazo",
                    "Pacotes Problematicos",
                    "Qtd de pacotes faltantes para 96%",
                    "SLA (%)",
                    "Status SLA",
                ]
            ],
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
