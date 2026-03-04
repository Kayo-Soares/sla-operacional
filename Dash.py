import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl

from Sla_Engine import calcular_sla, gerar_resumo

# =================================================
# CONFIG STREAMLIT
# =================================================
st.set_page_config(
    page_title="SLA Operacional",
    layout="wide"
)

st.title("📊 SLA Operacional – Análise Automática")

# =================================================
# COLUNAS PADRÃO (BASE FIXA)
# =================================================
COLUNAS_TABELA = [
    "Nome da base de entrega",   # coluna âncora
    "Sua cidade",
    "Coordenador",
    "Responsavel da base",
    "Qtd a entregar",
    "Qtd Entregues no prazo",
    "Pacotes Problematicos",
    "Qtd de pacotes faltantes para 96%",
    "SLA (%)",
    "Status SLA",
]

# =================================================
# BASE FIXA – REFERÊNCIA ORGANIZACIONAL (PROCV)
# =================================================
BASES_ORGANIZACAO = pd.DataFrame([
    # (lista mantida exatamente como você enviou)
    ("AUG -TO","Augustinopolis","ANA CUNHA","DHEYSONMAR FEITOSA LIMA"),
    ("AUX -TO","Araguaína","ANA CUNHA","DHEYSONMAR FEITOSA LIMA"),
    ("CDT -TO","Colinas do Tocantins","ANA CUNHA","DHEYSONMAR FEITOSA LIMA"),
    ("F GAI-TO","Guaraí","ANA CUNHA","DHEYSONMAR FEITOSA LIMA"),
    ("F DOM -PA","Dom Eliseu","ANA CUNHA","DHEYSONMAR FEITOSA LIMA"),
    ("F TLA-PA","Tailândia","ANA CUNHA","DHEYSONMAR FEITOSA LIMA"),
    ("F GNS-PA","Goianésia do Pará","ANA CUNHA","DHEYSONMAR FEITOSA LIMA"),
    ("TO PMW","Palmas","ANA CUNHA","Franciele Sousa Santos"),
    ("DNP -TO","Dianópolis","ANA CUNHA","Franciele Sousa Santos"),
    ("F PNA-TO","Porto Nacional","ANA CUNHA","Franciele Sousa Santos"),
    ("PMW 003-TO","Palmas","ANA CUNHA","Franciele Sousa Santos"),
    ("PMW 002-TO","Palmas","ANA CUNHA","Franciele Sousa Santos"),
    ("F GRP-TO","Gurupi","ANA CUNHA","Franciele Sousa Santos"),
    ("F PDT-TO","Paraíso do Tocantins","ANA CUNHA","Franciele Sousa Santos"),
    ("SMG -PA","São Miguel do Guamá","ANA CUNHA","JADSON OLIVEIRA DA CRUZ"),
    ("F PGM-PA","Paragominas","ANA CUNHA","JADSON OLIVEIRA DA CRUZ"),
    ("F CNC-PA","Concórdia do Pará","ANA CUNHA","JADSON OLIVEIRA DA CRUZ"),
    ("F GFN-PA","Garrafão do Norte","ANA CUNHA","JADSON OLIVEIRA DA CRUZ"),
    ("F MDR-PA","Mãe do Rio","ANA CUNHA","JADSON OLIVEIRA DA CRUZ"),
    ("F IPX-PA","Ipixuna do Pará","ANA CUNHA","JADSON OLIVEIRA DA CRUZ"),
    ("CPP -PA","Capitão Poço","ANA CUNHA","JADSON OLIVEIRA DA CRUZ"),
    ("BRG -PA","Bragança","ANA CUNHA","JADSON OLIVEIRA DA CRUZ"),
    ("CPN -PA","Capanema","ANA CUNHA","JADSON OLIVEIRA DA CRUZ"),
    ("SLP -PA","Salinópolis","ANA CUNHA","JADSON OLIVEIRA DA CRUZ"),
    ("VSU -PA","Viseu","ANA CUNHA","JADSON OLIVEIRA DA CRUZ"),
    ("BRV -PA","Breves","ANA CUNHA","VALDIR VIEIRA CORREA JUNIOR"),
    ("F BRV-PA","Breves","ORLEAN NASCIMENTO","PRICILA DO ESPIRITO SANTO DE LIMA"),
    ("SDA -PA","São Domingos do Araguaia","JOSÉ MARLON","FERNANDO SILVA OLIVEIRA"),
    ("F PCA-PA","Pacajá","JOSÉ MARLON","FERNANDO SILVA OLIVEIRA"),
    ("F ITI-PA","Itupiranga","JOSÉ MARLON","FERNANDO SILVA OLIVEIRA"),
    ("F JCD-PA","Jacundá","JOSÉ MARLON","FERNANDO SILVA OLIVEIRA"),
    ("F TUR-PA","Tucuruí","JOSÉ MARLON","FERNANDO SILVA OLIVEIRA"),
    ("NRE -PA","Novo Repartimento","JOSÉ MARLON","FERNANDO SILVA OLIVEIRA"),
    ("F NMB-PA","Marabá","JOSÉ MARLON","FERNANDO SILVA OLIVEIRA"),
    ("NMB -PA","Marabá","JOSÉ MARLON","FERNANDO SILVA OLIVEIRA"),
    ("AMP -PA","Marabá","JOSÉ MARLON","FERNANDO SILVA OLIVEIRA"),
    ("CNA -PA","Canaã dos Carajás","JOSÉ MARLON","MARIA NAGILA ARAUJO BASTOS"),
    ("CKS -PA","Parauapebas","JOSÉ MARLON","MARIA NAGILA ARAUJO BASTOS"),
    ("F ORL-PA","Ourilândia do Norte","JOSÉ MARLON","MARIA NAGILA ARAUJO BASTOS"),
    ("F SFX-PA","São Félix do Xingu","JOSÉ MARLON","MARIA NAGILA ARAUJO BASTOS"),
    ("F RDC -PA","Redenção","JOSÉ MARLON","MARIA NAGILA ARAUJO BASTOS"),
    ("F XIG-PA","Xinguara","JOSÉ MARLON","MARIA NAGILA ARAUJO BASTOS"),
    ("F ELD-PA","Eldorado do Carajás","JOSÉ MARLON","MARIA NAGILA ARAUJO BASTOS"),
    ("PA MRB","Marabá","JOSÉ MARLON","RAIMUNDO LIMA FILHO"),
    ("F DOM -PA","Dom Eliseu","JOSÉ MARLON","FERNANDO SILVA OLIVEIRA"),
    ("F TLA-PA","Tailândia","JOSÉ MARLON","FERNANDO SILVA OLIVEIRA"),
    ("F GNS -PA","Goianésia do Pará","JOSÉ MARLON","FERNANDO SILVA OLIVEIRA"),
    ("MCP FLUVIAL -AP","Macapá","ORLEAN NASCIMENTO","PRICILA DO ESPIRITO SANTO DE LIMA"),
    ("F MCP-AP","Macapá","ORLEAN NASCIMENTO","PRICILA DO ESPIRITO SANTO DE LIMA"),
    ("F MCP 02-AP","Macapá","ORLEAN NASCIMENTO","PRICILA DO ESPIRITO SANTO DE LIMA"),
    ("F MAC-AP","Macapá","ORLEAN NASCIMENTO","PRICILA DO ESPIRITO SANTO DE LIMA"),
    ("F BAO-PA","Baião","ORLEAN NASCIMENTO","CLAUDIO ROBERTO PANTOJA DE LIMA JUNIOR"),
    ("VCP -PA","Cametá","ORLEAN NASCIMENTO","CLAUDIO ROBERTO PANTOJA DE LIMA JUNIOR"),
    ("MCJ -PA","Mocajuba","ORLEAN NASCIMENTO","CLAUDIO ROBERTO PANTOJA DE LIMA JUNIOR"),
    ("PA ANA","Ananindeua","ORLEAN NASCIMENTO","DARLEY WALLLACE CUNHA QUARESMA"),
    ("BVD -PA","Benevides","ORLEAN NASCIMENTO","JOHAN RAFAEL QUEIROZ PACHECO"),
    ("VGA -PA","Vigia","ORLEAN NASCIMENTO","JOHAN RAFAEL QUEIROZ PACHECO"),
    ("CST -PA","Castanhal","ORLEAN NASCIMENTO","JOHAN RAFAEL QUEIROZ PACHECO"),
    ("F IGA-PA","Igarapé Açú","ORLEAN NASCIMENTO","JOHAN RAFAEL QUEIROZ PACHECO"),
    ("BRC -PA","Barcarena","ORLEAN NASCIMENTO","LOURIVAL RAIOL PORTAL FILHO"),
    ("ABT -PA","Abaetetuba","ORLEAN NASCIMENTO","LOURIVAL RAIOL PORTAL FILHO"),
    ("IGM -PA","Igarapé-Miri","ORLEAN NASCIMENTO","LOURIVAL RAIOL PORTAL FILHO"),
    ("F IGM-PA","Igarapé-Miri","ORLEAN NASCIMENTO","LOURIVAL RAIOL PORTAL FILHO"),
    ("MJU -PA","Moju","ORLEAN NASCIMENTO","LOURIVAL RAIOL PORTAL FILHO"),
    ("MRM -PA","Belém","ORLEAN NASCIMENTO","PRICILA DO ESPIRITO SANTO DE LIMA"),
    ("PDR -PA","Belém","ORLEAN NASCIMENTO","PRICILA DO ESPIRITO SANTO DE LIMA"),
    ("BEL -PA","Belém","ORLEAN NASCIMENTO","PRICILA DO ESPIRITO SANTO DE LIMA"),
    ("F SBZ-PA","Belém","ORLEAN NASCIMENTO","PRICILA DO ESPIRITO SANTO DE LIMA"),
    ("F PDP-PA","Ponta de Pedras","ORLEAN NASCIMENTO","PRICILA DO ESPIRITO SANTO DE LIMA"),
    ("ANA FLUVIAL - PA","Ananindeua","ORLEAN NASCIMENTO","PRICILA DO ESPIRITO SANTO DE LIMA"),
    ("F CRH-PA","Curralinho","ORLEAN NASCIMENTO","PRICILA DO ESPIRITO SANTO DE LIMA"),
    ("ICR -PA","Belém","ORLEAN NASCIMENTO","WALLESON BRAGA VIEIRA"),
    ("ANA -PA","Ananindeua","ORLEAN NASCIMENTO","WALLESON BRAGA VIEIRA"),
    ("F MTB-PA","Marituba","ORLEAN NASCIMENTO","WALLESON BRAGA VIEIRA"),
    ("F ANA-PA","Ananindeua","ORLEAN NASCIMENTO","WALLESON BRAGA VIEIRA"),
    ("COQ -PA","Ananindeua","ORLEAN NASCIMENTO","WALLESON BRAGA VIEIRA"),
    ("F TPN-PA","Belém","ORLEAN NASCIMENTO","WALLESON BRAGA VIEIRA"),
], columns=["Nome da base de entrega","Sua cidade","Coordenador","Responsavel da base"])


# =================================================
# FUNÇÃO COM CACHE
# =================================================
@st.cache_data(show_spinner=False)
def carregar_sla(arquivo_bytes):
    wb = openpyxl.load_workbook(arquivo_bytes, data_only=True)
    sheet = wb.active
    df = pd.DataFrame(sheet.values)
    df.columns = df.iloc[0]
    return df[1:].reset_index(drop=True)


# =================================================
# MENU
# =================================================
aba = st.sidebar.radio("Navegação", ["Visão Geral", "Resumo por Coordenador"])

# =================================================
# UPLOAD
# =================================================
arquivo = st.file_uploader("📄 Envie o arquivo de SLA", type=["xlsx"])

if arquivo:
    try:
        df_sla = carregar_sla(BytesIO(arquivo.read()))
        df = calcular_sla(df_sla)

        df["Qtd com tentativa de entrega"] = (
    pd.to_numeric(df["Qtd com tentativa de entrega"], errors="coerce")
    .fillna(0))
        
        df = (
    df
    .sort_values(
        by=[
            "Nome da base de entrega",
            "Qtd com tentativa de entrega",
            "Qtd a entregar"
        ],
        ascending=[True, False, False]
    )
    .drop_duplicates(
        subset=["Nome da base de entrega"],
        keep="first"
    )
)

        df = df.merge(BASES_ORGANIZACAO, on="Nome da base de entrega", how="left")
        df = df[df["Responsavel da base"].notna()]

        # ================= VISÃO GERAL =================
        if aba == "Visão Geral":

            c1, c2, c3 = st.columns(3)

            with c1:
                coord = st.selectbox(
                    "Coordenador",
                    ["Todos"] + sorted(df["Coordenador"].unique())
                )

            df_tmp = df.copy()
            if coord != "Todos":
                df_tmp = df_tmp[df_tmp["Coordenador"] == coord]

            with c2:
                resp = st.selectbox(
                    "Responsável",
                    ["Todos"] + sorted(df_tmp["Responsavel da base"].unique())
                )

            with c3:
                apenas_criticos = st.checkbox("Somente bases fora do SLA")

            df_f = df_tmp.copy()
            if resp != "Todos":
                df_f = df_f[df_f["Responsavel da base"] == resp]
            if apenas_criticos:
                df_f = df_f[df_f["SLA (%)"] < 96]

            resumo = gerar_resumo(df_f)

            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Total a entregar", resumo["Total a entregar"])
            k2.metric("Entregues no prazo", resumo["Entregues no prazo"])
            k3.metric("Pacotes problemáticos", resumo["Pacotes problemáticos"])
            k4.metric("SLA Geral (%)", f"{resumo['SLA Geral (%)']}%")
            k5.metric(
                "Bases fora do SLA",
                df_f[df_f["SLA (%)"] < 96]["Nome da base de entrega"].nunique()
            )

            st.subheader("🔥 Ranking das piores bases")
            ranking = (
    df_f[df_f["SLA (%)"] < 96]
    .sort_values("SLA (%)")
    .head(10)
)
            st.dataframe(
                ranking[
                    [
                        "Nome da base de entrega",
                        "Sua cidade",
                        "Coordenador",
                        "Responsavel da base",
                        "SLA (%)",
                        "Status SLA"
                    ]
                ].reset_index(drop=True),
                use_container_width=True
            )

            # EXPORTAÇÃO
            buffer = BytesIO()
            df_f[COLUNAS_TABELA].to_excel(buffer, index=False)

            st.download_button(
                "📤 Exportar Excel filtrado",
                buffer.getvalue(),
                file_name="sla_filtrado.xlsx"
            )

            st.divider()

            st.dataframe(
                df_f[COLUNAS_TABELA]
                .sort_values("Nome da base de entrega")
                .reset_index(drop=True),
                use_container_width=True,
            )

        # ================= RESUMO POR COORDENADOR =================
        else:
            resumo_coord = (
                df.groupby("Coordenador")
                .agg(
                    Bases=("Nome da base de entrega", "nunique"),
                    SLA_Médio=("SLA (%)", "mean"),
                    Bases_fora_SLA=("SLA (%)", lambda x: (x < 96).sum())
                )
                .reset_index()
                .sort_values("SLA_Médio")
            )

            resumo_coord["SLA_Médio"] = resumo_coord["SLA_Médio"].round(2)

            st.subheader("📌 Resumo por Coordenador")
            st.dataframe(
                resumo_coord.reset_index(drop=True),
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
