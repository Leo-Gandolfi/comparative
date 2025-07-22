import streamlit as st
import pandas as pd

st.set_page_config(page_title="Comparativo SAP x CSOD", layout="wide")
st.title("📊 Comparativo SAP x CSOD")

col1, col2 = st.columns(2)
with col1:
    csod_file = st.file_uploader("📤 Enviar arquivo da CSOD (.xlsx)", type=["xlsx"])
with col2:
    sap_file = st.file_uploader("📤 Enviar arquivo do SAP (.xlsx)", type=["xlsx"])

def encontrar_header_csod(file):
    for i in range(10):
        df_test = pd.read_excel(file, header=i, nrows=1)
        if "ID do Usuário" in df_test.columns and "Posição ID" in df_test.columns:
            return i
    return None

if csod_file and sap_file:
    st.success("Arquivos carregados com sucesso!")

    # Detecta a linha do cabeçalho na CSOD
    header_row = encontrar_header_csod(csod_file)
    if header_row is None:
        st.error("Colunas obrigatórias não encontradas no arquivo da CSOD.")
        st.stop()

    # Leitura dos dados
    csod = pd.read_excel(csod_file, header=header_row)
    sap = pd.read_excel(sap_file)

    # Normalizar nomes de colunas
    csod.columns = csod.columns.astype(str).str.replace(r"\n|\r|\t", "", regex=True)
    csod.columns = csod.columns.str.replace(r"\s+", " ", regex=True).str.strip()
    sap.columns = sap.columns.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()

    # 🔍 Debug: Mostrar colunas dos arquivos
    with st.expander("🔧 Debug - Colunas dos arquivos"):
        st.write("📄 Colunas CSOD:", list(csod.columns))
        st.write("📄 Colunas SAP:", list(sap.columns))



    # Verificar colunas obrigatórias
    if "ID do Usuário" not in csod.columns or "Posição ID" not in csod.columns:
        st.error("Arquivo da CSOD deve conter as colunas: 'ID do Usuário' e 'Posição ID'")
        st.stop()
    if "NP" not in sap.columns or "Cargo - Cód." not in sap.columns:
        st.error("Arquivo do SAP deve conter as colunas: 'NP' e 'Cargo - Cód.'")
        st.stop()

    # Renomear colunas para padronizar
    csod.rename(columns={"ID do Usuário": "ID"}, inplace=True)
    sap.rename(columns={"NP": "ID"}, inplace=True)
# Remover usuários afastados com base na descrição do centro de custo
    #sap = sap[~sap["Desc. C. Custo"].str.lower().str.contains("afastad", na=False)]

    # Limpeza forte dos IDs
    def limpar_ids(df, coluna):
        df[coluna] = (
            df[coluna]
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)  # remove final .0
            .str.extract(r"(\d+)", expand=False)   # extrai apenas números
            .str.strip()
        )
        return df
    
    csod = limpar_ids(csod, "ID")
    sap = limpar_ids(sap, "ID")

    # Remove IDs inválidos
    ids_invalidos = ("100008", "89", "70")

    csod = csod[csod["ID"].notna()]
    # 1. Filtrar IDs com descrição de custo que indicam afastamento
    ids_afastados = sap[sap["Desc. C. Custo"].str.lower().str.contains("afastad", na=False)]["ID"].unique()

    # 2. Manter apenas IDs válidos e não afastados
    sap = sap[sap["ID"].notna()]
    sap = sap[~sap["ID"].isin(ids_afastados)]
    csod = csod[~csod["ID"].isin(ids_afastados)]


    csod = csod[~csod["ID"].astype(str).str.startswith(ids_invalidos)]
    sap = sap[~sap["ID"].astype(str).str.startswith(ids_invalidos)]


    # 🔍 Debug: IDs válidos
    with st.expander("🔧 Debug - IDs válidos"):
        st.write("✅ Total IDs únicos CSOD:", csod["ID"].nunique())
        st.write("✅ Total IDs únicos SAP:", sap["ID"].nunique())
        st.write("IDs CSOD exemplo:", csod["ID"].unique()[:5])
        st.write("IDs SAP exemplo:", sap["ID"].unique()[:5])

    # 1. Usuários sem cargo na CSOD
    usuarios_sem_cargo = csod[csod["Posição ID"].isna() | (csod["Posição ID"].astype(str).str.strip() == "")]

    # 2. CSOD que não existe no SAP
    ids_csod = set(csod["ID"])
    ids_sap = set(sap["ID"])

    csod_nao_existe_no_sap = csod[csod["ID"].isin(ids_csod - ids_sap)]

    # 3. SAP que não existe na CSOD
    sap_nao_existe_no_csod = sap[sap["ID"].isin(ids_sap - ids_csod)]

    # 4. Cargos divergentes (por ID de cargo)
    def extrair_primeiro_id_cargo(valor):
        if pd.isna(valor):
            return ""
        parte = str(valor).split(",")[0]        # pega o primeiro cargo (antes da vírgula)
        parte = parte.split("-")[0].strip()     # pega o que está antes do "-" dentro dele
        return parte.zfill(8)

    csod["Cargo_ID"] = csod["Posição ID"].apply(extrair_primeiro_id_cargo)

    sap["Cargo_ID"] = sap["Cargo - Cód."].apply(lambda x: str(int(x)).zfill(8) if pd.notna(x) else "")

    # 🔍 Debug: Cargo ID normalizado
    with st.expander("🔧 Debug - Cargo IDs normalizados"):
        st.write("CSOD Cargo_ID exemplo:", csod["Cargo_ID"].dropna().unique()[:5])
        st.write("SAP Cargo_ID exemplo:", sap["Cargo_ID"].dropna().unique()[:5])

    csod_validos = csod[["ID", "Cargo_ID"]].dropna()
    sap_validos = sap[["ID", "Cargo_ID"]].dropna()
    csod_validos = csod[["ID", "Cargo_ID"]].dropna()

    sap_validos = sap[["ID", "Cargo_ID"]].dropna()

    csod_validos = csod_validos[csod_validos["Cargo_ID"].str.strip() != ""]
    sap_validos = sap_validos[sap_validos["Cargo_ID"].str.strip() != ""]

    comparativo = pd.merge(csod_validos, sap_validos, on="ID", how="inner", suffixes=("_CSOD", "_SAP"))
    cargos_divergentes = comparativo[comparativo["Cargo_ID_CSOD"] != comparativo["Cargo_ID_SAP"]]

    # Exibição
    st.header("📌 Resumo do Comparativo")
    st.metric("Usuário sem cargo na CSOD", len(usuarios_sem_cargo))
    st.metric("Usuário da CSOD não existente no SAP", len(csod_nao_existe_no_sap))
    st.metric("Usuário do SAP não existente na CSOD", len(sap_nao_existe_no_csod))
    st.metric("Usuário com cargo diferente entre CSOD e SAP", len(cargos_divergentes))
    st.caption(f"🔎 Total de usuários na CSOD: {len(csod)} | Total no SAP: {len(sap)}")

    with st.expander("👥 Usuários sem cargo na CSOD"):
        st.dataframe(usuarios_sem_cargo)
    with st.expander("🔍 CSOD não existe no SAP"):
        st.dataframe(csod_nao_existe_no_sap)
    with st.expander("🧾 SAP não existe na CSOD"):
        st.dataframe(sap_nao_existe_no_csod)
    with st.expander("⚖️ Cargos divergentes"):
        st.dataframe(cargos_divergentes)

    with pd.ExcelWriter("comparativo_sap_csod_resultado.xlsx", engine="openpyxl") as writer:
        usuarios_sem_cargo.to_excel(writer, index=False, sheet_name="Usuários sem cargo")
        csod_nao_existe_no_sap.to_excel(writer, index=False, sheet_name="CSOD não existe no SAP")
        sap_nao_existe_no_csod.to_excel(writer, index=False, sheet_name="SAP não existe na CSOD")
        cargos_divergentes.to_excel(writer, index=False, sheet_name="Cargos divergentes")

    with open("comparativo_sap_csod_resultado.xlsx", "rb") as f:
        st.download_button(
            label="📥 Baixar resultado em Excel",
            data=f,
            file_name="comparativo_sap_csod_resultado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("Por favor, envie os dois arquivos para gerar o comparativo.")
