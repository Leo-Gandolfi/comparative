
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Comparativo SAP x CSOD", layout="wide")
st.title("📊 Comparativo SAP x CSOD")

# Upload dos arquivos
col1, col2 = st.columns(2)
with col1:
    csod_file = st.file_uploader("📤 Enviar arquivo da CSOD (.xlsx)", type=["xlsx"], key="csod")
with col2:
    sap_file = st.file_uploader("📤 Enviar arquivo do SAP (.xlsx)", type=["xlsx"], key="sap")

# Processamento
if csod_file and sap_file:
    st.success("Arquivos carregados com sucesso!")

    # CSOD: começa na linha 8, ou seja header=7
    csod = pd.read_excel(csod_file, header=8)
    sap = pd.read_excel(sap_file)

    # Normaliza nomes de colunas
    # Normaliza nomes de colunas: remove espaços, quebras de linha, múltiplos espaços
    csod.columns = csod.columns.str.replace(r"\\n|\\r|\\t", "", regex=True)
    csod.columns = csod.columns.str.replace(r"\s+", " ", regex=True).str.strip()
    sap.columns = sap.columns.str.replace(r"\s+", " ", regex=True).str.strip()


    # Validação
    required_csod_cols = ["ID do Usuário", "Posição"]
    required_sap_cols = ["NP", "Cargo"]

    if not all(col in csod.columns for col in required_csod_cols):
        st.error(f"Arquivo da CSOD deve conter as colunas: {', '.join(required_csod_cols)}")
        st.stop()
    if not all(col in sap.columns for col in required_sap_cols):
        st.error(f"Arquivo do SAP deve conter as colunas: {', '.join(required_sap_cols)}")
        st.stop()

    # Renomeia para padronizar
    csod.rename(columns={"ID do Usuário": "ID", "Posição": "Cargo"}, inplace=True)
    sap.rename(columns={"NP": "ID"}, inplace=True)

    # Limpa os IDs
    csod["ID"] = csod["ID"].astype(str).str.strip()
    sap["ID"] = sap["ID"].astype(str).str.strip()

    # 1. Usuários sem cargo
    usuarios_sem_cargo = csod[csod["Cargo"].isna() | (csod["Cargo"].astype(str).str.strip() == "")]

    # 2. CSOD que não existe no SAP
    ids_csod = set(csod["ID"])
    ids_sap = set(sap["ID"])
    csod_nao_existe_no_sap = csod[csod["ID"].isin(ids_csod - ids_sap)]

    # 3. SAP que não existe na CSOD
    sap_nao_existe_no_csod = sap[sap["ID"].isin(ids_sap - ids_csod)]

    # 4. Cargos divergentes
    comparativo = pd.merge(
        csod[["ID", "Cargo"]],
        sap[["ID", "Cargo"]],
        on="ID",
        how="inner",
        suffixes=("_CSOD", "_SAP")
    )
    cargos_divergentes = comparativo[
        comparativo["Cargo_CSOD"].fillna("").str.strip() != comparativo["Cargo_SAP"].fillna("").str.strip()
    ]

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

    # Gerar Excel para download
    with pd.ExcelWriter("comparativo_sap_csod.xlsx", engine="openpyxl") as writer:
        usuarios_sem_cargo.to_excel(writer, index=False, sheet_name="Usuários sem cargo")
        csod_nao_existe_no_sap.to_excel(writer, index=False, sheet_name="CSOD não existe no SAP")
        sap_nao_existe_no_csod.to_excel(writer, index=False, sheet_name="SAP não existe na CSOD")
        cargos_divergentes.to_excel(writer, index=False, sheet_name="Cargos divergentes")

    with open("comparativo_sap_csod.xlsx", "rb") as f:
        st.download_button(
            label="📥 Baixar resultado em Excel",
            data=f,
            file_name="comparativo_sap_csod.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("Por favor, envie os dois arquivos para gerar o comparativo.")
