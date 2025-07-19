import streamlit as st
import pandas as pd

st.set_page_config(page_title="Comparativo SAP x CSOD", layout="wide")
st.title("📊 Comparativo SAP x CSOD")

col1, col2 = st.columns(2)
with col1:
    csod_file = st.file_uploader("📤 Enviar arquivo da CSOD (.xlsx)", type=["xlsx"])
with col2:
    sap_file = st.file_uploader("📤 Enviar arquivo do SAP (.xlsx)", type=["xlsx"])

if csod_file and sap_file:
    st.success("Arquivos carregados com sucesso!")

    # CSOD começa na linha 9 (header=8)
    csod = pd.read_excel(csod_file, header=8)
    sap = pd.read_excel(sap_file)

    # Normalizar nomes de colunas
    csod.columns = csod.columns.astype(str).str.replace(r"\n|\r|\t", "", regex=True)
    csod.columns = csod.columns.str.replace(r"\s+", " ", regex=True).str.strip()
    sap.columns = sap.columns.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()

    # Validar colunas
    if "ID do Usuário" not in csod.columns or "Posição ID" not in csod.columns:
        st.error("Arquivo da CSOD deve conter as colunas: 'ID do Usuário' e 'Posição ID'")
        st.stop()
    if "NP" not in sap.columns or "Cargo - Cód." not in sap.columns:
        st.error("Arquivo do SAP deve conter as colunas: 'NP' e 'Cargo - Cód.'")
        st.stop()

    # Renomear colunas para padronizar
    csod.rename(columns={"ID do Usuário": "ID"}, inplace=True)
    sap.rename(columns={"NP": "ID"}, inplace=True)



    # Filtrar apenas IDs numéricos
    csod = csod[csod["ID"].astype(str).str.isnumeric()]
    sap = sap[sap["ID"].astype(str).str.isnumeric()]

    # Garantir IDs como string padronizados
    csod["ID"] = csod["ID"].astype(str).str.strip()
    sap["ID"] = sap["ID"].astype(str).str.strip()

    # Remover o ID "100008" (Carolina Ortega)
    csod = csod[csod["ID"] != "100008"]
    sap = sap[sap["ID"] != "100008"]


    # Remover ID 100008 e IDs iniciados por '89' ou '70'
    csod = csod[~csod["ID"].str.startswith(("100008", "89", "70"))]
    sap = sap[~sap["ID"].str.startswith(("100008", "89", "70"))]


    # 1. Usuários sem cargo na CSOD
    usuarios_sem_cargo = csod[csod["Posição ID"].isna() | (csod["Posição ID"].astype(str).str.strip() == "")]

    # 2. CSOD que não existe no SAP
    ids_csod = set(csod["ID"])
    ids_sap = set(sap["ID"])
    csod_nao_existe_no_sap = csod[csod["ID"].isin(ids_csod - ids_sap)]

    # 3. SAP que não existe na CSOD
    sap_nao_existe_no_csod = sap[sap["ID"].isin(ids_sap - ids_csod)]

    # 4. Cargos divergentes usando ID do cargo

    # Extrair IDs e normalizar para string de 8 dígitos
    csod["Cargo_ID"] = csod["Posição ID"].astype(str).str.split("-").str[0].str.strip().str.zfill(8)
    sap["Cargo_ID"] = sap["Cargo - Cód."].apply(lambda x: str(int(x)).zfill(8) if pd.notna(x) else "")

    csod_validos = csod[["ID", "Cargo_ID"]].dropna()
    sap_validos = sap[["ID", "Cargo_ID"]].dropna()

    comparativo = pd.merge(csod_validos, sap_validos, on="ID", how="inner", suffixes=("_CSOD", "_SAP"))

    cargos_divergentes = comparativo[
        comparativo["Cargo_ID_CSOD"] != comparativo["Cargo_ID_SAP"]
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
