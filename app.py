import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="Atualização SIGET - LT 500kV Ceará Mirim II", 
    layout="wide",
    page_icon="⚡"
)

st.title("⚡ Painel de Atualização do SIGET")
st.subheader("LT 500kV Ceará Mirim II - João Pessoa II (SIGET - ANEEL)")

# 1. Upload do Arquivo
uploaded_file = st.sidebar.file_uploader("📂 Carregar planilha SIGET (.xlsx)", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    
    # Filtrar apenas abas SIGET
    siget_sheets = [sheet for sheet in xls.sheet_names if "SIGET" in sheet]
    if not siget_sheets:
        siget_sheets = xls.sheet_names
        
    selected_sheet = st.sidebar.selectbox(
        "📅 Selecione a Aba / Mês de Referência", 
        siget_sheets, 
        index=len(siget_sheets)-1
    )

    # Carrega os dados brutos da aba selecionada
    df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)

    st.info(f"Editando informações da aba: **{selected_sheet}**")

    # Extrai o trecho de dados referente às atividades (linha 6 em diante na estrutura SIGET)
    # Colunas: EAP (5), Descrição (6), Real Avanço % (10), Data Prevista Início (11), Data Prevista Conclusão (12), Data Efetiva Início (13), Data Efetiva Conclusão (14)
    activities_df = df_raw.iloc[6:, [5, 6, 10, 11, 12, 13, 14]].copy()
    activities_df.columns = [
        "EAP", 
        "Descrição da Atividade", 
        "Real Avanço (%)", 
        "Data Prevista Início", 
        "Data Prevista Conclusão", 
        "Data Efetiva Início", 
        "Data Efetiva Conclusão"
    ]

    # Tratamento de dados para exibição
    activities_df["Real Avanço (%)"] = pd.to_numeric(activities_df["Real Avanço (%)"], errors='coerce').fillna(0)
    
    date_cols = ["Data Prevista Início", "Data Prevista Conclusão", "Data Efetiva Início", "Data Efetiva Conclusão"]
    for col in date_cols:
        activities_df[col] = pd.to_datetime(activities_df[col], errors='coerce').dt.date

    # 2. Tabela Interativa para Atualização das Colunas Solicitadas
    st.write("### 📝 Atualização de Real Avanço e Datas Previstas/Efetivas")
    
    edited_df = st.data_editor(
        activities_df,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        column_config={
            "EAP": st.column_config.Column(disabled=True, width="medium"),
            "Descrição da Atividade": st.column_config.Column(disabled=True, width="large"),
            "Real Avanço (%)": st.column_config.NumberColumn(
                label="Real Avanço (%)",
                min_value=0, 
                max_value=100, 
                step=1, 
                format="%d%%"
            ),
            "Data Prevista Início": st.column_config.DateColumn(label="Prev. Início", format="DD/MM/YYYY"),
            "Data Prevista Conclusão": st.column_config.DateColumn(label="Prev. Conclusão", format="DD/MM/YYYY"),
            "Data Efetiva Início": st.column_config.DateColumn(label="Efet. Início", format="DD/MM/YYYY"),
            "Data Efetiva Conclusão": st.column_config.DateColumn(label="Efet. Conclusão", format="DD/MM/YYYY"),
        }
    )

    # 3. Processamento e Download
    st.markdown("---")
    if st.button("💾 Salvar Alterações na Planilha", type="primary"):
        # Atualiza os dados no DataFrame original mantendo os cabeçalhos do SIGET intactos
        df_updated = df_raw.copy()
        
        # Converte datas de volta para datetime para salvar no Excel
        updated_values = edited_df[[
            "Real Avanço (%)", 
            "Data Prevista Início", 
            "Data Prevista Conclusão", 
            "Data Efetiva Início", 
            "Data Efetiva Conclusão"
        ]].copy()
        
        for col in ["Data Prevista Início", "Data Prevista Conclusão", "Data Efetiva Início", "Data Efetiva Conclusão"]:
            updated_values[col] = pd.to_datetime(updated_values[col])

        df_updated.iloc[6:, [10, 11, 12, 13, 14]] = updated_values.values

        # Gera o arquivo em memória mantendo todas as abas originais
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Copia todas as abas existentes
            for sheet in xls.sheet_names:
                if sheet == selected_sheet:
                    df_updated.to_excel(writer, sheet_name=sheet, index=False, header=False)
                else:
                    df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
                    df_sheet.to_excel(writer, sheet_name=sheet, index=False, header=False)

        output.seek(0)
        
        st.success(f"✅ Aba **{selected_sheet}** atualizada com sucesso!")
        
        st.download_button(
            label="📥 Baixar Planilha Atualizada",
            data=output,
            file_name=f"SIGET_Atualizado_{selected_sheet}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("👈 Por favor, faça o upload da planilha SIGET no menu lateral para começar.")
