import streamlit as st
import pandas as pd
import io

# Configuração inicial da página web
st.set_page_config(page_title="Validador de Horas Extras", page_icon="⏱️", layout="centered")

st.title("⏱️ Validador de Horas Extras")
st.markdown("Faça o upload das planilhas abaixo para cruzar os dados do Líder com o RH.")

# 1. Componentes de Upload de Arquivos na Web
arquivo_lider = st.file_uploader("1. Envie a planilha do Líder (Ex: WK18.xlsx)", type=["xlsx"])
arquivo_rh = st.file_uploader("2. Envie o Relatório do RH", type=["xlsx"])

# Botão para iniciar a ação
if st.button("Cruzar Dados", type="primary"):
    
    if arquivo_lider is None or arquivo_rh is None:
        st.warning("⚠️ Por favor, envie as duas planilhas antes de continuar.")
    else:
        with st.spinner("Processando planilhas e cruzando dados..."):
            try:
                # 2. Leitura dos arquivos enviados pelo navegador
                df_wk18 = pd.read_excel(arquivo_lider, sheet_name='Quadro MO - AIRCON PCBA', skiprows=20)
                df_rh = pd.read_excel(arquivo_rh, sheet_name=0)

                # Limpeza das colunas
                df_wk18.columns = df_wk18.columns.astype(str).str.strip()
                df_rh.columns = df_rh.columns.astype(str).str.strip()

                # Preparação WK18
                wk18 = df_wk18[['Nome', 'Status']].dropna(subset=['Nome']).copy()
                wk18.rename(columns={'Nome': 'Nome_WK18', 'Status': 'Status_WK18'}, inplace=True)
                wk18['Nome_WK18'] = wk18['Nome_WK18'].astype(str)
                
                # Filtros de limpeza
                wk18 = wk18[~wk18['Nome_WK18'].str.isnumeric()]
                wk18 = wk18[wk18['Nome_WK18'].str.len() > 3]
                wk18 = wk18[~wk18['Nome_WK18'].str.strip().str.upper().eq('ATUAL')]
                
                wk18['Nome_Clean'] = wk18['Nome_WK18'].str.strip().str.upper()

                # Preparação RH
                rh = df_rh[['Nome.1', 'Status', 'Status HE']].dropna(subset=['Nome.1']).copy()
                rh.rename(columns={'Nome.1': 'Nome_RH', 'Status': 'Status_RH'}, inplace=True)
                rh['Nome_Clean'] = rh['Nome_RH'].str.strip().str.upper()
                rh = rh.drop_duplicates(subset=['Nome_Clean'])

                # Cruzamento
                merged = pd.merge(wk18, rh, on='Nome_Clean', how='left')
                merged['Status Final'] = merged['Status_RH'].fillna('Não está na planilha do RH')
                merged['Status HE Final'] = merged['Status HE'].fillna('-')

                final_df = merged[['Nome_WK18', 'Status Final', 'Status HE Final']].copy()
                final_df.rename(columns={
                    'Nome_WK18': 'Nome do Colaborador',
                    'Status Final': 'Status',
                    'Status HE Final': 'Status HE'
                }, inplace=True)

                # Regra de Cores
                def colorir_linhas(row):
                    if row['Status'] == 'Não está na planilha do RH':
                        return ['background-color: #FFC7CE'] * len(row) # Vermelho
                    else:
                        return ['background-color: #C6EFCE'] * len(row) # Verde

                styled_df = final_df.style.apply(colorir_linhas, axis=1)

                # 3. Preparando o arquivo Excel para Download na Memória (sem salvar no PC)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    styled_df.to_excel(writer, index=False, sheet_name="Resultado")
                
                # 4. Exibindo resultados na tela
                st.success(f"✅ Sucesso! {len(final_df)} colaboradores verificados.")
                
                # Mostra uma prévia interativa na própria página web
                st.subheader("Prévia dos Resultados")
                st.dataframe(final_df, use_container_width=True)

                # 5. Botão de Download do Excel Colorido
                st.download_button(
                    label="📥 Baixar Planilha com Cores",
                    data=buffer.getvalue(),
                    file_name="Resultado_Comparacao_HE.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Ocorreu um erro ao processar os arquivos. Verifique se a aba 'Quadro MO - AIRCON PCBA' existe.\n\nDetalhe do erro: {e}")