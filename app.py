import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import textwrap
import io
import zipfile

# --- CONFIGURAÇÕES DA PÁGINA ---
st.set_page_config(
    page_title="Gerador de Débitos",
    page_icon="✍️",
    layout="centered"
)

# --- FUNÇÕES CORE (LÓGICA ADAPTADA DO COLAB) ---

def format_brl(value):
    """Formata um número para o padrão brasileiro (1.234,56)"""
    try:
        a = '{:,.2f}'.format(float(value))
        b = a.replace(',', 'v')
        c = b.replace('.', ',')
        d = c.replace('v', '.')
        return d
    except (ValueError, TypeError):
        return "0,00"

def gerar_documento_word(context):
    """Gera um documento Word a partir de um dicionário de contexto."""
    try:
        doc = DocxTemplate("DEBIT - template.docx")
        doc.render(context)
        
        # Salva o documento em um buffer de memória
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
    except Exception as e:
        st.error(f"Erro ao gerar o documento: {e}")
        st.error("Verifique se o arquivo 'DEBIT - template.docx' está na mesma pasta do script.")
        return None

# --- INTERFACE DA APLICAÇÃO ---

st.title("✍️ Gerador de Formulários de Débito")
st.markdown("Aplicação para preenchimento automático de débitos, individualmente ou em massa via Excel.")

# --- ABAS PARA CADA FUNCIONALIDADE ---
tab1, tab2 = st.tabs(["📄 Formulário Individual", "📊 Múltiplos Débitos (Excel)"])

# --- ABA 1: FORMULÁRIO INDIVIDUAL ---
with tab1:
    st.header("Preenchimento de Débito Individual")

    with st.form("form_individual"):
        # Usando colunas para melhor organização
        col1, col2 = st.columns(2)

        with col1:
            w_escritorio = st.selectbox("Escritório", ["ASBZ SP", "ZUCCA BSB", "CONSULTING"], key="ind_esc")
            w_solicitante = st.text_input("Solicitante (Sigla)", key="ind_sol")
            w_cliente = st.number_input("Cliente (Número)", step=1, key="ind_cli")
            w_tipo_despesa = st.selectbox("Tipo de Despesa", ["MOTOCA", "CARTÓRIO", "CORREIOS", "OUTROS"], key="ind_tipo_desp")
            w_reembolsavel = st.radio("Reembolsável?", ["SIM", "NÃO"], horizontal=True, key="ind_reemb")
        
        with col2:
            w_centro_custo = st.text_input("Centro de Custo", key="ind_cc")
            w_os_caso = st.number_input("OS/Caso (Número)", step=1, key="ind_os")
            w_total_rs = st.number_input("Total R$", format="%.2f", key="ind_total")
            w_data_despesa = st.date_input("Data da Despesa", key="ind_data")
            w_adiantamento = st.radio("Tem adiantamento?", ["SIM", "NÃO"], horizontal=True, key="ind_adiant")

        w_observacao = st.text_area("Observação (Opcional)", height=150, key="ind_obs")
        
        submitted = st.form_submit_button("Gerar DEBIT")

        if submitted:
            # Coleta e processamento dos dados do formulário
            context = {
                's': w_solicitante.upper(), 'cc': w_centro_custo,
                'cl': w_cliente, 'osc': w_os_caso,
                'data': w_data_despesa.strftime('%d/%m/%Y') if w_data_despesa else '',
                'total': format_brl(w_total_rs)
            }
            
            # Lógica para dividir o texto da observação
            texto_completo_obs = w_observacao
            limite_de_caracteres = 97
            linhas_quebradas = textwrap.wrap(texto_completo_obs, width=limite_de_caracteres)
            placeholders_obs = ['obs', 'obs2', 'obs3', 'obs4', 'obs5']
            for i, placeholder in enumerate(placeholders_obs):
                context[placeholder] = linhas_quebradas[i] if i < len(linhas_quebradas) else ""

            # Lógica para os "checkboxes"
            context.update({k: 'X' if w_escritorio == v else '' for k, v in {'e1': 'ASBZ SP', 'e2': 'ZUCCA BSB', 'e3': 'CONSULTING'}.items()})
            context.update({k: 'X' if w_tipo_despesa == v else '' for k, v in {'m': 'MOTOCA', 'c': 'CARTÓRIO', 'co': 'CORREIOS', 'o': 'OUTROS'}.items()})
            context['si'] = 'X' if w_reembolsavel == 'SIM' else ''
            context['na'] = 'X' if w_reembolsavel == 'NÃO' else ''
            context['as'] = 'X' if w_adiantamento == 'SIM' else ''
            context['an'] = 'X' if w_adiantamento == 'NÃO' else ''
            
            # Geração e download do documento
            doc_buffer = gerar_documento_word(context)
            if doc_buffer:
                st.success("Documento gerado com sucesso!")
                filename = f"DEBIT_Cliente_{context['cl']}_Caso_{context['osc']}_{datetime.now().strftime('%Y-%m-%d')}.docx"
                st.download_button(
                    label="📥 Baixar Documento",
                    data=doc_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

# --- ABA 2: MÚLTIPLOS DÉBITOS (EXCEL) ---
with tab2:
    st.header("Geração de Múltiplos Débitos via Excel")
    
    st.info("""
    **Instruções:**
    1. Faça o upload de uma planilha Excel (`.xlsx`).
    2. A planilha **DEVE** conter as seguintes colunas (com o nome exato):
       - `Escritorio`, `Solicitante`, `CentroCusto`, `Cliente`, `OS_Caso`, `TipoDespesa`, `Total`, `DataDespesa`, `Reembolsavel`, `Adiantamento`, `Observacao`
    3. Os valores para colunas de múltipla escolha devem ser exatos (ex: "ASBZ SP", "MOTOCA", "SIM", "NÃO").
    """)

    uploaded_file = st.file_uploader("Escolha sua planilha Excel", type="xlsx")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.dataframe(df)

            if st.button("Gerar Todos os Débitos da Planilha"):
                with st.spinner("Gerando documentos... Isso pode levar um momento."):
                    # Cria um arquivo ZIP em memória
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for index, row in df.iterrows():
                            # Constrói o contexto para cada linha da planilha
                            context = {
                                's': str(row['Solicitante']).upper(), 'cc': row['CentroCusto'],
                                'cl': row['Cliente'], 'osc': row['OS_Caso'],
                                'data': pd.to_datetime(row['DataDespesa']).strftime('%d/%m/%Y') if pd.notna(row['DataDespesa']) else '',
                                'total': format_brl(row['Total'])
                            }

                            texto_completo_obs = str(row['Observacao']) if pd.notna(row['Observacao']) else ""
                            limite_de_caracteres = 97
                            linhas_quebradas = textwrap.wrap(texto_completo_obs, width=limite_de_caracteres)
                            placeholders_obs = ['obs', 'obs2', 'obs3', 'obs4', 'obs5']
                            for i, placeholder in enumerate(placeholders_obs):
                                context[placeholder] = linhas_quebradas[i] if i < len(linhas_quebradas) else ""
                            
                            # Lógica dos checkboxes
                            context.update({k: 'X' if row['Escritorio'] == v else '' for k, v in {'e1': 'ASBZ SP', 'e2': 'ZUCCA BSB', 'e3': 'CONSULTING'}.items()})
                            context.update({k: 'X' if row['TipoDespesa'] == v else '' for k, v in {'m': 'MOTOCA', 'c': 'CARTÓRIO', 'co': 'CORREIOS', 'o': 'OUTROS'}.items()})
                            context['si'] = 'X' if row['Reembolsavel'] == 'SIM' else ''
                            context['na'] = 'X' if row['Reembolsavel'] == 'NÃO' else ''
                            context['as'] = 'X' if row['Adiantamento'] == 'SIM' else ''
                            context['an'] = 'X' if row['Adiantamento'] == 'NÃO' else ''
                            
                            # Gera o documento em memória
                            doc_buffer = gerar_documento_word(context)
                            if doc_buffer:
                                filename = f"DEBIT_Cliente_{context['cl']}_Caso_{context['osc']}_{index+1}.docx"
                                zipf.writestr(filename, doc_buffer.getvalue())

                    zip_buffer.seek(0)
                    st.success("Todos os documentos foram gerados e compactados!")
                    st.download_button(
                        label="📥 Baixar todos os documentos (.zip)",
                        data=zip_buffer,
                        file_name=f"DEBITS_Gerados_{datetime.now().strftime('%Y-%m-%d')}.zip",
                        mime="application/zip"
                    )

        except Exception as e:
            st.error(f"Ocorreu um erro ao processar o arquivo Excel: {e}")
            st.error("Verifique se o nome das colunas está correto e se o formato dos dados é válido.")