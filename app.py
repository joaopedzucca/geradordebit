import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import textwrap
import io
import zipfile

# --- CONFIGURAÇÕES DA PÁGINA ---
st.set_page_config(
    page_title="Gerador de DEBITS",
    page_icon="✍️",
    layout="centered"
)

# --- FUNÇÕES CORE ---

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

def create_excel_template():
    """Cria um DataFrame modelo e o converte para um arquivo Excel em memória."""
    data = {
        'Escritorio': ['ASBZ SP'],
        'Solicitante': ['JDOE'],
        'CentroCusto': ['TI (Opcional)'],
        'Cliente': ['007 (Opcional)'],
        'OS_Caso': ['001 (Opcional)'],
        'TipoDespesa': ['CORREIOS'],
        'Total': [150.75],
        'DataDespesa': [datetime.now().date()],
        'Reembolsavel': ['SIM'],
        'Adiantamento': ['NÃO'],
        'Observacao': ['Exemplo de observação (Opcional).']
    }
    df = pd.DataFrame(data)
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)
    return buffer

def gerar_documento_word(context):
    """Gera um documento Word a partir de um dicionário de contexto."""
    try:
        doc = DocxTemplate("DEBIT - template.docx")
        doc.render(context)
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
    except Exception as e:
        st.error(f"Erro ao gerar o documento: {e}")
        st.error("Verifique se o arquivo 'DEBIT - template.docx' está na mesma pasta do script.")
        return None

# --- INTERFACE DA APLICAÇÃO ---

st.title("✍️ Gerador de Formulários de DEBIT")
st.markdown("Aplicação para preenchimento automático de DEBITS, individualmente ou em massa via Excel.")

tab1, tab2 = st.tabs(["📄 Formulário Individual", "📊 Múltiplos DEBITS (Excel)"])

# --- ABA 1: FORMULÁRIO INDIVIDUAL ---
with tab1:
    st.header("Preenchimento de DEBIT Individual")

    if 'doc_buffer' not in st.session_state:
        st.session_state.doc_buffer = None

    with st.form("form_individual"):
        col1, col2 = st.columns(2)

        with col1:
            w_escritorio = st.selectbox("Escritório*", ["ASBZ SP", "ZUCCA BSB", "CONSULTING"], key="ind_esc")
            w_solicitante = st.text_input("Solicitante (Sigla)*", key="ind_sol")
            # MUDANÇA AQUI: Removido o ícone de ajuda (parâmetro help)
            w_cliente = st.text_input("Cliente", key="ind_cli")
            w_tipo_despesa = st.selectbox("Tipo de Despesa*", ["MOTOCA", "CARTÓRIO", "CORREIOS", "OUTROS"], key="ind_tipo_desp")
            w_reembolsavel = st.radio("Reembolsável?*", ["SIM", "NÃO"], horizontal=True, key="ind_reemb")
        
        with col2:
            w_centro_custo = st.text_input("Centro de Custo", key="ind_cc")
            # MUDANÇA AQUI: Removido o ícone de ajuda (parâmetro help)
            w_os_caso = st.text_input("OS/Caso", key="ind_os")
            w_total_rs = st.number_input("Total R$*", format="%.2f", key="ind_total")
            w_data_despesa = st.date_input("Data da Despesa*", key="ind_data")
            # MUDANÇA AQUI: Texto do rótulo alterado
            w_adiantamento = st.radio("Tem adiantamento do cliente*", ["SIM", "NÃO"], horizontal=True, key="ind_adiant")

        w_observacao = st.text_area("Observação (Opcional)", height=150, key="ind_obs")
        
        submitted = st.form_submit_button("Gerar DEBIT")

        if submitted:
            context = {
                's': w_solicitante.upper(), 'cc': w_centro_custo,
                'cl': w_cliente, 'osc': w_os_caso,
                'data': w_data_despesa.strftime('%d/%m/%Y') if w_data_despesa else '',
                'total': format_brl(w_total_rs)
            }
            
            texto_completo_obs = w_observacao
            limite_de_caracteres = 97
            linhas_quebradas = textwrap.wrap(texto_completo_obs, width=limite_de_caracteres)
            placeholders_obs = ['obs', 'obs2', 'obs3', 'obs4', 'obs5']
            for i, placeholder in enumerate(placeholders_obs):
                context[placeholder] = linhas_quebradas[i] if i < len(linhas_quebradas) else ""

            context.update({k: 'X' if w_escritorio == v else '' for k, v in {'e1': 'ASBZ SP', 'e2': 'ZUCCA BSB', 'e3': 'CONSULTING'}.items()})
            context.update({k: 'X' if w_tipo_despesa == v else '' for k, v in {'m': 'MOTOCA', 'c': 'CARTÓRIO', 'co': 'CORREIOS', 'o': 'OUTROS'}.items()})
            context['si'] = 'X' if w_reembolsavel == 'SIM' else ''
            context['na'] = 'X' if w_reembolsavel == 'NÃO' else ''
            context['as'] = 'X' if w_adiantamento == 'SIM' else ''
            context['an'] = 'X' if w_adiantamento == 'NÃO' else ''
            
            st.session_state.doc_buffer = gerar_documento_word(context)
            st.session_state.filename_context = {'cl': w_cliente, 'osc': w_os_caso}
            if st.session_state.doc_buffer:
                st.success("DEBIT gerado com sucesso! Clique abaixo para baixar.")
            
    if st.session_state.doc_buffer:
        filename_context = st.session_state.get('filename_context', {'cl': 'N_A', 'osc': 'N_A'})
        filename = f"DEBIT_Cliente_{filename_context['cl']}_Caso_{filename_context['osc']}_{datetime.now().strftime('%Y-%m-%d')}.docx"
        st.download_button(
            label="📥 Baixar DEBIT",
            data=st.session_state.doc_buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# --- ABA 2: MÚLTIPLOS DEBITS (EXCEL) ---
with tab2:
    st.header("Geração de Múltiplos DEBITS via Excel")
    
    st.subheader("1. Baixe o Modelo")
    excel_template_buffer = create_excel_template()
    st.download_button(
        label="📥 Baixar planilha modelo (.xlsx)",
        data=excel_template_buffer,
        file_name="modelo_debits.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("2. Preencha e Faça o Upload")
    st.info("""
    **Instruções:**
    - Use o modelo baixado para garantir que os nomes das colunas estejam corretos.
    - **Importante:** Para as colunas `Cliente`, `OS_Caso` e `CentroCusto`, formate as células como **Texto** no Excel para permitir zeros à esquerda.
    """)

    uploaded_file = st.file_uploader("Escolha sua planilha Excel preenchida", type="xlsx")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, dtype={'Cliente': str, 'OS_Caso': str, 'CentroCusto': str})
            st.dataframe(df)

            if st.button("Gerar Todos os DEBITS da Planilha"):
                with st.spinner("Gerando documentos... Isso pode levar um momento."):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for index, row in df.iterrows():
                            context = {
                                's': str(row['Solicitante']).upper(), 'cc': str(row['CentroCusto']) if pd.notna(row['CentroCusto']) else '',
                                'cl': str(row['Cliente']) if pd.notna(row['Cliente']) else '', 
                                'osc': str(row['OS_Caso']) if pd.notna(row['OS_Caso']) else '',
                                'data': pd.to_datetime(row['DataDespesa']).strftime('%d/%m/%Y') if pd.notna(row['DataDespesa']) else '',
                                'total': format_brl(row['Total'])
                            }

                            texto_completo_obs = str(row['Observacao']) if pd.notna(row['Observacao']) else ""
                            limite_de_caracteres = 97
                            linhas_quebradas = textwrap.wrap(texto_completo_obs, width=limite_de_caracteres)
                            placeholders_obs = ['obs', 'obs2', 'obs3', 'obs4', 'obs5']
                            for i, placeholder in enumerate(placeholders_obs):
                                context[placeholder] = linhas_quebradas[i] if i < len(linhas_quebradas) else ""
                            
                            context.update({k: 'X' if row['Escritorio'] == v else '' for k, v in {'e1': 'ASBZ SP', 'e2': 'ZUCCA BSB', 'e3': 'CONSULTING'}.items()})
                            context.update({k: 'X' if row['TipoDespesa'] == v else '' for k, v in {'m': 'MOTOCA', 'c': 'CARTÓRIO', 'co': 'CORREIOS', 'o': 'OUTROS'}.items()})
                            context['si'] = 'X' if row['Reembolsavel'] == 'SIM' else ''
                            context['na'] = 'X' if row['Reembolsavel'] == 'NÃO' else ''
                            context['as'] = 'X' if row['Adiantamento'] == 'SIM' else ''
                            context['an'] = 'X' if row['Adiantamento'] == 'NÃO' else ''
                            
                            doc_buffer = gerar_documento_word(context)
                            if doc_buffer:
                                filename = f"DEBIT_Cliente_{context['cl']}_Caso_{context['osc']}_{index+1}.docx"
                                zipf.writestr(filename, doc_buffer.getvalue())

                    zip_buffer.seek(0)
                    st.success("Todos os documentos foram gerados e compactados!")
                    st.download_button(
                        label="📥 Baixar todos os DEBITS (.zip)",
                        data=zip_buffer,
                        file_name=f"DEBITS_Gerados_{datetime.now().strftime('%Y-%m-%d')}.zip",
                        mime="application/zip"
                    )
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar o arquivo Excel: {e}")
            st.error("Verifique se o nome das colunas está correto e se o formato dos dados é válido.")
