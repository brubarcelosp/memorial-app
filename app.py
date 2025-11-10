import streamlit as st
from io import BytesIO
from docx import Document

# Configuração básica da página
st.set_page_config(
    page_title="Gerador de Memorial - Sólido",
    layout="wide"
)

st.title("Gerador de Memorial – Sólido Design Urbano")

# Campos principais
tipo = st.selectbox(
    "Tipo de DOCX",
    [
        "Memorial Loteamento",
        "Memorial Condomínio",
        "Memorial Unificação",
        "Memorial Desmembramento",
        "Memorial Unificação e Desmembramento",
        "Memorial Resumo",
        "Solicitação de Análise"
    ]
)

col1, col2 = st.columns(2)
with col1:
    nome = st.text_input("Nome do empreendimento")
    cidade = st.text_input("Cidade/UF")
with col2:
    area_total = st.text_input("Área total (m²)")
    num_lotes = st.text_input("Número de lotes (se aplicar)")

arquivo_html = st.file_uploader(
    "Anexar relatório HTML/TXT do Civil 3D (opcional)",
    type=["html", "htm", "txt"]
)

# Botão para gerar memorial
if st.button("Gerar Memorial DOCX"):
    if not nome or not cidade or not area_total:
        st.error("Preencha pelo menos Nome, Cidade e Área total.")
    else:
        # Por enquanto: memorial simples (mock)
        # Depois vamos plugar aqui a lógica completa do teu código.
        doc = Document()
        doc.add_paragraph(f"Memorial: {tipo}")
        doc.add_paragraph(f"Empreendimento: {nome}")
        doc.add_paragraph(f"Cidade/UF: {cidade}")
        doc.add_paragraph(f"Área total: {area_total} m²")
        if num_lotes:
            doc.add_paragraph(f"Número de lotes: {num_lotes}")
        if arquivo_html:
            doc.add_paragraph(f"Arquivo base recebido: {arquivo_html.name}")

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.success("Memorial gerado com sucesso.")
        st.download_button(
            "Baixar memorial.docx",
            data=buffer,
            file_name="memorial.docx",
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument.wordprocessingml."
                "document"
            ),
        )
