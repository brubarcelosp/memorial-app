import streamlit as st
from memorial_core import preparar_doc
from io import BytesIO
from docx import Document

st.set_page_config(page_title="Gerador de Memorial - Sólido", layout="wide")

st.title("Gerador de Memorial – Sólido Design Urbano")

tipo = st.selectbox("Tipo de DOCX", [
    "Memorial Loteamento", "Memorial Condomínio",
    "Memorial Unificação", "Memorial Desmembramento",
    "Memorial Unificação e Desmembramento", "Memorial Resumo",
    "Solicitação de Análise"
])

nome = st.text_input("Nome do empreendimento")
cidade = st.text_input("Cidade/UF")
area_total = st.text_input("Área total (m²)")

arquivo_html = st.file_uploader("Anexar HTML/TXT (Civil 3D)", type=["html", "htm", "txt"])

if st.button("Gerar Memorial DOCX"):
    if not nome or not cidade or not area_total:
        st.error("Preencha os campos básicos.")
    else:
        doc = preparar_doc()
        doc.add_paragraph(f"Memorial: {tipo}")
        doc.add_paragraph(f"Empreendimento: {nome}")
        doc.add_paragraph(f"Cidade: {cidade}")
        doc.add_paragraph(f"Área total: {area_total} m²")

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.success("Memorial gerado com sucesso.")
        st.download_button(
            "Baixar Memorial DOCX",
            data=buffer,
            file_name="memorial.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
