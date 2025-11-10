import streamlit as st
from io import BytesIO
from typing import Dict
from memorial_core import (
    gerar_memorial_memorial,
    gerar_memorial_resumo,
    gerar_solicitacao_analise,
)

st.set_page_config(
    page_title="Gerador de Memorial - Sólido Design Urbano",
    layout="wide"
)

st.title("Gerador de Memorial – Sólido Design Urbano")

TIPOS = {
    "Memorial Condomínio": "condominio",
    "Memorial Loteamento": "loteamento",
    "Memorial Unificação": "unificacao",
    "Memorial Desmembramento": "desmembramento",
    "Memorial Unificação e Desmembramento": "unif_desm",
    "Memorial Resumo": "memorial_resumo",
    "Solicitação de Análise": "solicitacao_analise",
}

tipo_label = st.selectbox("Tipo de documento", list(TIPOS.keys()))
tipo = TIPOS[tipo_label]

def multi_file_uploader(label: str, types=None) -> Dict[str, bytes]:
    uploaded = st.file_uploader(label, type=types, accept_multiple_files=True)
    files: Dict[str, bytes] = {}
    if uploaded:
        for f in uploaded:
            files[f.name] = f.read()
    return files

st.markdown("### Dados básicos do empreendimento")

col1, col2 = st.columns(2)
with col1:
    nome = st.text_input("Nome do empreendimento")
    endereco = st.text_input("Endereço (com número ou s/nº)")
    bairro = st.text_input("Bairro")
    cidade = st.text_input("Cidade/UF (ex.: Porto Alegre/RS)")
    matricula = st.text_input("Matrícula(s) do imóvel")
with col2:
    area_total = st.text_input("Área total da gleba (m²)")
    perimetro = st.text_input("Perímetro (m)")
    num_lotes = st.text_input("Número de lotes")
    coord_fmt = st.selectbox(
        "Formato das coordenadas",
        options=[("UTM", "utm"), ("Graus decimais", "dec"), ("Graus-min-seg", "dms")],
        format_func=lambda x: x[0],
    )[1]

st.markdown("---")

extra = {}

# Opções específicas para loteamento/condomínio
if tipo in ("condominio", "loteamento"):
    st.markdown("#### Opções específicas")
    c1, c2 = st.columns(2)
    with c1:
        ane = st.selectbox("Área não edificante na testada?", ["Não", "Sim"])
    with c2:
        ane_largura = st.text_input("Largura área não edificante (m)", value="")
    extra["ane_enable"] = (ane == "Sim")
    extra["ane_largura"] = ane_largura or ""

    if tipo == "condominio":
        c3, c4 = st.columns(2)
        with c3:
            area_priv = st.text_input("Área total privativa (m²)")
        with c4:
            area_cond = st.text_input("Área total condominial (m²)")
        extra["area_priv"] = area_priv
        extra["area_cond"] = area_cond

# Memorial Resumo
if tipo == "memorial_resumo":
    st.markdown("#### Parâmetros específicos do Memorial Resumo")
    c1, c2 = st.columns(2)
    with c1:
        tipo_proj = st.selectbox(
            "Tipo de empreendimento",
            ["Condomínio", "Loteamento"],
        )
        usos = st.multiselect(
            "Usos do empreendimento",
            ["Residencial", "Comercial", "Industrial"],
            default=["Residencial"],
        )
        topografia = st.selectbox("Topografia", ["Acentuada", "Plana"])
    with c2:
        has_ai = st.checkbox("Possui área institucional?")
        has_restricao = st.checkbox("Possui área de restrição?")
    extra.update(
        tipo_proj_resumo="condominio" if tipo_proj == "Condomínio" else "loteamento",
        usos_multi=usos,
        topografia=topografia,
        has_ai=has_ai,
        has_restricao=has_restricao,
    )

# Solicitação de Análise
if tipo == "solicitacao_analise":
    st.markdown("#### Tipo de empreendimento para o ofício")
    tipo_proj_sol = st.selectbox(
        "Tipo",
        ["Condomínio fechado de lotes", "Loteamento de acesso controlado"],
    )
    extra["tipo_proj_resumo"] = (
        "condominio" if "Condomínio" in tipo_proj_sol else "loteamento"
    )

# UNIF / DESM
if tipo in ("unificacao", "desmembramento", "unif_desm"):
    st.markdown("#### Arquivos para UNIFICAÇÃO / DESMEMBRAMENTO")
    st.write(
        "- Para UNIFICAÇÃO: anexe o CivilReport HTML contendo o polígono da unificação.\n"
        "- Para DESMEMBRAMENTO: anexe os HTML/TXT das glebas gerados pelo Civil 3D."
    )

# Uploads
if tipo in ("condominio", "loteamento"):
    st.markdown("#### Arquivos de lotes / quadras (HTML/TXT do Civil 3D)")
    st.write("Anexe os arquivos exportados do Civil 3D com os lotes/quadras.")
    uploaded = multi_file_uploader("Arquivos HTML/TXT", ["html", "htm", "txt"])
elif tipo in ("unificacao", "desmembramento", "unif_desm"):
    uploaded = multi_file_uploader(
        "Arquivos HTML/TXT (CivilReport + glebas)", ["html", "htm", "txt"]
    )
else:
    uploaded = {}

st.markdown("---")

if st.button("Gerar documento"):
    if not nome or not cidade:
        st.error("Preencha pelo menos o Nome do empreendimento e Cidade/UF.")
    else:
        form = dict(
            tipo=tipo,
            nome=nome,
            endereco=endereco,
            bairro=bairro,
            cidade=cidade,
            matricula=matricula,
            area_total=area_total,
            perimetro=perimetro,
            num_lotes=num_lotes,
            coord_fmt=coord_fmt,
        )
        form.update(extra)

        if tipo in ("condominio", "loteamento", "unificacao", "desmembramento", "unif_desm"):
            docx_bytes, filename = gerar_memorial_memorial(form, uploaded)
        elif tipo == "memorial_resumo":
            docx_bytes, filename = gerar_memorial_resumo(form)
        elif tipo == "solicitacao_analise":
            docx_bytes, filename = gerar_solicitacao_analise(form)
        else:
            st.error("Tipo não suportado.")
            st.stop()

        st.success(f"Documento gerado: {filename}")
        st.download_button(
            "Baixar DOCX",
            data=docx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
