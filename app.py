import streamlit as st
from memorial_core import generate_docx, generate_excel

st.set_page_config(
    page_title="Gerador de Memorial - Sólido",
    layout="wide"
)

st.title("Gerador de Memorial (Streamlit)")
st.caption("Versão web equivalente ao notebook do Colab, com os mesmos tipos, campos e lógicas.")

# ===================== LOGOS =====================

with st.expander("Configuração de logos (opcional)"):
    header_logo_file = st.file_uploader("Logo de cabeçalho", type=["png", "jpg", "jpeg"], key="header_logo")
    footer_logo_file = st.file_uploader("Logo de rodapé", type=["png", "jpg", "jpeg"], key="footer_logo")
    watermark_logo_file = st.file_uploader("Marca d'água/canto", type=["png", "jpg", "jpeg"], key="watermark_logo")

header_logo = header_logo_file.read() if header_logo_file else None
footer_logo = footer_logo_file.read() if footer_logo_file else None
watermark_logo = watermark_logo_file.read() if watermark_logo_file else None

# ===================== TIPO =====================

TIPOS = [
    "Memorial Condomínio",
    "Memorial Loteamento",
    "Memorial Unificação",
    "Memorial Desmembramento",
    "Memorial Unificação e Desmembramento",
    "Memorial Resumo",
    "Solicitação de Análise",
]

tipo = st.selectbox("Tipo de DOCX:", TIPOS)

st.markdown("---")

# ===================== CAMPOS BÁSICOS =====================

col1, col2 = st.columns(2)

with col1:
    nome_emp = st.text_input("Empreendimento:", placeholder="Ex.: Golden View")
    endereco_emp = st.text_input("Endereço:", placeholder="Av. Principal, 123")
    bairro_emp = st.text_input("Bairro:", placeholder="Centro")
with col2:
    cidade_emp = st.text_input("Cidade (Cidade/UF):", placeholder="Portão/RS")
    area_total_emp = st.text_input("Área total da gleba (m²):", placeholder="123456,78")
    matricula_emp = st.text_input("Matrícula(s):", placeholder="17.051, 17.052, 17.053")

# ===================== CAMPOS ESPECÍFICOS POR TIPO =====================

perimetro_emp = ""
num_lotes_emp = 0
coord_fmt = "utm"
ane_drop = "Não"
ane_largura = ""
area_tot_priv_emp = ""
area_tot_cond_emp = ""
tipo_proj_resumo = "loteamento"
usos_multi = []
topografia = "Acentuada"
has_ai = False
has_restricao = False

if tipo in ("Memorial Condomínio", "Memorial Loteamento"):
    c1, c2 = st.columns(2)
    with c1:
        num_lotes_emp = st.number_input("Nº de lotes:", min_value=0, step=1)
        perimetro_emp = st.text_input("Perímetro da gleba (m):", placeholder="3456,78")
    with c2:
        coord_fmt = st.selectbox("Formato das coordenadas:",
                                 ["utm", "dec", "dms"],
                                 format_func=lambda v: {
                                     "utm": "UTM (SIRGAS 2000)",
                                     "dec": "Graus decimais",
                                     "dms": "Graus, minutos e segundos"
                                 }[v])
        ane_drop = st.selectbox("Possui área não edificante (faixa)?", ["Não", "Sim"])
        if ane_drop == "Sim":
            ane_largura = st.text_input("Largura da faixa não edificante (m):", "3,00")

    if tipo == "Memorial Condomínio":
        c3, c4 = st.columns(2)
        with c3:
            area_tot_priv_emp = st.text_input("Área total privativa (m²):", "")
        with c4:
            area_tot_cond_emp = st.text_input("Área total condominial (m²):", "")

elif tipo in ("Memorial Unificação", "Memorial Desmembramento", "Memorial Unificação e Desmembramento"):
    c1, c2 = st.columns(2)
    with c1:
        perimetro_emp = st.text_input("Perímetro (opcional):", "")
    with c2:
        coord_fmt = st.selectbox("Formato das coordenadas:",
                                 ["utm", "dec", "dms"],
                                 format_func=lambda v: {
                                     "utm": "UTM (SIRGAS 2000)",
                                     "dec": "Graus decimais",
                                     "dms": "Graus, minutos e segundos"
                                 }[v])

elif tipo == "Memorial Resumo":
    c1, c2 = st.columns(2)
    with c1:
        tipo_proj_resumo = st.selectbox(
            "Tipo de empreendimento:",
            ["loteamento", "condominio"],
            format_func=lambda v: "Loteamento" if v == "loteamento" else "Condomínio"
        )
        usos_multi = st.multiselect(
            "Usos do empreendimento:",
            ["Residencial", "Comercial", "Serviços", "Industrial"],
            default=["Residencial"]
        )
    with c2:
        topografia = st.selectbox("Topografia da gleba:", ["Acentuada", "Plana"])
        has_ai = st.checkbox("Possui área institucional?")
        has_restricao = st.checkbox("Possui área de restrição?")
        num_lotes_emp = st.number_input("Nº de lotes (informativo):", min_value=0, step=1)

elif tipo == "Solicitação de Análise":
    tipo_proj_resumo = st.selectbox(
        "Tipo de empreendimento:",
        ["loteamento", "condominio"],
        format_func=lambda v: "Loteamento" if v == "loteamento" else "Condomínio"
    )

# ===================== UPLOAD DE ARQUIVOS =====================

uploaded_files_dict = {}

if tipo in (
    "Memorial Condomínio",
    "Memorial Loteamento",
    "Memorial Unificação",
    "Memorial Desmembramento",
    "Memorial Unificação e Desmembramento",
):
    st.markdown("### Arquivos de apoio (HTML/TXT)")
    st.write(
        "- HTML/TXT de quadras/lotes (Parcel Report)\n"
        "- CivilReport para áreas gerais (viário, APP, verde, etc.)\n"
        "Use exatamente como no Colab."
    )
    up_files = st.file_uploader(
        "Anexar arquivos",
        type=["html", "htm", "txt"],
        accept_multiple_files=True
    )
    for f in up_files or []:
        uploaded_files_dict[f.name] = f.read()

# ===================== MONTAGEM DO FORM =====================

form = {
    "nome_emp": nome_emp,
    "endereco_emp": endereco_emp,
    "bairro_emp": bairro_emp,
    "cidade_emp": cidade_emp,
    "area_total_emp": area_total_emp,
    "perimetro_emp": perimetro_emp,
    "matricula_emp": matricula_emp,
    "num_lotes_emp": num_lotes_emp,
    "coord_fmt": coord_fmt,
    "ane_drop": ane_drop,
    "ane_largura": ane_largura,
    "area_tot_priv_emp": area_tot_priv_emp,
    "area_tot_cond_emp": area_tot_cond_emp,
    "tipo_proj_resumo": tipo_proj_resumo,
    "usos_multi": usos_multi,
    "topografia": topografia,
    "has_ai": has_ai,
    "has_restricao": has_restricao,
}

if "quadro_frac_ideal" not in st.session_state:
    st.session_state.quadro_frac_ideal = None

st.markdown("---")
b1, b2 = st.columns(2)

# ===================== GERAR DOCX =====================

with b1:
    if st.button("Gerar DOCX"):
        try:
            doc_bytes, filename, meta = generate_docx(
                tipo,
                form,
                uploaded_files_dict,
                header_logo=header_logo,
                footer_logo=footer_logo,
                watermark_logo=watermark_logo
            )

            # guarda quadro para Excel quando for condomínio
            st.session_state.quadro_frac_ideal = meta if meta else None

            st.success(f"DOCX gerado com sucesso: {filename}")
            st.download_button(
                "Baixar DOCX",
                data=doc_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Erro ao gerar DOCX: {e}")

# ===================== GERAR EXCEL =====================

with b2:
    show_excel_btn = tipo in (
        "Memorial Condomínio",
        "Memorial Unificação",
        "Memorial Desmembramento",
        "Memorial Unificação e Desmembramento",
    )
    if show_excel_btn:
        if st.button("Gerar Excel"):
            try:
                excel_bytes, xlsx_name = generate_excel(
                    tipo,
                    form,
                    uploaded_files_dict,
                    quadro_frac_ideal=st.session_state.quadro_frac_ideal
                )
                st.success(f"Excel gerado: {xlsx_name}")
                st.download_button(
                    "Baixar Excel",
                    data=excel_bytes,
                    file_name=xlsx_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Erro ao gerar Excel: {e}")
    else:
        st.write(" ")
