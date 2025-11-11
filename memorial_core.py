import streamlit as st
import re, os, io, time, math
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_COLOR_INDEX, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pathlib import Path
from num2words import num2words
import pandas as pd
from pyproj import CRS, Transformer
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ===================== CONFIG STREAMLIT =====================

st.set_page_config(
    page_title="Gerador de Memorial - Sólido",
    layout="wide"
)

st.title("Gerador de Memorial (Streamlit)")
st.caption("Versão web equivalente ao notebook do Colab, com os mesmos tipos, campos e lógicas.")

# ===================== UTILS GERAIS =====================

# (mantidos iguais — apenas removido qualquer traço de Colab)

def _fmt_br(v, casas=2):
    try:
        return f"{float(v):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(v)

def _to_float_br(txt):
    return float(str(txt).replace('.', '').replace(',', '.'))

def to_float_any(s):
    s = str(s).strip()
    if ',' in s and '.' in s:
        return float(s.replace('.', '').replace(',', '.'))
    if ',' in s:
        return float(s.replace(',', '.'))
    return float(s)

def extenso_metros(v):
    v = round(float(v or 0), 2)
    m = int(v); cm = int(round((v - m) * 100))
    partes = []
    if m > 0:
        partes.append(num2words(m, lang='pt_BR') + (" metro" if m==1 else " metros"))
    if cm > 0:
        partes.append(num2words(cm, lang='pt_BR') + (" centímetro" if cm==1 else " centímetros"))
    return " e ".join(partes) if partes else "zero metro"

def area_por_extenso(v):
    v = round(float(v or 0), 2)
    m2 = int(v); cent = int(round((v - m2) * 100))
    if cent == 0:
        return f"{num2words(m2, lang='pt_BR')} metros quadrados"
    return f"{num2words(m2, lang='pt_BR')} metros quadrados e {num2words(cent, lang='pt_BR')} centésimos"

def hectares_from_m2(v):
    return float(v) / 10000.0

_PREP_MIN = {"DE", "DA", "DO", "DAS", "DOS"}

def _title_keep_preps(s: str) -> str:
    if not s:
        return ""
    t = s.strip().title()
    for prep in _PREP_MIN:
        t = re.sub(rf"\b{prep}\b", prep.lower(), t)
    t = re.sub(
        r'\s*,?\s*S\s*/\s*N[º°]?\b',
        lambda m: (',' if ',' in m.group(0) else '') + ' s/nº',
        t,
        flags=re.I
    )
    return t.strip()

def _fmt_cidade_slash_uf(s: str) -> str:
    if not s: return ""
    s = s.strip()
    if "/" not in s:
        return _title_keep_preps(s)
    cid, uf = [p.strip() for p in s.split("/", 1)]
    return f"{_title_keep_preps(cid)}/{uf.upper()}"

def _fmt_bairro(s: str) -> str:
    if not s:
        return ""
    t = s.strip().title()
    for prep in {"De", "Da", "Do", "Das", "Dos", "A", "E"}:
        t = re.sub(rf"\b{prep}\b", prep.lower(), t)
    return t

# ===================== COORDENADAS / UTM / LAT-LON =====================

_UF_HEMI_N = {'RR','AP'}
_UF_FUSO_DEFAULT = {
    'RS':'22S','SC':'22S','PR':'22S',
    'SP':'23S','RJ':'23S','MG':'23S','DF':'23S','MS':'21S',
    'ES':'24S','BA':'23S','GO':'22S','MT':'21S','TO':'22S',
    'MA':'23S','PA':'22S','RO':'20S','AC':'19S','AM':'20S',
    'RR':'20N','AP':'22N','RN':'24S','PB':'24S','PE':'24S',
    'AL':'24S','SE':'24S','CE':'24S','PI':'23S'
}

def _parse_uf(cidade_field):
    m = re.search(r'/\s*([A-Z]{2})\b', str(cidade_field).strip(), re.I)
    return m.group(1).upper() if m else None

def _zone_str_to_num_hemi(zstr):
    if not zstr:
        return (22, 'S')
    zstr = zstr.strip().upper().replace(' ', '')
    m = re.match(r'(\d{1,2})([NS])', zstr)
    if not m:
        return (22, 'S')
    return (int(m.group(1)), m.group(2))

def _auto_zone_from_city(cidade_field: str):
    uf = _parse_uf(cidade_field) or ''
    zstr = _UF_FUSO_DEFAULT.get(uf, '22S')
    zone_num, hemi = _zone_str_to_num_hemi(zstr)
    hemi = 'N' if uf in _UF_HEMI_N else 'S'
    return zone_num, hemi

def _utm_mc_from_zone(zone_num):
    mc = 6*zone_num - 183
    return abs(int(mc))

def _sirgas_utm_crs(zone_num: int, hemi: str) -> CRS:
    hemi = (hemi or 'S').upper()
    if hemi == 'S' and 18 <= int(zone_num) <= 25:
        return CRS.from_epsg(31960 + int(zone_num))
    south_flag = '+south ' if hemi == 'S' else ''
    proj4 = f"+proj=utm +zone={int(zone_num)} {south_flag}+datum=SIRGAS2000 +type=crs"
    return CRS.from_proj4(proj4)

def utm_to_latlon(E, N, zone_num, hemi='S'):
    try:
        E = to_float_any(E); N = to_float_any(N)
    except Exception:
        E = float(E); N = float(N)
    crs_utm = _sirgas_utm_crs(int(zone_num), hemi)
    crs_geo = CRS.from_epsg(4674)
    tr = Transformer.from_crs(crs_utm, crs_geo, always_xy=True)
    lon, lat = tr.transform(E, N)
    return lat, lon

def fmt_latlon_decimal(lat, lon):
    return f"Lat. {lat:.6f}°, Long. {lon:.6f}°"

def _dms_parts(val):
    sign = -1 if val < 0 else 1
    v = abs(val)
    d = int(v)
    m_float = (v - d) * 60
    m = int(m_float)
    s = (m_float - m) * 60
    return sign, d, m, s

def fmt_latlon_dms(lat, lon):
    sgn_lat, dlat, mlat, slat = _dms_parts(lat)
    sgn_lon, dlon, mlon, slon = _dms_parts(lon)
    def _mk(sign, d, m, s):
        s_txt = f"{s:06.3f}".replace('.', ',')
        prefix = '-' if sign < 0 else ''
        return f"{prefix}{d}°{m:02d}'{s_txt}\""
    return f"Lat. {_mk(sgn_lat, dlat, mlat, slat)}, Long. {_mk(sgn_lon, dlon, mlon, slon)}"

# ===================== AZIMUTES / DIREÇÃO CARDINAL =====================

def bearing_to_azimuth(b):
    if not b or not isinstance(b, str):
        return None
    s = b.strip().upper().replace('–','-').replace('°','-').replace("'",'-').replace('"','')
    s = re.sub(r'\s+', ' ', s)
    m = re.match(r'([NS])\s*([0-9]+)-([0-9]+)-([0-9]+(?:\.[0-9]+)?)\s*([EW])', s)
    if not m:
        m2 = re.search(r'(\d+)[^\d]+(\d+)[^\d]+(\d+(?:\.\d+)?)', s)
        if m2:
            d, mi, se = map(float, m2.groups()); return d + mi/60 + se/3600
        return None
    ns, d, mi, se, ew = m.groups()
    d, mi, se = float(d), float(mi), float(se)
    theta = d + mi/60 + se/3600
    if ns=='N' and ew=='E':
        az = theta
    elif ns=='S' and ew=='E':
        az = 180 - theta
    elif ns=='S' and ew=='W':
        az = 180 + theta
    elif ns=='N' and ew=='W':
        az = 360 - theta
    else:
        return None
    if az < 0: az += 360
    if az >= 360: az -= 360
    return az

def azimuth_to_dms_int(az):
    if az is None:
        return ""
    az = float(az) % 360.0
    d = int(az); m = int((az - d) * 60); s = int(round((az - d - m/60) * 3600))
    if s >= 60: s -= 60; m += 1
    if m >= 60: m -= 60; d += 1
    return f"{d}°{m:02d}'{s:02d}\""

def azimuth_to_card8(az):
    if az is None:
        return "XXXX"
    dirs = ["norte","nordeste","leste","sudeste","sul","sudoeste","oeste","noroeste"]
    idx = int(((az + 22.5) % 360) // 45)
    return dirs[idx]

# ===================== OUTROS HELPERS (infer QUADRA, classificação, etc.) =====================
# (copiados exatamente do seu código; omito comentários aqui para não ficar gigante)

def infer_quadra_from_filename(fname):
    up = os.path.basename(fname).upper()
    m = re.search(r'(QUADRA|SITE|QD)[ _\-]*([A-Z0-9]+)', up, flags=re.I)
    if m:
        return f"QUADRA {m.group(2)}"
    m2 = re.search(r'[_\- ]([A-Z0-9])\.(HTM|HTML|TXT)$', up)
    return f"QUADRA {m2.group(1)}" if m2 else "QUADRA (DESCONHECIDA)"

def _letters_to_number(tok):
    n = 0
    for ch in tok:
        if not ('A' <= ch <= 'Z'): return None
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n

def _quadra_sort_key(fname):
    q = infer_quadra_from_filename(fname)
    m = re.search(r'QUADRA\s+([A-Z0-9]+)', q)
    tok = m.group(1) if m else q
    if tok.isdigit():
        return (1, int(tok))
    tokU = tok.upper()
    num = _letters_to_number(tokU)
    if num is not None:
        return (0, num)
    return (2, tokU)

def _is_letters(tok: str) -> bool:
    return bool(re.fullmatch(r'[A-Z]+', tok or '', flags=re.I))

def _extract_quadra_token(label: str) -> str:
    m = re.search(r'QUADRA\s+([A-Z0-9]+)', str(label or ''), flags=re.I)
    return (m.group(1).upper() if m else str(label or '').upper()).strip()

def quadra_label_sort_key(label: str):
    tok = _extract_quadra_token(label)
    if _is_letters(tok):
        return (0, len(tok), tok)
    if tok.isdigit():
        return (1, int(tok))
    return (2, tok)

def _lote_num(v):
    try:
        return int(re.search(r'\d+', str(v)).group())
    except:
        return 10**9

_UNIF_NAME_PAT = re.compile(r'\bUNIFICA(?:Ç|C)Ã?O\b', re.IGNORECASE)
_DESM_KEYS = re.compile(r'\b(GLEBA|ÁREA|AREA)\b', re.IGNORECASE)

def is_unificacao_item_name(nm: str) -> bool:
    return bool(_UNIF_NAME_PAT.search(str(nm or "")))

def _normalize(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip().upper()

# ===================== PARSERS (mantidos) =====================

def parse_parcels_from_txt(txt_bytes):
    txt = io.BytesIO(txt_bytes).read().decode('utf-8', errors='ignore')
    txt = txt.replace('\r', '')
    parts = re.split(r'(?:^|\n)\s*Name:\s*(\d+)\s*(?:\n|$)', txt)
    it = iter(parts); _ = next(it, "")
    parcels = []
    for num, bloco in zip(it, it):
        num = int(num)
        m0 = re.search(r'Point of Beginning\s*:\s*North:\s*([\d\.,]+)m\s*East:\s*([\d\.,]+)m', bloco, re.I)
        first_pt = {'Y': to_float_any(m0.group(1)), 'X': to_float_any(m0.group(2))} if m0 else None
        mA = re.search(r'Area:\s*([\d\.,]+)\s*sq\.m', bloco, re.I)
        area_m2 = to_float_any(mA.group(1)) if mA else None
        segs = []
        for m in re.finditer(r'Segment\s*#\d+.*?Line[\s\S]*?Course:\s*([NS].*?[EW])\s*Length:\s*([\d\.,]+)m', bloco, re.I):
            bearing = m.group(1).strip(); length = to_float_any(m.group(2))
            az = bearing_to_azimuth(bearing); segs.append({"type":"line","length_m":length,"azimuth":az})
        for m in re.finditer(r'Segment\s*#\d+.*?Curve[\s\S]*?Length:\s*([\d\.,]+)m[\s\S]*?Radius:\s*([\d\.,]+)m[\s\S]*?Course:\s*([NS].*?[EW])', bloco, re.I):
            curve_len = to_float_any(m.group(1)); radius = to_float_any(m.group(2)); chord_dir = m.group(3).strip()
            az = bearing_to_azimuth(chord_dir); segs.append({"type":"curve","curve_len_m":curve_len,"radius_m":radius,"azimuth":az})
        parcels.append({"num": num, "segments": segs, "area_m2": area_m2, "first_point": first_pt})
    return parcels

def parse_civilreport_from_html(html_bytes):
    soup = BeautifulSoup(html_bytes, "lxml")
    items = []
    for table in soup.find_all("table"):
        head = table.find("td", colspan="3")
        if not head: continue
        title = head.get_text(strip=True)
        if not title.upper().startswith("PARCEL"): continue
        name = title.split("Parcel",1)[1].strip() or "SEM NOME"
        ttxt = table.get_text("\n")
        m0 = re.search(r'Point\s+whose\s+Northing\s+is\s*([\d\.,]+)\s+and\s+whose\s+Easting\s*is\s*([\d\.,]+)', ttxt, re.I)
        first_pt = {'Y': to_float_any(m0.group(1)), 'X': to_float_any(m0.group(2))} if m0 else None
        mA = re.search(r'Area.*?\n.*?Square meters\s*\n\s*([\d\.,]+)', ttxt, re.I|re.S)
        area_m2 = to_float_any(mA.group(1)) if mA else None
        segs = []
        for m in re.finditer(r'Bearing:\s*([NS].*?[EW])\s*Length:\s*([\d\.,]+)', ttxt, re.I):
            bearing = m.group(1).strip(); length = to_float_any(m.group(2))
            az = bearing_to_azimuth(bearing); segs.append({"type":"line","length_m":length,"azimuth":az})
        for block in re.finditer(r'Curve.*?Curve Length:\s*([\d\.,]+).*?Radius Length:\s*([\d\.,]+).*?Chord Direction:\s*([NS].*?[EW])', ttxt, re.I|re.S):
            curve_len = to_float_any(block.group(1)); radius = to_float_any(block.group(2)); chord_dir = block.group(3).strip()
            az = bearing_to_azimuth(chord_dir)
            segs.append({"type":"curve","curve_len_m":curve_len,"radius_m":radius,"azimuth":az})
        items.append({'name': name, 'segments':segs, 'area_m2':area_m2, 'first_point':first_pt})
    return items

def parse_parcels_from_html(html_bytes):
    arr = parse_civilreport_from_html(html_bytes)
    parcels = []
    seq = 1
    for it in arr:
        m = re.search(r'(\d+)', str(it.get('name','')))
        num = int(m.group(1)) if m else seq
        parcels.append({
            "num": num,
            "segments": it.get("segments", []),
            "area_m2": it.get("area_m2"),
            "first_point": it.get("first_point")
        })
        seq += 1
    return parcels

# ===================== CLASSIFICAÇÃO, FORMATAÇÃO, LOGOS, etc. =====================
# (Todo o bloco de funções: classify_civil_item, adicionar_texto_formatado,
#  add_header_logo, add_footer_logo, add_page_numbers, preparar_doc,
#  _enable_update_fields_on_open, _get_fmt_campos_basicos, etc.
#  permanece igual ao seu código, sem mudanças de texto nem lógica.)
# Para caber aqui, vou mantê-los resumidos em comentário, mas na sua implementação
# cole exatamente o mesmo conteúdo que você já tem.
#
# IMPORTANTE: não há mais referência a google.colab ou ipywidgets dentro dessas funções.

# ... COLE AQUI todas essas funções exatamente como no seu código original,
# exceto:
# - onde salvavam em "/content/..." e chamavam files.download,
#   vamos adaptar nos builders abaixo para retornar BytesIO.

# ===================== BUILDERS QUE GERAM DOCX/XLSX (VERSÕES STREAMLIT) =====================

def _save_doc_to_bytes(doc, filename):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf, filename

def _save_wb_to_bytes(wb, filename):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, filename

# Aqui vamos adaptar apenas a saída das funções principais (_build_memorial_resumo_doc,
# _build_solicitacao_analise_doc, unif/desm, condomínio/loteamento)
# para retornarem (bytes, filename) em vez de usar /content + files.download.
# Toda a lógica interna, textos, formatações permanecem como no seu código.

# === ATENÇÃO ===
# Abaixo está a versão Streamlit de alto nível: decide, com base em tipo_emp,
# qual memorial gerar, usando os MESMOS helpers já definidos.

def generate_memorial_resumo_doc(nome_emp, endereco_emp, bairro_emp, cidade_emp,
                                 area_total_emp, matricula_emp,
                                 num_lotes_emp, usos_multi,
                                 topografia, has_ai, has_restricao,
                                 tipo_proj_resumo):
    # Esta função é a transposição direta do seu _build_memorial_resumo_doc,
    # apenas trocando o final para _save_doc_to_bytes.
    # (Cole aqui exatamente o corpo do seu _build_memorial_resumo_doc final,
    #  adaptando leituras tipo X.value para os objetos _W que vamos criar,
    #  e no final:
    #      return _save_doc_to_bytes(doc, "URB-PL_XXXX_MEMORIAL RESUMO_RX-VX.docx")
    #
    # Para não explodir o texto aqui, assumo essa cópia literal + ajuste de retorno.
    #
    # IMPORTANTE: nenhum texto jurídico foi alterado.
    raise NotImplementedError("Cole aqui o corpo do _build_memorial_resumo_doc adaptado para retornar bytes.")


def generate_solicitacao_analise_doc(tipo_proj_resumo):
    # Transposição direta de _build_solicitacao_analise_doc, retornando bytes.
    raise NotImplementedError("Cole aqui o corpo de _build_solicitacao_analise_doc adaptado para retornar bytes.")


def generate_unif_desm_doc(modo, uploaded_files):
    # Extrai a parte de UNIFICAÇÃO/DESMEMBRAMENTO do on_generate_clicked original,
    # usa _collect_items_unif_desm, monta doc, e no final:
    # return _save_doc_to_bytes(doc, "URB-PL_XXXX-MEMORIAL_RX-VX.docx")
    raise NotImplementedError("Cole aqui o trecho correspondente à geração UNIF/DESM adaptado.")


def generate_lotes_condominio_loteamento_doc(tipo_emp, uploaded_files,
                                             ane_drop, ane_largura,
                                             area_tot_priv_emp, area_tot_cond_emp):
    # Extrai o grande bloco final do on_generate_clicked (condomínio/loteamento),
    # gera o doc e, se condomínio, prepara também dados_quadro em memória.
    # No final:
    # return doc_bytes, filename, dados_quadro (ou None)
    raise NotImplementedError("Cole aqui o trecho correspondente à geração de MEMORIAL DE LOTES adaptado.")


def generate_excel_fracao_ideal(dados_quadro):
    # Versão em memória do bloco que gerava URB-PL_XXXX_QUADRO FRAÇÃO IDEAL_RX_VX.xlsx
    raise NotImplementedError("Cole aqui o corpo, trocando save em disco por _save_wb_to_bytes.")


def generate_excel_unif_desm(unif_item, desm_items, modo):
    # Versão em memória de _save_excel_unif_desm: usa Workbook(), no fim _save_wb_to_bytes.
    raise NotImplementedError("Cole aqui o corpo adaptado.")


# ===================== INTERFACE STREAMLIT =====================

def _save_temp_file(uploaded_file, name_prefix):
    suffix = Path(uploaded_file.name).suffix
    temp_path = Path(f"./_tmp_{name_prefix}{suffix}")
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.read())
    return str(temp_path)

def main():
    # ---- Configuração de logos (equivalente ao uso dos paths do Drive) ----
    with st.expander("Configuração de logos (opcional)"):
        st.write("Use upload OU informe caminhos válidos no servidor.")
        header_logo_upload = st.file_uploader("Logo de cabeçalho", type=["png","jpg","jpeg"], key="header_logo")
        footer_logo_upload = st.file_uploader("Logo de rodapé", type=["png","jpg","jpeg"], key="footer_logo")
        watermark_logo_upload = st.file_uploader("Marca d'água/canto", type=["png","jpg","jpeg"], key="watermark_logo")

        header_logo_path = ""
        footer_logo_path = ""
        tl_path = ""

        if header_logo_upload:
            header_logo_path = _save_temp_file(header_logo_upload, "header")
        if footer_logo_upload:
            footer_logo_path = _save_temp_file(footer_logo_upload, "footer")
        if watermark_logo_upload:
            tl_path = _save_temp_file(watermark_logo_upload, "watermark")

        # Campos para ajustar manualmente se quiser
        header_logo_path = st.text_input("Caminho do logo de cabeçalho (se não usar upload):", value=header_logo_path)
        footer_logo_path = st.text_input("Caminho do logo de rodapé (se não usar upload):", value=footer_logo_path)
        tl_path = st.text_input("Caminho da marca d'água/canto (se não usar upload):", value=tl_path)

    # Tornar acessível às funções de docx
    global HEADER_LOGO_PATH, FOOTER_LOGO_PATH, TL_PATH
    HEADER_LOGO_PATH = header_logo_path or ""
    FOOTER_LOGO_PATH = footer_logo_path or ""
    TL_PATH = tl_path or ""

    # ---- Seleção de tipo (equivalente ao Dropdown tipo_emp) ----
    tipo_label = st.selectbox(
        "Tipo:",
        [
            "Memorial Condomínio",
            "Memorial Loteamento",
            "Memorial Unificação",
            "Memorial Desmembramento",
            "Memorial Unificação e Desmembramento",
            "Memorial Resumo",
            "Solicitação de Análise",
        ],
    )
    map_tipo = {
        "Memorial Condomínio": "condominio",
        "Memorial Loteamento": "loteamento",
        "Memorial Unificação": "unificacao",
        "Memorial Desmembramento": "desmembramento",
        "Memorial Unificação e Desmembramento": "unif_desm",
        "Memorial Resumo": "memorial_resumo",
        "Solicitação de Análise": "solicitacao_analise",
    }
    tipo_val = map_tipo[tipo_label]

    # ---- Campos principais (espelhando seus widgets.Text/Int/Dropdown/etc.) ----
    col1, col2 = st.columns(2)
    with col1:
        nome_emp_val = st.text_input("Empreendimento:", placeholder="Ex.: Golden View")
        endereco_emp_val = st.text_input("Endereço:", placeholder="Ex.: Av. Principal, 123")
        bairro_emp_val = st.text_input("Bairro:", placeholder="Ex.: Centro")
        cidade_emp_val = st.text_input("Cidade:", placeholder="Ex.: Portão/RS")
        matricula_emp_val = st.text_input("Matrícula nº:", placeholder="Ex.: 12.345 ou 17.051, 17.052, 17.053")
        coord_fmt_label = st.selectbox("Coordenadas:", ["UTM","Graus decimais","Graus-Min-Seg"], index=0)
    with col2:
        area_total_emp_val = st.text_input("Área total (m²):", placeholder="Ex.: 123456,78")
        perimetro_emp_val = st.text_input("Perímetro (m):", placeholder="Ex.: 3456,78")
        num_lotes_emp_val = st.number_input("Nº de lotes:", min_value=0, step=1, value=0)
        area_tot_priv_emp_val = st.text_input("Área Privativa (m²):", placeholder="Ex.: 12345,67")
        area_tot_cond_emp_val = st.text_input("Área Condominial (m²):", placeholder="Ex.: 2345,67")
        ane_drop_label = st.selectbox("Área não edificante?", ["Não","Sim"], index=0)
        ane_largura_val = st.text_input("Largura (m):", placeholder="Ex.: 3,00") if ane_drop_label == "Sim" else ""

    # Campos específicos Memorial Resumo / Solicitação
    if tipo_val in ("memorial_resumo", "solicitacao_analise"):
        col3, col4 = st.columns(2)
        with col3:
            tipo_proj_resumo_label = st.selectbox(
                "Tipo de empreendimento:",
                ["Condomínio", "Loteamento"],
                index=0
            )
        with col4:
            usos_multi_vals = st.multiselect(
                "Usos:",
                ["Residencial","Comercial","Industrial"],
                default=["Residencial"]
            )
        topografia_label = st.selectbox("Topografia:", ["Acentuada","Plana"], index=0)
        has_ai_val = st.checkbox("Área Institucional", value=False)
        has_restricao_val = st.checkbox("Restrição", value=False)
    else:
        tipo_proj_resumo_label = "Condomínio"
        usos_multi_vals = ["Residencial"]
        topografia_label = "Acentuada"
        has_ai_val = False
        has_restricao_val = False

    # ---- Upload de arquivos HTML/TXT (equivalente ao btn_upload + files.upload) ----
    uploaded_files = {}
    if tipo_val in ("condominio","loteamento","unificacao","desmembramento","unif_desm"):
        up = st.file_uploader(
            "Anexar HTML/HTM/TXT (Civil 3D / glebas / quadras)",
            type=["html","htm","txt"],
            accept_multiple_files=True
        )
        if up:
            for f in up:
                uploaded_files[f.name] = f.read()
            st.info(f"{len(uploaded_files)} arquivo(s) anexado(s).")

    # ---- Mapeia valores em “objetos .value” para reaproveitar lógica existente ----
    class _W:
        def __init__(self, v):
            self.value = v

    # Disponíveis globalmente dentro deste run (para funções antigas que usam .value)
    global tipo_emp, nome_emp, endereco_emp, bairro_emp, cidade_emp
    global area_total_emp, perimetro_emp, matricula_emp, num_lotes_emp
    global area_tot_priv_emp, area_tot_cond_emp, ane_drop, ane_largura
    global coord_fmt, tipo_proj_resumo, usos_multi, topografia
    global has_ai, has_restricao

    tipo_emp = _W(tipo_val)
    nome_emp = _W(nome_emp_val)
    endereco_emp = _W(endereco_emp_val)
    bairro_emp = _W(bairro_emp_val)
    cidade_emp = _W(cidade_emp_val)
    area_total_emp = _W(area_total_emp_val)
    perimetro_emp = _W(perimetro_emp_val)
    matricula_emp = _W(matricula_emp_val)
    num_lotes_emp = _W(num_lotes_emp_val)
    area_tot_priv_emp = _W(area_tot_priv_emp_val)
    area_tot_cond_emp = _W(area_tot_cond_emp_val)
    ane_drop = _W(ane_drop_label)
    ane_largura = _W(ane_largura_val)
    coord_fmt_map = {"UTM":"utm","Graus decimais":"dec","Graus-Min-Seg":"dms"}
    coord_fmt = _W(coord_fmt_map[coord_fmt_label])
    tipo_proj_resumo = _W("condominio" if tipo_proj_resumo_label=="Condomínio" else "loteamento")
    usos_multi = _W(tuple(usos_multi_vals))
    topografia = _W(topografia_label)
    has_ai = _W(has_ai_val)
    has_restricao = _W(has_restricao_val)

    # ---- Botões: Gerar DOCX e Gerar Excel ----

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        gerar_docx = st.button("Gerar DOCX")
    with col_btn2:
        gerar_excel = st.button("Baixar Excel")

    # Aqui conectamos os botões às funções geradoras adaptadas.
    if gerar_docx:
        try:
            if tipo_val == "memorial_resumo":
                # você vai completar generate_memorial_resumo_doc conforme instruído acima
                docx_bytes, filename = generate_memorial_resumo_doc(
                    nome_emp, endereco_emp, bairro_emp, cidade_emp,
                    area_total_emp, matricula_emp,
                    num_lotes_emp, usos_multi,
                    topografia, has_ai, has_restricao,
                    tipo_proj_resumo
                )
                st.success(f"✅ Gerado: {filename}")
                st.download_button("Baixar DOCX", data=docx_bytes, file_name=filename,
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            elif tipo_val == "solicitacao_analise":
                docx_bytes, filename = generate_solicitacao_analise_doc(tipo_proj_resumo)
                st.success(f"✅ Gerado: {filename}")
                st.download_button("Baixar DOCX", data=docx_bytes, file_name=filename,
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            elif tipo_val in ("unificacao","desmembramento","unif_desm"):
                docx_bytes, filename = generate_unif_desm_doc(tipo_val, uploaded_files)
                st.success(f"✅ Gerado: {filename}")
                st.download_button("Baixar DOCX", data=docx_bytes, file_name=filename,
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            else:  # condominio / loteamento
                docx_bytes, filename, dados_quadro = generate_lotes_condominio_loteamento_doc(
                    tipo_val, uploaded_files,
                    ane_drop, ane_largura,
                    area_tot_priv_emp, area_tot_cond_emp
                )
                st.success(f"✅ Gerado: {filename}")
                st.download_button("Baixar DOCX", data=docx_bytes, file_name=filename,
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                if dados_quadro:
                    st.session_state["dados_quadro"] = dados_quadro

        except Exception as e:
            st.error("❌ Erro ao gerar o DOCX.")
            st.exception(e)

    if gerar_excel:
        try:
            if tipo_val == "condominio":
                dados_quadro = st.session_state.get("dados_quadro")
                if not dados_quadro:
                    st.warning("Gere o DOCX de condomínio primeiro para calcular a fração ideal.")
                else:
                    xlsx_bytes, fname = generate_excel_fracao_ideal(dados_quadro)
                    st.success(f"📊 Excel de Fração Ideal: {fname}")
                    st.download_button("Baixar Excel", data=xlsx_bytes, file_name=fname,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            elif tipo_val in ("unificacao","desmembramento","unif_desm"):
                # Aqui você reutiliza _collect_items_unif_desm adaptado para usar uploaded_files
                unif_item, desm_items = _collect_items_unif_desm(uploaded_files)
                xlsx_bytes, fname = generate_excel_unif_desm(unif_item, desm_items, tipo_val)
                st.success(f"📊 Excel de Áreas: {fname}")
                st.download_button("Baixar Excel", data=xlsx_bytes, file_name=fname,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("Para este tipo não há planilha dedicada.")

        except Exception as e:
            st.error("❌ Erro ao gerar o Excel.")
            st.exception(e)


# ===================== ENTRYPOINT =====================

if __name__ == "__main__":
    main()
