import os
import re
import io
import math
from io import BytesIO
from pathlib import Path
from datetime import datetime

from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_COLOR_INDEX, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from num2words import num2words
import pandas as pd
from pyproj import CRS, Transformer
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ===================== CONFIG / LOGOS =====================

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_WATERMARK = BASE_DIR / "assets" / "marca_dagua.png"
DEFAULT_HEADER_LOGO = BASE_DIR / "assets" / "logo_cabecalho.png"
DEFAULT_FOOTER_LOGO = BASE_DIR / "assets" / "logo_rodape.png"

# ===================== HELPERS GERAIS =====================

def _fmt_br(v, casas=2):
    try:
        return f"{float(v):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

def _to_float_br(txt: str) -> float:
    return float(str(txt).replace('.', '').replace(',', '.'))

def to_float_any(s):
    s = str(s).strip()
    if ',' in s and '.' in s:
        return float(s.replace('.', '').replace(',', '.'))
    if ',' in s:
        return float(s.replace(',', '.'))
    return float(s or 0)

def extenso_metros(v):
    v = round(float(v or 0), 2)
    m = int(v)
    cm = int(round((v - m) * 100))
    partes = []
    if m > 0:
        partes.append(num2words(m, lang='pt_BR') + (" metro" if m == 1 else " metros"))
    if cm > 0:
        partes.append(num2words(cm, lang='pt_BR') + (" centímetro" if cm == 1 else " centímetros"))
    return " e ".join(partes) if partes else "zero metro"

def area_por_extenso(v):
    v = round(float(v or 0), 2)
    m2 = int(v)
    cent = int(round((v - m2) * 100))
    if cent == 0:
        return f"{num2words(m2, lang='pt_BR')} metros quadrados"
    return f"{num2words(m2, lang='pt_BR')} metros quadrados e {num2words(cent, lang='pt_BR')} centésimos"

def hectares_from_m2(v):
    return float(v or 0) / 10000.0

# ===================== FORMATOS NOME / CIDADE =====================

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
    if not s:
        return ""
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

def _title_case_name(nome: str) -> str:
    nome = (nome or "").strip().lower()
    return ' '.join(w.capitalize() for w in nome.split())

def _cidade_sem_uf(txt: str) -> str:
    s = str(txt or "XXXX").strip()
    if "/" in s:
        s = s.split("/", 1)[0].strip()
    return s if s else "XXXX"

# ===================== COORDENADAS E UTM =====================

_UF_HEMI_N = {'RR', 'AP'}
_UF_FUSO_DEFAULT = {
    'RS': '22S', 'SC': '22S', 'PR': '22S',
    'SP': '23S', 'RJ': '23S', 'MG': '23S', 'DF': '23S', 'MS': '21S',
    'ES': '24S', 'BA': '23S', 'GO': '22S', 'MT': '21S', 'TO': '22S',
    'MA': '23S', 'PA': '22S', 'RO': '20S', 'AC': '19S', 'AM': '20S',
    'RR': '20N', 'AP': '22N', 'RN': '24S', 'PB': '24S', 'PE': '24S',
    'AL': '24S', 'SE': '24S', 'CE': '24S', 'PI': '23S'
}

def _parse_uf(cidade_field: str):
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

def _utm_mc_from_zone(zone_num: int) -> int:
    mc = 6 * int(zone_num) - 183
    return abs(int(mc))

def _sirgas_utm_crs(zone_num: int, hemi: str) -> CRS:
    hemi = (hemi or 'S').upper()
    if hemi == 'S' and 18 <= int(zone_num) <= 25:
        return CRS.from_epsg(31960 + int(zone_num))
    south_flag = '+south ' if hemi == 'S' else ''
    proj4 = f"+proj=utm +zone={int(zone_num)} {south_flag}+datum=SIRGAS2000 +type=crs"
    return CRS.from_proj4(proj4)

def utm_to_latlon(E, N, zone_num, hemi='S'):
    E = to_float_any(E)
    N = to_float_any(N)
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

# ===================== AZIMUTES =====================

def bearing_to_azimuth(b):
    if not b or not isinstance(b, str):
        return None
    s = b.strip().upper().replace('–', '-').replace('°', '-').replace("'", '-').replace('"', '')
    s = re.sub(r'\s+', ' ', s)
    m = re.match(r'([NS])\s*([0-9]+)-([0-9]+)-([0-9]+(?:\.[0-9]+)?)\s*([EW])', s)
    if not m:
        m2 = re.search(r'(\d+)[^\d]+(\d+)[^\d]+(\d+(?:\.\d+)?)', s)
        if m2:
            d, mi, se = map(float, m2.groups())
            return d + mi / 60 + se / 3600
        return None
    ns, d, mi, se, ew = m.groups()
    d, mi, se = float(d), float(mi), float(se)
    theta = d + mi / 60 + se / 3600
    if ns == 'N' and ew == 'E':
        az = theta
    elif ns == 'S' and ew == 'E':
        az = 180 - theta
    elif ns == 'S' and ew == 'W':
        az = 180 + theta
    elif ns == 'N' and ew == 'W':
        az = 360 - theta
    else:
        return None
    az %= 360
    return az

def azimuth_to_dms_int(az):
    if az is None:
        return ""
    az = float(az) % 360.0
    d = int(az)
    m = int((az - d) * 60)
    s = int(round((az - d - m / 60) * 3600))
    if s >= 60:
        s -= 60
        m += 1
    if m >= 60:
        m -= 60
        d += 1
    return f"{d}°{m:02d}'{s:02d}\""

def azimuth_to_card8(az):
    if az is None:
        return "XXXX"
    dirs = ["norte", "nordeste", "leste", "sudeste", "sul", "sudoeste", "oeste", "noroeste"]
    idx = int(((az + 22.5) % 360) // 45)
    return dirs[idx]

# ===================== LABELS / ORDEM =====================

def _normalize(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip().upper()

def infer_quadra_from_filename(fname):
    up = os.path.basename(fname).upper()
    m = re.search(r'(QUADRA|SITE|QD)[ _\-]*([A-Z0-9]+)', up, flags=re.I)
    if m:
        return f"QUADRA {m.group(2)}"
    m2 = re.search(r'[_\- ]([A-Z0-9])\.(HTM|HTML|TXT)$', up)
    return f"QUADRA {m2.group(1)}" if m2 else "QUADRA (DESCONHECIDA)"

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
    except Exception:
        return 10**9

# ===================== CLASSIFICAÇÃO CIVIL 3D =====================

_UNIF_NAME_PAT = re.compile(r'\bUNIFICA(?:Ç|C)Ã?O\b', re.IGNORECASE)

def is_unificacao_item_name(nm: str) -> bool:
    return bool(_UNIF_NAME_PAT.search(str(nm or "")))

def classify_civil_item(name):
    n = _normalize(name)
    if re.search(r'\b(ALARGAMENTO(S)?|ACESSO(S)?( DE SERVIÇO(S)?)?|RODOVI(A|Á)S?|RUA(S)?|AVENIDA(S)?|PEATONAL(IS)?|CANTEIRO(S)?|ACESSOS?)\b', n):
        return ('viario', 'DESCRIÇÃO DE SISTEMA VIÁRIO')
    if re.search(r'^(AVS?\b)|\bÁREA(S)? VERDE(S)?\b', n):
        return ('verde', 'DESCRIÇÃO DE ÁREAS VERDES')
    if 'ÁREA VERDE DE PRESERVAÇÃO' in n or 'AREA VERDE DE PRESERVACAO' in n:
        return ('verde_preservacao', 'DESCRIÇÃO DE ÁREA VERDE DE PRESERVAÇÃO')
    if re.search(r'(PRESERVAÇÃO PERMANENTE|PRESERVACAO PERMANENTE|\bAPP\b|RESTRIÇ|RESTRICAO|PRESERVAÇÃO AMBIENTAL|PRESERVACAO AMBIENTAL)', n):
        if 'RESTRI' in n:
            return ('app', 'DESCRIÇÃO DE RESTRIÇÕES')
        if 'PRESERVAÇÃO AMBIENTAL' in n or 'PRESERVACAO AMBIENTAL' in n:
            return ('app', 'DESCRIÇÃO DE ÁREA DE PRESERVAÇÃO AMBIENTAL')
        return ('app', 'DESCRIÇÃO DE ÁREA DE PRESERVAÇÃO PERMANENTE')
    if re.search(r'\bAI(\b|\s)|\bÁREA(S)? INSTITUCIONAL(IS)?\b|\bAREA(S)? INSTITUCIONAL(IS)?\b', n):
        return ('institucional', 'DESCRIÇÃO DE ÁREAS INSTITUCIONAIS')
    if re.search(r'RESERVA TÉCNICA|RESERVA TECNICA|\bETE\b|\bEBE\b|\bETA\b|\bEBA\b|ESTAÇÃO DE BOMBEAMENTO|ESTACAO DE BOMBEAMENTO|ESTAÇÃO DE TRATAMENTO|ESTACAO DE TRATAMENTO', n):
        return ('reserva_tecnica', 'DESCRIÇÃO DE RESERVA TÉCNICA')
    if 'REMANESCENTE' in n:
        return ('remanescente', 'DESCRIÇÃO DE ÁREA REMANESCENTE')
    if re.search(r'ÁREA(S)? CONDOMINIA(L|IS)|\bAC\s*\d+\b|AREA(S)? CONDOMINIA(L|IS)', n):
        return ('condominial', 'DESCRIÇÃO DE ÁREAS CONDOMINIAIS')
    if n.startswith('QUADRA'):
        return ('quadras', 'DESCRIÇÃO DE QUADRAS')
    return ('outros', 'DESCRIÇÃO DE OUTRAS ÁREAS')

# ===================== PARSERS CIVIL 3D =====================

def parse_parcels_from_txt(txt_bytes: bytes):
    txt = io.BytesIO(txt_bytes).read().decode('utf-8', errors='ignore')
    txt = txt.replace('\r', '')
    parts = re.split(r'(?:^|\n)\s*Name:\s*(\d+)\s*(?:\n|$)', txt)
    it = iter(parts)
    _ = next(it, "")
    parcels = []
    for num, bloco in zip(it, it):
        num = int(num)
        m0 = re.search(r'Point of Beginning\s*:\s*North:\s*([\d\.,]+)m\s*East:\s*([\d\.,]+)m', bloco, re.I)
        first_pt = {'Y': to_float_any(m0.group(1)), 'X': to_float_any(m0.group(2))} if m0 else None
        mA = re.search(r'Area:\s*([\d\.,]+)\s*sq\.m', bloco, re.I)
        area_m2 = to_float_any(mA.group(1)) if mA else None
        segs = []
        for m in re.finditer(r'Course:\s*([NS].*?[EW])\s*Length:\s*([\d\.,]+)m', bloco, re.I):
            bearing = m.group(1).strip()
            length = to_float_any(m.group(2))
            az = bearing_to_azimuth(bearing)
            segs.append({"type": "line", "length_m": length, "azimuth": az})
        parcels.append({"num": num, "segments": segs, "area_m2": area_m2, "first_point": first_pt})
    return parcels

def parse_civilreport_from_html(html_bytes: bytes):
    soup = BeautifulSoup(html_bytes, "lxml")
    items = []
    for table in soup.find_all("table"):
        head = table.find("td", colspan="3")
        if not head:
            continue
        title = head.get_text(strip=True)
        if not title.upper().startswith("PARCEL"):
            continue
        name = title.split("Parcel", 1)[1].strip() or "SEM NOME"
        ttxt = table.get_text("\n")
        m0 = re.search(r'Northing\s+is\s*([\d\.,]+)\s+and\s+whose\s+Easting\s*is\s*([\d\.,]+)', ttxt, re.I)
        first_pt = {'Y': to_float_any(m0.group(1)), 'X': to_float_any(m0.group(2))} if m0 else None
        mA = re.search(r'Square meters\s*\n\s*([\d\.,]+)', ttxt, re.I)
        area_m2 = to_float_any(mA.group(1)) if mA else None
        segs = []
        for m in re.finditer(r'Bearing:\s*([NS].*?[EW])\s*Length:\s*([\d\.,]+)', ttxt, re.I):
            bearing = m.group(1).strip()
            length = to_float_any(m.group(2))
            az = bearing_to_azimuth(bearing)
            segs.append({"type": "line", "length_m": length, "azimuth": az})
        items.append({'name': name, 'segments': segs, 'area_m2': area_m2, 'first_point': first_pt})
    return items

def parse_parcels_from_html(html_bytes: bytes):
    arr = parse_civilreport_from_html(html_bytes)
    parcels = []
    seq = 1
    for it in arr:
        m = re.search(r'(\d+)', str(it.get('name', '')))
        num = int(m.group(1)) if m else seq
        parcels.append({
            "num": num,
            "segments": it.get("segments", []),
            "area_m2": it.get("area_m2"),
            "first_point": it.get("first_point")
        })
        seq += 1
    return parcels

# ===================== DOCX HELPERS =====================

def _set_run_defaults(run, bold=False):
    run.font.name = 'Calibri'
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.bold = bool(bold)

def _add_hl(paragraph, txt="XXXX", bold=False):
    run = paragraph.add_run(txt)
    _set_run_defaults(run, bold=bold)
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return run

def heading(doc, text):
    h = doc.add_heading('', level=1)
    run = h.add_run(text)
    _set_run_defaults(run, bold=True)
    return h

def _heading_num(doc, idx, title):
    return heading(doc, f"{idx}. {title}")

def _run_xxxx(par):
    r = par.add_run("XXXX")
    _set_run_defaults(r)
    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return r

def add_page_numbers(document):
    section = document.sections[-1]
    p = section.footer.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE \\* MERGEFORMAT')
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    r.append(rPr)
    t = OxmlElement('w:t')
    t.text = ''
    r.append(t)
    fld.append(r)
    p._p.append(fld)

def add_footer_left_text(doc, lines, size_pt=10):
    for section in doc.sections:
        p = section.footer.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        for i, line in enumerate(lines):
            run = p.add_run(line)
            run.font.name = 'Calibri'
            run.font.size = Pt(size_pt)
            run.font.color.rgb = RGBColor(0, 0, 0)
            if i < len(lines) - 1:
                run.add_break()

def _apply_moderate_margins(doc):
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

def _add_image_run(paragraph, image, width=None):
    if image is None:
        return
    if isinstance(image, (str, Path)):
        if not os.path.exists(image):
            return
        paragraph.add_run().add_picture(str(image), width=width)
    elif isinstance(image, bytes):
        paragraph.add_run().add_picture(BytesIO(image), width=width)

def preparar_doc(header_logo=None, footer_logo=None, watermark_logo=None):
    doc = Document()
    _apply_moderate_margins(doc)

    for section in doc.sections:
        section.header_distance = Inches(0.8)
        hp = section.header.paragraphs[0] if section.header.paragraphs else section.header.add_paragraph()
        _add_image_run(hp, header_logo or DEFAULT_HEADER_LOGO, width=Inches(1.4))

    sec = doc.sections[0]
    wp = sec.header.add_paragraph()
    _add_image_run(wp, watermark_logo or DEFAULT_WATERMARK, width=Cm(6.46))

    for section in doc.sections:
        section.footer_distance = Inches(0.3)
        fp = section.footer.add_paragraph()
        fp.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        _add_image_run(fp, footer_logo or DEFAULT_FOOTER_LOGO, width=Inches(1.6))

    return doc

def _enable_update_fields_on_open(doc):
    settings_el = doc.settings._element
    for el in settings_el.iterchildren():
        if el.tag == qn('w:updateFields'):
            el.set(qn('w:val'), 'true')
            return
    upd = OxmlElement('w:updateFields')
    upd.set(qn('w:val'), 'true')
    settings_el.append(upd)

def _add_toc(doc):
    p = doc.add_paragraph()
    r = p.add_run()
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    fld_begin.set(qn('w:dirty'), 'true')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = r'TOC \o "1-3" \h \z \u'
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    r_tmp = OxmlElement('w:r')
    t_tmp = OxmlElement('w:t')
    t_tmp.text = "Sumário será atualizado ao abrir o documento…"
    r_tmp.append(t_tmp)
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    r._r.append(fld_begin)
    r._r.append(instr)
    r._r.append(fld_sep)
    r._r.append(r_tmp)
    r._r.append(fld_end)

def _add_centered(doc, text, bold=False):
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r = p.add_run(text)
    _set_run_defaults(r, bold=bold)

def _join_com_e(lista):
    lista = [str(x) for x in lista if str(x).strip()]
    if not lista:
        return ""
    if len(lista) == 1:
        return lista[0]
    return ", ".join(lista[:-1]) + " e " + lista[-1]

# ===================== TEXTO SEGMENTOS / ÁREAS =====================

def _propaga_vertices(first_point: dict, segments: list,
                      coord_fmt_str: str = 'utm',
                      zone_num: int = 22,
                      hemi: str = 'S'):
    if not first_point or not segments:
        return []
    x = float(first_point["X"])
    y = float(first_point["Y"])
    p_idx = 1
    rows = []
    for seg in segments:
        az = float(seg.get("azimuth") or 0.0)
        rad = math.radians(az)
        if seg.get("type") == "line":
            L = float(seg.get("length_m") or 0.0)
            dx = math.sin(rad) * L
            dy = math.cos(rad) * L
            x2, y2 = x + dx, y + dy
            dist = round(L, 2)
            raio = None
        else:
            arc = float(seg.get("curve_len_m") or 0.0)
            R_ = float(seg.get("radius_m") or 0.0)
            theta = (arc / R_) if R_ > 0 else 0.0
            chord = 2.0 * R_ * math.sin(theta / 2.0)
            dx = math.sin(rad) * chord
            dy = math.cos(rad) * chord
            x2, y2 = x + dx, y + dy
            dist = round(arc, 2)
            raio = round(R_, 2) if R_ else None

        if coord_fmt_str == 'utm':
            c1 = _fmt_br(x2, 2); c2 = _fmt_br(y2, 2)
        else:
            lat, lon = utm_to_latlon(x2, y2, zone_num, hemi)
            if coord_fmt_str == 'dec':
                c1 = f"{lon:.6f}°".replace('.', ',')
                c2 = f"{lat:.6f}°".replace('.', ',')
            else:
                c1 = fmt_latlon_dms(lat, lon)
                c2 = ""

        rows.append({
            "DE": f"P{p_idx}",
            "PARA": f"P{p_idx + 1}",
            "COORD_1": c1,
            "COORD_2": c2,
            "AZIMUTE": azimuth_to_dms_int(az),
            "DISTANCIA (m)": dist,
            "RAIO (m)": raio,
            "CONFRONTANTE": ""
        })
        x, y = x2, y2
        p_idx += 1
    return rows

def adicionar_texto_formatado(doc, texto):
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    bold_pat = (
        r'(?:LOTE\s+\d+\s*–\s*QUADRA\s+[A-Z0-9]+:)'
        r'|(?:LOTE\s+\d+\s+da\s+QUADRA\s+[A-Z0-9]+)'
        r'|(?:(?<!Y=\s)(?<!X=\s)\d{1,3}(?:\.\d{3})*,\d+m²)'
        r'|(?:(?<!Y=\s)(?<!X=\s)\d{1,3}(?:\.\d{3})*,\d+m)'
    )
    coord_pat = (
        r'(?:Y=\s*\d{1,3}(?:\.\d{3})*,\d+m|X=\s*\d{1,3}(?:\.\d{3})*,\d+m)'
        r'|(?:Lat\.\s*-?\d+\.\d+°\s*,\s*Long\.\s*-?\d+\.\d+°)'
        r'|(?:Lat\.\s*-?\d+°\d{2}\'\d{2}(?:,\d+)?\"\s*,\s*Long\.\s*-?\d+°\d{2}\'\d{2}(?:,\d+)?\")'
    )
    dms_pat = r'\d{1,3}°\d{2}\'\d{2}(?:,\d{1,3})?"'
    bold_marker_pat = r'\[\[B\]\](.*?)\[\[/B\]\]'

    tok = re.compile(
        f'({bold_pat})|(XXXX)|({coord_pat})|({dms_pat})|({bold_marker_pat})',
        flags=re.IGNORECASE | re.DOTALL
    )

    i = 0
    while i < len(texto):
        m = tok.search(texto, i)
        if not m:
            resto = texto[i:]
            if resto:
                run = p.add_run(resto)
                _set_run_defaults(run)
            break

        pref = texto[i:m.start()]
        if pref:
            run = p.add_run(pref)
            _set_run_defaults(run)

        if m.group(1):
            run = p.add_run(m.group(1)); _set_run_defaults(run, bold=True)
        elif m.group(2):
            _run_xxxx(p)
        elif m.group(3) or m.group(4):
            run = p.add_run(m.group(0)); _set_run_defaults(run)
        else:
            inner = re.sub(r'^\[\[B\]\]|\[\[/B\]\]$', '', m.group(5))
            run = p.add_run(inner); _set_run_defaults(run, bold=True)

        i = m.end()

# ===================== MEMORIAL RESUMO =====================

def build_memorial_resumo_doc(form, header_logo=None, footer_logo=None, watermark_logo=None):
    doc = preparar_doc(header_logo, footer_logo, watermark_logo)
    _enable_update_fields_on_open(doc)

    nome_fmt = _title_case_name(form.get('nome_emp','') or "")
    end_fmt = _title_keep_preps(form.get('endereco_emp','') or "")
    cid_fmt = _fmt_cidade_slash_uf(form.get('cidade_emp','') or "")
    bai_fmt = _fmt_bairro(form.get('bairro_emp','') or "")

    is_cond = (form.get('tipo_proj_resumo','loteamento') == 'condominio')
    tipo_lbl = "Condomínio fechado de lotes" if is_cond else "Loteamento de acesso controlado"

    # CAPA
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r = p.add_run("MEMORIAL DESCRITIVO"); _set_run_defaults(r, bold=True); r.font.size = Pt(14)

    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r1 = p.add_run(tipo_lbl + " "); _set_run_defaults(r1, bold=True); r1.font.size = Pt(14)
    r2a = p.add_run("“"); _set_run_defaults(r2a, bold=True); r2a.font.size = Pt(14)
    nome_txt = (nome_fmt or "")
    if not nome_txt.strip():
        rr = p.add_run("XXXX")
        _set_run_defaults(rr, bold=True); rr.font.size = Pt(14)
        rr.font.highlight_color = WD_COLOR_INDEX.YELLOW
    else:
        rr = p.add_run(nome_txt)
        _set_run_defaults(rr, bold=True); rr.italic = True; rr.font.size = Pt(14)
    r2b = p.add_run("”"); _set_run_defaults(r2b, bold=True); r2b.font.size = Pt(14)

    # Sumário
    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r = p.add_run("Sumário"); _set_run_defaults(r, bold=True); r.font.size = Pt(14)
    _add_toc(doc)

    idx = 1

    # INTRODUÇÃO
    h_intro = _heading_num(doc, idx, "INTRODUÇÃO")
    h_intro.paragraph_format.page_break_before = True
    idx += 1

    usos_sel = list(form.get('usos_multi', [])) or []
    usos_txt = _join_com_e(usos_sel) or "residencial"
    modo_lbl = "loteamento" if form.get('tipo_proj_resumo','loteamento') == 'loteamento' else "condomínio"

    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _set_run_defaults(p.add_run("O "))
    r = p.add_run(tipo_lbl + " "); _set_run_defaults(r, bold=True)
    p.add_run("“")
    r = p.add_run((nome_fmt or "").strip() or "XXXX")
    _set_run_defaults(r, bold=True); r.italic = True
    if not (form.get('nome_emp','') or "").strip():
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    p.add_run("”")
    _set_run_defaults(p.add_run(
        f" é um empreendimento por unidades autônomas a construir, com finalidade {usos_txt.lower()}, "
        f"configurado como {modo_lbl} de acesso controlado, concebido para oferecer infraestrutura completa, "
        "qualidade de vida, segurança e integração com o entorno."
    ))

    # PROPRIETÁRIO
    _heading_num(doc, idx, "PROPRIETÁRIO/INCORPORADORA"); idx += 1
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _run_xxxx(p); _set_run_defaults(p.add_run(", inscrita no CNPJ sob o nº ")); _run_xxxx(p)
    _set_run_defaults(p.add_run(
        ", proprietária e/ou incorporadora da gleba destinada ao empreendimento."
    ))

    # RESPONSÁVEL TÉCNICO
    _heading_num(doc, idx, "RESPONSÁVEL TÉCNICO PELO PROJETO URBANÍSTICO"); idx += 1
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _set_run_defaults(p.add_run(
        "SOLIDO - DESIGN URBANO LTDA., CNPJ nº 26.887.368/0001-07, CAU-RS 15335-4, "
        "responsável técnico pelo projeto urbanístico."
    ))

    # GLEBA
    _heading_num(doc, idx, "A GLEBA"); idx += 1
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _set_run_defaults(p.add_run("A gleba destinada ao empreendimento possui área total de "))
    if (form.get('area_total_emp','') or '').strip():
        try:
            v = _to_float_br(form.get('area_total_emp',''))
            _set_run_defaults(p.add_run(_fmt_br(v, 2) + "m²"))
            _set_run_defaults(p.add_run(f" ({area_por_extenso(v)}), "))
        except Exception:
            _run_xxxx(p); _set_run_defaults(p.add_run("m², "))
    else:
        _run_xxxx(p); _set_run_defaults(p.add_run("m², "))
    _set_run_defaults(p.add_run(
        f"localizada na {end_fmt or 'XXXX'}, bairro {bai_fmt or 'XXXX'}, município de {cid_fmt or 'XXXX'}, "
        "conforme matrículas e levantamento anexos."
    ))

    # TOPOGRAFIA
    _heading_num(doc, idx, "TOPOGRAFIA"); idx += 1
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    if form.get('topografia','Acentuada') == 'Acentuada':
        _set_run_defaults(p.add_run(
            "A área apresenta topografia acentuada, com declividades variadas, "
            "exigindo soluções específicas de implantação viária e drenagem."
        ))
    else:
        _set_run_defaults(p.add_run(
            "A área apresenta topografia predominantemente plana, "
            "facilitando a implantação do sistema viário e parcelamento."
        ))

    # ZONEAMENTO
    _heading_num(doc, idx, "ZONEAMENTO E ENQUADRAMENTO"); idx += 1
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _set_run_defaults(p.add_run(
        "O empreendimento está enquadrado na legislação urbanística municipal vigente, "
        "respeitando parâmetros de uso, ocupação, taxas, recuos e áreas obrigatórias."
    ))

    # SÍNTESE DO PROJETO
    _heading_num(doc, idx, "SÍNTESE DO PROJETO URBANÍSTICO"); idx += 1
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _set_run_defaults(p.add_run(
        "O projeto contempla sistema viário hierarquizado, lotes com dimensões compatíveis, "
        "áreas verdes e de lazer, áreas institucionais quando exigidas, e, no caso de condomínio, "
        "áreas condominiais de uso comum."
    ))
    if form.get('has_ai', False):
        p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        _set_run_defaults(p.add_run(
            "O empreendimento possui área(s) institucional(is), destinadas a equipamentos públicos "
            "ou comunitários, em atendimento à legislação."
        ))
    if form.get('has_restricao', False):
        p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        _set_run_defaults(p.add_run(
            "Há áreas de restrição definidas para proteção ambiental, faixas não edificantes "
            "ou condicionantes específicas, as quais constam nos desenhos anexos."
        ))

    # CONCLUSÃO
    _heading_num(doc, idx, "CONCLUSÃO"); idx += 1
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _set_run_defaults(p.add_run(
        "O presente Memorial Resumo tem por objetivo apresentar a concepção geral do empreendimento, "
        "comprovando sua adequação legal e fornecendo subsídios para análise técnica."
    ))

    # Data e assinaturas
    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    hoje = datetime.now()
    meses = ["janeiro","fevereiro","março","abril","maio","junho",
             "julho","agosto","setembro","outubro","novembro","dezembro"]
    _set_run_defaults(p.add_run(f"Porto Alegre, {hoje.day} de {meses[hoje.month-1]} de {hoje.year}."))

    _sec_assinaturas_resumo(doc)
    add_footer_left_text(doc, [
        "WWW.SOLIDO.ARQ.BR",
        "Avenida Ipiranga, 6681 – Prédio 99, Sala 906",
        "Porto Alegre – RS Brasil",
        "+ 55 51 99690-7857",
    ], size_pt=10)
    add_page_numbers(doc)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    filename = "URB-PL_XXXX_MEMORIAL RESUMO_RX-VX.docx"
    return bio, filename

def _sec_assinaturas_resumo(doc):
    _add_centered(doc, "ASSINATURAS", bold=True)
    doc.add_paragraph(); doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    r = p.add_run("SOLIDO - DESIGN URBANO LTDA."); _set_run_defaults(r, bold=True)
    p = doc.add_paragraph()
    _set_run_defaults(p.add_run("CAU-RS 15335-4"))

# ===================== SOLICITAÇÃO DE ANÁLISE =====================

def _pt_date(prefixo_cidade="Porto Alegre"):
    meses = ["janeiro","fevereiro","março","abril","maio","junho",
             "julho","agosto","setembro","outubro","novembro","dezembro"]
    hoje = datetime.now()
    return f"{prefixo_cidade}, {hoje.day} de {meses[hoje.month-1]} de {hoje.year}"

def _fmt_matriculas_plural(txt_raw: str):
    txt = (txt_raw or "").strip()
    if not txt:
        return ("matrícula", "XXXX")
    partes = [p.strip() for p in re.split(r'\s*(?:,|;| e )\s*', txt) if p.strip()]
    if len(partes) <= 1:
        return ("matrícula", partes[0] if partes else "XXXX")
    return ("matrículas", ", ".join(partes[:-1]) + " e " + partes[-1])

def build_solicitacao_analise_doc(form, header_logo=None, footer_logo=None, watermark_logo=None):
    doc = preparar_doc(header_logo, footer_logo, watermark_logo)

    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    _set_run_defaults(p.add_run(_pt_date("Porto Alegre")))

    cid_fmt = _fmt_cidade_slash_uf(form.get('cidade_emp','') or "")

    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    _set_run_defaults(p.add_run("À"))
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    _set_run_defaults(p.add_run(f"Prefeitura Municipal de {cid_fmt or 'XXXX'}"))
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    _add_hl(p, "Secretaria competente")

    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    _set_run_defaults(p.add_run("Objeto: Solicitação de análise de Projeto Urbanistico"))

    doc.add_paragraph()

    end_fmt = _title_keep_preps(form.get('endereco_emp',''))
    bai_fmt = _fmt_bairro(form.get('bairro_emp',''))
    tipo_cond_raw = form.get('tipo_proj_resumo','loteamento')
    tipo_cond_txt = "Loteamento de acesso controlado" if tipo_cond_raw == 'loteamento' else "Condomínio fechado de lotes"
    mats_t, mats_fmt = _fmt_matriculas_plural(form.get('matricula_emp',''))

    if (form.get('area_total_emp','') or '').strip():
        try:
            v = _to_float_br(form.get('area_total_emp',''))
            area_txt = _fmt_br(v, 2)
        except Exception:
            area_txt = "XXXX"
    else:
        area_txt = "XXXX"

    par = doc.add_paragraph(); par.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    r = par.add_run("SOLIDO - DESIGN URBANO LTDA., "); _set_run_defaults(r, bold=True)
    _set_run_defaults(par.add_run(
        "na qualidade de responsável técnico pelo projeto urbanístico, "
        "vem, por meio deste, requerer a análise técnica para implantação de um "
    ))
    r = par.add_run(tipo_cond_txt); _set_run_defaults(r, bold=True)
    _set_run_defaults(par.add_run(", inserido em gleba registrada sob "))
    _set_run_defaults(par.add_run(f"{mats_t} nº {mats_fmt}"))
    _set_run_defaults(par.add_run(
        f", com área total de "
    ))
    r = par.add_run(f"{area_txt}m²"); _set_run_defaults(r, bold=True)
    _set_run_defaults(par.add_run(
        f", situada na {end_fmt or 'XXXX'}, bairro {bai_fmt or 'XXXX'}, {cid_fmt or 'XXXX'}."
    ))

    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    _set_run_defaults(p.add_run("Para análise seguem anexos:"))
    for item in [
        "- Projeto urbanístico;",
        "- Memorial descritivo/resumo do empreendimento;",
        "- Ofício/requerimento específico, se houver;",
        "- RRT/ART dos responsáveis."
    ]:
        li = doc.add_paragraph(); li.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        _set_run_defaults(li.add_run(item))

    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    _set_run_defaults(p.add_run("Colocamo-nos à disposição para esclarecimentos e pedimos deferimento."))
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    _set_run_defaults(p.add_run("Atenciosamente,"))

    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    r = p.add_run("SOLIDO - DESIGN URBANO LTDA."); _set_run_defaults(r, bold=True)
    p = doc.add_paragraph()
    _set_run_defaults(p.add_run("CAU-RS 15335-4"))

    add_footer_left_text(doc, [
        "WWW.SOLIDO.ARQ.BR",
        "Avenida Ipiranga, 6681 – Prédio 99, Sala 906",
        "Porto Alegre – RS Brasil",
        "+ 55 51 99690-7857",
    ], size_pt=10)
    add_page_numbers(doc)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    filename = "URB-PL_XXXX_SOLICITAÇÃO DE ANÁLISE_RX-VX.docx"
    return bio, filename

# ===================== UNIFICAÇÃO / DESMEMBRAMENTO =====================

def collect_items_unif_desm(modo, uploaded_files_dict):
    items_unif = None
    items_desm = []

    civil_htmls = [
        (f, d) for f, d in uploaded_files_dict.items()
        if f.lower().endswith(('.html', '.htm')) and 'CIVILREPORT' in f.upper()
    ]
    other_htmls = [
        (f, d) for f, d in uploaded_files_dict.items()
        if f.lower().endswith(('.html', '.htm', '.txt')) and 'CIVILREPORT' not in f.upper()
    ]

    if modo in ('unificacao', 'unif_desm'):
        for fname, data in civil_htmls:
            arr = parse_civilreport_from_html(data)
            for it in arr:
                if is_unificacao_item_name(it.get('name') or ''):
                    items_unif = it

    if modo in ('desmembramento', 'unif_desm'):
        for fname, data in other_htmls:
            try:
                if fname.lower().endswith(('.html', '.htm')):
                    parcels = parse_parcels_from_html(data)
                else:
                    parcels = parse_parcels_from_txt(data)
                for p in parcels:
                    nm = f"GLEBA {p.get('num', 1)}"
                    items_desm.append((nm, {
                        'segments': p.get('segments', []),
                        'area_m2': p.get('area_m2', 0.0),
                        'first_point': p.get('first_point')
                    }))
            except Exception:
                pass

    return items_unif, items_desm

def _texto_ane(largura_m):
    num_sem = f"{_fmt_br(largura_m, 2)}m"
    return (
        f" Existe uma faixa não edificante com largura de {num_sem} "
        f"({extenso_metros(largura_m)}), conforme definido no projeto urbanístico."
    )

def _format_first_point(fp, coord_fmt, zone_num, hemi):
    if not fp:
        return None
    y = round(float(fp["Y"]), 2)
    x = round(float(fp["X"]), 2)
    if coord_fmt == 'utm':
        return f"ponto de coordenadas Y= {_fmt_br(y, 2)}m e X= {_fmt_br(x, 2)}m"
    lat, lon = utm_to_latlon(x, y, zone_num, hemi)
    if coord_fmt == 'dec':
        return f"ponto de coordenadas geográficas {fmt_latlon_decimal(lat, lon)}"
    return f"ponto de coordenadas geográficas {fmt_latlon_dms(lat, lon)}"

def _seg_texto_com_card(seg, dest_coord=None, tipo='line', coord_fmt='utm'):
    az = seg.get("azimuth")
    card = azimuth_to_card8(az)
    az_dms = azimuth_to_dms_int(az)
    dest_txt = ""
    if dest_coord:
        c1, c2 = dest_coord
        if coord_fmt == 'utm':
            dest_txt = f" até o ponto de coordenadas Y= {c2}m e X= {c1}m"
        else:
            dest_txt = f" até o ponto de coordenadas {c2} / {c1}"

    if tipo == 'line':
        lv = round(float(seg["length_m"]), 2)
        length = _fmt_br(lv, 2) + "m"
        return (
            f"daí segue, por reta, sentido {card}, medindo {length} ({extenso_metros(lv)}), "
            f"confrontando ao XXXX com XXXX{dest_txt}, seguindo por um azimute de {az_dms}; "
        )

    clv = round(float(seg["curve_len_m"]), 2)
    rv = round(float(seg["radius_m"]), 2)
    cl = _fmt_br(clv, 2) + "m"
    r = _fmt_br(rv, 2) + "m"
    return (
        f"daí segue, por curva, sentido {card}, medindo {cl} ({extenso_metros(clv)}) e raio de {r} ({extenso_metros(rv)}), "
        f"confrontando ao XXXX com XXXX{dest_txt}, seguindo por um azimute de {az_dms}; "
    )

def build_area_text(item_name, item, tipo_full, empreendimento, endereco, bairro, cidade,
                    ane_enable=False, ane_largura_m=None, coord_fmt='utm', zone_num=22, hemi='S',
                    ident_label_only=False, ident_label_text="Descrição do Imóvel:"):
    nome_norm = _normalize(item_name)
    area = item.get("area_m2") or 0
    area_fmt = _fmt_br(area, 2) + "m²"
    area_ext = area_por_extenso(area)

    cabeca = ""
    if ident_label_only:
        cabeca = f"[[B]]{ident_label_text}[[/B]] Um terreno urbano, irregular, sem benfeitorias, "
    else:
        cabeca = f"[[B]]{nome_norm}[[/B]]: Um terreno urbano, irregular, sem benfeitorias, "

    cabeca += (
        f"localizado na {endereco}, no bairro {bairro}, na cidade de {cidade}, "
        f"com área total de {area_fmt} ({area_ext}). "
    )

    if item.get("first_point"):
        fp_txt = _format_first_point(item["first_point"], coord_fmt, zone_num, hemi)
        if fp_txt:
            cabeca += f"Inicia-se a descrição no {fp_txt}; "

    rows = _propaga_vertices(
        item.get("first_point"),
        item.get("segments", []),
        coord_fmt_str=coord_fmt,
        zone_num=zone_num,
        hemi=hemi
    )

    partes = []
    segs = item.get("segments", []) or []
    for i, seg in enumerate(segs):
        dest = None
        if i < len(rows):
            dest = (rows[i]["COORD_1"], rows[i]["COORD_2"])
        if seg["type"] == "line":
            partes.append(_seg_texto_com_card(seg, dest_coord=dest, tipo='line', coord_fmt=coord_fmt))
        else:
            partes.append(_seg_texto_com_card(seg, dest_coord=dest, tipo='curve', coord_fmt=coord_fmt))

    corpo = "".join(partes)
    if corpo.endswith("; "):
        corpo = corpo[:-2] + ", "

    texto = cabeca + corpo + "chegando ao final da descrição do perímetro."
    if ane_enable and ane_largura_m:
        texto += _texto_ane(ane_largura_m)
    return texto

def build_memorial_text(parcel, quadra, tipo_full, empreendimento, endereco, bairro, cidade,
                        ane_enable=False, ane_largura_m=None, eh_condominio=False,
                        area_tot_priv=0.0, area_tot_cond=0.0, coord_fmt='utm', zone_num=22, hemi='S'):
    num = parcel["num"]
    area = parcel.get("area_m2") or 0

    cabeca = f"LOTE {num} – {quadra}: Um terreno urbano, irregular, sem benfeitorias, "
    cabeca += (
        f"localizado na {endereco}, no bairro {bairro}, na cidade de {cidade}, "
        f"constituído como LOTE {num} da {quadra}, "
    )

    if parcel.get("first_point"):
        fp_txt = _format_first_point(parcel["first_point"], coord_fmt, zone_num, hemi)
        if fp_txt:
            cabeca += f"inicia-se a descrição no {fp_txt}; "

    rows = _propaga_vertices(
        parcel.get("first_point"),
        parcel.get("segments", []),
        coord_fmt_str=coord_fmt,
        zone_num=zone_num,
        hemi=hemi
    )

    partes = []
    segs = parcel.get("segments", []) or []
    for i, seg in enumerate(segs):
        dest = None
        if i < len(rows):
            dest = (rows[i]["COORD_1"], rows[i]["COORD_2"])
        if seg["type"] == "line":
            partes.append(_seg_texto_com_card(seg, dest_coord=dest, tipo='line', coord_fmt=coord_fmt))
        else:
            partes.append(_seg_texto_com_card(seg, dest_coord=dest, tipo='curve', coord_fmt=coord_fmt))

    corpo = "".join(partes)
    if corpo.endswith("; "):
        corpo = corpo[:-2] + ", "

    texto = cabeca + corpo + "chegando ao final da descrição do perímetro."
    if ane_enable and ane_largura_m:
        texto += _texto_ane(ane_largura_m)

    if eh_condominio and area and area_tot_priv:
        fr = area / area_tot_priv
        area_comum = fr * (area_tot_cond or 0.0)
        area_total = area + area_comum
        texto += (
            f" Possui área real privativa de {_fmt_br(area, 2)}m², "
            f"área de uso comum de {_fmt_br(area_comum, 2)}m², "
            f"área real total de {_fmt_br(area_total, 2)}m², "
            f"correspondendo-lhe a fração ideal de {fr:.7f}."
        )

    return texto

def _sec_assinaturas_simples(doc):
    _add_centered(doc, "ASSINATURAS", bold=True)
    doc.add_paragraph(); doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    r = p.add_run("Responsável técnico"); _set_run_defaults(r, bold=True)
    p = doc.add_paragraph()
    r = p.add_run("SOLIDO - DESIGN URBANO LTDA."); _set_run_defaults(r)
    p = doc.add_paragraph()
    r = p.add_run("CAU-RS 15335-4"); _set_run_defaults(r)

def build_unif_desm_doc(modo, form, uploaded_files_dict,
                        header_logo=None, footer_logo=None, watermark_logo=None):
    unif_item, desm_items = collect_items_unif_desm(modo, uploaded_files_dict)
    doc = preparar_doc(header_logo, footer_logo, watermark_logo)

    nome_fmt = _title_case_name(form.get('nome_emp','') or "")
    end_fmt = _title_keep_preps(form.get('endereco_emp','') or "")
    cid_fmt = _fmt_cidade_slash_uf(form.get('cidade_emp','') or "")
    bai_fmt = _fmt_bairro(form.get('bairro_emp','') or "")
    zone_num, hemi = _auto_zone_from_city(form.get('cidade_emp','') or '')
    coord_fmt = form.get('coord_fmt','utm')

    heading(doc, "MEMORIAL DESCRITIVO DE UNIFICAÇÃO E/OU DESMEMBRAMENTO")

    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    mats_raw = (form.get('matricula_emp','') or '').strip()
    _set_run_defaults(p.add_run(
        "O presente memorial tem por finalidade descrever a unificação e/ou desmembramento "
        f"de área(s) de terras localizada(s) na {end_fmt or 'XXXX'}, bairro {bai_fmt or 'XXXX'}, "
        f"município de {cid_fmt or 'XXXX'}, objeto das matrículas {mats_raw or 'XXXX'}."
    ))

    if unif_item:
        heading(doc, "UNIFICAÇÃO")
        texto = build_area_text(
            unif_item.get('name',"UNIFICAÇÃO"),
            unif_item,
            "",
            nome_fmt or "XXXX",
            end_fmt or "XXXX",
            bai_fmt or "XXXX",
            cid_fmt or "XXXX",
            coord_fmt=coord_fmt,
            zone_num=zone_num,
            hemi=hemi,
            ident_label_only=True
        )
        adicionar_texto_formatado(doc, texto)

    if desm_items:
        heading(doc, "DESMEMBRAMENTO")
        def _first_int_or_big(s):
            m = re.search(r'(\d+)', _normalize(s))
            return int(m.group(1)) if m else 10**9
        desm_sorted = sorted(desm_items, key=lambda kv: (_first_int_or_big(kv[0]), _normalize(kv[0])))
        for nm, it in desm_sorted:
            texto = build_area_text(
                nm,
                it,
                "",
                nome_fmt or "XXXX",
                end_fmt or "XXXX",
                bai_fmt or "XXXX",
                cid_fmt or "XXXX",
                coord_fmt=coord_fmt,
                zone_num=zone_num,
                hemi=hemi,
                ident_label_only=True
            )
            adicionar_texto_formatado(doc, texto)

    _sec_assinaturas_simples(doc)
    add_footer_left_text(doc, [
        "WWW.SOLIDO.ARQ.BR",
        "Avenida Ipiranga, 6681 – Prédio 99, Sala 906",
        "Porto Alegre – RS Brasil",
        "+ 55 51 99690-7857",
    ], size_pt=10)
    add_page_numbers(doc)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    filename = "URB-PL_XXXX_UNIF_DESM_RX-VX.docx"
    return bio, filename

# ===================== CONDOMÍNIO / LOTEAMENTO =====================

def generate_excel_fracao_ideal(dados_quadro):
    if not dados_quadro:
        raise ValueError("Sem dados de fração ideal.")
    df = pd.DataFrame(dados_quadro)
    df['__quad_key__'] = df['Quadra'].map(lambda q: quadra_label_sort_key(f"QUADRA {q}"))
    df['__lote_key__'] = df['Lote'].map(_lote_num)
    df = df.sort_values(['__quad_key__', '__lote_key__']).drop(columns=['__quad_key__', '__lote_key__'])

    bio = BytesIO()
    df.to_excel(bio, index=False)
    bio.seek(0)

    wb = load_workbook(bio)
    ws = wb.active
    font_header = Font(name='Calibri', size=12, bold=True)
    font_cell = Font(name='Calibri', size=12)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin = Side(border_style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for r in ws.iter_rows():
        for c in r:
            c.alignment = center
            c.border = border
            c.font = font_header if c.row == 1 else font_cell

    for col in ws.columns:
        maxlen = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = max(12, maxlen + 2)

    bio2 = BytesIO()
    wb.save(bio2)
    bio2.seek(0)
    filename = "URB-PL_XXXX_QUADRO_FRACAO_IDEAL_RX-VX.xlsx"
    return bio2, filename

def generate_excel_unif_desm(modo, cidade, coord_fmt, uploaded_files_dict):
    unif_item, desm_items = collect_items_unif_desm(modo, uploaded_files_dict)
    if modo == 'unificacao' and not unif_item:
        raise ValueError("Nenhuma área de unificação encontrada.")
    if modo == 'desmembramento' and not desm_items:
        raise ValueError("Nenhuma gleba de desmembramento encontrada.")
    if modo == 'unif_desm' and not (unif_item or desm_items):
        raise ValueError("Nenhuma área para unificação/desmembramento encontrada.")

    zone_num, hemi = _auto_zone_from_city(cidade or '')

    wb = Workbook()
    wb.remove(wb.active)

    def nova_aba(nome):
        ws = wb.create_sheet(title=nome)
        for idx in range(1, 9):
            ws.column_dimensions[get_column_letter(idx)].width = 14
        ws.column_dimensions['C'].width = 17
        ws.column_dimensions['D'].width = 17
        ws.column_dimensions['F'].width = 17
        ws.column_dimensions['H'].width = 17
        return ws

    def headers(ws, row):
        if coord_fmt == 'utm':
            hC, hD = "COORD. X", "COORD. Y"
        else:
            hC, hD = "LATITUDE/LONGITUDE", ""
        headers = ["DE", "PARA", hC, hD, "AZIMUTE", "DISTANCIA (m)", "RAIO (m)", "CONFRONTANTE"]
        font_header = Font(name='Calibri', size=12, bold=True)
        center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        thin = Side(border_style='thin', color='000000')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=c, value=h)
            cell.font = font_header
            cell.alignment = center
            cell.border = border

    def append_block(ws, titulo, item, start_row):
        font_header = Font(name='Calibri', size=12, bold=True)
        font_cell = Font(name='Calibri', size=12)
        center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        thin = Side(border_style='thin', color='000000')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        yellow = PatternFill('solid', fgColor='FFF59D')

        max_col = 8
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=max_col)
        tcell = ws.cell(row=start_row, column=1, value=titulo)
        tcell.font = font_header
        tcell.alignment = center
        for c in range(1, max_col + 1):
            ws.cell(row=start_row, column=c).border = border

        header_row = start_row + 1
        headers(ws, header_row)

        rows = _propaga_vertices(
            item.get('first_point'),
            item.get('segments', []),
            coord_fmt_str=coord_fmt,
            zone_num=zone_num,
            hemi=hemi
        )
        r = header_row + 1
        for row in rows:
            vals = [
                row.get("DE",""), row.get("PARA",""),
                row.get("COORD_1",""), row.get("COORD_2",""),
                row.get("AZIMUTE",""), row.get("DISTANCIA (m)",""),
                row.get("RAIO (m)",""), row.get("CONFRONTANTE","")
            ]
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = font_cell
                cell.alignment = center
                cell.border = border
            if re.match(r'^P\d+$', str(vals[0])):
                ws.cell(row=r, column=1).fill = yellow
            if re.match(r'^P\d+$', str(vals[1])):
                ws.cell(row=r, column=2).fill = yellow
            r += 1
        return r + 1

    if modo in ('unificacao', 'unif_desm') and unif_item:
        ws_u = nova_aba("UNIFICACAO")
        r = 1
        area = unif_item.get("area_m2") or 0
        titulo = f"UNIFICAÇÃO - ÁREA: {_fmt_br(area, 2)}m²"
        r = append_block(ws_u, titulo, unif_item, r)

    if modo in ('desmembramento', 'unif_desm') and desm_items:
        ws_d = nova_aba("DESMEMBRAMENTO")
        r = 1
        def _first_int_or_big(s):
            m = re.search(r'(\d+)', _normalize(s))
            return int(m.group(1)) if m else 10**9
        desm_sorted = sorted(desm_items, key=lambda kv: (_first_int_or_big(kv[0]), _normalize(kv[0])))
        for nm, it in desm_sorted:
            area = it.get("area_m2") or 0
            titulo = f"{nm} - ÁREA: {_fmt_br(area, 2)}m²"
            r = append_block(ws_d, titulo, it, r)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    filename = "URB-PL_XXXX_VERTICES_UNIF_DESM_RX-VX.xlsx"
    return bio, filename

def build_condominio_loteamento_doc(modo, form, uploaded_files_dict,
                                    header_logo=None, footer_logo=None, watermark_logo=None):
    nome_fmt = _title_case_name(form.get('nome_emp','') or "")
    end_fmt = _title_keep_preps(form.get('endereco_emp','') or "")
    cid_fmt = _fmt_cidade_slash_uf(form.get('cidade_emp','') or "")
    bai_fmt = _fmt_bairro(form.get('bairro_emp','') or "")
    coord_fmt = form.get('coord_fmt','utm')

    lot_files = [
        (f, d) for f, d in uploaded_files_dict.items()
        if f.lower().endswith(('.html', '.htm', '.txt')) and 'CIVILREPORT' not in f.upper()
    ]
    civil_files = [
        (f, d) for f, d in uploaded_files_dict.items()
        if f.lower().endswith(('.html', '.htm')) and 'CIVILREPORT' in f.upper()
    ]

    file_parcels = []
    all_parcels = []
    for fname, data in lot_files:
        quadra = infer_quadra_from_filename(fname)
        if fname.lower().endswith(('.html', '.htm')):
            parcels = parse_parcels_from_html(data)
        else:
            parcels = parse_parcels_from_txt(data)
        parcels.sort(key=lambda p: p.get('num', 0))
        file_parcels.append((quadra, parcels))
        all_parcels.extend(parcels)

    file_parcels.sort(key=lambda qp: quadra_label_sort_key(qp[0]))
    for i, (quadra, parcels) in enumerate(file_parcels):
        parcels.sort(key=lambda p: int(p.get('num', 0)))
        file_parcels[i] = (quadra, parcels)

    tipo_full = "Condomínio Fechado de Lotes" if modo == 'condominio' else "Loteamento de Acesso Controlado"
    eh_condominio = (modo == 'condominio')

    area_tot_priv = 0.0
    area_tot_cond = 0.0
    if eh_condominio:
        if (form.get('area_tot_priv_emp','') or '').strip():
            try:
                area_tot_priv = _to_float_br(form.get('area_tot_priv_emp',''))
            except Exception:
                pass
        if (form.get('area_tot_cond_emp','') or '').strip():
            try:
                area_tot_cond = _to_float_br(form.get('area_tot_cond_emp',''))
            except Exception:
                pass

    ane_enable = (form.get('ane_drop') == 'Sim')
    ane_largura_m = None
    if ane_enable and (form.get('ane_largura','') or '').strip():
        try:
            ane_largura_m = _to_float_br(form.get('ane_largura',''))
        except Exception:
            ane_largura_m = None

    civil_items = []
    for fname, data in civil_files:
        civil_items.extend(parse_civilreport_from_html(data))

    grouped = {k: [] for k in [
        'remanescente', 'institucional', 'reserva_tecnica', 'app',
        'verde', 'verde_preservacao', 'viario', 'condominial',
        'quadras', 'outros'
    ]}
    for it in civil_items:
        cat, title = classify_civil_item(it['name'])
        grouped[cat].append((title, it))

    def _num_key(nm):
        m = re.search(r'(\d+)', _normalize(nm))
        return int(m.group(1)) if m else 10**9

    for cat in grouped:
        if cat == 'viario':
            grouped[cat].sort(key=lambda x: x[1]['name'])
        else:
            grouped[cat].sort(key=lambda x: (_num_key(x[1]['name']), _normalize(x[1]['name'])))

    doc = preparar_doc(header_logo, footer_logo, watermark_logo)
    heading(doc, "MEMORIAL DESCRITIVO")

    # texto inicial
    area_tot_fmt = area_tot_ext = ha_txt = perim_fmt = perim_ext = ""
    if (form.get('area_total_emp','') or '').strip():
        v = _to_float_br(form.get('area_total_emp',''))
        area_tot_fmt = _fmt_br(v, 2) + "m²"
        area_tot_ext = area_por_extenso(v)
        ha_txt = _fmt_br(hectares_from_m2(v), 2) + "ha"
    if (form.get('perimetro_emp','') or '').strip():
        pval = _to_float_br(form.get('perimetro_emp',''))
        perim_fmt = _fmt_br(pval, 2)
        perim_ext = extenso_metros(pval)

    zone_num, hemi = _auto_zone_from_city(form.get('cidade_emp','') or '')
    mc_w = _utm_mc_from_zone(zone_num)

    p1 = doc.add_paragraph(); p1.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _set_run_defaults(p1.add_run(
        "O presente memorial tem por finalidade descrever o parcelamento de solo de acordo com o projeto denominado "
    ))
    r = p1.add_run(tipo_full + " "); _set_run_defaults(r, bold=True)
    p1.add_run("“")
    rnome = p1.add_run(nome_fmt or "XXXX"); _set_run_defaults(rnome, bold=True); rnome.italic = True
    if not (form.get('nome_emp','') or '').strip():
        rnome.font.highlight_color = WD_COLOR_INDEX.YELLOW
    p1.add_run("”")
    _set_run_defaults(p1.add_run(
        f", em gleba situada frente à {end_fmt or 'XXXX'}, bairro {bai_fmt or 'XXXX'}, município de {cid_fmt or 'XXXX'}, "
        f"com área superficial de {area_tot_fmt or 'XXXXm²'}"
    ))
    if area_tot_ext:
        _set_run_defaults(p1.add_run(f" ({area_tot_ext}) - {ha_txt}"))
    if perim_fmt:
        _set_run_defaults(p1.add_run(
            f" e perímetro de {perim_fmt}m ({perim_ext}), "
        ))
    else:
        _set_run_defaults(p1.add_run(", "))
    _set_run_defaults(p1.add_run(
        f"conforme matrículas e plantas anexas."
    ))

    p2 = doc.add_paragraph(); p2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    if coord_fmt == 'utm':
        _set_run_defaults(p2.add_run(
            f"As coordenadas estão referidas ao Sistema Geodésico Brasileiro, Datum SIRGAS 2000, MC {mc_w}W, "
            "coordenadas plano retangulares UTM."
        ))
    elif coord_fmt == 'dec':
        _set_run_defaults(p2.add_run(
            "As coordenadas estão referidas ao Sistema Geodésico Brasileiro, Datum SIRGAS 2000, em graus decimais."
        ))
    else:
        _set_run_defaults(p2.add_run(
            "As coordenadas estão referidas ao Sistema Geodésico Brasileiro, Datum SIRGAS 2000, "
            "em graus, minutos e segundos."
        ))

    # Áreas especiais
    order = ['remanescente','institucional','reserva_tecnica','app',
             'verde','verde_preservacao','viario','condominial']
    for cat in order:
        itens = grouped.get(cat) or []
        if not itens:
            continue
        title = itens[0][0]
        heading(doc, title)
        for _, it in itens:
            texto = build_area_text(
                it['name'],
                it,
                tipo_full,
                nome_fmt or "XXXX",
                end_fmt or "XXXX",
                bai_fmt or "XXXX",
                cid_fmt or "XXXX",
                coord_fmt=coord_fmt,
                zone_num=zone_num,
                hemi=hemi
            )
            adicionar_texto_formatado(doc, texto)

    # Quadras (placeholder)
    if grouped.get('quadras'):
        heading(doc, "DESCRIÇÃO DE QUADRAS")
        for _, it in grouped['quadras']:
            texto = build_area_text(
                it['name'],
                it,
                tipo_full,
                nome_fmt or "XXXX",
                end_fmt or "XXXX",
                bai_fmt or "XXXX",
                cid_fmt or "XXXX",
                coord_fmt=coord_fmt,
                zone_num=zone_num,
                hemi=hemi
            )
            adicionar_texto_formatado(doc, texto)

    # Lotes
    heading(doc, "DESCRIÇÃO DE LOTES")

    dados_quadro = []
    for quadra, parcels in file_parcels:
        for parcel in parcels:
            texto_lote = build_memorial_text(
                parcel,
                quadra,
                tipo_full,
                nome_fmt or "XXXX",
                end_fmt or "XXXX",
                bai_fmt or "XXXX",
                cid_fmt or "XXXX",
                ane_enable=ane_enable,
                ane_largura_m=ane_largura_m,
                eh_condominio=eh_condominio,
                area_tot_priv=area_tot_priv,
                area_tot_cond=area_tot_cond,
                coord_fmt=coord_fmt,
                zone_num=zone_num,
                hemi=hemi
            )
            adicionar_texto_formatado(doc, texto_lote)

            if eh_condominio and parcel.get("area_m2") and area_tot_priv:
                area_priv = float(parcel['area_m2'])
                fr = area_priv / area_tot_priv
                area_comum = fr * (area_tot_cond or 0.0)
                area_total = area_priv + area_comum
                dados_quadro.append({
                    'Lote': str(parcel['num']),
                    'Quadra': quadra.replace("QUADRA","").strip(),
                    'Área Privativa (m²)': _fmt_br(area_priv, 2),
                    'Área Uso Comum (m²)': _fmt_br(area_comum, 2),
                    'Área Real Total (m²)': _fmt_br(area_total, 2),
                    'Fração Ideal': f"{fr:.7f}"
                })

    if eh_condominio and dados_quadro:
        # tabela dentro do doc
        dados_quadro.sort(key=lambda row: (quadra_label_sort_key("QUADRA "+row['Quadra']), _lote_num(row['Lote'])))
        tbl = doc.add_table(rows=1, cols=6)
        tbl.style = 'Table Grid'
        hdr = tbl.rows[0].cells
        headers = ["Lote","Quadra","Área Priv. (m²)","Área Uso Comum (m²)","Área Real Total (m²)","Fração Ideal"]
        for i, h in enumerate(headers):
            hdr[i].text = h
            for r in hdr[i].paragraphs:
                r.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                for run in r.runs:
                    _set_run_defaults(run, bold=True)
        for row in dados_quadro:
            cells = tbl.add_row().cells
            vals = [
                row['Lote'], row['Quadra'], row['Área Privativa (m²)'],
                row['Área Uso Comum (m²)'], row['Área Real Total (m²)'],
                row['Fração Ideal']
            ]
            for i, v in enumerate(vals):
                cells[i].text = v
                for par in cells[i].paragraphs:
                    par.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    for run in par.runs:
                        _set_run_defaults(run)

    _sec_assinaturas_simples(doc)
    add_footer_left_text(doc, [
        "WWW.SOLIDO.ARQ.BR",
        "Avenida Ipiranga, 6681 – Prédio 99, Sala 906",
        "Porto Alegre – RS Brasil",
        "+ 55 51 99690-7857",
    ], size_pt=10)
    add_page_numbers(doc)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    filename = "URB-PL_XXXX_MEMORIAL_RX-VX.docx"
    return bio, filename, (dados_quadro if eh_condominio else None)

# ===================== ORQUESTRADORES PÚBLICOS =====================

def generate_docx(tipo, form, uploaded_files_dict,
                  header_logo=None, footer_logo=None, watermark_logo=None):
    """
    Retorna (bytes_io, filename, meta)
    meta:
      - para Memorial Condomínio: quadro fração ideal (lista de dicts)
      - demais: None
    """
    if tipo == 'Memorial Resumo':
        bio, fname = build_memorial_resumo_doc(form, header_logo, footer_logo, watermark_logo)
        return bio, fname, None

    if tipo == 'Solicitação de Análise':
        bio, fname = build_solicitacao_analise_doc(form, header_logo, footer_logo, watermark_logo)
        return bio, fname, None

    if tipo in ('Memorial Unificação', 'Memorial Desmembramento', 'Memorial Unificação e Desmembramento'):
        modo = {
            'Memorial Unificação': 'unificacao',
            'Memorial Desmembramento': 'desmembramento',
            'Memorial Unificação e Desmembramento': 'unif_desm'
        }[tipo]
        bio, fname = build_unif_desm_doc(modo, form, uploaded_files_dict,
                                         header_logo, footer_logo, watermark_logo)
        return bio, fname, None

    if tipo in ('Memorial Condomínio', 'Memorial Loteamento'):
        modo = 'condominio' if tipo == 'Memorial Condomínio' else 'loteamento'
        bio, fname, quadro = build_condominio_loteamento_doc(
            modo, form, uploaded_files_dict,
            header_logo, footer_logo, watermark_logo
        )
        return bio, fname, quadro

    raise ValueError("Tipo inválido.")

def generate_excel(tipo, form, uploaded_files_dict, quadro_frac_ideal=None):
    if tipo == 'Memorial Condomínio':
        if not quadro_frac_ideal:
            raise ValueError("Sem dados de fração ideal (gere o DOCX do condomínio primeiro).")
        return generate_excel_fracao_ideal(quadro_frac_ideal)

    if tipo in ('Memorial Unificação', 'Memorial Desmembramento', 'Memorial Unificação e Desmembramento'):
        modo = {
            'Memorial Unificação': 'unificacao',
            'Memorial Desmembramento': 'desmembramento',
            'Memorial Unificação e Desmembramento': 'unif_desm'
        }[tipo]
        coord_fmt = form.get('coord_fmt', 'utm')
        cidade = form.get('cidade_emp', '')
        return generate_excel_unif_desm(modo, cidade, coord_fmt, uploaded_files_dict)

    raise ValueError("Este tipo não possui Excel dedicado.")
