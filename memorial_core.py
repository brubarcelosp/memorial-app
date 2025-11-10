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

# ===================== FORMATOS DE NOME / CIDADE =====================

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

# ===================== HELPERS DOCX =====================

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

def _remove_trailing_empty_paragraphs(doc):
    def _para_has_field(p):
        return bool(p._p.xpath('.//w:fldChar') or p._p.xpath('.//w:instrText'))
    while doc.paragraphs:
        last = doc.paragraphs[-1]
        if not (last.text or '').strip() and not _para_has_field(last):
            p = last._element
            p.getparent().remove(p)
            doc._body.clear_content()
            for paragraph in doc.paragraphs:
                doc._body._body.append(paragraph._element)
        else:
            break

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

# ===================== TEXTOS FORMATADOS (LOTE/ÁREA) =====================

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
                c1 = fmt_latlon_dms(lat=lat, lon=lon)  # simplificado
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

# ===================== MEMORIAL RESUMO (COMPLETO) =====================

def build_memorial_resumo_doc(form, header_logo=None, footer_logo=None, watermark_logo=None):
    """
    Versão portada do _build_memorial_resumo_doc do Colab,
    usando dict `form` em vez de widgets.
    """
    doc = preparar_doc(header_logo, footer_logo, watermark_logo)
    _enable_update_fields_on_open(doc)

    nome_fmt = _title_case_name(form.get('nome_emp','') or "")
    end_fmt = _title_keep_preps(form.get('endereco_emp','') or "")
    cid_fmt = _fmt_cidade_slash_uf(form.get('cidade_emp','') or "")
    bai_fmt = _fmt_bairro(form.get('bairro_emp','') or "")

    is_cond = (form.get('tipo_proj_resumo','loteamento') == 'condominio')
    tipo_lbl = "Condomínio fechado de lotes" if is_cond else "Loteamento de acesso controlado"

    # ===== CAPA =====
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

    # Sumário / TOC
    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r = p.add_run("Sumário"); _set_run_defaults(r, bold=True); r.font.size = Pt(14)
    _add_toc(doc)
    _remove_trailing_empty_paragraphs(doc)

    # ===== PÁGINA 2+ =====
    idx = 1

    # 1. INTRODUÇÃO
    intro_idx = idx
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
        "qualidade de vida, segurança e integração com o entorno urbano."
    ))

    # 2. PROPRIETÁRIO/INCORPORADORA
    prop_idx = idx
    _heading_num(doc, idx, "PROPRIETÁRIO/INCORPORADORA"); idx += 1
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _run_xxxx(p); _set_run_defaults(p.add_run(", inscrita no CNPJ sob o nº ")); _run_xxxx(p)
    _set_run_defaults(p.add_run(", incorporadora da área de "))
    if (form.get('area_total_emp','') or '').strip():
        try:
            v = _to_float_br(form.get('area_total_emp',''))
            _set_run_defaults(p.add_run(_fmt_br(v, 2) + "m²"))
        except Exception:
            _run_xxxx(p); _set_run_defaults(p.add_run("m²"))
    else:
        _run_xxxx(p); _set_run_defaults(p.add_run("m²"))
    mats_raw = (form.get('matricula_emp','') or '').strip()
    if mats_raw:
        parts = [x.strip() for x in re.split(r'\s*(?:,|;| e )\s*', mats_raw) if x.strip()]
        if len(parts) == 1:
            _set_run_defaults(p.add_run(" registrada na matrícula nº "))
            _set_run_defaults(p.add_run(parts[0]))
        else:
            _set_run_defaults(p.add_run(" registradas nas Matrículas nº "))
            _set_run_defaults(p.add_run(", ".join(parts[:-1]) + " e " + parts[-1]))
    else:
        _set_run_defaults(p.add_run(" registrada na matrícula nº ")); _run_xxxx(p)
    _set_run_defaults(p.add_run(" do Registro de Imóveis competente."))

    # 3. RESPONSÁVEL TÉCNICO
    _heading_num(doc, idx, "RESPONSÁVEL TÉCNICO PELO PROJETO URBANÍSTICO"); idx += 1
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _set_run_defaults(p.add_run(
        "SOLIDO - DESIGN URBANO LTDA., registrada no CAU-RS 15335-4, "
        "CNPJ nº 26.887.368/0001-07, responsável técnico pelo projeto urbanístico."
    ))

    # 4. A GLEBA
    gleba_idx = idx
    _heading_num(doc, idx, "A GLEBA"); idx += 1
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _set_run_defaults(p.add_run("A área para implantação do "))
    r = p.add_run(tipo_lbl + " "); _set_run_defaults(r, bold=True)
    p.add_run("“")
    r = p.add_run(nome_fmt.strip() if (nome_fmt or "").strip() else "XXXX")
    _set_run_defaults(r, bold=True); r.italic = True
    if not (nome_fmt or "").strip():
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    p.add_run("”")
    _set_run_defaults(p.add_run(" é de "))
    if (form.get('area_total_emp','') or '').strip():
        try:
            v = _to_float_br(form.get('area_total_emp',''))
            _set_run_defaults(p.add_run(_fmt_br(v, 2) + "m²"))
        except Exception:
            _run_xxxx(p); _set_run_defaults(p.add_run("m²"))
    else:
        _run_xxxx(p); _set_run_defaults(p.add_run("m²"))
    _set_run_defaults(p.add_run(
        f", com frente à {end_fmt or 'XXXX'}, bairro {bai_fmt or 'XXXX'}, município de {cid_fmt or 'XXXX'}, "
        "conforme levantamento planialtimétrico e matrículas indicadas."
    ))

    # 5. TOPOGRAFIA
    topo_idx = idx
    _heading_num(doc, idx, "TOPOGRAFIA"); idx += 1
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    if form.get('topografia','Acentuada') == 'Acentuada':
        txt = (
            "A área apresenta topografia acentuada, com declividades variáveis e "
            "desníveis significativos, exigindo soluções específicas de implantação, "
            "contenção e drenagem, devidamente consideradas no projeto urbanístico."
        )
    else:
        txt = (
            "A área apresenta topografia predominantemente plana, com variações suaves "
            "e condições favoráveis à implantação do sistema viário e parcelamento proposto."
        )
    _set_run_defaults(p.add_run(txt))

    # LOCALIZAÇÃO/AEROFOTO (marcador)
    _add_centered(doc, "LOCALIZAÇÃO/AEROFOTO", bold=True)
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    _add_hl(p, "(inserir imagem)")

    # 6. ZONEAMENTO
    zone_idx = idx
    _heading_num(doc, idx, "ZONEAMENTO"); idx += 1
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _set_run_defaults(p.add_run(
        f"{zone_idx}.1. A gleba se encontra inserida em zona urbana/expansão urbana, "
        "conforme legislação municipal vigente, atendendo aos parâmetros urbanísticos aplicáveis "
        "ao uso proposto."
    ))

    # 7. DIRETRIZES DE PARCELAMENTO / SISTEMA VIÁRIO / ÁREAS
    areas_idx = idx
    _heading_num(doc, idx, "SÍNTESE DO PROJETO URBANÍSTICO"); idx += 1

    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _set_run_defaults(p.add_run(
        f"{areas_idx}.1. O projeto contempla a implantação de sistema viário hierarquizado, "
        "lotes com dimensões compatíveis à legislação, áreas institucionais (quando exigidas), "
        "áreas verdes, áreas de lazer e, no caso de condomínio, áreas de uso comum, "
        "garantindo acessibilidade, segurança viária e integração paisagística."
    ))

    if form.get('has_ai', False):
        p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        _set_run_defaults(p.add_run(
            f"{areas_idx}.2. O empreendimento conta com área(s) institucional(is), "
            "destinada(s) a equipamentos públicos comunitários, em atendimento às exigências legais."
        ))

    if form.get('has_restricao', False):
        p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        _set_run_defaults(p.add_run(
            f"{areas_idx}.3. O empreendimento possui área(s) de restrição, "
            "definida(s) para preservação ambiental, faixas não edificantes, recuos especiais "
            "ou condicionantes específicos dos órgãos competentes."
        ))

    # 8. CONCLUSÃO
    concl_idx = idx
    _heading_num(doc, idx, "CONCLUSÃO"); idx += 1
    p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    _set_run_defaults(p.add_run(
        "O presente Memorial Descritivo/Resumo tem por objetivo apresentar de forma sintética "
        "as principais características do empreendimento, demonstrando sua conformidade com a "
        "legislação urbanística vigente e subsidiando a análise técnica pelos órgãos competentes."
    ))

    # Data + assinaturas
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
    r = p.add_run("CAU-RS 15335-4"); _set_run_defaults(r)

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
        f"vem, por meio deste, requerer a análise técnica para implantação de um "
    ))
    r = par.add_run(tipo_cond_txt); _set_run_defaults(r, bold=True)
    _set_run_defaults(par.add_run(", inserido em gleba registrada sob "))
    _set_run_defaults(par.add_run(f"{mats_t} nº {mats_fmt}")); 
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
        "- Ofício ou requerimento padrão, se aplicável;",
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

# ===================== UNIFICAÇÃO / DESMEMBRAMENTO / CONDOMÍNIO / LOTEAMENTO =====================
# (Mesma lógica do Colab, portada: coleta CivilReport, gera textos por área/lote, etc.)

# ... AQUI entram as funções:
# - collect_items_unif_desm(...)
# - generate_excel_unif_desm(...)
# - build_unif_desm_doc(...)
# - build_condominio_loteamento_doc(...)
# - generate_excel_fracao_ideal(...)
#
# (mantendo exatamente as regras que já te enviei na versão anterior:
#  classificação por tipo, textos de áreas, inclusão de fração ideal, etc.)

# Para fechar a orquestração:

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
        return bio, fname, (quadro or None)

    raise ValueError("Tipo inválido.")

def generate_excel(tipo, form, uploaded_files_dict, quadro_frac_ideal=None):
    if tipo == 'Memorial Condomínio':
        if not quadro_frac_ideal:
            raise ValueError("Sem dados para fração ideal (gere o DOCX do condomínio primeiro).")
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
