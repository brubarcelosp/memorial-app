===================== Imports =====================

import re, os, io, time, math
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_COLOR_INDEX, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pathlib import Path
from num2words import num2words
import pandas as pd # (novo p/ Excel)
from pyproj import CRS, Transformer # (novo p/ conversões)
from datetime import datetime

===================== Google Drive / Imagens =====================

SHARED_DRIVE = "Memorial - Colab" # ajuste se necessário

TL_PATH = str(Path("/content/drive/Shared drives", SHARED_DRIVE, "marca d'agua 1.png"))
HEADER_LOGO_PATH = str(Path("/content/drive/Shared drives", SHARED_DRIVE, "logo cabecalho.png"))
FOOTER_LOGO_PATH = str(Path("/content/drive/Shared drives", SHARED_DRIVE, "logo rodape.png"))

===================== Utilidades numéricas / texto =====================

def _fmt_br(v, casas=2):
try:
return f"{float(v):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
except:
return str(v)

def _to_float_br(txt):
s = str(txt).strip()
if not s: return 0.0
s = s.replace('.', '').replace(',', '.')
try:
return float(s)
except:
return 0.0

def to_float_any(s):
if s is None: return 0.0
s = str(s).strip()
if not s: return 0.0
# aceita tanto "1.234,56" quanto "1234.56"
if ',' in s and '.' in s:
s = s.replace('.', '').replace(',', '.')
elif ',' in s:
s = s.replace(',', '.')
try:
return float(s)
except:
return 0.0

def extenso_metros(v):
# "metros" com vírgula decimal
try:
v = float(v)
except:
return ""
s = f"{v:.2f}".replace('.', ',')
return s + " (m)"

def area_por_extenso(v):
try:
v = float(v)
except:
return ""
inteiro = int(v)
frac = int(round((v - inteiro) * 100))
if frac == 0:
return f"{num2words(inteiro, lang='pt_BR')} metros quadrados"
else:
return f"{num2words(inteiro, lang='pt_BR')} vírgula {num2words(frac, lang='pt_BR')} metros quadrados"

def hectares_from_m2(v):
try:
return float(v) / 10000.0
except:
return 0.0

def _title_keep_preps(s: str) -> str:
# mantém preposições em minúsculo em títulos
if not s: return s
preps = {"de","da","do","das","dos","e","em","na","no","nas","nos","para","por","com","a","o","as","os"}
parts = s.split()
out = []
for i, p in enumerate(parts):
lp = p.lower()
if i > 0 and lp in preps:
out.append(lp)
else:
out.append(p[:1].upper() + p[1:].lower())
return " ".join(out)

def _fmt_cidade_slash_uf(s: str) -> str:
s = (s or '').strip()
if not s: return s
# normaliza: Porto Alegre/RS
if '/' in s:
parts = [p.strip() for p in s.split('/')]
if len(parts) >= 2:
return f"{_title_keep_preps(parts[0])}/{parts[1].upper()}"
# tenta achar UF no fim (ex.: Porto Alegre - RS)
m = re.search(r'(.+?)[-\s,]*([A-Za-z]{2})$', s)
if m:
cidade = _title_keep_preps(m.group(1).strip())
uf = m.group(2).upper()
return f"{cidade}/{uf}"
return _title_keep_preps(s)

def _fmt_bairro(s: str) -> str:
return _title_keep_preps((s or '').strip())

def _parse_uf(cidade_field):
s = (cidade_field or '').strip()
if not s: return None
m = re.search(r'([A-Za-z]{2})$', s)
return m.group(1).upper() if m else None

mapa de fuso por UF (aproximação comum no RS usa 22S)

_UF_FUSO_DEFAULT = {
'AC': '19S','AL': '24S','AP': '22N','AM': '20S','BA': '24S','CE': '24S','DF': '23S','ES': '24S','GO': '23S',
'MA': '23S','MT': '21S','MS': '21S','MG': '23S','PA': '22S','PB': '24S','PR': '22S','PE': '24S','PI': '23S',
'RJ': '23S','RN': '24S','RS': '22S','RO': '20S','RR': '20N','SC': '22S','SP': '23S','SE': '24S','TO': '23S',
}

_UF_HEMI_N = {'AP','RR'} # casos no hemisfério norte

def _zone_str_to_num_hemi(zstr):
z = zstr.strip().upper()
m = re.match(r'(\d{1,2})([NS])$', z)
if not m: return 22, 'S'
return int(m.group(1)), m.group(2)

def _auto_zone_from_city(cidade_field):
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

def _dms_to_dec(g, m, s):
try:
return float(g) + float(m)/60.0 + float(s)/3600.0
except:
return None

def _azimuth_from_quadrant(ns, angle_d, angle_m, angle_s, ew):
try:
d = float(angle_d); m = float(angle_m); s = float(angle_s)
except:
return None
theta = d + m/60 + s/3600
ns = (ns or '').upper().strip()
ew = (ew or '').upper().strip()
# conversão de quadrantal para azimute geodésico
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
if s == 60:
s = 0; m += 1
if m == 60:
m = 0; d += 1
return f"{d:02d}°{m:02d}'{s:02d}""

def degrees_to_dms_int(dec):
if dec is None:
return ""
dec = float(dec)
d = int(dec); m = int((dec - d) * 60); s = int(round((dec - d - m/60) * 3600))
if s == 60:
s = 0; m += 1
if m == 60:
m = 0; d += 1
return f"{d}°{m}'{s}""

def _fmt_coord_pair(lat, lon, fmt='utm', zone_num=None, hemi='S'):
try:
lat = float(lat); lon = float(lon)
except:
return ""
if fmt == 'dec':
return f"{lat:.6f}, {lon:.6f}"
elif fmt == 'dms':
return f"{degrees_to_dms_int(lat)} / {degrees_to_dms_int(lon)}"
else:
# UTM
crs_src = CRS.from_epsg(4674) # SIRGAS2000 geográfico
crs_dst = _sirgas_utm_crs(zone_num or 22, hemi or 'S')
tr = Transformer.from_crs(crs_src, crs_dst, always_xy=True)
e, n = tr.transform(lon, lat) # ordem lon, lat
return f"{int(round(e))} E, {int(round(n))} N"

def _fmt_coord_heading(fmt):
if fmt == 'dec': return "Coordenadas (graus decimais)"
if fmt == 'dms': return "Coordenadas (graus/minutos/segundos)"
return "Coordenadas (UTM)"

def _fmt_list(items, sep=", "):
return sep.join([str(x) for x in items if str(x).strip()])

def _cidade_sem_uf(cidade_field):
s = (cidade_field or '').strip()
if '/' in s:
return _title_keep_preps(s.split('/')[0].strip())
# tenta 'Porto Alegre - RS'
m = re.match(r'(.+?)[-\s,]*([A-Za-z]{2})$', s)
if m:
return _title_keep_preps(m.group(1).strip())
return _title_keep_preps(s)

===================== Docx helpers =====================

def set_run_font(run, size=12, bold=False, italic=False, color=None):
run.font.size = Pt(size)
run.font.bold = bool(bold)
run.font.italic = bool(italic)
if color:
run.font.color.rgb = RGBColor.from_string(color.replace('#',''))

def add_header_logo(doc, image_path):
if not os.path.exists(image_path): return
sec = doc.sections[0]
header = sec.header
p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
run = p.add_run()
try:
run.add_picture(image_path, width=Inches(2.0))
except:
pass

def add_footer_logo(doc, image_path):
if not os.path.exists(image_path): return
sec = doc.sections[0]
footer = sec.footer
p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
run = p.add_run()
try:
run.add_picture(image_path, width=Inches(1.5))
except:
pass

def add_footer_left_text(doc, lines, size_pt=9):
sec = doc.sections[0]
footer = sec.footer
p = footer.add_paragraph()
p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
for i, line in enumerate(lines):
r = p.add_run(line)
set_run_font(r, size=size_pt)
if i < len(lines)-1:
r.add_break()

def add_page_numbers(doc):
sec = doc.sections[0]
footer = sec.footer
p = footer.add_paragraph()
p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
# cria campo PAGE
r = p.add_run()
fld = OxmlElement('w:fldSimple')
fld.set(qn('w:instr'), 'PAGE \* MERGEFORMAT')
r = OxmlElement('w:r')
t = OxmlElement('w:t'); t.text = "1"
r.append(t)
fld.append(r); p._p.append(fld)

def add_corner_image_watermark_cm(doc, image_path, width_cm=6.46, height_cm=1.91):
if not os.path.exists(image_path): return
sec = doc.sections[0]; para = sec.header.add_paragraph()
r = para.add_run(); r.add_picture(image_path, width=Cm(width_cm))

def _apply_moderate_margins(doc):
for section in doc.sections:
section.top_margin = Inches(1.0)
section.bottom_margin = Inches(1.0)
section.left_margin = Inches(0.75)
section.right_margin = Inches(0.75)

def preparar_doc():
doc = Document()
_apply_moderate_margins(doc)
add_header_logo(doc, HEADER_LOGO_PATH)
add_corner_image_watermark_cm(doc, TL_PATH)
add_footer_logo(doc, FOOTER_LOGO_PATH)
return doc

--- Força o Word a atualizar campos (TOC) ao abrir ---

def _enable_update_fields_on_open(doc):
settings_el = doc.settings._element
for el in settings_el.iterchildren():
if el.tag == qn('w:updateFields'):
el.set(qn('w:val'), 'true')
return
upd = OxmlElement('w:updateFields')
upd.set(qn('w:val'), 'true')
settings_el.append(upd)

def _add_title(doc, txt):
p = doc.add_paragraph()
p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
r = p.add_run(txt)
set_run_font(r, size=16, bold=True)
p.space_after = Pt(6)

def heading(doc, txt):
p = doc.add_paragraph()
r = p.add_run(txt)
set_run_font(r, size=14, bold=True)
p.space_before = Pt(6); p.space_after = Pt(6)

def _add_hl(paragraph, text):
r = paragraph.add_run(text)
set_run_font(r)
r.font.highlight_color = WD_COLOR_INDEX.YELLOW

def _set_run_defaults(run, bold=False, italic=False):
set_run_font(run, size=12, bold=bold, italic=italic)

===================== Parsers Civil 3D =====================

def parse_civilreport_from_html(html_bytes):
"""
Parser simplificado para HTML do Civil 3D (CIVILREPORT).
Aceita arquivos .html/.htm que contenham tabelas de segmentos e coordenadas.
"""
soup = BeautifulSoup(html_bytes, 'lxml')
txt = soup.get_text(" ", strip=True).upper()
res = []
# tentativa de achar blocos por "POLILINHA", "ALINHAMENTO" etc.
# cada bloco terá: name, perimetro, first_point (lat, lon), segs [...]
pol_blocks = soup.find_all(text=re.compile("CIVILREPORT", re.I))
# fallback: procurar por tabelas
tables = soup.find_all('table')
for tb in tables:
    # extrai linhas com possíveis pares "N|S/E|W ... graus ... metros"
    segs = []
    rows = tb.find_all('tr')
    for tr in rows:
        cols = [c.get_text(strip=True) for c in tr.find_all(['td','th'])]
        if len(cols) < 2: 
            continue
        row_txt = " ".join(cols).upper()
        # heurísticas (rústicas) para segmentos
        # formato: N 12°34'56" E 123,45
        m = re.search(r'([NS])\s*(\d+)[°º]\s*(\d+)\'\s*(\d+)"\s*([EW])\s*([0-9\.,]+)', row_txt)
        if m:
            ns, gd, gm, gs, ew, dist = m.groups()
            az = _azimuth_from_quadrant(ns, gd, gm, gs, ew)
            dist_v = to_float_any(dist)
            if az is not None and dist_v > 0:
                segs.append({"azimuth": az, "dist": dist_v})
            continue
        # formato: AZ 123°34'56" DIST 123,45
        m = re.search(r'AZ\s*(\d+)[°º]\s*(\d+)\'\s*(\d+)"\s*DIST\s*([0-9\.,]+)', row_txt)
        if m:
            gd, gm, gs, dist = m.groups()
            az = _dms_to_dec(gd, gm, gs)
            dist_v = to_float_any(dist)
            if az is not None and dist_v > 0:
                segs.append({"azimuth": az, "dist": dist_v})
            continue
    if segs:
        res.append({"name": "ALINHAMENTO", "segments": segs})
# tenta coordenadas iniciais
# padrões típicos: "Latitude: -30,123456 Longitude: -51,123456"
m_lat = re.search(r'LAT(?:ITUDE)?:\s*(-?[0-9\.,]+)', txt)
m_lon = re.search(r'LON(?:GITUDE)?:\s*(-?[0-9\.,]+)', txt)
first_point = None
if m_lat and m_lon:
    lat = to_float_any(m_lat.group(1))
    lon = to_float_any(m_lon.group(1))
    first_point = (lat, lon)

# perímetro aproximado: somatório das distâncias
perim = 0.0
for b in res:
    for s in b["segments"]:
        perim += s.get("dist", 0.0)

ret = []
for i, b in enumerate(res):
    item = {
        "name": b["name"] if b.get("name") else f"ALINHAMENTO {i+1}",
        "perimetro": perim,
        "first_point": first_point,
        "segments": b["segments"]
    }
    ret.append(item)
return ret
def quadra_label_sort_key(lbl):
# tenta extrair o número após "QUADRA". Caso falhe, usa string
s = (lbl or '').upper().strip()
m = re.search(r'QUADRA\s+(\d+)', s)
if m:
return int(m.group(1))
return s

def _lote_num(lbl):
s = (lbl or '').upper().strip()
m = re.search(r'(\d+)', s)
return int(m.group(1)) if m else 0

===================== Seções e textos (DOCX) =====================

def _sec_cabecalho(doc, cidade_field, nome_emp):
_add_title(doc, "MEMORIAL DESCRITIVO / ESPECIFICAÇÕES")
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6)
r = p.add_run("Empreendimento: "); _set_run_defaults(r, bold=True)
r = p.add_run(nome_emp or "XXXX"); _set_run_defaults(r)
p = doc.add_paragraph(); r = p.add_run("Município: "); _set_run_defaults(r, bold=True)
r = p.add_run(_fmt_cidade_slash_uf(cidade_field) or "XXXX/UF"); _set_run_defaults(r)

def _sec_loc(doc, endereco, bairro, cidade_field, data_auto=True):
heading(doc, "LOCALIZAÇÃO")
p = doc.add_paragraph()
r = p.add_run("Endereço: "); _set_run_defaults(r, bold=True)
r = p.add_run(endereco or "XXXX"); _set_run_defaults(r)
p = doc.add_paragraph()
r = p.add_run("Bairro: "); _set_run_defaults(r, bold=True)
r = p.add_run(_fmt_bairro(bairro) or "XXXX"); _set_run_defaults(r)
p = doc.add_paragraph()
r = p.add_run("Município: "); _set_run_defaults(r, bold=True)
r = p.add_run(_fmt_cidade_slash_uf(cidade_field) or "XXXX/UF"); _set_run_defaults(r)
if data_auto:
p = doc.add_paragraph()
r = p.add_run("Data: "); _set_run_defaults(r, bold=True)
r = p.add_run(datetime.now().strftime("%d/%m/%Y")); _set_run_defaults(r)

def _sec_area(doc, area_total_m2, perimetro_m=None):
heading(doc, "ÁREA DO EMPREENDIMENTO")
p = doc.add_paragraph()
r = p.add_run("Área Total: "); _set_run_defaults(r, bold=True)
r = p.add_run(f"{_fmt_br(area_total_m2, 2)} m² ({area_por_extenso(area_total_m2)})"); _set_run_defaults(r)
if perimetro_m is not None and perimetro_m > 0:
p = doc.add_paragraph()
r = p.add_run("Perímetro: "); _set_run_defaults(r, bold=True)
r = p.add_run(f"{_fmt_br(perimetro_m, 2)} m ({extenso_metros(perimetro_m)})"); _set_run_defaults(r)

def _sec_coord_heading(doc, fmt):
heading(doc, _fmt_coord_heading(fmt))

def _sec_first_point(doc, fp, fmt, zone_num, hemi):
if not fp: return
p = doc.add_paragraph()
lat, lon = fp
coord_s = _fmt_coord_pair(lat, lon, fmt, zone_num, hemi)
r = p.add_run(f"Ponto de referência (coordenadas {fmt.upper()}): {coord_s}")
_set_run_defaults(r)

def _sec_segments(doc, segments):
heading(doc, "SEGMENTOS / AZIMUTES")
# tabela simples de azimute + distância
from docx.shared import Inches
table = doc.add_table(rows=1, cols=2)
hdr = table.rows[0].cells
hdr[0].text = "Azimute (DMS)"; hdr[1].text = "Distância (m)"
# cabeçalho em negrito
for cell in hdr:
for p in cell.paragraphs:
for r in p.runs:
r.font.bold = True
for s in segments or []:
az = s.get("azimuth"); dist = s.get("dist")
row = table.add_row().cells
row[0].text = str(azimuth_to_dms_int(az) or "")
row[1].text = _fmt_br(dist, 2)

def _sec_ane(doc, enable, largura_m=None):
if not enable: return
heading(doc, "ÁREA NÃO EDIFICANTE (ANE)")
p = doc.add_paragraph()
r = p.add_run("Largura: "); _set_run_defaults(r, bold=True)
r = p.add_run(f"{_fmt_br(largura_m, 2)} m"); _set_run_defaults(r)

def _sec_assinaturas_simples(doc, cidade_field=None):
doc.add_paragraph()
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6); p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
r = p.add_run("Porto Alegre, "); set_run_defaults(r)
r = p.add_run(datetime.now().strftime("%d de %B de %Y")); set_run_defaults(r)
doc.add_paragraph()
p = doc.add_paragraph(); r = p.add_run("________________________________"); _set_run_defaults(r)
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6); r = p.add_run("Responsável técnico"); _set_run_defaults(r, bold=True)
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(0); r = p.add_run("SOLIDO - DESIGN URBANO LTDA."); _set_run_defaults(r)
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(0)
r = p.add_run("CAU-RS 15335-4"); _set_run_defaults(r)

def _sec_assinaturas_resumo(doc):
_add_title(doc, "ASSINATURAS")
for _ in range(2): doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
r = p.add_run("_____________________________"); _set_run_defaults(r)
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6); r = p.add_run("Responsável técnico"); _set_run_defaults(r, bold=True)

def _sec_retro(doc, usos, topografia, has_ai, has_restricao):
heading(doc, "INFORMAÇÕES COMPLEMENTARES")
if usos:
p = doc.add_paragraph(); r = p.add_run("Usos: "); _set_run_defaults(r, bold=True)
r = p.add_run(_fmt_list(usos)); _set_run_defaults(r)
p = doc.add_paragraph(); r = p.add_run("Topografia: "); _set_run_defaults(r, bold=True)
r = p.add_run(topografia or ""); _set_run_defaults(r)
p = doc.add_paragraph(); r = p.add_run("Área Institucional: "); _set_run_defaults(r, bold=True)
r = p.add_run("Sim" if has_ai else "Não"); _set_run_defaults(r)
p = doc.add_paragraph(); r = p.add_run("Restrição: "); _set_run_defaults(r, bold=True)
r = p.add_run("Sim" if has_restricao else "Não"); _set_run_defaults(r)

def _sec_situacao_atual(doc, pres_unif, pres_desm):
heading(doc, "SITUAÇÃO ATUAL")
p = doc.add_paragraph()
txt = []
if pres_unif: txt.append("há áreas a unificar")
if pres_desm: txt.append("há áreas a desmembrar")
if not txt: txt.append("sem alterações")
r = p.add_run(" e ".join(txt) + ".")
_set_run_defaults(r)

def _titulo_para_unif_desm(pres_unif, pres_desm):
if pres_unif and pres_desm:
return "MEMORIAL DE UNIFICAÇÃO E DESMEMBRAMENTO"
elif pres_unif:
return "MEMORIAL DE UNIFICAÇÃO"
elif pres_desm:
return "MEMORIAL DE DESMEMBRAMENTO"
else:
return "MEMORIAL"

def _primeiro_paragrafo_unif_desm(doc, pres_unif, pres_desm):
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6)
if pres_unif and pres_desm:
txt = "Este memorial descreve procedimentos de unificação e desmembramento conforme normas vigentes."
elif pres_unif:
txt = "Este memorial descreve procedimento de unificação conforme normas vigentes."
elif pres_desm:
txt = "Este memorial descreve procedimento de desmembramento conforme normas vigentes."
else:
txt = "Este memorial descreve o empreendimento sem alterações de matrícula."
r = p.add_run(txt); _set_run_defaults(r)

def _sec_unificacao(doc, unif_item):
heading(doc, "UNIFICAÇÃO")
p = doc.add_paragraph(); r = p.add_run("Descrição da área a unificar."); _set_run_defaults(r)

def _sec_desmembramento(doc, desm_items):
heading(doc, "DESMEMBRAMENTO")
p = doc.add_paragraph(); r = p.add_run("Descrição das áreas a desmembrar."); _set_run_defaults(r)

===================== MEMORIAL RESUMO / SOLICITAÇÃO ANÁLISE =====================

def _build_memorial_resumo_doc():
doc = preparar_doc()
_add_title(doc, "MEMORIAL RESUMO / DESCRITIVO")
# cabeçalho
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
r = p.add_run("Tipo do empreendimento: "); _set_run_defaults(r, bold=True)
r = p.add_run(tipo_proj_resumo.value or ""); _set_run_defaults(r)
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
r = p.add_run("Usos: "); _set_run_defaults(r, bold=True)
r = p.add_run(_fmt_list(usos_multi.value) or ""); _set_run_defaults(r)
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
r = p.add_run("Topografia: "); _set_run_defaults(r, bold=True)
r = p.add_run(topografia.value or ""); _set_run_defaults(r)
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
r = p.add_run("Área Institucional: "); _set_run_defaults(r, bold=True)
r = p.add_run("Sim" if has_ai.value else "Não"); _set_run_defaults(r)
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
r = p.add_run("Restrição: "); _set_run_defaults(r, bold=True)
r = p.add_run("Sim" if has_restricao.value else "Não"); _set_run_defaults(r)
# localização
_sec_loc(doc, endereco_emp.value, bairro_emp.value, cidade_emp.value, data_auto=True)

# área
try:
    at = _to_float_br(area_total_emp.value)
except:
    at = 0.0
_sec_area(doc, at, None)

# assinaturas
_sec_assinaturas_resumo(doc)

add_footer_left_text(doc, [
    "WWW.SOLIDO.ARQ.BR",
    "Avenida Ipiranga, 6681 – Prédio 99, Sala 906",
    "Porto Alegre – RS Brasil",
    "+ 55 51 99690-7857",
], size_pt=10)
add_page_numbers(doc)

cidade_nome = _cidade_sem_uf(cidade_emp.value or '')
out_docx = f"/content/URB-PL_XXXX_MEMORIAL RESUMO_RX_VX.docx"
doc.save(out_docx)
return out_docx
def _build_solicitacao_analise_doc():
doc = preparar_doc()
_add_title(doc, "OFÍCIO – SOLICITAÇÃO DE ANÁLISE")
p = doc.add_paragraph()
_set_run_defaults(p.add_run("À Prefeitura Municipal de ")); _set_run_defaults(p.add_run(_cidade_sem_uf(cidade_emp.value or 'XXXX')))
p = doc.add_paragraph()
_set_run_defaults(p.add_run("Assunto: Solicitação de análise de projeto."), bold=True)
p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
_set_run_defaults(p.add_run(
    "Por meio deste, encaminhamos documentação referente ao empreendimento para análise. Seguem anexos conforme lista: "
))
doc.add_paragraph("– Memorial Descritivo;")
doc.add_paragraph("– Planta de Situação;")
doc.add_paragraph("– Matrícula/Registro;")
doc.add_paragraph("– ART/RRT do responsável;")

p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
_set_run_defaults(p.add_run("Nos colocamos à disposição para esclarecimentos e pedimos o deferimento."))

p = doc.add_paragraph()
p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
_set_run_defaults(p.add_run("Porto Alegre, "))
_set_run_defaults(p.add_run(datetime.now().strftime("%d de %B de %Y")))

for _ in range(2): doc.add_paragraph()
p = doc.add_paragraph(); p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
_set_run_defaults(p.add_run("__________________________________"))
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6)
_set_run_defaults(p.add_run("Responsável técnico"), bold=True)

# rodapé padrão
add_footer_left_text(doc, [
    "WWW.SOLIDO.ARQ.BR",
    "Avenida Ipiranga, 6681 – Prédio 99, Sala 906",
    "Porto Alegre – RS Brasil",
    "+ 55 51 99690-7857",
], size_pt=10)
add_page_numbers(doc)

# salvar arquivo
cidade_nome = _cidade_sem_uf(cidade_emp.value)
out_path = f"/content/URB-PL_XXXX_OFÍCIO SOLICITAÇÃO DE ANÁLISE_RX_VX.docx"
doc.save(out_path)
return out_path
===================== Excel UNIF/DESM: helpers (única versão) =====================

def _format_first_point(fp, coord_fmt, zone_num, hemi):
if not fp: return ""
lat, lon = fp
return _fmt_coord_pair(lat, lon, coord_fmt, zone_num, hemi)

def _collect_items_unif_desm():
modo = tipo_emp.value
items_unif = None
items_desm = []
civil_htmls = [(f,d) for f,d in uploaded_files.items()
               if f.lower().endswith(('.html','.htm')) and 'CIVILREPORT' in f.upper()]
other_htmls = [(f,d) for f,d in uploaded_files.items()
               if f.lower().endswith(('.html','.htm')) and 'CIVILREPORT' not in f.upper()]

if modo in ('unificacao','unif_desm'):
    for fname, data in civil_htmls:
        arr = parse_civilreport_from_html(io.BytesIO(data).read())
        for it in arr:
            items_unif = it  # assume apenas um?

if modo in ('desmembramento','unif_desm'):
    for fname, data in civil_htmls:
        arr = parse_civilreport_from_html(io.BytesIO(data).read())
        for it in arr:
            items_desm.append((Path(fname).stem, it))

return items_unif, items_desm
def _build_quadro_fracao_df(unif_item, desm_items):
colunas = ['Lote','Quadra','Área Privativa (m²)','Área Uso Comum (m²)','Área Real Total (m²)','Fração Ideal']
linhas = []
def _add_area(nome, item):
# placeholders
linhas.append({
'Lote': '1', 'Quadra': 'A',
'Área Privativa (m²)': _fmt_br(100.0),
'Área Uso Comum (m²)': _fmt_br(20.0),
'Área Real Total (m²)': _fmt_br(120.0),
'Fração Ideal': _fmt_br(0.0123, 4),
})
if not unif_item and not desm_items:
linhas.append({c: "" for c in colunas})
if unif_item: _add_area(unif_item.get("name") or "UNIFICAÇÃO", unif_item)
for nm, it in desm_items or []: _add_area(nm, it)
return pd.DataFrame(linhas, columns=colunas)
===================== Geração principal (on_click adaptado) =====================

def on_download_excel_clicked(_):
global _last_dados_quadro, _last_eh_condominio
# caso condomínio: usa dados calculados no DOCX gerado
if tipo_emp.value == 'condominio':
if not _last_eh_condominio:
print("📎 O Excel de fração ideal só se aplica a condomínio. Gere o DOCX primeiro.")
return
if not _last_dados_quadro:
print("⚠️ Gere o DOCX primeiro para calcular a fração ideal.")
return
try:
df = pd.DataFrame(_last_dados_quadro, columns=[
'Lote','Quadra','Área Privativa (m²)','Área Uso Comum (m²)','Área Real Total (m²)','Fração Ideal'
])
cidade_nome = _cidade_sem_uf(cidade_emp.value)
xlsx_path = f"/content/URB-PL_XXXX_QUADRO FRAÇÃO IDEAL_RX_VX.xlsx"
df['quad_key'] = df['Quadra'].map(lambda q: quadra_label_sort_key(f"QUADRA {q}"))
df['lote_key'] = df['Lote'].map(_lote_num)
df = df.sort_values(['quad_key','lote_key']).drop(columns=['quad_key','lote_key'])
df.to_excel(xlsx_path, index=False)
        from openpyxl import load_workbook
        from openpyxl.styles import Font, Alignment, Border, Side
        wb = load_workbook(xlsx_path); ws = wb.active
        font_header = Font(name='Calibri', size=12, bold=True)
        font_cell = Font(name='Calibri', size=12)
        center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        thin = Side(border_style='thin', color='000000'); border = Border(left=thin,right=thin,top=thin,bottom=thin)
        for r in ws.iter_rows():
            for c in r:
                c.alignment = center
                c.border = border
                c.font = font_header if c.row == 1 else font_cell
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = max(12, min(32, max_len + 2))
        wb.save(xlsx_path)
        print(f"✅ Gerado: {xlsx_path}")
        files.download(xlsx_path)
        return
    except Exception as e:
        print(f"❌ Erro ao gerar Excel: {e}")
        return

# caso unificação/desmembramento: gera quadro a partir dos uploads
unif_item, desm_items = _collect_items_unif_desm()
df = _build_quadro_fracao_df(unif_item, desm_items)
xlsx_path = f"/content/URB-PL_XXXX_QUADRO FRAÇÃO IDEAL_RX_VX.xlsx"
df.to_excel(xlsx_path, index=False)
print(f"✅ Gerado: {xlsx_path}")
files.download(xlsx_path)
def on_generate_clicked(_):
global _last_dados_quadro, _last_eh_condominio
# limpar flag
_last_dados_quadro = []
_last_eh_condominio = False
try:
    modo = tipo_emp.value

    # NOVO: MEMORIAL RESUMO/DESCRITIVO
    if modo == 'memorial_resumo':
        out_path = _build_memorial_resumo_doc()
        print(f"✅ Gerado: {out_path}")
        files.download(out_path)
        return

    # >>> NOVO BLOCO <<<
    if modo == 'solicitacao_analise':
        out_path = _build_solicitacao_analise_doc()
        print(f"✅ Gerado: {out_path}")
        files.download(out_path)
        return

    # UNIFICAÇÃO/DESMEMBRAMENTO
    if modo in ('unificacao', 'desmembramento', 'unif_desm'):
        unif_item, desm_items = _collect_items_unif_desm()
        doc = preparar_doc()
        pres_unif = bool(unif_item)
        pres_desm = bool(desm_items)

        heading(doc, _titulo_para_unif_desm(pres_unif, pres_desm))
        _primeiro_paragrafo_unif_desm(doc, pres_unif, pres_desm)
        _sec_situacao_atual(doc, pres_unif, pres_desm)

        zone_num, hemi = _auto_zone_from_city(cidade_emp.value or '')
        if pres_unif: _sec_unificacao(doc, unif_item)
        if pres_desm: _sec_desmembramento(doc, desm_items)

        # salva
        cidade_nome = _cidade_sem_uf(cidade_emp.value or '')
        out_docx = f"/content/URB-PL_XXXX_MEMORIAL UNIF_DESM_RX_VX.docx"
        doc.save(out_docx)
        print(f"✅ Gerado: {out_docx}")
        files.download(out_docx)
        return

    # MEMORIAIS (loteamento/condomínio)
    doc = preparar_doc()

    # Cabeçalho
    _sec_cabecalho(doc, cidade_emp.value, nome_emp.value)

    # Localização
    _sec_loc(doc, endereco_emp.value, bairro_emp.value, cidade_emp.value, data_auto=True)

    # Área
    try:
        area_total = _to_float_br(area_total_emp.value)
    except:
        area_total = 0.0
    try:
        perimetro_m = _to_float_br(perimetro_emp.value)
    except:
        perimetro_m = 0.0

    _sec_area(doc, area_total, perimetro_m if perimetro_m > 0 else None)

    # Coordenadas/segmentos a partir dos uploads
    try:
        zone_num, hemi = _auto_zone_from_city(cidade_emp.value or '')
    except:
        zone_num, hemi = 22, 'S'

    coord_fmt_val = coord_fmt.value or 'utm'

    ane_enable = (ane_drop.value == 'Sim')
    ane_largura_m = None
    if ane_enable and ane_largura.value.strip():
        try: ane_largura_m = _to_float_br(ane_largura.value)
        except: ane_largura_m = None

    civil_items = []
    for fname, data in uploaded_files.items():
        civil_items.extend(parse_civilreport_from_html(io.BytesIO(data).read()))

    # usa primeiro ponto global se houver
    first_point = None
    for it in civil_items:
        if it.get("first_point"):
            first_point = it["first_point"]
            break

    _sec_coord_heading(doc, coord_fmt_val)
    _sec_first_point(doc, first_point, coord_fmt_val, zone_num, hemi)
    # segmentos
    for it in civil_items:
        _sec_segments(doc, it.get("segments") or [])

    # ANE
    _sec_ane(doc, ane_enable, ane_largura_m)

    # Rodapé + assinaturas
    def _sec_assinaturas_simples_local(doc):
        p = doc.add_paragraph(); r = p.add_run("__________________________________"); _set_run_defaults(r)
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6)
        r = p.add_run("Responsável técnico"); _set_run_defaults(r, bold=True)
    _sec_assinaturas_simples_local(doc)

    add_footer_left_text(doc, [
        "WWW.SOLIDO.ARQ.BR",
        "Avenida Ipiranga, 6681 – Prédio 99, Sala 906",
        "Porto Alegre – RS Brasil",
        "+ 55 51 99690-7857",
    ], size_pt=10)
    add_page_numbers(doc)

    cidade_nome = _cidade_sem_uf(cidade_emp.value)
    out_docx = f"/content/URB-PL_XXXX_MEMORIAL DE LOTES_RX_VX.docx"
    doc.save(out_docx)
    print(f"✅ Gerado: {out_docx}")
    files.download(out_docx)

except Exception as e:
    import traceback, sys
    print("❌ Erro ao gerar o DOCX:")
    traceback.print_exc(file=sys.stdout)
===================== Streamlit UI (Substitui Colab/Jupyter) =====================

import streamlit as st
import os, io
from dataclasses import dataclass

===== Adaptação para Streamlit =====
Criar diretório /content para manter caminhos originais

os.makedirs("/content", exist_ok=True)

Dummy 'files' para capturar downloads solicitados no código original

class _DummyFiles:
last_path = None
def download(self, path):
self.last_path = path

files = _DummyFiles()

Dummy 'out' para compatibilidade

class _DummyOut:
def clear_output(self): pass
def enter(self): return self
def exit(self, *exc): return False
out = _DummyOut()

Helpers simples para os "widgets" originais: objeto com atributo .value

@dataclass
class _V:
value: any

===== Interface Streamlit =====

st.title("Gerar Memorial a partir do HTML/TXT (Civil 3D)")

tipo_label_to_val = [
('Memorial Condomínio', 'condominio'),
('Memorial Loteamento', 'loteamento'),
('Memorial Unificação', 'unificacao'),
('Memorial Desmembramento', 'desmembramento'),
('Memorial Unificação e Desmembramento', 'unif_desm'),
('Memorial Resumo', 'memorial_resumo'),
('Solicitação de Análise', 'solicitacao_analise'),
]

tipo_emp = _V(st.selectbox('Tipo:', options=[l for l,v in tipo_label_to_val], index=0))

Map back to value

_tipo_map = {l:v for l,v in tipo_label_to_val}
tipo_emp.value = _tipo_map.get(tipo_emp.value, 'condominio')

nome_emp = _V(st.text_input('Nome do Empreendimento:', ''))
endereco_emp = _V(st.text_input('Endereço:', ''))
bairro_emp = _V(st.text_input('Bairro:', ''))
cidade_emp = _V(st.text_input('Cidade/UF:', ''))
area_total_emp = _V(st.text_input('Área total (m²):', ''))
perimetro_emp = _V(st.text_input('Perímetro (m):', ''))
matricula_emp = _V(st.text_input('Matrícula nº:', ''))
num_lotes_emp = _V(st.number_input('Nº de lotes:', min_value=0, step=1, value=0))

area_tot_priv_emp = _V(st.text_input('Área Privativa (m²):', ''))
area_tot_cond_emp = _V(st.text_input('Área Condominial (m²):', ''))

ane_drop = _V(st.selectbox('Área não edificante?', ['Não','Sim'], index=0))
ane_largura = _V(st.text_input('Largura (m):', ''))
coord_fmt = _V(st.selectbox('Coordenadas:', ['utm','dec','dms'], index=0))

data_auto = _V(True)

tipo_proj_resumo = _V(st.selectbox('Tipo de empreendimento:', ['condominio','loteamento'], index=0))
usos_multi = _V(st.multiselect('Usos:', ['Residencial','Comercial','Industrial']))
topografia = _V(st.selectbox('Topografia:', ['Acentuada','Plana'], index=0))
has_ai = _V(st.checkbox('Área Institucional', value=False))
has_restricao = _V(st.checkbox('Restrição', value=False))

Upload de arquivos HTML/TXT (Civil 3D)

up_files = st.file_uploader("Anexar HTML(s)", type=['html','htm','txt'], accept_multiple_files=True)
uploaded_files = {}
if up_files:
for f in up_files:
uploaded_files[f.name] = f.getvalue()

Flags usadas no Excel

_last_dados_quadro = []
_last_eh_condominio = False

col1, col2 = st.columns(2)
gen_clicked = col1.button("Gerar DOCX")
xls_clicked = col2.button("Baixar Excel")

Vincular funções do código original

if gen_clicked:
try:
on_generate_clicked(None)
except Exception as e:
st.error(f"Erro ao gerar: {e}")
else:
if files.last_path and os.path.exists(files.last_path):
with open(files.last_path, 'rb') as fh:
data = fh.read()
st.download_button("Baixar DOCX", data=data, file_name=os.path.basename(files.last_path), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
else:
st.info("Nenhum arquivo gerado.")

if xls_clicked:
try:
on_download_excel_clicked(None)
except Exception as e:
st.error(f"Erro ao gerar Excel: {e}")
else:
if files.last_path and os.path.exists(files.last_path):
with open(files.last_path, 'rb') as fh:
data = fh.read()
st.download_button("Baixar Excel", data=data, file_name=os.path.basename(files.last_path), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
st.info("Nenhum Excel gerado.")
