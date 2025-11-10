from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime

# ============================================================
# Funções utilitárias básicas (estilo do seu código original)
# ============================================================

def _set_run_defaults(run):
    """Define o estilo de texto padrão (Calibri 12)."""
    run.font.name = "Calibri"
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0, 0, 0)
    return run

def _novo_doc():
    doc = Document()
    return doc

# ============================================================
# Funções principais (simulando seus blocos do Colab)
# ============================================================

def gerar_memorial_memorial(form, arquivos):
    """
    Gera DOCX de memorial de condomínio/loteamento/unificação/desmembramento.
    """
    doc = _novo_doc()

    tipo = form.get("tipo", "")
    nome = form.get("nome", "")
    cidade = form.get("cidade", "")
    area_total = form.get("area_total", "")
    num_lotes = form.get("num_lotes", "")
    ane = form.get("ane_enable", False)
    ane_larg = form.get("ane_largura", "")

    # Cabeçalho
    p = doc.add_paragraph()
    run = _set_run_defaults(p.add_run(f"MEMORIAL {tipo.upper()}"))
    run.bold = True
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    doc.add_paragraph()
    p = doc.add_paragraph(f"Empreendimento: {nome}")
    _set_run_defaults(p.runs[0])
    p = doc.add_paragraph(f"Cidade/UF: {cidade}")
    _set_run_defaults(p.runs[0])
    p = doc.add_paragraph(f"Área total: {area_total} m²")
    _set_run_defaults(p.runs[0])
    if num_lotes:
        p = doc.add_paragraph(f"Número de lotes: {num_lotes}")
        _set_run_defaults(p.runs[0])
    if ane:
        p = doc.add_paragraph(f"Área não edificante: {ane_larg} m")
        _set_run_defaults(p.runs[0])

    if arquivos:
        doc.add_paragraph()
        doc.add_paragraph("Arquivos anexados:")
        for nome_arq in arquivos.keys():
            p = doc.add_paragraph(f"• {nome_arq}")
            _set_run_defaults(p.runs[0])

    doc.add_page_break()
    _assinatura(doc)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    filename = f"{tipo}_{nome}.docx".replace(" ", "_")
    return buffer.getvalue(), filename


def gerar_memorial_resumo(form):
    """
    Gera DOCX de memorial resumo (síntese do empreendimento).
    """
    doc = _novo_doc()

    nome = form.get("nome", "")
    cidade = form.get("cidade", "")
    area_total = form.get("area_total", "")
    tipo_proj = form.get("tipo_proj_resumo", "")
    usos = ", ".join(form.get("usos_multi", []))
    topografia = form.get("topografia", "")
    has_ai = "Sim" if form.get("has_ai") else "Não"
    has_restricao = "Sim" if form.get("has_restricao") else "Não"

    p = doc.add_paragraph("MEMORIAL RESUMO")
    _set_run_defaults(p.add_run())
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    doc.add_paragraph(f"Empreendimento: {nome}")
    doc.add_paragraph(f"Cidade/UF: {cidade}")
    doc.add_paragraph(f"Área total: {area_total} m²")
    doc.add_paragraph(f"Tipo: {tipo_proj}")
    doc.add_paragraph(f"Usos: {usos}")
    doc.add_paragraph(f"Topografia: {topografia}")
    doc.add_paragraph(f"Possui área institucional? {has_ai}")
    doc.add_paragraph(f"Possui área de restrição? {has_restricao}")

    doc.add_page_break()
    _assinatura(doc)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    filename = f"memorial_resumo_{nome}.docx".replace(" ", "_")
    return buffer.getvalue(), filename


def gerar_solicitacao_analise(form):
    """
    Gera DOCX de solicitação de análise.
    """
    doc = _novo_doc()
    nome = form.get("nome", "")
    cidade = form.get("cidade", "")
    tipo_proj = form.get("tipo_proj_resumo", "")
    area_total = form.get("area_total", "")

    p = doc.add_paragraph("SOLICITAÇÃO DE ANÁLISE DE PROJETO URBANÍSTICO")
    _set_run_defaults(p.add_run())
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    doc.add_paragraph()
    doc.add_paragraph(f"Prefeitura Municipal de {cidade.upper()}")
    doc.add_paragraph(f"Objeto: Solicitação de análise de Projeto Urbanístico")
    doc.add_paragraph(
        f"O {tipo_proj} '{nome}' possui área total de {area_total} m² e solicita análise conforme legislação vigente."
    )

    doc.add_paragraph()
    data_txt = datetime.today().strftime("%d de %B de %Y")
    doc.add_paragraph(f"Porto Alegre, {data_txt}.")

    doc.add_page_break()
    _assinatura(doc)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    filename = f"solicitacao_{nome}.docx".replace(" ", "_")
    return buffer.getvalue(), filename


def _assinatura(doc):
    doc.add_paragraph()
    p = doc.add_paragraph("_______________________________________")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p = doc.add_paragraph("Responsável Técnico")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
