# -*- coding: utf-8 -*-
r"""
scripts/gerador/renderers/docx_renderer.py - Fase 3, Slice 3: document_model ->
shadow_report.docx (spec Sec.4/5), com python-docx.

So conhece `gerador.document_model.Document` -- nunca report_data.json nem
caminho de planilha (isso e' trabalho do content_builder). Trocar o template
ou a biblioteca de renderizacao nao exige mexer em content_builder nem em
document_model.

Estilo: Arial 11pt no corpo (definido no estilo 'Normal', nao por formatacao
direta), titulos com niveis reais de outline (doc.add_heading), margens
basicas 2,5/3,0 cm. NAO pixel-perfect -- sem microespacamento, sem posicao
exata de imagem (spec Sec.5).

Marca de HITL/PLACEHOLDER/LIMITATION: realce amarelo + prefixo textual
inequivoco. Nenhum bloco com status != READY sai sem essa marca (spec Sec.5:
"nada de pendencia disfarcada de conteudo pronto").
"""
import os
import sys

from docx import Document as DocxDocument
from docx.shared import Pt, Cm
from docx.enum.text import WD_COLOR_INDEX
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(os.path.dirname(AQUI))
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402

_PAGE_DIMS = {
    "A4 retrato": (Cm(21.0), Cm(29.7)),
    "A3 retrato": (Cm(29.7), Cm(42.0)),
}

# P0.6: margens por PERFIL de secao (cover/body/photo), nao mais um unico
# valor fixo para o documento inteiro -- e' o que deixa a capa
# estruturalmente independente do corpo (spec P0.6: "capa A4 (margens
# proprias, titlePg), corpo A4, fotografico A3, volta para corpo A4").
# Os VALORES mm exatos de cada perfil (25/30/25/30 na capa, 30/20/20/30 no
# corpo etc.) sao P1.1 (config/visual_style_v3.yaml) -- fora do escopo desta
# rodada. Aqui os tres perfis ainda usam o generico historico (2,5/3,0 cm);
# so a MECANICA (perfil -> margem propria, capa com titlePg) e' nova em P0.
# (ordem da tupla: topo, base, esquerda, direita)
_MARGENS_POR_PERFIL = {
    "cover": (Cm(2.5), Cm(2.5), Cm(3.0), Cm(3.0)),
    "body": (Cm(2.5), Cm(2.5), Cm(3.0), Cm(3.0)),
    "photo": (Cm(2.5), Cm(2.5), Cm(3.0), Cm(3.0)),
}

_PREFIXOS = {
    # colchetes ASCII de proposito (nao angulos unicode <U+27E8>): o console
    # Windows padrao (PowerShell, cp1252) nao imprime esses caracteres e
    # quebraria qualquer log/teste que os reproduzisse -- marca visual no
    # DOCX nao pode depender de um glyph que nem sempre roundtrip-a.
    "HITL": "[HITL] ",
    "LIMITATION": "[LIMITAÇÃO DA FONTE] ",
    "PLACEHOLDER": "[PENDENTE] ",
}


def _setup_estilo(doc):
    """Corpo Arial 11pt DEFINIDO NO ESTILO (nao por formatacao direta em
    cada run) -- spec PADRAO_ATUAL_RELATORIO_SEMANAL.md §3."""
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(11)
    try:
        heading = doc.styles["Heading 1"]
        heading.font.name = "Arial"
    except KeyError:
        pass


def _set_page(section, page_setup, page_profile="body"):
    largura, altura = _PAGE_DIMS.get(page_setup, _PAGE_DIMS["A4 retrato"])
    section.page_width = largura
    section.page_height = altura
    topo, base, esquerda, direita = _MARGENS_POR_PERFIL.get(
        page_profile, _MARGENS_POR_PERFIL["body"])
    section.top_margin = topo
    section.bottom_margin = base
    section.left_margin = esquerda
    section.right_margin = direita
    # P0.6: titlePg -- a capa e' a UNICA secao com paginas-de-rosto distintas
    # (sem cabecalho/rodape do corpo, quando P1 os implementar). Nao existe
    # secao A0/A2 no Word (P0.1) -- so "cover"/"body"/"photo" sao validos.
    section.different_first_page_header_footer = (page_profile == "cover")


def _marcar(run, status):
    if status == "READY":
        return
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    run.bold = True


def _nota(doc, note):
    if not note:
        return
    p = doc.add_paragraph()
    r = p.add_run("Nota: %s" % note)
    r.italic = True
    r.font.size = Pt(9)


def _render_table(doc, block):
    headers = block.headers or []
    rows = block.rows or []
    ncols = len(headers) if headers else (len(rows[0]) if rows else 1)
    tbl = doc.add_table(rows=0, cols=max(ncols, 1))
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    if headers:
        cells = tbl.add_row().cells
        for i, h in enumerate(headers):
            cells[i].text = str(h)
            for p in cells[i].paragraphs:
                for r in p.runs:
                    r.bold = True
    for row in rows:
        cells = tbl.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = "" if v is None else str(v)
    if block.status != "READY":
        p = doc.add_paragraph()
        r = p.add_run(_PREFIXOS.get(block.status, "") + (block.note or "conteúdo da tabela não comprovado"))
        _marcar(r, block.status)
    else:
        _nota(doc, block.note)


def _render_block(doc, block):
    prefixo = _PREFIXOS.get(block.status, "") if block.status != "READY" else ""
    kind = block.kind

    if kind == "heading":
        p = doc.add_heading(block.text or "", level=max(1, min(int(block.level or 1), 9)))
        if block.status != "READY" and p.runs:
            _marcar(p.runs[0], block.status)
            p.runs[0].text = prefixo + p.runs[0].text
        _nota(doc, block.note)

    elif kind == "paragraph":
        p = doc.add_paragraph()
        r = p.add_run(prefixo + (block.text or ""))
        _marcar(r, block.status)
        _nota(doc, block.note)

    elif kind == "table":
        _render_table(doc, block)

    elif kind == "figure_placeholder":
        p = doc.add_paragraph()
        r = p.add_run(block.label or "[FIGURA — RENDERIZAÇÃO PENDENTE]")
        _marcar(r, block.status if block.status != "READY" else "PLACEHOLDER")
        _nota(doc, block.note)

    elif kind == "caption":
        p = doc.add_paragraph()
        r = p.add_run(block.text or "")
        r.italic = True
        _marcar(r, block.status)

    elif kind == "list":
        for item in (block.items or []):
            p = doc.add_paragraph(style="List Bullet")
            r = p.add_run(prefixo + item)
            _marcar(r, block.status)
        _nota(doc, block.note)

    elif kind == "toc":
        for item in (block.items or []):
            doc.add_paragraph(item)

    elif kind == "page_break":
        doc.add_page_break()

    elif kind == "cover_field":
        p = doc.add_paragraph()
        r = p.add_run(block.text or "")
        _marcar(r, block.status)

    else:
        raise ValueError("document_model.Block.kind desconhecido para o docx_renderer: %r" % kind)


def render(document, output_path):
    """`document`: gerador.document_model.Document. Grava em `output_path`
    (sempre dentro de output/shadow/, nunca no SharePoint -- assert_read_only
    protege contra acidente). Devolve output_path."""
    U.assert_read_only(output_path)

    doc = DocxDocument()
    _setup_estilo(doc)

    primeira = True
    perfil_atual = None
    for secao in document.sections:
        perfil_secao = (secao.page_setup, getattr(secao, "page_profile", "body"))
        if primeira:
            _set_page(doc.sections[0], *perfil_secao)
            perfil_atual = perfil_secao
            primeira = False
        elif perfil_secao != perfil_atual:
            # nova secao fisica do Word sempre que papel OU perfil mudar --
            # e' isso que separa capa/corpo/fotografico/corpo (P0.6) mesmo
            # quando dois deles compartilham o mesmo papel (A4).
            nova = doc.add_section(WD_SECTION.NEW_PAGE)
            _set_page(nova, *perfil_secao)
            perfil_atual = perfil_secao
        else:
            doc.add_page_break()

        for block in secao.blocks:
            _render_block(doc, block)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path
