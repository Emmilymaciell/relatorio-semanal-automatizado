# -*- coding: utf-8 -*-
r"""
scripts/gerador/annex_embed.py - Fase 3.3, Sec.4: anexos EMBEDDED_IMAGE
(Parecer Técnico/PA-NNN, Acompanhamento Ambiental, Acompanhamento SST) viram
IMAGEM DE PÁGINA CHEIA dentro do corpo do .docx, uma página do documento por
página REAL do PDF-fonte (`fitz`, 200 dpi), na ORDEM que `assembly_model.
montar_plano()` já calculou a partir de `config/annex_types.yaml` (nunca por
posição/número aqui).

Este módulo é camada de MONTAGEM/PIPELINE (mesmo nível de assembly_model.py),
não de CONTEÚDO: `content_builder.py` nunca abre arquivo nem sabe de PDF; é
`shadow_report.py` (comando `candidate`) quem chama `montar_blocos_embutidos()`
depois de `construir_documento()` e ANEXA os blocos resultantes na Section
"s11_anexos" já construída -- nunca dentro do próprio content_builder.

Os anexos MERGED_PDF (Diagrama Unifilar, Diário de Obras) NUNCA passam por
aqui -- ficam fora do Word (ver assembly_model.classificar_anexo).
"""
import os
import re
from collections import Counter

import fitz  # PyMuPDF

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(AQUI)
import sys  # noqa: E402
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402

from . import document_model as M  # noqa: E402
from . import assert_nao_e_golden  # noqa: E402

DPI_PADRAO = 200
_PT_PARA_MM = 25.4 / 72.0

# Correção 1 (Fase 3.3): rodapé interno "Pág. N/M" que o PRÓPRIO documento
# (ex.: PA-254 - Parecer.pdf) traz nas suas paginas -- tolerante a "Pág."/
# "Pag"/espacos ao redor da barra (verificado no PA-254 real: "Pág.1 / 6" ..
# "Pág.6 / 6", paginas 7-10 SEM esse rodape porque sao fichas FA-ambiental/
# FA-SST anexadas DENTRO do mesmo PDF, nao paginas do parecer).
_RE_RODAPE_PAGINA = re.compile(r"P[aá]g\.?\s*(\d+)\s*/\s*(\d+)", re.IGNORECASE)
_FAIXA_RODAPE = 0.20  # fracao inferior da pagina onde o rodape e procurado


def _detectar_paginas_proprias(doc, faixa_rodape=_FAIXA_RODAPE):
    """Correção 1: para um anexo EMBEDDED_IMAGE, usa SOMENTE as páginas do
    documento -- fonte -- ELE MESMO, nunca páginas de outro documento coladas
    dentro do mesmo arquivo PDF. Detecção determinística pelo rodapé interno
    "Pág. N/M" (regex tolerante, ver `_RE_RODAPE_PAGINA`), lido da faixa
    inferior (`faixa_rodape` da altura) de cada página.

    Devolve `(proprios, descartados, motivo)`, ambos listas de índices 0-based:
      - nenhuma página com o rodapé -> usa TODAS (`descartados=[]`), nunca
        corta as páginas de um anexo que simplesmente não tem esse rodapé
        (ex.: Acompanhamento Ambiental/SST, cada um já é seu próprio PDF de
        2 páginas, sem essa numeração interna);
      - alguma página com o rodapé -> o denominador M mais frequente entre
        os achados define "o documento próprio"; as páginas SEM o rodapé, ou
        com M diferente, são descartadas (nunca renderizadas)."""
    achados = {}
    for i in range(doc.page_count):
        page = doc[i]
        h = page.rect.height
        clip = fitz.Rect(0, h * (1 - faixa_rodape), page.rect.width, h)
        texto = page.get_text("text", clip=clip)
        m = _RE_RODAPE_PAGINA.search(texto)
        if m:
            achados[i] = (int(m.group(1)), int(m.group(2)))

    if not achados:
        return (list(range(doc.page_count)), [],
               "rodapé interno 'Pág. N/M' não encontrado em nenhuma página -- "
               "usando todas as %d páginas do PDF-fonte" % doc.page_count)

    m_dominante, freq = Counter(m for _n, m in achados.values()).most_common(1)[0]
    proprios = sorted(i for i, (_n, m) in achados.items() if m == m_dominante)
    descartados = [i for i in range(doc.page_count) if i not in proprios]
    motivo = ("rodapé interno 'Pág. N/%d' identifica as páginas do próprio documento "
             "(%d de %d páginas do PDF-fonte, achado em %d página(s)); as demais não "
             "trazem esse rodapé -- são conteúdo anexado DENTRO do mesmo arquivo PDF "
             "(ex.: fichas de Acompanhamento Ambiental/SST coladas ao final do Parecer), "
             "nunca páginas do próprio anexo" % (m_dominante, len(proprios), doc.page_count, freq))
    return proprios, descartados, motivo


def renderizar_paginas_anexo(source_path, output_dir, prefixo, *, dpi=DPI_PADRAO):
    """Renderiza a `dpi` em PNG SOMENTE as páginas PRÓPRIAS de `source_path`
    (PDF) -- Correção 1: detectadas por `_detectar_paginas_proprias()` via o
    rodapé interno "Pág. N/M" do próprio documento; páginas anexadas dentro
    do mesmo PDF-fonte (sem esse rodapé, ou com M diferente) são descartadas,
    nunca renderizadas/embutidas de novo. Grava em
    `output_dir/{prefixo}_pNN.png` (NN é a posição dentro do documento
    PRÓPRIO, não o índice bruto do PDF). Devolve `(paginas, corte)`:

      paginas: lista de dicts {path, width_mm, height_mm, page_index,
               n_paginas_total} -- a dimensão física vem SEMPRE do próprio
               PDF (MediaBox), nunca um valor fixo.
      corte:   {n_paginas_pdf, paginas_usadas, paginas_descartadas, motivo}
               -- proveniência do corte (1-based), para o manifesto/relate."""
    U.assert_read_only(output_dir)  # nunca grava dentro do SharePoint
    assert_nao_e_golden(source_path)  # nunca embute o relatorio oficial como "anexo"
    os.makedirs(output_dir, exist_ok=True)
    doc = fitz.open(U.lp(source_path))
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    paginas = []
    try:
        n_pdf = doc.page_count
        proprios, descartados, motivo = _detectar_paginas_proprias(doc)
        n_proprias = len(proprios)
        for pos, i in enumerate(proprios):
            page = doc[i]
            largura_mm = page.rect.width * _PT_PARA_MM
            altura_mm = page.rect.height * _PT_PARA_MM
            pix = page.get_pixmap(matrix=mat, alpha=False)
            nome = "%s_p%02d.png" % (prefixo, pos + 1)
            caminho = os.path.join(output_dir, nome)
            pix.save(caminho)
            paginas.append({
                "path": caminho, "width_mm": largura_mm, "height_mm": altura_mm,
                "page_index": pos, "n_paginas_total": n_proprias,
            })
    finally:
        doc.close()
    corte = {
        "n_paginas_pdf": n_pdf,
        "paginas_usadas": [i + 1 for i in proprios],
        "paginas_descartadas": [i + 1 for i in descartados],
        "motivo": motivo,
    }
    return paginas, corte


def _escalar_para_caber(largura_mm, altura_mm, usable_width_mm, usable_height_mm):
    """Escala (mantendo proporção) para caber dentro da área útil do corpo
    A4 -- 'página cheia' quer dizer o MAIOR tamanho possível dentro das
    margens do relatório (o cabeçalho/rodapé continuam visíveis, herdados
    da seção -- Sec.4), nunca ultrapassando a área útil nem distorcendo."""
    escala = usable_width_mm / largura_mm
    if altura_mm * escala > usable_height_mm:
        escala = usable_height_mm / altura_mm
    return round(largura_mm * escala, 2), round(altura_mm * escala, 2)


def montar_blocos_embutidos(embedded_annexes, output_dir, usable_width_mm, usable_height_mm,
                            *, dpi=DPI_PADRAO):
    """`embedded_annexes`: lista de `assembly_model.EmbeddedAnnexPlan` (já
    ordenada por `ordem` -- `montar_plano()` faz isso). Para cada anexo
    PRESENTE cuja fonte seja um PDF localizável, renderiza as páginas reais
    (fitz) e devolve `document_model.Block`s (`page_break` + `full_page_image`
    por página), na MESMA ordem dos anexos. Anexo ausente ou cuja fonte não
    seja PDF vira um único parágrafo LIMITATION -- nunca finge a página.

    Devolve `(blocos, resumo)`; `resumo` é uma lista de dicts
    {anexo_id, paginas, status} para o manifesto/relatório (Sec.6/relate)."""
    blocos = []
    resumo = []
    for anexo in embedded_annexes:
        if not anexo.presente or not anexo.source_path:
            blocos.append(M.page_break())
            blocos.append(M.paragraph(
                "Anexo “%s” não pôde ser embutido nesta rodada: fonte ausente "
                "ou não confirmada." % anexo.titulo,
                status="LIMITATION", note="anexo_id=%s presente=%s source_path=%s"
                    % (anexo.anexo_id, anexo.presente, anexo.source_path)))
            resumo.append({"anexo_id": anexo.anexo_id, "titulo": anexo.titulo,
                           "paginas": 0, "status": "AUSENTE"})
            continue

        if not U.exists(anexo.source_path) or U.isdir(anexo.source_path):
            blocos.append(M.page_break())
            blocos.append(M.paragraph(
                "Anexo “%s” não pôde ser embutido: caminho-fonte não encontrado "
                "(%s)." % (anexo.titulo, anexo.source_path),
                status="LIMITATION", note="anexo_id=%s" % anexo.anexo_id))
            resumo.append({"anexo_id": anexo.anexo_id, "titulo": anexo.titulo,
                           "paginas": 0, "status": "FONTE_NAO_ENCONTRADA"})
            continue

        if os.path.splitext(anexo.source_path)[1].lower() != ".pdf":
            blocos.append(M.page_break())
            blocos.append(M.paragraph(
                "Anexo “%s” não pôde ser embutido: a fonte não é um PDF "
                "renderizável (%s)." % (anexo.titulo, anexo.source_path),
                status="LIMITATION", note="anexo_id=%s source_path=%s"
                    % (anexo.anexo_id, anexo.source_path)))
            resumo.append({"anexo_id": anexo.anexo_id, "titulo": anexo.titulo,
                           "paginas": 0, "status": "FONTE_NAO_PDF"})
            continue

        paginas, corte = renderizar_paginas_anexo(anexo.source_path, output_dir, anexo.anexo_id, dpi=dpi)
        for pg in paginas:
            largura_mm, altura_mm = _escalar_para_caber(
                pg["width_mm"], pg["height_mm"], usable_width_mm, usable_height_mm)
            nota = ("página %d/%d de %s (fitz %ddpi) -- fonte: %s"
                    % (pg["page_index"] + 1, pg["n_paginas_total"], anexo.titulo, dpi,
                       anexo.source_path))
            if corte["paginas_descartadas"]:
                nota += " -- CORTE (Correção 1): %s" % corte["motivo"]
            blocos.append(M.page_break())
            blocos.append(M.full_page_image(
                pg["path"], width_mm=largura_mm, height_mm=altura_mm, status="READY",
                provenance=["annexes.%s" % anexo.anexo_id],
                label="[ANEXO %s — %s, pág. %d/%d]"
                     % (anexo.anexo_id, anexo.titulo, pg["page_index"] + 1, pg["n_paginas_total"]),
                note=nota))
        resumo.append({"anexo_id": anexo.anexo_id, "titulo": anexo.titulo,
                       "paginas": len(paginas), "status": "OK",
                       "n_paginas_pdf_fonte": corte["n_paginas_pdf"],
                       "paginas_usadas": corte["paginas_usadas"],
                       "paginas_descartadas": corte["paginas_descartadas"],
                       "motivo_corte": corte["motivo"]})
    return blocos, resumo
