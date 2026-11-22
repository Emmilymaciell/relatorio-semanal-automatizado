# -*- coding: utf-8 -*-
r"""
scripts/gerador/figuras/nc_graficos.py - NC: figura de notificacoes +
Grafico 2 (por status) + Grafico 3 (por natureza). Com PIL (spec Sec.2.5).

Fonte principal: `report_data["nonconformities"]` (ja extraido por
`scripts/gerador/extratores/nc.py` de `02_Acompanhamento de NC
automatizada.xlsm`, aba VISAO_GERAL, `Tbl_VISAO`). O Grafico 3 ("por
natureza") usa um campo (coluna `Natureza`) que o extrator anterior nao
capturava -- lido aqui diretamente do mesmo arquivo/aba/tabela (mesma fonte,
nenhuma nova planilha), caminho vindo do manifest
(`sources.nc_tracker.path`), nunca hardcoded.

`history_by_status` e o estado ATUAL da planilha (ela usa `TODAY()`), nao um
historico reconstruido da semana -- mantido como `SOURCE_LIMITATION`,
propagada para a nota da Figura/Grafico 2 (nunca forcado a bater com o
golden, spec Sec.2.5).
"""
import re
from collections import Counter

import openpyxl
from openpyxl.utils import column_index_from_string as cix

from .. import config as C  # noqa: E402
from . import (FiguraResult, mm_to_px, get_font, wrap_text, novo_canvas,
               salvar_png, desenhar_borda, text_width, COR_BORDA, COR_PRETO,
               COR_BRANCO, COR_SHADING_CABECALHO, PALETA_GRAFICOS, U,
               assert_nao_e_golden)

FIGURA_NOTIF_MM = (160.0, 63.82)
GRAFICO2_MM = (120.82, 61.48)
GRAFICO3_MM = (111.25, 67.37)

_CORES_CICLO = [PALETA_GRAFICOS[k] for k in
               ("azul_escuro", "azul_claro", "verde", "amarelo", "vermelho",
                "roxo", "laranja", "cinza_grade")]


# -----------------------------------------------------------------------------
# Figura de notificacoes (issued_in_period)
# -----------------------------------------------------------------------------
def gerar_figura_notificacoes(report_data, out_dir) -> FiguraResult:
    largura_mm, altura_mm = FIGURA_NOTIF_MM
    nc = (report_data or {}).get("nonconformities", {})
    issued = nc.get("issued_in_period")
    if issued is None:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION",
                            "nonconformities.issued_in_period ausente em report_data",
                            {"source_file": "report_data.json"})

    img, draw = novo_canvas(largura_mm, altura_mm)
    w_px, h_px = img.size
    fonte_titulo = get_font(10, bold=True)
    fonte_cab = get_font(8, bold=True)
    fonte = get_font(7.5)

    titulo_h = mm_to_px(6)
    draw.text((mm_to_px(1.5), mm_to_px(1)), "Notificações emitidas no período",
              font=fonte_titulo, fill=COR_PRETO)

    colunas = [("Documento", 0.14), ("Data de\nAbertura", 0.15), ("Status", 0.17),
              ("Descrição", 0.54)]
    larguras_px = [mm_to_px(largura_mm * p) for _, p in colunas]
    larguras_px[-1] += w_px - sum(larguras_px)

    cab_h = mm_to_px(9)
    y0 = titulo_h
    draw.rectangle([0, y0, w_px, y0 + cab_h], fill=COR_SHADING_CABECALHO())
    x = 0
    linha_cab_px = mm_to_px(3.2)
    for (nome, _p), wpx in zip(colunas, larguras_px):
        linhas_nome = nome.split("\n")
        ty = y0 + (cab_h - len(linhas_nome) * linha_cab_px) / 2
        for lt in linhas_nome:
            draw.text((x + mm_to_px(1), ty), lt, font=fonte_cab, fill=COR_PRETO)
            ty += linha_cab_px
        x += wpx

    corpo_y0 = y0 + cab_h
    if not issued:
        draw.text((mm_to_px(1.5), corpo_y0 + mm_to_px(3)),
                  "Nenhuma notificação de não conformidade emitida no período do relatório.",
                  font=fonte, fill=COR_PRETO)
    else:
        linha_h = (h_px - corpo_y0) / len(issued)
        y = corpo_y0
        for item in issued:
            data_br = item.get("data_abertura", "")
            if data_br and "-" in data_br:
                p = data_br.split("-")
                data_br = "%s/%s/%s" % (p[2], p[1], p[0])
            valores = [item.get("documento") or "", data_br, item.get("status") or "",
                      item.get("sentence") or item.get("irregularidade") or ""]
            x = 0
            for valor, wpx in zip(valores, larguras_px):
                ls = wrap_text(draw, str(valor), fonte, wpx - mm_to_px(2))
                draw.text((x + mm_to_px(1), y + mm_to_px(0.8)), ls[0], font=fonte, fill=COR_PRETO)
                x += wpx
            y += linha_h
            draw.line([(0, y), (w_px, y)], fill=COR_BORDA, width=1)

    desenhar_borda(draw, (0, y0, w_px - 1, h_px - 1), 1, COR_BORDA)
    x = 0
    for wpx in larguras_px[:-1]:
        x += wpx
        draw.line([(x, y0), (x, h_px)], fill=COR_BORDA, width=1)
    draw.line([(0, corpo_y0), (w_px, corpo_y0)], fill=COR_BORDA, width=1)

    caminho = salvar_png(img, out_dir, "figura9_notificacoes.png")
    prov = {
        "source_file": "report_data.json",
        "inputs": ["nonconformities.issued_in_period"],
        "method": "PIL: tabela nativa das NC emitidas no periodo (vazia = fato real, nao erro)",
        "n_notificacoes": len(issued),
    }
    return FiguraResult(caminho, largura_mm, altura_mm, "OK", None, prov)


# -----------------------------------------------------------------------------
# Grafico 2 - por status (history_by_status, ja em report_data)
# -----------------------------------------------------------------------------
def _grafico_pizza(titulo, contagem: Counter, largura_mm, altura_mm, out_dir, nome_arquivo,
                   nota_fonte=None):
    total = sum(contagem.values())
    img, draw = novo_canvas(largura_mm, altura_mm)
    w_px, h_px = img.size
    fonte_titulo = get_font(9.5, bold=True)
    fonte_legenda = get_font(7.5)

    titulo_h = mm_to_px(6)
    draw.text((mm_to_px(1), mm_to_px(0.5)), titulo, font=fonte_titulo, fill=COR_PRETO)
    rodape_h = mm_to_px(4.5) if nota_fonte else 0

    diam_px = min(w_px * 0.55, h_px - titulo_h - rodape_h - mm_to_px(2))
    cx = diam_px / 2 + mm_to_px(2)
    cy = titulo_h + (h_px - rodape_h - titulo_h) / 2
    bbox = [cx - diam_px / 2, cy - diam_px / 2, cx + diam_px / 2, cy + diam_px / 2]

    itens = [(k, v) for k, v in contagem.items() if v > 0]
    itens.sort(key=lambda kv: -kv[1])
    if total == 0:
        draw.ellipse(bbox, outline=COR_BORDA, width=2)
        draw.text((cx - mm_to_px(10), cy - mm_to_px(2)), "sem dados", font=fonte_legenda,
                  fill=COR_PRETO)
    else:
        ang0 = -90.0
        for i, (_status, n) in enumerate(itens):
            ang = 360.0 * n / total
            cor = _CORES_CICLO[i % len(_CORES_CICLO)]
            draw.pieslice(bbox, ang0, ang0 + ang, fill=cor, outline=COR_BRANCO)
            ang0 += ang

    lx = cx + diam_px / 2 + mm_to_px(6)
    ly = titulo_h + mm_to_px(1)
    quad = mm_to_px(2.6)
    for i, (status, n) in enumerate(itens):
        cor = _CORES_CICLO[i % len(_CORES_CICLO)]
        draw.rectangle([lx, ly, lx + quad, ly + quad], fill=cor)
        pct = 100.0 * n / total if total else 0
        rot = "%s: %d (%.0f%%)" % (status, n, pct)
        draw.text((lx + quad + mm_to_px(1.2), ly - mm_to_px(0.5)), rot, font=fonte_legenda,
                  fill=COR_PRETO)
        ly += mm_to_px(4.2)
        if ly > h_px - mm_to_px(4):
            lx += mm_to_px(38)
            ly = titulo_h + mm_to_px(1)

    if nota_fonte:
        fonte_nota = get_font(6)
        draw.text((mm_to_px(1), h_px - mm_to_px(3.5)), nota_fonte, font=fonte_nota,
                  fill=COR_BORDA)

    caminho = salvar_png(img, out_dir, nome_arquivo)
    return caminho


def gerar_grafico_status(report_data, out_dir) -> FiguraResult:
    largura_mm, altura_mm = GRAFICO2_MM
    nc = (report_data or {}).get("nonconformities", {})
    hist = nc.get("history_by_status")
    if not hist:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION",
                            "nonconformities.history_by_status ausente", {"source_file": "report_data.json"})

    contagem = Counter({k: v for k, v in hist.items() if isinstance(v, int)})
    rodovia = ((report_data or {}).get("road") or {}).get("rodovia") or "?"
    caminho = _grafico_pizza("NC por status (%s)" % rodovia, contagem, largura_mm, altura_mm, out_dir,
                             "grafico2_nc_por_status.png",
                             nota_fonte="fonte: Tbl_VISAO (estado atual -- TODAY())")
    prov = {
        "source_file": "report_data.json", "inputs": ["nonconformities.history_by_status"],
        "method": "PIL: pizza por status, cores por ciclo fixo",
        "confidence": "SOURCE_LIMITATION: a planilha usa TODAY() -- estado atual, nao "
                      "historico reconstruido da semana do relatorio (mesma ressalva de "
                      "extratores/nc.py); nunca forcado a bater com o golden",
    }
    return FiguraResult(caminho, largura_mm, altura_mm, "OK",
                        "history_by_status e o estado ATUAL da planilha (TODAY()), nao o "
                        "historico da semana -- SOURCE_LIMITATION propagada", prov)


# -----------------------------------------------------------------------------
# Grafico 3 - por natureza (precisa reabrir o tracker; coluna nao capturada
# pelo extrator anterior)
# -----------------------------------------------------------------------------
def _ler_tabela(ws, nome_tabela):
    if nome_tabela not in ws.tables:
        return None, None
    ref = ws.tables[nome_tabela].ref
    m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
    c1, r1, c2, r2 = cix(m.group(1)), int(m.group(2)), cix(m.group(3)), int(m.group(4))
    headers = [ws.cell(r1, c).value for c in range(c1, c2 + 1)]
    linhas = []
    for r in range(r1 + 1, r2 + 1):
        row, vazio = {}, True
        for i, c in enumerate(range(c1, c2 + 1)):
            v = ws.cell(r, c).value
            row[headers[i]] = v
            if v not in (None, ""):
                vazio = False
        if not vazio:
            linhas.append(row)
    return headers, linhas


def gerar_grafico_natureza(caminho_nc_tracker, rodovia, out_dir) -> FiguraResult:
    largura_mm, altura_mm = GRAFICO3_MM
    prov = {"source_file": caminho_nc_tracker, "sheet": "VISAO_GERAL", "tabela": "Tbl_VISAO"}
    if not caminho_nc_tracker or not U.exists(caminho_nc_tracker):
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION",
                            "arquivo do tracker de NC nao encontrado", prov)
    assert_nao_e_golden(caminho_nc_tracker)

    local = U.copiar_para_tmp(caminho_nc_tracker, "figuras_nc", "nc_tracker.xlsm")
    wb = openpyxl.load_workbook(local, data_only=True)
    if "VISAO_GERAL" not in wb.sheetnames:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION",
                            "aba VISAO_GERAL ausente", prov)
    ws = wb["VISAO_GERAL"]
    headers, linhas = _ler_tabela(ws, "Tbl_VISAO")
    if linhas is None:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION",
                            "tabela Tbl_VISAO ausente", prov)
    if "Natureza" not in headers:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION",
                            "coluna 'Natureza' ausente no cabecalho da Tbl_VISAO", prov)

    alvo_rod = C.norm_rodovia(rodovia)
    linhas_rodovia = [r for r in linhas if C.norm_rodovia(r.get("Rodovia")) == alvo_rod]
    contagem = Counter(str(r.get("Natureza") or "(vazio)").strip() for r in linhas_rodovia)
    if not contagem:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION",
                            "nenhuma linha de %s na Tbl_VISAO" % rodovia, prov)

    caminho = _grafico_pizza("NC por natureza (%s)" % rodovia, contagem, largura_mm, altura_mm,
                             out_dir, "grafico3_nc_por_natureza.png")
    prov.update({
        "method": ("PIL: pizza por 'Natureza' (coluna real da Tbl_VISAO, nao capturada pelo "
                  "extrator scripts/gerador/extratores/nc.py -- lida aqui do mesmo arquivo/"
                  "aba/tabela); filtro Rodovia == '%s'" % rodovia),
        "n_linhas": len(linhas_rodovia),
    })
    return FiguraResult(caminho, largura_mm, altura_mm, "OK", None, prov)
