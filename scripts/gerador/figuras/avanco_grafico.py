# -*- coding: utf-8 -*-
r"""
scripts/gerador/figuras/avanco_grafico.py - Gráfico 1: Avanço Global,
barras agrupadas por semana (Concluído x Concluído + Atacado), com PIL.

Fonte: `report_data["progress"]["series"]["relatorio_semanal"]` (ja extraida
por `scripts/gerador/extratores/avanco.py` a partir de `Grafico - Avanco
Global.xlsx`, aba 'Relatorio Semanal' -- spec Sec.2.2). Nao le planilha
diretamente: os 4 pontos ja vem prontos em `report_data.json`.

Dimensao: o GOLDEN_VISUAL_SPEC.md Sec.6 mede Figuras 6-9 e Graficos 2-3, mas
NAO mede o Grafico 1 -- so registra a regra geral ("grafico fica centralizado
a ~70-75% da largura util"). Sem medida direta, usamos 120 x 75 mm (75% de
160mm de largura, altura de aspecto de grafico de barras usual). Isso e uma
ESCOLHA DE DESIGN documentada aqui e na proveniencia, nao um valor medido no
golden -- reportado como tal (nao e uma contradicao, e uma lacuna: o golden
nao tem esse grafico dimensionado na spec de descoberta).
"""
from . import (FiguraResult, mm_to_px, get_font, novo_canvas, salvar_png,
               text_width, PALETA_GRAFICOS, COR_PRETO, COR_BRANCO)

LARGURA_MM = 120.0
ALTURA_MM = 75.0

COR_CONCLUIDO = PALETA_GRAFICOS["azul_escuro"]
COR_ATACADO = PALETA_GRAFICOS["azul_claro"]


def _fmt_data_curta(iso):
    # "2026-08-14" -> "14/08"
    try:
        partes = iso.split("-")
        return "%s/%s" % (partes[2], partes[1])
    except Exception:
        return iso


def gerar(report_data, out_dir) -> FiguraResult:
    prov_base = {"source_file": "report_data.json", "method": None}
    progress = (report_data or {}).get("progress", {})
    series = progress.get("series", {}).get("relatorio_semanal", {})
    linhas = series.get("linhas") if series.get("status") == "OK" else None

    if not linhas:
        return FiguraResult(None, LARGURA_MM, ALTURA_MM, "LIMITATION",
                            "progress.series.relatorio_semanal indisponivel ou vazio",
                            prov_base)

    img, draw = novo_canvas(LARGURA_MM, ALTURA_MM)
    w_px, h_px = img.size

    margem_esq = mm_to_px(14)
    margem_dir = mm_to_px(4)
    margem_topo = mm_to_px(6)
    margem_base = mm_to_px(12)
    area_x0, area_y0 = margem_esq, margem_topo
    area_x1, area_y1 = w_px - margem_dir, h_px - margem_base
    area_w, area_h = area_x1 - area_x0, area_y1 - area_y0

    fonte_eixo = get_font(7)
    fonte_valor = get_font(7, bold=True)
    fonte_legenda = get_font(7.5)

    y_max = 100.0
    # linhas de grade horizontais a cada 20%
    for pct in (0, 20, 40, 60, 80, 100):
        y = area_y1 - area_h * (pct / y_max)
        draw.line([(area_x0, y), (area_x1, y)], fill=PALETA_GRAFICOS["cinza_grade"], width=1)
        rot = "%d%%" % pct
        tw = text_width(draw, rot, fonte_eixo)
        draw.text((area_x0 - tw - mm_to_px(1.5), y - mm_to_px(1.5)), rot,
                  font=fonte_eixo, fill=COR_PRETO)
    draw.line([(area_x0, area_y0), (area_x0, area_y1)], fill=COR_PRETO, width=2)
    draw.line([(area_x0, area_y1), (area_x1, area_y1)], fill=COR_PRETO, width=2)

    n = len(linhas)
    grupo_w = area_w / n
    barra_w = grupo_w * 0.30
    gap_barras = grupo_w * 0.06

    for i, linha in enumerate(linhas):
        cx = area_x0 + grupo_w * (i + 0.5)
        concl = float(linha.get("concluido_pct") or 0)
        atac = float(linha.get("concluido_atacado_pct") or 0)

        for offset, valor, cor_barra in (
            (-1, concl, COR_CONCLUIDO), (1, atac, COR_ATACADO),
        ):
            bx0 = cx + offset * (gap_barras / 2 + (barra_w if offset > 0 else 0))
            bx1 = bx0 + (barra_w if offset > 0 else -barra_w)
            x0, x1 = min(bx0, bx1), max(bx0, bx1)
            by = area_y1 - area_h * (valor / y_max)
            draw.rectangle([x0, by, x1, area_y1], fill=cor_barra)
            rot = ("%.2f%%" % valor).replace(".", ",")
            tw = text_width(draw, rot, fonte_valor)
            draw.text(((x0 + x1) / 2 - tw / 2, by - mm_to_px(3.6)), rot,
                      font=fonte_valor, fill=COR_PRETO)

        data_rot = _fmt_data_curta(linha.get("data", ""))
        tw = text_width(draw, data_rot, fonte_eixo)
        draw.text((cx - tw / 2, area_y1 + mm_to_px(1.5)), data_rot,
                  font=fonte_eixo, fill=COR_PRETO)

    # legenda
    ly = mm_to_px(1.0)
    lx = area_x0
    quad = mm_to_px(2.6)
    draw.rectangle([lx, ly, lx + quad, ly + quad], fill=COR_CONCLUIDO)
    draw.text((lx + quad + mm_to_px(1.2), ly - mm_to_px(0.6)), "Concluído",
              font=fonte_legenda, fill=COR_PRETO)
    lx2 = lx + quad + mm_to_px(1.2) + text_width(draw, "Concluído", fonte_legenda) + mm_to_px(4)
    draw.rectangle([lx2, ly, lx2 + quad, ly + quad], fill=COR_ATACADO)
    draw.text((lx2 + quad + mm_to_px(1.2), ly - mm_to_px(0.6)), "Concluído + Atacado",
              font=fonte_legenda, fill=COR_PRETO)

    caminho = salvar_png(img, out_dir, "grafico1_avanco_global.png")

    provenance = dict(prov_base)
    provenance.update({
        "inputs": ["progress.series.relatorio_semanal.linhas"],
        "method": ("PIL: barras agrupadas por semana (Concluido / Concluido+Atacado), "
                   "rotulo de valor sobre a barra, eixo de datas dd/mm; dimensao 120x75mm "
                   "e escolha de design (golden nao mede este grafico especificamente, so "
                   "a regra geral de ~70-75%% da largura util do corpo)"),
        "confidence": "ALTA CONFIANÇA (valores diretos de report_data; dimensao nao medida no golden)",
    })
    return FiguraResult(caminho, LARGURA_MM, ALTURA_MM, "OK", None, provenance)
