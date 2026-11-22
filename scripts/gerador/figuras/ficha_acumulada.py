# -*- coding: utf-8 -*-
r"""
scripts/gerador/figuras/ficha_acumulada.py - Figura 1: Ficha de Acompanhamento
(tabela de estacas cumulativa), desenhada nativamente com PIL.

Fonte: `Arquivos Auxiliares/Ficha de Campo Acumulada GO-<n>.xlsx`, aba
`FICHA CAMPO` (spec Sec.2.1). O caminho vem do manifest/report_data (nunca
hardcoded), resolvido pelo chamador via `config/source_layout.yaml` ->
`auxiliares.ficha_acumulada`.

Layout real do arquivo (VERIFICADO nesta sessao, RD103/49): a tabela linear de
estacas comeca no cabecalho "Estaca inicial / Estaca final / Lado / Extensao
(km) / Servico / Situacao / Descricao / Observacoes" e termina na primeira
linha em branco -- depois dela vem o bloco dedicado OAC/OAE (outra tabela, com
outro cabecalho), que NAO faz parte desta figura (ela tem estrutura propria:
Escavacao/Corpo/Ala/Compactacao, nao "estaca a estaca" no mesmo sentido).
O cabecalho e localizado por texto normalizado, nunca por linha fixa (a linha
varia por obra/semana), para nao repetir o erro ja documentado em
`config/source_layout.yaml` (a spec antiga tinha a linha errada).
"""
import os
import re

import openpyxl
from PIL import Image, ImageDraw

from . import (FiguraResult, mm_to_px, pt_to_px, get_font, wrap_text, novo_canvas,
               salvar_png, desenhar_borda, COR_BORDA, COR_BRANCO, COR_PRETO,
               COR_SHADING_CABECALHO, U, assert_nao_e_golden, C, text_width)

LARGURA_MM = 160.0
# Correção 2 (Fase 3.3): área útil de altura do corpo A4 (GOLDEN_VISUAL_SPEC.md
# Sec.1.1 -- 247mm) -- a Figura 1 tem de caber numa página só, mantendo a
# largura útil (160mm) e a proporção/legibilidade da tabela (nunca reduzindo
# largura à custa de distorcer as colunas). Se mesmo compactada a altura
# exceder isto, o status vira LIMITATION (nunca corta linha sem registrar).
ALTURA_UTIL_A4_MM = 247.0
_ROTULOS_HEADER = {
    "estaca inicial": "estaca_inicial", "estaca final": "estaca_final",
    "lado": "lado", "extensao (km)": "extensao_km", "servico": "servico_codigo",
    "situacao": "situacao", "descricao": "descricao", "observacoes": "observacao",
}

# larguras proporcionais (mm), somando 160
_COLUNAS = [
    ("Estaca\nInicial", "estaca_inicial", 17.0),
    ("Estaca\nFinal", "estaca_final", 17.0),
    ("Lado", "lado", 9.0),
    ("Ext.\n(km)", "extensao_km", 12.0),
    ("Serviço", "servico_codigo", 23.0),
    ("Situação", "situacao", 16.0),
    ("Descrição", "descricao", 38.0),
    ("Observações", "observacao", 28.0),
]

LINHA_ALTURA_MIN_MM = 3.8
CABECALHO_ALTURA_MM = 8.0
FONTE_PT = 7.0
FONTE_CABECALHO_PT = 7.0
PAD_MM = 0.6


def _n(v):
    return U.norm(v) if v is not None else ""


_RE_ESTACA_I = re.compile(r"^estaca\s*i(nicial|\.)?$")
_RE_ESTACA_F = re.compile(r"^estaca\s*f(inal|\.)?$")


def _rotulo_canonico(v):
    """Rotulo normalizado -> rotulo canonico de _ROTULOS_HEADER ('' se nao
    reconhecido). Aceita 'Estaca i'/'Estaca f' (RD104) como inicial/final."""
    if v in _ROTULOS_HEADER:
        return v
    if _RE_ESTACA_I.match(v):
        return "estaca inicial"
    if _RE_ESTACA_F.match(v):
        return "estaca final"
    return ""


def _achar_cabecalho(ws, varredura_linhas=40, varredura_cols=40):
    for r in range(1, varredura_linhas + 1):
        achados = {}
        for c in range(1, varredura_cols + 1):
            v = _n(ws.cell(r, c).value)
            canon = _rotulo_canonico(v)
            if canon:
                achados[canon] = c
        if "estaca inicial" in achados and "servico" in achados:
            return r, achados
    return None, {}


def _estaca_str(km, sep, m):
    if km is None:
        return ""
    sep = sep if sep not in (None, "") else "+"
    m = m if m is not None else 0
    return "%s%s%s" % (km, sep, m)


def _ler_tabela(ws, header_row, col_map):
    """Le linhas a partir de header_row+1 ate a primeira linha totalmente
    vazia (fim da tabela linear -- o bloco OAC/OAE dedicado comeca depois,
    com outro cabecalho, e nao entra aqui)."""
    c_ei = col_map["estaca inicial"]
    c_ef = col_map.get("estaca final")
    c_serv = col_map["servico"]
    linhas = []
    r = header_row + 1
    while True:
        km_i = ws.cell(r, c_ei).value
        serv = ws.cell(r, c_serv).value
        # 3 colunas consecutivas por estaca (km, separador, m) -- padrao
        # identico ao da Ficha de Campo diaria (config/source_layout.yaml).
        # Para na PRIMEIRA linha totalmente vazia: o bloco OAC/OAE dedicado
        # (outra tabela, outro cabecalho) comeca logo depois de uma linha em
        # branco (VERIFICADO no arquivo real, RD103/49: linha 59 em branco
        # entre a tabela linear e o bloco OAE/OAC) -- nao faz parte desta
        # figura, entao nao se pode "pular" a linha em branco e continuar.
        if km_i in (None, "") and serv in (None, ""):
            break
        ei = _estaca_str(ws.cell(r, c_ei).value, ws.cell(r, c_ei + 1).value,
                         ws.cell(r, c_ei + 2).value)
        ef = ""
        if c_ef:
            ef = _estaca_str(ws.cell(r, c_ef).value, ws.cell(r, c_ef + 1).value,
                             ws.cell(r, c_ef + 2).value)
        linhas.append({
            "linha": r,
            "estaca_inicial": ei,
            "estaca_final": ef,
            "lado": ws.cell(r, col_map["lado"]).value if "lado" in col_map else None,
            "extensao_km": ws.cell(r, col_map["extensao (km)"]).value
                           if "extensao (km)" in col_map else None,
            "servico_codigo": str(serv).strip(),
            "situacao": ws.cell(r, col_map["situacao"]).value
                        if "situacao" in col_map else None,
            "descricao": ws.cell(r, col_map["descricao"]).value
                         if "descricao" in col_map else None,
            "observacao": ws.cell(r, col_map["observacoes"]).value
                          if "observacoes" in col_map else None,
        })
        r += 1
        if r - header_row > 500:  # salvaguarda
            break
    return linhas


def _mapas_taxonomia(road_key):
    """codigo -> nome canonico (servico) e letra -> nome canonico (situacao),
    lidos de config/services_taxonomy.yaml (nao hardcoded por obra)."""
    tax = C.services_taxonomy()
    codigo_para_nome = dict(tax.get("codigos_ficha_de_campo_por_obra", {}).get(road_key, {}))
    if not codigo_para_nome:
        codigo_para_nome = dict(tax.get("codigos_ficha_de_campo", {}))
    situacao = {k: v.get("canonical_name", k) for k, v in tax.get("situacao", {}).items()}
    return codigo_para_nome, situacao


def gerar(caminho_ficha_acumulada, road_key, out_dir) -> FiguraResult:
    prov_base = {"source_file": caminho_ficha_acumulada, "sheet": "FICHA CAMPO"}

    if not caminho_ficha_acumulada:
        return FiguraResult(None, LARGURA_MM, 0.0, "LIMITATION",
                            "caminho da Ficha de Campo Acumulada nao resolvido (NOT_FOUND)",
                            prov_base)
    assert_nao_e_golden(caminho_ficha_acumulada)
    if not U.exists(caminho_ficha_acumulada):
        return FiguraResult(None, LARGURA_MM, 0.0, "LIMITATION",
                            "arquivo nao encontrado: %s" % caminho_ficha_acumulada, prov_base)

    try:
        destino = "figura1_%s.xlsx" % U.slug(os.path.basename(caminho_ficha_acumulada))
        local = U.copiar_para_tmp(caminho_ficha_acumulada, "figuras_ficha_acumulada", destino)
        wb = openpyxl.load_workbook(local, data_only=True)
    except OSError as e:
        return FiguraResult(None, LARGURA_MM, 0.0, "LIMITATION",
                            "falha ao abrir arquivo: %s" % e, prov_base)

    aba = "FICHA CAMPO" if "FICHA CAMPO" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[aba]

    header_row, col_map = _achar_cabecalho(ws)
    if header_row is None:
        return FiguraResult(None, LARGURA_MM, 0.0, "LIMITATION",
                            "cabecalho da tabela de estacas nao encontrado (varredura de "
                            "40x40 celulas por texto normalizado)", prov_base)

    linhas = _ler_tabela(ws, header_row, col_map)
    if not linhas:
        return FiguraResult(None, LARGURA_MM, 0.0, "LIMITATION",
                            "tabela de estacas sem linhas de dados apos o cabecalho (linha %d)"
                            % header_row, prov_base)

    codigo_para_nome, situacao_para_nome = _mapas_taxonomia(road_key)

    def fmt_servico(codigo):
        return codigo_para_nome.get(codigo, codigo)

    def fmt_situacao(letra):
        if letra is None:
            return ""
        return situacao_para_nome.get(str(letra).strip(), str(letra))

    # ---- desenho ----
    img_provisorio, draw_provisorio = novo_canvas(LARGURA_MM, 10)
    fonte = get_font(FONTE_PT)
    fonte_cab = get_font(FONTE_CABECALHO_PT, bold=True)
    pad_px = mm_to_px(PAD_MM)

    larguras_px = [mm_to_px(w) for _, _, w in _COLUNAS]
    # ajusta arredondamento para bater exatamente a largura total
    larguras_px[-1] += mm_to_px(LARGURA_MM) - sum(larguras_px)

    linhas_render = []
    for row in linhas:
        valores = [
            row["estaca_inicial"], row["estaca_final"],
            row["lado"] or "", ("%.2f" % row["extensao_km"]) if isinstance(
                row["extensao_km"], (int, float)) else (row["extensao_km"] or ""),
            fmt_servico(row["servico_codigo"]), fmt_situacao(row["situacao"]),
            row["descricao"] or "", row["observacao"] or "",
        ]
        wrapped = []
        n_linhas_max = 1
        for v, w_px in zip(valores, larguras_px):
            ls = wrap_text(draw_provisorio, str(v), fonte, w_px - 2 * pad_px)
            wrapped.append(ls)
            n_linhas_max = max(n_linhas_max, len(ls))
        altura_px = max(mm_to_px(LINHA_ALTURA_MIN_MM),
                        n_linhas_max * (pt_to_px(FONTE_PT) + 2) + 2 * pad_px)
        linhas_render.append((wrapped, altura_px))

    cabecalho_altura_px = mm_to_px(CABECALHO_ALTURA_MM)
    altura_total_px = cabecalho_altura_px + sum(a for _, a in linhas_render)
    largura_total_px = sum(larguras_px)

    img = Image.new("RGB", (largura_total_px, altura_total_px), COR_BRANCO)
    draw = ImageDraw.Draw(img)

    # cabecalho sombreado (estilo secao 5 do GOLDEN_VISUAL_SPEC.md)
    draw.rectangle([0, 0, largura_total_px, cabecalho_altura_px], fill=COR_SHADING_CABECALHO())
    x = 0
    fonte_cab_altura = pt_to_px(FONTE_CABECALHO_PT) + 2
    for (titulo, _campo, _w), w_px in zip(_COLUNAS, larguras_px):
        linhas_titulo = titulo.split("\n")
        bloco_h = len(linhas_titulo) * fonte_cab_altura
        ty = (cabecalho_altura_px - bloco_h) / 2
        for lt in linhas_titulo:
            tw = text_width(draw, lt, fonte_cab)
            draw.text((x + w_px / 2 - tw / 2, ty), lt, font=fonte_cab, fill=COR_PRETO)
            ty += fonte_cab_altura
        x += w_px

    y = cabecalho_altura_px
    for wrapped, altura_px in linhas_render:
        x = 0
        for ls, w_px in zip(wrapped, larguras_px):
            ty = y + pad_px
            for linha_txt in ls:
                draw.text((x + pad_px, ty), linha_txt, font=fonte, fill=COR_PRETO)
                ty += pt_to_px(FONTE_PT) + 2
            x += w_px
        y += altura_px

    # bordas: contorno + linhas de coluna + linhas de linha (0.5pt no golden;
    # a 300dpi 0.5pt ~= 2px, usamos 1-2px para nitidez sem engrossar demais)
    borda_px = 2
    desenhar_borda(draw, (0, 0, largura_total_px - 1, altura_total_px - 1), borda_px, COR_BORDA)
    x = 0
    for w_px in larguras_px[:-1]:
        x += w_px
        draw.line([(x, 0), (x, altura_total_px)], fill=COR_BORDA, width=1)
    y = cabecalho_altura_px
    draw.line([(0, y), (largura_total_px, y)], fill=COR_BORDA, width=1)
    for _wrapped, altura_px in linhas_render:
        y += altura_px
        draw.line([(0, y), (largura_total_px, y)], fill=COR_BORDA, width=1)

    nome_arquivo = "figura1_ficha_acumulada.png"
    caminho_saida = salvar_png(img, out_dir, nome_arquivo)
    altura_mm = altura_total_px / 300.0 * 25.4

    provenance = dict(prov_base)
    provenance.update({
        "range": "linha %d (cabecalho) a %d" % (header_row, linhas[-1]["linha"]),
        "n_linhas": len(linhas),
        "method": ("PIL render nativo da tabela de estacas cumulativa; cabecalho "
                   "localizado por texto normalizado (nao por linha fixa); codigos de "
                   "servico/situacao traduzidos via config/services_taxonomy.yaml "
                   "(road_key=%s)" % road_key),
        "correcao_2_compactacao": (
            "linhas compactadas (fonte %.1fpt, preenchimento %.1fmm, cabecalho "
            "%.1fmm, minimo por linha %.1fmm) para caber a figura inteira numa "
            "unica pagina A4 (altura util %.0fmm); largura util (160mm) e "
            "proporcao das colunas mantidas -- nenhuma linha de dado foi "
            "cortada" % (FONTE_PT, PAD_MM, CABECALHO_ALTURA_MM, LINHA_ALTURA_MIN_MM,
                        ALTURA_UTIL_A4_MM)),
    })
    if altura_mm > ALTURA_UTIL_A4_MM:
        return FiguraResult(
            caminho_saida, LARGURA_MM, round(altura_mm, 2), "LIMITATION",
            "figura excede a altura util da pagina A4 (%.2fmm > %.0fmm) mesmo apos a "
            "compactacao da Correcao 2 -- %d linhas de dados; nao foi cortada nenhuma "
            "linha (todas as %d permanecem na imagem), mas ela pode invadir a pagina "
            "seguinte no DOCX" % (altura_mm, ALTURA_UTIL_A4_MM, len(linhas), len(linhas)),
            provenance)
    return FiguraResult(caminho_saida, LARGURA_MM, round(altura_mm, 2), "OK", None, provenance)
