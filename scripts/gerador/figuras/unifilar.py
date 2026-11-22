# -*- coding: utf-8 -*-
r"""
scripts/gerador/figuras/unifilar.py - UnifilarRenderer (Fase 3.4/v0.2):
renderiza a prancha do Diagrama Unifilar (Anexo IV) a partir da propria
planilha `UNIFILAR <obra>.xlsx`, via `openpyxl` (leitura de celulas/estilos)
+ `Pillow` (raster) -- SEM Word/Excel/LibreOffice/COM, e sem instalar nada.

## Por que isto existe

`docs/fontes/unifilar_renderizacao.md` (investigacao pontual, fechada em
2026-09-18) concluiu que nao ha PDF/imagem colada pronta para extrair: o
"desenho" do unifilar e' uma malha de CELULAS PINTADAS (fills + bordas +
mesclagens sobre uma grade de colunas de largura 1 char, uma por sub-trecho
de estaca), nao uma imagem. A unica via sem instalar software (Caminho 4 do
doc, avaliada mas nao implementada naquela investigacao) e' redesenhar essa
grade programaticamente lendo `openpyxl` -- exatamente o que este modulo faz.

## Genérico -- zero `if obra ==`

Tudo que varia entre obras (nome da aba, area de impressao, cores, larguras
de coluna, alturas de linha, mesclagens, textos, bordas, alinhamentos, logos)
vem da PROPRIA planilha (via `openpyxl`) ou de PARAMETROS explicitos
(`largura_pagina_mm`/`altura_pagina_mm`, que o chamador informa a partir de
`road_registry.yaml: obras.<OBRA>.formato_unifilar` -- ex.: "A0 paisagem" nas
3 obras, "A3 paisagem" na RD-102A). Nenhuma logica aqui distingue obras por
nome.

## Fidelidade -- deliberadamente NAO pixel-perfect (autorizado pelo briefing)

- Conversao largura de coluna (unidade "caracteres" do Excel) -> pixels usa a
  formula aproximada padrao (largura*7+5 para >=1 char, *12 para <1 -- a
  mesma usada por bibliotecas como `xlsxwriter`), nao a metrica exata da
  fonte Calibri 11 que o Excel usa internamente.
- Cores de tema (`theme`+`tint`) sao resolvidas lendo `xl/theme/theme1.xml`
  do proprio arquivo e aplicando um tint LINEAR sobre RGB (aproximacao usual;
  o Excel usa um algoritmo baseado em HSL). Cores diretas (`rgb`) sao exatas.
- Fontes: Arial (mesma fonte-base do resto do gerador), nao a fonte original
  da celula quando ela nao for Arial (ex.: alguns rotulos usam "Aptos Slab
  SemiBold" -- 9 celulas na amostra RD-103).
- O "ajuste a 1 pagina" do Excel na impressao (zoom exato) NAO e' replicado;
  em vez disso a grade inteira e' escalada uniformemente (preservando a
  proporcao) para caber na folha final com margem, e centralizada -- efeito
  visual equivalente, calculo mais simples.

Nenhuma dessas aproximacoes inventa DADO (cor/posicao/texto sempre vem de uma
celula real da planilha); elas so' afastam o resultado de ser pixel-identico
ao que o motor de impressao do Excel produziria. Registradas em
`provenance.limitacoes_fidelidade` do `FiguraResult` devolvido.
"""
import os
import re
import sys
import zipfile
import datetime as dt
from io import BytesIO

from PIL import Image, ImageDraw, ImageFont
import openpyxl
from openpyxl.utils import range_boundaries, get_column_letter

AQUI = os.path.dirname(os.path.abspath(__file__))
GERADOR = os.path.dirname(AQUI)
SCRIPTS = os.path.dirname(GERADOR)
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402

from . import FiguraResult, assert_nao_e_golden, _font_path  # noqa: E402

# -----------------------------------------------------------------------------
# constantes de folha (paisagem; o chamador escolhe A0 ou A3 via parametro)
# -----------------------------------------------------------------------------
A0_LARGURA_MM, A0_ALTURA_MM = 1189.0, 841.0
A3_LARGURA_MM, A3_ALTURA_MM = 420.0, 297.0

# DPI de saida deliberadamente MENOR que o padrao de 300dpi do resto do
# pacote `figuras/` (ver __init__.py) -- uma folha A0 a 300dpi teria ~14000 x
# 9900 px (~139 megapixels), pesada demais para gerar/gravar neste ambiente
# sem ganho pratico (a prancha vai para um anexo digital, nao para plotter).
# 150dpi ainda e' resolucao de impressao de qualidade para um desenho de
# grande formato (~7000 x 4950 px) e mantem o tempo de geracao em segundos.
DPI_UNIFILAR = 150
MARGEM_MM = 15.0

_EMU_POR_POLEGADA = 914400.0


def _mm_to_px(mm, dpi):
    return max(1, round(mm / 25.4 * dpi))


def _emu_to_px96(emu):
    return (emu or 0) / _EMU_POR_POLEGADA * 96.0


# -----------------------------------------------------------------------------
# tema de cores (xl/theme/theme1.xml) -- resolve Color(type='theme', tint=..)
# -----------------------------------------------------------------------------
# Ordem real no XML (<a:clrScheme>): dk1, lt1, dk2, lt2, accent1..6, hlink,
# folHlink. O INDICE de tema usado em openpyxl.Color.theme troca 0<->1 e
# 2<->3 em relacao a essa ordem (comportamento documentado do Excel/OOXML:
# "Background 1"/"Text 1" trocam de posicao) -- por isso o remapeamento
# abaixo, nao uma leitura direta por posicao.
_ORDEM_XML = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
             "accent4", "accent5", "accent6", "hlink", "folHlink"]
_REMAPA_INDICE_TEMA = {0: 1, 1: 0, 2: 3, 3: 2}  # demais indices (4..9) sao diretos


def _ler_tema_cores(caminho_xlsx):
    """Devolve lista de 10 cores hex (RRGGBB), na ordem do indice de tema do
    openpyxl (0..9 = Background1/Text1/Background2/Text2/Accent1..6).
    Fallback para uma paleta neutra se o tema nao existir/nao puder ser lido
    -- nunca lanca excecao (uma prancha sem tema legivel ainda deve ser
    gerada, so' com cores de tema aproximadas por cinza)."""
    try:
        with zipfile.ZipFile(caminho_xlsx) as z:
            nomes_tema = [n for n in z.namelist() if re.match(r"xl/theme/theme\d*\.xml$", n)]
            if not nomes_tema:
                raise FileNotFoundError("sem xl/theme/*.xml")
            xml = z.read(nomes_tema[0]).decode("utf-8", "ignore")
        bloco = re.search(r"<a:clrScheme.*?</a:clrScheme>", xml, re.S)
        if not bloco:
            raise ValueError("sem <a:clrScheme> no tema")
        bloco = bloco.group(0)
        cores_por_tag = {}
        for tag in _ORDEM_XML:
            m = re.search(r"<a:%s>.*?(?:val=\"([0-9A-Fa-f]{6})\"|lastClr=\"([0-9A-Fa-f]{6})\")"
                         r".*?</a:%s>" % (tag, tag), bloco, re.S)
            cores_por_tag[tag] = (m.group(1) or m.group(2)).upper() if m else "808080"
        ordem_indice = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                       "accent4", "accent5", "accent6"]
        cores = [cores_por_tag[t] for t in ordem_indice]
        # aplica a troca 0<->1, 2<->3
        finais = list(cores)
        for i, j in _REMAPA_INDICE_TEMA.items():
            finais[i] = cores[j]
        return finais
    except Exception:
        return ["808080"] * 10


def _aplicar_tint(rgb, tint):
    """Aproximacao LINEAR do tint do Excel (o real usa HSL) -- documentada
    como limitacao de fidelidade na docstring do modulo."""
    if not tint:
        return rgb
    if tint < 0:
        f = 1.0 + tint
        return tuple(max(0, min(255, int(round(ch * f)))) for ch in rgb)
    return tuple(max(0, min(255, int(round(ch * (1 - tint) + 255 * tint)))) for ch in rgb)


def _hex_para_rgb(hexstr):
    hexstr = (hexstr or "000000")[-6:]
    return tuple(int(hexstr[i:i + 2], 16) for i in (0, 2, 4))


def _resolver_cor(color, tema, cache):
    """`color`: um `openpyxl.styles.colors.Color` (fgColor de fill, ou color
    de border side). Devolve tupla RGB ou None (sem cor / tipo nao suportado
    -- NUNCA inventa uma cor)."""
    if color is None:
        return None
    tipo = getattr(color, "type", None)
    if tipo == "rgb":
        val = color.rgb
        if not isinstance(val, str) or len(val) < 6:
            return None
        chave = ("rgb", val)
    elif tipo == "theme":
        idx = color.theme if color.theme is not None else 0
        chave = ("theme", idx, round(color.tint or 0.0, 4))
    elif tipo == "indexed":
        chave = ("indexed", color.indexed)
    else:
        return None
    if chave in cache:
        return cache[chave]
    if tipo == "rgb":
        rgb = _hex_para_rgb(chave[1])
    elif tipo == "theme":
        base_hex = tema[idx] if 0 <= idx < len(tema) else "808080"
        rgb = _aplicar_tint(_hex_para_rgb(base_hex), chave[2])
    else:
        # paleta indexada legada -- so os indices mais comuns; qualquer outro
        # cai em cinza medio (aproximacao documentada, nao inventa cor viva)
        rgb = {0: (0, 0, 0), 1: (255, 255, 255), 2: (255, 0, 0), 3: (0, 255, 0),
              4: (0, 0, 255), 5: (255, 255, 0), 64: (0, 0, 0)}.get(chave[1], (128, 128, 128))
    cache[chave] = rgb
    return rgb


# -----------------------------------------------------------------------------
# geometria: largura de coluna / altura de linha -> pixels (dominio 96dpi)
# -----------------------------------------------------------------------------
def _largura_coluna_px96(largura_chars):
    if largura_chars is None:
        largura_chars = 8.43
    if largura_chars < 1:
        return max(1.0, largura_chars * 12.0 + 0.5)
    return max(1.0, largura_chars * 7.0 + 5.0)


def _altura_linha_px96(altura_pt):
    if altura_pt is None:
        altura_pt = 15.0
    return max(1.0, altura_pt * 96.0 / 72.0)


def _larguras_colunas_px96(ws, min_col, max_col):
    default = ws.sheet_format.defaultColWidth
    dims = [d for d in ws.column_dimensions.values() if d.width]
    out = []
    for c in range(min_col, max_col + 1):
        largura = default
        for d in dims:
            if d.min is not None and d.max is not None and d.min <= c <= d.max:
                largura = d.width
                break
        out.append(_largura_coluna_px96(largura))
    return out


def _alturas_linhas_px96(ws, min_row, max_row):
    default = ws.sheet_format.defaultRowHeight
    out = []
    for r in range(min_row, max_row + 1):
        dim = ws.row_dimensions.get(r)
        altura = dim.height if (dim and dim.height) else default
        out.append(_altura_linha_px96(altura))
    return out


def _offsets(tamanhos):
    off = [0.0]
    acc = 0.0
    for t in tamanhos:
        acc += t
        off.append(acc)
    return off


# -----------------------------------------------------------------------------
# mesclagens que cruzam a area de impressao
# -----------------------------------------------------------------------------
def _mapear_merges(ws, min_row, max_row, min_col, max_col):
    """Devolve (por_posicao_de_desenho, cobertas): `por_posicao_de_desenho`
    mapeia a 1a posicao (linha,coluna) DENTRO da area onde a mesclagem deve
    ser desenhada -> (linha_ancora_real, col_ancora_real, linha_fim_clip,
    col_fim_clip) -- a ancora real pode estar FORA da area (mesclagem que
    comeca antes do print_area) mas o valor/estilo sempre vem dela, nunca
    inventado. `cobertas` e' o conjunto de celulas que NAO devem ser
    desenhadas de novo (ja fazem parte de uma mesclagem desenhada alhures)."""
    por_pos = {}
    cobertas = set()
    for rng in ws.merged_cells.ranges:
        r0, c0, r1, c1 = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        if r1 < min_row or r0 > max_row or c1 < min_col or c0 > max_col:
            continue
        r0c, c0c = max(r0, min_row), max(c0, min_col)
        r1c, c1c = min(r1, max_row), min(c1, max_col)
        por_pos[(r0c, c0c)] = (r0, c0, r1c, c1c)
        for rr in range(r0c, r1c + 1):
            for cc in range(c0c, c1c + 1):
                if (rr, cc) != (r0c, c0c):
                    cobertas.add((rr, cc))
    return por_pos, cobertas


# -----------------------------------------------------------------------------
# texto
# -----------------------------------------------------------------------------
_FONT_CACHE = {}


def _fonte(tamanho_px, bold=False):
    tamanho_px = max(6, int(round(tamanho_px)))
    chave = (bold, tamanho_px)
    f = _FONT_CACHE.get(chave)
    if f is None:
        f = ImageFont.truetype(_font_path(bold=bold), tamanho_px)
        _FONT_CACHE[chave] = f
    return f


def _formatar_valor(valor):
    if valor is None or valor == "":
        return ""
    if isinstance(valor, dt.datetime):
        return valor.strftime("%d/%m/%Y") if valor.hour == valor.minute == 0 else \
            valor.strftime("%d/%m/%Y %H:%M")
    if isinstance(valor, dt.date):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, float):
        return str(int(valor)) if valor == int(valor) else ("%.2f" % valor)
    return str(valor).strip()


def _desenhar_texto_celula(img, draw, texto, rect, font, cor_rgb, alinh_h, alinh_v, rotacao):
    if not texto:
        return
    x0, y0, x1, y1 = rect
    if rotacao in (90, 180, 270):
        tw = draw.textlength(texto, font=font)
        th = font.size
        tmp = Image.new("RGBA", (int(tw) + 4, int(th) + 4), (255, 255, 255, 0))
        tdraw = ImageDraw.Draw(tmp)
        tdraw.text((2, 2), texto, font=font, fill=cor_rgb + (255,))
        tmp = tmp.rotate(rotacao, expand=True)
        px = int((x0 + x1) / 2 - tmp.width / 2)
        py = int((y0 + y1) / 2 - tmp.height / 2)
        img.paste(tmp, (px, py), tmp)
        return
    tw = draw.textlength(texto, font=font)
    th = font.size
    if alinh_h == "right":
        tx = x1 - tw - 2
    elif alinh_h == "left":
        tx = x0 + 2
    else:
        tx = (x0 + x1) / 2 - tw / 2
    if alinh_v == "top":
        ty = y0 + 1
    elif alinh_v == "bottom":
        ty = y1 - th - 1
    else:
        ty = (y0 + y1) / 2 - th / 2
    draw.text((tx, ty), texto, font=font, fill=cor_rgb)


# -----------------------------------------------------------------------------
# UnifilarRenderer
# -----------------------------------------------------------------------------
class UnifilarRenderer:
    """Le `caminho_xlsx` (copia local da planilha UNIFILAR da semana --
    `auxiliares.unifilar_semana` em `config/source_layout.yaml`) e renderiza
    a area de impressao da 1a aba cujo nome comeca com 'unifilar'
    (normalizado) como uma prancha raster, escalada para caber numa folha
    `largura_pagina_mm` x `altura_pagina_mm` (paisagem)."""

    def __init__(self, caminho_xlsx, *, largura_pagina_mm=A0_LARGURA_MM,
                altura_pagina_mm=A0_ALTURA_MM, dpi=DPI_UNIFILAR):
        self.caminho_original = caminho_xlsx
        self.largura_pagina_mm = largura_pagina_mm
        self.altura_pagina_mm = altura_pagina_mm
        self.dpi = dpi

    # -- passo 1: copia local + abertura ------------------------------------
    def _abrir(self):
        assert_nao_e_golden(self.caminho_original)
        destino = "unifilar_%s.xlsx" % U.slug(os.path.basename(self.caminho_original))
        self.caminho_local = U.copiar_para_tmp(self.caminho_original, "figuras_unifilar", destino)
        self.wb = openpyxl.load_workbook(self.caminho_local, data_only=True)
        return self.wb

    def _achar_aba(self):
        for nome in self.wb.sheetnames:
            if U.norm(nome).startswith("unifilar"):
                return self.wb[nome]
        return self.wb.worksheets[0]

    def _achar_area_impressao(self, ws):
        pa = ws.print_area
        if isinstance(pa, list):
            pa = pa[0] if pa else None
        if pa:
            ref = pa.split("!")[-1]
            min_col, min_row, max_col, max_row = range_boundaries(ref)
            return min_col, min_row, max_col, max_row, True
        # fallback: sem area de impressao definida -- usa a area usada da
        # aba inteira (documentado como limitacao, nao falha silenciosa)
        return 1, 1, ws.max_column, ws.max_row, False

    # -- passo 2: renderizacao ------------------------------------------------
    def renderizar(self, out_dir, nome_arquivo="figura_unifilar.png"):
        limitacoes = []
        self._abrir()
        ws = self._achar_aba()
        min_col, min_row, max_col, max_row, tinha_print_area = self._achar_area_impressao(ws)
        if not tinha_print_area:
            limitacoes.append("aba '%s' sem Print_Area definida -- usada a area total "
                              "utilizada (%d colunas x %d linhas) em vez de uma faixa "
                              "explicita" % (ws.title, ws.max_column, ws.max_row))

        larguras = _larguras_colunas_px96(ws, min_col, max_col)
        alturas = _alturas_linhas_px96(ws, min_row, max_row)
        offs_x = _offsets(larguras)
        offs_y = _offsets(alturas)
        total_w_px96, total_h_px96 = offs_x[-1], offs_y[-1]

        tema = _ler_tema_cores(self.caminho_local)
        cache_cor = {}

        # -- escala: cabe a grade inteira na folha, preservando proporcao,
        # com margem, centralizada (nao replica o "ajustar a 1 pagina" exato
        # do Excel -- ver limitacoes na docstring do modulo).
        area_util_mm_w = self.largura_pagina_mm - 2 * MARGEM_MM
        area_util_mm_h = self.altura_pagina_mm - 2 * MARGEM_MM
        total_w_mm_nativo = total_w_px96 * 25.4 / 96.0
        total_h_mm_nativo = total_h_px96 * 25.4 / 96.0
        fator_encaixe = min(area_util_mm_w / total_w_mm_nativo,
                            area_util_mm_h / total_h_mm_nativo)
        px_por_mm = self.dpi / 25.4
        k = (25.4 / 96.0) * fator_encaixe * px_por_mm  # native-px96 -> pixel final

        canvas_w = _mm_to_px(self.largura_pagina_mm, self.dpi)
        canvas_h = _mm_to_px(self.altura_pagina_mm, self.dpi)
        grade_w_px = total_w_px96 * k
        grade_h_px = total_h_px96 * k
        ox = (canvas_w - grade_w_px) / 2.0
        oy = (canvas_h - grade_h_px) / 2.0

        img = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
        draw = ImageDraw.Draw(img)

        def _x(idx_col_rel):
            return ox + offs_x[idx_col_rel] * k

        def _y(idx_row_rel):
            return oy + offs_y[idx_row_rel] * k

        por_pos_merge, cobertas = _mapear_merges(ws, min_row, max_row, min_col, max_col)

        n_preenchimentos = n_bordas = n_textos = 0
        for r in range(min_row, max_row + 1):
            ir = r - min_row
            for c in range(min_col, max_col + 1):
                ic = c - min_col
                if (r, c) in cobertas:
                    continue
                if (r, c) in por_pos_merge:
                    r_anc, c_anc, r_fim, c_fim = por_pos_merge[(r, c)]
                    cell = ws.cell(r_anc, c_anc)
                    ic_fim = c_fim - min_col
                    ir_fim = r_fim - min_row
                else:
                    cell = ws.cell(r, c)
                    ic_fim, ir_fim = ic, ir

                x0, x1 = _x(ic), _x(ic_fim + 1)
                y0, y1 = _y(ir), _y(ir_fim + 1)
                if x1 - x0 < 0.4 and y1 - y0 < 0.4:
                    continue  # celula sub-pixel nesta escala -- nada a desenhar

                fill = cell.fill
                if fill is not None and fill.patternType == "solid":
                    cor_fill = _resolver_cor(fill.fgColor, tema, cache_cor)
                    if cor_fill:
                        draw.rectangle([x0, y0, x1, y1], fill=cor_fill)
                        n_preenchimentos += 1

                borda = cell.border
                if borda is not None:
                    for lado, p0, p1 in (("top", (x0, y0), (x1, y0)),
                                         ("bottom", (x0, y1), (x1, y1)),
                                         ("left", (x0, y0), (x0, y1)),
                                         ("right", (x1, y0), (x1, y1))):
                        side = getattr(borda, lado, None)
                        if side is not None and side.style:
                            cor_borda = _resolver_cor(side.color, tema, cache_cor) or (0, 0, 0)
                            largura_linha = 2 if side.style in ("medium", "thick") else 1
                            draw.line([p0, p1], fill=cor_borda, width=largura_linha)
                            n_bordas += 1

                valor = _formatar_valor(cell.value)
                if valor:
                    tam_pt = (cell.font.size or 11) if cell.font else 11
                    negrito = bool(cell.font.bold) if cell.font else False
                    tam_px = max(6.0, tam_pt * 96.0 / 72.0 * k)
                    fonte = _fonte(tam_px, bold=negrito)
                    alinh = cell.alignment
                    cor_txt = (0, 0, 0)
                    rot = int(alinh.textRotation or 0) if alinh else 0
                    _desenhar_texto_celula(
                        img, draw, valor, (x0, y0, x1, y1), fonte, cor_txt,
                        (alinh.horizontal if alinh else None),
                        (alinh.vertical if alinh else None), rot)
                    n_textos += 1

        # -- logos do cabecalho (ancoras reais da planilha, nao posicao fixa)
        n_logos = 0
        for im in getattr(ws, "_images", []):
            try:
                a = im.anchor
                c0, r0 = a._from.col + 1, a._from.row + 1
                c1, r1 = a.to.col + 1, a.to.row + 1
                if not (min_col <= c0 <= max_col and min_row <= r0 <= max_row):
                    continue
                x0 = _x(c0 - min_col) + _emu_to_px96(a._from.colOff) * k
                y0 = _y(r0 - min_row) + _emu_to_px96(a._from.rowOff) * k
                x1 = _x(min(c1, max_col + 1) - min_col) + _emu_to_px96(a.to.colOff) * k
                y1 = _y(min(r1, max_row + 1) - min_row) + _emu_to_px96(a.to.rowOff) * k
                w, h = max(1, int(x1 - x0)), max(1, int(y1 - y0))
                dados = im._data() if hasattr(im, "_data") else im.ref.read()
                logo = Image.open(BytesIO(dados)).convert("RGBA")
                logo = logo.resize((w, h))
                img.paste(logo, (int(x0), int(y0)), logo)
                n_logos += 1
            except Exception as e:
                limitacoes.append("logo do cabecalho nao pode ser posicionado (%s)" % e)

        U.assert_read_only(out_dir)
        os.makedirs(out_dir, exist_ok=True)
        caminho_png = os.path.join(out_dir, nome_arquivo)
        U.assert_read_only(caminho_png)
        img.save(caminho_png, format="PNG", dpi=(self.dpi, self.dpi))

        limitacoes.extend([
            "largura de coluna (unidade 'caracteres' do Excel) convertida para pixel por "
            "formula aproximada (nao a metrica exata da fonte Calibri 11 do Excel)",
            "cores de tema resolvidas com tint LINEAR sobre RGB (o Excel usa HSL) -- "
            "aproximacao visual, mesma cor-base, luminosidade pode variar um pouco",
            "fonte sempre Arial (celulas com outra fonte, ex. 'Aptos Slab SemiBold', "
            "sao desenhadas em Arial no mesmo tamanho)",
            "zoom 'ajustar a 1 pagina' do Excel NAO replicado -- escala uniforme "
            "(preserva proporcao) calculada por esta rotina, nao pelo motor do Excel",
        ])

        provenance = {
            "source_file": self.caminho_original,
            "sheet": ws.title,
            "print_area": ws.print_area if tinha_print_area else None,
            "area_lida": "R%dC%d:R%dC%d" % (min_row, min_col, max_row, max_col),
            "n_colunas": max_col - min_col + 1, "n_linhas": max_row - min_row + 1,
            "n_preenchimentos_desenhados": n_preenchimentos,
            "n_bordas_desenhadas": n_bordas, "n_textos_desenhados": n_textos,
            "n_logos_posicionados": n_logos, "n_mesclagens_na_area": len(por_pos_merge),
            "dpi_saida": self.dpi, "pagina_mm": [self.largura_pagina_mm, self.altura_pagina_mm],
            "fator_encaixe_grade_na_pagina": round(fator_encaixe, 4),
            "method": ("openpyxl (celulas/estilos/mesclagens/logos) + Pillow (raster) -- "
                      "reconstrucao da grade da PROPRIA planilha, sem Word/Excel/LibreOffice/"
                      "COM; ver docs/fontes/unifilar_renderizacao.md Caminho 4"),
            "limitacoes_fidelidade": limitacoes,
        }
        note = ("prancha REAL renderizada a partir da planilha (%d preenchimentos, %d "
               "textos, %d mesclagens, %d logo(s)); NAO pixel-perfect frente ao que o "
               "Excel imprimiria -- ver provenance.limitacoes_fidelidade"
               % (n_preenchimentos, n_textos, len(por_pos_merge), n_logos))
        return FiguraResult(caminho_png, self.largura_pagina_mm, self.altura_pagina_mm,
                            "OK", note, provenance)


def gerar(caminho_unifilar_xlsx, out_dir, *, largura_pagina_mm=A0_LARGURA_MM,
         altura_pagina_mm=A0_ALTURA_MM, nome_arquivo="figura_unifilar.png") -> FiguraResult:
    """Ponto de entrada no padrao dos demais modulos de `figuras/`
    (`ficha_acumulada.gerar`, `avanco_grafico.gerar`, ...). Nunca falha
    silenciosamente: se a fonte nao existir, devolve LIMITATION; qualquer
    outro erro de leitura/renderizacao tambem vira LIMITATION com o motivo,
    nunca uma excecao nao tratada nem uma prancha inventada."""
    prov = {"source_file": caminho_unifilar_xlsx}
    if not caminho_unifilar_xlsx or not U.exists(caminho_unifilar_xlsx):
        return FiguraResult(None, largura_pagina_mm, altura_pagina_mm, "LIMITATION",
                            "arquivo UNIFILAR nao encontrado: %s" % caminho_unifilar_xlsx, prov)
    renderer = UnifilarRenderer(caminho_unifilar_xlsx, largura_pagina_mm=largura_pagina_mm,
                                altura_pagina_mm=altura_pagina_mm)
    try:
        return renderer.renderizar(out_dir, nome_arquivo=nome_arquivo)
    except Exception as e:
        return FiguraResult(None, largura_pagina_mm, altura_pagina_mm, "LIMITATION",
                            "falha ao renderizar o UNIFILAR a partir da planilha: %r" % e, prov)
