# -*- coding: utf-8 -*-
r"""
scripts/gerador/figuras/__init__.py - contrato comum e helpers compartilhados
dos modulos de figura da Fase 3.3 (FRENTE A).

Contrato (docs/arquitetura/FASE3_3_FINAL_CANDIDATE_SPEC.md Sec.1): todo modulo
de figura devolve um `FiguraResult`. PNG a 300 dpi sobre a dimensao em mm
(`px = mm / 25.4 * 300`).

Sem matplotlib, sem qrcode (Sec.0.1 - nao instalados neste ambiente): todo
desenho e feito com PIL (`ImageDraw`). Fonte: Arial (`C:\Windows\Fonts`), a
mesma exigida pelo `GOLDEN_VISUAL_SPEC.md` e por `config/visual_style_v3.yaml`
(so leitura -- este pacote nao grava em config/*.yaml).
"""
import os
import sys
import dataclasses
from typing import Optional, Dict, Any, List, Tuple

from PIL import Image, ImageDraw, ImageFont

AQUI = os.path.dirname(os.path.abspath(__file__))
GERADOR = os.path.dirname(AQUI)
SCRIPTS = os.path.dirname(GERADOR)
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from .. import assert_nao_e_golden  # noqa: E402
from .. import config as C  # noqa: E402

DPI = 300
FONT_DIR = r"C:\Windows\Fonts"


# -----------------------------------------------------------------------------
# contrato (spec Sec.1)
# -----------------------------------------------------------------------------
@dataclasses.dataclass
class FiguraResult:
    path: Optional[str]          # PNG gerado em output/shadow/<...>/figuras/, ou None
    largura_mm: float
    altura_mm: float
    status: str                  # OK | LIMITATION | PRESENTATION_DEPENDENCY | HITL
    note: Optional[str]
    provenance: Dict[str, Any]

    def to_dict(self) -> Dict[str, Any]:
        return dataclasses.asdict(self)


# -----------------------------------------------------------------------------
# conversao mm <-> px a 300 dpi
# -----------------------------------------------------------------------------
def mm_to_px(mm: float) -> int:
    return max(1, round(mm / 25.4 * DPI))


def px_to_mm(px: float) -> float:
    return px / DPI * 25.4


def pt_to_px(pt: float) -> int:
    """Tamanho de fonte: pt (1/72 polegada) -> px na resolucao de saida (300 dpi)."""
    return max(1, round(pt / 72.0 * DPI))


# -----------------------------------------------------------------------------
# cores (config/visual_style_v3.yaml - so LEITURA; nao e um dos 3 arquivos
# vedados pela spec e nao modificamos o yaml, so consumimos os valores)
# -----------------------------------------------------------------------------
def hex_to_rgb(h: str) -> Tuple[int, int, int]:
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


_VS = None


def visual_style():
    global _VS
    if _VS is None:
        _VS = C.visual_style()
    return _VS


def cor(nome: str, default: str = "000000") -> Tuple[int, int, int]:
    cores = visual_style().get("colors", {})
    return hex_to_rgb(cores.get(nome, default))


COR_TEXTO = lambda: cor("texto_corpo", "000000")           # noqa: E731
COR_LEGENDA = lambda: cor("legenda", "0D0D0D")              # noqa: E731
COR_SHADING_CABECALHO = lambda: cor("tabela_clima_shading", "F2F2F2")  # noqa: E731
COR_BORDA = (120, 120, 120)
COR_BRANCO = (255, 255, 255)
COR_PRETO = (0, 0, 0)

# paleta para graficos (nao definida no golden para Grafico 1 - escolha do
# shadow, documentada nos modulos que a usam)
PALETA_GRAFICOS = {
    "azul_escuro": (15, 71, 97),      # 0F4761 - mesma familia do H1 do golden
    "azul_claro": (79, 129, 189),     # 4F81BD - mesma familia do tema do golden
    "cinza_grade": (210, 210, 210),
    "verde": (60, 140, 90),
    "vermelho": (176, 60, 60),
    "amarelo": (218, 165, 32),
    "roxo": (120, 90, 160),
    "laranja": (214, 130, 50),
}


# -----------------------------------------------------------------------------
# fontes
# -----------------------------------------------------------------------------
_FONT_CACHE: Dict[Tuple[str, int], ImageFont.FreeTypeFont] = {}


def _font_path(bold=False, italic=False) -> str:
    if bold and italic:
        name = "arialbi.ttf"
    elif bold:
        name = "arialbd.ttf"
    elif italic:
        name = "ariali.ttf"
    else:
        name = "arial.ttf"
    p = os.path.join(FONT_DIR, name)
    return p if os.path.exists(p) else os.path.join(FONT_DIR, "arial.ttf")


def get_font(pt: float, bold: bool = False, italic: bool = False) -> ImageFont.FreeTypeFont:
    key = (_font_path(bold, italic), pt_to_px(pt))
    f = _FONT_CACHE.get(key)
    if f is None:
        f = ImageFont.truetype(key[0], key[1])
        _FONT_CACHE[key] = f
    return f


def text_width(draw: ImageDraw.ImageDraw, text: str, font) -> float:
    return draw.textlength(text, font=font)


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font, max_width_px: float) -> List[str]:
    """Quebra `text` em linhas que cabem em `max_width_px`, palavra a palavra."""
    if not text:
        return [""]
    palavras = str(text).split()
    linhas, atual = [], ""
    for w in palavras:
        candidata = (atual + " " + w).strip()
        if text_width(draw, candidata, font) <= max_width_px or not atual:
            atual = candidata
        else:
            linhas.append(atual)
            atual = w
    if atual:
        linhas.append(atual)
    return linhas or [""]


# -----------------------------------------------------------------------------
# canvas / salvamento
# -----------------------------------------------------------------------------
def novo_canvas(largura_mm: float, altura_mm: float, bg=COR_BRANCO):
    w, h = mm_to_px(largura_mm), mm_to_px(altura_mm)
    img = Image.new("RGB", (w, h), bg)
    return img, ImageDraw.Draw(img)


def salvar_png(img: Image.Image, out_dir: str, nome_arquivo: str) -> str:
    U.assert_read_only(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    caminho = os.path.join(out_dir, nome_arquivo)
    U.assert_read_only(caminho)
    img.save(caminho, format="PNG", dpi=(DPI, DPI))
    return caminho


def desenhar_borda(draw: ImageDraw.ImageDraw, box, largura_px=1, cor_borda=COR_BORDA):
    x0, y0, x1, y1 = box
    for i in range(largura_px):
        draw.rectangle([x0 + i, y0 + i, x1 - i, y1 - i], outline=cor_borda)


# re-export para os modulos de figura (evita repetir o import em cada um)
__all__ = [
    "FiguraResult", "DPI", "mm_to_px", "px_to_mm", "pt_to_px",
    "hex_to_rgb", "cor", "COR_TEXTO", "COR_LEGENDA", "COR_SHADING_CABECALHO",
    "COR_BORDA", "COR_BRANCO", "COR_PRETO", "PALETA_GRAFICOS",
    "get_font", "text_width", "wrap_text",
    "novo_canvas", "salvar_png", "desenhar_borda",
    "U", "assert_nao_e_golden", "C",
]
