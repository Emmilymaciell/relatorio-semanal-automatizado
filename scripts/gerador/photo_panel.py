# -*- coding: utf-8 -*-
r"""
scripts/gerador/photo_panel.py - Fase 3.2B, P0.4: Relatorio Fotografico como
IMAGEM COMPOSTA por servico.

    fotos do servico  ->  PhotoPanelComposer  ->  imagem composta  ->  Word

O golden usa 1 EMF por servico (5 EMF no documento inteiro; as ~46 fotos
ficam DENTRO dos 4 metarquivos maiores) -- nunca insere JPEG um a um
(VISUAL_DIFF_MASTER.md §C3). Esta e' a INTERFACE obrigatoria a partir desta
rodada: `content_builder.py` e `renderers/docx_renderer.py` NUNCA devem
iterar fotos e inserir imagem por imagem -- toda foto de um servico passa por
`PhotoPanelComposer.compose()`.

A implementacao desta rodada e' MINIMA (permitido pela spec): calcula o grid
correto (colunas/linhas/dimensoes em mm) a partir da contagem real de fotos,
mas ainda nao renderiza a imagem composta de verdade -- devolve PLACEHOLDER
(ou LIMITATION quando nao ha fotos resolvidas). O que a arquitetura proibe e'
so isso: contornar o componente. Emitir placeholder POR ELE e' aceitavel
agora (P0.4).

DEC-004 (docs/decisoes/DEC-004-formato-das-imagens-compostas.md): EMF nao e'
obrigatorio -- PNG de alta resolucao (ou SVG) serve, desde que fidelidade
visual, dimensao correta em mm, estabilidade no Word e boa qualidade no PDF.
`image_format` default aqui e' "PNG" por isso.

Especificacao do painel (GOLDEN_VISUAL_SPEC.md §7, predominante):
    papel:            A3 retrato, largura util 247 mm
    painel:           247,00 x 306,32 mm
    grid:             3 colunas fixas; linhas = ceil(n_fotos / 3)
    coluna de foto:   ~80 mm
    altura por foto:  livre, 50-57 mm conforme aspecto da origem
    legenda:          3 linhas ACIMA de cada foto (Estaca+Data / Coordenada / Descricao)
    linha incompleta: mantem celulas vazias COM BORDA -- nao expande as fotos
"""
import dataclasses
import math
from typing import Any, Dict, List, Optional, Sequence

PANEL_WIDTH_MM = 247.00
PANEL_HEIGHT_MM = 306.32
N_COLUNAS = 3
COLUNA_MM = 80.0
FOTO_ALTURA_MIN_MM = 50.0
FOTO_ALTURA_MAX_MM = 57.0
N_LINHAS_LEGENDA = 3  # Estaca+Data / Coordenada / Descricao -- acima de cada foto


@dataclasses.dataclass
class FotoRef:
    """Referencia minima a uma foto de origem. `overlay_read=False` no
    report_schema.yaml (T-17): nesta fase nao fazemos OCR do overlay impresso
    na foto, entao estaca/coordenada/descricao costumam vir vazios -- o grid
    e' calculado mesmo assim, so a legenda completa fica para outra rodada."""
    path: Optional[str] = None
    estaca: Optional[str] = None
    data: Optional[str] = None
    coordenada: Optional[str] = None
    descricao: Optional[str] = None


@dataclasses.dataclass
class PhotoPanelPlan:
    """Plano de grade do painel -- serializavel, e' o que um futuro
    compositor de imagem real (ou o renderer, entretanto) consome para saber
    quantas celulas desenhar."""
    servico: str
    n_fotos: int
    n_colunas: int
    n_linhas: int
    panel_width_mm: float
    panel_height_mm: float
    coluna_mm: float
    celulas_vazias_ultima_linha: int
    n_linhas_legenda: int = N_LINHAS_LEGENDA

    def to_dict(self) -> Dict[str, Any]:
        return dataclasses.asdict(self)


@dataclasses.dataclass
class PhotoPanelResult:
    """Resultado de `compose()`: ou uma imagem composta real (`image_path`
    preenchido, status READY) ou um placeholder pelo componente certo
    (status PLACEHOLDER/LIMITATION) -- nunca fotos individuais inseridas
    direto no documento."""
    plan: PhotoPanelPlan
    status: str  # "READY" | "PLACEHOLDER" | "LIMITATION"
    image_path: Optional[str] = None
    image_format: str = "PNG"  # DEC-004: EMF nao e' obrigatorio
    note: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        d = dataclasses.asdict(self)
        d["plan"] = self.plan.to_dict()
        return d


class PhotoPanelComposer:
    """fotos do servico -> PhotoPanelComposer -> imagem composta -> Word.

    `compose()` e' a UNICA porta de entrada para o Relatorio Fotografico:
    content_builder e docx_renderer nunca devem iterar fotos e inserir
    imagem por imagem (arquitetura proibida, VISUAL_DIFF_MASTER.md §C3).
    """

    def __init__(self, n_colunas: int = N_COLUNAS,
                panel_width_mm: float = PANEL_WIDTH_MM,
                panel_height_mm: float = PANEL_HEIGHT_MM,
                coluna_mm: float = COLUNA_MM):
        self.n_colunas = n_colunas
        self.panel_width_mm = panel_width_mm
        self.panel_height_mm = panel_height_mm
        self.coluna_mm = coluna_mm

    def plan(self, servico: str, fotos: Sequence[FotoRef]) -> PhotoPanelPlan:
        n = len(fotos)
        n_linhas = math.ceil(n / self.n_colunas) if n else 1
        vazias = (n_linhas * self.n_colunas - n) if n else self.n_colunas
        return PhotoPanelPlan(
            servico=servico, n_fotos=n, n_colunas=self.n_colunas, n_linhas=n_linhas,
            panel_width_mm=self.panel_width_mm, panel_height_mm=self.panel_height_mm,
            coluna_mm=self.coluna_mm, celulas_vazias_ultima_linha=vazias)

    def compose(self, servico: str, fotos: Sequence[FotoRef],
               output_path: Optional[str] = None) -> PhotoPanelResult:
        """Nesta rodada (P0.4) sempre devolve PLACEHOLDER/LIMITATION -- a
        composicao real da imagem (grade + molduras + legendas desenhadas)
        fica para a proxima rodada. O `plan` (grid dimensionado a partir da
        contagem real de fotos) ja esta correto agora."""
        p = self.plan(servico, fotos)
        if not fotos:
            return PhotoPanelResult(
                plan=p, status="LIMITATION", image_path=None,
                note=("sem fotos resolvidas para o servico '%s' nesta fatia -- grid "
                     "minimo reservado (%d coluna(s) x %d linha)"
                     % (servico, p.n_colunas, p.n_linhas)))
        return PhotoPanelResult(
            plan=p, status="PLACEHOLDER", image_path=output_path,
            note=("painel composto de %d foto(s) em grade %dx%d (%.2f x %.2f mm) -- "
                 "renderizacao da imagem composta pendente (DEC-004: PNG/SVG); "
                 "grid ja calculado corretamente pelo PhotoPanelComposer"
                 % (p.n_fotos, p.n_colunas, p.n_linhas, p.panel_width_mm, p.panel_height_mm)))
