# -*- coding: utf-8 -*-
r"""
scripts/gerador/document_model.py - Fase 3, Slice 3: modelo abstrato do
documento (spec Sec.4).

Estrutura DELIBERADAMENTE sem Word e sem caminho de arquivo -- so dataclasses
serializaveis em JSON. `content_builder.py` preenche este modelo a partir de
`report_data.json`; `renderers/docx_renderer.py` consome este modelo para
produzir o .docx. Trocar DOCX por outro formato de saida nao exige mexer em
extrator nem em content_builder -- so em um novo renderer.

kind de Block:
    "cover_field"        capa: {label, text=valor}
    "heading"             titulo de secao/subsecao: {text, level}
    "paragraph"           paragrafo de corpo: {text}
    "table"                tabela: {headers, rows}
    "figure_placeholder"  marcador de imagem/grafico pendente: {label}
    "figure"               imagem de figura/grafico JA RENDERIZADA (Fase 3.3,
                           contrato Sec.1 -- FiguraResult da Frente A):
                           {image_path, width_mm, height_mm}
    "caption"              legenda de figura/tabela: {text}
    "list"                  lista (ex.: anexos numerados): {items}
    "toc"                   indice: {items=[(numero, titulo, nivel)]}
    "page_break"           quebra de pagina explicita
    "full_page_image"     anexo EMBEDDED_IMAGE (Fase 3.3 Sec.4): 1 pagina do
                           PDF-fonte renderizada a imagem cheia:
                           {image_path, width_mm, height_mm}

status de Block/Section:
    "READY"        pronto, comprovado pela fonte
    "PLACEHOLDER"  posicao estrutural correta, conteudo real pendente (ex.: imagem)
    "HITL"         gerado, mas exige confirmacao humana antes de publicar
    "LIMITATION"   a fonte nao permite produzir este conteudo nesta fatia

`status != READY` SEMPRE entra em `marks` com o motivo (spec Sec.5): "todo
bloco com validation_status != READY recebe marca visual inequivoca no DOCX".
"""
import dataclasses
from typing import Any, Dict, List, Optional

STATUS_VALIDOS = ("READY", "PLACEHOLDER", "HITL", "LIMITATION")


@dataclasses.dataclass
class Block:
    kind: str
    status: str = "READY"
    provenance: List[str] = dataclasses.field(default_factory=list)
    text: Optional[str] = None
    level: Optional[int] = None
    label: Optional[str] = None
    items: Optional[List[str]] = None
    headers: Optional[List[str]] = None
    rows: Optional[List[List[str]]] = None
    note: Optional[str] = None
    marks: List[str] = dataclasses.field(default_factory=list)
    # P1 (renderer v2, design system): rotulo SEMANTICO opcional para o
    # renderer reconhecer um bloco com necessidade visual especial sem
    # depender de posicao/numero (ex.: "climate_table" para a tabela de
    # clima com as 8 larguras/bordas/sombreado exatos de P1.9, ou
    # "nc_status_table" para a tabela de status de NC -- Fase 3.3 Sec.3.3).
    # None em todo o resto -- nao muda nada do renderer v1 nem do
    # document_model.json ja gravado (campo aditivo, default None).
    role: Optional[str] = None
    # Fase 3.3 (Sec.1 contrato FiguraResult / Sec.4 anexos embutidos): campos
    # aditivos para os kinds "figure" e "full_page_image" -- imagem JA
    # renderizada em disco (nunca um caminho de planilha/Word). Default None
    # em todo o resto, nao muda nenhum document_model.json ja gravado.
    image_path: Optional[str] = None
    width_mm: Optional[float] = None
    height_mm: Optional[float] = None

    def __post_init__(self):
        if self.status not in STATUS_VALIDOS:
            raise ValueError("status invalido: %r (validos: %s)" % (self.status, STATUS_VALIDOS))
        if self.status != "READY" and "HITL" not in self.marks and self.status != "PLACEHOLDER":
            # PLACEHOLDER ja e visualmente inequivoco (rotulo "RENDERIZACAO
            # PENDENTE" no proprio texto/label); HITL/LIMITATION precisam da
            # marca explicita para o renderer nunca deixar passar disfarcado.
            self.marks = list(self.marks) + [self.status]

    def to_dict(self) -> Dict[str, Any]:
        return dataclasses.asdict(self)


PAGE_PROFILES_VALIDOS = ("cover", "body", "photo")


@dataclasses.dataclass
class Section:
    id: str
    number: str
    title: str
    level: int = 1
    page_setup: str = "A4 retrato"
    # P0.6: perfil estrutural da secao -- "cover" (capa, com titlePg e
    # margens proprias), "body" (corpo A4, o padrao) ou "photo" (fotografico
    # A3). Os VALORES exatos de margem por perfil sao P1 (config/
    # visual_style_v3.yaml, fora do escopo desta rodada); o que P0.6 exige e'
    # so que a capa seja estruturalmente distinguivel do corpo, para o
    # renderer poder aplicar titlePg e um profile de margem proprio a ela --
    # nunca inventar secao A0/A2 (essas nao existem no Word, ver P0.1).
    page_profile: str = "body"
    blocks: List[Block] = dataclasses.field(default_factory=list)

    def __post_init__(self):
        if self.page_profile not in PAGE_PROFILES_VALIDOS:
            raise ValueError("page_profile invalido: %r (validos: %s)"
                             % (self.page_profile, PAGE_PROFILES_VALIDOS))

    def to_dict(self) -> Dict[str, Any]:
        d = dataclasses.asdict(self)
        return d


@dataclasses.dataclass
class Document:
    meta: Dict[str, Any]
    sections: List[Section] = dataclasses.field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return dataclasses.asdict(self)

    def status_counts(self) -> Dict[str, int]:
        """Contagem de blocos por status, para o relatorio de checkpoint."""
        counts = {s: 0 for s in STATUS_VALIDOS}
        for sec in self.sections:
            for b in sec.blocks:
                counts[b.status] = counts.get(b.status, 0) + 1
        return counts

    def section_status(self) -> Dict[str, str]:
        """Status agregado por secao: pior status entre os blocos (LIMITATION
        > HITL > PLACEHOLDER > READY), para a tabela Secao -> status do
        checkpoint. Secao sem blocos -> READY (nada a comprovar)."""
        ordem = {"READY": 0, "PLACEHOLDER": 1, "HITL": 2, "LIMITATION": 3}
        out = {}
        for sec in self.sections:
            pior = "READY"
            for b in sec.blocks:
                if ordem[b.status] > ordem[pior]:
                    pior = b.status
            out["%s %s" % (sec.number, sec.title) if sec.number else sec.title] = pior
        return out


# -----------------------------------------------------------------------------
# fabricas de Block -- ergonomia para o content_builder
# -----------------------------------------------------------------------------
def heading(text, level=1, status="READY", provenance=None, note=None):
    return Block(kind="heading", text=text, level=level, status=status,
                provenance=provenance or [], note=note)


def paragraph(text, status="READY", provenance=None, note=None):
    return Block(kind="paragraph", text=text, status=status,
                provenance=provenance or [], note=note)


def table(headers, rows, status="READY", provenance=None, note=None, role=None):
    return Block(kind="table", headers=headers, rows=rows, status=status,
                provenance=provenance or [], note=note, role=role)


def figure_placeholder(label, status="PLACEHOLDER", provenance=None, note=None):
    return Block(kind="figure_placeholder", label=label, status=status,
                provenance=provenance or [], note=note)


def figure(image_path, width_mm, height_mm, status="READY", provenance=None,
          note=None, label=None):
    """Fase 3.3, contrato Sec.1: uma FiguraResult da Frente A (path/
    largura_mm/altura_mm/status/note/provenance) JA convertida num Block --
    usado quando `gerador.figuras_bridge.tentar()` devolve um resultado com
    `path` preenchido; nunca gerado por content_builder.py sozinho (que nao
    abre planilha nem desenha imagem)."""
    return Block(kind="figure", image_path=image_path, width_mm=width_mm,
                height_mm=height_mm, status=status, provenance=provenance or [],
                note=note, label=label)


def full_page_image(image_path, width_mm, height_mm, status="READY", provenance=None,
                    note=None, label=None):
    """Fase 3.3, Sec.4: uma pagina de anexo EMBEDDED_IMAGE ja renderizada
    (fitz, 200dpi) -- 1 bloco por pagina real do PDF-fonte."""
    return Block(kind="full_page_image", image_path=image_path, width_mm=width_mm,
                height_mm=height_mm, status=status, provenance=provenance or [],
                note=note, label=label)


def caption(text, status="READY", provenance=None):
    return Block(kind="caption", text=text, status=status, provenance=provenance or [])


def list_block(items, status="READY", provenance=None, note=None):
    return Block(kind="list", items=items, status=status,
                provenance=provenance or [], note=note)


def toc(items, status="READY"):
    """`items`: lista de "numero -- titulo" ja formatada (serializavel)."""
    return Block(kind="toc", items=items, status=status)


def page_break():
    return Block(kind="page_break", status="READY")


def cover_field(label, value, status="READY", provenance=None):
    texto = "%s: %s" % (label, value) if value not in (None, "") else "%s: ___" % label
    st = status if value not in (None, "") else "HITL"
    return Block(kind="cover_field", label=label, text=texto, status=st,
                provenance=provenance or [],
                note=None if value not in (None, "") else "campo nao comprovado")
