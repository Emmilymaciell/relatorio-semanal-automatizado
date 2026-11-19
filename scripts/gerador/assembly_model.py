# -*- coding: utf-8 -*-
r"""
scripts/gerador/assembly_model.py - Fase 3.2B, P0.2: camada de MONTAGEM,
separada explicitamente da camada de CONTEUDO.

`content_builder.py` decide O QUE vai no relatorio (document_model.Document).
Este modulo decide COMO o relatorio VIRA ARQUIVO:

    ReportDocument
    |-- main_document            -> DOCX -> PDF do corpo
    |   |-- sections              (A4 corpo, A4 capa, A3 fotografico)
    |   |-- embedded_annexes      anexos EMBEDDED_IMAGE (pag. cheia DENTRO do docx)
    |-- external_pdf_annexes      anexos MERGED_PDF (PDF externo, mesclado DEPOIS)
                |
          merge -> PDF final    (NAO executado nesta rodada -- so o PLANO)

P0.1 (DISCOVERY_CORRECTION): o document.xml do golden tem 9 secoes, todas A4
ou A3 retrato -- NAO existe secao A0 nem A2 no Word. As paginas A0 (unifilar)
e A2 (diario) do PDF final sao PDFs EXTERNOS mesclados depois da exportacao
(pdf_producer do golden = iLovePDF). Esse fato e' exatamente o que este
modulo formaliza: `external_pdf_annexes` nunca diz respeito a uma `Section`
do `main_document` -- e' uma lista paralela.

P0.3: o modo de integracao (EMBEDDED_IMAGE x MERGED_PDF) de cada anexo vem
SEMPRE de `config/annex_types.yaml`, indexado por `id`. Proibido decidir isso
por posicao/numero do anexo em qualquer lugar deste modulo.

Nesta rodada o merge NAO executa -- `montar_plano()` so descreve e serializa
o plano (`AssemblyPlan.to_dict()` -> JSON). Quem quiser executar o merge de
verdade (Checkpoint seguinte) le esse JSON; este modulo nao chama nenhuma
ferramenta de PDF.
"""
import dataclasses
import datetime as dt
import json
import os
from typing import Any, Dict, List, Optional, Sequence

from . import config as C


# -----------------------------------------------------------------------------
# excecao dedicada -- nunca deixar um anexo sem classificacao passar em silencio
# -----------------------------------------------------------------------------
class AnexoNaoClassificadoError(ValueError):
    """Levantada quando um `id` de anexo nao existe em config/annex_types.yaml.
    Preferimos falhar alto a inventar um `integration_mode` por adivinhacao
    (ex.: por posicao) -- e' exatamente o anti-padrao que o P0.3 proibe."""


_MODOS_VALIDOS = ("EMBEDDED_IMAGE", "MERGED_PDF")


def _get(obj, chave, default=None):
    """Le `chave` de `obj`, aceitando tanto dataclass/objeto (atributo) quanto
    dict (chave) -- as `Section` de document_model.py e os dicts de
    `report_data['annexes']` convivem nas chamadas deste modulo."""
    if isinstance(obj, dict):
        return obj.get(chave, default)
    return getattr(obj, chave, default)


def classificar_anexo(anexo_id: str) -> Dict[str, Any]:
    """Devolve a entrada de `config/annex_types.yaml#anexos` para `anexo_id`.
    Levanta `AnexoNaoClassificadoError` se o id nao estiver configurado --
    nunca inferir o modo por posicao/numero."""
    tipos = C.annex_types().get("anexos", {})
    entrada = tipos.get(anexo_id)
    if entrada is None:
        raise AnexoNaoClassificadoError(
            "anexo sem classificacao em config/annex_types.yaml: %r "
            "(tipos conhecidos: %s)" % (anexo_id, sorted(tipos)))
    modo = entrada.get("integration_mode")
    if modo not in _MODOS_VALIDOS:
        raise AnexoNaoClassificadoError(
            "integration_mode invalido para %r: %r (validos: %s)"
            % (anexo_id, modo, _MODOS_VALIDOS))
    return entrada


# -----------------------------------------------------------------------------
# dataclasses do plano -- tudo serializavel em JSON, nada de objeto Word/PDF
# -----------------------------------------------------------------------------
@dataclasses.dataclass
class SectionRef:
    """Referencia leve a uma `document_model.Section` -- so o que a montagem
    precisa saber (papel/pagina), nunca o conteudo dos blocos."""
    id: str
    number: str
    title: str
    page_setup: str

    def to_dict(self) -> Dict[str, Any]:
        return dataclasses.asdict(self)


@dataclasses.dataclass
class EmbeddedAnnexPlan:
    """Anexo EMBEDDED_IMAGE: vira imagem de pagina cheia DENTRO do
    main_document -- herda cabecalho e numeracao continua do relatorio."""
    anexo_id: str
    titulo: str
    ordem: Optional[int]
    presente: bool
    source_path: Optional[str] = None
    n_paginas_observado: Optional[int] = None
    extra: Dict[str, Any] = dataclasses.field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return dataclasses.asdict(self)


@dataclasses.dataclass
class ExternalPdfAnnexPlan:
    """Anexo MERGED_PDF: PDF externo, mesclado DEPOIS da exportacao DOCX->PDF.
    Nunca faz parte de `main_document.sections` -- nao existe secao A0/A2 no
    Word (P0.1). `merge_executed` fica sempre False nesta rodada."""
    anexo_id: str
    titulo: str
    posicao: Optional[str]
    papel_pdf: Optional[str]
    presente: bool
    source_path: Optional[str] = None
    n_paginas_esperado: Optional[int] = None
    cabecalho_e_numeracao: bool = False
    merge_executed: bool = False
    extra: Dict[str, Any] = dataclasses.field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return dataclasses.asdict(self)


@dataclasses.dataclass
class MainDocumentPlan:
    """O que vira o .docx (e depois, fora desta rodada, o PDF do corpo)."""
    sections: List[SectionRef] = dataclasses.field(default_factory=list)
    embedded_annexes: List[EmbeddedAnnexPlan] = dataclasses.field(default_factory=list)
    output_docx_path: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "sections": [s.to_dict() for s in self.sections],
            "embedded_annexes": [a.to_dict() for a in self.embedded_annexes],
            "output_docx_path": self.output_docx_path,
        }


@dataclasses.dataclass
class AssemblyPlan:
    """`ReportDocument` da spec P0.2: main_document + external_pdf_annexes.
    O merge para o PDF final NAO executa nesta rodada -- so descreve o plano.
    """
    report_id: str
    obra: str
    periodo: Dict[str, str]
    main_document: MainDocumentPlan
    external_pdf_annexes: List[ExternalPdfAnnexPlan] = dataclasses.field(default_factory=list)
    merge_executed: bool = False
    merge_note: str = (
        "P0.2: montagem descreve o plano, NAO executa o merge com os PDFs "
        "externos (Checkpoint 3.2B). Ver docs/arquitetura/FASE3_2B_RENDERER_V2_SPEC.md#P0.2."
    )
    gerado_em: str = dataclasses.field(
        default_factory=lambda: dt.datetime.now().astimezone().isoformat(timespec="seconds"))

    def to_dict(self) -> Dict[str, Any]:
        return {
            "report_id": self.report_id,
            "obra": self.obra,
            "periodo": self.periodo,
            "main_document": self.main_document.to_dict(),
            "external_pdf_annexes": [a.to_dict() for a in self.external_pdf_annexes],
            "merge_executed": self.merge_executed,
            "merge_note": self.merge_note,
            "gerado_em": self.gerado_em,
        }


# -----------------------------------------------------------------------------
# construcao do plano
# -----------------------------------------------------------------------------
def _titulo_anexo(anexo_id: str, entrada_config: Dict[str, Any], anexo_report_data: Dict[str, Any]) -> str:
    # o titulo de exibicao (com sufixo tipo "PA-254") continua sendo
    # responsabilidade do content_builder (secao 11 / lista de anexos);
    # aqui usamos o titulo canonico do tipo, para identificar o anexo no
    # plano de montagem.
    return entrada_config.get("titulo") or anexo_id


def montar_plano(*, report_id: str, obra: str, periodo: Dict[str, str],
                 sections: Sequence[Any], annexes: Sequence[Dict[str, Any]],
                 output_docx_path: Optional[str] = None) -> AssemblyPlan:
    """Monta o `AssemblyPlan` a partir das `sections` do document_model
    (Fase 3.2B nao muda o conteudo) e da lista de anexos que
    `extratores/anexos.montar()` (via `report_data['annexes']`) ja produz.

    `sections`: sequencia de `document_model.Section` OU dicts equivalentes
    (id/number/title/page_setup) -- ver `_get()`.
    `annexes`: sequencia de dicts com pelo menos a chave `id` (formato de
    `report_data['annexes']`).

    Classifica cada anexo por `classificar_anexo()` (config/annex_types.yaml)
    -- NUNCA por posicao/numero. Um anexo com id desconhecido faz a funcao
    falhar (`AnexoNaoClassificadoError`) em vez de ser silenciosamente
    ignorado ou classificado por adivinhacao.
    """
    section_refs = [
        SectionRef(id=_get(s, "id"), number=_get(s, "number"),
                  title=_get(s, "title"), page_setup=_get(s, "page_setup"))
        for s in sections
    ]

    embedded: List[EmbeddedAnnexPlan] = []
    external: List[ExternalPdfAnnexPlan] = []

    for a in annexes:
        anexo_id = a["id"]
        cfg = classificar_anexo(anexo_id)
        titulo = _titulo_anexo(anexo_id, cfg, a)
        presente = bool(a.get("presente")) if not isinstance(a.get("presente"), str) else False
        # "presente" pode vir como True/False ou "HITL" (parecer_tecnico
        # ambiguo, ver extratores/anexos.py) -- tratamos "HITL" como
        # nao-confirmado para fins de montagem (nao e' False, mas tambem
        # nao e' True sem revisao humana).
        extra = {}
        if a.get("presente") == "HITL":
            extra["presenca"] = "HITL"
            extra["candidatos"] = a.get("candidatos")

        if cfg["integration_mode"] == "EMBEDDED_IMAGE":
            embedded.append(EmbeddedAnnexPlan(
                anexo_id=anexo_id, titulo=titulo, ordem=cfg.get("ordem"),
                presente=presente, source_path=a.get("path"),
                n_paginas_observado=cfg.get("n_paginas_observado"),
                extra=extra,
            ))
        else:  # MERGED_PDF
            papel = a.get("papel_pdf") or cfg.get("papel_pdf_default")
            external.append(ExternalPdfAnnexPlan(
                anexo_id=anexo_id, titulo=titulo, posicao=cfg.get("posicao"),
                papel_pdf=papel, presente=presente, source_path=a.get("path"),
                n_paginas_esperado=cfg.get("n_paginas_observado"),
                cabecalho_e_numeracao=bool(cfg.get("cabecalho_e_numeracao", False)),
                merge_executed=False, extra=extra,
            ))

    embedded.sort(key=lambda e: (e.ordem if e.ordem is not None else 999))
    ordem_posicao = {"penultimo": 0, "ultimo": 1}
    external.sort(key=lambda e: ordem_posicao.get(e.posicao, 0.5))

    main_doc = MainDocumentPlan(sections=section_refs, embedded_annexes=embedded,
                                output_docx_path=output_docx_path)
    return AssemblyPlan(report_id=report_id, obra=obra, periodo=dict(periodo),
                        main_document=main_doc, external_pdf_annexes=external,
                        merge_executed=False)


def serializar(plano: AssemblyPlan, output_path: str) -> str:
    """Grava `plano` como JSON em `output_path`. Import tardio de
    `prog_util` para nao criar dependencia de import no topo do modulo (este
    arquivo tambem e' usado a partir de `tests/`, fora do pacote `scripts`)."""
    import sys
    aqui = os.path.dirname(os.path.abspath(__file__))
    scripts_dir = os.path.dirname(aqui)
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    import prog_util as U  # noqa: E402

    U.assert_read_only(output_path)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as fh:
        json.dump(plano.to_dict(), fh, ensure_ascii=False, indent=2, default=str)
    return output_path
