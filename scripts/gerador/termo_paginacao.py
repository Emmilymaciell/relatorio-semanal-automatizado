# -*- coding: utf-8 -*-
r"""
scripts/gerador/termo_paginacao.py - Fase 3.3, Sec.5: mecânica de "duas
passagens" do Termo de Encerramento.

    PASS 1  monta DOCX -> PDF do corpo -> + PDFs externos -> conta páginas reais
    PASS 2  atualiza o Termo com o número -> re-renderiza -> confere

Máximo 3 passagens; se não estabilizar, `PAGINATION_UNSTABLE`.

Word/Excel COM NÃO FUNCIONAM neste ambiente (CLAUDE.md, regra global; §0 da
FASE3_3_FINAL_CANDIDATE_SPEC.md) -- a conversão DOCX->PDF por COM está fora
de cogitação, e este módulo não tenta. `rodar()` implementa A MECÂNICA (o
loop de convergência, o teto de 3 passagens, a detecção de instabilidade) de
forma agnóstica a COMO as páginas são contadas: recebe um
`contador_paginas() -> Optional[int]` já resolvido por quem chama (hoje,
sempre `None`, porque não há PDF do corpo nesta rodada -- a conversão manual
é o Passo 8 da spec). Quando não há NENHUMA forma de contar páginas reais
disponível, `rodar()` devolve IMEDIATAMENTE `resultado="SEM_PDF"`: o Termo
fica PLACEHOLDER com `n_paginas_provisorio=true` (já implementado em
content_builder._secao10_termo) -- nunca inventa um número.
"""
import dataclasses
from typing import Callable, List, Optional

MAX_PASSAGENS = 3


@dataclasses.dataclass
class ResultadoPaginacao:
    resultado: str  # "SEM_PDF" | "ESTAVEL" | "PAGINATION_UNSTABLE"
    n_paginas: Optional[int] = None
    passagens: int = 0
    historico: List[Optional[int]] = dataclasses.field(default_factory=list)
    nota: str = ""

    def to_dict(self):
        return dataclasses.asdict(self)


def rodar(contador_paginas: Optional[Callable[[], Optional[int]]]) -> ResultadoPaginacao:
    """`contador_paginas`: callable SEM argumentos que devolve o número de
    páginas do PDF final NESTA passagem (corpo + anexos externos MERGED_PDF),
    ou `None` se ainda não houver PDF disponível. Cada chamada representa o
    ciclo completo de uma passagem: atualizar o Termo com o número da
    passagem anterior -> re-renderizar -> re-contar.

    `contador_paginas=None` (o caso desta rodada, Checkpoint 3.3A: sem
    conversão DOCX->PDF disponível) devolve `SEM_PDF` sem nenhuma iteração --
    é a MESMA saída de content_builder._secao10_termo quando `n_paginas is
    None`, só que auditável e serializável para o manifesto (Sec.6)."""
    if contador_paginas is None:
        return ResultadoPaginacao(
            resultado="SEM_PDF", n_paginas=None, passagens=0, historico=[],
            nota="nenhum contador de páginas disponível nesta rodada (sem PDF do corpo, "
                "conversão DOCX->PDF é manual -- Passo 8 da spec); Termo fica PLACEHOLDER "
                "com n_paginas_provisorio=true")

    historico = []
    anterior = None
    for i in range(1, MAX_PASSAGENS + 1):
        atual = contador_paginas()
        historico.append(atual)
        if atual is None:
            return ResultadoPaginacao(
                resultado="SEM_PDF", n_paginas=None, passagens=i, historico=historico,
                nota="contador devolveu None na passagem %d" % i)
        if anterior is not None and atual == anterior:
            return ResultadoPaginacao(
                resultado="ESTAVEL", n_paginas=atual, passagens=i, historico=historico,
                nota="paginação estabilizou em %d páginas na passagem %d" % (atual, i))
        anterior = atual

    return ResultadoPaginacao(
        resultado="PAGINATION_UNSTABLE", n_paginas=anterior, passagens=MAX_PASSAGENS,
        historico=historico,
        nota="não estabilizou em %d passagens (histórico: %s) -- registrado e seguido, "
            "conforme Sec.5" % (MAX_PASSAGENS, historico))
