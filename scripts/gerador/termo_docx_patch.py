# -*- coding: utf-8 -*-
r"""
scripts/gerador/termo_docx_patch.py - Fase 3.3B: atualiza o paragrafo do
Termo de Encerramento DENTRO do
`shadow_final_candidate_v{versao}.docx` ja gerado, com o numero real de
paginas (so conhecido depois da montagem do PDF completo).

Isto NAO substitui uma nova rodada do pipeline content_builder ->
docx_renderer_v2 (que exigiria reabrir report_data.json e regerar as 43 MB
do documento do zero, incluindo as 24 imagens). Em vez disso, edita
CIRURGICAMENTE, via `python-docx`, so o run do paragrafo do Termo e remove
a nota "n_paginas_provisorio=true" -- o resto do documento fica bit-a-bit
igual.

Por que precisa disto: o Termo definitivo exige o PDF ja montado (Sec.5 da
FASE3_3_FINAL_CANDIDATE_SPEC.md), mas o DOCX ja foi gerado ANTES da
montagem (Checkpoint 3.3A). Deixar o `.docx` "pronto" (com o numero real)
e' o que a Emmily pediu explicitamente na Fase 3.3B, mesmo sabendo que
reconverter esse `.docx` -> PDF exige Word (COM nao funciona aqui) -- essa
reconversao fica para quando ela quiser rodar manualmente de novo.
"""
import os
import sys

from docx.enum.text import WD_COLOR_INDEX

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(AQUI)
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402


def _e_heading_termo(paragraph):
    return (paragraph.style is not None
            and (paragraph.style.name or "").startswith("Heading")
            and paragraph.text.strip().upper() == "TERMO DE ENCERRAMENTO")


def atualizar_termo_docx(docx_path, *, texto_final):
    """Abre `docx_path` (NAO pode estar dentro do SharePoint -- e' sempre um
    output/shadow/... deste projeto), localiza o paragrafo de heading
    "TERMO DE ENCERRAMENTO" (estilo Heading, para nao confundir com a
    citacao homonima no INDICE), reescreve o proximo paragrafo (o run unico
    do template `s10_termo`) com `texto_final` -- sem o prefixo
    "[PENDENTE] ", sem destaque amarelo, sem negrito -- e remove o
    paragrafo de nota "Nota: n_paginas_provisorio=true..." logo abaixo, se
    existir. Grava por cima do MESMO arquivo (e' um artefato deste projeto,
    nao o SharePoint).

    Devolve um dict de diagnostico -- nunca finge sucesso: se o heading ou
    o paragrafo-alvo nao forem encontrados, `ok=False` e o arquivo NAO e'
    reescrito."""
    U.assert_read_only(docx_path)  # defesa: nunca deveria apontar pro SharePoint
    import docx  # import tardio -- so quando de fato vai editar

    documento = docx.Document(docx_path)
    paragrafos = documento.paragraphs

    idx_heading = None
    for i, p in enumerate(paragrafos):
        if _e_heading_termo(p):
            idx_heading = i
            break

    if idx_heading is None:
        return {"ok": False, "motivo": "heading 'TERMO DE ENCERRAMENTO' (estilo Heading) "
                                        "não encontrado no DOCX"}

    idx_paragrafo = idx_heading + 1
    if idx_paragrafo >= len(paragrafos):
        return {"ok": False, "motivo": "não há parágrafo após o heading do Termo"}

    p_termo = paragrafos[idx_paragrafo]
    if not p_termo.runs:
        return {"ok": False, "motivo": "parágrafo do Termo não tem nenhum run"}

    texto_antes = p_termo.runs[0].text
    r = p_termo.runs[0]
    # zera runs extras do mesmo paragrafo (o template so cria 1 run; isto e
    # so uma salvaguarda caso alguem tenha adicionado mais no futuro)
    for extra in p_termo.runs[1:]:
        extra.text = ""
    r.text = texto_final
    r.font.highlight_color = None
    r.bold = False

    nota_removida = False
    idx_nota = idx_paragrafo + 1
    if idx_nota < len(paragrafos):
        p_nota = paragrafos[idx_nota]
        if p_nota.text.strip().startswith("Nota: n_paginas_provisorio"):
            el = p_nota._p
            el.getparent().remove(el)
            nota_removida = True

    documento.save(docx_path)
    return {"ok": True, "texto_anterior": texto_antes, "texto_novo": texto_final,
            "nota_removida": nota_removida}
