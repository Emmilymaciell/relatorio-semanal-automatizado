# -*- coding: utf-8 -*-
"""
test_toc_titulo_indice.py — TEXTO-10 (Fase 3.4, v0.3)

Prova, no OOXML, que o título "ÍNDICE":

  T1  está VISUALMENTE presente no documento (o leitor vê o título)
  T2  NÃO usa "Heading 1"/"Heading 2"/"Heading 3" — usa o estilo próprio
      `TOC_TITLE`
  T3  o estilo `TOC_TITLE` NÃO deriva de um estilo de título e traz
      outlineLvl=9 na própria definição do estilo
  T4  o parágrafo traz outlineLvl=9 e numId=0 (não numerado)
  T5  visualmente equivalente ao h1 configurado (fonte, corpo, cor) e
      centralizado (VIS-04)
  T6  o campo TOC continua sendo `TOC \\o "1-3"` e NENHUM parágrafo
      alcançável por esse switch é o "ÍNDICE" — isto é, depois que o Word
      atualizar o campo, "ÍNDICE" não pode aparecer entre as entradas

POR QUE T2 EXISTE
-----------------
A v0.2 tentou resolver o mesmo bug com `outlineLvl=9` direto no parágrafo,
mantendo o estilo "Heading 1". Não funcionou: o switch `\\o` do campo TOC
coleta pelos ESTILOS DE TÍTULO EMBUTIDOS, usando o outline level DO ESTILO —
formatação direta no parágrafo não tira o parágrafo dessa varredura. O PDF
convertido pelo Word continuou trazendo "ÍNDICE ..... 2" como primeira
entrada do próprio índice, com o cache do campo vazio (ou seja, o Word
recalculou mesmo).

T2 é, portanto, o teste que a v0.2 teria reprovado. Se alguém voltar o título
para "Heading 1" achando que o outlineLvl basta, este teste acusa.

Execute:  python tests/test_toc_titulo_indice.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import re
import sys
import json
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
import yaml  # noqa: E402
import docx  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from gerador import content_builder as CB  # noqa: E402
from gerador.renderers import docx_renderer_v2 as R  # noqa: E402

OBRA, PERIODO = "RD103", ("2026-08-29", "2026-09-04")
ESTILOS_DE_TITULO = {"Heading 1", "Heading 2", "Heading 3",
                     "Título 1", "Título 2", "Título 3"}


def main():
    falhas = []
    caminho_data = os.path.join(U.OUT, "shadow", "%s_%s_%s" % (OBRA, *PERIODO),
                                "report_data.json")
    if not os.path.exists(caminho_data):
        print("SKIP: rode 'python scripts/shadow_report.py data --obra RD103 "
              "--inicio 2026-08-29 --fim 2026-09-04' antes deste teste")
        return 1
    with open(caminho_data, encoding="utf-8") as fh:
        report_data = json.load(fh)
    vs = yaml.safe_load(open(os.path.join(U.PROJ, "config", "visual_style_v3.yaml"),
                             encoding="utf-8"))

    documento = CB.construir_documento(report_data)
    out = os.path.join(U.PROJ, "tmp", "test_toc_titulo_indice.docx")
    R.render(documento, out)

    d = docx.Document(out)
    par_indice = [p for p in d.paragraphs if p.text.strip() == "ÍNDICE"]

    # ------------------------------------------------------------------ T1
    if len(par_indice) == 1:
        print("T1  OK   'ÍNDICE' presente no documento, uma vez")
    else:
        falhas.append("T1: esperado 1 parágrafo 'ÍNDICE', achei %d" % len(par_indice))
        print("FALHAS: T1 impede o restante"); print("RESULTADO: FALHOU"); return 1
    p = par_indice[0]

    # ------------------------------------------------------------------ T2
    if p.style.name in ESTILOS_DE_TITULO:
        falhas.append("T2: 'ÍNDICE' usa o estilo de título %r -- o switch \\o do campo "
                      "TOC vai alcançá-lo e ele volta a se auto-listar (era exatamente "
                      "o bug da v0.2)" % p.style.name)
    elif p.style.name != R.ESTILO_TITULO_INDICE:
        falhas.append("T2: 'ÍNDICE' usa %r, esperado %r"
                      % (p.style.name, R.ESTILO_TITULO_INDICE))
    else:
        print("T2  OK   'ÍNDICE' usa o estilo %r, que não é Heading 1/2/3"
              % R.ESTILO_TITULO_INDICE)

    # ------------------------------------------------------------------ T3
    try:
        st = d.styles[R.ESTILO_TITULO_INDICE]
        base = st.base_style.name if st.base_style is not None else None
        if base in ESTILOS_DE_TITULO:
            falhas.append("T3: %s deriva de %r (estilo de título) -- herdaria o "
                          "outline level e voltaria ao campo"
                          % (R.ESTILO_TITULO_INDICE, base))
        pPr_st = st.element.find(qn('w:pPr'))
        lvl = None
        if pPr_st is not None:
            el = pPr_st.find(qn('w:outlineLvl'))
            if el is not None:
                lvl = el.get(qn('w:val'))
        if lvl != "9":
            falhas.append("T3: outlineLvl do estilo %s == %r, esperado '9'"
                          % (R.ESTILO_TITULO_INDICE, lvl))
        if not falhas or all(not f.startswith("T3") for f in falhas):
            print("T3  OK   estilo deriva de %r e traz outlineLvl=9 na definição" % base)
    except KeyError:
        falhas.append("T3: estilo %s não existe no documento" % R.ESTILO_TITULO_INDICE)

    # ------------------------------------------------------------------ T4
    pPr = p._p.find(qn('w:pPr'))
    lvl_par = None
    numid = None
    if pPr is not None:
        el = pPr.find(qn('w:outlineLvl'))
        if el is not None:
            lvl_par = el.get(qn('w:val'))
        numPr = pPr.find(qn('w:numPr'))
        if numPr is not None:
            n = numPr.find(qn('w:numId'))
            if n is not None:
                numid = n.get(qn('w:val'))
    erros_t4 = []
    if lvl_par != "9":
        erros_t4.append("outlineLvl do parágrafo == %r, esperado '9'" % lvl_par)
    if numid != "0":
        erros_t4.append("numId == %r, esperado '0' (sem numeração)" % numid)
    if erros_t4:
        falhas.append("T4: " + "; ".join(erros_t4))
    else:
        print("T4  OK   parágrafo com outlineLvl=9 e numId=0")

    # ------------------------------------------------------------------ T5
    erros_t5 = []
    esperado_pt = vs["h1"]["fonte_pt"]
    obtido_pt = st.font.size.pt if st.font.size is not None else None
    if obtido_pt != esperado_pt:
        erros_t5.append("corpo %s pt, esperado %s" % (obtido_pt, esperado_pt))
    cor = st.font.color.rgb if st.font.color and st.font.color.rgb else None
    if str(cor) != vs["h1"]["cor"]:
        erros_t5.append("cor %s, esperado %s" % (cor, vs["h1"]["cor"]))
    if st.font.name != vs["fonts"]["principal"]:
        erros_t5.append("fonte %s, esperado %s" % (st.font.name, vs["fonts"]["principal"]))
    if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
        erros_t5.append("alinhamento %s, esperado CENTER (VIS-04)" % p.alignment)
    if erros_t5:
        falhas.append("T5: " + "; ".join(erros_t5))
    else:
        print("T5  OK   visualmente equivalente ao h1 (%s %s pt %s) e centralizado"
              % (vs["fonts"]["principal"], esperado_pt, vs["h1"]["cor"]))

    # ------------------------------------------------------------------ T6
    z = zipfile.ZipFile(out)
    xml = z.read("word/document.xml").decode("utf-8")
    instr = re.findall(r'<w:instrText[^>]*>([^<]*TOC[^<]*)</w:instrText>', xml)
    if not instr:
        falhas.append("T6: campo TOC não encontrado no document.xml")
    else:
        switch_o = re.search(r'\\o\s*"(\d)-(\d)"', instr[0])
        if switch_o:
            lo, hi = int(switch_o.group(1)), int(switch_o.group(2))
            # todo paragrafo que o `\o` alcanca = os que usam Heading lo..hi
            estilos_coletados = {"Heading %d" % n for n in range(lo, hi + 1)}
            coletados = [q.text.strip() for q in d.paragraphs
                         if q.style.name in estilos_coletados and q.text.strip()]
            if "ÍNDICE" in coletados:
                falhas.append("T6: 'ÍNDICE' está entre os %d parágrafos que o switch "
                              "\\o \"%d-%d\" coleta -- vai se auto-listar quando o Word "
                              "atualizar o campo" % (len(coletados), lo, hi))
            else:
                print("T6  OK   campo é %s; 'ÍNDICE' NÃO está entre os %d parágrafos "
                      "alcançados pelo switch" % (instr[0].strip(), len(coletados)))
        else:
            print("T6  OK   campo TOC sem switch \\o (%s) -- coleta por outline level, "
                  "e o 'ÍNDICE' está em 9" % instr[0].strip())

    print("=" * 74)
    if falhas:
        print("FALHAS (%d):" % len(falhas))
        for f in falhas:
            print("  X", f)
        print("RESULTADO: FALHOU")
        return 1
    print("RESULTADO: OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
