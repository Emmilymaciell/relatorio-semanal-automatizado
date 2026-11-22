# -*- coding: utf-8 -*-
"""
test_docx_renderer.py

Testa scripts/gerador/renderers/docx_renderer.py (Fase 3, Slice 3, spec §8
item 3):

  R1  o .docx gerado ABRE (python-docx consegue reabrir o arquivo salvo)
  R2  numero de secoes de pagina no docx bate com a quantidade de trocas de
      (`page_setup`, `page_profile`) no document_model (A4 -> A3 -> A4 para a
      secao 5.2; e a capa -- page_profile="cover" -- conta como troca mesmo
      sendo A4 igual ao corpo, P0.6)
  R3  titulos saem na ordem (heading paragraphs do docx, na ordem do
      document_model, comparando o texto sem o prefixo de marca)
  R4  todo `figure_placeholder` do document_model aparece no texto do docx
      (o rotulo "RENDERIZAÇÃO PENDENTE" nao se perde na renderizacao)
  R5  nenhum bloco HITL sai sem marca: todo texto de bloco HITL no
      document_model aparece PRECEDIDO do prefixo "[HITL]" no paragrafo
      correspondente do docx
  R6  idem para LIMITATION ("[LIMITAÇÃO DA FONTE]") e PLACEHOLDER (realce
      amarelo aplicado -- checado via WD_COLOR_INDEX.YELLOW no run)

Execute:  python tests/test_docx_renderer.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import sys
import json

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
import docx  # noqa: E402
from docx.enum.text import WD_COLOR_INDEX  # noqa: E402
from gerador import content_builder as CB  # noqa: E402
from gerador.renderers import docx_renderer  # noqa: E402

OBRA = "RD103"
PERIODO = ("2026-08-29", "2026-09-04")


def main():
    falhas = []
    caminho_data = os.path.join(U.OUT, "shadow", "%s_%s_%s" % (OBRA, *PERIODO), "report_data.json")
    if not os.path.exists(caminho_data):
        print("SKIP: rode 'python scripts/shadow_report.py data --obra RD103 --inicio "
             "2026-08-29 --fim 2026-09-04' antes deste teste")
        return 1
    with open(caminho_data, encoding="utf-8") as fh:
        report_data = json.load(fh)

    documento = CB.construir_documento(report_data)
    out_path = os.path.join(U.PROJ, "tmp", "test_docx_renderer_output.docx")
    docx_renderer.render(documento, out_path)

    # ---------------------------------------------------------------- R1
    try:
        d = docx.Document(out_path)
        print("R1  OK   %s abre sem erro (python-docx)" % out_path)
    except Exception as e:
        falhas.append("R1: docx nao abriu: %s: %s" % (type(e).__name__, e))
        print("FALHAS: R1 impede o restante do teste"); print("RESULTADO: FALHOU"); return 1

    # ---------------------------------------------------------------- R2
    def _perfil(s):
        return (s.page_setup, getattr(s, "page_profile", "body"))

    trocas_esperadas = 1
    ultimo = _perfil(documento.sections[0])
    for s in documento.sections[1:]:
        if _perfil(s) != ultimo:
            trocas_esperadas += 1
            ultimo = _perfil(s)
    if len(d.sections) == trocas_esperadas:
        print("R2  OK   %d secoes de pagina no docx == %d trocas de (page_setup, page_profile) no modelo"
             % (len(d.sections), trocas_esperadas))
    else:
        falhas.append("R2: docx tem %d secoes, esperado %d (trocas de page_setup/page_profile)"
                      % (len(d.sections), trocas_esperadas))

    # ---------------------------------------------------------------- R3
    titulos_modelo = [b.text for s in documento.sections for b in s.blocks if b.kind == "heading"]
    titulos_docx = [p.text.strip() for p in d.paragraphs if p.style.name.startswith("Heading")]

    def _sem_marca(t):
        for prefixo in ("[HITL] ", "[LIMITAÇÃO DA FONTE] ", "[PENDENTE] "):
            if t.startswith(prefixo):
                return t[len(prefixo):]
        return t

    titulos_docx_limpos = [_sem_marca(t) for t in titulos_docx]
    if titulos_docx_limpos == titulos_modelo:
        print("R3  OK   %d titulos saem na mesma ordem do document_model" % len(titulos_modelo))
    else:
        falhas.append("R3: titulos do docx == %s, esperado %s" % (titulos_docx_limpos, titulos_modelo))

    # ---------------------------------------------------------------- R4
    labels_modelo = [b.label for s in documento.sections for b in s.blocks
                     if b.kind == "figure_placeholder"]
    texto_docx_completo = "\n".join(p.text for p in d.paragraphs)
    faltando = [lbl for lbl in labels_modelo if lbl not in texto_docx_completo]
    if labels_modelo and not faltando:
        print("R4  OK   todos os %d figure_placeholder do modelo aparecem no texto do docx"
             % len(labels_modelo))
    else:
        falhas.append("R4: placeholders ausentes do docx: %s" % faltando)

    # ---------------------------------------------------------------- R5 / R6
    blocos_hitl = [b for s in documento.sections for b in s.blocks
                  if b.status == "HITL" and b.kind in ("paragraph", "heading")]
    blocos_limitation = [b for s in documento.sections for b in s.blocks
                         if b.status == "LIMITATION" and b.kind in ("paragraph", "heading")]

    def _paragrafo_com_texto_contendo(fragmento):
        """Acha o paragrafo que corresponde a ESTE bloco.

        Casar por `fragmento[:40]` era ambiguo: desde a Fase 3.4 (v0.2) o texto
        da secao 6 contem a mesma abertura "Durante o periodo abrangido por este
        relatorio..." que a frase negativa da secao 9, e o primeiro paragrafo
        encontrado (o da secao 6, READY) fazia o teste acusar falsamente a
        secao 9 de sair sem marca -- quando o docx trazia o [HITL] corretamente.
        Agora casa pelo texto INTEIRO do bloco, e so cai no prefixo longo se
        o texto tiver sido truncado na renderizacao."""
        if not fragmento:
            return None
        for p in d.paragraphs:
            if fragmento in p.text:
                return p
        chave = fragmento[:120]
        for p in d.paragraphs:
            if chave in p.text:
                return p
        return None

    sem_marca_hitl = []
    for b in blocos_hitl:
        p = _paragrafo_com_texto_contendo(b.text)
        if not p or "[HITL]" not in p.text:
            sem_marca_hitl.append(b.text)
    if blocos_hitl and not sem_marca_hitl:
        print("R5  OK   todos os %d blocos HITL saem com o prefixo [HITL] no docx" % len(blocos_hitl))
    else:
        falhas.append("R5: blocos HITL sem marca no docx: %s" % sem_marca_hitl)

    sem_marca_lim = []
    for b in blocos_limitation:
        p = _paragrafo_com_texto_contendo(b.text)
        if not p or "[LIMITAÇÃO DA FONTE]" not in p.text:
            sem_marca_lim.append(b.text)
    if blocos_limitation and not sem_marca_lim:
        print("R6  OK   todos os %d blocos LIMITATION saem com o prefixo [LIMITAÇÃO DA FONTE]"
             % len(blocos_limitation))
    else:
        falhas.append("R6: blocos LIMITATION sem marca no docx: %s" % sem_marca_lim)

    # realce amarelo em pelo menos um paragrafo de figure_placeholder (PLACEHOLDER)
    algum_realce_placeholder = False
    for p in d.paragraphs:
        if "RENDERIZAÇÃO PENDENTE" in p.text:
            for run in p.runs:
                if run.font.highlight_color == WD_COLOR_INDEX.YELLOW:
                    algum_realce_placeholder = True
    if algum_realce_placeholder:
        print("R6b OK   figure_placeholder sai com realce amarelo (marca visual, spec Sec.5)")
    else:
        falhas.append("R6b: nenhum paragrafo de figure_placeholder com realce amarelo")

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
