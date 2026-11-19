# -*- coding: utf-8 -*-
"""
test_content_builder.py

Testa scripts/gerador/content_builder.py (Fase 3, Slice 3, spec §8 item 2):

  B1  as 11 secoes numeradas (1..11) aparecem, na ordem, alem de capa e indice
  B2  Q-9: 1a ocorrencia do servico com sigla sai "Nome (SIGLA)" (secao 1),
      ocorrencias seguintes so a sigla (secao 4) -- nao heuristica, o
      rastreador e' compartilhado no documento inteiro
  B3  frase de dia sem execucao (30/08/2026) presente na secao 5.1
  B4  Q-8: dia sem Ficha de Campo (30/08) renderiza "-" na tabela climatica e
      o bloco da tabela sai HITL com nota -- nunca "Bom" nem repeticao do
      dia anterior
  B5  Q-10: secao 9 sem ocorrencia gera a frase negativa, HITL, nao bloqueia
      a montagem (documento continua tendo as 14 secoes)
  B6  document.status_counts() bate com a contagem manual de blocos
  B7  roda para as obras com report_data.json disponivel em output/shadow/,
      sem excecao e sem nenhum `if obra == `　na propria funcao (checagem
      estrutural: construir_documento nao referencia strings de obra)

Execute:  python tests/test_content_builder.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import re
import sys
import json
import inspect

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from gerador import content_builder as CB  # noqa: E402

PERIODO = ("2026-08-29", "2026-09-04")
SECOES_NUMERADAS_ESPERADAS = [str(n) for n in range(1, 12)]


def _carregar_report_data(obra):
    caminho = os.path.join(U.OUT, "shadow", "%s_%s_%s" % (obra, *PERIODO), "report_data.json")
    if not os.path.exists(caminho):
        return None
    with open(caminho, encoding="utf-8") as fh:
        return json.load(fh)


def main():
    falhas = []
    report_data = _carregar_report_data("RD103")
    if report_data is None:
        print("SKIP: rode 'python scripts/shadow_report.py data --obra RD103 --inicio "
             "2026-08-29 --fim 2026-09-04' antes deste teste")
        return 1

    doc = CB.construir_documento(report_data)

    # ---------------------------------------------------------------- B1
    numeros = [s.number for s in doc.sections if s.number]
    numeros_topo_unicos_em_ordem = []
    for n in numeros:
        raiz = n.split(".")[0]
        if raiz not in numeros_topo_unicos_em_ordem:
            numeros_topo_unicos_em_ordem.append(raiz)
    if numeros_topo_unicos_em_ordem == SECOES_NUMERADAS_ESPERADAS:
        print("B1  OK   11 secoes numeradas (1..11), na ordem: %s" % numeros_topo_unicos_em_ordem)
    else:
        falhas.append("B1: ordem das secoes numeradas == %s, esperado %s"
                      % (numeros_topo_unicos_em_ordem, SECOES_NUMERADAS_ESPERADAS))
    ids = [s.id for s in doc.sections]
    if "capa" in ids and "indice" in ids and ids.index("capa") < ids.index("indice") < ids.index("s1_apresentacao"):
        print("B1b OK   capa e indice presentes, antes da secao 1")
    else:
        falhas.append("B1b: ids das secoes == %s" % ids)

    # ---------------------------------------------------------------- B2 (Q-9)
    sec1 = next(s for s in doc.sections if s.id == "s1_apresentacao")
    sec4 = next(s for s in doc.sections if s.id == "s4_unifilar")
    texto_s1 = " ".join(b.text for b in sec1.blocks if b.kind == "paragraph" and b.text)
    texto_s4 = " ".join(b.text for b in sec4.blocks if b.kind == "paragraph" and b.text)
    tem_oae_completo_s1 = "Obras de Arte Especiais (OAE)" in texto_s1
    tem_oae_completo_s4 = "Obras de Arte Especiais (OAE)" in texto_s4
    tem_oae_sigla_s4 = re.search(r"(?<!\()\bOAE\b(?!\))", texto_s4) is not None
    if tem_oae_completo_s1 and not tem_oae_completo_s4 and tem_oae_sigla_s4:
        print("B2  OK   Q-9: 1a ocorrencia 'Obras de Arte Especiais (OAE)' na secao 1, "
             "so 'OAE' na secao 4")
    else:
        falhas.append("B2: 1a ocorrencia incorreta -- s1=%r s4=%r" % (texto_s1, texto_s4))

    # ---------------------------------------------------------------- B3
    sec5 = next(s for s in doc.sections if s.id == "s5_execucao_obra")
    frases = [b.text for b in sec5.blocks if b.kind == "paragraph" and b.text
             and "não houve execução" in b.text]
    if any("30/08/2026" in f for f in frases):
        print("B3  OK   frase de dia sem execucao presente: %r" % frases)
    else:
        falhas.append("B3: frase de dia sem execucao ausente ou sem a data esperada: %s" % frases)

    # ---------------------------------------------------------------- B4 (Q-8)
    tabela_clima = next(b for b in sec5.blocks if b.kind == "table")
    linha_clima = tabela_clima.rows[2]  # ["CLIMA", ...7 valores]
    idx_3008 = 1  # 2o dia do periodo (indice 0-based na linha, apos o rotulo "CLIMA")
    if linha_clima[0] == "CLIMA" and linha_clima[idx_3008 + 1] == "-":
        print("B4  OK   30/08 (sem ficha) renderiza '-' na tabela climatica")
    else:
        falhas.append("B4: linha CLIMA == %s, esperado '-' na posicao do dia 30/08" % linha_clima)
    if tabela_clima.status == "HITL" and tabela_clima.note:
        print("B4b OK   bloco da tabela climatica esta HITL, com nota: %r" % tabela_clima.note)
    else:
        falhas.append("B4b: tabela climatica status=%s note=%r, esperado HITL com nota"
                      % (tabela_clima.status, tabela_clima.note))
    # ---------------------------------------------------------------- B5 (Q-10)
    sec9 = next(s for s in doc.sections if s.id == "s9_interferencias")
    paragrafo9 = next(b for b in sec9.blocks if b.kind == "paragraph")
    if paragrafo9.status == "HITL" and "riscos ou interferências" in (paragrafo9.text or ""):
        print("B5  OK   secao 9 sem ocorrencia gera frase negativa, status HITL, nao bloqueia")
    else:
        falhas.append("B5: bloco da secao 9 == %s" % paragrafo9)
    if len(doc.sections) == 14:
        print("B5b OK   montagem nao bloqueada pela secao 9 HITL -- 14 secoes no documento")
    else:
        falhas.append("B5b: len(doc.sections) == %d, esperado 14" % len(doc.sections))

    # ---------------------------------------------------------------- B6
    manual = {"READY": 0, "PLACEHOLDER": 0, "HITL": 0, "LIMITATION": 0}
    for s in doc.sections:
        for b in s.blocks:
            manual[b.status] += 1
    if manual == doc.status_counts():
        print("B6  OK   status_counts() bate com contagem manual: %s" % manual)
    else:
        falhas.append("B6: status_counts()=%s, manual=%s" % (doc.status_counts(), manual))

    # ---------------------------------------------------------------- B7
    fonte = inspect.getsource(CB)
    if 'obra ==' not in fonte and 'obra=="' not in fonte.replace(" ", ""):
        print("B7  OK   content_builder.py nao contem `if obra == ...` (checagem estrutural)")
    else:
        falhas.append("B7: content_builder.py parece condicionar por nome de obra literal")

    rodadas = []
    for obra in U.OBRAS:
        rd = _carregar_report_data(obra)
        if rd is None:
            continue
        try:
            d = CB.construir_documento(rd)
            assert len(d.sections) == 14
            rodadas.append(obra)
        except Exception as e:
            falhas.append("B7b: %s levantou %s: %s" % (obra, type(e).__name__, e))
    if rodadas and not any(f.startswith("B7b") for f in falhas):
        print("B7b OK   construir_documento roda sem excecao para: %s (14 secoes cada)" % rodadas)
    elif not rodadas:
        print("B7b SKIP nenhum outro report_data.json disponivel em output/shadow/ nesta sessao")

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
