# -*- coding: utf-8 -*-
"""
test_regras_anexos.py

Valida as regras de composicao dos ANEXOS descobertas no historico
(saida de scripts/08_historico_anexos.py, 14 semanas x 4 obras).

  A1  os dois ultimos anexos sao sempre Diagrama Unifilar e Diario de Obras,
      nessa ordem
  A2  o bloco ambiental (Acompanhamento ambiental + SST) e QUINZENAL:
      aparece <=> (data_inicio - 2026-06-06) % 14 == 0
  A3  o anexo 'Desvio de cronograma acumulado' existiu ate 2026-07-25 e
      desapareceu a partir de 2026-08-01, junto com a secao 6 do relatorio
  A4  Acompanhamento ambiental e Acompanhamento SST andam SEMPRE juntos

Execute:  python tests/test_regras_anexos.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import sys
import json
import datetime

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(AQUI), "scripts"))
import prog_util as U

DADOS = os.path.join(U.OUT, "comparacoes", "historico_anexos.json")

ANCORA_QUINZENAL = datetime.date(2026, 6, 6)
FIM_DESVIO_CRONOGRAMA = datetime.date(2026, 7, 25)   # ultima semana com a secao 6


def main():
    if not os.path.exists(DADOS):
        print("Rode antes: python scripts/08_historico_anexos.py --ultimas 14")
        return 2
    d = json.load(open(DADOS, encoding="utf-8"))
    falhas = []
    cont = {"A1": [0, 0], "A2": [0, 0], "A3": [0, 0], "A4": [0, 0]}

    for obra, linhas in d["obras"].items():
        for l in linhas:
            if "categorias" not in l:
                continue
            di = datetime.date.fromisoformat(l["data_inicio"])
            cats = l["categorias"]
            rot = "%s n=%s %s" % (obra, l["numero"], di)

            # A1 — o Diario e SEMPRE o ultimo; o Unifilar vem imediatamente antes
            # no modelo vigente, e antes do Desvio de Cronograma no modelo antigo.
            cont["A1"][1] += 1
            esperado_cauda = (["DIAGRAMA_UNIFILAR", "DIARIO_DE_OBRAS"]
                              if di > FIM_DESVIO_CRONOGRAMA else
                              ["DIAGRAMA_UNIFILAR", "DESVIO_CRONOGRAMA", "DIARIO_DE_OBRAS"])
            if cats[-len(esperado_cauda):] == esperado_cauda:
                cont["A1"][0] += 1
            else:
                falhas.append("A1 %s: cauda = %s, esperada %s"
                              % (rot, cats[-len(esperado_cauda):], esperado_cauda))

            # A2
            cont["A2"][1] += 1
            esperado = ((di - ANCORA_QUINZENAL).days % 14) == 0
            tem = "ACOMP_AMBIENTAL" in cats and "ACOMP_SST" in cats
            if esperado == tem:
                cont["A2"][0] += 1
            else:
                falhas.append("A2 %s: esperado=%s obtido=%s" % (rot, esperado, tem))

            # A3
            cont["A3"][1] += 1
            tem_desvio = "DESVIO_CRONOGRAMA" in cats
            if tem_desvio == (di <= FIM_DESVIO_CRONOGRAMA):
                cont["A3"][0] += 1
            else:
                falhas.append("A3 %s: desvio_cronograma=%s" % (rot, tem_desvio))

            # A4
            cont["A4"][1] += 1
            if ("ACOMP_AMBIENTAL" in cats) == ("ACOMP_SST" in cats):
                cont["A4"][0] += 1
            else:
                falhas.append("A4 %s: ambiental e SST desacoplados: %s" % (rot, cats))

    print("=" * 74)
    print("VALIDACAO DAS REGRAS DE ANEXOS")
    print("=" * 74)
    desc = {
        "A1": "Diario sempre por ultimo; Unifilar imediatamente antes (modelo vigente)",
        "A2": "bloco ambiental quinzenal, ancora 2026-06-06",
        "A3": "Desvio de cronograma ate 2026-07-25, nunca depois",
        "A4": "Acompanhamento ambiental e SST sempre juntos",
    }
    for k, (ok, tot) in cont.items():
        print("  %s  %3d/%-3d  %s" % (k, ok, tot, desc[k]))

    if falhas:
        print("\nFALHAS (%d):" % len(falhas))
        for f in falhas[:20]:
            print("  X", f)
        print("\nRESULTADO: FALHOU")
        return 1
    print("\nRESULTADO: TODAS AS REGRAS DE ANEXO VALIDADAS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
