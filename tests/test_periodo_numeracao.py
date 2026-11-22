# -*- coding: utf-8 -*-
"""
test_periodo_numeracao.py

Testa scripts/gerador/periodo.py:
  T1  numero_relatorio('RD103', 2026-08-29) == 49 (caso de teste da spec,
      conferido diretamente pela pasta-semana real no SharePoint nesta sessao)
  T2  a regra (data_inicio - ancora)/7 + 1 bate com o numero real observado no
      NOME DA PASTA-SEMANA, para >= 20 semanas do regime vigente (>= 2026-01-03)
      em CADA uma das 4 obras (output/inventarios/relatorios_semanais_historico.json)
  T3  schema_version(2026-08-29) == 'v3' (CURRENT)

Execute:  python tests/test_periodo_numeracao.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import sys
import json
import datetime as dt

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from gerador import periodo as P  # noqa: E402

HIST = os.path.join(U.OUT, "inventarios", "relatorios_semanais_historico.json")
INICIO_REGIME_2 = dt.date(2026, 1, 3)


def main():
    falhas = []

    # ---------------------------------------------------------------- T1
    n = P.numero_relatorio("RD103", dt.date(2026, 8, 29))
    if n == 49:
        print("T1  OK   RD103/2026-08-29 -> numero %d (esperado 49)" % n)
    else:
        falhas.append("T1: RD103/2026-08-29 deu %s, esperado 49" % n)

    # ---------------------------------------------------------------- T2
    hist = json.load(open(HIST, encoding="utf-8"))["relatorios"]
    for obra in U.OBRAS:
        entradas = [r for r in hist if r.get("obra") == obra]
        ok, tot, checados = 0, 0, 0
        for r in entradas:
            try:
                da = dt.date.fromisoformat(r["pasta_data_inicio"])
                db = dt.date.fromisoformat(r["pasta_data_fim"])
            except (KeyError, ValueError, TypeError):
                continue
            di = min(da, db)  # 3 pastas tem as datas invertidas no nome (documentado)
            if di < INICIO_REGIME_2:
                continue
            numero_real = r.get("numero")
            if numero_real is None:
                continue
            checados += 1
            calc = P.numero_relatorio(obra, di)
            tot += 1
            if calc == numero_real:
                ok += 1
            else:
                falhas.append("T2 %s pasta=%r: calculado=%s real=%s"
                              % (obra, r.get("pasta"), calc, numero_real))
        print("T2  %-8s %3d/%-3d semanas do regime vigente batem com a regra "
             "(numero = (inicio-ancora)/7+1)" % (obra, ok, tot))
        if tot < 20:
            falhas.append("T2 %s: so %d semanas do regime vigente disponiveis "
                          "para testar (esperado >= 20)" % (obra, tot))

    # ---------------------------------------------------------------- T3
    v = P.schema_version(dt.date(2026, 8, 29))
    if v == "v3":
        print("T3  OK   schema_version(2026-08-29) == 'v3' (CURRENT)")
    else:
        falhas.append("T3: schema_version(2026-08-29) == %r, esperado 'v3'" % v)

    print("=" * 74)
    if falhas:
        print("FALHAS (%d):" % len(falhas))
        for f in falhas[:30]:
            print("  X", f)
        print("RESULTADO: FALHOU")
        return 1
    print("RESULTADO: OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
