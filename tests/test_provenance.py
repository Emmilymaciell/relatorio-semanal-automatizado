# -*- coding: utf-8 -*-
"""
test_provenance.py

Testa scripts/gerador/provenance.py contra a saida real de canonical.construir()
para RD103, 2026-08-29..2026-09-04 (spec Sec.3: "Todo valor de progress.*,
climate.*, nonconformities.*, report_number e services precisa ter entrada de
proveniencia. Campo derivado usa method='derived' e lista inputs").

  P1  'report_number' tem entrada de proveniencia
  P2  'progress.global_concluido' e 'progress.global_concluido_atacado' tem
      entrada, com source_file/sheet/cell apontando para o arquivo real
  P3  todo dia de climate.by_day COM clima lido tem entrada
      'climate.by_day.<data>' com source_file+sheet+cell
  P4  'nonconformities.issued_in_period' e 'nonconformities.history_by_status'
      tem entrada
  P5  'services' tem entrada, method == 'derived', com `inputs`
  P6  toda entrada tem os campos minimos: value, method, confidence, extracted_at

Execute:  python tests/test_provenance.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import sys
import datetime as dt

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
from gerador.canonical import construir  # noqa: E402

OBRA = "RD103"
DI, DF = dt.date(2026, 8, 29), dt.date(2026, 9, 4)


def main():
    falhas = []
    manifest, data, prov = construir(OBRA, DI, DF)
    d = prov.as_dict()

    # ---------------------------------------------------------------- P1
    if "report_number" in d:
        print("P1  OK   'report_number' tem entrada de proveniencia")
    else:
        falhas.append("P1: 'report_number' sem entrada de proveniencia")

    # ---------------------------------------------------------------- P2
    for chave in ("progress.global_concluido", "progress.global_concluido_atacado"):
        e = d.get(chave)
        if e and e.get("source_file") and e.get("sheet") and e.get("cell"):
            print("P2  OK   '%s' -> %s!%s (%s)" % (chave, e["sheet"], e["cell"],
                                                    os.path.basename(e["source_file"])))
        else:
            falhas.append("P2: '%s' sem entrada completa (source_file/sheet/cell): %s" % (chave, e))

    # ---------------------------------------------------------------- P3
    n_com_clima = 0
    for dia in data["climate"]["by_day"]:
        chave = "climate.by_day.%s" % dia["date"]
        e = d.get(chave)
        if not e:
            falhas.append("P3: sem entrada de proveniencia para '%s'" % chave)
            continue
        if dia["clima"] is not None:
            n_com_clima += 1
            if not (e.get("source_file") and e.get("sheet") and e.get("cell")):
                falhas.append("P3: '%s' com clima lido mas proveniencia incompleta: %s" % (chave, e))
    if n_com_clima == 6 and not any(f.startswith("P3") for f in falhas):
        print("P3  OK   6/6 dias com clima lido tem proveniencia completa "
             "(source_file+sheet+cell); 7/7 dias tem entrada (inclusive o ausente)")
    elif not any(f.startswith("P3") for f in falhas):
        falhas.append("P3: esperado 6 dias com clima lido, achou %d" % n_com_clima)

    # ---------------------------------------------------------------- P4
    for chave in ("nonconformities.issued_in_period", "nonconformities.history_by_status"):
        if chave in d:
            print("P4  OK   '%s' tem entrada de proveniencia" % chave)
        else:
            falhas.append("P4: '%s' sem entrada de proveniencia" % chave)

    # ---------------------------------------------------------------- P5
    e = d.get("services")
    if e and e.get("method") == "derived" and e.get("inputs"):
        print("P5  OK   'services' -> method=derived, inputs=%s" % e["inputs"])
    else:
        falhas.append("P5: 'services' sem entrada 'derived' com inputs: %s" % e)

    # ---------------------------------------------------------------- P6
    campos_min = {"value", "method", "confidence", "extracted_at"}
    incompletas = [k for k, v in d.items() if not campos_min.issubset(v)]
    if not incompletas:
        print("P6  OK   todas as %d entradas tem value/method/confidence/extracted_at" % len(d))
    else:
        falhas.append("P6: entradas incompletas: %s" % incompletas)

    print("=" * 74)
    if falhas:
        print("FALHAS (%d):" % len(falhas))
        for f in falhas:
            print("  X", f)
        print("RESULTADO: FALHOU")
        return 1
    print("RESULTADO: OK (%d chaves de proveniencia)" % len(d))
    return 0


if __name__ == "__main__":
    sys.exit(main())
