# -*- coding: utf-8 -*-
"""
test_canonical_rd103.py

Testa scripts/gerador/canonical.py para RD103, 2026-08-29..2026-09-04:
  C1  report_data.json tem as 18 chaves de topo (config/report_schema.yaml +
      spec Sec.3)
  C2  climate.by_day tem 7 entradas (uma por dia do periodo)
  C3  services nao vazio e cada item normalizado pela taxonomia (tem
      canonical_name e a key existe em services_taxonomy.servicos)
  C4  closing_term is None (placeholder explicito desta fase)
  C5  nunca um valor NUMERICO inventado: todo dia de climate.by_day sem
      fonte tem clima=None + _status != 'OK' (nunca um clima "chutado")
  C6  report_number bate com o do manifest (mesma fonte, sem duplicidade de logica)
  C7  Q-8 (spec Sec.1): 2026-08-30 (sem ficha) ganha clima_label=INDETERMINADO,
      validation_status=HITL e motivo -- nunca "Bom" nem repeticao do dia anterior;
      todo dia com clima=None tem os 3 campos, todo dia com clima!=None tem
      validation_status=OK
  C8  Q-9 (spec Sec.1): cada item de services[] expoe `sigla` (do
      services_taxonomy.yaml, nunca heuristica); para RD103/49 sabe-se que
      'Obras de Arte Especiais' tem sigla OAE

Execute:  python tests/test_canonical_rd103.py
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
import prog_util as U  # noqa: E402
from gerador import config as C  # noqa: E402
from gerador.canonical import construir  # noqa: E402

OBRA = "RD103"
DI, DF = dt.date(2026, 8, 29), dt.date(2026, 9, 4)

CHAVES_ESPERADAS = [
    "metadata", "road", "contract", "period", "report_number", "services",
    "climate", "progress", "field_records", "photos", "nonconformities",
    "environment", "sst", "meetings", "risks", "interferences", "annexes",
    "closing_term",
]


def main():
    falhas = []
    manifest, data, prov = construir(OBRA, DI, DF)

    # ---------------------------------------------------------------- C1
    chaves = list(data.keys())
    if len(chaves) == 18 and set(chaves) == set(CHAVES_ESPERADAS):
        print("C1  OK   18 chaves de topo, todas as esperadas presentes")
    else:
        falhas.append("C1: chaves == %s (%d), esperadas %s (%d)"
                      % (sorted(chaves), len(chaves), sorted(CHAVES_ESPERADAS), len(CHAVES_ESPERADAS)))

    # ---------------------------------------------------------------- C2
    by_day = data["climate"]["by_day"]
    if len(by_day) == 7:
        print("C2  OK   climate.by_day tem 7 entradas")
    else:
        falhas.append("C2: climate.by_day tem %d entradas, esperado 7" % len(by_day))
    datas_by_day = [d["date"] for d in by_day]
    datas_esperadas = [(DI + dt.timedelta(days=i)).isoformat() for i in range(7)]
    if datas_by_day != datas_esperadas:
        falhas.append("C2b: datas de climate.by_day == %s, esperado %s" % (datas_by_day, datas_esperadas))

    # ---------------------------------------------------------------- C3
    tax = C.services_taxonomy()
    chaves_taxonomia = set(tax.get("servicos", {}))
    if data["services"]:
        problemas = [s for s in data["services"]
                    if s.get("key") not in chaves_taxonomia or not s.get("canonical_name")]
        if not problemas:
            print("C3  OK   services nao vazio (%d), todos normalizados pela taxonomia: %s"
                 % (len(data["services"]), [s["canonical_name"] for s in data["services"]]))
        else:
            falhas.append("C3: services com problema de normalizacao: %s" % problemas)
    else:
        falhas.append("C3: services esta vazio (esperado nao-vazio para RD103 nesta semana)")

    # ---------------------------------------------------------------- C4
    if data["closing_term"] is None:
        print("C4  OK   closing_term is None")
    else:
        falhas.append("C4: closing_term == %r, esperado None" % data["closing_term"])

    # ---------------------------------------------------------------- C5
    problema_c5 = []
    for d in by_day:
        if d["clima"] is None and d.get("_status") == "OK":
            problema_c5.append(d)
        if d["clima"] is not None and d.get("_status") != "OK":
            problema_c5.append(d)
    if not problema_c5:
        print("C5  OK   clima e _status sempre coerentes (nunca clima=None com status=OK "
             "nem clima preenchido com status != OK)")
    else:
        falhas.append("C5: entradas incoerentes em climate.by_day: %s" % problema_c5)
    # dia sem ficha (2026-08-30) tem que ser um desses None coerentes
    dia_30 = next(d for d in by_day if d["date"] == "2026-08-30")
    if dia_30["clima"] is None and dia_30["_status"] == "NOT_FOUND":
        print("C5b OK   2026-08-30 (sem ficha) -> clima=None, _status=NOT_FOUND")
    else:
        falhas.append("C5b: 2026-08-30 == %s" % dia_30)

    # ---------------------------------------------------------------- C6
    if data["report_number"] == manifest["report"]["report_number"] == 49:
        print("C6  OK   report_number == 49 em report_data e manifest")
    else:
        falhas.append("C6: report_data=%s manifest=%s" % (data["report_number"],
                                                           manifest["report"]["report_number"]))

    # ---------------------------------------------------------------- C7 (Q-8)
    dia_30 = next(d for d in by_day if d["date"] == "2026-08-30")
    if dia_30.get("clima_label") == "INDETERMINADO" and dia_30.get("validation_status") == "HITL" \
       and dia_30.get("motivo"):
        print("C7  OK   2026-08-30 -> clima_label=INDETERMINADO, validation_status=HITL, "
             "motivo=%r" % dia_30["motivo"])
    else:
        falhas.append("C7: 2026-08-30 sem marcacao Q-8 esperada: %s" % dia_30)
    problema_c7b = []
    for d in by_day:
        if d["clima"] is None and not (d.get("clima_label") == "INDETERMINADO"
                                       and d.get("validation_status") == "HITL" and d.get("motivo")):
            problema_c7b.append(d)
        if d["clima"] is not None and d.get("validation_status") != "OK":
            problema_c7b.append(d)
    if not problema_c7b:
        print("C7b OK   todo dia com clima=None tem clima_label/validation_status/motivo; "
             "todo dia com clima lido tem validation_status=OK")
    else:
        falhas.append("C7b: dias incoerentes com a regra Q-8: %s" % problema_c7b)

    # ---------------------------------------------------------------- C8 (Q-9)
    if all("sigla" in s for s in data["services"]):
        oae = next((s for s in data["services"] if s["key"] == "OBRAS_DE_ARTE_ESPECIAIS"), None)
        if oae and oae["sigla"] == "OAE":
            print("C8  OK   todo item de services[] expoe 'sigla'; OBRAS_DE_ARTE_ESPECIAIS -> OAE")
        else:
            falhas.append("C8: OBRAS_DE_ARTE_ESPECIAIS sem sigla=OAE: %s" % oae)
    else:
        falhas.append("C8: nem todo item de services[] tem a chave 'sigla': %s" % data["services"])

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
