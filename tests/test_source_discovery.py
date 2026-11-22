# -*- coding: utf-8 -*-
"""
test_source_discovery.py

Testa scripts/gerador/source_discovery.py para RD103, 2026-08-29..2026-09-04
(caso de teste da spec Sec.1/6.2). Todos os fatos abaixo foram CONFIRMADOS por
leitura direta do SharePoint nesta sessao antes de escrever o codigo:

  S1  6 fichas de campo encontradas (falta so 2026-08-30)
  S2  '2026-08-30' em field_records_missing_days
  S3  5 auxiliares encontrados (accumulated_sheet, progress_chart, prog_progress,
      photo_workbook, progress_unifilar)
  S4  diario 'Diário de obra digital assinado\\2026-08-29 a 2026-09-04 - RD-103.pdf' localizado
  S5  nenhum caminho no manifest contem 'obsoleto' (normalizado)
  S6  o golden (docx/pdf) NAO esta em `sources` -- so em `golden_reference`,
      e la contem so path/sha1/n_paginas (metadados, nao conteudo)
  S7  golden_reference.pdf.n_paginas == 35 (contado por fitz, sem ler texto)
  S8  report_number == 49 (independente do numero embutido no nome da pasta)

Execute:  python tests/test_source_discovery.py
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
from gerador.source_discovery import construir_manifest  # noqa: E402

OBRA = "RD103"
DI, DF = dt.date(2026, 8, 29), dt.date(2026, 9, 4)


def _flatten_paths(obj, acc):
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k == "path" and isinstance(v, str):
                acc.append(v)
            else:
                _flatten_paths(v, acc)
    elif isinstance(obj, list):
        for v in obj:
            _flatten_paths(v, acc)


def main():
    falhas = []
    m = construir_manifest(OBRA, DI, DF)

    # ---------------------------------------------------------------- S1/S2
    fr = m["sources"]["field_records"]
    ok_days = sum(1 for f in fr if f["status"] == "OK")
    if ok_days == 6:
        print("S1  OK   6 fichas de campo encontradas")
    else:
        falhas.append("S1: %d fichas encontradas, esperado 6" % ok_days)

    if m["sources"]["field_records_missing_days"] == ["2026-08-30"]:
        print("S2  OK   field_records_missing_days == ['2026-08-30']")
    else:
        falhas.append("S2: field_records_missing_days == %s" % m["sources"]["field_records_missing_days"])

    # ---------------------------------------------------------------- S3
    auxiliares = ["accumulated_sheet", "progress_chart", "prog_progress",
                 "photo_workbook", "progress_unifilar"]
    achados = [a for a in auxiliares if m["sources"][a].get("status") == "OK"]
    if len(achados) == 5:
        print("S3  OK   5/5 auxiliares encontrados: %s" % achados)
    else:
        falhas.append("S3: so %d/5 auxiliares OK (%s); manifest=%s"
                      % (len(achados), achados,
                         {a: m["sources"][a].get("status") for a in auxiliares}))

    # ---------------------------------------------------------------- S4
    diario = m["sources"]["daily_reports"]
    nome_esperado = "2026-08-29 a 2026-09-04 - RD-103.pdf"
    if diario.get("status") == "OK" and diario["path"].endswith(nome_esperado):
        print("S4  OK   diario localizado: ...%s" % nome_esperado)
    else:
        falhas.append("S4: daily_reports == %s" % diario)

    # ---------------------------------------------------------------- S5
    caminhos = []
    _flatten_paths(m["sources"], caminhos)
    com_obsoleto = [c for c in caminhos if "obsoleto" in U.norm(c)]
    if not com_obsoleto:
        print("S5  OK   nenhum caminho contem 'obsoleto' (%d caminhos verificados)" % len(caminhos))
    else:
        falhas.append("S5: caminhos com 'obsoleto': %s" % com_obsoleto)

    # ---------------------------------------------------------------- S6/S7
    if "golden" not in json.dumps(m["sources"], ensure_ascii=False).lower():
        # nenhuma chave/valor de `sources` deveria conter a palavra "golden";
        # ele so existe em `golden_reference`, a parte separada do manifest.
        print("S6  OK   golden fora de `sources` (so em `golden_reference`)")
    else:
        falhas.append("S6: 'golden' aparece dentro de `sources` -- vazamento de referencia")
    docx_path = m["golden_reference"]["docx"]["path"] or ""
    pdf_path = m["golden_reference"]["pdf"]["path"] or ""
    if not any(docx_path in c or pdf_path in c for c in caminhos if pdf_path or docx_path):
        print("S6b OK   caminho do golden nao duplicado dentro de `sources`")
    else:
        falhas.append("S6b: caminho do golden tambem aparece em `sources`")

    if m["golden_reference"]["pdf"]["n_paginas"] == 35:
        print("S7  OK   golden_reference.pdf.n_paginas == 35")
    else:
        falhas.append("S7: n_paginas == %s, esperado 35"
                      % m["golden_reference"]["pdf"]["n_paginas"])

    # ---------------------------------------------------------------- S8
    if m["report"]["report_number"] == 49:
        print("S8  OK   report_number == 49")
    else:
        falhas.append("S8: report_number == %s" % m["report"]["report_number"])

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
