# -*- coding: utf-8 -*-
r"""
scripts/gerador/provenance.py - coletor de proveniencia.

Produz `report_provenance.json`: paralelo ao report_data.json, mesma estrutura
de chaves em caminho pontilhado (ex.: "progress.global_concluido"), cada uma
com {value, source_file, sheet, cell, method, confidence, extracted_at} (campo
derivado usa method="derived" e uma lista `inputs`).

"Nunca inventar numero, data, estaca, NC ou percentual": todo valor que entra
em report_data.json com origem em planilha/arquivo DEVE ter uma entrada aqui.
"""
import os
import sys
import json
import datetime as dt

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(AQUI)
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402


def _jsonable(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.isoformat()
    return v


def agora():
    return dt.datetime.now().astimezone().isoformat(timespec="seconds")


class Provenance:
    """Coletor de proveniencia. Uma instancia por execucao do `canonical.py`."""

    def __init__(self):
        self._entradas = {}

    def registrar(self, chave, value, source_file=None, sheet=None, cell=None,
                   method="direct extraction",
                   confidence="CONFIRMADO POR EVIDÊNCIA DIRETA",
                   inputs=None, **extra):
        entrada = {
            "value": _jsonable(value),
            "method": method,
            "confidence": confidence,
            "extracted_at": agora(),
        }
        if source_file is not None:
            entrada["source_file"] = source_file
        if sheet is not None:
            entrada["sheet"] = sheet
        if cell is not None:
            entrada["cell"] = cell
        if inputs is not None:
            entrada["inputs"] = inputs
        for k, v in extra.items():
            entrada[k] = _jsonable(v)
        self._entradas[chave] = entrada
        return entrada

    def derivado(self, chave, value, inputs,
                 confidence="ALTA CONFIANÇA (derivado deterministicamente)", **extra):
        return self.registrar(chave, value, method="derived",
                               confidence=confidence, inputs=inputs, **extra)

    def tem(self, chave):
        return chave in self._entradas

    def get(self, chave):
        return self._entradas.get(chave)

    def chaves(self):
        return list(self._entradas)

    def as_dict(self):
        return dict(self._entradas)

    def salvar(self, path):
        U.assert_read_only(path)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(self.as_dict(), fh, ensure_ascii=False, indent=2, default=str)
        return path
