# -*- coding: utf-8 -*-
r"""
scripts/gerador/extratores/ficha_acumulada.py - Ficha de Campo Acumulada.

Nenhum campo de config/report_schema.yaml (Slice 2) consome este arquivo
diretamente nesta iteracao -- ele alimenta a Figura 1 da secao "FICHA DE
ACOMPANHAMENTO" (classe G, imagem), fora do escopo desta fase (ver spec Sec.0
item 6: "fora de escopo: imagens"). Por isso a leitura aqui e deliberadamente
leve: so metadados + uma tentativa best-effort de achar a "data de corte", sem
inventar. Registrado no manifest (`sources.accumulated_sheet`) por completude
de descoberta (arquitetura da spec Sec.3).
"""
import os
import sys
import datetime as dt

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(os.path.dirname(AQUI))
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from .. import assert_nao_e_golden  # noqa: E402


def ler(caminho):
    if not caminho:
        return {"status": "NOT_FOUND"}
    assert_nao_e_golden(caminho)
    if not U.exists(caminho):
        return {"status": "NOT_FOUND", "path": caminho}

    try:
        destino = "acumulada_%s%s" % (U.slug(os.path.basename(caminho)),
                                       os.path.splitext(caminho)[1])
        local = U.copiar_para_tmp(caminho, "ficha_acumulada", destino)
        wb = openpyxl.load_workbook(local, data_only=True, read_only=True)
    except OSError as e:
        return {"status": "UNREADABLE", "path": caminho, "erro": str(e)}

    aba = wb.sheetnames[0]
    ws = wb[aba]
    data_corte = None
    celula_data_corte = None
    try:
        max_r = min(ws.max_row or 15, 15)
        max_c = min(ws.max_column or 40, 40)
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and "corte" in U.norm(v):
                    viz = ws.cell(r, c + 1).value
                    if isinstance(viz, dt.datetime):
                        data_corte = viz.date().isoformat()
                        celula_data_corte = "%d,%d" % (r, c + 1)
                    elif isinstance(viz, dt.date):
                        data_corte = viz.isoformat()
                        celula_data_corte = "%d,%d" % (r, c + 1)
    except Exception:
        pass

    return {
        "status": "OK", "path": caminho, "aba": aba,
        "n_linhas_aprox": ws.max_row, "n_colunas_aprox": ws.max_column,
        "data_corte": data_corte, "data_corte_celula": celula_data_corte,
        "data_corte_status": "OK" if data_corte else "NOT_FOUND",
        "nota": "leitura leve (metadados); tabela completa de estacas fora de escopo nesta fase",
    }
