# -*- coding: utf-8 -*-
r"""
scripts/gerador/extratores/avanco.py - avanco global da obra.

VERIFICADO DIRETAMENTE nesta sessao (2026-09-17) para RD103/2026-08-29..09-04:
  '[Programa] Avanço Global de Obra.xlsx' (copia da pasta da semana),
  aba 'Avanço Global - RD103', C3=0.32233518354079294, C7=0.421927187037607.
  'Gráfico - Avanço Global.xlsx', aba 'Relatório Semanal', ultima linha
  (2026-09-04): 32.23 / 42.19 -- bate com C3*100/C7*100. Aba 'Acumulado' tem a
  mesma ultima linha, serie historica completa desde 2026-03-13.
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


def _abrir(caminho, subdir, nome):
    assert_nao_e_golden(caminho)
    if not U.exists(caminho):
        return None, {"status": "NOT_FOUND", "path": caminho}
    try:
        local = U.copiar_para_tmp(caminho, subdir, nome)
        wb = openpyxl.load_workbook(local, data_only=True)
        return wb, None
    except OSError as e:
        return None, {"status": "UNREADABLE", "path": caminho, "erro": str(e)}


def ler_prog(caminho, road_key):
    """Le C3 (global_concluido) e C7 (global_concluido_atacado) da aba
    'Avanço Global - {road_key}' do workbook '[Programa] Avanço Global de Obra.xlsx'
    (a copia da PASTA DA SEMANA, nao o mestre compartilhado)."""
    if not caminho:
        return {"status": "NOT_FOUND"}
    wb, erro = _abrir(caminho, "avanco", "prog_avanco%s" % os.path.splitext(caminho)[1])
    if erro:
        return erro

    alvo = U.norm("Avanço Global - %s" % road_key)
    aba_real = next((sn for sn in wb.sheetnames if U.norm(sn) == alvo), None)
    if not aba_real:
        return {"status": "NOT_FOUND", "path": caminho,
                "erro": "aba 'Avanço Global - %s' ausente (abas: %s)" % (road_key, wb.sheetnames)}

    ws = wb[aba_real]
    c3, c7 = ws["C3"].value, ws["C7"].value
    if c3 is None or c7 is None:
        return {"status": "UNREADABLE", "path": caminho, "aba": aba_real,
                "erro": "C3/C7 sem valor calculado em cache (arquivo .xlsx sem "
                        "recalculo salvo?) -- ver T-12"}
    return {"status": "OK", "path": caminho, "aba": aba_real, "c3": c3, "c7": c7,
            "c3_pct": round(c3 * 100, 4), "c7_pct": round(c7 * 100, 4)}


def ler_grafico(caminho):
    """Le as series 'Relatório Semanal' e 'Acumulado' de
    'Gráfico - Avanço Global.xlsx'."""
    if not caminho:
        return {"status": "NOT_FOUND"}
    wb, erro = _abrir(caminho, "avanco", "grafico%s" % os.path.splitext(caminho)[1])
    if erro:
        return erro

    series = {}
    for aba_alvo, chave in (("Relatório Semanal", "relatorio_semanal"),
                             ("Acumulado", "acumulado")):
        nome_real = next((sn for sn in wb.sheetnames if U.norm(sn) == U.norm(aba_alvo)), None)
        if not nome_real:
            series[chave] = {"status": "NOT_FOUND"}
            continue
        ws = wb[nome_real]
        linhas = []
        for r in range(2, (ws.max_row or 1) + 1):
            data = ws.cell(r, 1).value
            concl = ws.cell(r, 2).value
            atac = ws.cell(r, 3).value
            if data is None:
                continue
            if isinstance(data, dt.datetime):
                data = data.date()
            if not isinstance(data, dt.date):
                # linha de cabecalho repetida dentro do bloco de dados (a aba
                # 'Relatorio Semanal' traz o cabecalho na propria area lida).
                continue
            linhas.append({
                "data": data.isoformat() if isinstance(data, dt.date) else str(data),
                "concluido_pct": concl, "concluido_atacado_pct": atac,
            })
        series[chave] = {"status": "OK", "aba": nome_real, "n": len(linhas), "linhas": linhas}

    return {"status": "OK", "path": caminho, "series": series}
