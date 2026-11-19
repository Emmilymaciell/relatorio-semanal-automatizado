# -*- coding: utf-8 -*-
r"""
scripts/gerador/extratores/ambiental.py - ciclo ambiental quinzenal + fichas
FA-* da pasta datada do periodo.

VERIFICADO DIRETAMENTE nesta sessao (2026-09-17): RD103, pasta
'6. Acompanhamento ambiental e SST\2026-09-02\' existe, com subpastas
'Ficha de acompanhamento ambiental\FA-016.1 - ...pdf' e
'Ficha de acompanhamento SST\FA-015.1 - ...pdf'.
(2026-08-29 - 2026-06-06).days == 84; 84 %% 14 == 0 -> ciclo aplica.

A leitura de KPIs do dashboard ('04_Acompanhamento - Licenca Ambiental.xlsx',
figura_indices / figura_condicionantes / figura_gestao) NAO e feita aqui: a
reconstrucao correta do painel do Power BI para uma data-base arbitraria e uma
cadeia de calculo grande e ja mapeada em scripts/12_analisar_ambiental.py, fora
do orcamento desta fatia. Fica MVP_LIMITATION -- registrado, nao inventado.
"""
import os
import re
import sys
import datetime as dt

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(os.path.dirname(AQUI))
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402

ANCORA_QUINZENAL = dt.date(2026, 6, 6)
CICLO_DIAS = 14


def ciclo_aplica(data_inicio):
    if isinstance(data_inicio, str):
        data_inicio = dt.date.fromisoformat(data_inicio)
    return (data_inicio - ANCORA_QUINZENAL).days % CICLO_DIAS == 0


def localizar_pasta_visita(pasta_ambiental_sst, data_inicio, data_fim):
    if not pasta_ambiental_sst or not U.exists(pasta_ambiental_sst):
        return None
    if isinstance(data_inicio, str):
        data_inicio = dt.date.fromisoformat(data_inicio)
    if isinstance(data_fim, str):
        data_fim = dt.date.fromisoformat(data_fim)
    for nome in U.listdir(pasta_ambiental_sst):
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", nome):
            continue
        d = dt.date.fromisoformat(nome)
        if data_inicio <= d <= data_fim:
            return os.path.join(pasta_ambiental_sst, nome)
    return None


def ler_fichas_da_visita(pasta_visita):
    if not pasta_visita or not U.exists(pasta_visita):
        return {"status": "NOT_APPLICABLE"}
    fa_ambiental = fa_sst = None
    for nome in U.listdir(pasta_visita):
        sub = os.path.join(pasta_visita, nome)
        if not U.isdir(sub):
            continue
        n = U.norm(nome)
        pdfs = [f for f in U.listdir(sub) if f.lower().endswith(".pdf")]
        if "ambiental" in n and pdfs:
            fa_ambiental = os.path.join(sub, pdfs[0])
        elif "sst" in n and pdfs:
            fa_sst = os.path.join(sub, pdfs[0])
    return {
        "status": "OK", "visita_data": os.path.basename(pasta_visita),
        "fa_ambiental": fa_ambiental, "fa_sst": fa_sst,
    }
