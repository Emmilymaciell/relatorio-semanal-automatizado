# -*- coding: utf-8 -*-
r"""
scripts/gerador/periodo.py - periodo do relatorio, numeracao e rotulos de dia.

Regras (ja provadas em tests/test_regras_periodo.py sobre as 193 semanas do
regime vigente, CONFIRMADO POR EVIDENCIA DIRETA):
  - regime sabado -> sexta, 7 dias, a partir de 2026-01-03
  - numero = (data_inicio - ancora_da_obra).days // 7 + 1
  - ancora vem de config/road_registry.yaml:obras.<OBRA>.ancora_semana_1
"""
import datetime as dt

from . import config as C

_DIAS_PT = {
    0: "SEGUNDA-FEIRA", 1: "TERÇA-FEIRA", 2: "QUARTA-FEIRA", 3: "QUINTA-FEIRA",
    4: "SEXTA-FEIRA", 5: "SÁBADO", 6: "DOMINGO",
}
_MES_PT = {1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
           7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez"}


def _d(v):
    if isinstance(v, dt.date):
        return v
    return dt.date.fromisoformat(str(v))


def ancora(obra):
    info = C.obra_info(obra)
    a = info.get("ancora_semana_1")
    if not a:
        raise ValueError("obra %s sem ancora_semana_1 em road_registry.yaml" % obra)
    return _d(a)


def numero_relatorio(obra, data_inicio):
    """numero = (data_inicio - ancora)/7 + 1 -- regra CONFIRMADA (R4)."""
    di = _d(data_inicio)
    anc = ancora(obra)
    return (di - anc).days // 7 + 1


def resolver_periodo(obra, inicio=None, fim=None, semana=None):
    """Resolve (data_inicio, data_fim) a partir de --inicio/--fim OU --semana.

    - inicio + fim: usa os dois.
    - inicio sem fim: fim = inicio + 6 dias.
    - semana: resolve data_inicio = ancora + (semana-1)*7 dias; fim = +6 dias.
    """
    if semana is not None:
        if inicio is not None:
            raise ValueError("informe --inicio OU --semana, nao os dois")
        di = ancora(obra) + dt.timedelta(weeks=int(semana) - 1)
        return di, di + dt.timedelta(days=6)

    if inicio is None:
        raise ValueError("informe --inicio (ou --semana)")
    di = _d(inicio)
    df = _d(fim) if fim else di + dt.timedelta(days=6)
    return di, df


def dias_do_periodo(data_inicio):
    """Lista de 7 dias do periodo, com rotulo em portugues, caixa alta."""
    di = _d(data_inicio)
    dias = []
    for i in range(7):
        d = di + dt.timedelta(days=i)
        dias.append({
            "date": d.isoformat(),
            "label": _DIAS_PT[d.weekday()],
            "data_dd_mmm": "%02d/%s" % (d.day, _MES_PT[d.month]),
            "indice": i,
        })
    return dias


def schema_version(data_inicio):
    """Resolve a versao do formato do documento (v1/v2/v3...) vigente na
    data_inicio, usando report_versions.yaml:resolucao. Default = padrao_se_indeterminado."""
    di = _d(data_inicio)
    versoes = C.report_versions()
    resolucao = versoes.get("resolucao", {})
    padrao = resolucao.get("padrao_se_indeterminado", "v1")
    for vid, v in versoes.get("report_versions", {}).items():
        ini = v.get("effective_from")
        fim = v.get("effective_until")
        ini = _d(ini) if ini else None
        fim = _d(fim) if fim else None
        if ini and di < ini:
            continue
        if fim and di > fim:
            continue
        return vid
    return padrao


def calendario_regime(data_inicio):
    """Devolve o id do regime de calendario (cal_v1/cal_v2) vigente na data."""
    di = _d(data_inicio)
    versoes = C.report_versions()
    for cid, c in versoes.get("calendario", {}).items():
        ini = _d(c["effective_from"])
        fim = c.get("effective_until")
        fim = _d(fim) if fim else None
        if di >= ini and (fim is None or di <= fim):
            return cid
    return None
