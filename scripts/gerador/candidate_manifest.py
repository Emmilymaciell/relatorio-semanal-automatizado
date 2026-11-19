# -*- coding: utf-8 -*-
r"""
scripts/gerador/candidate_manifest.py - Fase 3.3, Sec.6: manifesto de versão
do candidate -- hashes de FONTE e de CONFIGURAÇÃO, para que o PDF final
(quando existir) possa ser reproduzido a partir de um commit + um conjunto de
arquivos identificados por sha1 (nunca "confia na memória"; qualquer byte que
mudou na fonte ou na config aparece aqui).

`source_hashes` é coletado GENERICAMENTE: varre `report_data`/`manifest`
(saída de `gerador.canonical.construir`) procurando qualquer valor-string que
seja um caminho de arquivo REALMENTE EXISTENTE -- não lista campo por campo
(um novo campo de origem entra sozinho na próxima rodada, sem precisar tocar
neste módulo). `config_hashes` cobre todo `config/*.yaml` presente (fonte
única de configuração do renderer/gerador).
"""
import datetime as dt
import os
import subprocess

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(AQUI)
import sys  # noqa: E402
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402


def _git_commit():
    try:
        out = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=U.PROJ, stderr=subprocess.DEVNULL)
        return out.decode("ascii").strip()
    except Exception:
        return None


def _parece_caminho(v):
    return isinstance(v, str) and len(v) > 3 and (":\\" in v or v.startswith("\\\\") or "/" in v)


def coletar_caminhos_fonte(*estruturas):
    """Varre recursivamente dicts/lists procurando valores-string que sejam
    caminhos de arquivo existentes (nunca pasta) -- generico, sem `if campo
    == "..."`. Usado para achar toda fonte REAL referenciada em
    report_data.json/manifest.yaml desta rodada."""
    vistos = set()

    def _rec(v):
        if isinstance(v, dict):
            for vv in v.values():
                _rec(vv)
        elif isinstance(v, (list, tuple)):
            for vv in v:
                _rec(vv)
        elif _parece_caminho(v):
            try:
                if U.exists(v) and not U.isdir(v):
                    vistos.add(v)
            except (OSError, ValueError):
                pass

    for e in estruturas:
        _rec(e)
    return sorted(vistos)


def hashes_config():
    config_dir = os.path.join(U.PROJ, "config")
    out = {}
    if not os.path.isdir(config_dir):
        return out
    for nome in sorted(os.listdir(config_dir)):
        if nome.endswith(".yaml"):
            out[nome] = U.sha1(os.path.join(config_dir, nome))
    return out


def montar(*, obra, road_label, periodo, report_data, manifest,
          extra_source_paths=None, renderer_version="v2",
          candidate_version="v0.1", validation_result=None):
    """Monta o dict do manifesto (Sec.6). NAO grava em disco -- quem chama
    decide o caminho (sempre dentro de output/shadow/, ver
    shadow_report.cmd_candidate)."""
    caminhos = set(coletar_caminhos_fonte(report_data, manifest))
    for p in (extra_source_paths or []):
        if p and U.exists(p) and not U.isdir(p):
            caminhos.add(p)

    source_hashes = {}
    for caminho in sorted(caminhos):
        try:
            source_hashes[caminho] = U.sha1(caminho)
        except OSError as e:
            source_hashes[caminho] = "ERRO_AO_LER: %s" % e

    return {
        "candidate_version": candidate_version,
        "commit": _git_commit(),
        "generated_at": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
        "road": road_label,
        "obra_key": obra,
        "period": {"start": periodo[0], "end": periodo[1]},
        "source_hashes": source_hashes,
        "config_hashes": hashes_config(),
        "renderer_version": renderer_version,
        "validation_result": validation_result or {},
    }
