# -*- coding: utf-8 -*-
r"""
test_figuras.py - Fase 3.3, FRENTE A: testa scripts/gerador/figuras/* contra
os dados reais de RD103, 2026-08-29..2026-09-04 (output/shadow/... ja
existente nesta sessao).

  F1  cada modulo devolve FiguraResult (path/largura_mm/altura_mm/status/
      note/provenance) com status em {OK, LIMITATION, PRESENTATION_DEPENDENCY,
      HITL}
  F2  quando path nao e None: o PNG existe em disco e sua proporcao
      (largura_px/altura_px) bate com largura_mm/altura_mm dentro de 1%
  F3  painel_fotografico: cada servico resolvido tem grid 3 colunas x
      ceil(n_fotos/3) linhas (n_fotos = provenance['n_fotos'])
  F4  toda `provenance.source_file` nao-nula aponta para um arquivo real
      (existe no disco OU no SharePoint via prog_util.exists)
  F5  nenhum caminho de figura (path ou provenance.source_file) vem do
      golden (RO_NNN.n - <rodovia>.docx/pdf, ou raiz de pasta-semana)

Execute:  python tests/test_figuras.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import re
import sys
import json
import math

import yaml
from PIL import Image

AQUI = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.dirname(AQUI)
SCRIPTS = os.path.join(PROJ, "scripts")
sys.path.insert(0, SCRIPTS)

import prog_util as U  # noqa: E402
from gerador.figuras import ficha_acumulada, avanco_grafico, painel_fotografico  # noqa: E402
from gerador.figuras import ambiental, nc_graficos  # noqa: E402

OBRA_DIR = os.path.join(PROJ, "output", "shadow", "RD103_2026-08-29_2026-09-04")
OUT_DIR = os.path.join(OBRA_DIR, "figuras")

_STATUS_VALIDOS = {"OK", "LIMITATION", "PRESENTATION_DEPENDENCY", "HITL"}
_RE_GOLDEN = re.compile(r"RO_\d{3}\.\d+\s*-\s*.+\.(docx|pdf)$", re.IGNORECASE)


def _carregar_dados():
    with open(os.path.join(OBRA_DIR, "report_data.json"), encoding="utf-8") as f:
        report_data = json.load(f)
    with open(os.path.join(OBRA_DIR, "manifest.yaml"), encoding="utf-8") as f:
        manifest = yaml.safe_load(f)
    return report_data, manifest


def _checar_contrato(nome, r, falhas):
    campos = ("path", "largura_mm", "altura_mm", "status", "note", "provenance")
    for c in campos:
        if not hasattr(r, c):
            falhas.append("F1: %s: FiguraResult sem campo '%s'" % (nome, c))
            return
    if r.status not in _STATUS_VALIDOS:
        falhas.append("F1: %s: status invalido '%s'" % (nome, r.status))
    if not isinstance(r.provenance, dict):
        falhas.append("F1: %s: provenance nao e dict" % nome)


def _checar_png_e_proporcao(nome, r, falhas):
    if r.path is None:
        return
    if not os.path.exists(r.path):
        falhas.append("F2: %s: path declarado mas PNG nao existe: %s" % (nome, r.path))
        return
    with Image.open(r.path) as im:
        w_px, h_px = im.size
    prop_esperada = r.largura_mm / r.altura_mm
    prop_real = w_px / h_px
    erro_rel = abs(prop_real - prop_esperada) / prop_esperada
    if erro_rel > 0.01:
        falhas.append("F2: %s: proporcao PNG (%.4f) diverge de largura_mm/altura_mm "
                      "(%.4f) em %.2f%%" % (nome, prop_real, prop_esperada, erro_rel * 100))
    else:
        print("F2  OK   %s: %dx%d px, %.2fx%.2f mm (proporcao bate, erro %.3f%%)"
             % (nome, w_px, h_px, r.largura_mm, r.altura_mm, erro_rel * 100))


def _checar_provenance_real(nome, r, falhas):
    sf = (r.provenance or {}).get("source_file")
    if sf in (None, "report_data.json"):
        return
    if not (os.path.exists(sf) or U.exists(sf)):
        falhas.append("F4: %s: provenance.source_file nao existe: %s" % (nome, sf))
    else:
        print("F4  OK   %s: source_file real (%s)" % (nome, os.path.basename(sf)))


def _checar_nao_e_golden(nome, r, falhas):
    for caminho in (r.path, (r.provenance or {}).get("source_file")):
        if not caminho:
            continue
        base = os.path.basename(caminho)
        pai = os.path.basename(os.path.dirname(caminho))
        if _RE_GOLDEN.search(base):
            falhas.append("F5: %s: caminho parece vir do golden: %s" % (nome, caminho))
        if re.search(r"\d{4}\.\d{2}\.\d{2}-\d{4}\.\d{2}\.\d{2}.*Relat", pai, re.IGNORECASE) \
                and os.path.splitext(base)[1].lower() in (".docx", ".pdf"):
            falhas.append("F5: %s: caminho na raiz de pasta-semana (golden/correlato): %s"
                         % (nome, caminho))


def main():
    falhas = []
    os.makedirs(OUT_DIR, exist_ok=True)
    report_data, manifest = _carregar_dados()
    road_key = report_data["metadata"]["road_key"]
    rodovia = report_data["road"]["rodovia"]

    resultados = {}

    # ---- ficha_acumulada (Figura 1) -----------------------------------
    caminho_acumulada = manifest["sources"]["accumulated_sheet"]["path"]
    resultados["ficha_acumulada"] = ficha_acumulada.gerar(caminho_acumulada, road_key, OUT_DIR)

    # ---- avanco_grafico (Grafico 1) ------------------------------------
    resultados["avanco_grafico"] = avanco_grafico.gerar(report_data, OUT_DIR)

    # ---- painel_fotografico (o item central) ---------------------------
    paineis = painel_fotografico.montar(report_data, manifest, OUT_DIR)
    for chave, r in paineis.items():
        resultados["painel_fotografico[%s]" % chave] = r

    # ---- ambiental (Figuras 6, 7, 8) ------------------------------------
    caminho_licenca = manifest["sources"]["environmental"]["licenca_ambiental_planilha"]["path"]
    resultados["ambiental_indices"] = ambiental.gerar_indices(caminho_licenca, rodovia, OUT_DIR)
    resultados["ambiental_condicionantes"] = ambiental.gerar_condicionantes(
        caminho_licenca, rodovia, OUT_DIR)
    resultados["ambiental_qrcode"] = ambiental.gerar_qrcode(OUT_DIR)

    # ---- nc_graficos (Figura 9, Graficos 2 e 3) -------------------------
    resultados["nc_figura_notificacoes"] = nc_graficos.gerar_figura_notificacoes(
        report_data, OUT_DIR)
    resultados["nc_grafico_status"] = nc_graficos.gerar_grafico_status(report_data, OUT_DIR)
    caminho_nc = manifest["sources"]["nc_tracker"]["path"]
    resultados["nc_grafico_natureza"] = nc_graficos.gerar_grafico_natureza(
        caminho_nc, rodovia, OUT_DIR)

    # ---- F1/F2/F4/F5 por resultado --------------------------------------
    for nome, r in resultados.items():
        _checar_contrato(nome, r, falhas)
        _checar_png_e_proporcao(nome, r, falhas)
        _checar_provenance_real(nome, r, falhas)
        _checar_nao_e_golden(nome, r, falhas)

    n_ok_status = sum(1 for r in resultados.values() if r.status in _STATUS_VALIDOS)
    if n_ok_status == len(resultados) and not any(f.startswith("F1") for f in falhas):
        print("F1  OK   %d/%d FiguraResult com contrato e status validos"
             % (n_ok_status, len(resultados)))

    # ---- F3: grid do painel fotografico (3 colunas, ceil(n/3) linhas) ---
    n_grids_checados = 0
    for chave, r in paineis.items():
        if r.status not in ("OK", "HITL") or "n_fotos" not in r.provenance:
            continue
        n_fotos = r.provenance["n_fotos"]
        grid_esperado = "3 colunas x %d linhas" % math.ceil(n_fotos / 3)
        grid_real = r.provenance.get("grid")
        if grid_real != grid_esperado:
            falhas.append("F3: painel_fotografico[%s]: grid '%s' != esperado '%s' (n_fotos=%d)"
                          % (chave, grid_real, grid_esperado, n_fotos))
        else:
            n_grids_checados += 1
    if n_grids_checados and not any(f.startswith("F3") for f in falhas):
        print("F3  OK   %d paineis com grid 3 colunas x ceil(n/3) linhas confirmado"
             % n_grids_checados)
    elif not n_grids_checados:
        falhas.append("F3: nenhum painel fotografico com n_fotos para checar o grid")

    print("=" * 78)
    if falhas:
        print("FALHAS (%d):" % len(falhas))
        for f in falhas:
            print("  X", f)
        print("RESULTADO: FALHOU")
        return 1
    print("RESULTADO: OK (%d figuras testadas)" % len(resultados))
    return 0


if __name__ == "__main__":
    sys.exit(main())
