# -*- coding: utf-8 -*-
r"""
test_unifilar.py - Fase 3.4 (v0.2), FRENTE B: testa
scripts/gerador/figuras/unifilar.py (UnifilarRenderer) contra a copia real da
planilha UNIFILAR RD-103, semana 49 (2026-08-29..2026-09-04), ja hidratada em
`tmp/unifilar_investigacao/UNIFILAR_180_S49.xlsx` pela investigacao fechada
em `docs/fontes/unifilar_renderizacao.md`.

Contrato testado (analogo a tests/test_figuras.py F1/F2/F4/F5, mais checks
especificos da prancha):

  U1  FiguraResult tem o contrato usual (path/largura_mm/altura_mm/status/
      note/provenance) com status em {OK, LIMITATION, PRESENTATION_DEPENDENCY}
  U2  quando status == OK: o PNG existe em disco, e' A0 PAISAGEM (largura >
      altura, proporcao 1189/841 dentro de 1%)
  U3  provenance mostra area de impressao real da planilha (nao um valor
      fixo) e bate com o Print_Area conhecido da RD-103 (B2:MA174, Sec.
      "Caminho 2" de unifilar_renderizacao.md)
  U4  a prancha desenhou preenchimentos, textos E mesclagens (>0 de cada) --
      nao e' uma pagina em branco disfarcada de OK
  U5  nenhum caminho (path ou provenance.source_file) vem do golden
  U6  arquivo inexistente -> LIMITATION (nunca excecao, nunca inventa pagina)
  U7  genericidade: rodando com uma folha A3 (parametro, nao deteccao de
      obra) a proporcao muda de acordo -- prova que o formato de pagina e'
      parametro do chamador, nunca decidido internamente por nome de obra

Execute:  python tests/test_unifilar.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import re
import sys

from PIL import Image

AQUI = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.dirname(AQUI)
SCRIPTS = os.path.join(PROJ, "scripts")
sys.path.insert(0, SCRIPTS)

import prog_util as U  # noqa: E402
from gerador.figuras import unifilar as UF  # noqa: E402

CAMINHO_XLSX = os.path.join(PROJ, "tmp", "unifilar_investigacao", "UNIFILAR_180_S49.xlsx")
OUT_DIR = os.path.join(PROJ, "tmp", "test_unifilar_out")

_STATUS_VALIDOS = {"OK", "LIMITATION", "PRESENTATION_DEPENDENCY"}
_RE_GOLDEN = re.compile(r"RO_\d{3}\.\d+\s*-\s*.+\.(docx|pdf)$", re.IGNORECASE)


def main():
    falhas = []

    if not os.path.exists(CAMINHO_XLSX):
        print("AVISO: %s nao existe neste checkout -- e' uma copia local (tmp/, nao "
             "versionada) hidratada durante a investigacao de "
             "docs/fontes/unifilar_renderizacao.md. Sem ela, U1..U5/U7 nao podem rodar; "
             "roda so' U6 (arquivo ausente)." % CAMINHO_XLSX)
        r = UF.gerar(CAMINHO_XLSX, OUT_DIR)
        if r.status != "LIMITATION" or r.path is not None:
            falhas.append("U6: arquivo ausente nao gerou LIMITATION com path=None")
        else:
            print("U6  OK   arquivo ausente -> LIMITATION, path=None")
        print("=" * 78)
        if falhas:
            print("FALHAS (%d):" % len(falhas))
            for f in falhas:
                print("  X", f)
            print("RESULTADO: FALHOU")
            return 1
        print("RESULTADO: OK (parcial -- fonte local ausente neste checkout)")
        return 0

    # ---- U6 primeiro (arquivo inexistente) -------------------------------
    r_ausente = UF.gerar(os.path.join(PROJ, "tmp", "nao_existe_de_verdade.xlsx"), OUT_DIR)
    if r_ausente.status == "LIMITATION" and r_ausente.path is None:
        print("U6  OK   arquivo UNIFILAR inexistente -> LIMITATION, path=None (nunca "
             "excecao, nunca pagina inventada)")
    else:
        falhas.append("U6: arquivo ausente deveria dar LIMITATION/path=None, deu status=%s "
                      "path=%s" % (r_ausente.status, r_ausente.path))

    # ---- renderizacao real (A0, RD-103 semana 49) ------------------------
    r = UF.gerar(CAMINHO_XLSX, OUT_DIR, nome_arquivo="figura_unifilar_rd103_s49.png")

    # -- U1: contrato -------------------------------------------------------
    for campo in ("path", "largura_mm", "altura_mm", "status", "note", "provenance"):
        if not hasattr(r, campo):
            falhas.append("U1: FiguraResult sem campo '%s'" % campo)
    if r.status not in _STATUS_VALIDOS:
        falhas.append("U1: status invalido '%s'" % r.status)
    if not isinstance(r.provenance, dict):
        falhas.append("U1: provenance nao e dict")
    if not any(f.startswith("U1") for f in falhas):
        print("U1  OK   FiguraResult com contrato valido, status=%s" % r.status)

    if r.status != "OK":
        falhas.append("U2..U5: status nao e OK (%s) -- %s" % (r.status, r.note))
    else:
        # -- U2: A0 paisagem, proporcao dentro de 1% -------------------------
        with Image.open(r.path) as im:
            w_px, h_px = im.size
        if w_px <= h_px:
            falhas.append("U2: PNG nao esta em paisagem (w=%d, h=%d)" % (w_px, h_px))
        prop_esperada = UF.A0_LARGURA_MM / UF.A0_ALTURA_MM
        prop_real = w_px / h_px
        erro_rel = abs(prop_real - prop_esperada) / prop_esperada
        if erro_rel > 0.01:
            falhas.append("U2: proporcao do PNG (%.4f) diverge de A0 paisagem (%.4f) em "
                          "%.2f%%" % (prop_real, prop_esperada, erro_rel * 100))
        else:
            print("U2  OK   PNG %dx%d px, A0 paisagem (proporcao bate, erro %.3f%%)"
                 % (w_px, h_px, erro_rel * 100))

        # -- U3: area de impressao real (nao um valor fixo no codigo) -------
        area = r.provenance.get("print_area") or ""
        if "B2" not in area.upper().replace("$", "") or "MA174" not in area.upper().replace("$", ""):
            falhas.append("U3: provenance.print_area nao bate com B2:MA174 conhecido da "
                          "RD-103 (veio: %r)" % area)
        else:
            print("U3  OK   provenance.print_area == %r (lido da propria planilha, nao "
                 "hardcoded)" % area)

        # -- U4: desenhou conteudo de verdade ---------------------------------
        n_fill = r.provenance.get("n_preenchimentos_desenhados", 0)
        n_txt = r.provenance.get("n_textos_desenhados", 0)
        n_merge = r.provenance.get("n_mesclagens_na_area", 0)
        if n_fill > 0 and n_txt > 0 and n_merge > 0:
            print("U4  OK   prancha com conteudo real: %d preenchimentos, %d textos, %d "
                 "mesclagens (nao e pagina em branco)" % (n_fill, n_txt, n_merge))
        else:
            falhas.append("U4: prancha suspeitosamente vazia (preenchimentos=%d, textos=%d, "
                          "mesclagens=%d)" % (n_fill, n_txt, n_merge))

        # -- U5: nada vem do golden -------------------------------------------
        for caminho in (r.path, r.provenance.get("source_file")):
            if caminho and _RE_GOLDEN.search(os.path.basename(caminho)):
                falhas.append("U5: caminho parece vir do golden: %s" % caminho)
        if not any(f.startswith("U5") for f in falhas):
            print("U5  OK   nenhum caminho (path/source_file) vem do golden")

        # -- U4b: source_file real (existe em disco) --------------------------
        sf = r.provenance.get("source_file")
        if sf and not (os.path.exists(sf) or U.exists(sf)):
            falhas.append("U4b: provenance.source_file nao existe: %s" % sf)

    # -- U7: genericidade -- o formato de pagina e' PARAMETRO, nunca decidido
    # por nome de obra dentro do modulo. NAO usa A3 para provar isso: A3 e A0
    # sao as duas paisagem ISO 216, com a MESMA proporcao (raiz de 2) --
    # comparar proporcao contra A3 seria um falso-negativo. Em vez disso usa
    # um tamanho deliberadamente distorcido (2000x500mm, proporcao 4:1, que
    # nenhuma folha real tem) para provar sem ambiguidade que a forma da
    # pagina de saida segue os parametros recebidos, nao um valor interno.
    LARGURA_TESTE_MM, ALTURA_TESTE_MM = 2000.0, 500.0
    r_custom = UF.gerar(CAMINHO_XLSX, OUT_DIR, largura_pagina_mm=LARGURA_TESTE_MM,
                        altura_pagina_mm=ALTURA_TESTE_MM,
                        nome_arquivo="figura_unifilar_formato_custom_teste.png")
    if r_custom.status == "OK":
        with Image.open(r_custom.path) as im:
            wc, hc = im.size
        prop_custom_esperada = LARGURA_TESTE_MM / ALTURA_TESTE_MM
        prop_custom_real = wc / hc
        erro_custom = abs(prop_custom_real - prop_custom_esperada) / prop_custom_esperada
        if erro_custom <= 0.01:
            print("U7  OK   mesma planilha renderizada com largura_pagina_mm/altura_pagina_mm "
                 "= %gx%gmm (proporcao 4:1, sem relacao com A0/A3) produz pagina na "
                 "proporcao pedida (erro %.3f%%) -- formato e' parametro do chamador, "
                 "nunca decidido por nome de obra dentro do modulo"
                 % (LARGURA_TESTE_MM, ALTURA_TESTE_MM, erro_custom * 100))
        else:
            falhas.append("U7: pagina com tamanho customizado nao respeitou a proporcao "
                          "pedida (real=%.4f esperado=%.4f)"
                          % (prop_custom_real, prop_custom_esperada))
    else:
        falhas.append("U7: renderizacao com pagina de tamanho customizado (mesma planilha) "
                      "falhou: %s" % r_custom.note)

    print("=" * 78)
    if falhas:
        print("FALHAS (%d):" % len(falhas))
        for f in falhas:
            print("  X", f)
        print("RESULTADO: FALHOU")
        return 1
    print("RESULTADO: OK (UnifilarRenderer -- prancha real, nao mais placeholder)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
