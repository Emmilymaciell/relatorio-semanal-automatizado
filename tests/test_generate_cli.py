# -*- coding: utf-8 -*-
"""
tests/test_generate_cli.py - comando unico scripts/generate.py, SEM Word e SEM SharePoint.

  G1  parsing: normaliza RD-103/rd103, --end default = start+6, defaults de versao/modo
  G2  parsing: obra invalida e --end < --start sao rejeitados
  G3  veredito READY com tudo PASS
  G4  veredito BLOCKED por: FAIL do validador; gate nao-PASS (inclui N/A); validador ausente;
      gates ausentes; etapa com erro; guard BLOQUEIA
  G5  scorecard: contagens, pendencias (secoes nao-READY e figuras fora de OK) e Markdown
  G6  fluxo em disco: _concluir() grava summary+scorecard a partir de JSONs-fixture e devolve exit code
  G7  assert_read_only barra escrita no SharePoint
Roda com `python tests/test_generate_cli.py` ou com pytest.
"""
import json
import os
import shutil
import sys
import tempfile

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "scripts"))

import prog_util as U  # noqa: E402
import generate as G  # noqa: E402


def _validador(falhas=()):
    return {"n_paginas": 38, "resultados": [
        {"id": "V%d" % i, "regra": "regra %d" % i, "status": "FAIL" if ("V%d" % i) in falhas else "PASS",
         "detalhe": "x"} for i in range(1, 16)]}


def _gates(nao_pass=()):
    return {"n_paginas_final": 38, "gates": [
        {"id": "G%02d" % i, "gate": "g%d" % i,
         "status": nao_pass[0][1] if nao_pass and ("G%02d" % i) == nao_pass[0][0] else "CRITICAL_GATE_PASS",
         "evidencia": "e"} for i in range(1, 13)]}


def _manifest():
    return {"validation_result": {
        "document_status_counts": {"READY": 102, "HITL": 4, "LIMITATION": 5, "PLACEHOLDER": 1},
        "document_section_status": {"Capa": "READY", "1 APRESENTAÇÃO": "HITL",
                                    "6 ACOMPANHAMENTO AMBIENTAL": "LIMITATION"},
        "figuras_frente_a": {"ficha_acumulada": "OK", "painel_fotografico": {"p1": "OK", "p2": "FALHA"}}}}


OK = {"candidate": {"status": "OK"}, "pdf": {"status": "OK"}, "assemble": {"status": "OK"},
      "guard": {"status": "SKIPPED"}}


def _sc(**kw):
    base = dict(obra="RD103", periodo={"start": "2026-08-29", "end": "2026-09-04"}, versao="t1",
                base_versao=None, validador=_validador(), gates=_gates(), guard=None,
                manifest=_manifest(), etapas=OK)
    base.update(kw)
    return G.montar_scorecard(**base)


# ---------------------------------------------------------------- G1 / G2
def test_g1_parsing():
    a = G.parse_args(["--road", "RD-103", "--start", "2026-08-29"])
    assert a.obra == "RD103" and str(a.fim) == "2026-09-04"
    assert a.versao == "mr1" and a.modo == "HISTORICAL_SHADOW" and a.base_versao is None
    assert not a.skip_pdf and not a.pdf_manual and a.resume_from is None
    a = G.parse_args(["--road", "rd102a", "--start", "2026-08-29", "--end", "2026-09-04",
                      "--versao", "v9", "--modo", "CURRENT", "--base-versao", "v8",
                      "--pdf-manual", "--resume-from", "assemble"])
    assert a.obra == "RD102A" and a.modo == "CURRENT" and a.base_versao == "v8"
    assert a.pdf_manual and a.resume_from == "assemble"


def test_g2_parsing_invalido():
    try:
        G.parse_args(["--road", "RD999", "--start", "2026-08-29"])
        assert False, "deveria rejeitar obra"
    except ValueError:
        pass
    try:
        G.parse_args(["--road", "RD103", "--start", "2026-09-04", "--end", "2026-08-29"])
        assert False, "deveria rejeitar end < start"
    except SystemExit as e:
        assert e.code == 2


# ---------------------------------------------------------------- G3 / G4
def test_g3_ready():
    v, m = G.decidir_veredito(_validador(), _gates(), OK, None)
    assert v == G.VEREDITO_READY and m == []


def test_g4_blocked():
    v, m = G.decidir_veredito(_validador(("V7",)), _gates(), OK)
    assert v == G.VEREDITO_BLOCKED and any("V7" in x for x in m)
    v, m = G.decidir_veredito(_validador(), _gates((("G03", "CRITICAL_GATE_FAIL"),)), OK)
    assert v == G.VEREDITO_BLOCKED and any("G03" in x for x in m)
    v, m = G.decidir_veredito(_validador(), _gates((("G05", "N/A"),)), OK)
    assert v == G.VEREDITO_BLOCKED, "gate N/A tambem e' nao-PASS"
    assert G.decidir_veredito(None, _gates(), OK)[0] == G.VEREDITO_BLOCKED
    assert G.decidir_veredito(_validador(), None, OK)[0] == G.VEREDITO_BLOCKED
    et = dict(OK, assemble={"status": "ERRO"})
    assert G.decidir_veredito(_validador(), _gates(), et)[0] == G.VEREDITO_BLOCKED
    g = {"veredito": "BLOQUEIA", "regressoes": ["figura X: OK -> FALHA"]}
    assert G.decidir_veredito(_validador(), _gates(), OK, g)[0] == G.VEREDITO_BLOCKED
    g = {"veredito": "LIBERA_COM_JUSTIFICATIVA", "regressoes": []}
    assert G.decidir_veredito(_validador(), _gates(), OK, g)[0] == G.VEREDITO_READY


# ---------------------------------------------------------------- G5
def test_g5_scorecard():
    sc = _sc(marcadores=["[HITL] 1 APRESENTAÇÃO"])
    assert sc["veredito"] == G.VEREDITO_READY and sc["n_paginas"] == 38
    assert sc["validador"]["pass"] == 15 and sc["validador"]["fail"] == 0
    assert sc["gates"]["pass"] == 12 and sc["gates"]["total"] == 12
    assert sc["blocos_por_status"]["HITL"] == 4
    onde = [p["onde"] for p in sc["pendencias"]]
    assert "1 APRESENTAÇÃO" in onde and "painel_fotografico.p2" in onde
    assert "Capa" not in onde and "ficha_acumulada" not in onde
    md = G.scorecard_md(sc)
    assert "READY_FOR_SHADOW_REVIEW" in md and "Critical gates (12/12 PASS)" in md
    assert "[HITL] 1 APRESENTAÇÃO" in md
    sc2 = _sc(validador=_validador(("V2",)))
    assert sc2["veredito"] == G.VEREDITO_BLOCKED and sc2["validador"]["fail"] == 1
    assert "V2" in G.scorecard_md(sc2)


# ---------------------------------------------------------------- G6
def test_g6_fluxo_em_disco():
    tmp = tempfile.mkdtemp(prefix="prog_gen_test_")
    try:
        outdir = tmp
        v = "t1"
        for nome, obj in (("validacao_pdf_final_%s.json" % v, _validador()),
                          ("critical_gates_%s.json" % v, _gates()),
                          ("candidate_manifest_%s.json" % v, _manifest())):
            with open(os.path.join(outdir, nome), "w", encoding="utf-8") as fh:
                json.dump(obj, fh)
        a = G.parse_args(["--road", "RD103", "--start", "2026-08-29", "--versao", v])
        resumo = {"etapas": {"candidate": {"status": "OK"}, "assemble": {"status": "OK"}}}
        rc = G._concluir(a, outdir, resumo, os.path.join(outdir, "inexistente.docx"),
                         os.path.join(outdir, "inexistente_MONTADO.pdf"), True)
        assert rc == G.EXIT_READY
        assert os.path.exists(os.path.join(outdir, "scorecard_t1.json"))
        assert os.path.exists(os.path.join(outdir, "SCORECARD_t1.md"))
        s = json.load(open(os.path.join(outdir, "generation_summary_t1.json"), encoding="utf-8"))
        assert s["veredito"] == G.VEREDITO_READY and s["contagem"]["HITL"] == 4
        # agora com um FAIL no validador -> BLOCKED / exit 1
        with open(os.path.join(outdir, "validacao_pdf_final_%s.json" % v), "w", encoding="utf-8") as fh:
            json.dump(_validador(("V1",)), fh)
        resumo = {"etapas": {"candidate": {"status": "OK"}, "assemble": {"status": "OK"}}}
        assert G._concluir(a, outdir, resumo, "x.docx", "x.pdf", True) == G.EXIT_BLOCKED
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------- G7
def test_g7_read_only():
    try:
        G._gravar(os.path.join(U.SHAREPOINT, "x", "scorecard.md"), "x")
        assert False, "deveria bloquear"
    except PermissionError:
        pass


if __name__ == "__main__":
    falhas = 0
    for nome, fn in sorted(globals().items()):
        if nome.startswith("test_") and callable(fn):
            try:
                fn()
                print("PASS %s" % nome)
            except Exception as e:  # noqa: BLE001
                falhas += 1
                print("FAIL %s: %s: %s" % (nome, type(e).__name__, e))
    print("RESULTADO: %s" % ("OK" if not falhas else "%d FALHA(S)" % falhas))
    sys.exit(1 if falhas else 0)
