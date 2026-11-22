# -*- coding: utf-8 -*-
"""
test_comparison.py

Testa scripts/gerador/compare.py (Fase 3, Slice 3, spec §8 item 4):

  K1  o caso âncora clima-2026-09-01 classifica GOLDEN_DEFECT (shadow=Chuvoso,
      golden=Bom -- a fonte confirma o Shadow, o golden está errado)
  K2  OAE x "Obras de Arte Especiais (OAE)" classifica SEMANTICALLY_EQUIVALENT
  K3  nenhuma diferença cai em SHADOW_ERROR por falta de regra: um item
      sintético sem `regra_id` e com golden != shadow cai em HITL_REQUIRED
      COM motivo -- nunca SHADOW_ERROR silencioso
  K4  a condição de uma regra só se aplica quando bate nos dados reais: uma
      regra cuja condição é falsa para o item NÃO classifica -- cai no
      fallback (a regra não é "cega")
  K5  compare.py é a única camada que lê o golden -- carregar_golden() lê o
      RO_NNN.n sem levantar PermissionError (não chama assert_nao_e_golden);
      o resto do gerador CONTINUA bloqueando o mesmo caminho
      (assert_nao_e_golden do proprio caminho golden levanta PermissionError)
  K6  classes retornadas sempre pertencem ao conjunto de 9 classes válidas;
      contagem_por_classe soma o total de itens

Execute:  python tests/test_comparison.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import sys
import json

import yaml

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from gerador import assert_nao_e_golden  # noqa: E402
from gerador import content_builder as CB  # noqa: E402
from gerador import compare as CMP  # noqa: E402

OBRA = "RD103"
PERIODO = ("2026-08-29", "2026-09-04")


def main():
    falhas = []
    outdir = os.path.join(U.OUT, "shadow", "%s_%s_%s" % (OBRA, *PERIODO))
    caminho_manifest = os.path.join(outdir, "manifest.yaml")
    caminho_data = os.path.join(outdir, "report_data.json")
    if not (os.path.exists(caminho_manifest) and os.path.exists(caminho_data)):
        print("SKIP: rode 'python scripts/shadow_report.py data --obra RD103 --inicio "
             "2026-08-29 --fim 2026-09-04' antes deste teste")
        return 1
    with open(caminho_manifest, encoding="utf-8") as fh:
        manifest = yaml.safe_load(fh)
    with open(caminho_data, encoding="utf-8") as fh:
        report_data = json.load(fh)

    documento = CB.construir_documento(report_data)
    resultado = CMP.comparar(report_data, documento, manifest)

    if resultado.get("erro"):
        print("SKIP: golden indisponível para comparação nesta sessão (%s)" % resultado["erro"])
        return 1

    itens_por_id = {i["id"]: i for i in resultado["itens"]}

    # ---------------------------------------------------------------- K1
    item = itens_por_id.get("clima-2026-09-01")
    if item and item["classe"] == "GOLDEN_DEFECT" and item["golden"] == "Bom" and item["shadow"] == "Chuvoso":
        print("K1  OK   clima-2026-09-01 -> GOLDEN_DEFECT (shadow=Chuvoso, golden=Bom), "
             "evidencia presente: %s" % bool(item.get("evidencia")))
    else:
        falhas.append("K1: item clima-2026-09-01 == %s" % item)

    # ---------------------------------------------------------------- K2
    item = itens_por_id.get("text-servicos")
    if item and item["classe"] == "SEMANTICALLY_EQUIVALENT" \
       and "OAE" in str(item["golden"]) and "Obras de Arte Especiais (OAE)" in str(item["shadow"]):
        print("K2  OK   servicos: golden='OAE' x shadow contem 'Obras de Arte Especiais (OAE)' "
             "-> SEMANTICALLY_EQUIVALENT")
    else:
        falhas.append("K2: item text-servicos == %s" % item)

    # ---------------------------------------------------------------- K3
    regras = CMP._carregar_regras()
    item_sintetico = {"id": "sintetico-sem-regra", "eixo": "text", "campo": "x",
                      "golden": "valor do golden", "shadow": "valor do shadow (diferente)",
                      "shadow_status": None, "regra_id": None}
    classe, motivo, regra_aplicada = CMP.classificar(item_sintetico, regras)
    if classe == "HITL_REQUIRED" and motivo:
        print("K3  OK   item sem regra e golden != shadow -> HITL_REQUIRED com motivo (%r), "
             "nunca SHADOW_ERROR" % motivo)
    else:
        falhas.append("K3: classificar(sem regra, golden!=shadow) == (%s, %r), esperado "
                      "HITL_REQUIRED com motivo" % (classe, motivo))

    # ---------------------------------------------------------------- K4
    # a regra clima-2026-09-01 SO classifica GOLDEN_DEFECT se shadow=='Chuvoso'
    # e golden=='Bom' -- com outros valores, a condicao e falsa e o item cai
    # no fallback (aqui, EXACT, porque os dois lados sao iguais)
    item_condicao_falsa = {"id": "teste-condicao", "eixo": "data", "campo": "x",
                           "golden": "Bom", "shadow": "Bom", "shadow_status": None,
                           "regra_id": "clima-2026-09-01"}
    classe2, _motivo2, regra2 = CMP.classificar(item_condicao_falsa, regras)
    if classe2 == "EXACT" and regra2 is None:
        print("K4  OK   regra clima-2026-09-01 NAO se aplica quando a condicao nao bate nos "
             "dados (golden==shadow=='Bom') -- regra nao e' cega, cai no fallback (EXACT)")
    else:
        falhas.append("K4: classificar(condicao falsa) == (%s, regra=%s), esperado EXACT sem regra"
                      % (classe2, regra2))

    # ---------------------------------------------------------------- K5
    golden_path = (manifest.get("golden_reference", {}).get("docx", {}) or {}).get("path")
    try:
        assert_nao_e_golden(golden_path)
        falhas.append("K5: assert_nao_e_golden NAO levantou PermissionError para o proprio "
                      "caminho do golden -- guarda quebrada para o resto do gerador")
    except PermissionError:
        print("K5a OK   assert_nao_e_golden continua bloqueando o caminho do golden para o "
             "resto do gerador")
    golden_lido = CMP.carregar_golden(manifest)
    if golden_lido and not golden_lido.get("erro") and golden_lido.get("paragraphs"):
        print("K5b OK   compare.carregar_golden() le o mesmo caminho sem PermissionError "
             "(excecao documentada em FASE3_SLICE_3_SPEC.md §0.5) -- %d parágrafos"
             % len(golden_lido["paragraphs"]))
    else:
        falhas.append("K5b: compare.carregar_golden() nao leu o golden: %s"
                      % (golden_lido.get("erro") if golden_lido else None))

    # ---------------------------------------------------------------- K6
    invalidas = [i["classe"] for i in resultado["itens"] if i["classe"] not in CMP.CLASSES_VALIDAS]
    soma = sum(resultado["contagem_por_classe"].values())
    if not invalidas and soma == len(resultado["itens"]):
        print("K6  OK   todas as classes pertencem ao conjunto valido; contagem soma %d == %d itens"
             % (soma, len(resultado["itens"])))
    else:
        falhas.append("K6: classes invalidas=%s, soma=%d, len(itens)=%d"
                      % (invalidas, soma, len(resultado["itens"])))

    print("=" * 74)
    print("contagem_por_classe:", resultado["contagem_por_classe"])
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
