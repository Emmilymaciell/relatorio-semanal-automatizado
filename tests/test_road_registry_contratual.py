# -*- coding: utf-8 -*-
"""
test_road_registry_contratual.py

Testa o cadastro mestre (config/road_registry.yaml) e a exposicao em
report_data.json -> contract.* (spec Sec.3, §7): os campos contratuais usados
no DOCX (fora de escopo desta fatia) tem que vir do CADASTRO, nunca do
golden, e cada campo novo (instrumento_convocacao, ordem_servico,
ordem_servico_em, contrato_assinado_em) tem `contrato_provenance` com
arquivo + trecho -- OU fica NOT_FOUND + HITL quando nao comprovado.

  R1  RD102A, RD103 e RD104 tem instrumento_convocacao e ordem_servico(_em)
      preenchidos, cada um com `contrato_provenance` citando arquivo (dentro
      de '2_Documentacao Licitatoria\\3_Contrato Executora') e trecho
  R2  RD102A e RD103 NAO tem contrato_assinado_em inventado -- fica None,
      com `contrato_provenance.contrato_assinado_em.status == NOT_FOUND` e
      motivo nao vazio (nunca copiado do golden)
  R3  RD104 tem contrato_assinado_em confirmado (unica das 3 com bloco de
      assinatura textual extraivel no proprio contrato), com provenance
      citando pagina + trecho
  R4  report_data.json (canonical.construir) espelha os mesmos valores do
      cadastro em `contract.*`, e report_provenance.json tem entrada
      'contract.<campo>' para TODOS os 6 campos, nas 4 obras -- inclusive
      os NOT_FOUND (com validation_status HITL, nunca ausencia silenciosa)
  R5  nenhum arquivo citado em contrato_provenance aponta para o golden
      (RO_NNN.n - ....{docx,pdf})

Execute:  python tests/test_road_registry_contratual.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import re
import sys
import datetime as dt

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
from gerador import config as C  # noqa: E402
from gerador.canonical import construir  # noqa: E402

_RE_GOLDEN = re.compile(r"RO_\d{3}\.\d+\s*-\s*.+\.(docx|pdf)$", re.IGNORECASE)

CAMPOS_CONTRATO = ("contrato_obra", "contrato_prog", "instrumento_convocacao",
                   "contrato_assinado_em", "ordem_servico", "ordem_servico_em")


def main():
    falhas = []
    reg = C.road_registry()

    # ---------------------------------------------------------------- R1
    for obra in ("RD102A", "RD103", "RD104"):
        info = reg["obras"][obra]
        prov_c = info.get("contrato_provenance", {})
        problemas = []
        for campo in ("instrumento_convocacao", "ordem_servico", "ordem_servico_em"):
            if not info.get(campo):
                problemas.append("%s sem valor" % campo)
                continue
            origem = prov_c.get(campo)
            if not origem or not origem.get("arquivo") or not origem.get("trecho"):
                problemas.append("%s sem contrato_provenance completa: %s" % (campo, origem))
        if problemas:
            falhas.append("R1 %s: %s" % (obra, problemas))
        else:
            print("R1  OK   %-7s instrumento_convocacao=%s ordem_servico=%s/%s, todos com "
                 "provenance (arquivo+trecho)" % (obra, info["instrumento_convocacao"],
                                                   info["ordem_servico"], info["ordem_servico_em"]))

    # ---------------------------------------------------------------- R2
    for obra in ("RD102A", "RD103"):
        info = reg["obras"][obra]
        prov_c = info.get("contrato_provenance", {})
        st = prov_c.get("contrato_assinado_em", {})
        if info.get("contrato_assinado_em") is None and st.get("status") == "NOT_FOUND" \
           and st.get("motivo"):
            print("R2  OK   %-7s contrato_assinado_em=None + NOT_FOUND/HITL com motivo "
                 "(nao inventado, nao copiado do golden)" % obra)
        else:
            falhas.append("R2 %s: contrato_assinado_em=%r contrato_provenance=%s"
                          % (obra, info.get("contrato_assinado_em"), st))

    # ---------------------------------------------------------------- R3
    info461 = reg["obras"]["RD104"]
    prov461 = info461.get("contrato_provenance", {}).get("contrato_assinado_em", {})
    if info461.get("contrato_assinado_em") == dt.date(2025, 9, 16) and prov461.get("arquivo") \
       and prov461.get("trecho"):
        print("R3  OK   RD104 contrato_assinado_em=2025-09-16, com arquivo+trecho de origem")
    else:
        falhas.append("R3: RD104 contrato_assinado_em=%r provenance=%s"
                      % (info461.get("contrato_assinado_em"), prov461))

    # ---------------------------------------------------------------- R4
    for obra in C.obras_validas():
        manifest, data, prov = construir(obra, dt.date(2026, 8, 29), dt.date(2026, 9, 4))
        d = prov.as_dict()
        info = reg["obras"][obra]
        for campo in CAMPOS_CONTRATO:
            chave = "contract.%s" % campo
            if chave not in d:
                falhas.append("R4 %s: sem entrada de proveniencia '%s'" % (obra, chave))
                continue
            if data["contract"].get(campo) != info.get(campo) and \
               not (data["contract"].get(campo) is None and info.get(campo) is None):
                # contrato_prog vem do meta, nao do info -- tolerado
                if campo != "contrato_prog":
                    falhas.append("R4 %s.%s: report_data=%r cadastro=%r"
                                  % (obra, campo, data["contract"].get(campo), info.get(campo)))
            if data["contract"].get(campo) is None and \
               d[chave].get("validation_status") != "HITL":
                falhas.append("R4 %s.%s: valor None mas sem validation_status=HITL: %s"
                              % (obra, campo, d[chave]))
    if not any(f.startswith("R4") for f in falhas):
        print("R4  OK   report_data.contract.* == cadastro, com proveniencia (inclusive "
             "HITL para os NOT_FOUND) nas %d obras" % len(C.obras_validas()))

    # ---------------------------------------------------------------- R5
    achados_golden = []
    for obra, info in reg["obras"].items():
        for campo, origem in info.get("contrato_provenance", {}).items():
            arquivo = (origem or {}).get("arquivo")
            if arquivo and _RE_GOLDEN.search(os.path.basename(arquivo)):
                achados_golden.append((obra, campo, arquivo))
    if not achados_golden:
        print("R5  OK   nenhum contrato_provenance aponta para o golden")
    else:
        falhas.append("R5: contrato_provenance aponta para golden: %s" % achados_golden)

    print("=" * 74)
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
