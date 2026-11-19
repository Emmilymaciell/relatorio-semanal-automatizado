# -*- coding: utf-8 -*-
"""
tests/test_gates_genericos.py - multiobra-r1: critical gates sem valores da RD-103.

Ate a v0.3, G06 comparava com 32,23/42,19 fixos, G08 exigia zero NC e G10 exigia 5
anexos -- passavam na RD-103/49 por construcao. Estes testes fixam o comportamento
generico com textos minimos no formato do golden lido do DOCX (sem numeracao
automatica). Sem SharePoint, sem Word.
"""
import os
import sys
import tempfile

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "scripts"))

import montar_pdf_final as M  # noqa: E402

FALHAS = []


def chk(cod, ok, msg):
    print("%s  %s  %s" % (cod, "OK  " if ok else "FAIL", msg))
    if not ok:
        FALHAS.append(cod)


GOLDEN_UMA_NC = (
    "ACOMPANHAMENTO DE NÃO CONFORMIDADES:\n"
    "Não conformidades emitidas pela Programa no período:\n"
    "Durante o período ..., houve emissão de uma não conformidade emitida pelo Programa:\n"
    "NC-022: Emitida em 03/09/2026, durante a atividade ...\n"
    "Não conformidades emitidas pela Órgão Contratante no período:\n"
    "Durante o período ..., não houve emissão de não conformidades emitidas pela Órgão Contratante.\n")
GOLDEN_ZERO = (
    "Não conformidades emitidas pela Programa no período:\n"
    "Durante o período ..., não houve emissão de não conformidades emitidas pelo PROG.\n"
    "Não conformidades emitidas pela Órgão Contratante no período:\n"
    "Durante o período ..., não houve emissão de não conformidades emitidas pela Órgão Contratante.\n")
GOLDEN_FOTOS = (
    "Relatório Fotográfico\nTerraplenagem:\nFigura 2 - Relatório fotográfico RD-101.\n"
    "OAE\nFigura 3 - Relatório fotográfico RD-101.\n"
    "Obras Complementares\nFigura 7 - Relatório fotográfico RD-101.\n"
    "ACOMPANHAMENTO AMBIENTAL\n")


def main():
    # G08: a frase "não houve" da Órgão Contratante (8.2) nao pode zerar a contagem da Programa (8.1)
    chk("T1", M.golden_n_nc_prog(GOLDEN_UMA_NC) == 1,
        "8.1 com NC-022 conta 1 (a 8.2 diz 'não houve')")
    chk("T2", M.golden_n_nc_prog(GOLDEN_ZERO) == 0, "8.1 'não houve' conta 0")
    chk("T3", M.golden_n_nc_prog("texto sem secao 8") is None, "8.1 ausente -> None")

    # G07: subsecoes do fotografico no texto sem numeracao, ':' final removido
    chk("T4", M.golden_titulos_fotografico(GOLDEN_FOTOS) == [
        "Terraplenagem", "OAE", "Obras Complementares"], "titulos 5.2.x")

    # G10/G11: categorias -- NC numerada no golden == NC agrupada no shadow
    cats = [M._categoria_anexo(t) for t in (
        "PA-255", "Acompanhamento ambiental", "Acompanhamento saúde e segurança do trabalho",
        "NC-022", "Diagrama unifilar", "Diário de obra", "Ata de Reunião")]
    chk("T5", cats == ["parecer_tecnico", "acompanhamento_ambiental", "acompanhamento_sst",
                       "nao_conformidade", "diagrama_unifilar", "diario_de_obras",
                       "ata_reuniao"], "categorias de anexo: %s" % cats)
    chk("T6", M._categoria_anexo("nao_conformidade") == "nao_conformidade"
        and M._categoria_anexo("diario_de_obras") == "diario_de_obras", "ids internos")
    chk("T7", M._colapsar(["nao_conformidade", "nao_conformidade", "diario_de_obras"])
        == ["nao_conformidade", "diario_de_obras"], "NC-022 + NC-023 colapsam")

    # excecoes: so com entrada no yaml, so para gate FAIL, nunca para outra obra
    gates = [{"id": "G08", "status": "CRITICAL_GATE_FAIL", "evidencia": ""},
             {"id": "G05", "status": "CRITICAL_GATE_FAIL", "evidencia": ""},
             {"id": "G10", "status": "CRITICAL_GATE_PASS", "evidencia": ""}]
    res = M._aplicar_excecoes([dict(g) for g in gates], "RD104", "2026-08-29")
    st = {g["id"]: g["status"] for g in res}
    chk("T8", st == {"G08": "CRITICAL_GATE_PASS_COM_EXCECAO", "G05": "CRITICAL_GATE_FAIL",
                     "G10": "CRITICAL_GATE_PASS"}, "excecao GX-001 so no G08: %s" % st)
    chk("T9", res[0].get("excecao", {}).get("classe") == "GOLDEN_DEFECT"
        and res[0]["excecao"].get("evidencia"), "excecao carrega classe e evidencia")
    res2 = M._aplicar_excecoes([dict(g) for g in gates], "RD101", "2026-08-29")
    chk("T10", all(g["status"] != "CRITICAL_GATE_PASS_COM_EXCECAO" for g in res2),
        "outra obra nao herda a excecao")

    # nenhum valor da RD-103 sobrou no codigo dos gates
    src = open(M.__file__, encoding="utf-8").read()
    chk("T11", "32.23" not in src and "42.19" not in src and "== 5," not in src,
        "sem 32.23/42.19/len==5 fixos em montar_pdf_final.py")

    print("RESULTADO: %s" % ("OK" if not FALHAS else "FALHA %s" % FALHAS))
    return 1 if FALHAS else 0


if __name__ == "__main__":
    sys.exit(main())
