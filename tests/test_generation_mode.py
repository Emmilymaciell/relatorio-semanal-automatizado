# -*- coding: utf-8 -*-
"""
test_generation_mode.py — Fase 3.4 (v0.3)

Trava duas regras que nasceram do Checkpoint v0.2.

A) CURRENT × HISTORICAL_SHADOW
   O v0.2 afirmou no corpo do relatório "possui 3 licenças emitidas e 3
   pendentes" e "há 33 condicionantes próximas do vencimento" — números lidos
   HOJE sobre fontes que não guardam o estado da semana 49. O oficial diz 4/5
   e 1. Nem recalculando na data de emissão do oficial (11/09/2026) bate: dá
   31 naquele dia e 2 dez dias depois, sobre as MESMAS 45 condicionantes.

   M1  os dois modos existem e a inferência por data funciona
   M2  os dados irreconstituíveis estão declarados COM evidência
   M3  em HISTORICAL_SHADOW, nenhuma dessas afirmações é emitida
   M4  em HISTORICAL_SHADOW, o que substitui a afirmação é marcado
       LIMITATION (o leitor sabe que o dado existiu e não foi reconstruído)
   M5  em CURRENT, a redação normal volta — o modo histórico não pode
       emburrecer o relatório da semana corrente
   M6  nenhum número do golden vaza para o texto

B) Frase-ponte da seção 3 (TEXTO-01) é variação POR OBRA
   P1  road_registry traz a chave nas 4 obras
   P2  RD-101/RD-103 emitem; RD-102A/RD-104 não
   P3  o texto vem do catálogo, não está escrito no código

Execute:  python tests/test_generation_mode.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import re
import sys
import json

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from gerador import config as C  # noqa: E402
from gerador import content_builder as CB  # noqa: E402

# números que o relatório OFICIAL publica. Aparecem aqui só para provar que
# NÃO vazam para o candidate — o golden é referência de comparação, jamais
# fonte (CLAUDE.md, regra do Shadow Report).
NUMEROS_DO_GOLDEN = ("4 licenças emitidas", "5 licenças pendentes",
                     "1 condicionante próxima")

IDS_ESPERADOS = {"environment.licencas_kpis",
                 "environment.condicionantes_status",
                 "nonconformities.history_by_status"}

ESPERADO_PONTE = {"RD101": True, "RD102A": False, "RD103": True, "RD104": False}


def _texto_da_secao(doc, secao_id):
    s = [x for x in doc.sections if x.id == secao_id]
    if not s:
        return "", []
    blocos = [b for b in s[0].blocks if b.kind == "paragraph"]
    return "\n".join(b.text or "" for b in blocos), blocos


def main():
    falhas = []
    caminho = os.path.join(U.OUT, "shadow", "RD103_2026-08-29_2026-09-04", "report_data.json")
    if not os.path.exists(caminho):
        print("SKIP: gere o report_data da RD103 antes deste teste")
        return 1
    with open(caminho, encoding="utf-8") as fh:
        base = json.load(fh)

    # ------------------------------------------------------------------ M1
    modos = C.generation_modes().get("modos", {})
    if set(modos) != {"CURRENT", "HISTORICAL_SHADOW"}:
        falhas.append("M1: modos declarados == %s" % sorted(modos))
    else:
        import datetime as dt
        fim = dt.date(2026, 9, 4)
        folga = C.generation_modes()["inferencia"]["dias_apos_fim_para_historico"]
        perto = C.inferir_modo(fim, hoje=fim + dt.timedelta(days=folga))
        longe = C.inferir_modo(fim, hoje=fim + dt.timedelta(days=folga + 1))
        if perto != "CURRENT" or longe != "HISTORICAL_SHADOW":
            falhas.append("M1: inferência errada (folga=%d): %s / %s"
                          % (folga, perto, longe))
        else:
            print("M1  OK   dois modos; inferência vira HISTORICAL_SHADOW após %d dias"
                  % folga)

    # ------------------------------------------------------------------ M2
    decl = C.generation_modes().get("dados_nao_reconstruiveis_historicamente", [])
    ids = {d.get("id") for d in decl}
    sem_evidencia = [d.get("id") for d in decl
                     if not (d.get("evidencia") or "").strip()
                     or not (d.get("motivo") or "").strip()]
    if not IDS_ESPERADOS.issubset(ids):
        falhas.append("M2: faltam ids declarados: %s" % sorted(IDS_ESPERADOS - ids))
    elif sem_evidencia:
        falhas.append("M2: declarados sem motivo/evidência: %s" % sem_evidencia)
    else:
        print("M2  OK   %d dados irreconstituíveis declarados, todos com evidência"
              % len(decl))

    # ------------------------------------------------------------- M3/M4/M6
    rd_h = json.loads(json.dumps(base))
    doc_h = CB.construir_documento(rd_h, modo=C.MODO_HISTORICAL_SHADOW)
    if rd_h["metadata"].get("generation_mode") != C.MODO_HISTORICAL_SHADOW:
        falhas.append("M3: modo não foi gravado no report_data")
    texto_h, blocos_h = _texto_da_secao(doc_h, "s6_ambiental")
    plano_h = re.sub(r"\s+", " ", texto_h)

    afirma = []
    if re.search(r"possui\s+\d+\s+licenças?\s+emitidas", plano_h):
        afirma.append("contagem de licenças emitida como fato")
    if re.search(r"há\s+\d+\s+condicionantes?\s+próximas?", plano_h):
        afirma.append("contagem de condicionantes emitida como fato")
    if afirma:
        falhas.append("M3: em HISTORICAL_SHADOW o documento ainda afirma: %s" % afirma)
    else:
        print("M3  OK   HISTORICAL_SHADOW não afirma licenças nem condicionantes")

    marcados = [b for b in blocos_h
                if b.status == "LIMITATION" and "não foi reconstruída" in (b.text or "")]
    if len(marcados) < 2:
        falhas.append("M4: esperados 2 blocos LIMITATION substituindo as afirmações, "
                      "achei %d" % len(marcados))
    elif any(not (b.note or "").strip() for b in marcados):
        falhas.append("M4: bloco LIMITATION sem nota explicando o motivo")
    else:
        print("M4  OK   %d blocos LIMITATION no lugar das afirmações, com motivo"
              % len(marcados))

    vazou = [n for n in NUMEROS_DO_GOLDEN if n in plano_h]
    if vazou:
        falhas.append("M6: número do golden no texto do candidate: %s" % vazou)
    else:
        print("M6  OK   nenhum número do relatório oficial aparece no candidate")

    # ------------------------------------------------------------------ M5
    rd_c = json.loads(json.dumps(base))
    doc_c = CB.construir_documento(rd_c, modo=C.MODO_CURRENT)
    texto_c, blocos_c = _texto_da_secao(doc_c, "s6_ambiental")
    lim_c = [b for b in blocos_c
             if b.status == "LIMITATION" and "não foi reconstruída" in (b.text or "")]
    if lim_c:
        falhas.append("M5: modo CURRENT também suprimiu as afirmações (%d blocos) -- "
                      "o modo histórico não pode contaminar a semana corrente"
                      % len(lim_c))
    else:
        print("M5  OK   modo CURRENT não suprime nada (redação normal)")

    # ------------------------------------------------------------------ P1/P2
    reg = C.road_registry()["obras"]
    faltam = [o for o in ESPERADO_PONTE if "s3_frase_ponte_grafico" not in reg.get(o, {})]
    if faltam:
        falhas.append("P1: road_registry sem s3_frase_ponte_grafico em: %s" % faltam)
    else:
        print("P1  OK   s3_frase_ponte_grafico presente nas 4 obras")

    erros_p2 = []
    for obra, esperado in sorted(ESPERADO_PONTE.items()):
        if bool(reg[obra].get("s3_frase_ponte_grafico")) != esperado:
            erros_p2.append("%s: %s, esperado %s"
                            % (obra, reg[obra].get("s3_frase_ponte_grafico"), esperado))
    if erros_p2:
        falhas.append("P2 (registry): " + "; ".join(erros_p2))

    # e o efeito no documento, para a obra desta rodada
    frase = C._carregar("textos_v3.yaml")["textos"]["s3_frase_ponte_grafico"]["texto"]
    chave = re.sub(r"\s+", " ", frase).strip()[:60]
    s3_texto, _ = _texto_da_secao(doc_c, "s3_avanco_global")
    presente = chave in re.sub(r"\s+", " ", s3_texto)
    if presente != ESPERADO_PONTE["RD103"]:
        falhas.append("P2 (documento): RD-103 %s a frase-ponte, esperado %s"
                      % ("emite" if presente else "não emite", ESPERADO_PONTE["RD103"]))
    elif not erros_p2:
        print("P2  OK   RD-101/RD-103 emitem a frase-ponte; RD-102A/RD-104 não; "
              "a RD-103 a traz no documento")

    # ------------------------------------------------------------------ P3
    fonte_cb = open(os.path.join(SCRIPTS, "gerador", "content_builder.py"),
                    encoding="utf-8").read()
    if "Dessa forma, a análise conjunta" in fonte_cb:
        falhas.append("P3: a frase-ponte está escrita no código -- tem de vir de "
                      "textos_v3.yaml")
    elif re.search(r'if\s+obra\s*==|road_key\s*==\s*[\'"]GO', fonte_cb):
        falhas.append("P3: condicional por nome de obra no content_builder")
    else:
        print("P3  OK   frase vem do catálogo; nenhum condicional por nome de obra")

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
