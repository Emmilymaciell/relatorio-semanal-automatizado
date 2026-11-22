# -*- coding: utf-8 -*-
"""
test_regras_periodo.py

Valida, sobre as 193 pastas-semana inventariadas (nao so a amostra), as regras
deterministicas descobertas — separando os DOIS REGIMES DE CALENDARIO:

  Regime 1 (set/2025 a nov/2025): periodo SEXTA -> QUINTA, com varias
           irregularidades de implantacao.
  Regime 2 (a partir de 2026-01-03, apos o pivo de 29/11/2025): periodo
           SABADO -> SEXTA, estavel ate hoje.

Regras testadas no regime 2 (o unico que interessa a automacao):
  R1  data_inicio e sabado e data_fim e sexta
  R2  o periodo tem exatamente 7 dias
  R3  semanas consecutivas sao contiguas
  R4  numero = (data_inicio - ancora da obra)/7 + 1
  R5  toda semana concluida (nao a ultima) tem PDF

Execute:  python tests/test_regras_periodo.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import sys
import json
import datetime

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(AQUI), "scripts"))
import prog_util as U

INV = os.path.join(U.OUT, "inventarios", "inventario_relatorios_semanais.json")

# Fronteira do regime 2: a partir daqui TODAS as obras usam sabado->sexta.
# Descoberta em 07_estrutura_e_diff / analise de padroes de periodo.
INICIO_REGIME_2 = datetime.date(2026, 1, 3)

# Erros de digitacao nos nomes de pasta, ja documentados. Nao sao regra.
DATAS_INVERTIDAS = {("RD101", 1), ("RD101", 2), ("RD104", 7)}


def d(s):
    return datetime.date.fromisoformat(s)


def semanas(inv, obra):
    ss = [e for e in inv["obras"][obra]["entradas"]
          if e["tipo"] == "pasta" and "numero" in e and "data_inicio" in e]
    ss.sort(key=lambda e: e["numero"])
    return ss


def main():
    inv = json.load(open(INV, encoding="utf-8"))
    falhas, avisos = [], []
    tot_r2 = 0
    print("=" * 78)
    print("VALIDACAO DAS REGRAS DETERMINISTICAS")
    print("Regime 2 (sabado->sexta) a partir de %s" % INICIO_REGIME_2)
    print("=" * 78)

    for obra in U.OBRAS:
        ss = semanas(inv, obra)
        r2 = [s for s in ss if d(s["data_inicio"]) >= INICIO_REGIME_2]
        r1 = [s for s in ss if d(s["data_inicio"]) < INICIO_REGIME_2]
        tot_r2 += len(r2)
        ok = {"R1": 0, "R2": 0, "R4": 0, "R5": 0}

        ancora = d(r2[0]["data_inicio"]) - datetime.timedelta(weeks=r2[0]["numero"] - 1)
        ultima = max(s["numero"] for s in ss)

        for s in r2:
            di, df = d(s["data_inicio"]), d(s["data_fim"])
            if di.weekday() == 5 and df.weekday() == 4:
                ok["R1"] += 1
            else:
                falhas.append("%s n=%s R1: %s(%s)..%s(%s)"
                              % (obra, s["numero"], di, di.strftime("%a"), df, df.strftime("%a")))
            if (df - di).days == 6:
                ok["R2"] += 1
            else:
                falhas.append("%s n=%s R2: %d dias" % (obra, s["numero"], (df - di).days + 1))
            calc = (di - ancora).days // 7 + 1
            if calc == s["numero"]:
                ok["R4"] += 1
            else:
                falhas.append("%s n=%s R4: calculado %d" % (obra, s["numero"], calc))
            tem_pdf = any(f["ext"] == ".pdf" and not f.get("subpasta") for f in s.get("arquivos", []))
            if tem_pdf or s["numero"] == ultima:
                ok["R5"] += 1
            else:
                avisos.append("%s n=%s R5: sem PDF e nao e a ultima semana" % (obra, s["numero"]))

        cont = sum(1 for a, b in zip(r2, r2[1:])
                   if d(b["data_inicio"]) - d(a["data_fim"]) == datetime.timedelta(days=1))
        if cont != len(r2) - 1:
            falhas.append("%s R3: %d/%d transicoes contiguas" % (obra, cont, len(r2) - 1))

        n = len(r2)
        print("%-8s regime1=%2d  regime2=%2d | R1 %d/%d  R2 %d/%d  R3 %d/%d  R4 %d/%d  R5 %d/%d"
              % (obra, len(r1), n, ok["R1"], n, ok["R2"], n, cont, n - 1, ok["R4"], n, ok["R5"], n))
        print("         ancora (semana 1 projetada) = %s   [%s]"
              % (ancora, ancora.strftime("%a")))

    # --- regime 1: so caracterizacao, nao validacao ---
    print("\n-- Regime 1 (anterior a %s): %d semanas, NAO seguem sabado->sexta." %
          (INICIO_REGIME_2, 193 - tot_r2))
    print("   Padrao dominante: SEXTA->QUINTA, com semanas de 8 e 15 dias na implantacao")
    print("   e 3 pastas com datas invertidas no nome: %s"
          % ", ".join("%s n=%d" % x for x in sorted(DATAS_INVERTIDAS)))

    print("\n" + "=" * 78)
    if avisos:
        print("AVISOS (%d):" % len(avisos))
        for a in avisos:
            print("  ~", a)
    if falhas:
        print("FALHAS (%d):" % len(falhas))
        for f in falhas[:30]:
            print("  X", f)
        print("\nRESULTADO: FALHOU")
        return 1
    print("RESULTADO: %d/%d semanas do regime 2 validadas em R1..R5." % (tot_r2, tot_r2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
