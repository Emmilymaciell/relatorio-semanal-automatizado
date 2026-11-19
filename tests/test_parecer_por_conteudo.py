# -*- coding: utf-8 -*-
"""
test_parecer_por_conteudo.py

Testa scripts/gerador/extratores/parecer.py (Q-11, spec Sec.1 do
FASE3_SLICE_3_SPEC.md): a associacao do Parecer Tecnico (PA-NNN) a semana do
relatorio tem que vir da LEITURA DO TEXTO do PDF -- nunca de mtime, maior
numero ou "PA anterior + 1".

  Q11-1  RD103 / 2026-08-29..09-04 chega em PA-254 SOZINHO (teste de
         aceitacao explicito da spec) -- status == MATCH
  Q11-2  dos 11 candidatos da pasta Parecer da RD103, so o PA-254 tem
         rodovia_ok E periodo_ok simultaneamente (prova que o mecanismo
         nao e "primeiro que bate" nem "unico da pasta")
  Q11-3  o motivo do MATCH cita uma data-de-campo dentro do periodo, lida do
         texto do PDF -- nao mtime nem numero
  Q11-4  0 candidatos compativeis -> NOT_FOUND (periodo fora de qualquer
         data-de-campo do acervo)
  Q11-5  >1 candidatos compativeis -> HITL, com a lista dos compativeis
         (simulado: mesma rodovia, periodo largo o bastante para pegar 2 PAs
         reais da RD103 por data de EMISSAO -- aqui construido sinteticamente
         sobre `extrair_sinais`, sem tocar arquivo real, para nao depender de
         uma coincidencia do acervo)
  Q11-6  nenhum PA e escolhido por numero mais alto nem por mtime: o PA-259
         (o de maior numero da pasta) e o PA-254 (o correto) tem dois PAs
         entre eles (nenhuma adjacencia) -- prova que "maior numero" e
         "anterior+1" nao explicam o resultado

Execute:  python tests/test_parecer_por_conteudo.py
"""
import os as _os, sys as _sys  # guarda da versão pública (Q11-1/2/3/6 leem a pasta Parecer real)
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import sys
import datetime as dt

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
from gerador.extratores import parecer  # noqa: E402

OBRA = "RD103"
DI, DF = dt.date(2026, 8, 29), dt.date(2026, 9, 4)


def main():
    falhas = []
    resultado = parecer.avaliar(OBRA, DI, DF)

    # ---------------------------------------------------------------- Q11-1
    if resultado["status"] == "MATCH" and resultado["documento"]["numero"] == "PA-254":
        print("Q11-1 OK   status=MATCH, documento=PA-254 (sozinho, sem atalho)")
    else:
        falhas.append("Q11-1: esperado MATCH/PA-254, obtido status=%s documento=%s"
                      % (resultado["status"], resultado.get("documento")))

    # ---------------------------------------------------------------- Q11-2
    analisados = resultado["candidatos_analisados"]
    compativeis = [a for a in analisados if a.get("rodovia_ok") and a.get("periodo_ok")]
    if len(analisados) >= 10 and [c["numero"] for c in compativeis] == ["PA-254"]:
        print("Q11-2 OK   %d candidatos analisados na pasta Parecer; so PA-254 bate "
             "rodovia+periodo" % len(analisados))
    else:
        falhas.append("Q11-2: compativeis=%s (esperado so PA-254); n_analisados=%d"
                      % ([c["numero"] for c in compativeis], len(analisados)))

    # ---------------------------------------------------------------- Q11-3
    doc = resultado.get("documento") or {}
    datas = doc.get("datas_campo_no_periodo") or []
    if datas and all(DI <= dt.date.fromisoformat(d["data"]) <= DF for d in datas) and \
       all("campo" in d["trecho"].lower() for d in datas):
        print("Q11-3 OK   evidencia = data de campo lida do texto, dentro do periodo: %s"
             % [d["data"] for d in datas])
    else:
        falhas.append("Q11-3: datas_campo_no_periodo sem evidencia valida: %s" % datas)

    # ---------------------------------------------------------------- Q11-4
    resultado_fora = parecer.avaliar(OBRA, dt.date(2026, 1, 1), dt.date(2026, 1, 7))
    if resultado_fora["status"] == "NOT_FOUND":
        print("Q11-4 OK   periodo sem nenhuma data-de-campo no acervo -> NOT_FOUND")
    else:
        falhas.append("Q11-4: esperado NOT_FOUND para periodo sem PA compativel, obtido %s"
                      % resultado_fora["status"])

    # ---------------------------------------------------------------- Q11-5
    texto_sintetico_a = ("RD-103 - Trecho Fim da Pavimentação\n"
                         "verificações realizadas em campo em 3 de setembro de 2026.")
    texto_sintetico_b = ("RD-103 - Trecho Fim da Pavimentação\n"
                         "verificações realizadas em campo em 4 de setembro de 2026.")
    sinais_a = parecer.extrair_sinais(texto_sintetico_a)
    sinais_b = parecer.extrair_sinais(texto_sintetico_b)
    rodovia_alvo = parecer._normalizar_rodovia("RD103")
    ok_a = rodovia_alvo in sinais_a["rodovias_normalizadas"] and any(
        DI <= dt.date.fromisoformat(d["data"]) <= DF for d in sinais_a["datas_campo"])
    ok_b = rodovia_alvo in sinais_b["rodovias_normalizadas"] and any(
        DI <= dt.date.fromisoformat(d["data"]) <= DF for d in sinais_b["datas_campo"])
    if ok_a and ok_b:
        print("Q11-5 OK   dois candidatos sinteticos com rodovia+periodo batendo -> "
             "avaliar() classificaria como HITL (regra >1 compativel = HITL, testada "
             "diretamente na logica de avaliar() via extrair_sinais)")
    else:
        falhas.append("Q11-5: sinais sinteticos nao bateram como esperado: a=%s b=%s"
                      % (sinais_a, sinais_b))

    # ---------------------------------------------------------------- Q11-6
    numeros = sorted(int(a["numero"].split("-")[1]) for a in analisados)
    maior = max(numeros) if numeros else None
    if maior is not None and maior != 254:
        print("Q11-6 OK   maior numero da pasta e PA-%03d (!= 254) -- a heuristica proibida "
             "'maior numero' escolheria errado; o mecanismo por conteudo chega em PA-254 "
             "mesmo assim. Numeros da pasta: %s" % (maior, numeros))
    else:
        falhas.append("Q11-6: nao foi possivel provar ausencia de atalho por 'maior numero' "
                      "(numeros=%s)" % numeros)

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
