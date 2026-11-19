# -*- coding: utf-8 -*-
r"""
24_levantar_frase_ponte_s3.py - Fase 3.4 (v0.3), GF-1 item restante.

A v0.2 trouxe os 3 paragrafos conceituais da secao 3 (levantados por
20_levantar_texto_gf1.py), mas o levantamento CORTAVA o texto num limite de
caracteres e a ultima frase do bloco -- a frase-ponte que apresenta o
grafico -- ficou de fora. Este script levanta exatamente essa frase.

Mesmo metodo ja aprovado:
  - corpus v3 ja extraido em output/extracao/texto/
  - GOLDEN ATUAL EXCLUIDO (RD103_RO049.1) -- o golden serve para comparar,
    nunca como fonte
  - agrupa por esqueleto (numeros trocados por marcador) e reporta o
    predominante com quantos relatorios o sustentam

Saida: output/analises/texto_ponte_s3.json

Execute:  python scripts/24_levantar_frase_ponte_s3.py
"""
import os
import re
import sys
import json
import collections

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U  # noqa: E402

TEXTO_DIR = os.path.join(U.OUT, "extracao", "texto")
GOLDEN = "RD103_RO049.1"
SAIDA = os.path.join(U.OUT, "analises", "texto_ponte_s3.json")

# a frase fecha o bloco conceitual e chama o grafico; comeca por um conectivo
# de conclusao e cita "gr[áa]fico"
_RE_PONTE = re.compile(
    r"((?:Dessa forma|Desta forma|Assim|Portanto|Diante disso)[^.]{0,400}?"
    r"gr[áa]fico[^.]{0,400}?\.)",
    re.IGNORECASE)

# limites do bloco da secao 3.
# ATENCAO: o titulo aparece DUAS vezes -- no indice e no corpo. A primeira
# ocorrencia e' a entrada do indice ("AVANCO GLOBAL DA OBRA ......... 5") e
# pegar essa (foi o que a primeira versao deste script fez) delimita um bloco
# que nao contem a secao nenhuma. A do CORPO e' a que NAO vem seguida de
# pontilhado de tabulacao.
_RE_S3_INDICE = re.compile(r"AVAN[ÇC]O\s+GLOBAL\s+DA\s+OBRA\s*\.{4,}", re.IGNORECASE)
_RE_S3 = re.compile(r"AVAN[ÇC]O\s+GLOBAL\s+DA\s+OBRA", re.IGNORECASE)
_RE_PROX_SECAO = re.compile(r"\n\s*4\s*[\.\-]?\s*\n?\s*DIAGRAMA", re.IGNORECASE)


def _bloco_secao3(txt):
    """Devolve o corpo da secao 3, ou None."""
    for m in _RE_S3.finditer(txt):
        if _RE_S3_INDICE.match(txt, m.start()):
            continue                      # entrada do indice, nao o corpo
        resto = txt[m.end():]
        fim = _RE_PROX_SECAO.search(resto)
        return resto[:fim.start()] if fim else resto[:4000]
    return None


def _norm(s):
    return re.sub(r"\s+", " ", s).strip()


def _esqueleto(s):
    """Troca numeros por marcador, para agrupar variantes que so diferem em
    valor -- numero de frase-ponte, se houver, vem de report_data, nunca do
    template."""
    return re.sub(r"\d+[.,]?\d*", "<N>", s)


def main():
    if not os.path.isdir(TEXTO_DIR):
        print("ERRO: corpus nao encontrado em %s" % TEXTO_DIR)
        return 2

    arquivos = sorted(f for f in os.listdir(TEXTO_DIR)
                      if f.endswith(".txt") and GOLDEN not in f
                      and "OBSOLETO" not in f.upper())

    com_secao = 0
    contagem = collections.Counter()
    exemplos = {}
    por_obra = collections.defaultdict(collections.Counter)

    for nome in arquivos:
        with open(os.path.join(TEXTO_DIR, nome), encoding="utf-8", errors="replace") as fh:
            txt = fh.read()
        bloco = _bloco_secao3(txt)
        if bloco is None:
            continue
        com_secao += 1

        achados = _RE_PONTE.findall(bloco)
        if not achados:
            contagem["(sem frase-ponte)"] += 1
            por_obra[nome.split("_")[0]]["(sem frase-ponte)"] += 1
            continue
        frase = _norm(achados[-1])
        esq = _esqueleto(frase)
        contagem[esq] += 1
        por_obra[nome.split("_")[0]][esq] += 1
        exemplos.setdefault(esq, (nome, frase))

    total = sum(contagem.values())
    top = []
    for esq, n in contagem.most_common(8):
        nome, frase = exemplos.get(esq, (None, esq))
        top.append({"n": n, "pct": round(100.0 * n / total, 1) if total else 0.0,
                    "esqueleto": esq, "exemplo_arquivo": nome, "exemplo_texto": frase})

    saida = {
        "golden_excluido": GOLDEN,
        "metodo": ("corpus v3 em output/extracao/texto/, golden atual excluido; "
                   "bloco da secao 3 delimitado pelo titulo e pela secao 4; frase-ponte "
                   "= ultimo periodo que comeca por conectivo de conclusao e cita "
                   "'grafico'; agrupada por esqueleto (numeros -> <N>)"),
        "relatorios_analisados": len(arquivos),
        "relatorios_com_a_secao": com_secao,
        "variantes": len([k for k in contagem if k != "(sem frase-ponte)"]),
        "top": top,
        "por_obra": {o: dict(c.most_common(4)) for o, c in sorted(por_obra.items())},
    }
    os.makedirs(os.path.dirname(SAIDA), exist_ok=True)
    with open(U.lp(SAIDA), "w", encoding="utf-8") as fh:
        json.dump(saida, fh, ensure_ascii=False, indent=2)

    print("relatorios analisados (golden excluido): %d" % len(arquivos))
    print("com a secao 3: %d" % com_secao)
    print()
    for t in top:
        print("  %4d  %5.1f%%  %s" % (t["n"], t["pct"], (t["exemplo_texto"] or "")[:150]))
    print()
    print("por obra:")
    for o, c in saida["por_obra"].items():
        print("  %-8s %s" % (o, {k[:60]: v for k, v in c.items()}))
    print()
    print("gravado: %s" % SAIDA)
    return 0


if __name__ == "__main__":
    sys.exit(main())
