# -*- coding: utf-8 -*-
"""
17_levantar_texto_s1.py - levanta do corpus o texto literal do paragrafo
contratual de abertura da secao 1 (APRESENTACAO) e a redacao real da frase de
periodo/servicos.

Le SOMENTE os textos ja extraidos em output/extracao/texto/. Exclui o golden que
esta em comparacao (RD103_RO049.1) - regra "De onde vem os textos fixos" da
FASE3_SLICE_3_SPEC.md.
"""
import os
import re
import sys
import json
import collections

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

TEXTO_DIR = os.path.join(U.OUT, "extracao", "texto")
GOLDEN = "RD103_RO049.1"


def corpo_apresentacao(t):
    m = re.search(r"1\.\s*APRESENTA[ÇC][ÃA]O\s*\n(.{0,2600})", t, re.S)
    if not m:
        return None
    corpo = m.group(1)
    fim = re.search(r"\n\s*2\.\s", corpo)
    if fim:
        corpo = corpo[:fim.start()]
    return corpo


def paragrafos(corpo):
    # o PDF quebra linha no meio da frase; junta e reparte por frase final
    txt = re.sub(r"\s*\n\s*", " ", corpo).strip()
    txt = re.sub(r"\s{2,}", " ", txt)
    return txt


def main():
    arqs = sorted(f for f in os.listdir(TEXTO_DIR) if f.endswith(".txt"))
    contratual = collections.Counter()
    periodo = collections.Counter()
    exemplos = {}
    lidos = 0

    for a in arqs:
        if GOLDEN in a or "OBSOLETO" in a:
            continue
        t = open(os.path.join(TEXTO_DIR, a), encoding="utf-8", errors="replace").read()
        corpo = corpo_apresentacao(t)
        if not corpo:
            continue
        lidos += 1
        plano = paragrafos(corpo)

        # --- paragrafo contratual: do inicio ate "Parceiro."
        m = re.match(r"(O presente relat.{0,900}?Parceiro\.)", plano)
        if m:
            frase = m.group(1)
            # normaliza os campos variaveis por obra para achar o esqueleto
            esq = re.sub(r"rodovia RD[- ]?\d+\s?A?", "rodovia {RODOVIA}", frase)
            esq = re.sub(r"trecho:.*?extens[ãa]o:.*?km", "trecho: {TRECHO}, extensão: {EXT}", esq)
            esq = re.sub(r"contrato \d{2}/\d{4}", "contrato {CONTRATO_PROG}", esq)
            contratual[esq] += 1
            exemplos.setdefault(esq, (a, frase))

        # --- frase de periodo e servicos
        m2 = re.search(r"(As informa[çc][õo]es apresentadas neste relat.{0,260}?\.)\s", plano)
        if m2:
            f = m2.group(1)
            esq2 = re.sub(r"\d{2}/\d{2}/\d{4}", "{DATA}", f)
            esq2 = re.sub(r"(atividades referentes a|atividades de)\s.*?\.", r"\1 {SERVICOS}.", esq2)
            periodo[esq2] += 1
            exemplos.setdefault(esq2, (a, f))

    out = {
        "relatorios_lidos": lidos,
        "golden_excluido": GOLDEN,
        "paragrafo_contratual": [
            {"n": n, "exemplo_arquivo": exemplos[e][0], "esqueleto": e}
            for e, n in contratual.most_common(5)
        ],
        "frase_periodo_servicos": [
            {"n": n, "exemplo_arquivo": exemplos[e][0], "esqueleto": e}
            for e, n in periodo.most_common(5)
        ],
    }
    dest = os.path.join(U.OUT, "analises", "texto_s1_apresentacao.json")
    U.assert_read_only(dest)
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    with open(dest, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)

    print("relatorios com secao 1 legivel:", lidos)
    print("\n=== PARAGRAFO CONTRATUAL - variantes ===")
    for e, n in contratual.most_common(4):
        print("\n[%d ocorrencias] (ex.: %s)" % (n, exemplos[e][0]))
        print(e[:700])
    print("\n=== FRASE PERIODO/SERVICOS - variantes ===")
    for e, n in periodo.most_common(6):
        print("[%3d] %s" % (n, e[:190]))
    print("\njson:", dest)


if __name__ == "__main__":
    main()
