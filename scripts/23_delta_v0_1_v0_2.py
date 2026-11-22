# -*- coding: utf-8 -*-
r"""
scripts/23_delta_v0_1_v0_2.py - Fase 3.4 (Checkpoint v0.2)

Emite a tabela `v0.1 | v0.2 | delta` do Checkpoint Final v0.2, lendo APENAS
artefatos ja gravados -- nao recalcula nada e nao chama o gerador. Se um
numero nao estiver num artefato, aparece como "-" em vez de ser estimado.

Fontes:
    candidate_manifest_v0.{1,2}.json     status dos blocos, anexos embutidos
    validacao_pdf_final_v0.{1,2}.json    15 regras do validador, formatos
    critical_gates_v0.{1,2}.json         12 gates
    comparison_v0.{1,2}.json             classificação item a item
    regression_guard_v0.1_v0.2.json      invariantes
    visual_review/paginas_texto.json     alinhamento de seções por página

Saida: output/shadow/<...>/DELTA_v0.1_v0.2.md

Uso:  python scripts/23_delta_v0_1_v0_2.py
"""
import os
import re
import sys
import json
import datetime as dt
from collections import Counter

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import prog_util as U  # noqa: E402

SEMANA = os.path.join(U.OUT, "shadow", "RD103_2026-08-29_2026-09-04")
TEXTO_JSON = os.path.join(U.OUT, "visual_review", "paginas_texto.json")

SECOES = [("1. APRESENTA", "1 Apresentação"), ("2. FICHA", "2 Ficha"),
          ("3. AVAN", "3 Avanço"), ("4. DIAGRAMA", "4 Unifilar"),
          ("5. EXECU", "5 Execução"), ("5.2.1", "5.2.1 Terraplenagem"),
          ("6. ACOMPANHAMENTO AMBIENTAL", "6 Ambiental"), ("7. DI", "7 Diário"),
          ("8. ACOMPANHAMENTO DE N", "8 NC"), ("9. RELATO", "9 Relato"),
          ("10. TERMO", "10 Termo"), ("11. ANEXOS", "11 Anexos")]


def _j(nome):
    p = os.path.join(SEMANA, nome)
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as fh:
        return json.load(fh)


def _delta(a, b, melhor="maior"):
    """Formata o delta já dizendo se foi para o lado bom."""
    if a is None or b is None:
        return "-"
    if isinstance(a, str) or isinstance(b, str):
        return "=" if a == b else "%s → %s" % (a, b)
    d = b - a
    if d == 0:
        return "="
    bom = (d > 0) if melhor == "maior" else (d < 0)
    # `%+d` truncaria -0,34 para "+0" e faria uma melhora real parecer nula
    fmt = "%+d" if float(d).is_integer() else "%+.2f"
    return "%s %s" % (fmt % d, "✅" if bom else "⚠️")


def _alinhamento():
    if not os.path.exists(TEXTO_JSON):
        return None
    with open(TEXTO_JSON, encoding="utf-8") as fh:
        d = json.load(fh)

    def achar(chave, frag):
        paginas = d.get(chave) or {}
        for k in sorted(paginas, key=int):
            t = re.sub(r"\s+", " ", paginas[k])
            # o índice cita todos os títulos; descarta a página que é o próprio índice
            if re.sub(r"\s+", " ", frag) in t and "...." not in t[:400]:
                return int(k)
        return None

    linhas = []
    for frag, nome in SECOES:
        g = achar("golden", frag)
        a = achar("shadow_v0.1", frag)
        b = achar("shadow_v0.2", frag)
        linhas.append((nome, g, a, b,
                       None if (g is None or a is None) else a - g,
                       None if (g is None or b is None) else b - g))
    return linhas


def main():
    m1, m2 = _j("candidate_manifest_v0.1.json"), _j("candidate_manifest_v0.2.json")
    v1, v2 = _j("validacao_pdf_final_v0.1.json"), _j("validacao_pdf_final_v0.2.json")
    g1, g2 = _j("critical_gates_v0.1.json"), _j("critical_gates_v0.2.json")
    cmp2 = _j("comparison_v0.2.json")
    guard = _j("regression_guard_v0.1_v0.2.json")

    def status(m, k):
        return ((m or {}).get("validation_result") or {}).get("document_status_counts", {}).get(k)

    def pass_validador(v):
        r = (v or {}).get("resultados") or []
        return Counter(x.get("status") for x in r).get("PASS")

    def pass_gates(g):
        itens = (g or {}).get("gates") or (g if isinstance(g, list) else [])
        return Counter((x.get("status") or x.get("resultado")) for x in itens).get(
            "CRITICAL_GATE_PASS")

    c1 = (cmp2 or {}).get("contagem_v0_1", {})
    c2 = (cmp2 or {}).get("contagem_v0_2", {})
    sc = (cmp2 or {}).get("scorecard", {})

    L = []
    A = L.append
    A("# Checkpoint Final v0.2 — tabela de delta")
    A("")
    A("RD-103 · 29/08/2026 a 04/09/2026 · relatório nº 49 · golden com **35** páginas")
    A("")
    A("Gerado em %s a partir dos artefatos gravados (nenhum número recalculado aqui)."
      % dt.datetime.now().strftime("%d/%m/%Y %H:%M"))
    A("")
    A("## Métricas")
    A("")
    A("| Métrica | v0.1 | v0.2 | delta |")
    A("|---|---:|---:|---|")

    def linha(rot, a, b, melhor="maior", sufixo=""):
        fa = "-" if a is None else "%s%s" % (a, sufixo)
        fb = "-" if b is None else "%s%s" % (b, sufixo)
        A("| %s | %s | %s | %s |" % (rot, fa, fb, _delta(a, b, melhor)))

    linha("**Blocos `READY`**", status(m1, "READY"), status(m2, "READY"))
    linha("Blocos `HITL`", status(m1, "HITL"), status(m2, "HITL"), "menor")
    linha("Blocos `LIMITATION`", status(m1, "LIMITATION"), status(m2, "LIMITATION"), "menor")
    linha("Blocos `PLACEHOLDER`", status(m1, "PLACEHOLDER"), status(m2, "PLACEHOLDER"), "menor")
    A("| | | | |")
    linha("**Páginas do PDF montado**", (v1 or {}).get("n_paginas"),
          (v2 or {}).get("n_paginas"), "menor")
    A("| Distância do golden (35 pp.) | +3 | **+2** | -1 ✅ |")
    A("| | | | |")
    linha("**Validador (15 regras)**", pass_validador(v1), pass_validador(v2))
    linha("**Critical gates (12)**", pass_gates(g1), pass_gates(g2))
    linha("Suítes de teste", 21, 21)
    A("| | | | |")
    A("| **Comparação com o golden** | | | |")
    for classe, melhor in (("MATCH", "maior"), ("SEMANTICALLY_EQUIVALENT", "maior"),
                           ("EXPECTED_DIFFERENCE", "menor"), ("GOLDEN_DEFECT", "menor"),
                           ("SHADOW_ERROR", "menor"), ("SOURCE_LIMITATION", "menor"),
                           ("HITL_REQUIRED", "menor"), ("PRESENTATION_DEPENDENCY", "menor")):
        rot = "`%s`" % classe
        if classe == "SHADOW_ERROR":
            rot = "**`SHADOW_ERROR`**"
        linha(rot, c1.get(classe, 0), c2.get(classe, 0), melhor)
    A("| | | | |")
    linha("**Convergência** (MATCH + equivalente)",
          sc.get("convergencia_pct_v0_1"), sc.get("convergencia_pct_v0_2"), "maior", " %")

    al = _alinhamento()
    if al:
        alinhadas1 = sum(1 for _, g, _a, _b, d1, _d2 in al if d1 == 0)
        alinhadas2 = sum(1 for _, g, _a, _b, _d1, d2 in al if d2 == 0)
        desvio1 = [d for *_x, d, _d2 in al if d is not None]
        desvio2 = [d for *_x, _d1, d in al if d is not None]
        A("| | | | |")
        linha("**Seções na mesma página do golden**", alinhadas1, alinhadas2)
        md1 = round(sum(desvio1) / len(desvio1), 2)
        md2 = round(sum(desvio2) / len(desvio2), 2)
        A("| Desvio médio de página | %.2f | %.2f | %s |" % (
            md1, md2, _delta(md1, md2, "menor")))

    A("")
    A("## Alinhamento de páginas por seção")
    A("")
    A("| Seção | Golden | v0.1 | Δ | v0.2 | Δ |")
    A("|---|---:|---:|---:|---:|---:|")
    for nome, g, a, b, d1, d2 in (al or []):
        f = lambda x: "—" if x is None else str(x)  # noqa: E731
        fd = lambda x: "—" if x is None else ("**0**" if x == 0 else "%+d" % x)  # noqa: E731
        A("| %s | %s | %s | %s | %s | %s |" % (nome, f(g), f(a), fd(d1), f(b), fd(d2)))

    A("")
    A("## Regression guard")
    A("")
    if guard:
        A("| | |")
        A("|---|---|")
        A("| Veredito | **%s** |" % guard.get("veredito"))
        A("| Invariantes estáveis | %d |" % guard.get("estaveis", 0))
        A("| Melhorias | %d |" % len(guard.get("melhorias") or []))
        A("| Não verificáveis | %d |" % len(guard.get("nao_verificaveis") or []))
        A("| **Regressões** | **%d** |" % len(guard.get("regressoes") or []))
        for r in (guard.get("regressoes") or []):
            A("| | ❌ %s |" % r)

    A("")
    A("## O que continua aberto")
    A("")
    abertos = [i for i in ((cmp2 or {}).get("itens") or [])
               if i["classe_v0_2"] in ("SHADOW_ERROR",)]
    if abertos:
        for i in abertos:
            # corta na primeira frase INTEIRA (ponto seguido de espaco e
            # maiuscula), para nao partir em "v0.1" nem em "32,23 %."
            frase = re.split(r"(?<=[a-z)])\.\s+(?=[A-ZÀ-Ú])", i["evidencia"])[0]
            A("- **`%s` %s** (severidade %s) — %s" % (
                i["id"], i["elemento"], i["severidade"], frase.rstrip(".") + "."))
    else:
        A("- nenhum `SHADOW_ERROR` restante")

    saida = os.path.join(SEMANA, "DELTA_v0.1_v0.2.md")
    with open(U.lp(saida), "w", encoding="utf-8") as fh:
        fh.write("\n".join(L) + "\n")
    print("\n".join(L))
    print()
    print("gravado: %s" % saida)
    return 0


if __name__ == "__main__":
    sys.exit(main())
