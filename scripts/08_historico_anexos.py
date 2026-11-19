# -*- coding: utf-8 -*-
"""
08_historico_anexos.py

Responde a uma pergunta que a amostra de 2 semanas nao consegue responder:
o bloco ambiental (Parecer Tecnico PA-NNN + Acompanhamento ambiental + SST)
e QUINZENAL ou MENSAL?

Metodo: abre os PDFs das ultimas N semanas de cada obra e extrai APENAS a
lista de anexos da secao 11 — barato, sem extrair o documento inteiro.

Tambem cataloga a recorrencia de todos os tipos de anexo ao longo do tempo.

Saida: output/comparacoes/historico_anexos.json + .md
"""
import os
import sys
import re
import json
import argparse
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

import fitz

RE_ANEXO = re.compile(r"^\s*Anexo\s+([IVXLC]+)\s*[\-–—]\s*(.+?)\s*$", re.IGNORECASE | re.MULTILINE)
RE_SECANEXOS = re.compile(r"^\s*1[12]\.?\s*ANEXOS\s*$", re.IGNORECASE | re.MULTILINE)
RE_RO_FILE = re.compile(r"^RO[_\s-]*0*(\d{1,3})(?:\.(\d))?\s*-", re.IGNORECASE)


def normalizar_anexo(t):
    """Reduz o titulo do anexo a uma categoria estavel."""
    n = U.norm(t)
    if re.search(r"\bpa[\s\-]?\d", n):
        return "PARECER_TECNICO"
    if re.search(r"\bnc[\s\-]?\d", n):
        return "NAO_CONFORMIDADE"
    if "desvio" in n and "cronograma" in n:
        return "DESVIO_CRONOGRAMA"
    if "unifilar" in n:
        return "DIAGRAMA_UNIFILAR"
    if "diario" in n:
        return "DIARIO_DE_OBRAS"
    if "ambient" in n:
        return "ACOMP_AMBIENTAL"
    if "seguranca" in n or "sst" in n or "saude" in n:
        return "ACOMP_SST"
    if "ata" in n and "reuni" in n:
        return "ATA_REUNIAO"
    return "OUTRO:" + t[:40]


def lista_anexos(caminho):
    """Abre o PDF e devolve a lista de anexos. A secao e a 11 desde 2026-08-01
    e a 12 antes disso (quando existia a secao 6 DESVIO DE CRONOGRAMA)."""
    doc = fitz.open(U.lp(caminho))
    achados, pagina = [], None
    try:
        for i, page in enumerate(doc):
            r = page.rect
            if abs(r.width - 595) > 12:          # so paginas A4 retrato
                continue
            txt = page.get_text("text")
            if RE_SECANEXOS.search(txt):
                for m in RE_ANEXO.finditer(txt):
                    achados.append({"ordem": m.group(1), "titulo": m.group(2).strip(),
                                    "categoria": normalizar_anexo(m.group(2))})
                pagina = i + 1
                if achados:
                    break
        n_pag = doc.page_count
    finally:
        doc.close()
    return achados, pagina, n_pag


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ultimas", type=int, default=14, help="quantas semanas por obra")
    args = ap.parse_args()

    inv = json.load(open(os.path.join(U.OUT, "inventarios",
                                      "inventario_relatorios_semanais.json"), encoding="utf-8"))
    res = {"gerado_em": datetime.datetime.now().isoformat(timespec="seconds"),
           "ultimas_semanas": args.ultimas, "obras": {}}

    for obra in U.OBRAS:
        ss = [e for e in inv["obras"][obra]["entradas"]
              if e["tipo"] == "pasta" and "numero" in e and "data_inicio" in e]
        ss.sort(key=lambda e: e["numero"])
        linhas = []
        for s in ss[-args.ultimas:]:
            pdf = None
            for f in s.get("arquivos", []):
                if f["ext"] == ".pdf" and not f.get("subpasta") and RE_RO_FILE.match(f["nome"]):
                    pdf = f
                    break
            if not pdf:
                linhas.append({"numero": s["numero"], "data_inicio": s["data_inicio"],
                               "data_fim": s["data_fim"], "situacao": "sem PDF (em elaboracao)"})
                continue
            try:
                anexos, pagina, n_pag = lista_anexos(pdf["caminho"])
            except Exception as e:
                linhas.append({"numero": s["numero"], "data_inicio": s["data_inicio"],
                               "erro": str(e)})
                continue
            linhas.append({"numero": s["numero"], "data_inicio": s["data_inicio"],
                           "data_fim": s["data_fim"], "n_paginas": n_pag,
                           "pagina_anexos": pagina,
                           "anexos": anexos,
                           "categorias": [a["categoria"] for a in anexos]})
            print("  %-8s n=%-3d %s  %2d pgs  %d anexos: %s"
                  % (obra, s["numero"], s["data_inicio"], n_pag, len(anexos),
                     ", ".join(a["categoria"] for a in anexos)))
        res["obras"][obra] = linhas

    dest = os.path.join(U.OUT, "comparacoes")
    os.makedirs(U.assert_read_only(dest), exist_ok=True)
    with open(os.path.join(dest, "historico_anexos.json"), "w", encoding="utf-8") as fh:
        json.dump(res, fh, ensure_ascii=False, indent=1)

    # ---------- analise de periodicidade ----------
    L = ["# Histórico de anexos — periodicidade\n",
         "*Gerado por `scripts/08_historico_anexos.py`.*\n"]
    CATS = ["ATA_REUNIAO", "PARECER_TECNICO", "ACOMP_AMBIENTAL", "ACOMP_SST",
            "NAO_CONFORMIDADE", "DESVIO_CRONOGRAMA", "DIAGRAMA_UNIFILAR", "DIARIO_DE_OBRAS"]
    for obra, linhas in res["obras"].items():
        L.append("\n## %s\n" % obra)
        L.append("| nº | início | págs | " + " | ".join(c[:12] for c in CATS) + " |")
        L.append("|" + "---|" * (len(CATS) + 3))
        for l in linhas:
            if "categorias" not in l:
                L.append("| %s | %s | — | %s |" % (l["numero"], l.get("data_inicio", "?"),
                                                   " | ".join([l.get("situacao", l.get("erro", "?"))[:12]] +
                                                              [""] * (len(CATS) - 1))))
                continue
            cs = l["categorias"]
            L.append("| %s | %s | %s | %s |" % (
                l["numero"], l["data_inicio"], l["n_paginas"],
                " | ".join("✅" if c in cs else "" for c in CATS)))

    # intervalos do parecer tecnico
    L.append("\n## Periodicidade do Parecer Técnico (PA)\n")
    for obra, linhas in res["obras"].items():
        datas = [l["data_inicio"] for l in linhas
                 if "categorias" in l and "PARECER_TECNICO" in l["categorias"]]
        if len(datas) >= 2:
            dts = [datetime.date.fromisoformat(x) for x in datas]
            gaps = [(b - a).days // 7 for a, b in zip(dts, dts[1:])]
            L.append("- **%s**: PA em %s → intervalos de **%s semanas**"
                     % (obra, ", ".join(datas), ", ".join(map(str, gaps))))
        else:
            L.append("- **%s**: %d ocorrência(s) de PA na janela — insuficiente" % (obra, len(datas)))

    with open(os.path.join(dest, "historico_anexos.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(L))
    print("\nOK ->", os.path.join(dest, "historico_anexos.md"))


if __name__ == "__main__":
    main()
