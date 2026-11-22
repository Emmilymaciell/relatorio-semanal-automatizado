# -*- coding: utf-8 -*-
"""
03_extrair_pdf.py

Extracao estrutural READ-ONLY dos PDFs da amostra, com PyMuPDF.
Para cada PDF gera:
  output/extracao/pdf/<OBRA>_RO<NNN>.json  - estrutura completa (paginas, blocos, fontes, imagens)
  output/extracao/pdf/<OBRA>_RO<NNN>.txt   - texto por pagina, para leitura humana e diff

Nao usa OCR: o texto e extraido da camada de texto do PDF.
Os PDFs sao abertos direto do SharePoint em modo leitura (nunca salvos de volta).
"""
import os
import sys
import json
import argparse
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

import fitz  # PyMuPDF

DEST = os.path.join(U.OUT, "extracao", "pdf")


def cor_hex(c):
    return "#%06x" % (int(c) & 0xFFFFFF)


def extrair(caminho_pdf, rotulo):
    doc = fitz.open(U.lp(caminho_pdf))
    info = {
        "rotulo": rotulo,
        "arquivo": caminho_pdf,
        "n_paginas": doc.page_count,
        "metadata": {k: v for k, v in (doc.metadata or {}).items() if v},
        "toc": doc.get_toc(simple=True),
        "paginas": [],
    }
    fontes = Counter()
    tamanhos = Counter()
    dims = Counter()

    for i, page in enumerate(doc):
        r = page.rect
        dims[(round(r.width, 1), round(r.height, 1))] += 1
        pg = {
            "n": i + 1,
            "largura_pt": round(r.width, 1),
            "altura_pt": round(r.height, 1),
            "rotacao": page.rotation,
            "orientacao": "retrato" if r.height >= r.width else "paisagem",
            "blocos": [],
            "imagens": [],
            "desenhos": 0,
        }
        d = page.get_text("dict")
        for b in d.get("blocks", []):
            if b.get("type") != 0:
                continue
            linhas = []
            for ln in b.get("lines", []):
                spans = []
                for sp in ln.get("spans", []):
                    txt = sp.get("text", "")
                    if not txt.strip():
                        continue
                    f = sp.get("font", "")
                    t = round(sp.get("size", 0), 1)
                    fontes[f] += len(txt)
                    tamanhos[t] += len(txt)
                    spans.append({
                        "t": txt, "font": f, "size": t,
                        "cor": cor_hex(sp.get("color", 0)),
                        "flags": sp.get("flags", 0),
                        "bbox": [round(x, 1) for x in sp.get("bbox", [])],
                    })
                if spans:
                    linhas.append(spans)
            if linhas:
                pg["blocos"].append({
                    "bbox": [round(x, 1) for x in b.get("bbox", [])],
                    "linhas": linhas,
                })
        for im in page.get_images(full=True):
            xref = im[0]
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            pg["imagens"].append({
                "xref": xref, "largura_px": im[2], "altura_px": im[3],
                "bpc": im[4], "colorspace": im[5],
                "rects": [[round(x, 1) for x in (r.x0, r.y0, r.x1, r.y1)] for r in rects],
            })
        try:
            pg["desenhos"] = len(page.get_drawings())
        except Exception:
            pass
        info["paginas"].append(pg)

    info["resumo"] = {
        "fontes": fontes.most_common(),
        "tamanhos": sorted(tamanhos.items(), key=lambda kv: -kv[1]),
        "dimensoes_pagina": [{"w": k[0], "h": k[1], "n": v} for k, v in dims.most_common()],
        "total_imagens": sum(len(p["imagens"]) for p in info["paginas"]),
        "paginas_sem_texto": [p["n"] for p in info["paginas"] if not p["blocos"]],
    }
    doc.close()
    return info


def texto_plano(info):
    out = []
    for p in info["paginas"]:
        out.append("\n" + "=" * 78)
        out.append("### PAGINA %d  (%s %.0fx%.0f pt, %d imgs)"
                   % (p["n"], p["orientacao"], p["largura_pt"], p["altura_pt"], len(p["imagens"])))
        out.append("=" * 78)
        for b in p["blocos"]:
            for linha in b["linhas"]:
                out.append("".join(s["t"] for s in linha))
            out.append("")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--obras", nargs="*", default=U.OBRAS)
    args = ap.parse_args()

    am = json.load(open(os.path.join(U.PROJ, "state", "amostra.json"), encoding="utf-8"))
    os.makedirs(U.assert_read_only(DEST), exist_ok=True)

    for obra in args.obras:
        for s in am["obras"][obra]["selecionados"]:
            if not s["pdf"]:
                print("%s n=%s: sem PDF" % (obra, s["numero"]))
                continue
            rot = "%s_RO%03d" % (obra, s["numero"])
            info = extrair(s["pdf"]["caminho"], rot)
            info["obra"] = obra
            info["numero"] = s["numero"]
            info["data_inicio"] = s["data_inicio"]
            info["data_fim"] = s["data_fim"]
            with open(os.path.join(DEST, rot + ".json"), "w", encoding="utf-8") as fh:
                json.dump(info, fh, ensure_ascii=False)
            with open(os.path.join(DEST, rot + ".txt"), "w", encoding="utf-8") as fh:
                fh.write(texto_plano(info))
            r = info["resumo"]
            print("%-12s %3d pgs | %2d dims | %4d imgs | sem texto: %2d | fontes: %s"
                  % (rot, info["n_paginas"], len(r["dimensoes_pagina"]), r["total_imagens"],
                     len(r["paginas_sem_texto"]),
                     ", ".join(f for f, _ in r["fontes"][:3])))
    print("OK ->", DEST)


if __name__ == "__main__":
    main()
