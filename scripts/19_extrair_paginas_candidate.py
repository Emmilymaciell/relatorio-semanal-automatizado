# -*- coding: utf-8 -*-
r"""
scripts/19_extrair_paginas_candidate.py - Fase 3.3B/3.4: renderiza um candidate
MONTADO a 150 dpi e extrai o texto de cada pagina, para permitir a comparacao
exaustiva pagina a pagina contra o golden (35 paginas), cujas imagens e texto ja
existem em output/visual_review/golden/ e output/visual_review/paginas_texto.json.

Nasceu como 19_comparacao_exaustiva_v0_1.py, com a v0.1 fixa no codigo. Virou
parametrico na v0.2: a rodada de convergencia compara VARIAS versoes contra o
mesmo golden, e sobrescrever a extracao anterior destruiria justamente o termo
de comparacao v0.1 x v0.2.

Cada versao grava numa chave propria de paginas_texto.json ("shadow_v0.1",
"shadow_v0.2", ...). A chave legada "shadow" e' mantida apontando para a v0.1,
que e' o que os artefatos ja gravados (comparison_v0.1.json) assumem.

Nao le nada do SharePoint (opera sobre o PDF ja montado dentro do repositorio).
Nao decide classificacao -- so produz os insumos mecanicos (imagem + texto).

Uso:
    python scripts/19_extrair_paginas_candidate.py --versao v0.2
"""
import os
import json
import argparse

import fitz  # PyMuPDF

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)

SEMANA = os.path.join(RAIZ, "output", "shadow", "RD103_2026-08-29_2026-09-04")
TEXTO_JSON = os.path.join(RAIZ, "output", "visual_review", "paginas_texto.json")

DPI = 150

# a chave "shadow" e' anterior a existirem varias versoes; aponta para a v0.1
VERSAO_DA_CHAVE_LEGADA = "v0.1"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--versao", default="v0.2")
    ap.add_argument("--dpi", type=int, default=DPI)
    args = ap.parse_args()

    candidate_pdf = os.path.join(
        SEMANA, "shadow_final_candidate_%s_MONTADO.pdf" % args.versao)
    if not os.path.exists(candidate_pdf):
        print("ERRO: nao encontrei %s -- rode scripts/montar_pdf_final.py antes"
              % candidate_pdf)
        return 2

    out_dir = os.path.join(RAIZ, "output", "visual_review", "candidate_%s" % args.versao)
    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(candidate_pdf)

    zoom = args.dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    textos = {}
    formatos = {}
    for i, page in enumerate(doc):
        n = i + 1
        pix = page.get_pixmap(matrix=mat)
        pix.save(os.path.join(out_dir, "page_%03d.png" % n))

        textos[str(n)] = page.get_text("text")

        r = page.rect
        formatos[str(n)] = {
            "largura_pt": round(r.width, 1),
            "altura_pt": round(r.height, 1),
            "largura_mm": round(r.width * 25.4 / 72, 1),
            "altura_mm": round(r.height * 25.4 / 72, 1),
            "orientacao": "paisagem" if r.width > r.height else "retrato",
        }
        print("pagina %2d: %6.1f x %6.1f mm (%s)" % (
            n, formatos[str(n)]["largura_mm"], formatos[str(n)]["altura_mm"],
            formatos[str(n)]["orientacao"]))

    # preserva o golden e as OUTRAS versoes ja extraidas
    if os.path.exists(TEXTO_JSON):
        with open(TEXTO_JSON, "r", encoding="utf-8") as f:
            dados = json.load(f)
    else:
        dados = {"golden": {}}
    dados["shadow_%s" % args.versao] = textos
    if args.versao == VERSAO_DA_CHAVE_LEGADA:
        dados["shadow"] = textos

    with open(TEXTO_JSON, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

    with open(os.path.join(out_dir, "formatos_paginas.json"), "w", encoding="utf-8") as f:
        json.dump(formatos, f, ensure_ascii=False, indent=2)

    print("\nOK: %d paginas renderizadas em %s" % (doc.page_count, out_dir))
    print("Texto gravado em %s (chave 'shadow_%s')" % (TEXTO_JSON, args.versao))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
