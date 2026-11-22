# -*- coding: utf-8 -*-
r"""
tmp_extrair_paineis_semana.py - script TEMPORARIO (Tarefa 1 do briefing).

Copia os 8 relatorios (4 obras x 2 semanas: fim 2026-08-28 e 2026-09-04) do
SharePoint (READ-ONLY) para tmp/, renderiza as paginas da secao 6
(ACOMPANHAMENTO AMBIENTAL) com fitz e tambem extrai os xrefs de imagem ja
conhecidos (de scripts/12_analisar_ambiental.py::PAINEIS_LIDOS) como PNG,
para inspecao visual direta (Read tool) e decidir se o print e filtrado por
obra ou mostra os totais corporativos.
"""
import os
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U
import fitz

OUT_DIR = os.path.join(U.PROJ, "tmp", "amb_paineis_semana")
os.makedirs(OUT_DIR, exist_ok=True)

with open(os.path.join(U.OUT, "inventarios", "relatorios_semanais_historico.json"),
          encoding="utf-8") as f:
    HIST = json.load(f)["relatorios"]

ALVO_FIM = {"2026-08-28", "2026-09-04"}

# xref/pagina ja conhecidos de scripts/12_analisar_ambiental.py::PAINEIS_LIDOS
CONHECIDOS = {
    ("RD101", "2026-08-28"): {"indices": (13, 1527), "condicionantes": (14, 1539)},
    ("RD101", "2026-09-04"): {"indices": (13, 1553), "condicionantes": (14, 1566)},
    ("RD102A", "2026-08-28"): {"indices": (11, 835), "condicionantes": (12, 855)},
    ("RD102A", "2026-09-04"): {"indices": (11, 830), "condicionantes": (12, 842)},
    ("RD103", "2026-08-28"): {"indices": (11, 736), "condicionantes": (12, 747)},
    ("RD103", "2026-09-04"): {"indices": (11, 749), "condicionantes": (12, 761)},
    ("RD104", "2026-08-28"): {"indices": (11, 665), "condicionantes": (12, 681)},
    ("RD104", "2026-09-04"): {"indices": (10, 630), "condicionantes": (11, 642)},
}


def main():
    alvos = [r for r in HIST if r["pasta_data_fim"] in ALVO_FIM]
    print("Relatorios alvo:", len(alvos))
    for r in alvos:
        obra = r["obra"]
        fim = r["pasta_data_fim"]
        chave = (obra, fim)
        if chave not in CONHECIDOS:
            print("  [sem xref conhecido]", obra, fim)
            continue
        origem = r["caminho"]
        nome_local = "%s_%s.pdf" % (obra, fim)
        local = U.copiar_para_tmp(origem, "amb_paineis_semana_pdf", nome_local)
        print("  copiado:", local)

        doc = fitz.open(local)
        for tipo, (pagina, xref) in CONHECIDOS[chave].items():
            pg_idx = pagina - 1  # 1-based no inventario -> 0-based no fitz
            # 1) extrai a imagem raster bruta pelo xref conhecido (mais fiel)
            try:
                info = doc.extract_image(xref)
                ext = info["ext"]
                raw_path = os.path.join(OUT_DIR, "%s_%s_%s_raw_xref%d.%s" %
                                        (obra, fim, tipo, xref, ext))
                with open(raw_path, "wb") as fo:
                    fo.write(info["image"])
                print("    raw:", raw_path, info.get("width"), info.get("height"))
            except Exception as e:
                print("    ERRO extract_image", obra, fim, tipo, xref, repr(e))
            # 2) renderiza a pagina inteira tambem, para contexto (titulo,
            #    posicao do slicer, layout ao redor da imagem)
            try:
                page = doc[pg_idx]
                pix = page.get_pixmap(dpi=150)
                pg_path = os.path.join(OUT_DIR, "%s_%s_%s_pagina%d.png" %
                                       (obra, fim, tipo, pagina))
                pix.save(pg_path)
                print("    pagina:", pg_path)
            except Exception as e:
                print("    ERRO render pagina", obra, fim, tipo, pagina, repr(e))
        doc.close()


if __name__ == "__main__":
    main()
