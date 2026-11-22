# -*- coding: utf-8 -*-
r"""
11_amostrar_overlay_fotos.py - amostragem REPRODUTIVEL de fotografias de obra
para leitura visual do OVERLAY (metadata renderizada sobre a imagem).

CONTEXTO
--------
As fotos NAO tem EXIF util (D-30). Mas carregam metadata VISUAL "queimada" na
imagem: data, zona UTM, coordenadas E/N, estaca inicial/final, descricao do
servico, rodovia e "Numero de indice".

Distinguir sempre tres coisas:
  (1) metadata de ARQUIVO  (nome, mtime, tamanho)
  (2) EXIF                 (vazio nestas fotos)
  (3) metadata VISUAL      (overlay desenhado sobre os pixels)  <-- alvo deste script

ESTE SCRIPT NAO FAZ OCR. Ele apenas SELECIONA e COPIA amostras para tmp/,
de forma deterministica, para que sejam lidas com a ferramenta Read (visual).

USO
---
    python scripts/11_amostrar_overlay_fotos.py            # amostra padrao
    python scripts/11_amostrar_overlay_fotos.py --por-servico 3
    python scripts/11_amostrar_overlay_fotos.py --obra RD103 --semana 2026-08-22

SAIDAS
------
    tmp/overlay_fotos/<OBRA>/<AAAA-MM-DD>/<servico>__<nn>__<arquivo>.jpeg
    output/evidencias/overlay_fotos_amostra.json   (manifesto da amostra)

REGRAS DO PROJETO
-----------------
* SharePoint READ-ONLY: so leitura; toda escrita via prog_util.copiar_para_tmp.
* MAX_PATH: todo acesso passa por prog_util.lp()/listdir/isdir/stat_info.
"""
import os
import re
import sys
import json
import argparse
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

SUBPASTA_FOTOS = "3. Fotos"
DEST_SUB = "overlay_fotos"
MANIFESTO = os.path.join(U.OUT, "evidencias", "overlay_fotos_amostra.json")

# Semanas da amostra do projeto (regime sabado->sexta desde 2026-01-03)
# + uma semana antiga para comparar evolucao historica.
SEMANAS_PADRAO = [
    "2026-08-22",   # 2026-08-22 .. 2026-08-28
    "2026-08-29",   # 2026-08-29 .. 2026-09-04
    "2025-11-08",   # historica  (regime sexta->quinta)
]

EXT_IMG = (".jpeg", ".jpg", ".png", ".heic", ".webp")

# Os nomes de pasta-dia tem DOIS regimes de grafia:
#   AA_MM_DD  (sublinhado)  -> ate ~fev/2026
#   AA-MM-DD  (hifen)       -> a partir de ~mar/2026
_RE_DIA = re.compile(r"^(\d{2})[-_](\d{2})[-_](\d{2})$")


def data_da_pasta(nome):
    """'26-08-24' ou '26_08_24' -> date(2026, 8, 24). Devolve None se nao casar."""
    m = _RE_DIA.match(nome.strip())
    if not m:
        return None
    aa, mm, dd = (int(x) for x in m.groups())
    try:
        return date(2000 + aa, mm, dd)
    except ValueError:
        return None


def pasta_fotos(obra):
    raiz = U.pasta_semanal_raiz(obra)
    if not raiz:
        return None
    for d in U.listdir(raiz):
        if U.norm(d).startswith("3. foto"):
            return os.path.join(raiz, d)
    return None


def dias_da_obra(obra):
    """[(date, nome_pasta, caminho)] ordenado por data real, nao por string."""
    base = pasta_fotos(obra)
    if not base:
        return []
    out = []
    for nome in U.listdir(base):
        p = os.path.join(base, nome)
        if not U.isdir(p):
            continue
        d = data_da_pasta(nome)
        if d:
            out.append((d, nome, p))
    return sorted(out, key=lambda t: t[0])


def itens_do_dia(caminho_dia):
    """[(servico, arquivo, caminho)] - servico='(raiz)' quando a foto esta solta."""
    itens = []
    for nome in U.listdir(caminho_dia):
        p = os.path.join(caminho_dia, nome)
        if U.isdir(p):
            for arq in U.listdir(p):
                if arq.lower().endswith(EXT_IMG):
                    itens.append((nome, arq, os.path.join(p, arq)))
        elif nome.lower().endswith(EXT_IMG):
            itens.append(("(raiz)", nome, p))
    return itens


def selecionar(obra, inicio_semana, por_servico):
    """Selecao deterministica: para cada dia da semana, as `por_servico`
    primeiras fotos (ordem alfabetica) de cada pasta de servico."""
    fim = inicio_semana + timedelta(days=6)
    sel = []
    for d, nome_pasta, caminho in dias_da_obra(obra):
        if not (inicio_semana <= d <= fim):
            continue
        porsvc = {}
        for servico, arq, p in itens_do_dia(caminho):
            porsvc.setdefault(servico, []).append((arq, p))
        for servico in sorted(porsvc):
            lista = sorted(porsvc[servico])
            for i, (arq, p) in enumerate(lista[:por_servico], 1):
                sel.append({
                    "obra": obra,
                    "rodovia": U.RODOVIA_LABEL[obra],
                    "semana_inicio": inicio_semana.isoformat(),
                    "data_pasta": d.isoformat(),
                    "pasta_dia_grafia": nome_pasta,
                    "servico_pasta": servico,
                    "arquivo": arq,
                    "origem": p,
                    "ordem_no_servico": i,
                    "total_no_servico": len(lista),
                })
    return sel


def main():
    ap = argparse.ArgumentParser(description="Amostra fotos para leitura do overlay visual.")
    ap.add_argument("--obra", action="append", choices=U.OBRAS,
                    help="restringe a obra(s); default: todas")
    ap.add_argument("--semana", action="append",
                    help="inicio da semana AAAA-MM-DD; default: %s" % ", ".join(SEMANAS_PADRAO))
    ap.add_argument("--por-servico", type=int, default=2,
                    help="quantas fotos por pasta de servico (default 2)")
    ap.add_argument("--copiar", dest="copiar", action="store_true", default=True)
    ap.add_argument("--nao-copiar", dest="copiar", action="store_false",
                    help="so gera o manifesto, sem copiar para tmp/")
    args = ap.parse_args()

    obras = args.obra or U.OBRAS
    semanas = [date.fromisoformat(s) for s in (args.semana or SEMANAS_PADRAO)]

    manifesto = {"gerado_em": None, "por_servico": args.por_servico, "amostras": []}
    from datetime import datetime, timezone
    manifesto["gerado_em"] = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")

    for obra in obras:
        for ini in semanas:
            sel = selecionar(obra, ini, args.por_servico)
            for n, item in enumerate(sel, 1):
                st = U.stat_info(item["origem"])
                item["tamanho_bytes"] = st["tamanho_bytes"]
                item["mtime"] = st["modificado"]
                if args.copiar:
                    subdir = os.path.join(DEST_SUB, obra, item["data_pasta"])
                    nome = "%s__%02d__%s" % (U.slug(item["servico_pasta"]), n, item["arquivo"])
                    item["copia_tmp"] = U.copiar_para_tmp(item["origem"], subdir, nome)
                manifesto["amostras"].append(item)
            print("%-7s semana %s -> %d fotos" % (obra, ini.isoformat(), len(sel)))

    U.assert_read_only(MANIFESTO)
    os.makedirs(os.path.dirname(MANIFESTO), exist_ok=True)
    with open(MANIFESTO, "w", encoding="utf-8") as fh:
        json.dump(manifesto, fh, ensure_ascii=False, indent=2)
    print("\nTotal: %d fotos amostradas" % len(manifesto["amostras"]))
    print("Manifesto: %s" % MANIFESTO)
    if args.copiar:
        print("Copias:    %s" % os.path.join(U.TMP, DEST_SUB))


if __name__ == "__main__":
    main()
