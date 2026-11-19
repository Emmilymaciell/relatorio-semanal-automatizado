# -*- coding: utf-8 -*-
r"""
15_p08_ficha_acomp_semanal.py - fecha P-08.

Le a 'Ficha de Acompanhamento Semanal (Obra)' (arquivos de 10-900 MB porque
embutem fotos) SEM extrair midia: abre o zip por streaming e le apenas o XML das
planilhas. Escreve so em output/evidencias/.
"""
import os
import re
import sys
import json
import zipfile
import argparse
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NSP = "http://schemas.openxmlformats.org/package/2006/relationships"


def q(ns, t):
    return "{%s}%s" % (ns, t)


def sst(z):
    out = []
    try:
        info = z.getinfo("xl/sharedStrings.xml")
    except KeyError:
        return out
    if info.file_size > 80 * 1024 * 1024:
        return out
    with z.open("xl/sharedStrings.xml") as fh:
        for _, el in ET.iterparse(fh, events=("end",)):
            if el.tag == q(NS, "si"):
                out.append("".join(t.text or "" for t in el.iter(q(NS, "t"))))
                el.clear()
    return out


def rels(z, parte):
    d = os.path.dirname(parte)
    cam = (d + "/" if d else "") + "_rels/" + os.path.basename(parte) + ".rels"
    try:
        r = ET.fromstring(z.read(cam))
    except (KeyError, ET.ParseError):
        return {}
    return {x.get("Id"): x.get("Target") for x in r.findall(q(NSP, "Relationship"))}


def abas(z):
    r = ET.fromstring(z.read("xl/workbook.xml"))
    rl = rels(z, "xl/workbook.xml")
    out = []
    for s in r.find(q(NS, "sheets")):
        alvo = rl.get(s.get(q(NSR, "id")), "")
        parte = alvo if alvo.startswith("xl/") else "xl/" + alvo.lstrip("/")
        out.append({"nome": s.get("name"), "estado": s.get("state") or "visible",
                    "parte": parte})
    return out


def celulas(z, parte, strings, limite=60000):
    out = {}
    try:
        z.getinfo(parte)
    except KeyError:
        return out
    with z.open(parte) as fh:
        for _, el in ET.iterparse(fh, events=("end",)):
            if el.tag != q(NS, "row"):
                continue
            for c in el.findall(q(NS, "c")):
                t = c.get("t")
                v = c.find(q(NS, "v"))
                i = c.find(q(NS, "is"))
                val = None
                if t == "s" and v is not None and v.text is not None:
                    try:
                        val = strings[int(v.text)]
                    except (ValueError, IndexError):
                        val = None
                elif t == "inlineStr" and i is not None:
                    val = "".join(x.text or "" for x in i.iter(q(NS, "t")))
                elif v is not None:
                    val = v.text
                if val is None:
                    continue
                val = re.sub(r"\s+", " ", str(val)).strip()
                if val:
                    out[c.get("r")] = val
                if len(out) >= limite:
                    break
            el.clear()
    return out


def parse_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref or "")
    return (m.group(1), int(m.group(2))) if m else (None, 0)


def dissecar(caminho, max_linhas=400):
    st = U.stat_info(caminho)
    res = {"arquivo": os.path.basename(caminho), "caminho": caminho,
           "tamanho_MB": round(st["tamanho_bytes"] / 1e6, 1),
           "modificado": st["modificado"], "abas": []}
    with zipfile.ZipFile(U.lp(caminho)) as z:
        nomes = z.namelist()
        res["n_midias"] = sum(1 for n in nomes if n.startswith("xl/media/"))
        strings = sst(z)
        res["n_shared_strings"] = len(strings)
        for a in abas(z):
            cel = celulas(z, a["parte"], strings)
            porlin = {}
            for ref, v in cel.items():
                col, lin = parse_ref(ref)
                if col:
                    porlin.setdefault(lin, []).append((col, v))
            a["n_celulas"] = len(cel)
            a["linhas"] = [{"l": L, "cels": sorted(porlin[L])}
                           for L in sorted(porlin)[:max_linhas]]
            res["abas"].append(a)
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("caminho")
    ap.add_argument("--saida", default=None)
    ap.add_argument("--max", type=int, default=400)
    args = ap.parse_args()

    d = dissecar(args.caminho, args.max)
    print("#" * 80)
    print(d["arquivo"], "|", d["tamanho_MB"], "MB |", d["n_midias"], "midias |",
          d["modificado"])
    for a in d["abas"]:
        print("=" * 76)
        print("ABA: %s (%s) - %d celulas" % (a["nome"], a["estado"], a["n_celulas"]))
        for ln in a["linhas"]:
            print("  L%-4d %s" % (ln["l"], " | ".join(
                "%s=%s" % (c, v[:70]) for c, v in ln["cels"])))

    saida = args.saida or os.path.join(
        U.OUT, "evidencias", "P08_" + U.slug(d["arquivo"])[:60] + ".json")
    U.assert_read_only(saida)
    os.makedirs(os.path.dirname(saida), exist_ok=True)
    with open(saida, "w", encoding="utf-8") as fh:
        json.dump(d, fh, ensure_ascii=False, indent=1)
    print("\nOK ->", saida)


if __name__ == "__main__":
    main()
