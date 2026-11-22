# -*- coding: utf-8 -*-
"""
07_estrutura_e_diff.py

A partir da extracao dos PDFs (03_extrair_pdf.py):
  1. detecta a estrutura de secoes de cada relatorio;
  2. classifica cada pagina por formato de papel (A4/A3/A2/A0) -> identifica anexos;
  3. cataloga legendas de figuras/graficos/tabelas;
  4. faz o diff semantico semana x semana (mesma obra) e obra x obra (mesma semana).

Saida: output/comparacoes/*.md e *.json
"""
import os
import sys
import re
import json
import glob
import difflib
from collections import Counter, OrderedDict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

SRC = os.path.join(U.OUT, "extracao", "pdf")
DEST = os.path.join(U.OUT, "comparacoes")

# formatos ISO em pontos (1 pt = 1/72 pol), com tolerancia
FORMATOS = [
    ("A4 retrato",   595, 842), ("A4 paisagem",  842, 595),
    ("A3 retrato",   842, 1191), ("A3 paisagem", 1191, 842),
    ("A2 retrato",  1191, 1684), ("A2 paisagem", 1684, 1191),
    ("A1 retrato",  1684, 2384), ("A1 paisagem", 2384, 1684),
    ("A0 retrato",  2384, 3370), ("A0 paisagem", 3370, 2384),
]

RE_SEC = re.compile(r"^\s*(\d{1,2}(?:\.\d{1,2}){0,2})\.?\s*[\-–]?\s*(.+?)\s*$")
RE_TOC = re.compile(r"\.{4,}\s*\d+\s*$")
RE_LEGENDA = re.compile(r"^\s*(Figura|Gr[aá]fico|Tabela|Quadro|Foto(?:grafia)?)\s+(\d+)\s*[\-–:]\s*(.+?)\s*$",
                        re.IGNORECASE)
RE_ANEXO = re.compile(r"^\s*Anexo\s+([IVXLC]+)\s*[\-–]\s*(.+?)\s*$", re.IGNORECASE)


def formato(w, h, tol=12):
    for nome, fw, fh in FORMATOS:
        if abs(w - fw) <= tol and abs(h - fh) <= tol:
            return nome
    return "%dx%d pt" % (round(w), round(h))


def linhas_da_pagina(p):
    out = []
    for b in p["blocos"]:
        for ln in b["linhas"]:
            t = "".join(s["t"] for s in ln)
            if t.strip():
                out.append({"t": t.rstrip(), "size": max(s["size"] for s in ln),
                            "font": ln[0]["font"], "bbox": ln[0]["bbox"],
                            "bold": any(s["flags"] & 16 for s in ln)})
    return out


def analisar(info):
    """Estrutura de um relatorio: capa, indice, secoes, paginas, anexos, legendas."""
    r = {"obra": info["obra"], "numero": info["numero"],
         "data_inicio": info["data_inicio"], "data_fim": info["data_fim"],
         "n_paginas": info["n_paginas"], "arquivo": info["arquivo"]}

    # --- paginas por formato ---
    pgs = []
    for p in info["paginas"]:
        pgs.append({"n": p["n"], "formato": formato(p["largura_pt"], p["altura_pt"]),
                    "imgs": len(p["imagens"]), "linhas": linhas_da_pagina(p)})
    r["formatos"] = Counter(p["formato"] for p in pgs).most_common()

    # --- capa (pagina 1): pares ROTULO: valor ---
    capa = OrderedDict()
    txt_capa = [l["t"].strip() for l in pgs[0]["linhas"]]
    for t in txt_capa:
        m = re.match(r"^([A-ZÀ-Ú][A-ZÀ-Ú º°/\.]{2,40}):\s*(.*)$", t)
        if m:
            capa[m.group(1).strip()] = m.group(2).strip()
    r["capa_campos"] = capa
    r["capa_linhas"] = txt_capa

    # --- indice (pagina 2) ---
    toc = []
    for l in pgs[1]["linhas"]:
        t = l["t"].strip()
        if RE_TOC.search(t):
            t2 = re.sub(r"\.{2,}\s*(\d+)\s*$", r"|\1", t)
            partes = t2.rsplit("|", 1)
            titulo = partes[0].strip()
            pagina = int(partes[1]) if len(partes) > 1 and partes[1].strip().isdigit() else None
            m = RE_SEC.match(titulo)
            num = m.group(1) if m else None
            tit = m.group(2).strip() if m else titulo
            toc.append({"num": num, "titulo": tit, "pagina": pagina})
    r["indice"] = toc

    # --- titulos reais no corpo (tamanho de fonte maior + numeracao) ---
    corpo = []
    for p in pgs[2:]:
        if p["formato"] != "A4 retrato":
            continue
        for l in p["linhas"]:
            t = l["t"].strip()
            m = RE_SEC.match(t)
            if m and l["size"] >= 12 and len(t) < 90 and not RE_TOC.search(t):
                corpo.append({"num": m.group(1), "titulo": m.group(2).strip(),
                              "pagina": p["n"], "size": l["size"]})
    r["secoes_corpo"] = corpo

    # --- legendas ---
    leg = []
    for p in pgs:
        for l in p["linhas"]:
            m = RE_LEGENDA.match(l["t"].strip())
            if m:
                leg.append({"tipo": m.group(1).capitalize(), "n": int(m.group(2)),
                            "texto": m.group(3).strip(), "pagina": p["n"]})
    r["legendas"] = leg
    r["legendas_por_tipo"] = Counter(x["tipo"] for x in leg).most_common()

    # --- anexos ---
    anexos = []
    pagina_anexos = None
    for p in pgs:
        for l in p["linhas"]:
            m = RE_ANEXO.match(l["t"].strip())
            if m:
                anexos.append({"ordem": m.group(1), "titulo": m.group(2).strip(), "pagina": p["n"]})
                pagina_anexos = p["n"]
    r["anexos"] = anexos
    r["pagina_lista_anexos"] = pagina_anexos

    # blocos de paginas de anexo (tudo depois da pagina da lista)
    if pagina_anexos:
        blocos, atual = [], None
        for p in pgs:
            if p["n"] <= pagina_anexos:
                continue
            if atual is None or p["formato"] != atual["formato"]:
                atual = {"formato": p["formato"], "de": p["n"], "ate": p["n"], "imgs": p["imgs"]}
                blocos.append(atual)
            else:
                atual["ate"] = p["n"]
                atual["imgs"] += p["imgs"]
        r["blocos_anexo"] = blocos

    # --- texto corrido por secao (para diff) ---
    texto_secoes = OrderedDict()
    atual = "0. CAPA"
    for p in pgs:
        if p["formato"] != "A4 retrato":
            continue
        for l in p["linhas"]:
            t = l["t"].strip()
            if not t:
                continue
            m = RE_SEC.match(t)
            if m and l["size"] >= 12 and len(t) < 90 and not RE_TOC.search(t) and p["n"] > 2:
                atual = "%s %s" % (m.group(1), m.group(2).strip())
                texto_secoes.setdefault(atual, [])
                continue
            texto_secoes.setdefault(atual, []).append(t)
    r["texto_secoes"] = {k: "\n".join(v) for k, v in texto_secoes.items()}
    return r


def normalizar_para_diff(txt):
    """Remove numeros de pagina soltos e normaliza espacos."""
    linhas = []
    for l in txt.split("\n"):
        l = l.strip()
        if not l or re.fullmatch(r"\d{1,3}", l):
            continue
        linhas.append(re.sub(r"\s+", " ", l))
    return linhas


def diff_textos(a, b, rot_a, rot_b):
    out = []
    la, lb = normalizar_para_diff(a), normalizar_para_diff(b)
    sm = difflib.SequenceMatcher(None, la, lb, autojunk=False)
    ratio = sm.ratio()
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        for l in la[i1:i2]:
            out.append(("-", l))
        for l in lb[j1:j2]:
            out.append(("+", l))
    return ratio, out


def main():
    os.makedirs(U.assert_read_only(DEST), exist_ok=True)
    rels = {}
    for f in sorted(glob.glob(os.path.join(SRC, "*.json"))):
        info = json.load(open(f, encoding="utf-8"))
        r = analisar(info)
        rels["%s_%03d" % (r["obra"], r["numero"])] = r

    with open(os.path.join(DEST, "estrutura_relatorios.json"), "w", encoding="utf-8") as fh:
        json.dump(rels, fh, ensure_ascii=False, indent=1)

    # ---------- relatorio em Markdown ----------
    L = []
    L.append("# Estrutura detectada e diff semântico dos 8 relatórios\n")
    L.append("*Gerado por `scripts/07_estrutura_e_diff.py` a partir da extração dos PDFs.*\n")

    L.append("\n## 1. Formatos de página por relatório\n")
    L.append("| Relatório | Págs | " + " | ".join(
        sorted({f for r in rels.values() for f, _ in r["formatos"]})) + " |")
    todos_f = sorted({f for r in rels.values() for f, _ in r["formatos"]})
    L.append("|" + "---|" * (len(todos_f) + 2))
    for k, r in rels.items():
        d = dict(r["formatos"])
        L.append("| %s | %d | %s |" % (k, r["n_paginas"],
                                       " | ".join(str(d.get(f, "")) for f in todos_f)))

    L.append("\n## 2. Índice (TOC) por relatório\n")
    for k, r in rels.items():
        L.append("\n### %s — %s a %s (%d págs)\n" % (k, r["data_inicio"], r["data_fim"], r["n_paginas"]))
        for t in r["indice"]:
            L.append("- `%s` %s → p.%s" % (t["num"], t["titulo"], t["pagina"]))

    L.append("\n## 3. Anexos por relatório\n")
    for k, r in rels.items():
        L.append("\n### %s\n" % k)
        for a in r["anexos"]:
            L.append("- Anexo %s — %s" % (a["ordem"], a["titulo"]))
        if r.get("blocos_anexo"):
            L.append("\nBlocos de páginas após a lista:\n")
            for b in r["blocos_anexo"]:
                L.append("  - p.%d–%d — %s (%d imagens)" % (b["de"], b["ate"], b["formato"], b["imgs"]))

    L.append("\n## 4. Capa — campos detectados\n")
    chaves = []
    for r in rels.values():
        for c in r["capa_campos"]:
            if c not in chaves:
                chaves.append(c)
    L.append("| Campo | " + " | ".join(rels.keys()) + " |")
    L.append("|" + "---|" * (len(rels) + 1))
    for c in chaves:
        L.append("| **%s** | %s |" % (c, " | ".join(
            (r["capa_campos"].get(c, "—") or "—")[:40] for r in rels.values())))

    L.append("\n## 5. Legendas (figuras, gráficos, tabelas)\n")
    L.append("| Relatório | " + " | ".join(
        sorted({t for r in rels.values() for t, _ in r["legendas_por_tipo"]})) + " |")
    tipos = sorted({t for r in rels.values() for t, _ in r["legendas_por_tipo"]})
    L.append("|" + "---|" * (len(tipos) + 1))
    for k, r in rels.items():
        d = dict(r["legendas_por_tipo"])
        L.append("| %s | %s |" % (k, " | ".join(str(d.get(t, 0)) for t in tipos)))

    # ---------- diffs ----------
    L.append("\n## 6. Diff semana × semana (mesma obra)\n")
    pares_semana = []
    for obra in U.OBRAS:
        ks = sorted([k for k in rels if k.startswith(obra + "_")])
        if len(ks) == 2:
            pares_semana.append((ks[0], ks[1]))

    diffs = {}
    for a, b in pares_semana:
        ra, rb = rels[a], rels[b]
        L.append("\n### %s → %s\n" % (a, b))
        secs = list(OrderedDict.fromkeys(list(ra["texto_secoes"]) + list(rb["texto_secoes"])))
        for s in secs:
            ta, tb = ra["texto_secoes"].get(s, ""), rb["texto_secoes"].get(s, "")
            if not ta and not tb:
                continue
            if not ta:
                L.append("- **%s** — SÓ em %s" % (s, b)); continue
            if not tb:
                L.append("- **%s** — SÓ em %s" % (s, a)); continue
            ratio, mud = diff_textos(ta, tb, a, b)
            marca = "IDÊNTICO" if ratio == 1.0 else ("%.0f%% igual" % (ratio * 100))
            L.append("\n#### %s — %s" % (s, marca))
            if mud and ratio < 1.0:
                for sig, l in mud[:40]:
                    L.append("    %s %s" % (sig, l[:160]))
                if len(mud) > 40:
                    L.append("    ... (%d linhas a mais)" % (len(mud) - 40))
        diffs["%s->%s" % (a, b)] = True

    with open(os.path.join(DEST, "estrutura_e_diff.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(L))

    print("Relatorios analisados:", len(rels))
    for k, r in rels.items():
        print("  %-12s %2d pgs | %2d secoes TOC | %d anexos | %d legendas | formatos: %s"
              % (k, r["n_paginas"], len(r["indice"]), len(r["anexos"]), len(r["legendas"]),
                 ", ".join("%s(%d)" % (f, n) for f, n in r["formatos"])))
    print("OK ->", DEST)


if __name__ == "__main__":
    main()
