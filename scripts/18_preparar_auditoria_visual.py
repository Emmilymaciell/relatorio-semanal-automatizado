# -*- coding: utf-8 -*-
r"""
18_preparar_auditoria_visual.py - frente MECANICA da auditoria visual.

Renderiza as paginas do golden e do shadow em PNG na MESMA resolucao, extrai as
medidas exatas do DOCX golden (sectPr, styles, header/footer, numbering) e monta
o PAGE_ALIGNMENT_MAP.

NAO altera nada no SharePoint. O golden e copiado para tmp/ antes de ser aberto.
NAO implementa correcao nenhuma no renderer.

Uso:
    python -X utf8 scripts/18_preparar_auditoria_visual.py --dpi 150
"""
import os
import re
import sys
import json
import zipfile
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

import fitz  # PyMuPDF

OBRA = "RD103"
INI, FIM = "2026-08-29", "2026-09-04"
SHADOW_DIR = os.path.join(U.OUT, "shadow", "%s_%s_%s" % (OBRA, INI, FIM))
REVIEW_DIR = os.path.join(U.OUT, "visual_review")

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TWIP = 1 / 20.0          # 1 twip = 1/20 pt
PT_MM = 25.4 / 72.0


def golden_paths():
    """Caminhos do golden no SharePoint (READ-ONLY)."""
    pasta = U.pasta_relatorios(OBRA)
    for d in U.listdir(pasta):
        if d.startswith("2026.08.29-2026.09.04"):
            raiz = os.path.join(pasta, d)
            arqs = U.listdir(raiz)
            pdf = next((a for a in arqs if a.lower().endswith(".pdf")), None)
            docx = next((a for a in arqs if a.lower().endswith(".docx")), None)
            return (os.path.join(raiz, pdf) if pdf else None,
                    os.path.join(raiz, docx) if docx else None)
    return None, None


# -----------------------------------------------------------------------------
# 1. RENDERIZACAO
# -----------------------------------------------------------------------------
def render(pdf_path, destino, dpi):
    U.assert_read_only(destino)
    os.makedirs(destino, exist_ok=True)
    doc = fitz.open(U.lp(pdf_path))
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    paginas = []
    for i, page in enumerate(doc, 1):
        r = page.rect
        largura_mm, altura_mm = r.width * PT_MM, r.height * PT_MM
        nome = "page_%03d.png" % i
        page.get_pixmap(matrix=mat).save(os.path.join(destino, nome))
        paginas.append({
            "n": i, "arquivo": nome,
            "largura_pt": round(r.width, 1), "altura_pt": round(r.height, 1),
            "largura_mm": round(largura_mm, 1), "altura_mm": round(altura_mm, 1),
            "orientacao": "paisagem" if r.width > r.height else "retrato",
            "papel": papel(largura_mm, altura_mm),
            "rotacao": page.rotation,
            "n_imagens": len(page.get_images(full=True)),
            "texto": page.get_text().strip(),
        })
    doc.close()
    return paginas


PAPEIS = {"A4": (210, 297), "A3": (297, 420), "A2": (420, 594),
          "A1": (594, 841), "A0": (841, 1189)}


def papel(lmm, amm):
    a, b = sorted((lmm, amm))
    for nome, (pa, pb) in PAPEIS.items():
        if abs(a - pa) < 6 and abs(b - pb) < 6:
            return nome
    return "?"


# -----------------------------------------------------------------------------
# 2. MEDIDAS EXATAS DO DOCX GOLDEN
# -----------------------------------------------------------------------------
def ler_docx(caminho_docx):
    """Le o OOXML cru do golden. Copia para tmp/ antes (SharePoint e READ-ONLY)."""
    local = U.copiar_para_tmp(caminho_docx, "auditoria_visual")
    z = zipfile.ZipFile(local)
    partes = {n: z.read(n).decode("utf-8", "replace")
              for n in z.namelist()
              if n.endswith(".xml") and (n.startswith("word/") or n == "docProps/app.xml")}
    z.close()
    return local, partes


def dissecar_local(caminho_docx):
    """Mesma dissecacao, para um .docx que JA esta no projeto (nao copia)."""
    z = zipfile.ZipFile(caminho_docx)
    partes = {n: z.read(n).decode("utf-8", "replace")
              for n in z.namelist() if n.endswith(".xml") and n.startswith("word/")}
    z.close()
    doc = partes.get("word/document.xml", "")
    st = partes.get("word/styles.xml", "")
    padrao, sts = estilos(st)
    return {
        "sectPr": sectprs(doc),
        "docDefaults": padrao,
        "estilos": {k: v for k, v in sts.items() if v.get("tipo") == "paragraph"},
        "paragrafos": paragrafos_usados(doc),
        "tabelas": tabelas(doc),
        "headers": {n: {"texto": " ".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", partes[n]))[:300]}
                    for n in partes if re.match(r"word/header\d+\.xml", n)},
        "footers": {n: {"texto": " ".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", partes[n]))[:300],
                        "tem_campo_pagina": "PAGE" in partes[n]}
                    for n in partes if re.match(r"word/footer\d+\.xml", n)},
        "numbering_presente": "word/numbering.xml" in partes,
        "tem_header_footer": any(re.match(r"word/(header|footer)\d+\.xml", n) for n in partes),
    }


def _attr(tag, nome):
    m = re.search(r'%s%s="([^"]*)"' % (W.replace("{", "").replace("}", "") and "w:", nome), tag)
    return m.group(1) if m else None


def sectprs(document_xml):
    """Todos os <w:sectPr>: papel, orientacao, margens, header/footer distance."""
    out = []
    for m in re.finditer(r"<w:sectPr\b.*?</w:sectPr>|<w:sectPr\b[^>]*/>", document_xml, re.S):
        bloco = m.group(0)
        pg = re.search(r"<w:pgSz\b([^/>]*)/?>", bloco)
        mg = re.search(r"<w:pgMar\b([^/>]*)/?>", bloco)
        d = {"raw_len": len(bloco)}
        if pg:
            a = pg.group(1)
            d["pgSz_w_twips"] = _num(a, "w")
            d["pgSz_h_twips"] = _num(a, "h")
            d["orient"] = _str(a, "orient") or "portrait"
            if d["pgSz_w_twips"]:
                d["largura_mm"] = round(d["pgSz_w_twips"] * TWIP * PT_MM, 1)
                d["altura_mm"] = round(d["pgSz_h_twips"] * TWIP * PT_MM, 1)
                d["papel"] = papel(d["largura_mm"], d["altura_mm"])
        if mg:
            a = mg.group(1)
            for k in ("top", "right", "bottom", "left", "header", "footer", "gutter"):
                v = _num(a, k)
                if v is not None:
                    d["margem_%s_twips" % k] = v
                    d["margem_%s_mm" % k] = round(v * TWIP * PT_MM, 1)
        d["titlePg"] = "<w:titlePg" in bloco
        d["refs"] = re.findall(r'<w:(?:header|footer)Reference[^>]*w:type="(\w+)"[^>]*r:id="(\w+)"', bloco)
        out.append(d)
    return out


def _num(attrs, nome):
    m = re.search(r'w:%s="(-?\d+)"' % nome, attrs)
    return int(m.group(1)) if m else None


def _str(attrs, nome):
    m = re.search(r'w:%s="([^"]+)"' % nome, attrs)
    return m.group(1) if m else None


def estilos(styles_xml):
    """Estilos nomeados com fonte, corpo, espacamento, alinhamento e indentacao."""
    out = {}
    padrao = {}
    m = re.search(r"<w:docDefaults\b.*?</w:docDefaults>", styles_xml, re.S)
    if m:
        padrao = _props(m.group(0))
        padrao["_id"] = "docDefaults"
    for m in re.finditer(r"<w:style\b([^>]*)>(.*?)</w:style>", styles_xml, re.S):
        attrs, corpo = m.group(1), m.group(2)
        sid = _str(attrs, "styleId")
        tipo = _str(attrs, "type")
        nome = re.search(r'<w:name w:val="([^"]*)"', corpo)
        d = _props(corpo)
        d.update({"tipo": tipo, "nome": nome.group(1) if nome else sid,
                  "baseado_em": (re.search(r'<w:basedOn w:val="([^"]*)"', corpo) or [None, None])[1]
                  if re.search(r'<w:basedOn w:val="([^"]*)"', corpo) else None,
                  "default": _str(attrs, "default") == "1"})
        out[sid] = d
    return padrao, out


def _props(bloco):
    d = {}
    f = re.search(r'<w:rFonts\b([^/>]*)/?>', bloco)
    if f:
        for k in ("ascii", "hAnsi", "cs", "eastAsia"):
            v = _str(f.group(1), k)
            if v:
                d["fonte_%s" % k] = v
    sz = re.search(r'<w:sz w:val="(\d+)"', bloco)
    if sz:
        d["tamanho_pt"] = int(sz.group(1)) / 2.0
    if re.search(r"<w:b\b(?![^>]*w:val=\"0\")", bloco):
        d["negrito"] = True
    if re.search(r"<w:i\b(?![^>]*w:val=\"0\")", bloco):
        d["italico"] = True
    if re.search(r"<w:caps\b(?![^>]*w:val=\"0\")", bloco):
        d["caixa_alta"] = True
    c = re.search(r'<w:color w:val="([0-9A-Fa-f]{6}|auto)"', bloco)
    if c:
        d["cor"] = c.group(1)
    j = re.search(r'<w:jc w:val="(\w+)"', bloco)
    if j:
        d["alinhamento"] = j.group(1)
    sp = re.search(r"<w:spacing\b([^/>]*)/?>", bloco)
    if sp:
        a = sp.group(1)
        for k, rot in (("before", "espaco_antes"), ("after", "espaco_depois")):
            v = _num(a, k)
            if v is not None:
                d[rot + "_pt"] = round(v * TWIP, 1)
        ln = _num(a, "line")
        if ln:
            d["entrelinha"] = round(ln * TWIP, 1)
            d["entrelinha_regra"] = _str(a, "lineRule")
    ind = re.search(r"<w:ind\b([^/>]*)/?>", bloco)
    if ind:
        a = ind.group(1)
        for k, rot in (("left", "recuo_esq"), ("right", "recuo_dir"),
                       ("firstLine", "recuo_1a_linha"), ("hanging", "deslocamento")):
            v = _num(a, k)
            if v is not None:
                d[rot + "_mm"] = round(v * TWIP * PT_MM, 1)
    tabs = re.findall(r'<w:tab\b([^/>]*)/?>', bloco)
    ts = []
    for t in tabs:
        pos, val, lead = _num(t, "pos"), _str(t, "val"), _str(t, "leader")
        if pos is not None:
            ts.append({"pos_mm": round(pos * TWIP * PT_MM, 1), "tipo": val, "leader": lead})
    if ts:
        d["tabs"] = ts
    if re.search(r"<w:keepNext\b(?![^>]*w:val=\"0\")", bloco):
        d["keep_with_next"] = True
    if re.search(r"<w:keepLines\b(?![^>]*w:val=\"0\")", bloco):
        d["keep_together"] = True
    if re.search(r"<w:widowControl\b(?![^>]*w:val=\"0\")", bloco):
        d["widow_control"] = True
    return d


def paragrafos_usados(document_xml, limite=400):
    """Paragrafos na ordem, com estilo aplicado e overrides diretos."""
    out = []
    for m in re.finditer(r"<w:p\b[^>]*>(.*?)</w:p>", document_xml, re.S):
        corpo = m.group(1)
        texto = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", corpo))
        estilo = re.search(r'<w:pStyle w:val="([^"]*)"', corpo)
        ppr = re.search(r"<w:pPr>(.*?)</w:pPr>", corpo, re.S)
        d = {"texto": texto.strip()[:150],
             "estilo": estilo.group(1) if estilo else None,
             "overrides": _props(ppr.group(1)) if ppr else {}}
        rpr = re.search(r"<w:rPr>(.*?)</w:rPr>", corpo, re.S)
        if rpr:
            d["run"] = _props(rpr.group(1))
        if "<w:br w:type=\"page\"" in corpo:
            d["page_break"] = True
        if "<w:sectPr" in corpo:
            d["section_break"] = True
        if d["texto"] or d["estilo"] or d.get("page_break") or d.get("section_break"):
            out.append(d)
        if len(out) >= limite:
            break
    return out


def tabelas(document_xml):
    out = []
    for m in re.finditer(r"<w:tbl>(.*?)</w:tbl>", document_xml, re.S):
        corpo = m.group(1)
        grid = [int(x) for x in re.findall(r'<w:gridCol w:w="(\d+)"', corpo)]
        linhas = len(re.findall(r"<w:tr\b", corpo))
        larg = re.search(r'<w:tblW\b([^/>]*)/?>', corpo)
        bordas = re.findall(r'<w:(top|left|bottom|right|insideH|insideV)\b[^>]*w:sz="(\d+)"[^>]*w:color="([0-9A-Fa-f]{6}|auto)"', corpo)
        sh = re.findall(r'<w:shd\b[^>]*w:fill="([0-9A-Fa-f]{6}|auto)"', corpo)
        out.append({
            "linhas": linhas, "colunas": len(grid),
            "grid_twips": grid,
            "grid_mm": [round(g * TWIP * PT_MM, 1) for g in grid],
            "largura_total_mm": round(sum(grid) * TWIP * PT_MM, 1) if grid else None,
            "tblW": {"w": _num(larg.group(1), "w"), "type": _str(larg.group(1), "type")} if larg else None,
            "bordas": [{"lado": b[0], "sz_oitavos_pt": int(b[1]), "espessura_pt": int(b[1]) / 8.0,
                        "cor": b[2]} for b in bordas[:12]],
            "shading": list(dict.fromkeys(sh))[:8],
        })
    return out


# -----------------------------------------------------------------------------
# 3. PAGE ALIGNMENT MAP
# -----------------------------------------------------------------------------
ANCORAS = [
    ("capa", r"RELAT[ÓO]RIO SEMANAL DE"),
    ("indice", r"\b[ÍI]NDICE\b|SUM[ÁA]RIO"),
    ("s1_apresentacao", r"\bAPRESENTA[ÇC][ÃA]O\b"),
    ("s2_ficha", r"FICHA DE ACOMPANHAMENTO"),
    ("s3_avanco", r"AVAN[ÇC]O GLOBAL DA OBRA"),
    ("s4_unifilar", r"(DIAGRAMA )?UNIFILAR"),
    ("s5_execucao", r"EXECU[ÇC][ÃA]O DE OBRA"),
    ("s5_1_clima", r"(Dados|Condi[çc][õo]es) clim[áa]tic"),
    ("s5_2_fotografico", r"Relat[óo]rio Fotogr[áa]fico"),
    ("s6_ambiental", r"ACOMPANHAMENTO AMBIENTAL"),
    ("s7_diario", r"DI[ÁA]RIO DE OBRAS"),
    ("s8_nc", r"ACOMPANHAMENTO DE N[ÃA]O CONFORMIDADES"),
    ("s9_interferencias", r"RELATO DE (RISCOS E )?INTERFER[ÊE]NCIA"),
    ("s10_termo", r"TERMO DE ENCERRAMENTO"),
    ("s11_anexos", r"\bANEXOS\b"),
]


def secoes_por_pagina(paginas):
    """Em que pagina cada ancora de secao APARECE COMO TITULO (nao no indice)."""
    mapa = {}
    pag_indice = None
    for p in paginas:
        t = p["texto"]
        if pag_indice is None and re.search(ANCORAS[1][1], t):
            pag_indice = p["n"]
        for sid, rx in ANCORAS:
            if sid in mapa:
                continue
            if re.search(rx, t):
                # o indice cita todas as secoes; nao contar essa pagina
                if pag_indice is not None and p["n"] == pag_indice and sid not in ("indice", "capa"):
                    continue
                mapa[sid] = p["n"]
    return mapa


def montar_alignment(g_pags, s_pags):
    g_sec = secoes_por_pagina(g_pags)
    s_sec = secoes_por_pagina(s_pags)
    linhas = []
    for sid, _ in ANCORAS:
        gp, sp = g_sec.get(sid), s_sec.get(sid)
        if gp is None and sp is None:
            continue
        # quantas paginas do golden essa secao ocupa ate a proxima ancora
        prox = [v for k, v in sorted(g_sec.items(), key=lambda kv: kv[1]) if gp and v > gp]
        g_ate = (min(prox) - 1) if prox else (len(g_pags) if gp else None)
        n_g = (g_ate - gp + 1) if gp else 0
        prox_s = [v for k, v in sorted(s_sec.items(), key=lambda kv: kv[1]) if sp and v > sp]
        s_ate = (min(prox_s) - 1) if prox_s else (len(s_pags) if sp else None)
        n_s = (s_ate - sp + 1) if sp else 0
        if gp and sp:
            rel = "1:1" if n_g == n_s else "%d:%d" % (n_s, n_g)
            conf = "ALTA" if n_g == n_s else "MEDIA"
        else:
            rel, conf = ("0:%d" % n_g if gp else "%d:0" % n_s), "BAIXA"
        linhas.append({
            "secao": sid,
            "shadow_pagina": sp, "shadow_ate": s_ate, "shadow_n_paginas": n_s,
            "golden_pagina": gp, "golden_ate": g_ate, "golden_n_paginas": n_g,
            "relacao": rel, "confianca": conf,
            "papel_golden": g_pags[gp - 1]["papel"] if gp else None,
            "orient_golden": g_pags[gp - 1]["orientacao"] if gp else None,
            "papel_shadow": s_pags[sp - 1]["papel"] if sp else None,
            "orient_shadow": s_pags[sp - 1]["orientacao"] if sp else None,
        })
    return linhas


# -----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--shadow-pdf", default=None)
    args = ap.parse_args()

    U.assert_read_only(REVIEW_DIR)
    os.makedirs(REVIEW_DIR, exist_ok=True)

    g_pdf, g_docx = golden_paths()
    print("golden pdf :", g_pdf)
    print("golden docx:", g_docx)

    s_pdf = args.shadow_pdf or os.path.join(SHADOW_DIR, "shadow_report.pdf")
    tem_s_pdf = os.path.exists(s_pdf)

    print("\n-- renderizando golden @%d dpi" % args.dpi)
    g_pags = render(g_pdf, os.path.join(REVIEW_DIR, "golden"), args.dpi)
    print("   %d paginas" % len(g_pags))

    if tem_s_pdf:
        print("-- renderizando shadow @%d dpi" % args.dpi)
        s_pags = render(s_pdf, os.path.join(REVIEW_DIR, "shadow"), args.dpi)
        print("   %d paginas" % len(s_pags))
    else:
        s_pags = []
        print("-- SHADOW SEM PDF: %s" % s_pdf)
        print("   as medidas do shadow saem do proprio DOCX (exatas), nao de pixel.")

    print("-- lendo OOXML do golden")
    local, partes = ler_docx(g_docx)
    doc_xml = partes.get("word/document.xml", "")
    st_xml = partes.get("word/styles.xml", "")
    padrao, sts = estilos(st_xml)
    heads = {n: partes[n] for n in partes if re.match(r"word/header\d+\.xml", n)}
    foots = {n: partes[n] for n in partes if re.match(r"word/footer\d+\.xml", n)}

    print("-- lendo OOXML do shadow")
    s_docx = os.path.join(SHADOW_DIR, "shadow_report.docx")
    s_partes = dissecar_local(s_docx)

    dados = {
        "gerado_em": U.datetime.now().astimezone().isoformat(timespec="seconds"),
        "dpi": args.dpi,
        "golden": {"pdf": g_pdf, "docx": g_docx, "copia_local": local,
                   "n_paginas": len(g_pags), "paginas": g_pags},
        "shadow": {"pdf": s_pdf if tem_s_pdf else None,
                   "pdf_disponivel": tem_s_pdf,
                   "docx": s_docx,
                   "n_paginas": len(s_pags) if tem_s_pdf else None,
                   "paginas": s_pags},
        "shadow_docx": s_partes,
        "golden_docx": {
            "sectPr": sectprs(doc_xml),
            "docDefaults": padrao,
            "estilos": {k: v for k, v in sts.items()
                        if v.get("default") or re.match(r"(Heading|Ttulo|Titulo|Normal|Legenda|Caption|TOC|Sumrio|Sumario)", k, re.I)
                        or v.get("tipo") == "paragraph"},
            "paragrafos": paragrafos_usados(doc_xml),
            "tabelas": tabelas(doc_xml),
            "headers": {n: {"texto": " ".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", x))[:300],
                            "n_imagens": len(re.findall(r"<pic:pic\b|<w:drawing\b", x)),
                            "props": _props(x)} for n, x in heads.items()},
            "footers": {n: {"texto": " ".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", x))[:300],
                            "tem_campo_pagina": "PAGE" in x,
                            "props": _props(x)} for n, x in foots.items()},
            "numbering_presente": "word/numbering.xml" in partes,
        },
        "page_alignment_map": montar_alignment(g_pags, s_pags),
    }

    # o texto integral das paginas sai do JSON principal (volumoso) para um arquivo proprio
    textos = {"golden": {p["n"]: p.pop("texto") for p in g_pags},
              "shadow": {p["n"]: p.pop("texto") for p in s_pags}}

    for nome, obj in (("preparo_visual.json", dados), ("paginas_texto.json", textos)):
        dest = os.path.join(REVIEW_DIR, nome)
        U.assert_read_only(dest)
        with open(dest, "w", encoding="utf-8") as fh:
            json.dump(obj, fh, ensure_ascii=False, indent=1)
        print("   ->", dest)

    print("\n=== PAGE ALIGNMENT MAP ===")
    print("%-20s %-14s %-14s %-8s %s" % ("secao", "shadow", "golden", "relacao", "conf"))
    for l in dados["page_alignment_map"]:
        print("%-20s %-14s %-14s %-8s %s" % (
            l["secao"],
            "%s-%s" % (l["shadow_pagina"], l["shadow_ate"]) if l["shadow_pagina"] else "-",
            "%s-%s" % (l["golden_pagina"], l["golden_ate"]) if l["golden_pagina"] else "-",
            l["relacao"], l["confianca"]))


if __name__ == "__main__":
    main()
