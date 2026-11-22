# -*- coding: utf-8 -*-
r"""
04_extrair_docx.py

Extracao estrutural READ-ONLY dos .docx da amostra (8 documentos).
Le o .docx como ZIP direto do SharePoint (via prog_util.lp) e inspeciona o XML
interno; nada e escrito na arvore do SharePoint.

Produz, em output/extracao/docx/:
  <OBRA>_RO<NNN>.json       - dados estruturados completos de um documento
  <OBRA>_RO<NNN>.outline.txt- esqueleto legivel (titulos, tabelas, imagens, quebras)
  _comparativo.json         - o que e igual / o que varia entre os 8 documentos
  _comparativo.txt          - mesma comparacao em texto

Cobertura: docProps (core/app/custom), settings, styles (definicao completa dos
estilos efetivamente usados), sectPr, cabecalhos/rodapes, campos
(fldSimple/instrText), tabelas (grid, larguras, bordas, cabecalho repetido),
imagens (EMU->cm, inline/flutuante, legenda associada), word/media,
word/embeddings (OLE), numbering.xml.

Uso:
  python scripts/04_extrair_docx.py
  python scripts/04_extrair_docx.py --obras RD101 RD103
"""
import io
import os
import sys
import re
import json
import zipfile
import argparse
import collections

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

from lxml import etree

DEST = os.path.join(U.OUT, "extracao", "docx")

NS = {
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "v":   "urn:schemas-microsoft-com:vml",
    "o":   "urn:schemas-microsoft-com:office:office",
    "cp":  "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc":  "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "ep":  "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "cust": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "vt":  "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
W = "{%s}" % NS["w"]


# ---------------------------------------------------------------- conversoes
def tw_cm(v):
    """twips -> cm (1440 twips = 1 pol = 2.54 cm)."""
    try:
        return round(float(v) / 1440.0 * 2.54, 2)
    except (TypeError, ValueError):
        return None


def emu_cm(v):
    """EMU -> cm (360000 EMU = 1 cm)."""
    try:
        return round(float(v) / 360000.0, 2)
    except (TypeError, ValueError):
        return None


def hp_pt(v):
    """meio-ponto -> ponto."""
    try:
        return float(v) / 2.0
    except (TypeError, ValueError):
        return None


def val(el, path):
    """Valor do atributo w:val do primeiro no que casar com `path`."""
    if el is None:
        return None
    n = el.find(path, NS)
    if n is None:
        return None
    v = n.get(W + "val")
    return v if v is not None else True


def attrs(el):
    return {etree.QName(k).localname: v for k, v in el.attrib.items()} if el is not None else None


def txt(el):
    """Texto concatenado de um no (w:t, w:tab->TAB, w:br->quebra)."""
    if el is None:
        return ""
    out = []
    for n in el.iter():
        q = etree.QName(n).localname
        if q == "t":
            out.append(n.text or "")
        elif q == "tab":
            out.append("\t")
        elif q in ("br", "cr"):
            out.append("\n")
        elif q == "noBreakHyphen":
            out.append("-")
    return "".join(out)


def limpa(s):
    return re.sub(r"[ \t]+", " ", (s or "")).strip()


# ---------------------------------------------------------------- formatacao
def ler_rPr(rPr):
    """Formatacao de caractere -> dict compacto (so o que existe)."""
    if rPr is None:
        return {}
    d = {}
    f = rPr.find("w:rFonts", NS)
    if f is not None:
        a = attrs(f)
        nome = a.get("ascii") or a.get("hAnsi") or a.get("cs")
        tema = a.get("asciiTheme") or a.get("hAnsiTheme")
        if nome:
            d["fonte"] = nome
        if tema:
            d["fonte_tema"] = tema
    sz = val(rPr, "w:sz")
    if sz not in (None, True):
        d["tamanho_pt"] = hp_pt(sz)
    c = rPr.find("w:color", NS)
    if c is not None:
        cor = c.get(W + "val")
        if cor:
            d["cor"] = cor
        if c.get(W + "themeColor"):
            d["cor_tema"] = c.get(W + "themeColor")
    for tag, chave in (("w:b", "negrito"), ("w:i", "italico"), ("w:caps", "caixa_alta"),
                       ("w:smallCaps", "versalete"), ("w:strike", "tachado"),
                       ("w:vanish", "oculto")):
        n = rPr.find(tag, NS)
        if n is not None and n.get(W + "val") not in ("0", "false"):
            d[chave] = True
    u = rPr.find("w:u", NS)
    if u is not None:
        d["sublinhado"] = u.get(W + "val")
    hl = val(rPr, "w:highlight")
    if hl:
        d["realce"] = hl
    sh = rPr.find("w:shd", NS)
    if sh is not None and sh.get(W + "fill") not in (None, "auto"):
        d["fundo"] = sh.get(W + "fill")
    sp = rPr.find("w:spacing", NS)
    if sp is not None and sp.get(W + "val"):
        d["espacamento_car_tw"] = int(sp.get(W + "val"))
    lang = rPr.find("w:lang", NS)
    if lang is not None and lang.get(W + "val"):
        d["idioma"] = lang.get(W + "val")
    return d


def ler_pPr(pPr):
    """Formatacao de paragrafo -> dict compacto."""
    if pPr is None:
        return {}
    d = {}
    est = val(pPr, "w:pStyle")
    if est:
        d["estilo"] = est
    jc = val(pPr, "w:jc")
    if jc:
        d["alinhamento"] = jc
    o = val(pPr, "w:outlineLvl")
    if o is not None and o is not True:
        d["outline_lvl"] = int(o)
    sp = pPr.find("w:spacing", NS)
    if sp is not None:
        a = attrs(sp)
        e = {}
        if "before" in a:
            e["antes_pt"] = round(int(a["before"]) / 20.0, 1)
        if "after" in a:
            e["depois_pt"] = round(int(a["after"]) / 20.0, 1)
        if "line" in a:
            if a.get("lineRule") in ("auto", None):
                e["entrelinha"] = round(int(a["line"]) / 240.0, 2)
                e["entrelinha_regra"] = a.get("lineRule", "auto")
            else:
                e["entrelinha_pt"] = round(int(a["line"]) / 20.0, 1)
                e["entrelinha_regra"] = a.get("lineRule")
        if a.get("beforeAutospacing") == "1":
            e["antes_auto"] = True
        if a.get("afterAutospacing") == "1":
            e["depois_auto"] = True
        if e:
            d["espacamento"] = e
    ind = pPr.find("w:ind", NS)
    if ind is not None:
        a = attrs(ind)
        e = {}
        for k, nk in (("left", "esq_cm"), ("start", "esq_cm"), ("right", "dir_cm"),
                      ("end", "dir_cm"), ("firstLine", "primeira_linha_cm"),
                      ("hanging", "deslocamento_cm")):
            if k in a:
                e[nk] = tw_cm(a[k])
        if e:
            d["recuo"] = e
    npr = pPr.find("w:numPr", NS)
    if npr is not None:
        d["numeracao"] = {"numId": val(npr, "w:numId"), "ilvl": val(npr, "w:ilvl")}
    if pPr.find("w:keepNext", NS) is not None:
        d["manter_com_proximo"] = True
    if pPr.find("w:pageBreakBefore", NS) is not None:
        d["quebra_pagina_antes"] = True
    if pPr.find("w:contextualSpacing", NS) is not None:
        d["espacamento_contextual"] = True
    sh = pPr.find("w:shd", NS)
    if sh is not None and sh.get(W + "fill") not in (None, "auto"):
        d["fundo"] = sh.get(W + "fill")
    bd = pPr.find("w:pBdr", NS)
    if bd is not None:
        d["bordas"] = {etree.QName(c).localname: attrs(c) for c in bd}
    rpr = ler_rPr(pPr.find("w:rPr", NS))
    if rpr:
        d["rPr_paragrafo"] = rpr
    return d


def ler_bordas(el):
    """w:tblBorders / w:tcBorders -> dict lado -> {estilo, largura_pt, cor}."""
    if el is None:
        return None
    out = {}
    for c in el:
        a = attrs(c)
        lado = etree.QName(c).localname
        item = {"estilo": a.get("val")}
        if a.get("sz"):
            item["largura_pt"] = round(int(a["sz"]) / 8.0, 2)   # oitavos de ponto
        if a.get("color"):
            item["cor"] = a["color"]
        if a.get("themeColor"):
            item["cor_tema"] = a["themeColor"]
        out[lado] = item
    return out or None


# ---------------------------------------------------------------- docProps
def ler_props(z):
    p = {}
    if "docProps/core.xml" in z.namelist():
        x = etree.fromstring(z.read("docProps/core.xml"))
        mapa = {
            "titulo": "dc:title", "assunto": "dc:subject", "autor": "dc:creator",
            "palavras_chave": "cp:keywords", "descricao": "dc:description",
            "ultimo_autor": "cp:lastModifiedBy", "revisao": "cp:revision",
            "ultima_impressao": "cp:lastPrinted", "categoria": "cp:category",
            "criado_em": "dcterms:created", "modificado_em": "dcterms:modified",
        }
        core = {}
        for k, q in mapa.items():
            n = x.find(q, NS)
            if n is not None and (n.text or "").strip():
                core[k] = n.text.strip()
        p["core"] = core
    if "docProps/app.xml" in z.namelist():
        x = etree.fromstring(z.read("docProps/app.xml"))
        app = {}
        for n in x:
            nome = etree.QName(n).localname
            if n.text and n.text.strip():
                app[nome] = n.text.strip()
        p["app"] = app
    if "docProps/custom.xml" in z.namelist():
        x = etree.fromstring(z.read("docProps/custom.xml"))
        cu = {}
        for n in x.findall("cust:property", NS):
            v = "".join((c.text or "") for c in n)
            cu[n.get("name")] = v
        p["custom"] = cu
    return p


def ler_settings(z):
    if "word/settings.xml" not in z.namelist():
        return {}
    x = etree.fromstring(z.read("word/settings.xml"))
    s = {}
    at = x.find("w:attachedTemplate", NS)
    s["template_anexado_rid"] = at.get(NS_r("id")) if at is not None else None
    for tag, chave in (("w:defaultTabStop", "tab_padrao_tw"), ("w:zoom", "zoom"),
                       ("w:characterSpacingControl", "controle_espacamento"),
                       ("w:themeFontLang", "idioma_tema")):
        n = x.find(tag, NS)
        if n is not None:
            s[chave] = attrs(n) if len(n.attrib) > 1 else (n.get(W + "val") or n.get(W + "percent"))
    for tag, chave in (("w:trackChanges", "controle_alteracoes"),
                       ("w:evenAndOddHeaders", "cabecalho_par_impar"),
                       ("w:documentProtection", "protecao"),
                       ("w:mailMerge", "mala_direta"),
                       ("w:updateFields", "atualizar_campos_ao_abrir")):
        n = x.find(tag, NS)
        if n is not None:
            s[chave] = attrs(n) or True
    s["n_rsids"] = len(x.findall(".//w:rsid", NS))
    cm = x.find("w:compat", NS)
    if cm is not None:
        s["compat"] = [attrs(c) for c in cm.findall("w:compatSetting", NS)]
    return s


def NS_r(nome):
    return "{%s}%s" % (NS["r"], nome)


# ---------------------------------------------------------------- styles.xml
def ler_styles(z):
    if "word/styles.xml" not in z.namelist():
        return {}, {}
    x = etree.fromstring(z.read("word/styles.xml"))
    dd = x.find("w:docDefaults", NS)
    padroes = {}
    if dd is not None:
        padroes["run"] = ler_rPr(dd.find("w:rPrDefault/w:rPr", NS))
        padroes["paragrafo"] = ler_pPr(dd.find("w:pPrDefault/w:pPr", NS))
    estilos = {}
    for s in x.findall("w:style", NS):
        sid = s.get(W + "styleId")
        d = {
            "id": sid,
            "tipo": s.get(W + "type"),
            "nome": val(s, "w:name"),
            "baseado_em": val(s, "w:basedOn"),
            "proximo": val(s, "w:next"),
            "vinculado": val(s, "w:link"),
            "padrao": s.get(W + "default") == "1",
            "personalizado": s.get(W + "customStyle") == "1",
            "qFormat": s.find("w:qFormat", NS) is not None,
            "pPr": ler_pPr(s.find("w:pPr", NS)),
            "rPr": ler_rPr(s.find("w:rPr", NS)),
        }
        tp = s.find("w:tblPr", NS)
        if tp is not None:
            d["tblPr"] = {
                "bordas": ler_bordas(tp.find("w:tblBorders", NS)),
                "margens_celula": {etree.QName(c).localname: attrs(c)
                                   for c in tp.findall("w:tblCellMar/*", NS)} or None,
                "estilo_linha": val(tp, "w:tblStyleRowBandSize"),
            }
        d = {k: v for k, v in d.items() if v not in (None, {}, False)}
        estilos[sid] = d
    return padroes, estilos


# ---------------------------------------------------------------- numbering
def ler_theme(z):
    """Fontes e cores do tema (resolve majorHAnsi/minorHAnsi e accent1..6)."""
    alvo = next((n for n in z.namelist() if n.startswith("word/theme/")), None)
    if not alvo:
        return {}
    x = etree.fromstring(z.read(alvo))
    t = {"parte": alvo}
    fs = x.find(".//a:fontScheme", NS)
    if fs is not None:
        t["esquema_fonte"] = fs.get("name")
        for tag, chave in (("a:majorFont", "majorHAnsi"), ("a:minorFont", "minorHAnsi")):
            n = fs.find(tag + "/a:latin", NS)
            if n is not None:
                t[chave] = n.get("typeface")
    cs = x.find(".//a:clrScheme", NS)
    if cs is not None:
        t["esquema_cor"] = cs.get("name")
        cores = {}
        for c in cs:
            nome = etree.QName(c).localname
            srgb = c.find("a:srgbClr", NS)
            sysc = c.find("a:sysClr", NS)
            cores[nome] = (srgb.get("val") if srgb is not None else
                           (sysc.get("lastClr") if sysc is not None else None))
        t["cores"] = cores
    return t


def resolver_outline(estilos):
    """styleId -> nivel de outline efetivo (herdado por basedOn). None se nao for titulo."""
    cache = {}

    def f(sid, prof=0):
        if sid in cache:
            return cache[sid]
        if sid not in estilos or prof > 12:
            return None
        v = estilos[sid].get("pPr", {}).get("outline_lvl")
        if v is None:
            base = estilos[sid].get("baseado_em")
            v = f(base, prof + 1) if base else None
        cache[sid] = v
        return v
    return {k: f(k) for k in estilos}


def resolver_formato(estilos, padroes, sid):
    """Formatacao efetiva de um estilo: docDefaults + cadeia basedOn + o proprio."""
    cadeia = []
    cur, guard = sid, 0
    while cur and cur in estilos and guard < 12:
        cadeia.append(cur)
        cur = estilos[cur].get("baseado_em")
        guard += 1
    rpr = dict(padroes.get("run", {}))
    ppr = dict(padroes.get("paragrafo", {}))
    for k in reversed(cadeia):
        rpr.update(estilos[k].get("rPr", {}))
        p = dict(estilos[k].get("pPr", {}))
        esp = p.pop("espacamento", None)
        if esp:
            ppr.setdefault("espacamento", {})
            ppr["espacamento"] = {**ppr.get("espacamento", {}), **esp}
        ppr.update(p)
    return {"cadeia_basedOn": cadeia, "rPr_efetivo": rpr, "pPr_efetivo": ppr}


def ler_numbering(z):
    if "word/numbering.xml" not in z.namelist():
        return {}
    x = etree.fromstring(z.read("word/numbering.xml"))
    abst = {}
    for a in x.findall("w:abstractNum", NS):
        aid = a.get(W + "abstractNumId")
        niveis = {}
        for lv in a.findall("w:lvl", NS):
            i = lv.get(W + "ilvl")
            niveis[i] = {
                "formato": val(lv, "w:numFmt"),
                "texto": val(lv, "w:lvlText"),
                "inicio": val(lv, "w:start"),
                "estilo_paragrafo": val(lv, "w:pStyle"),
                "alinhamento": val(lv, "w:lvlJc"),
                "recuo": ler_pPr(lv.find("w:pPr", NS)).get("recuo"),
                "rPr": ler_rPr(lv.find("w:rPr", NS)),
            }
        abst[aid] = {
            "multinivel": val(a, "w:multiLevelType"),
            "nsid": (a.find("w:nsid", NS).get(W + "val") if a.find("w:nsid", NS) is not None else None),
            "niveis": niveis,
        }
    nums = {}
    for n in x.findall("w:num", NS):
        nums[n.get(W + "numId")] = val(n, "w:abstractNumId")
    return {"num_para_abstract": nums, "abstract": abst}


# ---------------------------------------------------------------- campos
RE_CAMPO = re.compile(r"^\s*([A-Z][A-Z0-9_]*)")


def coleta_campos(raiz):
    """fldSimple + instrText -> lista de {tipo, instrucao}."""
    campos = []
    for fs in raiz.iter(W + "fldSimple"):
        instr = fs.get(W + "instr") or ""
        m = RE_CAMPO.match(instr)
        campos.append({"forma": "fldSimple", "tipo": m.group(1) if m else "?",
                       "instrucao": limpa(instr), "resultado": limpa(txt(fs))[:120]})
    # instrText pode vir quebrado em varios runs; concatena entre begin e end
    buf = None
    for n in raiz.iter():
        q = etree.QName(n).localname
        if q == "fldChar":
            t = n.get(W + "fldCharType")
            if t == "begin":
                buf = []
            elif t in ("separate", "end") and buf is not None:
                instr = "".join(buf)
                if instr.strip():
                    m = RE_CAMPO.match(instr)
                    campos.append({"forma": "instrText", "tipo": m.group(1) if m else "?",
                                   "instrucao": limpa(instr)})
                buf = None if t == "end" else buf
        elif q == "instrText" and buf is not None:
            buf.append(n.text or "")
    return campos


# ---------------------------------------------------------------- imagens
def coleta_imagens(raiz, rels, contexto):
    """Imagens DrawingML (wp:inline/wp:anchor) + VML (v:imagedata)."""
    imgs = []
    for dr in raiz.iter(W + "drawing"):
        for tag, modo in (("{%s}inline" % NS["wp"], "em_linha"),
                          ("{%s}anchor" % NS["wp"], "flutuante")):
            for nd in dr.iter(tag):
                ext = nd.find("wp:extent", NS)
                dp = nd.find("wp:docPr", NS)
                blip = nd.find(".//a:blip", NS)
                rid = blip.get(NS_r("embed")) or blip.get(NS_r("link")) if blip is not None else None
                item = {
                    "modo": modo,
                    "nome_word": dp.get("name") if dp is not None else None,
                    "descricao": dp.get("descr") if dp is not None else None,
                    "rId": rid,
                    "alvo": rels.get(rid, {}).get("alvo") if rid else None,
                    "largura_cm": emu_cm(ext.get("cx")) if ext is not None else None,
                    "altura_cm": emu_cm(ext.get("cy")) if ext is not None else None,
                    "cx_emu": int(ext.get("cx")) if ext is not None else None,
                    "cy_emu": int(ext.get("cy")) if ext is not None else None,
                    "contexto": contexto,
                }
                if modo == "flutuante":
                    item["ancoragem"] = {
                        "atras_do_texto": nd.get("behindDoc") == "1",
                        "h_rel": (nd.find("wp:positionH", NS).get("relativeFrom")
                                  if nd.find("wp:positionH", NS) is not None else None),
                        "v_rel": (nd.find("wp:positionV", NS).get("relativeFrom")
                                  if nd.find("wp:positionV", NS) is not None else None),
                        "quebra_texto": (etree.QName(nd.find("wp:wrapSquare", NS) if
                                         nd.find("wp:wrapSquare", NS) is not None else
                                         (nd.find("wp:wrapNone", NS) if
                                          nd.find("wp:wrapNone", NS) is not None else nd)).localname),
                    }
                # recorte / borda da imagem
                sp = nd.find(".//pic:spPr", NS)
                if sp is not None and sp.find("a:ln/a:solidFill", NS) is not None:
                    item["tem_borda"] = True
                src = nd.find(".//a:srcRect", NS)
                if src is not None and src.attrib:
                    item["recorte"] = attrs(src)
                imgs.append(item)
    for vm in raiz.iter("{%s}imagedata" % NS["v"]):
        rid = vm.get(NS_r("id"))
        pai = vm.getparent()
        estilo = pai.get("style") if pai is not None else None
        imgs.append({"modo": "vml", "rId": rid,
                     "alvo": rels.get(rid, {}).get("alvo") if rid else None,
                     "estilo_vml": estilo, "contexto": contexto})
    return imgs


# ---------------------------------------------------------------- tabelas
def ler_tabela(tbl, rels, idx, nivel=0):
    tp = tbl.find("w:tblPr", NS)
    grid = [int(g.get(W + "w")) for g in tbl.findall("w:tblGrid/w:gridCol", NS)
            if g.get(W + "w")]
    tw = tp.find("w:tblW", NS) if tp is not None else None
    ind = tp.find("w:tblInd", NS) if tp is not None else None
    lay = tp.find("w:tblLayout", NS) if tp is not None else None
    look = tp.find("w:tblLook", NS) if tp is not None else None
    t = {
        "idx": idx,
        "nivel_aninhamento": nivel,
        "estilo": val(tp, "w:tblStyle") if tp is not None else None,
        "largura": attrs(tw),
        "largura_total_cm": round(sum(tw_cm(g) or 0 for g in grid), 2) if grid else None,
        "alinhamento": val(tp, "w:jc") if tp is not None else None,
        "recuo": attrs(ind),
        "layout": lay.get(W + "type") if lay is not None else "autofit",
        "tblLook": attrs(look),
        "bordas": ler_bordas(tp.find("w:tblBorders", NS)) if tp is not None else None,
        "margens_celula": ({etree.QName(c).localname: attrs(c)
                            for c in tp.findall("w:tblCellMar/*", NS)} if tp is not None else None),
        "grid_tw": grid,
        "grid_cm": [tw_cm(g) for g in grid],
        "n_colunas": len(grid),
        "linhas": [],
    }
    for ri, tr in enumerate(tbl.findall("w:tr", NS)):
        trPr = tr.find("w:trPr", NS)
        h = trPr.find("w:trHeight", NS) if trPr is not None else None
        linha = {
            "i": ri,
            "cabecalho_repetido": trPr is not None and trPr.find("w:tblHeader", NS) is not None,
            "nao_quebrar": trPr is not None and trPr.find("w:cantSplit", NS) is not None,
            "altura": attrs(h),
            "celulas": [],
        }
        for tc in tr.findall("w:tc", NS):
            tcPr = tc.find("w:tcPr", NS)
            w_ = tcPr.find("w:tcW", NS) if tcPr is not None else None
            sh = tcPr.find("w:shd", NS) if tcPr is not None else None
            vm = tcPr.find("w:vMerge", NS) if tcPr is not None else None
            cel = {
                "largura_tw": int(w_.get(W + "w")) if w_ is not None and (w_.get(W + "w") or "").isdigit() else None,
                "largura_cm": tw_cm(w_.get(W + "w")) if w_ is not None else None,
                "gridSpan": int(val(tcPr, "w:gridSpan") or 1) if tcPr is not None else 1,
                "vMerge": (vm.get(W + "val") or "continue") if vm is not None else None,
                "valign": val(tcPr, "w:vAlign") if tcPr is not None else None,
                "fundo": sh.get(W + "fill") if sh is not None else None,
                "bordas": ler_bordas(tcPr.find("w:tcBorders", NS)) if tcPr is not None else None,
                "texto": limpa(txt(tc)),
                "paragrafos": [],
            }
            for p in tc.findall("w:p", NS):
                fmt = ler_pPr(p.find("w:pPr", NS))
                rpr = [ler_rPr(r.find("w:rPr", NS)) for r in p.findall("w:r", NS)]
                rpr = [x for x in rpr if x]
                cel["paragrafos"].append({"texto": limpa(txt(p)), "pPr": fmt,
                                          "rPr": rpr[0] if rpr else {}})
            cel["imagens"] = coleta_imagens(tc, rels, "tabela%d" % idx)
            cel = {k: v for k, v in cel.items() if v not in (None, [], {}, 1) or k in ("texto",)}
            linha["celulas"].append(cel)
        linha = {k: v for k, v in linha.items() if v not in (None, False)}
        t["linhas"].append(linha)
    t["n_linhas"] = len(t["linhas"])
    t["assinatura"] = assinatura_tabela(t)
    return t


def assinatura_tabela(t):
    """Chave para reconhecer a mesma tabela entre documentos: 1a linha + n_colunas."""
    prim = ""
    if t["linhas"]:
        prim = " | ".join(U.norm(c.get("texto", "")) for c in t["linhas"][0]["celulas"])
    return "%dcol :: %s" % (t["n_colunas"], prim[:160])


# ---------------------------------------------------------------- sectPr
def ler_sectPr(sp, rels, i):
    pg = sp.find("w:pgSz", NS)
    mg = sp.find("w:pgMar", NS)
    a_pg, a_mg = attrs(pg) or {}, attrs(mg) or {}
    largura, altura = a_pg.get("w"), a_pg.get("h")
    orient = a_pg.get("orient") or "portrait"
    d = {
        "i": i,
        "tipo_quebra": val(sp, "w:type") or "nextPage(implicito)",
        "pagina_tw": {"largura": int(largura) if largura else None,
                      "altura": int(altura) if altura else None},
        "pagina_cm": {"largura": tw_cm(largura), "altura": tw_cm(altura)},
        "pagina_code": a_pg.get("code"),
        "orientacao": orient,
        "formato_papel": papel(largura, altura, orient),
        "margens_cm": {k: tw_cm(v) for k, v in a_mg.items() if k in
                       ("top", "right", "bottom", "left", "header", "footer", "gutter")},
        "margens_tw": a_mg,
        "colunas": attrs(sp.find("w:cols", NS)),
        "primeira_pagina_diferente": sp.find("w:titlePg", NS) is not None,
        "num_pagina": attrs(sp.find("w:pgNumType", NS)),
        "referencias": [],
    }
    for tag, papelx in (("w:headerReference", "cabecalho"), ("w:footerReference", "rodape")):
        for rf in sp.findall(tag, NS):
            rid = rf.get(NS_r("id"))
            d["referencias"].append({"papel": papelx, "tipo": rf.get(W + "type"),
                                     "rId": rid, "parte": rels.get(rid, {}).get("alvo")})
    return d


def papel(w_, h_, orient):
    if not w_ or not h_:
        return None
    a, b = sorted([int(w_), int(h_)])
    for nome, (pa, pb) in (("A5", (8391, 11906)), ("A4", (11906, 16838)),
                           ("A3", (16838, 23811)), ("Carta", (12240, 15840)),
                           ("Oficio", (12240, 20160))):
        if abs(a - pa) <= 30 and abs(b - pb) <= 40:
            return "%s %s" % (nome, "retrato" if orient == "portrait" else "paisagem")
    return "%.1fx%.1f cm" % (tw_cm(w_), tw_cm(h_))


# ---------------------------------------------------------------- rels
def ler_rels(z, parte):
    d = os.path.dirname(parte)
    caminho = "%s/_rels/%s.rels" % (d, os.path.basename(parte))
    if caminho not in z.namelist():
        return {}
    x = etree.fromstring(z.read(caminho))
    out = {}
    for r in x.findall("rel:Relationship", NS):
        out[r.get("Id")] = {"tipo": r.get("Type").rsplit("/", 1)[-1],
                            "alvo": r.get("Target"),
                            "externo": r.get("TargetMode") == "External"}
    return out


# ---------------------------------------------------------------- cabec/rodape
def ler_hf(z, parte, rels_doc):
    x = etree.fromstring(z.read("word/" + parte.lstrip("/")))
    rels = ler_rels(z, "word/" + parte.lstrip("/"))
    paras = []
    for p in x.iter(W + "p"):
        # ignora paragrafos dentro de tabela (serao vistos na tabela)
        paras.append({"texto": limpa(txt(p)), "pPr": ler_pPr(p.find("w:pPr", NS))})
    return {
        "parte": "word/" + parte.lstrip("/"),
        "paragrafos": [p for p in paras if p["texto"] or p["pPr"]],
        "texto": limpa(" | ".join(p["texto"] for p in paras if p["texto"])),
        "campos": coleta_campos(x),
        "imagens": coleta_imagens(x, rels, parte),
        "n_tabelas": len(list(x.iter(W + "tbl"))),
        "rels": rels,
    }


# ---------------------------------------------------------------- documento
RE_LEGENDA = re.compile(r"^\s*(figura|tabela|quadro|foto|gr[aá]fico|imagem)\s*[:\-–]?\s*(\d+)?",
                        re.I)


def abrir_zip(caminho):
    r"""Le o .docx inteiro para memoria e devolve um ZipFile.

    Necessario porque arquivos do OneDrive ainda nao hidratados falham com
    OSError [Errno 22] no acesso aleatorio que o zipfile faz (seek para o EOCD).
    A leitura sequencial forca o download; nada e escrito na origem.
    """
    ultimo = None
    for _ in range(3):
        try:
            buf = io.BytesIO()
            with open(U.lp(caminho), "rb") as fh:
                while True:
                    b = fh.read(4 * 1024 * 1024)
                    if not b:
                        break
                    buf.write(b)
            buf.seek(0)
            return zipfile.ZipFile(buf)
        except (OSError, zipfile.BadZipFile) as e:
            ultimo = e
    raise ultimo


def extrair(caminho, rotulo):
    info = {"rotulo": rotulo, "arquivo": caminho}
    with abrir_zip(caminho) as z:
        nomes = z.namelist()
        info["partes"] = sorted(
            [{"nome": n, "bytes": z.getinfo(n).file_size} for n in nomes],
            key=lambda d: d["nome"])
        info["propriedades"] = ler_props(z)
        info["settings"] = ler_settings(z)
        padroes, estilos = ler_styles(z)
        info["tema"] = ler_theme(z)
        outline_por_estilo = resolver_outline(estilos)
        info["outline_por_estilo"] = {k: v for k, v in outline_por_estilo.items()
                                      if v is not None}
        info["numeracao"] = ler_numbering(z)
        rels = ler_rels(z, "word/document.xml")
        info["rels_documento"] = rels

        # midia e embeddings
        info["midia"] = sorted(
            [{"nome": n, "bytes": z.getinfo(n).file_size, "ext": os.path.splitext(n)[1].lower()}
             for n in nomes if n.startswith("word/media/")], key=lambda d: d["nome"])
        info["embeddings"] = sorted(
            [{"nome": n, "bytes": z.getinfo(n).file_size, "ext": os.path.splitext(n)[1].lower()}
             for n in nomes if n.startswith("word/embeddings/")], key=lambda d: d["nome"])
        info["tem_ole"] = bool(info["embeddings"])
        info["outras_partes_notaveis"] = [n for n in nomes if n.startswith(("word/charts",
                                          "word/drawings", "customXml/", "word/glossary"))]

        x = etree.fromstring(z.read("word/document.xml"))
        body = x.find("w:body", NS)

        # ---- secoes
        secoes = []
        for i, sp in enumerate(body.iter(W + "sectPr")):
            secoes.append(ler_sectPr(sp, rels, i))
        info["secoes"] = secoes

        # ---- cabecalhos e rodapes
        hf = {}
        for rid, r in rels.items():
            if r["tipo"] in ("header", "footer"):
                try:
                    hf[r["alvo"]] = ler_hf(z, r["alvo"], rels)
                    hf[r["alvo"]]["papel"] = r["tipo"]
                    hf[r["alvo"]]["rId"] = rid
                except Exception as e:                       # pragma: no cover
                    hf[r["alvo"]] = {"erro": str(e)}
        info["cabecalhos_rodapes"] = hf

        # ---- corpo, em ordem
        paragrafos, tabelas, imagens, campos, titulos = [], [], [], [], []
        quebras = []
        pi = ti = 0
        pend_legenda = None

        def walk(el, dentro_tabela=False):
            nonlocal pi, ti, pend_legenda
            for n in el:
                q = etree.QName(n).localname
                if q == "p":
                    pPr = n.find("w:pPr", NS)
                    fmt = ler_pPr(pPr)
                    t = limpa(txt(n))
                    runs = []
                    for r in n.findall("w:r", NS):
                        rt = txt(r)
                        if not rt.strip():
                            continue
                        runs.append({"t": rt, "rPr": ler_rPr(r.find("w:rPr", NS))})
                    item = {"i": pi, "estilo": fmt.get("estilo"), "texto": t,
                            "pPr": fmt, "runs": runs[:8], "n_runs": len(runs)}
                    ims = coleta_imagens(n, rels, "p%d" % pi)
                    if ims:
                        for im in ims:
                            im["paragrafo"] = pi
                        imagens.extend(ims)
                        item["imagens"] = len(ims)
                    cps = coleta_campos(n)
                    if cps:
                        for c in cps:
                            c["paragrafo"] = pi
                        campos.extend(cps)
                        item["campos"] = [c["tipo"] for c in cps]
                    if n.find(".//w:br[@w:type='page']", NS) is not None:
                        item["quebra_pagina_manual"] = True
                        quebras.append({"paragrafo": pi, "tipo": "br page"})
                    if pPr is not None and pPr.find("w:sectPr", NS) is not None:
                        item["fim_de_secao"] = True
                        quebras.append({"paragrafo": pi, "tipo": "sectPr"})
                    lvl = fmt.get("outline_lvl")
                    if lvl is None and fmt.get("estilo"):
                        lvl = outline_por_estilo.get(fmt["estilo"])
                    # outlineLvl 9 = "corpo de texto" (nao entra no sumario)
                    if lvl is not None and lvl < 9 and t:
                        titulos.append({"paragrafo": pi, "estilo": fmt.get("estilo"),
                                        "nivel": lvl, "texto": t,
                                        "numeracao": fmt.get("numeracao") or
                                        estilos.get(fmt.get("estilo"), {})
                                               .get("pPr", {}).get("numeracao")})
                        item["titulo_nivel"] = lvl
                    if RE_LEGENDA.match(t) or (fmt.get("estilo") or "") in (
                            "Legenda", "Caption", "Titulodefiguras"):
                        item["legenda"] = True
                        if imagens and imagens[-1].get("paragrafo") in (pi, pi - 1, pi - 2):
                            imagens[-1]["legenda"] = t
                    paragrafos.append(item)
                    pi += 1
                elif q == "tbl":
                    t = ler_tabela(n, rels, ti, nivel=1 if dentro_tabela else 0)
                    t["apos_paragrafo"] = pi - 1
                    tabelas.append(t)
                    ti += 1
                    for im in t.pop("_imgs", []):
                        imagens.append(im)
                elif q in ("sdt",):
                    conteudo = n.find("w:sdtContent", NS)
                    sdtPr = n.find("w:sdtPr", NS)
                    alias = val(sdtPr, "w:alias") if sdtPr is not None else None
                    gal = sdtPr.find(".//w:docPartGallery", NS) if sdtPr is not None else None
                    if conteudo is not None:
                        paragrafos.append({"i": pi, "estilo": "__SDT__",
                                           "texto": "[controle de conteudo: %s]" %
                                                    (alias or (gal.get(W + "val") if gal is not None else "?")),
                                           "pPr": {}, "runs": [], "n_runs": 0,
                                           "sdt_galeria": gal.get(W + "val") if gal is not None else None})
                        pi += 1
                        walk(conteudo, dentro_tabela)

        walk(body)

        # imagens dentro de tabelas (coletadas na leitura da tabela)
        for t in tabelas:
            for ln in t["linhas"]:
                for c in ln["celulas"]:
                    for im in c.get("imagens", []) or []:
                        im["tabela"] = t["idx"]
                        imagens.append(im)

        info["paragrafos"] = paragrafos
        info["tabelas"] = tabelas
        info["imagens"] = imagens
        info["campos"] = campos
        info["titulos"] = titulos
        info["quebras"] = quebras

        # ---- estilos usados
        usados = collections.Counter(p["estilo"] for p in paragrafos if p.get("estilo"))
        usados.update({t["estilo"]: 1 for t in tabelas if t.get("estilo")})
        for h in hf.values():
            for p in h.get("paragrafos", []):
                if p["pPr"].get("estilo"):
                    usados[p["pPr"]["estilo"]] += 1
        info["estilos_usados"] = usados.most_common()
        info["estilos_definidos_usados"] = {k: estilos[k] for k in usados if k in estilos}
        # estilos ancestrais dos usados (basedOn)
        fila = list(usados)
        vistos = set(fila)
        while fila:
            k = fila.pop()
            base = estilos.get(k, {}).get("baseado_em")
            if base and base not in vistos:
                vistos.add(base)
                fila.append(base)
                if base in estilos:
                    info["estilos_definidos_usados"][base] = estilos[base]
        info["padroes_documento"] = padroes
        info["formato_efetivo"] = {k: resolver_formato(estilos, padroes, k)
                                   for k in info["estilos_definidos_usados"]}
        info["n_estilos_definidos"] = len(estilos)
        info["estilos_titulo_definidos"] = {
            k: v for k, v in estilos.items()
            if (v.get("nome") or "").lower().startswith(("heading", "título", "titulo"))
            or (v.get("pPr", {}).get("outline_lvl") is not None and v.get("tipo") == "paragraph")}

        info["resumo"] = {
            "n_paragrafos": len(paragrafos),
            "n_paragrafos_com_texto": sum(1 for p in paragrafos if p["texto"]),
            "n_tabelas": len(tabelas),
            "n_imagens": len(imagens),
            "n_secoes": len(secoes),
            "n_campos": len(campos),
            "tipos_campo": collections.Counter(c["tipo"] for c in campos).most_common(),
            "n_midia": len(info["midia"]),
            "bytes_midia": sum(m["bytes"] for m in info["midia"]),
            "formatos_midia": collections.Counter(m["ext"] for m in info["midia"]).most_common(),
            "tem_ole": info["tem_ole"],
            "orientacoes": collections.Counter(s["formato_papel"] for s in secoes).most_common(),
            "estilos_usados": usados.most_common(),
            "n_titulos": len(titulos),
        }
    return info


# ---------------------------------------------------------------- saida texto
def outline(info):
    L = []
    p = info["propriedades"]
    L.append("=" * 78)
    L.append("%s   %s" % (info["rotulo"], os.path.basename(info["arquivo"])))
    L.append("=" * 78)
    L.append("autor=%s  ultimo=%s  rev=%s  criado=%s  modif=%s  tempo_edicao=%smin" % (
        p.get("core", {}).get("autor"), p.get("core", {}).get("ultimo_autor"),
        p.get("core", {}).get("revisao"), p.get("core", {}).get("criado_em"),
        p.get("core", {}).get("modificado_em"), p.get("app", {}).get("TotalTime")))
    L.append("Word: paginas=%s palavras=%s template=%s AppVersion=%s" % (
        p.get("app", {}).get("Pages"), p.get("app", {}).get("Words"),
        p.get("app", {}).get("Template"), p.get("app", {}).get("AppVersion")))
    r = info["resumo"]
    L.append("secoes=%d paragrafos=%d tabelas=%d imagens=%d campos=%d midia=%d OLE=%s" % (
        r["n_secoes"], r["n_paragrafos"], r["n_tabelas"], r["n_imagens"],
        r["n_campos"], r["n_midia"], r["tem_ole"]))
    L.append("")
    L.append("--- SECOES ---")
    for s in info["secoes"]:
        L.append("  [%d] %s  %sx%s cm  margens T%s D%s B%s E%s  hdr=%s ftr=%s  refs=%s" % (
            s["i"], s["formato_papel"], s["pagina_cm"]["largura"], s["pagina_cm"]["altura"],
            s["margens_cm"].get("top"), s["margens_cm"].get("right"),
            s["margens_cm"].get("bottom"), s["margens_cm"].get("left"),
            s["margens_cm"].get("header"), s["margens_cm"].get("footer"),
            ",".join("%s:%s" % (x["papel"], x["parte"]) for x in s["referencias"]) or "-"))
    L.append("")
    L.append("--- CABECALHOS / RODAPES ---")
    for k, h in sorted(info["cabecalhos_rodapes"].items()):
        L.append("  %s (%s): texto=%r imgs=%d campos=%s" % (
            k, h.get("papel"), (h.get("texto") or "")[:80], len(h.get("imagens", [])),
            [c["tipo"] for c in h.get("campos", [])]))
        for im in h.get("imagens", []):
            L.append("        img %s %s  %sx%s cm" % (im.get("nome_word"), im.get("alvo"),
                                                      im.get("largura_cm"), im.get("altura_cm")))
    L.append("")
    L.append("--- CORPO (ordem) ---")
    tbl_por_p = collections.defaultdict(list)
    for t in info["tabelas"]:
        tbl_por_p[t["apos_paragrafo"]].append(t)
    img_por_p = collections.defaultdict(list)
    for im in info["imagens"]:
        if im.get("paragrafo") is not None:
            img_por_p[im["paragrafo"]].append(im)
    for pa in info["paragrafos"]:
        marca = ""
        if pa.get("fim_de_secao"):
            marca += " <<FIM DE SECAO>>"
        if pa.get("quebra_pagina_manual"):
            marca += " <<QUEBRA DE PAGINA>>"
        if pa.get("campos"):
            marca += " <<campos: %s>>" % ",".join(pa["campos"])
        t = pa["texto"]
        if t or marca or pa.get("imagens"):
            L.append("  %4d [%-16s] %s%s" % (pa["i"], pa.get("estilo") or "-", t[:150], marca))
        for im in img_por_p.get(pa["i"], []):
            L.append("       IMG %-12s %-16s %5.2f x %5.2f cm  %s  leg=%s" % (
                im.get("alvo"), im.get("nome_word"), im.get("largura_cm") or 0,
                im.get("altura_cm") or 0, im.get("modo"), (im.get("legenda") or "")[:60]))
        for t2 in tbl_por_p.get(pa["i"], []):
            L.append("       TABELA #%d  %dx%d  estilo=%s  larguras_cm=%s" % (
                t2["idx"], t2["n_linhas"], t2["n_colunas"], t2["estilo"], t2["grid_cm"]))
            for ln in t2["linhas"][:3]:
                L.append("         | " + " | ".join((c.get("texto") or "")[:28]
                                                    for c in ln["celulas"]))
    L.append("")
    L.append("--- MIDIA ---")
    for m in info["midia"]:
        L.append("  %-26s %9d %s" % (m["nome"], m["bytes"], m["ext"]))
    if info["embeddings"]:
        L.append("--- EMBEDDINGS (OLE) ---")
        for m in info["embeddings"]:
            L.append("  %-26s %9d" % (m["nome"], m["bytes"]))
    return "\n".join(L)


# ---------------------------------------------------------------- comparativo
def comparar(docs):
    c = {"documentos": [d["rotulo"] for d in docs], "igual": {}, "varia": {}}

    def campo(nome, fn):
        vals = {}
        for d in docs:
            try:
                vals[d["rotulo"]] = fn(d)
            except Exception:
                vals[d["rotulo"]] = None
        distintos = {json.dumps(v, ensure_ascii=False, sort_keys=True) for v in vals.values()}
        if len(distintos) == 1:
            c["igual"][nome] = list(vals.values())[0]
        else:
            c["varia"][nome] = vals

    campo("autor(dc:creator)", lambda d: d["propriedades"]["core"].get("autor"))
    campo("ultimo_autor", lambda d: d["propriedades"]["core"].get("ultimo_autor"))
    campo("criado_em", lambda d: d["propriedades"]["core"].get("criado_em"))
    campo("modificado_em", lambda d: d["propriedades"]["core"].get("modificado_em"))
    campo("revisao", lambda d: d["propriedades"]["core"].get("revisao"))
    campo("ultima_impressao", lambda d: d["propriedades"]["core"].get("ultima_impressao"))
    campo("Template", lambda d: d["propriedades"]["app"].get("Template"))
    campo("AppVersion", lambda d: d["propriedades"]["app"].get("AppVersion"))
    campo("Application", lambda d: d["propriedades"]["app"].get("Application"))
    campo("TotalTime_min", lambda d: d["propriedades"]["app"].get("TotalTime"))
    campo("Pages_Word", lambda d: d["propriedades"]["app"].get("Pages"))
    campo("Words_Word", lambda d: d["propriedades"]["app"].get("Words"))
    campo("custom_ContentTypeId", lambda d: d["propriedades"].get("custom", {}).get("ContentTypeId"))
    campo("n_secoes", lambda d: len(d["secoes"]))
    campo("formatos_pagina", lambda d: [s["formato_papel"] for s in d["secoes"]])
    campo("margens_cm_distintas", lambda d: sorted({json.dumps(s["margens_cm"], sort_keys=True)
                                                    for s in d["secoes"]}))
    campo("tipos_campo", lambda d: sorted(dict(d["resumo"]["tipos_campo"]).items()))
    campo("estilos_usados", lambda d: sorted(dict(d["estilos_usados"]).items()))
    campo("estilos_titulo_definidos", lambda d: sorted(d["estilos_titulo_definidos"]))
    campo("n_estilos_definidos", lambda d: d["n_estilos_definidos"])
    campo("padroes_documento", lambda d: d["padroes_documento"])
    campo("tem_OLE", lambda d: d["tem_ole"])
    campo("formatos_midia", lambda d: sorted(dict(d["resumo"]["formatos_midia"]).items()))
    campo("n_tabelas", lambda d: d["resumo"]["n_tabelas"])
    campo("n_imagens", lambda d: d["resumo"]["n_imagens"])
    campo("n_paragrafos", lambda d: d["resumo"]["n_paragrafos"])
    campo("titulos", lambda d: [t["texto"] for t in d["titulos"]])
    campo("estilos_de_titulo_em_uso", lambda d: sorted({t["estilo"] for t in d["titulos"]}))
    campo("niveis_de_titulo_em_uso", lambda d: sorted({t["nivel"] for t in d["titulos"]}))
    campo("tema_fontes", lambda d: {k: d["tema"].get(k) for k in
                                    ("esquema_fonte", "majorHAnsi", "minorHAnsi")})
    campo("tema_cores", lambda d: d["tema"].get("cores"))
    campo("formato_efetivo_Normal", lambda d: d["formato_efetivo"].get("Normal"))
    campo("formato_efetivo_titulo", lambda d: {k: v for k, v in d["formato_efetivo"].items()
                                               if k in ("Estilo1", "Ttulo1")})
    campo("formato_efetivo_Legenda", lambda d: d["formato_efetivo"].get("Legenda"))
    campo("numeracao_titulo", lambda d: [
        {"estilo": k, "numId": v["pPr_efetivo"].get("numeracao", {}).get("numId"),
         "abstract": d["numeracao"]["num_para_abstract"].get(
             v["pPr_efetivo"].get("numeracao", {}).get("numId")),
         "lvl0": d["numeracao"]["abstract"].get(
             d["numeracao"]["num_para_abstract"].get(
                 v["pPr_efetivo"].get("numeracao", {}).get("numId")), {})
             .get("niveis", {}).get("0")}
        for k, v in sorted(d["formato_efetivo"].items())
        if v["pPr_efetivo"].get("numeracao")])
    campo("estilos_tabela", lambda d: sorted({t["estilo"] for t in d["tabelas"]}))
    campo("assinaturas_tabela", lambda d: sorted(t["assinatura"] for t in d["tabelas"]))
    campo("texto_cabecalhos", lambda d: sorted(h.get("texto", "") for h in
                                               d["cabecalhos_rodapes"].values()
                                               if h.get("papel") == "header"))
    campo("texto_rodapes", lambda d: sorted(h.get("texto", "") for h in
                                            d["cabecalhos_rodapes"].values()
                                            if h.get("papel") == "footer"))
    campo("campos_rodape", lambda d: sorted({c["tipo"] for h in d["cabecalhos_rodapes"].values()
                                             if h.get("papel") == "footer"
                                             for c in h.get("campos", [])}))
    campo("larguras_imagem_cm", lambda d: sorted(collections.Counter(
        im.get("largura_cm") for im in d["imagens"]).most_common(), key=lambda kv: -kv[1])[:6])

    # tabelas recorrentes
    sig = collections.Counter()
    for d in docs:
        for s in {t["assinatura"] for t in d["tabelas"]}:
            sig[s] += 1
    c["tabelas_por_n_documentos"] = [{"assinatura": k, "n_docs": v}
                                     for k, v in sig.most_common()]
    # titulos recorrentes
    tsig = collections.Counter()
    for d in docs:
        for t in {U.norm(x["texto"]) for x in d["titulos"] if x["texto"]}:
            tsig[t] += 1
    c["titulos_por_n_documentos"] = [{"titulo": k, "n_docs": v} for k, v in tsig.most_common()]
    # entradas de sumario (estilo Sumario1/TOC1) recorrentes
    ssig = collections.Counter()
    for d in docs:
        ent = {U.norm(re.sub(r"\t.*$", "", p["texto"]))
               for p in d["paragrafos"]
               if (p.get("estilo") or "").lower().startswith(("sumrio", "sumario", "toc"))
               and p["texto"]}
        for e in ent:
            ssig[e] += 1
    c["entradas_sumario_por_n_documentos"] = [{"entrada": k, "n_docs": v}
                                              for k, v in ssig.most_common()]
    return c


def texto_comparativo(c):
    L = ["COMPARATIVO ENTRE OS %d DOCUMENTOS" % len(c["documentos"]),
         "docs: " + ", ".join(c["documentos"]), "",
         "=" * 78, "IGUAL EM TODOS", "=" * 78]
    for k, v in c["igual"].items():
        L.append("  %-28s = %s" % (k, json.dumps(v, ensure_ascii=False)[:600]))
    L += ["", "=" * 78, "VARIA", "=" * 78]
    for k, v in c["varia"].items():
        L.append("  %s:" % k)
        for doc, vv in v.items():
            L.append("      %-12s %s" % (doc, json.dumps(vv, ensure_ascii=False)[:500]))
    L += ["", "=" * 78, "TABELAS POR Nº DE DOCUMENTOS", "=" * 78]
    for t in c["tabelas_por_n_documentos"]:
        L.append("  %d/%d  %s" % (t["n_docs"], len(c["documentos"]), t["assinatura"]))
    L += ["", "=" * 78, "ENTRADAS DE SUMARIO POR Nº DE DOCUMENTOS", "=" * 78]
    for t in c["entradas_sumario_por_n_documentos"]:
        L.append("  %d/%d  %s" % (t["n_docs"], len(c["documentos"]), t["entrada"][:110]))
    return "\n".join(L)


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--obras", nargs="*", default=U.OBRAS)
    args = ap.parse_args()

    am = json.load(open(os.path.join(U.PROJ, "state", "amostra.json"), encoding="utf-8"))
    os.makedirs(U.assert_read_only(DEST), exist_ok=True)

    docs = []
    for obra in args.obras:
        for s in am["obras"][obra]["selecionados"]:
            if not s.get("docx"):
                print("%s n=%s: sem DOCX" % (obra, s["numero"]))
                continue
            rot = "%s_RO%03d" % (obra, s["numero"])
            info = extrair(s["docx"]["caminho"], rot)
            info.update({"obra": obra, "numero": s["numero"],
                         "data_inicio": s["data_inicio"], "data_fim": s["data_fim"],
                         "tamanho_bytes": s["docx"]["tamanho_bytes"]})
            with open(os.path.join(DEST, rot + ".json"), "w", encoding="utf-8") as fh:
                json.dump(info, fh, ensure_ascii=False, indent=1)
            with open(os.path.join(DEST, rot + ".outline.txt"), "w", encoding="utf-8") as fh:
                fh.write(outline(info))
            docs.append(info)
            r = info["resumo"]
            print("%-12s sec=%d par=%4d tab=%2d img=%3d campos=%2d midia=%2d OLE=%s | %s"
                  % (rot, r["n_secoes"], r["n_paragrafos"], r["n_tabelas"], r["n_imagens"],
                     r["n_campos"], r["n_midia"], r["tem_ole"],
                     ", ".join("%s(%d)" % kv for kv in r["orientacoes"])))

    if len(docs) > 1:
        c = comparar(docs)
        with open(os.path.join(DEST, "_comparativo.json"), "w", encoding="utf-8") as fh:
            json.dump(c, fh, ensure_ascii=False, indent=1)
        with open(os.path.join(DEST, "_comparativo.txt"), "w", encoding="utf-8") as fh:
            fh.write(texto_comparativo(c))
    print("OK ->", DEST)


if __name__ == "__main__":
    main()
