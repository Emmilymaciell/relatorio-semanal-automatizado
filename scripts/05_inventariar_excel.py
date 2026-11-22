# -*- coding: utf-8 -*-
r"""
05_inventariar_excel.py - disseca as 5 planilhas de 'Arquivos auxiliares' das
semanas da amostra (state/amostra.json).

Nao usa openpyxl para leitura: le o OOXML direto do zip (streaming), o que
permite tratar os RELATORIO FOTOGRAFICO de 70-170 MB sem carregar imagens.

REGRA: o SharePoint e READ-ONLY. Este script so le de la; escreve unicamente em
output/extracao/excel/ dentro do repositorio.

Uso:
    python scripts/05_inventariar_excel.py            # tudo
    python scripts/05_inventariar_excel.py --rapido   # pula RELATORIO FOTOGRAFICO
"""
import os
import re
import io
import sys
import json
import zipfile
import argparse
import datetime as dt
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

PROJ = U.PROJ
AMOSTRA = os.path.join(PROJ, "state", "amostra.json")
OUTDIR = os.path.join(U.OUT, "extracao", "excel")

# ---------------------------------------------------------------- namespaces
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKGREL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_DC = "http://purl.org/dc/elements/1.1/"
NS_DCTERMS = "http://purl.org/dc/terms/"
NS_EXTPROPS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS_CUSTPROPS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

MAX_CELULAS_POR_ABA = 40000
MAX_TXT = 300

# numFmtId built-in que representam data/hora
DATE_BUILTIN = set(range(14, 23)) | {27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
                                     45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58}
RE_DATEFMT = re.compile(r"(?<!\\)[ymdhs]", re.I)


def q(ns, tag):
    return "{%s}%s" % (ns, tag)


def txt(s, n=MAX_TXT):
    if s is None:
        return None
    s = re.sub(r"\s+", " ", str(s)).strip()
    return s if len(s) <= n else s[:n] + "…"


def col_letras(ref):
    m = re.match(r"([A-Z]+)", ref or "")
    return m.group(1) if m else ""


def col_num(letras):
    n = 0
    for c in letras:
        n = n * 26 + (ord(c) - 64)
    return n


def lin_num(ref):
    m = re.search(r"(\d+)$", ref or "")
    return int(m.group(1)) if m else 0


def serial_para_data(v, base1904=False):
    """Converte serial do Excel em ISO. None se implausivel."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if f < 1 or f > 80000:
        return None
    epoch = dt.datetime(1904, 1, 1) if base1904 else dt.datetime(1899, 12, 30)
    try:
        return (epoch + dt.timedelta(days=f)).isoformat(timespec="seconds")
    except OverflowError:
        return None


# ------------------------------------------------------------------ leitura
def ler(z, nome, limite=30 * 1024 * 1024):
    try:
        info = z.getinfo(nome)
    except KeyError:
        return None
    if info.file_size > limite:
        return None
    with z.open(nome) as fh:
        return fh.read()


def raiz(z, nome, limite=30 * 1024 * 1024):
    b = ler(z, nome, limite)
    if b is None:
        return None
    try:
        return ET.fromstring(b)
    except ET.ParseError:
        return None


# ------------------------------------------------------------ partes do xlsx
def props(z):
    out = {}
    r = raiz(z, "docProps/core.xml")
    if r is not None:
        mapa = {
            "titulo": q(NS_DC, "title"), "assunto": q(NS_DC, "subject"),
            "autor": q(NS_DC, "creator"), "ultimo_autor": q(NS_CP, "lastModifiedBy"),
            "criado": q(NS_DCTERMS, "created"), "modificado": q(NS_DCTERMS, "modified"),
            "categoria": q(NS_CP, "category"), "revisao": q(NS_CP, "revision"),
        }
        for k, tag in mapa.items():
            el = r.find(tag)
            if el is not None and el.text:
                out[k] = txt(el.text)
    r = raiz(z, "docProps/app.xml")
    if r is not None:
        for k, tag in [("aplicativo", "Application"), ("versao_app", "AppVersion"),
                       ("empresa", "Company"), ("gerente", "Manager")]:
            el = r.find(q(NS_EXTPROPS, tag))
            if el is not None and el.text:
                out[k] = txt(el.text)
        tp = r.find(q(NS_EXTPROPS, "TitlesOfParts"))
        if tp is not None:
            out["titulos_das_partes"] = [txt(e.text) for e in tp.iter()
                                         if e.tag.endswith("}lpstr") and e.text]
    r = raiz(z, "docProps/custom.xml")
    if r is not None:
        cust = {}
        for p in r:
            nome = p.get("name")
            val = None
            for ch in p:
                val = ch.text
            if nome:
                cust[nome] = txt(val)
        if cust:
            out["propriedades_customizadas"] = cust
    return out


def rels(z, parte):
    """Le os _rels de uma parte; devolve {rId: {tipo, alvo, modo}}."""
    d = os.path.dirname(parte)
    caminho = (d + "/" if d else "") + "_rels/" + os.path.basename(parte) + ".rels"
    r = raiz(z, caminho)
    out = {}
    if r is None:
        return out
    for rel in r.findall(q(NS_PKGREL, "Relationship")):
        out[rel.get("Id")] = {
            "tipo": (rel.get("Type") or "").rsplit("/", 1)[-1],
            "alvo": rel.get("Target"),
            "modo": rel.get("TargetMode"),
        }
    return out


def estilos(z):
    """Devolve (lista de numFmtId por xf, mapa numFmtId->codigo)."""
    r = raiz(z, "xl/styles.xml")
    if r is None:
        return [], {}
    fmts = {}
    nf = r.find(q(NS_MAIN, "numFmts"))
    if nf is not None:
        for f in nf.findall(q(NS_MAIN, "numFmt")):
            fmts[int(f.get("numFmtId"))] = f.get("formatCode")
    xfs = []
    cx = r.find(q(NS_MAIN, "cellXfs"))
    if cx is not None:
        for xf in cx.findall(q(NS_MAIN, "xf")):
            xfs.append(int(xf.get("numFmtId") or 0))
    return xfs, fmts


def eh_data_fmt(numfmt_id, codigo):
    if numfmt_id in DATE_BUILTIN:
        return True
    if codigo:
        limpo = re.sub(r'\[[^\]]*\]|"[^"]*"', "", codigo)
        return bool(RE_DATEFMT.search(limpo))
    return False


def shared_strings(z):
    b = ler(z, "xl/sharedStrings.xml", limite=60 * 1024 * 1024)
    if b is None:
        return []
    out = []
    for _, el in ET.iterparse(io.BytesIO(b), events=("end",)):
        if el.tag == q(NS_MAIN, "si"):
            out.append("".join(t.text or "" for t in el.iter(q(NS_MAIN, "t"))))
            el.clear()
    return out


def workbook(z):
    r = raiz(z, "xl/workbook.xml")
    if r is None:
        return {}
    wb = {"abas": [], "nomes_definidos": [], "referencias_externas": []}
    pr = r.find(q(NS_MAIN, "workbookPr"))
    wb["base_data_1904"] = bool(pr is not None and pr.get("date1904") in ("1", "true"))
    cp = r.find(q(NS_MAIN, "calcPr"))
    if cp is not None:
        wb["calc"] = dict(cp.attrib)
    sh = r.find(q(NS_MAIN, "sheets"))
    if sh is not None:
        for i, s in enumerate(sh.findall(q(NS_MAIN, "sheet"))):
            wb["abas"].append({
                "ordem": i + 1,
                "nome": s.get("name"),
                "sheetId": s.get("sheetId"),
                "rId": s.get(q(NS_REL, "id")),
                "estado": s.get("state") or "visible",
            })
    dn = r.find(q(NS_MAIN, "definedNames"))
    if dn is not None:
        for n in dn.findall(q(NS_MAIN, "definedName")):
            wb["nomes_definidos"].append({
                "nome": n.get("name"),
                "escopo_localSheetId": n.get("localSheetId"),
                "oculto": n.get("hidden"),
                "refere_a": txt(n.text, 500),
            })
    er = r.find(q(NS_MAIN, "externalReferences"))
    if er is not None:
        for e in er.findall(q(NS_MAIN, "externalReference")):
            wb["referencias_externas"].append(e.get(q(NS_REL, "id")))
    return wb


def links_externos(z):
    """xl/externalLinks/* -> arquivo-fonte + valores em cache. EVIDENCIA DE LINHAGEM."""
    saida = []
    alvos = [n for n in z.namelist()
             if n.startswith("xl/externalLinks/") and n.endswith(".xml")]
    for parte in sorted(alvos):
        item = {"parte": parte, "arquivo_alvo": None, "abas_referenciadas": [],
                "celulas_em_cache": [], "ddes": []}
        for rid, rel in rels(z, parte).items():
            if "externalLinkPath" in rel["tipo"] or rel["modo"] == "External":
                item["arquivo_alvo"] = rel["alvo"]
        r = raiz(z, parte)
        if r is not None:
            book = r.find(q(NS_MAIN, "externalBook"))
            if book is not None:
                sn = book.find(q(NS_MAIN, "sheetNames"))
                if sn is not None:
                    item["abas_referenciadas"] = [s.get("val") for s in sn]
                sd = book.find(q(NS_MAIN, "sheetDataSet"))
                if sd is not None:
                    for sheet in sd:
                        idx = sheet.get("sheetId")
                        for row in sheet:
                            for c in row:
                                v = c.find(q(NS_MAIN, "v"))
                                item["celulas_em_cache"].append({
                                    "aba_idx": idx, "ref": c.get("r"),
                                    "tipo": c.get("t"),
                                    "valor": txt(v.text) if v is not None else None,
                                })
            for d in r.iter(q(NS_MAIN, "ddeLink")):
                item["ddes"].append({"servico": d.get("ddeService"),
                                     "topico": d.get("ddeTopic")})
        item["celulas_em_cache"] = item["celulas_em_cache"][:400]
        saida.append(item)
    return saida


def aba(z, parte, sst, xfs, fmts, base1904, max_cel=MAX_CELULAS_POR_ABA):
    """Le uma worksheet em streaming. Devolve estrutura + celulas."""
    try:
        info = z.getinfo(parte)
    except KeyError:
        return None
    res = {
        "parte": parte, "bytes_xml": info.file_size, "dimensao": None,
        "linhas_com_dado": 0, "celulas_nao_vazias": 0, "truncado": False,
        "mescladas": [], "validacoes": [], "hyperlinks": [],
        "formatacao_condicional": 0, "celulas": [], "formulas": [],
        "datas": [], "oculta_linhas": 0, "oculta_colunas": 0,
        "config_impressao": {},
    }
    cels = []
    with z.open(parte) as fh:
        for ev, el in ET.iterparse(fh, events=("end",)):
            tag = el.tag
            if tag == q(NS_MAIN, "dimension"):
                res["dimensao"] = el.get("ref")
                el.clear()
            elif tag == q(NS_MAIN, "row"):
                res["linhas_com_dado"] += 1
                if el.get("hidden") in ("1", "true"):
                    res["oculta_linhas"] += 1
                for c in el.findall(q(NS_MAIN, "c")):
                    ref = c.get("r")
                    t = c.get("t")
                    s = c.get("s")
                    fel = c.find(q(NS_MAIN, "f"))
                    vel = c.find(q(NS_MAIN, "v"))
                    isel = c.find(q(NS_MAIN, "is"))
                    formula = fel.text if fel is not None else None
                    valor = None
                    if t == "s" and vel is not None and vel.text is not None:
                        try:
                            valor = sst[int(vel.text)]
                        except (ValueError, IndexError):
                            valor = None
                    elif t == "inlineStr" and isel is not None:
                        valor = "".join(x.text or "" for x in isel.iter(q(NS_MAIN, "t")))
                    elif vel is not None:
                        valor = vel.text
                    if valor is None and formula is None:
                        continue
                    res["celulas_nao_vazias"] += 1
                    if len(cels) >= max_cel:
                        res["truncado"] = True
                        continue
                    reg = {"ref": ref, "t": t, "v": txt(valor)}
                    if c.get("vm"):
                        # valor de metadados: marca de imagem-na-celula (richData)
                        reg["vm"] = c.get("vm")
                    if formula is not None:
                        reg["f"] = txt(formula, 500)
                        if fel.get("t"):
                            reg["f_tipo"] = fel.get("t")
                        if fel.get("ref"):
                            reg["f_ref"] = fel.get("ref")
                        if fel.get("si"):
                            reg["f_si"] = fel.get("si")
                    # data?
                    if t in (None, "n") and s is not None and valor is not None:
                        try:
                            nfid = xfs[int(s)]
                        except (ValueError, IndexError):
                            nfid = 0
                        if eh_data_fmt(nfid, fmts.get(nfid)):
                            iso = serial_para_data(valor, base1904)
                            if iso:
                                reg["data_iso"] = iso
                                res["datas"].append({"ref": ref, "iso": iso,
                                                     "serial": valor})
                    cels.append(reg)
                el.clear()
            elif tag == q(NS_MAIN, "mergeCell"):
                res["mescladas"].append(el.get("ref"))
                el.clear()
            elif tag == q(NS_MAIN, "dataValidation"):
                f1 = el.find(q(NS_MAIN, "formula1"))
                res["validacoes"].append({
                    "sqref": el.get("sqref"), "tipo": el.get("type"),
                    "formula1": txt(f1.text, 400) if f1 is not None else None,
                })
                el.clear()
            elif tag == q(NS_MAIN, "conditionalFormatting"):
                res["formatacao_condicional"] += 1
                el.clear()
            elif tag == q(NS_MAIN, "hyperlink"):
                res["hyperlinks"].append({"ref": el.get("ref"),
                                          "rId": el.get(q(NS_REL, "id")),
                                          "local": el.get("location")})
                el.clear()
            elif tag == q(NS_MAIN, "col"):
                if el.get("hidden") in ("1", "true"):
                    res["oculta_colunas"] += 1
                el.clear()
            elif tag == q(NS_MAIN, "pageSetup"):
                res["config_impressao"] = {k: v for k, v in el.attrib.items()
                                           if not k.startswith("{")}
                el.clear()
    res["celulas"] = cels
    res["mescladas"] = res["mescladas"][:400]
    res["datas"] = res["datas"][:200]
    vistas = {}
    for c in cels:
        if "f" in c:
            vistas.setdefault(c["f"], []).append(c["ref"])
    res["formulas"] = [{"formula": f, "n": len(r), "exemplos": r[:6]}
                       for f, r in sorted(vistas.items(), key=lambda x: -len(x[1]))]
    return res


def graficos(z):
    saida = []
    for parte in sorted(n for n in z.namelist()
                        if re.match(r"xl/charts/chart\d+\.xml$", n)):
        r = raiz(z, parte)
        if r is None:
            continue
        item = {"parte": parte, "titulo": None, "series": []}
        t = r.find(".//" + q(NS_CHART, "title"))
        if t is not None:
            item["titulo"] = txt("".join(x.text or "" for x in t.iter()
                                         if x.tag.endswith("}t")))
        tipos = set()
        for el in r.iter():
            if el.tag.startswith("{%s}" % NS_CHART) and el.tag.endswith("Chart"):
                tipos.add(el.tag.split("}")[1])
        item["tipos"] = sorted(tipos)
        for ser in r.iter(q(NS_CHART, "ser")):
            s = {}
            tx = ser.find(q(NS_CHART, "tx"))
            if tx is not None:
                fx = tx.find(".//" + q(NS_CHART, "f"))
                s["nome_ref"] = fx.text if fx is not None else None
                vs = [x.text for x in tx.iter(q(NS_CHART, "v")) if x.text]
                s["nome_cache"] = vs[-1] if vs else None
            cat = ser.find(q(NS_CHART, "cat"))
            if cat is not None:
                fx = cat.find(".//" + q(NS_CHART, "f"))
                s["categorias_ref"] = fx.text if fx is not None else None
                s["categorias_cache"] = [x.text for x in cat.iter(q(NS_CHART, "v"))][:60]
            val = ser.find(q(NS_CHART, "val"))
            if val is not None:
                fx = val.find(".//" + q(NS_CHART, "f"))
                s["valores_ref"] = fx.text if fx is not None else None
                s["valores_cache"] = [x.text for x in val.iter(q(NS_CHART, "v"))][:60]
            item["series"].append(s)
        item["rels"] = rels(z, parte)
        saida.append(item)
    return saida


def tabelas(z):
    saida = []
    for parte in sorted(n for n in z.namelist()
                        if re.match(r"xl/tables/table\d+\.xml$", n)):
        r = raiz(z, parte)
        if r is None:
            continue
        saida.append({
            "parte": parte, "nome": r.get("name"), "displayName": r.get("displayName"),
            "ref": r.get("ref"),
            "colunas": [c.get("name") for c in r.iter(q(NS_MAIN, "tableColumn"))],
        })
    return saida


def desenhos(z):
    saida = []
    for parte in sorted(n for n in z.namelist()
                        if re.match(r"xl/drawings/drawing\d+\.xml$", n)):
        info = z.getinfo(parte)
        item = {"parte": parte, "bytes_xml": info.file_size,
                "ancoras": 0, "imagens": 0, "formas": 0, "graficos": 0,
                "ancoras_exemplo": []}
        if info.file_size > 40 * 1024 * 1024:
            item["nao_analisado"] = "xml grande demais"
            saida.append(item)
            continue
        with z.open(parte) as fh:
            atual = None
            for ev, el in ET.iterparse(fh, events=("start", "end")):
                nome = el.tag.split("}")[-1]
                if ev == "start" and nome in ("twoCellAnchor", "oneCellAnchor",
                                              "absoluteAnchor"):
                    atual = {"tipo": nome, "de": None, "ate": None, "conteudo": None}
                elif ev == "end":
                    if nome in ("from", "to") and atual is not None:
                        c = el.find(q(NS_XDR, "col"))
                        l = el.find(q(NS_XDR, "row"))
                        v = "c%s,l%s" % (c.text if c is not None else "?",
                                         l.text if l is not None else "?")
                        atual["de" if nome == "from" else "ate"] = v
                    elif nome == "pic" and atual is not None:
                        atual["conteudo"] = "imagem"
                    elif nome == "sp" and atual is not None:
                        atual["conteudo"] = atual["conteudo"] or "forma"
                    elif nome == "graphicFrame" and atual is not None:
                        atual["conteudo"] = "grafico"
                    elif nome in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                        if atual:
                            item["ancoras"] += 1
                            if atual["conteudo"] == "imagem":
                                item["imagens"] += 1
                            elif atual["conteudo"] == "grafico":
                                item["graficos"] += 1
                            else:
                                item["formas"] += 1
                            if len(item["ancoras_exemplo"]) < 12:
                                item["ancoras_exemplo"].append(atual)
                        atual = None
                    el.clear()
        item["rels"] = rels(z, parte)
        saida.append(item)
    return saida


def midia(z):
    itens = []
    for n in z.namelist():
        if n.startswith("xl/media/"):
            i = z.getinfo(n)
            itens.append({"nome": os.path.basename(n), "bytes": i.file_size,
                          "bytes_comprimido": i.compress_size,
                          "ext": os.path.splitext(n)[1].lower()})
    itens.sort(key=lambda x: -x["bytes"])
    por_ext = {}
    for i in itens:
        e = por_ext.setdefault(i["ext"], {"n": 0, "bytes": 0})
        e["n"] += 1
        e["bytes"] += i["bytes"]
    return {
        "quantidade": len(itens),
        "bytes_total": sum(i["bytes"] for i in itens),
        "por_extensao": por_ext,
        "maiores": itens[:20],
        "todas": [{"nome": i["nome"], "bytes": i["bytes"]} for i in itens],
    }


def rich_data(z):
    """Imagens 'na celula' (Inserir > Imagem na celula). As fotos do RELATORIO
    FOTOGRAFICO ficam aqui, nao em xl/drawings/."""
    out = {"presente": False, "n_valores_ricos": 0, "n_rels_imagem": 0,
           "midia_referenciada": [], "tipos": []}
    if not any(n.startswith("xl/richData/") for n in z.namelist()):
        return out
    out["presente"] = True
    r = raiz(z, "xl/richData/rdrichvalue.xml")
    if r is not None:
        out["n_valores_ricos"] = len(list(r))
    rr = rels(z, "xl/richData/richValueRel.xml")
    alvos = [v["alvo"] for v in rr.values() if v["tipo"] == "image"]
    out["n_rels_imagem"] = len(alvos)
    out["midia_referenciada"] = [os.path.basename(a) for a in alvos]
    t = raiz(z, "xl/richData/rdRichValueTypes.xml")
    if t is not None:
        out["tipos"] = sorted({e.get("name") for e in t.iter() if e.get("name")})
    return out


def dissecar(caminho, com_celulas=True):
    st = U.stat_info(caminho)
    res = {
        "arquivo": os.path.basename(caminho),
        "caminho": caminho,
        "tamanho_bytes": st["tamanho_bytes"],
        "modificado": st["modificado"],
        "sha1": U.sha1(caminho) if st["tamanho_bytes"] < 5 * 1024 * 1024 else None,
    }
    with zipfile.ZipFile(U.lp(caminho)) as z:
        nomes = z.namelist()
        res["partes_zip"] = len(nomes)
        res["partes"] = sorted(n for n in nomes if not n.startswith("xl/media/"))
        res["propriedades"] = props(z)
        res["rels_workbook"] = rels(z, "xl/workbook.xml")
        wb = workbook(z)
        res["workbook"] = wb
        res["links_externos"] = links_externos(z)
        res["midia"] = midia(z)
        res["graficos"] = graficos(z)
        res["tabelas"] = tabelas(z)
        res["desenhos"] = desenhos(z)
        res["cadeia_calculo_presente"] = "xl/calcChain.xml" in nomes
        res["imagens_em_celula"] = rich_data(z)
        # abas
        sst = shared_strings(z)
        xfs, fmts = estilos(z)
        b1904 = wb.get("base_data_1904", False)
        res["abas"] = []
        rw = res["rels_workbook"]
        for a in wb.get("abas", []):
            rel = rw.get(a.get("rId") or "", {})
            alvo = rel.get("alvo") or ""
            parte = alvo if alvo.startswith("xl/") else "xl/" + alvo.lstrip("/")
            det = {"ordem": a["ordem"], "nome": a["nome"], "estado": a["estado"],
                   "parte": parte}
            if com_celulas:
                d = aba(z, parte, sst, xfs, fmts, b1904)
                if d:
                    det.update(d)
                    det["rels"] = rels(z, parte)
                    marcadas = [c["ref"] for c in d["celulas"] if "vm" in c]
                    det["celulas_com_imagem_na_celula"] = marcadas
            res["abas"].append(det)
    return res


# ----------------------------------------------------------------- amostra
def achar_aux(cam):
    for it in U.listdir(cam):
        p = os.path.join(cam, it)
        if U.isdir(p) and U.norm(it).startswith("arquivos aux"):
            return p
    return None


def classificar(nome):
    n = U.norm(nome)
    if "ficha de campo" in n:
        return "ficha_campo"
    if "grafico" in n and "avanco global" in n:
        return "grafico_avanco_global"
    if "fotografico" in n or "fotografico" in n.replace("ó", "o"):
        return "relatorio_fotografico"
    if n.startswith("unifilar") or " unifilar" in n:
        return "unifilar"
    if "programa" in n and "avanco global de obra" in n:
        return "prog_avanco_global"
    return "outro"


# ------------------------------------------------------------------ analise
def carregar(obra, semana, tipo):
    p = os.path.join(OUTDIR, "%s_S%02d_%s.json" % (obra, semana, tipo))
    if not os.path.exists(p):
        return None
    return json.load(open(p, encoding="utf-8"))


def mapa_celulas(doc, nome_aba, contem=False):
    for a in doc["abas"]:
        n = a.get("nome") or ""
        if (nome_aba in n) if contem else (n == nome_aba):
            return {c["ref"]: c for c in a.get("celulas", [])}, a
    return {}, None


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def analisar():
    """Confronta as planilhas entre si e entre semanas. Gera _analise.json."""
    dados = json.load(open(AMOSTRA, encoding="utf-8"))
    res = {"gerado_em": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
           "linhagem_grafico_vs_prog": [], "links_externos": [],
           "delta_semana_a_semana": [], "cabecalho_fotografico": []}

    for obra, info in dados["obras"].items():
        semanas = [s["numero"] for s in info["selecionados"]]
        for s in info["selecionados"]:
            n = s["numero"]
            g = carregar(obra, n, "grafico_avanco_global")
            c = carregar(obra, n, "prog_avanco_global")
            item = {"obra": obra, "semana": n,
                    "periodo": [s.get("data_inicio"), s.get("data_fim")]}
            if g:
                gc, _ = mapa_celulas(g, "Relat", contem=True)
                ga, _ = mapa_celulas(g, "Acumulado")
                linhas = sorted({lin_num(r) for r in gc if lin_num(r) >= 3})
                item["grafico_janela"] = [
                    {"linha": L, "data": (gc.get("A%d" % L) or {}).get("data_iso"),
                     "concluido_pct": num((gc.get("B%d" % L) or {}).get("v")),
                     "concl_atacado_pct": num((gc.get("C%d" % L) or {}).get("v"))}
                    for L in linhas]
                item["grafico_acumulado_n_linhas"] = len(
                    [r for r in ga if r.startswith("A") and lin_num(r) >= 2])
                item["grafico_tem_formula"] = any(
                    "f" in x for x in list(gc.values()) + list(ga.values()))
            if c:
                cc, _ = mapa_celulas(c, "Avanço Global - " + obra.replace("GO", "GO"))
                if not cc:
                    for a in c["abas"]:
                        if (a.get("nome") or "").replace(" ", "").endswith(obra):
                            cc = {x["ref"]: x for x in a.get("celulas", [])}
                item["prog_C3_concluido"] = num((cc.get("C3") or {}).get("v"))
                item["prog_C7_concl_atacado"] = num((cc.get("C7") or {}).get("v"))
                item["prog_C5_composto_concl"] = num((cc.get("C5") or {}).get("v"))
                item["prog_C9_composto_ca"] = num((cc.get("C9") or {}).get("v"))
                item["prog_C2_valor_contrato"] = num((cc.get("C2") or {}).get("v"))
                item["prog_links_externos"] = [l["arquivo_alvo"]
                                               for l in c["links_externos"]]
            # confronto
            if item.get("grafico_janela") and item.get("prog_C3_concluido") is not None:
                ult = item["grafico_janela"][-1]
                item["confronto"] = {
                    "grafico_ultima_linha_data": ult["data"],
                    "grafico_concluido": ult["concluido_pct"],
                    "prog_C3_x100": round(item["prog_C3_concluido"] * 100, 4),
                    "bate_concluido": ult["concluido_pct"] is not None and abs(
                        ult["concluido_pct"] - item["prog_C3_concluido"] * 100) < 0.005,
                    "grafico_concl_atacado": ult["concl_atacado_pct"],
                    "prog_C7_x100": round(item["prog_C7_concl_atacado"] * 100, 4)
                    if item.get("prog_C7_concl_atacado") is not None else None,
                    "bate_concl_atacado": (
                        ult["concl_atacado_pct"] is not None
                        and item.get("prog_C7_concl_atacado") is not None
                        and abs(ult["concl_atacado_pct"]
                                - item["prog_C7_concl_atacado"] * 100) < 0.005),
                }
            res["linhagem_grafico_vs_prog"].append(item)

            for tipo in ("ficha_campo", "unifilar", "prog_avanco_global",
                         "grafico_avanco_global", "relatorio_fotografico"):
                d = carregar(obra, n, tipo)
                if not d:
                    continue
                for l in d["links_externos"]:
                    res["links_externos"].append({
                        "obra": obra, "semana": n, "tipo": tipo,
                        "arquivo": d["arquivo"], "alvo": l["arquivo_alvo"],
                        "abas_alvo": l["abas_referenciadas"],
                        "n_celulas_cache": len(l["celulas_em_cache"])})
                if tipo == "relatorio_fotografico":
                    capa, _ = mapa_celulas(d, "CAPA")
                    res["cabecalho_fotografico"].append({
                        "obra": obra, "semana": n,
                        "AB1": (capa.get("AB1") or {}).get("v"),
                        "AB2_num_relatorio": (capa.get("AB2") or {}).get("v"),
                        "AB4_periodo": (capa.get("AB4") or {}).get("v"),
                        "A9_rodovia": (capa.get("A9") or {}).get("v"),
                        "S9_contrato": (capa.get("S9") or {}).get("v"),
                        "A12_construtora": (capa.get("A12") or {}).get("v"),
                        "S12_supervisora": (capa.get("S12") or {}).get("v"),
                        "A21_obra": (capa.get("A21") or {}).get("v"),
                        "A24_trecho": (capa.get("A24") or {}).get("v"),
                        "A27_extensao": (capa.get("A27") or {}).get("v"),
                        "A30_SRE": (capa.get("A30") or {}).get("v"),
                        "A58_mes": (capa.get("A58") or {}).get("v"),
                        "n_fotos_na_celula": d["imagens_em_celula"]["n_rels_imagem"],
                        "n_midia": d["midia"]["quantidade"],
                        "bytes_midia": d["midia"]["bytes_total"],
                        "abas": [{"nome": a["nome"], "estado": a["estado"],
                                  "fotos": len(a.get(
                                      "celulas_com_imagem_na_celula", []))}
                                 for a in d["abas"]],
                    })

        # delta semana N -> N+1
        if len(semanas) == 2:
            a_, b_ = sorted(semanas)
            for tipo in ("ficha_campo", "unifilar", "prog_avanco_global",
                         "grafico_avanco_global", "relatorio_fotografico"):
                da, db = carregar(obra, a_, tipo), carregar(obra, b_, tipo)
                if not da or not db:
                    continue
                d = {"obra": obra, "tipo": tipo, "de": a_, "para": b_,
                     "bytes": [da["tamanho_bytes"], db["tamanho_bytes"]],
                     "sha1_igual": bool(da.get("sha1")) and da["sha1"] == db["sha1"],
                     "abas_de": [x["nome"] for x in da["abas"]],
                     "abas_para": [x["nome"] for x in db["abas"]],
                     "por_aba": []}
                mb = {x["nome"]: x for x in db["abas"]}
                for aa in da["abas"]:
                    ab = mb.get(aa["nome"])
                    if not ab:
                        continue
                    ca = {x["ref"]: x.get("v") for x in aa.get("celulas", [])}
                    cb = {x["ref"]: x.get("v") for x in ab.get("celulas", [])}
                    mud = [r for r in set(ca) & set(cb) if ca[r] != cb[r]]
                    d["por_aba"].append({
                        "aba": aa["nome"],
                        "dim": [aa.get("dimensao"), ab.get("dimensao")],
                        "celulas": [aa.get("celulas_nao_vazias"),
                                    ab.get("celulas_nao_vazias")],
                        "so_em_N": len(set(ca) - set(cb)),
                        "so_em_N1": len(set(cb) - set(ca)),
                        "alteradas": len(mud),
                        "exemplos": sorted(
                            mud, key=lambda r: (col_num(col_letras(r)), lin_num(r))
                        )[:15],
                    })
                res["delta_semana_a_semana"].append(d)

    # o cache do link externo do [Programa] bate com a copia de UNIFILAR guardada
    # na propria pasta 'Arquivos auxiliares'? Se bater, sao a mesma versao.
    res["cache_link_x_unifilar_aux"] = []
    for obra, info in dados["obras"].items():
        for s in info["selecionados"]:
            n = s["numero"]
            c = carregar(obra, n, "prog_avanco_global")
            u = carregar(obra, n, "unifilar")
            if not c or not u:
                continue
            marca = obra.replace("GO", "")
            alvo = None
            for l in c["links_externos"]:
                if ("UNIFILAR%20" + marca) in (l["arquivo_alvo"] or ""):
                    alvo = l
            if not alvo:
                continue
            ok = div = 0
            exemplos = []
            for cc in alvo["celulas_em_cache"]:
                try:
                    nome_aba = alvo["abas_referenciadas"][int(cc["aba_idx"])]
                except (ValueError, IndexError):
                    continue
                mp, _ = mapa_celulas(u, nome_aba)
                real = (mp.get(cc["ref"]) or {}).get("v")
                a_, b_ = num(real), num(cc["valor"])
                igual = (abs(a_ - b_) < 1e-9) if (a_ is not None and b_ is not None) \
                    else (str(real) == str(cc["valor"]))
                if igual:
                    ok += 1
                else:
                    div += 1
                    if len(exemplos) < 6:
                        exemplos.append({"aba": nome_aba, "ref": cc["ref"],
                                         "cache": cc["valor"], "aux": real})
            res["cache_link_x_unifilar_aux"].append({
                "obra": obra, "semana": n, "alvo": alvo["arquivo_alvo"],
                "iguais": ok, "divergentes": div, "exemplos": exemplos})

    p = os.path.join(OUTDIR, "_analise.json")
    U.assert_read_only(p)
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(res, fh, ensure_ascii=False, indent=1)
    print("OK ->", p)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rapido", action="store_true",
                    help="pula os RELATORIO FOTOGRAFICO (70-170 MB)")
    ap.add_argument("--so", default=None, help="filtra por tipo")
    ap.add_argument("--obra", default=None)
    ap.add_argument("--analisar", action="store_true",
                    help="so confronta os JSON ja extraidos (nao le o SharePoint)")
    args = ap.parse_args()

    U.assert_read_only(OUTDIR)
    os.makedirs(OUTDIR, exist_ok=True)
    if args.analisar:
        analisar()
        return
    dados = json.load(open(AMOSTRA, encoding="utf-8"))

    resumo = {"gerado_em": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
              "itens": []}
    for obra, info in dados["obras"].items():
        if args.obra and obra != args.obra:
            continue
        for sem in info["selecionados"]:
            aux = achar_aux(sem["caminho"])
            if not aux:
                print("[!] sem 'Arquivos auxiliares':", obra, sem["numero"])
                continue
            for nome in U.listdir(aux):
                fp = os.path.join(aux, nome)
                if not nome.lower().endswith((".xlsx", ".xlsm")):
                    continue
                tipo = classificar(nome)
                if args.so and tipo != args.so:
                    continue
                if args.rapido and tipo == "relatorio_fotografico":
                    continue
                saida = os.path.join(
                    OUTDIR, "%s_S%02d_%s.json" % (obra, sem["numero"], tipo))
                print("->", obra, sem["numero"], tipo, nome, flush=True)
                try:
                    d = dissecar(fp, com_celulas=True)
                except Exception as e:  # noqa: BLE001
                    print("   ERRO:", type(e).__name__, e, flush=True)
                    continue
                d["obra"] = obra
                d["semana"] = sem["numero"]
                d["periodo"] = {"inicio": sem.get("data_inicio"),
                                "fim": sem.get("data_fim")}
                d["tipo"] = tipo
                d["pasta_auxiliares"] = aux
                U.assert_read_only(saida)
                with open(saida, "w", encoding="utf-8") as fh:
                    json.dump(d, fh, ensure_ascii=False, indent=1)
                resumo["itens"].append({
                    "obra": obra, "semana": sem["numero"], "tipo": tipo,
                    "arquivo": nome, "json": os.path.basename(saida),
                    "tamanho_bytes": d["tamanho_bytes"],
                    "modificado": d["modificado"],
                    "sha1": d.get("sha1"),
                    "abas": [{"nome": a["nome"], "estado": a["estado"],
                              "dimensao": a.get("dimensao"),
                              "celulas": a.get("celulas_nao_vazias"),
                              "n_formulas": len(a.get("formulas", [])),
                              "n_fotos_na_celula":
                                  len(a.get("celulas_com_imagem_na_celula", []))}
                             for a in d["abas"]],
                    "imagens_na_celula": d["imagens_em_celula"]["n_rels_imagem"],
                    "n_links_externos": len(d["links_externos"]),
                    "n_graficos": len(d["graficos"]),
                    "n_imagens": d["midia"]["quantidade"],
                    "bytes_imagens": d["midia"]["bytes_total"],
                    "n_nomes_definidos": len(d["workbook"].get("nomes_definidos", [])),
                    "propriedades": d["propriedades"],
                })

    rp = os.path.join(OUTDIR, "_resumo.json")
    antigo = {}
    if os.path.exists(rp):
        try:
            antigo = json.load(open(rp, encoding="utf-8"))
        except Exception:  # noqa: BLE001
            antigo = {}
    chave = lambda i: (i["obra"], i["semana"], i["tipo"])  # noqa: E731
    mescla = {chave(i): i for i in antigo.get("itens", [])}
    mescla.update({chave(i): i for i in resumo["itens"]})
    resumo["itens"] = [mescla[k] for k in sorted(mescla)]
    with open(rp, "w", encoding="utf-8") as fh:
        json.dump(resumo, fh, ensure_ascii=False, indent=1)
    print("OK ->", rp)


if __name__ == "__main__":
    main()
