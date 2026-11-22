# -*- coding: utf-8 -*-
r"""
10_analisar_nao_conformidades.py - disseca a planilha consolidada de Nao
Conformidades e cruza com a secao 8 dos Relatorios Semanais.

Fonte consolidada (CONFIRMADA PELA EMMILY):
    <SHAREPOINT>\Programa - 3. Desenvolvimento\G2 -  Engenharia\
        02_Acompanhamento de NC automatizada.xlsm

O que este script faz:
  1. Disseca o OOXML do .xlsm: abas (visiveis/ocultas), tabelas estruturadas,
     colunas, formulas, validacoes, nomes definidos, tabelas dinamicas,
     segmentacoes (slicers), graficos, customXml/PowerQuery e nomes de modulos
     e procedimentos do vbaProject.bin (SEM executar macro nenhuma).
  2. Le as duas tabelas mestras (Tbl_VISAO e Tbl_BASE) linha a linha.
  3. Extrai os quantitativos por status das 8 secoes 8.3 da amostra
     (output/comparacoes/estrutura_relatorios.json) e confronta com a planilha,
     tanto no estado atual quanto reconstruido na data de fechamento.
  4. Extrai as NCs citadas na secao 8.1, localiza o PDF original em
     '3. Notificacoes\Not. Geradas', extrai o texto e compara com a coluna
     Irregularidade da planilha e com a frase do relatorio.
  5. Verifica a regra "toda NC aberta dentro do periodo aparece na 8.1".

REGRA: o SharePoint e READ-ONLY. So le de la; escreve unicamente em
output/evidencias/ e tmp/ dentro do repositorio.

Uso:
    python scripts/10_analisar_nao_conformidades.py
    python scripts/10_analisar_nao_conformidades.py --sem-pdf
"""
import os
import re
import io
import sys
import json
import zipfile
import argparse
import datetime as dt
import collections
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

import warnings
warnings.filterwarnings("ignore")
import openpyxl
from openpyxl.utils import get_column_letter

PLANILHA = os.path.join(U.BASE_OBRAS, "02_Acompanhamento de NC automatizada.xlsm")
ESTRUTURA = os.path.join(U.OUT, "comparacoes", "estrutura_relatorios.json")
OUTDIR = os.path.join(U.OUT, "evidencias")
SAIDA = os.path.join(OUTDIR, "nao_conformidades.json")

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NSR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NSC = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
NSA = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

# ------------------------------------------------------------------ dominios
# LISTAS_VALIDACAO!I1:I7 - "Status da nao conformidade"
STATUS_OFICIAIS = ["Aberta", "Em Analise", "Aguardando Evidencias", "Encerrada",
                   "Aguardando Projeto Executivo", "Cancelada"]
STATUS_CANON = {
    "aberta": "Aberta",
    "em analise": "Em Análise",
    "aguardando evidencias": "Aguardando Evidências",
    "encerrada": "Encerrada",
    "aguardando projeto executivo": "Aguardando Projeto Executivo",
    "cancelada": "Cancelada",
}
# rotulo usado no texto da secao 8.3 -> status canonico da planilha
ROTULO_83 = {
    "em aberto": "Aberta",
    "aguardando projeto executivo": "Aguardando Projeto Executivo",
    "aguardando evidencias": "Aguardando Evidências",
    "em analise": "Em Análise",
    "canceladas": "Cancelada",
    "encerradas": "Encerrada",
}
# obras que agregam "Em Analise" dentro de "em aberto" na redacao da 8.3
AGREGA_EM_ANALISE = {"RD-101", "RD-104"}

PASTA_NOTIF = "3. Notificações"


def canon_status(v):
    return STATUS_CANON.get(U.norm(v), str(v).strip() if v is not None else "")


def iso(v):
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, dt.date):
        return v.isoformat()
    return v


def as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return None


def jsonable(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return iso(v)
    if v is None or isinstance(v, (str, int, float, bool)):
        return v
    return str(v)


# =========================================================== 1. anatomia OOXML
def abrir_zip(caminho):
    """OneDrive pode falhar no zipfile (OSError 22): le para BytesIO antes."""
    with open(U.lp(caminho), "rb") as fh:
        return zipfile.ZipFile(io.BytesIO(fh.read()))


def relmap(z, parte):
    d, b = os.path.dirname(parte), os.path.basename(parte)
    try:
        x = ET.fromstring(z.read(d + "/_rels/" + b + ".rels"))
    except KeyError:
        return {}
    return {r.get("Id"): r.get("Target") for r in x}


def anatomia(z):
    out = {}
    wbrels = relmap(z, "xl/workbook.xml")
    wb = ET.fromstring(z.read("xl/workbook.xml"))

    abas = []
    for sh in wb.find(NS + "sheets"):
        alvo = wbrels[sh.get(NSR + "id")]
        abas.append({"nome": sh.get("name"), "sheetId": sh.get("sheetId"),
                     "estado": sh.get("state") or "visible",
                     "parte": alvo.split("/")[-1]})
    out["abas"] = abas

    nomes = []
    for dn in (wb.find(NS + "definedNames") or []):
        nomes.append({"nome": dn.get("name"), "escopo_local": dn.get("localSheetId"),
                      "oculto": dn.get("hidden") == "1", "valor": (dn.text or "")[:200]})
    out["nomes_definidos"] = nomes

    # tabelas estruturadas
    tabs = []
    for n in sorted(x for x in z.namelist() if re.match(r"xl/tables/table\d+\.xml$", x)):
        t = ET.fromstring(z.read(n))
        af = t.find(NS + "autoFilter")
        filtros = []
        if af is not None:
            for fc in af.findall(NS + "filterColumn"):
                vals = [f.get("val") for f in fc.iter(NS + "filter")]
                filtros.append({"colId": fc.get("colId"), "valores": vals})
        tabs.append({
            "parte": n.split("/")[-1], "nome": t.get("displayName"), "ref": t.get("ref"),
            "colunas": [c.get("name") for c in t.find(NS + "tableColumns")],
            "autofiltro_salvo": filtros,
        })
    out["tabelas_estruturadas"] = tabs

    # em qual aba cada tabela / pivot / slicer / drawing mora
    por_aba = {}
    for a in abas:
        rm = relmap(z, "xl/worksheets/" + a["parte"])
        alvos = [v.split("/")[-1] for v in rm.values()]
        xml = z.read("xl/worksheets/" + a["parte"]).decode("utf8", "replace")
        sl = [rm.get(m) for m in re.findall(r'<x14:slicer[^>]*r:id="([^"]+)"', xml)]
        por_aba[a["nome"]] = {
            "tabelas": [x for x in alvos if x.startswith("table")],
            "pivots": [x for x in alvos if x.startswith("pivotTable")],
            "drawings": [x for x in alvos if x.startswith("drawing")],
            "slicers": [str(x).split("/")[-1] for x in sl if x],
        }
    out["objetos_por_aba"] = por_aba

    # tabelas dinamicas
    pivots = []
    for n in sorted(x for x in z.namelist() if re.match(r"xl/pivotTables/pivotTable\d+\.xml$", x)):
        p = ET.fromstring(z.read(n))
        loc = p.find(NS + "location")
        campos = p.find(NS + "pivotFields")
        rf = p.find(NS + "rowFields")
        cf = p.find(NS + "colFields")
        df = p.find(NS + "dataFields")
        pivots.append({
            "parte": n.split("/")[-1], "nome": p.get("name"), "cacheId": p.get("cacheId"),
            "destino": loc.get("ref") if loc is not None else None,
            "n_campos": len(campos) if campos is not None else 0,
            "rowFields": [e.get("x") for e in (rf if rf is not None else [])],
            "colFields": [e.get("x") for e in (cf if cf is not None else [])],
            "dataFields": [{"nome": e.get("name"), "fld": e.get("fld"),
                            "subtotal": e.get("subtotal")} for e in (df if df is not None else [])],
        })
    out["tabelas_dinamicas"] = pivots

    # cache da dinamica (origem + campos)
    caches = []
    for n in sorted(x for x in z.namelist() if re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml$", x)):
        c = ET.fromstring(z.read(n))
        src = c.find(NS + "cacheSource")
        ws = src.find(NS + "worksheetSource") if src is not None else None
        campos = []
        for i, cfd in enumerate(c.find(NS + "cacheFields")):
            si = cfd.find(NS + "sharedItems")
            itens = [e.get("v") for e in (si if si is not None else []) if e.get("v") is not None]
            campos.append({"indice": i, "nome": cfd.get("name"),
                           "itens_cacheados": itens[:12]})
        caches.append({"parte": n.split("/")[-1],
                       "tipo_origem": src.get("type") if src is not None else None,
                       "origem": ws.get("name") if ws is not None else None,
                       "campos": campos})
    out["pivot_caches"] = caches

    # segmentacoes de dados
    slicers = []
    for n in sorted(x for x in z.namelist() if re.match(r"xl/slicerCaches/slicerCache\d+\.xml$", x)):
        s = z.read(n).decode("utf8", "replace")
        m1 = re.search(r'\sname="([^"]+)"', s)
        m2 = re.search(r'sourceName="([^"]+)"', s)
        alvo_tab = re.search(r'<x15:tableSlicerCache[^>]*tableId="(\d+)"', s)
        slicers.append({"parte": n.split("/")[-1],
                        "nome": m1.group(1) if m1 else None,
                        "campo_origem": m2.group(1) if m2 else None,
                        "tipo": "tabela" if alvo_tab else "dinamica",
                        "pivots_ligados": re.findall(r'<pivotTable[^>]*name="([^"]+)"', s)})
    out["segmentacoes"] = slicers

    # graficos
    graficos = []
    for n in sorted(x for x in z.namelist() if re.match(r"xl/charts/chart\d+\.xml$", x)):
        c = ET.fromstring(z.read(n))
        titulo = " ".join((t.text or "") for t in c.iter(NSA + "t")).strip()
        tipos = sorted(set(e.tag.replace(NSC, "") for e in c.iter()
                           if e.tag.startswith(NSC) and e.tag.endswith("Chart")))
        refs = sorted(set(e.text for e in c.iter(NSC + "f") if e.text))
        cache = [e.text for e in c.iter(NSC + "v")][:20]
        graficos.append({"parte": n.split("/")[-1], "titulo": titulo[:120],
                         "tipos": tipos, "referencias": refs, "valores_cacheados": cache})
    out["graficos"] = graficos

    # conexoes externas / PowerQuery
    out["connections_xml"] = "xl/connections.xml" in z.namelist()
    cx = []
    for n in sorted(x for x in z.namelist() if x.startswith("customXml/item")):
        raw = z.read(n).decode("utf8", "replace")
        cx.append({"parte": n, "bytes": len(raw), "raiz": re.sub(r"\s+", " ", raw[:160]),
                   "menciona_powerquery": bool(re.search(r"(?i)DataMashup|query|Mashup", raw))})
    out["custom_xml"] = cx
    out["externalLinks"] = [x for x in z.namelist() if "externalLink" in x]

    # VBA - apenas leitura de strings, NENHUMA macro e executada
    vba = {"presente": "xl/vbaProject.bin" in z.namelist()}
    if vba["presente"]:
        b = z.read("xl/vbaProject.bin")
        vba["bytes"] = len(b)
        strs = set(s.decode("latin1") for s in re.findall(rb"[\x20-\x7E\xC0-\xFF]{4,60}", b))
        vba["modulos"] = sorted(s.split("=", 1)[1] for s in strs if s.startswith("Module="))
        alvo = re.compile(r"^(Gravar_|Imprimir_|Salvar_|Atualizar_|Limpar_|Aplicar|Conferir|"
                          r"Reconstruir_|Proxima|Garantir|Get|Format|Lista|Is|Safe|Aviso|Avisar|Macro_)")
        vba["procedimentos_inferidos"] = sorted(
            s for s in strs if alvo.match(s) and re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", s))
        vba["mensagens"] = sorted(s for s in strs
                                  if len(s) > 22 and " " in s and not s.startswith("Module="))[:60]
    out["vba"] = vba
    return out


# ======================================================= 2. leitura das tabelas
def ler_tabela(wv, wf, aba, ref):
    """Le uma tabela estruturada (ref tipo 'A3:AL132') devolvendo linhas + formulas."""
    m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
    c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
    from openpyxl.utils import column_index_from_string as cix
    ci1, ci2 = cix(c1), cix(c2)
    sv, sf = wv[aba], wf[aba]
    cols = [sv.cell(r1, c).value for c in range(ci1, ci2 + 1)]
    linhas, formulas = [], {}
    for r in range(r1 + 1, r2 + 1):
        d = {"_aba": aba, "_linha": r}
        vazia = True
        for i, c in enumerate(range(ci1, ci2 + 1)):
            v = sv.cell(r, c).value
            f = sf.cell(r, c).value
            d[cols[i]] = jsonable(v)
            if v not in (None, ""):
                vazia = False
            if isinstance(f, str) and f.startswith("="):
                chave = cols[i]
                formulas.setdefault(chave, {"coluna": chave,
                                            "celula_exemplo": "%s%d" % (get_column_letter(c), r),
                                            "formula": re.sub(r"\s+", " ", f), "ocorrencias": 0})
                formulas[chave]["ocorrencias"] += 1
        if not vazia:
            linhas.append(d)
    return cols, linhas, list(formulas.values())


def ler_listas_validacao(wv):
    s = wv["LISTAS_VALIDACAO"]
    listas = {}
    for c in range(1, s.max_column + 1):
        titulo = s.cell(1, c).value
        if not titulo:
            continue
        vals = []
        for r in range(2, s.max_row + 1):
            v = s.cell(r, c).value
            if v not in (None, ""):
                vals.append({"celula": "%s%d" % (get_column_letter(c), r), "valor": v})
        listas[str(titulo).strip()] = {"celula_titulo": "%s1" % get_column_letter(c),
                                       "valores": vals}
    return listas


# ===================================================== 3. secao 8.3 dos relatorios
RE_G2 = re.compile(r"Gr[áa]fico\s*2\s*apresenta a distribui[çc][ãa]o de n[ãa]o conformidades por status,\s*sendo\s*(.+?)\.\s*\n",
                   re.S | re.I)
RE_ITEM = re.compile(r"(\d+)\s+((?:em aberto|aguardando projeto executivo|aguardando evid[êe]ncias|"
                     r"em an[áa]lise|canceladas?|encerradas?))", re.I)
RE_ATRASO = re.compile(r"encontra-se\s*(?:em atraso\s*)?([a-zçãéêíóôõúâ\s]+?)\s*\((\d+)\)\s*"
                       r"(?:not|n[ãa]o)[^.]*?e\s*([a-zçãéêíóôõúâ\s]+?)\s*\((\d+)\)", re.I | re.S)


def extrair_83(texto):
    t = re.sub(r"\s+", " ", texto)
    m = re.search(r"Gr[áa]fico 2 apresenta a distribui[çc][ãa]o de n[ãa]o conformidades por status, sendo (.+?)\. Gr[áa]fico 2", t, re.I)
    if not m:
        m = re.search(r"por status, sendo (.+?)\.", t, re.I)
    if not m:
        return None, {}
    frase = m.group(1).strip()
    cont = {}
    for n, rot in RE_ITEM.findall(frase):
        cont[ROTULO_83[U.norm(rot).rstrip("s") + ("s" if U.norm(rot).endswith("s") else "")]
             if False else ROTULO_83.get(U.norm(rot), U.norm(rot))] = int(n)
    return frase, cont


def extrair_atraso(texto):
    t = re.sub(r"\s+", " ", texto)
    m = re.search(r"encontra-se\s*(?:em atraso\s*)?[^()]*?\((\d+)\)[^()]*?\((\d+)\)", t)
    if m:
        return {"em_atraso": int(m.group(1)), "encerradas": int(m.group(2))}
    return None


def contar_status(linhas, rodovia, col_rod, col_st, col_ab, col_enc, corte=None):
    """Conta status por rodovia. Com `corte`, reconstroi o estado naquela data:
    ignora NC aberta depois do corte e 'desencerra' as fechadas depois do corte."""
    c = collections.Counter()
    for r in linhas:
        if str(r.get(col_rod) or "").strip() != rodovia:
            continue
        if corte:
            ab = r.get(col_ab)
            ab = dt.date.fromisoformat(ab) if isinstance(ab, str) and re.match(r"\d{4}-\d\d-\d\d$", ab) else None
            if ab and ab > corte:
                continue
        st = canon_status(r.get(col_st))
        if corte and st == "Encerrada":
            en = r.get(col_enc)
            en = dt.date.fromisoformat(en) if isinstance(en, str) and re.match(r"\d{4}-\d\d-\d\d$", en) else None
            if en and en > corte:
                st = "(aberta na data de corte)"
        c[st] += 1
    return dict(c)


# =============================================== 4. NCs citadas na secao 8.1
RE_NC_81 = re.compile(r"(NC-\d+)\s*:\s*Emitida em\s*(\d{2}/\d{2}/\d{4})\s*,\s*(.+?)(?=\n\s*(?:•|NC-\d+:)|\Z)", re.S)


def extrair_81(texto):
    itens = []
    for m in RE_NC_81.finditer(texto):
        frase = re.sub(r"\s+", " ", m.group(3)).strip()
        itens.append({"nc": m.group(1), "data_citada": m.group(2), "descricao": frase})
    houve = "não houve" not in U.norm(texto).replace("nao houve", "não houve") \
        if False else ("nao houve" not in U.norm(texto))
    return {"houve_emissao": bool(itens) or houve, "itens": itens,
            "texto": re.sub(r"\s+", " ", texto).strip()}


def indexar_pdfs_nc(obra):
    """Mapeia NC-NNN -> caminho do PDF 'Nao Conformidade' em 3. Notificacoes."""
    raiz = os.path.join(U.BASE_OBRAS, obra, "3_Gestao Tecnica de Obra", PASTA_NOTIF)
    idx = collections.defaultdict(list)
    if not U.exists(raiz):
        return idx
    for dp, _dn, fn in U.walk(raiz):
        for f in fn:
            m = re.match(r"(NC-\d+)", f)
            if not m:
                continue
            idx[m.group(1)].append({"arquivo": os.path.join(dp, f),
                                    "eh_notificacao": bool(re.search(r"(?i)n[ãa]o\s*conformidade|notifica[çc][ãa]o de n[ãa]o", f)),
                                    "tamanho": U.stat_info(os.path.join(dp, f))["tamanho_bytes"],
                                    "modificado": U.stat_info(os.path.join(dp, f))["modificado"]})
    return idx


def texto_pdf(caminho):
    import fitz
    loc = U.copiar_para_tmp(caminho, "nc_docs", U.slug(os.path.basename(caminho))[:80] + ".pdf")
    d = fitz.open(loc)
    meta = {k: v for k, v in d.metadata.items() if v}
    pags = [pg.get_text() for pg in d]
    d.close()
    return meta, pags


CAMPOS_FORM = ["Natureza", "Criticidade", "Recorrência", "Recomendação",
               "Empresa notificada", "Contrato", "Rodovia"]


def campos_do_pdf(pags):
    t = "\n".join(pags)
    out = {}
    for c in CAMPOS_FORM:
        m = re.search(re.escape(c) + r":?\s*\n\s*([^\n]+)", t)
        if m:
            out[c] = m.group(1).strip()
    m = re.search(r"(NC-\d+)\s*\n\s*Notifica[çc][ãa]o de N[ãa]o Conformidade\s*\n\s*Versao\s*(\d+)\s*\n\s*(\d{2}/\d{2}/\d{4})", t)
    if m:
        out["_cabecalho"] = {"nc": m.group(1), "versao": m.group(2), "data": m.group(3)}
    m = re.search(r"Irregularidades constatadas:\s*(.*)", t, re.S)
    return out


def primeira_frase(txt):
    t = re.sub(r"\s+", " ", (txt or "")).strip()
    m = re.match(r"(.+?[.])(\s|$)", t)
    return m.group(1) if m else t


def frase_82_esperada(irregularidade, nc, data_br):
    """Template inferido da secao 8.1 (ver docs/mapeamento/NC_TEXTO_LINEAGE.md)."""
    f = primeira_frase(irregularidade)
    # remove a clausula temporal inicial ou interna
    f2 = re.sub(r"^Em inspe[çc][ãa]o realizada em \d{2}/\d{2}/\d{4},\s*", "", f)
    f2 = re.sub(r",?\s*realizada em \d{2}/\d{2}/\d{4},?", ",", f2)
    f2 = re.sub(r"\s+,", ",", f2).strip()
    if f2:
        f2 = f2[0].lower() + f2[1:]
    return "%s: Emitida em %s, %s" % (nc, data_br, f2)


def sim(a, b):
    """Similaridade por tokens (Jaccard) entre dois textos."""
    A = set(re.findall(r"\w+", U.norm(a)))
    B = set(re.findall(r"\w+", U.norm(b)))
    return round(len(A & B) / max(1, len(A | B)), 3)


def contido(trecho, corpo):
    """Fracao das frases de `trecho` que aparecem literalmente em `corpo`.

    Necessario porque o texto extraido do PDF sai fora de ordem (o PDF e a
    impressao de uma planilha, nao de um fluxo de texto): comparar o bloco
    inteiro daria falso negativo. Frase a frase o casamento e exato.
    """
    frases = [f for f in re.split(r"(?<=[.])\s+", re.sub(r"\s+", " ", trecho or "")) if len(f) > 25]
    if not frases:
        return None
    alvo = U.norm(corpo)
    ok = [f for f in frases if U.norm(f) in alvo]
    return {"frases": len(frases), "encontradas_literalmente": len(ok),
            "fracao": round(len(ok) / len(frases), 3),
            "nao_encontradas": [f[:160] for f in frases if f not in ok][:4]}


def conferir_campos(linha, corpo):
    """Cada campo estruturado da planilha aparece literalmente no PDF?"""
    corpo_n = U.norm(corpo)
    out = {}
    for c in ("Empresa ", "Rodovia", "Natureza", "Criticidade"):
        v = (linha or {}).get(c)
        if v:
            out[c.strip()] = {"planilha": v, "no_pdf": U.norm(v) in corpo_n}
    return out


# =================================================================== principal
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sem-pdf", action="store_true", help="pula a extracao dos PDFs das NCs")
    args = ap.parse_args()

    U.assert_read_only(OUTDIR)
    os.makedirs(OUTDIR, exist_ok=True)

    ev = {"gerado_em": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
          "script": os.path.basename(__file__)}

    # -------------------------------------------------------------- planilha
    print("[1/5] dissecando a planilha consolidada...")
    st = U.stat_info(PLANILHA)
    ev["planilha"] = {"caminho": PLANILHA, "sha1_1mb": U.sha1(PLANILHA, 1024 * 1024), **st,
                      "confianca_fonte": "CONFIRMADO POR FONTE HUMANA (Emmily) + "
                                         "CONFIRMADO POR EVIDÊNCIA DIRETA (este script)"}
    z = abrir_zip(PLANILHA)
    ev["anatomia"] = anatomia(z)

    local = U.copiar_para_tmp(PLANILHA, "nc", "02_Acompanhamento_de_NC_automatizada.xlsm")
    wv = openpyxl.load_workbook(local, data_only=True)
    wf = openpyxl.load_workbook(local, data_only=False)

    ev["listas_validacao"] = ler_listas_validacao(wv)

    tabelas = {t["nome"]: t for t in ev["anatomia"]["tabelas_estruturadas"]}
    aba_de = {}
    for aba, obj in ev["anatomia"]["objetos_por_aba"].items():
        for p in obj["tabelas"]:
            for t in ev["anatomia"]["tabelas_estruturadas"]:
                if t["parte"] == p:
                    aba_de[t["nome"]] = aba

    ev["tabelas_lidas"] = {}
    for nome in ("Tbl_VISAO", "Tbl_BASE"):
        t = tabelas[nome]
        cols, linhas, forms = ler_tabela(wv, wf, aba_de[nome], t["ref"])
        ev["tabelas_lidas"][nome] = {"aba": aba_de[nome], "ref": t["ref"], "colunas": cols,
                                     "n_linhas": len(linhas), "formulas_de_coluna": forms,
                                     "linhas": linhas}
        print("      %-10s aba=%-12s %3d linhas, %d colunas" % (nome, aba_de[nome], len(linhas), len(cols)))

    visao = ev["tabelas_lidas"]["Tbl_VISAO"]["linhas"]
    base = ev["tabelas_lidas"]["Tbl_BASE"]["linhas"]

    # ------------------------------------------------------------ relatorios
    print("[2/5] lendo a secao 8 dos 8 relatorios da amostra...")
    estr = json.load(io.open(ESTRUTURA, encoding="utf-8"))
    rels = []
    for chave, d in estr.items():
        secs = d.get("texto_secoes", {})
        t83 = t81 = t82 = ""
        for k, v in secs.items():
            kk = re.sub(r"\s+", " ", k).strip()
            if kk.startswith("8.3"):
                t83 = v
            elif kk.startswith("8.1"):
                t81 = v
            elif kk.startswith("8.2"):
                t82 = v
        frase, cont = extrair_83(t83)
        rels.append({
            "chave": chave, "obra": d["obra"], "numero": d["numero"],
            "rodovia": U.RODOVIA_LABEL[d["obra"]],
            "periodo_inicio": d.get("data_inicio"), "periodo_fim": d.get("data_fim"),
            "arquivo_pdf": d.get("arquivo"),
            "secao_8_3": {"frase_grafico2": frase, "contagens": cont,
                          "frase_atraso": extrair_atraso(t83), "texto": re.sub(r"\s+", " ", t83).strip()},
            "secao_8_1": extrair_81(t81),
            "secao_8_2": {"texto": re.sub(r"\s+", " ", t82).strip()},
        })
    ev["relatorios"] = rels

    # ------------------------------------------------------- cruzamento 8.3
    print("[3/5] cruzando os quantitativos da 8.3 com a planilha...")
    cruz = []
    for r in rels:
        rod = r["rodovia"]
        corte = dt.date.fromisoformat(r["periodo_fim"]) if r["periodo_fim"] else None
        atual_v = contar_status(visao, rod, "Rodovia", "Status ", "Data de Abertura", "Data de Encerramento")
        recon_v = contar_status(visao, rod, "Rodovia", "Status ", "Data de Abertura", "Data de Encerramento", corte)
        atual_b = contar_status(base, rod, "Rodovia", "Status", "Data_Notificacao", "Data de Encerramento")
        esperado = r["secao_8_3"]["contagens"]

        # "em aberto" de RD-101/RD-104 agrega Aberta + Em Analise + reabertas
        def conf(cont, agrega):
            c = dict(cont)
            if agrega and rod in AGREGA_EM_ANALISE:
                c["Aberta"] = c.get("Aberta", 0) + c.pop("Em Análise", 0) + c.pop("(aberta na data de corte)", 0)
            return c
        recon_aj = conf(recon_v, True)
        atual_aj = conf(atual_v, True)

        def bate(cont):
            return {k: (cont.get(k, 0), v, cont.get(k, 0) == v) for k, v in esperado.items()}
        b_atual, b_recon = bate(atual_aj), bate(recon_aj)
        cruz.append({
            "relatorio": r["chave"], "rodovia": rod, "periodo_fim": r["periodo_fim"],
            "relatorio_8_3": esperado,
            "Tbl_VISAO_hoje": atual_v, "Tbl_VISAO_na_data_de_corte": recon_v,
            "Tbl_BASE_hoje": atual_b,
            "confere_hoje": all(v[2] for v in b_atual.values()),
            "confere_na_data_de_corte": all(v[2] for v in b_recon.values()),
            "categorias_ok_hoje": sum(1 for v in b_atual.values() if v[2]),
            "categorias_ok_corte": sum(1 for v in b_recon.values() if v[2]),
            "categorias": len(esperado),
            "detalhe_corte": {k: {"planilha": v[0], "relatorio": v[1], "igual": v[2]}
                              for k, v in b_recon.items()},
            "total_relatorio": sum(esperado.values()),
            "total_planilha_corte": sum(recon_v.values()),
        })
        print("      %-12s %-8s 8.3=%s | hoje %d/%d | corte %d/%d | total %d vs %d" % (
            r["chave"], rod, esperado, cruz[-1]["categorias_ok_hoje"], len(esperado),
            cruz[-1]["categorias_ok_corte"], len(esperado),
            cruz[-1]["total_relatorio"], cruz[-1]["total_planilha_corte"]))
    ev["cruzamento_8_3"] = cruz

    # ------------------------------------------------- regra: NC do periodo na 8.1
    print("[4/5] verificando a regra 'NC aberta no periodo aparece na 8.1'...")
    regra = []
    for r in rels:
        rod, ini, fim = r["rodovia"], r["periodo_inicio"], r["periodo_fim"]
        no_periodo = [x for x in visao
                      if str(x.get("Rodovia") or "").strip() == rod
                      and isinstance(x.get("Data de Abertura"), str)
                      and ini <= x["Data de Abertura"] <= fim]
        citadas = {i["nc"] for i in r["secao_8_1"]["itens"]}
        esperadas = {re.match(r"(NC-\d+)", str(x.get("Documento") or "")).group(1)
                     for x in no_periodo if re.match(r"NC-\d+", str(x.get("Documento") or ""))}
        regra.append({"relatorio": r["chave"], "periodo": [ini, fim],
                      "nc_abertas_no_periodo": sorted(esperadas),
                      "nc_citadas_na_8_1": sorted(citadas),
                      "faltando_na_8_1": sorted(esperadas - citadas),
                      "citadas_sem_registro": sorted(citadas - esperadas),
                      "conforme": esperadas == citadas})
        if not regra[-1]["conforme"]:
            print("      !! %s faltando=%s extras=%s" % (r["chave"], regra[-1]["faltando_na_8_1"],
                                                         regra[-1]["citadas_sem_registro"]))
    ev["regra_8_1_vs_planilha"] = regra

    # ---------------------------------------------- NC individual x planilha x texto
    print("[5/5] comparando NC original x planilha x texto da 8.1...")
    lin = []
    if not args.sem_pdf:
        idx = {o: indexar_pdfs_nc(o) for o in U.OBRAS}
        alvos = []
        for r in rels:
            for it in r["secao_8_1"]["itens"]:
                alvos.append((r, it["nc"], it))
        # completa com NCs de obras sem citacao na amostra (RD-103 / RD-104)
        extras = [("RD103", "NC-026"), ("RD104", "NC-027"), ("RD104", "NC-028")]
        for obra, nc in extras:
            alvos.append((None, nc, {"nc": nc, "obra_forcada": obra}))

        for r, nc, it in alvos:
            obra = it.get("obra_forcada") or (r["obra"] if r else None)
            rod = U.RODOVIA_LABEL[obra]
            cands = [c for c in idx[obra].get(nc, []) if c["eh_notificacao"]]
            linha_v = next((x for x in visao
                            if str(x.get("Documento") or "").startswith(nc)
                            and str(x.get("Rodovia") or "").strip() == rod), None)
            reg = {"nc": nc, "obra": obra, "rodovia": rod,
                   "relatorio": r["chave"] if r else None,
                   "pdf_original": cands[0]["arquivo"] if cands else None,
                   "todos_arquivos_da_nc": [c["arquivo"] for c in idx[obra].get(nc, [])],
                   "linha_planilha": linha_v,
                   "texto_relatorio_8_1": it.get("descricao")}
            if cands:
                meta, pags = texto_pdf(cands[0]["arquivo"])
                reg["pdf_metadata"] = meta
                reg["pdf_n_paginas"] = len(pags)
                reg["pdf_gerado_por_excel"] = bool(re.search(r"(?i)excel", meta.get("creator", "") + meta.get("producer", "")))
                reg["pdf_texto_paginas"] = [re.sub(r"[ \t]+", " ", p).strip() for p in pags]
                reg["pdf_campos_formulario"] = campos_do_pdf(pags)
            irr = (linha_v or {}).get("Irregularidade ") or (linha_v or {}).get("Irregularidade")
            reg["irregularidade_planilha"] = irr
            reg["resumo_planilha"] = (linha_v or {}).get("Resumo")
            if irr:
                reg["primeira_frase_irregularidade"] = primeira_frase(irr)
                if reg.get("pdf_texto_paginas"):
                    corpo = " ".join(reg["pdf_texto_paginas"])
                    reg["irregularidade_planilha_no_pdf"] = contido(irr, corpo)
                    reg["campos_planilha_no_pdf"] = conferir_campos(linha_v, corpo)
                if it.get("descricao"):
                    data_br = None
                    m = re.search(r"Emitida em (\d{2}/\d{2}/\d{4})", "NC: Emitida em " + (it.get("data_citada") or ""))
                    dab = (linha_v or {}).get("Data de Abertura")
                    if isinstance(dab, str) and re.match(r"\d{4}-\d\d-\d\d$", dab):
                        data_br = "%s/%s/%s" % (dab[8:10], dab[5:7], dab[0:4])
                    prev = frase_82_esperada(irr, nc, data_br or "??/??/????")
                    real = "%s: Emitida em %s, %s" % (nc, it.get("data_citada"), it["descricao"])
                    reg["frase_prevista_pelo_template"] = prev
                    reg["frase_real_no_relatorio"] = real
                    reg["similaridade_template_x_real"] = sim(prev, real)
                    reg["similaridade_1a_frase_x_relatorio"] = sim(reg["primeira_frase_irregularidade"],
                                                                   it["descricao"])
            lin.append(reg)
            cc = reg.get("irregularidade_planilha_no_pdf") or {}
            print("      %-8s %-8s pdf=%s planilha=%s  irregularidade literal no PDF=%s/%s  template=%s" % (
                nc, rod, "sim" if cands else "NAO", "sim" if linha_v else "NAO",
                cc.get("encontradas_literalmente"), cc.get("frases"),
                reg.get("similaridade_template_x_real")))
    ev["linhagem_nc"] = lin

    with io.open(SAIDA, "w", encoding="utf-8") as fh:
        json.dump(ev, fh, ensure_ascii=False, indent=1, default=jsonable)
    print("\nOK -> %s (%.1f KB)" % (SAIDA, os.path.getsize(SAIDA) / 1024.0))


if __name__ == "__main__":
    main()
