# -*- coding: utf-8 -*-
r"""
09_cadeia_avanco_fisico.py — reconstroi e VERIFICA NUMERICAMENTE, elo a elo, a
cadeia do avanco fisico do Relatorio Semanal do Programa, usando a fonte real:

    C:\caminho\exemplo\PastaSincronizada\Programa - 2. Compartilhado interno\
        03_Gestao dos contratos\
            UNIFILAR (com nao conformidades)\UNIFILAR {147,178A,180,461}.xlsx   <- PRIMARY
            [Programa] Avanco Global de Obra.xlsx                                   <- INTERMEDIATE (mestre)

Elos verificados:
  E0  UNIFILAR!'UNIFILAR GO <n>'!G/H  ..... digitacao manual? (contagem de formulas)
  E1  UNIFILAR!'Base para BI'!E/F  =  'UNIFILAR GO <n>'!G/H          (formula interna)
  E2  [Programa]!'Avanco Global - <OBRA>'!H/I  =  '[k]Base para BI'!E/F  (link externo)
  E3  [Programa]!C3 = SUM(L12:L..) e C7 = SUM(N12:N..)                   (recalculo)
  E4  [Programa]!Acumulado (obra, data) = C3*100 / C7*100                (append semanal)
  E5  Grafico - Avanco Global.xlsx 'Relatorio Semanal' ultima linha  (copia manual)
  E6  Grafico 1 do PDF (EMF) — renderizado para conferencia visual

Tambem responde:
  T-11  a aba oculta AVANCO_GERAL das planilhas fotograficas contem o grafico?
  UNC   o caminho \\SERVIDOR\ e fonte alternativa ou reescrita de vinculo?

REGRA: o SharePoint e READ-ONLY. Este script so le de la. Escreve apenas em
output/evidencias/ e tmp/ dentro do repositorio.

Uso:
    python scripts/09_cadeia_avanco_fisico.py
    python scripts/09_cadeia_avanco_fisico.py --sem-render   # pula o PyMuPDF
"""
import os
import re
import io
import sys
import json
import zipfile
import argparse
import datetime as dt
import importlib.util
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

# ---- reaproveita o dissecador OOXML do script 05 (nome comeca com digito) ----
_spec = importlib.util.spec_from_file_location(
    "inv05", os.path.join(U.PROJ, "scripts", "05_inventariar_excel.py"))
INV = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(INV)

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKGREL = "http://schemas.openxmlformats.org/package/2006/relationships"

GESTAO = os.path.join(U.SHAREPOINT, "Programa - 2. Compartilhado interno",
                      "03_Gestao dos contratos")
PASTA_UNIFILAR = os.path.join(GESTAO, "UNIFILAR (com n\u00e3o conformidades)")
PROG_MESTRE = os.path.join(GESTAO, "[Programa] Avan\u00e7o Global de Obra.xlsx")

EVID = os.path.join(U.OUT, "evidencias")
TMP_G = os.path.join(U.TMP, "cadeia_avanco")
EXCEL_JSON = os.path.join(U.OUT, "extracao", "excel")
AMOSTRA = os.path.join(U.PROJ, "state", "amostra.json")

# obra -> sufixo do nome do arquivo UNIFILAR e nome da aba de grade
OBRA_MARCA = {"RD101": "147", "RD102A": "178A", "RD103": "180", "RD104": "461"}

TOL = 1e-9


# ------------------------------------------------------------------ utilidades
def agora():
    return dt.datetime.now().astimezone().isoformat(timespec="seconds")


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def igual(a, b, tol=TOL):
    fa, fb = num(a), num(b)
    if fa is None or fb is None:
        return str(a) == str(b)
    return abs(fa - fb) <= tol


def serial_iso(v):
    f = num(v)
    if f is None:
        return None
    return (dt.datetime(1899, 12, 30) + dt.timedelta(days=f)).date().isoformat()


def col_letras(ref):
    m = re.match(r"([A-Z]+)", ref or "")
    return m.group(1) if m else ""


def lin_num(ref):
    m = re.search(r"(\d+)$", ref or "")
    return int(m.group(1)) if m else 0


def ler_para_memoria(caminho):
    r"""OneDrive Files On-Demand quebra o zipfile com OSError [Errno 22].
    Ler sequencialmente para BytesIO antes de abrir o ZIP resolve."""
    buf = io.BytesIO()
    with open(U.lp(caminho), "rb") as fh:
        while True:
            b = fh.read(4 * 1024 * 1024)
            if not b:
                break
            buf.write(b)
    buf.seek(0)
    return buf


def dissecar_sharepoint(caminho):
    """Copia para tmp/ (hidrata o OneDrive) e disseca com o motor do script 05."""
    tmp = U.copiar_para_tmp(caminho, subdir="cadeia_avanco")
    d = INV.dissecar(tmp, com_celulas=True)
    st = U.stat_info(caminho)
    d["caminho_origem"] = caminho
    d["origem_tamanho_bytes"] = st["tamanho_bytes"]
    d["origem_modificado"] = st["modificado"]
    return d


def aba(doc, nome):
    for a in doc["abas"]:
        if a["nome"] == nome:
            return {c["ref"]: c for c in a.get("celulas", [])}, a
    return {}, None


def aba_contem(doc, fragmento):
    for a in doc["abas"]:
        if fragmento in (a["nome"] or ""):
            return {c["ref"]: c for c in a.get("celulas", [])}, a
    return {}, None


def v(mapa, ref):
    return (mapa.get(ref) or {}).get("v")


def f(mapa, ref):
    return (mapa.get(ref) or {}).get("f")


# ------------------------------------------------------- E0/E1 — dentro do UNIFILAR
RE_REF_INTERNA = re.compile(r"^'(?P<aba>[^']+)'!(?P<ref>\$?[A-Z]+\$?\d+)$")


def analisar_unifilar(doc, obra):
    """E0: G/H da aba de grade sao digitados? E1: Base para BI!E/F = grade!G/H?"""
    marca = OBRA_MARCA[obra]
    grade, ag = aba(doc, "UNIFILAR GO " + marca)
    bi, abi = aba(doc, "Base para BI")
    res = {
        "obra": obra,
        "arquivo": doc["arquivo"],
        "caminho": doc["caminho_origem"],
        "modificado": doc["origem_modificado"],
        "tamanho_bytes": doc["origem_tamanho_bytes"],
        "sha1": doc.get("sha1"),
        "autor": doc["propriedades"].get("autor"),
        "ultimo_autor": doc["propriedades"].get("ultimo_autor"),
        "criado": doc["propriedades"].get("criado"),
        "abas": [{"nome": a["nome"], "estado": a["estado"],
                  "dimensao": a.get("dimensao"),
                  "celulas": a.get("celulas_nao_vazias"),
                  "formulas": sum(1 for c in a.get("celulas", []) if "f" in c)}
                 for a in doc["abas"]],
        "nomes_definidos": doc["workbook"].get("nomes_definidos", []),
        "links_externos": [l["arquivo_alvo"] for l in doc["links_externos"]],
        "tem_conexao_ou_query": [p for p in doc["partes"]
                                 if re.search(r"connections|queryTable|pivotCache",
                                              p, re.I)],
        "data_corte_B8": {
            "aba": "Base para BI", "ref": "B8",
            "formula": f(bi, "B8"), "valor": v(bi, "B8"),
            "grade_B8": v(grade, "B8"),
        },
    }
    # E0 — colunas G (Atacado) e H (Concluido) da grade: formula ou digitado?
    gh = [(r, c) for r, c in grade.items()
          if re.match(r"^[GH]\d+$", r) and lin_num(r) >= 13]
    com_formula = sorted(r for r, c in gh if c.get("f"))
    res["E0_montante"] = {
        "aba": "UNIFILAR GO " + marca,
        "colunas": "G=Atacado, H=Conclu\u00eddo (linhas >= 13)",
        "celulas": len(gh),
        "com_formula": len(com_formula),
        "exemplos_com_formula": com_formula[:10],
        "veredito": ("DIGITACAO MANUAL — nenhuma formula, nenhum link externo, "
                     "nenhuma conexao de dados"
                     if not com_formula and not doc["links_externos"] else
                     "HA FORMULA — investigar"),
    }
    # E1 — Base para BI!E/F referenciam a grade?
    # so linhas da TABELA DE SERVICOS: Grupo (B) e Servico (C) preenchidos.
    linhas_servico = sorted(
        lin for lin in {lin_num(r) for r in bi if lin_num(r) >= 13}
        if v(bi, "B%d" % lin) and v(bi, "C%d" % lin))
    checks = []
    for lin in linhas_servico:
        for col in ("E", "F"):
            ref = "%s%d" % (col, lin)
            cel = bi.get(ref)
            if not cel:
                continue
            frm = cel.get("f")
            item = {"ref_base_bi": ref, "grupo": v(bi, "B%d" % lin),
                    "servico": v(bi, "C%d" % lin),
                    "formula": frm, "valor_base_bi": v(bi, ref),
                    "alvo": None, "valor_alvo": None, "confere": None}
            if frm:
                m = RE_REF_INTERNA.match(frm.strip())
                if m and m.group("aba") == "UNIFILAR GO " + marca:
                    alvo = m.group("ref").replace("$", "")
                    item["alvo"] = "'%s'!%s" % (m.group("aba"), alvo)
                    # celula vazia na grade e lida como 0 pelo Excel
                    bruto = v(grade, alvo)
                    item["valor_alvo"] = 0 if bruto is None else bruto
                    item["alvo_vazio"] = bruto is None
                    item["confere"] = igual(item["valor_base_bi"], item["valor_alvo"])
            else:
                item["alvo"] = "CONSTANTE (formula sobrescrita por valor)"
            checks.append(item)
    ok = sum(1 for c in checks if c["confere"] is True)
    div = sum(1 for c in checks if c["confere"] is False)
    const = sum(1 for c in checks if not c["formula"])
    res["E1_base_para_bi"] = {
        "aba": "Base para BI", "colunas": "E=Atacado, F=Conclu\u00eddo",
        "linhas_de_servico": linhas_servico,
        "celulas_avaliadas": len(checks), "com_formula_conferida": ok,
        "divergentes": div, "constantes_sem_formula": const,
        "constantes": [c["ref_base_bi"] for c in checks if not c["formula"]],
        "detalhe": checks,
    }
    # celulas-chave citadas no lineage anterior
    res["celulas_chave"] = {
        "E28": {"formula": f(bi, "E28"), "valor": v(bi, "E28"),
                "servico": v(bi, "C28"), "grupo": v(bi, "B28")},
        "F28": {"formula": f(bi, "F28"), "valor": v(bi, "F28"),
                "servico": v(bi, "C28"), "grupo": v(bi, "B28")},
    }
    return res


# ------------------------------------------------------- E2 — [Programa] x UNIFILAR
RE_REF_EXTERNA = re.compile(r"\[(?P<idx>\d+)\](?P<aba>[^!']+)'?!\$?(?P<col>[A-Z]+)\$?(?P<lin>\d+)")


def indice_links(doc):
    """externalLink<N>.xml -> (indice usado nas formulas, arquivo alvo, marca da obra)."""
    mapa = {}
    for l in doc["links_externos"]:
        m = re.search(r"externalLink(\d+)\.xml", l["parte"])
        if not m:
            continue
        idx = int(m.group(1))
        alvo = l["arquivo_alvo"] or ""
        mm = re.search(r"UNIFILAR%20(\w+)\.xlsx", alvo) or \
            re.search(r"UNIFILAR (\w+)\.xlsx", alvo)
        mapa[idx] = {"arquivo_alvo": alvo,
                     "marca": mm.group(1) if mm else None,
                     "abas": l["abas_referenciadas"],
                     "relativo": not alvo.lower().startswith("file:"),
                     "n_cache": len(l["celulas_em_cache"])}
    return mapa


def analisar_prog(doc, unifilares):
    """E2, E3 e E4 sobre um arquivo [Programa] (mestre ou copia semanal)."""
    links = indice_links(doc)
    saida = {
        "arquivo": doc["arquivo"],
        "caminho": doc.get("caminho_origem", doc.get("caminho")),
        "modificado": doc.get("origem_modificado", doc.get("modificado")),
        "sha1": doc.get("sha1"),
        "autor": doc["propriedades"].get("autor"),
        "ultimo_autor": doc["propriedades"].get("ultimo_autor"),
        "links_externos": links,
        "todos_relativos": all(l["relativo"] for l in links.values()) if links else None,
        "por_obra": {},
    }
    for obra in U.OBRAS:
        m, a = aba(doc, "Avan\u00e7o Global - " + obra)
        if not m:
            continue
        ult = max((lin_num(r) for r in m if col_letras(r) == "B" and lin_num(r) >= 12),
                  default=11)
        # E2 — H/I: formula externa, formula interna ou constante?
        e2 = []
        for lin in range(12, ult + 1):
            for col, rot in (("H", "Atacado"), ("I", "Conclu\u00eddo")):
                ref = "%s%d" % (col, lin)
                cel = m.get(ref)
                if not cel:
                    continue
                frm = cel.get("f")
                item = {"ref": ref, "tipo": rot,
                        "subevento": v(m, "D%d" % lin),
                        "presente_no_unifilar": v(m, "G%d" % lin),
                        "formula": frm, "valor": v(m, ref),
                        "classe": None, "alvo_unifilar": None,
                        "valor_no_unifilar": None, "confere": None}
                if not frm:
                    item["classe"] = "CONSTANTE"
                elif "[" in frm:
                    item["classe"] = "LINK_EXTERNO"
                    refs = RE_REF_EXTERNA.findall(frm)
                    total = 0.0
                    alvos = []
                    ok = True
                    for idx, nome_aba, colr, linr in refs:
                        info = links.get(int(idx))
                        if not info or not info["marca"]:
                            ok = False
                            break
                        uni = unifilares.get(info["marca"])
                        if not uni:
                            ok = False
                            break
                        bi, _ = aba(uni, nome_aba.strip().strip("'"))
                        val = num(v(bi, "%s%s" % (colr, linr)))
                        alvos.append("%s!'%s'!%s%s" % (
                            os.path.basename(info["arquivo_alvo"]),
                            nome_aba.strip().strip("'"), colr, linr))
                        if val is None:
                            ok = False
                            break
                        total += val
                    if ok and refs:
                        # SUM(E17:E20) tem apenas 1 par capturado -> nao somavel aqui
                        multi = "SUM(" in (frm or "").upper() and ":" in frm
                        item["alvo_unifilar"] = alvos
                        if not multi:
                            item["valor_no_unifilar"] = total
                            item["confere"] = igual(item["valor"], total, 1e-6)
                        else:
                            item["classe"] = "LINK_EXTERNO_FAIXA"
                else:
                    item["classe"] = "FORMULA_INTERNA"
                e2.append(item)
        diretos = [x for x in e2 if x["classe"] == "LINK_EXTERNO"]
        # E3 — recalculo de C3 e C7
        soma_L = sum(num(v(m, "L%d" % n)) or 0.0 for n in range(12, ult + 1))
        soma_N = sum(num(v(m, "N%d" % n)) or 0.0 for n in range(12, ult + 1))
        c3, c7 = num(v(m, "C3")), num(v(m, "C7"))
        saida["por_obra"][obra] = {
            "aba": a["nome"], "dimensao": a.get("dimensao"),
            "ultima_linha_eventograma": ult,
            "C2_valor_contrato": num(v(m, "C2")),
            "C3": {"formula": f(m, "C3"), "valor": c3, "x100": round(c3 * 100, 4) if c3 else None},
            "C7": {"formula": f(m, "C7"), "valor": c7, "x100": round(c7 * 100, 4) if c7 else None},
            "E2_quantidades": {
                "celulas_H_I": len(e2),
                "link_externo": sum(1 for x in e2 if x["classe"] == "LINK_EXTERNO"),
                "link_externo_faixa": sum(1 for x in e2 if x["classe"] == "LINK_EXTERNO_FAIXA"),
                "formula_interna": sum(1 for x in e2 if x["classe"] == "FORMULA_INTERNA"),
                "constante": sum(1 for x in e2 if x["classe"] == "CONSTANTE"),
                "conferidos_contra_unifilar": sum(1 for x in diretos if x["confere"] is True),
                "divergentes": sum(1 for x in diretos if x["confere"] is False),
                "detalhe": e2,
            },
            "E3_recalculo": {
                "SUM_L12_L%d" % ult: soma_L, "C3": c3,
                "confere_C3": igual(soma_L, c3, 1e-9),
                "SUM_N12_N%d" % ult: soma_N, "C7": c7,
                "confere_C7": igual(soma_N, c7, 1e-9),
            },
            "H18_I18": {
                "H18": {"formula": f(m, "H18"), "valor": v(m, "H18")},
                "I18": {"formula": f(m, "I18"), "valor": v(m, "I18")},
                "subevento_linha_18": v(m, "D18"),
            },
        }
    # E4 — aba Acumulado
    ac, aac = aba(doc, "Acumulado")
    linhas = []
    for n in sorted({lin_num(r) for r in ac if lin_num(r) >= 2}):
        linhas.append({"linha": n, "obra": v(ac, "A%d" % n),
                       "data_serial": v(ac, "B%d" % n),
                       "data": serial_iso(v(ac, "B%d" % n)),
                       "concluido_pct": num(v(ac, "C%d" % n)),
                       "concl_atacado_pct": num(v(ac, "D%d" % n))})
    saida["acumulado"] = {"aba": "Acumulado", "dimensao": aac.get("dimensao") if aac else None,
                          "n_linhas": len(linhas), "linhas": linhas}
    # a ultima linha de cada obra bate com C3/C7 daquela aba?
    conf = []
    for obra in U.OBRAS:
        rot = U.RODOVIA_LABEL[obra]
        da_obra = [l for l in linhas if l["obra"] == rot]
        if not da_obra or obra not in saida["por_obra"]:
            continue
        ult = da_obra[-1]
        po = saida["por_obra"][obra]
        conf.append({
            "obra": obra, "linha_acumulado": ult["linha"], "data": ult["data"],
            "acumulado_C": ult["concluido_pct"], "C3_x100": po["C3"]["x100"],
            "confere_concluido": igual(round(ult["concluido_pct"], 2),
                                       round(po["C3"]["x100"], 2), 5e-3),
            "acumulado_D": ult["concl_atacado_pct"], "C7_x100": po["C7"]["x100"],
            "confere_concl_atacado": igual(round(ult["concl_atacado_pct"], 2),
                                           round(po["C7"]["x100"], 2), 5e-3),
        })
    saida["E4_acumulado_x_C3_C7"] = conf
    return saida


# --------------------------------------- E5/E6 — semanas da amostra e o PDF
def carregar_json_semana(obra, semana, tipo):
    p = os.path.join(EXCEL_JSON, "%s_S%02d_%s.json" % (obra, semana, tipo))
    if not os.path.exists(p):
        return None
    return json.load(open(p, encoding="utf-8"))


def verificar_semanas(mestre):
    """E5: Grafico!'Relatorio Semanal' ultima linha = [Programa] semanal C3/C7 x100,
    e cada barra do grafico = uma linha do Acumulado do MESTRE."""
    if not os.path.exists(AMOSTRA):
        return {"erro": "state/amostra.json ausente"}
    am = json.load(open(AMOSTRA, encoding="utf-8"))
    acum = mestre["acumulado"]["linhas"]
    saida = []
    for obra, info in am["obras"].items():
        for s in info["selecionados"]:
            n = s["numero"]
            item = {"obra": obra, "semana": n,
                    "periodo": [s.get("data_inicio"), s.get("data_fim")]}
            cp = carregar_json_semana(obra, n, "prog_avanco_global")
            gr = carregar_json_semana(obra, n, "grafico_avanco_global")
            un = carregar_json_semana(obra, n, "unifilar")
            if cp:
                m, _ = aba_contem(cp, "Avan\u00e7o Global - " + obra)
                item["prog_semanal"] = {
                    "modificado": cp["modificado"],
                    "C3": num(v(m, "C3")), "C7": num(v(m, "C7")),
                    "links_externos": [l["arquivo_alvo"] for l in cp["links_externos"]],
                    "alvo_relativo": all(not (l["arquivo_alvo"] or "").lower()
                                         .startswith("file:")
                                         for l in cp["links_externos"]),
                }
            if un:
                bi, _ = aba_contem(un, "Base para BI")
                item["unifilar_semanal"] = {
                    "modificado": un["modificado"],
                    "B8_data_corte": v(bi, "B8"),
                    "E28": num(v(bi, "E28")), "F28": num(v(bi, "F28")),
                    "servico_linha_28": v(bi, "C28"),
                }
            if gr:
                g, _ = aba_contem(gr, "Relat")
                lins = sorted({lin_num(r) for r in g if lin_num(r) >= 3})
                item["grafico_janela"] = [
                    {"linha": L, "data": (g.get("A%d" % L) or {}).get("data_iso", "")[:10]
                     or serial_iso(v(g, "A%d" % L)),
                     "concluido_pct": num(v(g, "B%d" % L)),
                     "concl_atacado_pct": num(v(g, "C%d" % L))} for L in lins]
                item["grafico_tem_formula"] = any("f" in c for c in g.values())
            # E5 — ultima barra x C3/C7 da copia semanal
            if item.get("grafico_janela") and item.get("prog_semanal"):
                ult = item["grafico_janela"][-1]
                c3 = item["prog_semanal"]["C3"]
                c7 = item["prog_semanal"]["C7"]
                item["E5_ultima_barra_x_prog"] = {
                    "data": ult["data"],
                    "grafico_B": ult["concluido_pct"], "C3_x100": round(c3 * 100, 4),
                    "confere": igual(ult["concluido_pct"], round(c3 * 100, 2), 5e-3),
                    "grafico_C": ult["concl_atacado_pct"], "C7_x100": round(c7 * 100, 4),
                    "confere_ca": igual(ult["concl_atacado_pct"], round(c7 * 100, 2), 5e-3),
                }
            # E4b — TODAS as barras da janela existem no Acumulado do MESTRE?
            if item.get("grafico_janela"):
                rot = U.RODOVIA_LABEL[obra]
                casados = []
                for b in item["grafico_janela"]:
                    alvo = [l for l in acum if l["obra"] == rot and l["data"] == b["data"]]
                    if alvo:
                        a0 = alvo[-1]
                        casados.append({
                            "data": b["data"],
                            "grafico": [b["concluido_pct"], b["concl_atacado_pct"]],
                            "mestre_acumulado": [a0["concluido_pct"], a0["concl_atacado_pct"]],
                            "linha_mestre": a0["linha"],
                            "confere": (igual(b["concluido_pct"], a0["concluido_pct"], 5e-3)
                                        and igual(b["concl_atacado_pct"],
                                                  a0["concl_atacado_pct"], 5e-3)),
                        })
                    else:
                        casados.append({"data": b["data"], "grafico":
                                        [b["concluido_pct"], b["concl_atacado_pct"]],
                                        "mestre_acumulado": None, "confere": None})
                item["E4b_barras_x_acumulado_mestre"] = casados
            item["pdf"] = s.get("pdf", {}).get("caminho")
            saida.append(item)
    return saida


def renderizar_grafico1(saida_semanas):
    """E6: renderiza a pagina do 'Grafico 1' de cada PDF para conferencia visual."""
    try:
        import fitz
    except ImportError:
        return {"erro": "PyMuPDF ausente"}
    U.assert_read_only(TMP_G)
    os.makedirs(TMP_G, exist_ok=True)
    out = []
    for it in saida_semanas:
        pdf = it.get("pdf")
        if not pdf or not U.exists(pdf):
            continue
        doc = fitz.open(U.lp(pdf))
        pag = None
        for i in range(min(12, doc.page_count)):
            if "Gr\u00e1fico 1" in doc[i].get_text():
                pag = i
                break
        if pag is not None:
            fn = os.path.join(TMP_G, "%s_RO%03d_p%d_grafico1.png" %
                              (it["obra"], it["semana"], pag + 1))
            doc[pag].get_pixmap(dpi=140).save(fn)
            out.append({"obra": it["obra"], "semana": it["semana"],
                        "pdf": pdf, "pagina": pag + 1, "png": fn,
                        "observacao": "o Grafico 1 e imagem (EMF) — sem camada de "
                                      "texto; conferencia e visual"})
        doc.close()
    return out


# ---------------------------------------------------- T-11 — aba AVANCO_GERAL
def extrair_imagens_na_celula(caminho, nome_aba_frag, destino, rotulo):
    """Resolve vm -> metadata -> rdrichvalue -> richValueRel -> xl/media."""
    z = zipfile.ZipFile(ler_para_memoria(caminho))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    relwb = {r.get("Id"): r.get("Target")
             for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    parte = nome = estado = None
    for s in wb.find("{%s}sheets" % NS_MAIN):
        if nome_aba_frag in (s.get("name") or "").upper():
            alvo = relwb[s.get("{%s}id" % NS_REL)]
            parte = alvo if alvo.startswith("xl/") else "xl/" + alvo.lstrip("/")
            nome, estado = s.get("name"), s.get("state") or "visible"
    if not parte:
        return {"aba": None}
    celulas = {}
    for ev, el in ET.iterparse(z.open(parte), events=("end",)):
        if el.tag == "{%s}c" % NS_MAIN and el.get("vm"):
            celulas[el.get("r")] = int(el.get("vm"))
        if el.tag == "{%s}row" % NS_MAIN:
            el.clear()
    bks = list(ET.fromstring(z.read("xl/metadata.xml")).find("{%s}valueMetadata" % NS_MAIN))
    rvs = list(ET.fromstring(z.read("xl/richData/rdrichvalue.xml")))
    rels = {r.get("Id"): r.get("Target") for r in
            ET.fromstring(z.read("xl/richData/_rels/richValueRel.xml.rels"))}
    ordem = [e.get("{%s}id" % NS_REL)
             for e in ET.fromstring(z.read("xl/richData/richValueRel.xml"))]
    U.assert_read_only(destino)
    os.makedirs(destino, exist_ok=True)
    itens = []
    for ref, vm in sorted(celulas.items()):
        try:
            idx = int(bks[vm - 1][0].get("v"))
            k = int([e for e in rvs[idx] if e.tag.endswith("}v")][0].text)
            arq = os.path.basename(rels[ordem[k]])
            dados = z.read("xl/media/" + arq)
            fn = os.path.join(destino, "%s_%s_%s" % (rotulo, ref, arq))
            open(fn, "wb").write(dados)
            import hashlib
            itens.append({"celula": ref, "midia": arq, "bytes": len(dados),
                          "sha1": hashlib.sha1(dados).hexdigest(), "png": fn})
        except Exception as e:
            itens.append({"celula": ref, "erro": "%s: %s" % (type(e).__name__, e)})
    return {"aba": nome, "estado": estado, "imagens": itens}


def investigar_t11():
    if not os.path.exists(AMOSTRA):
        return {"erro": "state/amostra.json ausente"}
    am = json.load(open(AMOSTRA, encoding="utf-8"))
    dest = os.path.join(U.TMP, "avanco_geral")
    res = []
    for obra, info in am["obras"].items():
        for s in info["selecionados"]:
            cam = s["caminho"]
            aux = None
            for it in U.listdir(cam):
                if U.norm(it).startswith("arquivos aux"):
                    aux = os.path.join(cam, it)
            if not aux:
                continue
            foto = None
            for it in U.listdir(aux):
                if "fotogr" in U.norm(it) and it.lower().endswith(".xlsx"):
                    foto = os.path.join(aux, it)
            if not foto:
                continue
            rot = "%s_S%02d" % (obra, s["numero"])
            try:
                r = extrair_imagens_na_celula(foto, "AVAN", dest, rot)
            except Exception as e:
                r = {"erro": "%s: %s" % (type(e).__name__, e)}
            r.update({"obra": obra, "semana": s["numero"], "arquivo": foto})
            res.append(r)
    return res


# ------------------------------------------------------------------- principal
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sem-render", action="store_true",
                    help="nao renderiza as paginas do Grafico 1 com PyMuPDF")
    ap.add_argument("--sem-t11", action="store_true",
                    help="nao abre as planilhas fotograficas (70-170 MB cada)")
    args = ap.parse_args()

    U.assert_read_only(EVID)
    os.makedirs(EVID, exist_ok=True)

    ev = {"gerado_em": agora(),
          "pasta_fonte": PASTA_UNIFILAR,
          "arquivo_mestre": PROG_MESTRE,
          "niveis_de_confianca": [
              "CONFIRMADO POR FONTE HUMANA", "CONFIRMADO POR EVIDENCIA DIRETA",
              "CONFIRMADO POR CRUZAMENTO DE FONTES", "FORTE INFERENCIA",
              "HIPOTESE", "DESCONHECIDO"]}

    # ---- inventario da pasta
    ev["pasta_unifilar"] = {
        "existe": U.exists(PASTA_UNIFILAR),
        "arquivos": [dict(nome=n, **U.stat_info(os.path.join(PASTA_UNIFILAR, n)))
                     for n in U.listdir(PASTA_UNIFILAR)] if U.exists(PASTA_UNIFILAR) else [],
    }
    ev["irmaos_em_03_gestao_dos_contratos"] = sorted(U.listdir(GESTAO))

    # ---- UNIFILARes (PRIMARY)
    print("[1/6] dissecando os 4 UNIFILAR da pasta real...")
    unifilares_doc = {}
    ev["unifilar"] = {}
    for obra, marca in OBRA_MARCA.items():
        cam = os.path.join(PASTA_UNIFILAR, "UNIFILAR %s.xlsx" % marca)
        if not U.exists(cam):
            ev["unifilar"][obra] = {"erro": "arquivo ausente", "caminho": cam}
            continue
        d = dissecar_sharepoint(cam)
        unifilares_doc[marca] = d
        ev["unifilar"][obra] = analisar_unifilar(d, obra)
        print("      %-7s %s | E0 %s" % (obra, os.path.basename(cam),
                                         ev["unifilar"][obra]["E0_montante"]["veredito"]))

    # ---- [Programa] mestre (INTERMEDIATE)
    print("[2/6] dissecando o [Programa] mestre...")
    mestre_doc = dissecar_sharepoint(PROG_MESTRE)
    ev["prog_mestre"] = analisar_prog(mestre_doc, unifilares_doc)

    # ---- copias semanais + Grafico
    print("[3/6] conferindo as semanas da amostra...")
    ev["semanas"] = verificar_semanas(ev["prog_mestre"])

    # ---- UNC SERVIDOR
    print("[4/6] classificando o caminho UNC SERVIDOR...")
    alvos = []
    for it in ev["semanas"]:
        cs = it.get("prog_semanal")
        if cs:
            for a in cs["links_externos"]:
                alvos.append({"obra": it["obra"], "semana": it["semana"],
                              "modificado": cs["modificado"], "alvo": a})
    unc = [a for a in alvos if "SERVIDOR" in (a["alvo"] or "")]
    ev["caminho_unc_SERVIDOR"] = {
        "ocorrencias": unc,
        "total_alvos_examinados": len(alvos),
        "alvos_relativos_sharepoint": sum(1 for a in alvos
                                          if not (a["alvo"] or "").lower().startswith("file:")),
        "mestre_usa_caminho_relativo": ev["prog_mestre"]["todos_relativos"],
        "computador_desta_maquina": os.environ.get("COMPUTERNAME"),
        "classificacao": ("VINCULO REESCRITO — caminho legado/cache de outra estacao; "
                          "nao e uma fonte alternativa"),
    }

    # ---- T-11
    if not args.sem_t11:
        print("[5/6] T-11: aba oculta AVANCO_GERAL das planilhas fotograficas...")
        ev["T11_avanco_geral"] = investigar_t11()
    else:
        print("[5/6] T-11 pulado (--sem-t11)")

    # ---- E6 render
    if not args.sem_render:
        print("[6/6] renderizando as paginas do Grafico 1...")
        ev["E6_grafico1_render"] = renderizar_grafico1(ev["semanas"])
    else:
        print("[6/6] render pulado (--sem-render)")

    # ---- classificacao das fontes
    ev["classificacao_das_fontes"] = [
        {"arquivo": "UNIFILAR (com nao conformidades)\\UNIFILAR <n>.xlsx",
         "aba": "UNIFILAR GO <n>", "classe": "PRIMARY",
         "razao": "colunas G/H digitadas a mao; sem formula, link ou conexao"},
        {"arquivo": "UNIFILAR (com nao conformidades)\\UNIFILAR <n>.xlsx",
         "aba": "Base para BI", "classe": "INTERMEDIATE",
         "razao": "reorganiza a grade por servico; so formulas internas"},
        {"arquivo": "03_Gestao dos contratos\\[Programa] Avanco Global de Obra.xlsx",
         "aba": "Avanco Global - <OBRA>", "classe": "INTERMEDIATE",
         "razao": "pondera pelo eventograma; le os 4 UNIFILAR por link externo relativo"},
        {"arquivo": "[Programa] ... .xlsx", "aba": "Acumulado", "classe": "INTERMEDIATE",
         "razao": "serie historica append-only das 4 obras, em pontos percentuais"},
        {"arquivo": "[Programa] ... .xlsx", "aba": "Resumo - Avancos Fisicos",
         "classe": "PRESENTATION", "razao": "tabela f_AvancoFisico para BI"},
        {"arquivo": "<semana>\\Arquivos auxiliares\\[Programa] Avanco Global de Obra.xlsx",
         "aba": "todas", "classe": "INTERMEDIATE (snapshot)",
         "razao": "copia congelada do mestre na pasta da obra"},
        {"arquivo": "<semana>\\Arquivos auxiliares\\UNIFILAR <n>.xlsx",
         "aba": "todas", "classe": "PRIMARY (snapshot)",
         "razao": "copia congelada do UNIFILAR fonte"},
        {"arquivo": "<semana>\\Arquivos auxiliares\\Grafico - Avanco Global.xlsx",
         "aba": "Relatorio Semanal / Acumulado", "classe": "PRESENTATION",
         "razao": "zero formulas; existe para desenhar o barChart"},
        {"arquivo": "RO_<nnn> - <RODOVIA>.docx / .pdf", "aba": "Secao 3, Grafico 1",
         "classe": "OUTPUT", "razao": "EMF colado; sem camada de dados"},
        {"arquivo": "Controle de Cronograma\\...\\UNIFILAR <n> - Controle de cronograma.xlsx",
         "aba": "<MES>-<ANO>", "classe": "INTERMEDIATE (paralelo)",
         "razao": "arquivo mensal irmao do UNIFILAR; nao alimenta o [Programa]"},
        {"arquivo": "One Sheet\\<obra>\\<servico>\\*.jpeg", "aba": "-",
         "classe": "FORA DA CADEIA", "razao": "so fotos; nenhum numero"},
        {"arquivo": "Programa - Resumo BM's.xlsx", "aba": "RESUMO BM'S",
         "classe": "FORA DA CADEIA (candidato)",
         "razao": "registro de Boletins de Medicao; sem vinculo com o [Programa]"},
    ]

    dest = os.path.join(EVID, "cadeia_avanco_fisico.json")
    U.assert_read_only(dest)
    json.dump(ev, open(dest, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("\nOK ->", dest, os.path.getsize(dest), "bytes")

    # ---- resumo no terminal
    print("\n== RESUMO ==")
    for obra in U.OBRAS:
        u = ev["unifilar"].get(obra, {})
        if "E0_montante" not in u:
            continue
        e1 = u["E1_base_para_bi"]
        print("  %-7s E0: %d/%d celulas G/H com formula | E1: %d conferidas, "
              "%d divergentes, %d constantes | B8=%s" % (
                  obra, u["E0_montante"]["com_formula"], u["E0_montante"]["celulas"],
                  e1["com_formula_conferida"], e1["divergentes"],
                  e1["constantes_sem_formula"], u["data_corte_B8"]["valor"]))
    for obra, po in ev["prog_mestre"]["por_obra"].items():
        q = po["E2_quantidades"]
        print("  %-7s E2: link=%d faixa=%d interna=%d const=%d | conferidos %d/%d | "
              "E3 C3=%s C7=%s" % (
                  obra, q["link_externo"], q["link_externo_faixa"],
                  q["formula_interna"], q["constante"],
                  q["conferidos_contra_unifilar"], q["link_externo"],
                  po["E3_recalculo"]["confere_C3"], po["E3_recalculo"]["confere_C7"]))
    for c in ev["prog_mestre"]["E4_acumulado_x_C3_C7"]:
        print("  %-7s E4: %s -> %.2f / %.2f  (C3x100=%.4f C7x100=%.4f) %s" % (
            c["obra"], c["data"], c["acumulado_C"], c["acumulado_D"],
            c["C3_x100"], c["C7_x100"],
            "OK" if c["confere_concluido"] and c["confere_concl_atacado"] else "DIVERGE"))
    for it in ev["semanas"]:
        e5 = it.get("E5_ultima_barra_x_prog")
        b = it.get("E4b_barras_x_acumulado_mestre") or []
        ok = sum(1 for x in b if x["confere"] is True)
        print("  %-7s S%02d E5: %s | E4b: %d/%d barras batem com o Acumulado do mestre"
              % (it["obra"], it["semana"],
                 "OK" if e5 and e5["confere"] and e5["confere_ca"] else "?", ok, len(b)))


if __name__ == "__main__":
    main()
