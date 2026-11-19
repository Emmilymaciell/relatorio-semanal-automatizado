# -*- coding: utf-8 -*-
r"""
12_analisar_ambiental.py - fecha as pendencias P-09 (onde fica o "Dashboard
Ambiental") e T-05 (origem dos 3 paineis da secao 6).

O que faz:
  1. Disseca a planilha corporativa
     `G2 -  Engenharia\04_Acompanhamento - Licenca Ambiental.xlsx`
     (abas, colunas, formulas, KPIs, validacoes, nomes definidos, webextension).
  2. Extrai `Licencas` e `Condicionantes` como registros estruturados e liga
     cada linha a uma obra pela aba `Empresas` (coluna G = rodovia).
  3. Reimplementa em Python a formula de status das condicionantes (que usa
     TODAY()) e recalcula os KPIs do Dashboard para qualquer data-base -
     e assim reproduz o painel do Power BI em datas passadas.
  4. Confronta com:
       - os numeros escritos na secao 6 dos 8 relatorios da amostra;
       - os numeros LIDOS nas imagens dos paineis (registrados em PAINEIS_LIDOS).
  5. Inventaria o acervo ambiental da obra
     (`3_Gestao Tecnica de Obra\6. Acompanhamento ambiental e SST`):
     FA / PA / RA / RI / NT, datas das pastas-visita e ciclo quinzenal.
  6. Le a 1a pagina de cada PDF `PA-NNN` e extrai numero, obra, data da versao,
     data da verificacao em campo e o "Relatorio Quinzenal no NNN" citado.

REGRA: o SharePoint e READ-ONLY. So le de la. Escreve apenas em
output/evidencias/ambiental.json e (opcionalmente) tmp/.

Uso:
    python scripts/12_analisar_ambiental.py
    python scripts/12_analisar_ambiental.py --sem-pdf   # pula os PDFs PA-NNN
"""
import os
import re
import sys
import json
import argparse
import importlib
import datetime as dt
from collections import defaultdict, Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

M05 = importlib.import_module("05_inventariar_excel")

PLANILHA = os.path.join(U.BASE_OBRAS, "04_Acompanhamento - Licenca Ambiental.xlsx")
PASTA_AMB = os.path.join("3_Gestao Tecnica de Obra",
                         "6. Acompanhamento ambiental e SST")
ESTRUTURA = os.path.join(U.OUT, "comparacoes", "estrutura_relatorios.json")
SAIDA = os.path.join(U.OUT, "evidencias", "ambiental.json")

EPOCH = dt.date(1899, 12, 30)

# aba `Empresas`: coluna B (razao social) -> coluna G (rodovia). Lido do arquivo,
# mas fixado aqui tambem como chave de conferencia.
EMPRESA_OBRA_ESPERADO = {
    "Executora C": "RD103",
    "EXECUTORA B": "RD102A",
    "EXECUTORA A": "RD101",
    "EXECUTORA-D": "RD104",
}

# Valores LIDOS nas imagens dos paineis embutidos nos PDFs da amostra.
# Cada imagem e um recorte (JPEG 96 dpi) do Power BI publicado; nao ha camada
# de texto, entao os numeros abaixo foram transcritos por leitura da imagem.
# Referencia da imagem: <relatorio>, pagina, xref do objeto no PDF.
PAINEIS_LIDOS = {
    "RD101_047": {
        "indices": {"pagina": 13, "xref": 1527, "atualizado_em": "2026-08-28",
                    "total_licencas": 8, "licencas_emitidas": 8,
                    "licencas_pendentes": None,
                    "fases": {"Outorga de uso de recurso hídrico": 4,
                              "Licença Ambiental - Canteiro de obra": 1,
                              "Licença Ambiental - Oficinas Mecânicas": 1,
                              "Licença Ambiental - Tanque Aéreo": 1,
                              "Licença Ambiental Única": 1}},
        "condicionantes": {"pagina": 14, "xref": 1539, "atualizado_em": "2026-08-28",
                           "finalizadas": 11, "no_prazo": 1, "pendentes": None,
                           "proximo_a_vencer": 34, "total_rodovia": 46,
                           "pct": [23.91, 2.17, 73.91]},
    },
    "RD101_048": {
        "indices": {"pagina": 13, "xref": 1553, "atualizado_em": "2026-09-07",
                    "total_licencas": 8, "licencas_emitidas": 8,
                    "licencas_pendentes": None,
                    "fases": {"Outorga de uso de recurso hídrico": 4,
                              "Licença Ambiental - Canteiro de obra": 1,
                              "Licença Ambiental - Oficinas Mecânicas": 1,
                              "Licença Ambiental - Tanque Aéreo": 1,
                              "Licença Ambiental Única": 1},
                    "tabela_n_licenca": ["202500001", "202500002", "202500003",
                                         "202500004", "202500005", "202500006",
                                         "202500007", "202500008"]},
        "condicionantes": {"pagina": 14, "xref": 1566, "atualizado_em": "2026-09-07",
                           "finalizadas": 11, "no_prazo": 1, "pendentes": None,
                           "proximo_a_vencer": 34, "total_rodovia": 46,
                           "pct": [23.91, 2.17, 73.91]},
        "qr": {"pagina": 14, "xref": 1567, "px": [591, 591]},
    },
    "RD102A_048": {
        "indices": {"pagina": 11, "xref": 835, "atualizado_em": "2026-08-28",
                    "total_licencas": 8, "licencas_emitidas": 7,
                    "licencas_pendentes": 1,
                    "fases": {"Outorga de uso de recurso hídrico": 6,
                              "Licença Ambiental - Canteiro de obra": 1,
                              "Licença Ambiental Única": 1}},
        "condicionantes": {"pagina": 12, "xref": 855, "atualizado_em": "2026-08-28",
                           "finalizadas": 12, "no_prazo": 1, "pendentes": None,
                           "proximo_a_vencer": 31, "total_rodovia": 44,
                           "pct": [27.27, 2.27, 70.45]},
    },
    "RD102A_049": {
        "indices": {"pagina": 11, "xref": 830, "atualizado_em": "2026-09-07",
                    "total_licencas": 8, "licencas_emitidas": 7,
                    "licencas_pendentes": 1,
                    "fases": {"Outorga de uso de recurso hídrico": 6,
                              "Licença Ambiental - Canteiro de obra": 1,
                              "Licença Ambiental Única": 1}},
        "condicionantes": {"pagina": 12, "xref": 842, "atualizado_em": "2026-09-07",
                           "finalizadas": 14, "no_prazo": 28, "pendentes": None,
                           "proximo_a_vencer": 1, "total_rodovia": 43,
                           "pct": [32.56, 65.12, 2.33]},
    },
    "RD103_048": {
        "indices": {"pagina": 11, "xref": 736, "atualizado_em": "2026-08-28",
                    "total_licencas": 9, "licencas_emitidas": 4,
                    "licencas_pendentes": 5,
                    "fases": {"Outorga de uso de recurso hídrico": 5,
                              "Licença Ambiental - Tanque Aéreo": 1,
                              "Licença Ambiental - Usina": 1,
                              "Licença Ambiental Única": 1,
                              "Registro Ambiental": 1}},
        "condicionantes": {"pagina": 12, "xref": 747, "atualizado_em": "2026-08-28",
                           "finalizadas": 11, "no_prazo": 1, "pendentes": None,
                           "proximo_a_vencer": 33, "total_rodovia": 45,
                           "pct": [24.44, 2.22, 73.33]},
    },
    "RD103_049": {
        "indices": {"pagina": 11, "xref": 749, "atualizado_em": "2026-09-07",
                    "total_licencas": 9, "licencas_emitidas": 4,
                    "licencas_pendentes": 5,
                    "fases": {"Outorga de uso de recurso hídrico": 5,
                              "Licença Ambiental - Tanque Aéreo": 1,
                              "Licença Ambiental - Usina": 1,
                              "Licença Ambiental Única": 1,
                              "Registro Ambiental": 1}},
        "condicionantes": {"pagina": 12, "xref": 761, "atualizado_em": "2026-09-07",
                           "finalizadas": 14, "no_prazo": 29, "pendentes": None,
                           "proximo_a_vencer": 1, "total_rodovia": 44,
                           "pct": [31.82, 65.91, 2.27]},
    },
    "RD104_042": {
        "indices": {"pagina": 11, "xref": 665, "atualizado_em": "2026-08-28",
                    "total_licencas": 8, "licencas_emitidas": 6,
                    "licencas_pendentes": 3,
                    "observacao": ("unico caso em que Emitidas + Pendentes (9) "
                                   "nao fecha com o cartao Total de Licencas (8); "
                                   "a rosca mostra 62,5%/37,5% = 5 Vigente / 3 "
                                   "Pendente sobre 8"),
                    "fases": {"Outorga de uso de recurso hídrico": 5,
                              "Licença Ambiental - Canteiro de obra": 1,
                              "Licença Ambiental - Tanque Aéreo e Usina de Concreto": 1,
                              "Licença Ambiental Única": 1}},
        "condicionantes": {"pagina": 12, "xref": 681, "atualizado_em": "2026-08-28",
                           "finalizadas": 10, "no_prazo": 36, "pendentes": None,
                           "proximo_a_vencer": None, "total_rodovia": 46,
                           "pct": [21.74, 78.26, None]},
    },
    "RD104_043": {
        "indices": {"pagina": 10, "xref": 630, "atualizado_em": "2026-09-07",
                    "total_licencas": 8, "licencas_emitidas": 6,
                    "licencas_pendentes": 2,
                    "fases": {"Outorga de uso de recurso hídrico": 5,
                              "Licença Ambiental - Canteiro de obra": 1,
                              "Licença Ambiental - Tanque Aéreo e Usina de Concreto": 1,
                              "Licença Ambiental Única": 1},
                    "tabela_n_licenca": ["202500009", "202500010", "202500011",
                                         "202500012", "202500013", "202500014",
                                         None, None]},
        "condicionantes": {"pagina": 11, "xref": 642, "atualizado_em": "2026-09-07",
                           "finalizadas": 10, "no_prazo": 1, "pendentes": None,
                           "proximo_a_vencer": 35, "total_rodovia": 46,
                           "pct": [21.74, 2.17, 76.09]},
    },
}

# regra ja validada em 52/52 (D-27)
ANCORA_QUINZENAL = dt.date(2026, 6, 6)


# --------------------------------------------------------------- utilitarios
def col(ref):
    return re.match(r"[A-Z]+", ref or "").group(0)


def lin(ref):
    return int(re.search(r"\d+", ref or "0").group(0))


def serial(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def ser2date(v):
    f = serial(v)
    if f is None or f < 1 or f > 80000:
        return None
    return EPOCH + dt.timedelta(days=int(f))


def iso(d):
    return d.isoformat() if isinstance(d, dt.date) else None


def d_de(s):
    return dt.date.fromisoformat(s) if s else None


def linhas(doc, nome_aba):
    """Devolve {n_linha: {letra_coluna: celula}} de uma aba."""
    for a in doc["abas"]:
        if a.get("nome") == nome_aba:
            R = defaultdict(dict)
            for c in a.get("celulas", []):
                R[lin(c["ref"])][col(c["ref"])] = c
            return dict(R), a
    return {}, None


def v(r, letra):
    c = r.get(letra)
    return c.get("v") if c else None


# --------------------------------------------------- 1. dissecar a planilha
def dissecar_planilha():
    """Copia para tmp/ (forca hidratacao do OneDrive - P-06) e disseca."""
    local = U.copiar_para_tmp(PLANILHA, subdir="ambiental")
    doc = M05.dissecar(local, com_celulas=True)
    doc["caminho_origem"] = PLANILHA
    doc["caminho_lido"] = local
    doc["stat_origem"] = U.stat_info(PLANILHA)
    return doc


def ficha_planilha(doc):
    f = {
        "arquivo": os.path.basename(PLANILHA),
        "caminho": PLANILHA,
        "tamanho_bytes": doc["tamanho_bytes"],
        "modificado": doc["modificado"],
        "sha1": doc.get("sha1"),
        "propriedades": doc["propriedades"],
        "abas": [], "nomes_definidos": doc["workbook"].get("nomes_definidos", []),
        "links_externos": doc["links_externos"],
        "tabelas_estruturadas": doc["tabelas"],
        "graficos_nativos": len(doc["graficos"]),
        "tabelas_dinamicas": [n for n in doc["partes"] if "pivot" in n.lower()],
        "webextensions": [n for n in doc["partes"] if "webextension" in n],
        "midia": doc["midia"]["quantidade"],
    }
    for a in doc["abas"]:
        f["abas"].append({
            "ordem": a["ordem"], "nome": a["nome"], "estado": a["estado"],
            "dimensao": a.get("dimensao"),
            "celulas_nao_vazias": a.get("celulas_nao_vazias"),
            "linhas_ocultas": a.get("oculta_linhas"),
            "colunas_ocultas": a.get("oculta_colunas"),
            "mescladas": a.get("mescladas", [])[:20],
            "validacoes": a.get("validacoes", []),
            "formatacao_condicional": a.get("formatacao_condicional"),
            "n_hyperlinks": len(a.get("hyperlinks", [])),
            "formulas_distintas": len(a.get("formulas", [])),
        })
    return f


def cabecalhos(doc, aba, linha):
    R, _ = linhas(doc, aba)
    r = R.get(linha, {})
    return {k: r[k].get("v") for k in sorted(r, key=lambda x: (len(x), x))}


# ---------------------------------------------- 2. registros estruturados
def mapa_empresa_obra(doc):
    R, _ = linhas(doc, "Empresas")
    mapa, detalhe = {}, []
    for L in sorted(R):
        if L < 4:
            continue
        razao, rod = v(R[L], "B"), v(R[L], "G")
        if not razao:
            continue
        obra = ("GO" + re.sub(r"[^0-9A-Za-z]", "", str(rod)).upper()
                .replace("GO", "", 1)) if rod else None
        mapa[razao] = obra
        detalhe.append({"linha": L, "celula_razao": "Empresas!B%d" % L,
                        "razao_social": razao, "celula_rodovia": "Empresas!G%d" % L,
                        "rodovia": rod, "obra": obra,
                        "cnpj": v(R[L], "C"), "cidade": v(R[L], "E"),
                        "representante": v(R[L], "H"), "status": v(R[L], "L")})
    return mapa, detalhe


def extrair_licencas(doc, mapa):
    R, _ = linhas(doc, "Licenças")
    out = []
    for L in sorted(R):
        if L < 3:
            continue
        r = R[L]
        emp = v(r, "D")
        if not emp:
            continue
        out.append({
            "linha": L,
            "n_licenca": v(r, "B"), "n_processo": v(r, "C"),
            "empresa": emp, "obra": mapa.get(emp),
            "fase": v(r, "E"),
            "data_emissao": iso(ser2date(v(r, "F"))),
            "data_validade": iso(ser2date(v(r, "G"))),
            "orgao": v(r, "H"), "atividade": v(r, "I"),
            "status": v(r, "J"), "responsavel": v(r, "K"),
            "latitude": v(r, "L"), "longitude": v(r, "M"),
            "link": v(r, "P"),
        })
    return out


def extrair_condicionantes(doc, mapa):
    R, _ = linhas(doc, "Condicionantes")
    out = []
    for L in sorted(R):
        if L < 3:
            continue
        r = R[L]
        resp = v(r, "I")
        if not resp:
            continue
        out.append({
            "linha": L,
            "item": v(r, "A"), "n_licenca": v(r, "B"),
            "n_condicionante": v(r, "C"),
            "descricao": (v(r, "D") or "")[:160],
            "evidencias_em": v(r, "E"),
            "data_entrega": iso(ser2date(v(r, "F"))),
            "data_entrega_serial": serial(v(r, "F")),
            "status_cache": v(r, "G"),
            "observacoes": v(r, "H"),
            "responsavel": resp, "obra": mapa.get(resp),
            "vence_em_cache": v(r, "J"),
            "proxima_entrega_cache": v(r, "K"),
        })
    return out


# ------------------------------- 3. reimplementacao da formula de status
def status_condicionante(cond, hoje):
    """Reimplementa Condicionantes!G (formula compartilhada, mestre em G3):

    J = SE(H="Atendida";"Atendida"; SE(H="Finalizada";"Finalizada"; F-HOJE()))
    K = SE(J="Atendida"; DATA(ANO(F)+1;MES(F);DIA(F)); SE(H="Finalizada";...;F))
    L = SE(J="Atendida"; K-HOJE(); J)
    G = SE(J="Finalizada";"Finalizada";
          SE(J="Atendida"; faixa(L); faixa(J)))
    faixa(x) = "Pendente" se x<=0; "Proximo do vencimento" se x<=30; senao "No prazo"
    """
    h = (cond.get("observacoes") or "").strip()
    F = cond.get("data_entrega_serial")
    t = (hoje - EPOCH).days
    if h == "Finalizada":
        return "Finalizada"
    if F is None:
        return None
    if h == "Atendida":
        Fd = EPOCH + dt.timedelta(days=int(F))
        try:
            K = dt.date(Fd.year + 1, Fd.month, Fd.day)
        except ValueError:
            K = dt.date(Fd.year + 1, Fd.month, Fd.day - 1)
        x = (K - EPOCH).days - t
    else:
        x = F - t
    if x <= 0:
        return "Pendente"
    if x <= 30:
        return "Próximo do vencimento"
    return "No prazo"


def kpis(licencas, condicionantes, hoje):
    """Recalcula, por obra, os KPIs dos dois paineis do Power BI."""
    out = {}
    obras = sorted({c["obra"] for c in condicionantes if c["obra"]} |
                   {l["obra"] for l in licencas if l["obra"]})
    for o in obras:
        ls = [l for l in licencas if l["obra"] == o]
        cs = [c for c in condicionantes if c["obra"] == o]
        cnt = Counter(status_condicionante(c, hoje) for c in cs)
        tot = len(cs)
        out[o] = {
            # painel 1 - Indices de Acompanhamento Ambiental
            "total_licencas": len(ls),
            "licencas_emitidas": sum(1 for l in ls if l["status"] == "Vigente"),
            "licencas_pendentes": sum(1 for l in ls
                                      if l["status"] == "Aguardando a liberação"),
            "fases": dict(Counter(l["fase"] for l in ls)),
            # painel 2 - Status das Condicionantes
            "total_condicionantes": tot,
            "finalizadas": cnt.get("Finalizada", 0),
            "no_prazo": cnt.get("No prazo", 0),
            "pendentes": cnt.get("Pendente", 0),
            "proximo_a_vencer": cnt.get("Próximo do vencimento", 0),
            "pct_finalizadas": round(100.0 * cnt.get("Finalizada", 0) / tot, 2) if tot else None,
            "pct_no_prazo": round(100.0 * cnt.get("No prazo", 0) / tot, 2) if tot else None,
            "pct_proximo": round(100.0 * cnt.get("Próximo do vencimento", 0) / tot, 2) if tot else None,
        }
    return out


# ------------------------------------------- 4. numeros escritos na secao 6
RE_LIC = [
    (re.compile(r"possui\s+(\d+)\s+atos autorizativos", re.I), "total_licencas"),
    (re.compile(r"possui\s+(\d+)\s+licen[çc]as solicitadas", re.I), "total_licencas"),
    (re.compile(r"(?:sendo\s+)?(\d+)\s+licen[çc]as emitidas", re.I),
     "licencas_emitidas"),
    (re.compile(r"(\d+)\s+com status de atendimento", re.I), "licencas_emitidas"),
    (re.compile(r"(\d+)\s+licen[çc]as? (?:com pend[êe]ncias|pendentes)", re.I),
     "licencas_pendentes"),
]
# "possui N atos autorizativos validos, TODOS com status ... 'No prazo', SEM ...
# pendencias documentais" -> emitidas = N e pendentes = 0
RE_TODOS = re.compile(r"atos autorizativos ambientais\s+v[áa]lidos,\s*todos", re.I)
RE_SEM_PEND = re.compile(r"sem registros de vencimentos iminentes ou pend", re.I)
RE_COND = re.compile(r"h[áa]\s+(\d+)\s+condicionantes?\s+pr[óo]xim", re.I)
RE_COND_ZERO = re.compile(r"n[ãa]o h[áa] condicionantes", re.I)
RE_PA = re.compile(r"PA[-\s]?(\d{3})")


def texto_secao6(estr):
    for k, val in estr.get("texto_secoes", {}).items():
        if k.strip().startswith("6"):
            return re.sub(r"\s+", " ", val)
    return ""


def numeros_do_texto(txt):
    n = {}
    for rx, campo in RE_LIC:
        m = rx.search(txt)
        if m and campo not in n:
            n[campo] = int(m.group(1))
    if RE_TODOS.search(txt) and n.get("total_licencas") is not None:
        n["licencas_emitidas"] = n["total_licencas"]
        n["emitidas_implicita_por_todos"] = True
    if RE_SEM_PEND.search(txt):
        n["licencas_pendentes"] = 0
    m = RE_COND.search(txt)
    if m:
        n["proximo_a_vencer"] = int(m.group(1))
    elif RE_COND_ZERO.search(txt):
        n["proximo_a_vencer"] = 0
    pa = RE_PA.search(txt)
    if pa:
        n["parecer_tecnico"] = "PA-" + pa.group(1)
    lau = re.search(r"Licen[çc]a Ambiental [ÚU]nica n[ºo°]\s*(\d+)", txt, re.I)
    if lau:
        n["lau_citada"] = lau.group(1)
    n["cita_qr_code"] = bool(re.search(r"QR Code", txt, re.I))
    n["cita_quinzenal"] = bool(re.search(r"quinzenal", txt, re.I))
    return n


# --------------------------------------------- 5. acervo ambiental da obra
RE_DOC = re.compile(r"^(FA|PA|RA|RI|NT)\s*[-–]\s*(\d{3})(?:\.(\d))?", re.I)
RE_PASTA_DATA = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")


def inventariar_acervo():
    acervo = {}
    for obra in U.OBRAS:
        raiz = os.path.join(U.BASE_OBRAS, obra, PASTA_AMB)
        item = {"raiz": raiz, "existe": U.exists(raiz), "pastas_visita": [],
                "documentos": [], "outras_pastas": []}
        if not item["existe"]:
            acervo[obra] = item
            continue
        for nome in U.listdir(raiz):
            p = os.path.join(raiz, nome)
            if U.isdir(p) and RE_PASTA_DATA.match(nome):
                item["pastas_visita"].append(nome)
            elif U.isdir(p):
                item["outras_pastas"].append(nome)
        for dp, _dn, fn in U.walk(raiz):
            rel = os.path.relpath(dp, raiz).replace("\\", "/")
            for f in fn:
                m = RE_DOC.match(f.replace("–", "-"))
                if not m:
                    item["documentos"].append({"pasta": rel, "arquivo": f,
                                               "tipo": None})
                    continue
                item["documentos"].append({
                    "pasta": rel, "arquivo": f,
                    "tipo": m.group(1).upper(), "numero": int(m.group(2)),
                    "revisao": m.group(3),
                    "assunto": ("SST" if re.search(r"\bSST\b", f) else
                                ("ambiental" if "ambiental" in U.norm(f) else None)),
                    "data_pasta": rel.split("/")[0]
                    if RE_PASTA_DATA.match(rel.split("/")[0]) else None,
                })
        acervo[obra] = item
    return acervo


def testar_quinzena(acervo, periodos):
    """Cada pasta-visita cai dentro de um periodo semanal; esse periodo e
    quinzenal-de-anexo? Regra D-27: (inicio - 2026-06-06) mod 14 == 0."""
    res = []
    for obra, item in acervo.items():
        for d in item["pastas_visita"]:
            data = dt.date.fromisoformat(d)
            achou = None
            for p in periodos:
                if d_de(p["inicio"]) <= data <= d_de(p["fim"]):
                    achou = p
                    break
            if achou:
                delta = (d_de(achou["inicio"]) - ANCORA_QUINZENAL).days
                res.append({"obra": obra, "data_visita": d,
                            "periodo_inicio": achou["inicio"],
                            "periodo_fim": achou["fim"],
                            "delta_dias_ancora": delta,
                            "eh_periodo_de_anexo": delta % 14 == 0})
            else:
                res.append({"obra": obra, "data_visita": d, "periodo_inicio": None,
                            "eh_periodo_de_anexo": None})
    return res


def periodos_semanais(inicio, fim):
    """Regime vigente (D-02): sabado -> sexta desde 2026-01-03."""
    out = []
    d = inicio
    while d <= fim:
        out.append({"inicio": iso(d), "fim": iso(d + dt.timedelta(days=6))})
        d += dt.timedelta(days=7)
    return out


# ------------------------------------------------ 6. leitura dos PDFs PA-NNN
RE_VERIF = re.compile(
    r"verifica[çc][õo]es realizadas em campo em (\d{1,2}|1º)\s*d?e?\s*"
    r"(janeiro|fevereiro|mar[çc]o|abril|maio|junho|julho|agosto|setembro|"
    r"outubro|novembro|dezembro) de (\d{4})", re.I)
RE_QUINZ = re.compile(r"Relat[óo]rio Quinzenal n[ºo°]?\s*(\d{3})", re.I)
MESES = {"janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4,
         "maio": 5, "junho": 6, "julho": 7, "agosto": 8, "setembro": 9,
         "outubro": 10, "novembro": 11, "dezembro": 12}


def ler_pareceres(acervo):
    try:
        import fitz
    except ImportError:
        return [{"erro": "PyMuPDF indisponivel"}]
    out = []
    for obra, item in acervo.items():
        for d in item["documentos"]:
            if d.get("tipo") != "PA":
                continue
            p = os.path.join(U.BASE_OBRAS, obra, PASTA_AMB,
                             *d["pasta"].split("/")) if d["pasta"] != "." else \
                os.path.join(U.BASE_OBRAS, obra, PASTA_AMB)
            caminho = os.path.join(p, d["arquivo"])
            reg = {"obra": obra, "numero": d["numero"],
                   "id": "PA-%03d" % d["numero"], "arquivo": d["arquivo"],
                   "pasta": d["pasta"]}
            try:
                doc = fitz.open(U.lp(caminho))
                reg["n_paginas"] = doc.page_count
                t = re.sub(r"[ \t]+", " ", doc[0].get_text())
                doc.close()
            except Exception as e:  # noqa: BLE001
                reg["erro"] = "%s: %s" % (type(e).__name__, e)
                out.append(reg)
                continue
            t1 = re.sub(r"\s+", " ", t)
            m = re.search(r"(\d{2}/\d{2}/\d{4})", t1)
            reg["data_versao"] = m.group(1) if m else None
            m = re.search(r"Vers[ãa]o (\d+)", t1)
            reg["versao"] = m.group(1) if m else None
            m = re.search(r"(GO-\d+[A-Za-z_]*)", t1)
            reg["rodovia_no_documento"] = m.group(1) if m else None
            m = RE_VERIF.search(t1)
            if m:
                dia = 1 if m.group(1).startswith("1º") else int(m.group(1))
                mes = MESES.get(U.norm(m.group(2)), None)
                if mes:
                    reg["data_verificacao_campo"] = iso(
                        dt.date(int(m.group(3)), mes, dia))
            reg["relatorios_quinzenais_citados"] = sorted(
                set(RE_QUINZ.findall(t1)))
            reg["titulo"] = t1[:120]
            out.append(reg)
    out.sort(key=lambda x: x["numero"])
    return out


def copias_mensais():
    """A mesma planilha aparece copiada em `5. Relatorio Mensal da Obra\\
    RM-NNN.1 - <OBRA>\\(Arquivos )Auxiliar(es)`. Sao versoes de outra epoca,
    com duas abas a mais (`graficos-FA`, `graficos-licenciamento`) e layout de
    colunas diferente em `Condicionantes`. Nao sao a fonte do painel semanal."""
    out = []
    for obra in U.OBRAS:
        raiz = os.path.join(U.BASE_OBRAS, obra, "3_Gestao Tecnica de Obra",
                            "5. Relatório Mensal da Obra")
        if not U.exists(raiz):
            continue
        for rm in U.listdir(raiz):
            prm = os.path.join(raiz, rm)
            if not U.isdir(prm):
                continue
            for aux in U.listdir(prm):
                paux = os.path.join(prm, aux)
                if not U.isdir(paux) or not U.norm(aux).startswith("auxiliar") \
                        and not U.norm(aux).startswith("arquivos auxiliar"):
                    continue
                for f in U.listdir(paux):
                    if "licen" in U.norm(f) and f.lower().endswith(".xlsx"):
                        st = U.stat_info(os.path.join(paux, f))
                        out.append({"obra": obra, "rm": rm, "pasta": aux,
                                    "arquivo": f, **st})
    return out


# ------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sem-pdf", action="store_true")
    args = ap.parse_args()

    U.assert_read_only(SAIDA)
    os.makedirs(os.path.dirname(SAIDA), exist_ok=True)

    print("-> dissecando a planilha ambiental...", flush=True)
    doc = dissecar_planilha()
    mapa, empresas = mapa_empresa_obra(doc)
    licencas = extrair_licencas(doc, mapa)
    condicionantes = extrair_condicionantes(doc, mapa)
    print("   %d licencas, %d condicionantes, %d empresas"
          % (len(licencas), len(condicionantes), len(empresas)), flush=True)

    res = {
        "gerado_em": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
        "pendencias_fechadas": ["P-09", "T-05"],
        "fonte_humana": ("Emmily: o Dashboard Ambiental e um Power BI publicado "
                         "e a fonte dele e a planilha "
                         "04_Acompanhamento - Licenca Ambiental.xlsx"),
        "planilha": ficha_planilha(doc),
        "cabecalhos": {
            "Licenças!linha2": cabecalhos(doc, "Licenças", 2),
            "Condicionantes!linha2": cabecalhos(doc, "Condicionantes", 2),
            "Empresas!linha3": cabecalhos(doc, "Empresas", 3),
        },
        "dashboard_formulas": [],
        "empresas": empresas,
        "mapa_empresa_obra_confere": mapa == EMPRESA_OBRA_ESPERADO,
        "licencas": licencas,
        "condicionantes_resumo": {
            "n": len(condicionantes),
            "por_obra": dict(Counter(c["obra"] for c in condicionantes)),
            "status_em_cache": dict(Counter(c["status_cache"]
                                            for c in condicionantes)),
        },
        "condicionantes": condicionantes,
    }

    # formulas da aba Dashboard = definicao literal dos KPIs
    Rd, _ = linhas(doc, "Dashboard")
    for L in sorted(Rd):
        for letra, c in sorted(Rd[L].items()):
            if c.get("f"):
                res["dashboard_formulas"].append(
                    {"celula": "Dashboard!%s%d" % (letra, L),
                     "valor_em_cache": c.get("v"), "formula": c["f"]})

    # ------------------------------------------------ confronto com a secao 6
    estr = json.load(open(ESTRUTURA, encoding="utf-8"))
    conf = []
    for rot, e in sorted(estr.items()):
        txt = texto_secao6(e)
        n = numeros_do_texto(txt)
        painel = PAINEIS_LIDOS.get(rot, {})
        at = ((painel.get("condicionantes") or {}).get("atualizado_em")
              or (painel.get("indices") or {}).get("atualizado_em"))
        recalc = kpis(licencas, condicionantes, d_de(at)) if at else {}
        k = recalc.get(e["obra"], {})
        pi = painel.get("indices", {})
        pc = painel.get("condicionantes", {})
        conf.append({
            "relatorio": rot, "obra": e["obra"],
            "periodo": [e["data_inicio"], e["data_fim"]],
            "painel_atualizado_em": at,
            "painel_atualizado_igual_fim_periodo": at == e["data_fim"],
            "texto": n,
            "painel_lido": {"indices": pi, "condicionantes": pc},
            "planilha_recalculada_na_data": k,
            "confronto": {
                "total_licencas_texto_x_painel":
                    [n.get("total_licencas"), pi.get("total_licencas")],
                "emitidas_texto_x_painel":
                    [n.get("licencas_emitidas"), pi.get("licencas_emitidas")],
                "pendentes_texto_x_painel":
                    [n.get("licencas_pendentes"), pi.get("licencas_pendentes")],
                "proximo_a_vencer_texto_x_painel":
                    [n.get("proximo_a_vencer"),
                     pc.get("proximo_a_vencer") or (0 if "proximo_a_vencer" in pc
                                                    else None)],
                "total_condicionantes_painel_x_planilha":
                    [pc.get("total_rodovia"), k.get("total_condicionantes")],
                "finalizadas_painel_x_planilha":
                    [pc.get("finalizadas"), k.get("finalizadas")],
            },
            "lau_citada_x_planilha": {
                "citada": n.get("lau_citada"),
                "na_planilha": next((l["n_licenca"] for l in licencas
                                     if l["obra"] == e["obra"]
                                     and l["fase"] == "Licença Ambiental Única"),
                                    None),
            },
        })
    res["confronto_secao6"] = conf

    # marca os acertos
    def ok(a, b):
        return (a is not None and b is not None and a == b) or \
               (a in (0, None) and b in (0, None) and (a, b) != (None, None))

    res["placar"] = {
        "texto_x_painel_total_licencas": sum(
            1 for c in conf if ok(*c["confronto"]["total_licencas_texto_x_painel"])),
        "texto_x_painel_emitidas": sum(
            1 for c in conf if ok(*c["confronto"]["emitidas_texto_x_painel"])),
        "texto_x_painel_pendentes": sum(
            1 for c in conf if ok(*c["confronto"]["pendentes_texto_x_painel"])),
        "texto_x_painel_proximo_a_vencer": sum(
            1 for c in conf
            if ok(*c["confronto"]["proximo_a_vencer_texto_x_painel"])),
        "lau_citada_confere": sum(
            1 for c in conf
            if c["lau_citada_x_planilha"]["citada"]
            == c["lau_citada_x_planilha"]["na_planilha"]),
        "painel_x_planilha_total_condicionantes": sum(
            1 for c in conf
            if ok(*c["confronto"]["total_condicionantes_painel_x_planilha"])),
        "painel_x_planilha_finalizadas": sum(
            1 for c in conf
            if ok(*c["confronto"]["finalizadas_painel_x_planilha"])),
        "n_relatorios": len(conf),
    }

    # A tabela do painel 1 e linha-a-linha a aba `Licencas`?
    # (so para os relatorios em que a tabela foi transcrita da imagem)
    bat = []
    for rot, p in PAINEIS_LIDOS.items():
        alvo = (p.get("indices") or {}).get("tabela_n_licenca")
        if not alvo:
            continue
        obra = estr[rot]["obra"]
        na_planilha = {str(l["n_licenca"]) for l in licencas if l["obra"] == obra}
        bat.append({
            "relatorio": rot, "obra": obra,
            "n_licenca_no_painel": alvo,
            "encontradas_na_planilha": [a for a in alvo if a and a in na_planilha],
            "ausentes_na_planilha": [a for a in alvo if a and a not in na_planilha],
            "n_licencas_na_planilha_local": len(na_planilha),
        })
    res["tabela_painel_x_aba_licencas"] = bat

    # copias mensais da mesma planilha (5. Relatorio Mensal da Obra)
    res["copias_mensais"] = copias_mensais()

    # KPIs recalculados em datas de interesse
    res["kpis_por_data"] = {
        d: kpis(licencas, condicionantes, dt.date.fromisoformat(d))
        for d in ("2026-05-13", "2026-08-28", "2026-09-04", "2026-09-07")
    }

    # ---------------------------------------------------- acervo e quinzena
    print("-> inventariando o acervo ambiental...", flush=True)
    acervo = inventariar_acervo()
    per = periodos_semanais(dt.date(2025, 11, 1), dt.date(2026, 9, 30))
    # regime sabado->sexta so vale de 2026-01-03 em diante
    per = [p for p in per if d_de(p["inicio"]) >= dt.date(2026, 1, 3)] or per
    base = dt.date(2026, 1, 3)
    per = periodos_semanais(base, dt.date(2026, 9, 26))
    res["acervo_ambiental"] = {
        o: {"raiz": a["raiz"], "pastas_visita": a["pastas_visita"],
            "outras_pastas": a["outras_pastas"],
            "por_tipo": dict(Counter(d.get("tipo") for d in a["documentos"])),
            "documentos": a["documentos"]}
        for o, a in acervo.items()
    }
    res["teste_quinzenal"] = testar_quinzena(acervo, per)
    res["teste_quinzenal_placar"] = {
        "visitas_em_periodo_de_anexo": sum(1 for t in res["teste_quinzenal"]
                                           if t["eh_periodo_de_anexo"]),
        "visitas_fora": sum(1 for t in res["teste_quinzenal"]
                            if t["eh_periodo_de_anexo"] is False),
        "sem_periodo": sum(1 for t in res["teste_quinzenal"]
                           if t["eh_periodo_de_anexo"] is None),
        "total": len(res["teste_quinzenal"]),
        "observacao": ("so vale para visitas a partir de 2026-01-03, quando "
                       "comeca o regime sabado->sexta (D-02)"),
    }
    # so as visitas do regime atual
    recorte = [t for t in res["teste_quinzenal"]
               if t["data_visita"] >= "2026-01-03"]
    res["teste_quinzenal_placar"]["regime_atual"] = {
        "dentro": sum(1 for t in recorte if t["eh_periodo_de_anexo"]),
        "fora": sum(1 for t in recorte if t["eh_periodo_de_anexo"] is False),
        "total": len(recorte),
    }
    # o ciclo so estabiliza a partir de 2026-07-04
    est = [t for t in res["teste_quinzenal"] if t["data_visita"] >= "2026-07-04"]
    res["teste_quinzenal_placar"]["desde_2026_07_04"] = {
        "dentro": sum(1 for t in est if t["eh_periodo_de_anexo"]),
        "fora": sum(1 for t in est if t["eh_periodo_de_anexo"] is False),
        "total": len(est),
    }

    # rotacao semanal das visitas: dia da semana por obra
    DIAS = ["segunda", "terça", "quarta", "quinta", "sexta", "sábado", "domingo"]
    rot = defaultdict(list)
    por_quinzena = defaultdict(dict)
    for t in est:
        d = dt.date.fromisoformat(t["data_visita"])
        rot[t["obra"]].append(DIAS[d.weekday()])
        por_quinzena[t["periodo_inicio"]][t["obra"]] = t["data_visita"]
    res["rotacao_visitas"] = {
        "dia_da_semana_por_obra": {o: dict(Counter(v_)) for o, v_ in rot.items()},
        "ordem_por_quinzena": {
            k: [o for o, _ in sorted(v_.items(), key=lambda x: x[1])]
            for k, v_ in sorted(por_quinzena.items())},
    }

    # numeracao corporativa: PA / RA / RI sao sequencias unicas entre obras?
    for tipo in ("PA", "RA", "RI", "FA", "NT"):
        nums = defaultdict(list)
        for o, a in acervo.items():
            for d in a["documentos"]:
                if d.get("tipo") == tipo:
                    nums[d["numero"]].append(o)
        colisao = {n: sorted(set(v_)) for n, v_ in nums.items()
                   if len(set(v_)) > 1 or len(v_) > (2 if tipo == "FA" else 1)}
        res.setdefault("numeracao", {})[tipo] = {
            "n_documentos": sum(len(v_) for v_ in nums.values()),
            "faixa": [min(nums), max(nums)] if nums else None,
            "numeros_repetidos_entre_obras": colisao,
            "sequencia_unica_entre_obras": not colisao,
            "por_obra": {o: sorted(d["numero"] for d in a["documentos"]
                                   if d.get("tipo") == tipo)
                         for o, a in acervo.items()},
        }

    # ------------------------------------------------------ pareceres PA-NNN
    if not args.sem_pdf:
        print("-> lendo os PDFs PA-NNN...", flush=True)
        res["pareceres"] = ler_pareceres(acervo)
        # ordem de emissao x ordem da visita
        leva = [p for p in res["pareceres"] if 252 <= p["numero"] <= 255]
        res["leva_252_255"] = sorted(
            [{"id": p["id"], "obra": p["obra"],
              "data_verificacao_campo": p.get("data_verificacao_campo"),
              "data_versao": p.get("data_versao"),
              "relatorios_quinzenais_citados":
                  p.get("relatorios_quinzenais_citados")} for p in leva],
            key=lambda x: x["id"])

    with open(SAIDA, "w", encoding="utf-8") as fh:
        json.dump(res, fh, ensure_ascii=False, indent=1)
    print("OK ->", SAIDA)
    print(json.dumps(res["placar"], ensure_ascii=False, indent=1))
    print(json.dumps(res["teste_quinzenal_placar"], ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
