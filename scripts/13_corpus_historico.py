# -*- coding: utf-8 -*-
r"""
13_corpus_historico.py — CAMADA A: varredura programatica de TODO o corpus de
Relatorios Semanais de Acompanhamento de Obras do Programa (~185 PDFs, 4 obras).

O que faz
---------
1. INDICE HISTORICO  -> output/inventarios/relatorios_semanais_historico.json (+ .csv)
   Para cada PDF principal de cada pasta-semana extrai: obra, numero, revisao,
   periodo (pasta + capa), data de emissao (Termo de Encerramento + mtime),
   nº de paginas, formatos de papel, secoes nivel 1/2, subsecoes 5.2.x,
   anexos (com categoria normalizada), legendas e contagem de imagens.

2. CASOS RAROS       -> output/comparacoes/casos_raros.json
   (A) NC emitida pela Órgão Contratante com conteudo positivo
   (B) Relato real de riscos/interferencias em campo
   (C) Revisoes .2/.3 (nome, pasta, conteudo, data, diff .1 x .2)
   (D) Outras excecoes: palavras-chave + anomalias estatisticas

3. MARCOS DE MUDANCA -> embutido no JSON de casos raros + insumo do MD

4. ANALISE TEXTUAL   -> output/comparacoes/corpus_textual.json
   Clustering de paragrafos: FIXO / TEMPLATE / ESTRUTURADO / LIVRE.

Eficiencia
----------
- Le apenas o texto das paginas A4 retrato (largura ~595 pt) com page.get_text("text").
- Nao extrai imagens (so conta via page.get_images()).
- Cache de texto por relatorio em output/extracao/texto/*.txt.
- Progresso incremental em output/inventarios/_historico_parcial.jsonl
  (chave: caminho + tamanho + mtime). Reexecutar nao refaz o que ja foi feito.

REGRA: nunca escreve no SharePoint. Toda saida passa por U.assert_read_only().
"""
import os
import sys
import re
import csv
import json
import time
import argparse
import datetime
import difflib
import unicodedata
from collections import Counter, OrderedDict, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

import fitz  # PyMuPDF

# ---------------------------------------------------------------- caminhos ---
INV = os.path.join(U.OUT, "inventarios", "inventario_relatorios_semanais.json")
DIR_INV = os.path.join(U.OUT, "inventarios")
DIR_CMP = os.path.join(U.OUT, "comparacoes")
DIR_TXT = os.path.join(U.OUT, "extracao", "texto")
PARCIAL = os.path.join(DIR_INV, "_historico_parcial.jsonl")

# ---------------------------------------------------------------- formatos ---
FORMATOS = [
    ("A4 retrato", 595, 842), ("A4 paisagem", 842, 595),
    ("A3 retrato", 842, 1191), ("A3 paisagem", 1191, 842),
    ("A2 retrato", 1191, 1684), ("A2 paisagem", 1684, 1191),
    ("A1 retrato", 1684, 2384), ("A1 paisagem", 2384, 1684),
    ("A0 retrato", 2384, 3370), ("A0 paisagem", 3370, 2384),
    ("Carta retrato", 612, 792), ("Carta paisagem", 792, 612),
]


def formato(w, h, tol=14):
    for nome, fw, fh in FORMATOS:
        if abs(w - fw) <= tol and abs(h - fh) <= tol:
            return nome
    return "%dx%d pt" % (round(w), round(h))


# ----------------------------------------------------------------- regexes ---
# Nome do documento principal. Cobre as tres nomenclaturas historicas:
#   RO-001.1 - RELATORIO SEMANAL ... - 101.pdf     (set/2025)
#   RO_021.1 RD-101.pdf / RO_031.2 - RD-101.pdf    (atual)
#   4º Relatorio Semanal de Acompanhamento de Obras - RD101.pdf  (out-nov/2025)
RE_RO = re.compile(r"^RO[\s_\-]*0*(\d{1,3})(?:\.(\d+))?\b", re.IGNORECASE)
RE_RO_EXTENSO = re.compile(
    r"^(\d{1,3})\s*[ºo°ª]?\s*Relat[oó]rio\s+Semanal\s+de\s+Acompanhamento", re.IGNORECASE)

RE_TOC_LEADER = re.compile(r"\.{3,}\s*(\d{1,3})\s*$")
RE_NUM_SO = re.compile(r"^(\d{1,2}(?:\.\d{1,2}){0,2})\.?$")
RE_NUM_TIT = re.compile(r"^(\d{1,2}(?:\.\d{1,2}){0,2})\.?\s+(\S.*)$")
RE_LEGENDA = re.compile(
    r"^\s*(Figura|Gr[aá]fico|Tabela|Quadro|Foto(?:grafia)?)\s+(\d+)\s*[\-–—:]\s*(.+?)\s*$",
    re.IGNORECASE)
RE_ANEXO = re.compile(r"^\s*Anexo\s+([IVXLC]+|\d+)\s*[\-–—:]\s*(.+?)\s*$", re.IGNORECASE)
RE_DATA = re.compile(r"\b(\d{2})/(\d{2})/(\d{4})\b")
RE_PERIODO = re.compile(
    r"per[ií]odo\s*:?\s*(?:de\s*)?(\d{2}/\d{2}/\d{4})\s*(?:a|at[ée]|-|–)\s*(\d{2}/\d{2}/\d{4})",
    re.IGNORECASE)
RE_CAPA_CAMPO = re.compile(r"^([A-ZÀ-ÚÇ][A-ZÀ-ÚÇ º°/\.\-]{2,40}):\s*(.*)$")

MESES = {"janeiro": 1, "fevereiro": 2, "marco": 3, "abril": 4, "maio": 5, "junho": 6,
         "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11,
         "dezembro": 12}
# "Cidade, 24 de fevereiro de 2026." -> data de emissao do Termo de Encerramento
RE_DATA_EXTENSO = re.compile(
    r"([A-Za-zÀ-ú\s]{3,25}),\s*(\d{1,2})\s*de\s*([a-zç]+)\s*de\s*(\d{4})", re.IGNORECASE)

# (D) — termos de excecao. Regex por termo, aplicados sobre texto normalizado.
TERMOS_EXCECAO = OrderedDict([
    ("paralisacao",            r"paralisa(?:cao|coes|da|das|do|dos|r|ram|ndo)"),
    ("chuva_intensa",          r"chuva[s]?\s+(?:intensa|extrema|forte|volumosa|persistente)"),
    ("chuva_impacto",          r"(?:devido|em\s+razao|por\s+conta)\s+(?:a|as|ao|aos|da|das|de|do|dos)?\s*(?:chuva|precipitacao)"),
    ("impraticavel",           r"impratic[aá]vel|impraticabilidade"),
    ("sem_execucao",           r"sem\s+execu(?:cao|coes)|nao\s+houve\s+execu(?:cao)|nenhuma\s+atividade\s+(?:foi\s+)?execut"),
    ("acidente",               r"\bacidente[s]?\b|\bsinistro[s]?\b|\bincidente[s]?\b"),
    ("ata_extraordinaria",     r"ata\s+extraordin[aá]ria|reuniao\s+extraordin[aá]ria"),
    ("feriado",                r"\bferiado[s]?\b|\bponto\s+facultativo\b|\brecesso\b"),
    ("aditivo",                r"termo\s+aditivo|\baditivo[s]?\b|aditamento"),
    ("nova_licenca",           r"nova\s+licen[cç]a|renova(?:cao|r)\s+d[ae]\s+licen[cç]a|licen[cç]a\s+de\s+opera"),
    ("suspensao",              r"suspens(?:ao|oes)|suspens[oa][s]?\s|suspender"),
    ("embargo",                r"\bembargo[s]?\b|embargad[oa]"),
    ("atraso",                 r"\batraso[s]?\b|atrasad[oa]|em\s+atraso"),
    ("notificacao",            r"\bnotifica(?:cao|coes)\b|notificad[oa]"),
    ("recuperacao_cronograma", r"recupera(?:cao|r)\s+d[eo]\s+cronograma|plano\s+de\s+a(?:cao|celeracao)|recomposicao\s+de\s+cronograma"),
    ("mudanca_contratual",     r"altera(?:cao|coes)\s+contratual|repactua|reequilibrio|distrato|rescis"),
    ("desvio_cronograma",      r"desvio\s+de\s+cronograma"),
    ("nao_conformidade_ativa", r"\bnc\s*[\-–]?\s*\d|nao\s+conformidade\s+n[º°o]?\s*\d"),
    ("greve",                  r"\bgreve[s]?\b"),
    ("interdicao",             r"interdi(?:cao|tad)"),
    ("desapropriacao",         r"desapropria(?:cao|coes)"),
    ("interferencia_rede",     r"interfer[eê]ncia[s]?\s+(?:com|de|da|do)\s+(?:rede|adutora|linha|fibra|gasoduto|energia)"),
    ("circular_tecnica",       r"circular\s+t[eé]cnica"),
    ("mobilizacao_desmob",     r"desmobiliza(?:cao|r|ram)"),
])

# frases que indicam ausencia de ocorrencia (secoes 8.x e 9/10)
RE_NEGATIVO = re.compile(
    r"nao\s+houve|nao\s+foram|nao\s+ha\b|nao\s+existe|inexist|nao\s+se\s+aplica|"
    r"sem\s+registro|nenhum[ao]?\s|nada\s+a\s+(?:relatar|registrar)|"
    r"nao\s+foi\s+(?:identificad|registrad|constatad|emitid)")


def norm(s):
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.replace("ﬁ", "fi").replace("ﬂ", "fl").replace("‐", "-")
    s = s.replace("‑", "-").replace("‒", "-").replace("–", "-")
    s = s.replace("—", "-").replace("’", "'").replace(" ", " ")
    s = re.sub(r"[-¿̀-ͯ-]", " ", s)
    return re.sub(r"\s+", " ", s).strip().lower()


def limpar(s):
    """Limpeza leve preservando acento/caixa (para exibir o texto)."""
    s = s.replace("ﬁ", "fi").replace("ﬂ", "fl").replace(" ", " ")
    s = re.sub(r"[-]", "", s)
    return re.sub(r"[ \t]+", " ", s).strip()


def categoria_anexo(t):
    n = norm(t)
    if re.search(r"\bpa[\s\-]?\d", n) or ("parecer" in n and "tecnico" in n):
        return "PARECER_TECNICO"
    if re.search(r"\bnc[\s\-]?\d", n) or "nao conformidade" in n:
        return "NAO_CONFORMIDADE"
    if "desvio" in n and "cronograma" in n:
        return "DESVIO_CRONOGRAMA"
    if "unifilar" in n:
        return "DIAGRAMA_UNIFILAR"
    if "diario" in n:
        return "DIARIO_DE_OBRAS"
    if "ficha" in n and ("campo" in n or "acompanhamento" in n):
        return "FICHA_DE_CAMPO"
    if "ambient" in n:
        return "ACOMP_AMBIENTAL"
    if "seguranca" in n or re.search(r"\bsst\b", n) or "saude" in n:
        return "ACOMP_SST"
    if "ata" in n and "reuni" in n:
        return "ATA_REUNIAO"
    if "circular" in n:
        return "CIRCULAR_TECNICA"
    if "avanco" in n and "global" in n:
        return "AVANCO_GLOBAL"
    if "fotograf" in n:
        return "RELATORIO_FOTOGRAFICO"
    if "licenca" in n:
        return "LICENCA_AMBIENTAL"
    if "cronograma" in n:
        return "CRONOGRAMA"
    return "OUTRO: " + limpar(t)[:48]


# ------------------------------------------------------- selecao dos PDFs ----
def parse_nome(nome):
    """Devolve (numero, revisao, estilo) a partir do nome do arquivo, ou None."""
    base = os.path.splitext(nome)[0]
    m = RE_RO.match(base)
    if m:
        return int(m.group(1)), (int(m.group(2)) if m.group(2) else None), "RO_NNN"
    m = RE_RO_EXTENSO.match(base)
    if m:
        return int(m.group(1)), None, "EXTENSO"
    return None


def candidatos(entrada):
    """Devolve (principal, obsoletos[]) de uma pasta-semana."""
    principais, obsoletos = [], []
    for f in entrada.get("arquivos", []):
        if f["ext"] != ".pdf":
            continue
        p = parse_nome(f["nome"])
        if not p:
            continue
        numero, rev, estilo = p
        reg = dict(f, numero=numero, revisao=rev, estilo_nome=estilo)
        sub = norm(f.get("subpasta", ""))
        if not sub:
            principais.append(reg)
        elif "obsolet" in sub:
            obsoletos.append(reg)
    # se houver mais de um no raiz, o principal e o de maior revisao
    principais.sort(key=lambda r: ((r["revisao"] or 0), r["tamanho_bytes"]))
    principal = principais[-1] if principais else None
    return principal, obsoletos, principais[:-1] if len(principais) > 1 else []


# ------------------------------------------------------------ leitura PDF ----
def ler_pdf(caminho):
    """Le o PDF devolvendo texto das paginas A4 retrato + metadados de todas."""
    doc = fitz.open(U.lp(caminho))
    paginas, linhas, n_imgs = [], [], 0
    try:
        for i, page in enumerate(doc):
            r = page.rect
            fmt = formato(r.width, r.height)
            try:
                nimg = len(page.get_images(full=False))
            except Exception:
                nimg = 0
            n_imgs += nimg
            paginas.append({"n": i + 1, "formato": fmt, "imgs": nimg,
                            "w": round(r.width), "h": round(r.height)})
            if not fmt.startswith("A4 retrato"):
                continue
            txt = page.get_text("text")
            for ln in txt.split("\n"):
                ln = limpar(ln)
                if ln:
                    linhas.append((i + 1, ln))
        meta = dict(doc.metadata or {})
        n_paginas = doc.page_count
    finally:
        doc.close()
    return paginas, linhas, n_imgs, meta, n_paginas


def parse_toc(linhas):
    """Indice: linhas com dot-leader. Numero e titulo podem estar separados."""
    toc, i = [], 0
    # o indice esta nas primeiras paginas (<= 4)
    L = [(p, t) for p, t in linhas if p <= 4]
    while i < len(L):
        _, t = L[i]
        num = None
        m0 = RE_NUM_SO.match(t)
        if m0 and i + 1 < len(L):
            num = m0.group(1)
            i += 1
            _, t = L[i]
        m = RE_TOC_LEADER.search(t)
        if not m:
            i += 1
            continue
        pagina = int(m.group(1))
        titulo = RE_TOC_LEADER.sub("", t).strip(" .")
        if num is None:
            m2 = RE_NUM_TIT.match(titulo)
            if m2:
                num, titulo = m2.group(1), m2.group(2)
        titulo = titulo.strip(" .:")
        if titulo:
            toc.append({"num": num, "titulo": titulo, "pagina_indicada": pagina,
                        "nivel": (num.count(".") + 1) if num else None})
        i += 1
    return toc


def fatiar_secoes(linhas, toc):
    """Divide o corpo em secoes usando os titulos do indice como ancora."""
    alvo = {}
    for e in toc:
        if not e["num"]:
            continue
        k = norm("%s %s" % (e["num"], e["titulo"])).rstrip(" .:")
        alvo.setdefault(k, e)
        k2 = norm(e["titulo"]).rstrip(" .:")
        alvo.setdefault("__t__" + k2, e)

    corpo = [(p, t) for p, t in linhas if p >= 2]
    marcas = []
    vistos = set()
    for idx, (p, t) in enumerate(corpo):
        if RE_TOC_LEADER.search(t):
            continue
        base = norm(t).rstrip(" .:")
        e = alvo.get(base)
        if e is None:
            m = RE_NUM_TIT.match(t)
            if m:
                e = alvo.get("__t__" + norm(m.group(2)).rstrip(" .:"))
                if e and e["num"] != m.group(1):
                    e = None
        if e is None:
            continue
        chave = "%s|%s" % (e["num"], e["titulo"])
        if chave in vistos:
            continue
        vistos.add(chave)
        marcas.append((idx, e, p))

    marcas.sort()
    secoes = OrderedDict()
    for j, (idx, e, p) in enumerate(marcas):
        fim = marcas[j + 1][0] if j + 1 < len(marcas) else len(corpo)
        txt = [l for _, l in corpo[idx + 1:fim]]
        txt = [l for l in txt if not re.fullmatch(r"\d{1,3}", l)]
        secoes["%s %s" % (e["num"], e["titulo"])] = {
            "num": e["num"], "titulo": e["titulo"], "pagina_real": p,
            "texto": "\n".join(txt).strip(),
        }
    return secoes


ROMANOS = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X",
           "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX"]


def anexos_por_folha_separadora(linhas, pagina_anexos):
    """Modelo intermediario: cada anexo e uma folha A4 com so o titulo.
    Detecta paginas A4 depois da secao ANEXOS que tenham 1-3 linhas curtas."""
    porpag = OrderedDict()
    for p, t in linhas:
        if pagina_anexos and p < pagina_anexos:
            continue
        if re.fullmatch(r"\d{1,3}", t):
            continue
        porpag.setdefault(p, []).append(t)
    out = []
    for p, ls in porpag.items():
        ls = [l for l in ls if norm(l) not in ("anexos",)
              and not re.match(r"^\d{1,2}\.\s*anexos$", norm(l))]
        if not ls or len(ls) > 3:
            continue
        tit = " ".join(ls).strip()
        if len(tit) < 6 or len(tit) > 150:
            continue
        if norm(tit).startswith("termo de encerramento"):
            continue
        out.append({"ordem": ROMANOS[len(out)] if len(out) < len(ROMANOS) else str(len(out) + 1),
                    "titulo": tit, "categoria": categoria_anexo(tit)})
    return out


def achar_secao(secoes, *padroes):
    """Acha secao cujo titulo normalizado casa com algum padrao (regex)."""
    for k, v in secoes.items():
        n = norm(v["titulo"])
        for pad in padroes:
            if re.search(pad, n):
                return k, v
    return None, None


def analisar(reg, obra, entrada):
    caminho = reg["caminho"]
    paginas, linhas, n_imgs, meta, n_pag = ler_pdf(caminho)
    texto_a4 = "\n".join(t for _, t in linhas)

    r = OrderedDict()
    r["obra"] = obra
    r["pasta"] = entrada["nome"]
    r["pasta_data_inicio"] = entrada.get("data_inicio")
    r["pasta_data_fim"] = entrada.get("data_fim")
    r["pasta_numero"] = entrada.get("numero")
    r["arquivo"] = reg["nome"]
    r["subpasta"] = reg.get("subpasta", "")
    r["caminho"] = caminho
    r["estilo_nome"] = reg["estilo_nome"]
    r["numero"] = reg["numero"]
    r["revisao"] = reg["revisao"]
    r["tamanho_bytes"] = reg["tamanho_bytes"]
    r["mtime"] = reg["modificado"]
    r["n_paginas"] = n_pag
    r["n_imagens"] = n_imgs
    r["formatos"] = Counter(p["formato"] for p in paginas).most_common()
    r["paginas"] = paginas
    r["pdf_producer"] = (meta.get("producer") or "").strip()
    r["pdf_creation"] = (meta.get("creationDate") or "").strip()

    # ---- capa ----
    capa_linhas = [t for p, t in linhas if p == 1]
    capa = OrderedDict()
    for t in capa_linhas:
        m = RE_CAPA_CAMPO.match(t)
        if m:
            capa[m.group(1).strip()] = m.group(2).strip()
    r["capa_campos"] = capa
    r["capa_linhas"] = capa_linhas
    r["titulo_capa"] = next((t for t in capa_linhas if "RELAT" in t.upper() and ":" not in t), None)
    m = RE_PERIODO.search("\n".join(capa_linhas))
    r["capa_periodo_inicio"] = m.group(1) if m else None
    r["capa_periodo_fim"] = m.group(2) if m else None

    # ---- indice / secoes ----
    toc = parse_toc(linhas)
    r["indice"] = toc
    r["secoes_n1"] = [e["titulo"] for e in toc if e["nivel"] == 1]
    r["secoes_n2"] = ["%s %s" % (e["num"], e["titulo"]) for e in toc if e["nivel"] == 2]
    r["subsecoes_52x"] = ["%s %s" % (e["num"], e["titulo"]) for e in toc
                          if e["num"] and e["num"].startswith("5.2.")]
    r["assinatura_estrutura"] = " | ".join(
        "%s %s" % (e["num"] or "?", norm(e["titulo"])) for e in toc if e["nivel"] in (1, 2))
    r["assinatura_n1"] = " | ".join(norm(t) for t in r["secoes_n1"])

    secoes = fatiar_secoes(linhas, toc)
    r["secoes_encontradas"] = list(secoes.keys())

    # ---- anexos ----
    anexos = []
    kA, vA = achar_secao(secoes, r"^anexos?$", r"\banexos\b")
    fonte_anexos = vA["texto"] if vA else texto_a4
    for ln in fonte_anexos.split("\n"):
        m = RE_ANEXO.match(ln)
        if m:
            tit = m.group(2).strip()
            anexos.append({"ordem": m.group(1), "titulo": tit, "categoria": categoria_anexo(tit)})
    r["anexos_estilo"] = "LISTA_ANEXO_ROMANO" if anexos else None
    if not anexos and vA:
        # Modelo intermediario (~RO007 a RO030): a secao ANEXOS nao traz a lista
        # "Anexo I - ...". Cada anexo e uma folha separadora A4 com o titulo.
        anexos = anexos_por_folha_separadora(linhas, vA["pagina_real"])
        if anexos:
            r["anexos_estilo"] = "FOLHAS_SEPARADORAS"
    vistos, uniq = set(), []
    for a in anexos:
        k = (a["ordem"], norm(a["titulo"]))
        if k not in vistos:
            vistos.add(k)
            uniq.append(a)
    r["secao_anexos"] = kA
    r["anexos"] = uniq
    r["categorias_anexo"] = [a["categoria"] for a in uniq]

    # ---- legendas ----
    leg = []
    for p, t in linhas:
        m = RE_LEGENDA.match(t)
        if m:
            leg.append({"tipo": m.group(1).capitalize().replace("Grafico", "Gráfico"),
                        "n": int(m.group(2)), "texto": limpar(m.group(3)), "pagina": p})
    r["legendas"] = leg
    r["legendas_por_tipo"] = Counter(x["tipo"] for x in leg).most_common()

    # ---- termo de encerramento / data de emissao ----
    kT, vT = achar_secao(secoes, r"termo de encerramento")
    r["secao_termo"] = kT
    termo_txt = (vT["texto"] if vT else "")
    if not termo_txt:
        # fallback: o Termo pode nao ter virado secao no indice; procura no texto A4
        i = norm(texto_a4).find("termo de encerramento")
        if i >= 0:
            termo_txt = texto_a4[max(0, i - 40): i + 1200]
    emissao, local = None, None
    for m in RE_DATA_EXTENSO.finditer(termo_txt):
        mes = MESES.get(norm(m.group(3)))
        if mes:
            emissao = "%s-%02d-%02d" % (m.group(4), mes, int(m.group(2)))
            local = limpar(m.group(1)).split("\n")[-1].strip()
    r["termo_texto"] = termo_txt
    r["termo_datas"] = ["%s-%s-%s" % (m.group(3), m.group(2), m.group(1))
                        for m in RE_DATA.finditer(termo_txt)]
    r["data_emissao_termo"] = emissao
    r["local_emissao"] = local
    r["data_emissao_mtime"] = (reg["modificado"] or "")[:10]

    # ---- secoes-chave para casos raros ----
    kNC_orgao, vNC_orgao = achar_secao(secoes, r"orgao")
    kNC_prog, vNC_prog = achar_secao(secoes, r"emitidas pela programa")
    kNC_hist, vNC_hist = achar_secao(secoes, r"historico de nao conformidades")
    kRisco, vRisco = achar_secao(secoes, r"interferencias? em campo", r"relato de risco")
    kClima, vClima = achar_secao(secoes, r"condicoes climaticas")
    kApres, vApres = achar_secao(secoes, r"^apresentacao$")

    def pack(k, v):
        if not v:
            return None
        t = v["texto"].strip()
        return {"secao": k, "num": v["num"], "titulo": v["titulo"],
                "pagina": v["pagina_real"], "texto": t,
                "negativo": bool(RE_NEGATIVO.search(norm(t))) if t else None,
                "n_chars": len(t)}

    r["sec_nc_orgao"] = pack(kNC_orgao, vNC_orgao)
    r["sec_nc_prog"] = pack(kNC_prog, vNC_prog)
    r["sec_nc_historico"] = pack(kNC_hist, vNC_hist)
    r["sec_riscos"] = pack(kRisco, vRisco)
    r["sec_clima"] = pack(kClima, vClima)
    r["sec_apresentacao"] = pack(kApres, vApres)

    # ---- (D) varredura de termos de excecao ----
    n_texto = norm(texto_a4)
    achados = {}
    for termo, pad in TERMOS_EXCECAO.items():
        hits = []
        for m in re.finditer(pad, n_texto):
            ini = max(0, m.start() - 170)
            fim = min(len(n_texto), m.end() + 170)
            hits.append(n_texto[ini:fim])
            if len(hits) >= 4:
                break
        if hits:
            achados[termo] = {"n": len(list(re.finditer(pad, n_texto))), "trechos": hits}
    r["termos_excecao"] = achados

    # ---- cache de texto ----
    os.makedirs(U.assert_read_only(DIR_TXT), exist_ok=True)
    nome_txt = "%s_RO%03d%s.txt" % (obra, reg["numero"],
                                    (".%d" % reg["revisao"]) if reg["revisao"] else "")
    if reg.get("subpasta"):
        nome_txt = nome_txt[:-4] + "_OBSOLETO.txt"
    cam_txt = os.path.join(DIR_TXT, nome_txt)
    with open(cam_txt, "w", encoding="utf-8") as fh:
        fh.write(texto_a4)
    r["texto_cache"] = os.path.relpath(cam_txt, U.PROJ)

    # paragrafos por secao (para a analise textual)
    r["paragrafos"] = {k: v["texto"] for k, v in secoes.items()}
    return r


# ---------------------------------------------------------------- varredura --
def chave_arquivo(reg):
    return "%s|%s|%s" % (reg["caminho"], reg["tamanho_bytes"], reg["modificado"])


def varrer(args):
    inv = json.load(open(INV, encoding="utf-8"))
    tarefas = []
    for obra in U.OBRAS:
        if args.obra and obra != args.obra:
            continue
        ents = [e for e in inv["obras"][obra]["entradas"] if e["tipo"] == "pasta"]
        ents.sort(key=lambda e: (e.get("numero") or 0))
        for e in ents:
            principal, obsoletos, outros = candidatos(e)
            if principal:
                tarefas.append((obra, e, principal))
            for o in obsoletos + outros:
                tarefas.append((obra, e, o))

    feitos = {}
    if os.path.exists(PARCIAL) and not args.forcar:
        with open(PARCIAL, encoding="utf-8") as fh:
            for ln in fh:
                ln = ln.strip()
                if not ln:
                    continue
                try:
                    d = json.loads(ln)
                except Exception:
                    continue
                feitos[d["_chave"]] = d

    if args.limite:
        tarefas = tarefas[:args.limite]

    pendentes = [t for t in tarefas if chave_arquivo(t[2]) not in feitos]
    print("Tarefas: %d | ja feitas: %d | pendentes: %d"
          % (len(tarefas), len(tarefas) - len(pendentes), len(pendentes)))

    os.makedirs(U.assert_read_only(DIR_INV), exist_ok=True)
    modo = "w" if args.forcar else "a"
    t0 = time.time()
    with open(PARCIAL, modo, encoding="utf-8") as fh:
        for i, (obra, e, reg) in enumerate(pendentes, 1):
            try:
                r = analisar(reg, obra, e)
                r["_chave"] = chave_arquivo(reg)
                r["_erro"] = None
            except Exception as ex:
                r = {"_chave": chave_arquivo(reg), "obra": obra, "pasta": e["nome"],
                     "arquivo": reg["nome"], "caminho": reg["caminho"],
                     "numero": reg["numero"], "revisao": reg["revisao"],
                     "_erro": "%s: %s" % (type(ex).__name__, ex)}
                print("  ERRO %s %s -> %s" % (obra, reg["nome"], r["_erro"]))
            fh.write(json.dumps(r, ensure_ascii=False) + "\n")
            fh.flush()
            feitos[r["_chave"]] = r
            if r.get("_erro") is None:
                print("  [%3d/%3d] %-7s RO%03d%s %2s pgs %3d imgs %2d anexos %2d secoes (%.1fs)"
                      % (i, len(pendentes), obra, reg["numero"],
                         (".%d" % reg["revisao"]) if reg["revisao"] else "  ",
                         r["n_paginas"], r["n_imagens"], len(r["anexos"]),
                         len(r["indice"]), time.time() - t0))
                t0 = time.time()

    regs = [feitos[chave_arquivo(t[2])] for t in tarefas if chave_arquivo(t[2]) in feitos]
    return regs


# --------------------------------------------------------- pos-processamento -
def gravar_indice(regs):
    principais = [r for r in regs if not r.get("subpasta") and not r.get("_erro")]
    principais.sort(key=lambda r: (r["obra"], r["numero"], r.get("revisao") or 0))

    def enxuto(r):
        d = OrderedDict()
        for k in ("obra", "pasta", "pasta_data_inicio", "pasta_data_fim", "numero", "revisao",
                  "estilo_nome", "arquivo", "caminho", "capa_periodo_inicio", "capa_periodo_fim",
                  "data_emissao_termo", "data_emissao_mtime", "n_paginas", "n_imagens",
                  "tamanho_bytes", "formatos", "capa_campos", "titulo_capa",
                  "secoes_n1", "secoes_n2", "subsecoes_52x", "assinatura_n1",
                  "assinatura_estrutura", "anexos", "categorias_anexo", "anexos_estilo",
                  "legendas_por_tipo", "secao_anexos", "local_emissao",
                  "termo_texto", "pdf_producer", "texto_cache"):
            if k in r:
                d[k] = r[k]
        d["legendas"] = r.get("legendas", [])
        d["indice"] = r.get("indice", [])
        return d

    saida = {
        "gerado_em": datetime.datetime.now().isoformat(timespec="seconds"),
        "gerado_por": "scripts/13_corpus_historico.py",
        "n_relatorios": len(principais),
        "relatorios": [enxuto(r) for r in principais],
        "revisoes_obsoletas": [enxuto(r) for r in regs
                               if r.get("subpasta") and not r.get("_erro")],
        "erros": [{"obra": r["obra"], "arquivo": r["arquivo"], "erro": r["_erro"]}
                  for r in regs if r.get("_erro")],
    }
    cam = os.path.join(DIR_INV, "relatorios_semanais_historico.json")
    with open(U.assert_read_only(cam), "w", encoding="utf-8") as fh:
        json.dump(saida, fh, ensure_ascii=False, indent=1)

    camc = os.path.join(DIR_INV, "relatorios_semanais_historico.csv")
    cols = ["obra", "numero", "revisao", "pasta", "pasta_data_inicio", "pasta_data_fim",
            "capa_periodo_inicio", "capa_periodo_fim", "data_emissao_termo",
            "data_emissao_mtime", "n_paginas", "n_imagens", "tamanho_bytes",
            "n_secoes_n1", "n_secoes_n2", "n_subsecoes_52x", "n_anexos", "n_legendas",
            "formatos", "secoes_n1", "categorias_anexo", "arquivo"]
    with open(U.assert_read_only(camc), "w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(cols)
        for r in principais:
            w.writerow([
                r["obra"], r["numero"], r.get("revisao") or "", r["pasta"],
                r.get("pasta_data_inicio") or "", r.get("pasta_data_fim") or "",
                r.get("capa_periodo_inicio") or "", r.get("capa_periodo_fim") or "",
                r.get("data_emissao_termo") or "", r.get("data_emissao_mtime") or "",
                r["n_paginas"], r["n_imagens"], r["tamanho_bytes"],
                len(r["secoes_n1"]), len(r["secoes_n2"]), len(r["subsecoes_52x"]),
                len(r["anexos"]), len(r["legendas"]),
                ", ".join("%s(%d)" % (f, n) for f, n in r["formatos"]),
                " > ".join(r["secoes_n1"]),
                ", ".join(r["categorias_anexo"]),
                r["arquivo"]])
    print("OK ->", cam)
    print("OK ->", camc)
    return principais


# --------------------------------------------------------------- casos raros -
def data_ref(r):
    return r.get("pasta_data_inicio") or r.get("capa_periodo_inicio") or ""


def rotulo(r):
    return "%s RO%03d%s (%s a %s)" % (
        r["obra"], r["numero"], (".%d" % r["revisao"]) if r.get("revisao") else "",
        r.get("pasta_data_inicio") or "?", r.get("pasta_data_fim") or "?")


def casos_raros(regs, principais):
    C = OrderedDict()
    C["gerado_em"] = datetime.datetime.now().isoformat(timespec="seconds")
    C["universo"] = {"n_relatorios_principais": len(principais),
                     "n_revisoes_obsoletas": len([r for r in regs if r.get("subpasta")]),
                     "n_erros": len([r for r in regs if r.get("_erro")])}

    # ---------- A: NC emitida pela Órgão Contratante ----------
    A = {"criterio": "secao cujo titulo contem 'Órgão Contratante' (8.2 no modelo atual, 9.2 no antigo); "
                     "POSITIVO = texto sem marcador de ausencia",
         "positivos": [], "negativos": 0, "ausentes": [], "todos_textos_distintos": []}
    vistos = Counter()
    for r in principais:
        s = r.get("sec_nc_orgao")
        if not s:
            A["ausentes"].append(rotulo(r))
            continue
        vistos[norm(s["texto"])[:400]] += 1
        if s["negativo"] is False and s["n_chars"] > 0:
            A["positivos"].append({"relatorio": rotulo(r), "obra": r["obra"],
                                   "numero": r["numero"], "revisao": r.get("revisao"),
                                   "periodo": [r.get("pasta_data_inicio"), r.get("pasta_data_fim")],
                                   "secao": s["secao"], "pagina": s["pagina"],
                                   "texto": s["texto"], "arquivo": r["arquivo"]})
        else:
            A["negativos"] += 1
    A["variantes_de_texto"] = [{"n_ocorrencias": n, "amostra": t} for t, n in vistos.most_common()]
    # A2 — NC/NI cuja ORIGEM e a Órgão Contratante aparecem na tabela do historico (8.3 / 9.3),
    # ainda que a subsecao 8.2 "no periodo" nunca registre emissao.
    # Colunas da tabela 8.3/9.3, nesta ordem:
    #   Documento | Rodovia | Empresa | Irregularidade | Status | Criticidade |
    #   Natureza | Data Abertura | Data Encerramento | EMISSOR
    # No texto achatado, o emissor de uma linha aparece DEPOIS dela e ANTES do
    # documento da linha seguinte. Por isso o emissor e o primeiro token
    # Programa|Órgão Contratante|Parceiro que vem APOS o codigo do documento.
    # Conservador: so conta quando o emissor vem logo apos a(s) data(s) da linha
    # (fim de linha da tabela) e e seguido pelo codigo do documento da linha seguinte.
    RE_FIM_LINHA = re.compile(
        r"(\d{2}/\d{2}/\d{4})\s+(?:(\d{2}/\d{2}/\d{4})\s+)?(programa|orgao|parceiro)\s+"
        r"((?:nc|ni|rt|rm|rit)[\s\-]?0*\d{1,3}(?:[\-/]\d{1,4})?)\b")
    achados = OrderedDict()
    for r in principais:
        s = r.get("sec_nc_historico") or {}
        alvo = norm(s.get("texto", ""))
        if not alvo or alvo.find("documento rodovia empresa") < 0:
            continue                      # so trata a tabela, nao o texto corrido
        alvo = alvo[alvo.find("documento rodovia empresa"):]
        for m in RE_FIM_LINHA.finditer(alvo):
            if m.group(3) != "orgao":
                continue
            # o codigo APOS o emissor e o da linha seguinte; a linha cujo emissor
            # e a Órgão Contratante e a anterior -> pegamos o codigo anterior mais proximo.
            ant = list(re.finditer(
                r"\b((?:nc|ni|rt|rm|rit)[\s\-]?0*\d{1,3}(?:[\-/]\d{1,4})?)\b",
                alvo[:m.start()]))
            cod = re.sub(r"\s+", "-", ant[-1].group(1)).upper() if ant else "(nao identificado)"
            k = (r["obra"], cod)
            if k in achados:
                achados[k]["visto_em"].append(rotulo(r))
                continue
            achados[k] = {
                "obra": r["obra"], "codigo_inferido": cod, "emissor": "Órgão Contratante",
                "confianca": "FORTE INFERENCIA (tabela achatada pela extracao de PDF)",
                "primeira_ocorrencia": rotulo(r),
                "periodo": [r.get("pasta_data_inicio"), r.get("pasta_data_fim")],
                "trecho": alvo[max(0, ant[-1].start() if ant else m.start()):m.end()][:900],
                "visto_em": [rotulo(r)]}
    A["A2_nc_de_origem_orgao_no_historico"] = {
        "observacao": "A subsecao 8.2/9.2 nunca registra emissao no periodo, mas a tabela "
                      "de historico (8.3/9.3) tem coluna EMISSOR e nela aparecem "
                      "documentos emitidos pela Órgão Contratante.",
        "n": len(achados),
        "itens": [dict(v, n_relatorios=len(v["visto_em"]), visto_em=v["visto_em"][:8])
                  for v in achados.values()]}
    C["A_nc_orgao"] = A

    # ---------- B: riscos e interferencias ----------
    B = {"criterio": "secao 'RELATO DE (RISCOS E) INTERFERENCIAS EM CAMPO' (9 atual / 10 antigo)",
         "positivos": [], "negativos": 0, "ausentes": []}
    vistosB = Counter()
    for r in principais:
        s = r.get("sec_riscos")
        if not s:
            B["ausentes"].append(rotulo(r))
            continue
        vistosB[norm(s["texto"])[:400]] += 1
        if s["negativo"] is False and s["n_chars"] > 0:
            B["positivos"].append({"relatorio": rotulo(r), "obra": r["obra"],
                                   "numero": r["numero"],
                                   "periodo": [r.get("pasta_data_inicio"), r.get("pasta_data_fim")],
                                   "secao": s["secao"], "pagina": s["pagina"],
                                   "texto": s["texto"], "arquivo": r["arquivo"]})
        else:
            B["negativos"] += 1
    B["variantes_de_texto"] = [{"n_ocorrencias": n, "amostra": t} for t, n in vistosB.most_common()]
    C["B_riscos_interferencias"] = B

    # ---------- C: revisoes ----------
    porchave = {}
    for r in regs:
        if r.get("_erro"):
            continue
        porchave.setdefault((r["obra"], r["numero"]), []).append(r)
    Cc = {"criterio": "arquivos RO_NNN.rev com rev >= 2, em qualquer pasta",
          "ocorrencias": []}
    for (obra, numero), lst in sorted(porchave.items()):
        revs = [r for r in lst if (r.get("revisao") or 1) >= 2]
        if not revs:
            continue
        atual = max(lst, key=lambda r: (r.get("revisao") or 0))
        antigas = [r for r in lst if r is not atual]
        oc = {"obra": obra, "numero": numero,
              "pasta": atual["pasta"],
              "periodo": [atual.get("pasta_data_inicio"), atual.get("pasta_data_fim")],
              "versoes": []}
        for r in sorted(lst, key=lambda x: (x.get("revisao") or 0)):
            oc["versoes"].append({
                "arquivo": r["arquivo"], "revisao": r.get("revisao"),
                "subpasta": r.get("subpasta", "") or "(raiz da pasta-semana)",
                "mtime": r.get("mtime"), "n_paginas": r.get("n_paginas"),
                "n_imagens": r.get("n_imagens"), "tamanho_bytes": r.get("tamanho_bytes"),
                "data_emissao_termo": r.get("data_emissao_termo"),
                "n_anexos": len(r.get("anexos", [])),
                "assinatura_n1": r.get("assinatura_n1"),
            })
        # diff textual entre a menor e a maior revisao
        if antigas:
            velha = min(antigas, key=lambda x: (x.get("revisao") or 0))
            oc["diff"] = diff_relatorios(velha, atual)
        Cc["ocorrencias"].append(oc)
    C["C_revisoes"] = Cc

    # ---------- D: excecoes ----------
    D = {"por_termo": OrderedDict(), "anomalias": OrderedDict()}
    for termo in TERMOS_EXCECAO:
        ocs = []
        for r in principais:
            t = r.get("termos_excecao", {}).get(termo)
            if t:
                ocs.append({"relatorio": rotulo(r), "obra": r["obra"], "numero": r["numero"],
                            "periodo": [r.get("pasta_data_inicio"), r.get("pasta_data_fim")],
                            "n_ocorrencias": t["n"], "trechos": t["trechos"]})
        D["por_termo"][termo] = {"n_relatorios": len(ocs),
                                 "pct": round(100.0 * len(ocs) / max(1, len(principais)), 1),
                                 "ocorrencias": ocs if len(ocs) <= 25 else ocs[:25],
                                 "truncado": len(ocs) > 25}

    # D2 — anomalias estatisticas
    pags = [r["n_paginas"] for r in principais]
    media = sum(pags) / len(pags)
    dp = (sum((x - media) ** 2 for x in pags) / len(pags)) ** 0.5
    D["anomalias"]["paginas"] = {
        "media": round(media, 1), "desvio_padrao": round(dp, 1),
        "min": min(pags), "max": max(pags),
        "fora_2dp": [{"relatorio": rotulo(r), "n_paginas": r["n_paginas"],
                      "z": round((r["n_paginas"] - media) / dp, 2)}
                     for r in principais if abs(r["n_paginas"] - media) > 2 * dp]}

    # secoes raras (titulo nivel 1 que aparece em poucos relatorios)
    cnt_sec = Counter()
    onde_sec = defaultdict(list)
    for r in principais:
        for t in set(norm(x) for x in r["secoes_n1"]):
            cnt_sec[t] += 1
            onde_sec[t].append(rotulo(r))
    D["anomalias"]["secoes_n1_raras"] = [
        {"titulo": t, "n_relatorios": n, "onde": onde_sec[t][:12]}
        for t, n in sorted(cnt_sec.items(), key=lambda x: x[1]) if n <= 6]
    D["anomalias"]["secoes_n1_frequencia"] = [
        {"titulo": t, "n_relatorios": n} for t, n in cnt_sec.most_common()]

    # anexos raros
    cnt_ax = Counter()
    onde_ax = defaultdict(list)
    for r in principais:
        for c in set(r["categorias_anexo"]):
            cnt_ax[c] += 1
            onde_ax[c].append(rotulo(r))
    D["anomalias"]["anexos_raros"] = [
        {"categoria": c, "n_relatorios": n, "onde": onde_ax[c][:12]}
        for c, n in sorted(cnt_ax.items(), key=lambda x: x[1]) if n <= 6]
    D["anomalias"]["anexos_frequencia"] = [
        {"categoria": c, "n_relatorios": n} for c, n in cnt_ax.most_common()]

    # formatos de papel raros
    cnt_fmt = Counter()
    onde_fmt = defaultdict(list)
    for r in principais:
        for f, _ in r["formatos"]:
            cnt_fmt[f] += 1
            onde_fmt[f].append(rotulo(r))
    D["anomalias"]["formatos_raros"] = [
        {"formato": f, "n_relatorios": n, "onde": onde_fmt[f][:12]}
        for f, n in sorted(cnt_fmt.items(), key=lambda x: x[1]) if n <= 8]
    D["anomalias"]["formatos_frequencia"] = [
        {"formato": f, "n_relatorios": n} for f, n in cnt_fmt.most_common()]

    # semanas sem PDF principal
    C["D_excecoes"] = D
    C["marcos"] = detectar_marcos(principais)
    C["cobertura"] = cobertura(principais)

    cam = os.path.join(DIR_CMP, "casos_raros.json")
    os.makedirs(U.assert_read_only(DIR_CMP), exist_ok=True)
    with open(U.assert_read_only(cam), "w", encoding="utf-8") as fh:
        json.dump(C, fh, ensure_ascii=False, indent=1)
    print("OK ->", cam)
    return C


def diff_relatorios(a, b):
    """Diff linha a linha do texto A4 entre duas versoes."""
    def linhas(r):
        cam = os.path.join(U.PROJ, r["texto_cache"])
        if not os.path.exists(cam):
            return []
        out = []
        for l in open(cam, encoding="utf-8").read().split("\n"):
            l = l.strip()
            if l and not re.fullmatch(r"\d{1,3}", l):
                out.append(re.sub(r"\s+", " ", l))
        return out
    la, lb = linhas(a), linhas(b)
    if not la or not lb:
        return {"erro": "texto indisponivel"}
    sm = difflib.SequenceMatcher(None, la, lb, autojunk=False)
    mud = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        mud.append({"tipo": tag,
                    "removido": la[i1:i2][:25],
                    "adicionado": lb[j1:j2][:25]})
    return {"similaridade": round(sm.ratio(), 4),
            "n_linhas_a": len(la), "n_linhas_b": len(lb),
            "blocos_alterados": mud[:60],
            "truncado": len(mud) > 60}


def cobertura(principais):
    """Semanas do inventario sem PDF principal (relatorio em elaboracao/faltante)."""
    inv = json.load(open(INV, encoding="utf-8"))
    tem = {(r["obra"], r["numero"]) for r in principais}
    falt = []
    for obra in U.OBRAS:
        for e in inv["obras"][obra]["entradas"]:
            if e["tipo"] != "pasta":
                continue
            n = e.get("numero")
            if n is None or (obra, n) not in tem:
                falt.append({"obra": obra, "pasta": e["nome"], "numero": n,
                             "data_inicio": e.get("data_inicio"),
                             "n_arquivos": len(e.get("arquivos", []))})
    return {"pastas_sem_pdf_principal": falt}


# --------------------------------------------------------------- marcos ------
def detectar_marcos(principais):
    """Compara relatorio N com N-1 dentro da mesma obra e registra o que mudou."""
    marcos = []
    for obra in U.OBRAS:
        rs = sorted([r for r in principais if r["obra"] == obra], key=lambda r: r["numero"])
        for a, b in zip(rs, rs[1:]):
            mud = {}
            # secoes nivel 1
            sa = [norm(x) for x in a["secoes_n1"]]
            sb = [norm(x) for x in b["secoes_n1"]]
            if sa != sb:
                mud["secoes_n1"] = {"removidas": [x for x in sa if x not in sb],
                                    "adicionadas": [x for x in sb if x not in sa],
                                    "antes": sa, "depois": sb}
            # secoes nivel 2
            s2a = [norm(x) for x in a["secoes_n2"]]
            s2b = [norm(x) for x in b["secoes_n2"]]
            if s2a != s2b:
                mud["secoes_n2"] = {"removidas": [x for x in s2a if x not in s2b],
                                    "adicionadas": [x for x in s2b if x not in s2a]}
            # anexos
            ca, cb = set(a["categorias_anexo"]), set(b["categorias_anexo"])
            if ca != cb:
                mud["anexos"] = {"removidos": sorted(ca - cb), "adicionados": sorted(cb - ca)}
            # formatos
            fa = set(f for f, _ in a["formatos"])
            fb = set(f for f, _ in b["formatos"])
            if fa != fb:
                mud["formatos"] = {"removidos": sorted(fa - fb), "adicionados": sorted(fb - fa)}
            # nomenclatura do arquivo
            if a["estilo_nome"] != b["estilo_nome"]:
                mud["nomenclatura_arquivo"] = {"antes": a["arquivo"], "depois": b["arquivo"]}
            if a.get("anexos_estilo") != b.get("anexos_estilo"):
                mud["estilo_lista_anexos"] = {"antes": a.get("anexos_estilo"),
                                              "depois": b.get("anexos_estilo")}
            # titulos de anexo (nao so categoria)
            ta = set(norm(x["titulo"]) for x in a.get("anexos", []))
            tb = set(norm(x["titulo"]) for x in b.get("anexos", []))
            if ta != tb:
                mud["anexos_titulos"] = {"removidos": sorted(ta - tb)[:12],
                                         "adicionados": sorted(tb - ta)[:12]}
            # dia da semana do periodo
            def dow(r):
                d = r.get("pasta_data_inicio")
                if not d:
                    return None
                try:
                    return datetime.date.fromisoformat(d).strftime("%A")
                except Exception:
                    return None
            if dow(a) and dow(b) and dow(a) != dow(b):
                mud["dia_semana_inicio"] = {"antes": dow(a), "depois": dow(b),
                                            "data_antes": a.get("pasta_data_inicio"),
                                            "data_depois": b.get("pasta_data_inicio")}
            # capa
            cka, ckb = set(a["capa_campos"]), set(b["capa_campos"])
            if cka != ckb:
                mud["capa_campos"] = {"removidos": sorted(cka - ckb),
                                      "adicionados": sorted(ckb - cka)}
            if mud:
                marcos.append({"obra": obra,
                               "de": "RO%03d (%s)" % (a["numero"], a.get("pasta_data_inicio")),
                               "para": "RO%03d (%s)" % (b["numero"], b.get("pasta_data_inicio")),
                               "data_inicio_novo": b.get("pasta_data_inicio"),
                               "mudancas": mud})
    # agrupa por data para ver se a mudanca foi nas 4 obras
    porData = defaultdict(list)
    for m in marcos:
        porData[m["data_inicio_novo"]].append(m)
    consolidado = []
    for d in sorted(porData, key=lambda x: x or ""):
        ms = porData[d]
        tipos = sorted({k for m in ms for k in m["mudancas"]})
        consolidado.append({"data": d, "obras": sorted({m["obra"] for m in ms}),
                            "n_obras": len({m["obra"] for m in ms}),
                            "tipos_de_mudanca": tipos, "detalhe": ms})
    return {"transicoes": marcos, "consolidado_por_data": consolidado}


# ---------------------------------------------------- analise textual --------
RE_VARS = [
    (r"\b\d{2}/\d{2}/\d{4}\b", "{data}"),
    (r"\bRD[\s\-]?\d{3}[A-Z]?\b", "{rodovia}"),
    (r"\b\d{1,3}(?:\.\d{3})*,\d+\s*%", "{percentual}"),
    (r"\b\d+,\d+\s*km\b", "{extensao}"),
    (r"\b\d{1,3}(?:\.\d{3})*,\d+\b", "{numero}"),
    (r"\bn[ºo°]\s*\d+[\/\-]?\d*\b", "{numero_doc}"),
    (r"\b\d{2,4}[\/\-]\d{2,4}\b", "{codigo}"),
    (r"\b\d+\b", "{n}"),
]


def parametrizar(frase):
    s = frase
    for pad, sub in RE_VARS:
        s = re.sub(pad, sub, s, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", s).strip()


def analise_textual(principais):
    """Clustering de paragrafos por secao: FIXO / TEMPLATE / ESTRUTURADO / LIVRE."""
    por_secao = defaultdict(list)          # titulo_norm -> [(rot, paragrafo)]
    for r in principais:
        for k, txt in (r.get("paragrafos") or {}).items():
            m = RE_NUM_TIT.match(k) or re.match(r"^(\S+)\s+(.*)$", k)
            tit = norm(m.group(2)) if m else norm(k)
            for par in re.split(r"\n(?=[A-ZÀ-Ú0-9])", txt):
                par = re.sub(r"\s+", " ", par).strip()
                if len(par) < 25:
                    continue
                por_secao[tit].append((rotulo(r), par))

    total_rel = len(principais)
    saida = OrderedDict()
    saida["gerado_em"] = datetime.datetime.now().isoformat(timespec="seconds")
    saida["n_relatorios"] = total_rel
    saida["secoes"] = OrderedDict()

    for tit, itens in sorted(por_secao.items(), key=lambda x: -len(x[1])):
        rels_secao = len({r for r, _ in itens})
        grupos = defaultdict(list)
        for rot, par in itens:
            grupos[parametrizar(par)].append((rot, par))
        linhas = []
        for tpl, ocs in sorted(grupos.items(), key=lambda x: -len(x[1])):
            rels = {r for r, _ in ocs}
            brutos = {p for _, p in ocs}
            if len(rels) < 3:
                classe = "LIVRE"
            elif len(brutos) == 1:
                classe = "FIXO"
            elif "{" in tpl:
                classe = "TEMPLATE"
            else:
                classe = "ESTRUTURADO"
            linhas.append({
                "classe": classe,
                "n_relatorios": len(rels),
                "cobertura_pct": round(100.0 * len(rels) / max(1, rels_secao), 1),
                "n_variantes_brutas": len(brutos),
                "template": tpl[:900],
                "variaveis": sorted(set(re.findall(r"\{[a-z_]+\}", tpl))),
                "exemplo": sorted(brutos)[0][:900],
                "obras": sorted({r.split()[0] for r in rels}),
            })
        saida["secoes"][tit] = {
            "n_relatorios_com_a_secao": rels_secao,
            "n_paragrafos": len(itens),
            "resumo_classes": Counter(l["classe"] for l in linhas).most_common(),
            "grupos": linhas[:60],
        }

    # resumo global: quanto do corpus e fixo/template/livre
    glob_cls = Counter()
    glob_par = Counter()
    for tit, s in saida["secoes"].items():
        for g in s["grupos"]:
            glob_cls[g["classe"]] += 1
            glob_par[g["classe"]] += g["n_relatorios"]
    tot_g = sum(glob_cls.values()) or 1
    tot_p = sum(glob_par.values()) or 1
    saida["resumo_global"] = {
        "grupos_por_classe": glob_cls.most_common(),
        "pct_grupos": {k: round(100.0 * v / tot_g, 1) for k, v in glob_cls.items()},
        "ocorrencias_por_classe": glob_par.most_common(),
        "pct_ocorrencias": {k: round(100.0 * v / tot_p, 1) for k, v in glob_par.items()},
        "legenda": {
            "FIXO": "mesmo texto bruto em >= 3 relatorios - copiar literalmente",
            "TEMPLATE": "mesma forma, variaveis trocadas - parametrizar",
            "ESTRUTURADO": "mesma forma sem variavel detectada (rotulo/cabecalho de tabela)",
            "LIVRE": "aparece em < 3 relatorios - redacao original, exige IA ou humano"},
    }

    cam = os.path.join(DIR_CMP, "corpus_textual.json")
    with open(U.assert_read_only(cam), "w", encoding="utf-8") as fh:
        json.dump(saida, fh, ensure_ascii=False, indent=1)
    print("OK ->", cam)
    return saida


# ------------------------------------------------------------------- main ----
def main():
    ap = argparse.ArgumentParser(description="Varredura historica do corpus de Relatorios Semanais")
    ap.add_argument("--obra", help="limitar a uma obra (RD101|RD102A|RD103|RD104)")
    ap.add_argument("--limite", type=int, help="processar no maximo N PDFs (teste)")
    ap.add_argument("--forcar", action="store_true", help="ignorar o progresso incremental")
    ap.add_argument("--so-analise", action="store_true",
                    help="nao abrir PDFs; usar apenas o progresso ja salvo")
    args = ap.parse_args()

    if args.so_analise:
        regs = []
        with open(PARCIAL, encoding="utf-8") as fh:
            for ln in fh:
                if ln.strip():
                    regs.append(json.loads(ln))
        # deduplica por chave, mantendo o ultimo
        dd = OrderedDict((r["_chave"], r) for r in regs)
        regs = list(dd.values())
    else:
        regs = varrer(args)

    principais = gravar_indice(regs)
    casos_raros(regs, principais)
    analise_textual(principais)
    print("\nConcluido: %d relatorios principais." % len(principais))


if __name__ == "__main__":
    main()
