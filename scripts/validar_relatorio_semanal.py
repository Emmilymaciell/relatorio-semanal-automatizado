# -*- coding: utf-8 -*-
"""
validar_relatorio_semanal.py

VALIDADOR — le um Relatorio Semanal ja existente (PDF) e confere as regras
descobertas na engenharia reversa. NAO GERA NADA. Nao escreve no SharePoint.

E deliberadamente a primeira peca de software do projeto: risco zero, e serve
para PROVAR se o mapeamento esta correto antes de qualquer gerador.

Validacao VERSIONADA: a versao do formato e resolvida pela data do periodo
(config/report_versions.yaml), de modo que um relatorio de 2025 nao falha por
seguir corretamente o formato de 2025.

Uso:
    python scripts/validar_relatorio_semanal.py                 # amostra (8)
    python scripts/validar_relatorio_semanal.py --todos         # corpus inteiro
    python scripts/validar_relatorio_semanal.py --obra RD101
    python scripts/validar_relatorio_semanal.py --md docs/validacao/VALIDACAO_BASELINE_V2.md
"""
import os
import sys
import re
import json
import argparse
import datetime
from collections import Counter, OrderedDict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

import fitz
import yaml

PASS, FAIL, WARN, SKIP, NA = "PASS", "FAIL", "WARNING", "SKIP", "N/A"

FORMATOS = [
    ("A4 retrato", 595, 842), ("A4 paisagem", 842, 595),
    ("A3 retrato", 842, 1191), ("A3 paisagem", 1191, 842),
    ("A2 retrato", 1191, 1684), ("A2 paisagem", 1684, 1191),
    ("A1 paisagem", 2384, 1684),
    ("A0 retrato", 2384, 3370), ("A0 paisagem", 3370, 2384),
]

RE_RO_FILE = re.compile(r"^RO[_\s-]*0*(\d{1,3})(?:\.(\d))?\s*-", re.IGNORECASE)
RE_ANEXO = re.compile(r"^\s*Anexo\s+([IVXLC]+)\s*[\-–—]\s*(.+?)\s*$",
                      re.IGNORECASE | re.MULTILINE)
RE_SEC_ANEXOS = re.compile(r"^\s*1[12]\.?\s*ANEXOS\s*$", re.IGNORECASE | re.MULTILINE)
RE_TERMO = re.compile(
    r"constitu[ií]do\s+de\s+(\d+)\s*\(([^)]+)\)\s*p[aá]ginas?.*?"
    r"([A-Za-zÀ-ú]+),\s*(\d{1,2})\s+de\s+([a-zç]+)\s+de\s+(\d{4})",
    re.IGNORECASE | re.DOTALL)
RE_PERIODO_CAPA = re.compile(r"PER[IÍ]ODO:\s*(\d{2}/\d{2}/\d{4})\s*a\s*(\d{2}/\d{2}/\d{4})",
                             re.IGNORECASE)
RE_NUM_CAPA = re.compile(r"Relat[óo]rio:\s*0*(\d{1,3})", re.IGNORECASE)
RE_TAB_CLIMA = re.compile(r"\b(\d{1,2}/(?:jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez))\b",
                          re.IGNORECASE)

MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
         "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]


def formato(w, h, tol=12):
    for nome, fw, fh in FORMATOS:
        if abs(w - fw) <= tol and abs(h - fh) <= tol:
            return nome
    return "%dx%d pt" % (round(w), round(h))


def categoria_anexo(t):
    n = U.norm(t)
    if re.search(r"\bpa[\s\-]?\d", n):
        return "PARECER_TECNICO"
    if re.search(r"\bnc[\s\-]?\d", n):
        return "NAO_CONFORMIDADE"
    if "desvio" in n and "cronograma" in n:
        return "DESVIO_CRONOGRAMA"
    if "unifilar" in n:
        return "DIAGRAMA_UNIFILAR"
    if "diario" in n:
        return "DIARIO_DE_OBRAS"
    if "ambient" in n:
        return "ACOMP_AMBIENTAL"
    if "seguranca" in n or "sst" in n or "saude" in n:
        return "ACOMP_SST"
    if "ata" in n and "reuni" in n:
        return "ATA_REUNIAO"
    return "OUTRO"


# =============================================================================
# leitura
# =============================================================================
def ler_pdf(caminho):
    """Le o PDF e devolve os fatos necessarios as regras. So texto, sem imagens."""
    doc = fitz.open(U.lp(caminho))
    d = {"n_paginas": doc.page_count, "formatos": Counter(), "texto_a4": [],
         "anexos": [], "pagina_anexos": None}
    try:
        for i, page in enumerate(doc):
            r = page.rect
            fmt = formato(r.width, r.height)
            d["formatos"][fmt] += 1
            if fmt == "A4 retrato":
                t = page.get_text("text")
                d["texto_a4"].append(t)
                if RE_SEC_ANEXOS.search(t) and not d["anexos"]:
                    for m in RE_ANEXO.finditer(t):
                        d["anexos"].append({"ordem": m.group(1),
                                            "titulo": m.group(2).strip(),
                                            "categoria": categoria_anexo(m.group(2))})
                    if d["anexos"]:
                        d["pagina_anexos"] = i + 1
    finally:
        doc.close()
    d["texto"] = "\n".join(d["texto_a4"])
    return d


# =============================================================================
# regras
# =============================================================================
def r(cid, nome, status, detalhe="", evidencia=""):
    return {"id": cid, "regra": nome, "status": status,
            "detalhe": detalhe, "evidencia": evidencia}


def validar(ctx, cfg):
    """ctx: dados do relatorio + pdf lido. cfg: report_versions.yaml

    Cada regra so e aplicada a partir da data em que a caracteristica que ela
    testa passou a existir (cfg['vigencia_das_regras']). Antes disso: N/A.
    Um relatorio antigo NAO falha por seguir corretamente o formato da epoca.
    """
    res = []
    pdf = ctx["pdf"]
    ver = ctx["versao"]
    di = datetime.date.fromisoformat(ctx["data_inicio"])
    df = datetime.date.fromisoformat(ctx["data_fim"])
    vig = cfg.get("vigencia_das_regras", {})

    # o calendario e resolvido pela DATA, nao pela versao de formato
    cal = cfg["calendario"]["cal_v2"]
    lim = cfg["calendario"]["cal_v2"]["effective_from"]
    if isinstance(lim, str):
        lim = datetime.date.fromisoformat(lim)
    if di < lim:
        cal = cfg["calendario"]["cal_v1"]

    def vigente(cid):
        d = vig.get(cid, {}).get("desde")
        if d is None:
            return True
        if isinstance(d, str):
            d = datetime.date.fromisoformat(d)
        return di >= d

    def add(cid, nome, status, detalhe="", evidencia=""):
        if not vigente(cid):
            res.append(r(cid, nome, NA,
                         "anterior à vigência da regra (%s)" % vig[cid]["desde"],
                         vig[cid]["o_que_nasceu"]))
        else:
            res.append(r(cid, nome, status, detalhe, evidencia))

    # --- V1: periodo sabado->sexta ---
    if cal["id"] == "cal_v2":
        ok = di.weekday() == 5 and df.weekday() == 4 and (df - di).days == 6
        add("V1", "Período sábado→sexta, 7 dias",
            PASS if ok else FAIL,
            "%s(%s) a %s(%s)" % (di, di.strftime("%a"), df, df.strftime("%a")),
            "nome da pasta-semana")
    else:
        res.append(r("V1", "Período sábado→sexta, 7 dias", NA,
                     "calendário %s (sexta→quinta)" % cal["id"], "report_versions.calendario"))

    # --- V2: numero calculado bate com o declarado ---
    anc = ctx.get("ancora")
    if anc:
        calc = (di - datetime.date.fromisoformat(anc)).days // 7 + 1
        ok = calc == ctx["numero"]
        add("V2", "Nº = (início − âncora)/7 + 1",
            PASS if ok else FAIL,
            "declarado %d, calculado %d" % (ctx["numero"], calc),
            "road_registry.ancora_semana_1")

    # --- V3: periodo e numero da capa batem com a pasta ---
    m = RE_PERIODO_CAPA.search(pdf["texto"])
    if m:
        cap_i = datetime.datetime.strptime(m.group(1), "%d/%m/%Y").date()
        cap_f = datetime.datetime.strptime(m.group(2), "%d/%m/%Y").date()
        ok = cap_i == di and cap_f == df
        # a pasta pode ter as datas INVERTIDAS no nome (3 casos no acervo).
        # Nesse caso quem erra e o nome da pasta, nao a capa: WARNING, nao FAIL.
        invertida = (not ok) and cap_i == df and cap_f == di
        add("V3", "Período da capa == período da pasta",
            PASS if ok else (WARN if invertida else FAIL),
            ("capa %s a %s — datas INVERTIDAS no nome da pasta; a capa está correta"
             % (cap_i, cap_f)) if invertida else "capa %s a %s" % (cap_i, cap_f),
            "capa, campo PERÍODO")
    else:
        add("V3", "Período da capa == período da pasta", WARN,
            "campo PERÍODO não localizado na capa")

    mn = RE_NUM_CAPA.search(pdf["texto"])
    if mn:
        ok = int(mn.group(1)) == ctx["numero"]
        add("V4", "Nº da capa == nº da pasta",
            PASS if ok else FAIL,
            "capa %s, pasta %d" % (mn.group(1), ctx["numero"]), "capa")
    else:
        add("V4", "Nº da capa == nº da pasta", WARN, "campo Relatório não localizado")

    # --- V5: Diario de Obras = 7 paginas A2 paisagem ---
    n_a2 = pdf["formatos"].get("A2 paisagem", 0)
    esperado = (df - di).days + 1
    diario_nos_anexos = any(a["categoria"] == "DIARIO_DE_OBRAS" for a in pdf["anexos"])
    if n_a2 == 0 and diario_nos_anexos:
        # presente, mas em outro formato de papel (ocorre no historico: A4 paisagem)
        st = WARN
        det = ("Diário declarado nos anexos, mas não em A2 paisagem — formato de papel "
               "anterior à padronização de 2026-05-30")
    elif n_a2 == esperado:
        st, det = PASS, "A2 paisagem = %d, dias = %d" % (n_a2, esperado)
    elif n_a2 and n_a2 % esperado == 0:
        st = WARN
        det = ("A2 paisagem = %d = %d diários — provável diário em atraso "
               "(exceção documentada em report_versions.excecoes_legitimas)"
               % (n_a2, n_a2 // esperado))
    elif n_a2:
        st, det = WARN, "A2 paisagem = %d, esperado %d — diário incompleto" % (n_a2, esperado)
    else:
        st, det = FAIL, "nenhuma página A2 — Diário de Obras ausente"
    add("V5", "Diário de Obras: páginas A2 == dias do período", st, det,
        "formatos de página do PDF")

    # --- V6: Diagrama Unifilar = 1 pagina em formato grande ---
    n_uni = (pdf["formatos"].get("A0 paisagem", 0) + pdf["formatos"].get("A3 paisagem", 0)
             + pdf["formatos"].get("A1 paisagem", 0))
    add("V6", "Diagrama Unifilar: 1 página em formato grande",
        PASS if n_uni == 1 else FAIL,
        "A0/A1/A3 paisagem = %d" % n_uni, "formatos de página")

    # --- V7: Relatorio Fotografico em A3 retrato ---
    n_a3 = pdf["formatos"].get("A3 retrato", 0)
    add("V7", "Relatório Fotográfico em A3 retrato",
        PASS if n_a3 >= 1 else FAIL,
        "A3 retrato = %d páginas" % n_a3, "formatos de página")

    # --- V8: Termo de Encerramento declara o nº real de paginas ---
    mt = RE_TERMO.search(pdf["texto"])
    if mt:
        declarado = int(mt.group(1))
        ok = declarado == pdf["n_paginas"]
        add("V8", "Nº de páginas do Termo == páginas reais",
            PASS if ok else FAIL,
            "termo declara %d, PDF tem %d" % (declarado, pdf["n_paginas"]),
            "seção Termo de Encerramento")
        # --- V9: data de emissao do termo >= fim do periodo ---
        try:
            mes = MESES.index(U.norm(mt.group(5)).replace("marco", "março")) + 1
        except ValueError:
            mes = None
        if mes:
            demis = datetime.date(int(mt.group(6)), mes, int(mt.group(4)))
            ok = demis >= df
            add("V9", "Data de emissão posterior ao fim do período",
                PASS if ok else FAIL,
                "emissão %s, fim do período %s" % (demis, df),
                "Termo de Encerramento")
    else:
        add("V8", "Nº de páginas do Termo == páginas reais", WARN,
            "Termo de Encerramento não localizado no texto")

    # --- V10: tabela climatica com uma coluna por dia ---
    datas = set(RE_TAB_CLIMA.findall(pdf["texto"]))
    add("V10", "Tabela climática: %d datas distintas" % esperado,
        PASS if len(datas) >= esperado else WARN,
        "encontradas %d datas dd/mmm" % len(datas), "seção 5.1")

    # --- V11: cauda dos anexos ---
    cats = [a["categoria"] for a in pdf["anexos"]]
    if cats:
        cauda = (["DIAGRAMA_UNIFILAR", "DIARIO_DE_OBRAS"] if ver["id"] == "v3"
                 else ["DIAGRAMA_UNIFILAR", "DESVIO_CRONOGRAMA", "DIARIO_DE_OBRAS"])
        ok = cats[-len(cauda):] == cauda
        add("V11", "Anexos terminam em %s" % " + ".join(cauda),
            PASS if ok else FAIL,
            "cauda = %s" % cats[-3:], "seção de anexos")
    else:
        add("V11", "Anexos: cauda esperada", WARN, "lista de anexos não localizada")

    # --- V12: bloco ambiental quinzenal ---
    if cats:
        esperado_amb = ((di - datetime.date(2026, 6, 6)).days % 14) == 0
        tem = "ACOMP_AMBIENTAL" in cats and "ACOMP_SST" in cats
        ok = esperado_amb == tem
        add("V12", "Bloco ambiental quinzenal (âncora 2026-06-06)",
            PASS if ok else FAIL,
            "esperado=%s obtido=%s" % (esperado_amb, tem), "seção de anexos")
        # --- V13: ambiental e SST sempre juntos ---
        ok2 = ("ACOMP_AMBIENTAL" in cats) == ("ACOMP_SST" in cats)
        add("V13", "Acompanhamento ambiental e SST sempre juntos",
            PASS if ok2 else FAIL, "cats = %s" % cats, "seção de anexos")

    # --- V14: secao Desvio de Cronograma conforme a versao ---
    tem_desvio = "DESVIO_CRONOGRAMA" in cats
    deve = ver["id"] == "v2"
    if cats:
        add("V14", "Desvio de Cronograma conforme a versão %s" % ver["id"],
            PASS if tem_desvio == deve else FAIL,
            "presente=%s, esperado=%s" % (tem_desvio, deve),
            "report_versions.yaml")

    # --- V15: numero de secoes de nivel 1 ---
    n_sec = len(re.findall(r"^\s*(\d{1,2})\.\s*[A-ZÀ-Ú]", pdf["texto"], re.MULTILINE))
    add("V15", "Estrutura compatível com a versão %s (%d seções)"
        % (ver["id"], ver["n_secoes"]),
        PASS if n_sec >= ver["n_secoes"] else WARN,
        "títulos numerados detectados: %d" % n_sec, "corpo do relatório")

    return res


# =============================================================================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--todos", action="store_true", help="valida o corpus inteiro")
    ap.add_argument("--current", action="store_true",
                    help="valida apenas os relatorios da versao CURRENT (v3)")
    ap.add_argument("--obra", nargs="*", default=None)
    ap.add_argument("--md", default=None, help="caminho do relatorio Markdown")
    ap.add_argument("--ultimas", type=int, default=0, help="N ultimas semanas por obra")
    args = ap.parse_args()

    cfg = yaml.safe_load(open(os.path.join(U.PROJ, "config", "report_versions.yaml"),
                              encoding="utf-8"))
    reg = yaml.safe_load(open(os.path.join(U.PROJ, "config", "road_registry.yaml"),
                              encoding="utf-8"))
    inv = json.load(open(os.path.join(U.OUT, "inventarios",
                                      "inventario_relatorios_semanais.json"), encoding="utf-8"))
    amostra = json.load(open(os.path.join(U.PROJ, "state", "amostra.json"), encoding="utf-8"))

    obras = args.obra or U.OBRAS
    alvos = []
    for obra in obras:
        ss = [e for e in inv["obras"][obra]["entradas"]
              if e["tipo"] == "pasta" and "numero" in e and "data_inicio" in e]
        ss.sort(key=lambda e: e["numero"])
        if args.current:
            ini = cfg["report_versions"]["v3"]["effective_from"]
            if isinstance(ini, str):
                ini = datetime.date.fromisoformat(ini)
            sel = [s for s in ss if datetime.date.fromisoformat(s["data_inicio"]) >= ini]
        elif args.todos:
            sel = ss
        elif args.ultimas:
            sel = ss[-args.ultimas:]
        else:
            nums = [s["numero"] for s in amostra["obras"][obra]["selecionados"]]
            sel = [s for s in ss if s["numero"] in nums]
        for s in sel:
            pdf = next((f for f in s.get("arquivos", [])
                        if f["ext"] == ".pdf" and not f.get("subpasta")
                        and RE_RO_FILE.match(f["nome"])), None)
            if pdf:
                alvos.append((obra, s, pdf))

    def resolver_versao(data_inicio):
        d = datetime.date.fromisoformat(data_inicio)
        for vid in ("v3", "v2", "v1"):
            v = cfg["report_versions"][vid]
            ini = v["effective_from"]
            fim = v.get("effective_until")
            if d >= ini and (fim is None or d <= fim):
                return v
        return cfg["report_versions"][cfg["resolucao"]["padrao_se_indeterminado"]]

    linhas = []
    resumo = Counter()
    por_regra = {}
    print("Validando %d relatorios...\n" % len(alvos))
    for obra, s, pdff in alvos:
        try:
            pdf = ler_pdf(pdff["caminho"])
        except Exception as e:
            print("  ERRO %s n=%s: %s" % (obra, s["numero"], e))
            continue
        ver = resolver_versao(s["data_inicio"])
        ctx = {"obra": obra, "numero": s["numero"], "data_inicio": s["data_inicio"],
               "data_fim": s["data_fim"], "pdf": pdf, "versao": ver,
               "ancora": str(reg["obras"][obra].get("ancora_semana_1"))}
        res = validar(ctx, cfg)
        rot = "%s_%03d" % (obra, s["numero"])
        for x in res:
            resumo[x["status"]] += 1
            por_regra.setdefault(x["id"], Counter())[x["status"]] += 1
        linhas.append({"relatorio": rot, "obra": obra, "numero": s["numero"],
                       "periodo": "%s a %s" % (s["data_inicio"], s["data_fim"]),
                       "versao": ver["id"], "n_paginas": pdf["n_paginas"],
                       "resultados": res})
        falhas = [x for x in res if x["status"] == FAIL]
        warns = [x for x in res if x["status"] == WARN]
        marca = "FAIL" if falhas else ("WARN" if warns else "PASS")
        print("  %-12s %s  %2d pgs  [%s]  %s" % (rot, ver["id"], pdf["n_paginas"], marca,
              "; ".join("%s: %s" % (f["id"], f["detalhe"]) for f in falhas[:2])))

    dest = os.path.join(U.OUT, "validacao")
    os.makedirs(U.assert_read_only(dest), exist_ok=True)
    with open(os.path.join(dest, "resultado.json"), "w", encoding="utf-8") as fh:
        json.dump({"gerado_em": datetime.datetime.now().isoformat(timespec="seconds"),
                   "relatorios": linhas}, fh, ensure_ascii=False, indent=1)

    print("\n" + "=" * 74)
    print("RESUMO: %d PASS · %d FAIL · %d WARNING · %d N/A"
          % (resumo[PASS], resumo[FAIL], resumo[WARN], resumo[NA]))
    print("=" * 74)
    for rid in sorted(por_regra, key=lambda x: int(x[1:])):
        c = por_regra[rid]
        print("  %-4s  PASS %3d  FAIL %3d  WARN %3d  N/A %3d"
              % (rid, c[PASS], c[FAIL], c[WARN], c[NA]))

    if args.md:
        escrever_md(args.md, linhas, resumo, por_regra)
        print("\nMarkdown ->", args.md)
    return 1 if resumo[FAIL] else 0


def escrever_md(caminho, linhas, resumo, por_regra):
    U.assert_read_only(caminho)
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    L = ["# Validação do Relatório Semanal — baseline V2\n",
         "*Gerado por `scripts/validar_relatorio_semanal.py`. O validador apenas **lê**; não gera nada.*\n",
         "\n## Resumo\n",
         "| Resultado | Ocorrências |", "|---|---|",
         "| PASS | %d |" % resumo[PASS], "| FAIL | %d |" % resumo[FAIL],
         "| WARNING | %d |" % resumo[WARN], "| N/A | %d |" % resumo[NA],
         "\n## Por regra\n",
         "| Regra | PASS | FAIL | WARNING | N/A |", "|---|---|---|---|---|"]
    for rid in sorted(por_regra, key=lambda x: int(x[1:])):
        c = por_regra[rid]
        L.append("| %s | %d | %d | %d | %d |" % (rid, c[PASS], c[FAIL], c[WARN], c[NA]))
    L.append("\n## Detalhe por relatório\n")
    L.append("| Regra | Relatório | Versão | Resultado | Evidência |")
    L.append("|---|---|---|---|---|")
    for l in linhas:
        for x in l["resultados"]:
            if x["status"] in (PASS, NA):
                continue
            L.append("| %s — %s | %s | %s | **%s** | %s |"
                     % (x["id"], x["regra"], l["relatorio"], l["versao"],
                        x["status"], x["detalhe"]))
    L.append("\n*(linhas PASS omitidas; ver `output/validacao/resultado.json` para o detalhe completo)*\n")
    L.append("\n## Relatórios validados\n")
    L.append("| Relatório | Período | Versão | Págs | PASS | FAIL | WARN |")
    L.append("|---|---|---|---|---|---|---|")
    for l in linhas:
        c = Counter(x["status"] for x in l["resultados"])
        L.append("| %s | %s | %s | %d | %d | %d | %d |"
                 % (l["relatorio"], l["periodo"], l["versao"], l["n_paginas"],
                    c[PASS], c[FAIL], c[WARN]))
    open(caminho, "w", encoding="utf-8").write("\n".join(L))


if __name__ == "__main__":
    sys.exit(main())
