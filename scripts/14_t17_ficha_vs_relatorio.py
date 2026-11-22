# -*- coding: utf-8 -*-
r"""
14_t17_ficha_vs_relatorio.py - fecha T-17.

Le os JSON ja extraidos das Fichas de Campo (output/extracao/excel/*_ficha_campo.json),
detecta dinamicamente o cabecalho (a linha varia por obra!), extrai:
  - a LEGENDA de codigos de servico da propria ficha (bloco "Servicos a serem
    registrados:" na coluna D/F);
  - os codigos efetivamente usados na coluna Servico e as Descricoes associadas.
Cruza com as abas do RELATORIO FOTOGRAFICO da mesma semana.

Nao toca no SharePoint. Escreve so em output/evidencias/.
"""
import os
import re
import sys
import json
import glob

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

OUTDIR = os.path.join(U.OUT, "extracao", "excel")


def parse_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref or "")
    return (m.group(1), int(m.group(2))) if m else (None, None)


def grade(aba):
    """{linha: {coluna: valor}}"""
    g = {}
    for c in aba.get("celulas", []):
        col, lin = parse_ref(c["ref"])
        if col:
            g.setdefault(lin, {})[col] = c.get("v")
    return g


def achar_cabecalho(g):
    """Acha a linha do cabecalho da tabela e o mapa {campo: coluna}."""
    alvo = {"estaca i": "estaca_i", "estaca inicial": "estaca_i",
            "estaca f": "estaca_f", "estaca final": "estaca_f",
            "lado": "lado", "servico": "servico", "situacao": "situacao",
            "foto": "foto", "descricao": "descricao",
            "observacoes": "observacao", "observacao": "observacao"}
    for lin in sorted(g):
        cols = g[lin]
        mapa = {}
        for col, v in cols.items():
            n = U.norm(v)
            n = re.sub(r"\s*\(.*\)$", "", n)
            if n in alvo:
                mapa[alvo[n]] = col
            elif n.startswith("extensao"):
                mapa["extensao"] = col
        if "servico" in mapa and "situacao" in mapa:
            return lin, mapa
    return None, {}


def achar_legenda(g):
    """Bloco 'Servicos a serem registrados:' -> {codigo: nome}."""
    inicio = None
    col_cod = col_nome = None
    for lin in sorted(g):
        for col, v in g[lin].items():
            if "servicos a serem registrados" in U.norm(v):
                inicio, col_cod = lin, col
    if inicio is None:
        return {}, None
    # a coluna do nome e a proxima coluna preenchida a direita nas linhas seguintes
    leg = {}
    for lin in range(inicio + 1, inicio + 25):
        cols = g.get(lin, {})
        cod = cols.get(col_cod)
        if not cod or not re.match(r"^[A-Z]{3}$", str(cod).strip()):
            continue
        nome = None
        for c in sorted(cols, key=lambda x: (len(x), x)):
            if c != col_cod and c > col_cod and cols[c]:
                nome = cols[c]
                break
        leg[str(cod).strip()] = nome
        col_nome = col_nome or c
    return leg, (inicio, col_cod)


def analisar_ficha(path):
    doc = json.load(open(path, encoding="utf-8"))
    res = {"json": os.path.basename(path), "arquivo": doc.get("arquivo"),
           "caminho": doc.get("caminho"), "obra": doc.get("obra"),
           "semana": doc.get("semana"), "periodo": doc.get("periodo"),
           "legenda": {}, "cabecalho_linha": None, "colunas": {},
           "registros": [], "por_codigo": {}}
    for a in doc.get("abas", []):
        if U.norm(a.get("nome") or "") != "ficha campo":
            continue
        g = grade(a)
        res["legenda"], res["legenda_origem"] = achar_legenda(g)
        lin_h, mapa = achar_cabecalho(g)
        res["cabecalho_linha"], res["colunas"] = lin_h, mapa
        if not lin_h:
            break
        maxl = max(g) if g else 0
        for L in range(lin_h + 1, maxl + 1):
            cols = g.get(L, {})
            serv = cols.get(mapa.get("servico", ""))
            if not serv:
                continue
            reg = {"linha": L, "servico": str(serv).strip()}
            for campo in ("estaca_i", "estaca_f", "lado", "extensao",
                          "situacao", "foto", "descricao", "observacao"):
                if campo in mapa:
                    reg[campo] = cols.get(mapa[campo])
            res["registros"].append(reg)
        for r in res["registros"]:
            d = res["por_codigo"].setdefault(r["servico"], {"n": 0, "descricoes": {}})
            d["n"] += 1
            k = str(r.get("descricao") or "")
            d["descricoes"][k] = d["descricoes"].get(k, 0) + 1
        break
    return res


def abas_fotografico(obra, semana):
    p = os.path.join(OUTDIR, "%s_S%02d_relatorio_fotografico.json" % (obra, semana))
    if not os.path.exists(p):
        return None
    d = json.load(open(p, encoding="utf-8"))
    return [{"nome": a.get("nome"), "estado": a.get("estado"),
             "fotos": len(a.get("celulas_com_imagem_na_celula", []))}
            for a in d.get("abas", [])]


def main():
    saida = {"fichas": []}
    for p in sorted(glob.glob(os.path.join(OUTDIR, "*_ficha_campo.json"))):
        r = analisar_ficha(p)
        r["abas_relatorio_fotografico"] = abas_fotografico(r["obra"], r["semana"])
        saida["fichas"].append(r)

        print("=" * 80)
        print("%s | %s | semana %s | %s" % (r["obra"], r["arquivo"], r["semana"],
                                            r["periodo"]))
        print("cabecalho na linha %s -> %s" % (r["cabecalho_linha"], r["colunas"]))
        print("--- LEGENDA da propria ficha (%s) ---" % (r.get("legenda_origem"),))
        for k in r["legenda"]:
            print("   %-5s = %s" % (k, r["legenda"][k]))
        print("--- codigos USADOS na coluna Servico ---")
        for k in sorted(r["por_codigo"]):
            d = r["por_codigo"][k]
            descs = sorted(d["descricoes"].items(), key=lambda x: -x[1])
            print("   %-8s n=%-4d  descricoes: %s" % (
                k, d["n"], "; ".join("%s(%d)" % (a[:55], b) for a, b in descs[:6])))
        print("--- abas RELATORIO FOTOGRAFICO ---")
        for a in (r["abas_relatorio_fotografico"] or []):
            print("   %-28s %-8s fotos=%s" % (a["nome"], a["estado"], a["fotos"]))

    d = os.path.join(U.OUT, "evidencias")
    U.assert_read_only(d)
    os.makedirs(d, exist_ok=True)
    fp = os.path.join(d, "T17_ficha_vs_fotografico.json")
    with open(fp, "w", encoding="utf-8") as fh:
        json.dump(saida, fh, ensure_ascii=False, indent=1)
    print("\nOK ->", fp)


if __name__ == "__main__":
    main()
