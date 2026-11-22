# -*- coding: utf-8 -*-
r"""
16_p08_t17_fichas_diarias.py

Varre as Fichas de Campo DIARIAS de uma obra num intervalo de datas e extrai:
  - o bloco de Condicoes Climaticas (N15:P17) -> evidencia para P-08 / 5.1;
  - a tabela de servicos (cabecalho na linha 21) -> qual codigo (K) recebe as
    atividades descritas como OAC/Obras complementares (Q) -> evidencia T-17.

Le so o XML (nao extrai midia). Escreve so em output/evidencias/.
"""
import os
import re
import sys
import json
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from importlib import import_module
M = import_module("15_p08_ficha_acomp_semanal")


def grade(aba):
    g = {}
    for ln in aba["linhas"]:
        g[ln["l"]] = dict(ln["cels"])
    return g


def clima(g):
    """Acha 'Condicoes Climaticas' e devolve (rotulos, coluna marcada)."""
    for L in sorted(g):
        for col, v in g[L].items():
            if "condicoes climaticas" in U.norm(v):
                rot = g.get(L + 1, {})
                marca = g.get(L + 2, {})
                marcado = [rot.get(c) for c in marca
                           if str(marca[c]).strip().upper() in ("X", "1")]
                return {"linha_titulo": L, "rotulos": rot, "marcas": marca,
                        "clima_marcado": [m for m in marcado if m]}
    return None


def cabecalho(g):
    alvo = {"servico": "servico", "situacao": "situacao", "descricao": "descricao",
            "observacao": "observacao", "foto": "foto", "lado": "lado"}
    for L in sorted(g):
        mapa = {}
        for col, v in g[L].items():
            n = re.sub(r"\s*\(.*\)$", "", U.norm(v))
            if n in alvo:
                mapa[alvo[n]] = col
            elif n.startswith("estaca i"):
                mapa["estaca_i"] = col
            elif n.startswith("estaca f"):
                mapa["estaca_f"] = col
        if "servico" in mapa and "descricao" in mapa:
            return L, mapa
    return None, {}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("obra")
    ap.add_argument("--de", required=True, help="AAAA-MM-DD")
    ap.add_argument("--ate", required=True)
    args = ap.parse_args()

    raiz = U.pasta_semanal_raiz(args.obra)
    pasta = None
    for d in U.listdir(raiz):
        if U.norm(d).startswith("1. ficha"):
            pasta = os.path.join(raiz, d)
    arquivos = []
    for dp, dn, fn in U.walk(pasta):
        for f in fn:
            m = re.match(r"(\d{4}-\d{2}-\d{2})", f)
            if m and f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$"):
                if args.de <= m.group(1) <= args.ate:
                    arquivos.append((m.group(1), os.path.join(dp, f)))
    arquivos.sort()

    saida = {"obra": args.obra, "dias": []}
    for data, p in arquivos:
        d = M.dissecar(p, max_linhas=300)
        aba = next((a for a in d["abas"] if U.norm(a["nome"]) == "ficha campo"),
                   d["abas"][0])
        g = grade(aba)
        cl = clima(g)
        lin_h, mapa = cabecalho(g)
        regs = []
        if lin_h:
            for L in sorted(g):
                if L <= lin_h:
                    continue
                cols = g[L]
                s = cols.get(mapa.get("servico", ""))
                if not s:
                    continue
                regs.append({"linha": L, "servico": s,
                             "descricao": cols.get(mapa.get("descricao", "")),
                             "observacao": cols.get(mapa.get("observacao", "")),
                             "foto": cols.get(mapa.get("foto", ""))})
        item = {"data": data, "arquivo": os.path.basename(p), "caminho": p,
                "clima": cl, "cabecalho_linha": lin_h, "colunas": mapa,
                "registros": regs}
        saida["dias"].append(item)

        print("=" * 78)
        print(data, "|", os.path.basename(p))
        if cl:
            print("  CLIMA: rotulos=%s marcas=%s -> %s"
                  % (cl["rotulos"], cl["marcas"], cl["clima_marcado"]))
        else:
            print("  CLIMA: (bloco nao encontrado)")
        for r in regs:
            print("   L%-4d K=%-6s Q=%-38s foto=%-8s X=%s" % (
                r["linha"], str(r["servico"])[:6], str(r["descricao"] or "")[:38],
                str(r["foto"] or "")[:8], str(r["observacao"] or "")[:55]))

    print("\n### CRUZAMENTO T-17: descricao -> codigo de servico ###")
    mapa_t17 = {}
    for dia in saida["dias"]:
        for r in dia["registros"]:
            k = U.norm(r["descricao"] or "")
            mapa_t17.setdefault(k, {}).setdefault(str(r["servico"]), 0)
            mapa_t17[k][str(r["servico"])] += 1
    for k in sorted(mapa_t17):
        print("  %-45s -> %s" % (k[:45], mapa_t17[k]))
    saida["cruzamento_descricao_x_codigo"] = mapa_t17

    fp = os.path.join(U.OUT, "evidencias",
                      "T17_P08_fichas_diarias_%s_%s_%s.json"
                      % (args.obra, args.de, args.ate))
    U.assert_read_only(fp)
    os.makedirs(os.path.dirname(fp), exist_ok=True)
    with open(fp, "w", encoding="utf-8") as fh:
        json.dump(saida, fh, ensure_ascii=False, indent=1)
    print("\nOK ->", fp)


if __name__ == "__main__":
    main()
