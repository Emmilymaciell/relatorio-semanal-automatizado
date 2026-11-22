# -*- coding: utf-8 -*-
"""
01_inventariar_relatorios.py
Inventario READ-ONLY das pastas "5. Relatorio Semanal de Acompanhamento de Obras"
das 4 obras do PROG. Nao escreve nada no SharePoint.

Saida: output/inventarios/inventario_relatorios_semanais.json
       output/inventarios/semanas.csv
       output/inventarios/arquivos.csv
"""
import os
import sys
import json
import csv
import re
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

RE_PERIODO = re.compile(r"(\d{4})[.\-_](\d{2})[.\-_](\d{2})\s*[-a–]\s*(\d{4})[.\-_](\d{2})[.\-_](\d{2})")
RE_NUM = re.compile(r"(\d{1,4})\s*[º°o]\s*Relat", re.IGNORECASE)
RE_RO = re.compile(r"\bRO[_\s-]*(\d{1,3})(?:\.(\d))?", re.IGNORECASE)
IGNORAR = re.compile(r"^(~\$|\.~lock)|^desktop\.ini$|^Thumbs\.db$", re.IGNORECASE)


def parse_nome_semana(nome):
    """Extrai periodo e numero do nome da pasta-semana.
    Tolera o caso em que o dia grudou no numero (ex: '2026.02.0619o Relatorio')."""
    d = {}
    m = RE_PERIODO.search(nome)
    if m:
        try:
            d["data_inicio"] = str(datetime.date(int(m[1]), int(m[2]), int(m[3])))
            d["data_fim"] = str(datetime.date(int(m[4]), int(m[5]), int(m[6])))
        except ValueError:
            d["periodo_invalido"] = True
    m2 = RE_NUM.search(nome)
    if m2:
        n = int(m2[1])
        if n > 200:
            n = int(str(n)[-2:])   # dia colado no numero
            d["nome_irregular"] = True
        d["numero"] = n
    return d


def coletar_obra(obra):
    raiz = U.pasta_relatorios(obra)
    if not raiz:
        return {"erro": "pasta '5. Relatorio Semanal...' nao encontrada"}
    entradas = []
    for nome in U.listdir(raiz):
        if IGNORAR.search(nome):
            continue
        p = os.path.join(raiz, nome)
        is_dir = U.isdir(p)
        item = {"nome": nome, "caminho": p, "tipo": "pasta" if is_dir else "arquivo"}
        item.update(parse_nome_semana(nome))
        if is_dir:
            filhos = []
            for dirpath, _, filenames in U.walk(p):
                rel = os.path.relpath(dirpath, p)
                for f in filenames:
                    if IGNORAR.search(f):
                        continue
                    fp = os.path.join(dirpath, f)
                    d = {"nome": f,
                         "rel": f if rel == "." else os.path.join(rel, f),
                         "subpasta": "" if rel == "." else rel,
                         "ext": os.path.splitext(f)[1].lower(),
                         "caminho": fp}
                    try:
                        d.update(U.stat_info(fp))
                    except OSError as e:
                        d["erro"] = str(e)
                    mr = RE_RO.search(f)
                    if mr:
                        d["ro_num"] = int(mr[1])
                        d["ro_rev"] = mr[2]
                    filhos.append(d)
            item["arquivos"] = filhos
            item["qtd_arquivos"] = len(filhos)
        else:
            item["ext"] = os.path.splitext(nome)[1].lower()
            try:
                item.update(U.stat_info(p))
            except OSError as e:
                item["erro"] = str(e)
        entradas.append(item)
    return {"raiz": raiz, "entradas": entradas}


def main():
    res = {"gerado_em": datetime.datetime.now().isoformat(timespec="seconds"),
           "base": U.BASE_OBRAS, "obras": {}}
    for obra in U.OBRAS:
        res["obras"][obra] = coletar_obra(obra)

    out = os.path.join(U.OUT, "inventarios")
    os.makedirs(out, exist_ok=True)
    with open(os.path.join(out, "inventario_relatorios_semanais.json"), "w", encoding="utf-8") as fh:
        json.dump(res, fh, ensure_ascii=False, indent=2)

    with open(os.path.join(out, "semanas.csv"), "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["obra", "numero", "data_inicio", "data_fim", "pasta", "qtd_arquivos", "nome_irregular"])
        for obra, d in res["obras"].items():
            for e in d.get("entradas", []):
                if e["tipo"] == "pasta":
                    w.writerow([obra, e.get("numero", ""), e.get("data_inicio", ""), e.get("data_fim", ""),
                                e["nome"], e.get("qtd_arquivos", ""), e.get("nome_irregular", "")])

    with open(os.path.join(out, "arquivos.csv"), "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["obra", "numero", "pasta", "subpasta", "arquivo", "ext",
                    "tamanho_bytes", "modificado", "erro"])
        for obra, d in res["obras"].items():
            for e in d.get("entradas", []):
                for f in e.get("arquivos", []):
                    w.writerow([obra, e.get("numero", ""), e["nome"], f.get("subpasta", ""), f["nome"],
                                f["ext"], f.get("tamanho_bytes", ""), f.get("modificado", ""), f.get("erro", "")])

    erros = 0
    for obra, d in res["obras"].items():
        ps = [e for e in d.get("entradas", []) if e["tipo"] == "pasta"]
        nums = sorted(e["numero"] for e in ps if "numero" in e)
        er = sum(1 for e in ps for f in e.get("arquivos", []) if "erro" in f)
        erros += er
        falt = [n for n in range(1, max(nums) + 1) if n not in nums] if nums else []
        dup = sorted({n for n in nums if nums.count(n) > 1})
        print("%-8s %2d semanas | n=%s..%s | faltando %s | duplicados %s | sem numero %d | erros_stat %d"
              % (obra, len(ps), min(nums) if nums else "-", max(nums) if nums else "-",
                 falt, dup, sum(1 for e in ps if "numero" not in e), er))
    print("TOTAL erros de stat:", erros)
    print("OK ->", out)


if __name__ == "__main__":
    main()
