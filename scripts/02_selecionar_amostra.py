# -*- coding: utf-8 -*-
"""
02_selecionar_amostra.py

Seleciona a amostra de relatorios para a engenharia reversa:
  - ordena as semanas cronologicamente;
  - identifica a mais recente de cada obra e a marca como EM ELABORACAO (excluida);
  - seleciona os 2 relatorios concluidos imediatamente anteriores.

Criterio objetivo de "concluido": existe PDF final na pasta.
Tambem analisa a evolucao da nomenclatura/volume ao longo do tempo para
detectar trocas de modelo (que justificariam ampliar a amostra).

Saida: output/inventarios/amostra.json + state/amostra.json
"""
import os
import sys
import json
import re
import datetime
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

INV = os.path.join(U.OUT, "inventarios", "inventario_relatorios_semanais.json")
RE_RO_FILE = re.compile(r"^RO[_\s-]*0*(\d{1,3})(?:\.(\d))?\s*-", re.IGNORECASE)


def semanas_da_obra(dados):
    ss = [e for e in dados.get("entradas", []) if e["tipo"] == "pasta" and "numero" in e]
    ss.sort(key=lambda e: (e.get("data_inicio") or "0000-00-00", e["numero"]))
    return ss


def classificar(semana):
    """Separa os arquivos da semana em documento principal, auxiliares e outros."""
    doc = {"pdf": [], "docx": [], "aux": [], "outros": []}
    for f in semana.get("arquivos", []):
        sub = f.get("subpasta", "")
        principal = (sub == "") and bool(RE_RO_FILE.match(f["nome"]))
        if principal and f["ext"] == ".pdf":
            doc["pdf"].append(f)
        elif principal and f["ext"] == ".docx":
            doc["docx"].append(f)
        elif sub.lower().startswith("arquivos aux"):
            doc["aux"].append(f)
        else:
            doc["outros"].append(f)
    return doc


def assinatura_modelo(semana, cl):
    """Assinatura grosseira do 'modelo' daquela semana: padrao de nomes + conjunto de auxiliares."""
    principal = (cl["pdf"] + cl["docx"])
    nome = principal[0]["nome"] if principal else ""
    # generaliza o numero do relatorio para comparar nomes entre semanas
    padrao = re.sub(r"\d+", "#", nome)
    aux = sorted({re.sub(r"\d+", "#", a["nome"]) for a in cl["aux"]})
    outros = sorted({re.sub(r"\d+", "#", o["rel"]) for o in cl["outros"]})
    return {"padrao_nome_principal": padrao, "auxiliares": aux, "outros": outros}


def main():
    inv = json.load(open(INV, encoding="utf-8"))
    saida = {"gerado_em": datetime.datetime.now().isoformat(timespec="seconds"), "obras": {}}

    for obra in U.OBRAS:
        ss = semanas_da_obra(inv["obras"][obra])
        linhas = []
        for s in ss:
            cl = classificar(s)
            linhas.append({
                "numero": s["numero"], "data_inicio": s.get("data_inicio"), "data_fim": s.get("data_fim"),
                "pasta": s["nome"], "caminho": s["caminho"],
                "tem_pdf": bool(cl["pdf"]), "tem_docx": bool(cl["docx"]),
                "pdf": cl["pdf"][0] if cl["pdf"] else None,
                "docx": cl["docx"][0] if cl["docx"] else None,
                "qtd_aux": len(cl["aux"]), "qtd_outros": len(cl["outros"]),
                "assinatura": assinatura_modelo(s, cl),
            })

        mais_recente = linhas[-1]
        # concluidos = tem PDF final; pega os 2 ultimos antes do mais recente
        anteriores = [l for l in linhas[:-1] if l["tem_pdf"]]
        selecionados = anteriores[-2:]

        # evolucao de modelo: agrupa semanas por assinatura, em ordem
        grupos = []
        for l in linhas:
            chave = json.dumps(l["assinatura"], sort_keys=True, ensure_ascii=False)
            if grupos and grupos[-1]["chave"] == chave:
                grupos[-1]["ate"] = l["numero"]
                grupos[-1]["n"] += 1
            else:
                grupos.append({"chave": chave, "de": l["numero"], "ate": l["numero"], "n": 1,
                               "assinatura": l["assinatura"]})

        saida["obras"][obra] = {
            "total_semanas": len(linhas),
            "mais_recente": {k: mais_recente[k] for k in
                             ("numero", "data_inicio", "data_fim", "pasta", "tem_pdf", "tem_docx")},
            "excluido_por": "mais recente da serie; sem PDF final -> em elaboracao"
                            if not mais_recente["tem_pdf"] else "mais recente da serie",
            "selecionados": [{k: s[k] for k in
                              ("numero", "data_inicio", "data_fim", "pasta", "caminho",
                               "tem_pdf", "tem_docx", "pdf", "docx", "qtd_aux", "qtd_outros")}
                             for s in selecionados],
            "regimes_de_modelo": [{"de": g["de"], "ate": g["ate"], "n": g["n"],
                                   "padrao_nome": g["assinatura"]["padrao_nome_principal"],
                                   "auxiliares": g["assinatura"]["auxiliares"]} for g in grupos],
            "linhas": linhas,
        }

    os.makedirs(os.path.join(U.PROJ, "state"), exist_ok=True)
    for p in (os.path.join(U.OUT, "inventarios", "amostra.json"),
              os.path.join(U.PROJ, "state", "amostra.json")):
        with open(p, "w", encoding="utf-8") as fh:
            json.dump(saida, fh, ensure_ascii=False, indent=2)

    # ---- relatorio em tela ----
    for obra, d in saida["obras"].items():
        print("=" * 78)
        mr = d["mais_recente"]
        print("%s | %d semanas" % (obra, d["total_semanas"]))
        print("  EXCLUIDO (em elaboracao): n=%s  %s..%s  pdf=%s docx=%s"
              % (mr["numero"], mr["data_inicio"], mr["data_fim"], mr["tem_pdf"], mr["tem_docx"]))
        for s in d["selecionados"]:
            print("  SELECIONADO: n=%s  %s..%s  pdf=%s(%.1f MB) docx=%s(%.1f MB) aux=%d outros=%d"
                  % (s["numero"], s["data_inicio"], s["data_fim"],
                     "sim" if s["tem_pdf"] else "NAO",
                     (s["pdf"]["tamanho_bytes"] / 1e6) if s["pdf"] else 0,
                     "sim" if s["tem_docx"] else "NAO",
                     (s["docx"]["tamanho_bytes"] / 1e6) if s["docx"] else 0,
                     s["qtd_aux"], s["qtd_outros"]))
        print("  regimes de modelo detectados: %d" % len(d["regimes_de_modelo"]))
        for g in d["regimes_de_modelo"]:
            print("     n %3d..%-3d (%2d sem) nome=%-28s aux=%d"
                  % (g["de"], g["ate"], g["n"], g["padrao_nome"][:28], len(g["auxiliares"])))
    print("\nOK ->", os.path.join(U.OUT, "inventarios", "amostra.json"))


if __name__ == "__main__":
    main()
