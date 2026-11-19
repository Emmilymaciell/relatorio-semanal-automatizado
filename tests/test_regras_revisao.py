# -*- coding: utf-8 -*-
"""
test_regras_revisao.py

Valida a regra de versionamento descoberta nas 5 revisoes `.2` do acervo
(docs/mapeamento/REVISOES_E_VERSIONAMENTO.md):

  REV-1  a raiz da pasta-semana contem EXATAMENTE UMA revisao do relatorio
  REV-2  se existe subpasta de obsoletos, a revisao da raiz e MAIOR que as de la
  REV-3  a revisao vigente tem DOCX e PDF (salvo a semana em elaboracao)
  REV-4  revisoes obsoletas guardam APENAS o PDF
  REV-5  nenhuma revisao .3 ou superior existe no acervo

Execute:  python tests/test_regras_revisao.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import sys
import json
import re

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(AQUI), "scripts"))
import prog_util as U
import datetime

INV = os.path.join(U.OUT, "inventarios", "inventario_relatorios_semanais.json")

# REV-1 e REV-3 so sao EXIGIDAS no modelo vigente (v3). No regime de implantacao
# (2025 e inicio de 2026) houve semanas com so DOCX ou so PDF na raiz — e fato
# historico, nao defeito da regra. Antes desta data: apenas informativo.
INICIO_CURRENT = datetime.date(2026, 8, 1)
RE_RO = re.compile(r"^RO[_\s-]*0*(\d{1,3})(?:\.(\d))?\s*-", re.IGNORECASE)


def eh_obsoleto(subpasta):
    return "obsolet" in U.norm(subpasta or "")


def revisoes(arquivos, na_raiz=True):
    """Devolve {(numero, revisao): set(extensoes)} dos documentos principais."""
    out = {}
    for f in arquivos:
        sub = f.get("subpasta", "")
        if na_raiz and sub:
            continue
        if not na_raiz and not eh_obsoleto(sub):
            continue
        m = RE_RO.match(f["nome"])
        if not m or f["ext"] not in (".pdf", ".docx"):
            continue
        chave = (int(m.group(1)), int(m.group(2)) if m.group(2) else 1)
        out.setdefault(chave, set()).add(f["ext"])
    return out


def main():
    inv = json.load(open(INV, encoding="utf-8"))
    falhas, avisos = [], []
    cont = {k: [0, 0] for k in ("REV-1", "REV-2", "REV-3", "REV-4", "REV-5")}
    com_obsoleto = 0

    for obra in U.OBRAS:
        entradas = [e for e in inv["obras"][obra]["entradas"]
                    if e["tipo"] == "pasta" and "numero" in e]
        ultima = max(e["numero"] for e in entradas)
        for e in entradas:
            arqs = e.get("arquivos", [])
            raiz = revisoes(arqs, na_raiz=True)
            obs = revisoes(arqs, na_raiz=False)
            rot = "%s n=%s" % (obra, e["numero"])
            em_elaboracao = e["numero"] == ultima
            di = e.get("data_inicio")
            no_current = bool(di) and datetime.date.fromisoformat(di) >= INICIO_CURRENT
            registrar = falhas.append if no_current else avisos.append

            # REV-1
            cont["REV-1"][1] += 1
            if len(raiz) == 1:
                cont["REV-1"][0] += 1
            elif not raiz:
                registrar("REV-1 %s (%s): nenhuma revisao na raiz" % (rot, di))
            else:
                registrar("REV-1 %s: %d revisoes na raiz: %s" % (rot, len(raiz), sorted(raiz)))

            if not raiz:
                continue
            rev_raiz = max(r for _, r in raiz)

            # REV-2 e REV-4 (so onde ha obsoletos)
            if obs:
                com_obsoleto += 1
                cont["REV-2"][1] += 1
                if all(r < rev_raiz for _, r in obs):
                    cont["REV-2"][0] += 1
                else:
                    falhas.append("REV-2 %s: raiz .%d, obsoletos %s" % (rot, rev_raiz, sorted(obs)))
                cont["REV-4"][1] += 1
                extras = {k: v for k, v in obs.items() if v != {".pdf"}}
                if not extras:
                    cont["REV-4"][0] += 1
                else:
                    falhas.append("REV-4 %s: obsoleto com %s" % (rot, extras))

            # REV-3
            cont["REV-3"][1] += 1
            exts = raiz[(e["numero"], rev_raiz)] if (e["numero"], rev_raiz) in raiz \
                else list(raiz.values())[0]
            if exts == {".pdf", ".docx"}:
                cont["REV-3"][0] += 1
            elif em_elaboracao and exts == {".docx"}:
                cont["REV-3"][0] += 1     # em elaboracao: so DOCX e o esperado
            else:
                registrar("REV-3 %s (%s): raiz tem %s" % (rot, di, sorted(exts)))

            # REV-5
            cont["REV-5"][1] += 1
            maior = max([r for _, r in raiz] + [r for _, r in obs] or [1])
            if maior <= 2:
                cont["REV-5"][0] += 1
            else:
                falhas.append("REV-5 %s: revisao .%d encontrada" % (rot, maior))

    desc = {
        "REV-1": "raiz contem exatamente uma revisao",
        "REV-2": "revisao da raiz e maior que as obsoletas",
        "REV-3": "revisao vigente tem DOCX e PDF",
        "REV-4": "obsoletos guardam apenas o PDF",
        "REV-5": "nenhuma revisao .3 ou superior",
    }
    print("=" * 74)
    print("VALIDACAO DA REGRA DE VERSIONAMENTO")
    print("=" * 74)
    for k in ("REV-1", "REV-2", "REV-3", "REV-4", "REV-5"):
        ok, tot = cont[k]
        print("  %-6s %4d/%-4d  %s" % (k, ok, tot, desc[k]))
    print("\n  pastas-semana com subpasta de obsoletos: %d" % com_obsoleto)

    if avisos:
        print("\n  AVISOS historicos (anteriores a %s — nao contam como falha): %d"
              % (INICIO_CURRENT, len(avisos)))
        for a in avisos[:10]:
            print("     ~", a)
        if len(avisos) > 10:
            print("     ~ ... +%d" % (len(avisos) - 10))

    if falhas:
        print("\nFALHAS (%d):" % len(falhas))
        for f in falhas[:25]:
            print("  X", f)
        print("\nRESULTADO: FALHOU")
        return 1
    print("\nRESULTADO: REGRA DE VERSIONAMENTO VALIDADA")
    return 0


if __name__ == "__main__":
    sys.exit(main())
