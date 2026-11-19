# -*- coding: utf-8 -*-
"""
test_config_integridade.py

Testa a INTEGRIDADE dos arquivos de configuracao que passam a governar a
automacao futura. Estes testes representam o PADRAO CORRETO ATUAL — nao
codificam erro historico.

  C1  todo YAML de config carrega
  C2  nenhum alias da taxonomia de servicos aparece em duas entradas diferentes
  C3  toda entrada da taxonomia tem canonical_name e classe valida
  C4  os codigos citados em `codigo_ficha` existem na legenda da Ficha de Campo
  C5  exatamente UMA versao de formato esta `active` (o CURRENT)
  C6  os intervalos de vigencia das versoes nao se sobrepoem e nao deixam buraco
  C7  toda obra do road_registry tem ancora de semana 1 e ela cai num sabado
  C8  a versao CURRENT declara o mesmo numero de secoes que a lista de secoes

Execute:  python tests/test_config_integridade.py
"""
import os
import sys
import datetime

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(AQUI), "scripts"))
import prog_util as U
import yaml

CFG = os.path.join(U.PROJ, "config")
GRUPOS = ["servicos", "controle_tecnologico", "topografia", "meio_ambiente_e_sst", "eventos"]


def carregar(nome):
    return yaml.safe_load(open(os.path.join(CFG, nome), encoding="utf-8"))


def main():
    falhas, ok = [], []

    # ---------- C1 ----------
    cfgs = {}
    for nome in ("services_taxonomy.yaml", "report_versions.yaml",
                 "road_registry.yaml", "report_schema.yaml"):
        try:
            cfgs[nome] = carregar(nome)
            ok.append("C1 %s carrega" % nome)
        except Exception as e:
            falhas.append("C1 %s: %s" % (nome, e))
    if len(cfgs) < 4:
        print("FALHA ao carregar configs"); [print("  X", f) for f in falhas]; return 1

    tax = cfgs["services_taxonomy.yaml"]
    ver = cfgs["report_versions.yaml"]
    reg = cfgs["road_registry.yaml"]

    # ---------- C2: aliases sem colisao ----------
    dono = {}
    colisoes = []
    for g in GRUPOS:
        for chave, ent in tax.get(g, {}).items():
            for a in ent.get("aliases", []):
                k = U.norm(a)
                if k in dono and dono[k] != chave:
                    colisoes.append("'%s' em %s e %s" % (a, dono[k], chave))
                dono[k] = chave
    if colisoes:
        falhas += ["C2 alias duplicado: " + c for c in colisoes]
    else:
        ok.append("C2 %d aliases, nenhuma colisao" % len(dono))

    # ---------- C3: canonical_name e classe ----------
    classes = set(tax["classes"])
    n = 0
    for g in GRUPOS:
        for chave, ent in tax.get(g, {}).items():
            n += 1
            if not ent.get("canonical_name"):
                falhas.append("C3 %s sem canonical_name" % chave)
            c = ent.get("classe")
            if c not in classes:
                falhas.append("C3 %s tem classe invalida: %s" % (chave, c))
    if not any(f.startswith("C3") for f in falhas):
        ok.append("C3 %d entradas com canonical_name e classe valida" % n)

    # ---------- C4: codigos da ficha ----------
    legenda = set(tax["codigos_ficha_de_campo"])
    for g in GRUPOS:
        for chave, ent in tax.get(g, {}).items():
            cod = ent.get("codigo_ficha")
            if cod in (None, "null"):
                continue
            for c in (cod if isinstance(cod, list) else [cod]):
                if c not in legenda:
                    falhas.append("C4 %s referencia codigo inexistente: %s" % (chave, c))
    if not any(f.startswith("C4") for f in falhas):
        ok.append("C4 codigos referenciados existem na legenda (%d codigos)" % len(legenda))

    # ---------- C5: um unico CURRENT ----------
    ativos = [k for k, v in ver["report_versions"].items() if v["status"] == "active"]
    if len(ativos) == 1:
        ok.append("C5 CURRENT = %s" % ativos[0])
    else:
        falhas.append("C5 esperado 1 versao active, encontradas %s" % ativos)

    # ---------- C6: vigencias coerentes ----------
    vs = []
    for k, v in ver["report_versions"].items():
        ini = v["effective_from"]
        fim = v.get("effective_until")
        if isinstance(ini, str):
            ini = datetime.date.fromisoformat(ini)
        if isinstance(fim, str):
            fim = datetime.date.fromisoformat(fim)
        vs.append((ini, fim, k))
    vs.sort()
    for (i1, f1, k1), (i2, f2, k2) in zip(vs, vs[1:]):
        if f1 is None:
            falhas.append("C6 %s sem effective_until mas nao e a ultima" % k1)
        elif f1 >= i2:
            falhas.append("C6 sobreposicao entre %s e %s" % (k1, k2))
        elif (i2 - f1).days > 7:
            falhas.append("C6 lacuna de %d dias entre %s e %s" % ((i2 - f1).days, k1, k2))
    if vs[-1][1] is not None:
        falhas.append("C6 a versao mais recente deveria ter effective_until nulo")
    if not any(f.startswith("C6") for f in falhas):
        ok.append("C6 %d versoes com vigencias coerentes" % len(vs))

    # ---------- C7: ancoras ----------
    for obra, d in reg["obras"].items():
        a = d.get("ancora_semana_1")
        if not a:
            falhas.append("C7 %s sem ancora_semana_1" % obra); continue
        if isinstance(a, str):
            a = datetime.date.fromisoformat(a)
        if a.weekday() != 5:
            falhas.append("C7 %s: ancora %s cai em %s, esperado sabado" % (obra, a, a.strftime("%a")))
    if not any(f.startswith("C7") for f in falhas):
        ok.append("C7 %d obras com ancora em sabado" % len(reg["obras"]))

    # ---------- C8: contagem de secoes ----------
    cur = ver["report_versions"][ativos[0]] if ativos else None
    if cur:
        declarado = cur["n_secoes"]
        real = len(cur["secoes"])
        if declarado == real:
            ok.append("C8 CURRENT declara %d secoes e lista %d" % (declarado, real))
        else:
            falhas.append("C8 CURRENT declara %d secoes mas lista %d" % (declarado, real))
        nums = [s["n"] for s in cur["secoes"]]
        if nums != list(range(1, real + 1)):
            falhas.append("C8 numeracao das secoes nao e 1..%d: %s" % (real, nums))

    print("=" * 74)
    print("INTEGRIDADE DA CONFIGURACAO")
    print("=" * 74)
    for o in ok:
        print("  OK  ", o)
    if falhas:
        print("\nFALHAS (%d):" % len(falhas))
        for f in falhas:
            print("  X  ", f)
        print("\nRESULTADO: FALHOU")
        return 1
    print("\nRESULTADO: CONFIGURACAO INTEGRA")
    return 0


if __name__ == "__main__":
    sys.exit(main())
