# -*- coding: utf-8 -*-
r"""
test_ambiental_filtro_rodovia.py - Fase 3.4 (v0.2), FRENTE B: trava a regra
"filtrar por RODOVIA, nunca por empresa executora do CONTRATO" nos paineis
ambientais (Figura 6 "indices" e Figura 7 "condicionantes",
scripts/gerador/figuras/ambiental.py).

Contexto (docs/fontes/dashboard_ambiental_powerbi.md e docstring do modulo):
a v0.1 filtrava `Licencas`/`Condicionantes` pela EMPRESA EXECUTORA DO
CONTRATO (`road_registry.yaml: obras.<OBRA>.empresa_executora_contrato`).
Isso QUEBRAVA POR COMPLETO a RD-102A: a razao social na planilha
("EXECUTORA B") nao tem nenhuma substring em comum
com o nome do contrato ("Consórcio Executora B") -- o filtro antigo
devolvia sempre 0 linhas para essa obra, silenciosamente (LIMITATION).

O filtro correto usa a aba `Empresas` da PROPRIA planilha ambiental (coluna
"Endereco", que guarda a rodovia) para resolver "quais razoes sociais
atendem esta rodovia HOJE" -- nunca comparacao de string com o nome do
contrato.

Este teste tem DUAS camadas de protecao, para pegar uma regressao tanto se
alguem reintroduzir o filtro por empresa quanto se alguem so' "parecer"
corrigir sem de fato usar a aba Empresas:

  R1  as 4 obras (RD101, RD102A, RD103, RD104) resolvem OK (>0 licencas E
      >0 condicionantes) com o filtro atual -- nenhuma cai em LIMITATION
      por "nenhuma licenca/condicionante resolvida para a rodovia".
  R2  RD-102A em particular (o caso que quebrava por completo) resolve OK
      -- reconfirma o achado do briefing, nao so' "as 4 obras no agregado".
  R3  provenance["filtro"] de toda obra comeca com "Rodovia ==" (nunca
      "Empresa" / "Executora" / o nome do contrato) -- pega regressao que
      trocasse so' o RESULTADO mas nao o texto do metodo.
  R4  reproducao do filtro ANTIGO (substring do nome do CONTRATO dentro da
      razao social da planilha) mostra que ele devolveria 0 empresas
      resolvidas para RD-102A -- prova viva de que o bug era real e de que
      o teste discrimina entre os dois algoritmos (nao so' checa "nao e
      zero", checa que o algoritmo CERTO e' o que esta em uso).
  R5  o modulo scripts/gerador/figuras/ambiental.py nao importa
      `road_registry` nem referencia `empresa_executora_contrato` -- guarda
      estatica contra reintroducao do acoplamento errado (rodovia<->empresa
      deve vir SEMPRE da aba `Empresas` da planilha, nunca do cadastro de
      contrato).

Execute:  python tests/test_ambiental_filtro_rodovia.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import re
import sys
import json

import yaml

AQUI = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.dirname(AQUI)
SCRIPTS = os.path.join(PROJ, "scripts")
sys.path.insert(0, SCRIPTS)

import prog_util as U  # noqa: E402
from gerador import config as C  # noqa: E402
from gerador.figuras import ambiental as A  # noqa: E402

OBRAS = ["RD101", "RD102A", "RD103", "RD104"]
PERIODO = "2026-08-29_2026-09-04"
OBRA_QUE_QUEBRAVA = "RD102A"


def _dados_da_obra(obra):
    base = os.path.join(PROJ, "output", "shadow", "%s_%s" % (obra, PERIODO))
    with open(os.path.join(base, "report_data.json"), encoding="utf-8") as f:
        report_data = json.load(f)
    with open(os.path.join(base, "manifest.yaml"), encoding="utf-8") as f:
        manifest = yaml.safe_load(f)
    return report_data, manifest, base


def _norm_simples(s):
    """Normalizacao BEM mais fraca que prog_util.norm -- usada so' aqui para
    reproduzir o algoritmo do filtro ANTIGO (substring), sem reusar a mesma
    funcao de normalizacao do modulo atual (senao o teste R4 nao provaria
    nada -- estaria testando a mesma normalizacao dos dois lados)."""
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def _filtro_antigo_resolveria(wb, empresa_contrato):
    """Reproduz o algoritmo ANTIGO (SHADOW_ERROR, ja corrigido): filtrar
    'Licencas'/'Condicionantes' comparando a razao social da planilha contra
    o nome da EMPRESA EXECUTORA DO CONTRATO (road_registry), por substring
    normalizada. Devolve o numero de razoes sociais da aba 'Empresas' que o
    filtro antigo teria resolvido para essa empresa de contrato -- 0 quando
    quebra (caso documentado da RD-102A)."""
    nomes = [n for n in wb.sheetnames if U.norm(n).startswith("empresa")]
    ws = wb[nomes[0]]
    header_row = A._achar_header_row(ws)
    c_razao = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if isinstance(v, str) and U.norm(v) == "razao social":
            c_razao = c
            break
    alvo = _norm_simples(empresa_contrato)
    achados = 0
    for r in range(header_row + 1, ws.max_row + 1):
        razao = ws.cell(r, c_razao).value
        if not razao:
            continue
        razao_n = _norm_simples(razao)
        if alvo and (alvo in razao_n or razao_n in alvo):
            achados += 1
    return achados


def main():
    falhas = []
    caminho_licenca = C.corporativa("licenca_ambiental")
    wb = A._abrir(caminho_licenca)

    resultados = {}
    for obra in OBRAS:
        report_data, manifest, _ = _dados_da_obra(obra)
        rodovia = report_data["road"]["rodovia"]
        out_dir = os.path.join(PROJ, "tmp", "test_ambiental_filtro_rodovia", obra)
        os.makedirs(out_dir, exist_ok=True)
        r_lic = A.gerar_indices(caminho_licenca, rodovia, out_dir)
        r_cond = A.gerar_condicionantes(caminho_licenca, rodovia, out_dir)
        resultados[obra] = (rodovia, r_lic, r_cond)

    # ---- R1: as 4 obras resolvem OK, nenhuma cai em LIMITATION ----------
    for obra, (rodovia, r_lic, r_cond) in resultados.items():
        if r_lic.status != "OK" or not r_lic.provenance.get("n_licencas"):
            falhas.append("R1: %s (%s): Figura 6 nao resolveu (status=%s, n_licencas=%s) -- "
                          "%s" % (obra, rodovia, r_lic.status,
                                 r_lic.provenance.get("n_licencas"), r_lic.note))
        if r_cond.status != "OK" or not r_cond.provenance.get("n_condicionantes"):
            falhas.append("R1: %s (%s): Figura 7 nao resolveu (status=%s, n_condicionantes=%s) "
                          "-- %s" % (obra, rodovia, r_cond.status,
                                    r_cond.provenance.get("n_condicionantes"), r_cond.note))
    if not any(f.startswith("R1") for f in falhas):
        print("R1  OK   as 4 obras (%s) resolvem licencas E condicionantes com o filtro "
             "por rodovia" % ", ".join(OBRAS))

    # ---- R2: RD-102A em particular (o caso que quebrava) ----------------
    rodovia_178a, r_lic_178a, r_cond_178a = resultados[OBRA_QUE_QUEBRAVA]
    if r_lic_178a.status == "OK" and r_lic_178a.provenance.get("n_licencas", 0) > 0:
        print("R2  OK   %s (%s) resolve %d licencas -- o caso que quebrava por completo com "
             "o filtro por empresa esta OK com o filtro por rodovia"
             % (OBRA_QUE_QUEBRAVA, rodovia_178a, r_lic_178a.provenance["n_licencas"]))
    else:
        falhas.append("R2: %s nao resolveu licencas (status=%s) -- REGRESSAO para o bug "
                      "documentado (filtro por empresa quebrava esta obra)"
                      % (OBRA_QUE_QUEBRAVA, r_lic_178a.status))

    # ---- R3: provenance.filtro sempre diz "Rodovia ==", nunca empresa ---
    for obra, (rodovia, r_lic, r_cond) in resultados.items():
        for nome_fig, r in (("Figura 6", r_lic), ("Figura 7", r_cond)):
            filtro = r.provenance.get("filtro", "")
            if not filtro.startswith("Rodovia =="):
                falhas.append("R3: %s %s: provenance['filtro'] nao comeca com 'Rodovia ==': "
                              "%r" % (obra, nome_fig, filtro))
            for termo_proibido in ("Empresa ==", "empresa_executora", "Executora =="):
                if termo_proibido in filtro:
                    falhas.append("R3: %s %s: provenance['filtro'] contem termo de filtro por "
                                  "empresa (%r): %r" % (obra, nome_fig, termo_proibido, filtro))
    if not any(f.startswith("R3") for f in falhas):
        print("R3  OK   provenance['filtro'] de todas as 8 figuras (4 obras x 2 paineis) diz "
             "'Rodovia ==', nenhuma diz 'Empresa =='")

    # ---- R4: o filtro ANTIGO (substring do nome do contrato) quebraria --
    empresa_contrato_178a = C.obra_info(OBRA_QUE_QUEBRAVA)["empresa_executora_contrato"]
    achados_antigo = _filtro_antigo_resolveria(wb, empresa_contrato_178a)
    if achados_antigo == 0:
        print("R4  OK   filtro ANTIGO (substring de '%s' contra a razao social da planilha) "
             "resolve 0 empresas para %s -- confirma que o bug era real e que o teste "
             "discrimina entre o algoritmo certo e o errado"
             % (empresa_contrato_178a, OBRA_QUE_QUEBRAVA))
    else:
        falhas.append("R4: filtro antigo (substring do nome do contrato) resolveu %d "
                      "empresa(s) para %s -- o cenario de regressao que este teste "
                      "reproduz nao bate mais com o bug documentado; revisar o teste "
                      "(pode ser falso-negativo de protecao)" % (achados_antigo, OBRA_QUE_QUEBRAVA))

    # ---- R5: guarda estatica -- o modulo nao acopla rodovia<->empresa ---
    # Usa `ast` (nao regex sobre o texto) para olhar SO' o codigo executavel,
    # nunca a docstring do modulo -- que CITA de proposito
    # 'empresa_executora_contrato' e 'road_registry' ao explicar o bug
    # historico (isso e' esperado e nao e' regressao). O que nao pode existir
    # e' o modulo IMPORTAR config/road_registry, CHAMAR obra_info(), ou
    # ACESSAR o atributo/chave 'empresa_executora_contrato': a resolucao
    # rodovia<->empresa tem que vir SO' da aba 'Empresas' da propria
    # planilha (_ler_empresas_por_rodovia), nunca do cadastro de contrato.
    import ast
    caminho_modulo = os.path.join(SCRIPTS, "gerador", "figuras", "ambiental.py")
    with open(caminho_modulo, encoding="utf-8") as f:
        arvore = ast.parse(f.read(), filename=caminho_modulo)
    achados_proibidos = []
    for node in ast.walk(arvore):
        if isinstance(node, ast.Import):
            if any(a.name.split(".")[-1] in ("config", "road_registry") for a in node.names):
                achados_proibidos.append("import %s (linha %d)"
                                         % (node.names[0].name, node.lineno))
        elif isinstance(node, ast.ImportFrom):
            if (node.module or "").split(".")[-1] in ("config", "road_registry") or \
                    any(a.name in ("config", "road_registry", "obra_info") for a in node.names):
                achados_proibidos.append("from ... import ... (linha %d)" % node.lineno)
        elif isinstance(node, ast.Call):
            f_node = node.func
            nome = f_node.attr if isinstance(f_node, ast.Attribute) else \
                getattr(f_node, "id", None)
            if nome in ("obra_info", "road_registry"):
                achados_proibidos.append("chamada a %s() (linha %d)" % (nome, node.lineno))
        elif isinstance(node, ast.Attribute) and node.attr == "empresa_executora_contrato":
            achados_proibidos.append("acesso a .empresa_executora_contrato (linha %d)"
                                     % node.lineno)
        elif isinstance(node, ast.Subscript):
            slice_node = node.slice
            valor = getattr(slice_node, "value", slice_node)  # py3.8 vs 3.9+
            if isinstance(valor, ast.Constant) and valor.value == "empresa_executora_contrato":
                achados_proibidos.append("acesso a ['empresa_executora_contrato'] (linha %d)"
                                         % node.lineno)
    if achados_proibidos:
        falhas.append("R5: scripts/gerador/figuras/ambiental.py tem codigo (AST, fora da "
                      "docstring) acoplando rodovia<->empresa ao cadastro de contrato: %s -- "
                      "a chave deve vir SOMENTE da aba 'Empresas' da planilha"
                      % achados_proibidos)
    else:
        print("R5  OK   scripts/gerador/figuras/ambiental.py nao importa config/road_registry, "
             "nao chama obra_info() nem acessa 'empresa_executora_contrato' em codigo "
             "executavel (guarda AST; a mencao na docstring, ao explicar o bug historico, "
             "nao conta)")

    print("=" * 78)
    if falhas:
        print("FALHAS (%d):" % len(falhas))
        for f in falhas:
            print("  X", f)
        print("RESULTADO: FALHOU")
        return 1
    print("RESULTADO: OK (%d obras, filtro por rodovia confirmado sem regressao)"
         % len(OBRAS))
    return 0


if __name__ == "__main__":
    sys.exit(main())
