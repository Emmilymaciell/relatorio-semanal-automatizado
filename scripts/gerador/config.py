# -*- coding: utf-8 -*-
r"""
scripts/gerador/config.py - carrega e valida os YAMLs de config, com cache.

Fonte unica de verdade para as diferencas entre obras (road_registry,
source_layout) e para o modelo/versionamento do documento (report_schema,
report_versions) e a taxonomia de servicos (services_taxonomy).

Nada aqui contem `if obra == "..."`: toda diferenca por obra vem dos proprios
dados do YAML.
"""
import os
import re
import sys
import functools

import yaml

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(AQUI)
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402

CFG_DIR = os.path.join(U.PROJ, "config")

_NOMES = ("road_registry.yaml", "report_schema.yaml", "report_versions.yaml",
          "services_taxonomy.yaml", "source_layout.yaml", "annex_types.yaml",
          "visual_style_v3.yaml")


@functools.lru_cache(maxsize=None)
def _carregar(nome):
    caminho = os.path.join(CFG_DIR, nome)
    with open(caminho, encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def road_registry():
    return _carregar("road_registry.yaml")


def report_schema():
    return _carregar("report_schema.yaml")


def report_versions():
    return _carregar("report_versions.yaml")


def services_taxonomy():
    return _carregar("services_taxonomy.yaml")


def source_layout():
    return _carregar("source_layout.yaml")


def annex_types():
    """config/annex_types.yaml (P0.3) - classificacao EMBEDDED_IMAGE x
    MERGED_PDF por `id` de anexo. Fonte unica de verdade para o modo de
    integracao -- nunca decidir isso por posicao/numero no codigo."""
    return _carregar("annex_types.yaml")


def visual_style():
    """config/visual_style_v3.yaml (P1.0) - fonte unica do design system do
    renderer v2: pagina, margens por perfil, fontes, cores, corpo, h1/h2/h3,
    indice, tabela do clima, legendas, compatibilidade. Proibido numero
    magico espalhado em docx_renderer_v2.py -- tudo vem daqui."""
    return _carregar("visual_style_v3.yaml")


def generation_modes():
    """config/generation_modes.yaml (v0.3) - CURRENT x HISTORICAL_SHADOW e a
    lista dos dados que NAO se reconstroem historicamente (fonte guarda so o
    estado atual). Fonte unica dessa regra -- proibido decidir "isso e'
    volatil" no meio do content_builder."""
    return _carregar("generation_modes.yaml")


MODO_CURRENT = "CURRENT"
MODO_HISTORICAL_SHADOW = "HISTORICAL_SHADOW"


def inferir_modo(fim_periodo, hoje=None):
    """Modo PADRAO a partir da distancia entre o fim do periodo e hoje.

    Um relatorio real sai poucos dias depois do fim da semana (o oficial da
    semana 49 saiu 7 dias depois). Passada a folga configurada, gerar aquela
    semana e' necessariamente uma reconstrucao, e as fontes de estado atual ja
    nao descrevem o periodo.

    E' so o padrao: quem chama pode fixar o modo explicitamente, e o valor
    efetivamente usado e' gravado no report_data -- nunca re-inferido depois.
    """
    import datetime as _dt
    if isinstance(fim_periodo, str):
        fim_periodo = _dt.date.fromisoformat(fim_periodo)
    hoje = hoje or _dt.date.today()
    folga = (generation_modes().get("inferencia", {})
             .get("dias_apos_fim_para_historico", 21))
    return MODO_HISTORICAL_SHADOW if (hoje - fim_periodo).days > folga else MODO_CURRENT


def acao_historica(id_dado):
    """O que fazer com `id_dado` quando o modo e' HISTORICAL_SHADOW.

    Devolve None quando o dado NAO esta na lista -- isto e', quando ele e'
    reconstruivel e a redacao segue normal nos dois modos.
    """
    for d in generation_modes().get("dados_nao_reconstruiveis_historicamente", []):
        if d.get("id") == id_dado:
            return d.get("acao_em_historical_shadow")
    return None


def motivo_historico(id_dado):
    """Motivo + evidencia do dado nao reconstruivel, para a nota do bloco."""
    for d in generation_modes().get("dados_nao_reconstruiveis_historicamente", []):
        if d.get("id") == id_dado:
            return {"motivo": (d.get("motivo") or "").strip(),
                    "evidencia": (d.get("evidencia") or "").strip(),
                    "fonte": d.get("fonte"), "onde": d.get("onde")}
    return None


def validar():
    """Carrega os 5 YAMLs; levanta RuntimeError cedo se algum estiver quebrado."""
    erros = []
    for nome in _NOMES:
        try:
            doc = _carregar(nome)
            if not doc:
                erros.append("%s: vazio" % nome)
        except Exception as e:
            erros.append("%s: %s" % (nome, e))
    if erros:
        raise RuntimeError("Falha ao carregar configuracao: " + "; ".join(erros))
    return True


def obra_info(obra):
    reg = road_registry()
    obras = reg.get("obras", {})
    if obra not in obras:
        raise ValueError("obra desconhecida: %r (validas: %s)"
                          % (obra, sorted(obras)))
    return obras[obra]


def obras_validas():
    return list(road_registry().get("obras", {}))


# --------------------------------------------------------------------------
# resolucao de nomes de pasta variaveis (config/source_layout.yaml)
# --------------------------------------------------------------------------
def resolver_subpasta(base, prefixo=None, iguais=None, regex=None):
    """Acha, dentro de `base`, a 1a subpasta cujo nome normalizado comeca com
    `prefixo` OU e igual a um dos valores de `iguais`. Nao usa nome literal
    por obra - so prefixo/igualdade normalizados (config/source_layout.yaml).
    `regex` (opcional): padrao aplicado com re.search ao nome NORMALIZADO
    (sem acento, minusculo) -- tolera variacoes de numeracao e typos.
    Devolve o caminho completo ou None."""
    if not U.exists(base) or not U.isdir(base):
        return None
    for nome in U.listdir(base):
        caminho = os.path.join(base, nome)
        if not U.isdir(caminho):
            continue
        if "obsoleto" in U.norm(nome):
            continue
        n = U.norm(nome)
        if prefixo and n.startswith(U.norm(prefixo)):
            return caminho
        if iguais and n in {U.norm(v) for v in iguais}:
            return caminho
        if regex and re.search(regex, n):
            return caminho
    return None


def corporativa(chave, **subst):
    """Resolve `source_layout.yaml: corporativas.<chave>.caminho`, um caminho
    fixo (nao descoberto por prefixo/pasta) compartilhado pelas 4 obras --
    ex.: a planilha de licenca ambiental. Substitui `{base_obras}` e
    `{compartilhado}` pelas raizes de `source_layout.yaml: raizes`, e
    quaisquer outros `{...}` recebidos em `subst` (ex.: `pasta_obra_num`).

    Fonte unica de verdade do CAMINHO -- nunca literal duplicado dentro de um
    script (Fase 3.4, SOURCE FIX: decisao de qual copia da planilha ambiental
    e' a operacional, ver docs/fontes/acompanhamento_ambiental.md Sec.6,
    registrada aqui, nao no codigo)."""
    layout = source_layout()
    entrada = layout.get("corporativas", {}).get(chave)
    if not entrada or "caminho" not in entrada:
        raise ValueError("corporativas.%s.caminho nao definido em source_layout.yaml" % chave)
    raizes = layout.get("raizes", {})
    valores = {"base_obras": raizes.get("base_obras", ""),
              "compartilhado": raizes.get("compartilhado", "")}
    valores.update(subst)
    caminho_rel = entrada["caminho"].format(**valores)
    return os.path.join(U.SHAREPOINT, caminho_rel)


def resolver_arquivo(base, prefixo=None, iguais=None, extensoes=None, regex=None):
    """Como `resolver_subpasta`, mas para arquivos (nao pastas). `iguais`
    compara o nome SEM extensao (normalizado); `prefixo` compara o nome
    completo (com extensao), entao tambem funciona por prefixo puro."""
    if not U.exists(base) or not U.isdir(base):
        return None
    for nome in U.listdir(base):
        caminho = os.path.join(base, nome)
        if U.isdir(caminho):
            continue
        if "obsoleto" in U.norm(nome):
            continue
        if extensoes and os.path.splitext(nome)[1].lower() not in extensoes:
            continue
        n = U.norm(nome)
        n_sem_ext = U.norm(os.path.splitext(nome)[0])
        if prefixo and n.startswith(U.norm(prefixo)):
            return caminho
        if iguais and n_sem_ext in {U.norm(v) for v in iguais}:
            return caminho
        if regex and re.search(regex, n):
            return caminho
    return None


def norm_rodovia(s):
    """Normaliza rotulo de rodovia para COMPARACAO: sem acento, maiusculo, sem
    espacos/hifens/underscores. RD-102 A (road_registry) == RD-102A."""
    if not s:
        return ""
    return re.sub(r"[\s\-_]+", "", U.norm(s)).upper()
