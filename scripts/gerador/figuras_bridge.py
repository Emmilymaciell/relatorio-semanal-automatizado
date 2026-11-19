# -*- coding: utf-8 -*-
r"""
scripts/gerador/figuras_bridge.py - Fase 3.3: ponte de integracao com
scripts/gerador/figuras/ (pacote da FRENTE A), pelo CONTRATO do Sec.1 da
spec (docs/arquitetura/FASE3_3_FINAL_CANDIDATE_SPEC.md):

    @dataclass
    class FiguraResult:
        path: str | None
        largura_mm: float
        altura_mm: float
        status: str            # OK | LIMITATION | PRESENTATION_DEPENDENCY | HITL
        note: str | None
        provenance: dict

Este modulo e' territorio da FRENTE B (renderer/anexos/pipeline) -- NUNCA cria
nem edita nada dentro de scripts/gerador/figuras/. Import TOLERANTE: enquanto
o pacote (ou uma funcao especifica dele) nao existir, `tentar()` devolve None
e quem chama cai no placeholder atual, com o motivo -- nunca falha a
montagem do documento nem reimplementa a figura aqui.

`campo()` le um FiguraResult de forma duck-typed (dataclass OU dict), porque
esta ponte nao importa a classe da Frente A (evita acoplamento a um modulo
que pode nao existir ainda).
"""
import importlib


def tentar(modulo, funcao, **kwargs):
    """Tenta obter um FiguraResult de
    `gerador.figuras.<modulo>.<funcao>(**kwargs)`. Devolve o objeto (dataclass
    ou dict, ver `campo()`) ou None se o pacote/modulo/funcao ainda nao
    existir, ou se a chamada falhar por qualquer motivo -- uma figura que a
    Frente A ainda nao terminou nunca deve derrubar o pipeline da Frente B."""
    try:
        mod = importlib.import_module("gerador.figuras.%s" % modulo)
        fn = getattr(mod, funcao)
    except Exception:
        return None
    try:
        return fn(**kwargs)
    except Exception:
        return None


def campo(resultado, nome, default=None):
    """Le um campo de FiguraResult aceitando tanto dataclass/objeto quanto
    dict -- esta ponte nao conhece a classe concreta da Frente A."""
    if resultado is None:
        return default
    if isinstance(resultado, dict):
        return resultado.get(nome, default)
    return getattr(resultado, nome, default)


# Contrato Sec.1: FiguraResult.status usa o vocabulario
# OK|LIMITATION|PRESENTATION_DEPENDENCY|HITL -- diferente do vocabulario de
# document_model.Block.status (READY|PLACEHOLDER|HITL|LIMITATION). Esta
# ponte traduz um para o outro -- nunca deixa "OK" vazar para dentro de um
# Block (o document_model levantaria ValueError).
_MAPA_STATUS = {
    "OK": "READY",
    "LIMITATION": "LIMITATION",
    "HITL": "HITL",
    "PRESENTATION_DEPENDENCY": "LIMITATION",
}


def mapear_status(status_figura, default="PLACEHOLDER"):
    """Traduz um FiguraResult.status (contrato Sec.1) para um
    document_model.Block.status válido."""
    return _MAPA_STATUS.get(status_figura, default)


def disponivel(resultado):
    """True quando o resultado tem uma imagem de verdade para inserir
    (`path` preenchido) -- caso contrario quem chama deve manter o
    placeholder atual (figura ainda PENDING/LIMITATION/PRESENTATION_DEPENDENCY
    sem arquivo, ou pacote inexistente => resultado None)."""
    return bool(campo(resultado, "path"))
