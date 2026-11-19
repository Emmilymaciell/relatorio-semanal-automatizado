# -*- coding: utf-8 -*-
r"""
scripts/gerador/extratores/__init__.py - registry dos extratores determinísticos.

Cada modulo le UM tipo de fonte e devolve um dict com pelo menos a chave
`status` (OK | NOT_FOUND | UNREADABLE | NOT_APPLICABLE | MVP_LIMITATION).
Nenhum extrator usa IA; todos chamam `gerador.assert_nao_e_golden(caminho)`
antes de abrir um arquivo.
"""
from . import ficha_campo
from . import ficha_acumulada
from . import avanco
from . import nc
from . import ambiental
from . import fotos
from . import anexos

REGISTRY = {
    "ficha_campo": ficha_campo,
    "ficha_acumulada": ficha_acumulada,
    "avanco": avanco,
    "nc": nc,
    "ambiental": ambiental,
    "fotos": fotos,
    "anexos": anexos,
}
