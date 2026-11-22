# -*- coding: utf-8 -*-
r"""
scripts/gerador/__init__.py - pacote do relatorio-sombra (Fase 3, Slices 1 e 2).

Contem a guarda anti-vazamento do golden, exigida na Sec. 5 da spec:
    docs/arquitetura/FASE3_SLICE_1_2_SPEC.md

`assert_nao_e_golden(path)` levanta PermissionError se `path` apontar para o
relatorio oficial (RO_NNN.n - <rodovia>.docx/.pdf) ou para qualquer arquivo na
RAIZ de uma pasta-semana ("AAAA.MM.DD-AAAA.MM.DD Nº Relatorio Semanal\...").
Todo extrator deve chamar essa guarda antes de abrir um arquivo. O unico modulo
autorizado a *registrar* (nunca ler dados de) o golden e o source_discovery.
"""
import os
import re

# RO_NNN.n - RD-103.pdf / RO_049.1 - RD- 102A.docx / variações de espaço no hífen
_RE_GOLDEN_ARQUIVO = re.compile(
    r"RO_\d{3}\.\d+\s*-\s*.+\.(docx|pdf)$", re.IGNORECASE)

# "2026.08.29-2026.09.04 49º Relatório Semanal" (a pasta-semana em si)
_RE_PASTA_SEMANA = re.compile(
    r"\d{4}\.\d{2}\.\d{2}-\d{4}\.\d{2}\.\d{2}.*Relat", re.IGNORECASE)


def assert_nao_e_golden(path):
    """Levanta PermissionError se `path` for (ou estiver na raiz de) o golden.

    Duas condicoes disparam a guarda:
      1. o nome do arquivo bate com RO_NNN.n - <rodovia>.{docx,pdf}
      2. o arquivo esta na RAIZ de uma pasta cujo nome bate com o padrao de
         pasta-semana (qualquer arquivo solto ali, nao dentro de
         'Arquivos Auxiliares', e presumidamente golden ou documento correlato
         que nao deve alimentar a extracao).
    """
    nome = os.path.basename(path)
    pai = os.path.basename(os.path.dirname(path))

    if _RE_GOLDEN_ARQUIVO.search(nome):
        raise PermissionError(
            "BLOQUEADO: tentativa de ler o relatorio oficial (golden) como fonte: "
            + path)

    if _RE_PASTA_SEMANA.search(pai) and os.path.splitext(nome)[1].lower() in (
            ".docx", ".pdf"):
        raise PermissionError(
            "BLOQUEADO: arquivo na raiz de uma pasta-semana de Relatorio "
            "Semanal (presumido golden ou correlato, proibido como fonte): "
            + path)

    return path
