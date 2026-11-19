# -*- coding: utf-8 -*-
"""Guarda dos testes que dependem de dado real do contrato (versão pública).

Na versão pública, o SharePoint de origem e a pasta `output/` (inventários, shadow
reports, evidências extraídas dos relatórios reais) foram removidos. Os testes que
leem esses dados são pulados, a menos que PROG_DADOS_REAIS=1 esteja definido
(ambiente da autora, com o acervo original disponível).
"""
import os
import sys

MOTIVO = "depende de dado real do contrato — removido na versão pública"


def exigir_dados_reais(arquivo):
    if os.environ.get("PROG_DADOS_REAIS") == "1":
        return
    print("SKIP %s: %s (defina PROG_DADOS_REAIS=1 com o acervo original)"
          % (os.path.basename(arquivo), MOTIVO))
    sys.exit(0)
