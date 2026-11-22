# -*- coding: utf-8 -*-
"""
tests/test_cache_tmp.py - multiobra-r1: prog_util.copiar_para_tmp nao pode servir
a copia de OUTRA fonte com o mesmo tamanho.

O cache usa nome fixo por subdir (ex.: tmp/avanco/prog_avanco.xlsx) para todas as
obras e semanas. Ate o r1 so o tamanho era conferido: a RD-101 leu um arquivo
cacheado de outra rodada e o avanco saiu UNREADABLE (ou, pior, poderia sair o numero
de outra obra). Agora tamanho E mtime.
"""
import os
import sys
import tempfile
import time

TMPDIR = tempfile.mkdtemp(prefix="test_cache_tmp_")
os.environ["PROG_TMP"] = os.path.join(TMPDIR, "cache")
AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "scripts"))
import prog_util as U  # noqa: E402


def main():
    a = os.path.join(TMPDIR, "obra_a.xlsx")
    b = os.path.join(TMPDIR, "obra_b.xlsx")
    with open(a, "wb") as f:
        f.write(b"AAAA")
    with open(b, "wb") as f:
        f.write(b"BBBB")                       # mesmo tamanho, conteudo diferente
    os.utime(a, (time.time() - 3600, time.time() - 3600))
    d1 = U.copiar_para_tmp(a, "avanco", "prog_avanco.xlsx")
    d2 = U.copiar_para_tmp(b, "avanco", "prog_avanco.xlsx")
    conteudo = open(d2, "rb").read()
    falhas = []
    ok1 = d1 == d2 and conteudo == b"BBBB"
    print("C1  %s  mesma chave de cache, mesmo tamanho: serve a fonte nova (%r)"
          % ("OK  " if ok1 else "FAIL", conteudo))
    if not ok1:
        falhas.append("C1")
    m_antes = os.path.getmtime(d2)
    U.copiar_para_tmp(b, "avanco", "prog_avanco.xlsx")
    ok2 = os.path.getmtime(d2) == m_antes and int(m_antes) == int(os.path.getmtime(b))
    print("C2  %s  fonte inalterada: cache reaproveitado, mtime preservado"
          % ("OK  " if ok2 else "FAIL"))
    if not ok2:
        falhas.append("C2")
    print("RESULTADO: %s" % ("OK" if not falhas else "FALHA %s" % falhas))
    return 1 if falhas else 0


if __name__ == "__main__":
    sys.exit(main())
