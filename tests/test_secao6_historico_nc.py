# -*- coding: utf-8 -*-
"""
test_secao6_historico_nc.py

Testa scripts/gerador/extratores/nc.avaliar_reconstrucao_historica() (spec
Sec.2, §6): estado historico das NC e reconstruivel a partir dos documentos
individuais, ou so a partir da planilha Tbl_VISAO (TODAY(), estado atual)?

  H1  para RD103 (obra com maior volume de NC da amostra), o resultado e
      'UNRECONSTRUCTABLE_FROM_CURRENT_SOURCE', com motivo e impacto NAO
      vazios (nunca um bug silencioso)
  H2  a evidencia cita numeros REAIS levantados do acervo (nao um valor fixo
      copiado do golden nem de outra sessao) -- os campos min sao coerentes
      entre si (total_arquivos_individuais >= soma dos tipos contados)
  H3  a funcao roda para as 4 obras sem excecao (generalidade -- nenhum
      `if obra == ...`)
  H4  nunca usa o valor do golden: a funcao nao abre nenhum arquivo (so lista
      nomes) -- golden nunca aparece na pasta '3. Notificações', entao nem
      precisaria da guarda, mas confirmamos que a evidencia nao referencia
      nenhum caminho RO_NNN.n
  H5  o campo fica embutido em report_data.json ->
      nonconformities.historical_reconstruction (canonical.construir)

Execute:  python tests/test_secao6_historico_nc.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import re
import sys
import datetime as dt

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
from gerador import config as C  # noqa: E402
from gerador.extratores import nc  # noqa: E402
from gerador.canonical import construir  # noqa: E402

_RE_GOLDEN = re.compile(r"RO_\d{3}\.\d+\s*-\s*.+\.(docx|pdf)$", re.IGNORECASE)


def main():
    falhas = []

    # ---------------------------------------------------------------- H1
    r = nc.avaliar_reconstrucao_historica("RD103", contagem_status={"Encerrada": 25})
    if r["historical_nc_state"] == "UNRECONSTRUCTABLE_FROM_CURRENT_SOURCE" and \
       r.get("motivo") and r.get("impacto"):
        print("H1  OK   RD103 -> UNRECONSTRUCTABLE_FROM_CURRENT_SOURCE, com motivo e impacto")
    else:
        falhas.append("H1: %s" % r)

    # ---------------------------------------------------------------- H2
    ev = r["evidencia"]
    soma_tipos = sum(ev["contagem_por_tipo"].values())
    if ev["total_arquivos_individuais"] >= soma_tipos > 0:
        print("H2  OK   evidencia com numeros reais e coerentes: total=%d >= soma_tipos=%d"
             % (ev["total_arquivos_individuais"], soma_tipos))
    else:
        falhas.append("H2: evidencia incoerente: %s" % ev)

    # ---------------------------------------------------------------- H3
    for obra in C.obras_validas():
        try:
            r_obra = nc.avaliar_reconstrucao_historica(obra, contagem_status={"Encerrada": 1})
            assert r_obra["historical_nc_state"] in (
                "UNRECONSTRUCTABLE_FROM_CURRENT_SOURCE", "RECONSTRUCTABLE")
        except Exception as e:
            falhas.append("H3: %s levantou %s: %s" % (obra, type(e).__name__, e))
    if not any(f.startswith("H3") for f in falhas):
        print("H3  OK   avaliar_reconstrucao_historica roda para as %d obras sem excecao"
             % len(C.obras_validas()))

    # ---------------------------------------------------------------- H4
    texto_evidencia = repr(ev)
    if not _RE_GOLDEN.search(texto_evidencia):
        print("H4  OK   evidencia nao referencia nenhum caminho de golden (RO_NNN.n)")
    else:
        falhas.append("H4: evidencia referencia um caminho que parece golden: %s" % texto_evidencia)

    # ---------------------------------------------------------------- H5
    manifest, data, prov = construir("RD103", dt.date(2026, 8, 29), dt.date(2026, 9, 4))
    hr = data.get("nonconformities", {}).get("historical_reconstruction")
    if hr and hr.get("historical_nc_state") == "UNRECONSTRUCTABLE_FROM_CURRENT_SOURCE":
        print("H5  OK   report_data.json -> nonconformities.historical_reconstruction presente")
    else:
        falhas.append("H5: nonconformities.historical_reconstruction ausente ou inesperado: %s" % hr)

    print("=" * 74)
    if falhas:
        print("FALHAS (%d):" % len(falhas))
        for f in falhas:
            print("  X", f)
        print("RESULTADO: FALHOU")
        return 1
    print("RESULTADO: OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
