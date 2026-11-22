# -*- coding: utf-8 -*-
"""
test_sem_vazamento_golden.py

Prova a guarda anti-vazamento do golden (spec Sec.5):
`gerador.assert_nao_e_golden(path)` deve levantar PermissionError para:
  G1  o .docx do relatorio oficial (RO_NNN.n - <rodovia>.docx)
  G2  o .pdf do relatorio oficial (RO_NNN.n - <rodovia>.pdf)
  G3  variacoes de grafia observadas no acervo real (espaco antes do hifen,
      revisao de 2+ digitos)
  G4  qualquer arquivo solto na RAIZ de uma pasta-semana
  G5  NAO deve disparar para um arquivo legitimo dentro de 'Arquivos Auxiliares'
      (a extracao real precisa continuar funcionando)

Tambem confere que os extratores realmente chamam a guarda: monkeypatch em
`ficha_campo.assert_nao_e_golden` e força um caminho golden.

Execute:  python tests/test_sem_vazamento_golden.py
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from gerador import assert_nao_e_golden  # noqa: E402
from gerador.extratores import ficha_campo  # noqa: E402

SEMANA = (r"C:\caminho\exemplo\PastaSincronizada\Programa - 3. Desenvolvimento"
         r"\G2 -  Engenharia\RD103\3_Gestao Tecnica de Obra\4. Relatório Semanal Obras"
         r"\5. Relatório Semanal de Acompanhamento de Obras"
         r"\2026.08.29-2026.09.04 49º Relatório Semanal")


def _dispara(caminho):
    try:
        assert_nao_e_golden(caminho)
        return False
    except PermissionError:
        return True


def main():
    falhas = []

    casos_devem_disparar = {
        "G1 docx": os.path.join(SEMANA, "RO_049.1 - RD-103.docx"),
        "G2 pdf": os.path.join(SEMANA, "RO_049.1 - RD-103.pdf"),
        "G3a espaco antes do hifen": os.path.join(SEMANA, "RO_049.1 - RD- 102A.docx"),
        "G3b revisao 2 digitos": os.path.join(SEMANA, "RO_049.12 - RD-103.pdf"),
        "G4 arquivo solto na raiz da semana": os.path.join(SEMANA, "algum_anexo_solto.pdf"),
    }
    for rotulo, caminho in casos_devem_disparar.items():
        if _dispara(caminho):
            print("%-42s OK   PermissionError disparada" % rotulo)
        else:
            falhas.append("%s: NAO disparou para %s" % (rotulo, caminho))

    # G5 — nao pode disparar para um arquivo legitimo (falso positivo quebraria a extracao real)
    legitimo = os.path.join(SEMANA, "Arquivos Auxiliares", "[Programa] Avanço Global de Obra.xlsx")
    if _dispara(legitimo):
        falhas.append("G5: disparou (falso positivo) para arquivo legitimo em Arquivos Auxiliares: %s"
                      % legitimo)
    else:
        print("%-42s OK   nao disparou (falso positivo evitado)" % "G5 arquivo legitimo em auxiliares")

    # G6 — o extrator de fato chama a guarda antes de abrir o arquivo
    golden_docx = os.path.join(SEMANA, "RO_049.1 - RD-103.docx")
    try:
        ficha_campo.ler(golden_docx)
        falhas.append("G6: ficha_campo.ler() nao levantou PermissionError para o golden")
    except PermissionError:
        print("%-42s OK   extrator recusou abrir o golden" % "G6 extrator chama a guarda")

    print("=" * 74)
    if falhas:
        print("FALHAS (%d):" % len(falhas))
        for f in falhas:
            print("  X", f)
        print("RESULTADO: FALHOU")
        return 1
    print("RESULTADO: GUARDA ANTI-VAZAMENTO DO GOLDEN OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
