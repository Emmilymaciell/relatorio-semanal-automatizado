# -*- coding: utf-8 -*-
"""
test_assembly_model.py

Testa scripts/gerador/assembly_model.py (Fase 3.2B, P0.2/P0.3) e
config/annex_types.yaml.

  T1  os 7 tipos de anexo de config/annex_types.yaml existem, com
      integration_mode valido (EMBEDDED_IMAGE ou MERGED_PDF)
  T2  classificacao correta e' EXATAMENTE: diagrama_unifilar e
      diario_de_obras = MERGED_PDF; os outros 5 = EMBEDDED_IMAGE
  T3  classificar_anexo() levanta AnexoNaoClassificadoError para um id
      desconhecido -- nunca inventa modo por posicao/numero
  T4  montar_plano() classifica cada anexo pelo `id`, NAO pela posicao dele
      na lista de entrada (embaralhar a lista de anexos nao muda a
      classificacao de nenhum -- prova que nao ha `if numero <= 3`)
  T5  montar_plano() separa main_document.embedded_annexes (so
      EMBEDDED_IMAGE) de external_pdf_annexes (so MERGED_PDF);
      merge_executed fica False
  T6  o plano serializa em JSON valido e recarrega com as mesmas chaves
  T7  extratores/anexos.montar() so devolve ids que `classificar_anexo()`
      aceita (nenhuma excecao) -- integra P0.2 com o extrator real

Execute:  python tests/test_assembly_model.py
"""
import os
import sys
import json
import random

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from gerador import assembly_model as AM  # noqa: E402
from gerador import config as C  # noqa: E402
from gerador.extratores import anexos as EX_ANEXOS  # noqa: E402

TIPOS_ESPERADOS = {
    "parecer_tecnico": "EMBEDDED_IMAGE",
    "acompanhamento_ambiental": "EMBEDDED_IMAGE",
    "acompanhamento_sst": "EMBEDDED_IMAGE",
    "nao_conformidade": "EMBEDDED_IMAGE",
    "ata_reuniao": "EMBEDDED_IMAGE",
    "diagrama_unifilar": "MERGED_PDF",
    "diario_de_obras": "MERGED_PDF",
}


class _SecaoFake:
    """Substituto minimo de document_model.Section -- so os 4 campos que
    assembly_model.montar_plano le (via _get, por atributo)."""
    def __init__(self, id, number, title, page_setup):
        self.id, self.number, self.title, self.page_setup = id, number, title, page_setup


def _anexos_de_exemplo():
    """Um anexo de cada um dos 7 tipos -- ordem PROPOSITALMENTE fora de
    qualquer ordem "natural" (nao crescente por posicao), para a T4 provar
    que a classificacao nao depende de posicao."""
    return [
        {"id": "diario_de_obras", "presente": True, "path": "d.pdf"},
        {"id": "nao_conformidade", "presente": True, "path": "nc.pdf"},
        {"id": "diagrama_unifilar", "presente": True, "path": "u.pdf"},
        {"id": "ata_reuniao", "presente": True, "path": "ata.pdf"},
        {"id": "parecer_tecnico", "presente": True, "path": "pa.pdf", "pa_numero": "PA-254"},
        {"id": "acompanhamento_sst", "presente": True, "path": "sst.pdf"},
        {"id": "acompanhamento_ambiental", "presente": True, "path": "amb.pdf"},
    ]


def main():
    falhas = []

    # ---------------------------------------------------------------- T1/T2
    tipos = C.annex_types().get("anexos", {})
    faltando = sorted(set(TIPOS_ESPERADOS) - set(tipos))
    if faltando:
        falhas.append("T1: annex_types.yaml sem os ids: %s" % faltando)
    else:
        print("T1  OK   os 7 tipos de anexo existem em config/annex_types.yaml")

    divergentes = []
    for anexo_id, modo_esperado in TIPOS_ESPERADOS.items():
        entrada = tipos.get(anexo_id)
        if not entrada:
            continue
        modo = entrada.get("integration_mode")
        if modo not in ("EMBEDDED_IMAGE", "MERGED_PDF"):
            divergentes.append("%s: integration_mode invalido %r" % (anexo_id, modo))
        elif modo != modo_esperado:
            divergentes.append("%s: %s, esperado %s" % (anexo_id, modo, modo_esperado))
    if divergentes:
        falhas.append("T2: " + "; ".join(divergentes))
    else:
        print("T2  OK   EMBEDDED_IMAGE x MERGED_PDF corretos para os 7 tipos")

    # ---------------------------------------------------------------- T3
    try:
        AM.classificar_anexo("anexo_que_nao_existe")
        falhas.append("T3: classificar_anexo nao levantou para id desconhecido")
    except AM.AnexoNaoClassificadoError:
        print("T3  OK   classificar_anexo levanta AnexoNaoClassificadoError para id desconhecido")

    # ---------------------------------------------------------------- T4
    base = _anexos_de_exemplo()
    embaralhado = list(base)
    random.Random(42).shuffle(embaralhado)

    def _classificacao(anexos):
        return {a["id"]: AM.classificar_anexo(a["id"])["integration_mode"] for a in anexos}

    c1, c2 = _classificacao(base), _classificacao(embaralhado)
    if c1 == c2 == TIPOS_ESPERADOS:
        print("T4  OK   classificacao identica independente da posicao/ordem na lista de anexos")
    else:
        falhas.append("T4: classificacao mudou com a ordem -- base=%s embaralhado=%s" % (c1, c2))

    # ---------------------------------------------------------------- T5
    sections = [
        _SecaoFake("capa", "", "Capa", "A4 retrato"),
        _SecaoFake("s5_2_fotografico", "5.2", "Relatório Fotográfico", "A3 retrato"),
        _SecaoFake("s11_anexos", "11", "ANEXOS", "A4 retrato"),
    ]
    plano = AM.montar_plano(
        report_id="RO_049", obra="RD103",
        periodo={"start": "2026-08-29", "end": "2026-09-04"},
        sections=sections, annexes=base,
        output_docx_path=os.path.join(U.PROJ, "tmp", "shadow_report_v2_teste.docx"),
    )
    ids_embedded = {a.anexo_id for a in plano.main_document.embedded_annexes}
    ids_external = {a.anexo_id for a in plano.external_pdf_annexes}
    esperado_embedded = {k for k, v in TIPOS_ESPERADOS.items() if v == "EMBEDDED_IMAGE"}
    esperado_external = {k for k, v in TIPOS_ESPERADOS.items() if v == "MERGED_PDF"}
    if (ids_embedded == esperado_embedded and ids_external == esperado_external
            and plano.merge_executed is False):
        print("T5  OK   embedded_annexes=%s external_pdf_annexes=%s merge_executed=False"
             % (sorted(ids_embedded), sorted(ids_external)))
    else:
        falhas.append("T5: embedded=%s (esperado %s) external=%s (esperado %s) merge_executed=%s"
                      % (ids_embedded, esperado_embedded, ids_external, esperado_external,
                         plano.merge_executed))

    # posicao penultimo/ultimo dos MERGED_PDF, na ordem certa
    ordem_external = [a.anexo_id for a in plano.external_pdf_annexes]
    if ordem_external == ["diagrama_unifilar", "diario_de_obras"]:
        print("T5b OK   external_pdf_annexes na ordem penultimo/ultimo: %s" % ordem_external)
    else:
        falhas.append("T5b: ordem external_pdf_annexes == %s, esperado "
                      "['diagrama_unifilar', 'diario_de_obras']" % ordem_external)

    # ---------------------------------------------------------------- T6
    saida = os.path.join(U.PROJ, "tmp", "assembly_plan_teste.json")
    caminho = AM.serializar(plano, saida)
    with open(caminho, encoding="utf-8") as fh:
        recarregado = json.load(fh)
    chaves_esperadas = {"report_id", "obra", "periodo", "main_document",
                        "external_pdf_annexes", "merge_executed", "merge_note", "gerado_em"}
    if chaves_esperadas.issubset(recarregado.keys()) and recarregado["merge_executed"] is False:
        print("T6  OK   plano serializado e recarregado em JSON: %s" % caminho)
    else:
        falhas.append("T6: JSON recarregado sem as chaves esperadas: %s" % sorted(recarregado.keys()))
    os.remove(caminho)

    # ---------------------------------------------------------------- T7
    try:
        anexos_reais = EX_ANEXOS.montar(
            meetings=[{"path": "ata.docx"}],
            ambiental_applies=True, fa_ambiental="amb.pdf", fa_sst="sst.pdf",
            nc_no_periodo=["NC-026"],
            unifilar_path="unifilar.pdf", unifilar_exists=True,
            diario_path="diario.pdf", diario_exists=True,
            parecer={"status": "MATCH", "mecanismo": "texto",
                    "documento": {"numero": "PA-254", "path": "pa254.pdf"}},
        )
        for a in anexos_reais:
            AM.classificar_anexo(a["id"])
        print("T7  OK   extratores.anexos.montar() devolve %d anexos, todos classificaveis: %s"
             % (len(anexos_reais), [a["id"] for a in anexos_reais]))
    except Exception as e:
        falhas.append("T7: %s: %s" % (type(e).__name__, e))

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
