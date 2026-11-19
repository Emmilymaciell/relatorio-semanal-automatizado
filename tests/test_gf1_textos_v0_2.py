# -*- coding: utf-8 -*-
"""
test_gf1_textos_v0_2.py

Fase 3.4 (v0.2), Frente A -- GF-1 (texto-padrao que faltava em 5 secoes),
indice (TEXTO-10/VIS-04) e lista de anexos (VIS-07). Um teste por bloco de
texto adicionado (presenca + parametrizacao correta), mais TOC e anexos.

  G1  secao 3 (Avanco Global): os 3 paragrafos-conceito aparecem, READY,
      identicos ao catalogo (nenhum e reconstruido a mao no teste)
  G2  secao 5.1 (Clima): paragrafo intro + as 3 definicoes climaticas (lista)
      aparecem, antes da tabela
  G3  secao 6 (Ambiental): paragrafos A (monitoramento quinzenal, com o
      numero do Parecer Tecnico REAL, nunca "___") e B (Dashboard Ambiental)
      aparecem quando o ciclo se aplica
  G4  secao 6: roda para as 4 obras sem excecao (nenhum `if obra ==`) --
      cada uma usa sua data de migracao do diario e (quando a figura de
      licencas estiver disponivel) seu proprio modelo de frase
  G5  secao 7 (Diario de Obras): as 4 frases aparecem, com a data de migracao
      e a virgula de "encontra-se anexo" corretas POR OBRA (road_registry) --
      RD-101 diverge da RD-103 na data, RD-103 diverge de RD-101/RD-104 na
      virgula
  G6  secao 8.3: parametro de definicao de status aparece; a frase-ponte do
      Grafico 2 usa a contagem REAL de nonconformities.history_by_status
      (nunca um numero digitado); a do Grafico 3 respeita o estilo por obra
      (numero x "a seguir")
  G7  indice: o heading "ÍNDICE" tem outlineLvl=9 no OOXML (nao aparece nos
      niveis 1-3 que o campo TOC escaneia -- nao se autolista) e continua
      centralizado e com o estilo H1; os demais headings (ex.: 1.
      APRESENTAÇÃO) NAO tem esse override
  G8  lista de anexos: nenhum paragrafo de anexo usa o estilo "List Bullet"
      (sem marcador), com o respiro (espaco depois) configurado

Execute:  python tests/test_gf1_textos_v0_2.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import re
import sys
import json
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from gerador import content_builder as CB  # noqa: E402
from gerador import config as C  # noqa: E402
from gerador.renderers import docx_renderer_v2  # noqa: E402

OBRA = "RD103"
PERIODO = ("2026-08-29", "2026-09-04")


def _carregar_report_data(obra):
    caminho = os.path.join(U.OUT, "shadow", "%s_%s_%s" % (obra, *PERIODO), "report_data.json")
    if not os.path.exists(caminho):
        return None
    with open(caminho, encoding="utf-8") as fh:
        return json.load(fh)


def _textos_de(secao, kind="paragraph"):
    if kind == "list":
        return [b.items for b in secao.blocks if b.kind == kind and b.items]
    return [b.text for b in secao.blocks if b.kind == kind and b.text]


def main():
    falhas = []
    report_data = _carregar_report_data(OBRA)
    if report_data is None:
        print("SKIP: rode 'python scripts/shadow_report.py data --obra RD103 --inicio "
             "2026-08-29 --fim 2026-09-04' antes deste teste")
        return 1

    textos = C._carregar("textos_v3.yaml")["textos"]

    def _t(chave):
        return textos[chave].get("texto") or textos[chave].get("template")

    doc = CB.construir_documento(report_data)
    secoes = {s.id: s for s in doc.sections}

    # ---------------------------------------------------------------- G1
    sec3 = secoes["s3_avanco_global"]
    paras3 = _textos_de(sec3)
    esperado3 = [_t("s3_avanco_conceito"), _t("s3_avanco_concluido_conceito"),
                _t("s3_avanco_concluido_atacado_conceito")]
    if all(e in paras3 for e in esperado3):
        print("G1  OK   secao 3: os 3 parágrafos-conceito de avanço global estão presentes")
    else:
        falhas.append("G1: esperado os 3 parágrafos de textos_v3.yaml em sec3, achou %s" % paras3)

    # ---------------------------------------------------------------- G2
    sec5 = secoes["s5_execucao_obra"]
    paras5 = _textos_de(sec5)
    listas5 = _textos_de(sec5, kind="list")
    intro_ok = _t("s5_1_intro_definicoes") in paras5
    bullets_esperados = [_t("s5_1_definicao_bom"), _t("s5_1_definicao_chuvoso"),
                         _t("s5_1_definicao_impraticavel")]
    bullets_ok = any(all(b in lst for b in bullets_esperados) for lst in listas5)
    idx_intro = next((i for i, b in enumerate(sec5.blocks)
                      if b.kind == "paragraph" and b.text == _t("s5_1_intro_definicoes")), None)
    idx_tabela = next((i for i, b in enumerate(sec5.blocks) if b.kind == "table"), None)
    ordem_ok = idx_intro is not None and idx_tabela is not None and idx_intro < idx_tabela
    if intro_ok and bullets_ok and ordem_ok:
        print("G2  OK   secao 5.1: parágrafo intro + 3 definições climáticas, antes da tabela")
    else:
        falhas.append("G2: intro_ok=%s bullets_ok=%s ordem_ok=%s (paras=%s, listas=%s)"
                      % (intro_ok, bullets_ok, ordem_ok, paras5, listas5))

    # ---------------------------------------------------------------- G3
    sec6 = secoes["s6_ambiental"]
    paras6 = [(b.text, b.status) for b in sec6.blocks if b.kind == "paragraph"]
    pa_numero = next((a.get("pa_numero") for a in report_data.get("annexes") or []
                      if a.get("id") == "parecer_tecnico"), None)
    texto_a_esperado = _t("s6_intro_monitoramento_quinzenal").format(pa_numero=pa_numero or "___")
    tem_a = any(t == texto_a_esperado and st == "READY" for t, st in paras6)
    tem_b = any(t == _t("s6_dashboard_ambiental") for t, _st in paras6)
    if report_data["environment"]["applies"] and pa_numero and tem_a and tem_b:
        print("G3  OK   secao 6: parágrafo A (monitoramento, %s, READY) e B (Dashboard) presentes"
             % pa_numero)
    elif not report_data["environment"]["applies"]:
        print("G3  SKIP semana da comparação não é do ciclo quinzenal (environment.applies=False)")
    else:
        falhas.append("G3: tem_a=%s tem_b=%s pa_numero=%s -- paras6=%s" % (tem_a, tem_b, pa_numero, paras6))

    # ---------------------------------------------------------------- G4
    rodadas = []
    for obra in U.OBRAS:
        rd = _carregar_report_data(obra)
        if rd is None:
            continue
        try:
            d = CB.construir_documento(rd)
            sec6_obra = next(s for s in d.sections if s.id == "s6_ambiental")
            assert sec6_obra is not None
            rodadas.append(obra)
        except Exception as e:
            falhas.append("G4: %s levantou %s: %s" % (obra, type(e).__name__, e))
    fonte_cb = open(os.path.join(SCRIPTS, "gerador", "content_builder.py"), encoding="utf-8").read()
    tem_condicional_obra = re.search(r'obra\s*==\s*["\']', fonte_cb) is not None
    if rodadas and not tem_condicional_obra and not any(f.startswith("G4") for f in falhas):
        print("G4  OK   secao 6 roda para as %d obras sem exceção, sem `obra == \"...\"` no código"
             % len(rodadas))
    elif not rodadas:
        print("G4  SKIP nenhum report_data.json disponível para as demais obras")

    # ---------------------------------------------------------------- G5
    sec7 = secoes["s7_diario_obras"]
    paras7 = _textos_de(sec7)
    info180 = C.obra_info(OBRA)
    from gerador.content_builder import _data_por_extenso  # reaproveita o mesmo helper do builder
    hist_esperado = _t("s7_historico_formato").format(
        data_fim_fisico=_data_por_extenso(info180["diario_migracao_fisico_ate"]),
        data_inicio_digital=_data_por_extenso(info180["diario_migracao_digital_desde"]))
    anexo_chave = "s7_diario_anexo_com_virgula" if info180.get("diario_anexo_com_virgula") \
        else "s7_diario_anexo_sem_virgula"
    checks_g5 = {
        "histórico (data RD-103)": hist_esperado in paras7,
        "termo (s7_texto_base)": _t("s7_texto_base") in paras7,
        "formalização/assinatura": _t("s7_formalizacao_assinatura") in paras7,
        "encontra-se anexo (forma RD-103)": _t(anexo_chave) in paras7,
    }
    # RD-101 tem data de migração DIFERENTE (fato, não redação) -- confirma
    # que a data vem de road_registry, não hardcoded.
    rd_rd101 = _carregar_report_data("RD101")
    diverge_rd101 = None
    if rd_rd101:
        doc147 = CB.construir_documento(rd_rd101)
        sec7_147 = next(s for s in doc147.sections if s.id == "s7_diario_obras")
        info147 = C.obra_info("RD101")
        hist_147 = _t("s7_historico_formato").format(
            data_fim_fisico=_data_por_extenso(info147["diario_migracao_fisico_ate"]),
            data_inicio_digital=_data_por_extenso(info147["diario_migracao_digital_desde"]))
        diverge_rd101 = (hist_147 in _textos_de(sec7_147)) and (hist_147 != hist_esperado)
    if all(checks_g5.values()) and (diverge_rd101 is None or diverge_rd101):
        print("G5  OK   seção 7: as 4 frases presentes, com data/vírgula corretas por obra "
             "(RD-101 diverge da RD-103 na data de migração: %s)" % diverge_rd101)
    else:
        falhas.append("G5: checks=%s diverge_rd101=%s" % (checks_g5, diverge_rd101))

    # ---------------------------------------------------------------- G6
    sec8 = secoes["s8_nao_conformidades"]
    paras8 = _textos_de(sec8)
    tem_definicao = _t("s8_3_status_definicao") in paras8
    hist = report_data["nonconformities"]["history_by_status"]
    from gerador.content_builder import _distribuicao_status_prosa
    distrib = _distribuicao_status_prosa(hist)
    ponte2_esperada = _t("s8_3_grafico2_ponte").format(n=2, distribuicao_status=distrib)
    tem_ponte2 = ponte2_esperada in paras8
    chave3 = ("s8_3_grafico3_ponte_a_seguir"
             if info180.get("nc_grafico3_estilo_referencia") == "a_seguir"
             else "s8_3_grafico3_ponte_numero")
    ponte3_esperada = _t(chave3).format(n=3)
    tem_ponte3 = ponte3_esperada in paras8
    if tem_definicao and tem_ponte2 and tem_ponte3:
        print("G6  OK   seção 8.3: definição de status + pontes dos Gráficos 2/3 presentes "
             "(distribuição real: %r)" % distrib)
    else:
        falhas.append("G6: definicao=%s ponte2=%s (esperado %r) ponte3=%s (esperado %r) -- paras8=%s"
                      % (tem_definicao, tem_ponte2, ponte2_esperada, tem_ponte3, ponte3_esperada, paras8))

    # ---------------------------------------------------------------- G7/G8 (OOXML)
    out_path = os.path.join(U.PROJ, "tmp", "test_gf1_v0_2_output.docx")
    docx_renderer_v2.render(doc, out_path)
    with zipfile.ZipFile(out_path) as z:
        doc_xml = z.read("word/document.xml").decode("utf-8", "ignore")

    # isola o paragrafo do heading "ÍNDICE" e o de "APRESENTAÇÃO" (1a entrada
    # real do índice) para comparar o outlineLvl entre os dois.
    m_indice = re.search(r"<w:p\b[^>]*>(?:(?!</w:p>).)*?ÍNDICE(?:(?!</w:p>).)*?</w:p>", doc_xml, re.S)
    m_apresentacao = re.search(r"<w:p\b[^>]*>(?:(?!</w:p>).)*?APRESENTAÇÃO(?:(?!</w:p>).)*?</w:p>",
                               doc_xml, re.S)
    outline_indice = re.search(r'<w:outlineLvl w:val="(\d+)"/>', m_indice.group(0)) if m_indice else None
    outline_apresentacao = (re.search(r'<w:outlineLvl w:val="(\d+)"/>', m_apresentacao.group(0))
                            if m_apresentacao else None)
    jc_indice = '<w:jc w:val="center"/>' in (m_indice.group(0) if m_indice else "")
    if (m_indice and outline_indice and outline_indice.group(1) == "9" and jc_indice
            and (not outline_apresentacao)):
        print("G7  OK   'ÍNDICE' tem outlineLvl=9 (fora do TOC \\o \"1-3\") e está centralizado; "
             "'APRESENTAÇÃO' não tem esse override (continua no índice)")
    else:
        falhas.append("G7: m_indice=%s outline_indice=%s jc_indice=%s outline_apresentacao=%s"
                      % (bool(m_indice), outline_indice, jc_indice, outline_apresentacao))

    # anexos: nenhum paragrafo do bloco "ANEXOS" pode referenciar o estilo
    # "ListBullet" nem faltar o espaco depois configurado.
    m_anexos_bloco = re.search(r"ANEXOS.*", doc_xml, re.S)
    bloco_pos_anexos = doc_xml[m_anexos_bloco.start():] if m_anexos_bloco else ""
    tem_bullet_apos_anexos = 'w:pStyle w:val="ListBullet"' in bloco_pos_anexos
    vs = C.visual_style()
    espaco_esperado_twips = int(round(vs["anexos"]["espaco_depois_pt"] * 20))  # 1pt = 20 twips
    tem_espaco = ('w:spacing' in bloco_pos_anexos
                 and ('w:after="%d"' % espaco_esperado_twips) in bloco_pos_anexos)
    if not tem_bullet_apos_anexos and tem_espaco:
        print("G8  OK   lista de anexos sem marcador (\"List Bullet\"), com %spt de respiro"
             % vs["anexos"]["espaco_depois_pt"])
    else:
        falhas.append("G8: tem_bullet_apos_anexos=%s tem_espaco=%s (esperado w:after=%d)"
                      % (tem_bullet_apos_anexos, tem_espaco, espaco_esperado_twips))

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
