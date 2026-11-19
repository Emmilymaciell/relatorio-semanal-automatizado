# -*- coding: utf-8 -*-
"""
test_ooxml_estrutura_v2.py

Testa scripts/gerador/renderers/docx_renderer_v2.py (Fase 3.2B, P1 --
design system global) lendo o `.docx` GERADO, via python-docx + OOXML cru
(zipfile). NAO depende do Word (nenhum destes testes abre o arquivo no
Word/COM -- so leitura de XML).

Os ~18 testes obrigatorios da spec (FASE3_2B_RENDERER_V2_SPEC.md, "Testes
obrigatorios"):

  T01  margens corretas por PERFIL de secao (cover_a4 25/30/25/30,
       body_a4/photo_a3 30/20/20/30)
  T02  compatibilityMode == 15
  T03  numero de secoes e papeis bate com o document_model (4 secoes fisicas:
       capa A4, corpo A4, fotografico A3, corpo A4)
  T04  titlePg=true na secao da capa (so nela)
  T05  todo header do documento tem exatamente 2 imagens (os 2 logos)
  T06  todo header do documento tem a linha horizontal (borda inferior de
       paragrafo -- P1.2)
  T07  o logo direito (2a imagem) fica DENTRO da largura util de CADA secao
       (tab stop calculado por usable_width, nunca coordenada fixa que
       ultrapasse a pagina -- e' o defeito do header2 do golden, P1.2)
  T08  todo footer do documento existe e tem exatamente 1 paragrafo com
       campo PAGE real (instrText "PAGE"), alinhado a direita
  T09  Arial em todo w:rFonts dos estilos que o documento efetivamente usa
       (Normal, Heading 1/2/3, Header, Footer, Caption, List Bullet,
       Table Grid, TOC1/2/3) -- P1.13
  T10  H1 (estilo "Heading 1") = 16pt
  T11  H1 (estilo "Heading 1") = cor 000000
  T12  H1/H2/H3 tem numPr no PROPRIO ESTILO (numeracao automatica ligada ao
       nivel, nao a paragrafo nem a texto) -- P1.7
  T13  o abstractNum referenciado tem lvlText multinivel ("%1.", "%1.%2.",
       "%1.%2.%3.") -- confirma "5.1."/"5.2.1." por NIVEL, nao string
  T14  indice tem um campo TOC real (instrText contendo "TOC")
  T15  estilos TOC1/TOC2/TOC3 tem os tab stops esquerdos nos 3 niveis
       (8,47 / 12,7 / 16,93 mm)
  T16  os mesmos estilos TOC1/2/3 tem o tab direito com leader="dot"
       (nunca caracteres "." digitados)
  T17  a tabela do clima (8 colunas, DATA/DIA/CLIMA) tem as 8 larguras reais
       (15/20/20/20/22,5/25/20/20 mm), bordas 0,5pt (sz=4) e shading F2F2F2
       na coluna de rotulos
  T18  nenhum titulo (Heading 1/2/3) comeca com numero escrito no texto
       (ex.: "1. APRESENTAÇÃO") -- a numeracao e' SEMPRE automatica

Execute:  python tests/test_ooxml_estrutura_v2.py
"""
import os as _os, sys as _sys  # guarda da versão pública
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _dados_reais import exigir_dados_reais  # noqa: E402
exigir_dados_reais(__file__)
import os
import re
import sys
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
import docx  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from gerador import content_builder as CB  # noqa: E402
from gerador.renderers import docx_renderer_v2  # noqa: E402

OBRA = "RD103"
PERIODO = ("2026-08-29", "2026-09-04")

_MM_POR_TWIP = 25.4 / 1440.0


def _mm(twips):
    return round(twips * _MM_POR_TWIP, 2)


def _perto(a, b, tol=0.5):
    return a is not None and abs(a - b) <= tol


def main():
    falhas = []

    caminho_data = os.path.join(U.OUT, "shadow", "%s_%s_%s" % (OBRA, *PERIODO), "report_data.json")
    if not os.path.exists(caminho_data):
        print("SKIP: rode 'python scripts/shadow_report.py data --obra RD103 --inicio "
             "2026-08-29 --fim 2026-09-04' antes deste teste")
        return 1
    import json
    with open(caminho_data, encoding="utf-8") as fh:
        report_data = json.load(fh)

    documento = CB.construir_documento(report_data)
    out_path = os.path.join(U.PROJ, "tmp", "test_ooxml_v2_output.docx")
    docx_renderer_v2.render(documento, out_path)

    try:
        d = docx.Document(out_path)
    except Exception as e:
        print("FALHA FATAL: .docx v2 nao abriu: %s: %s" % (type(e).__name__, e))
        return 1

    with zipfile.ZipFile(out_path) as z:
        names = z.namelist()
        settings_xml = z.read("word/settings.xml").decode("utf-8", "ignore")
        styles_xml = z.read("word/styles.xml").decode("utf-8", "ignore")
        numbering_xml = z.read("word/numbering.xml").decode("utf-8", "ignore")
        doc_xml = z.read("word/document.xml").decode("utf-8", "ignore")
        header_names = sorted(n for n in names if re.match(r"word/header\d+\.xml$", n))
        footer_names = sorted(n for n in names if re.match(r"word/footer\d+\.xml$", n))
        headers_xml = {n: z.read(n).decode("utf-8", "ignore") for n in header_names}
        footers_xml = {n: z.read(n).decode("utf-8", "ignore") for n in footer_names}

    # ---------------------------------------------------------------- T01
    esperado = {
        "cover": (25.0, 30.0, 25.0, 30.0),
        "body": (30.0, 20.0, 20.0, 30.0),
    }
    secoes_modelo = documento.sections
    perfis_fisicos = []
    ultimo = None
    for s in secoes_modelo:
        chave = (s.page_setup, getattr(s, "page_profile", "body"))
        if chave != ultimo:
            perfis_fisicos.append(chave)
            ultimo = chave
    margens_ok = True
    detalhe_t01 = []
    for i, sec in enumerate(d.sections):
        perfil = perfis_fisicos[i][1] if i < len(perfis_fisicos) else "body"
        alvo = esperado["cover"] if perfil == "cover" else esperado["body"]
        medidas = (round(sec.top_margin.mm, 1), round(sec.right_margin.mm, 1),
                  round(sec.bottom_margin.mm, 1), round(sec.left_margin.mm, 1))
        ok = all(_perto(m, a, 0.2) for m, a in zip(medidas, alvo))
        detalhe_t01.append((i, perfil, medidas, alvo, ok))
        margens_ok = margens_ok and ok
    if margens_ok:
        print("T01 OK   margens por perfil corretas em %d secoes: %s" % (len(d.sections), detalhe_t01))
    else:
        falhas.append("T01: margens incorretas: %s" % detalhe_t01)

    # ---------------------------------------------------------------- T02
    m = re.search(r'w:name="compatibilityMode"[^>]*w:val="(\d+)"', settings_xml)
    compat = int(m.group(1)) if m else None
    if compat == 15:
        print("T02 OK   compatibilityMode == 15")
    else:
        falhas.append("T02: compatibilityMode == %r, esperado 15" % compat)

    # ---------------------------------------------------------------- T03
    if len(d.sections) == len(perfis_fisicos):
        print("T03 OK   %d secoes fisicas no docx == %d perfis (page_setup,page_profile) do modelo"
             % (len(d.sections), len(perfis_fisicos)))
    else:
        falhas.append("T03: %d secoes no docx, esperado %d (%s)"
                      % (len(d.sections), len(perfis_fisicos), perfis_fisicos))

    # ---------------------------------------------------------------- T04
    titlepg_flags = [bool(s.different_first_page_header_footer) for s in d.sections]
    if titlepg_flags and titlepg_flags[0] and not any(titlepg_flags[1:]):
        print("T04 OK   titlePg=true só na 1a secao (capa): %s" % titlepg_flags)
    else:
        falhas.append("T04: titlePg por secao == %s, esperado [True, False, False, ...]" % titlepg_flags)

    # ---------------------------------------------------------------- T05
    n_imgs_por_header = {n: len(re.findall(r"<a:blip[ >]", x)) for n, x in headers_xml.items()}
    if headers_xml and all(n == 2 for n in n_imgs_por_header.values()):
        print("T05 OK   todo header tem exatamente 2 imagens: %s" % n_imgs_por_header)
    else:
        falhas.append("T05: imagens por header == %s, esperado 2 em cada" % n_imgs_por_header)

    # ---------------------------------------------------------------- T06
    linha_por_header = {n: ("pBdr" in x and "<w:bottom" in x) for n, x in headers_xml.items()}
    if headers_xml and all(linha_por_header.values()):
        print("T06 OK   todo header tem a linha horizontal (pBdr/bottom): %s" % linha_por_header)
    else:
        falhas.append("T06: linha horizontal ausente em algum header: %s" % linha_por_header)

    # ---------------------------------------------------------------- T07
    # tab stop do logo direito (w:tab w:pos) tem que ficar DENTRO da largura
    # util da secao correspondente -- nunca alem da pagina (defeito do
    # header2 do golden, P1.2). Cada header corresponde a 1 secao fisica
    # nao-capa; casamos por ORDEM de aparicao (header1=1a secao com header,
    # header2=2a, ...), que e' a mesma ordem em que _render_header() os cria.
    secoes_com_header = [s for s in d.sections if not s.different_first_page_header_footer]
    logo_dentro_da_pagina = True
    detalhe_t07 = []
    for (hn, hx), sec in zip(headers_xml.items(), secoes_com_header):
        usable_mm = round(sec.page_width.mm - sec.left_margin.mm - sec.right_margin.mm, 1)
        m_tab = re.search(r'<w:tab w:val="left" w:pos="(\d+)"/>', hx) or re.search(
            r'<w:tab w:pos="(\d+)" w:val="left"/>', hx)
        pos_mm = _mm(int(m_tab.group(1))) if m_tab else None
        ok = pos_mm is not None and 0 < pos_mm < usable_mm
        detalhe_t07.append((hn, pos_mm, usable_mm, ok))
        logo_dentro_da_pagina = logo_dentro_da_pagina and ok
    if headers_xml and logo_dentro_da_pagina:
        print("T07 OK   logo direito dentro da largura util em toda secao: %s" % detalhe_t07)
    else:
        falhas.append("T07: logo direito fora/indefinido em alguma secao: %s" % detalhe_t07)

    # ---------------------------------------------------------------- T08
    footers_ok = True
    detalhe_t08 = []
    for n, fx in footers_xml.items():
        tem_page = bool(re.search(r"<w:instrText[^>]*>[^<]*PAGE", fx))
        alinhado = 'w:jc w:val="right"' in fx
        detalhe_t08.append((n, tem_page, alinhado))
        footers_ok = footers_ok and tem_page and alinhado
    if footers_xml and footers_ok:
        print("T08 OK   todo footer tem campo PAGE real alinhado a direita: %s" % detalhe_t08)
    else:
        falhas.append("T08: footer sem campo PAGE/alinhamento correto: %s" % detalhe_t08)

    # ---------------------------------------------------------------- T09
    estilos_relevantes = ("Normal", "Heading1", "Heading2", "Heading3", "Header",
                          "Footer", "Caption", "ListBullet", "TableGrid",
                          "TOC1", "TOC2", "TOC3")
    nao_arial = []
    for styleId in estilos_relevantes:
        m_bloco = re.search(r'<w:style [^>]*w:styleId="%s"[^>]*>(.*?)</w:style>' % re.escape(styleId),
                            styles_xml, re.DOTALL)
        if not m_bloco:
            continue
        for m_font in re.finditer(r'<w:rFonts[^>]*w:ascii="([^"]+)"', m_bloco.group(1)):
            if m_font.group(1) != "Arial":
                nao_arial.append((styleId, m_font.group(1)))
    if not nao_arial:
        print("T09 OK   Arial em w:rFonts de todos os estilos usados (%d checados)" % len(estilos_relevantes))
    else:
        falhas.append("T09: rFonts != Arial em estilos usados: %s" % nao_arial)

    # ---------------------------------------------------------------- T10/T11
    h1 = d.styles["Heading 1"]
    tam_ok = h1.font.size is not None and h1.font.size.pt == 16
    cor_ok = h1.font.color is not None and h1.font.color.rgb is not None and str(h1.font.color.rgb) == "000000"
    if tam_ok:
        print("T10 OK   Heading 1 == 16pt")
    else:
        falhas.append("T10: Heading 1 tamanho == %s, esperado 16pt"
                      % (h1.font.size.pt if h1.font.size else None))
    if cor_ok:
        print("T11 OK   Heading 1 cor == 000000")
    else:
        falhas.append("T11: Heading 1 cor == %s, esperado 000000"
                      % (h1.font.color.rgb if h1.font.color else None))

    # ---------------------------------------------------------------- T12
    numpr_por_nivel = {}
    for nome in ("Heading1", "Heading2", "Heading3"):
        m_bloco = re.search(r'<w:style [^>]*w:styleId="%s"[^>]*>(.*?)</w:style>' % nome,
                            styles_xml, re.DOTALL)
        numpr_por_nivel[nome] = bool(m_bloco and "<w:numPr>" in m_bloco.group(1))
    if all(numpr_por_nivel.values()):
        print("T12 OK   numPr no PROPRIO ESTILO de Heading 1/2/3 (numeracao automatica ligada ao nivel)")
    else:
        falhas.append("T12: numPr ausente no estilo: %s" % numpr_por_nivel)

    # ---------------------------------------------------------------- T13
    tem_multinivel = bool(re.search(r'w:lvlText w:val="%1\.%2\.', numbering_xml)) and bool(
        re.search(r'w:lvlText w:val="%1\.%2\.%3\.', numbering_xml))
    if tem_multinivel:
        print("T13 OK   abstractNum com lvlText multinivel (%1.%2. e %1.%2.%3.)")
    else:
        falhas.append("T13: numbering.xml sem lvlText multinivel esperado")

    # ---------------------------------------------------------------- T14
    tem_toc_field = bool(re.search(r"<w:instrText[^>]*>[^<]*\bTOC\b", doc_xml))
    if tem_toc_field:
        print("T14 OK   campo TOC real presente no indice")
    else:
        falhas.append("T14: nenhum campo TOC (instrText) encontrado em document.xml")

    # ---------------------------------------------------------------- T15/T16
    niveis_esperados_mm = {"TOC1": 8.47, "TOC2": 12.7, "TOC3": 16.93}
    tabs_ok, leader_ok = {}, {}
    for styleId, mm_esq in niveis_esperados_mm.items():
        m_bloco = re.search(r'<w:style [^>]*w:styleId="%s"[^>]*>(.*?)</w:style>' % styleId,
                            styles_xml, re.DOTALL)
        bloco = m_bloco.group(1) if m_bloco else ""
        # atributos de <w:tab .../> podem vir em qualquer ordem -- extrai
        # cada w:tab inteiro e le w:pos/w:leader independente da posicao.
        tags_tab = re.findall(r"<w:tab\b[^/]*/>", bloco)
        pos_encontradas = []
        tem_leader_dot = False
        for tag in tags_tab:
            m_pos = re.search(r'w:pos="(\d+)"', tag)
            if m_pos:
                pos_encontradas.append(_mm(int(m_pos.group(1))))
            if re.search(r'w:leader="dot"', tag):
                tem_leader_dot = True
        tabs_ok[styleId] = any(_perto(p, mm_esq, 0.3) for p in pos_encontradas)
        leader_ok[styleId] = tem_leader_dot
    if all(tabs_ok.values()):
        print("T15 OK   tab esquerdo por nivel do indice: %s" % niveis_esperados_mm)
    else:
        falhas.append("T15: tab stop esquerdo incorreto: %s (tabs_ok=%s)" % (niveis_esperados_mm, tabs_ok))
    if all(leader_ok.values()):
        print("T16 OK   leader=\"dot\" no tab direito de TOC1/2/3 (nunca pontos digitados)")
    else:
        falhas.append("T16: leader=dot ausente em algum nivel do indice: %s" % leader_ok)

    # ---------------------------------------------------------------- T17
    larguras_esperadas = [15.0, 20.0, 20.0, 20.0, 22.5, 25.0, 20.0, 20.0]
    tabela_clima = None
    for t in d.tables:
        if len(t.columns) == 8 and len(t.rows) == 3:
            primeira_col = [t.cell(r, 0).text.strip().upper() for r in range(3)]
            if primeira_col and primeira_col[0].startswith("DATA"):
                tabela_clima = t
                break
    if tabela_clima is None:
        falhas.append("T17: tabela do clima (8 colunas, DATA/DIA/CLIMA) nao encontrada")
    else:
        larguras_mm = [round(c.width.mm, 1) if c.width else None for c in tabela_clima.columns]
        larguras_ok = all(_perto(w, e, 0.3) for w, e in zip(larguras_mm, larguras_esperadas))
        tc_xml = tabela_clima._tbl.xml
        bordas_sz = set(re.findall(r'<w:tcBorders>.*?w:sz="(\d+)"', tc_xml, re.DOTALL))
        shading = set(re.findall(r'w:fill="([0-9A-Fa-f]{6})"', tc_xml))
        bordas_ok = bordas_sz == {"4"}
        shading_ok = "F2F2F2" in shading
        if larguras_ok and bordas_ok and shading_ok:
            print("T17 OK   tabela do clima: larguras=%s bordas_sz=%s shading=%s"
                 % (larguras_mm, bordas_sz, shading))
        else:
            falhas.append("T17: tabela do clima incorreta -- larguras=%s (ok=%s) bordas_sz=%s (ok=%s) "
                          "shading=%s (ok=%s)" % (larguras_mm, larguras_ok, bordas_sz, bordas_ok,
                                                  shading, shading_ok))

    # ---------------------------------------------------------------- T18
    titulos_com_numero = []
    for p in d.paragraphs:
        if p.style.name.startswith("Heading"):
            texto = p.text.strip()
            texto_sem_marca = re.sub(r"^\[(HITL|LIMITAÇÃO DA FONTE|PENDENTE)\]\s*", "", texto)
            if re.match(r"^\d+[.)]", texto_sem_marca):
                titulos_com_numero.append(texto)
    if not titulos_com_numero:
        print("T18 OK   nenhum titulo com numero escrito no texto (numeracao e' sempre automatica)")
    else:
        falhas.append("T18: titulos com numero LITERAL no texto: %s" % titulos_com_numero)

    print("=" * 74)
    if falhas:
        print("FALHAS (%d):" % len(falhas))
        for f in falhas:
            print("  X", f)
        print("RESULTADO: FALHOU")
        return 1
    print("RESULTADO: OK (18/18)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
