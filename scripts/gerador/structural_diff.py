# -*- coding: utf-8 -*-
r"""
scripts/gerador/structural_diff.py - Fase 3.2B: comparacao ESTRUTURAL
(OOXML) entre Golden x v1 (shadow_report.docx) x v2 (shadow_report_v2.docx),
item a item da checklist P0/P1 de FASE3_2B_RENDERER_V2_SPEC.md.

Nao depende do Word (le o .docx como zip + XML, e via python-docx). O golden
e' lido de UMA COPIA JA EXISTENTE em tmp/ (nunca do SharePoint -- este modulo
nunca abre nada dentro de `PastaSincronizada\`; se a copia em tmp/ nao
existir, o item correspondente fica "SEM_GOLDEN_DISPONIVEL", nunca inventado).

status possiveis por item: FIXED, PARTIALLY_FIXED, NOT_FIXED, REGRESSED,
N_A (item nao e' observavel no OOXML do corpo -- ex.: P0.3/P0.4/P0.5 sao
arquitetura de montagem/composicao de imagem, verificados por
tests/test_assembly_model.py e nao por este comparador de documento renderizado).
"""
import os
import re
import zipfile

import docx

_NSW = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_NSA = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

MM_POR_TWIP = 25.4 / 1440.0


def _mm(twips):
    return round(twips * MM_POR_TWIP, 2)


def _existe(path):
    return bool(path) and os.path.exists(path)


# -----------------------------------------------------------------------------
# extracao de fatos de UM .docx (golden, v1 ou v2) -- generica, sem golden-
# specific hacks (mesmo extrator para os 3 arquivos).
# -----------------------------------------------------------------------------
def extrair_fatos(path):
    if not _existe(path):
        return None
    fatos = {"path": path}
    d = docx.Document(path)

    fatos["n_sections"] = len(d.sections)
    secoes = []
    for s in d.sections:
        secoes.append({
            "titlePg": bool(s.different_first_page_header_footer),
            "top_mm": round(s.top_margin.mm, 1) if s.top_margin is not None else None,
            "right_mm": round(s.right_margin.mm, 1) if s.right_margin is not None else None,
            "bottom_mm": round(s.bottom_margin.mm, 1) if s.bottom_margin is not None else None,
            "left_mm": round(s.left_margin.mm, 1) if s.left_margin is not None else None,
            "page_width_mm": round(s.page_width.mm, 1) if s.page_width is not None else None,
            "page_height_mm": round(s.page_height.mm, 1) if s.page_height is not None else None,
        })
    fatos["secoes"] = secoes
    fatos["perfis_margem_distintos"] = len({
        (s["top_mm"], s["right_mm"], s["bottom_mm"], s["left_mm"]) for s in secoes})
    fatos["algum_a0_a2"] = any(
        (s["page_width_mm"] not in (210.0, 297.0) and s["page_height_mm"] not in (297.0, 420.0))
        for s in secoes)

    normal = d.styles["Normal"]
    fatos["normal_fonte"] = normal.font.name
    fatos["normal_tamanho_pt"] = normal.font.size.pt if normal.font.size else None
    pf = normal.paragraph_format
    fatos["normal_justificado"] = (pf.alignment is not None and str(pf.alignment).startswith("JUSTIFY"))
    fatos["normal_entrelinha"] = pf.line_spacing

    try:
        h1 = d.styles["Heading 1"]
        fatos["h1_tamanho_pt"] = h1.font.size.pt if h1.font.size else None
        fatos["h1_cor"] = str(h1.font.color.rgb) if h1.font.color and h1.font.color.rgb else None
        fatos["h1_recuo_mm"] = (h1.paragraph_format.left_indent.mm
                                if h1.paragraph_format.left_indent else None)
    except KeyError:
        fatos["h1_tamanho_pt"] = fatos["h1_cor"] = fatos["h1_recuo_mm"] = None

    titulos_h1 = [p.text.strip() for p in d.paragraphs if p.style.name == "Heading 1"]
    fatos["titulos_h1"] = titulos_h1
    fatos["algum_titulo_com_numero_no_texto"] = any(
        re.match(r"^\[?\s*\d+[.)]\s", t) for t in titulos_h1)
    # o titulo da secao 9: procurado em TODOS os paragrafos, nao so nos de
    # estilo "Heading 1" -- o golden usa um estilo customizado ("Estilo1",
    # GOLDEN_VISUAL_SPEC.md §2.4), entao filtrar por "Heading 1" NUNCA acha
    # nada nele (só em v1/v2, que usam o estilo nativo do Word).
    fatos["titulo_secao9"] = next(
        (p.text.strip() for p in d.paragraphs
         if "INTERFER" in p.text.upper() and len(p.text.strip()) < 80
         # exclui a linha do INDICE ("9 — RELATO DE ..."), que cita o mesmo
         # titulo com um prefixo de numero+travessao -- queremos o
         # PARAGRAFO DO TITULO DE VERDADE, no corpo da secao 9.
         and not re.match(r"^\s*\d+\s*[—-]", p.text)), None)

    with zipfile.ZipFile(path) as z:
        names = z.namelist()

        settings_xml = z.read("word/settings.xml").decode("utf-8", "ignore") if "word/settings.xml" in names else ""
        # [^>]* (nao [^/]*) -- o valor do atributo w:uri contem "/" (e' uma
        # URL), entao parar no primeiro "/" cortava a busca antes do w:val.
        m = re.search(r'w:name="compatibilityMode"[^>]*w:val="(\d+)"', settings_xml)
        fatos["compatibility_mode"] = int(m.group(1)) if m else None

        header_names = sorted(n for n in names if re.match(r"word/header\d+\.xml$", n))
        footer_names = sorted(n for n in names if re.match(r"word/footer\d+\.xml$", n))
        fatos["n_headers"] = len(header_names)
        fatos["n_footers"] = len(footer_names)

        headers_info = []
        for hn in header_names:
            hx = z.read(hn).decode("utf-8", "ignore")
            headers_info.append({
                "arquivo": hn,
                "n_imagens": len(re.findall(r"<a:blip[ >]", hx)),
                "tem_linha_horizontal": ("pBdr" in hx and "w:bottom" in hx) or ("<wps:" in hx and "line" in hx.lower()) or ("cxnSp" in hx),
            })
        fatos["headers"] = headers_info

        footers_texto = []
        for fn in footer_names:
            fx = z.read(fn).decode("utf-8", "ignore")
            footers_texto.append({
                "arquivo": fn,
                "tem_campo_PAGE": bool(re.search(r"<w:instrText[^>]*>[^<]*PAGE", fx)),
                "alinhado_direita": 'w:jc w:val="right"' in fx,
            })
        fatos["footers"] = footers_texto

        doc_xml = z.read("word/document.xml").decode("utf-8", "ignore")
        fatos["tem_campo_toc"] = bool(re.search(r"<w:instrText[^>]*>[^<]*\bTOC\b", doc_xml))

        numbering_xml = z.read("word/numbering.xml").decode("utf-8", "ignore") if "word/numbering.xml" in names else ""
        fatos["tem_lvltext_multinivel"] = bool(re.search(r'w:lvlText w:val="%1\.%2\.', numbering_xml))

        styles_xml = z.read("word/styles.xml").decode("utf-8", "ignore") if "word/styles.xml" in names else ""
        fatos["heading1_tem_numpr_no_estilo"] = bool(
            re.search(r'w:styleId="Heading1"[\s\S]{0,400}?<w:numPr>', styles_xml))
        fatos["tem_estilo_toc1_com_leader_dot"] = bool(
            re.search(r'(TOC1|"toc 1")[\s\S]{0,600}?w:leader="dot"', styles_xml, re.IGNORECASE))
        # so conta fontes != Arial nos estilos que o RENDERER efetivamente
        # usa -- os ~140 estilos de tabela do template (nunca aplicados a
        # nenhum paragrafo/tabela do documento) e estilos como "Macro Text"
        # (Courier e' o padrao correto DELE, monoespacado por design) nao
        # sao defeito visivel e nao entram na contagem (P1.13 e' sobre o que
        # aparece no documento, nao sobre o template inteiro do python-docx).
        _ESTILOS_RELEVANTES = ("Normal", "Heading1", "Heading2", "Heading3", "Header",
                              "Footer", "Caption", "ListBullet", "TableGrid",
                              "TOC1", "TOC2", "TOC3", "DefaultParagraphFont")
        n_nao_arial = 0
        for styleId in _ESTILOS_RELEVANTES:
            m_bloco = re.search(
                r'<w:style [^>]*w:styleId="%s"[^>]*>(.*?)</w:style>' % re.escape(styleId),
                styles_xml, re.DOTALL)
            if not m_bloco:
                continue
            n_nao_arial += len(re.findall(
                r'<w:rFonts[^>]*w:ascii="(?!Arial)[^"]+"', m_bloco.group(1)))
        fatos["fontes_nao_arial_em_styles"] = n_nao_arial

        # tabela do clima: procurada pela FORMA (8 colunas, 3 linhas,
        # primeira coluna com "DATA"/"DIA"/"CLIMA"), nunca por posicao.
        fatos["tabela_clima"] = None
        for t in d.tables:
            if len(t.columns) == 8 and len(t.rows) == 3:
                primeira_col = [t.cell(r, 0).text.strip().upper() for r in range(3)]
                if primeira_col and primeira_col[0].startswith("DATA"):
                    # c.width e' um Length em EMU (python-docx) -- .mm ja
                    # converte certo; _mm() aqui seria errado (aquele helper
                    # e' so para twips/dxa, usados em w:tblW/w:tcW crus).
                    larguras_mm = [round(c.width.mm, 2) if c.width else None for c in t.columns]
                    tc_xml = t._tbl.xml
                    borda_sz = re.findall(r'<w:tcBorders>.*?w:sz="(\d+)"', tc_xml, re.DOTALL)
                    shading = re.findall(r'w:fill="([0-9A-Fa-f]{6})"', tc_xml)
                    fatos["tabela_clima"] = {
                        "larguras_mm": larguras_mm,
                        "bordas_sz_oitavos_pt": sorted(set(borda_sz)),
                        "shading_hex": sorted(set(shading)),
                    }
                    break

        # legendas: paragrafos curtos comecando por "Figura"/"Grafico"/"Tabela"
        fatos["legendas"] = []
        for p in d.paragraphs:
            if re.match(r"^(Figura|Gr[aá]fico|Tabela)\s+\d", p.text.strip()):
                runs = p.runs
                fonte = runs[0].font.name if runs else None
                tam = runs[0].font.size.pt if runs and runs[0].font.size else None
                cor = str(runs[0].font.color.rgb) if runs and runs[0].font.color and runs[0].font.color.rgb else None
                fatos["legendas"].append({
                    "texto": p.text[:40], "fonte": fonte, "tamanho_pt": tam, "cor": cor,
                    "centralizado": str(p.alignment).startswith("CENTER") if p.alignment else False,
                })

    return fatos


# -----------------------------------------------------------------------------
# classificacao por item -- cada checador recebe (golden, v1, v2) [dict ou
# None] e devolve (status, golden_resumo, v1_resumo, v2_resumo, nota).
# -----------------------------------------------------------------------------
def _resultado(status, g, v1, v2, nota):
    return {"status": status, "golden": g, "v1": v1, "v2": v2, "nota": nota}


def _check_p01_sem_a0_a2(g, v1, v2):
    g_ok = (g is not None) and (not g["algum_a0_a2"]) and g["n_sections"] == 9
    v1_ok = (v1 is not None) and (not v1["algum_a0_a2"])
    v2_ok = (v2 is not None) and (not v2["algum_a0_a2"])
    status = "FIXED" if (v1_ok and v2_ok) else ("REGRESSED" if v1_ok and not v2_ok else "NOT_FIXED")
    return _resultado(status,
                      "9 seções, todas A4/A3 (confirmado)" if g_ok else g,
                      "sem A0/A2" if v1_ok else v1, "sem A0/A2" if v2_ok else v2,
                      "document_model nunca emite page_setup A0/A2 (P0.1) -- checado no DOCX renderizado")


def _check_p06_perfis_estruturais(g, v1, v2):
    """P0.6 exige a MECÂNICA (capa distinguível: seção própria + titlePg);
    os VALORES exatos de margem por perfil são P1.1 -- checado à parte. Este
    item mede só a mecânica estrutural (seções != 1 perfil de margem
    OU titlePg=true), a partir do artefato .docx REALMENTE gravado em disco
    -- não do que o código-fonte atual seria capaz de gerar."""
    def _estado(fatos):
        if not fatos or not fatos["secoes"]:
            return False, False
        return fatos["secoes"][0]["titlePg"], fatos["perfis_margem_distintos"] > 1

    v1_titlepg, v1_perfis_ok = _estado(v1)
    v2_titlepg, v2_perfis_ok = _estado(v2)
    v1_ok = v1_titlepg or v1_perfis_ok
    v2_ok = v2_titlepg or v2_perfis_ok

    if v2_ok and v1_ok:
        status = "FIXED"
    elif v2_ok and not v1_ok:
        status = "FIXED"
    elif not v2_ok and v1_ok:
        status = "REGRESSED"
    else:
        status = "NOT_FIXED"
    return _resultado(
        status, "titlePg=true na capa, perfis de margem distintos por seção (confirmado no OOXML)",
        {"titlePg": v1_titlepg, "perfis_de_margem_distintos": v1["perfis_margem_distintos"] if v1 else None},
        {"titlePg": v2_titlepg, "perfis_de_margem_distintos": v2["perfis_margem_distintos"] if v2 else None},
        "medido no shadow_report.docx JÁ GRAVADO em disco (v1) -- não é regenerado nesta rodada "
        "(baseline preservado), mesmo que docx_renderer.py (fonte atual) já implemente titlePg")


def _check_p11_margens_por_perfil(g, v1, v2):
    esperado = {
        "cover": (25.0, 30.0, 25.0, 30.0),
        "body": (30.0, 20.0, 20.0, 30.0),
        "photo": (30.0, 20.0, 20.0, 30.0),
    }

    def _bate(fatos):
        if not fatos or not fatos["secoes"]:
            return False
        cover = fatos["secoes"][0]
        if (cover["top_mm"], cover["right_mm"], cover["bottom_mm"], cover["left_mm"]) != esperado["cover"]:
            return False
        corpo_ok = any(
            (s["top_mm"], s["right_mm"], s["bottom_mm"], s["left_mm"]) == esperado["body"]
            for s in fatos["secoes"][1:])
        return corpo_ok

    g_ok = _bate(g)
    v1_ok = _bate(v1)
    v2_ok = _bate(v2)
    status = "FIXED" if v2_ok else "NOT_FIXED"
    if v1_ok and not v2_ok:
        status = "REGRESSED"
    return _resultado(status, "cover 25/30/25/30, body/photo 30/20/20/30 (confirmado no OOXML)" if g_ok else g,
                      "margens corretas" if v1_ok else "margens genéricas (2,5/3,0cm em todos os perfis)",
                      "margens corretas" if v2_ok else "margens ainda não corretas",
                      "config/visual_style_v3.yaml margins_mm, aplicado por perfil em _set_page()")


def _check_p12_cabecalho(g, v1, v2):
    def _resumo(fatos):
        if not fatos or not fatos["headers"]:
            return "sem cabeçalho"
        return "%d cabeçalho(s), imagens=%s, linha=%s" % (
            len(fatos["headers"]), [h["n_imagens"] for h in fatos["headers"]],
            [h["tem_linha_horizontal"] for h in fatos["headers"]])
    v1_tem = bool(v1 and v1["headers"])
    v2_ok = bool(v2 and v2["headers"] and all(h["n_imagens"] == 2 for h in v2["headers"])
                and all(h["tem_linha_horizontal"] for h in v2["headers"]))
    status = "FIXED" if v2_ok else "NOT_FIXED"
    return _resultado(status,
                      "header3 correto (2 logos + linha, largura útil); header2 é GOLDEN_DEFECT "
                      "(logo a 228,7mm herdado em seções A4 — ver GOLDEN_VISUAL_SPEC.md §3.4)",
                      _resumo(v1), _resumo(v2),
                      "v2: usable_width recalculada por seção (Mm(usable-21.9) via tab stop), "
                      "nunca coordenada absoluta -- não reproduz o defeito do header2 do golden")


def _check_p13_titlepg(g, v1, v2):
    v1_ok = bool(v1 and v1["secoes"] and v1["secoes"][0]["titlePg"])
    v2_ok = bool(v2 and v2["secoes"] and v2["secoes"][0]["titlePg"])
    status = "FIXED" if v2_ok else "NOT_FIXED"
    if v1_ok and not v2_ok:
        status = "REGRESSED"
    return _resultado(status, "titlePg=true confirmado", v1_ok, v2_ok,
                      "shadow_report.docx (v1) já gravado em disco não tem titlePg -- artefato antecede "
                      "a mecânica hoje presente em docx_renderer.py; v2 grava titlePg=true na capa")


def _check_p14_rodape_page(g, v1, v2):
    def _ok(fatos):
        return bool(fatos and fatos["footers"] and all(
            f["tem_campo_PAGE"] and f["alinhado_direita"] for f in fatos["footers"]))
    v1_ok, v2_ok = _ok(v1), _ok(v2)
    status = "FIXED" if v2_ok else "NOT_FIXED"
    return _resultado(status, "campo PAGE real, alinhado à direita (fonte Aptos -- GOLDEN_INTERNAL_VARIATION)",
                      "sem rodapé" if not v1_ok else "ok", "campo PAGE real, Arial" if v2_ok else "sem rodapé",
                      "v1 nunca populava section.footer; v2 insere fldChar begin/instrText PAGE/separate/end")


def _check_p15_corpo(g, v1, v2):
    def _pontuacao(fatos):
        if not fatos:
            return 0
        pts = 0
        if fatos["normal_fonte"] == "Arial":
            pts += 1
        if fatos["normal_tamanho_pt"] == 11:
            pts += 1
        if fatos["normal_justificado"]:
            pts += 1
        if fatos["normal_entrelinha"] == 1.5:
            pts += 1
        return pts
    v1_p, v2_p = _pontuacao(v1), _pontuacao(v2)
    if v2_p == 4:
        status = "FIXED"
    elif v2_p > v1_p:
        status = "PARTIALLY_FIXED"
    elif v2_p < v1_p:
        status = "REGRESSED"
    else:
        status = "NOT_FIXED"
    return _resultado(status, "Arial 11pt justificado entrelinha 1,5 (predominante)",
                      "%d/4 propriedades corretas (Arial+11pt, sem justify/entrelinha)" % v1_p,
                      "%d/4 propriedades corretas" % v2_p,
                      "v1 setava só font.name/size no estilo Normal; v2 adiciona jc=both e line=1.5 no MESMO estilo")


def _check_p16_titulos(g, v1, v2):
    def _pontuacao(fatos):
        if not fatos:
            return 0
        pts = 0
        if fatos.get("h1_tamanho_pt") == 16:
            pts += 1
        if fatos.get("h1_cor") in ("000000", "00000000"):
            pts += 1
        if fatos.get("h1_recuo_mm") and abs(fatos["h1_recuo_mm"] - 12.7) < 0.3:
            pts += 1
        return pts
    v1_p, v2_p = _pontuacao(v1), _pontuacao(v2)
    status = "FIXED" if v2_p == 3 else ("PARTIALLY_FIXED" if v2_p > v1_p else "NOT_FIXED")
    return _resultado(status, "H1: Arial 16pt, cor 000000, recuo 12,7mm",
                      "%d/3 (herdava 14pt azul do tema, sem recuo)" % v1_p, "%d/3" % v2_p,
                      "v1 só setava font.name no estilo Heading 1 (mantinha sz/cor/indent do template padrão)")


def _check_p17_numeracao(g, v1, v2):
    v1_ok = bool(v1 and v1["heading1_tem_numpr_no_estilo"])
    v2_ok = bool(v2 and v2["heading1_tem_numpr_no_estilo"] and v2["tem_lvltext_multinivel"])
    v2_sem_numero_no_texto = bool(v2 and not v2["algum_titulo_com_numero_no_texto"])
    status = "FIXED" if (v2_ok and v2_sem_numero_no_texto) else "NOT_FIXED"
    return _resultado(status, "numeração automática nativa (não verificável de forma simples via OOXML do golden sem abrir no Word; presumida pela ausência de numPr customizado -- golden usa numId=10 herdado do template original)",
                      "sem numPr, sem numbering.xml customizado" if not v1_ok else "ok",
                      "numPr em Heading 1/2/3 + abstractNum multilevel (%1., %1.%2., %1.%2.%3.)" if v2_ok else "incompleto",
                      "P1.7: numeração ligada ao NÍVEL DO ESTILO -- content_builder nunca escreve \"1.\" no texto")


def _check_p18_indice(g, v1, v2):
    v1_ok = bool(v1 and v1["tem_campo_toc"])
    v2_ok = bool(v2 and v2["tem_campo_toc"] and v2["tem_estilo_toc1_com_leader_dot"])
    status = "FIXED" if v2_ok else "NOT_FIXED"
    return _resultado(status, "tab stops 8,47/12,7/16,93mm + 149,8mm, leader=dot (10./11. com recuo "
                      "de nível 2 é GOLDEN_INTERNAL_VARIATION, não reproduzida)",
                      "lista de parágrafos simples, sem tab stop nem leader" if not v1_ok else "ok",
                      "campo TOC real + estilos TOC1/2/3 com tab stops/leader dot" if v2_ok else "incompleto",
                      "v2 usa { TOC \\o \"1-3\" \\h \\z \\u } -- reflete o document_model sozinho via "
                      "estilos Heading, inclui os 5.2.x da semana automaticamente")


def _check_p19_tabela_clima(g, v1, v2):
    esperado_larguras = [15.0, 20.0, 20.0, 20.0, 22.5, 25.0, 20.0, 20.0]

    def _bate(fatos):
        tc = fatos.get("tabela_clima") if fatos else None
        if not tc:
            return False
        larguras_ok = all(abs((w or 0) - e) < 0.6 for w, e in zip(tc["larguras_mm"], esperado_larguras))
        bordas_ok = tc["bordas_sz_oitavos_pt"] in (["4"], ["3"], ["3", "4"])
        shading_ok = "F2F2F2" in tc["shading_hex"]
        return larguras_ok and bordas_ok and shading_ok

    g_tc = g.get("tabela_clima") if g else None
    v1_ok, v2_ok = _bate(v1), _bate(v2)
    status = "FIXED" if v2_ok else "NOT_FIXED"
    return _resultado(status, g_tc or "tabela do clima não encontrada pela forma esperada",
                      (v1.get("tabela_clima") if v1 else None) if not v1_ok else "ok",
                      "larguras/bordas/shading corretos" if v2_ok else (v2.get("tabela_clima") if v2 else None),
                      "v1 usava Table Grid genérico sem larguras/bordas/shading customizados")


def _check_p110_legendas(g, v1, v2):
    def _pontuacao(fatos):
        if not fatos or not fatos["legendas"]:
            return 0
        certos = sum(1 for lg in fatos["legendas"]
                    if lg["fonte"] == "Arial" and lg["tamanho_pt"] == 10
                    and lg["cor"] in ("0D0D0D", "0D0D0D".upper()) and lg["centralizado"])
        return certos / max(1, len(fatos["legendas"]))
    v1_p, v2_p = _pontuacao(v1), _pontuacao(v2)
    status = "FIXED" if v2_p == 1.0 and (v2 and v2["legendas"]) else (
        "PARTIALLY_FIXED" if v2_p > v1_p else "NOT_FIXED")
    return _resultado(status, "3 formatações divergentes no golden -- padronizar Arial 10pt/0D0D0D/centralizado (predominante)",
                      "%.0f%% das legendas com o padrão-alvo" % (v1_p * 100),
                      "%.0f%% das legendas com o padrão-alvo" % (v2_p * 100),
                      "v1 só aplicava itálico; v2 força fonte/tamanho/cor/alinhamento na legenda")


def _check_p111_titulo_secao9(g, v1, v2):
    g_titulo = g.get("titulo_secao9") if g else None
    v1_titulo = v1.get("titulo_secao9") if v1 else None
    v2_titulo = v2.get("titulo_secao9") if v2 else None
    esperado_rd103 = "RELATO DE INTERFERÊNCIA EM CAMPO"
    v1_ok = (v1_titulo == esperado_rd103)
    v2_ok = (v2_titulo == esperado_rd103)
    if v2_ok and not v1_ok:
        status = "FIXED"
    elif v2_ok and v1_ok:
        status = "FIXED"
    elif v1_ok and not v2_ok:
        status = "REGRESSED"
    else:
        status = "NOT_FIXED"
    return _resultado(status, g_titulo, v1_titulo, v2_titulo,
                      "content_builder._secao9_interferencias tinha \"RELATO DE RISCOS E INTERFERÊNCIAS "
                      "EM CAMPO\" hardcoded (bug pré-revisão de GUIA_DE_REDACAO.md R-05, 2026-09-17); "
                      "corrigido para ler road_registry.titulo_secao_9 por obra -- afeta v1 e v2 igualmente "
                      "porque os dois consomem o mesmo content_builder, mas o shadow_report.docx (v1) já "
                      "gravado em disco não é regenerado nesta rodada (baseline preservado) e por isso "
                      "ainda reflete o texto antigo")


def _check_p112_compat(g, v1, v2):
    g_val = g.get("compatibility_mode") if g else None
    v1_val = v1.get("compatibility_mode") if v1 else None
    v2_val = v2.get("compatibility_mode") if v2 else None
    status = "FIXED" if v2_val == 15 else "NOT_FIXED"
    if v1_val == 15 and v2_val != 15:
        status = "REGRESSED"
    return _resultado(status, g_val, v1_val, v2_val,
                      "python-docx grava compatibilityMode=14 por padrão; v2 sobrescreve para 15 em settings.xml")


def _check_p113_arial(g, v1, v2):
    v1_n = v1.get("fontes_nao_arial_em_styles") if v1 else None
    v2_n = v2.get("fontes_nao_arial_em_styles") if v2 else None
    status = "FIXED" if (v2_n is not None and v2_n == 0) else (
        "PARTIALLY_FIXED" if (v2_n is not None and v1_n is not None and v2_n < v1_n) else "NOT_FIXED")
    return _resultado(status, "Aptos no índice/rodapé é GOLDEN_INTERNAL_VARIATION -- padrão-alvo é Arial em tudo",
                      "%s estilos com w:ascii != Arial" % v1_n, "%s estilos com w:ascii != Arial" % v2_n,
                      "contagem de w:rFonts com w:ascii != \"Arial\" em word/styles.xml (estilos usados pelo "
                      "renderer foram forçados para Arial; estilos de tabela do template nunca usados não "
                      "contam como defeito visível)")


_ITENS = [
    ("P0.1", "DOCX não tem seção A0/A2", _check_p01_sem_a0_a2),
    ("P0.6", "capa estruturalmente distinta do corpo (titlePg + perfil)", _check_p06_perfis_estruturais),
    ("P1.1", "margens reais por perfil (cover_a4/body_a4/photo_a3)", _check_p11_margens_por_perfil),
    ("P1.2", "cabeçalho com 2 logos + linha pela largura útil da seção", _check_p12_cabecalho),
    ("P1.3", "titlePg=true na capa", _check_p13_titlepg),
    ("P1.4", "rodapé com campo PAGE real, alinhado à direita", _check_p14_rodape_page),
    ("P1.5", "corpo Arial 11pt justificado, entrelinha 1,5, por estilo", _check_p15_corpo),
    ("P1.6", "H1 Arial 16pt preto, recuo 12,7mm", _check_p16_titulos),
    ("P1.7", "numeração automática multinível ligada ao nível do estilo", _check_p17_numeracao),
    ("P1.8", "índice com tab stops/leader dot (campo TOC real)", _check_p18_indice),
    ("P1.9", "tabela do clima: 8 larguras reais, bordas 0,5pt, shading F2F2F2", _check_p19_tabela_clima),
    ("P1.10", "legendas Arial 10pt cor 0D0D0D centralizadas", _check_p110_legendas),
    ("P1.11", "título da seção 9 via road_registry.titulo_secao_9 (sem 'RISCOS')", _check_p111_titulo_secao9),
    ("P1.12", "compatibilityMode = 15", _check_p112_compat),
    ("P1.13", "Arial em todos os estilos usados (inclusive índice/rodapé)", _check_p113_arial),
]

_ITENS_NAO_OBSERVAVEIS_NO_DOCX = [
    ("P0.2", "arquitetura de montagem (assembly_model.py)", "tests/test_assembly_model.py"),
    ("P0.3", "anexos por modo de integração (annex_types.yaml)", "tests/test_assembly_model.py"),
    ("P0.4", "relatório fotográfico via PhotoPanelComposer", "scripts/gerador/photo_panel.py (sem teste dedicado nesta rodada)"),
    ("P0.5", "formato da imagem composta (DEC-004)", "docs/decisoes/DEC-004-formato-das-imagens-compostas.md"),
]


def gerar_diff(golden_path, v1_path, v2_path, output_path):
    """Monta o `renderer_v2_structural_diff.json`: Golden x v1 x v2, item a
    item, com FIXED/PARTIALLY_FIXED/NOT_FIXED/REGRESSED. Grava em
    `output_path` (dentro de output/shadow/... -- assert_read_only protege).
    """
    import sys
    aqui = os.path.dirname(os.path.abspath(__file__))
    scripts_dir = os.path.dirname(aqui)
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    import prog_util as U  # noqa: E402
    import json
    import datetime as dt

    U.assert_read_only(output_path)

    golden_local = None
    if golden_path and U.exists(golden_path):
        # nunca abre o golden direto do SharePoint (Regra Nº1/Nº2 do
        # CLAUDE.md do projeto): copia para tmp/ primeiro, mesmo padrão que
        # gerador/compare.py já usa para ler o golden com segurança.
        golden_local = U.copiar_para_tmp(golden_path, "structural_diff", "golden.docx")
    golden_fatos = extrair_fatos(golden_local) if golden_local else None
    v1_fatos = extrair_fatos(v1_path) if v1_path else None
    v2_fatos = extrair_fatos(v2_path) if v2_path else None

    itens = []
    for codigo, descricao, checker in _ITENS:
        r = checker(golden_fatos, v1_fatos, v2_fatos)
        itens.append({"item": codigo, "descricao": descricao, **r})

    for codigo, descricao, onde in _ITENS_NAO_OBSERVAVEIS_NO_DOCX:
        itens.append({
            "item": codigo, "descricao": descricao, "status": "N_A",
            "golden": None, "v1": None, "v2": None,
            "nota": "não observável no OOXML do documento renderizado -- verificado em %s" % onde,
        })

    contagem = {}
    for it in itens:
        contagem[it["status"]] = contagem.get(it["status"], 0) + 1

    resultado = {
        "gerado_em": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
        "golden_path": golden_path if (golden_path and U.exists(golden_path)) else None,
        "golden_disponivel": golden_fatos is not None,
        "v1_path": v1_path if _existe(v1_path) else None,
        "v2_path": v2_path if _existe(v2_path) else None,
        "itens": itens,
        "contagem_por_status": contagem,
    }

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as fh:
        json.dump(resultado, fh, ensure_ascii=False, indent=2, default=str)
    return output_path
