# -*- coding: utf-8 -*-
r"""
scripts/gerador/renderers/docx_renderer_v2.py - Fase 3.2B, P1: renderer com o
DESIGN SYSTEM GLOBAL (config/visual_style_v3.yaml), consumindo o MESMO
`gerador.document_model.Document` do renderer v1 (docx_renderer.py).

Diferenca para o v1: v1 e' estrutural (Slice 3) -- margem generica, sem
cabecalho/rodape, sem numeracao automatica, sem indice de verdade. v2
implementa P1.0 a P1.13 da spec (FASE3_2B_RENDERER_V2_SPEC.md):

  P1.0  fonte unica config/visual_style_v3.yaml -- nenhum numero magico aqui,
        todo mm/pt/cor vem de `gerador.config.visual_style()`.
  P1.1  margens por PERFIL de secao (cover_a4/body_a4/photo_a3).
  P1.2  cabecalho com 2 logos + linha, calculados pela LARGURA UTIL da
        propria secao (nunca coordenada absoluta) -- nao reproduz o
        GOLDEN_DEFECT do header2 do golden.
  P1.3  titlePg na capa.
  P1.4  rodape com campo PAGE real, alinhado a direita, Arial.
  P1.5  corpo Arial 11pt justificado, entrelinha 1,5, por ESTILO.
  P1.6  H1/H2/H3 com fonte/tamanho/cor/recuo por estilo.
  P1.7  numeracao automatica multinivel (numbering.xml ligado ao NIVEL do
        estilo -- nunca prefixo de string no texto).
  P1.8  indice com tab stops/leader dot -- campo TOC real do Word (atualiza
        sozinho ao abrir/imprimir; NAO copia paginacao do golden).
  P1.9  tabela do clima com as 8 larguras reais, bordas 0,5pt, shading.
  P1.10 legendas Arial 10pt cor 0D0D0D centralizadas.
  P1.11 titulo da secao 9 vem de road_registry.titulo_secao_9 (corrigido no
        content_builder -- este renderer so aplica o estilo do titulo).
  P1.12 compatibilityMode = 15.
  P1.13 Arial em tudo, inclusive indice e rodape.

So conhece `gerador.document_model.Document` + `gerador.config.visual_style()`
-- nunca report_data.json nem caminho de planilha (mesma regra do v1).

NAO SOBRESCREVE `shadow_report.docx` (v1) -- e' sempre gravado num caminho
`..._v2.docx` explicito, decidido por quem chama `render()`.
"""
import os
import sys

from docx import Document as DocxDocument
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_COLOR_INDEX, WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(os.path.dirname(AQUI))
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from .. import config as C  # noqa: E402

_TWIPS_POR_MM = 1440.0 / 25.4  # 1 polegada = 1440 twips = 25.4mm


def _twips(mm):
    return int(round(mm * _TWIPS_POR_MM))


def _rgb(hexstr):
    return RGBColor.from_string(hexstr)


_PAGE_DIMS_MM = None  # preenchido por _pagina_mm() a partir do visual_style


def _pagina_mm(vs, page_setup):
    chave = "a3_mm" if "A3" in (page_setup or "") else "a4_mm"
    d = vs["page"][chave]
    return d["width"], d["height"]


_PERFIL_MARGEM = {"cover": "cover_a4", "body": "body_a4", "photo": "photo_a3"}


def _margens_mm(vs, page_profile):
    chave = _PERFIL_MARGEM.get(page_profile, "body_a4")
    return vs["margins_mm"][chave]


def _usable_width_mm(vs, page_setup, page_profile):
    """P1.2: `usable_width = page_width - left_margin - right_margin`, DA
    PROPRIA SECAO -- nunca coordenada absoluta. E' a formula documentada em
    assets/logos/README.md; funciona sozinha em A4 (corpo, 160mm) e em A3
    (fotografico, 247mm) sem nenhum `if`."""
    largura, _altura = _pagina_mm(vs, page_setup)
    m = _margens_mm(vs, page_profile)
    return largura - m["left"] - m["right"]


_PREFIXOS = {
    "HITL": "[HITL] ",
    "LIMITATION": "[LIMITAÇÃO DA FONTE] ",
    "PLACEHOLDER": "[PENDENTE] ",
}


# -----------------------------------------------------------------------------
# P1.0/P1.5/P1.6/P1.13 -- estilos (Normal, Heading 1/2/3), aplicados por
# ESTILO (nao formatacao direta em cada paragrafo/run).
# -----------------------------------------------------------------------------
def _setup_estilos(doc, vs):
    fonte = vs["fonts"]["principal"]

    normal = doc.styles["Normal"]
    normal.font.name = fonte
    normal.font.size = Pt(vs["body"]["fonte_pt"])
    normal.font.color.rgb = _rgb(vs["body"]["cor"])
    pf = normal.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = vs["body"]["entrelinha"]
    pf.space_after = Pt(vs["body"]["espaco_depois_pt"])

    for nome_estilo, nivel_cfg in (("Heading 1", vs["h1"]), ("Heading 2", vs["h2"]),
                                   ("Heading 3", vs["h3"])):
        try:
            st = doc.styles[nome_estilo]
        except KeyError:
            continue
        st.font.name = fonte
        st.font.size = Pt(nivel_cfg["fonte_pt"])
        st.font.color.rgb = _rgb(nivel_cfg["cor"])
        st.paragraph_format.left_indent = Mm(nivel_cfg["recuo_esquerdo_mm"])
        st.paragraph_format.keep_with_next = bool(nivel_cfg["keep_next"])
        st.paragraph_format.keep_together = bool(nivel_cfg["keep_lines"])
        # rFonts tambem precisa ir no rPr do estilo explicitamente -- alguns
        # visualizadores (e o proprio Word em runs herdados de tema) só
        # respeitam w:rFonts quando ele existe no proprio elemento, nao so
        # font.name (que python-docx grava em rFonts ascii/hAnsi/cs de
        # qualquer forma, mas garantimos aqui tambem eastAsia por robustez).
        rPr = st.element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rFonts.set(qn(attr), fonte)

    # P1.13: "Arial em tudo" -- inclusive nos estilos auxiliares que o
    # renderer efetivamente usa (cabecalho, rodape, legenda, lista de
    # anexos, tabelas genericas). Nao forcamos os ~140 estilos de tabela
    # embutidos do template que o documento nunca chega a usar -- isso so
    # infllaria o XML sem mudar nada visivel.
    for nome_estilo in ("Header", "Footer", "List Bullet", "Table Grid"):
        try:
            st = doc.styles[nome_estilo]
        except KeyError:
            continue
        try:
            st.font.name = fonte
        except (AttributeError, TypeError):
            pass

    # Fase 3.3, Sec.3.2: o estilo "Caption" do template do python-docx vem
    # ORFAO (9pt azul 4F81BD, herdado, ninguem usa -- lacuna registrada no
    # Checkpoint 3.2B). Corrigido aqui para o padrao PREDOMINANTE do golden
    # (Arial 10pt cor 0D0D0D centralizado, config/visual_style_v3.yaml
    # `captions`) -- e as legendas (kind="caption" em _render_block) passam a
    # USAR ESTE ESTILO em vez de formatacao direta por paragrafo/run.
    try:
        cap = doc.styles["Caption"]
        cap.font.name = fonte
        cap.font.size = Pt(vs["captions"]["fonte_pt"])
        cap.font.color.rgb = _rgb(vs["captions"]["cor"])
        cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rPr = cap.element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rFonts.set(qn(attr), fonte)
    except KeyError:
        pass


# -----------------------------------------------------------------------------
# P1.12 -- compatibilityMode = 15 (Word 2013+), + updateFields para o campo
# TOC/PAGE se recalcular ao abrir (nao dependemos do Word COM para gerar o
# .docx -- so avisamos o Word para atualizar quando ELA abrir de verdade).
# -----------------------------------------------------------------------------
def _setup_compatibilidade(doc, vs):
    settings_el = doc.settings.element
    compat = settings_el.find(qn('w:compat'))
    alvo = str(vs["compatibility_mode"])
    if compat is not None:
        achou = False
        for cs in compat.findall(qn('w:compatSetting')):
            if cs.get(qn('w:name')) == 'compatibilityMode':
                cs.set(qn('w:val'), alvo)
                achou = True
        if not achou:
            cs = OxmlElement('w:compatSetting')
            cs.set(qn('w:name'), 'compatibilityMode')
            cs.set(qn('w:uri'), 'http://schemas.microsoft.com/office/word')
            cs.set(qn('w:val'), alvo)
            compat.insert(0, cs)
    upd = settings_el.find(qn('w:updateFields'))
    if upd is None:
        upd = OxmlElement('w:updateFields')
        settings_el.insert(0, upd)
    upd.set(qn('w:val'), 'true')


# -----------------------------------------------------------------------------
# P1.7 -- numeracao automatica multinivel, LIGADA AO NIVEL DO ESTILO (nunca
# prefixo de string no texto do titulo -- o content_builder ja garante isso).
# -----------------------------------------------------------------------------
def _setup_numeracao(doc, vs):
    numbering_part = doc.part.numbering_part
    numbering_el = numbering_part.element

    ids_abstract = [int(a.get(qn('w:abstractNumId')))
                    for a in numbering_el.findall(qn('w:abstractNum'))]
    novo_abstract_id = (max(ids_abstract) + 1) if ids_abstract else 100

    abstract = OxmlElement('w:abstractNum')
    abstract.set(qn('w:abstractNumId'), str(novo_abstract_id))
    mlt = OxmlElement('w:multiLevelType')
    mlt.set(qn('w:val'), 'multilevel')
    abstract.append(mlt)

    for nivel in vs["numbering"]["niveis"]:
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(nivel["ilvl"]))
        start = OxmlElement('w:start')
        start.set(qn('w:val'), '1')
        lvl.append(start)
        numFmt = OxmlElement('w:numFmt')
        numFmt.set(qn('w:val'), nivel["formato"])
        lvl.append(numFmt)
        lvlText = OxmlElement('w:lvlText')
        lvlText.set(qn('w:val'), nivel["texto"])
        lvl.append(lvlText)
        lvlJc = OxmlElement('w:lvlJc')
        lvlJc.set(qn('w:val'), 'left')
        lvl.append(lvlJc)
        pPr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        # recuo do NUMERO ligado ao mesmo nivel do titulo correspondente
        # (ilvl 0/1/2 == heading level 1/2/3) -- coerente com P1.6, sem
        # numero magico: reaproveita os proprios recuos de h1/h2/h3.
        chave_nivel = ("h1", "h2", "h3")[nivel["ilvl"]] if nivel["ilvl"] < 3 else "h3"
        ind.set(qn('w:left'), str(_twips(vs[chave_nivel]["recuo_esquerdo_mm"])))
        pPr.append(ind)
        lvl.append(pPr)
        abstract.append(lvl)

    primeiro_num = numbering_el.find(qn('w:num'))
    if primeiro_num is not None:
        primeiro_num.addprevious(abstract)
    else:
        numbering_el.append(abstract)

    ids_num = [int(n.get(qn('w:numId'))) for n in numbering_el.findall(qn('w:num'))]
    novo_num_id = (max(ids_num) + 1) if ids_num else 100
    num = OxmlElement('w:num')
    num.set(qn('w:numId'), str(novo_num_id))
    absid = OxmlElement('w:abstractNumId')
    absid.set(qn('w:val'), str(novo_abstract_id))
    num.append(absid)
    numbering_el.append(num)

    for nome_estilo, ilvl in (("Heading 1", 0), ("Heading 2", 1), ("Heading 3", 2)):
        try:
            st = doc.styles[nome_estilo]
        except KeyError:
            continue
        pPr = st.element.get_or_add_pPr()
        numPr = OxmlElement('w:numPr')
        ilvl_el = OxmlElement('w:ilvl')
        ilvl_el.set(qn('w:val'), str(ilvl))
        numPr.append(ilvl_el)
        numId_el = OxmlElement('w:numId')
        numId_el.set(qn('w:val'), str(novo_num_id))
        numPr.append(numId_el)
        pPr.append(numPr)

    return novo_num_id


def _cancelar_numeracao_paragrafo(paragraph):
    """Remove a numeracao automatica herdada do ESTILO so PARA ESTE
    PARAGRAFO (numId=0 e' a forma padrao OOXML de suprimir localmente) --
    usado no titulo "ÍNDICE", que nao entra na sequencia 1..11 (P1.8)."""
    pPr = paragraph._p.get_or_add_pPr()
    numPr = OxmlElement('w:numPr')
    ilvl_el = OxmlElement('w:ilvl')
    ilvl_el.set(qn('w:val'), '0')
    numPr.append(ilvl_el)
    numId_el = OxmlElement('w:numId')
    numId_el.set(qn('w:val'), '0')
    numPr.append(numId_el)
    pPr.append(numPr)


ESTILO_TITULO_INDICE = "TOC_TITLE"


def _excluir_do_toc(paragraph):
    """Marca o paragrafo como corpo de texto para efeito de outline level.

    Sozinho isto NAO basta -- ver `_setup_estilo_titulo_indice`. Continua
    aplicado porque cobre o switch `\\u` do campo (que coleta pelo outline
    level aplicado), enquanto o estilo proprio cobre o `\\o`.
    """
    pPr = paragraph._p.get_or_add_pPr()
    outlineLvl = OxmlElement('w:outlineLvl')
    outlineLvl.set(qn('w:val'), '9')
    pPr.append(outlineLvl)


def _setup_estilo_titulo_indice(doc, vs):
    """Fase 3.4 (v0.3), TEXTO-10: estilo proprio para o titulo "ÍNDICE".

    A v0.2 tentou resolver isto com `outlineLvl=9` direto no paragrafo, e
    NAO funcionou -- o campo continuou listando "ÍNDICE ... 2" como primeira
    entrada de si mesmo. O motivo: o campo e' `TOC \\o "1-3"`, e o switch
    `\\o` coleta pelos ESTILOS DE TITULO EMBUTIDOS (Heading 1/2/3), usando o
    outline level DO ESTILO. Um override de formatacao direta no paragrafo
    nao tira o paragrafo dessa varredura; ele so afeta o `\\u`. Enquanto o
    titulo usar o estilo "Heading 1", ele entra no indice.

    A correcao e' estrutural: um estilo que NAO e' Heading 1/2/3, derivado de
    Normal, visualmente igual ao titulo desejado (mesma fonte, corpo e cor do
    h1, centralizado conforme VIS-04), sem numeracao e com outlineLvl=9 no
    proprio estilo. Assim o `\\o` nao o alcanca (nao e' estilo de titulo) e o
    `\\u` tambem nao (outline level e' corpo de texto).

    O que NAO se faz aqui: apagar a linha do indice depois, por
    pos-processamento. Isso esconderia o sintoma e quebraria de novo assim
    que o Word recalculasse o campo.
    """
    try:
        st = doc.styles[ESTILO_TITULO_INDICE]
    except KeyError:
        st = doc.styles.add_style(ESTILO_TITULO_INDICE, WD_STYLE_TYPE.PARAGRAPH)
    st.base_style = doc.styles["Normal"]
    st.hidden = False
    st.quick_style = True

    fonte = vs["fonts"]["principal"]
    st.font.name = fonte
    st.font.size = Pt(vs["h1"]["fonte_pt"])
    st.font.color.rgb = _rgb(vs["h1"]["cor"])
    st.font.bold = True

    pf = st.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER      # VIS-04
    pf.left_indent = Mm(0)
    pf.keep_with_next = bool(vs["h1"]["keep_next"])
    pf.keep_together = bool(vs["h1"]["keep_lines"])

    rPr = st.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), fonte)

    # outlineLvl=9 NO ESTILO: fecha tambem o caminho do `\u`
    pPr_st = st.element.get_or_add_pPr()
    for antigo in pPr_st.findall(qn('w:outlineLvl')):
        pPr_st.remove(antigo)
    outline = OxmlElement('w:outlineLvl')
    outline.set(qn('w:val'), '9')
    pPr_st.append(outline)
    return st


# -----------------------------------------------------------------------------
# P1.8 -- estilos TOC1/TOC2/TOC3 (tab stops + leader dot) e o campo TOC real.
# -----------------------------------------------------------------------------
def _setup_estilos_indice(doc, vs):
    fonte = vs["fonts"]["principal"]
    toc_cfg = vs["toc"]
    tab_direito_mm = toc_cfg["tab_direito_mm"]
    leader = WD_TAB_LEADER.DOTS if toc_cfg["leader"] == "dot" else WD_TAB_LEADER.SPACES
    for nivel in toc_cfg["niveis"]:
        nome = "TOC %d" % nivel["nivel"]
        try:
            st = doc.styles[nome]
        except KeyError:
            st = doc.styles.add_style(nome, WD_STYLE_TYPE.PARAGRAPH)
        st.font.name = fonte
        st.font.size = Pt(toc_cfg["fonte_pt"])
        st.paragraph_format.space_after = Pt(toc_cfg["espaco_depois_pt"])
        # limpa tab stops previos (redefinicao idempotente) e recria os 2:
        # esquerdo (posicao do numero/titulo) e direito com leader de pontos
        # (posicao da pagina) -- nunca caracteres "." digitados (P1.8).
        ts = st.paragraph_format.tab_stops
        ts.clear_all()
        ts.add_tab_stop(Mm(nivel["tab_esquerdo_mm"]), WD_TAB_ALIGNMENT.LEFT)
        ts.add_tab_stop(Mm(tab_direito_mm), WD_TAB_ALIGNMENT.RIGHT, leader)


def _add_field(paragraph, instr, texto_cache):
    """Constroi um campo OOXML completo (begin/instrText/separate/cache/end)
    -- usado para TOC (P1.8) e PAGE (P1.4). `texto_cache` e' o que aparece
    ANTES do Word recalcular o campo (F9/impressao/abertura com
    updateFields=true) -- nunca a paginacao do golden copiada."""
    r1 = paragraph.add_run()
    b = OxmlElement('w:fldChar')
    b.set(qn('w:fldCharType'), 'begin')
    b.set(qn('w:dirty'), 'true')
    r1._r.append(b)

    r2 = paragraph.add_run()
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = instr
    r2._r.append(it)

    r3 = paragraph.add_run()
    s = OxmlElement('w:fldChar')
    s.set(qn('w:fldCharType'), 'separate')
    r3._r.append(s)

    r4 = paragraph.add_run(texto_cache)

    r5 = paragraph.add_run()
    e = OxmlElement('w:fldChar')
    e.set(qn('w:fldCharType'), 'end')
    r5._r.append(e)
    return r4


def _render_toc_field(doc, vs):
    """P1.8: campo `TOC` real do Word, \\o "1-3" (niveis Heading 1-3, onde
    mora a numeracao automatica de P1.7) \\h (hyperlink) \\z (oculta tab
    leader ao ver na web) \\u (usa niveis de estrutura). Reflete o
    document_model SOZINHO -- inclui os `5.2.x` da semana automaticamente
    porque eles sao Heading 3 de verdade, nao lista fixa. NUNCA copia a
    paginacao do golden: o numero de pagina so existe quando o Word
    recalcula o campo (por isso settings.xml tem updateFields=true)."""
    p = doc.add_paragraph()
    _add_field(
        p, r' TOC \o "1-3" \h \z \u ',
        "Atualize o índice: clique com o botão direito aqui dentro e "
        "escolha “Atualizar campo” (ou selecione tudo e pressione F9).")
    for r in p.runs:
        r.font.name = vs["fonts"]["principal"]
        r.italic = True


def _render_footer_page_field(section, vs):
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = _add_field(p, " PAGE ", "1")
    for run in p.runs:
        run.font.name = vs["fonts"]["principal"]
        run.font.size = Pt(vs["footer"]["fonte_pt"])


# -----------------------------------------------------------------------------
# P1.2 -- cabecalho: 2 logos + linha, pela largura util DA SECAO.
# -----------------------------------------------------------------------------
def _caminho_logo(nome_arquivo):
    return os.path.join(U.PROJ, nome_arquivo)


def _render_header(section, vs, page_setup, page_profile):
    usable_mm = _usable_width_mm(vs, page_setup, page_profile)
    hdr_cfg = vs["header"]
    logo_parceiro = hdr_cfg["logo_parceiro"]
    logo_prog = hdr_cfg["logo_prog"]

    p = section.header.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    # posicao do logo Programa = usable_width - largura_do_proprio_logo (formula
    # de assets/logos/README.md) -- NUNCA coordenada absoluta; recalculada
    # aqui a cada secao com a largura util QUE ELA MESMA TEM.
    x_logo_prog_mm = usable_mm - logo_prog["largura_mm"]
    p.paragraph_format.tab_stops.add_tab_stop(Mm(x_logo_prog_mm), WD_TAB_ALIGNMENT.LEFT)

    caminho_parceiro = _caminho_logo(logo_parceiro["arquivo"])
    caminho_prog = _caminho_logo(logo_prog["arquivo"])

    if os.path.exists(caminho_parceiro):
        r1 = p.add_run()
        r1.add_picture(caminho_parceiro, width=Mm(logo_parceiro["largura_mm"]), height=Mm(logo_parceiro["altura_mm"]))
    else:
        r1 = p.add_run("[LOGO Parceiro PENDENTE — ver assets/logos/README.md]")
        _marcar(r1, "LIMITATION")

    r_tab = p.add_run()
    r_tab.add_tab()

    if os.path.exists(caminho_prog):
        r2 = p.add_run()
        r2.add_picture(caminho_prog, width=Mm(logo_prog["largura_mm"]), height=Mm(logo_prog["altura_mm"]))
    else:
        r2 = p.add_run("[LOGO Programa PENDENTE — ver assets/logos/README.md]")
        _marcar(r2, "LIMITATION")

    # linha horizontal: borda inferior de um paragrafo em BRANCO logo abaixo
    # dos logos. Por nao ter recuo (ind left/right = 0), a borda do Word
    # cobre exatamente a largura util da secao (margem a margem) -- e' a
    # MESMA formula usable_width, so que o proprio motor de paragrafo do
    # Word aplica sozinho (nunca um comprimento fixo em EMU escrito aqui).
    p_linha = section.header.add_paragraph()
    p_linha.paragraph_format.space_before = Pt(2)
    p_linha.paragraph_format.space_after = Pt(0)
    pPr = p_linha._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    # 0,4pt do golden (header3, P1.2) => sz em oitavos de ponto (~3).
    bottom.set(qn('w:sz'), str(max(2, round(hdr_cfg["linha_espessura_pt"] * 8))))
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pbdr.append(bottom)
    pPr.append(pbdr)


def _marcar(run, status):
    if status == "READY":
        return
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    run.bold = True


# -----------------------------------------------------------------------------
# pagina / margens (P1.1) + titlePg (P1.3)
# -----------------------------------------------------------------------------
def _set_page(section, vs, page_setup, page_profile):
    largura_mm, altura_mm = _pagina_mm(vs, page_setup)
    section.page_width = Mm(largura_mm)
    section.page_height = Mm(altura_mm)
    m = _margens_mm(vs, page_profile)
    section.top_margin = Mm(m["top"])
    section.bottom_margin = Mm(m["bottom"])
    section.left_margin = Mm(m["left"])
    section.right_margin = Mm(m["right"])
    section.header_distance = Mm(vs["header_distance_mm"])
    section.footer_distance = Mm(vs["footer_distance_mm"])
    try:
        section.gutter = Mm(vs["gutter_mm"])
    except Exception:
        pass
    section.different_first_page_header_footer = (page_profile == "cover")


# -----------------------------------------------------------------------------
# P1.9 -- tabela do clima: 8 larguras reais, bordas 0,5pt, shading F2F2F2.
# -----------------------------------------------------------------------------
def _set_cell_width(cell, mm):
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is None:
        tcW = OxmlElement('w:tcW')
        tcPr.append(tcW)
    tcW.set(qn('w:w'), str(_twips(mm)))
    tcW.set(qn('w:type'), 'dxa')


def _set_cell_borders(cell, sz_oitavos_pt):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement('w:%s' % edge)
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(sz_oitavos_pt))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), 'auto')
        borders.append(el)
    tcPr.append(borders)


def _set_cell_shading(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)


def _render_tabela_clima(doc, block, vs):
    cfg = vs["tabela_clima"]
    larguras = cfg["larguras_mm"]
    sz = max(2, round(cfg["borda_pt"] * 8))
    rows = block.rows or []
    ncols = len(larguras)

    tbl = doc.add_table(rows=0, cols=ncols)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    tblPr = tbl._tbl.tblPr
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(sum(_twips(m) for m in larguras)))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)

    tblGrid = tbl._tbl.tblGrid
    for gc in list(tblGrid.findall(qn('w:gridCol'))):
        tblGrid.remove(gc)
    for m in larguras:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(_twips(m)))
        tblGrid.append(gc)

    for row in rows:
        cells = tbl.add_row().cells
        for i, v in enumerate(row):
            texto = "" if v is None else str(v)
            cells[i].text = texto
            for p in cells[i].paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.name = vs["fonts"]["principal"]
            _set_cell_width(cells[i], larguras[i])
            _set_cell_borders(cells[i], sz)
            if i == 0:
                _set_cell_shading(cells[i], cfg["shading_coluna_rotulo"])

    if block.status != "READY":
        p = doc.add_paragraph()
        r = p.add_run(_PREFIXOS.get(block.status, "") + (block.note or "conteúdo da tabela não comprovado"))
        _marcar(r, block.status)
    elif block.note:
        _nota(doc, block.note)


def _render_tabela_status(doc, block, vs):
    """Fase 3.3, Sec.3.3: tabela de status de NC com o MESMO tratamento
    visual da tabela do clima (P1.9) -- bordas 0,5pt, cabecalho sombreado
    F2F2F2, Arial, larguras coerentes com o conteudo (config/
    visual_style_v3.yaml `tabela_nc`). Reusa os MESMOS helpers de largura/
    borda/sombreado da tabela do clima -- nao um caminho novo."""
    cfg = vs["tabela_nc"]
    larguras = cfg["larguras_mm"]
    sz = max(2, round(cfg["borda_pt"] * 8))
    headers = block.headers or []
    rows = block.rows or []
    ncols = len(larguras)

    tbl = doc.add_table(rows=0, cols=ncols)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    tblPr = tbl._tbl.tblPr
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(sum(_twips(m) for m in larguras)))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)

    tblGrid = tbl._tbl.tblGrid
    for gc in list(tblGrid.findall(qn('w:gridCol'))):
        tblGrid.remove(gc)
    for m in larguras:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(_twips(m)))
        tblGrid.append(gc)

    if headers:
        cells = tbl.add_row().cells
        for i, h in enumerate(headers):
            cells[i].text = str(h)
            for p in cells[i].paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.bold = True
                    r.font.name = vs["fonts"]["principal"]
            _set_cell_width(cells[i], larguras[i])
            _set_cell_borders(cells[i], sz)
            _set_cell_shading(cells[i], cfg["shading_cabecalho"])

    for row in rows:
        cells = tbl.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = "" if v is None else str(v)
            for p in cells[i].paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.name = vs["fonts"]["principal"]
            _set_cell_width(cells[i], larguras[i])
            _set_cell_borders(cells[i], sz)

    if block.status != "READY":
        p = doc.add_paragraph()
        r = p.add_run(_PREFIXOS.get(block.status, "") + (block.note or "conteúdo da tabela não comprovado"))
        _marcar(r, block.status)
    elif block.note:
        _nota(doc, block.note)


def _render_table_generico(doc, block, vs):
    headers = block.headers or []
    rows = block.rows or []
    ncols = len(headers) if headers else (len(rows[0]) if rows else 1)
    tbl = doc.add_table(rows=0, cols=max(ncols, 1))
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    if headers:
        cells = tbl.add_row().cells
        for i, h in enumerate(headers):
            cells[i].text = str(h)
            for p in cells[i].paragraphs:
                for r in p.runs:
                    r.bold = True
                    r.font.name = vs["fonts"]["principal"]
    for row in rows:
        cells = tbl.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = "" if v is None else str(v)
            for p in cells[i].paragraphs:
                for r in p.runs:
                    r.font.name = vs["fonts"]["principal"]
    if block.status != "READY":
        p = doc.add_paragraph()
        r = p.add_run(_PREFIXOS.get(block.status, "") + (block.note or "conteúdo da tabela não comprovado"))
        _marcar(r, block.status)
    else:
        _nota(doc, block.note)


def _nota(doc, note):
    if not note:
        return
    p = doc.add_paragraph()
    r = p.add_run("Nota: %s" % note)
    r.italic = True
    r.font.size = Pt(9)


# -----------------------------------------------------------------------------
# renderizacao de blocos
# -----------------------------------------------------------------------------
def _render_block(doc, block, vs, secao):
    prefixo = _PREFIXOS.get(block.status, "") if block.status != "READY" else ""
    kind = block.kind

    if kind == "heading":
        if secao.id == "indice" and block.text == "ÍNDICE":
            # TEXTO-10 (v0.3): estilo proprio, NAO "Heading 1" -- so assim o
            # switch `\o` do campo TOC deixa de alcancar este paragrafo. Ver
            # _setup_estilo_titulo_indice() para o porque de outlineLvl=9
            # sozinho nao ter resolvido na v0.2.
            p = doc.add_paragraph(block.text or "", style=ESTILO_TITULO_INDICE)
            _cancelar_numeracao_paragrafo(p)   # P1.8: fora da sequencia 1..11
            _excluir_do_toc(p)                 # fecha tambem o caminho do `\u`
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER   # VIS-04
        else:
            p = doc.add_heading(block.text or "",
                                level=max(1, min(int(block.level or 1), 9)))
        if block.status != "READY" and p.runs:
            _marcar(p.runs[0], block.status)
            p.runs[0].text = prefixo + p.runs[0].text
        _nota(doc, block.note)

    elif kind == "paragraph":
        p = doc.add_paragraph()
        if block.role == "anexos_item":
            # Fase 3.4 (v0.2), CONFIG FIX VIS-07: paragrafo simples, SEM
            # marcador de lista, alinhado a margem, com respiro entre itens
            # (config/visual_style_v3.yaml `anexos`) -- nunca o estilo
            # "List Bullet" (o CURRENT nao usa "•" na lista de anexos).
            cfg_anexos = vs.get("anexos", {})
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(cfg_anexos.get("espaco_depois_pt", 28))
        r = p.add_run(prefixo + (block.text or ""))
        _marcar(r, block.status)
        _nota(doc, block.note)

    elif kind == "table":
        if block.role == "climate_table":
            _render_tabela_clima(doc, block, vs)
        elif block.role == "nc_status_table":
            _render_tabela_status(doc, block, vs)
        else:
            _render_table_generico(doc, block, vs)

    elif kind == "figure_placeholder":
        p = doc.add_paragraph()
        r = p.add_run(block.label or "[FIGURA — RENDERIZAÇÃO PENDENTE]")
        _marcar(r, block.status if block.status != "READY" else "PLACEHOLDER")
        _nota(doc, block.note)

    elif kind == "figure":
        # Fase 3.3, Sec.1: FiguraResult da Frente A ja convertida em imagem
        # em disco (figuras_bridge.py) -- insere centralizada no tamanho
        # (mm) que a propria figura devolveu, nunca um valor fixo aqui.
        if block.image_path and os.path.exists(block.image_path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run()
            r.add_picture(block.image_path, width=Mm(block.width_mm), height=Mm(block.height_mm))
            if block.status != "READY":
                _nota(doc, block.note)
        else:
            p = doc.add_paragraph()
            r = p.add_run(block.label or "[FIGURA — IMAGEM AUSENTE]")
            _marcar(r, block.status if block.status != "READY" else "PLACEHOLDER")
            _nota(doc, block.note)

    elif kind == "full_page_image":
        # Fase 3.3, Sec.4: 1 pagina de anexo EMBEDDED_IMAGE (fitz 200dpi),
        # em pagina cheia, herdando cabecalho/rodape/numeracao da SECAO em
        # que o bloco esta (nao abre secao fisica nova -- ver annex_embed.py).
        if block.image_path and os.path.exists(block.image_path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run()
            r.add_picture(block.image_path, width=Mm(block.width_mm), height=Mm(block.height_mm))
        else:
            p = doc.add_paragraph()
            r = p.add_run(block.label or "[ANEXO — IMAGEM AUSENTE]")
            _marcar(r, "LIMITATION")
            _nota(doc, block.note)

    elif kind == "caption":
        # Fase 3.3, Sec.3.2: usa o ESTILO "Caption" (corrigido em
        # _setup_estilos para Arial 10pt 0D0D0D centralizado) -- nunca mais
        # formatacao direta por paragrafo/run. Fecha a lacuna do estilo
        # orfao registrada no Checkpoint 3.2B.
        p = doc.add_paragraph(style="Caption")
        r = p.add_run(block.text or "")
        _marcar(r, block.status)

    elif kind == "list":
        for item in (block.items or []):
            p = doc.add_paragraph(style="List Bullet")
            r = p.add_run(prefixo + item)
            _marcar(r, block.status)
        _nota(doc, block.note)

    elif kind == "toc":
        # P1.8: campo TOC real -- ignora `block.items` (v1 usava lista fixa;
        # v2 deixa o proprio Word montar a partir dos estilos Heading, que
        # ja carregam a numeracao automatica de P1.7 -- inclui os 5.2.x da
        # semana sozinho, sem lista fixa e sem copiar paginacao do golden).
        _render_toc_field(doc, vs)

    elif kind == "page_break":
        doc.add_page_break()

    elif kind == "cover_field":
        p = doc.add_paragraph()
        r = p.add_run(block.text or "")
        _marcar(r, block.status)

    else:
        raise ValueError("document_model.Block.kind desconhecido para o docx_renderer_v2: %r" % kind)


# -----------------------------------------------------------------------------
# orquestracao
# -----------------------------------------------------------------------------
def render(document, output_path):
    """`document`: gerador.document_model.Document (o MESMO modelo do v1).
    Grava em `output_path` (sempre dentro de output/shadow/ -- assert_read_only
    protege). Devolve output_path. NUNCA sobrescreve shadow_report.docx (v1):
    quem chama decide o nome do arquivo (sempre "..._v2.docx" nesta fase)."""
    U.assert_read_only(output_path)

    vs = C.visual_style()
    doc = DocxDocument()
    _setup_estilos(doc, vs)
    _setup_estilos_indice(doc, vs)
    _setup_estilo_titulo_indice(doc, vs)   # TEXTO-10 (v0.3)
    _setup_compatibilidade(doc, vs)
    _setup_numeracao(doc, vs)

    primeira = True
    perfil_fisico_atual = None   # (page_setup, page_profile) da secao fisica atual
    perfil_header_atual = None   # perfil (body/photo) cujo header/footer esta ativo

    for secao in document.sections:
        perfil_secao = (secao.page_setup, getattr(secao, "page_profile", "body"))
        page_profile = perfil_secao[1]

        if primeira:
            docx_section = doc.sections[0]
            _set_page(docx_section, vs, *perfil_secao)
            perfil_fisico_atual = perfil_secao
            primeira = False
            nova_secao_fisica = True
        elif perfil_secao != perfil_fisico_atual:
            docx_section = doc.add_section(WD_SECTION.NEW_PAGE)
            _set_page(docx_section, vs, *perfil_secao)
            perfil_fisico_atual = perfil_secao
            nova_secao_fisica = True
        else:
            # Fase 3.4 (v0.2), item 4 do plano: politica de quebra de pagina
            # configuravel (config/visual_style_v3.yaml `paginacao.
            # sem_quebra_antes`) -- o golden AGRUPA algumas secoes na mesma
            # pagina fisica (4+5, 7+8, 9+10); por padrao (id fora da lista)
            # mantem o comportamento anterior (quebra sempre).
            politica_sem_quebra = vs.get("paginacao", {}).get("sem_quebra_antes", [])
            if secao.id not in politica_sem_quebra:
                doc.add_page_break()
            docx_section = None
            nova_secao_fisica = False

        if nova_secao_fisica:
            if page_profile == "cover":
                # P1.3: capa com titlePg -- sem cabecalho nem rodape (a
                # pagina de rosto usa o header/footer "first page", que fica
                # em branco por nao ser populado aqui).
                pass
            else:
                if page_profile != perfil_header_atual:
                    docx_section.header.is_linked_to_previous = False
                    docx_section.footer.is_linked_to_previous = False
                    _render_header(docx_section, vs, secao.page_setup, page_profile)
                    _render_footer_page_field(docx_section, vs)
                    perfil_header_atual = page_profile
                else:
                    docx_section.header.is_linked_to_previous = True
                    docx_section.footer.is_linked_to_previous = True

        for block in secao.blocks:
            _render_block(doc, block, vs, secao)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path
