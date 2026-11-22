# -*- coding: utf-8 -*-
r"""
scripts/gerador/pdf_final_assembly.py - Fase 3.3B: montagem final do PDF
completo (corpo + anexos externos MERGED_PDF) e patch do Termo de
Encerramento no PDF ja montado, sem depender de Word/Excel COM.

Este modulo faz TRES coisas, deliberadamente separadas:

  1. gerar_placeholder_unifilar()   -- pagina A0 paisagem, gerada com
     reportlab, para o lugar do Diagrama Unifilar quando a fonte real
     (UNIFILAR 103.xlsx) nao pode ser exportada neste ambiente (Excel COM
     nao funciona). NUNCA inventa conteudo de diagrama -- so avisa.

  2. montar_pdf() / contar_paginas() -- merge determinístico dos PDFs, na
     ORDEM recebida (nunca decide ordem aqui -- quem chama resolve pelo
     assembly_plan.json), via `fitz` (PyMuPDF). Nao duplica pagina: cada
     PDF de entrada entra inteiro, uma vez.

  3. patch_termo_pdf() -- como Word/Excel COM nao funcionam aqui, a
     reconversao DOCX->PDF apos atualizar o Termo (PASS 2 da Sec.5 da
     FASE3_3_FINAL_CANDIDATE_SPEC.md) nao pode ser executada pela maquina.
     Em vez de deixar o PDF final com o texto da passagem anterior (___),
     este modulo aplica um patch DIRETO no PDF ja montado: localiza a
     pagina do Termo pelo titulo, apaga (redaction) o paragrafo-placeholder
     e a nota de rodape, e reescreve o paragrafo final com o numero real de
     paginas. Isso NAO usa Word/Excel/COM -- so `fitz`, que ja esta
     disponivel. E' uma tecnica de PATCH DE PDF, nao um substituto do
     pipeline DOCX->PDF: fica registrado explicitamente no retorno de
     `patch_termo_pdf()` (campo `metodo`), para nunca ser confundido com
     uma reconversao real pelo Word.

Zero IA. Nada de `if obra == "..."`: todo texto/numero vem de quem chama
(sempre a partir de report_data/manifest, nunca hardcoded aqui).
"""
import os
import re
import shutil

import fitz  # PyMuPDF

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(AQUI)
import sys  # noqa: E402
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402

_PT_PER_MM = 72.0 / 25.4

# -----------------------------------------------------------------------------
# 1. placeholder do Diagrama Unifilar (A0 paisagem)
# -----------------------------------------------------------------------------
_A0_LARGURA_MM = 1189.0
_A0_ALTURA_MM = 841.0


def _quebrar_linhas(texto, fonte, tamanho, largura_max_pt):
    """Quebra `texto` em linhas que cabem em `largura_max_pt`, medindo com a
    metrica real da fonte (reportlab.pdfmetrics) -- nunca um numero fixo de
    caracteres por linha (comprimento varia com a fonte/acentuacao)."""
    from reportlab.pdfbase.pdfmetrics import stringWidth

    palavras = texto.split()
    linhas, atual = [], ""
    for p in palavras:
        candidata = ("%s %s" % (atual, p)).strip()
        if stringWidth(candidata, fonte, tamanho) <= largura_max_pt or not atual:
            atual = candidata
        else:
            linhas.append(atual)
            atual = p
    if atual:
        linhas.append(atual)
    return linhas


def gerar_placeholder_unifilar(output_path, *, motivo, fonte_esperada,
                                titulo="ANEXO IV — DIAGRAMA UNIFILAR — EXPORTAÇÃO PENDENTE",
                                largura_mm=_A0_LARGURA_MM, altura_mm=_A0_ALTURA_MM):
    """Gera 1 pagina paisagem (default A0, 1189 x 841 mm -- `largura_mm`/
    `altura_mm` aceitam qualquer papel, ex. A3 na RD-102A via
    road_registry.yaml#formato_unifilar) com aviso textual claro -- NUNCA
    desenha conteudo de diagrama. E' o FALLBACK de `gerar_anexo_unifilar()`
    (Fase 3.4/v0.2): so e' chamada quando a prancha REAL (UnifilarRenderer,
    le a propria planilha) nao pode ser gerada. Usa reportlab (ja
    disponivel, nao precisa instalar nada). Cada bloco de texto e' quebrado
    em linhas que cabem na largura util da pagina (`_quebrar_linhas`) --
    nunca deixa uma linha unica estourar a margem."""
    from reportlab.lib.colors import HexColor, black, white
    from reportlab.pdfgen import canvas

    U.assert_read_only(output_path)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    largura_pt = largura_mm * _PT_PER_MM
    altura_pt = altura_mm * _PT_PER_MM
    margem = 40
    largura_util = largura_pt - 2 * margem - 80  # respiro extra dentro da faixa

    fonte_titulo, tam_titulo = "Helvetica-Bold", 60
    fonte_corpo, tam_corpo = "Helvetica", 30
    fonte_nota, tam_nota = "Helvetica-Oblique", 22

    linhas_titulo = _quebrar_linhas(titulo, fonte_titulo, tam_titulo, largura_util)
    linhas_motivo = _quebrar_linhas(motivo, fonte_corpo, tam_corpo, largura_util)
    linhas_fonte = _quebrar_linhas("Fonte esperada: %s" % fonte_esperada,
                                   fonte_corpo, tam_corpo, largura_util)
    linhas_rodape = _quebrar_linhas(
        "Página gerada automaticamente para preservar a estrutura e a paginação do "
        "relatório. Nenhum conteúdo de diagrama foi inventado.", fonte_nota, tam_nota,
        largura_pt - 2 * margem)

    entrelinha_titulo = tam_titulo * 1.25
    entrelinha_corpo = tam_corpo * 1.4
    faixa_altura = (len(linhas_titulo) * entrelinha_titulo + 30
                    + len(linhas_motivo) * entrelinha_corpo
                    + len(linhas_fonte) * entrelinha_corpo + 40)

    c = canvas.Canvas(output_path, pagesize=(largura_pt, altura_pt))

    c.setStrokeColor(black)
    c.setLineWidth(2)
    c.rect(margem, margem, largura_pt - 2 * margem, altura_pt - 2 * margem)

    cy = altura_pt / 2.0
    c.setFillColor(HexColor("#B00020"))
    c.rect(margem, cy - faixa_altura / 2.0, largura_pt - 2 * margem, faixa_altura, fill=1, stroke=0)

    y = cy + faixa_altura / 2.0 - entrelinha_titulo
    c.setFillColor(white)
    c.setFont(fonte_titulo, tam_titulo)
    for linha in linhas_titulo:
        c.drawCentredString(largura_pt / 2.0, y, linha)
        y -= entrelinha_titulo

    y -= 20
    c.setFont(fonte_corpo, tam_corpo)
    for linha in linhas_motivo + linhas_fonte:
        c.drawCentredString(largura_pt / 2.0, y, linha)
        y -= entrelinha_corpo

    c.setFillColor(black)
    c.setFont(fonte_nota, tam_nota)
    y_rodape = margem + 40 + (len(linhas_rodape) - 1) * (tam_nota * 1.3)
    for linha in linhas_rodape:
        c.drawCentredString(largura_pt / 2.0, y_rodape, linha)
        y_rodape -= tam_nota * 1.3

    c.showPage()
    c.save()
    return output_path


# -----------------------------------------------------------------------------
# 2. merge determinístico
# -----------------------------------------------------------------------------
def contar_paginas(path):
    doc = fitz.open(U.lp(path) if U.exists(path) else path)
    try:
        return doc.page_count
    finally:
        doc.close()


def montar_pdf(paths_ordenados, output_path):
    """Concatena os PDFs de `paths_ordenados`, NA ORDEM dada, sem duplicar
    nenhuma pagina -- cada PDF de entrada entra inteiro, uma unica vez."""
    U.assert_read_only(output_path)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    saida = fitz.open()
    try:
        for p in paths_ordenados:
            origem = U.lp(p) if U.exists(p) else p
            src = fitz.open(origem)
            try:
                saida.insert_pdf(src)
            finally:
                src.close()
        saida.save(output_path, garbage=3, deflate=True)
    finally:
        saida.close()
    return output_path


# -----------------------------------------------------------------------------
# 3. patch do Termo de Encerramento dentro do PDF ja montado
# -----------------------------------------------------------------------------
_RE_HEADING_TERMO = re.compile(r"TERMO\s+DE\s+ENCERRAMENTO", re.IGNORECASE)
_RE_MARCADOR_PLACEHOLDER = re.compile(r"\[PENDENTE\]|_{2,}")


def localizar_pagina_termo(doc):
    """Acha a pagina que contem o paragrafo do Termo (heading + marcador de
    placeholder), nunca a pagina do INDICE que so cita o titulo. Devolve o
    indice 0-based ou None."""
    for i in range(doc.page_count):
        t = doc[i].get_text("text")
        if _RE_HEADING_TERMO.search(t) and _RE_MARCADOR_PLACEHOLDER.search(t):
            return i
    return None


def _redigir_area(page, rect, cor=(1, 1, 1)):
    page.add_redact_annot(rect, fill=cor)


# fragmentos que identificam os blocos do Termo-placeholder na pagina. Sao
# trechos do proprio texto emitido pelo content_builder, nao marcas visuais:
# o realce amarelo e' desenho, nao e' pesquisavel.
_FRAGMENTOS_PLACEHOLDER = (
    "apresenta-se",            # paragrafo do Termo
    "data de emissão só é conhecida",   # linha de fechamento
    "n_paginas_provisorio",    # nota de rodape do placeholder
)

# distancia maxima entre um bloco e o seguinte para trata-los como o mesmo
# trecho (entrelinha 1,5 de Arial 11 fica em ~19 pt; 25 pt da folga sem
# alcancar o proximo paragrafo de verdade)
_GAP_CONTINUACAO_PT = 25.0

_RE_TITULO_SECAO = re.compile(r"^\d+(\.\d+)*\.?\s")


def _localizar_blocos_placeholder(page):
    """Devolve a bbox que cobre o Termo-placeholder NESTA pagina.

    Ate a v0.1 a geometria do patch era fixa (`Rect(x, 118, x, 205)` e
    vizinhos). Funcionava por acidente: o Termo era a unica coisa da pagina,
    logo abaixo do titulo. Quando a v0.2 mudou a politica de quebra e as
    secoes 9 e 10 passaram a dividir a mesma pagina, a faixa 118..270 pt
    deixou de conter o Termo e passou a conter o CORPO DA SECAO 9 e o TITULO
    da secao 10 -- que foram apagados, enquanto o placeholder sobrevivia
    intacto mais abaixo. Coordenada fixa nao sobrevive a mudanca de layout.

    A area sai entao de duas ancoras encontradas na propria pagina:

      inicio = logo ABAIXO do titulo "10. TERMO DE ENCERRAMENTO"
      fim    = base do ultimo bloco do placeholder

    Ancorar no titulo (e nao nos blocos de texto do paragrafo) importa: o
    PyMuPDF quebra o paragrafo justificado em varios blocos, e casar por
    fragmento pegava so a ultima linha -- a bbox comecava no meio do
    paragrafo e as primeiras linhas sobreviviam a redacao.

    Devolve (bbox, blocos) ou (None, []) se as ancoras nao forem achadas.
    """
    blocos = [b for b in page.get_text("dict").get("blocks", []) if b.get("type") == 0]

    def _txt(b):
        return "".join(span.get("text", "")
                       for linha in b.get("lines", [])
                       for span in linha.get("spans", []))

    rect_titulo = None
    for b in blocos:
        if _RE_HEADING_TERMO.search(_txt(b)):
            rect_titulo = fitz.Rect(b["bbox"])
            break
    if rect_titulo is None:
        return None, []

    # blocos do placeholder = os que estao ABAIXO do titulo e trazem uma das
    # marcas do texto provisorio
    achados = []
    for b in blocos:
        r = fitz.Rect(b["bbox"])
        if r.y0 < rect_titulo.y1:
            continue
        t = _txt(b)
        if any(f in t for f in _FRAGMENTOS_PLACEHOLDER):
            achados.append((r, t))
    if not achados:
        return None, []

    x0 = min(r.x0 for r, _ in achados)
    x1 = max(r.x1 for r, _ in achados)
    y1 = max(r.y1 for r, _ in achados)

    # o PyMuPDF quebra a nota de rodape em varios blocos e so o primeiro
    # carrega a marca ("n_paginas_provisorio"); sem absorver as continuacoes,
    # sobra lixo do placeholder na pagina ("copiado do golden)"). Entao a
    # area cresce para baixo enquanto o proximo bloco estiver colado (ate
    # _GAP_CONTINUACAO_PT) e nao for um titulo de secao -- um titulo marca o
    # inicio de outro conteudo, que NAO pode ser apagado.
    outros = sorted(((fitz.Rect(b["bbox"]), _txt(b)) for b in blocos),
                    key=lambda par: par[0].y0)
    mudou = True
    while mudou:
        mudou = False
        for r, t in outros:
            if r.y0 < y1 or r.y0 - y1 > _GAP_CONTINUACAO_PT:
                continue
            if _RE_TITULO_SECAO.match(t.strip()):
                continue
            y1 = max(y1, r.y1)
            x0, x1 = min(x0, r.x0), max(x1, r.x1)
            achados.append((r, t))
            mudou = True

    bbox = fitz.Rect(x0, rect_titulo.y1 + 1, x1, y1)
    return bbox, achados


_MAPA_LATIN1 = {"–": "-", "—": "-", "‘": "'", "’": "'",
               "“": '"', "”": '"', "…": "..."}


def _sanitizar_latin1(texto):
    """A fonte base-14 usada no patch (`helv`) só cobre Latin-1 (não
    WinAnsi/cp1252): travessão/aspas tipográficas fora do Latin-1 saem como
    '?' na página. Troca por equivalentes ASCII antes de inserir -- é uma
    perda tipográfica pequena e documentada (ver `nota` do retorno),
    NUNCA um problema de conteúdo (os caracteres acentuados normais,
    ç/ã/á/é/í/ó/ú/º/°, estão dentro do Latin-1 e não são afetados)."""
    for k, v in _MAPA_LATIN1.items():
        texto = texto.replace(k, v)
    return texto


def patch_termo_pdf(input_path, output_path, *, texto_paragrafo, texto_fechamento,
                     fontsize=11, fontname="helv"):
    """Reescreve, DENTRO do PDF ja montado, o paragrafo do Termo de
    Encerramento -- sem Word/Excel/COM. Apaga (redaction) a area do
    paragrafo-placeholder (texto + destaque amarelo) e a nota de rodape, e
    insere o texto final (sem prefixo [PENDENTE], sem negrito, sem
    destaque), no mesmo bloco/pagina.

    Devolve um dict com o resultado -- NUNCA finge sucesso: se o texto nao
    couber na caixa (`insert_textbox` devolve sobra negativa), `ok=False` e
    o PDF de entrada e' copiado sem alteracao para `output_path`."""
    U.assert_read_only(output_path)
    doc = fitz.open(input_path)
    try:
        idx = localizar_pagina_termo(doc)
        if idx is None:
            doc.close()
            shutil.copyfile(input_path, output_path)
            return {"ok": False, "motivo": "pagina do Termo não localizada no PDF montado",
                    "metodo": "patch_pdf_direto_sem_com", "pagina": None}

        page = doc[idx]
        texto_paragrafo = _sanitizar_latin1(texto_paragrafo)
        texto_fechamento = _sanitizar_latin1(texto_fechamento)

        # -----------------------------------------------------------------
        # geometria: sai do texto REALMENTE encontrado na pagina, nunca de
        # coordenada fixa -- ver _localizar_blocos_placeholder() para o
        # motivo (regressao da v0.2, em que o Termo deixou de ser o unico
        # conteudo da sua pagina).
        # -----------------------------------------------------------------
        bbox, blocos = _localizar_blocos_placeholder(page)
        if bbox is None:
            doc.close()
            shutil.copyfile(input_path, output_path)
            return {"ok": False, "metodo": "patch_pdf_direto_sem_com", "pagina": idx + 1,
                    "motivo": "blocos do Termo-placeholder não localizados na página %d "
                              "-- PDF NÃO alterado" % (idx + 1)}

        # margens horizontais reais do corpo A4, herdadas do layout do
        # renderer; a bbox medida e' a referencia, com uma folga minima para
        # o realce amarelo, que e' desenho e extrapola um pouco o glifo.
        # folga lateral e inferior para o realce amarelo (desenho, extrapola
        # o glifo). No TOPO nao ha folga: bbox.y0 ja vem 1 pt abaixo do
        # titulo da secao, e recuar aqui apagaria o proprio titulo.
        x0, x1 = bbox.x0 - 3, bbox.x1 + 3
        y0 = bbox.y0
        y1 = bbox.y1 + 3

        _redigir_area(page, fitz.Rect(x0, y0, x1, y1))
        page.apply_redactions()

        # o paragrafo ocupa o topo da area liberada; o fechamento
        # ("Cidade, <data>.") vem depois de uma linha em branco, como no
        # layout do Word.
        altura_total = y1 - y0
        altura_fechamento = fontsize * 2.2
        folga = fontsize * 1.6
        altura_paragrafo = altura_total - altura_fechamento - folga
        if altura_paragrafo <= fontsize:
            doc.close()
            shutil.copyfile(input_path, output_path)
            return {"ok": False, "metodo": "patch_pdf_direto_sem_com", "pagina": idx + 1,
                    "motivo": "área liberada (%.1f pt) não comporta parágrafo + fechamento "
                              "-- PDF NÃO alterado" % altura_total}

        rect_paragrafo = fitz.Rect(x0, y0, x1, y0 + altura_paragrafo)
        rect_fechamento = fitz.Rect(x0, y0 + altura_paragrafo + folga, x1, y1)

        sobra_par = page.insert_textbox(
            rect_paragrafo, texto_paragrafo, fontsize=fontsize, fontname=fontname,
            color=(0, 0, 0), align=fitz.TEXT_ALIGN_JUSTIFY)
        sobra_fech = page.insert_textbox(
            rect_fechamento, texto_fechamento, fontsize=fontsize, fontname=fontname,
            color=(0, 0, 0), align=fitz.TEXT_ALIGN_LEFT)

        coube = sobra_par >= 0 and sobra_fech >= 0
        if not coube:
            doc.close()
            shutil.copyfile(input_path, output_path)
            return {"ok": False, "metodo": "patch_pdf_direto_sem_com", "pagina": idx + 1,
                    "motivo": "texto final não coube na área do parágrafo original "
                              "(sobra_paragrafo=%.1f, sobra_fechamento=%.1f) -- PDF NÃO alterado"
                              % (sobra_par, sobra_fech)}

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path, garbage=3, deflate=True)
        return {"ok": True, "metodo": "patch_pdf_direto_sem_com", "pagina": idx + 1,
                "fonte_usada": fontname, "sobra_paragrafo_pt": round(sobra_par, 2),
                "sobra_fechamento_pt": round(sobra_fech, 2),
                "nota": "paragrafo do Termo reescrito diretamente no PDF (fitz/redaction), "
                        "SEM Word/Excel/COM; fonte 'helv' (Helvetica, metricamente compativel "
                        "com Arial) -- nao e' a mesma fonte incorporada pelo Word. Travessao "
                        "(–/—) trocado por hifen simples (fonte base-14 so cobre Latin-1); "
                        "acentuacao normal do portugues nao e' afetada."}
    finally:
        try:
            doc.close()
        except Exception:
            pass


# -----------------------------------------------------------------------------
# 4. Fase 3.4/v0.2 -- integracao do UnifilarRenderer (prancha REAL) na
#    montagem, com fallback automatico para o placeholder textual (item 1)
# -----------------------------------------------------------------------------
# rotulo de papel (road_registry.yaml#formato_unifilar, via assembly_plan.json)
# -> dimensoes ISO 216 em mm. So mapeia o ROTULO para dimensao -- generico,
# nenhuma logica aqui decide por nome de obra (RD102A cai em "a3 paisagem"
# porque e' o que a config dela diz, nao por um `if obra == "RD102A"`).
_PAPEL_PDF_MM = {
    "a0 paisagem": (_A0_LARGURA_MM, _A0_ALTURA_MM),
    "a3 paisagem": (420.0, 297.0),
}


def _dims_papel_pdf_mm(papel_pdf):
    """`papel_pdf` (ex.: 'A0 paisagem', 'A3 paisagem') -> (largura_mm,
    altura_mm). Rotulo nao reconhecido cai em A0 paisagem (o formato das 3
    obras que nao sao excecao) -- nunca lanca excecao por um rotulo
    inesperado de papel."""
    return _PAPEL_PDF_MM.get((papel_pdf or "").strip().lower(),
                             _PAPEL_PDF_MM["a0 paisagem"])


def _png_para_pdf_pagina(png_path, output_path, largura_mm, altura_mm):
    """Embute um PNG (prancha ja renderizada) como pagina unica de PDF, do
    tamanho de papel dado -- via `fitz` (ja disponivel, sem Word/Excel/
    LibreOffice/COM)."""
    U.assert_read_only(output_path)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    largura_pt = largura_mm * _PT_PER_MM
    altura_pt = altura_mm * _PT_PER_MM
    doc = fitz.open()
    try:
        page = doc.new_page(width=largura_pt, height=altura_pt)
        page.insert_image(fitz.Rect(0, 0, largura_pt, altura_pt), filename=png_path)
        doc.save(output_path, garbage=3, deflate=True)
    finally:
        doc.close()
    return output_path


def gerar_anexo_unifilar(output_path, *, caminho_xlsx, papel_pdf, fonte_esperada,
                         figuras_dir, motivo_fallback):
    """Fase 3.4/v0.2: tenta a prancha REAL do Diagrama Unifilar
    (`figuras.unifilar.UnifilarRenderer`, le a propria planilha via
    openpyxl+Pillow -- sem Word/Excel/LibreOffice/COM) e SO cai no
    placeholder textual A0 (`gerar_placeholder_unifilar`, item 1 deste
    modulo, ja usado desde a Fase 3.3B) se a renderizacao real falhar por
    qualquer motivo -- fonte ausente, planilha corrompida, erro de leitura,
    erro ao embutir o PNG no PDF. NUNCA deixa a montagem quebrar: sempre
    devolve um `output_path` valido, com o motivo registrado quando cai no
    fallback. Nunca inventa conteudo de diagrama em nenhum dos dois
    caminhos.

    Devolve um dict: `ok` (bool), `metodo`, `papel_pdf`, e o resto do
    contexto -- `provenance`/`note` da figura quando `ok=True`, `motivo`
    quando cai no placeholder."""
    from .figuras import unifilar as UF

    largura_mm, altura_mm = _dims_papel_pdf_mm(papel_pdf)
    resultado, erro_excecao = None, None
    try:
        resultado = UF.gerar(caminho_xlsx, figuras_dir, largura_pagina_mm=largura_mm,
                             altura_pagina_mm=altura_mm)
    except Exception as e:
        erro_excecao = repr(e)

    if resultado is not None and resultado.status == "OK" and resultado.path:
        try:
            _png_para_pdf_pagina(resultado.path, output_path, largura_mm, altura_mm)
            return {"ok": True, "metodo": "unifilar_renderer_openpyxl_pillow_fitz",
                    "papel_pdf": papel_pdf, "png_path": resultado.path,
                    "note": resultado.note, "provenance": resultado.provenance}
        except Exception as e:
            motivo = ("%s -- prancha real foi renderizada (%s) mas falhou ao ser "
                      "embutida no PDF (%r); caindo no placeholder"
                      % (motivo_fallback, resultado.path, e))
    else:
        motivo_real = (resultado.note if resultado is not None else
                      "exceção ao chamar UnifilarRenderer: %s" % erro_excecao)
        motivo = "%s -- %s" % (motivo_fallback, motivo_real)

    gerar_placeholder_unifilar(output_path, motivo=motivo, fonte_esperada=fonte_esperada,
                               largura_mm=largura_mm, altura_mm=altura_mm)
    return {"ok": False, "metodo": "placeholder_reportlab", "papel_pdf": papel_pdf,
            "motivo": motivo}
