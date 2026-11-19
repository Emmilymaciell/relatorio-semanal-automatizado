# -*- coding: utf-8 -*-
"""
tests/test_docx_to_pdf.py - conversão DOCX->PDF via Word COM isolado.

Gera um DOCX sintético (TOC + Heading 1/2 + PAGE no rodapé + 2 páginas), converte
com gerador.docx_to_pdf.converter() e confere:
  D1  status OK e PDF com 2 páginas
  D2  TOC recalculado (placeholder sumiu; entradas com número de página)
  D3  campo PAGE recalculado (rodapé "Pagina 1" / "Pagina 2")
  D4  nenhum WINWORD criado pelo teste sobrevive
  D5  o DOCX de entrada não foi alterado (sha1 igual)
Sem Word instalado: SKIP (sai 0).
"""
import hashlib
import os
import sys
import tempfile

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "scripts"))

from docx import Document  # noqa: E402
from docx.enum.text import WD_BREAK  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402


def _campo(par, instr, placeholder):
    for tipo, extra in (("begin", None), ("instr", instr), ("separate", None)):
        r = par.add_run()
        if tipo == "instr":
            t = OxmlElement("w:instrText")
            t.set(qn("xml:space"), "preserve")
            t.text = extra
            r._r.append(t)
        else:
            f = OxmlElement("w:fldChar")
            f.set(qn("w:fldCharType"), tipo)
            r._r.append(f)
    par.add_run(placeholder)
    f = OxmlElement("w:fldChar")
    f.set(qn("w:fldCharType"), "end")
    par.add_run()._r.append(f)


def _sha1(p):
    with open(p, "rb") as fh:
        return hashlib.sha1(fh.read()).hexdigest()


def main():
    try:
        import win32com.client  # noqa: F401
        import winreg
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application")
    except Exception:
        print("SKIP: Word/pywin32 indisponível")
        return 0
    import fitz
    from gerador import docx_to_pdf as D

    tmp = tempfile.mkdtemp(prefix="test_docx2pdf_")
    docx = os.path.join(tmp, "sint.docx")
    pdf = os.path.join(tmp, "sint.pdf")
    d = Document()
    d.add_paragraph("SUMARIO")
    _campo(d.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u', "PLACEHOLDER_TOC")
    d.add_heading("Capitulo Um", 1)
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    d.add_heading("Capitulo Dois", 1)
    d.add_heading("Secao 2.1", 2)
    f = d.sections[0].footer.paragraphs[0]
    f.add_run("Pagina ")
    _campo(f, "PAGE", "9")
    d.save(docx)
    sha_antes = _sha1(docx)

    res = D.converter(docx, pdf, timeout=180, atualizar_docx=False)
    falhas = []

    def chk(cod, ok, msg):
        print("%s  %s  %s" % (cod, "OK  " if ok else "FAIL", msg))
        if not ok:
            falhas.append(cod)

    chk("D1", res["status"] == "OK" and res.get("validacao_pdf", {}).get("n_paginas") == 2,
        "status=%s paginas=%s" % (res["status"], res.get("validacao_pdf", {}).get("n_paginas")))
    texto = []
    if os.path.exists(pdf):
        with fitz.open(pdf) as p:
            texto = [pg.get_text() for pg in p]
    t0 = texto[0] if texto else ""
    chk("D2", "PLACEHOLDER_TOC" not in t0 and "Capitulo Dois" in t0 and "Secao 2.1" in t0,
        "TOC recalculado")
    chk("D3", len(texto) == 2 and "Pagina 1" in texto[0] and "Pagina 2" in texto[1],
        "PAGE recalculado")
    vivos = [p for p in res.get("word_pids_proprios", []) if p in D._pids_word()]
    chk("D4", not vivos, "WINWORD próprios vivos: %s" % vivos)
    chk("D5", _sha1(docx) == sha_antes, "DOCX de entrada intacto")
    print("RESULTADO: %s" % ("OK" if not falhas else "FALHA %s" % falhas))
    return 1 if falhas else 0


if __name__ == "__main__":
    sys.exit(main())
