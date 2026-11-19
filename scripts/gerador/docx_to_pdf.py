"""
Conversão DOCX -> PDF via Word COM, isolada e com timeout.

Contrato:
  - trabalha sobre uma CÓPIA temporária do DOCX (o original nunca é aberto pelo Word);
  - sobe uma instância PRÓPRIA do Word (DispatchEx) num subprocesso;
  - identifica o PID dessa instância pela diferença de processos WINWORD antes/depois;
  - em timeout, mata SOMENTE esse PID — nunca um Word que a usuária tenha aberto;
  - atualiza campos (todas as stories), TOCs, repagina e exporta (wdExportFormatPDF=17);
  - valida o PDF (abre com fitz, conta páginas) antes de declarar sucesso.

Uso:
    python scripts/gerador/docx_to_pdf.py ENTRADA.docx SAIDA.pdf [--timeout 600]
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time


def _pids_word():
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True, errors="replace").stdout
    pids = set()
    for linha in out.splitlines():
        partes = [p.strip('"') for p in linha.split('","')]
        if len(partes) > 1 and partes[0].lower().startswith("winword"):
            try:
                pids.add(int(partes[1]))
            except ValueError:
                pass
    return pids


def _worker(docx, pdf, pidfile):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    antes = _pids_word()
    word = win32com.client.DispatchEx("Word.Application")
    novos = _pids_word() - antes
    with open(pidfile, "w") as f:
        json.dump(sorted(novos), f)
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            word.Options.UpdateLinksAtOpen = False
        except Exception:
            pass
        doc = word.Documents.Open(docx, ConfirmConversions=False, ReadOnly=False,
                                  AddToRecentFiles=False, Visible=False,
                                  OpenAndRepair=False, NoEncodingDialog=True)
        try:
            for story in doc.StoryRanges:
                r = story
                while r is not None:
                    r.Fields.Update()
                    try:
                        r = r.NextStoryRange
                    except Exception:
                        r = None
            for toc in doc.TablesOfContents:
                toc.Update()
            doc.Repaginate()
            # segunda passada: números de página do TOC dependem da paginação final
            for toc in doc.TablesOfContents:
                toc.UpdatePageNumbers()
            doc.Save()
            doc.ExportAsFixedFormat(pdf, 17, OpenAfterExport=False, OptimizeFor=0,
                                    CreateBookmarks=1, DocStructureTags=True)
        finally:
            doc.Close(0)
    finally:
        word.Quit(0)
        pythoncom.CoUninitialize()


def _matar(pids):
    for pid in pids:
        subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)


def validar_pdf(pdf):
    import fitz
    if not os.path.exists(pdf) or os.path.getsize(pdf) < 1000:
        return {"ok": False, "motivo": "PDF ausente ou vazio"}
    with fitz.open(pdf) as d:
        n = d.page_count
        texto = "".join(d[i].get_text() for i in range(min(n, 3)))
    return {"ok": n > 0, "n_paginas": n, "bytes": os.path.getsize(pdf),
            "amostra_texto": texto[:200]}


class _LockGlobal:
    """Serializa conversões entre processos. A identificação do PID próprio é pela
    diferença de processos WINWORD antes/depois; duas conversões simultâneas
    confundiriam essa diferença e o cleanup poderia matar a instância da outra."""

    def __init__(self, espera=1800):
        self.path = os.path.join(tempfile.gettempdir(), "prog_docx2pdf.lock")
        self.espera = espera

    def __enter__(self):
        import msvcrt
        self.fh = open(self.path, "a+")
        t0 = time.time()
        while True:
            try:
                self.fh.seek(0)
                msvcrt.locking(self.fh.fileno(), msvcrt.LK_NBLCK, 1)
                return self
            except OSError:
                if time.time() - t0 > self.espera:
                    self.fh.close()
                    raise TimeoutError("lock de conversão ocupado por > %ds" % self.espera)
                time.sleep(2)

    def __exit__(self, *exc):
        import msvcrt
        try:
            self.fh.seek(0)
            msvcrt.locking(self.fh.fileno(), msvcrt.LK_UNLCK, 1)
        finally:
            self.fh.close()


def converter(docx, pdf, timeout=600, atualizar_docx=True):
    """Converte; devolve dict com status. Nunca levanta por timeout — reporta.

    atualizar_docx=True devolve o DOCX com campos/TOC atualizados ao lado do original
    (sufixo _campos.docx), sem tocar no original.
    """
    with _LockGlobal():
        return _converter(docx, pdf, timeout, atualizar_docx)


def _converter(docx, pdf, timeout, atualizar_docx):
    docx = os.path.abspath(docx)
    pdf = os.path.abspath(pdf)
    if not os.path.exists(docx):
        return {"status": "ERRO", "motivo": "DOCX não existe: %s" % docx}
    tmpdir = tempfile.mkdtemp(prefix="prog_docx2pdf_")
    copia = os.path.join(tmpdir, "entrada.docx")
    saida = os.path.join(tmpdir, "saida.pdf")
    pidfile = os.path.join(tmpdir, "pid.json")
    shutil.copyfile(docx, copia)
    t0 = time.time()
    proc = subprocess.Popen([sys.executable, os.path.abspath(__file__), "--worker",
                             copia, saida, pidfile],
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
                            errors="replace")
    status = "OK"
    erro = ""
    try:
        _, err = proc.communicate(timeout=timeout)
        if proc.returncode != 0:
            status, erro = "ERRO", err[-2000:]
    except subprocess.TimeoutExpired:
        status, erro = "TIMEOUT", "excedeu %ds" % timeout
        proc.kill()
    # cleanup: mata só a instância criada por nós, se ainda viva
    pids = []
    if os.path.exists(pidfile):
        with open(pidfile) as f:
            pids = json.load(f)
    vivos = [p for p in pids if p in _pids_word()]
    if vivos:
        time.sleep(2)
        vivos = [p for p in pids if p in _pids_word()]
        _matar(vivos)
    res = {"status": status, "erro": erro, "segundos": round(time.time() - t0, 1),
           "word_pids_proprios": pids, "word_pids_mortos": vivos}
    if status == "OK":
        v = validar_pdf(saida)
        res["validacao_pdf"] = v
        if not v["ok"]:
            res["status"] = "PDF_INVALIDO"
        else:
            os.makedirs(os.path.dirname(pdf), exist_ok=True)
            shutil.copyfile(saida, pdf)
            if atualizar_docx:
                shutil.copyfile(copia, os.path.splitext(pdf)[0] + "_campos.docx")
    shutil.rmtree(tmpdir, ignore_errors=True)
    return res


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--worker":
        _worker(sys.argv[2], sys.argv[3], sys.argv[4])
        return
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("pdf")
    ap.add_argument("--timeout", type=int, default=600)
    a = ap.parse_args()
    res = converter(a.docx, a.pdf, a.timeout)
    print(json.dumps(res, ensure_ascii=False, indent=2))
    sys.exit(0 if res["status"] == "OK" else 1)


if __name__ == "__main__":
    main()
