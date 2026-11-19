# -*- coding: utf-8 -*-
"""
test_termo_patch_geometria.py — Fase 3.4 (v0.3)

Trava a lição da regressão da página 17 da v0.2.

O patch do Termo reescreve o parágrafo DENTRO do PDF já montado, por redaction
(sem Word/COM). Até a v0.2 ele usava coordenadas Y absolutas fixas:

    rect_paragrafo  = Rect(x0 - 3, 118, x1 + 3, 205)
    rect_fechamento = Rect(x0 - 3, 213, x1 + 3, 240)

Isso só funcionava enquanto o Termo era a única coisa da sua página. Quando a
política de quebra da v0.2 juntou as seções 9 e 10, a faixa 118..270 pt passou
a conter o CORPO DA SEÇÃO 9 e o TÍTULO DA SEÇÃO 10 — que foram apagados,
enquanto o placeholder sobrevivia intacto mais abaixo.

Os dois cenários abaixo são construídos aqui mesmo, em PDFs sintéticos, para
não dependerem de qual política de quebra está configurada hoje:

  G1  seção 10 sozinha numa página nova       (layout da v0.1)
  G2  seções 9 e 10 na MESMA página           (layout da v0.2)

Em ambos, depois do patch:
  - o texto final aparece, com o número real de páginas;
  - o placeholder some por inteiro (nada de "___" ou "[PENDENTE]");
  - o título "10. TERMO DE ENCERRAMENTO" continua lá;
  - o conteúdo da seção 9 continua lá (é o que G2 protege);
  - nenhum conteúdo acima do título é tocado.

G3 lê o código e garante que não voltou coordenada Y absoluta.

Execute:  python tests/test_termo_patch_geometria.py
"""
import os
import re
import sys
import inspect

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(AQUI), "scripts")
sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
import fitz  # noqa: E402
from gerador import pdf_final_assembly as PA  # noqa: E402

TEXTO_FINAL = ("O Relatório Semanal de Acompanhamento de Obras da RD103 - N° 49, "
               "referente à consolidação e à apresentação das atividades desenvolvidas "
               "durante o período de 29/08/2026 a 04/09/2026, no âmbito do Contrato "
               "Nº003/2025 - Parceiro X Executora C, apresenta-se constituído de "
               "37 (trinta e sete) páginas.")
FECHAMENTO = "Cidade, 21 de setembro de 2026."

PLACEHOLDER_PAR = ("[PENDENTE] O Relatório Semanal de Acompanhamento de Obras da RD103 "
                   "- N° 49, referente à consolidação e à apresentação das atividades "
                   "desenvolvidas durante o período de 29/08/2026 a 04/09/2026, no "
                   "âmbito do Contrato Nº003/2025 - Parceiro X Executora C, "
                   "apresenta-se constituído de ___ (___) páginas.")
PLACEHOLDER_FECH = "Cidade, ___ (data de emissão só é conhecida na montagem final)."
PLACEHOLDER_NOTA = ("Nota: n_paginas_provisorio=true -- Termo definitivo é gerado no "
                    "Checkpoint 3.3, após a montagem final (nunca copiado do golden)")

SECAO9_TITULO = "9. RELATO DE INTERFERÊNCIA EM CAMPO"
SECAO9_CORPO = ("Durante o período abrangido por este relatório, não foram registrados "
                "riscos ou interferências em campo.")
TERMO_TITULO = "10. TERMO DE ENCERRAMENTO"


def _escrever(page, y, texto, size=11):
    page.insert_textbox(fitz.Rect(85, y, 542, y + 90), texto,
                        fontsize=size, fontname="helv", align=fitz.TEXT_ALIGN_JUSTIFY)


def _pdf_cenario(caminho, junto):
    """junto=False -> Termo sozinho na 2a pagina (layout v0.1)
       junto=True  -> secoes 9 e 10 na MESMA pagina (layout v0.2)"""
    doc = fitz.open()
    p1 = doc.new_page(width=595.32, height=841.92)
    _escrever(p1, 100, "Conteúdo anterior que não pode ser tocado pelo patch.")

    if junto:
        p = doc.new_page(width=595.32, height=841.92)
        _escrever(p, 95, SECAO9_TITULO, size=14)
        _escrever(p, 125, SECAO9_CORPO)
        _escrever(p, 220, TERMO_TITULO, size=14)
        _escrever(p, 250, PLACEHOLDER_PAR)
        _escrever(p, 345, PLACEHOLDER_FECH)
        _escrever(p, 365, PLACEHOLDER_NOTA, size=8)
    else:
        p9 = doc.new_page(width=595.32, height=841.92)
        _escrever(p9, 95, SECAO9_TITULO, size=14)
        _escrever(p9, 125, SECAO9_CORPO)
        p = doc.new_page(width=595.32, height=841.92)
        _escrever(p, 100, TERMO_TITULO, size=14)
        _escrever(p, 130, PLACEHOLDER_PAR)
        _escrever(p, 225, PLACEHOLDER_FECH)
        _escrever(p, 245, PLACEHOLDER_NOTA, size=8)

    doc.save(caminho)
    doc.close()


def _checar(nome, caminho_in, caminho_out, falhas, exige_secao9_na_mesma_pagina):
    r = PA.patch_termo_pdf(caminho_in, caminho_out,
                           texto_paragrafo=TEXTO_FINAL, texto_fechamento=FECHAMENTO)
    if not r.get("ok"):
        falhas.append("%s: patch nao aplicou -- %s" % (nome, r.get("motivo")))
        return
    doc = fitz.open(caminho_out)
    texto = "\n".join(doc[i].get_text("text") for i in range(doc.page_count))
    pagina_patch = doc[r["pagina"] - 1].get_text("text")
    doc.close()

    erros = []
    if "37 (trinta e sete) páginas" not in re.sub(r"\s+", " ", texto):
        erros.append("texto final do Termo ausente")
    if "___" in texto:
        erros.append("lacuna '___' sobreviveu")
    if "[PENDENTE]" in texto:
        erros.append("marca [PENDENTE] sobreviveu")
    if "n_paginas_provisorio" in texto:
        erros.append("nota provisoria sobreviveu")
    if "TERMO DE ENCERRAMENTO" not in pagina_patch:
        erros.append("titulo '10. TERMO DE ENCERRAMENTO' foi apagado pelo patch")
    if "Conteúdo anterior" not in texto:
        erros.append("conteudo de outra pagina foi apagado")
    if exige_secao9_na_mesma_pagina:
        if "RELATO DE INTERFERÊNCIA" not in pagina_patch:
            erros.append("titulo da secao 9 apagado (era a regressao da v0.2)")
        if "interferências em campo" not in re.sub(r"\s+", " ", pagina_patch):
            erros.append("corpo da secao 9 apagado (era a regressao da v0.2)")
    else:
        if "RELATO DE INTERFERÊNCIA" not in texto:
            erros.append("secao 9 (em outra pagina) foi apagada")

    if erros:
        falhas.append("%s: %s" % (nome, "; ".join(erros)))
    else:
        print("%s  OK   patch na pagina %d; Termo reescrito, placeholder removido, "
              "vizinhanca intacta" % (nome, r["pagina"]))


def main():
    falhas = []
    tmp = os.path.join(U.PROJ, "tmp")
    os.makedirs(tmp, exist_ok=True)

    # ------------------------------------------------------------------ G1
    a_in = os.path.join(tmp, "test_termo_g1_in.pdf")
    a_out = os.path.join(tmp, "test_termo_g1_out.pdf")
    _pdf_cenario(a_in, junto=False)
    _checar("G1", a_in, a_out, falhas, exige_secao9_na_mesma_pagina=False)

    # ------------------------------------------------------------------ G2
    b_in = os.path.join(tmp, "test_termo_g2_in.pdf")
    b_out = os.path.join(tmp, "test_termo_g2_out.pdf")
    _pdf_cenario(b_in, junto=True)
    _checar("G2", b_in, b_out, falhas, exige_secao9_na_mesma_pagina=True)

    # ------------------------------------------------------------------ G3
    # nenhuma coordenada Y absoluta pode voltar ao calculo da area
    fonte = inspect.getsource(PA.patch_termo_pdf)
    corpo = "\n".join(l for l in fonte.splitlines()
                      if not l.strip().startswith("#") and '"""' not in l)
    suspeitas = re.findall(r"fitz\.Rect\(\s*[^)]*?\b\d{2,4}(?:\.\d+)?\s*,", corpo)
    literais = re.findall(r"fitz\.Rect\([^)]*\)", corpo)
    fixos = [l for l in literais if re.search(r",\s*\d{2,4}(\.\d+)?\s*,", l)]
    if fixos:
        falhas.append("G3: coordenada absoluta de volta em patch_termo_pdf: %s" % fixos[:2])
    elif "_localizar_blocos_placeholder" not in fonte:
        falhas.append("G3: patch_termo_pdf nao localiza mais a area pelo conteudo")
    else:
        print("G3  OK   area derivada de _localizar_blocos_placeholder(); nenhuma "
              "coordenada Y absoluta no calculo")

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
