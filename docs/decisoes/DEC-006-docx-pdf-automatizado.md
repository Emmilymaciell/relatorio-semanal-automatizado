# DEC-006 — Conversão DOCX→PDF automatizada (Word COM isolado)

**Data:** 2026-10-04 · **Status:** ACEITA · **Substitui:** a regra "Word COM não funciona — conversão manual" (STATE/checkpoints até v0.3)

## Contexto

Até a v0.3, a conversão do corpo DOCX em PDF era o único passo manual do pipeline, e travou a v0.3
inteira (o PDF "convertido" nunca chegou ao disco). As tentativas anteriores de Word COM travavam e
deixavam processo órfão.

## Teste controlado (2026-10-04)

`scripts/gerador/docx_to_pdf.py`, nesta ordem:

1. DOCX **sintético** (TOC `\o "1-3"` + Heading 1/2 + campo PAGE no rodapé + 2 páginas):
   OK em 16,7 s — TOC recalculado com números, PAGE 1/2, Word encerrado sem órfão.
2. O mesmo, com **outra instância do Word aberta** (simulando a da usuária): OK em 7,8 s; a
   instância alheia sobreviveu e encerrou sozinha.
3. Duas conversões **simultâneas**: OK, serializadas pelo lock (11,4 s e 10,5 s).
4. DOCX **real** da v0.3 (43,5 MB): OK em 43,9 s, 29 páginas, TOC correto (sem auto-listagem).

Teste permanente: `tests/test_docx_to_pdf.py` (D1..D5; SKIP sem Word).

## O que difere das tentativas anteriores

- `DispatchEx` (instância nova) dentro de **subprocesso** com timeout — o processo pai nunca fica
  preso a uma chamada COM bloqueada;
- trabalha sobre **cópia** em `%TEMP%`; o DOCX original nunca é aberto pelo Word;
- `Visible=False`, `DisplayAlerts=0`, `UpdateLinksAtOpen=False`, `ConfirmConversions=False`;
- identifica o PID próprio pela diferença de processos WINWORD e, em timeout, mata **só** ele;
- lock entre processos (`%TEMP%\prog_docx2pdf.lock`) — a diferença de PIDs só é confiável serializada;
- valida o PDF (fitz abre, >0 páginas) antes de declarar sucesso.

## Decisão

**`DOCX_TO_PDF = AUTOMATIZADO`.** Não é `AUTOMATIZAVEL_SE_WORD_FECHADO`: o teste 2 provou a
coexistência com um Word aberto. Fallback manual continua possível (`generate.py --pdf-manual`).

## Risco residual

Diálogo modal inesperado do Word (ex.: arquivo corrompido, add-in) → o timeout mata a instância
própria e a etapa reporta `TIMEOUT`; o pipeline para com instrução de conversão manual.
