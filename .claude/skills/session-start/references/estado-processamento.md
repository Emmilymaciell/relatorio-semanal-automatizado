# Estado de processamento incremental — onde olhar

Carregue só quando a sessão anterior mexeu em volume (corpus, candidate, anexos, figuras).
Para uma sessão de documentação ou configuração, pule este arquivo.

**Nunca reinicie trabalho concluído. Nunca pule pendência em silêncio.**

---

## Shadow report — a frente principal

Um "candidate" é uma versão do relatório reconstruído. Cada rodada tem artefatos próprios em
`output/shadow/<OBRA>_<inicio>_<fim>/`:

| Arquivo | Prova |
|---|---|
| `candidate_manifest_<ver>.json` | versão gerada; `validation_result.document_status_counts` traz READY/HITL/LIMITATION/PLACEHOLDER |
| `shadow_final_candidate_<ver>.docx` | corpo gerado |
| `shadow_final_candidate_<ver>.pdf` | **conversão manual** — se faltar, a rodada parou aqui |
| `shadow_final_candidate_<ver>_MONTADO.pdf` | montagem final (corpo + unifilar + diário) |
| `validacao_pdf_final_<ver>.json` | 15 regras do validador |
| `critical_gates_<ver>.json` | 12 gates |
| `comparison_<ver>.json` / `.md` | comparação item a item com o golden |
| `regression_guard_<base>_<nova>.json` | veredito LIBERA / LIBERA_COM_JUSTIFICATIVA / BLOQUEIA |

Para descobrir onde parou:

```bash
ls output/shadow/*/candidate_manifest_*.json          # versionado -> autoridade
ls output/shadow/*/shadow_final_candidate_*_MONTADO.pdf   # NAO versionado
```

> ⚠️ **Num clone novo os binários não existem** — `.docx`, `.pdf`, figuras e páginas de anexo
> ficam fora do Git (embutem fotos da obra e páginas de pareceres; ver `.gitignore`). Ausência
> de `_MONTADO.pdf` em máquina recém-clonada **não** significa rodada incompleta: significa que
> os derivados ainda não foram regerados ali.
>
> A **autoridade é o `candidate_manifest_<ver>.json`**, que é versionado e carrega os `sha1` das
> fontes e da configuração. Só conclua "rodada incompleta" quando o manifesto existir **e** a
> árvore tiver os demais derivados daquela rodada (sinal de que a máquina já trabalhou nela).

Quando for mesmo rodada incompleta, a causa é quase sempre a mesma: falta a conversão manual
DOCX→PDF.

> ⛔ **Word COM não funciona neste ambiente.** Testado várias vezes; trava e deixa processo órfão.
> A conversão DOCX→PDF é manual, feita pela Emmily. Não insista, e desconsidere qualquer
> afirmação de que páginas foram medidas via COM.

## Corpus histórico

`output/extracao/texto/` — um `.txt` por relatório (~187 legíveis).
`output/inventarios/` — inventário do ecossistema e progresso incremental.

> O relatório oficial da semana em estudo (**golden**) serve **só para comparar**. Nunca é fonte
> para preencher o shadow report. Há guarda em `scripts/gerador/__init__.py`.

## Verificações

```bash
for f in tests/test_*.py; do python -X utf8 "$f" >/dev/null 2>&1 || echo "FALHOU: $f"; done
python scripts/validar_relatorio_semanal.py --current
python scripts/21_regression_guard.py --obra RD103 \
    --inicio 2026-08-29 --fim 2026-09-04 --base v0.1 --nova v0.2
```

O guard tem 4 camadas: A fatos · B status · C DOCX · D **PDF montado**.
Sem o PDF montado ele **bloqueia** — e é esse o comportamento correto, não uma falha.

## O que reportar

- obra e semana correntes
- versões já montadas e a rodada incompleta, se houver
- veredito do último guard e o que ele apontou
- pendências com severidade (`comparison_<ver>.json`, campo `classe`/`severidade`)
- o que **não** precisa ser refeito
