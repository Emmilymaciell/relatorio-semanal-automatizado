# DEC-002 — Adotar o modelo de 11 seções (desde 2026-08-01) como alvo da automação

- **Data:** 2026-09-16
- **Situação:** aceita
- **Substitui parcialmente:** DEC-001 (que fixava a fronteira em 2026-06-20)
- **Afeta:** modelo canônico, regras de negócio, matriz de automação

---

## Contexto

A DEC-001 fixou a fronteira do "regime vigente" em **2026-06-20**, com base na
evolução do conjunto de arquivos auxiliares. Essa fronteira estava **incompleta**.

A varredura do conteúdo dos PDFs (`scripts/08_historico_anexos.py`) revelou uma
mudança **estrutural** posterior, que a análise de nomes de arquivo não podia ver:

> Em **2026-08-01**, nas quatro obras simultaneamente, a seção
> **6. DESVIO DE CRONOGRAMA** foi removida do relatório, junto com o anexo
> "Desvio de cronograma acumulado". Todas as seções seguintes foram renumeradas:
> o documento passou de **12 para 11 seções**.

Regra A3 em `tests/test_regras_anexos.py`: **52/52**, sem exceção.

## Decisão

1. O alvo da automação é o **modelo de 11 seções**, vigente desde **2026-08-01**.
2. A fronteira do "regime vigente" passa de 2026-06-20 para **2026-08-01**.
3. Toda regra estrutural do mapeamento é declarada **válida para o modelo vigente**.
   Regras validadas sobre semanas anteriores (p. ex. as de calendário, testadas
   desde 2026-01-03) mantêm sua própria fronteira, declarada caso a caso.
4. A amostra dos 8 relatórios **permanece válida**: todos são de 2026-08-22 em
   diante, portanto inteiramente dentro do modelo novo.

## Justificativa

Misturar os dois modelos produziria um mapeamento com uma seção fantasma e a
numeração errada em seis seções — exatamente o tipo de erro que o briefing pede
para evitar.

Vale registrar **por que a DEC-001 errou o alvo**: ela inferiu regimes a partir do
*conjunto de arquivos auxiliares*, e a mudança de 01/08 não mexeu em nenhum arquivo
auxiliar — foi uma mudança **dentro do documento**. Nomes de arquivo não bastam;
foi preciso abrir os PDFs, como o briefing determina.

## Consequências

- ✅ O modelo canônico (`config/report_schema.yaml`) descreve o documento atual.
- ✅ Ficou documentado um marco de processo que ninguém precisaria explicar.
- ⚠️ **O modelo vigente tem só ~7 semanas de vida** (5 relatórios concluídos por obra).
  É uma base estreita. A automação deve ser construída de forma que a **remoção ou
  inclusão de uma seção inteira** seja uma mudança de configuração, não de código —
  esse tipo de alteração já aconteceu duas vezes em quatro meses.
- ⚠️ Vale perguntar à Emmily se a remoção do "Desvio de Cronograma" é definitiva ou
  se a seção pode voltar. Registrado em `14_perguntas_para_emmily.md`.
