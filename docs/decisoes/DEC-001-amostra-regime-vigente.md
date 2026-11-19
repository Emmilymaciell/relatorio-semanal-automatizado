# DEC-001 — Restringir a amostra ao regime de modelo vigente

- **Data:** 2026-09-16
- **Situação:** aceita
- **Afeta:** seleção de amostra, regras de negócio, futuro motor de automação

---

## Contexto

O briefing pedia 2 relatórios por obra (8 no total) e autorizava ampliar a
amostra caso houvesse mudanças relevantes de versão/modelo ao longo do tempo,
desde que justificado.

O rastreamento do conjunto de arquivos auxiliares nas 193 pastas-semana mostrou
que o processo **mudou em marcos datados e simultâneos nas quatro obras**:

- **2026-02-07** — criação da pasta `Arquivos auxiliares` (3 planilhas);
- **2026-03-28** — entrada de `Gráfico - Avanço Global.xlsx`;
- **2026-04-04 a 2026-06-27** — entrada de `[Programa] Avanço Global de Obra.xlsx`;
- **2026-06-20** — substituição de `Ficha de campo` por `Ficha de Campo Acumulada`.

Antes de 2026-02-07 as pastas-semana continham apenas o par DOCX/PDF, sem
nenhum rastro dos insumos.

## Decisão

1. A amostra permanece com **8 relatórios** (2 por obra), toda ela dentro do
   regime vigente desde ~2026-06-20 (5 auxiliares fixos).
2. Relatórios anteriores a 2026-06-20 **não são usados como modelo** do processo
   atual. Serão consultados apenas como contraexemplo pontual, quando for preciso
   testar se uma regra é estrutural ou circunstancial.
3. Todo campo do mapeamento herda essa fronteira: uma regra validada só nas
   semanas ≥ 2026-06-20 é declarada válida **para o regime vigente**, não para
   o acervo inteiro.

## Justificativa

Ampliar a amostra para trás misturaria regimes de processo distintos e produziria
regras contraditórias — exatamente o risco que o briefing pede para evitar. As
duas semanas escolhidas (2026-08-22..08-28 e 2026-08-29..09-04) são idênticas
nas quatro obras, o que permite os dois eixos de comparação exigidos
(semana × semana e obra × obra) sobre a mesma janela de calendário.

## Consequências

- ✅ O mapeamento descreve o processo **como ele é hoje**, sem ruído histórico.
- ✅ Os marcos de mudança ficam documentados e datados, úteis para entender o
  histórico e para explicar divergências ao comparar relatórios antigos.
- ⚠️ Regras raras que só se manifestam em semanas atípicas (feriado, paralisação,
  chuva prolongada, mudança contratual) podem não aparecer na amostra. Mitigação:
  ao fim da fase 3, varrer o histórico procurando **seções ausentes ou
  excepcionais** e registrá-las em `13_pendencias.md`.
- ⚠️ A revisão `.2`+ não ocorre na amostra; a regra de revisão fica como
  `DESCONHECIDO` até varredura do histórico.
