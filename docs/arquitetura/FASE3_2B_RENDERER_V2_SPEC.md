# Fase 3.2B — P0 arquitetura + P1 design system

**Baseline:** `3fbf6f3` · **Criado em:** 2026-09-17
**Meta:** `output/shadow/RD103_2026-08-29_2026-09-04/shadow_report_v2.docx`
**Fora de escopo:** P2, P3, merge final do PDF, publicação, SharePoint.

Fontes de verdade: `docs/visual_review/GOLDEN_VISUAL_SPEC.md`,
`VISUAL_DIFF_MASTER.md`, `REVIEW_A/B/C.md`, `output/visual_review/visual_diff.json`.
**Não fazer nova engenharia reversa.**

---

## 0. Princípio

```
estrutura comprovada do Golden
+ design PREDOMINANTE do Golden
+ correção das inconsistências conhecidas
= renderer
```

| Classe | O que fazer |
|---|---|
| `GOLDEN_RULE` | **reproduzir** |
| `GOLDEN_INTERNAL_VARIATION` | implementar o **predominante**; registrar a exceção, não codificá-la |
| `GOLDEN_DEFECT` | **não reproduzir**; corrigir e documentar |

`shadow_report.docx` (v1) é baseline — **não sobrescrever**.

---

## P0 — ARQUITETURA

### P0.1 `DISCOVERY_CORRECTION` — o DOCX não tem A0 nem A2

O `document.xml` do golden tem **9 seções, todas A4 ou A3 retrato**. As páginas A0
(unifilar) e A2 (diário) do PDF final **não vêm do Word**: são PDFs externos mesclados
depois (`pdf_producer = iLovePDF`), sem cabeçalho e sem número de página.

**Corrigir toda documentação e configuração que ainda diga que o DOCX tem A0/A2**,
marcando como `DISCOVERY_CORRECTION` — em especial `config/report_schema.yaml`
(`s11_anexos.diagrama_unifilar.papel`, `diario_de_obras.papel`) e o que houver em
`report_versions.yaml` / `docs/mapeamento/`. O papel A0/A2 continua **verdadeiro para o
PDF final**; o que é falso é atribuí-lo a uma seção do Word.

### P0.2 Arquitetura de montagem

```
ReportDocument
├── main_document            -> DOCX -> PDF do corpo
│   ├── sections             (A4 corpo, A4 capa, A3 fotográfico)
│   ├── images               figuras e gráficos
│   └── embedded_annexes     anexos EMBEDDED_IMAGE
└── external_pdf_annexes     anexos MERGED_PDF
            ↓
      merge -> PDF final   (NÃO nesta rodada)
```

Camada nova e explícita, separando **conteúdo** de **montagem**. Sugestão:
`scripts/gerador/assembly_model.py`. O `content_builder` continua produzindo o
`document_model`; o assembly model diz **como o documento vira arquivo**.

Nesta rodada a montagem **não executa o merge** — apenas descreve o plano e o
serializa em `output/shadow/.../assembly_plan.json`.

### P0.3 Anexos por modo de integração — **semântico, não posicional**

Proibido `if anexo.numero <= 3`. O modo vem da **configuração do tipo de anexo**:

```yaml
# config/annex_types.yaml
anexos:
  parecer_tecnico:        {integration_mode: EMBEDDED_IMAGE, ordem: 1}
  acompanhamento_ambiental: {integration_mode: EMBEDDED_IMAGE, ordem: 2}
  acompanhamento_sst:     {integration_mode: EMBEDDED_IMAGE, ordem: 3}
  nao_conformidade:       {integration_mode: EMBEDDED_IMAGE}
  ata_reuniao:            {integration_mode: EMBEDDED_IMAGE}
  diagrama_unifilar:      {integration_mode: MERGED_PDF, posicao: penultimo, papel_pdf: "A0 paisagem"}
  diario_de_obras:        {integration_mode: MERGED_PDF, posicao: ultimo, papel_pdf: "A2 paisagem"}
```

| Modo | Comportamento |
|---|---|
| `EMBEDDED_IMAGE` | vira imagem de página cheia **dentro** do DOCX; herda cabeçalho e numeração |
| `MERGED_PDF` | PDF externo, mesclado **depois** da exportação; **sem** cabeçalho e **sem** número de página |

Evidência: golden p17–26 (PA-254 6 pág., ambiental 2, SST 2) têm logo e número;
p27–35 não têm.

### P0.4 Relatório Fotográfico — imagem composta por serviço

O golden usa **1 EMF por serviço** (5 EMF no documento; 16 JPEG para ~46 fotos — as
fotos estão dentro do metarquivo). **Não inserir JPEG um a um no Word.**

Modelar a cadeia, mesmo que o compositor ainda não produza imagem perfeita:

```
fotos do serviço  ->  PhotoPanelComposer  ->  imagem composta  ->  Word
```

Interface obrigatória agora; implementação pode ser mínima nesta rodada (por exemplo,
gerar a imagem com o grid e molduras vazias, ou continuar emitindo placeholder **desde
que pelo componente certo**). O que **não** pode é a arquitetura errada.

Especificação do painel (`GOLDEN_VISUAL_SPEC.md` §7): A3 retrato, largura útil 247 mm,
painel 247,00 × 306,32 mm, **3 colunas fixas**, linhas = `ceil(n/3)`, coluna ≈ 80 mm,
altura 50–57 mm por foto, legenda de 3 linhas **acima** de cada foto, linha incompleta
mantém células vazias com borda.

### P0.5 `DEC-004` — formato da imagem composta

**EMF não é obrigatório.** Era consequência do processo manual (colar do Excel).
A automação pode usar **PNG de alta resolução** (ou SVG), desde que: fidelidade visual,
dimensão correta em mm, estável no Word e boa qualidade no PDF.
Registrar em `docs/decisoes/DEC-004-formato-das-imagens-compostas.md`.

### P0.6 Seções do Word

Implementar as seções reais: **capa A4** (margens próprias, `titlePg`), **corpo A4**,
**fotográfico A3**, e voltar a **corpo A4**. Não inventar seção A0/A2.

---

## P1 — DESIGN SYSTEM

### P1.0 `config/visual_style_v3.yaml` — fonte única

Proibido número mágico espalhado pelo renderer. Centralizar: `page`, `margins`,
`fonts`, `colors`, `body`, `h1`, `h2`, `h3`, `toc`, `header`, `footer`, `tables`,
`captions`, `spacing`. O renderer **consome** esse arquivo.

### P1.1 Página e margens

| Perfil | T | D | B | E |
|---|---|---|---|---|
| `cover_a4` | 25 | 30 | 25 | 30 mm |
| `body_a4` | **30** | **20** | **20** | **30** mm |
| `photo_a3` | 30 | 20 | 20 | 30 mm |

`header distance` = `footer distance` = **12,5 mm**. `gutter` = 0.
Largura útil: A4 **160 mm**, A3 **247 mm**.

### P1.2 Cabeçalho — **sem reproduzir o defeito**

Dois logos e uma linha horizontal. O `header2` do golden é `GOLDEN_DEFECT`: foi
dimensionado para A3 e herdado em A4, jogando o logo do Programa para fora da folha.

**Regra obrigatória:** posição do logo direito e comprimento da linha calculados a
partir de `usable_width = page_width − left_margin − right_margin` **da própria seção**.
Nada de coordenada absoluta.

Referência (`header3`, correto para A4, `GOLDEN_VISUAL_SPEC.md` §3.3):

| Objeto | Tamanho | Posição |
|---|---|---|
| Logo Parceiro | 41,1 × 13,7 mm | à esquerda, offset 0 da margem |
| Logo Programa | 21,9 × 11,7 mm | **encostado à direita da área útil** (no A4 isso dá ≈137,7 mm) |
| Linha | `usable_width` × 0,4 pt | offset 0 da margem, ≈11 mm abaixo |

Em A4, `usable_width − 21,9 = 160 − 21,9 = 138,1 mm` — bate com os 137,7 mm medidos.
**É essa a fórmula**, e ela funciona sozinha no A3.

Se os arquivos de logo não estiverem disponíveis no repositório, **não inventar
imagem**: reservar o espaço, marcar `LIMITATION` e registrar de onde o logo deve vir.

### P1.3 Capa

`titlePg = true` na primeira seção → **capa sem cabeçalho e sem rodapé**.

### P1.4 Rodapé

Campo **`PAGE` real do Word** (não texto). Alinhado à **direita**, Arial, 12,5 mm da
borda. O golden herda Aptos aqui — `GOLDEN_INTERNAL_VARIATION`, **usar Arial**.

### P1.5 Corpo

Arial 11 pt, cor `000000`, **justificado** (`jc=both`), **entrelinha 1,5**,
espaço depois 0. Aplicar por **estilo**, não parágrafo a parágrafo.

### P1.6 Títulos

| | H1 | H2 | H3 |
|---|---|---|---|
| Fonte | Arial | Arial | Arial |
| Tamanho | **16 pt** | 14 pt | 12 pt |
| Cor | **`000000`** | `000000` | `000000` |
| Recuo esq. | **12,7 mm** | conforme nível | conforme nível |
| `keepNext`/`keepLines` | true | true | true |

> O golden **não tem H2/H3 dedicados** — usa `Estilo1` com override manual.
> Aqui o shadow é **melhor**: definir H2 e H3 de verdade. Registrar como melhoria
> deliberada, não como divergência.

### P1.7 Numeração automática multinível

Proibido escrever `"1. APRESENTAÇÃO"` no texto. Usar lista multinível do Word
(`numbering.xml`) ligada aos níveis dos estilos, produzindo `1.`, `5.1.`, `5.2.`,
`5.2.1.`, `8.1.` etc. **semanticamente**, por nível — não por prefixo de string.

### P1.8 Índice

Tab stops por nível, com `leader="dot"` (**nunca** pontos digitados):

| Nível | Tab esquerdo | Tab direito | Leader |
|---|---|---|---|
| 1 | **8,47 mm** (480 tw) | **149,8 mm** (8494 tw) | `dot` |
| 2 | **12,7 mm** (720 tw) | 149,8 mm | `dot` |
| 3 | **16,93 mm** (960 tw) | 149,8 mm | `dot` |

Arial 12 pt, espaço depois 5 pt.

**Número de página:** avaliar o campo `TOC` do Word (atualiza sozinho ao abrir/imprimir).
Se não for confiável, mecanismo controlado — mas **jamais** copiar a paginação do golden.
O índice tem de refletir **o novo documento** e sair do `document_model`, incluindo
automaticamente os níveis `5.2.x` existentes na semana. Nada de lista fixa.

> `GOLDEN_INTERNAL_VARIATION`: no golden os itens `10.` e `11.` usam recuo de nível 2
> sendo nível 1. **Não reproduzir.**

### P1.9 Tabela do clima

8 colunas: **15 / 20 / 20 / 20 / 22,5 / 25 / 20 / 20 mm** (total **162,5 mm**,
`tblW` em `dxa`). Bordas **0,5 pt** (`w:sz=4`) em todos os lados e internas.
Sombreado **`F2F2F2`** na **coluna de rótulos** (`DATA:` / `DIA:` / `CLIMA:`).

### P1.10 Legendas

Arial **10 pt**, cor **`0D0D0D`**, centralizado (predominante).
Numeração gerada por sequência independente: `Figura`, `Gráfico`, `Tabela`.

### P1.11 Q-12 — seção 9

Título vem de **`road_registry.titulo_secao_9`, por obra**.
Para RD-103: `RELATO DE INTERFERÊNCIA EM CAMPO`. **Não acrescentar "RISCOS"** —
ver a revisão da R-05 em `GUIA_DE_REDACAO.md`.

### P1.12 `compatibilityMode = 15`

Atualizar `settings.xml`, com teste que impeça o retorno silencioso ao 14.

### P1.13 Fontes

**Arial em tudo**, inclusive índice e rodapé. As ocorrências de Aptos no golden são
`GOLDEN_INTERNAL_VARIATION` — registrar, **não** implementar.

---

## Testes obrigatórios — OOXML, sem depender do Word

`tests/test_ooxml_estrutura_v2.py`, lendo o `.docx` gerado:

margens por perfil de seção · `compatibilityMode == 15` · nº de seções e papéis ·
`titlePg` na capa · header presente com 2 imagens + linha · **logo direito dentro da
largura útil em toda seção** · footer presente · campo `PAGE` real · Arial em todo
`rFonts` · H1 16 pt · H1 cor `000000` · `numPr` nos títulos (numeração automática) ·
tab stops do índice nos 3 níveis · `leader="dot"` · tabela do clima com as 8 larguras ·
bordas 0,5 pt · shading `F2F2F2` · **nenhum título com número escrito no texto**.

`tests/test_assembly_model.py`: anexos classificados por `integration_mode` vindo de
configuração; `EMBEDDED_IMAGE` vs `MERGED_PDF` corretos; plano serializado.

**Regressão:** as 15 suítes atuais continuam passando. Visual não pode quebrar conteúdo —
período, número, serviços, avanço, clima, PA-254, NC, anexos.

---

## Comparação estrutural

`output/shadow/.../renderer_v2_structural_diff.json`, comparando **Golden × v1 × v2**
para cada item P0/P1: `FIXED` · `PARTIALLY_FIXED` · `NOT_FIXED` · `REGRESSED`.

---

## Saída e parada

- `shadow_report_v2.docx` (o v1 **permanece**)
- `assembly_plan.json`, `renderer_v2_structural_diff.json`
- **Não** converter para PDF: o Word COM trava neste ambiente. **Não insistir.**
- **Parar** no Checkpoint 3.2B.
