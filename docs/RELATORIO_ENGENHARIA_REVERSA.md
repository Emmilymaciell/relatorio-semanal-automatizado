# Relatório de Engenharia Reversa — Relatório Semanal de Acompanhamento de Obras do Programa

**Data:** 2026-09-17 · **Baseline:** `discovery-v2` (substitui `discovery-v1`)
**Fase:** Descoberta **concluída e validada**
**Amostra analisada em profundidade:** 8 relatórios · **Corpus validado:** 177 relatórios · **Corpus textual:** 189
**Universo inventariado:** 193 pastas-semana · 3.944 arquivos do ecossistema · 34.839 fotos

---

## ⚡ O que mudou na discovery-v2

A v1 deixou 11 perguntas em aberto. A v2 incorporou as respostas humanas da Emmily,
fez cinco investigações dirigidas e validou o corpus inteiro. **Dez mudanças materiais:**

| # | Mudança | Antes (v1) | Agora (v2) |
|---|---|---|---|
| 1 | **A pasta do UNIFILAR existe** | "não existe no SharePoint" ❌ | `Programa - 2. Compartilhado interno\03_Gestao dos contratos\` ✅ |
| 2 | **A nascente do avanço físico** | desconhecida | **digitação manual** — 0 fórmulas em 259 células |
| 3 | **`\\SERVIDOR\`** | suspeita de fonte externa | vínculo reescrito; 24 de 28 alvos usam caminho relativo |
| 4 | **Fotos têm metadados** | "sem EXIF, sem data" ❌ | **overlay visual em 30/30**: data · UTM · estacas · serviço · rodovia · índice |
| 5 | **Coordenada UTM** | `DESCONHECIDO` | está impressa na própria foto — `VAL` |
| 6 | **Descrição da NC** | `MAN` | **`VAL`** — template reproduziu 3/3 com fidelidade **1,000** |
| 7 | **Contagens da 8.3** | `HIPÓTESE` | **`VAL`** — `VISAO_GERAL!Tbl_VISAO`, 4/4 exato em 4 obras |
| 8 | **Seção 9 com ocorrência** | `DESCONHECIDO` | **10 casos reais** → anatomia de 7 elementos |
| 9 | **Necessidade de IA** | ~11 % do conteúdo | **0 % obrigatório** — 69,8 % fixo, 22,3 % template |
| 10 | **Revisões** | não observadas | **5 revisões `.2`**, nenhuma `.3`; regra de versionamento testada |

**Validação:** 15 regras × 177 relatórios = **1634 PASS · 30 FAIL · 56 WARNING · 610 N/A**.
**No modelo vigente (v3): 298 PASS, 0 FAIL.** Os 30 erros são históricos e decrescentes
(v1: 27 · v2: 3 · **v3: 0**), e a análise de causa mostrou que **nenhum** decorre de
mudança de regra: são erro humano real (18) ou paginação do PDF (12).

Consolidação completa em `CONSOLIDACAO_P01_P11.md` *(não incluído na versão pública)* e
veredito de prontidão em [`GERADOR_READINESS_GATE.md`](GERADOR_READINESS_GATE.md).

> Nada neste documento foi explicado por uma pessoa. Tudo foi inferido dos documentos,
> dos seus metadados e da comparação entre versões. Onde não houve prova, está escrito
> `HIPÓTESE` ou `DESCONHECIDO`.
>
> **Este documento não propõe implementar nada ainda.** Ele é o entendimento a ser
> revisado antes de qualquer automação.

---

## Sumário executivo

O Relatório Semanal é um documento de **11 seções**, produzido semanalmente para quatro
obras, com período **sábado a sexta**, emitido cerca de uma semana após o fim do período,
e fechado **em lote para as quatro obras na mesma sessão de trabalho**.

Três coisas foram descobertas e valem mais que o resto:

1. **O processo é inteiramente manual e baseado em cópia.** Não existe template. O
   `.docx` de cada semana é literalmente o arquivo da semana anterior salvo com outro
   nome — os quatro documentos nasceram em setembro de 2025 e vêm sendo editados desde
   então (`cp:revision` ~2000, subindo ~30 por semana).

2. **O documento não contém dados — contém figuras.** Não há um único objeto OLE nem
   gráfico nativo. Fichas, unifilar, gráficos e painéis entram como **imagens EMF
   coladas**. Isso define a decisão de arquitetura mais cara do projeto.

3. **Quase tudo é rastreável.** Fechamos, com prova numérica, a cadeia do avanço
   físico, a do clima, a das fotos, a das não conformidades e a ambiental. O trabalho
   genuinamente humano se reduz a **selecionar qual foto entra** e **decidir sobre
   ocorrências atípicas** — a descrição da NC, que parecia manual, é extração.

**Correção da v1:** o alerta de que "a nascente do indicador principal estaria fora do
SharePoint" **estava errado**. A pasta existe em
`Programa - 2. Compartilhado interno\03_Gestao dos contratos\`, ao lado do `[Programa]` mestre,
e os vínculos são relativos. O que de fato ocorre é outra coisa, mais simples e mais
relevante: **a nascente é digitação manual** — as colunas `G` e `H` do UNIFILAR não têm
uma única fórmula em 259 células. Isso não impede a automação; apenas define onde ela
começa.

---

## 1. Como é estruturado o Relatório Semanal?

Onze seções, **idênticas nas quatro obras** (nenhuma tem seção a mais ou a menos):

| Nº | Seção | Papel |
|---|---|---|
| — | Capa | A4 retrato |
| — | Índice | A4 retrato |
| 1 | APRESENTAÇÃO | A4 |
| 2 | FICHA DE ACOMPANHAMENTO | A4, imagem EMF |
| 3 | AVANÇO GLOBAL DA OBRA | A4, imagem EMF |
| 4 | UNIFILAR / DIAGRAMA UNIFILAR | A4 |
| 5 | EXECUÇÃO DE OBRA (5.1 clima, 5.2 fotográfico) | **5.2 em A3 retrato** |
| 6 | ACOMPANHAMENTO AMBIENTAL | A4 |
| 7 | DIÁRIO DE OBRAS | A4 |
| 8 | ACOMPANHAMENTO DE NÃO CONFORMIDADES (8.1–8.3) | A4 |
| 9 | RELATO DE INTERFERÊNCIAS EM CAMPO | A4 |
| 10 | TERMO DE ENCERRAMENTO | A4 |
| 11 | ANEXOS | A4 + anexos |

**O papel identifica o conteúdo** — e isso virou regra de validação:

| Formato | Conteúdo | Quantidade |
|---|---|---|
| A4 retrato | corpo | 13–27 pg |
| **A3 retrato** | **seção 5.2, fotos** | 3–4 pg |
| **A2 paisagem** | **Diário de Obras** | **sempre 7** |
| A0 paisagem | Diagrama Unifilar | 1 (A3 na RD-102A) |

Detalhes em `docs/mapeamento/03_estrutura_relatorio.md` *(não incluído na versão pública)*.

---

## 2. Quais partes são universais?

- A estrutura de 11 seções e sua numeração.
- O título da capa.
- O texto introdutório da seção 5.2, palavra por palavra nas 4 obras.
- O texto da seção 8.2 ("não houve NC da Órgão Contratante"), idêntico em 8/8.
- O texto-base da seção 4 (Unifilar), variando só `{rodovia}`.
- O texto-base da seção 7 (Diário de Obras).
- O template do Termo de Encerramento (seção 10).
- A regra dos anexos: Unifilar e Diário sempre por último, nessa ordem.
- Os **logos** — os PNGs são byte-a-byte idênticos entre as 4 obras.

---

## 3. Quais variam por obra?

Sete campos de cadastro (rodovia, trecho, extensão, código SRE, lote, executora,
supervisora) mais o contrato, tudo em
[`config/road_registry.yaml`](../config/road_registry.yaml).

E divergências de **redação**, herdadas da cópia:

| Elemento | RD-101 | RD-102A | RD-103 | RD-104 |
|---|---|---|---|---|
| Seção 4 | UNIFILAR | DIAGRAMA UNIFILAR | DIAGRAMA UNIFILAR | UNIFILAR |
| Seção 5.1 | Condições climáticas | Condições climáticas | **Dados** climáticos | Condições climáticas |
| Seção 8.1 | **pela** Programa | pela | **pelo** | pela |
| Seção 9 | RELATO DE **RISCOS E** INTERFERÊNCIAS | INTERFERÊNCIAS | INTERFERÊNCIA (sing.) | INTERFERÊNCIAS |
| Unifilar | A0 | **A3** | A0 | A0 |
| Nº de página | ✅ | ❌ | ✅ | ❌ |
| Legenda de foto | genérica | genérica | genérica | **nomeia o serviço** |
| Tabela de bueiros | ✅ | ❌ | ❌ | ❌ |

---

## 4. Quais variam por semana?

- Número e período (determinísticos).
- A **lista de serviços da semana** — e com ela as subseções 5.2.x.
- As fotos.
- A linha CLIMA e a frase sobre dias sem execução.
- A lista de anexos.
- As contagens de não conformidades.
- O número de páginas (e portanto o Termo de Encerramento).

---

## 5. De onde vem cada dado?

Matriz completa em `08_data_lineage.md` *(não incluído na versão pública)*. As três cadeias
fechadas:

### 5.1 Avanço físico — CONFIRMADO

```
UNIFILAR!'Base para BI' E/F  →[link externo]→  [Programa]!'Avanço Global - <OBRA>' H18/I18
   →  C3 = SUM(L12:L..), C7 = SUM(N12:N..)  →[cópia manual]→  Gráfico!'Relatório Semanal'
   →  barChart  →  Gráfico 1 (EMF) na seção 3
```

Provas: 212/212 células de cache do link externo; 12/12 números em 6/6 semanas
(RD-101 nº 47: `C3 = 0,15425815718` → `B6 = 15,43`); 99/99 datas na série acumulada.

### 5.2 Clima — CONFIRMADO

`FICHA de Campo` diária, aba `FICHA CAMPO`: `N16/O16/P16` = Bom/Chuvoso/Impraticável,
marcação com `X` na linha 17. Verificado em 12 fichas.

E a frase *"No dia … não houve execução de atividades"* corresponde ao **dia sem ficha
de campo** — não a uma regra de domingo. Confirmado três vezes de forma independente
(ausência da ficha, ausência da pasta de fotos, ausência do serial de data no
fotográfico).

### 5.3 Fotos — CONFIRMADO

O campo `Indice` do Relatório Fotográfico **cai dentro do intervalo** da coluna `Foto`
da Ficha de Campo do mesmo dia, com estaca e serviço coincidentes. 5/5 verificados.

---

## 6. Quais arquivos participam?

| Arquivo | Papel | Onde |
|---|---|---|
| `Ficha de Campo Acumulada GO <obra>.xlsx` | seção 2 | `Arquivos auxiliares` |
| `Gráfico - Avanço Global.xlsx` | seção 3 | `Arquivos auxiliares` |
| `[Programa] Avanço Global de Obra.xlsx` | fonte do gráfico; **arquivo central das 4 obras** | `Arquivos auxiliares` |
| `UNIFILAR <obra>.xlsx` | anexo + elo da cadeia | `Arquivos auxiliares` |
| `RELATÓRIO FOTOGRÁFICO <nº>.xlsx` | seção 5.2 — **73 a 167 MB** | `Arquivos auxiliares` |
| `AAAA-MM-DD FICHA de Campo <obra>.xlsx` | clima, estacas, índices de foto | `1. Ficha de campo` |
| `3. Fotos\<AA-MM-DD>\<Serviço>\*.jpeg` | acervo (34.839 fotos, ~12,9 GB) | `3. Fotos` |
| `Diário de obra digital assinado\… .pdf` | seção 7 + anexo — **emitido pelo Parceiro** | `4. Diários de Obra` |
| `Ficha de Acompanhamento Semanal` | seções 5.1, 8, 9 (**RD-101 abandonou**) | `2. Ficha…` |
| `3. Notificações\Not. Geradas\…` | seção 8 e anexos NC | `3. Notificações` |
| `6. Acompanhamento ambiental e SST` | seção 6 e anexos | pasta homônima |
| ❓ `UNIFILAR (com não conformidades)` | **nascente do avanço físico — não localizado** | fora do SharePoint |

---

## 7. Quais anexos participam?

Dois fixos, sempre no fim: **Diagrama Unifilar** (penúltimo) e **Diário de Obras**
(último, 7 páginas A2). Antes deles, condicionais:

| Anexo | Regra | Confiança |
|---|---|---|
| Acompanhamento ambiental + SST | **quinzenal**: `(início − 2026-06-06) mod 14 == 0` | 52/52 |
| Parecer Técnico PA-NNN | irregular, **não segue o ciclo**; numeração corporativa única | ALTA |
| NC-NNN | NC emitida dentro do período | ALTA |
| Ata de Reunião | houve reunião no período | ALTA |
| ~~Desvio de cronograma~~ | **extinto em 2026-08-01** | 52/52 |

---

## 8. Quais textos são estáticos, templated ou precisam de IA?

| Tipo | Seções |
|---|---|
| **Estático** | 5.2 (intro), 8.2, 9, 4 (base), 7 (base), 5.1 (classificações), título da capa |
| **Template com variáveis** | 1 (par. 1 e 2), 10 (Termo — 100% determinístico) |
| **IA a partir de dados** | 1 (par. 3 e 4), 4 (frase de serviços), 6, 8.1 (moldura) |
| **Humano** | 8.1 (descrição da NC), casos extraordinários |

---

## 9. Como funciona a numeração?

`nº = (data_inicio − âncora da obra)/7 + 1`, validado em **144/144** semanas do regime
vigente. Âncoras: RD-101 `2025-10-04`, RD-102A e RD-103 `2025-09-27`, RD-104 `2025-11-08`.

Arquivo: `RO_<NNN>.<rev> - <RODOVIA>.<ext>`, com `NNN` em 3 dígitos. Só `.1` na amostra.

---

## 10. Como funciona o período?

**Sábado a sexta**, 7 dias, semanas contíguas — 144/144 desde `2026-01-03`.

Antes disso havia outro regime (**sexta→quinta**), com semanas de 8 e 15 dias na
implantação; a virada se deu por uma semana-pivô de 8 dias em 21–28/11/2025.

---

## 11. Como funciona o layout?

A4 retrato 21 × 29,7 cm; margens 2,5/3,0/2,5/3,0 cm (RD-103 usa 3,0/2,0/2,0/3,0 fora da
capa). **Uma única seção A3 retrato** por documento — a do Relatório Fotográfico, nunca
paisagem. Todas as quebras são `nextPage`.

Corpo em **Arial 11 pt**: por estilo `Normal` na RD-101/RD-102A, por **formatação direta
run a run** na RD-103/RD-104. Títulos 16 pt, definidos com cor `0F4761` no estilo mas
carimbados como preto no run.

**A hierarquia de títulos é falsa:** todos usam o estilo custom `Estilo1` com
`outlineLvl = 0`; o "5.2.1." vem do `w:ilvl` da lista, não do nível de título.

Imagens: 100% em linha no corpo, 100% flutuantes no cabeçalho, ajustadas à largura útil
(15,00 cm em A4 e 23,70 cm em A3; 16,00/24,70 na RD-103).

Duas tabelas apenas: a climática (3×8, em 8/8) e a de bueiros (52×6, só RD-101) — e
**nenhuma usa `tblHeader`**, então a de 52 linhas perde o cabeçalho ao virar a página.

Detalhes em `04_layout_visual.md` *(não incluído na versão pública)*.

---

## 12. Como entram fotografias, gráficos, tabelas e anexos?

**Todos como imagem EMF colada** (Colar Especial → Metarquivo), de 2 a 4,6 MB cada.
Não há OLE, não há gráfico nativo, não há dado estruturado no `.docx`.

As fotos chegam ao documento em blocos de 3, dentro das páginas A3, vindas das abas do
`RELATÓRIO FOTOGRÁFICO` (uma aba por serviço). Nessa planilha, cada foto é uma
**imagem na célula** (`xl/richData/`) com metadados imediatamente acima:
`Estaca · Coordenada UTM · Sensível · Índice · Descrição`.

Os anexos entram como PDFs concatenados, em ordem definida.

---

## 13. Qual parece ser o processo atual?

Reconstruído em `11_processo_atual_inferido.md` *(não incluído na versão pública)*.
Resumidamente: durante a semana, campo preenche ficha diária e despeja fotos; na
quarta/quinta seguinte, gabinete atualiza as planilhas em sequência (ficha acumulada →
unifilar → `[Programa]` + gráfico → fotográfico); no dia seguinte, copia o `.docx` da semana
anterior, troca período e número, recola as imagens, reescreve os textos, atualiza
índice e legendas à mão, anexa os PDFs e exporta.

---

## 14. O que pode ser automatizado?

Dos 44 elementos mapeados ([`12_matriz_automacao.md`](mapeamento/12_matriz_automacao.md)):

| Nível | v1 | **v2** |
|---|---|---|
| `DET` 100% determinística | 24 (55%) | **27 (59%)** |
| `VAL` extração validada | 8 (18%) | **12 (26%)** |
| `TPL` template parametrizado | — | **3 (7%)** |
| `IA` gerável por IA | 5 (11%) | **0 (0%)** |
| `HITL` human-in-the-loop | 3 (7%) | **3 (7%)** |
| `MAN` necessariamente manual | 1 (2%) | **1 (2%)** |
| `?` indeterminado | 3 (7%) | **0 (0%)** |

**Por volume de texto:** 69,8 % fixo · 22,3 % template · 7,8 % livre — e o "livre" se
decompõe em ~75 blocos `DET`, ~62 `VAL`, ~120 `TPL` e **apenas ~7 realmente humanos**.

**Ressalva honesta:** 59 % dos *elementos* são determinísticos, mas não são 59 % do
*esforço*. O esforço real está no Relatório Fotográfico (~40 %) e na produção das
imagens (~20 %). O maior ganho da automação é eliminar a **digitação de metadados que
já estão impressos na própria foto**.

---

## 15. Quais são as exceções?

- RD-102A: unifilar em A3, não A0; parágrafo permanente sobre a NC-025 na seção 4.
- RD-101: única com tabela de bueiros; abandonou a Ficha de Acompanhamento Semanal em
  11/05/2026; migrou o diário para digital em 25/04 (as outras em 21/04).
- RD-103: pasta `Arquivos Auxiliares` com A maiúsculo; margens diferentes; "Dados
  climáticos"; "pelo Programa".
- RD-104: pasta `4. Relatório Semanal **de** Obras`; legendas de foto nomeiam o serviço.
- RD-102A e RD-104 não numeram as páginas.

---

## 16. Quais são as incertezas?

**Resolvidas na v2:** a pasta do UNIFILAR (existe) · a fonte das contagens de NC
(`Tbl_VISAO`) · a origem das coordenadas UTM (overlay da foto) · o significado de
A/C/I (legenda da própria ficha) · o Dashboard Ambiental (Power BI sobre a planilha de
licença) · a seção 9 com interferência real (10 casos) · as revisões `.2` (5 casos).

**O que continua em aberto:**

1. 🟡 De onde vem o número que a pessoa **digita** em `G`/`H` do UNIFILAR — só fonte
   humana resolve. **Não bloqueia**: é entrada anterior ao gerador.
2. 🟢 Comportamento da seção 8.2 se a Órgão Contratante emitir NC — **133/133 negativas**, sem
   precedente. Cai em `HITL` se ocorrer.
3. 🟢 Revisão `.3` — nunca ocorreu; a regra já é genérica em `rev`.
4. 🟠 O modelo vigente tem ~7 semanas de vida — mitigado por `DEC-002`: estrutura é
   configuração, não código.

---

## 17. O que ainda preciso responder?

As 11 perguntas de `14_perguntas_para_emmily.md` *(não incluído na versão pública)*,
cada uma com o que já foi investigado, a hipótese atual e o impacto.

---

## 18. Qual arquitetura recomendo?

A pipeline sugerida no briefing **faz sentido**, com uma ressalva grande: o passo
"renderização Word" é muito mais caro do que parece, porque o documento atual é feito de
imagens, não de dados.

```
SharePoint (read-only)
   ↓ scanner: identifica obra e semana pelo calendário
   ↓ coleta: fichas diárias, fotos, planilhas auxiliares, diário do Parceiro, NC, ambiental
   ↓ extração estruturada  →  modelo canônico (config/report_schema.yaml)
   ↓ regras determinísticas: período, numeração, clima, anexos, termo, índice
   ↓ HITL: seleção de fotos entre faixas já qualificadas pela ficha de campo
   ↓ produção de imagens (decisão A/B/C pendente)
   ↓ narrativas por IA: seções 1, 4, 6, 8.1
   ↓ montagem do .docx  →  anexação dos PDFs  →  validação (V1–V7)  →  PDF
   ↓ revisão humana só nas exceções
```

**Três recomendações concretas:**

1. **Comece por um validador, não por um gerador.** Um programa que roda sobre os
   relatórios já existentes e confere as 7 regras (`V1`–`V7` do
   [`report_schema.yaml`](../config/report_schema.yaml)) tem risco zero, valor imediato,
   e — principalmente — **prova se este mapeamento está correto** antes de escrevermos
   qualquer gerador.

2. **Trate a estrutura como configuração, não como código.** Em 12 meses houve **sete**
   mudanças de formato, incluindo a remoção de uma seção inteira há 7 semanas. Um
   gerador com as seções codificadas quebra na próxima mudança.

3. **Resolva P-01 antes de automatizar o avanço físico.** Enquanto a fonte estiver fora
   do SharePoint, essa parte dependerá de alguém continuar colocando o arquivo na pasta.

---

## 19. Resumo das regras validadas por teste

| Teste | Regras | Casos |
|---|---|---|
| `tests/test_regras_periodo.py` | R1–R5: sábado→sexta, 7 dias, contiguidade, numeração, PDF ⇔ concluído | **144/144** |
| `tests/test_regras_anexos.py` | A1–A4: ordem dos anexos, bloco ambiental quinzenal, extinção do desvio de cronograma, ambiental+SST acoplados | **52/52** |

E as regras verificadas sobre a amostra (8/8): 7 páginas A2, 1 página de unifilar,
contagem de páginas do Termo, correspondência serviços ↔ subseções, ordem dos dois
anexos fixos, 7 colunas na tabela climática, data de emissão ↔ `mtime` do PDF.

---

## 20. Correções feitas durante a investigação

Registradas porque mostram onde a inferência rápida falha:

| Afirmação inicial | Correção | Como foi pega |
|---|---|---|
| "período sábado→sexta em 193/193" | só desde 2026-01-03; antes era sexta→quinta | teste sobre o acervo inteiro |
| "RD-104 não tem arquivos auxiliares" | tem — os caminhos estouravam MAX_PATH | prefixo `\\?\` |
| "páginas A3 são a Ficha de Acompanhamento" | são o Relatório Fotográfico | inspeção página a página |
| "bloco ambiental é mensal" | é **quinzenal** (52/52) — e o próprio texto já dizia | varredura de 14 semanas |
| "regime vigente desde 2026-06-20" (DEC-001) | desde **2026-08-01** (DEC-002) | abrir os PDFs, não só os nomes |

> A última é a mais instrutiva: a DEC-001 inferiu regimes a partir de **nomes de
> arquivo**, e a mudança de 01/08 não tocou em nenhum arquivo auxiliar — foi dentro do
> documento. Confirma a instrução do briefing: **não analisar somente nomes de arquivo**.
