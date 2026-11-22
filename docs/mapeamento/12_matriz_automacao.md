# 12 — Matriz de automação (V2)

**Atualizado em:** 2026-09-17 · **Versão:** 2 (substitui a V1 da `discovery-v1`)
**Base:** as 5 frentes de investigação concluídas + validação de 177 relatórios

---

## Níveis

| Nível | Sigla | Significado |
|---|---|---|
| 100 % determinística | `DET` | regra fechada, testável, sem julgamento |
| Extração validada | `VAL` | leitura mecânica de fonte estruturada + conferência |
| Template parametrizado | `TPL` | forma fixa, variáveis preenchidas |
| IA generativa grounded | `IA` | só sobre fatos estruturados, com verificação factual |
| Human-in-the-loop | `HITL` | a máquina propõe, a pessoa decide |
| Manual | `MAN` | julgamento puro |
| Indeterminado | `?` | falta evidência |

**Ordem de preferência (nunca inverter):**
`DET → VAL → TPL → IA grounded → HITL → MAN`

---

## 1. O que mudou da V1 para a V2

| Elemento | V1 | **V2** | Motivo |
|---|---|---|---|
| **Descrição da NC (8.1)** | `MAN` | **`VAL`** | O texto sai de `VISAO_GERAL!V` de `02_Acompanhamento de NC automatizada.xlsm`. O template reproduziu **3 de 3** frases com similaridade **1,000**. Não é caso de IA. |
| **Contagens de NC (8.3)** | `?` | **`VAL`** | Fonte confirmada: `VISAO_GERAL!Tbl_VISAO`, coluna `Status`, domínio fechado de 6 valores. Cruzamento **4/4 exato** em 4 das 8 seções; total fecha em 7/8. |
| **Metadados das fotos** | `MAN` (sem EXIF) | **`VAL`** | As fotos têm **overlay visual** em 30/30: data · UTM · estacas · serviço · rodovia · índice. `Estaca`, `Coordenada`, `Índice` e `Data` do Relatório Fotográfico são **transcritos do overlay**. |
| **Coordenada UTM** | `DESCONHECIDO` | **`VAL`** | Está impressa na própria foto. Encerra L-15. |
| **Seleção da fotografia** | `HITL` ~40 % | **`HITL`**, escopo menor | Escolher entre ~10 faixas já qualificadas por estaca/serviço, não entre ~300 fotos sem critério. |
| **Seção 9 com ocorrência real** | `DESCONHECIDO` | **`TPL` + `VAL` + `HITL`** | 10 casos históricos dão a anatomia: 7 elementos, 7 tipos de objeto. |
| **Painéis ambientais** | `VAL` 70 % | **`VAL`** + ressalva | São screenshots de Power BI (um é QR Code). A fonte de dados é a planilha de licença ambiental; painel 1 confere 8/8. |
| **Nº do Parecer PA** | `DET` (previsto) | **`VAL`** — nunca calcular | Contador **corporativo único** entre obras, atribuído na ordem da rotação de visitas. **Não é previsível**: a automação localiza o PA existente, jamais inventa o próximo. |
| **Avanço físico — origem** | `?` | **`MAN` na origem + `VAL` a jusante** | A nascente é **digitação manual** no UNIFILAR (0 fórmulas em 259 células). Tudo depois disso é automatizável. |
| **Texto geral** | 11 % gerável por IA | **0 % exige IA** | 69,8 % fixo · 22,3 % template · 7,8 % livre, e o "livre" é em boa parte falha de parametrização. |

---

## 2. Matriz por elemento

| Elemento | Hoje | Nível V2 | Fonte | Confiança |
|---|---|---|---|---|
| Período (datas) | manual | `DET` | calendário | 144/144 |
| Nº do relatório | manual | `DET` | `(início − âncora)/7 + 1` | 144/144 |
| Nome do arquivo e da pasta | manual | `DET` | padrão | 193/193 |
| Revisão (`.1`, `.2`) | manual | `DET` execução + `HITL` decisão | `REVISOES_E_VERSIONAMENTO.md` | 5/5 |
| Capa — 7 campos da obra | copiar | `DET` | `road_registry.yaml` | 8/8 |
| Índice | **manual** | `DET` | outline do documento | 8/8 |
| Seção 1 §1–2 | copiar | `DET` | template + cadastro | 189 rel. |
| Seção 1 §3 (período + serviços) | manual | `DET` | template + serviços da semana | 136 (72 %) |
| Seção 1 §4 (reuniões) | manual | `TPL` | condicional à ata | — |
| Seção 2 — Ficha (Figura 1) | colar EMF | `VAL` + render | `Ficha de Campo Acumulada` | confirmado |
| Seção 3 — Avanço (Gráfico 1) | colar EMF | `VAL` + render | cadeia UNIFILAR → `[Programa]` → Gráfico | 91/91 · 28/28 |
| Seção 4 — texto base | copiar | `DET` | template + `{rodovia}` | 56 (90 %) |
| Seção 4 — frase de serviços | manual | `DET` | serviços da semana | — |
| Seção 5.1 — texto das classificações | copiar | `DET` | fixo | 129 (100 %) |
| Seção 5.1 — DATA e DIA | manual | `DET` | período | 8/8 |
| Seção 5.1 — CLIMA | manual | `VAL` | `FICHA CAMPO!N17/O17/P17` | confirmado |
| Seção 5.1 — dias sem execução | manual | `DET` | dias sem Ficha de Campo | confirmado 3× |
| Seção 5.2 — introdução | copiar | `DET` | fixo | 123 (75 %) |
| Seção 5.2 — quais subseções | manual | `VAL` | abas do fotográfico | alta |
| Seção 5.2 — **metadados das fotos** | **digitar** | **`VAL`** | **overlay visual** | 30/30 |
| Seção 5.2 — **seleção da foto** | manual | **`HITL`** | faixas da Ficha de Campo | 13/13 |
| Seção 5.2 — montagem da página A3 | manual | `DET` | layout de blocos de 3 | 8/8 |
| Seção 5.2 — legendas | manual | `DET` | padrão + numeração | 8/8 |
| Seção 6 — texto | manual | `TPL` | template + nº do PA | 121 (72 %) |
| Seção 6 — 3 painéis | screenshot | `VAL` + render | planilha de licença ambiental | 8/8 |
| Seção 6 — nº do PA | manual | **`VAL`** (localizar, nunca calcular) | acervo ambiental | confirmado |
| Seção 7 — texto | copiar | `DET` | template + data de migração | 59 (48 %) |
| Seção 7 — ressalva de assinatura | manual | `TPL` | condicional | 5 casos |
| Seção 8.1 — frase negativa | copiar | `DET` | fixo | 54 |
| Seção 8.1 — **item de NC** | **manual** | **`VAL`** | `VISAO_GERAL!V` + `H` | **3/3 fidelidade 1,000** |
| Seção 8.2 | copiar | `DET` | fixo | **133/133** |
| Seção 8.3 — contagens | manual | **`VAL`** | `Tbl_VISAO`, coluna `Status` | 4/4 em 4 obras |
| Seção 8.3 — figura + 2 gráficos | colar | `VAL` + render | dinâmicas da planilha | confirmado |
| Seção 9 — frase negativa | copiar | `DET` | fixo por obra | 163 casos |
| Seção 9 — **ocorrência real** | manual | **`TPL` + `VAL` + `HITL`** | Ofício/PA + estaca + tipo | **10 casos** |
| Seção 10 — Termo | manual | `DET` | template + contagem final | 8/8 |
| Seção 11 — lista de anexos | manual | `DET` | anexos incluídos | 8/8 |
| Anexo Unifilar | manual | `DET` | sempre, penúltimo | 52/52 |
| Anexo Diário | manual | `DET` + `HITL` na exceção | 7 páginas A2 por período | 52/52 |
| Anexo ambiental + SST | manual | `DET` | quinzenal, âncora 2026-06-06 | 52/52 |
| Anexo NC | manual | `VAL` | NC do período | alta |
| Anexo Ata | manual | `VAL` | reunião no período | alta |
| Paginação | automática | `DET` | renderizador | — |
| Exportação PDF | manual | `DET` | — | — |

---

## 3. Placar V2

### 3.1 Por elemento (46 elementos)

| Nível | V1 | **V2** |
|---|---|---|
| `DET` | 24 (55 %) | **27 (59 %)** |
| `VAL` | 8 (18 %) | **12 (26 %)** |
| `TPL` | — | **3 (7 %)** |
| `IA` | 5 (11 %) | **0 (0 %)** |
| `HITL` | 3 (7 %) | **3 (7 %)** |
| `MAN` | 1 (2 %) | **1 (2 %)** ← só a digitação de origem no UNIFILAR |
| `?` | 3 (7 %) | **0 (0 %)** |

### 3.2 Por volume de texto produzido

`FIXO 69,8 %` · `TEMPLATE 22,3 %` · `LIVRE 7,8 %` — e o "livre" se decompõe em
~75 grupos `DET`, ~62 `VAL`, ~120 `TPL` e **~7 realmente `HITL`**.

### 3.3 Por esforço real — a leitura honesta

A contagem por elemento **subestima** onde o trabalho está. Estimativa por peso de
evidência (não cronometrada):

| Etapa | Peso do esforço hoje | Nível V2 | Ganho esperado |
|---|---|---|---|
| Montar o Relatório Fotográfico (selecionar ~30 fotos, digitar estaca/UTM/índice/descrição) | 🔴 ~40 % | `VAL` + `HITL` | **alto** — a digitação some, resta escolher |
| Produzir as imagens EMF (ficha, gráfico, painéis) | 🟠 ~20 % | `VAL` + render | alto |
| Reescrever os textos | 🟠 ~15 % | `DET`/`TPL` | **muito alto** — 92 % é fixo ou template |
| Atualizar índice, legendas e paginação | 🟡 ~10 % | `DET` | total |
| Atualizar as planilhas de avanço | 🟡 ~10 % | `MAN` origem + `VAL` | parcial — a digitação no UNIFILAR permanece |
| Anexar PDFs | 🟢 ~5 % | `DET` | total |

> **Leitura:** o maior ganho não está nos 59 % `DET` — está em eliminar a **digitação
> dos metadados das fotos**, que hoje é copiada à mão de uma informação que já está
> impressa na própria imagem.

---

## 4. Os três problemas fotográficos, separados

O briefing pediu explicitamente para não tratar "fotos" como um problema só.

| Subproblema | Nível | Situação |
|---|---|---|
| **1. Extração dos metadados** | **`VAL`** | Resolvido conceitualmente: overlay em 30/30, com data · UTM · estaca · serviço · rodovia · índice. Falta só decidir a técnica de leitura (ver ordem de preferência abaixo). |
| **2. Seleção da fotografia** | **`HITL`** | A máquina propõe as faixas candidatas (Ficha de Campo: ~10 por dia, já com estaca e serviço); a pessoa escolhe qual imagem representa melhor o serviço. |
| **3. Montagem do fotográfico** | **`DET`** | Layout de blocos de 3, uma aba por serviço, legendas padronizadas, ordenação por índice. |

**Ordem de preferência para obter cada campo** (usar visão/OCR só no fim da fila):
1. fonte estruturada original (Ficha de Campo, planilha)
2. nome do arquivo / pasta
3. cruzamento (data + índice) → linha da ficha
4. **overlay visual da foto**
5. visão/OCR/IA

> ⚠️ Cuidado documentado: **a pasta de serviço não é classificação confiável** — há
> fotos `PBA` e `PSL` arquivadas em `Terraplenagem`. O serviço correto vem de
> (data, índice) → linha da ficha. E a data do nome do arquivo é a hora de envio no
> WhatsApp, podendo cair no dia seguinte.

---

## 5. Ordem recomendada de implementação

| # | Entrega | Nível | Por quê nesta ordem |
|---|---|---|---|
| 1 | **Validador** (já existe) | `DET` | ✅ feito — 298 PASS / 0 FAIL no CURRENT |
| 2 | **Modelo canônico + extratores das fontes confirmadas** | `VAL` | ficha de campo, NC, avanço, ambiental — todas as cadeias já provadas |
| 3 | **Montador do esqueleto textual** | `DET`/`TPL` | 92 % do texto, sem depender de imagem |
| 4 | **Anexador de PDFs** | `DET` | regras de composição e ordem já validadas 52/52 |
| 5 | **Assistente do Relatório Fotográfico** | `VAL`+`HITL` | maior ganho de esforço |
| 6 | **Renderizador de imagens** | `VAL` | decisão A/B/C ainda pendente (`DEC-003`) |
| 7 | Narrativas por template | `TPL` | seção 9 com ocorrência, ressalvas |

**IA generativa não aparece nesta lista.** Continua disponível como recurso opcional
para o miolo descritivo de interferências e ressalvas atípicas, sempre `grounded` e com
verificação factual — nunca como mecanismo principal.
