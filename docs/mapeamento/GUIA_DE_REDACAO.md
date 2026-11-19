# Guia de Redação — Relatório Semanal de Acompanhamento de Obras

**Criado em:** 2026-09-16
**Resolve:** P-10
**Autoridade:** decisão da Emmily — o relatório futuro deve ser formal, técnico,
profissional, coerente, corporativo, gramaticalmente correto e **uniforme entre obras**.

> ⚠️ **Regra que não se quebra:** padroniza-se **forma**, nunca **fato**.
> Se duas obras têm contratadas diferentes, contratos diferentes ou extensões
> diferentes, isso permanece diferente. Uniformizar é sobre redação, capitalização,
> terminologia e estilo — não sobre dados.

---

## 1. Princípios

1. **Impessoalidade.** Terceira pessoa, voz passiva ou sujeito indeterminado.
   "Foram executados os serviços…", nunca "Nós executamos…".
2. **Tempo verbal.** Pretérito perfeito para o que ocorreu no período;
   presente para o que descreve o documento.
3. **Uma grafia por conceito.** O mesmo serviço, órgão ou documento tem sempre a
   mesma forma em todo o relatório e em todas as obras.
4. **Sem abreviação improvisada.** Sigla só quando consagrada e definida na
   primeira ocorrência.
5. **Frase negativa é padronizada.** "Não houve" tem uma única redação oficial
   por seção — hoje há quatro redações diferentes para a mesma ausência.

---

## 2. Nomenclatura oficial

### 2.1 Rodovias

| Forma oficial | Onde usar | Formas a **não** usar |
|---|---|---|
| `RD-101`, `RD-102A`, `RD-103`, `RD-104` | todo o corpo do relatório | `RD 101`, `RD-102 A`, `RD102A`, `#RD 103` |
| `RD101`, `RD102A`, `RD103`, `RD104` | **apenas** no Termo de Encerramento e em nome de arquivo | — |

> A capa da RD-102A hoje grafa `RD-102 A` (com espaço) e o termo grafa `RD102A`.
> **Padronizar a capa para `RD-102A`.**

### 2.2 Órgãos e empresas

| Forma oficial | Observação |
|---|---|
| `Parceiro` | Parceiro. Por extenso na 1ª ocorrência da seção 1. |
| `Órgão Contratante` | sempre em caixa alta |
| `Programa` | Consórcio do Programa. **Masculino: "o Programa"** |
| `Órgão Supervisor` | — |
| `Fundo Estadual de Infraestrutura` | Fundo Estadual de Infraestrutura |

> ⚠️ **Concordância do PROG.** Hoje três obras escrevem "emitidas **pela** Programa" e a
> RD-103 escreve "emitidas **pelo** Programa". Sendo *Consórcio*, o correto é **"pelo Programa"**.
> → **padronizar para `pelo Programa`** em todas as obras. É correção gramatical, não estilo.

Nomes de contratadas e supervisoras: usar exatamente o que consta de
`config/road_registry.yaml`, sem abreviar e sem inventar sigla.

### 2.3 Serviços

Usar sempre o `canonical_name` de `config/services_taxonomy.yaml`:

| Correto | Não usar |
|---|---|
| Obras de Arte Corrente | OAC, Obras de arte corrente |
| Obras de Arte Especiais | OAE |
| Terraplenagem | Terraplanagem |
| Pavimentação | — |
| Drenagem | — |
| Obras Complementares | Obras complementares, Obra complementares |

Sigla só entre parênteses após a primeira ocorrência, se necessário:
*"Obras de Arte Corrente (OAC)"*.

### 2.4 Documentos

| Forma oficial | Exemplo |
|---|---|
| Não Conformidade nº `NC-NNN` | NC-022 |
| Parecer Técnico nº `PA-NNN` | PA-255 |
| Contrato nº `NN/AAAA` | 001/2025 |
| Ordem de Serviço nº `NN/AAAA` | 01/2025 |
| Relatório Semanal nº `NNN` | 048 |

---

## 3. Datas, números e percentuais

| Elemento | Padrão | Exemplo |
|---|---|---|
| Data no corpo | `dd/mm/aaaa` | 29/08/2026 |
| Data por extenso (Termo) | `dd de <mês> de aaaa`, mês minúsculo | 11 de setembro de 2026 |
| Data na tabela climática | `dd/mmm` abreviado | 29/ago |
| Dia da semana (tabela climática) | caixa alta | SEXTA-FEIRA |
| Período | `de dd/mm/aaaa a dd/mm/aaaa` | de 29/08/2026 a 04/09/2026 |
| Decimal | vírgula | 40,00 km |
| Milhar | ponto | 151.600.155,38 |
| Percentual | vírgula, 2 casas, espaço antes do símbolo | 15,43 % |
| Extensão | número + espaço + `km` | 35,00 km (hoje a RD-102A grafa `35,00km`) |
| Estaca | `NNN+NN` | 548+00 |
| Intervalo de estacas | `Est. NNN+NN a NNN+NN` | Est. 548+00 a 573+00 |
| Nº do relatório | 3 dígitos com zero à esquerda | 048 |
| Nº de páginas (Termo) | algarismo + extenso entre parênteses | 39 (trinta e nove) páginas |

---

## 4. Frases padrão

As frases abaixo são o **texto oficial**. Onde há `{variável}`, o valor vem do
modelo canônico. Onde a frase é condicional, a condição está indicada.

### 4.1 Seção 1 — Apresentação

**Período e serviços** *(template)*:
> As informações apresentadas neste relatório referem-se ao período de
> **{data_inicio}** a **{data_fim}**. No intervalo considerado, foram registradas
> atividades de **{lista_de_servicos}**.

`{lista_de_servicos}`: nomes canônicos, minúsculos no corpo da frase, separados por
vírgula e com "e" antes do último.

**Reuniões** *(condicional)*:
- sem reunião:
  > Durante o período, não foram realizadas reuniões de alinhamento entre o Programa e a
  > empresa executora **{executora}** no canteiro de obras da Rodovia **{rodovia}**.
- com reunião:
  > Durante o período, foi realizada reunião de alinhamento entre o Programa e a empresa
  > executora **{executora}** no canteiro de obras da Rodovia **{rodovia}**, em
  > **{data}**, cuja ata integra os anexos deste relatório.

### 4.2 Seção 4 — Diagrama Unifilar

Texto-base **fixo** (idêntico nas 4 obras hoje — manter):
> O Diagrama Unifilar apresentado em anexo representa, de forma simplificada, o avanço
> físico acumulado da execução da Rodovia **{rodovia}**, indicando a progressão das
> frentes de serviço ao longo do trecho contratual. Esse instrumento permite visualizar,
> em um traçado linear, os segmentos concluídos, em execução e previstos, servindo como
> referência para o acompanhamento mensal e para a verificação da conformidade entre o
> avanço planejado e o efetivamente executado.

Complemento *(template)*:
> No intervalo considerado, foram registradas atividades de **{lista_de_servicos}**.

> Hoje a RD-103 escreve "No intervalo analisado, foram realizadas atividades referentes
> a…". **Padronizar** para a forma acima, igual à da seção 1.

### 4.3 Seção 5.1 — Condições climáticas

**Título oficial:** `Condições climáticas do período:` — a RD-103 usa "Dados climáticos";
**padronizar**.

Classificações (fixo): `Bom`, `Chuvoso`, `Impraticável`.

Dia sem execução *(condicional, um por dia sem ficha de campo)*:
> No dia **{data}** não houve execução de atividades.

Vários dias: consolidar em uma frase —
> Nos dias **{datas}** não houve execução de atividades.

### 4.4 Seção 5.2 — Relatório Fotográfico

Texto introdutório **fixo** (já idêntico nas 4 obras — manter):
> O relatório fotográfico demonstra de forma visual as informações relevantes das
> atividades desenvolvidas ao longo do trecho.

Legenda de figura *(template único para as 4 obras)*:
> Figura **{n}** – Relatório fotográfico – **{servico_canonico}** – **{rodovia}**.

> Hoje há três padrões diferentes. **Padronizar**, incorporando o serviço (hoje só a
> RD-104 faz isso, e é a forma mais informativa).

### 4.5 Seção 8.1 — Não conformidades emitidas pelo Programa

**Título oficial:** `Não conformidades emitidas pelo Programa no período:`

- sem NC:
  > Durante o período de execução de serviços abordado pelo presente relatório, não houve
  > emissão de não conformidades pelo PROG.
- com NC *(template + item por NC)*:
  > Durante o período de execução de serviços abordado pelo presente relatório, foram
  > emitidas **{n}** não conformidades pelo Programa, anexadas a este relatório:
  >
  > - **NC-{numero}**: emitida em **{data}**, durante **{atividade}**, em **{local}**,
  >   na estaca **{estaca}** da **{rodovia}**: **{ocorrencia}**.

> Hoje o texto usa "não conformidades **emitidas** ... **emitidas** pelo Programa" —
> redundância a eliminar. E "uma não conformidade" deve concordar com a quantidade real.

### 4.6 Seção 8.2 — Não conformidades emitidas pela Órgão Contratante

- sem NC:
  > Durante o período de execução de serviços abordado pelo presente relatório, não houve
  > emissão de não conformidades pela Órgão Contratante.

### 4.7 Seção 9 — Relato de riscos e interferências em campo

**Título oficial:** `RELATO DE RISCOS E INTERFERÊNCIAS EM CAMPO` — hoje há três variantes.

- sem ocorrência *(frase única para as 4 obras)*:
  > Durante o período abrangido por este relatório, não foram registrados riscos ou
  > interferências em campo.

> Hoje: "não houve relatos de interferências em campo" (RD-101), "não houve
> interferências em campo" (RD-102A), "não houve relato de interferência" (RD-103),
> "não houve relatos de interferências" (RD-104). **Quatro redações para o mesmo fato.**

### 4.8 Seção 10 — Termo de Encerramento

Template **fixo**, já uniforme nas 4 obras — manter exatamente:
> O Relatório Semanal de Acompanhamento de Obras da **{obra_sem_hifen}** – N° **{numero}**,
> referente à consolidação e à apresentação das atividades desenvolvidas durante o período
> de **{data_inicio}** a **{data_fim}**, no âmbito do Contrato Nº**{contrato}** – Parceiro X
> **{executora_contratual}**, apresenta-se constituído de **{n}** (**{n_extenso}**) páginas.
>
> **{cidade}**, **{data_emissao_extenso}**.

---

## 5. Capitalização

| Elemento | Regra | Exemplo |
|---|---|---|
| Títulos de seção nível 1 | CAIXA ALTA | `ACOMPANHAMENTO AMBIENTAL` |
| Títulos de nível 2 e 3 | Apenas a inicial maiúscula | `Condições climáticas do período:` |
| Nomes de serviço no corpo | minúsculo dentro da frase | "atividades de terraplenagem e drenagem" |
| Nomes de serviço como título/legenda | Iniciais maiúsculas | `Obras de Arte Corrente` |
| Siglas | caixa alta | `OAE`, `SST`, `NC` |
| Meses por extenso | minúsculo | "11 de setembro de 2026" |

> **Não usar dois-pontos no fim de título de seção.** Hoje existem
> `2. FICHA DE ACOMPANHAMENTO:` e `5. EXECUÇÃO DE OBRA:` com dois-pontos, e outros sem.
> **Padronizar: sem dois-pontos.**

---

## 6. Regras para a geração por IA

Quando um texto exigir composição (e não apenas substituição de variável), a IA é
**GROUNDING-ONLY**:

**Pode:** organizar fatos já estruturados em período gramatical correto; escolher
conectivos; ajustar concordância; ordenar por relevância; aplicar este guia.

**Não pode, em nenhuma hipótese, inventar ou inferir:**
fato · data · serviço · local · estaca · coordenada · status · origem · consequência ·
responsável · quantidade · percentual · número de documento.

**Entrada obrigatória:** objeto estruturado, nunca o PDF anterior.

```json
{
  "documento": "NC-022",
  "data": "2026-09-03",
  "rodovia": "RD-101",
  "local": "Ponte sobre o Rio Exemplo",
  "estaca": "1908",
  "servico": "Obras de Arte Especiais",
  "atividade": "perfuração para execução de estacas-raiz",
  "ocorrencia": "carreamento de material proveniente da escavação para o curso hídrico",
  "status": "em aberto"
}
```

**Saída obrigatoriamente verificada** por um *fact checker* que confira, campo a campo,
que todo número, data, nome próprio e sigla do texto gerado existe na entrada.
Se aparecer algo que não está na entrada → rejeitar e reprocessar.

**Ordem de preferência — nunca use IA onde código resolve:**

```
1. determinístico  →  2. extração estruturada  →  3. validação cruzada
→  4. template parametrizado  →  5. IA generativa grounded  →  6. HITL  →  7. manual
```

---

## 7. O que NÃO padronizar

- Dados cadastrais reais de cada obra (contratada, supervisora, contrato, extensão,
  trecho, código SRE, lote).
- Fatos do período (serviços executados, clima, NC, ocorrências).
- Datas de eventos reais (p. ex. a migração do diário para digital: 25/04 na RD-101 e
  21/04 nas demais — **é fato, não estilo**).
- Números e quantitativos.

---

## 8. Correções de redação a aplicar na automação

| # | Hoje | Correção | Tipo |
|---|---|---|---|
| R-01 | "pela Programa" (3 obras) / "pelo Programa" (RD-103) | **pelo Programa** | erro gramatical |
| R-02 | 4 redações de "não houve interferências" | frase única | inconsistência |
| R-03 | `UNIFILAR` / `DIAGRAMA UNIFILAR` | **DIAGRAMA UNIFILAR** | inconsistência |
| R-04 | `Condições` / `Dados` climáticos | **Condições climáticas do período** | inconsistência |
| R-05 | `RELATO DE INTERFERÊNCIA(S) EM CAMPO` (singular × plural) | **manter o título da obra** (`road_registry.titulo_secao_9`); padronizar só singular/plural | inconsistência ⚠️ **revisada — ver abaixo** |
| R-06 | `RD-102 A` na capa | **RD-102A** | erro de grafia |
| R-07 | `35,00km` | **35,00 km** | erro de formatação |
| R-08 | "Terraplanagem" | **Terraplenagem** | erro ortográfico |
| R-09 | Títulos com e sem dois-pontos | sem dois-pontos | inconsistência |
| R-10 | "não conformidades emitidas … emitidas pelo Programa" | eliminar redundância | erro de redação |
| R-11 | 3 padrões de legenda de foto | legenda única com serviço | inconsistência |
| R-12 | Numeração de página só em 2 obras | em todas | inconsistência |

Registro completo, com os erros **factuais**, em
[`ERROS_HERDADOS_E_CORRECOES.md`](ERROS_HERDADOS_E_CORRECOES.md).

---

## ⚠️ R-05 revisada em 2026-09-17 — padronizar forma não é alterar escopo

**A regra anterior estava errada.** Ela mandava padronizar as três variantes do título
da seção 9 para `RELATO DE RISCOS E INTERFERÊNCIAS EM CAMPO`. Isso **acrescenta
"RISCOS"** a três das quatro obras — e "riscos" não é uma variação de forma de
"interferências": é **outro escopo declarado** para a seção.

Os títulos reais no CURRENT:

| Obra | Título |
|---|---|
| RD-101 | `RELATO DE RISCOS E INTERFERÊNCIAS EM CAMPO` |
| RD-102A | `RELATO DE INTERFERÊNCIAS EM CAMPO` |
| RD-103 | `RELATO DE INTERFERÊNCIA EM CAMPO` |
| RD-104 | `RELATO DE INTERFERÊNCIAS EM CAMPO` |

**Regra vigente:** o título vem de `road_registry.titulo_secao_9`, **por obra**.
A padronização se limita a singular × plural (`INTERFERÊNCIA` → `INTERFERÊNCIAS`),
que é forma. **Nunca acrescentar ou remover "RISCOS"** — só a Emmily decide ampliar
ou reduzir o escopo de uma seção.

### Princípio geral, que vale para todas as regras deste guia

> **Padronização corrige forma: grafia, acentuação, singular/plural, maiúsculas,
> pontuação, espaçamento, redundância.**
> **Padronização nunca altera escopo, fato, responsável, data, local ou valor.**

Se uma regra de padronização mudar *o que a seção afirma cobrir*, ela deixou de ser
padronização e virou decisão editorial — e decisão editorial é da Emmily, não do guia.

Revisar as demais regras sob essa luz: `R-03` (`UNIFILAR` → `DIAGRAMA UNIFILAR`) e
`R-04` (`Dados` → `Condições climáticas`) são forma e continuam válidas.
