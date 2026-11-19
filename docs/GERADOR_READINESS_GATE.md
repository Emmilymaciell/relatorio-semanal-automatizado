# Readiness Gate — o gerador pode começar?

**Criado em:** 2026-09-17 · **Baseline:** `discovery-v2`
**Pergunta única:** *existe alguma coisa que tecnicamente impede começar a construir o gerador?*

**Resposta: NÃO.** Nenhum bloqueador técnico. Detalhamento na seção 3.

---

## 1. Critérios

Cada componente é avaliado em quatro eixos e recebe um veredito.

| Eixo | Pergunta |
|---|---|
| **Conhecimento suficiente?** | sabemos o que o elemento contém e como ele se comporta? |
| **Fonte conhecida?** | sabemos de onde o dado vem, com evidência? |
| **Regra conhecida?** | sabemos transformar a fonte no elemento final? |
| **Testável?** | existe regra de validação executável sobre relatórios reais? |

| Veredito | Significado |
|---|---|
| `READY` | pode ser implementado hoje, sem decisão humana pendente |
| `READY WITH HITL` | pode ser implementado; exige um ponto de decisão humana no fluxo |
| `PARTIAL` | implementável em parte; falta um detalhe que **não impede começar** |
| `BLOCKED` | não pode ser implementado — falta informação essencial |

---

## 2. Gate por componente

| Componente | Conhecimento | Fonte | Regra | Testável | Veredito | Bloqueio |
|---|---|---|---|---|---|---|
| **Capa** | ✅ | ✅ `road_registry.yaml` | ✅ | ✅ V3, V4 | **READY** | — |
| **Período** | ✅ | ✅ calendário | ✅ sábado→sexta | ✅ V1 (144/144) | **READY** | — |
| **Numeração** | ✅ | ✅ âncora por obra | ✅ `(início−âncora)/7+1` | ✅ V2 (144/144) | **READY** | — |
| **Cadastro da obra** | ✅ | ✅ `road_registry.yaml` | ✅ | ✅ C7 | **READY** | 2 campos a confirmar (supervisora, trecho) — não impede |
| **Índice** | ✅ | ✅ outline | ✅ derivado dos títulos | ✅ V15 | **READY** | — |
| **Avanço físico** | ✅ | ✅ UNIFILAR → `[Programa]` → Gráfico | ✅ cadeia provada 91/91 | ✅ | **READY** | origem é digitação manual — fora do escopo do gerador |
| **Clima** | ✅ | ✅ `FICHA CAMPO!N17/O17/P17` | ✅ X na coluna | ✅ V10 | **READY** | — |
| **Dias sem execução** | ✅ | ✅ ausência de ficha do dia | ✅ | ✅ | **READY** | — |
| **Serviços da semana** | ✅ | ✅ abas do fotográfico + ficha | ✅ taxonomia canônica | ✅ | **READY** | T-17 (OAC/Obras Complementares) — não impede |
| **Fotos — metadados** | ✅ | ✅ overlay visual (30/30) | ✅ transcrição | ✅ 13/13 índice | **READY** | técnica de leitura a definir — não impede |
| **Fotos — seleção** | ✅ | ✅ faixas da Ficha de Campo | ⚠️ critério editorial não registrado | ✅ parcial | **READY WITH HITL** | por desenho, não por falta de informação |
| **Relatório Fotográfico** | ✅ | ✅ uma aba por serviço | ✅ blocos de 3, A3 retrato | ✅ V7 | **READY WITH HITL** | depende da seleção acima |
| **Ambiental — texto** | ✅ | ✅ planilha de licença ambiental | ✅ template + nº do PA | ✅ V12, V13 | **READY** | — |
| **Ambiental — painéis** | ✅ | ✅ planilha (painel 1 confere 8/8) | ⚠️ hoje é screenshot de Power BI | ⚠️ | **PARTIAL** | decidir: reproduzir do dado ou capturar. Não impede as demais seções |
| **Ambiental — nº do PA** | ✅ | ✅ acervo ambiental | ✅ **localizar, nunca calcular** | ✅ | **READY** | contador corporativo não é previsível — por isso a regra é localizar |
| **SST** | ✅ | ✅ acervo ambiental | ✅ anda junto com o ambiental | ✅ V13 (52/52) | **READY** | — |
| **NC — contagens (8.3)** | ✅ | ✅ `VISAO_GERAL!Tbl_VISAO` | ✅ domínio fechado de 6 status | ✅ | **READY** | planilha é estado atual, sem histórico — cortar pela data de fechamento |
| **NC — descrição (8.1)** | ✅ | ✅ `VISAO_GERAL!V` + `H` | ✅ template 3/3 fidelidade 1,000 | ✅ | **READY** | — |
| **Interferências (seção 9)** | ✅ | ✅ Ofício/PA + estaca + tipo | ✅ anatomia de 7 elementos | ⚠️ 10 casos | **READY WITH HITL** | miolo descritivo é decisão humana |
| **Termo de Encerramento** | ✅ | ✅ documento montado | ✅ template + extenso | ✅ V8 (20 erros históricos detectados) | **READY** | — |
| **Anexos** | ✅ | ✅ regras de composição | ✅ ordem + periodicidade | ✅ V11, V12, V14 (52/52) | **READY** | — |
| **Diário de Obras** | ✅ | ✅ PDF do Parceiro | ✅ 7 páginas A2 + exceções | ✅ V5 | **READY** | exceção do diário em atraso já configurada |
| **Paginação** | ✅ | ✅ renderizador | ✅ | ✅ V8 | **READY** | — |
| **Revisão** | ✅ | ✅ estrutura de pastas | ✅ maior revisão na raiz | ✅ REV-1..5 | **READY WITH HITL** | decidir revisar é humano |
| **Renderização de imagens** | ⚠️ | ✅ planilhas-fonte | ⚠️ 3 caminhos possíveis | ❌ | **PARTIAL** | `DEC-003` pendente — decisão de arquitetura, não falta de conhecimento |

### Placar

| Veredito | Componentes |
|---|---|
| `READY` | **17** |
| `READY WITH HITL` | **5** |
| `PARTIAL` | **2** |
| `BLOCKED` | **0** |

---

## 3. Bloqueios reais

### 3.1 BLOQUEADORES — impedem começar

**Nenhum.**

### 3.2 NÃO BLOQUEADORES — precisam de decisão, mas o trabalho começa sem eles

| # | Item | Por que não bloqueia |
|---|---|---|
| 1 | **`DEC-003`** — como produzir as imagens (renderizar Excel, redesenhar nativo ou híbrido) | As etapas 2, 3 e 4 da ordem de implementação não dependem disso. A decisão pode ser tomada depois de o esqueleto textual estar de pé. |
| 2 | **Painéis ambientais** — reproduzir do dado ou capturar do Power BI | Caso particular do item 1. |
| 3 | **Digitação de origem no UNIFILAR** | É entrada humana **anterior** ao gerador. O gerador consome o resultado; não precisa saber de onde a pessoa tirou o número. |
| 4 | **Critério editorial de seleção de foto** | Resolvido por desenho: `HITL`. Não é falta de informação. |
| 5 | **NC da Órgão Contratante sem precedente** | 133/133 negativas. O gerador emite a negativa; se um dia houver, cai em `HITL`. |
| 6 | **Revisão `.3` sem precedente** | A regra de versionamento é genérica em `rev`; `.3` funciona sem mudança. |
| 7 | **T-17** (OAC/Obras Complementares fora da legenda da ficha) | Os serviços do relatório vêm das abas do fotográfico, que já os têm. |
| 8 | **2 campos de cadastro divergentes** (supervisora, trecho) | São 2 valores num YAML. A Emmily confirma e pronto. |

### 3.3 MELHORIAS FUTURAS — não entram nesta fase

| # | Item |
|---|---|
| 1 | Descobrir qual aplicativo gera o overlay das fotos (prioridade baixa declarada) |
| 2 | Corrigir a degradação da cadeia do `[Programa]` (101 de 230 células viraram valor colado) |
| 3 | Reconciliar os dois orçamentos divergentes entre UNIFILAR e `[Programa]` |
| 4 | Sincronizar `Tbl_BASE` com `Tbl_VISAO` na planilha de NC |
| 5 | Padronizar o nome da subpasta de obsoletos (`Obsoleto` vs `Relatório Semanal (Obsoleto)`) |

---

## 4. Pré-condições para a Fase 3

Tudo já disponível no repositório:

| Insumo | Onde |
|---|---|
| Modelo canônico | `config/report_schema.yaml` |
| Versões do formato | `config/report_versions.yaml` (CURRENT = `v3`) |
| Cadastro das obras | `config/road_registry.yaml` |
| Taxonomia de serviços | `config/services_taxonomy.yaml` |
| Regras de redação | `docs/mapeamento/GUIA_DE_REDACAO.md` |
| Textos fixos e templates | `docs/mapeamento/TEXTOS_E_TEMPLATES.md` |
| Linhagem das fontes | `docs/mapeamento/08_data_lineage.md` + `docs/fontes/*.md` |
| Erros a **não** reproduzir | `docs/mapeamento/ERROS_HERDADOS_E_CORRECOES.md` |
| Validador | `scripts/validar_relatorio_semanal.py` (15 regras) |
| Testes | `tests/` (4 suítes) |

---

## 5. Veredito

> **O gerador pode começar.** O caminho
> `fonte → modelo canônico → regra → validação → renderização`
> está documentado ponta a ponta, com evidência citável, para todos os componentes
> exceto a renderização de imagens — que é uma **decisão de arquitetura pendente**,
> não uma lacuna de conhecimento.
>
> Começar pelos componentes `READY` que não dependem de imagem entrega valor
> verificável antes de enfrentar a parte cara.
