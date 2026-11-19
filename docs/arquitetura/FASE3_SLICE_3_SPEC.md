# Fase 3 — Slice 3 — Esqueleto textual em DOCX + comparação

**Baseline:** Checkpoint 3.1 (`024cc02`) · **Criado em:** 2026-09-17
**Meta:** `output/shadow/RD103_2026-08-29_2026-09-04/shadow_report.docx` +
`comparison.json` + `comparison_report.md`.
**Fora de escopo:** PDF final, anexos montados, renderização real das imagens, pixel-perfect.

---

## 0. Regras que continuam valendo

1. SharePoint **READ-ONLY absoluto**; `prog_util.lp()` / `assert_read_only()`.
2. Não escrever `.py` com heredoc do Bash.
3. **Zero IA / zero LLM** no código gerado.
4. **Nada de `if obra == "..."`.** Diferenças em `config/`, `road_registry`, `taxonomy`, `schema`.
5. **Proibido usar o golden como fonte.** `assert_nao_e_golden()` continua ativa.
   A **única** camada autorizada a ler o golden é o comparador (`compare.py`),
   e só para classificar diferenças.
6. Tudo em `output/shadow/{OBRA}_{inicio}_{fim}/`. Nada publicado, nada sincronizado.

---

## 1. Decisões da Emmily a implementar

### Q-8 — clima de dia sem Ficha de Campo → `INDETERMINADO`

**Não** assumir default "Bom" nem repetição do dia anterior. Sem ficha e sem outra
fonte comprovada:

```json
{"date": "2026-08-30", "clima": null, "clima_label": "INDETERMINADO",
 "_status": "HITL", "validation_status": "HITL",
 "motivo": "sem Ficha de Campo e sem fonte alternativa comprovada"}
```

No DOCX a célula pode exibir o texto que o padrão do relatório usaria, **mas marcada**
(ver §5: placeholder/realce de HITL). Internamente permanece não comprovada.

No comparador, esse campo classifica como **`HITL_REQUIRED`** (categoria
`UNRESOLVED_SOURCE`), **nunca** `SHADOW_ERROR`.

> ✅ **Ratificado 2026-09-17 — dois campos distintos, não um.**
> `_status` = resultado da **extração** (`OK` / `NOT_FOUND` / `UNREADABLE` / …).
> `validation_status` = prontidão para o **documento** (`READY` / `HITL` / `LIMITATION`).
> São coisas diferentes e devem continuar separadas. O content builder e o renderer
> leem **`validation_status`**, não `_status`.

> ✅ **T-13 resolvido (`docs/mapeamento/T13_CLIMA_DIA_SEM_FICHA.md`): não existe regra.**
> No regime vigente, 23 casos discriminantes se repartem em três comportamentos —
> repetir o dia anterior (11), escrever `-` (6), escrever `Bom` (6). Nenhum domina.
> **Renderização adotada:** a célula do dia não comprovado recebe **`-`**, grafia que
> existe no acervo, não afirma fato falso, e o bloco vai marcado `HITL`.

### Q-9 — nomenclatura dos serviços → completa, formal, técnica

Fonte da grafia: `config/services_taxonomy.yaml` (`canonical_name`).
Regra de redação:

- **primeira ocorrência no documento:** `Obras de Arte Especiais (OAE)`
- **ocorrências seguintes:** `OAE`
- quando não houver sigla registrada na taxonomia, usar só o nome canônico.

A sigla vem de um campo da taxonomia (`sigla`), **não** de heurística sobre o texto.
Se a taxonomia não tiver `sigla` para uma entrada, acrescentar o campo — é configuração.

O comparador classifica `golden="OAE"` × `shadow="Obras de Arte Especiais (OAE)"`
como **`SEMANTICALLY_EQUIVALENT`**. Não é erro.

### Q-10 — seção 9 sem interferência → gerar, mas `TPL + HITL`

Quando nenhuma ocorrência for encontrada nas fontes consultadas:

```json
{"interferences": {"occurrences": [],
  "sources_consulted": [...],
  "text": "<frase negativa padrão da obra, de road_registry.titulo/texto>",
  "automation_class": "TPL",
  "validation_status": "HITL",
  "motivo": "ausência nas fontes não prova operacionalmente que não houve interferência"}}
```

Gera o texto, **não bloqueia** a montagem, e fica marcado para confirmação humana.
`automation_class` é `TPL`, **não** `DET`.

### Q-11 — associação do Parecer (PA) à semana → **por conteúdo**

Pipeline obrigatório, em `scripts/gerador/extratores/parecer.py`:

```
candidatos: todos os PA-NNN da pasta `6. Acompanhamento ambiental e SST\Parecer\` da obra
    ↓ ler o TEXTO do PDF (pdfplumber/PyMuPDF)
    ↓ extrair: rodovia mencionada · período/data · objeto
    ↓ casar rodovia com a obra E período com o período do relatório
    ↓
0 compatíveis  → NOT_FOUND
1 compatível   → MATCH
>1 compatíveis → HITL (listar todos)
```

**Proibido:** maior número, PA mais recente por `mtime`, PA anterior + 1.
A numeração é corporativa e não é previsível.

**Teste de aceitação:** para RD-103 / 2026-08-29..09-04 o mecanismo tem de chegar em
**PA-254** sozinho. O golden **não** pode participar dessa decisão — o teste apenas
confere o resultado no final.

---

## 2. §6 — estado histórico das NC

Antes de qualquer coisa: **não manipular a planilha** para simular 11/09.

Investigação pontual e barata: verificar se os **documentos individuais de NC**
(`3. Notificações\Not. Geradas\{BV,Programa,Supervisora}` e `Respostas\`) permitem
reconstruir data de emissão, mudanças de status, encerramento e evidência.

- **Se sim:** implementar a reconstrução e registrar a proveniência.
- **Se não:**

```json
{"historical_nc_state": "UNRECONSTRUCTABLE_FROM_CURRENT_SOURCE",
 "motivo": "Tbl_VISAO usa TODAY(); os documentos individuais não registram a linha do tempo de status",
 "impacto": "history_by_status reflete o estado atual, não o da semana"}
```

Isso é **limitação explícita**, não bug. **Nunca** preencher com o valor do golden.
Para relatórios novos o problema não existe — a leitura é contemporânea.

---

## 3. §7 — campos contratuais antes do DOCX

O primeiro DOCX **não pode** depender do golden para os dados contratuais.

1. Verificar o que `config/road_registry.yaml` já traz validado por obra.
2. Para o que faltar (RD-102A, RD-103, RD-104: instrumento de convocação, data de
   assinatura do contrato, ordem de serviço e data), fazer busca **pontual** em
   `{obra}\2_Documentacao Licitatoria\3_Contrato Executora\` e afins.
   **Busca pontual, não varredura ampla.**
3. Atualizar o cadastro mestre com o valor e a **origem** (arquivo + página/trecho).
4. O gerador lê do cadastro. Proveniência = o documento contratual que originou o registro.

Campo que continuar sem comprovação: fica `NOT_FOUND` + `HITL`, e o DOCX mostra o
marcador de pendência. **Não copiar do golden.**

> ⚠️ **Regra descoberta em 2026-09-17 — `ordem_servico_em` é a data da ASSINATURA DIGITAL.**
> A OS da RD-103 traz no corpo "26 de agosto de 2025", mas o PDF só foi gerado e assinado
> em **29/08/2025** (`creationDate D:20250829144957` + 3 carimbos `/Sig`). O próprio
> documento diz que o prazo corre "a partir da assinatura deste documento".
> Nas outras 3 obras as duas datas coincidem (RD101 19/09, RD102A 27/08, RD104 29/09) —
> a RD-103 é a única divergente, e por isso o erro passaria despercebido.
> **Aqui o golden estava certo e a extração estava errada.** Corrigido no cadastro.
> Qualquer campo de data vindo de PDF assinado deve preferir o carimbo de assinatura à
> data digitada no corpo, registrando a divergência quando houver.

---

## 4. Arquitetura — conteúdo separado de renderização

```
report_data.json
      ↓  content_builder.py     aplica textos fixos, templates e regras de redação
document_model (dataclasses)    estrutura abstrata, sem Word, sem caminho de arquivo
      ↓  renderers/docx_renderer.py
shadow_report.docx
```

**O content builder e o document model não conhecem caminho de arquivo nem planilha.**
Recebem `report.progress.completed`, `report.services`, `report.climate`… e pronto.
Trocar DOCX por PDF ou por outro template **não** pode exigir mexer em extrator.

### Novos módulos

```
scripts/gerador/
    document_model.py      dataclasses do modelo abstrato
    content_builder.py     report_data -> document_model
    textos.py              catálogo de textos fixos e templates (ou config/textos.yaml)
    renderers/
        __init__.py
        docx_renderer.py   document_model -> .docx (python-docx)
    compare.py             shadow x golden -> comparison.json + comparison_report.md
    extratores/parecer.py  Q-11
config/
    textos_v3.yaml         textos fixos e templates da estrutura CURRENT (v3)
    comparison_rules.yaml  casos conhecidos e sua classificação
```

### `document_model.py` — contrato

```python
@dataclass
class Block:
    kind: str          # "paragraph" | "heading" | "table" | "figure_placeholder"
                       # | "caption" | "list" | "toc" | "page_break" | "cover_field"
    status: str        # "READY" | "PLACEHOLDER" | "HITL" | "LIMITATION"
    provenance: list   # chaves de report_provenance.json que sustentam este bloco
    ...                # campos específicos por kind

@dataclass
class Section:
    id: str            # "s5_1_clima"
    number: str        # "5.1"
    title: str
    level: int
    page_setup: str    # "A4 retrato" | "A3 retrato" | ...
    blocks: list[Block]

@dataclass
class Document:
    meta: dict
    sections: list[Section]
```

Serializar o `document_model` em `output/shadow/.../document_model.json` — é o que
permite auditar "o que ia para o Word" sem abrir o Word.

---

## 5. O DOCX

Estrutura **CURRENT / v3 completa**, na ordem das 11 seções invariantes
(`config/report_schema.yaml`), incluindo capa, índice e Termo de Encerramento.

Conteúdo que deve sair pronto: capa · identificação · dados contratuais · índice ·
período · número · textos fixos · textos parametrizados · serviços · clima · avanço
(número, não figura) · NC (8.1, 8.2, 8.3 textuais) · ambiental textual ·
seção 9 · lista de anexos.

**Placeholders de imagem** — ocupando a posição estrutural correta:

```
[FIGURA 1 — FICHA DE CAMPO ACUMULADA — RENDERIZAÇÃO PENDENTE]
[GRÁFICO 1 — AVANÇO FÍSICO — RENDERIZAÇÃO PENDENTE]
[RELATÓRIO FOTOGRÁFICO — RENDERIZAÇÃO PENDENTE]
[FIGURA n — <legenda> — RENDERIZAÇÃO PENDENTE]
```

**Marcação de HITL** — todo bloco com `status != READY` recebe marca visual discreta
e inequívoca (ex.: realce ou prefixo `⟨HITL⟩`), para a Emmily enxergar de imediato o
que não está comprovado. Nada de pendência disfarçada de conteúdo pronto.

Estilo: fonte, tamanhos e hierarquia de títulos **básicos e coerentes**.
**Não** gastar tempo com margem de 2 mm, posição exata de imagem ou microespaçamento.

### De onde vêm os textos fixos — regra de origem

Os textos de **classe A** (boilerplate: a frase contratual da Apresentação, as
definições de Bom/Chuvoso/Impraticável, o parágrafo do Diário de Obras, o texto
explicativo da 8.3…) são **template do processo**, não dado da semana. Hoje eles se
propagam porque cada semana é uma cópia da anterior — não existe `.dotx`.

Regra para `config/textos_v3.yaml`:

1. O texto fixo é derivado do **corpus da descoberta** (`docs/mapeamento/TEXTOS_E_TEMPLATES.md`
   e `output/extracao/texto/`), **excluindo o relatório golden que está sendo comparado**
   (`RO_049.1 - RD-103`). Cada entrada registra de quantos relatórios foi confirmada.
2. Isso é **configuração**, como o `road_registry` — entra no repositório, é revisável,
   e não é relido a cada semana.
3. O que **nunca** pode vir daí: qualquer valor específico do período — data, número,
   percentual, estaca, NC, serviço, contagem, nº de páginas. Esses vêm sempre de
   `report_data.json`.

A guarda `assert_nao_e_golden()` continua valendo em tempo de execução: o gerador
não abre relatório nenhum.

### Termo de Encerramento (§14)

Implementar a lógica completa (número → número por extenso → frase), mas enquanto
houver placeholder, marcar explicitamente:

```
status: PLACEHOLDER
n_paginas_provisorio: true
nota: "paginação provisória — o Termo definitivo é gerado no Checkpoint 3.3, após a montagem final"
```

**Nunca** copiar as 35 páginas do golden.

---

## 6. Comparação (§15, §16)

`compare.py` → `comparison.json` + `comparison_report.md`, em 5 eixos:
`metadata`, `section`, `text`, `data`, `structure`.

Classificação de **cada** item — o comparador tem de responder *por que* é diferente:

| Classe | Quando |
|---|---|
| `EXACT` | idêntico |
| `SEMANTICALLY_EQUIVALENT` | mesmo fato, redação diferente (ex.: `OAE` × `Obras de Arte Especiais (OAE)`) |
| `EXPECTED_DIFFERENCE` | diferença deliberada desta iteração (placeholder, Termo provisório) |
| `GOLDEN_DEFECT` | a fonte sustenta o Shadow e o golden está errado — **com evidência citada** |
| `SHADOW_ERROR` | o Shadow está errado. **É o único que conta como falha do sistema** |
| `MISSING` | existe no golden, ausente no Shadow |
| `EXTRA` | existe no Shadow, ausente no golden |
| `HITL_REQUIRED` | não comprovado pela fonte; exige confirmação humana (ex.: clima de 30/08) |
| `SOURCE_LIMITATION` | a fonte não permite reproduzir (ex.: `TODAY()` das NC) |

Casos já conhecidos vão para `config/comparison_rules.yaml`, **com evidência**, e não
podem ser "regra cega": a regra declara a condição, e o comparador só aplica a classe
se a condição se verificar nos dados.

Caso âncora obrigatório:

```yaml
- id: clima-2026-09-01
  campo: climate.by_day[2026-09-01]
  classe: GOLDEN_DEFECT
  condicao: "shadow == 'Chuvoso' and golden == 'Bom'"
  evidencia: "2026-09-01 Ficha de Campo RD 103.xlsx, aba FICHA CAMPO, X em O17 (Chuvoso);
              mtime 2026-09-02, anterior à emissão do relatório em 2026-09-11"
```

O `comparison_report.md` traz a contagem por classe e, separadamente:
**(a)** erros reais do Shadow, **(b)** diferenças corretas, **(c)** placeholders.

Um score alto **nunca** pode esconder valor factual errado, anexo errado, NC errada
ou avanço errado.

---

## 7. Validador (§17)

Rodar `scripts/validar_relatorio_semanal.py` sobre o que for possível. Se ele só
aceitar PDF, fazer **o mínimo** para validar o primeiro output disponível — não
transformar isso em projeto paralelo. WARNING decorrente de placeholder deliberado
tem de sair rotulado **`MVP LIMITATION`**, nunca como conformidade.

---

## 8. Testes exigidos

1. `test_parecer_por_conteudo.py` — chega em PA-254 sozinho; `0/1/>1` candidatos
   produzem `NOT_FOUND`/`MATCH`/`HITL`.
2. `test_content_builder.py` — 11 seções na ordem; regra de 1ª ocorrência dos
   serviços; frase de dia sem execução; clima `INDETERMINADO` marcado `HITL`.
3. `test_docx_renderer.py` — o `.docx` abre; nº de seções; títulos na ordem;
   todo placeholder presente; nenhum bloco `HITL` sem marca.
4. `test_comparison.py` — o caso `clima-2026-09-01` classifica `GOLDEN_DEFECT`;
   `OAE` × `Obras de Arte Especiais (OAE)` classifica `SEMANTICALLY_EQUIVALENT`;
   nenhuma diferença cai em `SHADOW_ERROR` por falta de regra (default é
   `HITL_REQUIRED` com motivo, não silêncio).
5. `test_road_registry_contratual.py` — os campos contratuais usados no DOCX vêm do
   cadastro, e cada um tem origem registrada.

Mais as 9 suítes já existentes.

---

## 9. Definição de pronto (Checkpoint 3.2)

- [ ] Q-8, Q-9, Q-10, Q-11 implementadas conforme §1
- [ ] §6 respondido (reconstruível ou `UNRECONSTRUCTABLE_FROM_CURRENT_SOURCE`)
- [ ] §7 resolvido no cadastro mestre, sem golden
- [ ] `shadow_report.docx` gerado, estrutura v3 completa
- [ ] `document_model.json` serializado
- [ ] `comparison.json` + `comparison_report.md` com as 9 classes
- [ ] 14 suítes passando (9 antigas + 5 novas)
- [ ] `sources`/`data`/`build` rodam nas 4 obras sem exceção
- [ ] nada fora de `output/shadow/`
