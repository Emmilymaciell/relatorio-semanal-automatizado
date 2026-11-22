# Fase 3 — Slices 1 e 2 — Especificação de implementação

**Baseline:** `discovery-v2` (HEAD `a5bbdd6`) · **Criado em:** 2026-09-17
**Escopo desta spec:** Source Package (Slice 1) + Canonical Report Model + Provenance (Slice 2).
**Fora de escopo:** DOCX, PDF, imagens, comparação com o golden. Isso é Slice 3+.

---

## 0. Regras invioláveis

1. `C:\caminho\exemplo\PastaSincronizada\` é **READ-ONLY absoluto**.
   Todo acesso passa por `prog_util.lp()`; toda escrita passa por `prog_util.assert_read_only()`.
2. **MAX_PATH**: nunca usar `open()`/`os.walk()` cru em caminho do SharePoint.
3. **Não escrever `.py` com heredoc do Bash** (corrompe `\`). Usar Write/Edit.
4. **Zero LLM / zero IA** nesta fase. Só `DET`, `VAL`, `TPL`, `HITL`.
5. **Nada de `if obra == "RD103"`.** Toda diferença entre obras mora em
   `config/*.yaml`. O código é genérico.
6. **Proibido ler o relatório oficial para preencher dados.** Ver §5.

---

## 1. Caso de teste

| | |
|---|---|
| Obra | `RD103` (rótulo `RD-103`) |
| Período | `2026-08-29` (sábado) a `2026-09-04` (sexta) |
| Nº do relatório | **49** (conferir pela regra, não copiar) |
| Golden reference | `…\2026.08.29-2026.09.04 49º Relatório Semanal\RO_049.1 - RD-103.pdf` (35 páginas) |

O golden é **apenas referência de comparação futura**. Nesta spec ele só pode ser
**registrado** no manifest (caminho + sha1 + nº de páginas), nunca lido para extrair dado.

---

## 2. Layout real das fontes (já verificado — não redescobrir)

Raízes:

```
SHAREPOINT      = C:\caminho\exemplo\PastaSincronizada
BASE_OBRAS      = {SHAREPOINT}\Programa - 3. Desenvolvimento\G2 -  Engenharia     (dois espaços em "G2 -  Engenharia")
COMPARTILHADO   = {SHAREPOINT}\Programa - 2. Compartilhado interno\03_Gestao dos contratos
```

Por obra:

```
{BASE_OBRAS}\{pasta_obra}\3_Gestao Tecnica de Obra\
    3. Notificações\
        Not. Geradas\{BV|Programa|Supervisora}\
    4. Relat*  Semanal*  Obras\                     <-- NOME VARIA POR OBRA, resolver por glob
        1. Fichas de campo\
            2026-09-04 Ficha de Campo RD 103.xlsx   <-- uma por dia COM execução
        2. Ficha Acompanhamento Semanal (Obra)\     <-- NÃO alimenta o v3 (P-08). Só registrar.
        3. Fotos\
            26-09-04\{Terraplenagem|Pavimentação|Drenagem|OAE|...}\*.jpg
            (padrões observados: `NN-NN-NN` e `NN_NN_NN` — aceitar os dois)
        4. Diário de Obras\
            Diário de obra digital assinado\2026-08-29 a 2026-09-04 - RD-103.pdf
            Diário de obra físico - Atualizado\
        5. Relat* Semanal de Acompanhamento de Obras\
            2026.08.29-2026.09.04 49º Relatório Semanal\
                RO_049.1 - RD-103.docx      <-- GOLDEN, proibido como fonte
                RO_049.1 - RD-103.pdf       <-- GOLDEN, proibido como fonte
                Arquivos Auxiliares\        <-- nome varia: "Arquivos auxiliares" nas outras obras
                    Ficha de Campo Acumulada RD-103.xlsx
                    Gráfico - Avanço Global.xlsx
                    RELATORIO FOTOGRAFICO 49 - 103.xlsx
                    UNIFILAR 103.xlsx
                    [Programa] Avanço Global de Obra.xlsx
        Memorial de campo\
    6. Acompanhamento ambiental e SST\
        2026-09-02\                                  <-- pasta por data de inspeção
            Ficha de acompanhamento ambiental\FA-016.1 - Ficha acompanhamento ambiental - Executora C.pdf
            Ficha de acompanhamento SST\FA-015.1 - Ficha acompanhamento SST - Executora C.pdf
        Parecer\                                     <-- PA-NNN
```

Corporativas (uma só, compartilhada pelas 4 obras):

```
{BASE_OBRAS}\02_Acompanhamento de NC automatizada.xlsm        aba VISAO_GERAL, tabela Tbl_VISAO
{BASE_OBRAS}\04_Acompanhamento - Licenca Ambiental.xlsx
{COMPARTILHADO}\UNIFILAR (com não conformidades)\UNIFILAR 103.xlsx
{COMPARTILHADO}\[Programa] Avanço Global de Obra.xlsx             (mestre)
```

> ⚠️ `UNIFILAR OBSOLETO\` e `Avanço Global Mensal (OBSOLETO)\` **devem ser ignorados**.
> Regra geral: descartar qualquer caminho cujo nome normalizado contenha `obsoleto`.

**Fatos já confirmados para esta semana** (usar como teste de aceitação, não como atalho):

- Fichas de campo diárias presentes: `08-29, 08-31, 09-01, 09-02, 09-03, 09-04` → **6 de 7**.
  **Falta `2026-08-30` (domingo)** → gera a frase "No dia 30/08/2026 não houve execução de atividades."
- Pastas de foto presentes: `26-08-29, 26-08-31, 26-09-01, 26-09-02, 26-09-03, 26-09-04` — **coerente**
  com as fichas (dupla confirmação da regra `s5_1_clima.frase_sem_execucao`).
- Diário do período: `2026-08-29 a 2026-09-04 - RD-103.pdf`.
- Ambiental do período: pasta `2026-09-02`, com FA-016.1 (ambiental) e FA-015.1 (SST).

---

## 3. Arquitetura a construir

```
config/
    source_layout.yaml          NOVO — onde mora cada fonte, por padrão + override por obra
scripts/gerador/
    __init__.py
    config.py                   carrega e valida os 5 YAMLs; cache
    periodo.py                  período, nº do relatório, dias, rótulos de dia
    provenance.py               coletor de proveniência
    source_discovery.py         Slice 1 → manifest
    canonical.py                Slice 2 → report_data
    extratores/
        __init__.py             registry de extratores
        ficha_campo.py          clima (N16/O16/P16 + marca X na 17) e linhas de serviço/estaca
        ficha_acumulada.py      ficha acumulada (data de corte, tabela de estacas)
        avanco.py               [Programa] Avanço Global (C3, C7) e Gráfico - Avanço Global
        nc.py                   VISAO_GERAL!Tbl_VISAO → NC do período + histórico por status
        ambiental.py            04_Acompanhamento - Licenca Ambiental + pasta do período
        fotos.py                inventário de fotos por dia/serviço + vínculo com a ficha
        anexos.py               anexos fixos + condicionais, na ordem
scripts/shadow_report.py        CLI
tests/
    test_periodo_numeracao.py
    test_source_discovery.py
    test_canonical_rd103.py
    test_provenance.py
    test_sem_vazamento_golden.py
```

### CLI

```bash
python scripts/shadow_report.py sources --obra RD103 --inicio 2026-08-29 --fim 2026-09-04
python scripts/shadow_report.py data    --obra RD103 --inicio 2026-08-29 --fim 2026-09-04
```

- `--fim` é opcional: se ausente, `inicio + 6`.
- `--inicio` também aceita `--semana 49` (resolve pela âncora do `road_registry`).
- Saída sempre em `output/shadow/{OBRA}_{inicio}_{fim}/`.
- Nunca escrever fora de `output/` e `tmp/`.

### Saídas do Slice 1 — `manifest.yaml`

```yaml
report:
  road: RD-103
  road_key: RD103
  period_start: 2026-08-29
  period_end: 2026-09-04
  report_number: 49
  revision: 1
  schema_version: v3           # de report_versions.yaml, pela data
sources:
  field_records:      [{date, path, exists, size, sha1, mtime}]     # 1 por dia; ausentes marcados
  field_records_missing_days: [2026-08-30]
  accumulated_sheet:  {...}
  photos:             [{date, service_folder, files: [...], n}]
  progress_unifilar:  {...}    # UNIFILAR 103.xlsx (compartilhado, não obsoleto)
  prog_progress:      {...}    # [Programa] Avanço Global de Obra.xlsx (o da pasta da semana)
  progress_chart:     {...}    # Gráfico - Avanço Global.xlsx
  photo_workbook:     {...}    # RELATORIO FOTOGRAFICO 49 - 103.xlsx
  nc_tracker:         {...}    # 02_Acompanhamento de NC automatizada.xlsm
  nc_documents:       [...]    # notificações emitidas no período
  environmental:      {...}    # planilha + pasta do período + fichas FA-*
  daily_reports:      {...}    # diário de obras do período
  meetings:           [...]    # pasta de foto chamada "Reunião" / ata, se houver
  annexes:            [...]    # candidatos a anexo, já classificados
golden_reference:
  docx: {path, sha1}
  pdf:  {path, sha1, n_paginas}
  aviso: "REFERENCIA DE COMPARACAO — PROIBIDO COMO FONTE"
discovery:
  resolved_at: ...
  warnings: [...]
  not_found: [...]
```

Toda fonte que não for encontrada entra em `not_found` com o motivo — **nunca inventar caminho**.

### Saídas do Slice 2

`report_data.json` — o modelo canônico. Chaves conforme `config/report_schema.yaml`:

```
metadata, road, contract, period, report_number, services, climate, progress,
field_records, photos, nonconformities, environment, sst, meetings, risks,
interferences, annexes, closing_term
```

Regras de preenchimento (todas já provadas na discovery):

| Campo | Fonte | Método |
|---|---|---|
| `report_number` | `road_registry.ancora_semana_1` | `(inicio - ancora).days // 7 + 1` |
| `period.days[].label` | calendário | sábado…sexta, caixa alta |
| `climate.by_day[]` | Ficha de Campo diária, aba `FICHA CAMPO`, `N16/O16/P16` rotulam Bom/Chuvoso/Impraticável; a **linha 17** traz o `X` | leitura direta |
| `climate.days_without_work[]` | dias do período **sem** ficha de campo | derivado |
| `field_records[]` | Ficha de Campo **diária**, aba `FICHA CAMPO`: cabeçalho na **linha 21**, dados a partir da **22** — `B,C,D`=estaca inicial, `E,F,G`=final, `H`=lado, `J`=extensão, `K`=serviço (código), `M`=situação `A/C/I`, `O`=foto, `Q`=descrição, `X`=observação | leitura direta + `services_taxonomy` |

> ⚠️ **Correção 2026-09-17.** A primeira versão desta spec dava linha 24/25 e as colunas
> `I/J/L/M/O/AC` para a ficha diária. Isso é o layout da **Ficha de Campo ACUMULADA**,
> não da diária — as duas planilhas têm layouts diferentes e a spec as confundiu.
> Verificado célula a célula em `2026-09-04 Ficha de Campo RD 103.xlsx`.
> O extrator localiza o cabeçalho **por texto normalizado**, não por letra fixa.
| `services[]` | códigos distintos da coluna `J` + pastas de foto do período, normalizados por `services_taxonomy` | derivado; **é o dado que alimenta 3 pontos do documento** |
| `progress.global_concluido` | `[Programa] Avanço Global de Obra.xlsx`, aba `Avanço Global - RD103`, célula `C3` | leitura direta |
| `progress.global_concluido_atacado` | mesma aba, `C7` | leitura direta |
| `progress.series` | `Gráfico - Avanço Global.xlsx`, abas `Acumulado` e `Relatório Semanal` | leitura direta |
| `nonconformities.issued_in_period[]` | `VISAO_GERAL!Tbl_VISAO`: filtrar coluna `Rodovia` = obra, `Data de Abertura` (col. H) dentro do período | leitura direta |
| `nonconformities.issued_in_period[].sentence` | col. `V` (Irregularidade): 1ª sentença, sem a cláusula temporal, inicial minúscula, prefixada `NC-NNN Emitida em {dd/mm/aaaa},` | template já validado (fidelidade 1,000 em 3/3) |
| `nonconformities.history_by_status` | `Tbl_VISAO` col. `Status`, domínio de 6 valores de `LISTAS_VALIDACAO!I1:I7`, filtrado por rodovia | contagem. **Registrar o alerta: a planilha usa `TODAY()`, é estado atual, não histórico da semana** |
| `nonconformities.orgao` | sempre "não houve" (133/133) | constante, com `confidence: DESCONHECIDO_SE_HOUVER` |
| `environment.applies` | ciclo **quinzenal**, âncora `2026-06-06` | derivado |
| `environment.*` | `04_Acompanhamento - Licenca Ambiental.xlsx` + pasta datada | leitura direta |
| `photos[]` | inventário das pastas `26-MM-DD\{serviço}\*.jpg` | listagem. **Não fazer OCR do overlay nesta iteração** — registrar `overlay_read: false` e `status: READY_WITH_HITL` |
| `annexes[]` | regras de `report_schema.yaml` §`s11_anexos` | derivado + localização real dos PDFs |
| `closing_term` | `null` nesta fase — depende da paginação final | placeholder explícito |

`report_provenance.json` — **paralelo**, mesma estrutura de chaves em caminho pontilhado:

```json
{
  "progress.global_concluido": {
    "value": 0.4231,
    "source_file": "…\\[Programa] Avanço Global de Obra.xlsx",
    "sheet": "Avanço Global - RD103",
    "cell": "C3",
    "method": "direct extraction",
    "confidence": "CONFIRMADO POR EVIDÊNCIA DIRETA",
    "extracted_at": "…"
  }
}
```

Todo valor de `progress`, `climate`, `nonconformities`, `report_number`, `services`
e `annexes` **precisa** ter entrada de proveniência. Campo derivado usa
`"method": "derived"` e lista `"inputs": [...]`.

---

## 4. Tratamento de erro — obrigatório

Nunca abortar por uma fonte faltando. Cada fonte resolve para um de:

| Estado | Significado |
|---|---|
| `OK` | encontrada e lida |
| `NOT_FOUND` | não existe no caminho esperado |
| `UNREADABLE` | existe mas não hidrata (`ERROR_CLOUD_FILE_INVALID_REQUEST` / `OSError 22`) — ver T-12 |
| `NOT_APPLICABLE` | não se aplica a este período (ex.: ambiental fora do ciclo) |
| `MVP_LIMITATION` | deliberadamente não implementado nesta iteração |

Tudo que não for `OK` vai para `discovery.warnings` **e** para o campo correspondente
do `report_data.json` como `null` com um irmão `_status`.

---

## 5. Guarda anti-vazamento do golden — obrigatório

Implementar em `scripts/gerador/__init__.py` (ou `config.py`) uma função:

```python
def assert_nao_e_golden(path): ...
```

que levanta `PermissionError` se o caminho casar com o padrão de relatório oficial
(`RO_\d{3}\.\d+ - .*\.(docx|pdf)$`, ou qualquer arquivo na raiz de uma pasta-semana
`\d{4}\.\d{2}\.\d{2}-\d{4}\.\d{2}\.\d{2}.*Relat`). **Todo extrator chama essa guarda
antes de abrir o arquivo.** O `source_discovery` é a única parte autorizada a
registrar o caminho do golden — e só registra metadados (caminho, sha1, nº de páginas).

`tests/test_sem_vazamento_golden.py` prova que a guarda dispara.

---

## 6. Testes exigidos

Não basta "o script roda". Testar **valores**:

1. `test_periodo_numeracao.py` — nº 49 para RD103/2026-08-29; conferir também
   RD101, RD102A e RD104 contra o `relatorios_semanais_historico.json`
   (já existe em `output/inventarios/`) — **pelo menos 20 semanas por obra**.
2. `test_source_discovery.py` — para RD103 nesta semana: 6 fichas de campo,
   `2026-08-30` em `field_records_missing_days`, 5 auxiliares encontrados,
   diário `2026-08-29 a 2026-09-04 - RD-103.pdf` localizado, nenhum caminho
   contendo `obsoleto`, e o golden **fora** de `sources`.
3. `test_canonical_rd103.py` — `report_data.json` tem as 18 chaves de topo;
   `climate.by_day` tem 7 entradas; `services` não vazio e normalizado pela taxonomia;
   `closing_term is None`.
4. `test_provenance.py` — todo valor de `progress.*`, `climate.by_day[*]`,
   `report_number` e `nonconformities.*` tem entrada em `report_provenance.json`.
5. `test_sem_vazamento_golden.py` — a guarda dispara para o `.docx` e o `.pdf` do golden.

E as 4 suítes já existentes precisam continuar passando.

---

## 7. Generalidade

O mesmo comando tem de funcionar para as outras 3 obras sem alteração de código.
Como prova barata, `shadow_report.py sources` deve rodar sem exceção para
**RD101, RD102A e RD104** em uma semana v3 qualquer, mesmo que alguma fonte
apareça como `NOT_FOUND` (as outras obras têm variações de nome de pasta —
`4. Relatório Semanal de Obras`, `Arquivos auxiliares` minúsculo, etc.).
Essas variações moram em `config/source_layout.yaml`, resolvidas por glob
normalizado, **não** por `if`.

---

## 8. Definição de pronto (Checkpoint 3.1)

- [ ] `config/source_layout.yaml` criado e validado
- [ ] `manifest.yaml` gerado para RD-103 / 2026-08-29..09-04 com todo caminho apontando para arquivo real
- [ ] `report_data.json` preenchido
- [ ] `report_provenance.json` cobrindo os campos críticos
- [ ] 5 testes novos passando + 4 antigos passando
- [ ] `sources` roda nas 4 obras sem exceção
- [ ] Nenhuma escrita fora do repositório local
