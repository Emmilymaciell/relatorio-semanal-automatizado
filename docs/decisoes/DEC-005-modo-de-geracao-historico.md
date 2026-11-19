# DEC-005 — Modo de geração: `CURRENT` × `HISTORICAL_SHADOW`

**Data:** 2026-10-04 · **Status:** ACEITA · **Fase:** 3.4 (v0.3)

## Contexto

Reconstruir uma semana passada não é o mesmo que gerar o relatório da semana corrente. Para a
semana corrente, as fontes são contemporâneas do fato relatado. Para uma semana que já teve
relatório emitido, várias fontes **já mudaram**, e o estado que elas tinham na data de emissão não
é recuperável.

O Final Candidate v0.2 da RD-103 (semana 49, emitida em 11/09/2026) afirmou no corpo do relatório:

> "o empreendimento possui 3 licenças emitidas e 3 licenças pendentes"
> "há 33 condicionantes próximas do vencimento"

contra **4 e 5** e **1** do relatório oficial.

## Evidência

Recalculando as **mesmas 45 condicionantes** da RD-103 na planilha-fonte:

| `hoje` usado no recálculo | Finalizada | Próx. vencimento | No prazo | Pendente |
|---|---:|---:|---:|---:|
| 2026-09-11 (data de emissão do oficial) | 11 | **31** | 3 | 0 |
| 2026-09-21 | 11 | **2** | 1 | 31 |

Nem usando a data de emissão do próprio relatório oficial o número bate, e ele oscila de 31 para 2
em dez dias — porque "próximo do vencimento" é uma **função da data de hoje** (`Condicionantes!G`:
`<=0` Pendente, `<=30` Próximo do vencimento, senão No prazo) aplicada sobre um acervo que também
mudou.

Confiança: **CONFIRMADO** (medido em `scripts/gerador/figuras/ambiental.py` com as duas datas).

O erro não era o número. Era **emitir um número volátil como se fosse fato histórico apurado**.

## Decisão

1. Toda geração declara um modo: `CURRENT` ou `HISTORICAL_SHADOW`. O modo efetivo é gravado em
   `report_data.metadata.generation_mode` — nunca re-inferido na leitura.
2. `scripts/shadow_report.py candidate` usa `HISTORICAL_SHADOW` por padrão: este script **é** o
   gerador do shadow, que reconstrói semanas já emitidas.
3. Os dados que não se reconstroem estão declarados, **com evidência**, em
   `config/generation_modes.yaml` — não se decide "isso é volátil" dentro do `content_builder`.
4. Em `HISTORICAL_SHADOW`, a afirmação específica **não é emitida**. O parágrafo descritivo, que
   não afirma quantidade, permanece; no lugar da oração suprimida entra um bloco `LIMITATION`
   dizendo que aquele estado não foi reconstruído e por quê.
5. **Jamais copiar o número do golden.** O relatório oficial é referência de comparação, nunca
   fonte (regra do Shadow Report, `CLAUDE.md`).
6. Divergência causada por qualquer item dessa lista é `SOURCE_LIMITATION`, nunca `SHADOW_ERROR`:
   o extrator está certo; a fonte é que não guarda o passado.

## Alternativas consideradas

| Alternativa | Por que não |
|---|---|
| Recalcular com a data de emissão do relatório antigo | Não resolve: dá 31 contra 1 do oficial. O acervo mudou junto com a fórmula. |
| Emitir o número atual com uma ressalva textual | Continua afirmando um número que não descreve a semana; a ressalva não desfaz a afirmação. |
| Copiar os números do golden | Proibido — transformaria o golden em fonte e destruiria o valor do shadow test. |
| Omitir o parágrafo inteiro | Perde texto que é legítimo e padrão do corpus; só a oração final é insegura. |
| Snapshot histórico das planilhas | Não existe. Não há versionamento das fontes naquelas datas. |

## Consequências

- O `LIMITATION` da RD-103 subiu de 3 para 5 na v0.3. **Isso é melhora, não piora** — duas
  afirmações sem lastro viraram limitações declaradas, e o `READY` não caiu (102 nas duas versões).
- Como o regression guard lê "LIMITATION só desce", o movimento precisou de exceção declarada em
  `config/regression_exceptions.yaml`, que exige motivo **e** evidência e reporta `JUSTIFICADO` —
  nunca `OK` silencioso.
- Em `CURRENT`, a redação volta ao normal: o modo histórico não pode emburrecer o relatório da
  semana corrente. Há teste para isso (`M5` em `tests/test_generation_mode.py`).
- A comparação com o golden passa a ter mais `SOURCE_LIMITATION` e menos `SHADOW_ERROR`. É o
  retrato honesto: reduz erro real sem maquiar divergência legítima.

## Restrições

Esta decisão **não** autoriza omitir dado que seja reconstruível. A lista de
`dados_nao_reconstruiveis_historicamente` é fechada e cada entrada carrega a evidência de por que
o estado histórico não é recuperável. Acrescentar item exige a mesma prova.

## Verificação

`tests/test_generation_mode.py` — M1 a M6: os dois modos existem; os dados estão declarados com
evidência; em `HISTORICAL_SHADOW` nada é afirmado; o substituto é `LIMITATION` com motivo; em
`CURRENT` a redação normal volta; e **nenhum número do golden vaza** para o candidate.
