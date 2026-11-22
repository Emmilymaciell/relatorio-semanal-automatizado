# Sanitização desta versão pública

Cópia sanitizada de um projeto em uso real. Nada aqui é dado de cliente ou de terceiro: o que aparece como exemplo é inventado. Detalhe completo da sanitização: relatório interno da autora.

## O que foi removido

`output/` (inventários, relatórios de comparação e evidências reais), amostra real, checkpoints e engenharia reversa do acervo, logotipos do programa e do parceiro, arquivos de proveniência.

## O que foi trocado

Rodovias → `RD-1xx`/`RD-2xx`; programa/órgãos/parceiro/executoras → nomes genéricos; identificadores de código com o nome do programa → `prog_*`; contratos → números fictícios. Mantidos: a forma das frases-modelo (`config/textos_v3.yaml`, sem valores do período) e as datas de contrato/OS usadas pelo cálculo da numeração de semanas.

## O que é fictício

Logotipos são placeholders gerados; os testes que precisariam do acervo real são pulados com aviso explícito (`tests/_dados_reais.py`).

## Como verificar

```bat
for %f in (tests\test_*.py) do py -3.12 %f     :: cada suíte é um script; test_docx_to_pdf exige Word
```

Varreduras (2026-11-19): gitleaks `dir` só com falsos positivos justificados; busca por caminho local, e-mail, nomes do cliente e links privados sem ocorrências reais.
