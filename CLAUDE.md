# Programa — Engenharia Reversa do Relatório Semanal de Acompanhamento de Obras

Instruções para qualquer sessão de Claude Code que trabalhe neste repositório.

---

## ⛔ REGRA Nº 1 — O SHAREPOINT É READ-ONLY

```
C:\caminho\exemplo\PastaSincronizada\
```

Tudo abaixo deste caminho é **somente leitura, sem exceção**.

É **proibido**, dentro dessa árvore:

- excluir, renomear ou mover arquivos;
- alterar PDF, Word, Excel ou qualquer outro arquivo;
- sobrescrever arquivos;
- **criar** arquivos ou pastas;
- reorganizar diretórios;
- qualquer operação que sincronize alteração para o SharePoint via OneDrive.

Toda escrita — scripts, índices, cache, conversões, imagens extraídas, JSON,
Markdown, logs, outputs — acontece **exclusivamente dentro deste repositório**.

Para manipular um documento: copie para `tmp/` com
`prog_util.copiar_para_tmp()` ou trabalhe em memória. Nunca sobre o original.

O módulo `scripts/prog_util.py` expõe `assert_read_only(path)`, que levanta
`PermissionError` se um caminho de escrita cair dentro do SharePoint.
**Use-o em qualquer script que escreva arquivos.**

---

## ⚠️ REGRA Nº 2 — MAX_PATH

Os caminhos do SharePoint passam de 260 caracteres. `open()`, `os.stat()` e
`os.walk()` falham com `WinError 3` neles.

**Sempre** acesse arquivos do SharePoint através de `prog_util.lp(path)`, que
aplica o prefixo `\\?\` (extended-length path), ou dos wrappers do módulo
(`listdir`, `walk`, `isdir`, `exists`, `stat_info`, `copiar_para_tmp`, `sha1`).

Em PowerShell, use `-LiteralPath`.

---

## ⚠️ REGRA Nº 3 — heredoc do Bash corrompe backslash

O tool Bash deste ambiente **colapsa `\\` em `\`** dentro de heredocs, o que
quebra regex e caminhos Windows em arquivos Python.

**Escreva arquivos `.py` com a ferramenta Write ou Edit, nunca com
`cat > arquivo.py <<'EOF'`.** Heredoc só para conteúdo sem backslash.

---

## Objetivo do projeto

Fazer a **engenharia reversa** do processo de produção dos Relatórios Semanais
de Acompanhamento de Obras do Programa — descobrir, a partir dos documentos
existentes, como eles são feitos hoje, de onde vem cada dado, quais regras
existem e o que pode ser automatizado.

**Ordem obrigatória:** `ANALISAR → MAPEAR → VALIDAR → AUTOMATIZAR`.

🚫 **Não construir o gerador final (`gerar_relatorio.py`) nesta fase.**
A descoberta precisa ser revisada pela Emmily antes de qualquer implementação.

---

## Obras

| Pasta    | Rótulo   |
|----------|----------|
| `RD101`  | RD-101   |
| `RD102A` | RD-102A  |
| `RD103`  | RD-103   |
| `RD104`  | RD-104   |

Raiz: `Programa - 3. Desenvolvimento\G2 -  Engenharia\` (atenção: **dois espaços** em `G2 -  Engenharia`).

Relatórios em: `<OBRA>\3_Gestao Tecnica de Obra\4. Relatório Semanal (de) Obras\5. Relatório Semanal de Acompanhamento de Obras\`

---

## Níveis de confiança (usar em toda afirmação do mapeamento)

| Nível | Significado |
|---|---|
| `CONFIRMADO` | Verificado diretamente no arquivo, com evidência citada (caminho + página/célula + valor). |
| `ALTA CONFIANÇA` | Múltiplas evidências convergentes, sem contraexemplo. |
| `PROVÁVEL` | Uma evidência boa, ou padrão consistente em parte da amostra. |
| `HIPÓTESE` | Plausível, ainda não verificado. |
| `DESCONHECIDO` | Não descoberto. **Registrar como desconhecido é preferível a inventar.** |

Correlação não é origem. Nome parecido não é prova. Toda inferência importante
vai para o *evidence ledger* (`docs/mapeamento/08_data_lineage.md` e
`output/evidencias/`).

---

## Versionamento

Commits **apenas quando a Emmily pedir explicitamente** ("salva essa versão",
"commita", "atualiza o git", `/versionar`). Exceção já autorizada no briefing
inicial deste projeto: commits pequenos e semanticamente claros ao longo da
fase de descoberta.

Nunca: `push`, `reset --hard`, `clean -fd`, reescrita de histórico, troca de branch.

---

## Convenções

- Idioma dos documentos e commits: **português do Brasil**.
- Scripts em `scripts/`, numerados por ordem de execução (`01_`, `02_`, ...).
- Documentação viva em `docs/mapeamento/`.
- Conhecimento machine-readable em `config/*.yaml`.
- Estado da investigação em `STATE.md` — manter atualizado.

---

## Protocolo de sessão

**A conversa é descartável; o repositório é a fonte canônica.** Uma sessão precisa poder terminar
por completo e outra começar sem histórico nenhum — outra janela, outra conta, outro computador,
dias depois, depois de compactar.

| Comando | Faz | Para em |
|---|---|---|
| `/session-start` | reconstrói o contexto lendo só o repositório, reconcilia com Git e disco, mostra divergências | briefing — **nenhuma alteração**, aguardando autorização |
| `/session-close` | para o desenvolvimento novo, persiste o que só existe na conversa, registra validações reais | relatório — **nenhuma implementação nova** |

`.claude/skills/session-start/` e `.claude/skills/session-close/`.

### Onde cada informação mora

Cada arquivo tem **uma** função. Não transformar nenhum deles em depósito geral.

| Informação | Destino |
|---|---|
| regras permanentes | este `CLAUDE.md` |
| estado atual e próximo passo | `STATE.md` — snapshot, substitui; nunca diário |
| retomada sem conversa, marco | `docs/checkpoints/CHECKPOINT_*.md` (+ índice no README de lá) |
| histórico de sessões | `docs/checkpoints/SESSOES.md` — uma linha por sessão |
| decisão durável | `docs/decisoes/DEC-NNN-*.md` |
| conhecimento de domínio | `docs/mapeamento/`, `docs/fontes/`, `docs/validacao/` |
| regra machine-readable | `config/*.yaml` |
| histórico de alterações | Git |

**Evidência vence prosa.** Quando `STATE.md` e o repositório discordarem, o repositório está certo
— e a divergência se mostra, não se corrige em silêncio.
