---
name: session-close
description: Protocolo oficial de ENCERRAMENTO de sessão deste projeto. Para o desenvolvimento novo, reconcilia Git/disco/estado, persiste no repositório tudo que só existe na conversa (STATE.md, checkpoint, decisão, histórico), registra validações realmente executadas, confere a política de commit e para.
when_to_use: Quando a sessão vai acabar — fim de trabalho, contexto perto do limite, ou antes de compactar. Também quando houver conhecimento que não pode se perder se esta conversa sumir.
disable-model-invocation: true
argument-hint: "[observação final, opcional]"
---

# /session-close — encerrar persistindo

**Se esta conversa sumir agora, nenhum conhecimento importante pode sumir junto.**

**Pare o desenvolvimento novo.** Nada de feature, próxima tarefa, refatoração oportunista,
correção não relacionada ou aumento de escopo. Também proibido: descartar trabalho, `stash`
desnecessário, `reset`, apagar, esconder alteração.

E o mais importante: **não maquiar**. Teste não executado não "passou". Investigação não é
implementação. Intenção não é resultado. Hipótese não é fato.

## 1. Onde estamos

`pwd`, branch, HEAD completo e curto, `git status --porcelain`, commits desta sessão
(`git log --oneline` desde o HEAD de abertura).

Compare o persistido (`STATE.md`, checkpoints) contra Git, disco, artefatos e testes.
**A realidade verificável prevalece.** Divergência se mostra, não se apaga.

## 2. A pergunta obrigatória

> Existe alguma informação importante que está **apenas** no contexto desta conversa?

Varra a sessão: o que foi concluído, o que ficou pela metade, o que foi investigado e **não**
implementado, decisões, descobertas, hipóteses levantadas e descartadas, bugs, limitações, riscos,
pendências, bloqueios, perguntas abertas, erro cometido e corrigido no caminho.

Se houver — persista **antes** de encerrar. Essa é a razão de ser desta skill.

Um erro diagnosticado e corrigido vale registro: o raciocínio errado tende a voltar.

## 3. Onde cada coisa vai

Mapa canônico em `CLAUDE.md` → "Protocolo de sessão". Resumo:

| Informação | Destino |
|---|---|
| estado atual, próximo passo | `STATE.md` (**substitui**, não acumula — snapshot, não diário) |
| retomada sem conversa, marco | `docs/checkpoints/CHECKPOINT_<nome>.md` + linha no README de lá |
| histórico da sessão | `docs/checkpoints/SESSOES.md` (uma linha) |
| decisão durável | `docs/decisoes/DEC-NNN-<slug>.md` |
| conhecimento de domínio | `docs/mapeamento/`, `docs/fontes/`, `docs/validacao/` |
| regra machine-readable | `config/*.yaml` |
| regra permanente nova | `CLAUDE.md` |

Não duplique: o que está no checkpoint não precisa ser repetido inteiro no `STATE.md`.
Não crie DEC só para ter documentação — só decisão real, com alternativas e consequências.

Checkpoint **só quando houver marco ou risco de perda**. Sessão pequena: `STATE.md` + linha em
`SESSOES.md` bastam. Ver [references/persistencia.md](references/persistencia.md) para os modelos.

## 4. Validação — só o que couber ao que foi feito

Rode o que o trabalho da sessão exige, não a suíte inteira por hábito.

```bash
for f in tests/test_*.py; do python -X utf8 "$f" >/dev/null 2>&1 || echo "FALHOU: $f"; done
python scripts/validar_relatorio_semanal.py --current
```

Registre **comando exato** e **resultado real**. O que não puder rodar, registre como
*não executado*, com motivo e risco residual. Nunca escreva "testes passaram" sem execução.

## 5. Git — a política manda, não o fim da sessão

`CLAUDE.md`: **commit só quando a Emmily pedir explicitamente.** Terminar a sessão **não** é pedido.

Então: não commite por iniciativa própria. Reporte o que está pendente e pergunte, em uma linha,
se ela quer salvar. Se ela já tiver pedido nesta sessão, siga a skill `versionar`.

Nunca `push`, `reset --hard`, `clean -fd`, reescrita de histórico ou troca de branch.

Ao final reconfirme branch, HEAD e `git status`.

## 6. Próxima sessão

Uma unidade concreta e executável. Não vale "continuar a análise".
Objetivo, critério de pronto, dependências, arquivos de alto sinal, riscos, verificações esperadas.

## 7. Relatório

Monte conforme [references/encerramento.md](references/encerramento.md) e termine exatamente com:

> Handoff concluído. Nenhuma nova implementação foi iniciada.

**PARE.**
