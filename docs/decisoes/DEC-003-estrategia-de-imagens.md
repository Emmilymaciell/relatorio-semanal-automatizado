# DEC-003 — Estratégia de produção das imagens

**Data:** 2026-09-17 · **Situação:** ✅ **decidida (provisória)** · **Revisável após os shadow tests**
**Resolve:** `Q-1` de `docs/mapeamento/14_perguntas_para_emmily.md`

---

## Contexto

O relatório contém elementos que hoje chegam ao Word como **imagem colada** (EMF):
Ficha de Acompanhamento, gráfico de Avanço Global, painéis ambientais, figura de
acompanhamento de notificações, gráficos de NC e o Relatório Fotográfico.

Três caminhos foram levantados na descoberta:

| Opção | O que é | Custo | Risco |
|---|---|---|---|
| A | reproduzir tudo nativamente no DOCX | alto | divergência visual com o histórico |
| B | renderizar tudo a partir das planilhas-fonte | baixo | dependência de Excel/COM, menos controle |
| C | **híbrido** | médio | precisa de critério claro de corte |

---

## Decisão

**HÍBRIDO (opção C)** — decisão inicial do MVP.

**Nativo** (gerado pelo sistema, sem imagem):
tabelas simples · tabela climática · listas · lista de anexos · Termo de Encerramento ·
todo o texto corrido · gráficos simples cuja reprodução seja comprovadamente fiel.

**Renderizado a partir da fonte existente** (imagem produzida da planilha):
Ficha de Acompanhamento (acumulada) · painéis ambientais · figura de notificações ·
componentes complexos do Relatório Fotográfico.

---

## Justificativa

O objetivo do MVP é **fidelidade + rastreabilidade**, não reinventar layout. Reproduzir
nativamente um painel de Power BI ou a ficha acumulada no primeiro teste gastaria o
orçamento da fase em aparência, enquanto o risco real está nos **dados**.

O corte é: **se o elemento carrega dado que o sistema já extrai e sabe provar, vale
nativo; se é um artefato visual de uma ferramenta externa, renderiza-se da fonte.**

---

## Consequências

- O Slice 3 entrega DOCX com placeholder explícito
  (`[FIGURA — AVANÇO FÍSICO — PENDENTE DE RENDERIZAÇÃO]`) — isso é
  **MVP LIMITATION**, nunca conformidade.
- A renderização da planilha exige decidir o mecanismo (Excel COM × LibreOffice ×
  exportação prévia). Fica para o slice de imagens.
- Reavaliar depois do primeiro comparison report: se a similaridade visual do
  nativo for alta, migrar mais elementos para nativo.
