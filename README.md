# 📚 Livrarias para Baicho — Automação com Google Planilhas + Apps Script

Projeto de livraria virtual integrado ao **Google Planilhas + Google Apps Script**, atendendo a todos os requisitos da atividade **Automação de tarefas — Lógica: Procedimentos**.

---

## ✅ Requisitos atendidos

| Requisito | Implementação (`apps-script.gs` + Planilha) |
| :--- | :--- |
| **Google Planilhas + Apps Script** | Aba `Dados` (tabela `A:E` + Painel de Operações `G1:H5`) e aba `Log` (`A:H`) controladas por `apps-script.gs`. |
| **Comandos que executam tarefas automaticamente** | Geração automática de `ID`, validação de dados, proteção contra injeção de fórmulas (`@`), formatação de moeda (`R$`) e data/hora (`ULTIMA ATUALIZAÇÃO`) e auditoria automática na aba `Log`. |
| **Procedimento para registrar uma operação** | `registrarOperacao` — grava novos pedidos na aba `Dados` (pelo site via `doPost`, pelo Painel `H3:H5` ou por caixas de diálogo). |
| **Procedimento para atualizar informações** | `atualizarInformacoes` — localiza o pedido pelo `ID` e altera nome, livro e/ou valor, atualizando `ULTIMA ATUALIZAÇÃO` e gravando no `Log`. |
| **Procedimento para limpar campos** | `limparCampos` (limpa os campos do Painel `H2:H5` ou esvazia os campos de um pedido por `ID`), além de `limparPainel` e `limparCamposPedido`. |
| **Acionar por botões ou menus personalizados** | `onOpen()` cria automaticamente o menu personalizado **`📚 Automação`** no topo do Google Planilhas, e os mesmos nomes de procedimentos funcionam em botões (**Inserir > Desenho > Atribuir script**). |

---

## 🛠️ Como atualizar o Apps Script na Planilha

1. Abra a planilha vinculada e acesse **Extensões > Apps Script**.
2. Apague todo o conteúdo antigo do editor, cole **todo** o código de [`apps-script.gs`](apps-script.gs) e salve (**Ctrl+S**).
3. Recarregue a planilha (**F5**): o menu personalizado **`📚 Automação`** aparecerá na barra superior (na primeira execução o Google pedirá permissão do script).
4. **Atualizar a implantação do site (Web App):**
   - No editor do Apps Script, clique em **Implantar > Gerenciar implantações**.
   - Clique no ícone de **lápis (Editar)** > em *Versão*, selecione **Nova versão** > clique em **Implantar**.
   - A URL `/exec` em `script.js` continua a mesma.

### Opcional: criar botões visuais na planilha
Além do menu **`📚 Automação`**, você também pode criar botões na aba `Dados`:
1. Clique em **Inserir > Desenho**, desenhe um botão e clique em **Salvar e fechar**.
2. Clique nos **3 pontinhos** no canto do botão > **Atribuir script** e informe um dos nomes abaixo (sem parênteses):
   - `registrarOperacao`
   - `atualizarInformacoes`
   - `limparCampos`

---

## 🌐 Executar o site

Abra `index.html` no navegador. O catálogo, o carrinho e o envio automático de pedidos para a planilha funcionam com `script.js` e `style.css`.
