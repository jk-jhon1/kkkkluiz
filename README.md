# 📚 Livrarias para Baicho — Automação com Google Planilhas + Apps Script

Projeto de livraria virtual integrado ao **Google Planilhas + Google Apps Script**, atendendo a todos os requisitos da atividade **Automação de tarefas — Lógica: Procedimentos**.

---

## ✅ Requisitos atendidos

| Requisito | Implementação (`apps-script.gs` + Planilha) |
| :--- | :--- |
| **Google Planilhas + Apps Script** | Aba `Dados` (tabela `A:G` com `E-MAIL (GMAIL)` e `ENDEREÇO` + Painel de Operações `I1:J7`) e aba `Log` (`A:I` registrando `E-MAIL (GMAIL)` e `ENDEREÇO`) controladas por `apps-script.gs`. |
| **Comandos que executam tarefas automaticamente** | Geração automática de `ID`, validação de dados, proteção contra injeção de fórmulas (`@`), formatação de moeda (`R$`) e data/hora (`ULTIMA ATUALIZAÇÃO`) e auditoria completa na aba `Log` (com Gmail, endereço de entrega e forma de pagamento). |
| **Procedimento para registrar uma operação** | `registrarOperacao` — grava novos pedidos na aba `Dados` e no `Log` (pelo site via `doPost`, pelo Painel `J3:J7` ou por caixas de diálogo). |
| **Procedimento para atualizar informações** | `atualizarInformacoes` — localiza o pedido pelo `ID` e altera nome, e-mail, endereço, livro e/ou valor, atualizando `ULTIMA ATUALIZAÇÃO` e gravando no `Log`. |
| **Procedimento para limpar campos** | `limparCampos` (limpa os campos do Painel `J2:J7` ou esvazia os campos de um pedido por `ID`), além de `limparPainel` e `limparCamposPedido`. |
| **Acionar por botões ou menus personalizados** | `onOpen()` cria automaticamente o menu personalizado **`📚 Automação`** no topo do Google Planilhas, e os mesmos nomes de procedimentos funcionam em botões (**Inserir > Desenho > Atribuir script**). |

---

## 🛠️ Como atualizar o Apps Script na Planilha

1. Abra a planilha vinculada e acesse **Extensões > Apps Script**.
2. Apague todo o conteúdo antigo do editor, cole **todo** o código de [`apps-script.gs`](apps-script.gs) e salve (**Ctrl+S**).
3. **Atualizar a implantação do site (Web App):**
   - No editor do Apps Script, clique em **Implantar > Gerenciar implantações**.
   - Clique no ícone de **lápis (Editar)** > em *Versão*, selecione **Nova versão** > clique em **Implantar**.
   - A URL `/exec` em `script.js` continua exatamente a mesma.
4. Recarregue a planilha (**F5**): o menu personalizado **`📚 Automação`** aparecerá na barra superior.

### Botões de exclusão na planilha (`Dados`)
A aba `Dados` possui **três botões prontos** em `I9:J12` (caixas de clique rápido acionadas por `onEdit` + `onSelectionChange`):
- **`J10` (`Apagar Linha Selecionada`)** → clique primeiro na linha desejada da tabela (`A:G`) e depois marque `J10` para apagar essa linha e reorganizar as linhas de baixo.
- **`J11` (`Apagar Última Linha`)** → apaga somente a última linha com informações na tabela (`A:G`), preservando o cabeçalho e o painel.
- **`J12` (`Apagar Todos os Dados`)** → apaga todos os dados registrados na tabela (`A2:G...`) e limpa `J2:J7`, preservando o cabeçalho e o painel.

Para inserir também **3 botões gráficos flutuantes** clicáveis na aba `Dados` (`I14`, `I17` e `I20`), basta clicar no menu **`📚 Automação` > `🔘 Inserir os 3 botões na planilha`** (ou atribuir manualmente via **Inserir > Desenho > Atribuir script**):
- `apagarLinhaSelecionada`
- `apagarUltimaLinha`
- `apagarTodosOsDados`
- `registrarOperacao`
- `atualizarInformacoes`
- `limparCampos`

---

## 🌐 Executar o site

Abra `index.html` no navegador. O catálogo, o carrinho e o envio automático de pedidos (com nome, Gmail, endereço, forma de pagamento, livros e valor) para a planilha funcionam com `script.js` e `style.css`.
