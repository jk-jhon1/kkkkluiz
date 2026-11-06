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

### Botões de exclusão na planilha (`Dados`)
A aba `Dados` possui **dois botões prontos** em `G7:H9` (caixas de clique rápido acionadas por `onEdit`):
- **`H8` (`Apagar Última Linha`)** → apaga somente a última linha com informações na tabela (`A:E`), preservando o cabeçalho e o painel.
- **`H9` (`Apagar Todos os Dados`)** → apaga todos os dados registrados na tabela (`A2:E...`) e limpa `H2:H5`, preservando o cabeçalho e o painel.

Para inserir também **2 botões gráficos flutuantes** clicáveis na aba `Dados` (`G11` e `G14`), basta clicar no menu **`📚 Automação` > `🔘 Inserir os 2 botões na planilha`** (ou atribuir manualmente via **Inserir > Desenho > Atribuir script**):
- `apagarUltimaLinha`
- `apagarTodosOsDados`
- `registrarOperacao`
- `atualizarInformacoes`
- `limparCampos`

---

## 🌐 Executar o site

Abra `index.html` no navegador. O catálogo, o carrinho e o envio automático de pedidos para a planilha funcionam com `script.js` e `style.css`.
