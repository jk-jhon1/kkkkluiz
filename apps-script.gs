/**
 * Los Nanichos – Automação de Tarefas (Lógica: Procedimentos)
 * Ferramenta: Google Planilhas + Google Apps Script
 *
 * Como instalar/atualizar na planilha:
 * 1. Na planilha Google ("TRABALHO INDIO"), clique em: Extensões > Apps Script.
 * 2. Apague todo o código antigo do editor, cole TODO este arquivo e salve (Ctrl+S).
 * 3. Recarregue a planilha (F5): o menu personalizado "📚 Automação" aparecerá no topo!
 * 4. Para manter o site funcionando com a versão nova:
 *    No Apps Script, clique em Implantar > Gerenciar implantações > ícone de lápis (Editar) >
 *    Versão: "Nova versão" > Implantar. (A URL /exec usada no site permanece a mesma.)
 * 5. Opcional — Acionar também por Botões na planilha:
 *    Inserir > Desenho > desenhe um botão > Salvar e fechar > clique nos 3 pontinhos do botão >
 *    "Atribuir script" e digite um dos nomes abaixo (sem parênteses):
 *      - registrarOperacao
 *      - atualizarInformacoes
 *      - limparCampos
 *
 * Estrutura da aba "Dados":
 *   Tabela de pedidos (colunas A:E):
 *     A: ID | B: NOME_DO_CLIENTE | C: LIVRO | D: VALOR | E: ULTIMA ATUALIZAÇÃO
 *   Painel de Operações / Formulário (colunas G:H):
 *     H2: ID (para atualizar/limpar) | H3: NOME DO CLIENTE | H4: LIVRO | H5: VALOR (R$)
 *
 * Estrutura da aba "Log":
 *   A: DATA/HORA | B: USUARIO | C: AÇÃO | D: ID | E: NOME DO CLIENTE | F: LIVRO | G: VALOR | H: OBSERVAÇÃO
 */

const ID_PLANILHA = '1txnYqaemiW6WD_eE-t6JQiMpvzom3WaHUw11vDPf2WM';

// Células do Painel de Operações na aba "Dados"
const CELULA_ID = 'H2';
const CELULA_NOME = 'H3';
const CELULA_LIVRO = 'H4';
const CELULA_VALOR = 'H5';
const INTERVALO_PAINEL = 'H2:H5';

// ===============================================================
// MENU PERSONALIZADO (acionado automaticamente ao abrir a planilha)
// ===============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 Automação')
    .addItem('➕ Registrar operação', 'registrarOperacao')
    .addItem('✏️ Atualizar informações', 'atualizarInformacoes')
    .addItem('🧹 Limpar campos', 'limparCampos')
    .addSeparator()
    .addItem('🔍 Carregar pedido no painel (por ID)', 'carregarPedidoNoPainel')
    .addItem('🗑️ Limpar campos de um pedido (por ID)', 'limparCamposPedido')
    .addItem('🧹 Limpar apenas o painel (H2:H5)', 'limparPainel')
    .addSeparator()
    .addItem('⚠️ Limpar todos os pedidos', 'limparTudo')
    .addItem('🛠️ Configurar estrutura da planilha', 'configurarPlanilha')
    .addToUi();
}

// ===============================================================
// FUNÇÕES AUXILIARES DE PLANILHA, VALIDAÇÃO E LOG
// ===============================================================
function obterPlanilha() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(ID_PLANILHA);
}

function abaPedidos() {
  const planilha = obterPlanilha();
  return planilha.getSheetByName('Dados') || planilha.getSheets()[0];
}

function abaLog() {
  const planilha = obterPlanilha();
  let log = planilha.getSheetByName('Log');
  if (!log) {
    log = planilha.insertSheet('Log');
    log.getRange(1, 1, 1, 8).setValues([[
      'DATA/HORA',
      'USUARIO',
      'AÇÃO',
      'ID',
      'NOME DO CLIENTE',
      'LIVRO',
      'VALOR',
      'OBSERVAÇÃO'
    ]]);
  }
  return log;
}

function obterUsuario(origem) {
  if (origem === 'site') return 'Site (Web App)';
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'Usuário da planilha';
  } catch (e) {
    return 'Usuário da planilha';
  }
}

function registrarLog(acao, id, nome, livro, valor, observacao, origem) {
  try {
    const log = abaLog();
    const linha = log.getLastRow() + 1;
    const valorCelula = valor === '' || valor === undefined || valor === null || isNaN(Number(valor))
      ? ''
      : Number(valor);

    log.getRange(linha, 5, 1, 2).setNumberFormat('@');
    log.getRange(linha, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
    if (valorCelula !== '') {
      log.getRange(linha, 7).setNumberFormat('R$ #,##0.00');
    }

    log.getRange(linha, 1, 1, 8).setValues([[
      new Date(),
      obterUsuario(origem),
      acao,
      id || '-',
      nome || '-',
      livro || '-',
      valorCelula,
      observacao || ''
    ]]);
  } catch (erro) {
    // Falha no log não deve impedir a operação principal
    Logger.log('Aviso ao gravar log: ' + erro);
  }
}

function textoSeguro(valor, max) {
  return String(valor || '').trim().slice(0, max);
}

function converterValor(entrada) {
  if (typeof entrada === 'number') return entrada;
  let texto = String(entrada === undefined || entrada === null ? '' : entrada).trim();
  if (!texto) return NaN;
  texto = texto.replace(/[R$\s]/gi, '');
  if (texto.includes(',')) {
    texto = texto.replace(/\./g, '').replace(',', '.');
  }
  return Number(texto);
}

// Lê apenas as colunas A:E (ignora o Painel de Operações em G:H).
function estadoTabelaPedidos(aba) {
  const maxLinha = aba.getLastRow();
  if (maxLinha < 2) {
    return { ultimaLinha: 1, proximoId: 1 };
  }

  const dados = aba.getRange(2, 1, maxLinha - 1, 5).getValues();
  let ultimaLinha = 1;
  let maiorId = 0;

  for (let i = 0; i < dados.length; i++) {
    const temConteudo = dados[i].some(celula => celula !== '' && celula !== null);
    if (temConteudo) {
      ultimaLinha = i + 2;
    }
    const idNum = Number(dados[i][0]);
    if (dados[i][0] !== '' && !isNaN(idNum) && idNum > maiorId) {
      maiorId = idNum;
    }
  }

  return { ultimaLinha: ultimaLinha, proximoId: maiorId + 1 };
}

// Acha o número da linha do pedido com este ID nas colunas A:E (ou 0 se não existir).
function acharLinha(aba, id) {
  const idAlvo = Number(id);
  if (!idAlvo || isNaN(idAlvo) || idAlvo <= 0) return 0;

  const maxLinha = aba.getLastRow();
  if (maxLinha < 2) return 0;

  const ids = aba.getRange(2, 1, maxLinha - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] !== '' && Number(ids[i][0]) === idAlvo) {
      return i + 2; // +2: linha 1 é cabeçalho e o índice começa em 0
    }
  }
  return 0;
}

function formatarLinhaPedido(aba, linha) {
  const intervalo = aba.getRange(linha, 1, 1, 5);
  aba.getRange(linha, 2, 1, 2).setNumberFormat('@');
  aba.getRange(linha, 4).setNumberFormat('R$ #,##0.00');
  aba.getRange(linha, 5).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  aba.getRange(linha, 1).setHorizontalAlignment('center');
  aba.getRange(linha, 4, 1, 2).setHorizontalAlignment('center');
  intervalo.setBorders(true, true, true, true, true, true);
}

// ===============================================================
// 1) PROCEDIMENTO PARA REGISTRAR UMA OPERAÇÃO
//    Pode ser chamado pelo site/código: registrarOperacao(nome, livro, valor)
//    OU sem argumentos por botão / menu personalizado na planilha!
// ===============================================================
function registrarOperacao(nome, livro, valor, origem) {
  if (arguments.length === 0) {
    return acionarRegistrarOperacao();
  }

  const nomeLimpo = textoSeguro(nome, 80);
  const livroLimpo = textoSeguro(livro, 300);
  const valorNum = converterValor(valor);

  if (!nomeLimpo || !livroLimpo || isNaN(valorNum) || !(valorNum >= 0)) {
    return { ok: false, erro: 'dados inválidos' };
  }

  const aba = abaPedidos();
  const estado = estadoTabelaPedidos(aba);
  const linha = estado.ultimaLinha + 1;
  const id = estado.proximoId;
  const agora = new Date();

  formatarLinhaPedido(aba, linha);
  aba.getRange(linha, 1, 1, 5).setValues([[id, nomeLimpo, livroLimpo, valorNum, agora]]);

  const ehSite = origem === 'site';
  registrarLog(
    ehSite ? 'Registro pelo Site' : 'Registro Manual',
    id,
    nomeLimpo,
    livroLimpo,
    valorNum,
    ehSite ? 'Pedido recebido automaticamente pelo site' : 'Operação registrada via menu/botão',
    origem
  );

  return { ok: true, id: id };
}

function acionarRegistrarOperacao() {
  const ui = SpreadsheetApp.getUi();
  const aba = abaPedidos();

  let nome = textoSeguro(aba.getRange(CELULA_NOME).getValue(), 80);
  let livro = textoSeguro(aba.getRange(CELULA_LIVRO).getValue(), 300);
  let valorTexto = String(aba.getRange(CELULA_VALOR).getValue() ?? '').trim();

  const usouPainel = Boolean(nome || livro || valorTexto);

  // Se o Painel (H3:H5) estiver vazio, pergunta os dados em janelas interativas
  if (!usouPainel) {
    const respNome = ui.prompt('➕ Registrar Operação (1/3)', 'Informe o nome do cliente:', ui.ButtonSet.OK_CANCEL);
    if (respNome.getSelectedButton() !== ui.Button.OK) return;
    nome = textoSeguro(respNome.getResponseText(), 80);

    const respLivro = ui.prompt('➕ Registrar Operação (2/3)', 'Informe o(s) livro(s):', ui.ButtonSet.OK_CANCEL);
    if (respLivro.getSelectedButton() !== ui.Button.OK) return;
    livro = textoSeguro(respLivro.getResponseText(), 300);

    const respValor = ui.prompt('➕ Registrar Operação (3/3)', 'Informe o valor total (ex.: 49,90):', ui.ButtonSet.OK_CANCEL);
    if (respValor.getSelectedButton() !== ui.Button.OK) return;
    valorTexto = respValor.getResponseText();
  }

  const valorNum = converterValor(valorTexto);
  if (!nome || !livro || isNaN(valorNum) || !(valorNum >= 0)) {
    ui.alert(
      'Dados inválidos',
      'Preencha Nome do Cliente (H3), Livro (H4) e um Valor válido >= 0 (H5) antes de registrar.',
      ui.ButtonSet.OK
    );
    return;
  }

  const resultado = registrarOperacao(nome, livro, valorNum, 'planilha');
  if (!resultado.ok) {
    ui.alert('Erro ao registrar', resultado.erro, ui.ButtonSet.OK);
    return;
  }

  aba.getRange(INTERVALO_PAINEL).clearContent();
  obterPlanilha().toast('Pedido #' + resultado.id + ' registrado com sucesso!', '📚 Automação', 5);
  ui.alert('Sucesso!', 'Operação registrada com sucesso (ID #' + resultado.id + ').', ui.ButtonSet.OK);
}

// ===============================================================
// 2) PROCEDIMENTO PARA ATUALIZAR INFORMAÇÕES
//    Pode ser chamado via código: atualizarInformacoes(id, { nome, livro, valor })
//    OU sem argumentos por botão / menu personalizado na planilha!
// ===============================================================
function atualizarInformacoes(id, campos) {
  if (arguments.length === 0) {
    return acionarAtualizarInformacoes();
  }

  if (!campos || typeof campos !== 'object') {
    return { ok: false, erro: 'campos inválidos' };
  }

  const aba = abaPedidos();
  const linha = acharLinha(aba, id);
  if (!linha) return { ok: false, erro: 'pedido não encontrado' };

  if (campos.nome !== undefined) {
    const novoNome = textoSeguro(campos.nome, 80);
    if (!novoNome) return { ok: false, erro: 'nome inválido' };
    const c = aba.getRange(linha, 2);
    c.setNumberFormat('@');
    c.setValue(novoNome);
  }

  if (campos.livro !== undefined) {
    const novoLivro = textoSeguro(campos.livro, 300);
    if (!novoLivro) return { ok: false, erro: 'livro inválido' };
    const c = aba.getRange(linha, 3);
    c.setNumberFormat('@');
    c.setValue(novoLivro);
  }

  if (campos.valor !== undefined) {
    const valorNum = converterValor(campos.valor);
    if (isNaN(valorNum) || !(valorNum >= 0)) return { ok: false, erro: 'valor inválido' };
    const c = aba.getRange(linha, 4);
    c.setNumberFormat('R$ #,##0.00');
    c.setValue(valorNum);
  }

  aba.getRange(linha, 5)
    .setNumberFormat('dd/mm/yyyy hh:mm:ss')
    .setValue(new Date());

  const atualizado = aba.getRange(linha, 1, 1, 4).getValues()[0];
  registrarLog(
    'Atualização',
    atualizado[0],
    atualizado[1],
    atualizado[2],
    atualizado[3],
    'Informações atualizadas via menu/botão',
    'planilha'
  );

  return { ok: true, id: Number(id) };
}

function acionarAtualizarInformacoes() {
  const ui = SpreadsheetApp.getUi();
  const aba = abaPedidos();

  let idTexto = String(aba.getRange(CELULA_ID).getValue() ?? '').trim();
  const nomePainel = textoSeguro(aba.getRange(CELULA_NOME).getValue(), 80);
  const livroPainel = textoSeguro(aba.getRange(CELULA_LIVRO).getValue(), 300);
  const valorPainelTexto = String(aba.getRange(CELULA_VALOR).getValue() ?? '').trim();

  // Caso 1: usuário preencheu o ID (H2) e pelo menos um campo novo no Painel (H3:H5)
  if (idTexto && (nomePainel || livroPainel || valorPainelTexto)) {
    const camposPainel = {};
    if (nomePainel) camposPainel.nome = nomePainel;
    if (livroPainel) camposPainel.livro = livroPainel;
    if (valorPainelTexto) {
      const v = converterValor(valorPainelTexto);
      if (isNaN(v) || !(v >= 0)) {
        ui.alert('Valor inválido', 'Informe um valor numérico válido em H5.', ui.ButtonSet.OK);
        return;
      }
      camposPainel.valor = v;
    }

    const resPainel = atualizarInformacoes(Number(idTexto), camposPainel);
    if (!resPainel.ok) {
      ui.alert('Não foi possível atualizar', 'Erro: ' + resPainel.erro, ui.ButtonSet.OK);
      return;
    }

    aba.getRange(INTERVALO_PAINEL).clearContent();
    obterPlanilha().toast('Pedido #' + resPainel.id + ' atualizado!', '📚 Automação', 5);
    ui.alert('Sucesso!', 'Informações do pedido #' + resPainel.id + ' atualizadas com sucesso.', ui.ButtonSet.OK);
    return;
  }

  // Caso 2: fluxo guiado por janelas (pede o ID se não estiver em H2)
  if (!idTexto) {
    const respId = ui.prompt('✏️ Atualizar Informações', 'Informe o ID do pedido que deseja atualizar:', ui.ButtonSet.OK_CANCEL);
    if (respId.getSelectedButton() !== ui.Button.OK) return;
    idTexto = respId.getResponseText().trim();
  }

  const linha = acharLinha(aba, idTexto);
  if (!linha) {
    ui.alert('Pedido não encontrado', 'Não existe pedido com o ID "' + idTexto + '" na aba Dados.', ui.ButtonSet.OK);
    return;
  }

  const atual = aba.getRange(linha, 1, 1, 4).getValues()[0];
  const campos = {};

  const respNome = ui.prompt(
    '✏️ Atualizar Pedido #' + atual[0] + ' (1/3)',
    'Nome atual: ' + (atual[1] || '(vazio)') + '\nDigite o novo nome ou deixe em branco para manter:',
    ui.ButtonSet.OK_CANCEL
  );
  if (respNome.getSelectedButton() !== ui.Button.OK) return;
  if (respNome.getResponseText().trim()) {
    campos.nome = respNome.getResponseText().trim();
  }

  const respLivro = ui.prompt(
    '✏️ Atualizar Pedido #' + atual[0] + ' (2/3)',
    'Livro atual: ' + (atual[2] || '(vazio)') + '\nDigite o novo livro ou deixe em branco para manter:',
    ui.ButtonSet.OK_CANCEL
  );
  if (respLivro.getSelectedButton() !== ui.Button.OK) return;
  if (respLivro.getResponseText().trim()) {
    campos.livro = respLivro.getResponseText().trim();
  }

  const respValor = ui.prompt(
    '✏️ Atualizar Pedido #' + atual[0] + ' (3/3)',
    'Valor atual: ' + (atual[3] !== '' ? 'R$ ' + atual[3] : '(vazio)') + '\nDigite o novo valor ou deixe em branco para manter:',
    ui.ButtonSet.OK_CANCEL
  );
  if (respValor.getSelectedButton() !== ui.Button.OK) return;
  if (respValor.getResponseText().trim()) {
    const novoValor = converterValor(respValor.getResponseText());
    if (isNaN(novoValor) || !(novoValor >= 0)) {
      ui.alert('Valor inválido', 'O valor informado não é um número válido.', ui.ButtonSet.OK);
      return;
    }
    campos.valor = novoValor;
  }

  if (Object.keys(campos).length === 0) {
    ui.alert('Nenhuma alteração', 'Nenhum campo novo foi informado; os dados foram mantidos.', ui.ButtonSet.OK);
    return;
  }

  const resultado = atualizarInformacoes(Number(idTexto), campos);
  if (!resultado.ok) {
    ui.alert('Erro ao atualizar', resultado.erro, ui.ButtonSet.OK);
    return;
  }

  aba.getRange(INTERVALO_PAINEL).clearContent();
  obterPlanilha().toast('Pedido #' + resultado.id + ' atualizado!', '📚 Automação', 5);
  ui.alert('Sucesso!', 'Informações do pedido #' + resultado.id + ' atualizadas com sucesso.', ui.ButtonSet.OK);
}

// ===============================================================
// 3) PROCEDIMENTO PARA LIMPAR CAMPOS
//    Pode ser chamado via código: limparCampos(id)
//    OU sem argumentos por botão / menu personalizado na planilha!
// ===============================================================
function limparCampos(id) {
  if (arguments.length === 0) {
    return acionarLimparCampos();
  }

  const aba = abaPedidos();
  const linha = acharLinha(aba, id);
  if (!linha) return { ok: false, erro: 'pedido não encontrado' };

  const antes = aba.getRange(linha, 1, 1, 4).getValues()[0];
  aba.getRange(linha, 2, 1, 3).clearContent();
  aba.getRange(linha, 5)
    .setNumberFormat('dd/mm/yyyy hh:mm:ss')
    .setValue(new Date());

  registrarLog(
    'Limpar Campos',
    antes[0],
    antes[1],
    antes[2],
    antes[3],
    'Campos do pedido #' + antes[0] + ' esvaziados via procedimento',
    'planilha'
  );

  return { ok: true, id: Number(id) };
}

function acionarLimparCampos() {
  const ui = SpreadsheetApp.getUi();
  const aba = abaPedidos();
  const valoresPainel = aba.getRange(INTERVALO_PAINEL).getValues().map(l => String(l[0] ?? '').trim());
  const idPainel = valoresPainel[0];
  const temDadosFormulario = Boolean(valoresPainel[1] || valoresPainel[2] || valoresPainel[3]);

  // Se houver dados digitados nos campos do Painel (H3:H5), limpa o Painel imediatamente
  if (temDadosFormulario) {
    aba.getRange(INTERVALO_PAINEL).clearContent();
    registrarLog('Limpar Campos', '-', '-', '-', '', 'Campos do painel (H2:H5) limpos via menu/botão', 'planilha');
    obterPlanilha().toast('Campos do painel (H2:H5) limpos com sucesso!', '🧹 Limpar Campos', 4);
    return;
  }

  // Se apenas o ID (H2) estiver preenchido, pergunta se deseja limpar o pedido daquele ID ou só o painel
  if (idPainel) {
    const escolha = ui.alert(
      '🧹 Limpar Campos',
      'Deseja esvaziar os campos (nome, livro e valor) do pedido #' + idPainel + '?\n\n' +
      'Clique em "Sim" para limpar os campos do pedido #' + idPainel + ', ou "Não" para limpar apenas a célula H2 do painel.',
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (escolha === ui.Button.CANCEL) return;
    if (escolha === ui.Button.NO) {
      aba.getRange(INTERVALO_PAINEL).clearContent();
      obterPlanilha().toast('Campo ID do painel limpo.', '🧹 Limpar Campos', 4);
      return;
    }

    const res = limparCampos(Number(idPainel));
    aba.getRange(INTERVALO_PAINEL).clearContent();
    if (!res.ok) {
      ui.alert('Erro', 'Pedido #' + idPainel + ' não encontrado.', ui.ButtonSet.OK);
      return;
    }
    obterPlanilha().toast('Campos do pedido #' + res.id + ' limpos!', '🧹 Limpar Campos', 5);
    ui.alert('Concluído', 'Os campos do pedido #' + res.id + ' foram limpos (ID mantido).', ui.ButtonSet.OK);
    return;
  }

  // Se o painel já estiver vazio, oferece limpar os campos de um pedido pelo ID
  const resp = ui.alert(
    '🧹 Limpar Campos',
    'Os campos de entrada do painel (H2:H5) já estão limpos.\nDeseja limpar os campos de um pedido específico informando o ID?',
    ui.ButtonSet.YES_NO
  );
  if (resp === ui.Button.YES) {
    limparCamposPedido();
  }
}

// Limpa diretamente apenas as células de entrada do Painel (H2:H5).
function limparPainel() {
  const aba = abaPedidos();
  aba.getRange(INTERVALO_PAINEL).clearContent();
  registrarLog('Limpar Campos', '-', '-', '-', '', 'Campos do painel (H2:H5) limpos', 'planilha');
  obterPlanilha().toast('Campos do painel (H2:H5) limpos!', '🧹 Limpar Painel', 4);
}

// Limpa os campos (nome, livro e valor) de um pedido específico informado pelo ID.
function limparCamposPedido() {
  const ui = SpreadsheetApp.getUi();
  const aba = abaPedidos();

  let idTexto = String(aba.getRange(CELULA_ID).getValue() ?? '').trim();
  if (!idTexto) {
    const resp = ui.prompt(
      '🗑️ Limpar Campos de um Pedido',
      'Informe o ID do pedido cujos campos deseja esvaziar:',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    idTexto = resp.getResponseText().trim();
  }

  const linha = acharLinha(aba, idTexto);
  if (!linha) {
    ui.alert('Pedido não encontrado', 'Não existe pedido com o ID "' + idTexto + '".', ui.ButtonSet.OK);
    return;
  }

  const confirma = ui.alert(
    'Confirmar limpeza',
    'Esvaziar nome, livro e valor do pedido #' + idTexto + ' (mantendo o ID na linha)?',
    ui.ButtonSet.YES_NO
  );
  if (confirma !== ui.Button.YES) return;

  const resultado = limparCampos(Number(idTexto));
  aba.getRange(INTERVALO_PAINEL).clearContent();
  if (resultado.ok) {
    obterPlanilha().toast('Campos do pedido #' + resultado.id + ' esvaziados!', '🗑️ Limpar Campos', 5);
    ui.alert('Sucesso', 'Os campos do pedido #' + resultado.id + ' foram limpos.', ui.ButtonSet.OK);
  }
}

// Auxiliar: carrega os dados atuais de um pedido nas células H2:H5 para facilitar a edição.
function carregarPedidoNoPainel() {
  const ui = SpreadsheetApp.getUi();
  const aba = abaPedidos();

  let idTexto = String(aba.getRange(CELULA_ID).getValue() ?? '').trim();
  if (!idTexto) {
    const resp = ui.prompt('🔍 Carregar Pedido no Painel', 'Informe o ID do pedido:', ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    idTexto = resp.getResponseText().trim();
  }

  const linha = acharLinha(aba, idTexto);
  if (!linha) {
    ui.alert('Pedido não encontrado', 'Não existe pedido com o ID "' + idTexto + '".', ui.ButtonSet.OK);
    return;
  }

  const dados = aba.getRange(linha, 1, 1, 4).getValues()[0];
  aba.getRange(INTERVALO_PAINEL).setValues([
    [dados[0]],
    [dados[1]],
    [dados[2]],
    [dados[3]]
  ]);
  obterPlanilha().toast('Pedido #' + dados[0] + ' carregado em H2:H5. Altere e clique em Atualizar!', '🔍 Painel', 6);
}

// Apaga TODOS os pedidos (colunas A:E), preservando o cabeçalho e o Painel (G1:H5).
function limparTudo(confirmadoViaCodigo) {
  const aba = abaPedidos();

  if (arguments.length === 0) {
    const ui = SpreadsheetApp.getUi();
    const confirma = ui.alert(
      '⚠️ Atenção: Limpar TODOS os pedidos',
      'Tem certeza que deseja apagar todos os pedidos registrados na aba Dados?\n(O cabeçalho e o Painel de Operações serão mantidos.)',
      ui.ButtonSet.YES_NO
    );
    if (confirma !== ui.Button.YES) return;
  } else if (!confirmadoViaCodigo) {
    return { ok: false, erro: 'confirmação necessária' };
  }

  const estado = estadoTabelaPedidos(aba);
  if (estado.ultimaLinha > 1) {
    aba.getRange(2, 1, estado.ultimaLinha - 1, 5).clearContent();
  }

  registrarLog('Limpar Tudo', '-', '-', '-', '', 'Todos os pedidos da aba Dados foram limpos', 'planilha');

  if (arguments.length === 0) {
    obterPlanilha().toast('Todos os pedidos foram removidos.', '⚠️ Limpar Tudo', 5);
  }
  return { ok: true };
}

// Garante que os cabeçalhos da tabela (A1:E1), o Painel (G1:H5) e a aba Log existam.
function configurarPlanilha() {
  const aba = abaPedidos();
  aba.getRange('A1:E1').setValues([[
    'ID',
    'NOME_DO_CLIENTE',
    'LIVRO',
    'VALOR',
    'ULTIMA ATUALIZAÇÃO'
  ]]);
  aba.getRange('G1:H5').setValues([
    ['PAINEL DE OPERAÇÕES', ''],
    ['ID (atualizar/limpar)', ''],
    ['NOME DO CLIENTE', ''],
    ['LIVRO', ''],
    ['VALOR (R$)', '']
  ]);
  aba.getRange('G1:H1').merge();
  aba.getRange('H3:H4').setNumberFormat('@');
  aba.getRange('H5').setNumberFormat('R$ #,##0.00');
  abaLog();

  if (arguments.length === 0) {
    obterPlanilha().toast('Estrutura da planilha verificada e configurada!', '🛠️ Configuração', 5);
  }
}

// ===============================================================
// ENTRADA PELA INTERNET (chamada pelo site ao finalizar compra)
// ===============================================================
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // evita dois pedidos gravarem na mesma linha simultaneamente
    const dados = JSON.parse(e.postData.contents);
    return resposta(registrarOperacao(dados.nome, dados.livro, dados.valor, 'site'));
  } catch (erro) {
    return resposta({ ok: false, erro: 'falha ao registrar' });
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function resposta(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===============================================================
// FUNÇÕES DE TESTE RÁPIDO PELO EDITOR (opcionais)
// ===============================================================
function testeRegistrar() {
  Logger.log(registrarOperacao('Cliente Teste', '1x O Hobbit', 45.5));
}

function testeAtualizar() {
  Logger.log(atualizarInformacoes(1, { valor: 99.9 }));
}

function testeLimparCampos() {
  Logger.log(limparCampos(1));
}
