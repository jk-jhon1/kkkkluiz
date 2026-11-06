/**
 * Los Nanichos – Automação de Tarefas (Lógica: Procedimentos)
 * Ferramenta: Google Planilhas + Google Apps Script
 *
 * Como instalar/atualizar na planilha:
 * 1. Na planilha Google ("TRABALHO INDIO"), clique em: Extensões > Apps Script.
 * 2. Apague todo o código antigo do editor, cole TODO este arquivo e salve (Ctrl+S).
 * 3. Recarregue a planilha (F5):
 *    - As caixas de botão em H8 ("Apagar Última Linha") e H9 ("Apagar Todos os Dados") já funcionam na hora!
 *    - Para inserir também os 2 botões gráficos flutuantes na aba Dados, clique no menu do topo:
 *      "📚 Automação" > "🔘 Inserir os 2 botões na planilha".
 * 4. Se quiser atualizar também a versão do site (sem mudar o link):
 *    Implantar > Gerenciar implantações > lápis (Editar) > Versão: "Nova versão" > Implantar.
 *
 * Funções para botões (Inserir > Desenho > Atribuir script):
 *   - apagarUltimaLinha      -> apaga a última linha preenchida da tabela (A:E)
 *   - apagarTodosOsDados     -> apaga todos os dados registrados na tabela (A:E)
 *   - registrarOperacao      -> registra uma operação
 *   - atualizarInformacoes   -> atualiza informações de um pedido
 *   - limparCampos           -> limpa campos do painel ou de um pedido
 */

const ID_PLANILHA = '1txnYqaemiW6WD_eE-t6JQiMpvzom3WaHUw11vDPf2WM';

// Células do Painel de Operações e dos Botões Rápidos na aba "Dados"
const CELULA_ID = 'H2';
const CELULA_NOME = 'H3';
const CELULA_LIVRO = 'H4';
const CELULA_VALOR = 'H5';
const INTERVALO_PAINEL = 'H2:H5';
const LINHA_CHECK_ULTIMA = 8; // H8 -> Apagar Última Linha
const LINHA_CHECK_TUDO = 9;   // H9 -> Apagar Todos os Dados
const COLUNA_PAINEL_INPUT = 8; // Coluna H

// Imagens PNG em Base64 para gerar os 2 botões gráficos automaticamente na planilha
const PNG_BTN_APAGAR_ULTIMA =
  'iVBORw0KGgoAAAANSUhEUgAAAOYAAAAoCAMAAADdVnZNAAACBFBMVEX////CQQzDQQzQSQzcUAzfUgzgUgzVTAzoVgzqWAzVSwzqWxH75Nf518TqXRP0q4X74tXrYhv97+n////xlGP3wKP4ybD74NL3w6jrZR/4zbf2uJn4zrj62srvhk/rYBj+9/T++/rrZiDym2798ev86uDzpHv1so/507/99fDwi1b2upr++PX86N30qYPsbiz50r775NjxkmDufEH0rIb51sP++/nxlGT74tb98uz51cL++vjtcC/0qYL98er0ron99fHwiVTzpn7++ff63M32upv1sY375dn3v6L+/fztdTb62Mb3wqf2uJjymGnqWxD75tr86+L3wKT98Or+/v7618XwiVPzpX386d/1tpX+/Pv5z7r97+j98+787uf50Lvvg0r99O/qXhX51MH75tvtdjfvgUjtdznvf0b0r4r628vqWQ3qXxb3xqzvhU3sby7vh1D3xq3ynnLtdDT1tJPsaibtdjj87eb3waXymGr3xav4ybH0rYj+/f3ymWv4x67saSXxkF3xl2jxlmbuf0X++fbrYRr749bwjFj4y7PvhEzwiFLqXBLxkV/uej7qXxfxk2L4y7TynHDwjlvvgkn87OP1tpb4yrL63s/0qIH50bzwjVn1tZT98+3sbi3uej399vL87OT2u5zrZB7sbCr62cfrYxz1s5D508D2vJ7ynG/qWg9CeMHqAAAAAXRSTlMAQObYZgAAAxxJREFUaN7tmodTE1EQxjlSNyAiXxBEIyBBQQJIF1BBY0MIAtIFUbGjIhYUMaIIauyCIvaC3X/Sd5dyaY6gl9x45zeTmc3O7tv95W1eZu4lJkZN4hQrP8hYjtNodXrFSafVcFysuJUGIylURpN3Q2M5rdzNRFJaz35yJrk7iawMnHtkFTuxbhmFseU0cvcRaWkETEV/M3lpBUyd3G1EWjoBUy93G5GW/j+mpIqLX5SgAszFiUvko4z+0CbBHGLNN2MBSX+HmQws/YMqKUglWga+Rb7PNCxnzhWWlb9PCo6LCmZ6BpC5aiHrx/0ac16fTbCigpkFazZWC9XW5OTGryXKy7Zl5hcwT+G6jKJioET0JKG0rLwiDOZ6MFUKHVeioKrasmGjdVMN+S0WgMnH8fVq8zfzb7ak2XkjMFZazK3Yth07hNK2mp2o20X1DQWOajTS7iY0t+zhMb0eFtPa1tAeBrOjEyldNR7M7r096N3XB+wnMTUMZqYjBwcOMqOu/xBvBMZKinkYOHIUOMaXPk50Al2UQAMnW3CKTgODdIbH9HpYzFDAxgQPrRszj84C5ygRHSSmhsE8TxeAYWYk0yBvBMZKiunAxQEqwiW+9AjRZYzSlWYnm8AqaoSd+EZKfB4Wc1WcgjGia8D1UMxxKoaNqA1san2pYTBvUDowwYxGtxEYKyVmQi8ETd5k1RxEt9B1G5MVEy5W7Q5QSHcZps/jf1oMofMe3QceuL0utodeTDPz57J5ZZhiahhMMz10Y7qNoFgpMR8BaS7X41ZMsWrdY/VwDk/DNjJVy6rx381UK8P0efwxnzjxtK0cFk/TM+hxZIViiqkeTKvFYukLjxkUKyVmv3vVZ5hl1Z6/sBe9pIFXrytHG3j/G3bSjgNvRY//2d9e9q77/Yc5D+bHT3b0hWKKqR5M8UQOwQyKlRLTT6E/X5/ZaxrOuYWuFB1Jhvnla+lMLb7JzRNpTHOTPXf2+w+5eaTF/Nf0H1NJ0qvpkZdKHmCq5HG0Oi4XYjiD3I1EViZOTRd/bGxNip1bo0G8l1fHpbxa/mKhfP0EUSCpl9u+80MAAAAASUVORK5CYII=';

const PNG_BTN_APAGAR_TUDO =
  'iVBORw0KGgoAAAANSUhEUgAAAOYAAAAoCAMAAADdVnZNAAAB+1BMVEX///+aGxuZGxubGxuxHx/EIiLLIyPFIiKyHx+5ICDYJSXcJia4ICDeNzf////jVVXxra3zubn88vL409P419f++/v88PDtkZHdMTH99fX++vrfODjcKyv1xcX30ND99PT87u787e364eHfPDzqfX364+PsiYn76+vum5v53NzmaGjwpaX99vbtkJDgQkLocXHtk5P2ysr++fnodXXztLTeNTXvoaH99/f53t7gQ0PrhYX9+PjysbHnbW3ztbX30tLwpqbumZn53d3ulpbxrKz+/PziTEz2zMzysLDwo6PpeXncKir65eXxrq777Oz+/v72y8vmZWX52tr64uLxqqr76Oj3z8/cJyf75+fysrLjUVHtkpLxqaniTU3kW1viTk7kWVnul5f30dHvoKDdLCzdLy/lYGDgRETlY2PmZ2fgQEDlX1/hR0fkWFjrh4fyr6/penrytLTzurrtlZX+/f32yMjfPT3pe3vqgYHocnLumJjeMjL529vmamr0u7v2ycnmZGTodHTgPz/iUFDiT0/rhITnbGzztrboc3P0v7/0vLzqf3/zt7f98/P64ODriIjjVFT42dnsjY32x8f0urr41NTsjo71w8Pnamrvn5/87+/peHjjUlLvnZ3eMzPhRUX1xsbqfHz88fHkXFzeNDTeNjbvnJzqfn7cKSnueRX0AAAAAXRSTlMAQObYZgAAA0VJREFUaN7tmudXE0EUxcNQQhi4lABCIAlFhEgRQQEJiooKoSvGKAIiIHZFxAqiWLAgWLEX7P6ZzuxmN0FWSnDd44Z7TnLIPW/fe7+8mcmHwWAIJoUQnSrEDzKUkLDwCB3KGElIqERJiCmK6lRR0YRIszRq3YyaMnrnSaK17kRdmYh4+uh2xYqKEs4hEql1H2orjI9T3zuTK1zAjNC6DbUVsYKpI61g6kn/GDMGiNUi0SIw44D4pfeRAFHm5WEmAknJ5lUpqmOmpAKWtKW0ls7frDabHRk2W9xyMTPNmUBWttqYq5GZgzXsj1zkWRMcaynNzymwZBUyp8ia6ljH68lOLorXl5R6n7RhA3vfWJZVXpHHxrFJipadygqn3VHpjVZ2E1FF6WYLtviVXSBRYJjV2LoN2wWEgpod2LmL1tYVukpQTxsq0NjUzOtJDotpaa3b7Y+Zvgdte93Y52HR+w+wwcTKTjsyD9bE53uXgKIrYtIOlDT4l50vUWCYnUDXIaCbIxymtAe9tI96jjQxhH7gKD3GMSWHxWzxPSpgHgdO0HaW4CRwin+KlZ0uuOtPS7tO2fVingHS5CILJAoM0wW3hzpwliMMUHoOg/R84xA7Wy7QejgpvcgxJYfFXPoN8zLsbFTAFSE6m0XLjqejhG37q2Kwsus/zVll/5woIMy+ZPHALB9mCC5KR9B7DeWl3aO4zr/WInqD1ZMdFjM2d5o36S32lffziXSLQxAd1uPtO8CwNE0FV8Qct+AulYsskCggzHvA/dHRBy2YYAhJD2sx1DmOgoGJSVaP780qvklkRwEz3YG2Djd6PA1TqH3k4FtKch4/cT1twmSDEKzs8pO2ulk4aX1l508UEOYzvhQpfY5phvDipdPxinpev8kYrONQ/Mh7C7zzOXMxaYyrzTn1/gOlH2fsI5/4EpecImuz5XPrF2+0sst+N4eSzYVs4/mKLJAoEEw/zUIQ9JW9xjH0bTFPa6tlYX6fKS6bRIfWDGpjjlU4E6Z//NSa4e9i/s9awdSTggozXOs21JZRwAzTug21FUmC53LBQExaN6KuokkwXfwZCInW7bqNMknXuMKlfKRR6wt0NRQe5ncpHyz/YqF//QL+7Al68He2CAAAAABJRU5ErkJggg==';

// ===============================================================
// MENU PERSONALIZADO (acionado automaticamente ao abrir a planilha)
// ===============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 Automação')
    .addItem('🔘 Inserir os 2 botões na planilha', 'criarBotoesNaPlanilha')
    .addSeparator()
    .addItem('🗑️ Apagar última linha com dados', 'apagarUltimaLinha')
    .addItem('⚠️ Apagar todos os dados da planilha', 'apagarTodosOsDados')
    .addSeparator()
    .addItem('➕ Registrar operação', 'registrarOperacao')
    .addItem('✏️ Atualizar informações', 'atualizarInformacoes')
    .addItem('🧹 Limpar campos', 'limparCampos')
    .addSeparator()
    .addItem('🔍 Carregar pedido no painel (por ID)', 'carregarPedidoNoPainel')
    .addItem('🗑️ Limpar campos de um pedido (por ID)', 'limparCamposPedido')
    .addItem('🧹 Limpar apenas o painel (H2:H5)', 'limparPainel')
    .addItem('🛠️ Configurar estrutura da planilha', 'configurarPlanilha')
    .addToUi();
}

// ===============================================================
// ACIONAMENTO IMEDIATO PELOS BOTÕES DE CAIXA (H8 e H9 na aba Dados)
// ===============================================================
function onEdit(e) {
  if (!e || !e.range || !e.source) return;
  const aba = e.range.getSheet();
  const nomeAba = aba.getName();
  if (nomeAba !== 'Dados' && aba.getIndex() !== 1) return;

  const col = e.range.getColumn();
  const lin = e.range.getRow();
  if (col !== COLUNA_PAINEL_INPUT || (lin !== LINHA_CHECK_ULTIMA && lin !== LINHA_CHECK_TUDO)) {
    return;
  }

  if (e.range.getValue() !== true) return;

  // Desmarca imediatamente para funcionar como botão de clique
  e.range.setValue(false);

  if (lin === LINHA_CHECK_ULTIMA) {
    executarApagarUltimaLinha(e.source);
  } else if (lin === LINHA_CHECK_TUDO) {
    executarApagarTodosOsDados(e.source);
  }
}

// ===============================================================
// CRIA OS 2 BOTÕES GRÁFICOS CLICÁVEIS NA PLANILHA
// ===============================================================
function criarBotoesNaPlanilha() {
  const aba = abaPedidos();

  // Remove versões anteriores dos mesmos botões para não duplicar
  const imagens = aba.getImages();
  for (let i = 0; i < imagens.length; i++) {
    const titulo = imagens[i].getAltTextTitle();
    if (titulo === 'BTN_APAGAR_ULTIMA' || titulo === 'BTN_APAGAR_TUDO') {
      imagens[i].remove();
    }
  }

  // Botão 1: Apagar Última Linha (coluna G = 7, linha 11)
  const blobUltima = Utilities.newBlob(
    Utilities.base64Decode(PNG_BTN_APAGAR_ULTIMA),
    'image/png',
    'botao_apagar_ultima.png'
  );
  const imgUltima = aba.insertImage(blobUltima, 7, 11, 0, 4);
  imgUltima.assignScript('apagarUltimaLinha');
  imgUltima.setAltTextTitle('BTN_APAGAR_ULTIMA');

  // Botão 2: Apagar Todos os Dados (coluna G = 7, linha 14)
  const blobTudo = Utilities.newBlob(
    Utilities.base64Decode(PNG_BTN_APAGAR_TUDO),
    'image/png',
    'botao_apagar_tudo.png'
  );
  const imgTudo = aba.insertImage(blobTudo, 7, 14, 0, 4);
  imgTudo.assignScript('apagarTodosOsDados');
  imgTudo.setAltTextTitle('BTN_APAGAR_TUDO');

  mostrarToast('Os 2 botões gráficos foram inseridos em G11 e G14!', '🔘 Botões Criados');
}

// ===============================================================
// FUNÇÕES AUXILIARES DE PLANILHA, VALIDAÇÃO E LOG
// ===============================================================
function obterPlanilha(forcarPorId, planilhaExplicita) {
  if (planilhaExplicita) return planilhaExplicita;
  if (!forcarPorId) {
    try {
      const ativa = SpreadsheetApp.getActiveSpreadsheet();
      if (ativa) return ativa;
    } catch (e) {}
  }
  return SpreadsheetApp.openById(ID_PLANILHA);
}

function mostrarToast(mensagem, titulo, planilhaExplicita) {
  try {
    const alvo = planilhaExplicita || SpreadsheetApp.getActiveSpreadsheet();
    if (alvo) {
      alvo.toast(mensagem, titulo || '📚 Automação', 5);
    }
  } catch (e) {
    // Ignora se chamado fora da interface da planilha
  }
}

function abaPedidos(forcarPorId, planilhaExplicita) {
  const planilha = obterPlanilha(forcarPorId, planilhaExplicita);
  return planilha.getSheetByName('Dados') || planilha.getSheets()[0];
}

function abaLog(forcarPorId, planilhaExplicita) {
  const planilha = obterPlanilha(forcarPorId, planilhaExplicita);
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

function registrarLog(acao, id, nome, livro, valor, observacao, origem, planilhaExplicita) {
  try {
    const log = abaLog(origem === 'site', planilhaExplicita);
    const linha = log.getLastRow() + 1;
    const valorCelula = valor === '' || valor === undefined || valor === null || isNaN(Number(valor))
      ? ''
      : Number(valor);
    const usuario = origem === 'site' ? 'Site (Web App)' : 'Planilha (Menu/Botão)';

    log.getRange(linha, 5, 1, 2).setNumberFormat('@');
    log.getRange(linha, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
    if (valorCelula !== '') {
      log.getRange(linha, 7).setNumberFormat('R$ #,##0.00');
    }

    log.getRange(linha, 1, 1, 8).setValues([[
      new Date(),
      usuario,
      acao,
      id || '-',
      nome || '-',
      livro || '-',
      valorCelula,
      observacao || ''
    ]]);
  } catch (erro) {
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

// Lê apenas as colunas A:E (ignora o Painel de Operações e botões em G:H).
function estadoTabelaPedidos(aba) {
  const maxLinha = aba.getLastRow();
  if (maxLinha < 2) {
    return { ultimaLinha: 1, proximoId: 1 };
  }

  const dados = aba.getRange(2, 1, maxLinha - 1, 5).getValues();
  let ultimaLinha = 1;
  let maiorId = 0;

  for (let i = 0; i < dados.length; i++) {
    const temConteudo = dados[i].some(function(celula) {
      return celula !== '' && celula !== null && celula !== undefined;
    });
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
      return i + 2;
    }
  }
  return 0;
}

function formatarLinhaPedido(aba, linha) {
  try {
    const intervalo = aba.getRange(linha, 1, 1, 5);
    aba.getRange(linha, 4).setNumberFormat('R$ #,##0.00');
    aba.getRange(linha, 5).setNumberFormat('dd/mm/yyyy hh:mm:ss');
    aba.getRange(linha, 1).setHorizontalAlignment('center');
    aba.getRange(linha, 4, 1, 2).setHorizontalAlignment('center');
    intervalo.setBorder(true, true, true, true, true, true);
  } catch (erro) {
    Logger.log('Aviso ao formatar linha: ' + erro);
  }
}

// ===============================================================
// BOTÃO 1: APAGAR A ÚLTIMA LINHA QUE TEM INFORMAÇÕES
// ===============================================================
function executarApagarUltimaLinha(planilhaExplicita) {
  const aba = abaPedidos(false, planilhaExplicita);
  const estado = estadoTabelaPedidos(aba);

  if (estado.ultimaLinha <= 1) {
    mostrarToast('Não há nenhuma linha com informações para apagar.', '🗑️ Apagar Última Linha', planilhaExplicita);
    return { ok: false, erro: 'nenhuma linha com informações' };
  }

  const linhaAlvo = estado.ultimaLinha;
  const intervalo = aba.getRange(linhaAlvo, 1, 1, 5);
  const removidos = intervalo.getValues()[0];

  intervalo.clearContent();
  try {
    intervalo.setBorder(false, false, false, false, false, false);
  } catch (e) {}

  registrarLog(
    'Apagar Última Linha',
    removidos[0] || '-',
    removidos[1] || '-',
    removidos[2] || '-',
    removidos[3],
    'Última linha preenchida (linha ' + linhaAlvo + ') apagada via botão',
    'planilha',
    planilhaExplicita
  );

  mostrarToast(
    'Linha ' + linhaAlvo + ' apagada com sucesso!',
    '🗑️ Apagar Última Linha',
    planilhaExplicita
  );
  return { ok: true, linha: linhaAlvo, id: removidos[0] };
}

function apagarUltimaLinha() {
  const resultado = executarApagarUltimaLinha();
  if (!resultado.ok && arguments.length === 0) {
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Planilha vazia', 'Não há nenhuma linha com informações para apagar.', ui.ButtonSet.OK);
    } catch (e) {}
  }
  return resultado;
}

// ===============================================================
// BOTÃO 2: APAGAR TODOS OS DADOS DA PLANILHA
// ===============================================================
function executarApagarTodosOsDados(planilhaExplicita) {
  const aba = abaPedidos(false, planilhaExplicita);
  const estado = estadoTabelaPedidos(aba);

  if (estado.ultimaLinha > 1) {
    const intervalo = aba.getRange(2, 1, estado.ultimaLinha - 1, 5);
    intervalo.clearContent();
    try {
      intervalo.setBorder(false, false, false, false, false, false);
    } catch (e) {}
  }

  // Também limpa os campos digitados no Painel (H2:H5), mantendo os rótulos e cabeçalhos
  aba.getRange(INTERVALO_PAINEL).clearContent();

  registrarLog(
    'Apagar Todos os Dados',
    '-',
    '-',
    '-',
    '',
    'Todos os dados da planilha foram apagados via botão',
    'planilha',
    planilhaExplicita
  );

  mostrarToast('Todos os dados da planilha foram apagados!', '⚠️ Apagar Todos os Dados', planilhaExplicita);
  return { ok: true };
}

function apagarTodosOsDados(confirmadoViaCodigo) {
  if (arguments.length === 0) {
    const ui = SpreadsheetApp.getUi();
    const confirma = ui.alert(
      '⚠️ Apagar Todos os Dados',
      'Tem certeza que deseja apagar TODOS os dados registrados na planilha?\n(O cabeçalho e o painel serão preservados.)',
      ui.ButtonSet.YES_NO
    );
    if (confirma !== ui.Button.YES) return { ok: false, erro: 'cancelado' };
  } else if (!confirmadoViaCodigo) {
    return { ok: false, erro: 'confirmação necessária' };
  }

  return executarApagarTodosOsDados();
}

// Mantém compatibilidade com o nome anterior limparTudo
function limparTudo(confirmadoViaCodigo) {
  if (arguments.length === 0) {
    return apagarTodosOsDados();
  }
  return apagarTodosOsDados(confirmadoViaCodigo);
}

// ===============================================================
// 1) PROCEDIMENTO PARA REGISTRAR UMA OPERAÇÃO
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

  const ehSite = origem === 'site';
  const aba = abaPedidos(ehSite);
  const estado = estadoTabelaPedidos(aba);
  const linha = estado.ultimaLinha + 1;
  const id = estado.proximoId;
  const agora = new Date();

  aba.getRange(linha, 2, 1, 2).setNumberFormat('@');
  aba.getRange(linha, 1, 1, 5).setValues([[id, nomeLimpo, livroLimpo, valorNum, agora]]);
  formatarLinhaPedido(aba, linha);

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
  const brutoValor = aba.getRange(CELULA_VALOR).getValue();
  let valorTexto = String(brutoValor === undefined || brutoValor === null ? '' : brutoValor).trim();

  const usouPainel = Boolean(nome || livro || valorTexto);

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
  mostrarToast('Pedido #' + resultado.id + ' registrado com sucesso!', '📚 Automação');
  ui.alert('Sucesso!', 'Operação registrada com sucesso (ID #' + resultado.id + ').', ui.ButtonSet.OK);
}

// ===============================================================
// 2) PROCEDIMENTO PARA ATUALIZAR INFORMAÇÕES
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

  const brutoId = aba.getRange(CELULA_ID).getValue();
  let idTexto = String(brutoId === undefined || brutoId === null ? '' : brutoId).trim();
  const nomePainel = textoSeguro(aba.getRange(CELULA_NOME).getValue(), 80);
  const livroPainel = textoSeguro(aba.getRange(CELULA_LIVRO).getValue(), 300);
  const brutoValor = aba.getRange(CELULA_VALOR).getValue();
  const valorPainelTexto = String(brutoValor === undefined || brutoValor === null ? '' : brutoValor).trim();

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
    mostrarToast('Pedido #' + resPainel.id + ' atualizado!', '📚 Automação');
    ui.alert('Sucesso!', 'Informações do pedido #' + resPainel.id + ' atualizadas com sucesso.', ui.ButtonSet.OK);
    return;
  }

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
  mostrarToast('Pedido #' + resultado.id + ' atualizado!', '📚 Automação');
  ui.alert('Sucesso!', 'Informações do pedido #' + resultado.id + ' atualizadas com sucesso.', ui.ButtonSet.OK);
}

// ===============================================================
// 3) PROCEDIMENTO PARA LIMPAR CAMPOS
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
  const valoresPainel = aba.getRange(INTERVALO_PAINEL).getValues().map(function(l) {
    return String(l[0] === undefined || l[0] === null ? '' : l[0]).trim();
  });
  const idPainel = valoresPainel[0];
  const temDadosFormulario = Boolean(valoresPainel[1] || valoresPainel[2] || valoresPainel[3]);

  if (temDadosFormulario) {
    aba.getRange(INTERVALO_PAINEL).clearContent();
    registrarLog('Limpar Campos', '-', '-', '-', '', 'Campos do painel (H2:H5) limpos via menu/botão', 'planilha');
    mostrarToast('Campos do painel (H2:H5) limpos com sucesso!', '🧹 Limpar Campos');
    return;
  }

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
      mostrarToast('Campo ID do painel limpo.', '🧹 Limpar Campos');
      return;
    }

    const res = limparCampos(Number(idPainel));
    aba.getRange(INTERVALO_PAINEL).clearContent();
    if (!res.ok) {
      ui.alert('Erro', 'Pedido #' + idPainel + ' não encontrado.', ui.ButtonSet.OK);
      return;
    }
    mostrarToast('Campos do pedido #' + res.id + ' limpos!', '🧹 Limpar Campos');
    ui.alert('Concluído', 'Os campos do pedido #' + res.id + ' foram limpos (ID mantido).', ui.ButtonSet.OK);
    return;
  }

  const resp = ui.alert(
    '🧹 Limpar Campos',
    'Os campos de entrada do painel (H2:H5) já estão limpos.\nDeseja limpar os campos de um pedido específico informando o ID?',
    ui.ButtonSet.YES_NO
  );
  if (resp === ui.Button.YES) {
    limparCamposPedido();
  }
}

function limparPainel() {
  const aba = abaPedidos();
  aba.getRange(INTERVALO_PAINEL).clearContent();
  registrarLog('Limpar Campos', '-', '-', '-', '', 'Campos do painel (H2:H5) limpos', 'planilha');
  mostrarToast('Campos do painel (H2:H5) limpos!', '🧹 Limpar Painel');
}

function limparCamposPedido() {
  const ui = SpreadsheetApp.getUi();
  const aba = abaPedidos();

  const brutoId = aba.getRange(CELULA_ID).getValue();
  let idTexto = String(brutoId === undefined || brutoId === null ? '' : brutoId).trim();
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
    mostrarToast('Campos do pedido #' + resultado.id + ' esvaziados!', '🗑️ Limpar Campos');
    ui.alert('Sucesso', 'Os campos do pedido #' + resultado.id + ' foram limpos.', ui.ButtonSet.OK);
  }
}

function carregarPedidoNoPainel() {
  const ui = SpreadsheetApp.getUi();
  const aba = abaPedidos();

  const brutoId = aba.getRange(CELULA_ID).getValue();
  let idTexto = String(brutoId === undefined || brutoId === null ? '' : brutoId).trim();
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
  mostrarToast('Pedido #' + dados[0] + ' carregado em H2:H5. Altere e clique em Atualizar!', '🔍 Painel');
}

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
  aba.getRange('G7:H9').setValues([
    ['BOTÕES DE EXCLUSÃO', ''],
    ['Apagar Última Linha', false],
    ['Apagar Todos os Dados', false]
  ]);
  aba.getRange('G7:H7').merge();
  aba.getRange('H8:H9').insertCheckboxes();
  abaLog();
  criarBotoesNaPlanilha();

  if (arguments.length === 0) {
    mostrarToast('Estrutura e botões configurados!', '🛠️ Configuração');
  }
}

// ===============================================================
// ENTRADA PELA INTERNET (chamada pelo site ao finalizar compra)
// ===============================================================
function doGet() {
  return resposta({ ok: true, status: 'Web App ativo e funcionando!' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const dados = JSON.parse(e.postData.contents);
    return resposta(registrarOperacao(dados.nome, dados.livro, dados.valor, 'site'));
  } catch (erro) {
    return resposta({
      ok: false,
      erro: 'falha ao registrar: ' + (erro && erro.message ? erro.message : String(erro))
    });
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
