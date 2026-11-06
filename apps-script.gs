/**
 * Código do Google Apps Script (NÃO roda no site).
 *
 * Como usar:
 * 1. Abra o editor: planilha > Extensões > Apps Script.
 * 2. Apague o que estiver lá e cole TODO este arquivo. Salve (Ctrl+S).
 * 3. Para o site enxergar a mudança: Implantar > Gerenciar implantações >
 *    lápis (editar) > Versão: "Nova versão" > Implantar. A URL /exec continua a mesma.
 *
 * Colunas esperadas na primeira aba:
 *   A: ID | B: NOME_DO_CLIENTE | C: LIVRO | D: VALOR | E: ULTIMA ATUALIZAÇÃO
 *
 * Procedimentos:
 *   registrarOperacao(nome, livro, valor)        -> acrescenta um pedido (o site usa este)
 *   atualizarInformacoes(id, {nome, livro, valor}) -> altera um pedido existente
 *   limparCampos(id)                              -> esvazia os campos de UM pedido
 *   limparTudo()                                  -> apaga todos os pedidos (mantém o cabeçalho)
 *
 * Só o "registrar" é chamado pela internet (doPost). Os demais você roda pelo
 * editor (botão Executar), porque a URL do site é pública e qualquer pessoa
 * poderia apagar ou alterar pedidos se eles estivessem expostos.
 */

const ID_PLANILHA = '1txnYqaemiW6WD_eE-t6JQiMpvzom3WaHUw11vDPf2WM';

function abaPedidos() {
  return SpreadsheetApp.openById(ID_PLANILHA).getSheets()[0];
}

// Acha o número da linha do pedido com este ID (ou 0 se não existir).
function acharLinha(aba, id) {
  const ultima = aba.getLastRow();
  if (ultima < 2) return 0;
  const ids = aba.getRange(2, 1, ultima - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (Number(ids[i][0]) === Number(id)) return i + 2; // +2: cabeçalho + índice começa em 0
  }
  return 0;
}

function textoSeguro(valor, max) {
  return String(valor || '').trim().slice(0, max);
}

// ---------------------------------------------------------------
// 1) REGISTRAR UMA OPERAÇÃO
// ---------------------------------------------------------------
function registrarOperacao(nome, livro, valor) {
  nome = textoSeguro(nome, 80);
  livro = textoSeguro(livro, 300);
  valor = Number(valor);
  if (!nome || !livro || !(valor >= 0)) {
    return { ok: false, erro: 'dados inválidos' };
  }

  const aba = abaPedidos();
  const ultima = aba.getLastRow();
  const linha = ultima + 1;
  const id = ultima <= 1 ? 1 : Number(aba.getRange(ultima, 1).getValue()) + 1;

  // colunas de texto como "texto simples": um nome começando com "=" não vira fórmula
  aba.getRange(linha, 2, 1, 2).setNumberFormat('@');
  aba.getRange(linha, 1, 1, 5).setValues([[id, nome, livro, valor, new Date()]]);
  return { ok: true, id: id };
}

// ---------------------------------------------------------------
// 2) ATUALIZAR INFORMAÇÕES
//    campos: objeto com o que mudar, por exemplo { valor: 80 } ou { nome: 'Ana', livro: '...' }
// ---------------------------------------------------------------
function atualizarInformacoes(id, campos) {
  const aba = abaPedidos();
  const linha = acharLinha(aba, id);
  if (!linha) return { ok: false, erro: 'pedido não encontrado' };

  if (campos.nome !== undefined) {
    const c = aba.getRange(linha, 2);
    c.setNumberFormat('@');
    c.setValue(textoSeguro(campos.nome, 80));
  }
  if (campos.livro !== undefined) {
    const c = aba.getRange(linha, 3);
    c.setNumberFormat('@');
    c.setValue(textoSeguro(campos.livro, 300));
  }
  if (campos.valor !== undefined) {
    const valor = Number(campos.valor);
    if (!(valor >= 0)) return { ok: false, erro: 'valor inválido' };
    aba.getRange(linha, 4).setValue(valor);
  }

  aba.getRange(linha, 5).setValue(new Date()); // ULTIMA ATUALIZAÇÃO
  return { ok: true, id: Number(id) };
}

// ---------------------------------------------------------------
// 3) LIMPAR CAMPOS
// ---------------------------------------------------------------
// Esvazia nome, livro e valor de UM pedido (mantém o ID e marca a data).
function limparCampos(id) {
  const aba = abaPedidos();
  const linha = acharLinha(aba, id);
  if (!linha) return { ok: false, erro: 'pedido não encontrado' };

  aba.getRange(linha, 2, 1, 3).clearContent();
  aba.getRange(linha, 5).setValue(new Date());
  return { ok: true, id: Number(id) };
}

// Apaga TODOS os pedidos e mantém a linha de cabeçalho.
function limparTudo() {
  const aba = abaPedidos();
  const ultima = aba.getLastRow();
  if (ultima > 1) {
    aba.getRange(2, 1, ultima - 1, aba.getLastColumn()).clearContent();
  }
  return { ok: true };
}

// ---------------------------------------------------------------
// Entrada pela internet (o site): só registra pedidos.
// ---------------------------------------------------------------
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // evita dois pedidos gravarem na mesma linha

  try {
    const dados = JSON.parse(e.postData.contents);
    return resposta(registrarOperacao(dados.nome, dados.livro, dados.valor));
  } catch (erro) {
    return resposta({ ok: false, erro: 'falha ao registrar' });
  } finally {
    lock.releaseLock();
  }
}

function resposta(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
// Para testar pelo editor: escolha a função no menu e clique em Executar.
// Mexem na planilha de verdade, então use dados fictícios.
// ---------------------------------------------------------------
function testeRegistrar() {
  Logger.log(registrarOperacao('Teste', '1x Livro de teste', 10));
}
function testeAtualizar() {
  Logger.log(atualizarInformacoes(1, { valor: 99 })); // troque o 1 pelo ID que quer alterar
}
function testeLimparCampos() {
  Logger.log(limparCampos(1)); // troque o 1 pelo ID que quer esvaziar
}
