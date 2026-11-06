/**
 * Los Nanichos – Automação de Tarefas (Lógica: Procedimentos)
 * Ferramenta: Google Planilhas + Google Apps Script
 *
 * Como instalar/atualizar na planilha e no site:
 * 1. Na planilha Google ("TRABALHO INDIO"), clique em: Extensões > Apps Script.
 * 2. Apague todo o código antigo do editor, cole TODO este arquivo e salve (Ctrl+S).
 * 3. IMPORTANTE para o link do site atualizar:
 *    No Apps Script, clique em: Implantar > Gerenciar implantações > ícone de lápis (Editar)
 *    > em "Versão", selecione "Nova versão" > clique em "Implantar".
 *    (O link /exec continua exatamente o mesmo!)
 * 4. Volte na planilha e aperte F5:
 *    - Mesmo antes de reimplantar, ao clicar na planilha ou apertar F5 o script corrige
 *      automaticamente qualquer linha desalinhada!
 *    - Para ativar correção automática em segundo plano + 3 botões gráficos, clique em:
 *      "📚 Automação" > "🔘 Inserir os 3 botões na planilha".
 *
 * Funções para botões (Inserir > Desenho > Atribuir script):
 *   - apagarLinhaSelecionada -> apaga a(s) linha(s) selecionada(s) na tabela (A:G)
 *   - apagarUltimaLinha      -> apaga a última linha preenchida da tabela (A:G)
 *   - apagarTodosOsDados     -> apaga todos os dados registrados na tabela (A:G)
 *   - registrarOperacao      -> registra uma operação
 *   - atualizarInformacoes   -> atualiza informações de um pedido
 *   - limparCampos           -> limpa campos do painel ou de um pedido
 */

const ID_PLANILHA = '1txnYqaemiW6WD_eE-t6JQiMpvzom3WaHUw11vDPf2WM';

// Tabela principal na aba "Dados": colunas A até G (7 colunas)
// A: ID | B: NOME_DO_CLIENTE | C: E-MAIL (GMAIL) | D: ENDEREÇO | E: LIVRO | F: VALOR | G: ULTIMA ATUALIZAÇÃO
const COLUNAS_TABELA = 7;

// Células do Painel de Operações (I1:J7) e Botões Rápidos (I9:J12) na aba "Dados"
const CELULA_ID = 'J2';
const CELULA_NOME = 'J3';
const CELULA_EMAIL = 'J4';
const CELULA_ENDERECO = 'J5';
const CELULA_LIVRO = 'J6';
const CELULA_VALOR = 'J7';
const INTERVALO_PAINEL = 'J2:J7';

const COLUNA_PAINEL_BOTAO_IMG = 9;  // Coluna I (onde ficam os botões gráficos flutuantes)
const COLUNA_PAINEL_INPUT = 10;     // Coluna J (onde ficam os checkboxes J10:J12)
const LINHA_CHECK_SELECIONADA = 10; // J10 -> Apagar Linha Selecionada
const LINHA_CHECK_ULTIMA = 11;      // J11 -> Apagar Última Linha
const LINHA_CHECK_TUDO = 12;        // J12 -> Apagar Todos os Dados
const CHAVE_CACHE_SELECAO = 'ULTIMA_SELECAO_TABELA_DADOS';

const CABECALHO_DADOS = [
  'ID',
  'NOME_DO_CLIENTE',
  'E-MAIL (GMAIL)',
  'ENDEREÇO',
  'LIVRO',
  'VALOR',
  'ULTIMA ATUALIZAÇÃO'
];

const CABECALHO_LOG = [
  'DATA/HORA',
  'E-MAIL (GMAIL)',
  'ENDEREÇO',
  'AÇÃO',
  'ID',
  'NOME DO CLIENTE',
  'LIVRO',
  'VALOR',
  'OBSERVAÇÃO'
];

// Imagens PNG em Base64 para gerar os 3 botões gráficos automaticamente na planilha
const PNG_BTN_APAGAR_SELECIONADA =
  'iVBORw0KGgoAAAANSUhEUgAAAOYAAAAoCAMAAADdVnZNAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAACHFBMVEX///8lY+vr8P3///9zm/I2b+xUhe+sxPe4zPjW4vuxx/e/0fnP3fphjvA3cOyowffn7fz5+/7t8v3R3vqPrvQwauz1+P76+/4qZuvH1/nw9P3E1fnP3Prv8/19ovPj6/yJqvSat/XF1vny9f3b5ftnkvGlvvb2+P7h6fyPr/RCd+3c5vtxmfGTsfXJ2fp0m/Lz9v1lkfCLrPS+0fn4+v4pZuv9/f5ZiO8nZOtIfO4zbeyfuvbg6PyEp/MybOyZtvXd5/usw/c0buygu/b3+f7e5/wwa+xCeO3S3/qmv/aVs/X8/f5Lfu7M2vqwxveivPb5+v54nvLl7PytxPfs8f3+/v7L2vplkPDx9f0taevf6Pzi6vyXtPVslvHY4/uRsPTX4vvz9/2uxfezyfja5Puqwvfo7/wmZOuyx/dMf+5aie9OgO6WtPXQ3forZ+suaeu0yfhfjfBDee1ij/C1yvieufY8dO1Nf+5Qgu/T3/uhvPb7/P5Eee1mkvEvaus/de1Yh+9fjPBGeu55n/K5zfiUs/XI1/l6oPJ2nfI8c+1vmPGApPNJfe709/4qZ+txmvIxbOxpk/G7zvjJ2PpkkPBPge4+de1ymvK8z/h+o/O3y/hrlfHC0/nE1Pnn7vyYtfW+0Pnh6vypwfeIqfS6zvjU4PuOrvTD1PlqlPHu8/13nvJRgu/G1vlXh+81buydufacuPV8ofObuPUoZeu0JHlGAAAAAXRSTlMAQObYZgAAAAFiS0dEAIgFHUgAAAAHdElNRQfqCwYMAAAgqgXtAAADc0lEQVRo3u2a51MTURTFvRywRpAoTcCCUZpo0CUWEESIDZNQRcEKSmxYsKEYJIKKIhbsoigiKorY/kHf2yW7LIRJyGh2THJmMnn3zt6z57ebl3zYTJsWVKIAVVBAqkG1TuIfTq1z+IdT6xR+4tQ6hH8wtc7gJ06tI4QwQ5ghzBBmgOtfYYYhfJLC66kpTEcA033DnAHM9AFvFmaL73N0cz0GjYyaF62fv8ATkcrqL2PGxAJxkVMBjFdhqjQJZgIWJiYl6zwf6FE+Yy6CfjGWiGdemrLMsJxoxeLUuLR01slIjM1cyX3lThiyVkWvVmHywHzUmJbNizVrBb6QJ7hyTFhHtH4D0cbcvFTDpvXSlFzkF2wW9IVSc0tRWrF56VZSPGWrbSnj40wBczt2FGCzmDd1ZyFKdpHFmm4rRRnlmFFeUcl9XR12TNVua7UbzDhbCvbsZYuShBq+kCdE1cKwb38E4z2A7QfNyBWn5CLnEA7X1SeJzfgjqD3agCS74jkmzrxjelUc7zGPAydOgl/tMJwiakQRnSb7mQqcpXNAE53nvq4OO+aCa1CNeZEuAZfZYgY18YU8Iar5CgDTVcpGy7UIB+Ja+ZRcZAPN7BKIVg7gOjUDbYqny8oJ3CCHKo73mDY02CkTdfwk7UT16KCb5SaW6haVQSB+pulyhx1z2z1mJ8VI2cqkhTwxqjt36wFnFyTd41Ny0YVicll1oZvoPvBA8VTFuaOK4zXm6YXSuYofMl8bkQFFj1C8uq0Hj/nly6AnzFfujP3mUGOGU6uEKS2UCVGdbK+dacFdB6KrnU5ndat040YLdjefjrmbz+gp/3RNsGIfrkhqU8XxGvM5sLan50UVXjLfllcWmC73IrX9pZG58L05m28GuaPG1Ot0utfuMZUJUSWlB970oftSvAGN6W/fHZG2oatge9NYZ7W4mv3v+9Bon2iVkwfLQKYqjteYCdK9/4BB5vvxk5D5mey5yUMdVu6yjX3TfgG+Kh0VJteQe0xlQtSa4Yboyk/fiEa+1wjJugHJSC7yC8zCUKHUjLD1C3k/8smN1bXhbsNPhjnO3Ku9qWjiT9lG9uqFaYT+I/mA+Ws4q8iIKK2T/2vMcLOwbHDRb62T/13MwFAIM5AUwgwkBQ9mEHAG0TOUwOcMjgecwfG4Oij+faD1X1f8qz9VFfgKOvUvpgAAACV0RVh0ZGF0ZTpjcmVhdGUAMjAyNi0xMS0wNlQxMjowMDowMCswMDowMGIqurUAAAAldEVYdGRhdGU6bW9kaWZ5ADIwMjYtMTEtMDZUMTI6MDA6MDArMDA6MDATdwIJAAAAAElFTkSuQmCC';

const PNG_BTN_APAGAR_ULTIMA =
  'iVBORw0KGgoAAAANSUhEUgAAAOYAAAAoCAMAAADdVnZNAAACBFBMVEX////CQQzDQQzQSQzcUAzfUgzgUgzVTAzoVgzqWAzVSwzqWxH75Nf518TqXRP0q4X74tXrYhv97+n////xlGP3wKP4ybD74NL3w6jrZR/4zbf2uJn4zrj62srvhk/rYBj+9/T++/rrZiDym2798ev86uDzpHv1so/507/99fDwi1b2upr++PX86N30qYPsbiz50r775NjxkmDufEH0rIb51sP++/nxlGT74tb98uz51cL++vjtcC/0qYL98er0ron99fHwiVTzpn7++ff63M32upv1sY375dn3v6L+/fztdTb62Mb3wqf2uJjymGnqWxD75tr86+L3wKT98Or+/v7618XwiVPzpX386d/1tpX+/Pv5z7r97+j98+787uf50Lvvg0r99O/qXhX51MH75tvtdjfvgUjtdznvf0b0r4r628vqWQ3qXxb3xqzvhU3sby7vh1D3xq3ynnLtdDT1tJPsaibtdjj87eb3waXymGr3xav4ybH0rYj+/f3ymWv4x67saSXxkF3xl2jxlmbuf0X++fbrYRr749bwjFj4y7PvhEzwiFLqXBLxkV/uej7qXxfxk2L4y7TynHDwjlvvgkn87OP1tpb4yrL63s/0qIH50bzwjVn1tZT98+3sbi3uej399vL87OT2u5zrZB7sbCr62cfrYxz1s5D508D2vJ7ynG/qWg9CeMHqAAAAAXRSTlMAQObYZgAAAxxJREFUaN7tmodTE1EQxjlSNyAiXxBEIyBBQQJIF1BBY0MIAtIFUbGjIhYUMaIIauyCIvaC3X/Sd5dyaY6gl9x45zeTmc3O7tv95W1eZu4lJkZN4hQrP8hYjtNodXrFSafVcFysuJUGIylURpN3Q2M5rdzNRFJaz35yJrk7iawMnHtkFTuxbhmFseU0cvcRaWkETEV/M3lpBUyd3G1EWjoBUy93G5GW/j+mpIqLX5SgAszFiUvko4z+0CbBHGLNN2MBSX+HmQws/YMqKUglWga+Rb7PNCxnzhWWlb9PCo6LCmZ6BpC5aiHrx/0ac16fTbCigpkFazZWC9XW5OTGryXKy7Zl5hcwT+G6jKJioET0JKG0rLwiDOZ6MFUKHVeioKrasmGjdVMN+S0WgMnH8fVq8zfzb7ak2XkjMFZazK3Yth07hNK2mp2o20X1DQWOajTS7iY0t+zhMb0eFtPa1tAeBrOjEyldNR7M7r096N3XB+wnMTUMZqYjBwcOMqOu/xBvBMZKinkYOHIUOMaXPk50Al2UQAMnW3CKTgODdIbH9HpYzFDAxgQPrRszj84C5ygRHSSmhsE8TxeAYWYk0yBvBMZKiunAxQEqwiW+9AjRZYzSlWYnm8AqaoSd+EZKfB4Wc1WcgjGia8D1UMxxKoaNqA1san2pYTBvUDowwYxGtxEYKyVmQi8ETd5k1RxEt9B1G5MVEy5W7Q5QSHcZps/jf1oMofMe3QceuL0utodeTDPz57J5ZZhiahhMMz10Y7qNoFgpMR8BaS7X41ZMsWrdY/VwDk/DNjJVy6rx381UK8P0efwxnzjxtK0cFk/TM+hxZIViiqkeTKvFYukLjxkUKyVmv3vVZ5hl1Z6/sBe9pIFXrytHG3j/G3bSjgNvRY//2d9e9q77/Yc5D+bHT3b0hWKKqR5M8UQOwQyKlRLTT6E/X5/ZaxrOuYWuFB1Jhvnla+lMLb7JzRNpTHOTPXf2+w+5eaTF/Nf0H1NJ0qvpkZdKHmCq5HG0Oi4XYjiD3I1EViZOTRd/bGxNip1bo0G8l1fHpbxa/mKhfP0EUSCpl9u+80MAAAAASUVORK5CYII=';

const PNG_BTN_APAGAR_TUDO =
  'iVBORw0KGgoAAAANSUhEUgAAAOYAAAAoCAMAAADdVnZNAAAB+1BMVEX///+aGxuZGxubGxuxHx/EIiLLIyPFIiKyHx+5ICDYJSXcJia4ICDeNzf////jVVXxra3zubn88vL409P419f++/v88PDtkZHdMTH99fX++vrfODjcKyv1xcX30ND99PT87u787e364eHfPDzqfX364+PsiYn76+vum5v53NzmaGjwpaX99vbtkJDgQkLocXHtk5P2ysr++fnodXXztLTeNTXvoaH99/f53t7gQ0PrhYX9+PjysbHnbW3ztbX30tLwpqbumZn53d3ulpbxrKz+/PziTEz2zMzysLDwo6PpeXncKir65eXxrq777Oz+/v72y8vmZWX52tr64uLxqqr76Oj3z8/cJyf75+fysrLjUVHtkpLxqaniTU3kW1viTk7kWVnul5f30dHvoKDdLCzdLy/lYGDgRETlY2PmZ2fgQEDlX1/hR0fkWFjrh4fyr6/penrytLTzurrtlZX+/f32yMjfPT3pe3vqgYHocnLumJjeMjL529vmamr0u7v2ycnmZGTodHTgPz/iUFDiT0/rhITnbGzztrboc3P0v7/0vLzqf3/zt7f98/P64ODriIjjVFT42dnsjY32x8f0urr41NTsjo71w8Pnamrvn5/87+/peHjjUlLvnZ3eMzPhRUX1xsbqfHz88fHkXFzeNDTeNjbvnJzqfn7cKSnueRX0AAAAAXRSTlMAQObYZgAAA0VJREFUaN7tmudXE0EUxcNQQhi4lABCIAlFhEgRQQEJiooKoSvGKAIiIHZFxAqiWLAgWLEX7P6ZzuxmN0FWSnDd44Z7TnLIPW/fe7+8mcmHwWAIJoUQnSrEDzKUkLDwCB3KGElIqERJiCmK6lRR0YRIszRq3YyaMnrnSaK17kRdmYh4+uh2xYqKEs4hEql1H2orjI9T3zuTK1zAjNC6DbUVsYKpI61g6kn/GDMGiNUi0SIw44D4pfeRAFHm5WEmAknJ5lUpqmOmpAKWtKW0ls7frDabHRk2W9xyMTPNmUBWttqYq5GZgzXsj1zkWRMcaynNzymwZBUyp8ia6ljH68lOLorXl5R6n7RhA3vfWJZVXpHHxrFJipadygqn3VHpjVZ2E1FF6WYLtviVXSBRYJjV2LoN2wWEgpod2LmL1tYVukpQTxsq0NjUzOtJDotpaa3b7Y+Zvgdte93Y52HR+w+wwcTKTjsyD9bE53uXgKIrYtIOlDT4l50vUWCYnUDXIaCbIxymtAe9tI96jjQxhH7gKD3GMSWHxWzxPSpgHgdO0HaW4CRwin+KlZ0uuOtPS7tO2fVingHS5CILJAoM0wW3hzpwliMMUHoOg/R84xA7Wy7QejgpvcgxJYfFXPoN8zLsbFTAFSE6m0XLjqejhG37q2Kwsus/zVll/5woIMy+ZPHALB9mCC5KR9B7DeWl3aO4zr/WInqD1ZMdFjM2d5o36S32lffziXSLQxAd1uPtO8CwNE0FV8Qct+AulYsskCggzHvA/dHRBy2YYAhJD2sx1DmOgoGJSVaP780qvklkRwEz3YG2Djd6PA1TqH3k4FtKch4/cT1twmSDEKzs8pO2ulk4aX1l508UEOYzvhQpfY5phvDipdPxinpev8kYrONQ/Mh7C7zzOXMxaYyrzTn1/gOlH2fsI5/4EpecImuz5XPrF2+0sst+N4eSzYVs4/mKLJAoEEw/zUIQ9JW9xjH0bTFPa6tlYX6fKS6bRIfWDGpjjlU4E6Z//NSa4e9i/s9awdSTggozXOs21JZRwAzTug21FUmC53LBQExaN6KuokkwXfwZCInW7bqNMknXuMKlfKRR6wt0NRQe5ncpHyz/YqF//QL+7Al68He2CAAAAABJRU5ErkJggg==';

// ===============================================================
// MENU PERSONALIZADO (acionado automaticamente ao abrir a planilha)
// ===============================================================
function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createMenu('📚 Automação')
      .addItem('🔧 Corrigir alinhamento das colunas agora', 'corrigirPlanilhaManual')
      .addItem('🔘 Inserir os 3 botões + sincronização automática', 'criarBotoesNaPlanilha')
      .addSeparator()
      .addItem('🎯 Apagar linha selecionada', 'apagarLinhaSelecionada')
      .addItem('🗑️ Apagar última linha com dados', 'apagarUltimaLinha')
      .addItem('⚠️ Apagar todos os dados da planilha', 'apagarTodosOsDados')
      .addSeparator()
      .addItem('➕ Registrar operação', 'registrarOperacao')
      .addItem('✏️ Atualizar informações', 'atualizarInformacoes')
      .addItem('🧹 Limpar campos', 'limparCampos')
      .addSeparator()
      .addItem('🔍 Carregar pedido no painel (por ID)', 'carregarPedidoNoPainel')
      .addItem('🗑️ Limpar campos de um pedido (por ID)', 'limparCamposPedido')
      .addItem('🧹 Limpar apenas o painel (J2:J7)', 'limparPainel')
      .addItem('🛠️ Configurar estrutura da planilha', 'configurarPlanilha')
      .addToUi();
  } catch (err) {}

  try {
    const ss = (e && e.source) ? e.source : SpreadsheetApp.getActiveSpreadsheet();
    if (ss) corrigirLinhasDesalinhadas(ss);
  } catch (err) {}
}

// ===============================================================
// AUTO-CORREÇÃO DE LINHAS DESALINHADAS (CASO A IMPLANTAÇÃO ANTIGA RODE)
// ===============================================================
function extrairMetadadosLivro(texto) {
  const bruto = String(texto === undefined || texto === null ? '' : texto).trim();
  if (!bruto.includes(' || ')) {
    return { livro: bruto, email: '', endereco: '', pagamento: '' };
  }
  const partes = bruto.split(' || ').map(function(p) { return p.trim(); });
  return {
    livro: partes[0] || '',
    email: partes[1] || '',
    endereco: partes[2] || '',
    pagamento: partes[3] || ''
  };
}

function pareceDataOuTimestamp(valor) {
  if (valor instanceof Date) return true;
  const txt = String(valor === undefined || valor === null ? '' : valor).trim();
  return /^\d{2}\/\d{2}\/\d{4}/.test(txt);
}

function corrigirLinhasDesalinhadas(planilhaExplicita) {
  let corrigiuAlgo = false;
  try {
    const planilha = planilhaExplicita || SpreadsheetApp.getActiveSpreadsheet();
    if (!planilha) return false;

    // 1. Verifica e corrige a aba "Dados" (linhas antigas de 5 colunas -> 7 colunas)
    const aba = planilha.getSheetByName('Dados') || planilha.getSheets()[0];
    if (aba) {
      const cabDados = aba.getRange(1, 1, 1, COLUNAS_TABELA).getValues()[0];
      if (cabDados[0] !== 'ID' || cabDados[2] !== 'E-MAIL (GMAIL)') {
        aba.getRange(1, 1, 1, COLUNAS_TABELA).setValues([CABECALHO_DADOS]);
      }

      const maxDados = aba.getLastRow();
      if (maxDados >= 2) {
        const blocoDados = aba.getRange(2, 1, maxDados - 1, COLUNAS_TABELA).getValues();
        for (let i = 0; i < blocoDados.length; i++) {
          const r = blocoDados[i];
          const colF = r[5];
          const colG = r[6];
          const valorAntigo = converterValor(r[3]);
          // Linha gravada pelo Web App antigo: F e G vazias, D é número (valor) e E é data/hora
          if (
            (colF === '' || colF === null) &&
            (colG === '' || colG === null) &&
            r[0] !== '' &&
            r[3] !== '' &&
            !isNaN(valorAntigo) &&
            pareceDataOuTimestamp(r[4])
          ) {
            const linPlanilha = i + 2;
            const meta = extrairMetadadosLivro(r[2]);
            aba.getRange(linPlanilha, 2, 1, 4).setNumberFormat('@').setHorizontalAlignment('left');
            aba.getRange(linPlanilha, 1, 1, COLUNAS_TABELA).setValues([[
              r[0],
              r[1],
              meta.email || '-',
              meta.endereco || '-',
              meta.livro || r[2],
              valorAntigo,
              r[4]
            ]]);
            formatarBlocoPedidos(aba, linPlanilha, 1);
            corrigiuAlgo = true;
          }
        }
      }
    }

    // 2. Verifica e corrige a aba "Log" (linhas antigas de 8 colunas -> 9 colunas)
    const log = planilha.getSheetByName('Log');
    if (log) {
      const cabLog = log.getRange(1, 1, 1, 9).getValues()[0];
      if (cabLog[0] !== 'DATA/HORA' || cabLog[1] !== 'E-MAIL (GMAIL)' || cabLog[8] !== 'OBSERVAÇÃO') {
        log.getRange(1, 1, 1, 9).setValues([CABECALHO_LOG]);
      }

      const maxLog = log.getLastRow();
      if (maxLog >= 2) {
        const blocoLog = log.getRange(2, 1, maxLog - 1, 9).getValues();
        for (let j = 0; j < blocoLog.length; j++) {
          const l = blocoLog[j];
          const ehAntigo =
            (l[8] === '' || l[8] === null) &&
            (l[1] === 'Site (Web App)' || l[1] === 'Planilha (Menu/Botão)');
          if (ehAntigo) {
            const linLog = j + 2;
            const metaLog = extrairMetadadosLivro(l[5]);
            const emailLog =
              metaLog.email ||
              (l[1] === 'Planilha (Menu/Botão)' ? (obterEmailUsuarioPlanilha() || 'Usuário da Planilha') : '-');
            const enderecoLog = metaLog.endereco || '-';
            const livroLog = metaLog.livro || l[5] || '-';
            const valorNumLog = (l[6] === '' || l[6] === null || isNaN(converterValor(l[6])))
              ? ''
              : converterValor(l[6]);
            const obsLog = metaLog.pagamento
              ? ('Pagamento: ' + metaLog.pagamento + ' — Pedido pelo site')
              : (String(l[7] || '').replace('automaticamente ', '') || 'Pedido recebido pelo site');

            log.getRange(linLog, 2, 1, 3).setNumberFormat('@').setHorizontalAlignment('left');
            log.getRange(linLog, 6, 1, 2).setNumberFormat('@').setHorizontalAlignment('left');
            log.getRange(linLog, 9, 1, 1).setNumberFormat('@').setHorizontalAlignment('left');
            log.getRange(linLog, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss').setHorizontalAlignment('center');
            log.getRange(linLog, 5).setNumberFormat('0').setHorizontalAlignment('center');
            if (valorNumLog !== '') {
              log.getRange(linLog, 8).setNumberFormat('R$ #,##0.00').setHorizontalAlignment('center');
            }

            const intervaloLog = log.getRange(linLog, 1, 1, 9);
            intervaloLog.setValues([[
              l[0],
              emailLog,
              enderecoLog,
              l[2] || 'Registro pelo Site',
              l[3] || '-',
              l[4] || '-',
              livroLog,
              valorNumLog,
              obsLog
            ]]);
            try {
              intervaloLog.setBorder(true, true, true, true, true, true);
            } catch (e) {}
            corrigiuAlgo = true;
          }
        }
      }
    }
  } catch (err) {
    Logger.log('Aviso ao alinhar linhas: ' + err);
  }
  return corrigiuAlgo;
}

function sincronizarAutomaticamente() {
  const planilha = obterPlanilha(true);
  corrigirLinhasDesalinhadas(planilha);
}

function corrigirPlanilhaManual() {
  const planilha = obterPlanilha(false);
  corrigirLinhasDesalinhadas(planilha);
  instalarGatilhoAutomatico(planilha);
  mostrarToast('Colunas das abas Dados e Log verificadas e alinhadas!', '🔧 Alinhamento Concluído', planilha);
}

function instalarGatilhoAutomatico(planilha) {
  try {
    if (typeof ScriptApp === 'undefined' || !ScriptApp.getProjectTriggers) return;
    const gatilhos = ScriptApp.getProjectTriggers();
    for (let i = 0; i < gatilhos.length; i++) {
      if (gatilhos[i].getHandlerFunction() === 'sincronizarAutomaticamente') {
        return; // Já instalado
      }
    }
    const alvo = planilha || obterPlanilha(false);
    ScriptApp.newTrigger('sincronizarAutomaticamente')
      .forSpreadsheet(alvo)
      .onChange()
      .create();
  } catch (e) {
    // Sem permissão em gatilho simples; funciona normalmente quando chamado pelo menu
  }
}

// ===============================================================
// MEMORIZA A ÚLTIMA LINHA SELECIONADA NA TABELA (A:G) PARA O J10
// ===============================================================
function salvarSelecaoRecente(linhas) {
  const payload = JSON.stringify({ linhas: linhas, ts: Date.now() });
  try {
    CacheService.getScriptCache().put(CHAVE_CACHE_SELECAO, payload, 600);
  } catch (e) {}
  try {
    PropertiesService.getScriptProperties().setProperty(CHAVE_CACHE_SELECAO, payload);
  } catch (e) {}
}

function lerSelecaoRecente() {
  let texto = '';
  try {
    texto = CacheService.getScriptCache().get(CHAVE_CACHE_SELECAO) || '';
  } catch (e) {}
  if (!texto) {
    try {
      texto = PropertiesService.getScriptProperties().getProperty(CHAVE_CACHE_SELECAO) || '';
    } catch (e) {}
  }
  if (!texto) return [];
  try {
    const obj = JSON.parse(texto);
    if (!obj || !Array.isArray(obj.linhas)) return [];
    if (Date.now() - Number(obj.ts || 0) > 600000) return [];
    return obj.linhas;
  } catch (e) {
    return [];
  }
}

function limparSelecaoRecente() {
  try {
    CacheService.getScriptCache().remove(CHAVE_CACHE_SELECAO);
  } catch (e) {}
  try {
    PropertiesService.getScriptProperties().deleteProperty(CHAVE_CACHE_SELECAO);
  } catch (e) {}
}

function extrairLinhasDeIntervalos(intervalos) {
  const mapa = {};
  for (let i = 0; i < intervalos.length; i++) {
    const r = intervalos[i];
    if (!r) continue;
    // Considera seleções que tocam as colunas A:G (1..7), incluindo clique no número da linha
    if (r.getColumn() > COLUNAS_TABELA) continue;
    const inicio = r.getRow();
    const fim = inicio + r.getNumRows() - 1;
    for (let lin = inicio; lin <= fim; lin++) {
      if (lin >= 2) {
        mapa[lin] = true;
      }
    }
  }
  return Object.keys(mapa).map(Number).sort(function(a, b) { return a - b; });
}

function onSelectionChange(e) {
  if (!e || !e.range) return;
  const aba = e.range.getSheet();
  try {
    const ss = e.source || aba.getParent();
    if (ss) corrigirLinhasDesalinhadas(ss);
  } catch (err) {}

  if (aba.getName() !== 'Dados' && aba.getIndex() !== 1) return;

  let intervalos = [e.range];
  try {
    const lista = aba.getActiveRangeList();
    if (lista && lista.getRanges) {
      intervalos = lista.getRanges();
    }
  } catch (err) {}

  const linhas = extrairLinhasDeIntervalos(intervalos);
  if (linhas.length > 0) {
    salvarSelecaoRecente(linhas);
  }
}

// ===============================================================
// ACIONAMENTO IMEDIATO PELOS BOTÕES DE CAIXA (J10, J11 e J12)
// ===============================================================
function onEdit(e) {
  if (!e || !e.range || !e.source) return;
  corrigirLinhasDesalinhadas(e.source);

  const aba = e.range.getSheet();
  const nomeAba = aba.getName();
  if (nomeAba !== 'Dados' && aba.getIndex() !== 1) return;

  const col = e.range.getColumn();
  const lin = e.range.getRow();
  if (
    col !== COLUNA_PAINEL_INPUT ||
    (lin !== LINHA_CHECK_SELECIONADA && lin !== LINHA_CHECK_ULTIMA && lin !== LINHA_CHECK_TUDO)
  ) {
    return;
  }

  if (e.range.getValue() !== true) return;

  // Desmarca imediatamente para funcionar como botão de clique
  e.range.setValue(false);

  if (lin === LINHA_CHECK_SELECIONADA) {
    const linhasSalvas = lerSelecaoRecente();
    limparSelecaoRecente();
    if (linhasSalvas.length === 0) {
      mostrarToast(
        'Clique primeiro na linha que deseja apagar (colunas A:G) ou use o botão azul "Apagar Linha Selecionada" em I14!',
        '🎯 Apagar Linha Selecionada',
        e.source
      );
      return;
    }
    executarApagarLinhasSelecionadas(linhasSalvas, e.source);
  } else if (lin === LINHA_CHECK_ULTIMA) {
    executarApagarUltimaLinha(e.source);
  } else if (lin === LINHA_CHECK_TUDO) {
    executarApagarTodosOsDados(e.source);
  }
}

// ===============================================================
// CRIA OS 3 BOTÕES GRÁFICOS CLICÁVEIS NA PLANILHA
// ===============================================================
function criarBotoesNaPlanilha() {
  const planilha = obterPlanilha(false);
  corrigirLinhasDesalinhadas(planilha);
  instalarGatilhoAutomatico(planilha);

  const aba = abaPedidos(false, planilha);

  // Remove versões anteriores dos mesmos botões para não duplicar
  const imagens = aba.getImages();
  for (let i = 0; i < imagens.length; i++) {
    const titulo = imagens[i].getAltTextTitle();
    if (
      titulo === 'BTN_APAGAR_SELECIONADA' ||
      titulo === 'BTN_APAGAR_ULTIMA' ||
      titulo === 'BTN_APAGAR_TUDO'
    ) {
      imagens[i].remove();
    }
  }

  // Botão 1: Apagar Linha Selecionada (coluna I = 9, linha 14)
  const blobSel = Utilities.newBlob(
    Utilities.base64Decode(PNG_BTN_APAGAR_SELECIONADA),
    'image/png',
    'botao_apagar_selecionada.png'
  );
  const imgSel = aba.insertImage(blobSel, COLUNA_PAINEL_BOTAO_IMG, 14, 0, 4);
  imgSel.assignScript('apagarLinhaSelecionada');
  imgSel.setAltTextTitle('BTN_APAGAR_SELECIONADA');

  // Botão 2: Apagar Última Linha (coluna I = 9, linha 17)
  const blobUltima = Utilities.newBlob(
    Utilities.base64Decode(PNG_BTN_APAGAR_ULTIMA),
    'image/png',
    'botao_apagar_ultima.png'
  );
  const imgUltima = aba.insertImage(blobUltima, COLUNA_PAINEL_BOTAO_IMG, 17, 0, 4);
  imgUltima.assignScript('apagarUltimaLinha');
  imgUltima.setAltTextTitle('BTN_APAGAR_ULTIMA');

  // Botão 3: Apagar Todos os Dados (coluna I = 9, linha 20)
  const blobTudo = Utilities.newBlob(
    Utilities.base64Decode(PNG_BTN_APAGAR_TUDO),
    'image/png',
    'botao_apagar_tudo.png'
  );
  const imgTudo = aba.insertImage(blobTudo, COLUNA_PAINEL_BOTAO_IMG, 20, 0, 4);
  imgTudo.assignScript('apagarTodosOsDados');
  imgTudo.setAltTextTitle('BTN_APAGAR_TUDO');

  mostrarToast('Os 3 botões gráficos e a sincronização automática foram ativados!', '🔘 Botões Criados', planilha);
}

// ===============================================================
// FUNÇÕES AUXILIARES DE PLANILHA, E-MAIL, VALIDAÇÃO E LOG
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

function obterEmailUsuarioPlanilha() {
  try {
    const ativo = Session.getActiveUser().getEmail();
    if (ativo) return ativo;
  } catch (e) {}
  try {
    const efetivo = Session.getEffectiveUser().getEmail();
    if (efetivo) return efetivo;
  } catch (e) {}
  return '';
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
    log.getRange(1, 1, 1, 9).setValues([CABECALHO_LOG]);
  }
  return log;
}

function registrarLog(acao, id, nome, livro, valor, observacao, origem, planilhaExplicita, email, endereco) {
  try {
    const log = abaLog(origem === 'site', planilhaExplicita);
    const linha = log.getLastRow() + 1;
    const valorCelula = valor === '' || valor === undefined || valor === null || isNaN(Number(valor))
      ? ''
      : Number(valor);

    const emailFinal = textoSeguro(email, 120) ||
      obterEmailUsuarioPlanilha() ||
      (origem === 'site' ? 'Sem e-mail informado' : 'Usuário da Planilha');
    const enderecoFinal = textoSeguro(endereco, 200) || '-';

    log.getRange(linha, 2, 1, 3).setNumberFormat('@').setHorizontalAlignment('left');
    log.getRange(linha, 6, 1, 2).setNumberFormat('@').setHorizontalAlignment('left');
    log.getRange(linha, 9, 1, 1).setNumberFormat('@').setHorizontalAlignment('left');
    log.getRange(linha, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss').setHorizontalAlignment('center');
    if (valorCelula !== '') {
      log.getRange(linha, 8).setNumberFormat('R$ #,##0.00');
    }

    const intervaloLinha = log.getRange(linha, 1, 1, 9);
    intervaloLinha.setValues([[
      new Date(),
      emailFinal,
      enderecoFinal,
      acao,
      id || '-',
      nome || '-',
      livro || '-',
      valorCelula,
      observacao || ''
    ]]);
    try {
      log.getRange(linha, 5).setHorizontalAlignment('center');
      log.getRange(linha, 8).setHorizontalAlignment('center');
      intervaloLinha.setBorder(true, true, true, true, true, true);
    } catch (e) {}
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

function linhaTemDados(linhaArray) {
  return linhaArray.some(function(celula) {
    return celula !== '' && celula !== null && celula !== undefined;
  });
}

// Lê apenas as colunas A:G (ignora o Painel de Operações e botões em I:J).
function estadoTabelaPedidos(aba) {
  const maxLinha = aba.getLastRow();
  if (maxLinha < 2) {
    return { ultimaLinha: 1, proximoId: 1 };
  }

  const dados = aba.getRange(2, 1, maxLinha - 1, COLUNAS_TABELA).getValues();
  let ultimaLinha = 1;
  let maiorId = 0;

  for (let i = 0; i < dados.length; i++) {
    if (linhaTemDados(dados[i])) {
      ultimaLinha = i + 2;
    }
    const idNum = Number(dados[i][0]);
    if (dados[i][0] !== '' && !isNaN(idNum) && idNum > maiorId) {
      maiorId = idNum;
    }
  }

  return { ultimaLinha: ultimaLinha, proximoId: maiorId + 1 };
}

// Acha o número da linha do pedido com este ID nas colunas A:G (ou 0 se não existir).
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

function formatarBlocoPedidos(aba, linhaInicial, qtdLinhas) {
  try {
    aba.getRange(linhaInicial, 6, qtdLinhas, 1).setNumberFormat('R$ #,##0.00');
    aba.getRange(linhaInicial, 7, qtdLinhas, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
    aba.getRange(linhaInicial, 1, qtdLinhas, 1).setHorizontalAlignment('center');
    aba.getRange(linhaInicial, 6, qtdLinhas, 2).setHorizontalAlignment('center');
    aba.getRange(linhaInicial, 1, qtdLinhas, COLUNAS_TABELA).setBorder(true, true, true, true, true, true);
  } catch (erro) {
    Logger.log('Aviso ao formatar linha: ' + erro);
  }
}

// ===============================================================
// BOTÃO 1: APAGAR A(S) LINHA(S) SELECIONADA(S) NA TABELA
// ===============================================================
function executarApagarLinhasSelecionadas(linhasAlvo, planilhaExplicita) {
  const planilha = obterPlanilha(false, planilhaExplicita);
  corrigirLinhasDesalinhadas(planilha);
  const aba = abaPedidos(false, planilha);
  const estado = estadoTabelaPedidos(aba);

  if (estado.ultimaLinha <= 1) {
    mostrarToast('Não há dados na tabela para apagar.', '🎯 Apagar Linha Selecionada', planilha);
    return { ok: false, erro: 'planilha vazia' };
  }

  const mapaApagar = {};
  for (let i = 0; i < linhasAlvo.length; i++) {
    const n = Number(linhasAlvo[i]);
    if (n >= 2) mapaApagar[n] = true;
  }

  const totalLinhas = estado.ultimaLinha - 1;
  const todasAsLinhas = aba.getRange(2, 1, totalLinhas, COLUNAS_TABELA).getValues();
  const removidas = [];
  const mantidas = [];

  for (let i = 0; i < todasAsLinhas.length; i++) {
    const numLinha = i + 2;
    const conteudo = todasAsLinhas[i];
    if (mapaApagar[numLinha]) {
      if (linhaTemDados(conteudo)) {
        removidas.push({ linha: numLinha, dados: conteudo });
      }
    } else if (linhaTemDados(conteudo)) {
      mantidas.push(conteudo);
    }
  }

  if (removidas.length === 0) {
    mostrarToast(
      'A linha selecionada está vazia. Selecione uma linha com dados nas colunas A:G.',
      '🎯 Apagar Linha Selecionada',
      planilha
    );
    return { ok: false, erro: 'a linha selecionada está vazia' };
  }

  // Limpa o bloco anterior em A:G sem mexer nas colunas I:J do painel/botões
  const blocoAntigo = aba.getRange(2, 1, totalLinhas, COLUNAS_TABELA);
  blocoAntigo.clearContent();
  try {
    blocoAntigo.setBorder(false, false, false, false, false, false);
  } catch (e) {}

  // Regrava as linhas restantes compactadas (sem buracos em branco)
  if (mantidas.length > 0) {
    aba.getRange(2, 2, mantidas.length, 4).setNumberFormat('@');
    aba.getRange(2, 1, mantidas.length, COLUNAS_TABELA).setValues(mantidas);
    formatarBlocoPedidos(aba, 2, mantidas.length);
  }

  for (let j = 0; j < removidas.length; j++) {
    const item = removidas[j];
    const d = item.dados;
    registrarLog(
      'Apagar Linha Selecionada',
      d[0] || '-',
      d[1] || '-',
      d[4] || '-',
      d[5],
      'Linha ' + item.linha + ' apagada ao selecionar na planilha',
      'planilha',
      planilha,
      d[2] !== '-' ? d[2] : '',
      d[3]
    );
  }

  const descLinhas = removidas.map(function(r) { return r.linha; }).join(', ');
  mostrarToast(
    removidas.length === 1
      ? 'Linha ' + descLinhas + ' apagada com sucesso!'
      : 'Linhas ' + descLinhas + ' apagadas com sucesso!',
    '🎯 Apagar Linha Selecionada',
    planilha
  );

  return {
    ok: true,
    linhas: removidas.map(function(r) { return r.linha; }),
    ids: removidas.map(function(r) { return r.dados[0]; })
  };
}

function apagarLinhaSelecionada() {
  const planilha = obterPlanilha(false);
  corrigirLinhasDesalinhadas(planilha);
  const aba = abaPedidos(false, planilha);
  let intervalos = [];

  try {
    const lista = aba.getActiveRangeList();
    if (lista && lista.getRanges) {
      intervalos = lista.getRanges();
    }
  } catch (e) {}

  if (intervalos.length === 0) {
    try {
      const unico = aba.getActiveRange();
      if (unico) intervalos = [unico];
    } catch (e) {}
  }

  const linhas = extrairLinhasDeIntervalos(intervalos);
  if (linhas.length === 0) {
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        'Selecione uma linha da tabela',
        'Clique primeiro em qualquer célula da linha que deseja apagar (colunas A até G, a partir da linha 2) e depois clique no botão "Apagar Linha Selecionada".',
        ui.ButtonSet.OK
      );
    } catch (e) {}
    return { ok: false, erro: 'nenhuma linha da tabela selecionada' };
  }

  limparSelecaoRecente();
  const resultado = executarApagarLinhasSelecionadas(linhas, planilha);
  if (!resultado.ok && arguments.length === 0) {
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        'Aviso',
        resultado.erro === 'planilha vazia'
          ? 'Não há nenhuma linha com informações na tabela.'
          : 'A linha selecionada já está vazia. Clique em uma linha preenchida (colunas A até G).',
        ui.ButtonSet.OK
      );
    } catch (e) {}
  }
  return resultado;
}

// ===============================================================
// BOTÃO 2: APAGAR A ÚLTIMA LINHA QUE TEM INFORMAÇÕES
// ===============================================================
function executarApagarUltimaLinha(planilhaExplicita) {
  const planilha = obterPlanilha(false, planilhaExplicita);
  corrigirLinhasDesalinhadas(planilha);
  const aba = abaPedidos(false, planilha);
  const estado = estadoTabelaPedidos(aba);

  if (estado.ultimaLinha <= 1) {
    mostrarToast('Não há nenhuma linha com informações para apagar.', '🗑️ Apagar Última Linha', planilha);
    return { ok: false, erro: 'nenhuma linha com informações' };
  }

  const linhaAlvo = estado.ultimaLinha;
  const intervalo = aba.getRange(linhaAlvo, 1, 1, COLUNAS_TABELA);
  const removidos = intervalo.getValues()[0];

  intervalo.clearContent();
  try {
    intervalo.setBorder(false, false, false, false, false, false);
  } catch (e) {}

  registrarLog(
    'Apagar Última Linha',
    removidos[0] || '-',
    removidos[1] || '-',
    removidos[4] || '-',
    removidos[5],
    'Última linha preenchida (linha ' + linhaAlvo + ') apagada via botão',
    'planilha',
    planilha,
    removidos[2] !== '-' ? removidos[2] : '',
    removidos[3]
  );

  mostrarToast(
    'Linha ' + linhaAlvo + ' apagada com sucesso!',
    '🗑️ Apagar Última Linha',
    planilha
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
// BOTÃO 3: APAGAR TODOS OS DADOS DA PLANILHA
// ===============================================================
function executarApagarTodosOsDados(planilhaExplicita) {
  const planilha = obterPlanilha(false, planilhaExplicita);
  corrigirLinhasDesalinhadas(planilha);
  const aba = abaPedidos(false, planilha);
  const estado = estadoTabelaPedidos(aba);

  if (estado.ultimaLinha > 1) {
    const intervalo = aba.getRange(2, 1, estado.ultimaLinha - 1, COLUNAS_TABELA);
    intervalo.clearContent();
    try {
      intervalo.setBorder(false, false, false, false, false, false);
    } catch (e) {}
  }

  // Também limpa os campos digitados no Painel (J2:J7), mantendo os rótulos e cabeçalhos
  aba.getRange(INTERVALO_PAINEL).clearContent();
  limparSelecaoRecente();

  registrarLog(
    'Apagar Todos os Dados',
    '-',
    '-',
    '-',
    '',
    'Todos os dados da planilha foram apagados via botão',
    'planilha',
    planilha
  );

  mostrarToast('Todos os dados da planilha foram apagados!', '⚠️ Apagar Todos os Dados', planilha);
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
function registrarOperacao(nome, livro, valor, origem, email, endereco, pagamento) {
  if (arguments.length === 0) {
    return acionarRegistrarOperacao();
  }

  const meta = extrairMetadadosLivro(livro);
  const nomeLimpo = textoSeguro(nome, 80);
  const livroLimpo = textoSeguro(meta.livro || livro, 300);
  const valorNum = converterValor(valor);

  if (!nomeLimpo || !livroLimpo || isNaN(valorNum) || !(valorNum >= 0)) {
    return { ok: false, erro: 'dados inválidos' };
  }

  const ehSite = origem === 'site';
  const emailLimpo =
    textoSeguro(email || meta.email, 120) ||
    (!ehSite ? obterEmailUsuarioPlanilha() : '');
  const enderecoLimpo = textoSeguro(endereco || meta.endereco, 200);
  const pagamentoLimpo = textoSeguro(pagamento || meta.pagamento, 40);

  const planilha = obterPlanilha(ehSite);
  corrigirLinhasDesalinhadas(planilha);
  const aba = abaPedidos(ehSite, planilha);
  const estado = estadoTabelaPedidos(aba);
  const linha = estado.ultimaLinha + 1;
  const id = estado.proximoId;
  const agora = new Date();

  aba.getRange(linha, 2, 1, 4).setNumberFormat('@').setHorizontalAlignment('left');
  aba.getRange(linha, 1, 1, COLUNAS_TABELA).setValues([[
    id,
    nomeLimpo,
    emailLimpo || '-',
    enderecoLimpo || '-',
    livroLimpo,
    valorNum,
    agora
  ]]);
  formatarBlocoPedidos(aba, linha, 1);

  const obs = ehSite
    ? (pagamentoLimpo ? 'Pagamento: ' + pagamentoLimpo + ' — Pedido pelo site' : 'Pedido recebido pelo site')
    : 'Operação registrada via menu/botão';

  registrarLog(
    ehSite ? 'Registro pelo Site' : 'Registro Manual',
    id,
    nomeLimpo,
    livroLimpo,
    valorNum,
    obs,
    origem,
    planilha,
    emailLimpo,
    enderecoLimpo
  );

  return { ok: true, id: id, versao: 2 };
}

function acionarRegistrarOperacao() {
  const ui = SpreadsheetApp.getUi();
  const planilha = obterPlanilha(false);
  corrigirLinhasDesalinhadas(planilha);
  const aba = abaPedidos(false, planilha);

  let nome = textoSeguro(aba.getRange(CELULA_NOME).getValue(), 80);
  let email = textoSeguro(aba.getRange(CELULA_EMAIL).getValue(), 120);
  let endereco = textoSeguro(aba.getRange(CELULA_ENDERECO).getValue(), 200);
  let livro = textoSeguro(aba.getRange(CELULA_LIVRO).getValue(), 300);
  const brutoValor = aba.getRange(CELULA_VALOR).getValue();
  let valorTexto = String(brutoValor === undefined || brutoValor === null ? '' : brutoValor).trim();

  const usouPainel = Boolean(nome || email || endereco || livro || valorTexto);

  if (!usouPainel) {
    const respNome = ui.prompt('➕ Registrar Operação (1/5)', 'Informe o nome do cliente:', ui.ButtonSet.OK_CANCEL);
    if (respNome.getSelectedButton() !== ui.Button.OK) return;
    nome = textoSeguro(respNome.getResponseText(), 80);

    const respEmail = ui.prompt('➕ Registrar Operação (2/5)', 'Informe o e-mail (Gmail) do cliente:', ui.ButtonSet.OK_CANCEL);
    if (respEmail.getSelectedButton() !== ui.Button.OK) return;
    email = textoSeguro(respEmail.getResponseText(), 120);

    const respEnd = ui.prompt('➕ Registrar Operação (3/5)', 'Informe o endereço de entrega:', ui.ButtonSet.OK_CANCEL);
    if (respEnd.getSelectedButton() !== ui.Button.OK) return;
    endereco = textoSeguro(respEnd.getResponseText(), 200);

    const respLivro = ui.prompt('➕ Registrar Operação (4/5)', 'Informe o(s) livro(s):', ui.ButtonSet.OK_CANCEL);
    if (respLivro.getSelectedButton() !== ui.Button.OK) return;
    livro = textoSeguro(respLivro.getResponseText(), 300);

    const respValor = ui.prompt('➕ Registrar Operação (5/5)', 'Informe o valor total (ex.: 49,90):', ui.ButtonSet.OK_CANCEL);
    if (respValor.getSelectedButton() !== ui.Button.OK) return;
    valorTexto = respValor.getResponseText();
  }

  const valorNum = converterValor(valorTexto);
  if (!nome || !livro || isNaN(valorNum) || !(valorNum >= 0)) {
    ui.alert(
      'Dados inválidos',
      'Preencha pelo menos Nome do Cliente (J3), Livro (J6) e um Valor válido >= 0 (J7) antes de registrar.',
      ui.ButtonSet.OK
    );
    return;
  }

  const resultado = registrarOperacao(nome, livro, valorNum, 'planilha', email, endereco);
  if (!resultado.ok) {
    ui.alert('Erro ao registrar', resultado.erro, ui.ButtonSet.OK);
    return;
  }

  aba.getRange(INTERVALO_PAINEL).clearContent();
  mostrarToast('Pedido #' + resultado.id + ' registrado com sucesso!', '📚 Automação', planilha);
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

  const planilha = obterPlanilha(false);
  corrigirLinhasDesalinhadas(planilha);
  const aba = abaPedidos(false, planilha);
  const linha = acharLinha(aba, id);
  if (!linha) return { ok: false, erro: 'pedido não encontrado' };

  if (campos.nome !== undefined) {
    const novoNome = textoSeguro(campos.nome, 80);
    if (!novoNome) return { ok: false, erro: 'nome inválido' };
    const c = aba.getRange(linha, 2);
    c.setNumberFormat('@');
    c.setValue(novoNome);
  }

  if (campos.email !== undefined) {
    const novoEmail = textoSeguro(campos.email, 120);
    if (!novoEmail) return { ok: false, erro: 'e-mail inválido' };
    const c = aba.getRange(linha, 3);
    c.setNumberFormat('@');
    c.setValue(novoEmail);
  }

  if (campos.endereco !== undefined) {
    const novoEndereco = textoSeguro(campos.endereco, 200);
    if (!novoEndereco) return { ok: false, erro: 'endereço inválido' };
    const c = aba.getRange(linha, 4);
    c.setNumberFormat('@');
    c.setValue(novoEndereco);
  }

  if (campos.livro !== undefined) {
    const novoLivro = textoSeguro(campos.livro, 300);
    if (!novoLivro) return { ok: false, erro: 'livro inválido' };
    const c = aba.getRange(linha, 5);
    c.setNumberFormat('@');
    c.setValue(novoLivro);
  }

  if (campos.valor !== undefined) {
    const valorNum = converterValor(campos.valor);
    if (isNaN(valorNum) || !(valorNum >= 0)) return { ok: false, erro: 'valor inválido' };
    const c = aba.getRange(linha, 6);
    c.setNumberFormat('R$ #,##0.00');
    c.setValue(valorNum);
  }

  aba.getRange(linha, 7)
    .setNumberFormat('dd/mm/yyyy hh:mm:ss')
    .setValue(new Date());

  const atualizado = aba.getRange(linha, 1, 1, 6).getValues()[0];
  registrarLog(
    'Atualização',
    atualizado[0],
    atualizado[1],
    atualizado[4],
    atualizado[5],
    'Informações atualizadas via menu/botão',
    'planilha',
    planilha,
    atualizado[2] !== '-' ? atualizado[2] : '',
    atualizado[3]
  );

  return { ok: true, id: Number(id) };
}

function acionarAtualizarInformacoes() {
  const ui = SpreadsheetApp.getUi();
  const planilha = obterPlanilha(false);
  corrigirLinhasDesalinhadas(planilha);
  const aba = abaPedidos(false, planilha);

  const brutoId = aba.getRange(CELULA_ID).getValue();
  let idTexto = String(brutoId === undefined || brutoId === null ? '' : brutoId).trim();
  const nomePainel = textoSeguro(aba.getRange(CELULA_NOME).getValue(), 80);
  const emailPainel = textoSeguro(aba.getRange(CELULA_EMAIL).getValue(), 120);
  const enderecoPainel = textoSeguro(aba.getRange(CELULA_ENDERECO).getValue(), 200);
  const livroPainel = textoSeguro(aba.getRange(CELULA_LIVRO).getValue(), 300);
  const brutoValor = aba.getRange(CELULA_VALOR).getValue();
  const valorPainelTexto = String(brutoValor === undefined || brutoValor === null ? '' : brutoValor).trim();

  if (idTexto && (nomePainel || emailPainel || enderecoPainel || livroPainel || valorPainelTexto)) {
    const camposPainel = {};
    if (nomePainel) camposPainel.nome = nomePainel;
    if (emailPainel) camposPainel.email = emailPainel;
    if (enderecoPainel) camposPainel.endereco = enderecoPainel;
    if (livroPainel) camposPainel.livro = livroPainel;
    if (valorPainelTexto) {
      const v = converterValor(valorPainelTexto);
      if (isNaN(v) || !(v >= 0)) {
        ui.alert('Valor inválido', 'Informe um valor numérico válido em J7.', ui.ButtonSet.OK);
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
    mostrarToast('Pedido #' + resPainel.id + ' atualizado!', '📚 Automação', planilha);
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

  const atual = aba.getRange(linha, 1, 1, 6).getValues()[0];
  const campos = {};

  const respNome = ui.prompt(
    '✏️ Atualizar Pedido #' + atual[0] + ' (1/5)',
    'Nome atual: ' + (atual[1] || '(vazio)') + '\nDigite o novo nome ou deixe em branco para manter:',
    ui.ButtonSet.OK_CANCEL
  );
  if (respNome.getSelectedButton() !== ui.Button.OK) return;
  if (respNome.getResponseText().trim()) campos.nome = respNome.getResponseText().trim();

  const respEmail = ui.prompt(
    '✏️ Atualizar Pedido #' + atual[0] + ' (2/5)',
    'E-mail atual: ' + (atual[2] || '(vazio)') + '\nDigite o novo e-mail ou deixe em branco para manter:',
    ui.ButtonSet.OK_CANCEL
  );
  if (respEmail.getSelectedButton() !== ui.Button.OK) return;
  if (respEmail.getResponseText().trim()) campos.email = respEmail.getResponseText().trim();

  const respEnd = ui.prompt(
    '✏️ Atualizar Pedido #' + atual[0] + ' (3/5)',
    'Endereço atual: ' + (atual[3] || '(vazio)') + '\nDigite o novo endereço ou deixe em branco para manter:',
    ui.ButtonSet.OK_CANCEL
  );
  if (respEnd.getSelectedButton() !== ui.Button.OK) return;
  if (respEnd.getResponseText().trim()) campos.endereco = respEnd.getResponseText().trim();

  const respLivro = ui.prompt(
    '✏️ Atualizar Pedido #' + atual[0] + ' (4/5)',
    'Livro atual: ' + (atual[4] || '(vazio)') + '\nDigite o novo livro ou deixe em branco para manter:',
    ui.ButtonSet.OK_CANCEL
  );
  if (respLivro.getSelectedButton() !== ui.Button.OK) return;
  if (respLivro.getResponseText().trim()) campos.livro = respLivro.getResponseText().trim();

  const respValor = ui.prompt(
    '✏️ Atualizar Pedido #' + atual[0] + ' (5/5)',
    'Valor atual: ' + (atual[5] !== '' ? 'R$ ' + atual[5] : '(vazio)') + '\nDigite o novo valor ou deixe em branco para manter:',
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
  mostrarToast('Pedido #' + resultado.id + ' atualizado!', '📚 Automação', planilha);
  ui.alert('Sucesso!', 'Informações do pedido #' + resultado.id + ' atualizadas com sucesso.', ui.ButtonSet.OK);
}

// ===============================================================
// 3) PROCEDIMENTO PARA LIMPAR CAMPOS
// ===============================================================
function limparCampos(id) {
  if (arguments.length === 0) {
    return acionarLimparCampos();
  }

  const planilha = obterPlanilha(false);
  corrigirLinhasDesalinhadas(planilha);
  const aba = abaPedidos(false, planilha);
  const linha = acharLinha(aba, id);
  if (!linha) return { ok: false, erro: 'pedido não encontrado' };

  const antes = aba.getRange(linha, 1, 1, 6).getValues()[0];
  aba.getRange(linha, 2, 1, 5).clearContent();
  aba.getRange(linha, 7)
    .setNumberFormat('dd/mm/yyyy hh:mm:ss')
    .setValue(new Date());

  registrarLog(
    'Limpar Campos',
    antes[0],
    antes[1],
    antes[4],
    antes[5],
    'Campos do pedido #' + antes[0] + ' esvaziados via procedimento',
    'planilha',
    planilha,
    antes[2] !== '-' ? antes[2] : '',
    antes[3]
  );

  return { ok: true, id: Number(id) };
}

function acionarLimparCampos() {
  const ui = SpreadsheetApp.getUi();
  const planilha = obterPlanilha(false);
  corrigirLinhasDesalinhadas(planilha);
  const aba = abaPedidos(false, planilha);
  const valoresPainel = aba.getRange(INTERVALO_PAINEL).getValues().map(function(l) {
    return String(l[0] === undefined || l[0] === null ? '' : l[0]).trim();
  });
  const idPainel = valoresPainel[0];
  const temDadosFormulario = valoresPainel.slice(1).some(Boolean);

  if (temDadosFormulario) {
    aba.getRange(INTERVALO_PAINEL).clearContent();
    registrarLog('Limpar Campos', '-', '-', '-', '', 'Campos do painel (J2:J7) limpos via menu/botão', 'planilha', planilha);
    mostrarToast('Campos do painel (J2:J7) limpos com sucesso!', '🧹 Limpar Campos', planilha);
    return;
  }

  if (idPainel) {
    const escolha = ui.alert(
      '🧹 Limpar Campos',
      'Deseja esvaziar os campos do pedido #' + idPainel + '?\n\n' +
      'Clique em "Sim" para limpar os campos do pedido #' + idPainel + ', ou "Não" para limpar apenas a célula J2 do painel.',
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (escolha === ui.Button.CANCEL) return;
    if (escolha === ui.Button.NO) {
      aba.getRange(INTERVALO_PAINEL).clearContent();
      mostrarToast('Campo ID do painel limpo.', '🧹 Limpar Campos', planilha);
      return;
    }

    const res = limparCampos(Number(idPainel));
    aba.getRange(INTERVALO_PAINEL).clearContent();
    if (!res.ok) {
      ui.alert('Erro', 'Pedido #' + idPainel + ' não encontrado.', ui.ButtonSet.OK);
      return;
    }
    mostrarToast('Campos do pedido #' + res.id + ' limpos!', '🧹 Limpar Campos', planilha);
    ui.alert('Concluído', 'Os campos do pedido #' + res.id + ' foram limpos (ID mantido).', ui.ButtonSet.OK);
    return;
  }

  const resp = ui.alert(
    '🧹 Limpar Campos',
    'Os campos de entrada do painel (J2:J7) já estão limpos.\nDeseja limpar os campos de um pedido específico informando o ID?',
    ui.ButtonSet.YES_NO
  );
  if (resp === ui.Button.YES) {
    limparCamposPedido();
  }
}

function limparPainel() {
  const planilha = obterPlanilha(false);
  corrigirLinhasDesalinhadas(planilha);
  const aba = abaPedidos(false, planilha);
  aba.getRange(INTERVALO_PAINEL).clearContent();
  registrarLog('Limpar Campos', '-', '-', '-', '', 'Campos do painel (J2:J7) limpos', 'planilha', planilha);
  mostrarToast('Campos do painel (J2:J7) limpos!', '🧹 Limpar Painel', planilha);
}

function limparCamposPedido() {
  const ui = SpreadsheetApp.getUi();
  const planilha = obterPlanilha(false);
  corrigirLinhasDesalinhadas(planilha);
  const aba = abaPedidos(false, planilha);

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
    'Esvaziar os campos do pedido #' + idTexto + ' (mantendo o ID na linha)?',
    ui.ButtonSet.YES_NO
  );
  if (confirma !== ui.Button.YES) return;

  const resultado = limparCampos(Number(idTexto));
  aba.getRange(INTERVALO_PAINEL).clearContent();
  if (resultado.ok) {
    mostrarToast('Campos do pedido #' + resultado.id + ' esvaziados!', '🗑️ Limpar Campos', planilha);
    ui.alert('Sucesso', 'Os campos do pedido #' + resultado.id + ' foram limpos.', ui.ButtonSet.OK);
  }
}

function carregarPedidoNoPainel() {
  const ui = SpreadsheetApp.getUi();
  const planilha = obterPlanilha(false);
  corrigirLinhasDesalinhadas(planilha);
  const aba = abaPedidos(false, planilha);

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

  const dados = aba.getRange(linha, 1, 1, 6).getValues()[0];
  aba.getRange(INTERVALO_PAINEL).setValues([
    [dados[0]],
    [dados[1]],
    [dados[2] === '-' ? '' : dados[2]],
    [dados[3] === '-' ? '' : dados[3]],
    [dados[4]],
    [dados[5]]
  ]);
  mostrarToast('Pedido #' + dados[0] + ' carregado em J2:J7. Altere e clique em Atualizar!', '🔍 Painel', planilha);
}

function configurarPlanilha() {
  const planilha = obterPlanilha(false);
  corrigirLinhasDesalinhadas(planilha);
  instalarGatilhoAutomatico(planilha);

  const aba = abaPedidos(false, planilha);
  aba.getRange('A1:G1').setValues([CABECALHO_DADOS]);
  aba.getRange('I1:J7').setValues([
    ['PAINEL DE OPERAÇÕES', ''],
    ['ID (atualizar/limpar)', ''],
    ['NOME DO CLIENTE', ''],
    ['E-MAIL (GMAIL)', ''],
    ['ENDEREÇO', ''],
    ['LIVRO', ''],
    ['VALOR (R$)', '']
  ]);
  aba.getRange('I1:J1').merge();
  aba.getRange('J3:J6').setNumberFormat('@');
  aba.getRange('J7').setNumberFormat('R$ #,##0.00');
  aba.getRange('I9:J12').setValues([
    ['BOTÕES DE EXCLUSÃO', ''],
    ['Apagar Linha Selecionada', false],
    ['Apagar Última Linha', false],
    ['Apagar Todos os Dados', false]
  ]);
  aba.getRange('I9:J9').merge();
  aba.getRange('J10:J12').insertCheckboxes();

  const log = abaLog(false, planilha);
  log.getRange('A1:I1').setValues([CABECALHO_LOG]);

  criarBotoesNaPlanilha();

  if (arguments.length === 0) {
    mostrarToast('Estrutura e botões configurados!', '🛠️ Configuração', planilha);
  }
}

// ===============================================================
// ENTRADA PELA INTERNET (chamada pelo site ao finalizar compra)
// ===============================================================
function doGet() {
  return resposta({ ok: true, versao: 2, status: 'Web App ativo e funcionando!' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const dados = JSON.parse(e.postData.contents);
    const meta = extrairMetadadosLivro(dados.livro);
    return resposta(
      registrarOperacao(
        dados.nome,
        dados.livroLimpo || meta.livro || dados.livro,
        dados.valor,
        'site',
        dados.email || meta.email,
        dados.endereco || meta.endereco,
        dados.pagamento || meta.pagamento
      )
    );
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
