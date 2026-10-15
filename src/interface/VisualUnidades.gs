/**

 * Visual das planilhas das unidades (mesmo modelo da Mãe).

 * Aplicado sob demanda — NÃO roda no sync Enviado → Mãe.

 */



/**

 * Menu: aplica o visual em TODAS as planilhas de unidade.

 */

function menuAplicarVisualTodasUnidades() {

  var ui = SpreadsheetApp.getUi();

  var resp = ui.alert(

    'Visual nas unidades',

    'Aplica o modelo da Mãe em TODAS as planilhas Pedido:\n\n' +

      '• Faixa “PLANILHA DE PEDIDOS — MÊS XX” no topo\n' +

      '• Cabeçalho VERMELHO\n' +

      '• Coluna Produto AMARELA\n' +

      '• 🛒 Quantidade (âmbar) e Status com destaque\n' +

      '• Fundo ESCURO\n' +

      '• Nome do produto completo (coluna larga)\n\n' +

      'Pedidos preenchidos são mantidos.\nContinuar?',

    ui.ButtonSet.YES_NO

  );

  if (resp !== ui.Button.YES) return;



  try {

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    ss.toast('Aplicando visual… isso pode levar 1–2 min', 'Visual', 120);

    var r = aplicarVisualTodasUnidades_();

    SpreadsheetApp.flush();

    ui.alert(

      r.ok > 0 ? 'Visual aplicado' : 'Nenhuma unidade atualizada',

      'OK: ' +

        r.ok +

        '\nFalhas: ' +

        r.falhas.length +

        (r.detalhes.length ? '\n\n' + r.detalhes.slice(0, 14).join('\n') : '') +

        (r.falhas.length ? '\n\nFalhas:\n- ' + r.falhas.slice(0, 8).join('\n- ') : '') +

        '\n\nRecarregue a planilha da unidade (F5) para ver as cores.',

      ui.ButtonSet.OK

    );

  } catch (e) {

    ui.alert('Erro', e.message + '\n\n' + (e.stack || ''), ui.ButtonSet.OK);

  }

}



/**

 * @return {{ok:number, falhas:string[], detalhes:string[]}}

 */

function aplicarVisualTodasUnidades_() {

  var nomes =

    DADOS_INICIAIS && DADOS_INICIAIS.unidades

      ? DADOS_INICIAIS.unidades.slice()

      : listarUnidadesAtivas().map(function (u) {

          return u.nome;

        });



  var ok = 0;

  var falhas = [];

  var detalhes = [];



  for (var i = 0; i < nomes.length; i++) {

    var nome = nomes[i];

    try {

      SpreadsheetApp.getActiveSpreadsheet().toast(

        'Unidade ' + (i + 1) + '/' + nomes.length + ': ' + nome,

        'Visual',

        10

      );



      var id = resolverIdUnidadePeloNome_(nome);

      if (!id) throw new Error('não achei no Drive: Pedido - ' + nome);



      var ssFilha = SpreadsheetApp.openById(id);

      var aba = ssFilha.getSheetByName('Pedido');

      if (!aba) {

        aba = obterAbaPedidoFilha_(ssFilha);

      }



      try {

        aplicarDropdownPedido(ssFilha);

      } catch (eDrop) {}



      aplicarVisualPedidoUnidade_(aba);

      try {

        aplicarDropdownPedido(ssFilha, { somenteLinhasComDados: true });

      } catch (eDrop2) {}



      var headerLinha = obterLinhaCabecalhoPedido_(aba);

      var bg = String(aba.getRange(headerLinha, 1).getBackground() || '').toLowerCase();

      detalhes.push(

        nome +

          ' → ' +

          ssFilha.getName() +

          ' | cab=' +

          headerLinha +

          ' A' +

          headerLinha +

          '=' +

          bg +

          (bg.indexOf('ff0000') !== -1 || bg === CONFIG.TEMA.HEADER_BG.toLowerCase()

            ? ' ✓'

            : ' (cor diferente)')

      );



      atualizarIdUnidadeNaAba(nome, id);

      ok++;

      SpreadsheetApp.flush();

    } catch (e) {

      falhas.push(nome + ': ' + e.message);

    }

  }



  registrarLog('VISUAL_UNIDADES', ok + ' ok / ' + falhas.length + ' falhas');

  return { ok: ok, falhas: falhas, detalhes: detalhes };

}



/**

 * Aplica visual na aba Pedido (valores intactos).

 * @param {Sheet} aba

 */

function aplicarVisualPedidoUnidade_(aba) {

  if (!aba) throw new Error('aba Pedido vazia');

  var t = CONFIG.TEMA;



  if (planilhaPedidoUsaLayoutLegadoComData_(aba)) {

    migrarPlanilhaPedidoRemoverColunaData_(aba);

  } else if (planilhaPedidoUsaLayoutComEstoque_(aba)) {

    migrarPlanilhaPedidoRemoverColunaEstoque_(aba);

  }



  var headerLinha = obterLinhaCabecalhoPedido_(aba);



  try {

    aplicarMesReferenciaPedidoUnidade_(aba);

  } catch (eMes) {}



  aplicarVisualPedidoEmLinha_(aba, headerLinha, t);

}



/**

 * @param {Sheet} aba

 * @param {number} headerLinha 1-based

 * @param {Object} t tema

 */

function aplicarVisualPedidoEmLinha_(aba, headerLinha, t) {

  var nCols = obterNumColsPedido_();

  var colLetraFim = String.fromCharCode(64 + nCols);

  var lastRow = Math.max(aba.getLastRow(), headerLinha + 1);

  var endRow = Math.max(lastRow, headerLinha + 40);



  try {

    garantirCabecalhoPedidoUnidade_(aba);

  } catch (eHdr) {}



  try {

    var filtro = aba.getFilter();

    if (filtro) filtro.remove();

  } catch (eF) {}



  try {

    aba.setTabColor(t.TAB);

  } catch (eTab) {}



  var fundoEnd = Math.min(endRow + 30, 250);

  aba.getRange('A1:Z' + fundoEnd).setBackground(t.SHEET_BG);



  var headerA1 = 'A' + headerLinha + ':' + colLetraFim + headerLinha;

  var headerRange = aba.getRange(headerA1);

  headerRange.setBackground(t.HEADER_BG);

  headerRange.setFontColor(t.HEADER_FG);

  headerRange.setFontWeight('bold');

  headerRange.setFontSize(10);

  headerRange.setFontFamily('Arial');

  headerRange.setHorizontalAlignment('center');

  headerRange.setVerticalAlignment('middle');

  headerRange.setWrap(true);

  aba.setRowHeight(headerLinha, 48);

  aba.setFrozenRows(CONFIG.PEDIDO.HEADER_LINHA || headerLinha);



  var dataStart = headerLinha + 1;

  var dataA1 = 'A' + dataStart + ':' + colLetraFim + endRow;

  var dataRange = aba.getRange(dataA1);



  dataRange.setBackground(t.ROW_A);

  dataRange.setFontColor(t.TEXT_MUTED);

  dataRange.setFontSize(10);

  dataRange.setFontFamily('Arial');

  dataRange.setVerticalAlignment('middle');

  dataRange.setHorizontalAlignment('center');



  var prodA1 = 'A' + dataStart + ':A' + endRow;

  var prodRange = aba.getRange(prodA1);

  prodRange.setBackground(t.COL_DESTAQUE_BG);

  prodRange.setFontColor(t.COL_DESTAQUE_FG);

  prodRange.setFontWeight('bold');

  prodRange.setHorizontalAlignment('left');

  prodRange.setWrap(false);



  aba.setColumnWidth(CONFIG.PEDIDO.PRODUTO, 560);

  aba.setColumnWidth(CONFIG.PEDIDO.QUANTIDADE, 168);

  aba.setColumnWidth(CONFIG.PEDIDO.OBSERVACAO, 160);

  aba.setColumnWidth(CONFIG.PEDIDO.STATUS, 140);



  try {

    ajustarLarguraNomeCompleto_(aba, CONFIG.PEDIDO.PRODUTO, Math.max(1, lastRow - headerLinha));

    if (aba.getColumnWidth(CONFIG.PEDIDO.PRODUTO) < 480) {

      aba.setColumnWidth(CONFIG.PEDIDO.PRODUTO, 560);

    }

  } catch (eW) {

    aba.setColumnWidth(CONFIG.PEDIDO.PRODUTO, 560);

  }



  dataRange.setBorder(

    true,

    true,

    true,

    true,

    true,

    true,

    t.BORDER,

    SpreadsheetApp.BorderStyle.SOLID

  );



  headerRange.setBorder(

    true,

    true,

    true,

    true,

    true,

    true,

    t.BORDER,

    SpreadsheetApp.BorderStyle.SOLID

  );



  try {

    aba.getRange('A' + headerLinha + ':' + colLetraFim + endRow).createFilter();

  } catch (eFiltro) {}



  try {

    aplicarFormatacaoCondicionalStatus_(

      aba,

      dataStart,

      endRow - dataStart + 1,

      CONFIG.PEDIDO.STATUS

    );

  } catch (eCond) {}



  reforcarCabecalhoPedidoGeral_(aba, headerLinha, t);

  aplicarVisualColunaQuantidade_(aba, headerLinha, dataStart, endRow, t);

  aplicarVisualColunaObservacao_(aba, dataStart, endRow, t);

  aplicarVisualColunaStatus_(aba, headerLinha, dataStart, endRow, t);

  aplicarVisualCategoriasPedido_(aba, dataStart, endRow, t);



  try {

    aplicarDropdownPedido(aba.getParent(), { somenteLinhasComDados: true, linhaFim: endRow });

  } catch (eDrop) {}

}



/**

 * Destaque visual da coluna Quantidade (âmbar).

 * @param {Sheet} aba

 * @param {number} headerLinha

 * @param {number} dataStart

 * @param {number} endRow

 * @param {Object} t

 */

function aplicarVisualColunaQuantidade_(aba, headerLinha, dataStart, endRow, t) {

  if (!aba || endRow < dataStart) return;

  var colQtd = CONFIG.PEDIDO.QUANTIDADE;

  var nLinhas = endRow - dataStart + 1;



  aba

    .getRange(headerLinha, colQtd)

    .setBackground(t.COL_QTD_HEADER_BG)

    .setFontColor('#FFFFFF')

    .setFontWeight('bold')

    .setHorizontalAlignment('center')

    .setWrap(true);



  aba

    .getRange(dataStart, colQtd, nLinhas, 1)

    .setBackground(t.COL_QTD_BG)

    .setFontColor(t.COL_QTD_FG)

    .setFontWeight('bold')

    .setHorizontalAlignment('center')

    .setVerticalAlignment('middle');

}



/**

 * Destaque visual da coluna Status (Rascunho / Enviado), no mesmo padrão das demais.

 * @param {Sheet} aba

 * @param {number} headerLinha

 * @param {number} dataStart

 * @param {number} endRow

 * @param {Object} t

 */

function aplicarVisualColunaStatus_(aba, headerLinha, dataStart, endRow, t) {

  if (!aba || endRow < dataStart) return;

  var colStatus = CONFIG.PEDIDO.STATUS;

  var nLinhas = endRow - dataStart + 1;



  aba

    .getRange(headerLinha, colStatus)

    .setBackground(t.COL_STATUS_HEADER_BG)

    .setFontColor('#FFFFFF')

    .setFontWeight('bold')

    .setHorizontalAlignment('center')

    .setWrap(true);



  aba

    .getRange(dataStart, colStatus, nLinhas, 1)

    .setBackground(t.COL_STATUS_RASCUNHO_BG)

    .setFontColor(t.COL_STATUS_RASCUNHO_FG)

    .setFontWeight('bold')

    .setHorizontalAlignment('center')

    .setVerticalAlignment('middle');

}



/**

 * Coluna Observação: texto livre, sem dropdown (só Status envia o pedido).

 * @param {Sheet} aba

 * @param {number} dataStart

 * @param {number} endRow

 * @param {Object} t

 */

function aplicarVisualColunaObservacao_(aba, dataStart, endRow, t) {

  if (!aba || endRow < dataStart) return;

  var colObs = CONFIG.PEDIDO.OBSERVACAO;

  var nLinhas = endRow - dataStart + 1;

  aba

    .getRange(dataStart, colObs, nLinhas, 1)

    .clearDataValidations()

    .setBackground(t.ROW_A)

    .setFontColor(t.TEXT_MUTED)

    .setFontWeight('normal')

    .setHorizontalAlignment('left')

    .setVerticalAlignment('middle')

    .setWrap(true);

}



/**

 * Mantém vermelho no cabeçalho das colunas que não são quantidade/status.

 * @param {Sheet} aba

 * @param {number} headerLinha

 * @param {Object} t

 */

function reforcarCabecalhoPedidoGeral_(aba, headerLinha, t) {

  var cols = [CONFIG.PEDIDO.PRODUTO, CONFIG.PEDIDO.OBSERVACAO];

  for (var i = 0; i < cols.length; i++) {

    aba

      .getRange(headerLinha, cols[i])

      .setBackground(t.HEADER_BG)

      .setFontColor(t.HEADER_FG)

      .setFontWeight('bold');

  }

}



/**

 * Reaplica só as cores das linhas de dados (rápido).

 * @param {Sheet} aba

 * @param {number} headerLinha 1-based (linha do cabeçalho vermelho)

 * @param {number=} linhaFimCatalogo última linha gravada pelo catálogo

 */

function aplicarVisualLinhasPedidoAposCatalogo_(aba, headerLinha, linhaFimCatalogo) {

  if (!aba) return;

  var t = CONFIG.TEMA;

  var nCols = obterNumColsPedido_();

  var colLetraFim = String.fromCharCode(64 + nCols);

  var dataStart = headerLinha + 1;

  var lastRow = linhaFimCatalogo || Math.max(aba.getLastRow(), dataStart);

  if (lastRow < dataStart) return;



  try {

    garantirCabecalhoPedidoUnidade_(aba);

  } catch (eHdr) {}



  var dataA1 = 'A' + dataStart + ':' + colLetraFim + lastRow;

  var dataRange = aba.getRange(dataA1);

  dataRange.setBackground(t.ROW_A);

  dataRange.setFontColor(t.TEXT_MUTED);

  dataRange.setFontSize(10);

  dataRange.setFontFamily('Arial');

  dataRange.setVerticalAlignment('middle');

  dataRange.setHorizontalAlignment('center');



  aba

    .getRange('A' + dataStart + ':A' + lastRow)

    .setBackground(t.COL_DESTAQUE_BG)

    .setFontColor(t.COL_DESTAQUE_FG)

    .setFontWeight('bold')

    .setHorizontalAlignment('left')

    .setWrap(false);



  aplicarVisualColunaQuantidade_(aba, headerLinha, dataStart, lastRow, t);

  aplicarVisualColunaObservacao_(aba, dataStart, lastRow, t);

  aplicarVisualColunaStatus_(aba, headerLinha, dataStart, lastRow, t);

  reforcarCabecalhoPedidoGeral_(aba, headerLinha, t);

  aplicarVisualCategoriasPedido_(aba, dataStart, lastRow, t);



  try {

    aplicarDropdownPedido(aba.getParent(), { somenteLinhasComDados: true, linhaFim: lastRow });

  } catch (eDrop) {}

}



/**

 * Aplica faixas verdes nas linhas de categoria da aba Pedido.

 */

function aplicarVisualCategoriasPedido_(aba, linhaInicio, linhaFim, t) {

  if (linhaFim < linhaInicio) return;

  var nCols = obterNumColsPedido_();

  var colLetraFim = String.fromCharCode(64 + nCols);

  var valores = aba

    .getRange(linhaInicio, 1, linhaFim - linhaInicio + 1, nCols)

    .getDisplayValues();



  var faixas = [];

  var linhasCategoria = [];

  for (var i = 0; i < valores.length; i++) {

    var produto = String(valores[i][CONFIG.PEDIDO.PRODUTO - 1] || '').trim();

    if (!ehLinhaCategoriaPedido_(produto)) continue;



    var linha = linhaInicio + i;

    faixas.push('A' + linha + ':' + colLetraFim + linha);

    linhasCategoria.push(linha);

    try {
      aba
        .getRange(linha, CONFIG.PEDIDO.STATUS)
        .clearContent()
        .clearDataValidations();
    } catch (eClr) {}
  }



  if (faixas.length) {

    var lista = aba.getRangeList(faixas);

    lista

      .setBackground('#0F766E')

      .setFontColor('#FFFFFF')

      .setFontWeight('bold')

      .setFontSize(10)

      .setHorizontalAlignment('left')

      .setVerticalAlignment('middle');

    for (var r = 0; r < linhasCategoria.length; r++) {

      aba.setRowHeight(linhasCategoria[r], 26);

      aba
        .getRange(linhasCategoria[r], CONFIG.PEDIDO.STATUS)
        .setBackground('#0F766E')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

    }

  }

}

/**
 * Reaplica faixas de categoria (inclui coluna Status) após validações.
 * @param {Sheet} aba
 */
function corrigirVisualFaixasCategoria_(aba) {
  if (!aba) return;
  var inicio = CONFIG.PEDIDO.LINHA_DADOS || 3;
  var fim = Math.max(aba.getLastRow(), inicio);
  if (fim < inicio) return;
  aplicarVisualCategoriasPedido_(aba, inicio, fim, CONFIG.TEMA);
}



/**

 * @param {Sheet} aba

 * @param {number} dataStart

 * @param {number} nDataRows

 * @param {number} colStatus 1-based

 */

function aplicarFormatacaoCondicionalStatus_(aba, dataStart, nDataRows, colStatus) {

  var end = dataStart + nDataRows - 1;

  var produtos = aba
    .getRange(dataStart, CONFIG.PEDIDO.PRODUTO, nDataRows, 1)
    .getDisplayValues();

  var rangesStatus = [];
  var blocoInicio = -1;

  for (var i = 0; i < produtos.length; i++) {
    var linha = dataStart + i;
    var produto = String(produtos[i][0] || '').trim();
    var pular = ehLinhaCategoriaPedido_(produto) || !produto;

    if (pular) {
      if (blocoInicio >= 0) {
        rangesStatus.push(aba.getRange(blocoInicio, colStatus, linha - blocoInicio, 1));
        blocoInicio = -1;
      }
      continue;
    }
    if (blocoInicio < 0) blocoInicio = linha;
  }

  if (blocoInicio >= 0) {
    rangesStatus.push(aba.getRange(blocoInicio, colStatus, end - blocoInicio + 1, 1));
  }

  if (rangesStatus.length === 0) return;



  var t = CONFIG.TEMA;

  var rEnviado = SpreadsheetApp.newConditionalFormatRule()

    .whenTextEqualTo(CONFIG.VALORES.STATUS_ENVIADO)

    .setBackground(t.COL_STATUS_ENVIADO_BG)

    .setFontColor(t.COL_STATUS_ENVIADO_FG)

    .setRanges(rangesStatus)

    .build();



  var rRascunho = SpreadsheetApp.newConditionalFormatRule()

    .whenTextEqualTo(CONFIG.VALORES.STATUS_RASCUNHO)

    .setBackground(t.COL_STATUS_RASCUNHO_BG)

    .setFontColor(t.COL_STATUS_RASCUNHO_FG)

    .setRanges(rangesStatus)

    .build();



  var rCons = SpreadsheetApp.newConditionalFormatRule()

    .whenTextEqualTo(CONFIG.VALORES.STATUS_CONSOLIDADO)

    .setBackground(t.COL_STATUS_CONSOLIDADO_BG)

    .setFontColor(t.COL_STATUS_CONSOLIDADO_FG)

    .setRanges(rangesStatus)

    .build();



  var existentes = aba.getConditionalFormatRules() || [];

  var mantidas = [];

  for (var i = 0; i < existentes.length; i++) {

    var rngs = existentes[i].getRanges();

    var soStatus = false;

    for (var j = 0; j < rngs.length; j++) {

      if (rngs[j].getColumn() === colStatus) soStatus = true;

    }

    if (!soStatus) mantidas.push(existentes[i]);

  }

  aba.setConditionalFormatRules(mantidas.concat([rEnviado, rRascunho, rCons]));

}

