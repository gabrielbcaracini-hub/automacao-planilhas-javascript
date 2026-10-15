/**
 * Leitura flexível + reparo da planilha da unidade (sem Tabela/colunas tipadas).
 */

/**
 * Lê pedidos Enviado detectando cabeçalho automaticamente.
 * @param {Sheet} aba
 * @return {{enviados: Array.<Object>, diagnostico: string, todas: Array.<Object>}}
 */
function lerPedidosEnviadosFlexivel(aba) {
  var values = [];
  var displays = [];
  try {
    var rng = aba.getDataRange();
    values = rng.getValues();
    displays = rng.getDisplayValues();
  } catch (e) {
    var rng2 = aba.getRange(1, 1, Math.max(aba.getLastRow(), 1), Math.max(aba.getLastColumn(), 5));
    values = rng2.getValues();
    try {
      displays = rng2.getDisplayValues();
    } catch (e2) {
      displays = values;
    }
  }
  if (!displays || displays.length !== values.length) displays = values;

  var diag = [];
  diag.push('lastRow=' + aba.getLastRow() + ' lastCol=' + aba.getLastColumn());
  diag.push('linhas lidas=' + values.length);

  var tituloAtual = normalizarCabecalhoPedido_(
    displays[0] && displays[0][0]
  );
  var mapa = detectarColunasPedido_(displays);
  if (mapa.colStatus < 0 || mapa.colQtd < 0 || mapa.colProduto < 0) {
    mapa = detectarColunasPedido_(values);
  }
  diag.push(mapa.diag);
  if (tituloAtual) {
    diag.push('titulo=' + tituloAtual);
  }

  var todas = [];
  var enviados = [];
  var amostra = [];

  for (var i = mapa.headerRow + 1; i < values.length; i++) {
    var row = values[i];
    var rowD = displays[i] || row;
    var produto = String(rowD[mapa.colProduto] || row[mapa.colProduto] || '').trim();
    if (!produto) continue;
    // Faixa visual de categoria da aba Pedido: nunca é um item de pedido.
    if (ehLinhaCategoriaPedido_(produto)) continue;

    var qtd = row[mapa.colQtd];
    var qtdDisplay = rowD[mapa.colQtd];
    var qtdTexto = qtd === null || qtd === undefined || qtd === '' ? '' : String(qtd).trim();
    if (!qtdTexto && qtdDisplay !== null && qtdDisplay !== undefined) {
      qtdTexto = String(qtdDisplay).trim();
    }
    var status = String(rowD[mapa.colStatus] || row[mapa.colStatus] || '').trim();
    var item = {
      data: mapa.colData >= 0 ? row[mapa.colData] || '' : '',
      produto: produto,
      estoqueAtual: mapa.colEstoque >= 0 ? valorCelulaPreservandoZero_(row[mapa.colEstoque]) : '',
      quantidade: qtdTexto,
      observacao: mapa.colObs >= 0 ? String(row[mapa.colObs] || '').trim() : '',
      status: status || CONFIG.VALORES.STATUS_RASCUNHO
    };
    todas.push(item);

    var st = normalizarTexto(status)
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z]/g, '');
    // normalizarTexto já lowercases; remove acentos de forma simples
    st = status
      .toLowerCase()
      .replace(/á|à|ã|â/g, 'a')
      .replace(/é|ê/g, 'e')
      .replace(/í/g, 'i')
      .replace(/ó|ô|õ/g, 'o')
      .replace(/ú/g, 'u')
      .replace(/ç/g, 'c')
      .replace(/[^a-z]/g, '');

    if (amostra.length < 8) {
      amostra.push(
        produto.substring(0, 35) +
          ' | qtd=[' +
          qtdTexto +
          '] | status=[' +
          status +
          '] → ' +
          st
      );
    }

    if (!qtdTexto) continue;
    if (
      st.indexOf('enviado') === -1 &&
      st.indexOf('consolidado') === -1
    ) continue;
    enviados.push(item);
  }

  diag.push('Total produtos na aba: ' + todas.length);
  diag.push('Enviados: ' + enviados.length);
  diag.push('Amostra:\n' + (amostra.join('\n') || '(vazio)'));

  return { enviados: enviados, diagnostico: diag.join('\n'), todas: todas };
}

/**
 * @param {Array.<Array.<*>>} values
 * @return {Object}
 */
function detectarColunasPedido_(values) {
  var headerRow = -1;
  var colProduto = -1;
  var colQtd = -1;
  var colStatus = -1;
  var colData = -1;
  var colEstoque = -1;
  var colObs = -1;

  for (var r = 0; r < Math.min(20, values.length); r++) {
    var linhaMes = String(values[r][0] || values[r][1] || '')
      .toLowerCase()
      .replace(/á|à|ã|â/g, 'a')
      .replace(/é|ê/g, 'e')
      .replace(/í/g, 'i')
      .replace(/ó|ô|õ/g, 'o')
      .replace(/ú/g, 'u')
      .replace(/ç/g, 'c');
    if (linhaMes.indexOf('planilha de pedidos') !== -1) continue;

    for (var c = 0; c < values[r].length; c++) {
      var h = String(values[r][c] || '')
        .toLowerCase()
        .replace(/á|à|ã|â/g, 'a')
        .replace(/é|ê/g, 'e')
        .replace(/í/g, 'i')
        .replace(/ó|ô|õ/g, 'o')
        .replace(/ú/g, 'u')
        .replace(/ç/g, 'c')
        .replace(/[^a-z]/g, '');
      if (h === 'produto') {
        headerRow = r;
        colProduto = c;
      }
      if (h.indexOf('quantidad') === 0 || h === 'qtd') {
        colQtd = c;
        if (headerRow < 0) headerRow = r;
      }
      if (h === 'status' || h.indexOf('status') === 0) {
        colStatus = c;
        if (headerRow < 0) headerRow = r;
      }
      if (h === 'data') {
        colData = c;
        if (headerRow < 0) headerRow = r;
      }
      if (h.indexOf('estoque') === 0) {
        colEstoque = c;
        if (headerRow < 0) headerRow = r;
      }
      if (h.indexOf('observ') === 0) {
        colObs = c;
        if (headerRow < 0) headerRow = r;
      }
    }
    if (headerRow >= 0 && colProduto >= 0 && colQtd >= 0 && colStatus >= 0) break;
  }

  if (headerRow < 0 || colProduto < 0 || colQtd < 0 || colStatus < 0) {
    var linhaPadrao = (CONFIG.PEDIDO.HEADER_LINHA || 2) - 1;
    var layoutComData = false;
    for (var rf = 0; rf < Math.min(20, values.length); rf++) {
      var primeira = normalizarCabecalhoPedido_(values[rf][0]);
      if (primeira === 'data' && values[rf].length >= 6) {
        linhaPadrao = rf;
        layoutComData = true;
        break;
      }
    }
    if (layoutComData) {
      return {
        headerRow: linhaPadrao,
        colData: 0,
        colProduto: 1,
        colEstoque: 2,
        colQtd: 3,
        colObs: 4,
        colStatus: 5,
        diag: 'Cabeçalho legado — layout 6 colunas (com Data)'
      };
    }
    return {
      headerRow: linhaPadrao,
      colData: -1,
      colProduto: 0,
      colEstoque: -1,
      colQtd: 1,
      colObs: 2,
      colStatus: 3,
      diag: 'Cabeçalho não detectado — layout padrão 4 colunas (sem Estoque)'
    };
  }

  return {
    headerRow: headerRow,
    colData: colData,
    colProduto: colProduto,
    colEstoque: colEstoque,
    colQtd: colQtd,
    colObs: colObs,
    colStatus: colStatus,
    diag:
      'Cabeçalho linha ' +
      (headerRow + 1) +
      ' | Produto=' +
      (colProduto + 1) +
      ' Qtd=' +
      (colQtd + 1) +
      ' Status=' +
      (colStatus + 1)
  };
}

/**
 * Menu: repara Unidade Centro (recria aba Pedido sem tabela tipada) e puxa Enviados.
 */
function menuRepararUnidadeCentroEPuxar() {
  var ui = SpreadsheetApp.getUi();
  try {
    var id = resolverIdUnidadePeloNome_('UNIDADE CENTRO');
    if (!id) throw new Error('Não achei no Drive: Pedido - UNIDADE CENTRO');

    var ssFilha = SpreadsheetApp.openById(id);
    limparAbasErradasDaFilha(ssFilha);

    var abaAntiga = ssFilha.getSheetByName('Pedido');
    if (!abaAntiga) {
      abaAntiga = ssFilha.getSheets()[0];
    }

    var lido = lerPedidosEnviadosFlexivel(abaAntiga);

    // Recria aba limpa (evita erro de "colunas com tipo" / Tabela_1)
    var abaNova = recriarAbaPedidoLimpa_(ssFilha, abaAntiga, lido.todas);

    garantirAbaCatalogo(ssFilha);
    atualizarIdUnidadeNaAba('UNIDADE CENTRO', id);
    atualizarMaeComPedidosUnidade('UNIDADE CENTRO', lido.enviados);
    SpreadsheetApp.flush();

    ui.alert(
      'Unidade Centro reparado',
      'Arquivo: ' +
        ssFilha.getName() +
        '\nID: ' +
        id +
        '\nAbas: ' +
        ssFilha
          .getSheets()
          .map(function (s) {
            return s.getName();
          })
          .join(', ') +
        '\nAba Pedido nova: ' +
        abaNova.getName() +
        '\n\n' +
        lido.diagnostico +
        '\n\nEnviados gravados na MÃE: ' +
        lido.enviados.length +
        '\n\nAbra na MÃE: Resumo.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Recria a aba Pedido do zero (sem Tabela / tipos de coluna).
 * @param {Spreadsheet} ss
 * @param {Sheet} abaAntiga
 * @param {Array.<Object>} linhas
 * @return {Sheet}
 */
function recriarAbaPedidoLimpa_(ss, abaAntiga, linhas) {
  var nomeTmp = 'Pedido_LIMPO_' + Date.now();
  var nova = ss.insertSheet(nomeTmp);
  var nCols = obterNumColsPedido_();

  prepararLayoutPedidoUnidade_(nova);
  var dataStart = CONFIG.PEDIDO.LINHA_DADOS || 3;

  if (linhas && linhas.length) {
    var valores = linhas.map(function (item) {
      return montarLinhaPedidoUnidade_(
        item.produto,
        item.quantidade || '',
        item.observacao || '',
        item.status || CONFIG.VALORES.STATUS_RASCUNHO
      );
    });
    nova.getRange(dataStart, 1, valores.length, nCols).setValues(valores);

    var regra = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        [
          CONFIG.VALORES.STATUS_RASCUNHO,
          CONFIG.VALORES.STATUS_ENVIADO,
          CONFIG.VALORES.STATUS_CONSOLIDADO
        ],
        true
      )
      .setAllowInvalid(true)
      .build();
    nova.getRange(dataStart, CONFIG.PEDIDO.STATUS, valores.length, 1).setDataValidation(regra);
  }

  // Apaga a aba antiga problemática
  var nomeAntigo = abaAntiga.getName();
  try {
    if (ss.getSheets().length > 1) {
      ss.deleteSheet(abaAntiga);
    }
  } catch (e) {
    // se não der para apagar, renomeia
    try {
      abaAntiga.setName('Pedido_ANTIGO_ERRO');
    } catch (e2) {}
  }

  // Garante nome Pedido
  try {
    var existePedido = ss.getSheetByName('Pedido');
    if (existePedido && existePedido.getSheetId() !== nova.getSheetId()) {
      existePedido.setName('Pedido_OLD_' + Date.now());
    }
  } catch (e3) {}

  nova.setName('Pedido');
  try {
    aplicarVisualPedidoUnidade_(nova);
  } catch (eVis) {}
  return nova;
}

/**
 * @param {string} nomeUnidade
 * @return {string}
 */
function resolverIdUnidadePeloNome_(nomeUnidade) {
  var it = DriveApp.getFilesByName('Pedido - ' + nomeUnidade);
  while (it.hasNext()) {
    var f = it.next();
    if (f.isTrashed()) continue;
    if (f.getName().indexOf('COMPLETO') !== -1) continue;
    return f.getId();
  }
  if (DADOS_INICIAIS && DADOS_INICIAIS.idsPlanilhas && DADOS_INICIAIS.idsPlanilhas[nomeUnidade]) {
    return DADOS_INICIAIS.idsPlanilhas[nomeUnidade];
  }
  return '';
}
