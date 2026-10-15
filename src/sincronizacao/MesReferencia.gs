/**
 * Mês de referência nas planilhas das unidades (substitui coluna Data).
 */

/**
 * @param {Date=} data
 * @return {string}
 */
function obterTextoMesReferenciaPedido_(data) {
  data = data || new Date();
  var mes = Utilities.formatDate(data, Session.getScriptTimeZone(), 'MM');
  return 'PLANILHA DE PEDIDOS — MÊS ' + mes;
}

/**
 * @param {Sheet} aba
 * @param {Date=} data
 */
function aplicarMesReferenciaPedidoUnidade_(aba, data) {
  if (!aba) return;
  var linha = CONFIG.PEDIDO.MES_REFERENCIA_LINHA || 1;
  var nCols = obterNumColsPedido_();
  var texto = obterTextoMesReferenciaPedido_(data);
  var rng = aba.getRange(linha, 1, 1, nCols);
  try {
    rng.breakApart();
  } catch (eBreak) {}
  rng.merge();
  rng
    .setValue(texto)
    .setBackground('#0F766E')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontFamily('Arial')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  aba.setRowHeight(linha, 34);
}

/**
 * @return {number}
 */
function obterNumColsPedido_() {
  return CONFIG.PEDIDO.NUM_COLS || CONFIG.PEDIDO.HEADER.length;
}

/**
 * Linha do cabeçalho (Produto, Quantidade…) na aba Pedido.
 * @param {Sheet=} aba
 * @return {number} 1-based
 */
function obterLinhaCabecalhoPedido_(aba) {
  var padrao = CONFIG.PEDIDO.HEADER_LINHA || 2;
  if (!aba) return padrao;
  try {
    var lastCol = Math.max(aba.getLastColumn(), obterNumColsPedido_(), 6);
    var vals = aba
      .getRange(
        1,
        1,
        Math.min(15, Math.max(aba.getLastRow(), 1)),
        lastCol
      )
      .getValues();
    var mapa = detectarColunasPedido_(vals);
    if (mapa.headerRow >= 0) {
      return mapa.headerRow + 1;
    }
  } catch (e) {}
  return padrao;
}

/**
 * Planilha ainda com coluna Data ou sem faixa de mês no topo.
 * @param {Sheet} aba
 * @return {boolean}
 */
function planilhaPedidoUsaLayoutLegadoComData_(aba) {
  if (!aba) return false;
  var lastRow = Math.max(aba.getLastRow(), 1);
  var lastCol = Math.max(aba.getLastColumn(), obterNumColsPedido_(), 6);
  var vals = aba.getRange(1, 1, Math.min(15, lastRow), lastCol).getValues();
  var linha1Norm = normalizarTextoPedido_(String(vals[0].join(' ')));
  var temFaixaMes = linha1Norm.indexOf('planilha de pedidos') !== -1;
  var mapa = detectarColunasPedido_(vals);
  if (mapa.colData >= 0) return true;
  if (mapa.colProduto !== 0) return true;
  if (!temFaixaMes) return true;
  var headerEsperado = CONFIG.PEDIDO.HEADER_LINHA || 2;
  if (mapa.headerRow + 1 !== headerEsperado) return true;
  return false;
}

/**
 * @param {string} s
 * @return {string}
 */
function normalizarTextoPedido_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/á|à|ã|â/g, 'a')
    .replace(/é|ê/g, 'e')
    .replace(/í/g, 'i')
    .replace(/ó|ô|õ/g, 'o')
    .replace(/ú/g, 'u')
    .replace(/ç/g, 'c');
}

/**
 * Remove coluna Data, insere faixa de mês e regrava dados sem estoque.
 * @param {Sheet} aba
 * @param {Date=} dataMes
 */
function migrarPlanilhaPedidoRemoverColunaData_(aba, dataMes) {
  if (!aba) return;

  var eraLegado = planilhaPedidoUsaLayoutLegadoComData_(aba);
  if (!eraLegado) {
    if (planilhaPedidoUsaLayoutComEstoque_(aba)) {
      migrarPlanilhaPedidoRemoverColunaEstoque_(aba, dataMes);
      return;
    }
    prepararLayoutPedidoUnidade_(aba, dataMes);
    return;
  }

  var lastRow = Math.max(aba.getLastRow(), 1);
  var lastCol = Math.max(aba.getLastColumn(), 6);
  var vals = aba.getRange(1, 1, lastRow, lastCol).getValues();
  var mapa = detectarColunasPedido_(vals);
  var linhasPedido = [];

  for (var i = mapa.headerRow + 1; i < vals.length; i++) {
    var row = vals[i];
    var produto = String(row[mapa.colProduto] || '').trim();
    if (!produto) continue;
    if (ehLinhaCategoriaPedido_(produto)) {
      linhasPedido.push(montarLinhaCategoriaPedido_(produto.replace(/^[▸▶]\s*/, '')));
      continue;
    }
    linhasPedido.push(
      montarLinhaPedidoUnidade_(
        produto,
        mapa.colQtd >= 0
          ? row[mapa.colQtd] === null || row[mapa.colQtd] === undefined
            ? ''
            : String(row[mapa.colQtd]).trim()
          : '',
        mapa.colObs >= 0 ? String(row[mapa.colObs] || '').trim() : '',
        mapa.colStatus >= 0
          ? String(row[mapa.colStatus] || '').trim() || CONFIG.VALORES.STATUS_RASCUNHO
          : CONFIG.VALORES.STATUS_RASCUNHO
      )
    );
  }

  try {
    var filtro = aba.getFilter();
    if (filtro) filtro.remove();
  } catch (eFiltro) {}

  var linha1Norm = normalizarTextoPedido_(String(vals[0].join(' ')));
  if (linha1Norm.indexOf('planilha de pedidos') === -1 && mapa.headerRow === 0) {
    aba.insertRowBefore(1);
  }

  prepararLayoutPedidoUnidade_(aba, dataMes);

  var dataStart = CONFIG.PEDIDO.LINHA_DADOS || 3;
  var nCols = obterNumColsPedido_();
  var lr = Math.max(aba.getLastRow(), dataStart);
  var lc = Math.max(aba.getLastColumn(), lastCol);

  if (lr >= dataStart) {
    aba.getRange(dataStart, 1, lr - dataStart + 1, lc).clearContent();
  }

  if (linhasPedido.length) {
    aba.getRange(dataStart, 1, linhasPedido.length, nCols).setValues(linhasPedido);
  }

  while (aba.getLastColumn() > nCols) {
    try {
      aba.deleteColumn(nCols + 1);
    } catch (eDel) {
      break;
    }
  }
}

/**
 * Processa unidades em lotes (evita timeout do Google Apps Script).
 * @param {Date=} dataMes
 * @return {{ok:number, falhas:Array.<string>, concluido:boolean, total:number}}
 */
function migrarLayoutPedidoTodasUnidades_(dataMes) {
  var props = PropertiesService.getScriptProperties();
  var emAndamento = props.getProperty('MIGR_LAYOUT_EM_ANDAMENTO') === '1';
  if (!emAndamento) {
    iniciarMigracaoLayoutUnidadesEmLotes_(dataMes);
  }
  processarLoteMigracaoLayoutUnidades_();
  var ok = parseInt(props.getProperty('MIGR_LAYOUT_OK') || '0', 10);
  var falhas = JSON.parse(props.getProperty('MIGR_LAYOUT_FALHAS') || '[]');
  var idx = parseInt(props.getProperty('MIGR_LAYOUT_IDX') || '0', 10);
  var total = parseInt(props.getProperty('MIGR_LAYOUT_TOTAL') || '0', 10);
  var concluido = props.getProperty('MIGR_LAYOUT_EM_ANDAMENTO') !== '1';
  return { ok: ok, falhas: falhas, concluido: concluido, total: total, processadas: idx };
}

function iniciarMigracaoLayoutUnidadesEmLotes_(dataMes) {
  var unidades = obterListaUnidadesMigracaoLayout_();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('MIGR_LAYOUT_EM_ANDAMENTO', '1');
  props.setProperty('MIGR_LAYOUT_IDX', '0');
  props.setProperty('MIGR_LAYOUT_OK', '0');
  props.setProperty('MIGR_LAYOUT_FALHAS', '[]');
  props.setProperty('MIGR_LAYOUT_TOTAL', String(unidades.length));
  props.setProperty(
    'MIGR_LAYOUT_DATA',
    dataMes ? String(dataMes.getTime()) : String(Date.now())
  );
  removerGatilhosClockPorFuncao_('processarLoteMigracaoLayoutUnidades_');
}

function obterListaUnidadesMigracaoLayout_() {
  var unidades = listarUnidadesAtivas();
  if (unidades.length === 0 && DADOS_INICIAIS && DADOS_INICIAIS.unidades) {
    for (var d = 0; d < DADOS_INICIAIS.unidades.length; d++) {
      unidades.push({
        nome: DADOS_INICIAIS.unidades[d],
        spreadsheetId: DADOS_INICIAIS.idsPlanilhas[DADOS_INICIAIS.unidades[d]] || '',
        ativo: true
      });
    }
  }
  return unidades;
}

function processarLoteMigracaoLayoutUnidades_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('MIGR_LAYOUT_EM_ANDAMENTO') !== '1') return;

  var unidades = obterListaUnidadesMigracaoLayout_();
  var idx = parseInt(props.getProperty('MIGR_LAYOUT_IDX') || '0', 10);
  var ok = parseInt(props.getProperty('MIGR_LAYOUT_OK') || '0', 10);
  var falhas = JSON.parse(props.getProperty('MIGR_LAYOUT_FALHAS') || '[]');
  var dataMes = new Date(parseInt(props.getProperty('MIGR_LAYOUT_DATA') || String(Date.now()), 10));

  var LOTE = 4;
  var MAX_MS = 240000;
  var inicioMs = Date.now();
  var ssMae = null;

  try {
    ssMae = obterPlanilhaMae();
  } catch (eMae) {}

  var processadasNesteLote = 0;
  while (
    idx < unidades.length &&
    processadasNesteLote < LOTE &&
    Date.now() - inicioMs < MAX_MS
  ) {
    var u = unidades[idx];
    try {
      if (ssMae) {
        ssMae.toast(
          'Layout ' + (idx + 1) + '/' + unidades.length + ': ' + u.nome,
          'Unidades',
          10
        );
      }
      var id = String(u.spreadsheetId || '').trim();
      if (!id) id = resolverIdUnidadePeloNome_(u.nome);
      if (!id) throw new Error('ID não encontrado');
      var ssFilha = SpreadsheetApp.openById(id);
      var aba = obterAbaPedidoFilha_(ssFilha);
      migrarPlanilhaPedidoRemoverColunaData_(aba, dataMes);
      aplicarDropdownPedido(ssFilha, {
        somenteLinhasComDados: true,
        semCabecalho: true,
        linhaFim: Math.max(aba.getLastRow(), CONFIG.PEDIDO.LINHA_DADOS || 3)
      });
      ok++;
    } catch (e) {
      falhas.push(u.nome + ': ' + e.message);
    }
    idx++;
    processadasNesteLote++;
    try {
      SpreadsheetApp.flush();
    } catch (eFlush) {}
  }

  props.setProperty('MIGR_LAYOUT_IDX', String(idx));
  props.setProperty('MIGR_LAYOUT_OK', String(ok));
  props.setProperty('MIGR_LAYOUT_FALHAS', JSON.stringify(falhas));

  if (idx < unidades.length) {
    try {
      removerGatilhosClockPorFuncao_('processarLoteMigracaoLayoutUnidades_');
      ScriptApp.newTrigger('processarLoteMigracaoLayoutUnidades_')
        .timeBased()
        .after(4000)
        .create();
    } catch (eGat) {
      props.setProperty('MIGR_LAYOUT_EM_ANDAMENTO', '0');
      if (ssMae) {
        ssMae.toast('Gatilho falhou — rode o menu de novo', 'Layout', 15);
      }
    }
    return;
  }

  props.deleteProperty('MIGR_LAYOUT_EM_ANDAMENTO');
  removerGatilhosClockPorFuncao_('processarLoteMigracaoLayoutUnidades_');
  registrarLog('MIGR_LAYOUT_OK', ok + '/' + unidades.length + ' unidades');
  if (ssMae) {
    ssMae.toast('Layout concluído: ' + ok + ' unidades', 'Unidades', 12);
  }
}

/**
 * Menu: remove coluna Data e aplica faixa de mês em todas as unidades.
 */
function menuMigrarLayoutMesReferenciaUnidades() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Layout mês de referência',
    'Isso vai em TODAS as unidades:\n\n' +
      '• Remover a coluna Data\n' +
      '• Colocar “PLANILHA DE PEDIDOS — MÊS XX” no topo\n' +
      '• Manter pedidos já preenchidos\n\n' +
      'Pode levar alguns minutos (processa em lotes). Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Iniciando… processa unidade por unidade', 'Layout', 8);
    var r = migrarLayoutPedidoTodasUnidades_();
    SpreadsheetApp.flush();
    if (r.concluido) {
      ui.alert(
        'Layout atualizado',
        'Unidades processadas: ' +
          r.ok +
          '/' +
          r.total +
          '\nFalhas: ' +
          r.falhas.length +
          (r.falhas.length ? '\n\n- ' + r.falhas.slice(0, 6).join('\n- ') : '') +
          '\n\nRecarregue a planilha da unidade (F5).\nSomente Status tem seleção (Enviado).',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Processando em segundo plano',
        'Foram iniciadas ' +
          r.total +
          ' unidades.\n\n' +
          'O script continua sozinho (toast “Layout X/Y” na Mãe).\n' +
          'Aguarde 2–4 minutos e recarregue (F5) a planilha da unidade.\n\n' +
          'Não clique de novo no menu até terminar.',
        ui.ButtonSet.OK
      );
    }
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Planilha da unidade ainda com a coluna Estoque Atual.
 * @param {Sheet} aba
 * @return {boolean}
 */
function planilhaPedidoUsaLayoutComEstoque_(aba) {
  if (!aba) return false;
  var lastRow = Math.max(aba.getLastRow(), 1);
  var lastCol = Math.max(aba.getLastColumn(), obterNumColsPedido_(), 6);
  var vals = aba.getRange(1, 1, Math.min(15, lastRow), lastCol).getValues();
  var mapa = detectarColunasPedido_(vals);
  return mapa.colEstoque >= 0;
}

/**
 * Remove a coluna Estoque Atual das planilhas das unidades.
 * @param {Sheet} aba
 * @param {Date=} dataMes
 */
function migrarPlanilhaPedidoRemoverColunaEstoque_(aba, dataMes) {
  if (!aba) return;
  if (!planilhaPedidoUsaLayoutComEstoque_(aba)) {
    prepararLayoutPedidoUnidade_(aba, dataMes);
    return;
  }

  var lastRow = Math.max(aba.getLastRow(), 1);
  var lastCol = Math.max(aba.getLastColumn(), 6);
  var vals = aba.getRange(1, 1, lastRow, lastCol).getValues();
  var mapa = detectarColunasPedido_(vals);

  try {
    var filtro = aba.getFilter();
    if (filtro) filtro.remove();
  } catch (eFiltro) {}

  try {
    aba.getRange(1, 1, 1, lastCol).breakApart();
  } catch (eBreak) {}

  if (mapa.colEstoque >= 0) {
    aba.deleteColumn(mapa.colEstoque + 1);
  }

  prepararLayoutPedidoUnidade_(aba, dataMes);

  var nCols = obterNumColsPedido_();
  while (aba.getLastColumn() > nCols) {
    try {
      aba.deleteColumn(nCols + 1);
    } catch (eDel) {
      break;
    }
  }
}

/**
 * Faixa de mês + cabeçalho formal (sem coluna Data).
 * @param {Sheet} aba
 * @param {Date=} dataMes
 */
function prepararLayoutPedidoUnidade_(aba, dataMes) {
  var header = obterHeaderPedidoUnidade_();
  var headerLinha = CONFIG.PEDIDO.HEADER_LINHA || 2;
  var nCols = obterNumColsPedido_();
  aplicarMesReferenciaPedidoUnidade_(aba, dataMes);
  aba.getRange(headerLinha, 1, 1, nCols).setValues([header]);
}

/**
 * @param {Date=} data
 * @return {Date}
 */
function obterDataMesReferenciaPedido_(data) {
  data = data || new Date();
  var tz = Session.getScriptTimeZone();
  var mes = parseInt(Utilities.formatDate(data, tz, 'MM'), 10);
  var ano = parseInt(Utilities.formatDate(data, tz, 'yyyy'), 10);
  return new Date(ano, mes - 1, 1);
}

/**
 * Atualiza o mês de referência em todas as unidades ativas.
 * @param {Date=} data
 * @return {{ok:number, falhas:Array.<string>}}
 */
function atualizarMesReferenciaTodasUnidades_(data) {
  var resultado = { ok: 0, falhas: [] };
  var unidades = listarUnidadesAtivas();
  if (unidades.length === 0 && DADOS_INICIAIS && DADOS_INICIAIS.unidades) {
    for (var d = 0; d < DADOS_INICIAIS.unidades.length; d++) {
      unidades.push({
        nome: DADOS_INICIAIS.unidades[d],
        spreadsheetId: DADOS_INICIAIS.idsPlanilhas[DADOS_INICIAIS.unidades[d]] || '',
        ativo: true
      });
    }
  }

  for (var i = 0; i < unidades.length; i++) {
    var u = unidades[i];
    try {
      var id = String(u.spreadsheetId || '').trim();
      if (!id) id = resolverIdUnidadePeloNome_(u.nome);
      if (!id) throw new Error('ID não encontrado');
      var ssFilha = SpreadsheetApp.openById(id);
      var aba = obterAbaPedidoFilha_(ssFilha);
      aplicarMesReferenciaPedidoUnidade_(aba, data);
      resultado.ok++;
    } catch (e) {
      resultado.falhas.push(u.nome + ': ' + e.message);
    }
  }
  return resultado;
}
