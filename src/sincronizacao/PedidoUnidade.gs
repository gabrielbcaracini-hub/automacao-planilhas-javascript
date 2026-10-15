/**
 * Funções exclusivas das planilhas das unidades:
 * envio em lote do pedido e relatório Excel do que foi solicitado.
 */

/**
 * @return {boolean}
 */
function ehPlanilhaUnidadeAtiva_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return false;
    if (String(ss.getId()) === String(CONFIG.SPREADSHEET_MAE_ID || '').trim()) {
      return false;
    }
    var aba = ss.getSheetByName(CONFIG.ABAS.PEDIDO);
    if (aba) return true;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (ehAbaPedido_(sheets[i])) return true;
    }
  } catch (e) {}
  return false;
}

/**
 * @param {Spreadsheet} ss
 * @return {string}
 */
function obterNomeUnidadeDaPlanilha_(ss) {
  var unidade = encontrarUnidadePorId(ss.getId());
  if (unidade && unidade.nome) return unidade.nome;
  return String(ss.getName() || 'Unidade')
    .replace(/^Pedido\s*-\s*/i, '')
    .trim() || 'Unidade';
}

/**
 * @param {*} status
 * @return {string} rascunho|enviado|consolidado|outro
 */
function normalizarStatusPedido_(status) {
  var st = String(status || '')
    .toLowerCase()
    .replace(/á|à|ã|â/g, 'a')
    .replace(/é|ê/g, 'e')
    .replace(/í/g, 'i')
    .replace(/ó|ô|õ/g, 'o')
    .replace(/ú/g, 'u')
    .replace(/ç/g, 'c')
    .replace(/[^a-z]/g, '');
  if (st.indexOf('consolidado') !== -1) return 'consolidado';
  if (st.indexOf('enviado') !== -1) return 'enviado';
  if (st.indexOf('rascunho') !== -1) return 'rascunho';
  return st || 'rascunho';
}

/**
 * @param {*} valor
 * @param {*} display
 * @return {string}
 */
function obterTextoQuantidadePedido_(valor, display) {
  var qtdTexto =
    valor === null || valor === undefined || valor === '' ? '' : String(valor).trim();
  if (!qtdTexto && display !== null && display !== undefined) {
    qtdTexto = String(display).trim();
  }
  return qtdTexto;
}

/**
 * Marca como Enviado todos os itens com quantidade preenchida (exceto Consolidado).
 * @param {Spreadsheet} ss
 * @param {Sheet} aba
 * @return {{marcados:number, jaEnviados:number, consolidados:number}}
 */
function enviarPedidoCompletoUnidade_(ss, aba) {
  var rng = aba.getDataRange();
  var values = rng.getValues();
  var displays = rng.getDisplayValues();
  var mapa = detectarColunasPedido_(displays);
  if (mapa.colStatus < 0 || mapa.colQtd < 0) {
    mapa = detectarColunasPedido_(values);
  }
  if (mapa.colStatus < 0 || mapa.colQtd < 0) {
    throw new Error(
      'Não foi possível localizar as colunas “Quantidade do Pedido” e “Status”.'
    );
  }

  var colStatus = mapa.colStatus;
  var colQtd = mapa.colQtd;
  var colProduto = mapa.colProduto;
  var inicio = mapa.headerRow + 1;
  var linhasMarcar = [];
  var jaEnviados = 0;
  var consolidados = 0;

  for (var i = inicio; i < values.length; i++) {
    var produto = String(displays[i][colProduto] || values[i][colProduto] || '').trim();
    if (!produto || ehLinhaCategoriaPedido_(produto)) continue;

    var qtdTexto = obterTextoQuantidadePedido_(values[i][colQtd], displays[i][colQtd]);
    if (!qtdTexto || qtdTexto === '0') continue;

    var status = String(displays[i][colStatus] || values[i][colStatus] || '').trim();
    var stNorm = normalizarStatusPedido_(status);
    if (stNorm === 'consolidado') {
      consolidados++;
      continue;
    }
    if (stNorm === 'enviado') {
      jaEnviados++;
      continue;
    }
    linhasMarcar.push(i + 1);
  }

  if (!linhasMarcar.length && !jaEnviados) {
    throw new Error(
      'Nenhum item com quantidade preenchida para enviar.\n\n' +
        'Preencha a coluna “Quantidade do Pedido” nos itens desejados.'
    );
  }

  if (linhasMarcar.length) {
    var colStatus1Based = colStatus + 1;
    for (var m = 0; m < linhasMarcar.length; m++) {
      aba.getRange(linhasMarcar[m], colStatus1Based).setValue(CONFIG.VALORES.STATUS_ENVIADO);
    }
  }

  return {
    marcados: linhasMarcar.length,
    jaEnviados: jaEnviados,
    consolidados: consolidados
  };
}

/**
 * Menu da unidade: envia de uma vez todos os itens com quantidade.
 */
function menuEnviarPedidoCompleto() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resp = ui.alert(
    'Enviar pedido completo',
    'Todos os itens com “Quantidade do Pedido” preenchida serão marcados como Enviado nesta planilha.\n\n' +
      'Os pedidos só aparecerão no Resumo da Planilha Mãe quando o administrador usar ' +
      '“Puxar pedidos das unidades”.\n\nItens já Enviados ou Consolidados não serão alterados.\n\nContinuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    ss.toast('Marcando itens como Enviado…', 'Pedido', 30);
    var aba = obterAbaPedidoFilha_(ss);
    var resultado = enviarPedidoCompletoUnidade_(ss, aba);
    SpreadsheetApp.flush();

    ui.alert(
      'Pedido marcado como Enviado',
      'Novos itens marcados como Enviado: ' +
        resultado.marcados +
        (resultado.jaEnviados
          ? '\nJá estavam Enviados: ' + resultado.jaEnviados
          : '') +
        (resultado.consolidados
          ? '\nConsolidados (não alterados): ' + resultado.consolidados
          : '') +
        '\n\nPronto! Aguarde o administrador puxar os pedidos na Planilha Mãe ' +
        '(Sistema de Compras → Puxar pedidos das unidades).',
      ui.ButtonSet.OK
    );
    try {
      registrarLog(
        'ENVIO_PEDIDO_UNIDADE',
        obterNomeUnidadeDaPlanilha_(ss) +
          ' | novos=' +
          resultado.marcados +
          ' | ja=' +
          resultado.jaEnviados
      );
    } catch (eLog) {}
  } catch (e) {
    ui.alert('Erro ao enviar pedido', e.message, ui.ButtonSet.OK);
    try {
      registrarLog('ENVIO_PEDIDO_UNIDADE_ERRO', e.message);
    } catch (eLog2) {}
  }
}

/**
 * Itens com quantidade para o relatório Excel da unidade.
 * @param {Sheet} aba
 * @return {Array.<Object>}
 */
function coletarItensPedidoUnidadeParaRelatorio_(aba) {
  var lido = lerPedidosEnviadosFlexivel(aba);
  return lido.todas.filter(function (item) {
    var qtd = String(item.quantidade || '').trim();
    return qtd !== '' && qtd !== '0';
  });
}

/**
 * Converte itens da unidade para o formato usado pelo relatório Excel da Mãe.
 * @param {Sheet} aba
 * @param {string} nomeUnidade
 * @return {Array.<Array.<*>>}
 */
function montarLinhasConsolidadosUnidade_(aba, nomeUnidade) {
  var itens = coletarItensPedidoUnidadeParaRelatorio_(aba);
  if (!itens.length) {
    throw new Error(
      'Não há itens com quantidade preenchida para gerar o relatório.\n\nPreencha o pedido antes de exportar.'
    );
  }
  return itens.map(function (item) {
    return [
      nomeUnidade,
      item.produto,
      item.estoqueAtual,
      item.quantidade,
      item.observacao || '',
      ''
    ];
  });
}

/**
 * Menu da unidade: Excel com o resumo do pedido (mesmo modelo da Planilha Mãe).
 */
function menuGerarRelatorioExcelUnidade() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = obterAbaPedidoFilha_(ss);
    var nomeUnidade = obterNomeUnidadeDaPlanilha_(ss);
    ss.toast('Montando relatório Excel…', 'Excel', 45);
    var linhas = montarLinhasConsolidadosUnidade_(aba, nomeUnidade);
    var agora = new Date();
    var dataArquivo = Utilities.formatDate(
      agora,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd_HH-mm'
    );
    var arquivo = gerarRelatorioExcel_(ss, linhas, {
      nomeBase: 'Relatório de Pedidos - ' + nomeUnidade + ' - ' + dataArquivo,
      ocultarPregao: true
    });
    ui.alert(
      'Relatório Excel gerado',
      'Unidade: ' +
        nomeUnidade +
        '\nArquivo: ' +
        arquivo.getName() +
        '\n\nAbra ou baixe pelo link:\n' +
        arquivo.getUrl(),
      ui.ButtonSet.OK
    );
    try {
      registrarLog('EXCEL_PEDIDO_UNIDADE', nomeUnidade + ' | ' + arquivo.getUrl());
    } catch (eLog) {}
  } catch (e) {
    ui.alert('Erro ao gerar Excel', e.message, ui.ButtonSet.OK);
    try {
      registrarLog('EXCEL_PEDIDO_UNIDADE_ERRO', e.message);
    } catch (eLog2) {}
  }
}

/**
 * Monta o menu nas planilhas das unidades.
 */
function onOpenUnidade_() {
  try {
    ajustarLayoutPedidoUnidadeAoAbrir_();
  } catch (eOpen) {}
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sistema de Compras')
    .addItem('📤 Enviar pedido completo', 'menuEnviarPedidoCompleto')
    .addItem('📥 Baixar relatório Excel', 'menuGerarRelatorioExcelUnidade')
    .addToUi();
}

/**
 * Remove Estoque Atual (se ainda existir) e reaplica o destaque do Status.
 * @param {Spreadsheet=} ss
 */
function ajustarLayoutPedidoUnidadeAoAbrir_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var aba = obterAbaPedidoFilha_(ss);
  if (!aba) return;
  if (planilhaPedidoUsaLayoutComEstoque_(aba)) {
    migrarPlanilhaPedidoRemoverColunaEstoque_(aba);
  }
  var headerLinha = obterLinhaCabecalhoPedido_(aba);
  var dataStart = headerLinha + 1;
  var endRow = Math.max(aba.getLastRow(), dataStart);
  if (endRow < dataStart) return;
  var t = CONFIG.TEMA;
  aplicarVisualColunaQuantidade_(aba, headerLinha, dataStart, endRow, t);
  aplicarVisualColunaStatus_(aba, headerLinha, dataStart, endRow, t);
  aplicarFormatacaoCondicionalStatus_(
    aba,
    dataStart,
    endRow - dataStart + 1,
    CONFIG.PEDIDO.STATUS
  );
  aplicarVisualCategoriasPedido_(aba, dataStart, endRow, t);
}

/** Gatilho instalável ao abrir planilha de unidade. */
function onOpenUnidadeInstalavel_() {
  onOpenUnidade_();
}

/**
 * Instala gatilho onOpen nas unidades para exibir o menu Sistema de Compras.
 * @return {{instalados:number, falhas:string[]}}
 */
function instalarGatilhoOnOpenUnidades_() {
  var resultado = { instalados: 0, falhas: [] };
  removerGatilhosOnOpenUnidades_();
  var unidades = listarUnidadesAtivas();
  for (var i = 0; i < unidades.length; i++) {
    var id = String(unidades[i].spreadsheetId || '').trim();
    if (!id) {
      resultado.falhas.push(unidades[i].nome + ': sem ID');
      continue;
    }
    try {
      ScriptApp.newTrigger('onOpenUnidadeInstalavel_').forSpreadsheet(id).onOpen().create();
      resultado.instalados++;
    } catch (e) {
      resultado.falhas.push(unidades[i].nome + ': ' + e.message);
    }
  }
  return resultado;
}

/** Remove gatilhos onOpen das unidades (evita duplicados). */
function removerGatilhosOnOpenUnidades_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = triggers.length - 1; i >= 0; i--) {
    if (triggers[i].getHandlerFunction() === 'onOpenUnidadeInstalavel_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
