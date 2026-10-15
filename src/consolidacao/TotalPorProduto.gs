/**
 * Aba Pedidos das Unidades — situação simples da rodada:
 * unidades que já tiveram o pedido puxado pela Mãe × unidades ainda pendentes.
 */

/**
 * Garante a aba com o nome novo. Se existir “Total por Produto”, só renomeia.
 * @param {Spreadsheet} ssMae
 * @return {Sheet}
 */
function obterOuCriarAbaPedidosDasUnidades_(ssMae) {
  var nomeNovo = CONFIG.ABAS.TOTAL_PRODUTO;
  var aba = ssMae.getSheetByName(nomeNovo);
  if (aba) return aba;

  var antigos = [
    CONFIG.ABAS.TOTAL_PRODUTO_LEGADO,
    'Total por Produto',
    'Total de Produtos',
    'Total de Produto',
    'TOTAL DE PRODUTOS',
    'TOTAL DE PRODUTO'
  ];
  for (var i = 0; i < antigos.length; i++) {
    if (!antigos[i]) continue;
    aba = ssMae.getSheetByName(antigos[i]);
    if (!aba && typeof obterAbaPorNome_ === 'function') {
      aba = obterAbaPorNome_(ssMae, antigos[i]);
    }
    if (aba) {
      try {
        aba.setName(nomeNovo);
      } catch (eNome) {
        Logger.log('renomear Pedidos das Unidades: ' + eNome.message);
      }
      return aba;
    }
  }

  return ssMae.insertSheet(nomeNovo);
}

/**
 * Lê Consolidados e redesenha a aba (usado no onOpen e no pull).
 * @param {Spreadsheet=} ssMae
 */
function atualizarAbaSituacaoPedidosMae_(ssMae) {
  ssMae = ssMae || obterPlanilhaMae();
  var abaCons = ssMae.getSheetByName(CONFIG.ABAS.CONSOLIDADOS);
  var linhas = abaCons ? lerDados(abaCons, CONFIG.CONSOLIDADOS.HEADER.length) : [];
  escreverTotalPorProdutoFromConsolidados_(ssMae, linhas || [], {
    modoLeve: true,
    varrerFilhas: false
  });
}

/**
 * @param {Spreadsheet} ssMae
 * @param {Array.<Array.<*>>} linhasCons
 * @param {Object=} opcoes
 */
function escreverTotalPorProdutoFromConsolidados_(ssMae, linhasCons, opcoes) {
  opcoes = opcoes || {};
  var modoLeve = opcoes.modoLeve === true;
  var nCols = 2;
  var situacao = montarSituacaoPedidosUnidades_(linhasCons, ssMae, opcoes);
  var linhasPlanilha = montarLinhasSituacaoPedidos_(situacao);

  var aba = obterOuCriarAbaPedidosDasUnidades_(ssMae);

  try {
    var filtro = aba.getFilter();
    if (filtro) filtro.remove();
  } catch (eF) {}

  try {
    aba.getRange(1, 1, Math.min(aba.getMaxRows(), 80), Math.min(aba.getMaxColumns(), 6)).breakApart();
  } catch (eBr) {}

  var lastR = Math.max(aba.getLastRow(), 20);
  var lastC = Math.max(aba.getLastColumn(), 3, 5);
  try {
    aba.getRange(1, 1, lastR, lastC).clearContent();
    aba.getRange(1, 1, lastR, lastC).clearFormat();
  } catch (eLimpa) {}

  aba.getRange(1, 1, 1, nCols).setValues([
    [
      'JÁ PEDIRAM — ' + situacao.concluidas.length + '/' + situacao.total,
      'AINDA NÃO PEDIRAM — ' + situacao.pendentes.length + '/' + situacao.total
    ]
  ]);

  if (linhasPlanilha.length > 0) {
    aba.getRange(2, 1, linhasPlanilha.length, nCols).setValues(linhasPlanilha);
    var primeiraSobra = 2 + linhasPlanilha.length;
    if (lastR >= primeiraSobra) {
      aba.getRange(primeiraSobra, 1, lastR - primeiraSobra + 1, lastC).clearContent();
    }
  } else if (lastR >= 2) {
    aba.getRange(2, 1, lastR - 1, lastC).clearContent();
  }

  limparColunasExtrasTotalPorProduto_(aba, nCols, Math.max(lastR, linhasPlanilha.length + 1));
  aplicarVisualSituacaoPedidos_(aba, nCols, linhasPlanilha.length, situacao, modoLeve);
}

/**
 * Lista unidades ativas que já pediram e as que ainda não pediram.
 * Fontes: Consolidados, Resumo (colunas com quantidade) e, se pedido, as planilhas filhas.
 * @param {Array.<Array.<*>>} linhasCons
 * @param {Spreadsheet=} ssMae
 * @param {Object=} opcoes
 * @return {{concluidas:string[], pendentes:string[], total:number}}
 */
function montarSituacaoPedidosUnidades_(linhasCons, ssMae, opcoes) {
  opcoes = opcoes || {};
  var nomes = listarNomesUnidadesSituacao_(ssMae);
  var jaPediu = coletarChavesUnidadesQueJaPediram_(linhasCons, ssMae, opcoes);

  var concluidas = [];
  var pendentes = [];
  for (var u = 0; u < nomes.length; u++) {
    if (jaPediu[chaveUnidadeSituacao_(nomes[u])]) {
      concluidas.push(nomes[u]);
    } else {
      pendentes.push(nomes[u]);
    }
  }

  return {
    concluidas: concluidas,
    pendentes: pendentes,
    total: nomes.length
  };
}

/**
 * @param {*} nome
 * @return {string}
 */
function chaveUnidadeSituacao_(nome) {
  if (typeof normalizarChaveUnidade_ === 'function') {
    return normalizarChaveUnidade_(nome);
  }
  return String(nome || '')
    .toLowerCase()
    .replace(/á|à|ã|â/g, 'a')
    .replace(/é|ê/g, 'e')
    .replace(/í/g, 'i')
    .replace(/ó|ô|õ/g, 'o')
    .replace(/ú|ü/g, 'u')
    .replace(/ç/g, 'c')
    .replace(/[^a-z0-9]/g, '');
}

/**
 * Une todas as fontes que indicam “esta unidade já pediu”.
 * @param {Array.<Array.<*>>} linhasCons
 * @param {Spreadsheet=} ssMae
 * @param {boolean=} varrerFilhas
 * @return {Object.<string, boolean>}
 */
function coletarChavesUnidadesQueJaPediram_(linhasCons, ssMae, varrerFilhas) {
  var opcoes = varrerFilhas;
  if (opcoes === true || opcoes === false || opcoes == null) {
    opcoes = { varrerFilhas: varrerFilhas === true };
  }
  var jaPediu = {};

  function marcar(nome) {
    var chave = chaveUnidadeSituacao_(nome);
    if (chave) jaPediu[chave] = true;
  }

  for (var i = 0; i < (linhasCons || []).length; i++) {
    var unidade = String((linhasCons[i] && linhasCons[i][0]) || '').trim();
    if (unidade) marcar(unidade);
  }

  if (!opcoes.ignorarResumo) {
    try {
      marcarUnidadesComQuantidadeNoResumo_(ssMae, marcar);
    } catch (eRes) {}
  }

  if (opcoes.varrerFilhas) {
    try {
      marcarUnidadesComPedidoEnviadoNasFilhas_(ssMae, marcar);
    } catch (eFil) {}
  }

  return jaPediu;
}

/**
 * No Resumo (ITEM × unidades), qualquer quantidade preenchida conta como pedido.
 * @param {Spreadsheet} ssMae
 * @param {Function} marcar
 */
function marcarUnidadesComQuantidadeNoResumo_(ssMae, marcar) {
  if (!ssMae) return;
  var aba = ssMae.getSheetByName(CONFIG.ABAS.RESUMO);
  if (!aba || aba.getLastRow() < 2 || aba.getLastColumn() < 2) return;

  var lastCol = Math.min(aba.getLastColumn(), 40);
  var lastRow = Math.min(aba.getLastRow(), 500);
  var header = aba.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var corpo = aba.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  for (var c = 1; c < header.length; c++) {
    var nomeUn = String(header[c] || '').trim();
    if (!nomeUn) continue;
    var h = nomeUn.toLowerCase();
    if (h.indexOf('última') !== -1 || h.indexOf('ultima') !== -1) continue;
    if (h.indexOf('envio') !== -1 || h.indexOf('item') === 0) continue;

    for (var r = 0; r < corpo.length; r++) {
      var q = String(corpo[r][c] || '').trim();
      if (!q) continue;
      if (q.indexOf('▸') === 0) continue;
      marcar(nomeUn);
      break;
    }
  }
}

/**
 * Abre cada unidade e marca as que já têm item Enviado com quantidade.
 * @param {Spreadsheet} ssMae
 * @param {Function} marcar
 */
function marcarUnidadesComPedidoEnviadoNasFilhas_(ssMae, marcar) {
  var unidades = [];
  try {
    unidades = listarUnidadesAtivas(ssMae);
  } catch (eLista) {
    return;
  }

  for (var i = 0; i < unidades.length; i++) {
    var u = unidades[i];
    try {
      var id = String(u.spreadsheetId || '').trim();
      if (!id) continue;
      var ssFilha = SpreadsheetApp.openById(id);
      var aba = obterAbaPedidoFilha_(ssFilha);
      var lido = lerPedidosEnviadosFlexivel(aba);
      if (lido && lido.enviados && lido.enviados.length > 0) {
        marcar(u.nome);
      }
    } catch (eUn) {}
  }
}

/**
 * @param {{concluidas:string[], pendentes:string[]}} situacao
 * @return {Array.<Array.<string>>}
 */
function montarLinhasSituacaoPedidos_(situacao) {
  var n = Math.max(situacao.concluidas.length, situacao.pendentes.length, 1);
  var linhas = [];
  for (var i = 0; i < n; i++) {
    var ok = situacao.concluidas[i] || '';
    var falta = situacao.pendentes[i] || '';
    if (i === 0 && !ok && situacao.concluidas.length === 0) {
      ok = 'Nenhuma ainda';
    }
    if (i === 0 && !falta && situacao.pendentes.length === 0) {
      falta = 'Nenhuma — todas já pediram';
    }
    linhas.push([ok, falta]);
  }
  return linhas;
}

/**
 * Nomes das unidades ativas; se a aba Unidades falhar, usa a lista oficial.
 * @param {Spreadsheet=} ssMae
 * @return {string[]}
 */
function listarNomesUnidadesSituacao_(ssMae) {
  var nomes = [];
  try {
    nomes = listarTodasUnidades(ssMae)
      .filter(function (u) {
        return u.ativo !== false && u.nome;
      })
      .map(function (u) {
        return String(u.nome || '').trim();
      })
      .filter(function (n) {
        return n;
      });
  } catch (eLista) {}

  if (!nomes.length && typeof DADOS_INICIAIS !== 'undefined' && DADOS_INICIAIS.unidades) {
    nomes = DADOS_INICIAIS.unidades.slice();
  }
  return ordenarNomesUnidadesSituacao_(nomes);
}

/**
 * Ordem oficial das unidades (UNIDADE CENTRO, …, UNIDADE MODELO).
 * @param {string[]} nomes
 * @return {string[]}
 */
function ordenarNomesUnidadesSituacao_(nomes) {
  var ordem =
    typeof DADOS_INICIAIS !== 'undefined' && DADOS_INICIAIS.unidades ? DADOS_INICIAIS.unidades : [];
  var indice = {};
  for (var i = 0; i < ordem.length; i++) {
    indice[normalizarTexto(ordem[i])] = i;
  }
  return (nomes || []).slice().sort(function (a, b) {
    var ia = indice.hasOwnProperty(normalizarTexto(a)) ? indice[normalizarTexto(a)] : 999;
    var ib = indice.hasOwnProperty(normalizarTexto(b)) ? indice[normalizarTexto(b)] : 999;
    if (ia !== ib) return ia - ib;
    return String(a).localeCompare(String(b), 'pt-BR');
  });
}

/** Opções padrão para sync automático (modo leve). */
function opcoesSyncAutomaticoMae_() {
  return { modoLeve: true };
}

/**
 * Zera a aba na Nova rodada: todas as unidades voltam para pendente.
 * @param {Spreadsheet} ssMae
 */
function limparTotalPorProdutoNovaRodada_(ssMae) {
  try {
    removerGatilhosClockPorFuncao_('processarTotalPorProdutoPendente_');
    PropertiesService.getScriptProperties().deleteProperty('TOTAL_PRODUTO_PENDENTE');
  } catch (eGat) {}
  escreverTotalPorProdutoFromConsolidados_(ssMae, [], {
    modoLeve: true,
    ignorarResumo: true
  });
}

/**
 * Agenda atualização ~8s depois (evita sobrecarga junto com Resumo).
 */
function agendarAtualizacaoTotalPorProduto_() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('TOTAL_PRODUTO_PENDENTE', '1');
  try {
    removerGatilhosClockPorFuncao_('processarTotalPorProdutoPendente_');
    ScriptApp.newTrigger('processarTotalPorProdutoPendente_').timeBased().after(8000).create();
  } catch (e) {
    Logger.log('agendarAtualizacaoTotalPorProduto_: ' + e.message);
  }
}

/** Handler one-shot: lê Consolidados e atualiza a situação das unidades. */
function processarTotalPorProdutoPendente_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    agendarAtualizacaoTotalPorProduto_();
    return;
  }
  try {
    var ss = obterPlanilhaMae();
    var abaCons = ss.getSheetByName(CONFIG.ABAS.CONSOLIDADOS);
    var linhas = abaCons ? lerDados(abaCons, CONFIG.CONSOLIDADOS.HEADER.length) : [];
    escreverTotalPorProdutoFromConsolidados_(ss, linhas, { modoLeve: true });
    PropertiesService.getScriptProperties().deleteProperty('TOTAL_PRODUTO_PENDENTE');
    SpreadsheetApp.flush();
  } catch (e) {
    registrarLog('TOTAL_PRODUTO_PENDENTE_ERRO', e.message);
  } finally {
    lock.releaseLock();
    removerGatilhosClockPorFuncao_('processarTotalPorProdutoPendente_');
  }
}

/**
 * Remove colunas de versões anteriores (Produto / Total / Quem pediu).
 * @param {Sheet} aba
 * @param {number} nCols
 * @param {number} ateLinha
 */
function limparColunasExtrasTotalPorProduto_(aba, nCols, ateLinha) {
  try {
    var ateCol = Math.max(aba.getLastColumn(), nCols + 2, 8);
    if (ateCol <= nCols) return;
    var numLinhas = Math.min(aba.getMaxRows(), Math.max(ateLinha + 20, 40));
    var numColsExtra = ateCol - nCols;
    var rng = aba.getRange(1, nCols + 1, numLinhas, numColsExtra);
    rng.clearContent();
    rng.clearFormat();
    rng.setBackground(CONFIG.TEMA.SHEET_BG);
  } catch (e) {}
}

/**
 * Menu: abre a aba e atualiza a situação das unidades.
 */
function menuAbrirTotalPorProduto() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = obterPlanilhaMae();
    var abaCons = ss.getSheetByName(CONFIG.ABAS.CONSOLIDADOS);
    var linhas = abaCons ? lerDados(abaCons, CONFIG.CONSOLIDADOS.HEADER.length) : [];
    var situacao = montarSituacaoPedidosUnidades_(linhas, ss, { varrerFilhas: true });
    escreverTotalPorProdutoFromConsolidados_(ss, linhas, {
      modoLeve: false,
      varrerFilhas: true
    });
    var aba = obterOuCriarAbaPedidosDasUnidades_(ss);
    if (aba) {
      ss.setActiveSheet(aba);
    }
    SpreadsheetApp.flush();
    ui.alert(
      'Pedidos das unidades',
      situacao.total
        ? situacao.concluidas.length +
            ' unidade(s) já pediram.\n' +
            situacao.pendentes.length +
            ' ainda não pediram.\n\n' +
            'A lista já cruza Consolidados, Resumo e pedidos Enviado das unidades.'
        : 'Nenhuma unidade ativa cadastrada.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}

/**
 * @param {Sheet} aba
 * @param {number} nCols
 * @param {number} nLinhasCorpo
 * @param {Object} situacao
 * @param {boolean=} modoLeve
 */
function aplicarVisualSituacaoPedidos_(aba, nCols, nLinhasCorpo, situacao, modoLeve) {
  var t = CONFIG.TEMA;

  try {
    aba.setTabColor('#059669');
  } catch (eTab) {}

  try {
    var maxR = Math.min(aba.getMaxRows(), Math.max(nLinhasCorpo + 30, 60));
    var maxC = Math.max(nCols + 2, 8);
    aba.getRange(1, 1, maxR, maxC).setBackground(t.SHEET_BG);
  } catch (eBg) {}

  aba.setFrozenRows(1);
  try {
    aba.setFrozenColumns(0);
  } catch (eFr) {}

  aba
    .getRange(1, 1)
    .setBackground('#059669')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontFamily('Arial')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  aba
    .getRange(1, 2)
    .setBackground('#D97706')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontFamily('Arial')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  aba.setRowHeight(1, 46);

  aba.setColumnWidth(1, 360);
  aba.setColumnWidth(2, 360);

  if (nLinhasCorpo > 0) {
    var fundos = [];
    var cores = [];
    var pesos = [];
    for (var r = 0; r < nLinhasCorpo; r++) {
      var temOk = situacao && situacao.concluidas && situacao.concluidas[r];
      var temPend = situacao && situacao.pendentes && situacao.pendentes[r];
      fundos.push([temOk ? '#052E16' : t.ROW_A, temPend ? '#431407' : t.ROW_A]);
      cores.push([temOk ? '#86EFAC' : '#94A3B8', temPend ? '#FCD34D' : '#94A3B8']);
      pesos.push([temOk ? 'bold' : 'normal', temPend ? 'bold' : 'normal']);
    }

    aba
      .getRange(2, 1, nLinhasCorpo, nCols)
      .setBackgrounds(fundos)
      .setFontColors(cores)
      .setFontWeights(pesos)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setFontFamily('Arial')
      .setFontSize(12)
      .setWrap(true);

    for (var h = 0; h < nLinhasCorpo; h++) {
      aba.setRowHeight(h + 2, 34);
    }
  }

  aba
    .getRange(1, 1, 1 + nLinhasCorpo, nCols)
    .setBorder(true, true, true, true, true, true, t.BORDER, SpreadsheetApp.BorderStyle.SOLID);
}
