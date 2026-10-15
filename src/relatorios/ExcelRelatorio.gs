/**
 * Relatorio Excel (.xlsx) de pedidos do mes, agrupado por pregao.
 * Hierarquia: (PREGAO) -> PRODUTO (so os pedidos) -> UNIDADES + QTE.
 */

/**
 * Menu: gera arquivo Excel no modelo resumo por pregao.
 */
function menuGerarRelatorioExcel() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = obterPlanilhaMae();
    var linhas = obterLinhasConsParaRelatorioExcel_(ss);
    if (!linhas.length) {
      throw new Error(
        'Nao ha quantidades no Resumo/Consolidados para gerar o Excel.\n\n' +
          '1) Confira se as unidades marcaram Status = Enviado\n' +
          '2) Na Mae: Sistema de Compras -> Puxar pedidos das unidades\n' +
          '3) Depois gere o Excel de novo'
      );
    }

    ss.toast('Montando relatorio Excel...', 'Excel', 60);
    var arquivo = gerarRelatorioExcel_(ss, linhas);
    ui.alert(
      'Relatorio Excel gerado',
      'Arquivo: ' +
        arquivo.getName() +
        '\n\nLayout:\n' +
        '(PREGAO):\n' +
        'PRODUTO: (somente os que tiveram pedido)\n' +
        'UNIDADES: nome da unidade + QTE\n\n' +
        'Abra ou baixe pelo link:\n' +
        arquivo.getUrl(),
      ui.ButtonSet.OK
    );
    registrarLog('EXCEL_RELATORIO', arquivo.getName() + ' | ' + arquivo.getUrl());
  } catch (e) {
    ui.alert('Erro ao gerar Excel', e.message, ui.ButtonSet.OK);
    try {
      registrarLog('EXCEL_RELATORIO_ERRO', e.message);
    } catch (eLog) {}
  }
}

/**
 * Fonte do Excel: Consolidados; se vazio, le as quantidades da aba Resumo.
 * @param {Spreadsheet} ss
 * @return {Array.<Array.<*>>}
 */
function obterLinhasConsParaRelatorioExcel_(ss) {
  var linhas = [];
  try {
    var aba = ss.getSheetByName(CONFIG.ABAS.CONSOLIDADOS);
    if (aba) {
      linhas = lerDados(aba, CONFIG.CONSOLIDADOS.HEADER.length) || [];
      linhas = linhas.filter(function (row) {
        return (
          String(row[0] || '').trim() &&
          String(row[1] || '').trim() &&
          pedidoTemQuantidadeRelatorio_(row[3])
        );
      });
    }
  } catch (eCons) {}

  if (linhas.length) return linhas;

  try {
    linhas = extrairLinhasConsDaAbaResumo_(ss);
  } catch (eRes) {
    Logger.log('extrairLinhasConsDaAbaResumo_: ' + eRes.message);
  }
  return linhas || [];
}

/**
 * Quantidade preenchida e maior que zero (pedido de fato).
 * @param {*} valor
 * @return {boolean}
 */
function pedidoTemQuantidadeRelatorio_(valor) {
  if (valor === null || valor === undefined) return false;
  var s = String(valor).trim();
  if (!s) return false;
  var n = extrairNumeroQuantidade(valor);
  if (n === null) return true;
  return n > 0;
}

/**
 * Converte a matriz do Resumo (ITEM x unidades) em linhas de Consolidados.
 * @param {Spreadsheet} ss
 * @return {Array.<Array.<*>>}
 */
function extrairLinhasConsDaAbaResumo_(ss) {
  var aba = ss.getSheetByName(CONFIG.ABAS.RESUMO);
  if (!aba || aba.getLastRow() < 2) return [];

  var nCols = Math.min(Math.max(aba.getLastColumn(), 2), 40);
  var nRows = Math.min(aba.getLastRow(), 4000);
  var vals = aba.getRange(1, 1, nRows, nCols).getDisplayValues();
  var header = vals[0] || [];
  var unidades = [];
  for (var c = 1; c < header.length; c++) {
    var nomeUn = String(header[c] || '').trim();
    if (nomeUn) unidades.push({ col: c, nome: nomeUn });
  }
  if (!unidades.length) return [];

  var linhas = [];
  for (var r = 1; r < vals.length; r++) {
    var bruto = String(vals[r][0] || '').trim();
    if (!bruto) continue;
    if (/^▸/.test(bruto)) continue;
    if (/^PREG[AÃÁÀÂ]O\b/i.test(bruto)) continue;
    if (/^SEM PREG/i.test(bruto)) continue;

    var produto = limparNomeProdutoResumo_(bruto.replace(/^\s+/, ''));
    if (!produto) continue;

    for (var u = 0; u < unidades.length; u++) {
      var qtd = vals[r][unidades[u].col];
      if (!pedidoTemQuantidadeRelatorio_(qtd)) continue;
      linhas.push([unidades[u].nome, produto, '', qtd, '', '']);
    }
  }
  return linhas;
}

/**
 * @param {Spreadsheet} ssMae
 * @param {Array.<Array.<*>>} linhasCons
 * @return {File}
 */
function gerarRelatorioExcel_(ssMae, linhasCons, opcoes) {
  opcoes = opcoes || {};
  var dados = agruparDadosRelatorioPorPregao_(linhasCons, ssMae);
  if (!dados.pregoes.length) {
    throw new Error(
      'Nao ha produtos com quantidade pedida para montar o relatorio por pregao.'
    );
  }

  var agora = new Date();
  var tz = Session.getScriptTimeZone();
  var dataArquivo = Utilities.formatDate(agora, tz, 'yyyy-MM-dd_HH-mm');
  var nomeBase = opcoes.nomeBase || 'Relatorio de Pedidos - ' + dataArquivo;

  var montagem = montarLinhasRelatorioExcel_(dados, agora, tz);
  var temp = SpreadsheetApp.create(nomeBase);
  var aba = temp.getActiveSheet();
  aba.setName('Pedidos por Pregao');

  var nCols = 2;
  var nRows = montagem.linhas.length;
  if (nRows > 0) {
    var linhasNorm = normalizarLinhasRelatorioExcel_(montagem.linhas, nCols);
    aba.getRange(1, 1, nRows, nCols).setValues(linhasNorm);
  }
  aplicarVisualRelatorioExcel_(aba, montagem.meta, nRows, nCols);
  SpreadsheetApp.flush();

  var tempId = temp.getId();
  var blob = exportarPlanilhaComoXlsx_(tempId, nomeBase + '.xlsx');
  var pasta = obterPastaRelatoriosExportacao_(ssMae);
  var arquivo = pasta.createFile(blob);
  DriveApp.getFileById(tempId).setTrashed(true);
  return arquivo;
}

/**
 * Exporta planilha Google como .xlsx (getAs nativo falha no Apps Script).
 * @param {string} spreadsheetId
 * @param {string} nomeArquivo
 * @return {Blob}
 */
function exportarPlanilhaComoXlsx_(spreadsheetId, nomeArquivo) {
  var url =
    'https://docs.google.com/spreadsheets/d/' +
    encodeURIComponent(spreadsheetId) +
    '/export?format=xlsx';
  var resposta = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resposta.getResponseCode() !== 200) {
    throw new Error(
      'Nao foi possivel exportar Excel (HTTP ' + resposta.getResponseCode() + ').'
    );
  }
  return resposta
    .getBlob()
    .setName(nomeArquivo)
    .setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

/**
 * Agrupa pedidos: pregao -> produtos com quantidade -> unidades.
 * Produtos do mesmo pregao sem pedido nao entram.
 * @param {Array.<Array.<*>>} linhasCons
 * @param {Spreadsheet=} ssOrigem
 * @return {{pregoes:Array.<Object>, totalPregoes:number, totalProdutos:number, totalUnidades:number}}
 */
function agruparDadosRelatorioPorPregao_(linhasCons, ssOrigem) {
  var catalogo = {};
  try {
    var ssCatalogo = ssOrigem;
    if (!ssCatalogo || !ssCatalogo.getSheetByName(CONFIG.ABAS.PRODUTOS)) {
      ssCatalogo = obterPlanilhaMae();
    }
    catalogo = montarMapaCatalogoProdutoPregao_(ssCatalogo);
    catalogo = enriquecerCatalogoChavesFlexRelatorio_(catalogo);
  } catch (eCat) {}

  var mapa = {};
  var ordem = [];
  var unidadesVistas = {};

  for (var i = 0; i < (linhasCons || []).length; i++) {
    var row = linhasCons[i];
    var nomeUnidade = String(row[0] || '').trim();
    if (!nomeUnidade) continue;
    if (!pedidoTemQuantidadeRelatorio_(row[3])) continue;

    var produto = padronizarNomenclaturaProdutoResumo_(row[1]);
    if (!produto) continue;

    var chaveProd = normalizarTexto(produto);
    var chaveFlex =
      typeof chaveProdutoFlex_ === 'function' ? chaveProdutoFlex_(produto) : chaveProd;
    var info = catalogo[chaveProd] || catalogo[chaveFlex] || {};
    var pregaoTexto = String(info.pregao || '').trim() || 'SEM PREGAO CADASTRADO';
    var chavePregao = chaveAgrupamentoPregaoRelatorio_(pregaoTexto);

    if (!mapa[chavePregao]) {
      mapa[chavePregao] = {
        nome: pregaoTexto,
        produtosMap: {},
        ordemProdutos: []
      };
      ordem.push(chavePregao);
    }

    var bloco = mapa[chavePregao];
    var chaveItem = chaveFlex || chaveProd;
    if (!bloco.produtosMap[chaveItem]) {
      bloco.produtosMap[chaveItem] = {
        nome: String(produto || '').toUpperCase(),
        unidadesMap: {},
        ordemUnidades: []
      };
      bloco.ordemProdutos.push(chaveItem);
    }

    var item = bloco.produtosMap[chaveItem];
    var chaveUn = normalizarTexto(nomeUnidade);
    unidadesVistas[chaveUn] = true;
    if (!item.unidadesMap[chaveUn]) {
      item.unidadesMap[chaveUn] = {
        nome: nomeUnidade,
        quantidade: 0,
        texto: String(row[3] === 0 || row[3] === '0' ? '0' : row[3] || '').trim()
      };
      item.ordemUnidades.push(chaveUn);
    }
    var n = extrairNumeroQuantidade(row[3]);
    if (n !== null) item.unidadesMap[chaveUn].quantidade += n;
    else item.unidadesMap[chaveUn].texto = String(row[3] || '').trim();
  }

  ordem.sort(function (a, b) {
    if (typeof compararPregoesOrdem_ === 'function') {
      return compararPregoesOrdem_(mapa[a].nome, mapa[b].nome);
    }
    return String(mapa[a].nome || '').localeCompare(String(mapa[b].nome || ''), 'pt-BR');
  });

  var pregoes = [];
  var totalProdutos = 0;
  for (var p = 0; p < ordem.length; p++) {
    var bl = mapa[ordem[p]];
    bl.ordemProdutos.sort(function (a, b) {
      return String(bl.produtosMap[a].nome || '').localeCompare(
        String(bl.produtosMap[b].nome || ''),
        'pt-BR'
      );
    });
    var produtos = [];
    for (var j = 0; j < bl.ordemProdutos.length; j++) {
      var it = bl.produtosMap[bl.ordemProdutos[j]];
      it.ordemUnidades.sort(function (ua, ub) {
        return String(it.unidadesMap[ua].nome || '').localeCompare(
          String(it.unidadesMap[ub].nome || ''),
          'pt-BR'
        );
      });
      var unidades = [];
      var soma = 0;
      for (var u = 0; u < it.ordemUnidades.length; u++) {
        var un = it.unidadesMap[it.ordemUnidades[u]];
        var qtdExibir =
          un.quantidade > 0 ? String(un.quantidade) : un.texto;
        unidades.push({ nome: un.nome, quantidade: qtdExibir });
        soma += un.quantidade || 0;
      }
      if (!unidades.length) continue;
      produtos.push({ nome: it.nome, unidades: unidades, total: soma });
    }
    if (!produtos.length) continue;
    totalProdutos += produtos.length;
    pregoes.push({ nome: bl.nome, produtos: produtos });
  }

  return {
    pregoes: pregoes,
    totalPregoes: pregoes.length,
    totalProdutos: totalProdutos,
    totalUnidades: Object.keys(unidadesVistas).length
  };
}

/**
 * @param {Object.<string, Object>} catalogo
 * @return {Object.<string, Object>}
 */
function enriquecerCatalogoChavesFlexRelatorio_(catalogo) {
  catalogo = catalogo || {};
  var chaves = Object.keys(catalogo);
  for (var i = 0; i < chaves.length; i++) {
    var info = catalogo[chaves[i]];
    if (!info || !info.produto) continue;
    var flex =
      typeof chaveProdutoFlex_ === 'function'
        ? chaveProdutoFlex_(info.produto)
        : '';
    if (flex && !catalogo[flex]) catalogo[flex] = info;
  }
  return catalogo;
}

/**
 * @param {string} texto
 * @return {string}
 */
function chaveAgrupamentoPregaoRelatorio_(texto) {
  try {
    if (typeof normalizarChavePregaoRelatorio_ === 'function') {
      return normalizarChavePregaoRelatorio_(texto);
    }
  } catch (e) {}
  return normalizarTexto(texto);
}

/**
 * @param {Object} dados
 * @param {Date} agora
 * @param {string} tz
 * @return {{linhas: Array.<Array.<string>>, meta: Array.<Object>}}
 */
function montarLinhasRelatorioExcel_(dados, agora, tz) {
  var linhas = [];
  var meta = [];
  var dataTexto = Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm');

  linhas.push(['RELATORIO DE PEDIDOS POR PREGAO - ASSISTENCIA SOCIAL', '']);
  meta.push({ tipo: 'titulo' });
  linhas.push([
    'Gerado em ' +
      dataTexto +
      '   |   ' +
      dados.totalPregoes +
      ' pregoes   |   ' +
      dados.totalProdutos +
      ' produtos   |   ' +
      dados.totalUnidades +
      ' unidades',
    ''
  ]);
  meta.push({ tipo: 'info' });
  linhas.push(['', '']);
  meta.push({ tipo: 'vazio' });

  var pregoes = dados.pregoes || [];
  for (var p = 0; p < pregoes.length; p++) {
    var bloco = pregoes[p];
    linhas.push(['(PREGAO):', String(bloco.nome || '').trim()]);
    meta.push({ tipo: 'pregao' });
    linhas.push(['', '']);
    meta.push({ tipo: 'vazio' });

    var produtos = bloco.produtos || [];
    for (var i = 0; i < produtos.length; i++) {
      var item = produtos[i];
      linhas.push(['PRODUTO:', String(item.nome || '').trim()]);
      meta.push({ tipo: 'produto' });
      linhas.push(['UNIDADES:', 'QTE']);
      meta.push({ tipo: 'rotuloUnidades' });

      var unidades = item.unidades || [];
      for (var u = 0; u < unidades.length; u++) {
        linhas.push([
          String(unidades[u].nome || '').trim(),
          String(unidades[u].quantidade || '').trim()
        ]);
        meta.push({ tipo: 'unidade' });
      }

      linhas.push(['TOTAL:', item.total ? String(item.total) : '']);
      meta.push({ tipo: 'total' });

      if (i < produtos.length - 1) {
        linhas.push(['', '']);
        meta.push({ tipo: 'vazio' });
      }
    }

    if (p < pregoes.length - 1) {
      linhas.push(['', '']);
      meta.push({ tipo: 'separador' });
    }
  }

  return { linhas: linhas, meta: meta };
}

/**
 * Garante que cada linha tenha exatamente nCols colunas (exigencia do setValues).
 * @param {Array.<Array.<*>>} linhas
 * @param {number} nCols
 * @return {Array.<Array.<*>>}
 */
function normalizarLinhasRelatorioExcel_(linhas, nCols) {
  var resultado = [];
  for (var i = 0; i < linhas.length; i++) {
    var row = linhas[i] || [];
    var out = [];
    for (var c = 0; c < nCols; c++) {
      out.push(row[c] !== undefined && row[c] !== null ? row[c] : '');
    }
    resultado.push(out);
  }
  return resultado;
}

/**
 * Intervalo de uma linha (A1).
 * @param {Sheet} aba
 * @param {number} linha
 * @param {number} nCols
 * @return {Range}
 */
function intervaloLinhaRelatorioExcel_(aba, linha, nCols) {
  var colLetra = columnToLetterRelatorioExcel_(nCols);
  return aba.getRange('A' + linha + ':' + colLetra + linha);
}

/**
 * @param {number} col
 * @return {string}
 */
function columnToLetterRelatorioExcel_(col) {
  var temp = '';
  var n = col;
  while (n > 0) {
    var rem = (n - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    n = Math.floor((n - 1) / 26);
  }
  return temp;
}

/**
 * @param {Sheet} aba
 * @param {Array.<Object>} meta
 * @param {number} nRows
 * @param {number} nCols
 */
function aplicarVisualRelatorioExcel_(aba, meta, nRows, nCols) {
  aba.setHiddenGridlines(true);
  aba.setColumnWidth(1, 260);
  aba.setColumnWidth(2, 720);
  aba.setColumnWidth(3, 24);

  var maxR = Math.max(nRows + 8, 40);
  aba.getRange(1, 1, maxR, 4).setBackground('#F8FAFC');

  var refsPregaoA = [];
  var refsPregaoB = [];
  var refsProdutoA = [];
  var refsProdutoB = [];
  var refsRotulo = [];
  var refsUnidade = [];
  var refsTotal = [];
  var refsVazio = [];
  var linhasRotulo = [];
  var linhasUnidade = [];
  var linhasTotal = [];

  for (var r = 0; r < meta.length; r++) {
    var linha = r + 1;
    var tipo = meta[r].tipo;

    if (tipo === 'titulo') {
      try {
        aba.getRange(linha, 1, 1, nCols).merge();
      } catch (eT) {}
      aba
        .getRange(linha, 1, 1, nCols)
        .setBackground('#1D4ED8')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setFontSize(16)
        .setFontFamily('Arial')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      aba.setRowHeight(linha, 44);
      continue;
    }

    if (tipo === 'info') {
      try {
        aba.getRange(linha, 1, 1, nCols).merge();
      } catch (eI) {}
      aba
        .getRange(linha, 1, 1, nCols)
        .setBackground('#DBEAFE')
        .setFontColor('#1E3A8A')
        .setFontSize(11)
        .setFontFamily('Arial')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      aba.setRowHeight(linha, 28);
      continue;
    }

    if (tipo === 'pregao') {
      refsPregaoA.push('A' + linha);
      refsPregaoB.push('B' + linha);
      aba.setRowHeight(linha, 36);
      continue;
    }

    if (tipo === 'produto') {
      refsProdutoA.push('A' + linha);
      refsProdutoB.push('B' + linha);
      aba.setRowHeight(linha, 30);
      continue;
    }

    if (tipo === 'rotuloUnidades') {
      refsRotulo.push('A' + linha + ':B' + linha);
      linhasRotulo.push(linha);
      aba.setRowHeight(linha, 24);
      continue;
    }

    if (tipo === 'unidade') {
      refsUnidade.push('A' + linha + ':B' + linha);
      linhasUnidade.push(linha);
      aba.setRowHeight(linha, 24);
      continue;
    }

    if (tipo === 'total') {
      refsTotal.push('A' + linha + ':B' + linha);
      linhasTotal.push(linha);
      aba.setRowHeight(linha, 26);
      continue;
    }

    if (tipo === 'separador') {
      aba.setRowHeight(linha, 16);
      aba.getRange(linha, 1, 1, nCols).setBackground('#E2E8F0');
      continue;
    }

    refsVazio.push('A' + linha + ':B' + linha);
    aba.setRowHeight(linha, 10);
  }

  if (refsPregaoA.length) {
    aba
      .getRangeList(refsPregaoA)
      .setBackground('#1D4ED8')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(12)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    aba
      .getRangeList(refsPregaoB)
      .setBackground('#EFF6FF')
      .setFontColor('#1E3A8A')
      .setFontWeight('bold')
      .setFontSize(11)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);
  }

  if (refsProdutoA.length) {
    aba
      .getRangeList(refsProdutoA)
      .setBackground('#F59E0B')
      .setFontColor('#78350F')
      .setFontWeight('bold')
      .setFontSize(11)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    aba
      .getRangeList(refsProdutoB)
      .setBackground('#FEF3C7')
      .setFontColor('#78350F')
      .setFontWeight('bold')
      .setFontSize(11)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);
  }

  if (refsRotulo.length) {
    aba
      .getRangeList(refsRotulo)
      .setBackground('#E2E8F0')
      .setFontColor('#334155')
      .setFontWeight('bold')
      .setFontSize(9)
      .setFontFamily('Arial');
    for (var ru = 0; ru < linhasRotulo.length; ru++) {
      aba.getRange(linhasRotulo[ru], 2).setHorizontalAlignment('center');
    }
  }

  if (refsUnidade.length) {
    aba
      .getRangeList(refsUnidade)
      .setBackground('#FFFFFF')
      .setFontColor('#0F172A')
      .setFontSize(10)
      .setFontFamily('Arial');
    for (var un = 0; un < linhasUnidade.length; un++) {
      aba.getRange(linhasUnidade[un], 1).setFontWeight('bold');
      aba
        .getRange(linhasUnidade[un], 2)
        .setHorizontalAlignment('center')
        .setFontWeight('bold');
    }
  }

  if (refsTotal.length) {
    aba
      .getRangeList(refsTotal)
      .setBackground('#DCFCE7')
      .setFontColor('#14532D')
      .setFontWeight('bold')
      .setFontSize(10)
      .setFontFamily('Arial');
    for (var tot = 0; tot < linhasTotal.length; tot++) {
      aba.getRange(linhasTotal[tot], 2).setHorizontalAlignment('center');
    }
  }

  if (refsVazio.length) {
    aba.getRangeList(refsVazio).setBackground('#F8FAFC');
  }

  aba.setFrozenRows(2);
}

/**
 * Pasta no Drive para relatorios exportados (Excel).
 * @param {Spreadsheet} ss
 * @return {Folder}
 */
function obterPastaRelatoriosExportacao_(ss) {
  var nome = 'Relatórios - Sistema Prefeitura Compras';
  var arquivoMae = DriveApp.getFileById(ss.getId());
  var pais = arquivoMae.getParents();
  var pai = pais.hasNext() ? pais.next() : null;
  var pastas = pai ? pai.getFoldersByName(nome) : DriveApp.getFoldersByName(nome);
  if (pastas.hasNext()) return pastas.next();
  return pai ? pai.createFolder(nome) : DriveApp.createFolder(nome);
}
