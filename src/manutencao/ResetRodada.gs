/**
 * Nova rodada de compras — zera pedidos na Mãe e em todas as unidades.
 * Mantém: catálogo Produtos, aba Unidades (IDs), lista de produtos nas filhas.
 */

/**
 * Menu: apaga pedidos da rodada anterior (Mãe + unidades).
 */
function menuResetNovaRodadaCompras() {
  var ui = SpreadsheetApp.getUi();

  var r1 = ui.alert(
    'Nova rodada de compras',
    'Isso APAGA os pedidos da rodada anterior:\n\n' +
      '• Mãe: Resumo, Pedidos das Unidades e Consolidados\n' +
      '• Todas as unidades: quantidades, observações e Status → Rascunho\n' +
      '• Nas unidades: faixa “PLANILHA DE PEDIDOS — MÊS XX” atualizada para o mês atual\n' +
      '• Aba Produtos: remove cadastros de teste (TESTE, www, yyyy, etc.)\n' +
      '• Produtos reais com Ativo = Não PERMANECEM no cadastro\n\n' +
      'Mantém: aba Unidades (IDs/links) e o catálogo real (ativos e inativos).\n\n' +
      'Use quando for começar um novo mês/ciclo de pedidos.\n\n' +
      'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );
  if (r1 !== ui.Button.YES) return;

  var r2 = ui.alert(
    'Confirmação final',
    'Última chance: todos os pedidos Enviados serão limpos.\n\n' +
      'Recomendado: faça uma cópia da Planilha Mãe no Drive antes, se quiser guardar histórico.\n\n' +
      'Apagar mesmo assim?',
    ui.ButtonSet.YES_NO
  );
  if (r2 !== ui.Button.YES) return;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Limpando pedidos da mãe e das unidades…', 'Nova rodada', 90);
    var resultado = executarResetNovaRodadaCompras_();
    SpreadsheetApp.flush();
    if (resultado.continuando) {
      ui.alert(
        'Nova rodada em andamento',
        'A planilha mãe já foi limpa.\n' +
          'Unidades limpas agora: ' +
          resultado.unidadesOk +
          '\nAinda faltam: ' +
          resultado.faltando +
          '\n\nO sistema continua sozinho nas unidades que faltam.\n' +
          'O limite de 6 minutos do Google não interrompe a rodada.',
        ui.ButtonSet.OK
      );
      return;
    }
    ui.alert(
      'Nova rodada pronta',
      'Mãe: Resumo, Pedidos das Unidades e Consolidados limpos.\n' +
        'Produtos de teste removidos: ' +
        (resultado.testeRemovidos || 0) +
        '\n' +
        'Produtos reais com Ativo = Não foram mantidos.\n' +
        'Unidades resetadas: ' +
        resultado.unidadesOk +
        '\nFalhas: ' +
        resultado.falhas.length +
        (resultado.falhas.length
          ? '\n\n- ' + resultado.falhas.slice(0, 6).join('\n- ')
          : '') +
        '\n\nAs unidades já podem preencher e enviar de novo.\n' +
        'A faixa de mês de referência foi atualizada em cada unidade.\n' +
        'Depois: unidades enviam → Mãe usa “Puxar pedidos das unidades”.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro no reset', e.message, ui.ButtonSet.OK);
  }
}

/**
 * @return {{unidadesOk:number, falhas:string[], testeRemovidos:number}}
 */
function executarResetNovaRodadaCompras_() {
  var lock = adquirirLock('reset_rodada', 120);
  var falhas = [];
  var unidadesOk = 0;
  var testeRemovidos = 0;

  try {
    var ssMae = obterPlanilhaMae();

    // 1) Mãe primeiro — só zera pedidos. Não remonta o Resumo nem a aba Produtos.
    try {
      escreverSomenteValoresSemFormatar_(
        ssMae,
        CONFIG.ABAS.CONSOLIDADOS,
        CONFIG.CONSOLIDADOS.HEADER,
        []
      );
    } catch (eCons) {
      falhas.push('Consolidados: ' + eCons.message);
    }
    try {
      zerarQuantidadesAbaResumo_(ssMae);
    } catch (eRes) {
      falhas.push('Resumo: ' + eRes.message);
    }
    try {
      limparTotalPorProdutoNovaRodada_(ssMae);
    } catch (eTot) {
      falhas.push('Pedidos das Unidades: ' + eTot.message);
    }
    try {
      atualizarIndicadorUltimaAtualizacao_(ssMae);
    } catch (eInd) {}

    // 2) Unidades: só quantidade, observação e Status. A lista de produtos fica.
    var unidades = [];
    try {
      unidades = listarUnidadesAtivas(ssMae);
    } catch (eLista) {
      falhas.push('Lista de unidades: ' + eLista.message);
    }
    if (unidades.length === 0 && DADOS_INICIAIS && DADOS_INICIAIS.unidades) {
      for (var d = 0; d < DADOS_INICIAIS.unidades.length; d++) {
        unidades.push({
          nome: DADOS_INICIAIS.unidades[d],
          spreadsheetId: DADOS_INICIAIS.idsPlanilhas[DADOS_INICIAIS.unidades[d]] || '',
          ativo: true
        });
      }
    }

    var inicio = Date.now();
    var LIMITE_MS = 4 * 60 * 1000;
    var faltando = [];
    for (var i = 0; i < unidades.length; i++) {
      if (Date.now() - inicio > LIMITE_MS) {
        faltando = unidades.slice(i);
        break;
      }
      var u = unidades[i];
      try {
        try {
          SpreadsheetApp.getActiveSpreadsheet().toast(
            'Limpando pedidos ' + (i + 1) + '/' + unidades.length + ': ' + u.nome,
            'Nova rodada',
            8
          );
        } catch (eToast) {}
        var id = String(u.spreadsheetId || '').trim();
        if (!id) throw new Error('ID não encontrado');
        resetarPedidosPlanilhaUnidade_(SpreadsheetApp.openById(id));
        unidadesOk++;
      } catch (e) {
        falhas.push(u.nome + ': ' + e.message);
      }
    }

    if (faltando.length) {
      agendarContinuacaoResetNovaRodada_(faltando, falhas, unidadesOk);
      registrarLog(
        'RESET_RODADA',
        'mãe limpa; continua em ' + faltando.length + ' unidade(s)'
      );
      return {
        unidadesOk: unidadesOk,
        falhas: falhas,
        testeRemovidos: 0,
        continuando: true,
        faltando: faltando.length
      };
    }

    // 3) Testes por último, sem reformatação pesada da aba Produtos.
    try {
      testeRemovidos = removerProdutosTesteAbaProdutos_(ssMae, false, true);
    } catch (eTeste) {
      falhas.push('Limpeza de testes na Mãe: ' + eTeste.message);
    }

    var msg =
      'Reset rodada — resumo zerado, unidades ok: ' +
      unidadesOk +
      ', falhas: ' +
      falhas.length +
      ' — ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    registrarLog('RESET_RODADA', msg);

    try {
      limparPropsSyncPendente_();
      removerGatilhosClockPorFuncao_('processarTotalPorProdutoPendente_');
    } catch (eGat) {}

    return {
      unidadesOk: unidadesOk,
      falhas: falhas,
      testeRemovidos: testeRemovidos,
      continuando: false,
      faltando: 0
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Guarda as unidades que faltam e agenda a continuação em alguns segundos.
 * @param {Array.<Object>} unidades
 * @param {string[]} falhas
 * @param {number} unidadesOk
 */
function agendarContinuacaoResetNovaRodada_(unidades, falhas, unidadesOk) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(
    'RESET_RODADA_PENDENTE',
    JSON.stringify({
      unidades: (unidades || []).map(function (u) {
        return { nome: u.nome, spreadsheetId: u.spreadsheetId };
      }),
      falhas: falhas || [],
      unidadesOk: unidadesOk || 0
    })
  );
  try {
    removerGatilhosClockPorFuncao_('continuarResetNovaRodada_');
    ScriptApp.newTrigger('continuarResetNovaRodada_').timeBased().after(8000).create();
  } catch (eTrig) {
    registrarLog('RESET_RODADA_AGENDAR', eTrig.message);
  }
}

/**
 * Continua a limpeza das unidades se o primeiro passo chegou perto dos 6 minutos.
 */
function continuarResetNovaRodada_() {
  var lock = adquirirLock('reset_rodada', 120);
  var props = PropertiesService.getScriptProperties();
  try {
    try {
      removerGatilhosClockPorFuncao_('continuarResetNovaRodada_');
    } catch (eOld) {}
    var bruto = props.getProperty('RESET_RODADA_PENDENTE');
    if (!bruto) return;
    var estado = JSON.parse(bruto);
    var unidades = estado.unidades || [];
    var falhas = estado.falhas || [];
    var unidadesOk = estado.unidadesOk || 0;
    var inicio = Date.now();
    var LIMITE_MS = 4 * 60 * 1000;
    var faltando = [];

    for (var i = 0; i < unidades.length; i++) {
      if (Date.now() - inicio > LIMITE_MS) {
        faltando = unidades.slice(i);
        break;
      }
      var u = unidades[i];
      try {
        var id = String(u.spreadsheetId || '').trim();
        if (!id) throw new Error('ID não encontrado');
        resetarPedidosPlanilhaUnidade_(SpreadsheetApp.openById(id));
        unidadesOk++;
      } catch (e) {
        falhas.push(u.nome + ': ' + e.message);
      }
    }

    if (faltando.length) {
      agendarContinuacaoResetNovaRodada_(faltando, falhas, unidadesOk);
      registrarLog('RESET_RODADA', 'continua: faltam ' + faltando.length);
      return;
    }

    try {
      removerProdutosTesteAbaProdutos_(obterPlanilhaMae(), false, true);
    } catch (eTeste) {
      falhas.push('Limpeza de testes na Mãe: ' + eTeste.message);
    }
    props.deleteProperty('RESET_RODADA_PENDENTE');
    registrarLog(
      'RESET_RODADA',
      'concluída após continuação — unidades ok: ' +
        unidadesOk +
        ', falhas: ' +
        falhas.length
    );
  } finally {
    lock.releaseLock();
  }
}

/**
 * Apaga só as quantidades do Resumo. Gênero, pregão e produto permanecem.
 * @param {Spreadsheet} ssMae
 */
function zerarQuantidadesAbaResumo_(ssMae) {
  var aba = ssMae.getSheetByName(CONFIG.ABAS.RESUMO);
  if (!aba || aba.getLastRow() < 2 || aba.getLastColumn() < 2) return;

  var lastR = Math.min(aba.getLastRow(), 1500);
  var lastC = Math.min(aba.getLastColumn(), 40);
  var corpo = aba.getRange(2, 1, lastR - 1, lastC).getValues();
  var blocoIni = -1;

  function gravarBloco(de, ate) {
    if (de < 0 || ate < de) return;
    var n = ate - de + 1;
    var vazias = [];
    for (var i = 0; i < n; i++) {
      var linha = [];
      for (var c = 1; c < lastC; c++) linha.push('');
      vazias.push(linha);
    }
    aba.getRange(de + 2, 2, n, lastC - 1).setValues(vazias);
  }

  for (var r = 0; r < corpo.length; r++) {
    var rotulo = String(corpo[r][0] || '');
    if (rotulo.indexOf('   ') === 0) {
      if (blocoIni < 0) blocoIni = r;
    } else {
      gravarBloco(blocoIni, r - 1);
      blocoIni = -1;
    }
  }
  gravarBloco(blocoIni, corpo.length - 1);
}

/**
 * Função antiga: apagava da aba Produtos as linhas com Ativo = Não.
 * A Nova rodada NÃO chama mais esta função — inativos permanecem no cadastro.
 * @return {number} quantidade removida
 */
function removerProdutosInativosAbaProdutos_() {
  var ssMae = obterPlanilhaMae();
  var aba = obterAbaObrigatoria(ssMae, CONFIG.ABAS.PRODUTOS);
  var map = obterMapaColunasProdutos_(aba);
  var lastRow = aba.getLastRow();
  var lastCol = Math.max(aba.getLastColumn(), map.colAtivo + 1, CONFIG.PRODUTOS.HEADER.length, 6);

  if (lastRow < CONFIG.LINHA_DADOS) {
    return 0;
  }

  var numLinhas = lastRow - CONFIG.LINHA_DADOS + 1;
  var vals = aba.getRange(CONFIG.LINHA_DADOS, 1, numLinhas, lastCol).getValues();
  var mantidas = [];
  var removidas = 0;

  for (var i = 0; i < vals.length; i++) {
    var row = vals[i];
    var produto = String(row[map.colProduto] || '').trim();
    var codigo = String(row[map.colCodigo] || '').trim();
    if (!produto && !codigo) {
      continue;
    }
    if (ehAtivo(row[map.colAtivo])) {
      mantidas.push(row);
    } else {
      removidas++;
    }
  }

  try {
    aba.getRange(CONFIG.LINHA_DADOS, 1, numLinhas, lastCol).breakApart();
  } catch (eBreak) {}
  aba.getRange(CONFIG.LINHA_DADOS, 1, numLinhas, lastCol).clearContent();
  if (mantidas.length > 0) {
    aba.getRange(CONFIG.LINHA_DADOS, 1, mantidas.length, lastCol).setValues(mantidas);
  }

  try {
    aplicarValidacaoAtivoProdutos_(aba);
  } catch (eVal) {}
  try {
    organizarProdutosPorCategoria_(ssMae);
  } catch (eVis) {}

  return removidas;
}

/**
 * Zera pedidos na filha; mantém lista de produtos e dropdowns.
 * @param {Spreadsheet} ssFilha
 * @param {Array.<Object>} produtosCatalogo produtos ativos, com categoria
 */
function resetarPedidosPlanilhaUnidade_(ssFilha, produtosCatalogo) {
  var aba = obterAbaPedidoFilha_(ssFilha);
  var dataMes = new Date();
  try {
    aplicarMesReferenciaPedidoUnidade_(aba, dataMes);
  } catch (eMes) {}

  var lastRow = aba.getLastRow();
  var lastCol = Math.max(aba.getLastColumn(), obterNumColsPedido_(), 6);
  if (lastRow < 2) return;

  var vals = aba.getRange(1, 1, lastRow, lastCol).getValues();
  var displays = vals;
  try {
    displays = aba.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  } catch (eD) {}
  var mapa = detectarColunasPedido_(displays);
  if (mapa.colQtd < 0 || mapa.colStatus < 0 || mapa.colProduto < 0) {
    mapa = detectarColunasPedido_(vals);
  }
  if (mapa.colQtd < 0 || mapa.colStatus < 0) {
    throw new Error('não achei as colunas Quantidade/Status');
  }

  var rascunho = CONFIG.VALORES.STATUS_RASCUNHO;
  var nLinhas = lastRow - (mapa.headerRow + 1);
  if (nLinhas < 1) return;

  var colQtd = [];
  var colObs = mapa.colObs >= 0 ? [] : null;
  var colSt = [];
  for (var i = mapa.headerRow + 1; i < vals.length; i++) {
    var produto = String((displays[i] && displays[i][mapa.colProduto]) || vals[i][mapa.colProduto] || '').trim();
    if (!produto || ehLinhaCategoriaPedido_(produto)) {
      colQtd.push([vals[i][mapa.colQtd]]);
      if (colObs) colObs.push([vals[i][mapa.colObs]]);
      colSt.push([vals[i][mapa.colStatus]]);
      continue;
    }
    colQtd.push(['']);
    if (colObs) colObs.push(['']);
    colSt.push([rascunho]);
  }

  aba.getRange(mapa.headerRow + 2, mapa.colQtd + 1, colQtd.length, 1).setValues(colQtd);
  if (colObs) {
    aba.getRange(mapa.headerRow + 2, mapa.colObs + 1, colObs.length, 1).setValues(colObs);
  }
  aba.getRange(mapa.headerRow + 2, mapa.colStatus + 1, colSt.length, 1).setValues(colSt);
}

/**
 * Menu: apaga só cadastros de teste (Mãe + unidades), sem zerar pedidos reais.
 */
function menuLimparProdutosTeste() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.alert(
    'Limpar produtos de teste',
    'Isso APAGA da aba Produtos (e das unidades) cadastros de teste, como:\n\n' +
      '• TESTE, TESTE 01, TESTETSTETS\n' +
      '• www, yyyy\n' +
      '• texto batido no teclado (jhjhjh, asdf…)\n\n' +
      'Produtos reais (BOTA PVC, alimentos, etc.) permanecem.\n' +
      'Pedidos das unidades nos itens reais NÃO são apagados.\n\n' +
      'Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (r1 !== ui.Button.YES) return;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Removendo produtos de teste…', 'Limpeza', 90);
    var resultado = executarLimpezaProdutosTeste_();
    SpreadsheetApp.flush();
    ui.alert(
      'Limpeza concluída',
      'Produtos de teste removidos da Mãe: ' +
        resultado.removidos +
        (resultado.amostra.length
          ? '\n\nExemplos:\n- ' + resultado.amostra.slice(0, 8).join('\n- ')
          : '') +
        '\n\nUnidades atualizadas: ' +
        resultado.unidadesOk +
        '\nFalhas: ' +
        resultado.falhas.length +
        (resultado.falhas.length
          ? '\n\n- ' + resultado.falhas.slice(0, 6).join('\n- ')
          : '') +
        '\n\nSe for começar um ciclo novo, use também “Nova rodada — limpar pedidos”.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro na limpeza', e.message, ui.ButtonSet.OK);
  }
}

/**
 * @return {{removidos:number, amostra:string[], unidadesOk:number, falhas:string[]}}
 */
function executarLimpezaProdutosTeste_() {
  var lock = adquirirLock('limpeza_teste', 120);
  var falhas = [];
  try {
    resolverIdsUnidadesPeloDrive();
    var ssMae = obterPlanilhaMae();
    var removidosInfo = removerProdutosTesteAbaProdutos_(ssMae, true);
    try {
      sincronizarCadastroCategorias_();
    } catch (eCat) {
      falhas.push('Categorias: ' + eCat.message);
    }
    try {
      limparConsolidadosProdutosTeste_(ssMae);
    } catch (eCons) {
      falhas.push('Consolidados: ' + eCons.message);
    }

    var unidadesOk = 0;
    try {
      var sync = sincronizarCatalogoTodasUnidades(true, true, {
        modoRapido: true,
        mostrarProgresso: true
      });
      unidadesOk = sync.sucesso || 0;
      if (sync.falhas && sync.falhas.length) {
        falhas = falhas.concat(sync.falhas);
      }
    } catch (eSync) {
      falhas.push('Unidades: ' + eSync.message);
    }

    registrarLog(
      'LIMPEZA_TESTE',
      'removidos=' +
        removidosInfo.removidos +
        ' unidades=' +
        unidadesOk +
        ' falhas=' +
        falhas.length
    );
    return {
      removidos: removidosInfo.removidos,
      amostra: removidosInfo.amostra || [],
      unidadesOk: unidadesOk,
      falhas: falhas
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Remove da aba Produtos linhas de teste (TESTE, www, yyyy, teclado batido).
 * @param {Spreadsheet=} ssMae
 * @param {boolean=} detalhado
 * @return {number|{removidos:number, amostra:string[]}}
 */
function removerProdutosTesteAbaProdutos_(ssMae, detalhado, pularVisual) {
  ssMae = ssMae || obterPlanilhaMae();
  var aba = obterAbaObrigatoria(ssMae, CONFIG.ABAS.PRODUTOS);
  var map = obterMapaColunasProdutos_(aba);
  var lastRow = aba.getLastRow();
  var lastCol = Math.max(aba.getLastColumn(), map.colAtivo + 1, CONFIG.PRODUTOS.HEADER.length, 6);
  var amostra = [];
  var removidas = 0;

  if (lastRow < CONFIG.LINHA_DADOS) {
    return detalhado ? { removidos: 0, amostra: amostra } : 0;
  }

  var numLinhas = lastRow - CONFIG.LINHA_DADOS + 1;
  var vals = aba.getRange(CONFIG.LINHA_DADOS, 1, numLinhas, lastCol).getValues();
  var mantidas = [];

  for (var i = 0; i < vals.length; i++) {
    var row = vals[i];
    if (ehLinhaTituloCategoriaProdutos_(row)) {
      continue;
    }
    var produto = String(row[map.colProduto] || '').trim();
    var codigo = String(row[map.colCodigo] || '').trim();
    var pregao = String(row[map.colPregao] || '').trim();
    var categoria = String(row[map.colCategoria] || '').trim();
    if (!produto && !codigo) {
      continue;
    }
    if (ehProdutoTesteLixo_(produto, pregao, categoria)) {
      removidas++;
      if (amostra.length < 12) {
        amostra.push(produto || pregao || categoria || '(linha vazia)');
      }
      continue;
    }
    mantidas.push(row);
  }

  try {
    aba.getRange(CONFIG.LINHA_DADOS, 1, numLinhas, lastCol).breakApart();
  } catch (eBreak) {}
  aba.getRange(CONFIG.LINHA_DADOS, 1, numLinhas, lastCol).clearContent();
  if (mantidas.length > 0) {
    aba.getRange(CONFIG.LINHA_DADOS, 1, mantidas.length, lastCol).setValues(mantidas);
  }

  try {
    aplicarValidacaoAtivoProdutos_(aba);
  } catch (eVal) {}
  if (!pularVisual) {
    try {
      organizarProdutosPorCategoria_(ssMae);
    } catch (eVis) {}
  }

  return detalhado ? { removidos: removidas, amostra: amostra } : removidas;
}

/**
 * Tira dos Consolidados / Resumo as linhas cujo produto era teste.
 * @param {Spreadsheet} ssMae
 */
function limparConsolidadosProdutosTeste_(ssMae) {
  var aba = ssMae.getSheetByName(CONFIG.ABAS.CONSOLIDADOS);
  var linhas = aba ? lerDados(aba, CONFIG.CONSOLIDADOS.HEADER.length) : [];
  var filtradas = (linhas || []).filter(function (row) {
    var produto = String(row[CONFIG.CONSOLIDADOS.PRODUTO - 1] || '').trim();
    var unidade = String(row[CONFIG.CONSOLIDADOS.UNIDADE - 1] || '').trim();
    if (!produto && !unidade) return false;
    return !ehTextoLixoTeste_(produto);
  });
  escreverAbaSimples_(ssMae, CONFIG.ABAS.CONSOLIDADOS, CONFIG.CONSOLIDADOS.HEADER, filtradas);
  escreverResumoAgrupadoFromConsolidados_(ssMae, filtradas);
}

/**
 * @param {string} produto
 * @param {string=} pregao
 * @param {string=} categoria
 * @return {boolean}
 */
function ehProdutoTesteLixo_(produto, pregao, categoria) {
  if (ehTextoLixoTeste_(produto) || ehTextoLixoTeste_(categoria)) return true;
  // Pregão com a palavra teste não apaga um produto de verdade.
  // Ex.: MANGUEIRA com "PREGAO TESTE GABRIEL" continua no cadastro.
  var nome = String(produto || '').trim();
  return ehTextoLixoTeste_(pregao) && !nome;
}

/**
 * Detecta cadastro de teste: TESTE, www, yyyy, letras repetidas, teclado batido.
 * Não marca produtos reais (alimentos, papelaria, EPI, etc.).
 * @param {string} texto
 * @return {boolean}
 */
function ehTextoLixoTeste_(texto) {
  var bruto = String(texto || '').trim();
  if (!bruto) return false;
  var n = bruto
    .toLowerCase()
    .replace(/á|à|ã|â/g, 'a')
    .replace(/é|ê/g, 'e')
    .replace(/í/g, 'i')
    .replace(/ó|ô|õ/g, 'o')
    .replace(/ú/g, 'u')
    .replace(/ç/g, 'c');
  var compacto = n.replace(/[^a-z0-9]+/g, '');
  if (!compacto) return false;

  if (compacto.indexOf('teste') !== -1) return true;
  if (compacto.indexOf('www') !== -1) return true;
  if (/^(yyy+|www+|xxx+|zzz+|aaa+|qqq+|asdf+|qwer+|zxcv+)/.test(compacto)) {
    return true;
  }
  if (/^y+$/.test(compacto) || /^w+$/.test(compacto) || /^x+$/.test(compacto)) {
    return true;
  }
  if (/^([a-z])\1{3,}$/.test(compacto)) return true;
  if (/^([a-z]{2})\1{2,}$/.test(compacto)) return true;
  if (/^([a-z]{3})\1+$/.test(compacto) && compacto.length >= 6) return true;

  if (compacto.length >= 4 && !/[0-9]/.test(compacto)) {
    var unicos = {};
    for (var i = 0; i < compacto.length; i++) {
      unicos[compacto.charAt(i)] = true;
    }
    var qtd = Object.keys(unicos).length;
    if (qtd === 1) return true;
    if (qtd === 2 && compacto.length >= 6) return true;
  }

  return false;
}
