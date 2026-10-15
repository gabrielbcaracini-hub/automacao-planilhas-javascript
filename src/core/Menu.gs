/**
 * Menu da Planilha Mãe — só o que o cliente usa no dia a dia.
 * Funções técnicas ficam no script (não aparecem no menu):
 * menuAjustarVisualProdutos, menuAtualizarControleCategorias,
 * menuAbrirTotalPorProduto, menuMigrarLayoutMesReferenciaUnidades,
 * menuInstalarMenuUnidades, menuLimparProdutosTeste.
 */


/**
 * Devolve a aba Resumo colorida que foi arquivada ao recriar a planilha.
 * A aba nova ficou sem largura e sem faixas (gênero, pregão, produto).
 */
function restaurarResumoFormatadoSeArquivado_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  if (String(ss.getId()) !== String(CONFIG.SPREADSHEET_MAE_ID || '').trim()) return;

  var docProps = PropertiesService.getDocumentProperties();
  if (docProps.getProperty('RESUMO_VISUAL_RESTAURADO_2209') === '1') return;

  var legado = ss.getSheetByName('Resumo_LEGADO_20260921_225145');
  if (!legado) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (String(sheets[i].getName()).indexOf('Resumo_LEGADO_20260921') === 0) {
        legado = sheets[i];
        break;
      }
    }
  }

  var atual = ss.getSheetByName(CONFIG.ABAS.RESUMO);
  if (!legado || !atual || legado.getSheetId() === atual.getSheetId()) {
    docProps.setProperty('RESUMO_VISUAL_RESTAURADO_2209', '1');
    try {
      PropertiesService.getScriptProperties().setProperty('RESUMO_MODELO_PREGAO_V2', '1');
      docProps.setProperty('MATRIZ_SETEMBRO_2026', 'OK');
    } catch (eFlag) {}
    return;
  }

  var largura = 0;
  try {
    largura = atual.getColumnWidth(1);
  } catch (eL) {}
  if (largura >= 400) {
    docProps.setProperty('RESUMO_VISUAL_RESTAURADO_2209', '1');
    try {
      PropertiesService.getScriptProperties().setProperty('RESUMO_MODELO_PREGAO_V2', '1');
      docProps.setProperty('MATRIZ_SETEMBRO_2026', 'OK');
    } catch (eFlag2) {}
    return;
  }

  var idx = atual.getIndex();
  try {
    atual.setName('Resumo_SEM_VISUAL_20260922');
  } catch (eNome) {}
  try {
    atual.hideSheet();
  } catch (eHide) {}
  try {
    legado.showSheet();
  } catch (eShow) {}
  try {
    legado.setName(CONFIG.ABAS.RESUMO);
  } catch (eRen) {}
  try {
    ss.setActiveSheet(legado);
    ss.moveActiveSheet(idx);
  } catch (eMove) {}
  try {
    PropertiesService.getScriptProperties().setProperty('RESUMO_MODELO_PREGAO_V2', '1');
    docProps.setProperty('MATRIZ_SETEMBRO_2026', 'OK');
    docProps.setProperty('RESUMO_VISUAL_RESTAURADO_2209', '1');
  } catch (eProp) {}
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  try {
    if (ehPlanilhaUnidadeAtiva_()) {
      onOpenUnidade_();
      return;
    }
  } catch (eUn) {
    Logger.log('onOpen detecção unidade: ' + eUn.message);
  }

  try {
    ui.createMenu('Sistema de Compras')
      .addItem('➕ Nova linha na aba Produtos', 'menuInserirLinhaNovoProduto')
      .addItem('⚡ Atualizar produtos nas unidades agora', 'menuAtualizarProdutosUnidadesAgora')
      .addSeparator()
      .addItem('📥 Puxar pedidos das unidades', 'menuAtualizarResumoPedidosAgora')
      .addSeparator()
      .addItem('📥 Baixar relatório Excel', 'menuGerarRelatorioExcel')
      .addItem('📥 Relatório de Pregões ativos (Excel)', 'menuGerarRelatorioPregoes')
      .addItem('📥 Relatório de Pregões inativos (Excel)', 'menuGerarRelatorioPregoesInativos')
      .addSeparator()
      .addItem('🗑️ Nova rodada — limpar pedidos', 'menuResetNovaRodadaCompras')
      .addSeparator()
      .addItem('❓ Ajuda', 'menuAjuda')
      .addToUi();
  } catch (eMenu) {
    Logger.log('onOpen menu: ' + eMenu.message);
  }

  try {
    aplicarVisibilidadeAbasCliente_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (eRem) {
    Logger.log('onOpen visibilidade abas: ' + eRem.message);
  }
  try {
    removerGatilhosCatalogoAutomatico_();
  } catch (eG) {
    Logger.log('onOpen limpeza: ' + eG.message);
  }
  try {
    atualizarAbaSituacaoPedidosMae_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (eSit) {
    Logger.log('onOpen situação pedidos: ' + eSit.message);
  }
  try {
    restaurarResumoFormatadoSeArquivado_();
    completarResumoComProdutosDaAba_();
  } catch (eRest) {
    Logger.log('onOpen restaurar resumo: ' + eRest.message);
  }
  try {
    if (typeof importarMatrizSetembro2026SePendente_ === 'function') {
      importarMatrizSetembro2026SePendente_();
    }
  } catch (eMatriz) {
    Logger.log('onOpen importar matriz: ' + eMatriz.message);
  }
}

function menuAjuda() {
  var html = HtmlService.createHtmlOutputFromFile('Ajuda')
    .setWidth(760)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Como usar o sistema');
}

function menuInstalarMenuUnidades() {
  var ui = SpreadsheetApp.getUi();
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Instalando menu nas unidades…', 'Unidades', 60);
    var r = instalarGatilhoOnOpenUnidades_();
    ui.alert(
      'Menu instalado',
      'Planilhas das unidades com menu: ' +
        r.instalados +
        (r.falhas.length ? '\n\nFalhas:\n- ' + r.falhas.slice(0, 6).join('\n- ') : '') +
        '\n\nPeça às unidades que recarreguem a planilha (F5) para ver o menu.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}
