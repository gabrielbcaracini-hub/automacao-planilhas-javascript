/**
 * Setup inicial da Planilha Mãe e das Planilhas Filhas.
 */

/**
 * Prepara a Planilha Mãe: cria abas, cabeçalhos e exemplos iniciais (se vazias).
 */
function setupPlanilhaMae() {
  var ss = obterPlanilhaMae();

  var abaProdutos = garantirAbaComCabecalho(
    ss,
    CONFIG.ABAS.PRODUTOS,
    CONFIG.PRODUTOS.HEADER
  );
  var abaUnidades = garantirAbaComCabecalho(
    ss,
    CONFIG.ABAS.UNIDADES,
    CONFIG.UNIDADES.HEADER
  );
  garantirAbaComCabecalho(ss, CONFIG.ABAS.CONSOLIDADOS, CONFIG.CONSOLIDADOS.HEADER);
  garantirAbaComCabecalho(ss, CONFIG.ABAS.RESUMO, CONFIG.RESUMO.HEADER);
  garantirAbaComCabecalho(ss, CONFIG.ABAS.TOTAL_PRODUTO, CONFIG.TOTAL_PRODUTO.HEADER);
  garantirAbaComCabecalho(ss, CONFIG.ABAS.LOG, CONFIG.LOG.HEADER);

  // Exemplos apenas se a aba Produtos estiver vazia (sem lista UNIDADES ainda)
  if (abaProdutos.getLastRow() < CONFIG.LINHA_DADOS) {
    abaProdutos
      .getRange(CONFIG.LINHA_DADOS, 1, 3, CONFIG.PRODUTOS.HEADER.length)
      .setValues([
        ['001', 'Limpeza', 'Detergente', 'Un', '', CONFIG.VALORES.ATIVO_SIM],
        ['002', 'Escritório', 'Caneta Azul', 'Caixa', '', CONFIG.VALORES.ATIVO_SIM],
        ['003', 'Alimentos', 'Café 500g', 'Pacote', '', CONFIG.VALORES.ATIVO_SIM]
      ]);
  }

  // Validação Sim/Não na coluna Ativo de Produtos e Unidades
  aplicarValidacaoAtivo(abaProdutos, CONFIG.PRODUTOS.ATIVO);
  try {
    aplicarValidacaoAtivoProdutos_(abaProdutos);
  } catch (eVal) {}
  aplicarValidacaoAtivo(abaUnidades, CONFIG.UNIDADES.ATIVO);

  // Remove abas padrão vazias "Folha1" / "Sheet1" se existirem e houver outras abas
  removerAbaPadraoVazia(ss);

  registrarLog('SETUP_MAE', 'Planilha Mãe configurada com sucesso');
  try {
    aplicarVisibilidadeAbasCliente_(ss);
  } catch (eOculta) {}
  SpreadsheetApp.getUi().alert(
    'Setup concluído',
    'Abas da Planilha Mãe criadas/atualizadas.\n\nPróximos passos:\n1. Cadastre as unidades (Nome + Spreadsheet ID)\n2. Ajuste o catálogo em Produtos\n3. Menu → Instalar gatilhos\n4. Menu → Sincronizar catálogo',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Aplica validação Sim/Não em uma coluna.
 * Remove validações antigas da área de dados para evitar conflito
 * (ex.: Ativo mudou da coluna E para F após incluir Pregão).
 * @param {Sheet} aba
 * @param {number} coluna
 */
function aplicarValidacaoAtivo(aba, coluna) {
  var qtdLinhas = CONFIG.PEDIDO_LINHAS_VALIDACAO;
  var qtdCols = Math.max(aba.getLastColumn(), coluna, 10);
  aba.getRange(CONFIG.LINHA_DADOS, 1, qtdLinhas, qtdCols).clearDataValidations();

  var range = aba.getRange(CONFIG.LINHA_DADOS, coluna, qtdLinhas, 1);
  var regra = SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG.VALORES.ATIVO_SIM, CONFIG.VALORES.ATIVO_NAO], true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(regra);
}

/**
 * Remove Sheet1/Folha1 se estiver vazia e não for a única aba.
 * @param {Spreadsheet} ss
 */
function removerAbaPadraoVazia(ss) {
  var nomesPadrao = ['Folha1', 'Sheet1', 'Página1'];
  var abas = ss.getSheets();
  if (abas.length <= 1) {
    return;
  }
  for (var i = 0; i < abas.length; i++) {
    var nome = abas[i].getName();
    if (nomesPadrao.indexOf(nome) !== -1 && abas[i].getLastRow() === 0) {
      try {
        ss.deleteSheet(abas[i]);
      } catch (e) {
        Logger.log('Não removeu aba padrão: ' + e.message);
      }
    }
  }
}

/**
 * Entrada do menu — evita receber o objeto de evento do onOpen/menu como ID.
 */
function menuSetupPlanilhaFilha() {
  setupPlanilhaFilha(null);
}

/**
 * Configura uma planilha filha a partir do ID informado.
 * Cria abas Pedido e Catalogo, dropdown e proteção.
 * @param {string|null} spreadsheetId Opcional; se omitido, solicita via prompt.
 * @param {boolean=} mostrarAlerta
 * @return {string|null} ID configurado
 */
function setupPlanilhaFilha(spreadsheetId, mostrarAlerta) {
  if (mostrarAlerta === undefined || mostrarAlerta === null) {
    mostrarAlerta = true;
  }
  var ui = SpreadsheetApp.getUi();
  var id = typeof spreadsheetId === 'string' ? spreadsheetId.trim() : '';

  if (!id) {
    var resp = ui.prompt(
      'Setup Planilha Filha',
      'Cole o Spreadsheet ID da planilha da unidade:\n(está na URL: docs.google.com/spreadsheets/d/ID_AQUI/edit)',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) {
      return null;
    }
    id = String(resp.getResponseText() || '').trim();
  }

  if (!id) {
    if (mostrarAlerta) {
      ui.alert('ID não informado.');
    }
    return null;
  }

  var ssFilha = abrirPlanilhaFilha(id, id);
  garantirAbaComCabecalho(ssFilha, CONFIG.ABAS.PEDIDO, obterHeaderPedidoUnidade_());
  garantirAbaCatalogo(ssFilha);

  var nomes = [];
  try {
    nomes = listarNomesProdutosAtivos();
  } catch (e) {
    Logger.log('Catálogo ainda vazio na Mãe: ' + e.message);
  }

  sincronizarDropdownFilha(ssFilha, nomes);
  removerAbaPadraoVazia(ssFilha);
  try {
    aplicarVisualPedidoUnidade_(obterAbaPedidoFilha_(ssFilha));
  } catch (eVis) {}

  registrarLog('SETUP_FILHA', ssFilha.getName() + ' (' + id + ')');

  try {
    ScriptApp.newTrigger('onOpenUnidadeInstalavel_').forSpreadsheet(id).onOpen().create();
  } catch (eTrig) {
    Logger.log('Setup filha onOpen: ' + eTrig.message);
  }

  if (mostrarAlerta) {
    ui.alert(
      'Planilha filha pronta',
      'Planilha: ' +
        ssFilha.getName() +
        '\n\nNão esqueça de cadastrá-la na aba Unidades da Planilha Mãe:\nNome | ' +
        id +
        ' | Sim',
      ui.ButtonSet.OK
    );
  }

  return id;
}

/** Configura o sistema no modo manual e remove gatilhos de pedidos. */
function instalarGatilhos() {
  removerGatilhosDoProjeto();
  removerGatilhosCatalogoAutomatico_();

  var menuUnidades = { instalados: 0, falhas: [] };
  try {
    garantirGatilhosSyncPedidos_();
    removerGatilhosCatalogoAutomatico_();
    removerTriggerPainelPendente_();
    garantirGatilhoVisibilidadeAbas_();
    menuUnidades = instalarGatilhoOnOpenUnidades_();
  } catch (e) {
    registrarLog('GATILHOS_UNIDADES_ERRO', e.message);
  }

  registrarLog(
    'GATILHOS',
    'Pedidos em modo manual; catálogo também manual; menu unidades=' + menuUnidades.instalados
  );

  SpreadsheetApp.getUi().alert(
    'Modo manual configurado',
    'Use: Sistema de Compras → Puxar pedidos das unidades\n\n' +
      'Menu nas unidades instalado em ' +
      menuUnidades.instalados +
      ' planilha(s).' +
      (menuUnidades.falhas.length
        ? '\n\nAtenção:\n- ' + menuUnidades.falhas.slice(0, 4).join('\n- ')
        : ''),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Remove apenas os gatilhos deste projeto (pelos nomes das funções).
 */
function removerGatilhosDoProjeto() {
  var funcoes = {
    onEditProdutosInstalavel: true,
    consolidarPedidosAgendado: true,
    onEditPedidoUnidade: true,
    processarSyncPedidoAtrasado: true,
    processarSyncPedidoPendenteRapido_: true,
    executarSyncAtrasadoUnidade_: true,
    sincronizarTodasUnidadesAgendado: true,
    onOpenUnidadeInstalavel_: true
  };
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (funcoes[triggers[i].getHandlerFunction()]) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Lista gatilhos atuais (diagnóstico).
 */
function menuListarGatilhos() {
  var triggers = ScriptApp.getProjectTriggers();
  if (!triggers.length) {
    SpreadsheetApp.getUi().alert('Nenhum gatilho instalado.');
    return;
  }
  var linhas = triggers.map(function (t) {
    return '- ' + t.getHandlerFunction() + ' (' + t.getEventType() + ')';
  });
  SpreadsheetApp.getUi().alert('Gatilhos instalados', linhas.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}
