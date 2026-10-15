/**
 * Sync automático: unidade marca Enviado → aparece na Mãe (Resumo + Consolidados).
 */

/**
 * UM clique: conecta todas as unidades + puxa pedidos já Enviado.
 * Use este menu uma vez (ou se a conexão cair).
 */
function menuConectarTudoAgora() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var resp = ui.alert(
    'Conectar unidades à Planilha Mãe',
    'Isso faz as planilhas das unidades conversarem com a Mãe.\n\n' +
      'Depois de autorizar uma vez:\n' +
      '• Unidade preenche Quantidade + Status = Enviado\n' +
      '• Aparece sozinho no Resumo\n' +
      '• Sem precisar rodar script de novo\n\n' +
      'Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    ss.toast('1/4 Resolvendo IDs no Drive...', 'Conectar', 20);
    var nIds = resolverIdsUnidadesPeloDrive();

    ss.toast('2/4 Limpando gatilhos antigos...', 'Conectar', 20);
    var nTrig = instalarGatilhosUnidades();

    ss.toast('3/4 Puxando pedidos Enviado...', 'Conectar', 90);
    var resultado = consolidarPedidosComDiagnostico();

    ss.toast('4/4 Concluído', 'Conectar', 5);
    SpreadsheetApp.flush();

    ui.alert(
      'Conexão ativa',
      'IDs das unidades: ' +
        nIds +
        '\nGatilhos automáticos instalados: ' +
        nTrig +
        '\nPedidos Enviado na Mãe agora: ' +
        resultado.linhas +
        '\nUnidades lidas: ' +
        resultado.unidades +
        (resultado.diagnosticoUnidades
          ? '\n\n--- LEITURA POR UNIDADE ---\n' +
            resultado.diagnosticoUnidades.join('\n')
          : '') +
        (resultado.amostraUnidadeCentro
          ? '\n\n--- UNIDADE CENTRO ---\n' + resultado.amostraUnidadeCentro
          : '') +
        (resultado.falhas.length
          ? '\n\nAtenção:\n- ' + resultado.falhas.slice(0, 6).join('\n- ')
          : '') +
        '\n\nPronto. Abra a aba Resumo.\n' +
        'Daqui pra frente, Enviado atualiza sozinho.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    registrarLog('CONECTAR_TUDO_ERRO', e.message);
    ui.alert('Erro ao conectar', e.message + '\n\n' + (e.stack || ''), ui.ButtonSet.OK);
  }
}

/**
 * Instalação única do sync automático em todas as unidades.
 */
function menuInstalarSyncAutomatico() {
  desativarAutomacaoPedidos_();
  SpreadsheetApp.getUi().alert(
    'Sincronização manual',
    'A automação de pedidos foi desativada.\n\nUse:\nSistema de Compras → 📥 Puxar pedidos das unidades',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Menu cliente: puxa todos os Enviados das unidades para o Resumo agora.
 */
function menuAtualizarResumoPedidosAgora() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Puxando pedidos Enviado de todas as unidades…', 'Pedidos', 90);
    var r = forcarPuxarResumoAgora();
    // Sem flush — flush na Mãe sobrecarregada causa timeout do serviço Planilhas.
    var statusEnvio = { enviaram: 0, total: 0, naoEnviaram: [] };
    try {
      statusEnvio = calcularStatusEnvioUnidades_(ss);
    } catch (eSt) {}
    try {
      var abaResumo = ss.getSheetByName(CONFIG.ABAS.RESUMO);
      if (abaResumo) abaResumo.activate();
    } catch (eAt) {}
    ui.alert(
      r.publicado ? 'Pedidos puxados' : 'Resumo preservado',
      'Já enviaram: ' +
        statusEnvio.enviaram +
        ' de ' +
        statusEnvio.total +
        ' unidades' +
        (statusEnvio.naoEnviaram && statusEnvio.naoEnviaram.length
          ? '\nAinda faltam: ' + statusEnvio.naoEnviaram.join(', ')
          : statusEnvio.total
            ? '\nTodas as unidades já enviaram.'
            : '') +
        '\n\nPedidos no Resumo: ' +
        (r.linhas || 0) +
        (r.falhas && r.falhas.length ? '\n\nAtenção:\n- ' + r.falhas.slice(0, 6).join('\n- ') : '') +
        '\n\nAbra a aba Resumo (modelo por pregão).',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro ao atualizar Resumo', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Puxa Enviados sem interface — usado pelo menu e pelo clasp run.
 * @return {{linhas:number, unidades:number, falhas:string[], amostraLucy:string}}
 */
function forcarPuxarResumoAgora() {
  try {
    desativarAutomacaoPedidos_();
  } catch (eTrig) {
    Logger.log('Limpeza de gatilhos: ' + eTrig.message);
  }
  try {
    completarIdsUnidadesFaltantes_();
  } catch (eIds) {
    Logger.log('forcarPuxar IDs: ' + eIds.message);
  }
  return consolidarPedidosComDiagnostico();
}

/**
 * Reconectar leve: IDs + puxar agora (sem reinstalar gatilhos).
 */
function menuReconectarPlanilhas() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Corrigindo IDs...', 'Reconectar', 8);
    var nIds = resolverIdsUnidadesPeloDrive();
    ss.toast('Puxando pedidos...', 'Reconectar', 25);
    var resultado = consolidarPedidosComDiagnostico();
    ui.alert(
      'Pedidos atualizados',
      'IDs: ' +
        nIds +
        '\nNo Resumo: ' +
        resultado.linhas +
        '\nUnidades ok: ' +
        resultado.unidades +
        (resultado.falhas.length
          ? '\n\nFalhas:\n- ' + resultado.falhas.slice(0, 6).join('\n- ')
          : '') +
        '\n\nSe Enviado ainda não atualizar sozinho, use:\nSistema de Compras → 🔌 Conectar unidades à Mãe',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}

function menuTestarUnidadeCentro() {
  var ui = SpreadsheetApp.getUi();
  try {
    var id = resolverIdUnidadePeloNome_('UNIDADE CENTRO');
    if (!id) throw new Error('Não achei Pedido - UNIDADE CENTRO no Drive');
    atualizarIdUnidadeNaAba('UNIDADE CENTRO', id);
    var ssFilha = SpreadsheetApp.openById(id);
    var aba = obterAbaPedidoFilha_(ssFilha);
    var lido = lerPedidosEnviadosFlexivel(aba);
    var enviados = lido.enviados.map(function (p) {
      return p.produto + ' = ' + p.quantidade;
    });
    sincronizarUnidadePorId(id);
    ui.alert(
      'Teste Unidade Centro',
      'Planilha: ' +
        ssFilha.getName() +
        '\nPedidos Enviados: ' +
        enviados.length +
        '\n' +
        (enviados.slice(0, 8).join('\n') || '(nenhum com Status Enviado)') +
        '\n\nSync desta unidade na Mãe: concluído.\nConfira o Resumo.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Falha', e.message, ui.ButtonSet.OK);
  }
}

/** Remove toda automação de pedidos; o fluxo oficial é manual. */
function instalarGatilhosUnidades() {
  var r = desativarAutomacaoPedidos_();
  registrarLog(
    'SYNC_PEDIDOS_MANUAL',
    'Gatilhos automáticos removidos: ' + r.removidos
  );
  return 0;
}

/** @return {{removidos:number}} */
function desativarAutomacaoPedidos_() {
  var triggers = ScriptApp.getProjectTriggers();
  var removidos = 0;
  var funcoesSync = {
    onEditPedidoUnidade: true,
    consolidarPedidosAgendado: true,
    processarSyncPedidoAtrasado: true,
    processarSyncPedidoPendenteRapido_: true,
    executarSyncAtrasadoUnidade_: true,
    sincronizarTodasUnidadesAgendado: true
  };
  for (var i = triggers.length - 1; i >= 0; i--) {
    if (!funcoesSync[triggers[i].getHandlerFunction()]) continue;
    ScriptApp.deleteTrigger(triggers[i]);
    removidos++;
  }
  limparPropsSyncPendente_();
  return { removidos: removidos };
}

/** Compatibilidade: qualquer chamada antiga apenas desativa a automação. */
function recriarGatilhosSyncUnidades_() {
  var r = desativarAutomacaoPedidos_();
  return {
    backupAtivo: false,
    totalOnEdit: 0,
    removidos: r.removidos,
    falhas: []
  };
}

/** Compatibilidade: reinstala a automação completa. */
function reinstalarGatilhosOnEditUnidades_() {
  var r = desativarAutomacaoPedidos_();
  return {
    backupAtivo: false,
    totalOnEdit: 0,
    removidos: r.removidos,
    falhas: []
  };
}

/** Garante que o modo manual permaneça sem gatilhos de pedidos. */
function garantirGatilhosSyncPedidos_() {
  var r = desativarAutomacaoPedidos_();
  return {
    backupAtivo: false,
    totalOnEdit: 0,
    removidos: r.removidos,
    falhas: []
  };
}

/** Desativado: pedidos só são puxados pelo menu manual. */
function pullPedidosAoAbrirMae_() {
  return;
}

/**
 * Diagnóstico e correção de gatilhos onEdit + backup 30 min.
 */
function menuDiagnosticarSyncAutomatico() {
  var ui = SpreadsheetApp.getUi();
  try {
    resolverIdsUnidadesPeloDrive();
    var unidades = listarUnidadesAtivas();
    var sync = garantirGatilhosSyncPedidos_();
    var resultado = consolidarPedidosComDiagnostico();
    var linhas = [];
    linhas.push('Unidades ativas: ' + unidades.length);
    linhas.push('Gatilhos onEdit: ' + sync.totalOnEdit + ' / ' + unidades.length);
    linhas.push('Backup 30 min: ' + (sync.backupAtivo ? 'ATIVO' : 'INATIVO'));
    linhas.push('Gatilhos antigos removidos: ' + sync.removidos);
    linhas.push('Unidades lidas agora: ' + resultado.unidades);
    linhas.push('Pedidos encontrados: ' + resultado.linhas);
    if (resultado.diagnosticoUnidades) {
      linhas.push('');
      linhas.push('Leitura por unidade:');
      linhas.push(resultado.diagnosticoUnidades.join('\n'));
    }
    if (resultado.falhas.length) {
      linhas.push('');
      linhas.push('Falhas:');
      linhas.push('- ' + resultado.falhas.slice(0, 10).join('\n- '));
    }

    ui.alert('Diagnóstico sync automático', linhas.join('\n'), ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}

/** Diagnóstico sem interface para conferir gatilhos pela API do Apps Script. */
function diagnosticarGatilhosSyncRemoto() {
  var triggers = ScriptApp.getProjectTriggers();
  return {
    total: triggers.length,
    gatilhos: triggers.map(function (t) {
      return {
        funcao: t.getHandlerFunction(),
        evento: String(t.getEventType())
      };
    })
  };
}

/** Remove só gatilhos onEdit de unidade (para reconectar limpo). */
function removerGatilhosOnEditUnidadesDuplicados_() {
  var triggers = ScriptApp.getProjectTriggers();
  var vistos = {};
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn !== 'onEditPedidoUnidade') continue;
    try {
      var src = String(triggers[i].getTriggerSourceId() || '').trim();
      if (!src) continue;
      if (vistos[src]) {
        ScriptApp.deleteTrigger(triggers[i]);
      } else {
        vistos[src] = true;
      }
    } catch (e) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function removerGatilhosUnidades() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (
      fn === 'onEditPedidoUnidade' ||
      fn === 'sincronizarTodasUnidadesAgendado' ||
      fn === 'processarSyncPedidoAtrasado' ||
      fn === 'processarSyncPedidoPendenteRapido_' ||
      fn === 'executarSyncAtrasadoUnidade_'
    ) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function garantirTriggerConsolidacaoFrequente() {
  return garantirGatilhosSyncPedidos_();
}

/**
 * Um disparo ~6s após edições em sequência (qtd + status).
 */
function garantirTriggerSyncPedidoAtrasado_() {
  return garantirGatilhosSyncPedidos_();
}

function agendarSyncUnidadePedido_(spreadsheetId) {
  var props = PropertiesService.getScriptProperties();
  var id = String(spreadsheetId || '').trim();
  if (!id) return;
  var ids = lerFilaSyncPendente_(props);
  if (ids.indexOf(id) === -1) ids.push(id);
  salvarFilaSyncPendente_(props, ids);
  props.setProperty('SYNC_PENDENTE_TS', String(Date.now()));
}

/**
 * @param {GoogleAppsScript.Properties.Properties} props
 * @return {string[]}
 */
function lerFilaSyncPendente_(props) {
  props = props || PropertiesService.getScriptProperties();
  var raw = props.getProperty('SYNC_PENDENTE_IDS');
  if (raw) {
    try {
      var arr = JSON.parse(raw);
      if (Array.isArray(arr)) {
        return arr
          .map(function (x) {
            return String(x || '').trim();
          })
          .filter(function (x) {
            return !!x;
          });
      }
    } catch (eJson) {}
  }
  var legado = props.getProperty('SYNC_PENDENTE_ID');
  return legado ? [String(legado).trim()] : [];
}

/**
 * @param {GoogleAppsScript.Properties.Properties} props
 * @param {string[]} ids
 */
function salvarFilaSyncPendente_(props, ids) {
  if (!ids || !ids.length) {
    props.deleteProperty('SYNC_PENDENTE_IDS');
    props.deleteProperty('SYNC_PENDENTE_ID');
    props.deleteProperty('SYNC_PENDENTE_TS');
    return;
  }
  props.setProperty('SYNC_PENDENTE_IDS', JSON.stringify(ids));
  props.deleteProperty('SYNC_PENDENTE_ID');
}

/**
 * Limpa fila de sync pendente (ex.: após Nova rodada).
 */
function limparPropsSyncPendente_() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('SYNC_PENDENTE_ID');
  props.deleteProperty('SYNC_PENDENTE_IDS');
  props.deleteProperty('SYNC_PENDENTE_TS');
  props.deleteProperty('SYNC_AFTER_ID');
  props.deleteProperty('TOTAL_PRODUTO_PENDENTE');
}

/**
 * Remove gatilhos time-based de uma função (exceto everyMinutes do backup).
 * @param {string} nomeFuncao
 */
function removerGatilhosClockPorFuncao_(nomeFuncao) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = triggers.length - 1; i >= 0; i--) {
    if (triggers[i].getHandlerFunction() !== nomeFuncao) continue;
    try {
      ScriptApp.deleteTrigger(triggers[i]);
    } catch (eDel) {}
  }
}

/**
 * Disparo único ~7s — processa fila SYNC_PENDENTE (caminho rápido confiável).
 */
function criarGatilhoProcessarSyncPendenteRapido_() {
  try {
    removerGatilhosClockPorFuncao_('processarSyncPedidoPendenteRapido_');
    ScriptApp.newTrigger('processarSyncPedidoPendenteRapido_').timeBased().after(7000).create();
  } catch (e) {
    Logger.log('criarGatilhoProcessarSyncPendenteRapido_: ' + e.message);
  }
}

/** Handler one-shot (~7s) para processar sync pendente. */
function processarSyncPedidoPendenteRapido_() {
  try {
    processarSyncPedidoAtrasado();
  } finally {
    removerGatilhosClockPorFuncao_('processarSyncPedidoPendenteRapido_');
  }
}

/**
 * Processa sync pendente (debounce) e gatilhos .after() avulsos.
 */
function processarSyncPedidoAtrasado() {
  var props = PropertiesService.getScriptProperties();
  var ids = lerFilaSyncPendente_(props);
  if (!ids.length) return;
  var ts = Number(props.getProperty('SYNC_PENDENTE_TS') || '0');
  if (Date.now() - ts < 5500) return;

  var id = ids[0];
  try {
    var n = sincronizarUnidadePorId(id);
    if (n >= 0) {
      ids.shift();
      salvarFilaSyncPendente_(props, ids);
      if (ids.length) {
        props.setProperty('SYNC_PENDENTE_TS', String(Date.now()));
        criarGatilhoProcessarSyncPendenteRapido_();
      }
    }
  } catch (e) {
    registrarLog('SYNC_PENDENTE_ERRO', id + ': ' + e.message);
  }
}

/**
 * Disparo único ~6s depois (complementa o gatilho de 1 min).
 * @param {string} spreadsheetId
 */
function criarGatilhoSyncAtrasadoUnidade_(spreadsheetId) {
  try {
    removerGatilhosClockPorFuncao_('executarSyncAtrasadoUnidade_');
    ScriptApp.newTrigger('executarSyncAtrasadoUnidade_').timeBased().after(6000).create();
    PropertiesService.getScriptProperties().setProperty(
      'SYNC_AFTER_ID',
      String(spreadsheetId).trim()
    );
  } catch (e) {
    Logger.log('criarGatilhoSyncAtrasadoUnidade_: ' + e.message);
  }
}

function executarSyncAtrasadoUnidade_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SYNC_AFTER_ID');
  if (!id) {
    removerGatilhosClockPorFuncao_('executarSyncAtrasadoUnidade_');
    return;
  }
  try {
    var n = sincronizarUnidadePorId(id);
    if (n >= 0) {
      props.deleteProperty('SYNC_AFTER_ID');
      var ids = lerFilaSyncPendente_(props);
      var idx = ids.indexOf(id);
      if (idx !== -1) {
        ids.splice(idx, 1);
        salvarFilaSyncPendente_(props, ids);
      }
    }
  } catch (e) {
    registrarLog('SYNC_AFTER_ERRO', e.message);
  } finally {
    removerGatilhosClockPorFuncao_('executarSyncAtrasadoUnidade_');
  }
}

/**
 * Agenda sync da unidade ~6–7s após editar a coluna Status.
 * Sempre enfileira na fila (backup 1 min) + dispara caminhos rápidos.
 * @param {string} spreadsheetId
 */
function agendarSyncPedidoUnidadeDebounced_(spreadsheetId) {
  var id = String(spreadsheetId || '').trim();
  if (!id) return;
  agendarSyncUnidadePedido_(id);
  criarGatilhoProcessarSyncPendenteRapido_();
  criarGatilhoSyncAtrasadoUnidade_(id);
}

/**
 * Backup: a cada 1 min puxa Enviado de todas as unidades → Mãe.
 */
function sincronizarTodasUnidadesAgendado() {
  try {
    var r = consolidarPedidosComDiagnostico();
    registrarLog(
      'SYNC_AGENDADO_OK',
      (r.linhas || 0) + ' pedido(s) | unidades ' + (r.unidades || 0)
    );
  } catch (e) {
    registrarLog('SYNC_TODAS_AGENDADO_ERRO', e.message);
  }
}

/**
 * Gatilho automático na planilha da unidade.
 * Dispara somente quando a coluna Status é editada.
 * @param {Object} e
 */
function onEditPedidoUnidade(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (!sheet) return;
    if (!ehAbaPedido_(sheet)) return;

    if (e.range.getColumn() > 10) return;

    if (!rangeIncluiColunaStatus_(sheet, e.range)) return;

    var idFilha = e.source.getId();
    var n = sincronizarUnidadePorId(idFilha);
    if (n < 0) {
      registrarLog('SYNC_ONEDIT_OCUPADO', idFilha + ': backup em até 30 min');
    }
  } catch (err) {
    registrarLog('BRIDGE_ONEDIT_ERRO', err.message);
  }
}

/**
 * @param {Sheet} sheet
 * @param {Range} range
 * @return {boolean}
 */
function rangeIncluiColunaStatus_(sheet, range) {
  return rangeIncluiColunasPedido_(sheet, range, ['status']);
}

function rangeIncluiColunaStatusOuQuantidade_(sheet, range) {
  return rangeIncluiColunasPedido_(sheet, range, ['status', 'qtd']);
}

/**
 * @param {Sheet} sheet
 * @param {Range} range
 * @param {string[]} quais status|qtd
 * @return {boolean}
 */
function rangeIncluiColunasPedido_(sheet, range, quais) {
  try {
    var colStatus = CONFIG.PEDIDO.STATUS;
    var colQtd = CONFIG.PEDIDO.QUANTIDADE;
    var linha1 = String(sheet.getRange(1, 1).getDisplayValue() || '').toLowerCase();
    if (linha1.indexOf('planilha de pedidos') === -1) {
      var lr = Math.min(15, Math.max(sheet.getLastRow(), 1));
      var lc = Math.min(8, Math.max(sheet.getLastColumn(), 6));
      var vals = sheet.getRange(1, 1, lr, lc).getValues();
      var mapa = detectarColunasPedido_(vals);
      colStatus = mapa.colStatus >= 0 ? mapa.colStatus + 1 : CONFIG.PEDIDO.STATUS;
      colQtd = mapa.colQtd >= 0 ? mapa.colQtd + 1 : CONFIG.PEDIDO.QUANTIDADE;
    }
    var c1 = range.getColumn();
    var c2 = c1 + range.getNumColumns() - 1;
    if (quais.indexOf('status') !== -1 && colStatus >= c1 && colStatus <= c2) return true;
    if (quais.indexOf('qtd') !== -1 && colQtd >= c1 && colQtd <= c2) return true;
    return false;
  } catch (e) {
    var col = range.getColumn();
    return col === CONFIG.PEDIDO.STATUS || col === CONFIG.PEDIDO.QUANTIDADE;
  }
}

/**
 * @param {Sheet} sheet
 * @return {boolean}
 */
function ehAbaPedido_(sheet) {
  var nome = String(sheet.getName() || '').toLowerCase();
  if (nome === 'pedido' || nome.indexOf('pedido') !== -1) return true;
  try {
    var vals = sheet.getRange(1, 1, Math.min(8, Math.max(sheet.getLastRow(), 1)), Math.min(8, Math.max(sheet.getLastColumn(), 5))).getValues();
    for (var r = 0; r < vals.length; r++) {
      for (var c = 0; c < vals[r].length; c++) {
        var h = String(vals[r][c] || '')
          .toLowerCase()
          .replace(/[^a-z]/g, '');
        if (h === 'produto') return true;
      }
    }
  } catch (e) {}
  return false;
}

/**
 * Debounce por unidade (~4s).
 * @param {string} spreadsheetId
 * @return {boolean} true = pular
 */
function syncUnidadeEmDebounce_(spreadsheetId) {
  var props = PropertiesService.getScriptProperties();
  var key = 'DEB_SYNC_' + String(spreadsheetId).substring(0, 20);
  var agora = Date.now();
  var ultimo = Number(props.getProperty(key) || '0');
  if (agora - ultimo < 6000) return true;
  props.setProperty(key, String(agora));
  return false;
}

/**
 * @param {Spreadsheet} ssFilha
 * @return {Sheet}
 */
function obterAbaPedidoFilha_(ssFilha) {
  var aba = ssFilha.getSheetByName(CONFIG.ABAS.PEDIDO);
  if (aba) return aba;
  var sheets = ssFilha.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (ehAbaPedido_(sheets[i])) return sheets[i];
  }
  throw new Error('Sem aba Pedido em ' + ssFilha.getName());
}

/**
 * Lê uma unidade e atualiza Resumo + Consolidados na Mãe.
 * @param {string} spreadsheetId
 */
function sincronizarUnidadePorId(spreadsheetId) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    agendarSyncUnidadePedido_(spreadsheetId);
    return -1;
  }
  try {
    var unidade = encontrarUnidadePorId(spreadsheetId);
    if (!unidade) {
      var ssTmp = SpreadsheetApp.openById(spreadsheetId);
      var nomeArq = ssTmp.getName();
      var nomeUnidade = nomeArq.replace(/^Pedido\s*-\s*/i, '').trim();
      unidade = { nome: nomeUnidade, spreadsheetId: spreadsheetId, ativo: true };
    }

    var ssFilha = SpreadsheetApp.openById(spreadsheetId);
    var abaPedido = obterAbaPedidoFilha_(ssFilha);
    var lido = lerPedidosEnviadosFlexivel(abaPedido);
    if (!lido.enviados.length) {
      registrarLog(
        'SYNC_UNIDADE_SEM_ENVIADO',
        unidade.nome + ' — 0 enviado(s)\n' + (lido.diagnostico || '')
      );
      return 0;
    }
    atualizarMaeComPedidosUnidade(unidade.nome, lido.enviados);
    SpreadsheetApp.flush();

    registrarLog(
      'SYNC_UNIDADE_OK',
      unidade.nome + ' — ' + lido.enviados.length + ' enviado(s)'
    );
    return lido.enviados.length;
  } catch (eRun) {
    registrarLog('SYNC_UNIDADE_ERRO', spreadsheetId + ': ' + eRun.message);
    throw eRun;
  } finally {
    lock.releaseLock();
  }
}

/**
 * @param {string} spreadsheetId
 * @return {Object|null}
 */
function encontrarUnidadePorId(spreadsheetId) {
  var unidades = listarTodasUnidades();
  var id = String(spreadsheetId || '').trim();
  for (var i = 0; i < unidades.length; i++) {
    if (String(unidades[i].spreadsheetId || '').trim() === id) {
      return unidades[i];
    }
  }
  return null;
}

/**
 * Atualiza Resumo e Consolidados na Mãe.
 * @param {string} nomeUnidade
 * @param {Array.<Object>} pedidosEnviados
 */
function atualizarMaeComPedidosUnidade(nomeUnidade, pedidosEnviados) {
  var ssMae = obterPlanilhaMae();
  var chaveUn = normalizarTexto(nomeUnidade);

  var unidades = listarUnidadesAtivas();
  var nomesUnidades = unidades.map(function (u) {
    return u.nome;
  });
  if (nomesUnidades.length === 0 && DADOS_INICIAIS && DADOS_INICIAIS.unidades) {
    nomesUnidades = DADOS_INICIAIS.unidades.slice();
  }

  var temUnidade = false;
  for (var i0 = 0; i0 < nomesUnidades.length; i0++) {
    if (normalizarTexto(nomesUnidades[i0]) === chaveUn) {
      temUnidade = true;
      break;
    }
  }
  if (!temUnidade) {
    nomesUnidades.push(nomeUnidade);
  }

  // --- Consolidados (fonte dos dados) ---
  var novoCons = [];
  var abaConsOld = ssMae.getSheetByName(CONFIG.ABAS.CONSOLIDADOS);
  if (abaConsOld) {
    var consAtual = lerDados(abaConsOld, CONFIG.CONSOLIDADOS.HEADER.length);
    for (var k = 0; k < consAtual.length; k++) {
      if (normalizarTexto(consAtual[k][0]) === chaveUn) continue;
      novoCons.push(consAtual[k].slice(0, CONFIG.CONSOLIDADOS.HEADER.length));
    }
  }
  for (var q = 0; q < pedidosEnviados.length; q++) {
    novoCons.push([
      nomeUnidade,
      pedidosEnviados[q].produto,
      valorCelulaPreservandoZero_(pedidosEnviados[q].estoqueAtual),
      String(pedidosEnviados[q].quantidade || '').trim(),
      pedidosEnviados[q].observacao || '',
      pedidosEnviados[q].data || obterDataMesReferenciaPedido_()
    ]);
  }
  escreverAbaSimples_(ssMae, CONFIG.ABAS.CONSOLIDADOS, CONFIG.CONSOLIDADOS.HEADER, novoCons);

  // --- Resumo (visual por unidade, gerado a partir de Consolidados) ---
  escreverResumoAgrupadoFromConsolidados_(ssMae, novoCons, opcoesSyncAutomaticoMae_());

  atualizarIndicadorUltimaAtualizacao_(ssMae);
  try {
    aplicarVisibilidadeAbasCliente_(ssMae);
  } catch (eVis) {}

  SpreadsheetApp.flush();
}

/**
 * @param {string} s
 * @return {string}
 */
function chaveProdutoFlex_(s) {
  return normalizarTexto(s)
    .replace(/á|à|ã|â/g, 'a')
    .replace(/é|ê/g, 'e')
    .replace(/í/g, 'i')
    .replace(/ó|ô|õ/g, 'o')
    .replace(/ú/g, 'u')
    .replace(/ç/g, 'c')
    .replace(/[^a-z0-9]/g, '');
}

/** Mesma normalização de chaveProdutoFlex_ — usada para colunas de unidade na matriz. */
function normalizarChaveUnidade_(nome) {
  return chaveProdutoFlex_(nome);
}

/**
 * Grava aba na Mãe mantendo o visual (cabeçalho verde, zebrado, larguras).
 * Não apaga a aba — só atualiza valores + reaplica o modelo.
 * Se a aba tiver "Tabela" tipada (erro de coluna com tipo), recria e formata.
 * @param {Spreadsheet} ss
 * @param {string} nomeAba
 * @param {string[]} header
 * @param {Array.<Array.<*>>} linhas
 */
function escreverAbaSimples_(ss, nomeAba, header, linhas) {
  // Produtos: só grava valores, NÃO mexe no visual que o usuário configurou
  if (nomeAba === CONFIG.ABAS.PRODUTOS) {
    escreverSomenteValoresSemFormatar_(ss, nomeAba, header, linhas || []);
    return;
  }

  var estilo =
    nomeAba === CONFIG.ABAS.LISTA_PEDIDOS
      ? 'lista'
      : nomeAba === CONFIG.ABAS.RESUMO
        ? 'resumo'
        : nomeAba === CONFIG.ABAS.CONSOLIDADOS
          ? 'consolidados'
          : nomeAba === CONFIG.ABAS.UNIDADES
            ? 'unidades'
            : 'padrao';
  escreverAbaComModelo_(ss, nomeAba, header, linhas || [], estilo);
  if (nomeAba === CONFIG.ABAS.CONSOLIDADOS || nomeAba === CONFIG.ABAS.LOG) {
    ocultarAbaAdministrativaSeVisivel_(ss, nomeAba);
  }
}

/**
 * Grava dados sem alterar cores/formato (usado na aba Produtos).
 */
function escreverSomenteValoresSemFormatar_(ss, nomeAba, header, linhas) {
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) aba = ss.insertSheet(nomeAba);
  try {
    gravarValoresPreservandoAba_(aba, header, linhas);
  } catch (e) {
    var msg = String(e.message || e);
    if (msg.indexOf('tipo') !== -1 || msg.indexOf('type') !== -1) {
      try {
        ss.deleteSheet(aba);
      } catch (eDel) {
        aba.setName(nomeAba + '_OLD_' + new Date().getTime());
      }
      aba = ss.insertSheet(nomeAba);
      gravarValoresPreservandoAba_(aba, header, linhas);
    } else {
      throw e;
    }
  }
}

/**
 * @param {Spreadsheet} ss
 * @param {string} nomeAba
 * @param {string[]} header
 * @param {Array.<Array.<*>>} linhas
 * @param {string} estilo
 */
function escreverAbaComModelo_(ss, nomeAba, header, linhas, estilo) {
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) {
    aba = ss.insertSheet(nomeAba);
  }

  try {
    gravarValoresPreservandoAba_(aba, header, linhas);
  } catch (e) {
    var msg = String(e.message || e);
    // Tabela tipada / colunas com tipo → recria aba limpa e aplica o mesmo visual
    if (
      msg.indexOf('tipo') !== -1 ||
      msg.indexOf('type') !== -1 ||
      msg.indexOf('not allowed') !== -1 ||
      msg.indexOf('Tabela') !== -1
    ) {
      try {
        ss.deleteSheet(aba);
      } catch (eDel) {
        try {
          aba.setName(nomeAba + '_OLD_' + new Date().getTime());
        } catch (eRen) {}
      }
      aba = ss.insertSheet(nomeAba);
      gravarValoresPreservandoAba_(aba, header, linhas);
    } else {
      throw e;
    }
  }

  aplicarModeloVisualMae_(aba, header.length, linhas.length, estilo);
}

/**
 * @param {Sheet} aba
 * @param {string[]} header
 * @param {Array.<Array.<*>>} linhas
 */
function gravarValoresPreservandoAba_(aba, header, linhas) {
  var nCols = header.length;
  var nRows = linhas && linhas.length ? linhas.length : 0;
  var lastR = Math.max(aba.getLastRow(), 1);
  var lastC = Math.max(aba.getLastColumn(), nCols);

  try {
    var filtro = aba.getFilter();
    if (filtro) filtro.remove();
  } catch (eF) {}

  aba.getRange(1, 1, 1, nCols).setValues([header]);
  if (nRows > 0) {
    // getRange(linha, coluna, numeroDeLinhas, numeroDeColunas)
    aba.getRange(2, 1, nRows, nCols).setValues(linhas);
  }

  var primeiraLinhaSobra = nRows + 2;
  if (lastR >= primeiraLinhaSobra) {
    var linhasSobra = lastR - primeiraLinhaSobra + 1;
    aba.getRange(primeiraLinhaSobra, 1, linhasSobra, lastC).clearContent();
  }
  if (lastC > nCols) {
    var linhasLargura = Math.max(lastR, nRows + 1);
    aba.getRange(1, nCols + 1, linhasLargura, lastC - nCols).clearContent();
  }
}

/**
 * Modelo fixo: cabeçalho vermelho, coluna principal amarela, resto dark.
 * @param {Sheet} aba
 * @param {number} nCols
 * @param {number} nDataRows
 * @param {string} estilo
 */
function aplicarModeloVisualMae_(aba, nCols, nDataRows, estilo) {
  var t = CONFIG.TEMA;
  var totalRows = Math.max(1, nDataRows + 1);
  var rangeAll = aba.getRange(1, 1, totalRows, nCols);

  // Coluna amarela (destaque do nome/item/produto)
  var colAmarela = 1; // padrão: 1ª coluna
  if (estilo === 'resumo' || estilo === 'consolidados') {
    colAmarela = 2; // Produto
  }

  try {
    rangeAll.clearFormat();
  } catch (eClear) {}

  try {
    aba.setTabColor(t.TAB);
  } catch (eTab) {}

  try {
    var maxR = Math.min(aba.getMaxRows(), Math.max(totalRows + 50, 100));
    var maxC = Math.min(aba.getMaxColumns(), Math.max(nCols + 4, 22));
    aba.getRange(1, 1, maxR, maxC).setBackground(t.SHEET_BG);
  } catch (eBg) {}

  // Cabeçalho vermelho
  var headerRange = aba.getRange(1, 1, 1, nCols);
  headerRange.setBackground(t.HEADER_BG);
  headerRange.setFontColor(t.HEADER_FG);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setFontFamily('Arial');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setWrap(true);
  aba.setRowHeight(1, 42);

  aba.setFrozenRows(1);
  if (estilo === 'lista') {
    aba.setFrozenColumns(1);
  } else {
    aba.setFrozenColumns(0);
  }

  if (nDataRows > 0) {
    var dataRange = aba.getRange(2, 1, nDataRows, nCols);
    dataRange.setFontSize(10);
    dataRange.setFontFamily('Arial');
    dataRange.setVerticalAlignment('middle');

    var backgrounds = [];
    var fontColors = [];
    for (var r = 0; r < nDataRows; r++) {
      var corLinha = r % 2 === 1 ? t.ROW_B : t.ROW_A;
      var rowBg = [];
      var rowFg = [];
      for (var c = 0; c < nCols; c++) {
        if (c + 1 === colAmarela) {
          rowBg.push(t.COL_DESTAQUE_BG);
          rowFg.push(t.COL_DESTAQUE_FG);
        } else {
          rowBg.push(corLinha);
          rowFg.push(c === 0 ? t.TEXT : t.TEXT_MUTED);
        }
      }
      backgrounds.push(rowBg);
      fontColors.push(rowFg);
    }
    dataRange.setBackgrounds(backgrounds);
    dataRange.setFontColors(fontColors);

    // Coluna amarela: negrito + centralizado (como no modelo)
    aba
      .getRange(2, colAmarela, nDataRows, 1)
      .setFontWeight('bold')
      .setHorizontalAlignment(estilo === 'lista' ? 'left' : 'center');

    // Demais colunas centralizadas
    for (var c2 = 1; c2 <= nCols; c2++) {
      if (c2 === colAmarela) continue;
      aba.getRange(2, c2, nDataRows, 1).setHorizontalAlignment('center');
    }

    if (estilo === 'lista' && nCols >= 2) {
      aba.getRange(2, nCols, nDataRows, 1).setFontWeight('bold').setFontColor(t.ACCENT_QTD);
    }
  }

  // Larguras: unidades fixas; colunas de NOME/ITEM always auto (nome completo)
  if (estilo === 'lista') {
    for (var w = 2; w < nCols; w++) aba.setColumnWidth(w, 120);
    if (nCols >= 2) aba.setColumnWidth(nCols, 80);
    ajustarLarguraNomeCompleto_(aba, 1, nDataRows); // ITEM
  } else if (estilo === 'resumo') {
    aba.setColumnWidth(1, 180); // Unidade
    if (nCols >= 3) aba.setColumnWidth(3, 110); // Quantidade
    ajustarLarguraNomeCompleto_(aba, 2, nDataRows); // Produto
  } else if (estilo === 'consolidados') {
    aba.setColumnWidth(1, 160);
    aba.setColumnWidth(3, 100);
    if (nCols >= 4) aba.setColumnWidth(4, 160);
    if (nCols >= 5) aba.setColumnWidth(5, 120);
    ajustarLarguraNomeCompleto_(aba, 2, nDataRows); // Produto
  } else if (estilo === 'unidades') {
    aba.setColumnWidth(2, 360);
    if (nCols >= 3) aba.setColumnWidth(3, 80);
    ajustarLarguraNomeCompleto_(aba, 1, nDataRows); // Nome
  } else {
    for (var w2 = 1; w2 <= nCols; w2++) {
      aba.setColumnWidth(w2, w2 === 1 ? 180 : 140);
    }
    ajustarLarguraNomeCompleto_(aba, colAmarela, nDataRows);
  }

  rangeAll.setBorder(
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
    var filtro = aba.getFilter();
    if (filtro) filtro.remove();
  } catch (eRem) {}
  try {
    aba.getRange(1, 1, totalRows, nCols).createFilter();
  } catch (eFiltro) {}
}

/**
 * Garante que o nome do item/produto apareça inteiro (não corta).
 * Usa o maior texto da coluna + autoResize, com mínimo generoso.
 * @param {Sheet} aba
 * @param {number} col 1-based
 * @param {number} nDataRows
 */
function ajustarLarguraNomeCompleto_(aba, col, nDataRows) {
  var minW = 480;
  var maxW = 900;
  var largura = minW;

  try {
    // Calcula pela maior string (mais confiável que só autoResize)
    var maxLen = 12; // cabeçalho
    if (nDataRows > 0) {
      var vals = aba.getRange(2, col, nDataRows, 1).getDisplayValues();
      for (var i = 0; i < vals.length; i++) {
        var s = String(vals[i][0] || '');
        if (s.length > maxLen) maxLen = s.length;
      }
    }
    // ~8.5 px por caractere em Arial 10 bold + folga
    largura = Math.ceil(maxLen * 8.5 + 40);
  } catch (eCalc) {}

  try {
    aba.autoResizeColumn(col);
    var autoW = aba.getColumnWidth(col);
    if (autoW > largura) largura = autoW;
  } catch (eAuto) {}

  if (largura < minW) largura = minW;
  if (largura > maxW) largura = maxW;
  aba.setColumnWidth(col, largura);

  // Sem quebra forçada: nome em uma linha, coluna larga o suficiente
  try {
    if (nDataRows > 0) {
      aba.getRange(2, col, nDataRows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
    }
  } catch (eWrap) {
    try {
      if (nDataRows > 0) aba.getRange(2, col, nDataRows, 1).setWrap(false);
    } catch (e2) {}
  }
}
