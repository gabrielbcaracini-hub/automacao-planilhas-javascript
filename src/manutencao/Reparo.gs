/**
 * Reparo em etapas (segundo plano) — não trava a tela do Google Sheets.
 */

var ABAS_SO_MAE = {
  Produtos: true,
  Unidades: true,
  Consolidados: true,
  Resumo: true,
  Log: true
};

/**
 * Menu: inicia reparo em segundo plano e libera a tela na hora.
 */
function menuRepararSistemaCompleto() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Reparar em segundo plano',
    'O reparo completo trava se rodar na tela.\n\n' +
      'Agora ele roda em etapas automáticas (2–4 min).\n' +
      'Você pode continuar usando a planilha.\n\n' +
      'Acompanhe na aba Log.\n\nIniciar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    PropertiesService.getScriptProperties().setProperty('REPARO_PASSO', '1');
    PropertiesService.getScriptProperties().deleteProperty('CATALOGO_PENDENTE');
    agendarProximoPassoReparo(5000);
    ui.alert(
      'Reparo iniciado',
      'Rodando em segundo plano.\n\n' +
        'Enquanto isso, use também:\n' +
        '• Puxar pedidos Enviado agora (rápido)\n\n' +
        'Em alguns minutos confira Resumo + aba Log.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Menu RÁPIDO: só traz pedidos Enviado para a Mãe (sem sync de catálogo).
 */
function menuPuxarPedidosRapidoLegado_() {
  var ui = SpreadsheetApp.getUi();
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Puxando pedidos Enviado...', 'Sync', 20);
    aplicarIdsPlanilhasUnidades();
    var cons = consolidarPedidos(false, true);
    ui.alert(
      'Pedidos atualizados',
      'Linhas no Resumo: ' +
        cons.linhas +
        '\nUnidades lidas: ' +
        cons.unidades +
        (cons.falhas.length
          ? '\n\nFalhas:\n- ' + cons.falhas.slice(0, 6).join('\n- ')
          : '\n\nAbra a aba Resumo.'),
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Agenda o worker de reparo.
 * @param {number} depoisMs
 */
function agendarProximoPassoReparo(depoisMs) {
  removerTriggersReparo();
  ScriptApp.newTrigger('executarReparoEmPassos')
    .timeBased()
    .after(Math.max(1000, depoisMs || 3000))
    .create();
}

function removerTriggersReparo() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'executarReparoEmPassos') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Worker: um passo por execução (evita timeout).
 */
function executarReparoEmPassos() {
  var props = PropertiesService.getScriptProperties();
  var passo = Number(props.getProperty('REPARO_PASSO') || '0');
  if (!passo) {
    removerTriggersReparo();
    return;
  }

  try {
    if (passo === 1) {
      var n = aplicarIdsPlanilhasUnidades();
      registrarLog('REPARO_PASSO_1', 'IDs corrigidos: ' + n);
      props.setProperty('REPARO_PASSO', '2');
      agendarProximoPassoReparo(4000);
      return;
    }

    if (passo === 2) {
      var limpeza = limparAbasErradasTodasFilhas();
      registrarLog(
        'REPARO_PASSO_2',
        'Limpeza filhas: ' + limpeza.ok + ' ok / ' + limpeza.falhas.length + ' erros'
      );
      props.setProperty('REPARO_PASSO', '3');
      agendarProximoPassoReparo(4000);
      return;
    }

    if (passo === 3) {
      // Catálogo — o passo mais pesado
      var syncCat = sincronizarCatalogoTodasUnidades(true, true);
      registrarLog(
        'REPARO_PASSO_3',
        'Catálogo: ' + syncCat.sucesso + ' ok / ' + syncCat.falhas.length + ' erros'
      );
      props.setProperty('REPARO_PASSO', '4');
      agendarProximoPassoReparo(5000);
      return;
    }

    if (passo === 4) {
      var cons = consolidarPedidos(false, true);
      registrarLog(
        'REPARO_PASSO_4',
        'Pedidos na Mãe: ' + cons.linhas + ' linhas / ' + cons.unidades + ' unidades'
      );
      props.setProperty('REPARO_PASSO', '5');
      agendarProximoPassoReparo(4000);
      return;
    }

    if (passo === 5) {
      var nTrig = 0;
      try {
        nTrig = instalarGatilhosUnidades();
      } catch (e) {
        registrarLog('REPARO_PASSO_5_ERRO', e.message);
      }
      garantirTriggerConsolidacaoFrequente();
      removerGatilhosCatalogoAutomatico_();
      removerTriggerPainelPendente_();
      garantirGatilhoVisibilidadeAbas_();
      registrarLog('REPARO_PASSO_5', 'Sync automático: ' + nTrig + ' unidades');
      props.deleteProperty('REPARO_PASSO');
      removerTriggersReparo();
      registrarLog('REPARO_OK', 'Reparo em etapas finalizado');
      return;
    }

    props.deleteProperty('REPARO_PASSO');
    removerTriggersReparo();
  } catch (e) {
    registrarLog('REPARO_PASSO_ERRO', 'passo=' + passo + ' — ' + e.message);
    // tenta o mesmo passo de novo 1x
    var tent = Number(props.getProperty('REPARO_TENTATIVA') || '0');
    if (tent < 1) {
      props.setProperty('REPARO_TENTATIVA', String(tent + 1));
      agendarProximoPassoReparo(10000);
    } else {
      props.deleteProperty('REPARO_PASSO');
      props.deleteProperty('REPARO_TENTATIVA');
      removerTriggersReparo();
    }
  }
}

/**
 * @return {{ok: number, falhas: string[]}}
 */
function limparAbasErradasTodasFilhas() {
  var resultado = { ok: 0, falhas: [] };
  var unidades = listarUnidadesAtivas();

  for (var i = 0; i < unidades.length; i++) {
    var u = unidades[i];
    try {
      var ss = SpreadsheetApp.openById(u.spreadsheetId);
      limparAbasErradasDaFilha(ss);
      garantirAbaComCabecalho(ss, CONFIG.ABAS.PEDIDO, obterHeaderPedidoUnidade_());
      garantirAbaCatalogo(ss);
      resultado.ok++;
    } catch (e) {
      resultado.falhas.push(u.nome + ': ' + e.message);
    }
  }
  return resultado;
}

/**
 * @param {Spreadsheet} ssFilha
 */
function limparAbasErradasDaFilha(ssFilha) {
  var abas = ssFilha.getSheets();
  for (var i = abas.length - 1; i >= 0; i--) {
    var nome = abas[i].getName();
    if (!ABAS_SO_MAE[nome]) continue;
    if (ssFilha.getSheets().length <= 1) continue;
    try {
      ssFilha.deleteSheet(abas[i]);
    } catch (e) {
      Logger.log('Não removeu aba ' + nome + ': ' + e.message);
    }
  }
}
