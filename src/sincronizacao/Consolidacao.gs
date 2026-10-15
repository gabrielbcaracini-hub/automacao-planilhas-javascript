/**
 * Consolidação dos pedidos das Planilhas Filhas na Planilha Mãe.
 */

/**
 * Consolida pedidos das unidades ativas.
 * - Resumo: matriz ITEM × unidades
 * - Consolidados: apenas Status = Enviado (quantidade, observação e data)
 * @param {boolean=} marcarComoConsolidado Se true, altera Status nas filhas para Consolidado.
 * @return {{linhas: number, unidades: number, falhas: Array.<string>}}
 */
function consolidarPedidos(marcarComoConsolidado, semLock) {
  var lock = null;
  if (!semLock) {
    lock = adquirirLock('consolidacao', 45);
  }
  var resultado = { linhas: 0, unidades: 0, falhas: [] };

  try {
    var ssMae = obterPlanilhaMae();
    var unidades = listarUnidadesAtivas();
    var linhasConsolidadas = [];
    var matriz = {}; // chaveFlex -> { unidadeNorm -> qtd, __nome }

    for (var i = 0; i < unidades.length; i++) {
      var unidade = unidades[i];
      try {
        var ssFilha = abrirPlanilhaFilhaRapido(unidade.spreadsheetId, unidade.nome);
        var abaPedido = obterAbaPedidoFilha_(ssFilha);

        // Leitura flexível: funciona mesmo com Tabela_1 começando na linha 4+
        var lido = lerPedidosEnviadosFlexivel(abaPedido);
        var chaveUn = normalizarChaveUnidade_(unidade.nome);

        for (var r = 0; r < lido.enviados.length; r++) {
          var ped = lido.enviados[r];
          var produto = String(ped.produto || '').trim();
          var qtdTexto = String(ped.quantidade || '').trim();
          if (!produto || !qtdTexto) continue;

          var chaveProd = chaveProdutoFlex_(produto);
          if (!matriz[chaveProd]) {
            matriz[chaveProd] = { __nome: produto };
          }
          matriz[chaveProd][chaveUn] = qtdTexto;
          if (!matriz[chaveProd].__nome) matriz[chaveProd].__nome = produto;

          linhasConsolidadas.push([
            unidade.nome,
            produto,
            valorCelulaPreservandoZero_(ped.estoqueAtual),
            // Mantém a unidade informada pela unidade (ex.: 2KG, 500 g).
            qtdTexto,
            ped.observacao || '',
            ped.data || ''
          ]);
        }

        if (marcarComoConsolidado && lido.enviados.length > 0) {
          marcarEnviadosComoConsolidado_(abaPedido);
        }

        resultado.unidades++;
      } catch (e) {
        var erro = unidade.nome + ': ' + e.message;
        resultado.falhas.push(erro);
        registrarLog('CONSOLIDACAO_ERRO', erro);
      }
    }

    escreverAbaSimples_(
      ssMae,
      CONFIG.ABAS.CONSOLIDADOS,
      CONFIG.CONSOLIDADOS.HEADER,
      linhasConsolidadas
    );
    escreverResumoAgrupadoFromConsolidados_(ssMae, linhasConsolidadas, opcoesSyncAutomaticoMae_());
    atualizarIndicadorUltimaAtualizacao_(ssMae);

    resultado.linhas = contarItensConsolidados_(linhasConsolidadas);
    registrarLog(
      'CONSOLIDACAO_OK',
      resultado.linhas +
        ' pedido(s) no Resumo de ' +
        resultado.unidades +
        ' unidade(s). Marcar: ' +
        !!marcarComoConsolidado
    );

    return resultado;
  } finally {
    if (lock) {
      lock.releaseLock();
    }
  }
}

/**
 * Menu: consolidar sem alterar status nas filhas.
 */
function menuConsolidarPedidos() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Lendo pedidos das unidades...',
    'Atualizando',
    5
  );
  var resultado = consolidarPedidos(false);
  exibirResultadoConsolidacao(resultado, false);
}

/**
 * Menu: consolidar e marcar pedidos como Consolidado nas filhas.
 */
function menuConsolidarEMarcar() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Consolidar e marcar',
    'Isso irá consolidar os pedidos com status "Enviado" e alterá-los para "Consolidado" nas planilhas das unidades.\n\nDeseja continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) {
    return;
  }
  var resultado = consolidarPedidos(true);
  exibirResultadoConsolidacao(resultado, true);
}

/**
 * Entrada do trigger horário diário.
 */
function consolidarPedidosAgendado() {
  try {
    consolidarPedidos(false);
  } catch (e) {
    registrarLog('CONSOLIDACAO_AGENDADA_ERRO', e.message);
  }
}

/**
 * Marca linhas Enviado como Consolidado (cabeçalho flexível).
 * @param {Sheet} abaPedido
 */
function marcarEnviadosComoConsolidado_(abaPedido) {
  var values = abaPedido.getDataRange().getValues();
  var mapa = detectarColunasPedido_(values);
  if (mapa.colStatus < 0) return;
  for (var i = mapa.headerRow + 1; i < values.length; i++) {
    var status = String(values[i][mapa.colStatus] || '')
      .toLowerCase()
      .replace(/[^a-z]/g, '');
    if (status.indexOf('enviado') === -1) continue;
    abaPedido
      .getRange(i + 1, mapa.colStatus + 1)
      .setValue(CONFIG.VALORES.STATUS_CONSOLIDADO);
  }
}

/**
 * @param {{linhas: number, unidades: number, falhas: Array.<string>}} resultado
 * @param {boolean} marcou
 */
function exibirResultadoConsolidacao(resultado, marcou) {
  var texto =
    'Pedidos no Resumo: ' +
    resultado.linhas +
    '\nUnidades processadas: ' +
    resultado.unidades +
    '\nAbas atualizadas: Resumo + Consolidados.' +
    (marcou ? '\nStatus atualizado para Consolidado.' : '');

  if (resultado.falhas.length) {
    texto += '\n\nFalhas:\n- ' + resultado.falhas.join('\n- ');
  }

  SpreadsheetApp.getUi().alert('Consolidação concluída', texto, SpreadsheetApp.getUi().ButtonSet.OK);
}
