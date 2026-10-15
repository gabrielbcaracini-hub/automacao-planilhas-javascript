/**
 * Puxada direta e diagnóstica dos pedidos Enviado → Mãe.
 */

/**
 * Menu rápido com diagnóstico detalhado.
 */
function menuPuxarPedidosRapido() {
  var ui = SpreadsheetApp.getUi();
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Corrigindo IDs e lendo unidades...', 'Sync', 20);

    // Atualiza IDs pelo Drive (nome exato), não só pelo JSON
    var resolvidos = resolverIdsUnidadesPeloDrive();
    var diag = [];
    diag.push('IDs resolvidos pelo Drive: ' + resolvidos);

    var resultado = consolidarPedidosComDiagnostico();
    diag.push('Unidades lidas: ' + resultado.unidades);
    diag.push('Pedidos Enviado gravados no Resumo: ' + resultado.linhas);
    if (resultado.amostraUnidadeCentro) {
      diag.push('');
      diag.push('--- UNIDADE CENTRO ---');
      diag.push(resultado.amostraUnidadeCentro);
    }
    if (resultado.falhas.length) {
      diag.push('');
      diag.push('Falhas:');
      diag.push('- ' + resultado.falhas.slice(0, 8).join('\n- '));
    }

    SpreadsheetApp.flush();
    ui.alert('Resultado do pull', diag.join('\n'), ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Erro no pull', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Busca no Drive cada "Pedido - NOME" e atualiza o ID na aba Unidades.
 * Não apaga a aba se a busca falhar. Se houver várias cópias, usa a mais recente.
 * @return {number}
 */
function resolverIdsUnidadesPeloDrive() {
  var ss = obterPlanilhaMae();
  var aba = garantirAbaComCabecalho(ss, CONFIG.ABAS.UNIDADES, CONFIG.UNIDADES.HEADER);
  var existentes = {};
  var atuais = listarTodasUnidades();
  for (var e = 0; e < atuais.length; e++) {
    var chaveEx = normalizarTexto(atuais[e].nome);
    if (!chaveEx) continue;
    existentes[chaveEx] = atuais[e];
  }

  var nomes = [];
  var vistos = {};
  if (DADOS_INICIAIS && DADOS_INICIAIS.unidades) {
    for (var d = 0; d < DADOS_INICIAIS.unidades.length; d++) {
      var nomeSeed = DADOS_INICIAIS.unidades[d];
      var chaveSeed = normalizarTexto(nomeSeed);
      if (!chaveSeed || vistos[chaveSeed]) continue;
      vistos[chaveSeed] = true;
      nomes.push(nomeSeed);
    }
  }
  for (var a = 0; a < atuais.length; a++) {
    var chaveAtual = normalizarTexto(atuais[a].nome);
    if (!chaveAtual || vistos[chaveAtual]) continue;
    vistos[chaveAtual] = true;
    nomes.push(atuais[a].nome);
  }

  var linhas = [];
  var ok = 0;
  for (var i = 0; i < nomes.length; i++) {
    var nome = nomes[i];
    var chave = normalizarTexto(nome);
    var anterior = existentes[chave];
    var id = encontrarIdPedidoMaisRecente_(nome);
    if (!id && DADOS_INICIAIS && DADOS_INICIAIS.idsPlanilhas && DADOS_INICIAIS.idsPlanilhas[nome]) {
      id = DADOS_INICIAIS.idsPlanilhas[nome];
    }
    if (!id && anterior && anterior.spreadsheetId) {
      id = String(anterior.spreadsheetId).trim();
    }
    var ativo =
      anterior && !anterior.ativo ? CONFIG.VALORES.ATIVO_NAO : CONFIG.VALORES.ATIVO_SIM;
    if (id) ok++;
    linhas.push([nome, id || '', ativo]);
  }

  if (!linhas.length) {
    registrarLog('RESOLVER_IDS', 'Nenhum ID resolvido — aba Unidades preservada');
    return 0;
  }

  limparDadosAba(aba, CONFIG.UNIDADES.HEADER.length);
  aba.getRange(2, 1, linhas.length, 3).setValues(linhas);
  return ok;
}

/**
 * Entre várias planilhas "Pedido - NOME", escolhe a alterada por último.
 * @param {string} nomeUnidade
 * @return {string}
 */
function encontrarIdPedidoMaisRecente_(nomeUnidade) {
  var melhorId = '';
  var melhorTs = -1;
  try {
    var it = DriveApp.getFilesByName('Pedido - ' + nomeUnidade);
    while (it.hasNext()) {
      var f = it.next();
      if (f.isTrashed()) continue;
      if (f.getName().indexOf('COMPLETO') !== -1) continue;
      var ts = 0;
      try {
        var atualizado = f.getLastUpdated();
        ts = atualizado ? atualizado.getTime() : 0;
      } catch (eTs) {}
      if (!melhorId || ts > melhorTs) {
        melhorId = f.getId();
        melhorTs = ts;
      }
    }
  } catch (e) {
    Logger.log('Drive ' + nomeUnidade + ': ' + e.message);
  }
  return melhorId;
}

/**
 * Consolidação com diagnóstico do Unidade Centro.
 * @return {{linhas:number, unidades:number, falhas:string[], amostraUnidadeCentro:string}}
 */
function consolidarPedidosComDiagnostico() {
  var resultado = {
    linhas: 0,
    unidades: 0,
    falhas: [],
    amostraUnidadeCentro: '',
    amostraLucy: '',
    diagnosticoUnidades: [],
    publicado: false
  };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(45000)) {
    resultado.falhas.push('Outra sincronização em andamento — aguarde ~30s e tente de novo.');
    return resultado;
  }

  try {
  // NUNCA openById da Mãe aqui — isso causa o timeout
  // "tempo limite do serviço Planilhas ao acessar o documento".
  var ssMae = obterPlanilhaMae();
  try {
    removerGatilhosClockPorFuncao_('processarTotalPorProdutoPendente_');
    PropertiesService.getScriptProperties().deleteProperty('TOTAL_PRODUTO_PENDENTE');
  } catch (eTrig) {}

  var unidades = listarUnidadesAtivas(ssMae);
  var linhasCons = [];
  var matriz = {};
  var unidadesLidas = {};

  for (var i = 0; i < unidades.length; i++) {
    var u = unidades[i];
    try {
      var ssFilha = SpreadsheetApp.openById(u.spreadsheetId);
      var abaPedido = obterAbaPedidoFilha_(ssFilha);

      var lido = lerPedidosEnviadosFlexivel(abaPedido);
      var enviadosUnidade = lido.enviados.length;
      unidadesLidas[normalizarTexto(u.nome)] = true;
      resultado.diagnosticoUnidades.push(
        u.nome +
          ': ' +
          enviadosUnidade +
          ' pedido(s)' +
          (enviadosUnidade
            ? ' — ' + lido.enviados[0].produto + '=' + lido.enviados[0].quantidade
            : '')
      );

      for (var r = 0; r < lido.enviados.length; r++) {
        var ped = lido.enviados[r];
        linhasCons.push([
          u.nome,
          ped.produto,
          valorCelulaPreservandoZero_(ped.estoqueAtual),
          String(ped.quantidade || '').trim(),
          ped.observacao || '',
          ped.data || ''
        ]);
        var chaveProd = chaveProdutoFlex_(ped.produto);
        var chaveUn = normalizarChaveUnidade_(u.nome);
        if (!matriz[chaveProd]) matriz[chaveProd] = { __nome: ped.produto };
        matriz[chaveProd][chaveUn] = ped.quantidade;
        matriz[chaveProd].__nome = ped.produto;
      }

      if (
        normalizarTexto(u.nome).indexOf('unidade centro') !== -1 ||
        normalizarTexto(u.nome) === 'lucy'
      ) {
        var rotulo = normalizarTexto(u.nome) === 'lucy' ? 'amostraLucy' : 'amostraUnidadeCentro';
        resultado[rotulo] =
          'Arquivo: ' +
          ssFilha.getName() +
          '\nID: ' +
          u.spreadsheetId +
          '\nEnviados: ' +
          enviadosUnidade +
          '\n' +
          lido.diagnostico;
        if (rotulo === 'amostraLucy') resultado.amostraLucy = resultado[rotulo];
      }

      resultado.unidades++;
    } catch (e) {
      resultado.falhas.push(u.nome + ': ' + e.message);
    }
  }

  if (resultado.falhas.length > 0) {
    mesclarConsolidadosDeUnidadesNaoLidas_(ssMae, linhasCons, unidadesLidas);
  }

  var podePublicar =
    resultado.unidades > 0 &&
    (linhasCons.length > 0 || resultado.falhas.length === 0);
  if (podePublicar) {
    // A lista "Pedidos das Unidades" usa as linhas desta puxada.
    // Roda antes do Resumo: se a gravação do Resumo estourar o tempo,
    // a unidade que enviou nesta leitura ainda passa para "Já pediram".
    try {
      var situacaoPull = montarSituacaoPedidosUnidades_(linhasCons, ssMae, {
        varrerFilhas: false
      });
      escreverTotalPorProdutoFromConsolidados_(ssMae, linhasCons, { modoLeve: true });
      registrarLog(
        'QUADRO_PEDIDOS',
        situacaoPull.concluidas.length +
          '/' +
          situacaoPull.total +
          ' | ainda: ' +
          (situacaoPull.pendentes.join(', ') || 'nenhuma')
      );
    } catch (eQuadro) {
      resultado.falhas.push('Quadro Pedidos das Unidades: ' + eQuadro.message);
    }
    try {
      // Só valores — aplicarModeloVisualMae_ nos Consolidados causava timeout na Mãe.
      escreverSomenteValoresSemFormatar_(
        ssMae,
        CONFIG.ABAS.CONSOLIDADOS,
        CONFIG.CONSOLIDADOS.HEADER,
        linhasCons
      );
    } catch (eCons) {
      resultado.falhas.push('Gravar Consolidados: ' + eCons.message);
    }
    try {
      escreverResumoAgrupadoFromConsolidados_(ssMae, linhasCons, {});
      resultado.publicado = true;
    } catch (eWrite) {
      resultado.falhas.push('Gravar Resumo: ' + eWrite.message);
    }
  } else {
    resultado.falhas.push(
      resultado.unidades === 0
        ? 'Nenhuma unidade foi lida; o Resumo anterior foi preservado.'
        : 'Nenhum pedido Enviado foi lido; o Resumo anterior foi preservado.'
    );
  }

  resultado.linhas = contarItensConsolidados_(linhasCons);
  try {
    registrarLog('PULL_DIAG', resultado.linhas + ' pedidos / ' + resultado.unidades + ' unidades');
  } catch (eLog) {}

  try {
    var abaShow = ssMae.getSheetByName(CONFIG.ABAS.RESUMO);
    if (abaShow) abaShow.activate();
  } catch (eShow) {}
  return resultado;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mantém no lote as linhas antigas das unidades que falharam nesta leitura.
 * @param {Spreadsheet} ssMae
 * @param {Array.<Array.<*>>} linhasCons
 * @param {Object} unidadesLidas
 */
function mesclarConsolidadosDeUnidadesNaoLidas_(ssMae, linhasCons, unidadesLidas) {
  var abaConsOld = ssMae.getSheetByName(CONFIG.ABAS.CONSOLIDADOS);
  if (!abaConsOld || abaConsOld.getLastRow() < CONFIG.LINHA_DADOS) return;
  var antigas = lerDados(abaConsOld, CONFIG.CONSOLIDADOS.HEADER.length);
  for (var a = 0; a < antigas.length; a++) {
    var chaveAntiga = normalizarTexto(antigas[a][0]);
    if (!chaveAntiga || unidadesLidas[chaveAntiga]) continue;
    linhasCons.push(antigas[a].slice(0, CONFIG.CONSOLIDADOS.HEADER.length));
  }
}

/**
 * Só reescreve a aba Unidades se faltar ID. Evita apagar IDs bons a cada pull.
 * @return {number}
 */
function completarIdsUnidadesFaltantes_() {
  var atuais = listarTodasUnidades();
  var faltando = 0;
  for (var i = 0; i < atuais.length; i++) {
    if (!String(atuais[i].spreadsheetId || '').trim()) faltando++;
  }
  if (atuais.length > 0 && faltando === 0) return atuais.length;
  return resolverIdsUnidadesPeloDrive();
}

/**
 * @param {number} column
 * @return {string}
 */
function columnToLetter_(column) {
  var temp = '';
  var col = column;
  while (col > 0) {
    var rem = (col - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    col = Math.floor((col - 1) / 26);
  }
  return temp;
}
