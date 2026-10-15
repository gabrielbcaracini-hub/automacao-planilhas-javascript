/**

 * Sincronização do catálogo da Planilha Mãe para todas as Planilhas Filhas.

 * Propagação SOMENTE manual: menu → Atualizar produtos nas unidades.

 */



/**

 * Executa sync em todas as unidades ativas.

 * @param {boolean=} silencioso

 * @param {boolean=} semLock Se true, não usa LockService (para reparo).

 * @param {Object=} opcoes

 * @param {boolean=} opcoes.modoRapido pula visual completo e pós-processamento na Mãe

 * @param {boolean=} opcoes.mostrarProgresso toast por unidade + tempo total

 * @return {{sucesso: number, falhas: Array.<string>, segundos: number}}

 */

function sincronizarCatalogoTodasUnidades(silencioso, semLock, opcoes) {

  opcoes = opcoes || {};

  var modoRapido = opcoes.modoRapido !== false;

  var mostrarProgresso = opcoes.mostrarProgresso === true;

  var inicio = Date.now();

  var lock = null;

  var resultado = { sucesso: 0, falhas: [], segundos: 0 };



  if (!semLock) {

    try {

      lock = adquirirLock('sincronizacao', 60);

    } catch (eLock) {

      var msgLock =

        eLock.message ||

        'Outra operação já está em andamento. Aguarde 1 minuto, recarregue (F5) e tente de novo.';

      registrarLog('SYNC_ERRO', msgLock);

      if (!silencioso) {

        SpreadsheetApp.getUi().alert('Não foi possível iniciar', msgLock, SpreadsheetApp.getUi().ButtonSet.OK);

      }

      resultado.falhas.push(msgLock);

      return resultado;

    }

  }



  try {

    var validacao = validarUnicidadeProdutos();

    if (!validacao.ok) {

      var msgDup =

        'Produtos ativos com nome duplicado: ' + validacao.duplicados.join(', ');

      registrarLog('SYNC_ERRO', msgDup);

      if (!silencioso) {

        SpreadsheetApp.getUi().alert('Não foi possível sincronizar.\n\n' + msgDup, SpreadsheetApp.getUi().ButtonSet.OK);

        resultado.falhas.push(msgDup);

        return resultado;

      }

      throw new Error(msgDup);

    }



    try {

      sincronizarCadastroCategorias_();

    } catch (eCategorias) {

      registrarLog('SYNC_CATEGORIAS_ERRO', eCategorias.message);

    }



    var produtosCatalogo = listarProdutosAtivos();

    var nomes = produtosCatalogo.map(function (p) {

      return p.produto;

    });

    if (nomes.length === 0) {

      var msgZero =

        'Nenhum produto ATIVO na aba Produtos. Verifique a coluna Ativo (Sim/Não).';

      registrarLog('SYNC_ERRO', msgZero);

      if (!silencioso) {

        SpreadsheetApp.getUi().alert('Não foi possível sincronizar.\n\n' + msgZero, SpreadsheetApp.getUi().ButtonSet.OK);

        resultado.falhas.push(msgZero);

        return resultado;

      }

      throw new Error(msgZero);

    }



    var unidades = listarUnidadesAtivas();

    if (unidades.length === 0) {

      var msgSemUnidade = 'Nenhuma unidade ativa cadastrada na aba Unidades.';

      registrarLog('SYNC_AVISO', msgSemUnidade);

      if (!silencioso) {

        SpreadsheetApp.getUi().alert(msgSemUnidade, SpreadsheetApp.getUi().ButtonSet.OK);

      }

      return resultado;

    }



    var ssMae = null;

    if (mostrarProgresso) {

      try {

        ssMae = obterPlanilhaMae();

        ssMae.toast('0/' + unidades.length + ' unidades — preparando…', 'Atualizando catálogo', 8);

      } catch (eToast) {}

    }



    var opcoesFilha = modoRapido ? { aplicarVisual: false } : null;



    for (var i = 0; i < unidades.length; i++) {

      var unidade = unidades[i];

      try {

        if (mostrarProgresso && ssMae) {

          ssMae.toast(

            'Unidade ' + (i + 1) + '/' + unidades.length + ': ' + unidade.nome,

            'Atualizando catálogo',

            12

          );

        }



        var ssFilha = abrirPlanilhaFilhaRapido(unidade.spreadsheetId, unidade.nome);

        sincronizarDropdownFilha(ssFilha, nomes, produtosCatalogo, opcoesFilha);

        resultado.sucesso++;

        SpreadsheetApp.flush();

      } catch (e) {

        var erro = unidade.nome + ': ' + e.message;

        resultado.falhas.push(erro);

        registrarLog('SYNC_ERRO', erro);

      }

    }



    registrarLog(

      'SYNC_RESUMO',

      resultado.sucesso + '/' + unidades.length + ' unidades — ' + nomes.length + ' produto(s) ativo(s)'

    );



    if (!modoRapido) {

      try {

        organizarProdutosPorCategoria_();

      } catch (eVis) {

        registrarLog('SYNC_VISUAL_PRODUTOS', eVis.message);

      }

    } else {

      try {

        if (!ssMae) ssMae = obterPlanilhaMae();

        aplicarVisualAbaProdutosMae_(ssMae);

      } catch (eVisRap) {

        registrarLog('SYNC_VISUAL_PRODUTOS', eVisRap.message);

      }

    }

    // Sempre: tira Ativo=Não dos Consolidados e remonta o Resumo da Mãe
    try {
      if (!ssMae) ssMae = obterPlanilhaMae();
      removerPedidosProdutosInativosNaMae_();
    } catch (eRem) {
      registrarLog('SYNC_REM_INATIVOS', eRem.message);
    }

    try {
      if (!ssMae) ssMae = obterPlanilhaMae();
      if (mostrarProgresso) {
        ssMae.toast('Atualizando Resumo (produtos ativos)…', 'Catálogo', 20);
      }
      reconstruirResumoDeConsolidadosExistentes_(ssMae);
    } catch (eResumo) {
      registrarLog('SYNC_RESUMO_ERRO', eResumo.message);
    }



    resultado.segundos = Math.max(1, Math.round((Date.now() - inicio) / 1000));



    var resumo =

      'Sincronizadas: ' +

      resultado.sucesso +

      '/' +

      unidades.length +

      ' unidades em ' +

      resultado.segundos +

      's. Produtos ativos: ' +

      nomes.length +

      '.\nResumo da Mãe atualizado com o catálogo (pregão → gênero → produtos).\n' +

      'Pedidos Enviado: use “Puxar pedidos das unidades”.';



    if (resultado.falhas.length) {

      resumo += '\nFalhas:\n- ' + resultado.falhas.join('\n- ');

    }



    if (mostrarProgresso && ssMae) {

      ssMae.toast(

        resultado.sucesso + '/' + unidades.length + ' unidades em ' + resultado.segundos + 's',

        'Catálogo atualizado',

        10

      );

    }



    if (!silencioso) {

      SpreadsheetApp.getUi().alert('Sincronização concluída', resumo, SpreadsheetApp.getUi().ButtonSet.OK);

    } else if (!mostrarProgresso) {

      avisarUsuario(resumo, 'Sincronização');

    }



    return resultado;

  } catch (eGeral) {

    registrarLog('SYNC_ERRO', eGeral.message);

    if (!silencioso) {

      SpreadsheetApp.getUi().alert(

        'Erro na sincronização',

        eGeral.message + '\n\nSe aparecer "Script em execução", aguarde 1 minuto, recarregue (F5) e tente de novo.',

        SpreadsheetApp.getUi().ButtonSet.OK

      );

    }

    throw eGeral;

  } finally {

    if (lock) {

      lock.releaseLock();

    }

  }

}



/**

 * Mantém no histórico da Mãe pedidos já enviados de produtos inativados.

 */

function removerPedidosProdutosInativosNaMae_() {
  var ssMae = obterPlanilhaMae();
  var abaCons = ssMae.getSheetByName(CONFIG.ABAS.CONSOLIDADOS);
  var linhas = [];
  if (abaCons && abaCons.getLastRow() >= CONFIG.LINHA_DADOS) {
    linhas = lerDados(abaCons, CONFIG.CONSOLIDADOS.HEADER.length) || [];
  }

  var ativos = {};
  var todos = [];
  try {
    todos = listarTodosProdutos();
  } catch (eT) {}

  for (var i = 0; i < todos.length; i++) {
    var p = todos[i];
    if (!p.ativo) continue;
    var nomeLimpo = limparNomeProdutoResumo_(p.produto);
    var chave = normalizarTexto(nomeLimpo || p.produto);
    var chaveFlex =
      typeof chaveProdutoFlex_ === 'function'
        ? chaveProdutoFlex_(nomeLimpo || p.produto)
        : chave;
    var chaveBruta = normalizarTexto(p.produto);
    var chaveFlexBruta =
      typeof chaveProdutoFlex_ === 'function' ? chaveProdutoFlex_(p.produto) : chaveBruta;
    ativos[chave] = true;
    if (chaveFlex) ativos[chaveFlex] = true;
    if (chaveBruta) ativos[chaveBruta] = true;
    if (chaveFlexBruta) ativos[chaveFlexBruta] = true;
  }

  var filtradas = [];
  for (var r = 0; r < linhas.length; r++) {
    var row = linhas[r];
    var prod = limparNomeProdutoResumo_(row[1]);
    if (!prod) continue;
    var k = normalizarTexto(prod);
    var kf = typeof chaveProdutoFlex_ === 'function' ? chaveProdutoFlex_(prod) : k;
    if (ativos[k] || ativos[kf]) {
      filtradas.push(row);
    }
  }

  try {
    escreverSomenteValoresSemFormatar_(
      ssMae,
      CONFIG.ABAS.CONSOLIDADOS,
      CONFIG.CONSOLIDADOS.HEADER,
      filtradas
    );
  } catch (eWrite) {
    escreverAbaSimples_(ssMae, CONFIG.ABAS.CONSOLIDADOS, CONFIG.CONSOLIDADOS.HEADER, filtradas);
  }

  return filtradas;
}



/**

 * Menu: única forma de propagar produtos novos/inativados para as unidades.

 */

function menuAtualizarProdutosUnidadesAgora() {

  var ui = SpreadsheetApp.getUi();

  var resp = ui.alert(

    'Atualizar produtos nas unidades',

    'Isso vai enviar para TODAS as unidades só o que está LIBERADO neste mês.\n\n' +

      '• Categoria = Sim + produto Ativo = Sim → entra nas unidades e no Resumo.\n' +

      '• Categoria = Não → sai das unidades e do Resumo (o catálogo na aba Produtos permanece).\n' +

      '• Produto Ativo = Não → sai das unidades e do Resumo.\n\n' +

      'Assim o mês fica leve: se só alimentos estiver Sim, ninguém vê os 420 itens.\n\n' +

      'Pode levar 1–3 minutos. Você verá o progresso na tela.\n\nContinuar?',

    ui.ButtonSet.YES_NO

  );

  if (resp !== ui.Button.YES) {

    return;

  }



  try {

    removerGatilhosCatalogoAutomatico_();

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    ss.toast('Iniciando atualização nas unidades…', 'Catálogo', 5);

    return sincronizarCatalogoTodasUnidades(false, false, {

      modoRapido: true,

      mostrarProgresso: true

    });

  } catch (e) {

    ui.alert(

      'Erro ao atualizar produtos',

      (e.message || String(e)) +

        '\n\nDica: se aparecer "Script em execução", aguarde 1 minuto, recarregue (F5) e tente de novo.',

      ui.ButtonSet.OK

    );

  }

}



/**

 * Menu: sync completo (reorganiza Produtos na Mãe + Resumo).

 */

function menuSincronizarCatalogo() {

  var ui = SpreadsheetApp.getUi();

  var resp = ui.alert(

    'Sincronizar catálogo completo',

    'Igual a “Atualizar produtos nas unidades”, mas também reorganiza\n' +

      'a aba Produtos por categoria e atualiza o Resumo.\n\n' +

      'Use quando mudar categorias ou quiser o layout completo.\n\nContinuar?',

    ui.ButtonSet.YES_NO

  );

  if (resp !== ui.Button.YES) {

    return;

  }



  try {

    return sincronizarCatalogoTodasUnidades(false, false, {

      modoRapido: false,

      mostrarProgresso: true

    });

  } catch (e) {

    ui.alert('Erro na sincronização', e.message || String(e), ui.ButtonSet.OK);

  }

}



/**

 * Diagnóstico remoto do catálogo ativo.

 * @return {Object}

 */

function diagnosticarCatalogoAtivo_() {

  var produtos = listarProdutosAtivos();

  return {

    total: produtos.length,

    amostra: produtos.slice(0, 5).map(function (p) {

      return {

        produto: p.produto,

        categoria: p.categoria,

        ativo: p.ativo

      };

    })

  };

}



/**

 * Remove gatilhos legados de sync automático de catálogo (não usamos mais).

 */

function removerGatilhosCatalogoAutomatico_() {

  var funcoes = {

    onEditProdutosInstalavel: true,

    processarCatalogoPendente: true,

    processarCatalogoAtrasado_: true,

    executarCatalogoAtrasado_: true

  };

  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {

    var fn = triggers[i].getHandlerFunction();

    if (funcoes[fn]) {

      ScriptApp.deleteTrigger(triggers[i]);

    }

  }



  try {

    var props = PropertiesService.getScriptProperties();

    props.deleteProperty('CATALOGO_PENDENTE');

    props.deleteProperty('CATALOGO_PENDENTE_TS');

  } catch (e) {}

}



/** @deprecated Sync automático desativado — mantido só para limpar gatilhos antigos. */

function onEditProdutosInstalavel() {}



/** @deprecated Sync automático desativado. */

function processarCatalogoPendente() {}



/** @deprecated Sync automático desativado. */

function processarCatalogoAtrasado_() {}



/** @deprecated Sync automático desativado. */

function executarCatalogoAtrasado_() {}



/** @deprecated Use removerGatilhosCatalogoAutomatico_ */

function garantirTriggerCatalogoPendente() {

  removerGatilhosCatalogoAutomatico_();

}



/** @deprecated Use removerGatilhosCatalogoAutomatico_ */

function garantirTriggerCatalogoAtrasado_() {

  removerGatilhosCatalogoAutomatico_();

}


