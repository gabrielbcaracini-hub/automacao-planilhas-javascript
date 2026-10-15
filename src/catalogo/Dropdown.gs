/**
 * Gerencia a aba Catalogo e validações da aba Pedido nas planilhas filhas.
 * Somente a coluna Status tem dropdown (Enviado → sync com a Mãe).
 */

function garantirAbaCatalogo(ssFilha) {
  return garantirAbaComCabecalho(ssFilha, CONFIG.ABAS.CATALOGO, ['Produto']);
}

function atualizarCatalogoFilha(ssFilha, nomesProdutos) {
  var aba = garantirAbaCatalogo(ssFilha);
  var lastRow = Math.max(aba.getLastRow(), CONFIG.LINHA_DADOS);
  var numLinhas = lastRow - CONFIG.LINHA_DADOS + 1;
  if (aba.getLastRow() >= CONFIG.LINHA_DADOS) {
    aba.getRange(CONFIG.LINHA_DADOS, 1, numLinhas, 1).clearContent();
  }
  if (nomesProdutos.length === 0) return;
  var valores = nomesProdutos.map(function (nome) {
    return [nome];
  });
  aba.getRange(CONFIG.LINHA_DADOS, 1, valores.length, 1).setValues(valores);
}

function limparValidacoesEntradaPedidoUnidade_(abaPedido, linhaInicio, linhaFim) {
  if (!abaPedido || linhaFim < linhaInicio) return;
  var nLinhas = linhaFim - linhaInicio + 1;
  var colsLimpar = [
    CONFIG.PEDIDO.PRODUTO,
    CONFIG.PEDIDO.QUANTIDADE,
    CONFIG.PEDIDO.OBSERVACAO
  ];
  for (var i = 0; i < colsLimpar.length; i++) {
    abaPedido.getRange(linhaInicio, colsLimpar[i], nLinhas, 1).clearDataValidations();
  }
  if (abaPedido.getLastColumn() >= 6) {
    abaPedido.getRange(linhaInicio, 6, nLinhas, 1).clearDataValidations();
  }
}

/**
 * @param {Object=} opcoes
 * @param {boolean=} opcoes.semCabecalho não chama garantirCabecalho (mais rápido em lote)
 */
function aplicarDropdownPedido(ssFilha, opcoes) {
  opcoes = opcoes || {};
  var abaPedido = ssFilha.getSheetByName(CONFIG.ABAS.PEDIDO);
  if (!abaPedido) {
    abaPedido = garantirAbaComCabecalho(ssFilha, CONFIG.ABAS.PEDIDO, obterHeaderPedidoUnidade_());
  } else if (!opcoes.semCabecalho) {
    garantirCabecalhoPedidoSemReformatar_(abaPedido);
  }

  var colStatus = CONFIG.PEDIDO.STATUS;
  var qtdLinhas = CONFIG.PEDIDO_LINHAS_VALIDACAO;
  var regraStatus = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      [
        CONFIG.VALORES.STATUS_RASCUNHO,
        CONFIG.VALORES.STATUS_ENVIADO,
        CONFIG.VALORES.STATUS_CONSOLIDADO
      ],
      true
    )
    .setAllowInvalid(false)
    .setHelpText('Selecione Enviado para enviar o pedido à Planilha Mãe.')
    .build();

  aplicarValidacaoStatusPedido_(
    abaPedido,
    regraStatus,
    colStatus,
    qtdLinhas,
    opcoes.somenteLinhasComDados === true,
    opcoes.linhaFim
  );
}

/** Status em blocos (rápido) — pula faixas de categoria. */
function aplicarValidacaoStatusPedido_(
  abaPedido,
  regraStatus,
  colStatus,
  qtdLinhas,
  somenteLinhasComDados,
  linhaFim
) {
  var inicio = CONFIG.PEDIDO.LINHA_DADOS || 3;
  var fim = linhaFim || Math.max(abaPedido.getLastRow(), inicio);
  if (somenteLinhasComDados) {
    var limiteSuperior = inicio + qtdLinhas - 1;
    if (fim > limiteSuperior) fim = limiteSuperior;
  } else {
    fim = Math.max(inicio + qtdLinhas - 1, fim);
  }
  if (fim < inicio) return;

  limparValidacoesEntradaPedidoUnidade_(abaPedido, inicio, fim);
  abaPedido.getRange(inicio, colStatus, fim - inicio + 1, 1).clearDataValidations();

  var produtos = abaPedido
    .getRange(inicio, CONFIG.PEDIDO.PRODUTO, fim - inicio + 1, 1)
    .getDisplayValues();

  var blocos = [];
  var inicioBloco = -1;

  for (var i = 0; i < produtos.length; i++) {
    var linha = inicio + i;
    var produto = String(produtos[i][0] || '').trim();

    if (ehLinhaCategoriaPedido_(produto)) {
      if (inicioBloco >= 0) {
        blocos.push({ inicio: inicioBloco, tamanho: linha - inicioBloco });
        inicioBloco = -1;
      }
      continue;
    }
    if (!produto) {
      if (inicioBloco >= 0) {
        blocos.push({ inicio: inicioBloco, tamanho: linha - inicioBloco });
        inicioBloco = -1;
      }
      continue;
    }
    if (inicioBloco < 0) inicioBloco = linha;
  }

  if (inicioBloco >= 0) {
    blocos.push({ inicio: inicioBloco, tamanho: fim - inicioBloco + 1 });
  }

  for (var b = 0; b < blocos.length; b++) {
    var bloco = blocos[b];
    if (bloco.tamanho <= 0) continue;
    abaPedido
      .getRange(bloco.inicio, colStatus, bloco.tamanho, 1)
      .setDataValidation(regraStatus);
  }

  try {
    corrigirVisualFaixasCategoria_(abaPedido);
  } catch (eCat) {}
}

function limparValidacaoObservacaoPedido_(abaPedido, linhaInicio, linhaFim) {
  limparValidacoesEntradaPedidoUnidade_(abaPedido, linhaInicio, linhaFim);
}

function garantirCabecalhoPedidoSemReformatar_(aba) {
  garantirCabecalhoPedidoUnidade_(aba);
}

function protegerCatalogo(ssFilha) {
  var aba = garantirAbaCatalogo(ssFilha);
  try {
    aba.hideSheet();
  } catch (e) {}

  var protecoes = ssFilha.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protecoes.length; i++) {
    if (protecoes[i].getDescription() === 'CatalogoSistemaCompras') return;
  }

  try {
    var protecao = aba.protect().setDescription('CatalogoSistemaCompras');
    protecao.setWarningOnly(false);
    var eu = Session.getEffectiveUser();
    protecao.addEditor(eu);
    protecao.removeEditors(
      protecao.getEditors().filter(function (editor) {
        return editor.getEmail() !== eu.getEmail();
      })
    );
    if (protecao.canDomainEdit()) protecao.setDomainEdit(false);
  } catch (e) {}
}

function sincronizarDropdownFilha(ssFilha, nomesProdutos, produtosCatalogo, opcoes) {
  atualizarCatalogoFilha(ssFilha, nomesProdutos);
  preencherPedidoComCatalogo(ssFilha, nomesProdutos, produtosCatalogo, opcoes);
  protegerCatalogo(ssFilha);
}

function corrigirValidacoesPedidoTodasUnidades_() {
  var resultado = { ok: 0, falhas: [] };
  var unidades = obterListaUnidadesMigracaoLayout_();
  for (var i = 0; i < unidades.length; i++) {
    var u = unidades[i];
    try {
      var id = String(u.spreadsheetId || '').trim();
      if (!id) id = resolverIdUnidadePeloNome_(u.nome);
      if (!id) throw new Error('ID não encontrado');
      var ssFilha = SpreadsheetApp.openById(id);
      var aba = obterAbaPedidoFilha_(ssFilha);
      aplicarDropdownPedido(ssFilha, {
        somenteLinhasComDados: true,
        semCabecalho: true,
        linhaFim: Math.max(aba.getLastRow(), CONFIG.PEDIDO.LINHA_DADOS || 3)
      });
      resultado.ok++;
    } catch (e) {
      resultado.falhas.push(u.nome + ': ' + e.message);
    }
  }
  return resultado;
}
