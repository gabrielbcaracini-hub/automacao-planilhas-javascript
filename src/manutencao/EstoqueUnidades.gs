/**
 * Funções antigas de implantação do campo Estoque Atual.
 * O campo foi removido das planilhas das unidades — não use estes menus.
 */

function menuTestarEstoqueConselhoTutelar() {
  var ui = SpreadsheetApp.getUi();
  try {
    var resultado = testarEstoqueUnidadePorNome_('UNIDADE NORTE');
    ui.alert(
      'Piloto de estoque concluído',
      'Unidade atualizada: ' + resultado.unidade +
        '\nProdutos preservados: ' + resultado.produtos +
        '\n\nPreencha o campo “Estoque Atual” em um item, marque como Enviado e confirme se ele aparece no Resumo entre Data e Quantidade.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro no piloto de estoque', e.message, ui.ButtonSet.OK);
  }
}

function testarEstoqueUnidadePorNome_(nomeUnidade) {
  var ssMae = obterPlanilhaMae();
  migrarConsolidadosParaEstoque_(ssMae);

  var id = resolverIdUnidadePeloNome_(nomeUnidade);
  if (!id) throw new Error('Não encontrei a planilha da unidade ' + nomeUnidade + '.');

  var produtos = listarProdutosAtivos();
  var nomes = produtos.map(function (p) {
    return p.produto;
  });
  if (!nomes.length) throw new Error('Não há produtos ativos para aplicar o piloto.');

  var ssFilha = SpreadsheetApp.openById(id);
  preencherPedidoComCatalogo(ssFilha, nomes, produtos);
  sincronizarUnidadePorId(id);
  registrarLog('ESTOQUE_PILOTO', nomeUnidade + ' atualizado com campo Estoque Atual.');
  return { unidade: nomeUnidade, produtos: nomes.length };
}

/**
 * Aplica o campo Estoque Atual em todas as unidades ativas, preservando os
 * pedidos existentes e reconstruindo o Resumo após a migração.
 */
function menuAplicarEstoqueTodasUnidades() {
  var ui = SpreadsheetApp.getUi();
  var confirmar = ui.alert(
    'Aplicar campo Estoque Atual',
    'O novo campo será aplicado em todas as unidades ativas.\n\n' +
      'Pedidos, categorias, cores e dropdowns existentes serão preservados. ' +
      'Ao final, o Resumo será atualizado com a coluna Estoque Atual.\n\n' +
      'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmar !== ui.Button.YES) return;

  var ssMae = obterPlanilhaMae();
  var ok = 0;
  var falhas = [];
  try {
    ssMae.toast('Preparando campo de estoque…', 'Estoque nas unidades', 60);
    migrarConsolidadosParaEstoque_(ssMae);

    var produtos = listarProdutosAtivos();
    var nomes = produtos.map(function (p) {
      return p.produto;
    });
    if (!nomes.length) throw new Error('Não há produtos ativos no catálogo.');

    var unidades = listarUnidadesAtivas();
    for (var i = 0; i < unidades.length; i++) {
      var unidade = unidades[i];
      try {
        ssMae.toast(
          'Atualizando ' + (i + 1) + '/' + unidades.length + ': ' + unidade.nome,
          'Estoque nas unidades',
          12
        );
        var ssFilha = abrirPlanilhaFilhaRapido(unidade.spreadsheetId, unidade.nome);
        preencherPedidoComCatalogo(ssFilha, nomes, produtos);
        ok++;
      } catch (eUnidade) {
        falhas.push(unidade.nome + ': ' + eUnidade.message);
      }
    }

    // Releitura completa: preenche o estoque das unidades migradas no Resumo
    // e mantém vazio para unidades que eventualmente tenham falhado.
    var resultado = consolidarPedidos(false);
    var texto =
      'Unidades atualizadas: ' + ok + '/' + unidades.length +
      '\nPedidos consolidados: ' + resultado.linhas +
      '\nFalhas de atualização: ' + falhas.length;
    if (falhas.length) texto += '\n\n- ' + falhas.slice(0, 6).join('\n- ');
    if (resultado.falhas.length) texto += '\n\nFalhas na consolidação:\n- ' + resultado.falhas.slice(0, 6).join('\n- ');
    ui.alert('Campo de estoque aplicado', texto, ui.ButtonSet.OK);
    registrarLog('ESTOQUE_TODAS_UNIDADES', texto);
  } catch (e) {
    ui.alert('Erro ao aplicar estoque', e.message, ui.ButtonSet.OK);
    try {
      registrarLog('ESTOQUE_TODAS_UNIDADES_ERRO', e.message);
    } catch (eLog) {}
  }
}

/**
 * Converte Consolidados legado de 5 para 6 colunas, sem perder pedidos.
 */
function migrarConsolidadosParaEstoque_(ssMae) {
  var aba = ssMae.getSheetByName(CONFIG.ABAS.CONSOLIDADOS);
  if (!aba || aba.getLastRow() === 0) return;

  var valores = aba.getDataRange().getValues();
  var header = valores[0] || [];
  var terceiraColuna = normalizarTexto(header[2] || '');
  if (terceiraColuna.indexOf('estoque') === 0) return;

  var migradas = [];
  for (var i = 1; i < valores.length; i++) {
    var row = valores[i];
    if (!String(row[0] || '').trim() && !String(row[1] || '').trim()) continue;
    migradas.push([
      row[0] || '',
      row[1] || '',
      '',
      valorCelulaPreservandoZero_(row[2]),
      row[3] || '',
      row[4] || ''
    ]);
  }
  escreverAbaSimples_(ssMae, CONFIG.ABAS.CONSOLIDADOS, CONFIG.CONSOLIDADOS.HEADER, migradas);
  escreverResumoAgrupadoFromConsolidados_(ssMae, migradas);
}
