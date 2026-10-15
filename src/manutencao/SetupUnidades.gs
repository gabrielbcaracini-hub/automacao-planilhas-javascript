/**
 * Importação da lista UNIDADES + criação das planilhas das unidades.
 */

/**
 * Menu: cadastra produtos UNIDADES e cria planilhas filhas.
 */
function menuImportarUnidadesCompleto() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Importar Lista UNIDADES',
    'Isso vai:\n' +
      '1) Cadastrar os ' +
      DADOS_INICIAIS.produtos.length +
      ' produtos da lista UNIDADES\n' +
      '2) Criar/atualizar planilha para cada unidade (' +
      DADOS_INICIAIS.unidades.length +
      ')\n' +
      '3) Sincronizar o catálogo nas filhas\n\n' +
      'Pode levar alguns minutos. Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) {
    return;
  }

  try {
    var resultado = importarUnidadesCompleto();
    ui.alert(
      'Importação concluída',
      'Produtos: ' +
        resultado.produtos +
        '\nUnidades: ' +
        resultado.unidades +
        '\nPlanilhas criadas: ' +
        resultado.criadas +
        '\nPlanilhas já existentes: ' +
        resultado.reutilizadas +
        '\n\nAbra a aba "Unidades" (com os links/IDs).' +
        (resultado.falhas.length
          ? '\n\nFalhas:\n- ' + resultado.falhas.join('\n- ')
          : ''),
      ui.ButtonSet.OK
    );
  } catch (e) {
    registrarLog('IMPORT_SETUP_ERRO', e.message);
    ui.alert('Erro na importação', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Executa o fluxo completo de importação UNIDADES.
 * @return {{produtos: number, unidades: number, criadas: number, reutilizadas: number, falhas: string[]}}
 */
function importarUnidadesCompleto() {
  var lock = adquirirLock('import_unidades');
  var resultado = {
    produtos: 0,
    unidades: 0,
    criadas: 0,
    reutilizadas: 0,
    falhas: []
  };

  try {
    setupPlanilhaMaeSilencioso();
    resultado.produtos = popularProdutosUnidades();
    var unids = criarOuVincularPlanilhasUnidades(resultado);
    resultado.unidades = unids;
    sincronizarCatalogoTodasUnidades(true);
    registrarLog(
      'IMPORT_SETUP_OK',
      resultado.produtos +
        ' produtos, ' +
        resultado.unidades +
        ' unidades, criadas=' +
        resultado.criadas
    );
    return resultado;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Setup da mãe sem alertas de UI.
 */
function setupPlanilhaMaeSilencioso() {
  var ss = obterPlanilhaMae();
  garantirAbaComCabecalho(ss, CONFIG.ABAS.PRODUTOS, CONFIG.PRODUTOS.HEADER);
  garantirAbaComCabecalho(ss, CONFIG.ABAS.UNIDADES, CONFIG.UNIDADES.HEADER);
  garantirAbaComCabecalho(ss, CONFIG.ABAS.CONSOLIDADOS, CONFIG.CONSOLIDADOS.HEADER);
  garantirAbaComCabecalho(ss, CONFIG.ABAS.RESUMO, CONFIG.RESUMO.HEADER);
  garantirAbaComCabecalho(ss, CONFIG.ABAS.TOTAL_PRODUTO, CONFIG.TOTAL_PRODUTO.HEADER);
  garantirAbaComCabecalho(ss, CONFIG.ABAS.LOG, CONFIG.LOG.HEADER);
  removerAbaPadraoVazia(ss);
  try {
    aplicarVisibilidadeAbasCliente_(ss);
  } catch (eOculta) {}
}

/**
 * Substitui a aba Produtos pelos itens UNIDADES.
 * @return {number}
 */
function popularProdutosUnidades() {
  var ss = obterPlanilhaMae();
  var aba = garantirAbaComCabecalho(ss, CONFIG.ABAS.PRODUTOS, CONFIG.PRODUTOS.HEADER);

  // Limpa validação antiga (Ativo estava na coluna E antes do Pregão)
  var limparCols = Math.max(aba.getLastColumn(), CONFIG.PRODUTOS.HEADER.length, 10);
  aba.getRange(CONFIG.LINHA_DADOS, 1, CONFIG.PEDIDO_LINHAS_VALIDACAO, limparCols).clearDataValidations();
  limparDadosAba(aba, limparCols);

  var linhas = DADOS_INICIAIS.produtos.map(function (p) {
    return [p.codigo, p.categoria, p.produto, '', p.pregao || '', p.ativo || CONFIG.VALORES.ATIVO_SIM];
  });

  if (linhas.length > 0) {
    aba.getRange(CONFIG.LINHA_DADOS, 1, linhas.length, CONFIG.PRODUTOS.HEADER.length).setValues(linhas);
  }

  aplicarValidacaoAtivo(aba, CONFIG.PRODUTOS.ATIVO);
  try {
    aplicarValidacaoAtivoProdutos_(aba);
  } catch (eVal) {}
  return linhas.length;
}

/**
 * Cria planilha filha para cada unidade UNIDADES (ou reutiliza se já cadastrada com mesmo nome).
 * @param {{criadas: number, reutilizadas: number, falhas: string[]}} resultado
 * @return {number} quantidade de unidades ativas cadastradas
 */
function criarOuVincularPlanilhasUnidades(resultado) {
  var ssMae = obterPlanilhaMae();
  var abaUnidades = garantirAbaComCabecalho(ssMae, CONFIG.ABAS.UNIDADES, CONFIG.UNIDADES.HEADER);

  var existentes = listarTodasUnidades();
  var porNome = {};
  for (var i = 0; i < existentes.length; i++) {
    porNome[normalizarTexto(existentes[i].nome)] = existentes[i];
  }

  var nomes = listarNomesProdutosAtivos();
  var linhasUnidades = [];

  for (var u = 0; u < DADOS_INICIAIS.unidades.length; u++) {
    var nome = DADOS_INICIAIS.unidades[u];
    var chave = normalizarTexto(nome);
    var id = '';

    try {
      if (porNome[chave] && porNome[chave].spreadsheetId) {
        id = porNome[chave].spreadsheetId;
        setupPlanilhaFilha(id, false);
        resultado.reutilizadas++;
      } else {
        var ssFilha = SpreadsheetApp.create('Pedido - ' + nome);
        id = ssFilha.getId();
      garantirAbaComCabecalho(ssFilha, CONFIG.ABAS.PEDIDO, obterHeaderPedidoUnidade_());
        garantirAbaCatalogo(ssFilha);
        sincronizarDropdownFilha(ssFilha, nomes);
        preencherPedidoComCatalogo(ssFilha, nomes);
        removerAbaPadraoVazia(ssFilha);
        resultado.criadas++;
      }

      linhasUnidades.push([nome, id, CONFIG.VALORES.ATIVO_SIM]);
    } catch (e) {
      resultado.falhas.push(nome + ': ' + e.message);
      registrarLog('CRIAR_UNIDADE_ERRO', nome + ' — ' + e.message);
    }
  }

  limparDadosAba(abaUnidades, CONFIG.UNIDADES.HEADER.length);
  if (linhasUnidades.length > 0) {
    abaUnidades
      .getRange(CONFIG.LINHA_DADOS, 1, linhasUnidades.length, CONFIG.UNIDADES.HEADER.length)
      .setValues(linhasUnidades);
  }
  aplicarValidacaoAtivo(abaUnidades, CONFIG.UNIDADES.ATIVO);

  // Garante pedido pré-preenchido também nas reutilizadas
  for (var r = 0; r < linhasUnidades.length; r++) {
    try {
      var ss = abrirPlanilhaFilha(linhasUnidades[r][1], linhasUnidades[r][0]);
      preencherPedidoComCatalogo(ss, nomes);
    } catch (e2) {
      resultado.falhas.push(linhasUnidades[r][0] + ' (pedido): ' + e2.message);
    }
  }

  return linhasUnidades.length;
}

/**
 * Preenche/atualiza a aba Pedido com produtos ATIVOS agrupados por categoria;
 * remove inativos e mantém pedidos já preenchidos.
 * @param {Spreadsheet} ssFilha
 * @param {string[]} nomesProdutos nomes dos produtos ativos
 * @param {Array.<Object>=} produtosComCategoria itens com Produto/Categoria
 * @param {Object=} opcoes
 * @param {boolean=} opcoes.aplicarVisual padrão true; false acelera sync automático
 */
function preencherPedidoComCatalogo(ssFilha, nomesProdutos, produtosComCategoria, opcoes) {
  opcoes = opcoes || {};
  if (!nomesProdutos || nomesProdutos.length === 0) {
    throw new Error(
      'Nenhum produto ATIVO encontrado na aba Produtos. Sync cancelado para não apagar as unidades.'
    );
  }

  var aba = ssFilha.getSheetByName(CONFIG.ABAS.PEDIDO);
  var lido = aba ? lerPedidosEnviadosFlexivel(aba) : { todas: [] };
  if (!aba) {
    aba = garantirAbaComCabecalho(ssFilha, CONFIG.ABAS.PEDIDO, obterHeaderPedidoUnidade_());
  } else {
    if (planilhaPedidoUsaLayoutLegadoComData_(aba)) {
      migrarPlanilhaPedidoRemoverColunaData_(aba);
    } else if (planilhaPedidoUsaLayoutComEstoque_(aba)) {
      migrarPlanilhaPedidoRemoverColunaEstoque_(aba);
    } else {
      garantirCabecalhoPedidoSemReformatar_(aba);
    }
  }

  var mapaPrev = {};
  for (var i = 0; i < lido.todas.length; i++) {
    var prod = String(lido.todas[i].produto || '').trim();
    if (!prod) continue;
    mapaPrev[normalizarTexto(prod)] = lido.todas[i];
  }

  // Os nomes são a fonte obrigatória; as categorias só organizam o visual.
  // Assim, cada produto ativo continua sendo gravado mesmo sem categoria.
  var categoriasPorNome = {};
  var itensFonte = produtosComCategoria && produtosComCategoria.length
    ? produtosComCategoria
    : listarProdutosAtivos();
  for (var f = 0; f < itensFonte.length; f++) {
    var fonte = itensFonte[f] || {};
    var nomeFonte = String(fonte.produto || '').trim();
    if (nomeFonte) {
      categoriasPorNome[normalizarTexto(nomeFonte)] =
        String(fonte.categoria || '').trim() || 'SEM CATEGORIA';
    }
  }

  var itensOrdenados = [];
  for (var n = 0; n < nomesProdutos.length; n++) {
    var nomeCatalogo = String(nomesProdutos[n] || '').trim();
    if (!nomeCatalogo) continue;
    itensOrdenados.push({
      produto: nomeCatalogo,
      categoria: categoriasPorNome[normalizarTexto(nomeCatalogo)] || 'SEM CATEGORIA'
    });
  }
  itensOrdenados.sort(function (a, b) {
    return a.categoria.localeCompare(b.categoria, 'pt-BR') ||
      a.produto.localeCompare(b.produto, 'pt-BR');
  });

  var linhas = [];
  var categoriaAtual = '';
  for (var p = 0; p < itensOrdenados.length; p++) {
    var itemCatalogo = itensOrdenados[p];
    var nome = itemCatalogo.produto;
    if (itemCatalogo.categoria !== categoriaAtual) {
      categoriaAtual = itemCatalogo.categoria;
      linhas.push(montarLinhaCategoriaPedido_(categoriaAtual));
    }

    var prev = mapaPrev[normalizarTexto(nome)];
    if (prev) {
      linhas.push(
        montarLinhaPedidoUnidade_(
          nome,
          prev.quantidade || '',
          prev.observacao || '',
          prev.status || CONFIG.VALORES.STATUS_RASCUNHO
        )
      );
    } else {
      linhas.push(montarLinhaPedidoUnidade_(nome, '', '', CONFIG.VALORES.STATUS_RASCUNHO));
    }
  }

  try {
    aplicarMesReferenciaPedidoUnidade_(aba);
  } catch (eMesRef) {}

  var headerRow = obterLinhaCabecalhoPedido_(aba);

  var lastRow = Math.max(aba.getLastRow(), headerRow);
  var lastCol = Math.max(aba.getLastColumn(), CONFIG.PEDIDO.HEADER.length);
  var linhaFimCatalogo = headerRow + linhas.length;

  var dataStartPrep = headerRow + 1;
  var linhaFimPrep = Math.max(dataStartPrep + linhas.length - 1, lastRow);
  limparValidacoesEntradaPedidoUnidade_(aba, dataStartPrep, linhaFimPrep);

  if (linhas.length > 0) {
    var dataStart = headerRow + 1;
    var numLinhas = linhas.length;
    var numProdutos = 0;
    for (var v = 0; v < linhas.length; v++) {
      var txtProd = String(linhas[v][CONFIG.PEDIDO.PRODUTO - 1] || '').trim();
      if (txtProd && txtProd.indexOf('▸ ') !== 0) {
        numProdutos++;
      }
    }
    if (numProdutos !== nomesProdutos.length) {
      throw new Error(
        'Catálogo inválido: esperado ' +
          nomesProdutos.length +
          ' produto(s), preparado ' +
          numProdutos +
          ' produto(s).'
      );
    }

    aba
      .getRange(dataStart, 1, numLinhas, CONFIG.PEDIDO.HEADER.length)
      .setValues(linhas);
    while (aba.getLastColumn() > CONFIG.PEDIDO.HEADER.length) {
      try {
        aba.deleteColumn(CONFIG.PEDIDO.HEADER.length + 1);
      } catch (eDelCol) {
        break;
      }
    }
  }
  var primeiraLinhaExtra = headerRow + 1 + linhas.length;
  if (lastRow >= primeiraLinhaExtra) {
    aba
      .getRange(
        primeiraLinhaExtra,
        1,
        lastRow - primeiraLinhaExtra + 1,
        lastCol
      )
      .clearContent();
  }

  aplicarDropdownPedido(ssFilha, { somenteLinhasComDados: true, linhaFim: linhaFimCatalogo });
  try {
    aplicarVisualLinhasPedidoAposCatalogo_(aba, headerRow, linhaFimCatalogo);
  } catch (eVisLin) {
    Logger.log('Visual linhas pedido: ' + eVisLin.message);
  }
  try {
    corrigirVisualFaixasCategoria_(aba);
  } catch (eCat) {}
  try {
    aplicarDropdownPedido(ssFilha, { somenteLinhasComDados: true, linhaFim: linhaFimCatalogo });
  } catch (eDrop2) {}
  if (opcoes.aplicarVisual !== false) {
    try {
      aplicarVisualPedidoUnidade_(aba);
    } catch (eVis) {
      Logger.log('Visual unidade: ' + eVis.message);
    }
  }
}

/**
 * Teste seguro: organiza somente Pedido - UNIDADE NORTE por categoria.
 * Se a quantidade de produtos gravada divergir da Mãe, restaura a lista anterior.
 */
function menuTestarCategoriasConselhoTutelar() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Teste de categorias — Unidade Norte',
    'Organiza apenas a planilha Pedido - UNIDADE NORTE em faixas por categoria.\n\n' +
      'A lista atual será restaurada automaticamente se a validação falhar.\n\nContinuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    var resultado = testarCategoriasUnidadePorNome_('UNIDADE NORTE');
    ui.alert(
      'Teste concluído',
      'Produtos gravados: ' + resultado.produtos + '\nCategorias: ' + resultado.categorias +
        '\n\nAtualize Pedido - UNIDADE NORTE para conferir.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert(
      'Teste revertido',
      'A lista anterior foi preservada/restaurada.\n\nDetalhe: ' + e.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Aplica o visual agrupado por categoria em todas as unidades ativas.
 * Cada unidade é validada individualmente e restaurada se houver divergência.
 */
function menuAplicarCategoriasTodasUnidades() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Aplicar categorias em todas as unidades',
    'Organiza as abas Pedido das unidades em faixas por categoria.\n\n' +
      'Cada unidade é validada antes de concluir. Se uma falhar, somente ela será restaurada.\n\n' +
      'Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var unidades = listarUnidadesAtivas();
  var ok = 0;
  var falhas = [];
  var ssMae = obterPlanilhaMae();
  for (var i = 0; i < unidades.length; i++) {
    var unidade = unidades[i];
    try {
      ssMae.toast(
        'Unidade ' + (i + 1) + '/' + unidades.length + ': ' + unidade.nome,
        'Categorias',
        15
      );
      testarCategoriasUnidadePorNome_(unidade.nome);
      ok++;
    } catch (e) {
      falhas.push(unidade.nome + ': ' + e.message);
    }
  }

  SpreadsheetApp.flush();
  registrarLog('CATEGORIAS_UNIDADES', ok + ' OK / ' + falhas.length + ' falha(s)');
  ui.alert(
    'Categorias aplicadas',
    'Unidades atualizadas: ' + ok + '/' + unidades.length +
      '\nFalhas: ' + falhas.length +
      (falhas.length ? '\n\n- ' + falhas.slice(0, 8).join('\n- ') : '') +
      '\n\nAtualize uma planilha de unidade para visualizar.',
    ui.ButtonSet.OK
  );
}

/**
 * @param {string} nomeUnidade
 * @return {{produtos:number,categorias:number}}
 */
function testarCategoriasUnidadePorNome_(nomeUnidade) {
  var id = resolverIdUnidadePeloNome_(nomeUnidade);
  if (!id) throw new Error('Não encontrei Pedido - ' + nomeUnidade);

  var produtos = listarProdutosAtivos();
  if (!produtos.length) throw new Error('A Mãe não possui produtos ativos.');

  var ssFilha = SpreadsheetApp.openById(id);
  var aba = obterAbaPedidoFilha_(ssFilha);
  var lastRowAntes = Math.max(aba.getLastRow(), 1);
  var lastColAntes = Math.max(aba.getLastColumn(), CONFIG.PEDIDO.HEADER.length);
  var valoresAntes = aba.getRange(1, 1, lastRowAntes, lastColAntes).getValues();

  try {
    var lido = lerPedidosEnviadosFlexivel(aba);
    var anteriores = {};
    for (var a = 0; a < lido.todas.length; a++) {
      anteriores[normalizarTexto(lido.todas[a].produto)] = lido.todas[a];
    }

    produtos.sort(function (x, y) {
      var cx = String(x.categoria || 'SEM CATEGORIA');
      var cy = String(y.categoria || 'SEM CATEGORIA');
      return cx.localeCompare(cy, 'pt-BR') ||
        String(x.produto).localeCompare(String(y.produto), 'pt-BR');
    });

    var linhas = [];
    var categoriaAtual = '';
    var categorias = 0;
    for (var p = 0; p < produtos.length; p++) {
      var item = produtos[p];
      var categoria = String(item.categoria || '').trim() || 'SEM CATEGORIA';
      var produto = String(item.produto || '').trim();
      if (!produto) continue;

      if (categoria !== categoriaAtual) {
        categoriaAtual = categoria;
        categorias++;
        linhas.push(montarLinhaCategoriaPedido_(categoria));
      }

      var anterior = anteriores[normalizarTexto(produto)];
      linhas.push(
        anterior
          ? montarLinhaPedidoUnidade_(
              produto,
              anterior.quantidade || '',
              anterior.observacao || '',
              anterior.status || CONFIG.VALORES.STATUS_RASCUNHO
            )
          : montarLinhaPedidoUnidade_(produto, '', '', CONFIG.VALORES.STATUS_RASCUNHO)
      );
    }

    var dataStart = obterLinhaCabecalhoPedido_(aba) + 1;
    var linhaFimPrep = Math.max(dataStart + linhas.length - 1, aba.getLastRow());
    limparValidacoesEntradaPedidoUnidade_(aba, dataStart, linhaFimPrep);

    aba.getRange(dataStart, 1, linhas.length, obterNumColsPedido_()).setValues(linhas);
    if (aba.getLastRow() > dataStart + linhas.length - 1) {
      aba
        .getRange(dataStart + linhas.length, 1, aba.getLastRow() - dataStart - linhas.length + 1, lastColAntes)
        .clearContent();
    }

    var gravados = aba.getRange(dataStart, CONFIG.PEDIDO.PRODUTO, linhas.length, 1).getDisplayValues();
    var produtosGravados = 0;
    for (var g = 0; g < gravados.length; g++) {
      var texto = String(gravados[g][0] || '').trim();
      if (texto && texto.indexOf('▸ ') !== 0) produtosGravados++;
    }
    if (produtosGravados !== produtos.length) {
      throw new Error(
        'Validação falhou: esperado ' + produtos.length + ', gravado ' + produtosGravados + '.'
      );
    }

    aplicarDropdownPedido(ssFilha);
    aplicarVisualPedidoUnidade_(aba);
    SpreadsheetApp.flush();
    registrarLog('TESTE_CATEGORIAS_UNIDADE', nomeUnidade + ': ' + produtosGravados + ' produtos');
    return { produtos: produtosGravados, categorias: categorias };
  } catch (e) {
    aba.getRange(1, 1, Math.max(aba.getLastRow(), lastRowAntes), lastColAntes).clearContent();
    aba.getRange(1, 1, lastRowAntes, lastColAntes).setValues(valoresAntes);
    SpreadsheetApp.flush();
    throw e;
  }
}

/**
 * Reescreve a aba Unidades com os IDs corretos gravados em DADOS_INICIAIS.idsPlanilhas.
 * @return {number}
 */
function aplicarIdsPlanilhasUnidades() {
  if (!DADOS_INICIAIS || !DADOS_INICIAIS.idsPlanilhas) {
    return 0;
  }
  var ss = obterPlanilhaMae();
  var aba = garantirAbaComCabecalho(ss, CONFIG.ABAS.UNIDADES, CONFIG.UNIDADES.HEADER);
  var linhas = [];
  var nomes = DADOS_INICIAIS.unidades || [];
  for (var i = 0; i < nomes.length; i++) {
    var nome = nomes[i];
    var id = DADOS_INICIAIS.idsPlanilhas[nome] || '';
    if (!id) {
      // tenta por chave normalizada
      var chaves = Object.keys(DADOS_INICIAIS.idsPlanilhas);
      for (var k = 0; k < chaves.length; k++) {
        if (normalizarTexto(chaves[k]) === normalizarTexto(nome)) {
          id = DADOS_INICIAIS.idsPlanilhas[chaves[k]];
          break;
        }
      }
    }
    if (id) {
      linhas.push([nome, id, CONFIG.VALORES.ATIVO_SIM]);
    }
  }
  limparDadosAba(aba, CONFIG.UNIDADES.HEADER.length);
  if (linhas.length > 0) {
    aba.getRange(CONFIG.LINHA_DADOS, 1, linhas.length, CONFIG.UNIDADES.HEADER.length).setValues(linhas);
  }
  aplicarValidacaoAtivo(aba, CONFIG.UNIDADES.ATIVO);
  registrarLog('IDS_APLICADOS', linhas.length + ' unidade(s)');
  return linhas.length;
}

/**
 * Menu: aplica IDs corretos das planilhas filhas.
 */
function menuAplicarIdsPlanilhas() {
  var n = aplicarIdsPlanilhasUnidades();
  SpreadsheetApp.getUi().alert(
    'IDs atualizados',
    n + ' unidades com ID corrigido na aba Unidades.\n\nAgora rode: consolidar pedidos ou aguarde o sync automático.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Diagnóstico: tenta abrir cada unidade e contar pedidos com quantidade.
 */
function menuDiagnosticoPedidos() {
  aplicarIdsPlanilhasUnidades();
  var unidades = listarUnidadesAtivas();
  var linhas = [];
  for (var i = 0; i < unidades.length; i++) {
    var u = unidades[i];
    try {
      var ss = SpreadsheetApp.openById(u.spreadsheetId);
      var aba = ss.getSheetByName(CONFIG.ABAS.PEDIDO);
      if (!aba) {
        linhas.push(u.nome + ': SEM aba Pedido');
        continue;
      }
      var dados = lerDados(aba, CONFIG.PEDIDO.HEADER.length);
      var comQtd = 0;
      var amostra = [];
      for (var r = 0; r < dados.length; r++) {
        var q = dados[r][CONFIG.PEDIDO.QUANTIDADE - 1];
        var qtxt = q === null || q === undefined ? '' : String(q).trim();
        if (!qtxt) continue;
        comQtd++;
        if (amostra.length < 3) {
          amostra.push(
            String(dados[r][CONFIG.PEDIDO.PRODUTO - 1] || '') +
              '=' +
              qtxt +
              '/' +
              String(dados[r][CONFIG.PEDIDO.STATUS - 1] || '')
          );
        }
      }
      linhas.push(u.nome + ': OK — ' + comQtd + ' pedido(s)' + (amostra.length ? ' [' + amostra.join('; ') + ']' : ''));
    } catch (e) {
      linhas.push(u.nome + ': ERRO ' + e.message);
    }
  }
  SpreadsheetApp.getUi().alert('Diagnóstico', linhas.join('\n\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Menu: corrige os Spreadsheet IDs da aba Unidades buscando no Drive
 * as planilhas "Pedido - NOME DA UNIDADE".
 */
function menuCorrigirIdsUnidades() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Corrigir IDs das unidades',
    'Isso vai localizar no Drive as planilhas "Pedido - ..." e atualizar a aba Unidades com os IDs corretos.\n\nContinuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) {
    return;
  }

  try {
    var n = corrigirIdsUnidadesDoDrive();
    ui.alert('Pronto', 'Unidades atualizadas: ' + n, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}

/**
 * @return {number}
 */
function corrigirIdsUnidadesDoDrive() {
  var mapa = {};
  var it = DriveApp.searchFiles(
    "title contains 'Pedido -' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
  );
  while (it.hasNext()) {
    var f = it.next();
    var nomeArq = f.getName();
    if (nomeArq.indexOf('(COMPLETO)') !== -1) {
      continue;
    }
    if (nomeArq.indexOf('Pedido - ') !== 0) {
      continue;
    }
    var nomeUnidade = nomeArq.substring('Pedido - '.length).trim();
    mapa[normalizarTexto(nomeUnidade)] = { nome: nomeUnidade, id: f.getId() };
  }

  var ss = obterPlanilhaMae();
  var aba = garantirAbaComCabecalho(ss, CONFIG.ABAS.UNIDADES, CONFIG.UNIDADES.HEADER);
  var atuais = listarTodasUnidades();
  var linhas = [];
  var usados = {};

  for (var i = 0; i < atuais.length; i++) {
    var u = atuais[i];
    var chave = normalizarTexto(u.nome);
    var achado = mapa[chave];
    var id = achado ? achado.id : u.spreadsheetId;
    linhas.push([u.nome, id, u.ativo ? CONFIG.VALORES.ATIVO_SIM : CONFIG.VALORES.ATIVO_NAO]);
    usados[chave] = true;
  }

  if (linhas.length === 0 && typeof DADOS_INICIAIS !== 'undefined') {
    for (var j = 0; j < DADOS_INICIAIS.unidades.length; j++) {
      var nome = DADOS_INICIAIS.unidades[j];
      var m = mapa[normalizarTexto(nome)];
      if (m) {
        linhas.push([nome, m.id, CONFIG.VALORES.ATIVO_SIM]);
      }
    }
  }

  limparDadosAba(aba, CONFIG.UNIDADES.HEADER.length);
  if (linhas.length > 0) {
    aba.getRange(CONFIG.LINHA_DADOS, 1, linhas.length, CONFIG.UNIDADES.HEADER.length).setValues(linhas);
  }
  aplicarValidacaoAtivo(aba, CONFIG.UNIDADES.ATIVO);
  registrarLog('CORRIGIR_IDS', linhas.length + ' unidade(s)');
  return linhas.length;
}
