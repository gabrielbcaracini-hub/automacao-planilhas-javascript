/**
 * Aba Resumo — agrupada por PREGÃO (como o relatório Word de compras).
 * Estrutura: Pregão → Categoria → matriz ITEM × Unidades (quantidades).
 * Consolidados continua linha a linha (dados); Resumo é visual operacional.
 */

/**
 * Monta o Resumo a partir das linhas de Consolidados.
 * @param {Spreadsheet} ssMae
 * @param {Array.<Array.<*>>} linhasCons [Unidade, Produto, Estoque, Qtd, Obs, Data]
 * @param {Object=} opcoes
 */
/**
 * Monta o Resumo a partir das linhas de Consolidados.
 * @param {Spreadsheet} ssMae
 * @param {Array.<Array.<*>>} linhasCons [Unidade, Produto, Estoque, Qtd, Obs, Data]
 * @param {Object=} opcoes
 */
function escreverResumoAgrupadoFromConsolidados_(ssMae, linhasCons, opcoes) {
  opcoes = opcoes || {};
  var dados = montarDadosResumoPorPregao_(linhasCons, ssMae);
  escreverResumoPorPregao_(ssMae, dados, opcoes);
  // Total por Produto: só valores (sem formatação / sem trigger paralelo).
  try {
    escreverTotalPorProdutoSomenteValores_(ssMae, linhasCons);
  } catch (eTot) {
    try {
      registrarLog('TOTAL_PRODUTO_LEVE_ERRO', eTot.message);
    } catch (eLog) {}
  }
}

/**
 * Catálogo produto → pregão + categoria (aba Produtos).
 * @return {Object.<string, {produto:string, categoria:string, pregao:string}>}
 */
function montarMapaCatalogoProdutoPregao_(ssMae) {
  var mapa = {};
  var produtos = [];
  try {
    if (ssMae) {
      var aba = ssMae.getSheetByName(CONFIG.ABAS.PRODUTOS);
      if (aba) {
        produtos = listarProdutosDaAba_(aba);
      }
    }
    if (!produtos.length) {
      produtos = listarTodosProdutos();
    }
  } catch (e) {
    return mapa;
  }
  for (var i = 0; i < produtos.length; i++) {
    var p = produtos[i];
    var nome = String(p.produto || '').trim();
    if (!nome) continue;
    var chave = normalizarTexto(nome);
    var pregao = String(p.pregao || '').trim();
    try {
      pregao = normalizarTextoPregaoRelatorio_(pregao);
    } catch (eN) {}
    if (!pregao) pregao = 'SEM PREGÃO CADASTRADO';
    mapa[chave] = {
      produto: nome,
      categoria: String(p.categoria || '').trim() || 'SEM CATEGORIA',
      pregao: pregao
    };
  }
  return mapa;
}

/**
 * Lê produtos direto de uma aba (sem openById).
 * @param {Sheet} aba
 * @return {Array.<Object>}
 */
function listarProdutosDaAba_(aba) {
  var map = obterMapaColunasProdutos_(aba);
  var lastRow = aba.getLastRow();
  if (lastRow < CONFIG.LINHA_DADOS) return [];
  var numCols = Math.max(aba.getLastColumn(), map.colAtivo + 1, 6);
  var linhas = aba.getRange(CONFIG.LINHA_DADOS, 1, lastRow, numCols).getValues();
  var produtos = [];
  for (var i = 0; i < linhas.length; i++) {
    var row = linhas[i];
    var produto = String(row[map.colProduto] || '').trim();
    if (!produto) continue;
    produtos.push({
      produto: produto,
      categoria: String(row[map.colCategoria] || '').trim(),
      unidadeMedida: String(row[map.colUnidade] || '').trim(),
      pregao: String(row[map.colPregao] || '').trim(),
      ativo: ehAtivo(row[map.colAtivo])
    });
  }
  return produtos;
}

/**
 * Monta o Resumo agrupado por GÊNERO → pregão → produtos (colunas = unidades).
 * Assim alimentação fica toda junta, informática junta, etc.
 * Inclui produtos ATIVOS das categorias LIBERADAS no mês; quantidades vêm dos Consolidados.
 * @param {Array.<Array.<*>>} linhasCons
 * @param {Spreadsheet=} ssMae
 * @return {{unidades:string[], categorias:Array.<Object>, nCols:number, header:string[]}}
 */
function montarDadosResumoPorPregao_(linhasCons, ssMae) {
  var produtos = [];
  try {
    if (ssMae) {
      var abaP = ssMae.getSheetByName(CONFIG.ABAS.PRODUTOS);
      if (abaP) produtos = listarProdutosDaAba_(abaP);
    }
    if (!produtos.length) produtos = listarTodosProdutos();
  } catch (eP) {}

  var unidadesAtivas = [];
  try {
    if (ssMae) {
      unidadesAtivas = listarNomesUnidadesAtivasRapido_(ssMae);
    } else {
      unidadesAtivas = listarUnidadesAtivas()
        .map(function (u) {
          return String(u.nome || '').trim();
        })
        .filter(Boolean);
    }
  } catch (eU) {}

  var mapaQtd = montarMapaQuantidadesPorProdutoUnidade_(linhasCons || []);

  var disponibilidadeCategorias = {};
  try {
    disponibilidadeCategorias = obterDisponibilidadeCategorias_();
  } catch (eDisp) {}

  var mapaUnCons = {};
  for (var c0 = 0; c0 < (linhasCons || []).length; c0++) {
    var nomeUn0 = String(linhasCons[c0][0] || '').trim();
    if (nomeUn0) mapaUnCons[normalizarTexto(nomeUn0)] = nomeUn0;
  }
  var unidades = unidadesAtivas.slice();
  var chavesUn = {};
  for (var u0 = 0; u0 < unidades.length; u0++) {
    chavesUn[normalizarTexto(unidades[u0])] = true;
  }
  var extras = Object.keys(mapaUnCons);
  for (var e = 0; e < extras.length; e++) {
    if (!chavesUn[extras[e]]) {
      unidades.push(mapaUnCons[extras[e]]);
      chavesUn[extras[e]] = true;
    }
  }
  unidades.sort(function (a, b) {
    return a.localeCompare(b, 'pt-BR');
  });

  var catalogoPorChave = {};
  for (var i = 0; i < produtos.length; i++) {
    var p = produtos[i];
    var nomeLimpoIdx = limparNomeProdutoResumo_(p.produto);
    var chave = normalizarTexto(nomeLimpoIdx || p.produto);
    var chaveFlex =
      typeof chaveProdutoFlex_ === 'function' ? chaveProdutoFlex_(nomeLimpoIdx || p.produto) : chave;
    var chaveBruta = normalizarTexto(p.produto);
    var chaveFlexBruta =
      typeof chaveProdutoFlex_ === 'function' ? chaveProdutoFlex_(p.produto) : chaveBruta;
    catalogoPorChave[chave] = p;
    if (chaveFlex) catalogoPorChave[chaveFlex] = p;
    if (chaveBruta) catalogoPorChave[chaveBruta] = p;
    if (chaveFlexBruta) catalogoPorChave[chaveFlexBruta] = p;
  }

  // categoriasMap[chaveCat] = { nome, pregoesMap[chavePregao] = { nome, produtosMap, ordemProdutos } }
  var categoriasMap = {};
  var produtosNoResumo = {};

  function garantirBlocoProduto_(pregaoTexto, categoriaTexto) {
    var pregao = '';
    try {
      pregao = normalizarTextoPregaoRelatorio_(pregaoTexto || '');
    } catch (eN) {
      pregao = String(pregaoTexto || '').trim();
    }
    if (!pregao) pregao = 'SEM PREGÃO CADASTRADO';
    var cat = String(categoriaTexto || '').trim() || 'SEM CATEGORIA';
    var chaveCat = normalizarTexto(cat);
    var chavePregao =
      typeof normalizarChavePregaoRelatorio_ === 'function'
        ? normalizarChavePregaoRelatorio_(pregao)
        : normalizarTexto(pregao);

    if (!categoriasMap[chaveCat]) {
      categoriasMap[chaveCat] = {
        nome: cat,
        pregoesMap: {},
        ordemPregoes: []
      };
    }
    var blocoC = categoriasMap[chaveCat];
    if (!blocoC.pregoesMap[chavePregao]) {
      blocoC.pregoesMap[chavePregao] = {
        nome: pregao,
        produtosMap: {},
        ordemProdutos: []
      };
      blocoC.ordemPregoes.push(chavePregao);
    }
    return blocoC.pregoesMap[chavePregao];
  }

  function marcarChaves_(chaveProd, chaveFlexP, chaveBruta, chaveFlexBruta) {
    produtosNoResumo[chaveProd] = true;
    if (chaveFlexP) produtosNoResumo[chaveFlexP] = true;
    if (chaveBruta) produtosNoResumo[chaveBruta] = true;
    if (chaveFlexBruta) produtosNoResumo[chaveFlexBruta] = true;
  }

  function obterQtds_(chaveProd, chaveFlexP, chaveBruta, chaveFlexBruta) {
    return (
      mapaQtd[chaveProd] ||
      mapaQtd[chaveFlexP] ||
      mapaQtd[chaveBruta] ||
      mapaQtd[chaveFlexBruta] ||
      {}
    );
  }

  function temAlgumaQtd_(qtds) {
    if (!qtds) return false;
    var chavesQ = Object.keys(qtds);
    for (var q = 0; q < chavesQ.length; q++) {
      if (qtds[chavesQ[q]]) return true;
    }
    return false;
  }

  for (var ip = 0; ip < produtos.length; ip++) {
    var prod = produtos[ip];
    var nomeLimpo = limparNomeProdutoResumo_(prod.produto);
    var chaveProd = normalizarTexto(nomeLimpo || prod.produto);
    var chaveFlexP =
      typeof chaveProdutoFlex_ === 'function'
        ? chaveProdutoFlex_(nomeLimpo || prod.produto)
        : chaveProd;
    var chaveBruta = normalizarTexto(prod.produto);
    var chaveFlexBruta =
      typeof chaveProdutoFlex_ === 'function' ? chaveProdutoFlex_(prod.produto) : chaveBruta;
    var qtds = obterQtds_(chaveProd, chaveFlexP, chaveBruta, chaveFlexBruta);

    // Pedido já Enviado entra sempre, mesmo se a categoria ainda não estava liberada.
    // Sem quantidade: Ativo = Não ou categoria Não neste mês fica de fora.
    var temQtd = temAlgumaQtd_(qtds);
    if (!temQtd && !prod.ativo) continue;
    if (!temQtd && !categoriaLiberadaNoMes_(prod.categoria, disponibilidadeCategorias)) continue;

    marcarChaves_(chaveProd, chaveFlexP, chaveBruta, chaveFlexBruta);

    var blocoP = garantirBlocoProduto_(prod.pregao, prod.categoria);
    var rotulo = nomeLimpo;
    try {
      var rotMontado = limparNomeProdutoResumo_(montarRotuloItemPregao_(prod));
      if (rotMontado) rotulo = rotMontado;
    } catch (eR) {}
    if (!rotulo) rotulo = limparNomeProdutoResumo_(prod.produto);
    if (!blocoP.produtosMap[chaveProd]) {
      blocoP.produtosMap[chaveProd] = { rotulo: rotulo, qtds: qtds || {} };
      blocoP.ordemProdutos.push(chaveProd);
    }
  }

  var chavesPedidos = Object.keys(mapaQtd);
  for (var k = 0; k < chavesPedidos.length; k++) {
    var chavePedido = chavesPedidos[k];
    if (produtosNoResumo[chavePedido]) continue;
    if (catalogoPorChave[chavePedido]) continue;
    if (!temAlgumaQtd_(mapaQtd[chavePedido])) continue;

    var blocoOrf = garantirBlocoProduto_('SEM PREGÃO CADASTRADO', 'SEM CATEGORIA');
    if (!blocoOrf.produtosMap[chavePedido]) {
      blocoOrf.produtosMap[chavePedido] = {
        rotulo: limparNomeProdutoResumo_(chavePedido),
        qtds: mapaQtd[chavePedido]
      };
      blocoOrf.ordemProdutos.push(chavePedido);
    }
  }

  var chavesCat = Object.keys(categoriasMap).sort(function (a, b) {
    return compararGenerosResumo_(categoriasMap[a].nome, categoriasMap[b].nome);
  });

  var categorias = [];
  for (var cc = 0; cc < chavesCat.length; cc++) {
    var bc = categoriasMap[chavesCat[cc]];
    bc.ordemPregoes.sort(function (a, b) {
      if (typeof compararPregoesOrdem_ === 'function') {
        return compararPregoesOrdem_(bc.pregoesMap[a].nome, bc.pregoesMap[b].nome);
      }
      return bc.pregoesMap[a].nome.localeCompare(bc.pregoesMap[b].nome, 'pt-BR');
    });
    var pregoesBloco = [];
    for (var pp = 0; pp < bc.ordemPregoes.length; pp++) {
      var bp = bc.pregoesMap[bc.ordemPregoes[pp]];
      if (!bp.ordemProdutos.length) continue;
      bp.ordemProdutos.sort(function (a, b) {
        return bp.produtosMap[a].rotulo.localeCompare(bp.produtosMap[b].rotulo, 'pt-BR');
      });
      bp.produtos = bp.ordemProdutos.map(function (chavePr) {
        return bp.produtosMap[chavePr];
      });
      pregoesBloco.push(bp);
    }
    if (!pregoesBloco.length) continue;
    categorias.push({
      nome: bc.nome,
      pregoes: pregoesBloco
    });
  }

  var header = ['ITEM'].concat(
    unidades.map(function (n) {
      return String(n).toUpperCase();
    })
  );

  return {
    unidades: unidades,
    categorias: categorias,
    pregoes: [], // legado — escrita usa categorias
    nCols: header.length,
    header: header
  };
}

/**
 * Ordena gêneros: alimentação primeiro, depois limpeza, informática, papelaria…
 * @param {string} nomeA
 * @param {string} nomeB
 * @return {number}
 */
function compararGenerosResumo_(nomeA, nomeB) {
  var pa = prioridadeOrdenacaoGeneroResumo_(nomeA);
  var pb = prioridadeOrdenacaoGeneroResumo_(nomeB);
  if (pa !== pb) return pa - pb;
  return String(nomeA || '').localeCompare(String(nomeB || ''), 'pt-BR');
}

/**
 * @param {string} nomeCat
 * @return {number}
 */
function prioridadeOrdenacaoGeneroResumo_(nomeCat) {
  var n = normalizarTexto(nomeCat)
    .replace(/á|à|ã|â/g, 'a')
    .replace(/é|ê/g, 'e')
    .replace(/í/g, 'i')
    .replace(/ó|ô|õ/g, 'o')
    .replace(/ú/g, 'u')
    .replace(/ç/g, 'c');
  if (n.indexOf('aliment') !== -1 || n.indexOf('hortifruti') !== -1 || n.indexOf('genero aliment') !== -1) {
    return 1;
  }
  if (n.indexOf('limpeza') !== -1 || n.indexOf('higiene') !== -1) return 2;
  if (n.indexOf('informat') !== -1) return 3;
  if (
    n.indexOf('papelaria') !== -1 ||
    n.indexOf('expediente') !== -1 ||
    n.indexOf('escrever') !== -1 ||
    n.indexOf('escolar') !== -1
  ) {
    return 4;
  }
  if (n.indexOf('sem categoria') !== -1) return 90;
  return 50;
}

/**
 * @param {Array.<Array.<*>>} linhasCons
 * @return {Array.<{nome:string, itens:Array.<Object>}>}
 * @deprecated Mantido para compatibilidade; o Resumo usa montarDadosResumoPorPregao_.
 */
function agruparPedidosPorUnidade_(linhasCons) {
  var mapa = {};
  var ordem = [];

  for (var i = 0; i < linhasCons.length; i++) {
    var row = linhasCons[i];
    var nomeUn = String(row[0] || '').trim();
    var produto = String(row[1] || '').trim();
    if (!nomeUn || !produto) continue;

    var chave = normalizarTexto(nomeUn);
    if (!mapa[chave]) {
      mapa[chave] = { nome: nomeUn, itens: [] };
      ordem.push(chave);
    }
    mapa[chave].itens.push({
      produto: produto,
      estoqueAtual: valorCelulaPreservandoZero_(row[2]),
      quantidade: row[3],
      observacao: String(row[4] || '').trim(),
      data: row[5] || ''
    });
  }

  ordem.sort(function (a, b) {
    return mapa[a].nome.localeCompare(mapa[b].nome, 'pt-BR');
  });

  var blocos = [];
  for (var o = 0; o < ordem.length; o++) {
    var bl = mapa[ordem[o]];
    bl.itens.sort(function (x, y) {
      return String(x.produto).localeCompare(String(y.produto), 'pt-BR');
    });
    blocos.push(bl);
  }
  return blocos;
}

/**
 * Garante a aba Resumo no modelo por pregão.
 * Na 1ª execução: aposenta a aba antiga pesada (renomeia, sem delete) e cria aba limpa.
 * Não exige nenhum passo manual no menu.
 * @param {Spreadsheet} ssMae
 * @return {Sheet}
 */
function obterOuCriarAbaResumoPregoes_(ssMae) {
  return garantirAbaResumoModeloPregao_(ssMae);
}

/**
 * @param {Spreadsheet} ssMae
 * @return {Sheet}
 */
function garantirAbaResumoModeloPregao_(ssMae) {
  var nome = CONFIG.ABAS.RESUMO;
  var props = PropertiesService.getScriptProperties();
  var flag = props.getProperty('RESUMO_MODELO_PREGAO_V2');
  var tempNome = (CONFIG.ABAS.RESUMO_PREGOES_TEMP || 'Resumo Pregões');

  if (flag === '1') {
    var ok = ssMae.getSheetByName(nome);
    if (ok) return ok;
  }

  var abaTemp = ssMae.getSheetByName(tempNome);
  var abaResumo = ssMae.getSheetByName(nome);

  // Se existir "Resumo Pregões" leve, promove para "Resumo".
  if (abaTemp) {
    if (abaResumo && abaResumo.getSheetId() !== abaTemp.getSheetId()) {
      aposentarAbaResumoPesada_(abaResumo);
    }
    try {
      if (abaTemp.getName() !== nome) abaTemp.setName(nome);
      props.setProperty('RESUMO_MODELO_PREGAO_V2', '1');
      return abaTemp;
    } catch (eProm) {}
  }

  abaResumo = ssMae.getSheetByName(nome);
  if (abaResumo && flag !== '1') {
    var a1Resumo = '';
    try {
      a1Resumo = String(abaResumo.getRange(1, 1).getDisplayValue() || '')
        .trim()
        .toUpperCase();
    } catch (eA1) {}
    // Já é o Resumo por pregão (ITEM × unidades). Não aposentar — isso escondia as cores.
    if (a1Resumo === 'ITEM' || a1Resumo.indexOf('ITEM') === 0) {
      props.setProperty('RESUMO_MODELO_PREGAO_V2', '1');
      return abaResumo;
    }
    var idx = Math.max(1, abaResumo.getIndex());
    aposentarAbaResumoPesada_(abaResumo);
    var ainda = ssMae.getSheetByName(nome);
    if (ainda) {
      // Renomear falhou — grava mesmo assim na aba existente (só valores).
      props.setProperty('RESUMO_MODELO_PREGAO_V2', '1');
      return ainda;
    }
    var nova = ssMae.insertSheet(nome, idx);
    props.setProperty('RESUMO_MODELO_PREGAO_V2', '1');
    return nova;
  }

  if (!abaResumo) {
    abaResumo = ssMae.insertSheet(nome);
  }
  props.setProperty('RESUMO_MODELO_PREGAO_V2', '1');
  return abaResumo;
}

/**
 * Renomeia e oculta aba antiga (sem deleteSheet — evita timeout).
 * @param {Sheet} aba
 */
function aposentarAbaResumoPesada_(aba) {
  if (!aba) return;
  var sufixo = '';
  try {
    sufixo = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  } catch (eD) {
    sufixo = String(Date.now()).slice(-8);
  }
  try {
    aba.setName('Resumo_LEGADO_' + sufixo);
  } catch (eRen) {
    try {
      aba.setName('Resumo_LEGADO_' + String(Date.now()).slice(-8));
    } catch (eRen2) {}
  }
  try {
    aba.hideSheet();
  } catch (eHide) {}
}

/**
 * Grava a aba Resumo no layout por pregão (como o Word):
 * faixa PREGÃO → faixa CATEGORIA (tipo) → linhas de produto × unidades.
 * @param {Spreadsheet} ssMae
 * @param {{unidades:string[], pregoes:Array.<Object>, nCols:number, header:string[]}} dados
 * @param {Object=} opcoes
 */
function escreverResumoPorPregao_(ssMae, dados, opcoes) {
  opcoes = opcoes || {};
  var header = dados.header || ['ITEM'];
  var nCols = Math.max(header.length, 2);
  var unidades = dados.unidades || [];
  var linhaVazia = [];
  for (var z = 0; z < nCols; z++) linhaVazia.push('');

  var linhasPlanilha = [];
  var meta = [];
  var categorias = dados.categorias || [];

  // Compatibilidade: se vier no formato antigo (pregão → categoria), converte
  if (!categorias.length && dados.pregoes && dados.pregoes.length) {
    categorias = converterPregoesParaCategoriasResumo_(dados.pregoes);
  }

  if (!categorias.length) {
    linhasPlanilha.push(
      ['(Nenhum produto ativo no catálogo — cadastre na aba Produtos)'].concat(
        linhaVazia.slice(1)
      )
    );
    meta.push({ tipo: 'vazio' });
  }

  for (var c = 0; c < categorias.length; c++) {
    var cat = categorias[c];
    linhasPlanilha.push(
      [String(cat.nome || 'SEM CATEGORIA').toUpperCase()].concat(linhaVazia.slice(1))
    );
    meta.push({ tipo: 'categoria' });

    for (var p = 0; p < (cat.pregoes || []).length; p++) {
      var bloco = cat.pregoes[p];
      if (!bloco.produtos || !bloco.produtos.length) continue;

      linhasPlanilha.push(
        [rotuloPregaoResumo_(String(bloco.nome || '').trim())].concat(linhaVazia.slice(1))
      );
      meta.push({ tipo: 'pregao' });

      for (var j = 0; j < bloco.produtos.length; j++) {
        var it = bloco.produtos[j];
        var nomeItem = limparNomeProdutoResumo_(it.rotulo || '');
        var rowItem = ['   ' + nomeItem];
        for (var u = 0; u < unidades.length; u++) {
          var q =
            (it.qtds && it.qtds[normalizarTexto(unidades[u])]) ||
            (it.quantidades && it.quantidades[normalizarTexto(unidades[u])]) ||
            '';
          rowItem.push(q);
        }
        linhasPlanilha.push(rowItem);
        meta.push({ tipo: 'item' });
      }
    }

    if (c < categorias.length - 1) {
      linhasPlanilha.push(linhaVazia.slice());
      meta.push({ tipo: 'vazio' });
    }
  }

  var aba = garantirAbaResumoModeloPregao_(ssMae);

  try {
    var filtro = aba.getFilter();
    if (filtro) filtro.remove();
  } catch (eF) {}

  try {
    var lr = Math.min(Math.max(aba.getLastRow(), 1), 1500);
    var lc = Math.min(Math.max(aba.getLastColumn(), nCols), 40);
    if (lr >= 1) {
      aba.getRange(1, 1, lr, lc).clearContent();
    }
  } catch (eClear) {}

  var totalLinhas = 1 + linhasPlanilha.length;
  var matriz = [header].concat(linhasPlanilha);
  aba.getRange(1, 1, totalLinhas, nCols).setValues(matriz);

  // Mescla faixas de pregão/categoria na largura toda (nome completo + amarelo contínuo)
  try {
    mesclarFaixasResumoPorPregao_(aba, nCols, meta);
  } catch (eMerge) {}

  // Visual do modelo (faixas) — em lote, sem pintar célula a célula
  try {
    aplicarVisualResumoPorPregaoRapido_(aba, nCols, linhasPlanilha.length, meta);
  } catch (eVis) {
    try {
      aba.setFrozenRows(1);
      aba.setFrozenColumns(1);
    } catch (eFr) {}
  }
}

/**
 * Mescla linhas de pregão e categoria em toda a largura da matriz.
 * @param {Sheet} aba
 * @param {number} nCols
 * @param {Array.<Object>} meta
 */
function mesclarFaixasResumoPorPregao_(aba, nCols, meta) {
  if (!aba || nCols < 2 || !meta || !meta.length) return;
  // Remove merges antigos só na área útil (evita conflito)
  try {
    var lr = Math.min(1 + meta.length, 800);
    var merges = aba.getRange(1, 1, lr, nCols).getMergedRanges();
    for (var i = merges.length - 1; i >= 0; i--) {
      try {
        merges[i].breakApart();
      } catch (e1) {}
    }
  } catch (e2) {}

  for (var m = 0; m < meta.length; m++) {
    if (meta[m].tipo !== 'categoria') continue;
    try {
      aba.getRange(m + 2, 1, 1, nCols).merge();
    } catch (eM) {}
  }
}

/**
 * Total por Produto só com setValues (sem clearFormat / RichText / flush).
 * @param {Spreadsheet} ssMae
 * @param {Array.<Array.<*>>} linhasCons
 */
function escreverTotalPorProdutoSomenteValores_(ssMae, linhasCons) {
  try {
    escreverTotalPorProdutoFromConsolidados_(ssMae, linhasCons || [], { modoLeve: true });
  } catch (eAgr) {}
}

/**
 * @deprecated Não usar no menu — o pull já aplica o modelo por pregão sozinho.
 * Mantido só para compatibilidade / clasp run de emergência.
 */
function prepararResumoNovoModeloZerado_(ssMae) {
  var ss = ssMae || obterPlanilhaMae();
  try {
    PropertiesService.getScriptProperties().deleteProperty('RESUMO_MODELO_PREGAO_V2');
  } catch (eP) {}
  var unidades = listarNomesUnidadesAtivasRapido_(ss);
  var header = ['ITEM'].concat(
    unidades.map(function (n) {
      return String(n).toUpperCase();
    })
  );
  if (header.length < 2) header.push('(sem unidades)');
  escreverResumoPorPregao_(
    ss,
    { unidades: unidades, pregoes: [], nCols: header.length, header: header },
    {}
  );
  return { ok: true, unidades: unidades.length };
}

/** @deprecated Removido do menu — use Puxar pedidos. */
function menuPrepararResumoNovoModelo() {
  SpreadsheetApp.getUi().alert(
    'Não é necessário',
    'O Resumo por pregão é montado automaticamente ao usar:\n\n' +
      'Sistema de Compras → Puxar pedidos das unidades',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Visual mínimo: só cabeçalho + faixas de pregão/categoria (sem milhares de células).
 * @param {Sheet} aba
 * @param {number} nCols
 * @param {number} nLinhasCorpo
 * @param {Array.<Object>} meta
 */
function aplicarVisualResumoPorPregaoRapido_(aba, nCols, nLinhasCorpo, meta) {
  // Hierarquia visual clara:
  // Pregão = azul escuro (título) | Gênero = teal | Item = fundo claro (NÃO amarelo)
  var azulPregao = '#1E3A5F';
  var tealCat = '#0F766E';
  var azulHeader = '#1D4ED8';
  var itemFundo = '#FFFFFF';
  var itemZebrado = '#F8FAFC';
  var qtdFundo = '#FFFFFF';
  var qtdZebrado = '#EEF2FF';
  var qtdComValor = '#DCFCE7';

  try {
    aba.setTabColor(CONFIG.TEMA.TAB);
  } catch (eTab) {}

  aba.setFrozenRows(1);
  aba.setFrozenColumns(1);
  try {
    aba.setColumnWidth(1, 580);
  } catch (eW) {}
  try {
    for (var c = 2; c <= Math.min(nCols, 30); c++) {
      aba.setColumnWidth(c, 110);
    }
  } catch (eW2) {}

  aba
    .getRange(1, 1, 1, nCols)
    .setBackground(azulHeader)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(8)
    .setFontFamily('Arial')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  aba.getRange(1, 1).setHorizontalAlignment('left');
  aplicarQuebraNomeUnidadesResumo_(aba, nCols);

  if (nLinhasCorpo <= 0) return;

  try {
    aba
      .getRange(2, 1, nLinhasCorpo, nCols)
      .setBackground(itemFundo)
      .setFontColor('#111827')
      .setFontFamily('Arial')
      .setFontSize(9)
      .setFontWeight('normal')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    aba.getRange(2, 1, nLinhasCorpo, 1).setHorizontalAlignment('left');
  } catch (eBase) {}

  try {
    aba
      .getRange(2, 1, nLinhasCorpo, nCols)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  } catch (eClipAll) {}

  var refsPregao = [];
  var refsCat = [];
  var refsItem = [];
  var refsItemZ = [];
  var refsQtdA = [];
  var refsQtdB = [];
  var refsVazio = [];
  var letraFim = columnToLetterResumo_(nCols);
  var zebra = 0;

  for (var i = 0; i < meta.length; i++) {
    var linha = i + 2;
    if (meta[i].tipo === 'pregao') {
      refsPregao.push('A' + linha + ':' + letraFim + linha);
    } else if (meta[i].tipo === 'categoria') {
      refsCat.push('A' + linha + ':' + letraFim + linha);
    } else if (meta[i].tipo === 'item') {
      if (zebra % 2 === 0) refsItem.push('A' + linha);
      else refsItemZ.push('A' + linha);
      if (nCols > 1) {
        var refQ = 'B' + linha + ':' + letraFim + linha;
        if (zebra % 2 === 0) refsQtdA.push(refQ);
        else refsQtdB.push(refQ);
      }
      zebra++;
    } else if (meta[i].tipo === 'vazio') {
      refsVazio.push('A' + linha + ':' + letraFim + linha);
    }
  }

  // Pregão = título (azul) — bem diferente do item
  if (refsPregao.length) {
    aba
      .getRangeList(refsPregao)
      .setBackground(azulPregao)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }

  // Gênero/categoria = teal (já gostavam)
  if (refsCat.length) {
    aba
      .getRangeList(refsCat)
      .setBackground(tealCat)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // Item = amarelo + negrito (destaque do produto, como nas unidades)
  var amareloItem = '#FFFF00';
  if (refsItem.length) {
    aba
      .getRangeList(refsItem)
      .setBackground(amareloItem)
      .setFontColor('#000000')
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
  if (refsItemZ.length) {
    aba
      .getRangeList(refsItemZ)
      .setBackground(amareloItem)
      .setFontColor('#000000')
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  if (refsQtdA.length) {
    try {
      aba
        .getRangeList(refsQtdA)
        .setBackground(qtdFundo)
        .setFontColor('#111827')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    } catch (eQa) {}
  }
  if (refsQtdB.length) {
    try {
      aba
        .getRangeList(refsQtdB)
        .setBackground(qtdZebrado)
        .setFontColor('#111827')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    } catch (eQb) {}
  }

  if (refsVazio.length) {
    try {
      aba.getRangeList(refsVazio).setBackground('#FFFFFF').setFontColor('#FFFFFF');
    } catch (eV) {}
  }

  aplicarAlturasUniformesResumo_(aba, meta);
  aplicarQuebraNomeUnidadesResumo_(aba, nCols);
  aplicarQuebraNomePregaoResumo_(aba, meta);
}

/**
 * Cabecalho das unidades: quebra linha para o nome aparecer inteiro.
 * Os quadrados de quantidade continuam CLIP (altura uniforme).
 * @param {Sheet} aba
 * @param {number} nCols
 */
function aplicarQuebraNomeUnidadesResumo_(aba, nCols) {
  if (!aba || nCols < 2) return;
  try {
    aba
      .getRange(1, 2, 1, nCols - 1)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
  } catch (eW) {}
}

/**
 * Faixas de pregao: quebra linha na coluna ITEM para o nome completo aparecer.
 * @param {Sheet} aba
 * @param {Array.<Object>} meta
 */
function aplicarQuebraNomePregaoResumo_(aba, meta) {
  if (!aba || !meta || !meta.length) return;
  var refs = [];
  for (var i = 0; i < meta.length; i++) {
    if (meta[i].tipo !== 'pregao') continue;
    refs.push('A' + (i + 2));
  }
  if (!refs.length) return;
  try {
    aba
      .getRangeList(refs)
      .setWrap(true)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('left');
  } catch (eP) {}
}

/**
 * Todas as linhas de produto/quantidade com a mesma altura media.
 * @param {Sheet} aba
 * @param {Array.<Object>} meta
 */
function aplicarAlturasUniformesResumo_(aba, meta) {
  var ALTURA_CABECALHO = 48;
  var ALTURA_PREGAO = 56;
  var ALTURA_GENERO = 28;
  var ALTURA_ITEM = 28;
  var ALTURA_VAZIO = 10;
  try {
    aba.setRowHeight(1, ALTURA_CABECALHO);
  } catch (eH) {}
  for (var i = 0; i < (meta || []).length; i++) {
    var tipo = meta[i].tipo;
    var h = ALTURA_ITEM;
    if (tipo === 'pregao') h = ALTURA_PREGAO;
    else if (tipo === 'categoria') h = ALTURA_GENERO;
    else if (tipo === 'vazio') h = ALTURA_VAZIO;
    try {
      aba.setRowHeight(i + 2, h);
    } catch (eR) {}
  }
}

/**
 * Compat: converte estrutura antiga (pregão → categorias) para gênero → pregões.
 * @param {Array.<Object>} pregoes
 * @return {Array.<Object>}
 */
function converterPregoesParaCategoriasResumo_(pregoes) {
  var mapa = {};
  for (var p = 0; p < (pregoes || []).length; p++) {
    var bloco = pregoes[p];
    for (var c = 0; c < (bloco.categorias || []).length; c++) {
      var cat = bloco.categorias[c];
      var chave = normalizarTexto(cat.nome);
      if (!mapa[chave]) {
        mapa[chave] = { nome: cat.nome, pregoes: [] };
      }
      mapa[chave].pregoes.push({
        nome: bloco.nome,
        produtos: cat.produtos || []
      });
    }
  }
  var chaves = Object.keys(mapa).sort(function (a, b) {
    return compararGenerosResumo_(mapa[a].nome, mapa[b].nome);
  });
  return chaves.map(function (k) {
    return mapa[k];
  });
}

/**
 * Rótulo curto do pregão no Resumo (evita bloco enorme confuso).
 * Ex.: ▸ PREGÃO 22/2025 — FORNECEDOR LTDA (venc. 08/07/2026)
 * @param {string} nomeCompleto
 * @return {string}
 */
function rotuloPregaoResumo_(nomeCompleto) {
  var completo = String(nomeCompleto || '').trim();
  if (!completo) return '▸ PREGÃO';
  try {
    var p = parsearPartesPregao_(completo);
    if (p && p.numero) {
      var partes = ['▸ ' + p.numero];
      if (p.fornecedor) partes.push(p.fornecedor);
      if (p.vencimento) partes.push('venc. ' + p.vencimento);
      return partes.join(' — ');
    }
  } catch (e) {}
  // Fallback: limita tamanho se não conseguir parsear
  if (completo.length > 90) completo = completo.substring(0, 87) + '…';
  return '▸ PREGÃO: ' + completo;
}

/**
 * @param {number} col 1-based
 * @return {string}
 */
function columnToLetterResumo_(col) {
  var temp = '';
  var n = col;
  while (n > 0) {
    var rem = (n - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    n = Math.floor((n - 1) / 26);
  }
  return temp;
}

/**
 * Compatibilidade: chamadas antigas redirecionam para o layout por pregão.
 * @param {Spreadsheet} ssMae
 * @param {Array.<{nome:string, itens:Array.<Object>}>} blocos
 */
function escreverResumoAgrupado_(ssMae, blocos) {
  var linhasCons = [];
  for (var b = 0; b < (blocos || []).length; b++) {
    var bloco = blocos[b];
    for (var j = 0; j < (bloco.itens || []).length; j++) {
      var it = bloco.itens[j];
      linhasCons.push([
        bloco.nome,
        it.produto,
        it.estoqueAtual,
        it.quantidade,
        it.observacao || '',
        it.data || ''
      ]);
    }
  }
  escreverResumoPorPregao_(ssMae, montarDadosResumoPorPregao_(linhasCons, ssMae), {
    recriarAba: true,
    modoRapido: true
  });
}

/**
 * Remove filtro e todas as células mescladas. Merges na faixa verde
 * impediam o setValues — o Consolidados atualizava e o Resumo ficava velho.
 * @param {Sheet} aba
 * @param {number} nCols
 */
function prepararAbaResumoParaEscrita_(aba, nCols) {
  try {
    var filtro = aba.getFilter();
    if (filtro) filtro.remove();
  } catch (eF) {}

  // Só desfaz merges na área útil (layout por pregão quase não mescla).
  try {
    var lr = Math.min(Math.max(aba.getLastRow(), 2), 800);
    var lc = Math.min(Math.max(aba.getLastColumn(), nCols || 4), 40);
    var merges = aba.getRange(1, 1, lr, lc).getMergedRanges();
    for (var i = merges.length - 1; i >= 0; i--) {
      try {
        merges[i].breakApart();
      } catch (e1) {}
    }
  } catch (e2) {}

  try {
    var maxC = Math.max(aba.getLastColumn(), nCols || 4, 26);
    aba.getRange(1, 1, 1, maxC).breakApart();
  } catch (eH) {}
}

/**
 * @param {Sheet} aba
 */
function desfazerTodosMergesAba_(aba) {
  var tentativas = [
    [Math.max(aba.getLastRow(), 2), Math.max(aba.getLastColumn(), 4)],
    [Math.min(aba.getMaxRows(), 2000), Math.min(aba.getMaxColumns(), 26)],
    [1000, 26]
  ];
  for (var t = 0; t < tentativas.length; t++) {
    try {
      var merges = aba.getRange(1, 1, tentativas[t][0], tentativas[t][1]).getMergedRanges();
      for (var i = merges.length - 1; i >= 0; i--) {
        try {
          merges[i].breakApart();
        } catch (e1) {}
      }
    } catch (e2) {}
  }
}

/**
 * @param {Sheet} aba
 * @param {number} nCols
 * @param {Array.<Array.<*>>} linhasPlanilha
 * @param {number} lastR
 * @param {number} lastC
 */
function gravarCorpoResumo_(aba, nCols, linhasPlanilha, lastR, lastC) {
  if (linhasPlanilha.length > 0) {
    var numLinhasCorpo = linhasPlanilha.length;
    aba.getRange(2, 1, numLinhasCorpo, nCols).setValues(linhasPlanilha);
    var primeiraLinhaSobra = numLinhasCorpo + 2;
    var limparAteCol = Math.max(lastC, nCols);
    if (lastR >= primeiraLinhaSobra) {
      aba
        .getRange(primeiraLinhaSobra, 1, lastR - primeiraLinhaSobra + 1, limparAteCol)
        .clearContent();
    }
    // Limpa colunas à direita do novo layout (ex.: era 4 cols, agora 15).
    if (limparAteCol > nCols) {
      aba
        .getRange(1, nCols + 1, Math.max(lastR, numLinhasCorpo + 1), limparAteCol - nCols)
        .clearContent();
    }
  } else if (lastR >= 2) {
    aba.getRange(2, 1, lastR - 1, Math.max(lastC, nCols)).clearContent();
  }
}

/**
 * Troca a aba Resumo por uma nova (NÃO usa deleteSheet — evita timeout).
 * A antiga é renomeada e ocultada.
 * @param {Spreadsheet} ssMae
 * @return {Sheet}
 */
function recriarAbaResumo_(ssMae) {
  var antiga = ssMae.getSheetByName(CONFIG.ABAS.RESUMO);
  var idx = 1;
  if (antiga) {
    idx = Math.max(1, antiga.getIndex());
    var nomeOld = 'Resumo_OLD_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    try {
      antiga.setName(nomeOld);
    } catch (eRen) {
      try {
        antiga.setName('Resumo_OLD_' + String(Date.now()).slice(-8));
      } catch (eRen2) {}
    }
    try {
      antiga.hideSheet();
    } catch (eHide) {}
  }
  return ssMae.insertSheet(CONFIG.ABAS.RESUMO, idx);
}

/**
 * Lê só os nomes das unidades ativas (aba Unidades) — sem abrir planilhas filhas.
 * @param {Spreadsheet} ssMae
 * @return {string[]}
 */
function listarNomesUnidadesAtivasRapido_(ssMae) {
  var nomes = [];
  try {
    var aba = ssMae.getSheetByName(CONFIG.ABAS.UNIDADES);
    if (!aba || aba.getLastRow() < 2) return nomes;
    var vals = aba.getRange(2, 1, aba.getLastRow(), 3).getDisplayValues();
    for (var i = 0; i < vals.length; i++) {
      var nome = String(vals[i][0] || '').trim();
      if (!nome) continue;
      if (!ehAtivo(vals[i][2])) continue;
      nomes.push(nome);
    }
  } catch (e) {}
  nomes.sort(function (a, b) {
    return a.localeCompare(b, 'pt-BR');
  });
  return nomes;
}

/**
 * Reconstrói o Resumo a partir da aba Consolidados (fonte já gravada).
 * @param {Spreadsheet=} ssMae
 * @return {number} itens
 */
function reconstruirResumoDeConsolidadosExistentes_(ssMae) {
  ssMae = ssMae || obterPlanilhaMae();
  var abaCons = ssMae.getSheetByName(CONFIG.ABAS.CONSOLIDADOS);
  var linhas = abaCons ? lerDados(abaCons, CONFIG.CONSOLIDADOS.HEADER.length) : [];
  escreverResumoAgrupadoFromConsolidados_(ssMae, linhas, opcoesSyncAutomaticoMae_());
  return contarItensConsolidados_(linhas);
}

/**
 * Prepara o visual em lote antes da escrita do Resumo.
 */
function prepararVisualBaseResumoAntesDaEscrita_(aba, nCols, nLinhasCorpo, meta) {
  if (nLinhasCorpo <= 0) return;
  var t = CONFIG.TEMA;
  var produtos = [];
  var dadosProdutos = [];
  var secoes = [];
  var vazios = [];
  var colLetraFim = String.fromCharCode(64 + nCols);

  for (var i = 0; i < meta.length; i++) {
    var linha = i + 2;
    if (meta[i].tipo === 'item') {
      produtos.push('A' + linha);
      dadosProdutos.push('B' + linha + ':' + colLetraFim + linha);
    } else if (meta[i].tipo === 'secao') {
      secoes.push('A' + linha + ':' + colLetraFim + linha);
    } else {
      vazios.push('A' + linha + ':' + colLetraFim + linha);
    }
  }

  if (produtos.length) {
    aba.getRangeList(produtos)
      .setBackground(t.COL_DESTAQUE_BG)
      .setFontColor(t.COL_DESTAQUE_FG)
      .setFontWeight('bold')
      .setHorizontalAlignment('left');
    aba.getRangeList(dadosProdutos)
      .setBackground(t.ROW_A)
      .setFontColor(t.TEXT_MUTED)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }
  if (secoes.length) {
    aba.getRangeList(secoes)
      .setBackground('#0F766E')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
  }
  if (vazios.length) {
    aba.getRangeList(vazios).setBackground(t.SHEET_BG);
  }
}

/**
 * Padroniza o nome do produto exibido no Resumo (independente de como veio da unidade).
 * @param {*} texto
 * @return {string}
 */
function padronizarNomenclaturaProdutoResumo_(texto) {
  return limparNomeProdutoResumo_(texto);
}

/**
 * Nome limpo do item no Resumo: só o produto (sem o texto do pregão grudado).
 * Ex.: "BATATA PALHA 500G PREGÃO 116/2025 - FBA..." → "BATATA PALHA 500G"
 * @param {*} texto
 * @return {string}
 */
function limparNomeProdutoResumo_(texto) {
  var t = String(texto || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!t) return '';

  // Corta tudo a partir de "PREGÃO" / "PREGAO"
  t = t.replace(/\s*[-–—]?\s*PREG[AÃÁÀÂ]O\b[\s\S]*$/i, '');
  // Corta rastro de fornecedor/data colado sem a palavra pregão
  t = t.replace(/\s*[-–—]\s*DE\s+\d{1,2}\/\d{1,2}\/\d{2,4}[\s\S]*$/i, '');
  t = t.replace(/[\s\-–—]+$/g, '').trim();
  return t.toUpperCase();
}

/**
 * @param {*} valor
 * @return {string}
 */
function formatarDataResumo_(valor) {
  if (!valor) return '';
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor)) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(valor).trim();
}

/**
 * Menu: reaplica o visual do Resumo a partir dos Consolidados.
 */
function menuAtualizarVisualResumo() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ssMae = obterPlanilhaMae();
    ssMae.toast('Atualizando Resumo…', 'Visual', 15);
    var abaCons = ssMae.getSheetByName(CONFIG.ABAS.CONSOLIDADOS);
    var linhas = abaCons ? lerDados(abaCons, CONFIG.CONSOLIDADOS.HEADER.length) : [];
    var dados = montarDadosResumoPorPregao_(linhas, ssMae);
    escreverResumoPorPregao_(ssMae, dados, { recriarAba: false, modoRapido: true });
    ui.alert(
      'Resumo atualizado',
      'Layout por pregão:\n\n' +
        '• ▸ PREGÃO\n' +
        '• Categoria\n' +
        '• ITEM × Unidades\n\n' +
        'Recarregue a aba Resumo (F5) se necessário.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}

/**
 * @param {Sheet} aba
 * @param {Object} t
 */
function reforcarCabecalhoResumo_(aba, t) {
  var colsGerais = [CONFIG.RESUMO.PRODUTO, CONFIG.RESUMO.OBSERVACAO];
  for (var i = 0; i < colsGerais.length; i++) {
    aba
      .getRange(1, colsGerais[i])
      .setBackground(t.HEADER_BG)
      .setFontColor(t.HEADER_FG)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setWrap(true);
  }
  aba
    .getRange(1, CONFIG.RESUMO.ESTOQUE_ATUAL)
    .setBackground(t.COL_ESTOQUE_HEADER_BG)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);
  aba
    .getRange(1, CONFIG.RESUMO.QUANTIDADE)
    .setBackground(t.COL_QTD_HEADER_BG)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);
}

/**
 * Estilo estoque/quantidade nas linhas de produto do Resumo.
 * @param {Sheet} aba
 * @param {Array.<Object>} meta
 * @param {Object} t
 */
function aplicarVisualEstoqueQuantidadeResumo_(aba, meta, t) {
  var refsEst = [];
  var refsQtd = [];
  var refsObs = [];
  for (var i = 0; i < meta.length; i++) {
    if (meta[i].tipo !== 'item') continue;
    var linha = i + 2;
    refsEst.push(aba.getRange(linha, CONFIG.RESUMO.ESTOQUE_ATUAL).getA1Notation());
    refsQtd.push(aba.getRange(linha, CONFIG.RESUMO.QUANTIDADE).getA1Notation());
    refsObs.push(aba.getRange(linha, CONFIG.RESUMO.OBSERVACAO).getA1Notation());
  }
  if (refsEst.length) {
    aba
      .getRangeList(refsEst)
      .setNumberFormat('0')
      .setBackground(t.COL_ESTOQUE_BG)
      .setFontColor(t.COL_ESTOQUE_FG)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }
  if (refsQtd.length) {
    aba
      .getRangeList(refsQtd)
      .setNumberFormat('0')
      .setBackground(t.COL_QTD_BG)
      .setFontColor(t.COL_QTD_FG)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }
  if (refsObs.length) {
    aba
      .getRangeList(refsObs)
      .setBackground(t.ROW_A)
      .setFontColor(t.TEXT_MUTED)
      .setFontWeight('normal')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);
  }
}

/**
 * Visual do Resumo por pregão (matriz ITEM × unidades) — em lote, sem célula a célula.
 * @param {Sheet} aba
 * @param {number} nCols
 * @param {number} nLinhasCorpo
 * @param {Array.<Object>} meta
 * @param {number=} numUnidades
 * @param {Array.<Array.<*>>=} linhasPlanilha
 */
function aplicarVisualResumoPorPregao_(aba, nCols, nLinhasCorpo, meta, numUnidades, linhasPlanilha) {
  var t = CONFIG.TEMA;
  var totalRows = 1 + nLinhasCorpo;
  var azulPregao = '#17365D';
  var tealCat = '#0F766E';
  var azulHeader = '#2563EB';
  var verdeQtd = '#ECFDF5';
  var tealQtd = '#0F766E';
  var linhaFim = 1 + nLinhasCorpo;

  try {
    aba.setTabColor(t.TAB);
  } catch (eTab) {}

  var headerRange = aba.getRange(1, 1, 1, nCols);
  headerRange
    .setBackground(azulHeader)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(8)
    .setFontFamily('Arial')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  aba.getRange(1, 1).setHorizontalAlignment('left');
  aplicarQuebraNomeUnidadesResumo_(aba, nCols);
  aba.setFrozenRows(1);
  aba.setFrozenColumns(1);

  aba.setColumnWidth(1, 580);
  if (nCols >= 2) {
    try {
      aba.setColumnWidths(2, nCols - 1, 110);
    } catch (eW) {
      for (var colU = 2; colU <= Math.min(nCols, 20); colU++) {
        aba.setColumnWidth(colU, 110);
      }
    }
  }

  if (nLinhasCorpo <= 0) return;

  var fundos = [];
  var coresFonte = [];
  var pesosFonte = [];
  var alinhamentos = [];
  var linhasPregao = [];
  var linhasCategoria = [];

  for (var i = 0; i < meta.length; i++) {
    var tipo = meta[i].tipo;
    var fundoLinha = [];
    var corLinha = [];
    var pesoLinha = [];
    var alignLinha = [];
    var rowVals = (linhasPlanilha && linhasPlanilha[i]) || [];

    if (tipo === 'pregao') {
      for (var c1 = 0; c1 < nCols; c1++) {
        fundoLinha.push(azulPregao);
        corLinha.push('#FFFFFF');
        pesoLinha.push('bold');
        alignLinha.push('left');
      }
      fundos.push(fundoLinha);
      coresFonte.push(corLinha);
      pesosFonte.push(pesoLinha);
      alinhamentos.push(alignLinha);
      linhasPregao.push(i + 2);
      continue;
    }

    if (tipo === 'categoria') {
      for (var c2 = 0; c2 < nCols; c2++) {
        fundoLinha.push(tealCat);
        corLinha.push('#FFFFFF');
        pesoLinha.push('bold');
        alignLinha.push('center');
      }
      fundos.push(fundoLinha);
      coresFonte.push(corLinha);
      pesosFonte.push(pesoLinha);
      alinhamentos.push(alignLinha);
      linhasCategoria.push(i + 2);
      continue;
    }

    if (tipo === 'vazio') {
      for (var c3 = 0; c3 < nCols; c3++) {
        fundoLinha.push(t.SHEET_BG);
        corLinha.push(t.TEXT_MUTED);
        pesoLinha.push('normal');
        alignLinha.push('left');
      }
      fundos.push(fundoLinha);
      coresFonte.push(corLinha);
      pesosFonte.push(pesoLinha);
      alinhamentos.push(alignLinha);
      continue;
    }

    // item — destaca quantidade já no lote (sem pintar célula a célula)
    var zebra = i % 2 === 0 ? '#FFFFFF' : '#F8FAFC';
    for (var c4 = 0; c4 < nCols; c4++) {
      if (c4 === 0) {
        fundoLinha.push(t.COL_DESTAQUE_BG);
        corLinha.push(t.COL_DESTAQUE_FG);
        pesoLinha.push('bold');
        alignLinha.push('left');
      } else {
        var txt = String(rowVals[c4] || '').trim();
        if (txt) {
          fundoLinha.push(verdeQtd);
          corLinha.push(tealQtd);
          pesoLinha.push('bold');
        } else {
          fundoLinha.push(zebra);
          corLinha.push('#94A3B8');
          pesoLinha.push('normal');
        }
        alignLinha.push('center');
      }
    }
    fundos.push(fundoLinha);
    coresFonte.push(corLinha);
    pesosFonte.push(pesoLinha);
    alinhamentos.push(alignLinha);
  }

  var corpo = aba.getRange(2, 1, linhaFim, nCols);
  corpo
    .setBackgrounds(fundos)
    .setFontColors(coresFonte)
    .setFontWeights(pesosFonte)
    .setHorizontalAlignments(alinhamentos)
    .setVerticalAlignment('middle')
    .setFontFamily('Arial')
    .setFontSize(9);
  try {
    corpo.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  } catch (eClip) {}

  aplicarAlturasUniformesResumo_(aba, meta);
  aplicarQuebraNomeUnidadesResumo_(aba, nCols);
  aplicarQuebraNomePregaoResumo_(aba, meta);

  try {
    corpo.setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      t.BORDER,
      SpreadsheetApp.BorderStyle.SOLID
    );
    headerRange.setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      t.BORDER,
      SpreadsheetApp.BorderStyle.SOLID
    );
  } catch (eBord) {}

  // Fundo preto só na faixa imediatamente abaixo (rápido)
  try {
    var limpezaFim = Math.min(aba.getMaxRows(), totalRows + 15);
    if (limpezaFim > totalRows) {
      aba
        .getRange(totalRows + 1, 1, limpezaFim, Math.min(nCols + 4, 30))
        .setBackground(t.SHEET_BG);
    }
  } catch (eBg) {}
}

/**
 * Compatibilidade com chamadas antigas do visual por unidade.
 */
function aplicarVisualResumoAgrupado_(aba, nCols, nLinhasCorpo, meta) {
  aplicarVisualResumoPorPregao_(aba, nCols, nLinhasCorpo, meta, Math.max(0, nCols - 1), null);
}

/**
 * Conta itens (não contando títulos de unidade).
 * @param {Array.<Array.<*>>} linhasCons
 * @return {number}
 */
function contarItensConsolidados_(linhasCons) {
  var n = 0;
  for (var i = 0; i < linhasCons.length; i++) {
    if (String(linhasCons[i][1] || '').trim()) n++;
  }
  return n;
}

/**
 * Unidades ativas que já têm pedido na Consolidados vs. pendentes.
 * @param {Spreadsheet} ssMae
 * @return {{total:number, enviaram:number, naoEnviaram:string[], nomesEnviaram:string[]}}
 */
function calcularStatusEnvioUnidades_(ssMae) {
  var unidades = listarUnidadesAtivas();
  if (unidades.length === 0 && DADOS_INICIAIS && DADOS_INICIAIS.unidades) {
    for (var d = 0; d < DADOS_INICIAIS.unidades.length; d++) {
      unidades.push({
        nome: DADOS_INICIAIS.unidades[d],
        spreadsheetId: DADOS_INICIAIS.idsPlanilhas[DADOS_INICIAIS.unidades[d]] || '',
        ativo: true
      });
    }
  }

  var linhasCons = [];
  var abaCons = ssMae.getSheetByName(CONFIG.ABAS.CONSOLIDADOS);
  if (abaCons && abaCons.getLastRow() >= CONFIG.LINHA_DADOS) {
    linhasCons = lerDados(abaCons, CONFIG.CONSOLIDADOS.HEADER.length);
  }
  var comEnvio = {};
  try {
    comEnvio = coletarChavesUnidadesQueJaPediram_(linhasCons, ssMae, false);
  } catch (eColeta) {
    for (var i = 0; i < linhasCons.length; i++) {
      var nome = String(linhasCons[i][0] || '').trim();
      if (nome) comEnvio[chaveUnidadeSituacao_(nome)] = true;
    }
  }

  var nomesEnviaram = [];
  var naoEnviaram = [];
  for (var u = 0; u < unidades.length; u++) {
    var nomeUn = String(unidades[u].nome || '').trim();
    if (!nomeUn) continue;
    if (comEnvio[chaveUnidadeSituacao_(nomeUn)]) {
      nomesEnviaram.push(nomeUn);
    } else {
      naoEnviaram.push(nomeUn);
    }
  }

  nomesEnviaram.sort(function (a, b) {
    return a.localeCompare(b, 'pt-BR');
  });
  naoEnviaram.sort(function (a, b) {
    return a.localeCompare(b, 'pt-BR');
  });

  return {
    total: unidades.length,
    enviaram: nomesEnviaram.length,
    naoEnviaram: naoEnviaram,
    nomesEnviaram: nomesEnviaram
  };
}

/**
 * Garante só 4 colunas de cabeçalho (A–D) e limpa legado Observação/Data em E1+.
 * @param {Sheet} aba
 * @param {number} nCols
 * @param {string[]} header
 */
function normalizarLinhaCabecalhoResumo_(aba, nCols, header) {
  if (!aba) return;
  header = header || obterHeaderResumo_();
  nCols = nCols || header.length;
  var limparAte = Math.max(aba.getLastColumn(), nCols + 10, 12);
  try {
    aba.getRange(1, nCols + 1, 1, limparAte).breakApart();
  } catch (eBreak) {}
  aba.getRange(1, nCols + 1, 1, limparAte).clearContent().clearFormat();
  aba.getRange(1, 1, 1, nCols).setValues([header]);
}

/**
 * @deprecated use normalizarLinhaCabecalhoResumo_
 */
function limparColunaLegadoResumo_(aba, nCols) {
  normalizarLinhaCabecalhoResumo_(aba, nCols);
}

/**
 * Conta colunas do cabeçalho atual do Resumo (ITEM + unidades).
 * @param {Sheet} aba
 * @return {number}
 */
function obterNumColsResumoDinamicamente_(aba) {
  try {
    var last = Math.max(aba.getLastColumn(), 2);
    var vals = aba.getRange(1, 1, 1, last).getDisplayValues()[0];
    var n = 0;
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i] || '').trim()) n = i + 1;
    }
    if (n >= 2) return n;
  } catch (e) {}
  return Math.max(CONFIG.RESUMO.NUM_COLS || 2, 2);
}

/**
 * Lista pendências em 1 ou 2 colunas para caber no painel.
 * @param {string[]} nomes
 * @return {string[]}
 */
function formatarLinhasPendentesResumo_(nomes) {
  if (!nomes || nomes.length === 0) return [];
  if (nomes.length <= 6) {
    var simples = [];
    for (var s = 0; s < nomes.length; s++) {
      simples.push('• ' + nomes[s]);
    }
    return simples;
  }
  var metade = Math.ceil(nomes.length / 2);
  var linhas = [];
  for (var i = 0; i < metade; i++) {
    var esq = '• ' + nomes[i];
    var dir =
      i + metade < nomes.length
        ? '     • ' + nomes[i + metade]
        : '';
    linhas.push(esq + dir);
  }
  return linhas;
}

/**
 * Painel à direita do Resumo: quantas já enviaram e quais ainda faltam.
 * Fica nas colunas F–H, fora da tabela A–D, sem mesclar a área de pedidos.
 * @param {Spreadsheet} ssMae
 */
function atualizarIndicadorUltimaAtualizacao_(ssMae) {
  var aba = ssMae.getSheetByName(CONFIG.ABAS.RESUMO);
  if (!aba) return;
  var status = calcularStatusEnvioUnidades_(ssMae);
  desenharPainelPendenciasUnidades_(aba, status);
}

/**
 * Card visual do cliente. Não mexe em A–D (pedidos).
 * @param {Sheet} aba
 * @param {{enviaram:number, total:number, naoEnviaram:string[], nomesEnviaram:string[]}} status
 */
function desenharPainelPendenciasUnidades_(aba, status) {
  status = status || { enviaram: 0, total: 0, naoEnviaram: [], nomesEnviaram: [] };
  var nCols = obterNumColsResumoDinamicamente_(aba);
  var colEspaco = nCols + 1;
  var col = nCols + 2;
  var nPainel = 3;
  var linha = 2;
  var tema = CONFIG.TEMA;
  var pendentes = status.naoEnviaram || [];
  var enviaram = status.enviaram || 0;
  var total = status.total || 0;
  var linhasPainel = 6 + Math.max(pendentes.length, 1);
  var fundoPretoLinhas = Math.min(
    aba.getMaxRows(),
    Math.max(aba.getLastRow() + 40, 80, linhasPainel + 20)
  );
  var fundoPretoCols = nPainel + 4;

  try {
    aba.getRange(1, colEspaco, 1, Math.max(aba.getLastColumn() - colEspaco + 1, fundoPretoCols))
      .breakApart()
      .clearContent()
      .setBackground(tema.SHEET_BG)
      .setFontColor(tema.SHEET_BG);
  } catch (eH) {}

  try {
    var merges = aba.getRange(2, col, fundoPretoLinhas, nPainel).getMergedRanges();
    for (var m = merges.length - 1; m >= 0; m--) {
      try {
        merges[m].breakApart();
      } catch (eM) {}
    }
  } catch (eBreak) {}

  aba.getRange(2, colEspaco, fundoPretoLinhas - 1, fundoPretoCols)
    .clearContent()
    .setBackground(tema.SHEET_BG)
    .setFontColor(tema.SHEET_BG)
    .setBorder(false, false, false, false, false, false);
  aba.setColumnWidth(colEspaco, 28);
  aba.setColumnWidth(col, 230);
  aba.setColumnWidth(col + 1, 72);
  aba.setColumnWidth(col + 2, 90);

  var valores = [];
  valores.push(['SITUAÇÃO DAS UNIDADES', '', '']);
  valores.push(['Já enviaram o pedido', String(enviaram), 'de ' + total]);
  valores.push(['Ainda não enviaram', String(pendentes.length), pendentes.length === 1 ? 'unidade' : 'unidades']);
  valores.push(['', '', '']);
  if (pendentes.length > 0) {
    valores.push(['Unidades que faltam enviar', '', '']);
    for (var i = 0; i < pendentes.length; i++) {
      valores.push(['• ' + pendentes[i], '', '']);
    }
  } else if (total > 0) {
    valores.push(['Todas as unidades já enviaram o pedido.', '', '']);
  } else {
    valores.push(['Nenhuma unidade ativa cadastrada.', '', '']);
  }

  aba.getRange(linha, col, valores.length, nPainel).setValues(valores);

  var tit = aba.getRange(linha, col, 1, nPainel);
  tit.setBackground('#0F766E')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontFamily('Arial')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  aba.setRowHeight(linha, 34);

  var ok = aba.getRange(linha + 1, col, 1, nPainel);
  ok.setBackground('#14532D')
    .setFontColor('#BBF7D0')
    .setFontFamily('Arial')
    .setVerticalAlignment('middle');
  aba.getRange(linha + 1, col).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left');
  aba.getRange(linha + 1, col + 1)
    .setFontColor('#4ADE80')
    .setFontWeight('bold')
    .setFontSize(18)
    .setHorizontalAlignment('center');
  aba.getRange(linha + 1, col + 2).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left');
  aba.setRowHeight(linha + 1, 36);

  var falta = aba.getRange(linha + 2, col, 1, nPainel);
  falta.setBackground(pendentes.length ? '#78350F' : '#14532D')
    .setFontColor(pendentes.length ? '#FDE68A' : '#BBF7D0')
    .setFontFamily('Arial')
    .setVerticalAlignment('middle');
  aba.getRange(linha + 2, col).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left');
  aba.getRange(linha + 2, col + 1)
    .setFontColor(pendentes.length ? '#FBBF24' : '#4ADE80')
    .setFontWeight('bold')
    .setFontSize(18)
    .setHorizontalAlignment('center');
  aba.getRange(linha + 2, col + 2).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left');
  aba.setRowHeight(linha + 2, 36);

  aba.getRange(linha + 3, col, 1, nPainel).setBackground(tema.SHEET_BG);

  var linhaLista = linha + 4;
  if (pendentes.length > 0) {
    aba.getRange(linhaLista, col, 1, nPainel)
      .setBackground('#92400E')
      .setFontColor('#FFFBEB')
      .setFontWeight('bold')
      .setFontSize(10)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    aba.setRowHeight(linhaLista, 28);
    var corpoLista = aba.getRange(linhaLista + 1, col, pendentes.length, nPainel);
    corpoLista
      .setBackground('#FEF3C7')
      .setFontColor('#78350F')
      .setFontFamily('Arial')
      .setFontSize(11)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    aba.getRange(linhaLista + 1, col, pendentes.length, 1).setFontWeight('bold');
    for (var p = 0; p < pendentes.length; p++) {
      aba.setRowHeight(linhaLista + 1 + p, 24);
    }
  } else {
    aba.getRange(linhaLista, col, 1, nPainel)
      .setBackground('#14532D')
      .setFontColor('#BBF7D0')
      .setFontWeight('bold')
      .setFontSize(11)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);
    aba.setRowHeight(linhaLista, 32);
  }

  aba.getRange(linha, col, valores.length, nPainel).setBorder(
    true,
    true,
    true,
    true,
    false,
    false,
    '#0F766E',
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  var headerRange = aba.getRange(1, 1, 1, nCols);
  headerRange.setBackground(tema.HEADER_BG);
  headerRange.setFontColor(tema.HEADER_FG);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setWrap(true);
  reforcarCabecalhoResumo_(aba, tema);
}
