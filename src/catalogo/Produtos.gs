/**
 * Leitura e validação do catálogo de produtos (Planilha Mãe).
 * Detecta colunas mesmo com "#" extra ou cabeçalho customizado.
 */

/**
 * @param {Sheet} aba
 * @return {Object} colunas 0-based (colProduto, colAtivo, …)
 */
function obterMapaColunasProdutos_(aba) {
  var lastCol = Math.max(aba.getLastColumn(), CONFIG.PRODUTOS.HEADER.length, 8);
  var headers = aba.getRange(1, 1, 1, lastCol).getValues()[0];

  var map = {
    colCodigo: CONFIG.PRODUTOS.CODIGO - 1,
    colCategoria: CONFIG.PRODUTOS.CATEGORIA - 1,
    colProduto: CONFIG.PRODUTOS.PRODUTO - 1,
    colUnidade: CONFIG.PRODUTOS.UNIDADE_MEDIDA - 1,
    colPregao: CONFIG.PRODUTOS.PREGAO - 1,
    colAtivo: CONFIG.PRODUTOS.ATIVO - 1
  };

  for (var c = 0; c < headers.length; c++) {
    var h = normalizarHeaderProduto_(headers[c]);
    if (h === 'produto') map.colProduto = c;
    else if (h === 'ativo') map.colAtivo = c;
    else if (h === 'codigo') map.colCodigo = c;
    else if (h === 'categoria') map.colCategoria = c;
    else if (h === 'unidade') map.colUnidade = c;
    else if (h.indexOf('preg') === 0) map.colPregao = c;
  }

  return map;
}

/**
 * @param {*} valor
 * @return {string}
 */
function normalizarHeaderProduto_(valor) {
  return String(valor || '')
    .toLowerCase()
    .replace(/á|à|ã|â/g, 'a')
    .replace(/é|ê/g, 'e')
    .replace(/í/g, 'i')
    .replace(/ó|ô|õ/g, 'o')
    .replace(/ú/g, 'u')
    .replace(/ç/g, 'c')
    .replace(/[^a-z]/g, '');
}

/**
 * Aplica Sim/Não na coluna Ativo detectada.
 * @param {Sheet} aba
 */
function aplicarValidacaoAtivoProdutos_(aba) {
  var map = obterMapaColunasProdutos_(aba);
  var lastRow = aba.getLastRow();
  if (lastRow < CONFIG.LINHA_DADOS) return;

  var regra = SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG.VALORES.ATIVO_SIM, CONFIG.VALORES.ATIVO_NAO], true)
    .setAllowInvalid(false)
    .build();
  var valores = aba
    .getRange(CONFIG.LINHA_DADOS, 1, lastRow - CONFIG.LINHA_DADOS + 1, 1)
    .getDisplayValues();

  for (var i = 0; i < valores.length; i++) {
    if (ehLinhaTituloCategoriaProdutos_(valores[i])) continue;
    aba
      .getRange(CONFIG.LINHA_DADOS + i, map.colAtivo + 1)
      .clearDataValidations()
      .setDataValidation(regra);
  }
}

/**
 * Aplica Sim/Não somente na célula Ativo de uma nova linha de produto.
 * Evita alterar as faixas de categoria, que são células mescladas.
 */
function aplicarValidacaoAtivoLinhaProduto_(aba, linha, map) {
  var regra = SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG.VALORES.ATIVO_SIM, CONFIG.VALORES.ATIVO_NAO], true)
    .setAllowInvalid(false)
    .build();
  aba.getRange(linha, map.colAtivo + 1).clearDataValidations().setDataValidation(regra);
}

/**
 * Identifica uma linha-título de categoria criada pelo sistema.
 * @param {Array.<*>} row
 * @return {boolean}
 */
function ehLinhaTituloCategoriaProdutos_(row) {
  return String(row[0] || '').trim().indexOf('▸ ') === 0;
}

/**
 * Extrai o nome da categoria de uma linha-título.
 * @param {Array.<*>} row
 * @return {string}
 */
function obterCategoriaDaLinhaTituloProdutos_(row) {
  return String(row[0] || '').replace(/^▸\s*/, '').trim();
}

/**
 * Reorganiza a aba Produtos em blocos visuais por categoria.
 * Mantém as colunas de cada item (inclusive Pregão) para não afetar o catálogo.
 * @param {Spreadsheet=} ssMae
 * @return {number} quantidade de produtos organizados
 */
function organizarProdutosPorCategoria_(ssMae) {
  ssMae = ssMae || obterPlanilhaMae();
  var aba = obterAbaObrigatoria(ssMae, CONFIG.ABAS.PRODUTOS);
  var map = obterMapaColunasProdutos_(aba);
  var lastRow = aba.getLastRow();
  var nCols = Math.max(
    aba.getLastColumn(),
    CONFIG.PRODUTOS.HEADER.length,
    map.colAtivo + 1,
    6
  );
  var cabecalho = aba.getRange(1, 1, 1, nCols).getValues()[0];

  if (lastRow < CONFIG.LINHA_DADOS) return 0;

  var nRows = lastRow - CONFIG.LINHA_DADOS + 1;
  var valores = aba.getRange(CONFIG.LINHA_DADOS, 1, nRows, nCols).getValues();
  var produtos = [];

  for (var i = 0; i < valores.length; i++) {
    var row = valores[i];
    if (ehLinhaTituloCategoriaProdutos_(row)) continue;

    var produto = String(row[map.colProduto] || '').trim();
    if (!produto) continue;

    var categoria = String(row[map.colCategoria] || '').trim() || 'SEM CATEGORIA';
    var pregao = String(row[map.colPregao] || '').trim();
    produtos.push({
      categoria: categoria,
      produto: produto,
      pregao: pregao,
      valores: row.slice(0, nCols)
    });
  }

  produtos.sort(function (a, b) {
    var porCategoria = a.categoria.localeCompare(b.categoria, 'pt-BR');
    if (porCategoria) return porCategoria;
    var porPregao =
      typeof compararPregoesOrdem_ === 'function'
        ? compararPregoesOrdem_(a.pregao, b.pregao)
        : String(a.pregao || '').localeCompare(String(b.pregao || ''), 'pt-BR');
    if (porPregao) return porPregao;
    return a.produto.localeCompare(b.produto, 'pt-BR');
  });

  var saida = [];
  var categorias = [];
  var categoriaAtual = '';
  for (var p = 0; p < produtos.length; p++) {
    var item = produtos[p];
    if (item.categoria !== categoriaAtual) {
      categoriaAtual = item.categoria;
      var titulo = [];
      for (var c = 0; c < nCols; c++) titulo.push('');
      titulo[0] = '▸ ' + categoriaAtual;
      saida.push(titulo);
      categorias.push(saida.length);
    }
    saida.push(item.valores);
  }

  try {
    aba.getRange(CONFIG.LINHA_DADOS, 1, nRows, nCols).breakApart();
    aba.getRange(CONFIG.LINHA_DADOS, 1, nRows, nCols).clearContent().clearFormat();
    if (saida.length > 0) {
      aba.getRange(CONFIG.LINHA_DADOS, 1, saida.length, nCols).setValues(saida);
    }
    // Remove fisicamente as linhas vazias de testes/cadastros incompletos.
    // Os itens válidos e os títulos já foram reescritos de forma compacta.
    var linhasVazias = nRows - saida.length;
    if (linhasVazias > 0) {
      aba.deleteRows(CONFIG.LINHA_DADOS + saida.length, linhasVazias);
    }
    aplicarValidacaoAtivoProdutos_(aba);
    aplicarFaixasCategoriaProdutos_(aba, categorias, nCols);
  } catch (eTabela) {
    // O recurso "Tabela_1" do Google não permite mesclagem. Recria somente
    // a aba Produtos como intervalo comum e preserva todos os valores.
    aba = recriarAbaProdutosSemTabela_(
      ssMae,
      aba,
      cabecalho,
      saida,
      categorias,
      nCols
    );
  }

  aplicarVisualAbaProdutosMae_(ssMae);
  return produtos.length;
}

/**
 * Mescla e formata as faixas de título das categorias.
 */
function aplicarFaixasCategoriaProdutos_(aba, categorias, nCols) {
  for (var k = 0; k < categorias.length; k++) {
    var linhaTitulo = CONFIG.LINHA_DADOS + categorias[k] - 1;
    var faixa = aba.getRange(linhaTitulo, 1, 1, nCols);
    faixa.merge();
    faixa
      .setBackground('#0F766E')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(11)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    aba.setRowHeight(linhaTitulo, 30);
  }
}

/**
 * Converte uma aba Produtos criada como "Tabela" em intervalo normal.
 * @return {Sheet}
 */
function recriarAbaProdutosSemTabela_(ssMae, abaAntiga, cabecalho, saida, categorias, nCols) {
  var indiceOriginal = abaAntiga.getIndex();
  var nomeAntigo = 'Produtos tabela antiga ' + new Date().getTime();
  abaAntiga.setName(nomeAntigo);

  var nova = ssMae.insertSheet(CONFIG.ABAS.PRODUTOS);
  ssMae.setActiveSheet(nova);
  ssMae.moveActiveSheet(indiceOriginal);

  nova.getRange(1, 1, 1, nCols).setValues([cabecalho]);
  if (saida.length > 0) {
    nova.getRange(CONFIG.LINHA_DADOS, 1, saida.length, nCols).setValues(saida);
  }

  aplicarValidacaoAtivoProdutos_(nova);
  aplicarFaixasCategoriaProdutos_(nova, categorias, nCols);

  try {
    ssMae.deleteSheet(abaAntiga);
  } catch (eDel) {
    registrarLog('PRODUTOS_TABELA_ANTIGA', 'Aba antiga mantida: ' + eDel.message);
  }
  return nova;
}

/**
 * Retorna todos os produtos da aba Produtos.
 * @return {Array.<Object>}
 */
function listarTodosProdutos() {
  var ss = obterPlanilhaMae();
  var aba = obterAbaObrigatoria(ss, CONFIG.ABAS.PRODUTOS);
  var numCols = Math.max(aba.getLastColumn(), CONFIG.PRODUTOS.HEADER.length, 8);
  var linhas = lerDados(aba, numCols);
  var map = obterMapaColunasProdutos_(aba);
  var produtos = [];

  for (var i = 0; i < linhas.length; i++) {
    var row = linhas[i];
    var codigo = String(row[map.colCodigo] || '').trim();
    var produto = String(row[map.colProduto] || '').trim();

    if (!codigo && !produto) {
      continue;
    }
    if (!produto) {
      continue;
    }

    produtos.push({
      linha: i + CONFIG.LINHA_DADOS,
      codigo: codigo,
      categoria: String(row[map.colCategoria] || '').trim(),
      produto: produto,
      unidadeMedida: String(row[map.colUnidade] || '').trim(),
      pregao: String(row[map.colPregao] || '').trim(),
      ativo: ehAtivo(row[map.colAtivo]),
      ativoRaw: row[map.colAtivo]
    });
  }

  return produtos;
}

/**
 * Retorna apenas produtos ativos, ordenados pelo nome.
 * @return {Array.<Object>}
 */
function listarProdutosAtivos() {
  var disponibilidadeCategorias = obterDisponibilidadeCategorias_();
  var ativos = listarTodosProdutos().filter(function (p) {
    return (
      p.ativo &&
      p.produto &&
      categoriaLiberadaNoMes_(p.categoria, disponibilidadeCategorias)
    );
  });

  ativos.sort(function (a, b) {
    return a.produto.localeCompare(b.produto, 'pt-BR');
  });

  return ativos;
}

/**
 * Retorna apenas os nomes dos produtos ativos (fonte do dropdown).
 * @return {string[]}
 */
function listarNomesProdutosAtivos() {
  return listarProdutosAtivos().map(function (p) {
    return p.produto;
  });
}

/**
 * Valida unicidade de nomes de produtos ativos.
 * @return {{ok: boolean, duplicados: string[]}}
 */
function validarUnicidadeProdutos() {
  var ativos = listarProdutosAtivos();
  var vistos = {};
  var duplicados = [];

  for (var i = 0; i < ativos.length; i++) {
    var chave = normalizarTexto(ativos[i].produto);
    if (vistos[chave]) {
      if (duplicados.indexOf(ativos[i].produto) === -1) {
        duplicados.push(ativos[i].produto);
      }
    } else {
      vistos[chave] = true;
    }
  }

  return { ok: duplicados.length === 0, duplicados: duplicados };
}

/**
 * Menu: insere uma linha vazia dentro da categoria escolhida.
 */
function menuInserirLinhaNovoProduto() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ssMae = obterPlanilhaMae();
    var ssTela = SpreadsheetApp.getActiveSpreadsheet();
    var usandoMaeNaTela = ssTela && ssTela.getId() === ssMae.getId();
    var aba = usandoMaeNaTela
      ? obterAbaObrigatoria(ssTela, CONFIG.ABAS.PRODUTOS)
      : obterAbaObrigatoria(ssMae, CONFIG.ABAS.PRODUTOS);
    var map = obterMapaColunasProdutos_(aba);
    var numCols = Math.max(aba.getLastColumn(), map.colAtivo + 1, CONFIG.PRODUTOS.HEADER.length);
    var selecao = usandoMaeNaTela ? ssTela.getActiveRange() : null;
    var linhaSelecionada =
      selecao && selecao.getSheet().getSheetId() === aba.getSheetId()
        ? selecao.getRow()
        : 0;
    var lastRow = Math.max(aba.getLastRow(), CONFIG.LINHA_DADOS);
    var info = obterContextoInsercaoProduto_(
      aba,
      map,
      linhaSelecionada || CONFIG.LINHA_DADOS,
      lastRow,
      numCols
    );
    var categoria = info.categoria;

    if (!categoria) {
      var resposta = ui.prompt(
        'Categoria do novo produto',
        'Digite exatamente a categoria onde o item será cadastrado (ex.: ELETRODOMÉSTICOS).',
        ui.ButtonSet.OK_CANCEL
      );
      if (resposta.getSelectedButton() !== ui.Button.OK) return;
      categoria = String(resposta.getResponseText() || '').trim();
      info.pregaoRef = '';
      info.numeroPregao = -1;
    }
    if (!categoria) {
      throw new Error('Informe ou selecione uma categoria antes de criar o produto.');
    }

    var sugestao = info.numeroPregao > 0 ? String(info.numeroPregao) : '';
    var respNum = ui.prompt(
      'Número do pregão',
      'Esta categoria pode ter vários pregões (ex.: 05 e 156).\n\n' +
        'Digite o NÚMERO do pregão que você vai cadastrar.\n' +
        '• 05 → a linha entra depois dos 05, antes dos 156\n' +
        '• 156 → a linha entra depois dos 156\n\n' +
        (sugestao ? 'Sugestão da linha clicada: ' + sugestao : 'Ex.: 05'),
      ui.ButtonSet.OK_CANCEL
    );
    if (respNum.getSelectedButton() !== ui.Button.OK) return;
    var digitado = String(respNum.getResponseText() || '').trim();
    var numeroAlvo = extrairNumeroPregaoCadastro_(digitado || sugestao);
    info.numeroPregao = numeroAlvo;
    info.pregaoRef = numeroAlvo > 0 ? String(numeroAlvo) + '/0000' : '';

    var novaLinha = encontrarLinhaParaNovoPregao_(
      aba,
      map,
      categoria,
      info.pregaoRef,
      lastRow,
      numCols
    );
    aba.insertRowsBefore(novaLinha, 1);

    var row = [];
    for (var c = 0; c < numCols; c++) row.push('');
    row[map.colCategoria] = categoria;
    if (map.colAtivo >= 0) row[map.colAtivo] = CONFIG.VALORES.ATIVO_SIM;
    aba.getRange(novaLinha, 1, 1, numCols).setValues([row]);

    aplicarValidacaoAtivoLinhaProduto_(aba, novaLinha, map);
    aplicarVisualLinhasProdutos_(aba, novaLinha, novaLinha, true);
    aba.setActiveRange(aba.getRange(novaLinha, map.colProduto + 1));

    ui.alert(
      'Nova linha',
      'Linha criada em: ' +
        categoria +
        (numeroAlvo > 0
          ? '\nPosição: bloco do pregão ' + numeroAlvo + ' (depois dos itens com esse número).'
          : '\nPosição: fim desta categoria.') +
        '\n\nPreencha Código, Produto, Unidade e Pregão.\n' +
        'Deixe Ativo = Sim.\n\nDepois use:\n' +
        'Sistema de Compras → Atualizar produtos nas unidades agora',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}

function obterContextoInsercaoProduto_(aba, map, linha, lastRow, nCols) {
  var out = { categoria: '', pregaoRef: '', numeroPregao: -1, linhaModelo: 0 };
  if (linha < 1) linha = CONFIG.LINHA_DADOS;
  if (linha > lastRow) linha = lastRow;

  var row = aba.getRange(linha, 1, 1, nCols).getValues()[0];
  if (ehLinhaTituloCategoriaProdutos_(row)) {
    out.categoria =
      obterCategoriaDaLinhaTituloProdutos_(row) ||
      String(row[0] || '')
        .replace(/^▸\s*/, '')
        .trim();
    return out;
  }

  out.categoria = String(row[map.colCategoria] || '').trim();
  out.pregaoRef = String(row[map.colPregao] || '').trim();
  out.numeroPregao = extrairNumeroPregaoCadastro_(out.pregaoRef);
  if (String(row[map.colProduto] || '').trim()) out.linhaModelo = linha;

  if (!out.categoria) {
    for (var r = linha; r >= CONFIG.LINHA_DADOS; r--) {
      var acima = aba.getRange(r, 1, 1, nCols).getValues()[0];
      if (ehLinhaTituloCategoriaProdutos_(acima)) {
        out.categoria =
          obterCategoriaDaLinhaTituloProdutos_(acima) ||
          String(acima[0] || '')
            .replace(/^▸\s*/, '')
            .trim();
        break;
      }
      var cat = String(acima[map.colCategoria] || '').trim();
      if (cat) {
        out.categoria = cat;
        if (!out.pregaoRef) {
          out.pregaoRef = String(acima[map.colPregao] || '').trim();
          out.numeroPregao = extrairNumeroPregaoCadastro_(out.pregaoRef);
        }
        break;
      }
    }
  }
  return out;
}

function encontrarLinhaParaNovoPregao_(aba, map, categoria, pregaoRef, lastRow, nCols) {
  var start = CONFIG.LINHA_DADOS;
  if (lastRow < start) return start;
  var vals = aba.getRange(start, 1, lastRow - start + 1, nCols).getValues();
  var chaveCat = normalizarTexto(categoria);
  var numeroRef = extrairNumeroPregaoCadastro_(pregaoRef);
  var ultimaCategoria = -1;
  var ultimoAteNumero = -1;
  var tituloCategoria = -1;

  for (var i = 0; i < vals.length; i++) {
    var row = vals[i];
    var linhaAbs = start + i;
    if (ehLinhaTituloCategoriaProdutos_(row)) {
      var nomeTitulo =
        obterCategoriaDaLinhaTituloProdutos_(row) ||
        String(row[0] || '')
          .replace(/^▸\s*/, '')
          .trim();
      if (normalizarTexto(nomeTitulo) === chaveCat) tituloCategoria = linhaAbs;
      continue;
    }
    var cat = String(row[map.colCategoria] || '').trim();
    if (normalizarTexto(cat) !== chaveCat) continue;
    ultimaCategoria = linhaAbs;
    var nLinha = extrairNumeroPregaoCadastro_(row[map.colPregao]);
    if (numeroRef > 0 && nLinha > 0 && nLinha <= numeroRef) {
      ultimoAteNumero = linhaAbs;
    }
  }

  if (ultimoAteNumero >= 0) return ultimoAteNumero + 1;
  if (numeroRef > 0 && tituloCategoria >= 0) return tituloCategoria + 1;
  if (ultimaCategoria >= 0) return ultimaCategoria + 1;
  if (tituloCategoria >= 0) return tituloCategoria + 1;
  return lastRow + 1;
}

function extrairNumeroPregaoCadastro_(texto) {
  var t = String(texto || '').trim();
  var m = t.match(/(\d{1,4})\s*\/\s*(\d{2,4})/);
  if (m) return parseInt(m[1], 10);
  var soNumero = t.match(/^(\d{1,4})$/);
  if (soNumero) return parseInt(soNumero[1], 10);
  return -1;
}

function aplicarVisualAbaProdutosMae_(ssMae, linhaInicio, linhaFim) {
  ssMae = ssMae || obterPlanilhaMae();
  var aba = ssMae.getSheetByName(CONFIG.ABAS.PRODUTOS);
  if (!aba || aba.getLastRow() < CONFIG.LINHA_DADOS) return;

  var map = obterMapaColunasProdutos_(aba);
  var nCols = Math.max(aba.getLastColumn(), map.colAtivo + 1, map.colProduto + 1, 6);

  var hdr = aba.getRange(1, 1, 1, nCols);
  hdr
    .setBackground(CONFIG.TEMA.HEADER_BG)
    .setFontColor(CONFIG.TEMA.HEADER_FG)
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontFamily('Arial')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  aba.setRowHeight(1, 40);

  var start = linhaInicio || CONFIG.LINHA_DADOS;
  var end = linhaFim || aba.getLastRow();
  if (end < start) return;

  aplicarVisualLinhasProdutos_(aba, start, end);
}

/**
 * Formata linhas de produtos (código escuro, categoria ciano, produto amarelo, pregão escuro).
 * @param {Sheet} aba
 * @param {number} linhaInicio 1-based
 * @param {number} linhaFim 1-based inclusive
 * @param {boolean=} forcarLinhaVazia aplica o padrão mesmo sem Código/Produto
 */
function aplicarVisualLinhasProdutos_(aba, linhaInicio, linhaFim, forcarLinhaVazia) {
  var map = obterMapaColunasProdutos_(aba);
  var t = CONFIG.TEMA;
  var nCols = Math.max(aba.getLastColumn(), map.colAtivo + 1, 6);
  var numLinhas = linhaFim - linhaInicio + 1;
  if (numLinhas < 1) return;

  try {
    aba.setTabColor(t.TAB);
  } catch (eTab) {}

  var valores = aba.getRange(linhaInicio, 1, numLinhas, nCols).getDisplayValues();
  var blocos = [];
  var bloco = null;

  for (var i = 0; i < numLinhas; i++) {
    var row = valores[i];
    var r = linhaInicio + i;
    if (ehLinhaTituloCategoriaProdutos_(row)) {
      if (bloco) {
        blocos.push(bloco);
        bloco = null;
      }
      continue;
    }

    var produto = String(row[map.colProduto] || '').trim();
    var codigo = String(row[map.colCodigo] || '').trim();
    if (!produto && !codigo && !forcarLinhaVazia) {
      if (bloco) {
        blocos.push(bloco);
        bloco = null;
      }
      continue;
    }

    var idx = r - CONFIG.LINHA_DADOS;
    var corEscura = idx % 2 === 0 ? t.ROW_A : t.ROW_B;
    var fundoLinha = [];
    var corLinha = [];
    var pesoLinha = [];
    var alignLinha = [];
    for (var c2 = 0; c2 < nCols; c2++) {
      if (c2 === map.colProduto) {
        fundoLinha.push(t.COL_DESTAQUE_BG);
        corLinha.push(t.COL_DESTAQUE_FG);
        pesoLinha.push('bold');
        alignLinha.push('center');
      } else if (c2 === map.colCategoria) {
        fundoLinha.push(t.COL_CATEGORIA_BG);
        corLinha.push(t.COL_CATEGORIA_FG);
        pesoLinha.push('normal');
        alignLinha.push('center');
      } else {
        fundoLinha.push(corEscura);
        corLinha.push(t.TEXT);
        pesoLinha.push('normal');
        alignLinha.push('center');
      }
    }

    if (!bloco) {
      bloco = { start: r, fundos: [], cores: [], pesos: [], aligns: [] };
    }
    bloco.fundos.push(fundoLinha);
    bloco.cores.push(corLinha);
    bloco.pesos.push(pesoLinha);
    bloco.aligns.push(alignLinha);
  }
  if (bloco) blocos.push(bloco);

  for (var b = 0; b < blocos.length; b++) {
    var bl = blocos[b];
    var n = bl.fundos.length;
    var rng = aba.getRange(bl.start, 1, n, nCols);
    rng
      .setBackgrounds(bl.fundos)
      .setFontColors(bl.cores)
      .setFontWeights(bl.pesos)
      .setHorizontalAlignments(bl.aligns)
      .setVerticalAlignment('middle')
      .setFontFamily('Arial')
      .setFontSize(10)
      .setBorder(true, true, true, true, true, true, t.BORDER, SpreadsheetApp.BorderStyle.SOLID);
    if (map.colProduto >= 0) {
      try {
        aba.getRange(bl.start, map.colProduto + 1, n, 1).setWrap(true);
      } catch (eWrap) {}
    }
  }
}

/**
 * Menu: reaplica visual em todas as linhas da aba Produtos.
 */
function menuAjustarVisualProdutos() {
  try {
    aplicarVisualAbaProdutosMae_();
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert(
      'Visual aplicado',
      'Todas as linhas da aba Produtos (incluindo Expediente/Papelaria e as demais categorias) foram formatadas no mesmo padrão:\n\n' +
        '• Produto amarelo\n' +
        '• Categoria ciano\n' +
        '• Código, Pregão e Ativo no fundo escuro',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function menuOrganizarProdutosPorCategoria() {
  var ui = SpreadsheetApp.getUi();
  try {
    var total = organizarProdutosPorCategoria_();
    SpreadsheetApp.flush();
    ui.alert(
      'Produtos organizados',
      total +
        ' produto(s) organizados por categoria.\n\n' +
        'Para cadastrar outro item, selecione o título ou um item da categoria e use “Nova linha”.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro ao organizar produtos', e.message, ui.ButtonSet.OK);
  }
}
