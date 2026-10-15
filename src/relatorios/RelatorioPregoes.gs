/**
 * Relatórios mensais de Pregões em Excel:
 * nome do pregão + produtos que cada um disponibiliza.
 */

/**
 * Menu: relatório Excel dos pregões ATIVOS e dos produtos que cada um disponibiliza.
 */
function menuGerarRelatorioPregoes() {
  gerarRelatorioPregoesPeloMenu_('ativos');
}

/**
 * Menu: relatório Excel dos pregões INATIVOS (Ativo = Não) e dos produtos de cada um.
 */
function menuGerarRelatorioPregoesInativos() {
  gerarRelatorioPregoesPeloMenu_('inativos');
}

/**
 * @param {string} modo ativos|inativos
 */
function gerarRelatorioPregoesPeloMenu_(modo) {
  var ui = SpreadsheetApp.getUi();
  var titulo =
    modo === 'inativos' ? 'Relatório de pregões inativos' : 'Relatório de pregões ativos';
  try {
    var ss = obterPlanilhaMae();
    ss.toast('Montando ' + titulo.toLowerCase() + ' em Excel⬦', 'Excel', 90);
    var arquivo = gerarRelatorioPregoesExcel_(ss, { modo: modo });
    ui.alert(
      titulo + ' gerado (Excel)',
      'Arquivo: ' +
        arquivo.getName() +
        '\n\nLayout:\n' +
        'GENERO:\n' +
        '(PREGAO):\n' +
        'PRODUTO QUE DISPONIBILIZA:\n\n' +
        'Abra ou baixe pelo link:\n' +
        arquivo.getUrl(),
      ui.ButtonSet.OK
    );
    registrarLog(
      modo === 'inativos' ? 'RELATORIO_PREGOES_INATIVOS' : 'RELATORIO_PREGOES',
      arquivo.getName() + ' | ' + arquivo.getUrl()
    );
  } catch (e) {
    ui.alert('Erro ao gerar relatório', e.message, ui.ButtonSet.OK);
    try {
      registrarLog(
        modo === 'inativos' ? 'RELATORIO_PREGOES_INATIVOS_ERRO' : 'RELATORIO_PREGOES_ERRO',
        e.message
      );
    } catch (eLog) {}
  }
}

/**
 * @param {Spreadsheet} ss
 * @param {{modo?:string}=} opcoes
 * @return {File}
 */
function gerarRelatorioPregoesWord_(ss, opcoes) {
  opcoes = opcoes || {};
  var modo = opcoes.modo === 'inativos' ? 'inativos' : 'ativos';
  var produtos = listarTodosProdutos();
  var dados = montarDadosRelatorioPregoes_(produtos, modo);
  if (!dados.pregoes.length) {
    throw new Error(
      modo === 'inativos'
        ? 'Não há pregões inativos na aba Produtos.\n\nMarque Ativo = Não nos produtos que devem aparecer neste relatório.'
        : 'Não há pregões ativos cadastrados na aba Produtos.\n\nPreencha a coluna Pregão e deixe Ativo = Sim.'
    );
  }

  var agora = new Date();
  var tz = Session.getScriptTimeZone();
  var dataArquivo = Utilities.formatDate(agora, tz, 'yyyy-MM-dd_HH-mm');
  var nomeDoc =
    (modo === 'inativos' ? 'Relatório Pregões Inativos - ' : 'Relatório Pregões - ') +
    dados.tituloMes +
    ' - ' +
    dataArquivo;

  var doc = DocumentApp.create(nomeDoc);
  var body = doc.getBody();
  body.clear();
  try {
    body.setPageWidth(841.89);
    body.setPageHeight(595.28);
    body.setMarginTop(36);
    body.setMarginBottom(36);
    body.setMarginLeft(28);
    body.setMarginRight(28);
  } catch (eMarg) {}

  montarCapaListaPregoes_(body, dados, agora, tz);
  for (var p = 0; p < dados.pregoes.length; p++) {
    body.appendPageBreak();
    montarSecaoPregaoMatriz_(body, dados.pregoes[p]);
  }

  montarRodapeRelatorioPregoes_(doc, agora, tz, modo);
  doc.saveAndClose();

  var arquivo = DriveApp.getFileById(doc.getId());
  try {
    arquivo.setDescription(String(dados.pregoes.length) + ' pregão(ões)');
  } catch (eDesc) {}
  var pasta = obterPastaRelatoriosPregoes_(ss);
  arquivo.moveTo(pasta);
  return arquivo;
}

/**
 * Gera .xlsx com as mesmas informações do relatório de pregões.
 * @param {Spreadsheet} ss
 * @param {{modo?:string}=} opcoes
 * @return {File}
 */
function gerarRelatorioPregoesExcel_(ss, opcoes) {
  opcoes = opcoes || {};
  var modo = opcoes.modo === 'inativos' ? 'inativos' : 'ativos';
  var produtos = listarTodosProdutos();
  var dados = montarDadosRelatorioPregoes_(produtos, modo);
  if (!dados.pregoes.length) {
    throw new Error(
      modo === 'inativos'
        ? 'Não há pregões inativos na aba Produtos.\n\nMarque Ativo = Não nos produtos que devem aparecer neste relatório.'
        : 'Não há pregões ativos cadastrados na aba Produtos.\n\nPreencha a coluna Pregão e deixe Ativo = Sim.'
    );
  }

  var agora = new Date();
  var tz = Session.getScriptTimeZone();
  var dataArquivo = Utilities.formatDate(agora, tz, 'yyyy-MM-dd_HH-mm');
  var rotulo = modo === 'inativos' ? 'Pregões Inativos' : 'Pregões Ativos';
  var nomeBase = 'Relatório ' + rotulo + ' - ' + dados.tituloMes + ' - ' + dataArquivo;

  var temp = SpreadsheetApp.create(nomeBase + ' (temp)');
  var aba = temp.getSheets()[0];
  aba.setName('Relatório');
  preencherAbaRelatorioPregoesLegivel_(aba, dados, agora, tz, modo);
  SpreadsheetApp.flush();

  var blob = exportarPlanilhaComoXlsx_(temp.getId(), nomeBase + '.xlsx');
  var pasta = obterPastaRelatoriosPregoes_(ss);
  var arquivo = pasta.createFile(blob);
  try {
    arquivo.setDescription(String(dados.pregoes.length) + ' pregão(ões)');
  } catch (eDesc) {}
  DriveApp.getFileById(temp.getId()).setTrashed(true);
  return arquivo;
}

/**
 * Agrupa: Gênero �  Pregão �  produtos.
 * @param {Object} dados
 * @return {Array.<{nome:string, pregoes:Array.<Object>}>}
 */
function agruparRelatorioPregoesPorGenero_(dados) {
  var mapa = {};
  var ordem = [];
  var pregoes = (dados && dados.pregoes) || [];
  for (var p = 0; p < pregoes.length; p++) {
    var bl = pregoes[p];
    for (var c = 0; c < (bl.categorias || []).length; c++) {
      var cat = bl.categorias[c];
      var nomeCat = String(cat.nome || 'SEM CATEGORIA').trim() || 'SEM CATEGORIA';
      var chave = normalizarTexto(nomeCat);
      if (!mapa[chave]) {
        mapa[chave] = { nome: nomeCat, pregoes: [] };
        ordem.push(chave);
      }
      mapa[chave].pregoes.push({
        nome: bl.nome,
        partes: bl.partes || {},
        produtos: cat.produtos || []
      });
    }
  }
  ordem.sort(function (a, b) {
    return mapa[a].nome.localeCompare(mapa[b].nome, 'pt-BR');
  });
  return ordem.map(function (chave) {
    var bloco = mapa[chave];
    bloco.pregoes.sort(function (a, b) {
      return compararPregoesOrdem_(a.nome, b.nome);
    });
    return bloco;
  });
}

/**
 * Uma aba clara: GENERO / (PREGAO) / PRODUTO QUE DISPONIBILIZA.
 * @param {Sheet} aba
 * @param {Object} dados
 * @param {Date} agora
 * @param {string} tz
 * @param {string} modo
 */
function preencherAbaRelatorioPregoesLegivel_(aba, dados, agora, tz, modo) {
  var nCols = 2;
  var inativos = modo === 'inativos';
  var generos = agruparRelatorioPregoesPorGenero_(dados);
  var linhas = [];
  var meta = [];

  linhas.push([
    'RELATORIO DE PREGOES ' + (inativos ? 'INATIVOS' : 'ATIVOS') + ' - ' + dados.tituloMes,
    ''
  ]);
  meta.push('titulo');
  linhas.push([
    'Gerado em ' +
      Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm') +
      '   |   ' + dados.resumo.totalPregoes + ' pregoes   |   ' + dados.resumo.totalProdutos + ' produtos',
    ''
  ]);
  meta.push('info');
  linhas.push(['', '']);
  meta.push('vazio');

  for (var g = 0; g < generos.length; g++) {
    var gen = generos[g];
    linhas.push(['GENERO:', String(gen.nome || '').toUpperCase()]);
    meta.push('genero');
    linhas.push(['', '']);
    meta.push('vazio');

    for (var p = 0; p < (gen.pregoes || []).length; p++) {
      var bl = gen.pregoes[p];
      linhas.push(['(PREGAO):', String(bl.nome || '').trim()]);
      meta.push('pregao');
      linhas.push(['PRODUTO QUE DISPONIBILIZA:', '']);
      meta.push('rotuloProduto');

      var itens = bl.produtos || [];
      if (!itens.length) {
        linhas.push(['', '(nenhum produto)']);
        meta.push('produto');
      }
      for (var i = 0; i < itens.length; i++) {
        var item = itens[i];
        linhas.push([
          '',
          '⬢  ' + (item.rotulo || String(item.produto || '').toUpperCase())
        ]);
        meta.push('produto');
      }

      linhas.push(['', '']);
      meta.push('vazio');
    }

    if (g < generos.length - 1) {
      linhas.push(['', '']);
      meta.push('separador');
    }
  }

  aba.getRange(1, 1, linhas.length, nCols).setValues(linhas);
  aba.setHiddenGridlines(true);
  aba.setFrozenRows(2);
  aba.setColumnWidth(1, 260);
  aba.setColumnWidth(2, 720);
  aba.setColumnWidth(3, 24);

  var maxR = Math.max(linhas.length + 8, 40);
  aba.getRange(1, 1, maxR, 4).setBackground('#F8FAFC');

  var refsGeneroA = [];
  var refsGeneroB = [];
  var refsPregaoA = [];
  var refsPregaoB = [];
  var refsRotulo = [];
  var refsProduto = [];
  var refsVazio = [];

  for (var r = 0; r < meta.length; r++) {
    var tipo = meta[r];
    var linha = r + 1;
    if (tipo === 'titulo') {
      try {
        aba.getRange(linha, 1, 1, nCols).merge();
      } catch (eT) {}
      aba
        .getRange(linha, 1, 1, nCols)
        .setBackground('#1D4ED8')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setFontSize(16)
        .setFontFamily('Arial')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      aba.setRowHeight(linha, 44);
    } else if (tipo === 'info') {
      try {
        aba.getRange(linha, 1, 1, nCols).merge();
      } catch (eI) {}
      aba
        .getRange(linha, 1, 1, nCols)
        .setBackground('#DBEAFE')
        .setFontColor('#1E3A8A')
        .setFontSize(11)
        .setFontFamily('Arial')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      aba.setRowHeight(linha, 28);
    } else if (tipo === 'genero') {
      refsGeneroA.push('A' + linha);
      refsGeneroB.push('B' + linha);
      aba.setRowHeight(linha, 38);
    } else if (tipo === 'pregao') {
      refsPregaoA.push('A' + linha);
      refsPregaoB.push('B' + linha);
      aba.setRowHeight(linha, 36);
    } else if (tipo === 'rotuloProduto') {
      refsRotulo.push('A' + linha + ':B' + linha);
      aba.setRowHeight(linha, 30);
    } else if (tipo === 'produto') {
      refsProduto.push('A' + linha + ':B' + linha);
      aba.setRowHeight(linha, 28);
    } else if (tipo === 'separador') {
      aba.setRowHeight(linha, 18);
      aba.getRange(linha, 1, 1, nCols).setBackground('#E2E8F0');
    } else {
      refsVazio.push('A' + linha + ':B' + linha);
      aba.setRowHeight(linha, 12);
    }
  }

  if (refsGeneroA.length) {
    aba
      .getRangeList(refsGeneroA)
      .setBackground('#0F766E')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(13)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    aba
      .getRangeList(refsGeneroB)
      .setBackground('#14B8A6')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(13)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);
  }
  if (refsPregaoA.length) {
    aba
      .getRangeList(refsPregaoA)
      .setBackground('#1D4ED8')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(12)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    aba
      .getRangeList(refsPregaoB)
      .setBackground('#EFF6FF')
      .setFontColor('#1E3A8A')
      .setFontWeight('bold')
      .setFontSize(11)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);
  }
  if (refsRotulo.length) {
    aba
      .getRangeList(refsRotulo)
      .setBackground('#FEF3C7')
      .setFontColor('#92400E')
      .setFontWeight('bold')
      .setFontSize(11)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
  }
  if (refsProduto.length) {
    aba
      .getRangeList(refsProduto)
      .setBackground('#FFFFFF')
      .setFontColor('#111827')
      .setFontWeight('normal')
      .setFontSize(12)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setWrap(true);
  }
  if (refsVazio.length) {
    aba.getRangeList(refsVazio).setBackground('#F8FAFC');
  }
}

/**
 * @deprecated layout antigo do Excel de pregões
 */
function preencherAbaIndicePregoesExcel_(aba, dados, agora, tz, modo) {
  preencherAbaRelatorioPregoesLegivel_(aba, dados, agora, tz, modo);
}

/** @deprecated */
function preencherAbaListaPregoesExcel_(aba, dados) {}

/** @deprecated */
function preencherAbaVisualPregoesExcel_(aba, dados, agora, tz, modo) {
  preencherAbaRelatorioPregoesLegivel_(aba, dados, agora, tz, modo);
}
/**
 * @param {Array.<Object>} produtos
 * @param {string=} modo ativos|inativos
 * @return {Object}
 */
function montarDadosRelatorioPregoes_(produtos, modo) {
  modo = modo === 'inativos' ? 'inativos' : 'ativos';
  var pregoesMap = {};
  var totalProdutos = 0;

  for (var i = 0; i < produtos.length; i++) {
    var p = produtos[i];
    var pregaoBruto = String(p.pregao || '').trim();
    if (!pregaoBruto) continue;
    if (modo === 'inativos' ? p.ativo : !p.ativo) continue;

    var pregao = normalizarTextoPregaoRelatorio_(pregaoBruto);
    var chavePregao = normalizarChavePregaoRelatorio_(pregao);

    if (!pregoesMap[chavePregao]) {
      var partes = parsearPartesPregao_(pregao);
      pregoesMap[chavePregao] = {
        nome: pregao,
        partes: partes,
        categoriasMap: {},
        ordemCategorias: [],
        totalProdutos: 0
      };
    }

    var cat = String(p.categoria || 'SEM CATEGORIA').trim() || 'SEM CATEGORIA';
    var chaveCat = normalizarTexto(cat);
    if (!pregoesMap[chavePregao].categoriasMap[chaveCat]) {
      pregoesMap[chavePregao].categoriasMap[chaveCat] = { nome: cat, produtos: [] };
      pregoesMap[chavePregao].ordemCategorias.push(chaveCat);
    }

    pregoesMap[chavePregao].categoriasMap[chaveCat].produtos.push({
      produto: p.produto,
      pregao: pregao,
      rotulo: limparNomeProdutoRelatorio_(p.produto, pregao).toUpperCase() ||
        String(p.produto || '').toUpperCase(),
      unidadeMedida: String(p.unidadeMedida || '').trim(),
      ativo: p.ativo
    });
    pregoesMap[chavePregao].totalProdutos++;
    totalProdutos++;
  }

  var nomesPregoes = Object.keys(pregoesMap).sort(function (a, b) {
    return compararPregoesOrdem_(pregoesMap[a].nome, pregoesMap[b].nome);
  });

  var pregoes = nomesPregoes.map(function (chave) {
    var bloco = pregoesMap[chave];
    bloco.ordemCategorias.sort(function (a, b) {
      return bloco.categoriasMap[a].nome.localeCompare(bloco.categoriasMap[b].nome, 'pt-BR');
    });
    bloco.categorias = bloco.ordemCategorias.map(function (chaveCat) {
      var cat = bloco.categoriasMap[chaveCat];
      cat.produtos.sort(function (x, y) {
        return x.produto.localeCompare(y.produto, 'pt-BR');
      });
      return cat;
    });
    return bloco;
  });

  return {
    tituloMes: obterTituloMesRelatorioPregoes_(new Date()),
    modo: modo,
    pregoes: pregoes,
    resumo: {
      totalPregoes: pregoes.length,
      totalProdutos: totalProdutos
    }
  };
}

/**
 * Uniformiza traços e espaços do texto do pregão (mantém o conteúdo original).
 * @param {string} texto
 * @return {string}
 */
function normalizarTextoPregaoRelatorio_(texto) {
  return String(texto || '')
    .replace(/[\u2013\u2014\u2212]/g, '-')
    .replace(/\s*-\s*/g, ' - ')
    .replace(/\s{2,}/g, ' ')
    .replace(/\s*-\s*$/g, '')
    .trim();
}

/**
 * Número do pregão + ano + data DE (quando passou a valer).
 * Ex.: 140/2025 - DE 14/10/2025 - FBA SOLUÇÕES.
 * @param {string} texto
 * @return {{numero:number, ano:number, data:number, texto:string}}
 */
function extrairChaveOrdemPregao_(texto) {
  var t = String(texto || '');
  var numero = 999999;
  var ano = 9999;
  var data = 99999999;
  var m = t.match(/(\d{1,4})\s*\/\s*(\d{2,4})/);
  if (m) {
    numero = parseInt(m[1], 10);
    ano = parseInt(m[2], 10);
    if (ano < 100) ano += 2000;
  }
  var md = t.match(/\bDE\s+(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/i);
  if (md) {
    var dia = parseInt(md[1], 10);
    var mes = parseInt(md[2], 10);
    var anoDe = parseInt(md[3], 10);
    if (anoDe < 100) anoDe += 2000;
    data = anoDe * 10000 + mes * 100 + dia;
  }
  return {
    numero: numero,
    ano: ano,
    data: data,
    texto: t.toLowerCase()
  };
}

/**
 * Ordem: ano do pregão, número (14, 90, 140…), data em que passou a valer, texto.
 * @param {string} textoA
 * @param {string} textoB
 * @return {number}
 */
function compararPregoesOrdem_(textoA, textoB) {
  var a = extrairChaveOrdemPregao_(textoA);
  var b = extrairChaveOrdemPregao_(textoB);
  if (a.ano !== b.ano) return a.ano - b.ano;
  if (a.numero !== b.numero) return a.numero - b.numero;
  if (a.data !== b.data) return a.data - b.data;
  return a.texto.localeCompare(b.texto, 'pt-BR');
}

/**
 * Chave de agrupamento (ignora variações mínimas de espaço/traço).
 * @param {string} texto
 * @return {string}
 */
function normalizarChavePregaoRelatorio_(texto) {
  return normalizarTexto(normalizarTextoPregaoRelatorio_(texto)).replace(/\s+/g, ' ');
}

/**
 * Extrai número, data DE, fornecedor e vencimento do texto do pregão.
 * Ex.: PREGÒO 218/2023 - DE 04/12/23 - PROCOMP ... LTDA - 11/09/2026
 * @param {string} texto
 * @return {{completo:string, numero:string, data:string, fornecedor:string, vencimento:string}}
 */
function parsearPartesPregao_(texto) {
  var completo = normalizarTextoPregaoRelatorio_(texto);
  var partes = {
    completo: completo,
    numero: '',
    data: '',
    fornecedor: '',
    vencimento: ''
  };

  var mNum = completo.match(/PREG[AÒÁì�]O\s*([0-9]+\/[0-9]{2,4})/i);
  if (mNum) partes.numero = 'PREGÒO ' + mNum[1];

  var mData = completo.match(/\bDE\s+(\d{1,2}\/\d{1,2}\/\d{2,4})\b/i);
  if (mData) partes.data = mData[1];

  var mVenc = completo.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})\s*$/);
  if (mVenc) partes.vencimento = mVenc[1];

  var pedacos = completo.split(/\s+-\s+/);
  if (pedacos.length >= 3) {
    // [0]=PREGÒO..., [1]=DE data, [2]=fornecedor, [último]=vencimento (se for data)
    var inicioForn = 2;
    var fimForn = pedacos.length - 1;
    if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(String(pedacos[fimForn] || '').trim())) {
      partes.vencimento = String(pedacos[fimForn]).trim();
      fimForn--;
    }
    partes.fornecedor = pedacos.slice(inicioForn, fimForn + 1).join(' - ').trim();
  } else if (!partes.fornecedor) {
    // Fallback: tira prefixo PREGÒO e DE data, fica o restante
    var resto = completo
      .replace(/^PREG[AÒÁì�]O\s*[0-9]+\/[0-9]{2,4}\s*/i, '')
      .replace(/^-\s*/i, '')
      .replace(/^DE\s+\d{1,2}\/\d{1,2}\/\d{2,4}\s*-?\s*/i, '')
      .replace(/\s*-\s*\d{1,2}\/\d{1,2}\/\d{2,4}\s*$/, '')
      .trim();
    partes.fornecedor = resto;
  }

  return partes;
}

/**
 * @param {Object} produto
 * @return {string}
 */
function montarRotuloItemPregao_(produto) {
  var nome = limparNomeProdutoRelatorio_(produto.produto, produto.pregao);
  var medida = String(produto.unidadeMedida || '').trim();
  var medidaEhPregao =
    /preg/i.test(medida) ||
    /\d{1,2}\/\d{1,2}\/\d{2,4}/.test(medida) ||
    medida.length > 40;
  if (medida && !medidaEhPregao && nome.toLowerCase().indexOf(medida.toLowerCase()) === -1) {
    nome = limparNomeProdutoRelatorio_(nome + ' ' + medida, produto.pregao);
  }
  return formatarNomeProdutoRelatorio_(nome);
}

/**
 * Remove o texto do pregão grudado no nome do produto.
 * @param {*} texto
 * @param {string=} pregaoTexto
 * @return {string}
 */
function limparNomeProdutoRelatorio_(texto, pregaoTexto) {
  var t = String(texto || '').replace(/\s+/g, ' ').trim();
  if (!t) return '';
  try {
    t = t.normalize('NFC');
  } catch (eNorm) {}

  if (pregaoTexto) {
    var preg = String(pregaoTexto).replace(/\s+/g, ' ').trim();
    if (preg) {
      var idx = t.toLowerCase().indexOf(preg.toLowerCase());
      if (idx > 0) t = t.substring(0, idx).trim();
    }
  }

  t = t.replace(/\s*[-��]?\s*PREG\S{0,4}O\s+\d+\s*\/\s*\d{2,4}[\s\S]*$/i, '');
  t = t.replace(/\s*[-��]?\s*PREG[AÒÁì�A̒]O\b[\s\S]*$/i, '');
  t = t.replace(/\s*[-��]\s*DE\s+\d{1,2}\/\d{1,2}\/\d{2,4}[\s\S]*$/i, '');
  t = t.replace(/\s*[-��]\s*\d{1,2}\/\d{1,2}\/\d{2,4}\s*$/i, '');
  t = t.replace(/[\s\-��]+$/g, '').trim();
  t = t.replace(/([A-Za-zì-ú])(\d)/g, '$1 $2');
  t = t.replace(/(\d)([A-Za-zì-ú])/g, '$1 $2');
  return t.replace(/\s+/g, ' ').trim();
}

/**
 * Nome do produto em caixa amigável para leitura (não grita em caixa alta).
 * @param {string} texto
 * @return {string}
 */
function formatarNomeProdutoRelatorio_(texto) {
  var t = String(texto || '').replace(/\s+/g, ' ').trim();
  if (!t) return '';
  var pequenos = {
    de: true,
    da: true,
    do: true,
    das: true,
    dos: true,
    e: true,
    com: true,
    para: true,
    p: true
  };
  var partes = t.split(' ');
  var out = [];
  for (var i = 0; i < partes.length; i++) {
    var w = partes[i];
    if (!w) continue;
    if (w.indexOf('/') !== -1 || /^\d/.test(w)) {
      out.push(w.toUpperCase());
      continue;
    }
    var low = w.toLowerCase();
    if (i > 0 && pequenos[low]) {
      out.push(low);
      continue;
    }
    out.push(low.charAt(0).toUpperCase() + low.slice(1));
  }
  return out.join(' ');
}

/**
 * @param {Array.<Array.<*>>} linhasCons
 * @return {Object.<string, Object.<string, string>>}
 */
function montarMapaQuantidadesPorProdutoUnidade_(linhasCons) {
  var mapa = {};
  for (var i = 0; i < linhasCons.length; i++) {
    var row = linhasCons[i];
    var unidade = String(row[0] || '').trim();
    var produto = padronizarNomenclaturaProdutoResumo_(row[1]);
    var qtd = row[3];
    if (!unidade || !produto) continue;
    var qtdTexto =
      qtd === null || qtd === undefined || qtd === '' ? '' : String(qtd).trim();
    if (!qtdTexto) continue;

    var chaveProd = normalizarTexto(produto);
    var chaveFlex =
      typeof chaveProdutoFlex_ === 'function' ? chaveProdutoFlex_(produto) : chaveProd;
    if (!mapa[chaveProd]) mapa[chaveProd] = {};
    mapa[chaveProd][normalizarTexto(unidade)] = qtdTexto;
    // Alias flexível (acentos / pontuação) para casar com o catálogo
    if (chaveFlex && chaveFlex !== chaveProd) {
      if (!mapa[chaveFlex]) mapa[chaveFlex] = mapa[chaveProd];
      else {
        var uns = Object.keys(mapa[chaveProd]);
        for (var u = 0; u < uns.length; u++) {
          mapa[chaveFlex][uns[u]] = mapa[chaveProd][uns[u]];
        }
      }
    }
  }
  return mapa;
}

/**
 * @param {Date=} data
 * @return {string}
 */
function obterTituloMesRelatorioPregoes_(data) {
  data = data || new Date();
  var meses = [
    'JANEIRO',
    'FEVEREIRO',
    'MAR�!O',
    'ABRIL',
    'MAIO',
    'JUNHO',
    'JULHO',
    'AGOSTO',
    'SETEMBRO',
    'OUTUBRO',
    'NOVEMBRO',
    'DEZEMBRO'
  ];
  return meses[data.getMonth()] + ' ' + data.getFullYear();
}

function obterPaletaRelatorioPregoes_() {
  return {
    titulo: '#1B365D',
    texto: '#1F2937',
    suave: '#5C6770',
    linha: '#DEE2E6',
    fundo: '#F4F6F8',
    fundoMeta: '#C5D4E8',
    branco: '#FFFFFF'
  };
}

/**
 * Capa simples (mês + ativos/inativos). As grades começam na página seguinte.
 * @param {Body} body
 * @param {Object} dados
 * @param {Date} agora
 * @param {string} tz
 */
function montarCapaListaPregoes_(body, dados, agora, tz) {
  var cor = obterPaletaRelatorioPregoes_();
  var inativos = dados.modo === 'inativos';

  adicionarLogoPrefeituraNoRelatorio_(body);

  aplicarEstiloParagrafoRelatorio_(
    body.appendParagraph(obterNomeOrganizacao_()).setAlignment(DocumentApp.HorizontalAlignment.CENTER),
    { cor: cor.titulo, negrito: true, tamanho: 11 }
  );
  aplicarEstiloParagrafoRelatorio_(
    body
      .appendParagraph('Secretaria Municipal de Assistência e Desenvolvimento Social')
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER),
    { cor: cor.suave, tamanho: 9 }
  );

  body.appendParagraph('');

  aplicarEstiloParagrafoRelatorio_(
    body.appendParagraph(dados.tituloMes).setAlignment(DocumentApp.HorizontalAlignment.CENTER),
    { cor: cor.titulo, negrito: true, tamanho: 18, depois: 2 }
  );
  aplicarEstiloParagrafoRelatorio_(
    body
      .appendParagraph(inativos ? 'Pregões inativos' : 'Pregões ativos')
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER),
    { cor: cor.texto, tamanho: 11, depois: 2 }
  );
  aplicarEstiloParagrafoRelatorio_(
    body
      .appendParagraph(
        Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm') +
          '   ·   ' +
          dados.resumo.totalPregoes +
          ' pregões   ·   ' +
          dados.resumo.totalProdutos +
          ' produtos'
      )
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER),
    { cor: cor.suave, tamanho: 9, depois: 12 }
  );
}

/**
 * Página do pregão: título + categoria vermelha + itens disponibilizados.
 * @param {Body} body
 * @param {Object} blocoPregao
 */
function montarSecaoPregaoMatriz_(body, blocoPregao) {
  aplicarEstiloParagrafoRelatorio_(
    body.appendParagraph(String(blocoPregao.nome || '').toUpperCase()).setAlignment(
      DocumentApp.HorizontalAlignment.CENTER
    ),
    { cor: '#000000', negrito: true, tamanho: 11, depois: 4, sublinhado: true }
  );

  for (var c = 0; c < (blocoPregao.categorias || []).length; c++) {
    var categoria = blocoPregao.categorias[c];
    if (!categoria.produtos.length) continue;

    aplicarEstiloParagrafoRelatorio_(
      body
        .appendParagraph(String(categoria.nome || 'SEM CATEGORIA').toUpperCase())
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER),
      { cor: '#C41E3A', negrito: true, tamanho: 11, antes: 6, depois: 8, sublinhado: true }
    );

    montarTabelaItensPregao_(body, categoria.produtos);
  }
}

/**
 * Grade só com os itens do pregão (sem unidades).
 * @param {Body} body
 * @param {Array.<Object>} produtos
 */
function montarTabelaItensPregao_(body, produtos) {
  var linhas = [['ITEM']];
  for (var p = 0; p < produtos.length; p++) {
    var item = produtos[p];
    linhas.push([item.rotulo || String(item.produto || '').toUpperCase()]);
  }

  var tabela = body.appendTable(linhas);
  formatarTabelaItensPregao_(tabela);
}

/**
 * @param {Table} tabela
 */
function formatarTabelaItensPregao_(tabela) {
  try {
    tabela.setBorderColor('#000000');
    tabela.setBorderWidth(1);
    tabela.setColumnWidth(0, 640);
  } catch (eTab) {}

  estilizarCelulaFichaPregao_(tabela.getCell(0, 0), {
    fundo: '#FFFFFF',
    cor: '#000000',
    negrito: true,
    tamanho: 10,
    pad: 4
  });
  try {
    tabela.getCell(0, 0).getChild(0).asParagraph().setAlignment(
      DocumentApp.HorizontalAlignment.CENTER
    );
  } catch (eAl) {}

  for (var r = 1; r < tabela.getNumRows(); r++) {
    estilizarCelulaFichaPregao_(tabela.getCell(r, 0), {
      fundo: '#FFFFFF',
      cor: '#000000',
      negrito: true,
      tamanho: 10,
      pad: 3
    });
    try {
      tabela.getCell(r, 0).getChild(0).asParagraph().setAlignment(
        DocumentApp.HorizontalAlignment.LEFT
      );
    } catch (eAl2) {}
  }
}

/**
 * @param {TableCell} celula
 * @param {Object} op
 */
function estilizarCelulaFichaPregao_(celula, op) {
  op = op || {};
  try {
    celula.setBackgroundColor(op.fundo || '#FFFFFF');
    celula.setPaddingTop(op.pad || 4);
    celula.setPaddingBottom(op.pad || 4);
    celula.setPaddingLeft(10);
    celula.setPaddingRight(10);
  } catch (eCel) {}
  try {
    var filho = celula.getChild(0);
    if (filho && filho.getType() === DocumentApp.ElementType.PARAGRAPH) {
      filho.asParagraph().setAttributes({
        [DocumentApp.Attribute.FOREGROUND_COLOR]: op.cor || '#1F2937',
        [DocumentApp.Attribute.BOLD]: op.negrito === true,
        [DocumentApp.Attribute.FONT_SIZE]: op.tamanho || 10,
        [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
      });
      if (op.manter) {
        try {
          filho.asParagraph().setKeepWithNext(true);
        } catch (eKeep) {}
      }
    }
  } catch (eTxt) {}
}

/**
 * @param {Paragraph} paragrafo
 * @param {{cor?:string, negrito?:boolean, tamanho?:number, antes?:number, depois?:number, sublinhado?:boolean, manter?:boolean}} op
 */
function aplicarEstiloParagrafoRelatorio_(paragrafo, op) {
  op = op || {};
  paragrafo.setAttributes({
    [DocumentApp.Attribute.FOREGROUND_COLOR]: op.cor || '#1F2937',
    [DocumentApp.Attribute.BOLD]: op.negrito === true,
    [DocumentApp.Attribute.FONT_SIZE]: op.tamanho || 10,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
    [DocumentApp.Attribute.UNDERLINE]: op.sublinhado === true
  });
  if (op.antes != null) paragrafo.setSpacingBefore(op.antes);
  if (op.depois != null) paragrafo.setSpacingAfter(op.depois);
  if (op.manter) {
    try {
      paragrafo.setKeepWithNext(true);
    } catch (eKeep) {}
  }
  return paragrafo;
}

/**
 * @param {Document} doc
 * @param {Date} agora
 * @param {string} tz
 */
function montarRodapeRelatorioPregoes_(doc, agora, tz, modo) {
  var footer = doc.addFooter();
  footer
    .appendParagraph(
      'UNIDADES  |  Pregões ' +
        (modo === 'inativos' ? 'inativos' : 'ativos') +
        '  |  ' +
        Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm')
    )
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setAttributes({
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#000000',
      [DocumentApp.Attribute.FONT_SIZE]: 8,
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
    });
}


/**
 * @param {Spreadsheet} ss
 * @return {Folder}
 */
function obterPastaRelatoriosPregoes_(ss) {
  var nome = 'Relatórios Pregões - Sistema Prefeitura Compras';
  var arquivoMae = DriveApp.getFileById(ss.getId());
  var pais = arquivoMae.getParents();
  var pai = pais.hasNext() ? pais.next() : null;
  var pastas = pai ? pai.getFoldersByName(nome) : DriveApp.getFoldersByName(nome);
  if (pastas.hasNext()) return pastas.next();
  return pai ? pai.createFolder(nome) : DriveApp.createFolder(nome);
}
