/**
 * Relatório administrativo personalizado, gerado a partir dos Consolidados.
 */

function menuGerarRelatorioPersonalizado() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = obterPlanilhaMae();
    var aba = ss.getSheetByName(CONFIG.ABAS.CONSOLIDADOS);
    var linhas = aba ? lerDados(aba, CONFIG.CONSOLIDADOS.HEADER.length) : [];
    linhas = linhas.filter(function (row) {
      return String(row[0] || '').trim() && String(row[1] || '').trim();
    });
    if (!linhas.length) {
      throw new Error('Não há solicitações enviadas para gerar o relatório.');
    }

    ss.toast('Montando relatório personalizado…', 'PDF', 60);
    var arquivo = gerarRelatorioPersonalizado_(ss, linhas);
    ui.alert(
      'Relatório personalizado gerado',
      'Arquivo: ' + arquivo.getName() + '\n\nAbra pelo link:\n' + arquivo.getUrl(),
      ui.ButtonSet.OK
    );
    registrarLog('PDF_RELATORIO_PERSONALIZADO', arquivo.getName() + ' | ' + arquivo.getUrl());
  } catch (e) {
    ui.alert('Erro ao gerar relatório', e.message, ui.ButtonSet.OK);
    try {
      registrarLog('PDF_RELATORIO_ERRO', e.message);
    } catch (eLog) {}
  }
}

/**
 * @param {Spreadsheet} ss
 * @param {Array.<Array.<*>>} linhasCons [Unidade, Produto, Qtd, Obs, Data]
 * @return {File}
 */
function gerarRelatorioPersonalizado_(ss, linhasCons) {
  var dados = agruparDadosRelatorioPorUnidade_(linhasCons);
  var agora = new Date();
  var dataArquivo = Utilities.formatDate(
    agora,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd_HH-mm'
  );
  var nome = 'Relatório de Solicitação de Produtos - ' + dataArquivo;
  var doc = DocumentApp.create(nome);
  var body = doc.getBody();
  body.clear();

  var azul = '#17365D';
  var azulClaro = '#3F7FBA';
  var cinza = '#667085';

  adicionarLogoPrefeituraNoRelatorio_(body);
  var cabecalho = body.appendParagraph(obterNomeOrganizacao_());
  cabecalho
    .setAlignment(DocumentApp.HorizontalAlignment.LEFT)
    .setAttributes({
      [DocumentApp.Attribute.FOREGROUND_COLOR]: azul,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FONT_SIZE]: 15
    });
  var secretaria = body.appendParagraph(
    'Secretaria Municipal de Assistência e Desenvolvimento Social'
  );
  secretaria.setAttributes({
    [DocumentApp.Attribute.FOREGROUND_COLOR]: cinza,
    [DocumentApp.Attribute.FONT_SIZE]: 8
  });
  body.appendParagraph('________________________________________________________________________')
    .setAttributes({ [DocumentApp.Attribute.FOREGROUND_COLOR]: azul });

  body.appendParagraph('RELATÓRIO DE SOLICITAÇÃO DE PRODUTOS')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAttributes({
      [DocumentApp.Attribute.FOREGROUND_COLOR]: azul,
      [DocumentApp.Attribute.FONT_SIZE]: 21,
      [DocumentApp.Attribute.BOLD]: true
    });
  body.appendParagraph('Resumo de Solicitações por Unidade')
    .setAttributes({
      [DocumentApp.Attribute.FOREGROUND_COLOR]: azulClaro,
      [DocumentApp.Attribute.FONT_SIZE]: 11
    });
  body.appendParagraph(
    'Documento consolidado contendo os produtos solicitados pelas unidades, ' +
      'com suas respectivas datas e quantidades.'
  ).setAttributes({
    [DocumentApp.Attribute.FOREGROUND_COLOR]: cinza,
    [DocumentApp.Attribute.FONT_SIZE]: 8
  });

  body.appendParagraph('VISÃO GERAL')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAttributes({
      [DocumentApp.Attribute.FOREGROUND_COLOR]: azul,
      [DocumentApp.Attribute.FONT_SIZE]: 11,
      [DocumentApp.Attribute.BOLD]: true
    });
  var visao = body.appendTable([
    ['UNIDADES', 'REGISTROS', 'QUANTIDADE'],
    [
      String(dados.unidades.length),
      String(linhasCons.length),
      String(dados.quantidadeTotal)
    ]
  ]);
  formatarTabelaVisaoGeral_(visao, azul);

  body.appendParagraph('SOLICITAÇÕES POR UNIDADE')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAttributes({
      [DocumentApp.Attribute.FOREGROUND_COLOR]: azul,
      [DocumentApp.Attribute.FONT_SIZE]: 11,
      [DocumentApp.Attribute.BOLD]: true
    });

  for (var u = 0; u < dados.unidades.length; u++) {
    var unidade = dados.unidades[u];
    body.appendParagraph(unidade.nome.toUpperCase() + '  —  ' + unidade.itens.length + ' item(ns)')
      .setAttributes({
        [DocumentApp.Attribute.BACKGROUND_COLOR]: azul,
        [DocumentApp.Attribute.FOREGROUND_COLOR]: '#FFFFFF',
        [DocumentApp.Attribute.BOLD]: true,
        [DocumentApp.Attribute.FONT_SIZE]: 10
      });

    var tabelaDados = [['PRODUTO', 'DATA', 'ESTOQUE ATUAL', 'QUANTIDADE']];
    for (var i = 0; i < unidade.itens.length; i++) {
      var item = unidade.itens[i];
      tabelaDados.push([
        item.produto,
        formatarDataResumo_(item.data),
        String(valorCelulaPreservandoZero_(item.estoqueAtual)),
        String(item.quantidade)
      ]);
    }
    var tabela = body.appendTable(tabelaDados);
    formatarTabelaSolicitacoes_(tabela, azulClaro);
    body.appendParagraph('');
  }

  var footer = doc.addFooter();
  footer.appendParagraph(
    'Secretaria Municipal de Assistência e Desenvolvimento Social  |  Gerado em ' +
      Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
  ).setAttributes({
    [DocumentApp.Attribute.FOREGROUND_COLOR]: cinza,
    [DocumentApp.Attribute.FONT_SIZE]: 7
  });

  doc.saveAndClose();
  var blob = DriveApp.getFileById(doc.getId())
    .getAs(MimeType.PDF)
    .setName(nome + '.pdf');
  var pasta = obterPastaPdfsResumo_(ss);
  var arquivo = pasta.createFile(blob);
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return arquivo;
}

/**
 * Insere no documento o logotipo disponibilizado no portal oficial.
 * A falha de rede não impede a emissão do relatório.
 */
function adicionarLogoPrefeituraNoRelatorio_(body) {
  var urlLogo = (CONFIG.ORGANIZACAO && CONFIG.ORGANIZACAO.LOGO_URL) || '';
  try {
    var resposta = UrlFetchApp.fetch(urlLogo, { muteHttpExceptions: true });
    if (resposta.getResponseCode() !== 200) return;

    var linhaLogo = body.appendParagraph('');
    linhaLogo.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    var logo = linhaLogo.appendInlineImage(resposta.getBlob());
    logo.setWidth(135);
    logo.setHeight(55);
  } catch (e) {
    Logger.log('Não foi possível inserir o logo da Prefeitura: ' + e.message);
  }
}

function agruparDadosRelatorioPorUnidade_(linhasCons, ssMae) {
  var catalogo = {};
  try {
    catalogo = montarMapaCatalogoProdutoPregao_(ssMae || obterPlanilhaMae());
  } catch (eCat) {}

  var mapa = {};
  var ordem = [];
  var quantidadeTotal = 0;
  for (var i = 0; i < linhasCons.length; i++) {
    var row = linhasCons[i];
    var nome = String(row[0] || '').trim();
    if (!nome) continue;
    var chave = normalizarTexto(nome);
    if (!mapa[chave]) {
      mapa[chave] = { nome: nome, itens: [] };
      ordem.push(chave);
    }
    var produto = padronizarNomenclaturaProdutoResumo_(row[1]);
    var chaveProd = normalizarTexto(produto);
    var chaveFlex =
      typeof chaveProdutoFlex_ === 'function' ? chaveProdutoFlex_(produto) : chaveProd;
    var info = catalogo[chaveProd] || catalogo[chaveFlex] || {};
    var pregaoTexto = String(info.pregao || '').trim() || 'SEM PREGÃO CADASTRADO';
    var pregaoExibir = pregaoTexto;
    try {
      pregaoExibir = rotuloPregaoResumo_(pregaoTexto).replace(/^▸\s*/, '');
    } catch (eR) {}

    mapa[chave].itens.push({
      produto: produto,
      pregao: pregaoExibir,
      estoqueAtual: valorCelulaPreservandoZero_(row[2]),
      quantidade: row[3],
      data: row[5] || row[4] || ''
    });
    var numero = extrairNumeroQuantidade(row[3]);
    if (numero !== null) quantidadeTotal += numero;
  }
  ordem.sort(function (a, b) {
    return mapa[a].nome.localeCompare(mapa[b].nome, 'pt-BR');
  });
  return {
    unidades: ordem.map(function (chave) {
      var bloco = mapa[chave];
      bloco.itens.sort(function (a, b) {
        var cmpPregao =
          typeof compararPregoesOrdem_ === 'function'
            ? compararPregoesOrdem_(a.pregao, b.pregao)
            : String(a.pregao || '').localeCompare(String(b.pregao || ''), 'pt-BR');
        if (cmpPregao !== 0) return cmpPregao;
        return a.produto.localeCompare(b.produto, 'pt-BR');
      });
      return bloco;
    }),
    quantidadeTotal: quantidadeTotal
  };
}

function formatarTabelaVisaoGeral_(tabela, azul) {
  for (var c = 0; c < 3; c++) {
    tabela.getCell(0, c).setAttributes({
      [DocumentApp.Attribute.BACKGROUND_COLOR]: '#E9EEF5',
      [DocumentApp.Attribute.FOREGROUND_COLOR]: azul,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FONT_SIZE]: 8,
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
    });
    tabela.getCell(1, c).setAttributes({
      [DocumentApp.Attribute.FOREGROUND_COLOR]: azul,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FONT_SIZE]: 18,
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
    });
  }
}

function formatarTabelaSolicitacoes_(tabela, azulClaro) {
  var colunas = tabela.getRow(0).getNumCells();
  for (var c = 0; c < colunas; c++) {
    tabela.getCell(0, c).setAttributes({
      [DocumentApp.Attribute.BACKGROUND_COLOR]: azulClaro,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#FFFFFF',
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FONT_SIZE]: 8,
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
    });
  }
  for (var r = 1; r < tabela.getNumRows(); r++) {
    var bg = r % 2 === 0 ? '#F2F4F7' : '#FFFFFF';
    // Coluna Produto: mesmo padrão do Resumo (Arial 10, negrito, maiúsculas).
    tabela.getCell(r, 0).setAttributes({
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
      [DocumentApp.Attribute.FONT_SIZE]: 10,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#000000',
      [DocumentApp.Attribute.BACKGROUND_COLOR]: bg,
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT
    });
    for (var col = 1; col < colunas; col++) {
      tabela.getCell(r, col).setAttributes({
        [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
        [DocumentApp.Attribute.FONT_SIZE]: 8,
        [DocumentApp.Attribute.BACKGROUND_COLOR]: bg,
        [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER
      });
    }
  }
}
