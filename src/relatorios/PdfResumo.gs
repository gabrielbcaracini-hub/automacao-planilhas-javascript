/**
 * Geração de PDF da aba Resumo para controle administrativo da Planilha Mãe.
 */

/**
 * Menu: exporta todo o conteúdo atual da aba Resumo como PDF.
 */
function menuGerarPdfResumo() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = obterPlanilhaMae();
    var aba = ss.getSheetByName(CONFIG.ABAS.RESUMO);
    if (!aba || aba.getLastRow() < 2) {
      throw new Error('A aba Resumo não possui pedidos para exportar.');
    }

    ss.toast('Gerando PDF do Resumo…', 'PDF', 30);
    SpreadsheetApp.flush();

    var blob = gerarBlobPdfResumo_(ss, aba);
    var pasta = obterPastaPdfsResumo_(ss);
    var data = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd_HH-mm'
    );
    var arquivo = pasta
      .createFile(blob)
      .setName('Resumo de Pedidos - ' + data + '.pdf');

    ui.alert(
      'PDF gerado',
      'O PDF com todo o conteúdo do Resumo foi salvo no Drive.\n\n' +
        'Arquivo: ' +
        arquivo.getName() +
        '\nPasta: ' +
        pasta.getName() +
        '\n\nAbra pelo link:\n' +
        arquivo.getUrl(),
      ui.ButtonSet.OK
    );
    registrarLog('PDF_RESUMO', arquivo.getName() + ' | ' + arquivo.getUrl());
  } catch (e) {
    ui.alert('Erro ao gerar PDF', e.message, ui.ButtonSet.OK);
    try {
      registrarLog('PDF_RESUMO_ERRO', e.message);
    } catch (eLog) {}
  }
}

/**
 * Exporta somente a aba Resumo, mantendo sua formatação de blocos por unidade.
 * @return {Blob}
 */
function gerarBlobPdfResumo_(ss, abaResumo) {
  var base =
    'https://docs.google.com/spreadsheets/d/' +
    encodeURIComponent(ss.getId()) +
    '/export';
  var parametros = [
    'format=pdf',
    'gid=' + encodeURIComponent(abaResumo.getSheetId()),
    // Limita a exportação ao conteúdo real do Resumo; o fundo visual da
    // planilha pode se estender por muitas colunas e reduzir o PDF inteiro.
    'r1=0',
    'c1=0',
    'r2=' + encodeURIComponent(abaResumo.getLastRow()),
    'c2=' + encodeURIComponent(Math.max(abaResumo.getLastColumn(), 2)),
    'size=A4',
    'portrait=false',
    'fitw=true',
    'top_margin=0.35',
    'bottom_margin=0.35',
    'left_margin=0.35',
    'right_margin=0.35',
    'sheetnames=false',
    'printtitle=false',
    'pagenumbers=true',
    'gridlines=false',
    'fzr=true',
    'single=true',
    'attachment=true'
  ];

  var resposta = UrlFetchApp.fetch(base + '?' + parametros.join('&'), {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resposta.getResponseCode() !== 200) {
    throw new Error(
      'Não foi possível exportar o Resumo (HTTP ' +
        resposta.getResponseCode() +
        ').'
    );
  }
  return resposta.getBlob().setContentType('application/pdf');
}

/**
 * Obtém/cria a pasta onde os PDFs administrativos serão armazenados.
 * @return {Folder}
 */
function obterPastaPdfsResumo_(ss) {
  var nome = 'PDFs - Sistema Prefeitura Compras';
  var arquivoMae = DriveApp.getFileById(ss.getId());
  var pais = arquivoMae.getParents();
  var pai = pais.hasNext() ? pais.next() : null;
  var pastas = pai ? pai.getFoldersByName(nome) : DriveApp.getFoldersByName(nome);
  if (pastas.hasNext()) return pastas.next();
  return pai ? pai.createFolder(nome) : DriveApp.createFolder(nome);
}
