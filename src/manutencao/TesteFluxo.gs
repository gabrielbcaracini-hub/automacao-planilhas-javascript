/**
 * Teste automático do fluxo: grava pedido na unidade e confere na Mãe.
 */
function menuTesteFluxoCompleto() {
  SpreadsheetApp.getUi().alert(
    'Teste desligado',
    'Esse teste gravaria um pedido falso na planilha de uma unidade.\n\nEle não faz parte do uso do sistema.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * @return {string}
 */
function executarTesteFluxoCompleto_() {
  throw new Error(
    'Teste de fluxo desligado. Ele gravaria um pedido falso em uma unidade real.'
  );
  var linhas = [];
  var idUnidade = resolverIdUnidadePeloNome_('UNIDADE CENTRO');
  if (!idUnidade) {
    throw new Error('Não encontrei Pedido - UNIDADE CENTRO no Drive');
  }
  linhas.push('1) Unidade ID: ' + idUnidade);

  var ssFilha = SpreadsheetApp.openById(idUnidade);
  linhas.push('2) Arquivo: ' + ssFilha.getName());

  // Garante aba Pedido limpa (sem tabela tipada)
  limparAbasErradasDaFilha(ssFilha);
  var aba = ssFilha.getSheetByName('Pedido');
  if (!aba) {
    aba = ssFilha.insertSheet('Pedido');
  }

  // Lê estado atual
  var antes = lerPedidosEnviadosFlexivel(aba);
  linhas.push('3) Enviados antes: ' + antes.enviados.length);

  // Recria limpa preservando linhas
  var todas = antes.todas && antes.todas.length ? antes.todas : lerPedidosFallbackSimples_(aba);
  var produtoTeste = 'TESTE AUTO ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HHmmss');
  todas.push({
    data: new Date(),
    produto: produtoTeste,
    quantidade: '99',
    observacao: 'teste fluxo automatico',
    status: CONFIG.VALORES.STATUS_ENVIADO
  });

  recriarAbaPedidoLimpa_(ssFilha, aba, todas);
  SpreadsheetApp.flush();
  linhas.push('4) Pedido de teste gravado na unidade: ' + produtoTeste + ' = 99 / Enviado');

  // Sync para a mãe (mesmo caminho do gatilho)
  atualizarIdUnidadeNaAba('UNIDADE CENTRO', idUnidade);
  sincronizarUnidadePorId(idUnidade);
  SpreadsheetApp.flush();
  linhas.push('5) sincronizarUnidadePorId() executado');

  // Confere na Mãe
  var ssMae = SpreadsheetApp.openById(CONFIG.SPREADSHEET_MAE_ID);
  linhas.push('6) Mãe: ' + ssMae.getName() + ' (' + ssMae.getId() + ')');

  var abaResumo = ssMae.getSheetByName(CONFIG.ABAS.RESUMO);
  if (!abaResumo) {
    throw new Error('Aba Resumo não existe na Mãe');
  }
  var resumoVals = abaResumo.getDataRange().getValues();
  var achouResumo = false;
  var amostraResumo = [];
  for (var i = 1; i < resumoVals.length; i++) {
    var p = String(resumoVals[i][0] || '');
    var d = String(resumoVals[i][1] || '');
    var e = String(resumoVals[i][CONFIG.RESUMO.ESTOQUE_ATUAL - 1] || '');
    var q = String(resumoVals[i][CONFIG.RESUMO.QUANTIDADE - 1] || '');
    if (p.indexOf('UNIDADE:') !== -1 || !p.trim()) continue;
    if (amostraResumo.length < 5) amostraResumo.push(p + ' | ' + d + ' | estoque=' + e + ' | qtd=' + q);
    if (p.indexOf('TESTE AUTO') !== -1 || p === produtoTeste) {
      achouResumo = true;
    }
  }

  linhas.push('7) Resumo tem o teste? ' + (achouResumo ? 'SIM' : 'NÃO'));
  linhas.push('8) Amostra Resumo:\n' + (amostraResumo.join('\n') || '(vazio)'));
  linhas.push('9) Total linhas Resumo: ' + Math.max(0, resumoVals.length - 1));

  registrarLog(
    'TESTE_FLUXO',
    'resumo=' + achouResumo + ' produto=' + produtoTeste
  );

  if (!achouResumo) {
    linhas.push('');
    linhas.push('FALHA: o pedido de teste precisa aparecer no Resumo.');
  } else {
    linhas.push('');
    linhas.push('SUCESSO: Resumo atualizado.');
  }

  return linhas.join('\n');
}

/**
 * Fallback se lerPedidosEnviadosFlexivel não trouxer "todas".
 * @param {Sheet} aba
 * @return {Array.<Object>}
 */
function lerPedidosFallbackSimples_(aba) {
  var values = aba.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var produto = String(values[i][1] || '').trim();
    if (!produto) continue;
    out.push({
      data: values[i][0] || '',
      produto: produto,
      estoqueAtual: valorCelulaPreservandoZero_(values[i][2]),
      quantidade: values[i][3] === null || values[i][3] === undefined ? '' : String(values[i][3]).trim(),
      observacao: String(values[i][4] || '').trim(),
      status: String(values[i][5] || CONFIG.VALORES.STATUS_RASCUNHO).trim()
    });
  }
  return out;
}
