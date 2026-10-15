/**
 * Funções utilitárias reutilizáveis.
 */

/**
 * Nome da organização usado no cabeçalho dos relatórios e PDFs.
 * @return {string}
 */
function obterNomeOrganizacao_() {
  var org = CONFIG.ORGANIZACAO || {};
  return String(org.NOME || '').trim() || 'ORGANIZAÇÃO';
}

/**
 * Retorna a Planilha Mãe.
 * Prefere a planilha JÁ ABERTA (evita timeout do Sheets por openById na própria Mãe).
 * Só usa openById quando o script roda fora da Mãe (ex.: gatilho em filha).
 * @return {Spreadsheet}
 */
function obterPlanilhaMae() {
  var idMae = String(CONFIG.SPREADSHEET_MAE_ID || '').trim();
  try {
    var ativo = SpreadsheetApp.getActiveSpreadsheet();
    if (ativo) {
      if (!idMae || String(ativo.getId()) === idMae) {
        return ativo;
      }
    }
  } catch (eAtivo) {}

  if (idMae) {
    try {
      return SpreadsheetApp.openById(idMae);
    } catch (e) {
      Logger.log('Falha openById mãe: ' + e.message);
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Obtém uma aba pelo nome, ou null se não existir.
 * @param {Spreadsheet} ss
 * @param {string} nome
 * @return {Sheet|null}
 */
function obterAba(ss, nome) {
  return ss.getSheetByName(nome);
}

/**
 * Obtém uma aba ou lança erro descritivo.
 * @param {Spreadsheet} ss
 * @param {string} nome
 * @return {Sheet}
 */
function obterAbaObrigatoria(ss, nome) {
  var aba = obterAba(ss, nome);
  if (!aba) {
    throw new Error('Aba "' + nome + '" não encontrada em: ' + ss.getName());
  }
  return aba;
}

/**
 * Cria a aba se não existir e garante o cabeçalho na linha 1.
 * @param {Spreadsheet} ss
 * @param {string} nome
 * @param {string[]} header
 * @return {Sheet}
 */
function garantirAbaComCabecalho(ss, nome, header) {
  var aba = obterAba(ss, nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
  }
  if (nome === CONFIG.ABAS.PEDIDO) {
    prepararLayoutPedidoUnidade_(aba);
    aba.setFrozenRows(CONFIG.PEDIDO.HEADER_LINHA || 2);
    return aba;
  }
  var rangeHeader = aba.getRange(1, 1, 1, header.length);
  rangeHeader.setValues([header]);
  rangeHeader.setFontWeight('bold');
  // Não força cor clara aqui — o visual (Mãe/unidades) define as cores.
  // Se pintar verde, sobrescreve o tema dark/vermelho.
  aba.setFrozenRows(1);
  return aba;
}

/**
 * Limpa dados a partir da linha 2, preservando o cabeçalho.
 * getRange(row, column, numRows, numColumns) — 3º/4º args são quantidades.
 * @param {Sheet} aba
 * @param {number} numColunas
 */
function limparDadosAba(aba, numColunas) {
  var lastRow = aba.getLastRow();
  if (lastRow >= CONFIG.LINHA_DADOS) {
    var numLinhas = lastRow - CONFIG.LINHA_DADOS + 1;
    aba.getRange(CONFIG.LINHA_DADOS, 1, numLinhas, numColunas).clearContent();
  }
}

/**
 * Normaliza texto para comparação (trim + lowercase).
 * @param {*} valor
 * @return {string}
 */
function normalizarTexto(valor) {
  if (valor === null || valor === undefined) {
    return '';
  }
  return String(valor).trim().toLowerCase();
}

/**
 * Verifica se o valor representa "Sim" / ativo.
 * Aceita Sim, SIM, S, true, 1.
 * @param {*} valor
 * @return {boolean}
 */
function ehAtivo(valor) {
  var v = normalizarTexto(valor)
    .replace(/á|à|ã|â/g, 'a')
    .replace(/é|ê/g, 'e')
    .replace(/í/g, 'i')
    .replace(/ó|ô|õ/g, 'o')
    .replace(/ú/g, 'u')
    .replace(/ç/g, 'c');
  if (v === 'nao' || v === 'n' || v === 'false' || v === '0' || v === 'inativo') {
    return false;
  }
  return v === 'sim' || v === 's' || v === 'true' || v === '1' || v === 'ativo';
}

/**
 * Converte data/célula para string dd/MM/yyyy HH:mm ou valor original.
 * @param {*} valor
 * @return {string}
 */
function formatarData(valor) {
  if (!valor) {
    return '';
  }
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor)) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  }
  return String(valor);
}

/**
 * Preserva zero em estoque/quantidade; só null/undefined viram célula vazia.
 * @param {*} valor
 * @return {*}
 */
function valorCelulaPreservandoZero_(valor) {
  if (valor === null || valor === undefined || valor === '') {
    return '';
  }
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return serialNumericoDeDataSheets_(valor);
  }
  if (typeof valor === 'string') {
    var texto = valor.trim();
    var m = texto.match(/^(\d{1,2})\/(\d{1,2})\/(1899|1900)$/);
    if (m) {
      var dia = parseInt(m[1], 10);
      var mes = parseInt(m[2], 10);
      var ano = parseInt(m[3], 10);
      return serialNumericoDeDataSheets_(new Date(ano, mes - 1, dia));
    }
  }
  return valor;
}

/**
 * Converte Date do Sheets (estoque formatado como data) no número serial original.
 * 30/12/1899 = 0, 29/01/1900 = 30, 18/07/1900 = 200.
 * @param {Date} data
 * @return {number}
 */
function serialNumericoDeDataSheets_(data) {
  var tz = Session.getScriptTimeZone();
  var y = Number(Utilities.formatDate(data, tz, 'yyyy'));
  var m = Number(Utilities.formatDate(data, tz, 'MM'));
  var d = Number(Utilities.formatDate(data, tz, 'dd'));
  var utc = Date.UTC(y, m - 1, d);
  var epoch = Date.UTC(1899, 11, 30);
  return Math.round((utc - epoch) / 86400000);
}

/**
 * Cabeçalho da aba Pedido nas planilhas das unidades (formal + emoji).
 * @return {string[]}
 */
function obterHeaderPedidoUnidade_() {
  return CONFIG.PEDIDO.HEADER_UNIDADE || CONFIG.PEDIDO.HEADER;
}

/**
 * Cabeçalho da aba Resumo na Mãe (estoque x quantidade distintos).
 * @return {string[]}
 */
function obterHeaderResumo_() {
  return CONFIG.RESUMO.HEADER_RESUMO || CONFIG.RESUMO.HEADER;
}

/**
 * Normaliza texto de cabeçalho para comparação (sem emoji/acentos).
 * @param {*} texto
 * @return {string}
 */
function normalizarCabecalhoPedido_(texto) {
  return String(texto || '')
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
 * Linha de faixa de categoria na aba Pedido (▸ ou ▶).
 * @param {string} textoProduto
 * @return {boolean}
 */
function ehLinhaCategoriaPedido_(textoProduto) {
  var t = String(textoProduto || '').trim();
  return t.indexOf('▸ ') === 0 || t.indexOf('▶ ') === 0 || t.indexOf('▸') === 0 || t.indexOf('▶') === 0;
}

/**
 * Garante cabeçalho formal nas unidades (Produto | Quantidade | Observação | Status).
 * @param {Sheet} aba
 */
function garantirCabecalhoPedidoUnidade_(aba) {
  try {
    if (planilhaPedidoUsaLayoutComEstoque_(aba)) {
      migrarPlanilhaPedidoRemoverColunaEstoque_(aba);
      return;
    }
  } catch (eEst) {}

  var header = obterHeaderPedidoUnidade_();
  var headerLinha = CONFIG.PEDIDO.HEADER_LINHA || 2;
  var nCols = obterNumColsPedido_();
  var rowH = aba.getRange(headerLinha, 1, 1, nCols).getValues()[0];
  var produto = normalizarCabecalhoPedido_(rowH[CONFIG.PEDIDO.PRODUTO - 1]);
  var qtd = normalizarCabecalhoPedido_(rowH[CONFIG.PEDIDO.QUANTIDADE - 1]);
  var precisaAtualizar =
    produto.indexOf('produto') === -1 ||
    qtd.indexOf('quantidad') === -1 ||
    String(rowH[CONFIG.PEDIDO.QUANTIDADE - 1] || '').indexOf('🛒') === -1;
  if (precisaAtualizar) {
    aba.getRange(headerLinha, 1, 1, nCols).setValues([header]);
  }
  try {
    aplicarMesReferenciaPedidoUnidade_(aba);
  } catch (eMes) {}
}

/**
 * Linha de item da aba Pedido (4 colunas, sem estoque).
 * @param {string} produto
 * @param {*} quantidade
 * @param {string=} observacao
 * @param {string=} status
 * @return {Array.<*>}
 */
function montarLinhaPedidoUnidade_(produto, quantidade, observacao, status) {
  return [
    produto,
    quantidade === null || quantidade === undefined ? '' : quantidade,
    observacao || '',
    status || CONFIG.VALORES.STATUS_RASCUNHO
  ];
}

/**
 * Faixa de categoria na aba Pedido.
 * @param {string} categoria
 * @return {Array.<string>}
 */
function montarLinhaCategoriaPedido_(categoria) {
  return ['▸ ' + String(categoria || '').trim(), '', '', ''];
}

/**
 * Registra evento na aba Log da Planilha Mãe.
 * @param {string} acao
 * @param {string} detalhe
 */
function registrarLog(acao, detalhe) {
  try {
    var ss = obterPlanilhaMae();
    var aba = ss.getSheetByName(CONFIG.ABAS.LOG);
    if (!aba) {
      aba = ss.insertSheet(CONFIG.ABAS.LOG);
      aba.getRange(1, 1, 1, 3).setValues([CONFIG.LOG.HEADER]);
    }
    aba.appendRow([new Date(), acao, detalhe || '']);
  } catch (e) {
    Logger.log('Falha ao registrar log: ' + e.message);
  }
}

/**
 * Exibe toast amigável na planilha.
 * @param {string} mensagem
 * @param {string=} titulo
 */
function avisarUsuario(mensagem, titulo) {
  try {
    obterPlanilhaMae().toast(mensagem, titulo || 'Sistema de Compras', 8);
  } catch (e) {
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        mensagem,
        titulo || 'Sistema de Compras',
        8
      );
    } catch (e2) {
      Logger.log(mensagem);
    }
  }
}

/**
 * Tenta adquirir um bloqueio exclusivo para operações longas.
 * Faz até 3 tentativas para evitar erro quando outro sync acabou de rodar.
 * @param {string} nome
 * @param {number=} timeoutSegundos
 * @return {Lock}
 */
function adquirirLock(nome, timeoutSegundos) {
  var lock = LockService.getScriptLock();
  var timeoutMs = (timeoutSegundos || 30) * 1000;
  var tentativas = 3;
  for (var t = 0; t < tentativas; t++) {
    if (lock.tryLock(timeoutMs)) {
      return lock;
    }
    Utilities.sleep(2000);
  }
  throw new Error(
    'Outra operação já está em andamento (' +
      nome +
      '). Aguarde 30 segundos e tente de novo.'
  );
}

/**
 * Verifica se o sync automático deve ser ignorado por debounce.
 * @return {boolean} true se deve pular o sync
 */
function syncEmDebounce() {
  var props = PropertiesService.getScriptProperties();
  var ultimo = Number(props.getProperty(CONFIG.PROP_ULTIMO_SYNC) || '0');
  var agora = Date.now();
  if (agora - ultimo < CONFIG.SYNC_DEBOUNCE_MS) {
    return true;
  }
  props.setProperty(CONFIG.PROP_ULTIMO_SYNC, String(agora));
  return false;
}

/**
 * Abre planilha filha só pelo ID (rápido). Sem busca no Drive.
 * @param {string} spreadsheetId
 * @param {string} nomeUnidade
 * @return {Spreadsheet}
 */
function abrirPlanilhaFilhaRapido(spreadsheetId, nomeUnidade) {
  var id = spreadsheetId ? String(spreadsheetId).trim() : '';
  if (!id) {
    throw new Error('ID vazio: ' + nomeUnidade);
  }
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error(
      'Não abriu "' + nomeUnidade + '" (ID inválido?). ' + e.message
    );
  }
}

/**
 * Abre planilha filha pelo ID. Se falhar, busca no Drive por "Pedido - NOME".
 * @param {string} spreadsheetId
 * @param {string} nomeUnidade
 * @return {Spreadsheet}
 */
function abrirPlanilhaFilha(spreadsheetId, nomeUnidade) {
  var id = spreadsheetId ? String(spreadsheetId).trim() : '';
  var ultimoErro = '';

  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      ultimoErro = e.message;
    }
  }

  var titulo = 'Pedido - ' + nomeUnidade;
  try {
    var it = DriveApp.getFilesByName(titulo);
    while (it.hasNext()) {
      var arquivo = it.next();
      if (arquivo.isTrashed()) {
        continue;
      }
      var idCorreto = arquivo.getId();
      try {
        atualizarIdUnidadeNaAba(nomeUnidade, idCorreto);
      } catch (e2) {
        Logger.log('Não atualizou ID na aba Unidades: ' + e2.message);
      }
      registrarLog('ID_CORRIGIDO', nomeUnidade + ' → ' + idCorreto);
      return SpreadsheetApp.openById(idCorreto);
    }
  } catch (e3) {
    ultimoErro = (ultimoErro ? ultimoErro + ' | ' : '') + e3.message;
  }

  throw new Error(
    'Não foi possível abrir a planilha da unidade "' +
      nomeUnidade +
      '" (esperado: "' +
      titulo +
      '"). Verifique o ID na aba Unidades e se o arquivo existe no Drive. Detalhe: ' +
      ultimoErro
  );
}

/**
 * Atualiza o Spreadsheet ID de uma unidade na aba Unidades.
 * @param {string} nomeUnidade
 * @param {string} novoId
 */
function atualizarIdUnidadeNaAba(nomeUnidade, novoId) {
  var unidades = listarTodasUnidades();
  var ss = obterPlanilhaMae();
  var aba = obterAbaObrigatoria(ss, CONFIG.ABAS.UNIDADES);
  var chave = normalizarTexto(nomeUnidade);
  for (var i = 0; i < unidades.length; i++) {
    if (normalizarTexto(unidades[i].nome) === chave) {
      aba.getRange(unidades[i].linha, CONFIG.UNIDADES.SPREADSHEET_ID).setValue(novoId);
      return;
    }
  }
}

/**
 * Lê todas as linhas de dados (a partir da linha 2) de uma aba.
 * getRange(row, column, numRows, numColumns) — 3º/4º args são quantidades.
 * @param {Sheet} aba
 * @param {number} numColunas
 * @return {Array.<Array.<*>>}
 */
function lerDados(aba, numColunas) {
  var lastRow = aba.getLastRow();
  if (lastRow < CONFIG.LINHA_DADOS) {
    return [];
  }
  var numLinhas = lastRow - CONFIG.LINHA_DADOS + 1;
  return aba.getRange(CONFIG.LINHA_DADOS, 1, numLinhas, numColunas).getValues();
}

/**
 * Tenta extrair número de textos como "9KG", "4 kg", "10".
 * @param {*} valor
 * @return {number|null}
 */
function extrairNumeroQuantidade(valor) {
  if (valor === null || valor === undefined || valor === '') {
    return null;
  }
  if (typeof valor === 'number' && !isNaN(valor)) {
    return valor;
  }
  var s = String(valor).trim().replace(',', '.');
  var m = s.match(/(\d+(?:\.\d+)?)/);
  if (!m) {
    return null;
  }
  var n = Number(m[1]);
  return isNaN(n) ? null : n;
}

/**
 * Localiza aba pelo nome exato ou normalizado.
 * @param {Spreadsheet} ssMae
 * @param {string} nomeAba
 * @return {Sheet|null}
 */
function obterAbaPorNome_(ssMae, nomeAba) {
  var aba = ssMae.getSheetByName(nomeAba);
  if (aba) return aba;
  var alvo = normalizarTexto(nomeAba);
  var sheets = ssMae.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (normalizarTexto(sheets[i].getName()) === alvo) {
      return sheets[i];
    }
  }
  return null;
}

/** @param {Spreadsheet} ssMae @return {Sheet|null} */
function obterAbaListaPedidosLegada_(ssMae) {
  return obterAbaPorNome_(ssMae, CONFIG.ABAS.LISTA_PEDIDOS);
}

/**
 * Abas descontinuadas (Lista Pedidos + Painel/gráficos).
 * @return {string[]}
 */
function obterNomesAbasLegadas_() {
  return [CONFIG.ABAS.LISTA_PEDIDOS, CONFIG.ABAS.PAINEL, CONFIG.ABAS.PAINEL_DADOS];
}

/**
 * Remove gatilho e flag do Painel (não existe mais no sistema).
 */
function removerTriggerPainelPendente_() {
  PropertiesService.getScriptProperties().deleteProperty('PAINEL_PENDENTE');
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processarPainelPendente') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Ativa outra aba antes de ocultar/excluir (exigência do Google Sheets).
 * @param {Spreadsheet} ssMae
 * @param {string} nomeEvitar
 */
function ativarOutraAbaAntesDeRemover_(ssMae, nomeEvitar) {
  var preferidas = [
    CONFIG.ABAS.RESUMO,
    CONFIG.ABAS.PRODUTOS,
    CONFIG.ABAS.CONSOLIDADOS,
    CONFIG.ABAS.UNIDADES,
    CONFIG.ABAS.CATEGORIAS
  ];
  for (var p = 0; p < preferidas.length; p++) {
    var abaPref = ssMae.getSheetByName(preferidas[p]);
    if (abaPref && abaPref.getName() !== nomeEvitar) {
      abaPref.activate();
      SpreadsheetApp.flush();
      return;
    }
  }
  var sheets = ssMae.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== nomeEvitar) {
      sheets[i].activate();
      SpreadsheetApp.flush();
      return;
    }
  }
}

/**
 * Remove proteções que impedem ocultar/excluir a aba.
 * @param {Sheet} aba
 */
function liberarProtecoesAba_(aba) {
  try {
    var protSheet = aba.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var i = 0; i < protSheet.length; i++) {
      protSheet[i].remove();
    }
  } catch (e) {}
  try {
    var protRange = aba.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var j = 0; j < protRange.length; j++) {
      protRange[j].remove();
    }
  } catch (e2) {}
}

/**
 * @param {Spreadsheet} ssMae
 * @param {string} nomeAba
 * @return {{ok: boolean, modo: string, erro: string}}
 */
function executarRemocaoAbaPorNome_(ssMae, nomeAba) {
  var resultado = { ok: false, modo: '', erro: '' };
  var aba = obterAbaPorNome_(ssMae, nomeAba);
  if (!aba) {
    resultado.ok = true;
    resultado.modo = 'nao_encontrada';
    return resultado;
  }

  var nomeAtual = aba.getName();
  ativarOutraAbaAntesDeRemover_(ssMae, nomeAtual);
  aba = obterAbaPorNome_(ssMae, nomeAba);
  if (!aba) {
    resultado.ok = true;
    resultado.modo = 'nao_encontrada';
    return resultado;
  }

  liberarProtecoesAba_(aba);

  try {
    aba.getCharts().forEach(function (chart) {
      aba.removeChart(chart);
    });
  } catch (eChart) {}

  try {
    aba.hideSheet();
    SpreadsheetApp.flush();
    registrarLog('ABA_LEGADA_OCULTA', nomeAba);
    resultado.ok = true;
    resultado.modo = 'oculta';
    return resultado;
  } catch (eHide) {
    resultado.erro = eHide.message;
  }

  try {
    if (ssMae.getSheets().length > 1) {
      ssMae.deleteSheet(aba);
      SpreadsheetApp.flush();
      registrarLog('ABA_LEGADA_REMOVIDA', nomeAba);
      resultado.ok = true;
      resultado.modo = 'excluida';
      return resultado;
    }
  } catch (eDel) {
    resultado.erro = (resultado.erro ? resultado.erro + ' | ' : '') + eDel.message;
  }

  registrarLog('ABA_LEGADA_ERRO', nomeAba + ': ' + (resultado.erro || 'erro desconhecido'));
  return resultado;
}

/** @param {Spreadsheet=} ssMae @return {{ok: boolean, modo: string, erro: string}} */
function executarRemocaoAbaListaPedidos_(ssMae) {
  ssMae =
    ssMae ||
    (function () {
      try {
        return SpreadsheetApp.getActiveSpreadsheet();
      } catch (e) {
        return null;
      }
    })() ||
    obterPlanilhaMae();
  return executarRemocaoAbaPorNome_(ssMae, CONFIG.ABAS.LISTA_PEDIDOS);
}

/**
 * Remove abas legadas e automações do Painel (deixa o sistema mais leve).
 * @param {Spreadsheet=} ssMae
 * @return {boolean}
 */
function removerAbasLegadasSistema_(ssMae) {
  ssMae =
    ssMae ||
    (function () {
      try {
        return SpreadsheetApp.getActiveSpreadsheet();
      } catch (e) {
        return null;
      }
    })() ||
    obterPlanilhaMae();

  removerTriggerPainelPendente_();

  var nomes = obterNomesAbasLegadas_();
  var ok = true;
  for (var i = 0; i < nomes.length; i++) {
    var r = executarRemocaoAbaPorNome_(ssMae, nomes[i]);
    if (!r.ok) ok = false;
  }
  return ok;
}

/**
 * Remove a aba legada "Lista Pedidos" (não é mais usada pelo sistema).
 * @param {Spreadsheet=} ssMae
 * @return {boolean}
 */
function removerAbaListaPedidos_(ssMae) {
  return removerAbasLegadasSistema_(ssMae);
}

/** @deprecated use removerAbasLegadasSistema_ */
function ocultarAbaListaPedidos_(ssMae) {
  return removerAbasLegadasSistema_(ssMae);
}

/**
 * Oculta aba técnica após uso (Consolidados/Log).
 * @param {Spreadsheet} ssMae
 * @param {string} nomeAba
 */
function ocultarAbaAdministrativaSeVisivel_(ssMae, nomeAba) {
  try {
    var aba = obterAbaPorNome_(ssMae, nomeAba);
    if (!aba) return;
    if (typeof aba.isSheetHidden === 'function' && aba.isSheetHidden()) return;
    var ativa = ssMae.getActiveSheet();
    if (ativa && normalizarTexto(ativa.getName()) === normalizarTexto(nomeAba)) {
      ativarAbaOperacionalCliente_(ssMae);
    }
    aba.hideSheet();
  } catch (e) {}
}

/**
 * Ativa uma aba operacional (Produtos/Resumo/Categorias/Unidades).
 * @param {Spreadsheet} ssMae
 */
function ativarAbaOperacionalCliente_(ssMae) {
  var visiveis = obterNomesAbasVisiveisCliente_();
  for (var i = 0; i < visiveis.length; i++) {
    var aba = obterAbaPorNome_(ssMae, visiveis[i]);
    if (!aba) continue;
    try {
      if (typeof aba.isSheetHidden === 'function' && aba.isSheetHidden()) continue;
      aba.activate();
      return;
    } catch (e) {}
  }
}

/**
 * Abas visíveis para o cliente (operacional).
 * @return {string[]}
 */
function obterNomesAbasVisiveisCliente_() {
  return [
    CONFIG.ABAS.RESUMO,
    CONFIG.ABAS.TOTAL_PRODUTO,
    CONFIG.ABAS.PRODUTOS,
    CONFIG.ABAS.CATEGORIAS,
    CONFIG.ABAS.UNIDADES
  ];
}

/**
 * Abas técnicas — ficam ocultas, mas o script continua usando normalmente.
 * @return {string[]}
 */
function obterNomesAbasAdministrativas_() {
  return [CONFIG.ABAS.CONSOLIDADOS, CONFIG.ABAS.LOG];
}

/**
 * Oculta uma aba sem apagar (seguro para Consolidados/Log).
 * @param {Spreadsheet} ssMae
 * @param {string} nomeAba
 * @return {boolean}
 */
function ocultarAbaPorNome_(ssMae, nomeAba) {
  var aba = obterAbaPorNome_(ssMae, nomeAba);
  if (!aba) return true;
  if (typeof aba.isSheetHidden === 'function' && aba.isSheetHidden()) return true;

  liberarProtecoesAba_(aba);
  ativarAbaOperacionalCliente_(ssMae);

  aba = obterAbaPorNome_(ssMae, nomeAba);
  if (!aba) return true;
  if (typeof aba.isSheetHidden === 'function' && aba.isSheetHidden()) return true;

  try {
    aba.hideSheet();
    return true;
  } catch (e) {
    registrarLog('ABA_OCULTAR_ERRO', nomeAba + ': ' + e.message);
    return false;
  }
}

/**
 * Garante visão do cliente: Produtos, Resumo, Pedidos das Unidades, Categorias e Unidades visíveis;
 * Consolidados, Log e abas legadas ocultas/removidas.
 * @param {Spreadsheet=} ssMae
 * @return {boolean}
 */
function aplicarVisibilidadeAbasCliente_(ssMae) {
  ssMae =
    ssMae ||
    (function () {
      try {
        return SpreadsheetApp.getActiveSpreadsheet();
      } catch (e) {
        return null;
      }
    })() ||
    obterPlanilhaMae();

  ativarAbaOperacionalCliente_(ssMae);

  try {
    if (typeof obterOuCriarAbaPedidosDasUnidades_ === 'function') {
      obterOuCriarAbaPedidosDasUnidades_(ssMae);
    }
  } catch (eNomePed) {
    Logger.log('visibilidade Pedidos das Unidades: ' + eNomePed.message);
  }

  var admin = obterNomesAbasAdministrativas_();
  var ok = true;
  for (var i = 0; i < admin.length; i++) {
    if (!ocultarAbaPorNome_(ssMae, admin[i])) ok = false;
  }

  removerAbasLegadasSistema_(ssMae);

  var visiveis = obterNomesAbasVisiveisCliente_();
  for (var v = 0; v < visiveis.length; v++) {
    var abaV = obterAbaPorNome_(ssMae, visiveis[v]);
    if (!abaV) continue;
    if (typeof abaV.isSheetHidden === 'function' && abaV.isSheetHidden()) {
      try {
        abaV.showSheet();
      } catch (eShow) {
        ok = false;
      }
    }
  }

  ativarAbaOperacionalCliente_(ssMae);
  return ok;
}

/**
 * Gatilho instalável ao abrir — garante ocultação mesmo se o onOpen simples falhar.
 */
function garantirGatilhoVisibilidadeAbas_() {
  var idMae = String(CONFIG.SPREADSHEET_MAE_ID || '').trim();
  if (!idMae) return;

  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onOpenVisibilidadeAbasInstalavel_') {
      return;
    }
  }

  ScriptApp.newTrigger('onOpenVisibilidadeAbasInstalavel_')
    .forSpreadsheet(idMae)
    .onOpen()
    .create();
}

function onOpenVisibilidadeAbasInstalavel_() {
  try {
    aplicarVisibilidadeAbasCliente_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (e) {
    Logger.log('onOpenVisibilidade: ' + e.message);
  }
  try {
    garantirGatilhosSyncPedidos_();
    removerGatilhosCatalogoAutomatico_();
  } catch (eGat) {
    Logger.log('onOpenVisibilidade gatilhos: ' + eGat.message);
  }
  try {
    pullPedidosAoAbrirMae_();
  } catch (ePull) {
    Logger.log('onOpenVisibilidade pull: ' + ePull.message);
  }
}

/**
 * Menu: aplica a visão operacional (5 abas visíveis).
 */
function menuAplicarVisaoClienteAbas() {
  var ui = SpreadsheetApp.getUi();
  var ok = aplicarVisibilidadeAbasCliente_(SpreadsheetApp.getActiveSpreadsheet());
  ui.alert(
    ok ? 'Visão aplicada' : 'Visão parcial',
    ok
      ? 'O cliente verá apenas:\n• Produtos\n• Resumo\n• Pedidos das Unidades\n• Categorias\n• Unidades\n\nConsolidados e Log ficam ocultos (o sistema continua usando).'
      : 'Algumas abas não puderam ser ajustadas.\n\nClique na aba Resumo e tente de novo.',
    ui.ButtonSet.OK
  );
}

/**
 * Menu: remove abas legadas (Lista Pedidos + Painel de Compras).
 */
function menuRemoverAbasLegadasSistema() {
  menuAplicarVisaoClienteAbas();
}

/** @deprecated */
function menuRemoverAbaListaPedidosLegado() {
  menuRemoverAbasLegadasSistema();
}
