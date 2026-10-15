/**
 * Pregões e produtos da matriz 11/09/2026 que ainda não existiam no sistema.
 * Não substitui itens já cadastrados — só acrescenta os novos.
 */

var NOVOS_PREGOES_MATRIZ_SETEMBRO_2026 = [
  {
    categoria: 'ARES CONDICIONADOS',
    produto: 'AR SPLIT 24.000 BTUS + INSTALAÇÃO',
    pregao: 'PREGÃO 14/2025 – DE 20/03/2025 – FORNECEDOR EXEMPLO I LTDA – 19/03/2027',
    ativo: 'Sim'
  },
  {
    categoria: 'ARES CONDICIONADOS',
    produto: 'AR SPLIT 30.000 BTUS + INSTALAÇÃO',
    pregao: 'PREGÃO 14/2025 – DE 20/03/2025 – FORNECEDOR EXEMPLO I LTDA – 19/03/2027',
    ativo: 'Sim'
  },
  {
    categoria: 'ELETRODOMÉSTICOS',
    produto: 'LIQUIDIFICADOR INDUSTRIAL 1200 W',
    pregao: 'PREGÃO 156/2025 – DE 29/01/2026 – FORNECEDOR EXEMPLO H LTDA – 29/01/2027',
    ativo: 'Sim'
  },
  {
    categoria: 'ELETRODOMÉSTICOS',
    produto: 'REFRIGERADOR FROST FREE 310 LITROS',
    pregao: 'PREGÃO 156/2025 – DE 31/08/2026 – FORNECEDOR EXEMPLO K LTDA – 29/01/2027',
    ativo: 'Sim'
  },
  {
    categoria: 'GÊNERO ALIMENTÍCIO',
    produto: 'PÃO FRANCÊS KG',
    pregao: 'PREGÃO 49/2026 – DE 13/07/2026 – FORNECEDOR EXEMPLO L LTDA – 02/07/2027',
    ativo: 'Sim'
  },
  {
    categoria: 'GÊNERO ALIMENTÍCIO',
    produto: 'LEITE INTEGRAL CAIXA 1 LITRO',
    pregao: 'PREGÃO 49/2026 – DE 13/07/2026 – FORNECEDOR EXEMPLO L LTDA – 02/07/2027',
    ativo: 'Sim'
  },
  {
    categoria: 'MATERIAL DE LIMPEZA',
    produto: 'DESINFETANTE CONCENTRADO 5 LITROS',
    pregao: 'PREGÃO 01/2026 – DE 10/03/2026 – FORNECEDOR EXEMPLO M LTDA – 10/03/2027',
    ativo: 'Sim'
  },
  {
    categoria: 'MATERIAL DE LIMPEZA',
    produto: 'SACO DE LIXO 100 LITROS PACOTE 100 UN',
    pregao: 'PREGÃO 01/2026 – DE 10/03/2026 – FORNECEDOR EXEMPLO M LTDA – 10/03/2027',
    ativo: 'Sim'
  },
  {
    categoria: 'INFORMÁTICA',
    produto: 'TECLADO USB ABNT2',
    pregao: 'PREGÃO 63/2026 – DE 31/07/2026 – FORNECEDOR EXEMPLO N ME – 29/07/2027',
    ativo: 'Sim'
  },
  {
    categoria: 'PAPELARIA',
    produto: 'ENVELOPE OFÍCIO PACOTE 100 UN',
    pregao: 'PREGÃO 63/2026 – DE 31/07/2026 – FORNECEDOR EXEMPLO N ME – 29/07/2027',
    ativo: 'Sim'
  }
];

/**
 * Acrescenta só as linhas novas na aba Produtos (não mexe no que já existe).
 * @param {Spreadsheet=} ss
 * @return {{adicionados:number, ignorados:number, pregoes:string[], mensagem:string}}
 */
function acrescentarLinhasNovosPregoesMatriz_(ss) {
  ss = ss || obterPlanilhaMae();
  var aba = obterAbaObrigatoria(ss, CONFIG.ABAS.PRODUTOS);
  var map = obterMapaColunasProdutos_(aba);
  var existentes = listarTodosProdutos();
  var nomes = {};
  var maxCod = 0;
  var i;
  for (i = 0; i < existentes.length; i++) {
    var nomeEx = String(existentes[i].produto || '').trim();
    if (nomeEx) nomes[normalizarTexto(nomeEx)] = true;
    var codNum = parseInt(String(existentes[i].codigo || '').replace(/\D/g, ''), 10);
    if (codNum && codNum > maxCod) maxCod = codNum;
  }
  var pregoesExistentes = {};
  for (i = 0; i < existentes.length; i++) {
    var txtP = String(existentes[i].pregao || '');
    var mP = txtP.match(/PREG[AÃÁÂÀ]?O\s*(\d+\s*\/\s*\d{2,4})/i);
    if (mP) pregoesExistentes[String(mP[1]).replace(/\s+/g, '')] = true;
  }
  var lista = NOVOS_PREGOES_MATRIZ_SETEMBRO_2026 || [];
  var linhas = [];
  var ignorados = 0;
  var pregoesAdd = {};
  var nCols = Math.max(aba.getLastColumn(), CONFIG.PRODUTOS.HEADER.length, map.colAtivo + 1, 6);
  for (i = 0; i < lista.length; i++) {
    var item = lista[i];
    var produto = String(item.produto || '').trim();
    var pregao = String(item.pregao || '').trim();
    var categoria = String(item.categoria || '').trim() || 'SEM CATEGORIA';
    if (!produto || !pregao) {
      ignorados++;
      continue;
    }
    var mNovo = pregao.match(/PREG[AÃÁÂÀ]?O\s*(\d+\s*\/\s*\d{2,4})/i);
    var numNovo = mNovo ? String(mNovo[1]).replace(/\s+/g, '') : '';
    if (numNovo && pregoesExistentes[numNovo]) {
      ignorados++;
      continue;
    }
    if (nomes[normalizarTexto(produto)]) {
      ignorados++;
      continue;
    }
    nomes[normalizarTexto(produto)] = true;
    maxCod++;
    var row = [];
    var c;
    for (c = 0; c < nCols; c++) row.push('');
    row[map.colCodigo] = String(maxCod);
    row[map.colCategoria] = categoria;
    row[map.colProduto] = produto;
    row[map.colPregao] = pregao;
    row[map.colAtivo] = CONFIG.VALORES.ATIVO_SIM;
    linhas.push(row);
    if (numNovo) pregoesAdd[numNovo] = true;
  }
  if (linhas.length) {
    var dest = Math.max(aba.getLastRow(), 1) + 1;
    aba.getRange(dest, 1, linhas.length, nCols).setValues(linhas);
    try {
      aplicarValidacaoAtivoProdutos_(aba);
    } catch (eVal) {}
  }
  var chaves = Object.keys(pregoesAdd);
  chaves.sort();
  var mensagem =
    'Adicionados ' + linhas.length + ' produto(s) de ' + chaves.length + ' pregão(ões) novo(s).';
  return {
    adicionados: linhas.length,
    ignorados: ignorados,
    pregoes: chaves,
    mensagem: mensagem
  };
}

/**
 * Roda uma vez ao abrir a Mãe: cadastra pregões novos, organiza categorias e atualiza o Resumo.
 */
function importarMatrizSetembro2026SePendente_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss || String(ss.getId()) !== String(CONFIG.SPREADSHEET_MAE_ID || '').trim()) return;
  var props = PropertiesService.getDocumentProperties();
  var st = String(props.getProperty('MATRIZ_SETEMBRO_2026') || '');
  if (st === 'OK') return;

  if (st !== 'PRODUTOS') {
    var r = acrescentarLinhasNovosPregoesMatriz_(ss);
    props.setProperty('MATRIZ_SETEMBRO_2026', 'PRODUTOS');
    try {
      ss.toast((r && r.mensagem) || 'Pregões novos cadastrados.', 'Matriz', 10);
    } catch (eT) {}
    try {
      registrarLog('IMPORT_MATRIZ', (r && r.mensagem) || '');
    } catch (eL) {}
  }
  try {
    organizarProdutosPorCategoria_(ss);
  } catch (eOrg) {
    registrarLog('IMPORT_MATRIZ_ORG', eOrg.message);
  }
  try {
    sincronizarCadastroCategorias_();
  } catch (eCat) {
    registrarLog('IMPORT_MATRIZ_CAT', eCat.message);
  }
  try {
    reconstruirResumoDeConsolidadosExistentes_(ss);
  } catch (eRes) {
    registrarLog('IMPORT_MATRIZ_RESUMO', eRes.message);
    return;
  }
  props.setProperty('MATRIZ_SETEMBRO_2026', 'OK');
}

/**
 * Acrescenta pregões novos na aba Produtos, reorganiza categorias e atualiza o Resumo.
 * @return {{adicionados:number, ignorados:number, pregoes:string[], mensagem:string}}
 */
function importarNovosPregoesMatrizSetembro2026() {
  var ss = obterPlanilhaMae();
  var r = acrescentarLinhasNovosPregoesMatriz_(ss);
  try {
    organizarProdutosPorCategoria_(ss);
  } catch (eOrg) {
    registrarLog('IMPORT_MATRIZ_ORG', eOrg.message);
  }
  try {
    sincronizarCadastroCategorias_();
  } catch (eCat) {
    registrarLog('IMPORT_MATRIZ_CAT', eCat.message);
  }
  try {
    reconstruirResumoDeConsolidadosExistentes_(ss);
  } catch (eRes) {
    registrarLog('IMPORT_MATRIZ_RESUMO', eRes.message);
  }
  try {
    PropertiesService.getDocumentProperties().setProperty('MATRIZ_SETEMBRO_2026', 'OK');
  } catch (eProp) {}
  registrarLog('IMPORT_MATRIZ', r.mensagem + ' | ' + (r.pregoes || []).join(', '));
  return r;
}

function menuImportarNovosPregoesMatrizSetembro2026() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Importar pregões novos',
    'Vai acrescentar os pregões da matriz 11/09/2026 que ainda não existem no sistema,' +
      ' com os produtos de cada um, e atualizar o Resumo.\n\n' +
      'Os pregões que já estão cadastrados não serão alterados.\n\nContinuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Importando pregões novos…', 'Matriz', 60);
    var r = importarNovosPregoesMatrizSetembro2026();
    ui.alert('Pregões importados', r.mensagem + '\nIgnorados: ' + r.ignorados, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Erro na importação', e.message, ui.ButtonSet.OK);
  }
}

