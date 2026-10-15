/**
 * Controle do MÊS: quais categorias (e seus pregões) estão liberadas.
 * O catálogo completo fica na aba Produtos; unidades e Resumo só veem o que está Sim.
 */

/**
 * Lê a disponibilidade por categoria. Sem aba Categorias, mantém compatibilidade:
 * todas as categorias permanecem disponíveis.
 * @return {Object.<string, boolean>}
 */
function obterDisponibilidadeCategorias_() {
  var ss = obterPlanilhaMae();
  var aba = ss.getSheetByName(CONFIG.ABAS.CATEGORIAS);
  var mapa = {};
  if (!aba || aba.getLastRow() < CONFIG.LINHA_DADOS) return mapa;

  var valores = aba
    .getRange(
      CONFIG.LINHA_DADOS,
      1,
      aba.getLastRow() - CONFIG.LINHA_DADOS + 1,
      CONFIG.CATEGORIAS.HEADER.length
    )
    .getValues();
  for (var i = 0; i < valores.length; i++) {
    var categoria = String(valores[i][CONFIG.CATEGORIAS.CATEGORIA - 1] || '').trim();
    if (!categoria) continue;
    mapa[normalizarTexto(categoria)] = ehAtivo(
      valores[i][CONFIG.CATEGORIAS.ATIVO - 1]
    );
  }
  return mapa;
}

/**
 * Categoria liberada neste mês (aba Categorias = Sim).
 * Sem cadastro ainda, libera tudo (instalação nova).
 * Categoria nova, ainda não listada, fica bloqueada até a Mãe marcar Sim.
 * @param {string} categoria
 * @param {Object.<string, boolean>=} mapa
 * @return {boolean}
 */
function categoriaLiberadaNoMes_(categoria, mapa) {
  mapa = mapa || {};
  var chaves = Object.keys(mapa);
  if (!chaves.length) return true;
  var chave = normalizarTexto(categoria || 'SEM CATEGORIA');
  if (!Object.prototype.hasOwnProperty.call(mapa, chave)) return false;
  return mapa[chave] !== false;
}

/**
 * Recria a lista de categorias a partir da aba Produtos, preservando Sim/Não já definido.
 * Categoria que ainda não estava na lista entra como Sim.
 * @return {number}
 */
function sincronizarCadastroCategorias_() {
  var ss = obterPlanilhaMae();
  var disponibilidadeAtual = obterDisponibilidadeCategorias_();
  var temCadastro = Object.keys(disponibilidadeAtual).length > 0;
  var categorias = {};
  var todos = listarTodosProdutos();

  for (var i = 0; i < todos.length; i++) {
    var categoria = String(todos[i].categoria || '').trim() || 'SEM CATEGORIA';
    categorias[normalizarTexto(categoria)] = categoria;
  }

  var lista = Object.keys(categorias).map(function (chave) {
    return categorias[chave];
  });
  lista.sort(function (a, b) {
    return a.localeCompare(b, 'pt-BR');
  });

  var linhas = lista.map(function (categoria) {
    var chave = normalizarTexto(categoria);
    var disponivel;
    if (Object.prototype.hasOwnProperty.call(disponibilidadeAtual, chave)) {
      disponivel = disponibilidadeAtual[chave]
        ? CONFIG.VALORES.ATIVO_SIM
        : CONFIG.VALORES.ATIVO_NAO;
    } else if (!temCadastro) {
      disponivel = CONFIG.VALORES.ATIVO_SIM;
    } else {
      // Já está na aba Produtos e ainda não tinha linha em Categorias.
      // Entra como Sim para aparecer no Resumo. Não mexe em categoria já marcada Não.
      disponivel = CONFIG.VALORES.ATIVO_SIM;
    }
    return [categoria, disponivel];
  });

  escreverAbaSimples_(ss, CONFIG.ABAS.CATEGORIAS, CONFIG.CATEGORIAS.HEADER, linhas);
  var aba = ss.getSheetByName(CONFIG.ABAS.CATEGORIAS);
  aplicarValidacaoAtivo(aba, CONFIG.CATEGORIAS.ATIVO);
  aba.setColumnWidth(CONFIG.CATEGORIAS.CATEGORIA, 420);
  aba.setColumnWidth(CONFIG.CATEGORIAS.ATIVO, 190);
  return linhas.length;
}

/**
 * Menu: atualiza e abre a aba onde a Mãe libera/bloqueia categorias inteiras.
 */
function menuAtualizarControleCategorias() {
  var ui = SpreadsheetApp.getUi();
  try {
    var total = sincronizarCadastroCategorias_();
    var ss = obterPlanilhaMae();
    // Usa a planilha aberta na tela para trocar de aba de forma visível.
    var ssTela = SpreadsheetApp.getActiveSpreadsheet();
    var abaTela =
      ssTela && ssTela.getId() === ss.getId()
        ? ssTela.getSheetByName(CONFIG.ABAS.CATEGORIAS)
        : null;
    if (abaTela) {
      ssTela.setActiveSheet(abaTela);
    }
    ui.alert(
      'Categorias do mês',
      total +
        ' categoria(s) na aba Categorias.\n\n' +
        'O catálogo completo continua na aba Produtos.\n' +
        'No MÊS, unidades e Resumo só usam o que estiver Sim.\n\n' +
        '• Sim = categoria (e os pregões dela) liberada neste mês\n' +
        '• Não = some das unidades e do Resumo (fica só no catálogo)\n' +
        '• Categoria que ainda não estava na lista entra como Sim\n\n' +
        'Ex.: só alimentos neste mês → Sim em GÊNERO ALIMENTAÇÃO / ALIMENTÍCIO;\n' +
        'o resto em Não.\n\n' +
        'Depois de alterar Sim/Não, use no menu:\n' +
        '“⚡ Atualizar produtos nas unidades agora”.',
      ui.ButtonSet.OK
    );
    var syncAgora = ui.alert(
      'Enviar para as unidades agora?',
      'Deseja aplicar categorias + produtos ativos/inativos em todas as unidades agora?',
      ui.ButtonSet.YES_NO
    );
    if (syncAgora === ui.Button.YES) {
      menuAtualizarProdutosUnidadesAgora();
    }
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Coloca no Resumo os produtos ativos que ficaram de fora
 * porque a categoria existia na aba Produtos e não na aba Categorias.
 * Roda uma vez ao abrir a planilha mãe.
 */
function completarResumoComProdutosDaAba_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss || String(ss.getId()) !== String(CONFIG.SPREADSHEET_MAE_ID || '').trim()) return;
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('RESUMO_PRODUTOS_COMPLETOS_2309') === '1') return;

  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var nome = String(sheets[i].getName() || '');
    if (nome.indexOf('Resumo_SEM_VISUAL') === 0) {
      try { sheets[i].hideSheet(); } catch (eHide) {}
    }
  }

  sincronizarCadastroCategorias_();
  // Não reconstrói o Resumo a partir de Consolidados: essa aba pode estar vazia
  // e apagaria quantidades já visíveis. O menu Puxar pedidos grava as quantidades.
  props.setProperty('RESUMO_PRODUTOS_COMPLETOS_2309', '1');
  try {
    ss.toast('Resumo atualizado com os produtos ativos da aba Produtos.', 'Resumo', 8);
  } catch (eT) {}
}

