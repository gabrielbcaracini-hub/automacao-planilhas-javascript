/**
 * Gera um guia em PDF para apresentação e treinamento do Sistema de Compras.
 */
function menuGerarPdfGuiaSistema() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = obterPlanilhaMae();
    ss.toast('Gerando guia do sistema…', 'PDF', 60);
    var arquivo = gerarPdfGuiaSistema_(ss);
    ui.alert(
      'Guia do sistema gerado',
      'Arquivo: ' + arquivo.getName() + '\n\nAbra ou baixe pelo link:\n' + arquivo.getUrl(),
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro ao gerar guia', e.message, ui.ButtonSet.OK);
    try {
      registrarLog('GUIA_SISTEMA_PDF_ERRO', e.message);
    } catch (eLog) {}
  }
}

function gerarPdfGuiaSistema_(ss) {
  var agora = new Date();
  var data = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  var nome = 'Guia do Sistema Prefeitura Compras - ' + data;
  var doc = DocumentApp.create(nome);
  var body = doc.getBody();
  body.clear();

  var azul = '#17365D';
  var cinza = '#667085';
  adicionarLogoPrefeituraNoRelatorio_(body);

  body.appendParagraph(obterNomeOrganizacao_()).setAttributes({
    [DocumentApp.Attribute.FOREGROUND_COLOR]: azul,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_SIZE]: 14
  });
  body.appendParagraph('Secretaria Municipal de Assistência e Desenvolvimento Social').setAttributes({
    [DocumentApp.Attribute.FOREGROUND_COLOR]: cinza,
    [DocumentApp.Attribute.FONT_SIZE]: 8
  });
  body.appendParagraph('GUIA OPERACIONAL — SISTEMA DE COMPRAS')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAttributes({
      [DocumentApp.Attribute.FOREGROUND_COLOR]: azul,
      [DocumentApp.Attribute.BOLD]: true
    });
  body.appendParagraph(
    'Documento de apresentação sobre o fluxo de solicitação de produtos entre as unidades e a Planilha Mãe.'
  ).setAttributes({
    [DocumentApp.Attribute.FOREGROUND_COLOR]: cinza,
    [DocumentApp.Attribute.FONT_SIZE]: 9
  });

  adicionarTituloGuia_(body, '1. FLUXO PRINCIPAL');
  body.appendParagraph(
    'Planilha Mãe  →  Sincronizar catálogo  →  Unidade preenche pedido  →  Status Enviado  →  Mãe consolida'
  ).setAttributes({
    [DocumentApp.Attribute.BACKGROUND_COLOR]: '#E9EEF5',
    [DocumentApp.Attribute.FOREGROUND_COLOR]: azul,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_SIZE]: 10
  });
  adicionarTabelaGuia_(body, [
    ['ETAPA', 'COMO FUNCIONA'],
    ['1. Cadastro', 'O administrador cadastra produtos, categoria, pregão e situação Ativo na Planilha Mãe.'],
    ['2. Catálogo', 'Produtos e categorias ativos são enviados para todas as unidades.'],
    ['3. Solicitação', 'Cada unidade informa Quantidade e usa “Enviar pedido completo” (ou marca Enviado item a item).'],
    ['4. Consolidação', 'O pedido chega na Planilha Mãe (Resumo) quando o administrador puxa os pedidos.'],
    ['5. Relatório', 'Unidades baixam Excel do próprio pedido; o administrador gera relatórios consolidados na Mãe.']
  ], azul);

  adicionarTituloGuia_(body, '2. PRODUTOS E CATEGORIAS');
  adicionarTabelaGuia_(body, [
    ['SITUAÇÃO', 'RESULTADO'],
    ['Produto Ativo = Sim', 'Fica disponível para seleção nas unidades.'],
    ['Produto Ativo = Não', 'Sai das unidades após atualizar o catálogo; permanece na aba Produtos da Mãe para eventual reativação.'],
    ['Categoria ativa', 'Produtos daquela categoria aparecem agrupados nas unidades.'],
    ['Categoria inativa', 'Todos os produtos da categoria deixam de aparecer nas unidades, sem apagar o cadastro.']
  ], azul);

  adicionarTituloGuia_(body, '3. NOVA RODADA');
  body.appendParagraph(
    'Use no início de um novo ciclo de compras, após conferir ou guardar o histórico anterior.'
  );
  adicionarTabelaGuia_(body, [
    ['LOCAL', 'AÇÃO EXECUTADA'],
    ['Planilha Mãe', 'Limpa Resumo e Consolidados.'],
    ['Planilhas das unidades', 'Limpa data, quantidade e observação; o status volta para Rascunho.'],
    ['Visual por categoria', 'É mantido nas unidades.'],
    ['Produtos inativos', 'Itens com Ativo = Não permanecem na aba Produtos. Para voltar às unidades: Ativo = Sim e atualizar o catálogo.']
  ], azul);

  adicionarTituloGuia_(body, '4. COMPARTILHAMENTO');
  adicionarTabelaGuia_(body, [
    ['PESSOA', 'ACESSO RECOMENDADO'],
    ['Administrador do sistema', 'Mantém a propriedade da Mãe e das planilhas das unidades.'],
    ['Secretaria responsável', 'Editor da Mãe e das unidades; acompanha pedidos e compartilha cada unidade com seus responsáveis.'],
    ['Responsável pela unidade', 'Editor somente da respectiva planilha de unidade.'],
    ['Regra importante', 'Todos devem usar os arquivos originais compartilhados no Drive; não baixar nem criar cópias operacionais.']
  ], azul);

  adicionarTituloGuia_(body, '5. AUTOMAÇÕES UTILIZADAS');
  adicionarTabelaGuia_(body, [
    ['SCRIPT', 'RESPONSABILIDADE'],
    ['Menu.gs', 'Apresenta as opções operacionais na Planilha Mãe.'],
    ['Produtos.gs / Categorias.gs', 'Gerenciam cadastro, situação ativa e agrupamento por categoria.'],
    ['Sincronizacao.gs / Dropdown.gs', 'Atualizam catálogo, dropdowns e categorias nas unidades.'],
    ['Bridge.gs / Consolidacao.gs', 'Recebem os pedidos enviados e atualizam as abas consolidadas.'],
    ['Resumo.gs', 'Organiza as informações recebidas para conferência administrativa.'],
    ['ResetRodada.gs', 'Prepara o sistema para uma nova rodada de compras.'],
    ['RelatorioPersonalizado.gs', 'Gera o relatório administrativo de pedidos em PDF.']
  ], azul);

  var footer = doc.addFooter();
  footer.appendParagraph(
    'Sistema Prefeitura Compras | Guia gerado em ' +
      Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
  ).setAttributes({
    [DocumentApp.Attribute.FOREGROUND_COLOR]: cinza,
    [DocumentApp.Attribute.FONT_SIZE]: 7
  });

  doc.saveAndClose();
  var pdf = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF).setName(nome + '.pdf');
  var arquivo = obterPastaPdfsResumo_(ss).createFile(pdf);
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  registrarLog('GUIA_SISTEMA_PDF', arquivo.getName() + ' | ' + arquivo.getUrl());
  return arquivo;
}

function adicionarTituloGuia_(body, titulo) {
  body.appendParagraph(titulo)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAttributes({
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#17365D',
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FONT_SIZE]: 11
    });
}

function adicionarTabelaGuia_(body, valores, azul) {
  var tabela = body.appendTable(valores);
  for (var c = 0; c < valores[0].length; c++) {
    tabela.getCell(0, c).setAttributes({
      [DocumentApp.Attribute.BACKGROUND_COLOR]: azul,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#FFFFFF',
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FONT_SIZE]: 8
    });
  }
  for (var r = 1; r < tabela.getNumRows(); r++) {
    for (var col = 0; col < valores[0].length; col++) {
      tabela.getCell(r, col).setAttributes({
        [DocumentApp.Attribute.BACKGROUND_COLOR]: r % 2 === 0 ? '#F2F4F7' : '#FFFFFF',
        [DocumentApp.Attribute.FONT_SIZE]: 8
      });
    }
  }
}
