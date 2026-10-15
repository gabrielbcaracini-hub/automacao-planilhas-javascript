/**
 * Cadastro e leitura das unidades (planilhas filhas).
 */

/**
 * Lista todas as unidades cadastradas na aba Unidades.
 * @return {Array.<Object>}
 */
function listarTodasUnidades(ssOpt) {
  var ss = ssOpt || obterPlanilhaMae();
  var aba = obterAbaObrigatoria(ss, CONFIG.ABAS.UNIDADES);
  var linhas = lerDados(aba, CONFIG.UNIDADES.HEADER.length);
  var unidades = [];

  for (var i = 0; i < linhas.length; i++) {
    var row = linhas[i];
    var nome = String(row[CONFIG.UNIDADES.NOME - 1] || '').trim();
    var id = String(row[CONFIG.UNIDADES.SPREADSHEET_ID - 1] || '').trim();

    if (!nome && !id) {
      continue;
    }

    unidades.push({
      linha: i + CONFIG.LINHA_DADOS,
      nome: nome,
      spreadsheetId: id,
      ativo: ehAtivo(row[CONFIG.UNIDADES.ATIVO - 1])
    });
  }

  return unidades;
}

/**
 * Retorna somente unidades ativas com ID preenchido.
 * @return {Array.<Object>}
 */
function listarUnidadesAtivas(ssOpt) {
  var lista = listarTodasUnidades(ssOpt).filter(function (u) {
    return u.ativo && u.nome && u.spreadsheetId;
  });
  if (lista.length > 0) return lista;
  if (!DADOS_INICIAIS || !DADOS_INICIAIS.unidades) return lista;
  return DADOS_INICIAIS.unidades
    .map(function (nome) {
      return {
        nome: nome,
        spreadsheetId: (DADOS_INICIAIS.idsPlanilhas && DADOS_INICIAIS.idsPlanilhas[nome]) || '',
        ativo: true
      };
    })
    .filter(function (u) {
      return u.nome && u.spreadsheetId;
    });
}
