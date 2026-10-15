/**
 * Configurações centralizadas do sistema de solicitação de compras.
 * Altere nomes de abas, colunas e comportamentos apenas neste arquivo.
 */

var CONFIG = {
  /** ID da Planilha Mãe (obrigatório para sync automático a partir das filhas). */
  SPREADSHEET_MAE_ID: '',

  /**
   * Identificação da organização nos relatórios e PDFs gerados.
   * Preencha com os dados da sua instituição.
   */
  ORGANIZACAO: {
    /** Texto do cabeçalho dos relatórios. */
    NOME: 'PREFEITURA MUNICIPAL DE EXEMPLO',
    /** URL pública do logotipo (deixe vazio para gerar sem imagem). */
    LOGO_URL: ''
  },

  ABAS: {
    PRODUTOS: 'Produtos',
    CATEGORIAS: 'Categorias',
    UNIDADES: 'Unidades',
    LISTA_PEDIDOS: 'Lista Pedidos', // legado — removida automaticamente
    CONSOLIDADOS: 'Consolidados', // oculta para o cliente — scripts continuam usando
    RESUMO: 'Resumo',
    RESUMO_LEGADO: 'Resumo_LEGADO',
    /** Nome temporário usado na migração anti-timeout (pode existir na planilha). */
    RESUMO_PREGOES_TEMP: 'Resumo Pregões',
    TOTAL_PRODUTO: 'Pedidos das Unidades',
    TOTAL_PRODUTO_LEGADO: 'Total por Produto',
    PAINEL: 'Painel de Compras', // legado — removida automaticamente
    PAINEL_DADOS: 'Painel Dados', // legado — removida automaticamente
    LOG: 'Log', // oculta para o cliente — auditoria técnica
    PEDIDO: 'Pedido',
    CATALOGO: 'Catalogo'
  },

  /** Colunas da aba Produtos (1-based para getRange). */
  PRODUTOS: {
    CODIGO: 1,
    CATEGORIA: 2,
    PRODUTO: 3,
    UNIDADE_MEDIDA: 4,
    PREGAO: 5,
    ATIVO: 6,
    HEADER: ['Código', 'Categoria', 'Produto', 'Unidade', 'Pregão', 'Ativo']
  },

  /** Controle de disponibilidade de categorias nas planilhas das unidades. */
  CATEGORIAS: {
    CATEGORIA: 1,
    ATIVO: 2,
    HEADER: ['Categoria', 'Liberada neste mês']
  },

  /** Colunas da aba Unidades. */
  UNIDADES: {
    NOME: 1,
    SPREADSHEET_ID: 2,
    ATIVO: 3,
    HEADER: ['Nome Unidade', 'Spreadsheet ID', 'Ativo']
  },

  /** Colunas da aba Pedido (planilhas filhas — sem Data e sem Estoque Atual). */
  PEDIDO: {
    MES_REFERENCIA_LINHA: 1,
    HEADER_LINHA: 2,
    LINHA_DADOS: 3,
    NUM_COLS: 4,
    PRODUTO: 1,
    QUANTIDADE: 2,
    OBSERVACAO: 3,
    STATUS: 4,
    /** Cabeçalho interno / legado (detecção de colunas). */
    HEADER: ['Produto', 'Quantidade', 'Observação', 'Status'],
    /** Cabeçalho exibido nas planilhas das UNIDADES. */
    HEADER_UNIDADE: [
      'Produto',
      '🛒 Quantidade do Pedido',
      'Observação',
      'Status'
    ]
  },

  /** Colunas da aba Consolidados (Planilha Mãe). */
  CONSOLIDADOS: {
    UNIDADE: 1,
    PRODUTO: 2,
    ESTOQUE_ATUAL: 3,
    QUANTIDADE: 4,
    OBSERVACAO: 5,
    DATA: 6,
    HEADER: ['Unidade', 'Produto', 'Estoque Atual', 'Quantidade', 'Observação', 'Data']
  },

  /** Colunas da aba Resumo — layout dinâmico: ITEM + uma coluna por unidade. */
  RESUMO: {
    PRODUTO: 1,
    ESTOQUE_ATUAL: 2,
    QUANTIDADE: 3,
    OBSERVACAO: 4,
    NUM_COLS: 4,
    /** Cabeçalho legado (antes do layout por pregão). */
    HEADER: ['Produto', 'Estoque Atual', 'Quantidade', 'Observação'],
    /** Cabeçalho legado exibido nas unidades / versões antigas. */
    HEADER_RESUMO: [
      'Produto',
      '📦 Estoque Atual na Unidade',
      '🛒 Quantidade Requisitada',
      'Observação'
    ]
  },

  /** Colunas da aba Pedidos das Unidades — situação dos pedidos por unidade. */
  TOTAL_PRODUTO: {
    CONCLUIDAS: 1,
    PENDENTES: 2,
    HEADER: ['Já pediram', 'Ainda não pediram']
  },

  /** Colunas da aba Log. */
  LOG: {
    DATA: 1,
    ACAO: 2,
    DETALHE: 3,
    HEADER: ['Data/Hora', 'Ação', 'Detalhe']
  },

  /** Valores padronizados de status e ativo. */
  VALORES: {
    ATIVO_SIM: 'Sim',
    ATIVO_NAO: 'Não',
    STATUS_RASCUNHO: 'Rascunho',
    STATUS_ENVIADO: 'Enviado',
    STATUS_CONSOLIDADO: 'Consolidado'
  },

  /**
   * Intervalo mínimo (ms) entre syncs automáticos via onEdit.
   * Evita múltiplas sincronizações ao colar várias linhas.
   */
  SYNC_DEBOUNCE_MS: 8000,

  /** Chave em PropertiesService para controle de debounce do sync. */
  PROP_ULTIMO_SYNC: 'ULTIMO_SYNC_TS',

  /** Horário do trigger diário de consolidação (0–23). */
  CONSOLIDACAO_HORA: 7,

  /** Linha inicial de dados (após o cabeçalho). */
  LINHA_DADOS: 2,

  /** Quantidade de linhas de validação no dropdown do Pedido. */
  PEDIDO_LINHAS_VALIDACAO: 800,

  /**
   * Tema visual da Mãe (modelo fixo: cabeçalho vermelho + coluna destaque amarela + dark).
   * Aba Produtos usa o mesmo padrão em TODAS as linhas (código escuro, categoria ciano, produto amarelo).
   */
  TEMA: {
    HEADER_BG: '#FF0000',
    HEADER_FG: '#FFFFFF',
    COL_DESTAQUE_BG: '#FFFF00',
    COL_DESTAQUE_FG: '#000000',
    COL_CATEGORIA_BG: '#00FFFF',
    COL_CATEGORIA_FG: '#000000',
    ROW_A: '#0D1117',
    ROW_B: '#161B22',
    TEXT: '#FFFFFF',
    TEXT_MUTED: '#E6EDF3',
    BORDER: '#30363D',
    ACCENT_QTD: '#FFFFFF',
    SHEET_BG: '#010409',
    TAB: '#FF0000',
    /** Coluna Estoque — legado (Resumo antigo na Mãe). */
    COL_ESTOQUE_HEADER_BG: '#2563EB',
    COL_ESTOQUE_BG: '#DBEAFE',
    COL_ESTOQUE_FG: '#1E3A8A',
    /** Coluna Quantidade nas unidades — âmbar (o que será pedido). */
    COL_QTD_HEADER_BG: '#D97706',
    COL_QTD_BG: '#FEF3C7',
    COL_QTD_FG: '#92400E',
    /** Coluna Status nas unidades — destaque igual às demais colunas. */
    COL_STATUS_HEADER_BG: '#4F46E5',
    COL_STATUS_RASCUNHO_BG: '#E2E8F0',
    COL_STATUS_RASCUNHO_FG: '#0F172A',
    COL_STATUS_ENVIADO_BG: '#86EFAC',
    COL_STATUS_ENVIADO_FG: '#14532D',
    COL_STATUS_CONSOLIDADO_BG: '#93C5FD',
    COL_STATUS_CONSOLIDADO_FG: '#1E3A8A'
  }
};
