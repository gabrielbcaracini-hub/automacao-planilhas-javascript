/**
 * Dados iniciais de demonstração.
 *
 * ATENÇÃO: este arquivo contém APENAS dados fictícios, usados para o setup
 * inicial em um ambiente de testes. Substitua pelo seu próprio conteúdo antes
 * de usar o sistema de verdade.
 *
 * - idsPlanilhas: preencha com o ID de cada planilha de unidade. Deixe vazio
 *   para que `menuImportarUnidadesCompleto` crie as planilhas e preencha os IDs.
 * - unidades: define TAMBÉM a ordem oficial em que as unidades aparecem nas
 *   abas Resumo e Pedidos das Unidades.
 * - produtos: catálogo semente gravado na aba Produtos quando ela está vazia.
 */

var DADOS_INICIAIS = {
  idsPlanilhas: {
    'UNIDADE CENTRO': '',
    'UNIDADE NORTE': '',
    'UNIDADE SUL': '',
    'UNIDADE LESTE': '',
    'UNIDADE OESTE': '',
    'NÚCLEO SOCIAL I': '',
    'NÚCLEO SOCIAL II': '',
    'CENTRO DE CONVIVÊNCIA': '',
    'CENTRO COMUNITÁRIO': '',
    'CASA DE APOIO': '',
    'POSTO DE ATENDIMENTO': '',
    'ALMOXARIFADO CENTRAL': '',
    'SEDE ADMINISTRATIVA': '',
    'UNIDADE MODELO': ''
  },
  unidades: [
    'UNIDADE CENTRO',
    'UNIDADE NORTE',
    'UNIDADE SUL',
    'UNIDADE LESTE',
    'UNIDADE OESTE',
    'NÚCLEO SOCIAL I',
    'NÚCLEO SOCIAL II',
    'CENTRO DE CONVIVÊNCIA',
    'CENTRO COMUNITÁRIO',
    'CASA DE APOIO',
    'POSTO DE ATENDIMENTO',
    'ALMOXARIFADO CENTRAL',
    'SEDE ADMINISTRATIVA',
    'UNIDADE MODELO'
  ],
  produtos: [
    {
      codigo: '001',
      categoria: 'PAPELARIA',
      produto: 'PAPEL SULFITE BRANCO A4 RESMA',
      pregao: 'PREGÃO 10/2025 – DE 15/03/2025 – FORNECEDOR EXEMPLO A LTDA – 14/03/2027',
      ativo: 'Sim'
    },
    {
      codigo: '002',
      categoria: 'PAPELARIA',
      produto: 'CANETA ESFEROGRÁFICA AZUL CAIXA 50 UN',
      pregao: 'PREGÃO 10/2025 – DE 15/03/2025 – FORNECEDOR EXEMPLO A LTDA – 14/03/2027',
      ativo: 'Sim'
    },
    {
      codigo: '003',
      categoria: 'PAPELARIA',
      produto: 'PASTA SUSPENSA KRAFT UNIDADE',
      pregao: 'PREGÃO 10/2025 – DE 15/03/2025 – FORNECEDOR EXEMPLO A LTDA – 14/03/2027',
      ativo: 'Sim'
    },
    {
      codigo: '004',
      categoria: 'MATERIAL DE EXPEDIENTE',
      produto: 'GRAMPEADOR DE MESA 26/6',
      pregao: 'PREGÃO 22/2025 – DE 02/05/2025 – FORNECEDOR EXEMPLO B EIRELI – 01/05/2027',
      ativo: 'Sim'
    },
    {
      codigo: '005',
      categoria: 'MATERIAL DE EXPEDIENTE',
      produto: 'TESOURA ESCOLAR 13 CM',
      pregao: 'PREGÃO 22/2025 – DE 02/05/2025 – FORNECEDOR EXEMPLO B EIRELI – 01/05/2027',
      ativo: 'Sim'
    },
    {
      codigo: '006',
      categoria: 'GÊNERO ALIMENTÍCIO',
      produto: 'ARROZ TIPO 1 PACOTE 5 KG',
      pregao: 'PREGÃO 35/2025 – DE 10/06/2025 – FORNECEDOR EXEMPLO C LTDA – 09/06/2027',
      ativo: 'Sim'
    },
    {
      codigo: '007',
      categoria: 'GÊNERO ALIMENTÍCIO',
      produto: 'FEIJÃO CARIOCA PACOTE 1 KG',
      pregao: 'PREGÃO 35/2025 – DE 10/06/2025 – FORNECEDOR EXEMPLO C LTDA – 09/06/2027',
      ativo: 'Sim'
    },
    {
      codigo: '008',
      categoria: 'GÊNERO ALIMENTÍCIO',
      produto: 'ÓLEO DE SOJA FRASCO 900 ML',
      pregao: 'PREGÃO 35/2025 – DE 10/06/2025 – FORNECEDOR EXEMPLO C LTDA – 09/06/2027',
      ativo: 'Sim'
    },
    {
      codigo: '009',
      categoria: 'GÊNERO HORTIFRUTI',
      produto: 'BANANA NANICA KG',
      pregao: 'PREGÃO 41/2025 – DE 20/07/2025 – FORNECEDOR EXEMPLO D ME – 19/07/2027',
      ativo: 'Sim'
    },
    {
      codigo: '010',
      categoria: 'GÊNERO HORTIFRUTI',
      produto: 'BATATA LAVADA KG',
      pregao: 'PREGÃO 41/2025 – DE 20/07/2025 – FORNECEDOR EXEMPLO D ME – 19/07/2027',
      ativo: 'Sim'
    },
    {
      codigo: '011',
      categoria: 'MATERIAL DE LIMPEZA',
      produto: 'ÁGUA SANITÁRIA 1 LITRO',
      pregao: 'PREGÃO 58/2025 – DE 05/09/2025 – FORNECEDOR EXEMPLO E LTDA – 04/09/2027',
      ativo: 'Sim'
    },
    {
      codigo: '012',
      categoria: 'MATERIAL DE LIMPEZA',
      produto: 'DETERGENTE NEUTRO 500 ML',
      pregao: 'PREGÃO 58/2025 – DE 05/09/2025 – FORNECEDOR EXEMPLO E LTDA – 04/09/2027',
      ativo: 'Sim'
    },
    {
      codigo: '013',
      categoria: 'MATERIAL DE LIMPEZA',
      produto: 'PAPEL HIGIÊNICO ROLO 30 M FARDO 64',
      pregao: 'PREGÃO 58/2025 – DE 05/09/2025 – FORNECEDOR EXEMPLO E LTDA – 04/09/2027',
      ativo: 'Sim'
    },
    {
      codigo: '014',
      categoria: 'MATERIAL DESCARTÁVEL',
      produto: 'COPO DESCARTÁVEL 200 ML PACOTE 100 UN',
      pregao: 'PREGÃO 58/2025 – DE 05/09/2025 – FORNECEDOR EXEMPLO E LTDA – 04/09/2027',
      ativo: 'Sim'
    },
    {
      codigo: '015',
      categoria: 'INFORMÁTICA',
      produto: 'TONER COMPATÍVEL IMPRESSORA LASER PRETO',
      pregao: 'PREGÃO 77/2025 – DE 14/10/2025 – FORNECEDOR EXEMPLO F LTDA – 13/10/2027',
      ativo: 'Sim'
    },
    {
      codigo: '016',
      categoria: 'INFORMÁTICA',
      produto: 'MOUSE ÓPTICO USB',
      pregao: 'PREGÃO 77/2025 – DE 14/10/2025 – FORNECEDOR EXEMPLO F LTDA – 13/10/2027',
      ativo: 'Sim'
    },
    {
      codigo: '017',
      categoria: 'ELETRODOMÉSTICOS',
      produto: 'BEBEDOURO INDUSTRIAL INOX 25 LITROS',
      pregao: 'PREGÃO 05/2025 – DE 30/09/2025 – FORNECEDOR EXEMPLO G LTDA – 29/09/2026',
      ativo: 'Sim'
    },
    {
      codigo: '018',
      categoria: 'ELETRODOMÉSTICOS',
      produto: 'FREEZER HORIZONTAL 2 PORTAS 534 LITROS',
      pregao: 'PREGÃO 05/2025 – DE 30/09/2025 – FORNECEDOR EXEMPLO G LTDA – 29/09/2026',
      ativo: 'Sim'
    },
    {
      codigo: '019',
      categoria: 'ELETRODOMÉSTICOS',
      produto: 'FOGÃO INDUSTRIAL 4 BOCAS',
      pregao: 'PREGÃO 156/2025 – DE 29/01/2026 – FORNECEDOR EXEMPLO H LTDA – 29/01/2027',
      ativo: 'Sim'
    },
    {
      codigo: '020',
      categoria: 'ELETRODOMÉSTICOS',
      produto: 'FORNO MICRO-ONDAS 32 LITROS 110 OU 220',
      pregao: 'PREGÃO 156/2025 – DE 29/01/2026 – FORNECEDOR EXEMPLO H LTDA – 29/01/2027',
      ativo: 'Sim'
    },
    {
      codigo: '021',
      categoria: 'ARES CONDICIONADOS',
      produto: 'AR SPLIT 12.000 BTUS + INSTALAÇÃO',
      pregao: 'PREGÃO 14/2025 – DE 20/03/2025 – FORNECEDOR EXEMPLO I LTDA – 19/03/2027',
      ativo: 'Sim'
    },
    {
      codigo: '022',
      categoria: 'ARES CONDICIONADOS',
      produto: 'AR SPLIT 18.000 BTUS + INSTALAÇÃO',
      pregao: 'PREGÃO 14/2025 – DE 20/03/2025 – FORNECEDOR EXEMPLO I LTDA – 19/03/2027',
      ativo: 'Sim'
    },
    {
      codigo: '023',
      categoria: 'FERRAMENTAS / JARDINAGEM',
      produto: 'MANGUEIRA DE JARDIM 30 METROS',
      pregao: 'PREGÃO 90/2025 – DE 11/11/2025 – FORNECEDOR EXEMPLO J ME – 10/11/2027',
      ativo: 'Sim'
    },
    {
      codigo: '024',
      categoria: 'COMPRA DIRETA',
      produto: 'BOTA DE PVC CANO LONGO',
      pregao: 'COMPRA DIRETA',
      ativo: 'Sim'
    }
  ]
};
