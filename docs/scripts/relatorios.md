# Módulos de relatórios

Geração de arquivos em Excel, PDF e Word a partir dos dados consolidados.

Todos os relatórios são criados no Drive da conta que executa e a função devolve o
link do arquivo gerado.

---

## `ExcelRelatorio.gs`

**Caminho:** `src/relatorios/ExcelRelatorio.gs`
**Funções:** 15
**Formato:** `.xlsx`

### Objetivo

Gerar o relatório dos pedidos do mês, agrupado por pregão.

Hierarquia: **pregão → produto → unidades e quantidades**. Inclui apenas itens com
quantidade pedida.

### Como o `.xlsx` é produzido

O Apps Script não escreve `.xlsx` diretamente. A estratégia:

1. criar uma planilha temporária no Drive;
2. montar e formatar o conteúdo nela;
3. baixar essa planilha pelo endpoint de exportação do Google, via `UrlFetchApp`;
4. salvar o binário como arquivo `.xlsx`;
5. remover a planilha temporária.

### Funções

#### `menuGerarRelatorioExcel()`

Ponto de entrada do menu **📥 Baixar relatório Excel**. Exibe o link ao final.

#### `obterLinhasConsParaRelatorioExcel_(ss)`

Lê as linhas de `Consolidados`. Se a aba estiver vazia, cai para
`extrairLinhasConsDaAbaResumo_`, que reconstrói os dados a partir do `Resumo`.

#### `agruparDadosRelatorioPorPregao_(linhasCons, ssOrigem)`

| Retorno | Estrutura hierárquica pregão → produto → unidades |
|---|---|

Usa `chaveAgrupamentoPregaoRelatorio_` para agrupar e
`enriquecerCatalogoChavesFlexRelatorio_` para casar produtos cujo nome varia
ligeiramente entre as planilhas.

#### `montarLinhasRelatorioExcel_(dados, agora, tz)`

Converte a estrutura em linhas de planilha, com data e fuso.

#### `exportarPlanilhaComoXlsx_(spreadsheetId, nomeArquivo)`

| Parâmetros | Descrição |
|---|---|
| `spreadsheetId` | ID da planilha a exportar |
| `nomeArquivo` | Nome do arquivo final |

| Retorno | Arquivo `.xlsx` no Drive |
|---|---|

Requer o escopo `script.external_request`.

#### `aplicarVisualRelatorioExcel_(aba, meta, nRows, nCols)`

Formatação: cores, larguras, mesclagens e bordas.

#### `obterPastaRelatoriosExportacao_(ss)`

Obtém ou cria a pasta de destino no Drive, mantendo os relatórios organizados.

#### `pedidoTemQuantidadeRelatorio_(valor)`

Filtra itens sem quantidade.

### Abas relacionadas

`Consolidados` (leitura), `Resumo` (fallback), `Produtos` (agrupamento)

### Dependências

`Config.gs`, `Util.gs`, `Produtos.gs`, `RelatorioPregoes.gs` (ordenação)

### Limitações

- Consome cota de `UrlFetchApp` e de criação de arquivos.
- Relatórios muito grandes podem estourar o limite de execução.
- Os arquivos ficam no Drive da conta que executou e consomem a cota dela.
- Se a planilha temporária não for removida por causa de uma falha, resta lixo no
  Drive.

---

## `RelatorioPregoes.gs`

**Caminho:** `src/relatorios/RelatorioPregoes.gs`
**Funções:** 30
**Formatos:** `.xlsx` e `.docx`

### Objetivo

Listar cada pregão e os produtos que ele disponibiliza, em duas variantes: **ativos** e
**inativos**.

Este módulo também concentra as **funções de ordenação de pregões** usadas pelo resto
do sistema.

### Ordenação de pregões

Conjunto de funções compartilhado com `Resumo.gs` e `ExcelRelatorio.gs`. É o que
garante a mesma ordem no cadastro, no resumo e nos relatórios.

#### `parsearPartesPregao_(texto)`

| Parâmetros | `texto` — texto completo do pregão |
|---|---|
| Retorno | `{numero, ano, data, fornecedor, validade}` |

Interpreta o formato convencionado:

```
PREGÃO 156/2025 – DE 29/01/2026 – FORNECEDOR EXEMPLO H LTDA – 29/01/2027
       └──┬──┘      └────┬────┘    └────────┬────────────┘   └────┬────┘
        número        vigência            fornecedor           validade
```

#### `extrairChaveOrdemPregao_(texto)`

| Retorno | Chave numérica composta por ano, número e data |
|---|---|

#### `compararPregoesOrdem_(textoA, textoB)`

Comparador para `Array.sort`. Ordena por **número, ano e data de vigência**, todos
numericamente.

É por isso que `14/2025` vem antes de `101/2025` — a ordenação alfabética de texto
faria o contrário, porque compararia `'1'` com `'1'`, depois `'4'` com `'0'`.

#### `normalizarChavePregaoRelatorio_(texto)` / `normalizarTextoPregaoRelatorio_(texto)`

Normalizam o texto para agrupar variações de digitação do mesmo pregão.

### Geração

#### `menuGerarRelatorioPregoes()` / `menuGerarRelatorioPregoesInativos()`

Pontos de entrada do menu. Ambos chamam `gerarRelatorioPregoesPeloMenu_(modo)`.

| Modo | Filtro |
|---|---|
| `ativos` | Produtos com `Ativo = Sim` |
| `inativos` | Produtos com `Ativo = Não` |

#### `montarDadosRelatorioPregoes_(produtos, modo)`

Agrupa os produtos por pregão, já ordenados.

#### `gerarRelatorioPregoesExcel_(ss, opcoes)` / `gerarRelatorioPregoesWord_(ss, opcoes)`

Produzem as duas saídas a partir da mesma estrutura de dados.

#### `agruparRelatorioPregoesPorGenero_(dados)`

Subagrupa por categoria dentro de cada pregão.

#### `preencherAbaRelatorioPregoesLegivel_(aba, dados, agora, tz, modo)`

Monta a aba principal do Excel.

> `preencherAbaListaPregoesExcel_` tem corpo vazio: é uma aba que foi descontinuada,
> mas a função foi mantida para não quebrar as chamadas existentes.

#### `montarCapaListaPregoes_` / `montarSecaoPregaoMatriz_` / `montarTabelaItensPregao_`

Montagem do documento Word: capa, seção por pregão e tabela de itens.

#### `montarRotuloItemPregao_(produto)` / `limparNomeProdutoRelatorio_(texto, pregaoTexto)` / `formatarNomeProdutoRelatorio_(texto)`

Padronizam o nome exibido, removendo o texto do pregão quando ele vem grudado no nome
do produto.

#### `montarMapaQuantidadesPorProdutoUnidade_(linhasCons)`

Mapa produto × unidade com as quantidades, quando o relatório inclui pedidos.

#### `obterPaletaRelatorioPregoes_()`

Paleta de cores do relatório.

#### `obterPastaRelatoriosPregoes_(ss)`

Pasta de destino no Drive.

### Abas relacionadas

`Produtos` (leitura), `Consolidados` (quantidades, quando aplicável)

### Dependências

`Config.gs`, `Util.gs`, `Produtos.gs`

### Limitações

- A interpretação do pregão depende do formato convencionado. Textos fora do padrão
  entram num grupo genérico e ficam no fim da ordenação.
- Documentos Word longos podem se aproximar do limite de execução.
- O relatório reflete o estado atual da aba `Produtos`: itens cadastrados depois não
  aparecem até a próxima geração.

---

## `RelatorioPersonalizado.gs`

**Caminho:** `src/relatorios/RelatorioPersonalizado.gs`
**Funções:** 6
**Formato:** `.docx`

### Objetivo

Relatório administrativo com visão por unidade, a partir de `Consolidados`.

### Funções

#### `menuGerarRelatorioPersonalizado()`

Ponto de entrada.

#### `gerarRelatorioPersonalizado_(ss, linhasCons)`

| Parâmetros | Descrição |
|---|---|
| `ss` | `Spreadsheet` central |
| `linhasCons` | Linhas de `Consolidados` |

Monta o documento com cabeçalho institucional, tabela de visão geral e tabela de
solicitações por unidade.

O cabeçalho usa `obterNomeOrganizacao_()`, que lê `CONFIG.ORGANIZACAO.NOME`.

#### `adicionarLogoPrefeituraNoRelatorio_(body)`

Baixa a imagem de `CONFIG.ORGANIZACAO.LOGO_URL` e insere no documento.

Com `LOGO_URL` vazio, o relatório é gerado sem imagem — não há erro.

#### `agruparDadosRelatorioPorUnidade_(linhasCons, ssMae)`

Agrupa os pedidos por unidade.

#### `formatarTabelaVisaoGeral_(tabela, azul)` / `formatarTabelaSolicitacoes_(tabela, azulClaro)`

Formatação das tabelas do documento.

### Abas relacionadas

`Consolidados`, `Unidades` (leitura)

### Dependências

`Config.gs`, `Util.gs`, `Unidades.gs`

### Limitações

- O logotipo precisa estar em URL pública, acessível sem autenticação.
- Requer o escopo `script.external_request` para baixar a imagem.
- O documento é gerado sempre do zero; não há atualização incremental.

---

## `PdfResumo.gs`

**Caminho:** `src/relatorios/PdfResumo.gs`
**Funções:** 3
**Formato:** `.pdf`

### Objetivo

Exportar a aba `Resumo` como PDF, para arquivamento e envio.

### Funções

#### `menuGerarPdfResumo()`

Ponto de entrada.

#### `gerarBlobPdfResumo_(ss, abaResumo)`

| Retorno | `Blob` do PDF |
|---|---|

Usa o endpoint de exportação do Google com os parâmetros de página, margem e
orientação.

#### `obterPastaPdfsResumo_(ss)`

Pasta de destino no Drive.

### Abas relacionadas

`Resumo` (leitura)

### Dependências

`Config.gs`, `Util.gs`

### Limitações

- O PDF reproduz a aba como está: se o `Resumo` estiver desatualizado, o PDF também
  estará. Puxe os pedidos antes de gerar.
- Resumos muito largos (muitas unidades) podem ficar ilegíveis no PDF.

---

## `GuiaSistemaPdf.gs`

**Caminho:** `src/relatorios/GuiaSistemaPdf.gs`
**Funções:** 4
**Formato:** `.pdf`

### Objetivo

Gerar um guia do sistema em PDF, para apresentação e treinamento.

Diferente dos demais, este relatório **não consome dados**: o conteúdo é o texto
explicativo do próprio sistema.

### Funções

#### `menuGerarPdfGuiaSistema()`

Ponto de entrada.

#### `gerarPdfGuiaSistema_(ss)`

Monta o documento: capa com o nome da organização, seções de uso e tabelas de
referência. Converte para PDF e salva no Drive.

#### `adicionarTituloGuia_(body, titulo)` / `adicionarTabelaGuia_(body, valores, azul)`

Auxiliares de montagem.

### Dependências

`Config.gs`, `Util.gs`

### Limitações

- O conteúdo do guia está **escrito no código**. Mudanças no sistema exigem atualizar
  o texto manualmente em `gerarPdfGuiaSistema_`.
- O guia descreve o fluxo padrão; adaptações locais não são refletidas.
