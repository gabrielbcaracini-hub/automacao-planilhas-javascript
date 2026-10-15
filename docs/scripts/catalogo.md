# Módulos de catálogo

Gestão dos produtos: cadastro, categorias, validações e importação de pregões.

---

## `Produtos.gs`

**Caminho:** `src/catalogo/Produtos.gs`
**Funções:** 21

### Objetivo

Leitura, validação, organização e camada visual da aba `Produtos` — o catálogo mestre
do sistema.

### Funções principais

#### `obterMapaColunasProdutos_(aba)`

| Aspecto | Descrição |
|---|---|
| Parâmetros | `aba` — `Sheet` da aba `Produtos` |
| Retorno | Objeto com índices **0-based**: `colCodigo`, `colCategoria`, `colProduto`, `colUnidade`, `colPregao`, `colAtivo` |

Localiza as colunas pelo **texto do cabeçalho**, normalizado por
`normalizarHeaderProduto_`. Tolera cabeçalhos com `#` extra, acentos diferentes ou
colunas fora da ordem padrão.

É a função que permite ao sistema continuar lendo planilhas com layout antigo.

> Os índices são 0-based. Para montar um `getRange`, some 1: `map.colProduto + 1`.

#### `menuInserirLinhaNovoProduto()`

Ponto de entrada do menu **➕ Nova linha na aba Produtos**.

| Aspecto | Descrição |
|---|---|
| Parâmetros | Nenhum |
| Retorno | Nenhum |
| Entradas | Linha selecionada na aba `Produtos`; número do pregão informado pelo usuário |
| Saída | Uma linha em branco, formatada e posicionada |

Fluxo:

1. `obterContextoInsercaoProduto_` descobre a categoria e o pregão da linha selecionada.
2. Se não houver categoria, pergunta ao usuário.
3. Pergunta o **número do pregão** a cadastrar, sugerindo o número da linha selecionada.
4. `encontrarLinhaParaNovoPregao_` calcula onde inserir.
5. Insere a linha, preenche categoria e `Ativo = Sim`.
6. Aplica validação e formatação e posiciona o cursor na coluna Produto.

#### `encontrarLinhaParaNovoPregao_(aba, map, categoria, pregaoRef, lastRow, nCols)`

| Aspecto | Descrição |
|---|---|
| Parâmetros | `aba`, `map` (de `obterMapaColunasProdutos_`), `categoria`, `pregaoRef` (texto com o número), `lastRow`, `nCols` |
| Retorno | `number` — a linha onde inserir |

Regra de posicionamento. Varre as linhas da categoria e guarda a última cujo número de
pregão seja **menor ou igual** ao número informado:

| Situação | Resultado |
|---|---|
| Existem itens do mesmo número | Logo após o último deles |
| O número é novo, mas há números menores | Após o último número menor, mantendo a ordem |
| O número é menor que todos | Logo após o título da categoria |
| Nenhum número informado | No fim da categoria |

Isso resolve o caso de uma categoria com pregões diferentes (por exemplo, `05/2025` e
`156/2025` convivendo em `ELETRODOMÉSTICOS`): um item novo do pregão 05 entra no bloco
do 05, não no topo da categoria.

#### `extrairNumeroPregaoCadastro_(texto)`

| Aspecto | Descrição |
|---|---|
| Parâmetros | `texto` — texto do pregão ou só o número |
| Retorno | `number` — o número, ou `-1` se não reconhecer |

Aceita tanto o texto completo (`'PREGÃO 156/2025 – DE …'`) quanto o número isolado
(`'156'`), o que permite usar a mesma função na leitura da planilha e na resposta do
usuário.

#### `organizarProdutosPorCategoria_(ssMae)`

Reordena a aba agrupando por categoria e insere as faixas de título. Quando a aba está
como tabela tipada do Sheets, delega para `recriarAbaProdutosSemTabela_`, que recria a
aba sem esse formato — tabelas tipadas impedem a inserção livre de linhas.

#### `listarTodosProdutos()` / `listarProdutosAtivos()` / `listarNomesProdutosAtivos()`

| Função | Retorno |
|---|---|
| `listarTodosProdutos` | Todos os itens, incluindo inativos |
| `listarProdutosAtivos` | Só `Ativo = Sim` |
| `listarNomesProdutosAtivos` | Só os nomes, para montar validações |

Todas ignoram as linhas de título de categoria.

#### `validarUnicidadeProdutos()`

Detecta nomes de produto duplicados, comparando por `normalizarTexto`.

#### `aplicarVisualAbaProdutosMae_` / `aplicarVisualLinhasProdutos_` / `aplicarFaixasCategoriaProdutos_`

Camada visual: cores de `CONFIG.TEMA`, alternância de linhas, destaque da coluna de
categoria e faixas de título.

`aplicarVisualLinhasProdutos_` aceita um intervalo de linhas, o que permite formatar
só a linha recém-inserida em vez da aba inteira.

### Abas relacionadas

`Produtos` (leitura e escrita)

### Dependências

`Config.gs`, `Util.gs`

### Execução

`menuInserirLinhaNovoProduto` pelo menu. `menuOrganizarProdutosPorCategoria` e
`menuAjustarVisualProdutos` pelo editor.

### Limitações

- O posicionamento depende do número do pregão estar no formato `número/ano`. Textos
  fora desse padrão caem no fim da categoria.
- `organizarProdutosPorCategoria_` reescreve a aba inteira; em catálogos muito grandes
  pode se aproximar do limite de execução.
- Linhas de título de categoria são identificadas por heurística
  (`ehLinhaTituloCategoriaProdutos_`): uma linha de dados com formato parecido pode ser
  interpretada como título.

---

## `Categorias.gs`

**Caminho:** `src/catalogo/Categorias.gs`
**Funções:** 5

### Objetivo

Controlar quais categorias estão liberadas para pedido no mês.

### Funções

#### `obterDisponibilidadeCategorias_()`

| Aspecto | Descrição |
|---|---|
| Retorno | Mapa `categoria normalizada → boolean` |
| Entrada | Aba `Categorias` |

**Sem a aba `Categorias`, retorna todas as categorias como liberadas.** Essa escolha
mantém compatibilidade com instalações anteriores ao recurso.

#### `categoriaLiberadaNoMes_(categoria, mapa)`

Consulta o mapa. Categoria ausente conta como liberada.

#### `sincronizarCadastroCategorias_()`

Acrescenta à aba `Categorias` as categorias que existem em `Produtos` e ainda não
foram cadastradas. Preserva o `Sim`/`Não` já marcado.

#### `menuAtualizarControleCategorias()`

Ponto de entrada. Chama a sincronização e informa quantas categorias foram
acrescentadas.

#### `completarResumoComProdutosDaAba_()`

Completa a aba `Resumo` com produtos do catálogo que ainda não apareceram em nenhum
pedido, para que o resumo mostre a lista completa do que está liberado.

### Abas relacionadas

`Categorias` (leitura e escrita), `Produtos` (leitura), `Resumo` (escrita)

### Dependências

`Config.gs`, `Util.gs`, `Produtos.gs`

### Limitações

- Categorias são comparadas por texto normalizado; renomear uma categoria em
  `Produtos` cria uma entrada nova em `Categorias`.
- Não há remoção automática de categorias que deixaram de existir.

---

## `Dropdown.gs`

**Caminho:** `src/catalogo/Dropdown.gs`
**Funções:** 10

### Objetivo

Montar a aba `Catalogo` nas planilhas das unidades e aplicar as validações de dados da
aba `Pedido`.

### Funções

#### `garantirAbaCatalogo(ssFilha)`

Cria a aba `Catalogo` com o cabeçalho `['Produto']`, se não existir.

#### `atualizarCatalogoFilha(ssFilha, nomesProdutos)`

| Parâmetros | Descrição |
|---|---|
| `ssFilha` | `Spreadsheet` da unidade |
| `nomesProdutos` | `Array.<string>` |

Reescreve a lista de produtos válidos.

#### `aplicarDropdownPedido(ssFilha, opcoes)`

Aplica as validações na aba `Pedido`. **Somente a coluna Status tem lista suspensa** —
é a marcação de `Enviado` que dispara a sincronização.

As colunas Produto e Observação não recebem validação, para não travar a digitação.

#### `aplicarValidacaoStatusPedido_(...)`

Monta a lista suspensa com os valores de `CONFIG.VALORES`:
`Rascunho`, `Enviado`, `Consolidado`.

#### `protegerCatalogo(ssFilha)`

Protege a aba `Catalogo` contra edição pelo operador da unidade. O catálogo é definido
pela central.

#### `sincronizarDropdownFilha(ssFilha, nomesProdutos, produtosCatalogo, opcoes)`

Agrupa as operações acima numa chamada só. É o que a sincronização de catálogo usa.

#### `corrigirValidacoesPedidoTodasUnidades_()`

Reaplica as validações em todas as unidades. Rotina de reparo, para o caso de
validações terem sido apagadas manualmente.

### Abas relacionadas

`Catalogo` e `Pedido` nas unidades

### Dependências

`Config.gs`, `Util.gs`, `Unidades.gs`

### Limitações

- O número de linhas validadas vem de `CONFIG.PEDIDO_LINHAS_VALIDACAO` (800 por
  padrão). Catálogos maiores que isso ficam sem validação nas linhas excedentes.
- A proteção da aba `Catalogo` não impede o dono da planilha de removê-la.

---

## `NovosPregoesMatriz.gs`

**Caminho:** `src/catalogo/NovosPregoesMatriz.gs`
**Funções:** 4
**Dados:** `NOVOS_PREGOES_MATRIZ_SETEMBRO_2026`

### Objetivo

Importar um lote de pregões novos para a aba `Produtos` sem sobrescrever o que já
existe.

> O array de dados publicado neste repositório é **fictício** (10 itens). No sistema
> original continha o lote real de pregões.

### Funções

#### `acrescentarLinhasNovosPregoesMatriz_(ss)`

| Aspecto | Descrição |
|---|---|
| Parâmetros | `ss` — `Spreadsheet` opcional; sem ele, usa a central |
| Retorno | `{adicionados, ignorados, pregoes, mensagem}` |
| Entrada | `NOVOS_PREGOES_MATRIZ_SETEMBRO_2026` e a aba `Produtos` atual |
| Saída | Linhas acrescentadas em `Produtos` |

Compara cada item com o que já existe, por nome normalizado. Itens já cadastrados são
ignorados. **Nada é sobrescrito.** Os códigos continuam a partir do maior código já
existente.

#### `importarMatrizSetembro2026SePendente_()`

Executa a importação **uma única vez**, controlada por uma flag em
`DocumentProperties`. Chamada pelo `onOpen`.

#### `importarNovosPregoesMatrizSetembro2026()`

Importação direta, sem verificar a flag.

#### `menuImportarNovosPregoesMatrizSetembro2026()`

Ponto de entrada, com relatório do que foi importado.

### Abas relacionadas

`Produtos` (leitura e escrita)

### Dependências

`Config.gs`, `Util.gs`, `Produtos.gs`

### Execução

Automático na abertura (uma vez) ou manual pelo editor.

### Limitações

- A deduplicação é por **nome de produto**, não por nome mais pregão. Um produto com
  o mesmo nome em um pregão diferente é considerado duplicado e ignorado.
- O nome da função é específico de um lote. Para importar outro lote, substitua o
  array e remova a flag em `DocumentProperties`.
- Lotes grandes podem exceder o limite de execução; a importação não tem continuação
  automática.
