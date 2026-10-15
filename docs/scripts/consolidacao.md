# Módulos de consolidação

Visões derivadas de `Consolidados`: a matriz de pedidos e a situação da rodada.

---

## `Resumo.gs`

**Caminho:** `src/consolidacao/Resumo.gs`
**Funções:** 46 — o módulo mais extenso do sistema

### Objetivo

Construir e formatar a aba `Resumo`: uma matriz produto × unidade, agrupada por pregão
e por categoria.

É a aba que o administrador usa no dia a dia para enxergar a rodada inteira.

### Layout produzido

```
ITEM                              │ UNIDADE A │ UNIDADE B │ UNIDADE C │ …
──────────────────────────────────┼───────────┼───────────┼───────────┼───
▸ PREGÃO 10/2025 — FORNECEDOR A                                           ← faixa mesclada
  ▸ PAPELARIA                                                             ← faixa mesclada
    PAPEL SULFITE BRANCO A4 RESMA │     10    │     25    │      –    │
    CANETA ESFEROGRÁFICA AZUL     │      5    │      –    │     12    │
▸ PREGÃO 35/2025 — FORNECEDOR C
  ▸ GÊNERO ALIMENTÍCIO
    ARROZ TIPO 1 PACOTE 5 KG      │     40    │     30    │     20    │
```

Uma coluna por unidade, na ordem de `DADOS_INICIAIS.unidades`.

### Construção

#### `escreverResumoAgrupadoFromConsolidados_(ssMae, linhasCons, opcoes)`

Ponto de entrada da construção.

| Parâmetro | Tipo | Descrição |
|---|---|---|
| `ssMae` | `Spreadsheet` | Planilha central |
| `linhasCons` | `Array` | Linhas de `Consolidados` |
| `opcoes` | `Object` | `modoLeve` pula a formatação pesada |

| Aspecto | Descrição |
|---|---|
| Entradas | `Consolidados`, `Produtos`, `Unidades` |
| Saída | Aba `Resumo` reconstruída |

> `modoLeve` é usado no caminho automático de sincronização. Reaplicar toda a
> formatação a cada edição numa unidade seria lento demais.

#### `montarDadosResumoPorPregao_(linhasCons, ssMae)`

Transforma as linhas de `Consolidados` na estrutura hierárquica
**pregão → categoria → produto → quantidades por unidade**.

#### `montarMapaCatalogoProdutoPregao_(ssMae)`

Constrói o mapa produto → pregão a partir da aba `Produtos`. É o que permite agrupar
por pregão, já que `Consolidados` guarda apenas o nome do produto.

#### `agruparPedidosPorUnidade_(linhasCons)`

Reorganiza os pedidos por unidade, para montar as colunas.

#### `escreverResumoPorPregao_(ssMae, dados, opcoes)`

Escreve os valores na aba.

#### `mesclarFaixasResumoPorPregao_(aba, nCols, meta)`

Mescla as células das faixas de pregão e categoria.

### Ordenação

#### `compararGenerosResumo_(nomeA, nomeB)` / `prioridadeOrdenacaoGeneroResumo_(nomeCat)`

Definem a ordem das categorias dentro de cada pregão.

#### `converterPregoesParaCategoriasResumo_(pregoes)` / `rotuloPregaoResumo_(nomeCompleto)`

Montam o rótulo curto exibido na faixa, a partir do texto completo do pregão.

A ordenação numérica dos pregões usa `compararPregoesOrdem_`, definida em
`RelatorioPregoes.gs` e compartilhada entre resumo e relatórios — o que garante a
mesma ordem nos dois lugares.

### Formatação

| Função | Finalidade |
|---|---|
| `aplicarVisualResumoPorPregaoRapido_` | Versão leve, usada no caminho automático |
| `aplicarVisualResumoPorPregao_` | Versão completa |
| `aplicarQuebraNomeUnidadesResumo_` | Quebra de linha nos nomes das unidades |
| `aplicarQuebraNomePregaoResumo_` | Quebra de linha nas faixas de pregão |
| `aplicarAlturasUniformesResumo_` | Altura uniforme das linhas |
| `reforcarCabecalhoResumo_` | Reaplica o estilo do cabeçalho |
| `desenharPainelPendenciasUnidades_` | Painel de pendências no topo |
| `atualizarIndicadorUltimaAtualizacao_` | Carimbo de data/hora |

#### `menuAtualizarVisualResumo()`

Reaplica toda a formatação sem recalcular os dados.

### Situação das unidades

#### `calcularStatusEnvioUnidades_(ssMae)`

| Retorno | `{enviaram: [], pendentes: []}` |
|---|---|

Determina quais unidades já enviaram. Alimenta o painel de pendências.

#### `formatarLinhasPendentesResumo_(nomes)`

Formata a lista de pendentes para exibição.

### Manutenção

| Função | Finalidade |
|---|---|
| `garantirAbaResumoModeloPregao_` | Cria a aba no modelo atual |
| `obterOuCriarAbaResumoPregoes_` | Obtém ou cria a aba |
| `aposentarAbaResumoPesada_` | Arquiva uma aba antiga com formatação pesada |
| `recriarAbaResumo_` | Recria do zero |
| `desfazerTodosMergesAba_` | Remove todas as mesclagens antes de reescrever |
| `prepararResumoNovoModeloZerado_` | Prepara a aba zerada no modelo novo |
| `reconstruirResumoDeConsolidadosExistentes_` | Reconstrói sem consultar as unidades |
| `limparColunaLegadoResumo_` | Remove a coluna legada de estoque |
| `normalizarLinhaCabecalhoResumo_` | Corrige o cabeçalho |

> `reconstruirResumoDeConsolidadosExistentes_` é a prova prática de que `Consolidados`
> é a fonte de verdade: o `Resumo` inteiro pode ser refeito sem tocar nas unidades.

### Normalização de texto

`padronizarNomenclaturaProdutoResumo_` e `limparNomeProdutoResumo_` padronizam os
nomes exibidos, removendo o texto do pregão quando ele vem grudado no nome do produto.

### Abas relacionadas

`Consolidados`, `Produtos`, `Unidades` (leitura); `Resumo` (escrita)

### Dependências

`Config.gs`, `Util.gs`, `Produtos.gs`, `Unidades.gs`, `RelatorioPregoes.gs`
(para `compararPregoesOrdem_`)

### Execução

Automática após cada sincronização. Manual por `menuAtualizarVisualResumo` e
`menuPrepararResumoNovoModelo`.

### Limitações

- A aba é **reconstruída por completo** a cada atualização: edições manuais no
  `Resumo` se perdem.
- A formatação completa é cara. Em catálogos grandes com muitas unidades, pode se
  aproximar do limite de execução — daí a existência do `modoLeve`.
- As mesclagens exigem desfazer todas as anteriores antes de reescrever, o que torna
  a operação sensível a interrupções no meio.
- O agrupamento por pregão depende de o produto estar na aba `Produtos`. Um item em
  `Consolidados` sem correspondência no catálogo cai em um bloco genérico.

---

## `TotalPorProduto.gs`

**Caminho:** `src/consolidacao/TotalPorProduto.gs`
**Funções:** 18

### Objetivo

Manter a aba `Pedidos das Unidades`: quem já pediu e quem ainda não pediu.

O nome do arquivo é histórico — a aba antes mostrava totais por produto e foi
reaproveitada para a situação da rodada.

### Layout produzido

| Já pediram | Ainda não pediram |
|---|---|
| UNIDADE CENTRO | UNIDADE LESTE |
| UNIDADE NORTE | UNIDADE OESTE |
| UNIDADE SUL | … |

### Funções

#### `obterOuCriarAbaPedidosDasUnidades_(ssMae)`

Obtém a aba. Se encontrar a aba antiga `Total por Produto`, **renomeia** em vez de
criar outra — preserva o histórico de instalações antigas.

#### `atualizarAbaSituacaoPedidosMae_(ssMae)`

Atualiza a aba. Chamada pelo `onOpen` e após cada sincronização.

#### `montarSituacaoPedidosUnidades_(linhasCons, ssMae, opcoes)`

| Retorno | `{concluidas: [], pendentes: []}` |
|---|---|

Determina a situação de cada unidade.

#### `coletarChavesUnidadesQueJaPediram_(linhasCons, ssMae, varrerFilhas)`

Núcleo da regra. Uma unidade conta como "já pediu" se **qualquer** destes for
verdadeiro:

1. há registro dela em `Consolidados`;
2. há quantidade lançada na coluna dela no `Resumo`
   (`marcarUnidadesComQuantidadeNoResumo_`);
3. há itens marcados como `Enviado` na própria planilha dela
   (`marcarUnidadesComPedidoEnviadoNasFilhas_`).

A verificação tripla evita falso negativo quando o pedido chegou por um caminho
diferente do esperado.

> O parâmetro `varrerFilhas` controla a terceira verificação, que é a mais cara: abre
> cada planilha de unidade. No caminho automático ela é desligada.

#### `ordenarNomesUnidadesSituacao_(nomes)`

Ordena pela ordem oficial de `DADOS_INICIAIS.unidades`. Unidades fora dessa lista vão
para o fim, em ordem alfabética.

#### `agendarAtualizacaoTotalPorProduto_()` / `processarTotalPorProdutoPendente_()`

Adiam a atualização em ~8 s, para não concorrer com a reconstrução do `Resumo` na
mesma execução.

#### `limparTotalPorProdutoNovaRodada_(ssMae)`

Zera a aba: todas as unidades voltam para pendente. Chamada pela nova rodada.

#### `aplicarVisualSituacaoPedidos_(aba, nCols, nLinhasCorpo, situacao, modoLeve)`

Formatação, com versão leve para o caminho automático.

#### `chaveUnidadeSituacao_(nome)`

Normaliza o nome da unidade para comparação.

### Abas relacionadas

`Consolidados`, `Resumo`, `Unidades` (leitura); `Pedidos das Unidades` (escrita)
Opcionalmente, a aba `Pedido` de cada unidade

### Dependências

`Config.gs`, `Util.gs`, `Unidades.gs`, `DadosIniciais.gs`

### Execução

Automática ao abrir a central e após cada sincronização. Manual por
`menuAbrirTotalPorProduto`.

### Limitações

- Com `varrerFilhas` ativo, a função abre todas as planilhas das unidades: é a parte
  mais lenta da atualização.
- O vínculo é pelo **nome** da unidade, normalizado. Nomes divergentes entre a aba
  `Unidades` e os dados consolidados quebram a correspondência.
- A aba mostra apenas se a unidade pediu **alguma coisa**, não o que pediu nem se o
  pedido está completo.
