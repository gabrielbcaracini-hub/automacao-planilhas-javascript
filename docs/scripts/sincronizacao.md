# Módulos de sincronização

Ligação entre a planilha central e as planilhas das unidades, nos dois sentidos.

---

## `Sincronizacao.gs`

**Caminho:** `src/sincronizacao/Sincronizacao.gs`
**Funções:** 12
**Sentido:** central → unidades

### Objetivo

Distribuir o catálogo de produtos da central para todas as planilhas das unidades.

A propagação é **sempre manual**. Não há gatilho que distribua o catálogo
automaticamente: uma edição acidental na aba `Produtos` não pode reescrever as
planilhas de todas as unidades no meio de uma rodada.

### Funções

#### `sincronizarCatalogoTodasUnidades(silencioso, semLock, opcoes)`

Função central do módulo.

| Parâmetro | Tipo | Descrição |
|---|---|---|
| `silencioso` | `boolean` | Suprime os alertas na interface |
| `semLock` | `boolean` | Pula o `LockService` (quando o chamador já detém o lock) |
| `opcoes` | `Object` | Ajustes de comportamento, como pular a camada visual |

| Aspecto | Descrição |
|---|---|
| Retorno | Objeto com unidades processadas, falhas e contagens |
| Entradas | Abas `Produtos`, `Categorias` e `Unidades` |
| Saídas | Abas `Catalogo` e `Pedido` de cada unidade |

Para cada unidade ativa:

1. lê os produtos ativos das categorias liberadas;
2. reescreve a aba `Catalogo`;
3. monta a aba `Pedido` com os produtos agrupados por categoria;
4. aplica as validações de dados;
5. grava o mês de referência;
6. aplica a formatação visual.

**Quantidades já preenchidas são preservadas** quando o produto continua no catálogo.

Cada unidade é processada dentro do próprio `try/catch`: uma unidade inacessível não
interrompe as demais.

#### `menuAtualizarProdutosUnidadesAgora()`

Ponto de entrada do menu **⚡ Atualizar produtos nas unidades agora**. Exibe o
resultado consolidado ao final.

#### `removerPedidosProdutosInativosNaMae_()`

Remove de `Consolidados` e `Resumo` os pedidos de produtos que passaram a `Ativo = Não`.

#### `diagnosticarCatalogoAtivo_()`

Relatório do que seria distribuído: quantos produtos ativos, em quais categorias
liberadas. Útil para conferir antes de distribuir.

#### `removerGatilhosCatalogoAutomatico_()`

Remove gatilhos de distribuição automática criados por versões antigas. Chamada pelo
`onOpen`.

#### Funções vazias preservadas

`onEditProdutosInstalavel`, `processarCatalogoPendente`, `processarCatalogoAtrasado_`
e `executarCatalogoAtrasado_` têm corpo vazio.

Elas existem porque gatilhos instalados por versões anteriores ainda podem apontar
para esses nomes. Remover as funções faria esses gatilhos lançarem erro a cada
disparo. Mantê-las vazias é o que desativa a automação com segurança.

### Abas relacionadas

Central: `Produtos`, `Categorias`, `Unidades` (leitura)
Unidades: `Catalogo`, `Pedido` (escrita)

### Dependências

`Config.gs`, `Util.gs`, `Produtos.gs`, `Categorias.gs`, `Unidades.gs`, `Dropdown.gs`,
`VisualUnidades.gs`, `MesReferencia.gs`

### Limitações

- Percorre todas as unidades em uma execução; com muitas unidades, pode se aproximar
  do limite de 6 minutos.
- Uma unidade sem permissão de edição falha e é reportada, mas não bloqueia as outras.
- Produtos removidos do catálogo somem da aba `Pedido` da unidade, junto com as
  quantidades que tivessem sido lançadas neles.

---

## `Bridge.gs`

**Caminho:** `src/sincronizacao/Bridge.gs`
**Funções:** 48
**Sentido:** unidades → central

### Objetivo

Caminho de volta: detectar edições nas unidades, enfileirar, processar e escrever os
pedidos na central. É o módulo que sustenta a sincronização automática.

### Gatilho e detecção

#### `onEditPedidoUnidade(e)`

| Aspecto | Descrição |
|---|---|
| Parâmetros | `e` — evento de edição do Apps Script |
| Tipo | Gatilho **instalável**, registrado por planilha de unidade |

Verifica se a edição é relevante antes de fazer qualquer trabalho:

1. `ehAbaPedido_` — a edição foi na aba `Pedido`?
2. `rangeIncluiColunasPedido_` — tocou a coluna Status ou Quantidade?

Só então enfileira a unidade. Edições em outras abas ou colunas são descartadas de
imediato.

### Fila e debounce

#### `agendarSyncPedidoUnidadeDebounced_(spreadsheetId)`

Grava o ID da unidade na fila em `PropertiesService` e cria um gatilho temporizado
para ~7 s depois — **se ainda não houver um pendente**.

É o que transforma 30 edições em sequência numa única sincronização.

#### `lerFilaSyncPendente_(props)` / `salvarFilaSyncPendente_(props, ids)` / `limparPropsSyncPendente_()`

Manipulação da fila em `PropertiesService`.

#### `processarSyncPedidoPendenteRapido_()`

Executada pelo gatilho temporizado. Lê a fila e sincroniza cada unidade pendente.

#### `criarGatilhoSyncAtrasadoUnidade_(spreadsheetId)` / `executarSyncAtrasadoUnidade_()`

Nova tentativa para uma unidade que falhou, agendada alguns segundos depois.

### Escrita na central

#### `sincronizarUnidadePorId(spreadsheetId)`

| Aspecto | Descrição |
|---|---|
| Parâmetros | `spreadsheetId` — ID da planilha da unidade |
| Retorno | Resultado da sincronização |

Abre a unidade, lê os itens marcados como `Enviado` e atualiza a central.

#### `atualizarMaeComPedidosUnidade(nomeUnidade, pedidosEnviados)`

| Parâmetro | Tipo |
|---|---|
| `nomeUnidade` | `string` |
| `pedidosEnviados` | `Array.<{produto, quantidade, observacao}>` |

Substitui os registros daquela unidade em `Consolidados` e atualiza `Resumo` e
`Pedidos das Unidades`. Os dados das outras unidades não são tocados.

#### `chaveProdutoFlex_(s)` / `normalizarChaveUnidade_(nome)`

Geram as chaves de comparação usadas para casar produto e unidade entre as planilhas,
tolerando diferenças de acento, espaço e caixa.

### Menu e diagnóstico

| Função | Finalidade |
|---|---|
| `menuConectarTudoAgora` | Instala os gatilhos em todas as unidades e faz a primeira puxada |
| `menuAtualizarResumoPedidosAgora` | **📥 Puxar pedidos das unidades** |
| `menuReconectarPlanilhas` | Recria os gatilhos quando a conexão cai |
| `menuDiagnosticarSyncAutomatico` | Lista os gatilhos e aponta unidades sem gatilho |
| `menuInstalarSyncAutomatico` | Instala apenas os gatilhos |
| `menuTestarUnidadeCentro` | Teste pontual com uma unidade |

### Gestão de gatilhos

| Função | Finalidade |
|---|---|
| `instalarGatilhosUnidades` | Instala `onEdit` em cada unidade |
| `recriarGatilhosSyncUnidades_` | Remove e reinstala |
| `removerGatilhosOnEditUnidadesDuplicados_` | Limpa gatilhos duplicados |
| `removerGatilhosClockPorFuncao_` | Remove gatilhos temporizados órfãos |
| `garantirTriggerConsolidacaoFrequente` | Garante o gatilho diário |

### Escrita formatada

`escreverAbaSimples_`, `escreverSomenteValoresSemFormatar_`, `escreverAbaComModelo_`,
`gravarValoresPreservandoAba_` e `aplicarModeloVisualMae_` são variações de escrita com
diferentes níveis de formatação.

A versão "somente valores" existe para o caminho automático: reformatar a aba a cada
sincronização seria lento demais.

### Abas relacionadas

Unidades: `Pedido` (leitura)
Central: `Consolidados`, `Resumo`, `Pedidos das Unidades` (escrita)

### Dependências

`Config.gs`, `Util.gs`, `Unidades.gs`, `RepararUnidade.gs`, `Resumo.gs`,
`TotalPorProduto.gs`

### Limitações

- Gatilhos instaláveis rodam com as credenciais de **quem os instalou**. Se essa conta
  perder acesso a uma unidade, a sincronização daquela unidade para.
- Limite de 20 gatilhos por script por usuário: muitas unidades podem esbarrar nesse
  teto.
- A fila fica em `PropertiesService`, que tem limite de tamanho por propriedade.
- Edições simultâneas em muitas unidades podem gerar execuções concorrentes; o lock
  mitiga, mas não elimina.

---

## `Consolidacao.gs`

**Caminho:** `src/sincronizacao/Consolidacao.gs`
**Funções:** 6

### Objetivo

Varredura completa: percorre todas as unidades e remonta `Consolidados` do zero.

Diferente do `Bridge.gs`, que atualiza **uma** unidade por vez de forma incremental,
aqui a varredura é total.

### Funções

#### `consolidarPedidos(marcarComoConsolidado, semLock)`

| Parâmetro | Tipo | Descrição |
|---|---|---|
| `marcarComoConsolidado` | `boolean` | Muda o status dos itens lidos para `Consolidado` |
| `semLock` | `boolean` | Pula o `LockService` |

| Aspecto | Descrição |
|---|---|
| Retorno | `{linhas, unidades, falhas}` |
| Entradas | Aba `Pedido` de cada unidade ativa |
| Saídas | `Consolidados`, `Resumo`, `Pedidos das Unidades` |

#### `consolidarPedidosAgendado()`

Versão chamada pelo gatilho diário (`CONFIG.CONSOLIDACAO_HORA`), em modo silencioso.

#### `menuConsolidarPedidos()` / `menuConsolidarEMarcar()`

Pontos de entrada. A segunda marca os itens como `Consolidado` após a leitura.

#### `marcarEnviadosComoConsolidado_(abaPedido)`

Muda o status na planilha da unidade, sinalizando que o pedido já foi recolhido.

#### `exibirResultadoConsolidacao(resultado, marcou)`

Monta a mensagem de resultado para o usuário.

### Abas relacionadas

Unidades: `Pedido` (leitura e, opcionalmente, escrita do status)
Central: `Consolidados`, `Resumo`, `Pedidos das Unidades` (escrita)

### Dependências

`Config.gs`, `Util.gs`, `Unidades.gs`, `RepararUnidade.gs`, `Resumo.gs`,
`TotalPorProduto.gs`

### Limitações

- Percorre todas as unidades em uma execução: é a rotina mais sujeita ao limite de
  6 minutos.
- Marcar como `Consolidado` é irreversível pela interface; o operador precisa voltar o
  status manualmente.

---

## `PedidoUnidade.gs`

**Caminho:** `src/sincronizacao/PedidoUnidade.gs`
**Funções:** 14
**Contexto:** executa na planilha da **unidade**

### Objetivo

Funções que rodam no contexto da planilha de uma unidade: menu próprio, envio do
pedido em lote e relatório local.

### Funções

#### `ehPlanilhaUnidadeAtiva_()`

| Retorno | `boolean` |
|---|---|

Detecta se a planilha aberta é de uma unidade, comparando o ID com o cadastro. É o que
permite ao mesmo projeto Apps Script se comportar de forma diferente conforme a
planilha aberta.

#### `onOpenUnidade_()` / `onOpenUnidadeInstalavel_()`

Montam o menu da unidade e ajustam o layout da aba `Pedido` na abertura.

A versão instalável existe porque o gatilho simples `onOpen` não tem autorização para
abrir outras planilhas.

#### `enviarPedidoCompletoUnidade_(ss, aba)`

| Parâmetros | Descrição |
|---|---|
| `ss` | `Spreadsheet` da unidade |
| `aba` | `Sheet` da aba `Pedido` |

Marca como `Enviado` todos os itens com quantidade preenchida e dispara a
sincronização com a central.

#### `menuEnviarPedidoCompleto()`

Ponto de entrada do menu da unidade, com confirmação antes de enviar.

#### `menuGerarRelatorioExcelUnidade()`

Gera um `.xlsx` com o que aquela unidade pediu. Permite à unidade ter o próprio
comprovante sem acesso à central.

#### `instalarGatilhoOnOpenUnidades_()` / `removerGatilhosOnOpenUnidades_()`

Instalam e removem o gatilho `onOpen` nas planilhas das unidades. Executadas a partir
da central.

#### `normalizarStatusPedido_(status)` / `obterTextoQuantidadePedido_(valor, display)`

Normalizam status e quantidade, tolerando variações de digitação.

### Abas relacionadas

`Pedido` da unidade (leitura e escrita)

### Dependências

`Config.gs`, `Util.gs`, `Unidades.gs`, `Bridge.gs`, `MesReferencia.gs`

### Limitações

- Depende de a unidade estar cadastrada na aba `Unidades` com o ID correto.
- O relatório local consome a cota de Drive da conta que o gerou.

---

## `MesReferencia.gs`

**Caminho:** `src/sincronizacao/MesReferencia.gs`
**Funções:** 17

### Objetivo

Gerenciar o mês de referência na aba `Pedido` das unidades e migrar planilhas do
layout antigo.

### Funções

#### `obterTextoMesReferenciaPedido_(data)` / `obterDataMesReferenciaPedido_(data)`

Formatam e interpretam o texto do mês de referência.

#### `aplicarMesReferenciaPedidoUnidade_(aba, data)`

Escreve o mês na linha 1 da aba `Pedido`.

#### `obterLinhaCabecalhoPedido_(aba)`

Localiza a linha do cabeçalho. Em planilhas migradas é a linha 2; em layouts antigos
pode ser a linha 1.

#### `planilhaPedidoUsaLayoutLegadoComData_(aba)` / `planilhaPedidoUsaLayoutComEstoque_(aba)`

Detectam layouts antigos: um com coluna Data, outro com coluna Estoque Atual.

#### `migrarPlanilhaPedidoRemoverColunaData_(aba, dataMes)` / `migrarPlanilhaPedidoRemoverColunaEstoque_(aba, dataMes)`

Convertem a planilha para o layout atual preservando os dados.

#### `iniciarMigracaoLayoutUnidadesEmLotes_(dataMes)` / `processarLoteMigracaoLayoutUnidades_()`

Migração em blocos, com continuação agendada. Salva o progresso em
`PropertiesService` e retoma de onde parou.

#### `menuMigrarLayoutMesReferenciaUnidades()`

Ponto de entrada da migração.

#### `atualizarMesReferenciaTodasUnidades_(data)`

Atualiza o mês em todas as unidades. Chamada no início de cada rodada.

### Abas relacionadas

`Pedido` das unidades (leitura e escrita)

### Dependências

`Config.gs`, `Util.gs`, `Unidades.gs`

### Limitações

- A migração altera a estrutura da planilha da unidade: **faça backup antes**.
- A detecção de layout é heurística, baseada no texto do cabeçalho.
- Em muitas unidades, a migração roda em vários blocos e leva mais de uma execução
  para concluir.
