# Módulos de manutenção

Instalação, reparo, reinício de rodada e diagnóstico.

> Vários módulos desta página executam **operações destrutivas**. Leia a seção de
> limitações antes de rodar qualquer função em ambiente com dados reais.

---

## `Setup.gs`

**Caminho:** `src/manutencao/Setup.gs`
**Funções:** 8

### Objetivo

Preparar a planilha central e as planilhas das unidades, e instalar os gatilhos.

### Funções

#### `setupPlanilhaMae()`

| Parâmetros | Nenhum |
|---|---|
| Saída | Abas criadas com os cabeçalhos corretos |

Cria `Produtos`, `Categorias`, `Unidades`, `Consolidados`, `Resumo`,
`Pedidos das Unidades` e `Log`.

**Idempotente:** abas já existentes são preservadas. Exemplos iniciais só são gravados
quando a aba `Produtos` está vazia.

É a primeira função a executar numa instalação nova.

#### `setupPlanilhaFilha(spreadsheetId, mostrarAlerta)`

| Parâmetro | Tipo | Descrição |
|---|---|---|
| `spreadsheetId` | `string` | ID da planilha da unidade |
| `mostrarAlerta` | `boolean` | Exibe o resultado na interface |

Prepara uma planilha de unidade: cria `Pedido` e `Catalogo`, aplica cabeçalho,
validações e formatação.

#### `instalarGatilhos()`

Registra os gatilhos do projeto:

| Gatilho | Tipo |
|---|---|
| `consolidarPedidosAgendado` | Temporizado, diário em `CONFIG.CONSOLIDACAO_HORA` |
| `onOpenUnidadeInstalavel_` | Instalável, por planilha de unidade |

#### `removerGatilhosDoProjeto()`

Remove **todos** os gatilhos do projeto. Usada na desinstalação e para limpar
gatilhos órfãos.

#### `menuListarGatilhos()`

Lista os gatilhos registrados, com função e tipo. Primeira verificação quando a
automação não dispara.

#### `aplicarValidacaoAtivo(aba, coluna)`

Aplica a lista suspensa `Sim`/`Não`.

#### `removerAbaPadraoVazia(ss)`

Remove a aba "Página1" criada por padrão em planilhas novas.

### Dependências

`Config.gs`, `Util.gs`, `Dropdown.gs`, `Unidades.gs`

### Limitações

- `removerGatilhosDoProjeto` não pede confirmação e desliga toda a automação.
- O limite de 20 gatilhos por script por usuário se aplica.

---

## `SetupUnidades.gs`

**Caminho:** `src/manutencao/SetupUnidades.gs`
**Funções:** 14

### Objetivo

Importação inicial: criar as planilhas das unidades e popular o catálogo a partir de
`DADOS_INICIAIS`.

### Funções

#### `menuImportarUnidadesCompleto()` / `importarUnidadesCompleto()`

Fluxo completo de instalação:

1. `setupPlanilhaMaeSilencioso` — garante as abas;
2. `popularProdutosUnidades` — grava o catálogo semente;
3. `criarOuVincularPlanilhasUnidades` — cria ou reaproveita as planilhas;
4. `preencherPedidoComCatalogo` — monta a aba `Pedido` de cada uma.

Protegida por `adquirirLock` para impedir execução simultânea.

#### `popularProdutosUnidades()`

Grava `DADOS_INICIAIS.produtos` na aba `Produtos`.

> **Substitui o conteúdo da aba.** Não execute com um catálogo já em uso.

#### `criarOuVincularPlanilhasUnidades(resultado)`

Para cada unidade de `DADOS_INICIAIS.unidades`:

| Situação | Ação |
|---|---|
| ID preenchido em `idsPlanilhas` | Reaproveita a planilha existente |
| ID vazio, planilha existe no Drive com o nome esperado | Vincula |
| Nenhum dos dois | Cria uma planilha nova |

Registra nome e ID na aba `Unidades`.

#### `preencherPedidoComCatalogo(ssFilha, nomesProdutos, produtosComCategoria, opcoes)`

Monta a aba `Pedido` com os produtos agrupados por categoria, aplicando as linhas de
título e as validações.

#### `aplicarIdsPlanilhasUnidades()`

Grava na aba `Unidades` os IDs definidos em `DADOS_INICIAIS.idsPlanilhas`.

#### `corrigirIdsUnidadesDoDrive()` / `menuCorrigirIdsUnidades()`

Procura cada planilha pelo nome do arquivo no Drive e corrige o ID cadastrado.
Resolve o caso da planilha recriada, que ganhou um ID novo.

#### `menuDiagnosticoPedidos()`

Relatório do que cada unidade tem na aba `Pedido`: quantas linhas, quantos itens
enviados, se o layout está correto.

#### `menuAplicarCategoriasTodasUnidades()` / `testarCategoriasUnidadePorNome_(nomeUnidade)`

Aplicam o agrupamento por categoria nas unidades. A versão de teste roda em uma
unidade só, para conferir o resultado antes de aplicar a todas.

### Abas relacionadas

Central: `Produtos`, `Unidades` (escrita)
Unidades: `Pedido`, `Catalogo` (escrita)

### Dependências

`Config.gs`, `Util.gs`, `DadosIniciais.gs`, `Produtos.gs`, `Dropdown.gs`,
`VisualUnidades.gs`

### Limitações

- `popularProdutosUnidades` **substitui** a aba `Produtos`.
- A criação de planilhas consome cota de Drive e pode exceder o limite de execução
  com muitas unidades — a função retoma se executada de novo.
- As planilhas criadas pertencem à conta que executou.

---

## `ResetRodada.gs`

**Caminho:** `src/manutencao/ResetRodada.gs`
**Funções:** 13

### Objetivo

Encerrar a rodada: zerar os pedidos na central e nas unidades, preservando o catálogo.

### O que apaga e o que preserva

| Apaga | Preserva |
|---|---|
| Quantidades e observações nas unidades | Catálogo da aba `Produtos` |
| Status (volta para `Rascunho`) | **Itens com `Ativo = Não`** |
| Conteúdo de `Consolidados` | Cadastro da aba `Unidades` |
| Quantidades da aba `Resumo` | Lista de produtos nas unidades |
| Situação em `Pedidos das Unidades` | Gatilhos instalados |

Itens inativos **não são apagados**: um pregão suspenso pode voltar a valer.

### Estratégia

A limpeza é feita **no lugar**, alterando apenas valores. O catálogo, a formatação e
as validações não são reconstruídos.

Essa decisão veio de um problema concreto: a versão anterior recriava o catálogo e a
formatação de todas as unidades a cada rodada e estourava o limite de 6 minutos. A
limpeza por valores é ordens de grandeza mais rápida.

### Funções

#### `menuResetNovaRodadaCompras()`

Ponto de entrada do menu **🗑️ Nova rodada — limpar pedidos**, com confirmação.

#### `executarResetNovaRodadaCompras_()`

Executa a limpeza:

1. zera as quantidades do `Resumo` (`zerarQuantidadesAbaResumo_`);
2. limpa `Consolidados`;
3. zera `Pedidos das Unidades`;
4. percorre as unidades chamando `resetarPedidosPlanilhaUnidade_`;
5. se chegar perto do limite, agenda a continuação.

#### `resetarPedidosPlanilhaUnidade_(ssFilha, produtosCatalogo)`

Limpa quantidades e observações e volta o status para `Rascunho`, preservando a lista
de produtos e a formatação.

#### `agendarContinuacaoResetNovaRodada_(unidades, falhas, unidadesOk)` / `continuarResetNovaRodada_()`

Salvam o progresso em `PropertiesService` e retomam nas unidades restantes.

#### `removerProdutosInativosAbaProdutos_()`

Remove itens inativos da aba `Produtos`. **Não é chamada pela nova rodada** — existe
como limpeza manual para quando o administrador decide descartar definitivamente um
pregão encerrado.

#### `menuLimparProdutosTeste()` / `removerProdutosTesteAbaProdutos_(ssMae, detalhado, pularVisual)`

Removem itens de teste.

#### `ehProdutoTesteLixo_(produto, pregao, categoria)` / `ehTextoLixoTeste_(texto)`

Critérios de identificação: nomes genéricos de teste, linhas sem pregão válido.
**Não marcam** itens reais que estejam inativos.

### Abas relacionadas

Central: `Resumo`, `Consolidados`, `Pedidos das Unidades`, `Produtos`
Unidades: `Pedido`

### Dependências

`Config.gs`, `Util.gs`, `Unidades.gs`, `Resumo.gs`, `TotalPorProduto.gs`

### Limitações

- **Operação destrutiva.** Os pedidos da rodada anterior são perdidos. Gere os
  relatórios antes.
- Com muitas unidades, roda em vários blocos: a mensagem de conclusão pode aparecer
  antes de todas as unidades terem sido limpas. A continuação é automática.
- Uma unidade inacessível é reportada como falha e precisa ser limpa manualmente.

---

## `Reparo.gs`

**Caminho:** `src/manutencao/Reparo.gs`
**Funções:** 7

### Objetivo

Manutenção corretiva em etapas, executada em segundo plano, sem travar a interface.

### Funções

#### `menuRepararSistemaCompleto()`

Inicia o reparo e devolve o controle ao usuário imediatamente.

#### `executarReparoEmPassos()`

Executa um passo por vez e agenda o seguinte:

1. remove abas legadas e indevidas;
2. recria abas ausentes com o cabeçalho correto;
3. reaplica validações e formatação;
4. revalida os gatilhos.

#### `agendarProximoPassoReparo(depoisMs)` / `removerTriggersReparo()`

Controlam o encadeamento dos passos.

#### `limparAbasErradasTodasFilhas()` / `limparAbasErradasDaFilha(ssFilha)`

Removem das unidades abas que só deveriam existir na central. A lista está em
`ABAS_SO_MAE`.

### Dependências

`Config.gs`, `Util.gs`, `Setup.gs`, `Unidades.gs`, `Dropdown.gs`

### Limitações

- Remove abas sem confirmação por aba.
- Se interrompido no meio, pode deixar gatilhos de reparo órfãos: use
  `removerTriggersReparo`.

---

## `RepararUnidade.gs`

**Caminho:** `src/manutencao/RepararUnidade.gs`
**Funções:** 5

### Objetivo

Leitura tolerante e reparo da planilha de **uma** unidade.

### Funções

#### `lerPedidosEnviadosFlexivel(aba)`

| Parâmetros | `aba` — `Sheet` da aba `Pedido` |
|---|---|
| Retorno | `{enviados, diagnostico, todas}` |

Função de leitura mais usada do sistema. Detecta as colunas pelo cabeçalho em vez de
posições fixas, o que mantém a leitura funcionando com layouts antigos.

O campo `diagnostico` explica o que foi encontrado — é o que alimenta os relatórios de
diagnóstico.

#### `detectarColunasPedido_(values)`

| Retorno | Índices das colunas Produto, Quantidade, Observação e Status |
|---|---|

Procura o cabeçalho nas primeiras linhas e casa os nomes por texto normalizado.

#### `menuRepararUnidadeCentroEPuxar()`

Repara uma unidade e puxa os pedidos dela. Útil quando a aba `Pedido` foi convertida
em tabela tipada do Sheets, formato que impede a manipulação de linhas pelo script.

#### `recriarAbaPedidoLimpa_(ss, abaAntiga, linhas)`

Recria a aba `Pedido` sem tabela tipada, preservando os dados.

#### `resolverIdUnidadePeloNome_(nomeUnidade)`

Procura a planilha pelo nome do arquivo no Drive. Fallback quando o ID cadastrado não
abre.

### Dependências

`Config.gs`, `Util.gs`, `Unidades.gs`, `Dropdown.gs`

### Limitações

- `resolverIdUnidadePeloNome_` exige o escopo `drive` completo.
- Se houver mais de um arquivo com o mesmo nome no Drive, a escolha pode ser ambígua.
- `recriarAbaPedidoLimpa_` recria a aba: formatações personalizadas se perdem.

---

## `PullDiagnostico.gs`

**Caminho:** `src/manutencao/PullDiagnostico.gs`
**Funções:** 7

### Objetivo

Puxada manual dos pedidos com relatório detalhado do que foi lido em cada unidade.

Serve para responder à pergunta "por que o pedido da unidade X não apareceu?".

### Funções

#### `menuPuxarPedidosRapido()`

Ponto de entrada, com diagnóstico na tela.

#### `consolidarPedidosComDiagnostico()`

| Retorno | Resultado da consolidação com o diagnóstico por unidade |
|---|---|

Para cada unidade informa: se abriu, se encontrou a aba `Pedido`, quantas linhas leu,
quantos itens estavam `Enviado` e por que algum item foi descartado.

#### `resolverIdsUnidadesPeloDrive()` / `completarIdsUnidadesFaltantes_()`

Preenchem IDs ausentes buscando as planilhas pelo nome no Drive.

#### `encontrarIdPedidoMaisRecente_(nomeUnidade)`

Quando há mais de uma planilha com o nome esperado, escolhe a mais recente.

#### `mesclarConsolidadosDeUnidadesNaoLidas_(ssMae, linhasCons, unidadesLidas)`

Preserva em `Consolidados` os dados de unidades que não puderam ser lidas nesta
passagem. **Evita que uma unidade temporariamente inacessível seja apagada do
consolidado.**

### Dependências

`Config.gs`, `Util.gs`, `Unidades.gs`, `RepararUnidade.gs`, `Resumo.gs`,
`TotalPorProduto.gs`

### Limitações

- Mais lenta que a puxada comum, por causa do diagnóstico.
- A busca no Drive exige o escopo `drive` completo.

---

## `TesteFluxo.gs`

**Caminho:** `src/manutencao/TesteFluxo.gs`
**Funções:** 3
**Situação:** desativado

### Objetivo

Teste automático de ponta a ponta: gravaria um pedido em uma unidade e verificaria a
chegada na central.

### Por que está desligado

`menuTesteFluxoCompleto` exibe um aviso e não executa nada. O teste gravaria um
**pedido fictício em uma planilha de unidade real**, o que contaminaria os dados da
rodada em andamento.

O código foi mantido como registro da estratégia de teste; a execução foi desabilitada
deliberadamente.

### Funções

| Função | Situação |
|---|---|
| `menuTesteFluxoCompleto` | Exibe aviso e encerra |
| `executarTesteFluxoCompleto_` | Implementação preservada, não chamada |
| `lerPedidosFallbackSimples_` | **Em uso** — fallback de leitura usado por outros módulos |

> `lerPedidosFallbackSimples_` continua ativa: é o fallback quando
> `lerPedidosEnviadosFlexivel` não reconhece o layout.

### Limitações

Para reativar o teste, use uma planilha de unidade descartável. Nunca aponte para uma
unidade em produção.

---

## `EstoqueUnidades.gs`

**Caminho:** `src/manutencao/EstoqueUnidades.gs`
**Funções:** 4
**Situação:** legado

### Objetivo

Funções do campo "Estoque Atual", que foi removido das planilhas das unidades.

Mantidas para compatibilidade com planilhas que ainda tenham a coluna e para a
migração de dados antigos. **Não use os menus deste módulo.**

### Funções

| Função | Finalidade |
|---|---|
| `menuTestarEstoqueConselhoTutelar` | Teste pontual do campo em uma unidade |
| `testarEstoqueUnidadePorNome_` | Implementação do teste |
| `menuAplicarEstoqueTodasUnidades` | Aplicaria o campo em todas as unidades |
| `migrarConsolidadosParaEstoque_` | Migração de dados antigos |

### Limitações

Executar estas funções reintroduz um campo que o fluxo atual não usa. A coluna
`ESTOQUE_ATUAL` em `CONFIG.CONSOLIDADOS` existe apenas para manter a compatibilidade
da estrutura de `Consolidados`.
