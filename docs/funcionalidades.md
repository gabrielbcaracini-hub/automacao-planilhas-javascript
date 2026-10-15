# Funcionalidades

Cada funcionalidade do sistema, o que faz e como executar.

## Menu do operador

O menu **Sistema de Compras**, montado por `onOpen` em `src/core/Menu.gs`, expõe
apenas as ações do dia a dia:

```
Sistema de Compras
├── ➕ Nova linha na aba Produtos
├── ⚡ Atualizar produtos nas unidades agora
├── ─────────────────────────────
├── 📥 Puxar pedidos das unidades
├── ─────────────────────────────
├── 📥 Baixar relatório Excel
├── 📥 Relatório de Pregões ativos (Excel)
├── 📥 Relatório de Pregões inativos (Excel)
├── ─────────────────────────────
├── 🗑️ Nova rodada — limpar pedidos
├── ─────────────────────────────
└── ❓ Ajuda
```

As funções técnicas e de manutenção **não aparecem no menu** por decisão de projeto:
são executadas pelo editor do Apps Script, por quem administra o sistema. Isso evita
que o operador dispare por engano uma operação destrutiva.

---

## Catálogo de produtos

### Cadastrar um produto novo

**Menu:** ➕ Nova linha na aba Produtos
**Função:** `menuInserirLinhaNovoProduto`

Cria uma linha em branco já posicionada no lugar certo da aba `Produtos`.

1. Clique em qualquer linha da categoria onde o item entra.
2. Acione o menu.
3. O sistema pergunta o **número do pregão** que será cadastrado.
4. A linha é criada, formatada e com `Ativo = Sim`.
5. Preencha Código, Produto, Unidade e o texto completo do Pregão.

#### Posicionamento quando a categoria tem vários pregões

Este é o comportamento central da funcionalidade. Uma categoria pode conter pregões
com números diferentes:

```
ELETRODOMÉSTICOS
   FREEZER HORIZONTAL 2 PORTAS 534L    PREGÃO 05/2025  …
   BEBEDOURO INDUSTRIAL 25L INOX       PREGÃO 05/2025  …
   FOGÃO INDUSTRIAL 4 BOCAS            PREGÃO 156/2025 …
   FORNO MICRO-ONDAS 32L               PREGÃO 156/2025 …
```

Ao informar o número do pregão, o sistema posiciona a linha assim:

| Número informado | Posição da nova linha |
|---|---|
| `05` | Depois do último item do pregão 05, antes do bloco 156 |
| `156` | Depois do último item do pregão 156 |
| `80` (ainda não existe) | Entre os blocos 05 e 156, na ordem numérica |
| Nenhum | No fim da categoria |

Isso mantém a sequência por número de pregão mesmo quando um pregão antigo recebe um
item novo meses depois.

**Funções envolvidas:**

| Função | Papel |
|---|---|
| `obterContextoInsercaoProduto_` | Descobre categoria e pregão a partir da linha selecionada |
| `encontrarLinhaParaNovoPregao_` | Varre a categoria e calcula a linha de inserção |
| `extrairNumeroPregaoCadastro_` | Extrai o número do pregão do texto |

Depois de cadastrar, execute **⚡ Atualizar produtos nas unidades agora** para
distribuir.

### Desativar um produto

Mude `Ativo` para `Não` na aba `Produtos`.

O item para de ser distribuído às unidades, mas **permanece cadastrado**. Isso é
proposital: um pregão suspenso pode voltar a valer. A função `removerPedidosProdutosInativosNaMae_`
retira os pedidos pendentes desse item, e a nova rodada preserva os itens inativos.

### Organizar a aba por categoria

**Função:** `menuOrganizarProdutosPorCategoria`

Reordena a aba `Produtos` agrupando por categoria, insere as faixas de título e
reaplica a formatação. Internamente usa `organizarProdutosPorCategoria_` e, quando
necessário, `recriarAbaProdutosSemTabela_`.

### Importar um lote de pregões novos

**Função:** `menuImportarNovosPregoesMatrizSetembro2026`

Acrescenta à aba `Produtos` apenas os itens que **ainda não existem**, comparando por
nome normalizado. Nada é sobrescrito. Os códigos continuam a partir do maior código
já cadastrado.

A importação também roda sozinha na abertura da planilha
(`importarMatrizSetembro2026SePendente_`), uma única vez, controlada por uma flag em
`DocumentProperties`.

Para usar com os seus dados, substitua o array `NOVOS_PREGOES_MATRIZ_SETEMBRO_2026`
em `src/catalogo/NovosPregoesMatriz.gs`.

### Liberar categorias do mês

**Função:** `menuAtualizarControleCategorias`

Sincroniza a aba `Categorias` com as categorias existentes na aba `Produtos`. Marque
`Sim` ou `Não` na coluna "Liberada neste mês" para controlar o que cada unidade vê
na rodada.

Sem a aba `Categorias`, todas as categorias são consideradas liberadas.

### Ordenação dos pregões

Pregões são ordenados **numericamente**, não como texto. A comparação usa, em ordem:

1. o número do pregão;
2. o ano;
3. a data de vigência ("DE dd/mm/aaaa").

Por isso `14/2025` vem antes de `101/2025` — a ordenação alfabética faria o contrário.
A mesma regra vale no cadastro, na aba `Resumo` e nos relatórios.

**Funções:** `extrairChaveOrdemPregao_`, `compararPregoesOrdem_`, `parsearPartesPregao_`
(em `src/relatorios/RelatorioPregoes.gs`).

---

## Sincronização

### Distribuir o catálogo para as unidades

**Menu:** ⚡ Atualizar produtos nas unidades agora
**Função:** `menuAtualizarProdutosUnidadesAgora` → `sincronizarCatalogoTodasUnidades`

Para cada unidade ativa:

1. lê os produtos ativos das categorias liberadas;
2. reescreve a aba `Catalogo`;
3. monta a aba `Pedido` com os produtos agrupados por categoria;
4. aplica as validações de dados na coluna Status;
5. escreve o mês de referência no topo;
6. aplica a formatação visual.

**Quantidades já preenchidas são preservadas** quando o produto continua no catálogo.

A distribuição é **sempre manual**. Não existe gatilho que propague o catálogo
automaticamente: uma edição acidental não reescreve as planilhas de todas as unidades
no meio de uma rodada.

### Enviar o pedido (na unidade)

**Menu da unidade:** Enviar pedido completo
**Função:** `menuEnviarPedidoCompleto` → `enviarPedidoCompletoUnidade_`

Marca como `Enviado` todos os itens com quantidade preenchida e dispara a
sincronização com a central.

Também é possível marcar `Enviado` item a item na coluna Status: o gatilho `onEdit`
detecta a mudança.

### Puxar os pedidos

**Menu:** 📥 Puxar pedidos das unidades
**Função:** `menuAtualizarResumoPedidosAgora`

Percorre todas as unidades ativas, lê os itens marcados como `Enviado` e atualiza
`Consolidados`, `Resumo` e `Pedidos das Unidades`.

Use quando quiser forçar a atualização sem esperar o gatilho automático, ou quando
uma unidade avisar que enviou e o dado não apareceu.

A versão com diagnóstico detalhado é `menuPuxarPedidosRapido`
(`consolidarPedidosComDiagnostico`), que informa o que foi lido de cada unidade e por
que algo não foi importado.

### Sincronização automática

**Função:** `menuConectarTudoAgora`

Instala o gatilho `onEdit` em todas as unidades e faz a primeira puxada.

Fluxo quando a unidade edita:

```
unidade marca Enviado
   → onEditPedidoUnidade dispara
   → a unidade entra na fila em PropertiesService
   → gatilho temporizado roda após ~7 s
   → lê a unidade e atualiza Consolidados, Resumo e Pedidos das Unidades
```

O debounce (`CONFIG.SYNC_DEBOUNCE_MS`) agrupa edições em sequência numa execução só.

**Diagnóstico:** `menuDiagnosticarSyncAutomatico` lista os gatilhos ativos e aponta
unidades sem gatilho.

### Reconectar

**Função:** `menuReconectarPlanilhas`

Recria os gatilhos quando a conexão cai — por exemplo, se uma planilha de unidade foi
recriada ou se os gatilhos foram removidos.

### Mês de referência

**Função:** `atualizarMesReferenciaTodasUnidades_`

Escreve o mês da rodada no topo da aba `Pedido` de cada unidade. A migração de layout
(`menuMigrarLayoutMesReferenciaUnidades`) converte planilhas do formato antigo, que
usava uma coluna Data, para o formato atual. Processa em lotes com continuação
agendada.

---

## Consolidação

### `Consolidados`

Registro linha a linha: unidade, produto, quantidade, observação e data. É a fonte de
verdade dos pedidos — `Resumo` e `Pedidos das Unidades` são derivadas e podem ser
reconstruídas a partir dela (`reconstruirResumoDeConsolidadosExistentes_`).

### `Resumo`

Matriz produto × unidade, agrupada por pregão e categoria, montada por
`escreverResumoAgrupadoFromConsolidados_`.

- faixas de pregão e de categoria com mesclagem de células;
- uma coluna por unidade, na ordem de `DADOS_INICIAIS.unidades`;
- painel de pendências no topo (`desenharPainelPendenciasUnidades_`);
- indicador da última atualização.

**Reaplicar o visual:** `menuAtualizarVisualResumo`.

### `Pedidos das Unidades`

Duas colunas: "Já pediram" e "Ainda não pediram".

Uma unidade conta como "já pediu" se houver registro em `Consolidados`, quantidade
lançada no `Resumo` ou itens marcados como `Enviado` na própria planilha. A
verificação tripla evita falso negativo quando um pedido chegou por um caminho
diferente.

Atualiza ao puxar pedidos e ao abrir a central.

---

## Relatórios

Todos os relatórios são salvos no Drive da conta que executa e a função devolve o
link do arquivo.

| Relatório | Menu / Função | Formato |
|---|---|---|
| Pedidos do mês | 📥 Baixar relatório Excel → `menuGerarRelatorioExcel` | `.xlsx` |
| Pregões ativos | 📥 Relatório de Pregões ativos → `menuGerarRelatorioPregoes` | `.xlsx` |
| Pregões inativos | 📥 Relatório de Pregões inativos → `menuGerarRelatorioPregoesInativos` | `.xlsx` |
| Resumo em PDF | `menuGerarPdfResumo` | `.pdf` |
| Administrativo | `menuGerarRelatorioPersonalizado` | `.docx` |
| Guia do sistema | `menuGerarPdfGuiaSistema` | `.pdf` |

### Relatório de pedidos do mês

Hierarquia: **pregão → produto → unidades e quantidades**. Inclui apenas itens com
quantidade pedida.

A geração usa uma planilha temporária formatada e a exporta em `.xlsx` pelo endpoint
de exportação do Google, via `UrlFetchApp` (`exportarPlanilhaComoXlsx_`).

### Relatórios de pregões

Lista cada pregão e os produtos que ele disponibiliza, ordenados por número e data.
O modo "ativos" considera os itens com `Ativo = Sim`; o modo "inativos", os demais.
Produtos recém-cadastrados aparecem assim que estão na aba `Produtos`.

### Relatório administrativo

Documento Word com visão por unidade, cabeçalho com o nome da organização
(`CONFIG.ORGANIZACAO.NOME`) e logotipo opcional (`CONFIG.ORGANIZACAO.LOGO_URL`).

---

## Manutenção

### Nova rodada

**Menu:** 🗑️ Nova rodada — limpar pedidos
**Função:** `menuResetNovaRodadaCompras`

Encerra o ciclo e prepara o próximo.

| Apaga | Preserva |
|---|---|
| Quantidades e observações nas unidades | Catálogo da aba `Produtos` |
| Status (volta para `Rascunho`) | Itens com `Ativo = Não` |
| Conteúdo de `Consolidados` | Cadastro da aba `Unidades` |
| Quantidades da aba `Resumo` | Lista de produtos nas unidades |
| Situação em `Pedidos das Unidades` | Gatilhos instalados |

Itens marcados como `Ativo = Não` **não são apagados**: um pregão suspenso pode voltar
a valer.

A limpeza é feita **no lugar**, alterando apenas valores. O catálogo e a formatação
não são reconstruídos do zero — essa decisão existe porque a reconstrução completa
estourava o limite de 6 minutos. Se ainda assim o tempo se esgotar,
`agendarContinuacaoResetNovaRodada_` registra o progresso e `continuarResetNovaRodada_`
retoma automaticamente nas unidades restantes.

### Reparo do sistema

**Função:** `menuRepararSistemaCompleto`

Executa a manutenção em passos, em segundo plano, sem travar a interface:

1. remove abas legadas e indevidas;
2. recria abas ausentes com o cabeçalho correto;
3. reaplica validações e formatação;
4. revalida os gatilhos.

Cada passo agenda o seguinte (`agendarProximoPassoReparo`), o que contorna o limite de
execução.

### Reparar uma unidade específica

**Função:** `menuRepararUnidadeCentroEPuxar`

Recria a aba `Pedido` de uma unidade preservando os dados e puxa os itens enviados.
Útil quando a planilha da unidade foi convertida em tabela tipada ou teve o layout
alterado de forma incompatível.

### Corrigir IDs das unidades

**Função:** `menuCorrigirIdsUnidades` → `corrigirIdsUnidadesDoDrive`

Procura no Drive a planilha de cada unidade pelo nome do arquivo e atualiza o ID na
aba `Unidades`. Resolve o caso em que a planilha foi recriada e o ID antigo deixou de
existir.

### Visibilidade das abas

**Função:** `menuAplicarVisaoClienteAbas`

Oculta as abas administrativas (`Consolidados`, `Log`) e mantém visíveis apenas as
operacionais. Roda também na abertura da planilha.

### Limpar produtos de teste

**Função:** `menuLimparProdutosTeste`

Remove itens que atendem aos critérios de `ehProdutoTesteLixo_` — nomes genéricos de
teste, linhas sem pregão válido. **Não remove** itens reais marcados como inativos.

### Diagnóstico

| Função | O que mostra |
|---|---|
| `menuListarGatilhos` | Todos os gatilhos registrados |
| `menuDiagnosticarSyncAutomatico` | Estado da sincronização por unidade |
| `menuDiagnosticoPedidos` | O que cada unidade tem na aba `Pedido` |
| `consolidarPedidosComDiagnostico` | Puxada com relatório detalhado de leitura |

### Ajuda

**Menu:** ❓ Ajuda
**Função:** `menuAjuda`

Abre `src/interface/Ajuda.html` numa janela modal, com a explicação das abas e do
fluxo de uso.

---

## Funcionalidades desativadas

Mantidas no código para compatibilidade, mas fora de uso:

| Item | Situação |
|---|---|
| `TesteFluxo.gs` | Desligado. Gravaria um pedido fictício numa unidade real |
| `EstoqueUnidades.gs` | Legado. O campo "Estoque Atual" foi removido das unidades |
| Abas `Lista Pedidos`, `Painel de Compras`, `Painel Dados` | Legadas. Removidas automaticamente quando encontradas |
