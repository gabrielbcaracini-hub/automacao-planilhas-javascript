# Arquitetura

Este documento descreve como o sistema está organizado: onde o código vive, quais
estruturas de dados existem, e como os módulos se relacionam.

## 1. Topologia

O sistema tem dois tipos de planilha.

### Planilha central

Uma única planilha do Google Sheets com um projeto Apps Script vinculado. **Todo o
código do sistema está aqui.** Ela mantém o catálogo de produtos, o cadastro das
unidades, os dados consolidados e gera os relatórios.

### Planilhas de unidade

Uma planilha por unidade solicitante. Elas **não têm código próprio**. Recebem:

- a aba `Pedido`, montada e formatada pela central;
- a aba `Catalogo`, usada como fonte das validações de dados;
- gatilhos instaláveis (`onEdit` e `onOpen`) que apontam de volta para o projeto
  Apps Script da central.

Essa escolha concentra a manutenção: uma correção no código vale imediatamente para
todas as unidades, sem precisar republicar nada em cada planilha.

```
                     ┌──────────────────────────────┐
                     │      PLANILHA CENTRAL        │
                     │   (projeto Apps Script)      │
                     └───────┬──────────────▲───────┘
                             │              │
          catálogo (manual)  │              │  pedidos (onEdit + manual)
                             ▼              │
   ┌─────────────┬─────────────┬────────────┴──┬─────────────┐
   │ UNIDADE A   │ UNIDADE B   │  UNIDADE C    │     ...     │
   └─────────────┴─────────────┴───────────────┴─────────────┘
```

## 2. Estrutura de dados

As abas do Google Sheets funcionam como tabelas. Os nomes são configuráveis em
`CONFIG.ABAS`.

### Abas da planilha central

#### `Produtos` — catálogo mestre

| Coluna | Campo | Descrição |
|---|---|---|
| A | Código | Identificador sequencial do item |
| B | Categoria | Agrupamento (`PAPELARIA`, `GÊNERO ALIMENTÍCIO`, …) |
| C | Produto | Descrição do item |
| D | Unidade | Unidade de medida |
| E | Pregão | Texto completo: número, data de vigência, fornecedor e validade |
| F | Ativo | `Sim` / `Não` — controla se o item é distribuído |

A aba também contém **linhas de título de categoria**, identificadas por
`ehLinhaTituloCategoriaProdutos_`. Elas são decorativas e são ignoradas na leitura
dos dados.

O campo `Pregão` é um texto livre com formato convencionado:

```
PREGÃO 156/2025 – DE 29/01/2026 – FORNECEDOR EXEMPLO H LTDA – 29/01/2027
       └──┬──┘      └────┬────┘    └────────┬────────────┘   └────┬────┘
        número      vigência a partir de       fornecedor        validade
```

`parsearPartesPregao_` e `extrairChaveOrdemPregao_` extraem número, ano e data desse
texto para permitir a ordenação numérica correta.

#### `Categorias` — liberação mensal

| Coluna | Campo |
|---|---|
| A | Categoria |
| B | Liberada neste mês (`Sim` / `Não`) |

Funciona como filtro: só as categorias marcadas como liberadas são distribuídas às
unidades na rodada. Se a aba não existir, todas as categorias são consideradas
liberadas — isso mantém compatibilidade com instalações antigas.

#### `Unidades` — cadastro das planilhas filhas

| Coluna | Campo |
|---|---|
| A | Nome Unidade |
| B | Spreadsheet ID |
| C | Ativo |

É o índice que conecta a central às unidades. Quando um ID está vazio ou errado,
`resolverIdUnidadePeloNome_` procura a planilha pelo nome do arquivo no Drive.

#### `Consolidados` — dados brutos

| Coluna | Campo |
|---|---|
| A | Unidade |
| B | Produto |
| C | Estoque Atual *(legado)* |
| D | Quantidade |
| E | Observação |
| F | Data |

Registro linha a linha de cada item pedido. É a **fonte de verdade dos pedidos**:
as abas `Resumo` e `Pedidos das Unidades` são derivadas daqui e podem ser
reconstruídas a qualquer momento.

#### `Resumo` — matriz operacional

Layout dinâmico, construído em tempo de execução:

```
ITEM                              │ UNIDADE A │ UNIDADE B │ UNIDADE C │ …
──────────────────────────────────┼───────────┼───────────┼───────────┼───
▸ PREGÃO 10/2025 — FORNECEDOR A                                           ← faixa de pregão
  ▸ PAPELARIA                                                             ← faixa de categoria
    PAPEL SULFITE BRANCO A4 RESMA │     10    │     25    │      –    │
    CANETA ESFEROGRÁFICA AZUL     │      5    │      –    │     12    │
▸ PREGÃO 35/2025 — FORNECEDOR C
  ▸ GÊNERO ALIMENTÍCIO
    ARROZ TIPO 1 PACOTE 5 KG      │     40    │     30    │     20    │
```

Uma coluna por unidade, na ordem definida em `DADOS_INICIAIS.unidades`. Os blocos são
agrupados por pregão e, dentro dele, por categoria. A ordenação dos pregões usa
`compararPregoesOrdem_`, que compara número, ano e data de vigência numericamente.

#### `Pedidos das Unidades` — situação da rodada

| Coluna | Campo |
|---|---|
| A | Já pediram |
| B | Ainda não pediram |

Duas listas paralelas. Atualizada a cada puxada de pedidos e ao abrir a planilha.

#### `Log` — auditoria técnica

| Coluna | Campo |
|---|---|
| A | Data/Hora |
| B | Ação |
| C | Detalhe |

Oculta para o operador. Alimentada por `registrarLog`.

### Abas da planilha de unidade

#### `Pedido` — interface da unidade

| Linha | Conteúdo |
|---|---|
| 1 | Mês de referência |
| 2 | Cabeçalho |
| 3+ | Itens e linhas de categoria |

| Coluna | Campo |
|---|---|
| A | Produto |
| B | 🛒 Quantidade do Pedido |
| C | Observação |
| D | Status (`Rascunho` / `Enviado` / `Consolidado`) |

As linhas de categoria (`ehLinhaCategoriaPedido_`) separam visualmente os blocos e são
ignoradas na leitura.

#### `Catalogo` — apoio

Lista de produtos válidos, usada como origem das validações de dados. Protegida contra
edição por `protegerCatalogo`.

## 3. Módulos

### `core/` — base do sistema

| Arquivo | Responsabilidade |
|---|---|
| `Config.gs` | Objeto `CONFIG` com toda a parametrização: IDs, nomes de abas, posições de coluna, tema visual, tempos |
| `Menu.gs` | `onOpen` da central, construção do menu e janela de ajuda |
| `Util.gs` | Funções compartilhadas: acesso a abas, normalização de texto, locks, logs, visibilidade de abas, abertura de planilhas filhas |
| `Unidades.gs` | Leitura do cadastro de unidades, com fallback para `DADOS_INICIAIS` |

`Util.gs` é a camada mais usada: praticamente todo módulo depende de
`obterPlanilhaMae`, `obterAbaObrigatoria`, `normalizarTexto` e `registrarLog`.

### `catalogo/` — gestão dos produtos

| Arquivo | Responsabilidade |
|---|---|
| `Produtos.gs` | Leitura, validação, organização por categoria, inserção posicionada de linhas e camada visual da aba `Produtos` |
| `Categorias.gs` | Controle de quais categorias estão liberadas no mês |
| `Dropdown.gs` | Aba `Catalogo` e validações de dados nas unidades |
| `NovosPregoesMatriz.gs` | Importação incremental de lotes de pregões |

### `sincronizacao/` — ligação entre as planilhas

| Arquivo | Responsabilidade |
|---|---|
| `Sincronizacao.gs` | Distribuição do catálogo da central para as unidades |
| `Bridge.gs` | Caminho inverso: gatilhos `onEdit`, fila, debounce e escrita dos pedidos na central |
| `Consolidacao.gs` | Varredura de todas as unidades e montagem de `Consolidados` |
| `PedidoUnidade.gs` | Funções que rodam no contexto da planilha da unidade |
| `MesReferencia.gs` | Mês de referência e migração de layout das unidades |

### `consolidacao/` — visões derivadas

| Arquivo | Responsabilidade |
|---|---|
| `Resumo.gs` | Construção e formatação da matriz produto × unidade |
| `TotalPorProduto.gs` | Aba de situação: quem já pediu e quem falta |

### `relatorios/` — saídas

| Arquivo | Formato | Conteúdo |
|---|---|---|
| `ExcelRelatorio.gs` | `.xlsx` | Pedidos do mês agrupados por pregão |
| `RelatorioPregoes.gs` | `.xlsx` / `.docx` | Pregões ativos ou inativos e seus produtos |
| `RelatorioPersonalizado.gs` | `.docx` | Visão administrativa por unidade |
| `PdfResumo.gs` | `.pdf` | Exportação da aba `Resumo` |
| `GuiaSistemaPdf.gs` | `.pdf` | Guia de treinamento |

### `manutencao/` — operação e suporte

| Arquivo | Responsabilidade |
|---|---|
| `Setup.gs` | Criação das abas e instalação dos gatilhos |
| `SetupUnidades.gs` | Importação inicial e criação das planilhas das unidades |
| `ResetRodada.gs` | Nova rodada, com processamento em blocos |
| `Reparo.gs` | Reparo em etapas, em segundo plano |
| `RepararUnidade.gs` | Leitura flexível e reparo de uma unidade específica |
| `PullDiagnostico.gs` | Puxada manual com diagnóstico detalhado |
| `EstoqueUnidades.gs` | Funções legadas do campo Estoque Atual (desativadas) |
| `TesteFluxo.gs` | Teste de ponta a ponta, desligado por segurança |

### `interface/` — camada visual

| Arquivo | Responsabilidade |
|---|---|
| `VisualUnidades.gs` | Aplicação do tema nas planilhas das unidades |
| `Ajuda.html` | Janela modal de ajuda |

### `dados/` — dados de demonstração

| Arquivo | Responsabilidade |
|---|---|
| `DadosIniciais.gs` | `DADOS_INICIAIS`: unidades, IDs e catálogo semente (fictícios) |

Também define a **ordem oficial das unidades** nas abas `Resumo` e
`Pedidos das Unidades`.

## 4. Dependências entre módulos

```
Config.gs
   ↑ (todos dependem)
Util.gs ──────────────────┐
   ↑                      │
   ├── Produtos.gs ───────┤
   ├── Unidades.gs ───────┤
   │      ↑               │
   │   DadosIniciais.gs   │
   │                      ▼
   ├── Sincronizacao.gs ──→ Dropdown.gs, VisualUnidades.gs, MesReferencia.gs
   │
   ├── Bridge.gs ─────────→ RepararUnidade.gs, Resumo.gs, TotalPorProduto.gs
   │
   ├── Consolidacao.gs ───→ Resumo.gs, TotalPorProduto.gs
   │
   └── Relatórios ────────→ Produtos.gs, Resumo.gs, Consolidados
```

O Apps Script carrega todos os arquivos do projeto em um **único escopo global**.
Não há `import` ou `require`: qualquer função é visível para qualquer outra. A ordem
dos arquivos não importa, porque declarações `function` e `var` são içadas antes da
execução.

Por convenção, funções com sufixo `_` são internas; funções com prefixo `menu` são
pontos de entrada chamados pelo menu ou executados manualmente no editor.

## 5. Decisões de projeto

**`Consolidados` como fonte de verdade.** `Resumo` e `Pedidos das Unidades` são
visões derivadas e descartáveis. Isso permite reconstruir as duas abas sem consultar
as unidades novamente (`reconstruirResumoDeConsolidadosExistentes_`).

**Processamento em blocos.** O Apps Script encerra qualquer execução em 6 minutos.
Operações que percorrem todas as unidades gravam o progresso em `PropertiesService`
e agendam a continuação com `ScriptApp.newTrigger().timeBased().after(...)`.

**Debounce na sincronização.** Marcar vários itens como `Enviado` em sequência
dispararia uma sincronização por edição. `CONFIG.SYNC_DEBOUNCE_MS` e a fila em
`PropertiesService` agrupam essas edições em uma execução só.

**Leitura flexível de cabeçalho.** `detectarColunasPedido_` e
`obterMapaColunasProdutos_` localizam as colunas pelo texto do cabeçalho em vez de
posições fixas. O sistema continua lendo planilhas com layout antigo.

**Distribuição manual do catálogo.** A propagação do catálogo é sempre disparada pelo
operador, nunca por gatilho automático. Evita que uma edição acidental no catálogo
reescreva as planilhas de todas as unidades no meio de uma rodada.
