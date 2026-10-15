# Módulos core

Base do sistema: configuração, menu, utilitários e cadastro de unidades.

---

## `Config.gs`

**Caminho:** `src/core/Config.gs`
**Funções:** nenhuma — só o objeto `CONFIG`

### Objetivo

Centralizar toda a parametrização. É o único arquivo que precisa ser editado para
adaptar o sistema a outro ambiente.

### Estrutura

| Chave | Conteúdo |
|---|---|
| `SPREADSHEET_MAE_ID` | ID da planilha central. Vazio = usa a planilha ativa |
| `ORGANIZACAO` | `NOME` e `LOGO_URL` usados nos relatórios |
| `ABAS` | Nome de cada aba, incluindo as legadas |
| `PRODUTOS`, `CATEGORIAS`, `UNIDADES`, `PEDIDO`, `CONSOLIDADOS`, `RESUMO`, `TOTAL_PRODUTO`, `LOG` | Posição das colunas (1-based) e cabeçalhos |
| `VALORES` | Textos padronizados de `Ativo` e `Status` |
| `SYNC_DEBOUNCE_MS` | Intervalo mínimo entre sincronizações automáticas |
| `CONSOLIDACAO_HORA` | Hora do gatilho diário |
| `LINHA_DADOS` | Primeira linha de dados |
| `PEDIDO_LINHAS_VALIDACAO` | Linhas que recebem validação na aba `Pedido` |
| `TEMA` | Paleta de cores da camada visual |

### Dependências

Nenhuma. Todos os outros módulos dependem dele.

### Limitações

Carregado como `var` global. Alterações exigem salvar o projeto e recarregar a
planilha para que o `onOpen` rode de novo.

Detalhamento completo em [`../configuracao.md`](../configuracao.md).

---

## `Menu.gs`

**Caminho:** `src/core/Menu.gs`
**Funções:** 4

### Objetivo

Montar o menu da planilha central e executar as rotinas de abertura.

### Funções

#### `onOpen()`

Gatilho simples, disparado ao abrir a planilha.

| Aspecto | Descrição |
|---|---|
| Parâmetros | Nenhum |
| Retorno | Nenhum |
| Entradas | A planilha ativa |
| Saídas | Menu na interface; abas atualizadas |

Sequência:

1. Detecta se a planilha aberta é de uma unidade; se for, desvia para `onOpenUnidade_` e encerra.
2. Monta o menu **Sistema de Compras**.
3. Aplica a visibilidade das abas (`aplicarVisibilidadeAbasCliente_`).
4. Remove gatilhos de catálogo automático de versões antigas.
5. Atualiza a aba `Pedidos das Unidades`.
6. Restaura o visual do `Resumo`, se houver migração pendente.
7. Importa o lote de novos pregões, se ainda não importado.

Cada etapa tem `try/catch` próprio: uma falha isolada não impede o menu de aparecer.

#### `menuAjuda()`

Abre `Ajuda.html` em janela modal de 760×640.

> O arquivo HTML precisa se chamar exatamente **`Ajuda`** no projeto Apps Script:
> a chamada é `HtmlService.createHtmlOutputFromFile('Ajuda')`.

#### `menuInstalarMenuUnidades()`

Instala o gatilho `onOpen` nas planilhas das unidades, para que elas também tenham
menu. Delega para `instalarGatilhoOnOpenUnidades_`.

Retorna ao usuário quantas unidades receberam o menu e quais falharam. As unidades
precisam recarregar a planilha (F5) para ver o resultado.

#### `restaurarResumoFormatadoSeArquivado_()`

Rotina de migração pontual. Detecta o caso em que a aba `Resumo` foi recriada sem a
formatação e devolve a versão arquivada que mantinha as faixas e larguras.

Controlada por uma flag em `DocumentProperties` para rodar **uma vez só**.

### Dependências

`Config.gs`, `Util.gs`, `PedidoUnidade.gs`, `TotalPorProduto.gs`, `Categorias.gs`,
`NovosPregoesMatriz.gs`

### Execução

Automático ao abrir a planilha. `menuInstalarMenuUnidades` roda pelo editor.

### Limitações

- `onOpen` é gatilho **simples**: não pode abrir outras planilhas nem chamar serviços
  que exijam autorização. As funções que precisam disso são chamadas por gatilhos
  instaláveis.
- Se o `onOpen` demorar, a planilha abre antes de o menu aparecer.

---

## `Util.gs`

**Caminho:** `src/core/Util.gs`
**Funções:** 49

### Objetivo

Funções compartilhadas por todo o sistema: acesso a planilhas e abas, normalização de
texto, locks, log e controle de visibilidade de abas.

É o módulo mais usado — praticamente toda rotina depende de alguma função daqui.

### Funções principais

#### `obterPlanilhaMae()`

| Aspecto | Descrição |
|---|---|
| Parâmetros | Nenhum |
| Retorno | `Spreadsheet` |

Prefere a planilha **já aberta** e só recorre a `openById` quando o código roda fora
da central (por exemplo, num gatilho instalado numa unidade). Evita o timeout que
`openById` causa ao abrir a própria planilha ativa.

Se `CONFIG.SPREADSHEET_MAE_ID` estiver vazio e o código rodar fora da central, falha.

#### `obterAba(ss, nome)` / `obterAbaObrigatoria(ss, nome)`

| Função | Comportamento quando a aba não existe |
|---|---|
| `obterAba` | Retorna `null` |
| `obterAbaObrigatoria` | Lança exceção com mensagem descritiva |

#### `garantirAbaComCabecalho(ss, nome, header)`

Cria a aba se não existir e garante o cabeçalho. Idempotente: pode ser chamada
quantas vezes for preciso.

#### `normalizarTexto(valor)`

| Aspecto | Descrição |
|---|---|
| Parâmetros | `valor` — qualquer tipo |
| Retorno | `string` sem acento, sem espaço extra, em caixa única |

Base de **todas as comparações de texto** do sistema. Permite que
`'GÊNERO ALIMENTÍCIO'` e `'genero alimenticio '` sejam tratados como iguais.

#### `ehAtivo(valor)`

Interpreta a coluna `Ativo` de forma tolerante: aceita `Sim`, `SIM`, `sim`, `S`,
`true`, `1`.

#### `adquirirLock(nome, timeoutSegundos)`

Envolve o `LockService`. Impede que a mesma operação rode duas vezes ao mesmo tempo —
por exemplo, o gatilho diário disparando enquanto alguém puxa os pedidos pelo menu.

#### `abrirPlanilhaFilha(spreadsheetId, nomeUnidade)` / `abrirPlanilhaFilhaRapido(...)`

| Função | Comportamento |
|---|---|
| `abrirPlanilhaFilha` | Abre por ID; se falhar, procura pelo nome no Drive e corrige o cadastro |
| `abrirPlanilhaFilhaRapido` | Abre por ID e falha direto, sem fallback |

A versão rápida existe para laços sobre muitas unidades, onde a busca no Drive seria
cara demais.

#### `registrarLog(acao, detalhe)`

Acrescenta uma linha na aba `Log` com data/hora, ação e detalhe. Falha em silêncio se
a aba não existir — log nunca deve quebrar uma operação.

#### `aplicarVisibilidadeAbasCliente_(ssMae)`

Oculta as abas administrativas (`Consolidados`, `Log`) e mantém visíveis as
operacionais. Chamada no `onOpen`.

#### `obterNomeOrganizacao_()`

Devolve `CONFIG.ORGANIZACAO.NOME` com fallback para `'ORGANIZAÇÃO'`. Usada no
cabeçalho dos relatórios.

### Outras funções

| Grupo | Funções |
|---|---|
| Datas | `formatarData`, `serialNumericoDeDataSheets_`, `valorCelulaPreservandoZero_` |
| Cabeçalhos | `obterHeaderPedidoUnidade_`, `obterHeaderResumo_`, `normalizarCabecalhoPedido_`, `garantirCabecalhoPedidoUnidade_` |
| Linhas | `montarLinhaPedidoUnidade_`, `montarLinhaCategoriaPedido_`, `ehLinhaCategoriaPedido_` |
| Abas legadas | `removerAbasLegadasSistema_`, `ocultarAbaListaPedidos_`, `executarRemocaoAbaPorNome_` |
| Proteções | `liberarProtecoesAba_` |
| Leitura | `lerDados`, `extrairNumeroQuantidade`, `limparDadosAba` |

### Dependências

`Config.gs`. Nada mais.

### Limitações

- `normalizarTexto` remove acentos: textos que só diferem por acento passam a ser
  tratados como iguais.
- `extrairNumeroQuantidade` assume números no formato brasileiro (vírgula decimal).
- As funções de remoção de abas são destrutivas e não pedem confirmação.

---

## `Unidades.gs`

**Caminho:** `src/core/Unidades.gs`
**Funções:** 2

### Objetivo

Ler o cadastro de unidades.

### Funções

#### `listarTodasUnidades(ssOpt)`

| Aspecto | Descrição |
|---|---|
| Parâmetros | `ssOpt` — `Spreadsheet` opcional; sem ele, usa a central |
| Retorno | `Array.<{nome, spreadsheetId, ativo}>` |
| Entrada | Aba `Unidades` |

Se a aba estiver vazia ou não existir, cai para `DADOS_INICIAIS.unidades` e
`DADOS_INICIAIS.idsPlanilhas`. Isso permite que o sistema funcione logo após a
instalação, antes do primeiro setup.

#### `listarUnidadesAtivas(ssOpt)`

Mesmo retorno, filtrado por `ativo`. É a função usada pelas rotinas de sincronização.

### Dependências

`Config.gs`, `Util.gs`, `DadosIniciais.gs`

### Limitações

- Uma unidade com ID vazio é retornada assim mesmo; quem consome precisa tratar.
- O nome da unidade é a chave de ligação com `Consolidados` e `Resumo`. Renomear uma
  unidade na aba `Unidades` sem atualizar os dados já consolidados quebra o vínculo.
