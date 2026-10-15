# Configuração

Parâmetros, permissões e gatilhos do sistema.

## 1. `CONFIG` — parametrização central

Tudo que é configurável está em [`src/core/Config.gs`](../src/core/Config.gs), no
objeto global `CONFIG`. Nenhum outro arquivo precisa ser editado para adaptar o
sistema a um novo ambiente.

### Identificação

```javascript
SPREADSHEET_MAE_ID: '',
```

ID da planilha central.

| Valor | Comportamento |
|---|---|
| Vazio | Usa a planilha ativa. Funciona só quando o código roda na própria central |
| ID preenchido | Abre a central por ID. **Obrigatório** para os gatilhos das unidades |

`obterPlanilhaMae()` prefere a planilha já aberta e só recorre a `openById` quando o
código roda fora dela — isso evita timeouts desnecessários.

```javascript
ORGANIZACAO: {
  NOME: 'PREFEITURA MUNICIPAL DE EXEMPLO',
  LOGO_URL: ''
},
```

| Campo | Uso |
|---|---|
| `NOME` | Cabeçalho dos relatórios em Word e PDF |
| `LOGO_URL` | URL pública de uma imagem. Vazio gera o relatório sem logotipo |

O `LOGO_URL` é baixado com `UrlFetchApp`, então precisa ser acessível sem
autenticação.

### Nomes das abas

```javascript
ABAS: {
  PRODUTOS: 'Produtos',
  CATEGORIAS: 'Categorias',
  UNIDADES: 'Unidades',
  CONSOLIDADOS: 'Consolidados',
  RESUMO: 'Resumo',
  TOTAL_PRODUTO: 'Pedidos das Unidades',
  LOG: 'Log',
  PEDIDO: 'Pedido',
  CATALOGO: 'Catalogo'
}
```

Renomeie aqui se quiser outros nomes nas abas. As demais entradas (`LISTA_PEDIDOS`,
`PAINEL`, `RESUMO_LEGADO`, `TOTAL_PRODUTO_LEGADO`) existem para reconhecer e limpar
abas de versões anteriores — não as remova se estiver migrando uma instalação antiga.

Algumas abas ficam ocultas para o operador (`Consolidados`, `Log`), controladas por
`aplicarVisibilidadeAbasCliente_`.

### Posição das colunas

```javascript
PRODUTOS: {
  CODIGO: 1, CATEGORIA: 2, PRODUTO: 3,
  UNIDADE_MEDIDA: 4, PREGAO: 5, ATIVO: 6,
  HEADER: ['Código', 'Categoria', 'Produto', 'Unidade', 'Pregão', 'Ativo']
}
```

Índices 1-based, como o `getRange` do Sheets espera. Existem blocos equivalentes para
`CATEGORIAS`, `UNIDADES`, `PEDIDO`, `CONSOLIDADOS`, `RESUMO`, `TOTAL_PRODUTO` e `LOG`.

> Na prática, a leitura usa `obterMapaColunasProdutos_` e `detectarColunasPedido_`,
> que localizam as colunas **pelo texto do cabeçalho**. Os índices em `CONFIG` são o
> padrão de escrita e o fallback quando a detecção não encontra o cabeçalho.

### Valores padronizados

```javascript
VALORES: {
  ATIVO_SIM: 'Sim',
  ATIVO_NAO: 'Não',
  STATUS_RASCUNHO: 'Rascunho',
  STATUS_ENVIADO: 'Enviado',
  STATUS_CONSOLIDADO: 'Consolidado'
}
```

Alimentam as validações de dados (listas suspensas). Alterar aqui muda as opções
oferecidas nas planilhas das unidades — mas exige redistribuir o catálogo para que as
validações sejam reescritas.

### Comportamento

| Parâmetro | Padrão | O que faz |
|---|---|---|
| `SYNC_DEBOUNCE_MS` | `8000` | Intervalo mínimo entre sincronizações automáticas. Agrupa edições em sequência |
| `CONSOLIDACAO_HORA` | `7` | Hora (0–23) do gatilho diário de consolidação |
| `LINHA_DADOS` | `2` | Primeira linha de dados na central |
| `PEDIDO_LINHAS_VALIDACAO` | `800` | Quantidade de linhas que recebem validação na aba `Pedido` |
| `PROP_ULTIMO_SYNC` | `'ULTIMO_SYNC_TS'` | Chave de controle do debounce em `PropertiesService` |

Reduzir `SYNC_DEBOUNCE_MS` deixa a sincronização mais responsiva, mas aumenta o
consumo de cota de execução de gatilhos.

Aumentar `PEDIDO_LINHAS_VALIDACAO` acima do necessário deixa a abertura das planilhas
das unidades mais lenta.

### Tema visual

```javascript
TEMA: {
  HEADER_BG: '#FF0000',
  COL_DESTAQUE_BG: '#FFFF00',
  COL_CATEGORIA_BG: '#00FFFF',
  ROW_A: '#0D1117',
  ROW_B: '#161B22',
  ...
}
```

Cores em hexadecimal, aplicadas tanto na central quanto nas unidades. Depois de
alterar, rode `menuAjustarVisualProdutos` (central) e
`menuAplicarVisualTodasUnidades` (unidades).

## 2. Dados iniciais

[`src/dados/DadosIniciais.gs`](../src/dados/DadosIniciais.gs) define `DADOS_INICIAIS`:

```javascript
var DADOS_INICIAIS = {
  idsPlanilhas: { 'UNIDADE CENTRO': '', ... },
  unidades:     [ 'UNIDADE CENTRO', 'UNIDADE NORTE', ... ],
  produtos:     [ { codigo, categoria, produto, pregao, ativo }, ... ]
};
```

| Campo | Papel |
|---|---|
| `idsPlanilhas` | ID de cada planilha de unidade. Vazio faz o setup criar a planilha |
| `unidades` | Lista das unidades **e a ordem oficial** das colunas no `Resumo` |
| `produtos` | Catálogo semente, gravado quando a aba `Produtos` está vazia |

> A ordem de `unidades` determina a ordem das colunas na aba `Resumo` e das linhas em
> `Pedidos das Unidades` (`ordenarNomesUnidadesSituacao_`). Unidades que não estejam
> nessa lista vão para o fim, em ordem alfabética.

O conteúdo deste arquivo é **fictício**. Substitua pelo seu antes de usar.

## 3. Escopos OAuth

Declarados em [`src/appsscript.json`](../src/appsscript.json):

```json
{
  "timeZone": "America/Sao_Paulo",
  "runtimeVersion": "V8",
  "exceptionLogging": "STACKDRIVER",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/script.container.ui",
    "https://www.googleapis.com/auth/script.external_request"
  ]
}
```

| Escopo | Por que é necessário |
|---|---|
| `spreadsheets` | Ler e escrever na central e nas unidades |
| `drive` | Criar as planilhas das unidades, salvar relatórios, localizar arquivos por nome |
| `documents` | Gerar os relatórios em Word e PDF |
| `script.scriptapp` | Criar e remover gatilhos programaticamente |
| `script.container.ui` | Menu, caixas de diálogo e janela de ajuda |
| `script.external_request` | `UrlFetchApp`: exportação `.xlsx` e download do logotipo |

Ajuste `timeZone` para o seu fuso. Ele afeta as datas gravadas nos relatórios e o
horário do gatilho de consolidação.

### Por que o escopo `drive` completo

O sistema localiza planilhas pelo nome do arquivo quando um ID está ausente ou
incorreto (`resolverIdUnidadePeloNome_`). Isso exige busca no Drive, que o escopo
restrito `drive.file` não permite. Se você preencher todos os IDs manualmente e
remover essa funcionalidade, é possível reduzir o escopo.

## 4. Gatilhos

### Instalados por `instalarGatilhos`

| Função | Tipo | Quando dispara |
|---|---|---|
| `onOpen` | Simples | Ao abrir a central |
| `consolidarPedidosAgendado` | Temporizado | Diariamente, em `CONFIG.CONSOLIDACAO_HORA` |
| `onOpenUnidadeInstalavel_` | Instalável, por planilha | Ao abrir a planilha de uma unidade |

### Instalados por `menuConectarTudoAgora`

| Função | Tipo | Quando dispara |
|---|---|---|
| `onEditPedidoUnidade` | Instalável, por planilha | A cada edição na planilha de uma unidade |

### Criados sob demanda

Gatilhos temporizados de execução única, criados e removidos automaticamente para
retomar processos longos:

| Função | Finalidade |
|---|---|
| `processarSyncPedidoPendenteRapido_` | Processa a fila de sincronização |
| `executarSyncAtrasadoUnidade_` | Reprocessa uma unidade que falhou |
| `continuarResetNovaRodada_` | Continua a nova rodada no bloco seguinte |
| `processarTotalPorProdutoPendente_` | Atualiza a situação dos pedidos |
| `processarLoteMigracaoLayoutUnidades_` | Continua a migração de layout |
| `executarReparoEmPassos` | Executa o próximo passo do reparo |

Esses gatilhos se removem ao concluir. Se um processo for interrompido, podem ficar
gatilhos órfãos: `menuListarGatilhos` mostra o que existe e
`removerGatilhosDoProjeto` limpa tudo.

> **Limite do Google:** 20 gatilhos por script por usuário. Instalar `onEdit` e
> `onOpen` em muitas unidades pode se aproximar desse teto. `menuDiagnosticarSyncAutomatico`
> ajuda a verificar o estado atual.

## 5. Permissões

| Perfil | Acesso necessário |
|---|---|
| Administrador | Editor na central **e** em todas as planilhas das unidades |
| Operador de unidade | Editor apenas na planilha da própria unidade |

### Sobre a conta que executa

Os gatilhos instaláveis rodam **com as credenciais de quem os instalou**, não de quem
edita. Consequências:

- A conta administradora precisa de acesso de edição a todas as unidades.
- Os relatórios gerados são criados no Drive dessa conta e consomem a cota dela.
- Se essa conta perder o acesso a uma unidade, a sincronização daquela unidade para.
  Reinstale com `menuConectarTudoAgora`.

### Armazenamento

Cada relatório gerado vira um arquivo no Drive da conta administradora. Em uso
contínuo, o volume cresce. As funções `obterPastaRelatoriosExportacao_`,
`obterPastaRelatoriosPregoes_` e `obterPastaPdfsResumo_` organizam os arquivos em
pastas próprias, o que facilita a limpeza periódica.

## 6. Estado persistido

O sistema guarda estado entre execuções em `PropertiesService`.

### `ScriptProperties` — todo o projeto

| Chave | Conteúdo |
|---|---|
| `ULTIMO_SYNC_TS` | Timestamp da última sincronização (debounce) |
| `SYNC_PENDENTE_*` | Fila de unidades aguardando sincronização |
| `RESET_RODADA_*` | Progresso da nova rodada entre blocos |
| `TOTAL_PRODUTO_PENDENTE` | Marca a aba de situação como desatualizada |
| `MIGRACAO_LAYOUT_*` | Progresso da migração de layout |

### `DocumentProperties` — por planilha

| Chave | Conteúdo |
|---|---|
| `RESUMO_VISUAL_RESTAURADO_*` | Marca migrações visuais já aplicadas |
| `MATRIZ_SETEMBRO_2026` | Marca que o lote de pregões já foi importado |

Essas chaves funcionam como flags de migração: garantem que uma rotina de
atualização rode **uma vez só**, mesmo que a planilha seja aberta muitas vezes.

Para forçar a reexecução de uma migração, remova a chave correspondente pelo editor
do Apps Script.
