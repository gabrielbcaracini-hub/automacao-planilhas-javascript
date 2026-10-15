# Interface e dados

Camada visual, janela de ajuda, dados de demonstração e manifesto do projeto.

---

## `VisualUnidades.gs`

**Caminho:** `src/interface/VisualUnidades.gs`
**Funções:** 12

### Objetivo

Aplicar o tema visual nas planilhas das unidades, usando o mesmo modelo da central.

A formatação é aplicada **sob demanda**, nunca no caminho automático de
sincronização. Formatar células é a operação mais cara do `SpreadsheetApp`; fazer isso
a cada edição numa unidade tornaria a sincronização inviável.

### Funções

#### `menuAplicarVisualTodasUnidades()`

Ponto de entrada. Aplica o tema em todas as unidades ativas.

#### `aplicarVisualTodasUnidades_()`

Percorre as unidades chamando `aplicarVisualPedidoUnidade_`. Cada unidade é tratada
em `try/catch` próprio.

#### `aplicarVisualPedidoUnidade_(aba)`

| Parâmetros | `aba` — `Sheet` da aba `Pedido` |
|---|---|

Aplica o tema completo: cabeçalho, colunas destacadas, faixas de categoria e
alternância de linhas.

#### Formatação por coluna

| Função | Coluna | Destaque |
|---|---|---|
| `aplicarVisualColunaQuantidade_` | Quantidade | Âmbar — é o que a unidade preenche |
| `aplicarVisualColunaStatus_` | Status | Índigo no cabeçalho |
| `aplicarVisualColunaObservacao_` | Observação | Padrão |

As cores vêm de `CONFIG.TEMA`.

#### `aplicarFormatacaoCondicionalStatus_(aba, dataStart, nDataRows, colStatus)`

Formatação condicional por valor do status:

| Status | Cor |
|---|---|
| `Rascunho` | Cinza |
| `Enviado` | Verde |
| `Consolidado` | Azul |

Dá ao operador da unidade um retorno visual imediato sobre o que já foi enviado.

#### `aplicarVisualCategoriasPedido_(aba, linhaInicio, linhaFim, t)`

Formata as linhas de título de categoria como faixas.

#### `corrigirVisualFaixasCategoria_(aba)`

Reparo das faixas quando ficam desalinhadas após edições manuais.

#### `reforcarCabecalhoPedidoGeral_(aba, headerLinha, t)`

Reaplica o estilo do cabeçalho.

#### `aplicarVisualLinhasPedidoAposCatalogo_(aba, headerLinha, linhaFimCatalogo)`

Formata as linhas após o catálogo, usada quando o catálogo é reescrito.

### Abas relacionadas

`Pedido` das unidades (escrita de formatação)

### Dependências

`Config.gs`, `Util.gs`, `Unidades.gs`

### Execução

Manual, por `menuAplicarVisualTodasUnidades`. Também é chamada pela sincronização de
catálogo, salvo quando as opções pedem para pular o visual.

### Limitações

- É a operação mais cara do sistema. Com muitas unidades, pode exceder o limite de
  execução — rode em etapas se necessário.
- Formatações personalizadas feitas pela unidade são sobrescritas.
- A formatação condicional é recriada por completo a cada aplicação.

---

## `Ajuda.html`

**Caminho:** `src/interface/Ajuda.html`
**Tipo:** arquivo HTML do Apps Script

### Objetivo

Janela modal de ajuda, aberta pelo menu **❓ Ajuda**.

### Conteúdo

Página HTML autocontida, com CSS embutido, explicando:

- para que serve cada aba;
- o fluxo de uso do sistema;
- o que cada item do menu faz;
- as dúvidas mais comuns.

Não há JavaScript do lado do cliente: é conteúdo estático.

### Como é carregado

```javascript
function menuAjuda() {
  var html = HtmlService.createHtmlOutputFromFile('Ajuda')
    .setWidth(760)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Como usar o sistema');
}
```

> O arquivo precisa se chamar exatamente **`Ajuda`** no projeto Apps Script. O nome é
> passado como string e não é validado em tempo de compilação: um nome diferente só
> falha quando o usuário clica no menu.

### Histórico

Substituiu um `ui.alert` com texto corrido, que era difícil de ler. A janela modal
permite estrutura, hierarquia visual e títulos.

### Dependências

`Menu.gs`

### Limitações

- O conteúdo está escrito no HTML: mudanças no sistema exigem editar o arquivo.
- O tamanho da janela é fixo em 760×640.
- Caixas de diálogo modais do Apps Script não são redimensionáveis pelo usuário.

---

## `DadosIniciais.gs`

**Caminho:** `src/dados/DadosIniciais.gs`
**Funções:** nenhuma — só o objeto `DADOS_INICIAIS`

### Objetivo

Dados de demonstração para o setup inicial.

> **Todo o conteúdo deste arquivo é fictício.** No sistema original continha os dados
> reais. Substitua pelo seu conteúdo antes de usar.

### Estrutura

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
| `unidades` | Lista das unidades **e a ordem oficial** nas abas derivadas |
| `produtos` | Catálogo semente, gravado quando a aba `Produtos` está vazia |

### A ordem importa

`DADOS_INICIAIS.unidades` define a ordem das colunas na aba `Resumo` e das linhas em
`Pedidos das Unidades`, por meio de `ordenarNomesUnidadesSituacao_`.

Unidades que não constem nessa lista vão para o fim, em ordem alfabética.

### Quem usa

| Módulo | Uso |
|---|---|
| `Unidades.gs` | Fallback quando a aba `Unidades` está vazia |
| `SetupUnidades.gs` | Fonte da importação inicial |
| `TotalPorProduto.gs` | Ordem oficial das unidades |
| `Resumo.gs` | Ordem das colunas |
| `MesReferencia.gs` | Lista de unidades a migrar |
| `ResetRodada.gs` | Lista de unidades a limpar |

### Catálogo fictício incluído

24 produtos, cobrindo os casos relevantes para entender o sistema:

| Categoria | Pregões |
|---|---|
| `PAPELARIA` | 10/2025, 63/2026 |
| `MATERIAL DE EXPEDIENTE` | 22/2025 |
| `GÊNERO ALIMENTÍCIO` | 35/2025, 49/2026 |
| `GÊNERO HORTIFRUTI` | 41/2025 |
| `MATERIAL DE LIMPEZA` | 58/2025, 01/2026 |
| `MATERIAL DESCARTÁVEL` | 58/2025 |
| `INFORMÁTICA` | 77/2025, 63/2026 |
| `ELETRODOMÉSTICOS` | **05/2025 e 156/2025** |
| `ARES CONDICIONADOS` | 14/2025 |
| `FERRAMENTAS / JARDINAGEM` | 90/2025 |
| `COMPRA DIRETA` | sem pregão |

> `ELETRODOMÉSTICOS` tem **dois números de pregão de propósito**: é o caso que a
> inserção posicionada de linhas precisa tratar. Veja
> [`../funcionalidades.md`](../funcionalidades.md#posicionamento-quando-a-categoria-tem-vários-pregões).

A categoria `COMPRA DIRETA` cobre o caso de item sem pregão formal.

### Limitações

- Os IDs vêm vazios: o setup cria as planilhas. Para reaproveitar planilhas
  existentes, preencha antes de executar.
- O catálogo semente só é usado quando a aba `Produtos` está vazia.

---

## `appsscript.json`

**Caminho:** `src/appsscript.json`
**Tipo:** manifesto do projeto

### Objetivo

Declarar o runtime, o fuso horário, o registro de exceções e os escopos OAuth.

### Conteúdo

```json
{
  "timeZone": "America/Sao_Paulo",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/script.container.ui",
    "https://www.googleapis.com/auth/script.external_request"
  ],
  "executionApi": { "access": "MYSELF" }
}
```

### Campos

| Campo | Valor | Significado |
|---|---|---|
| `timeZone` | `America/Sao_Paulo` | Fuso das datas e do gatilho diário |
| `runtimeVersion` | `V8` | Runtime moderno, com sintaxe ECMAScript atual |
| `exceptionLogging` | `STACKDRIVER` | Erros registrados no Cloud Logging |
| `dependencies` | `{}` | Sem bibliotecas externas |
| `executionApi.access` | `MYSELF` | Só o dono pode chamar via API |

### Escopos

Detalhados em [`../configuracao.md`](../configuracao.md#3-escopos-oauth).

Declarar os escopos explicitamente, em vez de deixar o Apps Script inferi-los, dá
controle sobre o que é pedido ao usuário na tela de autorização.

### Como editar

No editor do Apps Script, o manifesto fica oculto por padrão. Ative em
**Configurações do projeto → Mostrar o arquivo de manifesto appsscript.json no editor**.

Com `clasp`, o arquivo é versionado normalmente e enviado pelo `clasp push`.

### Limitações

- Alterar escopos obriga o usuário a autorizar de novo na próxima execução.
- `runtimeVersion: V8` é incompatível com código escrito para o runtime antigo
  (Rhino). Este projeto já é V8.
- Mudar o `timeZone` afeta as datas gravadas e o horário do gatilho diário.
