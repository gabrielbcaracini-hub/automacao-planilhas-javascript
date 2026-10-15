# Documentação dos scripts

Documentação módulo a módulo. Cada página descreve, para cada arquivo: objetivo,
principais funções, parâmetros, entradas e saídas, abas relacionadas, dependências,
como executar e limitações conhecidas.

## Índice

| Página | Módulos cobertos |
|---|---|
| [`core.md`](core.md) | `Config.gs`, `Menu.gs`, `Util.gs`, `Unidades.gs` |
| [`catalogo.md`](catalogo.md) | `Produtos.gs`, `Categorias.gs`, `Dropdown.gs`, `NovosPregoesMatriz.gs` |
| [`sincronizacao.md`](sincronizacao.md) | `Sincronizacao.gs`, `Bridge.gs`, `Consolidacao.gs`, `PedidoUnidade.gs`, `MesReferencia.gs` |
| [`consolidacao.md`](consolidacao.md) | `Resumo.gs`, `TotalPorProduto.gs` |
| [`relatorios.md`](relatorios.md) | `ExcelRelatorio.gs`, `RelatorioPregoes.gs`, `RelatorioPersonalizado.gs`, `PdfResumo.gs`, `GuiaSistemaPdf.gs` |
| [`manutencao.md`](manutencao.md) | `Setup.gs`, `SetupUnidades.gs`, `ResetRodada.gs`, `Reparo.gs`, `RepararUnidade.gs`, `PullDiagnostico.gs`, `TesteFluxo.gs`, `EstoqueUnidades.gs` |
| [`interface-e-dados.md`](interface-e-dados.md) | `VisualUnidades.gs`, `Ajuda.html`, `DadosIniciais.gs`, `appsscript.json` |

## Convenções do código

Vale para todos os módulos.

### Sufixo `_` — função interna

```javascript
function obterDisponibilidadeCategorias_() { }
```

Uso interno. **Não aparecem na lista de funções executáveis** do editor do Apps
Script, o que impede que sejam disparadas por engano.

### Prefixo `menu` — ponto de entrada

```javascript
function menuGerarRelatorioExcel() { }
```

Chamadas pelo menu da planilha ou executadas manualmente no editor. Tratam exceções
e exibem o resultado ao usuário com `ui.alert`.

### Escopo global único

O Apps Script carrega todos os arquivos do projeto em um escopo só. Não há `import`
nem `require`, e a ordem dos arquivos não importa: declarações `function` e `var` são
içadas antes da execução.

Na prática, **todo módulo pode chamar qualquer função de qualquer outro**. As
dependências listadas nestas páginas são as relações reais de uso, não imports.

### Padrão de tratamento de erro

Pontos de entrada seguem sempre a mesma forma:

```javascript
function menuAlgumaCoisa() {
  var ui = SpreadsheetApp.getUi();
  try {
    // trabalho
    ui.alert('Título', 'Mensagem de sucesso', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}
```

Operações sobre várias unidades isolam cada iteração em `try/catch` próprio e
acumulam as falhas para reportar ao final.

### Tipos nos comentários

O código usa anotações JSDoc:

```javascript
/**
 * @param {Sheet} aba
 * @param {number} linha
 * @return {Object} colunas 0-based
 */
```

Tipos comuns: `Spreadsheet`, `Sheet`, `Range` (objetos do `SpreadsheetApp`).

### Índices de coluna

Há duas convenções no código, e a distinção importa:

| Contexto | Base | Exemplo |
|---|---|---|
| `CONFIG.PRODUTOS.*` e chamadas a `getRange` | **1-based** | `CODIGO: 1` |
| Mapas retornados por `obterMapaColunasProdutos_` | **0-based** | `map.colProduto` |

Daí o padrão `map.colProduto + 1` ao montar um `getRange`.
