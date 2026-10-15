# Instalação

Guia para montar o sistema em um **ambiente de testes próprio**.

> **Antes de começar.** Use planilhas novas e vazias. Não aponte este código para
> planilhas que contenham dados reais antes de validar o fluxo inteiro em um ambiente
> isolado. Várias funções reescrevem abas por completo.

## Pré-requisitos

| Requisito | Observação |
|---|---|
| Conta Google | Com acesso ao Google Sheets, Docs e Drive |
| Espaço no Drive | Os relatórios gerados consomem a cota da conta que executa |
| Node.js 16+ | Só para a Opção A (clasp) |

## Opção A — com clasp

O [clasp](https://github.com/google/clasp) é a CLI oficial do Google para projetos
Apps Script. Permite versionar o código em Git e publicar por linha de comando.

### 1. Instalar e autenticar

```bash
npm install -g @google/clasp
clasp login
```

O navegador abre para autorizar a conta.

### 2. Habilitar a API do Apps Script

Acesse https://script.google.com/home/usersettings e ative **Google Apps Script API**.
Sem isso o `clasp push` falha.

### 3. Criar a planilha central e o projeto

```bash
clasp create --type sheets --title "Sistema de Compras" --rootDir ./src
```

Esse comando cria a planilha, o projeto Apps Script vinculado e grava o `.clasp.json`.

Se a planilha já existir, pegue o `scriptId` em **Extensões → Apps Script →
Configurações do projeto** e crie o arquivo manualmente:

```bash
cp .clasp.json.example .clasp.json
```

```json
{
  "scriptId": "SEU_SCRIPT_ID",
  "rootDir": "src"
}
```

> `.clasp.json` está no `.gitignore`. Ele contém o ID do seu projeto e não deve ser
> versionado.

### 4. Enviar o código

```bash
clasp push
```

Confirme a substituição do manifesto quando solicitado.

```bash
clasp open
```

### 5. Autorizar os escopos

No editor do Apps Script, selecione a função `setupPlanilhaMae` e clique em **Executar**.
O Google pede autorização na primeira execução.

A tela de aviso "O Google não verificou este app" é esperada em projetos pessoais:
clique em **Avançado → Acessar (nome do projeto)**.

## Opção B — cópia manual

Sem Node.js nem linha de comando.

1. Crie uma planilha no Google Sheets. Ela será a planilha central.
2. Abra **Extensões → Apps Script**.
3. Para cada arquivo `.gs` em `src/`, crie um arquivo de script no editor e cole o
   conteúdo. Mantenha o mesmo nome (a divisão em pastas é opcional).
4. Crie um arquivo **HTML** chamado exatamente `Ajuda` e cole o conteúdo de
   `src/interface/Ajuda.html`.

   > O nome precisa ser `Ajuda`: `menuAjuda()` chama
   > `HtmlService.createHtmlOutputFromFile('Ajuda')`.

5. Em **Configurações do projeto**, marque **Mostrar o arquivo de manifesto
   appsscript.json no editor** e substitua o conteúdo pelo de `src/appsscript.json`.
6. Salve.

### Ordem dos arquivos

Não importa. O Apps Script carrega tudo em um escopo global único e as declarações
`function` e `var` são içadas antes da execução.

## Configuração mínima

Abra `src/core/Config.gs` e ajuste:

```javascript
SPREADSHEET_MAE_ID: '',   // vazio = usa a planilha ativa

ORGANIZACAO: {
  NOME: 'MINHA ORGANIZAÇÃO',
  LOGO_URL: ''
},
```

Deixar `SPREADSHEET_MAE_ID` vazio funciona enquanto o código roda na própria planilha
central. **Preencha com o ID real** antes de instalar os gatilhos das unidades: os
gatilhos executam fora da central e precisam saber para onde escrever.

O ID está na URL:

```
https://docs.google.com/spreadsheets/d/ESTE_TRECHO_É_O_ID/edit
```

Depois ajuste `src/dados/DadosIniciais.gs` com os nomes das suas unidades.

## Preparação do ambiente

Execute, nesta ordem, no editor do Apps Script:

### 1. `setupPlanilhaMae`

Cria as abas `Produtos`, `Categorias`, `Unidades`, `Consolidados`, `Resumo`,
`Pedidos das Unidades` e `Log` com os cabeçalhos corretos.

### 2. `menuImportarUnidadesCompleto`

- grava o catálogo semente de `DADOS_INICIAIS.produtos` na aba `Produtos`;
- cria uma planilha no Drive para cada unidade de `DADOS_INICIAIS.unidades`;
- registra nome e ID de cada uma na aba `Unidades`;
- monta a aba `Pedido` de cada planilha criada.

> Se as planilhas das unidades já existirem, preencha os IDs em
> `DADOS_INICIAIS.idsPlanilhas` antes de executar. O sistema reaproveita o que
> encontrar em vez de criar duplicatas.

Essa função percorre todas as unidades e é a mais sujeita ao limite de 6 minutos.
Com muitas unidades, pode ser necessário executá-la mais de uma vez — ela retoma de
onde parou.

### 3. `instalarGatilhos`

Registra os gatilhos do projeto:

| Gatilho | Evento | Função |
|---|---|---|
| Abertura da central | `onOpen` | Menu e atualização da situação |
| Consolidação diária | Temporizado | `consolidarPedidosAgendado` |
| Abertura das unidades | `onOpen` (por planilha) | `onOpenUnidadeInstalavel_` |

Confira o resultado com `menuListarGatilhos`.

### 4. `menuConectarTudoAgora`

Instala o gatilho `onEdit` em cada planilha de unidade e faz a primeira puxada dos
pedidos já marcados como `Enviado`.

### 5. Verificação

Recarregue a planilha central (F5). O menu **Sistema de Compras** deve aparecer.

| Verificação | Resultado esperado |
|---|---|
| Menu visível | **Sistema de Compras** na barra superior |
| Aba `Produtos` | Catálogo semente carregado |
| Aba `Unidades` | Uma linha por unidade, com ID preenchido |
| Planilhas no Drive | Uma planilha por unidade, com aba `Pedido` |
| `menuListarGatilhos` | Gatilhos registrados |
| **❓ Ajuda** | Janela modal abre |

## Compartilhamento

| Perfil | Precisa de acesso a |
|---|---|
| Administrador | Planilha central **e** todas as planilhas das unidades (editor) |
| Operador de unidade | Apenas a planilha da própria unidade (editor) |

O operador de uma unidade não precisa de acesso à central nem às planilhas das outras
unidades. A sincronização roda com as credenciais de quem instalou os gatilhos, não
com as de quem edita.

## Problemas comuns

| Sintoma | Causa provável | Solução |
|---|---|---|
| `clasp push` falha com erro de API | API do Apps Script desabilitada | Ative em script.google.com/home/usersettings |
| Menu não aparece | `onOpen` falhou | Recarregue; veja os logs em **Execuções** |
| "O Google não verificou este app" | Projeto pessoal sem verificação | **Avançado → Acessar** |
| "Exceeded maximum execution time" | Limite de 6 minutos | Execute de novo: as funções longas retomam de onde pararam |
| Unidade não sincroniza | Gatilho `onEdit` ausente | `menuConectarTudoAgora` |
| `Não encontrei a aba X` | Setup incompleto | Execute `setupPlanilhaMae` |
| Erro ao abrir planilha da unidade | ID errado ou sem permissão | `menuCorrigirIdsUnidades` |
| Ajuda não abre | Arquivo HTML com nome errado | Renomeie para `Ajuda` |

## Desinstalação

```
removerGatilhosDoProjeto      → remove os gatilhos da central
removerGatilhosUnidades       → remove os gatilhos das unidades
removerGatilhosOnOpenUnidades_→ remove os gatilhos de abertura das unidades
```

As planilhas e os dados permanecem; apenas a automação é desligada.
