# Segurança e anonimização

Relatório do que foi removido do código antes da publicação, por que, e como validar
que nada sensível permaneceu.

## Contexto

O código deste repositório vem de um sistema em produção, usado para apoiar um
processo administrativo real. O código em si — a lógica de automação — é o objeto do
portfólio. Os **dados** que acompanhavam esse código não são, e foram substituídos.

Nenhum arquivo original foi alterado durante a análise. O repositório foi montado como
uma cópia transformada.

## O que foi removido

| Categoria | Ocorrências | Tratamento |
|---|---|---|
| ID da planilha central | 1 | Substituído por string vazia em `CONFIG.SPREADSHEET_MAE_ID` |
| IDs das planilhas de unidade | 14 | Substituídos por strings vazias em `DADOS_INICIAIS.idsPlanilhas` |
| IDs dos projetos Apps Script | 2 | Removidos |
| Nomes reais das unidades | 14 | Substituídos por nomes genéricos |
| Nome do município | 6 ocorrências | Substituído por `CONFIG.ORGANIZACAO.NOME` |
| URL do logotipo institucional | 1 | Substituída por `CONFIG.ORGANIZACAO.LOGO_URL` |
| Catálogo real de produtos | 147 itens | Substituído por 24 itens fictícios |
| Lote de pregões reais | 1.645 linhas | Substituído por 10 itens fictícios |
| Razões sociais de fornecedores | em todos os registros acima | Substituídas por "FORNECEDOR EXEMPLO A…N" |
| Números e datas de pregões reais | em todos os registros acima | Substituídos por valores fictícios |

### O que a varredura **não** encontrou

Verificado por expressão regular em todos os arquivos:

- nenhum CPF ou CNPJ;
- nenhum telefone;
- nenhum endereço de e-mail;
- nenhuma credencial, token, chave de API ou segredo de cliente OAuth;
- nenhum dado pessoal de cidadãos ou de servidores;
- nenhum link privado do Google Drive.

O sistema, por desenho, nunca manipulou dados pessoais: ele trabalha com produtos,
quantidades e unidades administrativas. Os dados de identificação que existiam eram
**de infraestrutura** (IDs de planilhas) e **de contratos públicos** (números de
pregão e fornecedores).

## Como as credenciais são tratadas

O sistema **não armazena credenciais em lugar nenhum do código**. A autenticação é
inteiramente delegada ao Google:

- o Apps Script executa com a identidade da conta Google que autorizou o projeto;
- as permissões são declaradas como escopos OAuth em `appsscript.json` e concedidas
  na primeira execução;
- não há chave de API, token, senha ou arquivo de conta de serviço no projeto.

O `.gitignore` bloqueia preventivamente os arquivos que costumam carregar segredos em
projetos Apps Script:

```
.clasp.json          ID do projeto, específico de cada instalação
.clasprc.json        credenciais OAuth do clasp
client_secret*.json  segredo de cliente OAuth
credentials.json
token.json
service-account*.json
.env
```

> `.clasprc.json` fica no diretório home do usuário, não no repositório. Ele está
> listado por precaução: **esse arquivo contém um refresh token válido** e nunca deve
> ser versionado.

## Substituições aplicadas

### Identificadores renomeados

Nove identificadores do código original carregavam a sigla da secretaria ou o nome de
uma unidade específica. Todos foram renomeados para termos genéricos, por substituição
mecânica de texto, sem alteração de lógica:

| Publicado | Natureza do rename |
|---|---|
| `DADOS_INICIAIS` | Constante de dados |
| `DadosIniciais.gs` | Nome de arquivo |
| `SetupUnidades.gs` | Nome de arquivo |
| `menuImportarUnidadesCompleto` | Ponto de entrada |
| `importarUnidadesCompleto` | Função interna |
| `popularProdutosUnidades` | Função interna |
| `aplicarIdsPlanilhasUnidades` | Função interna |
| `menuRepararUnidadeCentroEPuxar` | Ponto de entrada |
| `menuTestarUnidadeCentro` | Ponto de entrada |

Cada rename foi conferido nas definições **e** em todas as chamadas.

### Parametrização introduzida

A única mudança estrutural: dados institucionais que estavam embutidos no código
viraram configuração.

```javascript
ORGANIZACAO: {
  NOME: 'PREFEITURA MUNICIPAL DE EXEMPLO',
  LOGO_URL: ''
}
```

Uma função auxiliar foi adicionada em `Util.gs`:

```javascript
function obterNomeOrganizacao_() {
  var org = CONFIG.ORGANIZACAO || {};
  return String(org.NOME || '').trim() || 'ORGANIZAÇÃO';
}
```

Os três pontos que escreviam o nome do município fixo no código
(`GuiaSistemaPdf.gs`, `RelatorioPersonalizado.gs`, `RelatorioPregoes.gs`) passaram a
chamar essa função.

### Dados fictícios

| Arquivo | Conteúdo publicado |
|---|---|
| `src/dados/DadosIniciais.gs` | 14 unidades genéricas, IDs vazios, 24 produtos fictícios |
| `src/catalogo/NovosPregoesMatriz.gs` | 10 itens fictícios no lugar de 1.645 linhas reais |
| `exemplos/dados-ficticios/` | CSVs de exemplo, todos fictícios |

Os produtos fictícios preservam a **estrutura** dos dados reais — mesmas categorias
típicas, mesmo formato de texto de pregão, mesma variedade de casos. Isso mantém os
exemplos úteis para entender o sistema, incluindo o caso de uma categoria com dois
números de pregão diferentes.

## Validação

As verificações executadas sobre o repositório final:

| Verificação | Método | Resultado |
|---|---|---|
| Sintaxe JavaScript | `node --check` em todos os `.gs` | 30/30 válidos |
| Referências entre funções | Comparação entre funções definidas e chamadas, antes e depois | Nenhuma referência quebrada |
| Identificadores renomeados | Conferência de definição e chamada para cada rename | 9/9 consistentes |
| Dados sensíveis | Busca por IDs, nomes reais e termos identificadores | Nenhuma ocorrência |
| Credenciais | Busca por e-mails, tokens, chaves e segredos | Nenhuma ocorrência |
| Dados pessoais | Busca por CPF, CNPJ e telefone | Nenhuma ocorrência |

Para repetir a verificação antes de um commit, busque pelos seus próprios
identificadores:

```bash
grep -rin "SEU_ID_DE_PLANILHA\|NOME_DA_SUA_ORGANIZACAO\|NOME_DE_UNIDADE" src/
```

## Preservação do comportamento

Nenhuma lógica foi reescrita. As alterações se limitam a:

1. **substituição de literais** — IDs, nomes e dados de exemplo;
2. **renome de identificadores** — aplicado de forma consistente a definições e chamadas;
3. **parametrização** — dados institucionais movidos para `CONFIG.ORGANIZACAO`;
4. **reorganização em pastas** — sem efeito no Apps Script, que usa escopo global único.

Nenhum algoritmo, estrutura de controle ou assinatura de função foi modificado.

## Limitação desta validação

As verificações acima cobrem **sintaxe e integridade de referências**. Elas não são
testes funcionais.

O Google Apps Script só executa dentro do ambiente do Google: não há como rodar as
funções fora dele. Validar o comportamento exigiria instalar o projeto em planilhas
reais e executar o fluxo — o que não foi feito, por decisão deliberada de não tocar
em ambiente de produção.

**O código publicado não foi executado.** Antes de usá-lo, monte um ambiente de
testes conforme [`instalacao.md`](instalacao.md) e valide o fluxo com planilhas novas.

## Recomendações para quem for usar

1. **Nunca versione `.clasp.json`.** Ele carrega o ID do seu projeto.
2. **Nunca versione `.clasprc.json`.** Ele carrega um refresh token válido.
3. **Mantenha IDs fora do código.** Prefira `CONFIG` e `PropertiesService` a literais
   espalhados pelos arquivos.
4. **Revise antes de publicar.** Rode uma busca por IDs e nomes reais antes de cada
   commit em repositório público.
5. **Use uma conta de serviço dedicada** para executar os gatilhos, em vez de uma
   conta pessoal. Os gatilhos instaláveis rodam com as credenciais de quem os instalou.
