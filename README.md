# Sistema de Solicitação de Compras — Automação de Planilhas com JavaScript

Automação completa de um processo de solicitação de compras distribuído entre uma
planilha central e várias planilhas de unidades, construída em **JavaScript (Google
Apps Script)** sobre o **Google Sheets**.

O sistema substitui a troca manual de planilhas por um fluxo automatizado: o catálogo
de produtos é distribuído de um ponto central, cada unidade registra o próprio pedido,
e os pedidos retornam consolidados para a planilha central, com relatórios gerados
automaticamente em Excel, PDF e Word.

> **Aviso sobre os dados.** Este repositório é uma versão anonimizada de um sistema que
> roda em produção. Todos os identificadores de planilhas, nomes de unidades, números
> de pregão, fornecedores e dados da instituição foram substituídos por valores
> fictícios. Nenhum dado real acompanha o código. Veja [`docs/seguranca.md`](docs/seguranca.md).

---

## Índice

- [O problema](#o-problema)
- [A solução](#a-solução)
- [Tecnologias](#tecnologias)
- [Funcionalidades](#funcionalidades)
- [Arquitetura](#arquitetura)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Utilização](#utilização)
- [Exemplos](#exemplos)
- [Resultados](#resultados)
- [Documentação completa](#documentação-completa)
- [Limitações conhecidas](#limitações-conhecidas)
- [Autor](#autor)

---

## O problema

O processo original de compras de uma estrutura administrativa com múltiplas unidades
funcionava assim:

1. A administração central mantinha uma lista de produtos disponíveis, derivada dos
   pregões vigentes.
2. Essa lista era enviada manualmente para cada unidade.
3. Cada unidade devolvia a própria planilha preenchida, em formatos e layouts diferentes.
4. Alguém consolidava tudo à mão para montar o pedido final e os relatórios.

Os pontos de falha eram claros: versões diferentes do catálogo circulando ao mesmo
tempo, retrabalho na consolidação, dificuldade em saber quais unidades já tinham
respondido, e nenhuma rastreabilidade sobre quem pediu o quê.

## A solução

Uma única base de código em Apps Script, vinculada à planilha central, que:

- **distribui** o catálogo de produtos para todas as planilhas de unidade a partir de
  um único cadastro;
- **recolhe** automaticamente os pedidos marcados como `Enviado` em cada unidade;
- **consolida** os dados em uma matriz produto × unidade, agrupada por pregão;
- **mostra** em tempo real quais unidades já pediram e quais estão pendentes;
- **gera** relatórios prontos em Excel, PDF e Word;
- **reinicia** a rodada de compras sem perder o catálogo cadastrado.

A operação inteira acontece dentro do Google Sheets, por um menu próprio. Quem usa o
sistema não precisa abrir o editor de scripts.

---

## Tecnologias

Tecnologias efetivamente utilizadas neste projeto:

| Tecnologia | Uso no projeto |
|---|---|
| **JavaScript (ECMAScript, runtime V8)** | Toda a lógica do sistema |
| **Google Apps Script** | Plataforma de execução, gatilhos e interface de menu |
| **Google Sheets (SpreadsheetApp)** | Banco de dados, interface do usuário e camada visual |
| **Google Drive (DriveApp)** | Criação das planilhas das unidades e armazenamento dos relatórios |
| **Google Docs (DocumentApp)** | Geração dos relatórios em Word e PDF |
| **HtmlService** | Janela modal de ajuda dentro da planilha |
| **PropertiesService** | Estado entre execuções, controle de debounce e retomada de processos longos |
| **LockService** | Prevenção de execuções simultâneas concorrentes |
| **ScriptApp** | Gatilhos programáticos (`onEdit`, `onOpen`, temporizados) |
| **UrlFetchApp** | Exportação de planilhas em `.xlsx` via endpoint de exportação do Google |
| **clasp** | Versionamento e publicação do código no projeto Apps Script |

Não há framework, não há backend próprio e não há banco de dados externo. Não há
Inteligência Artificial envolvida na execução deste sistema.

---

## Funcionalidades

### Catálogo de produtos

- Cadastro central de produtos com código, categoria, produto, unidade de medida,
  pregão e situação (`Ativo`).
- Organização automática da aba por categoria, com faixas de título e formatação
  aplicada por código.
- **Inserção inteligente de novas linhas:** ao cadastrar um item de um pregão já
  existente, a linha é criada imediatamente abaixo do último item **daquele mesmo
  número de pregão**, respeitando a ordenação por número e data em categorias que
  contêm pregões diferentes.
- Ordenação numérica real dos pregões (`14/2025` antes de `101/2025`), em vez de
  ordenação alfabética de texto, usando o número do pregão, o ano e a data de vigência.
- Importação incremental de novos lotes de pregões, sem sobrescrever o que já existe.
- Controle mensal de quais categorias ficam liberadas para as unidades.

### Sincronização central → unidades

- Propagação manual e controlada do catálogo para todas as planilhas de unidade.
- Montagem da aba `Pedido` de cada unidade com produtos agrupados por categoria,
  validação de dados na coluna de status e proteção da aba de catálogo.
- Mês de referência escrito no topo da planilha de cada unidade.
- Instalação programática do menu e dos gatilhos nas planilhas das unidades.

### Coleta unidade → central

- Gatilho `onEdit` em cada planilha de unidade: ao marcar um item como `Enviado`,
  a central é atualizada.
- Debounce e fila de sincronização para evitar execuções redundantes quando várias
  linhas são editadas em sequência.
- Puxada manual sob demanda pelo menu, com diagnóstico detalhado do que foi lido.
- Leitura flexível da planilha da unidade: as colunas são detectadas pelo cabeçalho,
  o que mantém a leitura funcionando mesmo com layouts antigos.

### Consolidação e acompanhamento

- Aba `Resumo`: matriz produto × unidade agrupada por pregão e por categoria, com
  quantidades por unidade.
- Aba `Consolidados`: registro linha a linha de cada pedido enviado.
- Aba `Pedidos das Unidades`: quais unidades já tiveram o pedido recolhido e quais
  continuam pendentes.
- Registro de log das ações do sistema para auditoria técnica.

### Relatórios

| Relatório | Formato | Conteúdo |
|---|---|---|
| Relatório de pedidos do mês | `.xlsx` | Pedidos agrupados por pregão → produto → unidades e quantidades |
| Pregões ativos | `.xlsx` | Cada pregão vigente e os produtos que disponibiliza |
| Pregões inativos | `.xlsx` | Mesmo formato, para pregões encerrados |
| Resumo | `.pdf` | Exportação direta da aba `Resumo` |
| Relatório administrativo | `.docx` / `.pdf` | Visão por unidade, com cabeçalho institucional |
| Guia do sistema | `.pdf` | Material de apresentação e treinamento |

### Manutenção

- **Nova rodada:** zera os pedidos da central e de todas as unidades preservando o
  catálogo, os itens marcados como inativos e os IDs cadastrados. Processa em blocos
  com continuação agendada, para não esbarrar no limite de 6 minutos por execução
  do Apps Script.
- Reparo em etapas, executado em segundo plano, para planilhas com layout corrompido.
- Diagnóstico de gatilhos e de conexão com as unidades.
- Resolução automática de IDs de planilhas pelo nome do arquivo no Drive.
- Controle de visibilidade das abas: o operador vê apenas as abas operacionais.

---

## Arquitetura

```
                     ┌──────────────────────────────┐
                     │      PLANILHA CENTRAL        │
                     │   (projeto Apps Script)      │
                     │                              │
                     │  Produtos    → catálogo      │
                     │  Categorias  → liberação     │
                     │  Unidades    → IDs           │
                     │  Consolidados→ dados brutos  │
                     │  Resumo      → matriz        │
                     │  Pedidos das Unidades        │
                     │  Log         → auditoria     │
                     └───────┬──────────────▲───────┘
                             │              │
          catálogo (manual)  │              │  pedidos (onEdit + manual)
                             ▼              │
   ┌─────────────┬─────────────┬────────────┴──┬─────────────┐
   │ UNIDADE A   │ UNIDADE B   │  UNIDADE C    │     ...     │
   │ aba Pedido  │ aba Pedido  │  aba Pedido   │             │
   │ aba Catalogo│ aba Catalogo│  aba Catalogo │             │
   └─────────────┴─────────────┴───────────────┴─────────────┘
```

O código vive **inteiramente no projeto Apps Script da planilha central**. As planilhas
das unidades não têm código próprio: elas recebem gatilhos instaláveis que apontam de
volta para esse mesmo projeto.

Organização do código-fonte:

```
src/
├── appsscript.json          Manifesto: escopos OAuth, fuso, runtime V8
├── core/                    Config, menu, utilitários, cadastro de unidades
├── catalogo/                Produtos, categorias, validações, novos pregões
├── sincronizacao/           Central ↔ unidades, gatilhos, mês de referência
├── consolidacao/            Aba Resumo e situação dos pedidos
├── relatorios/              Excel, PDF e Word
├── manutencao/              Setup, reparo, nova rodada, diagnóstico
├── interface/               Visual das unidades e janela de ajuda
└── dados/                   Dados iniciais fictícios para o setup
```

> As subpastas existem para organizar o repositório. O `clasp` as preserva como
> prefixo no nome dos arquivos dentro do editor do Apps Script, que exibe o resultado
> como pastas. Se você copiar os arquivos manualmente, a divisão em pastas é opcional —
> o Apps Script trata todos os arquivos do projeto como um único escopo global.

Detalhamento em [`docs/arquitetura.md`](docs/arquitetura.md).

---

## Instalação

Resumo. O passo a passo completo está em [`docs/instalacao.md`](docs/instalacao.md).

**Use sempre um ambiente de testes com planilhas novas.** Nunca aponte este código
para planilhas em produção antes de validar o fluxo.

### Opção A — com clasp (recomendada)

```bash
npm install -g @google/clasp
clasp login

git clone https://github.com/gabrielbcaracini-hub/automacao-planilhas-javascript.git
cd automacao-planilhas-javascript

cp .clasp.json.example .clasp.json
# edite .clasp.json e coloque o scriptId do SEU projeto Apps Script

clasp push
```

### Opção B — cópia manual

1. Crie uma planilha no Google Sheets (será a planilha central).
2. Abra **Extensões → Apps Script**.
3. Crie um arquivo para cada `.gs` de `src/` e cole o conteúdo.
4. Crie o arquivo HTML `Ajuda` e cole o conteúdo de `src/interface/Ajuda.html`.
5. Substitua o `appsscript.json` pelo de `src/appsscript.json`
   (ative "Mostrar arquivo de manifesto" nas configurações do projeto).

---

## Configuração

Toda a configuração fica em [`src/core/Config.gs`](src/core/Config.gs).

```javascript
var CONFIG = {
  // ID da planilha central. Deixe vazio para usar a planilha ativa.
  SPREADSHEET_MAE_ID: '',

  ORGANIZACAO: {
    NOME: 'PREFEITURA MUNICIPAL DE EXEMPLO',  // cabeçalho dos relatórios
    LOGO_URL: ''                              // logotipo (vazio = sem imagem)
  },

  ABAS: { /* nomes das abas */ },
  PRODUTOS: { /* posição das colunas */ },
  TEMA: { /* cores da camada visual */ },

  SYNC_DEBOUNCE_MS: 8000,     // intervalo mínimo entre sincronizações automáticas
  CONSOLIDACAO_HORA: 7        // hora do gatilho diário de consolidação
};
```

Os dados de demonstração ficam em [`src/dados/DadosIniciais.gs`](src/dados/DadosIniciais.gs):
nomes das unidades, IDs das planilhas e catálogo semente. Substitua pelo seu conteúdo.

Escopos OAuth exigidos, gatilhos criados e parâmetros ajustáveis estão descritos em
[`docs/configuracao.md`](docs/configuracao.md).

---

## Utilização

Depois da instalação, recarregue a planilha central. O menu **Sistema de Compras**
aparece na barra superior.

### Preparação inicial

| Passo | Ação |
|---|---|
| 1 | Execute `setupPlanilhaMae` no editor para criar as abas |
| 2 | Execute `menuImportarUnidadesCompleto` para criar as planilhas das unidades |
| 3 | Execute `instalarGatilhos` para registrar os gatilhos automáticos |
| 4 | Execute `menuConectarTudoAgora` para conectar as unidades |

### Operação mensal

| Passo | Ação | Onde |
|---|---|---|
| 1 | Cadastrar produtos novos — **➕ Nova linha na aba Produtos** | Central |
| 2 | Liberar as categorias do mês na aba `Categorias` | Central |
| 3 | **⚡ Atualizar produtos nas unidades agora** | Central |
| 4 | Preencher as quantidades e marcar `Enviado` | Unidade |
| 5 | **📥 Puxar pedidos das unidades** | Central |
| 6 | Conferir as abas `Resumo` e `Pedidos das Unidades` | Central |
| 7 | **📥 Baixar relatório Excel** | Central |
| 8 | **🗑️ Nova rodada — limpar pedidos** no fim do ciclo | Central |

### Cadastrando um produto de um pregão já existente

Quando uma categoria contém mais de um número de pregão (por exemplo `05/2025` e
`156/2025`), o sistema pergunta qual número está sendo cadastrado e posiciona a linha
nova no bloco correto, mantendo a sequência numérica.

Passo a passo completo em [`docs/funcionalidades.md`](docs/funcionalidades.md).

---

## Exemplos

A pasta [`exemplos/dados-ficticios/`](exemplos/dados-ficticios/) contém arquivos CSV
com o formato esperado por cada aba:

| Arquivo | Aba correspondente |
|---|---|
| `produtos.csv` | `Produtos` — catálogo central |
| `categorias.csv` | `Categorias` — liberação mensal |
| `unidades.csv` | `Unidades` — nome e ID de cada planilha |
| `pedido-unidade.csv` | `Pedido` — planilha de uma unidade |
| `consolidados.csv` | `Consolidados` — saída da consolidação |

Todos os dados são fictícios e servem para montar um ambiente de testes.

---

## Resultados

O que a automação entrega, em termos verificáveis no próprio código:

- **Uma única fonte de verdade para o catálogo.** Produtos são cadastrados em um lugar
  só e distribuídos a partir dele, o que elimina versões divergentes circulando entre
  as unidades.
- **Consolidação sem digitação manual.** Os pedidos chegam à planilha central pelo
  gatilho de edição ou pela puxada manual; ninguém recopia dados de uma planilha
  para outra.
- **Visibilidade do andamento da rodada.** A aba `Pedidos das Unidades` separa, a
  qualquer momento, quem já respondeu de quem ainda não respondeu.
- **Relatórios padronizados.** Os mesmos dados geram Excel, PDF e Word com formatação
  fixa, sem montagem manual.
- **Rastreabilidade.** A aba `Consolidados` guarda cada pedido linha a linha, com
  unidade, quantidade, observação e data; a aba `Log` registra as ações do sistema.
- **Ordenação correta dos pregões.** O agrupamento por número e data de vigência vale
  igualmente no cadastro, no resumo e nos relatórios.

Este projeto não mede tempo economizado nem produtividade, e nenhum número desse tipo
é afirmado aqui.

---

## Documentação completa

| Documento | Conteúdo |
|---|---|
| [`docs/arquitetura.md`](docs/arquitetura.md) | Estrutura de dados, abas, módulos e como se relacionam |
| [`docs/instalacao.md`](docs/instalacao.md) | Instalação passo a passo em ambiente de testes |
| [`docs/configuracao.md`](docs/configuracao.md) | Parâmetros, escopos OAuth, permissões e gatilhos |
| [`docs/funcionalidades.md`](docs/funcionalidades.md) | Cada funcionalidade e como executá-la |
| [`docs/fluxo-automacao.md`](docs/fluxo-automacao.md) | Fluxo completo dos dados e dos gatilhos |
| [`docs/seguranca.md`](docs/seguranca.md) | O que foi anonimizado e por quê |
| [`docs/scripts/`](docs/scripts/) | Documentação módulo a módulo |

---

## Limitações conhecidas

- **Limite de 6 minutos por execução.** Operações que percorrem todas as unidades
  (nova rodada, migração de layout, sincronização de catálogo) processam em blocos e
  agendam a continuação em um gatilho temporizado.
- **Cotas do Google Apps Script.** O número de execuções de gatilho, chamadas ao
  `UrlFetchApp` e criações de arquivo por dia é limitado pela conta usada.
- **Permissões de Drive.** A conta que executa precisa de acesso de edição a todas as
  planilhas das unidades. Os arquivos gerados consomem a cota de armazenamento dessa
  conta.
- **Gatilhos instaláveis por conta.** Os gatilhos `onEdit` das unidades pertencem à
  conta que os instalou; se essa conta perder o acesso, a sincronização automática
  para e precisa ser reinstalada.
- **Sem testes automatizados.** O Apps Script não roda fora do ambiente do Google.
  A validação deste repositório cobriu a sintaxe JavaScript de todos os arquivos e a
  integridade das referências entre funções, mas **as funcionalidades não foram
  executadas** — testá-las exigiria planilhas reais.
- **Nomes de arquivo com prefixo de pasta.** Ao enviar com `clasp`, os arquivos chegam
  ao editor do Apps Script com o caminho da pasta no nome (`core/Config`).
- **Dependência de layout.** A leitura das planilhas das unidades detecta colunas pelo
  cabeçalho; renomear cabeçalhos de forma muito diferente do esperado quebra a leitura.

---

## Autor

**Gabriel Braga**

Área de interesse: Automação Digital, Inteligência Artificial Generativa e
Desenvolvimento de Software.

- GitHub: [@gabrielbcaracini-hub](https://github.com/gabrielbcaracini-hub)

---

## Licença

Projeto publicado como portfólio técnico. Todos os direitos reservados.
O código pode ser lido e estudado; o uso, a cópia ou a redistribuição dependem de
autorização prévia do autor.
