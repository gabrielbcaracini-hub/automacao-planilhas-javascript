# Dados fictícios

Arquivos de exemplo com o formato esperado por cada aba do sistema.

**Todos os dados aqui são inventados.** Nomes de unidades, produtos, números de
pregão e razões sociais não correspondem a nada real. Servem para montar um ambiente
de testes e para entender a estrutura esperada em cada aba.

## Arquivos

| Arquivo | Aba | Onde fica |
|---|---|---|
| [`produtos.csv`](produtos.csv) | `Produtos` | Central |
| [`categorias.csv`](categorias.csv) | `Categorias` | Central |
| [`unidades.csv`](unidades.csv) | `Unidades` | Central |
| [`consolidados.csv`](consolidados.csv) | `Consolidados` | Central |
| [`pedido-unidade.csv`](pedido-unidade.csv) | `Pedido` | Planilha de uma unidade |

## Como usar

### Para montar um ambiente de testes

Prefira executar `menuImportarUnidadesCompleto`: ele monta tudo a partir de
`src/dados/DadosIniciais.gs`, que já contém os mesmos dados fictícios, e ainda cria
as planilhas das unidades.

### Para importar um CSV manualmente

1. Abra a aba de destino.
2. **Arquivo → Importar → Fazer upload**.
3. Escolha "Substituir planilha atual" ou "Substituir dados a partir da célula
   selecionada".
4. Em separador, escolha **vírgula**.

> Confira se os acentos vieram corretos. Os arquivos estão em UTF-8.

## Observações por arquivo

### `produtos.csv`

26 itens. Alguns detalhes são propositais:

- **`ELETRODOMÉSTICOS` tem dois números de pregão** (`05/2025` e `156/2025`). É o caso
  que a inserção posicionada de linhas precisa tratar: ao cadastrar um item novo do
  pregão 05, a linha entra no bloco do 05, não no topo da categoria.
- **`COMPRA DIRETA`** cobre o item sem pregão formal.
- **O item 026 está com `Ativo = Não`**, para exercitar o filtro de itens inativos e o
  relatório de pregões inativos.
- O pregão `63/2026` aparece em duas categorias diferentes, mostrando que um pregão
  pode atravessar categorias.

Formato do campo Pregão:

```
PREGÃO 156/2025 – DE 29/01/2026 – FORNECEDOR EXEMPLO H LTDA – 29/01/2027
       └──┬──┘      └────┬────┘    └────────┬────────────┘   └────┬────┘
        número        vigência            fornecedor           validade
```

O número e a data são extraídos desse texto para ordenar os pregões numericamente.

### `categorias.csv`

Mostra uma rodada em que nem tudo está liberado: `INFORMÁTICA`, `ELETRODOMÉSTICOS`,
`ARES CONDICIONADOS` e `COMPRA DIRETA` estão com `Não` e não seriam distribuídas às
unidades.

### `unidades.csv`

14 unidades. Os IDs são **placeholders sem validade** — não abrem nenhuma planilha.
Substitua pelos IDs reais das suas planilhas de teste.

A última unidade está com `Ativo = Não`, para exercitar o filtro de unidades ativas.

A ordem das linhas é a mesma de `DADOS_INICIAIS.unidades`, que define a ordem das
colunas na aba `Resumo`.

### `pedido-unidade.csv`

Como fica a aba `Pedido` de uma unidade no meio de uma rodada.

- A linha 1 traz o **mês de referência**.
- A linha 2 é o cabeçalho.
- Linhas que começam com `▸` são **títulos de categoria**: decorativas, ignoradas na
  leitura.
- Há itens em `Enviado` e em `Rascunho`: só os `Enviado` são recolhidos pela central.
- Alguns itens têm observação preenchida, outros não.

### `consolidados.csv`

Resultado da coleta: 29 linhas, 5 unidades que responderam.

Repare que nem todas as 14 unidades aparecem. As ausentes são exatamente as que a aba
`Pedidos das Unidades` listaria como pendentes.

Esta aba é a **fonte de verdade** dos pedidos: as abas `Resumo` e
`Pedidos das Unidades` são derivadas dela e podem ser reconstruídas a partir daqui.

## O que estes dados produzem

Com `consolidados.csv` carregado, a aba `Resumo` ficaria assim:

```
ITEM                              │ CENTRO │ NORTE │ SUL │ NÚCLEO I │ CONVIV. │
──────────────────────────────────┼────────┼───────┼─────┼──────────┼─────────┤
▸ PREGÃO 10/2025 — FORNECEDOR A
  ▸ PAPELARIA
    PAPEL SULFITE BRANCO A4 RESMA │   10   │   25  │  –  │     5    │    –    │
    CANETA ESFEROGRÁFICA AZUL     │    2   │    –  │ 12  │     –    │    –    │
▸ PREGÃO 35/2025 — FORNECEDOR C
  ▸ GÊNERO ALIMENTÍCIO
    ARROZ TIPO 1 PACOTE 5 KG      │   40   │   30  │ 20  │     –    │   50    │
    FEIJÃO CARIOCA PACOTE 1 KG    │   25   │   18  │  –  │     –    │   35    │
    ÓLEO DE SOJA FRASCO 900 ML    │   12   │    –  │  –  │     6    │    –    │
```

E a aba `Pedidos das Unidades`:

| Já pediram | Ainda não pediram |
|---|---|
| UNIDADE CENTRO | UNIDADE LESTE |
| UNIDADE NORTE | UNIDADE OESTE |
| UNIDADE SUL | NÚCLEO SOCIAL II |
| NÚCLEO SOCIAL I | CENTRO COMUNITÁRIO |
| CENTRO DE CONVIVÊNCIA | CASA DE APOIO |
| | POSTO DE ATENDIMENTO |
| | ALMOXARIFADO CENTRAL |
| | SEDE ADMINISTRATIVA |
