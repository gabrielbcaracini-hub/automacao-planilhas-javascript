# Fluxo de automação

Como os dados circulam entre a planilha central e as unidades, e o que dispara cada
etapa.

## Ciclo completo de uma rodada

```
┌─────────────────────────────────────────────────────────────────┐
│ 1. CADASTRO                                        [manual]     │
│    Administrador cadastra produtos na aba Produtos              │
│    menuInserirLinhaNovoProduto                                  │
└────────────────────────────┬────────────────────────────────────┘
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│ 2. LIBERAÇÃO                                       [manual]     │
│    Marca quais categorias valem no mês (aba Categorias)         │
│    menuAtualizarControleCategorias                              │
└────────────────────────────┬────────────────────────────────────┘
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│ 3. DISTRIBUIÇÃO                                    [manual]     │
│    Catálogo → todas as unidades                                 │
│    sincronizarCatalogoTodasUnidades                             │
│    • reescreve a aba Catalogo                                   │
│    • monta a aba Pedido agrupada por categoria                  │
│    • aplica validações e visual                                 │
│    • grava o mês de referência                                  │
└────────────────────────────┬────────────────────────────────────┘
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│ 4. PREENCHIMENTO                              [pela unidade]    │
│    Operador lança quantidades e marca Enviado                   │
└────────────────────────────┬────────────────────────────────────┘
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│ 5. COLETA                              [automático + manual]    │
│    onEditPedidoUnidade  →  fila  →  gatilho temporizado         │
│    ou menuAtualizarResumoPedidosAgora                           │
└────────────────────────────┬────────────────────────────────────┘
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│ 6. CONSOLIDAÇÃO                                 [automático]    │
│    Consolidados → Resumo → Pedidos das Unidades                 │
└────────────────────────────┬────────────────────────────────────┘
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│ 7. RELATÓRIOS                                      [manual]     │
│    Excel, PDF e Word gerados no Drive                           │
└────────────────────────────┬────────────────────────────────────┘
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│ 8. ENCERRAMENTO                                    [manual]     │
│    menuResetNovaRodadaCompras                                   │
│    Zera pedidos preservando o catálogo → volta ao passo 1       │
└─────────────────────────────────────────────────────────────────┘
```

As etapas 1, 2, 3, 7 e 8 exigem ação do administrador. As etapas 5 e 6 acontecem
sozinhas. A etapa 4 é a única feita pelas unidades.

## Sincronização automática em detalhe

O caminho da unidade até a central, item por item:

```
 PLANILHA DA UNIDADE                      PLANILHA CENTRAL
 ───────────────────                      ────────────────

 operador marca "Enviado"
         │
         ▼
 ┌──────────────────────┐
 │ onEditPedidoUnidade  │  gatilho instalável
 └──────────┬───────────┘
            │  a edição toca a coluna Status ou Quantidade?
            │  (rangeIncluiColunasPedido_)
            │
            ├── não → encerra
            │
            └── sim
                 ▼
       ┌───────────────────────────┐
       │ agendarSyncPedidoUnidade  │
       │ Debounced_                │
       │                           │
       │ grava o ID da unidade     │
       │ na fila em                │
       │ PropertiesService         │
       └───────────┬───────────────┘
                   │
                   │  já existe gatilho pendente?
                   │  sim → encerra (debounce)
                   │  não → cria gatilho após ~7 s
                   ▼
       ┌───────────────────────────┐
       │ processarSyncPedido       │
       │ PendenteRapido_           │
       └───────────┬───────────────┘
                   │  lê a fila e processa cada unidade
                   ▼
       ┌───────────────────────────┐        ┌─────────────────────┐
       │ sincronizarUnidadePorId   │───────▶│ Consolidados        │
       │                           │        ├─────────────────────┤
       │ lerPedidosEnviadosFlexivel│───────▶│ Resumo              │
       │ atualizarMaeComPedidos    │        ├─────────────────────┤
       │ Unidade                   │───────▶│ Pedidos das Unidades│
       └───────────┬───────────────┘        └─────────────────────┘
                   │
                   │  falhou?
                   └── sim → criarGatilhoSyncAtrasadoUnidade_ (nova tentativa)
```

### Por que a fila e o debounce

Marcar 30 itens como `Enviado` dispara 30 eventos `onEdit`. Sem controle, seriam 30
sincronizações completas da mesma unidade.

A fila resolve isso: cada evento apenas **registra o ID da unidade** em
`PropertiesService` e verifica se já existe um gatilho pendente. Se existe, encerra.
As 30 edições viram **uma** sincronização, executada alguns segundos depois.

`CONFIG.SYNC_DEBOUNCE_MS` (8000 ms por padrão) define a janela mínima entre
sincronizações da mesma unidade.

## Processamento em blocos

O Apps Script encerra qualquer execução em **6 minutos**. Operações que percorrem
todas as unidades usam o mesmo padrão:

```
┌────────────────────────────────────────────────┐
│ início                                         │
│   lê o progresso salvo em PropertiesService    │
└────────────────────┬───────────────────────────┘
                     ▼
         ┌───────────────────────┐
         │ para cada unidade:    │◀────────────┐
         │   processa            │             │
         │   salva o progresso   │             │
         └───────────┬───────────┘             │
                     │                         │
                     ▼                         │
         ┌───────────────────────┐             │
         │ perto do limite?      │             │
         └───┬───────────────┬───┘             │
             │ não           │ sim             │
             └───────────────┼─────────────────┘
                             ▼
         ┌─────────────────────────────────────┐
         │ agenda gatilho temporizado          │
         │ e encerra a execução                │
         └─────────────────┬───────────────────┘
                           ▼
         ┌─────────────────────────────────────┐
         │ o gatilho reinicia do ponto salvo   │
         │ e se remove ao concluir             │
         └─────────────────────────────────────┘
```

Rotinas que usam esse padrão:

| Rotina | Continuação |
|---|---|
| Nova rodada | `continuarResetNovaRodada_` |
| Migração de layout | `processarLoteMigracaoLayoutUnidades_` |
| Reparo do sistema | `executarReparoEmPassos` |
| Situação dos pedidos | `processarTotalPorProdutoPendente_` |
| Fila de sincronização | `processarSyncPedidoPendenteRapido_` |

Para o operador, uma operação longa pode parecer "terminada antes da hora". Ela
continua em segundo plano e conclui sozinha.

## Gatilhos ativos

| Gatilho | Tipo | Onde | Dispara |
|---|---|---|---|
| `onOpen` | Simples | Central | Ao abrir a planilha |
| `onOpenUnidadeInstalavel_` | Instalável | Cada unidade | Ao abrir a planilha |
| `onEditPedidoUnidade` | Instalável | Cada unidade | A cada edição |
| `consolidarPedidosAgendado` | Temporizado | Projeto | Diariamente |
| Gatilhos de continuação | Temporizado, único | Projeto | Sob demanda, removem-se ao concluir |

### O que `onOpen` faz na central

Além de montar o menu:

1. detecta se a planilha aberta é uma unidade e, nesse caso, desvia para `onOpenUnidade_`;
2. aplica a visibilidade das abas para o operador;
3. remove gatilhos de catálogo automático remanescentes de versões antigas;
4. atualiza a aba `Pedidos das Unidades`;
5. restaura o visual do `Resumo` se detectar migração pendente;
6. importa o lote de novos pregões, se ainda não tiver sido importado.

Cada etapa roda dentro do próprio `try/catch`: uma falha isolada não impede o menu de
aparecer.

## Tratamento de erros

### `try/catch` por operação

Operações que percorrem várias unidades isolam cada uma. Uma unidade inacessível não
interrompe as demais — a falha é registrada e reportada ao final.

```javascript
for (var i = 0; i < unidades.length; i++) {
  try {
    processarUnidade(unidades[i]);
    ok.push(unidades[i].nome);
  } catch (e) {
    falhas.push(unidades[i].nome + ': ' + e.message);
  }
}
```

### Locks

`adquirirLock` usa o `LockService` para impedir execuções simultâneas da mesma
operação — por exemplo, o gatilho diário disparando enquanto o administrador puxa os
pedidos manualmente.

### Reprocessamento

Quando a sincronização de uma unidade falha, `criarGatilhoSyncAtrasadoUnidade_`
agenda uma nova tentativa alguns segundos depois, em vez de descartar a atualização.

### Fallback de leitura

`lerPedidosEnviadosFlexivel` detecta as colunas pelo cabeçalho. Se o layout não for
reconhecido, `lerPedidosFallbackSimples_` assume o formato padrão. O sistema continua
lendo planilhas de versões anteriores.

### Resolução de IDs

Quando o ID cadastrado de uma unidade não abre, `resolverIdUnidadePeloNome_` procura
a planilha pelo nome do arquivo no Drive e atualiza o cadastro
(`atualizarIdUnidadeNaAba`). Cobre o caso da planilha recriada.

## Auditoria

| Origem | Conteúdo |
|---|---|
| Aba `Log` | Ações do sistema: data/hora, ação e detalhe (`registrarLog`) |
| Aba `Consolidados` | Cada item pedido, com unidade, quantidade, observação e data |
| **Execuções** do Apps Script | Histórico técnico, com erros e duração |
| Stackdriver | `exceptionLogging: "STACKDRIVER"` no manifesto |
