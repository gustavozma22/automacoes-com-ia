# Lições aprendidas com SAP GUI Scripting + VBScript

Anotações do que custou tempo durante o desenvolvimento. Servem para qualquer automação de tela no SAP.

## Armadilhas técnicas

| Armadilha | O que acontece | Como resolvi |
|---|---|---|
| `Children(i)` não devolve objeto | `Children.Count` funciona, mas o acesso por índice vem vazio | `Children.ElementAt(i)`, com `.Item(i)` de reserva |
| VBScript ignora maiúsculas | `Dim filho` e `Function Filho` são o mesmo nome; a chamada falha em silêncio | Linter próprio que detecta colisão de nomes |
| Identificador terminado em `_` | `_` é continuação de linha | Linter |
| `GetCell` devolve o texto **renderizado** | Coluna estreita trunca o valor lido | Alargar a coluna antes de ler a grade |
| A grade precisa de refresh | Lida cedo demais, um pedido aparece com dois itens | Ler primeiro; só forçar refresh se vier vazia |
| A grade às vezes só monta depois de um Enter | Leitura vazia vira retenção falsa | Ler em várias passadas |
| Rolar invalida o objeto da tabela | Cada rolagem é uma ida ao servidor | Buscar a tabela pelo ID de novo após rolar |
| `RowCount` não é a quantidade de itens | É o tamanho visível da tabela, com linhas vazias | Parar na primeira linha vazia |
| Tecla vai para a janela errada | `sendVKey` numa janela modal precisa ir para a `wnd[1]` | Usar exatamente a janela que a gravação usou |
| Modal aberta engole comandos | Tudo que vem depois some | Conferir explicitamente se a modal fechou |
| Número com separador local | `CDbl("1.234,56")` depende da configuração regional | Converter manualmente para centavos inteiros |
| Confirmação dupla ao gravar | Alguns documentos abrem uma segunda janela | Confirmar, checar se sobrou janela, e só então confirmar de novo |

## Método de trabalho

- **Nunca deduzir um ID de tela.** Usar a gravação do próprio SAP ou um script de "despejo" somente leitura que lista todos os campos da tela aberta. Rodar antes e depois de uma ação manual mostra o que mudou.
- **Não pular passos da gravação.** Se a gravação faz, o script faz, inclusive *onde* faz. Simplificar só depois de funcionar.
- **Verificador estático antes de pedir para rodar** (ver [`tools/lint_vbs.py`](../tools/lint_vbs.py)). E testar o verificador com defeitos injetados: o meu tinha uma checagem que nunca disparava.
- **Conferir a planilha contra o código ao inserir colunas.** Uma coluna nova desloca tudo e o script passa a ler o campo errado em silêncio.
- **Trava por sintoma é trava errada.** Perguntar a condição real (há retenção de imposto?), não um sinal indireto (qual é a forma de pagamento?).
- **Conferir a premissa contra os dados antes de gravá-la no código.** Duas de 37 chaves fugiam do padrão que "sempre" valia.
