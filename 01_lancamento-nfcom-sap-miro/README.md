# 01 · Lançamento de NFCom na MIRO

Executor local que lê um lote de notas de telecom (NFCom, modelo 62) numa planilha Excel e reproduz no SAP a sequência que o time fazia à mão.

## Fluxo por nota

1. **MIRO:** data, referência e montante.
2. Seleção múltipla de pedidos (sempre pela janela de seleção, mesmo com um pedido só).
3. **Valor líquido de cada item, casado pelo número do pedido** e não pela posição na grade.
4. Aba Detalhe: categoria da nota fiscal.
5. Tela Nota Fiscal: protocolo, data e hora de autorização, tipo de emissão, número aleatório da chave.
6. Rateio e conferência dos impostos por item, com tolerância configurável em centavos.
7. **Simular → conferir saldo zero → Lançar.**
8. Escreve STATUS, número do documento, erro, data/hora e usuário na própria linha e salva a cada nota.

## Travas de segurança

- Não faz login: usa a sessão SAP já aberta e autenticada.
- Confirma o ambiente (produção ou qualidade) e avisa se o usuário SAP for diferente do habitual.
- Ignora linhas com STATUS preenchido, o que impede lançamento em duplicidade.
- Não grava se houver erro, aviso ou saldo diferente de zero. Na dúvida, marca como NÃO LANÇADA com o motivo.
- Uma nota com problema não derruba o lote: registra, volta à tela inicial e segue.
- Em falha, grava um "despejo" da tela (todos os campos e valores) para diagnóstico.

## Arquivos

| Arquivo | Função |
|---|---|
| `LANCAR_NFCOM.vbs` | Executor (~2.700 linhas) |
| `MAPA_CAMPOS.txt` | IDs de tela e parâmetros, com a origem de cada um |
| `ferramentas/VER_IMPOSTO.vbs` | Sonda usada para mapear a tabela de impostos |

A bipagem do boleto logo após o lançamento usa também o `MAPA_CAMPOS_BOLETO.txt` do [projeto 03](../03_bipagem-boletos-sap), que deve ficar na mesma pasta do script.

## Como usar

1. Abrir o SAP GUI e fazer login.
2. Fechar a planilha de lote no Excel.
3. Deixar planilha, `MAPA_CAMPOS.txt` e o script na mesma pasta e dar duplo clique no script.
