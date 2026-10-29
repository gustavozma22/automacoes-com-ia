# 03 · Bipagem de boletos

Terceiro executor do processo. Não lança nota: volta nos documentos já lançados pelos projetos 01 e 02 e preenche, na linha do fornecedor, a **forma de pagamento** e o **código de barras** do boleto.

## Fluxo

1. Número do documento MIRO + ano → busca o documento contábil (FB03 → lista de documentos).
2. Abre em modo de alteração (FB02) e encontra a linha do fornecedor pela chave de lançamento.
3. Grava forma de pagamento e código de barras.
4. **Relê depois de salvar** para provar que gravou.

## Validação do código de barras

- Dígitos verificadores de campo e geral (módulo 10 e módulo 11).
- Boleto bancário × boleto de arrecadação (concessionárias, começa com 8): layouts diferentes.
- **Fator de vencimento:** desde fevereiro/2025 o fator reiniciou em 1000, e bancos diferentes usam bases que diferem em um dia. O script aceita as duas bases e trata a decodificação só como conferência auxiliar contra o vencimento impresso.

## Arquivos

| Arquivo | Função |
|---|---|
| `LANCAR_BOLETOS.vbs` | Executor |
| `MAPA_CAMPOS_BOLETO.txt` | Caminho de telas, mapeado com 6 sondas de leitura |
