# 06 · Validação de notas com IA

Etapa que alimenta os executores. Um assistente de IA (Claude, no modo Cowork) segue um roteiro fixo e, a cada rodada, processa **só os arquivos novos** da pasta do mês.

## Roteiro

1. Registra o início da rodada.
2. Encontra os arquivos que ainda não foram processados (log de arquivos por nome e hash).
3. Extrai o texto de cada PDF/XML.
4. Lê os campos: CNPJ do emitente, número, data, valor total, itens, modelo fiscal, protocolo.
5. **Conferência aritmética:** a soma dos itens precisa bater com o valor da nota.
6. Renomeia o arquivo num padrão fixo (empresa, número, documento, valor).
7. Procura a nota na base de pedidos e valida fornecedor, valor e pedidos.
8. Escreve o resultado na planilha de validação e devolve número e valor da nota para a base.
9. Remove da aba do mês as linhas que já foram totalmente lançadas.
10. Gera um relatório e registra a duração da rodada.

## Por que funciona

- **Incremental:** nunca reprocessa nem duplica uma nota já vista.
- **Sem acesso ao SAP:** a IA só lê arquivos e escreve planilhas. O lançamento é feito pelos executores locais, na sessão da pessoa.
- **Erros aparecem cedo:** problemas como uma nota de telecom emitida no modelo fiscal errado ou itens que não fecham com o total são barrados antes de virarem lançamento.
