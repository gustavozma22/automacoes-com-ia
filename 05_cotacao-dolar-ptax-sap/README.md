# 05 · Cotação do dólar (PTAX) no SAP

Busca a cotação PTAX de **venda** do dia na API pública do Banco Central (Olinda/PTAX) e preenche a tabela de câmbio do SAP (`V_TCURR`) nas duas direções (BRL→USD e USD→BRL) com a data de hoje.

- Se ainda não houver PTAX de fechamento publicada, avisa e não grava nada.
- Pede confirmação antes de gravar (`PEDIR_CONFIRMACAO = True`).
- Formata a data no padrão de cada lado: `DDMMAAAA` no SAP e `MM-DD-AAAA` na API.
