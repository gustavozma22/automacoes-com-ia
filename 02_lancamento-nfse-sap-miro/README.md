# 02 · Lançamento de NFS-e e recibos na MIRO

Executor para notas fiscais de serviço e recibos. Usa a mesma planilha de lote do projeto 01, e cada executor processa só as linhas do seu tipo.

## Quatro tipos de documento, quatro caminhos de tela

| | NFS-e | Recibo sem documento | Recibo de fatura | Recibo de boleto |
|---|---|---|---|---|
| Categoria fiscal | preenchida | não toca | não toca | não toca |
| Tela Nota Fiscal | sim | não abre | não abre | não abre |
| Referência | nº da NF | 4 últimos do pedido | nº da fatura | 4 últimos do nosso número |
| Confirmação ao gravar | uma | duas | duas | duas |

O **tipo vazio retém**, não vira recibo. No SAP, ausência de categoria significa recibo; na planilha, vazio significa "ainda não preenchi". Tratar os dois como a mesma coisa faria notas fiscais entrarem sem os dados fiscais.

## Destaques

- **Chave nacional da NFS-e (50 dígitos):** confere CNPJ, número, competência e dígito verificador a partir da chave. Chave pela metade não é completada: retém.
- Nem todo município usa o padrão nacional; o que decide é a presença da chave, não a cidade.
- **Modo DESPEJO** (somente leitura): fotografa todos os campos da tela aberta num `.txt`. Resolveu em um minuto um erro que a teoria não explicava.
- Só grava com saldo exatamente 0,00. Não arredonda nem completa dados.
- [`tools/lint_vbs.py`](../tools/lint_vbs.py): verificador estático que nasceu neste projeto e hoje roda no CI para todos os scripts. Detecta colisão de nomes (VBScript ignora maiúsculas), número errado de argumentos, blocos desbalanceados, variáveis não declaradas e chaves de mapa inexistentes.

## Arquivos

| Arquivo | Função |
|---|---|
| `LANCAR_SERVICOS.vbs` | Executor (~1.800 linhas) |
| `MAPA_CAMPOS_SERVICOS.txt` | IDs de tela e parâmetros |

Uso: `cscript LANCAR_SERVICOS.vbs` (lança) ou `cscript LANCAR_SERVICOS.vbs DESPEJO` (só lê a tela).
