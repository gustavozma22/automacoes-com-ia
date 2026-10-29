# 04 · Download de faturas e agendamento

Peças que fazem o processo rodar sozinho.

| Arquivo | Função |
|---|---|
| `BAIXAR_NFCOM_EMAILS.ps1` | Copia os anexos PDF/XML das pastas "Faturas *" de uma caixa compartilhada do Outlook para a pasta do mês. Não move, não apaga e não altera e-mails. |
| `VERIFICAR_LOTE_E_LANCAR.ps1` | Olha a planilha de lote; se houver nota pendente, chama o executor certo (NFCom ou NFS-e). |
| `LANCAR_TUDO.vbs` | Roda os dois executores em sequência sobre o mesmo lote. |
| `INSTALAR_AGENDAMENTOS.bat` / `REMOVER_AGENDAMENTOS.bat` | Criam/removem as tarefas no Agendador do Windows. |
| `config/NFCOM_EMAIL_REGRAS.exemplo.json` | Regras por fornecedor (exemplo com dados fictícios). |

## Destaques

- **Deduplicação real:** o anexo é salvo numa pasta temporária e o hash do arquivo é comparado com o que já existe na pasta, com o log de notas processadas e com o log do próprio script. Um anexo já tratado não volta como `_2`.
- Anexos que não são NFCom (NF-e modelo 55, eventos de cancelamento) vão para uma pasta de revisão.
- O disparo automático só acontece se **tudo** for verdade: dia útil, entre 7h e 20h, SAP aberto, planilha fechada, nenhum executor rodando e o mesmo conjunto de pendências não disparado na última hora.
- Mesmo agendado, o executor continua pedindo a confirmação de ambiente: a pessoa confirma na tela e o resto roda sozinho.

Agendamento padrão: download em dias úteis às 13h30; verificação do lote a cada 10 minutos.
