@echo off
rem Cria as duas tarefas no Agendador do Windows (rodar UMA vez, no seu PC):
rem   1) NFCOM - Baixar e-mails de faturas : dias uteis as 13:30
rem   2) NFCOM - Lancar lote        : a cada 10 min (so age se houver nota pendente no lote)
rem Para remover: REMOVER_AGENDAMENTOS.bat
set BASE=%~dp0
set BASE=%BASE:~0,-1%

schtasks /Create /F /TN "NFCOM - Baixar e-mails de faturas" /SC WEEKLY /D MON,TUE,WED,THU,FRI /ST 13:30 /TR "\"%BASE%\AGENDADO_BAIXAR_EMAILS.bat\""
schtasks /Create /F /TN "NFCOM - Lancar lote" /SC MINUTE /MO 10 /TR "\"%BASE%\AGENDADO_LANCAR_LOTE.bat\""

echo.
echo Tarefas criadas. Confira com: schtasks /Query /TN "NFCOM - Baixar e-mails de faturas"
echo                               schtasks /Query /TN "NFCOM - Lancar lote"
echo Importante: as duas rodam so com voce logado (precisam do Outlook e do SAP abertos).
pause
