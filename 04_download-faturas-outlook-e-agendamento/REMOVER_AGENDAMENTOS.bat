@echo off
schtasks /Delete /F /TN "NFCOM - Baixar e-mails de faturas"
schtasks /Delete /F /TN "NFCOM - Lancar lote"
echo Tarefas removidas.
pause
