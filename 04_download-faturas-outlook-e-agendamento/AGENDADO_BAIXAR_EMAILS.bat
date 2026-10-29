@echo off
rem Chamado pelo Agendador (dias uteis 13:30): baixa os anexos das pastas "Faturas *"
rem da caixa compartilhada de faturas para NFCOMs_<MES>. Sem pausa; o resultado fica em CONSULTA\NFCOM_EMAIL_run_log.txt.
pushd "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0BAIXAR_NFCOM_EMAILS.ps1" %*
popd
