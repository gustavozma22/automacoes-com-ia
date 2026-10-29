@echo off
rem Baixa os anexos das pastas "Faturas *" da caixa compartilhada de faturas para NFCOMs_<MES>.
rem Uso: duplo clique (baixa) ou "BAIXAR_NFCOM_EMAILS.bat -Simular" (so mostra).
pushd "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0BAIXAR_NFCOM_EMAILS.ps1" %*
popd
pause
