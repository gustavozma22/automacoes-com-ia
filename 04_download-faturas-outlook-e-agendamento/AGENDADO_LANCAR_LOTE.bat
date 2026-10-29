@echo off
rem Chamado pelo Agendador a cada 10 min: se houver nota pendente no lote,
rem dispara LANCAR_NFCOM.vbs / LANCAR_SERVICOS.vbs (ver VERIFICAR_LOTE_E_LANCAR.ps1).
rem Teste sem disparar: AGENDADO_LANCAR_LOTE.bat -Simular
pushd "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0VERIFICAR_LOTE_E_LANCAR.ps1" %*
popd
