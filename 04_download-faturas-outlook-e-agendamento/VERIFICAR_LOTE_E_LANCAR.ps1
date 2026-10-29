# =====================================================================
#  VERIFICAR_LOTE_E_LANCAR.ps1
#  Olha a aba LOTE de LANCAMENTO_NFCOM.xlsx. Se houver nota pendente
#  (ID preenchido e STATUS vazio), dispara o executor certo:
#     TIPO TA                         -> LANCAR_NFCOM.vbs
#     TIPO ZH / RECIBO ...            -> LANCAR_SERVICOS.vbs
#
#  NAO altera nenhum arquivo do processo (planilhas, .vbs, MAPA_*):
#  so le o lote (somente leitura) e chama os executores como estao.
#  Os executores continuam mostrando a CONFIRMACAO DE AMBIENTE (PRD/QAS)
#  e o aviso final: voce confirma na tela, e o resto roda sozinho.
#
#  So dispara se TUDO abaixo for verdade (senao nao faz nada e registra):
#    - dia util, entre 07:00 e 20:00
#    - SAP Logon aberto
#    - LANCAMENTO_NFCOM.xlsx nao esta aberta no Excel (~$ ausente)
#    - nenhum executor ja em andamento
#    - o mesmo conjunto de pendencias nao foi disparado nos ultimos 60 min
#
#  Uso: AGENDADO_LANCAR_LOTE.bat  (agendado a cada 10 min)  |  -Simular
#  Log: CONSULTA\LANCAR_AUTO_log.txt   Estado: CONSULTA\LANCAR_AUTO_estado.json
# =====================================================================
param([switch]$Simular, [int]$EsperaMin = 60)
$ErrorActionPreference = 'Stop'
$dir   = $PSScriptRoot
$raiz  = Split-Path $dir -Parent
$cons  = Join-Path $raiz 'CONSULTA'
$xlsx  = Join-Path $dir 'LANCAMENTO_NFCOM.xlsx'
$lock  = Join-Path $dir '~$LANCAMENTO_NFCOM.xlsx'
$log   = Join-Path $cons 'LANCAR_AUTO_log.txt'
$estP  = Join-Path $cons 'LANCAR_AUTO_estado.json'

function Log($m) { Add-Content -Path $log -Value ("[{0}] {1}" -f (Get-Date).ToString('s'), $m) -Encoding UTF8 }

# ---- janela de funcionamento ----
$agora = Get-Date
if ($agora.DayOfWeek -in 'Saturday','Sunday' -or $agora.Hour -lt 7 -or $agora.Hour -ge 20) { exit 0 }
if (-not (Test-Path $xlsx)) { Log "Nao achei $xlsx"; exit 1 }

# ---- le o lote em SOMENTE LEITURA (abre o zip sem travar o arquivo) ----
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Ler-Entrada($zip, $nome) {
    $e = $zip.GetEntry($nome); if (-not $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Close() }
}
$fs = New-Object System.IO.FileStream($xlsx, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
try {
    [xml]$wb   = Ler-Entrada $zip 'xl/workbook.xml'
    [xml]$rels = Ler-Entrada $zip 'xl/_rels/workbook.xml.rels'
    $ssTxt = Ler-Entrada $zip 'xl/sharedStrings.xml'
    $ss = @()
    if ($ssTxt) {
        [xml]$ssx = $ssTxt
        foreach ($si in $ssx.sst.si) { $ss += ((@($si.SelectNodes('.//*[local-name()="t"]')) | ForEach-Object { $_.'#text' }) -join '') }
    }
    $rid = $null
    foreach ($s in $wb.workbook.sheets.sheet) { if ($s.name -eq 'LOTE') { $rid = $s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') } }
    if (-not $rid) { Log "Aba LOTE nao encontrada"; exit 1 }
    $alvo = ($rels.Relationships.Relationship | Where-Object { $_.Id -eq $rid }).Target
    $alvo = $alvo.TrimStart('/'); if ($alvo -notlike 'xl/*') { $alvo = 'xl/' + $alvo }
    [xml]$sh = Ler-Entrada $zip $alvo
} finally { $zip.Dispose(); $fs.Dispose() }

function Valor-Celula($c) {
    if (-not $c) { return '' }
    switch ($c.t) {
        's'         { return [string]$ss[[int]$c.v] }
        'inlineStr' { return ((@($c.SelectNodes('.//*[local-name()="t"]')) | ForEach-Object { $_.'#text' }) -join '') }
        default     { return [string]$c.v }
    }
}
$pendTA = @(); $pendSV = @()
foreach ($row in $sh.worksheet.sheetData.row) {
    $r = [int]$row.r; if ($r -lt 2) { continue }
    $cel = @{}
    foreach ($c in $row.c) { $col = ($c.r -replace '\d', ''); $cel[$col] = $c }
    $id = (Valor-Celula $cel['A']).Trim()
    if ($id -eq '') { break }                       # lote e contiguo: para no primeiro ID vazio
    $tipo = (Valor-Celula $cel['F']).Trim().ToUpper()
    $status = (Valor-Celula $cel['R']).Trim()
    if ($status -ne '') { continue }
    if ($tipo -eq 'TA') { $pendTA += $id }
    elseif ($tipo -in 'ZH','RECIBO SEM DOC','RECIBO FATURA','RECIBO BOLETO') { $pendSV += $id }
}
if (($pendTA.Count + $pendSV.Count) -eq 0) { exit 0 }  # nada no lote: silencio

# ---- condicoes de seguranca ----
$todos = (($pendTA + $pendSV) | Sort-Object) -join ','
if (-not (Get-Process -Name saplogon -ErrorAction SilentlyContinue)) { Log "Pendentes ($todos) mas SAP Logon fechado: nada feito."; exit 0 }
if (Test-Path $lock) { Log "Pendentes ($todos) mas a planilha do lote esta aberta no Excel: nada feito."; exit 0 }
$exec = Get-CimInstance Win32_Process -Filter "Name='wscript.exe' OR Name='cscript.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -like '*LANCAR_*' }
if ($exec) { Log "Executor ja em andamento: nada feito."; exit 0 }
$est = $null
if (Test-Path $estP) { try { $est = Get-Content $estP -Raw -Encoding UTF8 | ConvertFrom-Json } catch {} }
if ($est -and $est.pendentes -eq $todos -and ((Get-Date) - [datetime]$est.disparado_em).TotalMinutes -lt $EsperaMin) { exit 0 }

Log ("Pendentes: TA={0} ({1}) | servicos={2} ({3}){4}" -f $pendTA.Count, ($pendTA -join ','), $pendSV.Count, ($pendSV -join ','), $(if ($Simular) { ' [SIMULACAO]' } else { '' }))
if ($Simular) { Write-Host "SIMULACAO: dispararia NFCOM=$($pendTA.Count) e SERVICOS=$($pendSV.Count)."; exit 0 }

[System.IO.File]::WriteAllText($estP, (@{ pendentes = $todos; disparado_em = (Get-Date).ToString('s') } | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($false)))

# ---- dispara (um de cada vez, esperando terminar; a confirmacao de ambiente aparece na tela) ----
if ($pendTA.Count -gt 0) {
    Log "Disparando LANCAR_NFCOM.vbs"
    Start-Process -FilePath 'wscript.exe' -ArgumentList ('"' + (Join-Path $dir 'LANCAR_NFCOM.vbs') + '"') -WorkingDirectory $dir -Wait
    Log "LANCAR_NFCOM.vbs terminou"
}
if ($pendSV.Count -gt 0) {
    Log "Disparando LANCAR_SERVICOS.vbs"
    Start-Process -FilePath 'wscript.exe' -ArgumentList ('"' + (Join-Path $dir 'LANCAR_SERVICOS.vbs') + '"') -WorkingDirectory $dir -Wait
    Log "LANCAR_SERVICOS.vbs terminou"
}
