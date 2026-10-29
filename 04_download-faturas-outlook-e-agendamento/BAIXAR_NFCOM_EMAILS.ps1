# =====================================================================
#  BAIXAR_NFCOM_EMAILS.ps1  (v2 - 26/10/2026)
#  Copia os anexos (PDF/XML) das pastas "Faturas *" da caixa compartilhada de
#  faturas do Outlook para a pasta do mes (NFCOMs_<MES>), como hoje e feito na mao.
#  NAO move, NAO apaga e NAO altera nenhum e-mail.
#
#  Base de regras: CONSULTA\NFCOM_EMAIL_REGRAS.json, gerado da BASE_PEDIDOS
#  (coluna RECEBIMENTO DA NOTA = E-MAIL). Operadora PORTAL/manual nao entra.
#
#  v2: o anexo e salvo primeiro numa pasta temporaria; o tamanho/hash
#  REAL do arquivo e comparado com (a) tudo que ja esta em NFCOMs_<MES>
#  (inclusive BOLETOS e _REVISAR_EMAIL), (b) o NFCOMs_processed_log.json
#  (arquivo renomeado por fora/MIRO) e (c) o log deste script. Assim um
#  anexo que voce ja tratou nao volta como "_2" / "_3".
#
#  Uso: BAIXAR_NFCOM_EMAILS.bat [-Simular] [-Dias 7]
#  Requer: Outlook desktop aberto com a caixa compartilhada de faturas. Nunca abre o SAP.
# =====================================================================
param(
    [string]$Raiz = (Split-Path $PSScriptRoot -Parent),
    [int]$Dias = 7,
    [switch]$Simular
)
$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)

$meses = 'JANEIRO','FEVEREIRO','MARCO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'
$mesNome = $meses[(Get-Date).Month - 1]
$pastaMes = Join-Path $Raiz ("NFCOMs_" + $mesNome)
$pastaBol = Join-Path $pastaMes ("BOLETOS_" + $mesNome)
$pastaRev = Join-Path $pastaMes "_REVISAR_EMAIL"
$cons     = Join-Path $Raiz 'CONSULTA'
$cfgPath  = Join-Path $cons 'NFCOM_EMAIL_REGRAS.json'
$logPath  = Join-Path $cons 'NFCOM_EMAIL_baixados.json'
$runLog   = Join-Path $cons 'NFCOM_EMAIL_run_log.txt'
$procPath = Join-Path $cons 'NFCOMs_processed_log.json'
$tmp      = Join-Path ([System.IO.Path]::GetTempPath()) ("nfcom_email_" + [guid]::NewGuid().ToString('N'))

if (-not (Test-Path $cfgPath)) { throw "Nao achei $cfgPath" }
if (-not (Test-Path $pastaMes)) {
    Write-Host "A pasta $pastaMes nao existe. Crie-a e rode de novo." -ForegroundColor Yellow
    exit 1
}
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
$cfg = Get-Content $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json

# ---- log deste script (dedupe por e-mail + nome + tamanho do anexo) ----
$baixados = @()
if (Test-Path $logPath) {
    $txt = Get-Content $logPath -Raw -Encoding UTF8
    if ($txt.Trim().Length -gt 0) { $baixados = @($txt | ConvertFrom-Json) }
}
$chaves = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($b in $baixados) { [void]$chaves.Add("$($b.entry_id)|$($b.nome_original)|$($b.tamanho)") }

# ---- indice do que ja existe (tamanho REAL -> caminhos) + processed_log ----
$porTamanho = @{}
Get-ChildItem -Path $pastaMes -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
    $k = [string]$_.Length
    if (-not $porTamanho.ContainsKey($k)) { $porTamanho[$k] = New-Object System.Collections.ArrayList }
    [void]$porTamanho[$k].Add($_.FullName)
}
$tamanhosProc = New-Object 'System.Collections.Generic.HashSet[string]'
if (Test-Path $procPath) {
    $pj = Get-Content $procPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($e in $pj) { if ($e.size_bytes) { [void]$tamanhosProc.Add([string]$e.size_bytes) } }
}

function Ja-Existe($arquivoTmp) {
    $len = [string](Get-Item $arquivoTmp).Length
    if ($porTamanho.ContainsKey($len)) {
        $h = (Get-FileHash $arquivoTmp -Algorithm SHA256).Hash
        foreach ($c in $porTamanho[$len]) { if ((Get-FileHash $c -Algorithm SHA256).Hash -eq $h) { return "ja esta em " + (Split-Path $c -Leaf) } }
    }
    if ($tamanhosProc.Contains($len)) { return "tamanho ja consta no processed_log (arquivo renomeado/lancado)" }
    return $null
}
function Nome-Seguro($n) { ($n -replace '[\\/:*?"<>|]', '-').Trim() }

# ---- Outlook: achar a caixa compartilhada de faturas ----
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace('MAPI')
$raizCaixa = $null
foreach ($st in $ns.Stores) {
    if ($st.DisplayName -like "*$($cfg.caixa_compartilhada)*") { $raizCaixa = $st.GetRootFolder(); break }
}
if (-not $raizCaixa) {
    $rcp = $ns.CreateRecipient($cfg.caixa_compartilhada); [void]$rcp.Resolve()
    $raizCaixa = $ns.GetSharedDefaultFolder($rcp, 6).Parent
}
Write-Host "Caixa: $($raizCaixa.Name)"

function Get-PastasFaturas($pasta, $prefixo, $nivel) {
    foreach ($f in $pasta.Folders) {
        if ($f.Name -like "$prefixo*") { $f }
        if ($nivel -lt 2) { Get-PastasFaturas $f $prefixo ($nivel + 1) }
    }
}
$pastas = @(Get-PastasFaturas $raizCaixa $cfg.prefixo_pasta_outlook 0)
Write-Host "Pastas '$($cfg.prefixo_pasta_outlook)*' encontradas: $($pastas.Count)"

$desde = (Get-Date).AddDays(-$Dias)
$novos = @(); $semRegra = @(); $revisar = @(); $ignorados = 0; $jaExistiam = @()

foreach ($pf in $pastas) {
    $op = $null
    foreach ($o in $cfg.operadoras_email) { if ($pf.Name -match $o.regex_pasta) { $op = $o; break } }
    if (-not $op) { $semRegra += $pf.Name; continue }

    $itens = $pf.Items
    $itens.Sort('[ReceivedTime]', $true)
    foreach ($m in $itens) {
        if ($m.Class -ne 43) { continue }
        if ($m.ReceivedTime -lt $desde) { break }
        foreach ($att in $m.Attachments) {
            if ($att.Type -ne 1) { continue }
            $ext = [System.IO.Path]::GetExtension($att.FileName).ToLower()
            if ($cfg.extensoes -notcontains $ext) { $ignorados++; continue }
            $chave = "$($m.EntryID)|$($att.FileName)|$($att.Size)"
            if ($chaves.Contains($chave)) { continue }

            $nome = Nome-Seguro $att.FileName
            $arqTmp = Join-Path $tmp ([guid]::NewGuid().ToString('N') + $ext)
            $att.SaveAsFile($arqTmp)

            $motivo = Ja-Existe $arqTmp
            $reg = [ordered]@{
                entry_id = $m.EntryID; nome_original = $att.FileName; tamanho = $att.Size
                tamanho_real = (Get-Item $arqTmp).Length
                pasta_outlook = $pf.Name; operadora = $op.fornecedor_base; assunto = $m.Subject
                recebido = $m.ReceivedTime.ToString('s'); baixado_em = (Get-Date).ToString('s')
                simulado = [bool]$Simular
            }
            if ($motivo) {
                $reg.tipo = 'ja_existia'; $reg.salvo_como = $motivo
                $jaExistiam += $nome
                if (-not $Simular) { $baixados += [pscustomobject]$reg; [void]$chaves.Add($chave) }
                Remove-Item $arqTmp -Force
                continue
            }

            if     ($nome -match $cfg.regex_revisar) { $dest = $pastaRev; $tipo = 'revisar' }
            elseif ($nome -match $cfg.regex_boleto)  { $dest = $pastaBol; $tipo = 'boleto' }
            elseif ($nome -match $cfg.regex_nota)    { $dest = $pastaMes; $tipo = 'nota' }
            else                                     { $dest = $pastaRev; $tipo = 'revisar' }

            $final = Join-Path $dest $nome; $k = 2
            while (Test-Path $final) {
                $final = Join-Path $dest ([System.IO.Path]::GetFileNameWithoutExtension($nome) + "_$k" + $ext); $k++
            }
            $reg.tipo = $tipo; $reg.salvo_como = $final
            if (-not $Simular) {
                if (-not (Test-Path $dest)) { New-Item -ItemType Directory -Path $dest | Out-Null }
                Move-Item $arqTmp $final
                $baixados += [pscustomobject]$reg; [void]$chaves.Add($chave)
                $k2 = [string]$reg.tamanho_real
                if (-not $porTamanho.ContainsKey($k2)) { $porTamanho[$k2] = New-Object System.Collections.ArrayList }
                [void]$porTamanho[$k2].Add($final)
            } else { Remove-Item $arqTmp -Force }
            $novos += [pscustomobject]$reg
            if ($tipo -eq 'revisar') { $revisar += $final }
        }
    }
}
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue

if (-not $Simular) { [System.IO.File]::WriteAllText($logPath, ($baixados | ConvertTo-Json -Depth 4), $utf8) }
$linha = "[{0}] {1} dias={2} novos={3} ja_existiam={4} revisar={5} ignorados(ext)={6} pastas sem regra={7}" -f (Get-Date).ToString('s'), $(if ($Simular) { 'SIMULACAO' } else { 'EXECUCAO' }), $Dias, $novos.Count, $jaExistiam.Count, $revisar.Count, $ignorados, ($semRegra -join ', ')
Add-Content -Path $runLog -Value $linha -Encoding UTF8

Write-Host ""
Write-Host ("Anexos NOVOS: {0}   (ja existiam e foram ignorados: {1})" -f $novos.Count, $jaExistiam.Count) -ForegroundColor Green
$novos | ForEach-Object { Write-Host (" [{0}] {1}  <-  {2}" -f $_.tipo, (Split-Path $_.salvo_como -Leaf), $_.pasta_outlook) }
if ($revisar.Count -gt 0) { Write-Host "Sem classificacao automatica / NF-e / cancelamento (veja _REVISAR_EMAIL): $($revisar.Count)" -ForegroundColor Yellow }
if ($semRegra.Count -gt 0) { Write-Host "Pastas Outlook sem regra na BASE_PEDIDOS (nao baixadas): $($semRegra -join ', ')" -ForegroundColor Yellow }
if ($Simular) { Write-Host "SIMULACAO: nada foi gravado em NFCOMs_$mesNome." -ForegroundColor Cyan }
