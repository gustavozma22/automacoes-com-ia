' =====================================================================
' Lançamento da cotação do dólar (PTAX venda) no SAP - V_TCURR
' Busca a cotação de VENDA do dia na API pública do Banco Central
' e preenche as duas linhas (BRL->USD e USD->BRL) com a data de hoje.
' =====================================================================
Option Explicit

' Identificação de navegador enviada à API do BC
Const USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"

' True = pede confirmação antes de gravar no SAP / False = grava direto
Const PEDIR_CONFIRMACAO = True

Dim hoje, dataSAP, dataAPI, cotacao
hoje    = Date
dataSAP = Pad2(Day(hoje)) & Pad2(Month(hoje)) & Year(hoje)      ' DDMMAAAA -> 07102026
dataAPI = Pad2(Month(hoje)) & "-" & Pad2(Day(hoje)) & "-" & Year(hoje) ' MM-DD-AAAA (formato da API)

' ---------------------------------------------------------------------
' 1) Buscar a cotação de venda no Banco Central
' ---------------------------------------------------------------------
cotacao = BuscarPtaxVenda(dataAPI)

If cotacao = "" Then
    MsgBox "Não há PTAX de fechamento publicada para " & FormatDateTime(hoje, 2) & "." & vbCrLf & vbCrLf & _
           "Possíveis motivos: o BC ainda não publicou o fechamento (sai por volta das 13h), " & _
           "ou hoje é fim de semana/feriado." & vbCrLf & vbCrLf & "Nada foi lançado no SAP.", _
           vbExclamation, "Cotação PTAX"
    WScript.Quit
End If

If PEDIR_CONFIRMACAO Then
    If MsgBox("Data: " & FormatDateTime(hoje, 2) & vbCrLf & _
              "PTAX venda: " & cotacao & vbCrLf & vbCrLf & _
              "Lançar no SAP?", vbYesNo + vbQuestion, "Cotação PTAX") <> vbYes Then
        WScript.Quit
    End If
End If

' ---------------------------------------------------------------------
' 2) Conectar no SAP (sessão já aberta)
' ---------------------------------------------------------------------
Dim SapGuiAuto, application, connection, session
Set SapGuiAuto  = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection  = application.Children(0)
Set session     = connection.Children(0)

' ---------------------------------------------------------------------
' 3) Preencher a tela (mesmos passos da gravação original)
' ---------------------------------------------------------------------
Const TBL = "wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/"

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00183"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").topNode = "F00200"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00183"
session.findById("wnd[0]/tbar[1]/btn[5]").press   ' Entradas novas

' Linha 0: BRL -> USD (cotação indireta)
session.findById(TBL & "ctxtV_TCURR-KURST[0,0]").text = "M"
session.findById(TBL & "ctxtV_TCURR-GDATU[1,0]").text = dataSAP
session.findById(TBL & "txtRFCU9-KURSM[2,0]").text    = cotacao
session.findById(TBL & "ctxtV_TCURR-FCURR[5,0]").text = "BRL"
session.findById(TBL & "ctxtV_TCURR-TCURR[10,0]").text = "USD"

' Linha 1: USD -> BRL (cotação direta)
session.findById(TBL & "ctxtV_TCURR-KURST[0,1]").text = "M"
session.findById(TBL & "ctxtV_TCURR-GDATU[1,1]").text = dataSAP
session.findById(TBL & "ctxtV_TCURR-FCURR[5,1]").text = "USD"
session.findById(TBL & "txtRFCU9-KURSP[7,1]").text    = cotacao
session.findById(TBL & "ctxtV_TCURR-TCURR[10,1]").text = "BRL"

session.findById("wnd[0]").sendVKey 0              ' Enter
session.findById("wnd[0]/tbar[0]/btn[11]").press   ' Gravar
session.findById("wnd[0]/tbar[0]/btn[15]").press   ' Sair

MsgBox "Cotação " & cotacao & " lançada para " & FormatDateTime(hoje, 2) & ".", vbInformation, "Cotação PTAX"


' =====================================================================
' Funções auxiliares
' =====================================================================

' Consulta a API PTAX do Banco Central e devolve a cotação de venda
' no formato do SAP (vírgula e 4 casas: "4,9935"). Vazio se não houver.
Function BuscarPtaxVenda(dataMMDDAAAA)
    Dim url, resp, re, m
    BuscarPtaxVenda = ""

    ' Mesma URL que você testou no navegador
    url = "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/" & _
          "CotacaoDolarDia(dataCotacao=@dataCotacao)?@dataCotacao='" & dataMMDDAAAA & _
          "'&$format=json"

    resp = BaixarTexto(url)

    Set re = New RegExp
    re.Pattern = """cotacaoVenda""\s*:\s*([0-9]+(\.[0-9]+)?)"
    If re.Test(resp) Then
        Set m = re.Execute(resp)(0)
        BuscarPtaxVenda = FormatarSAP(m.SubMatches(0))
    End If
End Function

' ---------------------------------------------------------------------
' Baixa o texto de uma URL tentando 3 caminhos, do mais simples ao mais
' robusto. Se todos falharem, mostra o erro detalhado de cada um.
' ---------------------------------------------------------------------
Function BaixarTexto(url)
    Dim progIds, pid, http, erros, txt
    erros = ""
    progIds = Array("MSXML2.XMLHTTP.6.0", "MSXML2.ServerXMLHTTP.6.0")

    For Each pid In progIds
        On Error Resume Next
        Err.Clear
        Set http = CreateObject(pid)
        If Err.Number = 0 Then
            http.Open "GET", url, False
            http.setRequestHeader "Cache-Control", "no-cache"
            ' O BC bloqueia (HTTP 403) chamadas que não parecem vir de um navegador
            http.setRequestHeader "User-Agent", USER_AGENT
            http.setRequestHeader "Accept", "application/json, text/plain, */*"
            http.setRequestHeader "Accept-Language", "pt-BR,pt;q=0.9"
            http.Send
        End If
        If Err.Number = 0 Then
            If http.Status = 200 Then
                BaixarTexto = http.responseText
                On Error GoTo 0
                Exit Function
            End If
            erros = erros & pid & ": HTTP " & http.Status & " " & Left(LimparHTML(http.responseText), 300) & vbCrLf
        Else
            erros = erros & pid & ": erro " & Hex(Err.Number) & " - " & Err.Description & vbCrLf
        End If
        On Error GoTo 0
    Next

    ' 3a tentativa: PowerShell (força TLS 1.2 e usa o proxy/credenciais do Windows)
    txt = BaixarViaPowerShell(url, erros)
    If txt <> "" Then
        BaixarTexto = txt
        Exit Function
    End If

    MsgBox "Não foi possível acessar a API do Banco Central." & vbCrLf & vbCrLf & _
           "Detalhes de cada tentativa:" & vbCrLf & erros & vbCrLf & _
           "Nada foi lançado no SAP.", vbCritical, "Cotação PTAX"
    WScript.Quit
End Function

' Tira as tags de uma página de erro HTML para a mensagem ficar legível
Function LimparHTML(s)
    Dim re
    Set re = New RegExp
    re.Global = True
    re.Pattern = "<[^>]+>"
    LimparHTML = re.Replace(s, " ")
    re.Pattern = "\s+"
    LimparHTML = Trim(re.Replace(LimparHTML, " "))
End Function

Function BaixarViaPowerShell(url, erros)
    Dim sh, ex, cmd, saida, errTxt
    BaixarViaPowerShell = ""
    cmd = "powershell -NoProfile -ExecutionPolicy Bypass -Command """ & _
          "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; " & _
          "$wc=New-Object Net.WebClient; " & _
          "$wc.Proxy.Credentials=[Net.CredentialCache]::DefaultCredentials; " & _
          "$wc.Encoding=[Text.Encoding]::UTF8; " & _
          "$wc.Headers.Add('User-Agent','" & USER_AGENT & "'); " & _
          "$wc.Headers.Add('Accept','application/json'); " & _
          "$wc.DownloadString('" & Replace(url, "'", "''") & "')"""

    On Error Resume Next
    Set sh = CreateObject("WScript.Shell")
    Set ex = sh.Exec(cmd)
    If Err.Number <> 0 Then
        erros = erros & "PowerShell: não foi possível executar - " & Err.Description & vbCrLf
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    saida  = ex.StdOut.ReadAll
    errTxt = ex.StdErr.ReadAll
    If InStr(saida, "cotacao") > 0 Or InStr(saida, """value""") > 0 Then
        BaixarViaPowerShell = saida
    Else
        erros = erros & "PowerShell: " & Left(Trim(errTxt & saida), 300) & vbCrLf
    End If
End Function

' "4.9935" -> "4,9935"  |  "4.99" -> "4,9900"  (independe da configuração regional)
Function FormatarSAP(valorComPonto)
    Dim partes, inteiro, decimais
    partes  = Split(valorComPonto, ".")
    inteiro = partes(0)
    If UBound(partes) > 0 Then decimais = partes(1) Else decimais = ""
    decimais = Left(decimais & "0000", 4)
    FormatarSAP = inteiro & "," & decimais
End Function

Function Pad2(n)
    Pad2 = Right("0" & n, 2)
End Function
