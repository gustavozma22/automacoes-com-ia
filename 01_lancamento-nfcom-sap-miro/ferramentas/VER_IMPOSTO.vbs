' =====================================================================
'  VER_IMPOSTO.vbs  --  fotografa a tela que estiver aberta no SAP
'  Projeto de automacao com IA
' =====================================================================
'
'  SO LEITURA. Nao abre transacao, nao navega, nao digita, nao grava
'  documento. Ele apenas percorre a tela atual e escreve o que achou.
'
'  COMO USAR
'     1. No SAP, chegue ate a nota com a divergencia e abra a tela
'        NOTA FISCAL, na aba do IMPOSTO - exatamente a tela onde voce
'        faria a correcao na mao.
'     2. NAO corrija ainda. Deixe a tela aberta como esta.
'     3. Duplo clique neste arquivo.
'     4. Me mande o CAMPOS_IMPOSTO.txt que aparecer nesta pasta.
'
'  Se quiser, rode duas vezes: uma ANTES de corrigir e uma DEPOIS.
'  Renomeie o primeiro arquivo para CAMPOS_IMPOSTO_ANTES.txt entre as
'  duas rodadas. Comparando os dois eu vejo exatamente o que mudou.
' =====================================================================

Option Explicit

Dim fso, pasta, rel, SapGuiAuto, aplicacao, conexao, sessao
Dim total, comTexto

Set fso = CreateObject("Scripting.FileSystemObject")
pasta = fso.GetParentFolderName(WScript.ScriptFullName)
Set rel = fso.CreateTextFile(pasta & "\CAMPOS_IMPOSTO.txt", True, False)
total = 0 : comTexto = 0

On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
Set aplicacao  = SapGuiAuto.GetScriptingEngine
Set conexao    = aplicacao.Children(0)
Set sessao     = conexao.Children(0)
If Err.Number <> 0 Then
    rel.WriteLine "Nao consegui falar com o SAP: " & Err.Description
    rel.Close
    MsgBox "Nao consegui falar com o SAP. O SAP esta aberto e logado?", vbCritical, "NFCOM"
    WScript.Quit
End If
Err.Clear

rel.WriteLine "FOTO DA TELA DO SAP - somente leitura"
rel.WriteLine "Gerado em: " & Now
rel.WriteLine "Sistema: " & sessao.Info.SystemName & " | Mandante: " & sessao.Info.Client & " | Usuario: " & sessao.Info.User
Err.Clear
rel.WriteLine "Transacao: " & sessao.Info.Transaction
rel.WriteLine "Titulo da janela: " & sessao.findById("wnd[0]").text
Err.Clear
rel.WriteLine "Barra de status: " & TextoDe("wnd[0]/sbar")
rel.WriteLine String(74, "-")
rel.WriteLine "CAMPOS COM TEXTO (tipo | id | [conteudo])"
rel.WriteLine ""

Varrer sessao.findById("wnd[0]"), 0
Err.Clear

If Existe("wnd[1]") Then
    rel.WriteLine ""
    rel.WriteLine String(74, "-")
    rel.WriteLine "JANELA wnd[1] ABERTA:"
    Varrer sessao.findById("wnd[1]"), 0
    Err.Clear
End If

rel.WriteLine ""
rel.WriteLine String(74, "-")
rel.WriteLine "Objetos percorridos: " & total & " | com texto: " & comTexto
rel.WriteLine ""
rel.WriteLine "Nenhuma tela foi alterada por este script."
rel.Close

MsgBox "Pronto. " & total & " objetos, " & comTexto & " com texto." & vbCrLf & vbCrLf & _
       "Nada foi alterado no SAP." & vbCrLf & _
       "Me mande o CAMPOS_IMPOSTO.txt desta pasta.", vbInformation, "NFCOM"
WScript.Quit

' =====================================================================
Sub Varrer(ByVal obj, ByVal prof)
    Dim k, alvo, cnt, t, tipo, id
    If prof > 16 Then Exit Sub
    On Error Resume Next

    id = ""
    id = obj.Id
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    total = total + 1

    tipo = "?"
    tipo = obj.Type
    Err.Clear

    t = ""
    t = obj.Text
    Err.Clear

    If Trim(t) <> "" Then
        comTexto = comTexto + 1
        rel.WriteLine "  " & tipo & " | " & Curto(id) & " | [" & t & "]"
    End If

    cnt = -1
    cnt = obj.Children.Count
    Err.Clear
    If cnt < 1 Then Exit Sub

    For k = 0 To cnt - 1
        Set alvo = PegarFilho(obj, k)
        If Not (alvo Is Nothing) Then Varrer alvo, prof + 1
        Err.Clear
    Next
End Sub

Function PegarFilho(pai, indice)
    Set PegarFilho = Nothing
    On Error Resume Next
    Set PegarFilho = pai.Children.ElementAt(indice)
    If Err.Number = 0 And Not (PegarFilho Is Nothing) Then Err.Clear : Exit Function
    Err.Clear
    Set PegarFilho = pai.Children.Item(indice)
    If Err.Number = 0 And Not (PegarFilho Is Nothing) Then Err.Clear : Exit Function
    Err.Clear
    Set PegarFilho = pai.Children(indice)
    If Err.Number <> 0 Then Set PegarFilho = Nothing
    Err.Clear
End Function

Function TextoDe(x)
    TextoDe = ""
    On Error Resume Next
    TextoDe = sessao.findById(x).Text
    Err.Clear
End Function

Function Curto(x)
    Curto = Replace(x, "/app/con[0]/ses[0]/", "")
End Function

Function Existe(x)
    Dim w
    Existe = False
    On Error Resume Next
    Set w = sessao.findById(x)
    If Err.Number = 0 Then Existe = True
    Err.Clear
End Function
