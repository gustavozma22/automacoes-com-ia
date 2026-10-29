' =====================================================================
'  LANCAR_TUDO.vbs
'
'  Roda os dois executores em sequencia, sobre a MESMA planilha de lote
'  (LANCAMENTO_NFCOM.xlsx, aba LOTE):
'
'     1. LANCAR_NFCOM.vbs     -> linhas com TIPO DE NF = TA
'     2. LANCAR_SERVICOS.vbs  -> linhas com TIPO DE NF = ZH / RECIBO ...
'
'  Cada um pula, sem escrever nada, as linhas que nao sao dele. Por isso
'  a ordem nao importa e rodar duas vezes nao duplica: o que ja tem
'  STATUS preenchido nenhum dos dois toca.
'
'  A confirmacao de ambiente (PRD/QAS) continua aparecendo DUAS vezes,
'  uma por executor. Isso e de proposito: e a unica trava que separa um
'  teste no QAS de um documento real no PRD, e ela nao se terceiriza.
' =====================================================================
Option Explicit

Const TITULO = "Lancar tudo - NFCOM + NFS-e"

Dim fso, shell, pastaBase, alvo, i, nomes, codigo, resumo
Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
pastaBase = fso.GetParentFolderName(WScript.ScriptFullName)

nomes = Array("LANCAR_NFCOM.vbs", "LANCAR_SERVICOS.vbs")

' --- os dois existem? ------------------------------------------------
For i = 0 To UBound(nomes)
    If Not fso.FileExists(pastaBase & "\" & nomes(i)) Then
        MsgBox "Nao achei o " & nomes(i) & " nesta pasta." & vbCrLf & vbCrLf & _
               pastaBase, 16, TITULO
        WScript.Quit
    End If
Next

' --- a planilha esta fechada? ----------------------------------------
If Not fso.FileExists(pastaBase & "\LANCAMENTO_NFCOM.xlsx") Then
    MsgBox "Nao achei a LANCAMENTO_NFCOM.xlsx nesta pasta.", 16, TITULO
    WScript.Quit
End If
If fso.FileExists(pastaBase & "\~$LANCAMENTO_NFCOM.xlsx") Then
    MsgBox "A LANCAMENTO_NFCOM.xlsx esta aberta no Excel." & vbCrLf & vbCrLf & _
           "Feche a planilha e rode de novo.", 16, TITULO
    WScript.Quit
End If

If MsgBox("Vou rodar os dois executores em sequencia, na mesma planilha:" & vbCrLf & vbCrLf & _
          "   1) NFCOM    (linhas TIPO = TA)" & vbCrLf & _
          "   2) NFS-e    (linhas TIPO = ZH / RECIBO ...)" & vbCrLf & vbCrLf & _
          "Cada um vai pedir a confirmacao de ambiente separadamente." & vbCrLf & _
          "O SAP precisa estar aberto e logado." & vbCrLf & vbCrLf & _
          "Continuar?", 4 + 32, TITULO) <> 6 Then
    WScript.Quit
End If

' --- roda um de cada vez, esperando terminar -------------------------
resumo = ""
For i = 0 To UBound(nomes)
    alvo = pastaBase & "\" & nomes(i)
    codigo = -1
    On Error Resume Next
    codigo = shell.Run("wscript.exe """ & alvo & """", 1, True)
    If Err.Number <> 0 Then
        resumo = resumo & nomes(i) & ": NAO consegui iniciar (" & Err.Description & ")" & vbCrLf
        Err.Clear
    Else
        resumo = resumo & nomes(i) & ": terminou" & vbCrLf
    End If
    On Error GoTo 0
Next

MsgBox "Sequencia encerrada." & vbCrLf & vbCrLf & resumo & vbCrLf & _
       "O resultado de cada nota esta na aba LOTE, nas colunas" & vbCrLf & _
       "STATUS / No DOC MIRO / ERRO. Volte ao Claude e peca" & vbCrLf & _
       "para importar o retorno.", 64, TITULO
