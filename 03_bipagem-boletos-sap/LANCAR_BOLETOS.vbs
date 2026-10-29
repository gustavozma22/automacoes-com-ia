' =====================================================================
'  LANCAR_BOLETOS.vbs   -   "bipagem" dos boletos
'
'  Terceiro executor do processo. Nao lanca nota nenhuma: ele volta
'  nos documentos JA LANCADOS pelo LANCAR_NFCOM.vbs / LANCAR_SERVICOS.vbs
'  e preenche, na linha do fornecedor do documento contabil:
'
'      forma de pagamento  (BSEG-ZLSCH)   <- coluna AN da planilha
'      codigo de barras    (RF05L-BRCDE)  <- coluna AP da planilha
'
'  Caminho percorrido para cada linha (descoberto pelas sondas
'  VER_BOLETO1..6 e registrado em MAPA_CAMPOS_BOLETO.txt):
'
'    numero da MIRO (coluna S) + ano da DATA DA FATURA
'      -> FB03 > Lista de documentos > BUKRS / GJAHR / AWKEY
'         (AWKEY = MIRO & ano, chave de origem 1:1)
'      -> a lista TEM de voltar exatamente 1 linha  ->  numero do FI
'      -> FB02 nesse FI
'      -> na grade de itens, a linha com chave de lancamento 31 ou 21
'      -> abre a linha, grava ZLSCH e BRCDE, salva
'      -> reabre em FB03 e LE os dois campos de volta como prova
'
'  TRAVAS (nenhum documento e tocado se alguma falhar):
'    - so processa linha com STATUS = LANCADA e numero de MIRO preenchido
'    - so processa linha com STATUS BIPAGEM (coluna AR) ainda vazio
'    - a busca por AWKEY tem de devolver exatamente 1 documento
'    - a grade do documento tem de ter exatamente 1 linha de fornecedor
'    - o montante dessa linha tem de bater com o MONTANTE da planilha
'    - o codigo de barras tem de ter 44 digitos, DV mod-11 valido, e o
'      valor embutido nas posicoes 10-19 tem de bater com o MONTANTE
'    - boleto de arrecadacao/concessionaria (comeca com 8, forma O):
'      DV geral na posicao 4 (mod 10 ou 11 conforme a posicao 3), valor
'      nas posicoes 5-15; linha digitavel de 48 digitos tambem aceita
'    - forma de pagamento sem boleto (S) recebe SO o ZLSCH
'
'  COMO USAR
'     1. Abrir o SAP e fazer login no ambiente certo.
'     2. FECHAR a planilha LANCAMENTO_NFCOM.xlsx no Excel.
'     3. Deixar planilha, MAPA_CAMPOS_BOLETO.txt e este arquivo na
'        mesma pasta.
'     4. Dois cliques neste arquivo.
'     5. Reabrir a planilha e conferir as colunas AQ..AT.
' =====================================================================
Option Explicit

' ---------- colunas da aba LOTE ----------
Const C_ID         = 1
Const C_EMPRESA    = 2
Const C_DTFATURA   = 3
Const C_MONTANTE   = 4
Const C_REFER      = 5
Const C_TIPONF     = 6
Const C_STATUS     = 18
Const C_DOCMIRO    = 19
Const C_FORNEC     = 23
Const C_FORMA      = 40        ' FORMA PGTO (ZLSCH)
Const C_LINHADIG   = 41        ' LINHA DIGITAVEL (47)
Const C_BARRAS     = 42        ' CODIGO DE BARRAS (44)
Const C_DOCCONT    = 43        ' DOC CONTABIL (FI)
Const C_BIPSTATUS  = 44        ' STATUS BIPAGEM
Const C_BIPERRO    = 45        ' ERRO BIPAGEM
Const C_BIPQUANDO  = 46        ' QUANDO BIPADO

Const MAX_LINHAS = 400
Const ForReading = 1
Const CENTAVOS_INVALIDO = -999999999
Const TITULO = "BIPAGEM"

Dim fso, pasta, arquivo, mapa
Dim excel, wb, ws
Dim SapGuiAuto, aplicacao, conexao, sessao
Dim quemRodou, quemWindows
Dim gradeAchada
Dim totalOK, totalNOK, pulados, r

Set fso = CreateObject("Scripting.FileSystemObject")
pasta   = fso.GetParentFolderName(WScript.ScriptFullName)
arquivo = pasta & "\LANCAMENTO_NFCOM.xlsx"

quemWindows = ""
On Error Resume Next
quemWindows = CreateObject("WScript.Network").UserName
Err.Clear
On Error GoTo 0

' ---------------------------------------------------------------------
' 1. Mapa de campos
' ---------------------------------------------------------------------
Set mapa = CreateObject("Scripting.Dictionary")
If Not CarregarMapa(pasta & "\MAPA_CAMPOS_BOLETO.txt", mapa) Then
    MsgBox "MAPA_CAMPOS_BOLETO.txt nao encontrado ou vazio nesta pasta.", vbCritical, TITULO
    WScript.Quit
End If

' ---------------------------------------------------------------------
' 2. Planilha
' ---------------------------------------------------------------------
If Not fso.FileExists(arquivo) Then
    MsgBox "LANCAMENTO_NFCOM.xlsx nao encontrado nesta pasta.", vbCritical, TITULO
    WScript.Quit
End If

On Error Resume Next
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Nao consegui abrir o Excel nesta maquina.", vbCritical, TITULO
    WScript.Quit
End If
excel.Visible = False
excel.DisplayAlerts = False
Set wb = excel.Workbooks.Open(arquivo)
If Err.Number <> 0 Then
    MsgBox "Nao consegui abrir a planilha: " & Err.Description, vbCritical, TITULO
    excel.Quit
    WScript.Quit
End If
Set ws = wb.Worksheets("LOTE")
If Err.Number <> 0 Then
    MsgBox "A planilha nao tem uma aba chamada LOTE.", vbCritical, TITULO
    wb.Close False
    excel.Quit
    WScript.Quit
End If
Err.Clear
On Error GoTo 0

If wb.ReadOnly Then
    MsgBox "A planilha abriu somente para leitura - provavelmente esta aberta no Excel." & vbCrLf & vbCrLf & _
           "Feche LANCAMENTO_NFCOM.xlsx e rode de novo.", vbCritical, TITULO
    wb.Close False
    excel.Quit
    WScript.Quit
End If

' ---------------------------------------------------------------------
' 3. Sessao SAP ja aberta
' ---------------------------------------------------------------------
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
    Encerrar "Nao encontrei o SAP GUI aberto." & vbCrLf & vbCrLf & _
             "Abra o SAP, faca login e rode este arquivo de novo."
End If
Set aplicacao = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Then
    Encerrar "O SAP GUI Scripting esta desabilitado nesta maquina ou neste servidor."
End If
Set conexao = aplicacao.Children(0)
Set sessao  = conexao.Children(0)
If Err.Number <> 0 Then
    Encerrar "Nenhuma sessao SAP ativa. Faca login no SAP primeiro."
End If
Err.Clear

Dim sapSistema, sapMandante, sapUsuario
sapSistema  = ""
sapMandante = ""
sapUsuario  = ""
sapSistema  = sessao.Info.SystemName
sapMandante = sessao.Info.Client
sapUsuario  = sessao.Info.User
Err.Clear
On Error GoTo 0

' --- ambiente ---------------------------------------------------------
' Este script ALTERA documento contabil ja lancado. A confirmacao de
' ambiente e sempre obrigatoria: a planilha nao sabe em qual sistema o
' lote foi lancado, e so a sessao aberta decide onde a gravacao cai.
Dim ambienteAtual, listaAmb, ambientePermitido, nAmb, itemAmb
ambienteAtual = UCase(Trim(sapSistema)) & "/" & Trim(sapMandante)
ambientePermitido = False

If Trim(mapa("SAP_AMBIENTES_PERMITIDOS")) <> "" Then
    listaAmb = Split(mapa("SAP_AMBIENTES_PERMITIDOS"), "|")
    For nAmb = 0 To UBound(listaAmb)
        itemAmb = UCase(Replace(Trim(listaAmb(nAmb)), " ", ""))
        If itemAmb = ambienteAtual Then
            ambientePermitido = True
        End If
    Next
    If Not ambientePermitido Then
        Encerrar "A sessao SAP aberta nao e um dos ambientes liberados." & vbCrLf & vbCrLf & _
                 "Liberados: " & mapa("SAP_AMBIENTES_PERMITIDOS") & vbCrLf & _
                 "Encontrado: " & ambienteAtual & vbCrLf & vbCrLf & _
                 "Nada foi alterado."
    End If
End If

If MsgBox("A BIPAGEM ALTERA DOCUMENTOS JA LANCADOS." & vbCrLf & vbCrLf & _
          "        >>>  " & ambienteAtual & "  <<<" & vbCrLf & vbCrLf & _
          "  Usuario: " & sapUsuario & vbCrLf & vbCrLf & _
          "Se aqui estiver PRD, a alteracao e real e fica no historico" & vbCrLf & _
          "de modificacoes do documento, em nome desse usuario." & vbCrLf & vbCrLf & _
          "Continuar?", vbYesNo + vbExclamation, TITULO & " - confirmar ambiente") <> vbYes Then
    Encerrar "Cancelado pelo usuario. Nada foi alterado."
End If

quemRodou = sapUsuario & "@" & sapSistema & "/" & sapMandante & " (win: " & quemWindows & ")"

On Error Resume Next
sessao.findById("wnd[0]").maximize
Err.Clear
On Error GoTo 0

' ---------------------------------------------------------------------
' 4. Percorre as linhas
' ---------------------------------------------------------------------
totalOK  = 0
totalNOK = 0
pulados  = 0

For r = 2 To MAX_LINHAS + 1
    If Celula(r, C_ID) = "" Then
        Exit For
    End If
    If Celula(r, C_BIPSTATUS) <> "" Then
        pulados = pulados + 1          ' ja bipada (ou retida) antes
    ElseIf UCase(Celula(r, C_STATUS)) <> "LANCADA" Then
        pulados = pulados + 1          ' nao lancada / lancada manualmente
    ElseIf Celula(r, C_DOCMIRO) = "" Then
        pulados = pulados + 1          ' sem numero de MIRO nao ha o que achar
    Else
        ProcessarLinha r
    End If
Next

wb.Save
wb.Close True
excel.Quit

MsgBox "Bipagem finalizada." & vbCrLf & vbCrLf & _
       "Bipadas: " & totalOK & vbCrLf & _
       "Retidas para conferencia: " & totalNOK & vbCrLf & _
       "Ignoradas (sem MIRO, nao lancadas ou ja bipadas): " & pulados & vbCrLf & vbCrLf & _
       "Confira as colunas AQ a AT da planilha.", vbInformation, TITULO

' =====================================================================
'  UMA LINHA DA PLANILHA
' =====================================================================
Sub ProcessarLinha(lin)
    Dim id, bukrs, miro, ano, montanteCent, forma, barras, linhaDig
    Dim docFI, erro, obs, i
    Dim semBoleto, valorBoletoCent, lidoZL, lidoBC

    id    = Celula(lin, C_ID)
    bukrs = Celula(lin, C_EMPRESA)
    miro  = SoDigitos(Celula(lin, C_DOCMIRO))
    ano   = AnoDaData(Celula(lin, C_DTFATURA))
    forma = UCase(Trim(Celula(lin, C_FORMA)))
    obs   = ""

    montanteCent = EmCentavos(Celula(lin, C_MONTANTE))

    ' ---------- conferencias que nem encostam no SAP ----------
    If bukrs = "" Then
        Resultado lin, "NAO BIPADA", "", "Linha sem EMPRESA."
        Exit Sub
    End If
    If Len(miro) < 8 Then
        Resultado lin, "NAO BIPADA", "", "Numero de MIRO invalido: [" & Celula(lin, C_DOCMIRO) & "]"
        Exit Sub
    End If
    If Len(ano) <> 4 Then
        Resultado lin, "NAO BIPADA", "", "Nao consegui tirar o ano da DATA DA FATURA: [" & Celula(lin, C_DTFATURA) & "]"
        Exit Sub
    End If
    If montanteCent = CENTAVOS_INVALIDO Or montanteCent <= 0 Then
        Resultado lin, "NAO BIPADA", "", "MONTANTE invalido: [" & Celula(lin, C_MONTANTE) & "]"
        Exit Sub
    End If
    If forma = "" Then
        Resultado lin, "NAO BIPADA", "", "Coluna FORMA PGTO (ZLSCH) vazia - preencha a partir da planilha BASE_PEDIDOS."
        Exit Sub
    End If
    If Not FormaValida(forma) Then
        Resultado lin, "NAO BIPADA", "", "Forma de pagamento [" & forma & "] fora da lista aceita (" & FormasAceitas() & ")."
        Exit Sub
    End If

    semBoleto = FormaSemBoleto(forma)

    barras   = SoDigitos(Celula(lin, C_BARRAS))
    linhaDig = SoDigitos(Celula(lin, C_LINHADIG))

    ' Boleto que ainda nao chegou nao e erro: e so a linha nao estar
    ' pronta. Sai SEM escrever nada, para ela continuar na fila da
    ' proxima rodada - se marcasse NAO BIPADA aqui, a trava da coluna
    ' AR trancaria a linha ate alguem limpar a mao.
    If Not semBoleto Then
        If barras = "" And linhaDig = "" Then
            pulados = pulados + 1
            Exit Sub
        End If
    End If

    If Not semBoleto Then
        ' sem codigo de barras mas com linha digitavel: converte aqui e
        ' grava na planilha, para a prova ficar registrada
        If barras = "" And (Len(linhaDig) = 47 Or Len(linhaDig) = 48) Then
            erro = ""
            barras = LinhaParaBarra(linhaDig, erro)
            If barras = "" Then
                Resultado lin, "NAO BIPADA", "", "Linha digitavel invalida: " & erro
                Exit Sub
            End If
            GravarTexto lin, C_BARRAS, barras
            obs = obs & "Codigo de barras derivado da linha digitavel. "
        End If

        If Len(barras) <> 44 Then
            Resultado lin, "NAO BIPADA", "", "Codigo de barras deve ter 44 digitos; veio com " & Len(barras) & ". Preencha a coluna AP ou AO."
            Exit Sub
        End If
        If Not DVBarraOK(barras) Then
            Resultado lin, "NAO BIPADA", "", "Digito verificador do codigo de barras nao confere: " & barras
            Exit Sub
        End If

        valorBoletoCent = ValorDaBarra(barras)
        If valorBoletoCent = 0 Then
            obs = obs & "Boleto sem valor embutido - nao deu para conferir contra o MONTANTE. "
        ElseIf valorBoletoCent <> montanteCent Then
            Resultado lin, "NAO BIPADA", "", "O valor dentro do codigo de barras (" & Formatar(valorBoletoCent) & _
                     ") nao bate com o MONTANTE da planilha (" & Formatar(montanteCent) & "). Boleto da nota errada?"
            Exit Sub
        End If
    End If

    ' ---------- passo 1: MIRO -> documento contabil ----------
    docFI = ""
    erro  = ""
    docFI = AcharDocumentoFI(bukrs, ano, miro, erro)
    If docFI = "" Then
        Resultado lin, "NAO BIPADA", "", erro
        Exit Sub
    End If
    GravarTexto lin, C_DOCCONT, docFI

    ' ---------- passo 2: abre em FB02 e acha a linha do fornecedor ----------
    erro = ""
    If Not AbrirDocumento(mapa("TCODE_ALTERA"), docFI, bukrs, ano, erro) Then
        Resultado lin, "NAO BIPADA", docFI, erro
        Exit Sub
    End If

    Dim nLinha, montanteLinha
    nLinha = -1
    montanteLinha = CENTAVOS_INVALIDO
    erro = ""
    If Not AcharLinhaFornecedor(nLinha, montanteLinha, erro) Then
        Resultado lin, "NAO BIPADA", docFI, erro
        Exit Sub
    End If

    If montanteLinha <> CENTAVOS_INVALIDO Then
        If Abs(montanteLinha) <> montanteCent Then
            Resultado lin, "NAO BIPADA", docFI, "A linha do fornecedor do documento " & docFI & " vale " & _
                     Formatar(Abs(montanteLinha)) & " e a planilha diz " & Formatar(montanteCent) & _
                     ". Documento errado - nada foi alterado."
            Exit Sub
        End If
    End If

    ' ---------- passo 3: abre a linha e grava ----------
    On Error Resume Next
    gradeAchada.setCurrentCell nLinha, mapa("LINHAS_COL_CHAVE")
    gradeAchada.doubleClickCurrentCell
    If Err.Number <> 0 Then
        Err.Clear
        Resultado lin, "NAO BIPADA", docFI, "Nao consegui abrir a linha do fornecedor do documento " & docFI & "."
        Exit Sub
    End If
    On Error GoTo 0

    lidoZL = LerCampo(mapa("ITEM_ZLSCH"))
    lidoBC = SoDigitos(LerCampo(mapa("ITEM_BRCDE")))

    Dim precisaGravar
    precisaGravar = False
    If UCase(Trim(lidoZL)) <> forma Then
        precisaGravar = True
    End If
    If Not semBoleto Then
        If lidoBC <> barras Then
            precisaGravar = True
        End If
    End If

    If Not precisaGravar Then
        VoltarAoMenu
        Resultado lin, "JA ESTAVA OK", docFI, Trim(obs & "Documento ja tinha os dois campos preenchidos com estes valores.")
        Exit Sub
    End If

    If Trim(lidoBC) <> "" And Not semBoleto Then
        If lidoBC <> barras Then
            obs = obs & "Codigo de barras anterior substituido: " & lidoBC & ". "
        End If
    End If

    erro = ""
    If Not Definir(mapa("ITEM_ZLSCH"), forma, erro) Then
        VoltarAoMenu
        Resultado lin, "NAO BIPADA", docFI, "Nao consegui escrever a forma de pagamento: " & erro
        Exit Sub
    End If
    If Not semBoleto Then
        erro = ""
        If Not Definir(mapa("ITEM_BRCDE"), barras, erro) Then
            VoltarAoMenu
            Resultado lin, "NAO BIPADA", docFI, "Nao consegui escrever o codigo de barras: " & erro
            Exit Sub
        End If

        ' Corrigido em 09/10/2026, com base numa gravacao de tela de uma
        ' tentativa manual bem-sucedida (nota de exemplo 1234-56, doc
        ' 5100000456): o SAP mostra o aviso "Vencimento do codigo de
        ' barras nao confere com valor do Titulo" assim que o campo de
        ' codigo de barras e confirmado - ANTES de chegar no botao
        ' Gravar, nao depois. Apertar Enter aqui, duas vezes (igual a
        ' gravacao manual), confirma esse aviso na hora; sem isso o
        ' Gravar que vem a seguir salva com o campo ZLSCH voltando
        ' vazio (foi o que aconteceu com as 4 notas do fornecedor em
        ' 08/10, mesmo tentando confirmar DEPOIS de gravar). O vencimento
        ' do documento ja foi conferido manualmente no FB03 e bate com o
        ' boleto - esse aviso e um falso positivo do SAP, nao um erro de
        ' dado real.
        On Error Resume Next
        sessao.findById("wnd[0]").sendVKey 0
        sessao.findById("wnd[0]").sendVKey 0
        Err.Clear
        On Error GoTo 0
    End If

    ' ---------- passo 4: salvar ----------
    Dim tipoMsg, txtMsg, txtPopup
    On Error Resume Next
    sessao.findById(mapa("BOTAO_SALVAR")).press
    If Err.Number <> 0 Then
        Err.Clear
        VoltarAoMenu
        Resultado lin, "NAO BIPADA", docFI, "Nao consegui acionar o botao de gravar."
        Exit Sub
    End If
    On Error GoTo 0

    ' popup de confirmacao, quando houver
    txtPopup = ""
    If ExisteJanela(1) Then
        txtPopup = TextoDaJanela(1)
        On Error Resume Next
        sessao.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Err.Number <> 0 Then
            Err.Clear
            sessao.findById("wnd[1]/tbar[0]/btn[0]").press
            Err.Clear
        End If
        On Error GoTo 0
        If Trim(txtPopup) <> "" Then
            obs = obs & "Popup ao gravar: " & txtPopup & " "
        End If
    End If

    ' Enter final depois de gravar (igual a gravacao manual) - limpa
    ' qualquer residuo de tela antes de ler a barra de status.
    On Error Resume Next
    sessao.findById("wnd[0]").sendVKey 0
    Err.Clear
    On Error GoTo 0

    tipoMsg = TipoDaBarra()
    txtMsg  = TextoDaBarra()
    If tipoMsg = "E" Or tipoMsg = "A" Then
        VoltarAoMenu
        Resultado lin, "NAO BIPADA", docFI, "O SAP recusou a gravacao (" & tipoMsg & "): " & txtMsg
        Exit Sub
    End If
    If Trim(txtMsg) <> "" Then
        obs = obs & "SAP: " & Trim(txtMsg) & " "
    End If

    VoltarAoMenu

    ' ---------- passo 5: prova - reabre em FB03 e le de volta ----------
    If UCase(Left(Trim(mapa("CONFERIR_APOS_SALVAR") & "S"), 1)) = "S" Then
        Dim provaZL, provaBC
        provaZL = ""
        provaBC = ""
        erro = ""
        If Not LerDeVolta(docFI, bukrs, ano, provaZL, provaBC, erro) Then
            Resultado lin, "BIPADA SEM CONFERENCIA", docFI, Trim(obs & "Gravou, mas nao consegui reabrir para conferir: " & erro)
            Exit Sub
        End If
        If UCase(Trim(provaZL)) <> forma Then
            Resultado lin, "NAO BIPADA", docFI, Trim(obs & "Depois de gravar, a forma de pagamento no documento esta [" & provaZL & "] e deveria ser [" & forma & "]. CONFIRA O DOCUMENTO " & docFI & ".")
            Exit Sub
        End If
        If Not semBoleto Then
            If SoDigitos(provaBC) <> barras Then
                Resultado lin, "NAO BIPADA", docFI, Trim(obs & "Depois de gravar, o codigo de barras no documento esta [" & provaBC & "] e deveria ser [" & barras & "]. CONFIRA O DOCUMENTO " & docFI & ".")
                Exit Sub
            End If
        End If
        obs = obs & "Conferido no documento depois de gravar. "
    End If

    Resultado lin, "BIPADA", docFI, Trim(obs)
End Sub

' =====================================================================
'  PASSO 1 - MIRO  ->  documento contabil
' =====================================================================
Function AcharDocumentoFI(bukrs, ano, miro, erro)
    Dim awkey, nLin, doc, lixo
    AcharDocumentoFI = ""
    erro = ""
    awkey = miro & ano

    On Error Resume Next
    sessao.findById("wnd[0]/tbar[0]/okcd").text = "/n" & mapa("TCODE_BUSCA")
    sessao.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        Err.Clear
        erro = "Nao consegui abrir a transacao " & mapa("TCODE_BUSCA") & "."
        Exit Function
    End If
    sessao.findById(mapa("BUSCA_BOTAO_LISTA")).press
    If Err.Number <> 0 Then
        Err.Clear
        erro = "Nao achei o botao de lista de documentos em " & mapa("TCODE_BUSCA") & "."
        Exit Function
    End If
    On Error GoTo 0

    lixo = ""
    If Not Definir(mapa("BUSCA_EMPRESA"), bukrs, lixo) Then
        erro = "Nao consegui preencher a empresa na busca: " & lixo
        Exit Function
    End If
    If Not Definir(mapa("BUSCA_EXERCICIO"), ano, lixo) Then
        erro = "Nao consegui preencher o exercicio na busca: " & lixo
        Exit Function
    End If
    If Not Definir(mapa("BUSCA_AWKEY"), awkey, lixo) Then
        erro = "Nao consegui preencher a chave de origem na busca: " & lixo
        Exit Function
    End If
    ' limpa filtros que possam ter sobrado da execucao anterior
    LimparCampo mapa("BUSCA_BELNR")
    LimparCampo mapa("BUSCA_XBLNR")

    On Error Resume Next
    sessao.findById("wnd[0]").sendVKey 8
    If Err.Number <> 0 Then
        Err.Clear
        erro = "Nao consegui executar a busca."
        Exit Function
    End If
    On Error GoTo 0

    If Not AcharGradeNaTela() Then
        erro = "A busca por chave de origem " & awkey & " nao devolveu nenhum documento contabil. " & TextoDaBarra()
        VoltarAoMenu
        Exit Function
    End If

    nLin = 0
    On Error Resume Next
    nLin = gradeAchada.RowCount
    Err.Clear
    On Error GoTo 0

    If nLin <> 1 Then
        erro = "A busca por chave de origem " & awkey & " devolveu " & nLin & " documentos. Esperado exatamente 1 - nada foi alterado."
        VoltarAoMenu
        Exit Function
    End If

    doc = Trim(Ler(gradeAchada, 0, mapa("BUSCA_COL_DOC")))
    VoltarAoMenu

    If SoDigitos(doc) = "" Then
        erro = "Achei a linha na lista mas nao consegui ler o numero do documento contabil."
        Exit Function
    End If
    AcharDocumentoFI = SoDigitos(doc)
End Function

' =====================================================================
'  PASSO 2 - abre o documento (FB02 para alterar, FB03 para conferir)
' =====================================================================
Function AbrirDocumento(tcode, docFI, bukrs, ano, erro)
    Dim lixo
    AbrirDocumento = False
    erro = ""

    On Error Resume Next
    sessao.findById("wnd[0]/tbar[0]/okcd").text = "/n" & tcode
    sessao.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        Err.Clear
        erro = "Nao consegui abrir a transacao " & tcode & "."
        Exit Function
    End If
    On Error GoTo 0

    lixo = ""
    If Not Definir(mapa("DOC_BELNR"), docFI, lixo) Then
        erro = "Nao consegui preencher o numero do documento em " & tcode & ": " & lixo
        Exit Function
    End If
    If Not Definir(mapa("DOC_BUKRS"), bukrs, lixo) Then
        erro = "Nao consegui preencher a empresa em " & tcode & ": " & lixo
        Exit Function
    End If
    If Not Definir(mapa("DOC_GJAHR"), ano, lixo) Then
        erro = "Nao consegui preencher o exercicio em " & tcode & ": " & lixo
        Exit Function
    End If

    On Error Resume Next
    sessao.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        Err.Clear
        erro = "Nao consegui entrar no documento " & docFI & " em " & tcode & "."
        Exit Function
    End If
    On Error GoTo 0

    If ExisteJanela(1) Then
        erro = "O SAP abriu um aviso ao entrar no documento " & docFI & ": " & TextoDaJanela(1)
        FecharPopups
        VoltarAoMenu
        Exit Function
    End If

    If TipoDaBarra() = "E" Or TipoDaBarra() = "A" Then
        erro = "O SAP recusou abrir o documento " & docFI & " em " & tcode & ": " & TextoDaBarra()
        VoltarAoMenu
        Exit Function
    End If

    AbrirDocumento = True
End Function

' =====================================================================
'  PASSO 3 - a linha do fornecedor dentro do documento
' =====================================================================
Function AcharLinhaFornecedor(nLinha, montanteCent, erro)
    Dim nLin, l, chave, achadas, primeira
    AcharLinhaFornecedor = False
    erro = ""
    nLinha = -1
    montanteCent = CENTAVOS_INVALIDO

    If Not AcharGradeNaTela() Then
        erro = "Nao achei a grade de itens do documento."
        VoltarAoMenu
        Exit Function
    End If

    nLin = 0
    On Error Resume Next
    nLin = gradeAchada.RowCount
    Err.Clear
    On Error GoTo 0

    achadas = 0
    primeira = -1
    For l = 0 To nLin - 1
        chave = Trim(Ler(gradeAchada, l, mapa("LINHAS_COL_CHAVE")))
        If EhChaveFornecedor(chave) Then
            achadas = achadas + 1
            If primeira < 0 Then
                primeira = l
            End If
        End If
    Next

    If achadas = 0 Then
        erro = "O documento nao tem nenhuma linha de fornecedor (chave " & mapa("CHAVES_FORNECEDOR") & ")."
        VoltarAoMenu
        Exit Function
    End If
    If achadas > 1 Then
        erro = "O documento tem " & achadas & " linhas de fornecedor. Esperado exatamente 1 - nada foi alterado."
        VoltarAoMenu
        Exit Function
    End If

    nLinha = primeira
    montanteCent = EmCentavos(Ler(gradeAchada, primeira, mapa("LINHAS_COL_MONTANTE")))
    AcharLinhaFornecedor = True
End Function

' =====================================================================
'  PASSO 5 - prova: reabre em FB03 e le os dois campos
' =====================================================================
Function LerDeVolta(docFI, bukrs, ano, valZL, valBC, erro)
    Dim nLinha, montanteLinha
    LerDeVolta = False
    valZL = ""
    valBC = ""
    erro = ""

    If Not AbrirDocumento(mapa("TCODE_BUSCA"), docFI, bukrs, ano, erro) Then
        Exit Function
    End If

    nLinha = -1
    montanteLinha = CENTAVOS_INVALIDO
    If Not AcharLinhaFornecedor(nLinha, montanteLinha, erro) Then
        Exit Function
    End If

    On Error Resume Next
    gradeAchada.setCurrentCell nLinha, mapa("LINHAS_COL_CHAVE")
    gradeAchada.doubleClickCurrentCell
    If Err.Number <> 0 Then
        Err.Clear
        erro = "nao consegui abrir a linha do fornecedor na conferencia"
        VoltarAoMenu
        Exit Function
    End If
    On Error GoTo 0

    valZL = LerCampo(mapa("ITEM_ZLSCH"))
    valBC = LerCampo(mapa("ITEM_BRCDE"))
    VoltarAoMenu
    LerDeVolta = True
End Function

' =====================================================================
'  PLANILHA
' =====================================================================
Function Celula(lin, col)
    Dim v
    Celula = ""
    On Error Resume Next
    v = ws.Cells(lin, col).Value
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    If IsNull(v) Or IsEmpty(v) Then
        Exit Function
    End If
    Celula = Trim(CStr(v))
End Function

Sub GravarTexto(lin, col, valor)
    On Error Resume Next
    ws.Cells(lin, col).NumberFormat = "@"
    ws.Cells(lin, col).Value = valor
    Err.Clear
    On Error GoTo 0
End Sub

Sub Resultado(lin, status, docFI, erro)
    On Error Resume Next
    ws.Cells(lin, C_BIPSTATUS).Value = status
    ws.Cells(lin, C_BIPERRO).Value   = Replace(erro, vbCrLf, " ")
    ws.Cells(lin, C_BIPQUANDO).Value = CStr(Now) & " | " & quemRodou
    If Trim(docFI) <> "" Then
        ws.Cells(lin, C_DOCCONT).NumberFormat = "@"
        ws.Cells(lin, C_DOCCONT).Value = docFI
    End If
    wb.Save
    Err.Clear
    On Error GoTo 0

    If status = "NAO BIPADA" Then
        totalNOK = totalNOK + 1
    Else
        totalOK = totalOK + 1
    End If

    VoltarAoMenu
End Sub

Sub Encerrar(msg)
    On Error Resume Next
    wb.Close False
    excel.Quit
    Err.Clear
    On Error GoTo 0
    MsgBox msg, vbCritical, TITULO
    WScript.Quit
End Sub

' =====================================================================
'  SAP - AUXILIARES
' =====================================================================
Function AcharGradeNaTela()
    Set gradeAchada = Nothing
    On Error Resume Next
    Varrer sessao.findById("wnd[0]")
    Err.Clear
    On Error GoTo 0
    AcharGradeNaTela = Not (gradeAchada Is Nothing)
End Function

Sub Varrer(obj)
    Dim filhos, i, f, tipo
    If Not (gradeAchada Is Nothing) Then
        Exit Sub
    End If
    On Error Resume Next
    tipo = obj.Type
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    If tipo = "GuiShell" Then
        If obj.SubType = "GridView" Then
            Set gradeAchada = obj
            Exit Sub
        End If
    End If
    Set filhos = Nothing
    Set filhos = obj.Children
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    For i = 0 To filhos.Count - 1
        Set f = Nothing
        Set f = filhos.ElementAt(i)
        If Err.Number <> 0 Then
            Err.Clear
        Else
            Varrer f
        End If
        If Not (gradeAchada Is Nothing) Then
            Exit Sub
        End If
    Next
    On Error GoTo 0
End Sub

Function Ler(g, lin, col)
    Ler = ""
    On Error Resume Next
    Ler = Trim(g.GetCellValue(lin, col))
    If Err.Number <> 0 Then
        Err.Clear
        Ler = ""
    End If
    On Error GoTo 0
End Function

Function LerCampo(id)
    Dim o
    LerCampo = ""
    On Error Resume Next
    Set o = Nothing
    Set o = sessao.findById(id)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    LerCampo = Trim(o.Text)
    Err.Clear
    On Error GoTo 0
End Function

Function Definir(id, valor, erro)
    Definir = False
    erro = ""
    On Error Resume Next
    sessao.findById(id).text = valor
    If Err.Number <> 0 Then
        erro = "campo " & id & " - " & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    Definir = True
End Function

Sub LimparCampo(id)
    On Error Resume Next
    sessao.findById(id).text = ""
    Err.Clear
    On Error GoTo 0
End Sub

Function ExisteJanela(n)
    Dim w
    ExisteJanela = False
    On Error Resume Next
    Set w = sessao.findById("wnd[" & n & "]")
    If Err.Number = 0 Then
        ExisteJanela = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

Function TextoDaJanela(n)
    TextoDaJanela = ""
    On Error Resume Next
    TextoDaJanela = Trim(sessao.findById("wnd[" & n & "]").Text)
    Err.Clear
    On Error GoTo 0
End Function

Sub FecharPopups()
    Dim guarda
    guarda = 0
    On Error Resume Next
    Do While ExisteJanela(1) And guarda < 6
        sessao.findById("wnd[1]").close
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
        guarda = guarda + 1
    Loop
    Err.Clear
    On Error GoTo 0
End Sub

Function TipoDaBarra()
    TipoDaBarra = ""
    On Error Resume Next
    TipoDaBarra = sessao.findById("wnd[0]/sbar").messageType
    Err.Clear
    On Error GoTo 0
End Function

Function TextoDaBarra()
    TextoDaBarra = ""
    On Error Resume Next
    TextoDaBarra = Trim(sessao.findById("wnd[0]/sbar").text)
    Err.Clear
    On Error GoTo 0
End Function

' Sai de onde estiver sem gravar nada. Se o SAP perguntar se quer
' perder os dados, a resposta e sim: quando chegamos aqui, ou ja
' gravamos, ou decidimos nao gravar.
Sub VoltarAoMenu()
    Dim guarda
    On Error Resume Next
    sessao.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    sessao.findById("wnd[0]").sendVKey 0
    Err.Clear
    guarda = 0
    Do While ExisteJanela(1) And guarda < 6
        sessao.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Err.Number <> 0 Then
            Err.Clear
            sessao.findById("wnd[1]").close
            If Err.Number <> 0 Then
                Err.Clear
                Exit Do
            End If
        End If
        guarda = guarda + 1
    Loop
    Err.Clear
    On Error GoTo 0
End Sub

' =====================================================================
'  BOLETO - validacao e conversao
' =====================================================================

' DV geral do codigo de barras: posicao 5, mod 11, pesos 2..9 ciclicos
' da direita para a esquerda sobre os outros 43 digitos.
'
' Boleto de ARRECADACAO / concessionaria (forma O - Operadora H, Operadora B, Operadora D,
' Operadora G...): comeca com 8. Layout diferente do bancario:
'   pos 1      = 8
'   pos 3      = identificador de valor: 6/7 -> DVs mod 10 ; 8/9 -> mod 11
'                (6 e 8 = valor real em reais; 7 e 9 = valor de referencia)
'   pos 4      = DV geral, calculado sobre os outros 43 digitos
'   pos 5-15   = valor (11 digitos, em centavos) quando pos 3 = 6 ou 8
' Linha digitavel: 48 digitos = 4 blocos de (11 digitos + 1 DV).
Function EhArrecadacao(b)
    EhArrecadacao = (Left(b, 1) = "8")
End Function

Function DVArrec(b43, ident)
    If ident = "6" Or ident = "7" Then
        DVArrec = DVMod10(b43)
    Else
        DVArrec = DVMod11Arrec(b43)
    End If
End Function

Function DVMod11Arrec(s)
    Dim i, peso, soma, resto
    soma = 0
    peso = 2
    For i = Len(s) To 1 Step -1
        soma = soma + CInt(Mid(s, i, 1)) * peso
        peso = peso + 1
        If peso > 9 Then
            peso = 2
        End If
    Next
    resto = soma Mod 11
    If resto < 2 Then
        DVMod11Arrec = 0
    Else
        DVMod11Arrec = 11 - resto
    End If
End Function

' Valor em centavos embutido no codigo de barras (0 = sem valor / referencia)
Function ValorDaBarra(b44)
    Dim ident
    If EhArrecadacao(b44) Then
        ident = Mid(b44, 3, 1)
        If ident = "6" Or ident = "8" Then
            ValorDaBarra = CDbl(Mid(b44, 5, 11))
        Else
            ValorDaBarra = 0
        End If
    Else
        ValorDaBarra = CDbl(Mid(b44, 10, 10))
    End If
End Function

Function DVBarraOK(b44)
    Dim b43, calc
    DVBarraOK = False
    If Len(b44) <> 44 Then
        Exit Function
    End If
    If EhArrecadacao(b44) Then
        If CStr(DVArrec(Left(b44, 3) & Mid(b44, 5), Mid(b44, 3, 1))) = Mid(b44, 4, 1) Then
            DVBarraOK = True
        End If
        Exit Function
    End If
    b43 = Left(b44, 4) & Mid(b44, 6)
    calc = DVMod11(b43)
    If CStr(calc) = Mid(b44, 5, 1) Then
        DVBarraOK = True
    End If
End Function

Function DVMod11(b43)
    Dim i, peso, soma, resto, d
    soma = 0
    peso = 2
    For i = Len(b43) To 1 Step -1
        d = CInt(Mid(b43, i, 1))
        soma = soma + d * peso
        peso = peso + 1
        If peso > 9 Then
            peso = 2
        End If
    Next
    resto = 11 - (soma Mod 11)
    If resto = 0 Or resto = 10 Or resto = 11 Then
        resto = 1
    End If
    DVMod11 = resto
End Function

Function DVMod10(bloco)
    Dim i, peso, soma, p
    soma = 0
    peso = 2
    For i = Len(bloco) To 1 Step -1
        p = CInt(Mid(bloco, i, 1)) * peso
        If p > 9 Then
            p = p - 9
        End If
        soma = soma + p
        If peso = 2 Then
            peso = 1
        Else
            peso = 2
        End If
    Next
    DVMod10 = (10 - (soma Mod 10)) Mod 10
End Function

' Linha digitavel de 47 digitos -> codigo de barras de 44.
' Confere os 3 DVs de campo (mod 10) e o DV geral (mod 11) antes de
' devolver: linha digitada errada e exatamente o que nao pode passar.
Function LinhaParaBarra(d47, erro)
    Dim c1, c2, c3, dv, fator, valor, livre, barra
    LinhaParaBarra = ""
    erro = ""
    If Len(d47) = 48 And EhArrecadacao(d47) Then
        LinhaParaBarra = LinhaArrecParaBarra(d47, erro)
        Exit Function
    End If
    If Len(d47) <> 47 Then
        erro = "a linha digitavel tem " & Len(d47) & " digitos e deveria ter 47"
        Exit Function
    End If

    c1 = Mid(d47, 1, 9)
    c2 = Mid(d47, 11, 10)
    c3 = Mid(d47, 22, 10)

    If CStr(DVMod10(c1)) <> Mid(d47, 10, 1) Then
        erro = "o digito do 1o campo nao confere"
        Exit Function
    End If
    If CStr(DVMod10(c2)) <> Mid(d47, 21, 1) Then
        erro = "o digito do 2o campo nao confere"
        Exit Function
    End If
    If CStr(DVMod10(c3)) <> Mid(d47, 32, 1) Then
        erro = "o digito do 3o campo nao confere"
        Exit Function
    End If

    dv    = Mid(d47, 33, 1)
    fator = Mid(d47, 34, 4)
    valor = Mid(d47, 38, 10)
    livre = Mid(d47, 5, 5) & Mid(d47, 11, 10) & Mid(d47, 22, 10)

    barra = Mid(d47, 1, 4) & dv & fator & valor & livre
    If Len(barra) <> 44 Then
        erro = "a montagem do codigo de barras saiu com " & Len(barra) & " digitos"
        Exit Function
    End If
    If Not DVBarraOK(barra) Then
        erro = "o digito verificador geral nao confere"
        Exit Function
    End If
    LinhaParaBarra = barra
End Function

' Linha digitavel de arrecadacao (48 digitos, comeca com 8) -> barra 44.
' Confere o DV de cada bloco de 11 e o DV geral antes de devolver.
Function LinhaArrecParaBarra(d48, erro)
    Dim i, bloco, ident, barra
    LinhaArrecParaBarra = ""
    erro = ""
    ident = Mid(d48, 3, 1)
    barra = ""
    For i = 0 To 3
        bloco = Mid(d48, i * 12 + 1, 11)
        If CStr(DVArrec(bloco, ident)) <> Mid(d48, i * 12 + 12, 1) Then
            erro = "o digito do " & (i + 1) & "o bloco nao confere"
            Exit Function
        End If
        barra = barra & bloco
    Next
    If Not DVBarraOK(barra) Then
        erro = "o digito verificador geral nao confere"
        Exit Function
    End If
    LinhaArrecParaBarra = barra
End Function

' =====================================================================
'  TEXTO E NUMERO
' =====================================================================
Function SoDigitos(t)
    Dim i, c, s
    s = ""
    For i = 1 To Len(t)
        c = Mid(t, i, 1)
        If c >= "0" And c <= "9" Then
            s = s & c
        End If
    Next
    SoDigitos = s
End Function

' DATA DA FATURA vem como ddmmaaaa (18082026) ou dd.mm.aaaa. Nos dois
' casos o ano sao os 4 ultimos digitos.
Function AnoDaData(t)
    Dim d
    AnoDaData = ""
    d = SoDigitos(t)
    If Len(d) < 8 Then
        Exit Function
    End If
    AnoDaData = Right(d, 4)
End Function

Function Formatar(cent)
    Dim s, neg
    neg = ""
    If cent < 0 Then
        neg = "-"
        cent = -cent
    End If
    s = CStr(cent)
    Do While Len(s) < 3
        s = "0" & s
    Loop
    Formatar = neg & Left(s, Len(s) - 2) & "," & Right(s, 2)
End Function

Function EmCentavos(t)
    Dim s, i, c, limpo, pos, parteInt, parteDec, digitos, n
    EmCentavos = CENTAVOS_INVALIDO
    s = Trim(t)
    If s = "" Then
        Exit Function
    End If

    limpo = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If (c >= "0" And c <= "9") Or c = "," Or c = "." Then
            limpo = limpo & c
        End If
    Next
    If limpo = "" Then
        Exit Function
    End If

    pos = 0
    For i = Len(limpo) To 1 Step -1
        c = Mid(limpo, i, 1)
        If c = "," Or c = "." Then
            pos = i
            Exit For
        End If
    Next

    parteInt = limpo
    parteDec = ""
    If pos > 0 Then
        If Len(limpo) - pos >= 1 And Len(limpo) - pos <= 2 Then
            parteInt = Left(limpo, pos - 1)
            parteDec = Mid(limpo, pos + 1)
        End If
    End If

    digitos = ""
    For i = 1 To Len(parteInt)
        c = Mid(parteInt, i, 1)
        If c >= "0" And c <= "9" Then
            digitos = digitos & c
        End If
    Next
    parteInt = digitos
    If parteInt = "" Then
        parteInt = "0"
    End If
    If Len(parteInt) > 12 Then
        Exit Function
    End If

    If Len(parteDec) = 0 Then
        parteDec = "00"
    End If
    If Len(parteDec) = 1 Then
        parteDec = parteDec & "0"
    End If

    n = CDbl(parteInt) * 100 + CDbl(parteDec)
    If InStr(Trim(t), "-") > 0 Then
        n = -n
    End If
    EmCentavos = n
End Function

' =====================================================================
'  LISTAS DO MAPA
' =====================================================================
Function FormasAceitas()
    FormasAceitas = Trim(mapa("FORMAS_VALIDAS"))
    If FormasAceitas = "" Then
        FormasAceitas = "O|S|X"
    End If
End Function

Function FormaValida(f)
    Dim lista, k
    FormaValida = False
    lista = Split(UCase(FormasAceitas()), "|")
    For k = 0 To UBound(lista)
        If Trim(lista(k)) = f Then
            FormaValida = True
        End If
    Next
End Function

Function FormaSemBoleto(f)
    Dim lista, k, texto
    FormaSemBoleto = False
    texto = Trim(mapa("FORMAS_SEM_BOLETO"))
    If texto = "" Then
        Exit Function
    End If
    lista = Split(UCase(texto), "|")
    For k = 0 To UBound(lista)
        If Trim(lista(k)) = f Then
            FormaSemBoleto = True
        End If
    Next
End Function

Function EhChaveFornecedor(chave)
    Dim lista, k, texto, c
    EhChaveFornecedor = False
    c = Trim(chave)
    If c = "" Then
        Exit Function
    End If
    texto = Trim(mapa("CHAVES_FORNECEDOR"))
    If texto = "" Then
        texto = "31|21"
    End If
    lista = Split(texto, "|")
    For k = 0 To UBound(lista)
        If Trim(lista(k)) = c Then
            EhChaveFornecedor = True
        End If
    Next
End Function

Function CarregarMapa(caminho, dic)
    Dim f, l, p
    CarregarMapa = False
    If Not fso.FileExists(caminho) Then
        Exit Function
    End If
    Set f = fso.OpenTextFile(caminho, ForReading)
    Do Until f.AtEndOfStream
        l = Trim(f.ReadLine)
        If l <> "" And Left(l, 1) <> "#" Then
            p = InStr(l, "=")
            If p > 1 Then
                dic(Trim(Left(l, p - 1))) = Trim(Mid(l, p + 1))
            End If
        End If
    Loop
    f.Close
    CarregarMapa = (dic.Count > 0)
End Function
