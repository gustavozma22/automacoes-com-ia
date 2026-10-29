' =====================================================================
'  LANCAR_NFCOM.vbs  --  Executor local de lancamento de NFCOMs
'  Projeto de automacao com IA - Contas a Pagar
'  Versao 3 - le e escreve direto na planilha LANCAMENTO_NFCOM.xlsx
' =====================================================================
'
'  O QUE ESTE ARQUIVO FAZ
'  ----------------------
'  Le a aba LOTE da planilha LANCAMENTO_NFCOM.xlsx (preenchida pelo
'  Claude a partir da VALIDACAO_IA) e, para cada linha, reproduz no
'  SAP a mesma sequencia que o time faz na mao:
'
'     MIRO -> data / referencia / montante
'          -> selecao multipla de pedidos
'          -> valor liquido de cada item, casado POR NUMERO DE PEDIDO
'          -> aba Detalhe: tipo de NF = TA
'          -> tela Nota Fiscal: protocolo, data, hora, tp.emissao,
'             numero aleatorio
'          -> Simular -> Lancar
'
'  e escreve o resultado (STATUS, Nº DOC MIRO, ERRO, QUANDO, EXECUTADO
'  POR) nas colunas cinzas da PROPRIA LINHA, salvando a cada nota.
'
'  PRINCIPIOS DE SEGURANCA (nao alterar sem aprovacao)
'  ---------------------------------------------------
'  * NAO faz login. Anexa-se a uma sessao SAP que o usuario ja abriu e
'    autenticou. Nenhuma senha e armazenada, digitada ou transmitida.
'  * IGNORA qualquer linha que ja tenha STATUS preenchido. E isso que
'    impede lancamento em duplicidade se a planilha rodar duas vezes.
'  * NAO grava se houver mensagem de erro ou aviso, ou se a janela de
'    simulacao nao abrir. Na duvida, marca NAO LANCADA.
'  * NUNCA troca PARK por POST por conta propria.
'  * Uma nota com problema nunca derruba o lote: registra o motivo,
'    volta para a tela inicial e segue para a proxima.
'
'  COMO USAR
'  ---------
'     1. Abrir o SAP GUI e logar normalmente.
'     2. FECHAR a planilha LANCAMENTO_NFCOM.xlsx no Excel.
'     3. Deixar planilha, MAPA_CAMPOS.txt e este arquivo na mesma pasta.
'     4. Duplo clique neste arquivo.
'     5. Reabrir a planilha e conferir as colunas cinzas.
' =====================================================================

Option Explicit

' --- colunas da aba LOTE ---------------------------------------------
Const C_ID        = 1
Const C_EMPRESA   = 2
Const C_DTFATURA  = 3
Const C_MONTANTE  = 4
Const C_REFER     = 5
Const C_TIPONF    = 6
Const C_PEDIDOS   = 7
Const C_VALORES   = 8
Const C_BRUTOS    = 9          ' valor BRUTO por pedido (liquido + ICMS)
Const C_NUMLOG    = 10
Const C_DTPROC    = 11
Const C_TMPROC    = 12
Const C_TPEMIS    = 13
Const C_ALEAT     = 14
Const C_MODELO    = 15
Const C_CHAVE     = 16
Const C_MODO      = 17
Const C_STATUS    = 18
Const C_DOCMIRO   = 19
Const C_ERRO      = 20
Const C_QUANDO    = 21
Const C_QUEM      = 22
Const C_FORMA       = 40   ' FORMA PGTO (ZLSCH)
Const C_LINHADIG    = 41   ' LINHA DIGITAVEL (47)
Const C_BARRAS      = 42   ' CODIGO DE BARRAS (44)
Const C_DOCCONT     = 43   ' DOC CONTABIL (FI)
Const C_BIPSTATUS   = 44   ' STATUS BIPAGEM
Const C_BIPERRO     = 45   ' ERRO BIPAGEM
Const C_BIPQUANDO   = 46   ' QUANDO BIPADO
Dim bipGradeAchada

' saldo interpretado como numero: sentinela de "nao consegui ler"
Const CENTAVOS_INVALIDO = -999999999

Const ForReading = 1
Dim geoGrade, obsQtd, obsImposto, obsAviso
' Teto de SEGURANCA, nao limite de trabalho. Desde que CelGrade passou a
' rolar as tabelas, nao existe mais numero maximo de pedidos por nota: o
' executor alcanca qualquer linha, com quantos pedidos a nota tiver. Este
' numero existe so para impedir laco desgovernado caso o SAP devolva
' celula indefinidamente.
Const MAX_ITENS  = 500
Const MAX_LINHAS = 400         ' tamanho da area de dados da planilha

Dim fso, pasta, mapa, sessao, aplicacao, conexao, SapGuiAuto
Dim excel, wb, ws, arquivo
Dim r, totalOK, totalNOK, quemRodou
Dim bipMapaOk

Set fso = CreateObject("Scripting.FileSystemObject")
pasta   = fso.GetParentFolderName(WScript.ScriptFullName)
arquivo = pasta & "\LANCAMENTO_NFCOM.xlsx"

Dim quemWindows
quemWindows = ""
On Error Resume Next
quemWindows = CreateObject("WScript.Network").UserName
Err.Clear
On Error GoTo 0
quemRodou = quemWindows      ' completado com o usuario SAP mais adiante

' ---------------------------------------------------------------------
' 1. Mapa de campos
' ---------------------------------------------------------------------
Set mapa = CreateObject("Scripting.Dictionary")
If Not CarregarMapa(pasta & "\MAPA_CAMPOS.txt", mapa) Then
    MsgBox "MAPA_CAMPOS.txt nao encontrado ou vazio.", vbCritical, "NFCOM"
    WScript.Quit
End If
' Mapa de campos da BIPAGEM (FB03/FB02), no MESMO dicionario - assim a
' bipagem, logo apos lancar, usa o mesmo mapa sem precisar de outro
' arquivo carregado a parte. Se nao existir, so desativa a bipagem
' automatica (cada nota continua lancando normalmente); nao aborta o
' lote.
bipMapaOk = CarregarMapa(pasta & "\MAPA_CAMPOS_BOLETO.txt", mapa)

' ---------------------------------------------------------------------
' 2. Planilha
' ---------------------------------------------------------------------
If Not fso.FileExists(arquivo) Then
    MsgBox "LANCAMENTO_NFCOM.xlsx nao encontrado nesta pasta.", vbCritical, "NFCOM"
    WScript.Quit
End If

On Error Resume Next
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Nao consegui abrir o Excel nesta maquina.", vbCritical, "NFCOM"
    WScript.Quit
End If
excel.Visible = False
excel.DisplayAlerts = False
Set wb = excel.Workbooks.Open(arquivo)
If Err.Number <> 0 Then
    MsgBox "Nao consegui abrir a planilha: " & Err.Description, vbCritical, "NFCOM"
    excel.Quit : WScript.Quit
End If
Set ws = wb.Worksheets("LOTE")
If Err.Number <> 0 Then
    MsgBox "A planilha nao tem uma aba chamada LOTE.", vbCritical, "NFCOM"
    wb.Close False : excel.Quit : WScript.Quit
End If
Err.Clear
On Error GoTo 0

If wb.ReadOnly Then
    MsgBox "A planilha abriu somente para leitura - provavelmente esta aberta no Excel." & vbCrLf & vbCrLf & _
           "Feche LANCAMENTO_NFCOM.xlsx e rode de novo.", vbCritical, "NFCOM"
    wb.Close False : excel.Quit : WScript.Quit
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
    Encerrar "O SAP GUI Scripting esta desabilitado nesta maquina ou neste servidor." & vbCrLf & vbCrLf & _
             "Peca a TI/Basis:" & vbCrLf & _
             "  - servidor: sapgui/user_scripting = TRUE (RZ11)" & vbCrLf & _
             "  - estacao:  SAP Logon > Opcoes > Acessibilidade e Script"
End If
Set conexao = aplicacao.Children(0)
Set sessao  = conexao.Children(0)
If Err.Number <> 0 Then Encerrar "Nenhuma sessao SAP ativa. Faca login no SAP primeiro."
Err.Clear

' --- em qual sistema estamos, e e mesmo esse? -------------------------
' O script nao faz login: ele entra na sessao que o usuario ja abriu.
' Logo, quem lanca no SAP e o usuario logado - e todo documento gravado
' fica com o nome dele. Por isso vale conferir, ANTES de lancar, que a
' sessao capturada e a que se pretendia usar.
Dim sapSistema, sapMandante, sapUsuario
sapSistema  = "" : sapMandante = "" : sapUsuario = ""
sapSistema  = sessao.Info.SystemName
sapMandante = sessao.Info.Client
sapUsuario  = sessao.Info.User
Err.Clear
On Error GoTo 0

' --- ambiente: pode haver mais de um liberado -------------------------
' Com um unico ambiente fixo, a conferencia podia ser silenciosa: so
' existia um lugar possivel. Liberando PRD e QAS ao mesmo tempo isso
' muda - o lote NAO sabe para qual ambiente foi montado, e a unica
' coisa que separa um teste no QAS de um documento real no PRD passa a
' ser a sessao que estiver aberta. Por isso, quando ha mais de um
' ambiente na lista, a confirmacao deixa de ser opcional.
Dim ambienteAtual, listaAmb, ambientePermitido, nAmb, itemAmb
ambienteAtual = UCase(Trim(sapSistema)) & "/" & Trim(sapMandante)
ambientePermitido = False

If mapa("SAP_AMBIENTES_PERMITIDOS") <> "" Then
    listaAmb = Split(mapa("SAP_AMBIENTES_PERMITIDOS"), "|")
    For nAmb = 0 To UBound(listaAmb)
        itemAmb = UCase(Replace(Trim(listaAmb(nAmb)), " ", ""))
        If itemAmb = ambienteAtual Then ambientePermitido = True
    Next

    If Not ambientePermitido Then
        Encerrar "A sessao SAP aberta nao e um dos ambientes liberados para este processo." & vbCrLf & vbCrLf & _
                 "Liberados: " & mapa("SAP_AMBIENTES_PERMITIDOS") & vbCrLf & _
                 "Encontrado: " & ambienteAtual & vbCrLf & vbCrLf & _
                 "Nada foi lancado. Deixe aberta apenas a janela do ambiente certo."
    End If

    Dim pulaConfirmacao
    pulaConfirmacao = (Len(Trim(mapa("SAP_USUARIO_ESPERADO"))) > 0 And _
                       UCase(Trim(sapUsuario)) = UCase(Trim(mapa("SAP_USUARIO_ESPERADO"))) And _
                       ambienteAtual = "PRD/100")
    If Not pulaConfirmacao Then
        If MsgBox("Prestes a lancar em:" & vbCrLf & vbCrLf & _
                  "        >>>  " & ambienteAtual & "  <<<" & vbCrLf & vbCrLf & _
                  "  Usuario: " & sapUsuario & vbCrLf & vbCrLf & _
                  "Este processo esta liberado para mais de um ambiente, e o lote" & vbCrLf & _
                  "nao sabe para qual dos dois foi montado. Confira antes de seguir:" & vbCrLf & _
                  "se aqui estiver PRD, os documentos sao reais." & vbCrLf & vbCrLf & _
                  "Continuar?", vbYesNo + vbExclamation, "NFCOM - confirmar ambiente") <> vbYes Then
            Encerrar "Lote cancelado pelo usuario. Nada foi lancado."
        End If
    End If

ElseIf mapa("SAP_SISTEMA_ESPERADO") <> "" And mapa("SAP_MANDANTE_ESPERADO") <> "" Then
    If UCase(Trim(sapSistema)) <> UCase(Trim(mapa("SAP_SISTEMA_ESPERADO"))) Or _
       Trim(sapMandante) <> Trim(mapa("SAP_MANDANTE_ESPERADO")) Then
        Encerrar "A sessao SAP aberta nao e a esperada para este lote." & vbCrLf & vbCrLf & _
                 "Esperado: " & mapa("SAP_SISTEMA_ESPERADO") & " / mandante " & mapa("SAP_MANDANTE_ESPERADO") & vbCrLf & _
                 "Encontrado: " & sapSistema & " / mandante " & sapMandante & vbCrLf & vbCrLf & _
                 "Nada foi lancado. Feche as outras janelas do SAP e deixe aberta apenas a do ambiente correto."
    End If
Else
    If MsgBox("Prestes a lancar em:" & vbCrLf & vbCrLf & _
              "  Sistema:  " & sapSistema & vbCrLf & _
              "  Mandante: " & sapMandante & vbCrLf & _
              "  Usuario:  " & sapUsuario & vbCrLf & vbCrLf & _
              "Todos os documentos serao gravados no SAP em nome desse usuario." & vbCrLf & vbCrLf & _
              "Continuar?", vbYesNo + vbQuestion, "NFCOM - confirmar ambiente") <> vbYes Then
        Encerrar "Lote cancelado pelo usuario. Nada foi lancado."
    End If
End If

' --- aviso de usuario diferente do habitual --------------------------
' Nao bloqueia: quem lanca e sempre quem esta logado, e um substituto
' legitimo deve conseguir seguir. O objetivo e so evitar que alguem
' rode sem perceber que a sessao esta em outro usuario.
If mapa("SAP_USUARIO_ESPERADO") <> "" Then
    If UCase(Trim(sapUsuario)) <> UCase(Trim(mapa("SAP_USUARIO_ESPERADO"))) Then
        If MsgBox("A sessao SAP esta no usuario " & sapUsuario & ", e o usuario habitual deste processo e " & _
                  mapa("SAP_USUARIO_ESPERADO") & "." & vbCrLf & vbCrLf & _
                  "Os documentos serao gravados em nome de " & sapUsuario & "." & vbCrLf & vbCrLf & _
                  "Continuar mesmo assim?", vbYesNo + vbExclamation, "NFCOM - usuario diferente do habitual") <> vbYes Then
            Encerrar "Lote cancelado pelo usuario. Nada foi lancado."
        End If
    End If
End If

' quem consta como autor no SAP, em qual ambiente, e quem apertou o botao
quemRodou = sapUsuario & "@" & sapSistema & "/" & sapMandante & " (win: " & quemWindows & ")"

On Error Resume Next
sessao.findById("wnd[0]").maximize
Err.Clear
On Error GoTo 0

' ---------------------------------------------------------------------
' 4. Percorre as linhas
' ---------------------------------------------------------------------
totalOK = 0 : totalNOK = 0
Dim pulados
pulados = 0

For r = 2 To MAX_LINHAS + 1
    If Celula(r, C_ID) = "" Then Exit For
    If Celula(r, C_STATUS) <> "" Then
        pulados = pulados + 1        ' ja processada em rodada anterior
    ElseIf Not MinhaLinha(Celula(r, C_TIPONF)) Then
        ' A planilha e compartilhada com o executor de NFS-e. Linha de
        ' outro tipo nao e erro: e so nao ser deste script. Passa reto
        ' SEM escrever nada, para o outro executor encontrar a linha
        ' com o STATUS ainda vazio.
        pulados = pulados + 1
    Else
        ProcessarLinha r
    End If
Next

wb.Save
wb.Close True
excel.Quit

MsgBox "Lote finalizado." & vbCrLf & vbCrLf & _
       "Lancadas: " & totalOK & vbCrLf & _
       "Retidas para conferencia: " & totalNOK & vbCrLf & _
       "Ja processadas antes (ignoradas): " & pulados & vbCrLf & vbCrLf & _
       "Confira as colunas cinzas da planilha.", vbInformation, "NFCOM"

' =====================================================================
'  UMA LINHA DA PLANILHA
' =====================================================================
Sub ProcessarLinha(lin)
    Dim id, bukrs, dataFatura, montante, referencia, tipoNF
    Dim pedidos, valores, brutos, numLog, dataProc, tempoProc
    Dim tpEmissao, numAleat, modeloNF, modo
    Dim erroMsg, docMiro, n, k, botao, msgFinal
    Dim pedTela, gradePed(), linhasGrade, contagem, idxLote, m, tentativa
    Dim tblSel, tblIt, colSel, colPed, colVal, colQtd, cel, qtdTxt
    Dim idPopup, idSaldo, saldoTxt, lidoPopup, conf, difCent, tolCent, falhou, lidoValor, qtdEdita, colQtdNome

    id         = Celula(lin, C_ID)
    bukrs      = Celula(lin, C_EMPRESA)
    dataFatura = Celula(lin, C_DTFATURA)
    montante   = Celula(lin, C_MONTANTE)
    referencia = Celula(lin, C_REFER)
    tipoNF     = Celula(lin, C_TIPONF)
    pedidos    = Split(Celula(lin, C_PEDIDOS), "|")
    valores    = Split(Celula(lin, C_VALORES), "|")
    brutos     = Split(Celula(lin, C_BRUTOS), "|")
    numLog     = Celula(lin, C_NUMLOG)
    dataProc   = Celula(lin, C_DTPROC)
    tempoProc  = Celula(lin, C_TMPROC)
    tpEmissao  = Celula(lin, C_TPEMIS)
    numAleat   = Celula(lin, C_ALEAT)
    modeloNF   = Celula(lin, C_MODELO)
    modo       = UCase(Celula(lin, C_MODO))
    If modo = "" Then modo = "POST"

    ' --- travas antes de tocar no SAP ---------------------------------
    If modeloNF <> "62" Then
        Resultado lin, "NAO LANCADA", "", "Modelo da NF = " & modeloNF & " (esperado 62). Parametrizacao a verificar - nao corrigido no lancamento."
        Exit Sub
    End If

    If UBound(pedidos) <> UBound(valores) Then
        Resultado lin, "NAO LANCADA", "", "Sao " & (UBound(pedidos) + 1) & " pedidos e " & (UBound(valores) + 1) & " valores liquidos - o rateio nao esta pareado."
        Exit Sub
    End If

    ' Os brutos so sao usados para corrigir a base do imposto. Se a coluna
    ' estiver vazia a nota lanca normalmente - apenas nao tera como
    ' corrigir divergencia de centavos, e ai fica retida como antes.
    If Celula(lin, C_BRUTOS) <> "" Then
        If UBound(pedidos) <> UBound(brutos) Then
            Resultado lin, "NAO LANCADA", "", "Sao " & (UBound(pedidos) + 1) & " pedidos e " & (UBound(brutos) + 1) & " valores brutos - nao estao pareados."
            Exit Sub
        End If
    End If

    If UBound(pedidos) + 1 > MAX_ITENS Then
        Resultado lin, "NAO LANCADA", "", "Nota com " & (UBound(pedidos) + 1) & " pedidos - acima do teto de seguranca de " & MAX_ITENS & ". Confira o rateio da planilha."
        Exit Sub
    End If

    If modo = "PARK" And mapa("MIRO_BOTAO_PRE_LANCAR") = "" Then
        Resultado lin, "NAO LANCADA", "", "Linha pede PRE-LANCAMENTO, mas MIRO_BOTAO_PRE_LANCAR nao esta mapeado. Nao lancada por seguranca."
        Exit Sub
    End If

    On Error Resume Next

    ' ------------------------------------------------------------------
    ' MIRO
    ' ------------------------------------------------------------------
    ' Limpa erro herdado: sem isso, uma falha da nota ANTERIOR sobrevive
    ' ate aqui e a nota seguinte morre com "Falha ao abrir a MIRO:" e
    ' descricao vazia, sem ter havido falha nenhuma.
    Err.Clear
    sessao.findById("wnd[0]/tbar[0]/okcd").text = "/nMIRO"
    sessao.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        Resultado lin, "NAO LANCADA", "", "Falha ao abrir a MIRO: " & Err.Description
        Err.Clear : Exit Sub
    End If

    ' --- janela de empresa ---------------------------------------------
    ' A MIRO pede a empresa na primeira nota de cada sessao. O campo e
    ' preenchido sozinho com o valor da coluna EMPRESA da planilha: se o
    ' ID nao estiver no mapa, o script procura na propria janela um campo
    ' de empresa (BUKRS). So retem se nao achar nenhum.
    If ExisteJanela(1) Then
        idPopup = mapa("MIRO_POPUP_EMPRESA")
        If idPopup = "" Then idPopup = BuscarId(sessao.findById("wnd[1]/usr"), "BUKRS", 0)
        Err.Clear
        If idPopup = "" Then
            FecharPopups
            Resultado lin, "NAO LANCADA", "", "Apareceu uma janela na entrada da MIRO e nao achei nela um campo de empresa. Titulo da janela anotado para mapeamento."
            Err.Clear : Exit Sub
        End If
        If bukrs = "" Then
            FecharPopups
            Resultado lin, "NAO LANCADA", "", "A MIRO pediu a empresa e a coluna EMPRESA desta linha esta vazia."
            Err.Clear : Exit Sub
        End If
        sessao.findById(idPopup).text = bukrs
        sessao.findById("wnd[1]").sendVKey 0
        Err.Clear
    End If

    ' --- cabecalho ----------------------------------------------------
    Definir mapa("MIRO_DATA_FATURA"), dataFatura
    Definir mapa("MIRO_REFERENCIA"),  referencia
    Definir mapa("MIRO_MONTANTE"),    montante
    If Err.Number <> 0 Then
        Resultado lin, "NAO LANCADA", "", "Falha no cabecalho da MIRO: " & Err.Description
        Err.Clear : Exit Sub
    End If

    ' --- pedidos, via selecao multipla --------------------------------
    If mapa("MIRO_BOTAO_MULTIPLOS_PEDIDOS") = "" Then
        Resultado lin, "NAO LANCADA", "", "Botao de selecao multipla de pedidos nao mapeado"
        Err.Clear : Exit Sub
    End If

    sessao.findById(mapa("MIRO_BOTAO_MULTIPLOS_PEDIDOS")).press
    If Err.Number <> 0 Or Not ExisteJanela(1) Then
        Resultado lin, "NAO LANCADA", "", "A janela de selecao multipla de pedidos nao abriu"
        Err.Clear : Exit Sub
    End If

    ' A janela de selecao mostra poucas linhas por vez. Com 10 pedidos e
    ' preciso rolar a tabela - por isso o acesso e feito pelo objeto da
    ' tabela, e nao por um ID de linha fixo.
    Set tblSel = sessao.findById(IdTabela(mapa("MIRO_PEDIDO_LINHA")))
    If Err.Number <> 0 Then
        FecharPopups
        Resultado lin, "NAO LANCADA", "", "Nao encontrei a tabela de pedidos da janela de selecao multipla"
        Err.Clear : Exit Sub
    End If
    colSel = IdxColuna(mapa("MIRO_PEDIDO_LINHA"))

    For n = 0 To UBound(pedidos)
        Set cel = CelGrade(IdTabela(mapa("MIRO_PEDIDO_LINHA")), colSel, n)
        If cel Is Nothing Then
            FecharPopups
            Resultado lin, "NAO LANCADA", "", "Nao consegui chegar na linha " & n & " da janela de selecao de pedidos (a nota tem " & (UBound(pedidos) + 1) & " pedidos) [" & geoGrade & "]"
            Err.Clear : Exit Sub
        End If
        cel.text = Trim(pedidos(n))
        If Err.Number <> 0 Then
            FecharPopups
            Resultado lin, "NAO LANCADA", "", "Falha ao informar o pedido " & Trim(pedidos(n)) & " na linha " & n
            Err.Clear : Exit Sub
        End If
    Next

    ' CONFERE o que ficou de fato na janela antes de transferir. Escrever
    ' numa tabela do SAP nao garante que o valor ficou: rolagem, validacao
    ' e round trip podem descartar. Sem essa conferencia, uma janela que
    ' perdeu pedidos vira "grade vazia" la na frente, sem explicacao.
    lidoPopup = ""
    For n = 0 To UBound(pedidos)
        Set cel = CelGrade(IdTabela(mapa("MIRO_PEDIDO_LINHA")), colSel, n)
        If cel Is Nothing Then
            FecharPopups
            Resultado lin, "NAO LANCADA", "", "Nao consegui reler a linha " & n & " da janela de pedidos [" & geoGrade & "]"
            Err.Clear : Exit Sub
        End If
        conf = ""
        conf = Trim(cel.Text)
        Err.Clear
        lidoPopup = lidoPopup & conf & "|"
        If conf <> Trim(pedidos(n)) Then
            FecharPopups
            Resultado lin, "NAO LANCADA", "", "A janela de pedidos nao guardou o que foi informado: na linha " & n & " esperava " & Trim(pedidos(n)) & " e esta " & conf & ". Lidos ate aqui: " & lidoPopup
            Err.Clear : Exit Sub
        End If
    Next

    ' volta a tabela para o topo antes de transferir
    Set cel = CelGrade(IdTabela(mapa("MIRO_PEDIDO_LINHA")), colSel, 0)
    Err.Clear

    ' "Verificar" (btn[0]) so e pressionado se estiver mapeado. O padrao e
    ' ir direto no "Transferir": em PRD a verificacao intermediaria deixou
    ' a janela num estado em que o Transferir nao adotava nada.
    If mapa("MIRO_BOTAO_CONTINUAR_PEDIDOS") <> "" Then
        sessao.findById(mapa("MIRO_BOTAO_CONTINUAR_PEDIDOS")).press
        Err.Clear
        erroMsg = MensagemDaBarra()
        If erroMsg <> "" Then
            FecharPopups
            Resultado lin, "NAO LANCADA", "", "Selecao de pedidos: " & erroMsg
            Err.Clear : Exit Sub
        End If
    End If

    sessao.findById(mapa("MIRO_BOTAO_ADOTAR_PEDIDOS")).press
    If Err.Number <> 0 Then
        FecharPopups
        Resultado lin, "NAO LANCADA", "", "Falha ao adotar os pedidos: " & Err.Description
        Err.Clear : Exit Sub
    End If

    erroMsg = MensagemDaBarra()
    If erroMsg <> "" Then
        Resultado lin, "NAO LANCADA", "", "Apos adotar os pedidos: " & erroMsg
        Err.Clear : Exit Sub
    End If

    ' Se a janela continua aberta, o Transferir nao aconteceu. Isso e uma
    ' coisa; grade vazia e outra. Separar as duas evita diagnostico errado.
    If ExisteJanela(1) Then
        ' O Transferir nao adotou e o SAP nao disse nada. O "Verificar"
        ' costuma fazer o SAP falar: aperta so para OUVIR a mensagem, nao
        ' para tentar de novo. Sem isso a falha volta muda e nao da para
        ' saber se o pedido nao tem entrada pendente ou se e outra coisa.
        erroMsg = ""
        If ExisteJanela(1) Then
            sessao.findById("wnd[1]/tbar[0]/btn[0]").press
            Err.Clear
            erroMsg = MensagemDaBarraCompleta()
        End If
        DespejarTela "POPUP_NAO_FECHOU.txt", "a janela de referencia a pedidos continuou aberta depois do Transferir (nota da linha " & lin & ")"
        FecharPopups
        If Trim(erroMsg) = "" Then erroMsg = "o SAP nao devolveu mensagem"
        Resultado lin, "NAO LANCADA", "", "A janela de referencia a pedidos nao fechou depois do Transferir - nenhum item foi adotado. O SAP disse: " & erroMsg & ". Pedidos conferidos na janela: " & lidoPopup & " - gerei o POPUP_NAO_FECHOU.txt"
        Err.Clear : Exit Sub
    End If

    ' ------------------------------------------------------------------
    ' Valor liquido de cada item
    ' ------------------------------------------------------------------
    ' A ordem dos pedidos na grade NAO importa - o que importa e que cada
    ' valor caia na linha do pedido a que pertence. Por isso o casamento
    ' e feito pelo NUMERO DO PEDIDO lido da propria grade.
    '
    ' A grade mostra poucas linhas por vez (3, no ambiente de testes), e
    ' uma nota pode ter ate 10 pedidos. Todo acesso passa por CelGrade,
    ' que rola a tabela quando a linha pedida esta fora da area visivel.
    If mapa("MIRO_PEDIDO_ITEM_LINHA") = "" Then
        Resultado lin, "NAO LANCADA", "", "MIRO_PEDIDO_ITEM_LINHA nao esta mapeado. Sem a coluna de pedido da grade nao da para saber de quem e cada linha - lancar manualmente."
        Err.Clear : Exit Sub
    End If

    Set tblIt = sessao.findById(IdTabela(mapa("MIRO_PEDIDO_ITEM_LINHA")))
    If Err.Number <> 0 Then
        Resultado lin, "NAO LANCADA", "", "Nao encontrei a grade de itens da MIRO"
        Err.Clear : Exit Sub
    End If
    colPed = IdxColuna(mapa("MIRO_PEDIDO_ITEM_LINHA"))
    colVal = IdxColuna(mapa("MIRO_VALOR_ITEM_LINHA"))

    ' Zerados POR NOTA. obsImposto era zerado so dentro de CorrigirBases,
    ' que so roda quando o saldo esta fora de zero - entao uma nota que
    ' fechava sozinha herdava a observacao da nota anterior e a planilha
    ' registrava uma correcao que nunca aconteceu naquela nota.
    obsQtd = ""
    obsImposto = ""
    obsAviso = ""

    ' 1) le o pedido de cada linha ate a grade acabar
    '
    ' Duas passadas: a grade as vezes so termina de montar depois de um
    ' Enter, entao se a primeira vier vazia confirma a tela e le de novo.
    ' O array cresce conforme a grade entrega linhas - a nota pode ter
    ' quantos pedidos tiver, nao ha numero maximo.
    linhasGrade = 0
    For tentativa = 1 To 2
        If tentativa = 2 Then
            If linhasGrade > 0 Then Exit For
            sessao.findById("wnd[0]").sendVKey 0
            Err.Clear
        End If

        ReDim gradePed(7)
        linhasGrade = 0
        Do While linhasGrade < MAX_ITENS
            Set cel = CelGrade(IdTabela(mapa("MIRO_PEDIDO_ITEM_LINHA")), colPed, linhasGrade)
            If cel Is Nothing Then Exit Do
            pedTela = ""
            pedTela = Trim(cel.Text)
            If Err.Number <> 0 Then Err.Clear : Exit Do
            If pedTela = "" Then Exit Do
            If linhasGrade > UBound(gradePed) Then ReDim Preserve gradePed(linhasGrade + 7)
            gradePed(linhasGrade) = pedTela
            linhasGrade = linhasGrade + 1
        Loop
        Err.Clear
    Next

    If linhasGrade = 0 Then
        DespejarTela "GRADE_VAZIA.txt", "grade de itens vazia apos adotar os pedidos da nota da linha " & lin
        Resultado lin, "NAO LANCADA", "", "A grade de itens voltou vazia depois de adotar os pedidos [" & geoGrade & "] barra: " & MensagemDaBarraCompleta() & " - gerei o GRADE_VAZIA.txt com a tela"
        Err.Clear : Exit Sub
    End If

    If linhasGrade >= MAX_ITENS Then
        Resultado lin, "NAO LANCADA", "", "A grade devolveu " & linhasGrade & " linhas e bateu no teto de seguranca (" & MAX_ITENS & ") - isso nao parece uma nota real. Lancar manualmente."
        Err.Clear : Exit Sub
    End If

    ' 2) um pedido em mais de uma linha exige rateio DENTRO do pedido,
    '    e isso a planilha nao informa
    For n = 0 To linhasGrade - 1
        contagem = 0
        For m = 0 To linhasGrade - 1
            If gradePed(m) = gradePed(n) Then contagem = contagem + 1
        Next
        If contagem > 1 Then
            Resultado lin, "NAO LANCADA", "", "O pedido " & gradePed(n) & " ocupa " & contagem & " linhas de item na MIRO. O rateio dentro do proprio pedido nao esta na planilha e nao se deduz - lancar manualmente."
            Err.Clear : Exit Sub
        End If
    Next

    ' 3) a grade e o rateio tem que descrever o mesmo conjunto
    If linhasGrade <> UBound(pedidos) + 1 Then
        Resultado lin, "NAO LANCADA", "", "A grade trouxe " & linhasGrade & " linhas de item e o rateio tem " & (UBound(pedidos) + 1) & " pedidos"
        Err.Clear : Exit Sub
    End If

    ' Coluna da QUANTIDADE - descoberta agora, com a grade ja montada.
    ' Se nao estiver no mapa, procura pelo nome tecnico (DRSEG-MENGE) e
    ' anota o ID achado no IDS_DESCOBERTOS.txt.
    colQtd = -1
    If mapa("MIRO_QTD_ITEM_LINHA") <> "" Then
        colQtd = IdxColuna(mapa("MIRO_QTD_ITEM_LINHA"))
    Else
        colQtd = AcharColunaGrade(IdTabela(mapa("MIRO_PEDIDO_ITEM_LINHA")), "DRSEG-MENGE")
    End If
    ' O indice vindo do mapa/descoberta falhou na nota de 9 pedidos
    ' (GetCell(0,4) falhou, 19/10/2026). Pergunta a propria tabela qual
    ' coluna e a quantidade, pelo nome tecnico/titulo; o que ela disser
    ' vale mais que o mapa.
    colQtdNome = ColunaQtdPorNome(IdTabela(mapa("MIRO_PEDIDO_ITEM_LINHA")))
    If colQtdNome >= 0 Then colQtd = colQtdNome
    ' A descoberta por nome falhou numa nota de 9 pedidos (19/10/2026) e o
    ' script seguiu em silencio SEM mexer na quantidade - a linha ficou
    ' sem quantidade, fora da fatura. A coluna e a 4 (txtDRSEG-MENGE[4,L],
    ' achada em 05/10/2026), entao e o ultimo recurso antes de desistir.
    If colQtd < 0 Then
        colQtd = 4
        obsAviso = obsAviso & "coluna da quantidade nao foi descoberta e foi assumida como 4; "
    End If

    ' 4) escreve o valor de cada pedido na linha DELE
    For n = 0 To linhasGrade - 1
        idxLote = -1
        For m = 0 To UBound(pedidos)
            If Trim(pedidos(m)) = gradePed(n) Then idxLote = m
        Next
        If idxLote < 0 Then
            Resultado lin, "NAO LANCADA", "", "A linha " & n & " da grade traz o pedido " & gradePed(n) & ", que nao esta no rateio desta nota"
            Err.Clear : Exit Sub
        End If
        ' QUANTIDADE: SEMPRE 1 quando o campo for digitavel (regra de
        ' 19/10/2026). Nao importa o que a MIRO trouxe (vazio, 0, 0,000):
        ' se o campo aceita digitacao, vai 1. Se nao for digitavel, nao
        ' encosta. Vai ANTES do valor: informar quantidade faz o SAP
        ' recalcular o montante da linha, e o valor liquido tem que ser a
        ' ultima coisa escrita.
        If colQtd >= 0 Then
            Set cel = CelQtd(IdTabela(mapa("MIRO_PEDIDO_ITEM_LINHA")), colQtd, n)
            If cel Is Nothing Then
                DespejarColunas "QTD_DEBUG.txt", IdTabela(mapa("MIRO_PEDIDO_ITEM_LINHA")), colQtd, n
                Resultado lin, "NAO LANCADA", "", "Nao consegui ler a quantidade da linha " & n & " (pedido " & gradePed(n) & ") coluna " & colQtd & " [" & geoGrade & "] - gerei o QTD_DEBUG.txt com as colunas da grade"
                Err.Clear : Exit Sub
            End If
            qtdTxt = ""
            qtdTxt = Trim(cel.Text)
            If Err.Number <> 0 Then
                Err.Clear
                Resultado lin, "NAO LANCADA", "", "Nao consegui ler a quantidade da linha " & n & " (pedido " & gradePed(n) & ")"
                Exit Sub
            End If
            qtdEdita = False
            qtdEdita = cel.Changeable
            If Err.Number <> 0 Then Err.Clear : qtdEdita = True
            If qtdEdita Then
                cel.text = "1"
                If Err.Number <> 0 Then
                    Err.Clear
                    Resultado lin, "NAO LANCADA", "", "A quantidade do pedido " & gradePed(n) & " e digitavel e nao consegui informar 1"
                    Exit Sub
                End If
                If QtdVazia(qtdTxt) Then
                    obsQtd = obsQtd & "quantidade do pedido " & gradePed(n) & " veio vazia e foi informada como 1; "
                End If
            End If
        End If

        Set cel = CelGrade(IdTabela(mapa("MIRO_PEDIDO_ITEM_LINHA")), colVal, n)
        If cel Is Nothing Then
            Resultado lin, "NAO LANCADA", "", "Nao consegui chegar na linha " & n & " da grade de itens [" & geoGrade & "]"
            Err.Clear : Exit Sub
        End If
        cel.text = Trim(valores(idxLote))
        If Err.Number <> 0 Then
            Resultado lin, "NAO LANCADA", "", "Falha ao informar o valor liquido do pedido " & gradePed(n) & " (" & Trim(valores(idxLote)) & ")"
            Err.Clear : Exit Sub
        End If
    Next

    ' 4b) CONFERE o valor de CADA linha da grade, relendo da tela.
    ' Escrever em celula de tabela do SAP nao garante que o valor ficou:
    ' rolagem de pagina e round trip podem descartar o digitado (caso
    ' NF 1001 Operadora C, 19/10/2026: uma linha ficou sem o liquido e o saldo
    ' fechou igual ao valor dela). Se a linha nao guardou, escreve de
    ' novo (ate 3 vezes); se ainda assim nao guardar, retem a nota
    ' dizendo QUAL pedido - nunca segue com linha sem valor.
    For n = 0 To linhasGrade - 1
        idxLote = -1
        For m = 0 To UBound(pedidos)
            If Trim(pedidos(m)) = gradePed(n) Then idxLote = m
        Next
        If idxLote >= 0 Then
            ' quantidade tambem e relida: linha sem quantidade nao entra na
            ' fatura (MIRO mostra "itens 8 de 9") e o saldo fica igual ao
            ' imposto da linha. Vazia ou zerada = informa 1, antes do valor.
            If colQtd >= 0 Then
                Set cel = CelQtd(IdTabela(mapa("MIRO_PEDIDO_ITEM_LINHA")), colQtd, n)
                qtdTxt = "?"
                If Not (cel Is Nothing) Then
                    qtdTxt = Trim(cel.Text)
                    If Err.Number <> 0 Then Err.Clear : qtdTxt = "?"
                End If
                If qtdTxt <> "?" And QtdVazia(qtdTxt) Then
                    cel.text = "1"
                    Err.Clear
                    obsQtd = obsQtd & "quantidade do pedido " & gradePed(n) & " estava [" & qtdTxt & "] na releitura e foi informada como 1; "
                End If
            End If
            lidoValor = ""
            For tentativa = 1 To 3
                Set cel = CelGrade(IdTabela(mapa("MIRO_PEDIDO_ITEM_LINHA")), colVal, n)
                lidoValor = "?"
                If Not (cel Is Nothing) Then
                    lidoValor = Trim(cel.Text)
                    If Err.Number <> 0 Then Err.Clear : lidoValor = "?"
                End If
                If EmCentavos(lidoValor) = EmCentavos(valores(idxLote)) And EmCentavos(lidoValor) <> CENTAVOS_INVALIDO Then Exit For
                If Not (cel Is Nothing) Then
                    cel.text = Trim(valores(idxLote))
                    Err.Clear
                    obsAviso = obsAviso & "valor do pedido " & gradePed(n) & " nao tinha ficado na grade (lido [" & lidoValor & "]) e foi informado de novo; "
                End If
            Next
            If EmCentavos(lidoValor) <> EmCentavos(valores(idxLote)) Or EmCentavos(lidoValor) = CENTAVOS_INVALIDO Then
                DespejarGrade "SALDO_FORA.txt", linhasGrade, colPed, colVal, colQtd
                Resultado lin, "NAO LANCADA", "", "O valor liquido do pedido " & gradePed(n) & " (linha " & (n + 1) & " da grade) nao ficou gravado: esperava " & Trim(valores(idxLote)) & " e a tela mostra [" & lidoValor & "] mesmo depois de 3 tentativas. Nada foi lancado - gerei o SALDO_FORA.txt."
                Err.Clear : Exit Sub
            End If
        End If
    Next

    sessao.findById("wnd[0]").sendVKey 0
    erroMsg = MensagemDaBarra()
    If erroMsg <> "" Then
        Resultado lin, "NAO LANCADA", "", "Itens: " & erroMsg
        Err.Clear : Exit Sub
    End If

    ' --- aba Detalhe: tipo de NF --------------------------------------
    Selecionar mapa("MIRO_ABA_DETALHE")
    Definir    mapa("MIRO_TIPO_NF"), tipoNF
    sessao.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        Resultado lin, "NAO LANCADA", "", "Falha ao informar o tipo de NF: " & Err.Description
        Err.Clear : Exit Sub
    End If

    ' --- tela da Nota Fiscal ------------------------------------------
    sessao.findById(mapa("MIRO_BOTAO_NOTA_FISCAL")).press
    If Err.Number <> 0 Then
        Resultado lin, "NAO LANCADA", "", "Falha ao abrir a tela de Nota Fiscal: " & Err.Description
        Err.Clear : Exit Sub
    End If

    Selecionar mapa("NF_ABA")
    Definir mapa("NF_DATA_PROCMTO"),  dataProc
    Definir mapa("NF_TEMPO_PROCMTO"), tempoProc
    Definir mapa("NF_NUM_LOG"),       numLog
    Definir mapa("NF_TP_EMISSAO"),    tpEmissao
    Definir mapa("NF_NUM_ALEATORIO"), numAleat
    sessao.findById("wnd[0]").sendVKey 0

    If Err.Number <> 0 Then
        Resultado lin, "NAO LANCADA", "", "Falha ao preencher a aba Nota Fiscal: " & Err.Description
        Err.Clear : Exit Sub
    End If

    erroMsg = MensagemDaBarra()
    If erroMsg <> "" Then
        Resultado lin, "NAO LANCADA", "", "Aba Nota Fiscal: " & erroMsg
        Err.Clear : Exit Sub
    End If

    If mapa("NF_MODELO_NF") <> "" Then
        Dim modeloTela
        modeloTela = ""
        modeloTela = Trim(sessao.findById(mapa("NF_MODELO_NF")).text)
        If Err.Number = 0 And modeloTela <> "" And modeloTela <> "62" Then
            Resultado lin, "NAO LANCADA", "", "Modelo da NF na tela = " & modeloTela & " (esperado 62). Parametrizacao incorreta - apontado como erro, nao corrigido."
            Err.Clear : Exit Sub
        End If
        Err.Clear
    End If

    sessao.findById(mapa("NF_BOTAO_VOLTAR")).press
    Err.Clear

    ' ------------------------------------------------------------------
    ' SALDO: tem que estar EXATAMENTE em 0,00
    ' ------------------------------------------------------------------
    ' Saldo diferente de zero significa que o que foi informado nao fecha
    ' com o que o pedido comporta. Nesse caso nao se lanca - nem se tenta
    ' simular. Se o campo do saldo nao puder ser lido, tambem nao se
    ' lanca: sem conseguir provar que esta zerado, a nota fica retida.
    ' A busca automatica procura primeiro pelo campo de TEXTO. Sem isso
    ' ela acharia antes o rotulo lblRM08M-DIFFERENZ, cujo texto e a
    ' palavra "Saldo" - sem digito nenhum - e toda nota seria retida.
    idSaldo = mapa("MIRO_SALDO")
    If idSaldo = "" Then idSaldo = BuscarId(sessao.findById("wnd[0]"), "txtRM08M-DIFFERENZ", 0)
    If idSaldo = "" Then idSaldo = BuscarId(sessao.findById("wnd[0]"), "txt%DIFFERENZ", 0)
    Err.Clear

    ' rotulo nao serve: precisa ser campo com valor
    If idSaldo <> "" Then
        If InStr(UCase(idSaldo), "/LBL") > 0 Then idSaldo = ""
    End If

    If idSaldo = "" Then
        Resultado lin, "NAO LANCADA", "", "Nao consegui localizar o campo Saldo na tela da MIRO. Sem poder conferir que esta em 0,00, a nota nao e lancada. Mapear MIRO_SALDO no MAPA_CAMPOS.txt."
        Err.Clear : Exit Sub
    End If

    ' deixa registrado o ID descoberto, para virar linha do mapa depois
    If mapa("MIRO_SALDO") = "" Then AnotarSaldoId idSaldo

    saldoTxt = ""
    saldoTxt = Trim(sessao.findById(idSaldo).text)
    If Err.Number <> 0 Then
        Resultado lin, "NAO LANCADA", "", "Nao consegui ler o campo Saldo (" & idSaldo & "). Nota nao lancada por seguranca."
        Err.Clear : Exit Sub
    End If

    If Not EhZero(saldoTxt) Then
        ' Divergencia de centavos: o SAP arredondou a base de imposto de
        ' algum item e ela deixou de bater com o valor BRUTO do pedido.
        ' Da para devolver o bruto a base - e SO isso. Qualquer coisa
        ' fora da tolerancia nao e arredondamento, e rateio errado: nesse
        ' caso a nota fica retida, como sempre foi.
        tolCent = 3
        If mapa("NF_TOLERANCIA_CENTAVOS") <> "" Then
            If IsNumeric(mapa("NF_TOLERANCIA_CENTAVOS")) Then tolCent = CInt(mapa("NF_TOLERANCIA_CENTAVOS"))
        End If
        Err.Clear

        difCent = EmCentavos(saldoTxt)

        If difCent = CENTAVOS_INVALIDO Then
            Resultado lin, "NAO LANCADA", "", "Nao consegui interpretar o saldo (" & saldoTxt & ") como numero. Nota retida por seguranca."
            Err.Clear : Exit Sub
        End If

        If Abs(difCent) > tolCent Then
            ' Diferenca grande: fotografa a grade de itens. Saber QUAL linha
            ' ficou diferente e o que resolve; o numero do saldo sozinho nao.
            DespejarTela "SALDO_FORA.txt", "saldo em " & saldoTxt & " acima da tolerancia (nota da linha " & lin & ")"
            DespejarGrade "SALDO_FORA.txt", linhasGrade, colPed, colVal, colQtd
            Resultado lin, "NAO LANCADA", "", "Saldo da MIRO em " & saldoTxt & " - acima da tolerancia de " & tolCent & " centavos, entao nao e arredondamento. Gerei o SALDO_FORA.txt com a grade de itens."
            Err.Clear : Exit Sub
        End If

        If Celula(lin, C_BRUTOS) = "" Then
            Resultado lin, "NAO LANCADA", "", "Saldo da MIRO em " & saldoTxt & " (dentro da tolerancia), mas a coluna VALORES BRUTOS desta linha esta vazia - sem o bruto do pedido nao da para corrigir a base do imposto."
            Err.Clear : Exit Sub
        End If

        falhou = CorrigirBases(pedidos, brutos, valores, gradePed, tolCent)
        If falhou <> "" Then
            Resultado lin, "NAO LANCADA", "", "Saldo em " & saldoTxt & ". Tentei corrigir a base do imposto e parei: " & falhou
            Err.Clear : Exit Sub
        End If

        saldoTxt = ""
        saldoTxt = Trim(sessao.findById(idSaldo).text)
        If Err.Number <> 0 Then
            Resultado lin, "NAO LANCADA", "", "Corrigi a base do imposto (" & obsImposto & ") mas depois nao consegui reler o saldo. Nota retida."
            Err.Clear : Exit Sub
        End If

        If Not EhZero(saldoTxt) Then
            Resultado lin, "NAO LANCADA", "", "Corrigi a base do imposto (" & obsImposto & ") mas o saldo ficou em " & saldoTxt & " em vez de 0,00. Nota retida."
            Err.Clear : Exit Sub
        End If
    End If

    ' ------------------------------------------------------------------
    ' Simular - a trava que decide automatico x humano
    ' ------------------------------------------------------------------
    If mapa("MIRO_BOTAO_SIMULAR") = "" Then
        Resultado lin, "NAO LANCADA", "", "Botao Simular nao mapeado - lancamento sem simulacao nao e permitido"
        Err.Clear : Exit Sub
    End If

    sessao.findById(mapa("MIRO_BOTAO_SIMULAR")).press
    If Err.Number <> 0 Then
        Resultado lin, "NAO LANCADA", "", "Falha ao simular: " & Err.Description
        Err.Clear : Exit Sub
    End If

    erroMsg = MensagemDaBarra()
    If erroMsg <> "" Then
        FecharPopups
        Resultado lin, "NAO LANCADA", "", "Simulacao: " & erroMsg
        Err.Clear : Exit Sub
    End If

    If Not ExisteJanela(1) Then
        Resultado lin, "NAO LANCADA", "", "A janela de simulacao nao abriu - lancamento interrompido por seguranca"
        Err.Clear : Exit Sub
    End If

    ' ------------------------------------------------------------------
    ' Lancar
    ' ------------------------------------------------------------------
    If modo = "PARK" Then
        botao = mapa("MIRO_BOTAO_PRE_LANCAR")
    Else
        botao = mapa("MIRO_BOTAO_GRAVAR")
    End If

    If botao = "" Then
        FecharPopups
        Resultado lin, "NAO LANCADA", "", "Botao de gravacao nao mapeado para o modo " & modo
        Err.Clear : Exit Sub
    End If

    sessao.findById(botao).press
    If Err.Number <> 0 Then
        FecharPopups
        Resultado lin, "NAO LANCADA", "", "Falha ao gravar: " & Err.Description
        Err.Clear : Exit Sub
    End If

    msgFinal = MensagemDaBarraCompleta()
    docMiro  = ExtrairNumeroDoc(msgFinal)

    ' a barra de status pode truncar a mensagem; o numero do documento
    ' tambem vem como parametro da mensagem
    If docMiro = "" Then
        Dim par
        par = ""
        par = sessao.findById("wnd[0]/sbar").messageParameter
        If Err.Number = 0 Then docMiro = ExtrairNumeroDoc(par)
        Err.Clear
    End If

    If docMiro <> "" Then
        If modo = "PARK" Then
            Resultado lin, "PRE-LANCADA", docMiro, "Aguardando liberacao humana"
        Else
            If obsQtd <> "" Or obsImposto <> "" Or obsAviso <> "" Then
                Resultado lin, "LANCADA", docMiro, "OBSERVACAO: " & obsQtd & obsImposto & obsAviso
            Else
                Resultado lin, "LANCADA", docMiro, ""
            End If
            ' Nota lancada (nao PARK): se esta linha ja chegou da validacao
            ' com a forma de pagamento e o codigo de barras prontos, bipa
            ' na hora - sem isso o boleto so seria anexado numa rodada
            ' separada do LANCAR_BOLETOS.vbs.
            BiparAposLancar lin, docMiro
        End If
    Else
        Resultado lin, "NAO LANCADA", "", "O SAP nao devolveu numero de documento. Retorno: " & msgFinal
    End If

    Err.Clear
    On Error GoTo 0
End Sub

' =====================================================================
'  PLANILHA
' =====================================================================

' Le uma celula como texto. A aba LOTE esta formatada como texto, entao
' .Value ja devolve string - mas CStr protege contra uma celula que o
' usuario tenha reformatado sem querer.
Function Celula(lin, col)
    Dim v
    Celula = ""
    On Error Resume Next
    v = ws.Cells(lin, col).Value
    If Err.Number <> 0 Then Err.Clear : Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    Celula = Trim(CStr(v))
End Function

' Escreve o resultado na propria linha e salva na hora, para que uma
' interrupcao no meio do lote nao perca o que ja foi lancado.
Sub Resultado(lin, status, doc, erro)
    On Error Resume Next
    ws.Cells(lin, C_STATUS).Value  = status
    ws.Cells(lin, C_DOCMIRO).Value = doc
    ws.Cells(lin, C_ERRO).Value    = Replace(erro, vbCrLf, " ")
    ws.Cells(lin, C_QUANDO).Value  = CStr(Now)
    ws.Cells(lin, C_QUEM).Value    = quemRodou
    wb.Save
    Err.Clear

    If status = "NAO LANCADA" Then
        totalNOK = totalNOK + 1
    Else
        totalOK = totalOK + 1
    End If

    ' volta ao menu inicial para a proxima nota nao herdar estado
    FecharPopups
    sessao.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    sessao.findById("wnd[0]").sendVKey 0
    If ExisteJanela(1) Then
        sessao.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Err.Number <> 0 Then
            Err.Clear
            FecharPopups
        End If
    End If
    Err.Clear
End Sub

Sub Encerrar(msg)
    On Error Resume Next
    wb.Close False
    excel.Quit
    Err.Clear
    MsgBox msg, vbCritical, "NFCOM"
    WScript.Quit
End Sub

' =====================================================================
'  SAP - AUXILIARES
' =====================================================================

' Procura, dentro de um container do SAP, o primeiro objeto cujo ID
' contenha o fragmento dado. Serve para achar campos sem depender de um
' ID fixo - util quando a tela muda de versao para versao.
Function BuscarId(obj, frag, prof)
    Dim i, filho, res, cnt
    BuscarId = ""
    If prof > 7 Then Exit Function
    On Error Resume Next
    If InStr(UCase(obj.Id), UCase(frag)) > 0 Then
        BuscarId = obj.Id
        Err.Clear
        Exit Function
    End If
    cnt = -1
    cnt = obj.Children.Count
    If Err.Number <> 0 Or cnt < 0 Then Err.Clear : Exit Function
    For i = 0 To cnt - 1
        Set filho = FilhoDe(obj, i)
        If Not (filho Is Nothing) Then
            res = BuscarId(filho, frag, prof + 1)
            If res <> "" Then
                BuscarId = res
                Err.Clear
                Exit Function
            End If
        End If
        Err.Clear
    Next
End Function

' A colecao de filhos do SAP GUI nao aceita indice direto em VBScript:
' Children.Count funciona, mas Children(i) nao devolve nada. O acesso
' correto e por ElementAt(i), com Item(i) como alternativa.
Function FilhoDe(obj, j)
    Set FilhoDe = Nothing
    On Error Resume Next
    Set FilhoDe = obj.Children.ElementAt(j)
    If Err.Number = 0 And Not (FilhoDe Is Nothing) Then Err.Clear : Exit Function
    Err.Clear
    Set FilhoDe = obj.Children.Item(j)
    If Err.Number = 0 And Not (FilhoDe Is Nothing) Then Err.Clear : Exit Function
    Err.Clear
    Set FilhoDe = obj.Children(j)
    If Err.Number <> 0 Then Set FilhoDe = Nothing
    Err.Clear
End Function

' "Esta exatamente zerado?" - conta digito a digito, sem depender de
' separador decimal nem de configuracao regional. Texto sem digito
' nenhum devolve False de proposito: nao deu para provar que e zero.
Function EhZero(t)
    Dim i, c, temDigito
    EhZero = False
    temDigito = False
    For i = 1 To Len(t)
        c = Mid(t, i, 1)
        If c >= "0" And c <= "9" Then
            temDigito = True
            If c <> "0" Then Exit Function
        End If
    Next
    EhZero = temDigito
End Function

' Grava o ID do saldo descoberto num arquivo ao lado, para virar linha
' do MAPA_CAMPOS.txt e a busca nao precisar se repetir.
' =====================================================================
'  CORRECAO DA BASE DO IMPOSTO
' =====================================================================
'  O SAP as vezes arredonda a base de imposto de um item, e ela deixa de
'  bater com o valor BRUTO do pedido. A diferenca reaparece como saldo de
'  alguns centavos na MIRO. A correcao e devolver o bruto para a base.
'
'  E o mesmo caminho que se faz na mao:
'     tela Nota Fiscal -> aba de itens -> F2 no item -> aba Imposto ->
'     marca as duas linhas, corrige BASE e OUTRAS BASES, aperta a
'     calculadora, volta. Repete em cada item que estiver fora.
'
'  De qual pedido e cada item: pelo VALOR. A base de um item tem que
'  bater com um dos brutos do rateio. Nao se usa a posicao na lista -
'  a ordem dos itens na Nota Fiscal nao e garantida.
'
'  Devolve "" quando terminou bem, ou a explicacao do que impediu.
'  NAO grava documento: so mexe na tela.
Function CorrigirBases(pedidos, brutos, valores, gradePed, tolCent)
    Dim n, k, cel, baseTxt, baseCent, dif
    Dim qtd, idTabItens, colFoco, idTabImp, colBase, colOutras, colTipo, confBase, idx
    Dim corrigidos, semImposto, temImposto

    CorrigirBases = "falha nao identificada na correcao da base"
    obsImposto = ""
    On Error Resume Next

    idTabItens = mapa("NF_TABELA_ITENS")
    colFoco    = IdxColuna(mapa("NF_ITEM_COLUNA_FOCO"))
    idTabImp   = mapa("NF_TABELA_IMPOSTO")
    colBase    = IdxColuna(mapa("NF_IMPOSTO_BASE"))
    colOutras  = IdxColuna(mapa("NF_IMPOSTO_OUTRAS_BASES"))
    colTipo    = IdxColuna(mapa("NF_IMPOSTO_TIPO"))

    If idTabItens = "" Or idTabImp = "" Or colFoco < 0 Or colBase < 0 Or colOutras < 0 Then
        CorrigirBases = "os campos da tela de imposto nao estao mapeados (NF_TABELA_ITENS, NF_ITEM_COLUNA_FOCO, NF_TABELA_IMPOSTO, NF_IMPOSTO_BASE, NF_IMPOSTO_OUTRAS_BASES)"
        Exit Function
    End If

    If mapa("NF_BOTAO_CALCULADORA") = "" Then
        CorrigirBases = "o botao da calculadora nao esta mapeado (NF_BOTAO_CALCULADORA)"
        Exit Function
    End If

    qtd = UBound(pedidos) + 1
    For k = 0 To qtd - 1
        If EmCentavos(brutos(k)) = CENTAVOS_INVALIDO Then
            CorrigirBases = "o valor bruto do pedido " & Trim(pedidos(k)) & " (" & Trim(brutos(k)) & ") nao e um numero"
            Exit Function
        End If
    Next

    sessao.findById(mapa("MIRO_BOTAO_NOTA_FISCAL")).press
    If Err.Number <> 0 Then
        CorrigirBases = "nao consegui reabrir a tela de Nota Fiscal: " & Err.Description
        Err.Clear : Exit Function
    End If
    Selecionar mapa("NF_ABA_ITENS")
    Err.Clear

    ' NAO se usa RowCount da tabela para saber quantos itens a nota tem:
    ' um TableControl do SAP reserva linhas de entrada vazias, entao ele
    ' devolve o tamanho da tabela (25, 40...) e nao a quantidade de itens.
    ' Foi assim que uma nota de 3 itens foi barrada por "ter 25".
    ' A conferencia de item sobrando e feita DEPOIS do laco, olhando se
    ' existe uma linha a mais COM CONTEUDO.
    corrigidos = 0
    semImposto = 0

    ' Percorre os pedidos do RATEIO, nao os itens da tela. Para cada um:
    '
    '   - pedido SEM imposto (bruto = liquido): nao abre o item. Nao ha
    '     base para recalcular, e cada abertura e uma ida ao servidor.
    '   - pedido COM imposto: abre o item na MESMA posicao que ele ocupa
    '     na grade da MIRO, e CONFERE pela base que abriu o item certo.
    '
    ' A posicao escolhe qual item abrir; o valor da base PROVA que foi o
    ' item certo. Se a prova falhar, retem - nunca corrige no escuro.
    For n = 0 To qtd - 1
        idx = -1
        For k = 0 To qtd - 1
            If Trim(pedidos(k)) = Trim(gradePed(n)) Then idx = k
        Next
        If idx < 0 Then
            VoltarDaNF
            CorrigirBases = "o pedido " & Trim(gradePed(n)) & " esta na grade da MIRO mas nao no rateio da planilha"
            Err.Clear : Exit Function
        End If

        If EmCentavos(brutos(idx)) = EmCentavos(valores(idx)) Then
            ' pedido sem imposto - nao abre o item
            semImposto = semImposto + 1
        Else
            Set cel = CelGrade(idTabItens, colFoco, n)
            If cel Is Nothing Then
                VoltarDaNF
                CorrigirBases = "nao alcancei o item " & (n + 1) & " na lista da Nota Fiscal [" & geoGrade & "]"
                Err.Clear : Exit Function
            End If

            cel.setFocus
            Err.Clear
            sessao.findById("wnd[0]").sendVKey 2
            If Err.Number <> 0 Then
                VoltarDaNF
                CorrigirBases = "nao consegui abrir o detalhe do item " & (n + 1)
                Err.Clear : Exit Function
            End If

            Selecionar mapa("NF_ABA_IMPOSTO_ITEM")
            Err.Clear

            Set cel = CelGrade(idTabImp, colBase, 0)
            If cel Is Nothing Then
                VoltarDoItem
                CorrigirBases = "nao encontrei a base do imposto no item " & (n + 1) & " [" & geoGrade & "]"
                Err.Clear : Exit Function
            End If

            baseTxt = ""
            baseTxt = Trim(cel.Text)
            Err.Clear
            If baseTxt = "" Then
                baseCent = 0
            Else
                baseCent = EmCentavos(baseTxt)
            End If
            If baseCent = CENTAVOS_INVALIDO Then
                VoltarDoItem
                CorrigirBases = "a base do imposto do item " & (n + 1) & " (" & baseTxt & ") nao e um numero"
                Err.Clear : Exit Function
            End If

            ' PROVA: a base tem que ser o bruto deste pedido, a menos dos
            ' centavos de arredondamento. Se estiver longe, ou a ordem da
            ' Nota Fiscal nao segue a da grade, ou o rateio esta errado -
            ' nos dois casos corrigir seria corrigir a coisa errada.
            dif = Abs(baseCent - EmCentavos(brutos(idx)))
            If dif > tolCent Then
                VoltarDoItem
                CorrigirBases = "abri o item " & (n + 1) & " esperando o pedido " & Trim(gradePed(n)) & " (bruto " & Trim(brutos(idx)) & ") e a base de la e " & baseTxt & " - diferenca de " & dif & " centavos. Nao da para garantir que e o item certo."
                Err.Clear : Exit Function
            End If

            If dif > 0 Then
                MarcarLinha idTabImp, 0
                MarcarLinha idTabImp, 1

                Set cel = CelGrade(idTabImp, colBase, 0)
                If cel Is Nothing Then
                    VoltarDoItem
                    CorrigirBases = "perdi a base do imposto do item " & (n + 1) & " na hora de corrigir"
                    Err.Clear : Exit Function
                End If
                cel.text = Trim(brutos(idx))
                If Err.Number <> 0 Then
                    VoltarDoItem
                    CorrigirBases = "nao consegui gravar a base do item " & (n + 1)
                    Err.Clear : Exit Function
                End If

                Set cel = CelGrade(idTabImp, colOutras, 1)
                If cel Is Nothing Then
                    VoltarDoItem
                    CorrigirBases = "nao encontrei OUTRAS BASES na segunda linha de imposto do item " & (n + 1)
                    Err.Clear : Exit Function
                End If
                cel.text = Trim(brutos(idx))
                If Err.Number <> 0 Then
                    VoltarDoItem
                    CorrigirBases = "nao consegui gravar OUTRAS BASES do item " & (n + 1)
                    Err.Clear : Exit Function
                End If

                ' o foco e a rolagem sao o round trip que manda a tela
                ' digitada de volta - sem eles a calculadora nao pega
                If colTipo >= 0 Then
                    Set cel = CelGrade(idTabImp, colTipo, 1)
                    If Not (cel Is Nothing) Then
                        cel.setFocus
                        Err.Clear
                    End If
                End If
                RolarTabela idTabImp, 1

                sessao.findById(mapa("NF_BOTAO_CALCULADORA")).press
                If Err.Number <> 0 Then
                    VoltarDoItem
                    CorrigirBases = "nao consegui acionar a calculadora no item " & (n + 1)
                    Err.Clear : Exit Function
                End If

                RolarTabela idTabImp, 0

                Set cel = CelGrade(idTabImp, colBase, 0)
                If cel Is Nothing Then
                    VoltarDoItem
                    CorrigirBases = "corrigi a base do item " & (n + 1) & " mas nao consegui reler para conferir"
                    Err.Clear : Exit Function
                End If
                confBase = ""
                confBase = Trim(cel.Text)
                Err.Clear
                If EmCentavos(confBase) <> EmCentavos(brutos(idx)) Then
                    VoltarDoItem
                    CorrigirBases = "escrevi " & Trim(brutos(idx)) & " na base do item " & (n + 1) & " e acionei a calculadora, mas a base voltou " & confBase & " - o recalculo nao pegou"
                    Err.Clear : Exit Function
                End If

                obsImposto = obsImposto & "item " & (n + 1) & " (pedido " & Trim(gradePed(n)) & "): base " & baseTxt & " corrigida para " & Trim(brutos(idx)) & "; "
                corrigidos = corrigidos + 1
            End If

            sessao.findById(mapa("NF_BOTAO_VOLTAR")).press
            Err.Clear
            Selecionar mapa("NF_ABA_ITENS")
            Err.Clear
        End If
    Next

    ' Sobrou item na Nota Fiscal? Pergunta certa: existe a linha seguinte
    ' e ela tem conteudo? Linha inexistente ou em branco significa que a
    ' lista acabou junto com o rateio, que e o esperado.
    Set cel = CelGrade(idTabItens, colFoco, qtd)
    If Not (cel Is Nothing) Then
        baseTxt = ""
        baseTxt = Trim(cel.Text)
        Err.Clear
        If baseTxt <> "" Then
            VoltarDaNF
            CorrigirBases = "a Nota Fiscal tem mais itens do que os " & qtd & " pedidos do rateio (a linha " & (qtd + 1) & " ainda traz conteudo) - nao da para saber a qual pedido pertence cada base"
            Err.Clear : Exit Function
        End If
    End If
    Err.Clear

    ' Todo item que estava fora foi corrigido - o laco acima passa por
    ' todos, nao para no primeiro. Se nenhum estava fora, a divergencia
    ' do saldo nao vem da base do imposto e nao ha o que fazer aqui.
    If corrigidos = 0 Then
        VoltarDaNF
        CorrigirBases = "percorri os " & qtd & " itens da Nota Fiscal (" & semImposto & " sem imposto) e nenhum dos tributados estava com a base diferente do bruto - a divergencia do saldo vem de outro lugar"
        Err.Clear : Exit Function
    End If

    If semImposto > 0 Then
        obsImposto = "corrigi " & corrigidos & " de " & (qtd - semImposto) & " itens tributados (" & semImposto & " item(ns) sem imposto, nao tocados) - " & obsImposto
    Else
        obsImposto = "corrigi " & corrigidos & " de " & qtd & " itens - " & obsImposto
    End If

    VoltarDaNF
    Err.Clear
    CorrigirBases = ""
End Function

' sai do detalhe do item e depois da tela da Nota Fiscal
' Rola a tabela ate a posicao pedida. Alem de posicionar, e uma ida ao
' servidor - e as vezes e so disso que a tela precisa.
Sub RolarTabela(idTbl, pos)
    Dim tbl
    On Error Resume Next
    Set tbl = sessao.findById(idTbl)
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    tbl.verticalScrollbar.position = pos
    Err.Clear
End Sub

Sub VoltarDoItem()
    On Error Resume Next
    sessao.findById(mapa("NF_BOTAO_VOLTAR")).press
    Err.Clear
    sessao.findById(mapa("NF_BOTAO_VOLTAR")).press
    Err.Clear
End Sub

Sub VoltarDaNF()
    On Error Resume Next
    sessao.findById(mapa("NF_BOTAO_VOLTAR")).press
    Err.Clear
End Sub

' Marca uma linha da tabela - equivale a clicar no quadradinho da linha.
Sub MarcarLinha(idTbl, n)
    Dim tbl
    On Error Resume Next
    Set tbl = sessao.findById(idTbl)
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    tbl.getAbsoluteRow(n).selected = True
    Err.Clear
End Sub

' Converte "1.234,56" / "-0,02" / "0,00 " em centavos inteiros.
' Imune a locale: nao usa CDbl sobre texto com separador.
' Devolve CENTAVOS_INVALIDO quando o texto nao tem numero.
Function EmCentavos(t)
    Dim s, i, c, limpo, pos, parteInt, parteDec, digitos, n
    EmCentavos = CENTAVOS_INVALIDO
    s = Trim(t)
    If s = "" Then Exit Function

    limpo = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If (c >= "0" And c <= "9") Or c = "," Or c = "." Then limpo = limpo & c
    Next
    If limpo = "" Then Exit Function

    ' o ultimo separador com 1 ou 2 digitos depois dele e o decimal
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
        If c >= "0" And c <= "9" Then digitos = digitos & c
    Next
    parteInt = digitos
    If parteInt = "" Then parteInt = "0"
    If Len(parteInt) > 12 Then Exit Function

    If Len(parteDec) = 0 Then parteDec = "00"
    If Len(parteDec) = 1 Then parteDec = parteDec & "0"

    n = CDbl(parteInt) * 100 + CDbl(parteDec)
    If InStr(Trim(t), "-") > 0 Then n = -n
    EmCentavos = n
End Function

' Procura, dentro de uma tabela do SAP, a primeira celula cujo ID
' contenha o fragmento dado, e devolve o INDICE DA COLUNA dela.
' Devolve -1 quando nao acha. Anota o ID no IDS_DESCOBERTOS.txt para
' que a coluna possa ser fixada no MAPA_CAMPOS.txt depois.
Function AcharColunaGrade(idTbl, frag)
    Dim tbl, i, cnt, filho, idf
    AcharColunaGrade = -1
    On Error Resume Next
    Set tbl = sessao.findById(idTbl)
    If Err.Number <> 0 Then Err.Clear : Exit Function
    cnt = -1
    cnt = tbl.Children.Count
    If Err.Number <> 0 Or cnt < 1 Then Err.Clear : Exit Function
    For i = 0 To cnt - 1
        Set filho = FilhoDe(tbl, i)
        If Not (filho Is Nothing) Then
            idf = ""
            idf = filho.Id
            If Err.Number = 0 Then
                If InStr(UCase(idf), UCase(frag)) > 0 Then
                    AcharColunaGrade = IdxColuna(idf)
                    AnotarIdDescoberto "MIRO_QTD_ITEM_LINHA", idf
                    Err.Clear
                    Exit Function
                End If
            End If
            Err.Clear
        End If
    Next
    Err.Clear
End Function

Sub AnotarIdDescoberto(rotulo, id)
    Dim f, limpo
    On Error Resume Next
    limpo = Replace(id, "/app/con[0]/ses[0]/", "")
    limpo = Replace(limpo, "," & "0]", "," & "{L}]")
    Set f = fso.OpenTextFile(pasta & "\IDS_DESCOBERTOS.txt", 8, True)
    If Err.Number <> 0 Then
        Err.Clear
        Set f = fso.CreateTextFile(pasta & "\IDS_DESCOBERTOS.txt", True, False)
        If Err.Number <> 0 Then Err.Clear : Exit Sub
    End If
    f.WriteLine rotulo & "=" & limpo & "    (achado em " & Now & ")"
    f.Close
    Err.Clear
End Sub

Sub AnotarSaldoId(id)
    Dim f
    On Error Resume Next
    Set f = fso.CreateTextFile(pasta & "\SALDO_ID.txt", True, False)
    f.WriteLine "Campo do Saldo encontrado automaticamente em " & Now & "."
    f.WriteLine "Cole a linha abaixo no MAPA_CAMPOS.txt para fixar:"
    f.WriteLine ""
    f.WriteLine "MIRO_SALDO=" & Replace(id, "/app/con[0]/ses[0]/", "")
    f.Close
    Err.Clear
End Sub

' --- tabelas do SAP que precisam de rolagem --------------------------
' O ID de uma celula tem a forma .../tbl<NOME>/<campo>[<coluna>,<linha>].
' Estas duas funcoes separam a tabela e o indice da coluna, para que o
' acesso possa ser feito pelo objeto da tabela - unico jeito de alcancar
' uma linha que esteja fora da area visivel.
' Escreve a tela inteira num arquivo, para diagnostico. Nao altera nada.
' Celula da QUANTIDADE de uma linha. Primeiro pelo GetCell (como todas as
' outras); se a tabela nao entregar a celula, tenta o campo pelo ID direto
' (txtDRSEG-MENGE[coluna,linha visivel]) - o SAP as vezes recusa GetCell
' numa coluna que o findById acha. Guarda em geoGrade o que tentou.
Function CelQtd(idTbl, colIdx, linhaAbs)
    Dim tbl, pos, rel, c, prefixo, k
    Set CelQtd = Nothing
    Set c = CelGrade(idTbl, colIdx, linhaAbs)
    If Not (c Is Nothing) Then Set CelQtd = c : Exit Function
    On Error Resume Next
    Set tbl = sessao.findById(idTbl)
    If Err.Number <> 0 Then Err.Clear : Exit Function
    pos = -1 : pos = tbl.verticalScrollbar.position
    If Err.Number <> 0 Then Err.Clear : pos = 0
    rel = linhaAbs - pos
    For Each prefixo In Array("txt", "ctxt")
        Set c = Nothing
        Set c = sessao.findById(idTbl & "/" & prefixo & "DRSEG-MENGE[" & colIdx & "," & rel & "]")
        If Err.Number = 0 And Not (c Is Nothing) Then
            geoGrade = geoGrade & " / achada por ID " & prefixo & "DRSEG-MENGE[" & colIdx & "," & rel & "]"
            Set CelQtd = c
            Exit Function
        End If
        Err.Clear
    Next
    geoGrade = geoGrade & " / findById MENGE[" & colIdx & "," & rel & "] tambem falhou"
End Function

' Procura, entre as colunas da tabela, a da quantidade - por nome tecnico
' (DRSEG-MENGE) ou titulo (Quantidade). -1 se nao achar.
Function ColunaQtdPorNome(idTbl)
    Dim tbl, total, i, col, nm, tt
    ColunaQtdPorNome = -1
    On Error Resume Next
    Set tbl = sessao.findById(idTbl)
    If Err.Number <> 0 Then Err.Clear : Exit Function
    total = -1
    total = tbl.Columns.Count
    If Err.Number <> 0 Or total < 1 Then Err.Clear : Exit Function
    For i = 0 To total - 1
        Set col = Nothing
        Set col = tbl.Columns.ElementAt(i)
        If Err.Number <> 0 Then Err.Clear : Set col = Nothing
        If Not (col Is Nothing) Then
            nm = "" : tt = ""
            nm = col.Name
            Err.Clear
            tt = col.Title
            Err.Clear
            If InStr(UCase(nm), "MENGE") > 0 Then ColunaQtdPorNome = i : Exit Function
            If UCase(Trim(tt)) = "QUANTIDADE" Then ColunaQtdPorNome = i : Exit Function
        End If
    Next
    Err.Clear
End Function

' Lista todas as colunas da grade (indice, nome, titulo) e o que a
' celula da linha informada tem em cada uma. Para achar a quantidade.
Sub DespejarColunas(nomeArq, idTbl, colTentada, linhaAbs)
    Dim arq, tbl, total, i, col, nm, tt, cel, tp, tx, ch, pos
    On Error Resume Next
    Set arq = fso.CreateTextFile(pasta & "\" & nomeArq, True, False)
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    arq.WriteLine "COLUNAS DA GRADE DE ITENS - " & Now & " - coluna tentada p/ quantidade: " & colTentada & " - linha absoluta " & linhaAbs
    Set tbl = sessao.findById(idTbl)
    total = -1 : total = tbl.Columns.Count
    pos = -1 : pos = tbl.verticalScrollbar.position
    arq.WriteLine "colunas=" & total & " posicao da rolagem=" & pos & " linhas visiveis=" & tbl.VisibleRowCount
    Err.Clear
    For i = 0 To total - 1
        nm = "" : tt = "" : tp = "-" : tx = "" : ch = ""
        Set col = Nothing
        Set col = tbl.Columns.ElementAt(i)
        nm = col.Name : Err.Clear
        tt = col.Title : Err.Clear
        Set cel = Nothing
        Set cel = tbl.GetCell(0, i)
        If Not (cel Is Nothing) Then
            tp = cel.Type : Err.Clear
            tx = cel.Text : Err.Clear
            ch = cel.Changeable : Err.Clear
        End If
        Err.Clear
        arq.WriteLine "  col " & i & " | nome=" & nm & " | titulo=" & tt & " | celula(0," & i & "): tipo=" & tp & " digitavel=" & ch & " texto=[" & tx & "]"
    Next
    arq.Close
    Err.Clear
End Sub

' Quantidade "sem quantidade": vazia, ou so zeros (0 / 0,000 / 0,00).
Function QtdVazia(t)
    Dim s
    s = Replace(Replace(Replace(Trim(t), ".", ""), ",", ""), " ", "")
    QtdVazia = (s = "" Or Replace(s, "0", "") = "")
End Function

' Acrescenta ao arquivo a grade de itens INTEIRA (todas as paginas):
' linha, pedido, valor liquido e quantidade como a tela mostra agora.
' O despejo de tela so enxerga a pagina visivel; aqui rola ate o fim.
Sub DespejarGrade(nomeArq, qtdLinhas, colunaPed, colunaVal, colunaQtd)
    Dim arq, k, cP, cV, cQ, tp, tv, tq
    On Error Resume Next
    Set arq = fso.OpenTextFile(pasta & "\" & nomeArq, 8, True)
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    arq.WriteLine String(74, "=")
    arq.WriteLine "GRADE DE ITENS COMPLETA (" & qtdLinhas & " linhas) - linha | pedido | valor | quantidade, como a tela mostra"
    For k = 0 To qtdLinhas - 1
        tp = "?" : tv = "?" : tq = "?"
        Set cP = Nothing : Set cV = Nothing : Set cQ = Nothing
        Set cP = CelGrade(IdTabela(mapa("MIRO_PEDIDO_ITEM_LINHA")), colunaPed, k)
        If Not (cP Is Nothing) Then tp = Trim(cP.Text)
        Err.Clear
        Set cV = CelGrade(IdTabela(mapa("MIRO_PEDIDO_ITEM_LINHA")), colunaVal, k)
        If Not (cV Is Nothing) Then tv = Trim(cV.Text)
        Err.Clear
        If colunaQtd >= 0 Then
            Set cQ = CelGrade(IdTabela(mapa("MIRO_PEDIDO_ITEM_LINHA")), colunaQtd, k)
            If Not (cQ Is Nothing) Then tq = Trim(cQ.Text)
            Err.Clear
        End If
        arq.WriteLine "  " & (k + 1) & " | " & tp & " | valor [" & tv & "] | qtd [" & tq & "]"
    Next
    arq.Close
    Err.Clear
End Sub

Sub DespejarTela(nomeArq, motivo)
    Dim arq, cx
    On Error Resume Next
    Set arq = fso.CreateTextFile(pasta & "\" & nomeArq, True, False)
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    arq.WriteLine "DESPEJO DE TELA - " & Now
    arq.WriteLine "Motivo: " & motivo
    arq.WriteLine "Transacao: " & sessao.Info.Transaction
    arq.WriteLine "Barra de status: " & MensagemDaBarraCompleta()
    arq.WriteLine String(74, "-")
    Err.Clear
    cx = 0
    EscreverRamo arq, sessao.findById("wnd[0]"), 0, cx
    If ExisteJanela(1) Then
        arq.WriteLine String(74, "-")
        arq.WriteLine "JANELA wnd[1] ABERTA:"
        EscreverRamo arq, sessao.findById("wnd[1]"), 0, cx
    End If
    arq.WriteLine String(74, "-")
    arq.WriteLine "objetos percorridos: " & cx
    arq.Close
    Err.Clear
End Sub

Sub EscreverRamo(arq, obj, prof, cx)
    Dim i, filho, cnt, t, tipo, id
    If prof > 14 Then Exit Sub
    On Error Resume Next
    id = ""
    id = obj.Id
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    cx = cx + 1
    tipo = "?"
    tipo = obj.Type
    Err.Clear
    t = ""
    t = obj.Text
    Err.Clear
    If Trim(t) <> "" Then
        arq.WriteLine "  " & tipo & " | " & Replace(id, "/app/con[0]/ses[0]/", "") & " | [" & t & "]"
    End If
    cnt = -1
    cnt = obj.Children.Count
    Err.Clear
    If cnt < 1 Then Exit Sub
    For i = 0 To cnt - 1
        Set filho = FilhoDe(obj, i)
        If Not (filho Is Nothing) Then EscreverRamo arq, filho, prof + 1, cx
        Err.Clear
    Next
End Sub

Function IdTabela(idCampo)
    IdTabela = Left(idCampo, InStrRev(idCampo, "/") - 1)
End Function

Function IdxColuna(idCampo)
    Dim a, b
    IdxColuna = -1
    a = InStrRev(idCampo, "[")
    If a = 0 Then Exit Function
    b = InStr(a, idCampo, ",")
    If b = 0 Then Exit Function
    IdxColuna = CInt(Mid(idCampo, a + 1, b - a - 1))
End Function

' Devolve a celula (coluna, linha absoluta) de uma tabela, rolando ate
' ela se necessario. Nothing quando a linha nao existe.
'
' Recebe o ID da tabela, e nao o objeto: toda rolagem e uma ida ao
' servidor e invalida o objeto anterior, por isso ele e re-buscado.
'
' A rolagem para frente e feita DE PAGINA EM PAGINA. Um salto direto
' para uma linha alem do fim da area ja montada e recusado pelo SAP -
' foi isso que impediu de chegar na linha 8 da janela de selecao de
' pedidos. Andando de pagina em pagina o SAP monta a proxima faixa.
Function CelGrade(idTbl, colIdx, linhaAbs)
    Dim tbl, pos, vis, novo, voltas
    Set CelGrade = Nothing
    geoGrade = ""
    On Error Resume Next

    Set tbl = sessao.findById(idTbl)
    If Err.Number <> 0 Then Err.Clear : Exit Function

    pos = -1
    pos = tbl.verticalScrollbar.position
    If Err.Number <> 0 Then                      ' tabela sem barra de rolagem
        Err.Clear
        geoGrade = "tabela sem rolagem"
        Set CelGrade = tbl.GetCell(linhaAbs, colIdx)
        If Err.Number <> 0 Then Err.Clear : Set CelGrade = Nothing
        Exit Function
    End If

    vis = 0
    vis = tbl.VisibleRowCount
    If Err.Number <> 0 Or vis < 1 Then Err.Clear : vis = 1
    geoGrade = "visiveis=" & vis & " posicao=" & pos

    ' para tras: salto direto e sempre aceito
    If linhaAbs < pos Then
        tbl.verticalScrollbar.position = linhaAbs
        Err.Clear
        Set tbl = sessao.findById(idTbl)
        If Err.Number <> 0 Then Err.Clear : Exit Function
        pos = tbl.verticalScrollbar.position
        Err.Clear
    End If

    ' para frente: uma pagina de cada vez
    voltas = 0
    Do While linhaAbs >= pos + vis
        voltas = voltas + 1
        If voltas > 60 Then
            geoGrade = geoGrade & " / travou rolando (60 paginas)"
            Exit Function
        End If
        tbl.verticalScrollbar.position = pos + vis
        If Err.Number <> 0 Then
            Err.Clear
            geoGrade = geoGrade & " / o SAP recusou rolar para " & (pos + vis)
            Exit Function
        End If
        Set tbl = sessao.findById(idTbl)
        If Err.Number <> 0 Then Err.Clear : Exit Function
        novo = -1
        novo = tbl.verticalScrollbar.position
        If Err.Number <> 0 Then Err.Clear : Exit Function
        vis = tbl.VisibleRowCount
        If Err.Number <> 0 Or vis < 1 Then Err.Clear : vis = 1
        If novo <= pos Then
            geoGrade = geoGrade & " / parou de rolar em " & novo & " (fim da tabela)"
            Exit Function
        End If
        pos = novo
        geoGrade = "visiveis=" & vis & " posicao=" & pos & " paginas=" & voltas
    Loop

    Set CelGrade = tbl.GetCell(linhaAbs - pos, colIdx)
    If Err.Number <> 0 Then
        Err.Clear
        geoGrade = geoGrade & " / GetCell(" & (linhaAbs - pos) & "," & colIdx & ") falhou"
        Set CelGrade = Nothing
    End If
End Function

Sub Definir(idCampo, valor)
    If idCampo = "" Then Exit Sub
    If valor = "" Then Exit Sub
    sessao.findById(idCampo).text = valor
End Sub

Sub Selecionar(idAba)
    If idAba = "" Then Exit Sub
    On Error Resume Next
    sessao.findById(idAba).select
    Err.Clear
End Sub

Function ExisteJanela(n)
    Dim w
    ExisteJanela = False
    On Error Resume Next
    Set w = sessao.findById("wnd[" & n & "]")
    If Err.Number = 0 Then ExisteJanela = True
    Err.Clear
End Function

Sub FecharPopups()
    Dim guarda
    guarda = 0
    On Error Resume Next
    Do While ExisteJanela(1) And guarda < 6
        sessao.findById("wnd[1]").close
        If Err.Number <> 0 Then Err.Clear : Exit Do
        guarda = guarda + 1
    Loop
    Err.Clear
End Sub

' Mensagem da barra de status SOMENTE se for erro (E), aviso (W) ou
' abend (A). Sucesso (S) devolve string vazia.
' So ERRO (E) e ABEND (A) retem a nota. AVISO (W) e informacao (I) sao
' registrados e a nota segue: aviso nao impede o SAP de gravar, e a
' trava do saldo mais a simulacao continuam valendo depois. Tratar W
' como fatal reteve uma nota boa por "vencimento liquido no passado".
Function MensagemDaBarra()
    Dim tipo, txt
    MensagemDaBarra = ""
    On Error Resume Next
    tipo = sessao.findById("wnd[0]/sbar").messageType
    txt  = sessao.findById("wnd[0]/sbar").text
    If Err.Number <> 0 Then Err.Clear : Exit Function
    If tipo = "E" Or tipo = "A" Then
        MensagemDaBarra = tipo & ": " & txt
    ElseIf tipo = "W" Or tipo = "I" Then
        If Trim(txt) <> "" Then
            If InStr(obsAviso, Trim(txt)) = 0 Then
                obsAviso = obsAviso & "AVISO DO SAP (" & tipo & "): " & Trim(txt) & "; "
            End If
        End If
    End If
    Err.Clear
End Function

Function MensagemDaBarraCompleta()
    MensagemDaBarraCompleta = ""
    On Error Resume Next
    MensagemDaBarraCompleta = sessao.findById("wnd[0]/sbar").text
    Err.Clear
End Function

' "Documento 5100000123 criado" -> 5100000123
Function ExtrairNumeroDoc(txt)
    Dim re, m
    ExtrairNumeroDoc = ""
    Set re = New RegExp
    re.Pattern = "\d{8,12}"
    re.Global = True
    Set m = re.Execute(txt)
    If m.Count > 0 Then ExtrairNumeroDoc = m(0).Value
End Function

' ---------------------------------------------------------------------
' A planilha LOTE e compartilhada entre este executor (NFCOM) e o
' LANCAR_SERVICOS.vbs (NFS-e e recibos). Quem separa um do outro e a
' coluna TIPO DE NF. Os tipos deste script vem do MAPA_CAMPOS.txt em
' TIPOS_NFCOM; sem a chave, vale TA, que e o que o lote sempre usou.
' ---------------------------------------------------------------------
Function MinhaLinha(tipo)
    Dim lista, k, t, aceitos
    MinhaLinha = False
    t = UCase(Trim(tipo))
    If t = "" Then Exit Function
    aceitos = mapa("TIPOS_NFCOM")
    If Trim(aceitos) = "" Then aceitos = "TA"
    lista = Split(UCase(aceitos), "|")
    For k = 0 To UBound(lista)
        If Trim(lista(k)) = t Then MinhaLinha = True
    Next
End Function

Function CarregarMapa(caminho, dic)
    Dim f, l, p
    CarregarMapa = False
    If Not fso.FileExists(caminho) Then Exit Function
    Set f = fso.OpenTextFile(caminho, ForReading)
    Do Until f.AtEndOfStream
        l = Trim(f.ReadLine)
        If l <> "" And Left(l, 1) <> "#" Then
            p = InStr(l, "=")
            If p > 1 Then dic(Trim(Left(l, p - 1))) = Trim(Mid(l, p + 1))
        End If
    Loop
    f.Close
    CarregarMapa = (dic.Count > 0)
End Function

' =====================================================================
'  BIPAGEM INLINE - integrada em 09/10/2026, a pedido: "quando eu
'  pedir para lancar o lote e a nota ja tiver validada e com o codigo
'  de barras no lote, quero que o script lance a nota e a bipe".
'
'  Chamada automaticamente logo apos cada nota ser LANCADA (nao PARK),
'  com doc MIRO em maos. Se a linha ja tiver FORMA PGTO preenchida (e,
'  quando a forma usa boleto, o codigo de barras ou a linha digitavel
'  tambem), volta no documento contabil ja gravado (FB02) e grava forma
'  de pagamento + codigo de barras - exatamente a mesma sequencia do
'  LANCAR_BOLETOS.vbs (incluindo a correcao de 09/10: confirmar com
'  Enter, duas vezes, o aviso "Vencimento do codigo de barras nao
'  confere", ALI na tela, antes de gravar - nao depois).
'
'  Se o boleto ainda nao chegou na planilha, ou nao ha mapa de bipagem
'  carregado, a nota continua LANCADA normalmente e so fica sem bipar -
'  o LANCAR_BOLETOS.vbs pega ela depois, numa rodada separada. Nada
'  aqui pode reter ou desfazer o lancamento MIRO, que ja aconteceu.
' =====================================================================
Sub BiparAposLancar(lin, docMiroTxt)
    Dim forma, barras, linhaDig, semBoleto, montanteCent, ano, miroDig
    Dim docFI, erro, obs, valorBoletoCent

    If Celula(lin, C_BIPSTATUS) <> "" Then Exit Sub   ' ja bipada - protecao
    If Not bipMapaOk Then Exit Sub                     ' sem mapa, nao tenta

    forma = UCase(Trim(Celula(lin, C_FORMA)))
    If forma = "" Then Exit Sub                        ' nada preenchido ainda - nao e erro

    If Not Bip_FormaValida(forma) Then
        ResultadoBip lin, "NAO BIPADA", "", "Forma de pagamento [" & forma & "] fora da lista aceita (" & Bip_FormasAceitas() & ")."
        Exit Sub
    End If

    semBoleto = Bip_FormaSemBoleto(forma)
    barras    = Bip_SoDigitos(Celula(lin, C_BARRAS))
    linhaDig  = Bip_SoDigitos(Celula(lin, C_LINHADIG))
    obs = ""

    If Not semBoleto Then
        If barras = "" And linhaDig = "" Then Exit Sub   ' boleto ainda nao chegou - fica para a proxima rodada
        If barras = "" And (Len(linhaDig) = 47 Or Len(linhaDig) = 48) Then
            erro = ""
            barras = Bip_LinhaParaBarra(linhaDig, erro)
            If barras = "" Then
                ResultadoBip lin, "NAO BIPADA", "", "Linha digitavel invalida: " & erro
                Exit Sub
            End If
            Bip_GravarTexto lin, C_BARRAS, barras
            obs = obs & "Codigo de barras derivado da linha digitavel. "
        End If
        If Len(barras) <> 44 Then
            ResultadoBip lin, "NAO BIPADA", "", "Codigo de barras deve ter 44 digitos; veio com " & Len(barras) & "."
            Exit Sub
        End If
        If Not Bip_DVBarraOK(barras) Then
            ResultadoBip lin, "NAO BIPADA", "", "Digito verificador do codigo de barras nao confere: " & barras
            Exit Sub
        End If
    End If

    montanteCent = EmCentavos(Celula(lin, C_MONTANTE))
    If Not semBoleto And montanteCent <> CENTAVOS_INVALIDO And montanteCent > 0 Then
        valorBoletoCent = Bip_ValorDaBarra(barras)
        If valorBoletoCent <> 0 And valorBoletoCent <> montanteCent Then
            ResultadoBip lin, "NAO BIPADA", "", "O valor dentro do codigo de barras (" & Bip_Formatar(valorBoletoCent) & _
                     ") nao bate com o MONTANTE da planilha (" & Bip_Formatar(montanteCent) & "). Boleto da nota errada?"
            Exit Sub
        End If
    End If

    miroDig = Bip_SoDigitos(docMiroTxt)
    ano = AnoDaDataBip(Celula(lin, C_DTFATURA))
    If Len(miroDig) < 8 Or Len(ano) <> 4 Then
        ResultadoBip lin, "NAO BIPADA", "", "Nao consegui montar a chave de origem (MIRO " & docMiroTxt & " / data " & Celula(lin, C_DTFATURA) & ")."
        Exit Sub
    End If

    ' ---------- passo 1: MIRO -> documento contabil ----------
    docFI = ""
    erro  = ""
    docFI = Bip_AcharDocumentoFI(Celula(lin, C_EMPRESA), ano, miroDig, erro)
    If docFI = "" Then
        ResultadoBip lin, "NAO BIPADA", "", erro
        Exit Sub
    End If
    Bip_GravarTexto lin, C_DOCCONT, docFI

    ' ---------- passo 2: FB02, linha do fornecedor ----------
    erro = ""
    If Not Bip_AbrirDocumento(mapa("TCODE_ALTERA"), docFI, Celula(lin, C_EMPRESA), ano, erro) Then
        ResultadoBip lin, "NAO BIPADA", docFI, erro
        Exit Sub
    End If

    Dim nLinha, montanteLinha
    nLinha = -1 : montanteLinha = CENTAVOS_INVALIDO : erro = ""
    If Not Bip_AcharLinhaFornecedor(nLinha, montanteLinha, erro) Then
        ResultadoBip lin, "NAO BIPADA", docFI, erro
        Exit Sub
    End If
    If montanteLinha <> CENTAVOS_INVALIDO And montanteCent <> CENTAVOS_INVALIDO Then
        If Abs(montanteLinha) <> montanteCent Then
            ResultadoBip lin, "NAO BIPADA", docFI, "A linha do fornecedor do documento " & docFI & " vale " & _
                     Bip_Formatar(Abs(montanteLinha)) & " e a planilha diz " & Bip_Formatar(montanteCent) & ". Documento errado - nada foi alterado."
            Exit Sub
        End If
    End If

    ' ---------- passo 3: abre a linha ----------
    On Error Resume Next
    bipGradeAchada.setCurrentCell nLinha, mapa("LINHAS_COL_CHAVE")
    bipGradeAchada.doubleClickCurrentCell
    If Err.Number <> 0 Then
        Err.Clear
        ResultadoBip lin, "NAO BIPADA", docFI, "Nao consegui abrir a linha do fornecedor do documento " & docFI & "."
        Exit Sub
    End If
    On Error GoTo 0

    Dim lidoZL, lidoBC, precisaGravar
    lidoZL = Bip_LerCampo(mapa("ITEM_ZLSCH"))
    lidoBC = Bip_SoDigitos(Bip_LerCampo(mapa("ITEM_BRCDE")))
    precisaGravar = False
    If UCase(Trim(lidoZL)) <> forma Then precisaGravar = True
    If Not semBoleto Then
        If lidoBC <> barras Then precisaGravar = True
    End If

    If Not precisaGravar Then
        Bip_VoltarAoMenu
        ResultadoBip lin, "JA ESTAVA OK", docFI, Trim(obs & "Documento ja tinha os dois campos preenchidos com estes valores.")
        Exit Sub
    End If

    erro = ""
    If Not Bip_Definir(mapa("ITEM_ZLSCH"), forma, erro) Then
        Bip_VoltarAoMenu
        ResultadoBip lin, "NAO BIPADA", docFI, "Nao consegui escrever a forma de pagamento: " & erro
        Exit Sub
    End If
    If Not semBoleto Then
        erro = ""
        If Not Bip_Definir(mapa("ITEM_BRCDE"), barras, erro) Then
            Bip_VoltarAoMenu
            ResultadoBip lin, "NAO BIPADA", docFI, "Nao consegui escrever o codigo de barras: " & erro
            Exit Sub
        End If

        ' Confirma o aviso de vencimento do codigo de barras AQUI, igual a
        ' gravacao manual que funcionou de primeira (nota de exemplo 1234,
        ' 09/10/2026) - dois Enter logo apos o campo, ANTES de gravar.
        On Error Resume Next
        sessao.findById("wnd[0]").sendVKey 0
        sessao.findById("wnd[0]").sendVKey 0
        Err.Clear
        On Error GoTo 0
    End If

    ' ---------- passo 4: salvar ----------
    On Error Resume Next
    sessao.findById(mapa("BOTAO_SALVAR")).press
    If Err.Number <> 0 Then
        Err.Clear
        Bip_VoltarAoMenu
        ResultadoBip lin, "NAO BIPADA", docFI, "Nao consegui acionar o botao de gravar."
        Exit Sub
    End If
    On Error GoTo 0

    Dim txtPopup
    txtPopup = ""
    If ExisteJanela(1) Then
        txtPopup = Bip_TextoDaJanela(1)
        On Error Resume Next
        sessao.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Err.Number <> 0 Then
            Err.Clear
            sessao.findById("wnd[1]/tbar[0]/btn[0]").press
            Err.Clear
        End If
        On Error GoTo 0
        If Trim(txtPopup) <> "" Then obs = obs & "Popup ao gravar: " & txtPopup & " "
    End If

    On Error Resume Next
    sessao.findById("wnd[0]").sendVKey 0
    Err.Clear
    On Error GoTo 0

    Dim tipoMsg, txtMsg
    tipoMsg = Bip_TipoDaBarra()
    txtMsg  = Bip_TextoDaBarra()
    If tipoMsg = "E" Or tipoMsg = "A" Then
        Bip_VoltarAoMenu
        ResultadoBip lin, "NAO BIPADA", docFI, "O SAP recusou a gravacao (" & tipoMsg & "): " & txtMsg
        Exit Sub
    End If
    If Trim(txtMsg) <> "" Then obs = obs & "SAP: " & Trim(txtMsg) & " "

    Bip_VoltarAoMenu

    ' ---------- passo 5: prova - reabre em FB03 e le os dois campos ----------
    Dim provaZL, provaBC
    provaZL = "" : provaBC = "" : erro = ""
    If Not Bip_LerDeVolta(docFI, Celula(lin, C_EMPRESA), ano, provaZL, provaBC, erro) Then
        ResultadoBip lin, "BIPADA SEM CONFERENCIA", docFI, Trim(obs & "Gravou, mas nao consegui reabrir para conferir: " & erro)
        Exit Sub
    End If
    If UCase(Trim(provaZL)) <> forma Then
        ResultadoBip lin, "NAO BIPADA", docFI, Trim(obs & "Depois de gravar, a forma de pagamento no documento esta [" & provaZL & "] e deveria ser [" & forma & "]. CONFIRA O DOCUMENTO " & docFI & ".")
        Exit Sub
    End If
    If Not semBoleto Then
        If Bip_SoDigitos(provaBC) <> barras Then
            ResultadoBip lin, "NAO BIPADA", docFI, Trim(obs & "Depois de gravar, o codigo de barras no documento esta [" & provaBC & "] e deveria ser [" & barras & "]. CONFIRA O DOCUMENTO " & docFI & ".")
            Exit Sub
        End If
    End If
    obs = obs & "Conferido no documento depois de gravar. "

    ResultadoBip lin, "BIPADA", docFI, Trim(obs)
End Sub

Sub ResultadoBip(lin, status, docFI, erro)
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
    Bip_VoltarAoMenu
End Sub

Function Bip_AcharDocumentoFI(bukrs, ano, miro, erro)
    Dim awkey, nLin, doc, lixo
    Bip_AcharDocumentoFI = ""
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
    If Not Bip_Definir(mapa("BUSCA_EMPRESA"), bukrs, lixo) Then
        erro = "Nao consegui preencher a empresa na busca: " & lixo
        Exit Function
    End If
    If Not Bip_Definir(mapa("BUSCA_EXERCICIO"), ano, lixo) Then
        erro = "Nao consegui preencher o exercicio na busca: " & lixo
        Exit Function
    End If
    If Not Bip_Definir(mapa("BUSCA_AWKEY"), awkey, lixo) Then
        erro = "Nao consegui preencher a chave de origem na busca: " & lixo
        Exit Function
    End If
    Bip_LimparCampo mapa("BUSCA_BELNR")
    Bip_LimparCampo mapa("BUSCA_XBLNR")

    On Error Resume Next
    sessao.findById("wnd[0]").sendVKey 8
    If Err.Number <> 0 Then
        Err.Clear
        erro = "Nao consegui executar a busca."
        Exit Function
    End If
    On Error GoTo 0

    If Not Bip_AcharGradeNaTela() Then
        erro = "A busca por chave de origem " & awkey & " nao devolveu nenhum documento contabil."
        Bip_VoltarAoMenu
        Exit Function
    End If

    nLin = 0
    On Error Resume Next
    nLin = bipGradeAchada.RowCount
    Err.Clear
    On Error GoTo 0

    If nLin <> 1 Then
        erro = "A busca por chave de origem " & awkey & " devolveu " & nLin & " documentos. Esperado exatamente 1 - nada foi alterado."
        Bip_VoltarAoMenu
        Exit Function
    End If

    doc = Trim(Bip_Ler(bipGradeAchada, 0, mapa("BUSCA_COL_DOC")))
    Bip_VoltarAoMenu

    If Bip_SoDigitos(doc) = "" Then
        erro = "Achei a linha na lista mas nao consegui ler o numero do documento contabil."
        Exit Function
    End If
    Bip_AcharDocumentoFI = Bip_SoDigitos(doc)
End Function

Function Bip_AbrirDocumento(tcode, docFI, bukrs, ano, erro)
    Dim lixo
    Bip_AbrirDocumento = False
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
    If Not Bip_Definir(mapa("DOC_BELNR"), docFI, lixo) Then
        erro = "Nao consegui preencher o numero do documento em " & tcode & ": " & lixo
        Exit Function
    End If
    If Not Bip_Definir(mapa("DOC_BUKRS"), bukrs, lixo) Then
        erro = "Nao consegui preencher a empresa em " & tcode & ": " & lixo
        Exit Function
    End If
    If Not Bip_Definir(mapa("DOC_GJAHR"), ano, lixo) Then
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
        erro = "O SAP abriu um aviso ao entrar no documento " & docFI & ": " & Bip_TextoDaJanela(1)
        FecharPopups
        Bip_VoltarAoMenu
        Exit Function
    End If

    If Bip_TipoDaBarra() = "E" Or Bip_TipoDaBarra() = "A" Then
        erro = "O SAP recusou abrir o documento " & docFI & " em " & tcode & ": " & Bip_TextoDaBarra()
        Bip_VoltarAoMenu
        Exit Function
    End If

    Bip_AbrirDocumento = True
End Function

Function Bip_AcharLinhaFornecedor(nLinha, montanteCent, erro)
    Dim nLin, l, chave, achadas, primeira
    Bip_AcharLinhaFornecedor = False
    erro = ""
    nLinha = -1
    montanteCent = CENTAVOS_INVALIDO

    If Not Bip_AcharGradeNaTela() Then
        erro = "Nao achei a grade de itens do documento."
        Bip_VoltarAoMenu
        Exit Function
    End If

    nLin = 0
    On Error Resume Next
    nLin = bipGradeAchada.RowCount
    Err.Clear
    On Error GoTo 0

    achadas = 0 : primeira = -1
    For l = 0 To nLin - 1
        chave = Trim(Bip_Ler(bipGradeAchada, l, mapa("LINHAS_COL_CHAVE")))
        If Bip_EhChaveFornecedor(chave) Then
            achadas = achadas + 1
            If primeira < 0 Then primeira = l
        End If
    Next

    If achadas = 0 Then
        erro = "O documento nao tem nenhuma linha de fornecedor (chave " & mapa("CHAVES_FORNECEDOR") & ")."
        Bip_VoltarAoMenu
        Exit Function
    End If
    If achadas > 1 Then
        erro = "O documento tem " & achadas & " linhas de fornecedor. Esperado exatamente 1 - nada foi alterado."
        Bip_VoltarAoMenu
        Exit Function
    End If

    nLinha = primeira
    montanteCent = EmCentavos(Bip_Ler(bipGradeAchada, primeira, mapa("LINHAS_COL_MONTANTE")))
    Bip_AcharLinhaFornecedor = True
End Function

Function Bip_LerDeVolta(docFI, bukrs, ano, valZL, valBC, erro)
    Dim nLinha, montanteLinha
    Bip_LerDeVolta = False
    valZL = "" : valBC = "" : erro = ""

    If Not Bip_AbrirDocumento(mapa("TCODE_BUSCA"), docFI, bukrs, ano, erro) Then Exit Function

    nLinha = -1 : montanteLinha = CENTAVOS_INVALIDO
    If Not Bip_AcharLinhaFornecedor(nLinha, montanteLinha, erro) Then Exit Function

    On Error Resume Next
    bipGradeAchada.setCurrentCell nLinha, mapa("LINHAS_COL_CHAVE")
    bipGradeAchada.doubleClickCurrentCell
    If Err.Number <> 0 Then
        Err.Clear
        erro = "nao consegui abrir a linha do fornecedor na conferencia"
        Bip_VoltarAoMenu
        Exit Function
    End If
    On Error GoTo 0

    valZL = Bip_LerCampo(mapa("ITEM_ZLSCH"))
    valBC = Bip_LerCampo(mapa("ITEM_BRCDE"))
    Bip_VoltarAoMenu
    Bip_LerDeVolta = True
End Function

Function Bip_AcharGradeNaTela()
    Set bipGradeAchada = Nothing
    On Error Resume Next
    Bip_Varrer sessao.findById("wnd[0]")
    Err.Clear
    On Error GoTo 0
    Bip_AcharGradeNaTela = Not (bipGradeAchada Is Nothing)
End Function

Sub Bip_Varrer(obj)
    Dim filhos, i, f, tipo
    If Not (bipGradeAchada Is Nothing) Then Exit Sub
    On Error Resume Next
    tipo = obj.Type
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    If tipo = "GuiShell" Then
        If obj.SubType = "GridView" Then
            Set bipGradeAchada = obj
            Exit Sub
        End If
    End If
    Set filhos = Nothing
    Set filhos = obj.Children
    If Err.Number <> 0 Then Err.Clear : Exit Sub
    For i = 0 To filhos.Count - 1
        Set f = Nothing
        Set f = FilhoDe(obj, i)
        If Not (f Is Nothing) Then Bip_Varrer f
        If Not (bipGradeAchada Is Nothing) Then Exit Sub
    Next
    On Error GoTo 0
End Sub

Function Bip_Ler(g, lin, col)
    Bip_Ler = ""
    On Error Resume Next
    Bip_Ler = Trim(g.GetCellValue(lin, col))
    If Err.Number <> 0 Then Err.Clear : Bip_Ler = ""
    On Error GoTo 0
End Function

Function Bip_LerCampo(id)
    Dim o
    Bip_LerCampo = ""
    On Error Resume Next
    Set o = Nothing
    Set o = sessao.findById(id)
    If Err.Number <> 0 Then Err.Clear : Exit Function
    Bip_LerCampo = Trim(o.Text)
    Err.Clear
    On Error GoTo 0
End Function

Function Bip_Definir(id, valor, erro)
    Bip_Definir = False
    erro = ""
    On Error Resume Next
    sessao.findById(id).text = valor
    If Err.Number <> 0 Then
        erro = "campo " & id & " - " & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    Bip_Definir = True
End Function

Sub Bip_LimparCampo(id)
    On Error Resume Next
    sessao.findById(id).text = ""
    Err.Clear
    On Error GoTo 0
End Sub

Sub Bip_GravarTexto(lin, col, valor)
    On Error Resume Next
    ws.Cells(lin, col).NumberFormat = "@"
    ws.Cells(lin, col).Value = valor
    Err.Clear
    On Error GoTo 0
End Sub

Function Bip_TextoDaJanela(n)
    Bip_TextoDaJanela = ""
    On Error Resume Next
    Bip_TextoDaJanela = Trim(sessao.findById("wnd[" & n & "]").Text)
    Err.Clear
    On Error GoTo 0
End Function

Function Bip_TipoDaBarra()
    Bip_TipoDaBarra = ""
    On Error Resume Next
    Bip_TipoDaBarra = sessao.findById("wnd[0]/sbar").messageType
    Err.Clear
    On Error GoTo 0
End Function

Function Bip_TextoDaBarra()
    Bip_TextoDaBarra = ""
    On Error Resume Next
    Bip_TextoDaBarra = Trim(sessao.findById("wnd[0]/sbar").text)
    Err.Clear
    On Error GoTo 0
End Function

Sub Bip_VoltarAoMenu()
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
            If Err.Number <> 0 Then Err.Clear : Exit Do
        End If
        guarda = guarda + 1
    Loop
    Err.Clear
    On Error GoTo 0
End Sub

' DV geral do codigo de barras: posicao 5, mod 11, pesos 2..9 ciclicos
' da direita para a esquerda sobre os outros 43 digitos.
' Boleto de ARRECADACAO / concessionaria (forma O): comeca com 8.
'   pos 3 = identificador de valor (6/7 -> mod 10 ; 8/9 -> mod 11)
'   pos 4 = DV geral sobre os outros 43 digitos
'   pos 5-15 = valor em centavos quando pos 3 = 6 ou 8
'   linha digitavel = 48 digitos (4 blocos de 11 + DV)
Function Bip_EhArrecadacao(b)
    Bip_EhArrecadacao = (Left(b, 1) = "8")
End Function

Function Bip_DVArrec(s, ident)
    Dim i, peso, soma, resto
    If ident = "6" Or ident = "7" Then
        Bip_DVArrec = Bip_DVMod10(s)
        Exit Function
    End If
    soma = 0 : peso = 2
    For i = Len(s) To 1 Step -1
        soma = soma + CInt(Mid(s, i, 1)) * peso
        peso = peso + 1
        If peso > 9 Then peso = 2
    Next
    resto = soma Mod 11
    If resto < 2 Then Bip_DVArrec = 0 Else Bip_DVArrec = 11 - resto
End Function

Function Bip_ValorDaBarra(b44)
    Dim ident
    If Bip_EhArrecadacao(b44) Then
        ident = Mid(b44, 3, 1)
        If ident = "6" Or ident = "8" Then Bip_ValorDaBarra = CDbl(Mid(b44, 5, 11)) Else Bip_ValorDaBarra = 0
    Else
        Bip_ValorDaBarra = CDbl(Mid(b44, 10, 10))
    End If
End Function

Function Bip_LinhaArrecParaBarra(d48, erro)
    Dim i, bloco, ident, barra
    Bip_LinhaArrecParaBarra = ""
    erro = ""
    ident = Mid(d48, 3, 1)
    barra = ""
    For i = 0 To 3
        bloco = Mid(d48, i * 12 + 1, 11)
        If CStr(Bip_DVArrec(bloco, ident)) <> Mid(d48, i * 12 + 12, 1) Then
            erro = "o digito do " & (i + 1) & "o bloco nao confere" : Exit Function
        End If
        barra = barra & bloco
    Next
    If Not Bip_DVBarraOK(barra) Then
        erro = "o digito verificador geral nao confere" : Exit Function
    End If
    Bip_LinhaArrecParaBarra = barra
End Function

Function Bip_DVBarraOK(b44)
    Dim b43, calc
    Bip_DVBarraOK = False
    If Len(b44) <> 44 Then Exit Function
    If Bip_EhArrecadacao(b44) Then
        If CStr(Bip_DVArrec(Left(b44, 3) & Mid(b44, 5), Mid(b44, 3, 1))) = Mid(b44, 4, 1) Then Bip_DVBarraOK = True
        Exit Function
    End If
    b43 = Left(b44, 4) & Mid(b44, 6)
    calc = Bip_DVMod11(b43)
    If CStr(calc) = Mid(b44, 5, 1) Then Bip_DVBarraOK = True
End Function

Function Bip_DVMod11(b43)
    Dim i, peso, soma, resto, d
    soma = 0 : peso = 2
    For i = Len(b43) To 1 Step -1
        d = CInt(Mid(b43, i, 1))
        soma = soma + d * peso
        peso = peso + 1
        If peso > 9 Then peso = 2
    Next
    resto = 11 - (soma Mod 11)
    If resto = 0 Or resto = 10 Or resto = 11 Then resto = 1
    Bip_DVMod11 = resto
End Function

Function Bip_DVMod10(bloco)
    Dim i, peso, soma, p
    soma = 0 : peso = 2
    For i = Len(bloco) To 1 Step -1
        p = CInt(Mid(bloco, i, 1)) * peso
        If p > 9 Then p = p - 9
        soma = soma + p
        If peso = 2 Then peso = 1 Else peso = 2
    Next
    Bip_DVMod10 = (10 - (soma Mod 10)) Mod 10
End Function

' Linha digitavel de 47 digitos -> codigo de barras de 44.
Function Bip_LinhaParaBarra(d47, erro)
    Dim c1, c2, c3, dv, fator, valor, livre, barra
    Bip_LinhaParaBarra = ""
    erro = ""
    If Len(d47) = 48 And Bip_EhArrecadacao(d47) Then
        Bip_LinhaParaBarra = Bip_LinhaArrecParaBarra(d47, erro)
        Exit Function
    End If
    If Len(d47) <> 47 Then
        erro = "a linha digitavel tem " & Len(d47) & " digitos e deveria ter 47"
        Exit Function
    End If

    c1 = Mid(d47, 1, 9)
    c2 = Mid(d47, 11, 10)
    c3 = Mid(d47, 22, 10)

    If CStr(Bip_DVMod10(c1)) <> Mid(d47, 10, 1) Then
        erro = "o digito do 1o campo nao confere" : Exit Function
    End If
    If CStr(Bip_DVMod10(c2)) <> Mid(d47, 21, 1) Then
        erro = "o digito do 2o campo nao confere" : Exit Function
    End If
    If CStr(Bip_DVMod10(c3)) <> Mid(d47, 32, 1) Then
        erro = "o digito do 3o campo nao confere" : Exit Function
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
    If Not Bip_DVBarraOK(barra) Then
        erro = "o digito verificador geral nao confere"
        Exit Function
    End If
    Bip_LinhaParaBarra = barra
End Function

Function Bip_SoDigitos(t)
    Dim i, c, s
    s = ""
    For i = 1 To Len(t)
        c = Mid(t, i, 1)
        If c >= "0" And c <= "9" Then s = s & c
    Next
    Bip_SoDigitos = s
End Function

' DATA DA FATURA vem como ddmmaaaa (18082026) ou dd.mm.aaaa - o ano sao
' os 4 ultimos digitos nos dois casos.
Function AnoDaDataBip(t)
    Dim d
    AnoDaDataBip = ""
    d = Bip_SoDigitos(t)
    If Len(d) < 8 Then Exit Function
    AnoDaDataBip = Right(d, 4)
End Function

Function Bip_Formatar(cent)
    Dim s, neg
    neg = ""
    If cent < 0 Then neg = "-" : cent = -cent
    s = CStr(cent)
    Do While Len(s) < 3
        s = "0" & s
    Loop
    Bip_Formatar = neg & Left(s, Len(s) - 2) & "," & Right(s, 2)
End Function

Function Bip_FormasAceitas()
    Bip_FormasAceitas = Trim(mapa("FORMAS_VALIDAS"))
    If Bip_FormasAceitas = "" Then Bip_FormasAceitas = "O|S|X"
End Function

Function Bip_FormaValida(f)
    Dim lista, k
    Bip_FormaValida = False
    lista = Split(UCase(Bip_FormasAceitas()), "|")
    For k = 0 To UBound(lista)
        If Trim(lista(k)) = f Then Bip_FormaValida = True
    Next
End Function

Function Bip_FormaSemBoleto(f)
    Dim lista, k, texto
    Bip_FormaSemBoleto = False
    texto = Trim(mapa("FORMAS_SEM_BOLETO"))
    If texto = "" Then Exit Function
    lista = Split(UCase(texto), "|")
    For k = 0 To UBound(lista)
        If Trim(lista(k)) = f Then Bip_FormaSemBoleto = True
    Next
End Function

Function Bip_EhChaveFornecedor(chave)
    Dim lista, k, texto, c
    Bip_EhChaveFornecedor = False
    c = Trim(chave)
    If c = "" Then Exit Function
    texto = Trim(mapa("CHAVES_FORNECEDOR"))
    If texto = "" Then texto = "31|21"
    lista = Split(texto, "|")
    For k = 0 To UBound(lista)
        If Trim(lista(k)) = c Then Bip_EhChaveFornecedor = True
    Next
End Function
