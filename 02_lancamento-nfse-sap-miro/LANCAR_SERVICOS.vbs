' =====================================================================
'  LANCAR_SERVICOS.vbs
'  Executor de lancamento de Notas Fiscais de Servico (NFS-e) na MIRO.
'  setembro 2026
'
'  Roda na maquina do usuario, anexado a sessao SAP JA AUTENTICADA.
'  Nao faz login, nao ve senha, nao escolhe usuario.
'
'  DOIS MODOS, um arquivo so:
'    LANCAR  (padrao) - percorre o lote e lanca
'    DESPEJO           - somente leitura: fotografa a tela SAP aberta
'                        num CAMPOS_<data>.txt e sai, sem abrir planilha
'                        nem tocar em nota. Pede-se pelo MODO no mapa ou
'                        por argumento: cscript LANCAR_SERVICOS.vbs DESPEJO
'
'  Le  : LANCAMENTO_NFCOM.xlsx (aba LOTE) - planilha UNIFICADA NFCOM+NFS-e
'  Grava: colunas W a AA da mesma linha
'  Config: MAPA_CAMPOS_SERVICOS.txt (todos os IDs de tela)
'
'  REGRAS DE OURO
'   - Linha com STATUS preenchido NUNCA e reprocessada (anti-duplicidade).
'   - Linha com LANCAR? diferente de SIM e ignorada, sem escrever nada.
'   - So grava com saldo em 0,00 exatos. Sem tolerancia: em servicos nao
'     existe rateio de ICMS a arredondar.
'   - Nao conserta, nao arredonda, nao completa dado nenhum.
' =====================================================================
Option Explicit

Const TITULO = "Lancamento NFS-e"

' --- colunas da aba LOTE ---
Const COL_LANCAR = 24 
Const COL_MOTIVO = 25 
Const COL_MES    = 26 
Const COL_EMPRESA= 27 
Const COL_COD    = 2  
Const COL_FORN   = 23 
Const COL_CNPJ   = 28 
Const COL_NF     = 5  
Const COL_EMIS   = 3  
Const COL_HORA   = 29 
Const COL_VALOR  = 4  
Const COL_PEDIDOS= 7  
Const COL_ITENS  = 30 
Const COL_RATEIO = 8  
Const COL_VERIF  = 31 
Const COL_CHAVE  = 32 
Const COL_AMBGER = 33 
Const COL_DOC9   = 34 
Const COL_TIPO   = 6  
Const COL_PGTO   = 35 
Const COL_RETENCAO= 36 
Const COL_INFO1  = 37 
Const COL_INFO2  = 38 
Const COL_INFO3  = 39 
Const COL_STATUS = 18 
Const COL_DOCMIRO= 19 
Const COL_ERRO   = 20 
Const COL_QUANDO = 21 
Const COL_EXEC   = 22 

Dim dicMapa, fso, pastaBase
Dim aplicacao, conexao, sessao
Dim appExcel, arqLote, abaLote
Dim ambiente, usuarioSap, usuarioWin
Dim linhaAtual, ultimaLinha
Dim qtdLancadas, qtdRetidas, qtdPuladas
Dim resumo
' Toda intervencao automatica do executor e registrada na planilha.
Dim obsGrade

Set fso = CreateObject("Scripting.FileSystemObject")
pastaBase = fso.GetParentFolderName(WScript.ScriptFullName)

' =====================================================================
'  1. Mapa de campos
' =====================================================================
Set dicMapa = CarregarMapa(pastaBase & "\MAPA_CAMPOS_SERVICOS.txt")
If dicMapa.Count = 0 Then
    MsgBox "Nao consegui ler o MAPA_CAMPOS_SERVICOS.txt." & vbCrLf & _
           "Ele precisa estar na mesma pasta deste script.", 16, TITULO
    WScript.Quit
End If

' Uma chave de configuracao que o executor nao respeita e pior do que
' chave nenhuma: quem a edita acha que mudou o comportamento. Este
' executor nao implementa correcao automatica de base em servicos, entao
' a tolerancia TEM de estar em zero.
If Len(Mapa("NF_TOLERANCIA_CENTAVOS")) > 0 Then
    If CLng(Mapa("NF_TOLERANCIA_CENTAVOS")) <> 0 Then
        MsgBox "NF_TOLERANCIA_CENTAVOS esta em " & Mapa("NF_TOLERANCIA_CENTAVOS") & _
               ", mas este executor nao corrige base de imposto em servicos." & vbCrLf & vbCrLf & _
               "Ele so lanca com saldo em 0,00 exatos. Volte a chave para 0, ou " & _
               "implemente a correcao antes de rodar.", 16, TITULO
        WScript.Quit
    End If
End If

' =====================================================================
'  2. Sessao SAP e trava de ambiente
' =====================================================================
On Error Resume Next
Set aplicacao = GetObject("SAPGUI").GetScriptingEngine
Set conexao = aplicacao.Children(0)
Set sessao = conexao.Children(0)
If Err.Number <> 0 Or sessao Is Nothing Then
    MsgBox "Nao encontrei uma sessao SAP aberta." & vbCrLf & _
           "Abra o SAP, faca login no ambiente certo e rode de novo.", 16, TITULO
    WScript.Quit
End If
Err.Clear
On Error GoTo 0

ambiente   = sessao.Info.SystemName & "/" & sessao.Info.Client
usuarioSap = sessao.Info.User
usuarioWin = NomeWindows()

' ---------------------------------------------------------------------
'  Modo DESPEJO: somente leitura, sai antes de qualquer trava de
'  ambiente - fotografar tela nao lanca nada e vale em qualquer sistema.
' ---------------------------------------------------------------------
If ModoDespejo() Then
    Dim arqDespejado
    arqDespejado = "CAMPOS_" & Carimbo() & ".txt"
    Despejar arqDespejado
    MsgBox "Despejo de tela gravado em:" & vbCrLf & vbCrLf & _
           pastaBase & "\" & arqDespejado & vbCrLf & vbCrLf & _
           "Somente leitura: nada foi digitado nem lancado." & vbCrLf & _
           "Sessao: " & ambiente & "   Usuario: " & usuarioSap, 64, TITULO
    WScript.Quit
End If

If Not AmbientePermitido(ambiente) Then
    MsgBox "A sessao SAP nao e um dos ambientes liberados." & vbCrLf & vbCrLf & _
           "Sessao aberta : " & ambiente & vbCrLf & _
           "Liberados     : " & Mapa("SAP_AMBIENTES_PERMITIDOS") & vbCrLf & vbCrLf & _
           "Feche as outras janelas do SAP e deixe so a certa.", 16, TITULO
    WScript.Quit
End If

' Com mais de um ambiente liberado, o lote nao sabe para qual foi
' montado: a unica coisa que separa teste de documento real e qual
' janela do SAP esta aberta. Por isso confirma sempre.
If InStr(Mapa("SAP_AMBIENTES_PERMITIDOS"), "|") > 0 And Not (Len(Trim(Mapa("SAP_USUARIO_ESPERADO"))) > 0 And UCase(Trim(usuarioSap)) = UCase(Trim(Mapa("SAP_USUARIO_ESPERADO"))) And ambiente = "PRD/100") Then
    If MsgBox("Prestes a lancar em:" & vbCrLf & vbCrLf & _
              "        >>>  " & ambiente & "  <<<" & vbCrLf & vbCrLf & _
              "  Usuario SAP: " & usuarioSap & vbCrLf & vbCrLf & _
              "Este processo esta liberado para mais de um ambiente, e o" & vbCrLf & _
              "lote nao sabe para qual dos dois foi montado. Confira:" & vbCrLf & _
              "se aqui estiver PRD, os documentos sao REAIS." & vbCrLf & vbCrLf & _
              "Continuar?", 4 + 48, TITULO) <> 6 Then
        WScript.Quit
    End If
End If

If Len(Mapa("SAP_USUARIO_ESPERADO")) > 0 Then
    If UCase(Trim(Mapa("SAP_USUARIO_ESPERADO"))) <> UCase(Trim(usuarioSap)) Then
        If MsgBox("A sessao esta logada em " & usuarioSap & ", e nao em " & _
                  Mapa("SAP_USUARIO_ESPERADO") & "." & vbCrLf & vbCrLf & _
                  "Todo documento sairá no nome de " & usuarioSap & "." & vbCrLf & vbCrLf & _
                  "Continuar?", 4 + 48, TITULO) <> 6 Then
            WScript.Quit
        End If
    End If
End If

' =====================================================================
'  3. Planilha de lote
' =====================================================================
Dim caminhoLote
caminhoLote = pastaBase & "\LANCAMENTO_NFCOM.xlsx"
If Not fso.FileExists(caminhoLote) Then
    MsgBox "Nao achei a LANCAMENTO_NFCOM.xlsx nesta pasta.", 16, TITULO
    WScript.Quit
End If

On Error Resume Next
Set appExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Nao consegui abrir o Excel.", 16, TITULO
    WScript.Quit
End If
Err.Clear
On Error GoTo 0

appExcel.Visible = False
appExcel.DisplayAlerts = False
Set arqLote = appExcel.Workbooks.Open(caminhoLote)

If arqLote.ReadOnly Then
    arqLote.Close False
    appExcel.Quit
    MsgBox "A planilha abriu somente para leitura - ela esta aberta no Excel." & vbCrLf & vbCrLf & _
           "Feche o Excel e rode de novo. Se houver um arquivo" & vbCrLf & _
           "~$LANCAMENTO_NFCOM.xlsx nesta pasta, ela esta aberta em algum lugar.", 16, TITULO
    WScript.Quit
End If

Set abaLote = arqLote.Worksheets("LOTE")
ultimaLinha = abaLote.Cells(abaLote.Rows.Count, COL_NF).End(-4162).Row   ' xlUp

qtdLancadas = 0
qtdRetidas = 0
qtdPuladas = 0
resumo = ""

' =====================================================================
'  4. Laco principal
' =====================================================================
For linhaAtual = 2 To ultimaLinha
    Dim numeroNf, decisao, statusAtual
    numeroNf    = Celula(linhaAtual, COL_NF)
    decisao     = UCase(Trim(Celula(linhaAtual, COL_LANCAR)))
    statusAtual = Trim(Celula(linhaAtual, COL_STATUS))

    If Len(numeroNf) = 0 Then
        ' linha em branco: ignora
    ElseIf EhLinhaNFCOM(Celula(linhaAtual, COL_TIPO)) Then
        ' A planilha e compartilhada com o LANCAR_NFCOM.vbs. Linha de
        ' NFCOM nao e erro aqui: e so nao ser deste script. Passa reto
        ' SEM escrever nada - inclusive sem Reter - para nao sujar a
        ' linha do outro executor.
        qtdPuladas = qtdPuladas + 1
    ElseIf Len(statusAtual) > 0 Then
        ' JA PROCESSADA. Esta e a trava anti-duplicidade: nunca reprocessa.
        qtdPuladas = qtdPuladas + 1
    ElseIf decisao <> "SIM" And decisao <> "SIM " Then
        ' retida pelo Claude ou pelo usuario. Nao escreve nada: a linha
        ' fica disponivel para virar SIM depois.
        qtdPuladas = qtdPuladas + 1
    Else
        ProcessarNota linhaAtual
        arqLote.Save
    End If
Next

arqLote.Save
arqLote.Close True
appExcel.Quit
Set appExcel = Nothing

MsgBox "Lote encerrado em " & ambiente & "." & vbCrLf & vbCrLf & _
       "Lancadas      : " & qtdLancadas & vbCrLf & _
       "Nao lancadas  : " & qtdRetidas & vbCrLf & _
       "Ignoradas     : " & qtdPuladas & vbCrLf & vbCrLf & _
       resumo, 64, TITULO


' =====================================================================
'  Uma nota
' =====================================================================
Sub ProcessarNota(lin)
    Dim numeroNf, empresaCod, dataEmis, horaEmis, valorNf
    Dim listaPedidos, listaItens, listaRateio
    Dim codVerif, chave, ambGer, doc9, digitoVerif, tipoNf, ehZH, temChave, referencia, retencao
    Dim mensagens, erroMsg
    Dim saldo, numeroDoc, aviso

    numeroNf   = Trim(Celula(lin, COL_NF))
    empresaCod = SoDigitos(Celula(lin, COL_COD))
    dataEmis   = SoDigitos(Celula(lin, COL_EMIS))
    horaEmis   = Trim(Celula(lin, COL_HORA))
    valorNf    = Moeda(Celula(lin, COL_VALOR))
    codVerif   = Trim(Celula(lin, COL_VERIF))
    chave      = SoDigitos(Celula(lin, COL_CHAVE))
    ambGer     = Trim(Celula(lin, COL_AMBGER))
    doc9       = Trim(Celula(lin, COL_DOC9))
    tipoNf     = UCase(Trim(Celula(lin, COL_TIPO)))
    retencao   = UCase(Trim(Celula(lin, COL_RETENCAO)))
    aviso      = ""
    obsGrade   = ""

    listaPedidos = Partir(Celula(lin, COL_PEDIDOS))
    listaItens   = Partir(Celula(lin, COL_ITENS))
    listaRateio  = Partir(Celula(lin, COL_RATEIO))

    ' ---- que tipo de linha e esta? decide o caminho inteiro ----
    If Not TipoAceito(tipoNf) Then
        Reter lin, "TIPO NF na planilha esta [" & tipoNf & "], que nao e um dos valores aceitos (" & _
                   Mapa("TIPOS_ACEITOS") & "). Vazio tambem retem: na planilha vazio significa " & _
                   "'ainda nao preenchi', e recibo precisa ser dito com todas as letras."
        Exit Sub
    End If
    ehZH = (tipoNf = "ZH")

    ' ---- conferencias antes de tocar no SAP ----
    If Len(empresaCod) = 0 Then
        Reter lin, "Empresa (codigo SAP) esta vazia na planilha."
        Exit Sub
    End If
    If Len(dataEmis) <> 8 Then
        Reter lin, "Data de emissao invalida: [" & Celula(lin, COL_EMIS) & "]."
        Exit Sub
    End If
    If Len(valorNf) = 0 Then
        Reter lin, "Valor da NF esta vazio na planilha."
        Exit Sub
    End If
    If UBound(listaPedidos) < 0 Then
        Reter lin, "Nenhum pedido informado na planilha."
        Exit Sub
    End If
    If UBound(listaPedidos) + 1 > CLng(Mapa("MAX_PEDIDOS_POR_NOTA")) Then
        Reter lin, "Nota com " & (UBound(listaPedidos) + 1) & " pedidos, acima da guarda de seguranca."
        Exit Sub
    End If
    If UBound(listaPedidos) > 0 Then
        If UBound(listaRateio) <> UBound(listaPedidos) Then
            Reter lin, "Nota com " & (UBound(listaPedidos) + 1) & " pedidos e " & _
                       (UBound(listaRateio) + 1) & " valores de rateio. Nao se deduz rateio."
            Exit Sub
        End If
    End If
    ' Chave, DOCNUM9, ambiente gerador e digito verificador so existem
    ' em NFS-e. Recibo nao tem nenhum deles e nao passa por estas travas.
    digitoVerif = ""
    temChave = False
    If ehZH Then
        ' A chave e tudo ou nada. Vazia = modelo municipal sem
        ' identificador nacional (ex.: municipio de exemplo): o caminho para no codigo
        ' de verificacao. Qualquer coisa entre vazio e 50 digitos e chave
        ' pela metade, e chave pela metade nao se completa.
        temChave = (Len(chave) = 50)
        If Len(chave) > 0 And Not temChave Then
            Reter lin, "Identificador Nacional com " & Len(chave) & " digitos: nao e vazio " & _
                       "nem tem os 50. Chave pela metade nao se completa."
            Exit Sub
        End If
        If Not temChave And UCase(Mapa("ZH_SEM_CHAVE_PERMITIDO")) <> "SIM" Then
            Reter lin, "Nota ZH sem Identificador Nacional, e o mapa nao libera esse caminho."
            Exit Sub
        End If
    End If
    If temChave Then
        If Len(doc9) <> 9 Then
            Reter lin, "DOCNUM9 com " & Len(doc9) & " digitos, e nao 9."
            Exit Sub
        End If
        If Len(ambGer) = 0 Then
            Reter lin, "Ambiente gerador vazio (sai da posicao 8 da chave)."
            Exit Sub
        End If

    ' As colunas AMBIENTE GERADOR e DOCNUM9 sao recortes da propria chave.
    ' Se divergirem, alguem editou uma das duas: retem em vez de escolher.
        If ambGer <> Mid(chave, 8, 1) Then
            Reter lin, "AMBIENTE GERADOR na planilha e [" & ambGer & "] mas a posicao 8 da chave " & _
                       "traz [" & Mid(chave, 8, 1) & "]. Nao escolho entre os dois."
            Exit Sub
        End If
        If doc9 <> Mid(chave, 41, 9) Then
            Reter lin, "DOCNUM9 na planilha e [" & doc9 & "] mas as posicoes 41-49 da chave trazem [" & _
                       Mid(chave, 41, 9) & "]. Nao escolho entre os dois."
            Exit Sub
        End If
        digitoVerif = Mid(chave, 50, 1)
    End If

    ' ---- retencao de imposto ----
    ' A pergunta certa e "esta nota tem imposto retido?", nao "como ela e
    ' paga". A versao anterior retinha por TIPO DE PAGAMENTO = BOLETO, o
    ' que reteve indevidamente o boleto de condominio, uma
    ' nota sem retencao nenhuma. Forma de pagamento so passa a importar
    ' quando os campos da aba Imp.ret.fonte estiverem mapeados.
    If retencao <> "SIM" And retencao <> "NAO" And retencao <> "NÃO" Then
        Reter lin, "Coluna RETENCAO esta [" & Celula(lin, COL_RETENCAO) & "]. " & _
                   "Precisa ser SIM ou NAO - ausencia de retencao nao se presume."
        Exit Sub
    End If
    If retencao = "SIM" Then
        If Len(Mapa("WT_TABELA")) = 0 Then
            Reter lin, "Nota com imposto retido, e os campos da aba Imp.ret.fonte ainda nao estao " & _
                       "mapeados: o executor nao sabe preencher a retencao nem conferir o IR. " & _
                       "Lancar esta nota manualmente."
            Exit Sub
        End If
    End If

    ' =============== a partir daqui, SAP ===============
    If Not AbrirMiroNaEmpresa(empresaCod) Then
        Reter lin, "Nao consegui fixar a empresa " & empresaCod & " na MIRO. " & _
                   "Titulo da janela: [" & Ler("wnd[0]") & "]."
        Exit Sub
    End If

    Escrever Mapa("MIRO_DATA_FATURA"), dataEmis
    referencia = ReferenciaDaLinha(tipoNf, numeroNf, listaPedidos)
    If Len(referencia) = 0 Then
        Reter lin, "Nao consegui montar a Referencia para o tipo " & tipoNf & _
                   ". Em recibo sem documento ela sai dos " & Mapa("REF_SEM_DOC_DIGITOS") & _
                   " ultimos digitos do pedido, e a nota tem de ter um pedido so."
        Exit Sub
    End If
    Escrever Mapa("MIRO_REFERENCIA"), referencia
    Escrever Mapa("MIRO_MONTANTE"), valorNf

    ' Local de negocio (BUPLA). No despejo de 02/09 ele estava VAZIO e o
    ' SAP recusou com "Local de negocio XXXX nao encontrado na estrutura
    ' do empreendimento". Enquanto MIRO_LOCAL_NEGOCIO_VALOR estiver vazio
    ' no mapa, o executor NAO escreve nada aqui e deixa o SAP derivar -
    ' preencher um local de negocio errado seria pior. Se o SAP recusar,
    ' a mensagem dele vai inteira para a coluna ERRO.
    If Len(Mapa("MIRO_LOCAL_NEGOCIO_VALOR")) > 0 Then
        Escrever Mapa("MIRO_LOCAL_NEGOCIO"), Mapa("MIRO_LOCAL_NEGOCIO_VALOR")
    End If

    ' ---- pedidos ----
    ' Corrigido em 05/10/2026: o LANCAR_NFCOM.vbs NUNCA usa atalho de pedido
    ' unico - sempre abre a janela de selecao multipla (botao XMSEL +
    ' Transferir), mesmo com um pedido so. Aqui o atalho (escrever direto em
    ' MIRO_PEDIDO_UNICO + Enter) deixou o campo do pedido preenchido mas a
    ' MIRO nao adotou o item: a grade voltou vazia nas duas leituras (Fornecedor X
    ' NF 0001234, pedido 4500000003, despejo GRADE_VAZIA_0001234.txt). A
    ' janela de selecao multipla e quem de fato aciona a MIRO a puxar o
    ' item (e so depois disso a quantidade e tratada, em PreencherGrade) -
    ' por isso agora se usa sempre ela, igual ao NFCOM, independente de
    ' quantos pedidos a nota tiver.
    If Not InformarVariosPedidos(listaPedidos) Then
        Despejar "POPUP_PEDIDOS_" & numeroNf & ".txt"
        Reter lin, "A janela de selecao multipla de pedidos nao fechou. SAP disse: [" & popupMsg & "]. " & _
                   "Veja o POPUP_PEDIDOS_" & numeroNf & ".txt na pasta."
        Exit Sub
    End If

    ' Corrigido em 05/10/2026, 3a tentativa: entre o Adotar e o
    ' preenchimento de quantidade/valor, a unica coisa que o script fazia
    ' na grade era o PrepararGrade (redimensiona a coluna de valor e
    ' aperta o botao REFRESH da tabela de itens) - e isso rodava
    ' incondicionalmente, ANTES de qualquer leitura, mesmo quando o item
    ' ja tinha chegado certinho pela janela de selecao multipla. Na Fornecedor X
    ' NF 0001234 (pedido 4500000003) esse refresh forcado derrubava o
    ' item recem-adotado, voltando a grade vazia. Repetir a MAO (sem
    ' forcar nenhum refresh) funcionou na hora. Agora o script le a
    ' grade do jeito que ela ja estiver PRIMEIRO; o PrepararGrade so
    ' entra como recurso, se essa primeira leitura vier vazia - e ele
    ' so existe por causa de um bug antigo e diferente (nota de um
    ' pedido lida como tendo dois itens).
    Dim gradePed, itensNaGrade, erroGrade
    gradePed = LerPedidosDaGrade()
    itensNaGrade = UBound(gradePed) + 1
    If itensNaGrade = 0 Then
        PrepararGrade
        gradePed = LerPedidosDaGrade()
        itensNaGrade = UBound(gradePed) + 1
    End If
    If itensNaGrade = 0 Then
        Despejar "GRADE_VAZIA_" & numeroNf & ".txt"
        Reter lin, "A grade de itens voltou vazia em todas as leituras: a MIRO nao adotou os " & _
                   "pedidos. Mande o GRADE_VAZIA_" & numeroNf & ".txt que apareceu na pasta."
        Exit Sub
    End If
    If itensNaGrade <> UBound(listaPedidos) + 1 Then
        Despejar "GRADE_DIVERGENTE_" & numeroNf & ".txt"
        Reter lin, "A grade trouxe " & itensNaGrade & " item(ns) para " & _
                   (UBound(listaPedidos) + 1) & " pedido(s) da planilha. Um pedido em mais de " & _
                   "uma linha retem: o rateio dentro do proprio pedido nao esta na planilha " & _
                   "e nao se deduz. [" & geoGrade & "] Veja o GRADE_DIVERGENTE_" & numeroNf & ".txt."
        Exit Sub
    End If

    ' ---- quantidade e valor liquido, linha por linha ----
    ' Vale para nota de um pedido tambem: quando a MIRO nao puxa o
    ' montante do pedido/item, e este passo que faz o saldo fechar.
    erroGrade = PreencherGrade(listaPedidos, listaRateio, gradePed)
    If Len(erroGrade) > 0 Then
        Reter lin, erroGrade
        Exit Sub
    End If
    Tecla 0

    ' ---- categoria da NF (aba Detalhe) - SOMENTE em NFS-e ----
    ' Recibo nao passa por aqui: e a AUSENCIA de categoria que faz do
    ' documento um recibo. Escrever zh numa linha de recibo mudaria a
    ' natureza fiscal do documento.
    If ehZH Then
        Selecionar Mapa("MIRO_ABA_DETALHE")
        Escrever Mapa("MIRO_TIPO_NF"), Mapa("MIRO_TIPO_NF_VALOR")
        Foco Mapa("MIRO_TIPO_NF")
        Tecla 0

        ' releitura como prova: apertar nao e prova de que ficou
        If UCase(Trim(Ler(Mapa("MIRO_TIPO_NF")))) <> UCase(Mapa("MIRO_TIPO_NF_VALOR")) Then
            Reter lin, "A categoria NF nao ficou: escrevi [" & Mapa("MIRO_TIPO_NF_VALOR") & _
                       "] e a tela devolveu [" & Trim(Ler(Mapa("MIRO_TIPO_NF"))) & "]."
            Exit Sub
        End If
    End If

    ' ---- tela Nota Fiscal - SOMENTE em NFS-e ----
    ' As tres gravacoes de recibo nao abrem esta tela em momento algum:
    ' recibo nao tem chave, nem codigo de verificacao, nem protocolo.
    If ehZH Then
        ' ---- tela Nota Fiscal ----
        Selecionar Mapa("MIRO_ABA_DADOS_BASICOS")
        Tecla CLng(Mapa("MIRO_TECLA_NOTA_FISCAL"))

        If Not Existe(Mapa("NF_ABA")) Then
            Despejar "NF_NAO_ABRIU_" & numeroNf & ".txt"
            Reter lin, "A tela Nota Fiscal nao abriu. Veja o NF_NAO_ABRIU_" & numeroNf & ".txt."
            Exit Sub
        End If

        ' Outras informacoes: uma linha de tela por linha de texto.
        ' O texto vem das tres colunas e de quebras de linha dentro delas.
        mensagens = MontarMensagens(lin)
        If UBound(mensagens) >= 0 Then
            Selecionar Mapa("NF_ABA_MENSAGENS")
            erroMsg = EscreverMensagens(mensagens)
            If Len(erroMsg) > 0 Then
                Reter lin, erroMsg
                Exit Sub
            End If
        End If

        ' identificacao da NFS-e
        Selecionar Mapa("NF_ABA")
        Escrever Mapa("NF_NUM_NF"), numeroNf
        ' Codigo de verificacao.
        ' COM chave (padrao nacional): copia literal, com hifen e tudo -
        ' foi assim que as gravacoes 1 e 3 escreveram (ABCD-1234,
        ' EFGH-5678) e assim que os documentos ja lancados sairam.
        ' SEM chave (modelo municipal, ex.: municipio de exemplo): saem os separadores
        ' e ficam letras e numeros, por decisao do dono do processo em
        ' 03/09/2026: AB12.3456.7890.1234567-X -> AB12345678901234567X.
        If Len(codVerif) > 0 Then
            Dim codParaTela
            If temChave Then
                codParaTela = Trim(codVerif)
            Else
                codParaTela = SemSeparadores(codVerif)
            End If
            Escrever Mapa("NF_COD_VERIFICACAO"), codParaTela
            If Trim(Ler(Mapa("NF_COD_VERIFICACAO"))) <> codParaTela Then
                Reter lin, "O codigo de verificacao nao ficou: escrevi [" & codParaTela & _
                           "] e a tela devolveu [" & Trim(Ler(Mapa("NF_COD_VERIFICACAO"))) & "]."
                Exit Sub
            End If
        End If
        Escrever Mapa("NF_DATA_PROCMTO"), dataEmis
        If Len(horaEmis) > 0 Then
            Escrever Mapa("NF_TEMPO_PROCMTO"), horaEmis
        Else
            Escrever Mapa("NF_TEMPO_PROCMTO"), "00:00:00"
        End If
        ' Daqui para baixo so existe em nota COM identificador nacional.
        ' No modelo municipal sem chave (ex.: municipio de exemplo) a sequencia termina
        ' na data e hora acima: volta para a MIRO e lanca.
        If temChave Then
            Escrever Mapa("NF_AMBIENTE_GERADOR"), ambGer
            Escrever Mapa("NF_DOCNUM9"), doc9
            Foco Mapa("NF_DOCNUM9")
            Tecla 0

            ' releitura como prova, nos dois campos que mais custaram no NFCOM
            If SoDigitos(Ler(Mapa("NF_DOCNUM9"))) <> doc9 Then
                Reter lin, "O DOCNUM9 nao ficou: escrevi [" & doc9 & "] e a tela devolveu [" & _
                           Trim(Ler(Mapa("NF_DOCNUM9"))) & "]."
                Exit Sub
            End If

            ' Digito verificador: campo proprio, DIGITADO depois do Enter no
            ' DOCNUM9. O SAP nao calcula (gravacao 3). Sem este campo mapeado a
            ' nota retem - chave incompleta nao se completa por conta propria.
            If Len(Mapa("NF_DIGITO_VERIFICADOR")) = 0 Then
                Reter lin, "O campo do digito verificador da chave nao esta mapeado."
                Exit Sub
            End If
            Escrever Mapa("NF_DIGITO_VERIFICADOR"), digitoVerif
            Foco Mapa("NF_DIGITO_VERIFICADOR")
            Tecla 0

            If SoDigitos(Ler(Mapa("NF_DIGITO_VERIFICADOR"))) <> digitoVerif Then
                Reter lin, "O digito verificador nao ficou: escrevi [" & digitoVerif & "] e a tela " & _
                           "devolveu [" & Trim(Ler(Mapa("NF_DIGITO_VERIFICADOR"))) & "]."
                Exit Sub
            End If
        End If

        Pressionar Mapa("NF_BOTAO_VOLTAR")
    End If

    ' ---- saldo: a trava que segura tudo ----
    If Not Existe(Mapa("MIRO_SALDO")) Then
        Despejar "SALDO_NAO_LIDO_" & numeroNf & ".txt"
        Reter lin, "Nao consegui LER o campo de saldo da MIRO. " & _
                   "Sem poder provar que esta zerado, a nota nao entra."
        Exit Sub
    End If
    saldo = Trim(Ler(Mapa("MIRO_SALDO")))
    If saldo <> "0,00" Then
        Reter lin, "Saldo da MIRO em " & saldo & " - so lanco com 0,00 exatos. " & _
                   "Em servicos a tolerancia e zero: confira o valor da NF contra o bruto do pedido."
        Exit Sub
    End If

    ' ---- simular e lancar ----
    If ehZH Then
        Pressionar Mapa("MIRO_BOTAO_SIMULAR")
    Else
        Tecla CLng(Mapa("MIRO_TECLA_SIMULAR"))
    End If
    If ContarJanelas() < 2 Then
        Despejar "SIMULACAO_NAO_ABRIU_" & numeroNf & ".txt"
        Reter lin, "A janela de simulacao nao abriu. Nada foi gravado. " & _
                   "Veja o SIMULACAO_NAO_ABRIU_" & numeroNf & ".txt."
        Exit Sub
    End If
    Pressionar Mapa("MIRO_BOTAO_GRAVAR")

    ' As gravacoes de recibo pressionam o Lancar duas vezes: depois da
    ' simulacao vem mais uma janela de confirmacao. Em vez de apertar as
    ' cegas, confere se sobrou janela aberta e so entao pressiona de novo.
    If ContarJanelas() >= 2 Then Pressionar Mapa("MIRO_BOTAO_GRAVAR")
    If ContarJanelas() >= 2 Then
        Despejar "JANELA_NAO_FECHOU_" & numeroNf & ".txt"
        Reter lin, "Depois de duas confirmacoes ainda sobrou janela aberta. " & _
                   "CONFIRA NO SAP antes de tentar de novo - pode ter gravado. " & _
                   "Veja o JANELA_NAO_FECHOU_" & numeroNf & ".txt."
        Exit Sub
    End If

    ' ---- numero do documento ----
    numeroDoc = NumeroNaBarra()
    If Len(numeroDoc) = 0 Then
        Despejar "SEM_NUMERO_DOC_" & numeroNf & ".txt"
        Reter lin, "Acionei o Lancar e a barra de status nao trouxe numero de documento. " & _
                   "Mensagem da tela: [" & Trim(Ler(Mapa("MIRO_BARRA_STATUS"))) & "]. " & _
                   "CONFIRA NO SAP antes de tentar de novo - pode ter gravado."
        Exit Sub
    End If

    If Len(obsGrade) > 0 Then aviso = "OBSERVACAO: " & obsGrade
    Gravar lin, "LANCADA", numeroDoc, aviso
    qtdLancadas = qtdLancadas + 1
    resumo = resumo & "NF " & numeroNf & " -> doc " & numeroDoc & vbCrLf
End Sub


'  O TIPO NF da linha decide o caminho. Vazio NAO vira recibo: na
'  planilha vazio significa "ainda nao preenchi".
' ---------------------------------------------------------------------
' A planilha LOTE e compartilhada com o LANCAR_NFCOM.vbs. A coluna
' TIPO NF separa os dois. Os tipos do OUTRO script vem do
' MAPA_CAMPOS_SERVICOS.txt em TIPOS_NFCOM; sem a chave, vale TA.
' Atencao: TIPO vazio NAO cai aqui - vazio continua retendo dentro de
' ProcessarNota, porque numa linha de servico vazio quer dizer
' "ainda nao preenchi", e isso tem de aparecer para o usuario.
' ---------------------------------------------------------------------
Function EhLinhaNFCOM(tipo)
    Dim lista, k, t, aceitos
    EhLinhaNFCOM = False
    t = UCase(Trim(tipo))
    If t = "" Then Exit Function
    aceitos = Mapa("TIPOS_NFCOM")
    If Trim(aceitos) = "" Then aceitos = "TA"
    lista = Split(UCase(aceitos), "|")
    For k = 0 To UBound(lista)
        If Trim(lista(k)) = t Then EhLinhaNFCOM = True
    Next
End Function

Function TipoAceito(tipo)
    Dim lista, k
    TipoAceito = False
    If Len(tipo) = 0 Then Exit Function
    lista = Split(UCase(Mapa("TIPOS_ACEITOS")), "|")
    For k = 0 To UBound(lista)
        If Trim(lista(k)) = tipo Then TipoAceito = True
    Next
End Function

'  Referencia (XBLNR):
'    ZH             -> numero da NF, copia literal
'    RECIBO FATURA  -> numero da fatura, na coluna NUMERO DA NF
'    RECIBO BOLETO  -> numero do boleto, na coluna NUMERO DA NF
'    RECIBO SEM DOC -> os N ultimos digitos do pedido (gravacao
'                      ReciboSemDoc: 4500001234 -> 1234)
Function ReferenciaDaLinha(tipo, numeroNf, pedidos)
    Dim n, quantos
    ReferenciaDaLinha = ""

    If tipo = "RECIBO SEM DOC" Then
        ' Sai do PEDIDO. Com mais de um pedido nao se decide de qual:
        ' retem em vez de escolher.
        If UBound(pedidos) <> 0 Then Exit Function
        n = SoDigitos(pedidos(0))
        quantos = 4
        If Len(Mapa("REF_SEM_DOC_DIGITOS")) > 0 Then quantos = CLng(Mapa("REF_SEM_DOC_DIGITOS"))

    ElseIf tipo = "RECIBO BOLETO" Then
        ' Sai do NOSSO NUMERO do banco, so os algarismos, contando o
        ' digito verificador: 000/00012345-6 -> 3456.
        n = SoDigitos(numeroNf)
        quantos = 4
        If Len(Mapa("REF_BOLETO_DIGITOS")) > 0 Then quantos = CLng(Mapa("REF_BOLETO_DIGITOS"))

    Else
        ' ZH e RECIBO FATURA: copia literal do numero do documento.
        ReferenciaDaLinha = Trim(numeroNf)
        Exit Function
    End If

    If Len(n) = 0 Then Exit Function
    ' Menos algarismos do que o recorte pede e sinal de dado incompleto:
    ' retem em vez de mandar uma referencia curta.
    If Len(n) < quantos Then Exit Function
    ReferenciaDaLinha = Right(n, quantos)
End Function


' =====================================================================
'  MIRO: abrir na empresa certa
' =====================================================================
Function AbrirMiroNaEmpresa(codEmpresa)
    AbrirMiroNaEmpresa = False

    Escrever "wnd[0]/tbar[0]/okcd", "/nmiro"
    Tecla 0

    ' Caminho 1: o popup que aparece SOZINHO na primeira MIRO da sessao.
    ' Gesto da gravacao 1: btn[0] da propria janela.
    If PopupEmpresaAberto() Then
        Escrever Mapa("MIRO_POPUP_EMPRESA"), codEmpresa
        Foco Mapa("MIRO_POPUP_EMPRESA")
        Pressionar Mapa("MIRO_POPUP_EMPRESA_OK")
    End If

    If EmpresaNaTela(codEmpresa) Then
        AbrirMiroNaEmpresa = True
        Exit Function
    End If

    ' Caminho 2: trocar a empresa pelo menu Processar > Mudar empresa.
    ' Gesto da gravacao 2 (Script2.vbs): Enter NA wnd[1], nao btn[0].
    If Not MudarEmpresaPeloMenu(codEmpresa) Then Exit Function

    AbrirMiroNaEmpresa = EmpresaNaTela(codEmpresa)
End Function

Function MudarEmpresaPeloMenu(codEmpresa)
    MudarEmpresaPeloMenu = False
    If Len(Mapa("MIRO_MENU_MUDAR_EMPRESA")) = 0 Then Exit Function

    Selecionar Mapa("MIRO_MENU_MUDAR_EMPRESA")
    If Not PopupEmpresaAberto() Then Exit Function

    Escrever Mapa("MIRO_POPUP_EMPRESA"), codEmpresa
    Foco Mapa("MIRO_POPUP_EMPRESA")
    TeclaEm Mapa("MIRO_POPUP_EMPRESA_TECLA_JANELA"), CLng(Mapa("MIRO_POPUP_EMPRESA_TECLA"))

    ' modal aberta engole todo comando que vem depois: confere que fechou
    If PopupEmpresaAberto() Then Exit Function

    MudarEmpresaPeloMenu = True
End Function

Function PopupEmpresaAberto()
    PopupEmpresaAberto = False
    If ContarJanelas() < 2 Then Exit Function
    If Not Existe(Mapa("MIRO_POPUP_EMPRESA")) Then Exit Function
    PopupEmpresaAberto = True
End Function

' Releitura como prova: o titulo da MIRO traz a empresa corrente
' ("Inserir fatura recebida: empresa 1000"). Sem essa conferencia, a
' segunda nota do lote herdaria a empresa da primeira.
Function EmpresaNaTela(codEmpresa)
    Dim titulo
    EmpresaNaTela = False
    If PopupEmpresaAberto() Then Exit Function
    titulo = Ler("wnd[0]")
    If Len(titulo) = 0 Then Exit Function
    If InStr(titulo, codEmpresa) = 0 Then Exit Function
    EmpresaNaTela = True
End Function


' =====================================================================
'  Selecao multipla de pedidos
' =====================================================================
Function InformarVariosPedidos(pedidos)
    Dim k
    InformarVariosPedidos = False
    popupMsg = ""

    Pressionar Mapa("MIRO_BOTAO_MULTIPLOS_PEDIDOS")
    If ContarJanelas() < 2 Then Exit Function

    ' escrita pelo objeto da grade (igual ao LANCAR_NFCOM), nao por ID fixo
    Dim colSel, celP, lidoP, confP
    colSel = IdxColuna(Mapa("MIRO_PEDIDO_LINHA"))
    lidoP = ""
    For k = 0 To UBound(pedidos)
        Set celP = CelGrade(IdTabela(Mapa("MIRO_PEDIDO_LINHA")), colSel, k)
        If celP Is Nothing Then
            popupMsg = "nao cheguei na linha " & k & " da grade [" & geoGrade & "]"
            Exit Function
        End If
        On Error Resume Next
        celP.Text = Trim(pedidos(k))
        If Err.Number <> 0 Then
            popupMsg = "erro ao escrever pedido " & pedidos(k) & " na linha " & k & ": " & Err.Description
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next

    ' releitura como prova: a janela guardou o que foi informado?
    For k = 0 To UBound(pedidos)
        Set celP = CelGrade(IdTabela(Mapa("MIRO_PEDIDO_LINHA")), colSel, k)
        confP = ""
        If Not (celP Is Nothing) Then
            On Error Resume Next
            confP = Trim(celP.Text)
            Err.Clear
            On Error GoTo 0
        End If
        lidoP = lidoP & confP & "|"
        If SoDigitos(confP) <> SoDigitos(pedidos(k)) Then
            popupMsg = "janela nao guardou o pedido: linha " & k & " esperava " & pedidos(k) & " e esta [" & confP & "]. Lidos: " & lidoP
            Exit Function
        End If
    Next
    Set celP = CelGrade(IdTabela(Mapa("MIRO_PEDIDO_LINHA")), colSel, 0)

    ' VAZIO DE PROPOSITO no mapa: o "Verificar" impedia o "Adotar" de
    ' transferir item nenhum. So age se alguem preencher a chave.
    If Len(Mapa("MIRO_BOTAO_CONTINUAR_PEDIDOS")) > 0 Then
        Pressionar Mapa("MIRO_BOTAO_CONTINUAR_PEDIDOS")
    End If

    Pressionar Mapa("MIRO_BOTAO_ADOTAR_PEDIDOS")
    If ContarJanelas() >= 2 Then
        ' popup nao fechou: o SAP quase sempre diz o porque na barra de
        ' status. Aperta "Verificar" so para OUVIR a mensagem (igual ao
        ' LANCAR_NFCOM) e guarda o texto para ir na coluna ERRO.
        popupMsg = BarraDoPopup()
        Pressionar "wnd[1]/tbar[0]/btn[0]"
        If Len(Trim(BarraDoPopup())) > 0 Then popupMsg = popupMsg & " | apos Verificar: " & BarraDoPopup()
        Exit Function
    End If

    InformarVariosPedidos = True
End Function

Dim popupMsg

' Texto da barra de status: da janela de cima, ou da principal.
Function BarraDoPopup()
    Dim t
    BarraDoPopup = ""
    On Error Resume Next
    t = ""
    t = sessao.findById("wnd[1]/sbar").Text
    Err.Clear
    If Len(Trim(t)) = 0 Then
        t = sessao.findById("wnd[0]/sbar").Text
        Err.Clear
    End If
    BarraDoPopup = Trim(t)
    On Error GoTo 0
End Function


' =====================================================================
'  Grade de itens
'  Baseado no LANCAR_NFCOM.vbs, que ja rodou este trecho em producao.
' =====================================================================

' Geometria da ultima tentativa de alcancar uma celula. Vai junto nas
' mensagens de erro: sem ela, "nao consegui chegar na linha 8" nao diz
' se faltou rolar, se o SAP recusou rolar, ou se a tabela acabou.
Dim geoGrade
geoGrade = ""

' O ID mapeado ja traz tabela e coluna: "..../tblX/txtDRSEG-EBELN[7,{L}]".
Function IdTabela(idCampo)
    Dim corte
    IdTabela = ""
    corte = InStrRev(idCampo, "/")
    If corte < 2 Then Exit Function
    IdTabela = Left(idCampo, corte - 1)
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

' Alcanca QUALQUER linha da grade, rolando se preciso.
' Regras que custaram caro no NFCOM:
'   - toda rolagem e ida ao servidor: buscar a tabela pelo ID DE NOVO;
'   - para tras, salto direto e aceito; para frente, uma pagina por vez;
'   - GetCell usa indice RELATIVO a posicao da barra, nao absoluto.
Function CelGrade(idTbl, colIdx, linhaAbs)
    Dim tbl, pos, vis, novo, voltas, maxVoltas
    Set CelGrade = Nothing
    geoGrade = ""
    If Len(idTbl) = 0 Or colIdx < 0 Then Exit Function

    maxVoltas = 60
    If Len(Mapa("MAX_ROLAGENS")) > 0 Then maxVoltas = CLng(Mapa("MAX_ROLAGENS"))

    On Error Resume Next
    Set tbl = sessao.findById(idTbl)
    If Err.Number <> 0 Or tbl Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    pos = -1
    pos = tbl.verticalScrollbar.Position
    If Err.Number <> 0 Then
        ' tabela sem barra de rolagem: indice absoluto serve
        Err.Clear
        geoGrade = "tabela sem rolagem"
        Set CelGrade = tbl.GetCell(linhaAbs, colIdx)
        If Err.Number <> 0 Then
            Err.Clear
            Set CelGrade = Nothing
        End If
        On Error GoTo 0
        Exit Function
    End If

    vis = 0
    vis = tbl.VisibleRowCount
    If Err.Number <> 0 Or vis < 1 Then
        Err.Clear
        vis = 1
    End If
    geoGrade = "visiveis=" & vis & " posicao=" & pos

    If linhaAbs < pos Then
        tbl.verticalScrollbar.Position = linhaAbs
        Err.Clear
        Set tbl = sessao.findById(idTbl)
        If Err.Number <> 0 Or tbl Is Nothing Then
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        pos = tbl.verticalScrollbar.Position
        Err.Clear
    End If

    voltas = 0
    Do While linhaAbs >= pos + vis
        voltas = voltas + 1
        If voltas > maxVoltas Then
            geoGrade = geoGrade & " / travou rolando (" & maxVoltas & " paginas)"
            On Error GoTo 0
            Exit Function
        End If
        tbl.verticalScrollbar.Position = pos + vis
        If Err.Number <> 0 Then
            Err.Clear
            geoGrade = geoGrade & " / o SAP recusou rolar para " & (pos + vis)
            On Error GoTo 0
            Exit Function
        End If
        Set tbl = sessao.findById(idTbl)
        If Err.Number <> 0 Or tbl Is Nothing Then
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        novo = -1
        novo = tbl.verticalScrollbar.Position
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        vis = tbl.VisibleRowCount
        If Err.Number <> 0 Or vis < 1 Then
            Err.Clear
            vis = 1
        End If
        If novo <= pos Then
            geoGrade = geoGrade & " / parou de rolar em " & novo & " (fim da tabela)"
            On Error GoTo 0
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
    On Error GoTo 0
End Function

' Prepara a grade antes de qualquer leitura, reproduzindo os gestos da
' gravacao ReciboBoleto1.vbs. Em GuiTableControl o GetCell devolve o
' texto RENDERIZADO: coluna estreita trunca o que o executor le, e a
' lista pode vir por montar. Alargar e dar REFRESH resolve os dois.
Sub PrepararGrade()
    Dim tbl, colVal, larg

    larg = Trim(Mapa("MIRO_LARGURA_COLUNA_VALOR"))
    colVal = IdxColuna(Mapa("MIRO_VALOR_ITEM_LINHA"))
    If Len(larg) > 0 And colVal >= 0 Then
        Set tbl = Objeto(IdTabela(Mapa("MIRO_VALOR_ITEM_LINHA")))
        If Not (tbl Is Nothing) Then
            On Error Resume Next
            tbl.Columns.ElementAt(colVal).Width = CLng(larg)
            Err.Clear
            On Error GoTo 0
        End If
    End If

    If Len(Mapa("MIRO_BOTAO_REFRESH_GRADE")) > 0 Then
        Pressionar Mapa("MIRO_BOTAO_REFRESH_GRADE")
        Tecla 0
    End If
End Sub

' Le o pedido de cada linha ate a grade acabar. TRES PASSADAS: a grade
' as vezes so termina de montar depois de um Enter e de mais um tempo
' de espera (corrigido em 05/10/2026 - ver comentario acima, na Fornecedor X
' NF 0001234).
Function LerPedidosDaGrade()
    Dim idTbl, colPed, tentativa, achadas, saida, cel, texto, maxItens
    idTbl = IdTabela(Mapa("MIRO_PEDIDO_ITEM_LINHA"))
    colPed = IdxColuna(Mapa("MIRO_PEDIDO_ITEM_LINHA"))
    maxItens = 500
    If Len(Mapa("MAX_PEDIDOS_POR_NOTA")) > 0 Then maxItens = CLng(Mapa("MAX_PEDIDOS_POR_NOTA"))

    LerPedidosDaGrade = Array()
    achadas = 0

    For tentativa = 1 To 3
        If tentativa > 1 Then
            If achadas > 0 Then Exit For
            WScript.Sleep 1500
            Tecla 0
        End If
        ReDim saida(maxItens)
        achadas = 0
        Do While achadas < maxItens
            Set cel = CelGrade(idTbl, colPed, achadas)
            If cel Is Nothing Then Exit Do
            texto = Trim(Ler(cel.Id))
            If Len(texto) = 0 Then Exit Do
            saida(achadas) = texto
            achadas = achadas + 1
        Loop
    Next

    If achadas = 0 Then Exit Function
    ReDim Preserve saida(achadas - 1)
    LerPedidosDaGrade = saida
End Function

' Escreve quantidade (so se vier vazia) e o valor liquido na linha do
' pedido DELE. Casamento por valor de pedido, nunca por posicao.
Function PreencherGrade(pedidos, rateio, gradePed)
    Dim idTbl, colVal, colQtd, n, m, alvo, candidatos, cel, qtdTxt
    PreencherGrade = ""
    idTbl = IdTabela(Mapa("MIRO_VALOR_ITEM_LINHA"))
    colVal = IdxColuna(Mapa("MIRO_VALOR_ITEM_LINHA"))
    If colVal < 0 Then
        PreencherGrade = "A coluna de valor da grade nao esta mapeada."
        Exit Function
    End If
    colQtd = ColunaQuantidade(idTbl)

    For n = 0 To UBound(gradePed)
        alvo = -1
        candidatos = 0
        For m = 0 To UBound(pedidos)
            If SoDigitos(pedidos(m)) = SoDigitos(gradePed(n)) Then
                alvo = m
                candidatos = candidatos + 1
            End If
        Next
        If candidatos > 1 Then
            PreencherGrade = "O pedido " & gradePed(n) & " aparece mais de uma vez na lista da " & _
                             "planilha. Empate nao se escolhe: nota retida."
            Exit Function
        End If
        If alvo < 0 Then
            PreencherGrade = "A linha " & n & " da grade traz o pedido " & gradePed(n) & _
                             ", que nao esta na lista de pedidos desta nota."
            Exit Function
        End If
        If Len(Trim(rateio(alvo))) = 0 Then
            PreencherGrade = "O pedido " & gradePed(n) & " esta sem VALOR LIQUIDO POR PEDIDO na " & _
                             "planilha. Nao invento valor fiscal."
            Exit Function
        End If

        ' QUANTIDADE ANTES DO VALOR: informar quantidade faz o SAP
        ' recalcular o montante da linha, entao o valor liquido tem de
        ' ser a ultima coisa escrita. Quantidade ja preenchida nao se toca.
        If colQtd >= 0 Then
            Set cel = CelGrade(idTbl, colQtd, n)
            If cel Is Nothing Then
                PreencherGrade = "Nao consegui ler a quantidade da linha " & n & " (pedido " & _
                                 gradePed(n) & ") [" & geoGrade & "]"
                Exit Function
            End If
            qtdTxt = Trim(Ler(cel.Id))
            If Len(qtdTxt) = 0 Then
                Escrever cel.Id, "1"
                If Trim(Ler(cel.Id)) <> "1" Then
                    PreencherGrade = "A quantidade do pedido " & gradePed(n) & _
                                     " veio vazia e nao consegui informar 1."
                    Exit Function
                End If
                obsGrade = obsGrade & "quantidade do pedido " & gradePed(n) & _
                           " veio vazia e foi informada como 1; "
            End If
        End If

        Set cel = CelGrade(idTbl, colVal, n)
        If cel Is Nothing Then
            PreencherGrade = "Nao consegui chegar na linha " & n & " da grade de itens [" & _
                             geoGrade & "]"
            Exit Function
        End If
        Escrever cel.Id, Trim(rateio(alvo))
        ' releitura como prova: escrever nao e prova de que ficou
        If Moeda(Ler(cel.Id)) <> Moeda(rateio(alvo)) Then
            PreencherGrade = "O valor liquido do pedido " & gradePed(n) & " nao ficou: escrevi [" & _
                             Trim(rateio(alvo)) & "] e a tela devolveu [" & Trim(Ler(cel.Id)) & "]."
            Exit Function
        End If
    Next
End Function

' Coluna da quantidade. Vazia no mapa = descobrir na tela, procurando
' DRSEG-MENGE entre os campos da tabela, e anotar o achado.
Function ColunaQuantidade(idTbl)
    Dim tbl, filhos, k, item, ident
    ColunaQuantidade = -1

    If Len(Mapa("MIRO_QTD_ITEM_LINHA")) > 0 Then
        ColunaQuantidade = IdxColuna(Mapa("MIRO_QTD_ITEM_LINHA"))
        Exit Function
    End If

    Set tbl = Objeto(idTbl)
    If tbl Is Nothing Then Exit Function
    On Error Resume Next
    Set filhos = tbl.Children
    If Err.Number <> 0 Or filhos Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    For k = 0 To filhos.Count - 1
        Set item = Pegar(filhos, k)
        If Not (item Is Nothing) Then
            ident = PropDe(item, "Id")
            If InStr(UCase(ident), "DRSEG-MENGE[") > 0 Then
                ColunaQuantidade = IdxColuna(ident)
                AnotarDescoberta "MIRO_QTD_ITEM_LINHA=" & Replace(ident, _
                    "[" & ColunaQuantidade & ",", "[" & ColunaQuantidade & ",{L}")
                Exit For
            End If
        End If
    Next
End Function

' Grava o ID descoberto para poder ser colado no mapa depois.
Sub AnotarDescoberta(texto)
    Dim arq
    On Error Resume Next
    Set arq = fso.OpenTextFile(pastaBase & "\IDS_DESCOBERTOS.txt", 8, True)
    arq.WriteLine Agora() & "  " & texto
    arq.Close
    Err.Clear
    On Error GoTo 0
End Sub


' =====================================================================
'  Planilha
' =====================================================================
Function Celula(lin, col)
    Dim v
    Celula = ""
    On Error Resume Next
    v = abaLote.Cells(lin, col).Text
    If Err.Number <> 0 Then
        Err.Clear
        v = ""
    End If
    On Error GoTo 0
    If IsNull(v) Then v = ""
    Celula = CStr(v)
End Function

Sub Gravar(lin, status, numeroDoc, observacao)
    abaLote.Cells(lin, COL_STATUS).Value = status
    abaLote.Cells(lin, COL_DOCMIRO).Value = numeroDoc
    abaLote.Cells(lin, COL_ERRO).Value = observacao
    abaLote.Cells(lin, COL_QUANDO).Value = Agora()
    abaLote.Cells(lin, COL_EXEC).Value = usuarioSap & " / " & ambiente & " / " & usuarioWin
End Sub

Sub Reter(lin, motivo)
    Gravar lin, "NAO LANCADA", "", motivo & MensagemSap()
    qtdRetidas = qtdRetidas + 1
    resumo = resumo & "NF " & Trim(Celula(lin, COL_NF)) & " -> retida" & vbCrLf
End Sub


' =====================================================================
'  Mapa de campos
' =====================================================================
Function CarregarMapa(caminho)
    Dim dic, arq, linha, pos, chave, valor
    Set dic = CreateObject("Scripting.Dictionary")
    Set CarregarMapa = dic
    If Not fso.FileExists(caminho) Then Exit Function
    Set arq = fso.OpenTextFile(caminho, 1, False)
    Do Until arq.AtEndOfStream
        linha = Trim(arq.ReadLine)
        If Len(linha) > 0 Then
            If Left(linha, 1) <> "#" Then
                pos = InStr(linha, "=")
                If pos > 1 Then
                    chave = Trim(Left(linha, pos - 1))
                    valor = Trim(Mid(linha, pos + 1))
                    If Not dic.Exists(chave) Then dic.Add chave, valor
                End If
            End If
        End If
    Loop
    arq.Close
End Function

Function Mapa(chave)
    Mapa = ""
    If dicMapa.Exists(chave) Then Mapa = dicMapa(chave)
End Function

' Troca {L} pelo numero da linha, para campos de tabela.
Function LinhaId(modelo, numero)
    LinhaId = Replace(modelo, "{L}", CStr(numero))
End Function


' =====================================================================
'  SAP - acesso a tela
' =====================================================================
Function Objeto(ident)
    Set Objeto = Nothing
    If Len(ident) = 0 Then Exit Function
    On Error Resume Next
    Set Objeto = sessao.findById(ident)
    If Err.Number <> 0 Then
        Err.Clear
        Set Objeto = Nothing
    End If
    On Error GoTo 0
End Function

Function Existe(ident)
    Dim o
    Set o = Objeto(ident)
    Existe = Not (o Is Nothing)
End Function

Function Ler(ident)
    Dim o
    Ler = ""
    Set o = Objeto(ident)
    If o Is Nothing Then Exit Function
    On Error Resume Next
    Ler = CStr(o.Text)
    If Err.Number <> 0 Then
        Err.Clear
        Ler = ""
    End If
    On Error GoTo 0
End Function

Sub Escrever(ident, valor)
    Dim o
    If Len(ident) = 0 Then Exit Sub
    Set o = Objeto(ident)
    If o Is Nothing Then Exit Sub
    On Error Resume Next
    o.Text = valor
    Err.Clear
    On Error GoTo 0
End Sub

Sub Foco(ident)
    Dim o
    If Len(ident) = 0 Then Exit Sub
    Set o = Objeto(ident)
    If o Is Nothing Then Exit Sub
    On Error Resume Next
    o.SetFocus
    Err.Clear
    On Error GoTo 0
End Sub

Sub Pressionar(ident)
    Dim o
    If Len(ident) = 0 Then Exit Sub
    Set o = Objeto(ident)
    If o Is Nothing Then Exit Sub
    On Error Resume Next
    o.Press
    Err.Clear
    On Error GoTo 0
End Sub

Sub Selecionar(ident)
    Dim o
    If Len(ident) = 0 Then Exit Sub
    Set o = Objeto(ident)
    If o Is Nothing Then Exit Sub
    On Error Resume Next
    o.Select
    Err.Clear
    On Error GoTo 0
End Sub

Sub Tecla(codigo)
    TeclaEm "wnd[0]", codigo
End Sub

' A gravacao 2 manda Enter na wnd[1], nao na wnd[0]. Tecla na janela
' errada nao faz nada e o popup fica aberto engolindo tudo depois.
Sub TeclaEm(janela, codigo)
    Dim o
    Set o = Objeto(janela)
    If o Is Nothing Then Exit Sub
    On Error Resume Next
    o.sendVKey codigo
    Err.Clear
    On Error GoTo 0
End Sub

Function ContarJanelas()
    ContarJanelas = 0
    On Error Resume Next
    ContarJanelas = sessao.Children.Count
    If Err.Number <> 0 Then ContarJanelas = 0
    Err.Clear
    On Error GoTo 0
End Function

' O numero do documento sai da barra de status. Sem numero, nao se
' afirma que gravou.
' A resposta costuma estar na barra de status. Toda retencao leva a
' mensagem do SAP junto do motivo - "a tela nao abriu" nao e motivo.
Function MensagemSap()
    Dim o, tipo, texto
    MensagemSap = ""
    Set o = Objeto(Mapa("MIRO_BARRA_STATUS"))
    If o Is Nothing Then Exit Function
    tipo = PropDe(o, "MessageType")
    texto = PropDe(o, "Text")
    If Len(Trim(texto)) = 0 Then Exit Function
    MensagemSap = "  |  SAP [" & Trim(tipo) & "]: " & Trim(texto)
End Function

Function NumeroNaBarra()
    Dim texto, k, ch, corrente, melhor
    NumeroNaBarra = ""
    texto = Ler(Mapa("MIRO_BARRA_STATUS"))
    corrente = ""
    melhor = ""
    For k = 1 To Len(texto)
        ch = Mid(texto, k, 1)
        If ch >= "0" And ch <= "9" Then
            corrente = corrente & ch
        Else
            If Len(corrente) > Len(melhor) Then melhor = corrente
            corrente = ""
        End If
    Next
    If Len(corrente) > Len(melhor) Then melhor = corrente
    If Len(melhor) >= 8 Then NumeroNaBarra = melhor
End Function


' =====================================================================
'  Texto e numero
' =====================================================================
' Tira apenas os separadores listados no mapa. Letras e numeros ficam:
' o codigo de verificacao nao e um numero, e uma identificacao.
Function SemSeparadores(texto)
    Dim lista, k, saida, ch
    lista = Mapa("COD_VERIFICACAO_LIMPAR")
    saida = ""
    For k = 1 To Len(texto)
        ch = Mid(texto, k, 1)
        If InStr(lista, ch) = 0 Then saida = saida & ch
    Next
    SemSeparadores = Trim(saida)
End Function

Function SoDigitos(texto)
    Dim k, ch, saida
    saida = ""
    For k = 1 To Len(texto)
        ch = Mid(texto, k, 1)
        If ch >= "0" And ch <= "9" Then saida = saida & ch
    Next
    SoDigitos = saida
End Function

' Quebra "4500000004 e 4500000005" ou "100,01 e 100,00" em vetor.
Function Partir(texto)
    Dim bruto, partes, k, limpo, saida, total
    bruto = Trim(texto)
    If Len(bruto) = 0 Then
        Partir = Array()
        Exit Function
    End If
    bruto = Replace(bruto, " E ", " e ")
    bruto = Replace(bruto, ";", " e ")
    partes = Split(bruto, " e ")
    total = -1
    ReDim saida(UBound(partes))
    For k = 0 To UBound(partes)
        limpo = Trim(partes(k))
        If Len(limpo) > 0 Then
            total = total + 1
            saida(total) = limpo
        End If
    Next
    If total < 0 Then
        Partir = Array()
    Else
        ReDim Preserve saida(total)
        Partir = saida
    End If
End Function

' Converte para o formato que o SAP espera, sem CDbl e sem depender de
' locale: passa por centavos inteiros.
Function Moeda(texto)
    Dim bruto, k, ch, digitos, virgula, casas, inteiro, centavos
    Moeda = ""
    bruto = Trim(texto)
    If Len(bruto) = 0 Then Exit Function

    digitos = ""
    virgula = 0
    casas = 0
    For k = 1 To Len(bruto)
        ch = Mid(bruto, k, 1)
        If ch >= "0" And ch <= "9" Then
            digitos = digitos & ch
            If virgula = 1 Then casas = casas + 1
        ElseIf ch = "," Or ch = "." Then
            ' o ultimo separador com 1 ou 2 digitos depois e o decimal
            If Len(bruto) - k <= 2 Then
                virgula = 1
                casas = 0
            End If
        End If
    Next
    If Len(digitos) = 0 Then Exit Function

    If virgula = 0 Then
        centavos = digitos & "00"
    ElseIf casas = 1 Then
        centavos = digitos & "0"
    ElseIf casas = 0 Then
        centavos = digitos & "00"
    Else
        centavos = digitos
    End If

    Do While Len(centavos) < 3
        centavos = "0" & centavos
    Loop
    inteiro = Left(centavos, Len(centavos) - 2)
    Do While Len(inteiro) > 1 And Left(inteiro, 1) = "0"
        inteiro = Mid(inteiro, 2)
    Loop
    Moeda = inteiro & "," & Right(centavos, 2)
End Function

Function Agora()
    Dim d
    d = Now()
    Agora = Right("00" & Day(d), 2) & "/" & Right("00" & Month(d), 2) & "/" & Year(d) & " " & _
            Right("00" & Hour(d), 2) & ":" & Right("00" & Minute(d), 2) & ":" & Right("00" & Second(d), 2)
End Function

' Carimbo de data e hora para nome de arquivo: 20260903_152401
Function Carimbo()
    Dim d
    d = Now()
    Carimbo = Right("0000" & Year(d), 4) & Right("00" & Month(d), 2) & _
              Right("00" & Day(d), 2) & "_" & Right("00" & Hour(d), 2) & _
              Right("00" & Minute(d), 2) & Right("00" & Second(d), 2)
End Function

Function NomeWindows()
    NomeWindows = ""
    On Error Resume Next
    NomeWindows = CreateObject("WScript.Network").UserName
    If Err.Number <> 0 Then NomeWindows = "(desconhecido)"
    Err.Clear
    On Error GoTo 0
End Function

' Modo pedido por argumento tem prioridade sobre o mapa: assim da para
' tirar um despejo sem editar arquivo nenhum.
Function ModoDespejo()
    Dim k, pedido
    ModoDespejo = False
    pedido = UCase(Trim(Mapa("MODO")))
    On Error Resume Next
    For k = 0 To WScript.Arguments.Count - 1
        If UCase(Trim(WScript.Arguments(k))) = "DESPEJO" Then pedido = "DESPEJO"
    Next
    Err.Clear
    On Error GoTo 0
    If pedido = "DESPEJO" Then ModoDespejo = True
End Function

Function AmbientePermitido(atual)
    Dim lista, k
    AmbientePermitido = False
    lista = Split(Mapa("SAP_AMBIENTES_PERMITIDOS"), "|")
    For k = 0 To UBound(lista)
        If UCase(Trim(lista(k))) = UCase(Trim(atual)) Then AmbientePermitido = True
    Next
End Function


'  Junta COL_INFO1..COL_INFO3 e as quebras de linha dentro delas numa
'  lista de linhas de texto, na ordem em que aparecem.
Function MontarMensagens(lin)
    Dim bruto, col, partes, k, saida, total
    total = -1
    ReDim saida(200)
    For col = COL_INFO1 To COL_INFO3
        bruto = Celula(lin, col)
        bruto = Replace(bruto, vbCrLf, vbLf)
        bruto = Replace(bruto, vbCr, vbLf)
        partes = Split(bruto, vbLf)
        For k = 0 To UBound(partes)
            If Len(Trim(partes(k))) > 0 Then
                total = total + 1
                saida(total) = Trim(partes(k))
            End If
        Next
    Next
    If total < 0 Then
        MontarMensagens = Array()
    Else
        ReDim Preserve saida(total)
        MontarMensagens = saida
    End If
End Function

'  Escreve e RELE. Texto fiscal nao e truncado nem remendado: se nao
'  couber ou nao ficar, a nota retem.
Function EscreverMensagens(mensagens)
    Dim k, limite, idCampo, devolvido
    EscreverMensagens = ""
    limite = 0
    If Len(Mapa("NF_MENSAGEM_LIMITE_CARACTERES")) > 0 Then
        limite = CLng(Mapa("NF_MENSAGEM_LIMITE_CARACTERES"))
    End If

    For k = 0 To UBound(mensagens)
        If limite > 0 Then
            If Len(mensagens(k)) > limite Then
                EscreverMensagens = "A linha " & (k + 1) & " de Outras informacoes tem " & _
                    Len(mensagens(k)) & " caracteres, acima do limite de " & limite & _
                    ". Nao trunco texto fiscal por conta propria."
                Exit Function
            End If
        End If
        idCampo = LinhaId(Mapa("NF_MENSAGEM_LINHA"), k)
        If Not Existe(idCampo) Then
            EscreverMensagens = "A nota tem " & (UBound(mensagens) + 1) & " linhas de Outras " & _
                "informacoes e a tela nao ofereceu a linha " & (k + 1) & "."
            Exit Function
        End If
        Escrever idCampo, mensagens(k)
    Next

    ' releitura como prova: apertar nao e prova de que ficou
    For k = 0 To UBound(mensagens)
        devolvido = Trim(Ler(LinhaId(Mapa("NF_MENSAGEM_LINHA"), k)))
        If devolvido <> Trim(mensagens(k)) Then
            EscreverMensagens = "A linha " & (k + 1) & " de Outras informacoes nao ficou: escrevi [" & _
                mensagens(k) & "] e a tela devolveu [" & devolvido & "]."
            Exit Function
        End If
    Next
End Function


' =====================================================================
'  Despejo de tela no momento da falha
'  Um despejo custa 40 linhas de codigo e resolve o que a teoria nao
'  resolve. Vale mais que qualquer hipotese.
' =====================================================================
Dim arqDespejo

Sub Despejar(nomeArquivo)
    Dim k, janelas
    On Error Resume Next
    Set arqDespejo = fso.CreateTextFile(pastaBase & "\" & nomeArquivo, True, False)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Anotar "DESPEJO DE FALHA - " & Agora()
    Anotar "Ambiente : " & ambiente & "   Usuario SAP: " & usuarioSap
    Anotar "Transacao: " & sessao.Info.Transaction & "   Tela: " & sessao.Info.ScreenNumber
    Anotar "Barra    : [" & Ler(Mapa("MIRO_BARRA_STATUS")) & "]"
    Anotar ""
    Set janelas = sessao.Children
    Anotar "JANELAS ABERTAS: " & janelas.Count
    Anotar ""
    For k = 0 To janelas.Count - 1
        Anotar "########## wnd[" & k & "] ##########"
        Caminhar Pegar(janelas, k), 0
        Anotar ""
    Next
    arqDespejo.Close
End Sub

Sub Caminhar(obj, nivel)
    Dim filhos, k, item, tipo, ident, valor, recuo
    If obj Is Nothing Then Exit Sub
    If nivel > 30 Then Exit Sub

    tipo = PropDe(obj, "Type")
    ident = PropDe(obj, "Id")
    valor = PropDe(obj, "Text")
    If nivel > 0 Then
        recuo = String(nivel * 2, " ")
    Else
        recuo = ""
    End If
    If Len(valor) = 0 And EhCampoTexto(tipo) Then valor = "<VAZIO>"
    Anotar recuo & "<" & tipo & ">  " & ident & "  nome=[" & PropDe(obj, "Name") & _
           "]  texto=[" & valor & "]"

    ' Tabela merece tratamento proprio: e nela que moram a grade de itens
    ' e a lista de pedidos, e um despejo sem o conteudo delas nao explica
    ' nada quando a grade volta vazia.
    If tipo = "GuiTableControl" Then AnotarTabela obj, nivel

    On Error Resume Next
    Set filhos = Nothing
    Set filhos = obj.Children
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    If filhos Is Nothing Then Exit Sub

    For k = 0 To filhos.Count - 1
        Set item = Pegar(filhos, k)
        If Not (item Is Nothing) Then Caminhar item, nivel + 1
    Next
End Sub

' Campo vazio NAO e a mesma coisa que campo inexistente: o despejo
' precisa distinguir os dois.
Function EhCampoTexto(tipo)
    EhCampoTexto = (tipo = "GuiTextField") Or (tipo = "GuiCTextField") Or _
                   (tipo = "GuiPasswordField") Or (tipo = "GuiComboBox")
End Function

' RowCount NAO e a quantidade de itens: inclui as linhas de entrada
' vazias. Uma nota de 3 itens devolve 25.
Sub AnotarTabela(tab, nivel)
    Dim totalLinhas, totalColunas, lin, col, cel, texto, linha, limite, cru, recuo
    recuo = ""
    If nivel >= 0 Then recuo = String((nivel + 1) * 2, " ")

    On Error Resume Next
    totalLinhas = tab.RowCount
    totalColunas = tab.Columns.Count
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Anotar recuo & "-- TABELA: RowCount=" & totalLinhas & " (inclui linhas vazias)" & _
           "  VisibleRowCount=" & PropDe(tab, "VisibleRowCount") & "  Colunas=" & totalColunas

    limite = totalLinhas - 1
    If limite > 40 Then limite = 40
    For lin = 0 To limite
        linha = ""
        cru = ""
        For col = 0 To totalColunas - 1
            texto = ""
            On Error Resume Next
            Set cel = Nothing
            Set cel = tab.GetCell(lin, col)
            If Not (cel Is Nothing) Then texto = cel.Text
            Err.Clear
            On Error GoTo 0
            linha = linha & "[" & texto & "]"
            cru = cru & texto
        Next
        If Trim(cru) <> "" Then
            Anotar recuo & "   linha " & lin & ": " & linha
        Else
            Anotar recuo & "   linha " & lin & ": <linha vazia>"
        End If
    Next
End Sub

' Children(i) nao devolve objeto: ElementAt, com Item de reserva.
Function Pegar(colecao, indice)
    Dim obtido
    Set obtido = Nothing
    On Error Resume Next
    Set obtido = colecao.ElementAt(indice)
    If obtido Is Nothing Then
        Err.Clear
        Set obtido = colecao.Item(indice)
    End If
    Err.Clear
    On Error GoTo 0
    Set Pegar = obtido
End Function

Function PropDe(obj, propriedade)
    Dim r
    r = ""
    On Error Resume Next
    Select Case propriedade
        Case "Type" r = obj.Type
        Case "Id"   r = obj.Id
        Case "Text" r = obj.Text
        Case "Name" r = obj.Name
        Case "VisibleRowCount" r = obj.VisibleRowCount
        Case "MessageType" r = obj.MessageType
    End Select
    If Err.Number <> 0 Then
        r = ""
        Err.Clear
    End If
    On Error GoTo 0
    PropDe = CStr(r)
End Function

Sub Anotar(texto)
    On Error Resume Next
    arqDespejo.WriteLine texto
    Err.Clear
    On Error GoTo 0
End Sub
