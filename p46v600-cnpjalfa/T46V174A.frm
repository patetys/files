VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmT46V174A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de cota��es sem agendamento de inspe��o"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   2190
   ClientWidth     =   12390
   Icon            =   "T46V174A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand cmdSolicitar 
      Height          =   360
      Left            =   11085
      TabIndex        =   1
      Top             =   8085
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1879
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Solicitar"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdSair 
      Height          =   360
      Left            =   11085
      TabIndex        =   0
      Top             =   8610
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1879
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "S&air"
      ForeColor       =   8388608
   End
   Begin TabDlg.SSTab TabInspecao 
      Height          =   8445
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   14896
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   10485760
      TabCaption(0)   =   "Syas"
      TabPicture(0)   =   "T46V174A.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flxCotacoes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Yasudanet"
      TabPicture(1)   =   "T46V174A.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "WebBrowser"
      Tab(1).ControlCount=   1
      Begin SHDocVwCtl.WebBrowser WebBrowser 
         Height          =   7875
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   11895
         ExtentX         =   20981
         ExtentY         =   13891
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid flxCotacoes 
         Height          =   7365
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   12991
         _Version        =   393216
         FixedCols       =   0
         ForeColor       =   8388608
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
End
Attribute VB_Name = "frmT46V174A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private objPedido As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED)
Private objPedLoc As clsA46V708A    'Objeto da classe clsA46V708A (TAB_PED_LOC)
Private objPedInsp As clsA46V717A    'Objeto da classe clsA46V717A (TAB_PED_INSP)

Private objstcPedido As stcA46V702B   'Objeto da classe de estrutura stcA46V702B (TAB_PED)
Private objstcPedInsp As stcA46V717B    'Objeto da classe de estrutura stcA46V717B (TAB_PED_INSP)

Private lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
Private lobjstcPedInsp As stcA46V717B     'Objeto da classe de estrutura stcA46V717B (TAB_PED_INSP)

Public gNosNumeroInsp As String     'Nosso n�mero.
Public gNumItemInsp As String   'N�mero do item da ap�lice.
Public gNumPedidoInsp As String     'N�mero de pedido.

Public gAbortSolic As Boolean    'Flag para abortar solicita��o de inspe��o da tela T46V175B

Private Sub cmdSair_Click()
    Unload Me
End Sub

Public Sub cmdSolicitar_Click()
    Dim lMensagem As String     'Mensagem.
    Dim bdBasDados As ADODB.Connection
    Dim lrst As ADODB.Recordset    'Este recordSet esta usando a propriedade Filter (nao implementado no YasRecordSet)
    Dim lSelect As String

    '1100008 - Inspe��o por Lote
    'Essa rotina foi re-estruturada para contemplar a requisi��o da ficha acima.

    On Error GoTo ControleErro

    With flxCotacoes
        If .Rows = 1 Then Exit Sub
        gNumPedidoInsp = .TextMatrix(.RowSel, 1)    'N�mero do pedido.
        gNumItemInsp = .TextMatrix(.RowSel, 2)   'N�mero do item.
        gNosNumeroInsp = .TextMatrix(.RowSel, 7)    'Nosso n�mero.

        If MsgBox("Deseja Solicitar Inspe��o para o Pedido " & gNumPedidoInsp & " Item " & gNumItemInsp & " ?", vbQuestion + vbYesNo, "Aten��o") = vbNo Then
            Exit Sub
        End If

        If gM46V999.gfDisponivel("frmT46V175B") Then
            Unload frmT46V175B
        End If
        frmT46V175B.Show vbModal, Me
        If gAbortSolic = True Then Exit Sub

        Screen.MousePointer = vbHourglass
        cmdSolicitar.Enabled = False
        cmdSair.Enabled = False
        flxCotacoes.Enabled = False

        'abre base de dados
        If Not gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, lMensagem) Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If

        'verifica se existe outros itens
        lSelect = "Select * From TAB_PED P Inner Join TAB_PED_LOC L"
        lSelect = lSelect & " On P.Nosso_Numero = L.Nosso_Numero "
        lSelect = lSelect & " Where P.Nosso_Numero = " & gM46V999.gfFormatarCamposSQL(gNosNumeroInsp)
        lSelect = lSelect & " And L.Insp_Obr = 1 And L.Insp_Agend = 0 "

        'Set lrst = New clsYasRecordSet
        Set lrst = New ADODB.Recordset
        Call lrst.Open(lSelect, bdBasDados)    'Este recordSet esta usando a propriedade Filter (nao implementado no YasRecordSet)

        If lrst.RecordCount > 1 Then
            If MsgBox("Existe mais de um item para o Pedido " & gNumPedidoInsp & vbCrLf & "Deseja gerar inspe��o para todos os itens com os mesmos dados de inspe��o?", vbQuestion + vbYesNo, "Aten��o") = vbNo Then
                lrst.Filter = " Num_Item = " & gNumItemInsp
            End If
        End If

        '3. Instancio o objeto.
        Set objPedido = New clsA46V702A
        Set objPedido.M46V999 = gM46V999

        Set objstcPedido = New stcA46V702B
        Set objstcPedido.M46V999 = gM46V999

        objstcPedido.Num_Ped = gNumPedidoInsp
        objstcPedido.NOSSO_NUMERO = gNosNumeroInsp

        '4. Busco os pedidos.
        If Not objPedido.mfCarregarColecao(objstcPedido, "", "", True) Then
            Call gM46V999.gpGraLog(1, objPedido.UltimoErro)
            Exit Sub
        End If

        '5. Se n�o encontrou registro, desabilita controles.
        If objPedido.Colecao.Count = 0 Then
            Call gM46V999.gpGraLog(1, "INSP0174 - N�o h� cota��o dentro do prazo de solicita��o de inspe��o.", vbInformation)
            Exit Sub
        End If
        Set objstcPedido = objPedido.Colecao.Item(1)

        Do Until lrst.EOF
            Set lobjstcPedLoc = objstcPedido.ColecaoPedLoc.Item(objstcPedido.NOSSO_NUMERO & Format(lrst("num_item"), "00000"))
            lobjstcPedLoc.Insp_Agend = True
            'salva dados da inspe��o
            If Not mfSalvarInspecao(lMensagem, lobjstcPedLoc, objstcPedido) Then
                Call gM46V999.gpGraLog(0, "mfSalvarInspecao ERRO - " & Err & " " & Error & Chr(13) & lMensagem)
            End If
            'altera status inspe��o PedLoc
            Set objPedLoc = New clsA46V708A
            Set objPedLoc.M46V999 = gM46V999

            If Not objPedLoc.mfAlterarAgendaInspecao(lobjstcPedLoc, "", "") Then
                lMensagem = objPedLoc.UltimoErro
                Exit Sub
            End If
            'gera arquivo de inspe��o
            Call gpGeraLogCalcInspecao(objstcPedido, lobjstcPedLoc)
            lrst.MoveNext
        Loop

    End With

    Set objPedido = Nothing
    Set objstcPedido = Nothing
    Set objPedLoc = Nothing


    'recarrega grid
    Call Form_Load

    cmdSolicitar.Enabled = True
    cmdSair.Enabled = True
    flxCotacoes.Enabled = True
    MsgBox "Gerada " & IIf(lrst.RecordCount > 1, "solicita��es", "solicita��o") & " de inspe��o para " & IIf(lrst.RecordCount > 1, lrst.RecordCount & " itens", "o item " & gNumItemInsp) & " da cota��o " & gNumPedidoInsp & "." & vbLf & "Utilize a op��o de Conex�o Internet Yasuda para completar o processo de solicita��o.", vbInformation
    Call gM46V999.gpFecha2(bdBasDados)
    Screen.MousePointer = vbDefault

    Exit Sub

ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V174A.cmdSolicitar_Click ERRO - " & Err & " " & Err.Description)
End Sub



Private Sub flxCotacoes_DblClick()
    cmdSolicitar.Enabled = True
    If flxCotacoes.Rows = 1 Then
        Call gM46V999.gpGraLog(1, "INSP0174 - N�o h� cota��o Syas dentro do prazo de solicita��o de inspe��o.", vbInformation)
        TabInspecao.Tab = 1
        cmdSolicitar.Enabled = False
        Exit Sub
    End If

    Call cmdSolicitar_Click
End Sub

Private Sub Form_Load()
    Dim lCount As Double    'Contador.
    Dim lCpfCnpj As String    'Retorno da fun��o de forma��o de cpf e cnpj.
    Dim lMensagem As String

    '1. Posiciona form.
    Call gM46V999.gpCenForm(Me)

    flxCotacoes.FixedCols = 0
    flxCotacoes.FixedRows = 1
    flxCotacoes.Cols = 10
    flxCotacoes.Row = 1

    If gM46V999.gCod_Corretor <> 0 Then
        If Not gM46V999.gfBuscaSusep(gM46V999.gCod_Corretor, gM46V999.gNumSusep, lMensagem) Then
            Exit Sub
        End If
    End If


    '3. Instancio o objeto.
    Set objPedido = New clsA46V702A
    Set objPedido.M46V999 = gM46V999

    Set objstcPedido = New stcA46V702B
    Set objstcPedido.M46V999 = gM46V999

    objstcPedido.Cod_Sit_Ped = e_SitPedido_ComCalculo
    objstcPedido.Cod_Tip_Doc = e_TipDoc_Cotacao

    '4. Busco os pedidos.
    gM46V999.gAgendarCot = True
    If Not objPedido.mfCarregarColecao(objstcPedido, "", "", True) Then
        Call gM46V999.gpGraLog(1, objPedido.UltimoErro)
        Exit Sub
    End If
    gM46V999.gAgendarCot = False

    '5. Se n�o encontrou registro, desabilita controles.
    cmdSolicitar.Enabled = True
    If objPedido.Colecao.Count = 0 Then
        flxCotacoes.Visible = False
        Call gM46V999.gpGraLog(1, "INSP0174 - N�o h� cota��o dentro do prazo de solicita��o de inspe��o.", vbInformation)
        cmdSolicitar.Enabled = False
        Exit Sub
    End If

    '6. Monta grades.
    With flxCotacoes
        'a) Posiciona quantidade de colunas e de linhas (apenas cabe�alho).
        .Cols = 10
        .Rows = 1
        'b) Posiciona linha de cabe�alho (com altura).
        .Row = 0
        'c) Posiciona cabe�alho nas colunas.
        .TextMatrix(0, 0) = "Ramo"
        .TextMatrix(0, 1) = "Pedido"
        .TextMatrix(0, 2) = "Item"
        .TextMatrix(0, 3) = "Segurado"
        .TextMatrix(0, 4) = "CNPJ/CPF"
        .TextMatrix(0, 5) = "In�cio Vig�ncia"
        .TextMatrix(0, 6) = "T�rmino Vig�ncia"
        .TextMatrix(0, 7) = "Nosso N�mero"
        .TextMatrix(0, 8) = "Insp.Obrigat�ria?"
        .TextMatrix(0, 9) = "Insp.Agendada?"
        'd) Posiciona largura das colunas.
        .ColWidth(0) = 1200
        .ColWidth(1) = 800
        .ColWidth(2) = 500
        .ColWidth(3) = 2800
        .ColWidth(4) = 1600
        .ColWidth(5) = 1200
        .ColWidth(6) = 1500
        .ColWidth(7) = 0
        .ColWidth(8) = 1300
        .ColWidth(9) = 1300
        'e) Posiciona alinhamento das colunas.
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignCenterCenter
        .Visible = True
    End With
    lCount = 1

    '7. Coloca registros na grade.
    With flxCotacoes
        For Each objstcPedido In objPedido.Colecao
            For Each lobjstcPedLoc In objstcPedido.ColecaoPedLoc
                If lobjstcPedLoc.Insp_Obr = True And lobjstcPedLoc.Insp_Agend = False Then
                    'a) Acrescenta linha na grade.
                    .Rows = .Rows + 1
                    If lobjstcPedLoc.Cod_Ramo = e_CodRamo_Condominio Then
                        .TextMatrix(lCount, 0) = "Condom�nio"
                    ElseIf lobjstcPedLoc.Cod_Ramo = e_CodRamo_Empresarial Then
                        .TextMatrix(lCount, 0) = "Empresarial"
                    ElseIf lobjstcPedLoc.Cod_Ramo = e_CodRamo_Residencial Then
                        .TextMatrix(lCount, 0) = "Residencial"
                    Else
                        .TextMatrix(lCount, 0) = "N�o definido"
                    End If
                    .TextMatrix(lCount, 1) = lobjstcPedLoc.Num_Ped
                    .TextMatrix(lCount, 2) = lobjstcPedLoc.Num_Item
                    .TextMatrix(lCount, 3) = objstcPedido.Nom_Segurado
                    Select Case Len(Trim(CStr(objstcPedido.Num_Cgc_Cpf)))
                    Case Is > 11
                        If IsNumeric(objstcPedido.Num_Cgc_Cpf) Then
                            Call gM46V999.gpForCNPJ(1, Format$(objstcPedido.Num_Cgc_Cpf, "00000000000000"), lCpfCnpj)
                        Else
                            Call gM46V999.gpForCNPJ(1, Right$(String(14, "0") & Trim(objstcPedido.Num_Cgc_Cpf), 14), lCpfCnpj)
                        End If
                    Case Else
                        Call gM46V999.gpForCPF(1, Format$(objstcPedido.Num_Cgc_Cpf, "00000000000"), lCpfCnpj)
                    End Select
                    .TextMatrix(lCount, 4) = lCpfCnpj
                    .TextMatrix(lCount, 5) = gM46V999.gfForData(2, objstcPedido.Dat_Ini_Vig)
                    .TextMatrix(lCount, 6) = gM46V999.gfForData(2, objstcPedido.Dat_Ter_Vig)
                    .TextMatrix(lCount, 7) = objstcPedido.NOSSO_NUMERO
                    If lobjstcPedLoc.Insp_Obr = True Then
                        .TextMatrix(lCount, 8) = "SIM"
                    Else
                        .TextMatrix(lCount, 8) = "N�O"
                    End If
                    If lobjstcPedLoc.Insp_Agend = True Then
                        .TextMatrix(lCount, 9) = "SIM"
                    Else
                        .TextMatrix(lCount, 9) = "N�O"
                    End If
                    lCount = lCount + 1
                    .Refresh
                End If
            Next lobjstcPedLoc
Proximo:
        Next objstcPedido
        cmdSolicitar.Enabled = True
        If .Rows = 1 Then
            Call gM46V999.gpGraLog(1, "INSP0174 - N�o h� cota��o Syas dentro do prazo de solicita��o de inspe��o.", vbInformation)
            TabInspecao.Tab = 1
            cmdSolicitar.Enabled = False
            Exit Sub
        End If

        .TopRow = 1
        .Refresh
        .ColSel = .Cols - 1
        .Row = 1
    End With
    'TabInspecao.Tab = 0
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V174A.Load ERRO - " & Err & " " & Err.Description)

End Sub

Function mfPesqCotAgend()
    Dim bdBasDados As ADODB.Connection    'Conex�o do Banco local das TAB_PED tempor�rias
    Dim lobjPedLoc As clsA46V708A    'Objeto da classe de itens.
    Dim lColecaoTemp As Collection
    Dim lMensagem As String    'Mensagem.
    Dim lrst As clsYasRecordSet    'Registro.
    Dim lSelect As String    'Select
    Dim lCpfCnpj As String    'Retorno da fun��o de forma��o de cpf e cnpj.


    Dim lCount As Double    'Contador.
    Dim lI As Double    'Auxiliar.
    Dim lColuna As Integer    'Coluna do grid.
    Dim lObjPed As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED).
    Dim lObjstcPed As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED).
    Dim lObjPedCorr As clsA46V705A    'Objeto da classe clsA46V705A (TAB_PED_CORR).
    Dim lObjstcPedCorr As stcA46V705B    'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).
    Dim lCod_Ramo As e_CodRamo


    On Error GoTo ControleErro

    mfPesqCotAgend = False

    If Not gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If

    'para testar
    gM46V999.gTipUsuario = e_Tipo_Usu_Corretor

    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        lSelect = "Select * from TAB_PED A inner join TAB_PED_LOC B "
        lSelect = lSelect & " ON a.Num_Ped = b.Num_Ped "
        lSelect = lSelect & " where Insp_Obr = 1 and Insp_Agend = 0 "
        '        lSelect = lSelect & " and a.Dat_Pri_Calculo >= " & gM46V999.gfForData(1, Date - 30)

        'Busca os pr�mios de coberturas
        If Not gM46V999.gfObtRegistro(bdBasDados, lSelect, lrst, lMensagem) Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Function
        End If

        'Se n�o encontrou registro, informa usu�rio.
        If lrst.EOF Then
            lMensagem = "FNL0174 - N�o h� cota��es pendentes de solicita��o de inspe��o." & vbLf
            Call gM46V999.gpGraLog(1, lMensagem)
        End If

        'Monta grade de resultados.
        With flxCotacoes
            'a) Posiciona quantidade de colunas e de linhas (apenas cabe�alho).
            .Cols = 8
            .Rows = 1
            'b) Posiciona linha de cabe�alho (com altura).
            .Row = 0
            'c) Posiciona cabe�alho nas colunas.
            .TextMatrix(0, 0) = "Ramo"
            .TextMatrix(0, 1) = "Pedido"
            .TextMatrix(0, 2) = "Item"
            .TextMatrix(0, 3) = "Segurado"
            .TextMatrix(0, 4) = "CNPJ/CPF"
            .TextMatrix(0, 5) = "In�cio Vig�ncia"
            .TextMatrix(0, 6) = "T�rmino Vig�ncia"
            .TextMatrix(0, 7) = "Nosso N�mero"
            'd) Posiciona largura das colunas.
            .ColWidth(0) = 1620
            .ColWidth(1) = 800
            .ColWidth(2) = 500
            .ColWidth(3) = 3200
            .ColWidth(4) = 1800
            .ColWidth(5) = 1200
            .ColWidth(6) = 1500
            .ColWidth(7) = 1500
            'e) Posiciona alinhamento das colunas.
            .ColAlignment(0) = flexAlignCenterCenter
            .ColAlignment(1) = flexAlignLeftCenter
            .ColAlignment(2) = flexAlignLeftCenter
            .ColAlignment(3) = flexAlignLeftCenter
            .ColAlignment(4) = flexAlignLeftCenter
            .ColAlignment(5) = flexAlignLeftCenter
            .ColAlignment(6) = flexAlignCenterCenter
            .ColAlignment(7) = flexAlignCenterCenter
            .Visible = True
        End With
        lCount = 1

        'Coloca registros na grade.
        flxCotacoes.Redraw = False
        Do While Not lrst.EOF
            flxCotacoes.Rows = flxCotacoes.Rows + 1
            flxCotacoes.CellPictureAlignment = 4
            If lrst("Cod_Ramo") = e_CodRamo_Condominio Then
                flxCotacoes.TextMatrix(lCount, 0) = "Condom�nio"
            ElseIf lrst("Cod_Ramo") = e_CodRamo_Empresarial Then
                flxCotacoes.TextMatrix(lCount, 0) = "Empresarial"
            ElseIf lrst("Cod_Ramo") = e_CodRamo_Residencial Then
                flxCotacoes.TextMatrix(lCount, 0) = "Residencial"
            Else
                flxCotacoes.TextMatrix(lCount, 0) = "N�o definido"
            End If
            flxCotacoes.TextMatrix(lCount, 1) = lrst("Num_Ped")
            flxCotacoes.TextMatrix(lCount, 2) = lrst("Num_Item")
            flxCotacoes.TextMatrix(lCount, 3) = lrst("Nom_Segurado")
            Select Case Len(CStr(lrst("Num_Cgc_Cpf")))
            Case Is > 11
                Call gM46V999.gpForCNPJ(1, Format$(lrst("Num_Cgc_Cpf"), "00000000000000"), lCpfCnpj)
            Case Else
                Call gM46V999.gpForCPF(1, Format$(lrst("Num_Cgc_Cpf"), "00000000000"), lCpfCnpj)
            End Select
            flxCotacoes.TextMatrix(lCount, 4) = lCpfCnpj
            flxCotacoes.TextMatrix(lCount, 5) = gM46V999.gfForData(2, lrst("Dat_Ini_Vig"))
            flxCotacoes.TextMatrix(lCount, 6) = gM46V999.gfForData(2, lrst("Dat_Ter_Vig"))
            flxCotacoes.TextMatrix(lCount, 7) = lrst("Nosso_Numero")
            lCount = lCount + 1
            lrst.MoveNext
        Loop
        Me.Refresh

        flxCotacoes.Redraw = True
        If flxCotacoes.Rows > 1 Then
            flxCotacoes.TopRow = flxCotacoes.Rows - 1
        End If
        flxCotacoes.Refresh
    End If

    mfPesqCotAgend = True
    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(1, Err.Description)
End Function

Private Function mfSalvarInspecao(pMensagem As String, pobjstcPedLoc As stcA46V708B, pobjstcPedido As stcA46V702B) As Boolean
    Dim lCNPJ As String

    On Error GoTo ControleErro

    Set objPedInsp = New clsA46V717A
    Set objPedInsp.M46V999 = gM46V999

    Set objPedInsp.ConexaoBD = objPedido.ConexaoBD

    mfSalvarInspecao = False

    If gM46V999.gfPreenchido(pobjstcPedLoc.objstcPedInsp.NOSSO_NUMERO) Then
        pobjstcPedLoc.objstcPedInsp.NOSSO_NUMERO = pobjstcPedLoc.NOSSO_NUMERO
        pobjstcPedLoc.objstcPedInsp.Num_Item = pobjstcPedLoc.Num_Item
        If Not objPedInsp.mfExcluir(pobjstcPedLoc.objstcPedInsp, "", "") Then
            pMensagem = objPedInsp.UltimoErro
            GoTo Finalizar
        End If
    End If

    With pobjstcPedLoc.objstcPedInsp
        .NOSSO_NUMERO = pobjstcPedLoc.NOSSO_NUMERO
        .Num_Item = pobjstcPedLoc.Num_Item
        .Dat_Cadastro = Left(gM46V999.gfForData(3), 8)
        .Num_Logr_Risco = pobjstcPedLoc.objstcPedInsp.Num_Logr_Risco
        .Num_Tel = frmT46V175B.txtTelefone.Text
        .Pess_Cont = gM46V999.gfRetirarCaracterEspecial(frmT46V175B.txtContato.Text)
        .Cod_Ramo = pobjstcPedLoc.Cod_Ramo
        .Num_Ped = pobjstcPedLoc.Num_Ped
    End With

    If Not objPedInsp.mfIncluir(pobjstcPedLoc.objstcPedInsp, "", "", "") Then
        pMensagem = objPedInsp.UltimoErro
        GoTo Finalizar
    End If

    If pobjstcPedido.Tip_Pessoa = e_TipPessoa_Fisica Then
        Call gM46V999.gpForCPF(2, frmT46V175B.txtCNPJ.Text, lCNPJ)
    Else
        Call gM46V999.gpForCNPJ(2, frmT46V175B.txtCNPJ.Text, lCNPJ)
    End If
    pobjstcPedido.Num_Cgc_Cpf = lCNPJ

    If Not objPedido.mfAlterar(pobjstcPedido, "", "") Then
        pMensagem = objPedido.UltimoErro
        GoTo Finalizar
    End If

    mfSalvarInspecao = True

Finalizar:
    Exit Function
ControleErro:
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarInspecao ERRO - " & Err & " " & Error & _
                              Chr(13) & pMensagem)
End Function

Private Sub TabInspecao_Click(PreviousTab As Integer)
    If PreviousTab = 1 Then
        '        flxCotacoes.Visible = True
        cmdSolicitar.Visible = True
        '        WebBrowser.Visible = False
    Else
        '        flxCotacoes.Visible = False
        cmdSolicitar.Visible = False
        '        WebBrowser.Visible = True
        WebBrowser.Navigate2 ("http://www.yasuda.com.br/InspecaoDiversos/Controller?Syas=1&numSusep=" & gM46V999.gNumSusep & "&login=" & gM46V999.gNomUsuario & "&action=interno&dataInicio=01/01/2010&dataFim=" & Format(Date, "dd/mm/yyyy") & "&codSituacao=00")
    End If

End Sub

