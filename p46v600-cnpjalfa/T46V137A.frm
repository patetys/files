VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V137A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fatores de Proteção"
   ClientHeight    =   4560
   ClientLeft      =   1545
   ClientTop       =   2145
   ClientWidth     =   8985
   ControlBox      =   0   'False
   Icon            =   "T46V137A.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picChecked 
      Height          =   255
      Left            =   90
      Picture         =   "T46V137A.frx":000C
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4275
      Width           =   255
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   300
      Left            =   7830
      TabIndex        =   0
      ToolTipText     =   "Volta ao item"
      Top             =   4275
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin MSFlexGridLib.MSFlexGrid flxProtecao 
      Height          =   4170
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   7355
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483648
      ForeColor       =   8388608
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   16777215
      ForeColorSel    =   12648447
      BackColorBkg    =   -2147483633
      GridColorFixed  =   4210752
      AllowBigSelection=   0   'False
      GridLines       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
   End
   Begin VB.Label Label1 
      Caption         =   "Selecione os fatores de Proteção clicando na primeira coluna"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   4320
      Width           =   5325
   End
End
Attribute VB_Name = "frmT46V137A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ColecaoFator As Collection    'Coleção.
Dim mLinhaFocus As Long    'Linha que recebeu o foco.
Dim mColunaFocus As Long    'Coluna que recebeu o foco.

Private Sub cmdOk_Click()
    Dim lI As Integer    'Utilizado no for...loop.
    Dim lLinha As Integer    'Linha.

'1. Efetuo as consistências
'For lI = 1 To flxProtecao.Rows - 1
'Next lI

'3. Inibir a tela
    On Error Resume Next
    Me.Hide
    On Error GoTo 0

End Sub

Private Sub flxProtecao_Click()
    Dim lColuna As Integer
    Dim lLinha As Integer

    If flxProtecao.Enabled = False Then Exit Sub

    With flxProtecao
        If .Row = 0 Then Exit Sub
        '        .Redraw = False
        lColuna = .Col
        lLinha = .Row
        If flxProtecao.Col = 0 Then
            'Ficha 901673
            If flxProtecao.CellPicture = picChecked Then
                'não permite que desmarque o Desconto Especial
                If (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor) And (UCase(flxProtecao.TextMatrix(lLinha, 2)) Like "*DESCONTO ESPECIAL*") Then
                    '1101045 - Tarifa 1022 e 1023 Novembro 2011
                    '                If (UCase(flxProtecao.TextMatrix(lLinha, 2)) Like "*DESCONTO ESPECIAL*") Then
                    Exit Sub
                End If

                'Supermercado SMART
                If gM46V999.gCod_Ramo = 112 And gM46V999.gCodProduto >= 1042 And (frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) >= 2178 And frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) <= 2182) Then
                    Exit Sub
                End If

                'Ficha 1200806

                If (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor) And InStr(frmT46V101A.objstcPedLoc.gTemProtecaoObrigatoria, flxProtecao.TextMatrix(lLinha, 1) & "@") > 0 Then
                    Exit Sub
                End If

                'Ficha 1000304
                If UCase(flxProtecao.TextMatrix(lLinha, 2)) Like "*RISCO EM SHOPPING CENTER*" Then
                    'habilitar os demais p/ poder marcar
                    gM46V999.gAchouShopping = False
                End If
                Set flxProtecao.CellPicture = LoadPicture("")
                .TextMatrix(lLinha, 4) = ""
                .TextMatrix(lLinha, 3) = ""
                .Col = 3
                .CellBackColor = &H80000000
            Else
                'Ficha 1000304
                If UCase(flxProtecao.TextMatrix(lLinha, 2)) Like "*RISCO EM SHOPPING CENTER*" Then
                    gM46V999.gAchouShopping = True
                Else
                    gM46V999.gAchouShopping = False
                End If

                Dim i As Integer
                i = 1
                
                If gM46V999.gCod_Ramo = 112 And gM46V999.gCodProduto >= 1080 And (gM46V111.gCod_plano < 1000 Or frmT46V101A.objstcPedido.getPlano.EhPlanoSegmentado) Then
                
                    If .TextMatrix(lLinha, 1) < 5 Then
                        Do While i < 5
                            If i <> lLinha And .TextMatrix(i, 4) = "Marcado" Then
                                .Row = i
                                Set flxProtecao.CellPicture = LoadPicture("")
                                .TextMatrix(i, 3) = ""
                                .TextMatrix(i, 4) = ""
                                .Col = 3
                                .CellBackColor = &H80000000
                                Exit Do
                            End If
                            i = i + 1
                        Loop
                    End If

                End If
                
                .Col = 0
                .Row = lLinha
                Set flxProtecao.CellPicture = picChecked
                .TextMatrix(lLinha, 4) = "Marcado"
                .CellPictureAlignment = flexAlignCenterCenter
                .Col = 3
                .CellBackColor = &H80000018
            End If
        End If

Pula_Marcar:
        .Col = lColuna
        .Row = lLinha
        '        .Redraw = True
        .Refresh
    End With

    'Ficha 1000304
    If gM46V999.gAchouShopping = True Then
        Dim lLin As Integer

        For lLin = 1 To flxProtecao.Rows - 1
            flxProtecao.Col = 0
            flxProtecao.Row = lLin
            If gM46V999.gCod_Ramo = e_CodRamo_Empresarial And flxProtecao.TextMatrix(lLin, 1) <> 9 Then
                Set flxProtecao.CellPicture = LoadPicture("")
                flxProtecao.TextMatrix(lLin, 4) = ""
                flxProtecao.Col = 3
                flxProtecao.CellBackColor = &H80000000
            End If
        Next lLin

        flxProtecao.Refresh
    End If

    'Desabilita prêmios.
    Call gpLimPremios
End Sub

Private Sub flxProtecao_KeyDown(KeyCode As Integer, Shift As Integer)

    If flxProtecao.Enabled = False Then Exit Sub

    With flxProtecao
        If .RowSel > 0 Then
            Select Case KeyCode
            Case 46
                If .ColSel = 3 Then
                    .Row = .RowSel
                    .Col = .ColSel
                    .Refresh
                    .TextMatrix(.RowSel, .ColSel) = ""
                End If
            Case 32
                If .ColSel = 0 Then
                    .Row = .RowSel
                    .Col = .ColSel
                    flxProtecao_Click
                End If
            End Select
        End If
    End With
    Call gpHabSalvar

End Sub

Private Sub flxProtecao_KeyPress(KeyAscii As Integer)

    If flxProtecao.Enabled = False Then Exit Sub

    With flxProtecao
        '1101045 - Tarifa 1022 e 1023 Novembro 2011
        If (.RowSel > 0) And (.ColSel = 3) And (.TextMatrix(.RowSel, 4) = "Marcado") And _
           (Not (UCase(.TextMatrix(.RowSel, 2)) Like "*DESCONTO ESPECIAL*")) Then
            Select Case KeyAscii
            Case 8, 44, 48 To 57
                If mLinhaFocus = .RowSel Then
                    mLinhaFocus = 0
                    mColunaFocus = 0
                    If .TextMatrix(.RowSel, .ColSel) = "" Then
                        'Call cmdExcluir_Click
                        Exit Sub
                    End If
                    .TextMatrix(.RowSel, .ColSel) = ""
                    .Row = .RowSel
                    .Col = .ColSel
                    .Refresh
                End If
            End Select
            Select Case KeyAscii
            Case 8              'BackSpace
                If Len(.TextMatrix(.RowSel, .ColSel)) > 0 Then
                    .TextMatrix(.RowSel, .ColSel) = Left(.TextMatrix(.RowSel, .ColSel), Len(.TextMatrix(.RowSel, .ColSel)) - 1)
                End If
            Case 13             'Enter
                If .RowSel < .Rows - 1 Then
                    .RowSel = .RowSel + 1
                    .Row = .RowSel
                End If
            Case 44             'Vírgula
                If InStr(1, .TextMatrix(.RowSel, .ColSel), Chr(44)) > 0 Then Exit Sub
                .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
            Case 48 To 57
                .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
            Case 32, 33, 35 To 38, 40 To 47, 61, 65 To 93, 95, 97 To 123, 126
                If Len(.TextMatrix(.RowSel, .ColSel)) < 30 Then
                    .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
                Else
                    Beep
                End If
            End Select
            Call gpHabSalvar
        End If
    End With

End Sub

Private Sub flxProtecao_RowColChange()
    flxProtecao_SelChange
End Sub

Private Sub flxProtecao_SelChange()
    flxProtecao.ColWidth(4) = 0
End Sub

Private Sub Form_Activate()
'1. Posiciona nome do form.
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)

    If flxProtecao.Rows = 1 Then
        'Call mpCarregarGrid
        '1100705 - Tarifa 1016 e 1017 Agosto 2011 - Item 7
        Call mpSelecionarFatores(frmT46V101A.objstcPedLoc.ColecaoPedFator)
    End If
    flxProtecao.Col = 0
    If flxProtecao.Rows > 1 Then
        flxProtecao.Row = 1
    End If
    If flxProtecao.Enabled = True Then
        flxProtecao.SetFocus
    End If

Finalizar:
    If gM46V999.gDesabilitarForms Or gM46V999.gDesabFormsCotacao Then
        If Me.Tag <> "Desabilitado" Then
            DesabilitaControles Me
            cmdOk.Enabled = True
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
        End If
    End If
End Sub

Private Sub Form_Load()
'1. Posiciona form.
    Call gM46V999.gpCenForm(Me)
    With flxProtecao
        .Rows = 1
        .Cols = 5
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Fator"
        .TextMatrix(0, 2) = "Descrição do Fator"
        .TextMatrix(0, 3) = "Complemento do Fator"
        .TextMatrix(0, 4) = "?"
        .Row = 0
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        .ColWidth(2) = 5000
        .ColWidth(3) = 3000
        .ColWidth(4) = 0
    End With

End Sub


Public Sub mpSelecionarFatores(pColecaoPedFator As Collection)
    Dim lobjstcFator As stcA46V720B    'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lobjstc As stcA46V101B    'Objeto da classe de estrutura stcA46V101A.
    Dim lblnAchou As Boolean    'Indica se achou  o fator ou não.
    Dim lblnAchouDesc As Boolean     'Indica se achou o fator "DESCONTO ESPECIAL" ou não
    Dim lLinha As Long    'Linhas.

    If frmT46V101A.cboPlano.ListIndex = -1 Then Exit Sub


    Call mpCarregarGrid
    flxProtecao.Rows = 1
    lLinha = 1
    For Each lobjstc In ColecaoFator
        lblnAchou = False
        lblnAchouDesc = False

        For Each lobjstcFator In pColecaoPedFator
            If lobjstcFator.Cod_Fator = lobjstc.Codigo Then
                lblnAchou = True
                Exit For
            End If
        Next lobjstcFator

        With flxProtecao
            .Rows = lLinha + 1
            .Row = lLinha

            'Ficha 901673
            If (UCase(lobjstc.Descricao) Like "*DESCONTO ESPECIAL*") And (frmT46V101A.objstcPedLoc.gLiberaDescontoEspecial = True) Then
                lblnAchou = True
                lblnAchouDesc = True
            End If

            If lblnAchou Then
                .Col = 0
                .CellPictureAlignment = flexAlignCenterCenter
                Set flxProtecao.CellPicture = picChecked
                'Ficha 1000304
                If UCase(lobjstc.Descricao) Like "*RISCO EM SHOPPING CENTER*" Then
                    gM46V999.gAchouShopping = True
                Else
                    gM46V999.gAchouShopping = False
                End If
            End If

            .TextMatrix(lLinha, 0) = ""
            .TextMatrix(lLinha, 1) = lobjstc.Codigo
            .TextMatrix(lLinha, 2) = lobjstc.Descricao

            If Not lblnAchou Then
                .TextMatrix(lLinha, 3) = ""
                .TextMatrix(lLinha, 4) = ""
            Else
                .Col = 3
                .CellBackColor = &H80000018
                If lblnAchouDesc = True Then
                    .TextMatrix(lLinha, 3) = " "
                Else
                    .TextMatrix(lLinha, 3) = lobjstcFator.Dsc_Compl_Fator
                End If
                .TextMatrix(lLinha, 4) = "Marcado"

            End If
        End With
        lLinha = lLinha + 1
    Next lobjstc
    flxProtecao.Col = 0
    If flxProtecao.Rows > 1 Then
        flxProtecao.Row = 1
    End If

    For Each lobjstc In ColecaoFator
        'Ficha 901673
        If gM46V999.gCod_Ramo = e_CodRamo_Residencial And _
           (lobjstc.Codigo = 7 Or _
            UCase(lobjstc.Descricao) Like "*DESCONTO ESPECIAL*") Then
            lblnAchou = True
            Exit For
        End If

    Next lobjstc

End Sub

Public Sub mpCarregarGrid()
    Dim lobjstc As stcA46V101B    'Objeto da classe de estrutura stcA46V101B.
    Dim lbdTbm_Tipo_Fator As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lMensagem As String     'Mensagem.
    Dim lrsTbm_Tipo_Fator As clsYasRecordSet    'Registro : Tbm_Tipo_Fator.
    Dim lLinha As Long    'Linhas.
    Dim lSelect As String
    Dim lCodSequencia As Integer

    Set ColecaoFator = New Collection
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Tipo_Fator, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If frmT46V101A.cboPlano.ListIndex = -1 Then
        Exit Sub
    End If

    lCodSequencia = gM46V999.gCodSequencia

NovaSequencia:
    'Ficha 1000601
    If gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
        If (frmT46V101A.objstcPedLoc.gLiberaDescontoEspecial = True) Then
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & gM46V999.gCod_Ramo
            'Selecionar apenas os fatores da modalidade(plano)
            lSelect = lSelect & " and Cod_Modalidade = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
            lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto
            lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia
            lSelect = lSelect & " order by DSC_FATOR"
        Else
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & gM46V999.gCod_Ramo
            'Selecionar apenas os fatores da modalidade(plano)
            lSelect = lSelect & " and Cod_Modalidade = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
            lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto
            lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia
            lSelect = lSelect & " and Cod_Fator <> 7 "    'Não carregar o item Desconto Especial
            lSelect = lSelect & " order by DSC_FATOR"
        End If
    Else
        lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
        lSelect = lSelect & " Where cod_Ramo = " & gM46V999.gCod_Ramo
        'Selecionar apenas os fatores da modalidade(plano)
        lSelect = lSelect & " and Cod_Modalidade = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
        lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto
        lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia
        lSelect = lSelect & " order by COD_FATOR"
    End If

    If gM46V999.gfObtRegistro(lbdTbm_Tipo_Fator, lSelect, lrsTbm_Tipo_Fator, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTbm_Tipo_Fator)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If lrsTbm_Tipo_Fator.EOF Then
        If lCodSequencia > 0 Then
            lCodSequencia = lCodSequencia - 1
            GoTo NovaSequencia
        End If
    End If

    Set ColecaoFator = New Collection
    Do While Not lrsTbm_Tipo_Fator.EOF
        Set lobjstc = New stcA46V101B
        lobjstc.Descricao = Trim(lrsTbm_Tipo_Fator("DSC_FATOR"))
        lobjstc.Codigo = lrsTbm_Tipo_Fator("COD_FATOR")
        ColecaoFator.Add lobjstc
        lrsTbm_Tipo_Fator.MoveNext
    Loop


    lLinha = 1
    flxProtecao.Rows = 1
    flxProtecao.Cols = 5
    For Each lobjstc In ColecaoFator
        flxProtecao.Rows = lLinha + 1
        flxProtecao.TextMatrix(lLinha, 0) = ""
        flxProtecao.TextMatrix(lLinha, 1) = lobjstc.Codigo
        flxProtecao.TextMatrix(lLinha, 2) = lobjstc.Descricao
        flxProtecao.TextMatrix(lLinha, 3) = ""
        flxProtecao.TextMatrix(lLinha, 4) = ""
        'Ficha 901673
        'permite que a digitação desmarque o Desconto Especial
        'Correção Ficha - 1000601 - Permitir Somente desconto nas tarifas 998 até 993

        '1101045 - Tarifa 1022 e 1023 Novembro 2011
        '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 8
        'Retirado na ficha 1200733
        '        If ((frmT46V101A.objstcPedido.Cod_Produto >= 1022 And frmT46V101A.objstcPedido.Cod_Produto <= 1033) Or _
                 '            (frmT46V101A.objstcPedido.Cod_Produto >= 1038 And frmT46V101A.objstcPedido.Cod_Produto <= 1043))
        If (frmT46V101A.objstcPedLoc.gLiberaDescontoEspecial = True) And _
           (UCase(lobjstc.Descricao) Like "*DESCONTO ESPECIAL*") Then
            flxProtecao.Row = lLinha
            flxProtecao.TextMatrix(lLinha, 4) = "Marcado"
            Set flxProtecao.CellPicture = picChecked
            flxProtecao.Col = 3
            flxProtecao.CellBackColor = &H80000018
        End If
        lLinha = lLinha + 1
    Next lobjstc
    Call gM46V999.gpFecha2(lbdTbm_Tipo_Fator)
End Sub

