VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V111A 
   Caption         =   "Selecione a classificação."
   ClientHeight    =   3630
   ClientLeft      =   2115
   ClientTop       =   2010
   ClientWidth     =   9930
   ControlBox      =   0   'False
   Icon            =   "T46V111A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   9930
   Begin VB.TextBox txtClassificacao 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1200
      MaxLength       =   130
      TabIndex        =   3
      ToolTipText     =   "Informe os dados para a busca."
      Top             =   0
      Width           =   8700
   End
   Begin VB.ListBox lstClassificacao 
      ForeColor       =   &H00800000&
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   9900
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   4455
      TabIndex        =   1
      ToolTipText     =   "Volta ao item com a classificação selecionada."
      Top             =   3240
      Width           =   1020
      _Version        =   65536
      _ExtentX        =   1799
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Busca: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Ordem da listagem da região."
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmT46V111A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ColecaoClassificacao As Collection   'Coleção.

Private Sub cmdOK_Click()
    Dim lobjstc As stcA46V101B                     'Objeto da classe de estrutura stcA46V101A.

    If lstClassificacao.ListIndex = -1 Then
        If lstClassificacao.ListCount > 0 Then
            lstClassificacao.ListIndex = 0
        End If
    End If
    If lstClassificacao.ListIndex <> -1 Then
        Set lobjstc = ColecaoClassificacao.Item(lstClassificacao.List(lstClassificacao.ListIndex))
        frmT46V101A.txtClassificacao.Text = lobjstc.Codigo & "-" & lobjstc.Descricao
        frmT46V101A.txtClassificacao.Tag = lobjstc.Chave
        frmT46V101A.objstcPedLoc.Cod_Clasf = Left(frmT46V101A.txtClassificacao.Tag, 4)
        frmT46V101A.objstcPedLoc.Cod_Compl_Clasf = Mid(frmT46V101A.txtClassificacao.Tag, 5, 2)
    End If

    If frmT46V101A.cboPlano.ListIndex <> -1 Then
        If frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) >= 2001 And _
           frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) <= 2031 Then
            frmT46V101A.txtDescBonus.Text = 0
            gM46V111.gDesc_Bonus = 0
        End If

        '1000744 - Item 4 - Plano Mercedes
        If (frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) >= 2021 And _
            frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) <= 2031) And (gM46V999.gCodProduto < 992) Then
            If Left(frmT46V101A.txtClassificacao.Text, 7) = "001 -11" Then
                gM46V111.gDesc_Coml = 33
                frmT46V101A.txtDescComercial.Text = 0
                frmT46V117A.txtDescComercial.Text = 0
            End If
            If Left(frmT46V101A.txtClassificacao.Text, 7) = "001 -21" Then
                gM46V111.gDesc_Coml = 20
                frmT46V101A.txtDescComercial.Text = 0
                frmT46V117A.txtDescComercial.Text = 0
            End If
            Dim ColecaoPedCobTemp As Collection
            Dim lobjstcCobert As stcA46V704B

            Set ColecaoPedCobTemp = New Collection
            Set lobjstcCobert = New stcA46V704B
            Set lobjstcCobert.M46V999 = gM46V999

            For Each lobjstcCobert In frmT46V101A.objstcPedLoc.ColecaoPedCobert
                lobjstcCobert.Per_Desc_Coml = frmT46V101A.txtDescComercial.Text
                ColecaoPedCobTemp.Add lobjstcCobert, lobjstcCobert.Class_ID
            Next lobjstcCobert

            Set frmT46V101A.objstcPedLoc.ColecaoPedCobert = Nothing
            Set frmT46V101A.objstcPedLoc.ColecaoPedCobert = ColecaoPedCobTemp
        End If

        'Ficha 901673 e 901772
        '1200903 - Tarifa - 1046 e 1047 - Item 13
        If (frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) >= 2051 And _
            frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) <= 2060 And frmT46V101A.objstcPedido.Cod_Produto < 1046) Or _
            (frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) >= 2111 And _
             frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) <= 2130) Then
            If Left(frmT46V101A.txtClassificacao.Text, 7) = "001 -11" Then
                gM46V111.gDesc_Coml = 20
                frmT46V101A.txtDescComercial.Text = 0
                frmT46V117A.txtDescComercial.Text = 0
            End If
            If Left(frmT46V101A.txtClassificacao.Text, 7) = "001 -21" Then
                gM46V111.gDesc_Coml = 5
                frmT46V101A.txtDescComercial.Text = 0
                frmT46V117A.txtDescComercial.Text = 0
            End If
            Dim ColecaoPedCobTemp2 As Collection
            Dim lobjstcCobert2 As stcA46V704B

            Set ColecaoPedCobTemp2 = New Collection
            Set lobjstcCobert2 = New stcA46V704B
            Set lobjstcCobert2.M46V999 = gM46V999

            For Each lobjstcCobert2 In frmT46V101A.objstcPedLoc.ColecaoPedCobert
                lobjstcCobert2.Per_Desc_Coml = frmT46V101A.txtDescComercial.Text
                ColecaoPedCobTemp2.Add lobjstcCobert2, lobjstcCobert2.Class_ID
            Next lobjstcCobert2

            Set frmT46V101A.objstcPedLoc.ColecaoPedCobert = Nothing
            Set frmT46V101A.objstcPedLoc.ColecaoPedCobert = ColecaoPedCobTemp2
        End If

        'Fichas 901282 e 901346 e 901673 e 901772
        If (frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) >= 2061 And _
            frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) <= 2100) Or _
            (frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) >= 2111 And _
             frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) <= 2130) Then
            If Left(frmT46V101A.txtClassificacao.Text, 7) = "001 -11" Then
                gM46V111.gDesc_Coml = 20
                frmT46V101A.txtDescComercial.Text = 0
                frmT46V117A.txtDescComercial.Text = 0
            End If
            If Left(frmT46V101A.txtClassificacao.Text, 7) = "001 -21" Then
                gM46V111.gDesc_Coml = 5
                frmT46V101A.txtDescComercial.Text = 0
                frmT46V117A.txtDescComercial.Text = 0
            End If
            Dim ColecaoPedCobTemp3 As Collection
            Dim lobjstcCobert3 As stcA46V704B

            Set ColecaoPedCobTemp3 = New Collection
            Set lobjstcCobert3 = New stcA46V704B
            Set lobjstcCobert3.M46V999 = gM46V999

            For Each lobjstcCobert3 In frmT46V101A.objstcPedLoc.ColecaoPedCobert
                lobjstcCobert3.Per_Desc_Coml = frmT46V101A.txtDescComercial.Text
                ColecaoPedCobTemp3.Add lobjstcCobert3, lobjstcCobert3.Class_ID
            Next lobjstcCobert3

            Set frmT46V101A.objstcPedLoc.ColecaoPedCobert = Nothing
            Set frmT46V101A.objstcPedLoc.ColecaoPedCobert = ColecaoPedCobTemp3
        End If

    End If
    If Not gCarga Then
        gAltDadosCobert = True
        Unload frmT46V117A
        Unload frmT46V179A
        gM46V999.gChamaSelecionarCoberturas = True
        'Call frmT46V117A.mpSelecionarCoberturas(frmT46V101A.objstcPedLoc.ColecaoPedCobert)
    End If
    Unload Me
End Sub

Private Sub Form_Load()
'1. Posiciona form.
    Call gM46V999.gpCenForm(Me)

    Call mpCarregarList
End Sub
Private Sub mpCarregarList()
    Dim lobjstc As stcA46V101B  'Objeto da classe de estrutura stcA46V101A.
    Dim lbdTAB_Class As ADODB.Connection  'Arquivo P00Tarifa.
    Dim lMensagem As String         'Mensagem.
    Dim rsTab_Class As clsYasRecordSet   'Registro, TAB_CLASF
    Dim lCount As Long        'Contador.
    Dim lSelect As String

    Set ColecaoClassificacao = New Collection
    If gM46V999.gCodProduto <= gM46V999.TarifaAntiga Then
        If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Class, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
        lSelect = " Select distinct dsc_clasf, COD_CLASF, " & _
                  "COD_COMPL_CLASF , tab_Clasf from TAB_CLASF " & _
                " Where cod_Ramo  = " & gM46V999.gCod_Ramo
    Else
        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Class, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
        lSelect = " Select distinct dsc_clasf, COD_CLASF, " & _
                  "COD_COMPL_CLASF , tab_Clasf from Tbm_Classificacao " & _
                " Where cod_Ramo  = " & gM46V999.gCod_Ramo
    End If

    If frmT46V101A.cboPlano.ListIndex <> -1 Then
        Select Case frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
        Case 1020, 1030, 1040, 1050, 1060
            lSelect = lSelect & " and  tab_clasf = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
        Case Else
            'simony
            If gM46V999.gCodProduto <= gM46V999.TarifaAntiga Then
                If Val(frmT46V101A.txtRisRegiao.Tag) <> 0 Then
                    lSelect = lSelect & " and  tab_clasf = " & Right(frmT46V101A.txtRisRegiao.Tag, 4)
                Else
                    lSelect = lSelect & " and  tab_clasf = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
                End If
            Else
                lSelect = lSelect & " and  tab_clasf = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
            End If
        End Select
    End If

    If gM46V999.gCodProduto > gM46V999.TarifaAntiga Then
        lSelect = lSelect & " and  cod_tarif  =  " & gM46V999.gCodProduto
        lSelect = lSelect & " and  cod_sequencia  <=  " & gM46V999.gCodSequencia
    End If

    If gM46V999.gfObtRegistro(lbdTAB_Class, lSelect, rsTab_Class, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_Class)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Do While Not rsTab_Class.EOF
        Set lobjstc = New stcA46V101B
        lobjstc.Descricao = Trim(rsTab_Class("DSC_CLASF"))
        lobjstc.Codigo = rsTab_Class("COD_CLASF") & "-" & Format(rsTab_Class("COD_COMPL_CLASF"), "00")
        lobjstc.Chave = rsTab_Class("COD_CLASF") & String(4 - Len(rsTab_Class("COD_CLASF")), " ") & Format(rsTab_Class("COD_COMPL_CLASF"), "00") & _
                        Format(rsTab_Class("tab_Clasf"), "0000")
        ColecaoClassificacao.Add lobjstc, lobjstc.Codigo & " - " & lobjstc.Descricao
        rsTab_Class.MoveNext
    Loop
    lCount = 0
    For Each lobjstc In ColecaoClassificacao
        lCount = lCount + 1
        lstClassificacao.AddItem lobjstc.Codigo & " - " & lobjstc.Descricao
        lstClassificacao.ItemData(lstClassificacao.NewIndex) = lCount
    Next lobjstc
End Sub

Private Sub Form_Activate()
'1. Posiciona controles.
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    If lstClassificacao.ListCount = 0 Then
        Call mpCarregarList
    End If
    'Para o planinho cacau show selecionar direto a classificação pois só tem uma.
    If frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) >= 2220 And frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) <= 2227 Then
        Call cmdOK_Click
    End If
End Sub

Private Sub Form_Resize()

    On Error GoTo Erro
    
    If Me.Height > 1500 And Me.Width > 1500 Then
        txtClassificacao.Width = Me.Width - (txtClassificacao.Left + 270)
        With lstClassificacao
            .Left = 30
            .Top = 40 + txtClassificacao.Height
            .Width = Me.Width - 270
            .Height = Me.Height - 1500
        End With
        
        cmdOk.Top = Me.Height - 1000
        cmdOk.Left = CInt(Me.Width / 2) - CInt(cmdOk.Width / 2)
    End If
    
    Exit Sub
Erro:
    Call gM46V999.gpGraLog(0, Me.Name + "_Form_Resize ERRO - " & Err & " " & Error)
End Sub

Private Sub lstClassificacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOK_Click
    End If
End Sub

Private Sub txtClassificacao_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lobjstc As stcA46V101B      'Objeto da classe de estrutura stcA46V101A.

    lstClassificacao.Clear
    For Each lobjstc In ColecaoClassificacao
        If InStr(1, lobjstc.Codigo & " - " & lobjstc.Descricao, UCase(txtClassificacao.Text)) > 0 Then
            lstClassificacao.AddItem lobjstc.Codigo & " - " & lobjstc.Descricao
        End If
    Next lobjstc
End Sub

Private Sub lstClassificacao_DblClick()
    Call cmdOK_Click
End Sub
