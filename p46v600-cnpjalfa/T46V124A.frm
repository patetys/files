VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmT46V124A 
   Caption         =   "Demonstrativo do parcelamento"
   ClientHeight    =   4245
   ClientLeft      =   4425
   ClientTop       =   1800
   ClientWidth     =   12030
   ControlBox      =   0   'False
   Icon            =   "T46V124A.frx":0000
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "T46V124A.frx":000C
   ScaleHeight     =   4245
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picUnchecked 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1890
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   3630
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picChecked 
      Height          =   255
      Left            =   1530
      Picture         =   "T46V124A.frx":031E
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3630
      Visible         =   0   'False
      Width           =   285
   End
   Begin TabDlg.SSTab tabResultado 
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   10160
      _ExtentX        =   17912
      _ExtentY        =   5900
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   706
      BackColor       =   12632256
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "1ª parcela a &vista"
      TabPicture(0)   =   "T46V124A.frx":0418
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flxVistaDebConta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "flxVistaCarne"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "1ª parcela a &30 dias"
      TabPicture(1)   =   "T46V124A.frx":0434
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxCartao"
      Tab(1).Control(1)=   "flx30diasDebConta"
      Tab(1).Control(2)=   "flx30diasCarne"
      Tab(1).ControlCount=   3
      Begin MSFlexGridLib.MSFlexGrid flxVistaCarne 
         Height          =   2865
         Left            =   0
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5054
         _Version        =   393216
         Rows            =   13
         Cols            =   5
         ForeColor       =   12582912
         ForeColorFixed  =   12582912
         BackColorSel    =   12582912
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483633
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid flxVistaDebConta 
         Height          =   2865
         Left            =   3330
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5054
         _Version        =   393216
         Rows            =   13
         Cols            =   5
         ForeColor       =   12582912
         ForeColorFixed  =   12582912
         BackColorSel    =   12582912
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483633
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid flx30diasCarne 
         Height          =   2865
         Left            =   -74970
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5054
         _Version        =   393216
         Rows            =   13
         Cols            =   5
         ForeColor       =   12582912
         ForeColorFixed  =   12582912
         BackColorSel    =   12582912
         ForeColorSel    =   16777215
         BackColorBkg    =   12632256
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid flx30diasDebConta 
         Height          =   2865
         Left            =   -71610
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   5054
         _Version        =   393216
         Rows            =   13
         Cols            =   5
         ForeColor       =   12582912
         ForeColorFixed  =   12582912
         BackColorSel    =   12582912
         ForeColorSel    =   16777215
         BackColorBkg    =   12632256
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid flxCartao 
         Height          =   2865
         Left            =   -68290
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5054
         _Version        =   393216
         Rows            =   13
         Cols            =   5
         ForeColor       =   12582912
         ForeColorFixed  =   12582912
         BackColorSel    =   12582912
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483633
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
   End
   Begin Threed.SSCommand cmdVoltar 
      Height          =   300
      Left            =   1515
      TabIndex        =   1
      ToolTipText     =   "Volta ao item."
      Top             =   3780
      Width           =   1800
      _Version        =   65536
      _ExtentX        =   3175
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Voltar"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdMosPremios 
      Height          =   300
      Left            =   3375
      TabIndex        =   2
      ToolTipText     =   "Volta ao item."
      Top             =   3780
      Width           =   1905
      _Version        =   65536
      _ExtentX        =   3360
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "Mostrar &prêmios do item"
      ForeColor       =   8388608
   End
   Begin VB.Label lblCartao 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cartão de Crédito"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   7070
      TabIndex        =   12
      Top             =   0
      Width           =   3105
   End
   Begin VB.Label lblDebConta 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Debito em conta"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3600
      TabIndex        =   8
      Top             =   0
      Width           =   3105
   End
   Begin VB.Label lblCarne 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Carnê"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   3345
   End
End
Attribute VB_Name = "frmT46V124A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjstcPed As stcA46V702B
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public UltimoErro As String           ' Descrição do último erro gerado

Public Sub gpInstanciarForm(pobjstcPed As stcA46V702B)
    Set mobjstcPed = pobjstcPed
    Load Me
End Sub

Private Function mfParcelaSelecionada() As Boolean
    Dim lI As Integer
    Dim lFra30diaSelecionada As Boolean
    Dim lFraVistaSelecionada As Boolean
    Dim lFra30diaDebContaSelecionada As Boolean
    Dim lFraVistaDebContaSelecionada As Boolean
    Dim lFraCartaoSelecionada As Boolean

    mfParcelaSelecionada = False
    lFra30diaDebContaSelecionada = False
    lFraVistaDebContaSelecionada = False
    lFraCartaoSelecionada = False

    With flxVistaDebConta
        For lI = 1 To .Rows - 1
            .Col = 4
            .Row = lI

            If .CellPicture = picChecked Then
                lFraVistaDebContaSelecionada = True
                Exit For
            End If
        Next
    End With

    With flx30diasDebConta
        For lI = 1 To .Rows - 1
            .Col = 4
            .Row = lI

            If .CellPicture = picChecked Then
                lFra30diaDebContaSelecionada = True
                Exit For
            End If
        Next
    End With
    
    If frmT46V101A.objstcPedido.Cod_Bco_Cobr = 999 And _
       frmT46V101A.objstcPedido.Cod_Agen_Bas_Cor = 9999 And _
       frmT46V101A.objstcPedido.Num_Cta_Bas_Cor = 9999999 Then
        With flxCartao
            For lI = 1 To .Rows - 1
                .Col = 4
                .row = lI
    
                If .CellPicture = picChecked Then
                    lFraCartaoSelecionada = True
                    Exit For
                End If
            Next
        End With
    End If

    If lFra30diaSelecionada = True _
       Or lFraVistaSelecionada = True _
       Or lFra30diaDebContaSelecionada = True _
       Or lFraVistaDebContaSelecionada = True _
       Or lFraCartaoSelecionada = True Then
        mfParcelaSelecionada = True
    End If
End Function

Private Sub mpMonGraFracionamento(ByVal pGrade As MSFlexGrid, pColunaImprimir As Boolean)
'Procedure: monta grade de fracionamento.

    With pGrade
        .Rows = 1
        .Row = 0
        .Cols = IIf(pColunaImprimir = True, 5, 4)
        'a) Posiciona fonte.
        .Font = 5
        .FontWidth = 3.35
        'b) Posiciona largura das colunas.
        .ColWidth(0) = 255
        .ColWidth(1) = 900    '1100
        .ColWidth(2) = 900    '1100
        .ColWidth(3) = 500

        'c) Posiciona altura da linha de cabeçalho.
        .RowHeight(0) = 220
        'd) Posiciona alinhamento das colunas.
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        'e) Posiciona títulos.
        .TextMatrix(0, 1) = "1ª"
        .TextMatrix(0, 2) = "Demais"
        .TextMatrix(0, 3) = "Juros"

        If pColunaImprimir = True Then
            .ColWidth(4) = 650
            .ColAlignment(4) = flexAlignCenterCenter
            .TextMatrix(0, 4) = "Imprimir"
        End If

    End With
End Sub
Private Sub mpPosValor(ByVal pGrade As MSFlexGrid, ByVal pLinha As Byte, ByVal pValParcela As Double, _
                       ByVal pValDemais As Double, ByVal pJuros As Double, ByVal pSelecionado As Boolean)
'Procedure: posiciona valor na grade.

'Parâmetros: pGrade........grade utilizada.
'            pLinha........número da linha.
'            pValParcela...valor da primeira parcela.
'            pValDemais....valor das demais parcelas.
'            pJuros........percentual de juros.

    With pGrade
        .Rows = pLinha + 1
        .TextMatrix(pLinha, 0) = pLinha  'Número de parcelas.
        .TextMatrix(pLinha, 1) = Format$(pValParcela, "standard")    'Primeira parcela.
        .TextMatrix(pLinha, 2) = Format$(pValDemais, "standard")    'Demais parcelas.
        .TextMatrix(pLinha, 3) = Format$(pJuros, "standard")  'Juros.

        If pGrade.Cols > 4 Then
            If pSelecionado = True Then    'Imprimir.
                .Row = pLinha: .Col = 4: Set .CellPicture = picChecked
            Else
                .Row = pLinha: .Col = 4: Set .CellPicture = Nothing
            End If
        End If
    End With
End Sub

Private Sub cmdMosPremios_Click()
    Me.Hide
    frmT46V123A.Show vbModeless, frmT46V101A
    If frmT46V123A.tabPremios.Visible = True And frmT46V123A.tabPremios.Enabled = True Then
        frmT46V123A.tabPremios.SetFocus
    End If
End Sub

Private Sub mpSalvar()
    Dim lI As Integer
    Dim lParcelas_Impressao_Avista As String
    Dim lParcelas_Impressao_30dias As String
    Dim lParcelas_Impressao_Cartao As String
    Dim lsSQl As String                     'Query.
    Dim lobjCtrlEmis As clsA46V712A    'Objeto da classe clsA46V712A (TAB_CTRL_EMIS)
    Dim lConexaoBD As ADODB.Connection

    If mfParcelaSelecionada = False Then
        Call gM46V999.gpGraLog(1, "Escolha ao menos uma opção de parcelamento.")
    Else
        lParcelas_Impressao_Avista = ""
        lParcelas_Impressao_30dias = ""
        lParcelas_Impressao_Cartao = ""

        With flxVistaDebConta
            If .Rows > 1 Then
                For lI = 1 To .Rows - 1
                    .Row = lI

                    If .CellPicture = frmT46V124A.picChecked Then
                        If lParcelas_Impressao_Avista <> "" Then lParcelas_Impressao_Avista = lParcelas_Impressao_Avista & ","
                        lParcelas_Impressao_Avista = lParcelas_Impressao_Avista & .TextMatrix(lI, 0)
                    End If
                Next
            End If
        End With

        With flx30diasDebConta
            If .Rows > 1 Then
                For lI = 1 To .Rows - 1
                    .Row = lI

                    If .CellPicture = frmT46V124A.picChecked Then
                        If lParcelas_Impressao_30dias <> "" Then lParcelas_Impressao_30dias = lParcelas_Impressao_30dias & ","
                        lParcelas_Impressao_30dias = lParcelas_Impressao_30dias & .TextMatrix(lI, 0)
                    End If
                Next
            End If
        End With
        
        If frmT46V101A.objstcPedido.Cod_Bco_Cobr = 999 And _
           frmT46V101A.objstcPedido.Cod_Agen_Bas_Cor = 9999 And _
           frmT46V101A.objstcPedido.Num_Cta_Bas_Cor = 9999999 Then
            With flxCartao
                If .Rows > 1 Then
                    For lI = 1 To .Rows - 1
                        .row = lI
    
                        If .CellPicture = frmT46V124A.picChecked Then
                            If lParcelas_Impressao_Cartao <> "" Then lParcelas_Impressao_Cartao = lParcelas_Impressao_Cartao & ","
                            lParcelas_Impressao_Cartao = lParcelas_Impressao_Cartao & .TextMatrix(lI, 0)
                        End If
                    Next
                End If
            End With
        End If

        Set lobjCtrlEmis = New clsA46V712A
        Set lobjCtrlEmis.M46V999 = gM46V999

        lsSQl = "UPDATE Tab_Ctrl_Emis_Compl Set"
        lsSQl = lsSQl & " Parcelas_Impressao_Avista = '" & lParcelas_Impressao_Avista & "',"
        lsSQl = lsSQl & " Parcelas_Impressao_30dias = '" & lParcelas_Impressao_30dias & "',"
        lsSQl = lsSQl & " Parcelas_Impressao_Cartao = '" & lParcelas_Impressao_Cartao & "',"
        lsSQl = lsSQl & " Parcelas_Impressao_AvistaDebito = '" & lParcelas_Impressao_Avista & "',"   ''20180615: Juliano - Ajuste impressao
        lsSQl = lsSQl & " Parcelas_Impressao_30diasDebito = '" & lParcelas_Impressao_30dias & "'"    ''20180615: Juliano - Ajuste impressao
        lsSQl = lsSQl & " WHERE"
        lsSQl = lsSQl & " Nosso_Numero = " & gM46V999.gfFormatarCamposSQL(gM46V111.gComNosNumero)

        If Not lobjCtrlEmis.mfConectar Then GoTo ControleErro
        If gM46V999.gfExeSQL(lobjCtrlEmis.ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
        mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Avista = lParcelas_Impressao_Avista
        mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_30dias = lParcelas_Impressao_30dias
        mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Cartao = lParcelas_Impressao_Cartao
        mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_30diasDebito = lParcelas_Impressao_30dias       ''20180615: Juliano - Ajuste impressao
        mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_AvistaDebito = lParcelas_Impressao_Avista       ''20180615: Juliano - Ajuste impressao

        Set lobjCtrlEmis = Nothing

        Me.Hide
    End If

    Exit Sub

ControleErro:
    Call gM46V999.gpGraLog(1, "mpSalvar ERRO - " & Err & " " & Err.Description & vbCrLf & " SQL: " & lsSQl)
End Sub
Private Sub cmdVoltar_Click()
    If mfParcelaSelecionada = True Then
        Call mpSalvar
    End If
    
    Me.Hide
    frmT46V101A.Show
End Sub

Private Sub flxCartao_Click()
'Marca ou desmarca o registro na grade.
    With flxCartao
        If .ColSel = 4 Then
            .Col = .ColSel

            If .CellPicture = picChecked Then
                Set .CellPicture = picUnchecked
            Else
                Set .CellPicture = picChecked
            End If
        End If
    End With
End Sub

Private Sub flxVistaCarne_Click()
'Marca ou desmarca o registro na grade.
    With flxVistaCarne
        If .ColSel = 4 Then
            .Col = .ColSel

            If .CellPicture = picChecked Then
                Set .CellPicture = picUnchecked
            Else
                Set .CellPicture = picChecked
            End If
        End If
    End With
End Sub

Private Sub flx30diasCarne_Click()
'Marca ou desmarca o registro na grade.
    With flx30diasCarne
        If .ColSel = 4 Then
            .Col = .ColSel

            If .CellPicture = picChecked Then
                Set .CellPicture = picUnchecked
            Else
                Set .CellPicture = picChecked
            End If
        End If
    End With
End Sub

Private Sub flxVistaDebConta_Click()
'Marca ou desmarca o registro na grade.
    With flxVistaDebConta
        If .ColSel = 4 Then
            .Col = .ColSel

            If .CellPicture = picChecked Then
                Set .CellPicture = picUnchecked
            Else
                Set .CellPicture = picChecked
            End If
        End If
    End With
End Sub

Private Sub flx30diasDebConta_Click()
'Marca ou desmarca o registro na grade.
    With flx30diasDebConta
        If .ColSel = 4 Then
            .Col = .ColSel

            If .CellPicture = picChecked Then
                Set .CellPicture = picUnchecked
            Else
                Set .CellPicture = picChecked
            End If
        End If
    End With
End Sub

Private Sub Form_Activate()
    If Not gTemErro_T46V124A Then
            gM46V999.gEfetuouCalculo = True
        frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    End If
End Sub

Private Sub Form_DblClick()
    Call gpPos123_124
End Sub

Private Sub Form_Load()
    Dim lAntJuros As Variant    'Juros de parcelas antecipadas (de 1 a 12).
    Dim lCusEmissao As Single       'Custo de emissão.
    Dim lCusParcelado As String         'Indicador de custo de emissão parcelado.
    Dim lErro As Variant    'Erros da DLL de cálculo de fracionamento.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lIniVigencia As String        'Início de vigência.
    Dim lMensagem As String     'Mensagem.
    Dim lOrdem As Byte    'Ordem da parcela.
    Dim lParcela As Variant   'Resultado do fracionamento.
    Dim lPosJuros As Variant    'Juros de parcelas postecipadas (de 1 a 12).
    Dim lQuaParcelas As Byte          'Quantidade de parcelas.
    Dim lQuaParcVista As Integer        'Quantidade de parcelas as Vista
    Dim lQuaParc30Dia As Integer        'Quantidade de parcelas as Vista
    Dim lQuaParcCartao As Integer        'Quantidade de parcelas as Vista
    Dim lIOF As Single    'VD - IOF.
    Dim lPreLiquido As Single       'VD - Prêmio líquido.
    Dim lobjstcLoc As stcA46V708B    'Objeto da classe (stc) de item
    Dim lobjstcCobert As stcA46V704B    'Objeto da classe (stc) de coberturas.
    Dim lobjstcJuros As stcA46V711B   'Objeto da classe (stc) de juros.
    Dim objstcPedido As stcA46V702B   'Objeto da classe de estrutura stcA46V702B (TAB_PED)
    Dim lColecaoTMP As Collection

    Screen.MousePointer = vbHourglass
    Me.Height = 5190
    Me.Width = 6750
    Call Reposicionar_Controles

    '1. Prepara grades.
    Call mpMonGraFracionamento(flxVistaCarne, False)
    Call mpMonGraFracionamento(flxVistaDebConta, True)
    Call mpMonGraFracionamento(flx30diasCarne, False)
    Call mpMonGraFracionamento(flx30diasDebConta, True)
    
    If frmT46V101A.objstcPedido.Cod_Bco_Cobr = 999 And _
       frmT46V101A.objstcPedido.Cod_Agen_Bas_Cor = 9999 And _
       frmT46V101A.objstcPedido.Num_Cta_Bas_Cor = 9999999 Then
        Call mpMonGraFracionamento(flxCartao, True)
        Me.flxCartao.Visible = True
        Me.lblCartao.Visible = True
    Else
        Me.flxCartao.Visible = False
        Me.lblCartao.Visible = False
    End If

    '2. Posiciona controle.
    gTemErro_T46V124A = False

    'c) Posiciona quantidade de parcelas.

    'Ficha 900665
    '    lQuaParcVista = gM46V999.gProQuaMaxParcelas
    '    lQuaParc30Dia = gM46V999.gProQuaMaxParcelas

    lQuaParcVista = mobjstcPed.Qtd_Parc
    lQuaParc30Dia = mobjstcPed.Qtd_Parc
    lQuaParcCartao = mobjstcPed.Qtd_Parc

    If mobjstcPed.Cod_Reg = 106 Or mobjstcPed.Cod_Reg = 107 Then
        lQuaParcVista = 1
    End If

        'Ajute retirar premio minimo do imobiliario
    If lQuaParcVista = 0 And lQuaParc30Dia = 0 And (mobjstcPed.Cod_Reg <> 120 And (mobjstcPed.Cod_Reg <> 129 Or mobjstcPed.Cod_Estip <> 20)) Then
        Call gpMensagem("FGEXXXX - Quantidade de parcelas inválida (zerada - fracionamento não " & _
                        "realizado).")
        gTemErro_T46V124A = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If InStr(UCase(mobjstcPed.getPlano.NomPlano), "MERCEDES") Or _
       InStr(UCase(mobjstcPed.getPlano.NomPlano), "RVA") Then
        lQuaParcVista = 11
    End If

    If mobjstcPed.Cod_Estip = 20 Then
        lQuaParcVista = 12
        lQuaParc30Dia = 12
    End If

    'Guarda o maior numero de parcelas
    If lQuaParcVista > lQuaParc30Dia Then
        lQuaParcelas = lQuaParcVista
    Else
        lQuaParcelas = lQuaParc30Dia
    End If

    'd) Posiciona data de início de vigência.
    lIniVigencia = gM46V999.gfForData(2, mobjstcPed.Dat_Ini_Vig)

    '4. Posiciona dados do form T46V123A.
    lPreLiquido = 0
    lCusEmissao = 0
    'a) Posiciona prêmio líquido e custo de emissão - valor determinado.
    With frmT46V123A.flxResValores
        .Col = 1      'Valor.
        .Row = 0      'Prêmio líquido.
        lPreLiquido = CSng(.Text)
        .Row = 1      'Custo de emissão.
        lCusEmissao = CSng(.Text)
    End With

    '5. Executa cálculo de fracionamento.

    Set lobjstcLoc = mobjstcPed.ColecaoPedLoc.Item(1)
    mobjstcPed.VAL_PR = lPreLiquido
    If gM46V998.gfCalFracionamento(mobjstcPed, lobjstcLoc.Tip_Emissao, lParcela, lMensagem, gM46V111.gProJurIOF, mobjstcPed.Cod_Reg, lIOF, mobjstcPed.Val_Parc_Min) = False Then
        If gM46V999.gfPreenchido(lMensagem) Then
            Call gpMensagem(lMensagem)
        End If
        gTemErro_T46V124A = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
        If mobjstcPed.Cod_Estip <> 20 Then
            If gM46V999.gParcMaxCorr <> 0 And mobjstcPed.Cod_Produto >= 1086 And mobjstcPed.Cod_Reg <> 106 And mobjstcPed.Cod_Reg <> 107 And mobjstcPed.Cod_Reg <> 120 Then
                lQuaParcelas = gM46V999.gParcMaxCorr
                lQuaParcVista = gM46V999.gParcMaxCorr
                lQuaParc30Dia = gM46V999.gParcMaxCorr
                lQuaParcCartao = gM46V999.gParcMaxCorr
            Else
                If mobjstcPed.Cod_Reg = 106 Or mobjstcPed.Cod_Reg = 107 Or mobjstcPed.Cod_Reg = 120 Then
                    'atualizar o campo na tela quando endosso
                    'Ajute retirar premio minimo do imobiliario
                    Dim gCarga_Ant As Boolean
                    gCarga_Ant = gCarga
                    gCarga = True
                    frmT46V102A.txtQuaParcelas = 1
                    gCarga = gCarga_Ant
                End If
            End If
        End If

    'a) Valor
    With frmT46V123A.flxResValores
        .TextMatrix(2, 1) = Format$(lIOF, "standard") & " "    'IOF. 'Valor.    ===>>> grade da janela frmT46V123A...
        .TextMatrix(3, 1) = Format$(lPreLiquido + lCusEmissao + lIOF, "standard") & " "    'Total.
    End With
    flxVistaCarne.Rows = 1
    flx30diasCarne.Rows = 1
    flxVistaDebConta.Rows = 1
    flx30diasDebConta.Rows = 1
    
    If frmT46V101A.objstcPedido.Cod_Bco_Cobr = 999 And _
       frmT46V101A.objstcPedido.Cod_Agen_Bas_Cor = 9999 And _
       frmT46V101A.objstcPedido.Num_Cta_Bas_Cor = 9999999 Then
        flxCartao.Rows = 1
    End If

    'marca todas as parcelas
    Dim lParcelas As String
    Dim i As Integer

    lParcelas = Empty

    For i = 1 To lQuaParcelas
        lParcelas = lParcelas & i & ","
    Next i

    If Len(lParcelas) > 0 Then
        lParcelas = Left(lParcelas, Len(lParcelas) - 1)
    End If

    'Flex Grid pgto Antecipado Carne
    For lOrdem = 1 To lQuaParcVista

        Set lobjstcJuros = mobjstcPed.getstcJuros(lOrdem, e_FormaPagamento_Carne)

        If lParcela(lOrdem, 1, e_FormaPagamento.e_FormaPagamento_Carne) <> 0 Then

            If Trim(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Avista) = Empty Or val(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Avista) = 0 Then
                mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Avista = lParcelas
            End If

            Call mpPosValor(flxVistaCarne, lOrdem, lParcela(lOrdem, 1, e_FormaPagamento.e_FormaPagamento_Carne), lParcela(lOrdem, 2, e_FormaPagamento.e_FormaPagamento_Carne), _
                            lobjstcJuros.Val_Ant, mfVerifica_Parcela_Imprimir(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Avista, lOrdem))

        End If
    Next lOrdem

    'Flex Grid pgto 30 dias Carne
    For lOrdem = 1 To lQuaParc30Dia

        Set lobjstcJuros = mobjstcPed.getstcJuros(lOrdem, e_FormaPagamento_Carne)

        If lParcela(lOrdem, 3, e_FormaPagamento.e_FormaPagamento_Carne) <> 0 Then

            If Trim(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_30dias) = Empty Or val(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_30dias) = 0 Then
                mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_30dias = lParcelas
            End If

            Call mpPosValor(flx30diasCarne, lOrdem, lParcela(lOrdem, 3, e_FormaPagamento.e_FormaPagamento_Carne), lParcela(lOrdem, 4, e_FormaPagamento.e_FormaPagamento_Carne), _
                            lobjstcJuros.Val_Pos, mfVerifica_Parcela_Imprimir(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_30dias, lOrdem))
        End If
    Next lOrdem

    'Flex Grid pgto Antecipado DebConta
    For lOrdem = 1 To lQuaParcVista

        Set lobjstcJuros = mobjstcPed.getstcJuros(lOrdem, e_FormaPagamento_Debito)

        If lParcela(lOrdem, 1, e_FormaPagamento.e_FormaPagamento_Debito) <> 0 Then

            If Trim(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Avista) = Empty Or val(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Avista) = 0 Then
                mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Avista = lParcelas
            End If

            Call mpPosValor(flxVistaDebConta, lOrdem, lParcela(lOrdem, 1, e_FormaPagamento.e_FormaPagamento_Debito), lParcela(lOrdem, 2, e_FormaPagamento.e_FormaPagamento_Debito), _
                            lobjstcJuros.Val_Ant, mfVerifica_Parcela_Imprimir(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Avista, lOrdem))

        End If
    Next lOrdem
    
    If frmT46V101A.objstcPedido.Cod_Bco_Cobr = 999 And _
       frmT46V101A.objstcPedido.Cod_Agen_Bas_Cor = 9999 And _
       frmT46V101A.objstcPedido.Num_Cta_Bas_Cor = 9999999 Then
        
        'Flex Grid pgto Antecipado Cartao
        For lOrdem = 1 To lQuaParcCartao
    
            Set lobjstcJuros = mobjstcPed.getstcJuros(lOrdem, 4)
    
            If lParcela(lOrdem, 1, 4) <> 0 Then
    
                If Trim(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Cartao) = Empty Or val(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Cartao) = 0 Then
                    mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Cartao = lParcelas
                End If
                
                Call mpPosValor(flxCartao, lOrdem, lParcela(lOrdem, 1, 4), lParcela(lOrdem, 2, 4), _
                                lobjstcJuros.Val_Ant, mfVerifica_Parcela_Imprimir(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_Cartao, lOrdem))
    
            End If
        Next lOrdem
        
    End If

    'Flex Grid pgto 30 dias DebConta
    For lOrdem = 1 To lQuaParc30Dia

        Set lobjstcJuros = mobjstcPed.getstcJuros(lOrdem, e_FormaPagamento_Debito)

        If lParcela(lOrdem, 3, e_FormaPagamento.e_FormaPagamento_Debito) <> 0 Then

            If Trim(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_30dias) = Empty Or val(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_30dias) = 0 Then
                mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_30dias = lParcelas
            End If
                
            '8000046897, Parcelamento Syas Kit
            If (gM46V999.gCod_Ramo = 112 And (mobjstcPed.ColecaoPedLoc.Item(1).Tip_Emissao = e_TipEmissao_RenCongenere Or mobjstcPed.ColecaoPedLoc.Item(1).Tip_Emissao = e_TipEmissao_SeguroNovo)) Then
                
                Dim lsSQl As String
                Dim strParcelas_Impressao_30dias As String
                
                'Verificar se já existe alguma opção de impressão na base.
                lsSQl = "SELECT Parcelas_Impressao_30dias FROM Tab_Ctrl_Emis_Compl (nolock) "
                lsSQl = lsSQl & " WHERE"
                lsSQl = lsSQl & " Nosso_Numero = " & gM46V999.gfFormatarCamposSQL(gM46V111.gComNosNumero)

                Dim lbdTabCtrl As ADODB.Connection
                Dim lrsTabCtrl As clsYasRecordSet  'Registro : Tab_Ctrl_Emis_Compl
                Dim lSelect As String
            
                If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTabCtrl, lMensagem) = False Then
                    Call gM46V999.gpGraLog(2, lMensagem)
                End If
            
                If Not gM46V999.gfObtRegistro(lbdTabCtrl, lsSQl, lrsTabCtrl, lMensagem) Then
                    Call gM46V999.gpGraLog(1, lMensagem)
                End If
                
                'Recupera valor.
                If Not lrsTabCtrl.EOF Then
                    strParcelas_Impressao_30dias = lrsTabCtrl("Parcelas_Impressao_30dias")
                End If
                
                'Se não tem nada gravado, atribui Empty ao inves del Parcelas.
                If Trim(strParcelas_Impressao_30dias) = Empty Or strParcelas_Impressao_30dias = "0" Then
                    mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_30dias = Empty
                End If
            
            End If

            Call mpPosValor(flx30diasDebConta, lOrdem, lParcela(lOrdem, 3, e_FormaPagamento.e_FormaPagamento_Debito), lParcela(lOrdem, 4, e_FormaPagamento.e_FormaPagamento_Debito), _
                            lobjstcJuros.Val_Pos, mfVerifica_Parcela_Imprimir(mobjstcPed.objstcCtrlEmis.Parcelas_Impressao_30dias, lOrdem))
        End If
    Next lOrdem

    'b) Valor de mercado.
    With frmT46V123A.flxResValores
        .TextMatrix(2, 1) = Format$(lIOF, "standard") & " "  'IOF. 'Valor.    ===>>> grade da janela frmT46V123A...
        'If lQuaParcelas = 1 Then
        '    .TextMatrix(3, 1) = Format$(lParcela(1, 5), "standard")
        'Else
        .TextMatrix(3, 1) = Format$(lPreLiquido + lCusEmissao + lIOF, "standard") & " "   'Total.
        'End If
    End With

    'Exibe formas de pagamento à Vista
    tabResultado.TabVisible(0) = (flxVistaCarne.Rows > 1)
    'Exibe formas de pagamento à 30 dias
    tabResultado.TabVisible(1) = (flx30diasCarne.Rows > 1)

    If tabResultado.TabVisible(0) = True Or tabResultado.TabVisible(1) = True Then
        Call mpSalvar
    End If

    '7. Centraliza e prepara form.
    Call gpPos123_124
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Resize()
'1. Se não estiver na forma normal, encerra processo.
    If Me.WindowState <> vbNormal Then
        Exit Sub
    End If

    '2. Altura e largura mínimas do form.
    If Me.Height < 4190 Then
        Me.Height = 5190
    End If
    If Me.Width < 6750 Then
        Me.Width = 6750
    End If

    '3. Dimensiona pasta de resultados.
    Call Reposicionar_Controles

    cmdVoltar.Top = Me.Height - 1200
    cmdMosPremios.Top = cmdVoltar.Top
End Sub

Private Function mfVerifica_Parcela_Imprimir(ByVal pParcelas As String, ByVal pParcela As Integer) As Boolean
    mfVerifica_Parcela_Imprimir = False
    Dim arrParcelas
    Dim i As Integer

    If Len(pParcelas) > 0 Then
        arrParcelas = Split(pParcelas, ",")

        For i = 0 To UBound(arrParcelas)
            If pParcela = CInt(arrParcelas(i)) Then
                mfVerifica_Parcela_Imprimir = True
                Exit For
            End If
        Next
        End If

    If mfVerifica_Parcela_Imprimir = False Then ' Apenas para permitir colocar um Breakpoint quando retorna false
        mfVerifica_Parcela_Imprimir = mfVerifica_Parcela_Imprimir
    End If

End Function

Private Sub Reposicionar_Controles()
'Tamanho
    tabResultado.Width = Me.Width - 300
    tabResultado.Height = Me.Height - 1800

    'Vista
    flxVistaCarne.Width = (tabResultado.Width / 2) - 250
    flxVistaCarne.Height = tabResultado.Height - 500
    lblCarne.Width = flxVistaCarne.Width

    flxVistaDebConta.Left = flxVistaCarne.Left + flxVistaCarne.Width + 10
    flxVistaDebConta.Width = (tabResultado.Width / 2) + 260
    flxVistaDebConta.Height = flxVistaCarne.Height
    lblDebConta.Left = lblCarne.Left + lblCarne.Width
    lblDebConta.Width = flxVistaDebConta.Width

    'Debito em Conta
    flx30diasCarne.Width = flxVistaCarne.Width
    flx30diasCarne.Height = flxVistaCarne.Height

    flx30diasDebConta.Left = flx30diasCarne.Left + flx30diasCarne.Width + 10
    flx30diasDebConta.Width = flxVistaDebConta.Width
    flx30diasDebConta.Height = flxVistaDebConta.Height



    'flxVistaDebConta

    '    flx30diasCarne
    '    flx30diasDebConta
    '
    '    lblCarne
    '    lblDebConta

    'Posição
    Me.Refresh
    DoEvents
    Exit Sub


End Sub

Private Sub tabResultado_Click(PreviousTab As Integer)
    If tabResultado.Tab = 0 Then
        flx30diasCarne.Visible = False
        flx30diasDebConta.Visible = False
        flxVistaCarne.Visible = True
        flxVistaDebConta.Visible = True
'        If frmT46V101A.objstcPedido.Cod_Bco_Cobr = 999 And _
'           frmT46V101A.objstcPedido.Cod_Agen_Bas_Cor = 9999 And _
'           frmT46V101A.objstcPedido.Num_Cta_Bas_Cor = 9999999 Then
            flxCartao.Visible = True
'        Else
'            flxCartao.Visible = False
'        End If
    Else
        flxVistaCarne.Visible = True
        flxVistaDebConta.Visible = True
        flx30diasCarne.Visible = True
        flx30diasDebConta.Visible = True
        flxCartao.Visible = False
        flx30diasDebConta.ZOrder 0
    End If
End Sub
