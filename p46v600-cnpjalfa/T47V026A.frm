VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmT47V026A 
   Caption         =   "Demonstrativo do parcelamento"
   ClientHeight    =   4515
   ClientLeft      =   4425
   ClientTop       =   1800
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4515
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabResultado 
      Height          =   3945
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6959
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
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
      TabPicture(0)   =   "T47V026A.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "panVista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "1ª parcela a &30 dias"
      TabPicture(1)   =   "T47V026A.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pan30dias"
      Tab(1).ControlCount=   1
      Begin Threed.SSPanel panVista 
         Height          =   3495
         Left            =   45
         TabIndex        =   1
         Top             =   45
         Width           =   3700
         _Version        =   65536
         _ExtentX        =   6526
         _ExtentY        =   6165
         _StockProps     =   15
         ForeColor       =   12582912
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin MSFlexGridLib.MSFlexGrid flxFraVista 
            Height          =   3435
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   15
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   6059
            _Version        =   393216
            Rows            =   13
            Cols            =   4
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
      End
      Begin Threed.SSPanel pan30dias 
         Height          =   3495
         Left            =   -74955
         TabIndex        =   3
         Top             =   45
         Width           =   3700
         _Version        =   65536
         _ExtentX        =   6526
         _ExtentY        =   6165
         _StockProps     =   15
         ForeColor       =   12582912
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin MSFlexGridLib.MSFlexGrid flxFra30dia 
            Height          =   3435
            Left            =   45
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   15
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   6059
            _Version        =   393216
            Rows            =   13
            Cols            =   4
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
      End
   End
   Begin Threed.SSCommand cmdVoltar 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   45
      TabIndex        =   5
      ToolTipText     =   "Volta ao item."
      Top             =   4170
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
      Left            =   1905
      TabIndex        =   6
      ToolTipText     =   "Volta ao item."
      Top             =   4170
      Width           =   1905
      _Version        =   65536
      _ExtentX        =   3351
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "Mostrar &prêmios do item"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT47V026A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjstcPed As stcA46V702B


Public Sub gpInstanciarForm(pobjstcPed As stcA46V702B)
    Set mobjstcPed = pobjstcPed
    If gM46V999.gfDisponivel("frmT47V026A") Then
        Unload Me
    End If
    Load Me
End Sub
Private Sub mpMonGraFracionamento(ByVal pGrade As MSFlexGrid)
    'Procedure: monta grade de fracionamento.

    With pGrade
        .Rows = 1
        .Row = 0
        'a) Posiciona fonte.
        .Font = 5
        .FontWidth = 3.35
        'b) Posiciona largura das colunas.
        .ColWidth(0) = 255
        .ColWidth(1) = 1100
        .ColWidth(2) = 1100
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
    End With
End Sub
Private Sub mpPosValor(ByVal pGrade As MSFlexGrid, ByVal pLinha As Byte, ByVal pValParcela As Double, _
                       ByVal pValDemais As Double, ByVal pJuros As Double)
    'Procedure: posiciona valor na grade.

    'Parâmetros: pGrade........grade utilizada.
    '            pLinha........número da linha.
    '            pValParcela...valor da primeira parcela.
    '            pValDemais....valor das demais parcelas.
    '            pJuros........percentual de juros.

    With pGrade
        .Rows = pLinha + 1
        .TextMatrix(pLinha, 0) = pLinha  'Número de parcelas.
        .TextMatrix(pLinha, 1) = Format$(pValParcela, "standard") 'Primeira parcela.
        .TextMatrix(pLinha, 2) = Format$(pValDemais, "standard") 'Demais parcelas.
        .TextMatrix(pLinha, 3) = Format$(pJuros, "standard")  'Juros.
    End With
End Sub
Private Sub cmdMosPremios_Click()
    Me.Hide
    frmT47V025A.Show vbModeless, frmT47V001A
    If frmT47V025A.tabPremios.Visible = True And frmT47V025A.tabPremios.Enabled = True Then
        frmT47V025A.tabPremios.SetFocus
    End If
End Sub
Private Sub cmdVoltar_Click()
    Me.Hide
    frmT47V001A.Show
End Sub

Private Sub Form_Activate()
    If gTemErro_T47V026A = True Then
        Unload Me
    End If
    gM46V999.gEfetuouCalculo = True
End Sub

Private Sub Form_DblClick()
    Call gpPos025_026
End Sub

Private Sub Form_Load()
    Dim lJuros          As Variant           'Juros da tela frmT47V004A
    Dim lAntJuros       As Variant          'Juros de parcelas antecipadas (de 1 a 12).
    Dim lCusEmissao     As Single           'Custo de emissão.
    Dim lCusParcelado   As String           'Indicador de custo de emissão parcelado.
    Dim lErro           As Variant          'Erros da DLL de cálculo de fracionamento.
    Dim lI              As Byte             'Utilizado com For...Next.
    Dim lIniVigencia    As String           'Início de vigência.
    Dim lMensagem       As String           'Mensagem.
    Dim lOrdem          As Byte             'Ordem da parcela.
    Dim lParcela        As Variant          'Resultado do fracionamento.
    Dim lPosJuros       As Variant          'Juros de parcelas postecipadas (de 1 a 12).
    Dim lQuaParcelas    As Byte             'Quantidade de parcelas.
    Dim lQuaParcVista   As Integer          'Quantidade de parcelas as Vista
    Dim lQuaParc30Dia   As Integer          'Quantidade de parcelas as Vista
    Dim lIOF            As Single           'VD - IOF.
    Dim lPreLiquido     As Single           'VD - Prêmio líquido.
    Dim lobjstcLoc      As stcA46V708B      'Objeto da classe (stc) de item
    Dim lobjstcCobert   As stcA46V704B      'Objeto da classe (stc) de coberturas.
    Dim lobjstcJuros    As stcA46V711B      'Objeto da classe (stc) de juros.
    Dim objstcPedido    As stcA46V702B      'Objeto da classe de estrutura stcA46V702B (TAB_PED)
    Dim lColecaoTMP     As Collection
    
    Screen.MousePointer = vbHourglass

    '1. Prepara grades.
    Call mpMonGraFracionamento(flxFra30dia)
    Call mpMonGraFracionamento(flxFraVista)

    '2. Posiciona controle.
    gTemErro_T47V026A = False
    
    'c) Posiciona quantidade de parcelas.
    
    'Ficha 900665
'    lQuaParcVista = gM46V999.gProQuaMaxParcelas
'    lQuaParc30Dia = gM46V999.gProQuaMaxParcelas
    
     lQuaParcVista = mobjstcPed.Qtd_Parc
     lQuaParc30Dia = mobjstcPed.Qtd_Parc
        
    If lQuaParcVista = 0 And lQuaParc30Dia = 0 And _
       (mobjstcPed.Cod_Reg <> 106 And mobjstcPed.Cod_Reg <> 107 And mobjstcPed.Cod_Reg <> 120 And _
        mobjstcPed.Cod_Reg <> 121 And mobjstcPed.Cod_Reg <> 122 And mobjstcPed.Cod_Reg <> 20) Then
        Call gpMensagem("FGEXXXX - Quantidade de parcelas inválida (zerada - fracionamento não " & _
                        "realizado).")
        gTemErro_T47V026A = True
        Screen.MousePointer = vbDefault
        Exit Sub
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
    With frmT47V025A.flxResValores
        .Col = 1      'Valor.
        .Row = 0      'Prêmio líquido.
        lPreLiquido = CSng(.Text)
        .Row = 1      'Custo de emissão.
        lCusEmissao = CSng(.Text)
    End With

    '5. Executa cálculo de fracionamento.
    lJuros = Format(frmT47V004A.txtJuros.Text, "0.00")
    
    Set lobjstcLoc = mobjstcPed.ColecaoPedLoc.Item(1)
    If gfCalFracionamento(lQuaParcelas, frmT47V001A.txtVigInicio.Text, frmT47V001A.txtVigTermino.Text, _
       lCusEmissao, IIf(mobjstcPed.gCustoParcelado = True, "S", "N"), _
       lPreLiquido, lJuros, lobjstcLoc.Tip_Emissao, mobjstcPed.Cod_Produto, _
       lParcela, lMensagem, gM46V111.gProJurIOF, mobjstcPed.Cod_Reg, lIOF, mobjstcPed.Val_Parc_Min) = False Then
        If gM46V999.gfPreenchido(lMensagem) Then
            Call gpMensagem(lMensagem)
        End If
        gTemErro_T47V026A = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    'a) Valor
    'If frmT47V001A.lVerTipEmissao = "E" Then lIOF = 0
    With frmT47V025A.flxResValores
        .TextMatrix(2, 1) = Format$(lIOF, "standard") & " " 'IOF. 'Valor.    ===>>> grade da janela frmT47V025A...
        .TextMatrix(3, 1) = Format$(lPreLiquido + lCusEmissao + lIOF, "standard") & " " 'Total.
    End With
    flxFraVista.Rows = 1
    flxFra30dia.Rows = 1
    'Flex Grid pgto Antecipado
    For lOrdem = 1 To lQuaParcVista
        If lParcela(lOrdem, 1) <> 0 Then
            Call mpPosValor(flxFraVista, lOrdem, lParcela(lOrdem, 1), lParcela(lOrdem, 2), _
                            lJuros)
        
        End If
    Next lOrdem
    'Flex Grid pgto 30 dias
    For lOrdem = 1 To lQuaParc30Dia
        If lParcela(lOrdem, 3) <> 0 Then
            Call mpPosValor(flxFra30dia, lOrdem, lParcela(lOrdem, 3), lParcela(lOrdem, 4), _
                                lJuros)
        End If
    Next lOrdem
    
    'b) Valor de mercado.
    With frmT47V025A.flxResValores
        .TextMatrix(2, 1) = Format$(lIOF, "standard") & " "  'IOF. 'Valor.    ===>>> grade da janela frmT47V025A...
        'If lQuaParcelas = 1 Then
        '    .TextMatrix(3, 1) = Format$(lParcela(1, 5), "standard")
        'Else
            .TextMatrix(3, 1) = Format$(lPreLiquido + lCusEmissao + lIOF, "standard") & " "   'Total.
        'End If
    End With

    'flxFraVista.Rows = lQuaParcVista + 1
    'flxFra30dia.Rows = lQuaParc30Dia + 1
    
    '7. Centraliza e prepara form.
    Call gpPos025_026
    
    'Exibe formas de pagamento à Vista
    tabResultado.TabVisible(0) = (flxFraVista.Rows > 1)
    'Exibe formas de pagamento à 30 dias
    tabResultado.TabVisible(1) = (flxFra30dia.Rows > 1)
    
End Sub

Private Sub Form_Resize()
    '1. Se não estiver na forma normal, encerra processo.
    If Me.WindowState <> vbNormal Then
        Exit Sub
    End If

    '2. Altura e largura mínimas do form.
    If Me.Height < 2700 Then
        Me.Height = 2700
    End If
    If Me.Width < 1000 Then
        Me.Width = 1000
    End If

    '3. Dimensiona pasta de resultados.
    With tabResultado
        'a) Tamanho da pasta de resultados.
        .Height = Me.Height - 850
        .Width = Me.ScaleWidth
        'b) Pasta 0: primeira parcela a vista.
        panVista.Height = .Height - 450
        panVista.Width = .Width - 100
        flxFraVista.Height = IIf(panVista.Height > 300, panVista.Height - 300, 0)
        flxFraVista.Width = panVista.Width - 100
        'c) Pasta 1: primeira parcela a 30 dias.
        pan30dias.Height = panVista.Height
        pan30dias.Width = panVista.Width
        flxFra30dia.Height = flxFraVista.Height
        flxFra30dia.Width = flxFraVista.Width
    End With
    cmdVoltar.Top = Me.Height - 750
    cmdMosPremios.Top = Me.Height - 750
End Sub

Private Sub tabResultado_Click(PreviousTab As Integer)
    If tabResultado.Tab = 0 Then
        flxFraVista.Visible = True
        flxFra30dia.Visible = False
    Else
        flxFraVista.Visible = False
        flxFra30dia.Visible = True
    End If
End Sub
