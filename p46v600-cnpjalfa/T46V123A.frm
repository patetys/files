VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmT46V123A 
   Caption         =   "Demonstrativo de prêmios"
   ClientHeight    =   5130
   ClientLeft      =   5010
   ClientTop       =   1680
   ClientWidth     =   3825
   ControlBox      =   0   'False
   ForeColor       =   &H00C00000&
   Icon            =   "T46V123A.frx":0000
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5130
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel panAgravarCalculo 
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   4260
      Width           =   3795
      _Version        =   65536
      _ExtentX        =   6694
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      Begin VB.TextBox txtAgravarCalculo 
         Height          =   285
         Left            =   1785
         MaxLength       =   6
         TabIndex        =   8
         Top             =   60
         Width           =   585
      End
      Begin Threed.SSCommand cmdAgravarCalculo 
         Height          =   300
         Left            =   2580
         TabIndex        =   9
         ToolTipText     =   "Volta ao item."
         Top             =   60
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "Recalcular"
         ForeColor       =   8388608
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Agravar Cálculo:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   1770
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "%"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   10
         Top             =   120
         Width           =   120
      End
   End
   Begin TabDlg.SSTab tabPremios 
      Height          =   4200
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3800
      _ExtentX        =   6694
      _ExtentY        =   7408
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   706
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
      TabCaption(0)   =   "&Resumo"
      TabPicture(0)   =   "T46V123A.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "panValores"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstResultado"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Itens"
      TabPicture(1)   =   "T46V123A.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxPreValores"
      Tab(1).ControlCount=   1
      Begin VB.ListBox lstResultado 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   1230
         Left            =   30
         TabIndex        =   5
         Top             =   2160
         Width           =   3705
      End
      Begin Threed.SSPanel panValores 
         Height          =   2085
         Left            =   45
         TabIndex        =   1
         Top             =   45
         Width           =   3645
         _Version        =   65536
         _ExtentX        =   6429
         _ExtentY        =   3678
         _StockProps     =   15
         Caption         =   "Prêmios"
         ForeColor       =   8388608
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
         Begin MSFlexGridLib.MSFlexGrid flxResValores 
            Height          =   1725
            Left            =   45
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   300
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   3043
            _Version        =   393216
            Rows            =   5
            ForeColor       =   8388608
            ForeColorFixed  =   12582912
            BackColorSel    =   12582912
            ForeColorSel    =   16777215
            BackColorBkg    =   -2147483633
            GridColorFixed  =   8421504
            WordWrap        =   -1  'True
            AllowBigSelection=   -1  'True
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxPreValores 
         Height          =   3645
         Left            =   -74955
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   70
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   6429
         _Version        =   393216
         Rows            =   15
         ForeColor       =   8388608
         ForeColorFixed  =   12582912
         BackColorSel    =   12582912
         BackColorBkg    =   -2147483626
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
   End
   Begin Threed.SSCommand cmdVoltar 
      Height          =   300
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Volta ao item."
      Top             =   4725
      Width           =   1800
      _Version        =   65536
      _ExtentX        =   3175
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Voltar"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdMosFracionamento 
      Height          =   300
      Left            =   1905
      TabIndex        =   4
      ToolTipText     =   "Volta ao item."
      Top             =   4725
      Width           =   1905
      _Version        =   65536
      _ExtentX        =   3351
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "Mostrar &fracionamento"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V123A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mobjstcPed As stcA46V702B

Private Sub mpPrepararGrids()
'Procedure: monta grade de prêmios.

    Dim lI As Long    'Utilizado com For...Next.

    With flxPreValores
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 1
        .Rows = mobjstcPed.ColecaoPedLoc.Count + 1
        .Font = 5
        .FontWidth = 3.35
        'Largura das colunas.
        .ColWidth(0) = 900
        'Altura das linhas.
        For lI = 0 To .Rows - 1
            .RowHeight(lI) = 220
        Next lI
        'Alinhamento da coluna.
        .ColAlignment(0) = flexAlignRightCenter
        'Títulos.
        .TextMatrix(0, 0) = "Item"
        .TextMatrix(0, 1) = "Prêmio Liquido"
        For lI = 1 To .Rows - 1
            .TextMatrix(lI, 0) = "Item " & lI
        Next lI
    End With
    flxPreValores.ColWidth(1) = 940
    If gM46V999.gNovaPropostaLMI And (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
        'Não Mostrar coluna de prêmios qdo LMI para o corretor
        flxPreValores.ColWidth(1) = 0
        flxPreValores.AllowUserResizing = flexResizeNone
    End If

    With flxResValores
        .Font = 5
        .FontWidth = 3.35
        .ColWidth(0) = 850
        .ColWidth(1) = 940
        .RowHeight(0) = 220
        .ColAlignment(0) = flexAlignRightCenter
        .ColAlignment(1) = flexAlignRightCenter
        .Rows = 4
        .FixedRows = 1
        .TextMatrix(0, 0) = "Pr. líquido"
        .TextMatrix(1, 0) = "Custo emissão"
        .TextMatrix(2, 0) = "IOF"
        .TextMatrix(3, 0) = "TOTAL"
    End With

End Sub

Public Sub gpInstanciarForm(pobjstcPed As stcA46V702B)
    Set mobjstcPed = pobjstcPed
    Load Me
End Sub

Private Sub cmdAgravarCalculo_Click()
    If gM46V999.gfPreenchido(txtAgravarCalculo.Text) Then
        mobjstcPed.objstcCtrlEmis.Per_Desc_Agrav = txtAgravarCalculo.Text
        Call frmT46V101A.mnuEdiCalPremios_Click
    End If
End Sub

Private Sub cmdMosFracionamento_Click()
    frmT46V124A.Show vbModeless, frmT46V101A
    Me.Hide
End Sub
Private Sub cmdVoltar_Click()
    Me.Hide
    frmT46V101A.Show
End Sub
Private Sub Form_Activate()
'1. Se houve erro ao montar T46V123A, encerra processo.
    If gTemErro_T46V123A = True Then
        Unload Me
    End If

    '2. Se form de fracionamento não está disponível, desabilita botão.
    If Not gM46V999.gfDisponivel("FRMT46V124A") Then
        cmdMosFracionamento.Visible = False
    End If
    gM46V999.gEfetuouCalculo = True
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
End Sub
Private Sub Form_DblClick()
    Call gpPos123_124
End Sub
Private Sub Form_Load()
    Dim lAuxNumItem As String       'Auxiliar: número item.
    Dim lCusEmissao As Double       'Custo de emissão.
    Dim lI As Long    'Utilizado com For...Next.
    Dim lMensagem As String     'Mensagem.
    Dim lPerIOF As Double   'Percentual de IOF.
    Dim lValIOF As Double   'Valor do IOF
    Dim lobjstcLoc As stcA46V708B    'Objeto da classe(stc) de itens.
    Dim lobjstcCobert As stcA46V704B    'Objeto da classe(stc) de coberturas.
    Dim lValTotItem As Double       'Valor total do item.
    Dim lValorTotal As Double       'Valor total
    Dim lPremioMinimo As Double

    Screen.MousePointer = vbHourglass

    '1. Posiciona controles.
    gTemErro_T46V123A = False

    lCusEmissao = mobjstcPed.Val_Cust_Emis
    lValIOF = mobjstcPed.Val_IOF

    '5. Posiciona grades.
    'a) Prepara grades.
    Call mpPrepararGrids
    panValores.Left = 50
    panValores.Width = tabPremios.Width - 100
    flxResValores.Width = panValores.Width - 100

    'c) Preenche grades.
    lI = 0
    lValorTotal = 0
    If gM46V111.gTipoGrupMulti = e_GrupoRamo_Multi Then
        For Each lobjstcLoc In mobjstcPed.ColecaoPedLoc
            lValTotItem = 0
            lI = lI + 1
            For Each lobjstcCobert In lobjstcLoc.ColecaoPedCobert
                lValTotItem = lValTotItem + lobjstcCobert.VAL_PR
            Next lobjstcCobert
            If mobjstcPed.Cod_Dizer <> 0 Then
                'Ajute retirar premio minimo do imobiliario
                If gM46V999.gPermitePrInformado Or gM46V111.gCod_plano = 20 Then
                    lPremioMinimo = lValTotItem
                    lValIOF = (lValTotItem * 7.38) / 100
                    mobjstcPed.Val_IOF = lValIOF
                Else
                lPremioMinimo = gM46V999.gProPreMinEndosso
                End If
            Else
				'Ajute retirar premio minimo do imobiliario
                If gM46V111.gCod_plano <> 20 Then
                	lPremioMinimo = gM46V999.gProPremioMinimo
                Else
                    lPremioMinimo = lValTotItem
                End If
            End If

            flxPreValores.TextMatrix(lI, 0) = lobjstcLoc.Num_Item
            If lValTotItem < 0 Then
                If lValTotItem > lPremioMinimo Then
                    flxPreValores.TextMatrix(lI, 1) = Format$(Round(0, 2), "standard")    'Prêmio
                Else
                    flxPreValores.TextMatrix(lI, 1) = Format$(Round(lValTotItem, 2), "standard")    'Prêmio
                End If
            Else
                If lValTotItem < lPremioMinimo Then
                    'Prêmio Mínimo
                    If lobjstcLoc.Tip_Emissao = e_TipEmissao_Endosso And lValTotItem < CDbl(8) Then
                        flxPreValores.TextMatrix(lI, 1) = Format$(Round(0, 2), "standard")    'Prêmio
                        lValTotItem = 0
                    Else
                        flxPreValores.TextMatrix(lI, 1) = Format$(Round(lPremioMinimo, 2), "standard")    'Prêmio
                    End If
                Else
                    flxPreValores.TextMatrix(lI, 1) = Format$(Round(lValTotItem, 2), "standard")    'Prêmio
                End If
            End If
            lValorTotal = lValorTotal + Round(lValTotItem, 2)
        Next lobjstcLoc
    ElseIf gM46V111.gTipoGrupMulti = e_GrupoRamo_Auto Then
        lValorTotal = mobjstcPed.objstcPedAuto.Val_Pr_Auto + mobjstcPed.objstcPedAuto.Val_Pr_Rc + mobjstcPed.objstcPedAuto.Val_Pr_App
    Else
        lValorTotal = mobjstcPed.VAL_PR
    End If
    
    If mobjstcPed.ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_Endosso Then
        If lValorTotal < mobjstcPed.gValPrLiqTotal And mobjstcPed.Cod_Dizer <> 10 Then
            lValorTotal = mobjstcPed.gValPrLiqTotal
        End If
    End If
    
    lValorTotal = Round(lValorTotal, 2)
    
    '6. Totaliza prêmios e posiciona grade de resumo.
    Call mpTotPremios(lValorTotal, lCusEmissao, lPerIOF, lValIOF, mobjstcPed.Cod_Reg, mobjstcPed.Qtd_Parc)

    '7. Centraliza e prepara form.
    Call gpPos123_124
    If mobjstcPed.LMI_Unico = 1 Then    'And (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
        flxPreValores.Visible = False
    End If

    '8. Atribui valor de agravo
    If mobjstcPed.objstcCtrlEmis.Per_Desc_Agrav <> 0 Then
        txtAgravarCalculo.Text = mobjstcPed.objstcCtrlEmis.Per_Desc_Agrav
    Else
        txtAgravarCalculo.Text = Empty
    End If

    Screen.MousePointer = vbDefault
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

    '3. Dimensiona pasta de prêmios.
    With tabPremios
        'a) Tamanho da pasta de prêmios.
        .Height = Me.Height - 1500
        .Width = Me.ScaleWidth
        'b) Pasta 0: resumo.
        panValores.Height = IIf(.Height > 930, .Height - 930, 0)
        If panValores.Height > 1850 Then
            panValores.Height = 1850
        End If
        panValores.Width = .Width / 2
        If panValores.Height > 360 Then
            flxResValores.Height = panValores.Height - 360
        Else
            flxResValores.Height = 0
        End If
        flxResValores.Width = panValores.Width - 100
        'c) Pasta 0: resumo - lista de resultado.
        lstResultado.Height = .Height - panValores.Height - 675
        lstResultado.Top = panValores.Height + 150
        lstResultado.Width = .Width - 100
        'd) Pasta 1: itens com valor de mercado.
        flxPreValores.Height = IIf(.Height > 530, .Height - 530, 0)
        flxPreValores.Width = .Width - 100
    End With

    '4. Posiciona agravar cálculo.
    panAgravarCalculo.Top = tabPremios.Height + 100

    '5. Posiciona botões de comando.
    cmdVoltar.Top = panAgravarCalculo.Top + panAgravarCalculo.Height + 100
    cmdMosFracionamento.Top = panAgravarCalculo.Top + panAgravarCalculo.Height + 100
End Sub

Private Sub mpTotPremios(ByVal pValLiq As Double, ByVal pCusEmissao As Double, _
                         ByVal pPerIOF As Double, ByVal pValIOF As Double, ByVal pCodReg As Integer, ByVal pQtdParc As Integer)



'Procedure: totaliza prêmios.

    Dim lTotal As Double  'Total.
    Dim lI1 As Byte    'Utilizado com For...Next.
    Dim lI2 As Byte    'Utilizado com For...Next.
    Dim lPreTotal As Double     'Prêmio total.
    Dim lPremioMinimo As Double

    If pCodReg > 100 Then
        'Ajute retirar premio minimo do imobiliario
        If gM46V999.gPermitePrInformado Or gM46V111.gCod_plano = 20 Then
            lPremioMinimo = pValLiq
        Else
        	lPremioMinimo = gM46V999.gProPreMinEndosso
        End If
    Else
		'Ajute retirar premio minimo do imobiliario
        If gM46V111.gCod_plano <> 20 Then
        	lPremioMinimo = gM46V999.gProPremioMinimo
        Else
            lPremioMinimo = pValLiq
        End If
    End If
    'AquiSimony


    If pCodReg > 100 Then
        If pValLiq > 0 And pCodReg <> 101 And pCodReg <> 102 And pCodReg <> 103 And pCodReg <> 104 Then
            pCodReg = 101
        End If
        If pValLiq = 0 Then
            pCodReg = 120
        End If
        If pValLiq < 0 And pCodReg <> 106 And pCodReg <> 107 Then
            pCodReg = 106
        End If
    End If

    If pCodReg = 2 Or _
       pCodReg = 3 Or _
       pCodReg = 102 Or _
       pCodReg = 103 Or _
       pCodReg = 202 Or _
       pCodReg = 203 Or _
       pQtdParc = 0 Then
        pPerIOF = 0
        pValIOF = 0
    End If

    If pCodReg < 100 Then
        If pValLiq < lPremioMinimo Then
            pValLiq = lPremioMinimo
        End If
    Else
        If pValLiq > 0 Then
            If pValLiq < lPremioMinimo Then
                pValLiq = lPremioMinimo
            End If
        ElseIf pValLiq < 0 Then
            If pValLiq > (lPremioMinimo * -1) And frmT46V101A.objstcPedido.Cod_Produto <= 1076 Then
                pValLiq = 0
            End If
        End If
    End If
    gCarga = True
    If pCodReg = 120 Or pCodReg = 106 Or pCodReg = 107 Or _
       Mid(Format(pCodReg, "000"), 3, 1) = 2 Or _
       Mid(Format(pCodReg, "000"), 3, 1) = 4 Or _
       Mid(Format(pCodReg, "000"), 3, 1) = 5 Then
        frmT46V120A.txtCusEmissao.Text = 0
        '    Else
        '        If pCodReg > 100 Then
        '            frmT46V120A.txtCusEmissao.Text = Format$(gM46V999.gProCusEndosso, "standard")
        '        Else
        '            frmT46V120A.txtCusEmissao.Text = Format$(gM46V999.gProCusApolice, "standard")
        '        End If
    End If
    pCusEmissao = frmT46V120A.txtCusEmissao.Text
    gCarga = False

    If mobjstcPed.Qtd_Parc = 1 Then
        lPreTotal = gM46V999.gfTrunca(pValLiq + pCusEmissao + pValIOF, 2)
    Else
        lPreTotal = pValLiq + pCusEmissao + pValIOF
    End If

    '2. Posiciona resumo.
    With flxResValores
        .Font = 5
        .FontWidth = 3.35
        .ColWidth(0) = 850
        .ColWidth(1) = 940
        .RowHeight(0) = 220
        .ColAlignment(0) = flexAlignRightCenter
        .ColAlignment(1) = flexAlignRightCenter
        .Rows = 5
        .FixedRows = 4
        .Row = 0
        .Col = 0
        .Text = "Pr. líquido"
        .Col = 1
        .Text = Format$(pValLiq, "standard") & " "
        .Row = 1
        .Col = 0
        .Text = "Custo emissão"
        .Col = 1
        .Text = Format$(pCusEmissao, "standard") & " "
        .Row = 2
        .Col = 0
        If gM46V999.gfPreenchido(pPerIOF) Then
            .Text = IIf(pPerIOF = 0, "IOF", "IOF (" & pPerIOF & "%)")
        Else
            .Text = "IOF"
        End If
        .Col = 1
        .Text = Format$(pValIOF, "standard") & " "
        .Row = 3
        .Col = 0
        .Text = "TOTAL"
        .Col = 1
        .Text = Format$(lPreTotal, "standard") & " "
        .RowHeight(4) = 0
    End With
End Sub

Private Sub txtAgravarCalculo_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",-")

    If KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then Exit Sub

    If Len(txtAgravarCalculo.Text) = 0 Then
        txtAgravarCalculo.Text = Chr(KeyAscii)
    Else
        Dim lCasasAposVirgula As Integer

        If InStr(txtAgravarCalculo.Text, Chr(44)) = 0 Then
            lCasasAposVirgula = 0
        Else
            lCasasAposVirgula = Len(txtAgravarCalculo.Text) + 1 - InStr(txtAgravarCalculo.Text, Chr(44))
        End If

        If CDbl(txtAgravarCalculo.Text & Chr(KeyAscii)) < 1000# And CDbl(txtAgravarCalculo.Text & Chr(KeyAscii)) > -100# And lCasasAposVirgula < 3 Then
            txtAgravarCalculo.Text = txtAgravarCalculo.Text & Chr(KeyAscii)
        Else
            Beep
        End If
    End If
    KeyAscii = 0
    txtAgravarCalculo.SelStart = Len(txtAgravarCalculo.Text)
End Sub
