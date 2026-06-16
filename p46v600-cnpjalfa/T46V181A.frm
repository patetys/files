VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V181A 
   Caption         =   "Coberturas Vida"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3480
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraControles 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   60
      TabIndex        =   34
      Tag             =   "Frame para habilitar/desabilitar controles."
      Top             =   60
      Width           =   7845
      Begin VB.Frame fraIndenizacoes 
         Caption         =   "Garantias"
         ForeColor       =   &H00800000&
         Height          =   2835
         Left            =   3990
         TabIndex        =   18
         Top             =   60
         Width           =   3795
         Begin VB.TextBox txtGarAuxCestaBasica 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1575
            TabIndex        =   32
            Tag             =   "1"
            ToolTipText     =   "Auxílio Funeral."
            Top             =   2400
            Width           =   2100
         End
         Begin VB.TextBox txtGarMorteConjuge 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1575
            TabIndex        =   30
            Tag             =   "1"
            ToolTipText     =   "Auxílio Funeral."
            Top             =   2040
            Width           =   2100
         End
         Begin VB.TextBox txtGarMorte 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1575
            TabIndex        =   20
            Tag             =   "1"
            ToolTipText     =   "Morta."
            Top             =   240
            Width           =   2100
         End
         Begin VB.TextBox txtGarIPA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1575
            TabIndex        =   22
            Tag             =   "1"
            ToolTipText     =   "IPA."
            Top             =   600
            Width           =   2100
         End
         Begin VB.TextBox txtGarIEA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1575
            TabIndex        =   24
            Tag             =   "1"
            ToolTipText     =   "IEA."
            Top             =   960
            Width           =   2100
         End
         Begin VB.TextBox txtGarIFPD 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1575
            TabIndex        =   26
            Tag             =   "1"
            ToolTipText     =   "IPDF."
            Top             =   1320
            Width           =   2100
         End
         Begin VB.TextBox txtGarAuxFuneral 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1575
            TabIndex        =   28
            Tag             =   "1"
            ToolTipText     =   "Auxílio Funeral."
            Top             =   1680
            Width           =   2100
         End
         Begin VB.Label lblX 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Aux. Cesta Básica:"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   14
            Left            =   60
            TabIndex        =   31
            ToolTipText     =   "Auxílio Funeral."
            Top             =   2400
            Width           =   1485
         End
         Begin VB.Label lblX 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Morte Cônjuge:"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   13
            Left            =   60
            TabIndex        =   29
            ToolTipText     =   "Auxílio Funeral."
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label lblX 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Morte:"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   8
            Left            =   60
            TabIndex        =   19
            ToolTipText     =   "Morte."
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label lblX 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "IPA:"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   9
            Left            =   60
            TabIndex        =   21
            ToolTipText     =   "IPA."
            Top             =   600
            Width           =   1485
         End
         Begin VB.Label lblX 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "IEA:"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   23
            ToolTipText     =   "IEA."
            Top             =   960
            Width           =   1485
         End
         Begin VB.Label lblX 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "IFPD:"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   11
            Left            =   60
            TabIndex        =   25
            ToolTipText     =   "IFPD."
            Top             =   1320
            Width           =   1485
         End
         Begin VB.Label lblX 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Auxílio Funeral:"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   12
            Left            =   60
            TabIndex        =   27
            ToolTipText     =   "Auxílio Funeral."
            Top             =   1680
            Width           =   1485
         End
      End
      Begin VB.ComboBox cboPlanoDeVida 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V181A.frx":0000
         Left            =   1860
         List            =   "T46V181A.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Plano de Vida."
         Top             =   60
         Width           =   2000
      End
      Begin VB.ComboBox cboCobertura 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V181A.frx":0004
         Left            =   1860
         List            =   "T46V181A.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Cobertura."
         Top             =   420
         Width           =   2000
      End
      Begin VB.TextBox txtQtdeVidas 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "1"
         ToolTipText     =   "Quantidade de Vidas"
         Top             =   750
         Width           =   630
      End
      Begin VB.TextBox txtLMGPorVida 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1860
         TabIndex        =   7
         Tag             =   "1"
         ToolTipText     =   "LMG/Vida."
         Top             =   1110
         Width           =   1380
      End
      Begin VB.TextBox txtLMGFuneralPorVida 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1860
         TabIndex        =   12
         Tag             =   "1"
         ToolTipText     =   "LMG/Vida."
         Top             =   1830
         Width           =   1860
      End
      Begin VB.TextBox txtAuxCestaBasPorVida 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1860
         TabIndex        =   17
         Tag             =   "1"
         ToolTipText     =   "LMG/Vida."
         Top             =   2550
         Width           =   1380
      End
      Begin Threed.SSPanel panInclusaoDeConjuge 
         Height          =   315
         Left            =   1890
         TabIndex        =   35
         Top             =   2220
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2831
         _ExtentY        =   556
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
         BevelOuter      =   0
         Begin VB.OptionButton optInclusaoDeConjuge 
            Caption         =   "&Não"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   570
            TabIndex        =   15
            ToolTipText     =   "NÃO tem sinistro."
            Top             =   0
            Value           =   -1  'True
            Width           =   600
         End
         Begin VB.OptionButton optInclusaoDeConjuge 
            Caption         =   "&Sim"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   0
            TabIndex        =   14
            ToolTipText     =   "SIM, tem sinistro."
            Top             =   0
            Width           =   600
         End
      End
      Begin Threed.SSPanel panAuxilioFuneral 
         Height          =   315
         Left            =   1890
         TabIndex        =   36
         Top             =   1500
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2831
         _ExtentY        =   556
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
         BevelOuter      =   0
         Begin VB.OptionButton optAuxilioFuneral 
            Caption         =   "&Não"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   600
            TabIndex        =   10
            ToolTipText     =   "NÃO tem sinistro."
            Top             =   30
            Value           =   -1  'True
            Width           =   600
         End
         Begin VB.OptionButton optAuxilioFuneral 
            Caption         =   "&Sim"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   30
            TabIndex        =   9
            ToolTipText     =   "SIM, tem sinistro."
            Top             =   30
            Width           =   600
         End
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plano de Vida: "
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   15
         TabIndex        =   0
         ToolTipText     =   "Plano de Vida"
         Top             =   60
         Width           =   1830
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cobertura: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   15
         TabIndex        =   2
         ToolTipText     =   "Cobertura"
         Top             =   420
         Width           =   1830
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Qtde. de Vidas:"
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   2
         Left            =   15
         TabIndex        =   4
         ToolTipText     =   "Quantiade de Vidas."
         Top             =   750
         Width           =   1845
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LMG/Vida:"
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         Left            =   15
         TabIndex        =   6
         ToolTipText     =   "LMG/Vida."
         Top             =   1110
         Width           =   1860
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Auxílio Funeral:"
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   4
         Left            =   15
         TabIndex        =   8
         ToolTipText     =   "Auxílio Funeral."
         Top             =   1470
         Width           =   1845
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LMG Funeral/Vida:"
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   5
         Left            =   15
         TabIndex        =   11
         ToolTipText     =   "Auxílio Funeral."
         Top             =   1830
         Width           =   1830
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inclusão de Cônjuge:"
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   6
         Left            =   15
         TabIndex        =   13
         ToolTipText     =   "Auxílio Funeral."
         Top             =   2190
         Width           =   1830
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aux Cesta Básica/Vida:"
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   7
         Left            =   15
         TabIndex        =   16
         ToolTipText     =   "Auxílio Funeral."
         Top             =   2550
         Width           =   1830
      End
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   3488
      TabIndex        =   33
      ToolTipText     =   "Volta aos dados complementares."
      Top             =   3120
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V181A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lAuxCestaBasPorVida As Currency
Private lLMGFuneralPorVida As Currency
Private lMorteConjuge As Currency
Private mbdTbm_Plano_Vida As ADODB.Connection
Private mrsTbm_Plano_Vida As clsYasRecordSet

Private Sub Form_Load()
    Dim lSelect As String
    Dim pMensagem As String

    Call gM46V999.gpCenForm(Me)

    'Carrega planos de vida.
    With cboPlanoDeVida
        .AddItem ("Não Contratado")
        .ItemData(.NewIndex) = 0
        .AddItem ("Capital Global")
        .ItemData(.NewIndex) = 1
    End With
    cboPlanoDeVida.ListIndex = 1
    cboPlanoDeVida.Enabled = False

    'Carrega coberturas.
    If mfConsulta_TBM_PLANO_VIDA(True, 1) = False Then Exit Sub

    If gM46V999.gCod_Ramo = e_CodRamo_Condominio Or gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
    
        With cboCobertura
            '1a opção do combo.
            .AddItem ("Não Contratado")    '1a opção do combo.
            .ItemData(.NewIndex) = 0
    
            'Demais opções do combo.
            Do While Not mrsTbm_Plano_Vida.EOF
                .AddItem (Trim(mrsTbm_Plano_Vida("DSC_PLANO_VIDA")))
                .ItemData(.NewIndex) = mrsTbm_Plano_Vida("COD_PLANO_VIDA")
                mrsTbm_Plano_Vida.MoveNext
            Loop
        End With
    Else
        With cboCobertura
            '1a opção do combo.
            .AddItem ("Não Aplicável")    '1a opção do combo.
            .ItemData(.NewIndex) = 0
    
        End With
    End If
    
    'Carrega auxilio funeral.
    If mfConsulta_TBM_PLANO_VIDA(True, 2, , 4) = False Then Exit Sub

    If mrsTbm_Plano_Vida.RecordCount = 0 Then
        lLMGFuneralPorVida = 0
    Else
        lLMGFuneralPorVida = mrsTbm_Plano_Vida("VAL_LMG_PADRAO")
    End If

    'Carrega auxilio cesta basica.
    If mfConsulta_TBM_PLANO_VIDA(True, 2, , 5) = False Then Exit Sub

    If mrsTbm_Plano_Vida.RecordCount = 0 Then
        lAuxCestaBasPorVida = 0
    Else
        lAuxCestaBasPorVida = mrsTbm_Plano_Vida("VAL_LMG_PADRAO")
    End If

    'Carrega morte (conjuge).
    If mfConsulta_TBM_PLANO_VIDA(True, 2, , 6) = False Then Exit Sub

    If mrsTbm_Plano_Vida.RecordCount = 0 Then
        lMorteConjuge = 0
    Else
        lMorteConjuge = mrsTbm_Plano_Vida("VAL_LMG_PADRAO")
    End If
End Sub

Private Sub cboCobertura_Click()
    'Limpar os campos
    If cboCobertura.ListIndex = 0 Then
        txtQtdeVidas.Text = Empty
        txtLMGPorVida.Text = Empty
        optAuxilioFuneral(1).Value = True
        optInclusaoDeConjuge(1).Value = True
    End If
    Call mpCalculaGarantias
End Sub

Private Sub cmdOk_Click()
    Dim lAchou As Boolean
    Dim lCodPlanoVida As Integer
    Dim lInd As Integer

    If cboCobertura.ListIndex > 0 Then
        'Validações
        If Val(txtQtdeVidas.Text) = 0 Then
            MsgBox "Por favor, preencher corretamente o campo Qtde de Vidas!", vbExclamation
            txtQtdeVidas.SetFocus
            Exit Sub
        End If

        If Val(txtLMGPorVida.Text) = 0 Then
            MsgBox "Por favor, preencher corretamente o campo LMG/Vida!", vbExclamation
            txtLMGPorVida.SetFocus
            Exit Sub
        End If

        If Val(Replace(Replace(txtGarMorte.Text, ".", Empty), ".", ",")) > 9999999999# Then
            MsgBox "Por favor, o valor da garantia de morte não pode ser superior a 9.999.999.999,99!", vbExclamation
            txtLMGPorVida.SetFocus
            Exit Sub
        End If

        lCodPlanoVida = cboCobertura.ItemData(cboCobertura.ListIndex)
        frmT46V101A.flxCobertura.Rows = 1
        Call gpLimPremios
        Call frmT46V117A.mpIncluirCoberturasCapitalGlobal
        Me.Hide

    Else
        gCarga = True
        frmT46V101A.optVida(1).Value = True
        gCarga = False
        Call frmT46V117A.mpExcluirCoberturasCapitalGlobal
        Unload Me
    End If

End Sub

Private Sub optAuxilioFuneral_Click(Index As Integer)
    Call mpCalculaGarantias
End Sub

Private Sub optInclusaoDeConjuge_Click(Index As Integer)
    Call mpCalculaGarantias
End Sub

Private Sub txtLMGPorVida_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

    If KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then Exit Sub

    If Val(txtLMGPorVida.Text) >= 1000000000# Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLMGPorVida_LostFocus()
    txtLMGPorVida.Text = Format(txtLMGPorVida.Text, "#,###,###,##0.00")
    Call mpCalculaGarantias
End Sub

Private Sub txtQtdeVidas_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Public Sub mpCalculaGarantias()
    txtGarMorte.Text = "0,00"
    If InStr(UCase(Trim(cboCobertura.Text)), "MORTE") Then
        If gM46V999.gfPreenchido(txtLMGPorVida.Text) And gM46V999.gfPreenchido(txtQtdeVidas.Text) Then
            txtGarMorte.Text = Format(txtLMGPorVida.Text * txtQtdeVidas.Text, "#,##0.00")
        End If
    End If

    txtGarIPA.Text = "0,00"
    If InStr(UCase(Trim(cboCobertura.Text)), "IPA") Then
        If gM46V999.gfPreenchido(txtLMGPorVida.Text) And gM46V999.gfPreenchido(txtQtdeVidas.Text) Then
            txtGarIPA.Text = Format(txtLMGPorVida.Text * txtQtdeVidas.Text, "#,##0.00")
        End If
    End If

    txtGarIEA.Text = "0,00"
    If InStr(UCase(Trim(cboCobertura.Text)), "IEA") Then
        If gM46V999.gfPreenchido(txtLMGPorVida.Text) And gM46V999.gfPreenchido(txtQtdeVidas.Text) Then
            txtGarIEA.Text = Format(txtLMGPorVida.Text * txtQtdeVidas.Text, "#,##0.00")
        End If
    End If

    txtGarIFPD.Text = "0,00"
    If InStr(UCase(Trim(cboCobertura.Text)), "IFPD") Then
        If gM46V999.gfPreenchido(txtLMGPorVida.Text) And gM46V999.gfPreenchido(txtQtdeVidas.Text) Then
            txtGarIFPD.Text = Format(txtLMGPorVida.Text * txtQtdeVidas.Text, "#,##0.00")
        End If
    End If

    txtLMGFuneralPorVida.Text = "0,00"
    txtGarAuxFuneral.Text = "0,00"
    If optAuxilioFuneral(0).Value = True Then
        txtLMGFuneralPorVida.Text = Format(lLMGFuneralPorVida, "#,##0.00")
        If gM46V999.gfPreenchido(txtQtdeVidas.Text) Then
            txtGarAuxFuneral.Text = Format(lLMGFuneralPorVida * txtQtdeVidas.Text, "#,##0.00")
        End If
    End If

    txtGarMorteConjuge = "0,00"
    If optInclusaoDeConjuge(0).Value = True Then
        If gM46V999.gfPreenchido(txtLMGPorVida.Text) And gM46V999.gfPreenchido(txtQtdeVidas.Text) Then
            txtGarMorteConjuge.Text = Format(txtLMGPorVida.Text * txtQtdeVidas.Text * 0.5, "#,##0.00")
        End If
    End If

    txtAuxCestaBasPorVida.Text = "0,00"
    txtGarAuxCestaBasica.Text = "0,00"
    If UCase(Trim(cboCobertura.Text)) = "MORTE/IPA/IEA/IFPD" Then
        txtAuxCestaBasPorVida.Text = Format(lAuxCestaBasPorVida, "#,##0.00")
        If gM46V999.gfPreenchido(txtQtdeVidas.Text) Then
            txtGarAuxCestaBasica.Text = Format(lAuxCestaBasPorVida * txtQtdeVidas.Text, "#,##0.00")
        End If
    End If
End Sub

Private Function mfConsulta_TBM_PLANO_VIDA(pAgrupar As Boolean, pTipPlanoVida As Integer, Optional pPesquisaCodCobert As Boolean, Optional pCodPlanoVida As Integer) As Boolean
    Dim lSelect As String
    Dim pMensagem As String

    mfConsulta_TBM_PLANO_VIDA = False

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, mbdTbm_Plano_Vida, pMensagem) = False Then
        pMensagem = "Erro ao abrir base de dados (TBM_PLANO_VIDA)."
        Call gM46V999.gpGraLog(2, pMensagem)
        End
    End If

    lSelect = "SELECT COD_PLANO_VIDA"
    lSelect = lSelect & ", DSC_PLANO_VIDA"
    lSelect = lSelect & ", VAL_LMG_PADRAO"
    lSelect = lSelect & IIf(pPesquisaCodCobert = False, "", ", COD_COBERT")
    lSelect = lSelect & " FROM TBM_PLANO_VIDA"
    lSelect = lSelect & " WHERE COD_RAMO = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " AND COD_MODALIDADE = " & gM46V111.gCod_Plano
    lSelect = lSelect & " AND TIP_PLANO_VIDA = " & pTipPlanoVida
    lSelect = lSelect & " AND COD_TARIF = " & gM46V999.gCodProduto
    lSelect = lSelect & " AND COD_SEQUENCIA = " & gM46V999.gCodSequencia
    lSelect = lSelect & IIf(pCodPlanoVida = 0, "", " AND COD_PLANO_VIDA = " & pCodPlanoVida)
    lSelect = lSelect & IIf(pAgrupar = False, "", " GROUP BY COD_PLANO_VIDA, DSC_PLANO_VIDA, VAL_LMG_PADRAO")

    If gM46V999.gfObtRegistro(mbdTbm_Plano_Vida, lSelect, mrsTbm_Plano_Vida, pMensagem) = False Then
        Call gM46V999.gpFecha2(mbdTbm_Plano_Vida)
        Call gM46V999.gpGraLog(2, pMensagem)
        Exit Function
    End If

    Call gM46V999.gpFecha2(mbdTbm_Plano_Vida)

    mfConsulta_TBM_PLANO_VIDA = True
End Function

Private Sub txtQtdeVidas_LostFocus()
    Call mpCalculaGarantias
End Sub

