VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V112A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo de emiss�o"
   ClientHeight    =   3870
   ClientLeft      =   2100
   ClientTop       =   2430
   ClientWidth     =   10935
   ControlBox      =   0   'False
   Icon            =   "T46V112A.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnvio 
      Caption         =   "Envio de Documento para: "
      ForeColor       =   &H00800000&
      Height          =   660
      Left            =   8640
      TabIndex        =   23
      Top             =   1680
      Width           =   2160
      Begin VB.ComboBox cboTipEnvio 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1890
      End
   End
   Begin VB.OptionButton optTipEmissao 
      Caption         =   "&Endosso"
      ForeColor       =   &H00800000&
      Height          =   600
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2445
      Width           =   1035
   End
   Begin VB.OptionButton optTipEmissao 
      Caption         =   "Renova��o de &seguro de outra seguradora (cong�nere)"
      ForeColor       =   &H00800000&
      Height          =   960
      Index           =   2
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Renova��o de cong�nere."
      Top             =   1440
      Width           =   1965
   End
   Begin VB.OptionButton optTipEmissao 
      Caption         =   "Renova��o de seguro &Yasuda"
      ForeColor       =   &H00800000&
      Height          =   600
      Index           =   1
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Renova��o Yasuda."
      Top             =   720
      Width           =   1995
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   420
      Left            =   3120
      TabIndex        =   12
      ToolTipText     =   "Volta � proposta."
      Top             =   3300
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdEndosso 
      Height          =   420
      Left            =   8700
      TabIndex        =   13
      ToolTipText     =   "Volta � proposta."
      Top             =   3300
      Width           =   2010
      _Version        =   65536
      _ExtentX        =   3545
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Texto para Endosso"
      ForeColor       =   8388608
   End
   Begin VB.Frame frmCodReg 
      Caption         =   "C�digo de Registro"
      ForeColor       =   &H00800000&
      Height          =   660
      Left            =   2250
      TabIndex        =   16
      Top             =   1680
      Width           =   6240
      Begin VB.ComboBox cboCodReg 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame fraSeguradora 
      Caption         =   "Seguradora: "
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   2280
      TabIndex        =   22
      Top             =   30
      Width           =   8520
      Begin VB.ComboBox cboSeguradora 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V112A.frx":000C
         Left            =   120
         List            =   "T46V112A.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Nome da seguradora cong�nere."
         Top             =   240
         Width           =   8265
      End
   End
   Begin VB.Frame fraApolice 
      Caption         =   "Dados da Ap�lice"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   2250
      TabIndex        =   17
      Top             =   825
      Width           =   8520
      Begin VB.TextBox txtRenItem 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3135
         MaxLength       =   5
         TabIndex        =   6
         ToolTipText     =   "Item da ap�lice de cong�nere."
         Top             =   240
         Width           =   705
      End
      Begin VB.TextBox txtRenApolice 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   5
         ToolTipText     =   "N�mero da cota��o."
         Top             =   240
         Width           =   1680
      End
      Begin Threed.SSPanel panX 
         Height          =   300
         Index           =   1
         Left            =   4800
         TabIndex        =   18
         Top             =   240
         Width           =   1440
         _Version        =   65536
         _ExtentX        =   2540
         _ExtentY        =   529
         _StockProps     =   15
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   0
         BevelInner      =   1
         Alignment       =   6
         Begin VB.OptionButton optSinistroCon 
            Caption         =   "&Sim"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "SIM, tem sinistro."
            Top             =   40
            Width           =   600
         End
         Begin VB.OptionButton optSinistroCon 
            Caption         =   "&N�o"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   690
            TabIndex        =   8
            ToolTipText     =   "N�O tem sinistro."
            Top             =   40
            Value           =   -1  'True
            Width           =   600
         End
      End
      Begin Threed.SSCommand cmdBusca 
         Height          =   300
         Left            =   6360
         TabIndex        =   9
         ToolTipText     =   "Busca os dados da ap�lice."
         Top             =   240
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Busca dados de ap�lice"
         ForeColor       =   8388608
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   21
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sinistro? "
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   7
         Left            =   3960
         TabIndex        =   19
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ap�lice: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame frmTipoEndosso 
      Caption         =   "Tipo de endosso:"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   2250
      TabIndex        =   14
      Top             =   2475
      Width           =   8520
      Begin VB.ComboBox cboDizer 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   300
         Width           =   8055
      End
   End
   Begin VB.OptionButton optTipEmissao 
      Caption         =   "Seguro &novo"
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Seguro novo."
      Top             =   150
      Width           =   1395
   End
   Begin Threed.SSCommand cmdCancelar 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   4725
      TabIndex        =   24
      ToolTipText     =   "Volta � proposta."
      Top             =   3300
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V112A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTipEmissaoAnt As e_TipEmissao
Private mTipEmissao As e_TipEmissao
Private mCodDizerAnt As e_CodDizer

Private Sub cboConSeguradora_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Sub PosicionaObj()
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then

        frmCodReg.Visible = False
        frmTipoEndosso.Visible = False
        optTipEmissao(3).Visible = False

        cboTipEnvio.Visible = True
        fraEnvio.Visible = True

    End If
End Sub


Private Sub cboCodReg_Change()
    cboCodReg_Click
End Sub

Private Sub cboCodReg_Click()
'1. Habilita op��o de salvar.
    If gCarga = True Then Exit Sub
    Call gpLimPremios
End Sub

Private Sub cboDizer_Change()
    cboDizer_Click
End Sub

Private Sub cboDizer_Click()
'1. Habilita op��o de salvar.
    If gCarga = True Then Exit Sub
    Call gpLimPremios
    gM46V111.gCod_Dizer = cboDizer.ItemData(cboDizer.ListIndex)
End Sub

Private Sub cboSeguradora_Change()
    Call gpHabSalvar
End Sub

Private Sub cboTipEnvio_Change()
    cboTipEnvio_Click
End Sub

Private Sub cboTipEnvio_Click()
'1. Habilita op��o de salvar.
    If gCarga = True Then Exit Sub
    Call gpLimPremios
End Sub

Private Sub cmdBusca_Click()
    Dim lbdP00APOL As ADODB.Connection    'Arquivo - P00APOL.
    Dim lBonus As Long
    Dim lCodSeguradora As Integer     'C�digo da seguradora.
    Dim lIndSinistro As Byte      'Indicador de sinistro.
    Dim lNumItem As Integer    'N�mero do item.
    Dim lMensagem As String    'Mensagem.
    Dim rsAPOL As clsYasRecordSet    'Registro: P00APOL.
    Dim lI As Long    'Auxiliar.
    Dim lobjstcClau As stcA46V703B    'Objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU)
    Dim ColecaoClau As Collection    'Cole��o de cl�usulas.
    Dim lobjstcCobert As stcA46V704B    'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
    Dim ColecaoCobert As Collection    'Cole��o de coberturas.
    Dim lRegistro As String    'Registro.
    Dim lIdeTabela As String
    Dim lSelect As String
    If Not txtRenApolice.Enabled Then Exit Sub

    Screen.MousePointer = vbHourglass

    '1. Confirma com usu�rio.
    lMensagem = "AGE0018 - Confirma busca de dados de ap�lice? A busca de dados de ap�lice " & _
                "desconsiderar� os dados existentes."
    If MsgBox(lMensagem, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Call gM46V999.gpGraLog(0, lMensagem & " N�o.")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Call gM46V999.gpGraLog(0, lMensagem & " Sim.")

    '2. Consiste n�mero de ap�lice anterior.
    If Not gM46V999.gfPreenchido(txtRenApolice.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                                  "N�mero da ap�lice inv�lido (renova��o Yasuda - n�o preenchido).")
        If txtRenApolice.Visible = True And txtRenApolice.Enabled = True Then
            txtRenApolice.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Not IsNumeric(txtRenApolice.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                                  "N�mero da ap�lice inv�lido (renova��o Yasuda - n�o num�rico).")
        If txtRenApolice.Visible = True And txtRenApolice.Enabled = True Then
            txtRenApolice.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If val(txtRenApolice.Text) = 0 Then
        Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                                  "N�mero da ap�lice inv�lido (renova��o Yasuda - zerado).")
        If txtRenApolice.Visible = True And txtRenApolice.Enabled = True Then
            txtRenApolice.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Len(Trim(txtRenApolice.Text)) < 10 Then
        txtRenApolice.Text = Trim$(txtRenApolice.Text)
        If Len(txtRenApolice.Text) < 10 Then
            txtRenApolice.Text = String(10 - Len(txtRenApolice.Text), "0") & txtRenApolice.Text
        End If
    End If

    '3. Consiste n�mero do item da ap�lice.
    If Not gM46V999.gfPreenchido(txtRenItem.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                                  "N�mero do item da ap�lice inv�lido (renova��o Yasuda - n�o preenchido).")
        If txtRenItem.Visible = True And txtRenItem.Enabled = True Then
            txtRenItem.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Not IsNumeric(txtRenItem.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                                  "N�mero do item da ap�lice inv�lido (renova��o Yasuda - n�o num�rico).")
        If txtRenItem.Visible = True And txtRenItem.Enabled = True Then
            txtRenItem.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If val(txtRenItem.Text) = 0 Then
        Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                                  "N�mero do item da ap�lice inv�lido (renova��o Yasuda - zerado).")
        If txtRenItem.Visible = True And txtRenItem.Enabled = True Then
            txtRenItem.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Len(Trim(txtRenItem.Text)) < 5 Then
        txtRenItem.Text = Trim$(txtRenItem.Text)
        If Len(txtRenItem.Text) < 5 Then
            txtRenItem.Text = String(5 - Len(txtRenItem.Text), "0") & txtRenItem.Text
        End If
    End If

    '4. Obt�m ap�lice/item.
    If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, lbdP00APOL, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'If gM46V999.gBaseLocal = True Then
    lIdeTabela = "TAB_APOL_LOC"
    'Else
    '    lIdeTabela = "GED.DBO.TAB_PED_LOC"
    'End If

    lSelect = "SELECT * FROM " & lIdeTabela & _
            " WHERE " & lIdeTabela & ".Num_Apol = " & txtRenApolice.Text & _
            " AND " & lIdeTabela & ".Num_Item = " & txtRenItem.Text

    If gM46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdP00APOL)
        Call gM46V999.gpGraLog(1, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If rsAPOL.EOF = True Then
        'If gM46V999.gBaseLocal = True Then
        lIdeTabela = "TAB_APOL"
        'Else
        '    lIdeTabela = "GED.DBO.TAB_PED"
        'End If

        lSelect = "SELECT * from " & lIdeTabela & _
                " WHERE " & lIdeTabela & ".Num_Apol = " & txtRenApolice.Text
        If gM46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdP00APOL)
            Call gM46V999.gpGraLog(1, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If rsAPOL.EOF = True Then
            lMensagem = "FNL0008 - Ap�lice " & txtRenApolice.Text & " n�o localizada."
        Else
            lMensagem = "FNL0009 - Item  " & txtRenItem.Text & " n�o localizado na ap�lice " & _
                        txtRenApolice.Text & "."
        End If

        'proposta eletr�nica
        If IsNumeric(rsAPOL("NUM_CGC_CPF")) Then
            mAntCNPJ = Format(val(rsAPOL("NUM_CGC_CPF")), "00000000000000")
        Else
            mAntCNPJ = Right$(String(14, "0") & Trim(rsAPOL("NUM_CGC_CPF")), 14)
        End If
        mAntSeg = rsAPOL("NOM_SEGURADO")

        Call gM46V999.gpFecha2(lbdP00APOL)
        Call gM46V999.gpGraLog(1, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    gCarga = True
    frmT46V101A.txtRisTipLogradouro.Text = gM46V999.gfTratarNulo(rsAPOL("Cod_Tip_Logr_Risco"), tpoString)
    frmT46V101A.txtRisNomLogradouro.Text = gM46V999.gfTratarNulo(rsAPOL("Nom_Logr_Risco"), tpoString)
    frmT46V101A.txtRisComplemento.Text = gM46V999.gfTratarNulo(rsAPOL("Dsc_Compl_Risco"), tpoString)
    frmT46V101A.txtRisBairro.Text = gM46V999.gfTratarNulo(rsAPOL("Nom_Bairro_Risco"), tpoString)
    frmT46V101A.txtRisCidade.Text = gM46V999.gfTratarNulo(rsAPOL("Nom_Cid_Risco"), tpoString)
    frmT46V101A.cboRisUF.ListIndex = -1
    If gM46V999.gfPreenchido(gM46V999.gfTratarNulo(rsAPOL("Sig_UF_Risco"), tpoString)) Then
        For lI = 0 To (frmT46V101A.cboRisUF.ListCount - 1)
            If gM46V999.gfTratarNulo(rsAPOL("Sig_UF_Risco"), tpoString) = frmT46V101A.cboRisUF.List(lI) Then
                frmT46V101A.cboRisUF.ListIndex = lI
                Exit For
            End If
        Next lI
    End If
    frmT46V101A.txtRisCEP.Text = gM46V999.gfTratarNulo(rsAPOL("Num_CEP_Risco"), tpoNumerico)
    frmT46V101A.txtDescProgr.Text = Format(gM46V999.gfTratarNulo(rsAPOL("Per_Desc_Progres"), TpoValor), "standard")
    frmT46V101A.txtISBas.Text = Format(gM46V999.gfTratarNulo(rsAPOL("Per_IS_Vr"), TpoValor), "standard")
    frmT46V101A.txtCoefISVR.Text = Format(gM46V999.gfTratarNulo(rsAPOL("Val_Coef_IS_Vr"), TpoValor), "standard")
    frmT46V101A.txtValoRisco.Text = Format(gM46V999.gfTratarNulo(rsAPOL("Val_Vr"), TpoValor), "standard")
    frmT46V101A.txtClassificacao.Tag = gM46V999.gfTratarNulo(rsAPOL("Cod_Clasf"), tpoNumerico) & gM46V999.gfTratarNulo(rsAPOL("Cod_Compl_Clasf"), tpoNumerico) & Format(gM46V999.gfTratarNulo(rsAPOL("Tab_Clasf"), tpoNumerico), "0000")
    frmT46V101A.txtClassificacao.Text = gM46V999.gfTratarNulo(rsAPOL("Cod_Clasf"), tpoNumerico) & "-" & gM46V999.gfTratarNulo(rsAPOL("Dsc_Clasf"), tpoString)
    'Ficha 1000305
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
        Select Case gM46V999.gfTratarNulo(rsAPOL("Cod_Constr"), tpoNumerico)
        Case e_CodContr_Superior
            frmT46V101A.cboConstrucao.ListIndex = 1
        Case e_CodContr_Inferior
            frmT46V101A.cboConstrucao.ListIndex = 2
        Case e_CodContr_Mista
            frmT46V101A.cboConstrucao.ListIndex = 3
        Case Else
            frmT46V101A.cboConstrucao.ListIndex = 1
        End Select
    Else
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Constr"), tpoNumerico) = e_CodContr_Superior Then
            frmT46V101A.optContrucao(0).Value = True
        Else
            frmT46V101A.optContrucao(1).Value = True
        End If
    End If

    'Ficha 1000305
    '== Tipo Cobertura
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
        Select Case gM46V999.gfTratarNulo(rsAPOL("Tip_Cobert"), tpoNumerico)
        Case e_TipCobert_PredioeConteudo
            frmT46V101A.cboTipCobert.ListIndex = 0
        Case e_TipCobert_Predio
            frmT46V101A.cboTipCobert.ListIndex = 1
        Case e_TipCobert_Conteudo
            frmT46V101A.cboTipCobert.ListIndex = 2
        Case Else
            frmT46V101A.cboTipCobert.ListIndex = 0
        End Select
    Else
        frmT46V101A.cboTipCobert.ListIndex = 0
    End If
    lBonus = gM46V999.gfTratarNulo(rsAPOL("Per_Bonus"), tpoNumerico)
    ''thais sinistro
    If gM46V999.gfTratarNulo(rsAPOL("Ind_Sinistro"), tpoNumerico) Then    'Com sinistro
        lBonus = lBonus - 5
    Else
        lBonus = lBonus + 5
    End If
    If lBonus < 0 Then lBonus = 0


    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
        If gM46V999.gCodProduto >= 942 Then
            'Soma mais 5 da fidelidade.
            lBonus = lBonus + 5
        End If
        'If lBonus > 30 Then lBonus = 30
        If lBonus > 25 Then lBonus = 25
    ElseIf gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
        If gM46V999.gCodProduto >= 946 Then
            'Soma mais 5 da fidelidade.
            lBonus = lBonus + 5
            If lBonus > 25 Then lBonus = 25
            'If lBonus > 30 Then lBonus = 30
        Else
            If lBonus > 25 Then lBonus = 25
        End If
    Else
        If lBonus > 25 Then lBonus = 25
    End If

    'Ficha 901673 e 901772  'Ficha 1000099 - Planos 2131 a 2138 Corretor Schroeder
    If (gM46V111.gCod_plano >= 2001 And gM46V111.gCod_plano <= 2031) Or (gM46V111.gCod_plano >= 2051 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2138) Then
        lBonus = 0
    End If

    'Ser� carregado no c�lculo
    frmT46V101A.txtDescBonus.Text = lBonus

    frmT46V101A.txtQuaEmpregados.Text = gM46V999.gfTratarNulo(rsAPOL("Qtd_Empreg"), tpoNumerico)

    'Seguran�a
    If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_01"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_02"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_03"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_04"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_05"), tpoString) = "S" Then
        frmT46V101A.optSeguranca(0).Value = True
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_01"), tpoString) = "S" Then
            frmT46V128A.chkSeguranca(0).Value = vbChecked
        Else
            frmT46V128A.chkSeguranca(0).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_02"), tpoString) = "S" Then
            frmT46V128A.chkSeguranca(1).Value = vbChecked
        Else
            frmT46V128A.chkSeguranca(1).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_03"), tpoString) = "S" Then
            frmT46V128A.chkSeguranca(2).Value = vbChecked
        Else
            frmT46V128A.chkSeguranca(2).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_04"), tpoString) = "S" Then
            frmT46V128A.chkSeguranca(3).Value = vbChecked
        Else
            frmT46V128A.chkSeguranca(3).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_05"), tpoString) = "S" Then
            frmT46V128A.chkSeguranca(4).Value = vbChecked
        Else
            frmT46V128A.chkSeguranca(4).Value = vbUnchecked
        End If
    Else
        frmT46V101A.optSeguranca(1).Value = True
        Unload frmT46V128A
    End If

    If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_01"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_02"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_03"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_04"), tpoString) = "S" Then
        frmT46V101A.optProtecao(0).Value = True
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_01"), tpoString) = "S" Then
            frmT46V126A.chkProtecao(0).Value = vbChecked
        Else
            frmT46V126A.chkProtecao(0).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_02"), tpoString) = "S" Then
            frmT46V126A.chkProtecao(1).Value = vbChecked
        Else
            frmT46V126A.chkProtecao(1).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_03"), tpoString) = "S" Then
            frmT46V126A.chkProtecao(2).Value = vbChecked
        Else
            frmT46V126A.chkProtecao(2).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_04"), tpoString) = "S" Then
            frmT46V126A.chkProtecao(3).Value = vbChecked
        Else
            frmT46V126A.chkProtecao(3).Value = vbUnchecked
        End If
    Else
        frmT46V101A.optProtecao(1).Value = True
        Unload frmT46V126A
    End If

    frmT46V101A.txtRisRegiao.Tag = Format(gM46V999.gfTratarNulo(rsAPOL("Cod_Regiao"), tpoNumerico), "00000") & Format(gM46V999.gfTratarNulo(rsAPOL("Cod_Produto"), tpoNumerico), "00000")
    gM46V111.gCod_Regiao = val(Right(frmT46V101A.txtRisRegiao.Tag, 4))
    frmT46V101A.objstcPedLoc.Cod_Regiao = gM46V999.gfTratarNulo(rsAPOL("Cod_Regiao"), tpoNumerico)
    frmT46V101A.objstcPedLoc.Cod_Produto = gM46V999.gfTratarNulo(rsAPOL("Cod_Produto"), tpoNumerico)
    frmT46V101A.txtRisRegiao.Text = frmT46V101A.objstcPedLoc.gDesc_Regiao

    'Clausulas
    'If gM46V999.gBaseLocal = True Then
    lIdeTabela = "TAB_APOL_CLAU"
    'Else
    '    lIdeTabela = "GED.DBO.TAB_PED_CLAU"
    'End If

    lSelect = " Select * from " & lIdeTabela & " Where Num_Apol = " & rsAPOL("Num_Apol")
    lSelect = lSelect & " and num_item = " & rsAPOL("Num_Item")

    If Not gM46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) Then
        Exit Sub
    End If
    'Clausula do benefici�rio.
    If Not rsAPOL.EOF Then
        If rsAPOL("Cod_Clau") = 4 Or rsAPOL("Cod_Clau") = 1 Then
            frmT46V101A.optBeneficiario(0).Value = True
        Else
            frmT46V101A.optBeneficiario(1).Value = True
        End If
    End If

    'If gM46V999.gBaseLocal = True Then
    lIdeTabela = "TAB_APOL_COBERT"
    'Else
    '    lIdeTabela = "GED.DBO.TAB_PED_COBERT"
    'End If

    'Coberturas
    lSelect = " Select * from " & lIdeTabela & " Where Num_Apol = " & txtRenApolice.Text
    lSelect = lSelect & " and num_item = " & txtRenItem.Text

    If Not gM46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) Then
        Exit Sub
    End If
    Set ColecaoCobert = New Collection
    Do While Not rsAPOL.EOF
        Set lobjstcCobert = New stcA46V704B
        Set lobjstcCobert.M46V999 = gM46V999

        With lobjstcCobert
            .Cod_Cobert = gM46V999.gfTratarNulo(rsAPOL("Cod_Cobert"), tpoNumerico)
            .Per_Desc_Coml = gM46V999.gfTratarNulo(rsAPOL("Per_Desc_Coml"), tpoNumerico)
            .Per_Desc_Exp = gM46V999.gfTratarNulo(rsAPOL("Per_Desc_Exp"), tpoNumerico)
        End With
        ColecaoCobert.Add lobjstcCobert, lobjstcCobert.Class_ID
        rsAPOL.MoveNext
    Loop
    'Cobertura
    Unload frmT46V117A
    If ColecaoCobert.Count > 0 Then
        'Call frmT46V117A.mpSelecionarCoberturas(ColecaoCobert)
        gM46V999.gChamaSelecionarCoberturas = True
    End If

    Set ColecaoClau = Nothing
    Set ColecaoCobert = Nothing
    gCarga = False
    Call gpHabSalvar
    gAltT46V112A = True

    '7. Reposiciona informa��es do item e retorna ao form de item.
    Me.Hide
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancelar_Click()
    Me.Hide
End Sub

Private Sub cmdEndosso_Click()
    frmT46V148A.Show vbModal, Me
End Sub

Private Sub cmdOk_Click()
    Dim lTipEmissao As Integer      'Tipo de emiss�o.
    Dim lI As Integer    'For...next.
    Dim lOpcao As Integer    'Op��o.


'1. Posiciona tipo de emiss�o.
    If optTipEmissao(0) = True Then         'Seguro novo.
        lTipEmissao = 100
    ElseIf optTipEmissao(1) = True Then     'Renova��o Yasuda.
        lTipEmissao = 200
    ElseIf optTipEmissao(2) = True Then     'Renova��o de cong�nere.
        lTipEmissao = 101
    ElseIf optTipEmissao(3) = True Then     'Endosso
        lTipEmissao = 300
    Else
        lTipEmissao = 0
    End If

    frmT46V101A.txtSegurado.Enabled = True
    frmT46V102A.txtSegurado.Enabled = True

    '2. Consiste dados.
    Select Case lTipEmissao
    Case 100    'Seguro novo.
        'Nada a consistir.
        txtRenApolice.Text = ""
        txtRenItem.Text = ""
        cboSeguradora.ListIndex = -1
        cboDizer.ListIndex = -1
    Case 101    'Renova��o de cong�nere.
        If Not gM46V999.gfPreenchido(cboSeguradora.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0071 - " & gM46V999.gPrefixo & _
                                      "Escolher uma seguradora cong�nere! (Renova��o de cong�nere).")
            If cboSeguradora.Visible And cboSeguradora.Enabled Then
                cboSeguradora.SetFocus
            End If
            Exit Sub
        End If
    Case 200    'Renova��o de seguro Yasuda.
        'Proposta Eletr�nica
        If gM46V999.gNovaTransmissaoMulti = True Then
            frmT46V101A.txtSegurado.Enabled = False
            frmT46V102A.txtSegurado.Enabled = False
        End If
        If Not gM46V999.gfPreenchido(txtRenApolice.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                                      "N�mero da ap�lice inv�lido (renova��o Yasuda - n�o preenchido).")
            If txtRenApolice.Visible And txtRenApolice.Enabled Then
                txtRenApolice.SetFocus
            End If
            Exit Sub
        End If
        If Not IsNumeric(txtRenApolice.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                                      "N�mero da ap�lice inv�lido (renova��o Yasuda - n�o num�rico).")
            If txtRenApolice.Visible And txtRenApolice.Enabled Then
                txtRenApolice.SetFocus
            End If
            Exit Sub
        End If
        If val(txtRenApolice.Text) = 0 Then
            Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                                      "N�mero da ap�lice inv�lido (renova��o Yasuda - zerado).")
            If txtRenApolice.Visible And txtRenApolice.Enabled Then
                txtRenApolice.SetFocus
            End If
            Exit Sub
        End If
        If Not gM46V999.gfPreenchido(txtRenItem.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & "N�mero do item da ap�lice inv�lido " & _
                                      "(renova��o Yasuda - n�o preenchido).")
            If txtRenItem.Visible And txtRenItem.Enabled Then
                txtRenItem.SetFocus
            End If
            Exit Sub
        End If
        If Not IsNumeric(txtRenItem.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                                      "N�mero do item da ap�lice inv�lido (renova��o Yasuda - n�o num�rico).")
            If txtRenItem.Visible And txtRenItem.Enabled Then
                txtRenItem.SetFocus
            End If
            Exit Sub
        End If
        If val(txtRenItem.Text) = 0 Then
            Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                                      "N�mero do item da ap�lice inv�lido (renova��o Yasuda - zerado).")
            If txtRenItem.Visible And txtRenItem.Enabled Then
                txtRenItem.SetFocus
            End If
            Exit Sub
        End If
    Case 300    'Endosso
        If Not gM46V999.gfPreenchido(cboDizer.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0071 - " & gM46V999.gPrefixo & _
                                      "Escolher o c�digo do dizer correto! (Endosso Yasuda).")
            If cboDizer.Visible And cboDizer.Enabled Then
                cboDizer.SetFocus
            End If
            Exit Sub
        End If
        If Not gM46V999.gfPreenchido(txtRenApolice.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0071 - " & gM46V999.gPrefixo & _
                                      "N�mero da ap�lice inv�lido (Endosso Yasuda - n�o preenchido).")
            If txtRenApolice.Visible And txtRenApolice.Enabled Then
                txtRenApolice.SetFocus
            End If
            Exit Sub
        End If
        If Not IsNumeric(txtRenApolice.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0071 - " & gM46V999.gPrefixo & _
                                      "N�mero da ap�lice inv�lido (Endosso Yasuda - n�o num�rico).")
            If txtRenApolice.Visible And txtRenApolice.Enabled Then
                txtRenApolice.SetFocus
            End If
            Exit Sub
        End If
        If val(txtRenApolice.Text) = 0 Then
            Call gM46V999.gpGraLog(1, "FGE0071 - " & gM46V999.gPrefixo & _
                                      "N�mero da ap�lice inv�lido (Endosso Yasuda - zerado).")
            If txtRenApolice.Visible And txtRenApolice.Enabled Then
                txtRenApolice.SetFocus
            End If
            Exit Sub
        End If
        If cboDizer.ItemData(cboDizer.ListIndex) <> e_CodDizerInclusao Then
            If Not gM46V999.gfPreenchido(txtRenItem.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0072 - " & gM46V999.gPrefixo & "Endosso do item da ap�lice inv�lido " & _
                                          "(renova��o Yasuda - n�o preenchido).")
                If txtRenItem.Visible = True And txtRenItem.Enabled = True Then
                    txtRenItem.SetFocus
                End If
                Exit Sub
            End If
            If Not IsNumeric(txtRenItem.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0072 - " & gM46V999.gPrefixo & _
                                          "N�mero do item da ap�lice inv�lido (Endosso Yasuda - n�o num�rico).")
                If txtRenItem.Visible = True And txtRenItem.Enabled = True Then
                    txtRenItem.SetFocus
                End If
                Exit Sub
            End If
            If val(txtRenItem.Text) = 0 Then
                Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                                          "N�mero do item da ap�lice inv�lido (Endosso Yasuda - zerado).")
                If txtRenItem.Visible = True And txtRenItem.Enabled = True Then
                    txtRenItem.SetFocus
                End If
                Exit Sub
            End If
        End If
        gM46V111.gCod_Dizer = cboDizer.ItemData(cboDizer.ListIndex)
    End Select

    If mCodDizerAnt <> gM46V111.gCod_Dizer Then
        frmT46V101A.mpHabEndosso
        mCodDizerAnt = gM46V111.gCod_Dizer
    End If
    
    '' 2016-08-24: Selecionar SISCOTA se for Endosso ou Renovacao Sompo. Nao da para pegar pelo objStc.
    'If lTipEmissao = 200 Or lTipEmissao = 300 Then
    '    If frmT46V102A.CmbTipCotac.Enabled Then
    '       frmT46V102A.CmbTipCotac.ListIndex = 2
    '    End If
    'End If
    
    'Acesso Syas Privil�gios
    If Not gM46V999.M46O001.PermiteQQDescontoComercial = True Then
        If gM46V111.gCod_plano < 2000 Then frmT46V101A.txtDescComercial.Text = "0,00"
        If gM46V111.gCod_plano < 2000 Then frmT46V101A.txtDescComercial.Enabled = False
    End If
    
    '3. Torna form invis�vel.
    Me.Hide
End Sub

Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    Me.Caption = " Tipo de Emiss�o  -  PI: " & frmT46V101A.objstcPedido.NUM_PI
    PosicionaObj
Finalizar:
    If gM46V999.gDesabilitarForms Or gM46V999.gDesabFormsCotacao Then
        If Me.Tag <> "Desabilitado" Then
            DesabilitaControles Me
            cmdOk.Enabled = True
            If gM46V999.gDesabFormsCotacao = True And gM46V999.gTipoAcesso = DIGITACAO Then
                cboTipEnvio.Enabled = True
                cboCodReg.Enabled = True
            End If
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
        End If
    End If
    
    Dim lTipEmissao As Integer
    
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        If gM46V999.gDesabFormsTransmi = True Then
            
            For lTipEmissao = 0 To optTipEmissao.Count - 1
                If optTipEmissao(lTipEmissao).Value Then
                    Exit For
                End If
            Next
                
            DesabilitaControles Me
            optTipEmissao(lTipEmissao).Value = True
            
            cmdOk.Enabled = True
            cmdEndosso.Enabled = True
            If gM46V999.gTipoAcesso = DIGITACAO Then
                cboTipEnvio.Enabled = True
                cboCodReg.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub optConSinistro_Click(index As Integer)
    Call gpLimPremios
End Sub

Private Sub optConSinistro_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Dim lbdTab_Dizer As ADODB.Connection    'Arquivo: P00Tarifa.
    Dim lMensagem As String     'Mensagem.
    Dim lrsTab_Dizer As clsYasRecordSet    'Registro : Tab_Dizer.
    Dim lLinha As Long    'Linhas.
    Dim lSelect As String
    Call gM46V999.gpCenForm(Me)  'Centraliza form.

    '1. Posiciona controles do help.
    'Inicio como Altera��o
    txtRenItem.Enabled = True
    txtRenApolice.Text = ""
    txtRenItem.Text = ""

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Dizer, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    lSelect = " Select * from tab_dizer "
    If gM46V999.gfObtRegistro(lbdTab_Dizer, lSelect, lrsTab_Dizer, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTab_Dizer)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    cboDizer.Clear
    'Do While Not lrsTab_Dizer.EOF
    '    cboDizer.AddItem Format(lrsTab_Dizer("Cod_Dizer"), "00") & " - " & lrsTab_Dizer("Dsc_Dizer")
    '    cboDizer.ItemData(cboDizer.NewIndex) = lrsTab_Dizer("Cod_Dizer")
    '    lrsTab_Dizer.MoveNext
    'Loop

    cboDizer.AddItem "Altera��es"
    cboDizer.ItemData(cboDizer.NewIndex) = 12

    cboDizer.AddItem "Inclus�o de item"
    cboDizer.ItemData(cboDizer.NewIndex) = 1

    cboDizer.AddItem "Exclus�o de item"
    cboDizer.ItemData(cboDizer.NewIndex) = 4

    cboDizer.AddItem "Cancelamento da ap�lice"
    cboDizer.ItemData(cboDizer.NewIndex) = 10

    cboDizer.ListIndex = -1
    If gM46V111.gCod_Dizer <> 0 Then
        cboDizer.ListIndex = gM46V999.gfRetornarListIndex(cboDizer, gM46V111.gCod_Dizer)
    End If

    cboCodReg.Clear

    mCodDizerAnt = gM46V111.gCod_Dizer

    cboCodReg.AddItem "001-Ap�lice com cobran�a de pr�mio"
    cboCodReg.AddItem "002-Ap�lice sem cobran�a de IOF e custo"
    cboCodReg.AddItem "003-Ap�lice sem cobran�a de IOF"
    cboCodReg.AddItem "004-Ap�lice sem cobran�a de custo"
    cboCodReg.AddItem "005-Ap�lice de Moeda Estrangeira"
    cboCodReg.AddItem "006-Ap�lice coletivo Acidentes Pessoais"
    cboCodReg.AddItem "007-Ap�lice coletivo pagamento trimestral"
    cboCodReg.AddItem "008-Ap�lice coletivo pagamento semestral"
    '8000047867, C�digo de registro
    'cboCodReg.AddItem "020-Ap�lice a averbar (sem cobran�a de pr�mio)"
    cboCodReg.AddItem "101-Endosso com cobran�a de pr�mio"
    cboCodReg.AddItem "102-Endosso sem cobran�a de IOF e custo"
    cboCodReg.AddItem "103-Endosso sem cobran�a de IOF"
    cboCodReg.AddItem "104-Endosso sem cobran�a de custo"
    cboCodReg.AddItem "105-Endosso de Moeda Estrangeira"
    cboCodReg.AddItem "106-Endosso restitui��o de pr�mio - sem cancelamento da ap�lice"
    cboCodReg.AddItem "107-Endosso restitui��o de pr�mio - com cancelamento da ap�lice"
    cboCodReg.AddItem "120-Endosso sem movimento de pr�mio"
    cboCodReg.AddItem "121-Endosso sem movimento de pr�mio - com pr�mio negativo"
    cboCodReg.AddItem "122-Endosso sem movimento de pr�mio - com pr�mio positivo"
    cboCodReg.AddItem "129-Endosso de cancelamento de ap�lice sem devolu��o de pr�mio"
    cboCodReg.AddItem "141-Endosso de cancelamento de ap�lice"
    cboCodReg.AddItem "142-Endosso de cancelamento de endosso"
    cboCodReg.AddItem "143-Endosso de cancelamento de fatura mensal"
    cboCodReg.AddItem "201-Fatura mensal com cobran�a de pr�mio"
    cboCodReg.AddItem "202-Fatura mensal sem cobran�a de pr�mio"
    cboCodReg.AddItem "203-Fatura mensal sem cobran�a de IOF"
    cboCodReg.AddItem "204-Fatura mensal sem cobran�a de custo"
    cboCodReg.AddItem "205-Fatura mensal de Moeda Estrangeira"
    cboCodReg.AddItem "400-Bilhete de seguro"
    cboCodReg.AddItem "600-Sinistros avisados - Direto"
    cboCodReg.AddItem "601-Sinistros avisados - Cosseguros Aceitos"
    cboCodReg.AddItem "610-Sinistros pagos / recuperados - Direto"
    cboCodReg.AddItem "611-Sinistros pagos / recuperados - Cosseguro Aceito"
    cboCodReg.AddItem "630-Sinistros avisados - Direto, com n�mero de sinistro"
    cboCodReg.AddItem "631-Sinistros avisados - Cosseguros, com n�mero de sinistro"
    cboCodReg.AddItem "650-Estorno de sinistros pagos - Direto"
    cboCodReg.AddItem "651-Estorno de sinistros pagos - Cosseguros Aceitos"
    cboCodReg.AddItem "800-Comiss�es"

    cboCodReg.ListIndex = -1

    cboTipEnvio.AddItem "Conforme cadastro"
    cboTipEnvio.AddItem "Segurado"
    cboTipEnvio.AddItem "Corretor"
    cboTipEnvio.ListIndex = 0

    'Call mpAjustaForm


    Call gM46V999.gpFecha2(lbdTab_Dizer)

End Sub

Private Sub optSinistroCon_Click(index As Integer)
    frmT46V101A.optSinistroYas(index).Value = optSinistroCon(index).Value
    '1. Habilita op��o de salvar.
    Call gpLimPremios
End Sub

Private Sub optSinistroYas_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub optTipEmissao_Click(index As Integer)
    Dim bdP0044700 As ADODB.Connection    'Arquivo - P0044700.
    Dim lMensagem As String         'Mensagem.
    Dim rsP0044700 As clsYasRecordSet    'Registro - P0044700.
    Dim lSelect As String
    '1. Habilita op��o de salvar.
    Call gpLimPremios

    'If gCarga = True Then Exit Sub

    '2. Posiciona quadros de informa��es e tipo de emiss�o.
    Select Case index
    Case 0      'Novo
        fraSeguradora.Visible = False
        fraApolice.Visible = False
        frmT46V101A.cmdBonus.Enabled = True
        frmT46V101A.cmdBonus.Picture = frmT46V101A.cmdExtra(1).Picture
        If (gM46V999.gCodProduto >= 932 Or gM46V999.gCodProduto = 0) And (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
            frmT46V101A.cmdBonus.Enabled = False
            frmT46V101A.txtDescBonus.Text = "0"
            frmT46V101A.cmdBonus.Picture = frmT46V101A.cmdExtra(2).Picture
        End If
        frmTipoEndosso.Visible = False
        frmT46V148A.txtTextoEndosso = ""
        cmdEndosso.Visible = False
        mTipEmissao = e_TipEmissao_SeguroNovo
    Case 1      'Yasuda
        fraSeguradora.Visible = False
        fraApolice.Visible = True
        If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
            frmT46V101A.cmdBonus.Enabled = False
            frmT46V101A.cmdBonus.Picture = frmT46V101A.cmdExtra(2).Picture
            optSinistroCon(0).Enabled = False
            optSinistroCon(1).Enabled = False
            frmT46V101A.optSinistroYas(0).Enabled = False
            frmT46V101A.optSinistroYas(1).Enabled = False

        End If

        frmTipoEndosso.Visible = False
        frmT46V148A.txtTextoEndosso = ""
        cmdEndosso.Visible = False
        mTipEmissao = e_TipEmissao_RenYasuda
        cmdBusca.Visible = True
        txtRenApolice.MaxLength = 10

    Case 2      'Congenere
        If cboSeguradora.ListCount = 0 Then
            If gM46V999.gfAbrBasDados("P0044700", gM46V999.gTipP0044700, gM46V999.gEndP0044700, bdP0044700, lMensagem) = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            lSelect = "SELECT CSeguradora, Seguradora FROM YTAB_SEGURADORA"
            If gM46V999.gfObtRegistro(bdP0044700, lSelect, rsP0044700, lMensagem) = False Then
                Call gM46V999.gpFecha2(bdP0044700)
                Call gM46V999.gpGraLog(2, lMensagem)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            If rsP0044700.EOF = False Then
                With cboSeguradora
                    While Not rsP0044700.EOF
                        Select Case val(rsP0044700!CSeguradora)
                        Case 641
                            'Yasuda - n�o incluir.
                        Case 754, 872, 883
                            'Seguradoras sem nome - inv�lidas - n�o incluir.
                        Case Else
                            .AddItem Trim(rsP0044700!Seguradora)
                            .ItemData(.NewIndex) = val(rsP0044700!CSeguradora)
                        End Select
                        rsP0044700.MoveNext
                    Wend
                    If (gM46V999.gCodProduto < 988) And (gM46V999.gTipoAcesso = DIGITACAO) Then
                        .AddItem Trim("**N�o informado")
                        .ItemData(.NewIndex) = 0
                    End If
                End With

            End If
            Call gM46V999.gpFecha2(bdP0044700)
        End If
        fraSeguradora.Visible = True
        fraApolice.Visible = True
        optSinistroCon(0).Enabled = True
        optSinistroCon(1).Enabled = True
        frmT46V101A.optSinistroYas(0).Enabled = True
        frmT46V101A.optSinistroYas(1).Enabled = True
        frmT46V101A.cmdBonus.Enabled = True
        frmT46V101A.cmdBonus.Picture = frmT46V101A.cmdExtra(1).Picture
        frmTipoEndosso.Visible = False
        frmT46V148A.txtTextoEndosso = ""
        cmdEndosso.Visible = False
        mTipEmissao = e_TipEmissao_RenCongenere
        cmdBusca.Visible = False
        txtRenApolice.MaxLength = 15
    Case 3      'Endosso
        fraSeguradora.Visible = False
        fraApolice.Visible = True
        If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
            frmT46V101A.cmdBonus.Enabled = False
            frmT46V101A.cmdBonus.Picture = frmT46V101A.cmdExtra(1).Picture
        End If
        frmTipoEndosso.Visible = True
        cmdEndosso.Visible = True
        mTipEmissao = e_TipEmissao_Endosso
        cmdBusca.Visible = True
    End Select
    Call PosicionaObj
    If gM46V999.gfDisponivel("frmT46V117A") And frmT46V101A.cboPlano.ListIndex > -1 Then
        ' CR1425 - Quando for plano, n�o tem bonus
        If mTipEmissao = e_TipEmissao_RenYasuda And (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) _
           And frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) >= 2001 Then
            frmT46V101A.cmdBonus.Enabled = False
            frmT46V101A.cmdBonus.Picture = frmT46V101A.cmdExtra(1).Picture
        End If
        ' CR1425 - Fim

        frmT46V117A.cmdBonus.Enabled = frmT46V101A.cmdBonus.Enabled
        frmT46V117A.cmdBonus.Picture = frmT46V101A.cmdBonus.Picture
    End If

    If mTipEmissaoAnt = 0 Then
        mTipEmissaoAnt = mTipEmissao
    Else
        'If gM46V999.gCodProduto >= 932 Then
        '    If mTipEmissaoAnt = e_TipEmissao_RenCongenere And mTipEmissao <> e_TipEmissao_RenCongenere Then
        '        'Se antes era congenere e agora n�o � , ent�o acrescentar 5 no bonus
        '         frmT46V101A.txtDescBonus.Text = Val(frmT46V101A.txtDescBonus.Text) + 5
        '    ElseIf mTipEmissaoAnt <> e_TipEmissao_RenCongenere And mTipEmissao = e_TipEmissao_RenCongenere Then
        '        'Se antes n�o era congenere e agora � , ent�o retirar 5 do bonus
        '        frmT46V101A.txtDescBonus.Text = Val(frmT46V101A.txtDescBonus.Text) - 5
        '    End If
        'End If
        mTipEmissaoAnt = mTipEmissao
    End If
End Sub
Private Sub optTipEmissao_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtConApolice_Change(index As Integer)
    Call gpHabSalvar
End Sub
Private Sub txtRenApolice_Change()
    Call gpHabSalvar
End Sub
Private Sub txtRenApolice_GotFocus()
    With txtRenApolice
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtRenApolice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtRenItem_Change()
    Call gpHabSalvar
End Sub
Private Sub txtRenItem_GotFocus()
    With txtRenItem
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtRenItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
