VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT47V016A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo de emissão"
   ClientHeight    =   3495
   ClientLeft      =   2100
   ClientTop       =   2430
   ClientWidth     =   10890
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnvio 
      Caption         =   "Envio de Documento para: "
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   8640
      TabIndex        =   22
      Top             =   1050
      Width           =   2160
      Begin VB.ComboBox cboTipEnvio 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1890
      End
   End
   Begin VB.OptionButton optTipEmissao 
      Caption         =   "&Endosso"
      ForeColor       =   &H00800000&
      Height          =   600
      Index           =   3
      Left            =   165
      TabIndex        =   2
      Top             =   1890
      Width           =   1035
   End
   Begin VB.OptionButton optTipEmissao 
      Caption         =   "Renovação de seguro &Yasuda"
      ForeColor       =   &H00800000&
      Height          =   600
      Index           =   1
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Renovação Yasuda."
      Top             =   900
      Width           =   1995
   End
   Begin VB.OptionButton optTipEmissao 
      Caption         =   "Seguro &novo"
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Seguro novo."
      Top             =   165
      Width           =   1395
   End
   Begin Threed.SSCommand cmdOk 
      Default         =   -1  'True
      Height          =   420
      Left            =   4320
      TabIndex        =   12
      ToolTipText     =   "Volta à proposta."
      Top             =   2910
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
      Left            =   6330
      TabIndex        =   13
      ToolTipText     =   "Volta à proposta."
      Top             =   2940
      Width           =   2010
      _Version        =   65536
      _ExtentX        =   3545
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Texto para Endosso"
      ForeColor       =   8388608
   End
   Begin VB.Frame frmCodReg 
      Caption         =   "Código de Registro"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   2250
      TabIndex        =   15
      Top             =   1050
      Width           =   6240
      Begin VB.ComboBox cboCodReg 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame fraSeguradora 
      Caption         =   "Seguradora: "
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   2265
      TabIndex        =   21
      Top             =   2835
      Visible         =   0   'False
      Width           =   8520
      Begin VB.ComboBox cboSeguradora 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Nome da seguradora congênere."
         Top             =   240
         Width           =   8265
      End
   End
   Begin VB.Frame fraApolice 
      Caption         =   "Dados da Apólice"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   2280
      TabIndex        =   16
      Top             =   165
      Width           =   8520
      Begin VB.TextBox txtRenItem 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3135
         MaxLength       =   5
         TabIndex        =   5
         ToolTipText     =   "Item da apólice de congênere."
         Top             =   240
         Width           =   705
      End
      Begin VB.TextBox txtRenApolice 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   4
         ToolTipText     =   "Número da cotação."
         Top             =   240
         Width           =   1680
      End
      Begin Threed.SSPanel panX 
         Height          =   300
         Index           =   1
         Left            =   4800
         TabIndex        =   17
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
         Begin VB.OptionButton optSinistroYas 
            Caption         =   "&Sim"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "SIM, tem sinistro."
            Top             =   40
            Width           =   600
         End
         Begin VB.OptionButton optSinistroYas 
            Caption         =   "&Não"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   690
            TabIndex        =   7
            ToolTipText     =   "NÃO tem sinistro."
            Top             =   40
            Value           =   -1  'True
            Width           =   600
         End
      End
      Begin Threed.SSCommand cmdBusca 
         Cancel          =   -1  'True
         Height          =   300
         Left            =   6360
         TabIndex        =   8
         ToolTipText     =   "Busca os dados da apólice."
         Top             =   240
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Busca dados de apólice"
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
         TabIndex        =   20
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Apólice: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sinistro? "
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   7
         Left            =   3960
         TabIndex        =   18
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame frmTipoEndosso 
      Caption         =   "Tipo de endosso:"
      ForeColor       =   &H00800000&
      Height          =   870
      Left            =   2265
      TabIndex        =   14
      Top             =   1890
      Width           =   8520
      Begin VB.ComboBox cboDizer 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T47V016A.frx":0000
         Left            =   165
         List            =   "T47V016A.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   315
         Width           =   8145
      End
   End
End
Attribute VB_Name = "frmT47V016A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTipEmissaoAnt  As e_TipEmissao
Private mTipEmissao     As e_TipEmissao
    
Dim lNumApolice     As String   'Número da apólice.
Dim lNumIteApolice  As String   'Número do item da apólice.


Private Sub cboConSeguradora_KeyPress(Index As Integer, KeyAscii As Integer)
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
    Dim lI  As Integer
    
    '1. Habilita opção de salvar.
    If gCarga = True Then Exit Sub
    Call gpLimPremios
    
    If cboCodReg.ListIndex > -1 Then
        gCod_Registro = Mid(cboCodReg.Text, 1, 3)
    Else
        gCod_Registro = 0
    End If

End Sub

Private Sub cboDizer_Change()
    cboDizer_Click
End Sub

Private Sub cboDizer_Click()
    '1. Habilita opção de salvar.
    If gCarga = True Then Exit Sub
    Call gpLimPremios
    If cboDizer.ListIndex > -1 Then
        gM46V111.gCod_Dizer = cboDizer.ItemData(cboDizer.ListIndex)
    Else
        gM46V111.gCod_Dizer = 0
    End If
End Sub

Private Sub cboSeguradora_Change()
    Call gpHabSalvar
End Sub

Private Sub cboTipEnvio_Change()
    cboTipEnvio_Click
End Sub

Private Sub cboTipEnvio_Click()
    '1. Habilita opção de salvar.
    If gCarga = True Then Exit Sub
End Sub

Private Sub cmdBusca_Click()
    Dim lbdP00APOL          As ADODB.Connection 'Arquivo - P00APOL.
    Dim lbdBasDados         As ADODB.Connection 'Arquivo - P00MPED
    Dim lBonus              As Long
    Dim lCodSeguradora      As Integer          'Código da seguradora.
    Dim lIndSinistro        As Byte             'Indicador de sinistro.
    Dim lNumItem            As Integer          'Número do item.
    Dim lMensagem           As String           'Mensagem.
    Dim rsAPOL              As clsYasRecordSet  'Registro: P00APOL.
    Dim lrsTabela           As clsYasRecordSet  'Registro: P00MPED.
    Dim lI                  As Long             'Auxiliar.
    Dim lobjstcClau         As stcA46V703B      'Objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU)
    Dim ColecaoClau         As Collection       'Coleção de cláusulas.
    Dim lobjstcCobert       As stcA46V704B      'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
    Dim ColecaoCobert       As Collection       'Coleção de coberturas.
    Dim lRegistro           As String           'Registro.
    Dim lIdeTabela          As String
    Dim lSemItem            As Boolean          'Renovacao sem item gravado na TAB_PED_LOC
    Dim lSelect     As String
    If Not txtRenApolice.Enabled Then Exit Sub
    
    Screen.MousePointer = vbHourglass

    lSemItem = False
    
    '1. Confirma com usuário.
    lMensagem = "AGE0018 - Confirma busca de dados de apólice? A busca de dados de apólice " & _
                "desconsiderará os dados existentes."
    If MsgBox(lMensagem, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Call gM46V999.gpGraLog(0, lMensagem & " Não.")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Call gM46V999.gpGraLog(0, lMensagem & " Sim.")

    '2. Consiste número de apólice anterior.
    If Not gM46V999.gfPreenchido(txtRenApolice.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                         "Número da apólice inválido (renovação Yasuda - não preenchido).")
        If txtRenApolice.Visible = True And txtRenApolice.Enabled = True Then
            txtRenApolice.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Not IsNumeric(txtRenApolice.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                         "Número da apólice inválido (renovação Yasuda - não numérico).")
        If txtRenApolice.Visible = True And txtRenApolice.Enabled = True Then
            txtRenApolice.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Val(txtRenApolice.Text) = 0 Then
        Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                         "Número da apólice inválido (renovação Yasuda - zerado).")
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

    '3. Consiste número do item da apólice.
    If Not optTipEmissao(3).Value = True Then
        If Not gM46V999.gfPreenchido(txtRenItem.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                             "Número do item da apólice inválido (renovação Yasuda - não preenchido).")
            If txtRenItem.Visible = True And txtRenItem.Enabled = True Then
                txtRenItem.SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If Not IsNumeric(txtRenItem.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                             "Número do item da apólice inválido (renovação Yasuda - não numérico).")
            If txtRenItem.Visible = True And txtRenItem.Enabled = True Then
                txtRenItem.SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If Val(txtRenItem.Text) = 0 Then
            Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                             "Número do item da apólice inválido (renovação Yasuda - zerado).")
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
    End If
    
    '4. Verifica se a apólice foi emitida pelo Syas
    gM46V999.gBaseLocal = False
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, lbdBasDados, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If gM46V999.gBaseLocal = True Then
        lIdeTabela = "TAB_PED_COBERT"
    Else
        lIdeTabela = "GED.DBO.TAB_PED_COBERT"
    End If
    
    lSelect = "SELECT * FROM " & lIdeTabela & " WHERE Num_Apol = " & Format(txtRenApolice.Text, "0000000000")
    If gM46V999.gfObtRegistro(lbdBasDados, lSelect, lrsTabela, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdBasDados)
        Call gM46V999.gpGraLog(1, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If lrsTabela.EOF = True And optTipEmissao(3).Value = True Then
        Call gM46V999.gpGraLog(1, "A apólice " & Format(txtRenApolice.Text, "0000000000") & " foi emitida no terminal. O endosso deve ser feito pelo terminal.")
        Screen.MousePointer = vbDefault
        Unload Me
        End
    End If
    
'sandra-aqui
    Set ColecaoCobert = New Collection
    Do While Not lrsTabela.EOF
        Set lobjstcCobert = New stcA46V704B
        Set lobjstcCobert.M46V999 = gM46V999
        
        
        With lobjstcCobert
            .Num_PI = gM46V999.gfTratarNulo(lrsTabela("Num_pi"), tpoNumerico)
            .Tab_Cobert = gM46V999.gfTratarNulo(lrsTabela("Tab_Cobert"), tpoNumerico)
            .Cod_Cobert = gM46V999.gfTratarNulo(lrsTabela("Cod_Cobert"), tpoNumerico)
        End With
        ColecaoCobert.Add lobjstcCobert, lobjstcCobert.Class_ID
        lrsTabela.MoveNext
    Loop
    
    '5. Obtém apólice/item.
    gM46V999.gBaseLocal = False
    If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, lbdP00APOL, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If gM46V999.gBaseLocal = True Then
        lIdeTabela = "TAB_APOL_LOC"
    Else
        lIdeTabela = "GED.DBO.TAB_PED_LOC"
    End If
    
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
        lSemItem = True
        If gM46V999.gBaseLocal = True Then
            lIdeTabela = "TAB_APOL"
        Else
            lIdeTabela = "GED.DBO.TAB_PED"
        End If
        
        lSelect = "SELECT * from " & lIdeTabela & _
            " WHERE " & lIdeTabela & ".Num_Apol = " & txtRenApolice.Text
        If gM46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdP00APOL)
            Call gM46V999.gpGraLog(1, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If rsAPOL.EOF = True Then
            lMensagem = "FNL0008 - Apólice " & txtRenApolice.Text & " não localizada."
        Else
            If lSemItem = True Then GoTo Sem_Item
            lMensagem = "FNL0009 - Item  " & txtRenItem.Text & " não localizado na apólice " & _
                        txtRenApolice.Text & "."
        End If
        Call gM46V999.gpFecha2(lbdP00APOL)
        Call gM46V999.gpGraLog(1, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    gCarga = True
    gM46V999.gBaseLocal = False
    frmT47V001A.txtApolice.Text = gM46V999.gfTratarNulo(rsAPOL("Num_Apol"), tpoString)
    frmT47V001A.txtRisTipLogradouro.Text = gM46V999.gfTratarNulo(rsAPOL("Cod_Tip_Logr_Risco"), tpoString)
    frmT47V001A.txtRisNomLogradouro.Text = gM46V999.gfTratarNulo(rsAPOL("Nom_Logr_Risco"), tpoString)
    frmT47V001A.txtRisComplemento.Text = gM46V999.gfTratarNulo(rsAPOL("Dsc_Compl_Risco"), tpoString)
    frmT47V001A.txtRisBairro.Text = gM46V999.gfTratarNulo(rsAPOL("Nom_Bairro_Risco"), tpoString)
    frmT47V001A.txtRisCidade.Text = gM46V999.gfTratarNulo(rsAPOL("Nom_Cid_Risco"), tpoString)
    frmT47V001A.cboRisUF.ListIndex = -1
    If gM46V999.gfPreenchido(gM46V999.gfTratarNulo(rsAPOL("Sig_UF_Risco"), tpoString)) Then
        For lI = 0 To (frmT47V001A.cboRisUF.ListCount - 1)
            If gM46V999.gfTratarNulo(rsAPOL("Sig_UF_Risco"), tpoString) = frmT47V001A.cboRisUF.List(lI) Then
                frmT47V001A.cboRisUF.ListIndex = lI
                Exit For
            End If
        Next lI
    End If
    frmT47V001A.txtRisCEP.Text = gM46V999.gfTratarNulo(rsAPOL("Num_CEP_Risco"), tpoNumerico)
    frmT47V001A.txtISBas.Text = Format(gM46V999.gfTratarNulo(rsAPOL("Per_IS_Vr"), TpoValor), "standard")
    frmT47V001A.txtCoefISVR.Text = Format(gM46V999.gfTratarNulo(rsAPOL("Val_Coef_IS_Vr"), TpoValor), "standard")
    frmT47V001A.txtValoRisco.Text = Format(gM46V999.gfTratarNulo(rsAPOL("Val_Vr"), TpoValor), "standard")
    frmT47V001A.txtClassificacao.Tag = gM46V999.gfTratarNulo(rsAPOL("Cod_Clasf"), tpoNumerico) & gM46V999.gfTratarNulo(rsAPOL("Cod_Compl_Clasf"), tpoNumerico) & Format(gM46V999.gfTratarNulo(rsAPOL("Tab_Clasf"), tpoNumerico), "0000")
    frmT47V001A.txtClassificacao.Text = gM46V999.gfTratarNulo(rsAPOL("Cod_Clasf"), tpoNumerico) & "-" & gM46V999.gfTratarNulo(rsAPOL("Dsc_Clasf"), tpoString)
    GoTo SAIDA

Sem_Item:
    gCarga = True
    gM46V999.gBaseLocal = False
    frmT47V001A.txtApolice.Text = gM46V999.gfTratarNulo(rsAPOL("Num_Apol"), tpoString)
    frmT47V001A.txtSegurado.Text = gM46V999.gfTratarNulo(rsAPOL("Nom_Segurado"), tpoString)
    If gM46V999.gfPreenchido(gM46V999.gfTratarNulo(rsAPOL("Cod_Estip"), tpoNumerico)) Then
        For lI = 0 To (frmT47V001A.cboPlano(1).ListCount - 1)
            If gM46V999.gfTratarNulo(rsAPOL("Cod_Estip"), tpoNumerico) = Left(Val(frmT47V001A.cboPlano(0).List(lI)), 4) Then
                frmT47V001A.cboPlano(1).ListIndex = lI
                frmT47V001A.cboPlano(0).ListIndex = lI
                Exit For
            End If
        Next lI
    End If

    gM46V999.gBaseLocal = False
    If gM46V999.gBaseLocal = True Then
        lIdeTabela = "TAB_APOL"
    Else
        lIdeTabela = "GED.DBO.TAB_PED_ENDER"
    End If
    
    lSelect = "SELECT * from " & lIdeTabela & _
        " WHERE " & lIdeTabela & ".Num_Apol = " & txtRenApolice.Text
    If gM46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdP00APOL)
        Call gM46V999.gpGraLog(1, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Not rsAPOL.EOF = True Then
        frmT47V001A.txtRisTipLogradouro.Text = gM46V999.gfTratarNulo(rsAPOL("Cod_Tip_Logr"), tpoString)
        frmT47V001A.txtRisNomLogradouro.Text = gM46V999.gfTratarNulo(rsAPOL("Nom_Logr"), tpoString)
        frmT47V001A.txtRisComplemento.Text = gM46V999.gfTratarNulo(rsAPOL("Dsc_Compl_Logr"), tpoString)
        frmT47V001A.txtRisBairro.Text = gM46V999.gfTratarNulo(rsAPOL("Nom_Bairro"), tpoString)
        frmT47V001A.txtRisCidade.Text = gM46V999.gfTratarNulo(rsAPOL("Nom_Cid"), tpoString)
        frmT47V001A.cboRisUF.ListIndex = -1
        If gM46V999.gfPreenchido(gM46V999.gfTratarNulo(rsAPOL("Sig_UF"), tpoString)) Then
            For lI = 0 To (frmT47V001A.cboRisUF.ListCount - 1)
                If gM46V999.gfTratarNulo(rsAPOL("Sig_UF"), tpoString) = frmT47V001A.cboRisUF.List(lI) Then
                    frmT47V001A.cboRisUF.ListIndex = lI
                    Exit For
                End If
            Next lI
        End If
        frmT47V001A.txtRisCEP.Text = gM46V999.gfTratarNulo(rsAPOL("Num_CEP"), tpoNumerico)
    End If
    
SAIDA:
    gCarga = False
    Call gpHabSalvar

    '7. Reposiciona informações do item e retorna ao form de item.
    Me.Hide
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEndosso_Click()
    frmT47V024A.Show vbModal, Me
End Sub

Private Sub cmdOK_Click()
    Dim lTipEmissao As Integer      'Tipo de emissão.
    Dim lTipEndosso As Integer      'Tipo de endosso.
    Dim lI          As Integer      'For...next.
    Dim lOpcao      As Integer      'Opção.
    Dim lMensagem   As String               'Mensagem.
    Dim lNosNumero  As String           'Nosso número.
    Dim lNovNosNumero   As String           'Novo: Nosso Numero.
    Dim bdP00APOL       As ADODB.Connection 'Base de dados: P00Apol.
    
    '0. Verifica tipo de endosso.
    If frmT47V001A.lVerTipEmissao = "E" Then
        If cboDizer.ListIndex > -1 Then
            Select Case cboDizer.ItemData(cboDizer.ListIndex)
                Case 1
                    lTipEndosso = 302
                Case 4
                    lTipEndosso = 303
                Case 10
                    lTipEndosso = 301
                Case Else
                    lTipEndosso = 300
            End Select
        Else
            Call gM46V999.gpGraLog(1, "FGE0071 - " & "Escolher o código do dizer correto! (Endosso Yasuda).")
            If cboDizer.Visible And cboDizer.Enabled Then
                cboDizer.SetFocus
            End If
            Exit Sub
        End If
    End If
    
    '1. Posiciona tipo de emissão.
    If optTipEmissao(0) = True Then         'Seguro novo.
        lTipEmissao = 100
        frmT47V001A.lVerTipEmissao = "S"
    ElseIf optTipEmissao(1) = True Then     'Renovação Yasuda.
        lTipEmissao = 200
        frmT47V001A.lVerTipEmissao = "R"
    ElseIf optTipEmissao(3) = True Then     'Endosso.
        lTipEmissao = lTipEndosso
        frmT47V001A.lVerTipEmissao = "E"
    Else
        lTipEmissao = 0
    End If
    

    '2. Consiste dados.
    Select Case lTipEmissao
        Case 100    'Seguro novo.
            'Nada a consistir.
            txtRenApolice.Text = ""
            txtRenItem.Text = ""
            frmT47V001A.txtApolice.Text = ""
            cboSeguradora.ListIndex = -1
            cboDizer.ListIndex = -1
        Case 101    'Renovação de congênere.
            If Not gM46V999.gfPreenchido(cboSeguradora.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0071 - " & gM46V999.gPrefixo & _
                                 "Escolher uma seguradora congênere! (Renovação de congênere).")
                If cboSeguradora.Visible And cboSeguradora.Enabled Then
                    cboSeguradora.SetFocus
                End If
                Exit Sub
            End If
        Case 200    'Renovação de seguro Yasuda.
            If Not gM46V999.gfPreenchido(txtRenApolice.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                                 "Número da apólice inválido (renovação Yasuda - não preenchido).")
                If txtRenApolice.Visible And txtRenApolice.Enabled Then
                    txtRenApolice.SetFocus
                End If
                Exit Sub
            End If
            If Not IsNumeric(txtRenApolice.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                                 "Número da apólice inválido (renovação Yasuda - não numérico).")
                If txtRenApolice.Visible And txtRenApolice.Enabled Then
                    txtRenApolice.SetFocus
                End If
                Exit Sub
            End If
            If Val(txtRenApolice.Text) = 0 Then
                Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                                 "Número da apólice inválido (renovação Yasuda - zerado).")
                If txtRenApolice.Visible And txtRenApolice.Enabled Then
                    txtRenApolice.SetFocus
                End If
                Exit Sub
            End If
            frmT47V001A.txtApolice.Text = Val(txtRenApolice.Text)
            If Not gM46V999.gfPreenchido(txtRenItem.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & "Número do item da apólice inválido " & _
                                 "(renovação Yasuda - não preenchido).")
                If txtRenItem.Visible And txtRenItem.Enabled Then
                    txtRenItem.SetFocus
                End If
                Exit Sub
            End If
            If Not IsNumeric(txtRenItem.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                                 "Número do item da apólice inválido (renovação Yasuda - não numérico).")
                If txtRenItem.Visible And txtRenItem.Enabled Then
                    txtRenItem.SetFocus
                End If
                Exit Sub
            End If
            If Val(txtRenItem.Text) = 0 Then
                Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                                 "Número do item da apólice inválido (renovação Yasuda - zerado).")
                If txtRenItem.Visible And txtRenItem.Enabled Then
                    txtRenItem.SetFocus
                End If
                Exit Sub
            End If
        Case 300, 301, 302, 303    'Endosso
            If Not gM46V999.gfPreenchido(cboDizer.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0071 - " & gM46V999.gPrefixo & _
                                 "Escolher o código do dizer correto! (Endosso Yasuda).")
                If cboDizer.Visible And cboDizer.Enabled Then
                    cboDizer.SetFocus
                End If
                Exit Sub
            End If
            If Not gM46V999.gfPreenchido(txtRenApolice.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0071 - " & gM46V999.gPrefixo & _
                                 "Número da apólice inválido (Endosso Yasuda - não preenchido).")
                If txtRenApolice.Visible And txtRenApolice.Enabled Then
                    txtRenApolice.SetFocus
                End If
                Exit Sub
            End If
            If Not IsNumeric(txtRenApolice.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0071 - " & gM46V999.gPrefixo & _
                                 "Número da apólice inválido (Endosso Yasuda - não numérico).")
                If txtRenApolice.Visible And txtRenApolice.Enabled Then
                    txtRenApolice.SetFocus
                End If
                Exit Sub
            End If
            If Val(txtRenApolice.Text) = 0 Then
                Call gM46V999.gpGraLog(1, "FGE0071 - " & gM46V999.gPrefixo & _
                                 "Número da apólice inválido (Endosso Yasuda - zerado).")
                If txtRenApolice.Visible And txtRenApolice.Enabled Then
                    txtRenApolice.SetFocus
                End If
                Exit Sub
            End If
            frmT47V001A.txtApolice.Text = Val(txtRenApolice.Text)
            If Not gM46V999.gfPreenchido(txtRenItem.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0072 - " & gM46V999.gPrefixo & "Endosso do item da apólice inválido " & _
                                 "(renovação Yasuda - não preenchido).")
                If txtRenItem.Visible = True And txtRenItem.Enabled = True Then
                    txtRenItem.SetFocus
                End If
                Exit Sub
            End If
       
    End Select
    
    '3. Verifica o código de registro
    If cboCodReg.ListIndex = -1 Then
        Call gM46V999.gpGraLog(1, "FGE0073 - " & "Escolher o código do registro correto!")
        If cboCodReg.Visible And cboCodReg.Enabled Then
            cboCodReg.SetFocus
        End If
        Exit Sub
    End If

'Consistência retirada do P01O003.CBL e incluída aqui
    If Mid(cboCodReg.Text, 1, 3) = 106 Or Mid(cboCodReg.Text, 1, 3) = 107 Then
        If Mid(cboTipEnvio.Text, 1, 8) <> "Corretor" Then
            Call gM46V999.gpGraLog(1, "FGE0074 - " & "Para Cod.Registro 106 ou 107 o tipo de envio deve ser CORRETOR")
            If cboTipEnvio.Visible = True And cboTipEnvio.Enabled = True Then
                cboTipEnvio.ListIndex = 2
            End If
        End If
    End If

 'Consistência retirada do P01O002.CBL e incluída aqui
    If frmT47V001A.lVerTipEmissao = "E" Then
        If gM46V999.gCod_Ramo = 112 Or gM46V999.gCod_Ramo = 113 Or gM46V999.gCod_Ramo = 114 Then
            'If cboDizer.ItemData(cboDizer.ListIndex) <> 1 And 2 And 3 And 4 And 5 And 6 And 7 And 10 And 11 And 12 And 18 Then
             If InStr(" " & cboDizer.ItemData(cboDizer.ListIndex) & ",", " 1, 2, 3, 4, 5, 6, 7, 10, 11, 12, 18,") > 0 Then
                Call gM46V999.gpGraLog(1, "FGE0075 - " & "Código do Dizer inválido")
                If cboDizer.Visible And cboDizer.Enabled Then
                    cboDizer.SetFocus
                End If
            End If
        Else
            'If cboDizer.ItemData(cboDizer.ListIndex) <> 12 And 18 Then
            If InStr(" " & cboDizer.ItemData(cboDizer.ListIndex) & ",", " 12, 18,") > 0 Then
                Call gM46V999.gpGraLog(1, "FGE0075 - " & "Código do Dizer do endosso deve ser preenchido com 12 ou 18")
                If cboDizer.Visible And cboDizer.Enabled Then
                    cboDizer.SetFocus
                End If
            End If
        End If
    End If
   
    If gM46V999.gCod_Ramo = 112 And (gM46V999.gCodProduto = 1020 Or 1030 Or 1040 Or 1050 Or 1060) Then
        If cboDizer.ItemData(cboDizer.ListIndex) = 5 Or 7 Then
            Call gM46V999.gpGraLog(1, "FGE0075 - " & "Dizer inválido para este produto")
            If cboDizer.Visible And cboDizer.Enabled Then
                cboDizer.SetFocus
            End If
        End If
    End If
    
    If gM46V999.gCod_Ramo = 113 Then
        'If cboDizer.ItemData(cboDizer.ListIndex) <> 3 And 6 And 10 And 11 And 12 And 18 Then
        If InStr(" " & cboDizer.ItemData(cboDizer.ListIndex) & ",", " 3, 6, 10, 11, 12, 18,") > 0 Then
            Call gM46V999.gpGraLog(1, "FGE0075 - " & "Dizer inválido para este produto")
            If cboDizer.Visible And cboDizer.Enabled Then
                cboDizer.SetFocus
            End If
        End If
    End If
    
    
    'Nosso_numero
    If (frmT47V001A.lVerTipEmissao = "E" Or frmT47V001A.lVerTipEmissao = "R") And frmT47V001A.mCarregarPI = False Then        'Posiciona nosso número.
        'Abre arquivo.
            gM46V999.gBaseLocal = False
            If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, bdP00APOL, lMensagem) = False Then
                Exit Sub
            End If
            
            If Not gM46V999.gfPreenchido(lNovNosNumero) Then
                If gM46V999.gfCalNosNumero(lNosNumero, lMensagem) = False Then
                    Call gM46V999.gpFecha2(bdP00APOL)
                    Exit Sub
                End If
                lNovNosNumero = lNosNumero
            Else
                lNosNumero = lNovNosNumero
            End If
            
           
            lNumApolice = Val(txtRenApolice.Text)      'Número da apólice.
            lNumIteApolice = 0
            gM46V999.gBaseLocal = False
            If gM46V001.gfGraRenYasuda(lNumApolice, lNumIteApolice, lNosNumero, lMensagem, lMensagem, "") = False Then
                If gM46V999.gfPreenchido(lMensagem) Then
                    Call gM46V999.gpGraLog(1, lMensagem)
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            Call frmT47V001A.mpCarregarPI
    End If
    
'    If objPedido Is Nothing Then
'        Set objPedido = New clsA46V702A
'    End If
'    If mInclusao Then
'        If Not objPedido.mfIncluir(objstcPedido) Then
'            pMensagem = objPedido.UltimoErro
'            Exit Function
'        End If
'    Else
'        If Not objPedido.mfAlterar(objstcPedido, mComUltData, mComUltUsuario) Then
'            pMensagem = objPedido.UltimoErro
'            Exit Function
'        End If
'    End If

    '3. Torna form invisível.
    Me.Hide
End Sub

Private Sub Form_Activate()
    Me.Caption = " Tipo de Emissão  -  PI: " & frmT47V001A.txtNumPI.Text
    PosicionaObj
Finalizar:
    If gM46V999.gDesabilitarForms Or gM46V999.gDesabFormsCotacao Then
        If Me.Tag <> "Desabilitado" Then
            DesabilitaControles Me
            cmdOk.Enabled = True
            If gM46V999.gDesabFormsCotacao = True And gM46V999.gTipoAcesso = DIGITACAO Then
                cboTipEnvio.Enabled = True
            End If
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
        End If
    End If
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        If gM46V999.gDesabFormsTransmi = True Then
            DesabilitaControles Me
            cmdOk.Enabled = True
            cmdEndosso.Enabled = True
        End If
    End If
End Sub

Private Sub optConSinistro_Click(Index As Integer)
    Call gpLimPremios
End Sub

Private Sub optConSinistro_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
Dim lbdTab_Dizer        As ADODB.Connection     'Arquivo: P00Tarifa.
Dim lMensagem           As String               'Mensagem.
Dim lrsTab_Dizer        As clsYasRecordSet      'Registro : Tab_Dizer.
Dim lLinha              As Long                 'Linhas.
Dim lSelect     As String
    Call gM46V999.gpCenForm(Me)  'Centraliza form.

    '1. Posiciona controles do help.
    'Inicio como Alteração
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
    Do While Not lrsTab_Dizer.EOF
        cboDizer.AddItem Format(lrsTab_Dizer("Cod_Dizer"), "00") & " - " & lrsTab_Dizer("Dsc_Dizer")
        cboDizer.ItemData(cboDizer.NewIndex) = lrsTab_Dizer("Cod_Dizer")
        lrsTab_Dizer.MoveNext
    Loop
    
    cboDizer.ListIndex = -1
    If gM46V111.gCod_Dizer <> 0 Then
        cboDizer.ListIndex = gM46V999.gfRetornarListIndex(cboDizer, gM46V111.gCod_Dizer)
    End If
        
    If gCod_Registro <> 0 Then
        For lLinha = 0 To cboCodReg.ListCount - 1
            If Mid(cboCodReg.List(lLinha), 1, 3) = gCod_Registro Then
                cboCodReg.ListIndex = lLinha
                Exit For
            End If
        Next
    End If

   cboTipEnvio.AddItem "Conforme cadastro"
   cboTipEnvio.AddItem "Segurado"
   cboTipEnvio.AddItem "Corretor"
   cboTipEnvio.ListIndex = 0

   Call gM46V999.gpFecha2(lbdTab_Dizer)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub optSinistroYas_Click(Index As Integer)
    '1. Habilita opção de salvar.
    Call gpLimPremios
End Sub

Private Sub optSinistroYas_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub optTipEmissao_Click(KeyAscii As Integer)
    Dim lMensagem   As String           'Mensagem.
    
    '2. Posiciona quadros de informações e tipo de emissão.
    If optTipEmissao(0).Value = True Then    'Novo
            fraSeguradora.Visible = False
            fraApolice.Visible = False
            frmTipoEndosso.Visible = False
            cmdEndosso.Visible = False
            mTipEmissao = e_TipEmissao_SeguroNovo
            txtRenApolice.Text = ""
            txtRenItem.Text = ""
            frmT47V001A.lVerTipEmissao = "S"
    End If
    If optTipEmissao(1).Value = True Then    'Renovação Yasuda
            fraSeguradora.Visible = False
            fraApolice.Visible = True
            frmTipoEndosso.Visible = False
            cmdEndosso.Visible = False
            mTipEmissao = e_TipEmissao_RenYasuda
            cmdBusca.Visible = True
            txtRenApolice.MaxLength = 10
            frmT47V001A.lVerTipEmissao = "R"
    End If
    
    '1200287 - Syas Multi - RD - Codigo de Registro
    cboCodReg.Clear
    cboCodReg.AddItem "001-Apólice com cobrança de prêmio"
    cboCodReg.AddItem "002-Apólice sem cobrança de prêmio"
    cboCodReg.AddItem "003-Apólice sem cobrança de IOF"
    cboCodReg.AddItem "004-Apólice sem cobrança de custo"
    cboCodReg.AddItem "005-Apólice de Moeda Estrangeira"
    cboCodReg.AddItem "006-Apólice coletivo Acidentes Pessoais"
    cboCodReg.AddItem "007-Apólice coletivo pagamento trimestral"
    cboCodReg.AddItem "008-Apólice coletivo pagamento semestral"
    cboCodReg.AddItem "020-Apólice a averbar (sem cobrança de prêmio)"
    
    If optTipEmissao(3).Value = True Then    'Endosso
            fraSeguradora.Visible = False
            fraApolice.Visible = True
            frmTipoEndosso.Visible = True
            cmdEndosso.Visible = True
            mTipEmissao = e_TipEmissao_Endosso
            cmdBusca.Visible = True
            frmT47V001A.lVerTipEmissao = "E"
            
            '1200287 - Syas Multi - RD - Codigo de Registro
            cboCodReg.Clear
            cboCodReg.AddItem "101-Endosso com cobrança de IOF e Custo"
            cboCodReg.AddItem "102-Endosso sem cobrança de IOF e Custo"
            cboCodReg.AddItem "103-Endosso sem cobrança de IOF"
            cboCodReg.AddItem "104-Endosso sem cobrança de custo"
            cboCodReg.AddItem "105-Endosso de Moeda Estrangeira"
            cboCodReg.AddItem "106-Endosso restituição de prêmio - sem cancelamento da apólice"
            cboCodReg.AddItem "107-Endosso restituição de prêmio - com cancelamento da apólice"
            cboCodReg.AddItem "120-Endosso sem movimento de prêmio"
            cboCodReg.AddItem "121-Endosso sem movimento de prêmio - com prêmio negativo"
            cboCodReg.AddItem "122-Endosso sem movimento de prêmio - com prêmio positivo"
            cboCodReg.AddItem "129-Endosso de cancelamento de apólice sem estorno (devolução ?) de prêmio"
            cboCodReg.AddItem "141-Endosso de cancelamento de apólice"
            cboCodReg.AddItem "142-Endosso de cancelamento de endosso"
            cboCodReg.AddItem "143-Endosso de cancelamento de fatura mensal"
            cboCodReg.AddItem "201-Fatura mensal com cobrança de prêmio"
            cboCodReg.AddItem "202-Fatura mensal sem cobrança de prêmio"
            cboCodReg.AddItem "203-Fatura mensal sem cobrança de IOF"
            cboCodReg.AddItem "204-Fatura mensal sem cobrança de custo"
            cboCodReg.AddItem "205-Fatura mensal de Moeda Estrangeira"
            cboCodReg.AddItem "400-Bilhete de seguro"
            cboCodReg.AddItem "600-Sinistros avisados - Direto"
            cboCodReg.AddItem "601-Sinistros avisados - Cosseguros Aceitos"
            cboCodReg.AddItem "610-Sinistros pagos / recuperados - Direto"
            cboCodReg.AddItem "611-Sinistros pagos / recuperados - Cosseguro Aceito"
            cboCodReg.AddItem "630-Sinistros avisados - Direto, com número de sinistro"
            cboCodReg.AddItem "631-Sinistros avisados - Cosseguros, com número de sinistro"
            cboCodReg.AddItem "650-Estorno de sinistros pagos - Direto"
            cboCodReg.AddItem "651-Estorno de sinistros pagos - Cosseguros Aceitos"
            cboCodReg.AddItem "800-Comissões"
    End If
    Call PosicionaObj
   
    If mTipEmissaoAnt = 0 Then
        mTipEmissaoAnt = mTipEmissao
    Else
        mTipEmissaoAnt = mTipEmissao
    End If
    
    frmT47V001A.mTipoEmissao = mTipEmissao
End Sub
Private Sub optTipEmissao_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtConApolice_Change(Index As Integer)
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
