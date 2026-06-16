VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V183A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inclusão de Inquilino"
   ClientHeight    =   2130
   ClientLeft      =   2085
   ClientTop       =   1800
   ClientWidth     =   5925
   ControlBox      =   0   'False
   Icon            =   "T46V183A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbTpPessoa 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "T46V183A.frx":000C
      Left            =   1140
      List            =   "T46V183A.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   1395
   End
   Begin VB.Frame fraDocumento 
      Caption         =   "Documento"
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   60
      TabIndex        =   6
      ToolTipText     =   "Documento (Apenas para pessoa Física)"
      Top             =   870
      Width           =   5775
      Begin VB.ComboBox cboIdeTipDocumento 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V183A.frx":002C
         Left            =   630
         List            =   "T46V183A.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Tipo do documento de identificação do segurado."
         Top             =   270
         Width           =   2505
      End
      Begin VB.TextBox txtNumeroDocumento 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3150
         TabIndex        =   9
         Top             =   270
         Width           =   2535
      End
      Begin VB.Label lblTipoDocumento 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo:  "
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   90
         TabIndex        =   7
         ToolTipText     =   "Tipo do documento de identificação do segurado."
         Top             =   270
         Width           =   540
      End
   End
   Begin VB.TextBox txtCpfCnpj 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4110
      TabIndex        =   5
      Top             =   480
      Width           =   1635
   End
   Begin VB.TextBox txtInquilino 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1140
      MaxLength       =   60
      TabIndex        =   1
      Top             =   90
      Width           =   4620
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   2460
      TabIndex        =   10
      ToolTipText     =   "Volta à proposta."
      Top             =   1740
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin VB.Label lblTipoPessoa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Pessoa: "
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   60
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   480
      Width           =   1065
   End
   Begin VB.Label lblCpfCnpj 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CNPJ: "
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3330
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   480
      Width           =   765
   End
   Begin VB.Label lblInquilino 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inquilino:"
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   60
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   90
      Width           =   1065
   End
End
Attribute VB_Name = "frmT46V183A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbdTab_Ped_Loc_Pessoas As ADODB.Connection
Private mrsTab_Ped_Loc_Pessoas As clsYasRecordSet
Private lVerificaUnload As Boolean

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
    lVerificaUnload = False
    cmbTpPessoa.ListIndex = 0
    Call CarregarCombo(cmbTpPessoa.ListIndex)
End Sub

Private Sub cmbTpPessoa_Click()
    Call CarregarCombo(cmbTpPessoa.ListIndex)
    If cmbTpPessoa.ListIndex = 0 Then
        lblCpfCnpj = "CNPJ: "
        fraDocumento.Enabled = False
        txtNumeroDocumento.Text = Empty
    Else
        lblCpfCnpj = "CPF: "
        fraDocumento.Enabled = True
    End If
    txtCpfCnpj.Text = Empty
End Sub

Private Sub cmbTpPessoa_KeyPress(KeyAscii As Integer)
    With cmbTpPessoa
        If KeyAscii = vbKeyReturn Then
            SendKeys "{TAB}"
        End If
    End With
End Sub

Private Sub cmdOk_Click()
    If Trim(txtInquilino.Text) = Empty Or cmbTpPessoa.Text = Empty Or Trim(txtCpfCnpj.Text) = Empty Then
        lVerificaUnload = True
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lMensagem As String

    If lVerificaUnload = True Then
        lVerificaUnload = False
    Else
        Exit Sub
    End If

    lMensagem = "AGE0016 - " & gM46V999.gPrefixo & "Excluída opção de inquilino. " & _
                "Confirma esta opção?"
    If MsgBox(lMensagem, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        Call gM46V999.gpGraLog(0, lMensagem & " Sim.")
        gCarga = True
        frmT46V101A.optInquilino(0).Value = False
        frmT46V101A.optInquilino(1).Value = True
        gCarga = False
        Call gM46V999.gpGraLog(0, lMensagem & " Sim.")
        On Error Resume Next
    Else
        Call gM46V999.gpGraLog(0, lMensagem & " Não.")
        txtInquilino.SetFocus
        Cancel = 1
    End If

End Sub

Private Sub txtCpfCnpj_GotFocus()
    With txtCpfCnpj
        .Text = Format(Replace(Replace(Replace(.Text, ".", ""), "-", ""), "/", ""), "#0")
    End With
End Sub

Private Sub txtCpfCnpj_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    Case vbKeyDelete, vbKeyBack
        Exit Sub

    Case vbKeyReturn
        SendKeys "{TAB}"

    Case 22    'Ctrl+V
        Dim lcopy As String

        lcopy = Replace(Replace(Replace(Clipboard.GetText, ".", Empty), "-", Empty), "/", Empty)

        If IsNumeric(lcopy) = True Then

            If UCase(cmbTpPessoa.Text) = "FÍSICA" Then    'CPF
                If Len(lcopy) > 11 Then
                    KeyAscii = 0
                    Beep
                End If
            Else    'CNPJ
                If Len(lcopy) > 14 Then
                    KeyAscii = 0
                    Beep
                End If
            End If
        Else
            KeyAscii = 0
            Beep
        End If

    Case 48 To 57    'Zero a nove
        If UCase(cmbTpPessoa.Text) = "FÍSICA" Then    'CPF
            If Len(txtCpfCnpj.Text) >= 11 Then
                KeyAscii = 0
                Beep
            End If
        Else
            If Len(txtCpfCnpj.Text) >= 14 Then    'CNPJ
                KeyAscii = 0
                Beep
            End If
        End If

    Case Else
        KeyAscii = 0
        Beep

    End Select

End Sub

Private Sub txtCpfCnpj_LostFocus()
    Call FormataCpfCnpj
End Sub

Private Sub FormataCpfCnpj()
    Dim lCpfCnpj As String

    With txtCpfCnpj
        lCpfCnpj = Format(Replace(Replace(Replace(.Text, ".", ""), "-", ""), "/", ""), "#0")

        If UCase(cmbTpPessoa.Text) = "FÍSICA" Then
            Call gM46V999.gpForCPF(1, lCpfCnpj, lCpfCnpj)
        Else
            Call gM46V999.gpForCNPJ(1, lCpfCnpj, lCpfCnpj)
        End If

        .Text = lCpfCnpj
    End With
End Sub

Private Sub txtInquilino_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub CarregarCombo(ByVal pTipoPessoa As Integer)
    
    With cboIdeTipDocumento
        .Clear
        If pTipoPessoa = 0 Then 'Juridica
            .AddItem ""
            .ItemData(.NewIndex) = 0
            .Enabled = False
        Else
            .AddItem ""
            .ItemData(.NewIndex) = 0
            .Enabled = False
            .AddItem "RG / Carteira de identidade"
            .ItemData(.NewIndex) = 1
            .AddItem "CNH / Carteira Nacional de Habilitação"
            .ItemData(.NewIndex) = 2
            .AddItem "Carteira de trabalho"
            .ItemData(.NewIndex) = 3
            .AddItem "Entidades de Classe"
            .ItemData(.NewIndex) = 7
            .AddItem "RNE / Identidade estrangeira"
            .ItemData(.NewIndex) = 8
            .AddItem "Passaporte"
            .ItemData(.NewIndex) = 9
            .AddItem "PIS"
            .ItemData(.NewIndex) = 10
            .Enabled = True
        End If
        .ListIndex = 0
    End With
End Sub

Public Sub CarregarDados(ByRef pPessoa As stcA46V735B)
    Dim lNumeroDoc As String
    
    txtInquilino.Text = pPessoa.Nom_Pessoa
    cmbTpPessoa.ListIndex = pPessoa.Tip_Pessoa - 1    'Enum e_TipPessoa começa do um
    CarregarCombo pPessoa.Tip_Pessoa
    
    If UCase(cmbTpPessoa.Text) = "FÍSICA" Then
        Call gM46V999.gpForCPF(1, pPessoa.Num_CPF_CNPJ, lNumeroDoc)
    Else
        Call gM46V999.gpForCNPJ(1, pPessoa.Num_CPF_CNPJ, lNumeroDoc)
    End If
    txtCpfCnpj.Text = lNumeroDoc
    cboIdeTipDocumento.ListIndex = pPessoa.Tip_Doc
    txtNumeroDocumento.Text = pPessoa.Num_Doc
End Sub

Public Sub SalvarDados(ByRef pPessoa As stcA46V735B)
   
    With pPessoa
        .Tip_Loc_Pessoa = e_TipLocPessoa_Inquilino
        .Nom_Pessoa = gM46V999.gfRetirarCaracterEspecial(txtInquilino.Text, True)
        .Tip_Pessoa = cmbTpPessoa.ListIndex + 1    'Enum e_TipPessoa começa do um
        
        If gM46V999.gfPreenchido(txtCpfCnpj.Text) = False Then
            .Num_CPF_CNPJ = 0
        Else
            .Num_CPF_CNPJ = Replace(Replace(Replace(txtCpfCnpj.Text, ".", ""), "-", ""), "/", "")
        End If
        .Tip_Doc = cboIdeTipDocumento.ListIndex
        .Num_Doc = txtNumeroDocumento.Text
    End With
   
End Sub
