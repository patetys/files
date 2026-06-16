VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V105A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Endereço de cobrança"
   ClientHeight    =   1980
   ClientLeft      =   1365
   ClientTop       =   2490
   ClientWidth     =   8640
   ControlBox      =   0   'False
   Icon            =   "T46V105A.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1980
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtComplemento 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3570
      MaxLength       =   13
      TabIndex        =   9
      ToolTipText     =   "Complemento do endereço do segurado (número, andar, apartamento, conjunto, etc.)."
      Top             =   420
      Width           =   4965
   End
   Begin VB.TextBox txtNumero 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   7
      ToolTipText     =   "Complemento do endereço do segurado (número, andar, apartamento, conjunto, etc.)."
      Top             =   420
      Width           =   885
   End
   Begin VB.ComboBox cboUF 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "T46V105A.frx":000C
      Left            =   5265
      List            =   "T46V105A.frx":0061
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   17
      ToolTipText     =   "Unidade da federação do endereço do segurado."
      Top             =   1050
      Width           =   810
   End
   Begin VB.TextBox txtNumeroComplemento 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4995
      MaxLength       =   20
      TabIndex        =   19
      ToolTipText     =   "Complemento do endereço do segurado (número, andar, apartamento, conjunto, etc.)."
      Top             =   1500
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.TextBox txtBairro 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1485
      MaxLength       =   30
      TabIndex        =   11
      ToolTipText     =   "Bairro do endereço do segurado."
      Top             =   735
      Width           =   2925
   End
   Begin VB.TextBox txtNomLogradouro 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3990
      MaxLength       =   50
      TabIndex        =   5
      ToolTipText     =   "Nome do logradouro do endereço do segurado."
      Top             =   100
      Width           =   4560
   End
   Begin VB.TextBox txtTipLogradouro 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3390
      MaxLength       =   3
      TabIndex        =   4
      ToolTipText     =   "Tipo do logradouro do endereço do segurado (Rua, Av., etc.)."
      Top             =   100
      Width           =   600
   End
   Begin VB.TextBox txtCidade 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1485
      MaxLength       =   30
      TabIndex        =   15
      ToolTipText     =   "Cidade do endereço do segurado."
      Top             =   1050
      Width           =   2925
   End
   Begin VB.TextBox txtTelefones 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5265
      MaxLength       =   9
      TabIndex        =   13
      ToolTipText     =   "Telefone do endereço do segurado (Digitar somente números)."
      Top             =   735
      Width           =   3285
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   3840
      TabIndex        =   18
      ToolTipText     =   "Volta aos dados complementares."
      Top             =   1470
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdCEP 
      Height          =   250
      HelpContextID   =   42
      Left            =   100
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Busca endereço do segurado."
      Top             =   120
      Width           =   250
      _Version        =   65536
      _ExtentX        =   441
      _ExtentY        =   441
      _StockProps     =   78
      ForeColor       =   12582912
      Picture         =   "T46V105A.frx":00D1
   End
   Begin VB.TextBox txtCEP 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1485
      MaxLength       =   9
      TabIndex        =   2
      ToolTipText     =   "CEP do endereço do segurado (nnnnn-nnn)."
      Top             =   100
      Width           =   900
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Complemento: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   8
      ToolTipText     =   "Complemento do endereço do segurado."
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Número: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Complemento do endereço do segurado."
      Top             =   420
      Width           =   1395
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UF: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   6
      Left            =   4425
      TabIndex        =   16
      ToolTipText     =   "Unidade da federação do endereço do segurado."
      Top             =   1050
      Width           =   855
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bairro: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Bairro do endereço do segurado."
      Top             =   735
      Width           =   1395
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Logradouro: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   2385
      TabIndex        =   3
      ToolTipText     =   "Logradouro do endereço do segurado."
      Top             =   100
      Width           =   1005
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cidade: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Cidade do endereço do segurado."
      Top             =   1050
      Width           =   1395
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telefone: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   8
      Left            =   4410
      TabIndex        =   12
      ToolTipText     =   "Telefone(s) do endereço do segurado."
      Top             =   735
      Width           =   855
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CEP: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "CEP do endereço do segurado."
      Top             =   100
      Width           =   1395
   End
End
Attribute VB_Name = "frmT46V105A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAltEndereco As Boolean    'Indica que alterou endereço.
Private mAntCEP As String    'Anterior: CEP.
Private Sub mpVerEndCobranca()
'Procedure: verifica endereço do segurado.

    Dim lEndereco As String    'Endereço.
    Dim lTamanho As Integer    'Tamanho.

'1. Se não alterou endereço, não precisa verificar.
    If mAltEndereco = False Then
        Exit Sub
    End If

    '2. Verifica endereço.
    If gM46V999.gfPreenchido(txtTipLogradouro.Text) Then
        lEndereco = Trim$(txtTipLogradouro.Text)
    End If
    If gM46V999.gfPreenchido(txtNomLogradouro.Text) Then
        If gM46V999.gfPreenchido(lEndereco) Then
            lEndereco = lEndereco & " " & Trim$(txtNomLogradouro.Text)
        Else
            lEndereco = Trim$(txtNomLogradouro.Text)
        End If
    End If

    If gM46V999.gfPreenchido(lEndereco) Then
        lEndereco = lEndereco & " " & Trim$(txtNumero.Text) & " " & Trim$(txtComplemento.Text)
    Else
        lEndereco = Trim$(txtNumero.Text) & " " & Trim$(txtComplemento.Text)
    End If
        
    lTamanho = Len(lEndereco)
    If lTamanho > 50 Then
        Call gM46V999.gpGraLog(1, "AGE0036 - " & gM46V999.gPrefixo & "Endereço de cobrança completo: " & vbLf & vbLf & _
                                  UCase(lEndereco) & vbLf & vbLf & "O endereço completo está com " & lTamanho & _
                                " posições (tamanho máximo permitido no sistema = 50).")
    End If
    mAltEndereco = False
End Sub
Private Sub cboUF_Click()
'Observação: Lista do combo de unidades da federação carregadas em tempo de projeto.
'            Qualquer alteração deve ser feita na propriedade LIST do controle.

    Call gpHabSalvar
End Sub
Private Sub cboUF_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub cboUF_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then     'Teclou backspace.
        cboUF.ListIndex = -1
    End If
End Sub
Private Sub cmdCEP_Click()
    Dim lCEP As String    'CEP.
    Dim lTipLogradouro As String      'Tipo de logradouro.
    Dim lLogradouro As String   'Logradouro.
    Dim lBairro As String    'Bairro.
    Dim lCidade As String    'Cidade.
    Dim lUF As String    'UF.

    Call frmT46V104A.mpInstanciarForm(lCEP, lTipLogradouro, lLogradouro, lBairro, lCidade, lUF)
    txtCEP.Text = lCEP
    txtTipLogradouro.Text = lTipLogradouro
    txtNomLogradouro.Text = lLogradouro
    txtBairro.Text = lBairro
    txtCidade.Text = lCidade
    If gM46V999.gfPreenchido(lUF) Then
        cboUF.Text = lUF
    End If
End Sub
Private Sub cmdOk_Click()
    Me.Hide
End Sub
Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
Finalizar:
    If gM46V999.gDesabilitarForms Then
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

'1. Centraliza form (posiciona logo abaixo do endereço completo).
    Me.Left = frmT46V102A.Left
    Me.Top = frmT46V102A.Top + frmT46V102A.panEndereco.Top + frmT46V102A.panEndereco.Height

End Sub
Private Sub txtBairro_Change()
    Call gpHabSalvar
End Sub
Private Sub txtBairro_GotFocus()
    With txtBairro
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtBairro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtBairro_LostFocus()
    txtBairro.Text = gM46V999.gfRetirarCaracterEspecial(txtBairro.Text, True)
End Sub

Private Sub txtCEP_Change()
    Call gpHabSalvar
End Sub
Private Sub txtCEP_GotFocus()
    With txtCEP
        If gM46V999.gfPreenchido(.Text) Then
            mAntCEP = .Text
        Else
            mAntCEP = ""
        End If
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtCEP_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtCEP_LostFocus()
    Dim lbdP0043000 As ADODB.Connection    'Arquivo - P0043000.
    Dim lMensagem As String         'Mensagem.
    Dim lrsP0043000 As clsYasRecordSet  'Registro - P0043000.
    Dim lPos As Byte
    Dim lCEP As String
    Screen.MousePointer = vbHourglass
    Dim lSelect As String
    With txtCEP
        '1. Trata campo do CEP.
        If gM46V999.gfPreenchido(mAntCEP) Then
            If gM46V999.gfPreenchido(.Text) Then
                If mAntCEP = .Text Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            Else
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Else
            If Not gM46V999.gfPreenchido(.Text) Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        .Text = Trim$(.Text)
        Select Case Len(.Text)
        Case Is < 8
            lPos = InStr(1, .Text, "-")
            If lPos > 6 Then
                .Text = .Text
            ElseIf lPos < 6 And lPos <> 0 Then
                .Text = Format(Left(.Text, lPos - 1), "00000") & "-" & Format(Left(Mid(.Text, lPos + 1), 3), "000")
            Else
                .Text = Format(Format(.Text, "00000000"), "00000-000")
            End If
        Case 8
            lPos = InStr(1, .Text, "-")
            If lPos > 6 Then
                .Text = .Text
            ElseIf lPos <= 6 And lPos <> 0 Then
                .Text = Format(Left(.Text, lPos - 1), "00000") & Left(Mid(.Text, lPos + 1) & "000", 3)
            End If

            If Not IsNumeric(.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0039 - " & gM46V999.gPrefixo & _
                                          "CEP de pernoite do veículo inválido (não numérico).")
                .Text = ""
                If .Visible = True And .Enabled = True Then
                    .SetFocus
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            .Text = Mid$(.Text, 1, 5) & "-" & Mid$(.Text, 6, 3)
            lCEP = .Text
        Case 9
            If Not IsNumeric(Mid$(.Text, 1, 5)) Or Not IsNumeric(Mid$(.Text, 7, 3)) Then
                Call gM46V999.gpGraLog(1, "FGE0074 - " & gM46V999.gPrefixo & "CEP inválido (endereço do segurado - " & .Text & _
                                        " - não numérico - fora do formato nnnnn-nnn).")
                .Text = ""
                If .Visible = True And .Enabled = True Then
                    .SetFocus
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            If Mid$(.Text, 6, 1) <> "-" Then
                Call gM46V999.gpGraLog(1, "FGE0074 - " & gM46V999.gPrefixo & "CEP inválido (endereço do segurado - " & .Text & _
                                        " - não tem traço na posição 6 - fora do formato nnnnn-nnn).")
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End Select
        '3. Obtém CEP.
        'a) Abre arquivo.
        If gM46V999.gfAbrBasDados("P0043000", gM46V999.gTipP0043000, gM46V999.gEndP0043000, lbdP0043000, lMensagem) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'b) Obtém registro.
        lSelect = "SELECT * FROM P0043000 WHERE Num_CEP = '" & Left$(.Text, 5) & Right$(.Text, 3) & "'"
        If gM46V999.gfObtRegistro(lbdP0043000, lSelect, lrsP0043000, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdP0043000)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If lrsP0043000.EOF = True Then
            Call gM46V999.gpGraLog(1, "FNL0007 - " & gM46V999.gPrefixo & "CEP não localizado (" & .Text & ").")
            Call gM46V999.gpFecha2(lbdP0043000)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'c) Posiciona endereço.
        If gM46V999.gfPreenchido(lrsP0043000!TIP_LOGR) Then
            txtTipLogradouro.Text = lrsP0043000!TIP_LOGR
        Else
            txtTipLogradouro.Text = ""
        End If
        If gM46V999.gfPreenchido(lrsP0043000!Nom_Logr) Then
            txtNomLogradouro.Text = Left(lrsP0043000!Nom_Logr, IIf(InStr(1, lrsP0043000!Nom_Logr, "-") > 0, InStr(1, lrsP0043000!Nom_Logr, "-") - 2, 200))
        Else
            txtNomLogradouro.Text = ""
        End If
        txtNumero.Text = ""
        txtComplemento.Text = ""
        If gM46V999.gfPreenchido(lrsP0043000!Nom_Bairro) Then
            txtBairro.Text = lrsP0043000!Nom_Bairro
        Else
            txtBairro.Text = ""
        End If
        If gM46V999.gfPreenchido(lrsP0043000!Nom_Cid) Then
            txtCidade.Text = lrsP0043000!Nom_Cid
        Else
            txtCidade.Text = ""
        End If
        If gM46V999.gfPreenchido(lrsP0043000!Sig_UF) Then
            cboUF.Text = lrsP0043000!Sig_UF
        Else
            cboUF.ListIndex = -1
        End If
        'd) Fecha tabela e arquivo.
        Call gM46V999.gpFecha2(lbdP0043000)
    End With

    Screen.MousePointer = vbDefault
End Sub
Private Sub txtCidade_Change()
    Call gpHabSalvar
End Sub
Private Sub txtCidade_GotFocus()
    With txtCidade
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCidade_LostFocus()
    txtCidade.Text = gM46V999.gfRetirarCaracterEspecial(txtCidade.Text, True)
End Sub

Private Sub txtNumero_Change()
    mAltEndereco = True
    Call gpHabSalvar
End Sub
Private Sub txtNumero_GotFocus()
    With txtNumero
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtNumero_LostFocus()
    If gM46V999.gfPreenchido(txtNumero.Text) Then
        Call mpVerEndCobranca
    End If
    txtNumero.Text = gM46V999.gfRetirarCaracterEspecial(txtNumero.Text, True)
End Sub

Private Sub txtComplemento_Change()
    mAltEndereco = True
    Call gpHabSalvar
End Sub
Private Sub txtComplemento_GotFocus()
    With txtComplemento
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtComplemento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtComplemento_LostFocus()
    If gM46V999.gfPreenchido(txtComplemento.Text) Then
        Call mpVerEndCobranca
    End If
    txtComplemento.Text = gM46V999.gfRetirarCaracterEspecial(txtComplemento.Text, True)
End Sub
Private Sub txtNomLogradouro_Change()
    mAltEndereco = True
    Call gpHabSalvar
End Sub
Private Sub txtNomLogradouro_GotFocus()
    With txtNomLogradouro
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtNomLogradouro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtNomLogradouro_LostFocus()
    If gM46V999.gfPreenchido(txtNomLogradouro.Text) Then
        Call mpVerEndCobranca
    End If
    txtNomLogradouro.Text = gM46V999.gfRetirarCaracterEspecial(txtNomLogradouro.Text, True)
End Sub
Private Sub txtTelefones_Change()
    Call gpHabSalvar
End Sub
Private Sub txtTelefones_GotFocus()
    With txtTelefones
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtTelefones_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, 48 To 57  'Permite Back-Space e números (0 a 9)
    Case 13  'Se teclou enter move cursor para proximo campo
        KeyAscii = 0
        SendKeys "{TAB}"
    Case Else  'Não permite demais teclas
        KeyAscii = 0
    End Select
End Sub
Private Sub txtTelefones_LostFocus()
    If gM46V999.gfPreenchido(txtTelefones.Text) Then
        txtTelefones.Text = Trim$(txtTelefones.Text)
    End If
    txtTelefones.Text = gM46V999.gfRetirarCaracterEspecial(txtTelefones.Text, True)
End Sub
Private Sub txtTipLogradouro_Change()
    mAltEndereco = True
    Call gpHabSalvar
End Sub
Private Sub txtTipLogradouro_GotFocus()
    With txtTipLogradouro
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtTipLogradouro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtTipLogradouro_LostFocus()
    If gM46V999.gfPreenchido(txtTipLogradouro.Text) Then
        Call mpVerEndCobranca
    End If
    txtTipLogradouro.Text = gM46V999.gfRetirarCaracterEspecial(txtTipLogradouro.Text, True)
End Sub
