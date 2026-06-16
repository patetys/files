VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V119A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indicador do seguro"
   ClientHeight    =   1725
   ClientLeft      =   3315
   ClientTop       =   1740
   ClientWidth     =   5010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "T46V119A.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtIndConta 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1400
      MaxLength       =   9
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Número da conta corrente do indicador do seguro."
      Top             =   900
      Width           =   1200
   End
   Begin VB.TextBox txtIndAgencia 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1400
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Agência da conta do indicador do seguro."
      Top             =   500
      Width           =   800
   End
   Begin VB.ComboBox cboIndBanco 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1400
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Banco da conta do indicador do seguro."
      Top             =   100
      Width           =   3500
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   300
      Left            =   1900
      TabIndex        =   6
      ToolTipText     =   "Volta à proposta."
      Top             =   1300
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conta corrente: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   37
      Left            =   100
      TabIndex        =   4
      ToolTipText     =   "Número da conta corrente do indicador do seguro."
      Top             =   900
      Width           =   1300
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agência: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   36
      Left            =   100
      TabIndex        =   2
      ToolTipText     =   "Agência da conta do indicador do seguro."
      Top             =   500
      Width           =   1300
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Banco: "
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   35
      Left            =   100
      TabIndex        =   0
      ToolTipText     =   "Banco da conta do indicador do seguro."
      Top             =   100
      Width           =   1300
   End
End
Attribute VB_Name = "frmT46V119A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboIndBanco_Click()
    Dim lAuxiliar As String    'Auxiliar.
    Dim lCodBanco As String    'Código do banco.
    Dim lMensagem As String    'Mensagem.
    Dim lTamAgencia As Byte     'Tamanho do campo de agência.
    Dim lTamConta As Byte   'Tamanho do campo de conta.

'1. Habilita salvar.
    Call gpHabSalvar

    '2. Posiciona código do banco.
    With cboIndBanco
        If .ListIndex = -1 Then
            lCodBanco = "000"
        Else
            lCodBanco = Left$(.Text, 3)
        End If
    End With

    '3. Posiciona e verifica campos de agência e conta.
    Select Case lCodBanco
    Case "237"      'Bradesco.
        lTamAgencia = 6
        lTamConta = 9
    Case "409"      'Unibanco.
        lTamAgencia = 4
        lTamConta = 8
    Case "341"      'Itaú.
        lTamAgencia = 4
        lTamConta = 7
    Case "347"      'Sudameris.
        lTamAgencia = 3
        lTamConta = 11
    Case "479"      'Boston.
        lTamAgencia = 3
        lTamConta = 9
    Case "356"      'Real.
        lTamAgencia = 4
        lTamConta = 11
    Case Else       'Outros bancos.
        lTamAgencia = 4
        lTamConta = 9
    End Select
    If gM46V999.gfPreenchido(txtIndAgencia.Text) Then
        If Len(Trim(txtIndAgencia.Text)) > lTamAgencia Then
            txtIndAgencia.Text = Left$(Trim(txtIndAgencia.Text), lTamAgencia)
        End If
        txtIndAgencia.MaxLength = lTamAgencia
        If gfVerAgencia(2, lCodBanco, txtIndAgencia.Text, lAuxiliar, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
        Else
            If txtIndAgencia.Text <> lAuxiliar Then
                txtIndAgencia.Text = lAuxiliar
            End If
        End If
    End If
    If gM46V999.gfPreenchido(txtIndConta.Text) Then
        If Len(Trim(txtIndConta.Text)) > lTamConta Then
            txtIndConta.Text = Left$(Trim(txtIndConta.Text), lTamConta)
        End If
        txtIndConta.MaxLength = lTamConta
        If gfVerConta(2, lCodBanco, txtIndConta.Text, lAuxiliar, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
        Else
            If txtIndConta.Text <> lAuxiliar Then
                txtIndConta.Text = lAuxiliar
            End If
        End If
    End If
End Sub
Private Sub cboIndBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub cboIndBanco_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then     'Teclou backspace.
        cboIndBanco.ListIndex = -1
    End If
End Sub
Private Sub cmdOk_Click()
    Me.Hide
End Sub
Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    If cboIndBanco.Visible = True And cboIndBanco.Enabled = True Then
        cboIndBanco.SetFocus
    End If
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
    Dim lMensagem As String    'Mensagem.

'1. Centraliza form.
    Call gM46V999.gpCenForm(Me)

    '2. Carrega combo de bancos.
    If gfCarCboBancos(4, cboIndBanco, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
End Sub
Private Sub txtIndAgencia_Change()
    Call gpHabSalvar
End Sub
Private Sub txtIndAgencia_GotFocus()
    With txtIndAgencia
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtIndAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtIndAgencia_LostFocus()
    Dim lAuxiliar As String    'Auxiliar.
    Dim lCodBanco As String    'Código do banco.
    Dim lMensagem As String    'Mensagem.

    If gM46V999.gfPreenchido(txtIndAgencia.Text) Then
        With cboIndBanco
            If .ListIndex = -1 Then
                lCodBanco = "000"
            Else
                lCodBanco = Left$(.Text, 3)
            End If
        End With
        If gfVerAgencia(2, lCodBanco, txtIndAgencia.Text, lAuxiliar, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If
        If txtIndAgencia.Text <> lAuxiliar Then
            txtIndAgencia.Text = lAuxiliar
        End If
    End If
End Sub
Private Sub txtIndConta_Change()
    Call gpHabSalvar
End Sub
Private Sub txtIndConta_GotFocus()
    With txtIndConta
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtIndConta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtIndConta_LostFocus()
    Dim lAuxiliar As String    'Auxiliar.
    Dim lCodBanco As String    'Código do banco.
    Dim lMensagem As String    'Mensagem.

    If gM46V999.gfPreenchido(txtIndConta.Text) Then
        With cboIndBanco
            If .ListIndex = -1 Then
                lCodBanco = "000"
            Else
                lCodBanco = Left$(.Text, 3)
            End If
        End With
        If gfVerConta(2, Left$(lCodBanco, 3), txtIndConta.Text, lAuxiliar, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If
        If txtIndConta.Text <> lAuxiliar Then
            txtIndConta.Text = lAuxiliar
        End If
    End If
End Sub
