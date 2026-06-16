VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT47V010A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Débito em conta"
   ClientHeight    =   1740
   ClientLeft      =   1860
   ClientTop       =   1545
   ClientWidth     =   4830
   Icon            =   "T47V010A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDebConta 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1300
      MaxLength       =   9
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Número da conta corrente para débito."
      Top             =   915
      Width           =   1200
   End
   Begin VB.TextBox txtDebAgencia 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1300
      MaxLength       =   4
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Agência da conta de débito."
      Top             =   525
      Width           =   800
   End
   Begin VB.ComboBox cboDebBanco 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1300
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Banco da conta de débito."
      Top             =   120
      Width           =   3500
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Volta à proposta."
      Top             =   1335
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
      Index           =   2
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Número da conta corrente para débito."
      Top             =   915
      Width           =   1305
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agência: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Agência da conta de débito."
      Top             =   525
      Width           =   1305
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Banco: "
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Banco da conta de débito."
      Top             =   120
      Width           =   1305
   End
End
Attribute VB_Name = "frmT47V010A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDebBanco_Click()
    Dim lAuxiliar   As String   'Auxiliar.
    Dim lCodBanco   As String   'Código do banco.
    Dim lMensagem   As String   'Mensagem.
    Dim lTamAgencia As Byte     'Tamanho do campo de agência.
    Dim lTamConta   As Byte     'Tamanho do campo de conta.

    '1. Habilita salvar.
    If gM46V999.PerfilAcesso <> e_Perfil_Triagem Then
        Call gpHabSalvar
    End If

    '2. Posiciona código do banco.
    With cboDebBanco
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
            lTamAgencia = 4
            lTamConta = 11
        Case "479"      'Boston.
            lTamAgencia = 3
            lTamConta = 9
        Case "356"      'Real.
            lTamAgencia = 4
            lTamConta = 11
        Case "399"      'HSBC
            lTamAgencia = 4
            lTamConta = 12
        Case Else       'Outros bancos.
            lTamAgencia = 4
            lTamConta = 9
    End Select
    If gM46V999.gfPreenchido(txtDebAgencia.Text) Then
        If Len(Trim(txtDebAgencia.Text)) > lTamAgencia Then
            txtDebAgencia.Text = Left$(Trim(txtDebAgencia.Text), lTamAgencia)
        End If
        txtDebAgencia.MaxLength = lTamAgencia
        If gfVerAgencia(1, lCodBanco, txtDebAgencia.Text, lAuxiliar, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
        Else
            If txtDebAgencia.Text <> lAuxiliar Then
                txtDebAgencia.Text = lAuxiliar
            End If
        End If
    Else
        txtDebAgencia.MaxLength = lTamAgencia
    End If
    If gM46V999.gfPreenchido(txtDebConta.Text) Then
        If Len(Trim(txtDebConta.Text)) > lTamConta Then
            txtDebConta.Text = Left$(Trim(txtDebConta.Text), lTamConta)
        End If
        txtDebConta.MaxLength = lTamConta
        If gfVerConta(1, lCodBanco, txtDebConta.Text, lAuxiliar, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
        Else
            If txtDebConta.Text <> lAuxiliar Then
                txtDebConta.Text = lAuxiliar
            End If
        End If
    Else
        txtDebConta.MaxLength = lTamConta
    End If
End Sub

Private Sub cboDebBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub cboDebBanco_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then     'Teclou backspace.
        cboDebBanco.ListIndex = -1
    End If
End Sub
Private Sub cmdOK_Click()
    Me.Hide
End Sub
Private Sub Form_Activate()

     If cboDebBanco.Visible = True And cboDebBanco.Enabled = True Then
        cboDebBanco.SetFocus
    End If
Finalizar:
    If gM46V999.gDesabilitarForms Then
        If Me.Tag <> "Desabilitado" Then
            DesabilitaControles Me
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
        End If
    End If
    
End Sub
Public Sub mpInstanciarForm(pTipoPagamento As Integer)
    
    Me.Show vbModal, frmT47V004A
 
End Sub

Private Sub Form_Load()
    Dim lMensagem   As String   'Mensagem.

    '1. Centraliza form.
    Call gM46V999.gpCenForm(Me)

    '2. Carrega combos de bancos.
    If gfCarCboBancos(3, cboDebBanco, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
End Sub
Private Sub txtDebAgencia_Change()
    If gM46V999.PerfilAcesso <> e_Perfil_Triagem Then
        Call gpHabSalvar
    End If
End Sub
Private Sub txtDebAgencia_GotFocus()
    With txtDebAgencia
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtDebAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtDebAgencia_LostFocus()
    Dim lAuxiliar   As String   'Auxiliar.
    Dim lCodBanco   As String   'Código do banco.
    Dim lMensagem   As String   'Mensagem.

    If gM46V999.gfPreenchido(txtDebAgencia.Text) Then
        With cboDebBanco
            If .ListIndex = -1 Then
                lCodBanco = "000"
            Else
                lCodBanco = Left$(.Text, 3)
            End If
        End With
        If gfVerAgencia(1, lCodBanco, txtDebAgencia.Text, lAuxiliar, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If
        If txtDebAgencia.Text <> lAuxiliar Then
            txtDebAgencia.Text = lAuxiliar
        End If
    End If
End Sub
Private Sub txtDebConta_Change()
    If gM46V999.PerfilAcesso <> e_Perfil_Triagem Then
        Call gpHabSalvar
    End If
End Sub
Private Sub txtDebConta_GotFocus()
    With txtDebConta
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtDebConta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtDebConta_LostFocus()
    Dim lAuxiliar   As String   'Auxiliar.
    Dim lCodBanco   As String   'Código do banco.
    Dim lMensagem   As String   'Mensagem.

    If gM46V999.gfPreenchido(txtDebConta.Text) Then
        With cboDebBanco
            If .ListIndex = -1 Then
                lCodBanco = "000"
            Else
                lCodBanco = Left$(.Text, 3)
            End If
        End With
        If gfVerConta(1, lCodBanco, txtDebConta.Text, lAuxiliar, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If
        If txtDebConta.Text <> lAuxiliar Then
            txtDebConta.Text = lAuxiliar
        End If
    End If
End Sub

