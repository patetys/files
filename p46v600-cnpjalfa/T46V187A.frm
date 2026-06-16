VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V187A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Meio para devolução de créditos futuros - Restituição"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSCommand cmdCopiar 
      Height          =   360
      Left            =   1020
      TabIndex        =   5
      ToolTipText     =   "Volta à proposta."
      Top             =   1440
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "Usar Conta Débito"
      ForeColor       =   8388608
   End
   Begin VB.Frame Frame 
      Height          =   1875
      Left            =   30
      TabIndex        =   1
      Top             =   900
      Width           =   5055
      Begin VB.OptionButton optCredito 
         Caption         =   "Não Informado"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   210
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton optCredito 
         Caption         =   "Não. Declaro que estou ciente que eventuais devoluções de prêmio serão realizadas através de ordem de pagamento pelo Banco Itaú."
      ForeColor       =   &H00800000&
         Height          =   945
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   4815
   End
      Begin VB.OptionButton optCredito 
         Caption         =   "Sim"
      ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   630
         Width           =   705
      End
   End
   Begin VB.Frame fraContaCre 
      Height          =   1575
      Left            =   30
      TabIndex        =   6
      Top             =   2850
      Width           =   5055
   Begin VB.TextBox txtResAgencia 
      ForeColor       =   &H00800000&
      Height          =   285
         Left            =   1425
      MaxLength       =   4
         TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Agência da conta de débito."
         Top             =   645
      Width           =   800
   End
      Begin VB.TextBox txtResConta 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1425
         MaxLength       =   12
         TabIndex        =   12
         Tag             =   "1"
         ToolTipText     =   "Número da conta corrente para débito."
         Top             =   1035
         Width           =   1200
      End
      Begin VB.ComboBox cboResBanco 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V187A.frx":0000
         Left            =   1440
         List            =   "T46V187A.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Banco da conta de débito."
         Top             =   240
         Width           =   3500
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco: "
      ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Banco da conta de débito."
         Top             =   240
      Width           =   1305
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agência: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
         Left            =   120
         TabIndex        =   9
      ToolTipText     =   "Agência da conta de débito."
         Top             =   645
      Width           =   1305
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
         Caption         =   "Conta corrente: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Número da conta corrente para débito."
         Top             =   1035
         Width           =   1305
      End
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   1553
      TabIndex        =   13
      ToolTipText     =   "Volta à proposta."
      Top             =   4530
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   2633
      TabIndex        =   14
      ToolTipText     =   "Volta à proposta."
      Top             =   4530
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "Sair"
      ForeColor       =   8388608
   End
   Begin VB.Label Label 
      Caption         =   $"T46V187A.frx":0004
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   5025
   End
End
Attribute VB_Name = "frmT46V187A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboResBanco_Click()
    Dim lTamConta As Byte   'Tamanho do campo de conta.

    '1. Habilita salvar.
    If gM46V999.PerfilAcesso <> e_Perfil_Triagem Then
        Call gpHabSalvar
    End If

    '3. Posiciona e verifica campos de agência e conta.
    Select Case Left$(cboResBanco.Text, 3)
    Case "001"    'Projeto 11080 - Debito automatico - Banco do Brasil
        lTamConta = 10
    Case "237"      'Bradesco.
        lTamConta = 10
    Case "409"      'Unibanco.
        lTamConta = 8
    Case "341"      'Itaú.
        lTamConta = 7
    Case "347"      'Sudameris.
        lTamConta = 11
    Case "479"      'Boston.
        lTamConta = 9
    Case "356"      'Real.
        lTamConta = 11
    Case "399"      'HSBC
        lTamConta = 12
    Case "033", "33"
        lTamConta = 10
    Case "047", "47"
        lTamConta = 10
    Case 422 'SAFRA
        lTamConta = 8
    Case 756 'BANCOOB
        lTamConta = 12
    Case Else       'Outros bancos.
        lTamConta = 9
    End Select
    
        txtResConta.MaxLength = lTamConta
End Sub

Private Sub cboResBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub cboResBanco_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then     'Teclou backspace.
        cboResBanco.ListIndex = -1
    End If
End Sub

Private Sub cmdCopiar_Click()
    Dim intC As Integer
    Dim numBanco As String
    Dim numAgencia As String
    Dim numConta As String
    
    numBanco = Empty
    numAgencia = Empty
    numConta = Empty
    
    If frmT46V102A.optForPagamento(1).Value Then
        If frmT46V110A.cboDebBanco.ListIndex > -1 Then
            numBanco = frmT46V110A.cboDebBanco.ItemData(frmT46V110A.cboDebBanco.ListIndex)
            numAgencia = Left(frmT46V110A.txtDebAgencia.Text, 4)
            numConta = frmT46V110A.txtDebConta.Text
        End If
    Else
        If frmT46V102A.optForPagamento(2).Value Then
            If frmT46V185A.cboDebBanco.ListIndex > -1 Then
                numBanco = frmT46V185A.cboDebBanco.ItemData(frmT46V185A.cboDebBanco.ListIndex)
                numAgencia = Left(frmT46V185A.txtDebAgencia.Text, 4)
                numConta = frmT46V185A.txtDebConta.Text
            End If
        End If
    End If
   
    If numBanco <> Empty Then
            For intC = 0 To cboResBanco.ListCount - 1
                If cboResBanco.ItemData(intC) = numBanco Then
                    cboResBanco.ListIndex = intC
                    Exit For
                End If
            Next intC
    
    Else
        cboResBanco.ListIndex = -1
    End If
    
    txtResAgencia.Text = numAgencia
    txtResConta.Text = numConta
End Sub

Private Sub cmdOk_Click()
    Dim clsA46V102 As clsA46V102A      'DLL : consiste dados comuns (nível 2).
    Dim lMensagem As String
    Dim lValida As Boolean

    If optCredito(1).Value Then
        If cboResBanco.ListIndex = -1 Or txtResAgencia.Text = Empty Or txtResConta.Text = Empty Then
            MsgBox "Necessário preencher corretamente os dados para restituição de prêmio", vbCritical, "Atenção"
            cboResBanco.SetFocus
            Exit Sub
        End If
    End If

    '1001180 - Validação de DV na Triagem
    If gM46V999.PerfilAcesso = e_Perfil_Triagem Then
        Set clsA46V102 = New clsA46V102A    'DLL : consiste dados comuns (nível 2).
        Set clsA46V102.M46V999 = gM46V999

        With clsA46V102
            .Pro_IndBanco = cboResBanco.ItemData(cboResBanco.ListIndex)
            .Pro_IndAgencia = Val(Replace(Replace(txtResAgencia.Text, "X", "0"), "-", ""))
            .Pro_IndConta = Val(Replace(Replace(txtResConta.Text, "X", "0"), "-", ""))
            lValida = .mfVerConta(1, cboResBanco.ItemData(cboResBanco.ListIndex), Val(Replace(txtResAgencia.Text, "-", "")), Val(Replace(txtResConta.Text, "-", "")), lMensagem)
            If (lValida = True) And (lMensagem = "") Then
                Me.Hide
            Else
                MsgBox lMensagem, vbCritical, "Atenção"
            End If
        End With
        Set clsA46V102 = Nothing
    
    Else
        Me.Hide
    End If
End Sub

Private Sub cmdSair_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()

    On Error GoTo Erro
    
    If gM46V999.PerfilAcesso <> e_Perfil_Triagem Then
        frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
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
    
    Exit Sub
Erro:
    gM46V999.gpGraLog 1, Err.Number & " - " & Err.Description, vbCritical
    GoTo Finalizar
End Sub

Public Sub mpInstanciarForm(pTipoPagamento As Integer)

    If gM46V999.PerfilAcesso = e_Perfil_Triagem Then
        Me.Show vbModal, frmT46V012A
    Else
        Me.Hide
        Me.Show vbModal, frmT46V102A
    End If

End Sub

Private Sub Form_Load()
    Dim lMensagem As String    'Mensagem.

    '1. Centraliza form.
    Call gM46V999.gpCenForm(Me)

    '2. Carrega combos de bancos.
    If gfCarCboBancos(6, cboResBanco, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    
    optCredito(0).Value = True
End Sub

Private Sub optCredito_Click(index As Integer)

    Select Case index
        Case 0, 2
            txtResAgencia.Text = Empty
            txtResConta.Text = Empty
            cboResBanco.ListIndex = -1
            fraContaCre.Enabled = False
            cmdCopiar.Enabled = False
        Case 1
            fraContaCre.Enabled = True
            cmdCopiar.Enabled = True
    End Select
    
End Sub

Private Sub txtResAgencia_Change()
    If gM46V999.PerfilAcesso <> e_Perfil_Triagem Then
        Call gpHabSalvar
    End If
End Sub
Private Sub txtResAgencia_GotFocus()
    With txtResAgencia
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtResAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtResAgencia_LostFocus()

    If gM46V999.gfPreenchido(txtResAgencia.Text) Then
        txtResAgencia.Text = Right("0000" & txtResAgencia.Text, 4)
    End If
End Sub
Private Sub txtResConta_Change()
    If gM46V999.PerfilAcesso <> e_Perfil_Triagem Then
        Call gpHabSalvar
    End If
End Sub
Private Sub txtResConta_GotFocus()
    With txtResConta
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtResConta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtResConta_LostFocus()
    Dim lAuxiliar As String    'Auxiliar.
    Dim lCodBanco As String    'Código do banco.
    Dim lMensagem As String    'Mensagem.

    If gM46V999.gfPreenchido(txtResConta.Text) Then
        With cboResBanco
            If .ListIndex = -1 Then
                lCodBanco = "000"
            Else
                lCodBanco = Left$(.Text, 3)
            End If
        End With
        
        lAuxiliar = Left(txtResConta.Text, txtResConta.MaxLength - 1)
        lAuxiliar = Replace(lAuxiliar, "-", Empty)
        lAuxiliar = Right(String(txtResConta.MaxLength - 1, "0") & lAuxiliar, txtResConta.MaxLength - 1)
        lAuxiliar = Left(lAuxiliar, Len(lAuxiliar) - 1) & "-" & Right(lAuxiliar, 1)
        txtResConta.Text = lAuxiliar
        
        If gfVerConta(1, lCodBanco, txtResConta.Text, lAuxiliar, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If
        
    End If

End Sub

