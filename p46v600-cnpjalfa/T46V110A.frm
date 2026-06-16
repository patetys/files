VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V110A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Débito em conta"
   ClientHeight    =   1920
   ClientLeft      =   1350
   ClientTop       =   2370
   ClientWidth     =   5010
   ControlBox      =   0   'False
   Icon            =   "T46V110A.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboDebBanco 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "T46V110A.frx":000C
      Left            =   1400
      List            =   "T46V110A.frx":000E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Banco da conta de débito."
      Top             =   100
      Width           =   3500
   End
   Begin VB.TextBox txtDebAgencia 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1400
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Agência da conta de débito."
      Top             =   500
      Width           =   800
   End
   Begin VB.TextBox txtDebConta 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1400
      MaxLength       =   12
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Número da conta corrente para débito."
      Top             =   900
      Width           =   1200
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   1905
      TabIndex        =   6
      ToolTipText     =   "Volta à proposta."
      Top             =   1530
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin VB.Label lblTexto 
      Caption         =   "Somente para contas originadas no Banco Real."
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   90
      TabIndex        =   7
      Top             =   1260
      Visible         =   0   'False
      Width           =   4845
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Banco: "
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   0
      Left            =   100
      TabIndex        =   0
      ToolTipText     =   "Banco da conta de débito."
      Top             =   100
      Width           =   1300
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agência: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   100
      TabIndex        =   2
      ToolTipText     =   "Agência da conta de débito."
      Top             =   500
      Width           =   1300
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conta corrente: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   100
      TabIndex        =   4
      ToolTipText     =   "Número da conta corrente para débito."
      Top             =   900
      Width           =   1300
   End
End
Attribute VB_Name = "frmT46V110A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDebBanco_Click()
    Dim lAuxiliar As String    'Auxiliar.
    Dim lCodBanco As String    'Código do banco.
    Dim lMensagem As String    'Mensagem.
    Dim lTamAgencia As Byte     'Tamanho do campo de agência.
    Dim lTamConta As Byte   'Tamanho do campo de conta.

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

    lblTexto.Visible = False    '1200429 - Tarifa 1036 e 1037 Maio 2012 - Item 10

    '3. Posiciona e verifica campos de agência e conta.
    Select Case lCodBanco
    Case "001"    'Projeto 11080 - Debito automatico - Banco do Brasil
        lTamAgencia = 6
        lTamConta = 10
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
        lblTexto.Caption = "Somente para contas originadas no Banco Real."
        lblTexto.Visible = True    '1200429 - Tarifa 1036 e 1037 Maio 2012 - Item 10
    Case "399"      'HSBC
        lTamAgencia = 4
        lTamConta = 12
    Case "033", "33"
        lTamAgencia = 4
        lTamConta = 10
    Case "047", "47"
        lTamAgencia = 4
        lTamConta = 10
    Case 422 'SAFRA
        lTamAgencia = 4
        lTamConta = 8
    Case 756 'BANCOOB
        lTamAgencia = 4
        lTamConta = 10
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
    
    If Not gCarga And Me.Visible Then
        If lCodBanco = "001" Then
        
            MsgBox "Você optou por pagamento em débito automático através do Banco do Brasil. Por exigência da instituição bancária, o débito deve ser autorizado pelo titular da conta junto aos canais do Banco. Essa autorização será necessária somente para a primeira parcela do seguro. Se não houver autorização junto ao Banco, o débito não será efetuado o que implicará em cancelamento do seguro." & _
                   vbLf & "É importante manter os dados cadastrais atualizados para que as comunicações de débito e demais informações não sejam interrompidas.", vbInformation, Me.Caption
            
        End If
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
Private Sub cmdOk_Click()
    Dim clsA46V102 As clsA46V102A      'DLL : consiste dados comuns (nível 2).
    Dim lMensagem As String
    Dim lValida As Boolean

    '1001180 - Validação de DV na Triagem
    If gM46V999.PerfilAcesso = e_Perfil_Triagem Then
        Set clsA46V102 = New clsA46V102A    'DLL : consiste dados comuns (nível 2).
        Set clsA46V102.M46V999 = gM46V999

        With clsA46V102
            .Pro_IndBanco = cboDebBanco.ItemData(cboDebBanco.ListIndex)
            .Pro_IndAgencia = Val(Replace(Replace(txtDebAgencia.Text, "X", "0"), "-", ""))
            .Pro_IndConta = Val(Replace(Replace(txtDebConta.Text, "X", "0"), "-", ""))
            lValida = .mfVerConta(1, cboDebBanco.ItemData(cboDebBanco.ListIndex), Val(Replace(txtDebAgencia.Text, "-", "")), Val(Replace(txtDebConta.Text, "-", "")), lMensagem)
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

Private Sub Form_Activate()
    Call Form_Load

    If gM46V999.PerfilAcesso <> e_Perfil_Triagem Then
        frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    End If

    If cboDebBanco.Visible = True And cboDebBanco.Enabled = True Then
        cboDebBanco.SetFocus
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
Public Sub mpInstanciarForm(pTipoPagamento As Integer)

    If gM46V999.PerfilAcesso = e_Perfil_Triagem Then
        Me.Show vbModal, frmT46V012A
    Else
        Me.Show vbModal, frmT46V102A
    End If

End Sub

Private Sub Form_Load()
    Dim lMensagem As String    'Mensagem.
    Dim lCpfCnpj As String
    Dim lDatEmiss As Long
    Dim lOpcao As Byte
    Dim lCDebSel As Integer

'1. Centraliza form.
    Call gM46V999.gpCenForm(Me)

    '2 Carrega combos de bancos.
    With frmT46V101A
        If .objstcPedLoc.Tip_Emissao <= 200 Then
            lDatEmiss = 0
        Else
            lDatEmiss = .objstcPedido.gStcPedidoApolice.Dat_Emis
        End If
        
        '[Verifica se PF/PJ]
        If frmT46V102A.optTipPessoa(0).Value = True Then    '[Pessoa física]
            Call gM46V999.gpForCPF(2, frmT46V102A.txtCPF.Text, lCpfCnpj)
        Else
            Call gM46V999.gpForCNPJ(2, frmT46V102A.txtCNPJ.Text, lCpfCnpj)
        End If
        
        If gM46V999.mfObtemBancoBoleto(.objstcPedLoc.Tip_Emissao, _
                                       lCpfCnpj, _
                                       .objstcPedido.Cod_Produto, _
                                       .objstcPedido.Cod_Segurado, _
                                       .objstcPedido.Num_Apol, _
                                       gM46V999.CDblx(frmT46V102A.txtCorretor(1).Tag), _
                                       lDatEmiss) Then
            lOpcao = 5
        Else
            lOpcao = 3
        End If
    End With
    
    lCDebSel = cboDebBanco.ListIndex
    If gfCarCboBancos(lOpcao, cboDebBanco, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    cboDebBanco.ListIndex = lCDebSel
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
    Dim lAuxiliar As String    'Auxiliar.
    Dim lCodBanco As String    'Código do banco.
    Dim lMensagem As String    'Mensagem.

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
            txtDebAgencia.Text = ""
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
    Dim lAuxiliar As String    'Auxiliar.
    Dim lCodBanco As String    'Código do banco.
    Dim lMensagem As String    'Mensagem.

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
