VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V175B 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dados obrigatórios para inspeção"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand cmdCancelar 
      Height          =   360
      Left            =   3735
      TabIndex        =   4
      Top             =   1380
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1879
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
   Begin VB.Frame FrmDados 
      ForeColor       =   &H00800000&
      Height          =   1900
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4905
      Begin VB.TextBox txtCNPJ 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1830
         MaxLength       =   18
         TabIndex        =   2
         ToolTipText     =   "CNPJ do segurado."
         Top             =   885
         Width           =   1600
      End
      Begin VB.TextBox txtTelefone 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1830
         MaxLength       =   14
         TabIndex        =   0
         ToolTipText     =   "Telefone da inspeção. Formato: (99)9999-9999"
         Top             =   285
         Width           =   1200
      End
      Begin VB.TextBox txtContato 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   1
         ToolTipText     =   "Contato de inspeção."
         Top             =   585
         Width           =   2520
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   360
         Left            =   2415
         TabIndex        =   3
         Top             =   1380
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   635
         _StockProps     =   78
         Caption         =   "&Ok"
         ForeColor       =   8388608
      End
      Begin VB.Label lblCPFCNPJ 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CPF/CNPJ:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   900
         Width           =   1665
      End
      Begin VB.Label lblTelefone 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefone Inspeção: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   270
         Width           =   1665
      End
      Begin VB.Label lblContato 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Contato: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   585
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmT46V175B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objPedido As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED)
Private objPedLoc As clsA46V708A    'Objeto da classe clsA46V708A (TAB_PED_LOC)
Private objPedInsp As clsA46V717A    'Objeto da classe clsA46V717A (TAB_PED_INSP)

Private objstcPedido As stcA46V702B   'Objeto da classe de estrutura stcA46V702B (TAB_PED)
Private objstcPedLoc As stcA46V708B   'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
Private objstcPedInsp As stcA46V717B    'Objeto da classe de estrutura stcA46V717B (TAB_PED_INSP)

Private mAntControle As Control   'Anterior: controle.
Private Tecla As Integer
Private lTipoPessoa As Integer

Private Sub cmdCancelar_Click()
    frmT46V174A.gAbortSolic = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim clsA99V001 As clsA99V001    'Cálculo de dígitos.
    Dim pMensagem As String
    Dim lContato As String
    Dim lAuxiliar As String
    Dim lDig As String

    If gM46V999.gfPreenchido(txtTelefone.Text) = False Then
        MsgBox ("FGE0105 - Telefone de contato para inspeção inválido (não preenchido).")
        Exit Sub
    Else
        '1000477 - Telefone de Contato Formatado
        '1 - valida se esta totalmente preenchido
        If Len(txtTelefone.Text) < 13 Then
            MsgBox ("FGE0105 - Telefone de contato para inspeção inválido (Favor informar o DDD e o número do telefone com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
            Exit Sub
        Else
            '2 - valida formado (99)9999-9999
            If Mid(txtTelefone.Text, 1, 1) <> "(" Or Mid(txtTelefone.Text, 4, 1) <> ")" Or Mid(txtTelefone.Text, 9, 1) <> "-" Then
                MsgBox ("FGE0105 - Telefone de contato para inspeção inválido (Favor informar o DDD e o número do telefone com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                Exit Sub
            Else
                '3 - não permite todos os numeros do telefone iguais
                If InStr(1, "0000-0000 1111-1111 2222-2222 3333-3333 4444-4444 5555-5555 6666-6666 7777-7777 8888-8888 9999-9999", Mid(txtTelefone.Text, 5, Len(txtTelefone.Text)), vbTextCompare) > 0 Then
                    MsgBox ("FGE0105 - Telefone de contato para inspeção inválido (Favor informar o DDD e o número do telefone com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                    Exit Sub
                End If
            End If
        End If
    End If

    lContato = gM46V999.gfRetirarCaracterEspecial(txtContato.Text, True, False, True)
    txtContato.Text = lContato
    If Len(Trim(lContato)) < 4 Then
        MsgBox "O Contato deve ter pelo menos 4 caracteres. Não é possível solicitar o Agendamento.", vbCritical
        Exit Sub
    End If

    'Ficha 1100102
    If Val(txtCNPJ.Text) <> 0 Then
        If Not gM46V999.gfPreenchido(txtCNPJ.Text) Then
            MsgBox "FNP0175 - CPF/CNPJ do segurado não preenchido.", vbCritical
            Exit Sub
        End If

        Select Case lTipoPessoa
        Case 1  'Pessoa jurídica.
            lAuxiliar = gM46V001.gfForCNPJ(txtCNPJ.Text)
            lDig = Right(lAuxiliar, 2)
            Set clsA99V001 = New clsA99V001     'Calcula dígitos.
            Set clsA99V001.M46V999 = gM46V999

            If clsA99V001.DIG_CNPJ(lAuxiliar, lDig, pMensagem) = False Then
                MsgBox "FGE0175 - CNPJ inválido.", vbCritical
                Exit Sub
            End If
            Set clsA99V001 = Nothing

        Case 2  'Pessoa física.
            lAuxiliar = gM46V001.gfForCPF(txtCNPJ.Text)
            lDig = Right(lAuxiliar, 2)
            Set clsA99V001 = New clsA99V001     'Calcula dígitos.
            Set clsA99V001.M46V999 = gM46V999

            If Not clsA99V001.DIG_CPF(lAuxiliar, lDig, pMensagem) Then
                MsgBox "FGE0175 - CPF inválido.", vbCritical
                Exit Sub
            End If
            Set clsA99V001 = Nothing
        End Select
    Else
        txtCNPJ.Text = " "
    End If

    If gM46V999.gfPreenchido(txtTelefone.Text) = False Or gM46V999.gfPreenchido(txtContato.Text) = False Or gM46V999.gfPreenchido(txtCNPJ.Text) = False Then
        MsgBox "Informações necessárias para a inspeção não foram preenchidas. Não é possível solicitar o Agendamento.", vbCritical
        Exit Sub
    End If


    frmT46V174A.gAbortSolic = False
    Me.Hide
    '    Unload Me
End Sub

Private Sub Form_Load()
    Dim lobjstcPedLoc As stcA46V708B
    Dim lCPF_CNPJ As String

    Screen.MousePointer = vbDefault
    Call gM46V999.gpCenForm(Me)

    '1. Instancio o objeto.
    Set objPedido = New clsA46V702A
    Set objPedido.M46V999 = gM46V999

    Set objstcPedido = New stcA46V702B
    Set objstcPedido.M46V999 = gM46V999

    objstcPedido.NOSSO_NUMERO = frmT46V174A.gNosNumeroInsp
    objstcPedido.Num_Ped = frmT46V174A.gNumPedidoInsp

    '2. Busco os pedidos.
    If Not objPedido.mfCarregarColecao(objstcPedido, "", "", True) Then
        Call gM46V999.gpGraLog(1, objPedido.UltimoErro)
        Exit Sub
    End If

    For Each objstcPedido In objPedido.Colecao
        For Each lobjstcPedLoc In objstcPedido.ColecaoPedLoc
            If lobjstcPedLoc.Num_Item = frmT46V174A.gNumItemInsp Then
                txtTelefone.Text = lobjstcPedLoc.objstcPedInsp.Num_Tel
                txtContato.Text = gM46V999.gfRetirarCaracterEspecial(lobjstcPedLoc.objstcPedInsp.Pess_Cont, True, False, True)
                txtCNPJ.Text = Trim(objstcPedido.gCNPJ_CPF_Formatado)
                lTipoPessoa = objstcPedido.Tip_Pessoa
                Set objstcPedLoc = lobjstcPedLoc
            End If
        Next lobjstcPedLoc
    Next objstcPedido

End Sub


Private Sub txtCNPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtContato_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtContato_LostFocus()
    txtContato.Text = gM46V999.gfRetirarCaracterEspecial(txtContato.Text, True, False, True)
End Sub

Private Sub txtTelefone_KeyPress(KeyAscii As Integer)
'Telefone de Contato Formatado
    Tecla = KeyAscii
    Select Case KeyAscii
    Case 48 To 57   'permite de 0 a 9
        'limita tamanho do campo
        If Len(txtTelefone.Text) >= 14 And txtTelefone.SelLength <> Len(txtTelefone.Text) Then KeyAscii = 0: Exit Sub   '1200178 - Telefone 9 Digitos
        'se foco estiver no inicio e o campo ainda nao foi todo preenchido
        If txtTelefone.SelStart <= 3 And Len(txtTelefone.Text) > 0 And txtTelefone.SelLength <> Len(txtTelefone.Text) Then
            If txtTelefone.SelStart <= 3 And Mid(txtTelefone.Text, 4, 1) = ")" Then
                KeyAscii = 0: Exit Sub
            End If
            txtTelefone.Text = Mid(txtTelefone.Text, 1, IIf(txtTelefone.SelStart = 0, 1, txtTelefone.SelStart)) & Chr(KeyAscii) & Mid(txtTelefone.Text, IIf(txtTelefone.SelStart = 0, 1, txtTelefone.SelStart) + 1, Len(txtTelefone.Text))
            txtTelefone.SelStart = IIf(txtTelefone.SelStart <= 1, txtTelefone.SelStart + 2, Len(txtTelefone.Text))
            KeyAscii = 0
            '            Else   '1200178 - Telefone 9 Digitos (desabilitado)
            '                If Mid(txtTelefone.Text, 9, 1) = "-" And txtTelefone.SelStart < 9 Then KeyAscii = 0: Exit Sub
        End If
    Case vbKeyReturn
        KeyAscii = 0
        SendKeys "{TAB}"
    Case vbKeyBack  'permite back
    Case Else   'nao permite demais teclas
        KeyAscii = 0
    End Select
End Sub

Private Sub txtTelefone_Change()
'Telefone de Contato Formatado
    If Tecla = vbKeyBack Then Exit Sub
    Select Case Len(txtTelefone.Text)
    Case 1
        If InStr(1, txtTelefone.Text, "(", vbTextCompare) = 0 Then
            txtTelefone.Text = "(" & txtTelefone.Text
            txtTelefone.SelStart = Len(txtTelefone.Text)
        End If
    Case 3
        If InStr(1, txtTelefone.Text, ")", vbTextCompare) = 0 Then
            txtTelefone.Text = txtTelefone.Text & ")"
            txtTelefone.SelStart = Len(txtTelefone.Text)
        End If
    End Select
End Sub

Private Sub txtTelefone_KeyDown(KeyCode As Integer, Shift As Integer)
'Telefone de Contato Formatado
    Select Case KeyCode
    Case vbKeyDelete
        If txtTelefone.SelLength <> Len(txtTelefone.Text) Then
            Select Case Mid(txtTelefone.Text, txtTelefone.SelStart + 1, 1)
            Case "(", ")", "-"
                KeyCode = 0
            End Select
        End If
    Case vbKeyBack
        If txtTelefone.SelStart > 0 And txtTelefone.SelStart <> Len(txtTelefone.Text) Then
            Select Case Mid(txtTelefone.Text, txtTelefone.SelStart, 1)
            Case "(", ")", "-"
                KeyCode = 0
                txtTelefone.SelStart = txtTelefone.SelStart - 1
            End Select
        End If
    End Select
End Sub

Private Sub txtTelefone_GotFocus()
    Set mAntControle = ActiveControl
    With txtTelefone
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtCNPJ_LostFocus()
    Dim lAuxiliar As String    'Auxiliar.

    If lTipoPessoa = e_TipPessoa_Fisica Then
        Call gM46V999.gpForCPF(1, txtCNPJ.Text, lAuxiliar)
    Else
        Call gM46V999.gpForCNPJ(1, txtCNPJ.Text, lAuxiliar)
    End If

    If lAuxiliar <> "" Then
        txtCNPJ.Text = lAuxiliar
    Else
        If Len(txtCNPJ.Text) = 18 And InStr(1, txtCNPJ.Text, ".") = 0 Then
            txtCNPJ.Text = ""
        Else
            If lAuxiliar <> "" Then
                txtCNPJ.Text = lAuxiliar
            End If
        End If
    End If
End Sub

'1200178 - Telefone 9 Digitos
Private Sub txtTelefone_LostFocus()
    With txtTelefone
        If InStr(1, .Text, "-", vbTextCompare) > 0 Then Exit Sub
        Select Case Len(.Text)
        Case 12    '(99)9999-9999
            .Text = Left(.Text, 8) & "-" & Right(.Text, 4)
        Case 13    '(99)99999-9999
            .Text = Left(.Text, 9) & "-" & Right(.Text, 4)
        End Select
    End With
End Sub

