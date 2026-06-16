VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V193A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contatos de Inspeção"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8130
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   8130
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPontoReferencia 
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   960
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   1860
      Width           =   7050
   End
   Begin VB.TextBox txtEmail 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   960
      MaxLength       =   60
      TabIndex        =   14
      ToolTipText     =   "Contato de inspeção."
      Top             =   1530
      Width           =   7050
   End
   Begin VB.Frame Frame 
      Caption         =   "Contatos"
      ForeColor       =   &H00800000&
      Height          =   1365
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   7935
      Begin VB.TextBox txtContato 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   885
         MaxLength       =   40
         TabIndex        =   10
         ToolTipText     =   "Contato de inspeção."
         Top             =   930
         Width           =   4620
      End
      Begin VB.TextBox txtTelefone 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   6555
         MaxLength       =   14
         TabIndex        =   12
         ToolTipText     =   "Telefone da inspeção. Formato: (99)9999-9999"
         Top             =   930
         Width           =   1240
      End
      Begin VB.TextBox txtContato 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   885
         MaxLength       =   40
         TabIndex        =   6
         ToolTipText     =   "Contato de inspeção."
         Top             =   600
         Width           =   4620
      End
      Begin VB.TextBox txtTelefone 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   6555
         MaxLength       =   14
         TabIndex        =   8
         ToolTipText     =   "Telefone da inspeção. Formato: (99)9999-9999"
         Top             =   600
         Width           =   1240
      End
      Begin VB.TextBox txtContato 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   885
         MaxLength       =   40
         TabIndex        =   2
         ToolTipText     =   "Contato de inspeção."
         Top             =   270
         Width           =   4620
      End
      Begin VB.TextBox txtTelefone 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   6555
         MaxLength       =   14
         TabIndex        =   4
         ToolTipText     =   "Telefone da inspeção. Formato: (99)9999-9999"
         Top             =   270
         Width           =   1240
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nome 3:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   150
         TabIndex        =   9
         Top             =   930
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefone 3: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5610
         TabIndex        =   11
         Top             =   930
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nome 2:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   150
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefone 2: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5610
         TabIndex        =   7
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblContato 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nome 1:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   735
      End
      Begin VB.Label lblTelefone 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefone 1: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5580
         TabIndex        =   3
         Top             =   270
         Width           =   975
      End
   End
   Begin Threed.SSCommand cmdCopiarDadoInspecao 
      Height          =   480
      Left            =   60
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Copiar dados de inspeção para os demais itens."
      Top             =   2430
      Width           =   1260
      _Version        =   65536
      _ExtentX        =   2222
      _ExtentY        =   847
      _StockProps     =   78
      Caption         =   "Replicar"
      ForeColor       =   12582912
      Picture         =   "T46V193A.frx":0000
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   480
      Left            =   1380
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Copiar dados de inspeção para os demais itens."
      Top             =   2430
      Width           =   1260
      _Version        =   65536
      _ExtentX        =   2222
      _ExtentY        =   847
      _StockProps     =   78
      Caption         =   "Ok"
      ForeColor       =   12582912
      Picture         =   "T46V193A.frx":001C
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ponto de Referência:"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   60
      TabIndex        =   15
      Top             =   1860
      Width           =   885
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Email:"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   60
      TabIndex        =   13
      Top             =   1530
      Width           =   885
   End
End
Attribute VB_Name = "frmT46V193A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'1000477 - Telefone de Contato Formatado
Private Tecla As Integer

Public Sub CarregarDados(ByRef pobjstcPedInsp As stcA46V717B)

    'Contato da Inspeção
    txtContato(0).Text = gM46V999.gfRetirarCaracterEspecial(pobjstcPedInsp.Pess_Cont, False, True, True)
    txtContato(1).Text = gM46V999.gfRetirarCaracterEspecial(pobjstcPedInsp.PESS_CONTATO_2, False, True, True)
    txtContato(2).Text = gM46V999.gfRetirarCaracterEspecial(pobjstcPedInsp.PESS_CONTATO_3, False, True, True)
    
    'Telefone da inspeção
    txtTelefone(0).Text = pobjstcPedInsp.Num_Tel
    txtTelefone(1).Text = pobjstcPedInsp.NUM_TEL_2
    txtTelefone(2).Text = pobjstcPedInsp.NUM_TEL_3
    
    txtEmail.Text = pobjstcPedInsp.DSC_EMAIL
    txtPontoReferencia.Text = pobjstcPedInsp.DSC_PONTO_REF
End Sub

Public Sub SalvarDados(ByRef pobjstcPedInsp As stcA46V717B)
    
    With pobjstcPedInsp
        .Num_Tel = txtTelefone(0).Text
        .NUM_TEL_2 = txtTelefone(1).Text
        .NUM_TEL_3 = txtTelefone(2).Text
        .Pess_Cont = gM46V999.gfRetirarCaracterEspecial(txtContato(0).Text, False, True, True)
        .PESS_CONTATO_2 = gM46V999.gfRetirarCaracterEspecial(txtContato(1).Text, False, True, True)
        .PESS_CONTATO_3 = gM46V999.gfRetirarCaracterEspecial(txtContato(2).Text, False, True, True)
        .DSC_EMAIL = txtEmail.Text
        .DSC_PONTO_REF = txtPontoReferencia.Text
    End With
End Sub

Public Function mfAlterarDadosInspecao(ByRef lMensagem As String, ByRef pobjstcPedLoc As stcA46V708B, ByVal pNumInsp As String) As Boolean
    Dim lUpdate As String
    Dim lbdBasDados As ADODB.Connection    'Base de dados.

    On Error GoTo ControleErros

    If Not gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, lbdBasDados, lMensagem) Then
        GoTo ControleErros
    End If

    lUpdate = " Update Tab_Ped_Loc SET Num_Insp = " & pNumInsp
    lUpdate = lUpdate & " WHERE Nosso_Numero = " & gM46V999.gfFormatarCamposSQL(pobjstcPedLoc.NOSSO_NUMERO) & " AND Num_Item = " & pobjstcPedLoc.Num_Item

    If gM46V999.gfExeSQL(lbdBasDados, lUpdate, lMensagem) <> 0 Then GoTo ControleErros

    lUpdate = " Update Tab_Ped_Insp SET Num_Tel = " & gM46V999.gfFormatarCamposSQL(txtTelefone(0).Text) & ", Pess_Cont = " & gM46V999.gfFormatarCamposSQL(gM46V999.gfRetirarCaracterEspecial(txtContato(0).Text, False, True, True))
    lUpdate = lUpdate & " WHERE Nosso_Numero = " & gM46V999.gfFormatarCamposSQL(pobjstcPedLoc.NOSSO_NUMERO) & " AND Num_Item = " & pobjstcPedLoc.Num_Item

    If gM46V999.gfExeSQL(lbdBasDados, lUpdate, lMensagem) <> 0 Then GoTo ControleErros

    lUpdate = " Update tab_ped_insp_compl SET NUM_TEL_2 = " & gM46V999.gfFormatarCamposSQL(txtTelefone(1).Text) & ", PESS_CONTATO_2 = " & gM46V999.gfFormatarCamposSQL(gM46V999.gfRetirarCaracterEspecial(txtContato(1).Text, False, True, True))
    lUpdate = lUpdate & " ,NUM_TEL_3 = " & gM46V999.gfFormatarCamposSQL(txtTelefone(2).Text) & ", PESS_CONTATO_3 = " & gM46V999.gfFormatarCamposSQL(gM46V999.gfRetirarCaracterEspecial(txtContato(2).Text, False, True, True))
    lUpdate = lUpdate & " WHERE Nosso_Numero = " & gM46V999.gfFormatarCamposSQL(pobjstcPedLoc.NOSSO_NUMERO) & " AND Num_Item = " & pobjstcPedLoc.Num_Item

    If gM46V999.gfExeSQL(lbdBasDados, lUpdate, lMensagem) <> 0 Then GoTo ControleErros
    mfAlterarDadosInspecao = True

    Exit Function
ControleErros:
    mfAlterarDadosInspecao = False
    If Err.Number > 0 Then
        lMensagem = "mfAlterarDadosInspecao ERRO - " & Err.Number & " " & Err.Description & _
                     vbCrLf & " SQL: " & lUpdate
    Else
        lMensagem = "mfAlterarDadosInspecao ERRO - " & lMensagem & _
             vbCrLf & " SQL: " & lUpdate
    End If
End Function

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub txtTelefone_LostFocus(index As Integer)
    With txtTelefone(index)
        If InStr(1, .Text, "-", vbTextCompare) > 0 Then Exit Sub
        Select Case Len(.Text)
        Case 12    '(99)9999-9999
            .Text = Left(.Text, 8) & "-" & Right(.Text, 4)
        Case 13    '(99)99999-9999
            .Text = Left(.Text, 9) & "-" & Right(.Text, 4)
        End Select
    End With
End Sub

Private Sub txtTelefone_Change(index As Integer)
    Call gpLimPremios
    '1000477 - Telefone de Contato Formatado
    If Tecla = vbKeyBack Then Exit Sub
    
    With txtTelefone(index)
        Select Case Len(.Text)
        Case 1
            If InStr(1, .Text, "(", vbTextCompare) = 0 Then
                .Text = "(" & .Text
                .SelStart = Len(.Text)
            End If
        Case 3
            If InStr(1, .Text, ")", vbTextCompare) = 0 Then
                .Text = .Text & ")"
                .SelStart = Len(.Text)
            End If
        End Select
    End With
End Sub

Private Sub txtTelefone_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    '1000477 - Telefone de Contato Formatado
    With txtTelefone(index)
        Select Case KeyCode
        Case vbKeyDelete
            If .SelLength <> Len(.Text) Then
                Select Case Mid(.Text, .SelStart + 1, 1)
                Case "(", ")", "-"
                    KeyCode = 0
                End Select
            End If
        Case vbKeyBack
            If .SelStart > 0 And .SelStart <> Len(.Text) Then
                Select Case Mid(.Text, .SelStart, 1)
                Case "(", ")", "-"
                    KeyCode = 0
                    .SelStart = .SelStart - 1
                End Select
            End If
        End Select
    End With
End Sub

Private Sub txtTelefone_GotFocus(index As Integer)
    With txtTelefone(index)
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtTelefone_KeyPress(index As Integer, KeyAscii As Integer)
'1000477 - Telefone de Contato Formatado
    Tecla = KeyAscii
    
    With txtTelefone(index)
        Select Case KeyAscii
        Case 48 To 57   'permite de 0 a 9
            'limita tamanho do campo
            If Len(.Text) >= 14 And .SelLength <> Len(.Text) Then KeyAscii = 0: Exit Sub   '1200178 - Telefone 9 Digitos
            'se foco estiver no inicio e o campo ainda nao foi todo preenchido
            If .SelStart <= 3 And Len(.Text) > 0 And .SelLength <> Len(.Text) Then
                If .SelStart <= 3 And Mid(.Text, 4, 1) = ")" Then
                    KeyAscii = 0: Exit Sub
                End If
                .Text = Mid(.Text, 1, IIf(.SelStart = 0, 1, .SelStart)) & Chr(KeyAscii) & Mid(.Text, IIf(.SelStart = 0, 1, .SelStart) + 1, Len(.Text))
                .SelStart = IIf(.SelStart <= 1, .SelStart + 2, Len(.Text))
                KeyAscii = 0
            End If
            
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{TAB}"
        Case vbKeyBack  'permite back
        Case Else   'nao permite demais teclas
            KeyAscii = 0
        End Select
    End With
End Sub

Private Sub txtContato_Change(index As Integer)
    Call gpLimPremios
End Sub

Private Sub txtContato_GotFocus(index As Integer)
    With txtContato(index)
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtContato_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtContato_LostFocus(index As Integer)
    txtContato(index).Text = gM46V999.gfRetirarCaracterEspecial(txtContato(index).Text, True, False, True)
End Sub

Private Sub cmdCopiarDadoInspecao_Click()
    Call frmT46V101A.ReplicarDadoInspecao
End Sub
