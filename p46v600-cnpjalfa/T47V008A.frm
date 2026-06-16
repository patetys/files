VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT47V008A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cosseguro"
   ClientHeight    =   2460
   ClientLeft      =   1860
   ClientTop       =   1545
   ClientWidth     =   5190
   Icon            =   "T47V008A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboConSeguradora 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   1
      ItemData        =   "T47V008A.frx":000C
      Left            =   240
      List            =   "T47V008A.frx":000E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      ToolTipText     =   "Seguradora para confirmação do bônus."
      Top             =   375
      Width           =   3855
   End
   Begin VB.ComboBox cboConSeguradora 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   2
      ItemData        =   "T47V008A.frx":0010
      Left            =   240
      List            =   "T47V008A.frx":0012
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      ToolTipText     =   "Seguradora para confirmação do bônus."
      Top             =   675
      Width           =   3855
   End
   Begin VB.ComboBox cboConSeguradora 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   3
      ItemData        =   "T47V008A.frx":0014
      Left            =   240
      List            =   "T47V008A.frx":0016
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      ToolTipText     =   "Seguradora para confirmação do bônus."
      Top             =   975
      Width           =   3855
   End
   Begin VB.ComboBox cboConSeguradora 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   4
      ItemData        =   "T47V008A.frx":0018
      Left            =   240
      List            =   "T47V008A.frx":001A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      ToolTipText     =   "Seguradora para confirmação do bônus."
      Top             =   1275
      Width           =   3855
   End
   Begin VB.ComboBox cboConSeguradora 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   5
      ItemData        =   "T47V008A.frx":001C
      Left            =   240
      List            =   "T47V008A.frx":001E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      ToolTipText     =   "Seguradora para confirmação do bônus."
      Top             =   1575
      Width           =   3855
   End
   Begin VB.TextBox txtCosPercentual 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   1
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "0,00"
      Top             =   375
      Width           =   1000
   End
   Begin VB.TextBox txtCosPercentual 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   2
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "0,00"
      Top             =   675
      Width           =   1000
   End
   Begin VB.TextBox txtCosPercentual 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   3
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "0,00"
      Top             =   975
      Width           =   1000
   End
   Begin VB.TextBox txtCosPercentual 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   4
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "0,00"
      Top             =   1275
      Width           =   1000
   End
   Begin VB.TextBox txtCosPercentual 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   5
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "0,00"
      Top             =   1575
      Width           =   1000
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   300
      Left            =   2040
      TabIndex        =   5
      ToolTipText     =   "Volta à proposta."
      Top             =   2055
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCheck chkCosseguro 
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Indicador de corretor."
      Top             =   435
      Width           =   195
      _Version        =   65536
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   78
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCheck chkCosseguro 
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Indicador de corretor."
      Top             =   735
      Width           =   195
      _Version        =   65536
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   78
      ForeColor       =   -2147483642
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCheck chkCosseguro 
      Height          =   195
      Index           =   3
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "Indicador de corretor."
      Top             =   1035
      Width           =   195
      _Version        =   65536
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   78
      Caption         =   "7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCheck chkCosseguro 
      Height          =   195
      Index           =   4
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "Indicador de corretor."
      Top             =   1335
      Width           =   195
      _Version        =   65536
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   78
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCheck chkCosseguro 
      Height          =   195
      Index           =   5
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "Indicador de corretor."
      Top             =   1635
      Width           =   195
      _Version        =   65536
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   78
      ForeColor       =   -2147483642
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Percentual"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   17
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cosseguro"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   15
      Width           =   3765
   End
End
Attribute VB_Name = "frmT47V008A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mpExcCosseguro(ByVal pLinha As Byte)
    'Procedure: excluir cosseguro (deixa linha de cosseguro em branco).

    'Parâmetro...pLinha...Linha onde será excluído cosseguro.

    Call gpHabSalvar    'Posiciona controles de salvar e desfazer.
    chkCosseguro(pLinha).Value = False
    cboConSeguradora(pLinha).ListIndex = -1
    cboConSeguradora(pLinha).Enabled = False
    txtCosPercentual(pLinha).Enabled = False
    txtCosPercentual(pLinha).Text = "0,00"
End Sub
Private Sub mpIncCosseguro(ByVal pLinha As Byte)
    'Procedure: incluir cosseguro (obtém dados do cosseguro no form T46V117A).

    'Parâmetro...pLinha...Linha onde será incluído cosseguro.

    Call gpHabSalvar        'Posiciona controles de salvar e desfazer.
    chkCosseguro(pLinha).Value = True
    
    cboConSeguradora(pLinha).Enabled = True
    txtCosPercentual(pLinha).Enabled = True
    txtCosPercentual(pLinha).Text = "0,00"
    gIndCosseguro = pLinha
    If gCarga = False Then
        If cboConSeguradora(pLinha).Enabled And cboConSeguradora(pLinha).Visible Then
            cboConSeguradora(pLinha).SetFocus
        End If
    End If
End Sub

Private Sub chkCosseguro_Click(Index As Integer, Value As Integer)
    Call gpHabSalvar
    Select Case Value
        Case False
            Call mpExcCosseguro(Index)
        Case True
            Call mpIncCosseguro(Index)
    End Select
End Sub
Private Sub cmdOK_Click()
Dim lCos1               As Integer
Dim lCos2               As Integer
Dim lCos3               As Integer
Dim lCos4               As Integer
Dim lCos5               As Integer

    'verifico se há duplicidade
    If Me.cboConSeguradora(1).ListIndex <> -1 Then
        lCos1 = cboConSeguradora(1).ItemData(cboConSeguradora(1).ListIndex)
    End If
    If Me.cboConSeguradora(2).ListIndex <> -1 Then
        lCos2 = cboConSeguradora(2).ItemData(cboConSeguradora(2).ListIndex)
    End If
    If Me.cboConSeguradora(3).ListIndex <> -1 Then
        lCos3 = cboConSeguradora(3).ItemData(cboConSeguradora(3).ListIndex)
    End If
    If Me.cboConSeguradora(4).ListIndex <> -1 Then
        lCos4 = cboConSeguradora(4).ItemData(cboConSeguradora(4).ListIndex)
    End If
    If Me.cboConSeguradora(5).ListIndex <> -1 Then
        lCos5 = cboConSeguradora(5).ItemData(cboConSeguradora(5).ListIndex)
    End If
    
    If lCos1 <> 0 Then
        If lCos1 = lCos2 Or lCos1 = lCos3 Or lCos1 = lCos4 Or lCos1 = lCos5 Then
            Call gM46V999.gpGraLog(1, "FGE0092 - Cosseguro inválido. Duplicidade de seguradoras no cosseguro.")
            Exit Sub
        End If
    End If
    
    If lCos2 <> 0 Then
        If lCos2 = lCos3 Or lCos2 = lCos4 Or lCos2 = lCos5 Then
            Call gM46V999.gpGraLog(1, "FGE0092 - Cosseguro inválido. Duplicidade de seguradoras no cosseguro.")
            Exit Sub
        End If
    End If
    
    If lCos3 <> 0 Then
        If lCos3 = lCos4 Or lCos3 = lCos5 Then
            Call gM46V999.gpGraLog(1, "FGE0092 - Cosseguro inválido. Duplicidade de seguradoras no cosseguro.")
            Exit Sub
        End If
    End If
    
    If lCos4 <> 0 Then
        If lCos4 = lCos5 Then
            Call gM46V999.gpGraLog(1, "FGE0092 - Cosseguro inválido. Duplicidade de seguradoras no cosseguro.")
            Exit Sub
        End If
    End If
    
    Me.Hide
End Sub
Private Sub Form_Activate()
    
Finalizar:
    If gM46V999.gDesabilitarForms Or gM46V999.gDesabFormsCotacao Then
        If Me.Tag <> "Desabilitado" Then
            DesabilitaControles Me
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
        End If
    End If
    
End Sub
Private Sub Form_Load()
    Dim bdP0044700  As ADODB.Connection 'Arquivo - P0044700.
    Dim lMensagem   As String           'Mensagem.
    Dim rsP0044700  As clsYasRecordSet  'Registro - P0044700.
Dim lSelect     As String
    '1. Centraliza form.
    Call gM46V999.gpCenForm(Me)

    '2. Carrega combo de seguradoras.
    If gM46V999.gfAbrBasDados("P0044700", gM46V999.gTipP0044700, gM46V999.gEndP0044700, bdP0044700, lMensagem) = False Then
        Exit Sub
    End If
    lSelect = "SELECT CSeguradora, Seguradora FROM YTAB_SEGURADORA"
    If gM46V999.gfObtRegistro(bdP0044700, lSelect, rsP0044700, lMensagem) = False Then
        Call gM46V999.gpFecha2(bdP0044700)
        Exit Sub
    End If
    While Not rsP0044700.EOF
        Select Case rsP0044700!CSeguradora
            Case 641
                'Yasuda - não incluir.
            Case 754, 872, 883
                'Seguradoras sem nome - inválidas - não incluir.
            Case Else
                cboConSeguradora(1).AddItem rsP0044700!CSeguradora & " - " & rsP0044700!Seguradora
                cboConSeguradora(1).ItemData(cboConSeguradora(1).NewIndex) = rsP0044700!CSeguradora
                cboConSeguradora(2).AddItem rsP0044700!CSeguradora & " - " & rsP0044700!Seguradora
                cboConSeguradora(2).ItemData(cboConSeguradora(2).NewIndex) = rsP0044700!CSeguradora
                cboConSeguradora(3).AddItem rsP0044700!CSeguradora & " - " & rsP0044700!Seguradora
                cboConSeguradora(3).ItemData(cboConSeguradora(3).NewIndex) = rsP0044700!CSeguradora
                cboConSeguradora(4).AddItem rsP0044700!CSeguradora & " - " & rsP0044700!Seguradora
                cboConSeguradora(4).ItemData(cboConSeguradora(4).NewIndex) = rsP0044700!CSeguradora
                cboConSeguradora(5).AddItem rsP0044700!CSeguradora & " - " & rsP0044700!Seguradora
                cboConSeguradora(5).ItemData(cboConSeguradora(5).NewIndex) = rsP0044700!CSeguradora
        End Select
        rsP0044700.MoveNext
    Wend
    Call gM46V999.gpFecha2(bdP0044700)
End Sub

Private Sub lblCodCosseguro_Click(Index As Integer)

End Sub

Private Sub txtCosPercentual_Change(Index As Integer)
    Call gpHabSalvar
End Sub

Private Sub txtCosPercentual_gotfocus(Index As Integer)
    With txtCosPercentual(Index)
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtCosPercentual_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCosPercentual_LostFocus(Index As Integer)
    If gM46V999.gfPreenchido(txtCosPercentual(Index).Text) = False Then txtCosPercentual(Index).Text = "0"
    txtCosPercentual(Index).Text = Format(txtCosPercentual(Index).Text, "standard")
End Sub

