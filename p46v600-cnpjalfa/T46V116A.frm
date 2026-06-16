VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V116A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cosseguradoras da proposta"
   ClientHeight    =   2550
   ClientLeft      =   2655
   ClientTop       =   2520
   ClientWidth     =   5400
   ControlBox      =   0   'False
   Icon            =   "T46V116A.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCosPercentual 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   5
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   14
      Text            =   "0,00"
      Top             =   1680
      Width           =   1000
   End
   Begin VB.TextBox txtCosPercentual 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   4
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   11
      Text            =   "0,00"
      Top             =   1380
      Width           =   1000
   End
   Begin VB.TextBox txtCosPercentual 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   3
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   8
      Text            =   "0,00"
      Top             =   1080
      Width           =   1000
   End
   Begin VB.TextBox txtCosPercentual 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   2
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   5
      Text            =   "0,00"
      Top             =   780
      Width           =   1000
   End
   Begin VB.TextBox txtCosPercentual 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   1
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "0,00"
      Top             =   480
      Width           =   1000
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   300
      Left            =   2160
      TabIndex        =   15
      ToolTipText     =   "Volta à proposta."
      Top             =   2160
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
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Indicador de corretor."
      Top             =   540
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
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Indicador de corretor."
      Top             =   840
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
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Indicador de corretor."
      Top             =   1140
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
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Indicador de corretor."
      Top             =   1440
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
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Indicador de corretor."
      Top             =   1740
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
   Begin VB.ComboBox cboConSeguradora 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   5
      ItemData        =   "T46V116A.frx":000C
      Left            =   360
      List            =   "T46V116A.frx":000E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      ToolTipText     =   "Seguradora para confirmação do bônus."
      Top             =   1680
      Width           =   3855
   End
   Begin VB.ComboBox cboConSeguradora 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   4
      ItemData        =   "T46V116A.frx":0010
      Left            =   360
      List            =   "T46V116A.frx":0012
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "Seguradora para confirmação do bônus."
      Top             =   1380
      Width           =   3855
   End
   Begin VB.ComboBox cboConSeguradora 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   3
      ItemData        =   "T46V116A.frx":0014
      Left            =   360
      List            =   "T46V116A.frx":0016
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      ToolTipText     =   "Seguradora para confirmação do bônus."
      Top             =   1080
      Width           =   3855
   End
   Begin VB.ComboBox cboConSeguradora 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   2
      ItemData        =   "T46V116A.frx":0018
      Left            =   360
      List            =   "T46V116A.frx":001A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Seguradora para confirmação do bônus."
      Top             =   780
      Width           =   3855
   End
   Begin VB.ComboBox cboConSeguradora 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   1
      ItemData        =   "T46V116A.frx":001C
      Left            =   360
      List            =   "T46V116A.frx":001E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Seguradora para confirmação do bônus."
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cosseguro"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   120
      Width           =   3765
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Percentual"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   4200
      TabIndex        =   16
      Top             =   105
      Width           =   1005
   End
End
Attribute VB_Name = "frmT46V116A"
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

Private Sub cboConSeguradora_Change(Index As Integer)
    Call gpLimPremios
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
    Dim lCos1 As Integer
    Dim lCos2 As Integer
    Dim lCos3 As Integer
    Dim lCos4 As Integer
    Dim lCos5 As Integer

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
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
Finalizar:
    If gM46V999.gDesabilitarForms Or gM46V999.gDesabFormsCotacao Then
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
    Dim ConexaoBD As ADODB.Connection       'Arquivo - P0044700.
    Dim lMensagem As String         'Mensagem.
    Dim rsP0044700 As clsYasRecordSet    'Registro - P0044700.
    Dim UltimoErro As String
    Dim lSelect As String
    '1. Centraliza form.
    Call gM46V999.gpCenForm(Me)

    '2. Carrega combo de seguradoras.
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, ConexaoBD, UltimoErro) = False Then
        Call gM46V999.gpGraLog(2, UltimoErro)
        Exit Sub
    End If
    
    lSelect = "SELECT CSeguradora, Seguradora FROM GED.DBO.TAB_SEGURADORA"
    If gM46V999.gfObtRegistro(ConexaoBD, lSelect, rsP0044700, lMensagem) = False Then
        Call gM46V999.gpFecha2(ConexaoBD)
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
    Call gM46V999.gpFecha2(ConexaoBD)
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

    KeyAscii = ConsisteKeyNum(KeyAscii, ",")
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCosPercentual_LostFocus(Index As Integer)
    If gM46V999.gfPreenchido(txtCosPercentual(Index).Text) = False Then txtCosPercentual(Index).Text = "0"
    txtCosPercentual(Index).Text = Format(gM46V999.CDblx(txtCosPercentual(Index).Text), "standard")
End Sub
