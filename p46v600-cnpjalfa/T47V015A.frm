VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT47V015A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custo de emissão"
   ClientHeight    =   1395
   ClientLeft      =   1860
   ClientTop       =   1545
   ClientWidth     =   4185
   Icon            =   "T47V015A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCusEmissao 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2415
      MaxLength       =   10
      TabIndex        =   3
      ToolTipText     =   "Valor do custo de emissão."
      Top             =   120
      Width           =   1485
   End
   Begin Threed.SSPanel panX 
      Height          =   285
      Left            =   2415
      TabIndex        =   0
      ToolTipText     =   "Custo de emissão parcelado?"
      Top             =   525
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   503
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.OptionButton optCusEmissao 
         Caption         =   "&Não"
         ForeColor       =   &H00800000&
         Height          =   200
         Index           =   1
         Left            =   800
         TabIndex        =   2
         ToolTipText     =   "Custo de emissão NÃO parcelado."
         Top             =   50
         Width           =   600
      End
      Begin VB.OptionButton optCusEmissao 
         Caption         =   "&Sim"
         ForeColor       =   &H00800000&
         Height          =   200
         Index           =   0
         Left            =   100
         TabIndex        =   1
         ToolTipText     =   "SIM, custo de emissão parcelado."
         Top             =   50
         Value           =   -1  'True
         Width           =   600
      End
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   300
      Left            =   1425
      TabIndex        =   4
      ToolTipText     =   "Volta à proposta."
      Top             =   915
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
      Caption         =   "Custo de emissão: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Custo de emissão."
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Custo de emissão parcelado? "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Custo de emissão parcelado?"
      Top             =   525
      Width           =   2295
   End
End
Attribute VB_Name = "frmT47V015A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
    If frmT47V001A.lVerTipEmissao = "E" Then
        gM46V999.gProCusEndosso = frmT47V015A.txtCusEmissao.Text
    Else
        gM46V999.gProCusApolice = frmT47V015A.txtCusEmissao.Text
    End If
    Me.Hide
End Sub
Private Sub Form_Activate()
Finalizar:
    If gM46V999.gDesabilitarForms Or gM46V999.gDesabFormsCotacao Then
        If gM46V999.gTipoAcesso = DIGITACAO And gM46V999.gDesabFormsCotacao = True Then GoTo Habilita
        If Me.Tag <> "Desabilitado" Then
            DesabilitaControles Me
        End If
    Else
Habilita:
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
        End If
    End If
    
End Sub
Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        txtCusEmissao.Enabled = False
    End If
End Sub
Private Sub optCusEmissao_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtCusEmissao_GotFocus()
    With txtCusEmissao
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtCusEmissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtCusEmissao_LostFocus()
    With txtCusEmissao
        If gM46V999.gfPreenchido(.Text) Then
            If IsNumeric(.Text) Then
                .Text = Format$(CDbl(.Text), "standard")
            End If
        End If
    End With
End Sub

