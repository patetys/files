VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT47V003A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de prazo"
   ClientHeight    =   1080
   ClientLeft      =   1860
   ClientTop       =   1545
   ClientWidth     =   3585
   Icon            =   "T47V003A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optTipPrazo 
      Caption         =   "Prazo curto"
      ForeColor       =   &H00800000&
      Height          =   200
      Index           =   0
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Prazo curto."
      Top             =   240
      Width           =   1250
   End
   Begin VB.OptionButton optTipPrazo 
      Caption         =   "Pro rata"
      ForeColor       =   &H00800000&
      Height          =   200
      Index           =   1
      Left            =   1365
      TabIndex        =   1
      ToolTipText     =   "Pro tata."
      Top             =   240
      Width           =   950
   End
   Begin VB.OptionButton optTipPrazo 
      Caption         =   "Prazo longo"
      ForeColor       =   &H00800000&
      Height          =   200
      Index           =   2
      Left            =   2325
      TabIndex        =   0
      ToolTipText     =   "Prazo longo."
      Top             =   240
      Width           =   1250
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   300
      Left            =   1215
      TabIndex        =   3
      ToolTipText     =   "Volta à proposta."
      Top             =   645
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT47V003A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
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
    Call gM46V999.gpCenForm(Me)
End Sub
Private Sub optTipPrazo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

