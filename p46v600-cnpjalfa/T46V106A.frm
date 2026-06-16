VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V106A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de prazo"
   ClientHeight    =   870
   ClientLeft      =   3765
   ClientTop       =   2130
   ClientWidth     =   3510
   ControlBox      =   0   'False
   Icon            =   "T46V106A.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optTipPrazo 
      Caption         =   "Prazo longo"
      ForeColor       =   &H00800000&
      Height          =   200
      Index           =   2
      Left            =   2250
      TabIndex        =   2
      ToolTipText     =   "Prazo longo."
      Top             =   100
      Width           =   1250
   End
   Begin VB.OptionButton optTipPrazo 
      Caption         =   "Pro rata"
      ForeColor       =   &H00800000&
      Height          =   200
      Index           =   1
      Left            =   1300
      TabIndex        =   1
      ToolTipText     =   "Pro tata."
      Top             =   100
      Width           =   950
   End
   Begin VB.OptionButton optTipPrazo 
      Caption         =   "Prazo curto"
      ForeColor       =   &H00800000&
      Height          =   200
      Index           =   0
      Left            =   50
      TabIndex        =   0
      ToolTipText     =   "Prazo curto."
      Top             =   100
      Width           =   1250
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   300
      Left            =   1150
      TabIndex        =   3
      ToolTipText     =   "Volta à proposta."
      Top             =   500
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V106A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOk_Click()
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
    Call gM46V999.gpCenForm(Me)
End Sub
Private Sub optTipPrazo_Click(Index As Integer)
    Call gpLimPremios
    If gCarga = False Then
        gAltComum = True
    End If
End Sub
Private Sub optTipPrazo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
