VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V126A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proteção"
   ClientHeight    =   1830
   ClientLeft      =   4080
   ClientTop       =   1965
   ClientWidth     =   2880
   ControlBox      =   0   'False
   Icon            =   "T46V126A.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkProtecao 
      Caption         =   "Detecção e alarme"
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   3
      Left            =   100
      TabIndex        =   4
      ToolTipText     =   "possui detecção e alarme?"
      Top             =   1000
      Width           =   2760
   End
   Begin VB.CheckBox chkProtecao 
      Caption         =   "Sprinklers (chuveiros automáticos)"
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   2
      Left            =   100
      TabIndex        =   3
      ToolTipText     =   "possui sprinklers?"
      Top             =   700
      Width           =   2760
   End
   Begin VB.CheckBox chkProtecao 
      Caption         =   "Hidrantes"
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   1
      Left            =   100
      TabIndex        =   2
      ToolTipText     =   "possui hidrantes?"
      Top             =   400
      Width           =   2760
   End
   Begin VB.CheckBox chkProtecao 
      Caption         =   "Extintor"
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   0
      Left            =   100
      TabIndex        =   1
      ToolTipText     =   "possui extintor?"
      Top             =   100
      Width           =   2760
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   300
      Left            =   960
      TabIndex        =   0
      ToolTipText     =   "Volta ao item."
      Top             =   1440
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V126A"
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
End Sub
Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
End Sub
Private Sub chkProtecao_Click(index As Integer)
    Call gpHabSalvar
End Sub
Private Sub chkProtecao_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

