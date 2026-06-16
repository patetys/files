VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V128A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Segurança"
   ClientHeight    =   3315
   ClientLeft      =   2895
   ClientTop       =   2385
   ClientWidth     =   5385
   ControlBox      =   0   'False
   Icon            =   "T46V128A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkSeguranca 
      Caption         =   $"T46V128A.frx":000C
      ForeColor       =   &H00800000&
      Height          =   600
      Index           =   4
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Possui grades metálicas?"
      Top             =   2100
      Width           =   5160
   End
   Begin VB.CheckBox chkSeguranca 
      Caption         =   "Vigilância exclusiva e permanente do imóvel entendendo-se ainda como tal, porteiros fixos."
      ForeColor       =   &H00800000&
      Height          =   600
      Index           =   0
      Left            =   100
      TabIndex        =   3
      ToolTipText     =   "Possui vigilância?"
      Top             =   50
      Width           =   5160
   End
   Begin VB.CheckBox chkSeguranca 
      Caption         =   "Alarmes."
      ForeColor       =   &H00800000&
      Height          =   600
      Index           =   1
      Left            =   100
      TabIndex        =   2
      ToolTipText     =   "Possui alarmes?"
      Top             =   550
      Width           =   5160
   End
   Begin VB.CheckBox chkSeguranca 
      Caption         =   "Porteiros eletrônicos."
      ForeColor       =   &H00800000&
      Height          =   600
      Index           =   2
      Left            =   100
      TabIndex        =   1
      ToolTipText     =   "Possue porteiros eletrônicos."
      Top             =   1050
      Width           =   5160
   End
   Begin VB.CheckBox chkSeguranca 
      Caption         =   "Inexistência de terreno baldio,  imóvel desocupado   ou em construção na vizinhança do risco."
      ForeColor       =   &H00800000&
      Height          =   600
      Index           =   3
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Inexistência de terreno baldio?"
      Top             =   1550
      Width           =   5160
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   300
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "Volta ao item."
      Top             =   2880
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V128A"
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
Private Sub chkSeguranca_Click(index As Integer)
    Call gpHabSalvar
End Sub
Private Sub chkSeguranca_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub


