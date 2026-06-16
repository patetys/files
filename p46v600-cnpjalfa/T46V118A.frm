VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V118A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cláusula de beneficiário"
   ClientHeight    =   1335
   ClientLeft      =   2040
   ClientTop       =   1815
   ClientWidth     =   6330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "T46V118A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBenCPF 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1425
      MaxLength       =   18
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "CPF/CNPJ do beneficiário."
      Top             =   525
      Width           =   1800
   End
   Begin VB.TextBox txtBemNome 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1425
      MaxLength       =   60
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Nome do beneficiário."
      Top             =   120
      Width           =   4755
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   300
      Left            =   2640
      TabIndex        =   0
      ToolTipText     =   "Volta à proposta."
      Top             =   960
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&OK"
      ForeColor       =   8388608
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CPF/CNPJ:"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Número da conta corrente para débito."
      Top             =   525
      Width           =   1305
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nome:"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Agência da conta de débito."
      Top             =   120
      Width           =   1305
   End
End
Attribute VB_Name = "frmT46V118A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
'1. Posiciona form.
    Call gM46V999.gpCenForm(Me)
End Sub

Private Sub txtBemNome_Change()
    Call gpLimPremios
End Sub

Private Sub txtBenCPF_Change()
    Call gpLimPremios
End Sub

Private Sub txtBenCPF_LostFocus()
    Dim lretorno As String    'Retorno da função que formata o CPF ou CNPJ.

    txtBenCPF.Text = Replace(Replace(Replace(txtBenCPF.Text, ".", ""), "-", ""), "\", "")
    If Len(txtBenCPF.Text) <= 11 Then
        Call gM46V999.gpForCPF(1, txtBenCPF.Text, lretorno)
    Else
        Call gM46V999.gpForCNPJ(1, txtBenCPF.Text, lretorno)
    End If
    txtBenCPF.Text = lretorno
End Sub
