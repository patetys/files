VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT47V009A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carnê"
   ClientHeight    =   1125
   ClientLeft      =   1860
   ClientTop       =   1545
   ClientWidth     =   4980
   Icon            =   "T47V009A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboCarBanco 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2700
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Banco para emissão de carnê."
      Top             =   120
      Width           =   2200
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   1950
      TabIndex        =   1
      ToolTipText     =   "Volta à proposta."
      Top             =   615
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
      Caption         =   "Banco: "
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2700
   End
End
Attribute VB_Name = "frmT47V009A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboCarBanco_Click()
    Call gpHabSalvar
End Sub
Private Sub cboCarBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub cboCarBanco_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then     'Teclou backspace.
        cboCarBanco.ListIndex = -1
    End If
End Sub
Private Sub cmdOK_Click()
    Me.Hide
End Sub
Private Sub Form_Activate()
    
    If cboCarBanco.Visible = True And cboCarBanco.Enabled = True Then
        cboCarBanco.SetFocus
    End If
End Sub
Private Sub Form_Load()
    Dim lMensagem   As String   'Mensagem.

    '1. Centraliza form.
    Call gM46V999.gpCenForm(Me)

    '2. Carrega combo de bancos.
    If gfCarCboBancos(2, cboCarBanco, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    Select Case gM46V999.PerfilAcesso
        Case 1, 2, 5
            'cboCarBanco.AddItem "215"
            'cboCarBanco.ItemData(cboCarBanco.NewIndex) = 215
    End Select
End Sub

