VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V194A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Observações do Local"
   ClientHeight    =   3570
   ClientLeft      =   1740
   ClientTop       =   1755
   ClientWidth     =   6480
   ControlBox      =   0   'False
   Icon            =   "T46V194A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtObservacao 
      ForeColor       =   &H00800000&
      Height          =   3135
      Left            =   50
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   50
      Width           =   6375
   End
   Begin Threed.SSCommand cmdSalvar 
      Height          =   300
      Left            =   5400
      TabIndex        =   1
      ToolTipText     =   "Volta à proposta."
      Top             =   3240
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Salvar"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdCancelar 
      Height          =   300
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Volta à proposta."
      Top             =   3240
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V194A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
    txtObservacao.Text = ""
    Me.Hide
End Sub

Private Sub cmdSalvar_Click()
    Call gpHabSalvar
    Me.Hide
End Sub

Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
Finalizar:

    If Me.Tag = "Desabilitado" Then
        HabilitaControles Me
    End If

    cmdSalvar.Enabled = True

End Sub

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
End Sub
