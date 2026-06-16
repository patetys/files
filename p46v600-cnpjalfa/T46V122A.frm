VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V122A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Observação"
   ClientHeight    =   3570
   ClientLeft      =   1740
   ClientTop       =   1755
   ClientWidth     =   6480
   ControlBox      =   0   'False
   Icon            =   "T46V122A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtObservacao 
      ForeColor       =   &H00800000&
      Height          =   3135
      Left            =   50
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   50
      Width           =   6375
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Volta à proposta."
      Top             =   3240
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&OK"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V122A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
Finalizar:
    If gM46V999.gDesabilitarForms Then
        If Me.Tag <> "Desabilitado" Then
            '1001266 - Validar CPF e CNPJ na Cotação e Proposta
            '''            DesabilitaControles Me
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
        End If
    End If

    If gM46V999.gCod_Ramo = e_CodRamo_Roubos Or gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        txtObservacao.MaxLength = 2000
    Else
        txtObservacao.MaxLength = 600
    End If

    '1001266 - Validar CPF e CNPJ na Cotação e Proposta
    '''    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
    '''        If gM46V999.gDesabFormsTransmi = True Then
    '''            DesabilitaControles Me
    '''        End If
    '''    End If

    cmdOk.Enabled = True

End Sub

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
End Sub


Private Sub txtObservacao_Change()
    Call gpHabSalvar
End Sub
