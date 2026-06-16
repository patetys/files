VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V108A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagamento antecipado"
   ClientHeight    =   1320
   ClientLeft      =   5325
   ClientTop       =   3525
   ClientWidth     =   2115
   ControlBox      =   0   'False
   Icon            =   "T46V108A.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2115
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPagAntBloqueto 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   550
      MaxLength       =   9
      TabIndex        =   0
      ToolTipText     =   "Número de bloqueto do pagamento antecipado."
      Top             =   350
      Width           =   1000
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   550
      TabIndex        =   1
      ToolTipText     =   "Volta à proposta."
      Top             =   800
      Width           =   1000
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bloqueto"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   26
      Left            =   550
      TabIndex        =   2
      ToolTipText     =   "Número de bloqueto do pagamento antecipado."
      Top             =   100
      Width           =   1000
   End
End
Attribute VB_Name = "frmT46V108A"
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
    If gM46V999.gDesabilitarForms Then
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
    txtPagAntBloqueto.Text = ""
End Sub
Private Sub txtPagAntBloqueto_Change()
    Call gpHabSalvar
End Sub
Private Sub txtPagAntBloqueto_GotFocus()
    With txtPagAntBloqueto
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtPagAntBloqueto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtPagAntBloqueto_LostFocus()
    Dim lAuxiliar As String     'Auxiliar.

    With txtPagAntBloqueto
        If gM46V999.gfPreenchido(.Text) Then
            .Text = Trim$(.Text)
            If Len(.Text) = 9 Then
                Exit Sub
            End If
            If IsNumeric(.Text) Then
                lAuxiliar = Format$(CDbl(.Text), "00000000")
                .Text = Mid$(lAuxiliar, 1, 7) & "-" & Mid$(lAuxiliar, 8, 1)
            End If
        End If
    End With
End Sub
