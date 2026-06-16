VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V173A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informações LMI"
   ClientHeight    =   1605
   ClientLeft      =   2760
   ClientTop       =   2775
   ClientWidth     =   2805
   Icon            =   "T46V173A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   105
      TabIndex        =   3
      Top             =   1230
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdCancelar 
      Height          =   300
      Left            =   1710
      TabIndex        =   4
      Top             =   1230
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
   Begin VB.Frame FrmReferencia 
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   2625
      Begin VB.TextBox TxtReferCotac 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   90
         MaxLength       =   20
         TabIndex        =   1
         Top             =   630
         Width           =   2370
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nosso Número Cotação LMI"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   270
         Width           =   2370
      End
   End
   Begin VB.Frame FrmDesc 
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   2625
      Begin VB.TextBox TxtDescLMI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   345
         TabIndex        =   0
         Top             =   630
         Width           =   1875
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desconto LMI Único"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   345
         TabIndex        =   5
         Top             =   270
         Width           =   1875
      End
   End
End
Attribute VB_Name = "frmT46V173A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lPerDescLMI As Double
Dim lReferencia As String
Private Sub cmdCancelar_Click()
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    lPerDescLMI = gM46V999.CDblx(TxtDescLMI)
    lReferencia = TxtReferCotac

    If TxtReferCotac.Visible Then
        If Len(Trim(lReferencia)) <> 20 Then
            Call gM46V999.gpGraLog(1, "Nosso número inválido (deve conter 20 caracteres)")
            Exit Sub
        End If
    End If
    Me.Hide
End Sub

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
End Sub

Public Sub gpInstanciarForm(ByRef pPerDescLMI, ByVal pTipCotac As String, ByRef pReferencia)
    lPerDescLMI = pPerDescLMI
    TxtDescLMI = Format(gM46V999.CDblx(lPerDescLMI), "##0.00000")
    TxtReferCotac = pReferencia
    If gM46V999.gNovaPropostaLMI Then
        TxtDescLMI.Enabled = True
    Else
        TxtDescLMI.Enabled = False
    End If
    If pTipCotac = e_TipDoc_Cotacao Then
        FrmDesc.Visible = True
        FrmReferencia.Visible = False
    Else
        FrmDesc.Visible = False
        FrmReferencia.Visible = True
    End If
    Me.Show vbModal
    pReferencia = lReferencia
    pPerDescLMI = lPerDescLMI
End Sub

Private Sub TxtDescLMI_Change()
    Call gpHabSalvar
    gAltComum = True
End Sub

Private Sub TxtDescLMI_GotFocus()
    With TxtDescLMI
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub TxtDescLMI_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")
End Sub

Private Sub TxtDescLMI_LostFocus()
    TxtDescLMI.Text = Format(gM46V999.CDblx(TxtDescLMI), "##0.00000")
End Sub

