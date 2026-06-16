VERSION 5.00
Begin VB.Form frmT46V186A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edição da observação de franquias"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   420
      Left            =   3712
      TabIndex        =   1
      Top             =   1950
      Width           =   1410
   End
   Begin VB.TextBox txtClausula 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1755
      Left            =   45
      MaxLength       =   8000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   8745
   End
End
Attribute VB_Name = "frmT46V186A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Dim i As Integer
                 
    If txtClausula.Text <> Empty Then
        'Coloca asterisco
        If txtClausula.Text <> Empty And Left(txtClausula.Text, 1) <> "*" Then
            txtClausula.Text = "* " & Trim(txtClausula.Text)
        End If
        
        For i = 1 To frmT46V117A.flxSelecionadas.Rows - 1
            If frmT46V117A.flxSelecionadas.TextMatrix(i, ColCodCobert) = frmT46V179A.flxCobertura.TextMatrix(frmT46V179A.flxCobertura.Row, 0) Then
                If Left(frmT46V117A.flxSelecionadas.TextMatrix(i, ColTextoFranquia), 1) <> "*" Then
                    frmT46V117A.flxSelecionadas.TextMatrix(i, ColTextoFranquia) = "*" & frmT46V117A.flxSelecionadas.TextMatrix(i, ColTextoFranquia)
                    frmT46V179A.flxCobertura.TextMatrix(frmT46V179A.flxCobertura.Row, 2) = "*" & frmT46V179A.flxCobertura.TextMatrix(frmT46V179A.flxCobertura.Row, 2)
                End If
                Exit For
            End If
        Next
            
    Else
        'Retira asterisco
        For i = 1 To frmT46V117A.flxSelecionadas.Rows - 1
            If frmT46V117A.flxSelecionadas.TextMatrix(i, ColCodCobert) = frmT46V179A.flxCobertura.TextMatrix(frmT46V179A.flxCobertura.Row, 0) Then
                If Left(frmT46V117A.flxSelecionadas.TextMatrix(i, ColTextoFranquia), 1) = "*" Then
                    frmT46V117A.flxSelecionadas.TextMatrix(i, ColTextoFranquia) = Right(frmT46V117A.flxSelecionadas.TextMatrix(i, ColTextoFranquia), Len(frmT46V117A.flxSelecionadas.TextMatrix(i, ColTextoFranquia)) - 1)
                    frmT46V179A.flxCobertura.TextMatrix(frmT46V179A.flxCobertura.Row, 2) = Right(frmT46V179A.flxCobertura.TextMatrix(frmT46V179A.flxCobertura.Row, 2), Len(frmT46V179A.flxCobertura.TextMatrix(frmT46V179A.flxCobertura.Row, 2)) - 1)
                End If
                Exit For
            End If
        Next
    End If
 
    frmT46V179A.flxCobertura.TextMatrix(frmT46V179A.flxCobertura.Row, 4) = txtClausula.Text
    Unload Me
End Sub

Private Sub Form_Load()
    '1. Posiciona form.
    Call gM46V999.gpCenForm(Me)
    
    txtClausula.Text = frmT46V179A.flxCobertura.TextMatrix(frmT46V179A.flxCobertura.Row, 4)
End Sub
