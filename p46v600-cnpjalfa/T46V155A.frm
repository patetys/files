VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V155A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Beneficiários"
   ClientHeight    =   3270
   ClientLeft      =   390
   ClientTop       =   1620
   ClientWidth     =   9930
   ControlBox      =   0   'False
   Icon            =   "T46V155A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTextoEspec 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2955
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9885
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   300
      Left            =   4560
      TabIndex        =   1
      ToolTipText     =   "Volta à proposta."
      Top             =   2970
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&OK"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V155A"
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
    If txtTextoEspec.Enabled And txtTextoEspec.Visible Then
        txtTextoEspec.SetFocus
    End If
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
End Sub





Private Sub txtTextoEspec_Change()
    Dim AuxArray As Variant
    Dim lNewTexto As String
    Dim RestTexto As String
    Dim lTextoOriginal As String
    Dim i As Integer
    Dim Pri As Boolean
    Dim lTotLinhas As Integer
    Dim PosCursor As Long
    Static Alterando As Boolean


    If Alterando Or gCarga Then Exit Sub

    If Trim(txtTextoEspec) = "" Then Exit Sub

    Alterando = True
    lNewTexto = ""
    RestTexto = ""
    lTotLinhas = 0
    Pri = True


    PosCursor = txtTextoEspec.SelStart
    txtTextoEspec.Text = gM46V999.gfRetirarCaracterEspecial(txtTextoEspec.Text, True, False)
    lTextoOriginal = txtTextoEspec.Text

    If InStr(lTextoOriginal, vbCrLf) <> 0 Then
        AuxArray = Split(lTextoOriginal, vbCrLf)
        For i = 0 To UBound(AuxArray)
            AuxArray(i) = RestTexto & AuxArray(i)
            RestTexto = ""
            If Len(AuxArray(i)) > 80 Then
                RestTexto = Mid(AuxArray(i), 81)
                lNewTexto = lNewTexto & Mid(AuxArray(i), 1, 80) & vbCrLf
                If (Mid(lTextoOriginal, PosCursor + 1, 1) = vbCr Or Len(lTextoOriginal) = PosCursor) And Pri Then
                    PosCursor = PosCursor + 2
                    Pri = False
                End If
                lTotLinhas = lTotLinhas + 1
            Else
                lNewTexto = lNewTexto & AuxArray(i) & vbCrLf
                lTotLinhas = lTotLinhas + 1
            End If
        Next
        If RestTexto <> "" Then
            txtTextoEspec = lNewTexto & RestTexto
        Else
            txtTextoEspec = Mid(lNewTexto, 1, Len(lNewTexto) - 2)
        End If
    Else
        If Len(lTextoOriginal) > 80 Then
            Do While True
                lNewTexto = lNewTexto & Mid(lTextoOriginal, 1, 80) & vbCrLf
                lTextoOriginal = Mid(lTextoOriginal, 81)
                If Len(lTextoOriginal) = 0 Then
                    Exit Do
                End If
            Loop
            txtTextoEspec = Mid(lNewTexto, 1, Len(lNewTexto) - 2)
            PosCursor = PosCursor + 2
        End If
    End If
    If lTotLinhas > 12 Then
        MsgBox "Total de Linhas é superior ao limite máximo (12)!" & vbCr & "O excedente será desprezado!", vbInformation
    End If
    txtTextoEspec.SelStart = PosCursor
    Call gpLimPremios
    Alterando = False
End Sub

