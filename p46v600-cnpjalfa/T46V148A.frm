VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V148A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Texto para Endosso"
   ClientHeight    =   5310
   ClientLeft      =   1215
   ClientTop       =   1770
   ClientWidth     =   10110
   ControlBox      =   0   'False
   Icon            =   "T46V148A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCodDizer 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7740
      TabIndex        =   2
      Top             =   4815
      Width           =   645
   End
   Begin Threed.SSCommand cmdFechar 
      Height          =   390
      Left            =   4170
      TabIndex        =   1
      ToolTipText     =   "Volta à proposta."
      Top             =   4905
      Width           =   1770
      _Version        =   65536
      _ExtentX        =   3122
      _ExtentY        =   688
      _StockProps     =   78
      Caption         =   "&Fechar"
      ForeColor       =   0
   End
   Begin Threed.SSCommand cmdBusca 
      Height          =   375
      Left            =   8415
      TabIndex        =   3
      ToolTipText     =   "Volta à proposta."
      Top             =   4815
      Width           =   1635
      _Version        =   65536
      _ExtentX        =   2884
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Busca Texto Padrão"
      ForeColor       =   8388608
   End
   Begin VB.TextBox txtTextoEndosso 
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
      Height          =   4800
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   10110
   End
End
Attribute VB_Name = "frmT46V148A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBusca_Click()
    Dim lbdTab_Ped_Dizer_Pad As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lrsTab_Ped_Dizer_Pad As clsYasRecordSet    'Registro : Tab_Ped_Dizer_Pad.
    Dim lMensagem As String        'Mensagem.
    Dim lLinha As Long       'Linhas.
    Dim lSelect As String
    If Not txtTextoEndosso.Enabled Then Exit Sub

    txtCodDizer_LostFocus
    If txtCodDizer = "0" Then Exit Sub

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Ped_Dizer_Pad, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If


    lSelect = " Select * from Tab_Ped_Dizer_Pad  "
    lSelect = lSelect & " Where cod_Dizer_Pad = " & txtCodDizer
    lSelect = lSelect & " order by num_seq"

    If gM46V999.gfObtRegistro(lbdTab_Ped_Dizer_Pad, lSelect, lrsTab_Ped_Dizer_Pad, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTab_Ped_Dizer_Pad)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    gCarga = True
    txtTextoEndosso = ""
    If lrsTab_Ped_Dizer_Pad.EOF Then
        MsgBox "Mensagem Padrão '" & txtCodDizer & "' não encontrada!", vbInformation
    Else
        Do While Not lrsTab_Ped_Dizer_Pad.EOF
            txtTextoEndosso = txtTextoEndosso & lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo") & vbCrLf
            lrsTab_Ped_Dizer_Pad.MoveNext
        Loop
        txtTextoEndosso = Mid(txtTextoEndosso, 1, Len(txtTextoEndosso) - 2)
    End If
    Call gM46V999.gpFecha2(lbdTab_Ped_Dizer_Pad)
    gCarga = False
    txtTextoEndosso_Change

End Sub


Private Sub cmdOk_Click()
End Sub

Private Sub cmdFechar_Click()
    Me.Hide
End Sub


Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    If txtTextoEndosso.Enabled And txtTextoEndosso.Visible Then
        txtTextoEndosso.SetFocus
    End If
Finalizar:
    If gM46V999.gDesabilitarForms Or gM46V999.gDesabFormsCotacao Then
        If Me.Tag <> "Desabilitado" Then
            DesabilitaControles Me
            cmdFechar.Enabled = True
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





Private Sub txtCodDizer_LostFocus()
    txtCodDizer = gM46V999.CDblx(txtCodDizer)
End Sub


Private Sub txtTextoEndosso_Change()
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

    If Trim(txtTextoEndosso) = "" Then Exit Sub

    Alterando = True
    lNewTexto = ""
    RestTexto = ""
    lTotLinhas = 0
    Pri = True


    PosCursor = txtTextoEndosso.SelStart
    txtTextoEndosso.Text = gM46V999.gfRetirarCaracterEspecial(txtTextoEndosso.Text, True, False)
    lTextoOriginal = txtTextoEndosso.Text

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
            txtTextoEndosso = lNewTexto & RestTexto
        Else
            txtTextoEndosso = Mid(lNewTexto, 1, Len(lNewTexto) - 2)
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
            txtTextoEndosso = Mid(lNewTexto, 1, Len(lNewTexto) - 2)
            PosCursor = PosCursor + 2
        End If
    End If
    If lTotLinhas > 60 Then
        MsgBox "Total de Linhas é superior ao limite máximo (60)!" & vbCr & "O excedente será desprezado!", vbInformation
    End If
    txtTextoEndosso.SelStart = PosCursor
    Call gpLimPremios
    Alterando = False
End Sub

