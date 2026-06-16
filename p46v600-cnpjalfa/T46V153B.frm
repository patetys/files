VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmT46V153B 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro Técnico da Especificação da Proposta"
   ClientHeight    =   3030
   ClientLeft      =   2625
   ClientTop       =   2400
   ClientWidth     =   5835
   Icon            =   "T46V153B.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmDados 
      Height          =   2445
      Left            =   45
      TabIndex        =   4
      Top             =   495
      Width           =   5730
      Begin VB.TextBox TxtContratante 
         Height          =   285
         Left            =   135
         MaxLength       =   50
         TabIndex        =   2
         Top             =   945
         Width           =   5460
      End
      Begin VB.TextBox TxtPrazoObra 
         Height          =   285
         Left            =   135
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxtPrazoManut 
         Height          =   285
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxtLocalExposicao 
         Height          =   780
         Left            =   135
         MaxLength       =   80
         TabIndex        =   3
         Top             =   1575
         Width           =   5460
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prazo da Obra"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   135
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prazo Manutenção"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   1620
         TabIndex        =   7
         Top             =   135
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contratante"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   6
         Top             =   720
         Width           =   5460
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local de Exposição"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1350
         Width           =   5460
      End
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   4095
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":128C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":15A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":18C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":1D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":202C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":2346
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":2660
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":2AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":2DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":30E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":3538
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153B.frx":398A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar barFerramentas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salvar"
            Object.ToolTipText     =   "Gravar Itens da Especificação"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1000
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Fecha Especificação"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmT46V153B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub mpSalvar()
    If IsDate(TxtPrazoObra) Then
        frmT46V143B.CmdAbreTxt(e_Tipo_Espec_PrazoObra).Tag = TxtPrazoObra
    Else
        frmT46V143B.CmdAbreTxt(e_Tipo_Espec_PrazoObra).Tag = ""
    End If
    If IsDate(TxtPrazoManut) Then
        frmT46V143B.CmdAbreTxt(e_Tipo_Espec_PrazoManut).Tag = TxtPrazoManut
    Else
        frmT46V143B.CmdAbreTxt(e_Tipo_Espec_PrazoManut).Tag = ""
    End If
    frmT46V143B.CmdAbreTxt(e_Tipo_Espec_Contratante).Tag = Trim(TxtContratante)
    frmT46V143B.CmdAbreTxt(e_Tipo_Espec_LocalExpos).Tag = Trim(TxtLocalExposicao)
    Call gpLimPremios
End Sub

Private Sub barFerramentas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case "SALVAR"
        Call mpSalvar
        Unload Me
    Case "SAIR"
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
End Sub


Private Sub TxtContratante_GotFocus()
    With TxtContratante
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub


Private Sub TxtContratante_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtLocalExposicao_GotFocus()
    With TxtLocalExposicao
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub


Private Sub TxtLocalExposicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtPrazoManut_GotFocus()
    With TxtPrazoManut
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub


Private Sub TxtPrazoManut_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtPrazoManut_LostFocus()
    TxtPrazoManut.Text = gM46V999.gfForData(9, TxtPrazoManut.Text)
    If IsDate(TxtPrazoManut.Text) = False Then
        TxtPrazoManut.Text = ""
    End If
End Sub


Private Sub TxtPrazoObra_GotFocus()
    With TxtPrazoObra
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub


Private Sub TxtPrazoObra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtPrazoObra_LostFocus()
    TxtPrazoObra.Text = gM46V999.gfForData(9, TxtPrazoObra.Text)
    If IsDate(TxtPrazoObra.Text) = False Then
        TxtPrazoObra.Text = ""
    End If
End Sub


