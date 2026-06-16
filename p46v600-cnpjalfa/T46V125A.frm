VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V125A 
   Caption         =   "Itens da proposta"
   ClientHeight    =   4395
   ClientLeft      =   750
   ClientTop       =   1335
   ClientWidth     =   3810
   ControlBox      =   0   'False
   Icon            =   "T46V125A.frx":0000
   KeyPreview      =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4395
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flxItens 
      Height          =   3800
      Left            =   100
      TabIndex        =   0
      ToolTipText     =   "Duplo clique sobre o item, aciona o mesmo."
      Top             =   100
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   6694
      _Version        =   393216
      Cols            =   5
      ForeColor       =   8388608
      ForeColorFixed  =   12582912
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   -2147483633
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Volta à janela de item."
      Top             =   4050
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V125A.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V125A.frx":006A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V125A.frx":00C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V125A.frx":0126
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V125A.frx":0184
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V125A.frx":01E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V125A.frx":0240
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V125A.frx":029E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmT46V125A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Me.Hide
    frmT46V101A.SetFocus
End Sub
Private Sub flxItens_DblClick()
    Call cmdOk_Click
End Sub
Private Sub flxItens_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Form_Activate()
'1. Posiciona nome do form.
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)

    '2. Se primeira vez, posiciona form.
    Me.Visible = True
    Me.Refresh

    '3. Coloca foco na grade.
    On Error Resume Next
    With flxItens
        .Col = 0
        .SetFocus
    End With
    On Error GoTo 0
End Sub
Private Sub Form_DblClick()
    Dim lAltura As Integer    'Altura original.
    Dim lLargura As Integer    'Largura original.
    Dim lMarEsquerda As Integer   'Margem esquerda.
    Dim lMarSuperior As Integer   'Margem superior.

'1. Calcula margens.
    lAltura = 4800
    lLargura = 3950
    With frmT46V101A
        lMarEsquerda = .Left + .Width - lLargura
        lMarSuperior = .Top + 1060
    End With

    '2. Posiciona form.
    Me.Move lMarEsquerda, lMarSuperior, lLargura, lAltura
End Sub
Private Sub Form_Load()
'1. Centraliza form.
'Height = 6000
'Call gM46V999.gpCenForm(Me)

'2. Prepara grade.
    Me.Visible = False
    With flxItens
        .Rows = 1
        .Cols = 4
        .Row = 0
        .FontWidth = 3.35
        'a) Posiciona largura das colunas.
        .ColWidth(0) = 405
        .ColWidth(1) = 2505
        .ColWidth(2) = 2175
        .ColWidth(3) = 945
        'b) Posiciona altura da linha de cabeçalho.
        .RowHeight(0) = 220
        'c) Posiciona alinhamento das colunas.
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        'd) Posiciona títulos.
        .Col = 0
        .Text = "Item"
        .Col = 1
        .Text = "Classificação"
        .Col = 2
        .Text = "Região"
        .Col = 3
        .Text = "CEP Risco"
    End With
    Call Form_DblClick
End Sub
Private Sub Form_Resize()
'1. Se não estiver na forma normal, encerra processo.
    If Me.WindowState <> vbNormal Then
        Exit Sub
    End If

    '2. Verifica tamanhos mínimos.
    If Me.Height < 1500 Then
        Me.Height = 1500
    End If
    If Me.Width < 3900 Then
        Me.Width = 3900
    End If

    '3. Posiciona grade.
    flxItens.Width = Me.ScaleWidth - 200
    flxItens.Height = Me.Height - 1000

    '4. Posiciona botão de comando.
    cmdOk.Top = flxItens.Top + flxItens.Height + 150
End Sub
