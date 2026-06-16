VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT47V007A 
   Caption         =   "Busca de endereço"
   ClientHeight    =   5295
   ClientLeft      =   1860
   ClientTop       =   1545
   ClientWidth     =   8925
   Icon            =   "T47V007A.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   8925
   Begin MSFlexGridLib.MSFlexGrid flxCEP 
      Height          =   2870
      Left            =   100
      TabIndex        =   0
      Top             =   2400
      Width           =   8800
      _ExtentX        =   15505
      _ExtentY        =   5054
      _Version        =   393216
      Cols            =   7
      ForeColor       =   12582912
      ForeColorFixed  =   12582912
      ForeColorSel    =   12582912
      WordWrap        =   -1  'True
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin Threed.SSPanel panX 
      Height          =   495
      Index           =   2
      Left            =   4500
      TabIndex        =   1
      Top             =   495
      Width           =   4410
      _Version        =   65536
      _ExtentX        =   7779
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.ComboBox cboUF 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T47V007A.frx":000C
         Left            =   2000
         List            =   "T47V007A.frx":0061
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Unidade da federação do endereço do segurado."
         Top             =   80
         Width           =   960
      End
      Begin Threed.SSCheck chkUF 
         Height          =   200
         Left            =   100
         TabIndex        =   3
         ToolTipText     =   "Considerar unidade da federação na pesquisa?"
         Top             =   120
         Width           =   1900
         _Version        =   65536
         _ExtentX        =   3360
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "&Unidade da federação"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
   Begin Threed.SSPanel panX 
      Height          =   495
      Index           =   1
      Left            =   105
      TabIndex        =   4
      Top             =   495
      Width           =   4950
      _Version        =   65536
      _ExtentX        =   8731
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.TextBox txtCEP 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   5
         ToolTipText     =   "CEP considerado na pesquisa."
         Top             =   90
         Width           =   1000
      End
      Begin Threed.SSCheck chkCEP 
         Height          =   200
         Left            =   100
         TabIndex        =   6
         ToolTipText     =   "Considerar CEP na pesquisa?"
         Top             =   120
         Width           =   1100
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "C&EP"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
   Begin Threed.SSPanel panX 
      Height          =   495
      Index           =   3
      Left            =   105
      TabIndex        =   7
      Top             =   945
      Width           =   8805
      _Version        =   65536
      _ExtentX        =   15531
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.TextBox txtNomLogradouro 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   8
         ToolTipText     =   "Nome ou parte do nome do logradouro considerado na pesquisa."
         Top             =   90
         Width           =   3100
      End
      Begin Threed.SSPanel panLogradouro 
         Height          =   300
         Index           =   1
         Left            =   6300
         TabIndex        =   9
         ToolTipText     =   "Opção de pesquisa para nome do logradouro."
         Top             =   105
         Width           =   2400
         _Version        =   65536
         _ExtentX        =   4233
         _ExtentY        =   529
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optLogradouro2 
            Caption         =   "&Qualquer parte"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   800
            TabIndex        =   11
            ToolTipText     =   "Considerar qualquer parte do nome do logradouro na pesquisa."
            Top             =   50
            Width           =   1500
         End
         Begin VB.OptionButton optLogradouro2 
            Caption         =   "&Início"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   100
            TabIndex        =   10
            ToolTipText     =   "Considerar início do nome do logradouro na pesquisa."
            Top             =   50
            Width           =   700
         End
      End
      Begin Threed.SSPanel panLogradouro 
         Height          =   300
         Index           =   0
         Left            =   4395
         TabIndex        =   12
         ToolTipText     =   "Opção de pesquisa para nome do logradouro."
         Top             =   105
         Width           =   1905
         _Version        =   65536
         _ExtentX        =   3360
         _ExtentY        =   529
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optLogradouro1 
            Caption         =   "&Completo"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   100
            TabIndex        =   14
            ToolTipText     =   "Considerar nome completo do logradouro na pesquisa."
            Top             =   50
            Width           =   1000
         End
         Begin VB.OptionButton optLogradouro1 
            Caption         =   "&Parte"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   1100
            TabIndex        =   13
            ToolTipText     =   "Considerar parte do nome do logradouro na pesquisa."
            Top             =   50
            Width           =   700
         End
      End
      Begin Threed.SSCheck chkLogradouro 
         Height          =   200
         Left            =   100
         TabIndex        =   15
         ToolTipText     =   "Considerar logradouro na pesquisa?"
         Top             =   120
         Width           =   1100
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "&Logradouro"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
   Begin Threed.SSPanel panX 
      Height          =   495
      Index           =   4
      Left            =   105
      TabIndex        =   16
      Top             =   1395
      Width           =   8805
      _Version        =   65536
      _ExtentX        =   15531
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.TextBox txtBairro 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   17
         ToolTipText     =   "Nome ou parte do nome do bairro considerado na pesquisa."
         Top             =   90
         Width           =   2400
      End
      Begin Threed.SSPanel panBairro 
         Height          =   300
         Index           =   0
         Left            =   4395
         TabIndex        =   18
         ToolTipText     =   "Opção de pesquisa para nome do bairro."
         Top             =   105
         Width           =   1905
         _Version        =   65536
         _ExtentX        =   3360
         _ExtentY        =   529
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optBairro1 
            Caption         =   "&Completo"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   100
            TabIndex        =   20
            ToolTipText     =   "Considerar nome completo do bairro na pesquisa."
            Top             =   50
            Width           =   1000
         End
         Begin VB.OptionButton optBairro1 
            Caption         =   "&Parte"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   1100
            TabIndex        =   19
            ToolTipText     =   "Considerar parte do nome do bairro na pesquisa."
            Top             =   50
            Width           =   700
         End
      End
      Begin Threed.SSPanel panBairro 
         Height          =   300
         Index           =   1
         Left            =   6300
         TabIndex        =   21
         ToolTipText     =   "Opção de pesquisa para nome do bairro."
         Top             =   105
         Width           =   2400
         _Version        =   65536
         _ExtentX        =   4233
         _ExtentY        =   529
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optBairro2 
            Caption         =   "&Qualquer parte"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   800
            TabIndex        =   23
            ToolTipText     =   "Considerar qualquer parte do nome do bairro na pesquisa."
            Top             =   50
            Width           =   1500
         End
         Begin VB.OptionButton optBairro2 
            Caption         =   "&Início"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   100
            TabIndex        =   22
            ToolTipText     =   "Considerar início do nome do bairro na pesquisa."
            Top             =   50
            Width           =   700
         End
      End
      Begin Threed.SSCheck chkBairro 
         Height          =   200
         Left            =   100
         TabIndex        =   24
         ToolTipText     =   "Considerar bairro na pesquisa?"
         Top             =   120
         Width           =   1100
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "&Bairro"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
   Begin Threed.SSPanel panX 
      Height          =   495
      Index           =   5
      Left            =   105
      TabIndex        =   25
      ToolTipText     =   "Considerar cidade na pesquisa?"
      Top             =   1845
      Width           =   8805
      _Version        =   65536
      _ExtentX        =   15531
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.TextBox txtCidade 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   26
         ToolTipText     =   "Nome ou parte do nome da cidade considerada na pesquisa."
         Top             =   90
         Width           =   2400
      End
      Begin Threed.SSPanel panCidade 
         Height          =   300
         Index           =   0
         Left            =   4395
         TabIndex        =   27
         ToolTipText     =   "Opção de pesquisa para nome da cidade."
         Top             =   105
         Width           =   1905
         _Version        =   65536
         _ExtentX        =   3360
         _ExtentY        =   529
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optCidade1 
            Caption         =   "&Completo"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   100
            TabIndex        =   29
            ToolTipText     =   "Considerar nome completo da cidade na pesquisa."
            Top             =   50
            Width           =   1000
         End
         Begin VB.OptionButton optCidade1 
            Caption         =   "&Parte"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   1100
            TabIndex        =   28
            ToolTipText     =   "Considerar parte do nome da cidade na pesquisa."
            Top             =   50
            Width           =   700
         End
      End
      Begin Threed.SSPanel panCidade 
         Height          =   300
         Index           =   1
         Left            =   6300
         TabIndex        =   30
         ToolTipText     =   "Opção de pesquisa para nome da cidade."
         Top             =   105
         Width           =   2400
         _Version        =   65536
         _ExtentX        =   4233
         _ExtentY        =   529
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optCidade2 
            Caption         =   "&Qualquer parte"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   800
            TabIndex        =   32
            ToolTipText     =   "Considerar qualquer parte do nome da cidade na pesquisa."
            Top             =   50
            Width           =   1500
         End
         Begin VB.OptionButton optCidade2 
            Caption         =   "&Início"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   100
            TabIndex        =   31
            ToolTipText     =   "Considerar início do nome da cidade na pesquisa."
            Top             =   50
            Width           =   700
         End
      End
      Begin Threed.SSCheck chkCidade 
         Height          =   200
         Left            =   100
         TabIndex        =   33
         Top             =   120
         Width           =   1100
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Ci&dade"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   8450
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V007A.frx":00D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V007A.frx":03EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V007A.frx":083D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V007A.frx":0B57
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V007A.frx":0E71
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar barFerramentas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pesquisar"
            Object.ToolTipText     =   "Pesquisar endereços - F3."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Selecionar"
            Object.ToolTipText     =   "Selecionar endereço - F6."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Parar"
            Object.ToolTipText     =   "Interromper processo."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair do programa."
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmT47V007A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mInterromper    As Boolean      'Indica que usuário clicou botão interromper.
Dim mCEP                As String       'CEP.
Dim mTipLogradouro      As String       'Tipo de logradouro.
Dim mLogradouro         As String       'Logradouro.
Dim mBairro             As String       'Bairro.
Dim mCidade             As String       'Cidade.
Dim mUF                 As String       'UF.

Public Sub mpInstanciarForm(ByRef pCEP As String, ByRef pTipLogradouro As String, ByRef pLogradouro As String, _
                            ByRef pBairro As String, ByRef pCidade As String, ByRef pUF As String)
    Me.Show vbModal, frmT47V004A
    
    pCEP = mCEP
    pTipLogradouro = mTipLogradouro
    pLogradouro = mLogradouro
    pBairro = mBairro
    pCidade = mCidade
    pUF = mUF
    
End Sub
Private Sub barFerramentas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
        Case "PESQUISAR"
            Call mnuArqPesCEP_Click
        Case "SELECIONAR"
            Call mnuArqSelCEP_Click
        Case "PARAR"
            mInterromper = True
        Case "SAIR"
            Unload Me
    End Select
End Sub
Private Sub cboUF_Click()
    If flxCEP.Visible = True Then
        flxCEP.Visible = False
    
        barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
        barFerramentas.Buttons(4).Enabled = False   'Selecionar.
    End If
End Sub
Private Sub cboUF_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub chkBairro_Click(Value As Integer)
    If flxCEP.Visible = True Then
        flxCEP.Visible = False
        barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
        barFerramentas.Buttons(4).Enabled = False   'Selecionar.
    End If
    With txtBairro
        .Visible = Value
        On Error Resume Next
        If .Visible = True And .Enabled = True Then
            .SetFocus
        End If
        On Error GoTo 0
    End With
    panBairro(0).Visible = Value
    panBairro(1).Visible = Value
End Sub
Private Sub chkBairro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub chkCEP_Click(Value As Integer)
    If flxCEP.Visible = True Then
        flxCEP.Visible = False
        barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
        barFerramentas.Buttons(4).Enabled = False   'Selecionar.
    End If
    With txtCEP
        .Visible = Value
        On Error Resume Next
        If .Visible = True And .Enabled = True Then
            .SetFocus
        End If
        On Error GoTo 0
    End With
End Sub
Private Sub chkCEP_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub chkCidade_Click(Value As Integer)
    If flxCEP.Visible = True Then
        flxCEP.Visible = False
        barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
        barFerramentas.Buttons(4).Enabled = False   'Selecionar.
    End If
    With txtCidade
        .Visible = Value
        On Error Resume Next
        If .Visible = True And .Enabled = True Then
            .SetFocus
        End If
        On Error GoTo 0
    End With
    panCidade(0).Visible = Value
    panCidade(1).Visible = Value
End Sub
Private Sub chkCidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub chkLogradouro_Click(Value As Integer)
    If flxCEP.Visible = True Then
        flxCEP.Visible = False
        barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
        barFerramentas.Buttons(4).Enabled = False   'Selecionar.
    End If
    With txtNomLogradouro
        .Visible = Value
        On Error Resume Next
        .SetFocus
        On Error GoTo 0
    End With
    panLogradouro(0).Visible = Value
    panLogradouro(1).Visible = Value
End Sub
Private Sub chkLogradouro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub chkUF_Click(Value As Integer)
    If flxCEP.Visible = True Then
        flxCEP.Visible = False
        
        barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
        barFerramentas.Buttons(4).Enabled = False   'Selecionar.
    End If
    With cboUF
        .Visible = Value
        On Error Resume Next
        .SetFocus
        On Error GoTo 0
    End With
End Sub
Private Sub chkUF_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub cmdSair_Click()
    'Observação: controle criado para usuário sair do FORM utilizando tecla ESC. Invisível pelo usuário.
    Unload Me
End Sub
Private Sub flxCEP_DblClick()
    Call mnuArqSelCEP_Click
End Sub
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass

    '1. Centraliza form.
    Height = 6000
    Call gM46V999.gpCenForm(Me)

    '3. Prepara grade.
    With flxCEP
        .Row = 0
        'a) Posiciona largura das colunas.
        .ColWidth(0) = 500
        .ColWidth(1) = 1000
        .ColWidth(2) = 2700
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1000
        .ColWidth(6) = 900
        'b) Posiciona altura da linha de cabeçalho.
        .RowHeight(0) = 480
        'c) Posiciona alinhamento das colunas.
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        'd) Posiciona títulos.
        .Col = 1
        .Text = "Tipo do logradouro"
        .Col = 2
        .Text = "Logradouro"
        .Col = 3
        .Text = "Bairro"
        .Col = 4
        .Text = "Cidade"
        .Col = 5
        .Text = "Unidade da federação"
        .Col = 6
        .Text = "CEP"
    End With

    '4. Posiciona opções iniciais.
    chkCEP.Value = 0
    chkUF.Value = 0
    chkLogradouro.Value = 0
    chkBairro.Value = 0
    chkCidade.Value = 0
    optLogradouro1(0) = True
    optBairro1(0) = True
    optCidade1(0) = True

    '5. Desabilita opção de seleção.
    
    barFerramentas.Buttons(4).Enabled = False   'Selecionar.

    Screen.MousePointer = vbDefault
End Sub
Private Sub mnuArqPesCEP_Click()
    'Grade.
    '   Coluna 1 = tipo do logradouro.
    '   Coluna 2 = nome do logradouro.
    '   Coluna 3 = nome do bairro.
    '   Coluna 4 = nome da cidade.
    '   Coluna 5 = sigla da unidade da federação.
    '   Coluna 6 = CEP.

    Dim lbdP0043000 As ADODB.Connection 'Arquivo - P0043000.
    Dim lMensagem   As String           'Mensagem.
    Dim lrsP0043000 As clsYasRecordSet  'Registro - P0043000.
    Dim lSelect     As String
    Screen.MousePointer = vbHourglass

    '1. Consiste opções de entrada.
    If chkCEP.Value = False And chkUF.Value = False And chkLogradouro.Value = False And _
       chkBairro.Value = False And chkCidade.Value = False Then
        Call gM46V999.gpGraLog(1, "FNP0045 - " & gM46V999.gPrefixo & "Opção de busca de endereço não selecionada.")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If chkCEP.Value = True Then
        If Not gM46V999.gfPreenchido(txtCEP.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0067 - " & gM46V999.gPrefixo & _
                             "CEP inválido (busca de endereços - não preenchido).")
            If txtCEP.Visible = True And txtCEP.Enabled = True Then
                txtCEP.SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If Len(Trim(txtCEP.Text)) <> 9 Then
            Call gM46V999.gpGraLog(1, "FGE0067 - " & gM46V999.gPrefixo & "CEP inválido (busca de endereços - " & _
                             txtCEP.Text & " - não tem 9 posições - fora do formato nnnnn-nnn).")
            If txtCEP.Visible = True And txtCEP.Enabled = True Then
                txtCEP.SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If Not IsNumeric(Left$(txtCEP.Text, 5)) And Mid$(txtCEP.Text, 6, 1) <> "-" And _
           Not IsNumeric(Right$(txtCEP.Text, 3)) Then
            Call gM46V999.gpGraLog(1, "FGE0067 - " & gM46V999.gPrefixo & "CEP inválido (busca de endereços - " & _
                             txtCEP.Text & " - não numérico - fora do formato nnnnn-nnn).")
            If txtCEP.Visible = True And txtCEP.Enabled = True Then
                txtCEP.SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    If chkUF.Value = True Then
        If cboUF.ListIndex = -1 Then
            Call gM46V999.gpGraLog(1, "FNP0043 - " & gM46V999.gPrefixo & "Unidade da federação não preenchida.")
            If cboUF.Visible = True And cboUF.Enabled = True Then
                cboUF.SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    If chkLogradouro.Value = True Then
        If Not gM46V999.gfPreenchido(txtNomLogradouro.Text) Then
            Call gM46V999.gpGraLog(1, "FNP0039 - " & gM46V999.gPrefixo & "Nome do logradouro não preenchido.")
            If txtNomLogradouro.Visible = True And txtNomLogradouro.Enabled = True Then
                txtNomLogradouro.SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    If chkBairro.Value = True Then
        If Not gM46V999.gfPreenchido(txtBairro.Text) Then
            Call gM46V999.gpGraLog(1, "FNP0036 - " & gM46V999.gPrefixo & "Bairro não preenchido.")
            If txtBairro.Visible = True And txtBairro.Enabled = True Then
                txtBairro.SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    If chkCidade.Value = True Then
        If Not gM46V999.gfPreenchido(txtCidade.Text) Then
            Call gM46V999.gpGraLog(1, "FNP0037 - " & gM46V999.gPrefixo & "Cidade não preenchida.")
            If txtCidade.Visible = True And txtCidade.Enabled = True Then
                txtCidade.SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    '2. Abre arquivo.
    If gM46V999.gfAbrBasDados("P0043000", gM46V999.gTipP0043000, gM46V999.gEndP0043000, lbdP0043000, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '3. Monta comando SELECT.
    lSelect = "SELECT * FROM P0043000"
    If chkCEP.Value = True Then
        lSelect = lSelect & " WHERE Num_CEP = '" & Left$(txtCEP.Text, 5) & Right$(txtCEP.Text, 3) & "'"
    End If
    If chkLogradouro.Value = True Then
        lSelect = lSelect & IIf(InStr(lSelect, "WHERE") = 0, " WHERE ", " AND ")
        If optLogradouro1(0) = True Then        'Nome completo.
            lSelect = lSelect & "Nom_Logr = '" & txtNomLogradouro.Text & "'"
        Else                                    'Parte do nome.
            If optLogradouro2(0) = True Then    'Início do nome.
                lSelect = lSelect & "Nom_Logr LIKE '" & gM46V999.gfForAspas(txtNomLogradouro.Text) & "%'"
            Else                                'Qualquer parte do nome.
                lSelect = lSelect & "Nom_Logr LIKE '%" & gM46V999.gfForAspas(txtNomLogradouro.Text) & "%'"
            End If
        End If
    End If
    If chkBairro.Value = True Then
        lSelect = lSelect & IIf(InStr(lSelect, "WHERE") = 0, " WHERE ", " AND ")
        If optBairro1(0) = True Then            'Nome completo.
            lSelect = lSelect & "Nom_Bairro = '" & txtBairro.Text & "'"
        Else                                    'Parte do nome.
            If optBairro2(0) = True Then        'Início do nome.
                lSelect = lSelect & "Nom_Bairro LIKE '" & gM46V999.gfForAspas(txtBairro.Text) & "%'"
            Else                                'Qualquer parte do nome.
                lSelect = lSelect & "Nom_Bairro LIKE '%" & gM46V999.gfForAspas(txtBairro.Text) & "%'"
            End If
        End If
    End If
    If chkCidade.Value = True Then
        lSelect = lSelect & IIf(InStr(lSelect, "WHERE") = 0, " WHERE ", " AND ")
        If optCidade1(0) = True Then            'Nome completo.
            lSelect = lSelect & "Nom_Cid = '" & txtCidade.Text & "'"
        Else                                    'Parte do nome.
            If optCidade2(0) = True Then        'Início do nome.
                lSelect = lSelect & "Nom_Cid LIKE '" & gM46V999.gfForAspas(txtCidade.Text) & "%'"
            Else                                'Qualquer parte do nome.
                lSelect = lSelect & "Nom_Cid LIKE '%" & gM46V999.gfForAspas(txtCidade.Text) & "%'"
            End If
        End If
    End If
    If chkUF.Value = True Then
        lSelect = lSelect & IIf(InStr(lSelect, "WHERE") = 0, " WHERE ", " AND ") & _
                            "Sig_UF = '" & cboUF.Text & "'"
    End If
    lSelect = lSelect & " ORDER BY Nom_Logr"

    '4. Busca dados.
    If InStr(lSelect, "LIKE") <> 0 Or InStr(lSelect, "AND") <> 0 Then
        lMensagem = "AGE0028 - " & gM46V999.gPrefixo & "Busca de informações DEMORADA e uma vez iniciado, não " & _
                    "há como interromper processo. Confirma busca de informações?"
        If MsgBox(lMensagem, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            Call gM46V999.gpGraLog(0, lMensagem & " Sim.")
        Else
            Call gM46V999.gpGraLog(0, lMensagem & " Não.")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    If gM46V999.gfObtRegistro(lbdP0043000, lSelect, lrsP0043000, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdP0043000)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If lrsP0043000.EOF = True Then
        lMensagem = "FNL0006 - " & gM46V999.gPrefixo & "Endereço não localizado"
        If chkCEP.Value = True Then
            lMensagem = lMensagem & " (CEP = " & txtCEP.Text
        End If
        If chkLogradouro.Value = True Then
            If optLogradouro1(0) = True Then        'Nome completo.
                If InStr(lMensagem, "0") = 0 Then
                    lMensagem = lMensagem & " (nome completo do logradouro = " & txtNomLogradouro.Text
                Else
                    lMensagem = lMensagem & "; nome completo do logradouro = " & txtNomLogradouro.Text
                End If
            Else                                    'Parte do nome.
                If optLogradouro2(0) = True Then    'Início do nome.
                    If InStr(lMensagem, "0") = 0 Then
                        lMensagem = lMensagem & " (início do nome do logradouro = " & _
                                    txtNomLogradouro.Text
                    Else
                        lMensagem = lMensagem & "; início do nome do logradouro = " & _
                                    txtNomLogradouro.Text
                    End If
                Else                                'Qualquer parte do nome.
                    If InStr(lMensagem, "0") = 0 Then
                        lMensagem = lMensagem & " (qualquer parte do nome do logradouro = " & _
                                    txtNomLogradouro.Text
                    Else
                        lMensagem = lMensagem & "; qualquer parte do nome do logradouro = " & _
                                    txtNomLogradouro.Text
                    End If
                End If
            End If
        End If
        If chkBairro.Value = True Then
            If optBairro1(0) = True Then            'Nome completo.
                If InStr(lMensagem, "0") = 0 Then
                    lMensagem = lMensagem & " (nome completo do bairro = " & txtBairro.Text
                Else
                    lMensagem = lMensagem & "; nome completo do bairro = " & txtBairro.Text
                End If
            Else                                    'Parte do nome.
                If optBairro2(0) = True Then        'Início do nome.
                    If InStr(lMensagem, "0") = 0 Then
                        lMensagem = lMensagem & " (início do nome do bairro = " & txtBairro.Text
                    Else
                        lMensagem = lMensagem & "; início do nome do bairro = " & txtBairro.Text
                    End If
                Else                                'Qualquer parte do nome.
                    If InStr(lMensagem, "0") = 0 Then
                        lMensagem = lMensagem & " (qualquer parte do nome do bairro = " & txtBairro.Text
                    Else
                        lMensagem = lMensagem & "; qualquer parte do nome do bairro = " & txtBairro.Text
                    End If
                End If
            End If
        End If
        If chkCidade.Value = True Then
            If optCidade1(0) = True Then            'Nome completo.
                If InStr(lMensagem, "0") = 0 Then
                    lMensagem = lMensagem & " (nome completo da cidade = " & txtCidade.Text
                Else
                    lMensagem = lMensagem & "; nome completo da cidade = " & txtCidade.Text
                End If
            Else                                    'Parte do nome.
                If optCidade2(0) = True Then        'Início do nome.
                    If InStr(lMensagem, "0") = 0 Then
                        lMensagem = lMensagem & " (início do nome da cidade = " & txtCidade.Text
                    Else
                        lMensagem = lMensagem & "; início do nome da cidade = " & txtCidade.Text
                    End If
                Else                                'Qualquer parte do nome.
                    If InStr(lMensagem, "0") = 0 Then
                        lMensagem = lMensagem & vbLf & " (qualquer parte do nome da cidade = " & _
                                    txtCidade.Text
                    Else
                        lMensagem = lMensagem & vbLf & "; qualquer parte do nome da cidade = " & _
                                    txtCidade.Text
                    End If
                End If
            End If
        End If
        If chkUF.Value = True Then
            If InStr(lMensagem, "0") = 0 Then
                lMensagem = lMensagem & " (unidade da federação = " & cboUF.Text
            Else
                lMensagem = lMensagem & "; unidade da federação = " & cboUF.Text
            End If
        End If
        lMensagem = lMensagem & ")."
        Call gM46V999.gpGraLog(1, lMensagem)
        Call gM46V999.gpFecha2(lbdP0043000)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '5. Monta grade.
    With flxCEP
        Screen.MousePointer = vbDefault
        mInterromper = False
        .Rows = 1
        .Visible = True
        Do While Not lrsP0043000.EOF
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0
            .Text = .Row
            .Col = 1
            If gM46V999.gfPreenchido(lrsP0043000!TIP_LOGR) Then
                .Text = lrsP0043000!TIP_LOGR
            End If
            .Col = 2
            If gM46V999.gfPreenchido(lrsP0043000!Nom_Logr) Then
                .Text = lrsP0043000!Nom_Logr
            End If
            .Col = 3
            If gM46V999.gfPreenchido(lrsP0043000!Nom_Bairro) Then
                .Text = lrsP0043000!Nom_Bairro
            End If
            .Col = 4
            If gM46V999.gfPreenchido(lrsP0043000!Nom_Cid) Then
                .Text = lrsP0043000!Nom_Cid
            End If
            .Col = 5
            If gM46V999.gfPreenchido(lrsP0043000!Sig_UF) Then
                .Text = lrsP0043000!Sig_UF
            End If
            .Col = 6
            If gM46V999.gfPreenchido(lrsP0043000!Num_CEP) Then
                .Text = Left$(lrsP0043000!Num_CEP, 5) & "-" & Right$(lrsP0043000!Num_CEP, 3)
            End If
            If .Row > 16 Then
                .TopRow = .Row - 15
            End If
            .Refresh
            DoEvents
            If mInterromper = True Then
                Exit Do
            End If
            lrsP0043000.MoveNext
        Loop
        Call gM46V999.gpFecha2(lbdP0043000)

        '6. Reposiciona janela.
        If .Rows <> 1 Then
            .Visible = True
            
            barFerramentas.Buttons(2).Enabled = False   'Pesquisar.
            barFerramentas.Buttons(4).Enabled = True    'Selecionar.
            .TopRow = 1
        End If
    End With

    Screen.MousePointer = vbDefault
End Sub
Private Sub mnuArqSelCEP_Click()
    Dim lI  As Byte     'Utilizado com For...Next.

    Screen.MousePointer = vbHourglass

    With flxCEP
        '1. Se grade vazia, encerra processo.
        If .Rows = 1 Then
            Unload Me
            Exit Sub
        End If

        '2. Se indicada linha de título, encerra processo.
        If .Row = 0 Then
            Unload Me
            Exit Sub
        End If
        .Col = 1        'Tipo do logradouro.
        mTipLogradouro = .Text
        .Col = 2        'Nome do logradouro.
        mLogradouro = .Text
        .Col = 3        'Bairro.
        mBairro = .Text
        .Col = 4        'Cidade.
        mCidade = .Text
        .Col = 5        'UF.
        mUF = .Text
        .Col = 6        'CEP.
        mCEP = .Text
    End With
    Unload Me
    Screen.MousePointer = vbDefault
End Sub
Private Sub mnuArqSaiCEP_Click()
    Unload Me
End Sub
Private Sub optBairro1_Click(Index As Integer)
    If flxCEP.Visible = True Then
        flxCEP.Visible = False
        
        barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
        barFerramentas.Buttons(4).Enabled = False   'Selecionar.
    End If
    Select Case Index
        Case 0      'Completo.
            optBairro2(0) = True
            optBairro2(1).Enabled = False
        Case 1      'Parte.
            optBairro2(1).Enabled = True
    End Select
End Sub
Private Sub optBairro1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optBairro2_Click(Index As Integer)
    If flxCEP.Visible = True Then
        flxCEP.Visible = False
        
        barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
        barFerramentas.Buttons(4).Enabled = False   'Selecionar.
    End If
End Sub
Private Sub optBairro2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optCidade1_Click(Index As Integer)
    If flxCEP.Visible = True Then
        flxCEP.Visible = False
        
        barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
        barFerramentas.Buttons(4).Enabled = False   'Selecionar.
    End If
    Select Case Index
        Case 0      'Completo.
            optCidade2(0) = True
            optCidade2(1).Enabled = False
        Case 1      'Parte.
            optCidade2(1).Enabled = True
    End Select
End Sub
Private Sub optCidade1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optCidade2_Click(Index As Integer)
    If flxCEP.Visible = True Then
        flxCEP.Visible = False
        
        barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
        barFerramentas.Buttons(4).Enabled = False   'Selecionar.
    End If
End Sub
Private Sub optCidade2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optLogradouro1_Click(Index As Integer)
    If flxCEP.Visible = True Then
        flxCEP.Visible = False
        
        barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
        barFerramentas.Buttons(4).Enabled = False   'Selecionar.
    End If
    Select Case Index
        Case 0      'Completo.
            optLogradouro2(0).Enabled = False
            optLogradouro2(1).Enabled = False
        Case 1      'Parte.
            optLogradouro2(0).Enabled = True
            optLogradouro2(1).Enabled = True
    End Select
End Sub
Private Sub optLogradouro1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optLogradouro2_Click(Index As Integer)
    If flxCEP.Visible = True Then
        flxCEP.Visible = False
        
        barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
        barFerramentas.Buttons(4).Enabled = False   'Selecionar.
    End If
End Sub
Private Sub optLogradouro2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtBairro_GotFocus()
    With txtBairro
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtBairro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        If flxCEP.Visible = True Then
            flxCEP.Visible = False
            
            barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
            barFerramentas.Buttons(4).Enabled = False   'Selecionar.
        End If
    End If
End Sub

Private Sub txtBairro_LostFocus()
    txtBairro.Text = gM46V999.gfRetirarCaracterEspecial(txtBairro.Text, True)
End Sub

Private Sub txtCEP_GotFocus()
    With txtCEP
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtCEP_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        If flxCEP.Visible = True Then
            flxCEP.Visible = False
            
            barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
            barFerramentas.Buttons(4).Enabled = False   'Selecionar.
        End If
    End If
End Sub
Private Sub txtCEP_LostFocus()
    If Len(Trim(txtCEP.Text)) = 9 Then
        Exit Sub
    End If
    If Len(Trim(txtCEP.Text)) = 8 Then
        If IsNumeric(txtCEP.Text) Then
            txtCEP.Text = Left$(txtCEP.Text, 5) & "-" & Right$(txtCEP.Text, 3)
        End If
    End If
End Sub
Private Sub txtCidade_GotFocus()
    With txtCidade
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        If flxCEP.Visible = True Then
            flxCEP.Visible = False
            
            barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
            barFerramentas.Buttons(4).Enabled = False   'Selecionar.
        End If
    End If
End Sub

Private Sub txtCidade_LostFocus()
    txtCidade.Text = gM46V999.gfRetirarCaracterEspecial(txtCidade.Text, True)
End Sub

Private Sub txtNomLogradouro_GotFocus()
    With txtNomLogradouro
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtNomLogradouro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        If flxCEP.Visible = True Then
            flxCEP.Visible = False
            
            barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
            barFerramentas.Buttons(4).Enabled = False   'Selecionar.
        End If
    End If
End Sub

Private Sub txtNomLogradouro_LostFocus()
    txtNomLogradouro.Text = gM46V999.gfRetirarCaracterEspecial(txtNomLogradouro.Text, True)
End Sub


