VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmT46V003A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pesquisa cota��o/proposta/ap�lice/protocolo"
   ClientHeight    =   5520
   ClientLeft      =   2655
   ClientTop       =   2475
   ClientWidth     =   10395
   Icon            =   "T46V003A.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picLegenda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   60
      Picture         =   "T46V003A.frx":000C
      ScaleHeight     =   345
      ScaleWidth      =   7425
      TabIndex        =   39
      Top             =   5070
      Width           =   7455
   End
   Begin Threed.SSFrame fraX 
      Height          =   2100
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   10335
      _Version        =   65536
      _ExtentX        =   18230
      _ExtentY        =   3704
      _StockProps     =   14
      Caption         =   "Crit�rios de pesquisa"
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
      Begin VB.ListBox lstAviso 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   450
         Left            =   810
         TabIndex        =   20
         Top             =   1245
         Visible         =   0   'False
         Width           =   7965
      End
      Begin Threed.SSFrame fraSitDocumento 
         Height          =   555
         Left            =   4305
         TabIndex        =   6
         Top             =   255
         Width           =   5940
         _Version        =   65536
         _ExtentX        =   10477
         _ExtentY        =   979
         _StockProps     =   14
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optSitDocumento 
            Caption         =   "Inspe��es n�o Solicitadas"
            ForeColor       =   &H00800000&
            Height          =   400
            Index           =   3
            Left            =   4515
            TabIndex        =   38
            Top             =   105
            Width           =   1395
         End
         Begin VB.OptionButton optSitDocumento 
            Caption         =   "Apenas propostas transmitidas"
            ForeColor       =   &H00800000&
            Height          =   400
            Index           =   2
            Left            =   2900
            TabIndex        =   9
            Top             =   100
            Width           =   1600
         End
         Begin VB.OptionButton optSitDocumento 
            Caption         =   "Apenas propostas n�o transmitidas"
            ForeColor       =   &H00800000&
            Height          =   400
            Index           =   1
            Left            =   1200
            TabIndex        =   8
            Top             =   100
            Width           =   1600
         End
         Begin VB.OptionButton optSitDocumento 
            Caption         =   "Todas propostas"
            ForeColor       =   &H00800000&
            Height          =   400
            Index           =   0
            Left            =   100
            TabIndex        =   7
            Top             =   100
            Width           =   1000
         End
      End
      Begin VB.TextBox txtCorretor 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6705
         MaxLength       =   40
         TabIndex        =   19
         Top             =   1300
         Width           =   3400
      End
      Begin VB.TextBox txtIniVigencia 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7110
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1650
         Width           =   1335
      End
      Begin VB.TextBox txtTerVigencia 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   8760
         MaxLength       =   10
         TabIndex        =   25
         Top             =   1650
         Width           =   1335
      End
      Begin VB.TextBox txtSegurado 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1100
         MaxLength       =   40
         TabIndex        =   13
         Top             =   1300
         Width           =   3350
      End
      Begin VB.TextBox txtProdutor 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1100
         MaxLength       =   40
         TabIndex        =   16
         Top             =   1650
         Width           =   3350
      End
      Begin VB.TextBox txtNumPedido 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7320
         MaxLength       =   6
         TabIndex        =   11
         Top             =   900
         Width           =   800
      End
      Begin Threed.SSFrame fraDocumento 
         Height          =   550
         Left            =   100
         TabIndex        =   2
         Top             =   250
         Width           =   4200
         _Version        =   65536
         _ExtentX        =   7408
         _ExtentY        =   970
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optDocumento 
            Caption         =   "Protocolo de transmiss�o"
            ForeColor       =   &H00800000&
            Height          =   400
            Index           =   2
            Left            =   2400
            TabIndex        =   5
            Top             =   100
            Width           =   1200
         End
         Begin VB.OptionButton optDocumento 
            Caption         =   "Ap�lices vigentes"
            ForeColor       =   &H00800000&
            Height          =   400
            Index           =   1
            Left            =   1400
            TabIndex        =   4
            Top             =   100
            Width           =   900
         End
         Begin VB.OptionButton optDocumento 
            Caption         =   "Cota��es e propostas"
            ForeColor       =   &H00800000&
            Height          =   400
            Index           =   0
            Left            =   100
            TabIndex        =   3
            Top             =   100
            Width           =   1200
         End
      End
      Begin Threed.SSFrame fraRamo 
         Height          =   495
         Left            =   105
         TabIndex        =   36
         Top             =   750
         Width           =   5370
         _Version        =   65536
         _ExtentX        =   9472
         _ExtentY        =   873
         _StockProps     =   14
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cboRamo 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "T46V003A.frx":7A0A
            Left            =   90
            List            =   "T46V003A.frx":7A0C
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   135
            Width           =   5200
         End
      End
      Begin VB.Label lblSegurado 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Segurado: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   100
         TabIndex        =   12
         Top             =   1300
         Width           =   1000
      End
      Begin VB.Label lblNumPedido 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N�mero do pedido: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5715
         TabIndex        =   10
         Top             =   900
         Width           =   1605
      End
      Begin VB.Label lblVigencia 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "de"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   6705
         TabIndex        =   22
         Top             =   1650
         Width           =   405
      End
      Begin VB.Label lblCodProdutor 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   1650
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblCodCorretor 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4550
         TabIndex        =   18
         Top             =   1300
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblVigencia 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "a"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   8460
         TabIndex        =   24
         Top             =   1650
         Width           =   300
      End
      Begin VB.Label lblVigencia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vig�ncia: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   5700
         TabIndex        =   21
         Top             =   1650
         Width           =   1005
      End
      Begin VB.Label lblProdutor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produtor: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   100
         TabIndex        =   14
         Top             =   1650
         Width           =   1000
      End
      Begin VB.Label lblCorretor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Corretor: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5700
         TabIndex        =   17
         Top             =   1305
         Width           =   1005
      End
   End
   Begin VB.PictureBox picUnchecked 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   8385
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   31
      Top             =   5000
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picChecked 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      Height          =   255
      Left            =   8085
      Picture         =   "T46V003A.frx":7A0E
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   30
      Top             =   5000
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSComctlLib.ImageList lstImagens 
      Left            =   8370
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V003A.frx":7F50
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V003A.frx":826A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V003A.frx":86BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V003A.frx":89D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V003A.frx":8E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V003A.frx":9142
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V003A.frx":945C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V003A.frx":9776
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V003A.frx":9A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V003A.frx":9DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V003A.frx":A0C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxResultado 
      Height          =   2100
      Left            =   0
      TabIndex        =   26
      Top             =   2895
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   3704
      _Version        =   393216
      Rows            =   7
      Cols            =   10
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   12582912
      BackColorSel    =   8454016
      ForeColorSel    =   8388608
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin Threed.SSCommand cmdProAntigos 
      Height          =   270
      Left            =   8715
      TabIndex        =   28
      Top             =   5040
      Visible         =   0   'False
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   476
      _StockProps     =   78
      Caption         =   "Protocolos antigos"
      ForeColor       =   8388608
   End
   Begin MSComctlLib.Toolbar barFerramentas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "lstImagens"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pesquisar"
            Object.ToolTipText     =   "Pesquisar documentos - F3."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Selecionar"
            Object.ToolTipText     =   "Selecionar documento - F6."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar crit�rios de pesquisa."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copiar"
            Object.ToolTipText     =   "Copiar documento selecionado."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir documento selecionado."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair da pesquisa."
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Height          =   300
      Left            =   -15
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Width           =   300
   End
   Begin VB.CheckBox chkTodos 
      Caption         =   "Marcar todos"
      ForeColor       =   &H00C00000&
      Height          =   200
      Left            =   1750
      TabIndex        =   33
      Top             =   2600
      Width           =   1250
   End
   Begin VB.CheckBox chkExcluir 
      Caption         =   "Marcar para excluir"
      ForeColor       =   &H00C00000&
      Height          =   200
      Left            =   50
      TabIndex        =   32
      Top             =   2600
      Width           =   1650
   End
   Begin Threed.SSPanel panMensagem 
      Height          =   705
      Left            =   525
      TabIndex        =   34
      Top             =   2160
      Visible         =   0   'False
      Width           =   9000
      _Version        =   65536
      _ExtentX        =   15875
      _ExtentY        =   1235
      _StockProps     =   15
      Caption         =   "Mensagem..."
      ForeColor       =   12582912
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BorderWidth     =   2
      BevelInner      =   1
      FloodColor      =   16761024
      FloodShowPct    =   0   'False
   End
   Begin VB.Label lblObservacao 
      AutoSize        =   -1  'True
      Caption         =   "* Proposta j� transmitida."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6480
      TabIndex        =   27
      Top             =   5280
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label LblObservacao2 
      AutoSize        =   -1  'True
      Caption         =   "* Cota��o expirada."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   8760
      TabIndex        =   35
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuArqPesquisar 
         Caption         =   "&Pesquisar"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArqSelecionar 
         Caption         =   "&Selecionar"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sai&r"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "&Editar"
      Begin VB.Menu mnuEdiLimpar 
         Caption         =   "&Limpar crit�rios de pesquisa"
      End
      Begin VB.Menu mnuEdiCopiar 
         Caption         =   "&Copiar documento selecionado"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuEdiExcluir 
         Caption         =   "E&xcluir documento selecionado"
      End
   End
   Begin VB.Menu mnuExibir 
      Caption         =   "E&xibir"
      Begin VB.Menu mnuExiProAntigos 
         Caption         =   "&Protocolos antigos"
      End
   End
End
Attribute VB_Name = "frmT46V003A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lCarga As Boolean    'Indicador de carga
Private mNum_Apolice As String    'N�mero da ap�lice selecionado.
Private mNum_Item As String    'N�mero do item selecionado.
Private mPesqApolice As Boolean    'Indica se a tela est� preparada para pesquisar ap�lice ou n�o.
'Somente ser� "true" se for a pesquisa para a tela de renova��o.
Private mRegChecked As String    'Registros selecionados para exclus�o (armazena "Nosso Numero").
Private mTipDocumento As Byte   'Tipo do documento.
Private mCod_Ramo As Integer
Private mCod_Modalidade As Integer
Private lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)

Private Sub mpDesControles()
'Procedure: desabilita controles (habilita pesquisar).

    On Error GoTo ControleErro

    '1. Grade apenas com cabe�alho - invis�vel.
    flxResultado.Rows = 1
    flxResultado.Visible = False

    '2. Op��o de pesquisar - HABILITAR.
    mnuArqPesquisar.Enabled = True
    barFerramentas.Buttons(2).Enabled = True    'Pesquisar.

    '3. Op��o de selecionar.
    mnuArqSelecionar.Enabled = False
    barFerramentas.Buttons(4).Enabled = False   'Selecionar.

    '4. Op��o de copiar documento selecionado.
    mnuEdiCopiar.Enabled = False
    barFerramentas.Buttons(7).Enabled = False   'Copiar documento selecionado.

    '5. Op��o de excluir documento selecionado.
    mnuEdiExcluir.Enabled = False
    barFerramentas.Buttons(8).Enabled = False   'Excluir documento selecionado.

    '6. Observa��o de proposta j� transmitida.
    '    lblObservacao.Visible = False

    '7. Bot�o de protocolos antigos.
    cmdProAntigos.Visible = False

    '8. Caixa de sele��o de Multiplos registros para exclus�o
    chkExcluir.Visible = False
    chkExcluir.Value = 0
    chkTodos.Visible = False
    chkTodos.Value = 0

    '1100773 - Sinaliza��o Consulta de Propostas e Cota��es
    picLegenda.Visible = (gM46V999.gTipUsuario <> e_Tipo_Usu_Yasuda)

    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mpDesControle ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub mpHabControles()
'Procedure: habilita controles

    On Error GoTo ControleErro

    '1. Op��o de selecionar - HABILITAR.
    mnuArqSelecionar.Enabled = True
    barFerramentas.Buttons(4).Enabled = True    'Selecionar.

    '2. Op��o de copiar documento selecionado - HABILITAR APENAS PARA COTA��ES E PROPOSTAS.
    If optDocumento(0) = True Then      'Cota��es e propostas.
        mnuEdiCopiar.Enabled = True
        barFerramentas.Buttons(7).Enabled = True    'Copiar documento selecionado.
    End If

    '3. Op��o de excluir documento selecionado - HABILITAR PARA COTA��ES E PROPOSTAS e PROTOCOLOS.
    If optDocumento(0) = True Or optDocumento(2) = True Then
        mnuEdiExcluir.Enabled = True
        barFerramentas.Buttons(8).Enabled = True    'Excluir documento selecionado.
    End If
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mpHabControles ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub mpMenNaoEncontrado()
    Dim lMensagem As String    'Mensagem.
    Dim lCod_Ramo As e_CodRamo

    lMensagem = "FNL0011 - N�o localizado registro que satisfa�a os crit�rios selecionados." & vbLf
    If gM46V999.gfPreenchido(txtNumPedido.Text) Then
        lMensagem = lMensagem & vbLf & "N�mero do pedido = " & txtNumPedido.Text & "."
    Else
        If optDocumento(0) = True Then
            lMensagem = lMensagem & vbLf & "Tipo de documento = cota��es e propostas"
            If optSitDocumento(0) = True Then
                lMensagem = lMensagem & "."
            Else
                If optSitDocumento(1) = True Then
                    lMensagem = lMensagem & " (apenas propostas n�o transmitidas)."
                Else
                    If optSitDocumento(2) = True Then
                        lMensagem = lMensagem & " (apenas propostas transmitidas)."
                    Else
                        lMensagem = lMensagem & " (apenas cota��es sem solicita��o de agM46V999.gEndamento)."
                    End If
                End If
            End If
        Else
            If optDocumento(1) = True Then
                lMensagem = lMensagem & vbLf & "Tipo de documento = ap�lices vigentes."
            Else
                lMensagem = lMensagem & vbLf & "Tipo de documento = protocolo de transmiss�o."
            End If
        End If
        If gM46V999.gfPreenchido(txtSegurado.Text) Then
            lMensagem = lMensagem & vbLf & "Nome do segurado que contenha """ & txtSegurado.Text & """."
        End If
        If gM46V999.gfPreenchido(txtProdutor.Text) Then
            lMensagem = lMensagem & vbLf & "Nome do produtor que contenha """ & txtProdutor.Text & """."
        End If
        If gM46V999.gfPreenchido(txtCorretor.Text) Then
            If gM46V999.gfPreenchido(lblCodCorretor.Caption) Then
                lMensagem = lMensagem & vbLf & "Corretor = " & lblCodCorretor.Caption & " - " & _
                            txtCorretor.Text & "."
            End If
        End If
        If gM46V999.gfPreenchido(txtIniVigencia.Text) And gM46V999.gfPreenchido(txtTerVigencia.Text) Then
            lMensagem = lMensagem & vbLf & "Vig�ncia: de " & txtIniVigencia.Text & " a " & _
                        txtTerVigencia.Text & "."
        Else
            If gM46V999.gfPreenchido(txtIniVigencia.Text) Then
                lMensagem = lMensagem & vbLf & "Vig�ncia em " & txtIniVigencia.Text & "."
            Else
                If gM46V999.gfPreenchido(txtTerVigencia.Text) Then
                    lMensagem = lMensagem & vbLf & "Vig�ncia em " & txtTerVigencia.Text & "."
                End If
            End If
        End If
        lCod_Ramo = cboRamo.ItemData(cboRamo.ListIndex)

        If lCod_Ramo <> e_CodRamo_NaoDefinido Then
            lMensagem = lMensagem & vbLf & "Ramo =" & cboRamo.List(cboRamo.ListIndex) & "."
        End If
    End If
    Call gM46V999.gpGraLog(1, lMensagem)
End Sub

Private Sub mpCarregarRamos()
'Procedure: obt�m Ramos.
    Dim lbdTbm_Ramos As ADODB.Connection    'Base de dados - Tbm_Ramos.
    Dim lMensagem As String     'Mensagem.
    Dim lrsTbm_Ramos As clsYasRecordSet    'Registro - Tbm_Ramos.
    Dim lSelect As String
    '1. Abre arquivo.
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Ramos, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    '2. Obt�m Todos os Ramos.
    lSelect = "SELECT * from Tbm_Ramo "
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then   'Corretor
        lSelect = lSelect & " where Tipo_Usuario = 'C' "
    End If
    lSelect = lSelect & " order by Cod_Ramo "

    If gM46V999.gfObtRegistro(lbdTbm_Ramos, lSelect, lrsTbm_Ramos, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTbm_Ramos)
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    If lrsTbm_Ramos.EOF = True Then
        Call gM46V999.gpFecha2(lbdTbm_Ramos)
        Call gM46V999.gpGraLog(1, "FNL0055 - " & gM46V999.gPrefixo & "Tabela de Ramos n�o localizado.")
        Exit Sub
    End If

    cboRamo.Clear
    cboRamo.AddItem "Todos os Ramos"
    cboRamo.ItemData(cboRamo.NewIndex) = 0

    '3. Carrega Combo
    Do While Not lrsTbm_Ramos.EOF
        cboRamo.AddItem lrsTbm_Ramos("Dsc_Ramo")
        cboRamo.ItemData(cboRamo.NewIndex) = lrsTbm_Ramos("Cod_Ramo")
        lrsTbm_Ramos.MoveNext
    Loop

    cboRamo.ListIndex = 0
    '4. Fecha

    Call gM46V999.gpFecha2(lbdTbm_Ramos)

End Sub

Private Sub mpPesApolices()
'Procedure: pesquisa cota��es e propostas.
    Dim lCod_Ramo As e_CodRamo

    Dim lMensagem As String    'Mensagem.
    Dim lListaNossoNumero As String         'Lista de nosso_numero.
    Dim lCount As Double    'Contador.
    Dim lrst As clsYasRecordSet    'Registro de ap�lice.
    Dim bdP00APOL As ADODB.Connection    'Arquivo de ap�lice (Conex�o).
    Dim lCpfCnpj As String    'Retorno da fun��o de forma��o de cpf e cnpj.
    Dim lCod_Corr As String    'Codigo do corretor.
    Dim lRegistro As String    'Registro.
    Dim lNomeTab As String
    Dim lSelect As String
    On Error GoTo ControleErro

    '1. Posiciona controles.
    panMensagem.Visible = True
    panMensagem.Caption = "Pesquisando..."
    Me.Refresh

    If gM46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
    Else
        lNomeTab = "GED.DBO.TAB_PED"
    End If

    '2. Abre base de dados
    If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, bdP00APOL, _
                              lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    '.3 Montagem do comando Select.
    lListaNossoNumero = ""
    If gM46V999.gfPreenchido(lblCodCorretor.Caption) Then
        lSelect = "SELECT * FROM " & lNomeTab & "_Corr"
        lSelect = lSelect & " Where Cod_Corr= " & lblCodCorretor.Caption

        If Not gM46V999.gfObtRegistro(bdP00APOL, lSelect, lrst, lMensagem) Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If
        lListaNossoNumero = ""
        Do While Not lrst.EOF
            If gM46V999.gfPreenchido(lListaNossoNumero) Then lListaNossoNumero = lListaNossoNumero & ","
            If gM46V999.gBaseLocal = True Then
                lListaNossoNumero = lListaNossoNumero & "'" & lrst("Nosso_Numero") & "'"
            Else
                lListaNossoNumero = lListaNossoNumero & "'" & lrst("Num_Pi") & "'"
            End If
            lrst.MoveNext
        Loop
    End If

    Call gM46V999.gfVerCorrPermitido(gM46V999.gCod_Corretor, 0, lRegistro)
    'Busco os dados da ap�lice

    'Ficha 1000782 - N�o permitir Renova��o LMI para Corretores
    lSelect = "SELECT distinct " & lNomeTab & ".* , " & lNomeTab & "_LOC.NOM_LOGR_RISCO "
    lSelect = lSelect & " FROM (" & lNomeTab & " inner join " & lNomeTab & "_Corr"
    lSelect = lSelect & " on " & lNomeTab & ".Num_APol = " & lNomeTab & "_Corr.NUm_Apol and"
    'Ficha 1000782 - N�o permitir Renova��o LMI para Corretores
    lSelect = lSelect & " " & lNomeTab & ".Num_ENDO = " & lNomeTab & "_Corr.NUm_ENDO) "
    lSelect = lSelect & " inner join " & lNomeTab & "_LOC "
    lSelect = lSelect & " on " & lNomeTab & "_LOC.NUM_APOL = " & lNomeTab & ".NUM_APOL "
    If Not gM46V999.gBaseLocal Then
        lSelect = lSelect & " AND " & lNomeTab & "_LOC.NUM_ENDO = " & lNomeTab & ".NUM_ENDO "
    End If

    If Not gM46V999.gBaseLocal = True Then
        lSelect = lSelect & " and " & lNomeTab & ".NUM_PI IN (SELECT NUM_PI "
        lSelect = lSelect & " From [tab_ctrl_emis]"
        lSelect = lSelect & " WHERE  [cod_iden_stat_acao] = 90 ) "
    End If

    'Nome do segurado.
    If gM46V999.gfPreenchido(txtSegurado.Text) Then
        lSelect = gM46V999.gfWhereAnd(lSelect) & " Nom_Segurado like '%" & UCase(txtSegurado.Text) & "%'"
    End If

    'In�cio da vig�ncia.
    If gM46V999.gfPreenchido(txtIniVigencia.Text) Then
        lSelect = gM46V999.gfWhereAnd(lSelect) & " Dat_Ini_Vig = " & gM46V999.gfForData(1, txtIniVigencia.Text)
    End If

    'T�rmino da vig�ncia.
    If gM46V999.gfPreenchido(txtTerVigencia.Text) Then
        lSelect = gM46V999.gfWhereAnd(lSelect) & " Dat_Ter_Vig = " & gM46V999.gfForData(1, txtTerVigencia.Text)
    End If

    'Produtor.
    If gM46V999.gfPreenchido(lblCodProdutor.Caption) Then
        lSelect = gM46V999.gfWhereAnd(lSelect) & " Cod_Prod = " & lblCodProdutor.Caption
    End If

    lCod_Ramo = cboRamo.ItemData(cboRamo.ListIndex)
    If lCod_Ramo <> e_CodRamo_NaoDefinido Then
        lSelect = gM46V999.gfWhereAnd(lSelect) & " " & lNomeTab & ".Cod_Ramo = " & lCod_Ramo
    End If

    If gM46V999.gfPreenchido(lListaNossoNumero) Then
        If gM46V999.gBaseLocal = True Then
            lSelect = gM46V999.gfWhereAnd(lSelect) & lNomeTab & ".Nosso_Numero in ( " & lListaNossoNumero & ")"
        Else
            lSelect = gM46V999.gfWhereAnd(lSelect) & lNomeTab & ".Num_pi in ( " & lListaNossoNumero & ")"
        End If
    End If

    'Ficha 1000782 - N�o permitir Renova��o LMI para Corretores
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        lSelect = gM46V999.gfWhereAnd(lSelect) & lNomeTab & "_LOC.NOM_LOGR_RISCO not like '%CONFORME ESPEC%'"
    End If

    lSelect = gM46V999.gfWhereAnd(lSelect) & " Dat_Ter_Vig >= " & gM46V999.gfForData(1, Date - 30)
    lSelect = gM46V999.gfWhereAnd(lSelect) & " cod_dizer not in (4,10) "
    lSelect = gM46V999.gfWhereAnd(lSelect) & lNomeTab & ".num_apol > 0 "

    lSelect = lSelect & " ORDER BY " & lNomeTab & ".Num_Ped"

    '4. obtenho o registro.
    If Not gM46V999.gfObtRegistro(bdP00APOL, lSelect, lrst, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    '5. Se n�o encontrou registro, informa usu�rio.
    If lrst.EOF Then
        Call mpMenNaoEncontrado
        panMensagem.Visible = False
    End If

    '6. Monta grade de resultados.

    With flxResultado
        'a) Posiciona quantidade de colunas e de linhas (apenas cabe�alho).
        .Cols = 11
        .Rows = 1
        'b) Posiciona linha de cabe�alho (com altura).
        .Row = 0
        'c) Posiciona cabe�alho nas colunas.
        .TextMatrix(0, 0) = "Excluir"
        .TextMatrix(0, 1) = "Ramo"
        .TextMatrix(0, 2) = "Ap�lice"
        .TextMatrix(0, 3) = "Segurado"
        .TextMatrix(0, 4) = "CNPJ/CPF"
        .TextMatrix(0, 5) = "In�cio Vig�ncia"
        .TextMatrix(0, 6) = "T�rmino Vig�ncia"
        .TextMatrix(0, 7) = "Nosso n�mero"
        .TextMatrix(0, 8) = "Situa��o do documento"
        .TextMatrix(0, 9) = "Produto/Modalidade"
        .TextMatrix(0, 10) = "CodRamo"
        'd) Posiciona largura das colunas.
        .ColWidth(0) = 0
        .ColWidth(1) = 1620
        .ColWidth(2) = 1200
        .ColWidth(3) = 2100
        .ColWidth(4) = 1300
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        'e) Posiciona alinhamento das colunas.
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        '.ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignCenterCenter
        .Visible = True
    End With
    lCount = 1

    '7. Coloca registros na grade.
    flxResultado.Redraw = False
    Do While Not lrst.EOF
        flxResultado.Rows = flxResultado.Rows + 1
        'b) Coluna 0 - Excluir - checkBox para marcar itens a serem exclu�dos.
        Set flxResultado.CellPicture = picUnchecked.Picture
        flxResultado.CellPictureAlignment = 4
        If lrst("Cod_Ramo") = e_CodRamo_Condominio Then
            flxResultado.TextMatrix(lCount, 1) = "Condom�nio"
        ElseIf lrst("Cod_Ramo") = e_CodRamo_Empresarial Then
            flxResultado.TextMatrix(lCount, 1) = "Empresarial"
        ElseIf lrst("Cod_Ramo") = e_CodRamo_Residencial Then
            flxResultado.TextMatrix(lCount, 1) = "Residencial"
        Else
            flxResultado.TextMatrix(lCount, 1) = "N�o definido"
        End If
        flxResultado.TextMatrix(lCount, 2) = lrst("Num_Apol")
        flxResultado.TextMatrix(lCount, 3) = lrst("Nom_Segurado")
        Select Case Len(CStr(lrst("Num_Cgc_Cpf")))
        Case Is > 11
            If IsNumeric(lrst("Num_Cgc_Cpf")) Then
                Call gM46V999.gpForCNPJ(1, Format$(lrst("Num_Cgc_Cpf"), "00000000000000"), lCpfCnpj)
            Else
                Call gM46V999.gpForCNPJ(1, Right$(String(14, "0") & Trim(lrst("Num_Cgc_Cpf")), 14), lCpfCnpj)
            End If
        Case Else
            Call gM46V999.gpForCPF(1, Format$(lrst("Num_Cgc_Cpf"), "00000000000"), lCpfCnpj)
        End Select
        flxResultado.TextMatrix(lCount, 4) = lCpfCnpj
        flxResultado.TextMatrix(lCount, 5) = gM46V999.gfForData(2, lrst("Dat_Ini_Vig"))
        flxResultado.TextMatrix(lCount, 6) = gM46V999.gfForData(2, lrst("Dat_Ter_Vig"))
        flxResultado.TextMatrix(lCount, 7) = 0
        flxResultado.TextMatrix(lCount, 8) = 0
        flxResultado.TextMatrix(lCount, 9) = lrst("Cod_Estip")
        flxResultado.TextMatrix(lCount, 10) = lrst("Cod_Ramo")
        lCount = lCount + 1
        '        flxResultado.Refresh
        lrst.MoveNext
    Loop
    panMensagem.Visible = False
    Me.Refresh

    flxResultado.Redraw = True
    If flxResultado.Rows > 1 Then
        flxResultado.TopRow = flxResultado.Rows - 1
    End If
    flxResultado.Refresh
    Call mpHabControles
    flxResultado.Row = flxResultado.Rows - 1
    flxResultado.Col = 0
    Call flxResultado_SelChange
    If flxResultado.Visible = True And flxResultado.Enabled = True Then
        flxResultado.SetFocus
    End If
    chkExcluir.Visible = False

    panMensagem.Visible = False

    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mpPesApolice ERRO - " & Err & " " & Err.Description)
End Sub

Private Sub mpPesCotPropostas()
'Procedure: pesquisa cota��es e propostas.

    Dim lMensagem As String    'Mensagem.
    Dim lListaNossoNumero As String         'Lista de nosso_numero.
    Dim lCount As Double    'Contador.
    Dim lI As Double    'Auxiliar.
    Dim lColuna As Integer    'Coluna do grid.
    Dim lObjPed As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED).
    Dim lObjstcPed As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED).
    Dim lObjPedCorr As clsA46V705A    'Objeto da classe clsA46V705A (TAB_PED_CORR).
    Dim lObjstcPedCorr As stcA46V705B    'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).
    Dim lCod_Ramo As e_CodRamo
    Dim lobjstcTrans As stcA46V718B
    Dim lobjTrans As clsA46V718A


    On Error GoTo ControleErro

    '1. Posiciona controles.
    panMensagem.Visible = True
    panMensagem.Caption = "Pesquisando..."
    Me.Refresh

    '2. Busco os registro com o corretor informado  (se informado).
    lListaNossoNumero = ""
    If gM46V999.gfPreenchido(lblCodCorretor.Caption) Then
        Set lObjPedCorr = New clsA46V705A
        Set lObjPedCorr.M46V999 = gM46V999

        Set lObjstcPedCorr = New stcA46V705B
        Set lObjstcPedCorr.M46V999 = gM46V999

        lObjstcPedCorr.Cod_Corr = lblCodCorretor.Caption
        If Not lObjPedCorr.mfCarregarColecao(lObjstcPedCorr, "", lListaNossoNumero) Then
            Call gM46V999.gpGraLog(1, lObjPedCorr.UltimoErro)
            panMensagem.Visible = False
            Exit Sub
        End If
    ElseIf gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        Set lObjPedCorr = New clsA46V705A
        Set lObjPedCorr.M46V999 = gM46V999

        Set lObjstcPedCorr = New stcA46V705B
        Set lObjstcPedCorr.M46V999 = gM46V999

        lObjstcPedCorr.Cod_Corr = gM46V999.gCod_Corretor
        If Not lObjPedCorr.mfCarregarColecao(lObjstcPedCorr, "", lListaNossoNumero) Then
            Call gM46V999.gpGraLog(1, lObjPedCorr.UltimoErro)
            panMensagem.Visible = False
            Exit Sub
        End If
    End If

    '3. Instancio os objetos
    Set lObjPed = New clsA46V702A
    Set lObjPed.M46V999 = gM46V999

    Set lObjstcPed = New stcA46V702B
    Set lObjstcPed.M46V999 = gM46V999

    lCod_Ramo = cboRamo.ItemData(cboRamo.ListIndex)
    If lCod_Ramo <> e_CodRamo_NaoDefinido Then
        lObjstcPed.Cod_Ramo = lCod_Ramo
    End If

    'N�mero do pedido.
    If gM46V999.gfPreenchido(txtNumPedido.Text) Then
        lObjstcPed.Num_Ped = txtNumPedido.Text
    End If
    'Situa��o do documento.
    If optSitDocumento(1) = True Then       'Propostas n�o transmitidas.
        lObjstcPed.Cod_Sit_Ped = e_SitPedido_Nenhum
    Else
        If optSitDocumento(2) = True Then       'Propostas transmitidas.
            lObjstcPed.Cod_Sit_Ped = e_SitPedido_Transmitido
        End If
    End If
    'Nome do segurado.
    If gM46V999.gfPreenchido(txtSegurado.Text) Then
        lObjstcPed.Nom_Segurado = txtSegurado.Text
    End If
    'Vig�ncia.
    If gM46V999.gfPreenchido(txtIniVigencia.Text) Then
        lObjstcPed.Dat_Ini_Vig = gM46V999.gfForData(1, txtIniVigencia.Text)
    End If
    If gM46V999.gfPreenchido(txtTerVigencia.Text) Then
        lObjstcPed.Dat_Ter_Vig = gM46V999.gfForData(1, txtTerVigencia.Text)
    End If
    'Produtor.
    If gM46V999.gfPreenchido(lblCodProdutor.Caption) Then
        lObjstcPed.Cod_Prod = lblCodProdutor.Caption
    End If
    lObjstcPed.Num_Apol = 9999999999#
    lObjstcPed.NossoNumeroIn = lListaNossoNumero

    If gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        lObjstcPed.Cod_Segurado = gM46V999.gCodSeguradoYas
    End If

    '4. Busco os dados com os par�metros de pesquisa informados.
    '###PROJETO INSPECAO - inclu�do o TRUE
    If Not lObjPed.mfCarregarColecao(lObjstcPed, "", "", True) Then    '1200708 - Limite Inicio de Vigencia - alterado para carregar itens dependentes
        Call gM46V999.gpGraLog(1, lObjPed.UltimoErro)
        panMensagem.Visible = False
        Exit Sub
    End If

    '5. Se n�o encontrou registro, informa usu�rio.
    If lObjPed.Colecao.Count = 0 Then
        Call mpMenNaoEncontrado
        GoTo Finalizar
    End If

    '6. Monta grade de resultados.
    With flxResultado
        'a) Posiciona quantidade de colunas e de linhas (apenas cabe�alho).
        .Cols = 12  '07/01/2010 - Ficha: 0901900 - C�pia de cota��o/proposta c/ �ltima tarifa
        .Rows = 1
        'b) Posiciona linha de cabe�alho (com altura).
        .Row = 0
        'c) Posiciona cabe�alho nas colunas.
        .TextMatrix(0, 0) = "Excluir"
        .TextMatrix(0, 1) = "Ramo"
        .TextMatrix(0, 2) = "Pedido"
        .TextMatrix(0, 3) = "Segurado"
        .TextMatrix(0, 4) = "CNPJ/CPF"
        .TextMatrix(0, 5) = "In�cio Vig�ncia"
        .TextMatrix(0, 6) = "T�rmino Vig�ncia"
        .TextMatrix(0, 7) = "Nosso n�mero"
        .TextMatrix(0, 8) = "Situa��o do documento"
        .TextMatrix(0, 9) = "Expira��o"
        .TextMatrix(0, 10) = "Ramo"    '07/01/2010 - Ficha: 0901900 - C�pia de cota��o/proposta c/ �ltima tarifa
        .TextMatrix(0, 11) = "Situa��o"
        'd) Posiciona largura das colunas.
        .ColWidth(0) = 0
        .ColWidth(1) = 1965
        .ColWidth(2) = 720
        .ColWidth(3) = 2100
        .ColWidth(4) = 1560
        .ColWidth(5) = 1155
        .ColWidth(6) = 1350
        .ColWidth(7) = 1500
        .ColWidth(8) = 1095
        .ColWidth(9) = 0
        .ColWidth(10) = 960
        .ColWidth(11) = 0
        'e) Posiciona alinhamento das colunas.
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .Visible = True
    End With
    lCount = 1

    '7. Coloca registros na grade.
    flxResultado.Redraw = False
    If optSitDocumento(3) = True Then
        For Each lObjstcPed In lObjPed.Colecao
            For Each lobjstcPedLoc In lObjstcPed.ColecaoPedLoc
                If lobjstcPedLoc.Insp_Obr = True And lobjstcPedLoc.Insp_Agend = False And lObjstcPed.DAT_DIGIT > 0 Then
                    If DateDiff("D", gM46V999.gfForData(2, lObjstcPed.DAT_DIGIT), Date) <= 35 Then
                        flxResultado.Rows = flxResultado.Rows + 1
                        'b) Coluna 0 - Excluir - checkBox para marcar itens a serem exclu�dos.
                        'Set flxResultado.CellPicture = picUnchecked.Picture
                        'flxResultado.CellPictureAlignment = 4
                        flxResultado.TextMatrix(lCount, 1) = lObjstcPed.gDesc_Ramo
                        flxResultado.TextMatrix(lCount, 2) = lObjstcPed.Num_Ped
                        flxResultado.TextMatrix(lCount, 3) = lObjstcPed.Nom_Segurado
                        flxResultado.TextMatrix(lCount, 4) = lObjstcPed.gCNPJ_CPF_Formatado
                        flxResultado.TextMatrix(lCount, 5) = gM46V999.gfForData(2, lObjstcPed.Dat_Ini_Vig)
                        flxResultado.TextMatrix(lCount, 6) = gM46V999.gfForData(2, lObjstcPed.Dat_Ter_Vig)
                        flxResultado.TextMatrix(lCount, 7) = lObjstcPed.NOSSO_NUMERO
                        flxResultado.TextMatrix(lCount, 8) = lObjstcPed.Cod_Sit_Ped
                        flxResultado.TextMatrix(lCount, 10) = lObjstcPed.Cod_Ramo  '07/01/2010 - Ficha: 0901900 - C�pia de cota��o/proposta c/ �ltima tarifa
                        flxResultado.TextMatrix(lCount, 11) = lObjstcPed.Cod_Tip_Doc
                        'Verifico se est� expirada.
                        flxResultado.TextMatrix(lCount, 9) = "0"
                        If lObjstcPed.Dat_Pri_Calculo <> 0 And (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
                            '                        If (gM46V999.gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gM46V999.gfForData(1, Format(Date, "dd/mm/yyyy"))) > 15 And lObjstcPed.Cod_Produto >= 978) Or _
                                                     '                           (gM46V999.gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gM46V999.gfForData(1, Format(Date, "dd/mm/yyyy"))) > 30 And lObjstcPed.Cod_Produto < 978) Then
                            '1200708 - Limite Inicio de Vigencia
                            If (gM46V999.gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gM46V999.gfForData(1, Format(Date, "dd/mm/yyyy"))) > 15 And lObjstcPed.Cod_Produto < 1044) Or _
                               ((lObjstcPed.Cod_Produto >= 1044) And _
                                (gM46V999.gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gM46V999.gfForData(1, Format(Date, "dd/mm/yyyy"))) > 30 And lobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda) Or _
                                (gM46V999.gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gM46V999.gfForData(1, Format(Date, "dd/mm/yyyy"))) > 15 And lobjstcPedLoc.Tip_Emissao <> e_TipEmissao_RenYasuda)) Then
                                'Cota��o expirada
                                flxResultado.TextMatrix(lCount, 9) = "1"
                                flxResultado.Row = lCount
                                lColuna = flxResultado.Col
                                For lI = 0 To (flxResultado.Cols - 1)
                                    flxResultado.Col = lI
                                    flxResultado.CellForeColor = &HC000C0
                                    '                                flxResultado.CellFontBold = True
                                Next lI
                                '                            If Not LblObservacao2.Visible Then
                                '                                LblObservacao2.Visible = True
                                '                            End If
                            End If
                        End If
                        If lObjstcPed.Cod_Sit_Ped = e_SitPedido_Transmitido Then          'J� transmitido.
                            flxResultado.Row = lCount
                            lColuna = flxResultado.Col
                            For lI = 0 To (flxResultado.Cols - 1)
                                flxResultado.Col = lI
                                flxResultado.CellForeColor = &HFF0000
                                '                            flxResultado.CellFontBold = True
                            Next lI
                            flxResultado.Col = lColuna
                            '                        lblObservacao.Visible = True
                        End If
                        lCount = lCount + 1
                        'Call flxResultado_SelChange
                    End If
                End If
            Next lobjstcPedLoc
        Next lObjstcPed
    Else
        For Each lObjstcPed In lObjPed.Colecao
            flxResultado.Rows = flxResultado.Rows + 1
            'b) Coluna 0 - Excluir - checkBox para marcar itens a serem exclu�dos.
            'Set flxResultado.CellPicture = picUnchecked.Picture
            'flxResultado.CellPictureAlignment = 4
            flxResultado.TextMatrix(lCount, 1) = lObjstcPed.gDesc_Ramo
            flxResultado.TextMatrix(lCount, 2) = lObjstcPed.Num_Ped
            flxResultado.TextMatrix(lCount, 3) = lObjstcPed.Nom_Segurado
            flxResultado.TextMatrix(lCount, 4) = lObjstcPed.gCNPJ_CPF_Formatado
            flxResultado.TextMatrix(lCount, 5) = gM46V999.gfForData(2, lObjstcPed.Dat_Ini_Vig)
            flxResultado.TextMatrix(lCount, 6) = gM46V999.gfForData(2, lObjstcPed.Dat_Ter_Vig)
            flxResultado.TextMatrix(lCount, 7) = lObjstcPed.NOSSO_NUMERO
            flxResultado.TextMatrix(lCount, 8) = lObjstcPed.Cod_Sit_Ped
            flxResultado.TextMatrix(lCount, 10) = lObjstcPed.Cod_Ramo  '07/01/2010 - Ficha: 0901900 - C�pia de cota��o/proposta c/ �ltima tarifa
            flxResultado.TextMatrix(lCount, 11) = lObjstcPed.Cod_Tip_Doc
            'Verifico se est� expirada.
            flxResultado.TextMatrix(lCount, 9) = "0"
            If lObjstcPed.Dat_Pri_Calculo <> 0 And (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
                '                If (gM46V999.gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gM46V999.gfForData(1, Format(Date, "dd/mm/yyyy"))) > 15 And lObjstcPed.Cod_Produto >= 978) Or _
                                 '                (gM46V999.gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gM46V999.gfForData(1, Format(Date, "dd/mm/yyyy"))) > 30 And lObjstcPed.Cod_Produto < 978) Then
                If lObjstcPed.ColecaoPedLoc.Count > 0 Then Set lobjstcPedLoc = lObjstcPed.ColecaoPedLoc(1)
                If (gM46V999.gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gM46V999.gfForData(1, Format(Date, "dd/mm/yyyy"))) > 15 And lObjstcPed.Cod_Produto < 1044) Or _
                   ((lObjstcPed.Cod_Produto >= 1044) And _
                    (gM46V999.gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gM46V999.gfForData(1, Format(Date, "dd/mm/yyyy"))) > 30 And lobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda) Or _
                    (gM46V999.gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gM46V999.gfForData(1, Format(Date, "dd/mm/yyyy"))) > 15 And lobjstcPedLoc.Tip_Emissao <> e_TipEmissao_RenYasuda)) Then
                    'Cota��o expirada
                    flxResultado.TextMatrix(lCount, 9) = "1"
                    flxResultado.Row = lCount
                    lColuna = flxResultado.Col
                    For lI = 0 To (flxResultado.Cols - 1)
                        flxResultado.Col = lI
                        flxResultado.CellForeColor = &HC000C0
                        '                        flxResultado.CellFontBold = True
                    Next lI
                    '                    If Not LblObservacao2.Visible Then
                    '                        LblObservacao2.Visible = True
                    '                    End If
                End If
            End If
            If lObjstcPed.Cod_Sit_Ped = e_SitPedido_Transmitido Then          'J� transmitido.
                flxResultado.Row = lCount
                lColuna = flxResultado.Col
                For lI = 0 To (flxResultado.Cols - 1)
                    flxResultado.Col = lI
                    flxResultado.CellForeColor = &HFF0000
                    '                    flxResultado.CellFontBold = True
                Next lI
                flxResultado.Col = lColuna
                '                lblObservacao.Visible = True
            End If


            '1100773 - Sinaliza��o Consulta de Propostas e Cota��es - Corre��o
            If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then    'And pObjstc.Cod_Sit_Ped = e_SitPedido_Transmitido Then
                Set lobjTrans = New clsA46V718A
                Set lobjTrans.M46V999 = gM46V999

                Set lobjstcTrans = New stcA46V718B
                lobjstcTrans.NOSSO_NUMERO = lObjstcPed.NOSSO_NUMERO
                If Not lobjTrans.mfCarregarColecao(lobjstcTrans, "", "") Then
                    'UltimoErro = objTransm.UltimoErro
                    Call gM46V999.gpGraLog(0, lobjTrans.UltimoErro)
                    Exit Sub
                End If
                If lobjTrans.Colecao.Count > 0 Then
                    Set lobjstcTrans = lobjTrans.Colecao.Item(1)
                    flxResultado.Row = lCount
                    lColuna = flxResultado.Col
                    For lI = 0 To (flxResultado.Cols - 1)
                        flxResultado.Col = lI
                        Select Case val(lobjstcTrans.Cod_Acao)
                        Case 80   'Recusada
                            flxResultado.CellForeColor = &HFF&
                        Case 81   'Cancelada
                            flxResultado.CellForeColor = &H80FF&
                        Case 90   'Emitida
                            flxResultado.CellForeColor = &H8000&
                        End Select
                    Next lI
                    flxResultado.Col = lColuna
                End If
                Set lobjTrans = Nothing
                Set lobjstcTrans = Nothing
            End If


            lCount = lCount + 1
            'Call flxResultado_SelChange
        Next lObjstcPed
    End If

    Call mpHabControles
    panMensagem.Visible = False
    Me.Refresh

    If flxResultado.Visible = True And flxResultado.Enabled = True Then
        flxResultado.SetFocus
    End If

    flxResultado.Redraw = True
    If flxResultado.Rows > 1 Then flxResultado.Row = 1
    flxResultado.Col = 0
    If flxResultado.Rows > 1 Then flxResultado.TopRow = 1
    flxResultado.Refresh
    '8. Mostra a caixa de sele��o para exclus�o multipla.
    chkExcluir.Visible = True
    chkExcluir.Value = 0
    '    Me.Height = 6000
    flxResultado.Height = 2100
    '    lblObservacao.Top = 5000
    panMensagem.Visible = False

    '9. Destruo os objetos.
    Set lObjPed = Nothing
    Set lObjstcPed = Nothing
    Set lObjPedCorr = Nothing
    Set lObjstcPedCorr = Nothing

    Exit Sub
Finalizar:
    panMensagem.Visible = False
    Me.Refresh
    Screen.MousePointer = vbDefault
    Set lObjPed = Nothing
    Set lObjstcPed = Nothing
    Set lObjPedCorr = Nothing
    Set lObjstcPedCorr = Nothing
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mpPesCotProposta ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub mpPesProTransmissao()
'Procedure: pesquisa protocolo de transmiss�o.

    Dim lData As String    'Campo auxiliar: data.
    Dim lDir As String    'Utilizado com a fun��o DIR.
    Dim lMensagem As String     'Mensagem.
    Dim lobjTransm As clsA46V718A    'Objeto da classe clsA46V718A (TAB_PED_TRANSM)
    Dim lobjstcTransm As stcA46V718B    'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM)

    On Error GoTo ControleErro

    '1. Instancio os objetos
    Set lobjTransm = New clsA46V718A
    Set lobjTransm.M46V999 = gM46V999

    Set lobjstcTransm = New stcA46V718B


    '2. Carrego os par�metro de pesquisa.
    'Obs.: Se n�mero do pedido est� preenchido, utiliza apenas esta informa��o no SELECT.
    If gM46V999.gfPreenchido(txtNumPedido.Text) Then
        lobjstcTransm.Num_Ped = Format$(txtNumPedido.Text, "000000")
    Else
        If gM46V999.gfPreenchido(txtSegurado.Text) Then
            lobjstcTransm.Nom_Segurado = gM46V999.gfForAspas(UCase(txtSegurado.Text))
        End If
        'b) Vig�ncia.
        If gM46V999.gfPreenchido(txtIniVigencia.Text) Then
            lData = gM46V999.gfForData(1, txtIniVigencia.Text)
            lobjstcTransm.Dat_Ini_Vig = lData
        End If
        If gM46V999.gfPreenchido(txtTerVigencia.Text) Then
            lData = gM46V999.gfForData(1, txtTerVigencia.Text)
            lobjstcTransm.Dat_Ter_Vig = lData
        End If
    End If

    '3. Obt�m registros de protocolo.
    If Not lobjTransm.mfCarregarColecao(lobjstcTransm, "", "") Then
        Call gM46V999.gpGraLog(1, lobjTransm.UltimoErro)
        Exit Sub
    End If

    '4. Se n�o encontrou registro, informa usu�rio.
    If lobjTransm.Colecao.Count = 0 Then
        lDir = Dir("*.Y01")
        If gM46V999.gfPreenchido(lDir) Then
            Call mnuExiProAntigos_Click
            Exit Sub
        Else
            Call mpMenNaoEncontrado
            Exit Sub
        End If
    End If

    '5. Monta grade de resultados.
    With flxResultado
        'a) Posiciona quantidade de colunas e de linhas (apenas cabe�alho).
        .Cols = 14
        .Rows = 1
        'b) Posiciona linha de cabe�alho (com altura).
        .Row = 0
        .RowHeight(0) = 480
        'c) Posiciona cabe�alho nas colunas.
        .TextMatrix(0, 0) = "Tipo"
        .TextMatrix(0, 1) = "Pedido"
        .TextMatrix(0, 2) = "Segurado"
        .TextMatrix(0, 3) = "In�cio de vig�ncia"
        .TextMatrix(0, 4) = "T�rmino de vig�ncia"
        .TextMatrix(0, 5) = "Tipo logradouro"
        .TextMatrix(0, 6) = "Logradouro"
        .TextMatrix(0, 7) = "N�mero logradouro."
        .TextMatrix(0, 8) = "Complemento"
        .TextMatrix(0, 9) = "Bairro"
        .TextMatrix(0, 10) = "Cidade"
        .TextMatrix(0, 11) = "UF"
        .TextMatrix(0, 12) = "CEP"
        .TextMatrix(0, 13) = "Nosso n�mero"
        'd) Posiciona largura das colunas.
        .ColWidth(0) = 1620
        .ColWidth(1) = 1000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1500
        .ColWidth(7) = 1000
        .ColWidth(8) = 2000
        .ColWidth(9) = 2000
        .ColWidth(10) = 2000
        .ColWidth(11) = 2000
        .ColWidth(12) = 2000
        .ColWidth(13) = 2000
        'e) Posiciona alinhamento das colunas.
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignLeftCenter
        .ColAlignment(8) = flexAlignLeftCenter
        .ColAlignment(9) = flexAlignLeftCenter
        .ColAlignment(10) = flexAlignLeftCenter
        .ColAlignment(11) = flexAlignLeftCenter
        .ColAlignment(12) = flexAlignLeftCenter
        .ColAlignment(13) = flexAlignLeftCenter

        .Visible = True
    End With

    '6. Coloca registros na grade.
    flxResultado.Redraw = False
    With flxResultado
        For Each lobjstcTransm In lobjTransm.Colecao
            'a) Acrescenta linha na grade.
            .Rows = .Rows + 1
            .Row = .Rows - 1
            '===1) Tipo do documento.
            .TextMatrix(.Row, 0) = "Protocolo"
            '===2) N�mero do pedido.
            .TextMatrix(.Row, 1) = Format$(lobjstcTransm.Num_Ped, "000000")
            '===3) Nome do segurado.
            If gM46V999.gfPreenchido(lobjstcTransm.Nom_Segurado) Then
                .TextMatrix(.Row, 2) = lobjstcTransm.Nom_Segurado
            Else
                .TextMatrix(.Row, 2) = ""
            End If
            '===4) In�cio de vig�ncia.
            If gM46V999.gfPreenchido(lobjstcTransm.Dat_Ini_Vig) Then
                .TextMatrix(.Row, 3) = gM46V999.gfForData(2, lobjstcTransm.Dat_Ini_Vig)
            Else
                .TextMatrix(.Row, 3) = ""
            End If
            '===5) T�rmino de vig�ncia.
            If gM46V999.gfPreenchido(lobjstcTransm.Dat_Ter_Vig) Then
                .TextMatrix(.Row, 4) = gM46V999.gfForData(2, lobjstcTransm.Dat_Ter_Vig)
            Else
                .TextMatrix(.Row, 4) = ""
            End If
            '===6) Tipo do logradouro
            If gM46V999.gfPreenchido(lobjstcTransm.Cod_Tip_Logr_Risco) Then
                .TextMatrix(.Row, 5) = lobjstcTransm.Cod_Tip_Logr_Risco
            Else
                .TextMatrix(.Row, 5) = ""
            End If
            '===7) logradouro.
            If gM46V999.gfPreenchido(lobjstcTransm.Nom_Logr_Risco) Then
                .TextMatrix(.Row, 6) = lobjstcTransm.Nom_Logr_Risco
            Else
                .TextMatrix(.Row, 6) = ""
            End If
            '===8) N�mero logradouro.
            If gM46V999.gfPreenchido(lobjstcTransm.Num_Risco) Then
                .TextMatrix(.Row, 7) = lobjstcTransm.Num_Risco
            Else
                .TextMatrix(.Row, 7) = ""
            End If
            '===9) Complemento.
            If gM46V999.gfPreenchido(lobjstcTransm.Dsc_Compl_Risco) Then
                .TextMatrix(.Row, 8) = lobjstcTransm.Dsc_Compl_Risco
            Else
                .TextMatrix(.Row, 8) = ""
            End If
            '===10) Complemento.
            If gM46V999.gfPreenchido(lobjstcTransm.Nom_Bairro_Risco) Then
                .TextMatrix(.Row, 9) = lobjstcTransm.Nom_Bairro_Risco
            Else
                .TextMatrix(.Row, 9) = ""
            End If
            '===11) Cidade.
            If gM46V999.gfPreenchido(lobjstcTransm.Nom_Cid_Risco) Then
                .TextMatrix(.Row, 10) = lobjstcTransm.Nom_Cid_Risco
            Else
                .TextMatrix(.Row, 10) = ""
            End If
            '===12) UF.
            If gM46V999.gfPreenchido(lobjstcTransm.Sig_UF_Risco) Then
                .TextMatrix(.Row, 11) = lobjstcTransm.Sig_UF_Risco
            Else
                .TextMatrix(.Row, 11) = ""
            End If
            '===13) CEP.
            If gM46V999.gfPreenchido(lobjstcTransm.Num_Cep_Risco) Then
                .TextMatrix(.Row, 12) = lobjstcTransm.Num_Cep_Risco
            Else
                .TextMatrix(.Row, 12) = ""
            End If
            '===14) Nosso n�mero.
            .TextMatrix(.Row, 13) = lobjstcTransm.NOSSO_NUMERO
            If .Row > 5 Then
                .TopRow = .Row - 4
            End If
            '.Refresh
        Next lobjstcTransm
        .Redraw = True
        '.TopRow = 1
        .Refresh
        Call mpHabControles
        '.Row = 1
        .Row = .Rows - 1
        .Col = 0
        Call flxResultado_SelChange
        If .Visible = True And .Enabled = True Then
            .SetFocus
        End If
    End With

    '7. Fecha tabelas e bases de dados.
    Set lobjstcTransm = Nothing
    Set lobjTransm = Nothing
    'cmdProAntigos.Visible = True
    flxResultado.Height = 2100
    '    Me.Height = 6000
    cmdProAntigos.Top = 5200
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mpPesProTransmissao ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub barFerramentas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case "PESQUISAR"
        Call mnuArqPesquisar_Click
    Case "SELECIONAR"
        Call mnuArqSelecionar_Click
    Case "LIMPAR"
        Call mnuEdiLimpar_Click
    Case "COPIAR"
        Call mnuEdiCopiar_Click
    Case "EXCLUIR"
        Call mnuEdiExcluir_Click
    Case "SAIR"
        Call mnuSair_Click
    End Select
    If UCase(Button.Key) <> "SAIR" Then
        If flxResultado.Enabled And flxResultado.Visible Then
            If flxResultado.Rows > 1 Then
                flxResultado.ColSel = flxResultado.Cols - 1
                flxResultado.SetFocus
            End If
        End If
    Else
        If gM46V999.gfDisponivel("frmT46V002A") = True Then
            frmT46V002A.SetFocus
            Exit Sub
        End If
        If gM46V999.gfDisponivel("frmT46V011A") = True Then
            frmT46V011A.SetFocus
            Exit Sub
        End If
        If gM46V999.gfDisponivel("frmT46V178A") = True Then
            frmT46V178A.SetFocus
            Exit Sub
        End If
        If gM46V999.gfDisponivel("frmT46V001A") = True Then
            frmT46V001A.SetFocus
        End If



    End If
End Sub

Private Sub chkExcluir_Click()
'Exibe a coluna para marcar registros a serem excluidos.

    If chkExcluir.Value = 1 Then
        '1- Mostra a caixa de sele��o utilizada para marcar todos s registros
        chkTodos.Visible = True
        chkTodos.Value = 0
        '2- Exibe a coluna utilizada para exclus�o
        With flxResultado
            .ColWidth(0) = 700
        End With
        '3-
        ' Op��o de selecionar.
        mnuArqSelecionar.Enabled = False
        barFerramentas.Buttons(4).Enabled = False   'Selecionar.

        ' Op��o de copiar documento selecionado.
        mnuEdiCopiar.Enabled = False
        barFerramentas.Buttons(7).Enabled = False   'Copiar documento selecionado.
        '4- Desmarca todos os registros
        Call mpAlteraTodos(False)
        '5- Chama a fun��o para remover a cor verde da linha selecionada
        flxResultado.CellBackColor = vbWhite
        Call flxResultado_SelChange
    Else
        '1- Oculta a caixa de sele��o utilizada para marcar todos s registros
        chkTodos.Visible = False
        chkTodos.Value = 0
        '2- Exibe a coluna utilizada para exclus�o
        With flxResultado
            .ColWidth(0) = 0
        End With
        '3-
        ' Op��o de selecionar.
        mnuArqSelecionar.Enabled = True
        barFerramentas.Buttons(4).Enabled = True   'Selecionar.

        ' Op��o de copiar documento selecionado.
        mnuEdiCopiar.Enabled = True
        barFerramentas.Buttons(7).Enabled = True   'Copiar documento selecionado.
        '4- Desmarca todos os registros
        Call mpAlteraTodos(False)
    End If
End Sub
Private Sub chkTodos_Click()

    If chkTodos.Value = 1 Then
        'Marca todos os registros
        Call mpAlteraTodos(True)
    Else
        'Desmarca todos
        Call mpAlteraTodos(False)
    End If

End Sub

Private Sub cmdProAntigos_Click()
    Call mnuExiProAntigos_Click
End Sub
Private Sub cmdSair_Click()
'Observa��o: controle criado para usu�rio sair do FORM utilizando tecla ESC. Invis�vel pelo usu�rio.
    Unload Me
End Sub
Private Sub flxResultado_Click()
    Dim lColuna As Integer    'Coluna.
    Dim lLinha As Integer    'Linha.
    Dim lSelecao As String    'Sele��o.

'1. Somente quando est� selecionada a op��o "Marcar para Excluir" e o tipo da consulta �
'   "Cota��es e Propostas" - para executar o item 2 quando clicado em qualquer coluna da grade.
    If mTipDocumento = 1 And chkExcluir.Visible = True And chkExcluir.Value = 1 Then
        'Altera posi��o da coluna.
        flxResultado.Col = 0
    Else
        Exit Sub
    End If

    '2. Marca ou desmarca o registro na grade.
    With flxResultado
        lColuna = .Col
        lLinha = .Row
        If flxResultado.Col = 0 Then
            If flxResultado.CellPicture = picChecked Then
                Set flxResultado.CellPicture = picUnchecked
                .Col = .Col + 7       'Coluna de nosso n�mero.
                lSelecao = .Text
                'Quando desmarcar um registro, remove o n�mero de pedido da lista.
                mRegChecked = Replace(mRegChecked, lSelecao & ",", "")
                'Antes de utilizar o campo no comando SQL, n�o esquecer de remover a �ltima v�rgula
                'contida na vari�vel.
            Else
                Set flxResultado.CellPicture = picChecked
                .CellPictureAlignment = flexAlignCenterCenter
                .Col = .Col + 7        'Coluna de nosso n�mero.
                lSelecao = .Text
                mRegChecked = mRegChecked & lSelecao & ","
            End If
        End If
        .Col = lColuna
        .Row = lLinha


    End With
End Sub
Private Sub flxResultado_DblClick()
'Somente quando est� selecionada a op��o "Marcar para Excluir".
    If mTipDocumento = 1 And chkExcluir.Visible = True And chkExcluir.Value = 1 Then
        'Est� halibitado para selecionar arquivos para excluir - sai do procedimento.
        Exit Sub
    End If

    'Abre o arquivo selecionado.
    Call mnuArqSelecionar_Click
End Sub
Private Sub flxResultado_SelChange()
    Dim lColuna As Integer      'Coluna.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lJ As Integer    'Utilizado com For...Next.
    Dim lLinha As Integer     'Linha.

    With flxResultado
        '1. Posiciona linha e coluna.
        lLinha = .Row
        lColuna = .Col

        '2. Se linha j� est� com fundo verde, nada a fazer.
        'If .CellBackColor = vbGreen Then
        '    Exit Sub
        'End If

        '.Redraw = False

        '3. Deixa todas as linhas com fundo branco.
        'For lI = 1 To (.Rows - 1)
        '    .Row = lI
        '    For lJ = 0 To (.Cols - 1)
        '        .Col = lJ
        '        .CellBackColor = vbWhite
        '    Next lJ
        'Next lI

        'Somente quando esta selecionado a op��o "Marcar para Excluir"
        If mTipDocumento = 1 _
           And chkExcluir.Visible = True _
           And chkExcluir.Value = 1 Then
            'Esta halibitado para selecionar arquivos para excluir

            'Op��o de copiar documento selecionado.
            mnuEdiCopiar.Enabled = False
            barFerramentas.Buttons(7).Enabled = False   'Copiar documento selecionado.

            'Reposiciona linha e coluna
            '.Row = lLinha
            '.Col = lColuna

            'Sai do procedimento
            Exit Sub
        End If

        '4. Deixa linha selecionada com fundo verde.
        '.Row = lLinha
        'For lJ = 0 To (.Cols - 1)
        '    .Col = lJ
        '    .CellBackColor = vbGreen
        'Next lJ

        '5. Reposiciona linha e coluna.
        '.Row = lLinha
        '.Col = lColuna
        '.Redraw = True
        '.Refresh
    End With


End Sub
Private Sub Form_Activate()

'1. Posiciona nome do form
    frmT46V001A.panForm = Mid$(Me.Name, 4)

    '2. Centraliza form.
    Call gM46V999.gpCenForm(Me)
End Sub
Private Sub Form_Load()

'1. Desabilita controles.
    Call mpDesControles

    '2. Preencher Combo de Ramo com Tabela Tbm_Ramo
    Call mpCarregarRamos

End Sub
Private Sub mpPesquisaDocumento()
    Dim lMensagem As String    'Mensagem.
    Dim lTemNumPedido As Boolean    'Indica que tem n�mero do pedido para pesquisa.

    On Error GoTo ControleErro
    '1. Executa eventos LostFocus pendentes.
    '   Se foco est� em algum campo que possue evento LostFocus, o mesmo deve ser acionado.
    '   Observa��o: � necess�rio verificar todos os campos que possuem evento LOSTFOCUS.
    Select Case UCase$(ActiveControl.Name)
    Case "TXTCORRETOR"
        If gM46V999.gfPreenchido(txtCorretor.Text) Then
            Call txtCorretor_LostFocus
            If Not gM46V999.gfPreenchido(lblCodCorretor.Caption) Then
                Exit Sub
            End If
        End If
    Case "TXTPRODUTOR"
        If gM46V999.gfPreenchido(txtProdutor.Text) Then
            Call txtProdutor_LostFocus
            If Not gM46V999.gfPreenchido(lblCodProdutor.Caption) Then
                Exit Sub
            End If
        End If
    Case "TXTINIVIGENCIA"
        If gM46V999.gfPreenchido(txtIniVigencia.Text) Then
            Call txtIniVigencia_LostFocus
        End If
    Case "TXTTERVIGENCIA"
        If gM46V999.gfPreenchido(txtTerVigencia.Text) Then
            Call txtTerVigencia_LostFocus
        End If
    End Select

    '2. Consiste crit�rios de pesquisa.
    'a) Tipo de documento.
    If optDocumento(0) = False And optDocumento(1) = False And optDocumento(2) = False Then
        Call gM46V999.gpGraLog(1, "FNP0013 - Tipo de documento n�o selecionado.")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    'b) N�mero do pedido preenchido.
    lTemNumPedido = False
    If txtNumPedido.Visible = True Then
        If gM46V999.gfPreenchido(txtNumPedido.Text) Then
            If Not IsNumeric(txtNumPedido.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0080 - N�mero do pedido n�o num�rico.")
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            lTemNumPedido = True
        End If
    End If

    'c) Ramo
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        If gM46V999.gfPreenchido(txtSegurado.Text) = False And gM46V999.gfPreenchido(txtNumPedido.Text) = False Then
            If cboRamo.ListIndex = 0 Then
                Call gM46V999.gpGraLog(1, "FRA0001 - Ramo n�o selecionado.")
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If

    'd) N�mero do pedido n�o preenchido.
    If lTemNumPedido = False Then
        '=== Vig�ncia.
        If txtIniVigencia.Visible = True Then
            If gM46V999.gfPreenchido(txtIniVigencia.Text) Then
                If gM46V999.gfVerData(2, txtIniVigencia.Text) = False Then
                    Call gM46V999.gpGraLog(1, "FGE0018 - Vig�ncia inv�lida (data de in�cio de vig�ncia inv�lida).")
                    If txtIniVigencia.Visible = True And txtIniVigencia.Enabled = True Then
                        txtIniVigencia.SetFocus
                    End If
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
        If txtTerVigencia.Visible = True Then
            If gM46V999.gfPreenchido(txtTerVigencia.Text) Then
                If gM46V999.gfVerData(2, txtTerVigencia.Text) = False Then
                    Call gM46V999.gpGraLog(1, "FGE0081 - Vig�ncia inv�lida (data de t�rmino de vig�ncia " & _
                                              "inv�lida).")
                    If txtTerVigencia.Visible = True And txtTerVigencia.Enabled = True Then
                        txtTerVigencia.SetFocus
                    End If
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
        If txtIniVigencia.Visible = True And txtTerVigencia.Visible = True And _
           gM46V999.gfPreenchido(txtIniVigencia.Text) And gM46V999.gfPreenchido(txtTerVigencia.Text) Then
            If gM46V999.gfForData(1, txtIniVigencia.Text) > gM46V999.gfForData(1, txtTerVigencia.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0082 - Vig�ncia inv�lida (data de in�cio de vig�ncia " & _
                                          txtIniVigencia.Text & ") maior que a data de t�rmino de vig�ncia (" & _
                                          txtTerVigencia.Text & ").")
                If txtIniVigencia.Visible = True And txtIniVigencia.Enabled = True Then
                    txtIniVigencia.SetFocus
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If

    '3. Obt�m dados.
    lMensagem = ""
    Select Case mTipDocumento
    Case 1
        Call mpPesCotPropostas
    Case 2
        Call mpPesApolices
    Case 3
        Call mpPesProTransmissao
    Case Else
        Call gM46V999.gpGraLog(1, "FNP0013 - Tipo de documento n�o selecionado.")
        Screen.MousePointer = vbDefault
        Exit Sub
    End Select
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mpPesquisaDocumento ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mPesqApolice = False
End Sub

Private Sub mnuArqPesquisar_Click()
    Screen.MousePointer = vbHourglass
    Call mpPesquisaDocumento
    Screen.MousePointer = vbDefault
End Sub
Private Sub mnuArqSelecionar_Click()
    Dim lMensagem As String    'Mensagem.
    Dim lNosNumero As String    'Nosso n�mero.
    Dim lNumApolice As String    'N�mero da ap�lice.
    Dim lNumIteApolice As String  'N�mero do item da ap�lice.
    Dim lNumPedido As String    'N�mero de pedido.
    Dim lSitDocumento As Byte   'Situa��o do documento.

    Screen.MousePointer = vbHourglass
    On Error GoTo ControleErro

    With flxResultado
        '1. Verifica se tem documento selecionado.
        If .Row = 0 Then
            Call gM46V999.gpGraLog(1, "FNP0014 - Documento n�o selecionado.")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'If .CellBackColor <> vbGreen Then
        '    Call gM46V999.gpGraLog(1, "FNP0014 - Documento n�o selecionado.")
        '    Screen.MousePointer = vbDefault
        '    Exit Sub
        'End If
        'Caso cota��o expirada exibe mensagem com op��o para exclus�o
        If mTipDocumento = 1 Then
            If .Cols > 11 Then
                If .TextMatrix(.Row, 11) = 1 Then
                    If .TextMatrix(.Row, 9) = 1 Then
                        lMensagem = "AGE0046 - Cota��o " & lNumPedido & " (nosso n�mero " & _
                                    lNosNumero & ") expirou o prazo de validade, e n�o ser� exibida. Deseja excluir este item da lista?"
                        If MsgBox(lMensagem, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                            Call gM46V999.gpGraLog(0, lMensagem & " SIM.")
                            Call mnuEdiExcluir_Click
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        Else
                            Call gM46V999.gpGraLog(0, lMensagem & " N�O.")
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        '2. Verifica tipo do documento.
        If optDocumento(0) = False And optDocumento(1) = False And optDocumento(2) = False Then
            Call gM46V999.gpGraLog(1, "FNP0013 - Tipo de documento n�o selecionado.")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        '3. Trata sele��o.
        If mPesqApolice = True Then
            'Se for pesquisa de ap�lice , dever� ent�o retornar a ap�lice para a tela de origem
            mNum_Apolice = .TextMatrix(.Row, 2)      'N�mero da ap�lice.
            mNum_Item = 0
            mCod_Ramo = .TextMatrix(.Row, 10)
            mCod_Modalidade = .TextMatrix(.Row, 9)
            Unload Me
        Else
            Select Case mTipDocumento
            Case 1      'Cota��o ou proposta - grava registro de comunica��o.
                lNumPedido = .TextMatrix(.Row, 2)       'N�mero do pedido.
                lNosNumero = .TextMatrix(.Row, 7)       'Nosso n�mero.
                lSitDocumento = .TextMatrix(.Row, 8)    'Situa��o do documento.
                If lSitDocumento = 3 Then               'J� transmitido.
                    lMensagem = "AGE0019 - Proposta " & lNumPedido & " (nosso n�mero " & _
                                lNosNumero & ") j� transmitida. Confirma abertura desta proposta?"
                    If MsgBox(lMensagem, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        Call gM46V999.gpGraLog(0, lMensagem & " SIM.")
                    Else
                        Call gM46V999.gpGraLog(0, lMensagem & " N�O.")
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                If gM46V111.gfGraRegComunicacao(lNosNumero, lMensagem) = False Then
                    Call gM46V999.gpGraLog(2, lMensagem)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                Me.Hide
            Case 2      'Ap�lice - cria renova��o de ap�lice.
                lNumApolice = .TextMatrix(.Row, 2)     'N�mero da ap�lice.
                lNumIteApolice = 0
                lMensagem = ""
                If gM46V001.gfGraRenYasuda(lNumApolice, lNumIteApolice, lNosNumero, _
                                           lMensagem, lMensagem, "") = False Then
                    If gM46V999.gfPreenchido(lMensagem) Then
                        Call gM46V999.gpGraLog(1, lMensagem)
                    End If
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                    If gM46V999.gfPreenchido(lMensagem) Then
                        Call gM46V999.gpGraLog(1, lMensagem)
                    End If
                End If
                Me.Hide
            Case 3      'Protocolo de transmiss�o - exibe informa��es do protocolo.
                gM46V111.gNumPedido = .TextMatrix(.Row, 1)   'N�mero do pedido.
                gM46V111.gComNosNumero = .TextMatrix(.Row, 13)  'Nosso n�mero.
                frmT46V006A.Show 1
            Case Else
                Call gM46V999.gpGraLog(1, "FTE0015 - Tipo de documento desconhecido.")
                Screen.MousePointer = vbDefault
                Exit Sub
            End Select
        End If
    End With

    Screen.MousePointer = vbDefault
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V003A.mnuArqSelecionar_Click ERRO - " & Err & " " & Err.Description)
End Sub

Private Sub mnuEdiCopiar_Click()
    Dim lAntNosNumero As String     'Anterior: nosso n�mero.
    Dim lAntNumPedido As String     'Anterior: n�mero do pedido.
    Dim lMensagem As String    'Mensagem.
    Dim lNovNosNumero As String     'Novo: nosso n�mero.
    Dim lNovNumPedido As String     'Novo: n�mero do pediro.
    Dim lNumApolice As String   'N�mero da ap�lice.
    Dim lReentrada As Integer
    Dim objPed As clsA46V702A
    Dim objstcPed As stcA46V702B
    Dim objTransm As clsA46V718A    'Objeto da classe clsA46V718A (TAB_PED_TRANSM).
    Dim objstcTransm As stcA46V718B   'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM).

    On Error GoTo ControleErro
    Screen.MousePointer = vbHourglass

    With flxResultado
        '1. Verifica se tem documento selecionado.
        If .Row = 0 Then
            Call gM46V999.gpGraLog(1, "FNP0014 - Documento n�o selecionado.")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'If .CellBackColor <> vbGreen Then
        '    Call gM46V999.gpGraLog(1, "FNP0014 - Documento n�o selecionado.")
        '    Screen.MousePointer = vbDefault
        '    Exit Sub
        'End If

        '2. Posiciona controles.
        gM46V999.gPrefixo = " Pedido " & .TextMatrix(.Row, 2) & ". "
        lAntNumPedido = .TextMatrix(.Row, 2)    'N�mero do pedido.
        lAntNosNumero = .TextMatrix(.Row, 7)   'Nosso n�mero.

        '3. Confirma com usu�rio.
        lMensagem = "AGE0009 - " & gM46V999.gPrefixo & "Confirma c�pia deste documento?"
        If MsgBox(lMensagem, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Call gM46V999.gpGraLog(0, lMensagem & " N�O.")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Call gM46V999.gpGraLog(0, lMensagem & " SIM.")
        lReentrada = 2
        If .TextMatrix(.Row, 8) = "3" Then    'Proposta transmitida
            'Verifico se � a c�pia de uma proposta cancelada
            Set objTransm = New clsA46V718A      'Objeto da classe clsA46V718A (TAB_PED_TRANSM).
            Set objTransm.M46V999 = gM46V999

            Set objstcTransm = New stcA46V718B         'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM).
            objstcTransm.NOSSO_NUMERO = lAntNosNumero
            If Not objTransm.mfCarregarColecao(objstcTransm, "", "") Then
                Call gM46V999.gpGraLog(1, objTransm.UltimoErro)
                Exit Sub
            End If
            If objTransm.Colecao.Count > 0 Then
                Set objstcTransm = objTransm.Colecao.Item(1)

                If val(objstcTransm.Cod_Acao) = 80 Or val(objstcTransm.Cod_Acao) = 81 Then
                    If val(objstcTransm.Dat_Emissao) = 0 Then GoTo Reentrada
                    If gM46V999.gfForData(1, Format(CDate(gM46V999.gfForData(2, objstcTransm.Dat_Emissao)) + 7, "dd/mm/yyyy")) >= gM46V999.gHoje Then
Reentrada:
                        lMensagem = "Esta � uma proposta transmitida. Trata-se de uma reentrada de proposta?"
                        If MsgBox(lMensagem, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                            lReentrada = 1
                        End If
                    End If
                End If
            End If
        End If
        '4. Calcula nosso n�mero e n�mero do pedido.
        If gM46V999.gfCalNosNumero(lNovNosNumero, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If gM46V999.gfCalNumPedido(lNovNumPedido, lMensagem) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        Dim lCodProduto As Integer
        Dim lInputBox As String
        lInputBox = InputBox("Digite a tarifa que ser� usada na copia (em branco usar� a tarifa da original): ", Me.Caption)
        
        Dim lMsg_Erro As String
        Dim lTarifaAtual_Impar As Integer
        
        lTarifaAtual_Impar = gM46V999.gfTarifaAtual_Impar(lMsg_Erro)
        If gM46V999.gfPreenchido(lMsg_Erro) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - " & lMsg_Erro)
            Exit Sub
        End If
        
        If lInputBox <> Empty Then
            If Not IsNumeric(lInputBox) Then
                Call gM46V999.gpGraLog(1, "AGE0002 - Tarifa inv�lida.")
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                If lInputBox > lTarifaAtual_Impar Then
                    Call gM46V999.gpGraLog(1, "AGE0002 - Tarifa inv�lida.")
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            lCodProduto = lInputBox
        Else
            lCodProduto = 0
        End If
        
        '5. Copia documento.
        gM46V999.gOrigemChamada = e_Origem_Kit
        If gM46V999.gfCopProposta(lAntNosNumero, lNovNosNumero, lNovNumPedido, lMensagem, 0, 0, .TextMatrix(.Row, 10), , , lReentrada, lCodProduto) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If lReentrada = 1 Then
            Set objPed = New clsA46V702A
            Set objPed.M46V999 = gM46V999

            Set objstcPed = New stcA46V702B
            Set objstcPed.M46V999 = gM46V999

            objstcPed.NOSSO_NUMERO = lNovNosNumero
            If Not objPed.mfAlterarReentrada(objstcPed, "", lReentrada, lAntNosNumero, "") Then
                Call gM46V999.gpGraLog(1, objPed.UltimoErro)
                Exit Sub
            End If
        End If

        'g) Conclus�o.
        Call gM46V999.gpGraLog(1, "AGE0002 - Processo conclu�do: c�pia realizada para pedido " & lNovNumPedido & _
                                " (nosso n�mero " & lNovNosNumero & ").")

        '6. Abre novo documento.
        .Col = 0    'Tipo do documento.
        gM46V999.gPrefixo = gM46V999.gPrefixo & " " & lNovNumPedido & ". "
        If gM46V111.gfGraRegComunicacao(lNovNosNumero, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
        Me.Hide
    End With
    Exit Sub
ControleErro:
    Screen.MousePointer = vbDefault
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V003A.mnuEdiCopiar_Click ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub mnuEdiExcluir_Click()
'Procedure: exclui documento selecionada (grava N no campo Cod_Ativo).


    Dim lMensagem As String     'Mensagem.
    Dim lTotalExcluir As Long           'Total de registros a serem excluidos
    Dim lRegExcluir As Variant      'Lista com os nosso_numeros para excluir.
    Dim lCount As Long    'Utilizado no for ... loop
    Dim objPedido As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED).
    Dim objstcPedido As stcA46V702B   'Objeto da classe de estrutura stcA46V702B (TAB_PED).
    Dim objTransm As clsA46V718A    'Objeto da classe clsA46V718A (TAB_PED_TRANSM).
    Dim objstcTransm As stcA46V718B   'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM).
    Dim lNumPedido As Long        'N�mero do pedido.
    Dim lNosNumero As String      'Nosso n�mero.

'1- Verifica se existe lista de arquivos para excluir
    On Error GoTo ControleErro

    If gM46V999.gfPreenchido(mRegChecked) = True Then

        'Separo na lista
        lRegExcluir = Split(mRegChecked, ",")

        'Verifica o tipo de documento selecionado
        If mTipDocumento <> 1 Then
            Call gM46V999.gpGraLog(1, "FGE0083 - Opera��o de exclus�o inv�lida para este tipo de documento.")
            'Limpa lista de arquivos selecionados para exclus�o
            Call mpAlteraTodos(False)
            Exit Sub
        End If

        'Confirma com usu�rio.
        lTotalExcluir = UBound(lRegExcluir)
        lMensagem = "AGE0023 - Confirma exclus�o de " & lTotalExcluir & _
                  " registro(s)?"
        If MsgBox(lMensagem, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Call gM46V999.gpGraLog(0, lMensagem & " N�O.")
            Exit Sub
        End If
        Call gM46V999.gpGraLog(0, lMensagem & " SIM.")

        For lCount = 0 To lTotalExcluir
            If gM46V999.gfPreenchido(lRegExcluir(lCount)) Then
                Call gM46V999.gpExcluiRegistros(CStr(lRegExcluir(lCount)), 0, "", "", lMensagem)
            End If
        Next lCount


        'Finaliza��o OK
        Call mpDesControles
        Call gM46V999.gpGraLog(1, "AGE0002 - Processo conclu�do: " & lTotalExcluir & _
                                " registro(s) exclu�do(s) com sucesso.")

    Else
        '2- Exclus�o de unico arquivo
        With flxResultado
            '1. Verifica tipo de documento e se tem documento selecionado.
            If optDocumento(0) = False And optDocumento(1) = False And optDocumento(2) = False Then
                Call gM46V999.gpGraLog(1, "FNP0013 - Tipo de documento n�o selecionado.")
                Exit Sub
            End If
            If mTipDocumento = 2 Then
                Call gM46V999.gpGraLog(1, "FGE0083 - Opera��o de exclus�o inv�lida para este tipo de documento.")
                Exit Sub
            End If
            If .Row = 0 Then
                Call gM46V999.gpGraLog(1, "FNP0014 - Documento n�o selecionado.")
                Exit Sub
            End If
            'If .CellBackColor <> vbGreen Then
            '    Call gM46V999.gpGraLog(1, "FNP0014 - Documento n�o selecionado.")
            '    Exit Sub
            'End If

            '2. Posiciona controles.
            Select Case mTipDocumento
            Case 1      'Cota��o ou proposta.
                lNumPedido = .TextMatrix(.Row, 2)       'N�mero do pedido - quando tipo de Doc = 1.
                lNosNumero = .TextMatrix(.Row, 7)      'Nosso n�mero.
            Case 3      'Protocolo.
                lNumPedido = .TextMatrix(.Row, 1)               'N�mero do pedido.
                lNosNumero = .TextMatrix(.Row, 13)       'Nosso n�mero.
            Case Else
                lNosNumero = String(20, "0")
            End Select

            '3. Confirma com usu�rio.
            lMensagem = "AGE0023 - Confirma exclus�o?" & vbLf & vbLf & "Documento " & lNumPedido & _
                      " com nosso n�mero " & lNosNumero & "?"
            If MsgBox(lMensagem, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                Call gM46V999.gpGraLog(0, lMensagem & " N�O.")
                Exit Sub
            End If
            Call gM46V999.gpGraLog(0, lMensagem & " SIM.")

            '4. Processa exclus�o.
            Select Case mTipDocumento
            Case 1      'Cota��o ou proposta.
                'a) Verifica registro de dados comuns.
                Set objPedido = New clsA46V702A
                Set objPedido.M46V999 = gM46V999

                Set objstcPedido = New stcA46V702B
                Set objstcPedido.M46V999 = gM46V999

                objstcPedido.NOSSO_NUMERO = lNosNumero
                If Not objPedido.mfCarregarColecao(objstcPedido, "", "", False) Then
                    Call gM46V999.gpGraLog(2, objPedido.UltimoErro)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                Set objstcPedido = objPedido.Colecao.Item(1)

                If UCase(objstcPedido.Cod_User_Liber) <> UCase(gM46V999.gNomUsuario6) Then
                    lMensagem = "AGE0023 - Confirma exclus�o?" & vbLf & vbLf & "Documento " & _
                                lNumPedido & " salvo anteriormente por " & objstcPedido.Cod_User_Liber & "."
                    If MsgBox(lMensagem, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                        Call gM46V999.gpGraLog(0, lMensagem & " N�O.")
                        Exit Sub
                    End If
                    Call gM46V999.gpGraLog(0, lMensagem & " SIM.")
                End If
                'b) Exclui registro de dados comuns.
                Call gM46V999.gpExcluiRegistros(lNosNumero, 0, "", "", lMensagem)
                Call gM46V999.gpGraLog(1, "AGE0002 - Processo conclu�do: documento " & lNumPedido & _
                                        " exclu�do com sucesso.")
            Case 3      'Protocolo.
                Set objTransm = New clsA46V718A
                Set objTransm.M46V999 = gM46V999

                Set objstcTransm = New stcA46V718B

                objstcTransm.NOSSO_NUMERO = lNosNumero
                objstcTransm.Num_Ped = lNumPedido

                If Not objTransm.mfExcluir(objstcTransm, "", "") Then
                    Call gM46V999.gpGraLog(1, objTransm.UltimoErro)
                    Exit Sub
                End If
                Call gM46V999.gpGraLog(1, "AGE0002 - Processo conclu�do: documento " & lNumPedido & _
                                        " exclu�do com sucesso.")
            End Select
            If .Rows = 2 Then
                Call mpDesControles
            Else
                .RemoveItem .Row
                If .Rows = 1 Then
                    Call mpDesControles
                End If
            End If
        End With
    End If
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V003A.mnuEdiExcluir_Click ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub mnuEdiLimpar_Click()
    Call mpDesControles
    If mPesqApolice = False Then
        optDocumento(0).Value = True
    End If
    txtNumPedido.Text = ""
    txtSegurado.Text = ""
    lblProdutor.Visible = True
    txtProdutor.Visible = True
    txtProdutor.Text = ""
    lblCodCorretor.Caption = ""
    txtCorretor.Text = ""
    flxResultado.Rows = 1
    txtIniVigencia.Text = ""
    txtTerVigencia.Text = ""
    If mPesqApolice = False Then
        If txtNumPedido.Visible = True And txtNumPedido.Enabled = True Then
            txtNumPedido.SetFocus
        End If
    End If
End Sub
Private Sub mnuExiProAntigos_Click()
    If flxResultado.Rows > 1 Then
        gM46V111.gNumPedido = flxResultado.TextMatrix(flxResultado.RowSel, 1)
        gM46V111.gComNosNumero = flxResultado.TextMatrix(flxResultado.RowSel, 13)
        frmT46V006A.Show 1
    End If
End Sub
Private Sub mnuSair_Click()
    Unload Me
End Sub
Private Sub optDocumento_Click(index As Integer)

'Toda vez que o tipo de documento for alterado
'Desabilita os controles
    Call mpDesControles

    Select Case index
    Case 0      'Cota��es e propostas.
        mTipDocumento = 1
        lblNumPedido.Visible = True
        txtNumPedido.Visible = True
        lblProdutor.Visible = True
        txtProdutor.Visible = True
        lblCorretor.Visible = True
        txtCorretor.Visible = True
        fraSitDocumento.Enabled = True
        optSitDocumento(0).Enabled = True
        optSitDocumento(1).Enabled = True
        optSitDocumento(2).Enabled = True
        '###PROJETO INSPECAO 10018
        'DESCOMENTAR LINHAS
        optSitDocumento(3).Enabled = True

        optSitDocumento(0) = True
    Case 1      'Ap�lices vigentes.
        mTipDocumento = 2
        lblNumPedido.Visible = False
        txtNumPedido.Visible = False
        lblProdutor.Visible = True
        txtProdutor.Visible = True
        lblCorretor.Visible = True
        txtCorretor.Visible = True
        fraSitDocumento.Enabled = False
        optSitDocumento(0).Value = False
        optSitDocumento(0).Enabled = False
        optSitDocumento(1).Value = False
        optSitDocumento(1).Enabled = False
        optSitDocumento(2).Value = False
        optSitDocumento(2).Enabled = False
        optSitDocumento(3).Value = False
        optSitDocumento(3).Enabled = False
    Case 2      'Protocolo de transmiss�o.
        mTipDocumento = 3
        lblNumPedido.Visible = True
        txtNumPedido.Visible = True
        lblProdutor.Visible = False
        txtProdutor.Visible = False
        lblCorretor.Visible = False
        txtCorretor.Visible = False
        fraSitDocumento.Enabled = False
        optSitDocumento(0).Value = False
        optSitDocumento(0).Enabled = False
        optSitDocumento(1).Value = False
        optSitDocumento(1).Enabled = False
        optSitDocumento(2).Value = False
        optSitDocumento(2).Enabled = False
        optSitDocumento(3).Value = False
        optSitDocumento(3).Enabled = False
    End Select
End Sub
Private Sub optDocumento_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optSitDocumento_Click(index As Integer)
    If flxResultado.Rows <> 1 Then
        Call mpDesControles
    End If
End Sub
Private Sub txtCorretor_Change()
    If lCarga = False Then
        lblCodCorretor.Caption = ""
    End If
End Sub
Private Sub txtCorretor_GotFocus()
    With txtCorretor
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtCorretor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtCorretor_LostFocus()
    Dim lCorretor As String    'Corretor.
    Dim lI As Double             'Utilizado com For...Next.
    Dim lMensagem As String    'Mensagem.
    Dim lQuaCorretores As Double  'Quantidade de corretores selecionados.
    Dim lSaida() As Variant    'Resultado da busca de corretores.

    On Error GoTo ControleErro

    '1. Se c�digo do corretor j� preenchido, nada a fazer.
    If gM46V999.gfPreenchido(lblCodCorretor.Caption) Then
        Exit Sub
    End If
    lblCodCorretor.Caption = "000000"

    '2. Posiciona controle.
    Screen.MousePointer = vbHourglass

    '3. Se n�o preenchido, encerra processo.
    If Not gM46V999.gfPreenchido(txtCorretor.Text) Then
        lblCodCorretor.Caption = ""
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '4. Posiciona campo com o que foi informado pelo usu�rio.
    lCorretor = Trim$(txtCorretor.Text)

    '5. Se foram digitadas at� 6 posi��es num�ricas, trata de c�digo de corretor.
    '   Obt�m registro e posiciona dados.
    If IsNumeric(lCorretor) And Len(lCorretor) <= 6 Then
        With lstAviso
            .Clear
            .AddItem "Aguarde um momento... Estamos buscando dados do corretor " & lCorretor & "..."
            .ZOrder 0
            .Visible = True
            .Refresh
            If gM46V999.gfObtCorretor(1, lCorretor, lSaida, lMensagem) = True Then
                lblCodCorretor.Caption = lSaida(1, 1)
                lCarga = True
                txtCorretor.Text = lSaida(1, 3)
                lCarga = False
            Else
                Call gM46V999.gpGraLog(1, lMensagem)
                lblCodCorretor.Caption = ""
                txtCorretor.Text = ""
                If txtCorretor.Visible = True And txtCorretor.Enabled = True Then
                    txtCorretor.SetFocus
                End If
            End If
            .Visible = False
        End With
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '6. Se preenchido com menos de 3 posi��es, encerra processo.
    If Len(lCorretor) < 3 Then
        Call gM46V999.gpGraLog(1, "FGE0075 - Corretor inv�lido (digitar pelo menos 3 posi��es).")
        lblCodCorretor.Caption = ""
        txtCorretor.Text = ""
        If txtCorretor.Visible = True And txtCorretor.Enabled = True Then
            txtCorretor.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '7. Obt�m registros com informa��o do usu�rio.
    'a) Emite aviso de espera.
    With lstAviso
        .Clear
        .AddItem "Aguarde um momento... Estamos buscando dados de corretor que possuem " & lCorretor & _
                 "..."
        .ZOrder 0
        .Visible = True
        .Refresh
    End With
    'b) Obt�m lista de corretores
    If gM46V999.gfObtCorretor(2, lCorretor, lSaida, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        lstAviso.Visible = False
        lblCodCorretor.Caption = ""
        txtCorretor.Text = ""
        If txtCorretor.Visible = True And txtCorretor.Enabled = True Then
            txtCorretor.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'c) Se campo de sa�da estiver vazio, nada a fazer.
    If IsNull(lSaida) Then
        lstAviso.Visible = False
        lstAviso.Refresh
        lblCodCorretor.Caption = ""
        txtCorretor.Text = ""
        If txtCorretor.Visible = True And txtCorretor.Enabled = True Then
            txtCorretor.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'd) Posiciona quantidade de corretores encontrados.
    lQuaCorretores = UBound(lSaida)
    'e) Se encontrou apenas 1 corretor, posiciona dados deste corretor e encerra processo.
    If lQuaCorretores = 1 Then
        lblCodCorretor.Caption = lSaida(1, 1)
        lCarga = True
        txtCorretor.Text = lSaida(1, 3)
        lCarga = False
        lstAviso.Visible = False
        lstAviso.Refresh
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'f) Se encontrou mais de um corretor, abre janela com grade de corretores.
    Load frmT46V004A
    With frmT46V004A.flxConsulta
        .Rows = lQuaCorretores + 1
        For lI = 1 To lQuaCorretores
            .TextMatrix(lI, 0) = lSaida(lI, 1)  'C�digo do corretor.
            .TextMatrix(lI, 1) = lSaida(lI, 2)  'Nome do corretor.
            .TextMatrix(lI, 2) = lSaida(lI, 3)  'Nome abreviado do corretor.
        Next lI
        Screen.MousePointer = vbDefault
        lCarga = True
        Call frmT46V004A.mpInstanciarForm(1)
        lCarga = False
        lstAviso.Visible = False
        lstAviso.Refresh
        If Not gM46V999.gfPreenchido(txtCorretor.Text) Then
            lblCodCorretor.Caption = ""
            If txtCorretor.Visible = True And txtCorretor.Enabled = True Then
                txtCorretor.SetFocus
            End If
        End If
    End With

    Screen.MousePointer = vbDefault
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V003A.txtCorretor_LostFocus ERRO - " & Err & " " & Err.Description)
End Sub

Private Sub txtIniVigencia_GotFocus()
    With txtIniVigencia
        .SelStart = 0
        .SelLength = 10
    End With
End Sub
Private Sub txtIniVigencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtIniVigencia_LostFocus()
    txtIniVigencia.Text = gM46V999.gfForData(9, txtIniVigencia.Text)
End Sub

Private Sub txtNumPedido_GotFocus()
    With txtNumPedido
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtNumPedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'KeyAscii = 0
        'SendKeys "{TAB}"
        mnuArqPesquisar_Click
    End If
End Sub
Private Sub txtProdutor_GotFocus()
    With txtProdutor
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtProdutor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtProdutor_LostFocus()
    Dim bdP0044700 As ADODB.Connection    'P0044700: controle de documentos.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lMensagem As String         'Mensagem.
    Dim lProdutor As String         'Produtor.
    Dim rsPRODUTOR As clsYasRecordSet    'Registros de produtores.
    Dim lSelect As String
    lCarga = False
    On Error GoTo ControleErro

    '1. Se c�digo do produtor j� preenchido, nada a fazer.
    If gM46V999.gfPreenchido(lblCodProdutor.Caption) Then
        Exit Sub
    End If

    '2. Se produtor n�o preenchido, encerra processo.
    If Not gM46V999.gfPreenchido(txtProdutor.Text) Then
        lblCodProdutor.Caption = ""
        Exit Sub
    End If

    '3. Posiciona controle.
    Screen.MousePointer = vbHourglass

    '4. Posiciona campo com o que foi informado pelo usu�rio.
    lProdutor = Trim$(txtProdutor.Text)

    '5. Se produtor � num�rico e tamanho superior a 4, emite mensagem e encerra processo.
    If IsNumeric(lProdutor) Then
        If Len(lProdutor) > 4 Then
            txtProdutor.Text = ""
            Call gM46V999.gpGraLog(1, "FGE0084 - Produtor inv�lido (c�digo do produtor com mais de 4 posi��es).")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    '6. Emite aviso.
    With lstAviso
        .Clear
        .AddItem "Aguarde um momento... Estamos buscando dados do produtor " & lProdutor & "..."
        .ZOrder 0
        .Visible = True
        .Refresh
    End With

    '7. Obt�m registros de produtores.
    If gM46V999.gfAbrBasDados("P0044700", gM46V999.gTipP0044700, gM46V999.gEndP0044700, bdP0044700, lMensagem) = False Then
        txtProdutor.Text = ""
        Call gM46V999.gpGraLog(2, lMensagem)
        lstAviso.Visible = False
        lstAviso.Refresh
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If IsNumeric(lProdutor) Then
        'Se produtor num�rico, trata-se do c�digo do produtor.
        lSelect = "SELECT COD_PROD, NOM_PROD FROM YTAB_PRODUTOR WHERE COD_PROD = " & lProdutor
        If gM46V999.gfObtRegistro(bdP0044700, lSelect, rsPRODUTOR, lMensagem) = False Then
            txtProdutor.Text = ""
            Call gM46V999.gpFecha2(bdP0044700)
            Call gM46V999.gpGraLog(2, lMensagem)
            lstAviso.Visible = False
            lstAviso.Refresh
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If rsPRODUTOR.EOF = True Then
            txtProdutor.Text = ""
            Call gM46V999.gpFecha2(bdP0044700)
            Call gM46V999.gpGraLog(1, "FGE0085 - Produtor inv�lido (c�digo de produtor n�o localizado: " & _
                                      lProdutor & ").")
            lstAviso.Visible = False
            lstAviso.Refresh
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        'Se produtor n�o num�rico, trata-se do nome do produtor.
        lSelect = "SELECT COD_PROD, NOM_PROD FROM YTAB_PRODUTOR WHERE NOM_PROD LIKE '%" & _
                  gM46V999.gfForAspas(UCase(lProdutor)) & "%'"
        If gM46V999.gfObtRegistro(bdP0044700, lSelect, rsPRODUTOR, lMensagem) = False Then
            txtProdutor.Text = ""
            Call gM46V999.gpFecha2(bdP0044700)
            Call gM46V999.gpGraLog(2, lMensagem)
            lstAviso.Visible = False
            lstAviso.Refresh
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If rsPRODUTOR.EOF = True Then
            txtProdutor.Text = ""
            Call gM46V999.gpFecha2(bdP0044700)
            Call gM46V999.gpGraLog(1, "FGE0086 - " & gM46V999.gPrefixo & _
                                      "Produtor inv�lido (n�o localizado produtor com nome " & _
                                      Trim$(lProdutor) & ").")
            lstAviso.Visible = False
            lstAviso.Refresh
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    '8. Se encontrou apenas um produtor, posiciona dados deste produtor e encerra processo.
    rsPRODUTOR.MoveLast
    If rsPRODUTOR.RecordCount = 1 Then
        lblCodProdutor.Caption = rsPRODUTOR!Cod_Prod
        lCarga = True
        txtProdutor.Text = rsPRODUTOR!NOM_PROD
        lCarga = False
        Call gM46V999.gpFecha2(bdP0044700)
        lstAviso.Visible = False
        lstAviso.Refresh
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '9. Se encontrou mais de um produtor, abre janela com grade de produtores.
    Load frmT46V004A
    With frmT46V004A.flxConsulta
        .Rows = rsPRODUTOR.RecordCount + 1
        rsPRODUTOR.MoveFirst
        lI = 0
        While Not rsPRODUTOR.EOF
            lI = lI + 1
            .TextMatrix(lI, 0) = rsPRODUTOR!Cod_Prod    'C�digo do produtor.
            .TextMatrix(lI, 1) = rsPRODUTOR!NOM_PROD    'Nome do produtor.
            rsPRODUTOR.MoveNext
        Wend
        Screen.MousePointer = vbDefault
        lCarga = True
        Call frmT46V004A.mpInstanciarForm(2)
        lCarga = False
        lstAviso.Visible = False
        lstAviso.Refresh
        If Not gM46V999.gfPreenchido(txtProdutor.Text) Then
            lblCodProdutor.Caption = ""
            If txtProdutor.Visible = True And txtProdutor.Enabled = True Then
                txtProdutor.SetFocus
            End If
        End If
    End With
    Call gM46V999.gpFecha2(bdP0044700)

    Screen.MousePointer = vbDefault
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V003A.txtProdutor_LostFocus ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub txtSegurado_GotFocus()
    With txtSegurado
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtSegurado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtTerVigencia_GotFocus()
    With txtTerVigencia
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtTerVigencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtTerVigencia_LostFocus()
    txtTerVigencia.Text = gM46V999.gfForData(9, txtTerVigencia.Text)
End Sub
Public Sub mpInstanciarForm(ByRef pNum_Apolice As String, ByRef pNum_Item As String, Optional ByRef pCod_Ramo As Integer, Optional ByRef pCod_Modalidade As Integer)
'Inicio.
    mPesqApolice = True
    mNum_Apolice = ""
    optDocumento(1).Value = True
    fraDocumento.Enabled = False

    Me.Show vbModal
    'Retorno
    pNum_Apolice = mNum_Apolice
    pNum_Item = mNum_Item
    pCod_Ramo = mCod_Ramo
    pCod_Modalidade = mCod_Modalidade

    mPesqApolice = False
End Sub
Private Sub mpAlteraTodos(ByVal pFlag As Boolean)
'Procedure utilizada para selecinar ou desmarar arquivos para exclu��o

'Par�metro de entrada...pFlag...true = selecionar todos e false = desmarcar todos.

    Dim lColuna As Integer    'Coluna.
    Dim lI As Integer    'Contador do FOR.
    Dim lLinha As Integer    'Linha.
    Dim lSelecao As String    'Sele��o.

'1. Limpa vari�vel.
    mRegChecked = Empty

    '2. Marca ou desmarca todos os registro na grade.
    With flxResultado
        lColuna = .Col
        lLinha = .Row
        If pFlag = True Then        'Marca todos os registros exibidos na grade.
            For lI = 1 To (.Rows - 1)
                .Row = lI
                .Col = 0
                'Seleciona o registro.
                Set flxResultado.CellPicture = picChecked
                .Col = .Col + 7    'Coluna de nosso n�mero.
                'Armazena na vari�vel
                lSelecao = .Text
                mRegChecked = mRegChecked & lSelecao & ","
            Next lI
        Else                        'Desmarca todos os registor exibidos na grade.
            For lI = 1 To (.Rows - 1)
                .Row = lI
                .Col = 0
                Set flxResultado.CellPicture = picUnchecked
            Next lI
        End If
        .Col = lColuna
        .Row = lLinha
    End With
End Sub
