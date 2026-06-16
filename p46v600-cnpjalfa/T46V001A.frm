VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmT46V001A 
   BackColor       =   &H80000004&
   Caption         =   "SYAS - Sistema Yasuda Seguros - Multisseguros"
   ClientHeight    =   7935
   ClientLeft      =   1485
   ClientTop       =   1530
   ClientWidth     =   11280
   Icon            =   "T46V001A.frx":0000
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7935
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.Frame fraEmissao 
      Caption         =   "Funções do sistema"
      ForeColor       =   &H00800000&
      Height          =   7440
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   2800
      Begin MSComDlg.CommonDialog cdgHelp 
         Left            =   2280
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSCommand cmdSegNovo 
         Height          =   550
         Left            =   100
         TabIndex        =   1
         ToolTipText     =   "Criar cotação de nova apólice."
         Top             =   250
         Width           =   550
         _Version        =   65536
         _ExtentX        =   970
         _ExtentY        =   970
         _StockProps     =   78
         MousePointer    =   99
         MouseIcon       =   "T46V001A.frx":030A
         Picture         =   "T46V001A.frx":0624
      End
      Begin Threed.SSCommand cmdPesquisa 
         Height          =   550
         Left            =   100
         TabIndex        =   5
         ToolTipText     =   "Pesquisar cotações, propostas, apólices e protocolo."
         Top             =   1410
         Width           =   550
         _Version        =   65536
         _ExtentX        =   970
         _ExtentY        =   970
         _StockProps     =   78
         MousePointer    =   99
         MouseIcon       =   "T46V001A.frx":093E
         Picture         =   "T46V001A.frx":0C58
      End
      Begin Threed.SSCommand cmdSair 
         Cancel          =   -1  'True
         Height          =   550
         Left            =   100
         TabIndex        =   13
         ToolTipText     =   "Sair do sistema."
         Top             =   6660
         Width           =   555
         _Version        =   65536
         _ExtentX        =   970
         _ExtentY        =   970
         _StockProps     =   78
         MousePointer    =   99
         MouseIcon       =   "T46V001A.frx":0F72
         Picture         =   "T46V001A.frx":128C
      End
      Begin Threed.SSCommand cmdRenovacao 
         Height          =   550
         Left            =   100
         TabIndex        =   3
         ToolTipText     =   "Criar cotação de renovação de apólice."
         Top             =   830
         Width           =   550
         _Version        =   65536
         _ExtentX        =   970
         _ExtentY        =   970
         _StockProps     =   78
         MousePointer    =   99
         MouseIcon       =   "T46V001A.frx":15A6
         Picture         =   "T46V001A.frx":18C0
      End
      Begin Threed.SSCommand cmdTransmitir 
         Height          =   550
         Left            =   100
         TabIndex        =   7
         ToolTipText     =   "Transmitir propostas."
         Top             =   2570
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   979
         _StockProps     =   78
         MousePointer    =   99
         MouseIcon       =   "T46V001A.frx":1BDA
         Picture         =   "T46V001A.frx":1EF4
      End
      Begin Threed.SSCommand cmdAjuda 
         Height          =   550
         Left            =   105
         TabIndex        =   11
         ToolTipText     =   "Acionar ajuda do sistema."
         Top             =   6075
         Width           =   555
         _Version        =   65536
         _ExtentX        =   970
         _ExtentY        =   970
         _StockProps     =   78
         MousePointer    =   99
         MouseIcon       =   "T46V001A.frx":2346
         Picture         =   "T46V001A.frx":2660
      End
      Begin Threed.SSCommand cmdManCorretor 
         Height          =   550
         Left            =   100
         TabIndex        =   9
         ToolTipText     =   "Acionar manual do corretor."
         Top             =   5505
         Width           =   555
         _Version        =   65536
         _ExtentX        =   970
         _ExtentY        =   970
         _StockProps     =   78
         MousePointer    =   99
         MouseIcon       =   "T46V001A.frx":297A
         Picture         =   "T46V001A.frx":2C94
      End
      Begin MSComctlLib.ImageList imgLista 
         Left            =   1700
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "T46V001A.frx":2FAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "T46V001A.frx":3400
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "T46V001A.frx":371A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin Threed.SSCommand cmdEndosso 
         Height          =   555
         Left            =   105
         TabIndex        =   24
         ToolTipText     =   "Criar cotação de endosso de apólice."
         Top             =   1980
         Width           =   555
         _Version        =   65536
         _ExtentX        =   970
         _ExtentY        =   970
         _StockProps     =   78
         MousePointer    =   99
         MouseIcon       =   "T46V001A.frx":3B6C
         Picture         =   "T46V001A.frx":3E86
      End
      Begin Threed.SSCommand cmdPerfil 
         Height          =   555
         Left            =   90
         TabIndex        =   25
         ToolTipText     =   "Alterar Propostas."
         Top             =   4905
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   979
         _StockProps     =   78
         MousePointer    =   99
         MouseIcon       =   "T46V001A.frx":42D8
         Picture         =   "T46V001A.frx":45F2
      End
      Begin Threed.SSCommand cmdImpressao 
         Height          =   555
         Left            =   90
         TabIndex        =   27
         ToolTipText     =   "Impressão das apólices."
         Top             =   4305
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   979
         _StockProps     =   78
         MousePointer    =   99
         MouseIcon       =   "T46V001A.frx":4A44
         Picture         =   "T46V001A.frx":4D5E
      End
      Begin Threed.SSCommand cmdAgendaInspecao 
         Height          =   555
         Left            =   105
         TabIndex        =   29
         ToolTipText     =   "Solicitar agendamento de inspeção das cotações."
         Top             =   3150
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   979
         _StockProps     =   78
         MousePointer    =   99
         MouseIcon       =   "T46V001A.frx":51B0
         Picture         =   "T46V001A.frx":54CA
      End
      Begin VB.Label lblAgendaInspecao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Solicitar Inspeção"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   690
         MouseIcon       =   "T46V001A.frx":57E4
         MousePointer    =   99  'Custom
         TabIndex        =   30
         ToolTipText     =   "Solicitar agendamento de inspeção das cotações."
         Top             =   3325
         Width           =   2055
      End
      Begin VB.Label lblImpressao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Impressão de apólice"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   690
         MouseIcon       =   "T46V001A.frx":5AEE
         MousePointer    =   99  'Custom
         TabIndex        =   28
         ToolTipText     =   "Fechar a fatura."
         Top             =   4485
         Width           =   2055
      End
      Begin VB.Label LblPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Alterar Perfil"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   690
         MouseIcon       =   "T46V001A.frx":5DF8
         MousePointer    =   99  'Custom
         TabIndex        =   26
         ToolTipText     =   "Transmitir propostas."
         Top             =   5105
         Width           =   2055
      End
      Begin VB.Label lblEndosso 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Endosso"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   700
         MouseIcon       =   "T46V001A.frx":6102
         MousePointer    =   99  'Custom
         TabIndex        =   23
         ToolTipText     =   "Criar cotação de endosso de apólice."
         Top             =   2185
         Width           =   2055
      End
      Begin VB.Label lblManCorretor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Manual do corretor"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   705
         MouseIcon       =   "T46V001A.frx":6544
         MousePointer    =   99  'Custom
         TabIndex        =   10
         ToolTipText     =   "Acionar manual do corretor."
         Top             =   5705
         Width           =   2055
      End
      Begin VB.Label lblSair 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Sair"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   705
         MouseIcon       =   "T46V001A.frx":684E
         MousePointer    =   99  'Custom
         TabIndex        =   14
         ToolTipText     =   "Sair do sistema."
         Top             =   6855
         Width           =   2055
      End
      Begin VB.Label lblAjuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Ajuda do SYAS"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   705
         MouseIcon       =   "T46V001A.frx":6B58
         MousePointer    =   99  'Custom
         TabIndex        =   12
         ToolTipText     =   "Acionar ajuda do sistema."
         Top             =   6270
         Width           =   2055
      End
      Begin VB.Label lblTransmitir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Transmitir propostas"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   705
         MouseIcon       =   "T46V001A.frx":6E62
         MousePointer    =   99  'Custom
         TabIndex        =   8
         ToolTipText     =   "Transmitir propostas."
         Top             =   2765
         Width           =   2055
      End
      Begin VB.Label lblRenovacao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Renovação"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   700
         MouseIcon       =   "T46V001A.frx":716C
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "Criar cotação de renovação de apólice."
         Top             =   1025
         Width           =   2050
      End
      Begin VB.Label lblPesquisa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Pesquisa"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   700
         MouseIcon       =   "T46V001A.frx":7476
         MousePointer    =   99  'Custom
         TabIndex        =   6
         ToolTipText     =   "Pesquisar cotações, propostas, apólices e protocolo."
         Top             =   1605
         Width           =   2050
      End
      Begin VB.Label lblSegNovo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Seguro novo"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   700
         MouseIcon       =   "T46V001A.frx":7780
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "Criar cotação de nova apólice."
         Top             =   445
         Width           =   2050
      End
   End
   Begin Threed.SSPanel panAviso 
      Height          =   1000
      Left            =   3825
      TabIndex        =   15
      Top             =   2940
      Visible         =   0   'False
      Width           =   4000
      _Version        =   65536
      _ExtentX        =   7056
      _ExtentY        =   1764
      _StockProps     =   15
      Caption         =   "Aguarde...   Atualizando dados do sistema..."
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.21
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Outline         =   -1  'True
   End
   Begin Threed.SSPanel panRodape 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   7515
      Width           =   11280
      _Version        =   65536
      _ExtentX        =   19897
      _ExtentY        =   741
      _StockProps     =   15
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin Threed.SSPanel panForm 
         Height          =   300
         Left            =   60
         TabIndex        =   17
         ToolTipText     =   "Nome do programa"
         Top             =   60
         Width           =   1000
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "form"
         ForeColor       =   8388608
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         FloodColor      =   8388608
      End
      Begin Threed.SSPanel panUsuario 
         Height          =   300
         Left            =   8000
         TabIndex        =   21
         ToolTipText     =   "Código do usuário."
         Top             =   60
         Width           =   1600
         _Version        =   65536
         _ExtentX        =   2822
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Usuário XXXXXXXX"
         ForeColor       =   8388608
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         FloodColor      =   8388608
      End
      Begin Threed.SSPanel panVersao 
         Height          =   300
         Left            =   4800
         TabIndex        =   19
         ToolTipText     =   "Versão do sistema."
         Top             =   60
         Width           =   1600
         _Version        =   65536
         _ExtentX        =   2822
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Versão dd/mm/aaaa"
         ForeColor       =   8388608
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         FloodColor      =   8388608
      End
      Begin Threed.SSPanel panData 
         Height          =   300
         Left            =   9600
         Negotiate       =   -1  'True
         TabIndex        =   22
         ToolTipText     =   "Data do dia de hoje."
         Top             =   60
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "dd/mm/aaaa"
         ForeColor       =   8388608
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         FloodColor      =   8388608
      End
      Begin Threed.SSPanel panBasDados 
         Height          =   300
         Left            =   1100
         TabIndex        =   18
         ToolTipText     =   "Localização das bases de dados."
         Top             =   60
         Width           =   3700
         _Version        =   65536
         _ExtentX        =   6526
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ForeColor       =   8388608
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         FloodColor      =   8388608
      End
      Begin Threed.SSPanel panCorretor 
         Height          =   300
         Left            =   6400
         TabIndex        =   20
         ToolTipText     =   "Código do corretor."
         Top             =   60
         Width           =   1600
         _Version        =   65536
         _ExtentX        =   2822
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Corretor XXXXXXX"
         ForeColor       =   8388608
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         FloodColor      =   8388608
      End
   End
   Begin VB.Image imgYasuda 
      Height          =   7455
      Left            =   2880
      Top             =   0
      Width           =   8895
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuArqEmissao 
         Caption         =   "&SYAS - Emissão"
      End
      Begin VB.Menu mnuSepSegNovo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArqSegNovo 
         Caption         =   "Segur&o novo"
      End
      Begin VB.Menu mnuArqRenovacao 
         Caption         =   "Renov&ação de seguro"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuArqEndosso 
         Caption         =   "&Endosso"
      End
      Begin VB.Menu mnuSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArqPesquisa 
         Caption         =   "&Pesquisa cotações, propostas, apólices e protocolo"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArqTransmissao 
         Caption         =   "&Transmissão de propostas"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuAgendarInspecao 
         Caption         =   "&Agendar Inspeção"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuArqManCorretor 
         Caption         =   "&Manual do corretor"
      End
      Begin VB.Menu mnuSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sai&r"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuFerramentas 
      Caption         =   "&Ferramentas"
      Begin VB.Menu mnuFerLocDados 
         Caption         =   "&Localização dos dados"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFerImpressao 
         Caption         =   "Impressão de apólice/endosso"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFerCadCCorretagem 
         Caption         =   "Cadastro de co-corretagem"
      End
      Begin VB.Menu mnuFerRegPadrao 
         Caption         =   "Registro padrão"
      End
      Begin VB.Menu mnuSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFerImpressaoTransmissao 
         Caption         =   "Impressão de Transmissão"
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "Aj&uda"
      Begin VB.Menu mnuAjuSYAS 
         Caption         =   "Aj&uda do SYAS"
      End
      Begin VB.Menu mnuSep08 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAjuSobre 
         Caption         =   "So&bre o SYAS"
      End
      Begin VB.Menu mnuManutencao 
         Caption         =   "Manutenção"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmT46V001A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO.
'
'No evento Form_Load é necessário colocar a data da versão do módulo para garantir que o usuário esteja
'com o módulo mais atualizado.
'
'ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private mErrCarga As Boolean    'Indicador de erro na carga do form.
Private mFezConexao As Boolean  'Indicador de que fêz conexão.
Private mPreConexao As Boolean  'Indicador de que precisa fazer conexão.

Private Sub mpDesabilita()
'Procedure: desabilita botões (deixa apenas o botão de conexão habilitado) e posiciona controle.
    On Error GoTo ControleErro
    '1. Desabilita botões.
    mnuArquivo.Enabled = False
    mnuFerramentas.Enabled = False
    'mnuAjuda.Enabled = False
    cmdSegNovo.Visible = False
    lblSegNovo.Visible = False
    cmdRenovacao.Visible = False
    lblRenovacao.Visible = False
    cmdPesquisa.Visible = False
    lblPesquisa.Visible = False
    cmdTransmitir.Visible = False
    lblTransmitir.Visible = False
    cmdManCorretor.Visible = False
    lblManCorretor.Visible = False
    cmdEndosso.Visible = False
    lblEndosso.Visible = (cmdEndosso.Visible)

    cmdImpressao.Visible = False
    lblImpressao.Visible = False

    '###PROJETO INSPECAO 10018
    'DESCOMENTAR AS LINHAS ABAIXO
    cmdAgendaInspecao.Visible = False
    lblAgendaInspecao.Visible = False

    '2. Posiciona controles.
    mPreConexao = True      'Precisa conexão?
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mpDesabilita ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub mpLimBasDados()
'Procedure: limpeza das bases de dados.

    Dim bdP00MPED As ADODB.Connection    'Base de dados: P00MPED.
    Dim lDatLimite As Double          'Data limite.
    Dim lMensagem As String         'Mensagem.
    Dim rsTab_Ped As clsYasRecordSet    'Registro - Tab_Ped.
    Dim lSelect As String
    Dim lDelete As String

    On Error GoTo ControleErro

    '1. Prepara quadro de aviso.
    panAviso.Caption = "Aguarde..." & vbLf & vbLf & "Processando limpeza nos arquivos..."
    panAviso.Refresh

    '2. Abre base de dados.
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdP00MPED, lMensagem) = False Then
        Call gM46V999.gpGraLog(0, lMensagem)
        panAviso.Visible = False
        panAviso.Refresh
        Exit Sub
    End If

    '3. Obtém registros de dados comuns das registros a serem excluídos.
    If Month(Now) <= 3 Then
        lDatLimite = CDbl(Format$(Now, "yyyymmdd")) - 10000 + 900    'Subtrai 1 do ano e soma 9 no mês.
    Else
        lDatLimite = CDbl(Format$(Now, "yyyymmdd")) - 300            'Subtrai 2 do mês.
    End If
    lDatLimite = lDatLimite * 1000000
    lSelect = "SELECT * FROM Tab_Ped WHERE Dat_Digit < " & lDatLimite
    If gM46V999.gfObtRegistro(bdP00MPED, lSelect, rsTab_Ped, lMensagem) = False Then
        Call gM46V999.gpFecha2(bdP00MPED)
        Call gM46V999.gpGraLog(0, lMensagem)
        panAviso.Visible = False
        panAviso.Refresh
        Exit Sub
    End If
    If rsTab_Ped.EOF = True Then
        'Não tem registro anterior a 2 meses.
        Call gM46V999.gpFecha2(bdP00MPED)
        panAviso.Visible = False
        panAviso.Refresh
        Exit Sub
    End If

    '4. Processa limpeza.
    Do While Not rsTab_Ped.EOF
        'a) Posiciona mensagem de exclusão.
        lMensagem = "AGE0030 - Exclusão realizada com sucesso." & vbLf & vbLf & "Documento " & _
                    rsTab_Ped!Num_Ped & " com nosso número " & rsTab_Ped!NOSSO_NUMERO & _
                  " e data da última atualização " & gM46V999.gfForData(2, rsTab_Ped!DAT_DIGIT) & "."
        Call gM46V999.gpGraLog(0, lMensagem)
        panAviso.Caption = lMensagem
        panAviso.Refresh
        'b) Exclui registros das tabelas.
        lDelete = "DELETE FROM Tab_Ped_Clau WHERE Nosso_Numero = '" & rsTab_Ped!NOSSO_NUMERO & "'"
        If gM46V999.gfExeSQL(bdP00MPED, lDelete, lMensagem) <> 0 Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Do
        End If
        lDelete = "DELETE FROM Tab_Ped_Cobert WHERE Nosso_Numero = '" & rsTab_Ped!NOSSO_NUMERO & "'"
        If gM46V999.gfExeSQL(bdP00MPED, lDelete, lMensagem) <> 0 Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Do
        End If
        lDelete = "DELETE FROM Tab_Ped_Corr WHERE Nosso_Numero = '" & rsTab_Ped!NOSSO_NUMERO & "'"
        If gM46V999.gfExeSQL(bdP00MPED, lDelete, lMensagem) <> 0 Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Do
        End If
        lDelete = "DELETE FROM Tab_Ped_Ender WHERE Nosso_Numero = '" & rsTab_Ped!NOSSO_NUMERO & "'"
        If gM46V999.gfExeSQL(bdP00MPED, lDelete, lMensagem) <> 0 Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Do
        End If
        lDelete = "DELETE FROM Tab_Ped_Auto WHERE Nosso_Numero = '" & rsTab_Ped!NOSSO_NUMERO & "'"
        If gM46V999.gfExeSQL(bdP00MPED, lDelete, lMensagem) <> 0 Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Do
        End If
        lDelete = "DELETE FROM Tab_Ped_Loc WHERE Nosso_Numero = '" & rsTab_Ped!NOSSO_NUMERO & "'"
        If gM46V999.gfExeSQL(bdP00MPED, lDelete, lMensagem) <> 0 Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Do
        End If
        'c) Exclui registro da tabela de pedidos.
        lDelete = "DELETE FROM Tab_Ped WHERE Nosso_Numero = '" & rsTab_Ped!NOSSO_NUMERO & "'"
        If gM46V999.gfExeSQL(bdP00MPED, lDelete, lMensagem) <> 0 Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Do
        End If
        'd) Posiciona próximo registro.
        rsTab_Ped.MoveNext
    Loop

    '6. Encerra processo.
    Call gM46V999.gpFecha2(bdP00MPED)
    Call gM46V999.gpGraLog(0, "AGE0002 - Processo concluído: limpeza das bases de dados.")
    panAviso.Visible = False
    panAviso.Refresh
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mpLimBasDados ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub mpLimHistorico()
'Procedure: limpa arquivo histórico.

    Dim lAuxData As String    'Auxiliar: data.
    Dim lDatLimite As Double    'Data limite.
    Dim lEndSYASMULTLOG As String   'Endereço do arquivo SYASMULT.LOG.
    Dim lEndSYASMULTTMP As String   'Endereço do arquivo SYASMULT.TMP.
    Dim lNumSYASMULTLOG As Integer  'Utilizado com FreeFile para o arquivo SYASMULT.LOG.
    Dim lNumSYASMULTTMP As Integer  'Utilizado com FreeFile para o arquivo SYASMULT.TMP.
    Dim lRegistro As String    'Registro.
    Dim lErro As Integer

    On Error GoTo ControleErro
    lErro = 1
    '1. Posiciona endereços dos arquivos histórico e temporário.
    lEndSYASMULTLOG = gM46V999.gAppPath & "SYASMULT_" & Format(Date, "DD_MM_YYYY") & ".LOG"
    lEndSYASMULTTMP = gM46V999.gAppPath & "SYASMULT.TMP"
    lErro = 2
    '2. Abre arquivo histórico.
    If Not gM46V999.gfPreenchido(Dir(lEndSYASMULTLOG)) Then
        Call gM46V999.gpGraLog(0, "LIMPANDO O HISTORICO")
    End If
    lErro = 3
    lNumSYASMULTLOG = FreeFile
    On Error Resume Next
    Open lEndSYASMULTLOG For Input Shared As #lNumSYASMULTLOG
    If Err <> 0 Then
        Call gM46V999.gpGraLog(2, "FTE0076 - Erro ao abrir arquivo " & lEndSYASMULTLOG & ": " & Err & " - " & _
                                  Error & vbLf & vbLf & "O sistema será encerrado.")
        End
    End If
    On Error GoTo ControleErro
    lErro = 4
    '3. Posiciona data limite e verifica primeiro registro do arquivo histórico.
    lDatLimite = gM46V999.CDblx(Format$(Now - 7, "yyyymmdd"))
    lErro = 5
    Line Input #lNumSYASMULTLOG, lRegistro
    Do
        lErro = 6
        If Len(Trim(lRegistro)) < 10 Then
            Line Input #lNumSYASMULTLOG, lRegistro
        Else
            lErro = 7
            lAuxData = Mid$(lRegistro, 7, 4) & Mid$(lRegistro, 4, 2) & Mid$(lRegistro, 1, 2)
            lErro = 8
            If Not IsNumeric(lAuxData) Then
                Line Input #lNumSYASMULTLOG, lRegistro
            Else
                lErro = 9
                If gM46V999.CDblx(lAuxData) >= gM46V999.CDblx(lDatLimite) Then
                    lErro = 10
                    'O primeiro registro do arquivo já é superior à data limite - não precisa limpar.
                    On Error Resume Next
                    Close #lNumSYASMULTLOG
                    Err = 0
                    Exit Sub
                End If
                lErro = 11
                Exit Do
            End If
        End If
    Loop
    lErro = 12
    '4. Prepara quadro de aviso.
    panAviso.Caption = "Aguarde... " & vbLf & vbLf & "Processando limpeza no arquivo histórico..."
    panAviso.Left = (Me.ScaleWidth - panAviso.Width) / 2
    panAviso.Top = (Me.ScaleHeight - panAviso.Height) / 2
    panAviso.Visible = True
    panAviso.Refresh

    '5. Abre arquivo temporário.
    lNumSYASMULTTMP = FreeFile
    On Error Resume Next
    Open lEndSYASMULTTMP For Output Shared As #lNumSYASMULTTMP
    If Err <> 0 Then
        Call gM46V999.gpGraLog(2, "FTE0076 - Erro ao criar arquivo " & lEndSYASMULTTMP & _
                                " (arquivo temporário para limpeza do histórico - " & Err & " - " & Error & ").")
        End
    End If
    On Error GoTo ControleErro

    '6. Grava no arquivo temporário registros com data até 7 dias antes da data do dia.
    Do While Not EOF(lNumSYASMULTLOG)
        If Len(Trim(lRegistro)) > 10 Then
            lAuxData = Mid$(lRegistro, 7, 4) & Mid$(lRegistro, 4, 2) & Mid$(lRegistro, 1, 2)
            If IsNumeric(lAuxData) Then
                If CDbl(lAuxData) >= lDatLimite Then
                    Print #lNumSYASMULTTMP, lRegistro
                End If
            End If
        End If
        Line Input #lNumSYASMULTLOG, lRegistro
    Loop

    Close #lNumSYASMULTLOG
    Close #lNumSYASMULTTMP
    Kill lEndSYASMULTLOG
    Name lEndSYASMULTTMP As lEndSYASMULTLOG
    Call gM46V999.gpGraLog(0, "AGE0002 - Processo concluído: limpeza do arquivo histórico " & lEndSYASMULTLOG & _
                              ".")

    '7. Encerra processo.
    panAviso.Visible = False
    panAviso.Refresh
    Exit Sub
ControleErro:
    'Call gM46V999.gpGraLog(2, Err & " " & Err.Description, vbCritical)
    'Call gM46V999.gpGraLog(2, "mpLimHistorico ERRO - " & Err & " " & Err.Description)

    'Call MsgBox("mpLimHistorico ERRO - " & Err & " " & Err.Description & " LinhaErro " & lErro)

    Close #lNumSYASMULTLOG
    Close #lNumSYASMULTTMP


End Sub

Private Sub cmdAgendaInspecao_Click()
    Call mnuAgendarInspecao_Click
End Sub

Private Sub cmdAjuda_Click()

    If gM46V999.gfDisponivel("frmT46V101A") = True Then
        'Válido somente para o Projeto P46V101, não interfere quando executado pelo P46V001.
        Exit Sub
    End If
    Call mnuAjuSYAS_Click
End Sub

Private Sub cmdImpressao_Click()
    Call mnuFerImpressao_Click
End Sub

Private Sub cmdManCorretor_Click()
    Call mnuArqManCorretor_Click
End Sub

Private Sub cmdPerfil_Click()
    frmT46V164A.Show 1
End Sub

Private Sub cmdRenovacao_Click()
    Call mnuArqRenovacao_Click
End Sub
Private Sub cmdEndosso_Click()
    Call mnuArqEndosso_Click
End Sub
Private Sub cmdSegNovo_Click()
    Call mnuArqSegNovo_Click
End Sub
Private Sub cmdPesquisa_Click()
    Call mnuArqPesquisa_Click
End Sub
Private Sub cmdSair_Click()
    Call mnuSair_Click
End Sub
Private Sub cmdTransmitir_Click()
    Call mnuArqTransmissao_Click
End Sub
Private Sub Form_Activate()
    panForm.Caption = Mid$(Me.Name, 4)
    If mErrCarga = False Then
        If gM46V999.gfDisponivel("frmT46V101A") Then
            If frmT46V101A.Visible = True Then
                On Error Resume Next
                frmT46V101A.SetFocus
            End If
        End If
    Else
        Call gM46V999.gpGraLog(0, " mErrCarga " & mErrCarga)
        Screen.MousePointer = vbDefault
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    Dim lCodCorretor As String    'Código do corretor.
    Dim lDatGEISEDI2LOG As Double       'Data da versão do módulo GEISEDI2.LOG.
    Dim lDatGEISEDI2NET As Double       'Data da versão do módulo GEISEDI2.NET.
    Dim lDatP46V001EXE As Double      'Data da versão do módulo P46V001.EXE.
    Dim lDatUltCN As Double    'Data da versão do módulo do arquivo a ser considerado:
'GEISEDI2.LOG ou GEISEDI2.NET.
    Dim lDatVersao1 As String   'Data da versão no formato dd/mm/aaaa.
    Dim lDatVersao2 As Double   'Data da versão no formato aaaammdd.
    Dim lEndGeisedi2LOG As String       'Endereço do arquivo GEISEDI2.LOG ou ERROR.LOG. (LOCAL).
    Dim lEndGEISEDI2NET As String       'Endereço do arquivo GEISEDI2.LOG ou GEISEDI2.NET (REDE).
    Dim lEndP46V001BMP As String      'Endereço da imagem - P46V001.BMP.
    Dim lLargura As Integer    'Largura.
    Dim lMensagem As String    'Mensagem.
    Dim lPosicao As Integer    'Utilizado com InStr.
    Dim lPriVez As Boolean    'Indica que é primeira vez que está executando o sistema.
    Dim lQuaDiasL As Double    'Quantidade de dias do arquivo LOCAL.
    Dim lQuaDiasR As Double    'Quantidade de dias do arquivo da REDE.
    Dim lQuaUltCN As Double    'Quantidade de dias sem conexão.
    Dim objTabUser As clsA46V725A    'Objeto da classe clsA28V713A (Tab_User)
    Dim objstcTabUser As stcA46V725B    'Objeto da classe de estrutura stcA28V713B (Tab_User)
    Dim lCaminhoBase As String
    Dim lPropTabela As String
    Dim lErro As Integer
    Dim lDataProc As String
    Dim RelacaoUsuBookKit() As String
    Dim i As Integer
    
    lMensagem = "Início "
    'Call gM46V999.gpGraLog(0, lMensagem)
    Screen.MousePointer = vbHourglass
    lErro = 1
    On Error GoTo ControleErro

    '1. Verifica se form já está ativo e data do sistema.
    If App.PrevInstance = True Then
        Call gM46V999.gpGraLog(2, "AGE0003 - O programa de cálculo SYAS - MULTISSEGUROS já está ativo neste " & _
                                  "microcomputador." & vbLf & vbLf & "Utilize o programa que já está ativo.")
        End
    End If
    lErro = 2

    '    gM46V999.EhTeste = True
    If gM46V999.EhTeste Then
        MsgBox "Versão de Teste"
    End If

    '2. ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO ATENÇÃO.
    '   A qualquer alteração de qualquer módulo do SYAS (P46V001 e P46V100), alterar a linha abaixo.
    lMensagem = "Verificando a versão "
    lDatVersao1 = Format(FileDateTime(gM46V999.gAppPath & App.EXEName & ".EXE"), "dd/mm/yyyy")    'Data da versão: dd/mm/aaaa.
    lDatVersao2 = gM46V999.gfForData(1, lDatVersao1)     'Data da versão: aaaammdd.
    lMensagem = "Imagem "

    lMensagem = "Versão  " & FileDateTime(gM46V999.gAppPath & App.EXEName & ".EXE")
    'Call gM46V999.gpGraLog(0, lMensagem)

    '3. Carrega imagem.
    If Right$(App.Path, 1) = "\" Then
        lEndP46V001BMP = App.Path & "P46V001.jpg"
    Else
        lEndP46V001BMP = App.Path & "\P46V001.jpg"
    End If
    If Dir(lEndP46V001BMP) = "" Then
        If Right$(App.Path, 1) = "\" Then
            lEndP46V001BMP = App.Path & "P46V001.gif"
        Else
            lEndP46V001BMP = App.Path & "\P46V001.gif"
        End If
    End If
    On Error Resume Next

    lMensagem = "Imagem  " & lEndP46V001BMP
    'Call gM46V999.gpGraLog(0, lMensagem)

    imgYasuda.Picture = LoadPicture(lEndP46V001BMP)

    lMensagem = "Erro " & Error
    'Call gM46V999.gpGraLog(0, lMensagem)

    On Error GoTo ControleErro

    '4. Exibe form.
    panRodape.Visible = False
    fraEmissao.Visible = False
    lMensagem = "Vai mostrar a tela frmT46V001A "
    'Call gM46V999.gpGraLog(0, lMensagem)

    Me.Show

    '5. Obtém informações iniciais do sistema: tipo do usuário, endereços das bases de dados, endereços
    '   dos arquivos de impressão, indicador de uso do auxílio automático, nome ou código do usuário.
    lMensagem = "Vai executar o gM46V999.gpInicia "

    Call gM46V999.gpInicia(lPropTabela)

    lMensagem = "Vai executar o gM46V001.gpAtualizaIni "
    'Call gM46V999.gpGraLog(0, lMensagem)

    Call gM46V001.gpAtualizaIni

    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And gfIsMetaframe = False Then
        'Será executado apenas pelo P46V101
        If Not RegisterODBCDatabase("WORKFLOW", "SQL Server", True, "Description=BaseWorkflow" & vbCr & "Server=" & gM46V999.gSerGed & vbCr & "Database=" & gM46V999.gEndGed) Then
            MsgBox "Falha na criação do ODBC. Favor avisar HelpDesk!", vbCritical, "Criação de ODBC para SQL Server"
        End If
    End If

    'Call gM46V999.gpGraLog(0, lMensagem)
        '11.40
    'mfVerificaVersao

    Call gM46V999.gpGraLog(0, "Endereço " & gM46V999.gEndP00MULTI & ", Tipo " & gM46V999.gTipoHomologacao)

    '6. Encerra programa caso não tenha parâmetro para ser aberto.
    'Encerra programa caso não tenha sido chamado por outro.
    'Executável do SYAS que informou o parametro "YAS" para o acesso
    If UCase(PegaParam(1)) <> "/YAS" And Dir(gM46V999.gAppPath & "YASUDA.EXE") <> "" Then
        Call gM46V999.gpGraLog(2, "Iniciar o sistema através do programa " & vbLf & vbLf & gM46V999.gAppPath & _
                                  "YASUDA.EXE" & " .")
        End
    End If

    '7. gM46V999.gVerDataMaq indica se a função que verifica se houve ajuste na máquina deve ser startada ou não.
    '   Zero indica que não verifica e 1 indica que verifica se houve ajuste.
    gM46V999.gVerDataMaq = 1

    '8. Verifica se é usuário de digitação: perfil de acesso 4 indica utilização do sistema de corretor.
    gM46V999.PerfilAcesso = e_Perfil_CALC


    '9. Se for usuário corretor: posiciona indicador de primeira vez de execução do programa e obtém
    '    código do corretor.
    lMensagem = "Busca código do corretor "
    'Call gM46V999.gpGraLog(0, lMensagem)

    Call gpBuscaCodigoCorretor(lCodCorretor, lPriVez)

    'Ficha 1200030 - Bloqueio do Multiseguros para corretor 0622104
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        If CSng(lCodCorretor) = 622104 Then
            Call gM46V999.gpGraLog(1, "Corretor não cadastrado.")
            End
        End If
    End If

    '10. Se for usuário corretor, na primeira vez, aciona processo de localização das bases de dados.
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor And lPriVez = True Then
        lMensagem = "Localiza a base de Dados "
        Call mnuFerLocDados_Click
    End If

    '11. Limpa arquivo histórico.
    lMensagem = "Limpa Histórico"
    'Call gM46V999.gpGraLog(0, lMensagem)

    Call mpLimHistorico
    Err = 0
    lErro = 101
    lMensagem = "Voltou do limpa Histórico"
    'Call gM46V999.gpGraLog(0, lMensagem)

    '12. Emite mensagem de entrada.
    '    Obs.: Este comando precisa ser emitido depois de obter dados do arquivo de inicialização e
    '          também da verificação da data do arquivo log.
    Call gM46V999.gpGraLog(0, "AGE0001 - Início do processamento do programa " & App.EXEName & ".")
    lMensagem = "Verifica data do sistema"
    'Call gM46V999.gpGraLog(0, lMensagem)
    '13. Verifica data do sistema.

    If UCase(App.EXEName) = "P46V001" Then
        If Dir(gM46V999.gAppPath & "P46V001.EXE") <> "" Then
            lDatP46V001EXE = gM46V999.CDblx(Format(FileDateTime(gM46V999.gAppPath & "P46V001.EXE"), "yyyymmdd"))
        End If
    Else
        If Dir(gM46V999.gAppPath & "P46V101.EXE") <> "" Then
            lDatP46V001EXE = gM46V999.CDblx(Format(FileDateTime(gM46V999.gAppPath & "P46V101.EXE"), "yyyymmdd"))
        End If
    End If
    If lDatP46V001EXE > gM46V999.gHoje Then
        Call gM46V999.gpGraLog(2, "FTE0077 - Data do microcomputador (" & gM46V999.gForHoje & _
                                  ") menor que a data do programa (" & gM46V999.gfForData(2, lDatP46V001EXE) & ")." & _
                                  vbLf & vbLf & "O sistema será encerrado.")
        End     'Encerra programa.
    End If

    lMensagem = "Posiciona controle"
    'Call gM46V999.gpGraLog(0, lMensagem)

    '14. Posiciona controle.
    mErrCarga = False

    If gM46V999.gTrvSistema = True Then
        Call mpDesabilita
    End If
    lMensagem = "Posiciona Form "
    'Call gM46V999.gpGraLog(0, lMensagem)
    '15. Posiciona form - botões de comando e rodapé.
    'a) Botões de comando.
    fraEmissao.Visible = True
    fraEmissao.Height = Me.Height - 1150
    lblSair.Top = fraEmissao.Height - 450
    cmdSair.Top = lblSair.Top - 240
    lblAjuda.Top = lblSair.Top - 580
    cmdAjuda.Top = cmdSair.Top - 580
    lblManCorretor.Top = lblSair.Top - 1160
    cmdManCorretor.Top = cmdSair.Top - 1160

    lMensagem = "Verifica usuario Yasuda " & gM46V999.gNomUsuario
    'Call gM46V999.gpGraLog(0, lMensagem)

    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        Set objTabUser = New clsA46V725A
        Set objTabUser.M46V999 = gM46V999

        'Forçar usuario
        Set objstcTabUser = New stcA46V725B
        objstcTabUser.login_user_rede = UCase(gM46V999.gNomUsuario)
        gM46V999.gOrigemChamada = e_Origem_Kit
        If Not objTabUser.mfCarregarColecao(objstcTabUser) Then
            lMensagem = objTabUser.UltimoErro
            Call gM46V999.gpGraLog(2, lMensagem)
            gM46V999.gFinalizarPrograma = True
        End If
        If gM46V999.gFinalizarPrograma = True Then End
    End If
    
    mnuManutencao.Visible = (UCase(gM46V999.gNomUsuario) = "RSPAIXAO") Or _
                            (UCase(gM46V999.gNomUsuario) = "LVDELAZARI") Or _
                            (UCase(gM46V999.gNomUsuario) = "FSOUZA") Or _
                            (UCase(gM46V999.gNomUsuario) = "JMIWAI") Or _
                            (UCase(gM46V999.gNomUsuario) = "CYBER05") Or _
                            (UCase(gM46V999.gNomUsuario) = "SMANILLI") Or _
                            (UCase(gM46V999.gNomUsuario) = "FRODRIGUES") Or _
                            (UCase(gM46V999.gNomUsuario) = "TCINOUE") Or _
                            (UCase(gM46V999.gNomUsuario) = "BDRMELO")

    If gM46V999.gTipoAcesso = DIGITACAO Then
        lMensagem = "Tipo acesso " & gM46V999.gTipoAcesso
        'Call gM46V999.gpGraLog(0, lMensagem)
        If App.EXEName = "P46V001" Then
            Call gfShell("P46V101", True, vbNormalFocus)
            End
        End If
    End If

    lMensagem = "Desabilita opções "
    cmdImpressao.Visible = False
    lblImpressao.Visible = False
    mnuFerImpressao.Visible = False
    cmdPerfil.Visible = False
    LblPerfil.Visible = False
    mnuFerCadCCorretagem.Enabled = False
    mnuFerCadCCorretagem.Visible = False
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        cmdPerfil.Visible = True
        LblPerfil.Visible = True
        If App.EXEName = "P46V101" Then
            cmdImpressao.Visible = True
            lblImpressao.Visible = True
            mnuFerImpressao.Visible = True
        End If
        mnuFerCadCCorretagem.Visible = True
        mnuFerCadCCorretagem.Enabled = True
    End If

    cmdEndosso.Visible = (gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda)
    lblEndosso.Visible = (cmdEndosso.Visible)
    mnuArqEndosso.Visible = True

    'c) Desabilita as opções
    If (gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda) Then

        '###PROJETO INSPECAO 10018
        'DESCOMENTAR AS LINHAS ABAIXO
        cmdAgendaInspecao.Visible = False
        lblAgendaInspecao.Visible = False
        If gM46V999.PerfilAcesso <> e_Perfil_CALC And gM46V999.PerfilAcesso <> e_Perfil_CALC_Gestor Then           'Se não for Corretor.
            mnuArquivo.Enabled = False
            'mnuFerramentas.Enabled = False
            'mnuAjuda.Enabled = False
            cmdSegNovo.Visible = False
            lblSegNovo.Visible = False
            cmdRenovacao.Visible = False
            lblRenovacao.Visible = False
            cmdManCorretor.Visible = False
            lblManCorretor.Visible = False
            cmdEndosso.Visible = False
            lblEndosso.Visible = (cmdEndosso.Visible)
            mnuArqEndosso.Visible = False

            mnuFerImpressao.Visible = False
            cmdImpressao.Visible = False
            lblImpressao.Visible = False
            mnuArqTransmissao.Visible = False

            cmdTransmitir.Visible = False
            lblTransmitir.Visible = False
        End If
    End If

    If gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        cmdTransmitir.Visible = False
        lblTransmitir.Visible = False
        '###PROJETO INSPECAO 10018
        'DESCOMENTAR AS LINHAS ABAIXO
        cmdAgendaInspecao.Visible = False
        lblAgendaInspecao.Visible = False
    End If
    If gM46V999.PerfilAcesso = e_Perfil_Operacao Then
        cmdImpressao.Visible = True
        lblImpressao.Visible = True
    End If


    lMensagem = "Monta Rodape"
    'Call gM46V999.gpGraLog(0, lMensagem)

    'd) Rodapé.
    panRodape.Visible = True
    panForm.Caption = "T46V001A"
    lPosicao = InStr(UCase$(gM46V999.gEndP00MPED), "P00MPED.MDB")

    Select Case lPosicao
    Case 0
        lCaminhoBase = "Cotação: Não encontrado"
    Case 1
        lCaminhoBase = "Cotação: " & gM46V999.gAppPath
    Case Else
        lCaminhoBase = "Cotação: " & Left$(gM46V999.gEndP00MPED, lPosicao - 2)
    End Select

    '    lPosicao = InStr(UCase$(gM46V999.gEndP00MULTI), "P00MULTI.MDB")
    '
    '    Select Case lPosicao
    '        Case 0
    '            lCaminhoBase = lCaminhoBase & " - Tarifa: Não encontrado"
    '        Case 1
    '            lCaminhoBase = lCaminhoBase & " - Tarifa: " & gM46V999.gAppPath
    '        Case Else
    '            lCaminhoBase = lCaminhoBase & " - Tarifa: " & Left$(gM46V999.gEndP00MULTI, lPosicao - 2)
    '    End Select

    panBasDados.Caption = lCaminhoBase
    panVersao.Caption = "Versão " & lDatVersao1
    If UCase(lCodCorretor) = "YASUDA" Then
        panCorretor.Caption = lCodCorretor
        gM46V999.gCod_Corretor = 0
    Else
        panCorretor.Caption = "Corretor " & lCodCorretor
        gM46V999.gCod_Corretor = val(lCodCorretor)
    End If
    panUsuario.Caption = "Usuário " & gM46V999.gNomUsuario6
    lMensagem = "Ve SYASEDI"
    'Call gM46V999.gpGraLog(0, lMensagem)

    If Trim$(gM46V999.gNomUsuario6) = "000000" Then
        'Exclui o arquivo SyasEdi.EXE / EapScript.ini
        'para que o codigo do usuario seja solicitado novamente
        If gM46V999.gfPreenchido(gM46V999.gAppPath & "SYASEDI.EXE") = True Then
            On Error Resume Next
            Kill gM46V999.gAppPath & "SYASEDI.EXE"
            On Error GoTo 0
        End If

        If gM46V999.gfPreenchido(gM46V999.gAppPath & "Eapscript.INI") = True Then
            On Error Resume Next
            Kill gM46V999.gAppPath & "Eapscript.INI"
            On Error GoTo 0
        End If
        lMensagem = "Ve EAPSCRIPT001.INI"
        If gM46V999.gfPreenchido(gM46V999.gAppPath & "EAPSCRIPT001.INI") = True Then
            On Error Resume Next
            Kill gM46V999.gAppPath & "EAPSCRIPT001.INI"
            Name gM46V999.gAppPath & "EAPSCRIPT001.001" As gM46V999.gAppPath & "EAPSCRIPT001.003"
            On Error GoTo 0
        End If

        If gM46V999.gfPreenchido(gM46V999.gAppPath & "Eapscript006.INI") = True Then
            On Error Resume Next
            Kill gM46V999.gAppPath & "Eapscript006.INI"
            Name gM46V999.gAppPath & "Eapscript006.001" As gM46V999.gAppPath & "Eapscript006.003"
            On Error GoTo 0
        End If

        Call gM46V999.gpGraLog(1, "AGE0113 - Código de usuário cadastrado é inválido, solicite um novo código com a Central de Atendimento." & Chr(13) & _
                                  "Ao acessar o sistema novamente o código de usuario de transmissão será solicitado.")
        lMensagem = "Finaliza"
        'Call gM46V999.gpGraLog(0, lMensagem)
        End
    End If
    lMensagem = "Continua Rodape"
    'Call gM46V999.gpGraLog(0, lMensagem)

    panData.Caption = gM46V999.gForHoje
    lLargura = panForm.Width + panBasDados.Width + panVersao.Width + panCorretor.Width + _
               panUsuario.Width + panData.Width
    If panRodape.Width > lLargura Then
        panBasDados.Width = panRodape.Width + panBasDados.Width - lLargura - 120
    End If
    panForm.Left = 60
    panBasDados.Left = panForm.Left + panForm.Width
    panVersao.Left = panBasDados.Left + panBasDados.Width
    panCorretor.Left = panVersao.Left + panVersao.Width
    panUsuario.Left = panCorretor.Left + panCorretor.Width
    panData.Left = panUsuario.Left + panUsuario.Width
    'e) Refresh.
    Me.Refresh

    '16. Executa opção de menu do SYAS Emissão.
    'a) Posiciona controles.
    '   Obs.: não utilizar "SYAS - Sistema YASUDA Seguros - Emissão" por causa da instrução AppActivate.
    Me.Caption = "SYAS - Sistema YASUDA Multisseguros"
    fraEmissao.Height = Me.Height - 1200
    fraEmissao.Visible = True
    cmdSair.Cancel = True
    'b) Exibe form.
    Me.Refresh

    '17. Posiciona help.
    App.HelpFile = gM46V999.gAppPath & "SYASMULT.HLP"

    '18. Verifica última conexão - apenas para corretor.
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then      'Corretor.
        'a) Posiciona controles.
        mPreConexao = False
        mFezConexao = False

        'Obtem informações da ultima conexão
        'Procura o Arquivo na mesma pasta de onde se encontra
        ' as bases tarifárias
        If gM46V999.gfPreenchido(gM46V999.gEndSYASUpDt) = True Then
            lEndGEISEDI2NET = gM46V999.gEndSYASUpDt & "GEISEDI2.NET"
        Else
            lEndGEISEDI2NET = gM46V999.gEndServidor & "GEISEDI2.NET"
        End If
        If gM46V999.gfPreenchido(Dir(lEndGEISEDI2NET)) Then
            lDatGEISEDI2NET = CDbl(Format$(FileDateTime(lEndGEISEDI2NET), "yyyymmdd"))
        Else
            lDatGEISEDI2NET = 0
        End If

        'Obtem dias sem conexão dos arquivos rede
        If lDatGEISEDI2NET = 0 Then
            'Caso não encontre arquivos
            lQuaDiasR = 0
        Else
            If lDatGEISEDI2NET < gM46V999.gHoje Then
                lQuaDiasR = gM46V999.gfDifDatas(lDatGEISEDI2NET, gM46V999.gHoje)
            Else
                lQuaDiasR = 0
            End If
        End If

        'atribui qual arquivo será ultilizado na verificação
        'Rede ou Local
        lQuaUltCN = lQuaDiasR
        lDatUltCN = lDatGEISEDI2NET

        'Faz a Verificação
        If gM46V999.gNomUsuario <> "447627" Then
            Call mpVerificaUltCn(lDatUltCN, lQuaUltCN)
        End If
    End If

    If Trim(gM46V999.gSerSyasPortal) <> "" Then
        lPropTabela = "Syasportal.dbo."
        gM46V999.gPrograma = "D046O003"
    End If
    gM46V999.gPrograma = App.EXEName
    If Dir(gM46V999.gAppPath & "syas_robo.exe") = "" Then
'        lMensagem = "Verifica Versão do Banco P00MULTI"
'        'Call gM46V999.gpGraLog(0, lMensagem)
'        If mfObtemVtarMult(lMensagem) = False Then
'            Call gM46V999.gpGraLog(1, lMensagem, vbInformation)
'            cmdSair_Click
'            Exit Sub
'        End If

        lMensagem = "Ver Registro Padrão"
        'Call gM46V999.gpGraLog(0, lMensagem)

        '24. Verifica a existencia de novas tabelas
        lMensagem = "Cria Tabelas"
        'Call gM46V999.gpGraLog(0, lMensagem)

        'Retirado em 05/10/2009
        If gM46V001.mfCriaTabelas(lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If


        '23. Verifica existência do registro padrão.
        If gM46V999.mfVerRegPadrao(lMensagem) = False Then

            Call gM46V999.gpGraLog(0, "SAIU DO Registro Padrão")

            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If

        'Classe
        If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
            Call mpVerRecotacao
        End If
    End If

'    If UCase(PegaParam(1)) = "/BOOK" Then
     RelacaoUsuBookKit = Split(gM46V999.gRelUsuBookKit, ";")
     
     For i = 0 To UBound(RelacaoUsuBookKit())
         If UCase(gM46V999.gNomUsuario) = UCase(RelacaoUsuBookKit(i)) Then
            lDataProc = ""
            lDataProc = InputBox("Caso deseje executar o book, preencha a data de processamento formato DDMM:", Me.Caption)
            If lDataProc = "" Then
                GoTo SeguirProcesso
            End If
            panAviso.Visible = True
            panAviso.Caption = "Aguarde conferência do book de multi"
    '        If Not IsNumeric(PegaParam(2)) Then
    '            Call gM46V999.gpGraLog(1, "Parametro de data inválido")
    '            End
    '        End If
            Call mpAnalisaBOOK("BKCN", lDataProc)
            Call mpAnalisaBOOK("BKEM", lDataProc)
            Call mpAnalisaBOOK("BKRE", lDataProc)
            End
        End If
    Next
    
    ''' 2016-08-11 : Removida limitacao
    ''' 2016-07-12 : Verifica perfil de acesso
    ''' gM46V999.gTipoAcesso = CALCULO
    ''' gM46V999.PerfilAcesso = e_...
    ''' gM46V999.gPermiteImpressaoFtp = False
    'Dim ArrUsuarios() As String
    'ArrUsuarios = Split(gM46V999.gRelUsuImpressaoFtp, ";")
    'For i = 0 To UBound(ArrUsuarios())
    '    If UCase(ArrUsuarios(i)) = UCase(gM46V999.gNomUsuario) Then
    '        mnuFerImpressaoTransmissao.Visible = True
    '        Exit For
    '    End If
    'Next
    
SeguirProcesso:
    '19. Encerra processo.
    gM46V111.gNum_Pi = 0
    Screen.MousePointer = vbDefault
    Select Case gM46V999.gTipoAcesso
    Case DIGITACAO
        lMensagem = "vai pesquisar "
        'Call gM46V999.gpGraLog(0, lMensagem)
        Call mnuArqPesquisa_Click
    End Select

    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V001.Load ERRO - " & lMensagem & lErro & " - " & Err & " " & Err.Description)
End Sub

Private Sub mpAnalisaBOOK(pArquivo As String, pDataProc As String)
    Dim lNumArq As Integer
    Dim lEndArquivo As String
    Dim lLinha As String
    Dim lPI As Long
    Dim lAcao As Long
    Dim lInicioPI As Boolean
    Dim lNovaPag As Boolean
    Dim lSql As String
    Dim lData As String
    Dim lrst As ADODB.Recordset
    Dim lPIAnt As Long
    Dim lItemAnt As Long
    Dim lMsg As String
    Dim lFranqVidros As String
    Dim lI As Integer
    Dim lFormulario As String
    Dim lLista As String
    Dim lNomeArq As String
    Dim lNumApol As Double
    Dim lNumEndo As Long
    Dim lNumApolAnt As Double
    Dim lNumEndoAnt As Long
    Dim lAuxiliar As String
    Dim clsA46V101A As clsA46V101A
    Dim lObjPed As clsA46V702A
    Dim lObjstcPed As stcA46V702B
    Dim lObjstcPedOrig As stcA46V702B
    Dim lIndColetiva As Boolean
    Dim lobjstcPedLoc As stcA46V708B
    Dim lobjstcPedLocOrig As stcA46V708B
    'Dim ColecaoPedLoc            As Collection               'Coleção de itens.

    Dim ColecaoCoberturaTMP As Collection
    Dim lobjstcCobertura As stcA46V704B  'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
    Dim lobjstcCoberturaOrig As stcA46V704B  'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
    'Dim ColecaoClausula     As Collection
    Dim lobjstcPedClau As stcA46V703B    'Objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU)
    Dim lObjstcPedClauOrig As stcA46V703B    'Objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU)

    'Dim ColecaoBem          As Collection
    Dim lobjstcBem As stcA46V715B    'Objeto da classe de estrutura stcA46V715B (TAB_PED_LOC_BEM)
    Dim lobjstcBemOrig As stcA46V715B    'Objeto da classe de estrutura stcA46V715B (TAB_PED_LOC_BEM)


    Dim lObjstcInsp As stcA46V717B    'Objeto da classe de estrutura stcA46V717B (TAB_PED_INSP)
    'Dim ColecaoFatoProt     As Collection
    Dim lobjstcFator As stcA46V720B    'Objeto da classe de estrutura stcA46V720B (Tab_Ped_Fat_Risco)
    Dim lobjstcFatorOrig As stcA46V720B    'Objeto da classe de estrutura stcA46V720B (Tab_Ped_Fat_Risco)


    Dim lObjstcPedCorr As stcA46V705B
    Dim lObjstcPedCorrOrig As stcA46V705B
    'Dim ColecaoPedCorr           As Collection               'Coleção de corretores.7
    Dim lobjstcPedCoss As stcA46V706B
    Dim lobjstcPedCossOrig As stcA46V706B
    'Dim ColecaoPedCoss           As Collection               'Coleção de cosseguros.

    Dim lobjstcPedEndereco As stcA46V707B       'Objeto de endereço.
    Dim lbjstcCtrlEmis As stcA46V712B    'Objeto de controle de emissão.

    'Dim ColecaoTextoEndo         As Collection               'Objeto de Texto para Endosso.

    'Dim ColecaoTextoEspec        As Collection               'Objeto de Texto para Especificação Padrão.

    Dim lPontoInicial As Long
    Dim lTamanho As Long
    Dim lDataArq As Long

    'lData = PegaParam(2)
    lData = pDataProc


    lNomeArq = pArquivo

    lNumArq = FreeFile
    lEndArquivo = "\\Spx20103psluapl\import$\" & lNomeArq & lData & ".BMK"
    
    On Error Resume Next
    Open lEndArquivo For Input Access Read As #lNumArq
    If Err <> 0 Then
        On Error GoTo 0
        Exit Sub
    End If


    On Error GoTo ControleErro
    lLista = ""
    lNumApol = 0
    lNumEndo = 0
    
    Call gM46V999.gpGraLog(0, "Inicio Processo: " & lNomeArq)
    
    Do While Not EOF(lNumArq)
        Line Input #lNumArq, lLinha

        lNumApol = Mid(lLinha, 13, 10)
        lNumEndo = Mid(lLinha, 23, 6)
        If lNumApol <> lNumApolAnt And lNumEndo <> lNumEndoAnt Then
            '''Realiza a comparação
            'MsgBox lFormulario

        End If

        'If lNumEndo = 0 Then
        '    GoTo Proximo
        'End If
        lNumApolAnt = lNumApol
        lNumEndoAnt = lNumEndo

        lFormulario = Mid(lLinha, 35, 8)
        Select Case lFormulario
        Case "L02C0001"
            Call gM46V999.gpGraLog(0, "Processamento L02C0001")
            lIndColetiva = False
            '            073000 01  FORM-CAPA-WS.                                                11
            Set lobjstcPedLoc = New stcA46V708B
            Set lObjstcPed = New stcA46V702B
            lPontoInicial = 1
            ''                073100     03  COD-RAMO-CAPA          PIC 9(03)      VALUE ZEROS.
            lTamanho = 3
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Cod_Ramo = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '                073200     03  COD-ENVI-CAPA          PIC X(01)      VALUE SPACES.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 7
            '                073300     03  COD-CORR-CAPA          PIC 9(07)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '                073400     03  ORDEM-CAPA             PIC 9(01)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '                073500     03  NUM-APOL-CAPA          PIC 9(10)      VALUE ZEROS.       11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 6
            '                073600     03  NUM-ENDO-CAPA          PIC 9(06)      VALUE ZEROS.       11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '                073700     03  NUM-ITEM-CAPA          PIC 9(05)      VALUE ZEROS.       11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '                073800     03  FLAG-COL-CAPA          PIC 9(01)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lIndColetiva = IIf(lAuxiliar <> "1", True, False)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 8
            '                073900     03  COD-FORM-CAPA          PIC X(08)      VALUE 'L02C0001'.  11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 15
            '                074000     03  VIA-CAPA               PIC X(15)      VALUE SPACES.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho


            lTamanho = 60
            '                074100     03  NOM-SEGUR-CAPA         PIC X(60)      VALUE SPACES.      11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Nom_Segurado = lAuxiliar

            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '                074200     03  NRO-APOL-CAPA          PIC 9(10)      VALUE ZEROS.       11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '                074300     03  HIFEN-APOL-CAPA        PIC X(01)      VALUE SPACES.      11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '                074400     03  DIG-APOL-CAPA          PIC 9(01)      VALUE ZEROS.       11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '                074500     03  DIA-EMIS-CAPA          PIC 9(02)      VALUE ZEROS.       11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '                074600     03  BARRA-1-EMIS-CAPA      PIC X(01)      VALUE SPACES.      11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '                074700     03  MES-EMIS-CAPA          PIC 9(02)      VALUE ZEROS.       11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '                074800     03  BARRA-2-EMIS-CAPA      PIC X(01)      VALUE SPACES.      11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 4
            '                074900     03  ANO-EMIS-CAPA          PIC 9(04)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 30


            '----------------------------------corretor-------------------------------------
            '                075000     03  NOM-CORR-1-CAPA        PIC X(30)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 30
            '                075100     03  NOM-CORR-2-CAPA        PIC X(30)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 30
            '                075200     03  NOM-CORR-3-CAPA        PIC X(30)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            '-------------------------------------------------------------------------------

            lTamanho = 2
            '                075300     03  CARNE-CAPA             PIC X(02)      VALUE SPACES.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '                075400     03  NUM-ITEM-DE-CAPA       PIC 9(05)      VALUE ZEROS.       11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '                075500     03  NUM-ITEM-ATE-CAPA      PIC 9(05)      VALUE ZEROS.       11
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 150
            '                075600     03  TEXTO-1-CAPA           PIC X(150)     VALUE SPACES.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 150
            '                075700     03  TEXTO-2-CAPA           PIC X(150)     VALUE SPACES.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)

        Case "L02C0003"
            Call gM46V999.gpGraLog(0, "Processamento L02C0003")
            Set lobjstcPedLoc = New stcA46V708B
            Set lObjstcPed = New stcA46V702B
            lPontoInicial = 1
            lTamanho = 3
            '081500     03  COD-RAMO-CAPA          PIC 9(03)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Cod_Ramo = lAuxiliar

            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '081600     03  COD-ENVI-CAPA          PIC X(01)      VALUE SPACES.      14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 7
            '081700     03  COD-CORR-CAPA          PIC 9(07)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '081800     03  ORDEM-CAPA             PIC 9(01)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '081900     03  NUM-APOL-CAPA          PIC 9(10)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 6
            '082000     03  NUM-ENDO-CAPA          PIC 9(06)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '082100     03  NUM-ITEM-CAPA          PIC 9(05)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '082200     03  FLAG-COL-CAPA          PIC 9(01)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lIndColetiva = IIf(lAuxiliar <> "1", True, False)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 8
            '082300     03  COD-FORM-CAPA          PIC X(08)      VALUE 'L02C0003'.  14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 15
            '082400     03  VIA-CAPA               PIC X(15)      VALUE SPACES.      14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 60
            '082500     03  NOM-SEGUR-CAPA         PIC X(60)      VALUE SPACES.      14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Nom_Segurado = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '082600     03  NRO-APOL-CAPA          PIC 9(10)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '082700     03  DIA-EMIS-CAPA          PIC 9(02)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '082800     03  BARRA-1-EMIS-CAPA      PIC X(01)      VALUE SPACES.      14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '082900     03  MES-EMIS-CAPA          PIC 9(02)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '083000     03  BARRA-2-EMIS-CAPA      PIC X(01)      VALUE SPACES.      14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 4
            '083100     03  ANO-EMIS-CAPA          PIC 9(04)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 30
            '083200     03  NOM-CORR-1-CAPA        PIC X(30)      VALUE SPACES.      14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 30
            '083300     03  NOM-CORR-2-CAPA        PIC X(30)      VALUE SPACES.      14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 30
            '083400     03  NOM-CORR-3-CAPA        PIC X(30)      VALUE SPACES.      14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '083500     03  CARNE-CAPA             PIC X(02)      VALUE SPACES.      14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 150
            '083600     03  TEXTO-1-CAPA           PIC X(150)     VALUE SPACES.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 150
            '083700     03  TEXTO-2-CAPA           PIC X(150)     VALUE SPACES.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
        Case "L02C0005"
            Call gM46V999.gpGraLog(0, "Processamento L02C0005")
            '            075900 01  FORM-APOL-WS.
            lPontoInicial = 1
            lTamanho = 3
            '076000     03  COD-RAMO-APOL          PIC 9(03)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '076100     03  COD-ENVI-APOL          PIC X(01)      VALUE SPACES.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 7
            '076200     03  COD-CORR-APOL          PIC 9(07)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '076300     03  ORDEM-APOL             PIC 9(01)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '076400     03  NUM-APOL-APOL          PIC 9(10)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 6
            '076500     03  NUM-ENDO-APOL          PIC 9(06)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            If Not lobjstcPedLoc Is Nothing Then
                If lobjstcPedLoc.Num_Item <> 0 Then
                    On Error Resume Next
                    lObjstcPed.ColecaoPedLoc.Add lobjstcPedLoc, lobjstcPedLoc.Class_ID
                    If Err.Number <> 0 Then
                        MsgBox ""
                    End If
                    On Error GoTo ControleErro
                End If
            End If
            Set lobjstcPedLoc = New stcA46V708B
            '076600     03  NUM-ITEM-APOL          PIC 9(05)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Num_Item = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '076700     03  FLAG-COL-APOL          PIC 9(01)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            If lAuxiliar = 0 Then
                GoTo Proximo
            End If
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 8
            '076800     03  COD-FORM-APOL          PIC X(08)      VALUE 'L02C0005'.  12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 20
            '076900     03  NOM-RAMO-APOL          PIC X(20)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 30
            '077000     03  NUM-SUSEP-APOL         PIC X(30)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '077100     03  NRO-APOL-APOL          PIC 9(10)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '077200     03  HIFEN-APOL-APOL        PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '077300     03  DIG-APOL-APOL          PIC 9(01)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '077400     03  NRO-ITEM-APOL          PIC 9(05)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 15
            '077500     03  NRO-PROP-APOL          PIC X(15)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 4
            '077600     03  COD-UNID-PROD-APOL     PIC 9(04)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Cod_Unid_Prod = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '077700     03  HIFEN-UNID-APOL        PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 3
            '077800     03  COD-AGEN-ANGAR-APOL    PIC 9(03)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '077900     03  HIFEN-PROD-APOL        PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 4
            '078000     03  COD-PROD-APOL          PIC 9(04)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Cod_Prod = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '078100     03  NRO-APOL-ANT-APOL      PIC 9(10)      BLANK WHEN ZEROS.  12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Num_Apol_An = val(lAuxiliar)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '078200     03  HIFEN-APOL-ANT-APOL    PIC X(01)         VALUE SPACES.   12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '078300     03  DIG-APOL-ANT-APOL      PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '078400     03  DIA-INI-VIG-APOL       PIC 9(02)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Dat_Ini_Vig = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '078500     03  BARRA-1-INI-APOL       PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '078600     03  MES-INI-VIG-APOL       PIC 9(02)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Dat_Ini_Vig = lAuxiliar & Format(lObjstcPed.Dat_Ini_Vig, "00")
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '078700     03  BARRA-2-INI-APOL       PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 4
            '078800     03  ANO-INI-VIG-APOL       PIC 9(04)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Dat_Ini_Vig = lAuxiliar & Format(lObjstcPed.Dat_Ini_Vig, "0000")
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '078900     03  DIA-TER-VIG-APOL       PIC 9(02)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Dat_Ter_Vig = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '079000     03  BARRA-1-TER-APOL       PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '079100     03  MES-TER-VIG-APOL       PIC 9(02)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Dat_Ter_Vig = lAuxiliar & Format(lObjstcPed.Dat_Ter_Vig, "00")
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '079200     03  BARRA-2-TER-APOL       PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 4
            '079300     03  ANO-TER-VIG-APOL       PIC 9(04)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Dat_Ter_Vig = lAuxiliar & Format(lObjstcPed.Dat_Ter_Vig, "0000")
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 60
            '079400     03  NOME-SEGUR-APOL        PIC X(60)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Nom_Segurado = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 18
            '079500     03  CNPJ-CPF-SEGUR-APOL    PIC X(18)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Num_Cgc_Cpf = Replace(Replace(Replace(lAuxiliar, ".", ""), "-", ""), "/", "")
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 3
            '079600     03  COD-TIP-LOGR-APOL      PIC X(03)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Cod_Tip_Logr = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 50
            '079700     03  NOME-LOGR-APOL         PIC X(50)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Nom_Logr = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 20
            '079800     03  DSC-COMPL-APOL         PIC X(20)         VALUE SPACES.   12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Dsc_Compl_Logr = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 25
            '079900     03  BAIRRO-APOL            PIC X(25)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Nom_Bairro = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 25
            '080000     03  CIDADE-APOL            PIC X(25)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Nom_Cid = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '080100     03  SIG-UF-APOL            PIC X(02)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Sig_UF = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 3
            '080200     03  TIPO-LOGR-RISCO-APOL   PIC X(03)         VALUE SPACES.   12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            If Not gM46V999.gfPreenchido(lAuxiliar) Then
                'deve ser coletiva
                lobjstcPedLoc.Num_Item = 0
            End If
            lobjstcPedLoc.Cod_Tip_Logr_Risco = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 50
            '080300     03  NOME-LOGR-RISCO-APOL   PIC X(50)         VALUE SPACES.   12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Nom_Logr_Risco = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '080400     03  NUM-LOGR-RISCO-APOL    PIC ---------9 BLANK WHEN ZEROS.  12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            Set lobjstcPedLoc.objstcPedInsp = New stcA46V717B
            lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco = val(lAuxiliar)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 20
            '080500     03  DSC-COMPL-RISCO-APOL   PIC X(20)         VALUE SPACES.   12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Dsc_Compl_Risco = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 25
            '080600     03  BAIRRO-RISCO-APOL      PIC X(25)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Nom_Bairro_Risco = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 25
            '080700     03  CIDADE-RISCO-APOL      PIC X(25)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Nom_Cid_Risco = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '080800     03  SIG-UF-RISCO-APOL      PIC X(02)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Sig_UF_Risco = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '080900     03  NUM-CEP-RISCO-1-APOL   PIC X(05)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Num_Cep_Risco = val(lAuxiliar)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '081000     03  HIFEN-CEP-APOL         PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 3
            '081100     03  NUM-CEP-RISCO-2-APOL   PIC X(03)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Num_Cep_Risco = lobjstcPedLoc.Num_Cep_Risco & lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 14
            '081200     03  COD-CLASF-APOL         PIC X(14)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Cod_Clasf = Left(lAuxiliar, 3)
            lobjstcPedLoc.Cod_Compl_Clasf = Mid(lAuxiliar, 6, 2)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 60
            '081300     03  DSC-CLASF-APOL         PIC X(60)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Dsc_Clasf = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 20
            '081400     03  VAL-RISC-APOL          PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99          12
            '081500                                               VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            If IsNumeric(lAuxiliar) Then
                lobjstcPedLoc.Val_Vr = lAuxiliar
            Else
                lobjstcPedLoc.Val_Vr = val(lAuxiliar)
            End If
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 15
            '081800     03  NOM-CORR-1-APOL        PIC X(15)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            Set lObjstcPedCorr = New stcA46V705B
            lObjstcPedCorr.ComparaCorretor = lAuxiliar
            If gM46V999.gfPreenchido(lAuxiliar) Then
                If lObjstcPed.ColecaoPedCorr.Count = 0 Then
                    lObjstcPedCorr.Cod_Corr = 1
                    On Error Resume Next
                    lObjstcPed.ColecaoPedCorr.Add lObjstcPedCorr, lObjstcPedCorr.Class_ID
                    If Err.Number <> 0 Then
                        MsgBox ""
                    End If
                    On Error GoTo ControleErro
                End If
            End If
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '081900     03  ABRE-PAR-1-APOL        PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 14
            '082000     03  COD-CORR-1-APOL        PIC 9(14).                        12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho


            lTamanho = 1
            '082300     03  FECHA-PAR-1-APOL       PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 15
            '082400     03  NOM-CORR-2-APOL        PIC X(15)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            Set lObjstcPedCorr = New stcA46V705B
            lObjstcPedCorr.ComparaCorretor = lAuxiliar
            If gM46V999.gfPreenchido(lAuxiliar) Then
                If lObjstcPed.ColecaoPedCorr.Count = 1 Then
                    lObjstcPedCorr.Cod_Corr = 2
                    On Error Resume Next
                    lObjstcPed.ColecaoPedCorr.Add lObjstcPedCorr, lObjstcPedCorr.Class_ID
                    If Err.Number <> 0 Then
                        MsgBox ""
                    End If
                    On Error GoTo ControleErro
                End If
            End If
            lTamanho = 1
            '082500     03  ABRE-PAR-2-APOL        PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 14
            '082600     03  COD-CORR-2-APOL        PIC 9(14).                        12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho

            lTamanho = 1
            '082900     03  FECHA-PAR-2-APOL       PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 15
            '083000     03  NOM-CORR-3-APOL        PIC X(15)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            Set lObjstcPedCorr = New stcA46V705B
            lObjstcPedCorr.ComparaCorretor = lAuxiliar
            If gM46V999.gfPreenchido(lAuxiliar) Then
                If lObjstcPed.ColecaoPedCorr.Count = 2 Then
                    lObjstcPedCorr.Cod_Corr = 3
                    On Error Resume Next
                    lObjstcPed.ColecaoPedCorr.Add lObjstcPedCorr, lObjstcPedCorr.Class_ID
                    If Err.Number <> 0 Then
                        MsgBox ""
                    End If
                    On Error GoTo ControleErro
                End If
            End If
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '083100     03  ABRE-PAR-3-APOL        PIC X(01)      VALUE SPACES.      12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho

            lTamanho = 1
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            '083500     03  FECHA-PAR-3-APOL       PIC X(01)      VALUE SPACES.      12

            '------------------------------------------------------------------------
            '083700     03  NOM-COBERT-APOL        OCCURS         20    TIMES.       00

            'CR49835: Arquivo de Comparação - Geração Book Produto Empresarial
            '083700     03  NOM-COBERT-APOL        OCCURS         30    TIMES.       00
            'For lI = 0 To 19
            For lI = 0 To 29
                lTamanho = 4
                Set lobjstcCobertura = New stcA46V704B
                lobjstcCobertura.Num_Item = lobjstcPedLoc.Num_Item
                '083800         07  COD-COBERT-APOL    PIC X(04)      VALUE SPACES.      00
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lobjstcCobertura.Cod_Cobert = val(lAuxiliar)
                lTamanho = 1
                '083900         07  HIFEN-COBERT-APOL  PIC X(01)      VALUE SPACES.      00
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 50
                '084000         07  DSC-COBERT-APOL    PIC X(50)      VALUE SPACES.      00
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 20
                '084100         07  VAL-IS-APOL        PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         00
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                If IsNumeric(lAuxiliar) Then
                    lobjstcCobertura.Val_IS = lAuxiliar
                Else
                    lobjstcCobertura.Val_IS = val(lAuxiliar)
                End If
                lTamanho = 2
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lobjstcCobertura.Num_Meses_PI = val(lAuxiliar)
                lTamanho = 2
                '084700         07  QTD-DIAS-APOL      PIC ZZ         VALUE ZEROS.       00
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho

                lTamanho = 6
                '085000         07  PER-FRANQ-APOL     PIC ZZ9,99     VALUE ZEROS.       00

                'CR49835: Arquivo de Comparação - Geração Book Produto Empresarial
                '085000         07  PER-FRANQ-APOL-R   PIC Z9,99     VALUE ZEROS.        00
                lAuxiliar = Replace(Mid(lLinha, lPontoInicial, lTamanho), "*", " ")
                lPontoInicial = lPontoInicial + lTamanho
                lobjstcCobertura.Per_Franq = val(lAuxiliar)

                'CR49835: Arquivo de Comparação - Geração Book Produto Empresarial
                'lTamanho = 10
                lTamanho = 16
                '085300         07  VAL-FRANQ-MIN-APOL PIC ZZZ.ZZ9,99.                   00

                'CR49835: Arquivo de Comparação - Geração Book Produto Empresarial
                '085300         07  VAL-FRANQ-MIN-APOL-R PIC Z.ZZZ.ZZZ.ZZ9,99.           00
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                If IsNumeric(lAuxiliar) Then
                    lobjstcCobertura.Val_Franq_Min = lAuxiliar
                Else
                    lobjstcCobertura.Val_Franq_Min = val(lAuxiliar)
                End If
                If lobjstcCobertura.Cod_Cobert > 0 Then
                    On Error Resume Next
                    lobjstcPedLoc.ColecaoPedCobert.Add lobjstcCobertura, lobjstcCobertura.Class_ID
                    If Err.Number <> 0 Then
                        MsgBox ""
                    End If
                    On Error GoTo ControleErro
                End If
            Next
            '------------------------------------------------------------------------------'

            For lI = 0 To 4
                Set lobjstcPedCoss = New stcA46V706B
                '------------------------------------------------------------------------------'
                '085700     03  COSSEG-APOL            OCCURS         5     TIMES.       12
                lTamanho = 25
                '085800         05  NOM-COSS-APOL      PIC X(25)      VALUE SPACES.      12
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lobjstcPedCoss.Nom_Seguradora = lAuxiliar
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 5
                '085900         05  PER-COSS-APOL      PIC Z9,99.                        12
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                If IsNumeric(lAuxiliar) Then
                    lobjstcPedCoss.Per_Coss = lAuxiliar
                Else
                    lobjstcPedCoss.Per_Coss = val(lAuxiliar)
                End If
                lobjstcPedCoss.Cod_Coss = lI
                lPontoInicial = lPontoInicial + lTamanho
                If lobjstcPedCoss.Per_Coss > 0 Then
                    On Error Resume Next
                    lObjstcPed.ColecaoPedCoss.Add lobjstcPedCoss, lobjstcPedCoss.Class_ID
                    On Error GoTo ControleErro
                End If
            Next

            '---------------------------------------------------------------------------

            lTamanho = 20
            '086000     03  PREM-LIQ-TOT-APOL      PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            'If lIndColetiva = False Or (lIndColetiva = True And lobjstcPedLoc.Num_Item = 0) Then
                lObjstcPed.VAL_PR = lAuxiliar
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 20
                '086300     03  DESC-FRM-PAG-APOL      PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         12
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 20
                '086600     03  JUROS-FRAC-APOL        PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         12
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lObjstcPed.objstcCtrlEmis.Per_Juros_Neg = lAuxiliar
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 20
                '086900     03  CUST-APOL-APOL         PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         12
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lObjstcPed.Val_Cust_Emis = lAuxiliar
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 20
                '087200     03  VAL-IOF-APOL           PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         13
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lObjstcPed.Val_IOF = lAuxiliar
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 20
                '087500     03  VAL-TOT-APOL           PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         13
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 1
                '087800     03  FLAG-COLETIVA-APOL     PIC 9(01)      VALUE ZEROS.       12
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
            'End If
            '088000 01  FORM-COBERT.
            '088100     05  NOM-COBERT             OCCURS        150    TIMES.       12
            '088200         07  COD-COBERT         PIC X(04)      VALUE SPACES.      12
            '088300         07  HIFEN-COBERT       PIC X(01)      VALUE SPACES.      12
            '088400         07  DSC-COBERT         PIC X(50)      VALUE SPACES.      12
            '088500         07  VAL-IS-COBERT      PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         12
            '088600         07  VAL-IS-COBERT-R    REDEFINES VAL-IS-COBERT           12
            '088700                                PIC X(20).                        12
            '088800         07  PI-MES-COBERT      PIC ZZ         VALUE ZEROS.       12
            '088900         07  PI-MES-COBERT-R    REDEFINES PI-MES-COBERT           12
            '089000                                PIC X(02).
            '089100         07  QTD-DIAS-COBERT    PIC ZZ         VALUE ZEROS.       12
            '089200         07  QTD-DIAS-COBERT-R  REDEFINES QTD-DIAS-COBERT         12
            '089300                                PIC X(02).
            '089400         07  PER-FRANQ-COBERT   PIC ZZ9,99     VALUE ZEROS.       12
            '089500         07  PER-FRANQ-COBERT-R REDEFINES PER-FRANQ-COBERT        12
            '089600                                PIC X(006).                       12
            '089700         07  VAL-FRANQ-COBERT   PIC ZZZ.ZZ9,99.                   12
            '089800         07  VAL-FRANQ-COBERT-R                                   12
            '089900                                REDEFINES VAL-FRANQ-COBERT        12
            '090000                                PIC X(10).                        12
ProximoForm:
        Case "L02C0006"
            Call gM46V999.gpGraLog(0, "Processamento L02C0006")
            lPontoInicial = 1
            lTamanho = 3
            '03  COD-RAMO-ENDO          PIC 9(03)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '084100     03  COD-ENVI-ENDO          PIC X(01)      VALUE SPACES.      14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 7
            '084200     03  COD-CORR-ENDO          PIC 9(07)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '084300     03  ORDEM-ENDO             PIC 9(01)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '084400     03  NUM-APOL-ENDO          PIC 9(10)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 6
            '084500     03  NUM-ENDO-ENDO          PIC 9(06)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '084600     03  NUM-ITEM-ENDO          PIC 9(05)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            If Not lobjstcPedLoc Is Nothing Then
                If lobjstcPedLoc.Num_Item <> 0 And lobjstcPedLoc.Num_Item <> lAuxiliar Then
                    On Error Resume Next
                    lObjstcPed.ColecaoPedLoc.Add lobjstcPedLoc, lobjstcPedLoc.Class_ID
                    If Err.Number <> 0 Then
                        MsgBox ""
                    End If
                    On Error GoTo ControleErro
    
                    Set lobjstcPedLoc = New stcA46V708B
                End If
            End If
            If lobjstcPedLoc Is Nothing Then
                Set lobjstcPedLoc = New stcA46V708B
            End If
            lobjstcPedLoc.Num_Item = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '084700     03  FLAG-COL-ENDO          PIC 9(01)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 8
            '084800     03  COD-FORM-ENDO          PIC X(08)      VALUE 'L02C0006'.  15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 20
            '084900     03  NOM-RAMO-ENDO          PIC X(20)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 30
            '085000     03  NUM-SUSEP-ENDO         PIC X(30)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 6
            '085100     03  NRO-ENDO-ENDO          PIC 9(06)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '085200     03  NRO-APOL-ENDO          PIC 9(10)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Num_Apol_An = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '085300     03  NRO-ITEM-ENDO          PIC 9(05)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 15
            '085400     03  NRO-PROP-ENDO          PIC X(15)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '085500     03  DIA-INI-VIG-ENDO       PIC 9(02)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Dat_Ini_Vig = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '085600     03  BARRA-1-INI-ENDO       PIC X(01)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '085700     03  MES-INI-VIG-ENDO       PIC 9(02)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Dat_Ini_Vig = lAuxiliar & Format(lObjstcPed.Dat_Ini_Vig, "00")
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '085800     03  BARRA-2-INI-ENDO       PIC X(01)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 4
            '085900     03  ANO-INI-VIG-ENDO       PIC 9(04)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Dat_Ini_Vig = lAuxiliar & Format(lObjstcPed.Dat_Ini_Vig, "0000")
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '086000     03  DIA-TER-VIG-ENDO       PIC 9(02)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Dat_Ter_Vig = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '086100     03  BARRA-1-TER-ENDO       PIC X(01)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '086200     03  MES-TER-VIG-ENDO       PIC 9(02)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Dat_Ter_Vig = lAuxiliar & Format(lObjstcPed.Dat_Ter_Vig, "00")
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '086300     03  BARRA-2-TER-ENDO       PIC X(01)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 4
            '086400     03  ANO-TER-VIG-ENDO       PIC 9(04)      VALUE ZEROS.       15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Dat_Ter_Vig = lAuxiliar & Format(lObjstcPed.Dat_Ter_Vig, "0000")
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 60
            '086500     03  NOME-SEGUR-ENDO        PIC X(60)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Nom_Segurado = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 18
            '086600     03  CNPJ-CPF-SEGUR-ENDO    PIC X(18)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lObjstcPed.Num_Cgc_Cpf = Replace(Replace(Replace(lAuxiliar, "-", ""), ".", ""), "/", "")
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 80
            '086700     03  DSC-DIZER-1-ENDO       PIC X(80)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 80
            '086800     03  DSC-DIZER-2-ENDO       PIC X(80)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 3
            '086900     03  TIPO-LOGR-RISCO-ENDO   PIC X(03)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Cod_Tip_Logr_Risco = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 50
            '087000     03  NOME-LOGR-RISCO-ENDO   PIC X(50)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Nom_Logr_Risco = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '087100     03  NUM-LOGR-RISCO-ENDO    PIC ---------9 BLANK WHEN ZEROS.  15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco = val(lAuxiliar)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 20
            '087200     03  DSC-COMPL-RISCO-ENDO   PIC X(20)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Dsc_Compl_Risco = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 25
            '087300     03  BAIRRO-RISCO-ENDO      PIC X(25)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Nom_Bairro_Risco = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 25
            '087400     03  CIDADE-RISCO-ENDO      PIC X(25)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Nom_Cid_Risco = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '087500     03  SIG-UF-RISCO-ENDO      PIC X(02)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Sig_UF_Risco = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '087600     03  NUM-CEP-RISCO-1-ENDO   PIC X(05)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Num_Cep_Risco = IIf(Trim(lAuxiliar) = "", 0, lAuxiliar)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '087700     03  HIFEN-CEP-ENDO         PIC X(01)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 3
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Num_Cep_Risco = lobjstcPedLoc.Num_Cep_Risco & Format(lAuxiliar, "000")
            lPontoInicial = lPontoInicial + lTamanho
            '087800     03  NUM-CEP-RISCO-2-ENDO   PIC X(03)      VALUE SPACES.      15
            '--------------------------------------------------------------------------'
            lTamanho = 14
            '089000     03  COD-CLASF-ENDO         PIC X(14)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Cod_Clasf = Left(lAuxiliar, 3)
            lobjstcPedLoc.Cod_Compl_Clasf = Mid(lAuxiliar, 6, 2)

            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 60
            '089100     03  DSC-CLASF-ENDO         PIC X(60)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lobjstcPedLoc.Dsc_Clasf = lAuxiliar
            lPontoInicial = lPontoInicial + lTamanho
            For lI = 1 To 30
                Set lobjstcCobertura = New stcA46V704B
                '087900     03  NOM-COBERT-ENDO        OCCURS         20    TIMES.       15
                lTamanho = 4
                '088000         05  COD-COBERT-ENDO    PIC X(04)      VALUE SPACES.      15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lobjstcCobertura.Num_Item = lobjstcPedLoc.Num_Item
                lobjstcCobertura.Cod_Cobert = val(lAuxiliar)
                If lobjstcCobertura.Cod_Cobert = 0 Then
                    If Not gM46V999.gfPreenchido(Mid(lLinha, lPontoInicial, 85)) Then
                        Dim lA As Integer
                        Do
                            If Mid(lLinha, lPontoInicial, 1) <> " " Then
                                GoTo ContinuarForaCobertura
                            End If
                            lPontoInicial = lPontoInicial + 1
                        Loop

                    End If
                End If
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 1
                '088100         05  HIFEN-COBERT-ENDO  PIC X(01)      VALUE SPACES.      15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 50
                '088200         05  DSC-COBERT-ENDO    PIC X(50)      VALUE SPACES.      15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 20
                '088300         05  VAL-IS-ENDO        PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                If IsNumeric(lAuxiliar) Then
                    lobjstcCobertura.Val_IS = lAuxiliar
                End If

                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 10
                '088600         05  VAL-PR-LIQ-ENDO    PIC ---.--9,99.                   15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)

                If IsNumeric(lAuxiliar) Then
                    lobjstcCobertura.VAL_PR = lAuxiliar
                End If
                lPontoInicial = lPontoInicial + lTamanho

                If lobjstcCobertura.Cod_Cobert > 0 Then
                    On Error Resume Next
                    lobjstcPedLoc.ColecaoPedCobert.Add lobjstcCobertura, lobjstcCobertura.Class_ID
                    If Err.Number <> 0 Then
                        MsgBox ""
                    End If
                    On Error GoTo ControleErro
                End If
            Next lI
            '---------------------------------------------------------------------------'
ContinuarForaCobertura:
            lTamanho = 15
            '089200     03  NOM-CORR-1-ENDO        PIC X(15)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            Set lObjstcPed.ColecaoPedCorr = New Collection
            Set lObjstcPedCorr = New stcA46V705B
            lObjstcPedCorr.ComparaCorretor = lAuxiliar
            lObjstcPedCorr.Cod_Corr = 1
            On Error Resume Next
            lObjstcPed.ColecaoPedCorr.Add lObjstcPedCorr, lObjstcPedCorr.Class_ID
            If Err.Number <> 0 Then
                MsgBox ""
            End If
            On Error GoTo ControleErro

            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '089300     03  ABRE-PAR-1-ENDO        PIC X(01)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 14
            '089400     03  COD-CORR-1-ENDO        PIC 9(14).                        15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '089700     03  FECHA-PAR-1-ENDO       PIC X(01)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 15
            '089800     03  NOM-CORR-2-ENDO        PIC X(15)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            Set lObjstcPedCorr = New stcA46V705B
            lObjstcPedCorr.ComparaCorretor = lAuxiliar
            lObjstcPedCorr.Cod_Corr = 2
            If gM46V999.gfPreenchido(lObjstcPedCorr.ComparaCorretor) Then
                On Error Resume Next
                lObjstcPed.ColecaoPedCorr.Add lObjstcPedCorr, lObjstcPedCorr.Class_ID
                If Err.Number <> 0 Then
                    MsgBox ""
                End If
                On Error GoTo ControleErro
            End If

            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '089900     03  ABRE-PAR-2-ENDO        PIC X(01)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 14
            '090000     03  COD-CORR-2-ENDO        PIC 9(14).                        15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '090300     03  FECHA-PAR-2-ENDO       PIC X(01)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 15
            '090400     03  NOM-CORR-3-ENDO        PIC X(15)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            Set lObjstcPedCorr = New stcA46V705B
            lObjstcPedCorr.ComparaCorretor = lAuxiliar
            lObjstcPedCorr.Cod_Corr = 3
            If gM46V999.gfPreenchido(lObjstcPedCorr.ComparaCorretor) Then
                On Error Resume Next
                lObjstcPed.ColecaoPedCorr.Add lObjstcPedCorr, lObjstcPedCorr.Class_ID
                If Err.Number <> 0 Then
                    MsgBox ""
                End If
                On Error GoTo ControleErro
            End If
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '090500     03  ABRE-PAR-3-ENDO        PIC X(01)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 14
            '090600     03  COD-CORR-3-ENDO        PIC 9(14).                        15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '090900     03  FECHA-PAR-3-ENDO       PIC X(01)      VALUE SPACES.      15
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 20
            '091000     03  PREM-LIQ-TOT-ENDO      PIC -.---.---.---.--9,99.         15
            'If lIndColetiva = False Or (lIndColetiva = True And lobjstcPedLoc.Num_Item = 0) Then
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lObjstcPed.VAL_PR = lAuxiliar
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 20
                '091300     03  DESC-FRM-PAG-ENDO      PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 5
                '091600     03  PER-JUROS-FRAC-ENDO    PIC Z9,99.                        15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lObjstcPed.objstcCtrlEmis.Per_Juros_Neg = lAuxiliar
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 20
                '091900     03  JUROS-FRAC-ENDO        PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 20
                '092200     03  CUST-APOL-ENDO         PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lObjstcPed.Val_Cust_Emis = lAuxiliar
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 20
                '092500     03  VAL-IOF-ENDO           PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lObjstcPed.Val_IOF = lAuxiliar
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 20
                '092800     03  VAL-TOT-ENDO           PIC -.---.---.---.--9,99.         15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 30
                '093100     03  DSC-RODAPE-1-ENDO      PIC X(30)      VALUE SPACES.      15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 30
                '093200     03  DSC-RODAPE-2-ENDO      PIC X(30)      VALUE SPACES.      15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 30
                '093300     03  LOCAL-ENDO             PIC X(30)      VALUE SPACES.      15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 2
                '093400     03  DIA-EMIS-ENDO          PIC 9(02)      VALUE ZEROS.       15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 1
                '093500     03  BARRA-1-EMIS-ENDO      PIC X(01)      VALUE SPACES.      15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 2
                '093600     03  MES-EMIS-ENDO          PIC 9(02)      VALUE ZEROS.       15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 1
                '093700     03  BARRA-2-EMIS-ENDO      PIC X(01)      VALUE SPACES.      15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 4
                '093800     03  ANO-EMIS-ENDO          PIC 9(04)      VALUE ZEROS.       15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 20
                '093900     03  VAL-RISC-ENDO          PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99          15
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                If Not IsNumeric(lAuxiliar) Then
                    lobjstcPedLoc.Val_Vr = val(lAuxiliar)
                Else
                    lobjstcPedLoc.Val_Vr = lAuxiliar
                End If

            'End If
        Case "L02C0007", "L02C0008"
            Call gM46V999.gpGraLog(0, "Processamento L02C0007, L02C0008")
            '090100 01  FORM-CLAU-WS.
            lPontoInicial = 1
            lTamanho = 3
            '090200     03  COD-RAMO-CLAU          PIC 9(03)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '090300     03  COD-ENVI-CLAU          PIC X(01)      VALUE SPACES.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 7
            '090400     03  COD-CORR-CLAU          PIC 9(07)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '090500     03  ORDEM-CLAU             PIC 9(01)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '090600     03  NUM-APOL-CLAU          PIC 9(10)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 6
            '090700     03  NUM-ENDO-CLAU          PIC 9(06)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '090800     03  NUM-ITEM-CLAU          PIC 9(05)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '090900     03  FLAG-COL-CLAU          PIC 9(01)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 8
            '091000     03  COD-FORM-CLAU          PIC X(08)      VALUE 'L02C0007'.  13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            '-------------------------------------------------------------------------'
            If lobjstcPedLoc.ColecaoPedClau.Count > 0 Then
                Set lobjstcPedLoc.ColecaoPedClau = New Collection
            End If
            If lFormulario = "L02C0008" Then
                For lI = 0 To 26
                    Set lobjstcPedClau = New stcA46V703B
                    '091100     03  NOM-CLAUSULA-CLAU      OCCURS         20    TIMES.       13
                    lobjstcPedClau.Num_Item = lobjstcPedLoc.Num_Item
                    lTamanho = 3
                    '091200         05  COD-CLAUSULA-CLAU  PIC X(03)      VALUE SPACES.      13
                    lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                    lPontoInicial = lPontoInicial + lTamanho
                    lobjstcPedClau.Cod_Clau = val(lAuxiliar)
                    lTamanho = 1
                    '091300         05  TRACO-CLAU         PIC X(01)      VALUE SPACES.      13
                    lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                    lPontoInicial = lPontoInicial + lTamanho
                    lTamanho = 75
                    '091400         05  DSC-CLAUSULA-CLAU  PIC X(75)      VALUE SPACES.      13
                    lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                    lPontoInicial = lPontoInicial + lTamanho
                    If lobjstcPedClau.Cod_Clau > 0 And lobjstcPedClau.Cod_Clau <> 50 Then
                        On Error Resume Next
                        lobjstcPedLoc.ColecaoPedClau.Add lobjstcPedClau, lobjstcPedClau.Class_ID
                        If Err.Number <> 0 Then
                            MsgBox ""
                        End If
                        On Error GoTo ControleErro
                    End If
                Next
            Else
                For lI = 0 To 19
                    Set lobjstcPedClau = New stcA46V703B
                    '091100     03  NOM-CLAUSULA-CLAU      OCCURS         20    TIMES.       13
                    lobjstcPedClau.Num_Item = lobjstcPedLoc.Num_Item
                    lTamanho = 3
                    '091200         05  COD-CLAUSULA-CLAU  PIC X(03)      VALUE SPACES.      13
                    lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                    lPontoInicial = lPontoInicial + lTamanho
                    lobjstcPedClau.Cod_Clau = val(lAuxiliar)
                    lTamanho = 1
                    '091300         05  TRACO-CLAU         PIC X(01)      VALUE SPACES.      13
                    lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                    lPontoInicial = lPontoInicial + lTamanho
                    lTamanho = 75
                    '091400         05  DSC-CLAUSULA-CLAU  PIC X(75)      VALUE SPACES.      13
                    lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                    lPontoInicial = lPontoInicial + lTamanho
                    If lobjstcPedClau.Cod_Clau > 0 And lobjstcPedClau.Cod_Clau <> 50 Then
                        On Error Resume Next
                        lobjstcPedLoc.ColecaoPedClau.Add lobjstcPedClau, lobjstcPedClau.Class_ID
                        If Err.Number <> 0 Then
                            MsgBox ""
                        End If
                        On Error GoTo ControleErro
                    End If
                Next
            End If
            '--------------------------------------------------------------------------'

            '--------------------------------------------------------------------------'
            For lI = 1 To 5
                '091500     03  OUTROS-SEGUROS-CLAU    OCCURS         5     TIMES.       13
                lTamanho = 35
                '091600         05  DSC-OUTROS-CLAU    PIC X(35)      VALUE SPACES.      13
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
            Next lI
            '--------------------------------------------------------------------------'
            lTamanho = 30
            '091700     03  LOCAL-CLAU             PIC X(30)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '091800     03  DIA-EMIS-CLAU          PIC 9(02)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '091900     03  BARRA-1-EMIS-CLAU      PIC X(01)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '092000     03  MES-EMIS-CLAU          PIC 9(02)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '092100     03  BARRA-2-EMIS-CLAU      PIC X(01)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 4
            '092200     03  ANO-EMIS-CLAU          PIC 9(04)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho

        Case "L02C0011"
            Call gM46V999.gpGraLog(0, "Processamento L02C0011")
            '092600 01  FORM-PROT-WS.                                                13
            lPontoInicial = 1
            lTamanho = 3
            '092700     03  COD-RAMO-PROT          PIC 9(03)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '092800     03  COD-ENVI-PROT          PIC X(01)      VALUE SPACES.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 7
            '092900     03  COD-CORR-PROT          PIC 9(07)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '093000     03  ORDEM-PROT             PIC 9(01)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '093100     03  NUM-APOL-PROT          PIC 9(10)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 6
            '093200     03  NUM-ENDO-PROT          PIC 9(06)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '093300     03  NUM-ITEM-PROT          PIC 9(05)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '093400     03  FLAG-COL-PROT          PIC 9(01)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 8
            '093500     03  COD-FORM-PROT          PIC X(08)      VALUE 'L02C0011'.  13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '093600     03  NRO-APOL-PROT          PIC 9(10)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '093700     03  HIFEN-DIG-PROT         PIC X(01)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '093800     03  DIG-APOL-PROT          PIC 9(01)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 6
            '093900     03  NRO-ENDO-PROT          PIC 9(06)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '094000     03  NRO-ITEM-PROT          PIC 9(05)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 60
            '094100     03  NOME-SEGUR-PROT        PIC X(60)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 3
            '094200     03  TIPO-LOGR-RISCO-PROT   PIC X(03)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 50
            '094300     03  NOME-LOGR-RISCO-PROT   PIC X(50)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '094400     03  NUM-LOGR-RISCO-PROT    PIC ---------9 BLANK WHEN ZEROS.  13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 20
            '094500     03  DSC-COMPL-RISCO-PROT   PIC X(20)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 25
            '094600     03  BAIRRO-RISCO-PROT      PIC X(25)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 25
            '094700     03  CIDADE-RISCO-PROT      PIC X(25)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '094800     03  SIG-UF-RISCO-PROT      PIC X(02)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '094900     03  NUM-CEP-RISCO-1-PROT   PIC X(05)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '095000     03  HIFEN-CEP-PROT         PIC X(01)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 3
            '095100     03  NUM-CEP-RISCO-2-PROT   PIC X(03)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            '--------------------------------------------------------------------------

            For lI = 0 To 44
                Set lobjstcFator = Nothing
                '095200     03  NOM-PROTEC-PROT        OCCURS         45    TIMES.       13
                lTamanho = 1
                '095300         05  ABRE-FATOR-PROT    PIC X(01)      VALUE SPACES.      13
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)

                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 1
                '095400         05  TEM-PROTEC-PROT    PIC X(01)      VALUE SPACES.      13
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                If gM46V999.gfPreenchido(lAuxiliar) Then
                    Set lobjstcFator = New stcA46V720B
                    lobjstcFator.Num_Item = lobjstcPedLoc.Num_Item
                End If
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 2
                '095500         05  FECHA-FATOR-PROT   PIC X(02)      VALUE SPACES.      13
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 50
                '095600         05  DSC-PROTEC-PROT    PIC X(50)      VALUE SPACES.      13
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                If Not lobjstcFator Is Nothing Then
                    lobjstcFator.Desc_Fator = lAuxiliar
                    lobjstcFator.Cod_Fator = lI
                End If
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 41
                '095700         05  FILLER             PIC X(41)      VALUE SPACES.      13
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 1
                '095800         05  HIFEN-FATOR-PROT   PIC X(01)      VALUE SPACES.      13
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 1
                '095900         05  FILLER             PIC X(01)      VALUE SPACES.      13
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 18
                '096000         05  DSC-COMPL-PROT     PIC X(18)      VALUE SPACES.      13
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                If Not lobjstcFator Is Nothing Then
                    On Error Resume Next
                    lobjstcPedLoc.ColecaoPedFator.Add lobjstcFator, lobjstcFator.Class_ID
                    If Err.Number <> 0 Then
                        MsgBox ""
                    End If
                    On Error GoTo ControleErro
                End If

            Next
            '--------------------------------------------------------------------------'
        Case "L02C0013"
            Call gM46V999.gpGraLog(0, "Processamento L02C0013")

            '                096200 01  FORM-BENS-WS.
            lPontoInicial = 1
            lTamanho = 3
            '096300     03  COD-RAMO-BEM           PIC 9(03)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '096400     03  COD-ENVI-BEM           PIC X(01)      VALUE SPACES.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 7
            '096500     03  COD-CORR-BEM           PIC 9(07)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '096600     03  ORDEM-BEM              PIC 9(01)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '096700     03  NUM-APOL-BEM           PIC 9(10)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 6
            '096800     03  NUM-ENDO-BEM           PIC 9(06)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '096900     03  NUM-ITEM-BEM           PIC 9(05)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '097000     03  FLAG-COL-BEM           PIC 9(01)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 8
            '097100     03  COD-FORM-BEM           PIC X(08)      VALUE 'L02C0013'.  13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '097200     03  NRO-APOL-BEM           PIC 9(10)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 6
            '097300     03  NRO-ENDO-BEM           PIC 9(06)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '097400     03  NRO-ITEM-BEM           PIC 9(05)      VALUE ZEROS.       13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 60
            '097500     03  NOME-SEGUR-BEM         PIC X(60)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 3
            '097600     03  TIPO-LOGR-RISCO-BEM    PIC X(03)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 50
            '097700     03  NOME-LOGR-RISCO-BEM    PIC X(50)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '097800     03  NUM-LOGR-RISCO-BEM     PIC ---------9 BLANK WHEN ZEROS.  13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 20
            '097900     03  DSC-COMPL-RISCO-BEM    PIC X(20)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 25
            '098000     03  BAIRRO-RISCO-BEM       PIC X(25)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 25
            '098100     03  CIDADE-RISCO-BEM       PIC X(25)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '098200     03  SIG-UF-RISCO-BEM       PIC X(02)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '098300     03  NUM-CEP-RISCO-1-BEM    PIC X(05)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '098400     03  HIFEN-CEP-BEM          PIC X(01)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 3
            '098500     03  NUM-CEP-RISCO-2-BEM    PIC X(03)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            '--------------------------------------------------------------------------'
            '098600     03  NOM-BENS-BEM           OCCURS         45    TIMES.       13
            lTamanho = 50
            '098700         05  DSC-BENS-BEM       PIC X(50)      VALUE SPACES.      13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 20
            '098800         05  VAL-BENS-BEM       PIC Z.ZZZ.ZZZ.ZZZ.ZZ9,99.         13
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            '--------------------------------------------------------------------------'
        Case "L02C0015"
            Call gM46V999.gpGraLog(0, "Processamento L02C0015")

            lPontoInicial = 1
            lTamanho = 3
            '03  COD-RAMO-FRANQ         PIC 9(03)      VALUE ZEROS.       16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '102700     03  COD-ENVI-FRANQ         PIC X(01)      VALUE SPACES.      16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 7
            '102800     03  COD-CORR-FRANQ         PIC 9(07)      VALUE ZEROS.       16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '102900     03  ORDEM-FRANQ            PIC 9(01)      VALUE ZEROS.       16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '103000     03  NUM-APOL-FRANQ         PIC 9(10)      VALUE ZEROS.       16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 6
            '103100     03  NUM-ENDO-FRANQ         PIC 9(06)      VALUE ZEROS.       16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '103200     03  NUM-ITEM-FRANQ         PIC 9(05)      VALUE ZEROS.       16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '103300     03  FLAG-COL-FRANQ         PIC 9(01)      VALUE ZEROS.       16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 8
            '103400     03  COD-FORM-FRANQ         PIC X(08)      VALUE 'L02C0015'.  16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '103500     03  NRO-APOL-FRANQ         PIC 9(10)      VALUE ZEROS.       16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 6
            '103600     03  NRO-ENDO-FRANQ         PIC 9(06)      VALUE ZEROS.       16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '103700     03  NRO-ITEM-FRANQ         PIC 9(05)      VALUE ZEROS.       16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 60
            '103800     03  NOME-SEGUR-FRANQ       PIC X(60)      VALUE SPACES.      16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 3
            '103900     03  TIPO-LOGR-RISCO-FRANQ  PIC X(03)      VALUE SPACES.      16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 50
            '104000     03  NOME-LOGR-RISCO-FRANQ  PIC X(50)      VALUE SPACES.      16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 10
            '104100     03  NUM-LOGR-RISCO-FRANQ   PIC ---------9 BLANK WHEN ZEROS.  16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 20
            '104200     03  DSC-COMPL-RISCO-FRANQ  PIC X(20)      VALUE SPACES.      16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 25
            '104300     03  BAIRRO-RISCO-FRANQ     PIC X(25)      VALUE SPACES.      16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 25
            '104400     03  CIDADE-RISCO-FRANQ     PIC X(25)      VALUE SPACES.      16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2
            '104500     03  SIG-UF-RISCO-FRANQ     PIC X(02)      VALUE SPACES.      16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '104600     03  NUM-CEP-RISCO-1-FRANQ  PIC X(05)      VALUE SPACES.      16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '104700     03  HIFEN-CEP-FRANQ        PIC X(01)      VALUE SPACES.      16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 3
            '104800     03  NUM-CEP-RISCO-2-FRANQ  PIC X(03)      VALUE SPACES.      16
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            '--------------------------------------------------------------------------'
            Set ColecaoCoberturaTMP = New Collection
            For lI = 1 To 25
                '104900     03  NOM-FRANQUIAS-FRANQ    OCCURS         25    TIMES.       16
                lTamanho = 4
                '105000         05 COD-COBERT-FRANQ    PIC X(04)      VALUE SPACES.
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                For Each lobjstcCobertura In lobjstcPedLoc.ColecaoPedCobert
                    If lobjstcCobertura.Cod_Cobert = val(lAuxiliar) Then
                        GoTo carregarInfoFranq
                    End If
                Next
carregarInfoFranq:
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 1
                '105100         05 HIFEN-COBERT-FRANQ  PIC X(01)      VALUE SPACES.
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 14
                '105200         05 DSC-COBERT-FRANQ    PIC X(14)      VALUE SPACES.      16
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 1
                '105300         05 FILLER              PIC X(01)      VALUE SPACES.
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 5
                '105400         05 PER-FRANQ           PIC Z9,99.
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                If Not lobjstcCobertura Is Nothing Then
                    lobjstcCobertura.Per_Franq = lAuxiliar
                End If
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 28
                '105700         05 TEXTO1-FRANQ        PIC X(30)      VALUE SPACES.
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 10
                '105800         05 VAL-MIN-FRANQ       PIC Z.ZZ9,99.                     16
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                If Not lobjstcCobertura Is Nothing Then
                    lobjstcCobertura.Val_Franq_Min = lAuxiliar
                End If
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 1
                '106100         05 FILLER              PIC X(01)      VALUE SPACES.
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 8
                '106200         05 TEXTO2-FRANQ        PIC X(08)      VALUE SPACES.
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 2
                '106300         05 QTD-DIAS-FRANQ      PIC 99.
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                If Not lobjstcCobertura Is Nothing Then
                    lobjstcCobertura.Num_Meses_PI = val(lAuxiliar)
                End If
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 1
                '106600         05 FILLER              PIC X(01)      VALUE SPACES.
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 5
                '106700         05 TEXTO3-FRANQ        PIC X(05)      VALUE SPACES.
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                lPontoInicial = lPontoInicial + lTamanho
                lTamanho = 1811
                '106800     03  FILLER                 PIC X(1811)    VALUE SPACES.      16
                lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
                If Not lobjstcCobertura Is Nothing Then
                    On Error Resume Next
                    ColecaoCoberturaTMP.Add lobjstcCobertura, lobjstcCobertura.Class_ID
                    If Err.Number <> 0 Then
                        MsgBox ""
                    End If
                    On Error GoTo ControleErro
                End If
            Next

            If ColecaoCoberturaTMP.Count = lobjstcPedLoc.ColecaoPedCobert.Count Then
                Set lobjstcPedLoc.ColecaoPedCobert = New Collection
                Set lobjstcPedLoc.ColecaoPedCobert = ColecaoCoberturaTMP
            End If
        Case "L02C0025"
            Call gM46V999.gpGraLog(0, "Processamento L02C0025")

            '099200 01  FORM-PRAZO-WS.                                               14
            lPontoInicial = 1
            lTamanho = 3
            '099300     03  COD-RAMO-PRAZ          PIC 9(03)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 201
            '099400     03  COD-ENVI-PRAZ          PIC X(01)      VALUE SPACES.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 207
            '099500     03  COD-CORR-PRAZ          PIC 9(07)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 201
            '099600     03  ORDEM-PRAZ             PIC 9(01)      VALUE ZEROS.
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 2010
            '099700     03  NUM-APOL-PRAZ          PIC 9(10)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 206
            '099800     03  NUM-ENDO-PRAZ          PIC 9(06)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 5
            '099900     03  NUM-ITEM-PRAZ          PIC 9(05)      VALUE ZEROS.       14
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 1
            '100000     03  FLAG-COL-PRAZ          PIC 9(01)      VALUE ZEROS.       12
            lAuxiliar = Mid(lLinha, lPontoInicial, lTamanho)
            lPontoInicial = lPontoInicial + lTamanho
            lTamanho = 8
            '100100     03  COD-FORM-PRAZ          PIC X(08)      VALUE 'L02C0025'.  14
        Case "L02C0017"    '??
        Case "L02C0021"    '??
        Case "L02C0027", "L02C0029", "L02C0031"    'assist
        Case "L02C0033"    'clau 50

        End Select

        '        If lFormulario = "L02C0040" Or lFormulario = "L02C0029" Then
        If lFormulario = "L02C0147" Then
            Call gM46V999.gpGraLog(0, "Tratamento do formulario L02C0147")
            On Error Resume Next
            If Not lobjstcPedLoc Is Nothing Then
                Call gM46V999.gpGraLog(0, "Adiciona lobjstcPedLoc")
                lObjstcPed.ColecaoPedLoc.Add lobjstcPedLoc, lobjstcPedLoc.Class_ID
            End If
            If Err.Number <> 0 Then
                MsgBox ""
            End If
            On Error GoTo ControleErro


            Set lObjPed = New clsA46V702A
            Set lObjPed.M46V999 = gM46V999
            Set lObjstcPedOrig = New stcA46V702B
            lObjstcPedOrig.Num_Apol = lNumApol
            lObjstcPedOrig.Num_Endo = lNumEndo
            If Not lObjPed.mfCarregarColecao(lObjstcPedOrig, "", "", True) Then
                Call gM46V999.gpGraLog(1, lObjPed.UltimoErro)
                Exit Sub
            End If
            lMsg = ""
            If lObjPed.Colecao.Count = 0 Then
                If Not lObjPed.mfCarregarColecao(lObjstcPedOrig, "GED.DBO.", "", True) Then
                    Call gM46V999.gpGraLog(1, lObjPed.UltimoErro)
                    Exit Sub
                End If
                GoTo Comparar
            Else

Comparar:
                Set lObjstcPedOrig = lObjPed.Colecao.Item(1)
                'If lObjstcPedOrig.NOSSO_NUMERO <> lObjstcPed.NOSSO_NUMERO Then
                If lObjstcPedOrig.Cod_Ramo <> lObjstcPed.Cod_Ramo Then
                    lMsg = lMsg & "Ramo: " & lObjstcPedOrig.Cod_Ramo & "/" & lObjstcPed.Cod_Ramo
                End If
                'If lObjstcPedOrig.Cod_Org <> lObjstcPed.Cod_Org Then
                'If lObjstcPedOrig.Num_Ped <> lObjstcPed.Num_Ped Then
                'If lObjstcPedOrig.Num_Apol <> lObjstcPed.Num_Apol Then
                'If lObjstcPedOrig.Num_Endo <> lObjstcPed.Num_Endo Then
                If lObjstcPedOrig.Num_Apol_An <> lObjstcPed.Num_Apol_An Then
                    lMsg = lMsg & "Apólice anterior: " & lObjstcPedOrig.Num_Apol_An & "/" & lObjstcPed.Num_Apol_An
                End If
                'If lObjstcPedOrig.Cod_Reg <> lObjstcPed.Cod_Reg Then
                'If lObjstcPedOrig.Cod_Bco_Cobr <> lObjstcPed.Cod_Bco_Cobr Then
                'If lObjstcPedOrig.Cod_Agen_Angar <> lObjstcPed.Cod_Agen_Angar Then
                'If lObjstcPedOrig.Cod_Agen_Indicr <> lObjstcPed.Cod_Agen_Indicr Then
                'If lObjstcPedOrig.Num_Cta_Indicr <> lObjstcPed.Num_Cta_Indicr Then
                If lObjstcPedOrig.Cod_Unid_Prod <> lObjstcPed.Cod_Unid_Prod And lObjstcPed.Cod_Unid_Prod <> 0 Then
                    lMsg = lMsg & "Unidade prod: " & lObjstcPedOrig.Cod_Unid_Prod & "/" & lObjstcPed.Cod_Unid_Prod
                End If
                If lObjstcPedOrig.Cod_Prod <> lObjstcPed.Cod_Prod And lObjstcPed.Cod_Prod <> 0 Then
                    lMsg = lMsg & "Produtor: " & lObjstcPedOrig.Cod_Prod & "/" & lObjstcPed.Cod_Prod
                End If
                'If lObjstcPedOrig.Cod_Agen_Bas_Cor <> lObjstcPed.Cod_Agen_Bas_Cor Then
                'If lObjstcPedOrig.Cod_Dig_Agen_Bas_Cor <> lObjstcPed.Cod_Dig_Agen_Bas_Cor Then
                'If lObjstcPedOrig.Num_Cta_Bas_Cor <> lObjstcPed.Num_Cta_Bas_Cor Then
                'If lObjstcPedOrig.mNum_Cta_Bas_Cor_Compl <> lObjstcPed.mNum_Cta_Bas_Cor_Compl Then
                'If lObjstcPedOrig.Cod_Segurado <> lObjstcPed.Cod_Segurado Then
                'If lObjstcPedOrig.Dig_Cod_Segurado <> lObjstcPed.Dig_Cod_Segurado Then
                If lObjstcPedOrig.Nom_Segurado <> Trim(lObjstcPed.Nom_Segurado) Then
                    lMsg = lMsg & "Nome do segurado: " & lObjstcPedOrig.Nom_Segurado & "/" & lObjstcPed.Nom_Segurado
                End If
                'If lObjstcPedOrig.Dsc_Compl_Segurado <> lObjstcPed.Dsc_Compl_Segurado Then
                If lObjstcPedOrig.Num_Cgc_Cpf <> lObjstcPed.Num_Cgc_Cpf Then
                    lMsg = lMsg & "CPF/CNPJ : " & lObjstcPedOrig.Num_Cgc_Cpf & "/" & lObjstcPed.Num_Cgc_Cpf
                End If
                If lObjstcPedOrig.Cod_Tip_Logr <> Trim(lObjstcPed.Cod_Tip_Logr) And gM46V999.gfPreenchido(lObjstcPed.Cod_Tip_Logr) Then
                    lMsg = lMsg & "Tip Logr: " & lObjstcPedOrig.Cod_Tip_Logr & "/" & lObjstcPed.Cod_Tip_Logr
                End If
                If lObjstcPedOrig.Nom_Logr <> Trim(lObjstcPed.Nom_Logr) And gM46V999.gfPreenchido(lObjstcPed.Nom_Logr) Then
                    lMsg = lMsg & "Nome Logr: " & lObjstcPedOrig.Nom_Logr & "/" & lObjstcPed.Nom_Logr
                End If
                If lObjstcPedOrig.Dsc_Compl_Logr <> Trim(lObjstcPed.Dsc_Compl_Logr) And gM46V999.gfPreenchido(lObjstcPed.Dsc_Compl_Logr) Then
                    lMsg = lMsg & "Compl Logr: " & lObjstcPedOrig.Dsc_Compl_Logr & "/" & lObjstcPed.Dsc_Compl_Logr
                End If
                If Left(lObjstcPedOrig.Nom_Bairro, 25) <> Trim(lObjstcPed.Nom_Bairro) And gM46V999.gfPreenchido(lObjstcPed.Nom_Bairro) Then
                    lMsg = lMsg & "Bairro : " & lObjstcPedOrig.Nom_Bairro & "/" & lObjstcPed.Nom_Bairro
                End If
                If lObjstcPedOrig.Nom_Cid <> Trim(lObjstcPed.Nom_Cid) And gM46V999.gfPreenchido(lObjstcPed.Nom_Cid) Then
                    lMsg = lMsg & "Cidade : " & lObjstcPedOrig.Nom_Cid & "/" & lObjstcPed.Nom_Cid
                End If
                If lObjstcPedOrig.Sig_UF <> lObjstcPed.Sig_UF And gM46V999.gfPreenchido(lObjstcPed.Sig_UF) Then
                    lMsg = lMsg & "UF : " & lObjstcPedOrig.Sig_UF & "/" & lObjstcPed.Sig_UF
                End If
                If lObjstcPedOrig.Num_CEP <> lObjstcPed.Num_CEP Then
                    '    lMsg = lMsg & "CEP: " & lObjstcPedOrig.Num_CEP & "/" & lObjstcPed.Num_CEP
                End If
                'If lObjstcPedOrig.Num_Fone <> lObjstcPed.Num_Fone Then
                'If lObjstcPedOrig.Cod_Estip <> lObjstcPed.Cod_Estip Then
                'If lObjstcPedOrig.Val_Coef_Cms_Pad <> lObjstcPed.Val_Coef_Cms_Pad Then
                If lObjstcPedOrig.Dat_Ini_Vig <> lObjstcPed.Dat_Ini_Vig Then
                    lMsg = lMsg & "Inicio vigência: " & lObjstcPedOrig.Dat_Ini_Vig & "/" & lObjstcPed.Dat_Ini_Vig
                End If
                If lObjstcPedOrig.Dat_Ter_Vig <> lObjstcPed.Dat_Ter_Vig Then
                    lMsg = lMsg & "Término vigência: " & lObjstcPedOrig.Dat_Ter_Vig & "/" & lObjstcPed.Dat_Ter_Vig
                End If
                'If lObjstcPedOrig.Dat_Emis <> lObjstcPed.Dat_Emis Then
                'If lObjstcPedOrig.Cod_Moed <> lObjstcPed.Cod_Moed Then
'                If lObjstcPedOrig.Per_Juro_Frac <> lObjstcPed.Per_Juro_Frac Then
'                    lMsg = lMsg & "Juros : " & lObjstcPedOrig.Per_Juro_Frac & "/" & lObjstcPed.Per_Juro_Frac
'                End If
                'If lObjstcPedOrig.Qtd_Parc <> lObjstcPed.Qtd_Parc Then
                'If lObjstcPedOrig.Qtd_Anexo <> lObjstcPed.Qtd_Anexo Then
                'If lObjstcPedOrig.Cod_Iden_Praz <> lObjstcPed.Cod_Iden_Praz Then
                'If lObjstcPedOrig.Cod_Pr_Min_Endo <> lObjstcPed.Cod_Pr_Min_Endo Then
                'If lObjstcPedOrig.Cod_Pr_Min_Parc <> lObjstcPed.Cod_Pr_Min_Parc Then
                'If lObjstcPedOrig.Val_Pr_Antec <> lObjstcPed.Val_Pr_Antec Then
                'If lObjstcPedOrig.Per_Adic_Form_Pag <> lObjstcPed.Per_Adic_Form_Pag Then
                'If lObjstcPedOrig.Per_Desc_Form_Pag <> lObjstcPed.Per_Desc_Form_Pag Then
                'If lObjstcPedOrig.Val_IS <> lObjstcPed.Val_IS Then
                'If lObjstcPedOrig.Val_IS_2 <> lObjstcPed.Val_IS_2 Then
                If lObjstcPedOrig.VAL_PR <> lObjstcPed.VAL_PR Then
                    'lMsg = lMsg & "Premio líquido: " & lObjstcPedOrig.VAL_PR & "/" & lObjstcPed.VAL_PR
                End If
                'If lObjstcPedOrig.Val_Pr_Ref_Tarif <> lObjstcPed.Val_Pr_Ref_Tarif Then
                'If lObjstcPedOrig.Val_Pr_Ref_Resg <> lObjstcPed.Val_Pr_Ref_Resg Then
                'If lObjstcPedOrig.Dat_Ini_Viag <> lObjstcPed.Dat_Ini_Viag Then
                'If lObjstcPedOrig.Cod_Tip_Liber <> lObjstcPed.Cod_Tip_Liber Then
                'If lObjstcPedOrig.Cod_Iden_Alt_Corr <> lObjstcPed.Cod_Iden_Alt_Corr Then
                'If lObjstcPedOrig.Cod_User_Digit <> lObjstcPed.Cod_User_Digit Then
                'If lObjstcPedOrig.DAT_DIGIT <> lObjstcPed.DAT_DIGIT Then
                'If lObjstcPedOrig.HOR_DIGIT <> lObjstcPed.HOR_DIGIT Then
                'If lObjstcPedOrig.Cod_User_Liber <> lObjstcPed.Cod_User_Liber Then
                'If lObjstcPedOrig.Dat_Liber <> lObjstcPed.Dat_Liber Then
                'If lObjstcPedOrig.Hor_Liber <> lObjstcPed.Hor_Liber Then
                'If lObjstcPedOrig.Cod_Sit_Ped <> lObjstcPed.Cod_Sit_Ped Then
                'If lObjstcPedOrig.Cod_Tip_Doc <> lObjstcPed.Cod_Tip_Doc Then
                'If lObjstcPedOrig.Cod_Dizer <> lObjstcPed.Cod_Dizer Then
                'If lObjstcPedOrig.Qtd_Copia <> lObjstcPed.Qtd_Copia Then
                If lObjstcPedOrig.Val_Cust_Emis <> lObjstcPed.Val_Cust_Emis Then
                    lMsg = lMsg & "Custo de emissão: " & lObjstcPedOrig.Val_Cust_Emis & "/" & lObjstcPed.Val_Cust_Emis
                End If
                'If lObjstcPedOrig.Val_Parc_Min <> lObjstcPed.Val_Parc_Min Then
                'If lObjstcPedOrig.Val_Juro <> lObjstcPed.Val_Juro Then
                If lIndColetiva = False Or (lIndColetiva = True And lobjstcPedLoc.Num_Item = 0) Then
                    If Round(lObjstcPedOrig.Val_IOF, 0) <> Round(lObjstcPed.Val_IOF, 0) And lObjstcPedOrig.Val_IOF <> 0 Then
                        lMsg = lMsg & "IOF : " & lObjstcPedOrig.Val_IOF & "/" & lObjstcPed.Val_IOF
                    End If
                End If
                'If lObjstcPedOrig.Num_Prop <> lObjstcPed.Num_Prop Then
                'If lObjstcPedOrig.Cod_Posi_Atual <> lObjstcPed.Cod_Posi_Atual Then
                'If lObjstcPedOrig.Dat_Mov_Ini <> lObjstcPed.Dat_Mov_Ini Then
                'If lObjstcPedOrig.Dat_Mov_Ter <> lObjstcPed.Dat_Mov_Ter Then
                'If lObjstcPedOrig.Tip_Cust_Emis <> lObjstcPed.Tip_Cust_Emis Then
                'If lObjstcPedOrig.Val_Cambio_Emis <> lObjstcPed.Val_Cambio_Emis Then
                'If lObjstcPedOrig.Iden <> lObjstcPed.Iden Then
                'If lObjstcPedOrig.Num_PI <> lObjstcPed.Num_PI Then
                'If lObjstcPedOrig.NossoNumeroIn <> lObjstcPed.NossoNumeroIn Then
                'If lObjstcPedOrig.Cod_Produto <> lObjstcPed.Cod_Produto Then
                'If lObjstcPedOrig.Cod_Protocolo <> lObjstcPed.Cod_Protocolo Then
                'If lObjstcPedOrig.Cod_Sequencia <> lObjstcPed.Cod_Sequencia Then
                'If lObjstcPedOrig.Tip_Pessoa <> lObjstcPed.Tip_Pessoa Then
                'If lObjstcPedOrig.Cod_Ramo_Ativ <> lObjstcPed.Cod_Ramo_Ativ Then
                'If lObjstcPedOrig.Cod_Doc <> lObjstcPed.Cod_Doc Then
                'If lObjstcPedOrig.Num_Doc <> lObjstcPed.Num_Doc Then
                'If lObjstcPedOrig.Org_Exp <> lObjstcPed.Org_Exp Then
                'If lObjstcPedOrig.Dat_Emis_Doc <> lObjstcPed.Dat_Emis_Doc Then
                'If lObjstcPedOrig.Nom_Empresa <> lObjstcPed.Nom_Empresa Then
                'If lObjstcPedOrig.Sexo <> lObjstcPed.Sexo Then
                'If lObjstcPedOrig.Dat_Pri_Calculo <> lObjstcPed.Dat_Pri_Calculo Then
                'If lObjstcPedOrig.Ult_Tarifa <> lObjstcPed.Ult_Tarifa Then
                'If lObjstcPedOrig.Cod_Unicidade <> lObjstcPed.Cod_Unicidade Then
                'If lObjstcPedOrig.Cod_Ver_Unici <> lObjstcPed.Cod_Ver_Unici Then
                'If lObjstcPedOrig.Val_PriParcela <> lObjstcPed.Val_PriParcela Then
                'If lObjstcPedOrig.val_DemParcelas <> lObjstcPed.val_DemParcelas Then
                'If lObjstcPedOrig.Desc_Plano <> lObjstcPed.Desc_Plano Then
                'If lObjstcPedOrig.Val_RendaFatu <> lObjstcPed.Val_RendaFatu Then
                'If lObjstcPedOrig.LMI_Unico <> lObjstcPed.LMI_Unico Then
                'If lObjstcPedOrig.Per_Desc_LMI <> lObjstcPed.Per_Desc_LMI Then
                'If lObjstcPedOrig.Num_Cotac_LMI <> lObjstcPed.Num_Cotac_LMI Then
                'If lObjstcPedOrig.Tip_Transmissao <> lObjstcPed.Tip_Transmissao Then
                'If lObjstcPedOrig.Ind_CondGerais <> lObjstcPed.Ind_CondGerais Then
                'If lObjstcPedOrig.Ind_Reentrada <> lObjstcPed.Ind_Reentrada Then
                'If lObjstcPedOrig.Dig_PI_Emissao <> lObjstcPed.Dig_PI_Emissao Then
                'If lObjstcPedOrig.Nosso_Numero_Original <> lObjstcPed.Nosso_Numero_Original Then'
VoltarItem:
                For Each lobjstcPedLocOrig In lObjstcPedOrig.ColecaoPedLoc
                    For Each lobjstcPedLoc In lObjstcPed.ColecaoPedLoc
                        If lobjstcPedLoc.Num_Item = lobjstcPedLocOrig.Num_Item Then
                            With lobjstcPedLocOrig
                                If .Cod_Tip_Logr_Risco <> Trim(lobjstcPedLoc.Cod_Tip_Logr_Risco) Then
                                    lMsg = lMsg & "Tipo de logradouro : " & .Cod_Tip_Logr_Risco & "/" & lobjstcPedLoc.Cod_Tip_Logr_Risco
                                End If
                                If .Nom_Logr_Risco <> Trim(lobjstcPedLoc.Nom_Logr_Risco) Then
                                    lMsg = lMsg & "Nome logradouro : " & .Nom_Logr_Risco & "/" & lobjstcPedLoc.Nom_Logr_Risco
                                End If
                                If .Dsc_Compl_Risco <> Trim(lobjstcPedLoc.Dsc_Compl_Risco) Then
                                    lMsg = lMsg & "Complemento logradouro : " & .Dsc_Compl_Risco & "/" & lobjstcPedLoc.Dsc_Compl_Risco
                                End If
                                If Left(.Nom_Bairro_Risco, 25) <> Trim(lobjstcPedLoc.Nom_Bairro_Risco) Then
                                    lMsg = lMsg & "Bairro : " & .Nom_Bairro_Risco & "/" & lobjstcPedLoc.Nom_Bairro_Risco
                                End If
                                If .Nom_Cid_Risco <> Trim(lobjstcPedLoc.Nom_Cid_Risco) Then
                                    lMsg = lMsg & "Cidade : " & .Nom_Cid_Risco & "/" & lobjstcPedLoc.Nom_Cid_Risco
                                End If
                                If .Sig_UF_Risco <> lobjstcPedLoc.Sig_UF_Risco Then
                                    lMsg = lMsg & "UF : " & .Sig_UF_Risco & "/" & lobjstcPedLoc.Sig_UF_Risco
                                End If
                                If .Num_Cep_Risco <> lobjstcPedLoc.Num_Cep_Risco Then
                                    lMsg = lMsg & "CEP : " & .Num_Cep_Risco & "/" & lobjstcPedLoc.Num_Cep_Risco
                                End If
                                '.Per_Desc_Apar = Per_Desc_Apar
                                '.Per_Desc_Progres = Per_Desc_Progres
                                '.Per_IS_Vr = Per_IS_Vr
                                '.Val_Coef_IS_Vr = Val_Coef_IS_Vr
                                If val(.Val_Vr) <> val(lobjstcPedLoc.Val_Vr) Then
                                    lMsg = lMsg & "Valor do risco: " & .Val_Vr & "/" & lobjstcPedLoc.Val_Vr
                                End If
                                '.Tab_Clasf = Tab_Clasf
                                If .Cod_Clasf <> lobjstcPedLoc.Cod_Clasf Then
                                    lMsg = lMsg & "Cod Clasf: " & .Cod_Clasf & "/" & lobjstcPedLoc.Cod_Clasf
                                End If
                                If .Cod_Compl_Clasf <> lobjstcPedLoc.Cod_Compl_Clasf Then
                                    lMsg = lMsg & "Compl Clasf: " & .Cod_Compl_Clasf & "/" & lobjstcPedLoc.Cod_Compl_Clasf
                                End If
                                lAuxiliar = lobjstcPedLoc.Dsc_Clasf
                                lAuxiliar = Replace(lAuxiliar, "  ", " ")
                                lAuxiliar = Replace(lAuxiliar, "  ", " ")
                                lAuxiliar = Replace(lAuxiliar, "  ", " ")
                                lAuxiliar = Replace(lAuxiliar, "  ", " ")
                                If Left("* " & .Dsc_Clasf, 39) <> Left(Replace(Replace(Replace(Trim(lAuxiliar), "    ", " "), "   ", " "), "  ", " "), 39) And .Dsc_Clasf <> "" Then
                                    If Left(.Dsc_Clasf, 39) <> Left(Replace(Replace(Replace(Trim(lAuxiliar), "    ", " "), "   ", " "), "  ", " "), 39) And .Dsc_Clasf <> "" Then
                                        lMsg = lMsg & "Desc Clasf: " & .Dsc_Clasf & "/" & lAuxiliar
                                    End If
                                End If
                                '.Cod_Classe_Ocup = Cod_Classe_Ocup
                                '.Cod_Constr = Cod_Constr
                                '.Cod_User_Digit = Cod_User_Digit
                                '.DAT_DIGIT = DAT_DIGIT
                                '.HOR_DIGIT = HOR_DIGIT
                                '.Cod_User_Liber = Cod_User_Liber
                                '.Dat_Liber = Dat_Liber
                                '.Hor_Liber = Hor_Liber
                                '.Cod_Sit_item = Cod_Sit_item
                                '.Per_Bonus = Per_Bonus
                                '.Qtd_Empreg = Qtd_Empreg
                                '.Num_Insp = Num_Insp
                                '.Cod_Iden_Segur_01_V = Cod_Iden_Segur_01_V
                                '.Cod_Iden_Segur_02_A = Cod_Iden_Segur_02_A
                                '.Cod_Iden_Segur_03_P = Cod_Iden_Segur_03_P
                                '.Cod_Iden_Segur_04_I = Cod_Iden_Segur_04_I
                                '.Cod_Iden_Segur_05_G = Cod_Iden_Segur_05_G
                                '.Cod_Iden_Segur_06 = Cod_Iden_Segur_06
                                '.Cod_Iden_Segur_07 = Cod_Iden_Segur_07
                                '.Cod_Iden_Segur_08 = Cod_Iden_Segur_08
                                '.Cod_Iden_Segur_09 = Cod_Iden_Segur_09
                                '.Cod_Iden_Segur_10 = Cod_Iden_Segur_10
                                '.Per_Tot_Segur = Per_Tot_Segur
                                '.Cod_Iden_Protec_01_E = Cod_Iden_Protec_01_E
                                '.Cod_Iden_Protec_02_H = Cod_Iden_Protec_02_H
                                '.Cod_Iden_Protec_03_S = Cod_Iden_Protec_03_S
                                '.Cod_Iden_Protec_04_D = Cod_Iden_Protec_04_D
                                '.Cod_Iden_Protec_05 = Cod_Iden_Protec_05
                                '.Cod_Iden_Protec_06 = Cod_Iden_Protec_06
                                '.Cod_Iden_Protec_07 = Cod_Iden_Protec_07
                                '.Cod_Iden_Protec_08 = Cod_Iden_Protec_08
                                '.Cod_Iden_Protec_09 = Cod_Iden_Protec_09
                                '.Cod_Iden_Protec_10 = Cod_Iden_Protec_10
                                '.Per_Tot_Protec = Per_Tot_Protec
                                '.Cod_Produto = Cod_Produto
                                '.Cod_Regiao = Cod_Regiao
                                '.Cod_Iden_Tarif = Cod_Iden_Tarif
                                '.Num_PI = Num_PI
                                '.Num_Item_Ant = Num_Item_Ant
                                '.Tip_Emissao = Tip_Emissao
                                '.Num_Apol_Anterior = Num_Apol_Anterior
                                '.Num_Item_Anterior = Num_Item_Anterior
                                '.Ind_Sinistro = Ind_Sinistro

                                '1000347 - Inclusão de informações da congênere em renovações
                                '.Ren_Con_Seguradora = Ren_Con_Seguradora
                                '.Ren_Con_Apolice = Ren_Con_Apolice
                                '.Ren_Con_Item = Ren_Con_Item

                                '1200134 - Inspeção obrigatório por causa de sinistro
                                '.Ind_Sin_Inspec = Ind_Sin_Inspec

VoltarClausula:
                                For Each lObjstcPedClauOrig In .ColecaoPedClau
                                    For Each lobjstcPedClau In lobjstcPedLoc.ColecaoPedClau

                                        If lobjstcPedClau.Num_Item = lObjstcPedClauOrig.Num_Item And _
                                           lobjstcPedClau.Cod_Clau = lObjstcPedClauOrig.Cod_Clau Then

                                            lobjstcPedLoc.ColecaoPedClau.Remove lobjstcPedClau.Class_ID
                                            .ColecaoPedClau.Remove lObjstcPedClauOrig.Class_ID
                                            GoTo VoltarClausula
                                        End If
                                    Next lobjstcPedClau
                                    If lObjstcPedClauOrig.Cod_Clau = 2 Or lObjstcPedClauOrig.Cod_Clau = 1 Or lObjstcPedClauOrig.Cod_Clau = 403 Then
                                        .ColecaoPedClau.Remove lObjstcPedClauOrig.Class_ID
                                    End If
                                Next lObjstcPedClauOrig


                                If .ColecaoPedClau.Count <> lobjstcPedLoc.ColecaoPedClau.Count Then
                                    lMsg = lMsg & "Quantidade de clausulas "
                                End If

                                If lNumApol = 11779107 Then
                                    MsgBox ""
                                End If
VoltarCobertura:
                                For Each lobjstcCoberturaOrig In .ColecaoPedCobert
                                    For Each lobjstcCobertura In lobjstcPedLoc.ColecaoPedCobert
                                        If lobjstcCobertura.Num_Item = lobjstcCoberturaOrig.Num_Item And _
                                           lobjstcCobertura.Cod_Cobert = lobjstcCoberturaOrig.Cod_Cobert Then
                                            If val(lobjstcCobertura.Val_IS) <> val(lobjstcCoberturaOrig.Val_IS) Then
                                                lMsg = lMsg & "Cobertura " & lobjstcCobertura.Cod_Cobert & " Val IS: " & lobjstcCobertura.Val_IS & "/" & lobjstcCoberturaOrig.Val_IS
                                            End If
                                            '.Per_Taxa = Per_Taxa
                                            '.Per_Desc_Coml = Per_Desc_Coml
                                            '.Per_Desc_Exp = Per_Desc_Exp
                                            If lobjstcCobertura.VAL_PR <> lobjstcCoberturaOrig.VAL_PR Then
                                                'lMsg = lMsg & "Cobertura " & lobjstcCobertura.Cod_Cobert & " Val PR: " & lobjstcCobertura.VAL_PR & "/" & lobjstcCoberturaOrig.VAL_PR
                                            End If

                                            '.Val_Pr_Anu = Val_Pr_Anu
                                            '.Val_Pr_Tarif = Val_Pr_Tarif
                                            '.Val_Pr_Tarif_Anu = Val_Pr_Tarif_Anu
                                            '.Val_Pr_Resg = Val_Pr_Resg
                                            '.Val_Pr_Resg_Anu = Val_Pr_Resg_Anu
                                            If lobjstcCobertura.Per_Franq <> lobjstcCoberturaOrig.Per_Franq And lNumEndo = 0 Then
                                                lMsg = lMsg & "Cobertura " & lobjstcCobertura.Cod_Cobert & " Per_Franq : " & lobjstcCobertura.Per_Franq & "/" & lobjstcCoberturaOrig.Per_Franq
                                            End If
                                            If lobjstcCobertura.Val_Franq_Min <> lobjstcCoberturaOrig.Val_Franq_Min And lNumEndo = 0 And lobjstcCobertura.Num_Meses_PI = 0 Then
                                                lMsg = lMsg & "Cobertura " & lobjstcCobertura.Cod_Cobert & " Val_Franq_Min : " & lobjstcCobertura.Val_Franq_Min & "/" & lobjstcCoberturaOrig.Val_Franq_Min
                                            End If
                                            '.Val_Franq_Max = Val_Franq_Max
                                            '.Num_PI = Num_PI
                                            If lobjstcCobertura.Num_Meses_PI <> lobjstcCoberturaOrig.Num_Meses_PI And lNumEndo = 0 Then
                                                lMsg = lMsg & "Cobertura " & lobjstcCobertura.Cod_Cobert & " Num_Meses_PI: " & lobjstcCobertura.Num_Meses_PI & "/" & lobjstcCoberturaOrig.Num_Meses_PI
                                            End If
                                            '.Cod_Classe = Cod_Classe
                                            '.Formula = Formula
                                            '.Val_Fator_Franq = Val_Fator_Franq
                                            '.Alterar_Cobert = Alterar_Cobert
                                            '.Val_Risco_Cob = Val_Risco_Cob
                                            '.Nom_Cobertura = Nom_Cobertura
                                            '.Cod_Grupo_Cobertura = Cod_Grupo_Cobertura
                                            .ColecaoPedCobert.Remove lobjstcCoberturaOrig.Class_ID
                                            lobjstcPedLoc.ColecaoPedCobert.Remove lobjstcCobertura.Class_ID
                                            GoTo VoltarCobertura:
                                        End If
                                    Next lobjstcCobertura
                                Next lobjstcCoberturaOrig


                                If .ColecaoPedCobert.Count <> lobjstcPedLoc.ColecaoPedCobert.Count Then
                                    lMsg = lMsg & "Quantidade de coberturas "
                                End If

VoltarFator:
                                For Each lobjstcFatorOrig In .ColecaoPedFator
                                    For Each lobjstcFator In lobjstcPedLoc.ColecaoPedFator
                                        If lobjstcFatorOrig.Num_Item = lobjstcFator.Num_Item And _
                                           lobjstcFatorOrig.Desc_Fator = lobjstcFator.Desc_Fator Then
                                            .ColecaoPedFator.Remove lobjstcFatorOrig.Class_ID
                                            lobjstcPedLoc.ColecaoPedFator.Remove lobjstcFator.Class_ID
                                            GoTo VoltarFator
                                        End If
                                    Next lobjstcFator
                                Next lobjstcFatorOrig

                                If .ColecaoPedFator.Count <> lobjstcPedLoc.ColecaoPedFator.Count Then
                                    lMsg = lMsg & "Quantidade de fator de proteção"
                                End If


                                'Set .ColecaoPedLocBem = New Collection
                                'For Each lobjstcBem In Me.ColecaoPedLocBem
                                '    Set lobjstcBem = lobjstcBem.gCopiaBem(pNum_Item)
                                '    Set lobjstcBem.ObjPai = Me
                                '    .ColecaoPedLocBem.Add lobjstcBem
                                'Next lobjstcBem

                                'Set .ColecaoPedSeguro = New Collection
                                'For Each lobjstcSeguro In Me.ColecaoPedSeguro
                                '    Set lobjstcSeguro = lobjstcSeguro.gCopiaSeguro(pNum_Item)
                                '    Set lobjstcSeguro.ObjPai = Me
                                '    .ColecaoPedSeguro.Add lobjstcSeguro, lobjstcSeguro.Class_ID
                                'Next lobjstcSeguro
                                '
                                'Set .objstcPedObs = objstcPedObs.gCopiaObs(pNum_Item)
                                '
                                ' Set .objstcPedInsp = objstcPedInsp.gCopiaInsp(pNum_Item)

                                'Set .ColecaoLMI = New Collection
                                'For Each lobjstcLMI In Me.ColecaoLMI
                                '    Set lobjstcLMI = lobjstcLMI.gCopiaLMI(pNum_Item)
                                '    .ColecaoLMI.Add lobjstcLMI, lobjstcLMI.Class_ID
                                'Next lobjstcLMI

                                '.Tip_Cobert = Tip_Cobert
                            End With
                        End If
                    Next lobjstcPedLoc
                Next lobjstcPedLocOrig
            End If

            lAuxiliar = lNumApol & "-" & lNumEndo & " - "
            If gM46V999.gfPreenchido(lMsg) Then
                lLista = lLista & vbCrLf & lAuxiliar & lMsg
            Else
                lLista = lLista & vbCrLf & lAuxiliar & "OK"
            End If
        End If
Proximo:
    Loop
    Close #lNumArq
    Call EnviarEmail("simony@yasuda.com.br", "simony@yasuda.com.br;luana.delazari@cybersolutions.com.br;frank.souza@cybersolutions.com.br;aline.iozzi@cybersolutions.com.br;", , "Conferência do book_ " & lNomeArq & "_" & lData, Replace(lLista, vbCrLf, "<br>"))
    'Call EnviarEmail("tjgomes@yasuda.com.br", "tjgomes@yasuda.com.br;", , "Conferência do book_ " & lNomeArq, Replace(lLista, vbCrLf, "<br>"))
    Exit Sub
ControleErro:
    Call EnviarEmail("simony@yasuda.com.br", "simony@yasuda.com.br", , "Conferência do book_ " & lNomeArq, Err.Description & "/" & lNumApol & "/" & lNumEndo & "/" & lFormulario)
End Sub
Public Function EnviarEmail(ByVal pFrom As String, ByVal pTo As String, _
                            Optional ByVal pCc As String, _
                            Optional ByVal pSubject As String, _
                            Optional ByVal pBody As String, _
                            Optional ByVal pAttach As String, _
                            Optional ByVal pOpenMessage As Boolean, _
                            Optional ByVal pBcc As String) As Boolean

    On Error GoTo TratarErro

    EnviarEmail = False

    Dim sMessage As String
    Dim oMessage As Object
    Dim lMensagem As String

    Set oMessage = CreateObject("CDO.Message")
    If Not pOpenMessage Then
        With oMessage.Configuration.Fields
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "spx20858psluexc"
            .Update
        End With
        oMessage.From = pFrom
        If pTo <> vbNullString Then oMessage.To = pTo
        If pCc <> vbNullString Then oMessage.CC = pCc
        If pBcc <> vbNullString Then oMessage.BCC = pBcc
        oMessage.Subject = pSubject
        oMessage.HTMLBody = pBody
        oMessage.MimeFormatted = True
        oMessage.Send
    Else
        sMessage = pTo
        If pCc <> vbNullString Then sMessage = sMessage & "&cc=" & pCc
        If pSubject <> vbNullString Then sMessage = sMessage & "&subject=" & pSubject
        If pBody <> vbNullString Then sMessage = sMessage & "&body=" & Replace(pBody, vbCrLf, "%0D%0A")
        If pAttach <> vbNullString Then sMessage = sMessage & "&attachment=" & pAttach

        ShellExecute 0, "open", "mailto:" & sMessage, "", "", vbNormalFocus
    End If
    EnviarEmail = True
    Exit Function
TratarErro:
    lMensagem = Err.Description
    Call gM46V999.gpGraLog(1, lMensagem)
End Function

Private Sub mpVerRecotacao()
    Dim lNovoNossoNumero As String
    Dim lNovReferencia As String
    Dim lMensagem As String
    Dim lobjCotacao As clsA28V721A
    Dim lobjstcCotacao As stcA28V721B


    gM46V111.gComNosNumero = ""
    Call gM46V111.gpObterRegComunicacao(False, "", "")
    If gM46V999.gfPreenchido(gM46V111.gComNosNumero) Then
        Set lobjstcCotacao = New stcA28V721B
        lobjstcCotacao.NOSSO_NUMERO = gM46V111.gComNosNumero
        lobjstcCotacao.STATUS = -1
        Set lobjCotacao = New clsA28V721A
        Set lobjCotacao.M46V999 = gM46V999

        If Not lobjCotacao.gfCarregarColecao(lobjstcCotacao, "") Then
            Call gM46V999.gpGraLog(1, lobjCotacao.UltimoErro)
            Exit Sub
        End If
        If lobjCotacao.Colecao.Count <> 0 Then

            Set lobjstcCotacao = lobjCotacao.Colecao.Item(1)

            If gM46V999.gfCalNumPedido(lNovReferencia, lMensagem) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            '4. Calcula nosso número e número do pedido.
            If gM46V999.gfCalNosNumero(lNovoNossoNumero, lMensagem) = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            gM46V999.gOrigemChamada = e_Origem_Kit
            If Not gM46V999.gfCopProposta(gM46V111.gComNosNumero, lNovoNossoNumero, lNovReferencia, lMensagem) Then
                Call gM46V999.gpGraLog(1, lMensagem)
                Exit Sub
            End If
            If gM46V111.gfGraRegComunicacao(lNovoNossoNumero, lMensagem) = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
                Exit Sub
            End If
            lobjstcCotacao.STATUS = -1
            If Not lobjCotacao.gfExcluir(lobjstcCotacao, "") Then
                Call gM46V999.gpGraLog(1, lobjCotacao.UltimoErro)
                Exit Sub
            End If
            lobjstcCotacao.NOSSO_NUMERO = lNovoNossoNumero
            lobjstcCotacao.STATUS = -1
            Set lobjCotacao = New clsA28V721A
            Set lobjCotacao.M46V999 = gM46V999

            If Not lobjCotacao.gfIncluir(lobjstcCotacao, "") Then
                Call gM46V999.gpGraLog(1, lobjCotacao.UltimoErro)
                Exit Sub
            End If
            Call gM46V999.gpGraLog(1, "Para a recotação ser realizada, a cotação anterior foi copiada para a referência " & lNovReferencia & ".(Nosso número : " & lNovoNossoNumero & ")")

            If App.EXEName = "P46V101" Then
                Unload Me
            Else
                If gfShell("P46V100", False) = True Then
                    Unload Me
                End If
            End If
        End If
    End If
End Sub

Sub mfVerificaVersao()
    Dim lbdTAB_Produto As ADODB.Connection    'Base de dados .
    Dim lbdTAB_Produto2 As ADODB.Connection    'Base de dados .
    Dim lMensagem As String    'Mensagem.
    Dim lrsTAB_Produto As clsYasRecordSet    'Registro .
    Dim lrsTAB_Produto2 As clsYasRecordSet    'Registro .
    Dim NomePasta As String
    Dim lSelect As String

    On Error Resume Next

    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Or UCase(gM46V999.gAppPath) = "S:\SYASPROD\DIVERSOS\" Then
        GoSub ExcluirBaseTemp
        Exit Sub
    End If

    Dim fso As New FileSystemObject
'    Dim objShell As Object
'    Set objShell = CreateObject("WScript.Shell")
'    NomePasta = objShell.ExpandEnvironmentStrings("%APPDATA%")
'    NomePasta = NomePasta & "\SYAS\"
'    fso.CreateFolder (NomePasta)
'

'/********  Comentar caso use a pasta APPDATA do usuario ****************/
    NomePasta = Empty
    If gfIsMetaframe() Then
        NomePasta = "L:\User\Syas\"
        MkDir NomePasta
        If Err = 0 Then
            Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & NomePasta)
        Else
            NomePasta = Empty
        End If
    End If

    If NomePasta = Empty Then
        GoSub ExcluirBaseTemp
        NomePasta = "C:\Temp\"
        If Not fso.FolderExists(NomePasta) Then
            fso.CreateFolder (NomePasta)
        End If
        
        NomePasta = NomePasta & "Syas\"
        If Not fso.FolderExists(NomePasta) Then
            fso.CreateFolder (NomePasta)
        End If
    End If
'/********  Comentar caso use a pasta APPDATA do usuario ****************/

    If Not fso.FileExists(NomePasta & "P00MULTI.MDB") Then
        GoTo Copia
    End If
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Produto, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, NomePasta & "P00MULTI.MDB", lbdTAB_Produto2, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If


    lSelect = "SELECT * FROM Tab_Produto "
    lSelect = lSelect & " order by Dat_Alt desc "

    If gM46V999.gfObtRegistro(lbdTAB_Produto, lSelect, lrsTAB_Produto, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_Produto)
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If gM46V999.gfObtRegistro(lbdTAB_Produto2, lSelect, lrsTAB_Produto2, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_Produto)
        Call gM46V999.gpFecha2(lbdTAB_Produto2)
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If Format(lrsTAB_Produto!Dat_Alt, "00000000000000") > Format(lrsTAB_Produto2!Dat_Alt, "00000000000000") Then
        GoTo Copia
    End If

    gM46V999.gEndP00MULTI = NomePasta & "P00MULTI.MDB"
    gM46V999.gTipoHomologacao = 1
    Exit Sub

Copia:
    panAviso.Caption = "Atualizando Banco Dados Local. Tempo de Espera:      5 minutos. Por favor aguarde."
    panAviso.Visible = True
    panAviso.Refresh

    Call gM46V999.gpFecha2(lbdTAB_Produto)
    Call gM46V999.gpFecha2(lbdTAB_Produto2)
    Err = 0
    FileCopy gM46V999.gEndP00MULTI, NomePasta & "P00MULTI.MDB"
    panAviso.Visible = False
    If Err <> 0 Then
        MsgBox Error
    End If
    If Dir(NomePasta & "P00MULTI.MDB") = "" Then
        Call gM46V999.gpFecha2(lbdTAB_Produto)
        Call gM46V999.gpFecha2(lbdTAB_Produto2)
        Call gM46V999.gpGraLog(2, "Não foi possível copia a base para a base CITRIX! Sistema Encerrado.")
        End
    End If
    Call gM46V999.gpGraLog(0, "P00MULTI copiado com sucesso de: " & gM46V999.gEndP00MULTI & " para: " & NomePasta & "P00MULTI.MDB")
    gM46V999.gEndP00MULTI = NomePasta & "P00MULTI.MDB"
    gM46V999.gTipoHomologacao = 1
    Exit Sub

ExcluirBaseTemp:

    gM46V999.gTipoHomologacao = 0

    If Dir("L:\") = "" Then
        Return
    End If
    If Dir("L:\USER\SYAS\P00MULTI.MDB") <> "" Then
        On Error Resume Next
        Err = 0
        Kill "L:\USER\SYAS\P00MULTI.MDB"
        If Err <> 0 Then
            Call gM46V999.gpGraLog(2, "Não foi possível remover a base CITRIX! Sistema Encerrado.")
            End
        End If
    End If
    If Dir("C:\TEMP\SYAS\P00MULTI.MDB") <> "" Then
        On Error Resume Next
        Err = 0
        Kill "C:\TEMP\SYAS\P00MULTI.MDB"
        If Err <> 0 Then
            Call gM46V999.gpGraLog(2, "Não foi possível remover a base CITRIX! Sistema Encerrado.")
            End
        End If
    End If
    Return

End Sub



Private Sub imgLogo_Click()
End Sub

Private Sub lblEndosso_Click()
    Call mnuArqEndosso_Click
End Sub

Private Sub lblImpressao_Click()
    Call mnuFerImpressao_Click
End Sub

Private Sub mnuAgendarInspecao_Click()
    frmT46V174A.Show vbModal
End Sub

Private Sub mnuFerCadCCorretagem_Click()
    frmT46V170A.Show vbModal
End Sub

Private Sub mnuFerImpressao_Click()
    frmT46V163A.Show vbModal
End Sub

Private Sub mnuArqEndosso_Click()
'    Call gM46V999.gpGraLog(0, "AGE0050 - Opção selecionada: arquivo - endosso.")
    Call gM46V999.gpGraLog(0, "AGE0004 - Opção selecionada: arquivo - endosso.")

    gM46V111.gNum_Pi = 0
    On Error GoTo ControleErro

    With frmT46V178A
        '.Caption = "Endosso"
        .Show 1
        If Forms.Count <> 1 Then
            Unload frmT46V178A
            If App.EXEName = "P46V101" Then
                Call gpP46v100Main
            Else
                If gfShell("P46V100") = True Then
                    Unload Me
                End If
            End If
        End If
    End With
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mnuArqEndosso ERRO - " & Err & " " & Err.Description)

End Sub

''2016-07-11: Nova funcionalidade de impressao com dados do FTP
Private Sub mnuFerImpressaoTransmissao_Click()
    Call gM46V999.gpGraLog(0, "AGE0004 - Opção selecionada: ferramentas - impressão de transmissão.")
    
    frmT46V198A.Show vbModal  ''Novo form de pesquisa na FTP
End Sub

Private Sub mnuFerRegPadrao_Click()
    Dim lMensagem As String    'Mensagem.

    Call gM46V999.gpGraLog(0, "AGE0004 - Opção selecionada: ferramentas - registro padrão.")

    '1. Grava registro de comunicação.
    If gM46V111.gfGraRegComunicacao("PADRAO", lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    '2. Passa o controle para o programa P46V100.
    If App.EXEName = "P46V101" Then
        Call gpP46v100Main
    Else
        If gfShell("P46V100") = True Then
            Unload Me
        End If
    End If

End Sub

Private Sub Form_Resize()
'Posiciona logo.
    On Error GoTo ControleErro



    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V001A.Resize ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If gM46V999.gfDisponivel("frmT46V100A") Then
        Unload frmT46V100A
    End If
    Call gM46V999.gpGraLog(0, "AGE0007 - Encerrado processamento do programa " & App.EXEName & ".")
End Sub

Private Sub lblAjuda_Click()
    Call mnuAjuSYAS_Click
End Sub
Private Sub lblManCorretor_Click()
    Call mnuArqManCorretor_Click
End Sub
Private Sub lblRenovacao_Click()
    Call mnuArqRenovacao_Click
End Sub
Private Sub lblSegNovo_Click()
    Call mnuArqSegNovo_Click
End Sub
Private Sub lblPesquisa_Click()
    Call mnuArqPesquisa_Click
End Sub
Private Sub lblSair_Click()
    Call mnuSair_Click
End Sub
Private Sub lblTransmitir_Click()
    Call mnuArqTransmissao_Click
End Sub
Private Sub mnuAjuSobre_Click()
    Call gM46V999.gpGraLog(0, "AGE0004 - Opção selecionada: ajuda - sobre.")

    frmT46V010A.Show vbModal
End Sub
Private Sub mnuAjuSYAS_Click()
    Call gM46V999.gpGraLog(0, "AGE0004 - Opção selecionada: ajuda - SYAS.")

    With cdgHelp
        .CancelError = True
        On Error Resume Next
        .HelpCommand = cdlHelpForceFile
        .HelpFile = App.HelpFile
        .ShowHelp
        On Error GoTo 0
    End With
End Sub
Private Sub mnuArqManCorretor_Click()
    Dim lMensagem As String    'Mensagem.

    Call gM46V999.gpGraLog(0, "AGE0004 - Opção selecionada: arquivo - manual do corretor.")

    'Call gM46V999.gpGraLog(1, "Manual do corretor em fase de ajuste.")
    'Call Form1.Show
    lMensagem = "AGE0008 - Este microcomputador possui o Acrobat Reader 5.0 instalado?"
    Select Case MsgBox(lMensagem, vbYesNo + vbQuestion, "Manual do corretor.")
    Case vbYes
        Call gM46V999.gpGraLog(0, lMensagem & " SIM.")
        'WinExec "Explorer.exe " & gM46V999.gAppPath & "Tela Diversos.PDF", 10
        ShellExecute 0, "open", gM46V999.gAppPath & "Tela Diversos.PDF", "", "", 1
    Case vbNo
        Call gM46V999.gpGraLog(0, lMensagem & " NÃO.")
        lMensagem = "AGE0039 - Deseja instalar o Acrobat Reader 5.0?"
        If MsgBox(lMensagem, vbYesNo + vbQuestion, "Manual do corretor.") = vbYes Then
            Call gM46V999.gpGraLog(0, lMensagem & " SIM.")
            'WinExec gM46V999.gAppPath & "ar500ptb.exe", 9
            'WinExec "Explorer.exe " & gM46V999.gAppPath & "Tela Diversos.PDF", 10
            'ShellExecute 0, "open", "acrord32", gM46V999.gAppPath & "Tela Diversos.PDF", vbNull, 3
            ShellExecute 0, "open", gM46V999.gAppPath & "Tela Diversos.PDF", "", "", 1
        Else
            Call gM46V999.gpGraLog(0, lMensagem & " NÃO.")
        End If
    End Select
End Sub
Private Sub mnuArqRenovacao_Click()
    Call gM46V999.gpGraLog(0, "AGE0004 - Opção selecionada: arquivo - renovação.")
    gM46V111.gNum_Pi = 0
    On Error GoTo ControleErro

    With frmT46V002A
        .Caption = "Cotação - Renovação"
        .Show 1
        If Forms.Count <> 1 Then
            Unload frmT46V002A
            If App.EXEName = "P46V101" Then
                Call gpP46v100Main
            Else
                If gfShell("P46V100") = True Then
                    Unload Me
                End If
            End If
        End If
    End With
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mnuArqRenovacao ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub mnuArqPesquisa_Click()
    Call gM46V999.gpGraLog(0, "AGE0004 - Opção selecionada: arquivo - pesquisa cotações, propostas, apólices " & _
                              "e protocolo.")

    On Error GoTo ControleErro
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And _
       gM46V999.PerfilAcesso <> e_Perfil_CALC And gM46V999.PerfilAcesso <> e_Perfil_CALC_Gestor Then    'Yasuda
        frmT46V012A.Show 1
        If gM46V999.gfDisponivel("frmT46V012A") = True Then
            If frmT46V012A.cboOpcao.ListIndex = 4 Then Exit Sub
            Unload frmT46V012A
            If App.EXEName = "P46V101" Then
                Call gpP46v100Main
            Else
                If gfShell("P46V100") = True Then
                    While Forms.Count <> 0
                        Unload Forms(0)
                    Wend
                End If
            End If
        End If
    Else                        'Corretor ou Cotação
        frmT46V003A.Show 1
        If gM46V999.gfDisponivel("frmT46V003A") = True Then
            Unload frmT46V003A
            If App.EXEName = "P46V101" Then
                Call gpP46v100Main
            Else
                If gfShell("P46V100") = True Then
                    While Forms.Count <> 0
                        Unload Forms(0)
                    Wend
                End If
            End If
        End If
    End If
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mnuArqPesquisa ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub mnuArqSegNovo_Click()
    Call gM46V999.gpGraLog(0, "AGE0004 - Opção selecionada: arquivo - seguro novo.")
    gM46V111.gNum_Pi = 0
    If gM46V001.gfSegNovo("", "", 0, 0, 0, "") = True Then
        'Passa o controle para o programa P46V100.
        If gM46V999.gfDisponivel("frmT46V024A") = True Then
            Unload frmT46V024A
        End If
        If gM46V999.gfDisponivel("frmT46V012A") = True Then
            Unload frmT46V012A
        End If
        If App.EXEName = "P46V101" Then
            Call gpP46v100Main
        Else
            If gfShell("P46V100") = True Then
                Unload Me
            End If
        End If
    End If
End Sub
Private Sub mnuArqTransmissao_Click()
    Call gM46V999.gpGraLog(0, "AGE0004 - Opção selecionada: arquivo - transmitir proposta.")

    On Error GoTo ControleErro

    '1. Verifica se corretor pode transmitir.
    '**************    TODOS OS CORREOTRES PODEM TRANSMITIR
    '    If gfVerTransmissao = False Then
    '        Exit Sub
    '    End If

    '2. Abre janela para processo de transmissão.
    frmT46V005A.Show 1
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mnuArqTransmissao ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub mnuFerLocDados_Click()
    Call gM46V999.gpGraLog(0, "AGE0004 - Opção selecionada: ferramentas - localização dos dados.")

    On Error GoTo ControleErro
    If UCase(gM46V999.gEndServidor) <> "S:\SYASPROD\DADOS" Then
        frmT46V007A.Show vbModal
    End If
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mnuFerLocDados ERRO - " & Err & " " & Err.Description)
End Sub

Private Sub mnuManutencao_Click()
    frmT46V201A.Show vbModal
End Sub

Private Sub mnuSair_Click()
    Unload Me

    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then      'Corretor.
        Call gfShell("P28V000", True)
    End If
    End
End Sub
Private Sub panBasDados_DblClick()
    MsgBox "Endereço das bases de dados: " & panBasDados.Caption
End Sub
Private Sub panCorretor_DblClick()
    MsgBox "Código do corretor: " & Mid$(panCorretor.Caption, 10)
End Sub
Private Sub panData_DblClick()
    MsgBox "Data de hoje registrado neste microcomputador: " & panData.Caption
End Sub
Private Sub panForm_DblClick()
    MsgBox "Programa ativo: " & panForm.Caption
End Sub
Private Sub panUsuario_DblClick()
    MsgBox "Código do usuário: " & Mid(panUsuario.Caption, 9)
End Sub
Private Sub panVersao_DblClick()
    MsgBox "Data da versão do programa: " & Mid(panVersao.Caption, 8)
End Sub

Private Sub mpVerificaUltCn(ByVal lDatArqConsiderado As Double, ByVal lUltimaCN As Integer)
'Procedure utilizada para verificar a data da ultima conexão

'Parâmetros de entrada...lDatArqConsiderado...Data do arquivo considerado na verificação.
'                        lUltimaCN............Quantidade de dias da última conexão.

    Dim lDatP2842400EXE As Double   'P2842400.EXE: data do arquivo.
    Dim lEndP2842400EXE As String   'P2842400.EXE: endereço do arquivo.
    Dim lMensagem As String    'Mensagem.
    Dim lNumP2842400EXE As Integer  'P2842400.EXE: utilizado com FreeFile.

    On Error GoTo ControleErro

    Select Case lDatArqConsiderado
    Case 0      'Arquivo de conexão não existe.
        'Obs.: Se data do arquivo P2842400.EXE não é a data do dia, não precisa dar mensagem.
        lEndP2842400EXE = gM46V999.gAppPath & "P2842400.EXE"
        If gM46V999.gfPreenchido(Dir(lEndP2842400EXE)) Then
            lDatP2842400EXE = CDbl(Format$(FileDateTime(lEndP2842400EXE), "yyyymmdd"))
        Else
            lDatP2842400EXE = 0
        End If
        If lDatP2842400EXE <> gM46V999.gHoje Then
            Call gM46V999.gpGraLog(1, "AGE0010 - Registro de conexão com a Yasuda não localizado." & vbLf & _
                                      vbLf & "Providenciar conexão com a Yasuda para garantir a versão " & _
                                      "atualizada do sistema e não ocasionar rejeição de proposta.")
        End If
        lNumP2842400EXE = FreeFile
        Open lEndP2842400EXE For Output As #lNumP2842400EXE
        Print #lNumP2842400EXE, lMensagem
        Close #lNumP2842400EXE
    Case Else
        Select Case lUltimaCN
        Case 1
            Call gM46V999.gpGraLog(1, "AGE0011 - Última conexão com a Yasuda: " & _
                                      gM46V999.gfForData(2, lDatArqConsiderado) & " (a " & lUltimaCN & " dias)." & _
                                      vbLf & vbLf & "Providenciar conexão com a Yasuda para " & _
                                      "garantir a versão atualizada do sistema e não ocasionar " & _
                                      "rejeição de propostas.")
        Case Is >= 2
            Call gM46V999.gpGraLog(1, "AGE0012 - Última conexão com a Yasuda: " & _
                                      gM46V999.gfForData(2, lDatArqConsiderado) & " (a " & lUltimaCN & _
                                    " dias). Conectar com a Yasuda para garantir o processo de " & _
                                      "cálculo. Sistema indisponível até próxima conexão.")
            Call mpDesabilita
        End Select
    End Select
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mpVerificaUltCn ERRO - " & Err & " " & Err.Description)
End Sub
Private Sub mpEnvioSYASINFO()
'Procedure: envia arquivos LOG contendo informações sobre as versões dos programas utilizados pelo
'           usuário.

    Dim lDirArqLog As String    'Diretório do arquivo log.
    Dim lEndArqLog As String    'Endereço do arquivo log.
    Dim lEndGEISEDI2NET As String   'Endereço do arquivo GEISEDI2.net.
    Dim lEndSYASINFO001 As String   'Endereço da Copia do arquivo SYASINFO.LOG.
    Dim lEndSYASINFO006 As String   'Endereço do SYASINFO.LOG.
    Dim lFlag As Boolean    'Flag: True = continua e False = sai da procedure.
    Dim lLCoduser As String    'Long     'Código do usuário de transmissão em formato Long.
    Dim lMsg As String    'Mensagem.
    Dim lNumGEISEDI2net As String   'Arquivo GEISEDI2.NET.
    Dim lNumSYASINFO001 As String   'Arquivo SyasInfo.
    Dim lNumSYASINFO006 As String   'Arquivo SyasInfo.
    Dim lRegistro As String    'Registro do arquivo.
    Dim lSCoduser As String    'Código do usuário de transmissão em formato String.
    Dim lSpliRegistro() As String   'Registro (Split).

    On Error GoTo 0

    '1. Inicializa variável.
    lFlag = True

    '2. Inicia cópia dos logs para a pasta remessa - obtém caminho origem do arquivo a ser copiado.
    If gM46V999.gfPreenchido(gM46V999.gEndSYASUpDt) = True Then
        lDirArqLog = gM46V999.gEndSYASUpDt
    Else
        lDirArqLog = gM46V999.gEndServidor
    End If

    '3. Copia arquivos.
    'a) Obtém primeiro arquivo a ser copiado.
    lEndArqLog = lDirArqLog & Dir(lDirArqLog & "*YY.006")
    Do While lFlag = True
        If gM46V999.gfPreenchido(Dir(lDirArqLog & "*YY.006")) = True Then
            FileCopy lEndArqLog, gM46V999.gAppPath & "REMESSA\" & Mid(lEndArqLog, Len(lEndArqLog) - 11)
            'b) Exclui da origem o arquivo já copiado.
            Kill lEndArqLog
        Else
            lFlag = False
        End If
        'c) Obtém próximo arquivo a ser copiado.
        lEndArqLog = lDirArqLog & Dir(lDirArqLog & "*YY.006")
    Loop

    '4. Obtém código do usuário que está enviando o log.
    'a) Reincia variável - redundância evita problemas inesperados.
    lFlag = False
    If gfBuscaUsuarioTransmissao(lLCoduser, lMsg) = False Then
        'b) Sai do procedimento - não obteve código de usuário.
        Call gM46V999.gpGraLog(2, lMsg)
        Exit Sub
    End If
    lSCoduser = Format(lLCoduser, "000000")

    '5. Verifica se usuário que está iniciando a transmissão é o usuário de recebe freqüentemente as
    '   atualizações do SYAS, através do arquivo GEISEDI2.net.
    If gM46V999.gfPreenchido(gM46V999.gEndSYASUpDt) = True Then
        If gM46V999.gfPreenchido(Dir(gM46V999.gEndSYASUpDt & "GEISEDI2.NET")) = True Then
            'Cria réplica do arquivo do GEISEDI2.NET
            FileCopy gM46V999.gEndSYASUpDt & "GEISEDI2.NET", gM46V999.gAppPath & "REMESSA\GEISEDI2.NET"
        Else
            'Não existe informações de rede
            lMsg = gM46V999.gEndSYASUpDt & "GEISEDI2.NET" & vbLf & vbLf & "Arquivo não localizado."
            Call gM46V999.gpGraLog(0, lMsg)
            Exit Sub
        End If
    Else
        If gM46V999.gfPreenchido(Dir(gM46V999.gEndServidor & "GEISEDI2.NET")) = True Then
            'Cria réplica do arquivo do GEISEDI2.NET
            FileCopy gM46V999.gEndServidor & "GEISEDI2.NET", gM46V999.gAppPath & "REMESSA\GEISEDI2.NET"
        Else
            'Não existe informações de rede
            lMsg = gM46V999.gEndServidor & "GEISEDI2.NET" & vbLf & vbLf & "Arquivo não localizado."
            Call gM46V999.gpGraLog(0, lMsg)
            Exit Sub
        End If
    End If

    '6. Inicia leitura do arquivo GEISEDI2.NET.
    lEndGEISEDI2NET = gM46V999.gAppPath & "REMESSA\GEISEDI2.NET"
    lNumGEISEDI2net = FreeFile
    Open lEndGEISEDI2NET For Input Access Read As #lNumGEISEDI2net
    If Err <> 0 Then
        Call gM46V999.gpGraLog(1, "FTE0078 - Erro ao abrir arquivo " & lEndGEISEDI2NET & ": " & Err & " - " & _
                                  Error & vbLf & vbLf & "O sistema será encerrado.")
        End
    Else
        Do While Not EOF(lNumGEISEDI2net)
            Line Input #lNumGEISEDI2net, lRegistro
            lSpliRegistro = Split(lRegistro, "$")
            If UBound(lSpliRegistro()) >= 1 Then
                If Format(Trim(lSpliRegistro(1)), "000000") = Format(Trim(lSCoduser), "000000") Then
                    'Usuário é o mesmo da rede - altera variável flag para continuar o processamento.
                    lFlag = True
                    Exit Do
                Else
                    'Sai do procedimento - mantém a variável lFlag como False para não continuar
                    'procedimento.
                    Exit Do
                End If
            Else
                'Arquivo existende está danificado ou com informações incorretas - exclui para evitar
                'problemas.
                If gM46V999.gfPreenchido(gM46V999.gEndSYASUpDt) = True Then
                    On Error Resume Next
                    Kill gM46V999.gEndSYASUpDt & "GEISEDI2.NET"
                    On Error GoTo 0
                Else
                    On Error Resume Next
                    Kill gM46V999.gEndServidor & "GEISEDI2.NET"
                    On Error GoTo 0
                End If
                'Sai do procedimento - mantém variável lFlag como False para não continuar procedimento.
                Exit Do
            End If
        Loop
        Close #lNumGEISEDI2net
    End If
    Kill lEndGEISEDI2NET

    '7. Somente se o usuário da rede for identificado.
    If lFlag = True Then
        'a) Reescreve o arquivo SYASINFO.LOG - informando o usuário REDE.
        '   Atribui o caminho dos arquivos na variável e réplica.
        If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & "REMESSA\" & Trim(lSCoduser) & "YY.006")) = True Then
            'Atribui o caminho dos arquivos SYASINFO.LOG na variável.
            lEndSYASINFO006 = gM46V999.gAppPath & "REMESSA\" & Trim(lSCoduser) & "YY.006"
            lEndSYASINFO001 = gM46V999.gAppPath & "REMESSA\" & Trim(lSCoduser) & "YY.001"
            'Cria réplica do arquivo do usuário REDE.
            FileCopy lEndSYASINFO006, lEndSYASINFO001
        Else
            'Sai do procedimento - não existe log do usuário.
            lMsg = gM46V999.gAppPath & "REMESSA\" & Trim(lSCoduser) & "YY.006" & vbLf & vbLf & _
                   "Arquivo não localizado - a comunicação será iniciada porém não disponibilizará " & _
                   "atualização automaticamente."
            Call gM46V999.gpGraLog(2, lMsg)
            Exit Sub
        End If
        'Inicia a edição do log.
        If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & "REMESSA\" & Trim(lSCoduser) & "YY.001")) = True Then
            lNumSYASINFO001 = FreeFile
            Open lEndSYASINFO001 For Input Access Read As #lNumSYASINFO001
            If Err <> 0 Then
                Call gM46V999.gpGraLog(1, "FTE0079 - Erro ao abrir arquivo " & lEndSYASINFO001 & ": " & Err & _
                                        " - " & Error & vbLf & vbLf & "O sistema será encerrado.")
                End
            Else
                lNumSYASINFO006 = FreeFile
                Open lEndSYASINFO006 For Output Access Write As #lNumSYASINFO006
                Do While Not EOF(lNumSYASINFO001)
                    Line Input #lNumSYASINFO001, lRegistro
                    'Encontrou título da base de dados a ser alterado o caminho.
                    If Left(lRegistro, 7) = "USUARIO" Then
                        Print #lNumSYASINFO006, "USUARIO$REDE"
                    Else
                        'Reescreve sem alteções.
                        Print #lNumSYASINFO006, lRegistro
                    End If
                Loop
                Close #lNumSYASINFO006
                Close #lNumSYASINFO001
                Kill lEndSYASINFO001
            End If
        Else
            'Não existe réplica do arquivo.
            lMsg = gM46V999.gAppPath & "REMESSA\" & Trim(lSCoduser) & "YY.001" & vbLf & vbLf & _
                   "Arquivo não localizado - a comunicação será iniciada porém não disponibilizará " & _
                   "atualização automaticamente."
            Call gM46V999.gpGraLog(2, lMsg)
            Exit Sub
        End If
    Else
        'Usuário é estação - enviará arquivos porém não serão disponibilizados atualizações automáticas.
        Exit Sub
    End If
End Sub


'*******************************************
' gM46V999.mfVerRegPadrao agora contida no BAS 999
'*******************************************






