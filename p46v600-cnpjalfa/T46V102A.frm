VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmT46V102A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dados de proposta"
   ClientHeight    =   9390
   ClientLeft      =   2415
   ClientTop       =   1740
   ClientWidth     =   8970
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "T46V102A.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9390
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNumeroComplemento 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3300
      MaxLength       =   20
      TabIndex        =   146
      ToolTipText     =   "Complemento do endere�o do segurado (n�mero, andar, apartamento, conjunto, etc.)."
      Top             =   8220
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox picExcluir 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      Picture         =   "T46V102A.frx":000C
      ScaleHeight     =   255
      ScaleWidth      =   285
      TabIndex        =   145
      TabStop         =   0   'False
      Top             =   4530
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picIncluir 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "T46V102A.frx":0419
      ScaleHeight     =   255
      ScaleWidth      =   285
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   4530
      Visible         =   0   'False
      Width           =   285
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   300
      Left            =   1920
      TabIndex        =   142
      ToolTipText     =   "Novo bem."
      Top             =   4170
      Visible         =   0   'False
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Novo"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdExcluir 
      Height          =   300
      Left            =   840
      TabIndex        =   143
      ToolTipText     =   "Excluir o bem."
      Top             =   4170
      Visible         =   0   'False
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Excluir"
      ForeColor       =   8388608
   End
   Begin VB.PictureBox picChecked 
      Height          =   255
      Left            =   120
      Picture         =   "T46V102A.frx":0813
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   4170
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picUnchecked 
      Height          =   255
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   140
      Top             =   4170
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.ComboBox cboTipTelefone 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "T46V102A.frx":090D
      Left            =   1470
      List            =   "T46V102A.frx":090F
      Style           =   2  'Dropdown List
      TabIndex        =   139
      Top             =   8190
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   3090
      MaxLength       =   60
      TabIndex        =   138
      Top             =   8190
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSFlexGridLib.MSFlexGrid flxTelefones 
      Height          =   1575
      Left            =   0
      TabIndex        =   71
      Top             =   3510
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      AllowBigSelection=   0   'False
      HighLight       =   0
      AllowUserResizing=   1
   End
   Begin VB.TextBox txtDatEmis 
      ForeColor       =   &H00800000&
      Height          =   300
      HelpContextID   =   40
      Left            =   4125
      MaxLength       =   10
      TabIndex        =   104
      ToolTipText     =   "Data agendada para emiss�o"
      Top             =   7710
      Width           =   1140
   End
   Begin VB.TextBox txtDigCodSegurado 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   137
      ToolTipText     =   "C�digo do segurado."
      Top             =   285
      Width           =   300
   End
   Begin VB.TextBox txtReferencia 
      ForeColor       =   &H00800000&
      Height          =   300
      HelpContextID   =   40
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   95
      ToolTipText     =   "Refer�ncia do Corretor"
      Top             =   7350
      Width           =   1515
   End
   Begin VB.ComboBox CmbTipCotac 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "T46V102A.frx":0911
      Left            =   5625
      List            =   "T46V102A.frx":0913
      Style           =   2  'Dropdown List
      TabIndex        =   11
      ToolTipText     =   "Tipo de Cota��o (se houver)"
      Top             =   285
      Width           =   1575
   End
   Begin VB.TextBox txtNumSiscota 
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   7215
      MaxLength       =   12
      TabIndex        =   12
      ToolTipText     =   "N�mero da Cota��o (SisCota)"
      Top             =   285
      Width           =   1770
   End
   Begin VB.TextBox txtCodSegurado 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1230
      MaxLength       =   6
      TabIndex        =   8
      ToolTipText     =   "C�digo do segurado."
      Top             =   285
      Width           =   1455
   End
   Begin VB.TextBox txtSegurado 
      ForeColor       =   &H00800000&
      Height          =   285
      HelpContextID   =   44
      Left            =   1230
      MaxLength       =   60
      TabIndex        =   2
      ToolTipText     =   "Nome do segurado."
      Top             =   0
      Width           =   5205
   End
   Begin Threed.SSPanel panProduto 
      Height          =   4200
      Left            =   5400
      TabIndex        =   72
      ToolTipText     =   "Dados e op��es do produto."
      Top             =   5115
      Width           =   3600
      _Version        =   65536
      _ExtentX        =   6350
      _ExtentY        =   7408
      _StockProps     =   15
      ForeColor       =   -2147483630
      BackColor       =   12632256
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
      Begin VB.TextBox txtJuros 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   300
         HelpContextID   =   40
         Left            =   2385
         MaxLength       =   5
         TabIndex        =   108
         ToolTipText     =   "Juros."
         Top             =   330
         Width           =   825
      End
      Begin VB.TextBox txtQuaParcelas 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   300
         HelpContextID   =   40
         Left            =   2385
         MaxLength       =   3
         TabIndex        =   106
         ToolTipText     =   "Quantidade de parcelas."
         Top             =   30
         Width           =   1125
      End
      Begin Threed.SSPanel panProp 
         Height          =   945
         Index           =   8
         Left            =   45
         TabIndex        =   121
         ToolTipText     =   "Tem pagamento antecipado?"
         Top             =   2515
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   1667
         _StockProps     =   15
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
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin VB.TextBox txtValorPgAnt 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2280
            MaxLength       =   14
            TabIndex        =   128
            ToolTipText     =   "Valor Pagamento Antecipado"
            Top             =   600
            Width           =   1185
         End
         Begin VB.OptionButton optPagAntecipado 
            Caption         =   "A 30 dias"
            ForeColor       =   &H00800000&
            Height          =   230
            Index           =   2
            Left            =   1200
            TabIndex        =   125
            ToolTipText     =   "N�O tem pagamento antecipado."
            Top             =   380
            Width           =   1900
         End
         Begin Threed.SSCommand cmdPagAntBloqueto 
            Height          =   240
            Left            =   3200
            TabIndex        =   124
            TabStop         =   0   'False
            ToolTipText     =   "Exibir informa��es de pagamento antecipado."
            Top             =   70
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T46V102A.frx":0915
         End
         Begin VB.OptionButton optPagAntecipado 
            Caption         =   "Antecipada"
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   1
            Left            =   1200
            TabIndex        =   123
            ToolTipText     =   "SIM, tem pagamento antecipado."
            Top             =   75
            Width           =   1900
         End
         Begin Threed.SSPanel panProp 
            Height          =   850
            Index           =   9
            Left            =   45
            TabIndex        =   122
            Top             =   75
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   1499
            _StockProps     =   15
            Caption         =   "1� Parcela"
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
            BorderWidth     =   1
            FloodColor      =   8388608
         End
         Begin Threed.SSCommand cmdPagAntDebito 
            Height          =   240
            Left            =   3195
            TabIndex        =   126
            TabStop         =   0   'False
            ToolTipText     =   "Exibir informa��es de pagamento antecipado."
            Top             =   350
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T46V102A.frx":0C2F
         End
         Begin VB.Label lblPagAnt 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "R$"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2040
            TabIndex        =   127
            Top             =   630
            Visible         =   0   'False
            Width           =   210
         End
      End
      Begin Threed.SSPanel panProp 
         Height          =   1420
         Index           =   10
         Left            =   45
         TabIndex        =   111
         ToolTipText     =   "Forma de pagamento."
         Top             =   630
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   2505
         _StockProps     =   15
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
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin VB.OptionButton optForPagamento 
            Caption         =   "Cart�o de Cr�dito"
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   3
            Left            =   1200
            TabIndex        =   147
            ToolTipText     =   "Pagamento atrav�s de d�bito em conta corrente."
            Top             =   1100
            Width           =   1875
         End
         Begin VB.OptionButton optForPagamento 
            Caption         =   "D�bito autom�tico com 1� em boleto"
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   2
            Left            =   1200
            TabIndex        =   116
            ToolTipText     =   "Pagamento atrav�s de d�bito em conta corrente."
            Top             =   700
            Width           =   1875
         End
         Begin VB.OptionButton optForPagamento 
            Caption         =   "Carn�"
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   0
            Left            =   1200
            TabIndex        =   112
            ToolTipText     =   "Pagamento atrav�s de carn�."
            Top             =   120
            Width           =   750
         End
         Begin VB.OptionButton optForPagamento 
            Caption         =   "D�bito em c/c"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   114
            ToolTipText     =   "Pagamento atrav�s de d�bito em conta corrente."
            Top             =   430
            Width           =   1400
         End
         Begin Threed.SSCommand cmdForPagCarne 
            Height          =   240
            Left            =   3200
            TabIndex        =   113
            TabStop         =   0   'False
            ToolTipText     =   "Exibir informa��es de forma de pagamento."
            Top             =   100
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T46V102A.frx":0F49
         End
         Begin Threed.SSPanel panProp 
            Height          =   1370
            Index           =   11
            Left            =   75
            TabIndex        =   110
            Top             =   45
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   2417
            _StockProps     =   15
            Caption         =   "Forma de pagamento"
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
            BorderWidth     =   1
         End
         Begin Threed.SSCommand cmdForPagDebito 
            Height          =   240
            Left            =   3200
            TabIndex        =   115
            TabStop         =   0   'False
            ToolTipText     =   "Exibir informa��es de forma de pagamento."
            Top             =   430
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T46V102A.frx":1263
         End
         Begin Threed.SSCommand cmdForPagDebito1oBoleto 
            Height          =   240
            Left            =   3200
            TabIndex        =   117
            TabStop         =   0   'False
            ToolTipText     =   "Exibir informa��es de forma de pagamento."
            Top             =   745
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T46V102A.frx":157D
         End
         Begin Threed.SSCommand cmdForPagCartao 
            Height          =   240
            Left            =   3200
            TabIndex        =   148
            TabStop         =   0   'False
            ToolTipText     =   "Exibir informa��es de forma de pagamento."
            Top             =   1100
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T46V102A.frx":1897
         End
      End
      Begin Threed.SSPanel panProp 
         Height          =   705
         Index           =   3
         Left            =   45
         TabIndex        =   129
         ToolTipText     =   "Forma de pagamento."
         Top             =   3465
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6174
         _ExtentY        =   1235
         _StockProps     =   15
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
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin VB.OptionButton optVencimento 
            Caption         =   "N�o"
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   1
            Left            =   1200
            TabIndex        =   133
            ToolTipText     =   "N�o ajustar a data das parcelas"
            Top             =   345
            Width           =   1400
         End
         Begin VB.OptionButton optVencimento 
            Caption         =   "Sim"
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   0
            Left            =   1200
            TabIndex        =   131
            ToolTipText     =   "Data das parcelas ajustadas."
            Top             =   120
            Width           =   750
         End
         Begin Threed.SSCommand cmdAjustarDatas 
            Height          =   240
            Left            =   3200
            TabIndex        =   132
            TabStop         =   0   'False
            ToolTipText     =   "Exibir informa��es de forma de pagamento."
            Top             =   235
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T46V102A.frx":1897
         End
         Begin Threed.SSPanel panProp 
            Height          =   645
            Index           =   4
            Left            =   45
            TabIndex        =   130
            Top             =   45
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   1138
            _StockProps     =   15
            Caption         =   "Ajustar data das parcelas"
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
            BorderWidth     =   1
         End
      End
      Begin Threed.SSCommand cmdJuros 
         Height          =   240
         Left            =   3240
         TabIndex        =   109
         TabStop         =   0   'False
         ToolTipText     =   "Exibir informa��es de forma de pagamento."
         Top             =   360
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   423
         _StockProps     =   78
         ForeColor       =   12582912
         Picture         =   "T46V102A.frx":1BB1
      End
      Begin Threed.SSPanel panProp 
         Height          =   465
         Index           =   6
         Left            =   45
         TabIndex        =   118
         ToolTipText     =   "Forma de pagamento."
         Top             =   2050
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   820
         _StockProps     =   15
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
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSCommand cmdRestituicao 
            Height          =   240
            Left            =   3200
            TabIndex        =   120
            TabStop         =   0   'False
            ToolTipText     =   "Exibir informa��es de forma de pagamento."
            Top             =   100
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T46V102A.frx":1ECB
         End
         Begin Threed.SSPanel panProp 
            Height          =   255
            Index           =   7
            Left            =   75
            TabIndex        =   119
            Top             =   135
            Width           =   3045
            _Version        =   65536
            _ExtentX        =   5371
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Meio para dev cred futuros - Restitui��o"
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
            BorderWidth     =   1
         End
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Juros das parcelas: "
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   16
         Left            =   45
         TabIndex        =   107
         ToolTipText     =   "Juros das parcelas"
         Top             =   330
         Width           =   2340
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quantidade de parcelas: "
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   24
         Left            =   45
         TabIndex        =   105
         ToolTipText     =   "Quantidade de parcelas"
         Top             =   30
         Width           =   2340
      End
   End
   Begin Threed.SSPanel panCorretores 
      Height          =   1245
      Left            =   0
      TabIndex        =   73
      Top             =   5115
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9349
      _ExtentY        =   2205
      _StockProps     =   15
      BackColor       =   12632256
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
      Begin VB.TextBox txtCorretor 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   50
         TabIndex        =   80
         ToolTipText     =   "Terceiro corretor."
         Top             =   890
         Width           =   4500
      End
      Begin VB.TextBox txtCorretor 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   50
         TabIndex        =   78
         ToolTipText     =   "Segundo corretor."
         Top             =   600
         Width           =   4500
      End
      Begin VB.TextBox txtCorretor 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   50
         TabIndex        =   76
         ToolTipText     =   "Primeiro corretor."
         Top             =   320
         Width           =   4500
      End
      Begin VB.TextBox txtCorComissao 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   4550
         MaxLength       =   5
         TabIndex        =   81
         Text            =   "0,00"
         ToolTipText     =   "Comiss�o do terceiro corretor."
         Top             =   890
         Width           =   700
      End
      Begin VB.TextBox txtCorComissao 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   4550
         MaxLength       =   5
         TabIndex        =   79
         Text            =   "0,00"
         ToolTipText     =   "Comiss�o do segundo corretor."
         Top             =   600
         Width           =   700
      End
      Begin VB.TextBox txtCorComissao 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   4550
         MaxLength       =   5
         TabIndex        =   77
         Text            =   "0,00"
         ToolTipText     =   "Comiss�o do primeiro corretor."
         Top             =   320
         Width           =   700
      End
      Begin VB.Label lblProp 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Corretor"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   13
         Left            =   50
         TabIndex        =   74
         Top             =   30
         Width           =   4500
      End
      Begin VB.Label lblProp 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   14
         Left            =   4550
         TabIndex        =   75
         Top             =   30
         Width           =   700
      End
   End
   Begin Threed.SSPanel panEndereco 
      Height          =   1635
      Left            =   0
      TabIndex        =   43
      ToolTipText     =   "Dados do segurado."
      Top             =   1890
      Width           =   9000
      _Version        =   65536
      _ExtentX        =   15875
      _ExtentY        =   2884
      _StockProps     =   15
      ForeColor       =   -2147483630
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
      Begin VB.TextBox txtComplemento 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3960
         MaxLength       =   13
         TabIndex        =   53
         ToolTipText     =   "Complemento do endere�o do segurado (n�mero, andar, apartamento, conjunto, etc.)."
         Top             =   330
         Width           =   4905
      End
      Begin VB.TextBox txtNumero 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   51
         ToolTipText     =   "Complemento do endere�o do segurado (n�mero, andar, apartamento, conjunto, etc.)."
         Top             =   330
         Width           =   885
      End
      Begin VB.TextBox txtEmail 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   59
         ToolTipText     =   "E-mail do segurado."
         Top             =   930
         Width           =   2490
      End
      Begin VB.TextBox txtnumregistro 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7350
         MaxLength       =   10
         TabIndex        =   70
         ToolTipText     =   "Nr. do registro da categoria."
         Top             =   1260
         Width           =   1560
      End
      Begin VB.ComboBox cboRegCons 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V102A.frx":21E5
         Left            =   5055
         List            =   "T46V102A.frx":21E7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   68
         ToolTipText     =   "Sigla do conselho da categoria."
         Top             =   1260
         Width           =   1335
      End
      Begin Threed.SSPanel panProp 
         Height          =   285
         Index           =   1
         Left            =   2310
         TabIndex        =   63
         ToolTipText     =   "Endere�o de cobran�a � diferente do endere�o do segurado?"
         Top             =   1275
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   503
         _StockProps     =   15
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
         BorderWidth     =   0
         BevelOuter      =   1
         Alignment       =   6
         Begin VB.OptionButton optEndCobranca 
            Caption         =   "&N�o"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   645
            TabIndex        =   65
            ToolTipText     =   "N�O, endere�o de cobran�a � igual ao endere�o do segurado."
            Top             =   50
            Width           =   600
         End
         Begin VB.OptionButton optEndCobranca 
            Caption         =   "&Sim"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   45
            TabIndex        =   64
            ToolTipText     =   "SIM, endere�o de cobran�a � diferente do endere�o do segurado."
            Top             =   50
            Width           =   600
         End
         Begin Threed.SSCommand cmdEndCobranca 
            Height          =   255
            Left            =   1275
            TabIndex        =   66
            TabStop         =   0   'False
            ToolTipText     =   "Exibir endere�o de cobran�a."
            Top             =   15
            Width           =   255
            _Version        =   65536
            _ExtentX        =   441
            _ExtentY        =   441
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T46V102A.frx":21E9
         End
      End
      Begin VB.TextBox txtCidade 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   55
         ToolTipText     =   "Cidade do endere�o do segurado."
         Top             =   630
         Width           =   2500
      End
      Begin VB.ComboBox cboUF 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V102A.frx":2503
         Left            =   5040
         List            =   "T46V102A.frx":2558
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   61
         ToolTipText     =   "Unidade da federa��o do endere�o do segurado."
         Top             =   930
         Width           =   810
      End
      Begin VB.TextBox txtBairro 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5040
         MaxLength       =   30
         TabIndex        =   57
         ToolTipText     =   "Bairro do endere�o do segurado."
         Top             =   630
         Width           =   3900
      End
      Begin VB.TextBox txtTipLogradouro 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3345
         MaxLength       =   3
         TabIndex        =   48
         ToolTipText     =   "Tipo do logradouro do endere�o do segurado (Rua, Av., etc.)."
         Top             =   30
         Width           =   600
      End
      Begin VB.TextBox txtCEP 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   46
         ToolTipText     =   "CEP do endere�o do segurado (nnnnn-nnn)."
         Top             =   30
         Width           =   900
      End
      Begin Threed.SSCommand cmdCEP 
         Height          =   250
         HelpContextID   =   42
         Left            =   50
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Busca endere�o do segurado."
         Top             =   50
         Width           =   250
         _Version        =   65536
         _ExtentX        =   441
         _ExtentY        =   441
         _StockProps     =   78
         ForeColor       =   12582912
         Picture         =   "T46V102A.frx":25C8
      End
      Begin VB.TextBox txtNomLogradouro 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3945
         MaxLength       =   50
         TabIndex        =   49
         ToolTipText     =   "Nome do logradouro do endere�o do segurado."
         Top             =   30
         Width           =   5010
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Complemento: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   31
         Left            =   2340
         TabIndex        =   52
         ToolTipText     =   "Complemento do endere�o do segurado."
         Top             =   330
         Width           =   1605
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E-mail: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   30
         Left            =   30
         TabIndex        =   58
         ToolTipText     =   "Telefone(s) do endere�o do segurado."
         Top             =   930
         Width           =   1395
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nr.Registro:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   29
         Left            =   6420
         TabIndex        =   69
         ToolTipText     =   "Registro da categoria."
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cat.Conselho:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   28
         Left            =   3960
         TabIndex        =   67
         ToolTipText     =   "Registro no conselho da categoria."
         Top             =   1275
         Width           =   1125
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Endere�o cobran�a diferente?"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   9
         Left            =   45
         TabIndex        =   62
         ToolTipText     =   "Endere�o de cobran�a � diferente do endere�o do segurado?"
         Top             =   1275
         Width           =   2250
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UF: "
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   6
         Left            =   3945
         TabIndex        =   60
         ToolTipText     =   "Unidade da federa��o do endere�o do segurado."
         Top             =   930
         Width           =   1080
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cidade: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   5
         Left            =   45
         TabIndex        =   54
         ToolTipText     =   "Cidade do endere�o do segurado."
         Top             =   630
         Width           =   1395
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bairro: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   3945
         TabIndex        =   56
         ToolTipText     =   "Bairro do endere�o do segurado."
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   45
         TabIndex        =   50
         ToolTipText     =   "Complemento do endere�o do segurado."
         Top             =   315
         Width           =   1395
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Logradouro: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   2340
         TabIndex        =   47
         ToolTipText     =   "Logradouro do endere�o do segurado."
         Top             =   30
         Width           =   1005
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CEP: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   7
         Left            =   300
         TabIndex        =   45
         ToolTipText     =   "CEP do endere�o do segurado."
         Top             =   30
         Width           =   1140
      End
   End
   Begin Threed.SSPanel panProp 
      Height          =   945
      Index           =   0
      Left            =   0
      TabIndex        =   82
      Top             =   6390
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   1667
      _StockProps     =   15
      ForeColor       =   -2147483630
      BackColor       =   12632256
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
      Begin Threed.SSCommand cmdUniProdutiva 
         Height          =   285
         HelpContextID   =   42
         Left            =   4950
         TabIndex        =   85
         TabStop         =   0   'False
         ToolTipText     =   "Seleciona unidade produtiva."
         Top             =   30
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   503
         _StockProps     =   78
         ForeColor       =   12582912
         Enabled         =   0   'False
         Picture         =   "T46V102A.frx":28E2
      End
      Begin Threed.SSCommand cmdProdutor 
         Height          =   285
         HelpContextID   =   42
         Left            =   4950
         TabIndex        =   88
         TabStop         =   0   'False
         ToolTipText     =   "Seleciona produtor."
         Top             =   330
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   503
         _StockProps     =   78
         ForeColor       =   12582912
         Enabled         =   0   'False
         Picture         =   "T46V102A.frx":2BFC
      End
      Begin Threed.SSCommand cmdCosseguro 
         Height          =   285
         Left            =   4950
         TabIndex        =   93
         TabStop         =   0   'False
         ToolTipText     =   "Exibir informa��es de cosseguro."
         Top             =   600
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   503
         _StockProps     =   78
         ForeColor       =   12582912
         Picture         =   "T46V102A.frx":2F16
      End
      Begin Threed.SSPanel panProp 
         Height          =   285
         Index           =   5
         Left            =   1900
         TabIndex        =   90
         Top             =   600
         Width           =   3050
         _Version        =   65536
         _ExtentX        =   5380
         _ExtentY        =   503
         _StockProps     =   15
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
         BorderWidth     =   0
         BevelOuter      =   1
         Alignment       =   6
         Begin VB.OptionButton optCosseguro 
            Caption         =   "&N�o"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   650
            TabIndex        =   92
            ToolTipText     =   "N�O tem cess�o de cosseguro."
            Top             =   50
            Width           =   600
         End
         Begin VB.OptionButton optCosseguro 
            Caption         =   "&Sim"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   50
            TabIndex        =   91
            ToolTipText     =   "SIM, tem cess�o de cosseguro."
            Top             =   50
            Width           =   600
         End
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cess�o de cosseguro? "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   12
         Left            =   50
         TabIndex        =   89
         ToolTipText     =   "Cess�o de cosseguro?"
         Top             =   600
         Width           =   1850
      End
      Begin VB.Label lblNomProdutor 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   950
         TabIndex        =   87
         ToolTipText     =   "Nome do produtor."
         Top             =   330
         Width           =   4000
      End
      Begin VB.Label lblProdutor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produtor: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   50
         TabIndex        =   86
         Top             =   330
         Width           =   900
      End
      Begin VB.Label lblUnidade 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidade: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   45
         TabIndex        =   83
         ToolTipText     =   "Unidade produtiva."
         Top             =   45
         Width           =   900
      End
      Begin VB.Label lblNomUnidade 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   950
         TabIndex        =   84
         ToolTipText     =   "Nome da unidade produtiva."
         Top             =   30
         Width           =   4000
      End
   End
   Begin Threed.SSPanel panProp 
      Height          =   285
      Index           =   15
      Left            =   7250
      TabIndex        =   4
      Top             =   0
      Width           =   1750
      _Version        =   65536
      _ExtentX        =   3087
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "4"
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
      BorderWidth     =   0
      BevelOuter      =   1
      Alignment       =   6
      Begin VB.OptionButton optTipPessoa 
         Caption         =   "&F�sica"
         ForeColor       =   &H00800000&
         Height          =   200
         Index           =   0
         Left            =   50
         TabIndex        =   5
         ToolTipText     =   "Pessoa f�sica."
         Top             =   50
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton optTipPessoa 
         Caption         =   "&Jur�dica"
         ForeColor       =   &H00800000&
         Height          =   200
         Index           =   1
         Left            =   800
         TabIndex        =   6
         ToolTipText     =   "Pessoa jur�dica."
         Top             =   50
         Width           =   900
      End
   End
   Begin Threed.SSCommand cmdExtra 
      Height          =   240
      Index           =   2
      Left            =   210
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   7665
      Visible         =   0   'False
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   78
      ForeColor       =   12582912
      Picture         =   "T46V102A.frx":3230
   End
   Begin Threed.SSCommand cmdExtra 
      Height          =   240
      Index           =   1
      Left            =   0
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   7665
      Visible         =   0   'False
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   78
      ForeColor       =   12582912
      Picture         =   "T46V102A.frx":354A
   End
   Begin Threed.SSCommand cmdSiscad 
      Height          =   300
      Index           =   0
      Left            =   3015
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Buscar dados SISCAD"
      Top             =   285
      Width           =   360
      _Version        =   65536
      _ExtentX        =   635
      _ExtentY        =   529
      _StockProps     =   78
      ForeColor       =   12582912
      Picture         =   "T46V102A.frx":3864
   End
   Begin Threed.SSCommand cmdSiscad 
      Height          =   300
      Index           =   1
      Left            =   3390
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Atualizar dados do segurado."
      Top             =   285
      Width           =   360
      _Version        =   65536
      _ExtentX        =   635
      _ExtentY        =   529
      _StockProps     =   78
      ForeColor       =   12582912
      Picture         =   "T46V102A.frx":3B7E
   End
   Begin Threed.SSCommand cmdVoltar 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Volta ao item."
      Top             =   8220
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Voltar ao item"
      ForeColor       =   8388608
   End
   Begin Threed.SSPanel panPesFisica 
      Height          =   1305
      Left            =   0
      TabIndex        =   20
      Top             =   600
      Width           =   9000
      _Version        =   65536
      _ExtentX        =   15875
      _ExtentY        =   2302
      _StockProps     =   15
      BackColor       =   13160660
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
      Begin VB.TextBox txtDataNascimento 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   36
         ToolTipText     =   "Data de nascimento do segurado."
         Top             =   630
         Width           =   1470
      End
      Begin VB.ComboBox cboEstadoCivil 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V102A.frx":3CD8
         Left            =   5505
         List            =   "T46V102A.frx":3CDA
         Style           =   2  'Dropdown List
         TabIndex        =   38
         ToolTipText     =   "Estado civil do segurado."
         Top             =   630
         Width           =   3435
      End
      Begin VB.ComboBox cboRendaMensal 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V102A.frx":3CDC
         Left            =   4755
         List            =   "T46V102A.frx":3CDE
         Style           =   2  'Dropdown List
         TabIndex        =   42
         ToolTipText     =   "Renda Mensal - Pessoa F�sica"
         Top             =   960
         Width           =   4185
      End
      Begin VB.TextBox txtIdeEmpresa 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2145
         MaxLength       =   30
         TabIndex        =   40
         ToolTipText     =   "N�mero do documento de identifica��o do segurado."
         Top             =   960
         Width           =   2050
      End
      Begin VB.TextBox txtIdeDatEmissao 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7485
         MaxLength       =   10
         TabIndex        =   26
         ToolTipText     =   "Data de emiss�o do documento de identifica��o do segurado."
         Top             =   50
         Width           =   1470
      End
      Begin VB.TextBox txtIdeOrgExpedidor 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5910
         MaxLength       =   20
         TabIndex        =   30
         ToolTipText     =   "�rg�o expedidor do documento de identifica��o do segurado."
         Top             =   345
         Width           =   1125
      End
      Begin VB.TextBox txtIdeNumDocumento 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2985
         MaxLength       =   20
         TabIndex        =   28
         ToolTipText     =   "N�mero do documento de identifica��o do segurado."
         Top             =   350
         Width           =   1605
      End
      Begin VB.ComboBox cboIdeTipDocumento 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V102A.frx":3CE0
         Left            =   2985
         List            =   "T46V102A.frx":3CE2
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "Tipo do documento de identifica��o do segurado."
         Top             =   50
         Width           =   3100
      End
      Begin VB.TextBox txtCPF 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   50
         MaxLength       =   14
         TabIndex        =   22
         ToolTipText     =   "CPF do segurado."
         Top             =   660
         Width           =   1250
      End
      Begin Threed.SSPanel panProp 
         Height          =   285
         Index           =   2
         Left            =   7530
         TabIndex        =   32
         Top             =   345
         Width           =   1420
         _Version        =   65536
         _ExtentX        =   2505
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "4"
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
         BorderWidth     =   0
         BevelOuter      =   1
         Alignment       =   6
         Begin VB.OptionButton optSexo 
            Caption         =   "Fem"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   735
            TabIndex        =   34
            ToolTipText     =   "Pessoa jur�dica."
            Top             =   50
            Width           =   615
         End
         Begin VB.OptionButton optSexo 
            Caption         =   "Masc"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   50
            TabIndex        =   33
            ToolTipText     =   "Pessoa f�sica."
            Top             =   50
            Width           =   735
         End
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Data de nascimento: "
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   33
         Left            =   1350
         TabIndex        =   35
         ToolTipText     =   "Data de nascimento do segurado."
         Top             =   630
         Width           =   1665
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado Civil: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   32
         Left            =   4470
         TabIndex        =   37
         ToolTipText     =   "Estado civil do segurado."
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Renda: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   27
         Left            =   4080
         TabIndex        =   41
         ToolTipText     =   "�rg�o expedidor do documento de identifica��o do segurado."
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblProp 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sexo: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   11
         Left            =   7035
         TabIndex        =   31
         ToolTipText     =   "�rg�o expedidor do documento de identifica��o do segurado."
         Top             =   345
         Width           =   495
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Profiss�o: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   10
         Left            =   1350
         TabIndex        =   39
         ToolTipText     =   "N�mero do documento de identifica��o do segurado."
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Data de emiss�o: "
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   23
         Left            =   6075
         TabIndex        =   25
         ToolTipText     =   "Data de emiss�o do documento de identifica��o do segurado."
         Top             =   45
         Width           =   1395
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N�m. do documento: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   21
         Left            =   1350
         TabIndex        =   27
         ToolTipText     =   "N�mero do documento de identifica��o do segurado."
         Top             =   345
         Width           =   1680
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo do documento: "
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   20
         Left            =   1350
         TabIndex        =   23
         ToolTipText     =   "Tipo do documento de identifica��o do segurado."
         Top             =   45
         Width           =   1680
      End
      Begin VB.Label lblProp 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CPF"
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   18
         Left            =   45
         TabIndex        =   21
         ToolTipText     =   "CPF do segurado."
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "�rg�o expedidor: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   22
         Left            =   4575
         TabIndex        =   29
         ToolTipText     =   "�rg�o expedidor do documento de identifica��o do segurado."
         Top             =   345
         Width           =   1335
      End
   End
   Begin Threed.SSPanel panPesJuridica 
      Height          =   945
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   9000
      _Version        =   65536
      _ExtentX        =   15875
      _ExtentY        =   1676
      _StockProps     =   15
      BackColor       =   13160660
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
      Begin VB.ComboBox cboFaturamento 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V102A.frx":3CE4
         Left            =   3400
         List            =   "T46V102A.frx":3CE6
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Faturamento - Pessoa Jur�dica"
         Top             =   480
         Width           =   5505
      End
      Begin VB.TextBox txtCNPJ 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   50
         MaxLength       =   18
         TabIndex        =   15
         ToolTipText     =   "CNPJ do segurado."
         Top             =   450
         Width           =   1600
      End
      Begin VB.ComboBox cboRamAtividade 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V102A.frx":3CE8
         Left            =   3400
         List            =   "T46V102A.frx":3CEA
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Ramo de atividade do segurado."
         Top             =   150
         Width           =   5500
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Faturamento: "
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   25
         Left            =   1800
         TabIndex        =   18
         ToolTipText     =   "Faturamento"
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label lblProp 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CNPJ"
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   19
         Left            =   50
         TabIndex        =   14
         ToolTipText     =   "CNPJ do segurado."
         Top             =   150
         Width           =   1600
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ramo de atividade: "
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   16
         ToolTipText     =   "Ramo de atividade do segurado."
         Top             =   150
         Width           =   1605
      End
   End
   Begin Threed.SSPanel panProlabore 
      Height          =   285
      Index           =   0
      Left            =   2610
      TabIndex        =   96
      Top             =   7350
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   503
      _StockProps     =   15
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
      BorderWidth     =   0
      BevelOuter      =   1
      Alignment       =   6
      Begin VB.OptionButton optProlabore 
         Caption         =   "&Sim"
         ForeColor       =   &H00800000&
         Height          =   200
         Index           =   0
         Left            =   975
         TabIndex        =   98
         ToolTipText     =   "SIM, tem prolaborista."
         Top             =   50
         Width           =   600
      End
      Begin VB.OptionButton optProlabore 
         Caption         =   "&N�o"
         ForeColor       =   &H00800000&
         Height          =   200
         Index           =   1
         Left            =   1605
         TabIndex        =   99
         ToolTipText     =   "N�O tem prolaborista."
         Top             =   50
         Width           =   600
      End
      Begin Threed.SSCommand cmdProlabore 
         Height          =   285
         HelpContextID   =   42
         Left            =   2340
         TabIndex        =   100
         TabStop         =   0   'False
         ToolTipText     =   "Exibir informa��es dos prolaboristas."
         Top             =   0
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   503
         _StockProps     =   78
         ForeColor       =   12582912
         Picture         =   "T46V102A.frx":3CEC
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prolabore?"
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   97
         ToolTipText     =   "Refer�ncia"
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.TextBox txtTelefones 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   102
      ToolTipText     =   "Telefone(s) do endere�o do segurado."
      Top             =   7710
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data para emiss�o:"
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   0
      Left            =   2640
      TabIndex        =   103
      ToolTipText     =   "Refer�ncia"
      Top             =   7710
      Width           =   1500
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Refer�ncia: "
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   27
      Left            =   0
      TabIndex        =   94
      ToolTipText     =   "Refer�ncia"
      Top             =   7350
      Width           =   1080
   End
   Begin VB.Label lblProp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cota��o: "
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   15
      Left            =   4860
      TabIndex        =   136
      ToolTipText     =   "Cota��o"
      Top             =   285
      Width           =   765
   End
   Begin VB.Label lblProp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C�d. segurado: "
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   26
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "C�digo e nome do segurado."
      Top             =   285
      Width           =   1215
   End
   Begin VB.Label lblProp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pessoa: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   17
      Left            =   6450
      TabIndex        =   3
      Top             =   0
      Width           =   800
   End
   Begin VB.Label lblProp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Segurado: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "C�digo e nome do segurado."
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblProp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telefones:"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   8
      Left            =   30
      TabIndex        =   101
      ToolTipText     =   "Telefone(s) do endere�o do segurado."
      Top             =   7710
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Menu mnuExtra 
      Caption         =   "Extra"
      Visible         =   0   'False
      Begin VB.Menu mnuDadSegurado 
         Caption         =   "Dados do &segurado"
      End
      Begin VB.Menu mnuEndCobranca 
         Caption         =   "E&ndere�o de cobran�a"
      End
      Begin VB.Menu mnuEndSegurado 
         Caption         =   "&Endere�o do segurado"
      End
      Begin VB.Menu mnuProdutor 
         Caption         =   "Pr&odutor"
      End
      Begin VB.Menu mnuUniProdutiva 
         Caption         =   "&Unidade produtiva"
      End
   End
End
Attribute VB_Name = "frmT46V102A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mEstaCarregando As Boolean
Private mAltEndereco As Boolean         'Indica que houve altera��o no endere�o.
Private mAntCEP As String               'Anterior: CEP.
Private mAntControle As Control         'Anterior: controle.
Private strCorretor As String
Private strCodigoCorretor As String

Private Const mMaxLength_DDD = 2
Private Const mMaxLength_Telefone = 10
Private Const mMaxLength_Ramal = 4      'Alterado de 10 para 4 por compatibilidade com a web
Private Const mMaxLength_Obs = 255

Private Sub mpUnidadeSegurado(ByRef pUnidade As Integer, ByRef pMensagem As String)
    Dim lbdP0044700 As ADODB.Connection    'Arquivo de conex�o.
    Dim lrsP0044700 As clsYasRecordSet  'Registro P0044700.
    Dim lSelect As String

    On Error GoTo Erro
    pUnidade = 0
    lSelect = "SELECT * FROM YTAB_SEG_UNID WHERE cod_segurado = " & txtCodSegurado.Text
    If gM46V999.gfAbrBasDados("P0044700", gM46V999.gTipP0044700, gM46V999.gEndP0044700, lbdP0044700, pMensagem) = False Then
        Exit Sub
    End If
    If gM46V999.gfObtRegistro(lbdP0044700, lSelect, lrsP0044700, pMensagem) = False Then
        Exit Sub
    End If
    If Not lrsP0044700.EOF Then
        pUnidade = lrsP0044700("Cod_Unidade")
    End If
    Exit Sub

Erro:
    pUnidade = 0
End Sub
Private Function mfPosUniProProdutor(ByVal pCodCorretor As String, ByRef pMensagem As String) As Boolean
'Procedure: posiciona unidade produtiva e produtor.

'Par�metro de entrada...pCodCorretor...C�digo do corretor.
'Par�metro de sa�da.....pMensagem......Mensagem.

    Dim lCodProdutor As Integer    'C�digo do produtor.
    Dim lCodUniProdutiva As Integer    'C�digo da unidade produtiva.
    Dim lCodUniProdAux As Integer    'C�digo da unidade produtiva auxiliar
    Dim lNomProdutor As String    'Nome do produtor.
    Dim lNomUniProdutiva As String    'Nome da unidade produtiva.
    Dim lTotalRegistros As Integer    'Total de registros na tabela YTAB_UNID_CORR_PROD
    Dim lListarUnidade As Boolean    'Indicador se � para listar as unidades para o usu�rios escolher
    mfPosUniProProdutor = False

    '1. Limpa campos.
    lblNomUnidade.Tag = " "
    lblNomUnidade.Caption = " "
    lblNomProdutor.Tag = " "
    lblNomProdutor.Caption = " "

    '2. Obt�m dados.
    If gM46V999.gfObtUniProdutiva(pCodCorretor, lCodUniProdutiva, lCodProdutor, pMensagem, lTotalRegistros) = False Then
        Exit Function
    End If

    lListarUnidade = True
    lCodUniProdAux = 0
    If gM46V999.gfPreenchido(txtCodSegurado.Text) = True And Trim(txtCodSegurado.Text) <> "0" Then
        Call mpUnidadeSegurado(lCodUniProdAux, pMensagem)
        pMensagem = ""
        If lCodUniProdAux <> 0 Then
            lListarUnidade = False
            lCodUniProdutiva = lCodUniProdAux
        End If
    End If

    '****-- 0501113 - Disponibilizada op��o de sele��o da unidade quando corretora com mais de um registro
    '****--           na TAB_UNID_CORR_PROD para todos os usu�rios do Syas.
    'Select Case gPerAcesso
    '    Case 1, 2, 5
    If lTotalRegistros > 1 Then
        If lListarUnidade = True Then
            Call frmT46V114A.mpInstanciarForm(CLng(pCodCorretor), lCodUniProdutiva, lCodProdutor)
        End If
    End If
    'End Select
    If gM46V999.gfNomUniProdutiva(lCodUniProdutiva, lNomUniProdutiva, pMensagem) = False Then
        Exit Function
    End If
    If gM46V999.gfNomProdutor(lCodProdutor, lNomProdutor, pMensagem) = False Then
        Exit Function
    End If

    '3. Posiciona informa��es.
    lblNomUnidade.Tag = lCodUniProdutiva
    lblNomUnidade.Caption = lNomUniProdutiva
    lblNomProdutor.Tag = lCodProdutor
    lblNomProdutor.Caption = lNomProdutor

    mfPosUniProProdutor = True
End Function
Private Sub mpVerEndSegurado()
'Procedure: verifica endere�o do segurado.

    Dim lEndereco As String    'Endere�o.
    Dim lTamanho As Integer    'Tamanho.

'1. Se n�o houve altera��o no endere�o, n�o verificar.
    If mAltEndereco = False Then
        Exit Sub
    End If

    '2. Verifica endere�o.
    If gM46V999.gfPreenchido(txtTipLogradouro.Text) Then
        lEndereco = Trim$(txtTipLogradouro.Text)
    End If
    If gM46V999.gfPreenchido(txtNomLogradouro.Text) Then
        If gM46V999.gfPreenchido(lEndereco) Then
            lEndereco = lEndereco & " " & Trim$(txtNomLogradouro.Text)
        Else
            lEndereco = Trim$(txtNomLogradouro.Text)
        End If
    End If
    If gM46V999.gfPreenchido(lEndereco) Then
        lEndereco = lEndereco & " " & Trim$(txtNumero.Text) & " " & Trim$(txtComplemento.Text)
    Else
        lEndereco = Trim$(txtNumero.Text) & " " & Trim$(txtComplemento.Text)
    End If
    lTamanho = Len(lEndereco)
    If lTamanho > 50 Then
        Call gM46V999.gpGraLog(1, "AGE0035 - " & gM46V999.gPrefixo & "Endere�o completo do segurado: " & vbLf & vbLf & _
                                  UCase(lEndereco) & vbLf & vbLf & "O endere�o completo est� com " & lTamanho & _
                                " posi��es (tamanho m�ximo permitido no sistema = 50).")
    End If
    mAltEndereco = False
End Sub

Private Sub mnuProlaboristas()
    If optProlabore(0).Value = False Then
        If gM46V999.gDesabilitarForms = True Then Exit Sub
        optProlabore(0).Value = True
    Else
        'If gM46V999.gfDisponivel("frmT46V180A") Then
        'TODO:
        'Limpar os dados dos Prolaboristas. frmT46V180A.Show
        'Else
        frmT46V180A.Show vbModal, Me
        'End If
    End If
End Sub

Private Sub cboIdeTipDocumento_Click()
    Call gpHabSalvar
End Sub
Private Sub cboIdeTipDocumento_GotFocus()
    Set mAntControle = ActiveControl
End Sub
Private Sub cboIdeTipDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub cboIdeTipDocumento_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then     'Teclou backspace.
        cboIdeTipDocumento.ListIndex = -1
    End If
End Sub
Private Sub cboEstadoCivil_Click()
    Call gpHabSalvar
End Sub
Private Sub cboEstadoCivil_GotFocus()
    Set mAntControle = ActiveControl
End Sub
Private Sub cboEstadoCivil_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub cboEstadoCivil_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then     'Teclou backspace.
        cboEstadoCivil.ListIndex = -1
    End If
End Sub
Private Sub cboRamAtividade_Click()
    Call gpHabSalvar
End Sub
Private Sub cboRamAtividade_GotFocus()
    Set mAntControle = ActiveControl
End Sub
Private Sub cboRamAtividade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub cboRamAtividade_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then     'Teclou backspace.
        cboRamAtividade.ListIndex = -1
    End If
End Sub
'Ficha 1000307
Private Sub cboRegCons_Change()
    Call mpCarregarRegConselho
    Call mpPosReg
End Sub

'Ficha 1000307
Private Sub cboRegCons_Click()
    Call gpHabSalvar
End Sub

'Ficha 1000307
Private Sub cboRegCons_GotFocus()
    Set mAntControle = ActiveControl
End Sub
'Ficha 1000307
Private Sub cboRegCons_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
'Ficha 1000307
Private Sub cboRegCons_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then     'Teclou backspace.
        cboRegCons.ListIndex = -1
    End If

End Sub

Private Sub cboRegCons_LostFocus()
    If cboRegCons.ListIndex > -1 Then
        cboRegCons.Tag = cboRegCons.ItemData(cboRegCons.ListIndex)
    Else
        cboRegCons.Tag = 0
    End If
End Sub

Private Sub cboTipTelefone_LostFocus()
    If cboTipTelefone.Visible = True And Len(cboTipTelefone.Tag) > 0 Then
        '        If flxTelefones.TextMatrix(Left(cboTipTelefone.Tag, InStr(cboTipTelefone.Tag, ",")), Right(cboTipTelefone.Tag, InStr(cboTipTelefone.Tag, ",") - 1)) <> cboTipTelefone.Text Then
        '            flxTelefones.TextMatrix(Left(cboTipTelefone.Tag, InStr(cboTipTelefone.Tag, ",")), 3) = Empty
        '        End If
        flxTelefones.TextMatrix(Left(cboTipTelefone.Tag, InStr(cboTipTelefone.Tag, ",")), Right(cboTipTelefone.Tag, InStr(cboTipTelefone.Tag, ",") - 1)) = cboTipTelefone.Text
    End If
End Sub

Private Sub cboUF_Click()
'Observa��o: Lista do combo de unidades da federa��o carregadas em tempo de projeto.
'            Qualquer altera��o deve ser feita na propriedade LIST do controle.

    Dim lI As Integer    'Utilizado com For...Next.

    Call gpHabSalvar
    If optEndCobranca(1) = True Then    'Endere�o de cobran�a igual.
        For lI = 0 To (frmT46V105A.cboUF.ListCount - 1)
            If frmT46V105A.cboUF.List(lI) = cboUF.Text Then
                frmT46V105A.cboUF.ListIndex = lI
                Exit For
            End If
        Next lI
    End If
End Sub
Private Sub cboUF_GotFocus()
    Set mAntControle = ActiveControl
End Sub
Private Sub cboUF_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub cboUF_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then     'Teclou backspace.
        cboUF.ListIndex = -1
    End If
End Sub

Private Sub CmbTipCotac_Change()
    CmbTipCotac_Click
End Sub

Private Sub CmbTipCotac_Click()
    If CmbTipCotac.ListIndex = 2 Then
        frmT46V212A.ChkBloqueiaWeb.Enabled = True
        frmT46V212A.ChkBloqueiaWeb.ForeColor = 8388608
    Else
        frmT46V212A.ChkBloqueiaWeb.value = False
        'frmT46V212A.ChkBloqueiaWeb.Enabled = False  '20171219: Removido disabled (Juliano/Adriana)
        frmT46V212A.ChkBloqueiaWeb.ForeColor = vbGrayText
    End If
    If gCarga = True Then Exit Sub
    Call gpLimPremios
End Sub

Private Sub cmdAjustarDatas_Click()
    frmT46V121A.Show vbModal, Me
End Sub

Private Sub cmdAjustarDatas_GotFocus()
    Set mAntControle = ActiveControl
End Sub

Private Sub cmdCEP_Click()
    Dim lCEP As String    'CEP.
    Dim lTipLogradouro As String      'Tipo de logradouro.
    Dim lLogradouro As String   'Logradouro.
    Dim lBairro As String    'Bairro.
    Dim lCidade As String    'Cidade.
    Dim lUF As String    'UF.

    Call frmT46V104A.mpInstanciarForm(lCEP, lTipLogradouro, lLogradouro, lBairro, lCidade, lUF)
    txtCep.Text = lCEP
    txtTipLogradouro.Text = lTipLogradouro
    txtNomLogradouro.Text = lLogradouro
    txtBairro.Text = lBairro
    txtCidade.Text = lCidade
    If gM46V999.gfPreenchido(lUF) Then
        cboUF.Text = lUF
    End If
End Sub
Private Sub cmdCosseguro_Click()
    frmT46V116A.Show vbModal, Me
    '1100023 - Imprimir Informa��es de Cosseguro na Cota��o
    If Not mAntControle Is Nothing Then
        If mAntControle.Visible = True And mAntControle.Enabled = True Then
            mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
        End If
    End If
End Sub
Private Sub cmdEndCobranca_Click()
    Call mnuEndCobranca_Click
    On Error Resume Next
    mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
End Sub

Private Sub cmdExcluir_Click()
    Dim lRowSel As Integer

    cboTipTelefone.Visible = False
    Text1.Visible = False

    With flxTelefones
        If .Rows > 2 Then
            lRowSel = .RowSel
            If .RowSel <> 0 Then
                .RemoveItem .RowSel
            End If

            'Resize da coluna de observa��o
            .Col = 5
            .ColSel = .Col
            If .Rows > 5 Then
                .ColWidth(.Col) = 3220    'com scrollbar
            Else
                .ColWidth(.Col) = 3480    'sem scrollbar
            End If

        Else
            Call mpIniciaGridTelefones(1)
        End If
    End With
End Sub

Private Sub cmdForPagCartao_Click()
    frmT46V213A.Show vbModal, Me
End Sub

Private Sub cmdJuros_Click()
    Set mAntControle = cmdJuros
    frmT46V142A.Show vbModal, Me
    
    txtJuros.Text = Format(frmT46V142A.getJuros(IIf(optForPagamento(0).Value, e_FormaPagamento_Carne, e_FormaPagamento_Debito), optPagAntecipado(1).Value, val(txtQuaParcelas.Text)), "standard")
    
    If mAntControle.Visible = True And mAntControle.Enabled = True Then
        mAntControle.SetFocus
    End If
End Sub

Private Sub cmdNovo_Click()
    cboTipTelefone.Visible = False
    Text1.Visible = False

    With flxTelefones
        .Rows = .Rows + 1
        Call mpIniciaGridTelefones(.Rows - 1)
        If .Enabled And .Visible Then
            .SetFocus
        End If
        .RowHeight(.Row) = cboTipTelefone.Height
        .Row = .Rows - 1
        .RowSel = .Rows - 1
        .ColSel = 1
        .Col = 1
    End With
End Sub

Private Sub cmdForPagCarne_Click()
    frmT46V109A.Show vbModal, Me
    If mAntControle.Visible = True And mAntControle.Enabled = True Then
        mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
    End If
    optForPagamento(0).Value = True
End Sub

Private Sub cmdForPagCarne_GotFocus()
    Set mAntControle = ActiveControl
End Sub

Private Sub cmdForPagDebito_Click()
    Call frmT46V110A.mpInstanciarForm(2)
    If Not gM46V999.gDesabilitarForms Then
        If mAntControle.Visible = True And mAntControle.Enabled = True Then
            mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
        End If
    End If
End Sub

Private Sub cmdForPagDebito_GotFocus()
    Set mAntControle = ActiveControl
End Sub

Private Sub cmdForPagDebito1oBoleto_Click()
    frmT46V185A.Show vbModal, Me

    '    If Not gM46V999.gDesabilitarForms Then
    '        If mAntControle.Visible = True And mAntControle.Enabled = True Then
    '            mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
    '        End If
    '    End If
    optForPagamento(2).Value = True
    If optForPagamento(2).Visible Then
        optForPagamento(2).SetFocus
    End If
End Sub

Private Sub cmdPagAntBloqueto_Click()
    frmT46V108A.Show vbModal, Me
    If mAntControle.Visible = True And mAntControle.Enabled = True And Me.Tag <> "Desabilitado" Then
        mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
    End If
End Sub

Private Sub cmdPagAntBloqueto_GotFocus()
    Set mAntControle = ActiveControl
End Sub

Private Sub cmdPagAntDebito_Click()
    Call frmT46V110A.mpInstanciarForm(1)
    If mAntControle.Visible = True And mAntControle.Enabled = True Then
        mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
    End If
End Sub

Private Sub cmdPagAntDebito_GotFocus()
    Set mAntControle = ActiveControl
End Sub

Private Sub cmdProdutor_Click()

    If gM46V999.gDesabilitarForms Then Exit Sub

    Call gpHabSalvar

    gProChamado = "T46V115A"
    frmT46V115A.Show vbModal, Me
    If mAntControle.Visible = True And mAntControle.Enabled = True Then
        mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
    End If
End Sub

Private Sub cmdProlabore_Click()
    frmT46V180A.Show vbModal, Me
    If Not mAntControle Is Nothing Then
        If mAntControle.Visible = True And mAntControle.Enabled = True Then
            mAntControle.SetFocus
        End If
    End If
End Sub

Private Sub cmdSiscad_Click(index As Integer)
    Dim lCPF_CNPJ As String
    Dim lBlnContemDados As Boolean

    If gM46V999.gDesabilitarForms Then Exit Sub

    If optTipPessoa(0).Value = True Then    'F�sica
        Call gM46V999.gpForCPF(2, txtCPF.Text, lCPF_CNPJ)
    Else    'Jur�dica
        Call gM46V999.gpForCNPJ(2, txtCnpj.Text, lCPF_CNPJ)
    End If

    If InStr(gM46V999.gEndExeSiscad, "\") = 0 Then
        gM46V999.gEndExeSiscad = App.Path & "\" & gM46V999.gEndExeSiscad
    End If

    On Error GoTo ControleErro

    If index = 0 Then
        If Dir(gM46V999.gEndExeSiscad) <> "" Then
            Dim lCpfCnpjSiscad As String
            If IsNumeric(lCPF_CNPJ) Then
                lCpfCnpjSiscad = Format(lCPF_CNPJ, String(14, "0"))
            Else
                lCpfCnpjSiscad = Right$(String(14, "0") & Trim(lCPF_CNPJ), 14)
            End If
            Call Shell(gM46V999.gEndExeSiscad & " " & lCpfCnpjSiscad & "0" & Format(frmT46V101A.objstcPedido.objstcCtrlEmis.NUM_PI, String(10, "0")), vbNormalFocus)
        End If
        Call mpAtualizaDadosSegurado(lBlnContemDados, True)
    Else
        Call mpAtualizaDadosSegurado(lBlnContemDados, False)
    End If
    Exit Sub
ControleErro:
    If Err.Number = 53 Then
        Call gM46V999.gpGraLog(1, "FTE0025 - Arquivo n�o localizado: Sistema SISCAD (Verificar endere�o informado no arquivo de inicializa��o).")
    End If
End Sub

Private Sub cmdUniProdutiva_Click()

    If gM46V999.gDesabilitarForms Then Exit Sub

    Call gpHabSalvar

    gProChamado = "T46V114A"
    frmT46V114A.Show vbModal, Me
    If mAntControle.Visible = True And mAntControle.Enabled = True Then
        mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
    End If
End Sub
Private Sub cmdVoltar_Click()
'1. Torna form invis�vel.
    Dim lAchouCobert As Boolean
    Dim i As Integer
    
    'Verifica cobertura 1215 - Bicicleta especifica do corretor 908624 (Asteca)
    If frmT46V101A.objstcPedido.Cod_Produto >= 1094 And frmT46V101A.objstcPedido.Cod_Ramo = e_CodRamo_Residencial Then
        
        '8000046259, Problema na cobertura de Bicicleta
        If frmT46V101A.objstcPedido.Cod_Estip <> 30 Then 'Supremo pode comercializar.
        
            If frmT46V102A.txtCorretor(1).Tag <> "908624" Then '908624 - (Asteca)
                lAchouCobert = False
                For i = 1 To frmT46V117A.flxCobertura.Rows - 1
                    If frmT46V117A.flxCobertura.TextMatrix(i, 0) = 1215 Then
                        frmT46V117A.flxCobertura.RemoveItem (i)
                        lAchouCobert = True
                        Exit For
                    End If
                Next i
                
                If Not lAchouCobert Then
                    For i = 1 To frmT46V117A.flxSelecionadas.Rows - 1
                        If frmT46V117A.flxSelecionadas.TextMatrix(i, 0) = 1215 Then
                            frmT46V117A.flxSelecionadas.RemoveItem (i)
                            frmT46V101A.objstcPedLoc.ColecaoPedCobert.Remove (i)
                        End If
                    Next i
                End If
            Else
                For i = 1 To frmT46V117A.flxSelecionadas.Rows - 1
                    If frmT46V117A.flxSelecionadas.TextMatrix(i, 0) = 1215 Then
                        lAchouCobert = True
                        Exit For
                    End If
                Next i
            
                If Not lAchouCobert Then
                    Call frmT46V117A.mpSelecionarCoberturas(frmT46V101A.objstcPedLoc.ColecaoPedCobert, False, True, True)
                End If
            End If
        End If
    End If
    
    If frmT46V101A.lblTipDocumento.Caption = "2" Then   'Proposta.
        If optForPagamento(0).Enabled = True Then
            If optForPagamento(0) = True Then   'Carn�.
                With frmT46V109A
                    If gM46V999.gfDisponivel("frmT46V109A") = True Then
                        If .cboCarBanco.ListIndex = -1 Then
                            .Show vbModal, Me
                        End If
                    Else
                        .Show vbModal, Me
                    End If
                End With
            ElseIf optForPagamento(1) = True Then    'D�bito em conta.
                With frmT46V110A
                    If gM46V999.gfDisponivel("frmT46V110A") = True Then
                        If .cboDebBanco.ListIndex = -1 Then
                            .Show vbModal, Me
                        ElseIf Not gM46V999.gfPreenchido(.txtDebAgencia.Text) Then
                            .Show vbModal, Me
                        ElseIf Not gM46V999.gfPreenchido(.txtDebConta.Text) Then
                            .Show vbModal, Me
                        End If
                    Else
                        .Show vbModal, Me
                    End If
                End With
            ElseIf optForPagamento(2) = True Then    'D�bito autom�tico com 1� em boleto.
                With frmT46V185A
                    If gM46V999.gfDisponivel("frmT46V185A") = True Then
                        If .cboDebBanco.ListIndex = -1 Then
                            .Show vbModal, Me
                        ElseIf Not gM46V999.gfPreenchido(.txtDebAgencia.Text) Then
                            .Show vbModal, Me
                        ElseIf Not gM46V999.gfPreenchido(.txtDebConta.Text) Then
                            .Show vbModal, Me
                        End If
                    Else
                        .Show vbModal, Me
                    End If
                End With
            ElseIf CLng(txtQuaParcelas.Text) <> 0 Then   'Nada selecionado.
                Call gM46V999.gpGraLog(1, "FGE0085 - Forma de pagamento inv�lida (n�o selecionada).")
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If

    Me.Hide
    frmT46V101A.Show
End Sub

Private Sub flxTelefones_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call cboTipTelefone_LostFocus
    Call Text1_LostFocus

    cboTipTelefone.Visible = False
    Text1.Visible = False

    Select Case flxTelefones.Col

    Case 0
        If flxTelefones.CellPicture = picChecked Then
            Set flxTelefones.CellPicture = picUnchecked
        Else
            Set flxTelefones.CellPicture = picChecked
        End If

    Case 2, 3, 4, 5
        Text1.Move flxTelefones.Left + flxTelefones.CellLeft, _
                   flxTelefones.Top + flxTelefones.CellTop, _
                   flxTelefones.CellWidth, flxTelefones.CellHeight
        Text1.Tag = flxTelefones.Row & "," & flxTelefones.Col

        Select Case flxTelefones.ColSel    'Tamanho do campo conforme SQL (Tab_Ped_Telefones)
        Case 2: Text1.MaxLength = mMaxLength_DDD
        Case 3: Text1.MaxLength = mMaxLength_Telefone
        Case 4: Text1.MaxLength = mMaxLength_Ramal
        Case 5: Text1.MaxLength = mMaxLength_Obs
        End Select

        Text1.Text = flxTelefones.Text
        Text1.ZOrder
        Text1.Visible = True

    Case 1
        cboTipTelefone.Move flxTelefones.Left + flxTelefones.CellLeft, _
                            flxTelefones.Top + flxTelefones.CellTop, _
                            flxTelefones.CellWidth
        cboTipTelefone.Tag = flxTelefones.Row & "," & flxTelefones.Col
        If Len(Trim(flxTelefones.Text)) > 0 Then
            cboTipTelefone.Text = flxTelefones.Text
        End If
        cboTipTelefone.ZOrder
        cboTipTelefone.Visible = True

    Case 6
        Call cmdNovo_Click

    Case 7
        Call cmdExcluir_Click

    End Select

End Sub

Private Sub Form_Activate()
    Dim lI As Long
    Dim loptProlaboreAnt As Boolean

    loptProlaboreAnt = False
    '1. Posiciona controles.
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)

    If optTipPessoa(0).Value = True Then
        'Pessoa f�sica.
        panPesFisica.Visible = True
        panPesJuridica.Visible = False

        'Fichas 1200168 e 1200169 Inclus�o dos campos de listagem Renda e Faixa de Faturamento Pessoa f�sica e jur�dica
        '        'Ficha 900417
        '        If gM46V999.gCodProduto >= 968 Then
        '            lblProp(10).Caption = "Profiss�o: "
        '        Else
        '            lblProp(10).Caption = "Empr. onde trabalha: "
        '        End If

        'Ficha 1000307
        If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
            lblProp(28).Visible = True
            lblProp(29).Visible = True
            cboRegCons.Visible = True
            cboRegCons.Enabled = True
            txtnumregistro.Visible = True
            txtnumregistro.Enabled = True
        Else
            lblProp(28).Visible = False
            lblProp(29).Visible = False
            cboRegCons.Visible = False
            cboRegCons.Enabled = False
            txtnumregistro.Visible = False
            txtnumregistro.Enabled = False
        End If

    Else    'Pessoa jur�dica.
        panPesFisica.Visible = False
        panPesJuridica.Visible = True

        'Ficha 1000307
        lblProp(28).Visible = False
        lblProp(29).Visible = False
        cboRegCons.Visible = False
        cboRegCons.Enabled = False
        txtnumregistro.Visible = False
        txtnumregistro.Enabled = False
        cboRegCons.ListIndex = -1
        
        Dim blnVisualizaCondominioConstituido As Boolean
        blnVisualizaCondominioConstituido = False
        
        With frmT46V101A
            blnVisualizaCondominioConstituido = gM46V999.gCod_Ramo = e_CodRamo_Condominio And .objstcPedLoc.Tip_Emissao <> e_TipEmissao_Endosso And gM46V999.gCodProduto >= 1100
            If Not blnVisualizaCondominioConstituido And gM46V999.gCod_Ramo = e_CodRamo_Condominio And .objstcPedLoc.Tip_Emissao = e_TipEmissao_Endosso Then
                gM46V999.gNum_Apolice = .objstcPedido.Num_Apol_An
                If gM46V999.gTarifa_Apolice >= 1100 Then
                    blnVisualizaCondominioConstituido = True
                End If
            End If
        End With
        
        If blnVisualizaCondominioConstituido Then
            cboRamAtividade.Visible = False
            lblProp(1).Visible = False
        Else
            cboRamAtividade.Visible = True
            lblProp(1).Visible = True
        End If
    End If

    '2. Verifica programa chamado.
    If gM46V999.gfPreenchido(gProChamado) Then
        Select Case gProChamado
        Case "T46V102A"
        Case "T46V104A"     'Busca endere�o.
            If txtTipLogradouro.Visible = True And txtTipLogradouro.Enabled = True Then
                txtTipLogradouro.SetFocus
            End If
        Case "T46V112A"     'Corretores.
            'Nada a fazer.
        Case "T46V114A"     'Unidades produtivas.
            'Nada a fazer.
        Case "T46V115A"     'Produtores
            'Nada a fazer.
        Case "T46V125A"     'Grade de itens.
            'Nada a fazer.
            'Ficha 1000307
        Case "T46V111A"     'Lista de classifica��es
            If frmT46V102A.cboRegCons.ListIndex = -1 Then
                Call mpCarregarRegConselho
            End If
        Case Else
            Call gM46V999.gpGraLog(2, "FTE2011 - " & gM46V999.gPrefixo & "Programa chamado desconhecido (" & _
                                      gProChamado & " - janela T46V102A).")
        End Select
        gProChamado = ""
    End If

    lblX(0).Visible = False
    txtDatEmis.Visible = False
    Select Case gM46V999.gTipUsuario
    Case e_Tipo_Usu_Corretor, e_Tipo_Usu_CorretorYasuda
        cmdSiscad(0).Visible = False
        cmdSiscad(1).Visible = False
        txtJuros.Visible = False
        lblProp(16).Visible = False
        lblProp(15).Visible = False
        CmbTipCotac.Visible = False
        txtNumSiscota.Visible = False
        '1101158 - Numero de Referencia do Corretor
        lblX(27).Visible = True
        If txtReferencia = "" Then txtReferencia = "0"
        txtReferencia.Visible = True
        txtValorPgAnt.Visible = False
        lblPagAnt.Visible = False

        'Proposta Eletr�nica
        If gM46V999.gNovaTransmissaoMulti = True Then
            If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
                If frmT46V112A.optTipEmissao(1).Value = True Then  'Renova��o
                    If frmT46V101A.optTipPessoa(0).Value = True Then
                        Call gM46V999.gpForCPF(2, txtCnpj.Text, mAtuCNPJ)  'F�sica
                    Else
                        Call gM46V999.gpForCNPJ(2, txtCnpj.Text, mAtuCNPJ)    'Jur�dica
                    End If
                    If mAtuCNPJ <> mAntCNPJ Then
                        frmT46V101A.txtSegurado.Enabled = True
                        txtSegurado.Enabled = True
                    Else
                        frmT46V101A.txtSegurado.Text = mAntSeg
                        frmT46V101A.txtSegurado.Enabled = False
                        txtSegurado.Text = mAntSeg
                        txtSegurado.Enabled = False
                    End If
                End If
            End If
        End If

    Case Else
        cmdSiscad(0).Visible = True
        cmdSiscad(1).Visible = True
        If gM46V999.PerfilAcesso = e_Perfil_DIG_Gestor Then
            lblX(0).Visible = True
            txtDatEmis.Visible = True
        End If
    End Select

    Me.Caption = "Dados da Proposta - PI: " & frmT46V101A.objstcPedido.objstcCtrlEmis.NUM_PI
        
    '' 2016-08-24: Selecionar SISCOTA se for Endosso ou Renovacao Sompo
    'If frmT46V101A.objstcPedLoc.Tip_Emissao = e_TipEmissao_Endosso Or frmT46V101A.objstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
        'If CmbTipCotac.Enabled Then
        '    CmbTipCotac.ListIndex = 2
        'End If
    'End If
    
    Call mpHabControles

Finalizar:
    If gM46V999.gDesabilitarForms Then
        If Me.Tag <> "Desabilitado" Then
            DesabilitaControles Me
            '[FRH.07]->
            '[Controles que necessitam ser DESATIVADOS no 102A]
            cmdUniProdutiva.Enabled = False: cmdUniProdutiva.Picture = cmdExtra(2).Picture
            cmdProdutor.Enabled = False: cmdProdutor.Picture = cmdExtra(2).Picture
            cmdForPagCarne.Enabled = False: cmdForPagCarne.Picture = cmdExtra(2).Picture
            cmdSiscad(0).Enabled = False: cmdSiscad(0).Picture = cmdExtra(2).Picture
            cmdCEP.Enabled = False: cmdCEP.Picture = cmdExtra(2).Picture
            cmdSiscad(1).Enabled = False
            '<-[FRH]
            'Ficha 900654
            If frmT46V101A.objstcPedido.mTransmissao = True Then
                If frmT46V101A.objstcPedido.Cod_Bco_Cobr = 999 And _
                   frmT46V101A.objstcPedido.Cod_Agen_Bas_Cor = 9999 And _
                   frmT46V101A.objstcPedido.Num_Cta_Bas_Cor = 9999999 And _
                   frmT46V101A.objstcPedido.Cod_Tip_Doc = e_TipDoc_Proposta Then
                    cmdAjustarDatas.Enabled = False
                    cmdAjustarDatas.Picture = cmdExtra(2).Picture
                    cmdJuros.Enabled = True: cmdJuros.Picture = cmdExtra(1).Picture
                    cmdForPagCartao.Enabled = True: cmdForPagCartao.Picture = cmdExtra(1).Picture
                Else
                    cmdAjustarDatas.Enabled = True
                    cmdAjustarDatas.Picture = cmdExtra(1).Picture
                End If
            End If
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
            '[FRH.07]->
            '[REATIVANDO controles no 102A]
            cmdUniProdutiva.Enabled = True: cmdUniProdutiva.Picture = cmdExtra(1).Picture
            cmdProdutor.Enabled = True: cmdProdutor.Picture = cmdExtra(1).Picture
            
            If frmT46V101A.objstcPedido.Cod_Bco_Cobr = 999 And _
                   frmT46V101A.objstcPedido.Cod_Agen_Bas_Cor = 9999 And _
                   frmT46V101A.objstcPedido.Num_Cta_Bas_Cor = 9999999 And _
                   frmT46V101A.objstcPedido.Cod_Tip_Doc = e_TipDoc_Proposta Then
                    cmdForPagCarne.Enabled = False: cmdForPagCarne.Picture = cmdExtra(2).Picture
                    txtValorPgAnt.Enabled = False
                    cmdJuros.Enabled = True: cmdJuros.Picture = cmdExtra(1).Picture
                    cmdForPagCartao.Enabled = True: cmdForPagCartao.Picture = cmdExtra(1).Picture
            Else
                cmdForPagCarne.Enabled = True: cmdForPagCarne.Picture = cmdExtra(1).Picture
                cmdSiscad(0).Enabled = True: cmdSiscad(0).Picture = cmdExtra(1).Picture
                cmdCEP.Enabled = True: cmdCEP.Picture = cmdExtra(1).Picture
                cmdSiscad(1).Enabled = True
            End If
            '<-[FRH]
        End If
    End If

    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        If lblNomUnidade.Caption = "MATRIZ" And gM46V999.CDblx(txtCorretor(1).Tag) <> 0 Then
            txtCorretor_Change (1)
            txtCorretor_LostFocus (1)
        End If
    End If

    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        If gM46V999.gDesabFormsTransmi = True Then
            
            loptProlaboreAnt = frmT46V102A.optProlabore(0).Value
            gCarga = True
            If frmT46V101A.objstcPedido.mCotacaoInterna = False Then
                DesabilitaControles Me
            End If
            
            frmT46V102A.optProlabore(0).Value = loptProlaboreAnt
            gCarga = False
            cmdVoltar.Enabled = True
            cmdSiscad(0).Enabled = True
            cmdSiscad(1).Enabled = True
            optTipPessoa(0).Enabled = True
            optTipPessoa(1).Enabled = True
            'corre��o - thais
            If frmT46V101A.objstcPedido.mTransmissao = True Then
                If frmT46V101A.objstcPedido.Cod_Bco_Cobr = 999 And _
                   frmT46V101A.objstcPedido.Cod_Agen_Bas_Cor = 9999 And _
                   frmT46V101A.objstcPedido.Num_Cta_Bas_Cor = 9999999 And _
                   frmT46V101A.objstcPedido.Cod_Tip_Doc = e_TipDoc_Proposta Then
                    cmdAjustarDatas.Enabled = False
                    cmdAjustarDatas.Picture = cmdExtra(2).Picture
                    txtValorPgAnt.Enabled = False
                    cmdForPagCartao.Enabled = True: cmdForPagCartao.Picture = cmdExtra(1).Picture
                Else
                    'Ficha 1000094
                    txtValorPgAnt.Enabled = True
                    cmdAjustarDatas.Enabled = True
                    cmdAjustarDatas.Picture = cmdExtra(1).Picture
                End If
                
                'AJuste para habilitar pro-labore somente para imobiliario
                panProlabore(0).Enabled = (frmT46V101A.objstcPedido.Cod_Estip = 20)
                optProlabore(0).Enabled = (frmT46V101A.objstcPedido.Cod_Estip = 20)
                optProlabore(1).Enabled = (frmT46V101A.objstcPedido.Cod_Estip = 20)
                cmdProlabore.Enabled = (frmT46V101A.objstcPedido.Cod_Estip = 20)
            End If
        End If
    End If
    cmdVoltar.Enabled = True

End Sub

Public Sub mpAtualizaDadosSegurado(ByRef pContemDados As Boolean, pConfirmar As Boolean)

    Dim lCPF_CNPJ As String
    Dim BDSisCad As ADODB.Connection
    Dim rsSisCad As clsYasRecordSet
    Dim rsSyasCob As clsYasRecordSet
    Dim lMensagem As String
    Dim lLibAltEnd As Boolean
    Dim lLibAltNome As Boolean
    Dim lSelect As String

    pContemDados = False

    If pConfirmar = True Then    'SisCadAberto 1
        ' If MsgBox("Deseja atualizar os dados do segurado?", vbYesNo + vbQuestion, Me.Caption) = vbNo Then Exit Sub
        MsgBox "Pressione Ok, ap�s terminar de atualizar os dados do SisCad!", vbOKOnly, Me.Caption
    End If
    Screen.MousePointer = vbHourglass
    If optTipPessoa(0).Value = True Then    'F�sica
        Call gM46V999.gpForCPF(2, txtCPF.Text, lCPF_CNPJ)
    Else    'Jur�dica
        Call gM46V999.gpForCNPJ(2, txtCnpj.Text, lCPF_CNPJ)
    End If
    If gM46V999.gfPreenchido(lCPF_CNPJ) And Trim(lCPF_CNPJ) <> "0" Then  'CNPJ_CPF.
        If gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, BDSisCad, lMensagem) = False Then
            Exit Sub
        End If

        lSelect = " select * from p0044800.dbo.t0044800 Where "
        lSelect = lSelect & " num_cgc_cpf = '" & Trim(lCPF_CNPJ) & "'"
        If gM46V999.gfObtRegistro(BDSisCad, lSelect, rsSisCad, lMensagem) = False Then
            Call gM46V999.gpFecha2(BDSisCad)
            Call gM46V999.gpGraLog(1, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If Not rsSisCad.EOF Then
            If IsNull(rsSisCad("TIP_IRREGULAR")) Then
                lLibAltEnd = False
            Else
                lLibAltEnd = (rsSisCad("TIP_IRREGULAR") = 9)
            End If
            If IsNull(rsSisCad("ALTERA_NOME")) Then
                lLibAltNome = False
            Else
                lLibAltNome = (rsSisCad("ALTERA_NOME") = 9)
            End If
            If lLibAltNome = True And pConfirmar = True Then
                If MsgBox("Nome do segurado liberado para altera��o!" & vbCrLf & "Atualizar tela com os dados do SISCAD?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                    lLibAltNome = False
                End If
            End If
            If lLibAltEnd = True And pConfirmar = True Then
                If MsgBox("Endere�o do segurado e cobran�a liberados para altera��o!" & vbCrLf & "Atualizar tela com os dados do SISCAD?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                    lLibAltEnd = False
                End If
            End If
            If lLibAltNome = False Then

                'Ficha 900417
                If gM46V999.gCodProduto >= 968 Then
                    txtSegurado.Text = Trim(rsSisCad("NOME_RAZAO"))
                Else
                    txtSegurado.Text = Trim(rsSisCad("NOME_RAZAO")) & "   " & rsSisCad("NOM_LOCTRAB")
                End If

                If IsNull(rsSisCad("ALTERA_NOME")) Then
                    txtSegurado.Enabled = False
                ElseIf rsSisCad("ALTERA_NOME") <> 9 Then
                    txtSegurado.Enabled = False
                Else
                    txtSegurado.Enabled = True
                End If
            Else
                txtSegurado.Enabled = True
            End If
            If lLibAltEnd = False Then
                txtTipLogradouro.Text = rsSisCad("COD_TIP_LOGR")
                txtNomLogradouro.Text = rsSisCad("NOM_LOGR")
                
                If Len(Trim(txtNumero.Text & txtComplemento.Text)) = 0 Then
                    txtNumero.Text = gM46V999.gfTratarNulo(rsSisCad("NUM_LOGR"), tpoNumerico)
                    txtComplemento.Text = Left(Trim(rsSisCad("DSC_COMPL_LOGR")), 13)
                End If
                
                txtBairro.Text = rsSisCad("NOM_BAIRRO")
                txtCidade.Text = rsSisCad("NOM_CID")
                If Trim(rsSisCad("SIG_UF")) <> "" Then
                    cboUF.Text = rsSisCad("SIG_UF")
                End If
                txtCep.Text = Format(rsSisCad("Num_CEP"), "00000-000")
                If rsSisCad("TIP_IRREGULAR") <> 9 Then
                    txtCep.Enabled = False
                    txtTipLogradouro.Enabled = False
                    txtNomLogradouro.Enabled = False
                    txtNumero.Enabled = False
                    txtComplemento.Enabled = False
                    txtBairro.Enabled = False
                    txtCidade.Enabled = False
                    cboUF.Enabled = False
                    txtCep.Enabled = False
                Else
                    txtCep.Enabled = True
                    txtTipLogradouro.Enabled = True
                    txtNomLogradouro.Enabled = True
                    txtNumero.Enabled = True
                    txtComplemento.Enabled = True
                    txtBairro.Enabled = True
                    txtCidade.Enabled = True
                    cboUF.Enabled = True
                    txtCep.Enabled = True
                End If

            Else
                txtCep.Enabled = True
                txtTipLogradouro.Enabled = True
                txtNomLogradouro.Enabled = True
                txtNumero.Enabled = True
                txtComplemento.Enabled = True
                txtBairro.Enabled = True
                txtCidade.Enabled = True
                cboUF.Enabled = True
                txtCep.Enabled = True
            End If
            '1001046 - Telefone do Segurado 8 Digitos
            txtTelefones.Text = gM46V999.gfTratarNulo(rsSisCad("NUM_FONE"), tpoString)

            '[FRH.08]->
            txtDigCodSegurado.Enabled = False

            'Ficha 1000495 - Plano Belacity
            '''            'Ficha 901282
            '''            'Ficha 901562
            '''            If gM46V111.gCod_Plano <> 2050 And gM46V111.gCod_Plano <> 2101 And gM46V111.gCod_Plano <> 2102 And (gM46V111.gCod_Plano < 2060 Or gM46V111.gCod_Plano > 2070) Then
            '''                If gM46V999.gTipUsuario <> e_Tipo_Usu_Yasuda Then
            '''                  txtCodSegurado.Text = rsSisCad("COD_IDENTIFICADOR")
            '''                  txtCodSegurado.Enabled = False
            '''                Else
            '''                  txtCodSegurado.Enabled = True
            '''                End If
            '''            '<-[FRH]
            '''            End If

            If (rsSisCad("SEXO") = "M") Then
                optSexo(0).Value = True
            Else
                optSexo(1).Value = True
            End If
            cboIdeTipDocumento.ListIndex = gM46V999.gfRetornarListIndex(cboIdeTipDocumento, val(Trim(rsSisCad("TIP_DOCUM"))))
                txtIdeNumDocumento.Text = rsSisCad("NUM_DOCUM")
            txtIdeOrgExpedidor.Text = rsSisCad("NOM_ORG_EXPED")
            txtIdeDatEmissao.Text = gM46V999.gfForData(2, rsSisCad("DAT_EXPED"))
            txtDataNascimento.Text = gM46V999.gfForData(2, rsSisCad("DAT_NASC"))

            'Ficha 900417
            'A partir da vers�o 968 este campo ir� conter a PROFISS�O e n�o mais o local de trabalho
            If gM46V999.gCodProduto >= 968 Then
                txtIdeEmpresa.Text = rsSisCad("NOM_PROFISSAO")
            Else
                txtIdeEmpresa.Text = rsSisCad("NOM_LOCTRAB")
            End If
            txtTelefones.Enabled = False

            optSexo(0).Enabled = False
            optSexo(1).Enabled = False
            cboIdeTipDocumento.Enabled = False
            txtIdeNumDocumento.Enabled = False
            txtIdeOrgExpedidor.Enabled = False
            txtIdeDatEmissao.Enabled = False
            txtDataNascimento.Enabled = False
            txtIdeEmpresa.Enabled = False

            Select Case Trim(rsSisCad("RAMO_ATIV_EMPR"))
            Case "A"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 1)
            Case "B"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 2)
            Case "C"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 3)
            Case "D"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 4)
            Case "E"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 5)
            Case "F"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 6)
            Case "G"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 7)
            Case "H"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 8)
            Case "I"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 9)
            Case "J"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 10)
            Case "K"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 11)
            Case "L"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 12)
            Case "M"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 13)
            Case "N"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 14)
            Case "O"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 15)
            Case "P"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 16)
            Case "Q"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 17)
                'Ficha 1000365 - Inclus�o Ramos de Atividade
            Case "R"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 18)
            Case "S"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 19)
            Case "T"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 20)
            Case "U"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 21)
            Case "V"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 22)
            Case "W"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 23)
            Case "X"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 24)
            Case "Y"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 25)
            Case "Z"
                cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 26)
            End Select
            cboRamAtividade.Enabled = False

            If optTipPessoa(0).Value = True Then    'F�sica
                '                txtRendaMensal.Text = Format(gM46V999.gfTratarNulo(rsSisCad("RendaFatu"), TpoValor), "standard")
                '                txtRendaMensal.Enabled = False
                'Fichas 1200168 e 1200169 Inclus�o dos campos de listagem Renda e Faixa de Faturamento Pessoa f�sica e jur�dica
                cboRendaMensal.ListIndex = gM46V999.gfRetornarListIndex(cboRendaMensal, val(Trim(gM46V999.gfTratarNulo(rsSisCad("RendaFatu"), tpoNumerico))))
                cboRendaMensal.Enabled = False
                'Ficha 1000307
                If gM46V999.gCod_Ramo = e_CodRamo_Empresarial And gObrRegistro = True Then
                    lblProp(28).Visible = True
                    lblProp(29).Visible = True
                    cboRegCons.Visible = True
                    cboRegCons.Enabled = True
                    txtnumregistro.Visible = True
                    txtnumregistro.Enabled = True
                    '                    Call mpCarregarRegConselho
                    Call cboRegCons_Change
                Else
                    lblProp(28).Visible = False
                    lblProp(29).Visible = False
                    cboRegCons.Visible = False
                    cboRegCons.Enabled = False
                    txtnumregistro.Visible = False
                    txtnumregistro.Enabled = False
                    txtnumregistro.Text = " "
                End If
            Else
                '                txtFaturamento.Text = Format(gM46V999.gfTratarNulo(rsSisCad("RendaFatu"), TpoValor), "standard")
                '                txtFaturamento.Enabled = False
                'Fichas 1200168 e 1200169 Inclus�o dos campos de listagem Renda e Faixa de Faturamento Pessoa f�sica e jur�dica
                cboFaturamento.ListIndex = gM46V999.gfRetornarListIndex(cboFaturamento, val(Trim(rsSisCad("RendaFatu"))))
                cboFaturamento.Enabled = False
                'Ficha 1000307
                lblProp(28).Visible = False
                lblProp(29).Visible = False
                cboRegCons.Visible = False
                cboRegCons.Enabled = False
                txtnumregistro.Visible = False
                txtnumregistro.Enabled = False
                txtnumregistro.Text = " "
            End If

            lSelect = " select * from p0044800.dbo.T0044807 Where "
            lSelect = lSelect & " num_pi = " & frmT46V101A.objstcPedido.objstcCtrlEmis.NUM_PI
            lSelect = lSelect & " and num_cgc_cpf = '" & Trim(lCPF_CNPJ) & "'"
            If gM46V999.gfObtRegistro(BDSisCad, lSelect, rsSyasCob, lMensagem) = False Then
                Call gM46V999.gpFecha2(BDSisCad)
                Call gM46V999.gpGraLog(1, lMensagem)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            With frmT46V105A
                If lLibAltEnd = False Then
                    .txtCep.Enabled = False
                    .txtTipLogradouro.Enabled = False
                    .txtNomLogradouro.Enabled = False
                    .txtNumero.Enabled = False
                    .txtComplemento.Enabled = False
                    .txtBairro.Enabled = False
                    .txtCidade.Enabled = False
                    .cboUF.Enabled = False
                    .txtCep.Enabled = False
                    .txtTelefones.Enabled = False
                Else
                    .txtCep.Enabled = True
                    .txtTipLogradouro.Enabled = True
                    .txtNomLogradouro.Enabled = True
                    .txtNumero.Enabled = True
                    .txtComplemento.Enabled = True
                    .txtBairro.Enabled = True
                    .txtCidade.Enabled = True
                    .cboUF.Enabled = True
                    .txtCep.Enabled = True
                    .txtTelefones.Enabled = True
                End If
                If rsSyasCob.EOF Then
                    .txtNomLogradouro.Text = ""
                Else
                    .txtTipLogradouro.Text = rsSyasCob("COD_TIP_LOGR")
                    .txtNomLogradouro.Text = rsSyasCob("NOM_LOGR")
                    .txtNumero.Text = gM46V999.gfTratarNulo(rsSyasCob("NUM_LOGR"), tpoNumerico)
                    .txtComplemento.Text = rsSyasCob("DSC_COMPL_LOGR")
                    .txtBairro.Text = rsSyasCob("NOM_BAIRRO")
                    .txtCidade.Text = rsSyasCob("NOM_CID")
                    .txtCep.Text = Format(rsSyasCob("Num_CEP"), "00000-000")
                    If Trim(rsSyasCob("SIG_UF")) <> "" Then
                        .cboUF.Text = rsSyasCob("SIG_UF")
                    End If
                    If Not IsNull(rsSyasCob("num_fone")) Then
                        '1001046 - Telefone do Segurado 8 Digitos
                        .txtTelefones.Text = rsSyasCob("num_fone")
                    Else
                        .txtTelefones.Text = 0
                    End If
                    gCarga = True
                    optEndCobranca(0).Value = True
                    gCarga = False
                    pContemDados = True
                End If
            End With

            'OBTEM A LISTA DE TELEFONES DO SISCAD
            lSelect = " select * from p0044800.dbo.TAB_CONTATOS Where "
            lSelect = lSelect & " num_pi = " & frmT46V101A.objstcPedido.objstcCtrlEmis.NUM_PI
            lSelect = lSelect & " and num_cgc_cpf = '" & Trim(lCPF_CNPJ) & "'"
            If gM46V999.gfObtRegistro(BDSisCad, lSelect, rsSyasCob, lMensagem) = False Then
                Call gM46V999.gpFecha2(BDSisCad)
                Call gM46V999.gpGraLog(1, lMensagem)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If rsSyasCob.RecordCount > 0 Then
                Call LimpaGrid
                
                Dim lTelefone As String
                Dim lCarregaTelefone As Boolean
                Dim i As Integer
                
                Do Until rsSyasCob.EOF

                    If InStr(1, rsSyasCob!Num_Fone, "-") = 0 And Len(rsSyasCob!Num_Fone) < 7 Then
                        lCarregaTelefone = False
                    Else
                        lCarregaTelefone = True
                        lTelefone = Replace(rsSyasCob!Num_Fone, "-", Empty)
                    End If
                    
                    If lCarregaTelefone Then
                        i = 1
                        Do While i < flxTelefones.Rows
                            If Replace(flxTelefones.TextMatrix(i, 3), "-", Empty) = lTelefone Then
                                lCarregaTelefone = False
                                Exit Do
                            End If
                            i = i + 1
                        Loop
                    End If
                    
                    If lCarregaTelefone Then
                        'Nova linha na grid de telefone
                        If flxTelefones.TextMatrix(flxTelefones.Row, 3) <> Empty Then
                            Call cmdNovo_Click
                        End If
                        
                        flxTelefones.TextMatrix(flxTelefones.Row, 1) = cboTipTelefone.List(rsSyasCob!Tipo_Fone - 1)
                        flxTelefones.TextMatrix(flxTelefones.Row, 2) = rsSyasCob!Num_DDD_Fone
                        'tratamento para nono digito
                        If InStr(1, rsSyasCob!Num_Fone, "-") = 0 And Len(rsSyasCob!Num_Fone) < 9 Then
                            flxTelefones.TextMatrix(flxTelefones.Row, 3) = Mid(rsSyasCob!Num_Fone, 1, 4) & "-" & Mid(rsSyasCob!Num_Fone, 5, 4)
                        Else
                            flxTelefones.TextMatrix(flxTelefones.Row, 3) = Mid(rsSyasCob!Num_Fone, 1, 5) & "-" & Mid(rsSyasCob!Num_Fone, 6, 4)
                        End If
                        
                        flxTelefones.TextMatrix(flxTelefones.Row, 4) = IIf(IsNull(rsSyasCob!Num_Ramal) Or rsSyasCob!Num_Ramal = 0, "", rsSyasCob!Num_Ramal)
                        flxTelefones.TextMatrix(flxTelefones.Row, 5) = IIf(IsNull(rsSyasCob!DSC_OBS), "", rsSyasCob!DSC_OBS)
    
                        If rsSyasCob!COD_FONE_PRINCIPAL = 1 Then
                            flxTelefones.Col = 0
                            Set flxTelefones.CellPicture = picChecked
                        Else
                            flxTelefones.Col = 0
                            Set flxTelefones.CellPicture = picUnchecked
                        End If
                    End If

                    rsSyasCob.MoveNext
                Loop

            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim lI As Long

    '1. Posiciona form.
    Call gM46V999.gpCenForm(Me)

    '2. Posiciona controle.
    mAltEndereco = False

    '3. Carrega combo de ramos de atividade para pessoa jur�dica.
    With cboRamAtividade
        .AddItem "Agricultura, pecu�ria, produ��o florestal, pesca e aquicultura"
        .ItemData(.NewIndex) = 1
        .AddItem "Ind�stria extrativa"
        .ItemData(.NewIndex) = 2
        .AddItem "Ind�stria de transforma��o"
        .ItemData(.NewIndex) = 3
        .AddItem "Eletrecidade e g�s"
        .ItemData(.NewIndex) = 4
        .AddItem "�gua, esgoto, atividades de gest�o de res�duos e descontamina��o"
        .ItemData(.NewIndex) = 5
        .AddItem "Constru��o"
        .ItemData(.NewIndex) = 6
        .AddItem "Com�rcio; Repara��o de ve�culos automotores e motocicletas"
        .ItemData(.NewIndex) = 7
        .AddItem "Transporte, armazenagem ou comunica��es"
        .ItemData(.NewIndex) = 8
        .AddItem "Alojamento e alimenta��o"
        .ItemData(.NewIndex) = 9
        .AddItem "Informa��o e comunica��o"
        .ItemData(.NewIndex) = 10
        .AddItem "Atividades financeiras, de seguros e servi�os relacionados"
        .ItemData(.NewIndex) = 11
        .AddItem "Atividades imobili�rias"
        .ItemData(.NewIndex) = 12
        .AddItem "Atividades profissionais, cient�ficas e t�cnicas"
        .ItemData(.NewIndex) = 13
        .AddItem "Atividades administrativas e servi�os complementares"
        .ItemData(.NewIndex) = 14
        .AddItem "Administra��o p�blica, defesa e seguridade social"
        .ItemData(.NewIndex) = 15
        .AddItem "Educa��o"
        .ItemData(.NewIndex) = 16
        .AddItem "Sa�de humana e servi�os sociais"
        .ItemData(.NewIndex) = 17
        .AddItem "Artes, cultura, esporte e recrea��o"
        .ItemData(.NewIndex) = 18
        .AddItem "Outras atividades de servi�os"
        .ItemData(.NewIndex) = 19
        .AddItem "Servi�os dom�sticos"
        .ItemData(.NewIndex) = 20
        .AddItem "Organismos internacionais e outras institui��es extraterritoriais"
        .ItemData(.NewIndex) = 21
    End With

    '4. Carrega combo de tipos de documento de identifica��o para pessoa jur�dica.
    '[FRH] - F.900530
    
    With cboIdeTipDocumento
        .AddItem "RG / Carteira de identidade"
        .ItemData(.NewIndex) = 1
        
'Adicionado por Luiz Alberto em 09/12/2015
''        .AddItem "CNH / Carteira Nacional de Habilita��o"
''        .ItemData(.NewIndex) = 2
''        .AddItem "Carteira de trabalho"
''        .ItemData(.NewIndex) = 3
''
''        If gM46V999.gCodProduto < 1086 Then
''            .AddItem "Certid�o de casamento"
''            .ItemData(.NewIndex) = 4
''            .AddItem "Certid�o de nascimento"
''            .ItemData(.NewIndex) = 5
''            .AddItem "Certificado de reservista"
''            .ItemData(.NewIndex) = 6
''        End If
''
''        '1000753 - Inclusao - Tipo de Documento - Entidade de Classe
''        .AddItem "Entidades de Classe"
''        .ItemData(.NewIndex) = 7
        .AddItem "RNE / Identidade estrangeira"
        .ItemData(.NewIndex) = 8
''        .AddItem "Passaporte"
''        .ItemData(.NewIndex) = 9
''        .AddItem "PIS"
''        .ItemData(.NewIndex) = 10
''
''        If gM46V999.gCodProduto < 1086 Then
''            .AddItem "T�tulo de eleitor"
''            .ItemData(.NewIndex) = 12
''        End If
    End With


    '5. Carrega estado civil
    With cboEstadoCivil
        .AddItem "N�o Informado"
        .ItemData(.NewIndex) = 0
        .AddItem "Solteiro"
        .ItemData(.NewIndex) = 1
        .AddItem "Casado/Uni�o est�vel"
        .ItemData(.NewIndex) = 2
        .AddItem "Separado/Desquitado/Divorciado"
        .ItemData(.NewIndex) = 3
        .AddItem "Vi�vo"
        .ItemData(.NewIndex) = 4
    End With

    With CmbTipCotac
        .AddItem "Sem Cota��o"
        .ItemData(.NewIndex) = 0
        .AddItem "Proposta Avulsa"
        .ItemData(.NewIndex) = 1
        .AddItem "Siscota"
        .ItemData(.NewIndex) = 2
        .ListIndex = 0
    End With
    
    '' 2016-08-24: Selecionar SISCOTA se for Endosso ou Renovacao Sompo
    'If frmT46V101A.objstcPedLoc.Tip_Emissao = e_TipEmissao_Endosso Or frmT46V101A.objstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
    '   CmbTipCotac.ListIndex = 2
    'End If
    
    'Fichas 1200168 e 1200169 Inclus�o dos campos de listagem Renda e Faixa de Faturamento Pessoa f�sica e jur�dica
    With cboFaturamento
        .AddItem "N�o desejo informar": .ItemData(.NewIndex) = 0
        .AddItem "At� R$ 500 mil": .ItemData(.NewIndex) = 1
        .AddItem "De R$ 500 mil a 1 milh�o": .ItemData(.NewIndex) = 2
        .AddItem "De R$ 1 milh�o a 3 milh�es": .ItemData(.NewIndex) = 3
        .AddItem "De R$ 3 milh�es a 6 milh�es": .ItemData(.NewIndex) = 4
        .AddItem "De 6 milh�es a 15 milh�es": .ItemData(.NewIndex) = 5
        .AddItem "De 15 milh�es a 40 milh�es": .ItemData(.NewIndex) = 6
        .AddItem "De R$ 40 milh�es a 80 milh�es": .ItemData(.NewIndex) = 7
        .AddItem "De R$ 80 milh�es a 150 milh�es": .ItemData(.NewIndex) = 8
        .AddItem "Acima de 150 milh�es": .ItemData(.NewIndex) = 9
        .ListIndex = -1
    End With

    'Fichas 1200168 e 1200169 Inclus�o dos campos de listagem Renda e Faixa de Faturamento Pessoa f�sica e jur�dica
    With cboRendaMensal
        .AddItem "N�o desejo informar": .ItemData(.NewIndex) = 0
        .AddItem "Menor ou igual a � sal�rio m�nimo": .ItemData(.NewIndex) = 1
        .AddItem "Maior que � sal�rio m�nimo at�1 sal�rio m�nimo": .ItemData(.NewIndex) = 2
        .AddItem "Maior que 1 sal�rio m�nimo at�2 sal�rios m�nimos": .ItemData(.NewIndex) = 3
        .AddItem "Maior que 2 sal�rios m�nimos at� 3 sal�rios m�nimos": .ItemData(.NewIndex) = 4
        .AddItem "Maior que 3 sal�rios m�nimos at� 5 sal�rios m�nimos": .ItemData(.NewIndex) = 5
        .AddItem "Maior que 5 sal�rios m�nimos at� 10 sal�rios m�nimos": .ItemData(.NewIndex) = 6
        .AddItem "Maior que 10 sal�rios m�nimos at� 20 sal�rios m�nimos": .ItemData(.NewIndex) = 7
        .AddItem "Maior que 20 sal�rios m�nimos": .ItemData(.NewIndex) = 8
    End With

    '[FRH]-> Se usu�rio corretor..
    '1001046 - Telefone do Segurado 8 Digitos
    '    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
    txtTelefones.MaxLength = 9    '..permite armazenar Fone de 8 dig.     '1200178 - Telefone 9 Digitos
    '    Else
    '     txtTelefones.MaxLength = 7
    '    End If
    '<-[FRH]

    'Tarifa 1040 e 1041 - Desabilitar Cosseguro para Corretor
    lblProp(12).Visible = (gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda)
    panProp(5).Visible = (gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda)
    cmdCosseguro.Visible = (gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda)
    
    

    With cboTipTelefone
        .AddItem "Residencial"
        .ItemData(.NewIndex) = 1
        .AddItem "Celular"
        .ItemData(.NewIndex) = 2
        .AddItem "Comercial"
        .ItemData(.NewIndex) = 3
        .AddItem "FAX"
        .ItemData(.NewIndex) = 4
    End With

    With flxTelefones
        .FixedRows = 1
        .FixedCols = 0
        .Rows = 2
        .Cols = 9
        .TextMatrix(0, 0) = "Principal"
        .TextMatrix(0, 1) = "Tipo"
        .TextMatrix(0, 2) = "DDD"
        .TextMatrix(0, 3) = "Telefone"
        .TextMatrix(0, 4) = "Ramal"
        .TextMatrix(0, 5) = "Observa��o"
        .TextMatrix(0, 6) = Empty
        .TextMatrix(0, 7) = Empty
        .TextMatrix(0, 8) = Empty
        .ColWidth(8) = 0
        .RowHeight(1) = cboTipTelefone.Height
        Call mpIniciaGridTelefones(1)
    End With
    
    If frmT46V101A.objstcPedido.objstcCtrlEmis.Ind_Bloq_Web = 1 Then
        frmT46V212A.ChkBloqueiaWeb.value = 1
    End If
End Sub

Private Sub lblNomUnidade_Change()
    If gCarga = False Then
        gAltComum = True
    End If
End Sub

Private Sub mnuEndCobranca_Click()
    frmT46V105A.Show vbModal, Me
End Sub

Private Sub mnuEndSegurado_Click()
    Dim lI As Byte    'Utilizado com For...Next.

'Monta form.
    Load frmT46V104A
    With frmT46V104A
        'a) CEP.
        If gM46V999.gfPreenchido(txtCep.Text) Then
            .chkCEP.Value = 1
            .txtCep.Text = txtCep.Text
        End If
        'b) UF.
        .cboUF.Clear
        For lI = 0 To (cboUF.ListCount - 1)
            .cboUF.AddItem cboUF.List(lI)
        Next lI
        If cboUF.ListIndex <> -1 Then
            .chkUF.Value = 1
            .cboUF.ListIndex = cboUF.ListIndex
        End If
        'c) Logradouro.
        If gM46V999.gfPreenchido(txtNomLogradouro.Text) Then
            .chkLogradouro.Value = 1
            .txtNomLogradouro.Text = txtNomLogradouro.Text
        End If
        'd) Bairro.
        If gM46V999.gfPreenchido(txtBairro.Text) Then
            .chkBairro.Value = 1
            .txtBairro.Text = txtBairro.Text
        End If
        'e) Cidade.
        If gM46V999.gfPreenchido(txtCidade.Text) Then
            .chkCidade.Value = 1
            .txtCidade.Text = txtCidade.Text
        End If
    End With
    gProChamado = "T46V104A"
    frmT46V104A.Show vbModal, Me
End Sub
Private Sub mnuProdutor_Click()
    Call gpHabSalvar

    gProChamado = "T46V115A"
    frmT46V115A.Show vbModal, Me
End Sub
Private Sub mnuUniProdutiva_Click()
    Call gpHabSalvar

    gProChamado = "T46V114A"
    frmT46V114A.Show vbModal, Me
End Sub

Private Sub optCosseguro_Click(index As Integer)
    Call gpHabSalvar
    Select Case index
    Case 0
        cmdCosseguro.Enabled = True
        cmdCosseguro.Picture = cmdExtra(1).Picture
        If gCarga = False Then
            frmT46V116A.Show vbModal, Me
        End If
    Case 1
        cmdCosseguro.Enabled = False
        cmdCosseguro.Picture = cmdExtra(2).Picture
    End Select
    If gCarga = False Then
        gAltComum = True
    End If
End Sub
Private Sub optCosseguro_GotFocus(index As Integer)
    Set mAntControle = ActiveControl
End Sub
Private Sub optEndCobranca_Click(index As Integer)
    Dim lI As Byte    'Utilizado com For...Next.

    Select Case index
    Case 0      'Sim.
        cmdEndCobranca.Enabled = True
        cmdEndCobranca.Picture = cmdExtra(1).Picture
    Case 1      'N�o tem endere�o de cobran�a diferente
        cmdEndCobranca.Enabled = False
        cmdEndCobranca.Picture = cmdExtra(2).Picture
    End Select
    If gCarga = False Then
        Call gpHabSalvar
        Select Case index
        Case 0      'Sim.
            If gM46V999.gfPreenchido(txtTipLogradouro.Text) Then
                frmT46V105A.txtTipLogradouro.Text = txtTipLogradouro
            End If
            If gM46V999.gfPreenchido(txtNomLogradouro.Text) Then
                frmT46V105A.txtNomLogradouro.Text = txtNomLogradouro.Text
            End If
            If gM46V999.gfPreenchido(txtNumero.Text) Then
                frmT46V105A.txtNumero.Text = txtNumero.Text
            End If
            If gM46V999.gfPreenchido(txtComplemento.Text) Then
                frmT46V105A.txtComplemento.Text = txtComplemento.Text
            End If
            If gM46V999.gfPreenchido(txtBairro.Text) Then
                frmT46V105A.txtBairro.Text = txtBairro.Text
            End If
            If gM46V999.gfPreenchido(txtCidade.Text) Then
                frmT46V105A.txtCidade.Text = txtCidade.Text
            End If
            If cboUF.ListIndex <> -1 Then
                If frmT46V105A.cboUF.ListCount > 0 Then
                    For lI = 0 To (frmT46V105A.cboUF.ListCount - 1)
                        If frmT46V105A.cboUF.List(lI) = cboUF.Text Then
                            frmT46V105A.cboUF.ListIndex = lI
                            Exit For
                        End If
                    Next lI
                End If
            End If
            If gM46V999.gfPreenchido(txtCep.Text) Then
                frmT46V105A.txtCep.Text = txtCep.Text
            End If
            If gM46V999.gfPreenchido(txtTelefones.Text) Then
                frmT46V105A.txtTelefones.Text = txtTelefones.Text
            End If
            frmT46V105A.Show vbModal, Me
        Case 1      'N�o tem endere�o de cobran�a diferente
            On Error Resume Next
            Unload frmT46V105A
            On Error GoTo 0
        End Select
    End If
End Sub
Private Sub optEndCobranca_GotFocus(index As Integer)
    Set mAntControle = ActiveControl
End Sub
Private Sub optEndCobranca_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optForPagamento_Click(index As Integer)
    Call gpHabSalvar
    Select Case index
    Case 0  'Carn�.
        'panProp(8).Enabled = True
        optPagAntecipado(1).Enabled = True
        optPagAntecipado(2).Enabled = True

        cmdForPagCarne.Enabled = True
        cmdForPagCarne.Picture = cmdExtra(1).Picture
        cmdForPagDebito.Enabled = False
        cmdForPagDebito.Picture = cmdExtra(2).Picture
        cmdForPagDebito1oBoleto.Enabled = False
        cmdForPagDebito1oBoleto.Picture = cmdExtra(2).Picture
        'Ficha 900046
        If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
            cmdPagAntBloqueto.Visible = True
        End If
        If gCarga = False Then
            frmT46V109A.Show vbModal, Me
        End If
        If optPagAntecipado(2).Value = False Then
            If gM46V999.gfDisponivel("frmT46V110A") Then
                Unload frmT46V110A
            End If
        End If
        If gM46V999.gfDisponivel("frmT46V185A") Then
            Unload frmT46V185A
        End If
    Case 1  'D�bito autom�tico.
        'panProp(8).Enabled = True
        optPagAntecipado(1).Enabled = True
        optPagAntecipado(2).Enabled = True

        cmdForPagCarne.Enabled = False
        cmdForPagCarne.Picture = cmdExtra(2).Picture
        cmdForPagDebito.Enabled = True
        cmdForPagDebito.Picture = cmdExtra(1).Picture
        cmdForPagDebito1oBoleto.Enabled = False
        cmdForPagDebito1oBoleto.Picture = cmdExtra(2).Picture
        'Ficha 900046
        If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
            cmdPagAntBloqueto.Visible = False
        End If
        If gCarga = False Then
            Call frmT46V110A.mpInstanciarForm(2)
        End If
        If gM46V999.gfDisponivel("frmT46V109A") Then
            Unload frmT46V109A
        End If
        If gM46V999.gfDisponivel("frmT46V185A") Then
            Unload frmT46V185A
        End If
    Case 2  'D�bito autom�tico com 1� em boleto.
        'panProp(8).Enabled = False
        optPagAntecipado(1).Value = True
        optPagAntecipado(1).Enabled = False
        optPagAntecipado(2).Enabled = False

        cmdForPagCarne.Enabled = False
        cmdForPagCarne.Picture = cmdExtra(2).Picture
        cmdForPagDebito.Enabled = False
        cmdForPagDebito.Picture = cmdExtra(2).Picture
        cmdForPagDebito1oBoleto.Enabled = True
        cmdForPagDebito1oBoleto.Picture = cmdExtra(1).Picture

        If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
            cmdPagAntBloqueto.Visible = False
        End If
        If gCarga = False Then
            frmT46V185A.Show vbModal, Me
        End If
        If gM46V999.gfDisponivel("frmT46V109A") Then
            Unload frmT46V109A
        End If
        If gM46V999.gfDisponivel("frmT46V110A") Then
            Unload frmT46V110A
        End If
    End Select
    
    txtJuros.Text = Format(frmT46V142A.getJuros(IIf(optForPagamento(0).Value, e_FormaPagamento_Carne, e_FormaPagamento_Debito), optPagAntecipado(1).Value, val(txtQuaParcelas.Text)), "standard")
End Sub
Private Sub optForPagamento_GotFocus(index As Integer)
    Set mAntControle = ActiveControl
End Sub
Private Sub optForPagamento_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optPagAntecipado_Click(index As Integer)
    Dim lobjstcBoleto As stcA28V720B
    Dim lobjBoleto As clsA28V720A
    Call gpLimPremios

    Select Case index
    Case 1  'Sim - Antecipado.
        'Ficha 900046
        If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
            cmdPagAntBloqueto.Enabled = True
            cmdPagAntBloqueto.Picture = cmdExtra(1).Picture
            cmdPagAntDebito.Enabled = False
            cmdPagAntDebito.Picture = cmdExtra(2).Picture

            txtValorPgAnt.Visible = True
            frmT46V108A.txtPagAntBloqueto.Enabled = True
        Else
            txtValorPgAnt.Text = "0,00"
            txtValorPgAnt.Visible = False
        End If
        If gCarga = False Then
            '***BOLETO***
            'Ficha 900046
            If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And optForPagamento(0).Value = True Then
                'Somente com usu�rio yasuda � que
                'poder� digitar o n�mero do boleto. Caso contr�rio este
                'n�mero ser� gerado na rotina de impress�o de formul�rio.
                frmT46V108A.Show vbModal, Me
            End If
        End If
        If gM46V999.gfDisponivel("frmT46V110A") = True Then
            If optForPagamento(1) = False Then
                Unload frmT46V110A
            End If
        End If
        If gM46V999.gfDisponivel("frmT46V185A") = True Then
            If optForPagamento(2) = False Then
                Unload frmT46V185A
            End If
        End If
        If txtValorPgAnt.Visible Then
            txtValorPgAnt.Enabled = True
            If txtValorPgAnt.Enabled = True Then txtValorPgAnt.SetFocus    'Exibe mensagem "Invalid procedure call" se ainda n�o estiver pronto.
        End If
    Case 2  'N�o - a 30 dias.
        cmdPagAntBloqueto.Enabled = False
        cmdPagAntBloqueto.Picture = cmdExtra(2).Picture
        cmdPagAntDebito.Enabled = True
        cmdPagAntDebito.Picture = cmdExtra(1).Picture
        'Ficha 900046
        txtValorPgAnt.Text = "0,00"
        txtValorPgAnt.Visible = False
        frmT46V108A.txtPagAntBloqueto.Text = "0"
        frmT46V108A.txtPagAntBloqueto.Enabled = False
        'Ficha 900046
        '            If gCarga = False Then
        '                Call frmT46V110A.mpInstanciarForm(1)
        '                gCarga = True
        '                optForPagamento(1).Value = True
        '                gCarga = False
        '            End If
        If gM46V999.gfDisponivel("frmT46V108A") = True Then
            Unload frmT46V108A
        End If
        If txtValorPgAnt.Visible Then
            txtValorPgAnt.Enabled = True
            txtValorPgAnt.SetFocus
        End If
    Case 9  'N�o.
        cmdPagAntBloqueto.Enabled = False
        cmdPagAntBloqueto.Picture = cmdExtra(2).Picture
        cmdPagAntDebito.Enabled = False
        cmdPagAntDebito.Picture = cmdExtra(2).Picture
        If gM46V999.gfDisponivel("frmT46V108A") = True Then
            Unload frmT46V108A
        End If
        If gM46V999.gfDisponivel("frmT46V110A") = True Then
            If optForPagamento(1) = False Then
                Unload frmT46V110A
            End If
        End If
        If gM46V999.gfDisponivel("frmT46V185A") = True Then
            If optForPagamento(2) = False Then
                Unload frmT46V185A
            End If
        End If
        txtValorPgAnt.Text = "0,00"
        txtValorPgAnt.Enabled = False
    End Select

    txtJuros.Text = Format(frmT46V142A.getJuros(IIf(optForPagamento(0).Value, e_FormaPagamento_Carne, e_FormaPagamento_Debito), optPagAntecipado(1).Value, val(txtQuaParcelas.Text)), "standard")
    Call txtQuaParcelas_LostFocus   'Considera informa��o de quantidade de parcelas.
End Sub
Private Sub optPagAntecipado_GotFocus(index As Integer)
    Set mAntControle = ActiveControl
End Sub
Private Sub optPagAntecipado_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub optProlabore_Click(index As Integer)
    Dim lMensagem As String

    If gCarga = True Then
        Exit Sub
    End If

    Call gpHabSalvar

    Select Case index
    Case 0  'Sim.
        Dim i As Integer
        Dim blnCorretorPreenchido As Boolean
        For i = 1 To frmT46V102A.txtCorretor.Count
            If frmT46V102A.txtCorretor(i).Tag <> "" Then
                blnCorretorPreenchido = True
                Exit For
            End If
        Next i

        If blnCorretorPreenchido Then
            Call mnuProlaboristas
        Else
            MsgBox "Para atribuir os prolaboristas, primeiro selecione o corretor.", vbInformation + vbOKOnly, "Erro"
            gCarga = True
            optProlabore(1).Value = True
            gCarga = False
            Exit Sub
        End If

    Case 1  'N�o.
        lMensagem = "AGE0016 - " & gM46V999.gPrefixo & "Exclu�da op��o de prolabore. " & _
                    "Lista de prolaboristas ser� desconsiderada. Confirma esta op��o?"
        If MsgBox(lMensagem, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Call gM46V999.gpGraLog(0, lMensagem & " Sim.")
            On Error Resume Next
            Unload frmT46V180A
            On Error GoTo 0
            Exit Sub
        End If
        Call gM46V999.gpGraLog(0, lMensagem & " N�o.")
        optProlabore(0).Value = True
        If optProlabore(0).Visible And optProlabore(0).Enabled Then
            optProlabore(0).SetFocus
        End If
    End Select

End Sub

Private Sub optSexo_GotFocus(index As Integer)
    Set mAntControle = ActiveControl
End Sub
Private Sub optSexo_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optTipPessoa_Click(index As Integer)
    frmT46V101A.optTipPessoa(index) = True
    Select Case index
    Case 0  'Pessoa f�sica.
        panPesFisica.Visible = True
        panPesJuridica.Visible = False

        'Fichas 1200168 e 1200169 Inclus�o dos campos de listagem Renda e Faixa de Faturamento Pessoa f�sica e jur�dica
        '           'Ficha 900417
        '            If gM46V999.gCodProduto >= 968 Then
        '                lblProp(10).Caption = "          Profiss�o: "
        '            Else
        '                lblProp(10).Caption = "Empr. onde trabalha: "
        '            End If

        'Ficha 1000307
        If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
            lblProp(28).Visible = True
            lblProp(29).Visible = True
            cboRegCons.Visible = True
            cboRegCons.Enabled = True
            txtnumregistro.Visible = True
            txtnumregistro.Enabled = True
            Call cboRegCons_Change
        Else
            lblProp(28).Visible = False
            lblProp(29).Visible = False
            cboRegCons.Visible = False
            cboRegCons.Enabled = False
            txtnumregistro.Visible = False
            txtnumregistro.Enabled = False
        End If

    Case 1  'Pessoa jur�dica.
        panPesFisica.Visible = False
        panPesJuridica.Visible = True

        'Ficha 1000307
        lblProp(28).Visible = False
        lblProp(29).Visible = False
        cboRegCons.Visible = False
        cboRegCons.Enabled = False
        txtnumregistro.Visible = False
        txtnumregistro.Enabled = False
        gObrRegistro = False
                
        Dim blnVisualizaCondominioConstituido As Boolean
        blnVisualizaCondominioConstituido = False
        
        With frmT46V101A
            blnVisualizaCondominioConstituido = gM46V999.gCod_Ramo = e_CodRamo_Condominio And .objstcPedLoc.Tip_Emissao <> e_TipEmissao_Endosso And gM46V999.gCodProduto >= 1100
            If Not blnVisualizaCondominioConstituido And gM46V999.gCod_Ramo = e_CodRamo_Condominio And .objstcPedLoc.Tip_Emissao = e_TipEmissao_Endosso Then
                gM46V999.gNum_Apolice = .objstcPedido.Num_Apol_An
                If gM46V999.gTarifa_Apolice >= 1100 Then
                    blnVisualizaCondominioConstituido = True
                End If
            End If
        End With
        
        If blnVisualizaCondominioConstituido Then
            cboRamAtividade.Visible = False
            lblProp(1).Visible = False
        Else
            cboRamAtividade.Visible = True
            lblProp(1).Visible = True
        End If
        
    End Select
    If gCarga = False Then
        gAltComum = True
    End If
End Sub
Private Sub optTipPessoa_GotFocus(index As Integer)
    Set mAntControle = ActiveControl
End Sub
Private Sub optTipPessoa_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub optVencimento_Click(index As Integer)
    Dim lMensagem As String

    '1. Posiciona controles.
    Select Case index
    Case 0  'Sim.
        cmdAjustarDatas.Enabled = True
        cmdAjustarDatas.Picture = cmdExtra(1).Picture
        If IsNumeric(txtQuaParcelas) Then
        If txtQuaParcelas = 1 And optPagAntecipado(1).Value = True Then
            lMensagem = "FGE0120 - N�o � permitido ajustar vencimento quando somente uma parcela."
            MsgBox lMensagem
            If txtQuaParcelas.Visible Then txtQuaParcelas.SetFocus
            Exit Sub
        End If
        End If
    Case 1  'N�o.
        cmdAjustarDatas.Enabled = False
        cmdAjustarDatas.Picture = cmdExtra(2).Picture
    End Select

    If gCarga = True Then
        Exit Sub
    End If
    Call gpHabSalvar
    '4. Processa op��o do usu�rio.
    Select Case index
    Case 0  'Sim.
        Call cmdAjustarDatas_Click
    Case 1  'N�o.
        lMensagem = "AGE0016 - " & gM46V999.gPrefixo & "Exclu�da op��o de ajuste de vencimento. " & _
                    "Datas ajustadas ser�o desconsideradas. Confirma esta op��o?"
        If MsgBox(lMensagem, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Call gM46V999.gpGraLog(0, lMensagem & " Sim.")
            On Error Resume Next
            Unload frmT46V121A
            On Error GoTo 0
            Exit Sub
        End If
        Call gM46V999.gpGraLog(0, lMensagem & " N�o.")
        optVencimento(0).Value = True
        If optVencimento(0).Visible And optVencimento(0).Enabled Then
            optVencimento(0).SetFocus
        End If
    End Select
End Sub

Private Sub Text1_LostFocus()
    Dim lcol As Integer
    Dim lrow As Integer

         
    If Len(Text1.Tag) Then
        lrow = Left(Text1.Tag, InStr(Text1.Tag, ","))
        lcol = Right(Text1.Tag, InStr(Text1.Tag, ",") - 1)

        If lrow < flxTelefones.Rows Then
            flxTelefones.TextMatrix(lrow, lcol) = Text1.Text
        End If
    End If
    Text1.Visible = False
    
    'Adicionado por Luiz Alberto em 11/12/2015
    If lcol = 3 And Trim(Text1.Text) <> "" Then
        If Mid(Text1.Text, 1, 1) = 0 Then
           MsgBox "Telefone Inv�lido, n�mero come�ando com zero. Favor Verificar!", vbCritical, "Aten��o"
           Text1.Text = ""
           flxTelefones.TextMatrix(lrow, lcol) = Text1.Text
           'Text1.SetFocus
           Exit Sub
         End If
        
        If Len(Text1.Text) < 8 Then
           MsgBox "Telefone Inv�lido, n�mero inferior a 8 d�gitos. Favor Verificar!", vbCritical, "Aten��o"
           Text1.Text = ""
           flxTelefones.TextMatrix(lrow, lcol) = Text1.Text
           'Text1.SetFocus
           Exit Sub
         End If
         
        If Not VALIDATEXTO(Text1.Text) Then
           MsgBox "Telefone Inv�lido/Viciado. Favor Verificar!", vbCritical, "Aten��o"
           Text1.Text = ""
           flxTelefones.TextMatrix(lrow, lcol) = Text1.Text
           'Text1.SetFocus
           Exit Sub
         End If
        'TRATAMENTO NONO DIGITO
        If InStr(1, Text1.Text, "-") = 0 Then
            If Len(Text1.Text) < 9 Then
                flxTelefones.TextMatrix(lrow, lcol) = Mid(Text1.Text, 1, 4) & "-" & Mid(Text1.Text, 5, 4)
            Else
                flxTelefones.TextMatrix(lrow, lcol) = Mid(Text1.Text, 1, 5) & "-" & Mid(Text1.Text, 6, 4)
            End If
        Else
            flxTelefones.TextMatrix(lrow, lcol) = Text1.Text
        End If
    Else
        flxTelefones.TextMatrix(lrow, lcol) = Text1.Text
    End If
  
End Sub
'Adicionado por Luiz Alberto em 11/12/2015
Private Sub txtIdeNumDocumento_LostFocus()
    If Not VALIDATEXTO(txtIdeNumDocumento.Text) And gM46V999.gfPreenchido(txtIdeNumDocumento.Text) Then
       MsgBox "Numero do documento Inv�lido. Favor Verificar!", vbInformation, "Aten��o"
       txtIdeNumDocumento.Text = ""
       Exit Sub
     End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

    Case 8    'BackSpace
        Exit Sub

    Case 13    'Enter
        Exit Sub

    Case 22    'Ctrl+V
        Select Case flxTelefones.Col
        Case 5    'Observa��o
            Text1.Text = Trim(Clipboard.GetText)
        Case 2    'DDD
            If IsNumeric(Trim(Clipboard.GetText)) Then
                Text1.Text = Trim(Clipboard.GetText)
            End If
        Case 3    'Telefone
            If IsNumeric(Trim(Clipboard.GetText)) Then
                Text1.Text = Trim(Clipboard.GetText)
            End If

        Case 4    'Ramal
            If IsNumeric(Trim(Clipboard.GetText)) Then
                Text1.Text = Trim(Clipboard.GetText)
            End If
        End Select

    Case 48 To 57    'Zero a nove

        Select Case flxTelefones.Col

        Case 5    'Observa��o
            If Len(Text1.Text) < mMaxLength_Obs Then
                Exit Sub
            End If

        Case 2    'DDD
            Text1.Text = Text1.Text & Chr(KeyAscii)

        Case 3    'Telefone
            Text1.Text = Text1.Text & Chr(KeyAscii)

        Case 4    'Ramal
            Text1.Text = Text1.Text & Chr(KeyAscii)

        End Select

    Case 32, 33, 35 To 38, 40 To 47, 61, 65 To 93, 95, 97 To 123, 126
        If flxTelefones.Col = 5 Then
            Exit Sub
        End If

    Case Else
        Beep
    End Select

    KeyAscii = 0
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub txtBairro_Change()
    Call gpHabSalvar
    If optEndCobranca(1) = True Then    'Endere�o de cobran�a igual.
        frmT46V105A.txtBairro.Text = txtBairro.Text
    End If
End Sub
Private Sub txtBairro_GotFocus()
    Set mAntControle = ActiveControl
    With txtBairro
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtBairro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtBairro_LostFocus()
    txtBairro.Text = gM46V999.gfRetirarCaracterEspecial(txtBairro.Text, True)
End Sub

Private Sub txtCEP_Change()
    Call gpHabSalvar
End Sub
Private Sub txtCEP_GotFocus()
    Set mAntControle = ActiveControl
    With txtCep
        If gM46V999.gfPreenchido(.Text) Then
            mAntCEP = .Text
        Else
            mAntCEP = ""
        End If
        .SelStart = 0
        .SelLength = .MaxLength
        .Text = Replace(.Text, "-", "")
    End With
End Sub
Private Sub txtCEP_KeyPress(KeyAscii As Integer)
    Dim lI As Byte
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    Select Case KeyAscii
    Case 48 To 57
        'Caracteres v�lidos
        lI = txtCep.SelStart
        If Len(txtCep.Text) = 8 And lI <> 0 Then
            KeyAscii = 0
        End If
    Case 8
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtCEP_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(txtCep.Text) > 8 And InStr(1, txtCep.Text, "-") = 0 Then
        txtCep.Text = Left(txtCep.Text, 8)
    End If
End Sub

Private Sub txtCEP_LostFocus()
    Dim lbdP0043000 As ADODB.Connection    'Arquivo - P0043000.
    Dim lMensagem As String         'Mensagem.
    Dim lrsP0043000 As clsYasRecordSet  'Registro - P0043000.
    Dim lPos As Byte
    Dim lCEP As String
    Dim lSelect As String
    Screen.MousePointer = vbHourglass

    With txtCep
        '1. Trata campo do CEP.
        If gM46V999.gfPreenchido(mAntCEP) Then
            If gM46V999.gfPreenchido(.Text) Then
                If mAntCEP = .Text Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            Else
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Else
            If Not gM46V999.gfPreenchido(.Text) Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        .Text = Trim$(.Text)
        Select Case Len(.Text)
        Case Is < 8
            lPos = InStr(1, .Text, "-")
            If lPos > 6 Then
                .Text = .Text
            ElseIf lPos < 6 And lPos <> 0 Then
                .Text = Format(Left(.Text, lPos - 1), "00000") & "-" & Format(Left(Mid(.Text, lPos + 1), 3), "000")
            Else
                .Text = Format(Format(.Text, "00000000"), "00000-000")
            End If
        Case 8
            lPos = InStr(1, .Text, "-")
            If lPos > 6 Then
                .Text = .Text
            ElseIf lPos <= 6 And lPos <> 0 Then
                .Text = Format(Left(.Text, lPos - 1), "00000") & Left(Mid(.Text, lPos + 1) & "000", 3)
            End If

            If Not IsNumeric(.Text) Then
                Call gM46V999.gpGraLog(1, "FGE0039 - " & gM46V999.gPrefixo & _
                                          "CEP de pernoite do ve�culo inv�lido (n�o num�rico).")
                .Text = ""
                If .Visible = True And .Enabled = True Then
                    .SetFocus
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            .Text = Mid$(.Text, 1, 5) & "-" & Mid$(.Text, 6, 3)
            lCEP = .Text
        Case 9
            If Not IsNumeric(Mid$(.Text, 1, 5)) Or Not IsNumeric(Mid$(.Text, 7, 3)) Then
                Call gM46V999.gpGraLog(1, "FGE0074 - " & gM46V999.gPrefixo & "CEP inv�lido (endere�o do segurado - " & .Text & _
                                        " - n�o num�rico - fora do formato nnnnn-nnn).")
                .Text = ""
                If .Visible = True And .Enabled = True Then
                    .SetFocus
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            If Mid$(.Text, 6, 1) <> "-" Then
                Call gM46V999.gpGraLog(1, "FGE0074 - " & gM46V999.gPrefixo & "CEP inv�lido (endere�o do segurado - " & .Text & _
                                        " - n�o tem tra�o na posi��o 6 - fora do formato nnnnn-nnn).")
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End Select
        '3. Obt�m CEP.
        'a) Abre arquivo.
        If gM46V999.gfAbrBasDados("P0043000", gM46V999.gTipP0043000, gM46V999.gEndP0043000, lbdP0043000, lMensagem) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'b) Obt�m registro.
        lSelect = "SELECT * FROM P0043000 WHERE Num_CEP = '" & Left$(.Text, 5) & Right$(.Text, 3) & "'"
        If gM46V999.gfObtRegistro(lbdP0043000, lSelect, lrsP0043000, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdP0043000)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If lrsP0043000.EOF = True Then
            Call gM46V999.gpGraLog(1, "FNL0007 - " & gM46V999.gPrefixo & "CEP n�o localizado (" & .Text & ").")
            Call gM46V999.gpFecha2(lbdP0043000)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'c) Posiciona endere�o.
        If gM46V999.gfPreenchido(lrsP0043000!tip_logr) Then
            txtTipLogradouro.Text = lrsP0043000!tip_logr
        Else
            txtTipLogradouro.Text = ""
        End If
        If gM46V999.gfPreenchido(lrsP0043000!nom_logr) Then
            txtNomLogradouro.Text = Left(lrsP0043000!nom_logr, IIf(InStr(1, lrsP0043000!nom_logr, "-") > 0, InStr(1, lrsP0043000!nom_logr, "-") - 2, 200))
        Else
            txtNomLogradouro.Text = ""
        End If
        txtNumero.Text = ""
        txtComplemento.Text = ""
        If gM46V999.gfPreenchido(lrsP0043000!Nom_Bairro) Then
            txtBairro.Text = lrsP0043000!Nom_Bairro
        Else
            txtBairro.Text = ""
        End If
        If gM46V999.gfPreenchido(lrsP0043000!Nom_Cid) Then
            txtCidade.Text = lrsP0043000!Nom_Cid
        Else
            txtCidade.Text = ""
        End If
        If gM46V999.gfPreenchido(lrsP0043000!Sig_UF) Then
            cboUF.Text = lrsP0043000!Sig_UF
        Else
            cboUF.ListIndex = -1
        End If
        'd) Fecha tabela e arquivo.
        Call gM46V999.gpFecha2(lbdP0043000)
    End With

    Screen.MousePointer = vbDefault
End Sub
Private Sub txtCidade_Change()
    Call gpHabSalvar
    If optEndCobranca(1) = True Then    'Endere�o de cobran�a igual.
        frmT46V105A.txtCidade.Text = txtCidade.Text
    End If
End Sub
Private Sub txtCidade_GotFocus()
    With txtCidade
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub
Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCidade_LostFocus()
    txtCidade.Text = gM46V999.gfRetirarCaracterEspecial(txtCidade.Text, True)
End Sub

Private Sub txtCNPJ_Change()
    Call gpHabSalvar
End Sub
Private Sub txtCNPJ_GotFocus()
    With txtCnpj
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub
Private Sub txtCNPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtCNPJ_LostFocus()
    Dim lAuxiliar As String    'Auxiliar.

    Call gM46V999.gpForCNPJ(1, txtCnpj.Text, lAuxiliar)
    If lAuxiliar <> "" Then
        txtCnpj.Text = lAuxiliar
    Else
        If Len(txtCnpj.Text) = 18 And InStr(1, txtCnpj.Text, ".") = 0 Then
            txtCnpj.Text = ""
        Else
            If lAuxiliar <> "" Then
                txtCnpj.Text = lAuxiliar
            End If
        End If
    End If

    'Proposta Eletr�nica
    If gM46V999.gNovaTransmissaoMulti = True Then
        If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
            If frmT46V101A.optTipPessoa(0).Value = True Then
                Call gM46V999.gpForCPF(2, txtCnpj.Text, mAtuCNPJ)  'F�sica
            Else
                Call gM46V999.gpForCNPJ(2, txtCnpj.Text, mAtuCNPJ)    'Jur�dica
            End If
            If mAtuCNPJ <> mAntCNPJ Then
                frmT46V101A.txtSegurado.Enabled = True
                txtSegurado.Enabled = True
            Else
                frmT46V101A.txtSegurado.Text = mAntSeg
                frmT46V101A.txtSegurado.Enabled = False
                txtSegurado.Text = mAntSeg
                txtSegurado.Enabled = False
            End If
        End If
    End If
    If cmdSiscad(1).Visible And gM46V999.PerfilAcesso <> e_Perfil_CALC And gM46V999.PerfilAcesso <> e_Perfil_CALC_Gestor Then
        Call cmdSiscad_Click(1)
    End If

End Sub
Private Sub txtCodSegurado_Change()
    Call gpHabSalvar
End Sub
Private Sub txtCodSegurado_LostFocus()
    txtCodSegurado.Text = Format(val(txtCodSegurado.Text), "000000")
    frmT46V101A.txtSegurado.Tag = txtCodSegurado.Text
    
       ''2016-08-05: Check Exception de Juros
    Call frmT46V101A.CheckExceptionJuros
End Sub
Private Sub txtNumero_Change()
    mAltEndereco = True
    Call gpHabSalvar
    If optEndCobranca(1) = True Then    'Endere�o de cobran�a igual.
        frmT46V105A.txtNumero.Text = txtNumero.Text
    End If
End Sub
Private Sub txtNumero_GotFocus()
    With txtNumero
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub
Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtNumero_LostFocus()
    If gM46V999.gfPreenchido(txtNumero.Text) Then
        Call mpVerEndSegurado
    End If
    txtNumero.Text = gM46V999.gfRetirarCaracterEspecial(txtNumero.Text, True)
End Sub
Private Sub txtComplemento_Change()
    mAltEndereco = True
    Call gpHabSalvar
    If optEndCobranca(1) = True Then    'Endere�o de cobran�a igual.
        frmT46V105A.txtComplemento.Text = txtComplemento.Text
    End If
End Sub
Private Sub txtComplemento_GotFocus()
    With txtComplemento
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub
Private Sub txtComplemento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtComplemento_LostFocus()
    If gM46V999.gfPreenchido(txtComplemento.Text) Then
        Call mpVerEndSegurado
    End If
    txtComplemento.Text = gM46V999.gfRetirarCaracterEspecial(txtComplemento.Text, True)
End Sub
Private Sub txtCorretor_Change(index As Integer)
    txtCorretor(index).Tag = ""
    Call gpHabSalvar
    If gCarga = False Then
        gAltComum = True
    End If
End Sub

Private Sub txtCorretor_GotFocus(index As Integer)
    strCorretor = txtCorretor(index).Text
    strCodigoCorretor = txtCorretor(index).Tag
    Set mAntControle = ActiveControl
End Sub
Private Sub txtCorretor_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Public Sub txtCorretor_LostFocus(index As Integer)
    Dim lCodProdutor As Integer    'C�digo do produtor.
    Dim lCodUniProdutiva As Integer    'C�digo da unidade produtiva.
    Dim lCorretor As String    'Corretor.
    Dim lI As Double             'Utilizado com For...Next.
    Dim lMensagem As String    'Mensagem.
    Dim lNomProdutor As String    'Nome do produtor.
    Dim lNomUniProdutiva As String    'Nome da unidade produtiva.
    Dim lQuaCorretores As Double    'Quantidade de corretores selecionados.
    Dim lSaida() As Variant    'Resultado da busca de corretores.
    Dim lTotal As Double             'Total
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass

    '1. Se c�digo do corretor preenchido, encerrar processo.
    If gM46V999.gfPreenchido(txtCorretor(index).Tag) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    DoEvents
    
    'Carregar novamente as coberturas por corretor pois foi incluido a funcionalidade de franquia por corretor
    If Not gCarga Then
        gAltDadosCobert = True
        Unload frmT46V117A
        Unload frmT46V179A
        gM46V999.gChamaSelecionarCoberturas = True
    End If
    
    '2. Limpa campos de unidade produtiva e produtor.
    If index = 1 Then
        lblNomUnidade.Tag = " "
        lblNomUnidade.Caption = " "
        lblNomProdutor.Tag = " "
        lblNomProdutor.Caption = " "
    End If

    '3. Se nome do corretor n�o preenchido, encerra processo.
    If Not gM46V999.gfPreenchido(txtCorretor(index).Text) Then
        txtCorComissao(index).Text = "0,00"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '4. Posiciona campo com o que foi informado pelo usu�rio.
    lCorretor = gM46V999.gfRetirarCaracterEspecial(Trim$(txtCorretor(index).Text))

    '5. Se foram digitadas at� 6 posi��es num�ricas, trata de c�digo de corretor.
    '   Obt�m corretor, unidade produtiva e produtor.
    If IsNumeric(lCorretor) And Len(CStr(val(lCorretor))) <= 6 Then
        If gM46V999.gfObtCorretor(1, val(lCorretor), lSaida, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            txtCorretor(index).Text = ""
            If txtCorretor(index).Visible = True And txtCorretor(index).Enabled = True Then
                txtCorretor(index).SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'Se for corretor 1 e tarifa >= 877 verifico se corretor escolhido � permitido.
        If index = 1 And gCodCorrInst <> lSaida(1, 1) And gCodCorrInst <> 0 And gM46V999.gCodProtocolo = 0 Then
            'Ficha 1000153 - Automatiza��o do agrupamento de corretores no Syas ap�s o cadastro realizado pelo help-desk.
            If gM46V999.gfVerCorrPermitido_TBS_Corr_Grupo(gCodCorrInst, lSaida(1, 1), CInt(gM46V999.gCod_Ramo), lMensagem) = False Then
                Call gM46V999.gpGraLog(1, "FGE0017 - C�digo corretor n�o permitido para este usu�rio (" & gM46V999.gNomUsuario & ")")
                txtCorretor(index).Text = ""
                If txtCorretor(index).Visible = True And txtCorretor(index).Enabled = True Then
                    txtCorretor(index).SetFocus
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        txtCorretor(index).Text = lSaida(1, 3)
        txtCorretor(index).Tag = lSaida(1, 1)
        If index = 1 Then
            If mfPosUniProProdutor(txtCorretor(1).Tag, lMensagem) = False Then
                Call gM46V999.gpGraLog(1, lMensagem)
                txtCorretor(index).Text = ""
                If txtCorretor(index).Visible = True And txtCorretor(index).Enabled = True Then
                    txtCorretor(index).SetFocus
                End If
                If gfCarCboBancos(2, frmT46V109A.cboCarBanco, lMensagem, txtCorretor(index).Tag) = False Then
                    Call gM46V999.gpGraLog(2, lMensagem)
                    Exit Sub
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If

        'Ficha 1000601
        If Not IsNumeric(lCorretor) Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        '        'Corre��o Ficha - 1000601 - Permitir Somente desconto nas tarifas 998 at� 993
        '        If gM46V999.gCodProduto <= 999 Then
        '            If lCorretor = 425201 Or lCorretor = 383901 Or lCorretor = 611501 Or lCorretor = 680301 Or lCorretor = 18301 And _
                     '                (gM46V999.gCodProduto >= 988 And gM46V999.gCodProduto <= 993) And gM46V111.gCod_Plano = 1 Then
        '                gM46V998.gLiberaDescontoEspecial = True
        '                frmT46V101A.optProtecao(1).Value = True
        '            Else
        '                gM46V998.gLiberaDescontoEspecial = False
        '                frmT46V101A.optProtecao(1).Value = True
        '            End If
        '        Else
        '            gM46V998.gLiberaDescontoEspecial = False
        '            frmT46V101A.optProtecao(1).Value = True
        '        End If
        '        '1101045 - Tarifa 1022 e 1023 Novembro 2011
        '        If (frmT46V101A.objstcPedido.Cod_Produto >= 1022 And frmT46V101A.objstcPedido.Cod_Produto <= 1033) Or _
                 '            (frmT46V101A.objstcPedido.Cod_Produto >= 1038) Then     '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 8
        '            gM46V998.gLiberaDescontoEspecial = True
        '        Else
        '            gM46V998.gLiberaDescontoEspecial = False
        '        End If
        '
        '        'Ajustado na ficha 1200733 1042 e 1043
        '        If (frmT46V101A.objstcPedido.Cod_Produto >= 1042 And frmT46V101A.objstcPedido.Cod_Produto <= 1043) And _
                 '           gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
        '            gM46V998.gLiberaDescontoEspecial = True
        '        End If

        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '5. Se preenchido com menos de 3 posi��es, encerra processo.
    If Len(lCorretor) < 3 Then
        Call gM46V999.gpGraLog(1, "FGE0075 - Corretor inv�lido (digitar pelo menos 3 posi��es).")
        txtCorretor(index).Text = ""
        If txtCorretor(index).Visible = True And txtCorretor(index).Enabled = True Then
            txtCorretor(index).SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '6. Obt�m registros com informa��o do usu�rio.
    'a) Obt�m lista de corretores
    If gM46V999.gfObtCorretor(2, lCorretor, lSaida, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        txtCorretor(index).Text = ""
        If txtCorretor(index).Visible = True And txtCorretor(index).Enabled = True Then
            txtCorretor(index).SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'b) Se campo de sa�da estiver vazio, nada a fazer.
    If IsNull(lSaida) Then
        txtCorretor(index).Text = ""
        If txtCorretor(index).Visible = True And txtCorretor(index).Enabled = True Then
            txtCorretor(index).SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'c) Posiciona quantidade de corretores encontrados.
    lQuaCorretores = UBound(lSaida)
    'd) Se encontrou apenas 1 corretor, posiciona dados deste corretor e encerra processo.
    If lQuaCorretores = 1 Then
        'Se for corretor 1 e tarifa >= 877 verifico se corretor escolhido � permitido.
        If index = 1 And gCodCorrInst <> lSaida(1, 1) And gCodCorrInst <> 0 And gM46V999.gCodProtocolo = 0 Then
            'Ficha 1000153 - Automatiza��o do agrupamento de corretores no Syas ap�s o cadastro realizado pelo help-desk.
            If gM46V999.gfVerCorrPermitido_TBS_Corr_Grupo(gCodCorrInst, lSaida(1, 1), CInt(gM46V999.gCod_Ramo), lMensagem) = False Then
                Call gM46V999.gpGraLog(1, "FGE0017 - C�digo corretor n�o permitido para este usu�rio (" & gM46V999.gNomUsuario & ")")
                txtCorretor(index).Text = ""
                If txtCorretor(index).Visible = True And txtCorretor(index).Enabled = True Then
                    txtCorretor(index).SetFocus
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        txtCorretor(index).Text = lSaida(1, 3)
        txtCorretor(index).Tag = lSaida(1, 1)
        If index = 1 Then
            If mfPosUniProProdutor(txtCorretor(1).Tag, lMensagem) = False Then
                Call gM46V999.gpGraLog(1, lMensagem)
            End If
        End If
        'Ficha 1000601
        If Not IsNumeric(lSaida(1, 1)) Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        '        'Corre��o Ficha - 1000601 - Permitir Somente desconto nas tarifas 998 at� 993
        '        If gM46V999.gCodProduto <= 999 Then
        '            If lSaida(1, 1) = 425201 Or lSaida(1, 1) = 383901 Or lSaida(1, 1) = 611501 Or lSaida(1, 1) = 680301 Or lSaida(1, 1) = 18301 And _
                     '                (gM46V999.gCodProduto >= 988 And gM46V999.gCodProduto <= 993) And gM46V111.gCod_Plano = 1 Then
        '                gM46V998.gLiberaDescontoEspecial = True
        '                frmT46V101A.optProtecao(1).Value = True
        '            Else
        '                gM46V998.gLiberaDescontoEspecial = False
        '                frmT46V101A.optProtecao(1).Value = True
        '            End If
        '        Else
        '            gM46V998.gLiberaDescontoEspecial = False
        '            frmT46V101A.optProtecao(1).Value = True
        '        End If
        '        '1101045 - Tarifa 1022 e 1023 Novembro 2011
        '        If (frmT46V101A.objstcPedido.Cod_Produto >= 1022 And frmT46V101A.objstcPedido.Cod_Produto <= 1033) Or _
                 '           (frmT46V101A.objstcPedido.Cod_Produto >= 1038) Then  '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 8
        '            gM46V998.gLiberaDescontoEspecial = True
        '        Else
        '            gM46V998.gLiberaDescontoEspecial = False
        '        End If

        If gfCarCboBancos(2, frmT46V109A.cboCarBanco, lMensagem, txtCorretor(index).Tag) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        Else
'            Boleto santander
            For i = 0 To frmT46V109A.cboCarBanco.ListCount - 1
                '2016-08-22: Alterar de Itau para Santander
                If txtCorretor(index).Tag = 15701 Then 'mercedes
                    If frmT46V109A.cboCarBanco.ItemData(i) = 33 Then
                        frmT46V109A.cboCarBanco.ListIndex = i
                        Exit For
                    End If
                Else
                    If frmT46V109A.cboCarBanco.ItemData(i) = frmT46V101A.objstcPedido.Cod_Bco_Cobr Then
                        frmT46V109A.cboCarBanco.ListIndex = i
                        Exit For
                    End If
                End If
                '2016-08-22
            Next i
        End If

        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    'e) Se encontrou mais de um corretor, abre janela com grade de corretores.
    Load frmT46V113A
    With frmT46V113A.flxCorretores
        .Visible = True
        .Rows = lQuaCorretores + 1
        For lI = 1 To lQuaCorretores
            .Row = lI
            .Col = 0        'C�digo do corretor.
            .Text = lSaida(lI, 1)
            .Col = 1        'Nome do corretor.
            .Text = lSaida(lI, 2)
            .Col = 2        'Nome abreviado do corretor.
            .Text = lSaida(lI, 3)
        Next lI
        .Row = 1
        .Col = 1
    End With
    Screen.MousePointer = vbDefault
    frmT46V113A.Show vbModal, Me
    txtCorretor(index).Text = gNomCorretor
    txtCorretor(index).Tag = gCodCorretor
    If txtCorretor(index).Visible = True And txtCorretor(index).Enabled = True Then
        txtCorretor(index).SetFocus
    End If
    If index = 1 Then
        If gM46V999.gfPreenchido(txtCorretor(index).Tag) Then
            If gCodCorrInst <> txtCorretor(index).Tag And gCodCorrInst <> 0 And gM46V999.gCodProtocolo = 0 Then
                'Ficha 1000153 - Automatiza��o do agrupamento de corretores no Syas ap�s o cadastro realizado pelo help-desk.
                If gM46V999.gfVerCorrPermitido_TBS_Corr_Grupo(gCodCorrInst, txtCorretor(index).Tag, CInt(gM46V999.gCod_Ramo), lMensagem) = False Then
                    Call gM46V999.gpGraLog(1, "FGE0017 - C�digo corretor n�o permitido para este usu�rio (" & gM46V999.gNomUsuario & ")")
                    txtCorretor(index).Text = ""
                    If txtCorretor(index).Visible = True And txtCorretor(index).Enabled = True Then
                        txtCorretor(index).SetFocus
                    End If
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            If mfPosUniProProdutor(txtCorretor(1).Tag, lMensagem) = False Then
                Call gM46V999.gpGraLog(1, lMensagem)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub txtCorretor_Validate(index As Integer, Cancel As Boolean)
    If optProlabore(1).Value Or LTrim(RTrim(strCorretor)) = LTrim(RTrim(txtCorretor(index).Text)) Then
        Exit Sub
    End If

    Dim intlinhaGrid As Integer
    Dim blnExisteRegistro As Boolean
    blnExisteRegistro = False

    With frmT46V180A
        For intlinhaGrid = 1 To .flxProlaboristas.Rows - 1
            If LTrim(RTrim(txtCorretor(index).Text)) = LTrim(RTrim(.flxProlaboristas.TextMatrix(intlinhaGrid, 1))) Or LTrim(RTrim(strCorretor)) = LTrim(RTrim(.flxProlaboristas.TextMatrix(intlinhaGrid, 1))) Then
                blnExisteRegistro = True
                Exit For
            End If
        Next intlinhaGrid
    End With

    With frmT46V180A
        intlinhaGrid = 1
        Do While intlinhaGrid < .flxProlaboristas.Rows
            'refaz a numera��o das colunas no grid.
            .flxProlaboristas.TextMatrix(intlinhaGrid, 0) = Format(intlinhaGrid, "000")
            If LTrim(RTrim(strCorretor)) = LTrim(RTrim(.flxProlaboristas.TextMatrix(intlinhaGrid, 1))) Then
                If .flxProlaboristas.Rows = 2 Then

                    .flxProlaboristas.TextMatrix(intlinhaGrid, 0) = Format(intlinhaGrid, "000")
                    .flxProlaboristas.TextMatrix(intlinhaGrid, 1) = ""
                    .flxProlaboristas.TextMatrix(intlinhaGrid, 2) = ""
                    .flxProlaboristas.TextMatrix(intlinhaGrid, 3) = ""
                    .flxProlaboristas.TextMatrix(intlinhaGrid, 4) = ""
                    .flxProlaboristas.TextMatrix(intlinhaGrid, 5) = ""
                    'caso n�o reste mais prolaboristas, deve setar para n�o a op��o.
                    gCarga = True
                    optProlabore(1).Value = True
                    gCarga = False
                    Exit Do
                Else
                    .flxProlaboristas.RemoveItem (intlinhaGrid)
                    intlinhaGrid = 0
                End If
            End If
            intlinhaGrid = intlinhaGrid + 1
        Loop
    End With

End Sub
'Adicionado por Luiz Alberto em 11/12/2015
'*****************************************
Private Function VALIDATEXTO(X As String) As Boolean
Dim f As Integer
Dim G As Integer
Dim H As Integer

Dim Z As String * 6

Z = X

If Len(Z) < 6 Then
    VALIDATEXTO = False
Else
    'Recebe o primeiro caractere da string em c�digo asc.
    G = Asc(Left(Z, 1))
    H = 0
    'F recebe a segunda posi��o de Z.
    For f = 2 To Len(Z)
        'H soma o resultado do primeiro numero - o segundo numero e atribui a H.
        H = H + (G - Asc(Mid(Z, f, 1)))
    Next
    If H = 0 And Z = StrReverse(Z) Or H = -15 Then
        VALIDATEXTO = False
    Else
        VALIDATEXTO = True
    End If
End If
End Function

Private Sub txtCPF_Change()
    Call gpHabSalvar
End Sub
Private Sub txtCPF_GotFocus()
    With txtCPF
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub
Private Sub txtCPF_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtCPF_LostFocus()
    Dim lAuxiliar As String    'Auxiliar.

    Call gM46V999.gpForCPF(1, txtCPF.Text, lAuxiliar)
    If Len(lAuxiliar) = 14 And InStr(1, lAuxiliar, ".") = 0 Then
        txtCPF.Text = ""
    Else
        txtCPF.Text = lAuxiliar
    End If
    If cmdSiscad(0).Visible And gM46V999.PerfilAcesso <> e_Perfil_CALC And gM46V999.PerfilAcesso <> e_Perfil_CALC_Gestor Then
        Call cmdSiscad_Click(1)
    End If
End Sub

Private Sub txtDatEmis_Change()
    Call gpHabSalvar
End Sub

Private Sub txtDatEmis_GotFocus()
    With txtDatEmis
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub

Private Sub txtDatEmis_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDatEmis_LostFocus()
    txtDatEmis.Text = gM46V999.gfForData(9, txtDatEmis.Text)
    If IsDate(txtDatEmis.Text) = False Then
        txtDatEmis.Text = ""
    End If
End Sub

Private Sub txtFaturamento_Change()
'    Call gpHabSalvar
End Sub

Private Sub txtFaturamento_GotFocus()
'    Set mAntControle = ActiveControl
'    With txtFaturamento
'        .SelStart = 0
'        .SelLength = .MaxLength
'    End With
End Sub

Private Sub txtFaturamento_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        SendKeys "{TAB}"
'    End If
End Sub

Private Sub txtFaturamento_LostFocus()
'    If IsNumeric(txtFaturamento.Text) Then
'        txtFaturamento.Text = Format(txtFaturamento.Text, "standard")
'    Else
'        txtFaturamento.Text = ""
'    End If
End Sub

Private Sub txtIdeDatEmissao_Change()
    Call gpHabSalvar
End Sub
Private Sub txtIdeDatEmissao_GotFocus()
    With txtIdeDatEmissao
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub
Private Sub txtIdeDatEmissao_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, "/")
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtIdeDatEmissao_LostFocus()
'Verifica data.
    txtIdeDatEmissao.Text = gM46V999.gfForData(9, txtIdeDatEmissao.Text)
    
    'Adicionado por Luiz Alberto em 09/12/2015
    If Not IsDate(txtIdeDatEmissao.Text) Then
       MsgBox "Favor entrar com a data no formato: DD/MM/YYYY", vbInformation, "Aviso"
       txtIdeDatEmissao.Text = ""
       Exit Sub
    End If
    
    If CDate(txtIdeDatEmissao.Text) >= Date Then
       MsgBox "A data de emiss�o dever� ser menor que a data atual.", vbCritical, "Aten��o"
       txtIdeDatEmissao.Text = ""
       Exit Sub
    End If
    
'    If IsDate(txtIdeDatEmissao.Text) = False Then
'        txtIdeDatEmissao.Text = ""
'    End If
End Sub
Private Sub txtDataNascimento_Change()
    Call gpHabSalvar
End Sub
Private Sub txtDataNascimento_GotFocus()
    With txtDataNascimento
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub
Private Sub txtDataNascimento_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, "/")
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtDataNascimento_LostFocus()
'Verifica data.

    txtDataNascimento.Text = gM46V999.gfForData(9, txtDataNascimento.Text)
    
    'Adicionado por Luiz Alberto em 09/12/2015
    If Not IsDate(txtDataNascimento.Text) Then
       MsgBox "Favor entrar com a data no formato: DD/MM/YYYY", vbInformation, "Aviso"
       txtDataNascimento.Text = ""
       Exit Sub
    End If
    
    If CDate(txtDataNascimento.Text) >= Date Then
       MsgBox "A data de nascimento dever� ser menor que a data atual.", vbCritical, "Aten��o"
       txtDataNascimento.Text = ""
       Exit Sub
    End If

'    If IsDate(txtDataNascimento.Text) = False Then
'        txtDataNascimento.Text = ""
'    End If
    
End Sub
Private Sub txtIdeEmpresa_GotFocus()
    With txtIdeEmpresa
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub
Private Sub txtIdeempresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtIdeNumDocumento_Change()
    Call gpHabSalvar
End Sub
Private Sub txtIdeNumDocumento_GotFocus()
    With txtIdeNumDocumento
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub
Private Sub txtIdeNumDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtIdeOrgExpedidor_Change()
    Call gpHabSalvar
End Sub
Private Sub txtIdeOrgExpedidor_GotFocus()
    With txtIdeOrgExpedidor
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub
Private Sub txtIdeOrgExpedidor_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        SendKeys "{TAB}"
'    End If

    'Adicionado por Luiz Alberto em 09/12/2015
    'Transformar letras minusculas em Mai�sculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    ' teclas adicionais permitidas
    If KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 32 Then
       KeyAscii = KeyAscii
       Exit Sub
    End If
    
    ' Intercepta um c�digo ASCII recebido e admite somente letras
    If InStr("A��BC�DE��FGHI�JKLMNOPQRSTU�VWXYZ", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = KeyAscii
    End If
    
End Sub
Private Sub txtCorComissao_Change(index As Integer)
    Call gpLimPremios
    If gCarga = False Then
        gAltComum = True
    End If
End Sub
Private Sub txtCorComissao_GotFocus(index As Integer)
    On Error GoTo ControleErro
    With txtCorComissao(index)
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
    Exit Sub
ControleErro:
    Err.Clear
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = ActiveControl.MaxLength
    Set mAntControle = ActiveControl
    
End Sub
Private Sub txtCorComissao_KeyPress(index As Integer, KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtCorComissao_LostFocus(index As Integer)
    Dim lI As Byte             'Utilizado com For...Next.
    Dim lTotComissao As Double    'Total de comiss�o.

    If index < txtCorComissao.LBound Or index > txtCorComissao.UBound Then
        Load txtCorComissao(index)
    End If

'1. Procedimento para edi��o da comiss�o.
    With txtCorComissao(index)
        If gM46V999.gfPreenchido(.Text) Then
            If IsNumeric(.Text) Then
                .Text = Format$(CDbl(.Text), "standard")
            Else
                .Text = Format$("0,00", "standard")
            End If
        Else
            .Text = Format$("0,00", "standard")
        End If

        If index = 2 Or index = 3 Then
            If val(txtCorretor(index).Tag) = 0 Then
                .Text = "0,00"
            End If
        End If
        If CDbl(.Text) >= 100 Then
            .Text = 99.99
        End If
    End With

    '2. Verifica total de comiss�o.
    lTotComissao = 0
    For lI = 1 To 3
        If gM46V999.gfPreenchido(txtCorComissao(lI).Text) Then
            If IsNumeric(txtCorComissao(lI).Text) Then
                lTotComissao = lTotComissao + CDbl(txtCorComissao(lI).Text)
            End If
        End If
    Next lI

    If gM46V999.gTipoAcesso = DIGITACAO And lTotComissao = 0 Then
        frmT46V101A.txtComissao.Text = ""
    Else
        frmT46V101A.txtComissao.Text = Trim$(Format$(lTotComissao, "standard"))
    End If

    Screen.MousePointer = vbDefault


End Sub

Private Sub txtJuros_Change()
    Call gpHabSalvar
End Sub

Private Sub txtJuros_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")
End Sub


Private Sub txtJuros_LostFocus()
    Dim lI As Integer
    Dim lTip As Integer

    If gCarga = True Then Exit Sub
    txtJuros.Text = Format(gM46V999.CDblx(txtJuros.Text), "standard")
    If Not IsNumeric(txtJuros) Then
        'For lI = 1 To 12
        '    frmT46V142A.txtAntJuro(lI).Text = Format(3, "standard")
        '    frmT46V142A.txtPosJuro(lI).Text = Format(3, "standard")
        'Next

        'Simonyjuros
        'Trocar entrada

        'frmT46V142A.txtAntJuro(1).Text = Format(gM46O001.gProJurMinAnt01, "standard")
        'frmT46V142A.txtAntJuro(2).Text = Format(gM46O001.gProJurMinAnt02, "standard")
        'frmT46V142A.txtAntJuro(3).Text = Format(gM46O001.gProJurMinAnt03, "standard")
        'frmT46V142A.txtAntJuro(4).Text = Format(gM46O001.gProJurMinAnt04, "standard")
        'frmT46V142A.txtAntJuro(5).Text = Format(gM46O001.gProJurMinAnt05, "standard")
        'frmT46V142A.txtAntJuro(6).Text = Format(gM46O001.gProJurMinAnt06, "standard")
        'frmT46V142A.txtAntJuro(7).Text = Format(gM46O001.gProJurMinAnt07, "standard")
        'frmT46V142A.txtAntJuro(8).Text = Format(gM46O001.gProJurMinAnt08, "standard")
        'frmT46V142A.txtAntJuro(9).Text = Format(gM46O001.gProJurMinAnt09, "standard")
        'frmT46V142A.txtAntJuro(10).Text = Format(gM46O001.gProJurMinAnt10, "standard")
        'frmT46V142A.txtAntJuro(11).Text = Format(gM46O001.gProJurMinAnt11, "standard")
        'frmT46V142A.txtAntJuro(12).Text = Format(gM46O001.gProJurMinAnt12, "standard")
        'frmT46V142A.txtPosJuro(1).Text = Format(gM46O001.gProJurMinPos01, "standard")
        'frmT46V142A.txtPosJuro(2).Text = Format(gM46O001.gProJurMinPos02, "standard")
        'frmT46V142A.txtPosJuro(3).Text = Format(gM46O001.gProJurMinPos03, "standard")
        'frmT46V142A.txtPosJuro(4).Text = Format(gM46O001.gProJurMinPos04, "standard")
        'frmT46V142A.txtPosJuro(5).Text = Format(gM46O001.gProJurMinPos05, "standard")
        'frmT46V142A.txtPosJuro(6).Text = Format(gM46O001.gProJurMinPos06, "standard")
        'frmT46V142A.txtPosJuro(7).Text = Format(gM46O001.gProJurMinPos07, "standard")
        'frmT46V142A.txtPosJuro(8).Text = Format(gM46O001.gProJurMinPos08, "standard")
        'frmT46V142A.txtPosJuro(9).Text = Format(gM46O001.gProJurMinPos09, "standard")
        'frmT46V142A.txtPosJuro(10).Text = Format(gM46O001.gProJurMinPos10, "standard")
        'frmT46V142A.txtPosJuro(11).Text = Format(gM46O001.gProJurMinPos11, "standard")
        'frmT46V142A.txtPosJuro(12).Text = Format(gM46O001.gProJurMinPos12, "standard")

        If frmT46V102A.optForPagamento(0).Value = True Then lTip = e_FormaPagamento_Carne
        If frmT46V102A.optForPagamento(1).Value = True Then lTip = e_FormaPagamento_Debito
        
        If frmT46V101A.objstcPedido.Cod_Tip_Doc = e_TipDoc_Proposta And _
           frmT46V101A.objstcPedido.Cod_Bco_Cobr = 999 And _
           frmT46V101A.objstcPedido.Cod_Agen_Bas_Cor = 9999 And _
           frmT46V101A.objstcPedido.Num_Cta_Bas_Cor = 9999999 Then
            If frmT46V102A.optForPagamento(3).Value = True Then lTip = 4
        End If

        frmT46V142A.txtAntJuro(1).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinAnt01(lTip), "standard")
        frmT46V142A.txtAntJuro(2).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinAnt02(lTip), "standard")
        frmT46V142A.txtAntJuro(3).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinAnt03(lTip), "standard")
        frmT46V142A.txtAntJuro(4).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinAnt04(lTip), "standard")
        frmT46V142A.txtAntJuro(5).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinAnt05(lTip), "standard")
        frmT46V142A.txtAntJuro(6).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinAnt06(lTip), "standard")
        frmT46V142A.txtAntJuro(7).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinAnt07(lTip), "standard")
        frmT46V142A.txtAntJuro(8).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinAnt08(lTip), "standard")
        frmT46V142A.txtAntJuro(9).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinAnt09(lTip), "standard")
        frmT46V142A.txtAntJuro(10).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinAnt10(lTip), "standard")
        frmT46V142A.txtAntJuro(11).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinAnt11(lTip), "standard")
        frmT46V142A.txtAntJuro(12).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinAnt12(lTip), "standard")
        frmT46V142A.txtPosJuro(1).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinPos01(lTip), "standard")
        frmT46V142A.txtPosJuro(2).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinPos02(lTip), "standard")
        frmT46V142A.txtPosJuro(3).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinPos03(lTip), "standard")
        frmT46V142A.txtPosJuro(4).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinPos04(lTip), "standard")
        frmT46V142A.txtPosJuro(5).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinPos05(lTip), "standard")
        frmT46V142A.txtPosJuro(6).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinPos06(lTip), "standard")
        frmT46V142A.txtPosJuro(7).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinPos07(lTip), "standard")
        frmT46V142A.txtPosJuro(8).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinPos08(lTip), "standard")
        frmT46V142A.txtPosJuro(9).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinPos09(lTip), "standard")
        frmT46V142A.txtPosJuro(10).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinPos10(lTip), "standard")
        frmT46V142A.txtPosJuro(11).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinPos11(lTip), "standard")
        frmT46V142A.txtPosJuro(12).Text = Format(frmT46V101A.objstcPedido.getProduto.ProJurMinPos12(lTip), "standard")



        'Ficha 900046
        If frmT46V102A.optPagAntecipado(2).Value = True And IsNumeric(txtQuaParcelas) Then
            txtJuros = frmT46V142A.txtPosJuro(txtQuaParcelas).Text
        Else
            txtJuros = frmT46V142A.txtAntJuro(txtQuaParcelas).Text
        End If
    Else
        For lI = 1 To 12    '<- originalmente 10
            frmT46V142A.txtAntJuro(lI).Text = Format(gM46V999.CDblx(txtJuros), "standard")
            frmT46V142A.txtPosJuro(lI).Text = Format(gM46V999.CDblx(txtJuros), "standard")
        Next
    End If
    Call gpLimPremios
End Sub

Private Sub txtNomLogradouro_Change()
    mAltEndereco = True
    Call gpHabSalvar
    If optEndCobranca(1) = True Then    'Endere�o de cobran�a igual.
        frmT46V105A.txtNomLogradouro.Text = txtNomLogradouro.Text
    End If
End Sub
Private Sub txtNomLogradouro_GotFocus()
    With txtNomLogradouro
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub
Private Sub txtNomLogradouro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtNomLogradouro_LostFocus()
    If gM46V999.gfPreenchido(txtNomLogradouro.Text) Then
        Call mpVerEndSegurado
    End If
    txtNomLogradouro.Text = gM46V999.gfRetirarCaracterEspecial(txtNomLogradouro.Text, True)
End Sub

'Ficha 1000307
Private Sub txtnumregistro_Change()
    Call gpHabSalvar
    If gM46V999.gfPreenchido(txtnumregistro.Text) Then
        txtnumregistro.Text = Trim$(txtnumregistro.Text)
    End If
    txtnumregistro.Text = gM46V999.gfRetirarCaracterEspecial(txtnumregistro.Text, True)
End Sub

Private Sub txtNumSiscota_Change()
    If gCarga = True Then Exit Sub
    Call gpLimPremios
End Sub

Private Sub txtNumSiscota_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, "")
End Sub


Private Sub txtQuaParcelas_Change()
    If gCarga = True Then Exit Sub
    txtJuros.Text = Format(frmT46V142A.getJuros(IIf(optForPagamento(0).Value, e_FormaPagamento_Carne, e_FormaPagamento_Debito), optPagAntecipado(1).Value, val(txtQuaParcelas.Text)), "standard")
    Call gpLimPremios
End Sub
Private Sub txtQuaParcelas_GotFocus()
    With txtQuaParcelas
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub
Private Sub txtQuaParcelas_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtQuaParcelas_LostFocus()

    If gM46V999.gfPreenchido(txtQuaParcelas.Text) Then
        If IsNumeric(txtQuaParcelas.Text) Then
            If CInt(Me.txtQuaParcelas.Text) > frmT46V101A.objstcPedido.getProduto.ProQuaMaxParcelas And frmT46V101A.objstcPedido.Cod_Produto < 1086 Then
                Me.txtQuaParcelas.Text = frmT46V101A.objstcPedido.getProduto.ProQuaMaxParcelas
            End If
            If CInt(Me.txtQuaParcelas.Text) = 0 Then
                If gM46V111.gCod_Dizer = 0 Then
                    txtQuaParcelas.Text = frmT46V101A.objstcPedido.getProduto.ProQuaMaxParcelas
                End If
            End If
            
            'Nova tarifa
            If optPagAntecipado(1).Value = True And CDbl(txtQuaParcelas.Text) = 1 Then
                panProp(3).Enabled = False
                optVencimento(0).ForeColor = vbGrayText
                optVencimento(1).ForeColor = vbGrayText
                cmdAjustarDatas.Enabled = False
                optVencimento(1).Value = True
            Else
                panProp(3).Enabled = True
                optVencimento(0).ForeColor = 8388608
                optVencimento(1).ForeColor = 8388608
                cmdAjustarDatas.Enabled = True
            End If
        Else
            txtQuaParcelas.Text = frmT46V101A.objstcPedido.getProduto.ProQuaMaxParcelas
        End If
    Else
        If gM46V999.gTipoAcesso <> DIGITACAO Then
            If gM46V111.gCod_Dizer = 0 Then
                txtQuaParcelas.Text = frmT46V101A.objstcPedido.getProduto.ProQuaMaxParcelas
            Else
                txtQuaParcelas.Text = 0
            End If
        End If
    End If

End Sub

Private Sub txtReferencia_Change()
    If gCarga = True Then Exit Sub
    Call gpLimPremios
End Sub

Private Sub txtRendaMensal_Change()
'    Call gpHabSalvar
End Sub

Private Sub txtRendaMensal_GotFocus()
' Set mAntControle = ActiveControl
'    With txtRendaMensal
'        .SelStart = 0
'        .SelLength = .MaxLength
'    End With
End Sub

Private Sub txtRendaMensal_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        SendKeys "{TAB}"
'    End If
End Sub

Private Sub txtRendaMensal_LostFocus()
'    If IsNumeric(txtRendaMensal.Text) Then
'        txtRendaMensal.Text = Format(txtRendaMensal.Text, "standard")
'    Else
'        txtRendaMensal.Text = ""
'    End If
End Sub

Private Sub txtSegurado_Change()
    Call gpHabSalvar
    frmT46V101A.txtSegurado.Text = txtSegurado.Text

    'Proposta Eletr�nica
    If gM46V999.gNovaTransmissaoMulti = True Then
        If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
            If frmT46V101A.optTipPessoa(0).Value = True Then
                Call gM46V999.gpForCPF(2, txtCnpj.Text, mAtuCNPJ)  'F�sica
            Else
                Call gM46V999.gpForCNPJ(2, txtCnpj.Text, mAtuCNPJ)    'Jur�dica
            End If
            If mAtuCNPJ <> mAntCNPJ Then
                frmT46V101A.txtSegurado.Enabled = True
                txtSegurado.Enabled = True
            Else
                frmT46V101A.txtSegurado.Text = mAntSeg
                frmT46V101A.txtSegurado.Enabled = False
                txtSegurado.Text = mAntSeg
                txtSegurado.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub txtSegurado_GotFocus()
    With txtSegurado
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub

Private Sub txtSegurado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSegurado_LostFocus()
    txtSegurado.Text = gM46V999.gfRetirarCaracterEspecial(txtSegurado.Text, True)
End Sub

Private Sub txtTelefones_Change()
    Call gpHabSalvar
    If optEndCobranca(1) = True Then    'Endere�o de cobran�a igual.
        frmT46V105A.txtTelefones.Text = txtTelefones.Text
    End If
End Sub

Private Sub txtTelefones_GotFocus()
    With txtTelefones
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub

Private Sub txtTelefones_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtTelefones_LostFocus()
    If gM46V999.gfPreenchido(txtTelefones.Text) Then
        txtTelefones.Text = Trim$(txtTelefones.Text)
    End If
    txtTelefones.Text = gM46V999.gfRetirarCaracterEspecial(txtTelefones.Text, True)
End Sub

Private Sub txtTipLogradouro_Change()
    mAltEndereco = True
    txtTipLogradouro.Text = gM46V999.gfRetirarCaracterEspecial(txtTipLogradouro.Text, True)
    Call gpHabSalvar
    If optEndCobranca(1) = True Then    'Endere�o de cobran�a igual.
        frmT46V105A.txtTipLogradouro.Text = txtTipLogradouro.Text
    End If
End Sub

Private Sub txtTipLogradouro_GotFocus()
    With txtTipLogradouro
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub

Private Sub txtTipLogradouro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtTipLogradouro_LostFocus()
    If gM46V999.gfPreenchido(txtTipLogradouro.Text) Then
        Call mpVerEndSegurado
    End If
End Sub

Private Sub txtValorPgAnt_Change()
    If gCarga = True Then Exit Sub
    Call gpLimPremios
End Sub

Private Sub txtValorPgAnt_GotFocus()
    With txtValorPgAnt
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    Set mAntControle = ActiveControl
End Sub


Private Sub txtValorPgAnt_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtValorPgAnt_LostFocus()
    With txtValorPgAnt
        If gM46V999.gfPreenchido(.Text) Then
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "standard")
            Else
                .Text = Format(0, "standard")
            End If
        Else
            .Text = Format(0, "standard")
        End If
    End With
End Sub

'Ficha 1000307
Public Sub mpCarregarRegConselho()
'Procedure: obt�m Reg.Conselho.
    Dim lbdTAB_RegCons As ADODB.Connection    'Base de dados - TBM_REG_PROFISSOES.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lMensagem As String    'Mensagem.
    Dim lrsTAB_RegCons As clsYasRecordSet    'Registro - TBM_REG_PROFISSOES.
    Dim lCod_Clasf As String  'C�digo da Classifica��o
    Dim lCod_Compl_Clasf As String        'Complemento da Classifica��o
    Dim lSelect As String
    If gM46V999.gCod_Ramo = 0 Then Exit Sub

    gObrRegistro = False

    lCod_Clasf = Left(frmT46V101A.txtClassificacao.Tag, 4)
    lCod_Compl_Clasf = Mid(frmT46V101A.txtClassificacao.Tag, 5, 2)

    '1. Abre arquivo.
    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_RegCons, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    '2. Obt�m Reg.Conselho informado.
    lSelect = "SELECT DISTINCT a.Cod_Reg, a.Dsc_Reg FROM TBM_REG_PROFISSOES A, TBM_CLASF_PROFISSOES B "
    lSelect = lSelect & " WHERE a.Cod_Reg = b.Cod_Reg"
    lSelect = lSelect & " and b.Cod_Ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " and b.Cod_Modalidade = " & gM46V111.gCod_plano
    lSelect = lSelect & " and b.Cod_Clasf = '" & Trim(lCod_Clasf) & "'"    '1000871 - Inclus�o de Org�o Categoria Conselho
    lSelect = lSelect & " and b.Cod_Compl_Clasf = '" & Trim(lCod_Compl_Clasf) & "'"    '1000871 - Inclus�o de Org�o Categoria Conselho
    lSelect = lSelect & " order by a.Cod_Reg"


    If gM46V999.gfObtRegistro(lbdTAB_RegCons, lSelect, lrsTAB_RegCons, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_RegCons)
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If Not lrsTAB_RegCons.EOF Then gObrRegistro = True

    cboRegCons.Clear
    Do While Not lrsTAB_RegCons.EOF
        If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
            cboRegCons.AddItem lrsTAB_RegCons("Dsc_Reg")
            cboRegCons.ItemData(cboRegCons.NewIndex) = lrsTAB_RegCons("Cod_Reg")
        End If
        lrsTAB_RegCons.MoveNext
    Loop

    Call gM46V999.gpFecha2(lbdTAB_RegCons)

End Sub

Private Sub mpPosReg()
    Dim i As Integer
    cboRegCons.ListIndex = -1
    For i = 0 To cboRegCons.ListCount - 1
        If cboRegCons.ItemData(i) = cboRegCons.Tag Then
            cboRegCons.ListIndex = i
            Exit For
        End If
    Next
End Sub

'Fichas 1200168 e 1200169 Inclus�o dos campos de listagem Renda e Faixa de Faturamento Pessoa f�sica e jur�dica
Private Sub cboFaturamento_Click()
    Call gpHabSalvar
End Sub

'Fichas 1200168 e 1200169 Inclus�o dos campos de listagem Renda e Faixa de Faturamento Pessoa f�sica e jur�dica
Private Sub cboRendaMensal_Click()
    Call gpHabSalvar
End Sub

Private Sub mpIniciaGridTelefones(Linha As Integer)
    With flxTelefones
        .Row = Linha
        .TextMatrix(.Row, 0) = ""
        .TextMatrix(.Row, 1) = ""
        .TextMatrix(.Row, 2) = ""
        .TextMatrix(.Row, 3) = ""
        .TextMatrix(.Row, 4) = ""
        .TextMatrix(.Row, 5) = ""
        .TextMatrix(.Row, 6) = ""
        .TextMatrix(.Row, 7) = ""
        .TextMatrix(.Row, 8) = ""

        .Col = 0
        Set .CellPicture = picChecked

        'Resize da coluna de observa��o
        .Col = 5
        If .Rows > 5 Then
            .ColWidth(.Col) = 3220    'com scrollbar
        Else
            .ColWidth(.Col) = 3480    'sem scrollbar
        End If

        .Col = 6
        .ColWidth(.Col) = picIncluir.Width
        Set .CellPicture = picIncluir
        .Refresh

        .Col = 7
        .ColWidth(.Col) = picExcluir.Width
        Set .CellPicture = picExcluir

        .Refresh
        DoEvents
    End With
End Sub

Private Sub LimpaGrid()
    With flxTelefones
        .FixedRows = 1
        .FixedCols = 0
        .Rows = 2
        .Cols = 9
        .TextMatrix(0, 0) = "Principal"
        .TextMatrix(0, 1) = "Tipo"
        .TextMatrix(0, 2) = "DDD"
        .TextMatrix(0, 3) = "Telefone"
        .TextMatrix(0, 4) = "Ramal"
        .TextMatrix(0, 5) = "Observa��o"
        .TextMatrix(0, 6) = Empty
        .TextMatrix(0, 7) = Empty
        .TextMatrix(0, 8) = Empty
        .ColWidth(8) = 0
        .RowHeight(1) = cboTipTelefone.Height
        Call mpIniciaGridTelefones(1)
    End With
End Sub

Private Sub cmdRestituicao_Click()
    Call frmT46V187A.mpInstanciarForm(2)
    If Not gM46V999.gDesabilitarForms Then
        If Not mAntControle Is Nothing Then
            If mAntControle.Visible = True And mAntControle.Enabled = True Then
                mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
            End If
        End If
    End If
End Sub

Private Sub mpHabControles()
    Me.AutoRedraw = False

'    If gM46V999.gCodProduto > 9999 Then
        panProp(6).Visible = True
        panProp(8).Top = panProp(6).Top + panProp(6).Height + 15
'    Else
'        panProp(6).Visible = False
'        panProp(8).Top = panProp(10).Top + panProp(10).Height + 15
'    End If
    
    If gM46V999.gfPreenchido(txtQuaParcelas.Text) Then
        If optPagAntecipado(1).Value = True And txtQuaParcelas.Text = 1 Then
            panProp(3).Enabled = False
            optVencimento(0).ForeColor = vbGrayText
            optVencimento(1).ForeColor = vbGrayText
        Else
            panProp(3).Enabled = True
            optVencimento(0).ForeColor = 8388608
            optVencimento(1).ForeColor = 8388608
        End If
    End If

    ''2016-05-09:Juliano-Bloq. se endosso
    If (frmT46V101A.isEndosso = True) Then
        txtSegurado.Enabled = False
        optTipPessoa(0).Enabled = False
        optTipPessoa(1).Enabled = False
        txtCPF.Enabled = False
        txtCnpj.Enabled = False
    End If
    
    panProp(3).Top = panProp(8).Top + panProp(8).Height + 15
    panProduto.Height = panProp(3).Top + panProp(3).Height + 30
    Me.Height = panProduto.Top + panProduto.Height + 500
    cmdVoltar.Top = Me.Height - cmdVoltar.Height - 500
    Me.AutoRedraw = True
End Sub
