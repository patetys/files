VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT47V004A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danos de proposta"
   ClientHeight    =   5880
   ClientLeft      =   8040
   ClientTop       =   2100
   ClientWidth     =   9000
   Icon            =   "T47V004A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSegurado 
      ForeColor       =   &H00800000&
      Height          =   285
      HelpContextID   =   44
      Left            =   1230
      MaxLength       =   60
      TabIndex        =   0
      ToolTipText     =   "Nome do segurado."
      Top             =   0
      Width           =   5205
   End
   Begin VB.TextBox txtCodSegurado 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1230
      MaxLength       =   6
      TabIndex        =   3
      ToolTipText     =   "C�digo do segurado."
      Top             =   285
      Width           =   1455
   End
   Begin VB.TextBox txtNumSiscota 
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   7215
      MaxLength       =   12
      TabIndex        =   8
      ToolTipText     =   "N�mero da Cota��o (SisCota)"
      Top             =   285
      Width           =   1770
   End
   Begin VB.ComboBox CmbTipCotac 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "T47V004A.frx":000C
      Left            =   5625
      List            =   "T47V004A.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   7
      ToolTipText     =   "Tipo de Cota��o (se houver)"
      Top             =   285
      Width           =   1575
   End
   Begin VB.TextBox txtReferencia 
      ForeColor       =   &H00800000&
      Height          =   300
      HelpContextID   =   40
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   50
      ToolTipText     =   "Refer�ncia do Corretor"
      Top             =   5040
      Width           =   1515
   End
   Begin VB.TextBox txtDigCodSegurado 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   4
      ToolTipText     =   "C�digo do segurado."
      Top             =   285
      Width           =   300
   End
   Begin VB.TextBox txtDatEmis 
      ForeColor       =   &H00800000&
      Height          =   300
      HelpContextID   =   40
      Left            =   4125
      MaxLength       =   10
      TabIndex        =   52
      ToolTipText     =   "Data agendada para emiss�o"
      Top             =   5040
      Width           =   1140
   End
   Begin Threed.SSPanel panProduto 
      Height          =   3060
      Left            =   5400
      TabIndex        =   51
      ToolTipText     =   "Dados e op��es do produto."
      Top             =   2805
      Width           =   3600
      _Version        =   65536
      _ExtentX        =   6350
      _ExtentY        =   5397
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.TextBox txtQuaParcelas 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   300
         HelpContextID   =   40
         Left            =   2750
         MaxLength       =   3
         TabIndex        =   30
         ToolTipText     =   "Quantidade de parcelas."
         Top             =   30
         Width           =   800
      End
      Begin VB.TextBox txtJuros 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   300
         HelpContextID   =   40
         Left            =   2750
         MaxLength       =   5
         TabIndex        =   31
         ToolTipText     =   "Juros."
         Top             =   330
         Width           =   800
      End
      Begin Threed.SSPanel panProp 
         Height          =   945
         Index           =   8
         Left            =   45
         TabIndex        =   54
         ToolTipText     =   "Tem pagamento antecipado?"
         Top             =   1365
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
         Begin VB.OptionButton optPagAntecipado 
            Caption         =   "Antecipada"
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   1
            Left            =   1200
            TabIndex        =   36
            ToolTipText     =   "SIM, tem pagamento antecipado."
            Top             =   75
            Width           =   1900
         End
         Begin VB.OptionButton optPagAntecipado 
            Caption         =   "A 30 dias"
            ForeColor       =   &H00800000&
            Height          =   230
            Index           =   2
            Left            =   1200
            TabIndex        =   38
            ToolTipText     =   "N�O tem pagamento antecipado."
            Top             =   380
            Width           =   1900
         End
         Begin VB.TextBox txtValorPgAnt 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2280
            MaxLength       =   14
            TabIndex        =   40
            ToolTipText     =   "Valor Pagamento Antecipado"
            Top             =   600
            Width           =   1185
         End
         Begin Threed.SSCommand cmdPagAntBloqueto 
            Height          =   240
            Left            =   3200
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "Exibir informa��es de pagamento antecipado."
            Top             =   70
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T47V004A.frx":0010
         End
         Begin Threed.SSPanel panProp 
            Height          =   840
            Index           =   9
            Left            =   45
            TabIndex        =   55
            Top             =   45
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   1482
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
            TabIndex        =   39
            TabStop         =   0   'False
            ToolTipText     =   "Exibir informa��es de pagamento antecipado."
            Top             =   350
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T47V004A.frx":032A
         End
         Begin VB.Label lblPagAnt 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "R$"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2040
            TabIndex        =   56
            Top             =   630
            Visible         =   0   'False
            Width           =   210
         End
      End
      Begin Threed.SSPanel panProp 
         Height          =   705
         Index           =   10
         Left            =   45
         TabIndex        =   57
         ToolTipText     =   "Forma de pagamento."
         Top             =   630
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6174
         _ExtentY        =   1235
         _StockProps     =   15
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.19
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
            Caption         =   "D�bito em c/c"
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   1
            Left            =   1200
            TabIndex        =   34
            ToolTipText     =   "Pagamento atrav�s de d�bito em conta corrente."
            Top             =   350
            Width           =   1400
         End
         Begin VB.OptionButton optForPagamento 
            Caption         =   "Carn�"
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   0
            Left            =   1200
            TabIndex        =   32
            ToolTipText     =   "Pagamento atrav�s de carn�."
            Top             =   120
            Width           =   750
         End
         Begin Threed.SSCommand cmdForPagCarne 
            Height          =   240
            Left            =   3200
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Exibir informa��es de forma de pagamento."
            Top             =   100
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T47V004A.frx":0644
         End
         Begin Threed.SSPanel panProp 
            Height          =   610
            Index           =   11
            Left            =   50
            TabIndex        =   58
            Top             =   50
            Width           =   1100
            _Version        =   65536
            _ExtentX        =   1940
            _ExtentY        =   1076
            _StockProps     =   15
            Caption         =   "Forma de pagamento"
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.18
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
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Exibir informa��es de forma de pagamento."
            Top             =   350
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T47V004A.frx":095E
         End
      End
      Begin Threed.SSPanel panProp 
         Height          =   705
         Index           =   3
         Left            =   45
         TabIndex        =   59
         ToolTipText     =   "Forma de pagamento."
         Top             =   2290
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6174
         _ExtentY        =   1235
         _StockProps     =   15
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.19
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
            Caption         =   "Sim"
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   0
            Left            =   1200
            TabIndex        =   41
            ToolTipText     =   "Data das parcelas ajustadas."
            Top             =   120
            Width           =   750
         End
         Begin VB.OptionButton optVencimento 
            Caption         =   "N�o"
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   1
            Left            =   1200
            TabIndex        =   43
            ToolTipText     =   "N�o ajustar a data das parcelas"
            Top             =   350
            Width           =   1400
         End
         Begin Threed.SSCommand cmdAjustarDatas 
            Height          =   240
            Left            =   3200
            TabIndex        =   42
            TabStop         =   0   'False
            ToolTipText     =   "Exibir informa��es de forma de pagamento."
            Top             =   100
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T47V004A.frx":0C78
         End
         Begin Threed.SSPanel panProp 
            Height          =   610
            Index           =   4
            Left            =   50
            TabIndex        =   60
            Top             =   50
            Width           =   1100
            _Version        =   65536
            _ExtentX        =   1940
            _ExtentY        =   1076
            _StockProps     =   15
            Caption         =   "Ajustar data das parcelas"
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.18
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
         Caption         =   "Quantidade de parcelas: "
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   24
         Left            =   50
         TabIndex        =   62
         ToolTipText     =   "Quantidade de parcelas"
         Top             =   30
         Width           =   2700
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Juros das parcelas: "
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   16
         Left            =   50
         TabIndex        =   61
         ToolTipText     =   "Juros das parcelas"
         Top             =   330
         Width           =   2700
      End
   End
   Begin Threed.SSPanel panCorretores 
      Height          =   1245
      Left            =   0
      TabIndex        =   63
      Top             =   2805
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9349
      _ExtentY        =   2205
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.28
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.TextBox txtCorComissao 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   4550
         MaxLength       =   5
         TabIndex        =   25
         Text            =   "0,00"
         ToolTipText     =   "Comiss�o do primeiro corretor."
         Top             =   320
         Width           =   700
      End
      Begin VB.TextBox txtCorComissao 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   4550
         MaxLength       =   5
         TabIndex        =   27
         Text            =   "0,00"
         ToolTipText     =   "Comiss�o do segundo corretor."
         Top             =   600
         Width           =   700
      End
      Begin VB.TextBox txtCorComissao 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   4550
         MaxLength       =   5
         TabIndex        =   29
         Text            =   "0,00"
         ToolTipText     =   "Comiss�o do terceiro corretor."
         Top             =   890
         Width           =   700
      End
      Begin VB.TextBox txtCorretor 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   50
         TabIndex        =   24
         ToolTipText     =   "Primeiro corretor."
         Top             =   320
         Width           =   4500
      End
      Begin VB.TextBox txtCorretor 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   50
         TabIndex        =   26
         ToolTipText     =   "Segundo corretor."
         Top             =   600
         Width           =   4500
      End
      Begin VB.TextBox txtCorretor 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   50
         TabIndex        =   28
         ToolTipText     =   "Terceiro corretor."
         Top             =   890
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
         TabIndex        =   65
         Top             =   30
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
         TabIndex        =   64
         Top             =   30
         Width           =   4500
      End
   End
   Begin Threed.SSPanel panEndereco 
      Height          =   1245
      Left            =   0
      TabIndex        =   66
      ToolTipText     =   "Dados do segurado."
      Top             =   1565
      Width           =   9000
      _Version        =   65536
      _ExtentX        =   15875
      _ExtentY        =   2205
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.28
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.TextBox TxtComplCobr 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5715
         MaxLength       =   20
         TabIndex        =   117
         ToolTipText     =   "Complemento do endere�o do segurado (n�mero, andar, apartamento, conjunto, etc.)."
         Top             =   870
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtNumCobr 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5115
         MaxLength       =   20
         TabIndex        =   116
         ToolTipText     =   "Complemento do endere�o do segurado (n�mero, andar, apartamento, conjunto, etc.)."
         Top             =   870
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtComplResid 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4590
         MaxLength       =   20
         TabIndex        =   115
         ToolTipText     =   "Complemento do endere�o do segurado (n�mero, andar, apartamento, conjunto, etc.)."
         Top             =   855
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtNumResid 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4005
         MaxLength       =   20
         TabIndex        =   114
         ToolTipText     =   "Complemento do endere�o do segurado (n�mero, andar, apartamento, conjunto, etc.)."
         Top             =   870
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtNomLogradouro 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3945
         MaxLength       =   50
         TabIndex        =   15
         ToolTipText     =   "Nome do logradouro do endere�o do segurado."
         Top             =   30
         Width           =   5010
      End
      Begin VB.TextBox txtCEP 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   13
         ToolTipText     =   "CEP do endere�o do segurado (nnnnn-nnn)."
         Top             =   30
         Width           =   900
      End
      Begin VB.TextBox txtTipLogradouro 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3345
         MaxLength       =   3
         TabIndex        =   14
         ToolTipText     =   "Tipo do logradouro do endere�o do segurado (Rua, Av., etc.)."
         Top             =   30
         Width           =   600
      End
      Begin VB.TextBox txtComplemento 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   16
         ToolTipText     =   "Complemento do endere�o do segurado (n�mero, andar, apartamento, conjunto, etc.)."
         Top             =   320
         Width           =   2500
      End
      Begin VB.TextBox txtBairro 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5040
         MaxLength       =   30
         TabIndex        =   17
         ToolTipText     =   "Bairro do endere�o do segurado."
         Top             =   320
         Width           =   3900
      End
      Begin VB.ComboBox cboUF 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T47V004A.frx":0F92
         Left            =   8130
         List            =   "T47V004A.frx":0FE7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         ToolTipText     =   "Unidade da federa��o do endere�o do segurado."
         Top             =   600
         Width           =   810
      End
      Begin VB.TextBox txtCidade 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   18
         ToolTipText     =   "Cidade do endere�o do segurado."
         Top             =   600
         Width           =   2500
      End
      Begin VB.TextBox txtTelefones 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5055
         MaxLength       =   7
         TabIndex        =   19
         ToolTipText     =   "Telefone(s) do endere�o do segurado."
         Top             =   600
         Width           =   2310
      End
      Begin Threed.SSPanel panProp 
         Height          =   285
         Index           =   1
         Left            =   2310
         TabIndex        =   67
         ToolTipText     =   "Endere�o de cobran�a � diferente do endere�o do segurado?"
         Top             =   885
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
            Caption         =   "&Sim"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   45
            TabIndex        =   21
            ToolTipText     =   "SIM, endere�o de cobran�a � diferente do endere�o do segurado."
            Top             =   50
            Width           =   600
         End
         Begin VB.OptionButton optEndCobranca 
            Caption         =   "&N�o"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   645
            TabIndex        =   22
            ToolTipText     =   "N�O, endere�o de cobran�a � igual ao endere�o do segurado."
            Top             =   50
            Width           =   600
         End
         Begin Threed.SSCommand cmdEndCobranca 
            Height          =   255
            Left            =   1275
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Exibir endere�o de cobran�a."
            Top             =   15
            Width           =   255
            _Version        =   65536
            _ExtentX        =   441
            _ExtentY        =   441
            _StockProps     =   78
            ForeColor       =   12582912
            Picture         =   "T47V004A.frx":1057
         End
      End
      Begin Threed.SSCommand cmdCEP 
         Height          =   250
         HelpContextID   =   42
         Left            =   50
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Busca endere�o do segurado."
         Top             =   50
         Width           =   250
         _Version        =   65536
         _ExtentX        =   441
         _ExtentY        =   441
         _StockProps     =   78
         ForeColor       =   12582912
         Picture         =   "T47V004A.frx":1371
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CEP: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   7
         Left            =   300
         TabIndex        =   75
         ToolTipText     =   "CEP do endere�o do segurado."
         Top             =   30
         Width           =   1140
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Logradouro: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   2340
         TabIndex        =   74
         ToolTipText     =   "Logradouro do endere�o do segurado."
         Top             =   30
         Width           =   1005
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N�/Complemento: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   45
         TabIndex        =   73
         ToolTipText     =   "Complemento do endere�o do segurado."
         Top             =   315
         Width           =   1395
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bairro:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   3945
         TabIndex        =   72
         ToolTipText     =   "Bairro do endere�o do segurado."
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cidade: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   5
         Left            =   45
         TabIndex        =   71
         ToolTipText     =   "Cidade do endere�o do segurado."
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UF:"
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   6
         Left            =   7365
         TabIndex        =   70
         ToolTipText     =   "Unidade da federa��o do endere�o do segurado."
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefones:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   8
         Left            =   3945
         TabIndex        =   69
         ToolTipText     =   "Telefone(s) do endere�o do segurado."
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Endere�o cobran�a diferente?"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   9
         Left            =   45
         TabIndex        =   68
         ToolTipText     =   "Endere�o de cobran�a � diferente do endere�o do segurado?"
         Top             =   885
         Width           =   2250
      End
   End
   Begin Threed.SSPanel panProp 
      Height          =   945
      Index           =   0
      Left            =   0
      TabIndex        =   76
      Top             =   4080
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   1667
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin Threed.SSCommand cmdUniProdutiva 
         Height          =   285
         HelpContextID   =   42
         Left            =   4950
         TabIndex        =   45
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
         Picture         =   "T47V004A.frx":168B
      End
      Begin Threed.SSCommand cmdProdutor 
         Height          =   285
         HelpContextID   =   42
         Left            =   4950
         TabIndex        =   47
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
         Picture         =   "T47V004A.frx":19A5
      End
      Begin Threed.SSCommand cmdCosseguro 
         Height          =   285
         Left            =   4950
         TabIndex        =   77
         TabStop         =   0   'False
         ToolTipText     =   "Exibir informa��es de cosseguro."
         Top             =   600
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   503
         _StockProps     =   78
         ForeColor       =   12582912
         Picture         =   "T47V004A.frx":1CBF
      End
      Begin Threed.SSPanel panProp 
         Height          =   285
         Index           =   5
         Left            =   1900
         TabIndex        =   78
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
            Caption         =   "&Sim"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   50
            TabIndex        =   48
            ToolTipText     =   "SIM, tem cess�o de cosseguro."
            Top             =   50
            Width           =   600
         End
         Begin VB.OptionButton optCosseguro 
            Caption         =   "&N�o"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   650
            TabIndex        =   49
            ToolTipText     =   "N�O tem cess�o de cosseguro."
            Top             =   50
            Width           =   600
         End
      End
      Begin VB.Label lblNomUnidade 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   950
         TabIndex        =   44
         ToolTipText     =   "Nome da unidade produtiva."
         Top             =   30
         Width           =   4000
      End
      Begin VB.Label lblUnidade 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidade: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   45
         TabIndex        =   81
         ToolTipText     =   "Unidade produtiva."
         Top             =   45
         Width           =   900
      End
      Begin VB.Label lblProdutor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produtor: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   50
         TabIndex        =   80
         Top             =   330
         Width           =   900
      End
      Begin VB.Label lblNomProdutor 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   950
         TabIndex        =   46
         ToolTipText     =   "Nome do produtor."
         Top             =   330
         Width           =   4000
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cess�o de cosseguro? "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   12
         Left            =   50
         TabIndex        =   79
         ToolTipText     =   "Cess�o de cosseguro?"
         Top             =   600
         Width           =   1850
      End
   End
   Begin Threed.SSPanel panProp 
      Height          =   285
      Index           =   15
      Left            =   7250
      TabIndex        =   82
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
         Caption         =   "&Jur�dica"
         ForeColor       =   &H00800000&
         Height          =   200
         Index           =   1
         Left            =   800
         TabIndex        =   2
         ToolTipText     =   "Pessoa jur�dica."
         Top             =   50
         Width           =   900
      End
      Begin VB.OptionButton optTipPessoa 
         Caption         =   "&F�sica"
         ForeColor       =   &H00800000&
         Height          =   200
         Index           =   0
         Left            =   50
         TabIndex        =   1
         ToolTipText     =   "Pessoa f�sica."
         Top             =   50
         Value           =   -1  'True
         Width           =   750
      End
   End
   Begin Threed.SSCommand cmdExtra 
      Height          =   240
      Index           =   2
      Left            =   210
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   4995
      Visible         =   0   'False
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   78
      ForeColor       =   12582912
      Picture         =   "T47V004A.frx":1FD9
   End
   Begin Threed.SSCommand cmdExtra 
      Height          =   240
      Index           =   1
      Left            =   0
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   4995
      Visible         =   0   'False
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   78
      ForeColor       =   12582912
      Picture         =   "T47V004A.frx":22F3
   End
   Begin Threed.SSCommand cmdSiscad 
      Height          =   300
      Index           =   0
      Left            =   3015
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Buscar dados SISCAD"
      Top             =   285
      Width           =   360
      _Version        =   65536
      _ExtentX        =   635
      _ExtentY        =   529
      _StockProps     =   78
      ForeColor       =   12582912
      Picture         =   "T47V004A.frx":260D
   End
   Begin Threed.SSCommand cmdSiscad 
      Height          =   300
      Index           =   1
      Left            =   3390
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Atualizar dados do segurado."
      Top             =   285
      Width           =   360
      _Version        =   65536
      _ExtentX        =   635
      _ExtentY        =   529
      _StockProps     =   78
      ForeColor       =   12582912
      Picture         =   "T47V004A.frx":2927
   End
   Begin Threed.SSCommand cmdVoltar 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   45
      TabIndex        =   53
      ToolTipText     =   "Volta ao item."
      Top             =   5520
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Voltar ao item"
      ForeColor       =   8388608
   End
   Begin Threed.SSPanel panPesJuridica 
      Height          =   945
      Left            =   0
      TabIndex        =   85
      Top             =   600
      Width           =   9000
      _Version        =   65536
      _ExtentX        =   15875
      _ExtentY        =   1676
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.3
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.ComboBox cboRamAtividade 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T47V004A.frx":2A81
         Left            =   3400
         List            =   "T47V004A.frx":2A83
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Ramo de atividade do segurado."
         Top             =   150
         Width           =   5500
      End
      Begin VB.TextBox txtCNPJ 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   50
         MaxLength       =   18
         TabIndex        =   9
         ToolTipText     =   "CNPJ do segurado."
         Top             =   450
         Width           =   1600
      End
      Begin VB.TextBox txtFaturamento 
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   3400
         MaxLength       =   18
         TabIndex        =   11
         ToolTipText     =   "Valor do faturamento."
         Top             =   480
         Width           =   1770
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ramo de atividade: "
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   88
         ToolTipText     =   "Ramo de atividade do segurado."
         Top             =   150
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
         TabIndex        =   87
         ToolTipText     =   "CNPJ do segurado."
         Top             =   150
         Width           =   1600
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Faturamento: "
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   25
         Left            =   1800
         TabIndex        =   86
         ToolTipText     =   "Faturamento"
         Top             =   480
         Width           =   1605
      End
   End
   Begin Threed.SSPanel panPesFisica 
      Height          =   945
      Left            =   0
      TabIndex        =   89
      Top             =   600
      Width           =   9000
      _Version        =   65536
      _ExtentX        =   15875
      _ExtentY        =   1676
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.3
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.TextBox txtCPF 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   50
         MaxLength       =   14
         TabIndex        =   96
         ToolTipText     =   "CPF do segurado."
         Top             =   450
         Width           =   1250
      End
      Begin VB.ComboBox cboIdeTipDocumento 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T47V004A.frx":2A85
         Left            =   2985
         List            =   "T47V004A.frx":2A87
         Style           =   2  'Dropdown List
         TabIndex        =   95
         ToolTipText     =   "Tipo do documento de identifica��o do segurado."
         Top             =   50
         Width           =   3100
      End
      Begin VB.TextBox txtIdeNumDocumento 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2985
         MaxLength       =   20
         TabIndex        =   94
         ToolTipText     =   "N�mero do documento de identifica��o do segurado."
         Top             =   350
         Width           =   1605
      End
      Begin VB.TextBox txtIdeOrgExpedidor 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5910
         MaxLength       =   20
         TabIndex        =   93
         ToolTipText     =   "�rg�o expedidor do documento de identifica��o do segurado."
         Top             =   345
         Width           =   1125
      End
      Begin VB.TextBox txtIdeDatEmissao 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7485
         MaxLength       =   10
         TabIndex        =   92
         ToolTipText     =   "Data de emiss�o do documento de identifica��o do segurado."
         Top             =   50
         Width           =   1470
      End
      Begin VB.TextBox txtIdeEmpresa 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2985
         MaxLength       =   30
         TabIndex        =   91
         ToolTipText     =   "N�mero do documento de identifica��o do segurado."
         Top             =   600
         Width           =   2925
      End
      Begin VB.TextBox txtRendaMensal 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7040
         MaxLength       =   18
         TabIndex        =   90
         ToolTipText     =   "N�mero do documento de identifica��o do segurado."
         Top             =   600
         Width           =   1920
      End
      Begin Threed.SSPanel panProp 
         Height          =   285
         Index           =   2
         Left            =   7530
         TabIndex        =   97
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
            Caption         =   "Masc"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   50
            TabIndex        =   99
            ToolTipText     =   "Pessoa f�sica."
            Top             =   50
            Width           =   735
         End
         Begin VB.OptionButton optSexo 
            Caption         =   "Fem"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   735
            TabIndex        =   98
            ToolTipText     =   "Pessoa jur�dica."
            Top             =   50
            Width           =   615
         End
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "�rg�o expedidor: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   22
         Left            =   4575
         TabIndex        =   107
         ToolTipText     =   "�rg�o expedidor do documento de identifica��o do segurado."
         Top             =   345
         Width           =   1335
      End
      Begin VB.Label lblProp 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CPF"
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   18
         Left            =   50
         TabIndex        =   106
         ToolTipText     =   "CPF do segurado."
         Top             =   150
         Width           =   1250
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo do documento: "
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   20
         Left            =   1350
         TabIndex        =   105
         ToolTipText     =   "Tipo do documento de identifica��o do segurado."
         Top             =   45
         Width           =   1680
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N�m. do documento: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   21
         Left            =   1350
         TabIndex        =   104
         ToolTipText     =   "N�mero do documento de identifica��o do segurado."
         Top             =   345
         Width           =   1680
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Data de emiss�o: "
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   23
         Left            =   6075
         TabIndex        =   103
         ToolTipText     =   "Data de emiss�o do documento de identifica��o do segurado."
         Top             =   45
         Width           =   1395
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Empr. onde trabalha: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   10
         Left            =   1350
         TabIndex        =   102
         ToolTipText     =   "N�mero do documento de identifica��o do segurado."
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label lblProp 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sexo: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   11
         Left            =   7035
         TabIndex        =   101
         ToolTipText     =   "�rg�o expedidor do documento de identifica��o do segurado."
         Top             =   345
         Width           =   495
      End
      Begin VB.Label lblProp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Renda mensal: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   27
         Left            =   5880
         TabIndex        =   100
         ToolTipText     =   "�rg�o expedidor do documento de identifica��o do segurado."
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Label lblProp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Segurado: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   113
      ToolTipText     =   "C�digo e nome do segurado."
      Top             =   0
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
      TabIndex        =   112
      Top             =   0
      Width           =   800
   End
   Begin VB.Label lblProp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C�d. segurado: "
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   26
      Left            =   0
      TabIndex        =   111
      ToolTipText     =   "C�digo e nome do segurado."
      Top             =   285
      Width           =   1215
   End
   Begin VB.Label lblProp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cota��o: "
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   15
      Left            =   4860
      TabIndex        =   110
      ToolTipText     =   "Cota��o"
      Top             =   285
      Width           =   765
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Refer�ncia: "
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   27
      Left            =   0
      TabIndex        =   109
      ToolTipText     =   "Refer�ncia"
      Top             =   5040
      Width           =   1020
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data para emiss�o:"
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   0
      Left            =   2640
      TabIndex        =   108
      ToolTipText     =   "Refer�ncia"
      Top             =   5040
      Width           =   1500
   End
End
Attribute VB_Name = "frmT47V004A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAltEndereco    As Boolean  'Indica que houve altera��o no endere�o.
Private mAntCEP         As String   'Anterior: CEP.
Private mAntControle    As Control  'Anterior: controle.

Private Sub mpUnidadeSegurado(ByRef pUnidade As Integer, ByRef pMensagem As String)
    Dim lbdP0044700     As ADODB.Connection     'Arquivo de conex�o.
    Dim lrsP0044700     As clsYasRecordSet      'Registro P0044700.
    Dim lSelect         As String
    
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

    Dim lCodProdutor        As Integer  'C�digo do produtor.
    Dim lCodUniProdutiva    As Integer  'C�digo da unidade produtiva.
    Dim lCodUniProdAux      As Integer  'C�digo da unidade produtiva auxiliar
    Dim lNomProdutor        As String   'Nome do produtor.
    Dim lNomUniProdutiva    As String   'Nome da unidade produtiva.
    Dim lTotalRegistros     As Integer  'Total de registros na tabela YTAB_UNID_CORR_PROD
    Dim lListarUnidade      As Boolean  'Indicador se � para listar as unidades para o usu�rios escolher
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
                    Call frmT47V005A.mpInstanciarForm(CLng(pCodCorretor), lCodUniProdutiva, lCodProdutor)
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

    Dim lEndereco   As String   'Endere�o.
    Dim lTamanho    As Integer  'Tamanho.

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
    If gM46V999.gfPreenchido(txtComplemento.Text) Then
        If gM46V999.gfPreenchido(lEndereco) Then
            lEndereco = lEndereco & " " & Trim$(txtComplemento.Text)
        Else
            lEndereco = Trim$(txtComplemento.Text)
        End If
    End If
    lTamanho = Len(lEndereco)
    If lTamanho > 50 Then
        Call gM46V999.gpGraLog(1, "AGE0035 - " & gM46V999.gPrefixo & "Endere�o completo do segurado: " & vbLf & vbLf & _
                         UCase(lEndereco) & vbLf & vbLf & "O endere�o completo est� com " & lTamanho & _
                         " posi��es (tamanho m�ximo permitido no sistema = 50).")
    End If
    mAltEndereco = False
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
Private Sub cboUF_Click()
    'Observa��o: Lista do combo de unidades da federa��o carregadas em tempo de projeto.
    '            Qualquer altera��o deve ser feita na propriedade LIST do controle.

    Dim lI  As Integer  'Utilizado com For...Next.

    Call gpHabSalvar
    If optEndCobranca(1) = True Then    'Endere�o de cobran�a igual.
        For lI = 0 To (frmT47V006A.cboUF.ListCount - 1)
            If frmT47V006A.cboUF.List(lI) = cboUF.Text Then
                frmT47V006A.cboUF.ListIndex = lI
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
    If gCarga = True Then Exit Sub
    Call gpLimPremios
End Sub

Private Sub cmdCEP_Click()
Dim lCEP            As String       'CEP.
Dim lTipLogradouro  As String       'Tipo de logradouro.
Dim lLogradouro     As String       'Logradouro.
Dim lBairro         As String       'Bairro.
Dim lCidade         As String       'Cidade.
Dim lUF             As String       'UF.

    Call frmT47V007A.mpInstanciarForm(lCEP, lTipLogradouro, lLogradouro, lBairro, lCidade, lUF)
    txtCEP.Text = lCEP
    txtTipLogradouro.Text = lTipLogradouro
    txtNomLogradouro.Text = lLogradouro
    txtBairro.Text = lBairro
    txtCidade.Text = lCidade
    If gM46V999.gfPreenchido(lUF) Then
        cboUF.Text = lUF
    End If
End Sub
Private Sub cmdCosseguro_Click()
    frmT47V008A.Show vbModal, Me
    If mAntControle.Visible = True And mAntControle.Enabled = True Then
        mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
    End If
End Sub
Private Sub cmdEndCobranca_Click()
    Call mnuEndCobranca_Click
    On Error Resume Next
    mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
End Sub
Private Sub cmdProdutor_Click()

    If gM46V999.gDesabilitarForms Then Exit Sub
    
    Call gpHabSalvar

    frmT47V012A.Show vbModal, Me
    If mAntControle.Visible = True And mAntControle.Enabled = True Then
        mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
    End If
End Sub

Private Sub cmdSiscad_Click(Index As Integer)
    Dim lCPF_CNPJ   As String
    Dim lBlnContemDados As Boolean
    
    If gM46V999.gDesabilitarForms Then Exit Sub
    
    If optTipPessoa(0).Value = True Then 'F�sica
        Call gM46V999.gpForCPF(2, txtCPF.Text, lCPF_CNPJ)
    Else 'Jur�dica
        Call gM46V999.gpForCNPJ(2, txtCNPJ.Text, lCPF_CNPJ)
    End If

    If InStr(1, gM46V999.gEndExeSiscad, "\") = 0 Then
        gM46V999.gEndExeSiscad = App.Path & "\" & gM46V999.gEndExeSiscad
    End If

    On Error GoTo ControleErro
    
    If Index = 0 Then
        If Dir(gM46V999.gEndExeSiscad) <> "" Then
            Dim lCpfCnpjSiscad As String
            If IsNumeric(lCPF_CNPJ) Then
                lCpfCnpjSiscad = Format(lCPF_CNPJ, String(14, "0"))
            Else
                lCpfCnpjSiscad = Right$(String(14, "0") & Trim(lCPF_CNPJ), 14)
            End If
            Call Shell(gM46V999.gEndExeSiscad & " " & lCpfCnpjSiscad & Format(frmT47V001A.objstcPedido.objstcCtrlEmis.Num_PI, String(10, "0")), vbNormalFocus)
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

    gProChamado = "T47V005A"
    frmT47V005A.Show vbModal, Me
    If mAntControle.Visible = True And mAntControle.Enabled = True Then
        mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
    End If
End Sub
Private Sub cmdForPagCarne_Click()
    frmT47V009A.Show vbModal, Me
    If mAntControle.Visible = True And mAntControle.Enabled = True Then
        mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
    End If
    optForPagamento(0).Value = True
End Sub

Private Sub cmdForPagCarne_GotFocus()
    Set mAntControle = ActiveControl
End Sub

Private Sub cmdForPagDebito_Click()
    Call frmT47V010A.mpInstanciarForm(2)
    If Not gM46V999.gDesabilitarForms Then
        If mAntControle.Visible = True And mAntControle.Enabled = True Then
            mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
        End If
    End If
End Sub

Private Sub cmdForPagDebito_GotFocus()
    Set mAntControle = ActiveControl
End Sub

Private Sub cmdPagAntBloqueto_Click()
    frmT47V011A.Show vbModal, Me
    If mAntControle.Visible = True And mAntControle.Enabled = True And Me.Tag <> "Desabilitado" Then
        mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
    End If
End Sub

Private Sub cmdPagAntBloqueto_GotFocus()
    Set mAntControle = ActiveControl
End Sub

Private Sub cmdPagAntDebito_Click()
    Call frmT47V010A.mpInstanciarForm(1)
    If mAntControle.Visible = True And mAntControle.Enabled = True Then
        mAntControle.SetFocus       'Volta foco para campo onde estava posicionado.
    End If
End Sub

Private Sub cmdPagAntDebito_GotFocus()
    Set mAntControle = ActiveControl
End Sub

Private Sub cmdVoltar_Click()
    '1. Torna form invis�vel.
    If frmT47V001A.lblTipDocumento.Caption = "2" Then   'Proposta.
        If optForPagamento(0).Enabled = True Then
            If optForPagamento(0) = True Then   'Carn�.
                With frmT47V009A
                    If gM46V999.gfDisponivel("frmT47V009A") = True Then
                        If .cboCarBanco.ListIndex = -1 Then
                            .Show vbModal, Me
                        End If
                    Else
                        .Show vbModal, Me
                    End If
                End With
            ElseIf optForPagamento(1) = True Then 'D�bito em conta.
                With frmT47V010A
                    If gM46V999.gfDisponivel("frmT47V110A") = True Then
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
    frmT47V001A.Show
End Sub

Private Sub Form_Activate()
Dim lI As Long

        '1. Posiciona controles.
    frmT47V004A.Caption = Mid$(Me.Name, 4)
    
    If optTipPessoa(0).Value = True Then
        'Pessoa f�sica.
        panPesFisica.Visible = True
        panPesJuridica.Visible = False
            
        'Ficha 900417
        If gM46V999.gCodProduto >= 968 Then
            lblProp(10).Caption = "          Profiss�o: "
        Else
            lblProp(10).Caption = "Empr. onde trabalha: "
        End If
    
    Else 'Pessoa jur�dica.
        panPesFisica.Visible = False
        panPesJuridica.Visible = True
        
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
            lblX(27).Visible = False
            txtReferencia = "VAZIO"
            txtReferencia.Visible = False
            txtValorPgAnt.Visible = False
            lblPagAnt.Visible = False
        Case Else
            cmdSiscad(0).Visible = True
            cmdSiscad(1).Visible = True
            If gM46V999.PerfilAcesso = e_Perfil_DIG_Gestor Then
                lblX(0).Visible = True
                txtDatEmis.Visible = True
            End If
    End Select
    
    'Me.Caption = "Dados da Proposta - PI: " & frmT47V004A.objstcPedido.objstcCtrlEmis.Num_PI
    
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
                cmdAjustarDatas.Enabled = True
                cmdAjustarDatas.Picture = cmdExtra(1).Picture
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
            '[FRH.07]->
            '[REATIVANDO controles no 102A]
                cmdUniProdutiva.Enabled = True: cmdUniProdutiva.Picture = cmdExtra(1).Picture
                cmdProdutor.Enabled = True: cmdProdutor.Picture = cmdExtra(1).Picture
                cmdForPagCarne.Enabled = True: cmdForPagCarne.Picture = cmdExtra(1).Picture
                cmdSiscad(0).Enabled = True: cmdSiscad(0).Picture = cmdExtra(1).Picture
                cmdCEP.Enabled = True: cmdCEP.Picture = cmdExtra(1).Picture
                cmdSiscad(1).Enabled = True
            '<-[FRH]
        End If
    End If
    
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        If lblNomUnidade.Caption = "MATRIZ" And gM46V999.CDblx(TxtCorretor(1).Tag) <> 0 Then
            txtCorretor_Change (1)
            txtCorretor_LostFocus (1)
        End If
    End If
    
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        If gM46V999.gDesabFormsTransmi = True Then
            DesabilitaControles Me
            cmdVoltar.Enabled = True
            cmdSiscad(0).Enabled = True
            cmdSiscad(1).Enabled = True
            optTipPessoa(0).Enabled = True
            optTipPessoa(1).Enabled = True
            cmdAjustarDatas.Enabled = True
            cmdAjustarDatas.Picture = cmdExtra(1).Picture
            txtValorPgAnt.Enabled = True
        End If
    End If
    cmdVoltar.Enabled = True
    
        
End Sub

Public Sub mpAtualizaDadosSegurado(ByRef pContemDados As Boolean, pConfirmar As Boolean)

Dim lCPF_CNPJ   As String
Dim BDSisCad    As ADODB.Connection
Dim rsSisCad    As clsYasRecordSet
Dim rsSyasCob   As clsYasRecordSet
Dim lMensagem   As String
Dim lLibAltEnd  As Boolean
Dim lLibAltNome As Boolean
Dim lSelect     As String
    pContemDados = False

    If pConfirmar = True Then 'SisCadAberto 1
       ' If MsgBox("Deseja atualizar os dados do segurado?", vbYesNo + vbQuestion, Me.Caption) = vbNo Then Exit Sub
        MsgBox "Pressione Ok, ap�s terminar de atualizar os dados do SisCad!", vbOKOnly, Me.Caption
    End If
    Screen.MousePointer = vbHourglass
    If optTipPessoa(0).Value = True Then 'F�sica
        Call gM46V999.gpForCPF(2, txtCPF.Text, lCPF_CNPJ)
    Else 'Jur�dica
        Call gM46V999.gpForCNPJ(2, txtCNPJ.Text, lCPF_CNPJ)
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
                txtSegurado.Text = Trim(rsSisCad("NOME_RAZAO")) & "   " & rsSisCad("NOM_LOCTRAB")
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
                txtComplemento.Text = rsSisCad("DSC_COMPL_LOGR")
                txtBairro.Text = rsSisCad("NOM_BAIRRO")
                txtCidade.Text = rsSisCad("NOM_CID")
                
                If Trim(rsSisCad("SIG_UF")) <> "" Then
                    cboUF.Text = rsSisCad("SIG_UF")
                End If
                txtCEP.Text = Format(rsSisCad("Num_CEP"), "00000-000")
                If rsSisCad("TIP_IRREGULAR") <> 9 Then
                    txtCEP.Enabled = False
                    txtTipLogradouro.Enabled = False
                    txtNomLogradouro.Enabled = False
                    txtComplemento.Enabled = False
                    txtBairro.Enabled = False
                    txtCidade.Enabled = False
                    cboUF.Enabled = False
                    txtCEP.Enabled = False
                Else
                    txtCEP.Enabled = True
                    txtTipLogradouro.Enabled = True
                    txtNomLogradouro.Enabled = True
                    txtComplemento.Enabled = True
                    txtBairro.Enabled = True
                    txtCidade.Enabled = True
                    cboUF.Enabled = True
                    txtCEP.Enabled = True
                End If
                
            Else
                txtCEP.Enabled = True
                txtTipLogradouro.Enabled = True
                txtNomLogradouro.Enabled = True
                txtComplemento.Enabled = True
                txtBairro.Enabled = True
                txtCidade.Enabled = True
                cboUF.Enabled = True
                txtCEP.Enabled = True
            End If
            txtTelefones.Text = rsSisCad("NUM_DDD_FONE") & rsSisCad("NUM_FONE")
            
            '[FRH.08]->
            txtDigCodSegurado.Enabled = False
            
            If gM46V999.gTipUsuario <> e_Tipo_Usu_Yasuda Then
              txtCodSegurado.Text = rsSisCad("COD_IDENTIFICADOR")
              txtCodSegurado.Enabled = False
            Else
              txtCodSegurado.Enabled = True
            End If
            '<-[FRH]
            
            If (rsSisCad("SEXO") = "M") Then
                optSexo(0).Value = True
            Else
                optSexo(1).Value = True
            End If
            cboIdeTipDocumento.ListIndex = gM46V999.gfRetornarListIndex(cboIdeTipDocumento, Val(Trim(rsSisCad("TIP_DOCUM"))))
            txtIdeNumDocumento.Text = rsSisCad("NUM_DOCUM")
            txtIdeOrgExpedidor.Text = rsSisCad("NOM_ORG_EXPED")
            txtIdeDatEmissao.Text = gM46V999.gfForData(2, rsSisCad("DAT_EXPED"))
            txtIdeEmpresa.Text = IIf(IsNull(rsSisCad("NOM_LOCTRAB")), "", rsSisCad("NOM_LOCTRAB"))
                        
            txtTelefones.Enabled = False
            
            optSexo(0).Enabled = False
            optSexo(1).Enabled = False
            cboIdeTipDocumento.Enabled = False
            txtIdeNumDocumento.Enabled = False
            txtIdeOrgExpedidor.Enabled = False
            txtIdeDatEmissao.Enabled = False
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
                Case "R"
                    cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 18)
                Case "S"
                    cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 19)
                Case "T"
                    cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 20)
                Case "U"
                    cboRamAtividade.ListIndex = gM46V999.gfRetornarListIndex(cboRamAtividade, 21)
            End Select
            cboRamAtividade.Enabled = False
                        
            lSelect = " select * from p0044800.dbo.T0044807 Where "
            lSelect = lSelect & " num_pi = " & frmT47V001A.objstcPedido.objstcCtrlEmis.Num_PI
            lSelect = lSelect & " and num_cgc_cpf = '" & Trim(lCPF_CNPJ) & "'"
            If gM46V999.gfObtRegistro(BDSisCad, lSelect, rsSyasCob, lMensagem) = False Then
                Call gM46V999.gpFecha2(BDSisCad)
                Call gM46V999.gpGraLog(1, lMensagem)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            With frmT47V006A
                If lLibAltEnd = False Then
                    .txtCEP.Enabled = False
                    .txtTipLogradouro.Enabled = False
                    .txtNomLogradouro.Enabled = False
                    .txtComplemento.Enabled = False
                    .txtBairro.Enabled = False
                    .txtCidade.Enabled = False
                    .cboUF.Enabled = False
                    .txtCEP.Enabled = False
                    .txtTelefones.Enabled = False
                Else
                    .txtCEP.Enabled = True
                    .txtTipLogradouro.Enabled = True
                    .txtNomLogradouro.Enabled = True
                    .txtComplemento.Enabled = True
                    .txtBairro.Enabled = True
                    .txtCidade.Enabled = True
                    .cboUF.Enabled = True
                    .txtCEP.Enabled = True
                    .txtTelefones.Enabled = True
                End If
                If rsSyasCob.EOF Then
                    .txtNomLogradouro.Text = ""
                Else
                    .txtTipLogradouro.Text = rsSyasCob("COD_TIP_LOGR")
                    .txtNomLogradouro.Text = rsSyasCob("NOM_LOGR")
                    .txtComplemento.Text = rsSyasCob("DSC_COMPL_LOGR")
                    .txtBairro.Text = rsSyasCob("NOM_BAIRRO")
                    .txtCidade.Text = rsSyasCob("NOM_CID")
                    .txtCEP.Text = Format(rsSyasCob("Num_CEP"), "00000-000")
                    If Trim(rsSyasCob("SIG_UF")) <> "" Then
                        .cboUF.Text = rsSyasCob("SIG_UF")
                    End If
                    If Not IsNull(rsSyasCob("num_fone")) Then
                        .txtTelefones.Text = rsSyasCob("num_ddd_fone") & rsSyasCob("num_fone")
                    Else
                        .txtTelefones.Text = 0
                    End If
                    gCarga = True
                    
                    'Interface Simony
                    'TxtNumCobr = gM46V999.Check(rsSyasCob("DSC_COMPL_LOGR_SN"))
                    'TxtNumResid = gM46V999.Check(rsSyasCob("DSC_COMPL_LOGR_RESID_SN"))
                    'TxtComplResid = Left(Trim(gM46V999.Check(rsSyasCob("NUM_LOGR_RESID"))), 13)
                    'TxtComplCobr = Left(Trim(gM46V999.Check(rsSyasCob("NUM_LOGR"))), 13)
                    
                    If Trim(.txtComplemento.Text) = "" Then
                        If CDblx(TxtNumCobr) <> 0 Then
                            .txtComplemento.Text = TxtNumCobr & " " & TxtComplCobr
                        Else
                            .txtComplemento.Text = TxtComplCobr
                        End If
                    End If
                    
                    If Trim(txtComplemento.Text) = "" Then
                        If CDblx(TxtNumResid) <> 0 Then
                            txtComplemento.Text = TxtNumResid & " " & TxtComplResid
                        Else
                            txtComplemento.Text = TxtComplResid
                        End If
                    End If
                    
                    optEndCobranca(0).Value = True
                    gCarga = False
                    pContemDados = True
                End If
            End With
        Else
            txtSegurado.Text = ""
            txtTipLogradouro.Text = ""
            txtNomLogradouro.Text = ""
            txtComplemento.Text = ""
            txtBairro.Text = ""
            txtCidade.Text = ""
            cboUF.ListIndex = -1
            txtCEP.Text = ""
            txtTelefones.Text = ""
            cboIdeTipDocumento.ListIndex = -1
            txtIdeNumDocumento.Text = ""
            txtIdeOrgExpedidor.Text = ""
            txtIdeDatEmissao.Text = ""
            txtIdeEmpresa.Text = ""
            cboRamAtividade.ListIndex = -1
            frmT47V006A.txtCEP.Text = ""
            frmT47V006A.txtTipLogradouro.Text = ""
            frmT47V006A.txtNomLogradouro.Text = ""
            frmT47V006A.txtComplemento.Text = ""
            frmT47V006A.txtBairro.Text = ""
            frmT47V006A.txtCidade.Text = ""
            frmT47V006A.cboUF.ListIndex = -1
            frmT47V006A.txtTelefones.Text = ""
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
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
        .AddItem "Eletricidade e g�s"
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
    With cboIdeTipDocumento
        .AddItem "RG / Carteira de identidade"
        .ItemData(.NewIndex) = 1
        .AddItem "RNE / Identidade estrangeira"
        .ItemData(.NewIndex) = 2
        .AddItem "CNH / Carteira Nacional de Habilita��o"
        .ItemData(.NewIndex) = 3
        .AddItem "Certid�o de nascimento"
        .ItemData(.NewIndex) = 4
        .AddItem "Certid�o de casamento"
        .ItemData(.NewIndex) = 5
        .AddItem "Passaporte"
        .ItemData(.NewIndex) = 6
        .AddItem "Carteira de trabalho"
        .ItemData(.NewIndex) = 7
        .AddItem "Certificado de reservista"
        .ItemData(.NewIndex) = 8
        .AddItem "PIS"
        .ItemData(.NewIndex) = 9
        .AddItem "T�tulo de eleitor"
        .ItemData(.NewIndex) = 10
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
    
    '[FRH]-> Se usu�rio corretor..
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
     txtTelefones.MaxLength = 8 '..permite armazenar Fone de 8 dig.
    Else
     txtTelefones.MaxLength = 7
    End If
    '<-[FRH]
End Sub

Private Sub mnuEndCobranca_Click()
    frmT47V006A.Show vbModal, Me
End Sub
Private Sub mnuProdutor_Click()
    Call gpHabSalvar

    gProChamado = "T47V012A"
    frmT47V012A.Show vbModal, Me
End Sub
Private Sub mnuUniProdutiva_Click()
    Call gpHabSalvar

    gProChamado = "T47V005A"
    frmT47V005A.Show vbModal, Me
End Sub
Private Sub optCosseguro_Click(Index As Integer)
    Call gpHabSalvar
    Select Case Index
        Case 0
            cmdCosseguro.Enabled = True
            cmdCosseguro.Picture = cmdExtra(1).Picture
            If gCarga = False Then
                frmT47V008A.Show vbModal, Me
            End If
        Case 1
            cmdCosseguro.Enabled = False
            cmdCosseguro.Picture = cmdExtra(2).Picture
    End Select
    
End Sub
Private Sub optCosseguro_GotFocus(Index As Integer)
    Set mAntControle = ActiveControl
End Sub
Private Sub optEndCobranca_Click(Index As Integer)
    Dim lI  As Byte     'Utilizado com For...Next.

    Select Case Index
        Case 0      'Sim.
            cmdEndCobranca.Enabled = True
            cmdEndCobranca.Picture = cmdExtra(1).Picture
        Case 1      'N�o tem endere�o de cobran�a diferente
            cmdEndCobranca.Enabled = False
            cmdEndCobranca.Picture = cmdExtra(2).Picture
    End Select
    If gCarga = False Then
        Call gpHabSalvar
        Select Case Index
            Case 0      'Sim.
                If gM46V999.gfPreenchido(txtTipLogradouro.Text) Then
                    frmT47V006A.txtTipLogradouro.Text = txtTipLogradouro
                End If
                If gM46V999.gfPreenchido(txtNomLogradouro.Text) Then
                    frmT47V006A.txtNomLogradouro.Text = txtNomLogradouro.Text
                End If
                If gM46V999.gfPreenchido(txtComplemento.Text) Then
                    frmT47V006A.txtComplemento.Text = txtComplemento.Text
                End If
                If gM46V999.gfPreenchido(txtBairro.Text) Then
                    frmT47V006A.txtBairro.Text = txtBairro.Text
                End If
                If gM46V999.gfPreenchido(txtCidade.Text) Then
                    frmT47V006A.txtCidade.Text = txtCidade.Text
                End If
                If cboUF.ListIndex <> -1 Then
                    If frmT47V006A.cboUF.ListCount > 0 Then
                        For lI = 0 To (frmT47V006A.cboUF.ListCount - 1)
                            If frmT47V006A.cboUF.List(lI) = cboUF.Text Then
                                frmT47V006A.cboUF.ListIndex = lI
                                Exit For
                            End If
                        Next lI
                    End If
                End If
                If gM46V999.gfPreenchido(txtCEP.Text) Then
                    frmT47V006A.txtCEP.Text = txtCEP.Text
                End If
                If gM46V999.gfPreenchido(txtTelefones.Text) Then
                    frmT47V006A.txtTelefones.Text = txtTelefones.Text
                End If
                frmT47V006A.Show vbModal, Me
            Case 1      'N�o tem endere�o de cobran�a diferente
                On Error Resume Next
                Unload frmT47V006A
                On Error GoTo 0
        End Select
    End If
End Sub
Private Sub optEndCobranca_GotFocus(Index As Integer)
    Set mAntControle = ActiveControl
End Sub
Private Sub optEndCobranca_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optForPagamento_Click(Index As Integer)
    Call gpHabSalvar
    Select Case Index
        Case 0  'Carn�.
            cmdForPagCarne.Enabled = True
            cmdForPagCarne.Picture = cmdExtra(1).Picture
            cmdForPagDebito.Enabled = False
            cmdForPagDebito.Picture = cmdExtra(2).Picture
'Ficha 900046
            If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
                cmdPagAntBloqueto.Visible = True
            End If
            If gCarga = False Then
                frmT47V009A.Show vbModal, Me
            End If
            If optPagAntecipado(2).Value = False Then
                If gM46V999.gfDisponivel("frmT47V010A") Then
                    Unload frmT47V010A
                End If
            End If
        Case 1  'D�bito autom�tico.
            cmdForPagCarne.Enabled = False
            cmdForPagCarne.Picture = cmdExtra(2).Picture
            cmdForPagDebito.Enabled = True
            cmdForPagDebito.Picture = cmdExtra(1).Picture
'Ficha 900046
            If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
                cmdPagAntBloqueto.Visible = False
            End If
            If gCarga = False Then
                If frmT47V001A.mCarregarPI = False Then Call frmT47V010A.mpInstanciarForm(2)
            End If
            If gM46V999.gfDisponivel("frmT47V009A") Then
                Unload frmT47V009A
            End If
    End Select
End Sub
Private Sub optForPagamento_GotFocus(Index As Integer)
    Set mAntControle = ActiveControl
End Sub
Private Sub optForPagamento_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optPagAntecipado_Click(Index As Integer)
Dim lobjstcBoleto As stcA28V720B
Dim lobjBoleto  As clsA28V720A
    Call gpLimPremios
    Select Case Index
        Case 1  'Sim - Antecipado.
        'Ficha 900046
            If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
                cmdPagAntBloqueto.Enabled = True
                cmdPagAntBloqueto.Picture = cmdExtra(1).Picture
                cmdPagAntDebito.Enabled = False
                cmdPagAntDebito.Picture = cmdExtra(2).Picture

                txtValorPgAnt.Visible = True
                frmT47V011A.txtPagAntBloqueto.Enabled = True
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
                    frmT47V011A.Show vbModal, Me
                End If
            End If
            If gM46V999.gfDisponivel("frmT47V110A") = True Then
                If optForPagamento(1) = False Then
                    Unload frmT47V010A
                End If
            End If
            If txtValorPgAnt.Visible Then
                txtValorPgAnt.Enabled = True
                txtValorPgAnt.SetFocus
            End If
        Case 2  'N�o - a 30 dias.
            cmdPagAntBloqueto.Enabled = False
            cmdPagAntBloqueto.Picture = cmdExtra(2).Picture
            cmdPagAntDebito.Enabled = True
            cmdPagAntDebito.Picture = cmdExtra(1).Picture
'Ficha 900046
            txtValorPgAnt.Text = "0,00"
            txtValorPgAnt.Visible = False
            frmT47V011A.txtPagAntBloqueto.Text = "0"
            frmT47V011A.txtPagAntBloqueto.Enabled = False
'Ficha 900046
'            If gCarga = False Then
'                Call frmT47V110A.mpInstanciarForm(1)
'                gCarga = True
'                optForPagamento(1).Value = True
'                gCarga = False
'            End If
            If gM46V999.gfDisponivel("frmT47V011A") = True Then
                Unload frmT47V011A
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
            If gM46V999.gfDisponivel("frmT47V011A") = True Then
                Unload frmT47V011A
            End If
            If gM46V999.gfDisponivel("frmT47V010A") = True Then
                If optForPagamento(1) = False Then
                    Unload frmT47V010A
                End If
            End If
            txtValorPgAnt.Text = "0,00"
            txtValorPgAnt.Enabled = False
    End Select

    Call txtQuaParcelas_LostFocus   'Considera informa��o de quantidade de parcelas.
End Sub
Private Sub optPagAntecipado_GotFocus(Index As Integer)
    Set mAntControle = ActiveControl
End Sub
Private Sub optPagAntecipado_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optSexo_GotFocus(Index As Integer)
    Set mAntControle = ActiveControl
End Sub
Private Sub optSexo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optTipPessoa_Click(Index As Integer)
    frmT47V001A.optTipPessoa(Index) = True
    Select Case Index
        Case 0  'Pessoa f�sica.
            panPesFisica.Visible = True
            panPesJuridica.Visible = False
        Case 1  'Pessoa jur�dica.
            panPesFisica.Visible = False
            panPesJuridica.Visible = True
    End Select

End Sub
Private Sub optTipPessoa_GotFocus(Index As Integer)
    Set mAntControle = ActiveControl
End Sub
Private Sub optTipPessoa_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtBairro_Change()
    Call gpHabSalvar
    If optEndCobranca(1) = True Then    'Endere�o de cobran�a igual.
        frmT47V006A.txtBairro.Text = txtBairro.Text
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
    With txtCEP
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
            lI = txtCEP.SelStart
            If Len(txtCEP.Text) = 8 And lI <> 0 Then
                KeyAscii = 0
            End If
        Case 8
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCEP_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(txtCEP.Text) > 8 And InStr(1, txtCEP.Text, "-") = 0 Then
        txtCEP.Text = Left(txtCEP.Text, 8)
    End If
End Sub

Private Sub txtCEP_LostFocus()
    Dim lbdP0043000 As ADODB.Connection 'Arquivo - P0043000.
    Dim lMensagem   As String           'Mensagem.
    Dim lrsP0043000 As clsYasRecordSet  'Registro - P0043000.
    Dim lPos        As Byte
    Dim lCEP        As String
    Dim lSelect     As String
    Screen.MousePointer = vbHourglass

    With txtCEP
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
        If gM46V999.gfPreenchido(lrsP0043000!TIP_LOGR) Then
            txtTipLogradouro.Text = lrsP0043000!TIP_LOGR
        Else
            txtTipLogradouro.Text = ""
        End If
        If gM46V999.gfPreenchido(lrsP0043000!Nom_Logr) Then
            txtNomLogradouro.Text = Left(lrsP0043000!Nom_Logr, IIf(InStr(1, lrsP0043000!Nom_Logr, "-") > 0, InStr(1, lrsP0043000!Nom_Logr, "-") - 2, 200))
        Else
            txtNomLogradouro.Text = ""
        End If
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
        frmT47V006A.txtCidade.Text = txtCidade.Text
    End If
End Sub
Private Sub txtCidade_GotFocus()
    Set mAntControle = ActiveControl
    With txtCidade
        .SelStart = 0
        .SelLength = .MaxLength
    End With
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
    Set mAntControle = ActiveControl
    With txtCNPJ
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtCNPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtCNPJ_LostFocus()
    Dim lAuxiliar   As String   'Auxiliar.

    Call gM46V999.gpForCNPJ(1, txtCNPJ.Text, lAuxiliar)
    If lAuxiliar <> "" Then
        txtCNPJ.Text = lAuxiliar
    Else
        If Len(txtCNPJ.Text) = 18 And InStr(1, txtCNPJ.Text, ".") = 0 Then
            txtCNPJ.Text = ""
        Else
            If lAuxiliar <> "" Then
                txtCNPJ.Text = lAuxiliar
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
    txtCodSegurado.Text = Format(Val(txtCodSegurado.Text), "0000000")
    frmT47V001A.txtSegurado.Tag = txtCodSegurado.Text
End Sub
Private Sub txtComplemento_Change()
    mAltEndereco = True
    Call gpHabSalvar
    If optEndCobranca(1) = True Then    'Endere�o de cobran�a igual.
        frmT47V006A.txtComplemento.Text = txtComplemento.Text
    End If
End Sub
Private Sub txtComplemento_GotFocus()
    Set mAntControle = ActiveControl
    With txtComplemento
        .SelStart = 0
        .SelLength = .MaxLength
    End With
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
Private Sub txtCorretor_Change(Index As Integer)
    TxtCorretor(Index).Tag = ""
    Call gpHabSalvar
End Sub
Private Sub txtCorretor_GotFocus(Index As Integer)
    Set mAntControle = ActiveControl
End Sub
Private Sub txtCorretor_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtCorretor_LostFocus(Index As Integer)
    Dim lCodProdutor        As Integer  'C�digo do produtor.
    Dim lCodUniProdutiva    As Integer  'C�digo da unidade produtiva.
    Dim lCorretor           As String   'Corretor.
    Dim lI                  As Double   'Utilizado com For...Next.
    Dim lMensagem           As String   'Mensagem.
    Dim lNomProdutor        As String   'Nome do produtor.
    Dim lNomUniProdutiva    As String   'Nome da unidade produtiva.
    Dim lQuaCorretores      As Double   'Quantidade de corretores selecionados.
    Dim lSaida()            As Variant  'Resultado da busca de corretores.
    Dim lTotal              As Double   'Total
    Screen.MousePointer = vbHourglass

    '1. Se c�digo do corretor preenchido, encerrar processo.
    If gM46V999.gfPreenchido(TxtCorretor(Index).Tag) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '2. Limpa campos de unidade produtiva e produtor.
    If Index = 1 Then
        lblNomUnidade.Tag = " "
        lblNomUnidade.Caption = " "
        lblNomProdutor.Tag = " "
        lblNomProdutor.Caption = " "
    End If

    '3. Se nome do corretor n�o preenchido, encerra processo.
    If Not gM46V999.gfPreenchido(TxtCorretor(Index).Text) Then
        txtCorComissao(Index).Text = "0,00"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '4. Posiciona campo com o que foi informado pelo usu�rio.
    lCorretor = gM46V999.gfRetirarCaracterEspecial(Trim$(TxtCorretor(Index).Text))

    '5. Se foram digitadas at� 6 posi��es num�ricas, trata de c�digo de corretor.
    '   Obt�m corretor, unidade produtiva e produtor.
    If IsNumeric(lCorretor) And Len(CStr(Val(lCorretor))) <= 6 Then
        If gM46V999.gfObtCorretor(1, Val(lCorretor), lSaida, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            TxtCorretor(Index).Text = ""
            If TxtCorretor(Index).Visible = True And TxtCorretor(Index).Enabled = True Then
                TxtCorretor(Index).SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        TxtCorretor(Index).Text = lSaida(1, 3)
        TxtCorretor(Index).Tag = lSaida(1, 1)
        If Index = 1 Then
            If mfPosUniProProdutor(TxtCorretor(1).Tag, lMensagem) = False Then
                Call gM46V999.gpGraLog(1, lMensagem)
                TxtCorretor(Index).Text = ""
                If TxtCorretor(Index).Visible = True And TxtCorretor(Index).Enabled = True Then
                    TxtCorretor(Index).SetFocus
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '5. Se preenchido com menos de 3 posi��es, encerra processo.
    If Len(lCorretor) < 3 Then
        Call gM46V999.gpGraLog(1, "FGE0075 - Corretor inv�lido (digitar pelo menos 3 posi��es).")
        TxtCorretor(Index).Text = ""
        If TxtCorretor(Index).Visible = True And TxtCorretor(Index).Enabled = True Then
            TxtCorretor(Index).SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '6. Obt�m registros com informa��o do usu�rio.
    'a) Obt�m lista de corretores
    If gM46V999.gfObtCorretor(2, lCorretor, lSaida, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        TxtCorretor(Index).Text = ""
        If TxtCorretor(Index).Visible = True And TxtCorretor(Index).Enabled = True Then
            TxtCorretor(Index).SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'b) Se campo de sa�da estiver vazio, nada a fazer.
    If IsNull(lSaida) Then
        TxtCorretor(Index).Text = ""
        If TxtCorretor(Index).Visible = True And TxtCorretor(Index).Enabled = True Then
            TxtCorretor(Index).SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'c) Posiciona quantidade de corretores encontrados.
    lQuaCorretores = UBound(lSaida)
    'd) Se encontrou apenas 1 corretor, posiciona dados deste corretor e encerra processo.
    If lQuaCorretores = 1 Then
        
        TxtCorretor(Index).Text = lSaida(1, 3)
        TxtCorretor(Index).Tag = lSaida(1, 1)
        If Index = 1 Then
            If mfPosUniProProdutor(TxtCorretor(1).Tag, lMensagem) = False Then
                Call gM46V999.gpGraLog(1, lMensagem)
            End If
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'e) Se encontrou mais de um corretor, abre janela com grade de corretores.
    Load frmT47V013A
    With frmT47V013A.flxCorretores
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
    frmT47V013A.Show vbModal, Me
    TxtCorretor(Index).Text = gNomCorretor
    TxtCorretor(Index).Tag = gCodCorretor
    If TxtCorretor(Index).Visible = True And TxtCorretor(Index).Enabled = True Then
        TxtCorretor(Index).SetFocus
    End If
    If Index = 1 Then
        If gM46V999.gfPreenchido(TxtCorretor(Index).Tag) Then
            If mfPosUniProProdutor(TxtCorretor(1).Tag, lMensagem) = False Then
                Call gM46V999.gpGraLog(1, lMensagem)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If

    Screen.MousePointer = vbDefault
    
End Sub
Private Sub txtCPF_Change()
    Call gpHabSalvar
End Sub
Private Sub txtCPF_GotFocus()
    Set mAntControle = ActiveControl
    With txtCPF
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtCPF_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtCPF_LostFocus()
    Dim lAuxiliar   As String   'Auxiliar.

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
    Set mAntControle = ActiveControl
    With txtDatEmis
        .SelStart = 0
        .SelLength = .MaxLength
    End With
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

Private Sub txtIdeDatEmissao_Change()
    Call gpHabSalvar
End Sub
Private Sub txtIdeDatEmissao_GotFocus()
    Set mAntControle = ActiveControl
    With txtIdeDatEmissao
        .SelStart = 0
        .SelLength = .MaxLength
    End With
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
    If IsDate(txtIdeDatEmissao.Text) = False Then
        txtIdeDatEmissao.Text = ""
    End If
End Sub
Private Sub txtIdeEmpresa_GotFocus()
    Set mAntControle = ActiveControl
    With txtIdeEmpresa
        .SelStart = 0
        .SelLength = .MaxLength
    End With
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
    Set mAntControle = ActiveControl
    With txtIdeNumDocumento
        .SelStart = 0
        .SelLength = .MaxLength
    End With
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
    Set mAntControle = ActiveControl
    With txtIdeOrgExpedidor
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtIdeOrgExpedidor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtCorComissao_GotFocus(Index As Integer)
    Set mAntControle = ActiveControl
    With txtCorComissao(Index)
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtCorComissao_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtCorComissao_LostFocus(Index As Integer)
    Dim lI              As Byte     'Utilizado com For...Next.
    Dim lTotComissao    As Double   'Total de comiss�o.

    '1. Procedimento para edi��o da comiss�o.
    With txtCorComissao(Index)
        If gM46V999.gfPreenchido(.Text) Then
            If IsNumeric(.Text) Then
                .Text = Format$(CDbl(.Text), "standard")
            Else
                .Text = Format$("0,00", "standard")
            End If
        Else
            .Text = Format$("0,00", "standard")
        End If
        
        If Index = 2 Or Index = 3 Then
            If Val(TxtCorretor(Index).Tag) = 0 Then
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
        frmT47V001A.TxtComissao.Text = ""
    Else
        frmT47V001A.TxtComissao.Text = Trim$(Format$(lTotComissao, "standard"))
    End If

End Sub

Private Sub txtJuros_Change()
    Call gpLimPremios
End Sub

Private Sub txtJuros_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")
End Sub


Private Sub txtJuros_LostFocus()
Dim lI As Integer
    If gCarga = True Then Exit Sub
    txtJuros.Text = Format(gM46V999.CDblx(txtJuros.Text), "standard")
    Call gpLimPremios
End Sub

Private Sub txtNomLogradouro_Change()
    mAltEndereco = True
    Call gpHabSalvar
    If optEndCobranca(1) = True Then    'Endere�o de cobran�a igual.
        frmT47V006A.txtNomLogradouro.Text = txtNomLogradouro.Text
    End If
End Sub
Private Sub txtNomLogradouro_GotFocus()
    Set mAntControle = ActiveControl
    With txtNomLogradouro
        .SelStart = 0
        .SelLength = .MaxLength
    End With
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

Private Sub txtNumSiscota_Change()
    If gCarga = True Then Exit Sub
    Call gpLimPremios
End Sub

Private Sub txtNumSiscota_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, "")
End Sub


Private Sub txtQuaParcelas_Change()
    If gCarga = True Then Exit Sub
    Call gpLimPremios
End Sub
Private Sub txtQuaParcelas_GotFocus()
    Set mAntControle = ActiveControl
    With txtQuaParcelas
        .SelStart = 0
        .SelLength = .MaxLength
    End With
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
            If CInt(Me.txtQuaParcelas.Text) > gM46V999.gProQuaMaxParcelas Then
                Me.txtQuaParcelas.Text = gM46V999.gProQuaMaxParcelas
            End If
'            If CInt(Me.txtQuaParcelas.Text) = 0 Then txtQuaParcelas.Text = gM46V999.gProQuaMaxParcelas
        Else
            txtQuaParcelas.Text = gM46V999.gProQuaMaxParcelas
        End If
    Else
        If gM46V999.gTipoAcesso <> DIGITACAO Then
            txtQuaParcelas.Text = gM46V999.gProQuaMaxParcelas
        End If
    End If
    
End Sub


Private Sub txtReferencia_Change()
    If gCarga = True Then Exit Sub
    Call gpLimPremios
End Sub

Private Sub txtSegurado_Change()
    Call gpHabSalvar
    frmT47V001A.txtSegurado.Text = txtSegurado.Text
End Sub
Private Sub txtSegurado_GotFocus()
    Set mAntControle = ActiveControl
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
Private Sub txtSegurado_LostFocus()
    txtSegurado.Text = gM46V999.gfRetirarCaracterEspecial(txtSegurado.Text, True)
End Sub
Private Sub txtTelefones_Change()
    Call gpHabSalvar
    If optEndCobranca(1) = True Then    'Endere�o de cobran�a igual.
        frmT47V006A.txtTelefones.Text = txtTelefones.Text
    End If
End Sub
Private Sub txtTelefones_GotFocus()
    Set mAntControle = ActiveControl
    With txtTelefones
        .SelStart = 0
        .SelLength = .MaxLength
    End With
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
        frmT47V006A.txtTipLogradouro.Text = txtTipLogradouro.Text
    End If
End Sub
Private Sub txtTipLogradouro_GotFocus()
    Set mAntControle = ActiveControl
    With txtTipLogradouro
        .SelStart = 0
        .SelLength = .MaxLength
    End With
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
    Set mAntControle = ActiveControl
    With txtValorPgAnt
        .SelStart = 0
        .SelLength = .MaxLength
    End With
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

Private Sub cmdAjustarDatas_Click()
    frmT47v019A.Show vbModal, Me
End Sub

Private Sub cmdAjustarDatas_GotFocus()
    Set mAntControle = ActiveControl
End Sub

Private Sub optVencimento_Click(Index As Integer)
Dim lMensagem As String
    
    '1. Posiciona controles.
    Select Case Index
        Case 0  'Sim.
            cmdAjustarDatas.Enabled = True
            cmdAjustarDatas.Picture = cmdExtra(1).Picture
        Case 1  'N�o.
            cmdAjustarDatas.Enabled = False
            cmdAjustarDatas.Picture = cmdExtra(2).Picture
    End Select

    If gCarga = True Then
        Exit Sub
    End If
    Call gpHabSalvar
    '4. Processa op��o do usu�rio.
    If frmT47V001A.mCarregarDatas = False Then Exit Sub
    
    Select Case Index
        Case 0  'Sim.
            Call cmdAjustarDatas_Click
        Case 1  'N�o.
            lMensagem = "AGE0016 - " & gM46V999.gPrefixo & "Exclu�da op��o de ajuste de vencimento. " & _
                        "Datas ajustadas ser�o desconsideradas. Confirma esta op��o?"
            If MsgBox(lMensagem, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call gM46V999.gpGraLog(0, lMensagem & " Sim.")
                On Error Resume Next
                Unload frmT47v019A
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


Private Sub txtCorComissao_Change(Index As Integer)
    Call gpLimPremios
    If gCarga = False Then
        gAltComum = True
    End If
End Sub

