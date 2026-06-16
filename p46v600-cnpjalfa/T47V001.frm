VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmT47V001A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalhes do risco"
   ClientHeight    =   7155
   ClientLeft      =   4860
   ClientTop       =   1410
   ClientWidth     =   9510
   Icon            =   "T47V001.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   9510
   Begin Threed.SSPanel panItemLocRisco 
      Height          =   1980
      Left            =   0
      TabIndex        =   102
      Top             =   3000
      Visible         =   0   'False
      Width           =   9450
      _Version        =   65536
      _ExtentX        =   16669
      _ExtentY        =   3492
      _StockProps     =   15
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
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   1
      Alignment       =   1
      Begin VB.CommandButton cmdExcluirCobertura 
         Caption         =   "- C"
         Height          =   255
         Left            =   8880
         TabIndex        =   28
         ToolTipText     =   "Excluir a cobertura."
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdAdicionarCobertura 
         Caption         =   "+ C"
         Height          =   255
         Left            =   8880
         TabIndex        =   27
         ToolTipText     =   "Adicionar uma cobertura."
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdExcluirModalidade 
         Caption         =   "- M"
         Height          =   255
         Left            =   8880
         TabIndex        =   26
         ToolTipText     =   "Excluir a modalidade."
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdAdicionarTab 
         Caption         =   "+ M"
         Height          =   255
         Left            =   8880
         TabIndex        =   25
         ToolTipText     =   "Adicionar uma modalidade."
         Top             =   360
         Width           =   495
      End
      Begin VB.ComboBox CmbCoberturas 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   0
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtModalidade 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   105
         Text            =   "Selecione uma modalidade:"
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox cboPlano 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   24
         Text            =   "cboPlano"
         ToolTipText     =   "Modalidade (plano)."
         Top             =   600
         Width           =   4425
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxGridCoberturaDemais 
         Height          =   1425
         Index           =   0
         Left            =   90
         TabIndex        =   103
         Top             =   375
         Visible         =   0   'False
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   2514
         _Version        =   393216
         ForeColor       =   8388608
         ForeColorFixed  =   8388608
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   0
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.TabStrip TabMod 
         Height          =   1935
         Left            =   45
         TabIndex        =   107
         Top             =   -30
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   3413
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin Threed.SSPanel panItem 
      Height          =   375
      Left            =   0
      TabIndex        =   46
      Top             =   2040
      Visible         =   0   'False
      Width           =   9480
      _Version        =   65536
      _ExtentX        =   16722
      _ExtentY        =   661
      _StockProps     =   15
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
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Begin VB.TextBox txtClassificacao 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4860
         TabIndex        =   13
         ToolTipText     =   "Classifica��o"
         Top             =   50
         Width           =   4245
      End
      Begin VB.ComboBox cboPlano 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Plano."
         Top             =   45
         Width           =   2625
      End
      Begin Threed.SSCommand cmdClassificacao 
         Height          =   285
         Left            =   9120
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Busca da classifica��o."
         Top             =   45
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   503
         _StockProps     =   78
         Picture         =   "T47V001.frx":0442
      End
      Begin VB.Label lblAntNumItem 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   8535
         TabIndex        =   49
         Top             =   90
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Classifica��o: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   8
         Left            =   3690
         TabIndex        =   48
         Top             =   45
         Width           =   1165
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Modalidade:"
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   47
         Top             =   45
         Width           =   1020
      End
   End
   Begin Threed.SSPanel panComum 
      Height          =   630
      Left            =   0
      TabIndex        =   52
      Top             =   1320
      Visible         =   0   'False
      Width           =   9480
      _Version        =   65536
      _ExtentX        =   16722
      _ExtentY        =   1111
      _StockProps     =   15
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
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Begin VB.TextBox txtComissao 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   8715
         MaxLength       =   5
         TabIndex        =   6
         ToolTipText     =   "Comiss�o."
         Top             =   30
         Width           =   720
      End
      Begin Threed.SSPanel panProduto 
         Height          =   285
         Left            =   4860
         TabIndex        =   53
         Top             =   315
         Width           =   4575
         _Version        =   65536
         _ExtentX        =   8070
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   12582912
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
         BorderWidth     =   0
         BevelOuter      =   0
         BevelInner      =   1
         Begin Threed.SSCommand cmdProduto 
            Height          =   240
            Left            =   4305
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Busca protocolo."
            Top             =   30
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   8388608
            Picture         =   "T47V001.frx":075C
         End
         Begin VB.Label lblMulti 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Produto: "
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   7
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label lblNomProduto 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1095
            TabIndex        =   10
            ToolTipText     =   "Protocolo."
            Top             =   0
            Width           =   3225
         End
         Begin VB.Label lblTipDocumento 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   230
            Left            =   50
            TabIndex        =   55
            Top             =   30
            Visible         =   0   'False
            Width           =   150
         End
      End
      Begin Threed.SSPanel panVigencia 
         Height          =   285
         Left            =   45
         TabIndex        =   57
         Top             =   315
         Width           =   4815
         _Version        =   65536
         _ExtentX        =   8493
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
         BevelOuter      =   0
         BevelInner      =   1
         Begin VB.TextBox txtVigInicio 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   900
            MaxLength       =   10
            TabIndex        =   7
            ToolTipText     =   "In�cio de vig�ncia (dd/mm/aaaa)."
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtVigTermino 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2670
            MaxLength       =   10
            TabIndex        =   8
            ToolTipText     =   "T�rmino de vig�ncia (dd/mm/aaaa)."
            Top             =   0
            Width           =   975
         End
         Begin Threed.SSCommand cmdPrazo 
            Height          =   240
            Left            =   4550
            TabIndex        =   59
            TabStop         =   0   'False
            ToolTipText     =   "Tipo do prazo."
            Top             =   30
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   8388608
            Picture         =   "T47V001.frx":0A76
         End
         Begin VB.Label lblMulti 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "T�rmino: "
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   6
            Left            =   1870
            TabIndex        =   61
            Top             =   0
            Width           =   795
         End
         Begin VB.Label lblMulti 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "In�cio: "
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   0
            TabIndex        =   60
            ToolTipText     =   "Data de in�cio da vig�ncia."
            Top             =   0
            Width           =   900
         End
         Begin VB.Label lblPrazo 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3650
            TabIndex        =   9
            ToolTipText     =   "Tempo de vig�ncia."
            Top             =   0
            Width           =   900
         End
      End
      Begin Threed.SSPanel panSegurado 
         Height          =   285
         Left            =   45
         TabIndex        =   63
         Top             =   30
         Width           =   4815
         _Version        =   65536
         _ExtentX        =   8493
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
         BevelOuter      =   0
         BevelInner      =   1
         Begin VB.TextBox txtSegurado 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   900
            MaxLength       =   60
            TabIndex        =   3
            ToolTipText     =   "Nome do segurado."
            Top             =   0
            Width           =   3905
         End
         Begin VB.Label lblMulti 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Segurado: "
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   0
            TabIndex        =   64
            ToolTipText     =   "Nome do segurado."
            Top             =   0
            Width           =   900
         End
      End
      Begin Threed.SSPanel panTipPessoa 
         Height          =   285
         Left            =   5950
         TabIndex        =   99
         Top             =   30
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2999
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
         Begin VB.OptionButton optTipPessoa 
            Caption         =   "&F�sica"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   20
            TabIndex        =   4
            ToolTipText     =   "Pessoa f�sica."
            Top             =   50
            Width           =   750
         End
         Begin VB.OptionButton optTipPessoa 
            Caption         =   "&Jur�dica"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   770
            TabIndex        =   5
            ToolTipText     =   "Pessoa jur�dica."
            Top             =   50
            Width           =   900
         End
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comiss�o: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   7650
         TabIndex        =   101
         ToolTipText     =   "Comiss�o."
         Top             =   30
         Width           =   1065
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pessoa: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   4860
         TabIndex        =   100
         ToolTipText     =   "Tipo de pessoa."
         Top             =   30
         Width           =   1095
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
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":0D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":13C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":16DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":19F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":1D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":202C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":2346
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":2798
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":2AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":2F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":321E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":3538
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":3852
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":3B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":3E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":41A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":48F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":4C0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":505E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":5378
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":57CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V001.frx":5C1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar barFerramentas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   75
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Novo"
            Object.ToolTipText     =   "Novo item."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Abrir"
            Object.ToolTipText     =   "Abrir itens."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar item."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir documento."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Visualizar"
            Object.ToolTipText     =   "Visualizar impress�o de documento."
            ImageIndex      =   18
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Arquivo"
            Object.ToolTipText     =   "Gerar documento em arquivo."
            ImageIndex      =   16
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir item."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copiar"
            Object.ToolTipText     =   "Copiar item."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Desfazer"
            Object.ToolTipText     =   "Desfazer digita��o."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Transmitir"
            Object.ToolTipText     =   "Transmitir documento"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bens"
            Object.ToolTipText     =   "Lista de bens"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Proposta"
            Object.ToolTipText     =   "Dados da proposta"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Recotacao"
            Object.ToolTipText     =   "Recota��o"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair do programa."
            ImageIndex      =   13
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1000
         EndProperty
      EndProperty
      Begin Threed.SSCommand cmdExtra 
         Height          =   240
         Index           =   2
         Left            =   8200
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   423
         _StockProps     =   78
         ForeColor       =   12582912
         Picture         =   "T47V001.frx":5F36
      End
      Begin Threed.SSCommand cmdExtra 
         Height          =   240
         Index           =   1
         Left            =   7950
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   423
         _StockProps     =   78
         ForeColor       =   12582912
         Picture         =   "T47V001.frx":6250
      End
   End
   Begin Threed.SSPanel panValores 
      Height          =   1245
      Left            =   0
      TabIndex        =   82
      Top             =   5070
      Visible         =   0   'False
      Width           =   9465
      _Version        =   65536
      _ExtentX        =   16695
      _ExtentY        =   2196
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
      BevelOuter      =   1
      Begin VB.TextBox TxtValor6 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7845
         MaxLength       =   14
         TabIndex        =   38
         ToolTipText     =   "Data in�cio viagem."
         Top             =   330
         Width           =   1575
      End
      Begin VB.TextBox TxtValor5 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7845
         MaxLength       =   14
         TabIndex        =   37
         ToolTipText     =   "C�mbio na emiss�o."
         Top             =   45
         Width           =   1575
      End
      Begin VB.TextBox TxtValor7 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7845
         MaxLength       =   14
         TabIndex        =   39
         ToolTipText     =   "Pr�mio Auto."
         Top             =   615
         Width           =   1575
      End
      Begin VB.TextBox TxtValor4 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4530
         MaxLength       =   14
         TabIndex        =   36
         ToolTipText     =   "Valor em Risco."
         Top             =   615
         Width           =   1575
      End
      Begin VB.TextBox TxtValor3 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4530
         MaxLength       =   14
         TabIndex        =   35
         ToolTipText     =   "Valor IS base."
         Top             =   330
         Width           =   1575
      End
      Begin VB.TextBox TxtValor2 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4530
         MaxLength       =   14
         TabIndex        =   34
         ToolTipText     =   "Coeficiente IS / VR."
         Top             =   60
         Width           =   1575
      End
      Begin VB.TextBox txtISBas 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1290
         MaxLength       =   14
         TabIndex        =   33
         ToolTipText     =   "Valor IS base."
         Top             =   615
         Width           =   1575
      End
      Begin VB.TextBox txtCoefISVR 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1290
         MaxLength       =   14
         TabIndex        =   32
         ToolTipText     =   "Coeficiente IS / VR."
         Top             =   320
         Width           =   1575
      End
      Begin VB.TextBox txtValoRisco 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1290
         MaxLength       =   14
         TabIndex        =   31
         ToolTipText     =   "Valor em Risco."
         Top             =   30
         Width           =   1575
      End
      Begin Threed.SSPanel panoptBem 
         Height          =   300
         Left            =   1320
         TabIndex        =   109
         Top             =   915
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   529
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
         BevelInner      =   1
         Alignment       =   6
         Begin VB.OptionButton optBens 
            Caption         =   "&N�o"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   650
            TabIndex        =   111
            ToolTipText     =   "N�O tem bem"
            Top             =   40
            Width           =   600
         End
         Begin VB.OptionButton optBens 
            Caption         =   "&Sim"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   50
            TabIndex        =   110
            ToolTipText     =   "SIM, tem bem"
            Top             =   40
            Width           =   600
         End
         Begin Threed.SSCommand cmdBens 
            Height          =   240
            Left            =   1300
            TabIndex        =   112
            TabStop         =   0   'False
            ToolTipText     =   "Exibir bem."
            Top             =   30
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
            Enabled         =   0   'False
            Picture         =   "T47V001.frx":656A
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bens? "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   30
         TabIndex        =   108
         Top             =   915
         Width           =   1275
      End
      Begin VB.Label lblValor6 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "In�cio da Viagem:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6075
         TabIndex        =   91
         Top             =   330
         Width           =   1755
      End
      Begin VB.Label lblValor5 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cambio na Emiss�o:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6075
         TabIndex        =   90
         Top             =   45
         Width           =   1755
      End
      Begin VB.Label lblValor7 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pr�mio Auto:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6075
         TabIndex        =   89
         Top             =   615
         Width           =   1755
      End
      Begin VB.Label lblValor3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pr�mio Tarif�rio:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2895
         TabIndex        =   88
         Top             =   330
         Width           =   1620
      End
      Begin VB.Label lblValor4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pr�mio Comercial:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2895
         TabIndex        =   87
         Top             =   615
         Width           =   1620
      End
      Begin VB.Label lblValor2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pr�mio Resseguro:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2895
         TabIndex        =   86
         Top             =   45
         Width           =   1620
      End
      Begin VB.Label lblCoefISVR 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LMG/VR:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   30
         TabIndex        =   85
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label lblValorRisco 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Val. em risco: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   30
         TabIndex        =   84
         Top             =   30
         Width           =   1275
      End
      Begin VB.Label lblISBas 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LMG. B�sico: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   30
         TabIndex        =   83
         Top             =   615
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Height          =   250
      Left            =   0
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   0
      Width           =   300
   End
   Begin Threed.SSPanel panBotao 
      Height          =   645
      Left            =   0
      TabIndex        =   97
      Top             =   6345
      Width           =   9480
      _Version        =   65536
      _ExtentX        =   16722
      _ExtentY        =   1138
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
      BevelOuter      =   0
      BevelInner      =   1
      Begin Threed.SSPanel panBotDadProposta 
         Height          =   555
         Left            =   1995
         TabIndex        =   58
         ToolTipText     =   "Exibir dados de proposta."
         Top             =   45
         Width           =   990
         _Version        =   65536
         _ExtentX        =   1746
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "F5 Dados de proposta"
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
         BevelWidth      =   3
         BorderWidth     =   0
      End
      Begin Threed.SSPanel panBotSair 
         Height          =   555
         Left            =   8730
         TabIndex        =   98
         ToolTipText     =   "Sair do sistema."
         Top             =   45
         Width           =   660
         _Version        =   65536
         _ExtentX        =   1164
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "Esc Sair"
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
         BevelWidth      =   3
         BorderWidth     =   0
      End
      Begin Threed.SSPanel panBotValidar 
         Height          =   555
         Left            =   60
         TabIndex        =   54
         ToolTipText     =   "Validar dados."
         Top             =   45
         Width           =   885
         _Version        =   65536
         _ExtentX        =   1561
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "F3    Validar"
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
         BevelWidth      =   3
         BorderWidth     =   0
      End
      Begin Threed.SSPanel panBotAltTipEmissao 
         Height          =   555
         Left            =   3015
         TabIndex        =   62
         ToolTipText     =   "Exibir tipo de emiss�o"
         Top             =   45
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "F6 Tipo de emiss�o"
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
         BevelWidth      =   3
         BorderWidth     =   0
      End
      Begin Threed.SSPanel panBotTransmitir 
         Height          =   555
         Left            =   7590
         TabIndex        =   80
         ToolTipText     =   "Transmiss�o"
         Top             =   45
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "Transmiss�o"
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
         BevelWidth      =   3
         BorderWidth     =   0
      End
      Begin Threed.SSPanel panBotEspecificar 
         Height          =   555
         Left            =   975
         TabIndex        =   113
         ToolTipText     =   "Exibir dados da Especifica��o."
         Top             =   45
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   979
         _StockProps     =   15
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
         BevelWidth      =   3
         BorderWidth     =   0
      End
      Begin Threed.SSPanel panBotConverter 
         Height          =   555
         Left            =   5250
         TabIndex        =   114
         ToolTipText     =   "Converter"
         Top             =   45
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   979
         _StockProps     =   15
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
         BevelWidth      =   3
         BorderWidth     =   0
         Enabled         =   0   'False
      End
      Begin Threed.SSPanel panBotCobertura 
         Height          =   555
         Left            =   4110
         TabIndex        =   115
         ToolTipText     =   "Exibir coberturas."
         Top             =   45
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "F7 Prolabore Agenciamento"
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
         BevelWidth      =   3
         BorderWidth     =   0
      End
      Begin Threed.SSPanel panBotObservacoes 
         Height          =   555
         Left            =   6450
         TabIndex        =   116
         ToolTipText     =   "Observa��es"
         Top             =   45
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   979
         _StockProps     =   15
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
         BevelWidth      =   3
         BorderWidth     =   0
         Enabled         =   0   'False
      End
   End
   Begin Threed.SSPanel panLocRisco 
      Height          =   960
      Left            =   0
      TabIndex        =   65
      Top             =   2475
      Visible         =   0   'False
      Width           =   9480
      _Version        =   65536
      _ExtentX        =   16722
      _ExtentY        =   1693
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
      BevelOuter      =   0
      BevelInner      =   1
      Begin VB.TextBox txtNumItem 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   8715
         MaxLength       =   5
         TabIndex        =   20
         ToolTipText     =   "N�mero do item."
         Top             =   30
         Width           =   735
      End
      Begin VB.TextBox txtRisCidade 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1430
         MaxLength       =   30
         TabIndex        =   29
         ToolTipText     =   "Cidade."
         Top             =   610
         Width           =   2640
      End
      Begin VB.TextBox txtRisComplemento 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3450
         MaxLength       =   20
         TabIndex        =   22
         ToolTipText     =   "Complemento."
         Top             =   320
         Width           =   2295
      End
      Begin VB.TextBox txtRisCEP 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   17
         ToolTipText     =   "CEP de risco. (xxxxx-xxx)"
         Top             =   30
         Width           =   915
      End
      Begin VB.TextBox txtRisNomLogradouro 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3940
         MaxLength       =   50
         TabIndex        =   19
         ToolTipText     =   "Nome do logradouro do endere�o do segurado."
         Top             =   30
         Width           =   3930
      End
      Begin VB.TextBox txtRisTipLogradouro 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3450
         MaxLength       =   3
         TabIndex        =   18
         ToolTipText     =   "Tipo do logradouro do endere�o do segurado (Rua, Av., etc.)."
         Top             =   30
         Width           =   495
      End
      Begin VB.ComboBox cboRisUF 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T47V001.frx":6884
         Left            =   4710
         List            =   "T47V001.frx":68D9
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   30
         ToolTipText     =   "Unidade da federa��o do local de risco."
         Top             =   610
         Width           =   690
      End
      Begin VB.TextBox txtRisBairro 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6675
         MaxLength       =   30
         TabIndex        =   23
         ToolTipText     =   "Bairro do local de risco."
         Top             =   320
         Width           =   2760
      End
      Begin VB.TextBox txtRisNumero 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1430
         MaxLength       =   8
         TabIndex        =   21
         ToolTipText     =   "N�mero do logradouro "
         Top             =   315
         Width           =   915
      End
      Begin Threed.SSCommand cmdLocalRisco 
         Height          =   880
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   1552
         _StockProps     =   78
         Enabled         =   0   'False
         Picture         =   "T47V001.frx":6949
      End
      Begin Threed.SSCommand cmdCepRisco 
         Height          =   285
         Left            =   330
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   503
         _StockProps     =   78
         ForeColor       =   8388608
         Picture         =   "T47V001.frx":837B
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   9
         Left            =   7890
         TabIndex        =   104
         Top             =   30
         Width           =   825
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cidade: "
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   14
         Left            =   330
         TabIndex        =   73
         Top             =   615
         Width           =   1095
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Complemento: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   12
         Left            =   2350
         TabIndex        =   72
         Top             =   315
         Width           =   1100
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bairro: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   13
         Left            =   5760
         TabIndex        =   71
         Top             =   315
         Width           =   900
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CEP: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   10
         Left            =   330
         TabIndex        =   70
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label lblCEPSave 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   600
         TabIndex        =   69
         Top             =   30
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Logradouro: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   11
         Left            =   2350
         TabIndex        =   68
         ToolTipText     =   "Logradouro do endere�o do segurado."
         Top             =   30
         Width           =   1100
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UF: "
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   15
         Left            =   4080
         TabIndex        =   67
         ToolTipText     =   "Unidade da federa��o do endere�o do segurado."
         Top             =   615
         Width           =   615
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N�mero:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   20
         Left            =   330
         TabIndex        =   66
         Top             =   315
         Width           =   1095
      End
   End
   Begin Threed.SSPanel panAguarde 
      Height          =   780
      Left            =   240
      TabIndex        =   74
      Top             =   2265
      Visible         =   0   'False
      Width           =   9195
      _Version        =   65536
      _ExtentX        =   16219
      _ExtentY        =   1376
      _StockProps     =   15
      ForeColor       =   12582912
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BevelInner      =   1
      Outline         =   -1  'True
   End
   Begin Threed.SSPanel panPesquisa 
      Height          =   780
      Left            =   0
      TabIndex        =   50
      Top             =   450
      Width           =   9480
      _Version        =   65536
      _ExtentX        =   16722
      _ExtentY        =   1376
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
      BevelOuter      =   0
      BevelInner      =   1
      Begin VB.TextBox txtDataRecebimento 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7935
         MaxLength       =   10
         TabIndex        =   117
         ToolTipText     =   "N�mero do PI"
         Top             =   60
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox txtOrdem 
         Height          =   285
         Left            =   6930
         MaxLength       =   10
         TabIndex        =   45
         Top             =   660
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtApolice 
         Height          =   285
         Left            =   4335
         MaxLength       =   10
         TabIndex        =   42
         ToolTipText     =   "N�mero da Ap�lice"
         Top             =   45
         Width           =   1185
      End
      Begin VB.TextBox txtEndosso 
         Height          =   285
         Left            =   6780
         MaxLength       =   6
         TabIndex        =   43
         ToolTipText     =   "N�mero do Endosso"
         Top             =   45
         Width           =   1050
      End
      Begin VB.TextBox txtFatura 
         Height          =   285
         Left            =   5385
         MaxLength       =   2
         TabIndex        =   44
         Top             =   660
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtNumPI 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "N�mero do PI"
         Top             =   45
         Width           =   1425
      End
      Begin VB.CheckBox chkLider 
         Caption         =   "Yasuda"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1980
         TabIndex        =   41
         Top             =   720
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtLider 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   945
         MaxLength       =   10
         TabIndex        =   40
         ToolTipText     =   "In�cio de vig�ncia (dd/mm/aaaa)."
         Top             =   675
         Visible         =   0   'False
         Width           =   975
      End
      Begin Threed.SSPanel panConfirmar 
         Height          =   435
         Left            =   7890
         TabIndex        =   2
         ToolTipText     =   "Confirma o tipo de ramo."
         Top             =   255
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   767
         _StockProps     =   15
         Caption         =   "Confirmar"
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
         BevelWidth      =   2
         BorderWidth     =   0
         BevelOuter      =   1
         BevelInner      =   2
      End
      Begin VB.ComboBox cboRamo 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T47V001.frx":8695
         Left            =   945
         List            =   "T47V001.frx":8697
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   345
         Width           =   6885
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordem:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   24
         Left            =   6135
         TabIndex        =   95
         ToolTipText     =   "Tipo de ramo."
         Top             =   660
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fatura: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   23
         Left            =   4455
         TabIndex        =   94
         ToolTipText     =   "Tipo de ramo."
         Top             =   660
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Endosso:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   22
         Left            =   5565
         TabIndex        =   93
         ToolTipText     =   "Tipo de ramo."
         Top             =   45
         Width           =   1215
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ap�lice: "
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   18
         Left            =   3420
         TabIndex        =   92
         ToolTipText     =   "Tipo de ramo."
         Top             =   45
         Width           =   900
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Protocolo Interno:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   17
         Left            =   45
         TabIndex        =   79
         Top             =   45
         Width           =   1440
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L�der:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   16
         Left            =   30
         TabIndex        =   78
         ToolTipText     =   "Data de in�cio da vig�ncia."
         Top             =   675
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ramo: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   30
         TabIndex        =   51
         ToolTipText     =   "Tipo de ramo."
         Top             =   345
         Width           =   900
      End
   End
   Begin VB.Label LblTamanho 
      AutoSize        =   -1  'True
      Caption         =   "LblTamanho"
      Height          =   195
      Left            =   240
      TabIndex        =   81
      Top             =   5760
      Visible         =   0   'False
      Width           =   870
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgRamo 
      Height          =   255
      Left            =   0
      Top             =   430
      Width           =   9015
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuArqNovaConsultaPI 
         Caption         =   "Nova consulta PI"
      End
      Begin VB.Menu mnuArqNovItem 
         Caption         =   "&Novo item"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuArqAbrItem 
         Caption         =   "&Abrir item"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArqSalvar 
         Caption         =   "&Salvar"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArqImprimir 
         Caption         =   "Imprimir"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuArqVisualizar 
         Caption         =   "Visualizar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuArqExcProposta 
         Caption         =   "E&xcluir proposta"
      End
      Begin VB.Menu mnuSep03 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuArqSaiItem 
         Caption         =   "Sai&r"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "&Editar"
      Begin VB.Menu mnuEdiDesfazer 
         Caption         =   "&Desfazer"
      End
      Begin VB.Menu mnuSep06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdiCalPremios 
         Caption         =   "&Validar"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdiCopItem 
         Caption         =   "&Copiar item"
      End
      Begin VB.Menu mnuEdiExcItem 
         Caption         =   "&Excluir item"
      End
      Begin VB.Menu mnuSep08 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdiBusProduto 
         Caption         =   "Buscar pr&oduto"
      End
      Begin VB.Menu nuSep010 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEdiEspecificar 
         Caption         =   "Especificac�o"
         Shortcut        =   {F4}
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuExibir 
      Caption         =   "E&xibir"
      Begin VB.Menu mnuExiBem 
         Caption         =   "&Bem"
      End
      Begin VB.Menu mnuExiBeneficiarios 
         Caption         =   "Bene&fici�rios"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExiCustoEmis 
         Caption         =   "C&usto de emiss�o"
      End
      Begin VB.Menu mnuExiDadProposta 
         Caption         =   "&Dados da proposta"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuExiJuros 
         Caption         =   "&Juros"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExiObservacoes 
         Caption         =   "Obser&va��es"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExiOutrosSeg 
         Caption         =   "&Outros seguros"
         Shortcut        =   {F8}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExiPrazo 
         Caption         =   "Pra&zo"
      End
      Begin VB.Menu mnuExiProtecao 
         Caption         =   "&Prote��o"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExiSeguranca 
         Caption         =   "&Seguran�a"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExiTipoEmissao 
         Caption         =   "&Tipo de emiss�o"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuTextoEndosso 
         Caption         =   "Texto para &Endosso"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuExiAjuste 
         Caption         =   "Ajusta Pr�mio"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExiClausulas 
         Caption         =   "C&l�usulas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExiEspecAnt 
         Caption         =   "Especifica��o Antiga"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuExiProAgenc 
         Caption         =   "Prolabori e Agenciamento"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "Aj&uda"
      Begin VB.Menu mnuAjuSYAS 
         Caption         =   "Aj&uda do SYAS"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAjuSobre 
         Caption         =   "So&bre o SYAS"
      End
   End
   Begin VB.Menu mnuSituacao 
      Caption         =   "Situa��o"
   End
End
Attribute VB_Name = "frmT47V001A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mLinhaFocus         As Integer
Dim mColunaFocus        As Integer
Dim GridFocado          As String
Dim EstaCarregando      As Boolean
Dim mAntCEP             As String
Dim EstaAlterando       As Boolean
Dim PosicaoAnterior     As Boolean
Dim mComUltData         As Double
Dim mComUltUsuario      As String
Public objstcPedido     As stcA46V702B
Private objstcPedLoc    As stcA46V708B
Private objPedido       As clsA46V702A
Private mInclusao       As Boolean
Private mFlag_Emitir    As Boolean
Private mFlag_Suspender As Boolean
Public mTipoEmissao     As Integer
Public lVerTipEmissao   As String           'Tipo de emiss�o informado no P46V101
Public mCarregarPI      As Boolean
Public mCarregarDatas   As Boolean


Private Sub mpVerComum()
    'Procedure: consiste dados comuns
    
    Dim lTipPessoa          As Byte             'Tipo da pessoa: 1 = pessoa jur�dica e 2 = pessoa f�sica.
    Dim lQutdeParcelas      As Integer          'Quantidade de parcelas
    Dim lI                  As Integer          'Utilizado no for...next
    Dim lMaxParcela         As Integer          'N�mero de parcelas m�ximo permitido.
    Dim clsA46V102          As clsA46V102A      'DLL : consiste dados comuns (n�vel 2).
    Dim lObjstcVenc         As stcA46V713B      'Objeto da classe de estrutura stcA46V713B (TAB_PED_Venc)
    Dim lObjstcCorr         As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_Corr)
    Dim lobjstcCoss         As stcA46V706B      'Objeto da classe de estrutura stcA46V706B (TAB_PED_Coss)
    Dim lobjstcItem         As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcCobert       As stcA46V704B      'Objeto da classe (stc) de coberturas.
    Dim lCosseguros(5, 2)   As String           'Cosseguros.
    Dim lCod_Retorno        As Integer
    Dim clsA99V001          As clsA99V001
    Dim lMensagem           As String
    Dim lDigito             As String
    Dim lAuxiliar           As String
    Dim lPercAux            As Single
    Dim lComissao           As Single
    Dim lValLMG             As Long
    Dim lCodCorr            As Long
    Dim lCalcDat1           As Date
    Dim lCalcDat2           As Long
    Dim lBlnSiscad          As Boolean      'Indicador se existe dados no Siscad
    Dim lValidade           As Long
    Dim txtobs              As String
    Dim lClassif            As String
    
    gM46V999.gpInicializar_gErros
    lPercAux = 0
    lCodCorr = 0
    lComissao = 0
  
    If objstcPedido.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = objstcPedido.ColecaoPedCorr.Item(1)
        lPercAux = lObjstcCorr.Per_Corr
        lCodCorr = lObjstcCorr.Cod_Corr
    End If
    
    Dim lPlanoComercial    As Boolean
    Dim lPlanoNaoComercial As Boolean
    
    For Each lobjstcItem In objstcPedido.ColecaoPedLoc
        'Ficha 1000308 - Permitir itens com planos diferentes no produto Roubo (c�lculo coletivo)
        'Roubo n�o permite c�lculo de itens (Riscos Comerciais e/ou Industriais) com os demais
        If lobjstcItem.Cod_Ramo = e_CodRamo_Roubos Then
             If lobjstcItem.Tab_Clasf = 2 Then lPlanoComercial = True
             If lobjstcItem.Tab_Clasf <> 2 Then lPlanoNaoComercial = True
        End If
        For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
            If lobjstcCobert.Cod_Cobert = 1 Then
                lValLMG = lobjstcCobert.Val_IS
                Exit For
            End If
        Next lobjstcCobert
    Next lobjstcItem
    
    'Ficha 1000308 - Permitir itens com planos diferentes no produto Roubo (c�lculo coletivo)
    'Roubo n�o permite c�lculo de itens (Riscos Comerciais e/ou Industriais) com os demais
    If lPlanoComercial = True And lPlanoNaoComercial = True Then
        Call gM46V999.gpMenErro("FGE0050 - N�o � permitido c�lculo de itens (Riscos Comerciais e/ou Industriais) com os demais." & lMensagem)
    End If

'Ficha 900653
    Dim lString1, lString2, lPosicao

    lString1 = objstcPedido.Nom_Segurado
    lString2 = "ESPOLIO"

    lPosicao = InStr(lString1, lString2)

    If (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) And lPosicao > 0 Then
        Call gM46V999.gpMenErro("FGE0050 - Palavra 'ESPOLIO' n�o permitida no nome do Segurado" & lMensagem)
    End If
    
'Ficha 901782
    If UCase(objstcPedido.Nom_Segurado) Like "*E/OU*" Then
       Call gM46V999.gpMenErro("FGE0052 - Nome inv�lido (n�o utilizar E/OU)." & lMensagem)
    End If
    
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And gM46V999.gTipoAcesso = CALCULO Then
        If Not gM46V999.gfPreenchido(objstcPedido.objstcCtrlEmis.Num_Cotac) Or objstcPedido.objstcCtrlEmis.Num_Cotac = 0 Then
           If gM46O001.gPermiteSemSiscota = False Then
                Call gM46V999.gpMenErro("FGE0088 - N�mero do Siscota inv�lido (n�o informado)." & lMensagem)
           End If
        End If
    End If

    Dim lStatusSiscota  As Integer
    Dim lRamoSiscota    As Long
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And gM46V999.gTipoAcesso = CALCULO Then
        If gM46V999.gfPreenchido(objstcPedido.objstcCtrlEmis.Num_Cotac) And objstcPedido.objstcCtrlEmis.Num_Cotac > 0 Then
            Call gM46V111.gfVerStatusSiscota(objstcPedido, lStatusSiscota, lRamoSiscota)
            If gM46O001.gPermiteSemSiscota = False And (lStatusSiscota < 12 Or lStatusSiscota > 13) Then
                Call gM46V999.gpMenErro("FGE0089 - N�mero do Siscota inv�lido (a cota��o n�o est� pendente)." & lMensagem)
            End If
            If gM46O001.gPermiteSemSiscota = False And lRamoSiscota <> objstcPedido.Cod_Ramo Then
                Call gM46V999.gpMenErro("FGE0089 - N�mero do Siscota inv�lido (a cota��o n�o pertence ao mesmo ramo)." & lMensagem)
            End If
        End If
    End If

    If lPercAux = 0 Then
        Call gM46V999.gpMenErro("FNP1010 - Comiss�o deve estar preenchida!")
    End If
    
    If objstcPedido.ValTotalAgenciamento > 100 Then
        Call gM46V999.gpMenErro("FNP1010 - Total Comiss�o Agenciamento n�o pode ser mais do que 100%!")
    End If
    
    If objstcPedido.ValTotalProlabore > 100 Then
        Call gM46V999.gpMenErro("FNP1010 - Total Comiss�o Prolabore n�o pode ser mais do que 100%")
    End If
    
    If gM46V999.gfVerEtiqueta(lCodCorr, objstcPedido.Cod_Reg, objstcPedido.Tip_Pessoa, objstcPedido.objstcCtrlEmis.Cod_Envi_Doc, lMensagem) Then
        Call gM46V999.gpMenErro("FGE0019 - " & lMensagem)
    Else
        'Ficha:0801295
        'Novas op��es "Tipo de Emiss�o=Conforme cadastro":
        '"O" -> Conforme Cadastro Corretor e "E" -> Conforme Cadastro Segurado
        'a defini��o se "O" ou "E" � feita dentro da gfVerEtiqueta
        If lMensagem = "O" Then '<- digitacao
            objstcPedido.objstcCtrlEmis.Cod_Envi_Doc = "O"
        ElseIf lMensagem = "E" Then
            objstcPedido.objstcCtrlEmis.Cod_Envi_Doc = "E"
        End If
    End If
        
    'Ficha:900048 Isen��o de IOF para segurados:
    '-Contendo as palavras-chaves no nome;
    '-Ramo de atividade -> Administra��o publica, defesa e seguridade social;
    '-C�digo de registro 1.
    '--------------------------------------------------------------------------------------------------------------------------------------------
    Dim X As Long
    Dim DadosChave
    
    DadosChave = Split(UCase(Trim("Estado,Distrito,Federal,Estadual,Municipal,Munic�pio,Municipio,Prefeitura,Governo,P�blic,Public,Secretaria,Minist�rio,Ministerio")), ",")
    For X = 0 To UBound(DadosChave)
    
        If InStr(1, UCase(Trim(objstcPedido.Nom_Segurado)), DadosChave(X)) > 0 _
        And UCase(Trim(objstcPedido.Cod_Ramo_Ativ)) = 12 _
        And objstcPedido.Cod_Reg = 1 Then
            If MsgBox("Emitir com isen��o de IOF?", vbYesNo, App.Title) = vbYes Then
            
                frmT47V016A.cboCodReg.ListIndex = 2 '                         <-Reposicionar controle
                objstcPedido.Cod_Reg = Mid(frmT47V016A.cboCodReg.Text, 1, 3) '<-Armazenar na cole��o
            End If
            Exit For
        End If
    Next X
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    If Right(objstcPedido.Dat_Ini_Vig, 4) = "0229" Then
          Call gM46V999.gpMenErro("FGE0046 - Vig�ncia inv�lida (data de in�cio com ano " & _
                                                "bissexto - utilizar 28/02 ou 01/03).")
                                                
    End If
    
    If objstcPedido.Dat_Ini_Vig = 0 Then
          Call gM46V999.gpMenErro("FGE0017 - " & gM46V999.gPrefixo & _
                         "Vig�ncia inv�lida (data de in�cio de vig�ncia inv�lida).")
    End If
  
    If objstcPedido.Dat_Ter_Vig = 0 Then
          Call gM46V999.gpMenErro("FGE0017 - " & gM46V999.gPrefixo & _
                         "Vig�ncia inv�lida (data de t�rmino de vig�ncia inv�lida).")
    End If
    
    Dim mDatTerVig  As Date
    Dim lAuxMsg As String
    Dim lobjPedObs          As clsA46V714A      'Objeto da classe de observa��o.
    'Consiste vig�ncia e renova��o
    Select Case objstcPedLoc.Tip_Emissao
        Case e_TipEmissao_RenYasuda
            If objPedido.mfRetornaTerminoVigencia(objstcPedido) <> 0 Then
                mDatTerVig = gM46V999.gfForData(2, objPedido.mfRetornaTerminoVigencia(objstcPedido))
                If gM46V999.gfForData(1, mDatTerVig + 30) < objstcPedido.Dat_Ini_Vig Then
                    Call gM46V999.gpMenErro("FGE0111 - In�cio de vig�ncia superior a 30 dias " & _
                    "do final de vig�ncia da ap�lice anterior! Favor calcular seguro novo.")
                End If
            End If
        Case e_TipEmissao_RenCongenere
            lAuxMsg = "� obrigat�rio o envio de c�pia da ap�lice da cong�nere."
            lAuxMsg = lAuxMsg & "Caso o seguro tenha vencido a mais de 30 dias ser� classificado como Seguro Novo, sem concess�o de b�nus."
            lAuxMsg = gM46V999.gfRetirarCaracterEspecial(lAuxMsg, True)
            
           
            'Verifico se deve ser gravada a observa��o de renova��o.
            objstcPedLoc.objstcPedObs.Observacao = Replace(lAuxMsg, vbCrLf, "&%")
            Set lobjPedObs = New clsA46V714A
            Set lobjPedObs.M46V999 = gM46V999
            
            Set lobjPedObs.ConexaoBD = objPedido.ConexaoBD
            If Not lobjPedObs.mfExcluir(objstcPedLoc.objstcPedObs, "", "") Then
                lMensagem = lobjPedObs.UltimoErro
                Call gM46V999.gpMenErro("FNP1002 - " & lMensagem)
            End If
            With objstcPedLoc.objstcPedObs
                .NOSSO_NUMERO = objstcPedLoc.NOSSO_NUMERO
                .Num_Item = objstcPedLoc.Num_Item
            End With
            If Not lobjPedObs.mfIncluir(objstcPedLoc.objstcPedObs, "", "", "") Then
                lMensagem = lobjPedObs.UltimoErro
                Call gM46V999.gpMenErro("FNP1001 - " & lMensagem)
            End If
    End Select
            
    '2. Consiste dados comuns - n�vel 2.
    Set clsA46V102 = New clsA46V102A    'DLL : consiste dados comuns (n�vel 2).
    Set clsA46V102.M46V999 = gM46V999
    With clsA46V102
        'a) Prepara dados para consist�ncia de comum n�vel 2.
        .Pro_CarBanco = objstcPedido.Cod_Bco_Cobr
        'Obs.: O campo Ven_Parcela1_VM cont�m a data de vencimento da primeira parcela do carn�.
        If objstcPedido.ColecaoPedVenc.Count > 0 Then
            Set lObjstcVenc = objstcPedido.ColecaoPedVenc.Item(1)
            .Pro_CarDatPriParcela = lObjstcVenc.DAT_VENC
        End If
        .Pro_Cob_Bairro = objstcPedido.Nom_Bairro
        .Pro_Cob_CEP = Format(objstcPedido.Num_CEP, "00000000")
        .Pro_Cob_Cidade = objstcPedido.Nom_Cid
        If gM46V999.gfPreenchido(objstcPedido.Dsc_Compl_Logr) Then
            .Pro_Cob_Complemento = objstcPedido.Dsc_Compl_Logr
        Else
            .Pro_Cob_Complemento = ""
        End If
        .Pro_Cob_NomLogradouro = objstcPedido.Nom_Logr
        .Pro_Cob_TipLogradouro = objstcPedido.Cod_Tip_Logr
        .Pro_Cob_UF = objstcPedido.Sig_UF
        If objstcPedido.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = objstcPedido.ColecaoPedCorr.Item(1)
            .Pro_CodCorretor = lObjstcCorr.Cod_Corr
        Else
            .Pro_CodCorretor = 0
        End If
        lCosseguros(1, 1) = 0
        lCosseguros(1, 2) = 0
        lCosseguros(2, 1) = 0
        lCosseguros(2, 2) = 0
        lCosseguros(3, 1) = 0
        lCosseguros(3, 2) = 0
        lCosseguros(4, 1) = 0
        lCosseguros(4, 2) = 0
        lCosseguros(5, 1) = 0
        lCosseguros(5, 2) = 0
        lI = 0
        For Each lobjstcCoss In objstcPedido.ColecaoPedCoss
            lI = lI + 1
            lCosseguros(lI, 1) = lobjstcCoss.Cod_Coss
            lCosseguros(lI, 2) = lobjstcCoss.Per_Coss
        Next
        .Pro_Cosseguros = lCosseguros
        .Pro_DebAgencia = objstcPedido.gAgenciaDig
        .Pro_DebBanco = objstcPedido.Cod_Bco_Cobr
        .Pro_DebConta = objstcPedido.Num_Cta_Bas_Cor
        'Verificar se consegue fazer a consist�ncia,
        'pois para Yasuda foi gravado sem digito
        If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
            .Pro_DebConta = .Pro_DebConta & Right(frmT47V010A.txtDebConta, 1)
        End If
        .Pro_End_Bairro = objstcPedido.objstcPedEndereco.Nom_Bairro
        .Pro_End_CEP = Format(objstcPedido.objstcPedEndereco.Num_CEP, "00000000")
        .Pro_End_Cidade = objstcPedido.objstcPedEndereco.Nom_Cid
        If gM46V999.gfPreenchido(objstcPedido.objstcPedEndereco.Dsc_Compl_Logr) Then
            .Pro_End_Complemento = objstcPedido.objstcPedEndereco.Dsc_Compl_Logr
        Else
            .Pro_End_Complemento = ""
        End If
        .Pro_End_NomLogradouro = objstcPedido.objstcPedEndereco.Nom_Logr
        .Pro_End_TipLogradouro = objstcPedido.objstcPedEndereco.Cod_Tip_Logr
        .Pro_End_UF = objstcPedido.objstcPedEndereco.Sig_UF
        'Proposta: forma de pagamento.
        '1 = carn� e 2 = d�bito em conta.
        .Pro_ForPagamento = IIf(frmT47V004A.optForPagamento(0).Value = True, 1, 2)
        .Pro_IndAgencia = objstcPedido.gAgenciaDig
        .Pro_IndBanco = objstcPedido.Cod_Bco_Cobr
        .Pro_IndConta = objstcPedido.Num_Cta_Bas_Cor
        .Pro_IndCobranca = IIf(frmT47V004A.optEndCobranca(0).Value = True, 1, 2)
        .Pro_IndCosseguros = IIf(objstcPedido.ColecaoPedCoss.Count > 0, 1, 2)
        .Pro_NomSegurado = objstcPedido.Nom_Segurado
        .Pro_PagAntBloqueto = objstcPedido.objstcCtrlEmis.Num_Bloq
        If frmT47V004A.optPagAntecipado(1).Value Then
            .Pro_PagAntecipado = 1
        ElseIf frmT47V004A.optPagAntecipado(2).Value Then
            .Pro_PagAntecipado = 2
        Else
            .Pro_PagAntecipado = 0
        End If
        .Pro_QtdParcelas = objstcPedido.Qtd_Parc
        .Pro_UniProdutiva = objstcPedido.Cod_Unid_Prod
        .Pro_VigInicio = objstcPedido.Dat_Ini_Vig
        .Pro_TipPessoa = objstcPedido.Tip_Pessoa
        .Pro_CNPJ_CPF = objstcPedido.Num_Cgc_Cpf
        .Pro_CodSegurado = objstcPedido.Cod_Segurado
        'b) Executa consist�ncia.
        On Error Resume Next
        Call .VER_COMUM_NIVEL_2
        If Err <> 0 Then
            Call gM46V999.gpGraLog(2, "FTE0009 - " & gM46V999.gPrefixo & "Erro ao consistir dados comuns n�vel 2 (" & _
                             Err & " - " & Error & ").")
            On Error GoTo 0
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        On Error GoTo 0
        
        Dim lZerar_campoCodigo As Boolean: lZerar_campoCodigo = False
        If Left(objstcPedido.Cod_Segurado, 4) = 7777 And objstcPedido.objstcCtrlEmis.Cod_Corr <> 1601 Then
            lZerar_campoCodigo = True
        Else
            Dim bdGed As ADODB.Connection
            Dim lrst As clsYasRecordSet
            Dim lSelect As String
            Dim lCpfCnpj As String
        
            '[Verifica se PF/PJ]
            lCpfCnpj = objstcPedido.Num_Cgc_Cpf
                    
            '[Abrir base GED.tab_cliente]
            If Not gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, bdGed, lMensagem) Then Call gM46V999.gpGraLog(1, lMensagem)
                    
            '[Verificar se CPF/CNPJ dif. de 0 e C�d.= 0]
            If lCpfCnpj <> "0" And objstcPedido.Cod_Segurado = 0 Then

              '[Obter os C�digos cadastrados na GED de acordo com o CPF/CNPJ digitado]
                lSelect = "SELECT COD_CLI FROM tab_cliente where CGC_CPF='" & Trim(lCpfCnpj) & "'"
                If Not gM46V999.gfObtRegistro(bdGed, lSelect, lrst, lMensagem) Then Call gM46V999.gpGraLog(1, lMensagem)
        
                '[Caso v�rios c�digos retornados, selecionar o primeiro]
                If Not (lrst.EOF And lrst.BOF) Then
                    lrst.MoveLast: lrst.MoveFirst
                    If lrst.RecordCount = 1 Then
                        objstcPedido.Cod_Segurado = IIf(Len(lrst("Cod_Cli")) > 6, 0, Val(lrst("Cod_Cli")))
                    End If
                End If
              
            '[Verificar se c�d. dif. de 0 no campo]
            ElseIf objstcPedido.Cod_Segurado <> 0 Then
            
                '[Obter os CPFs/CNPJs cadastrados na GED com o c�digo digitado]
                lSelect = "SELECT * FROM tab_cliente where COD_CLI=" & objstcPedido.Cod_Segurado
                If Not gM46V999.gfObtRegistro(bdGed, lSelect, lrst, lMensagem) Then Call gM46V999.gpGraLog(1, lMensagem)
                                
                '[C�digo inv�lido:]
                If lrst.EOF And lrst.BOF Then
                    lZerar_campoCodigo = True
                End If
            End If
        End If
            
        If lZerar_campoCodigo Then
            objstcPedido.Cod_Segurado = 0
        End If
        '[Limpar Mem�ria]
        lSelect = ""
        Set lrst = Nothing
        Set bdGed = Nothing
    End With
    Set clsA46V102 = Nothing
    
    '3. Consiste dados de identifica��o do proponente.
    If optTipPessoa(0).Value = True Then
        lTipPessoa = 2  'Pessoa f�sica.
    Else
        lTipPessoa = 1  'Pessoa jur�dica.
    End If
    'PE 0401710 - Mensagem de erro quando n�o informado dados SUSEP.
    Select Case lTipPessoa
        Case 1  'Pessoa jur�dica.
            If frmT47V004A.cboRamAtividade.ListIndex = -1 Then
                Call gM46V999.gpMenErro("FNP0001 - " & gM46V999.gPrefixo & "Ramo de atividade do proponente n�o " & _
                               "selecionado (F5 - dados proposta).")
            End If
            If Not gM46V999.gfPreenchido(frmT47V004A.txtCNPJ.Text) Then
                    Call gM46V999.gpMenErro("FNP0002 - CNPJ do segurado n�o preenchido." & _
                                   "(F5 - dados proposta).")
            End If
        Case 2  'Pessoa f�sica.
            With frmT47V004A
                'a) Tipo de documento de identifica��o.
                If .cboIdeTipDocumento.ListIndex = -1 Then
                    Call gM46V999.gpMenErro("FNP0003 - Tipo do documento do proponente n�o selecionado " & _
                                   "(F5 - dados proposta).")
                End If
                'b) N�mero do documento.
                If Not gM46V999.gfPreenchido(.txtIdeNumDocumento.Text) Then
                    Call gM46V999.gpMenErro("FNP0004 - N�mero do documento do proponente n�o preenchido " & _
                                   "(F5 - dados proposta).")
                End If
                
                If Not gM46V999.gfPreenchido(.txtCPF.Text) Then
                    Call gM46V999.gpMenErro("FNP0005 - CPF do segurado n�o preenchido." & _
                                   "(F5 - dados proposta).")
                End If
                
                If .optSexo(0).Value = False And .optSexo(1).Value = False Then
                    Call gM46V999.gpMenErro("FNP0006 - Indicador de sexo do segurado n�o selecionado  (F5 - dados proposta).")
                                   
                End If
            End With
    End Select
    If objstcPedido.Qtd_Parc = 0 And _
      (objstcPedido.Cod_Reg <> 106 And objstcPedido.Cod_Reg <> 107 And objstcPedido.Cod_Reg <> 120 And _
       objstcPedido.Cod_Reg <> 121 And objstcPedido.Cod_Reg <> 122 And objstcPedido.Cod_Reg <> 20) Then
        Call gM46V999.gpMenErro("FGE0111 - Quantidade de parcelas inv�lida (n�o informada)")
    End If
    If gM46V999.gfDisponivel("frmT47V011A") Then
        If gM46V999.gfPreenchido(frmT47V011A.txtPagAntBloqueto.Text) Then
            lAuxiliar = Replace(frmT47V011A.txtPagAntBloqueto.Text, "-", "")
            If Val(lAuxiliar) > 0 Then
                Set clsA99V001 = New clsA99V001
                Set clsA99V001.M46V999 = gM46V999
                
                Call clsA99V001.DIG_BLOQUETO(Format(lAuxiliar, "000000000"), lDigito, lMensagem)
                If Right(lAuxiliar, 1) <> lDigito Then
                    Call gM46V999.gpMenErro("FGE0059 - Pagamento antecipado com bloqueto - n�mero do bloqueto com d�gito inv�lido.")
                End If
            End If
        End If
    End If
    Dim lLimDias As Integer
    lLimDias = 31
    
    Select Case Left(frmT47V004A.txtCodSegurado.Text, 6)
       Case 777701, 777702, 777704, 777705, 777706, 777707, 777709, 777710, 777711, 777712, 777713, 777714, 777716, 777733, 777787, 777788
           lLimDias = 60
    End Select
    
'Ficha 901170
    Select Case Val(frmT47V004A.txtCodSegurado.Text)
       Case 126, 503863
'               lLimDias = 60
'Ficha 901651
            lLimDias = 90
    End Select

    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
'Ficha 900745
        lLimDias = 99
    End If
    
    If frmT47V004A.optVencimento(0).Value = True Then
        If frmT47v019A.txtDtVenc(1).Text <> "" Then
            For lI = 1 To Val(frmT47V004A.txtQuaParcelas.Text)
                With frmT47v019A
                    If lI = 1 Then
                        If CDate(.txtDtVenc(lI).Text) > CDate(Date + lLimDias) Then
                            Call gM46V999.gpMenErro("FGE0111 - Data de vencimento da parcela " & lI & " n�o pode ser superior � " & lLimDias & " dias (" & .txtDtVenc(lI).Text & ").")
                        End If
                    End If
                    If Not gM46V999.gfPreenchido(.txtDtVenc(lI).Text) Then
                        Call gM46V999.gpMenErro("FGE0112 - Data de vencimento da parcela " & lI & " n�o preenchida.")
                    Else
                        If Not IsDate(.txtDtVenc(lI).Text) Then
                            Call gM46V999.gpMenErro("FGE0006 - Data de vencimento da parcela " & lI & " inv�lida (" & .txtDtVenc(lI).Text & ").")
                        End If
                        If CDate(.txtDtVenc(lI).Text) < CDate(Date + 8) Then
                            Call gM46V999.gpMenErro("FGE0007 - Data de vencimento da parcela " & lI & " inv�lida (inferior a " & CDate(Date + 8) & ").")
                        End If
                        If CDate(.txtDtVenc(lI).Text) > CDate(txtVigTermino.Text) - 10 Then
                            Call gM46V999.gpMenErro("FGE0110 - Data de vencimento da parcela " & lI & " inv�lida (Superior ao t�rmino da vig�ncia).")
                        End If
'Ficha 900046
                        If lI = 1 And frmT47V004A.optPagAntecipado(2).Value = True And CDate(.txtDtVenc(lI).Text) < CDate(Date + 20) Then
                            Call gM46V999.gpMenErro("FGE0009 - Data de vencimento da parcela " & lI & " inv�lida (inferior a " & CDate(Date + 20) & ").")
                        End If
                        If lI = 2 And frmT47V004A.optPagAntecipado(2).Value = False And CDate(.txtDtVenc(lI).Text) < CDate(Date + 20) Then
                            Call gM46V999.gpMenErro("FGE0009 - Data de vencimento da parcela " & lI & " inv�lida (inferior a " & CDate(Date + 20) & ").")
                        End If
                        If lI > 1 Then
                            If CDate(.txtDtVenc(lI).Text) > CDate(.txtDtVenc(lI - 1).Text) + lLimDias Then
                                Call gM46V999.gpMenErro("FGE0109 - Data de vencimento da parcela " & lI & " inv�lida (Superior a " & lLimDias & " da parcela anterior).")
                            End If
'Ficha 901136
                            If DateDiff("d", CDate(.txtDtVenc(lI - 1).Text), CDate(.txtDtVenc(lI).Text)) < 9 Then
                                Call gM46V999.gpMenErro("FGE0109 - Data de vencimento da parcela " & lI & " inv�lida (um vencimento deve ter no m�nimo 9 dias de diferen�a do outro vencimento).")
                            End If
                        End If
                        If lI < objstcPedido.Qtd_Parc Then
                            If gM46V999.gfPreenchido(.txtDtVenc(lI + 1).Text) Then
                                If CDate(.txtDtVenc(lI).Text) > CDate(.txtDtVenc(lI + 1).Text) Then
                                    Call gM46V999.gpMenErro("FGE0009 - Data de vencimento da parcela " & lI & " maior que a data de vencimento da parcela " & lI + 1)
                                End If
                            End If
                        End If
                   End If
                End With
            Next
        End If
    End If
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        If objstcPedido.Cod_Prod = 0 Then
            Call gM46V999.gpMenErro("FGE0015 - Produtor inv�lido (N�o informado).")
        End If
        If objstcPedido.Cod_Unid_Prod = 0 Then
            Call gM46V999.gpMenErro("FGE0016 - Unidade inv�lida (N�o informada).")
        End If
    End If
    
    Dim lQuaDias  As Long
    If objstcPedido.Dat_Ini_Vig <> 0 And objstcPedido.Dat_Ter_Vig <> 0 And (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) And objstcPedido.Ind_CotacaoSolic = True Then
        lQuaDias = gM46V999.gfDifDatas(objstcPedido.Dat_Ini_Vig, objstcPedido.Dat_Ter_Vig)
        If lQuaDias <> 365 And lQuaDias <> 366 And objstcPedido.Ind_CotacaoSolic = True Then
            Call gM46V999.gpMenErro("FNP0011 - Vig�ncia inv�lida. Deve ser anual.")
        End If
        
        If frmT47V016A.optTipEmissao(1).Value = True Then
            If gM46V999.gCodProduto >= 978 Then
                lValidade = gfCalcData(objstcPedido.Dat_Pri_Calculo, 15)
            Else
                lValidade = gfCalcData(objstcPedido.Dat_Pri_Calculo, 30)
            End If
        Else
            lValidade = gfCalcData(objstcPedido.Dat_Pri_Calculo, 15)
        End If
        
        If objstcPedido.Dat_Ini_Vig > lValidade And lValidade <> 0 Then
            'Call  gM46V999.gpMenErro("FNP0011 - Vig�ncia inv�lida. Deve ser menor que o prazo da cota��o.")
        End If
    End If
    
    If objstcPedido.Dat_Emis > 0 Then
        If gM46V999.gfDifDatas(gM46V999.gHoje, objstcPedido.Dat_Ini_Vig) > 45 Then
            Call gM46V999.gpMenErro("FGE0135 - S� � permitida agendar uma emiss�o emiss�o com at� 45 dias de anteced�ncia..")
        End If
        If gM46V999.gfDifDatas(objstcPedido.Dat_Ini_Vig, objstcPedido.Dat_Emis) > 5 Then
            Call gM46V999.gpMenErro("FGE0135 - Data da emiss�o agendada acima de cinco dias do in�cio da vig�ncia.")
        End If
    End If
    
    'Consist�ncia retirada do P01O003.CBL e inclu�da aqui
    If objstcPedido.Cod_Reg = 106 Or objstcPedido.Cod_Reg = 107 Then
        If objstcPedido.objstcCtrlEmis.Cod_Envi_Doc <> "C" Then
            Call gM46V999.gpMenErro("FGE0136 - " & "Para Cod.Registro 106 ou 107 o tipo de envio deve ser CORRETOR.")
        End If
    End If

    Select Case gM46V999.gCod_Ramo
        Case e_CodRamo_Resp_Civil
            If IsDate(txtVigInicio.Text) Then
                If CDate(txtVigInicio.Text) >= CDate("01/01/2013") Then
                    Select Case objstcPedido.Cod_Reg
                        Case 1, 3, 101, 103
                            Call gM46V999.gpMenErro("FGE0136 - " & "A Partir de 01/01/2013, n�o � permitido a cobran�a de Custos. Alterar Cod. Registro")
                    End Select
                End If
            End If
    End Select

'N�o � para zerar
    If gM46V999.gQuaErros <> 0 Then
  '      Call mpEmiMensagens(gM46V999.gPrefixo)
  '      panAguarde.Visible = False
  '      panAguarde.Refresh
  '      frmT47V017A.Show vbModeless, Me
        Exit Sub
    End If
    
    '**BOLETO
    'If frmT46V102A.optPagAntecipado(1).Value = True And frmT46V102A.optForPagamento(0).Value And lBoleto = 0 And (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
    '    lBoleto = 123456
    'End If
    Call gfExportarParaBaseTemp(objstcPedido.NOSSO_NUMERO, objstcPedido.Num_Ped, objstcPedido.objstcCtrlEmis.Num_Bloq, lMensagem)
    '1. Consiste dados comuns - n�vel 1 (dados para c�lculo).

'Sandra - TIRAR COBOL e substituir pela rotina mpConsisteDadosComum
'    Call mpConsisteComum(lCod_Retorno)         'DLL
   
    Call mpConsisteDadosComum(lCod_Retorno)
    
    Exit Sub
    
End Sub

Private Sub mpCarregarRamos()
    'Procedure: obt�m Ramos.
    Dim lMensagem       As String           'Mensagem.
    Dim lrsTbm_Ramos    As clsYasRecordSet  'Registro - Tbm_Ramos.
    Dim lBDConexao      As ADODB.Connection
    Dim lSelect         As String
    
    If Not gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lBDConexao, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If
    '2. Obt�m Todos os Ramos.
    lSelect = "SELECT * from Tbm_Ramo "
    lSelect = lSelect & " WHERE COD_RAMO NOT IN (310,112,113,114,211,212,540,550, 110,120,130,170,330,331,340,350,430,440,480,490,530,560,570,640,660,680,711,730,750,810,812,820,830,930,931,932)"
    lSelect = lSelect & " order by Cod_Ramo "
    
    If gM46V999.gfObtRegistro(lBDConexao, lSelect, lrsTbm_Ramos, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    If lrsTbm_Ramos.EOF = True Then
        Call gM46V999.gpGraLog(1, "FNL0055 - Tabela de Ramos n�o localizado.")
        Exit Sub
    End If
    
    cboRamo.Clear
    '3. Carrega Combo
    Do While Not lrsTbm_Ramos.EOF
        Select Case lrsTbm_Ramos("Cod_Ramo")
            Case 510, 950, 710
                cboRamo.AddItem lrsTbm_Ramos("Dsc_Ramo")
                cboRamo.ItemData(cboRamo.NewIndex) = lrsTbm_Ramos("Cod_Ramo")
        End Select
        lrsTbm_Ramos.MoveNext
    Loop
    'cboRamo.ListIndex = 0
    
    
End Sub

Private Sub cboPlano_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboPlano_LostFocus(Index As Integer)
    If cboPlano(1).ListIndex = -1 Then
        Call gM46V999.gpGraLog(1, "� necess�rio selecionar uma modalidade.")
        cboPlano(1).SetFocus
        If flxGridCoberturaDemais.UBound > 0 Then TabMod.Tabs.Item(flxGridCoberturaDemais.UBound).Selected = True
        Exit Sub
    End If
End Sub

Private Sub cboRisUF_Change()
    Call gpLimPremios
End Sub

Private Sub chkLider_Click()
    Call gpHabSalvar
End Sub

Private Sub CmbCoberturas_Click()
Dim lI  As Integer
Dim lI2 As Integer

    Call gpHabSalvar
    lI = Val(TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag)
    For lI2 = 1 To flxGridCoberturaDemais(lI).Rows - 1
        If flxGridCoberturaDemais(lI).TextMatrix(lI2, 0) <> "" And CmbCoberturas.ListIndex <> -1 And mLinhaFocus <> lI2 Then
            If CmbCoberturas.ItemData(CmbCoberturas.ListIndex) = flxGridCoberturaDemais(lI).TextMatrix(lI2, 0) Then
                Call gM46V999.gpGraLog(1, "Esta cobertura j� foi selecionada para esta modalidade.")
                CmbCoberturas.ListIndex = 0
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub cmdAdicionarCobertura_Click()
Dim lI  As Integer
    'lI = TabMod.SelectedItem.Index
    lI = TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag
    Call flxGridCoberturaDemais_Click(lI)
    If GridFocado = "" Then
        If flxGridCoberturaDemais(lI).Visible = True Then
            GridFocado = "flxGridCoberturaDemais" & lI
        End If
    End If

    Call mpIncluirLinha
End Sub

Private Sub cmdAdicionarTab_Click()
Dim lI  As Long
    
    If TabMod.Tabs.Item(TabMod.Tabs.Count).Caption = "" Then
        TabMod.Tabs.Item(TabMod.Tabs.Count).Selected = True
        txtModalidade.Visible = True
        cboPlano(1).Visible = True
    Else
        TabMod.Tabs.Add
        TabMod.Tabs.Item(TabMod.Tabs.Count).Selected = True
        
        Load flxGridCoberturaDemais(flxGridCoberturaDemais.UBound + 1)
        TabMod.Tabs.Item(TabMod.Tabs.Count).Tag = flxGridCoberturaDemais.UBound
        mpCarregarGridCoberturaMod (flxGridCoberturaDemais.UBound)
        For lI = 0 To flxGridCoberturaDemais.UBound
            On Error Resume Next
            flxGridCoberturaDemais(lI).Visible = False
            On Error GoTo 0
        Next lI
        txtModalidade.Visible = True
        cboPlano(1).Visible = True
        cboPlano(1).ListIndex = -1
        cboPlano(1).SetFocus
    End If

End Sub

Private Sub cmdCepRisco_Click()
Dim lCEP            As String       'CEP.
Dim lTipLogradouro  As String       'Tipo de logradouro.
Dim lLogradouro     As String       'Logradouro.
Dim lBairro         As String       'Bairro.
Dim lCidade         As String       'Cidade.
Dim lUF             As String       'UF.

    Call frmT47V007A.mpInstanciarForm(lCEP, lTipLogradouro, lLogradouro, lBairro, lCidade, lUF)
    
    txtRisCEP.Text = lCEP
    mAntCEP = lCEP
    
    txtRisTipLogradouro.Text = lTipLogradouro
    txtRisNomLogradouro.Text = lLogradouro
    If Mid(cboRamo.Text, 1, 3) = "710" Then
        txtRisNomLogradouro.Text = "CONFORME ESPECIFICACAO EM ANEXO"
    End If
    txtRisBairro.Text = lBairro
    txtRisCidade.Text = lCidade
    If gM46V999.gfPreenchido(lUF) Then
        cboRisUF.Text = lUF
    End If
    Call txtRisCEP_LostFocus
End Sub


Private Sub cmdExcluirCobertura_Click()
Dim lI  As Integer
    lI = TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag
    If GridFocado = "" Then
        
        If flxGridCoberturaDemais(lI).Visible = True Then
            GridFocado = "flxGridCoberturaDemais" & lI
        End If
    End If
    Call mpExcluirLinha
End Sub

Private Sub cmdExcluirModalidade_Click()
Dim lI  As Integer

    'lI = TabMod.SelectedItem.Index
    lI = TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag
    If TabMod.Tabs.Count = 1 Then
        TabMod.Tabs.Clear
        TabMod.Tabs.Add
        TabMod.Tabs.Item(TabMod.Tabs.Count).Tag = 0
        flxGridCoberturaDemais(0).Rows = 1
        flxGridCoberturaDemais(0).Rows = 2
        flxGridCoberturaDemais(0).FixedRows = 1
        cboPlano(1).ListIndex = -1
        txtModalidade.Visible = True
        flxGridCoberturaDemais(0).Visible = False
        cboPlano(1).Visible = True
    Else
        
        If cboPlano(1).ListIndex = -1 Then
            If cboPlano(1).ListCount > 0 Then
                cboPlano(1).ListIndex = 0
            End If
        End If
        TabMod.Tabs.Remove TabMod.SelectedItem.Index
        If lI = 0 Then
            TabMod.Tabs(1).Selected = True
            Exit Sub
        End If
        Unload flxGridCoberturaDemais(lI)
        If TabMod.Tabs.Count = 1 Then
            TabMod.Tabs(1).Selected = True
        Else
            TabMod.Tabs(TabMod.Tabs.Count - 1).Selected = True
        End If
    End If
End Sub


Private Sub cmdPrazo_Click()
    Call mnuExiPrazo_Click
End Sub

Private Sub cmdProduto_Click()
    Call mnuEdiBusProduto_Click
End Sub


Private Sub Form_Activate()
Dim lMudou  As Boolean

    Select Case UCase(gProChamado)
        Case "T47V014A"
                gProChamado = ""
                With frmT47V014A.flxItens
                    '===1) Verifica se mudou posi��o na grade de itens.
                    lMudou = False
                    If .Row <> 0 Then
                        If gM46V999.gfPreenchido(txtNumItem.Text) Then
                            .Col = 0    'N�mero do item.
                            If Val(txtNumItem.Text) <> Val(.Text) Then
                                lMudou = True
                            End If
                        Else
                            lMudou = True
                        End If
                        If lMudou = True Then
                            '===a) Obt�m novo item.
                            .Col = 0    'N�mero do item.
                            On Error Resume Next
                            Set objstcPedLoc = objstcPedido.ColecaoPedLoc.Item(objstcPedLoc.NOSSO_NUMERO & Format(.Text, "00000"))
                            On Error GoTo 0
                            '===b) Trata resultado.
                            If objstcPedLoc.Num_Item = Val(txtNumItem.Text) Then
                                Call gM46V999.gpGraLog(2, "FTE0004 - " & gM46V999.gPrefixo & "Item n�o localizado (" & .Text & _
                                                 " - navega��o entre itens).")
                            Else
                                gCarga = True
                                Call mpPosItem(objstcPedLoc)
                                gCarga = False
                            End If
                        End If
                    End If
                End With
    End Select
    
    frmT47V001A.Refresh
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lMensagem       As String   'Mensagem.


    'Verifica se form est� desabilitado
    If gM46V999.gDesabilitarForms Then
        Exit Sub
    End If

    'Confirma com usu�rio e verifica se [gSinalRegPadExcluido] para n�o exibir msg para salvar quando excluindo RP.
    If mnuArqSalvar.Enabled = True Then
        lMensagem = "AGE2015 - " & gM46V999.gPrefixo & "Deseja salvar as altera��es?"
        Select Case MsgBox(lMensagem, vbYesNo + vbDefaultButton2 + vbQuestion)
            Case vbYes
                If mfSalvar() = False Then
                    lMensagem = "AGE0024 - Houve erro ao salvar informa��es da proposta. Confirma sa�da " & _
                                "do sistema?"
                    If MsgBox(lMensagem, vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                        Exit Sub
                    End If
                    barFerramentas.Buttons(21).Image = 13   'Sair.
                    Cancel = True
                    Exit Sub
                End If
        End Select
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Observa��o: controle criado para usu�rio sair do FORM utilizando tecla ESC. Invis�vel pelo usu�rio.
    Call mnuArqSaiItem_Click
End Sub

Private Sub mnuArqExcProposta_Click()
    Dim lMensagem   As String           'Mensagem.

    Screen.MousePointer = vbHourglass

    '1. Confirma com usu�rio.
    lMensagem = "AGE0023 - " & gM46V999.gPrefixo & "Confirma exclus�o?" & vbLf & vbLf & _
                "Proposta " & gM46V111.gNumPedido & " com nosso n�mero " & gM46V111.gComNosNumero & "."
    If MsgBox(lMensagem, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
        Call gM46V999.gpGraLog(0, lMensagem & " N�o.")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Call gM46V999.gpGraLog(0, lMensagem & " Sim.")

    '2. Exclui dados comuns.
    gM46V999.gBaseLocal = False
    Call gM46V999.gpExcluiRegistros(objstcPedido.NOSSO_NUMERO, 0, "", "", lMensagem)
    
    '8. Emite mensagem.
    Call gM46V999.gpGraLog(1, "AGE0030 - " & gM46V999.gPrefixo & "Exclus�o realizada com sucesso." & vbLf & vbLf & _
                     "Proposta " & gM46V111.gNumPedido & " com nosso n�mero " & objstcPedido.NOSSO_NUMERO & ".")
NaoExcluir:
    Call mpDesSalvar
    Screen.MousePointer = vbDefault
    
    Unload Me
End Sub

Private Sub mnuArqNovaConsultaPI_Click()
    txtNumPI.Text = ""
    panComum.Visible = False
    panValores.Visible = False
    panLocRisco.Visible = False
    panItem.Visible = False
    panItemLocRisco.Visible = False
    Set objstcPedido = Nothing
    Set objstcPedLoc = Nothing
    panPesquisa.Enabled = True
    cboRamo.Enabled = True
    Call mpControleBotoes(False)
    gM46V111.gNum_Pi = 0

End Sub

Private Sub mpControleBotoes(pHabilitar As Boolean)
    
    barFerramentas.Buttons("Novo").Enabled = pHabilitar
    mnuArqNovItem.Enabled = pHabilitar
    
    barFerramentas.Buttons("Abrir").Enabled = pHabilitar
    mnuArqAbrItem.Enabled = pHabilitar
    
    barFerramentas.Buttons("Salvar").Enabled = pHabilitar
    mnuArqSalvar.Enabled = pHabilitar
    
    barFerramentas.Buttons("Excluir").Enabled = pHabilitar
    mnuArqExcProposta.Enabled = pHabilitar
    mnuEdiExcItem.Enabled = pHabilitar
    
    
    barFerramentas.Buttons("Copiar").Enabled = pHabilitar
    mnuEdiCopItem.Enabled = pHabilitar
    
    barFerramentas.Buttons("Desfazer").Enabled = pHabilitar
    mnuEdiDesfazer.Enabled = pHabilitar
    
    barFerramentas.Buttons("Bens").Enabled = pHabilitar
    mnuExiBem.Enabled = pHabilitar
    
    barFerramentas.Buttons("Proposta").Enabled = pHabilitar
    mnuExiDadProposta.Enabled = pHabilitar
    
    mnuExiPrazo.Enabled = pHabilitar
    mnuExiTipoEmissao.Enabled = pHabilitar
    
    mnuExiEspecAnt.Enabled = pHabilitar
    
    mnuTextoEndosso.Enabled = pHabilitar
    
End Sub
Private Sub mnuArqSalvar_Click()
    If gM46V111.gNum_Pi > 0 Then
        Screen.MousePointer = vbHourglass
        If mnuArqSalvar.Enabled = True Then
            Call mfSalvar
        End If
        Call mpValidar
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub mpVerIteNiv_2(pobjstcItem As stcA46V708B, Optional ByRef pValorCoberturas As Double, Optional ByRef pISCoberturas As Double)
    'Procedure: consiste item n�vel 2: consiste todos dados do item.
        'Procedure: consiste item n�vel 2: consiste todos dados do item.
Dim lobjstcCobert   As stcA46V704B
Dim lDig            As String
Dim clsA99V001      As clsA99V001
Dim lMensagem       As String
Dim lAuxiliar       As String
Dim lNumCobert      As Integer
Dim lSomaCobert     As Double
Dim lArrayCoberts() As Integer
Dim lArrayISCoberts() As Double
Dim i As Integer
Dim lDescontoComercial As Double
Dim lComissao As Double
Dim lobjstcPedClau   As stcA46V703B     'Objeto da classe de estrutura stcA46V704B (TAB_PED_CLAU).
Dim lDesconto       As Double
Dim pulaLimite      As Boolean
On Error GoTo ControleErro

    '3. Endere�o do segurado.
    If panLocRisco.Visible = True Then
        If Not gM46V999.gfPreenchido(pobjstcItem.Cod_Tip_Logr_Risco) Then
            Call gM46V999.gpMenErro("FGE0095 - Endere�o de risco inv�lido (tipo do logradouro n�o preenchido).")
        Else
            If Not gM46V999.gfPreenchido(pobjstcItem.Nom_Logr_Risco) Then
                Call gM46V999.gpMenErro("FGE0096 - Endere�o de risco inv�lido (nome do logradouro n�o preenchido).")
            Else
                If pobjstcItem.objstcPedInsp.Num_Logr_Risco = 0 And Not gM46V999.gfPreenchido(pobjstcItem.Dsc_Compl_Risco) Then
                    Call gM46V999.gpMenErro("FGE0097 - Endere�o de risco inv�lido (N�mero n�o preenchido).")
                Else
                    If Not gM46V999.gfPreenchido(pobjstcItem.Nom_Bairro_Risco) Then
                        Call gM46V999.gpMenErro("FGE0098 - Endere�o de risco inv�lido (bairro n�o preenchido).")
                    Else
                        If Not gM46V999.gfPreenchido(pobjstcItem.Nom_Cid_Risco) Then
                            Call gM46V999.gpMenErro("FGE0099 - Endere�o de risco inv�lido (cidade n�o " & _
                                           "preenchida).")
                        Else
                            If Not gM46V999.gfPreenchido(pobjstcItem.Sig_UF_Risco) Then
                                Call gM46V999.gpMenErro("FGE0100 - Endere�o de risco inv�lido (unidade da " & _
                                               "federa��o n�o preenchida).")
                            Else
                                If Not gM46V999.gfPreenchido(pobjstcItem.Num_Cep_Risco) Then
                                    Call gM46V999.gpMenErro("FGE0101 - Endere�o de risco inv�lido (CEP n�o " & _
                                                   "preenchido).")
                                Else
                                    If Val(pobjstcItem.Num_Cep_Risco) = 0 Then
                                        Call gM46V999.gpMenErro("FGE0102 - Endere�o de risco inv�lido (CEP " & _
                                                       "zerado).")
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
        
    If gM46V999.PerfilAcesso = e_Perfil_DIG_Gestor Then
        lNumCobert = 0
        If pobjstcItem.ColecaoPedCobert.Count = 0 Then
            Call gM46V999.gpMenErro("FGE0107 - Nenhuma cobertura foi selecionada")
        End If
        For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert
            If lobjstcCobert.Val_IS = 0 And gM46V111.gCod_Dizer = 0 And objstcPedido.Cod_Reg <> 20 Then  'N�o consiste para endosso
                Call gM46V999.gpMenErro("FGE0107 - Valor L.M.G da cobertura " & lobjstcCobert.Cod_Cobert & " inv�lido (n�o informado).")
            End If
            If lobjstcCobert.Val_IS < 100# And gM46V111.gCod_Dizer = 0 And objstcPedido.Cod_Reg <> 20 Then
                Call gM46V999.gpMenErro("FGE0112 - Valor L.M.G da cobertura " & lobjstcCobert.Cod_Cobert & " inv�lido (menor que 100,00).")
            End If

            lNumCobert = lNumCobert + 1
            ReDim Preserve lArrayCoberts(lNumCobert)
            ReDim Preserve lArrayISCoberts(lNumCobert)
            lArrayCoberts(lNumCobert) = lobjstcCobert.Cod_Cobert
            lArrayISCoberts(lNumCobert) = lobjstcCobert.Val_IS
            pValorCoberturas = pValorCoberturas + lobjstcCobert.VAL_PR
            pISCoberturas = pISCoberturas + lobjstcCobert.Val_IS
        Next
        If lNumCobert > 22 Then
            Call gM46V999.gpMenErro("AGE0052 - Quantidade de coberturas selecionadas acima do limite (22). ")
        End If
    End If
    Exit Sub
ControleErro:

MsgBox "mpVerIteNiv_2 " & Error

End Sub

Private Sub mpEmiMensagens(ByVal pPrefixo As String)
    'Procedure: emite mensagens (transfere da matriz para a grade de mensagens).

    'Par�metro de entrada...pPrefixo...Prefixo da mensagem.

    Dim lI          As Integer  'Utilizado com For...Next.
    Dim lMensagem1  As String   'Mensagem.
    Dim lMensagem2  As String   'Mensagem.
    Dim lPosicao    As Integer  'Utilizado com InStr.

    '1. Se o form T28V136A ainda n�o estiver dispon�vel, abrir e deixar a grade limpa.
    If gM46V999.gfDisponivel("frmT47V017A") = False Then
        Load frmT47V017A
        frmT47V017A.flxMensagens.Rows = 0
    End If

    '2. Posiciona mensagens.
    With frmT47V017A
        For lI = 1 To gM46V999.gQuaErros
            lMensagem1 = pPrefixo & gM46V999.gerros(lI)
            .lblAuxMensagem.Width = .flxMensagens.ColWidth(0) - 300
            .lblAuxMensagem.Caption = Trim$(lMensagem1)
            .flxMensagens.Rows = .flxMensagens.Rows + 1
            .flxMensagens.Row = .flxMensagens.Rows - 1
            .flxMensagens.RowHeight(.flxMensagens.Row) = .lblAuxMensagem.Height
            .flxMensagens.Text = .lblAuxMensagem.Caption
            
            lPosicao = InStr(gM46V999.gerros(lI), " - ")
            lMensagem2 = Left$(gM46V999.gerros(lI), lPosicao + 2) & gM46V999.gPrefixo & pPrefixo & _
                         Mid$(gM46V999.gerros(lI), lPosicao + 3)
            Call gM46V999.gpGraLog(0, lMensagem2)
            .Caption = .Caption
        Next lI
        gM46V999.gpInicializar_gErros
    End With
End Sub

Private Function mpValidar() As Boolean
Dim lobjstcItem As stcA46V708B
Dim lobjstcCobert   As stcA46V704B
Dim lValorCoberturas As Double
Dim lModalidade As Double
Dim lISCoberturas As Double
Dim lMensagem As String

    gM46V999.gpInicializar_gErros
    mpValidar = False

    Unload frmT47V017A

    If gM46V999.PerfilAcesso = e_Perfil_DIG_Gestor Then
        'Verificar valor dos pr�mios das coberturas
        lValorCoberturas = 0
        For Each lobjstcItem In objstcPedido.ColecaoPedLoc
            Call mpVerIteNiv_2(lobjstcItem, lValorCoberturas, lISCoberturas)
        Next lobjstcItem
            
        If Val(lValorCoberturas) > 0 Then
            If Val(lValorCoberturas) <> Val(objstcPedido.VAL_PR) Then
                Call gM46V999.gpMenErro("AGE0052 - Valor do pr�mio total informado (" & Format(objstcPedido.VAL_PR, "standard") & " difere do pr�mio das coberturas " & Format(lValorCoberturas, "standard"))
            End If
        Else
            If objstcPedido.Cod_Reg <> 2 And objstcPedido.Cod_Reg <> 20 And objstcPedido.Cod_Reg <> 120 And objstcPedido.Cod_Reg <> 202 And _
               objstcPedido.Cod_Reg <> 3 And objstcPedido.Cod_Reg <> 102 And objstcPedido.Cod_Reg <> 103 And objstcPedido.Cod_Reg <> 203 Then
                If lVerTipEmissao = "S" Or lVerTipEmissao = "R" Then
                    Call gM46V999.gpMenErro("AGE0053 - Falta Modalidade / Valor do pr�mio n�o informado.")
                End If
            Else
               gM46V999.gDesabilitarForms = False
            End If
        End If
        
        If Val(lISCoberturas) <> Val(objstcPedido.Val_IS) Then
            lMensagem = "AGE0054 - LMG das coberturas / modalidade est� divergente do LMG B�sico informado!" & vbLf & vbLf & _
                "Confirma somat�ria de pr�mios divergentes?"
            If MsgBox(lMensagem, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call gM46V999.gpGraLog(0, lMensagem & " SIM.")
            Else
                Call gM46V999.gpGraLog(0, lMensagem & " N�O.")
                Exit Function
            End If
       End If
    End If
    
    If lVerTipEmissao = "E" And objstcPedido.ColecaoTextoEndo.Count = 0 Then
        Call gM46V999.gpMenErro("AGE0055 - Texto do Endosso n�o encontrado!")
    End If

    If gM46V999.gQuaErros <> 0 Then
        Call mpEmiMensagens(gM46V999.gPrefixo)
        panAguarde.Visible = False
        panAguarde.Refresh
        frmT47V017A.Show vbModeless, Me
        Exit Function
    End If

    Call mfCalPremios
  
    'Verificar qtde de coberturas contratadas por modalidade
'    For Each lobjstcCobert In objstcPedLoc.ColecaoPedCobert
'        lModalidade = lobjstcCobert.Tab_Cobert
'        Call mpVerQtdeCoberturas(lobjstcItem, lModalidade)
'    Next lobjstcCobert
'

    If gM46V999.gQuaErros <> 0 Then
        Call mpEmiMensagens(gM46V999.gPrefixo)
        panAguarde.Visible = False
        panAguarde.Refresh
        frmT47V017A.Show vbModeless, Me
        Exit Function
    End If
    
    panAguarde.Caption = "Aguarde... Efetuando totaliza��o..."
    panAguarde.Refresh
    DoEvents

    mnuSituacao.Caption = Chr(183) & Space(10) & "PI: " & frmT47V001A.txtNumPI.Text & " -  Consistido"
    
    If gM46V111.gCod_plano > 0 Then mpValidar = True
    
    Screen.MousePointer = vbDefault
End Function
Function mfSalvar() As Boolean
'********************Vari�veis padr�o*************************************
Dim lCkPoint        As String      'Variavel utilizada para determinar pontos onde o c�digo j� passou.
Dim lCkStr          As String       'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
                                    'lCkPoint. (posi��o das variaveis em ordem de declara��o)
'********************Vari�veis Declaradas para o processamento************
Dim lMensagem       As String       'Mensagem.
Dim lAtivo          As String
Dim lBlnSiscad          As Boolean      'Indicador se existe dados no Siscad
Dim lI              As Integer



    
    Err = 0
    lCkPoint = ""
    gM46V999.gBaseLocal = False
    lCkStr = ""
    If GridFocado <> "" Then
        Call flxGridCoberturaDemais_Click(Val(Replace(UCase(GridFocado), "FLXGRIDCOBERTURADEMAIS", "")))
    End If
    mfSalvar = False
    On Error GoTo ControleErro
    '1. Verifica situa��o do documento.
    lCkPoint = "1"
    If objstcPedido Is Nothing Then
        Set objstcPedido = New stcA46V702B
        Set objstcPedido.M46V999 = gM46V999
        
        Set objstcPedLoc = New stcA46V708B
        Set objstcPedLoc.ObjPai = objstcPedido
        
    End If
    If objstcPedido.Cod_Sit_Ped = e_SitPedido_Transmitido Then
        lMensagem = "AGE0020 - Proposta " & gM46V111.gNumPedido & " (nosso n�mero " & gM46V111.gComNosNumero & _
                    ") j� transmitida." & vbLf & vbLf & _
                    "Foi acionado o processo de salvar os dados da proposta e este processo modifica " & _
                    "a indica��o de que a proposta j� foi transmitida." & vbLf & vbLf & _
                    "Confirma a execu��o deste processo?"
        If MsgBox(lMensagem, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Call gM46V999.gpGraLog(0, lMensagem & " SIM.")
        Else
            Call gM46V999.gpGraLog(0, lMensagem & " N�O.")
            mfSalvar = True
            Exit Function
        End If
    End If
    lCkPoint = "2"
    
    '1. Se foco est� em algum campo que possue evento LostFocus, o mesmo deve ser acionado.
    'Obs.: Estes comandos foram necess�rios pois em um caso muito especial n�o existe nenhum campo com
    '      com foco e ocorreu erro.
    On Error Resume Next
    lAtivo = UCase$(ActiveControl.Name)
    On Error GoTo 0
    If gM46V999.gCod_Ramo = e_CodRamo_Roubos Or gM46V999.gCod_Ramo = e_CodRamo_Riscos_diversos Then
        If Val(Replace(frmT47V004A.txtCEP.Text, "-", "")) = 0 And Val(Replace(Me.txtRisCEP.Text, "-", "")) <> 0 Then
            With frmT47V004A
                .txtTipLogradouro = Me.txtRisTipLogradouro.Text
                .txtNomLogradouro = Me.txtRisNomLogradouro.Text
                .txtComplemento = Me.txtRisNumero & " - " & Me.txtRisComplemento.Text
                .txtBairro = Me.txtRisBairro.Text
                .txtCidade = Me.txtRisCidade.Text
                .cboUF.Text = Me.cboRisUF.Text
                .txtCEP = Me.txtRisCEP.Text
            End With
        End If
    End If
    
    If gM46V999.gfPreenchido(lAtivo) Then
        Select Case lAtivo
            Case "TXTNUMITEM"
                Call txtNumItem_LostFocus
            Case "TXTRISCEP"
                If gM46V999.gfPreenchido(txtRisCEP.Text) Then
                    Call txtRisCEP_LostFocus
                    If Not gM46V999.gfPreenchido(txtRisCEP.Text) Then
                        Exit Function
                    End If
                End If
            Case "TXTCOMISSAO"
                If gM46V999.gfPreenchido(txtComissao.Text) Then
                    Call TxtComissao_LostFocus
                    If Not gM46V999.gfPreenchido(txtComissao.Text) Then
                        Exit Function
                    End If
                End If
            Case "TXTVIGINICIO"
                If gM46V999.gfPreenchido(txtVigInicio.Text) Then
                    Call txtVigInicio_LostFocus
                    If Not gM46V999.gfPreenchido(txtVigInicio.Text) Then
                        Exit Function
                    End If
                End If
            Case "TXTVIGTERMINO"
                If gM46V999.gfPreenchido(txtVigTermino.Text) Then
                    Call txtVigTermino_LostFocus
                    If Not gM46V999.gfPreenchido(txtVigTermino.Text) Then
                        Exit Function
                    End If
                End If
            Case "TXTISBAS"
                If gM46V999.gfPreenchido(txtISBas.Text) Then
                    Call txtISBas_LostFocus
                    If Not gM46V999.gfPreenchido(txtISBas.Text) Then
                        Exit Function
                    End If
                End If
            Case "TXTVALORISCO"
                If gM46V999.gfPreenchido(txtValoRisco.Text) Then
                    Call txtValoRisco_LostFocus
                    If Not gM46V999.gfPreenchido(txtValoRisco.Text) Then
                        Exit Function
                    End If
                End If
        End Select
    End If
    
    'Atualizar com os dados do Siscad
    lCkPoint = "2AA"
    
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And (gM46V999.PerfilAcesso <> e_Perfil_CALC And gM46V999.PerfilAcesso <> e_Perfil_CALC_Gestor) Then
       Call frmT47V004A.mpAtualizaDadosSegurado(lBlnSiscad, False)
       If lBlnSiscad = False Then
           lMensagem = "FGE0020 - Dados do Siscad n�o localizados."
           MsgBox lMensagem, vbInformation
            If Not UCase(gM46V999.gSerDIVERSOS) Like "*MIX*" Then
                   Exit Function
            End If
            lMensagem = ""
       End If
    End If
    
    
    '2. Salvo os dados de comum
    lCkPoint = "2A"
    If Not mfSalvarComum(lMensagem) Then GoTo ControleErro
        
    lCkPoint = "2B"
    '3. Salvo os dados do item
    If Not mfSalvarItem(lMensagem, objstcPedLoc) Then GoTo ControleErro
    
    lCkPoint = "18"
    'Salvar Texto da Especifica��o (Beneficiarios)
    'tem que ser depois que salva o item para utilizar a colecao de coberturas atualizada.
    If Not mfSalvarTextoEspecPadrao(lMensagem) Then GoTo ControleErro
    
    lCkPoint = "4"
    '4. Atualizo a grade de itens
    Call mpCriGraItens
    lCkPoint = "5"
    '5. Desabilito o salvar.
    Call mpDesSalvar
    lCkPoint = "6"
    
    mfSalvar = True

    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(1, lMensagem, vbCritical)
        'Obtem o inicio da mensagem de erro
    lMensagem = "mfSalvar ERRO - " & Err & " " & Error & _
                Chr(13) & "CheckPoint " & lCkPoint & " - " & lMensagem
    Call gM46V999.gpGraLog(0, lMensagem)
    
    mfSalvar = False

End Function
Private Sub mpDesSalvar()
    'Procedure: desabilita op��es de salvar e desfazer.

    mnuArqSalvar.Enabled = False
    mnuEdiDesfazer.Enabled = False
    barFerramentas.Buttons(4).Enabled = False   'Salvar.
    barFerramentas.Buttons(12).Enabled = False  'Desfazer.
    If gM46V999.gCod_Ramo <> e_CodRamo_NaoDefinido Then
        mnuArqImprimir.Enabled = True
        mnuArqVisualizar.Enabled = True
    End If
    barFerramentas.Buttons(6).Enabled = True   'Imprimir.
    barFerramentas.Buttons(7).Enabled = True   'Visualizar.
End Sub

Private Sub mpCriGraItens()
    'Procedure: cria grade de itens.

    Dim lMensagem   As String           'Mensagem.
    Dim lSaida      As Variant          'Dados de sa�da da fun��o OBTEM_VEICULO.
    Dim pobjstc     As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lCodAnt     As String           'C�digo do ve�culo anterior
    Dim lDescAnt    As String           'Descri��o do ve�culo anterior
        
    '3. Posiciona dados na grade.
    With frmT47V014A.flxItens
        'a) Inicializa grade.
        .Rows = 1
        'b) Sem registro de item.
        If objstcPedido.ColecaoPedLoc.Count = 0 Then
            .Row = 0
            Exit Sub
        End If
        'c) Monta grade.
        For Each pobjstc In objstcPedido.ColecaoPedLoc
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .TextMatrix(.Rows - 1, 0) = pobjstc.Num_Item   'N�mero do item.
            .TextMatrix(.Rows - 1, 1) = Format(pobjstc.Num_Cep_Risco, "00000-000")
        Next pobjstc
        .ColSel = 0
        .Sort = 1
        'd) Seleciona primeira linha.
        .Row = 1
        .Col = 0    'N�mero do item.
    End With
End Sub


Private Function mfSalvarItem(pMensagem As String, pobjstcPedLoc As stcA46V708B) As Boolean
'********************Vari�veis padr�o*************************************
Dim lCkPoint        As String      'Variavel utilizada para determinar pontos onde o c�digo j� passou.
Dim lCkStr          As String       'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
                                    'lCkPoint. (posi��o das variaveis em ordem de declara��o)
'********************Vari�veis Declaradas para o processamento************
Dim objPedidoLoc      As clsA46V708A      'Objeto da classe clsA46V708A (TAB_PED_LOC).
Dim lChave          As Long             'Chave
Dim lI              As Integer          'Retorno do INSTR.

Dim lobjstcItem     As stcA46V708B
Dim lobjstcCobert   As stcA46V704B
Dim lColecaoItemTemp    As Collection
Dim lBonus          As Integer

    Err = 0
    lCkPoint = ""
    lCkStr = ""
    On Error GoTo ControleErro
    Set objPedidoLoc = New clsA46V708A
    Set objPedidoLoc.M46V999 = gM46V999
    

    lCkPoint = "1"
    With pobjstcPedLoc
        '== N�mero do item anterior.
        .Cod_Ramo = objstcPedido.Cod_Ramo
        .NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
        .Num_Ped = objstcPedido.Num_Ped
        If lblAntNumItem.Caption <> txtNumItem.Text Then
            .Num_Item_Ant = lblAntNumItem.Caption
        Else
            .Num_Item_Ant = txtNumItem.Text
        End If
        '== N�mero do item.
        .Num_Item = txtNumItem.Text
        '== C�digo do ramo.
        .Cod_Ramo = objstcPedido.Cod_Ramo
        lCkPoint = "2"
        '== Risco - Tipo de logradouro.
        .Cod_Tip_Logr_Risco = txtRisTipLogradouro.Text
        '== Risco - Logradouro.
        .Nom_Logr_Risco = txtRisNomLogradouro.Text
        '== Risco - Complemento.
        .Dsc_Compl_Risco = txtRisComplemento.Text
        '== Risco - Bairro.
        .Nom_Bairro_Risco = txtRisBairro.Text
        '== Risco - Cidade.
        .Nom_Cid_Risco = txtRisCidade.Text
        '== Risco - UF.
        If cboRisUF.ListIndex <> -1 Then
            .Sig_UF_Risco = cboRisUF.List(cboRisUF.ListIndex)
        Else
            .Sig_UF_Risco = ""
        End If
        '== Risco - CEP.
        .Num_Cep_Risco = Val(Replace(txtRisCEP.Text, "-", ""))
        
        .Num_Insp = 0
        
        If cboPlano(1).ListIndex = -1 Then
            .Cod_Plano = 0
        Else
            .Cod_Plano = cboPlano(1).ItemData(cboPlano(1).ListIndex)
        End If

        gM46V111.gCod_plano = .Cod_Plano
        lCkPoint = "3"
        '== Desconto progressivo.
        .Per_Desc_Progres = 0
        
        
        '== IS B�sica
        If gM46V999.gfPreenchido(txtISBas.Text) = False Then txtISBas.Text = "0,00"
        If CDbl(Val(txtValoRisco.Text)) <> 0 Then
            .Per_IS_Vr = Val(txtISBas.Text) * 100 / Val(txtValoRisco.Text)
        Else
            .Per_IS_Vr = 0
        End If
        
        '= Coeficiente IS/VR
        If gM46V999.gfPreenchido(txtCoefISVR.Text) = False Then txtCoefISVR.Text = "0,00"
        .Val_Coef_IS_Vr = txtCoefISVR.Text
        
        lCkPoint = "4"
        '== Valor de risco
        .Val_Vr = txtValoRisco.Text
        
        '== Classifica��o
        If txtClassificacao.Tag <> "" Then
            '.Tab_Clasf = Right(txtClassificacao.Tag, 4)
            .Tab_Clasf = 0
            .Cod_Clasf = Left(txtClassificacao.Tag, 4)
            .Cod_Compl_Clasf = Mid(txtClassificacao.Tag, 5, 2)
            lI = InStr(1, txtClassificacao.Text, "-")
            lI = InStr(lI + 1, txtClassificacao.Text, "-")
            .Dsc_Clasf = Mid(txtClassificacao.Text, lI + 1)
        Else
            .Tab_Clasf = 0
            .Cod_Clasf = ""
            .Cod_Compl_Clasf = ""
            .Dsc_Clasf = txtClassificacao.Text
        End If
        lCkPoint = "5"
        '.Cod_Classe_Ocup
        '== Constru��o
        .Cod_Constr = 0
        
        
        '==Usu�rio de libera��o
        lCkPoint = "6"
        .Cod_User_Liber = gM46V999.gNomUsuario6
        .Dat_Liber = Left(gM46V999.gfForData(3), 8)
        '[FRH]->
        objstcPedido.Dat_Liber = .Dat_Liber
        '<-[FRH]
        .Hor_Liber = Mid(gM46V999.gfForData(3), 9, 4)
        .Cod_User_Digit = .Cod_User_Liber
        .DAT_DIGIT = .Dat_Liber
        .HOR_DIGIT = .Hor_Liber
        
        '== Situa��o do item
        .Cod_Sit_item = e_SitItem_NaoCalc
        
        '== Percentual de desconto de bonus.
        .Per_Bonus = 0
        
        '==Quantidade de empregados.
        .Qtd_Empreg = 0
                
        lCkPoint = "7"
        'Identificadores de seguran�a
        .Cod_Iden_Segur_01_V = "N"
        .Cod_Iden_Segur_02_A = "N"
        .Cod_Iden_Segur_03_P = "N"
        .Cod_Iden_Segur_04_I = "N"
        .Cod_Iden_Segur_05_G = "N"
        lCkPoint = "8"
        .Cod_Iden_Segur_06 = "N"
        .Cod_Iden_Segur_07 = "N"
        .Cod_Iden_Segur_08 = "N"
        .Cod_Iden_Segur_09 = "N"
        .Cod_Iden_Segur_10 = "N"
        '== Percentual total de seguran�a
        .Per_Tot_Segur = .gPer_Tot_Segur
        
        '== Identificador de prote��o.
        .Cod_Iden_Protec_01_E = "N"
        .Cod_Iden_Protec_02_H = "N"
        .Cod_Iden_Protec_03_S = "N"
        .Cod_Iden_Protec_04_D = "N"
        lCkPoint = "9"
        .Cod_Iden_Protec_05 = "N"
        .Cod_Iden_Protec_06 = "N"
        .Cod_Iden_Protec_07 = "N"
        .Cod_Iden_Protec_08 = "N"
        .Cod_Iden_Protec_09 = "N"
        .Cod_Iden_Protec_10 = "N"
        '== Percentual total de prote��o.
        
        .Per_Tot_Protec = .gPer_Tot_Protec
        '== Desconto de aparelho.
        .Per_Desc_Apar = .gPer_Tot_Protec
        
        '== produto da tabela de regi�o.
        .Cod_Produto = gM46V111.gCod_plano
        .Cod_Iden_Tarif = gM46V999.gCodProduto
        .NUM_PI = txtNumPI.Text
        lCkPoint = "10"
        .Num_Apol_Anterior = Val(txtApolice.Text)
        .Num_Item_Anterior = 0
        .Ind_Sinistro = e_SimNao_Nao
        .Tip_Emissao = mTipoEmissao
        
    End With
    lCkPoint = "11"
    gM46V999.gBaseLocal = False
    Set pobjstcPedLoc.ObjPai = objstcPedido
    Set pobjstcPedLoc.M46V999 = gM46V999
    
    If mInclusao = True Then
        If Not objPedidoLoc.mfIncluir(pobjstcPedLoc, "", "", "") Then
            pMensagem = objPedidoLoc.UltimoErro
            Exit Function
        End If
    Else

        If Not objPedidoLoc.mfAlterar(pobjstcPedLoc, "", "") Then
            pMensagem = objPedidoLoc.UltimoErro
            Exit Function
        End If
    End If
    lCkPoint = "12"
    'Salvo as coberturas
    If Not mfSalvarCoberturas(pMensagem, objstcPedLoc) Then GoTo Finalizar
    lCkPoint = "15"
    'Salvo os bens
    If Not mfSalvarBens(pMensagem, objstcPedLoc) Then GoTo Finalizar
    lCkPoint = "19"
    
    If Not mfSalvarInspecao(pMensagem, objstcPedLoc) Then GoTo Finalizar
    
    'removo e depois adiciono na cole��o de itens novamente.
    
    If objstcPedido.ColecaoPedLoc.Count > 0 Then
        objstcPedido.ColecaoPedLoc.Remove pobjstcPedLoc.NOSSO_NUMERO & Format(pobjstcPedLoc.Num_Item_Ant, "00000")
    End If
    'Verifico em que lugar da cole��o ser� adicionado.
    '� feito isso para que n�o seja exibido com os n�mero fora de ordem
    lCkPoint = "20"
    lChave = gpPosicaoColecao(pobjstcPedLoc.Num_Item)
    
    Set pobjstcPedLoc.ObjPai = objstcPedido
    If objstcPedido.ColecaoPedLoc.Count > 0 And objstcPedido.ColecaoPedLoc.Count <> lChave - 1 Then
        objstcPedido.ColecaoPedLoc.Add pobjstcPedLoc, pobjstcPedLoc.Class_ID, lChave
    Else
        objstcPedido.ColecaoPedLoc.Add pobjstcPedLoc, pobjstcPedLoc.Class_ID
    End If
    lCkPoint = "21"
    
    lblAntNumItem.Caption = txtNumItem.Text
    mfSalvarItem = True
Finalizar:
    Set objPedidoLoc = Nothing
    Exit Function
ControleErro:
    'Obtem o inicio da mensagem de erro
    Call gM46V999.gpGraLog(0, "mfSalvarItem ERRO - " & Err & " " & Error & _
                 "CheckPoint " & lCkPoint & " - " & pMensagem)
                
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    mfSalvarItem = False
End Function
Public Function gpPosicaoColecao(pItem As Long) As Long
Dim lCount As Long

    For lCount = 1 To frmT47V014A.flxItens.Rows - 1
        If pItem = frmT47V014A.flxItens.TextMatrix(lCount, 0) Then
            gpPosicaoColecao = lCount
            Exit Function
        End If
    Next
    'Se houve altera��o de n�mero e o n�mero for menor que algum dos itens...
    For lCount = 1 To frmT47V014A.flxItens.Rows - 1
        If pItem < frmT47V014A.flxItens.TextMatrix(lCount, 0) Then
            gpPosicaoColecao = lCount
            Exit Function
        End If
    Next
    'Se houve altera��o de n�mero e o n�mero for maior que os demais...
    For lCount = 1 To frmT47V014A.flxItens.Rows - 1
        If pItem > frmT47V014A.flxItens.TextMatrix(lCount, 0) Then
            gpPosicaoColecao = lCount
        End If
    Next
End Function

Private Function mfSalvarInspecao(pMensagem As String, pobjstcPedLoc As stcA46V708B) As Boolean
Dim lobjInspecao    As clsA46V717A      'Objeto da classe clsA46V717A (TAB_PED_INSP).
    
    On Error GoTo ControleErro

    mfSalvarInspecao = False

Incluir_Inspec:
    pMensagem = " "
    Set lobjInspecao = New clsA46V717A
    Set lobjInspecao.M46V999 = gM46V999
    
    Set lobjInspecao.ConexaoBD = objPedido.ConexaoBD
    If gM46V999.gfPreenchido(pobjstcPedLoc.objstcPedInsp.NOSSO_NUMERO) Then
        pobjstcPedLoc.objstcPedInsp.NOSSO_NUMERO = pobjstcPedLoc.NOSSO_NUMERO
        pobjstcPedLoc.objstcPedInsp.Num_Item = pobjstcPedLoc.Num_Item
        If Not lobjInspecao.mfExcluir(pobjstcPedLoc.objstcPedInsp, "", "") Then
            pMensagem = lobjInspecao.UltimoErro
            GoTo Finalizar
        End If
    End If
    
    With pobjstcPedLoc.objstcPedInsp
        .NOSSO_NUMERO = pobjstcPedLoc.NOSSO_NUMERO
        .Num_Item = pobjstcPedLoc.Num_Item
        .Dat_Cadastro = Left(gM46V999.gfForData(3), 8)
        .Num_Logr_Risco = Val(txtRisNumero.Text)
        .Num_Tel = ""   'txtTelefone.Text
        .Pess_Cont = "" 'txtContato.Text
        .Cod_Ramo = objstcPedido.Cod_Ramo
        .Num_Ped = pobjstcPedLoc.Num_Ped
    End With
        
    If Not lobjInspecao.mfIncluir(pobjstcPedLoc.objstcPedInsp, "", "", "") Then
        If UCase(lobjInspecao.UltimoErro) Like "*DUPLICATE*" Then
            If Not lobjInspecao.mfExcluir(pobjstcPedLoc.objstcPedInsp, "", "") Then
                pMensagem = lobjInspecao.UltimoErro
                GoTo Finalizar
            Else
                Set lobjInspecao = Nothing
                GoTo Incluir_Inspec
            End If
        Else
            pMensagem = lobjInspecao.UltimoErro
            GoTo Finalizar
        End If
    End If
    
    mfSalvarInspecao = True
    
Finalizar:
    Set lobjInspecao = Nothing
    Exit Function
ControleErro:
    Set lobjInspecao = Nothing
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarInspecao ERRO - " & Err & " " & Error & _
                Chr(13) & pMensagem)
End Function

Private Function mfSalvarCoberturas(pMensagem As String, pobjstcPedLoc As stcA46V708B) As Boolean
Dim objPedidoCobert       As clsA46V704A
Dim lobjstcPedCobert    As stcA46V704B
Dim lI                  As Integer
Dim lI2                 As Integer
Dim lI3                 As Integer
Dim lNaoIncluirCompl    As Boolean

    On Error GoTo ControleErro
    
Incluir_Cobert:
    Set objPedidoCobert = New clsA46V704A
    Set objPedidoCobert.M46V999 = gM46V999
    
    Set lobjstcPedCobert = New stcA46V704B
    Set lobjstcPedCobert.M46V999 = gM46V999
    Set lobjstcPedCobert.ObjPai = pobjstcPedLoc
    
    If pobjstcPedLoc.ColecaoPedCobert.Count > 0 Then
        lobjstcPedCobert.NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
        lobjstcPedCobert.Num_Item = pobjstcPedLoc.Num_Item_Ant
        If Not objPedidoCobert.mfExcluir(lobjstcPedCobert, "", "") Then
            pMensagem = objPedidoCobert.UltimoErro
            GoTo Finalizar
        End If
    End If
    
    Set pobjstcPedLoc.ColecaoPedCobert = New Collection
    
    For lI3 = 1 To TabMod.Tabs.Count
        lI2 = TabMod.Tabs.Item(lI3).Tag
        For lI = 1 To flxGridCoberturaDemais(lI2).Rows - 1
            Set lobjstcPedCobert = New stcA46V704B
            Set lobjstcPedCobert.M46V999 = gM46V999
            Set lobjstcPedCobert.ObjPai = pobjstcPedLoc

            If flxGridCoberturaDemais(lI2).TextMatrix(lI, 0) <> "" Then
                 With lobjstcPedCobert
                     .NOSSO_NUMERO = pobjstcPedLoc.NOSSO_NUMERO
                     .Cod_Cobert = flxGridCoberturaDemais(lI2).TextMatrix(lI, 0)
                     
                     .Num_Item = pobjstcPedLoc.Num_Item
                     If gM46V999.gfPreenchido(flxGridCoberturaDemais(lI2).TextMatrix(lI, 7)) = False Then
                         .Per_Desc_Coml = 0
                     Else
                         .Per_Desc_Coml = flxGridCoberturaDemais(lI2).TextMatrix(lI, 7)
                     End If
                     If gM46V999.gfPreenchido(flxGridCoberturaDemais(lI2).TextMatrix(lI, 8)) = False Then
                         .Per_Desc_Exp = 0
                     Else
                         .Per_Desc_Exp = flxGridCoberturaDemais(lI2).TextMatrix(lI, 8)
                     End If
                     If gM46V999.gfPreenchido(flxGridCoberturaDemais(lI2).TextMatrix(lI, 5)) = False Then
                         .Per_Franq = 0
                     Else
                         .Per_Franq = flxGridCoberturaDemais(lI2).TextMatrix(lI, 5)
                     End If
                     If .Per_Franq > 100 Then
                        .Per_Franq = 99
                    End If
                     .Per_Taxa = 0
                     .Tab_Cobert = gM46V111.gCod_plano
                     .Val_Franq_Max = 0
                     If gM46V999.gfPreenchido(flxGridCoberturaDemais(lI2).TextMatrix(lI, 6)) = False Then
                         .Val_Franq_Min = 0
                     Else
                         .Val_Franq_Min = flxGridCoberturaDemais(lI2).TextMatrix(lI, 6)
                     End If
                     If gM46V999.gfPreenchido(flxGridCoberturaDemais(lI2).TextMatrix(lI, 2)) = False Then
                         .Val_IS = 0
                     Else
                         .Val_IS = flxGridCoberturaDemais(lI2).TextMatrix(lI, 2)
                     End If
                     If gM46V999.gfPreenchido(flxGridCoberturaDemais(lI2).TextMatrix(lI, 9)) = False Then
                         .VAL_PR = 0
                     Else
                         .VAL_PR = flxGridCoberturaDemais(lI2).TextMatrix(lI, 9)
                     End If
                     .Tab_Cobert = Val(flxGridCoberturaDemais(lI2).Tag)
                     .Val_Pr_Anu = .VAL_PR
                     .Val_Pr_Resg = .VAL_PR
                     .Val_Pr_Resg_Anu = .VAL_PR
                     .Val_Pr_Tarif = .VAL_PR
                     .Val_Pr_Tarif_Anu = .VAL_PR
                     .Cod_Ramo = objstcPedido.Cod_Ramo
                     .Num_Ped = objstcPedido.Num_Ped
                     .Num_Meses_PI = 0
                     .Val_Fator_Franq = 0
                     .Val_Risco_Cob = 0
                     .Alterar_Cobert = 0
                     .NUM_PI = gM46V111.gNum_Pi
                     If .Cod_Ramo = e_CodRamo_Condominio Or _
                        .Cod_Ramo = e_CodRamo_Empresarial Or _
                        .Cod_Ramo = e_CodRamo_Residencial Then
                        lNaoIncluirCompl = False
                    Else
                        lNaoIncluirCompl = True
                    End If
                End With
                

                If Not objPedidoCobert.mfIncluir(lobjstcPedCobert, "", "", lNaoIncluirCompl, False, "") Then
                    If UCase(objPedidoCobert.UltimoErro) Like "*DUPLICATE*" Then
                        If Not objPedidoCobert.mfExcluir(lobjstcPedCobert, "", "", lNaoIncluirCompl) Then
                            pMensagem = objPedidoCobert.UltimoErro
                            GoTo Finalizar
                        Else
                            Set lobjstcPedCobert = Nothing
                            Set objPedidoCobert = Nothing
                            GoTo Incluir_Cobert
                        End If
                    Else
                        pMensagem = objPedidoCobert.UltimoErro
                        GoTo Finalizar
                    End If
                End If
                Set lobjstcPedCobert.ObjPai = pobjstcPedLoc
                pobjstcPedLoc.ColecaoPedCobert.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
           End If
        Next lI
    Next lI3
    mfSalvarCoberturas = True
Finalizar:
    Set lobjstcPedCobert = Nothing
    Set objPedidoCobert = Nothing
    Exit Function
    
ControleErro:
    Set lobjstcPedCobert = Nothing
    Set objPedidoCobert = Nothing
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarCoberturas ERRO - " & Err & " " & Error & _
                Chr(13) & pMensagem)
End Function

Private Function mfSalvarCosseguro(pMensagem As String) As Boolean
Dim objPedidoCoss         As clsA46V706A
Dim lobjstcPedCoss      As stcA46V706B
Dim lI                  As Integer


    On Error GoTo ControleErro
    Set objPedidoCoss = New clsA46V706A
    Set objPedidoCoss.M46V999 = gM46V999
    
    Set lobjstcPedCoss = New stcA46V706B
    If objstcPedido.ColecaoPedCoss.Count > 0 Then
        lobjstcPedCoss.NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
        If Not objPedidoCoss.mfExcluir(lobjstcPedCoss, "", "", "") Then
            pMensagem = objPedidoCoss.UltimoErro
            GoTo Finalizar
        End If
    End If
    Set objstcPedido.ColecaoPedCoss = New Collection
    If frmT47V004A.optCosseguro(0).Value = True Then
        For lI = 1 To 5
            If frmT47V008A.cboConSeguradora(lI).ListIndex <> -1 Then
                Set lobjstcPedCoss = New stcA46V706B
                With lobjstcPedCoss
                    .NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
                    .Cod_Coss = frmT47V008A.cboConSeguradora(lI).ItemData(frmT47V008A.cboConSeguradora(lI).ListIndex)
                    .Nom_Seguradora = frmT47V008A.cboConSeguradora(lI).Text
                    .Per_Coss = frmT47V008A.txtCosPercentual(lI).Text
                    .Cod_Ramo = objstcPedido.Cod_Ramo
                    .Num_Ped = objstcPedido.Num_Ped
                    .NUM_PI = gM46V111.gNum_Pi
                End With
                If Not objPedidoCoss.mfIncluir(lobjstcPedCoss, "", "", "") Then
                    pMensagem = objPedidoCoss.UltimoErro
                    GoTo Finalizar
                End If
                objstcPedido.ColecaoPedCoss.Add lobjstcPedCoss, lobjstcPedCoss.Class_ID
            End If
        Next
    End If
    mfSalvarCosseguro = True
Finalizar:
    Set lobjstcPedCoss = Nothing
    Set objPedidoCoss = Nothing
    Exit Function
ControleErro:
    Set lobjstcPedCoss = Nothing
    Set objPedidoCoss = Nothing
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarCosseguro ERRO - " & Err & " " & Error & _
                Chr(13) & pMensagem)
End Function
Private Function mfSalvarCtrlEmiss(pMensagem As String) As Boolean
Dim lobjCtrlEmis As clsA46V712A
Dim lNumPI      As Long

    On Error GoTo ControleErro
    
    With objstcPedido.objstcCtrlEmis
        .NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
        If gM46V999.CDblx(.NUM_PI) = 0 Then
            If gM46V111.gNum_Pi = 0 Then
                gM46V111.gNum_Pi = Right(Year(Date), 2) & Format(objstcPedido.Num_Ped, "00000000")
            End If
        End If
        .NUM_PI = gM46V111.gNum_Pi
        If gM46V999.gfPreenchido(frmT47V001A.txtDataRecebimento.Text) Then
            .Dat_Rec_Cia = frmT47V001A.txtDataRecebimento.Text
        End If
        .Num_Bloq = Left(Format(Val(Replace(frmT47V011A.txtPagAntBloqueto.Text, "-", "")), "00000000"), "7")
        .DIG_BLOQUETO = Val(Right(frmT47V011A.txtPagAntBloqueto.Text, 1))
        .COD_UNID = objstcPedido.Cod_Unid_Prod
        .Cod_Ramo = objstcPedido.Cod_Ramo
        .Dat_Ini_Vig = objstcPedido.Dat_Ini_Vig
        .Dat_Ter_Vig = objstcPedido.Dat_Ter_Vig
        .Cod_Corr = Val(frmT47V004A.TxtCorretor(1).Tag)
        .Per_Corr = frmT47V004A.txtCorComissao(1).Text
        .Cod_Corr_2 = Val(frmT47V004A.TxtCorretor(2).Tag)
        .Per_Corr_2 = frmT47V004A.txtCorComissao(2).Text
        .Cod_Corr_3 = Val(frmT47V004A.TxtCorretor(3).Tag)
        .Per_Corr_3 = frmT47V004A.txtCorComissao(3).Text
        
        .Num_Ref_Corr = frmT47V004A.txtReferencia
        
        .Nom_Seg = objstcPedido.Nom_Segurado
        .Qtd_Anex = 0
        .Num_Apol_Ant = Val(frmT47V001A.txtApolice.Text)
        .Dsc_Obs_Prop = ""
        
        .Tip_Cotac = frmT47V004A.CmbTipCotac.ItemData(frmT47V004A.CmbTipCotac.ListIndex)
        If objstcPedido.Cod_Dizer <> 0 And .Tip_Cotac = "1" Then
            .Tip_Cotac = "3"
        End If
        .Num_Cotac = Format(gM46V999.CDblx(frmT47V004A.txtNumSiscota.Text), "000000000000")
        
'Ficha 900046
        If frmT47V004A.optPagAntecipado(1).Value = True Then
            .Cod_Iden_Pag_Antec = e_SimNao_Sim
        Else
            .Cod_Iden_Pag_Antec = e_SimNao_Nao
        End If
        If objstcPedido.Cod_Reg <> 106 And objstcPedido.Cod_Reg <> 107 And objstcPedido.Cod_Reg <> 120 And _
           objstcPedido.Cod_Reg <> 121 And objstcPedido.Cod_Reg <> 122 And objstcPedido.Cod_Reg <> 20 Then
            If objstcPedido.Qtd_Parc = 0 Then objstcPedido.Qtd_Parc = 1
        End If
        If objstcPedido.Qtd_Parc >= 2 Then
            .Per_Juros_Neg = frmT47V004A.txtJuros.Text
        Else
            If .Cod_Iden_Pag_Antec = e_SimNao_Sim Then
                .Per_Juros_Neg = 0
            Else
                .Per_Juros_Neg = frmT47V004A.txtJuros.Text
            End If
        End If
    
        .Cod_Envi_Doc = "C"
        If .DAT_DIGIT = 0 Then
            .DAT_DIGIT = gM46V999.gHoje
            .HOR_DIGIT = objstcPedido.HOR_DIGIT & "00"
        End If
        .Cod_Iden_Stat_Acao = 0
        .Cod_User_Digit = objstcPedido.Cod_User_Digit
        .Dat_Digit_MF = gM46V999.gHoje
        .Hor_Digit_MF = objstcPedido.HOR_DIGIT & "00"
        .Cod_User_Digit_MF = ""
        .Val_Pr_Tot = objstcPedido.gValorTotal
        .Cod_Iden_Classe = 0
        .Cod_Iden_Cobr = 0
        .Num_Prop = objstcPedido.Num_Ped
        .Dat_Canc_Apol = 0
        Select Case objstcPedLoc.Tip_Emissao
            Case e_TipEmissao_SeguroNovo
                .Cod_Tip_Doc = 1
            Case e_TipEmissao_RenYasuda
                .Cod_Tip_Doc = 3
            Case Else
                .Cod_Tip_Doc = 2
        End Select
        .Dat_Emis = 0
        .Num_Apol = 0
        .Num_Endo = 0
        .Val_Pr_Emis = objstcPedido.VAL_PR
        .Val_Pr_Liq = objstcPedido.VAL_PR
        If frmT47V004A.optCosseguro(0).Value = True Then
            .Per_Tot_Coss = CDbl(frmT47V008A.txtCosPercentual(1).Text) + CDbl(frmT47V008A.txtCosPercentual(2).Text) + _
                            CDbl(frmT47V008A.txtCosPercentual(3).Text) + CDbl(frmT47V008A.txtCosPercentual(4).Text) + _
                            CDbl(frmT47V008A.txtCosPercentual(5).Text)
        Else
            .Per_Tot_Coss = 0
        End If
        .Cod_Iden_Stat_Resp = 0
        .Cod_Iden_SAC = 0
        .Cod_Prod = objstcPedido.Cod_Prod
        .Val_IS = objstcPedido.Val_IS
        .Per_Desc_Coml = 0
        
        .Num_Cgc_Cpf = objstcPedido.Num_Cgc_Cpf
        .Cod_Iden_Pessoa = objstcPedido.Tip_Pessoa
        
        .Num_Logr_Cobranca = CDblx(frmT47V004A.TxtNumCobr.Text)
        .Num_Logr_Segurado = CDblx(frmT47V004A.TxtNumResid.Text)
        .Dsc_Compl_Cobranca = frmT47V004A.TxtComplCobr.Text
        .Dsc_Compl_Segurado = frmT47V004A.TxtComplResid.Text
        
    End With
    Set lobjCtrlEmis = New clsA46V712A
    Set lobjCtrlEmis.M46V999 = gM46V999
    
    Set lobjCtrlEmis.ConexaoBD = objPedido.ConexaoBD
    
    'Se for inclusao
    If mInclusao = True Then
        If Not lobjCtrlEmis.mfIncluir(objstcPedido.objstcCtrlEmis, "", "", "") Then
            pMensagem = lobjCtrlEmis.UltimoErro
            GoTo Finalizar
        End If
    Else
        If Not lobjCtrlEmis.mfAlterar(objstcPedido.objstcCtrlEmis, "", "") Then
            pMensagem = lobjCtrlEmis.UltimoErro
            GoTo Finalizar
        End If
    End If
    mfSalvarCtrlEmiss = True
Finalizar:
    Set lobjCtrlEmis = Nothing
    Exit Function
ControleErro:
    Set lobjCtrlEmis = Nothing
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarCtrlEmis ERRO - " & Err & " " & Err.Description & _
                Chr(13) & pMensagem)
End Function

Private Function mfSalvarComum(pMensagem As String) As Boolean
'********************Vari�veis padr�o*************************************
Dim lCkPoint        As String      'Variavel utilizada para determinar pontos onde o c�digo j� passou.
Dim lCkStr          As String       'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
                                    'lCkPoint. (posi��o das variaveis em ordem de declara��o)
'********************Vari�veis Declaradas para o processamento************

Dim lSaida      As String       'Sa�da
Dim lAuxiliar   As String
Dim lNovoPedido As String
Dim lAuxDig     As String
Dim lNossoNumero As String
    Err = 0
    lCkPoint = ""
    lCkStr = ""
    On Error GoTo ControleErro
    lCkPoint = "1"
    If (gM46V999.gCodProduto & gM46V999.gCodSequencia) > objstcPedido.Ult_Tarifa Then
        objstcPedido.Ult_Tarifa = gM46V999.gCodProduto & gM46V999.gCodSequencia
    End If
    lCkPoint = "2"
    With objstcPedido
        '===Ramo
        .Num_Apol_An = Val(frmT47V001A.txtApolice.Text)
        
        .Cod_Reg = Mid(frmT47V016A.cboCodReg.Text, 1, 3)
        
        .Dat_Emis = 0
        If gM46V999.gfPreenchido(frmT47V004A.txtDatEmis.Text) And gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
            .Dat_Emis = gM46V999.gfForData(1, frmT47V004A.txtDatEmis.Text)
        End If
              
        .Num_Apol = 0
        .Num_Endo = 0
        
        .Cod_Bco_Cobr = 0
        '.Cod_Agen_Angar
        .Cod_Agen_Indicr = 0
        .Num_Cta_Indicr = 0
        lCkPoint = "3"
        If frmT47V004A.optForPagamento(1).Value = True Then 'se d�bito , salvo os dados de d�bito
            If frmT47V010A.cboDebBanco.ListIndex <> -1 Then
                .Cod_Bco_Cobr = frmT47V010A.cboDebBanco.ItemData(frmT47V010A.cboDebBanco.ListIndex)
            Else
                .Cod_Bco_Cobr = 215
            End If
            If .Cod_Bco_Cobr = 237 Then
                .Cod_Dig_Agen_Bas_Cor = Right(Format(Val(Replace(frmT47V010A.txtDebAgencia.Text, "-", "")), "00000"), 1)
                .Cod_Agen_Bas_Cor = Left(Format(Val(Replace(frmT47V010A.txtDebAgencia.Text, "-", "")), "00000"), 4)
            Else
                .Cod_Dig_Agen_Bas_Cor = 0
                .Cod_Agen_Bas_Cor = Val(Replace(frmT47V010A.txtDebAgencia.Text, "-", ""))
            End If
            If gM46V999.gfPreenchido(frmT47V010A.txtDebAgencia.Text) Then
                If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
            'Ficha 1000833
                    If .Cod_Bco_Cobr = 399 Then
                        lAuxiliar = gM46V999.CDblx(Mid(frmT47V010A.txtDebConta.Text, 1, InStr(frmT47V010A.txtDebConta.Text, "-") - 1))
                        lAuxDig = gM46V999.CDblx(Mid(frmT47V010A.txtDebConta.Text, InStr(frmT47V010A.txtDebConta.Text, "-") + 1, 1))
                        .Num_Cta_Bas_Cor = Mid(lAuxiliar, 1, 7) & lAuxDig
                    Else
                        lAuxiliar = gM46V999.CDblx(Mid(frmT47V010A.txtDebConta.Text, 1, InStr(frmT47V010A.txtDebConta.Text, "-") - 1))
                        .Num_Cta_Bas_Cor = Mid(lAuxiliar, 1, 7)
                    End If
                Else
                    .Num_Cta_Bas_Cor = Val(Replace(frmT47V010A.txtDebConta.Text, "-", ""))
                End If
            End If
        Else
            .Cod_Bco_Cobr = 215
               'Carn�
            .Cod_Agen_Bas_Cor = 0
            .Num_Cta_Bas_Cor = 0
            .Cod_Dig_Agen_Bas_Cor = 0
            If frmT47V009A.cboCarBanco.ListIndex <> -1 Then
                .Cod_Bco_Cobr = frmT47V009A.cboCarBanco.ItemData(frmT47V009A.cboCarBanco.ListIndex)
            Else
                .Cod_Bco_Cobr = 215
            End If
        End If
        lCkPoint = "4"
        '===Unidade
        .Cod_Unid_Prod = IIf(Val(frmT47V004A.lblNomUnidade.Tag) < 1000, 0, Val(frmT47V004A.lblNomUnidade.Tag))
        '===Produtor
        .Cod_Prod = Val(frmT47V004A.lblNomProdutor.Tag)
        '.Cod_Agen_Bas_Cor
        '.Num_Cta_Bas_Cor
        
        
        '[FRH.08] ->
        '01.[N�o carregar C�digo do SisCAD, habilitar campo "c�d. Segurado"]
        '02.[Validar c�digo digitado com o GED.TAB_CLIENTE.COD_CLI]
        '03.[Calcular digito e adicionar ao objstcPedido.Dig_Cod_Segurado]
                             
        If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        
            Dim lZerar_campoCodigo As Boolean: lZerar_campoCodigo = False
            
            If Left(Val(frmT47V004A.txtCodSegurado.Text), 4) = 7777 And Val(frmT47V004A.TxtCorretor(1).Tag) <> 1601 Then
                lZerar_campoCodigo = True
            Else
        
                Dim bdGed As ADODB.Connection
                Dim lrst As clsYasRecordSet
                Dim lSelect As String
                Dim lMensagem As String
                Dim lCpfCnpj As String
            
                '[Verifica se PF/PJ]
                If frmT47V004A.optTipPessoa(0).Value = True Then '[Pessoa f�sica]
                    Call gM46V999.gpForCPF(2, frmT47V004A.txtCPF.Text, lCpfCnpj)
                Else
                    Call gM46V999.gpForCNPJ(2, frmT47V004A.txtCNPJ.Text, lCpfCnpj)
                End If
                        
                '[Abrir base GED.tab_cliente]
                If Not gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, bdGed, lMensagem) Then Call gM46V999.gpGraLog(1, lMensagem)
                        
                '[Verificar se CPF/CNPJ dif. de 0 e C�d.= 0]
                If gM46V999.gfPreenchido(lCpfCnpj) And Trim(lCpfCnpj) <> "0" And Val(frmT47V004A.txtCodSegurado.Text) = 0 Then
    
                  '[Obter os C�digos cadastrados na GED de acordo com o CPF/CNPJ digitado]
                    lSelect = "SELECT COD_CLI FROM tab_cliente where CGC_CPF='" & Trim(lCpfCnpj) & "'"
                    If Not gM46V999.gfObtRegistro(bdGed, lSelect, lrst, lMensagem) Then Call gM46V999.gpGraLog(1, lMensagem)
            
                    '[Caso v�rios c�digos retornados, selecionar o primeiro]
                    If Not (lrst.EOF And lrst.BOF) Then
                        lrst.MoveLast: lrst.MoveFirst
                        If lrst.RecordCount = 1 Then
                            frmT47V004A.txtCodSegurado.Text = IIf(Len(lrst("Cod_Cli")) > 6, 0, Val(lrst("Cod_Cli")))
                        End If
                    End If
                  
                '[Verificar se c�d. dif. de 0 no campo]
                ElseIf Val(frmT47V004A.txtCodSegurado.Text) <> 0 Then
                
                    '[Obter os CPFs/CNPJs cadastrados na GED com o c�digo digitado]
                    lSelect = "SELECT * FROM tab_cliente where COD_CLI=" & CStr(gM46V999.gfTratarNulo(frmT47V004A.txtCodSegurado.Text, tpoNumerico))
                    If Not gM46V999.gfObtRegistro(bdGed, lSelect, lrst, lMensagem) Then Call gM46V999.gpGraLog(1, lMensagem)
                                    
                    '[C�digo inv�lido:]
                    If lrst.EOF And lrst.BOF Then
                        lZerar_campoCodigo = True
                    End If
                End If
            End If
                
            If lZerar_campoCodigo Then
                frmT47V004A.txtCodSegurado.Text = 0
                .Cod_Segurado = frmT47V004A.txtCodSegurado.Text
                frmT47V004A.txtDigCodSegurado.Text = ""
                .Dig_Cod_Segurado = frmT47V004A.txtDigCodSegurado.Text
            Else
                .Cod_Segurado = Val(frmT47V004A.txtCodSegurado.Text)
                If .Cod_Segurado <> 0 Then
                    '[Calcular o d�gito]
                    Dim lDigito As String
                    Dim clsA99V001 As clsA99V001
                    Set clsA99V001 = New clsA99V001
                    Set clsA99V001.M46V999 = gM46V999
                    
                    If clsA99V001.DIG_YASUDA(.Cod_Segurado, lDigito, pMensagem) = False Then Exit Function
                    
                    .Dig_Cod_Segurado = CStr(lDigito)
                    frmT47V004A.txtDigCodSegurado.Text = CStr(lDigito)
                    
                    Set clsA99V001 = Nothing
                Else
                    .Dig_Cod_Segurado = ""
                End If
            End If

            '[Limpar Mem�ria]
            lSaida = ""
            lSelect = ""
            Set lrst = Nothing
            Set bdGed = Nothing
        Else
            '[Continuar as rotinas originais]
            .Cod_Segurado = Val(frmT47V004A.txtCodSegurado.Text)
            
            frmT47V004A.txtDigCodSegurado.Visible = False
            .Dig_Cod_Segurado = ""
        End If
        '<-[FRH]
        '===Nome do segurado.
        .Nom_Segurado = Trim(txtSegurado.Text)
        .Dsc_Compl_Segurado = ""
        
        lCkPoint = "5"
        '===CPF ou CNPJ
        If frmT47V004A.optTipPessoa(0).Value = True Then 'F�sica
                                
            .Tip_Pessoa = e_TipPessoa_Fisica
            Call gM46V999.gpForCPF(2, frmT47V004A.txtCPF.Text, lSaida)
            .Num_Cgc_Cpf = lSaida
            If frmT47V004A.cboIdeTipDocumento.ListIndex <> -1 Then
                .Cod_Doc = frmT47V004A.cboIdeTipDocumento.ItemData(frmT47V004A.cboIdeTipDocumento.ListIndex)
            Else
                .Cod_Doc = 0
            End If
            .Num_Doc = frmT47V004A.txtIdeNumDocumento.Text
            .Org_Exp = frmT47V004A.txtIdeOrgExpedidor.Text
            .Dat_Emis_Doc = gM46V999.gfForData(1, frmT47V004A.txtIdeDatEmissao.Text)
            .Nom_Empresa = frmT47V004A.txtIdeEmpresa.Text
            If frmT47V004A.optSexo(0).Value = True Then
                objstcPedido.Sexo = "M"
            ElseIf frmT47V004A.optSexo(1).Value = True Then
                objstcPedido.Sexo = "F"
            Else
                objstcPedido.Sexo = ""
            End If
            .Cod_Ramo_Ativ = 0
        Else
            Call gM46V999.gpForCNPJ(2, frmT47V004A.txtCNPJ.Text, lSaida)
            .Num_Cgc_Cpf = lSaida
            .Tip_Pessoa = e_TipPessoa_Juridica
            If frmT47V004A.cboRamAtividade.ListIndex <> -1 Then
                .Cod_Ramo_Ativ = frmT47V004A.cboRamAtividade.ItemData(frmT47V004A.cboRamAtividade.ListIndex)
            Else
                .Cod_Ramo_Ativ = 0
            End If
            .Cod_Doc = 0
            .Num_Doc = ""
            .Org_Exp = ""
            .Dat_Emis_Doc = 0
            .Nom_Empresa = ""
            .Sexo = ""
        End If

        lCkPoint = "6"
        If frmT47V004A.optEndCobranca(0).Value = True Then
            '===Tipo de logradouro.
            .Cod_Tip_Logr = frmT47V006A.txtTipLogradouro.Text
            '===Logradouro.
            .Nom_Logr = frmT47V006A.txtNomLogradouro.Text
            '===Complemento do logradouro.
            .Dsc_Compl_Logr = frmT47V006A.txtComplemento.Text
            '===Bairro.
            .Nom_Bairro = frmT47V006A.txtBairro.Text
            '===CIdade.
            .Nom_Cid = frmT47V006A.txtCidade.Text
            '===UF.
            If frmT47V006A.cboUF.ListIndex <> -1 Then
                .Sig_UF = frmT47V006A.cboUF.List(frmT47V006A.cboUF.ListIndex)
            Else
                .Sig_UF = ""
            End If
            lCkPoint = "7"
            '===CEP.
            .Num_CEP = Val(Replace(frmT47V006A.txtCEP.Text, "-", ""))
        Else
                    '===Tipo de logradouro.
            .Cod_Tip_Logr = frmT47V004A.txtTipLogradouro.Text
            '===Logradouro.
            .Nom_Logr = frmT47V004A.txtNomLogradouro.Text
            '===Complemento do logradouro.
            .Dsc_Compl_Logr = frmT47V004A.txtComplemento.Text
            '===Bairro.
            .Nom_Bairro = frmT47V004A.txtBairro.Text
            '===CIdade.
            .Nom_Cid = frmT47V004A.txtCidade.Text
            '===UF.
            If frmT47V004A.cboUF.ListIndex <> -1 Then
                .Sig_UF = frmT47V004A.cboUF.List(frmT47V004A.cboUF.ListIndex)
            Else
                .Sig_UF = ""
            End If
            lCkPoint = "7A"
            '===CEP.
            .Num_CEP = Val(Replace(frmT47V004A.txtCEP.Text, "-", ""))
        End If
        '===Telefone
        lAuxiliar = Replace(frmT47V004A.txtTelefones.Text, "-", "")
        lAuxiliar = Replace(lAuxiliar, "(", "")
        lAuxiliar = Replace(lAuxiliar, ")", "")
        .Num_Fone = Val(lAuxiliar)
        
        If gM46V111.gTipoGrupMulti = e_GrupoRamo_Multi And panItemLocRisco.Enabled = True Then
            If cboPlano(1).ListIndex > -1 Then
                .Cod_Estip = cboPlano(1).ItemData(cboPlano(1).ListIndex)
            Else
                .Cod_Estip = 0
            End If
        Else
            .Cod_Estip = 0
        End If
        .Desc_Plano = cboPlano(1).Text
        '.Val_Coef_Cms_Pad
        '===Data de in�cio da vig�ncia.
        If gM46V999.gfPreenchido(txtVigInicio.Text) Then
            .Dat_Ini_Vig = gM46V999.gfForData(1, txtVigInicio.Text)
        Else
            .Dat_Ini_Vig = 0
        End If
        '===Data de t�rmino da vig�ncia.
        If gM46V999.gfPreenchido(txtVigTermino.Text) Then
            .Dat_Ter_Vig = gM46V999.gfForData(1, txtVigTermino.Text)
        Else
            .Dat_Ter_Vig = 0
        End If
        
        If Not gfCarregaPerIOF(.Dat_Ini_Vig, pMensagem) Then Exit Function
        
        '.Dat_Emis
        '.Cod_Moed
        '===Juros. (utilizado os juros daa tabela TAB_PED_JUROS
        lCkPoint = "8"
        If IsNumeric(frmT47V004A.txtJuros.Text) Then
            .Per_Juro_Frac = CDbl(frmT47V004A.txtJuros.Text)
        Else
            .Per_Juro_Frac = 0
        End If
        '===Quantidade de parcelas.
        .Qtd_Parc = Val(frmT47V004A.txtQuaParcelas.Text)
        If .Qtd_Parc = 0 And _
          (objstcPedido.Cod_Reg <> 106 And objstcPedido.Cod_Reg <> 107 And objstcPedido.Cod_Reg <> 120 And _
           objstcPedido.Cod_Reg <> 121 And objstcPedido.Cod_Reg <> 122 And objstcPedido.Cod_Reg <> 20) Then
            .Qtd_Parc = 1
            frmT47V004A.txtQuaParcelas.Text = 1
        End If
        If lVerTipEmissao = "E" Then
            .Qtd_Anexo = 0
        Else
            .Qtd_Anexo = objstcPedido.ColecaoPedLoc.Count
        End If
        '===Tipo de prazo.
        If frmT47V003A.optTipPrazo(1).Value = True Then
            .Cod_Iden_Praz = e_TipoPrazo_ProRata
        ElseIf frmT47V003A.optTipPrazo(0).Value = True Then
            .Cod_Iden_Praz = e_TipoPrazo_PrazoCurto
        ElseIf frmT47V003A.optTipPrazo(2).Value = True Then
            .Cod_Iden_Praz = e_TipoPrazo_PrazoLongo
        Else
            .Cod_Iden_Praz = e_TipoPrazo_ProRata
        End If
        .Cod_Pr_Min_Endo = 1
        .Cod_Pr_Min_Parc = 1
        If gM46V999.gCod_Ramo = 113 Then
            If gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031 Then
                .Cod_Pr_Min_Parc = 2
            End If
        End If
        If (frmT47V004A.optPagAntecipado(2).Value = True) Or (frmT47V004A.optPagAntecipado(1).Value = True And frmT47V004A.optForPagamento(1).Value = True) Then
            .Val_Pr_Antec = 0
        Else
            If gM46V999.CDblx(frmT47V004A.txtQuaParcelas.Text) = 0 Then
                .Val_Pr_Antec = 0
            ElseIf gM46V999.CDblx(frmT47V004A.txtValorPgAnt.Text) = 0 Then
                .Val_Pr_Antec = 1
            Else
                .Val_Pr_Antec = gM46V999.CDblx(frmT47V004A.txtValorPgAnt.Text)
            End If
        End If
        '.Per_Adic_Form_Pag
        '.Per_Desc_Form_Pag
        'Simony
        '====Valor da LMG
        If .Cod_Ramo <> e_CodRamo_Automoveis Then
            If gM46V999.gfPreenchido(txtISBas.Text) = False Then txtISBas.Text = "0,00"
            .Val_IS = txtISBas.Text
        End If
        
        '.Val_IS_2
        If .Cod_Ramo <> e_CodRamo_Automoveis Then
            If gM46V999.gfPreenchido(TxtValor4.Text) = False Then TxtValor4.Text = "0,00"
            .VAL_PR = TxtValor4.Text
        End If
        
        '====Valor do Pr�mio Resseguro or Valor LMG 2
        If .Cod_Ramo <> e_CodRamo_Automoveis Then
            If gM46V999.gfPreenchido(TxtValor2.Text) = False Then TxtValor2.Text = "0,00"
            .Val_Pr_Ref_Resg = TxtValor2.Text
        End If
        
        '====Valor do Pr�mio Tarif�rio
        If .Cod_Ramo <> e_CodRamo_Automoveis Then
            If gM46V999.gfPreenchido(TxtValor3.Text) = False Then TxtValor3.Text = "0,00"
            .Val_Pr_Ref_Tarif = TxtValor3.Text
        End If
        
        '====In�cio da Viagem
        If .Cod_Ramo <> e_CodRamo_Automoveis Then
            If gM46V999.gfPreenchido(TxtValor7.Text) Then
                .Dat_Ini_Viag = gM46V999.gfForData(1, TxtValor7.Text)
            End If
        End If
        
        '.Cod_Tip_Liber
        '.Cod_Iden_Alt_Corr
        '.Cod_User_Digit
        .DAT_DIGIT = Left(gM46V999.gfForData(3), 8)
        .HOR_DIGIT = Mid(gM46V999.gfForData(3), 9, 4)
        .Cod_User_Liber = gM46V999.gNomUsuario6
        .Dat_Liber = Left(gM46V999.gfForData(3), 8)
        .Hor_Liber = Mid(gM46V999.gfForData(3), 9, 4)
        .Cod_Sit_Ped = e_SitPedido_SemCalculo
        lCkPoint = "9"
        
        Select Case .Cod_Reg
            Case Is < 100
                .Cod_Tip_Doc = 0
            Case 101 To 105
                .Cod_Tip_Doc = 1
            Case 106, 107
                .Cod_Tip_Doc = 2
            Case 120, 121, 122, 129
                .Cod_Tip_Doc = 3
        End Select
        
        .Cod_Dizer = gM46V111.gCod_Dizer
        .Qtd_Copia = 3
        
        If Mid(.Num_Cgc_Cpf, 1, 8) = "59291534" Then
            If gM46V999.gCod_Ramo = 112 Or gM46V999.gCod_Ramo = 150 Or gM46V999.gCod_Ramo = 710 Then
               .Qtd_Copia = 4
            End If
        End If
        
        If gM46V999.gfPreenchido(frmT47V015A.txtCusEmissao.Text) = False Then frmT47V015A.txtCusEmissao.Text = "0"
        .Val_Cust_Emis = frmT47V015A.txtCusEmissao.Text
        
        If lVerTipEmissao = "E" Then
            '.Val_IOF = 0
            'gProJurIOF = 0
            .Val_Cust_Emis = 0
            frmT47V015A.txtCusEmissao.Text = 0
        End If
        .Val_IOF = Round((.VAL_PR + .Val_Cust_Emis) * (gM46V111.gProJurIOF / 100), 2)
        
        
        If .Cod_Reg = 101 And .Val_Cust_Emis = 0 Then .Cod_Reg = 104
        If .Cod_Reg = 103 And .Val_Cust_Emis = 0 Then .Cod_Reg = 102
        
        gCod_Registro = .Cod_Reg
        If .VAL_PR = 0 Then .Cod_Reg = 120
        
        If .VAL_PR = 0 And gM46V111.gCod_Dizer = 0 Then
            .Cod_Reg = 20
        End If
        If .VAL_PR = 0 And gM46V111.gCod_Dizer <> 0 Then
            .Cod_Reg = 120
        End If
        
        
        If .Cod_Reg = 2 Or .Cod_Reg = 4 Or .Cod_Reg = 102 Or .Cod_Reg = 104 Or .Cod_Reg = 120 Or .Cod_Reg = 106 Or .Cod_Reg = 107 Or .Cod_Reg = 20 Then
            .Val_Cust_Emis = 0
        End If
        If .Cod_Reg = 3 Or .Cod_Reg = 103 Or .Cod_Reg = 20 Or .Cod_Reg = 120 Or .Cod_Reg = 106 Or .Cod_Reg = 107 Then
            .Val_IOF = 0
        End If
        
        .Cod_Posi_Atual = "T"
        If objstcPedido.mStatusDivSiscota = 0 Or gM46V999.gTipoAcesso = DIGITACAO Then
            mnuSituacao.Caption = Chr(183) & Space(10) & "PI: " & objstcPedido.objstcCtrlEmis.NUM_PI & " -  N�o consistido"
        End If
        
        '.Dat_Mov_Ini
        '.Dat_Mov_Ter
        If frmT47V015A.optCusEmissao(0).Value Then 'Sim.
'Ficha 900046
            If frmT47V004A.optPagAntecipado(2).Value = True Then    '30 dias
                .Tip_Cust_Emis = 1
            ElseIf frmT47V004A.optPagAntecipado(1).Value = True And frmT47V004A.optForPagamento(0).Value = True Then    'boleto
                .Tip_Cust_Emis = 5
            ElseIf frmT47V004A.optPagAntecipado(1).Value = True And frmT47V004A.optForPagamento(1).Value = True Then    'ant com d�bito
                .Tip_Cust_Emis = 3
            End If
        Else    'n�o
            If frmT47V004A.optPagAntecipado(2).Value = True Then  '30 dias
                .Tip_Cust_Emis = 0
            ElseIf frmT47V004A.optPagAntecipado(1).Value = True And frmT47V004A.optForPagamento(0).Value = True Then    'boleto
                .Tip_Cust_Emis = 4
            ElseIf frmT47V004A.optPagAntecipado(1).Value = True And frmT47V004A.optForPagamento(1).Value = True Then    'ant com debito
                .Tip_Cust_Emis = 2
            End If
        End If
        If .Val_Cust_Emis = 0 And (.Cod_Reg = 106 Or .Cod_Reg = 107 Or .Cod_Reg = 120 Or .Cod_Reg = 20) Then
            .Tip_Cust_Emis = 0
        End If
        '====In�cio da Viagem
        If gM46V999.gfPreenchido(TxtValor5.Text) = False Then TxtValor5.Text = "0"
        .Val_Cambio_Emis = TxtValor5.Text
        
        '.Iden
        
        'Verifico o nosso n�mero e o n�mero do pedido
        mInclusao = False
        If gM46V999.gfPreenchido(.NOSSO_NUMERO) = False Then
            gM46V111.gComNosNumero = ""
            
            Call gM46V999.gfCalNosNumero(lNossoNumero, lMensagem)
            gM46V111.gComNosNumero = lNossoNumero
            .NOSSO_NUMERO = gM46V111.gComNosNumero
            Call gM46V999.gfCalNumPedido(lNovoPedido, lMensagem)
            .Num_Ped = lNovoPedido
            Set objPedido = Nothing
            mInclusao = True
            gM46V111.gNumPedido = .Num_Ped
            gM46V999.gPrefixo = "Proposta " & Format$(gM46V111.gNumPedido, "000000") & " - "
        End If
            
                
        If gM46V111.gNum_Pi = 0 Then
            gM46V111.gNum_Pi = txtNumPI.Text
        End If
            
        .NUM_PI = gM46V111.gNum_Pi
        
        If gM46V999.gfPreenchido(txtCoefISVR.Text) = False Then txtCoefISVR.Text = "0"
        .Val_Coef_Cms_Pad = txtCoefISVR.Text
        
        .Cod_Produto = gM46V999.gCodProduto
        .Cod_Protocolo = 0
        .Cod_Sequencia = 0
        .Ind_CondGerais = 1
    End With
    lCkPoint = "10"
    If objPedido Is Nothing Then
        Set objPedido = New clsA46V702A
        Set objPedido.M46V999 = gM46V999
    End If
    If mInclusao Then
        If Not objPedido.mfIncluir(objstcPedido, "", "", "") Then
            pMensagem = objPedido.UltimoErro
            Exit Function
        End If
    Else
        If Not objPedido.mfAlterar(objstcPedido, "", "", mComUltData, mComUltUsuario) Then
            pMensagem = objPedido.UltimoErro
            Exit Function
        End If
    End If
    lCkPoint = "11"
    mComUltData = Left(gM46V999.gfForData(3), 8)
    mComUltUsuario = gM46V999.gNomUsuario6
    lCkPoint = "11A"
    'Salvar ctrl_emiss
    If Not mfSalvarCtrlEmiss(pMensagem) Then Exit Function
    lCkPoint = "12"
    'Salvo o endere�o
    If Not mfSalvarEndereco(pMensagem) Then Exit Function
    lCkPoint = "13A"
    'simony
    'Salvar Auto
    If objstcPedido.Cod_Ramo = e_CodRamo_Automoveis Then
        If Not mfSalvarAuto(pMensagem) Then Exit Function
    End If
    lCkPoint = "13B"
    'Salvo os corretores
    If Not mfSalvarCorretores(pMensagem) Then Exit Function
    lCkPoint = "14"
    'Salvo os cosseguros
    If Not mfSalvarCosseguro(pMensagem) Then Exit Function
    lCkPoint = "15"
    'Salvar Texto para Endosso
    If Not mfSalvarTextoEndosso(pMensagem) Then Exit Function
    lCkPoint = "16"
    'Salvar Especificacao
    If Not mfSalvarTextoEspecPadrao(pMensagem) Then Exit Function
    lCkPoint = "17"
    'Salvar vencimento
    If Not mfSalvarVencimento(pMensagem) Then Exit Function
    lCkPoint = "18"
    'Salvo os prolabore e agenciamento
    If Not mfSalvarProAgenc(pMensagem) Then Exit Function
    lCkPoint = "19"
    
    mfSalvarComum = True
    Exit Function
ControleErro:
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarComum ERRO - " & Err & " " & Err.Description & _
                Chr(13) & "CheckPoint " & lCkPoint & " " & pMensagem)
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
End Function
Private Function mfSalvarVencimento(pMensagem As String) As Boolean
Dim lObjPedVenc         As clsA46V713A
Dim lObjstcPedVenc      As stcA46V713B
Dim lI                  As Integer
    
    On Error GoTo ControleErro
    
    Set lObjPedVenc = New clsA46V713A
    Set lObjPedVenc.M46V999 = gM46V999
    Set lObjstcPedVenc = New stcA46V713B
    
    
    If objstcPedido.ColecaoPedVenc.Count > 0 Then
        lObjstcPedVenc.NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
        If Not lObjPedVenc.mfExcluir(lObjstcPedVenc, "", "", "") Then
            pMensagem = lObjPedVenc.UltimoErro
            GoTo Finalizar
        End If
    End If
    Set objstcPedido.ColecaoPedVenc = New Collection
    If frmT47V004A.optVencimento(0).Value = True Then
        If gM46V999.gfDisponivel("frmT47v019A") Then
            If frmT47v019A.txtDtVenc(1).Text <> "" Then
                
                For lI = 1 To Val(frmT47V004A.txtQuaParcelas.Text) '<- originalmente 10
                
                    If gM46V999.gfPreenchido(frmT47v019A.txtDtVenc(lI).Text) And frmT47v019A.txtDtVenc(lI).Enabled = True Then
'Ficha 900046
                        If lI = 1 And frmT47V004A.optPagAntecipado(1).Value = True Then
                            frmT47v019A.txtDtVenc(1).Text = Format(Date + 8, "dd/mm/yyyy")
                        End If
                        Set lObjstcPedVenc = New stcA46V713B
                        lObjstcPedVenc.NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
                        lObjstcPedVenc.Cod_Ramo = objstcPedido.Cod_Ramo
                        lObjstcPedVenc.Num_Ped = objstcPedido.Num_Ped
                        lObjstcPedVenc.Num_Parc = lI
                        lObjstcPedVenc.DAT_VENC = gM46V999.gfForData(1, Format(frmT47v019A.txtDtVenc(lI).Text, "dd/mm/yyyy"))
                        If Not lObjPedVenc.mfIncluir(lObjstcPedVenc, "", "", "") Then
                            pMensagem = lObjPedVenc.UltimoErro
                            GoTo Finalizar
                        End If
                        objstcPedido.ColecaoPedVenc.Add lObjstcPedVenc, lObjstcPedVenc.Class_ID
                    End If
                Next lI
            End If
        End If
    End If
    mfSalvarVencimento = True
Finalizar:
    Set lObjPedVenc = Nothing
    Set lObjstcPedVenc = Nothing
    Exit Function
ControleErro:
    Set lObjPedVenc = Nothing
    Set lObjstcPedVenc = Nothing
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarVencimento ERRO - " & Err & " " & Err.Description & _
                Chr(13) & pMensagem)
End Function
Private Function mfSalvarCorretores(pMensagem As String) As Boolean
Dim objPedidoCorr         As clsA46V705A
Dim lObjstcPedCorr      As stcA46V705B
Dim lI                  As Integer
Dim lJ                  As Integer
Dim lGravaPerCorr       As Boolean

    On Error GoTo ControleErro
    gM46V999.gBaseLocal = False
    Set objPedidoCorr = New clsA46V705A
    Set objPedidoCorr.M46V999 = gM46V999
    
    Set lObjstcPedCorr = New stcA46V705B
    If objstcPedido.ColecaoPedCorr.Count > 0 Then
        lObjstcPedCorr.NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
        
        If Not objPedidoCorr.mfExcluir(lObjstcPedCorr, "", "", "") Then
            pMensagem = objPedidoCorr.UltimoErro
            GoTo Finalizar
        End If
    End If
    
    Set objstcPedido.ColecaoPedCorr = New Collection
   
    For lJ = 1 To 3
        gCod_RamoCorr = 0
        lGravaPerCorr = True
        If gM46V999.gCod_Ramo <> e_CodRamo_Automoveis Then
            gCod_RamoCorr = objstcPedido.Cod_Ramo
        Else
            If lJ = 1 Then
                gCod_RamoCorr = e_CodRamoAuto
                If objstcPedido.objstcPedAuto.Val_Pr_Auto = 0 Then
                    lGravaPerCorr = False
                End If
            ElseIf lJ = 2 And objstcPedido.objstcPedAuto.Val_Pr_Rc <> 0 Then
                gCod_RamoCorr = e_CodRamoRC
            ElseIf lJ = 3 And objstcPedido.objstcPedAuto.Val_Pr_App <> 0 Then
                gCod_RamoCorr = e_CodRamoapp
            End If
        End If
        If gCod_RamoCorr <> 0 Then
            For lI = 1 To 3
                If Val(frmT47V004A.TxtCorretor(lI).Tag) <> 0 Then
                    Set lObjstcPedCorr = New stcA46V705B
                    With lObjstcPedCorr
                        .NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
                        .Cod_Corr = frmT47V004A.TxtCorretor(lI).Tag
                        .Num_Ordem_Corr = lI
                        If lGravaPerCorr Then .Per_Corr = gM46V999.CDblx(frmT47V004A.txtCorComissao(lI).Text)
                        .Cod_Ramo = gCod_RamoCorr
                        .Num_Ped = objstcPedido.Num_Ped
                        .NUM_PI = gM46V111.gNum_Pi
                    End With
                    If Not objPedidoCorr.mfIncluir(lObjstcPedCorr, "", "", "") Then
                        pMensagem = objPedidoCorr.UltimoErro
                        GoTo Finalizar
                    End If
                    objstcPedido.ColecaoPedCorr.Add lObjstcPedCorr, lObjstcPedCorr.Class_ID
                End If
            Next lI
        End If
        If gM46V999.gCod_Ramo <> e_CodRamo_Automoveis Then
            Exit For
        End If
    Next lJ
            
SelecionaRamosAuto:
Terminou:
    mfSalvarCorretores = True
Finalizar:
    Set lObjstcPedCorr = Nothing
    Set objPedidoCorr = Nothing
    Exit Function
    
ControleErro:
    Set lObjstcPedCorr = Nothing
    Set objPedidoCorr = Nothing
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarCorretores ERRO - " & Err & " " & Error & _
                Chr(13) & pMensagem)
End Function

Private Function mfSalvarProAgenc(pMensagem As String) As Boolean
Dim objPedidoProAg       As clsA46V733A
Dim lObjstcPedProAg      As stcA46V733B
Dim lI                  As Integer
Dim i                  As Integer
Dim lJ                  As Integer

    On Error GoTo ControleErro
    gM46V999.gBaseLocal = False
    Set objPedidoProAg = New clsA46V733A
    Set objPedidoProAg.M46V999 = gM46V999
    
    Set lObjstcPedProAg = New stcA46V733B
    If objstcPedido.ColecaoProlabore.Count > 0 Then
        lObjstcPedProAg.NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
        
        If Not objPedidoProAg.mfExcluir(lObjstcPedProAg, "", "", "") Then
            pMensagem = objPedidoProAg.UltimoErro
            GoTo Finalizar
        End If
    End If
    
    Set objstcPedido.ColecaoProlabore = New Collection
   
   
    If Not gM46V999.gfDisponivel("frmT47V028A") Then
        mfSalvarProAgenc = True
        GoTo Finalizar
    End If
   
    For lJ = 1 To 3
        gCod_RamoCorr = 0
        If gM46V999.gCod_Ramo <> e_CodRamo_Automoveis Then
            gCod_RamoCorr = objstcPedido.Cod_Ramo
        Else
            If lJ = 1 Then
                gCod_RamoCorr = e_CodRamoAuto
            ElseIf lJ = 2 And objstcPedido.objstcPedAuto.Val_Pr_Rc <> 0 Then
                gCod_RamoCorr = e_CodRamoRC
            ElseIf lJ = 3 And objstcPedido.objstcPedAuto.Val_Pr_App <> 0 Then
                gCod_RamoCorr = e_CodRamoapp
            End If
        End If
        If gCod_RamoCorr <> 0 Then
            'Grava Prolabore
            For i = 1 To frmT47V028A.flxProlabore.Rows - 1
                If Val(Mid(frmT47V028A.flxProlabore.TextMatrix(i, 0), 1, 7)) <> 0 Then
                    Set lObjstcPedProAg = New stcA46V733B
                    With lObjstcPedProAg
                    
                        .NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
                        .NUM_PI = gM46V111.gNum_Pi
                        .Num_Apol = objstcPedido.Num_Apol
                        .Num_Seq = lI
                        .Tip_Canal = 2    'Prolabore
                        .Cod_Corretor = objstcPedido.objstcCtrlEmis.Cod_Corr
                        .Cod_Canal = Mid(frmT47V028A.flxProlabore.TextMatrix(i, 0), 1, 7)
                        .Perc_Canal = frmT47V028A.flxProlabore.TextMatrix(i, 1)
                        .Qtd_Parc_Apol = objstcPedido.Qtd_Parc
                        .Qtd_Parc_Pag = objstcPedido.Qtd_Parc
                        
                    End With
                    lI = lI + 1
                    If Not objPedidoProAg.mfIncluir(lObjstcPedProAg, "", "", "") Then
                        pMensagem = objPedidoProAg.UltimoErro
                        GoTo Finalizar
                    End If
                    objstcPedido.ColecaoProlabore.Add lObjstcPedProAg, lObjstcPedProAg.Class_ID
                End If
            Next
            
            'Grava Agenciamento
            For i = 1 To frmT47V028A.flxAgenciamento.Rows - 1
                If Val(Mid(frmT47V028A.flxAgenciamento.TextMatrix(i, 0), 1, 7)) <> 0 Then
                    Set lObjstcPedProAg = New stcA46V733B
                    With lObjstcPedProAg
                        .NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
                        .NUM_PI = gM46V111.gNum_Pi
                        .Num_Apol = objstcPedido.Num_Apol
                        .Num_Seq = lI
                        .Tip_Canal = 2    'Prolabore
                        .Cod_Corretor = objstcPedido.objstcCtrlEmis.Cod_Corr
                        .Cod_Canal = Mid(frmT47V028A.flxAgenciamento.TextMatrix(i, 0), 1, 7)
                        .Perc_Canal = frmT47V028A.flxAgenciamento.TextMatrix(i, 1)
                        .Qtd_Parc_Apol = objstcPedido.Qtd_Parc
                        .Qtd_Parc_Pag = objstcPedido.Qtd_Parc
                        
                    End With
                    lI = lI + 1
                    If Not objPedidoProAg.mfIncluir(lObjstcPedProAg, "", "", "") Then
                        pMensagem = objPedidoProAg.UltimoErro
                        GoTo Finalizar
                    End If
                    objstcPedido.ColecaoProlabore.Add lObjstcPedProAg, lObjstcPedProAg.Class_ID
                End If
            Next
        End If
        If gM46V999.gCod_Ramo <> e_CodRamo_Automoveis Then
            Exit For
        End If
    Next lJ
            
    mfSalvarProAgenc = True
Finalizar:
    Set lObjstcPedProAg = Nothing
    Set objPedidoProAg = Nothing
    Exit Function
    
ControleErro:
    Set lObjstcPedProAg = Nothing
    Set objPedidoProAg = Nothing
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarProAgenc ERRO - " & Err & " " & Error & _
                Chr(13) & pMensagem)
End Function

Private Function mfSalvarBens(pMensagem As String, pobjstcPedLoc As stcA46V708B) As Boolean
Dim lobjPedBem       As clsA46V715A         'Objeto da classe clsA46V715A (TAB_PED_LOC_BEM).
Dim lobjstcPedBem    As stcA46V715B         'Objeto da classe de estrutura stcA46V715B (TAB_PED_LOC_BEM).
Dim lI               As Integer             'Utilizado no for...loop.
Dim lCheck           As String

    On Error GoTo ControleErro
    lCheck = 1
    Set lobjPedBem = New clsA46V715A
    Set lobjPedBem.M46V999 = gM46V999
    
    If pobjstcPedLoc.ColecaoPedLocBem.Count > 0 Then
        Set lobjstcPedBem = New stcA46V715B
        lobjstcPedBem.NOSSO_NUMERO = pobjstcPedLoc.NOSSO_NUMERO
        lobjstcPedBem.Num_Item = pobjstcPedLoc.Num_Item
        lCheck = 2
        If Not lobjPedBem.mfExcluir(lobjstcPedBem, "", "") Then
            pMensagem = lobjPedBem.UltimoErro
            GoTo Finalizar
        End If
    End If
    lCheck = 3
    Set pobjstcPedLoc.ColecaoPedLocBem = New Collection
    lCheck = 4
    For lI = 1 To frmT47V018A.flxBens.Rows - 1
        lCheck = 5
        If frmT47V018A.flxBens.ColWidth(1) = 0 And _
           gM46V999.gfPreenchido(frmT47V018A.flxBens.TextMatrix(lI, 2)) And Trim(frmT47V018A.flxBens.TextMatrix(lI, 2)) <> "0" Then
           GoTo Salvar
        End If
        If (gM46V999.gCod_Ramo = e_CodRamo_Empresarial Or _
           gM46V999.gCod_Ramo = e_CodRamo_Residencial Or _
           gM46V999.gCod_Ramo = e_CodRamo_Condominio) And _
           gM46V999.gfPreenchido(frmT47V018A.flxBens.TextMatrix(lI, 1)) And Trim(frmT47V018A.flxBens.TextMatrix(lI, 1)) <> "0" Then
           GoTo Salvar
        Else
            GoTo Proximo
        End If
           
Salvar:
        Set lobjstcPedBem = New stcA46V715B
        With lobjstcPedBem
            .NOSSO_NUMERO = pobjstcPedLoc.NOSSO_NUMERO
            .Num_Item = pobjstcPedLoc.Num_Item
            .Cod_Ramo = objstcPedido.Cod_Ramo
            .Num_Ped = objstcPedido.Num_Ped
            .Dsc_Bem = gM46V999.gfRetirarCaracterEspecial(frmT47V018A.flxBens.TextMatrix(lI, 2), True)
            lCheck = 6
            If IsNumeric(frmT47V018A.flxBens.TextMatrix(lI, 3)) = True Then
                .Val_Bem = gM46V999.CDblx(frmT47V018A.flxBens.TextMatrix(lI, 3))
            Else
                .Val_Bem = 0
            End If
            .Cod_Bem = gM46V999.CDblx(frmT47V018A.flxBens.TextMatrix(lI, 4))
            .Cod_Seq = gM46V999.CDblx(frmT47V018A.flxBens.TextMatrix(lI, 0))
            lCheck = 7
            If Not lobjPedBem.mfIncluir(lobjstcPedBem, "", "", "") Then
                pMensagem = lobjPedBem.UltimoErro
                GoTo Finalizar
            End If
            lCheck = 8
            pobjstcPedLoc.ColecaoPedLocBem.Add lobjstcPedBem
        End With
Proximo:
    Next lI
    mfSalvarBens = True
Finalizar:
    Set lobjPedBem = Nothing
    Set lobjstcPedBem = Nothing
    Exit Function
ControleErro:
    Set lobjPedBem = Nothing
    Set lobjstcPedBem = Nothing
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarBens ERRO - " & Err & " " & Error & _
                Chr(13) & "Check " & lCheck & pMensagem)
End Function

Private Function mfSalvarAuto(pMensagem As String) As Boolean
Dim objPedidoAuto   As clsA46V719A
    
    On Error GoTo ControleErro
    
    With objstcPedido.objstcPedAuto
        .NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
        If gM46V999.gfPreenchido(TxtValor2.Text) = False Then TxtValor2.Text = "0"
        If gM46V999.gfPreenchido(TxtValor3.Text) = False Then TxtValor3.Text = "0"
        If gM46V999.gfPreenchido(TxtValor4.Text) = False Then TxtValor4.Text = "0"
        If gM46V999.gfPreenchido(TxtValor5.Text) = False Then TxtValor5.Text = "0"
        If gM46V999.gfPreenchido(TxtValor6.Text) = False Then TxtValor6.Text = "0"
        .Val_Is_Auto = Val(txtISBas.Text)
        .Val_Is_Rc = TxtValor2.Text
        .Val_Is_App = TxtValor3.Text
        .Val_Pr_Auto = TxtValor4.Text
        .Val_Pr_Rc = TxtValor5.Text
        .Val_Pr_App = TxtValor6.Text
        .Cod_Ramo = objstcPedido.Cod_Ramo
        .Num_Ped = objstcPedido.Num_Ped
        .NUM_PI = gM46V111.gNum_Pi
    End With
    Set objPedidoAuto = New clsA46V719A
    Set objPedidoAuto.M46V999 = gM46V999
    
    Set objPedidoAuto.ConexaoBD = objPedido.ConexaoBD
    If Not objPedidoAuto.mfAlterar(objstcPedido.objstcPedAuto, "", "") Then
        pMensagem = objPedidoAuto.UltimoErro
        GoTo Finalizar
    End If
    mfSalvarAuto = True
Finalizar:
    Set objPedidoAuto = Nothing
    Exit Function
ControleErro:
    Set objPedidoAuto = Nothing
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarAuto ERRO - " & Err & " " & Err.Description & _
                Chr(13) & pMensagem)
End Function

Private Function mfSalvarEndereco(pMensagem As String) As Boolean
Dim objPedidoEnd   As clsA46V707A
    
    On Error GoTo ControleErro
    
    With objstcPedido.objstcPedEndereco
        .NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
        .Cod_Tip_Logr = frmT47V004A.txtTipLogradouro.Text
        .Nom_Logr = frmT47V004A.txtNomLogradouro.Text
        .Dsc_Compl_Logr = frmT47V004A.txtComplemento.Text
        .Nom_Bairro = frmT47V004A.txtBairro.Text
        .Nom_Cid = frmT47V004A.txtCidade.Text
        If frmT47V004A.cboUF.ListIndex <> -1 Then
            .Sig_UF = frmT47V004A.cboUF.List(frmT47V004A.cboUF.ListIndex)
        Else
            .Sig_UF = ""
        End If
        .Num_CEP = Val(Replace(frmT47V004A.txtCEP.Text, "-", ""))
        .Num_Fone = Val(Replace(Replace(Replace(frmT47V004A.txtTelefones.Text, "-", ""), "(", ""), ")", ""))
        .Cod_Ramo = objstcPedido.Cod_Ramo
        .Num_Ped = objstcPedido.Num_Ped
        .NUM_PI = gM46V111.gNum_Pi
    End With
    Set objPedidoEnd = New clsA46V707A
    Set objPedidoEnd.M46V999 = gM46V999
    Set objPedidoEnd.ConexaoBD = objPedido.ConexaoBD
    If mInclusao = True Then
        If Not objPedidoEnd.mfIncluir(objstcPedido.objstcPedEndereco, "", "", "") Then
            pMensagem = objPedidoEnd.UltimoErro
            GoTo Finalizar
        End If
    Else
        If Not objPedidoEnd.mfAlterar(objstcPedido.objstcPedEndereco, "", "") Then
            pMensagem = objPedidoEnd.UltimoErro
            GoTo Finalizar
        End If
    End If
    mfSalvarEndereco = True
Finalizar:
    Set objPedidoEnd = Nothing
    Exit Function
ControleErro:
    Set objPedidoEnd = Nothing
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarEndereco ERRO - " & Err & " " & Err.Description & _
                Chr(13) & pMensagem)
End Function

Private Sub mnuEdiBusProduto_Click()
Dim lProdAnt            As Integer
Dim lbdTAB_Cobert       As ADODB.Connection         'Arquivo: P00Tarifa.
Dim lMensagem           As String                   'Mensagem.
Dim lrsTab_Cobert       As clsYasRecordSet
Dim lIni                As Integer
Dim lFim                As Integer
Dim lobjstcLoc          As stcA46V708B
Dim lColecaoTMP         As Collection
Dim lClasfAnt           As String
Dim lValidade           As Long
Dim lHoje               As Long
Dim lSelect             As String
    Call gpLimPremios

    If gM46V111.gNum_Pi > 0 Then
        '1. Verifica in�cio de vig�ncia.
        If Not gM46V999.gfPreenchido(txtVigInicio.Text) Then
            Call gM46V999.gpGraLog(1, "FNP0011 - " & gM46V999.gPrefixo & _
                             "Vig�ncia inv�lida (data de in�cio de vig�ncia n�o preenchida).")
            Exit Sub
        End If
        If gM46V999.gfVerData(2, txtVigInicio.Text) = False Then
            Call gM46V999.gpGraLog(1, "FGE0018 - " & gM46V999.gPrefixo & _
                             "Vig�ncia inv�lida (data de in�cio de vig�ncia inv�lida).")
            Exit Sub
        End If
    
        lProdAnt = gM46V999.gCodProduto
        
        '3. Posiciona controles.
        gProChamado = "T47V021A"
    
        '4. Exibe form T47V021A.
        frmT47V021A.Show vbModal, Me
        
        'Verifico se est� tempo altera��o de tarifa.
        
        'Ficha 901354
        If gM46V999.gCodProduto & gM46V999.gCodSequencia < objstcPedido.Ult_Tarifa And (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
            If frmT47V016A.optTipEmissao(1).Value = True Then
                If gM46V999.gCodProduto >= 978 Then
                    lValidade = gfCalcData(objstcPedido.Dat_Pri_Calculo, 15)
                Else
                    lValidade = gfCalcData(objstcPedido.Dat_Pri_Calculo, 30)
                End If
            Else
                lValidade = gfCalcData(objstcPedido.Dat_Pri_Calculo, 15)
            End If
                    
            lHoje = Format(Date, "yyyymmdd")
            
            If lHoje > lValidade Then
                Call gM46V999.gpGraLog(1, "FGE0094 - Tarifa selecionada (" & gM46V999.gCodProduto & ") n�o pode ser utilizada para esta cota��o pois ultrapassou o prazo de validade " & lValidade)
                gM46V999.gCodProduto = lProdAnt
                Call mpObtProduto(gM46V999.gCodProduto, gM46V999.gCod_Ramo)
                lblNomProduto.Caption = gM46O001.gProNome
                lblNomProduto.Refresh
                Exit Sub
            End If
        End If
       
        If lProdAnt <> gM46V999.gCodProduto Then
            If MsgBox("AGE0019 - " & gM46V999.gPrefixo & "Altera��o de Tarifa. " & _
                        "Regi�o, classifica��o e lista de coberturas " & _
                        IIf(objstcPedido.ColecaoPedLoc.Count > 1, "de todos os itens ", "") & _
                        "ser�o desconsideradas. Confirma esta op��o?", vbQuestion + vbYesNo) <> vbYes Then
                gM46V999.gCodProduto = lProdAnt
                Call mpObtProduto(gM46V999.gCodProduto, gM46V999.gCod_Ramo)
                lblNomProduto.Caption = gM46O001.gProNome
                lblNomProduto.Refresh
                Exit Sub
            End If
            
            Call mpCarregarPlano
            If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Cobert, lMensagem) = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
                Exit Sub
            End If

            lClasfAnt = txtClassificacao.Tag
            lSelect = " Select top 1 * from TAB_COBERT"
            lSelect = lSelect & " Where cod_Ramo = " & gM46V999.gCod_Ramo
            lSelect = lSelect & " and COD_TARIF_INI <= " & lProdAnt
            lSelect = lSelect & " and COD_TARIF_TER >= " & lProdAnt
            If gM46V999.gfObtRegistro(lbdTAB_Cobert, lSelect, lrsTab_Cobert, lMensagem) = False Then
                Call gM46V999.gpGraLog(1, lMensagem)
                Exit Sub
            End If
            If Not lrsTab_Cobert.EOF Then
                lIni = lrsTab_Cobert("Cod_Tarif_Ini")
                lFim = lrsTab_Cobert("Cod_Tarif_Ter")
                lSelect = " Select top 1 * from TAB_COBERT"
                lSelect = lSelect & " Where cod_Ramo = " & gM46V999.gCod_Ramo
                lSelect = lSelect & " and COD_TARIF_INI <= " & gM46V999.gCodProduto
                lSelect = lSelect & " and COD_TARIF_TER >= " & gM46V999.gCodProduto
                If gM46V999.gfObtRegistro(lbdTAB_Cobert, lSelect, lrsTab_Cobert, lMensagem) = False Then
                    Call gM46V999.gpGraLog(1, lMensagem)
                    Exit Sub
                End If
                If Not lrsTab_Cobert.EOF Then
                    If lIni <> lrsTab_Cobert("Cod_Tarif_Ini") Or lFim <> lrsTab_Cobert("Cod_Tarif_Ter") Then
                          If objstcPedido.ColecaoPedLoc.Count > 1 Then
                            Set lColecaoTMP = objstcPedido.ColecaoPedLoc
                            Set objstcPedido.ColecaoPedLoc = New Collection
                            For Each lobjstcLoc In lColecaoTMP
                                Set lobjstcLoc.ColecaoPedCobert = New Collection
                                lobjstcLoc.Cod_Sit_item = e_SitItem_NaoCalc
                                lobjstcLoc.Cod_Iden_Tarif = gM46V999.gCodProduto
                                objstcPedido.ColecaoPedLoc.Add lobjstcLoc, lobjstcLoc.Class_ID
                            Next lobjstcLoc
                          End If
                     End If
                End If
            End If
        End If
        Call mpObtProduto(gM46V999.gCodProduto, gM46V999.gCod_Ramo)
    
        'Carregar coberturas
        Call mpCarregarCoberturas

        lblNomProduto.Caption = gM46O001.gProNome
        lblNomProduto.Refresh
    End If
End Sub

Private Sub mnuEdiExcItem_Click()
    Dim lMensagem   As String   'Mensagem.
    Dim lQtdIten    As Long
    Dim lUltItem    As String
    Dim lItemAtual  As String
    Dim lUltItemAtual   As String
    Dim lobjstcItem As stcA46V708B
    Dim lChave      As Long
    
    lUltItemAtual = Format(frmT47V001A.txtNumItem.Text, "00000")
    Screen.MousePointer = vbHourglass
    lQtdIten = objstcPedido.ColecaoPedLoc.Count
    lItemAtual = lblAntNumItem.Caption
    Select Case Left$(gM46V999.gPrefixo, 15)
        Case "Registro padr�o"
            If mfExcRegPadrao(lMensagem) = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
            End If
            Screen.MousePointer = vbDefault
            Call mnuArqSaiItem_Click
        Case Else
            If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
                If objstcPedido.Cod_Tip_Doc = e_TipDoc_Proposta Then
                    If objstcPedido.Dat_Pri_Calculo <> gM46V999.gfForData(1, Date) Then
                        If objstcPedido.ColecaoPedLoc.Count <> objstcPedLoc.Num_Item Or _
                           objstcPedido.ColecaoPedLoc.Count = 1 Then
                            Screen.MousePointer = vbDefault
                            Call gM46V999.gpGraLog(1, "Exclus�o v�lida somente no mesmo dia da digita��o do PI, ou para o �ltimo item!")
                            Exit Sub
                        End If
                    End If
                End If
            End If
            If mfExcItem(lMensagem) = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
            End If
            
            'Se for uma coletiva ent�o pegar o �ltimo item e realizar uma c�pia para o n�mero exclu�do e
            'depois excluir o �ltimo item
            If lQtdIten > 1 Then
                lUltItem = Format(frmT47V014A.flxItens.TextMatrix(frmT47V014A.flxItens.Rows - 1, 0), "00000")
                If lUltItemAtual <> lItemAtual Then
                    Call mfCopItem(gM46V111.gComNosNumero, lUltItem, lItemAtual, lMensagem)
                    Call mpCriGraItens
                    Set objstcPedLoc = objstcPedido.ColecaoPedLoc.Item(objstcPedido.ColecaoPedLoc.Count)
                    Set lobjstcItem = objstcPedLoc.gCopiaItem(CLng(lItemAtual))
                    lobjstcItem.objstcPedInsp.Num_Logr_Risco = objstcPedLoc.objstcPedInsp.Num_Logr_Risco
                    lobjstcItem.objstcPedInsp.Dat_Cadastro = objstcPedLoc.objstcPedInsp.Dat_Cadastro
                    lobjstcItem.objstcPedInsp.Num_Tel = objstcPedLoc.objstcPedInsp.Num_Tel
                    lobjstcItem.objstcPedInsp.Pess_Cont = objstcPedLoc.objstcPedInsp.Pess_Cont
                    
                    lChave = gpPosicaoColecao(lobjstcItem.Num_Item)
                    If objstcPedido.ColecaoPedLoc.Count > 0 And objstcPedido.ColecaoPedLoc.Count <> lChave - 1 Then
                        objstcPedido.ColecaoPedLoc.Add lobjstcItem, lobjstcItem.Class_ID, lChave
                    Else
                        objstcPedido.ColecaoPedLoc.Add lobjstcItem, lobjstcItem.Class_ID
                    End If
                    Call mfExcUltimoItem(lUltItem, lMensagem)
                    Call mpCriGraItens
                    Set objstcPedLoc = objstcPedido.ColecaoPedLoc.Item(1)
                    '1000897 - Renumerar itens apos exclusao em coletivas
                    If objstcPedido.ColecaoPedLoc.Count >= 1 Then
                        gCarga = True
                        Call mpPosItem(objstcPedLoc)
                        gCarga = False
                    End If
                    
                End If
            End If
                    
            Screen.MousePointer = vbDefault
        End Select
    End Sub
Private Function mfExcItem(ByRef pMensagem As String) As Boolean
    'Func�o: exclui item.

    Dim lEncontrou  As Boolean          'Indica que encontrou o registro.
    Dim lLinha      As Byte             'Linha selecionada.
    Dim lMensagem   As String           'Mensagem.

    mfExcItem = False

    '1. Confirma com usu�rio.
    lMensagem = "AGE0023 - " & gM46V999.gPrefixo & "Confirma exclus�o?" & vbLf & vbLf & "Item " & _
                lblAntNumItem.Caption & "."
    If MsgBox(lMensagem, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        Call gM46V999.gpGraLog(0, lMensagem & " Sim.")
    Else
        Call gM46V999.gpGraLog(0, lMensagem & " N�o.")
        mfExcItem = True
        Exit Function
    End If
Retorno:
    gM46V999.gBaseLocal = False
    Call gM46V999.gpExcluiRegistros(gM46V111.gComNosNumero, lblAntNumItem.Caption, "", "", pMensagem)
    If objstcPedido.ColecaoPedLoc.Count > 0 Then
        'Removo o item da cole��o de itens.
        objstcPedido.ColecaoPedLoc.Remove (CStr(gM46V111.gComNosNumero) & Format(CStr(lblAntNumItem.Caption), "00000"))
        
        '6. Atualiza grade de itens e exibe outro item.
        With frmT47V014A.flxItens
            If .Rows = 2 Then
                .Rows = 1
                .Row = 0
            Else
                lLinha = .Row
                If .Row > 0 Then .RemoveItem .Row
                If lLinha > (.Rows - 1) Then
                    lLinha = .Rows - 1
                End If
                .Row = lLinha
                If objstcPedido.ColecaoPedLoc.Count > 0 Then
                    lEncontrou = False
                    .Col = 0    'N�mero do item.
                    'busco o pr�ximo item
                    Set objstcPedLoc = objstcPedido.ColecaoPedLoc.Item(1)
                    gCarga = True
                    Call mpPosItem(objstcPedLoc)
                    gCarga = False
                    Call gpHabSalvar
                    mfExcItem = True
                    Exit Function
                End If
            End If
        End With
    End If
    
    '7. N�o tem outro item. Exclui registro de dados comuns e registros relacionados.
    If objPedido Is Nothing Then
        Set objPedido = New clsA46V702A
        Set objPedido.M46V999 = gM46V999
    End If
    If gM46V999.gfPreenchido(objstcPedido.NOSSO_NUMERO) Then
        If Not objPedido.mfExcluir(objstcPedido, "", "", "") Then
            pMensagem = objPedido.UltimoErro
            Exit Function
        End If
    End If
        
    Call gM46V999.gpGraLog(1, "AGE0030 - " & gM46V999.gPrefixo & "Exclus�o realizada com sucesso." & vbLf & vbLf & _
                     "Documento com nosso n�mero " & gM46V111.gComNosNumero & ".")
    Screen.MousePointer = vbDefault
    
    Call mpDesSalvar
    Call mnuArqSaiItem_Click
    

    mfExcItem = True
End Function

Private Sub mpPosItem(pobjstcPedLoc As stcA46V708B)
'********************Vari�veis padr�o*************************************
Dim lCkPoint        As String      'Variavel utilizada para determinar pontos onde o c�digo j� passou.
Dim lCkStr          As String       'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
                                    'lCkPoint. (posi��o das variaveis em ordem de declara��o)
'********************Vari�veis Declaradas para o processamento************
Dim lI              As Integer          'Utilizado no for...next.
Dim lCEP            As String           'CEP.
Dim lCount          As Integer          'Utilizado no for...next.
Dim lobjstcCobert   As stcA46V704B     'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT).
Dim lobjstcPedClau   As stcA46V703B     'Objeto da classe de estrutura stcA46V704B (TAB_PED_CLAU).

Dim lobjstcBem      As stcA46V715B      'Objeto da classe de estrutura stcA46V715B (TAB_PED_LOC_BEM).
Dim lobjstcSeguro   As stcA46V716B      'Objeto da classe de estrutura stcA46V716B (TAB_PED_LOC_SEGURO).
Dim lretorno        As String           'Retorno
Dim lobjstcLMI      As stcA46V728B      'Objeto da classe de estrutura stcA46V728B (TAB_PED_LOC_LMI).

    
    On Error GoTo ControleErro
    Err = 0
    lCkPoint = ""
    lCkStr = ""
    
    lCkPoint = "1"
    If pobjstcPedLoc Is Nothing Then
        Set pobjstcPedLoc = New stcA46V708B
        Set pobjstcPedLoc.M46V999 = gM46V999
    End If
    pobjstcPedLoc.Cod_Ramo = objstcPedido.Cod_Ramo
    '1. N�mero do item
    If pobjstcPedLoc.Num_Item = 0 Then pobjstcPedLoc.Num_Item = 1
    txtNumItem.Text = Format(pobjstcPedLoc.Num_Item, "00000")
    lblAntNumItem.Caption = txtNumItem.Text
    
    
    lCkPoint = "2"
    '3. Tipo de logradouro do Risco
    txtRisTipLogradouro.Text = pobjstcPedLoc.Cod_Tip_Logr_Risco
    
    '4. Nome do logradouro
    txtRisNomLogradouro.Text = pobjstcPedLoc.Nom_Logr_Risco
    
    '5. Complemento
    txtRisComplemento.Text = pobjstcPedLoc.Dsc_Compl_Risco
    
    '6. Bairro
    txtRisBairro.Text = pobjstcPedLoc.Nom_Bairro_Risco
    
    '7. Cidade
    txtRisCidade.Text = pobjstcPedLoc.Nom_Cid_Risco
    
    lCkPoint = "3"
    '8. UF Risco
    cboRisUF.ListIndex = -1
    If gM46V999.gfPreenchido(pobjstcPedLoc.Sig_UF_Risco) Then
        For lI = 0 To (cboRisUF.ListCount - 1)
            If pobjstcPedLoc.Sig_UF_Risco = cboRisUF.List(lI) Then
                cboRisUF.ListIndex = lI
                Exit For
            End If
        Next lI
    End If
    lCkPoint = "4"
    '9. CEP Risco
    If pobjstcPedLoc.Num_Cep_Risco > 0 Then
        lCEP = Format(pobjstcPedLoc.Num_Cep_Risco, "00000000")
        txtRisCEP.Text = Left(lCEP, 5) & "-" & Right(lCEP, 3)
        mAntCEP = txtRisCEP.Text
        If pobjstcPedLoc.Cod_Regiao = 0 And pobjstcPedLoc.Cod_Iden_Tarif <= gM46V999.TarifaAntiga Then
            gCarga = False
            txtRisCEP_LostFocus
            gCarga = True
        End If
    Else
        txtRisCEP.Text = ""
    End If
    
    mAntCEP = txtRisCEP.Text
    
    lCkPoint = "5"
    '10. N�mero do endere�o
    txtRisNumero.Text = pobjstcPedLoc.objstcPedInsp.Num_Logr_Risco
    
    '13. Valor de Risco
    txtValoRisco.Text = Format(pobjstcPedLoc.Val_Vr, "standard")
    
    '14. IS Valor de risco
    txtISBas.Text = Format(pobjstcPedLoc.Val_Vr * pobjstcPedLoc.Per_IS_Vr / 100, "standard")
    Call txtISBas_LostFocus
    
    lCkPoint = "6"
    '15. Classificacao
    txtClassificacao.Tag = 0
    If cboPlano(0).ListCount > 0 Then
        cboPlano(0).ListIndex = -1
        If pobjstcPedLoc.Tab_Clasf <> 0 Then
            For lI = 0 To (cboPlano(0).ListCount - 1)
                If pobjstcPedLoc.Tab_Clasf = cboPlano(0).ItemData(lI) Then
                    cboPlano(0).ListIndex = lI
                    Exit For
                End If
            Next lI
        End If
        
        gM46V111.gCod_plano = pobjstcPedLoc.Cod_Plano
    End If
    objstcPedido.Desc_Plano = cboPlano(0).Text
    
    If cboPlano(0).ListIndex <> -1 Then
        pobjstcPedLoc.Cod_Plano = cboPlano(0).ItemData(cboPlano(0).ListIndex)
        pobjstcPedLoc.Tab_Clasf = 0
        gM46V111.gCod_plano = pobjstcPedLoc.Cod_Plano
    End If
    
    lCkPoint = "7"
   
    '22. Cobertura
    'Carregar coberturas
    TabMod.Tabs.Clear
    TabMod.Tabs.Add
    TabMod.Tabs.Item(TabMod.Tabs.Count).Tag = 0
    For lI = 0 To flxGridCoberturaDemais.UBound
        If lI <> 0 Then
            Unload flxGridCoberturaDemais(lI)
        End If
    Next
    flxGridCoberturaDemais(0).Rows = 1
    flxGridCoberturaDemais(0).Rows = 2
    flxGridCoberturaDemais(0).FixedRows = 1
    Dim lI2 As Integer
    Dim lI3 As Integer
    Dim lUltModalidade  As Long
    lUltModalidade = 0
    lI = 0
    lI2 = -1
    
    
    For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
        If lUltModalidade <> lobjstcCobert.Tab_Cobert Then
            If lUltModalidade <> 0 Then
                TabMod.Tabs.Add
                Load flxGridCoberturaDemais(flxGridCoberturaDemais.UBound + 1)
                mpCarregarGridCoberturaMod (flxGridCoberturaDemais.UBound)
                TabMod.Tabs.Item(TabMod.Tabs.Count).Tag = flxGridCoberturaDemais.UBound
                flxGridCoberturaDemais(flxGridCoberturaDemais.UBound).Rows = 2
                flxGridCoberturaDemais(flxGridCoberturaDemais.UBound).Visible = False
            End If
            lI2 = lI2 + 1
            lI = 0
        End If
        If gM46V999.gfPreenchido(TabMod.Tabs(lI2 + 1).Caption) = False Then
            For lI3 = 0 To (cboPlano(1).ListCount - 1)
                If lobjstcCobert.Tab_Cobert = cboPlano(1).ItemData(lI3) Then
                    TabMod.Tabs(lI2 + 1).Caption = Trim(cboPlano(1).List(lI3))
                    If cboPlano(1).ListIndex = -1 Then cboPlano(1).ListIndex = lI3
                    Exit For
                End If
            Next lI3
        End If
        lI = lI + 1
        If flxGridCoberturaDemais(lI2).Rows <= lI Then
            flxGridCoberturaDemais(lI2).Rows = flxGridCoberturaDemais(lI2).Rows + 1
        End If
        Set lobjstcCobert.M46V999 = gM46V999
        Set lobjstcCobert.ObjPai = pobjstcPedLoc
        
'        If lUltModalidade = lobjstcCobert.Tab_Cobert Then
            flxGridCoberturaDemais(lI2).TextMatrix(lI, 0) = lobjstcCobert.Cod_Cobert
            flxGridCoberturaDemais(lI2).TextMatrix(lI, 1) = lobjstcCobert.gNomeCobertura
            flxGridCoberturaDemais(lI2).TextMatrix(lI, 2) = Format(lobjstcCobert.Val_IS, "standard")
            flxGridCoberturaDemais(lI2).TextMatrix(lI, 3) = Format(lobjstcCobert.Val_IS, "standard")
            flxGridCoberturaDemais(lI2).TextMatrix(lI, 4) = lobjstcCobert.Tab_Cobert
            flxGridCoberturaDemais(lI2).TextMatrix(lI, 5) = Format(lobjstcCobert.Per_Franq, "standard")
            flxGridCoberturaDemais(lI2).TextMatrix(lI, 6) = Format(lobjstcCobert.Val_Franq_Min, "standard")
            flxGridCoberturaDemais(lI2).TextMatrix(lI, 7) = Format(lobjstcCobert.Per_Desc_Coml, "standard")
            flxGridCoberturaDemais(lI2).TextMatrix(lI, 8) = Format(lobjstcCobert.Per_Desc_Exp, "standard")
            flxGridCoberturaDemais(lI2).TextMatrix(lI, 9) = Format(lobjstcCobert.VAL_PR, "standard")
            flxGridCoberturaDemais(lI2).Tag = lobjstcCobert.Tab_Cobert
'        End If
        lUltModalidade = lobjstcCobert.Tab_Cobert
    Next lobjstcCobert
    If pobjstcPedLoc.ColecaoPedCobert.Count > 0 Then
        flxGridCoberturaDemais(0).Visible = True
        txtModalidade.Visible = False
        cboPlano(1).Visible = False
    Else
        If pobjstcPedLoc.ColecaoPedCobert.Count = 0 Then
            If objstcPedido.Cod_Estip > 0 Then
                lUltModalidade = objstcPedido.Cod_Estip
                Load flxGridCoberturaDemais(flxGridCoberturaDemais.UBound + 1)
                mpCarregarGridCoberturaMod (flxGridCoberturaDemais.UBound)
                TabMod.Tabs.Item(TabMod.Tabs.Count).Tag = flxGridCoberturaDemais.UBound
                flxGridCoberturaDemais(flxGridCoberturaDemais.UBound).Rows = 2
                flxGridCoberturaDemais(flxGridCoberturaDemais.UBound).Visible = False
                lI2 = lI2 + 1
                lI = 0
                If gM46V999.gfPreenchido(TabMod.Tabs(lI2 + 1).Caption) = False Then
                    For lI3 = 0 To (cboPlano(1).ListCount - 1)
                        If lUltModalidade = cboPlano(1).ItemData(lI3) Then
                            TabMod.Tabs(lI2 + 1).Caption = Trim(cboPlano(1).List(lI3))
                            If cboPlano(1).ListIndex = -1 Then cboPlano(1).ListIndex = lI3
                            Exit For
                        End If
                    Next lI3
                End If
            End If
        End If
    End If
    
    
    lCkPoint = "8"
    '25. Bens
    If gM46V999.gfDisponivel("frmT47V018A") Then
        Unload frmT47V018A
    End If
    lI = 0
    optBens(1).Value = True
    For Each lobjstcBem In pobjstcPedLoc.ColecaoPedLocBem
        With frmT47V018A.flxBens
            lI = lI + 1
            If lI > 1 Then
                .Rows = .Rows + 1
            End If
            .TextMatrix(.Rows - 1, 0) = Format(.Rows - 1, "000")
            .TextMatrix(.Rows - 1, 1) = lobjstcBem.gTipBem
            .TextMatrix(.Rows - 1, 2) = lobjstcBem.Dsc_Bem
            .TextMatrix(.Rows - 1, 3) = Format(lobjstcBem.Val_Bem, "#,###,##0.00")
            .TextMatrix(.Rows - 1, 4) = lobjstcBem.Cod_Bem
            .RowHeight(.Rows - 1) = frmT47V018A.CmbBens.Height
        End With
        optBens(0).Value = True
    Next lobjstcBem
    
    lCkPoint = "12"
    '28. Tipo de emiss�o
    If gM46V999.gfDisponivel("frmT47V016A") Then
        Unload frmT47V016A
    End If
        
    With frmT47V016A
        If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_SeguroNovo Or lVerTipEmissao = "S" Then
            .optTipEmissao(0).Value = True
        ElseIf pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Or lVerTipEmissao = "R" Then
            .optTipEmissao(1).Value = True
            .txtRenApolice.Text = pobjstcPedLoc.Num_Apol_Anterior
            .txtRenItem.Text = pobjstcPedLoc.Num_Item_Anterior
            If pobjstcPedLoc.Ind_Sinistro = e_SimNao_Sim Then
                .optSinistroYas(0).Value = True
            Else
                .optSinistroYas(1).Value = True
            End If
        ElseIf pobjstcPedLoc.Tip_Emissao = e_TipEmissao_Endosso Or lVerTipEmissao = "E" Then
            .txtRenApolice.Text = pobjstcPedLoc.Num_Apol_Anterior
            .txtRenItem.Text = pobjstcPedLoc.Num_Item_Anterior
            .optTipEmissao(3).Value = True
        Else
            .optTipEmissao(0).Value = True  'Novo
        End If

        'sandra-aqui
        If gCod_Registro > 0 Then objstcPedido.Cod_Reg = gCod_Registro
        
        If objstcPedido.Cod_Reg = 0 Then
            If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_Endosso Or lVerTipEmissao = "E" Then
                objstcPedido.Cod_Reg = 101
                gCod_Registro = 101
            Else
                objstcPedido.Cod_Reg = 1
                gCod_Registro = 1
            End If
        End If
        For lI = 0 To .cboCodReg.ListCount - 1
            If Mid(.cboCodReg.List(lI), 1, 3) = objstcPedido.Cod_Reg Then
                .cboCodReg.ListIndex = lI
                Exit For
            End If
        Next
        If objstcPedido.objstcCtrlEmis.Cod_Envi_Doc = "C" Then
            .cboTipEnvio.ListIndex = 2
        '[FRH]Ficha:0801295
        ElseIf objstcPedido.objstcCtrlEmis.Cod_Envi_Doc = "X" Or objstcPedido.objstcCtrlEmis.Cod_Envi_Doc = "" Then
            .cboTipEnvio.ListIndex = 0
        ElseIf objstcPedido.objstcCtrlEmis.Cod_Envi_Doc = "O" Or _
               objstcPedido.objstcCtrlEmis.Cod_Envi_Doc = "E" Then
           objstcPedido.objstcCtrlEmis.Cod_Envi_Doc = "X"
           .cboTipEnvio.ListIndex = 0
        Else
            .cboTipEnvio.ListIndex = 1
        End If
    End With
    
    lCkPoint = "14"
    'caption
    Call mpPosCaption(objstcPedido.Cod_Tip_Doc, objstcPedido.Num_Ped, pobjstcPedLoc.Num_Item, objstcPedido.Cod_Sit_Ped)
    If txtSegurado.Enabled Then
        On Error Resume Next
        Call txtSegurado.SetFocus
    End If
    Exit Sub
ControleErro:
    'Obtem o inicio da mensagem de erro
    lCkStr = "mpPosItem ERRO - " & Err & " " & Error & _
                Chr(13) & "CheckPoint " & lCkPoint & " - "
    'Completa a mensagem de erro
    Call gM46V999.gpGraLog(0, lCkStr)
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
End Sub


Private Sub mpPosCaption(ByVal pTipDocumento As Byte, ByVal pNumReferencia As String, _
                         ByVal pNumItem As String, ByVal pSitDocumento As Byte)
    'Procedure: posiciona caption do form.

    'Par�metros de entrada...pTipDocumento....Tipo do documento.
    '                        pNumReferencia...N�mero de refer�ncia.
    '                        pNumItem.........N�mero do item.
    '                        pSitDocumento....Situa��o do documento.

    With Me
        Select Case pTipDocumento
            Case 1
                .Caption = "Cota��o "
            Case 2
                .Caption = "Proposta "
            Case Else
                .Caption = "Documento "
        End Select
        .Caption = .Caption & ": " & objstcPedido.NOSSO_NUMERO & " "

        .Caption = .Caption & " - Pedido: " & Format(pNumReferencia, "000000") & " - Item " & pNumItem
        
        
        If pSitDocumento = 3 Then
            .Caption = .Caption & " - PROPOSTA J� TRANSMITIDA"
        End If
    End With
End Sub
Private Sub mpPosMenu(ByVal pTipDocumento As Byte)
    'Procedure: posiciona menu, barra de ferramentas, vistoria e quadro de op��es.

    'Par�metro de entrada...pTipDocumento...Tipo de documento.

    '1. Posiciona controles
    Select Case pTipDocumento
    
        Case 1, 3  'Cota��o.
            'a) Barra de menu.
            'mnuEdiConCotProposta.Visible = True
            'mnuEdiConProCotacao.Visible = False
            'b) Barra de ferramentas.
            barFerramentas.Buttons(20).ToolTipText = "Converter cota��o em proposta."
            'c) Bot�es de comando no rodap�.
        Case 2, 4 'Proposta.
            'a) Barra de menu.
            mnuEdiBusProduto.Visible = False
            'mnuEdiConCotProposta.Visible = False
            'mnuEdiConProCotacao.Visible = True
            'b) Barra de ferramentas.
            'barFerramentas.Buttons(20).ToolTipText = "Converter proposta em cota��o."
            'c) Bot�es de comando no rodap�.
    End Select

    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        mnuExiPrazo.Visible = False
        cmdPrazo.Enabled = False
    End If


End Sub
Private Sub mpPosRamo()
Dim i As Integer
    cboRamo.ListIndex = -1
    For i = 0 To cboRamo.ListCount - 1
        If cboRamo.ItemData(i) = objstcPedido.Cod_Ramo Then
            cboRamo.ListIndex = i
            Exit For
        End If
    Next
End Sub

Private Function mfPosComum(ByRef pMensagem As String) As Boolean
    'Fun��o: monta dados comuns.
    'Par�metro de entrada...pTab_Ped...registro.
    'Par�metro de sa�da.....pMensagem...Mensagem.
'********************Vari�veis padr�o*************************************
    Dim lCkPoint        As String      'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lCkStr          As String       'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
                                        'lCkPoint. (posi��o das variaveis em ordem de declara��o)
    Dim clsA99V001          As clsA99V001
'********************Vari�veis Declaradas para o processamento************
    Dim lAuxiliar           As String           'Auxiliar.
    Dim lCEP                As String           'CEP.
    Dim lCodProdutor        As Integer          'C�digo do produtor.
    Dim lCodRegistro        As Long             'C�digo do Registro
    Dim lCodUniProdutiva    As Integer          'C�digo da unidade produtiva.
    Dim lComissao           As Double           'Comiss�o.
    Dim lI                  As Byte             'Utilizado com For...Next.
    Dim lNomProdutor        As String           'Nome do produtor.
    Dim lNomUniProdutiva    As String           'Nome da unidade produtiva.
    Dim lObjstcCorr         As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).
    Dim lObjstcVenc         As stcA46V713B      'Objeto da classe de estrutura stcA46V713B (TAB_PED_VENC).
    Dim lobjstcTextoEndo    As stcA46V722B      'Objeto da classe de estrutura stcA46V722B (TAB_PED_DIZER).
    Dim lobjstcTextoEspec   As stcA46V726B      'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).
    Dim lretorno            As Integer
    Dim lDigito             As String
    Dim lobjCtrlEmis    As clsA46V712A      'Objeto da classe clsA46V712A (TAB_CTRL_EMIS)
    Dim lobjstcCtrlEmis As stcA46V712B      'Objeto da classe de estrutura stcA46V712B (TAB_CTRL_EMIS)
    Dim lBlnSiscad          As Boolean
    Call mpControleBotoes(True)
    Err = 0
    lCkPoint = ""
    lCkStr = ""
    On Error GoTo ControleErro
    mfPosComum = False
    gCarga = True
    '1. Posiciona controles.
    lCkPoint = "1"
    'a) T�tulo, barra de menu, barra de ferramentas e bot�es.
    Call mpPosMenu(objstcPedido.Cod_Tip_Doc)
    
    'b) N�mero do pedido.
    gM46V111.gNumPedido = objstcPedido.Num_Ped
    
    gM46V999.gPrefixo = "Proposta " & Format$(gM46V111.gNumPedido, "000000") & " - "
    
    'c) C�digo do Dizer
    gM46V111.gCod_Dizer = objstcPedido.Cod_Dizer

    'd) C�digo de Registro
'    gCod_Registro = objstcPedido.Cod_Reg

    '2. T46V101A: item (dados comuns - panComum).
    'a) Ramo
    lCkPoint = "2"
    'b) Preencher Combo de Ramo com Tabela Tbm_Ramos
    'simony
    
    If objstcPedido.Cod_Ramo = e_CodRamo_NaoDefinido Then
        panLocRisco.Visible = False
        panItem.Visible = False
        panComum.Visible = False
        panValores.Visible = False
        cboRamo.ListIndex = -1
        panBotAltTipEmissao.Enabled = False
        panBotValidar.Enabled = False
        panBotDadProposta.Enabled = False
        mnuExibir.Enabled = False
        mnuEditar.Enabled = False
        barFerramentas.Visible = False
        mnuArqNovItem.Enabled = False
        mnuArqAbrItem.Enabled = False
        mnuArqExcProposta.Enabled = False
        mnuArqImprimir.Enabled = False
        mnuArqVisualizar.Enabled = False
        Call mpPosCaption(1, gM46V111.gNumPedido, "1", objstcPedido.Cod_Sit_Ped)
    Else
        'posiciona o ramo escolhido
        mpPosRamo
        cmdSair.Top = 1215
    End If
    objstcPedido.Cod_Ramo = gM46V999.gCod_Ramo
    'Simony
    gM46V111.gTipoGrupMulti = e_GrupoRamo_NDefinido
    
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Or _
       gM46V999.gCod_Ramo = e_CodRamo_Condominio Or _
       gM46V999.gCod_Ramo = e_CodRamo_Residencial Or _
       gM46V999.gCod_Ramo = e_CodRamo_Riscos_diversos Or _
       gM46V999.gCod_Ramo = e_CodRamo_GarantiaEstendida Or _
       gM46V999.gCod_Ramo = e_CodRamo_Roubos Then
       gM46V111.gTipoGrupMulti = e_GrupoRamo_Multi
    ElseIf gM46V999.gCod_Ramo = e_CodRamo_D_P_E_M Or _
           gM46V999.gCod_Ramo = e_CodRamo_Ap_bilhete Then
        gM46V111.gTipoGrupMulti = e_GrupoRamo_570_820
    ElseIf gM46V999.gCod_Ramo = e_CodRamo_Automoveis Then
        gM46V111.gTipoGrupMulti = e_GrupoRamo_Auto
    ElseIf gM46V999.gCod_Ramo = e_CodRamo_Acidentes_pessoais Then
        gM46V111.gTipoGrupMulti = e_GrupoRamo_Acidentes_pessoais
    ElseIf gM46V999.gCod_Ramo = e_CodRamo_Transp_internacional Or _
           gM46V999.gCod_Ramo = e_CodRamo_Transp_nacional Or _
           gM46V999.gCod_Ramo = e_CodRamo_Rc_Transportador Or _
           gM46V999.gCod_Ramo = e_CodRamo_Rcf_dc Then
        gM46V111.gTipoGrupMulti = e_GrupoRamo_Transp_Diversos
    ElseIf gM46V999.gCod_Ramo <> e_CodRamo_NaoDefinido Then
        gM46V111.gTipoGrupMulti = e_GrupoRamo_Generico
    End If
        
    If gM46V999.gCod_Ramo <> e_CodRamo_NaoDefinido Then
        panPesquisa.Enabled = False
        cboRamo.Enabled = False
        barFerramentas.Visible = True
        Call mpHabControles
    End If
    
    'Carrega ap�lice
    txtApolice.Text = objstcPedido.Num_Apol_An
    
    ''thais
    'panLocRisco.Visible = True
    panItem.Visible = False
    panLocRisco.Visible = True
    panLocRisco.Top = panItem.Top
    panItemLocRisco.Visible = True
    panItemLocRisco.Top = 3100
    'panItemLocRisco.Top = 3900
    
    If objstcPedido.Cod_Ramo <> 0 Then
        Call mpPosRamo
    End If
    ''thais
    lCkPoint = "3"
    
    'b) Nome do segurado.
    txtSegurado.Text = Trim(objstcPedido.Nom_Segurado)
    
    'c) Tipo de pessoa.
    If objstcPedido.Tip_Pessoa = e_TipPessoa_Fisica Then
        optTipPessoa(0).Value = True    'Pessoa f�sica
    Else
        optTipPessoa(1).Value = True    'Pessoa jur�dica
    End If
    lCkPoint = "4"
    'd. Posiciona a tarifa
    Call mpObtProduto(objstcPedido.Cod_Produto, objstcPedido.Cod_Ramo)
    If objstcPedido.Ult_Tarifa = 0 Then
        objstcPedido.Ult_Tarifa = objstcPedido.Cod_Produto & objstcPedido.Cod_Sequencia
    End If
    lblNomProduto.Caption = gM46O001.gProNome
    lblNomProduto.Refresh
    lCkPoint = "5"
    
    'e) Vig�ncia.
    If objstcPedido.Dat_Ini_Vig = 0 Then
        txtVigInicio.Text = ""
    Else
        txtVigInicio.Text = gM46V999.gfForData(2, objstcPedido.Dat_Ini_Vig)
    End If
    If objstcPedido.Dat_Ter_Vig = 0 Then
        txtVigTermino.Text = ""
    Else
        txtVigTermino.Text = gM46V999.gfForData(2, objstcPedido.Dat_Ter_Vig)
    End If
    
    If Not gfCarregaPerIOF(objstcPedido.Dat_Ini_Vig, pMensagem) Then Exit Function
        
    'f) Produto.
    lblNomProduto.Caption = gM46O001.gProNome
    
    'g) Tipo de documento.
    lblTipDocumento.Caption = e_TipDoc_Proposta

    'h) Comiss�o.
    'Retirado conf. 701722
    'lComissao = gProComBasica
    'txtComissao.Text = Format$(lComissao, "standard")
    lCkPoint = "6"
    
    'i) plano.
    '2. Posiciona plano
    
    
    
    '3. T46V102A: dados comuns.
    With frmT47V004A
        'a) Dados do segurado (panSegurado).
        '===1) Nome do segurado.
        .txtSegurado.Text = Trim(objstcPedido.Nom_Segurado)
        .txtCodSegurado.Text = objstcPedido.Cod_Segurado
        
        If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
            .txtDigCodSegurado.Text = objstcPedido.Dig_Cod_Segurado
        End If
        
        
        .txtCodSegurado.Enabled = True
        '===2) CNPJ/CPF.
        .panPesJuridica.Visible = False
        .panPesFisica.Visible = True
        If objstcPedido.Tip_Pessoa = 0 And gM46V999.gfPreenchido(objstcPedido.Num_Cgc_Cpf) And Trim(objstcPedido.Num_Cgc_Cpf) <> "0" Then
            If Len(CStr(objstcPedido.Num_Cgc_Cpf)) <= 11 Then
                objstcPedido.Tip_Pessoa = e_TipPessoa_Fisica
            Else
                objstcPedido.Tip_Pessoa = e_TipPessoa_Juridica
            End If
        End If
        Select Case objstcPedido.Tip_Pessoa
            Case e_TipPessoa_Juridica   'Pessoa jur�dica.
                .panPesJuridica.Visible = True
                .panPesFisica.Visible = False
            Case e_TipPessoa_Fisica   'Pessoa f�sica.
                .panPesJuridica.Visible = False
                .panPesFisica.Visible = True
        End Select
        lCkPoint = "7"
        Select Case objstcPedido.Tip_Pessoa
            Case e_TipPessoa_Juridica   'Pessoa jur�dica.
                .txtCNPJ.Text = Trim(objstcPedido.gCNPJ_CPF_Formatado)
                .cboRamAtividade.ListIndex = -1
                If gM46V999.gfPreenchido(objstcPedido.Cod_Ramo_Ativ) Then
                    For lI = 0 To (.cboRamAtividade.ListCount - 1)
                        If objstcPedido.Cod_Ramo_Ativ = .cboRamAtividade.ItemData(lI) Then
                            .cboRamAtividade.ListIndex = lI
                            Exit For
                        End If
                    Next lI
                End If
                .optTipPessoa(1).Value = True
            Case e_TipPessoa_Fisica   'Pessoa f�sica.
                .txtCPF.Text = Trim(objstcPedido.gCNPJ_CPF_Formatado)
                .cboIdeTipDocumento.ListIndex = -1
                If gM46V999.gfPreenchido(objstcPedido.Cod_Doc) Then
                    For lI = 0 To (.cboIdeTipDocumento.ListCount - 1)
                        If objstcPedido.Cod_Doc = .cboIdeTipDocumento.ItemData(lI) Then
                            .cboIdeTipDocumento.ListIndex = lI
                            Exit For
                        End If
                    Next lI
                End If
                .txtIdeNumDocumento.Text = objstcPedido.Num_Doc
                .txtIdeOrgExpedidor.Text = objstcPedido.Org_Exp
                .txtIdeDatEmissao.Text = gM46V999.gfForData(2, objstcPedido.Dat_Emis_Doc)
                .txtIdeEmpresa.Text = objstcPedido.Nom_Empresa
                If objstcPedido.Sexo = "M" Then
                    .optSexo(0).Value = True
                Else
                    .optSexo(1).Value = True
                End If
                .optTipPessoa(0).Value = True
        End Select
        lCkPoint = "8"
        'b) Endere�o do segurado (panEndereco).
        If objstcPedido.objstcPedEndereco Is Nothing Then
            Set objstcPedido.objstcPedEndereco = New stcA46V707B
            Set objstcPedido.objstcPedEndereco.M46V999 = gM46V999
        End If
        '===1) Tipo do logradouro.
        .txtTipLogradouro.Text = Trim(objstcPedido.objstcPedEndereco.Cod_Tip_Logr)
        '===2) Nome do logradouro.
        .txtNomLogradouro.Text = objstcPedido.objstcPedEndereco.Nom_Logr
        '===3) Complemento.
        .txtComplemento.Text = objstcPedido.objstcPedEndereco.Dsc_Compl_Logr
        '===4) Bairro.
        .txtBairro.Text = objstcPedido.objstcPedEndereco.Nom_Bairro
        '===5) Cidade.
        .txtCidade.Text = objstcPedido.objstcPedEndereco.Nom_Cid
        '===6) Unidade da federa��o.
        .cboUF.ListIndex = -1
        If gM46V999.gfPreenchido(objstcPedido.objstcPedEndereco.Sig_UF) Then
            For lI = 0 To (.cboUF.ListCount - 1)
                If objstcPedido.objstcPedEndereco.Sig_UF = .cboUF.List(lI) Then
                    .cboUF.ListIndex = lI
                    Exit For
                End If
            Next lI
        End If
        '===7) CEP.
        If objstcPedido.objstcPedEndereco.Num_CEP = 0 Then
            .txtCEP.Text = ""
        Else
            lCEP = Format(objstcPedido.objstcPedEndereco.Num_CEP, "00000000")
            .txtCEP.Text = Left(lCEP, 5) & "-" & Right(lCEP, 3)
        End If
        lCkPoint = "9"
        '===8) Telefones.
        .txtTelefones.Text = objstcPedido.objstcPedEndereco.Num_Fone
       
        'Chama o siscad
        Call frmT47V004A.mpAtualizaDadosSegurado(lBlnSiscad, False)
        
        '===9) Endere�o de cobran�a.
        If objstcPedido.gEnd_Cob_Diferente = False Then
            .optEndCobranca(1) = True    'Nao
        Else
            .optEndCobranca(0) = True    'Sim.
        End If
        
        'Posiciona Unidade e produtor
        Call mpPosiUnidProd(pMensagem)
        
        
        lCkPoint = "12"
        '===3) Cosseguros.
        gCarga = True
        .optCosseguro(1).Value = True   'N�o.
        If objstcPedido.ColecaoPedCoss.Count > 0 Then
            .optCosseguro(0).Value = True   'Sim.
        Else
            .optCosseguro(1).Value = True   'N�o.
        End If
        'd) Corretores (panCorretores).
        Call mpPosCorretor(objstcPedido.ColecaoPedCorr)
        
        lCkPoint = "13"
        'e) Produto (panProduto).
        '===1) Quantidade de parcelas.
        If gM46V999.gProQuaMaxParcelas = 0 Then
            gM46V999.gProQuaMaxParcelas = gM46V999.gProQuaMaxParcelasTAB
        End If
        If objstcPedido.Qtd_Parc = 0 And _
           (objstcPedido.Cod_Reg <> 106 And objstcPedido.Cod_Reg <> 107 And objstcPedido.Cod_Reg <> 120 And _
            objstcPedido.Cod_Reg <> 121 And objstcPedido.Cod_Reg <> 122 And objstcPedido.Cod_Reg <> 20) Then
            '>>>>>fatiitus
            If gM46V999.gTipoAcesso = DIGITACAO Then
                .txtQuaParcelas.Text = ""
            Else
                .txtQuaParcelas.Text = gM46V999.gProQuaMaxParcelas
            End If
            '>>>>>fatiitus
        Else
            If objstcPedido.Qtd_Parc > gM46V999.gProQuaMaxParcelas Then
                .txtQuaParcelas.Text = gM46V999.gProQuaMaxParcelas
            Else
                .txtQuaParcelas.Text = objstcPedido.Qtd_Parc
            End If
        End If
        
        If gM46V999.gfPreenchido(.txtQuaParcelas.Text) Then
            If IsNumeric(.txtQuaParcelas.Text) Then
                .optForPagamento(0).Enabled = True
                .optForPagamento(1).Enabled = True
            End If
        End If
        lCkPoint = "14"
        mCarregarDatas = True
        If objstcPedido.ColecaoPedVenc.Count > 0 Then
            frmT47V004A.optVencimento(0).Value = True
            frmT47V004A.optVencimento(1).Value = False
        Else
            mCarregarDatas = False
            frmT47V004A.optVencimento(0).Value = False
            frmT47V004A.optVencimento(1).Value = True
        End If

        frmT47V004A.txtJuros.Text = Format(objstcPedido.Per_Juro_Frac, "0.00")
            
        
        If gM46V999.PerfilAcesso = e_Perfil_CALC Then
            frmT47V004A.txtValorPgAnt.Text = "0"
        Else
            frmT47V004A.txtValorPgAnt.Text = Format(objstcPedido.Val_Pr_Antec, "0.00")
        End If
        
        If objstcPedido.Cod_Dizer <> 0 And objstcPedido.objstcCtrlEmis.Tip_Cotac = "3" Then
            objstcPedido.objstcCtrlEmis.Tip_Cotac = "1"
        End If
        
        For lI = 0 To (frmT47V004A.CmbTipCotac.ListCount - 1)
            If gM46V999.CDblx(objstcPedido.objstcCtrlEmis.Tip_Cotac) = frmT47V004A.CmbTipCotac.ItemData(lI) Then
                frmT47V004A.CmbTipCotac.ListIndex = lI
                Exit For
            End If
        Next lI
        
        '.Num_pi
        gM46V111.gNum_Pi = objstcPedido.objstcCtrlEmis.NUM_PI
        
        If gM46V111.gNum_Pi = 0 Then
            If gM46V999.gfPreenchido(txtNumPI.Text) Then
                If IsNumeric(txtNumPI.Text) Then
                    gM46V111.gNum_Pi = txtNumPI.Text
                End If
            Else
                gM46V111.gNum_Pi = objstcPedido.NUM_PI
            End If
            objstcPedido.objstcCtrlEmis.NUM_PI = gM46V111.gNum_Pi
        End If
        
        frmT47V004A.txtNumSiscota.Text = Format(gM46V999.CDblx(objstcPedido.objstcCtrlEmis.Num_Cotac), "000000000000")
        frmT47V004A.txtReferencia = objstcPedido.objstcCtrlEmis.Num_Ref_Corr
        
        If objstcPedido.Dat_Emis <> 0 Then
            frmT47V004A.txtDatEmis.Text = gM46V999.gfForData(2, objstcPedido.Dat_Emis)
        End If
        
        lCkPoint = "15"
        'Pagamento antecipado
        '***BOLETO***
'Ficha 900046
        .cmdPagAntDebito.Visible = False
        
        If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
            .cmdPagAntBloqueto.Visible = False
        End If
        
        If objstcPedido.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
            '***BOLETO***
            If (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) And (Left(objstcPedido.objstcCtrlEmis.Num_Bloq, 1) = 5 Or objstcPedido.objstcCtrlEmis.Num_Bloq = 0) Then
                .cmdPagAntBloqueto.Visible = False
            Else
                .cmdPagAntBloqueto.Visible = True
            End If
            If objstcPedido.objstcCtrlEmis.Num_Bloq <> 0 Then
                .optPagAntecipado(1).Value = True       'Sim - Antecipado com bloqueto.
                .optForPagamento(0).Value = True 'carne
                If (objstcPedido.objstcCtrlEmis.Num_Bloq <> 0 And Left(objstcPedido.objstcCtrlEmis.Num_Bloq, 1) <> 5) Or gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
                    frmT47V011A.txtPagAntBloqueto.Text = objstcPedido.objstcCtrlEmis.Num_Bloq & "-" & objstcPedido.objstcCtrlEmis.DIG_BLOQUETO
                End If
            ElseIf (objstcPedido.objstcCtrlEmis.Num_Bloq = 0 And objstcPedido.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 And objstcPedido.Cod_Agen_Bas_Cor <> 0) Then
                .optPagAntecipado(1).Value = True       'Sim - Antecipado com D�bito em conta corrente
                .optForPagamento(1).Value = True 'debito
                frmT47V010A.cboDebBanco.ListIndex = -1
                If objstcPedido.Cod_Bco_Cobr = 347 Then
                    objstcPedido.Cod_Bco_Cobr = 356
                End If
                For lI = 0 To (frmT47V010A.cboDebBanco.ListCount - 1)
                    If objstcPedido.Cod_Bco_Cobr = frmT47V010A.cboDebBanco.ItemData(lI) Then
                        frmT47V010A.cboDebBanco.ListIndex = lI
                        Exit For
                    End If
                Next lI
                frmT47V010A.txtDebAgencia.Text = objstcPedido.Cod_Agen_Bas_Cor
                frmT47V010A.txtDebConta.Text = objstcPedido.Num_Cta_Bas_Cor
            Else
                .optPagAntecipado(1).Value = True       'Sim - Bloqueto.
            End If
        Else
            .optPagAntecipado(2).Value = True
        End If
        lCkPoint = "16"
        lAuxiliar = ""
        'Forma de pagamento.
        If objstcPedido.Cod_Agen_Bas_Cor <> 0 Then      'D�bito
            .optForPagamento(1) = True
            frmT47V010A.cboDebBanco.ListIndex = -1
            For lI = 0 To (frmT47V010A.cboDebBanco.ListCount - 1)
                If objstcPedido.Cod_Bco_Cobr = frmT47V010A.cboDebBanco.ItemData(lI) Then
                    frmT47V010A.cboDebBanco.ListIndex = lI
                    Exit For
                End If
            Next lI
            frmT47V010A.txtDebAgencia.Text = objstcPedido.gAgenciaDig
            frmT47V010A.txtDebConta.Text = objstcPedido.Num_Cta_Bas_Cor
            If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
                lDigito = 0
                Set clsA99V001 = New clsA99V001
                Set clsA99V001.M46V999 = gM46V999
                
                Select Case objstcPedido.Cod_Bco_Cobr
                    Case 1  'Projeto 11080 - Debito automatico - Banco do Brasil
                        Call clsA99V001.DIG_BANCOBRASIL(Format(objstcPedido.Num_Cta_Bas_Cor, "0000000"), lDigito, lAuxiliar)
                    Case 237
                        Call clsA99V001.DIG_BRADESCO(Format(objstcPedido.Num_Cta_Bas_Cor, "0000000"), lDigito, lAuxiliar)
                    Case 341
                        Call clsA99V001.DIG_ITAU(Format(objstcPedido.gAgenciaDig, "0000"), Format(objstcPedido.Num_Cta_Bas_Cor, "00000"), lDigito, lAuxiliar)
                    Case 347
                        If objstcPedido.gAgenciaDig >= 1000 Then
                            Call clsA99V001.DIG_SUDAMERIS_NOVO(Format(objstcPedido.gAgenciaDig, "0000"), Format(objstcPedido.Num_Cta_Bas_Cor, "0000000"), lDigito, lAuxiliar)
                        Else
                            Call clsA99V001.DIG_SUDAMERIS(Format(objstcPedido.gAgenciaDig, "000"), Format(objstcPedido.Num_Cta_Bas_Cor, "000000000"), lDigito, lAuxiliar)
                        End If
                    Case 409
                         Call clsA99V001.DIG_UNIBANCO(Format(objstcPedido.gAgenciaDig, "0000"), Format(objstcPedido.Num_Cta_Bas_Cor, "000000"), lDigito, lAuxiliar)
                    Case 479
                        Call clsA99V001.DIG_BOSTON(Format(objstcPedido.Num_Cta_Bas_Cor, "0000000"), lDigito, lAuxiliar)
                    Case 356
                        Call clsA99V001.DIG_REAL(Format(objstcPedido.gAgenciaDig, "0000"), Format(objstcPedido.Num_Cta_Bas_Cor, "0000000"), lDigito, lAuxiliar)
                End Select
                frmT47V010A.txtDebConta.Text = frmT47V010A.txtDebConta.Text & "-" & lDigito
                Set clsA99V001 = Nothing
            End If
        Else        'Carn�
            .optForPagamento(0) = True
            If objstcPedido.Cod_Bco_Cobr = 0 And (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then objstcPedido.Cod_Bco_Cobr = 409
            For lI = 0 To (frmT47V009A.cboCarBanco.ListCount - 1)
                If objstcPedido.Cod_Bco_Cobr = frmT47V009A.cboCarBanco.ItemData(lI) Then
                    frmT47V009A.cboCarBanco.ListIndex = lI
                    Exit For
                End If
            Next lI
        End If
    End With
    lCkPoint = "17"
    '4. T46V105A: endere�o de cobran�a.
    With frmT47V006A
        If objstcPedido.gEnd_Cob_Diferente = False Then
            'a) Tipo do logradouro.
            .txtTipLogradouro.Text = Trim(objstcPedido.objstcPedEndereco.Cod_Tip_Logr)
            'b) Nome do logradouro.
            .txtNomLogradouro.Text = objstcPedido.objstcPedEndereco.Nom_Logr
            'c) Complemento.
            .txtComplemento.Text = objstcPedido.objstcPedEndereco.Dsc_Compl_Logr
            'd) Bairro.
            .txtBairro.Text = objstcPedido.objstcPedEndereco.Nom_Bairro
            'e) Cidade.
            .txtCidade.Text = objstcPedido.Nom_Cid
            'f) Unidade da federa��o.
            .cboUF.ListIndex = -1
            For lI = 0 To (.cboUF.ListCount - 1)
                If objstcPedido.objstcPedEndereco.Sig_UF = .cboUF.List(lI) Then
                    .cboUF.ListIndex = lI
                    Exit For
                End If
            Next lI
            'g) CEP.
            If objstcPedido.objstcPedEndereco.Num_CEP = 0 Then
                .txtCEP.Text = ""
            Else
                lCEP = Format(objstcPedido.objstcPedEndereco.Num_CEP, "00000000")
                .txtCEP.Text = Left(lCEP, 5) & "-" & Right(lCEP, 3)
            End If
            'h) Telefones.
            .txtTelefones.Text = objstcPedido.objstcPedEndereco.Num_Fone
        Else
            'a) Tipo do logradouro.
            .txtTipLogradouro.Text = Trim(objstcPedido.Cod_Tip_Logr)
            'b) Nome do logradouro.
            .txtNomLogradouro.Text = objstcPedido.Nom_Logr
            'c) Complemento.
            .txtComplemento.Text = objstcPedido.Dsc_Compl_Logr
            'd) Bairro.
            .txtBairro.Text = objstcPedido.Nom_Bairro
            'e) Cidade.
            .txtCidade.Text = objstcPedido.Nom_Cid
            'f) Unidade da federa��o.
            .cboUF.ListIndex = -1
            For lI = 0 To (.cboUF.ListCount - 1)
                If objstcPedido.Sig_UF = .cboUF.List(lI) Then
                    .cboUF.ListIndex = lI
                    Exit For
                End If
            Next lI
            'g) CEP.
            If objstcPedido.Num_CEP = 0 Then
                .txtCEP.Text = ""
            Else
                lCEP = Format(objstcPedido.Num_CEP, "00000000")
                .txtCEP.Text = Left(lCEP, 5) & "-" & Right(lCEP, 3)
            End If
            'h) Telefones.
            .txtTelefones.Text = objstcPedido.Num_Fone
        End If
    End With
    lCkPoint = "18"
    '5. T46V106A: prazo.
    Call txtVigTermino_LostFocus        'Posiciona quadro de prazo.
    With frmT47V003A
        .optTipPrazo(1).Value = True
        Select Case objstcPedido.Cod_Iden_Praz
            Case e_TipoPrazo_ProRata    'Pro-rata.
                .optTipPrazo(1).Value = True
            Case e_TipoPrazo_PrazoCurto 'Prazo curto.
                .optTipPrazo(0).Value = True
            Case e_TipoPrazo_PrazoLongo 'Prazo longo.
                .optTipPrazo(2).Value = True
            Case Else
                .optTipPrazo(1).Value = True
        End Select
    End With
    lCkPoint = "19"
    '6. T46V116A: cosseguros (se houver algum cosseguro).
    Call mpPosCosseguro(objstcPedido.ColecaoPedCoss)
    lCkPoint = "20"
    '7. T46V120A: custo de emiss�o.
    With frmT47V015A
        'a) Custo de emissao.
        .txtCusEmissao.Text = Format$(objstcPedido.Val_Cust_Emis, "standard")
        If lVerTipEmissao = "E" Then
            gM46V999.gProCusEndosso = objstcPedido.Val_Cust_Emis
        Else
            gM46V999.gProCusApolice = objstcPedido.Val_Cust_Emis
        End If
        'b) Custo de emiss�o parcelado.
        If objstcPedido.NUM_PI = 0 Then
                .optCusEmissao(0).Value = True     'Sim.
        Else
            If objstcPedido.gCustoParcelado = True Then
                .optCusEmissao(0).Value = True     'Sim.
            Else
                .optCusEmissao(1).Value = True     'N�o.
            End If
        End If
    End With
      
    lCkPoint = "21"
    '8. Vencimento das parcelas
    frmT47V004A.optVencimento(1).Value = True
    For Each lObjstcVenc In objstcPedido.ColecaoPedVenc
        frmT47v019A.txtDtVenc(lObjstcVenc.Num_Parc).Text = gM46V999.gfForData(2, lObjstcVenc.DAT_VENC)
        frmT47V004A.optVencimento(0).Value = True
        gM46V999.gProDatUltParcela = lObjstcVenc.DAT_VENC
    Next lObjstcVenc
    
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Or gM46V999.PerfilAcesso = e_Perfil_DIG_IST Then
'Ficha 900046
        If frmT47V004A.optPagAntecipado(2).Value = False And _
                frmT47V004A.optVencimento(0).Value = True Then
            frmT47v019A.txtDtVenc(1).Text = Format(Date + 7, "dd/mm/yyyy")
        End If
    End If
  
    '10. Controles diversos (vari�veis) - estes comandos precisam estar neste local (por �ltimo).
    mComUltData = objstcPedido.Dat_Liber
    mComUltUsuario = objstcPedido.Cod_User_Liber
    
    'simony deixar o que vier do arquivo
    'If gM46V999.CDblx(txtComissao.Text) = 0 Then
    '    txtComissao.Text = Format(gM46O001.gProComBasica, "standard")
    '    frmT47V004A.txtCorComissao(1).Text = Format(gM46O001.gProComBasica, "standard")
    'End If
    
    'simony
    '11. Valores de Autom�vel
    If objstcPedido.Cod_Ramo = e_CodRamo_Automoveis Then
        txtISBas.Text = Format(objstcPedido.objstcPedAuto.Val_Is_Auto, "standard")
        TxtValor2.Text = Format(objstcPedido.objstcPedAuto.Val_Is_Rc, "standard")
        TxtValor3.Text = Format(objstcPedido.objstcPedAuto.Val_Is_App, "standard")
        TxtValor4.Text = Format(objstcPedido.objstcPedAuto.Val_Pr_Auto, "standard")
        TxtValor5.Text = Format(objstcPedido.objstcPedAuto.Val_Pr_Rc, "standard")
        TxtValor6.Text = Format(objstcPedido.objstcPedAuto.Val_Pr_App, "standard")
    Else
        txtISBas.Text = Format(objstcPedido.Val_IS, "standard")
        TxtValor2.Text = Format(objstcPedido.Val_Pr_Ref_Resg, "standard")
        TxtValor3.Text = Format(objstcPedido.Val_Pr_Ref_Tarif, "standard")
        TxtValor4.Text = Format(objstcPedido.VAL_PR, "standard")
        TxtValor5.Text = Format(objstcPedido.Val_Cambio_Emis, "standard")
        TxtValor6.Text = 0
        TxtValor7.Text = Format(objstcPedido.Dat_Ini_Viag)
    End If
    lCkPoint = "22"
    Dim lNumPag As Integer
    '13. T47V024: Texto para Endosso.
    If objstcPedido.ColecaoTextoEndo.Count > 0 Then
        With frmT47V024A
            .txtTextoEndosso = ""
            For Each lobjstcTextoEndo In objstcPedido.ColecaoTextoEndo
                 .txtTextoEndosso = .txtTextoEndosso & lobjstcTextoEndo.Dsc_Dizer_Endo & vbCrLf
            Next lobjstcTextoEndo
            .txtTextoEndosso = Mid(.txtTextoEndosso, 1, Len(.txtTextoEndosso) - 2)
        End With
    End If
    
    lI = 1
    
    '14. T46V155: Texto para Especifica��o.
    lCkPoint = "23"
    If gM46V999.gCodProduto > gM46V999.TarifaAntiga Then
        If objstcPedido.ColecaoTextoEspec.Count > 0 Then
            For Each lobjstcTextoEspec In objstcPedido.ColecaoTextoEspec
                With frmT47V022A
                  .txtTextoEspec = ""
                  .txtTextoEspec = lobjstcTextoEspec.Espec_lin1 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin2 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin3 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin4 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin5 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin6 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin7 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin8 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin9 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin10 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin11 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin12 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin13 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin14 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin15 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin16 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin17 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin18 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin19 & vbCrLf
                  .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin20
                  .flxEspec.Rows = lI + 1
                  .flxEspec.TextMatrix(lI, 0) = lI
                  .flxEspec.TextMatrix(lI, 1) = lobjstcTextoEspec.Cod_Espec
                  .flxEspec.TextMatrix(lI, 2) = .txtTextoEspec
                  lI = lI + 1
                End With
            Next lobjstcTextoEspec
        End If
    End If
    
    lCkPoint = "24"
    'S� para teste, correto false
    gM46V999.gPendenteWorkflow = False
    gM46V999.gDesabilitarForms = False
    gM46V999.gAbrirWorkflow = False
    'S� para teste
    'Desabilita a tela pois TAB_Ctrl_Emis est� pendente
    If objstcPedido.objstcCtrlEmis.Cod_Iden_Stat_Acao <> 0 Then
        gM46V999.gPendenteWorkflow = True
        gM46V999.gDesabilitarForms = True
        gM46V999.gAbrirWorkflow = True
    End If
    
    lCkPoint = "25"
    
    frmT47V004A.TxtNumCobr.Text = CDblx(objstcPedido.objstcCtrlEmis.Num_Logr_Cobranca)
    frmT47V004A.TxtNumResid.Text = CDblx(objstcPedido.objstcCtrlEmis.Num_Logr_Segurado)
    frmT47V004A.TxtComplCobr.Text = objstcPedido.objstcCtrlEmis.Dsc_Compl_Cobranca
    frmT47V004A.TxtComplResid.Text = objstcPedido.objstcCtrlEmis.Dsc_Compl_Segurado
    
    lCkPoint = "26"
    mfPosComum = True
    lAuxiliar = ""
    
    If objstcPedido.Num_Apol <> 0 Then
        lAuxiliar = lAuxiliar & " emitido - Ap�lice: " & Format(objstcPedido.Num_Apol, "0000000000")
        lAuxiliar = lAuxiliar & " Endosso: " & objstcPedido.Num_Endo
    End If
    If lAuxiliar = "" Then
        Select Case objstcPedido.Cod_Posi_Atual
            Case "T"
                If objstcPedido.Cod_Sit_Ped = e_SitPedido_ComCalculo Then
                    lAuxiliar = " - Consistido "
                Else
                    lAuxiliar = " - N�o consistido "
                End If
            Case "9"
                lAuxiliar = " - Emiss�o Pendente"
            Case e_IdeEmitida_Emitir
                lAuxiliar = " - Transmitido"
            Case e_IdeEmitida_Suspender
                lAuxiliar = " - Suspenso"
        End Select
    End If
    lAuxiliar = Chr(183) & Space(10) & "PI: " & objstcPedido.objstcCtrlEmis.NUM_PI & lAuxiliar
    If objstcPedido.Ind_CotacaoSolic = False Or gM46V999.gTipoAcesso = DIGITACAO Then
        'mnuSituacao.Caption = lAuxiliar
    End If
    'MsgBox lAuxiliar, vbInformation
        
    gM46V111.gComNosNumero = objstcPedido.NOSSO_NUMERO
    
    lCkPoint = "31"
    Call mpPosPro_Agenc(objstcPedido.ColecaoProlabore, "P")
    Call mpPosPro_Agenc(objstcPedido.ColecaoAgenciador, "A")
    
    Exit Function
ControleErro:
    pMensagem = "mfPosComum ERRO - " & Err & " " & Error & _
                Chr(13) & "CheckPoint " & lCkPoint & " - "
    
   
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    'Obtem o inicio da mensagem de erro
    pMensagem = "mfPosComum ERRO - " & Err & " " & Error & _
                Chr(13) & "CheckPoint " & lCkPoint & " - "
    
   
    'Completa a mensagem de erro
    pMensagem = pMensagem & lCkStr
    
    Call gM46V999.gpGraLog(1, pMensagem)
    
    mfPosComum = False
    Exit Function
    Resume Next
End Function
Private Sub mpPosCorretor(pColecaoCorretor As Collection)
    'Procedure: posiciona corretor.

    'Par�metros...pLinha.........linha do corretor.
    '             pCodCorretor...c�digo do corretor.
    '             pPerCorretor...percentual do corretor.

    Dim lMensagem   As String       'Mensagem.
    Dim lSaida      As Variant      'Resultado da fun��o de busca de dados do corretor.
    Dim lObjstcCorr As stcA46V705B  'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).
    Dim lTotal      As Double       'Total de comiss�o.
    
    With frmT47V004A
        For Each lObjstcCorr In pColecaoCorretor
            Set lObjstcCorr.M46V999 = gM46V999
            If lObjstcCorr.Num_Ordem_Corr = 0 Then
                lObjstcCorr.Num_Ordem_Corr = 1
            End If
            .TxtCorretor(lObjstcCorr.Num_Ordem_Corr).Tag = ""
            .TxtCorretor(lObjstcCorr.Num_Ordem_Corr).Text = ""
            .txtCorComissao(lObjstcCorr.Num_Ordem_Corr).Text = ""
            If lObjstcCorr.Cod_Corr <> 0 Then
                .TxtCorretor(lObjstcCorr.Num_Ordem_Corr).Text = lObjstcCorr.gNom_Corretor
                    'O comando seguinte que move o c�digo do corretor deve ser executado depois do
                    'comando que posiciona o nome do corretor.
                .TxtCorretor(lObjstcCorr.Num_Ordem_Corr).Tag = lObjstcCorr.Cod_Corr
                .txtCorComissao(lObjstcCorr.Num_Ordem_Corr).Text = Format$(gM46V999.CDblx(lObjstcCorr.Per_Corr), "standard")
                lTotal = lTotal + gM46V999.CDblx(lObjstcCorr.Per_Corr)
            Else
                .txtCorComissao(lObjstcCorr.Num_Ordem_Corr).Text = Format$(0, "standard")
            End If
        Next lObjstcCorr
    End With
    
    If gM46V999.gTipoAcesso = DIGITACAO And lTotal = 0 Then
        txtComissao.Text = ""
    Else
        txtComissao.Text = Format(lTotal, "standard")
    End If
    
End Sub

Private Sub mpPosPro_Agenc(pColecaoProAgenc As Collection, pTipoCorr As String)
    'Procedure: posiciona corretor.

    'Par�metros...pLinha.........linha do corretor.
    '             pCodCorretor...c�digo do corretor.
    '             pPerCorretor...percentual do corretor.

    Dim lMensagem   As String       'Mensagem.
    Dim lSaida      As Variant      'Resultado da fun��o de busca de dados do corretor.
    Dim lObjstcCorr As stcA46V733B  'Objeto da classe de estrutura stcA46V733B (Tab_Ped_ProAgenc).
    Dim Corretor    As String
    Dim Tipo        As String
    Dim i As Integer
    
    If pColecaoProAgenc.Count = 0 Then Exit Sub
    
    
    With frmT47V028A
        Load frmT47V028A
        .flxAgenciamento.Rows = 1
        .flxProlabore.Rows = 1
        For Each lObjstcCorr In pColecaoProAgenc
            Set lObjstcCorr.M46V999 = gM46V999
            Corretor = Format(lObjstcCorr.Cod_Canal, String(7, "0")) & " - " & lObjstcCorr.gNom_Canal
            'Tipo = IIf(lObjstcCorr.Tip_Pagto = "R", "Retido na Fonte", "Ap�s Fatura")
            If pTipoCorr = "A" Then
                .flxAgenciamento.Rows = .flxAgenciamento.Rows + 1
                i = .flxAgenciamento.Rows - 1
                .flxAgenciamento.TextMatrix(i, 0) = Corretor
                .flxAgenciamento.TextMatrix(i, 1) = Format(lObjstcCorr.Perc_Canal, "00.00")
                .flxAgenciamento.TextMatrix(i, 2) = Tipo
            Else
                .flxProlabore.Rows = .flxProlabore.Rows + 1
                i = .flxProlabore.Rows - 1
                .flxProlabore.TextMatrix(i, 0) = Corretor
                .flxProlabore.TextMatrix(i, 1) = Format(lObjstcCorr.Perc_Canal, "00.00")
                .flxProlabore.TextMatrix(i, 2) = Tipo
            End If
        Next lObjstcCorr
        On Error Resume Next
        .flxAgenciamento.FixedRows = 1
        .flxProlabore.FixedRows = 1
    End With
    
End Sub

Private Sub mpPosCosseguro(pColecaoCosseguro As Collection)
    'Procedure: posiciona cosseguro.

    'Par�metros de entrada...lLinha........linha do cosseguro.
    '                        pCodigo.......c�digo do cosseguro.
    '                        pPercentual...percentual do cosseguro.

    Dim lMensagem   As String       'Mensagem.
    Dim lCount      As Long         'Contador.
    Dim lobjstcCoss As stcA46V706B  'Objeto da classe de estrutura stcA46V706B (TAB_PED_COSS).
    Dim lLinha      As Integer      'Linhas
    Dim lColecao    As Collection
    
    lLinha = 0
    Set lColecao = pColecaoCosseguro
    Set objstcPedido.ColecaoPedCoss = New Collection
    With frmT47V008A
        For Each lobjstcCoss In lColecao
            '1. Limpa campos de sa�da.
            lLinha = lLinha + 1
            .cboConSeguradora(lLinha).ListIndex = -1
            .txtCosPercentual(lLinha).Text = ""
    
            '2. Se c�digo n�o preenchido, encerra processo.
            If lobjstcCoss.Cod_Coss = 0 Then
                Exit Sub
            End If
                    
            '3. Posiciona campo de sa�da - c�digo.
            For lCount = 0 To .cboConSeguradora(lLinha).ListCount - 1
                If .cboConSeguradora(lLinha).ItemData(lCount) = lobjstcCoss.Cod_Coss Then
                    .cboConSeguradora(lLinha).ListIndex = lCount
                    .chkCosseguro(lLinha).Value = True
                    lobjstcCoss.Nom_Seguradora = .cboConSeguradora(lLinha).Text
                    objstcPedido.ColecaoPedCoss.Add lobjstcCoss, lobjstcCoss.Class_ID
                    Exit For
                End If
            Next
    
            '4. Posiciona percentual.
            .txtCosPercentual(lLinha).Text = lobjstcCoss.Per_Coss
            
        Next lobjstcCoss
    End With
End Sub

Private Sub mpPosiUnidProd(pMensagem As String)
Dim lCodUniProdutiva    As Integer
Dim lCodProdutor        As Integer
Dim lNomUniProdutiva    As String
Dim lNomProdutor        As String
Dim lObjstcCorr         As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).
            
    'c) Unidade produtiva, produtor e cosseguro (panUniProdutiva).
    '===1) Unidade produtiva.
    '===2) Produtor.
    lCodUniProdutiva = 0
    lCodProdutor = 0
    lCodUniProdutiva = objstcPedido.Cod_Unid_Prod
    lCodProdutor = objstcPedido.Cod_Prod
    If lCodUniProdutiva = 0 And lCodProdutor = 0 Then
        If objstcPedido.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = objstcPedido.ColecaoPedCorr.Item(1)
            If lObjstcCorr.Cod_Corr <> 0 Then
                Call gM46V999.gfObtUniProdutiva(lObjstcCorr.Cod_Corr, lCodUniProdutiva, lCodProdutor, pMensagem)
            End If
        End If
    End If
    frmT47V004A.lblNomUnidade.Tag = " "
    frmT47V004A.lblNomUnidade.Caption = " "
    frmT47V004A.lblNomProdutor.Tag = " "
    frmT47V004A.lblNomProdutor.Caption = " "
    
    If lCodUniProdutiva <> 0 Then
        If objstcPedido.Cod_Unid_Prod = 0 Then
            objstcPedido.Cod_Unid_Prod = lCodUniProdutiva
        End If
        lNomUniProdutiva = objstcPedido.gNom_UnidadeProd
        frmT47V004A.lblNomUnidade.Tag = lCodUniProdutiva
        frmT47V004A.lblNomUnidade.Caption = lNomUniProdutiva
    End If
    
    If lCodProdutor = 0 And lCodUniProdutiva <> 0 Then
        If objstcPedido.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = objstcPedido.ColecaoPedCorr.Item(1)
            If lObjstcCorr.Cod_Corr <> 0 Then
                Call gM46V999.gfObtProdutor(lObjstcCorr.Cod_Corr, lCodUniProdutiva, lCodProdutor, pMensagem)
            End If
        End If
    End If
    
    If lCodProdutor <> 0 Then
        If objstcPedido.Cod_Prod = 0 Then
            objstcPedido.Cod_Prod = lCodProdutor
        End If
        lNomProdutor = objstcPedido.gNom_Produtor
        frmT47V004A.lblNomProdutor.Tag = lCodProdutor
        frmT47V004A.lblNomProdutor.Caption = lNomProdutor
    End If
End Sub

Private Sub mpHabControles()
Dim lTotal As Double
Dim lMensagem   As String

    'Habilita bot�es e menus
    mnuExibir.Enabled = True
    mnuEditar.Enabled = True
    
    panBotAltTipEmissao.Enabled = True
    
    panBotValidar.Enabled = True
    
    panBotDadProposta.Enabled = True
        
    mnuArqImprimir.Enabled = True
    mnuArqVisualizar.Enabled = True
    
    cmdSair.Top = 0
    cmdSair.Visible = False
    
    mnuArqNovItem.Enabled = True
    mnuArqAbrItem.Enabled = True
    mnuArqExcProposta.Enabled = True
    barFerramentas.Visible = True
                
    panValores.Visible = True
    panPesquisa.Visible = True
    
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
        lblValorRisco.Left = 10
        txtValoRisco.Left = 1005
        
        lblCoefISVR.Left = 10
        txtCoefISVR.Left = 1005
        
        lblISBas.Left = 10
        txtISBas.Left = 1005
        
        panValores.Visible = True
        
    ElseIf gM46V999.gCod_Ramo = e_CodRamo_Condominio Then
        lblValorRisco.Left = 10
        txtValoRisco.Left = 1005
        
        lblCoefISVR.Left = 10
        txtCoefISVR.Left = 1005
        
        lblISBas.Left = 10
        txtISBas.Left = 1005
        
    ElseIf gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
        optTipPessoa(0).Value = True
        panValores.Visible = True
    ElseIf gM46V999.gCod_Ramo = e_CodRamo_Roubos Or gM46V999.gCod_Ramo = e_CodRamo_Riscos_diversos Then
        lblValorRisco.Left = 10
        txtValoRisco.Left = 1005
        
        lblCoefISVR.Left = 10
        txtCoefISVR.Left = 1005
        
        lblISBas.Left = 10
        txtISBas.Left = 1005
        panValores.Visible = True
    
    Else
        If gM46V111.gTipoGrupMulti = e_GrupoRamo_Generico Then
            lblValor2.Caption = "Pr�mio Resseguro: "
            lblValor3.Caption = "Pr�mio Tarif�rio: "
            lblValor4.Caption = "Pr�mio Comercial: "
            lblValor5.Caption = "C�mbio Emiss�o: "
            lblValor6.Visible = False
            lblValor7.Visible = False
            TxtValor6.Visible = False
            TxtValor7.Visible = False
            
        ElseIf gM46V111.gTipoGrupMulti = e_GrupoRamo_570_820 Then
            lblValor2.Caption = "Lim.M�ximo 2: "
            lblValor3.Caption = "Pr�mio Tarif�rio: "
            lblValor4.Caption = "Pr�mio Comercial: "
            lblValor5.Caption = "C�mbio Emiss�o: "
            lblValor6.Visible = False
            lblValor7.Visible = False
            TxtValor6.Visible = False
            TxtValor7.Visible = False

        ElseIf gM46V111.gTipoGrupMulti = e_GrupoRamo_Auto Then
            lblValor2.Caption = "Lim.M�ximo RC: "
            lblValor3.Caption = "Lim.M�ximo APP: "
            lblValor4.Caption = "Pr�mio Auto: "
            lblValor5.Caption = "Pr�mio RC: "
            lblValor6.Caption = "Pr�mio APP: "
            lblValor7.Caption = "In�cio Viagem: "

        ElseIf gM46V111.gTipoGrupMulti = e_GrupoRamo_Transp_Diversos Then
            lblValor2.Caption = "Pr�mio Resseguro: "
            lblValor3.Caption = "Pr�mio Tarif�rio: "
            lblValor4.Caption = "Pr�mio Comercial: "
            lblValor5.Caption = "C�mbio Emiss�o: "
            lblValor6.Visible = False
            TxtValor6.Visible = False
            lblValor7.Caption = "In�cio Viagem: "

        ElseIf gM46V111.gTipoGrupMulti = e_GrupoRamo_Acidentes_pessoais Then
            lblValor2.Caption = "Lim.M�ximo 2: "
            lblValor3.Caption = "Pr�mio Tarif�rio: "
            lblValor4.Caption = "Pr�mio Comercial: "
            lblValor5.Caption = "C�mbio Emiss�o: "
            lblValor7.Caption = "In�cio Viagem: "
            lblValor6.Visible = False
            TxtValor6.Visible = False
        End If
    End If

    lblValor2.ToolTipText = Replace(lblValor2.Caption, ":", "")
    lblValor3.ToolTipText = Replace(lblValor3.Caption, ":", "")
    lblValor4.ToolTipText = Replace(lblValor4.Caption, ":", "")
    lblValor5.ToolTipText = Replace(lblValor5.Caption, ":", "")
    lblValor6.ToolTipText = Replace(lblValor6.Caption, ":", "")
    lblValor7.ToolTipText = Replace(lblValor7.Caption, ":", "")

    TxtValor2.ToolTipText = lblValor2.ToolTipText
    TxtValor3.ToolTipText = lblValor3.ToolTipText
    TxtValor4.ToolTipText = lblValor4.ToolTipText
    TxtValor5.ToolTipText = lblValor5.ToolTipText
    TxtValor6.ToolTipText = lblValor6.ToolTipText
    TxtValor7.ToolTipText = lblValor7.ToolTipText
    If objstcPedido.Cod_Produto <> 0 Then
        gM46V999.gCodProduto = objstcPedido.Cod_Produto
    Else
        objstcPedido.Cod_Produto = gM46V999.gCodProduto
    End If
          
   Call mpCarregarPlano
            
    'Posiciona a tarifa
    Call mpObtProduto(objstcPedido.Cod_Produto, gM46V999.gCod_Ramo)
    lblNomProduto.Caption = gM46O001.gProNome
    lblNomProduto.Refresh
    If txtComissao.Text <> "" Then
        If gM46V999.CDblx(txtComissao.Text) = 0 Then
            txtComissao.Text = Format(gM46O001.gProComBasica, "standard")
            frmT47V004A.txtCorComissao(1).Text = Format(gM46O001.gProComBasica, "standard")
        End If
    Else
        txtComissao.Text = Format(gM46O001.gProComBasica, "standard")
        frmT47V004A.txtCorComissao(1).Text = Format(gM46O001.gProComBasica, "standard")
    End If
        
End Sub

Private Sub barFerramentas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
        Case "NOVO"
            Call mnuArqNovItem_Click
        Case "ABRIR"
            Call mnuArqAbrItem_Click
        Case "SALVAR"
            Call mnuArqSalvar_Click
        Case "IMPRIMIR"
            Call mnuArqImprimir_Click
        Case "VISUALIZAR"
            Call mnuArqVisualizar_Click
        Case "ARQUIVO"
            Call mnuArqCriArqImagem_Click
        Case "EXCLUIR"
            Call mnuEdiExcItem_Click
        Case "COPIAR"
            Call mnuEdiCopItem_Click
        Case "DESFAZER"
            Call mnuEdiDesfazer_Click
        Case "EXIDADPROPOSTA"
            Call mnuExiDadProposta_Click
        Case "BENS"
            Call mnuExiBem_Click
        Case "PROPOSTA"
            Call mnuExiDadProposta_Click
        Case "SAIR"
            Call mnuArqSaiItem_Click
        Case Else
            Call gM46V999.gpGraLog(1, "FTE2001 - Bot�o desconhecido na barra de ferramentas (" & _
                             UCase(Button.Key) & ").")
    End Select
End Sub

Private Sub mnuExiCustoEmis_Click()
    If gM46V111.gNum_Pi > 0 Then
        frmT47V015A.Show vbModal
    End If
End Sub

Private Sub mnuExiDadProposta_Click()
    If gM46V111.gNum_Pi > 0 Then
        frmT47V004A.Show vbModal, Me
    End If
End Sub

Private Sub mnuEdiCopItem_Click()
    Dim lMensagem       As String       'Mensagem.
    Dim lmfNovNumItem   As String       'Novo n�mero de item.
    Dim lDatAlt         As Double       'Data de altera��o dos registros
    Dim lobjstcItem     As stcA46V708B  'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Screen.MousePointer = vbHourglass
    If gM46V111.gNum_Pi > 0 Then
        '1. Verifica se tem dado a salvar.
        If mnuArqSalvar.Enabled = True Then
            If mfSalvar() = False Then
                Call gM46V999.gpGraLog(1, lMensagem)
                Screen.MousePointer = vbHourglass
                Exit Sub
            End If
        End If
    
        '2. Confirma com usu�rio.
        lMensagem = "AGE0025 - " & gM46V999.gPrefixo & "Confirma c�pia do item " & lblAntNumItem.Caption & "?"
        If MsgBox(lMensagem, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            Call gM46V999.gpGraLog(0, lMensagem & " Sim.")
        Else
            Call gM46V999.gpGraLog(0, lMensagem & " N�o.")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        '3. Obt�m novo n�mero de item.
        If mfNovNumItem(gM46V111.gComNosNumero, lmfNovNumItem) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        lDatAlt = gM46V999.gfForData(3)
        
        '4. Copia item.
        If mfCopItem(gM46V111.gComNosNumero, txtNumItem.Text, lmfNovNumItem, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        Set lobjstcItem = objstcPedLoc.gCopiaItem(CLng(lmfNovNumItem))
        objstcPedido.ColecaoPedLoc.Add lobjstcItem, lobjstcItem.Class_ID
        
        'f) Conclus�o.
        Call gM46V999.gpGraLog(1, "AGE0029 - " & gM46V999.gPrefixo & "C�pia realizada com sucesso: do item " & _
                         txtNumItem.Text & " para item " & lmfNovNumItem & ".")
    
        '5. Atualiza grade de itens.
        Call mpCriGraItens
        On Error Resume Next
        Set objstcPedLoc = objstcPedido.ColecaoPedLoc.Item(objstcPedLoc.NOSSO_NUMERO & Format(lmfNovNumItem, "00000"))
        On Error GoTo 0
        gCarga = True
        Call mpPosItem(objstcPedLoc)
        gCarga = False
        Call gpHabSalvar
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Function mfCopItem(ByVal pNosNumero As String, ByVal pAntNumItem As String, _
                           ByVal pmfNovNumItem As String, ByRef pMensagem As String) As Boolean
    'Fun��o: monta comando INSERT gen�rico para dados de proposta.

    'Par�metros de entrada...pAntNumItem....Anterior: n�mero do item.
    '                        pmfNovNumItem....Novo: n�mero do item.
    'Par�metro de sa�da......pMensagem......Mensagem.
    '********************Vari�veis padr�o*************************************
    Dim lCkPoint        As String      'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lCkStr          As String       'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
                                        'lCkPoint. (posi��o das variaveis em ordem de declara��o)
    '********************Vari�veis Declaradas para o processamento************
    Dim lbdBasDados As ADODB.Connection 'Base de dados.
    Dim lflTabela   As clsYasField      'Campos da tabela.
    Dim lrsTabela   As clsYasRecordSet  'Registro da tabela.
    Dim lCount      As Long             'Utilizado no for...next.
    Dim lIdeTabela  As String           'Nome da tabela.
    Dim lContinuar  As Boolean          'Controle para saber se continua ou n�o.
    Dim lSelect     As String
    Dim lInsert     As String
    Err = 0
    lCkPoint = ""
    lCkStr = ""
    On Error GoTo ControleErro
    mfCopItem = False

    '1. Obt�m registros da tabela.
    lCkPoint = "1"
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, lbdBasDados, _
            pMensagem) = False Then Exit Function
    lCkPoint = "2"
    For lCount = 1 To 7
        lContinuar = True
        Select Case lCount
            Case 1
                lIdeTabela = "TAB_PED_LOC"
            Case 2
                lIdeTabela = "TAB_PED_CLAU"
            Case 3
                lIdeTabela = "TAB_PED_COBERT"
            Case 4
                lIdeTabela = "TAB_PED_LOC_BEM"
            Case 5
                lIdeTabela = "TAB_PED_LOC_SEGURO"
            Case 6
                lIdeTabela = "TAB_PED_OBS"
            Case 7
                lIdeTabela = "TAB_PED_INSP"
        End Select
        lSelect = "SELECT * FROM " & lIdeTabela & " WHERE Nosso_Numero = '" & pNosNumero & _
              "' AND Num_Item = " & pAntNumItem
        lCkPoint = "3"
        If gM46V999.gfObtRegistro(lbdBasDados, lSelect, lrsTabela, pMensagem) = False Then
            Call gM46V999.gpFecha2(lbdBasDados)
            Exit Function
        End If
        lCkPoint = "4"
        If lrsTabela.EOF = True Then
            If lIdeTabela = "TAB_PED_LOC" Then
                pMensagem = "FTE2010 - " & gM46V999.gPrefixo & "Item n�o localizado (" & pAntNumItem & _
                                 " - c�pia do item)."
                Exit Function
            End If
            lContinuar = False
        End If
        lCkPoint = "5"
        If lContinuar = True Then
            '2. Grava nova cota��o/proposta.
            While Not lrsTabela.EOF
                lCkPoint = "6"
                lInsert = "INSERT INTO " & lIdeTabela & " ("
                For Each lflTabela In lrsTabela.Fields
                    lInsert = lInsert & lflTabela.Name & ", "
                Next lflTabela
                lInsert = Left$(lInsert, Len(lInsert) - 2) & ") VALUES ("
                For Each lflTabela In lrsTabela.Fields
                    Select Case UCase(lflTabela.Name)
                        Case "NUM_ITEM"
                            lInsert = lInsert & pmfNovNumItem & ", "                          'Num_Item.
                        Case "SIT_DOCUMENTO"
                            lInsert = lInsert & "1, "       'DOCUMENTO N�O CALCULADO.       'Sit_Documento.
                        Case "DAT_LIBER"
                            lInsert = lInsert & Left(gM46V999.gfForData(3), 8) & ", "                        'Dat_Alt.
                        Case "DAT_DIGIT"
                            lInsert = lInsert & Left(gM46V999.gfForData(3), 8) & ", "                        'Dat_DIGIT
                        Case "COD_USER_LIBER"
                            lInsert = lInsert & "'COPIA',"                                  'Cod_User_Alt.
                        Case Else
                            Select Case lflTabela.Type_
                                Case adBigInt, adBinary, adInteger, adLongVarBinary, adSmallInt, _
                                     adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, _
                                     adUnsignedTinyInt                          'Num�rico - inteiro.
                                    If IsNull(lflTabela.Value) Then
                                        lInsert = lInsert & "0, "
                                    Else
                                        lInsert = lInsert & lflTabela.Value & ", "
                                    End If
                                    lCkPoint = "7"
                                Case adCurrency, adSingle, adDouble, adNumeric             'Num�rico - com v�rgula.
                                    If IsNull(lflTabela.Value) Then
                                        lInsert = lInsert & "0, "
                                    Else
                                        lInsert = lInsert & gM46V999.gfForValor(1, lflTabela.Value) & ", "
                                    End If
                                    lCkPoint = "8"
                                Case adChar, adVarChar, adLongVarChar                           'Texto.
                                    If gM46V999.gfPreenchido(lflTabela.Value) Then
                                        lInsert = lInsert & "'" & gM46V999.gfForAspas(lflTabela.Value) & "', "
                                    Else
                                        lInsert = lInsert & "' ', "
                                    End If
                                    lCkPoint = "9"
                                Case adBoolean                                  'Sim/N�o.
                                    If IsNull(lflTabela.Value) Then
                                        lInsert = lInsert & "TRUE, "
                                    Else
                                        lInsert = lInsert & lflTabela.Value & ", "
                                    End If
                                    lCkPoint = "10"
                                Case Else
                                    Call gM46V999.gpFecha2(lbdBasDados)
                                    pMensagem = "FTE0008 - Tabela " & lIdeTabela & " com tipo do campo " & _
                                                lflTabela.Name & " diferente dos previstos."
                                    Exit Function
                            End Select
                    End Select
                Next lflTabela
                lCkPoint = "12"
                If Mid(lInsert, Len(lInsert), 1) = "," Then
                    lInsert = Left(lInsert, Len(lInsert) - 1) & ")"
                Else
                    lInsert = Left(lInsert, Len(lInsert) - 2) & ")"
                End If
                lCkPoint = "13"
                If gM46V999.gfExeSQL(lbdBasDados, lInsert, pMensagem) <> 0 Then
                    Call gM46V999.gpFecha2(lbdBasDados)
                    Exit Function
                End If
                lCkPoint = "14"
                lrsTabela.MoveNext
            Wend
        End If
        lCkPoint = "15"
    Next lCount
    
    mfCopItem = True
    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
        'Obtem o inicio da mensagem de erro
    pMensagem = "mfCopItem ERRO - " & Err & " " & Error & _
                Chr(13) & "CheckPoint " & lCkPoint & " - "
    'Carrega a variavel lCkStr
    On Error Resume Next
        lCkStr = lCount
        lCkStr = lCkStr & " | " & lIdeTabela
        lCkStr = lCkStr & " | " & lContinuar
    On Error GoTo 0
    'Completa a mensagem de erro
    pMensagem = pMensagem & lCkStr
    
    Call gM46V999.gpGraLog(0, pMensagem)
    
    mfCopItem = False

End Function


Private Function mfNovNumItem(ByVal pNosNumero As String, ByRef pNumItem As String) As Boolean
    'Fun��o: calcula novo n�mero de item.

    'Par�metro de entrada...pNosNumero...Nosso n�mero.
    'Par�metro de sa�da.....pNumItem.....N�mero do item.
    '********************Vari�veis padr�o*************************************
    Dim lCkPoint        As String      'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lCkStr          As String       'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
                                    'lCkPoint. (posi��o das variaveis em ordem de declara��o)
    '********************Vari�veis Declaradas para o processamento************
    Dim lobjstcPedLoc   As stcA46V708B  'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)

    On Error GoTo ControleErro
    Err = 0
    lCkPoint = ""
    lCkStr = ""
    
    mfNovNumItem = False
    '1. Obt�m m�ximo n�mero de item utilizado.
    lCkPoint = "1"
    pNumItem = 1
    For Each lobjstcPedLoc In objstcPedido.ColecaoPedLoc
        If lobjstcPedLoc.Num_Item + 1 > pNumItem Then
            pNumItem = lobjstcPedLoc.Num_Item + 1
        End If
        lCkPoint = "2"
    Next lobjstcPedLoc
    lCkPoint = "3"
    pNumItem = Format(pNumItem, "00000")
    mfNovNumItem = True
    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    'Obtem o inicio da mensagem de erro
    lCkStr = "mfNovNumItem ERRO - " & Err & " " & Error & _
                Chr(13) & "CheckPoint " & lCkPoint & " - "
    'Carrega a variavel lCkStr
    On Error Resume Next
        lCkStr = lCkStr & " | " & pNumItem
    On Error GoTo 0
    
    Call gM46V999.gpGraLog(0, lCkStr)
    
    mfNovNumItem = False
End Function

Private Sub mnuEdiDesfazer_Click()
Dim lMensagem   As String       'Mensagem
    '2. Obt�m registro e posiciona dados.
    If gM46V111.gNum_Pi > 0 Then
        gCarga = True
        Call mfPosComum(lMensagem)
        Call mpPosItem(objstcPedLoc)
        Call mpDesSalvar
        gCarga = False
    End If
End Sub

Private Sub mpCarregarPlano()
'Procedure: obt�m Plano.
    Dim lMensagem       As String           'Mensagem.
    Dim lrsTbm_Plano    As clsYasRecordSet  'Registro - Tbm_Ramos.
    Dim lBDConexao      As ADODB.Connection
    Dim lTamMax         As Long
    Dim lSelect         As String
    If cboRamo.ListIndex <> -1 Then
        'gM46V999.gBaseLocal = True
        If Not gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lBDConexao, lMensagem) Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If
        '2. Obt�m Todos os Ramos.
        lSelect = "SELECT * from Tab_Plano "
        lSelect = lSelect & " where cod_ramo = " & cboRamo.ItemData(cboRamo.ListIndex)
        If txtVigInicio.Text <> "" Then
            lSelect = lSelect & " And Dat_Vig_Ini <= " & gM46V999.gfForData(1, txtVigInicio.Text)
        Else
            lSelect = lSelect & " And Dat_Vig_Ini <= " & gM46V999.gfForData(1, Date)
        End If
        lSelect = lSelect & " order by Cod_Produto "
        
        If gM46V999.gfObtRegistro(lBDConexao, lSelect, lrsTbm_Plano, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
        If lrsTbm_Plano.EOF = True Then
            Call gM46V999.gpGraLog(1, "FNL0055 - Tabela de planos n�o localizado.")
            Exit Sub
        End If
        
        cboPlano(0).Clear
        cboPlano(1).Clear
        '3. Carrega Combo
        lTamMax = 0
        Do While Not lrsTbm_Plano.EOF
            cboPlano(0).AddItem lrsTbm_Plano("Cod_Produto") & " - " & lrsTbm_Plano("Nom_Prot")
            cboPlano(0).ItemData(cboPlano(0).NewIndex) = lrsTbm_Plano("Cod_Produto")
            cboPlano(1).AddItem lrsTbm_Plano("Cod_Produto") & " - " & lrsTbm_Plano("Nom_Prot")
            If Len(Trim(lrsTbm_Plano("nom_prot"))) > lTamMax Then
                lTamMax = Len(Trim(lrsTbm_Plano("nom_prot")))
            End If
            cboPlano(1).ItemData(cboPlano(1).NewIndex) = lrsTbm_Plano("Cod_Produto")
            lrsTbm_Plano.MoveNext
        Loop
        
        If cboPlano(0).ListCount = 1 Then
            cboPlano(0).ListIndex = 0
        End If
        
        If lTamMax > 50 Then
            'cboPlano(1).Width
        End If
        '4. Fecha
    End If
End Sub


Private Sub cboRamo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub cmdSair_Click()
    Unload Me
    End
End Sub

Private Sub mnuExiEspecAnt_Click()
    frmT47V022A.Show 1
End Sub

Private Sub mnuExiPrazo_Click()
    If gM46V111.gNum_Pi > 0 Then
        frmT47V003A.Show vbModal
    End If
End Sub


Private Sub mnuExiProAgenc_Click()
        frmT47V028A.Show 1
End Sub

Private Sub mnuSituacao_Click()
    If gM46V999.gfPreenchido(mnuSituacao.Caption) Then
        MsgBox mnuSituacao.Caption
    End If
End Sub

Private Sub optTipPessoa_Click(Index As Integer)
    Call gpLimPremios
    frmT47V004A.optTipPessoa(Index).Value = True
End Sub

Private Sub panBotAltTipEmissao_Click()
Dim lMensagem   As String       'Mensagem

    If gM46V111.gNum_Pi > 0 Then
        If objstcPedido.Class_ID_PI > 0 Then
            gCarga = True
            Call mfPosComum(lMensagem)
            Call mpPosItem(objstcPedLoc)
            gCarga = False
            Call gpHabSalvar
        End If
        Call mnuExiTipoEmissao_Click
        On Error Resume Next
    End If
    On Error GoTo 0
End Sub
Private Sub mnuExiTipoEmissao_Click()
    gProChamado = "T47V016A"
    On Error Resume Next
    
    frmT47V016A.Show vbModal, Me
    On Error GoTo 0
End Sub

Private Sub panBotCobertura_Click()
    mnuExiProAgenc_Click
End Sub

Private Sub panBotDadProposta_Click()
    Call mnuExiDadProposta_Click
End Sub

Private Sub panBotSair_Click()
    'Observa��o: controle criado para usu�rio sair do FORM utilizando tecla ESC. Invis�vel pelo usu�rio.
    Call mnuArqSaiItem_Click
End Sub


Private Sub mnuArqNovItem_Click()
    Dim lMensagem   As String           'Mensagem.
    Dim lNumItem    As String           'N�mero do item.
    Dim BDConexao   As ADODB.Connection 'Conex�o.
    Dim lobjstcPedLoc  As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC).
    Dim lTipEmissao  As e_TipEmissao
    If gM46V111.gNum_Pi > 0 Then
        Screen.MousePointer = vbHourglass
    
        '1. Salva registro (se n�o houver nada para salvar, altera pelo menos a "situa��o do documento").
        If mfSalvar = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        '2. Fecha janelas auxilires se houver.
        On Error Resume Next
        Unload frmT47V014A
        On Error GoTo 0
    
        '3. Obt�m novo n�mero de item.
        If mfNovNumItem(gM46V111.gComNosNumero, lNumItem) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If objstcPedLoc.Tip_Emissao = e_TipEmissao_RenCongenere Then
            lTipEmissao = e_TipEmissao_RenCongenere
        Else
            lTipEmissao = e_TipEmissao_SeguroNovo
        End If
        '4. Cria registro na tabela P0042200.
        If gM46V999.gfGraItem(BDConexao, gM46V111.gComNosNumero, gM46V111.gNumPedido, lNumItem, lMensagem, _
                 lTipEmissao, lobjstcPedLoc) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        lobjstcPedLoc.Cod_Ramo = gM46V999.gCod_Ramo
        objstcPedido.ColecaoPedLoc.Add lobjstcPedLoc, lobjstcPedLoc.Class_ID
        gCarga = True
        'Posiciona o item.
        Call mpPosItem(lobjstcPedLoc)
        
        gCarga = False
    
        Set objstcPedLoc = objstcPedido.ColecaoPedLoc.Item(lobjstcPedLoc.Class_ID)
        '6. Atualiza grade de itens.
        Call mpCriGraItens
    
        '7. Posiciona foco.
        On Error Resume Next
        txtRisCEP.SetFocus
        On Error GoTo 0
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub mnuArqAbrItem_Click()
    Dim lMensagem   As String   'Mensagem.

    If gM46V111.gNum_Pi > 0 Then
        Screen.MousePointer = vbHourglass
        
        '1. Verifica se tem dado a salvar.
        If mnuArqSalvar.Enabled = True Then
            If mfSalvar = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    
        '2. Exibe rela��o de itens.
        gProChamado = "T47V014A"
        Screen.MousePointer = vbDefault
        frmT47V014A.Show vbModeless, Me
        If frmT47V014A.cmdOk.Visible = True And frmT47V014A.cmdOk.Enabled = True Then
            frmT47V014A.cmdOk.SetFocus
        End If
    End If
End Sub
Private Sub mnuArqImprimir_Click()
    Screen.MousePointer = vbHourglass

    Call mpImprimir(1)
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuArqVisualizar_Click()
    Screen.MousePointer = vbHourglass
    ChDir (gM46V999.gAppPath)
    Call mpImprimir(0)
    
    Screen.MousePointer = vbDefault

End Sub
Private Sub mnuArqCriArqImagem_Click()
    Screen.MousePointer = vbHourglass

    Call mpImprimir(2)

    Screen.MousePointer = vbDefault
End Sub

Private Sub panBotTransmitir_Click()
    Call mnuArqTransmitir_Click
End Sub

Private Sub panBotValidar_Click()
    Call mnuEdiCalPremios_Click
End Sub

Private Sub panConfirmar_Click()
Dim lMensagem As String
Dim lDesTarifa As String

    If gM46V999.gfPreenchido(txtNumPI.Text) Then
        If Not IsNumeric(txtNumPI.Text) Then
            Call gM46V999.gpGraLog(1, "N�mero de PI inv�lido. Deve ser num�rico")
            Exit Sub
        End If
        txtNumPI.Text = Format(txtNumPI.Text, "0000000000")
        If txtNumPI.Text < 1000000000 Then
            Call gM46V999.gpGraLog(1, "N�mero de PI inv�lido.")
            Exit Sub
        End If

        Set objstcPedido = Nothing
        
        Call mpCarregarPI
        
        If mCarregarPI = False Then Exit Sub
        If objstcPedido Is Nothing Then Exit Sub
        
        If gM46V999.gfBusProduto(gM46V999.gHoje, Val(lblNomProduto.Tag), 0, 0, lDesTarifa, lMensagem, 710, (objstcPedido.Cod_Reg > 100)) = _
            False Then
            Call gM46V999.gpGraLog(1, lMensagem)
           Exit Sub
        End If
       
        If Left(frmT47V016A.cboCodReg.Text, 3) = "020" Or Left(frmT47V016A.cboCodReg.Text, 3) = "120" Or Left(frmT47V016A.cboCodReg.Text, 3) = "121" Or Left(frmT47V016A.cboCodReg.Text, 3) = "122" Then
            gM46V999.gProPremioMinimo = 0
            gM46V999.gProCusApolice = 0
        End If
        
        If Not objstcPedido Is Nothing Then
            panComum.Visible = True
            panValores.Visible = True
            Call mpCriGraItens
            Call mpDesSalvar
            gM46V111.gNum_Pi = txtNumPI.Text
        End If
    Else
        Call gM46V999.gpGraLog(1, "N�mero de PI n�o informado.")
        If txtNumPI.Visible = True Then
            txtNumPI.SetFocus
        End If
    End If
    
    If mCarregarPI = False Then
        Select Case lVerTipEmissao
        Case "S" 'Seguro Novo
            frmT47V016A.optTipEmissao(0).Value = True
            Call panBotAltTipEmissao_Click
    '        Exit Sub 'vai abrir a tela para digita��o
        Case "R" 'Renova��o
            frmT47V016A.optTipEmissao(1).Value = True
            Call panBotAltTipEmissao_Click
        Case "E" 'Endosso
            frmT47V016A.optTipEmissao(3).Value = True
            Call panBotAltTipEmissao_Click
        End Select
    End If
    
    Call mpConfiguraTela
    
End Sub
Public Sub mpCarregarPI()
Dim lobjstcCtrl     As stcA46V712B
Dim lObjstcCorr     As stcA46V705B
Dim lobjCtrl        As clsA46V712A
Dim lobjstcCobert   As stcA46V704B
Dim lobjCobert      As clsA46V704A
Dim lSelect         As String
Dim lrst            As clsYasRecordSet
Dim lBDGed          As ADODB.Connection
Dim lMensagem       As String
Dim lI              As Long

mCarregarPI = False

    If gM46V999.gfPreenchido(txtNumPI.Text) Then
        If IsNumeric(txtNumPI.Text) Then
            If Left(txtNumPI.Text, 2) > Right(Year(Date), 2) Then
                Call gM46V999.gpGraLog(1, "N�mero de PI inv�lido.")
                Exit Sub
            End If
            Set lobjstcCtrl = New stcA46V712B
            Set lobjCtrl = New clsA46V712A
            Set lobjCtrl.M46V999 = gM46V999
            
            lobjstcCtrl.NUM_PI = txtNumPI.Text
            gM46V999.gBaseLocal = False
            If Not lobjCtrl.mfCarregarColecao(lobjstcCtrl, "", "") Then
                Call gM46V999.gpGraLog(1, lobjCtrl.UltimoErro)
                Exit Sub
            End If
            'gM46V999.gBaseLocal = True
            If lobjCtrl.Colecao.Count > 0 Then
                'J� est� salvo no sistema
                'Carrega as tab peds....
                
                mCarregarPI = True
                
                Set lobjstcCtrl = lobjCtrl.Colecao.Item(1)
                
                Select Case lobjstcCtrl.Cod_Ramo
                    Case 510, 310, 112, 113, 114, 211, 212, 540, 550, 110, 120, 130, 170, 330, 331, 340, 350, 430, 440, 480, 490, 530, 560, 570, 640, 660, 680, 711, 730, 750, 810, 812, 820, 830, 930, 931, 932
                        Call gM46V999.gpGraLog(1, "Ramo n�o permitido.", vbCritical)
                        Exit Sub
                End Select
                If cboRamo.ListIndex = -1 Then
                    cboRamo.ListIndex = gM46V999.gfRetornarListIndex(cboRamo, lobjstcCtrl.Cod_Ramo)
                Else
                    If lobjstcCtrl.Cod_Ramo <> cboRamo.ItemData(cboRamo.ListIndex) Then
                        cboRamo.ListIndex = gM46V999.gfRetornarListIndex(cboRamo, lobjstcCtrl.Cod_Ramo)
                    End If
                End If
                
                Set objstcPedido = New stcA46V702B
                Set objstcPedido.M46V999 = gM46V999
                
                Set objPedido = New clsA46V702A
                Set objPedido.M46V999 = gM46V999
                
                gM46V999.gBaseLocal = False
                objstcPedido.NOSSO_NUMERO = lobjstcCtrl.NOSSO_NUMERO
                If Not objPedido.mfCarregarColecao(objstcPedido, "", "", True) Then
                    Call gM46V999.gpGraLog(1, objPedido.UltimoErro)
                    Exit Sub
                End If
                If objPedido.Colecao.Count > 0 Then
                    Set objstcPedido = objPedido.Colecao.Item(1)
                    If objstcPedido.ColecaoPedLoc.Count > 0 Then
                        Set objstcPedLoc = objstcPedido.ColecaoPedLoc.Item(1)
                    Else
                        Set objstcPedLoc = New stcA46V708B
                    End If
                End If
                
                Set lobjstcCobert = New stcA46V704B
                Set lobjCobert = New clsA46V704A
                Set lobjCobert.M46V999 = gM46V999
                Set lobjstcCobert.ObjPai = objstcPedLoc
                
                gM46V999.gBaseLocal = False
                lobjstcCobert.NOSSO_NUMERO = lobjstcCtrl.NOSSO_NUMERO
                If Not lobjCobert.mfCarregarColecao(lobjstcCobert, "", "") Then
                    Call gM46V999.gpGraLog(1, lobjCobert.UltimoErro)
                    Exit Sub
                End If
                If lobjCobert.Colecao.Count > 0 Then
                    Set lobjstcCobert = lobjCobert.Colecao.Item(1)
                End If
                
                gCarga = True
                Call mfPosComum(lMensagem)
                Call mpPosItem(objstcPedLoc)
                gCarga = False
            Else 'n�o est� salvo no sistema
                Set objstcPedido = New stcA46V702B
                Set objstcPedido.M46V999 = gM46V999
                
                Set objstcPedLoc = New stcA46V708B

                'Verifico se tem tab_ctrl_emis do ged
                If Not gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, lBDGed, lMensagem) Then
                    Call gM46V999.gpGraLog(1, lMensagem)
                    Exit Sub
                End If
                lSelect = " Select * from ged.dbo.tab_ctrl_emis where num_pi = " & txtNumPI.Text
                If Not gM46V999.gfObtRegistro(lBDGed, lSelect, lrst, lMensagem) Then
                    Call gM46V999.gpGraLog(1, lMensagem)
                    Exit Sub
                End If
                If Not lrst.EOF Then
                    objstcPedido.Nom_Segurado = lrst("Nom_Seg")
                    objstcPedido.Dat_Ini_Vig = lrst("Dat_Ini_Vig")
                    objstcPedido.Dat_Ter_Vig = lrst("Dat_Ter_Vig")
                    objstcPedido.Tip_Pessoa = IIf(Len(lrst("Num_Cgc_Cpf")) <= 11, e_TipPessoa_Fisica, e_TipPessoa_Juridica)
                    objstcPedido.Cod_Ramo = lrst("Cod_Ramo")
                    objstcPedido.Num_Cgc_Cpf = lrst("num_cgc_cpf")
                    objstcPedido.objstcCtrlEmis.Num_Bloq = lrst("num_bloq")
                    objstcPedLoc.Num_Item = 1
'                    objstcPedido.Nosso_Numero = lRst("Nosso_numero")
                    
                    Set lObjstcCorr = New stcA46V705B
                    lObjstcCorr.Cod_Corr = lrst("Cod_Corr")
                    lObjstcCorr.Per_Corr = lrst("Per_Corr")
                    If lObjstcCorr.Per_Corr > 100 Then lObjstcCorr.Per_Corr = lObjstcCorr.Per_Corr / 100
                    objstcPedido.ColecaoPedCorr.Add lObjstcCorr, lObjstcCorr.Class_ID
                    
                    If lrst("Cod_Corr_2") <> 0 Then
                        Set lObjstcCorr = New stcA46V705B
                        lObjstcCorr.Cod_Corr = lrst("Cod_Corr_2")
                        lObjstcCorr.Per_Corr = lrst("Per_Corr_2")
                        If lObjstcCorr.Per_Corr > 100 Then lObjstcCorr.Per_Corr = lObjstcCorr.Per_Corr / 100
                        objstcPedido.ColecaoPedCorr.Add lObjstcCorr, lObjstcCorr.Class_ID
                    End If
                        
                    If lrst("Cod_Corr_3") <> 0 Then
                        Set lObjstcCorr = New stcA46V705B
                        lObjstcCorr.Cod_Corr = lrst("Cod_Corr_3")
                        lObjstcCorr.Per_Corr = lrst("Per_Corr_3")
                        If lObjstcCorr.Per_Corr > 100 Then lObjstcCorr.Per_Corr = lObjstcCorr.Per_Corr / 100
                        objstcPedido.ColecaoPedCorr.Add lObjstcCorr, lObjstcCorr.Class_ID
                    End If
                    
                    
                Else
                    'Verifico se tem triagem
                    If Mid(Format(txtNumPI.Text, "0000000000"), 3, 1) = "5" Then
                        lSelect = " Select * from ged.dbo.tab_triagem_Transp where num_pi = " & txtNumPI.Text
                        If Not gM46V999.gfObtRegistro(lBDGed, lSelect, lrst, lMensagem) Then
                            Call gM46V999.gpGraLog(1, lMensagem)
                            Exit Sub
                        End If
                        If Not lrst.EOF Then
                            objstcPedido.Dat_Ini_Vig = lrst("Dat_Ini_Vig")
                            objstcPedido.Dat_Ter_Vig = lrst("dat_fim_vig")
                            objstcPedido.Cod_Ramo = lrst("Cod_Ramo")
                        End If
                        
                    Else
                        lSelect = " Select * from ged.dbo.tab_triagem where num_pi = " & txtNumPI.Text
                        If Not gM46V999.gfObtRegistro(lBDGed, lSelect, lrst, lMensagem) Then
                            Call gM46V999.gpGraLog(1, lMensagem)
                            Exit Sub
                        End If
                        If Not lrst.EOF Then
                            If Not IsNull(lrst("Dat_Ini_Vig")) Then
                                objstcPedido.Dat_Ini_Vig = lrst("Dat_Ini_Vig")
                            End If
                            If Not IsNull(lrst("dat_fim_vig")) Then
                                objstcPedido.Dat_Ter_Vig = lrst("dat_fim_vig")
                            End If
                            If Not IsNull(lrst("Cod_Ramo")) Then
                                objstcPedido.Cod_Ramo = lrst("Cod_Ramo")
                            End If
                        End If
                    End If
                End If
                gCarga = True
                If objstcPedido.Cod_Ramo = 0 Then
                    If cboRamo.ListIndex <> -1 Then
                        objstcPedido.Cod_Ramo = cboRamo.ItemData(cboRamo.ListIndex)
                    Else
                        Call gM46V999.gpGraLog(1, "Selecione o ramo.")
                        If cboRamo.Visible And cboRamo.Enabled Then
                            cboRamo.SetFocus
                        End If
                        Set objstcPedido = Nothing
                        Exit Sub
                    End If
                Else
                    Select Case objstcPedido.Cod_Ramo
                         Case 510, 310, 112, 113, 114, 211, 212, 540, 550, 110, 120, 130, 170, 330, 331, 340, 350, 430, 440, 480, 490, 530, 560, 570, 640, 660, 680, 711, 730, 750, 810, 812, 820, 830, 930, 931, 932
                            Call gM46V999.gpGraLog(1, "Este cadastro n�o permite os ramos 310, 112, 113, 114, 211, 212, 540, 550, 310, 112, 113, 114, 211, 212, 540, 550, 110, 120, 130, 170, 330, 331, 340, 350, 430, 440, 480, 490, 530, 560, 570, 640, 660, 680, 711, 730, 750, 810, 812, 820, 830, 930, 931, 932")
                            Set objstcPedido = Nothing
                            Exit Sub
                    End Select
                End If
                'sandra-aqui
                Set lobjstcCobert = New stcA46V704B
                Set lobjCobert = New clsA46V704A
                Set lobjCobert.M46V999 = gM46V999
                Set lobjstcCobert.ObjPai = objstcPedLoc
                
                gM46V999.gBaseLocal = False
                lobjstcCobert.NOSSO_NUMERO = lobjstcCtrl.NOSSO_NUMERO
                If gM46V999.gfPreenchido(lobjstcCobert.NOSSO_NUMERO) Then
                    If Not lobjCobert.mfCarregarColecao(lobjstcCobert, "", "") Then
                        Call gM46V999.gpGraLog(1, lobjCobert.UltimoErro)
                        Exit Sub
                    End If
                End If
                Call mfPosComum(lMensagem)
                Call mpPosItem(objstcPedLoc)
                gCarga = False
                mCarregarPI = True
            End If
        End If
    End If
End Sub

Private Sub TabMod_BeforeClick(Cancel As Integer)
  If cboPlano(1).ListIndex = -1 Then Cancel = True
End Sub

Private Sub TabMod_Click()
Dim lI  As Long

    For lI = 0 To flxGridCoberturaDemais.UBound
        On Error Resume Next
        flxGridCoberturaDemais(lI).Visible = False
        On Error GoTo 0
    Next lI
    
    If gM46V999.gfPreenchido(TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag) Then
        lI = TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag

        If cboPlano(1).ListIndex = -1 Then Exit Sub
        If flxGridCoberturaDemais.UBound >= (lI) Then
            flxGridCoberturaDemais(lI).Visible = True
            flxGridCoberturaDemais(lI).ZOrder
            txtModalidade.Visible = False
            cboPlano(1).Visible = False
        End If
    End If
    
End Sub
Private Sub txtApolice_Change()
    Call gpHabSalvar
End Sub

Private Sub txtApolice_GotFocus()
    With txtApolice
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub


Private Sub txtApolice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtClassificacao_Change()
    Call gpLimPremios
End Sub

Private Sub txtCoefISVR_Change()
    Call gpLimPremios
End Sub

Private Sub txtCoefISVR_GotFocus()
    With txtCoefISVR
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtCoefISVR_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtComissao_Change()
    Call gpLimPremios
End Sub

Private Sub txtComissao_GotFocus()
    With txtComissao
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtComissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtComissao_LostFocus()
    If IsNumeric(txtComissao.Text) Then
        txtComissao.Text = Format(txtComissao.Text, "standard")
    Else
        txtComissao.Text = Format(0, "standard")
    End If
    frmT47V004A.txtCorComissao(1).Text = txtComissao.Text
End Sub


Private Sub txtEndosso_Change()
    Call gpHabSalvar
End Sub

Private Sub txtEndosso_GotFocus()
    With txtEndosso
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtEndosso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtFatura_GotFocus()
    With txtFatura
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtFatura_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtISBas_Change()
    Call gpLimPremios
End Sub

Private Sub txtISBas_GotFocus()
    With txtISBas
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtISBas_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtISBas_LostFocus()
Dim lI As Integer

    With txtISBas
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
    If CDbl(Val(txtValoRisco.Text)) <> 0 Then
        txtCoefISVR.Text = CDbl(Val(txtISBas.Text)) / CDbl(Val(txtValoRisco.Text))
    Else
        txtCoefISVR.Text = "0,00"
    End If
End Sub


Private Sub txtLider_Change()
    Call gpHabSalvar
End Sub

Private Sub txtLider_GotFocus()
    With txtLider
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtLider_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtNumItem_DblClick()
    Call gpLimPremios
End Sub

Private Sub txtNumItem_GotFocus()
    With txtNumItem
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtNumItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNumItem_LostFocus()
Dim lobjstcLoc  As stcA46V708B
    If IsNumeric(txtNumItem.Text) Then
        txtNumItem.Text = Format(Val(txtNumItem.Text), "00000")
    Else
        txtNumItem.Text = ""
    End If
    
    If txtNumItem.Text <> lblAntNumItem.Caption Then
        For Each lobjstcLoc In objstcPedido.ColecaoPedLoc
            If lobjstcLoc.Num_Item = Val(txtNumItem.Text) Then
                Call gM46V999.gpGraLog(1, "N�mero do item duplicado. Informar outro n�mero.")
                txtNumItem.Text = lblAntNumItem.Caption
                txtNumItem.SetFocus
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub txtNumPI_GotFocus()
    With txtNumPI
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtNumPI_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtOrdem_GotFocus()
    With txtOrdem
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOrdem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtRisBairro_Change()
    Call gpLimPremios
End Sub

Private Sub txtRisBairro_GotFocus()
    With txtRisBairro
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRisBairro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtRisCEP_Change()
    Call gpLimPremios
End Sub

Private Sub txtRisCEP_GotFocus()
    With txtRisCEP
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRisCEP_KeyPress(KeyAscii As Integer)
Dim lI As Byte
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
     Select Case KeyAscii
        Case 48 To 57
            'Caracteres v�lidos
            lI = txtRisCEP.SelStart
            If Len(txtRisCEP.Text) = 8 And lI <> 0 Then
                KeyAscii = 0
            End If
        Case 8
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtRisCEP_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(txtRisCEP.Text) > 8 And InStr(1, txtRisCEP.Text, "-") = 0 Then
        txtRisCEP.Text = Left(txtRisCEP.Text, 8)
    End If
End Sub

Private Sub txtRisCEP_LostFocus()
'Rotina Alterada para atender a nova Precifica��o
    Dim lbdP0043000 As ADODB.Connection 'Arquivo - P0043000.
    Dim lBDP00Multi As ADODB.Connection 'Arquivo - P00Multi.
    Dim lMensagem   As String           'Mensagem.
    Dim lrsP0043000 As clsYasRecordSet  'Registro - P0043000.
    Dim lPos        As Byte
    Dim lCEP        As String
    Dim lSelect     As String
    If gCarga Then Exit Sub
    


    Screen.MousePointer = vbHourglass
    
    With txtRisCEP
        '1. Trata campo do CEP.
        If gM46V999.gfPreenchido(mAntCEP) Then
            If gM46V999.gfPreenchido(.Text) Then
                If Replace(mAntCEP, "-", "") = Replace(.Text, "-", "") Then
                    .Text = Left(.Text, 5) & "-" & Right(.Text, 3)
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
                    Call gM46V999.gpGraLog(1, "FGE0001 - " & gM46V999.gPrefixo & "CEP inv�lido (endere�o do segurado - " & .Text & _
                                     " - n�o num�rico - fora do formato nnnnn-nnn).")
                    .Text = ""
                    If .Visible = True And .Enabled = True Then
                        .SetFocus
                    End If
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
        End Select
        .Text = Format(Replace(Trim(.Text), "-", ""), "00000-000")
        If Len(Trim(.Text)) = 8 Then
            If IsNumeric(.Text) Then
                .Text = Left$(.Text, 5) & "-" & Right$(.Text, 3)
            End If
        End If
    
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
            Call gM46V999.gpGraLog(1, "FNL0001 - " & gM46V999.gPrefixo & "CEP n�o localizado (" & .Text & ").")
            Call gM46V999.gpFecha2(lbdP0043000)
            Screen.MousePointer = vbDefault
            mAntCEP = "0"
            Exit Sub
        End If
        'c) Posiciona endere�o.
        If gM46V999.gfPreenchido(lrsP0043000!TIP_LOGR) Then
            txtRisTipLogradouro.Text = lrsP0043000!TIP_LOGR
        Else
            txtRisTipLogradouro.Text = ""
        End If
        If gM46V999.gfPreenchido(lrsP0043000!Nom_Logr) Then
            txtRisNomLogradouro.Text = Left(lrsP0043000!Nom_Logr, IIf(InStr(1, lrsP0043000!Nom_Logr, "-") > 0, InStr(1, lrsP0043000!Nom_Logr, "-") - 2, 200))
        Else
            txtRisNomLogradouro.Text = ""
        End If
        'Ficha 1300287
        If Mid(cboRamo.Text, 1, 3) = "710" Then
            txtRisNomLogradouro.Text = "CONFORME ESPECIFICACAO EM ANEXO"
        End If
        txtRisComplemento.Text = ""
        If gM46V999.gfPreenchido(lrsP0043000!Nom_Bairro) Then
            txtRisBairro.Text = lrsP0043000!Nom_Bairro
        Else
            txtRisBairro.Text = ""
        End If
        If gM46V999.gfPreenchido(lrsP0043000!Nom_Cid) Then
            txtRisCidade.Text = lrsP0043000!Nom_Cid
        Else
            txtRisCidade.Text = ""
        End If
        If gM46V999.gfPreenchido(lrsP0043000!Sig_UF) Then
            cboRisUF.Text = lrsP0043000!Sig_UF
        Else
            cboRisUF.ListIndex = -1
        End If
        'd) Fecha tabela e arquivo.
        Call gM46V999.gpFecha2(lbdP0043000)
        mAntCEP = .Text
    End With
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub cboPlano_Click(Index As Integer)
Dim lI  As Long
    Call gpHabSalvar
    If Index = 1 Then
        'Verifico se esta modalidade j� foi selecionada antes
        If cboPlano(1).ListIndex <> -1 Then
            For lI = 1 To TabMod.Tabs.Count
                If TabMod.Tabs.Item(lI).Tag = cboPlano(1).ItemData(cboPlano(1).ListIndex) Then
                    Call gM46V999.gpGraLog(1, "Esta modalidade j� foi selecionada antes.")
                    Exit Sub
                End If
            Next lI
        End If
        If flxGridCoberturaDemais.UBound = 0 Then
            flxGridCoberturaDemais(0).Visible = True
            txtModalidade.Visible = False
            cboPlano(1).Visible = False
            
            If cboPlano(1).ListIndex <> -1 Then
                TabMod.Tabs.Item(1).Caption = Trim(cboPlano(1).Text)

                flxGridCoberturaDemais(0).Tag = cboPlano(1).ItemData(cboPlano(1).ListIndex)
                Call mpCarregarCoberturas
            End If
        Else
            If cboPlano(1).ListIndex <> -1 Then
                lI = TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag
                flxGridCoberturaDemais(lI).Visible = True
                flxGridCoberturaDemais(lI).ZOrder
                flxGridCoberturaDemais(lI).Tag = cboPlano(1).ItemData(cboPlano(1).ListIndex)
                TabMod.Tabs.Item(TabMod.SelectedItem.Index).Caption = Trim(cboPlano(1).Text)
                Call mpCarregarCoberturas
            End If
        End If
    End If
    
    gM46V999.gProPreMinParcela = gM46V999.gProPreMinParcelaTAB
    gM46V999.gProPremioMinimo = gM46V999.gProPremioMinimoTAB
    gM46V999.gProQuaMaxParcelas = gM46V999.gProQuaMaxParcelasTAB
        
    
End Sub

Private Sub cmdClassificacao_Click()
Dim lCodClasf   As String
Dim lDescricao As String
    Call gpLimPremios
    If cboPlano(0).ListIndex = -1 Then
        Call gM46V999.gpGraLog(1, "FNP0010 - Plano inv�lido (n�o informado).")
        Exit Sub
    ElseIf gM46V999.gCodProduto = 0 Then
        Call gM46V999.gpGraLog(1, "FGE0091 - Tarifa inv�lida (n�o informada).")
        Exit Sub
    ElseIf gM46V999.gCodProduto > gM46V999.TarifaAntiga Then
    Else
        Select Case cboPlano(0).ItemData(cboPlano(0).ListIndex)
            Case Is > 1010
            Case Else
                'If Val(txtRisRegiao.Tag) = 0 Then
                '    Call gM46V999.gpGraLog(1, " FGE0104 - Regi�o n�o informada. Informar um CEP de risco v�lido.")
                '    Exit Sub
                'End If
        End Select
    End If
    Call frmT47V002A.gpIntanciarForm(lCodClasf, lDescricao)
    txtClassificacao.Text = lDescricao
    txtClassificacao.Tag = lCodClasf
End Sub
Private Sub mpCarregarClassificacao()
Dim lCodClasf   As String
Dim lDescricao  As String
    If gCarga = True Or cboPlano(0).ListIndex = -1 Then Exit Sub
    Call frmT47V002A.gpIntanciarForm(lCodClasf, lDescricao)
    txtClassificacao.Text = lDescricao
    txtClassificacao.Tag = lCodClasf
    
End Sub

Private Sub mpCarregarProduto()
'Procedure: obt�m Plano.
    Dim lMensagem       As String           'Mensagem.
    Dim rsTabProd       As clsYasRecordSet  'Registro - Tbm_Ramos.
    Dim lBDConexao      As ADODB.Connection
    Dim lSelect         As String
    
    If Not gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lBDConexao, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If
    '2. Obt�m Todos os Ramos.
    lSelect = "SELECT * from Tab_Produto WHERE Dat_Vig_Ini <= " & gM46V999.gHoje
    lSelect = lSelect & " and cod_ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " ORDER BY Cod_Produto DESC , Cod_Sequencia desc"
    
    If gM46V999.gfObtRegistro(lBDConexao, lSelect, rsTabProd, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    If rsTabProd.EOF = True Then
        Call gM46V999.gpGraLog(1, "FNL0055 - Tabela de produto n�o localizado.")
        Exit Sub
    End If
            
    If Not rsTabProd.EOF Then
        gM46V999.gCodProduto = gM46V999.Check(rsTabProd!Cod_Produto)
        gM46V999.gCodSequencia = gM46V999.Check(rsTabProd!Cod_Sequencia)
        gM46O001.gProNome = gM46V999.Check(rsTabProd("Nom_Prot"))
    End If
    
    lblNomProduto.Caption = gM46O001.gProNome
    lblNomProduto.Refresh
    '4. Fecha
End Sub
Private Sub mpObtProduto(ByVal pCodProduto As Integer, ByVal pCodRamo As Long)
    'Procedure: obt�m produto e informa��es dependentes de produto.

    'Par�metros de entrada...pCodProduto.....c�digo do produto.
    '                        pCodProtocolo...c�digo do protocolo.
    '                        pCodSequencia...c�digo da seq��ncia.

    Dim lbdTAB_PROD     As ADODB.Connection 'Base de dados - TAB_PROD.
    Dim lCodProtocolo   As Long             'C�digo do protocolo.
    Dim lCodSequencia   As Long             'C�digo da seq��ncia.
    Dim lI              As Byte             'Utilizado com For...Next.
    Dim lMensagem       As String           'Mensagem.
    Dim lPosicao        As Integer          'Utilizado com InStr.
    Dim lrsTAB_PROD     As clsYasRecordSet  'Registro - TAB_PROD.
    Dim lCarga          As Boolean
    Dim lSelect         As String
    
    '1. Inicializa vari�veis globais - dados do produto.
        
    gM46O001.gProNome = " "
    gM46V999.gCodProduto = 0
    gM46V999.gCodSequencia = 0
    

    '2. Se produto zerado, encerra processo.
    If pCodProduto = 0 Or pCodRamo = 0 Then
        Exit Sub
    End If

    '3. Abre arquivo.
    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_PROD, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        gM46V999.gCodProduto = 0
        gM46V999.gCodSequencia = 0
        Exit Sub
    End If

    '4. Obt�m produto informado.
    lSelect = "SELECT * FROM TAB_PRODUTO WHERE Cod_Produto = " & pCodProduto & " AND Cod_Ramo = " & _
              pCodRamo
    If gM46V999.gfObtRegistro(lbdTAB_PROD, lSelect, lrsTAB_PROD, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_PROD)
        Call gM46V999.gpGraLog(2, lMensagem)
        gM46V999.gCodProduto = 0
        
        gM46V999.gCodSequencia = 0
        Exit Sub
    End If
    If lrsTAB_PROD.EOF = True Then
        Call gM46V999.gpFecha2(lbdTAB_PROD)
        Call gM46V999.gpGraLog(1, "FNL0002 - " & gM46V999.gPrefixo & "Produto n�o localizado (" & _
                         Format$(pCodProduto, "000") & "/" & Format$(pCodRamo, "000") & ").")
        gM46V999.gCodProduto = 0
        gM46V999.gCodSequencia = 0
        Exit Sub
    End If

    '5. Carrega vari�veis globais.
    gM46V999.gCodProduto = lrsTAB_PROD!Cod_Produto
    gM46V999.gCodSequencia = lrsTAB_PROD!Cod_Sequencia
    gM46O001.gProNome = lrsTAB_PROD!Nom_Prot
    gM46V999.gProQuaMaxParcelasTAB = lrsTAB_PROD!Qua_Max_Parcelas
    gM46V999.gProCusEndosso = gM46V999.Check(lrsTAB_PROD!Val_Cus_Emis_Endo)
    gM46V999.gProCusApolice = gM46V999.Check(lrsTAB_PROD!Val_Cus_Emis_Apol)

    If lVerTipEmissao = "E" Then
        frmT47V015A.txtCusEmissao.Text = gM46V999.gProCusEndosso
    Else
        frmT47V015A.txtCusEmissao.Text = gM46V999.gProCusApolice
    End If

    Call gM46V999.gpFecha2(lbdTAB_PROD)
        
End Sub

Private Sub cboRamo_Click()
    Call gpHabSalvar
    If cboRamo.ListIndex <> -1 Then
        gM46V999.gCod_Ramo = cboRamo.ItemData(cboRamo.ListIndex)
    End If
    'Carregar produto
    Call mpCarregarProduto
    'Carregar plano
    Call mpCarregarPlano
    'Carregar coberturas
    Call mpCarregarCoberturas
End Sub

Private Sub CmbCoberturas_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lI  As Integer
    lI = TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag
    Call gpHabSalvar
        With flxGridCoberturaDemais(lI)
            Select Case KeyCode
                Case 37
                    .ColSel = .ColSel - 1
                    .Col = .Col - 1
                    flxGridCoberturaDemais_Click (lI)
                    .SetFocus
                Case 39
                    .ColSel = .ColSel + 1
                    .Col = .Col + 1
                    flxGridCoberturaDemais_Click (lI)
                    .SetFocus
            End Select
        End With
End Sub


Private Sub CmbCoberturas_LostFocus()
Dim lI  As Integer
    lI = Val(TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag)
    
    If mLinhaFocus = 0 Then mLinhaFocus = flxGridCoberturaDemais(lI).Row
    If CmbCoberturas.ListIndex <> -1 Then
        On Error Resume Next
        If flxGridCoberturaDemais(lI).TextMatrix(mLinhaFocus, 0) = "" Then flxGridCoberturaDemais(lI).TextMatrix(mLinhaFocus, 0) = "0"
        flxGridCoberturaDemais(lI).TextMatrix(mLinhaFocus, 1) = CmbCoberturas.Text
        flxGridCoberturaDemais(lI).TextMatrix(mLinhaFocus, 0) = CmbCoberturas.ItemData(CmbCoberturas.ListIndex)
    Else
        flxGridCoberturaDemais(lI).TextMatrix(mLinhaFocus, 1) = ""
    End If
    Call mpAtualTamanho
    CmbCoberturas.Visible = False
    DoEvents
End Sub

Private Sub flxGridCoberturaDemais_Click(Index As Integer)

    If flxGridCoberturaDemais(Index).ColSel = 1 And flxGridCoberturaDemais(Index).Text <> "" Then Exit Sub
    
    If flxGridCoberturaDemais(Index).Tag = "Travado" Then
        flxGridCoberturaDemais(Index).Col = 0
        flxGridCoberturaDemais(Index).ColSel = flxGridCoberturaDemais(Index).Cols - 1
        Exit Sub
    End If
    
    If flxGridCoberturaDemais(Index).ColSel = mColunaFocus And flxGridCoberturaDemais(Index).RowSel = mLinhaFocus Then
        Exit Sub
    End If
    
    If EstaCarregando Then Exit Sub
    EstaCarregando = True
   
    If CmbCoberturas.Visible Then
        DoEvents
    End If
    
    If Not PosicaoAnterior Then
        mColunaFocus = flxGridCoberturaDemais(Index).ColSel
        mLinhaFocus = flxGridCoberturaDemais(Index).RowSel
    End If

    If flxGridCoberturaDemais(Index).ColSel = 1 Then
        If Not CmbCoberturas.Visible Then
            CmbCoberturas.Left = flxGridCoberturaDemais(Index).CellLeft + 70
            CmbCoberturas.Width = flxGridCoberturaDemais(Index).CellWidth + 10
            CmbCoberturas.Top = flxGridCoberturaDemais(Index).Top + flxGridCoberturaDemais(Index).CellTop - 25
            CmbCoberturas.ListIndex = gM46V999.gfRetornarListIndex(CmbCoberturas, flxGridCoberturaDemais(Index).TextMatrix(flxGridCoberturaDemais(Index).RowSel, 0))
            CmbCoberturas.ZOrder 0
            CmbCoberturas.Visible = True
            CmbCoberturas.SetFocus
        End If
    Else
        flxGridCoberturaDemais(Index).TextMatrix(mLinhaFocus, mColunaFocus) = Format(flxGridCoberturaDemais(Index).TextMatrix(mLinhaFocus, mColunaFocus), "standard")
        If mColunaFocus = 5 Then
            If IsNumeric(flxGridCoberturaDemais(Index).TextMatrix(mLinhaFocus, mColunaFocus)) Then
                If CDbl(flxGridCoberturaDemais(Index).TextMatrix(mLinhaFocus, mColunaFocus)) >= 100 Then
                    Call gM46V999.gpGraLog(1, "Percentual de franquia inv�lido. Deve ser menor que 100%")
                    flxGridCoberturaDemais(Index).TextMatrix(mLinhaFocus, mColunaFocus) = ""
                End If
            End If
        End If
        flxGridCoberturaDemais(Index).Refresh
    End If
    mColunaFocus = flxGridCoberturaDemais(Index).ColSel
    mLinhaFocus = flxGridCoberturaDemais(Index).RowSel
    
    EstaCarregando = False

End Sub



Private Sub flxGridCoberturaDemais_GotFocus(Index As Integer)
    GridFocado = "flxGridCoberturaDemais" & Index
    mLinhaFocus = 0
    flxGridCoberturaDemais_Click (Index)
End Sub

Private Sub flxGridCoberturaDemais_LostFocus(Index As Integer)
Dim lI      As Integer
    For lI = 1 To flxGridCoberturaDemais(Index).Rows - 1
        If IsNumeric(flxGridCoberturaDemais(Index).TextMatrix(lI, 5)) Then
            If CDbl(flxGridCoberturaDemais(Index).TextMatrix(lI, 5)) >= 100 Then
                Call gM46V999.gpGraLog(1, "Percentual de franquia inv�lido. Deve ser menor que 100%")
                flxGridCoberturaDemais(Index).TextMatrix(lI, 5) = ""
            End If
        End If
    Next lI
    
            
    GridFocado = ""
End Sub

Private Sub flxGridCoberturaDemais_RowColChange(Index As Integer)
    PosicaoAnterior = True
    flxGridCoberturaDemais_Click (Index)
    PosicaoAnterior = False
End Sub
Private Sub flxGridCoberturaDemais_KeyPress(Index As Integer, KeyAscii As Integer)
    Call gpHabSalvar
    If flxGridCoberturaDemais(Index).Tag = "Travado" Then
        flxGridCoberturaDemais(Index).Col = 0
        flxGridCoberturaDemais(Index).ColSel = flxGridCoberturaDemais(Index).Cols - 1
        Exit Sub
    End If
    
    With flxGridCoberturaDemais(Index)
        If .RowSel > 0 And (.ColSel > 1) Then
            Select Case KeyAscii
                Case 8              'BackSpace
                    If Len(.TextMatrix(.RowSel, .ColSel)) > 0 Then
                        .TextMatrix(.RowSel, .ColSel) = Left(.TextMatrix(.RowSel, .ColSel), Len(.TextMatrix(.RowSel, .ColSel)) - 1)
                        EstaAlterando = True
                    End If
                Case 13             'Enter
                    If .RowSel < .Rows - 1 Then
                        .RowSel = .RowSel + 1
                        .Row = .RowSel
                    End If
                Case 48 To 57, 58, 44, 45
                        If Len(CStr(Val(.TextMatrix(.RowSel, .ColSel)))) < 13 Then
                            .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
                            EstaAlterando = True
                        Else
                            Beep
                        End If
                Case 22
                    .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Clipboard.GetText
                    EstaAlterando = True
'                Case 32, 33, 35 To 38, 40 To 47, 61, 65 To 93, 95, 97 To 123, 126, 58
                Case Else
                    'If .ColSel = 3 Then
                    '    .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
                    '    EstaAlterando = True
                    'End If
            End Select
            Call mpAtualTamanho
            If KeyAscii <> 9 Then
                '.SetFocus
            End If
        End If
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        Call mpIncluirLinha
    End If
    
End Sub

Private Sub mpIncluirLinha()
Dim lI  As Integer
    
    lI = TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag
    If flxGridCoberturaDemais(lI).Tag = "Travado" Then
        Exit Sub
    End If
    
    
    If GridFocado = "flxGridCoberturaDemais" & lI Then
        With flxGridCoberturaDemais(lI)
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .RowSel = .Rows - 1
            .ColSel = 0
            .Col = 1
            .RowHeight(.Row) = CmbCoberturas.Height
            .SetFocus
            .ColSel = 1
            .RowSel = .Rows - 1
            flxGridCoberturaDemais_Click (lI)
        End With
    End If

End Sub
Private Sub mpExcluirLinha()
Dim lI  As Integer
    
    lI = TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag
    If flxGridCoberturaDemais(lI).Tag = "Travado" Then
        Exit Sub
    End If
    
    
    If GridFocado = "flxGridCoberturaDemais" & lI Then
        With flxGridCoberturaDemais(lI)
            If .RowSel <> 1 Then
                If .Rows = 2 Then
                    .Rows = 1
                    .Rows = 2
                    .FixedRows = 1
                Else
                    If mLinhaFocus = .RowSel Then
                        mLinhaFocus = mLinhaFocus - 1
                    End If
                    .RemoveItem .RowSel
                End If
            End If
        End With
    End If

End Sub


Private Sub flxGridCoberturaDemais_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim ColAnt As Integer
    Call gpHabSalvar
    
    If flxGridCoberturaDemais(Index).Tag = "Travado" Then
        flxGridCoberturaDemais(Index).Col = 0
        flxGridCoberturaDemais(Index).ColSel = flxGridCoberturaDemais(Index).Cols - 1
        Exit Sub
    End If
    
    With flxGridCoberturaDemais(Index)
    ColAnt = .Cols
    .Redraw = False
        Select Case KeyCode
            Case 46
                .Row = .RowSel
                .Col = .ColSel
                .Refresh
                .TextMatrix(.RowSel, .ColSel) = ""
                Call mpAtualTamanho
                EstaAlterando = True
                '.SetFocus
        End Select
     .Cols = ColAnt
    .Redraw = True
        
    End With
End Sub

Private Sub mpAtualTamanho()

'Rotina vai avaliar conte�do dos campos da grid do Grid de projeto e
'ajustar o tamanho da linha usando um campo de lbl para determinar o tamanho

Dim Tamanho As Long
Dim i As Integer
Dim ColAnt As Integer
Dim Rowant As Integer
Dim lI  As Integer
    
    lI = Val(TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag)
    
Static ProcessaInterno As Boolean
    If ProcessaInterno Then Exit Sub
    ProcessaInterno = True
    flxGridCoberturaDemais(lI).Redraw = False
    
    ColAnt = flxGridCoberturaDemais(lI).Col
    Rowant = flxGridCoberturaDemais(lI).Row
    Tamanho = CmbCoberturas.Height
    If mLinhaFocus >= flxGridCoberturaDemais(lI).Rows Then
        mLinhaFocus = flxGridCoberturaDemais(lI).Rows - 1
    End If
    flxGridCoberturaDemais(lI).Row = mLinhaFocus
    For i = 0 To 4
        flxGridCoberturaDemais(lI).Col = i
        If flxGridCoberturaDemais(lI).CellWidth > 80 Then
            LblTamanho.Caption = flxGridCoberturaDemais(lI).Text
            LblTamanho.Width = flxGridCoberturaDemais(lI).CellWidth - 80
        
            If LblTamanho.Height > Tamanho Then
                Tamanho = LblTamanho.Height + 50
            End If
        End If
    Next
    flxGridCoberturaDemais(lI).RowHeight(flxGridCoberturaDemais(lI).Row) = Tamanho
    flxGridCoberturaDemais(lI).Col = ColAnt
    flxGridCoberturaDemais(lI).Row = Rowant
    flxGridCoberturaDemais(lI).Redraw = True
    ProcessaInterno = False
End Sub

Private Sub mnuArqSaiItem_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Dim lMensagem           As String           'Mensagem.
    Dim objTabUser      As clsA46V725A      'Objeto da classe clsA28V713A (Tab_User)
    Dim objstcTabUser   As stcA46V725B      'Objeto da classe de estrutura stcA28V713B (Tab_User)

    '1. Centraliza form.
    Call gM46V999.gpCenForm(Me)
    
    lMensagem = "Verifica usuario Yasuda " & gM46V999.gNomUsuario

    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        Set objTabUser = New clsA46V725A
        Set objTabUser.M46V999 = gM46V999
        
        Set objstcTabUser = New stcA46V725B
        
        objstcTabUser.login_user_rede = UCase(gM46V999.gNomUsuario)
        If Not objTabUser.mfCarregarColecao(objstcTabUser) Then
            lMensagem = objTabUser.UltimoErro
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
        If gM46V999.gFinalizarPrograma = True Then End
    End If

    'Comentar para testar
    txtNumPI.Text = Mid(PegaParam(1), 2, 10)
    txtDataRecebimento.Text = Mid(PegaParam(1), 13, 8)
    lVerTipEmissao = Mid(PegaParam(1), 22, 1)
    
'    'Descomentar para testar
'    lVerTipEmissao = "S"
'    txtNumPI.Text = 1229000005
'    txtDataRecebimento = 20121204

'    'Descomentar para testar
'    lVerTipEmissao = "R"
'    txtNumPI.Text = 1120036849
'    txtDataRecebimento = 20120201
    
'    'Descomentar para testar
    'lVerTipEmissao = "E"
    'txtNumPI.Text = 1120036391
    'txtDataRecebimento = 20111214
    
'    'Descomentar para testar
'Simony
    'lVerTipEmissao = "S"
    'txtNumPI.Text = "1320900001"
    'txtDataRecebimento = 20130515
    
    
'    gM46V999.PerfilAcesso = e_Perfil_DIG_IST
'    gM46V999.gTipoAcesso = DIGITACAO
'
        
    Call mpCarregarRamos

    '2. Prepara grade de coberturas.
    '1200077 - SyasRD Ajuste Perfil Servi�os Operacionais
'    If gM46V999.PerfilAcesso = e_Perfil_DIG_IST Or gM46V999.PerfilAcesso = e_Perfil_DIG_Yas Then panItemLocRisco.Enabled = False
    
    Call mpCarregarGridCoberturaMod(0)
    Call mpCarregarCoberturas
    Call mpControleBotoes(False)
    Call panConfirmar_Click
    If Mid(cboRamo.Text, 1, 3) = "710" Then
        txtRisNomLogradouro.Text = "CONFORME ESPECIFICACAO EM ANEXO"
    End If
    
End Sub
Private Sub mpCarregarGridCoberturaMod(indice As Integer)
        '2. Prepara grade de coberturas.
    With flxGridCoberturaDemais(indice)
        'a) Quantidade inicial de colunas e linhas.
        .Cols = 13
        .Rows = 1
        .TextMatrix(0, 0) = "C�digo"
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "Cobertura"
'        .ColWidth(1) = 3405
        .ColWidth(1) = 4500
        .TextMatrix(0, 2) = IIf(gM46V999.gCod_Ramo = e_CodRamo_Resp_Civil, "Lim.Max.Indeniza��o", "Lim.Max.Garantia")
'        .ColWidth(2) = 1350
        .ColWidth(2) = 1500
        .TextMatrix(0, 3) = "Valor do Risco"
        .ColWidth(3) = 0
        .TextMatrix(0, 4) = "TAB_COBERT"
        .ColWidth(4) = 0
        .TextMatrix(0, 5) = "%Franq"
        .ColWidth(5) = IIf(gM46V999.gCod_Ramo = e_CodRamo_Resp_Civil, 1500, 0)
        .TextMatrix(0, 6) = "$ Franq"
        .ColWidth(6) = .ColWidth(5)
        .TextMatrix(0, 7) = "Desconto comercial"
        .ColWidth(7) = 0
        .TextMatrix(0, 8) = "Desconto experi�ncia"
        .ColWidth(8) = 0
        .TextMatrix(0, 9) = "Pr�mio l�quido"
'        .ColWidth(9) = 1095
        .ColWidth(9) = 1200
        .TextMatrix(0, 10) = "%P.O.S"
        .ColWidth(10) = IIf(gM46V999.gCod_Ramo = e_CodRamo_Resp_Civil, 1500, 0)
        .TextMatrix(0, 11) = "$P.O.S"
        .ColWidth(11) = .ColWidth(10)
        .TextMatrix(0, 12) = "$Dedut�vel"
        .ColWidth(12) = IIf(gM46V999.gCod_Ramo = e_CodRamo_Resp_Civil, 1500, 0)
        .Col = 0
        .Rows = 3
        .FixedRows = 1
        .TopRow = 1
        .Row = 1
        .RowHeight(1) = CmbCoberturas.Height
        .RowHeight(2) = CmbCoberturas.Height
    End With
End Sub
Private Sub mpCarregarCoberturas(Optional pLimparGrid As Boolean = True)

    Dim lMensagem       As String
    Dim lRsAux          As clsYasRecordSet
    Dim lBDConexao      As ADODB.Connection
    Dim lI              As Integer
    Dim lSelect         As String
       
    'lI = TabMod.SelectedItem.Index - 1
    lI = Val(TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag)
    If cboRamo.ListIndex <> -1 And (cboPlano(0).ListIndex <> -1 Or cboPlano(1).ListIndex <> -1) Then
        If pLimparGrid Then
            flxGridCoberturaDemais(lI).Rows = 1
            flxGridCoberturaDemais(lI).Rows = 2
            flxGridCoberturaDemais(lI).FixedRows = 1
        End If
        gM46V999.gBaseLocal = False
        If Not gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lBDConexao, lMensagem) Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If
        'gM46V999.gBaseLocal = True
        lSelect = "SELECT  Cod_Cobert, Tab_Cobert, min(Dsc_Cobert)  as Descricao from Tab_Cobert "
        lSelect = lSelect & " where cod_ramo = " & cboRamo.ItemData(cboRamo.ListIndex)
        
        If cboPlano(0).ListIndex <> -1 Then
            lSelect = lSelect & " and tab_cobert  = " & cboPlano(0).ItemData(cboPlano(0).ListIndex)
        Else
            If flxGridCoberturaDemais(lI).Tag <> "" Then
                lSelect = lSelect & " and tab_cobert  = " & flxGridCoberturaDemais(lI).Tag  'cboPlano(1).ItemData(cboPlano(1).ListIndex)
            Else
                lSelect = lSelect & " and tab_cobert  = " & cboPlano(1).ItemData(cboPlano(1).ListIndex)
            End If
        End If

        lSelect = lSelect & " and COD_TARIF_INI <= " & gM46V999.gCodProduto
        lSelect = lSelect & " and COD_TARIF_TER >= " & gM46V999.gCodProduto
        lSelect = lSelect & " group by Cod_Cobert, Tab_Cobert "
        lSelect = lSelect & " order by min(Dsc_Cobert)"
        If gM46V999.gfObtRegistro(lBDConexao, lSelect, lRsAux, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
        
        CmbCoberturas.Clear
        CmbCoberturas.Tag = 0
        
        If lRsAux.EOF = True Then
            Call gM46V999.gpGraLog(1, "ERR0055 - Tabela de Cobertura n�o localizada.")
            Exit Sub
        End If
        
        Do While Not lRsAux.EOF
            CmbCoberturas.AddItem gM46V999.Check(lRsAux!Descricao)
            CmbCoberturas.ItemData(CmbCoberturas.ListCount - 1) = gM46V999.Check(lRsAux!Cod_Cobert)
            If gM46V999.Check(lRsAux!Cod_Cobert) = 1 Then
                CmbCoberturas.Tag = CmbCoberturas.ListCount - 1
            End If
            lRsAux.MoveNext
        Loop
        
        lI = Val(TabMod.Tabs.Item(TabMod.SelectedItem.Index).Tag)
        If CmbCoberturas.Tag >= 0 Then
            On Error Resume Next
            CmbCoberturas.ListIndex = CmbCoberturas.Tag
            CmbCoberturas.ZOrder 0
            CmbCoberturas.Visible = True
            Call CmbCoberturas_LostFocus
            CmbCoberturas.Visible = False
        End If
    End If
End Sub

Private Sub txtRisCidade_Change()
    Call gpLimPremios
End Sub

Private Sub txtRisCidade_GotFocus()
    With txtRisCidade
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRisCidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtRisComplemento_Change()
    Call gpLimPremios
End Sub

Private Sub txtRisComplemento_GotFocus()
    With txtRisComplemento
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRisComplemento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtRisNomLogradouro_Change()
    Call gpLimPremios
End Sub

Private Sub txtRisNomLogradouro_GotFocus()
    With txtRisNomLogradouro
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRisNomLogradouro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtRisNumero_Change()
    Call gpLimPremios
End Sub
Private Sub txtRisNumero_GotFocus()
    With txtRisNumero
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRisNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtRisTipLogradouro_Change()
    Call gpLimPremios
End Sub

Private Sub txtRisTipLogradouro_GotFocus()
    With txtRisTipLogradouro
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRisTipLogradouro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSegurado_Change()
    Call gpLimPremios
    frmT47V004A.txtSegurado.Text = txtSegurado.Text
    
End Sub

Private Sub txtSegurado_GotFocus()
    With txtSegurado
        .Text = Replace(.Text, "-", "")
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


Private Sub TxtValor2_Change()
    Call gpHabSalvar
End Sub

Private Sub TxtValor2_GotFocus()
    With TxtValor2
'        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub TxtValor2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtValor2_LostFocus()
    TxtValor2.Text = Format(TxtValor2.Text, "standard")
End Sub

Private Sub TxtValor3_Change()
    Call gpLimPremios
End Sub

Private Sub TxtValor3_GotFocus()
    With TxtValor3
'        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub TxtValor3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtValor3_LostFocus()
    TxtValor3.Text = Format(TxtValor3.Text, "standard")
End Sub

Private Sub TxtValor4_Change()
    Call gpLimPremios
End Sub

Private Sub TxtValor4_GotFocus()
    With TxtValor4
'        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub TxtValor4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtValor4_LostFocus()
    TxtValor4.Text = Format(TxtValor4.Text, "standard")
End Sub

Private Sub TxtValor5_Change()
    Call gpLimPremios
End Sub

Private Sub TxtValor5_GotFocus()
    With TxtValor5
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub TxtValor5_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtValor5_LostFocus()
    TxtValor5.Text = Format(TxtValor5.Text, "standard")
End Sub

Private Sub TxtValor6_Change()
    Call gpLimPremios
End Sub

Private Sub TxtValor6_GotFocus()
    With TxtValor6
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub TxtValor6_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtValor6_LostFocus()
    TxtValor6.Text = gM46V999.gfForData(9, TxtValor6.Text)
End Sub

Private Sub TxtValor7_Change()
    Call gpLimPremios
End Sub

Private Sub TxtValor7_GotFocus()
    With TxtValor7
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub TxtValor7_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtValor7_LostFocus()
    TxtValor7.Text = Format(TxtValor7.Text, "standard")
End Sub

Private Sub txtValoRisco_Change()
    Call gpLimPremios
End Sub

Private Sub txtValoRisco_GotFocus()
    With txtValoRisco
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtValoRisco_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtValoRisco_LostFocus()
    txtValoRisco.Text = Format(txtValoRisco.Text, "standard")
    If CDbl(Val(txtValoRisco.Text)) <> 0 Then
        txtCoefISVR.Text = CDbl(Val(txtISBas.Text)) / CDbl(Val(txtValoRisco.Text))
        If Val(txtCoefISVR.Text) >= 10 Then
            Call gM46V999.gpGraLog(1, "FGE0111 - Coeficiente LMG/VR inv�lido.(maior que 9,99)")
            If txtValoRisco.Visible And txtValoRisco.Enabled Then
                txtValoRisco.SetFocus
            End If
        End If
    Else
        txtCoefISVR.Text = "0,00"
        txtValoRisco.Text = "0,00"
    End If
End Sub

Private Sub txtVigInicio_Change()
    Call gpLimPremios
End Sub

Private Sub txtVigInicio_GotFocus()
    With txtVigInicio
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtVigInicio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtVigInicio_LostFocus()
Dim lPrazo  As String
    
    txtVigInicio.Text = gM46V999.gfForData(9, txtVigInicio.Text)
    If gM46V999.gfPreenchido(txtVigTermino.Text) = False And gM46V999.gfPreenchido(txtVigInicio.Text) = True Then
        txtVigTermino.Text = Format(DateAdd("yyyy", 1, CDate(txtVigInicio.Text)), "dd/mm/yyyy")
    End If
    Call mpPosPrazo(txtVigInicio.Text, txtVigTermino.Text, lPrazo)
    lblPrazo.Caption = lPrazo
End Sub
Public Sub mpPosPrazo(ByVal pInicio As String, ByVal pTermino As String, ByRef pPrazo As String)
    'Procedure: posiciona prazo.
    '********************Vari�veis padr�o*************************************
    Dim lCkPoint        As String      'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lCkStr          As String       'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
                                        'lCkPoint. (posi��o das variaveis em ordem de declara��o)
    '********************Vari�veis Declaradas para o processamento************
    Dim lAuxQuaMeses    As Double   'Auxiliar para c�lculo.
    Dim lDatInicio      As Double   'Data de in�cio.
    Dim lDatTermino     As Double   'Data de t�rmino.
    Dim lDesPrazo       As Boolean  'Indicador para desabilitar prazo.
    Dim lQuaDias        As Double   'Quantidade de dias.
    Dim lQuaMeses       As Double   'Quantidade de meses.


    On Error GoTo ControleErro
    Err = 0
    lCkPoint = ""
    lCkStr = ""
    
    '1. Se data inv�lida, encerra processo.
    lCkPoint = "1"
    lDesPrazo = False
    If Not gM46V999.gfPreenchido(pInicio) Then
        lDesPrazo = True
    Else
        If Not gM46V999.gfPreenchido(pTermino) Then
            lDesPrazo = True
        Else
            If gM46V999.gfVerData(2, pInicio) = False Then
                lDesPrazo = True
            Else
                If gM46V999.gfVerData(2, pTermino) = False Then
                    lDesPrazo = True
                Else
                    lDatInicio = gM46V999.gfForData(1, pInicio)
                    lDatTermino = gM46V999.gfForData(1, pTermino)
                    If lDatInicio > lDatTermino Then
                        lDesPrazo = True
                    End If
                End If
            End If
        End If
    End If
    lCkPoint = "2"
    If lDesPrazo = True Then        'Desabilita prazo.
        pPrazo = " "
        frmT47V003A.optTipPrazo(0) = False
        frmT47V003A.optTipPrazo(1) = False
        frmT47V003A.optTipPrazo(2) = False
        frmT47V003A.optTipPrazo(0).Enabled = False
        frmT47V003A.optTipPrazo(1).Enabled = False
        frmT47V003A.optTipPrazo(2).Enabled = False
        Exit Sub
    End If
    lCkPoint = "3"
    '2. Calcula prazo de vig�ncia:
    'a) Dias.
    lQuaDias = gM46V999.gfDifDatas(lDatInicio, lDatTermino)
    If lQuaDias = 366 And Mid(lDatInicio, 5, 4) = Mid(lDatTermino, 5, 4) Then
        lQuaDias = 365
    End If
    lCkPoint = "4"
    'b) Meses.
    lAuxQuaMeses = lQuaDias / 30
    lQuaMeses = Int(lAuxQuaMeses)
    lAuxQuaMeses = lAuxQuaMeses - lQuaMeses
    If lAuxQuaMeses > 0.5 Then
        lQuaMeses = lQuaMeses + 1
    End If
    lCkPoint = "5"
    '3. Posiciona informa��es.
    If lQuaDias <= 365 Then
        Select Case lQuaDias
            Case 0
                pPrazo = "0"
            Case 1
                pPrazo = "1 dia"
            Case 366
                pPrazo = "365 dias"
            Case Else
                pPrazo = lQuaDias & " dias"
        End Select
        frmT47V003A.optTipPrazo(0).Enabled = True
        frmT47V003A.optTipPrazo(1).Enabled = True
        'Alterado Simony 22/08/2007
        'If lQuaDias < 365 Then
        '    frmT47V003A.optTipPrazo(0).Value = True     'Prazo curto
        'Else
        '    frmT47V003A.optTipPrazo(1).Value = True     'Pro rata
        'End If
        If frmT47V003A.optTipPrazo(2).Value = True Then
            frmT47V003A.optTipPrazo(2).Value = False
            frmT47V003A.optTipPrazo(0).Value = True
        End If
        frmT47V003A.optTipPrazo(2).Enabled = False
        If frmT47V003A.optTipPrazo(0).Value = False And frmT47V003A.optTipPrazo(1).Value = False Then
            frmT47V003A.optTipPrazo(1).Value = True
        End If
    Else
        Select Case lQuaMeses
            Case 0
                pPrazo = "0"
            Case 1
                pPrazo = "1 m�s"
            Case Else
                pPrazo = lQuaMeses & " meses"
        End Select
        frmT47V003A.optTipPrazo(2).Enabled = True
        frmT47V003A.optTipPrazo(2).Value = True 'longo
        frmT47V003A.optTipPrazo(0).Enabled = False
        frmT47V003A.optTipPrazo(1).Enabled = False
    End If
    lCkPoint = "6"
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    'Obtem o inicio da mensagem de erro
    lCkStr = "mpPosPrazo ERRO - " & Err & " " & Error & _
                Chr(13) & "CheckPoint " & lCkPoint & " - "
    'Carrega a variavel lCkStr
    On Error Resume Next
    lCkStr = lCkStr & " | " & lAuxQuaMeses
    lCkStr = lCkStr & " | " & lDatInicio
    lCkStr = lCkStr & " | " & lDatTermino
    lCkStr = lCkStr & " | " & lDesPrazo
    lCkStr = lCkStr & " | " & lQuaDias
    lCkStr = lCkStr & " | " & lQuaMeses
    On Error GoTo 0
    
    Call gM46V999.gpGraLog(0, lCkStr)
End Sub


Private Sub txtVigTermino_Change()
    Call gpLimPremios
End Sub

Private Sub txtVigTermino_GotFocus()
    With txtVigTermino
        .Text = Replace(.Text, "-", "")
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtVigTermino_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtVigTermino_LostFocus()
Dim lPrazo As String
    txtVigTermino.Text = gM46V999.gfForData(9, txtVigTermino.Text)
    Call mpPosPrazo(txtVigInicio.Text, txtVigTermino.Text, lPrazo)
    lblPrazo.Caption = lPrazo
End Sub
Private Sub optBens_Click(Index As Integer)
Dim lMensagem As String
    '1. Posiciona controles.
    Select Case Index
        Case 0  'Sim.
            cmdBens.Enabled = True
            cmdBens.Picture = cmdExtra(1).Picture
        Case 1  'N�o.
            cmdBens.Enabled = False
            cmdBens.Picture = cmdExtra(2).Picture
    End Select

    '2. Se a��o do programa ou processo de carga, encerra processo.
    If gCarga = True Then
        Exit Sub
    End If
    

    '4. Processa op��o do usu�rio.
    Select Case Index
        Case 0  'Sim.
            Call cmdBens_Click
        Case 1  'N�o.
            lMensagem = "AGE0016 - " & gM46V999.gPrefixo & "Exclu�da op��o de bens. " & _
                        "Lista de bens ser� desconsiderada. Confirma esta op��o?"
            If MsgBox(lMensagem, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call gM46V999.gpGraLog(0, lMensagem & " Sim.")
                On Error Resume Next
                Unload frmT47V018A
                On Error GoTo 0
                Exit Sub
            End If
            Call gM46V999.gpGraLog(0, lMensagem & " N�o.")
            gCarga = True
            optBens(0).Value = True
            If optBens(0).Visible And optBens(0).Enabled Then
                optBens(0).SetFocus
            End If
            gCarga = True
    End Select
End Sub

Private Sub optBens_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub cmdBens_Click()
    Call mnuExiBem_Click
End Sub
Private Sub mnuExiBem_Click()
    If gM46V111.gNum_Pi > 0 Then
        If optBens(0).Value = False Then
            optBens(0).Value = True
        Else
            frmT47V018A.Show vbModal, Me
        End If
    End If
End Sub

Private Sub mnuArqTransmitir_Click()
Dim lMsgErr As String

    If mfTransmitir(False, lMsgErr) = False Then
        Call gM46V999.gpGraLog(2, lMsgErr, vbCritical)
    End If
End Sub

Private Function GravaInterfaceSAP(ByRef pobjPedido As clsA46V702A, ByRef pobjstcPedido As stcA46V702B, ByRef pMensagem As String) As Boolean
    Dim lrsTab_Cobert As clsYasRecordSet
    Dim lSql As String
    Dim lCodRamoSec As Integer
    Dim lNumApolEndosso As Double
    Dim lCodRetorno As Integer
    Dim lDescricao As String

    On Error GoTo ControleErro
    
    Call gM46V999.gpGraLog(0, "GravaInterfaceSAP - Num_pi: " & pobjstcPedido.NUM_PI)
    
    GravaInterfaceSAP = False
    lCodRamoSec = 0

    lSql = "SELECT TOP 1  COD_RAMO_SEC  FROM RamosDiversos..Tab_Ped_Cobert AS  F with(nolock) "
    lSql = lSql & " INNER JOIN RamosDiversos..Tab_Ped  AS P  with(nolock)  ON P.NOSSO_NUMERO = F.NOSSO_NUMERO"
    lSql = lSql & " INNER JOIN RamosDiversos..Tab_Cobert AS C with(nolock) ON  C.COD_RAMO = F.COD_RAMO AND C.TAB_COBERT = COD_ESTIP"
    lSql = lSql & " AND  COD_PRODUTO = COD_TARIF_INI AND C.COD_COBERT = F.COD_COBERT"
    lSql = lSql & " WHERE F.NOSSO_NUMERO = " & gM46V999.gfFormatarCamposSQL(pobjstcPedido.NOSSO_NUMERO) & " and COD_RAMO_SEC <> p.Cod_Ramo"
     
    If Not gM46V999.gfObtRegistro(pobjPedido.ConexaoBD, lSql, lrsTab_Cobert, pMensagem) Then
        Call gM46V999.gpGraLog(1, pMensagem)
        Exit Function
    End If
    
    If Not lrsTab_Cobert.EOF Then
        lCodRamoSec = lrsTab_Cobert("COD_RAMO_SEC")
    End If
    Set lrsTab_Cobert = Nothing
    
    If lCodRamoSec = 0 Then
        lCodRamoSec = pobjstcPedido.Cod_Ramo
    End If
    
    If pobjstcPedido.Cod_Dizer <> 0 Then
        lNumApolEndosso = pobjstcPedido.Num_Apol_An
    End If

'    @NUM_CGC_CPF    bigint
'    @NUM_PI bigint
'    @NUM_LOGR_COBR  int
'    @NUM_LOGR_RESID int
'    @DSC_COMPL_LOGR_COBR_SN varchar
'    @DSC_COMPL_LOGR_RESID_SN    varchar
'    @COD_RAMO   smallint
'    @COD_PROD   int
'    @COD_RAMORAMO   smallint
'    @NUM_APOL_ENDO  bigint
'    @CDRET  tinyint
'    @DSC_RET    varchar
    
    
    lSql = "SET NOCOUNT ON BEGIN"
    lSql = lSql & " DECLARE @CDRET tinyint,"
    lSql = lSql & "         @DSC_RET varchar(max)"
    lSql = lSql & "    exec P0044800.dbo.sp_InterfaceCadastro '" & Trim(pobjstcPedido.Num_Cgc_Cpf) & "'"
    lSql = lSql & ", " & pobjstcPedido.NUM_PI
    lSql = lSql & ", " & gM46V999.gfTratarNulo(Val(pobjstcPedido.objstcCtrlEmis.Num_Logr_Cobranca), tpoNumerico)
    lSql = lSql & ", " & gM46V999.gfTratarNulo(Val(pobjstcPedido.objstcCtrlEmis.Num_Logr_Segurado), tpoNumerico)
    lSql = lSql & ", " & gM46V999.gfFormatarCamposSQL(pobjstcPedido.objstcCtrlEmis.Dsc_Compl_Cobranca)
    lSql = lSql & ", " & gM46V999.gfFormatarCamposSQL(pobjstcPedido.objstcCtrlEmis.Dsc_Compl_Segurado)
    lSql = lSql & ", " & gM46V999.gfTratarNulo(pobjstcPedido.Cod_Ramo, tpoNumerico)
    lSql = lSql & ", " & gM46V999.gfTratarNulo(pobjstcPedido.Cod_Estip, tpoNumerico)
    lSql = lSql & ", " & gM46V999.gfTratarNulo(lCodRamoSec, tpoNumerico)
    lSql = lSql & ", " & gM46V999.gfTratarNulo(lNumApolEndosso, tpoNumerico)
    lSql = lSql & ",  @CDRET OUTPUT "
    lSql = lSql & ",  @DSC_RET OUTPUT "
    lSql = lSql & " SELECT @CDRET as CDRET, @DSC_RET as DSC_RET"
    lSql = lSql & " END;"
    
    Call gM46V999.gpGraLog(0, "GravaInterfaceSAP: " & lSql)
    
'    Call gM46V999.gpGraLog(0, lSql)
    Dim bdGed As ADODB.Connection
    If Not gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, bdGed, pMensagem) Then Call gM46V999.gpGraLog(1, pMensagem)
    
    If Not gM46V999.gfObtRegistro(bdGed, lSql, lrsTab_Cobert, pMensagem) Then
        Call gM46V999.gpGraLog(1, pMensagem)
        Exit Function
    End If

    Call gM46V999.gpGraLog(0, "GravaInterfaceSAP: " & lrsTab_Cobert("CDRET") & " - " & lrsTab_Cobert("DSC_RET"))

    If lrsTab_Cobert("CDRET") <> 0 Then
        pMensagem = lrsTab_Cobert("CDRET") & " - " & lrsTab_Cobert("DSC_RET")
        Call gM46V999.gpGraLog(1, pMensagem)
        Exit Function
    End If

    GravaInterfaceSAP = True
    
    Set lrsTab_Cobert = Nothing
    bdGed.Close
    Set bdGed = Nothing
    
    Exit Function
    
ControleErro:
    Call gM46V999.gpGraLog(1, "GravaInterfaceSAP: " & Err & " " & Err.Description, vbCritical)
End Function

Private Function mfTransmitir(ByVal pAutomatico As Boolean, _
                              ByRef pMsgErr As String, Optional pGerarAutEspec As Boolean = False) As Boolean
Dim lArqIni         As String           'Endere�o do arquivo SYASPROD.INI.
Dim lArqTransmissao As String           'Endere�o do arquivo de transmiss�o.
Dim lCodCorretor    As Long             'C�digo do corretor.
Dim lMensagem       As String           'Mensagem.
Dim lTipDocumento   As Byte             'Tipo do documento.
Dim lUsuario        As String 'Long     'C�digo do usu�rio de transmiss�o.

Dim lTipoTransmissao    As Long     'Tipo de transmiss�o
Dim lResultado          As Integer  'Resultado da transmiss�o
Dim lOK                 As String

Dim lAcao               As Integer  'Fluxo
Dim lRodar              As Boolean


Err.Clear
mfTransmitir = False
On Error GoTo TrataErro

    Screen.MousePointer = vbHourglass
    lOK = ""
    If pAutomatico = True Then
    'Projeto 9075 **-- Transmissao automatica somente no perfil de corretor
        Call gM46V999.gpGraLog(0, "Transmissao automatica somente no perfil de corretor")
        mfTransmitir = True
        GoTo SAIDA
    End If
    If frmT47V004A.TxtCorretor(1).Text = "" Then
        lMensagem = "FGE0044 - Corretor inv�lido (primeiro corretor n�o preenchido)."
        GoTo TrataErro
    End If
    'Salva dados.
    If barFerramentas.Buttons("Salvar").Enabled = True Then
        If mfSalvar() = False Then
            lMensagem = "FGE0044 - Problemas ao salvar dados."
            GoTo TrataErro
        End If
'                gAltComum = False
    End If
    
    If gM46V999.gPendenteWorkflow = True Then
        lMensagem = "Necess�rio que a tarefa do workflow seja liberada."
        GoTo TrataErro
    End If
    
    'Se for passo de digita��o, verifico se existem erros apontados na guia de triagem.
    
'            If mfExisteErrosTriagem Then
'                gM46V999.gEfetuouCalculo = False
'            End If
    
    Call frmT47V020A.mpInstanciarForm(lTipoTransmissao, lAcao)
    
    mFlag_Emitir = False
    mFlag_Suspender = False

    If lTipoTransmissao > 0 Then
        Screen.MousePointer = vbHourglass
        'Call gpTransmitirPropostaOCX(lTipoTransmissao, lResultado, objPedido)
        lRodar = True
        
        Select Case lTipoTransmissao
            Case 1
'                objstcPedido.Cod_Posi_Atual = e_IdeEmitida_Emitir
                objstcPedido.Cod_Sit_Ped = 1    'Consistido
                objstcPedido.Cod_Posi_Atual = "T"
                lOK = "transmitido"
                mFlag_Emitir = True
                
'                'Interface Simony 03/2015
'                pMsgErr = Empty
'                If Not GravaInterfaceSAP(objPedido, objstcPedido, pMsgErr) Then
'                    Call MsgBox(pMsgErr)
'                    Screen.MousePointer = vbDefault
'                    Call gM46V999.gpGraLog(0, "mfTransmitir: NUM_PI - " & objstcPedido.objstcCtrlEmis.NUM_PI & " - " & pMsgErr)
'                    Exit Function
'                End If
                
            Case 2
                objstcPedido.Cod_Posi_Atual = e_IdeEmitida_Suspender
                objstcPedido.Cod_Sit_Ped = 2    'Suspenso
                lOK = "suspenso"
                lRodar = False
                mFlag_Suspender = True
            Case 3
                objstcPedido.Cod_Posi_Atual = e_IdeEmitida_EmitirProd
'                lOK = "emitido produto"
                lOK = "suspenso produto"
                mFlag_Emitir = False
                mFlag_Suspender = False
            Case 5
                objstcPedido.Cod_Posi_Atual = e_IdeEmitida_SuspenderProd
                objstcPedido.Cod_Sit_Ped = 2    'Suspenso
                lOK = "suspenso produto"
                mFlag_Emitir = False
                mFlag_Suspender = False
            Case 6
                objstcPedido.Cod_Posi_Atual = e_IdeEmitida_Reemissao
                lOK = "reemiss�o"
                mFlag_Emitir = False
                mFlag_Suspender = False
            Case Else
                objstcPedido.Cod_Posi_Atual = e_IdeEmitida_Nao
                lOK = "recusado"
                mFlag_Emitir = False
                mFlag_Suspender = False
        End Select
        
        'Interface Toshio 31/08/2015
        Call gM46V999.gpGraLog(0, "mfTransmitir: NUM_PI - " & objstcPedido.objstcCtrlEmis.NUM_PI)
        pMsgErr = Empty
        If Not GravaInterfaceSAP(objPedido, objstcPedido, pMsgErr) Then
            Call MsgBox("mfTransmitir (GravaInterfaceSAP): " & pMsgErr)
            Screen.MousePointer = vbDefault
            Call gM46V999.gpGraLog(0, "mfTransmitir: NUM_PI - " & objstcPedido.objstcCtrlEmis.NUM_PI & " - " & pMsgErr)
            Exit Function
        End If
        
        objstcPedido.DAT_DIGIT = Left(gM46V999.gfForData(3), 8)
        If Not objPedido.mfAlterar(objstcPedido, "", "", mComUltData, mComUltUsuario) Then
            lMensagem = objPedido.UltimoErro
            GoTo TrataErro
        End If
        
        objstcPedido.objstcCtrlEmis.Dat_Digit_MF = Left(gM46V999.gfForData(3), 8)
        
        If lRodar = False And lAcao <> 2 Then
            objstcPedido.objstcCtrlEmis.Cod_Iden_Stat_Acao = lAcao
        Else
            objstcPedido.objstcCtrlEmis.Cod_Iden_Stat_Acao = 0
        End If
        
        Dim lobjCtrlEmis As clsA46V712A
        Set lobjCtrlEmis = New clsA46V712A
        Set lobjCtrlEmis.M46V999 = gM46V999
        
        Set lobjCtrlEmis.ConexaoBD = objPedido.ConexaoBD
        If Not lobjCtrlEmis.mfAlterar(objstcPedido.objstcCtrlEmis, "", "") Then
            lMensagem = objPedido.UltimoErro
            GoTo TrataErro
        End If
                        
'                Call gfVerSegSiscota(objstcPedido)--------------ver
        Screen.MousePointer = vbDefault
        'Se ocorreu tudo OK finalizar o sistema
        If lOK <> "" Then
            MsgBox "Pi: " & objstcPedido.objstcCtrlEmis.NUM_PI & " ser� " & lOK & ".", vbInformation
            Call mnuArqSaiItem_Click
            mfTransmitir = True
            GoTo SAIDA
        End If
    End If

'************* Transmitido com sucesso
    mfTransmitir = True
'************* Transmitido com sucesso

SAIDA:
    '9. Encerra processo.
    GoSub FechaObj
    Screen.MousePointer = vbDefault
    Err.Clear
    On Error GoTo 0
    Exit Function
    
TrataErro:
    If gM46V999.gfPreenchido(lMensagem) = False Then
        lMensagem = Err & " - " & Error
    End If
    pMsgErr = "mfTransmitir Erro:" & lMensagem
    Call gM46V999.gpGraLog(0, "mfTransmitir: " & pMsgErr & "- Num_PI: " & objstcPedido.objstcCtrlEmis.NUM_PI)
    GoTo SAIDA
    
FechaObj:
    Return
End Function

Private Sub mnuEdiCalPremios_Click()
    panAguarde.Visible = True
    panAguarde.Caption = "Validando as informa��es..."
    panAguarde.Refresh
    
    Screen.MousePointer = vbHourglass
    
    If lVerTipEmissao = "E" Or lVerTipEmissao = "R" Then
        If Val(txtApolice.Text) = 0 Then
            Call gM46V999.gpGraLog(1, "N�mero da Ap�lice anterior n�o informado.")
            Exit Sub
        End If
    End If
    
    If gM46V111.gNum_Pi > 0 Then
        mnuArqSalvar.Enabled = True
        Call mfSalvar
        If mpValidar = True Then
            gM46V999.gEfetuouCalculo = True
            gM46V999.gPodeEmitir = True
        Else
            gM46V999.gEfetuouCalculo = False
            gM46V999.gPodeEmitir = False
        End If
    End If
    
    panAguarde.Visible = False
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub mpConsisteComum(pCod_Retorno As Integer)
 Dim lretorno As String
        
    lretorno = gM46V111.gfExecDLL(objstcPedido, objstcPedLoc, e_TipProc_ConsComum, 0, pCod_Retorno, gM46V999.gTipoHomologacao)

    If InStr(1, lretorno, "NUMERO DO ITEM DEVE SER PREENCHIDO") <> 0 Then
        lretorno = gM46V111.gfExecDLL(objstcPedido, objstcPedLoc, e_TipProc_ConsComum, 1, pCod_Retorno, gM46V999.gTipoHomologacao)
    End If
    
    If InStr(1, lretorno, "OPCAO INVALIDA") <> 0 Then Exit Sub
   
    If InStr(1, lretorno, "Existem erros de consist�ncia!") <> 0 Then
        'Busca os erros de consist�ncia.
        Call mpCapturaErros("Consist�ncia do comum")
    ElseIf gM46V999.gfPreenchido(lretorno) = True Then 'retornou sem erros
        Call gM46V999.gpMenErro("Consist�ncia do comum - " & lretorno)
    End If
    
    DoEvents
    
End Sub

Private Sub mpConsisteItem(pItem As Long, ByRef pCod_Retorno As Integer)
Dim lretorno As String
Dim lbdP00MPED      As ADODB.Connection 'Base de dados: P0042300.
Dim lRstErro        As clsYasRecordSet
Dim lMensagem       As String
Dim lSelect         As String

    'N�o consiste Itens dos ramos diferente de Multi
    If gM46V111.gTipoGrupMulti <> e_GrupoRamo_Multi Then Exit Sub
    
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, lbdP00MPED, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
        
    lSelect = " select count(*) as cobert from Tab_Ped_Cobert where nosso_numero = '" & objstcPedido.NOSSO_NUMERO & "' And Num_Item = " & pItem
    If gM46V999.gfObtRegistro(lbdP00MPED, lSelect, lRstErro, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    If Not lRstErro.EOF Then
        'Call gM46V999.gpGraLog(2, "Antes Consistir *** Tab_Cobert ==> " & lRstErro("cobert") & " *** ITEM " & pItem)
    End If
    
    lretorno = gM46V111.gfExecDLL(objstcPedido, objstcPedLoc, e_TipProc_ConsItem, pItem, pCod_Retorno, gM46V999.gTipoHomologacao)
    
    lSelect = " select count(*) as cobert from Tab_Ped_Cobert where nosso_numero = '" & objstcPedido.NOSSO_NUMERO & "' And Num_Item = " & pItem
    If gM46V999.gfObtRegistro(lbdP00MPED, lSelect, lRstErro, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    If Not lRstErro.EOF Then
        'Call gM46V999.gpGraLog(2, "Depois de Consistir *** Tab_Cobert ==> " & lRstErro("cobert") & " *** ITEM " & pItem)
    End If


    If InStr(1, lretorno, "Existem erros de consist�ncia!") <> 0 Then
        'Busca os erros de consist�ncia.
        Call mpCapturaErros("Consist�ncia do item " & pItem)
        pCod_Retorno = 0
    ElseIf gM46V999.gfPreenchido(lretorno) = True Then 'retornou sem erros
        Call gM46V999.gpMenErro("Item " & Format(pItem, "00000") & "  " & lretorno)
    End If
    
    DoEvents
End Sub


'Sandra - COBOL
Private Sub mpConsisteDadosComum(pCod_Retorno As Integer)
 Dim lretorno As String

    If lVerTipEmissao = "E" Then
        lretorno = gfPesqDadosComum(objstcPedido, objstcPedLoc, e_TipProc_ConsComum, 0, pCod_Retorno, gM46V999.gTipoHomologacao)
    Else
        lretorno = gfPesqDadosComum(objstcPedido, objstcPedLoc, e_TipProc_ConsComum, 1, pCod_Retorno, gM46V999.gTipoHomologacao)
    End If
        
    If pCod_Retorno > 0 Then frmT47V020A.optOpcao(0).Enabled = False
    
    Select Case pCod_Retorno
            Case 1
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "PROCESSO INTERNO SEM DADOS DE COMUM")
            Case 2
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "PROCESSO INTERNO INEXISTENTE")
            Case 3
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "PROCESSO INTERNO EMITIDO")
            Case 4
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "PROCESSO INTERNO CANCELADO")
            Case 5
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "APOLICE ANTERIOR DIFERE DA TRIADA")
            Case 6
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "APOLICE JA EMITIDA")
            Case 7
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "CODIGO DO RAMO NAO CONFERE COM A APOLICE")
            Case 8
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "COD SEGURADO NAO PODE SER SER ALTERNATIVO")
            Case 9
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "CODIGO DE REGISTRO NAO PERMITE CUSTO")
            Case 10
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "OBRIGATORIO CUSTO PARCELADO")
            Case 11
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "OBRIGATORIO CUSTO ZERADO")
            Case 12
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "TERMINO INFERIOR A INICIO DE VIGENCIA")
            Case 13
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "TERMINO IGUAL AO INICIO DE VIGENCIA")
            Case 14
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "REFERENCIA N�O PREENCHIDA")
            Case 15
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "DIGITAR CORRETOR-1 MARCH E CORRETOR-2 NOMMINAL")
            Case 16
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "PARA CORRETOR NOMMINAL DEVE EXISTIR CORRETOR MARCH")
            Case 17
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "DEVE SER DIGITADO O CODIGO DA SEGURATEC")
            Case 18
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "PARA BASCOL, DEVE SER DIGITADO O CODIGO DA SEGURATEC")
            Case 19
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "PARA PGTO.ANTEC. A VISTA, JURO DEVE SER ZEROS")
            Case 20
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "CORRETOR FOI DIGITADO MAIS DE UMA VEZ")
            Case 21
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "QDO CORRETOR FOR PTMM EH OBRIGATORIO DIG TMM")
            Case 22
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "QDO CORRETOR FOR TMM EH OBRIGATORIO DIG PTMM")
            Case 23
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "YASUDA NAO PODE SER COSSEGURADORA")
            Case 24
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "COSSEGURADORA 955 PERMITIDO SOMENTE PARA RAMO VIDA")
            Case 25
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "NAO E MAIS PERMITIDO COSSEGURO SEASUL")
            Case 26
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "TIPO DE ENVIO DEVE SER 'S' PARA SEGURADO E 'C' PARA CORRETOR")
            Case 27
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "TIPO DE COTACAO DEVE SER 0, 1, 2 OU 3")
            Case 28
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "NUMERO DA COTACAO DEVE SER NUMERICO")
            Case 29
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "CORRETAGEM VVD/PALLAS,DIGITAR PRIMEIRO PALLAS")
            Case 30
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "QUANDO FOR CORRETOR PALLAS UTILIZAR TAMBEM VVD")
            Case 31
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "LOCALIDADE DA PALLAS E VVD DEVE SER IGUAIS")
            Case 32
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "VVD DEVE TER 70 % DA COMISSAO")
            Case 33
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "PALLAS DEVE TER 30 % DA COMISSAO")
            Case 34
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "PAGAMENTO ANTECIPADO NAO CADASTRADO")
            Case 35
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "FOI INFORMADO DEBITO DA 1A PARCELA, POREM EXISTE PAGAMENTO ANTECIPADO CADASTRADO")
            Case 36
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "PREMIO ANTECIPADO DEVE SER MAIOR QUE CUSTO + IOF")
            Case 37
                Call gM46V999.gpMenErro("Consist�ncia do comum - " & "PARA ESTE CODIGO DE REGISTRO JURO NAO DEVE SER PREENCHIDO")

    End Select
  
    DoEvents
    
End Sub

'Sandra - COBOL
Private Sub mpConsisteDadosItem(pItem As Long, ByRef pCod_Retorno As Integer)
Dim lretorno As String
Dim lbdP00MPED      As ADODB.Connection 'Base de dados: P0042300.
Dim lRstErro        As clsYasRecordSet
Dim lMensagem       As String
Dim lSelect         As String
On Error GoTo ControleErro
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, lbdP00MPED, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
        
    lSelect = " select count(*) as cobert from Tab_Ped_Cobert where nosso_numero = '" & objstcPedido.NOSSO_NUMERO & "' And Num_Item = " & pItem
    If gM46V999.gfObtRegistro(lbdP00MPED, lSelect, lRstErro, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    If Not lRstErro.EOF Then
        'Call gM46V999.gpGraLog(2, "Antes Consistir *** Tab_Cobert ==> " & lRstErro("cobert") & " *** ITEM " & pItem)
    End If
    
    lretorno = gfPesqDadosItem(objstcPedido, objstcPedLoc, e_TipProc_ConsItem, pItem, pCod_Retorno, gM46V999.gTipoHomologacao)
    
    lSelect = " select count(*) as cobert from Tab_Ped_Cobert where nosso_numero = '" & objstcPedido.NOSSO_NUMERO & "' And Num_Item = " & pItem
    If gM46V999.gfObtRegistro(lbdP00MPED, lSelect, lRstErro, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    If Not lRstErro.EOF Then
        'Call gM46V999.gpGraLog(2, "Depois de Consistir *** Tab_Cobert ==> " & lRstErro("cobert") & " *** ITEM " & pItem)
    End If

    lretorno = gfPesqDadosItem(objstcPedido, objstcPedLoc, e_TipProc_ConsItem, pItem, pCod_Retorno, gM46V999.gTipoHomologacao)

    If pCod_Retorno > 0 Then frmT47V020A.optOpcao(0).Enabled = False
  
    Select Case pCod_Retorno
            Case 1
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "PROCESSO INTERNO SEM ITENS CADASTRADOS")
           Case 2
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "ITEM DA APOLICE JA EMITIDO")
            Case 3
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "INSPECAO 9999999 INVALIDA")
            Case 4
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "CODIGO DA CLASSIFICACAO NAO ESTA PREENCHIDO")
            Case 5
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "COMPLEMENTO DA CLASSIFICACAO NAO ESTA PREENCHIDO")
            Case 6
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "PERCENTUAL DE BONUS INVALIDO PARA TELEMAR")
            Case 7
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "PERCENTUAL DE DESCONTO PROGRESSIVO INVALIDO PARA TELEMAR")
            Case 8
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "CODIGO DE CONSTRUCAO INVALIDA PARA ESTA CLASSIFICACAO")
            Case 9
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "PERCENTUAL DESC. COMERCIAL DA COBERTURA 9999 INVALIDO PARA TELEMAR")
            Case 10
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "QUANTIDADE DE EMPREGADOS DEVE SER ZEROS")
            Case 11
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "INFORME A QUANTIDADE DE EMPREGADOS")
            Case 12
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "FALTA CLAUSULA 0001 OU 0002")
            Case 13
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "DIGITE SOMENTE CLAUSULA 0001 OU 0002")
            Case 14
                Call gM46V999.gpMenErro("Consist�ncia dos itens - " & "PARA CLASSIFICACAO 197 E 134, A CLAUSULA DEVE SER 0001")
    End Select
   
    DoEvents
    Exit Sub
    
ControleErro:
    MsgBox "Erro  mpConsisteDadosItem " & Err

End Sub

Private Sub mpCapturaErros(pPref As String)
Dim lBDP00Multi     As ADODB.Connection
Dim lRstErro        As clsYasRecordSet
Dim lMensagem       As String
Dim lSelect         As String
Dim lDelete         As String

    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    
    lSelect = " Select * from Tab_Erro"
    
    If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lRstErro, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
        
    Do While Not lRstErro.EOF
        If gM46V999.PerfilAcesso = e_Perfil_CALC_Gestor And gM46V111.gNum_Pi = 0 Then GoTo LBL_GRAVAR
        If gM46V999.PerfilAcesso <> e_Perfil_DIG_Gestor And gM46V999.PerfilAcesso <> e_Perfil_CALC_Gestor Then GoTo LBL_GRAVAR
   
GoTo LBL_FIM:
LBL_GRAVAR:
    Call gM46V999.gpMenErro(pPref & " - " & Trim(lRstErro("DSC_ERRO")))
LBL_FIM:
     lRstErro.MoveNext
    Loop
    
    lDelete = " delete from Tab_Erro"
    If gM46V999.gfExeSQL(lBDP00Multi, lDelete, lMensagem) <> 0 Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    
    Call gM46V999.gpFecha2(lBDP00Multi)
End Sub

Private Sub mpVerQtdeCoberturas(pobjstcItem As stcA46V708B, ByRef pModalidade As Double)
    'Procedure: consiste item n�vel 2: consiste todos dados do item.
        'Procedure: consiste item n�vel 2: consiste todos dados do item.
Dim lobjstcCobert   As stcA46V704B
Dim lDig            As String
Dim clsA99V001      As clsA99V001
Dim lMensagem       As String
Dim lAuxiliar       As String
Dim lNumCobert      As Integer

    lNumCobert = 0
    For Each lobjstcCobert In objstcPedLoc.ColecaoPedCobert
        If lobjstcCobert.Tab_Cobert = pModalidade Then
            lNumCobert = lNumCobert + 1
        End If
    Next
    If lNumCobert = 1 Then
        Call gM46V999.gpMenErro("AGE0053 - Obrigat�rio a contrata��o de pelo menos uma cobertura acess�ria para a Modalidade " & pModalidade & ".")
    End If
        
End Sub

Private Function mfSalvarTextoEspecPadrao(pMensagem As String) As Boolean
Dim lobjTextoEspec        As clsA46V726A
Dim lobjstcTextoEspec     As stcA46V726B
Dim ColecaoPedCobTemp     As Collection
Dim lobjstcCobert         As stcA46V704B
Dim lObjstcCorr           As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).
Dim lobjstcItem           As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)Dim lNumPI                As Long
Dim lTextoInc             As String
Dim lTotLin               As Long
Dim lSeq                  As Long
Dim AuxArray              As Variant
Dim i                     As Integer
Dim lNum_Pag              As Integer
Dim lCod_Espec            As Integer
Dim lLinha                As Integer
Dim lCorretor             As Double
Dim lCod_Clasf            As String
Dim lCod_Compl            As String
Dim lCod_Regiao           As Double
Dim lCod_Produto          As Double
Dim lMensagem             As String
Dim lRegiao               As Integer
Dim lAuxiliar     As String
Dim lAuxiliar1     As String
                    
On Error GoTo ControleErro
    
    lNum_Pag = 1
    
    Set lobjTextoEspec = New clsA46V726A
    Set lobjTextoEspec.M46V999 = gM46V999
    
    Set lobjstcTextoEspec = New stcA46V726B
    Set lobjTextoEspec.ConexaoBD = objPedido.ConexaoBD
    If objstcPedido.ColecaoTextoEspec.Count > 0 Then
        lobjstcTextoEspec.NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
        'Vai excluir tudo pois essa rotina vai salvar al�m
        ' dos benefici�rios as demais cl�usulas
        ' lobjstcTextoEspec.Cod_Espec = 8500 e
        ' as franquias das coberturas 56, 57 e 60
        If Not lobjTextoEspec.mfExcluir(lobjstcTextoEspec, "", "", "") Then
            pMensagem = lobjTextoEspec.UltimoErro
            GoTo Finalizar
        End If
    End If
    
    Set objstcPedido.ColecaoTextoEspec = New Collection
    
    lSeq = 0
    lTotLin = 1
    
    Set lobjstcTextoEspec = New stcA46V726B
   
SalvarEspecificacao:
   
    With frmT47V022A
  
      For lLinha = 1 To .flxEspec.Rows - 1
        lSeq = 0
        lTotLin = 1
              
        If gM46V999.CDblx(.flxEspec.TextMatrix(lLinha, 1)) <> 0 Then
             lCod_Espec = .flxEspec.TextMatrix(lLinha, 1)
             If InStr(.flxEspec.TextMatrix(lLinha, 2), vbCrLf) <> 0 Then
                AuxArray = Split(.flxEspec.TextMatrix(lLinha, 2), vbCrLf)
                lTextoInc = AuxArray(lSeq)
                lTotLin = UBound(AuxArray)
             Else
                 lSeq = 1
                 lTextoInc = .flxEspec.TextMatrix(lLinha, 2)
             End If
            
             Set lobjstcTextoEspec = New stcA46V726B
            
             With lobjstcTextoEspec
                 .NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
                 .Cod_Org = 0
                 .Cod_Ramo = gM46V999.gCod_Ramo
                 .Num_Ped = objstcPedido.Num_Ped
                 .Num_Apol = 0
                 .Num_Endo = 0
                 .Cod_Espec = lCod_Espec
                 .Num_pag = lNum_Pag
                 .Qtd_vias = 3
                 For i = 1 To lTotLin + 1
                     If i = 1 Then .Espec_lin1 = lTextoInc
                     If i = 2 Then
                        .Espec_lin2 = lTextoInc
                        If InStr(1, .Espec_lin1, "** FRANQUIA **", vbTextCompare) > 0 Or _
                           InStr(1, .Espec_lin2, "** FRANQUIA **", vbTextCompare) > 0 Then
                            GoTo Proximo
                        End If
                     End If
                     If i = 3 Then .Espec_lin3 = lTextoInc
                     If i = 4 Then .Espec_lin4 = lTextoInc
                     If i = 5 Then .Espec_lin5 = lTextoInc
                     If i = 6 Then .Espec_lin6 = lTextoInc
                     If i = 7 Then .Espec_lin7 = lTextoInc
                     If i = 8 Then .Espec_lin8 = lTextoInc
                     If i = 9 Then .Espec_lin9 = lTextoInc
                     If i = 10 Then .Espec_lin10 = lTextoInc
                     If i = 11 Then .Espec_lin11 = lTextoInc
                     If i = 12 Then .Espec_lin12 = lTextoInc
                     If i = 13 Then .Espec_lin13 = lTextoInc
                     If i = 14 Then .Espec_lin14 = lTextoInc
                     If i = 15 Then .Espec_lin15 = lTextoInc
                     If i = 16 Then .Espec_lin16 = lTextoInc
                     If i = 17 Then .Espec_lin17 = lTextoInc
                     If i = 18 Then .Espec_lin18 = lTextoInc
                     If i = 19 Then .Espec_lin19 = lTextoInc
                     If i = 20 Then .Espec_lin20 = lTextoInc
                     If lSeq < lTotLin And lSeq + 1 < 20 Then
                         lSeq = lSeq + 1
                         lTextoInc = AuxArray(lSeq)
                     Else
                         Exit For
                     End If
                 Next
                 .NUM_PI = gM46V111.gNum_Pi
             End With
             
             If Not lobjTextoEspec.mfIncluir(lobjstcTextoEspec, "", "", "") Then
                 pMensagem = lobjTextoEspec.UltimoErro
                 GoTo Finalizar
             End If
             
             objstcPedido.ColecaoTextoEspec.Add lobjstcTextoEspec, lobjstcTextoEspec.Class_ID
             lNum_Pag = lNum_Pag + 1
         End If
Proximo:
      Next
    End With

    mfSalvarTextoEspecPadrao = True
Finalizar:
    Set lobjTextoEspec = Nothing
    Exit Function
ControleErro:
    Set lobjTextoEspec = Nothing
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarTextoEspecifica��o ERRO - " & Err & " " & Err.Description & _
                Chr(13) & pMensagem)
End Function

Private Sub mnuTextoEndosso_Click()
    frmT47V024A.Show vbModal, Me
End Sub

Private Function mfSalvarTextoEndosso(pMensagem As String) As Boolean
Dim lobjTextoEndo         As clsA46V722A
Dim lobjstcTextoEndo      As stcA46V722B
Dim lNumPI                As Long
Dim lTextoInc             As String
Dim lTotLin               As Long
Dim lSeq                  As Long
Dim AuxArray              As Variant
    
    On Error GoTo ControleErro
    Set lobjTextoEndo = New clsA46V722A
    Set lobjTextoEndo.M46V999 = gM46V999
    
    Set lobjstcTextoEndo = New stcA46V722B
    Set lobjTextoEndo.ConexaoBD = objPedido.ConexaoBD
    If objstcPedido.ColecaoTextoEndo.Count > 0 Then
        lobjstcTextoEndo.NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
        If Not lobjTextoEndo.mfExcluir(lobjstcTextoEndo, "", "", "") Then
            pMensagem = lobjTextoEndo.UltimoErro
            GoTo Finalizar
        End If
    End If
    
    If Not gM46V999.gfDisponivel("frmT47V024A") Then
        mfSalvarTextoEndosso = True
        Exit Function
    End If
   
    lSeq = 0
    lTotLin = 0
    Set objstcPedido.ColecaoTextoEndo = New Collection
    
    If InStr(frmT47V024A.txtTextoEndosso, vbCrLf) <> 0 Then
       AuxArray = Split(frmT47V024A.txtTextoEndosso, vbCrLf)
       lTextoInc = AuxArray(lSeq)
       lTotLin = UBound(AuxArray)
    Else
        lTextoInc = frmT47V024A.txtTextoEndosso
    End If
   
   
OutraLinha:
   
    Set lobjstcTextoEndo = New stcA46V722B
   
    With lobjstcTextoEndo
        .NOSSO_NUMERO = objstcPedido.NOSSO_NUMERO
        .Cod_Org = 0
        .Cod_Ramo = gM46V999.gCod_Ramo
        .Num_Ped = objstcPedido.Num_Ped
        .Num_Apol = 0
        .Num_Endo = 0
        .Num_Seq = lSeq + 1
        .Dsc_Dizer_Endo = lTextoInc
    End With
    
    If Not lobjTextoEndo.mfIncluir(lobjstcTextoEndo, "", "", "") Then
        pMensagem = lobjTextoEndo.UltimoErro
        GoTo Finalizar
    End If
    
    objstcPedido.ColecaoTextoEndo.Add lobjstcTextoEndo, lobjstcTextoEndo.Class_ID
    
    If lSeq < lTotLin And lSeq + 1 < 20 Then
        lSeq = lSeq + 1
        lTextoInc = AuxArray(lSeq)
        GoTo OutraLinha
    End If
    
    mfSalvarTextoEndosso = True
Finalizar:
    Set lobjTextoEndo = Nothing
    Exit Function
ControleErro:
    Set lobjTextoEndo = Nothing
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(0, "mfSalvarTextoEndosso ERRO - " & Err & " " & Err.Description & _
                Chr(13) & pMensagem)
End Function


Private Function mfCalPremios() As Boolean
    'Fun��o: calcula pr�mios.
    '********************Vari�veis padr�o*************************************
    Dim lCkPoint        As String      'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lCkStr          As String       'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
                                        'lCkPoint. (posi��o das variaveis em ordem de declara��o)
    '********************Vari�veis Declaradas para o processamento************
    Dim flxFraVista         As MSFlexGrid       'Indicador da grade.
    Dim lI                  As Long             'Utilizada no for...loop
    Dim lJuros              As Variant          'Juros.
    Dim lMensagem           As String           'Mensagem.
    Dim lParcela1           As Double           'Quantidade de parcelas para antecipada.
    Dim lParcelaN           As Double           'Quantidade de parcelas para postecipada.
    Dim lPrefixo            As String           'Prefixo de mensagem.
    Dim lUltData            As Double           'Data da �ltima atualiza��o.
    Dim lQutdeParcelas      As Integer          'Quantidade de parcelas.
    Dim lobjstcItem         As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    
    Dim clsA46V102          As clsA46V102A      'DLL : consiste dados comuns (n�vel 2).
    Dim lObjstcVenc         As stcA46V713B      'Objeto da classe de estrutura stcA46V713B (TAB_PED_Venc)
    
    Dim lParcelasGrid       As Integer          ' Numero de parcelas exibidas no grid
    Dim lColecaoTemp        As Collection       'Cole��o tempor�ria de itens.
    
    Dim lretorno            As Integer          'Retorno da chamada da DLL.

    Dim lTotPr              As Single
    Dim lTotPr_Resg         As Single
    Dim lTotPr_Tarif        As Single
    Dim lobjstcCobert       As stcA46V704B      'Objeto da classe (stc) de coberturas.
    Dim lBoleto             As Long
    Dim lItemRenova         As String
    Static lsCod_Unicidade  As Double
    Static lsNOSSO_NUMERO   As String
    
    
    Dim lNovaColecao        As Collection       'Cole��o tempor�ria de itens.
    
    If gM46V999.gDesabilitarForms Then Exit Function

    
    mfCalPremios = False
    Screen.MousePointer = vbHourglass
    Err = 0
    lCkPoint = ""
    lCkStr = ""
    '1. Limpa grade de mensagens e par�metros de sa�da.
    On Error Resume Next
    Unload frmT47V017A
    Unload frmT47V025A
    On Error GoTo ControleErro
    
    'Zero os contadores de erros
    gM46V999.gQuaErrosNormas = 0
    gM46V999.gQuaErrosDemais = 0
    
    gM46V999.gPodeEmitir = True

    lCkPoint = "1"
  
    'Verifica parcelas
    If Not mfConsParcelas(lParcela1, lParcelaN) Then Exit Function
    lCkPoint = "2"
    
    '>>>>>fatiitus ok
    'Verifica juros
    If gM46V999.gTipoAcesso = DIGITACAO Then
        If Not mfConsJuros() Then Exit Function
        If frmT47V004A.optPagAntecipado(1).Value = True And CDblx(frmT47V004A.txtQuaParcelas) <> 0 Then
            '1200693 - RD Pagamento Antecipado
            If frmT47V004A.optForPagamento(0).Value Then
                If Not gM46V999.gfPreenchido(frmT47V004A.txtValorPgAnt.Text) Then
                    Call gM46V999.gpGraLog(1, "FGE0112 - Valor do pagamento antecipado inv�lido (n�o preenchido - campo obrigat�rio quando pagamento antecipado).")
                    Exit Function
                Else
                    If CDbl(frmT47V004A.txtValorPgAnt.Text) = 0 Then
                        Call gM46V999.gpGraLog(1, "FGE0112 - Valor do pagamento antecipado inv�lido (zerado - campo obrigat�rio quando pagamento antecipado).")
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    '>>>>>fatiitus
   
    Screen.MousePointer = vbHourglass
    '3. Salva dados.
Salvar:
    'If mnuArqSalvar.Enabled = True Or objstcPedido.objstcCtrlEmis.Num_PI = 0 Then
        If mfSalvar = False Then
            Exit Function
        End If
    'End If
    
    Screen.MousePointer = vbHourglass
    gAltComum = False
    lCkPoint = "3"
    
    '4. Posiciona controles.
    panAguarde.Visible = True
    panAguarde.ZOrder 0
    panAguarde.Caption = "C�lculo de pr�mios. Consistindo os dados de comum..."
    panAguarde.Refresh
    gM46V999.gpInicializar_gErros

    '5. COnsit�ncia dos dados de comum
    
    'lTipDocumento = objstcPedido.Cod_Tip_Doc
    
    'b) Itens.
    If objstcPedido.ColecaoPedLoc.Count = 0 Then
        Call gM46V999.gpGraLog(2, "FTE0011 - " & gM46V999.gPrefixo & _
                         "N�o localizado registro de item (c�lculo de pr�mios).")
        panAguarde.Visible = False
        Exit Function
    End If
    
    'Exportando os dados para as bases temp
    lBoleto = objstcPedido.objstcCtrlEmis.Num_Bloq
    
    '**BOLETO
    If frmT47V004A.optPagAntecipado(1).Value = True And frmT47V004A.optForPagamento(0).Value And lBoleto = 0 And (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
        lBoleto = 123456
    End If
    
    If Not gfExportarParaBaseTemp(objstcPedido.NOSSO_NUMERO, objstcPedido.Num_Ped, lBoleto, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If
    
    
    '6. Consiste dados comuns - proposta.
    Call mpVerComum
    
    lCkPoint = "5"
    If gM46V999.gQuaErros <> 0 Then
    'N�o � para zerar
        'Call mpEmiMensagens(gM46V999.gPrefixo)
        'panAguarde.Visible = False
        'panAguarde.Refresh
        'frmT47V017A.Show vbModeless, Me
        Exit Function
    End If
    
    
    '7. Processa item.
    'Vou copiar a cole��o porque durante o c�lculo os itens sofrem altera��o
    'e s�o removidos e adicionados novamente na cole��o, ent�o para que n�o aconte�a
    'de reprocessar o mesmo item, ent�o � feita a c�pia da cole��o.
    
    lCkPoint = "6"
    Set lColecaoTemp = objstcPedido.ColecaoPedLoc
    
    Set lNovaColecao = New Collection
    
    For Each lobjstcItem In lColecaoTemp
              
        'Ficha 0801749
        '[Verifico se existe mais de um item renovando o mesmo item da ap�lice.]
        If lobjstcItem.Tip_Emissao = e_TipEmissao_RenYasuda Then
            If InStr(1, lItemRenova, ";" & lobjstcItem.Num_Apol_Anterior & ";" & lobjstcItem.Num_Item_Anterior & ";") > 0 Then
                Call gM46V999.gpMenErro("AGE0154 - Os itens " & Mid(lItemRenova, InStr(1, lItemRenova, ";" & lobjstcItem.Num_Apol_Anterior & ";" & lobjstcItem.Num_Item_Anterior & ";") - 4, 4) & " e " & Format(lobjstcItem.Num_Item, "0000") & " est�o renovando o mesmo item (" & Format(lobjstcItem.Num_Item_Anterior, "0000") & ") da ap�lice " & Format(lobjstcItem.Num_Apol_Anterior, "0000000000") & ".")
            End If
            lItemRenova = lItemRenova & ";" & Format(lobjstcItem.Num_Item, "0000") & ";" & lobjstcItem.Num_Apol_Anterior & ";" & lobjstcItem.Num_Item_Anterior & ";"
        End If
        If mfCalPreItem(objstcPedido, objstcPedLoc, lobjstcItem, lNovaColecao) = False Then
            lPrefixo = " Item " & lobjstcItem.Num_Item & " "
            'Call mpEmiMensagens(lPrefixo)
            Call gM46V999.gpMenErro(lPrefixo)
        End If
        gM46V999.gpInicializar_gErros
        'i) Posiciona no pr�ximo item.
    Next lobjstcItem
    
    Set objstcPedido.ColecaoPedLoc = lNovaColecao
    
    lCkPoint = "11"

    '9. Se houve erro, encerra processo.
'    frmT47V020A.optOpcao(0).Enabled = True
    If gM46V999.gfDisponivel("frmT47V017A") Then
        panAguarde.Visible = False
        panAguarde.Refresh
        frmT47V017A.Show vbModeless, Me
        'gM46V999.gpInicializar_gErros
        Exit Function
    End If

    If gM46V999.gQuaErros <> 0 Then
    'N�o � para zerar
        'Call mpEmiMensagens(gM46V999.gPrefixo)
        'panAguarde.Visible = False
        'panAguarde.Refresh
        'frmT47V017A.Show vbModeless, Me
        Exit Function
    End If


    '10. Efetua totaliza��o.
    panAguarde.Caption = "Aguarde... Efetuando totaliza��o..."
    panAguarde.Refresh
    DoEvents
    'a) Abre o form T28V123A para c�lculo do fracionamento.
    lCkPoint = "12"
    Call frmT47V025A.gpInstanciarForm(objstcPedido)
    If gTemErro_T47V025A = True Then
        panAguarde.Visible = False
        Unload frmT47V025A
        If gM46V999.gfDisponivel("frmT47V017A") Then
            frmT47V017A.Show vbModeless, Me
        End If
        Exit Function
    End If

    lCkPoint = "13"
    '11. Efetua fracionamento.
    panAguarde.Caption = "Aguarde... Efetuando fracionamento..."
    panAguarde.Refresh
    DoEvents
    
    '12. Consiste dados comuns - n�vel 3.
      Set clsA46V102 = New clsA46V102A    'DLL : consiste dados comuns (n�vel 3).
      Set clsA46V102.M46V999 = gM46V999
      
      With clsA46V102
          'a) Prepara dados para consist�ncia de comum n�vel 3.
          
          If objstcPedido.ColecaoPedVenc.Count > 0 Then
              Set lObjstcVenc = objstcPedido.ColecaoPedVenc.Item(1)
              .Pro_CarDatPriParcela = lObjstcVenc.DAT_VENC
          End If
          
          'Proposta: forma de pagamento.
          '1 = carn� e 2 = d�bito em conta.
          If objstcPedido.Num_Cta_Bas_Cor <> 0 Then
              .Pro_ForPagamento = 2
          Else
              .Pro_ForPagamento = 1
          End If
          
'Ficha 900046
          If objstcPedido.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
              If objstcPedido.objstcCtrlEmis.Num_Bloq <> 0 Then
                  .Pro_PagAntecipado = 1
              Else
                  .Pro_PagAntecipado = 2
              End If
          Else
              .Pro_PagAntecipado = 0
          End If
          .Pro_QtdParcelas = objstcPedido.Qtd_Parc
          .Pro_CodSegurado = objstcPedido.Cod_Segurado

          'b) Executa consist�ncia.
          On Error Resume Next
          
          Call .VER_COMUM_NIVEL_3(objstcPedido)
          
          lI = 0
          
          For lI = 1 To gM46V999.gQuaErros
              Call gpMensagem(gM46V999.gerros(lI))
          Next
          
          'c) Se houve erro, encerra processo.
          If gM46V999.gfDisponivel("frmT47V017A") Then
              panAguarde.Visible = False
              panAguarde.Refresh
              frmT47V017A.Show vbModeless, Me
              Exit Function
          End If
  
          If Err <> 0 Then
              Call gM46V999.gpGraLog(1, "FTE0009 - " & gM46V999.gPrefixo & "Erro ao consistir dados comuns n�vel 3 (" & _
                               Err & " - " & Error & ").")
              On Error GoTo 0
              Exit Function
          End If
          
      End With

    Call frmT47V026A.gpInstanciarForm(objstcPedido)
    If gTemErro_T47V026A = True Then
        panAguarde.Visible = False
        Unload frmT47V026A
        If gM46V999.gfDisponivel("frmT47V017A") Then
            frmT47V017A.Show vbModeless, Me
        End If
        Exit Function
    Else
        lQutdeParcelas = Val(frmT47V004A.txtQuaParcelas.Text)
        
        If lQutdeParcelas = 0 And (objstcPedido.Cod_Reg = 106 Or objstcPedido.Cod_Reg = 107 Or _
           objstcPedido.Cod_Reg = 120 Or objstcPedido.Cod_Reg = 20 Or objstcPedido.Cod_Reg = 121 Or objstcPedido.Cod_Reg = 122) Then
            GoTo Pula_Parcelamento
        End If
        
        Do
            If frmT47V004A.optPagAntecipado(1) = True Then
                Set flxFraVista = frmT47V026A.flxFraVista
                'Numero de parcelas � o informado pelo usuario
                lParcelasGrid = lQutdeParcelas
            Else
                If frmT47V004A.optPagAntecipado(2) = True Then
                    Set flxFraVista = frmT47V026A.flxFra30dia
                    'Numero de parcelas � o informado pelo usuario
                    lParcelasGrid = lQutdeParcelas
                Else
                    Call mpPosResParcelamento("0", "0,00", "0,00", "0,00")
                    Unload frmT47V026A
                    Exit Do
                End If
            End If
            With frmT47V025A.lstResultado
                If gM46V999.gfPreenchido(lParcelasGrid) Then
                    If IsNumeric(lParcelasGrid) Then
                        If Val(lParcelasGrid) = 0 Then
                            .AddItem "Parcelamento n�o calculado:"
                            .AddItem "- quantidade de parcelas zerada."
                        Else
                            If Val(lParcelasGrid) <= gM46V999.gProQuaMaxParcelas Then
                                If flxFraVista.Rows = 1 Then
                                    gCarga = True
                                    frmT47V004A.txtQuaParcelas.Text = 0
                                    gCarga = False
                                    lParcela1 = 0
                                    lParcelaN = 0
                                    lJuros = 0
                                    Call mpPosResParcelamento(frmT47V004A.txtQuaParcelas.Text, _
                                                              lParcela1, lParcelaN, lJuros)
                                ElseIf flxFraVista.Rows <= CByte(lParcelasGrid) Then
                                    If CByte(lParcelasGrid) = 1 Then
                                        .AddItem "Fracionamento n�o contempla 1 parcela."
                                    Else
                                        .AddItem "Fracionamento n�o contempla " & _
                                                 lParcelasGrid & " parcelas."
                                        .AddItem "N�mero m�ximo de parcelas permitido: " & (flxFraVista.Rows - 1)
                                        gCarga = True
                                        frmT47V004A.txtQuaParcelas.Text = (flxFraVista.Rows - 1)
                                        gCarga = False
                                    End If
                                Else
                                    flxFraVista.Row = CByte(lParcelasGrid)
                                    flxFraVista.Col = 1  'Valor da primeira parcela.
                                    If flxFraVista.Text = "" Then flxFraVista.Text = 0
                                    lParcela1 = CDbl(flxFraVista.Text)
                                    flxFraVista.Col = 2  'Valor das demais parcelas.
                                    If flxFraVista.Text = "" Then flxFraVista.Text = 0
                                    lParcelaN = CDbl(flxFraVista.Text)
                                    flxFraVista.Col = 3  'Juro.
                                    lJuros = flxFraVista.Text
                                    Call mpPosResParcelamento(lParcelasGrid, _
                                                              lParcela1, lParcelaN, lJuros)
                                End If
                            Else
                                flxFraVista.Row = CByte(gM46V999.gProQuaMaxParcelas)
                                flxFraVista.Col = 1  'Valor da primeira parcela.
                                lParcela1 = CDbl(flxFraVista.Text)
                                flxFraVista.Col = 2  'Valor das demais parcelas.
                                lParcelaN = CDbl(flxFraVista.Text)
                                flxFraVista.Col = 3  'Juro.
                                lJuros = flxFraVista.Text
                                Call mpPosResParcelamento(gM46V999.gProQuaMaxParcelas, lParcela1, lParcelaN, _
                                                          lJuros)
                            End If
                        End If
                    Else
                        .AddItem "Parcelamento n�o calculado:"
                        .AddItem "- quantidade de parcelas n�o num�rica."
                    End If
                Else
                    .AddItem "Parcelamento n�o calculado:"
                    .AddItem "- quantidade de parcelas n�o preenchida."
                End If
            End With
'                Unload frmT47V026A
            Exit Do
        Loop
    End If

Pula_Parcelamento:

    '13. Atualiza registro de dados comuns com dados calculados.
    'a) Posiciona dados.
    lUltData = gM46V999.gfForData(3)
            
    lCkPoint = "15"
    'c) atualiza��o dos dados
    '===1) Situa��o do documento
    'Sandra (*)
    'objstcPedido.Cod_Sit_Ped = e_SitPedido_ComCalculo
    'AQUI
    '===2) Custo de emiss�o.
    objstcPedido.Val_Cust_Emis = frmT47V015A.txtCusEmissao.Text

'Ficha 900046
    If (frmT47V004A.optPagAntecipado(2).Value = True) Or (frmT47V004A.optPagAntecipado(1).Value = True And frmT47V004A.optForPagamento(1).Value = True) Then
        objstcPedido.Val_Pr_Antec = 0
    Else
        If gM46V999.CDblx(frmT47V004A.txtQuaParcelas.Text) = 0 Then
            objstcPedido.Val_Pr_Antec = 0
        ElseIf gM46V999.CDblx(frmT47V004A.txtValorPgAnt.Text) = 0 Then
            If gM46V999.gTipoAcesso <> DIGITACAO Then
                objstcPedido.Val_Pr_Antec = lParcela1
            End If
        Else
            objstcPedido.Val_Pr_Antec = gM46V999.CDblx(frmT47V004A.txtValorPgAnt.Text)
        End If
    End If
        
     '===16) Data e usu�rio da �ltima altera��o.
    objstcPedido.Val_IOF = frmT47V025A.flxResValores.TextMatrix(2, 1)
    'objstcPedido.VAL_PR = frmT47V025A.flxResValores.TextMatrix(0, 1)
    
    lTotPr = 0
    lTotPr_Resg = 0
    lTotPr_Tarif = 0
    
    For Each lobjstcItem In objstcPedido.ColecaoPedLoc
        If lobjstcItem.ColecaoPedCobert.Count > 0 Then
            For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
            'RESSEGURO
                If objstcPedido.Dat_Ini_Vig >= 20090501 Then
                    lTotPr = lTotPr + lobjstcCobert.VAL_PR
                    lTotPr_Resg = lTotPr
                    lTotPr_Tarif = lTotPr
                Else
                    lTotPr = lTotPr + lobjstcCobert.VAL_PR
                    lTotPr_Resg = lTotPr_Resg + lobjstcCobert.Val_Pr_Resg
                    lTotPr_Tarif = lTotPr_Tarif + lobjstcCobert.Val_Pr_Tarif
                End If
            Next lobjstcCobert
        Else
            lTotPr = lTotPr + objstcPedido.VAL_PR
            lTotPr_Resg = lTotPr_Resg + objstcPedido.Val_Pr_Ref_Resg
            lTotPr_Tarif = lTotPr_Tarif + objstcPedido.Val_Pr_Ref_Tarif
        End If
    Next lobjstcItem
    
    If objstcPedido.Cod_Reg = 20 Or objstcPedido.Cod_Reg = 120 Or objstcPedido.Cod_Reg = 121 Or objstcPedido.Cod_Reg = 122 Then
            gM46V999.gProPremioMinimo = 0
            gM46V999.gProCusApolice = 0
    End If
    
    If lTotPr > 0 And lTotPr < gM46V999.gProPremioMinimo Then
        lTotPr = gM46V999.gProPremioMinimo
    End If
    If lTotPr_Resg > 0 And lTotPr_Resg < gM46V999.gProPremioMinimo Then
        lTotPr_Resg = gM46V999.gProPremioMinimo
    End If
    If lTotPr_Resg > 0 And lTotPr_Tarif < gM46V999.gProPremioMinimo Then
        lTotPr_Tarif = gM46V999.gProPremioMinimo
    End If
    
'RESSEGURO
    If objstcPedido.Dat_Ini_Vig >= 20090501 Then
        objstcPedido.VAL_PR = lTotPr
        objstcPedido.Val_Pr_Ref_Resg = lTotPr
        objstcPedido.Val_Pr_Ref_Tarif = lTotPr
    Else
        objstcPedido.VAL_PR = lTotPr
        objstcPedido.Val_Pr_Ref_Resg = lTotPr_Resg
        objstcPedido.Val_Pr_Ref_Tarif = lTotPr_Tarif
    End If
    
    objstcPedido.Dat_Liber = Left(lUltData, 8)
    objstcPedido.Hor_Liber = Mid(lUltData, 9, 4)
    objstcPedido.Cod_User_Liber = gM46V999.gNomUsuario6
    objstcPedido.Qtd_Parc = frmT47V004A.txtQuaParcelas.Text
    
    '===17) Altera o registro.
    'a)
    If Not objPedido.mfAlterar(objstcPedido, "", "", mComUltData, mComUltUsuario) Then
        Call gM46V999.gpGraLog(2, objPedido.UltimoErro)
        Exit Function
    End If
    'b) Posiciona informa��es de data e usu�rio da �ltima atualiza��o.
    mComUltData = Left(lUltData, 8)
    mComUltUsuario = gM46V999.gNomUsuario6
    lCkPoint = "16"
    '13. Calcula e grava c�digo de verifica��o da unicidade da proposta.
    'a) Calcula c�digo de verifica��o.
'    If mfCalCodUnicidade_V1(objstcPedido, lMensagem) = False Then
'        Call gM46V999.gpGraLog(1, lMensagem)
'        Exit Function
'    End If
    objstcPedido.Dat_Liber = mComUltData
    objstcPedido.Cod_User_Liber = mComUltUsuario
    If Not objPedido.mfAlterar(objstcPedido, "", "", mComUltData, mComUltUsuario) Then
        lMensagem = objPedido.UltimoErro
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If
    lCkPoint = "17"
    
    If Not mfSalvarTextoEspecPadrao(lMensagem) Then GoTo ControleErro

    'Vai gerar o arquivo com o c�lculo
    'mnuArqTransmitir_Click 'n�o gerar
    panAguarde.Visible = False
    
    '14. Exibe form de totaliza��o
    frmT47V025A.Show vbModeless, Me
    If gM46V111.gCod_plano > 0 Then frmT47V020A.optOpcao(0).Enabled = True
    mfCalPremios = True
    If objstcPedido.mStatusDivSiscota = 0 Or gM46V999.gTipoAcesso = DIGITACAO Then
        mnuSituacao.Caption = Chr(183) & Space(10) & "PI: " & objstcPedido.objstcCtrlEmis.NUM_PI & " -  Consistido"
    End If
    
    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    'Obtem o inicio da mensagem de erro
    lMensagem = "mfCalPremios ERRO - " & Err & " " & Error & _
                Chr(13) & "CheckPoint " & lCkPoint & " - "
    
    
Resume Next
    'Carrega a variavel lCkStr
    On Error Resume Next
        lCkStr = " | " & ""
    On Error GoTo 0
    'Completa a mensagem de erro
    lMensagem = lMensagem & lCkStr
    
    Call gM46V999.gpGraLog(0, lMensagem)
    
    mfCalPremios = False
End Function

Private Function mfConsParcelas(ByRef pParcela1 As Double, pParcelaN As Double) As Boolean
    'Fun��o: calcula pr�mios.
    '********************Vari�veis padr�o*************************************
    Dim lCkPoint        As String      'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lCkStr          As String       'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
                                        'lCkPoint. (posi��o das variaveis em ordem de declara��o)
    '********************Vari�veis Declaradas para o processamento************
    Dim lMaxParcela         As Integer          'N�mero m�ximo de parcelas permitido
    Dim lMaxParcelaVista    As Integer          'N�mero m�ximo de parcelas permitido a vista
    Dim lMaxParcela30Dia    As Integer          'N�mero m�ximo de parcelas permitido 30 dias
    Dim lDataUltVencimento  As Double           'Data do ultimo vencimento possivel
    Dim lQutdeParcelas      As Integer          'Quantidade de parcelas
    Dim lAuxiliar           As String
    Dim lTipPessoa          As Integer
    Dim lMensagem           As String
    
    pParcela1 = 0
    pParcelaN = 0
    
    On Error GoTo ControleErro
    lCkPoint = "1"
    '2. Verifica quantidade de parcelas.
    If Not gM46V999.gfPreenchido(frmT47V004A.txtQuaParcelas.Text) And _
      (objstcPedido.Cod_Reg <> 106 And objstcPedido.Cod_Reg <> 107 And objstcPedido.Cod_Reg <> 120 And _
       objstcPedido.Cod_Reg <> 121 And objstcPedido.Cod_Reg <> 122 And objstcPedido.Cod_Reg <> 20) Then
        Screen.MousePointer = vbDefault
        Call gM46V999.gpGraLog(1, "FGE0013 - " & gM46V999.gPrefixo & "Quantidade de parcelas inv�lida (n�o preenchida).")
        Exit Function
    End If
    If Not IsNumeric(frmT47V004A.txtQuaParcelas.Text) And _
      (objstcPedido.Cod_Reg <> 106 And objstcPedido.Cod_Reg <> 107 And objstcPedido.Cod_Reg <> 120 And _
       objstcPedido.Cod_Reg <> 121 And objstcPedido.Cod_Reg <> 122 And objstcPedido.Cod_Reg <> 20) Then
        Screen.MousePointer = vbDefault
        Call gM46V999.gpGraLog(1, "FGE0014 - " & gM46V999.gPrefixo & "Quantidade de parcelas inv�lida (n�o num�rica).")
        Exit Function
    End If
    
    If Val(frmT47V004A.txtQuaParcelas.Text) = 0 And _
      (objstcPedido.Cod_Reg <> 106 And objstcPedido.Cod_Reg <> 107 And objstcPedido.Cod_Reg <> 120 And _
       objstcPedido.Cod_Reg <> 121 And objstcPedido.Cod_Reg <> 122 And objstcPedido.Cod_Reg <> 20) Then
        frmT47V004A.txtQuaParcelas.Text = 1
    End If
    
    'Guarda na variavel o numero de parcelas informado
    lQutdeParcelas = Val(frmT47V004A.txtQuaParcelas.Text)
    
    If frmT47V004A.optTipPessoa(0).Value = True Then
        lTipPessoa = 2  'Pessoa f�sica.
    Else
        lTipPessoa = 1  'Pessoa jur�dica.
    End If
    
    
    lCkPoint = "2"
    mfConsParcelas = True
    Exit Function
ControleErro:
    mfConsParcelas = False

End Function

Private Function mfConsJuros()
    'Fun��o: Consiste juros.
    
    If Not gM46V999.gfPreenchido(frmT47V004A.txtJuros.Text) Then
        Screen.MousePointer = vbDefault
        Call gM46V999.gpGraLog(1, "FNP0048 - " & gM46V999.gPrefixo & "Juros inv�lido (n�o preenchido).")
        Exit Function
    End If
    
    If Not IsNumeric(frmT47V004A.txtJuros.Text) Then
        Screen.MousePointer = vbDefault
        Call gM46V999.gpGraLog(1, "FNN0005 - " & gM46V999.gPrefixo & "Juros inv�lido (n�o num�rico).")
        Exit Function
    End If
    
    mfConsJuros = True
    Exit Function
    
ControleErro:
    mfConsJuros = False

End Function


Private Sub mpPosResParcelamento(ByVal pQuaParcelas As String, ByVal pParcela1 As String, _
                                 ByVal pParcelaN As String, ByVal pJuros As String)
    'Procedure: posiciona resultado do parcelamento.
Dim Aux As String

    With frmT47V025A.lstResultado
        .Clear
        If Not gM46V999.gfPreenchido(frmT47V004A.txtQuaParcelas.Text) Then
            If .ListCount = 0 Then
                .AddItem "Parcelamento n�o calculado:"
            End If
            .AddItem "- quantidade de parcelas n�o informada."
        Else
            If Not IsNumeric(frmT47V004A.txtQuaParcelas.Text) Then
                If .ListCount = 0 Then
                    .AddItem "Parcelamento n�o calculado:"
                End If
                .AddItem "- quantidade de parcelas n�o num�rica."
            End If
        End If
        If .ListCount <> 0 Then
            Exit Sub
        End If
        If frmT47V004A.optPagAntecipado(1) = True Then
            Aux = "Pagamento antecipado "
            If gM46V999.CDblx(frmT47V004A.txtValorPgAnt) <> 0 Then
                Aux = Aux & ":" & gM46V999.CDblx(frmT47V004A.txtValorPgAnt)
            End If
        Else
            Aux = "Pagamento postecipado"
        End If
        If gM46V999.CDblx(frmT47V004A.txtValorPgAnt) <> 0 Then
            If gM46V999.CDblx(frmT47V004A.txtValorPgAnt) <> gM46V999.CDblx(pParcela1) Then
                Aux = Aux & " - Diferen�a entre a "
                .AddItem Aux
                Aux = "1� Parcela ser� ratiada nas demais parcelas."
            End If
        End If
        .AddItem Aux
        .AddItem ""
        .AddItem "Quantidade de parcelas: " & pQuaParcelas
        If pQuaParcelas = 1 Then
            .AddItem "Valor da parcela: " & Format$(pParcela1, "standard")
        Else
            .AddItem "Primeira parcela: " & Format$(pParcela1, "standard")
            .AddItem "Demais parcelas: " & Format$(pParcelaN, "standard")
        End If
        .AddItem "Juro: " & Format$(pJuros, "standard") & "%"
        If pParcela1 = 0 Then
            .AddItem "Sem movimento de pr�mio."
        End If
    End With
End Sub


Private Sub mpImprimir(ByVal pOpcao As Byte)
    'Procedure: imprime ou visualiza.

    'Par�metro de entrada...pOpcao......0 = visualizar impress�o.
    '                                   1 = imprimir.
    '                                   2 = criar arquivo imagem.

    Dim lAuxiliar   As String           'Auxiliar.
    Dim lDocumento  As Byte             'Documento:
                                        '1 = cota��o simples.
                                        '2 = cota��o coletiva.
                                        '3 = proposta simples.
                                        '4 = proposta coletiva.
    Dim lImagem     As String           'Nome para arquivo imagem.
    Dim lMensagem   As String           'Mensagem.
    Dim lobjstcLoc  As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lCount      As Integer          'Contador.
    Dim lNomeArqIST  As String
    Dim lobjstcCob  As stcA46V704B
    Dim colecaoTMP  As Collection
    Dim colecaoTMP2  As Collection
    Dim lCoberturas     As String
    Dim lbdTAB_Cobert   As ADODB.Connection
    Dim lrsTab_Cobert   As clsYasRecordSet
    Dim lI              As Integer
    Dim lArray          As Variant
    Dim lDatDigitBoleto As Long
    Dim lSelect         As String
    
    '1. Verifica se os dados j� est�o calculados (as bases de dados e os recordsets ficar�o abertos).
''    If mfVerDados(lDocumento, lMensagem) = False Then
''        Call gM46V999.gpGraLog(1, lMensagem & " (impress�o)")
''        Exit Sub
''    End If

    '2. Confirma com usu�rio.
    Select Case pOpcao
        Case 0
            lMensagem = "AGE0041 - " & gM46V999.gPrefixo & "Confirma visualizar impress�o do documento?"
        Case 1
            If gM46V999.PerfilAcesso = e_Perfil_DIG_IST Or gM46V999.PerfilAcesso = e_Perfil_DIG_Yas Then
                lMensagem = "AGE0038 - " & gM46V999.gPrefixo & "Confirma gera��o do arquivo para confer�ncia?"
            Else
                lMensagem = "AGE0038 - " & gM46V999.gPrefixo & "Confirma impress�o do documento?"
            End If
        Case 2
            lMensagem = "AGE0034 - " & gM46V999.gPrefixo & "Confirma criar arquivo imagem do documento?"
    End Select
    
    'Ficha 1000400 - Bloquear todos os campos da cota��o quando esta for impressa ou for gerado seu arquivo (SyasMulti Vers�o Interna)
    If gM46V999.gTipoAcesso = CALCULO And gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And (UCase(gM46V999.gSerDIVERSOS) Like "*MIX*" = False) Then
        lMensagem = lMensagem & vbCrLf & "Aten��o ap�s a confirma��o, esse c�lculo n�o poder� ser alterado."
    End If
    
    Select Case MsgBox(lMensagem, vbYesNo + vbExclamation + vbDefaultButton3)
        Case vbYes
            Call gM46V999.gpGraLog(0, lMensagem & " Sim.")
        Case vbNo
            Call gM46V999.gpGraLog(0, lMensagem & " N�o.")
            Exit Sub
    End Select
    
    lNomeArqIST = ""
    If pOpcao = 1 And (gM46V999.PerfilAcesso = e_Perfil_DIG_IST Or gM46V999.PerfilAcesso = e_Perfil_DIG_Yas) Then
        On Error Resume Next
        lAuxiliar = Dir(gM46V999.gDir_Ist)
        If Err <> 0 Then
            Call gM46V999.gpGraLog(1, "Voc� n�o tem permiss�o para gerar arquivo de confer�ncia! Acesso Negado.")
            Exit Sub
        End If
        lNomeArqIST = gM46V999.gDir_Ist & "PI" & Format(objstcPedido.objstcCtrlEmis.NUM_PI, String(10, "0")) & ".txt"
        On Error Resume Next
        Kill lNomeArqIST
        On Error GoTo 0
    End If
    
    If objstcPedido.Dat_Pri_Calculo = 0 Then
        objstcPedido.Dat_Pri_Calculo = Left(gM46V999.gfForData(3), 8)
        If Not objPedido.mfAlterar(objstcPedido, "", "", mComUltData, mComUltUsuario) Then
            Call gM46V999.gpGraLog(2, objPedido.UltimoErro)
        End If
    '### PROJETO INSPECAO
''        If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
''            Call mfAgendarInspecaoProduto(objstcPedido, lMensagem)
''        End If
    End If
    
        
    'Primeiramente, as coberturas ser�o ordenadas.
    Set colecaoTMP2 = New Collection
    For Each lobjstcLoc In objstcPedido.ColecaoPedLoc
        Set colecaoTMP = New Collection
        lCoberturas = ""
        For Each lobjstcCob In lobjstcLoc.ColecaoPedCobert
            If gM46V999.gfPreenchido(lCoberturas) Then
                lCoberturas = lCoberturas & ", "
            End If
            lCoberturas = lCoberturas & lobjstcCob.Cod_Cobert
        Next lobjstcCob
        
        If lCoberturas = "" Then Exit For
        
        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Cobert, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
            
        lSelect = " Select * from Tab_Cobert where cod_cobert in (" & lCoberturas & ")"
        lSelect = lSelect & " and cod_Ramo = " & gM46V999.gCod_Ramo
        lSelect = lSelect & " and tab_cobert = " & lobjstcLoc.Tab_Clasf 'Ficha 1000308 - Permitir itens com planos diferentes no produto Roubo (c�lculo coletivo)
        lSelect = lSelect & " and COD_TARIF_INI <= " & gM46V999.gCodProduto
        lSelect = lSelect & " and COD_TARIF_TER >= " & gM46V999.gCodProduto
        lSelect = lSelect & " order by NUM_ORDEM_COBERT"
        
        If gM46V999.gfObtRegistro(lbdTAB_Cobert, lSelect, lrsTab_Cobert, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdTAB_Cobert)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        lCoberturas = ""
        Do While Not lrsTab_Cobert.EOF
            lCoberturas = lCoberturas & Format(lrsTab_Cobert("Cod_Cobert"), "00") & ";"
            lrsTab_Cobert.MoveNext
        Loop
        
        lArray = Split(lCoberturas, ";")
        For lI = 0 To UBound(lArray)
            If gM46V999.gfPreenchido(lArray(lI)) Then
                Set lobjstcCob = lobjstcLoc.ColecaoPedCobert.Item(lobjstcLoc.Class_ID & Format(lArray(lI), "0000") & Format(lobjstcLoc.Tab_Clasf, "00000")) 'Ficha 1000308 - Permitir itens com planos diferentes no produto Roubo (c�lculo coletivo)
                colecaoTMP.Add lobjstcCob, lobjstcCob.Class_ID
            End If
        Next lI
        Set lobjstcLoc.ColecaoPedCobert = Nothing
        Set lobjstcLoc.ColecaoPedCobert = colecaoTMP
        colecaoTMP2.Add lobjstcLoc, lobjstcLoc.Class_ID
    Next lobjstcLoc
    Set objstcPedido.ColecaoPedLoc = Nothing
    Set objstcPedido.ColecaoPedLoc = colecaoTMP2
    
   '3. Imprime documento.
      
    If gM46V999.gCod_Ramo = 710 Then
            lImagem = "ROUBO-" & Format$(objstcPedido.Num_Ped, "00000") & ".WMF"
            Select Case lDocumento
                Case 1      'Cota��o simples.
                    lAuxiliar = Left$(lImagem, 11) & "-Capa.WMF"
                    '1000308 - SyasMulti - Roubo - Cota��o Coletiva com Planos Distintos nos Itens
                    If gfImpDocumento(pOpcao, "ROUBOCOT1", lAuxiliar, objstcPedido, objstcPedLoc, 1, 2, lNomeArqIST) = False Then
                        Exit Sub
                    End If
                    lAuxiliar = Left$(lImagem, 11) & "-Anexo.WMF"
                    '1000308 - SyasMulti - Roubo - Cota��o Coletiva com Planos Distintos nos Itens
                    If gfImpDocumento(pOpcao, "ROUBOCOT2", lAuxiliar, objstcPedido, objstcPedLoc, 2, 2, lNomeArqIST) = False Then
                        Exit Sub
                    End If
                Case 2      'Cota��o coletiva.
                    'a) Imprime capa.
                    lAuxiliar = Left$(lImagem, 11) & "-capa.WMF"
                    '1000308 - SyasMulti - Roubo - Cota��o Coletiva com Planos Distintos nos Itens
                    If gfImpDocumento(pOpcao, "ROUBOCOT3", lAuxiliar, objstcPedido, objstcPedLoc, 1, (2 * objstcPedido.ColecaoPedLoc.Count) + 1, lNomeArqIST) = False Then
                        Exit Sub
                    End If
                    lCount = 1
                    'b) Imprime itens.
                    lAuxiliar = Left$(lImagem, 11) & "-item .WMF"
                    '1000308 - SyasMulti - Roubo - Cota��o Coletiva com Planos Distintos nos Itens
                    If gfImpDocumento(pOpcao, "ROUBOCOT4", lAuxiliar, objstcPedido, lobjstcLoc, 2, (2 * objstcPedido.ColecaoPedLoc.Count) + 1, lNomeArqIST) = False Then
                        Exit Sub
                    End If
                    'c) Bens
                    lAuxiliar = Left$(lImagem, 11) & "-bens.WMF"
                    '1000308 - SyasMulti - Roubo - Cota��o Coletiva com Planos Distintos nos Itens
                    Call gfImpDocumento(pOpcao, "ROUBOCOT5", lAuxiliar, objstcPedido, objstcPedLoc, (2 + objstcPedido.ColecaoPedLoc.Count), (2 * objstcPedido.ColecaoPedLoc.Count) + 1, lNomeArqIST)
            End Select
    End If
    
Continuar:
    
    If pOpcao = 1 And (gM46V999.PerfilAcesso = e_Perfil_DIG_IST Or gM46V999.PerfilAcesso = e_Perfil_DIG_Yas) And Trim(Dir(lNomeArqIST)) <> "" And lNomeArqIST <> "" Then
        Call gM46V999.gpGraLog(1, "Arquivo: " & lNomeArqIST & " gerado com sucesso!")
        Exit Sub
    End If

    'Salvo em Word
    If pOpcao = 2 Then
        Call gpGeraFiguraArqWord(lImagem, 0)
    End If

    '4. Emite mensagem.
    Select Case pOpcao
        Case 0
            Call gM46V999.gpGraLog(0, "AGE0033 - " & gM46V999.gPrefixo & "Documento visualizado.")
        Case 1
            Call gM46V999.gpGraLog(0, "AGE0032 - " & gM46V999.gPrefixo & "Documento impresso.")
        Case 2
            lAuxiliar = gM46V999.gAppPath & lImagem
            If Dir(lAuxiliar) <> "" Then
                Call gM46V999.gpGraLog(1, "AGE0031 - " & gM46V999.gPrefixo & "Criado(s) arquivo(s) imagem do documento." & _
                             vbLf & vbLf & "Veja arquivo(s) " & lAuxiliar & ".")
            Else
                lAuxiliar = Replace(lAuxiliar, "WMF", "DOC")
                Call gM46V999.gpGraLog(1, "AGE0031 - " & gM46V999.gPrefixo & "Criado(s) arquivo(s) imagem do documento." & _
                             vbLf & vbLf & "Veja arquivo(s) " & lAuxiliar & ".")
            End If
    End Select

    'Ficha 1000400 - Bloquear todos os campos da cota��o quando esta for impressa ou for gerado seu arquivo (SyasMulti Vers�o Interna)
    If objstcPedido.Dat_Pri_Calculo <> 0 And gM46V999.gTipoAcesso = CALCULO And gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And (UCase(gM46V999.gSerDIVERSOS) Like "*MIX*" = False) Then
        panAguarde.Caption = "Esta cota��o n�o pode ser alterado. Para altera��es realize c�pia desta refer�ncia."
        panAguarde.Visible = True
        gM46V999.gDesabilitarForms = True
'''        Call mpVerDesabilitar
    End If

End Sub



Public Sub gpGeraFiguraArqWord(pArquivo As String, ByRef pRetorno As Byte, Optional ByVal pOrigem As Integer = 0)
Dim lAplicacao      As Object
Dim lDocumento      As Object
Dim lArquivo        As String
Dim lArray          As Variant
Dim lI              As Integer
Dim lListaArquivos  As String

    On Error GoTo ControleErro
    lArquivo = Dir(gM46V999.gAppPath & Replace(pArquivo, ".WMF", "*.WMF"))
    Do While lArquivo <> ""
        lListaArquivos = lListaArquivos & lArquivo & ";"
        lArquivo = Dir()
    Loop
    lArray = Split(lListaArquivos, ";")
    
    pOrigem = gM46V999.CDblx(pOrigem)
    
    'Retirado por Simony 05/12/2007
    'If UBound(lArray) > 1 Then
    '    Call mpGeraFiguraArqWordSeparado(pArquivo, pRetorno)
    'End If
    
    If UBound(lArray) > 0 Then
        pRetorno = 0
        Set lAplicacao = CreateObject("Word.Application")
        
        lAplicacao.Documents.Add
        If pOrigem <> 1 Then
            lAplicacao.Visible = True
        End If
        
        Set lDocumento = lAplicacao.Documents(1)
        
        lDocumento.PageSetup.LineNumbering.Active = False
        lDocumento.PageSetup.TopMargin = lAplicacao.CentimetersToPoints(0.1)
        lDocumento.PageSetup.BottomMargin = lAplicacao.CentimetersToPoints(0.1)
        lDocumento.PageSetup.LeftMargin = lAplicacao.CentimetersToPoints(1)
        lDocumento.PageSetup.RightMargin = lAplicacao.CentimetersToPoints(0.1)
        
        'Call gM46V999.gpGraLog(0,UBound(lArray) & " folhas.")
        For lI = (UBound(lArray) - 1) To 0 Step -1
            If UCase(lArray(lI)) Like "*ITENS*" Then
                lDocumento.PageSetup.Orientation = 1
            End If
                        
            lDocumento.InlineShapes.AddPicture filename:=gM46V999.gAppPath & lArray(lI), _
                        LinkToFile:=False, SaveWithDocument:=True
        Next lI
        
        lDocumento.Protect password:="yasuda21", NoReset:=False, Type:=2
        If Dir(gM46V999.gAppPath & Replace(pArquivo, "WMF", "DOC")) <> "" Then
            Kill gM46V999.gAppPath & Replace(pArquivo, "WMF", "DOC")
        End If
        lDocumento.SaveAs gM46V999.gAppPath & Replace(pArquivo, "WMF", "DOC")
        If pOrigem <> 1 Then
            lAplicacao.WindowState = 1
        Else
            lAplicacao.Documents.Close
        End If
        Set lAplicacao = Nothing
        Set lDocumento = Nothing
        
        For lI = (UBound(lArray) - 1) To 0 Step -1
            Kill gM46V999.gAppPath & lArray(lI)
        Next lI
        
    End If
    pRetorno = 1
ControleErro:
    On Error Resume Next
    Set lAplicacao = Nothing
                
End Sub

Private Function mfExcRegPadrao(ByRef pMensagem As String) As Boolean
    'Func�o: exclui registro padr�o.

    mfExcRegPadrao = False

    '1. Confirma com usu�rio.
    pMensagem = "AGE0023 - Confirma exclus�o?" & vbLf & vbLf & "Registro padr�o."
    If MsgBox(pMensagem, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        Call gM46V999.gpGraLog(0, pMensagem & " Sim.")
    Else
        Call gM46V999.gpGraLog(0, pMensagem & " N�o.")
        mfExcRegPadrao = True
        Exit Function
    End If

    '2. Posiciona nosso n�mero do registro padr�o.
    Call gM46V999.gpExcluiRegistros("PADRAO" & gM46V999.gNomUsuario6, 0, "", "", pMensagem)

    mfExcRegPadrao = True
End Function


Private Function mfExcUltimoItem(pNumItem As String, ByRef pMensagem As String) As Boolean
    'Func�o: exclui item.

    Dim lEncontrou  As Boolean          'Indica que encontrou o registro.
    Dim lLinha      As Byte             'Linha selecionada.
    Dim lMensagem   As String           'Mensagem.

    mfExcUltimoItem = False
    
Retorno:
    
    Call gM46V999.gpExcluiRegistros(gM46V111.gComNosNumero, CLng(pNumItem), "", "", pMensagem)
    'Removo o item da cole��o de itens.
    objstcPedido.ColecaoPedLoc.Remove (CStr(gM46V111.gComNosNumero) & Format(CStr(pNumItem), "00000"))
    
    '6. Atualiza grade de itens e exibe outro item.
    With frmT47V014A.flxItens
        .Rows = .Rows - 1
    End With

    mfExcUltimoItem = True
    
End Function

Sub mfVerificaVersao()
    Dim lbdTAB_Produto   As ADODB.Connection 'Base de dados .
    Dim lbdTAB_Produto2  As ADODB.Connection 'Base de dados .
    Dim lMensagem        As String           'Mensagem.
    Dim lrsTAB_Produto   As clsYasRecordSet  'Registro .
    Dim lrsTAB_Produto2  As clsYasRecordSet  'Registro .
    Dim NomePasta        As String
    Dim lTotErr          As Integer
    Dim lSelect         As String
    
    On Error Resume Next

   ' If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Or UCase(gM46V999.gAppPath) = "S:\SYASPROD\DIVERSOS\" Then
   '     GoSub ExcluirBaseTemp
   '     Exit Sub
   ' End If
    
    If gfIsMetaframe() Then
        NomePasta = "L:\User\Syas\"
        MkDir "L:\User\"
        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & NomePasta)
    Else
        GoSub ExcluirBaseTemp
        NomePasta = "C:\Temp\Syas\"
        MkDir "C:\Temp\"
    End If
    
    MkDir NomePasta
    
    lTotErr = 0
    
    If Dir(NomePasta & "P00MULTI.MDB") = "" Then
        GoTo Copia:
    End If
    
    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Produto, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    
    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, NomePasta & "P00MULTI.MDB", lbdTAB_Produto2, lMensagem) = False Then
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
    panAguarde.Caption = "Atualizando Banco Dados Local. Tempo de Espera:      5 minutos. Por favor aguarde."
    panAguarde.Visible = True
    panAguarde.Refresh

    Call gM46V999.gpFecha2(lbdTAB_Produto)
    Call gM46V999.gpFecha2(lbdTAB_Produto2)
    Err = 0
    FileCopy gM46V999.gEndP00MULTI, NomePasta & "P00MULTI.MDB"
    panAguarde.Visible = False
    If Err <> 0 Then
        MsgBox Error
    End If
    If Dir(NomePasta & "P00MULTI.MDB") = "" Then
          Call gM46V999.gpFecha2(lbdTAB_Produto)
          Call gM46V999.gpFecha2(lbdTAB_Produto2)
          Call gM46V999.gpGraLog(2, "N�o foi poss�vel copia a base para a base CITRIX! Sistema Encerrado.")
          End
    End If
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
           Call gM46V999.gpGraLog(2, "N�o foi poss�vel remover a base CITRIX! Sistema Encerrado.")
           End
        End If
    End If
    Return
    
End Sub

Private Function mfCalPreItem(ByRef pobjstcPed As stcA46V702B, ByVal pobjstcPedLoc As stcA46V708B, ByVal pobjstcItem As stcA46V708B, ByRef lNovaColecao As Collection) As Boolean
    'Fun��o: Calcula pr�mios do item.
    'Par�metros de entrada...
    '                        pAtuStsItem.........Indica se o status do pr�mio pode ou n�o ser atualizado.
    'Fun��o: calcula pr�mios.
    '********************Vari�veis padr�o*************************************
    Dim lCkPoint        As String      'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lCkStr          As String       'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
                                        'lCkPoint. (posi��o das variaveis em ordem de declara��o)
    '********************Vari�veis Declaradas para o processamento************
    Dim lAguarde            As String           'Mensagem de aguarde.
    Dim lbdMulti            As ADODB.Connection 'Conex�o do Banco local das TAB_PED tempor�rias
    Dim lChave              As Long             'Posi��o na cole��o.
    Dim lErro               As Variant          'Erros da DLL de c�lculo de pr�mio.
    Dim lI                  As Byte             'Utilizado com For...Next.
    Dim lMensagem           As String           'Mensagem.
    Dim lMsgInsp            As String           'Mensagem de inspe��o
    Dim lPrefixo            As String           'Prefixo de mensagem.
    Dim lretorno            As Integer          'Retorno.
    Dim lrst                As clsYasRecordSet  'Registro.
    Dim lSelect             As String           'Select
    Dim lUltData            As Double           'Data da �ltima atualiza��o.
    Dim lobjstcCobert       As stcA46V704B      'Objeto da classe (stc) de coberturas.
    Dim lobjCobert          As clsA46V704A      'Objeto da classe de coberturas.
    Dim lobjPedObs          As clsA46V714A      'Objeto da classe de observa��o.
    Dim lobjPedLoc          As clsA46V708A      'Objeto da classe de itens.
    Dim lObjPed             As clsA46V702A      'Objeto da classe de estrutura stcA46V702A (TAB_PED)
    Dim lBonus              As Long
    Dim lColecaoTemp        As Collection
    Dim AtualizaProgres     As Boolean
    
    Dim lCod_Aceit          As Integer
    On Error GoTo ControleErro
    
    mfCalPreItem = False

    'Calcula
    'a) Prepara controles de mensagens.
    lAguarde = "C�lculo de pr�mios. Item " & pobjstcItem.Num_Item & ". " & vbLf
    lPrefixo = "Item " & pobjstcItem.Num_Item & " - "
    
    
    'b) Calcula pr�mios e salva registro com resultado do c�lculo.
    panAguarde.Caption = lAguarde & "Efetuando c�lculo de pr�mios do item..."
    panAguarde.Refresh

    'Calcula
    Set lobjPedLoc = New clsA46V708A
    Set lobjPedLoc.M46V999 = gM46V999
    
    Set lObjPed = New clsA46V702A
    Set lObjPed.M46V999 = gM46V999
    Set lobjPedLoc.ConexaoBD = lObjPed.ConexaoBD
    
    'Sandra - TIRAR COBOL e substituir pela rotina mpConsisteDadosComum e mpConsisteDadosItem
'    Call mpConsisteItem(pobjstcItem.Num_Item, lRetorno)         'DLL
     Call mpConsisteDadosItem(pobjstcItem.Num_Item, lretorno)

    'Call mpConsisteDadosItem(pobjstcItem.Num_Item, lRetorno)
    lCkPoint = "8"
    
    'Consiste item - n�vel 2.
    gM46V111.gCod_plano = pobjstcPed.Cod_Estip
        
    ''Resseguro facultativo
    'Call gpInicializarIdentificacaoResseguro(objstcPedido)
    
    panAguarde.Caption = lAguarde & "Consistindo dados do item ap�s o c�lculo... "
    panAguarde.Refresh

    Call mpVerIteNiv_2(pobjstcItem)
    lCkPoint = "9"
    
    'd) Consiste item - normas
    'panAguarde.Caption = lAguarde & "Consistindo as normas..."
    'panAguarde.Refresh
    
    lCkPoint = "10"
    'Verifica Regras de Aceita��o Pontuais
    
'    If VerificarNormas Then
'        lCod_Aceit = 0
'        lCod_Aceit = mfClasseAceit(pobjstcItem, pobjstcPed)
'
'        If lCod_Aceit = 3 Then
'            Call mpInformarErro("Proposta sem aceita��o.", 2, True)
'        ElseIf lCod_Aceit = 4 Then
'            Call mpInformarErro("Para essa proposta solicitar cota��o a Unidade de Produtos Diversos.", 2, True)
'        End If
'
'        lCkPoint = "11"
'    End If
    
'    Call mpVerNormas(pobjstcPed, pobjstcPedLoc, pobjstcItem.Num_Item, lRetorno, lCod_Aceit, lMsgInsp)
'
'    If gM46V999.gfPreenchido(lMsgInsp) = True Then
'        If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
'            If gM46V999.gfPreenchido(pobjstcItem.objstcPedInsp.Num_Tel) = False Then
'                Call mpInformarErro("Telefone de contato para inspe��o inv�lido (n�o preenchido).", 2, True)
'            End If
'
'            If gM46V999.gfPreenchido(pobjstcItem.objstcPedInsp.Pess_Cont) = False Or Len(pobjstcItem.objstcPedInsp.Pess_Cont) <= 3 Then
'                Call mpInformarErro("Contato para inspe��o inv�lido (n�o preenchido).", 2, True)
'            End If
'        End If
'
'        lRetorno = 0
'    End If
    
    If Not gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdMulti, lMensagem) Then
        Call gM46V999.gpMenErro(lMensagem)
        Exit Function
    End If
    
    AtualizaProgres = True
    
    lSelect = "Select * from TAB_PED_LOC"
    lSelect = lSelect & " where num_item = " & pobjstcItem.Num_Item
    lSelect = lSelect & " and num_apol = 0"
    
    'Busca os pr�mios de coberturas
    If Not gM46V999.gfObtRegistro(lbdMulti, lSelect, lrst, lMensagem) Then
        Call gM46V999.gpMenErro(lMensagem)
        Exit Function
    End If
    
    If Not lrst.EOF Then
        pobjstcItem.Per_Desc_Progres = lrst("Per_Desc_Progres")
    End If
    
    If Not lobjPedLoc.mfAlterar(pobjstcItem, "", "") Then
        lMensagem = lobjPedLoc.UltimoErro
        Exit Function
    End If
    
    lCkPoint = "50"
    
'    'Antes de Carregar objeto vai calcular novamente
'    If gM46V999.gCodProduto > gM46V999.TarifaAntiga Then
'        If gM46V111.gTipoGrupMulti = e_GrupoRamo_Multi Then
'            pobjstcItem.Cod_Sit_item = e_SitItem_NaoCalc
'           If Not mfCalNewPrecif(pobjstcItem, pobjstcPed, AtualizaProgres, lMensagem) Then
'                Call mpInformarErro(lMensagem, 2, True)
'                Call  gM46V999.gpMenErro(lMensagem)
'                Exit Function
'            End If
'        End If
'    End If
    
    gCarga = True
    'txtDescProgr.Text = pobjstcItem.Per_Desc_Progres
    gCarga = False
    
    'Salva resultado do c�lculo.
    'a) Atualiza registro do item.
    lUltData = CDbl(gM46V999.gfForData(3))
    pobjstcItem.Cod_Sit_item = e_SitItem_Calculado

    'If pobjstcItem.Num_Item = Val(txtNumItem.Text) Then
    '    gIteUltUsuario = gM46V999.gNomUsuario6
    '    gIteUltData = lUltData
    '    Call mpCarregarCobertGrid(pobjstcItem.ColecaoPedCobert)
    'End If

    'retirado por simony
    'pobjstcPed.ColecaoPedLoc.Remove pobjstcItem.Class_ID
    
    'Verifico se deve ser gravada a observa��o de inspe��o.
    If gM46V999.gfPreenchido(lMsgInsp) = True Then
        If InStr(1, gM46V999.gfRetirarCaracterEspecial(pobjstcItem.objstcPedObs.Observacao, True), gM46V999.gfRetirarCaracterEspecial(lMsgInsp, True)) = 0 Then
            Set lobjPedObs = New clsA46V714A
            Set lobjPedObs.M46V999 = gM46V999
            Set lobjPedObs.ConexaoBD = lObjPed.ConexaoBD
            
            If Not lobjPedObs.mfExcluir(pobjstcPedLoc.objstcPedObs, "", "") Then
                lMensagem = lobjPedObs.UltimoErro
                Exit Function
            End If
            
            With pobjstcItem.objstcPedObs
                .NOSSO_NUMERO = pobjstcItem.NOSSO_NUMERO
                .Num_Item = pobjstcItem.Num_Item
            End With
            
            If Not lobjPedObs.mfIncluir(pobjstcItem.objstcPedObs, "", "", "") Then
                lMensagem = lobjPedObs.UltimoErro
                Exit Function
            End If
        End If
    End If
        
    'pobjstcPed.ColecaoPedLoc.Add pobjstcItem, pobjstcItem.Class_ID
    
    lNovaColecao.Add pobjstcItem, pobjstcItem.Class_ID
        
    mfCalPreItem = True
    Exit Function
    
ControleErro:
    MsgBox "Erro  mfCalPreItem " & Err & lMensagem

End Function

Private Function mpConfiguraTela()
    Select Case gM46V999.gCod_Ramo
        Case e_CodRamo_Resp_Civil
            lblISBas.Caption = "Lim.Max Garant: "
            
            lblValorRisco.Enabled = False
            txtValoRisco.Enabled = False
            
            lblCoefISVR.Enabled = False
            txtCoefISVR.Enabled = False
        Case Else
            lblISBas.Caption = "LMG. B�sico: "
            
            lblValorRisco.Enabled = True
            txtValoRisco.Enabled = True
            
            lblCoefISVR.Enabled = True
            txtCoefISVR.Enabled = True
    End Select
End Function
