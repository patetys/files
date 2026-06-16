VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmT46V143B 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Especificação do Risco"
   ClientHeight    =   3585
   ClientLeft      =   1365
   ClientTop       =   2790
   ClientWidth     =   5685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "T46V143B.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand cmdExtra 
      Height          =   240
      Index           =   2
      Left            =   4980
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   78
      ForeColor       =   12582912
      Picture         =   "T46V143B.frx":030A
   End
   Begin Threed.SSCommand cmdExtra 
      Height          =   240
      Index           =   1
      Left            =   4770
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   78
      ForeColor       =   12582912
      Picture         =   "T46V143B.frx":0624
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
            Picture         =   "T46V143B.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":128C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":15A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":18C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":1BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":1EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":2346
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":2660
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":297A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":2C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":30E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":3400
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":371A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":3B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143B.frx":3FBE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar barFerramentas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Visualizar"
            Object.ToolTipText     =   "Visualizar o formulário."
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Fecha Especificação"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel PanTexto 
      Height          =   3030
      Left            =   45
      TabIndex        =   3
      Top             =   495
      Width           =   5565
      _Version        =   65536
      _ExtentX        =   9816
      _ExtentY        =   5345
      _StockProps     =   15
      Caption         =   "Textos"
      ForeColor       =   8388608
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Begin Threed.SSPanel Pan_Quadro 
         Height          =   2580
         Left            =   90
         TabIndex        =   4
         ToolTipText     =   "Dados do segurado."
         Top             =   360
         Width           =   5370
         _Version        =   65536
         _ExtentX        =   9472
         _ExtentY        =   4551
         _StockProps     =   15
         BackColor       =   -2147483638
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   16
            Left            =   2685
            TabIndex        =   50
            Top             =   2175
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCheck CkEndLMI 
               Height          =   195
               Left            =   105
               TabIndex        =   51
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Endereços LMI"
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
            End
            Begin Threed.SSCommand CmdAbreLMI 
               Height          =   240
               Left            =   2310
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   15
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":42D8
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   8
            Left            =   2685
            TabIndex        =   5
            Top             =   960
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   8
               Left            =   2295
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":45F2
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   8
               Left            =   90
               TabIndex        =   7
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Cláusulas"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   14
            Left            =   75
            TabIndex        =   8
            Top             =   660
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   14
               Left            =   2295
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":490C
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   14
               Left            =   90
               TabIndex        =   10
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Contratante"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   12
            Left            =   75
            TabIndex        =   11
            Top             =   60
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   12
               Left            =   2295
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":4C26
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   12
               Left            =   90
               TabIndex        =   13
               Top             =   30
               Width           =   2145
               _Version        =   65536
               _ExtentX        =   3784
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Prazo da Obra"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   6
            Left            =   2685
            TabIndex        =   14
            Top             =   360
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   6
               Left            =   2295
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":4F40
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   6
               Left            =   90
               TabIndex        =   16
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Coberturas"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   13
            Left            =   75
            TabIndex        =   17
            Top             =   360
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   13
               Left            =   2295
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":525A
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   13
               Left            =   90
               TabIndex        =   19
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Prazo da Manutenção"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   7
            Left            =   2685
            TabIndex        =   20
            Top             =   660
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   7
               Left            =   2295
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":5574
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   7
               Left            =   90
               TabIndex        =   22
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Franquias"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   11
            Left            =   2685
            TabIndex        =   23
            Top             =   1860
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   11
               Left            =   2295
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":588E
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   11
               Left            =   105
               TabIndex        =   25
               Top             =   15
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Observações"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   15
            Left            =   75
            TabIndex        =   26
            Top             =   960
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   15
               Left            =   2295
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":5BA8
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   15
               Left            =   90
               TabIndex        =   28
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Local de Exposição"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   1
            Left            =   75
            TabIndex        =   29
            Top             =   1260
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   1
               Left            =   2295
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":5EC2
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   31
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Objeto Segurado"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   4
            Left            =   60
            TabIndex        =   32
            Top             =   2175
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   4
               Left            =   2295
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":61DC
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   4
               Left            =   90
               TabIndex        =   34
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Perímetro da Cobertura"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   9
            Left            =   2685
            TabIndex        =   35
            Top             =   1260
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   9
               Left            =   2295
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":64F6
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   9
               Left            =   90
               TabIndex        =   37
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Condições Especiais"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   3
            Left            =   75
            TabIndex        =   38
            Top             =   1860
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   3
               Left            =   2295
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":6810
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   3
               Left            =   90
               TabIndex        =   40
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Característica"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   2
            Left            =   75
            TabIndex        =   41
            Top             =   1560
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   2
               Left            =   2295
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":6B2A
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   43
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Riscos Cobertos"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   10
            Left            =   2685
            TabIndex        =   44
            Top             =   1560
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   10
               Left            =   2295
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":6E44
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   10
               Left            =   90
               TabIndex        =   46
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Outras Informações"
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
            End
         End
         Begin Threed.SSPanel pantxt 
            Height          =   300
            Index           =   5
            Left            =   2685
            TabIndex        =   47
            Top             =   60
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   529
            _StockProps     =   15
            ForeColor       =   12582912
            BackColor       =   -2147483637
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
            Begin Threed.SSCommand CmdAbreTxt 
               Height          =   240
               Index           =   5
               Left            =   2310
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   30
               Width           =   240
               _Version        =   65536
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   78
               ForeColor       =   12582912
               Enabled         =   0   'False
               Picture         =   "T46V143B.frx":715E
            End
            Begin Threed.SSCheck CkCodTexto 
               Height          =   195
               Index           =   5
               Left            =   90
               TabIndex        =   49
               Top             =   30
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "Especificação do Serviço"
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
            End
         End
      End
   End
   Begin VB.Menu mnu_arquivo 
      Caption         =   "&Arquivos"
      Begin VB.Menu mnu_imprimir 
         Caption         =   "Im&primir"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_visualizar 
         Caption         =   "&Visualizar"
      End
      Begin VB.Menu mnu_traco2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_sair 
         Caption         =   "Sai&r"
      End
   End
End
Attribute VB_Name = "frmT46V143B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub barFerramentas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case "VISUALIZAR"
        'Verifica se tem dado a salvar.
        If frmT46V101A.mnuArqSalvar.Enabled = True Then
            If MsgBox("Para visualizar é necessário salvar toda a proposta!" & Chr(13) & "Deseja salvar a proposta nesse momento?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
            If frmT46V101A.mfSalvar = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        Call frmT46V163A.gpInstanciarForm(frmT46V101A.objstcPedido, 1, e_TipDoc_Cotacao)
    Case "SAIR"
        Call mnu_sair_Click
    End Select
End Sub

Private Sub CkCodTexto_Click(Index As Integer, Value As Integer)
    Static lsCancel As Boolean
    Dim lUltTag As String

    If lsCancel = True Then
        lsCancel = False
        GoTo SUB_SAIDA
    End If

    If CkCodTexto.Item(Index).Value = True Then
        CmdAbreTxt(Index).Picture = cmdExtra(1).Picture
        CmdAbreTxt(Index).Enabled = True

        If gCarga = True Then
            GoTo SUB_SAIDA
        Else
            lUltTag = CkCodTexto(Index).Tag
            Call CmdAbreTxt_Click(Index)
            If lUltTag <> CkCodTexto(Index).Tag Then
                Call gpHabSalvar
            End If
        End If
    Else
        CmdAbreTxt(Index).Picture = cmdExtra(2).Picture
        CmdAbreTxt(Index).Enabled = False
        '----
        If gM46V999.gfPreenchido(CmdAbreTxt.Item(Index).Tag) = True Then
            If MsgBox("Esta opção excluirá o texto já informado." & Chr(13) & _
                      "Confirma a exclusão?", vbYesNo, Me.Caption) = vbYes Then
                CkCodTexto.Item(Index).Tag = ""   'A sequencia fica armazenada no tag deste objeto
                CmdAbreTxt.Item(Index).Tag = ""  'A texto RTF fica armazenado no tag deste objeto
                Call gpHabSalvar
            Else
                CmdAbreTxt(Index).Picture = cmdExtra(1).Picture
                CmdAbreTxt(Index).Enabled = True
                lsCancel = True     'marca p/ não entrar em loop
                CkCodTexto.Item(Index).Value = True
            End If
        End If
    End If

SUB_SAIDA:
    Exit Sub

End Sub


'********************ADRIANO**************************

Private Sub CkEndLMI_Click(Value As Integer)
    Static lsCancel As Boolean

    If lsCancel = True Then
        lsCancel = False
        GoTo SUB_SAIDA
    End If
    If CkEndLMI.Value = True Then   'selecionado
        CmdAbreLMI.Picture = cmdExtra(1).Picture
        CmdAbreLMI.Enabled = True

        If gCarga = True Then
            Exit Sub
        Else
            frmT46V168A.Show vbModal, Me
        End If
    Else
        CmdAbreLMI.Picture = cmdExtra(2).Picture
        CmdAbreLMI.Enabled = False
        If gCarga = True Then
            Exit Sub
        Else
            If frmT46V101A.chkLMI.Value = vbChecked Then
                If MsgBox("Esta opção excluirá os LMI(s) já informado(s)." & Chr(13) & _
                          "Confirma a exclusão?", vbYesNo, Me.Caption) = vbYes Then
                    Unload frmT46V168A
                Else
                    lsCancel = True     'marca p/ não entrar em loop
                    CmdAbreLMI.Enabled = True
                    CmdAbreLMI.Picture = cmdExtra(1).Picture
                    CkEndLMI.Value = vbChecked
                End If
            Else
                Unload frmT46V168A
            End If
        End If
    End If

SUB_SAIDA:
    Exit Sub

End Sub

Private Sub CmdAbreLMI_Click()
    frmT46V168A.Show vbModal, Me
End Sub

Private Sub CmdAbreTxt_Click(Index As Integer)
'Dim lstcTexto       As stcA06V045B
    Dim lstcTexto As stcA46V724B
    Dim lUltTag As String
    If gCarga = True Then
        GoTo SUB_SAIDA
    End If
    lUltTag = CmdAbreTxt(Index).Tag
    If Index < 12 Then
        Set lstcTexto = New stcA46V724B
        With lstcTexto
            .Tipo_Espec = Index
            .Nom_Espec = CkCodTexto(Index).Caption
            .Dsc_Espec = CmdAbreTxt(Index).Tag
        End With

        'Abre o Formulário editor de Textos
        Call frmT46V153A.gpInstanciarForm(lstcTexto)
        CmdAbreTxt.Item(Index).Tag = lstcTexto.Dsc_Espec
        If Trim(CmdAbreTxt.Item(Index).Tag) = "" Then
            CkCodTexto(Index).Value = 0
        End If
        CkCodTexto(Index).SetFocus
    Else
        gCarga = True
        frmT46V153B.TxtPrazoObra = CmdAbreTxt(e_Tipo_Espec_PrazoObra).Tag
        frmT46V153B.TxtPrazoManut = CmdAbreTxt(e_Tipo_Espec_PrazoManut).Tag
        frmT46V153B.TxtContratante = CmdAbreTxt(e_Tipo_Espec_Contratante).Tag
        frmT46V153B.TxtLocalExposicao = CmdAbreTxt(e_Tipo_Espec_LocalExpos).Tag
        frmT46V153B.Show 1
        If Trim(CmdAbreTxt(e_Tipo_Espec_PrazoObra).Tag) = "" Then
            CkCodTexto(e_Tipo_Espec_PrazoObra).Value = 0
        Else
            CkCodTexto(e_Tipo_Espec_PrazoObra).Value = 1
        End If
        If Trim(CmdAbreTxt(e_Tipo_Espec_PrazoManut).Tag) = "" Then
            CkCodTexto(e_Tipo_Espec_PrazoManut).Value = 0
        Else
            CkCodTexto(e_Tipo_Espec_PrazoManut).Value = 1
        End If
        If Trim(CmdAbreTxt(e_Tipo_Espec_Contratante).Tag) = "" Then
            CkCodTexto(e_Tipo_Espec_Contratante).Value = 0
        Else
            CkCodTexto(e_Tipo_Espec_Contratante).Value = 1
        End If
        If Trim(CmdAbreTxt(e_Tipo_Espec_LocalExpos).Tag) = "" Then
            CkCodTexto(e_Tipo_Espec_LocalExpos).Value = 0
        Else
            CkCodTexto(e_Tipo_Espec_LocalExpos).Value = 1
        End If
        CkCodTexto(e_Tipo_Espec_PrazoObra).SetFocus
        gCarga = False
    End If
    If lUltTag <> CmdAbreTxt(Index).Tag Then
        Call gpHabSalvar
    End If
SUB_SAIDA:
    Exit Sub

End Sub


Private Sub Form_Activate()
Finalizar:

'[FRH.07]->
    Dim lEstadoEmissao As Boolean
    lEstadoEmissao = True

    Select Case frmT46V101A.objstcPedido.objstcCtrlEmis.Cod_Iden_Stat_Acao
    Case 70, 80, 81, 90
        lEstadoEmissao = False
    End Select

    If lEstadoEmissao Then
        HabilitaControles Me
    Else
        DesabilitaControles Me, , frmT46V101A.objstcPedido.objstcCtrlEmis.Cod_Iden_Stat_Acao
    End If
    '<-[FRH]

    pantxt(16).Visible = False
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        pantxt(16).Visible = frmT46V101A.chkLMI.Visible
        If gM46V999.gNovaPropostaLMI Or frmT46V101A.chkLMI.Value = vbUnchecked Then
            CkEndLMI.Enabled = False
        End If
    Else
        pantxt(16).Visible = False
    End If

End Sub

Private Sub Form_Load()
    Me.Height = 4275
    Call gM46V999.gpCenForm(Me)
End Sub

Private Sub mnu_sair_Click()
    Me.Hide
End Sub


