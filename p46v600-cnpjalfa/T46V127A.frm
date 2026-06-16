VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V127A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sobre"
   ClientHeight    =   4455
   ClientLeft      =   1890
   ClientTop       =   3045
   ClientWidth     =   6180
   ControlBox      =   0   'False
   ForeColor       =   &H00C00000&
   Icon            =   "T46V127A.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraX 
      Height          =   4335
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin VB.PictureBox Picture1 
         FillColor       =   &H8000000F&
         ForeColor       =   &H8000000F&
         Height          =   2175
         Left            =   120
         Picture         =   "T46V127A.frx":000C
         ScaleHeight     =   2115
         ScaleWidth      =   2235
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
      Begin Threed.SSCommand cmdOk 
         Cancel          =   -1  'True
         Default         =   -1  'True
         Height          =   330
         Left            =   4320
         TabIndex        =   6
         Top             =   3720
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1764
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "&Ok"
      End
      Begin Threed.SSPanel panX 
         Height          =   1200
         Left            =   240
         TabIndex        =   4
         Top             =   2880
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6174
         _ExtentY        =   2117
         _StockProps     =   15
         Caption         =   $"T46V127A.frx":135A
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
      Begin Threed.SSPanel panLicenca 
         Height          =   705
         Left            =   2640
         TabIndex        =   5
         Top             =   1560
         Width           =   3195
         _Version        =   65536
         _ExtentX        =   5636
         _ExtentY        =   1244
         _StockProps     =   15
         Caption         =   "Yasuda Seguros S/A"
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
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   45
         X2              =   5850
         Y1              =   2610
         Y2              =   2610
      End
      Begin VB.Label lblX 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Este produto está licenciado para:"
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   1
         Left            =   2640
         TabIndex        =   3
         Top             =   1200
         Width           =   3195
      End
      Begin VB.Label lblVersao 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Versão"
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   2640
         TabIndex        =   2
         Top             =   840
         Width           =   3195
      End
      Begin VB.Label lblX 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SYAS - Sistema Yasuda Seguros            "
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   2640
         TabIndex        =   1
         Top             =   480
         Width           =   3195
      End
   End
End
Attribute VB_Name = "frmT46V127A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOk_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = "T46V127A"
End Sub
Private Sub Form_Load()
'Centraliza form.
    Call gM46V999.gpCenForm(Me)

    lblVersao.Caption = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    panLicenca.Caption = "Yasuda Seguros S/A"
End Sub
