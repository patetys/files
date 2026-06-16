VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V010A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sobre"
   ClientHeight    =   4440
   ClientLeft      =   2295
   ClientTop       =   2235
   ClientWidth     =   6120
   ControlBox      =   0   'False
   ForeColor       =   &H00C00000&
   Icon            =   "T46V010A.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4440
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraX 
      Height          =   4335
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin VB.PictureBox Picture1 
         Height          =   2175
         Left            =   120
         Picture         =   "T46V010A.frx":000C
         ScaleHeight     =   2115
         ScaleWidth      =   2235
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin Threed.SSCommand cmdOk 
         Cancel          =   -1  'True
         Default         =   -1  'True
         Height          =   420
         Left            =   4320
         TabIndex        =   2
         Top             =   3630
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   741
         _StockProps     =   78
         Caption         =   "&Ok"
         ForeColor       =   8388608
      End
      Begin Threed.SSPanel panX 
         Height          =   1200
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6174
         _ExtentY        =   2117
         _StockProps     =   15
         Caption         =   $"T46V010A.frx":135A
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
      Begin Threed.SSPanel panLicenca 
         Height          =   705
         Left            =   2640
         TabIndex        =   4
         Top             =   1560
         Width           =   3195
         _Version        =   65536
         _ExtentX        =   5636
         _ExtentY        =   1244
         _StockProps     =   15
         Caption         =   "Yasuda Seguros S/A"
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
         BevelOuter      =   1
      End
      Begin VB.Label lblX 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SYAS - Sistema Yasuda Seguros  - MULTI"
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   0
         Left            =   2640
         TabIndex        =   7
         Top             =   480
         Width           =   3195
      End
      Begin VB.Label lblVersao 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Versão"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2640
         TabIndex        =   6
         Top             =   840
         Width           =   3195
      End
      Begin VB.Label lblX 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Este produto está licenciado para:"
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   1
         Left            =   2640
         TabIndex        =   5
         Top             =   1200
         Width           =   3195
      End
   End
End
Attribute VB_Name = "frmT46V010A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOk_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    frmT46V001A.panForm.Caption = "T46V010A"
End Sub
Private Sub Form_Load()
'Centraliza form.
    Call gM46V999.gpCenForm(Me)

    lblVersao.Caption = "Versão: " & Format$(App.Major, "000") & "." & Format$(App.Minor, "000") & "." & _
                        Format$(App.Revision, "000")
    panLicenca.Caption = "Yasuda Seguros S/A"
End Sub
