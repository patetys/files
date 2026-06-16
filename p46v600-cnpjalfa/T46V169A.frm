VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V169A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Triagem X Digitação"
   ClientHeight    =   3540
   ClientLeft      =   2025
   ClientTop       =   1935
   ClientWidth     =   8835
   Icon            =   "T46V169A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel panAviso 
      Height          =   3540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8835
      _Version        =   65536
      _ExtentX        =   15584
      _ExtentY        =   6244
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
      BevelWidth      =   3
      BevelInner      =   1
      Outline         =   -1  'True
      Begin MSFlexGridLib.MSFlexGrid flxResultado 
         Height          =   2580
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4551
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   10485760
         BackColorBkg    =   -2147483624
         GridLines       =   0
         GridLinesFixed  =   0
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdOk 
         Height          =   300
         Left            =   7560
         TabIndex        =   2
         Top             =   3000
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Ok"
         ForeColor       =   12582912
      End
      Begin VB.Label lblAviso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Triagem X Digitação - Divergências localizadas."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmT46V169A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub gpInstanciarForm(pErros As String)
    Dim lArray As Variant

    Dim lI As Long


    lArray = Split(pErros, ";")
    For lI = 0 To UBound(lArray) - 1
        If gM46V999.gfPreenchido(lArray(lI)) Then
            With flxResultado
                If lI <> 1 Then
                    .Rows = .Rows + 1
                End If
                .ColAlignment(0) = 0
                .TextMatrix(.Rows - 1, 0) = "- " & Trim(Replace(Replace(lArray(lI), Chr(9), ""), vbCrLf, " "))
            End With
        End If
    Next
    Screen.MousePointer = vbDefault
    Me.Show vbModal

End Sub
Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    With flxResultado
        .Cols = 1
        .Rows = 1
        .FixedRows = 0
        .ColWidth(0) = panAviso.Width - 100
    End With
    Call gM46V999.gpCenForm(Me)
End Sub


