VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V113A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleciona corretor"
   ClientHeight    =   4725
   ClientLeft      =   135
   ClientTop       =   510
   ClientWidth     =   9000
   Icon            =   "T46V113A.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4725
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flxCorretores 
      Height          =   4080
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Visible         =   0   'False
      Width           =   8800
      _ExtentX        =   15505
      _ExtentY        =   7197
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ForeColor       =   8388608
      ForeColorFixed  =   12582912
      ForeColorSel    =   12582912
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin Threed.SSCommand cmdOk 
      Default         =   -1  'True
      Height          =   300
      Left            =   3400
      TabIndex        =   1
      ToolTipText     =   "Volta à proposta."
      Top             =   4300
      Width           =   1000
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdCancelar 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   4600
      TabIndex        =   2
      ToolTipText     =   "Volta à proposta."
      Top             =   4300
      Width           =   1000
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V113A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()
    Call flxCorretores_DblClick
End Sub
Private Sub flxCorretores_DblClick()
    flxCorretores.Col = 0   'Código do corretor.
    gCodCorretor = Format$(flxCorretores.Text, "0000000")
    flxCorretores.Col = 2   'Nome abreviado do corretor.
    gNomCorretor = flxCorretores.Text
    Unload Me
End Sub
Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
End Sub
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass

    '1. Centraliza form.
    Height = 6000
    Call gM46V999.gpCenForm(Me)

    '2. Prepara grade.
    With flxCorretores
        .Row = 0
        'a) Posiciona largura das colunas.
        .ColWidth(0) = 1000
        .ColWidth(1) = 4800
        .ColWidth(2) = 3750
        'b) Posiciona títulos.
        .Col = 0
        .Text = "Código"
        .Col = 1
        .Text = "Nome completo"
        .Col = 2
        .Text = "Nome abreviado"
    End With

    Screen.MousePointer = vbDefault
End Sub
