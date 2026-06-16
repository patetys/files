VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V004A 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleciona corretor"
   ClientHeight    =   4620
   ClientLeft      =   330
   ClientTop       =   660
   ClientWidth     =   8805
   Icon            =   "T46V004A.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4620
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flxConsulta 
      Height          =   4080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8800
      _ExtentX        =   15531
      _ExtentY        =   7197
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ForeColor       =   12582912
      ForeColorFixed  =   12582912
      ForeColorSel    =   12582912
      BackColorBkg    =   12632256
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin Threed.SSCommand cmdOk 
      Default         =   -1  'True
      Height          =   300
      Left            =   3300
      TabIndex        =   1
      Top             =   4200
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdCancelar 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   4500
      TabIndex        =   2
      Top             =   4200
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V004A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mTipo As Byte       'Controle que define se a tela é para visualizar lista de corretores
'ou se é para visualizar lista de produtor
'1 - Corretor
'2 - Produtor
Private Sub cmdCancelar_Click()
    If mTipo = 1 Then
        frmT46V003A.lblCodCorretor.Caption = ""
        frmT46V003A.txtCorretor.Text = ""
    Else
        frmT46V003A.lblCodProdutor.Caption = ""
        frmT46V003A.txtProdutor.Text = ""
    End If
    Unload Me
End Sub
Public Sub mpInstanciarForm(pTipo As Byte)
    mTipo = pTipo
    If mTipo = 1 Then
        Me.Caption = "Selecione o corretor"
    Else
        Me.Caption = "Selecione o produtor"
    End If
    Me.Show vbModal
End Sub
Private Sub cmdOk_Click()
    With flxConsulta
        'Não tirar desta ordem, pois ao mover alguma informação para o controle txtCorretor o programa
        'limpa o controle lblCodCorretor.
        If mTipo = 1 Then
            frmT46V003A.txtCorretor.Text = .TextMatrix(.Row, 2)
            frmT46V003A.lblCodCorretor.Caption = .TextMatrix(.Row, 0)
        Else
            frmT46V003A.txtProdutor.Text = .TextMatrix(.Row, 1)
            frmT46V003A.lblCodProdutor.Caption = .TextMatrix(.Row, 0)
        End If
    End With
    Unload Me
End Sub
Private Sub flxConsulta_DblClick()
    Call cmdOk_Click
End Sub
Private Sub Form_Activate()
    frmT46V001A.panForm = Mid$(Me.Name, 4)
    On Error Resume Next
    flxConsulta.Row = 1
    flxConsulta.Col = 1
    If flxConsulta.Visible = True And flxConsulta.Enabled = True Then
        flxConsulta.SetFocus
    End If
    On Error GoTo 0
End Sub
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass

    '1. Centraliza form.
    Height = 5000
    Call gM46V999.gpCenForm(Me)

    '2. Prepara grade.
    '   Grade é preenchida pelo programa chamador.
    With flxConsulta
        'a) Posiciona largura das colunas.
        .ColWidth(0) = 0
        .ColWidth(1) = 4800
        .ColWidth(2) = 3750
        'b) Posiciona títulos.
        If mTipo = 1 Then
            .TextMatrix(0, 0) = "Código do corretor"
            .TextMatrix(0, 1) = "Nome completo"
            .TextMatrix(0, 2) = "Nome abreviado"
        Else
            .TextMatrix(0, 0) = "Código do produtor"
            .TextMatrix(0, 1) = "Nome do produtor"
            .TextMatrix(0, 2) = "Nome abreviado"
            .ColWidth(2) = 0
        End If
    End With
    Screen.MousePointer = vbDefault
End Sub
