VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V136A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensagens"
   ClientHeight    =   4425
   ClientLeft      =   3870
   ClientTop       =   2010
   ClientWidth     =   3855
   Icon            =   "T46V136A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   300
      Left            =   100
      TabIndex        =   0
      ToolTipText     =   "Volta à janela de item."
      Top             =   4100
      Width           =   1000
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin MSFlexGridLib.MSFlexGrid flxMensagens 
      Height          =   4000
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   3800
      _ExtentX        =   6694
      _ExtentY        =   7064
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   12648447
      ForeColor       =   8388608
      ForeColorFixed  =   4210752
      BackColorBkg    =   12648447
      GridColor       =   14737632
      GridColorFixed  =   14737632
      WordWrap        =   -1  'True
      Appearance      =   0
   End
   Begin VB.Label lblAuxMensagem 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label auxiliar de mensagem"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1750
      TabIndex        =   2
      Top             =   4100
      Visible         =   0   'False
      Width           =   2010
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmT46V136A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
    Dim lAcao As Integer
    Dim lMensagem As String
    Dim lI As Long

    If gM46V999.gQuaErrosDemais = 0 And gM46V999.gQuaErrosNormas > 0 And (gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or _
                                                                          gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) And frmT46V101A.objstcPedido.Ind_CotacaoSolic = False Then
        'Verifica se o usuário deseja solicitar uma cotação para o produto
        'DESCOMENTAR CLASSE 4
        For lI = 0 To flxMensagens.Rows - 1
            lMensagem = lMensagem & flxMensagens.TextMatrix(lI, 0) & vbCrLf
        Next lI
        Call frmT46V172A.gpInstanciarForm(lAcao, frmT46V101A.objPedido.Colecao, lMensagem)
        If lAcao = 1 Then
            gM46V999.gFinalizarPrograma = True
        End If
    End If
    Unload Me
End Sub
Private Sub flxMensagens_DblClick()
    If gM46V999.gfPreenchido(flxMensagens.Text) Then
        MsgBox flxMensagens.Text
    End If
End Sub
Private Sub Form_Activate()
'1. Posiciona nome do form.
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)

    '2. Coloca foco na grade.
    On Error Resume Next
    With flxMensagens
        .Row = 0
        .Col = 0
        If .Enabled And .Visible Then
            .SetFocus
        End If
    End With
    On Error GoTo 0
End Sub
Private Sub Form_Load()
    Dim lMarEsquerda As Integer   'Margem esquerda.
    Dim lMarSuperior As Integer   'Margem superior.

'1. Monta grade de mensagens.
    With flxMensagens
        .Cols = 1
        .ColWidth(0) = .Width - 300
    End With

    '2. Posiciona form.
    With frmT46V101A
        lMarEsquerda = .Left + .Width - Me.Width
        lMarSuperior = .Top + 1060
    End With
    Me.Move lMarEsquerda, lMarSuperior
End Sub
