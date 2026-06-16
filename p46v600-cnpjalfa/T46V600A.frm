VERSION 5.00
Begin VB.Form FrmT46V600A 
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   4845
   ClientTop       =   1980
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   825
      Left            =   1155
      TabIndex        =   0
      Top             =   450
      Width           =   2445
   End
End
Attribute VB_Name = "FrmT46V600A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim clsA46V601          As clsA46V601A      'Classe: cria arquivo de transmissão.
Dim lEndArqTransmissao As String
Dim lEndArqTransInsp   As String
Dim pMensagem As String
Dim objstcPedido As stcA46V702B

    lEndArqTransmissao = "C:\User_c\Projetos\SyasMulti\P46V100\Remessa\169949_00628951963000000001.008"

Call gM46V999.gpInicia("")

Set clsA46V601 = New clsA46V601A
Set clsA46V601.M46V999 = gM46V999

'If Not clsA46V601.gfRecArqTransmissao(lEndArqTransmissao, pMensagem, objstcPedido) Then
'    MsgBox pMensagem
'    Exit Sub
'End If

lEndArqTransInsp = gM46V999.gAppPath & "447627_01025249496007823866_1.055"

If Not clsA46V601.gfRecArqTransInsp(lEndArqTransInsp, pMensagem, objstcPedido) Then
    MsgBox pMensagem
    Exit Sub
End If
    Exit Sub


    'lEndArqTransInsp = gM46V999.gAppPath & "447627_01025249496007823866_1.055"
    '
    'If Not clsA46V601.gfRecArqTransInsp(lEndArqTransInsp, pMensagem, objstcPedido) Then
    '    MsgBox pMensagem
    '    Exit Sub
    'End If
    '
    'Set clsA46V601 = Nothing
Dim auxunicidade As Long

'If gNomeTabImport = "" Then
    auxunicidade = objstcPedido.Cod_Unicidade
    
    If gM46V999.gfCalCodUnicidade_V1(objstcPedido, pMensagem) = False Then
        Call gM46V999.gpGraLog(1, pMensagem)
        Exit Sub
    End If
    
    
    If auxunicidade <> objstcPedido.Cod_Unicidade Then
        pMensagem = "IIHHHH"
    End If
'End If

End Sub


