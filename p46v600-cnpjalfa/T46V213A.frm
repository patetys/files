VERSION 5.00
Begin VB.Form frmT46V213A 
   Caption         =   "Número da Transação"
   ClientHeight    =   720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   720
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTransactionId 
      Enabled         =   0   'False
      Height          =   325
      Left            =   1920
      TabIndex        =   1
      Top             =   200
      Width           =   2535
   End
   Begin VB.Label lblTransactionID 
      Caption         =   "Número da Transação:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmT46V213A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
    txtTransactionId = gfObterDadosTransacaoCartao(frmT46V101A.objstcPedido.nosso_numero)
    
End Sub


Public Function gfObterDadosTransacaoCartao(ByRef nosso_numero As String) As String
    
    Dim lConexaoDiv As ADODB.Connection
    Dim lSelect As String
    Dim lrst As clsYasRecordSet
    Dim lTransactionId As String
    Dim lMensagem As String    'Mensagem.
    
    Dim methodName As String
    methodName = "gfObterDadosTransacaoCartao"
    
    Err.Clear: UltimoErro = ""

    On Error GoTo ControleErro

    '- Abre as bases que serão utilizadas
    If Not gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gM46V999.gTipDIVERSOS, gM46V999.gEndDIVERSOS, lConexaoDiv, lMensagem) Then
        Call gM46V999.gpMenErro(lMensagem)
        GoTo SAIDA
    End If

    lSelect = "SELECT TransactionId FROM RamosDiversos..Tab_Controle_Cartao WHERE Nosso_Numero = (SELECT NOSSO_NUMERO_OLD FROM RamosDiversos..TAB_PED_NUM_COPIA "
    lSelect = lSelect & "WHERE NOSSO_NUMERO_NEW = " & gM46V999.gfFormatarCamposSQL(nosso_numero) & ")"
    
    If Not gM46V999.gfObtRegistro(lConexaoDiv, lSelect, lrst, lMensagem) Then
        Call gM46V999.gpMenErro(lMensagem)
        GoTo SAIDA
    End If
    
    If Not lrst.EOF Then
        lTransactionId = lrst("TransactionId")
    Else
        lTransactionId = ""
    End If
    
    gfObterDadosTransacaoCartao = lTransactionId

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:

    GoTo SAIDA

FechaObj:
    Set lConexaoDiv = Nothing
    Set lrst = Nothing
    
    Return
    
End Function
