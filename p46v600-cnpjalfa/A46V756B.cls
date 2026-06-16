VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V756B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V756B - Classe de Dados Cartão de Crédito
' Data/Hora: 07/02/2018
' ======================================================================================================================================================

Public NUM_PI_EMIT          As Double 'PK
Public COD_RAMO_EMIT        As Double 'PK
Public TransactionId        As String 'PK
Public AUTHCODE             As String 'PK
Public DTMOV                As String
Public TRANSACTIONDATE      As String
Public TRANSACTIONSTATUS    As Integer
Public TRANSACTIONTYPE      As Integer
Public ORDERNUMBER          As Long
Public ORDERTOTAL           As Double
Public ERROR                As String
Public XML_SEND             As String
Public XML_RECEIVE          As String
Public NUM_TEMPO_TOTAL      As Long
Public TIMESATAMPBW         As String

' ======================================================================================================================================================

'Private ObjStcPai As stcA46V708B
Private mClsM46V999 As clsM46V999

Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
End Sub

' ======================================================================================================================================================
'Public Property Get Class_ID() As String
'    Class_ID = CStr(NOSSO_NUMERO) & _
'               CStr(Format(Ped_Loc_Pessoa_ID, "00000"))
'End Property

'Public Property Get gCopiaPessoa(pNum_Item As Long) As stcA46V735B
'    Dim lObjstcObs As stcA46V735B

'    Set lObjstcObs = New stcA46V735B
'    With lObjstcObs
'        .NOSSO_NUMERO = NOSSO_NUMERO
'        .Num_Apol = Num_Apol
'        .NUM_PI = NUM_PI
'        .Num_Item = pNum_Item
'        '.Ped_Loc_Pessoa_ID = Ped_Loc_Pessoa_ID
'        .Tip_Loc_Pessoa = Tip_Loc_Pessoa
'        .Nom_Pessoa = Nom_Pessoa
'        .Tip_Pessoa = Tip_Pessoa
'        .Num_CPF_CNPJ = Num_CPF_CNPJ
'        .Tip_Cobert = Tip_Cobert
'        .Per_Participacao = Per_Participacao
'        .Tip_Doc = Tip_Doc
'        .Num_Doc = Num_Doc
'    End With

'    Set gCopiaPessoa = lObjstcObs
'End Property

Public Property Get M46V999() As clsM46V999
    If mClsM46V999 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V999 não atribuído para a classe (" & Me.Class_Name & ")"
    End If

Retorno:
    Set M46V999 = mClsM46V999
End Property

Public Property Get Class_Name() As String
    Class_Name = "A46V756B"
End Property
Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
End Property

'Public Property Set ObjPai(pobjstc As stcA46V708B)
'    Set ObjStcPai = pobjstc
'End Property

'Public Property Get Cod_Unicidade() As Long
'    If Not ObjStcPai Is Nothing Then
'       Cod_Unicidade = ObjStcPai.Cod_Unicidade
'    End If
'End Property





