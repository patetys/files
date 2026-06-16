VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V747B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V747B - Tab_Ped_All_Risks
' Data/Hora: 08/11/2016
' Sistema: Multi
' Rodrigo Berçot
' ======================================================================================================================================================
Public Num_Item As Long         'PK         'Número do item.
Public nosso_numero As String   'Nosso Número
Public Cod_Org As Long          'Código do órgão.
Public Cod_Ramo As e_CodRamo    'Código do ramo.
Public NUM_PI As Double         'Número do PI
Public Num_Ped As Long          'Número do pedido.
Public Num_Apol As Double       'Número da apólice.
Public Num_Endo As Long         'Número do endosso.
Public Cod_Seq As Long          'Código de Sequencia.
Public Descricao As String      'Descrição da bicicleta.
Public Valor As Double          'Valor da bicicleta.
' ======================================================================================================================================================
Private ObjStcPai As stcA46V708B
' ======================================================================================================================================================

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V747B"
End Property

Public Property Set ObjPai(pobjstc As stcA46V708B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get ObjPai() As stcA46V708B
    Set ObjPai = ObjStcPai
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property

Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
End Property

Public Property Get M46V999() As clsM46V999
    If mClsM46V999 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V999 não atribuído para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V999 já
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46V999 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46V999 = mClsM46V999
End Property

Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
End Sub

Public Property Get gCopiaAllRisks(pNumItem As Long) As stcA46V747B
    Dim lobjstcAllRisks As stcA46V747B
    Set lobjstcAllRisks = New stcA46V747B
    Set lobjstcAllRisks.M46V999 = M46V999

    With lobjstcAllRisks
        .nosso_numero = nosso_numero
        .Cod_Org = Cod_Org
        .Cod_Ramo = Cod_Ramo
        .NUM_PI = NUM_PI
        .Num_Ped = Num_Ped
        .Num_Apol = Num_Apol
        .Num_Endo = Num_Endo
        .Num_Item = Num_Item
        .Cod_Seq = Cod_Seq
        .Descricao = Descricao
        .Valor = Valor
    End With
    Set gCopiaAllRisks = lobjstcAllRisks
End Property

Public Function getRpt() As stcA46V747Brpt
    Static lrpt As stcA46V747Brpt

    If lrpt Is Nothing Then
        Set lrpt = New stcA46V747Brpt
        Set lrpt.stcOrigem = Me
    End If

    Set getRpt = lrpt

End Function
