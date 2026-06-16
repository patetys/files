VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V739B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V739B --> Tabela Tab_Ped_Textos
' Data/Hora: 23/09/2014 - 14:18:00
' Sistema: Multi
' Daniel Del Bianco Costa
' ======================================================================================================================================================

Public nosso_numero As String          'Nosso número.
Public Cod_Org As Integer              'Código do órgão.
Public Cod_Ramo As e_CodRamo           'Código do ramo.
Public Num_Ped As Long                 'Número do pedido.
Public Num_Apol As Double                'Sem uso - número da apólice.
Public Num_Endo As Long                'Sem uso - número do endosso.
Public Num_Item As Long                'Número do item.
'Public NUM_PI As Long                  'Número do Pi.
Public NUM_PI As Double                  'Número do Pi.
Public Cod_Cobert As Integer           'Código da cobertura.
Public Nom_Cobertura As String         'Nome da cobertura.
Public Tipo_Texto As Integer           'Código do Grupo das Coberturas .
Public Cod_Texto As Integer
Public Desc_Texto As String
Public NUM_SEQ As Integer
' ======================================================================================================================================================
Private ObjStcPai As stcA46V708B
' ======================================================================================================================================================

Private mClsM46V999 As clsM46V999
'Private mRpt As stcA46V739rpt

Public Property Get Class_Name() As String
    Class_Name = "A46V739B"
End Property

' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Num_Item, "00000")) & _
               CStr(Format(Cod_Cobert, "0000")) & _
               CStr(Format(NUM_SEQ, "0000"))
End Property

Public Property Get gCopiaCobert(pNumItem As Long) As stcA46V739B
    Dim lobjstcCobert As stcA46V739B

    Set lobjstcCobert = New stcA46V739B
    Set lobjstcCobert.M46V999 = M46V999

    With lobjstcCobert
        .nosso_numero = nosso_numero
        .Cod_Org = Cod_Org
        .Num_Ped = .Num_Ped
        .Cod_Ramo = Cod_Ramo
        .Num_Apol = Num_Apol
        .Num_Endo = Num_Endo
        .Num_Item = pNumItem
        .NUM_PI = NUM_PI
        .Cod_Cobert = Cod_Cobert
        .NUM_SEQ = NUM_SEQ
        .Tipo_Texto = Tipo_Texto
        .Cod_Texto = Cod_Texto
        .Desc_Texto = Desc_Texto
    End With
    
    Set gCopiaCobert = lobjstcCobert

SAIDA:
    GoSub FechaObj
    Exit Property

FechaObj:
    Set lobjstcCobert = Nothing
    Return
End Property

Public Property Set ObjPai(pobjstc As stcA46V708B)
    
    Set ObjStcPai = pobjstc

End Property

Public Property Get ObjPai() As stcA46V708B
    If ObjStcPai Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto ObjStcPai(Tipo:stcA46V708B) não atribuído para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe stcA46V708B já
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (stcA46V708B) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set ObjPai = ObjStcPai
End Property

'Public Function getRpt() As stcA46V739rpt
'
'    If mRpt Is Nothing Then
'        Set mRpt = New stcA46V739rpt
'        Set mRpt.stcOrigem = Me
'    End If
'
'    Set getRpt = mRpt
'
'End Function

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

Public Property Get M46V111() As clsM46V111
    Set M46V111 = M46V999.M46V111
End Property

Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
    Set ObjStcPai = Nothing
    'Set mRpt = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
    Set ObjStcPai = Nothing
    'Set mRpt = Nothing
End Sub




