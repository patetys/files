VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V714B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V714B -- Observacao
' Data/Hora: 21/09/2005 - 10:28:32
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================
Public Num_Item As Long    'PK
Public nosso_numero As String    'PK
Public Observacao As String
' ======================================================================================================================================================
Private ObjStcPai As stcA46V708B
' ======================================================================================================================================================
Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Num_Item, "00000"))
End Property

Public Property Get gCopiaObs(pNumItem As Long) As stcA46V714B
    Dim lObjstcObs As stcA46V714B

    Set lObjstcObs = New stcA46V714B
    With lObjstcObs
        .nosso_numero = nosso_numero
        .Num_Item = pNumItem
        .Observacao = Observacao
    End With
    Set gCopiaObs = lObjstcObs
End Property


Public Property Set ObjPai(pobjstc As stcA46V708B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property
