VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V722B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V722B -- Texto para Endosso
' Data/Hora: 19/10/2005 - 10:28:32
' Sistema: Multi
' Simony Sales
' ======================================================================================================================================================
Public nosso_numero As String    'PK
Public Cod_Org As Integer    'Código do órgão.
Public Cod_Ramo As e_CodRamo    'Código do ramo.
Public Num_Ped As Long    'Número do pedido.
Public Num_Apol As Double     'Sem uso - número da apólice.
Public Num_Endo As Long     'Sem uso - número do endosso
Public NUM_SEQ As Integer
Public Dsc_Dizer_Endo As String         'Local de Exposicao.

' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B
' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(NUM_SEQ, "00"))
End Property

Public Property Get gCopiaTextoEndosso() As stcA46V722B
    Dim lObjstcObs As stcA46V722B

    Set lObjstcObs = New stcA46V722B
    With lObjstcObs
        .nosso_numero = nosso_numero
        .Cod_Org = Cod_Org
        .Cod_Ramo = Cod_Ramo
        .Num_Ped = Num_Ped
        .Num_Apol = Num_Apol
        .Num_Endo = Num_Endo
        .NUM_SEQ = NUM_SEQ
        .Dsc_Dizer_Endo = Dsc_Dizer_Endo
    End With
    Set gCopiaTextoEndosso = lObjstcObs
End Property

Public Property Set ObjPai(pobjstc As stcA46V702B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get ObjPai() As stcA46V702B
    Set ObjPai = ObjStcPai
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property
