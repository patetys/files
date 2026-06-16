VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V711B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V711B --Tab_Ped_Juros
' Data/Hora: 16/09/2005 - 14:36:18
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================

Public nosso_numero As String
Public Num_Parcela As Long
Private Val_Pos_ As Double
Private Val_Ant_ As Double
Public Tip_Compensacao As Integer
' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B
' ======================================================================================================================================================


Public Property Get Val_Ant() As Double
    Val_Ant = Val_Ant_
End Property

Public Property Let Val_Ant(Value As Double)
    'Debug.Assert Num_Parcela <> 7
    Val_Ant_ = Value
End Property

Public Property Get Val_Pos() As Double
    Val_Pos = Val_Pos_
End Property

Public Property Let Val_Pos(Value As Double)
    'Debug.Assert Num_Parcela <> 1
    Val_Pos_ = Value
End Property

Public Property Get Class_ID() As String '
    Class_ID = CStr(nosso_numero) & _
               Format(Num_Parcela, "00") & _
               Format(Tip_Compensacao, "00")
End Property

Public Property Set ObjPai(pobjstc As stcA46V702B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property
