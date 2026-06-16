VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcParcelamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcParcelamento --> Demostrativo de Parcelamento
' Data/Hora: 05/10/2011 - 15:50:52
' Sistema: Multi
' Marcos Barra
' ======================================================================================================================================================

Public nosso_numero As String    'PK   'Nosso número.
Public Num_Parcela As Integer    'PK
Public TipPagamento As Integer    'PK

Public Val_Primeira_Vista As Double
Public Val_Demais_Vista As Double
Public Val_Juros_Vista As Double

Public Val_Primeira_30dias As Double
Public Val_Demais_30dias As Double
Public Val_Juros_30dias As Double

Private ObjStcPai As stcA46V702B

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Num_Parcela, "00")) & _
               CStr(Format(TipPagamento, "00"))
End Property

Public Property Set ObjPai(pobjstc As stcA46V702B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get ObjPai() As stcA46V702B
    Set ObjPai = ObjStcPai
End Property

Public Property Get EhParcelaValida() As Boolean
    EhParcelaValida = Not (Val_Primeira_Vista = 0 And _
                           Val_Demais_Vista = 0 And _
                           Val_Primeira_30dias = 0 And _
                           Val_Demais_30dias = 0)
End Property

Public Function getRpt() As stcParcelamentoRpt
    Static lrpt As stcParcelamentoRpt

    If lrpt Is Nothing Then
        Set lrpt = New stcParcelamentoRpt
        Set lrpt.stcOrigem = Me
    End If

    Set getRpt = lrpt

End Function

