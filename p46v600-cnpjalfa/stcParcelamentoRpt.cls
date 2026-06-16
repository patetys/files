VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcParcelamentoRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStcParcelamento As New stcParcelamento

Public Property Set stcOrigem(ByRef pStcParcelamento As stcParcelamento)
    Set mStcParcelamento = pStcParcelamento
End Property

Public Property Get Num_Parcela() As String
    Num_Parcela = Format(mStcParcelamento.Num_Parcela, "00")
End Property

Public Property Get Val_Primeira_Vista() As String
    Dim lretorno As String
    If mStcParcelamento.Val_Primeira_Vista > 0 Or _
       mStcParcelamento.Val_Demais_Vista > 0 Then
        lretorno = Format(mStcParcelamento.Val_Primeira_Vista, "#,###,##0.00")
    Else
        lretorno = "-"
    End If
    Val_Primeira_Vista = lretorno
End Property
Public Property Get Val_Demais_Vista() As String
    Dim lretorno As String
    If mStcParcelamento.Val_Primeira_Vista > 0 Or _
       mStcParcelamento.Val_Demais_Vista > 0 Then
        lretorno = Format(mStcParcelamento.Val_Demais_Vista, "#,###,##0.00")
    Else
        lretorno = "-"
    End If
    Val_Demais_Vista = lretorno
End Property
Public Property Get Val_Juros_Vista() As String
    Dim lretorno As String
    If mStcParcelamento.Val_Primeira_Vista > 0 Or _
       mStcParcelamento.Val_Demais_Vista > 0 Then
        lretorno = Format(mStcParcelamento.Val_Juros_Vista, "0.00") & "%"
    Else
        lretorno = "-"
    End If
    Val_Juros_Vista = lretorno
End Property

Public Property Get Val_Primeira_30dias() As String
    Dim lretorno As String
    If mStcParcelamento.Val_Primeira_30dias > 0 Or _
       mStcParcelamento.Val_Demais_30dias > 0 Then
        lretorno = Format(mStcParcelamento.Val_Primeira_30dias, "#,###,##0.00")
    Else
        lretorno = "-"
    End If
    Val_Primeira_30dias = lretorno
End Property
Public Property Get Val_Demais_30dias() As String
    Dim lretorno As String
    If mStcParcelamento.Val_Primeira_30dias > 0 Or _
       mStcParcelamento.Val_Demais_30dias > 0 Then
        lretorno = Format(mStcParcelamento.Val_Demais_30dias, "#,###,##0.00")
    Else
        lretorno = "-"
    End If
    Val_Demais_30dias = lretorno
End Property
Public Property Get Val_Juros_30dias() As String
    Dim lretorno As String
    If mStcParcelamento.Val_Primeira_30dias > 0 Or _
       mStcParcelamento.Val_Demais_30dias > 0 Then
        lretorno = Format(mStcParcelamento.Val_Juros_30dias, "0.00") & "%"
    Else
        lretorno = "-"
    End If
    Val_Juros_30dias = lretorno
End Property
