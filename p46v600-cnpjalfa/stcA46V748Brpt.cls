VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V748Brpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStcA46V748B As stcA46V748B

Public Property Set stcOrigem(ByRef pStcA46V748B As stcA46V748B)
    Set mStcA46V748B = pStcA46V748B
End Property

Public Property Get Sequencia() As String
    Dim lretorno As String
    
    lretorno = mStcA46V748B.Cod_Seq
    lretorno = IIf(lretorno = "", " ", lretorno)
    Sequencia = lretorno
End Property

Public Property Get Descricao() As String
    Dim lretorno As String
    
    lretorno = mStcA46V748B.Descricao
    lretorno = IIf(lretorno = "", " ", lretorno)
    Descricao = lretorno
End Property

Public Property Get Valor() As String
    Dim lretorno As String
    
    If mStcA46V748B.Valor > 0 Then
        lretorno = "R$ " & Format(mStcA46V748B.Valor, "standard")
    End If
    
    lretorno = IIf(lretorno = "", " ", lretorno)
    Valor = lretorno
End Property

Private Sub Class_Initialize()
    Set mStcA46V748B = Nothing
End Sub

Private Sub Class_Terminate()
    Set mStcA46V748B = Nothing
End Sub
