VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMDynaTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Enum EP_DT   'Entry Point DynaTrace
    epMetodoPublico = 1
    epMetodoInterno = 0
End Enum

Public Enum eSeveridadeErro
    eSeveridadeAlerta = 0
    eSeveridadeErro_Baixo = 1
    eSeveridadeErro_Medio = 2
    eSeveridadeErro_Alto = 3
End Enum

Public Sub Iniciar(ByVal pNomeAgente As String)
   
End Sub

Public Function gravaLog(a As Long, b As String, c As Integer, Optional d As Boolean = False)
'    ''20180314:Juliano - LOG
'    'pBasDados.CommandTimeout = 30
'    If mClsM46V999 Is Nothing Then
'        Set mClsM46V999 = New clsM46V999
'    End If
'
'    Call mClsM46V999.gpGraLog(0, " DynaTrace -> gravaLog -> " & Format(Time, "hh:mm:ss") & " Param B: " & B)
End Function

Public Function finalizarMetodoInterno(a As Long, b As Boolean, Optional c As String = "", Optional d As Integer = 0)

End Function

Public Function iniciarMetodoInterno(a As String, b As String)

End Function

Public Function iniciarMetodoPublico(a As String, b As String, c As String)

End Function

Public Function finalizarMetodoPublico(a, b, Optional c, Optional d, Optional e, Optional f)

End Function
