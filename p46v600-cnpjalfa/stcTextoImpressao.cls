VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcTextoImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcTextoImpressao --
' Data/Hora: 21/11/2011 - 10:28:32
' Sistema: Multi
' ======================================================================================================================================================
Public Sequencia As String      'PK
Public Descricao As String
' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(Sequencia)
End Property
