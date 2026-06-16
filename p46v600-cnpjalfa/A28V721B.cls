VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA28V721B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA28V721B
' Data/Hora: 08/05/2008 - 15:52:33
' Sistema: P28V101
' Analista: tjgomes
' Objetivo:
' ======================================================================================================================================================

' ======================================================================================================================================================
Public CCotacao As String    'PK
Public nosso_numero As String    'PK
Public Dat_Cotacao As Long
Public Email As String
Public Nome As String
Public Telefone As String
Public STATUS As Long
Public Senha As String
Public CCotacaoOrig As String
Public Usuario As String
Public MaxDesc As Double




' ======================================================================================================================================================

' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(CCotacao) & CStr(nosso_numero)
End Property

