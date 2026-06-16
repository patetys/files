VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V753B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V753B - Tab_Ped_SISCOTA
' Data/Hora: 19/10/2017
' Sistema: Multi
' Juliano Iwai
' ======================================================================================================================================================
Public IdSiscota As Long
Public nosso_numero As String
Public Num_Ped As Long
Public IdGrupo As String
Public Siscota As String
Public STATUS As Integer

Public ColecaoLog As Collection
Public ColecaoObs As Collection

'''0: indefinido, 1: ativo, 2: removido/inativo
'Public Enum eStatusSISCOTA       'Status Tab_Ped_SISCOTA
'    eStatus_Indefinido = 0
'    eStatus_Ativo = 1
'    eStatus_RemovidoInativo = 2
'End Enum

Public Property Get Class_Name() As String
    Class_Name = "A46V753B"
End Property

Private Sub Class_Initialize()
    'Set mClsM46V999 = Nothing
    Set ColecaoLog = New Collection
    Set ColecaoObs = New Collection
End Sub

Private Sub Class_Terminate()
    'Set mClsM46V999 = Nothing
    Set ColecaoLog = Nothing
    Set ColecaoObs = Nothing
End Sub

