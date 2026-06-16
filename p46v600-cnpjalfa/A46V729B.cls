VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46v729B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46v729B
' Data/Hora: 28/03/2008 - 11:05:11
' Sistema: P46V101
' Analista: fatiitus
' ======================================================================================================================================================

Public NUM_PI       As Double 'PK
Public nosso_numero As String
Public Dat_Triagem  As Double
Public User_Triagem As String
Public STATUS       As e_status_triagem
Public Enum e_status_triagem
    e_status_triagem_NDefinido = 0
    e_status_triagem_OK = 1
    e_status_triagem_NLoc = 2
    e_status_triagem_ErroTriagem = 3
    e_status_triagem_ErroSiscad = 4
    e_status_triagem_ErroPremio = 5
    e_status_triagem_Erro = 6
End Enum

Public Property Get Class_ID() As String
    Class_ID = CStr(NUM_PI)
End Property
