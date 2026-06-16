VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V701B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V701B --> Tabela Tab_ped_ult
' Data/Hora: 05/09/2005 - 16:47:03
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================
Public Ult_Num_Ped As Long    'PK
Public Cod_User_Alt As String
Public Dat_Alt As Double
' ======================================================================================================================================================

' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(Ult_Num_Ped)
End Property

