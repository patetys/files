VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V724B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V724B
' Data/Hora: 17/04/2007 - 15:03:10
' Sistema: P46V100
' Analista: Simony
' ======================================================================================================================================================
Public Tipo_Espec     As e_Tipo_Espec 'PK Tip
Public Nom_Espec      As String
Public Dsc_Espec      As String



Public Property Get Class_ID() As String
    Class_ID = CStr(Tipo_Espec)
End Property


