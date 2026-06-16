VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V731B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46v731B
' Data/Hora: 16/06/2010 - 10:30:00
' Sistema: P46V600
' Analista: Sandra M. De Rosa
' ======================================================================================================================================================

Public nosso_numero As String
Public Num_Ped As Long
Public Num_Item As Long
Public Num_Cotac As String
Public Num_Insp As Long
Public Cod_Cobert As Long
Public Lmg_Cobert As Double

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero)
End Property

