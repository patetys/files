VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46O001B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================================================================================
' Class Name: stcA46O001B --Retorno de Consulta de Apólices
' Data/Hora: 08/06/2011 - 15:00:00
' Sistema: SyasMulti- Web
' Simony Sales Andrade
' ======================================================================================================================================================

Public Dsc_Ramo As String
Public Num_Apol As Double
Public Num_Item As Long
Public NomeSegurado As String
Public NumCNPJCPF As String
Public IniVigencia As String
Public TerVigencia As String
Public CodModalidade As Long
Public Cod_Ramo As Long
' ======================================================================================================================================================


Public Property Get Class_ID() As String '
    Class_ID = CStr(Num_Apol) & CStr(Num_Item)
End Property


