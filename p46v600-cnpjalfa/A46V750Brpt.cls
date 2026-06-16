VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V750Brpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================================================================================
' Class Name: stcA46V750Brpt - Tab_Ped_Obs_SISCOTA join Tab_Ped_Obs_SISCOTA_Log
' Data/Hora: 19/10/2017
' Sistema: Multi
' Juliano Iwai
'
'   Procedure RamosDiversos -> proc_BuscarObservacoesSiscota
'        SIS.IdSiscota,
'        SIS.Num_Ped,
'        SIS.Nosso_Numero,
'        SIS.SISCOTA,
'        SIS.IdGrupo,
'        SOBS.IdObs,
'        SOBS.Status,
'        SOBS.CodOrigem,
'        SOBS.Observacao,
'        SLOG.Data,
'        SOBS.Login
' ======================================================================================================================================================
Public IdObs As Long
Public IdSiscota As Long
Public nosso_numero As String
Public Num_Ped As Long
Public IdGrupo As String
Public Siscota As String
Public STATUS As Integer
Public StatusSiscota As Integer
Public CodOrigem As Integer
Public Observacao As String
Public Data As Date
Public Login As String

'''0: indefinido, 1: ativo, 2: removido/inativo
'Public Enum eStatusObsSISCOTA       'Status Tab_Ped_Obs_SISCOTA
'    eStatus_Indefinido = 0
'    eStatus_Ativo = 1
'    eStatus_RemovidoInativo = 2
'End Enum

Public Property Get Class_Name() As String
    Class_Name = "A46V750Brpt"
End Property

Private Sub Class_Initialize()
    'Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
    'Set mClsM46V999 = Nothing
End Sub

