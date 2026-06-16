VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V746B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V746B - Tab_Ped_Liberacao_Serasa
' Data/Hora: 06/09/2016
' Sistema: Multi
' Rodrigo P S Rogério
' ======================================================================================================================================================

'Public NUM_PI                   As Long             'Num_Pi.
Public NUM_PI                   As Double             'Num_Pi.
Public nosso_numero             As String
Public Cod_User_Alt             As String           'Código da cobertura.
Public Dat_Alter                As Long             'Data de Alteracao.

' ======================================================================================================================================================

Public Property Get Class_Name() As String
    Class_Name = "A46V746B"
End Property

