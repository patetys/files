VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V752B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V752B - Tab_Ped_SISCOTA_Anexo
' Data/Hora: 27/11/2017
' Sistema: Multi
' Juliano Iwai
' ======================================================================================================================================================
Public IdAnexo As Long
Public IdObs As Long
Public DataAnexo As Date
Public NomeArquivo As String
Public NomeArquivoFisico As String
Public PastaArquivo As String
Public Sequencia As Integer

Public Property Get Class_Name() As String
    Class_Name = "A46V752B"
End Property
