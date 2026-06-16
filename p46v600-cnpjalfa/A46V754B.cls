VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V754B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V754B - Tab_Ped_SISCOTA_LOG
' Data/Hora: 27/11/2017
' Sistema: Multi
' Juliano Iwai
' ======================================================================================================================================================
Public IdLog As Long
Public IdSiscota As Long
Public Data As Date
Public Tipo As Integer
Public Descricao As String

'Public Enum eTipoOperacao
'    ''0:Indefinido, Web 1: Inclusão,  2:Alteração, 3:Exclusão, 4:Outros;  Cotador Interno 11: Inclusão, 12:Alteração, 13:Exclusão, 14:Outros
'    eTipoOperacao_Indefinido = 0
'    eTipoOperacao_Incluir = 1
'    eTipoOperacao_Alterar = 2
'    eTipoOperacao_Excluir = 3
'    eTipoOperacao_Outros = 4
'    eTipoOperacao_CotadorInternt_Incluir = 11
'    eTipoOperacao_CotadorInternt_Alterar = 12
'End Enum

Public Property Get Class_Name() As String
    Class_Name = "A46V754B"
End Property

