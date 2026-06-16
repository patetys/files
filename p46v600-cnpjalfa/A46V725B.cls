VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V725B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V725B
' Simony Sales
' ======================================================================================================================================================

Public login_user_rede As String    'PK
Public nom_user As String
Public cod_user_msp As String
Public ender_ip As String
Public rede As String
Public cod_sisger As String
Public prior_siscad As Long
Public prior_scr As Long
Public prior_multi As Long
Public prior_m_taxa As Long
Public prior_m_dcoml As Long
Public prior_m_normas As Long
Public cod_Depto As Long
Public cod_grupo As Long
Public Prior_Certif As Long
Public prior_reclama As Long
Public prior_m_franquia As Long

''###PROJETO SISCOTA
''DESCOMENTAR LINHAS
Public prior_m_siscota As Long
Public TipoUsuario As String
Public TipoAcesso As Integer
Public codCorretor As Long
Public PropTabela As String


' ======================================================================================================================================================

' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(login_user_rede)
End Property

