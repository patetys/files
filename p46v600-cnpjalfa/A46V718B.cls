VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V718B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V718B
' Data/Hora: 04/10/2005 - 13:42:13
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================
Public nosso_numero As String           'PK
Public Num_Ped As String
Public Num_Item As String
Public Nom_Arq_Envio As String
Public Dat_Hor_Envio As String
Public Nom_Arq_Retorno As String
Public Dat_Hor_Retorno As String
Public Ind_Proposta As String
Public Nom_Segurado As String
Public Cod_Ramo As String
Public Dat_Ini_Vig As String
Public Dat_Ter_Vig As String
Public Cod_Tip_Logr_Risco As String
Public Nom_Logr_Risco As String
Public Dsc_Compl_Risco As String
Public Nom_Bairro_Risco As String
Public Nom_Cid_Risco As String
Public Sig_UF_Risco As String
Public Num_Cep_Risco As String
Public Num_Risco As String
Public Tip_Emissao As String
Public NUM_PI As String
Public Cod_Acao As String
Public Dat_Emissao As String
Public Dat_Ult_Inf As String
Public Ind_Cancel As String


' ======================================================================================================================================================

' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero)
End Property

