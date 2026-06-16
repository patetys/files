VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V730B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46v730B
' Data/Hora: 28/11/2008 - 10:30:00
' Sistema: P46V101
' Analista: Sandra M. De Rosa
' ======================================================================================================================================================

Public nosso_numero As String
Public Num_Ped As Long
Public Num_Item As Long
Public Num_Apol_Ant As Double
Public Nom_Segurado As String
Public Num_Tel As String
Public Num_Cgc_Cpf As String
Public Cod_Ramo As Long
Public Cod_Org As Long
Public Dat_Ini_Vig As Double
Public Dat_Ter_Vig As Double
Public Cod_Tip_Logr_Risco As String
Public Nom_Logr_Risco As String
Public Dsc_Compl_Risco As String
Public Nom_Bairro_Risco As String
Public Nom_Cid_Risco As String
Public Sig_UF_Risco As String
Public Num_Cep_Risco As Long
Public Num_Risco As Long
Public Dat_Trans As Double
Public Cod_Iden_Pessoa As String
Public Cod_Unid_Prod As Long
Public Cod_Prod As Long
Public Cod_Corr As Long
Public Cod_Cobert_01 As Long
Public Cod_Cobert_02 As Long
Public Cod_Cobert_03 As Long
Public Cod_Cobert_04 As Long
Public Cod_Cobert_05 As Long
Public Cod_Cobert_06 As Long
Public Cod_Cobert_07 As Long
Public Cod_Cobert_08 As Long
Public Cod_Cobert_09 As Long
Public Cod_Cobert_10 As Long
Public Val_IS_Cobert_01 As Long
Public Val_IS_Cobert_02 As Long
Public Val_IS_Cobert_03 As Long
Public Val_IS_Cobert_04 As Long
Public Val_IS_Cobert_05 As Long
Public Val_IS_Cobert_06 As Long
Public Val_IS_Cobert_07 As Long
Public Val_IS_Cobert_08 As Long
Public Val_IS_Cobert_09 As Long
Public Val_IS_Cobert_10 As Long
Public Cod_Clasf As String
Public Cod_Compl_Clasf As String
Public Pess_Cont As String
Public Cod_Produto As Long
Public Flag_Trans As Long
Public Num_Cotac As String
Public Num_Insp As Long
Public Cod_Status_I As Integer
Public Dat_Confirm As Double
Public Dat_Final As Double
Public Dat_Cancel As Double

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero)
End Property
