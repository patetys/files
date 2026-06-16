VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V727B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ======================================================================================================================================================
' Data/Hora: 12/09/2007 - 11:14:18
' Analista: tjgomes
' ======================================================================================================================================================

' Properties...
' ======================================================================================================================================================
Public NUM_BOLETO       As Long  'PK
Public NOSSO_NUMERO     As String
Public COD_UNID         As Long
Public COD_RAMO         As Long
Public COD_BANCO        As Long
Public NOM_SEGURADO     As String
Public VAL_PR           As Double
Public DAT_VENC         As Long
Public COD_CORR         As Long
Public NUM_LINHA_DIGIT  As String
Public COD_USER         As String
Public DAT_DIGIT        As Long
Public HOR_DIGIT        As Long
Public STATUS           As Long
' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(NUM_BOLETO)
End Property

Public Property Get gDigitoBloq() As String
Dim clsA99V001  As clsA99V001     'DLL: calcula dígitos.
Dim lDigito     As String
Dim lMensagem   As String
Dim lAuxiliar   As String
    Set clsA99V001 = New clsA99V001     'DLL: calcula dígitos.
    lAuxiliar = Format(NUM_BOLETO & "0", "000000000")
    If clsA99V001.DIG_BLOQUETO(lAuxiliar, lDigito, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
    End If
    gDigitoBloq = lDigito
    Set clsA99V001 = Nothing
    
End Property

