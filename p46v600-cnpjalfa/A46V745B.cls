VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V745B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V745B - Log_Tab_Ped_Cobert_Batente
' Data/Hora: 08/04/2016
' Sistema: Multi
' Rodrigo P S Rogério
' ======================================================================================================================================================

Public nosso_numero             As String           'Nosso número.
Public Cod_Cobert               As Integer          'Código da cobertura.
Public Num_Item                 As Long             'Número do item.
Public Dat_Calculo              As Long             'Data do Calculo
Public Usuario                  As String           'Usuário
Public Status_Batente           As eStatusBatente   'Status do Batente
Public Val_Pr_Base_Apol         As Currency         'Base de calculo do Batente
Public Val_Pr_Pago_Apol         As Currency         'Valor Pago pelo segurado
Public Formula_Apol             As String           'Formula
Public Per_Dif_Batente          As Double           'Percentual de limite do batente utilizada
Public Per_Dif_Renovacao        As Double           'Percentual de diferença apurada
Public Val_Pr_Calc              As Currency         'Valor do Premio calculada
Public Formula                  As String           'Formula

' ======================================================================================================================================================

Public Enum eStatusBatente     'Status Batente
    eStatusBatente_Aplicado = 1
    eStatusBatente_LimiteInferior = 2
    eStatusBatente_ApolicePossuiEndosso = 3
    eStatusBatente_QtdItensDiferente = 4
    eStatusBatente_CEPEnderecoRiscoDiferente = 5
    eStatusBatente_AtividadeDiferente = 6
    eStatusBatente_ComissaoDiferente = 7
    eStatusBatente_SuprimirBatente = 8
End Enum

Public Property Get Class_Name() As String
    Class_Name = "A46V745B"
End Property
