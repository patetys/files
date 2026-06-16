VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcConsistPreCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public NumItem              As Long
Public IDConsistencia       As E_IDConsist
Public TipoConsistencia     As E_TipConsist
Public Mensagem             As String
Public RespostaQuest        As E_RespQuest
Public RespostaValidada     As Boolean

Public Enum E_TipConsist
    E_TipConsist_ND = 0
    E_TipConsist_Quest = 1
    E_TipConsist_Aviso = 2
End Enum

Public Enum E_RespQuest
    E_RespQuest_ND = 0
    E_RespQuest_Sim = 1
    E_RespQuest_Nao = 2
    E_RespQuest_OK = 3
End Enum

Public Enum E_IDConsist
    E_IDConsist_ND = 0
    E_IDConsist_DescontoLMI = 1
    E_IDConsist_IsencaoIOF = 2
    E_IDConsist_RelacaoBens = 3
    E_IDConsist_RcOperacoes = 4
    E_IDConsist_EndDivergente = 5
    E_IDConsist_Cob28_Empresarial = 6   '1200004 - Tarifa 1028 e 1029 Janeiro 2012 - Item 2
    E_IDConsist_Cob28_Aviso_Empresarial = 7
    E_IDConsist_Aviso_InspRisco = 100
    E_IDConsist_Aviso_ArqTransmInsp = 101
End Enum

Private Sub Class_Initialize()
    TipoConsistencia = E_TipConsist_ND
    RespostaQuest = E_RespQuest_ND
    RespostaValidada = False
End Sub

Public Property Get Class_ID() As String
    Class_ID = Format(NumItem, "000") & "-" & _
               Format(IDConsistencia, "000")
End Property

