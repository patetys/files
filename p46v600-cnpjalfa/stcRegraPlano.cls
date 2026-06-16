VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcRegraPlano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public CodRegra As eRegraPlano
Public ValRegra As Double


Public Enum eRegraPlano
    eRegraPlano_ND = 0
    '-----
    eRegraPlano_ValorRisco = 1
    eRegraPlano_Unidade = 2
    eRegraPlano_Corretor = 3
    eRegraPlano_Comissao = 4
    eRegraPlano_DescontoComercial = 5
    eRegraPlano_DescontoProgressivo = 6
    eRegraPlano_ParcelaMinima = 7
    eRegraPlano_CustoApolice = 8
    eRegraPlano_QtdeParcelas = 9
    eRegraPlano_CodigoSegurado = 10
    eRegraPlano_PremioMinimoPlano = 11
    eRegraPlano_CEPRisco = 12
    eRegraPlano_NumeroLogradouro = 13
    '-----
    eRegraPlano_1aParcelaVista = 20
    eRegraPlano_1aParcela30Dias = 21
    eRegraPlano_2aParcelaVista = 22
    eRegraPlano_2aParcela30Dias = 23
    eRegraPlano_3aParcelaVista = 24
    eRegraPlano_3aParcela30Dias = 25
    eRegraPlano_4aParcelaVista = 26
    eRegraPlano_4aParcela30Dias = 27
    eRegraPlano_5aParcelaVista = 28
    eRegraPlano_5aParcela30Dias = 29
    eRegraPlano_6aParcelaVista = 30
    eRegraPlano_6aParcela30Dias = 31
    eRegraPlano_7aParcelaVista = 32
    eRegraPlano_7aParcela30Dias = 33
    eRegraPlano_8aParcelaVista = 34
    eRegraPlano_8aParcela30Dias = 35
    eRegraPlano_9aParcelaVista = 36
    eRegraPlano_9aParcela30Dias = 37
    eRegraPlano_10aParcelaVista = 38
    eRegraPlano_10aParcela30Dias = 39
    eRegraPlano_11aParcelaVista = 40
    eRegraPlano_11aParcela30Dias = 41
    '----
    eRegraPlano_PremioMinimoCobertura = 50
    eRegraPlano_PremioMinimoEndosso = 51
    eRegraPlano_ParcelaMinimaApolAntecipada = 52
    eRegraPlano_ParcelaMinimaApolPostecipada = 53
    eRegraPlano_ParcelaMinimaEndosso = 54
End Enum

Public Property Get Class_ID() As String
    Class_ID = CStr(CodRegra)
End Property
