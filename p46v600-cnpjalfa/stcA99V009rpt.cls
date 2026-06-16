VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA99V009rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mstcA99V009B As stcA99V009B

Public Property Set stcOrigem(ByRef pstcA99V009B As stcA99V009B)
    Set mstcA99V009B = pstcA99V009B
End Property

Public Property Get Apolice() As String
    Apolice = mstcA99V009B.getApolice
End Property

Public Property Get endosso() As String
    endosso = mstcA99V009B.getEndosso
End Property

Public Property Get parcelaPagar() As String
    parcelaPagar = mstcA99V009B.getQuantidade
End Property

Public Property Get Emissao() As String
    Emissao = mstcA99V009B.getDataDocumento
End Property

Public Property Get numeroBanco() As String
    numeroBanco = mstcA99V009B.numeroBanco
End Property

Public Property Get nomeBanco() As String
    nomeBanco = mstcA99V009B.nomeBanco
End Property

Public Property Get codBanco() As String
    codBanco = mstcA99V009B.numeroBanco
End Property

Public Property Get agenciaCodigoCedente() As String
    agenciaCodigoCedente = mstcA99V009B.AgenciaContaMask
End Property

Public Property Get vencimento() As String
    vencimento = mstcA99V009B.getVencimento
End Property

Public Property Get solicitacao() As String
    solicitacao = mstcA99V009B.getDataDocumento
End Property

Public Property Get ramo() As String
    ramo = mstcA99V009B.ramo
End Property

Public Property Get linhaDigitavel() As String
    linhaDigitavel = mstcA99V009B.linhaDigitavel
End Property

Public Property Get localPagamento() As String
    localPagamento = mstcA99V009B.getLocalPagamento
End Property

Public Property Get cedente() As String
    cedente = mstcA99V009B.getNomeCedente
End Property

Public Property Get cnpjCedente() As String
    cnpjCedente = mstcA99V009B.getCNPJCedente
End Property

Public Property Get dataDocumento() As String
    dataDocumento = mstcA99V009B.getDataDocumento
End Property

Public Property Get documento() As String
    documento = mstcA99V009B.NumeroDoc
End Property

Public Property Get especieDoc() As String
    especieDoc = mstcA99V009B.getEspecieDocumento
End Property

Public Property Get aceite() As String
    aceite = mstcA99V009B.getAceite
End Property

Public Property Get processamento() As String
    processamento = mstcA99V009B.getDataProcessamento
End Property

Public Property Get nossoNumeroDocumento() As String
    nossoNumeroDocumento = mstcA99V009B.NumeroDoc
End Property

Public Property Get nossoNumeroBanco() As String
    nossoNumeroBanco = mstcA99V009B.getNossoNumeroMask
End Property

Public Property Get usoBanco() As String
    usoBanco = mstcA99V009B.getUsoBanco
End Property

Public Property Get carteira() As String
    carteira = mstcA99V009B.getCarteira
End Property

Public Property Get valorDocumento() As String
    valorDocumento = Format(mstcA99V009B.valorDocumento, "standard")
End Property

Public Property Get instrucoes() As String
    instrucoes = Replace(mstcA99V009B.getInstrucao, "<br>", vbCrLf)
End Property

Public Property Get sacado() As String
    sacado = mstcA99V009B.NomeSacado
End Property

Public Property Get cpfcnpjsacado() As String
    cpfcnpjsacado = mstcA99V009B.CPF_CNPJ_Sacado
End Property

Public Property Get codigoBarras() As String
    codigoBarras = mstcA99V009B.CodigoBarra
End Property
