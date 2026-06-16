VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA99V009B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public aceite As String
Public AgenciaCedente As String
Public AgenciaContaMask As String
Public Apolice As String
Public carteira As String
Public CodConv As String
Public CodigoBarra As String
Public CodigoBarraBinario As String
Public ContaCedente As String
Public Corretor As String
Public dataDocumento As Long
Public DataProcessamento As Long
'Public Desconto             As String
Public DigitoVerificador As Long
Public DigYasNossoNumero As Long
Public DigBanco As String
Public DV1 As Long
Public DV2 As Long
Public DV3 As Long
Public endosso As String
Public EspecieDocumento As String
Public IdentifCIA As String
Public Instrucao As String
Public linhaDigitavel As String
Public LinhaWN As String
Public LinhaWN_Fonte As String
Public LocalImagensBarra As String
Public LocalImagensBanco As String
Public localPagamento As String
Public MensagemDuvida As String
Public MensagemPgto As String
'Public MoraMulta            As String
Public nomeBanco As String
Public NomeCedente As String
Public NomeSacado As String
Public NossoNumero As String
Public NossoNumeroMask As String
Public numeroBanco As String
Public NumeroDoc As String
'Public OutrasDeducoes       As String
'Public OutrosAcrescimos     As String
Public PathImagemBanco As String
Public Quantidade As String
Public ramo As String
Public usoBanco As String
'Public ValorCobrado         As String
'Public ValorCorrigido       As String
Public valorDocumento As Double
Public vencimento As Long
Public Vigencia As String
Public cnpjCedente As String
Public CPF_CNPJ_Sacado As String

Public CarregarDados As Boolean
Public ConexaoBD As ADODB.Connection
Public NomeTabela As String

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A99V009B"
End Property

Public Property Get getVencimento() As String
    Dim lAuxiliar As String
    lAuxiliar = Format(Me.vencimento, "00000000")
    lAuxiliar = Right(lAuxiliar, 2) & "/" & Mid(lAuxiliar, 5, 2) & "/" & Left(lAuxiliar, 4)
    getVencimento = lAuxiliar
End Property
Public Property Get getDataDocumento() As String
    Dim lAuxiliar As String
    If Me.dataDocumento = 0 Then
        Me.dataDocumento = Format(Date, "yyyymmdd")
    End If
    lAuxiliar = Format(Me.dataDocumento, "00000000")
    lAuxiliar = Right(lAuxiliar, 2) & "/" & Mid(lAuxiliar, 5, 2) & "/" & Left(lAuxiliar, 4)
    getDataDocumento = lAuxiliar
End Property

Public Property Get getDataProcessamento() As String
    Dim lAuxiliar As String
    If Me.DataProcessamento = 0 Then
        If Me.numeroBanco = 399 Or val(Me.numeroBanco) = 33 Then 'Boleto Santander
            Me.DataProcessamento = 0
        Else
            Me.DataProcessamento = Format(Date, "yyyymmdd")
        End If
    End If
    If Me.DataProcessamento <> 0 Then
        lAuxiliar = Format(Me.DataProcessamento, "00000000")
        lAuxiliar = Right(lAuxiliar, 2) & "/" & Mid(lAuxiliar, 5, 2) & "/" & Left(lAuxiliar, 4)
    Else
        lAuxiliar = ""
    End If
    getDataProcessamento = lAuxiliar
End Property
Public Property Get getLocalPagamento() As String
'Informações tratadas
    If M46V999.gfPreenchido(Me.localPagamento) = False Then
        Me.localPagamento = "PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO"

        If Me.numeroBanco = 399 Then
            Me.localPagamento = "PAGAR PREFERENCIALMENTE EM AGÊNCIA HSBC"
        ElseIf val(Me.numeroBanco) = 33 Then
            Me.localPagamento = "PAGAR PREFERENCIALMENTE EM AGÊNCIA SANTANDER"
        End If
    End If
    getLocalPagamento = Me.localPagamento
End Property
Public Property Get getInstrucao() As String
    If Not M46V999.gfPreenchido(Me.Instrucao) Then
        Me.Instrucao = "NÃO RECEBER APÓS O VENCIMENTO, CASO ESTA DATA OCORRA NOS FINAIS " & _
                       "DE SEMANA OU FERIADO, RECEBER NO 1º DIA ÚTIL SUBSEQUENTE.<br><br>" & _
                       "EXCLUSIVO PARA PAGAMENTO ANTECIPADO."
    End If
    getInstrucao = Me.Instrucao
End Property
Public Property Get getUsoBanco() As String
    If Not M46V999.gfPreenchido(Me.usoBanco) And Me.numeroBanco = 237 Then
        Me.usoBanco = "CIP - 163"
    End If
    If Not M46V999.gfPreenchido(Me.usoBanco) And Me.numeroBanco = 409 Then
        Me.usoBanco = "CVT 07744-7"
    End If
    getUsoBanco = Me.usoBanco
End Property
Public Property Get getEspecieDocumento() As String
    If Not M46V999.gfPreenchido(Me.EspecieDocumento) Then
        If Me.numeroBanco = 237 Then
            Me.EspecieDocumento = "NS"
        ElseIf Me.numeroBanco = 399 Or val(Me.numeroBanco) = 33 Then 'Boleto Santander
            Me.EspecieDocumento = ""
        Else
            Me.EspecieDocumento = "AP"
        End If
    End If
    getEspecieDocumento = Me.EspecieDocumento
End Property
Public Property Get getAceite() As String
    If Not M46V999.gfPreenchido(Me.aceite) Then
        If Me.numeroBanco = 399 Or val(Me.numeroBanco) = 33 Then 'Boleto Santander
            Me.aceite = ""
        Else
            Me.aceite = "S"
        End If
    End If
    getAceite = Me.aceite
End Property
Public Property Get getQuantidade() As String
    If Not M46V999.gfPreenchido(Me.Quantidade) Then
        Me.Quantidade = "01"
    End If
    getQuantidade = Me.Quantidade
End Property
Public Property Get getNomeCedente() As String
    Dim lNomeSeguradora As String
    If Not M46V999.gfPreenchido(Me.NomeCedente) Then
        ' Obter o nome da Seguradora
        lNomeSeguradora = "Yasuda Maritima"
        If M46V999.gDataCorteSompo <> "" Then
            If IsNumeric(M46V999.gDataCorteSompo) Then
                If CLng(M46V999.gDataCorteSompo) <= M46V999.gHoje Then
                    lNomeSeguradora = "Sompo"
                End If
            End If
        End If
    
        Me.NomeCedente = lNomeSeguradora & " Seguros S.A."
    End If
    getNomeCedente = Me.NomeCedente
End Property
Public Property Get getCNPJCedente() As String
    If Not M46V999.gfPreenchido(Me.cnpjCedente) Then
        'CR54548 - Projeto de Incorporação - Produto Compreensivo
        'Me.cnpjCedente = "60.405.925/0001-44"
        Me.cnpjCedente = "61.383.493/0001-80"
    End If
    getCNPJCedente = Me.cnpjCedente
End Property
Public Property Get getApolice() As String
    If Not M46V999.gfPreenchido(Me.Apolice) Then
        Me.Apolice = "0000000000"
    End If
    getApolice = Me.Apolice
End Property
Public Property Get getEndosso() As String
    If Not M46V999.gfPreenchido(Me.endosso) Then
        Me.endosso = "000000"
    End If
    getEndosso = Me.endosso
End Property
Public Property Get getCarteira() As String
    Dim lAuxiliar As String

    If Me.carteira = "0" Then
        lAuxiliar = ""
    ElseIf IsNumeric(Me.carteira) Then
        lAuxiliar = Format(Me.carteira, "00")
    Else
        lAuxiliar = Me.carteira
    End If
    Me.carteira = lAuxiliar
    If Me.numeroBanco = 399 Or val(Me.numeroBanco) = 33 Then 'Boleto Santander
        Me.carteira = "CNR"
    End If
    getCarteira = Me.carteira
End Property

Public Property Get getNossoNumeroMask() As String
    Dim lDig As String
    Dim lSomaAux As Double
    Select Case Me.numeroBanco
    Case 341
        Me.NossoNumeroMask = Format(Me.carteira, "000") & "/" & Format(Me.NossoNumero, "00000000") & "-" & Me.DigBanco
    Case 237
        Me.DigBanco = mfRetornarDivNumBradesco
        Me.NossoNumeroMask = Format(Me.carteira, "00") & "/" & Format(Me.NossoNumero, "00000000000") & "-" & Me.DigBanco
    Case 356
        Me.AgenciaContaMask = Me.AgenciaCedente & "/" & Me.ContaCedente & "-" & Me.DigBanco
        Me.NossoNumeroMask = "18" & Format(Me.NossoNumero, "00000000000")
    Case 399
        lDig = Modulo_11Base9(Me.NossoNumero)
        lSomaAux = CDbl((Format(Me.NossoNumero, "00000000000000") & lDig & 4)) + Me.ContaCedente + (Mid(Me.vencimento, 7, 2) & Mid(Me.vencimento, 5, 2) & Mid(Me.vencimento, 3, 2))
        Me.NossoNumeroMask = Format(Me.NossoNumero, "00000000000000") & lDig & "4" & Modulo_11Base9(Format(lSomaAux, "0000000000000000"))
    Case Else
        Me.NossoNumeroMask = Format(Me.NossoNumero, "00000000000000") & "-" & Me.DigYasNossoNumero
    End Select
    getNossoNumeroMask = Me.NossoNumeroMask
End Property
Private Function mfRetornarDivNumBradesco() As String
    Dim lAuxiliar As String

    lAuxiliar = Format(Me.carteira, "00")
    lAuxiliar = lAuxiliar & Format(Me.NossoNumero, "00000000000")
    mfRetornarDivNumBradesco = Modulo_11(lAuxiliar)

End Function
Private Function Modulo_11Base9(pEntrada As String) As String

    Dim lPesos As Integer
    Dim lTotal As Integer
    Dim contador As Integer
    Dim lResto As Integer
    Dim lResultado As Integer
    Dim lI As Integer
    Dim digito As String
    Dim lAuxiliar As String
    lPesos = 9             '923456789234...'
    lTotal = 0
    digito = "0"

    lAuxiliar = pEntrada

    For lI = Len(pEntrada) To 1 Step -1
        lResultado = lResultado + (CByte(Mid(lAuxiliar, lI, 1)) * lPesos)
        lPesos = lPesos - 1
        If lPesos = 1 Then
            lPesos = 9
        End If
    Next lI


    '4. Divide por 11 e obtém resto.
    lResto = lResultado Mod 11


    '5. Posiciona dígito.
    Select Case lResto
    Case 10
        Modulo_11Base9 = "0"
    Case Else
        Modulo_11Base9 = lResto
    End Select


End Function

Private Function Modulo_11(pEntrada As String) As String

    Dim lPesos As Integer
    Dim lTotal As Integer
    Dim contador As Integer
    Dim lResto As Integer
    Dim lResultado As Integer
    Dim lI As Integer
    Dim digito As String
    Dim lAuxiliar As String
    lPesos = 2             '2765432765432
    lTotal = 0
    digito = "0"

    lAuxiliar = pEntrada

    For lI = 1 To Len(pEntrada)
        lResultado = lResultado + (CByte(Mid(lAuxiliar, lI, 1)) * lPesos)
        lPesos = lPesos - 1
        If lPesos = 1 Then
            lPesos = 7
        End If

    Next lI


    '4. Divide por 11 e obtém resto.
    lResto = lResultado Mod 11

    lResto = 11 - lResto

    '5. Posiciona dígito.
    Select Case lResto
    Case 10
        Modulo_11 = "P"
    Case 11
        Modulo_11 = "0"
    Case Else
        Modulo_11 = lResto
    End Select
End Function

Public Function getRpt() As stcA99V009rpt
    Static lrpt As stcA99V009rpt

    If lrpt Is Nothing Then
        Set lrpt = New stcA99V009rpt
        Set lrpt.stcOrigem = Me
    End If

    Set getRpt = lrpt

End Function

Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
End Property

Public Property Get M46V999() As clsM46V999
    If mClsM46V999 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V999 não atribuído para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V999 já
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46V999 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46V999 = mClsM46V999
End Property

Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
End Sub
