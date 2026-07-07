VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V102A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private clsA99V001 As clsA99V001    'Cálculo de dígitos.
Public Pro_CarBanco As Integer    'Pagamento com carnê: código do banco.
Public Pro_CarDatPriParcela As Double   'Pagamento com carnê: data de vencimento da primeira
'parcela do carnê - formato aaaammdd.
Public Pro_Cob_Bairro As String    'Proposta: endereço de cobrança - bairro.
Public Pro_Cob_CEP As String    'Proposta: endereço de cobrança - CEP.
Public Pro_Cob_Cidade As String    'Proposta: endereço de cobrança - cidade.
Public Pro_Cob_Complemento As String  'Proposta: endereço de cobrança - complemento.
Public Pro_Cob_NomLogradouro As String    'Proposta: endereço de cobrança - nome do logradouro.
Public Pro_Cob_TipLogradouro As String    'Proposta: endereço de cobrança - tipo do logradouro.
Public Pro_Cob_UF As String    'Proposta: endereço de cobrança - unidade da federação.
Public Pro_CodCorretor As Long    'Proposta: código do corretor.
Public Pro_Cosseguros As Variant    'Proposta: cosseguros.
Public Pro_DebAgencia As Long    'Proposta: débito em conta - agência.
Public Pro_DebBanco As Integer    'Proposta: débito em conta - banco.
Public Pro_DebConta As Double    'Proposta: débito em conta - conta.
Public Pro_End_Bairro As String    'Proposta: endereço do segurado - bairro.
Public Pro_End_CEP As String    'Proposta: endereço do segurado - CEP.
Public Pro_End_Cidade As String    'Proposta: endereço do segurado - cidade.
Public Pro_End_Complemento As String  'Proposta: endereço do segurado - complemento.
Public Pro_End_NomLogradouro As String    'Proposta: endereço do segurado - nome do logradouro.
Public Pro_End_TipLogradouro As String    'Proposta: endereço do segurado - tipo do logradouro.
Public Pro_End_UF As String    'Proposta: endereço do segurado - unidade da federação.
Public Pro_ForPagamento As Byte    'Proposta: forma de pagamento.
'1 = carnê e 2 = débito em conta.
Public Pro_IndAgencia As Long    'Proposta: indicador do seguro - agência.
Public Pro_IndBanco As Integer    'Proposta: indicador do seguro - banco.
Public Pro_IndConta As Double    'Proposta: indicador do seguro - conta.
Public Pro_IndCobranca As Byte    'Proposta: indicador do endereço de cobrança diferente.
Public Pro_IndCosseguros As Byte  'Proposta: indicador de cessão de cosseguro.
Public Pro_NomSegurado As String    'Proposta: nome do segurado.
Public Pro_PagAntBloqueto As Double    'Proposta: pagamento antecipado - bloqueto.
Public Pro_PagAntecipado As Byte  'Proposta: indicador de pagamento antecipado.
Public Pro_UniProdutiva As Integer    'Proposta: unidade produtiva.
Public Pro_VigInicio As Double    'Proposta: vigência - início - formato aaaammdd.

Public Pro_TipPessoa As Byte    'Proposta: tipo de pessoa.
Public Pro_CNPJ_CPF As Double    'Proposta: CNPJ/CPF.
Public Pro_CodSegurado As Long
Public Pro_QtdParcelas As Integer
Public Pro_Cod_Reg As Integer
Public Pro_CreAgencia As Long    'Proposta: débito em conta - agência.
Public Pro_CreBanco As Integer    'Proposta: débito em conta - banco.
Public Pro_CreConta As Double    'Proposta: débito em conta - conta.
Public Pro_Tip_Res_Cre As Integer

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V102A"
End Property

Public Sub VER_COMUM_NIVEL_2()
'Procedure: consiste dados comuns nível 2.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lCodSegurado As String    'Código do segurado.
    Dim lDigito As String    'Dígito.
    Dim lEncontrou As Boolean    'Indica se encontrou registro.
    Dim lEndereco As String    'Endereço completo.
    Dim lI1 As Byte             'Utilizado com For...Next.
    Dim lI2 As Byte             'Utilizado com For...Next.
    Dim lMensagem As String    'Mensagem.
    Dim lQuaCosseguros As Byte    'Quantidade de cosseguros.
    Dim lQuaDias As Double    'Quantidade de dias.
    Dim lTamanho As Integer    'Tamanho.
    Dim lTotPercentuais As Double   'Total do percentuais.
    Dim lBanco As Integer
    Dim lAgencia As Long
    Dim lConta As Long

    '11.32 Verifica Black list
    If M46V999.gPrograma = "D046O003" Then
        Dim lBlackList As Boolean
        
        lMensagem = Empty
        lBlackList = M46V999.Is_BlackList(M46V999.gCodProduto, M46V999.gCod_Ramo, Pro_TipPessoa, Pro_CNPJ_CPF, lMensagem)
        
        If M46V999.gfPreenchido(lMensagem) Then
            Call M46V999.gpGraLog(0, lMensagem)
            Call M46V999.gpMenErro(lMensagem)
        Else
            If lBlackList Then
                Call M46V999.gpMenErro("FGE0036 - Por motivos técnicos não é possível prosseguir com este calculo.")
            End If
        End If
    End If
    'FIM - 11.32 Verifica Black list
    
'3. Endereço do segurado.
    If Not M46V999.gfPreenchido(Pro_End_TipLogradouro) Then
        Call M46V999.gpMenErro("FGE0037 - Endereço do segurado inválido (tipo do logradouro não preenchido).")
    End If
    If Not M46V999.gfPreenchido(Pro_End_NomLogradouro) Then
        Call M46V999.gpMenErro("FGE0038 - Endereço do segurado inválido (nome do logradouro não preenchido).")
    End If
    If Not M46V999.gfPreenchido(Pro_End_Complemento) Then
        Call M46V999.gpMenErro("FGE0108 - Complemento do endereço do segurado inválido (Complemento do logradouro não preenchido).")
    End If
    Pro_End_TipLogradouro = Trim$(Pro_End_TipLogradouro)
    Pro_End_NomLogradouro = Trim$(Pro_End_NomLogradouro)
    Pro_End_Complemento = Trim$(Pro_End_Complemento)
    lEndereco = Pro_End_TipLogradouro & " " & Pro_End_NomLogradouro
    If M46V999.gfPreenchido(Pro_End_Complemento) Then
        lEndereco = lEndereco & " " & Pro_End_Complemento
    End If
    'lTamanho = Len(lEndereco)
    lTamanho = 0
    If lTamanho > 50 Then
        Call M46V999.gpMenErro("FGE0039 - Endereço do segurado inválido (endereço completo com " & _
                               lTamanho & _
                             " posições - tamanho máximo permitido no sistema de 50 posições)." & _
                               vbLf & vbLf & UCase(lEndereco))
    End If
    If Not M46V999.gfPreenchido(Pro_End_Bairro) Then
        Call M46V999.gpMenErro("FGE0040 - Endereço do segurado inválido (bairro não preenchido).")
    End If
    If Not M46V999.gfPreenchido(Pro_End_Cidade) Then
        Call M46V999.gpMenErro("FGE0041 - Endereço do segurado inválido (cidade não " & _
                               "preenchida).")
    End If
    If Not M46V999.gfPreenchido(Pro_End_UF) Then
        Call M46V999.gpMenErro("FGE0042 - Endereço do segurado inválido (unidade da " & _
                               "federação não preenchida).")
    End If
    If Not M46V999.gfPreenchido(Pro_End_CEP) Then
        Call M46V999.gpMenErro("FGE0043 - Endereço do segurado inválido (CEP não " & _
                               "preenchido).")
    End If
    If val(Pro_End_CEP) = 0 Then
        Call M46V999.gpMenErro("FGE0044 - Endereço do segurado inválido (CEP " & _
                               "zerado).")
    End If
    If Len(Trim(Pro_End_CEP)) <> 8 Then
        Call M46V999.gpMenErro("FGE0045 - Endereço do segurado inválido " & _
                               "(CEP com tamanho diferente de 8 posições).")
    End If

    Pro_End_Bairro = Trim$(Pro_End_Bairro)
    Pro_End_Cidade = Trim$(Pro_End_Cidade)

    '4. Endereço de cobrança.
    If Pro_IndCobranca = 1 Then     'Indica endereço de cobrança diferente posicionado.
        If Not M46V999.gfPreenchido(Pro_Cob_TipLogradouro) And Mid(Pro_Cob_NomLogradouro, 1, 4) <> "FIN-" Then
            Call M46V999.gpMenErro("FGE0046 - Endereço de cobrança inválido (tipo do logradouro não preenchido).")
        End If
        If Not M46V999.gfPreenchido(Pro_Cob_NomLogradouro) Then
            Call M46V999.gpMenErro("FGE0047 - Endereço de cobrança inválido (nome do logradouro não " & _
                                   "preenchido).")
        End If
        If Not M46V999.gfPreenchido(Pro_End_Complemento) Then
            Call M46V999.gpMenErro("FGE0109 - Complemento do endereço de cobrança inválido (Complemento do logradouro não preenchido).")
        End If
        Pro_Cob_TipLogradouro = Trim$(Pro_Cob_TipLogradouro)
        Pro_Cob_NomLogradouro = Trim$(Pro_Cob_NomLogradouro)
        Pro_Cob_Complemento = Trim$(Pro_Cob_Complemento)
        lEndereco = Pro_Cob_TipLogradouro & " " & Pro_Cob_NomLogradouro
        If M46V999.gfPreenchido(Pro_Cob_Complemento) Then
            lEndereco = lEndereco & " " & Pro_Cob_Complemento
        End If
        'lTamanho = Len(lEndereco)
        lTamanho = 0
        If lTamanho > 50 Then
            Call M46V999.gpMenErro("FGE0048 - Endereço de cobrança inválido (endereço completo " & _
                                   "com " & lTamanho & " posições - tamanho máximo permitido no " & _
                                   "sistema de 50 posições)." & vbLf & vbLf & UCase(lEndereco))
        End If
        If Not M46V999.gfPreenchido(Pro_Cob_Bairro) Then
            Call M46V999.gpMenErro("FGE0049 - Endereço de cobrança inválido (bairro não " & _
                                   "preenchido).")
        End If
        If Not M46V999.gfPreenchido(Pro_Cob_Cidade) Then
            Call M46V999.gpMenErro("FGE0050 - Endereço de cobrança inválido (cidade não " & _
                                   "preenchida).")
        End If
        If Not M46V999.gfPreenchido(Pro_Cob_UF) Then
            Call M46V999.gpMenErro("FGE0051 - Endereço de cobrança inválido (unidade " & _
                                   "da federação não preenchida).")
        End If
        If Not M46V999.gfPreenchido(Pro_Cob_CEP) Then
            Call M46V999.gpMenErro("FGE0056 - Endereço de cobrança inválido (CEP " & _
                                   "não preenchido).")
        End If
        If val(Pro_Cob_CEP) = 0 Then
            Call M46V999.gpMenErro("FGE0057 - Endereço de cobrança inválido " & _
                                   "(CEP zerado).")
        End If
        If Len(Trim(Pro_Cob_CEP)) <> 8 Then
            Call M46V999.gpMenErro("FGE0058 - Endereço de cobrança " & _
                                   "inválido (CEP com tamanho diferente " & _
                                   "de 8 posições).")
        End If
    End If
    Pro_Cob_Bairro = Trim$(Pro_Cob_Bairro)
    Pro_Cob_Cidade = Trim$(Pro_Cob_Cidade)

    'Ficha 900591
    M46V999.gPodeEmitir = True
    
    If Not (Pro_Cod_Reg = 106 Or Pro_Cod_Reg = 107) Then
        lBanco = Pro_DebBanco
        lAgencia = Pro_DebAgencia
        lConta = Pro_DebConta
    
        If Pro_ForPagamento > 1 And Pro_ForPagamento <> 4 Then '2 = débito em conta. 3 = débito com primeira em boleto
            If lBanco <> 0 And Pro_Cod_Reg <> 129 Then
                    If mfVerConta(1, lBanco, lAgencia, lConta, lMensagem) = False Then
                        Call M46V999.gpMenErro(lMensagem)
                    End If
            Else
                If lBanco = 0 Then  'DebConta
                    Call M46V999.gpMenErro("FGE0073 - É necessário informar os dados de Debito em conta.")
                End If
            End If
        End If
        
    End If

    If Pro_Tip_Res_Cre = 1 And Pro_ForPagamento <> 4 Then
            If Pro_CreBanco <> 0 Then
                lBanco = Pro_CreBanco
                lAgencia = Pro_CreAgencia
                lConta = Pro_CreConta
                If lAgencia <> 0 And Pro_Cod_Reg <> 129 Then
                    If mfVerConta(3, lBanco, lAgencia, lConta, lMensagem) = False Then
                        Call M46V999.gpMenErro(lMensagem)
                    End If
                End If
        Else
            Call M46V999.gpMenErro("FGE0073 - É necessário informar os dados conta de restituição.")
        End If
    End If

    '1001266 - Validar CPF e CNPJ na Cotação e Proposta
    '1. e 2. CNPJ/CPF.
    '    Select Case Pro_TipPessoa
    '        Case 1      'Pessoa jurídica - trata-se de CNPJ.
    '            If IsNull(Pro_CNPJ_CPF) Then
    '                Call M46V999.gpMenErro("FNP0040 - CNPJ do segurado não preenchido.")
    '            Else
    '                If Pro_CNPJ_CPF = 0 Then
    '                    Call M46V999.gpMenErro("FNP0040 - CNPJ do segurado não preenchido.")
    '                Else
    '                    lAuxiliar = Format$(Pro_CNPJ_CPF, "00000000000000")
    '                    lAuxiliar = Mid$(lAuxiliar, 1, 2) & "." & Mid$(lAuxiliar, 3, 3) & "." & _
                         '                                Mid$(lAuxiliar, 6, 3) & "/" & Mid$(lAuxiliar, 9, 4) & "-" & _
                         '                                Mid$(lAuxiliar, 13, 2)
    '                    Set clsA99V001 = New clsA99V001     'Calcula dígitos.
    '                    Set clsA99V001.M46V999 = M46V999
    '                    If clsA99V001.DIG_CNPJ(lAuxiliar, lDigito, lMensagem) = False Then
    '                        Call M46V999.gpMenErro(lMensagem)
    '                    End If
    '                    Set clsA99V001 = Nothing
    '                End If
    '            End If
    '        Case 2      'Pessoa física - trata-se de CPF.
    '            If IsNull(Pro_CNPJ_CPF) Then
    '                Call M46V999.gpMenErro("FNP0034 - CPF do segurado não preenchido.")
    '            Else
    '                If Pro_CNPJ_CPF = 0 Then
    '                    Call M46V999.gpMenErro("FNP0034 - CPF do segurado não preenchido.")
    '                Else
    '                    lAuxiliar = Format$(Pro_CNPJ_CPF, "00000000000")
    '                    lAuxiliar = Mid$(lAuxiliar, 1, 3) & "." & Mid$(lAuxiliar, 4, 3) & "." & _
                         '                                Mid$(lAuxiliar, 7, 3) & "-" & Mid$(lAuxiliar, 10, 2)
    '                    Set clsA99V001 = New clsA99V001     'Calcula dígitos.
    '                    Set clsA99V001.M46V999 = M46V999
    '                    If clsA99V001.DIG_CPF(lAuxiliar, lDigito, lMensagem) = False Then
    '                        Call M46V999.gpMenErro(lMensagem)
    '                    End If
    '                    Set clsA99V001 = Nothing
    '                End If
    '            End If
    '    End Select

    If (Pro_CarBanco = 409 Or Pro_DebBanco = 409) And M46V999.gTipoAcesso = CALCULO Then
        Call M46V999.gpMenErro("FGE0144 - Não será possível efetuar pagamento com a opção de banco Unibanco (Cod. 409). Favor selecionar outro banco.")
    End If

    '1201070 - Tarifa - 1050 e 1051 - Item 1 - Planos TGP
    If M46V999.gCod_Ramo = e_CodRamo_Residencial And _
       M46V999.M46V111.gCod_plano >= 2278 And _
       M46V999.M46V111.gCod_plano <= 2297 Then
        If Pro_CarBanco = e_Banco_ND Or Pro_DebAgencia = 0 Or Pro_DebConta = 0 Then
            Call M46V999.gpMenErro("Para Plano TGP, o pagamento só pode ser Débito em Conta Corrente.")
        End If
    End If

End Sub

Private Function mfVerBoston(ByVal pAgencia As Long, ByVal pConta As Double, ByRef pMensagem As String) _
        As Boolean
'Função: verifica informações de conta Boston.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetro de saída......pMensagem...Mensagem.

    Dim lAgencia As String    'Código da agência.
    Dim lConta As String  'Número da conta.
    Dim lDigito As String   'Dígito.

    mfVerBoston = False

    '1. Verifica agência.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                  " - Banco de Boston - mais que 4 posições)."
        Exit Function
    End If
    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    If pConta > 99999999 Then
        pMensagem = "FGE2076 - Número de conta corrente inválido (" & pConta & _
                  " - Banco de Boston - mais que 8 posições)."
        Exit Function
    End If
    lConta = Format(pConta, "00000000")

    '3. Calcula e verifica dígito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_BOSTON(Left$(lConta, 7), lDigito, pMensagem) = False Then
        Exit Function
    End If
    Set clsA99V001 = Nothing
    If lDigito <> Mid$(lConta, 8, 1) Then
        pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco de Boston - " & pConta & ")."
        'Ficha 900591
        '1001180 - Validação de DV na Triagem
        '        If M46V999.gTipoAcesso = DIGITACAO Then
        '            MsgBox pMensagem, vbCritical
        '            M46V999.gPodeEmitir = False
        '            mfVerBoston = True
        Exit Function
        '        Else
        '            Exit Function
        '        End If
    End If

    mfVerBoston = True
End Function
Private Function mfVerBradesco(ByVal pAgencia As Long, ByVal pConta As Double, _
                               ByRef pMensagem As String, Optional blnVerificarAgencia As Boolean = True) As Boolean
'Função: verifica informações de conta Bradesco.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetro de saída......pMensagem...Mensagem.

    Dim lAgencia As String    'Código da agência.
    Dim lConta As String  'Número da conta.
    Dim lDigito As String   'Dígito.

    mfVerBradesco = False
    
    If blnVerificarAgencia Then
        '1. Verifica agência.
        If pAgencia > 99999 Then
            pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                      " - Banco Bradescon - mais que 5 posições)."
            Exit Function
        End If
        lAgencia = Format$(pAgencia, "00000")
        Set clsA99V001 = New clsA99V001
        Set clsA99V001.M46V999 = M46V999
    
        If clsA99V001.DIG_BRADESCO(Left$(lAgencia, 4), lDigito, pMensagem) = False Then
            Exit Function
        End If
        Set clsA99V001 = Nothing
        If lDigito <> Mid$(lAgencia, 5, 1) Then
            pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                      " - Banco Bradesco - dígito inválido)."
            Exit Function
        End If
    End If

    '2. Verifica conta.
    If pConta > 99999999 Then
        pMensagem = "FGE2076 - Número de conta corrente inválido (" & pConta & _
                  " - Banco Bradesco - mais que 8 posições)."
        Exit Function
    End If
    lConta = Format(pConta, "00000000")
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_BRADESCO(Left$(lConta, 7), lDigito, pMensagem) = False Then
        Exit Function
    End If
    Set clsA99V001 = Nothing
    If lDigito <> Mid$(lConta, 8, 1) Then
        pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco Bradesco - " & pConta & ")."
        'Ficha 900591
        '1001180 - Validação de DV na Triagem
        '        If M46V999.gTipoAcesso = DIGITACAO Then
        '            MsgBox pMensagem, vbCritical
        '            M46V999.gPodeEmitir = False
        '            mfVerBradesco = True
        Exit Function
        '        Else
        '            Exit Function
        '        End If
    End If

    mfVerBradesco = True
End Function
Public Function mfVerConta(ByVal pOpcao As Byte, ByVal pBanco As Integer, ByVal pAgencia As Long, _
                           ByVal pConta As Double, ByRef pMensagem As String) As Boolean
'Função: verifica informações de conta corrente.

'Parâmetros de entrada...pOpcao......1 = Conta para débito.
'                                    2 = Indicador de seguro.
'                                    3 = Conta de restituição
'                        pBanco......Código do banco.
'                        pAgencia....Código da agência.
'                        pConta......Código da conta.
'Parâmetro de saída......pMensagem...Mensagem.

    mfVerConta = False

    '1. Verifica opção.
    If Not M46V999.gfPreenchido(pOpcao) Then
        pMensagem = "FTE0006 - Opção inválida (verifica conta corrente - opção não preenchida)."
        Exit Function
    End If
    If Not IsNumeric(pOpcao) Then
        pMensagem = "FTE0006 - Opção inválida (verifica conta corrente - opção não numérica)."
        Exit Function
    End If
    If pOpcao <> 1 And pOpcao <> 2 And pOpcao <> 3 Then
        pMensagem = "FTE0006 - Opção inválida (verifica conta corrente - opção diferente de 1, 2 e 3: " & _
                    pOpcao & ")."
        Exit Function
    End If

    '2. Verifica banco.
    If pBanco = 0 Then
        Select Case pOpcao
        Case 1  'Conta para débito.
            pMensagem = "FGE0095 - Código do banco inválido (conta para débito - banco não " & _
                        "preenchido)."
        Case 2  'Indicador do seguro.
            pMensagem = "FGE0095 - Código do banco inválido (conta do indicador de seguro - " & _
                        "banco não preenchido )."
        Case 3  'conta de restituição.
            pMensagem = "FGE0095 - Código do banco inválido (conta de restituição - " & _
                        "banco não preenchido )."
        End Select
        Exit Function
    End If
    
    Dim codigo As Integer, Nome As String, nomeAbreviado As String, isCarne As Boolean, isDebito As Boolean, isRestituicao As Boolean
    pMensagem = Empty
    
    If (M46V999.GetBancoByID(pBanco, codigo, Nome, nomeAbreviado, isCarne, isDebito, isRestituicao, pMensagem) = False Or codigo = 0) Then
        pMensagem = "FGE0095 - Código do banco inválido. " & pMensagem
        Exit Function
    End If
    
    If pOpcao = 1 Then  'Conta para débito.
        ''20170116: original
        'Select Case pBanco
        '    'Projeto 11080 - Debito automatico - Banco do Brasil
        'Case 1, 237, 341, 409, 479, 356, 399, 33, 47, 422, 756    '09029
        '    'Nada a fazer.
        'Case Else
        '    pMensagem = "FGE0095 - Código do banco inválido (conta para débito - " & pBanco & ")."
        '    Exit Function
        'End Select
        ''20170116: fim
        If codigo = 0 Or isDebito = False Then
            pMensagem = "FGE0095 - Código do banco inválido para conta débito (" & pBanco & ")."
            Exit Function
        End If
    End If

    If pOpcao = 3 Then  'Conta para restituição.
        ''20170116: original já comentado
        '        Select Case pBanco
        '        Case 1, 104, 748, 399, 756, 3, 394, 320, 47, 289, 634, 21, 604, 477, 376, 263, 745, 641, 29, 422, 479, 456, 464, 341, 237, 653, 370, 611, 707, 487, 610, 33, 41
        '            'Nada a fazer.
        '        Case Else
        '            pMensagem = "FGE0095 - Código do banco inválido (conta de restituição - " & pBanco & ")."
        '            Exit Function
        '        End Select
        ''20170116: fim
        If codigo = 0 Or isRestituicao = False Then
            pMensagem = "FGE0095 - Código do banco inválido para conta restituição (" & pBanco & ")."
            Exit Function
        End If
    End If


    '3. Verifica agência.
    If IsNull(pAgencia) Then
        Select Case pOpcao
        Case 1      'Conta para débito.
            pMensagem = "FGE0073 - Código de agência inválido (conta para débito - não preenchido)."
        Case 2      'Indicador de seguro.
            pMensagem = "FNP0073 - Código de agência inválido (conta do indicador do seguro - não preenchido)."
        Case 3      'conta de restituição.
            pMensagem = "FNP0073 - Código de agência inválido (conta de restituição - não preenchido)."
        End Select
        Exit Function
    End If
    
    If pAgencia = 0 Then
        Select Case pOpcao
        Case 1      'Conta para débito.
            pMensagem = "FGE0073 - Código de agência inválido (conta para débito - não preenchido)."
        Case 2      'Indicador de seguro.
            pMensagem = "FGE0073 - Código de agência inválido (conta do indicador do seguro - não preenchido)."
        Case 3      'conta de restituição.
            pMensagem = "FGE0073 - Código de agência inválido (conta de restituição - não preenchido)."
        End Select
        Exit Function
    End If

    '3. Verifica conta.
    If IsNull(pConta) Then
        Select Case pOpcao
        Case 1      'Conta para débito.
            pMensagem = "FNP2002 - Número da conta corrente não preenchido (conta para débito)."
        Case 2      'Indicador de seguro.
            pMensagem = "FNP2002 - Número da conta corrente não preenchido (conta do indicador do seguro)."
        Case 3      'conta de restituição.
            pMensagem = "FNP2002 - Número da conta corrente não preenchido (conta de restituição)."
        End Select
        Exit Function
    End If
    If pConta = 0 Then
        Select Case pOpcao
        Case 1      'Conta para débito.
            pMensagem = "FNP2002 - Número da conta corrente não preenchido (conta para débito)."
        Case 2      'Indicador de seguro.
            pMensagem = "FNP2002 - Número da conta corrente não preenchido (conta do indicador do seguro)."
        Case 3      'conta de restituição.
            pMensagem = "FNP2002 - Número da conta corrente não preenchido (conta de restituição)."
        End Select
        Exit Function
    End If

    '4. Verificações específicas por banco.
    Select Case pBanco
    Case 1  'Projeto 11080 - Debito automatico - Banco do Brasil
        If mfVerBancoBrasil(pAgencia, pConta, pMensagem, pOpcao <> 3) = False Then
            Exit Function
        End If
    Case 237    'Bradesco.
        If mfVerBradesco(pAgencia, pConta, pMensagem, pOpcao <> 3) = False Then
            Exit Function
        End If
    Case 341    'Itaú.
        If mfVerItau(pAgencia, pConta, pMensagem) = False Then
            Exit Function
        End If
    Case 347    'Sudameris.
        If mfVerSudameris(pAgencia, pConta, pMensagem) = False Then
            Exit Function
        End If
    Case 409    'Unibanco.
        If mfVerUnibanco(pAgencia, pConta, pMensagem) = False Then
            Exit Function
        End If
    Case 479    'Boston.
        If mfVerBoston(pAgencia, pConta, pMensagem) = False Then
            Exit Function
        End If
    Case 356    'Real
        If mfVerReal(pAgencia, pConta, pMensagem) = False Then
            Exit Function
        End If
    Case 399
        If mfVerHSBC(pAgencia, pConta, pMensagem) = False Then
            Exit Function
        End If
    Case 33
        If mfVerSantander(pAgencia, pConta, pMensagem) = False Then
            Exit Function
        End If
    Case 47
        If mfVerBanese(pAgencia, pConta, pMensagem) = False Then
            Exit Function
        End If
    'Nova Tarifação
    Case 422    'Safra
        If mfVerSAFRA(pAgencia, pConta, pMensagem) = False Then
            Exit Function
        End If
    Case 756    'BANCOOB
        If mfVerBANCOOB(pAgencia, pConta, pMensagem) = False Then
            Exit Function
        End If
    'Novo Convênio banco BRB
    Case 70    'BANCOBRB
        If mfVerBANCOBRB(pAgencia, pConta, pMensagem) = False Then
            Exit Function
        End If
    Case Else
        ''If pOpcao <> 3 Then
        If codigo = 0 Then      ''Codigo=0 : nao achou banco
            Exit Function
        End If
    End Select

    mfVerConta = True
End Function

Private Function mfVerHSBC(ByVal pAgencia As Long, ByVal pConta As Double, ByRef pMensagem As String) _
        As Boolean
'Função: verifica informações de conta HSBC

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetro de saída......pMensagem...Mensagem.

    Dim lAgencia As String    'Código da agência.
    Dim lConta As String  'Número da conta.
    Dim lDigito As String   'Dígito.

    mfVerHSBC = False

    '1. Verifica agência.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                  " - Banco HSBC- mais que 4 posições)."
        Exit Function
    End If

    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    If pConta > 99999999999# Then
        pMensagem = "FGE2076 - Número de conta corrente inválido (" & pConta & _
                  " - Banco HSBC - mais que 11 posições)."
        Exit Function
    End If

    lConta = Format(pConta, "00000000000")

    '3. Calcula e verifica dígito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_HSBC(lAgencia, Left$(lConta, 10), lDigito, pMensagem) = False Then
        Exit Function
    End If

    Set clsA99V001 = Nothing

    If lDigito <> Mid$(lConta, 11, 1) Then
        pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco HSBC- " & pConta & ")."
        'Ficha 900591
        '1001180 - Validação de DV na Triagem
        '        If M46V999.gTipoAcesso = DIGITACAO Then
        '            MsgBox pMensagem, vbCritical
        '            M46V999.gPodeEmitir = False
        '            mfVerHSBC = True
        Exit Function
        '        Else
        '            Exit Function
        '        End If
    End If

    mfVerHSBC = True
End Function


Private Function mfVerItau(ByVal pAgencia As Long, ByVal pConta As Double, ByRef pMensagem As String) _
        As Boolean
'Função: verifica informações de conta Itaú.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetro de saída......pMensagem...Mensagem.

    Dim lAgencia As String    'Código da agência.
    Dim lConta As String  'Número da conta.
    Dim lDigito As String   'Dígito.

    mfVerItau = False

    '1. Verifica agência.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                  " - Banco Itaú - mais que 4 posições)."
        Exit Function
    End If
    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    If pConta > 999999 Then
        pMensagem = "FGE2076 - Número de conta corrente inválido (" & pConta & _
                  " - Banco Itaú - mais que 6 posições)."
        Exit Function
    End If
    lConta = Format(pConta, "000000")

    '3. Calcula e verifica dígito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_ITAU(lAgencia, Left$(lConta, 5), lDigito, pMensagem) = False Then
        Exit Function
    End If
    Set clsA99V001 = Nothing
    If lDigito <> Mid$(lConta, 6, 1) Then
        pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco Itaú - " & pConta & ")."
        'Ficha 900591
'        If M46V999.gTipoAcesso = DIGITACAO Then
'            MsgBox pMensagem, vbCritical
'            M46V999.gPodeEmitir = False
'            mfVerItau = True
'            Exit Function
'        Else
            Exit Function
'        End If
    End If

    mfVerItau = True
End Function
Private Function mfVerSudameris(ByVal pAgencia As Long, ByVal pConta As Double, _
                                ByRef pMensagem As String) As Boolean
'Função: verifica informações de conta Sudameris.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetro de saída......pMensagem...Mensagem.

    Dim lAgencia As String    'Código da agência.
    Dim lConta As String  'Número da conta.
    Dim lDigito As String   'Dígito.

    mfVerSudameris = False

    '1. Verifica agência.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                  " - Banco Sudameris - mais que 4 posições)."
        Exit Function
    End If
    If pAgencia >= 1000 Then
        lAgencia = Format$(pAgencia, "0000")
    Else
        lAgencia = Format$(pAgencia, "000")
    End If

    '2. Verifica conta.
    If pConta > 9999999999# Then
        pMensagem = "FGE2076 - Número de conta corrente inválido (" & pConta & _
                  " - Banco Sudameris - mais que 10 posições)."
        Exit Function
    End If
    If pAgencia >= 1000 Then
        lConta = Format$(pConta, "00000000")
    Else
        lConta = Format$(pConta, "0000000000")
    End If

    '3. Calcula e verifica dígito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If pAgencia >= 1000 Then
        If clsA99V001.DIG_SUDAMERIS_NOVO(lAgencia, Left$(lConta, 7), lDigito, pMensagem) = False Then
            Exit Function
        End If
    Else
        If clsA99V001.DIG_SUDAMERIS(lAgencia, Left$(lConta, 9), lDigito, pMensagem) = False Then
            Exit Function
        End If
    End If
    Set clsA99V001 = Nothing
    If pAgencia >= 1000 Then
        If lDigito <> Mid$(lConta, 8, 1) Then
            pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco Sudameris - " & pConta & ")."
            'Ficha 900591
            '1001180 - Validação de DV na Triagem
            '            If M46V999.gTipoAcesso = DIGITACAO Then
            '                MsgBox pMensagem, vbCritical
            '                M46V999.gPodeEmitir = False
            '                mfVerSudameris = True
            Exit Function
            '            Else
            '                Exit Function
            '            End If
        End If
    Else
        If lDigito <> Mid$(lConta, 10, 1) Then
            pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco Sudameris - " & pConta & ")."
            'Ficha 900591
            '1001180 - Validação de DV na Triagem
            '            If M46V999.gTipoAcesso = DIGITACAO Then
            '                MsgBox pMensagem, vbCritical
            '                M46V999.gPodeEmitir = False
            '                mfVerSudameris = True
            Exit Function
            '            Else
            '                Exit Function
            '            End If
        End If
    End If

    mfVerSudameris = True
End Function
Private Function mfVerReal(ByVal pAgencia As Long, ByVal pConta As Double, _
                           ByRef pMensagem As String) As Boolean
'Função: verifica informações de conta Real.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetro de saída......pMensagem...Mensagem.

    Dim lAgencia As String    'Código da agência.
    Dim lConta As String  'Número da conta.
    Dim lDigito As String   'Dígito.

    mfVerReal = False

    '1. Verifica agência.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                  " - Banco Real - mais que 4 posições)."
        Exit Function
    End If
    lAgencia = Format$(pAgencia, "0000")


    '2. Verifica conta.
    If pConta > 9999999999# Then
        pMensagem = "FGE2076 - Número de conta corrente inválido (" & pConta & _
                  " - Banco Real - mais que 10 posições)."
        Exit Function
    End If
    lConta = Format$(pConta, "00000000")


    '3. Calcula e verifica dígito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_REAL(lAgencia, Left$(lConta, 7), lDigito, pMensagem) = False Then
        Exit Function
    End If

    Set clsA99V001 = Nothing
    If lDigito <> Mid$(lConta, 8, 1) Then
        pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco Real - " & pConta & ")."
        'Ficha 900591
        '1001180 - Validação de DV na Triagem
        '        If M46V999.gTipoAcesso = DIGITACAO Then
        '            MsgBox pMensagem, vbCritical
        '            M46V999.gPodeEmitir = False
        '            mfVerReal = True
        Exit Function
        '        Else
        '            Exit Function
        '        End If
    End If

    mfVerReal = True
End Function

Private Function mfVerUnibanco(ByVal pAgencia As Long, ByVal pConta As Double, _
                               ByRef pMensagem As String) As Boolean
'Função: verifica informações de conta Unibanco.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetro de saída......pMensagem...Mensagem.

    Dim lAgencia As String    'Código da agência.
    Dim lConta As String  'Número da conta.
    Dim lDigito As String   'Dígito.

    mfVerUnibanco = False

    '1. Verifica agência.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                  " - Banco Unibanco - mais que 4 posições)."
        Exit Function
    End If
    lAgencia = Format(pAgencia, "0000")
    '2. Verifica conta.
    If pConta > 9999999 Then
        pMensagem = "FGE2076 - Número de conta corrente inválido (" & pConta & _
                  " - Banco Unibanco - mais que 7 posições)."
        Exit Function
    End If
    lConta = Format(pConta, "0000000")

    '3. Calcula e verifica dígito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_UNIBANCO(lAgencia, Left$(lConta, 6), lDigito, pMensagem) = False Then
        Exit Function
    End If
    Set clsA99V001 = Nothing
    If lDigito <> Mid$(lConta, 7, 1) Then
        pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco Unibanco - " & pConta & ")."
        'Ficha 900591
        '1001180 - Validação de DV na Triagem
        '        If M46V999.gTipoAcesso = DIGITACAO Then
        '            MsgBox pMensagem, vbCritical
        '            M46V999.gPodeEmitir = False
        '            mfVerUnibanco = True
        Exit Function
        '        Else
        '            Exit Function
        '        End If
    End If

    mfVerUnibanco = True
End Function

Private Function mfVerSantander(ByVal pAgencia As Long, ByVal pConta As Double, _
                                ByRef pMensagem As String) As Boolean
'Função: verifica informações de conta Sudameris.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetro de saída......pMensagem...Mensagem.

    Dim lAgencia As String    'Código da agência.
    Dim lConta As String  'Número da conta.
    Dim lDigito As String   'Dígito.

    mfVerSantander = False

    '1. Verifica agência.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                  " - Banco Santander - mais que 4 posições)."
        Exit Function
    End If

    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    If pConta > 999999999 Then
        pMensagem = "FGE2076 - Número de conta corrente inválido (" & pConta & _
                  " - Banco Santander - mais que 8 posições)."
        Exit Function
    End If

    lConta = Format$(pConta, "000000000")


    '3. Calcula e verifica dígito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_SANTANDER(lAgencia, Left$(lConta, 8), lDigito, pMensagem) = False Then
        Exit Function
    End If

    Set clsA99V001 = Nothing


    If lDigito <> Mid$(lConta, 9, 1) Then
        pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco Santander - " & pConta & ")."
        'Ficha 900591
'        If M46V999.gTipoAcesso = DIGITACAO Then
'            MsgBox pMensagem, vbCritical
'            M46V999.gPodeEmitir = False
'            mfVerSantander = True
'            Exit Function
'        Else
            Exit Function
'        End If
    End If

    mfVerSantander = True
End Function
Private Function mfVerBanese(ByVal pAgencia As Long, ByVal pConta As Double, _
                             ByRef pMensagem As String) As Boolean
'Função: verifica informações de conta Sudameris.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetro de saída......pMensagem...Mensagem.

    Dim lAgencia As String    'Código da agência.
    Dim lConta As String  'Número da conta.
    Dim lDigito As String   'Dígito.

    mfVerBanese = False

    '1. Verifica agência.
    If pAgencia > 999 Then
        pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                  " - Banco Banese - mais que 3 posições)."
        Exit Function
    End If

    lAgencia = Format$(pAgencia, "000")

    '2. Verifica conta.
    If pConta > 999999999 Then
        pMensagem = "FGE2076 - Número de conta corrente inválido (" & pConta & _
                  " - Banco Banese - mais que 9 posições)."
        Exit Function
    End If

    lConta = Format$(pConta, "000000000")


    '3. Calcula e verifica dígito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_BANESE(lAgencia, Left$(lConta, 8), lDigito, pMensagem) = False Then
        Exit Function
    End If

    Set clsA99V001 = Nothing


    If lDigito <> Mid$(lConta, 9, 1) Then
        pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco Banese - " & pConta & ")."
        'Ficha 900591
'        If M46V999.gTipoAcesso = DIGITACAO Then
'            MsgBox pMensagem, vbCritical
'            M46V999.gPodeEmitir = False
'            mfVerBanese = True
'            Exit Function
'        Else
            Exit Function
'        End If
    End If

    mfVerBanese = True
End Function


Public Sub VER_COMUM_NIVEL_3(ByRef objstcPedido As stcA46V702B)   'Procedure: consiste dados comuns nível 3.

    Dim lObjstcVenc As stcA46V713B  'Objeto da classe de estrutura stcA46V713A (TAB_PED_VENC)



    If objstcPedido.Cod_Reg <> 120 And objstcPedido.Cod_Reg <> 106 And objstcPedido.Cod_Reg <> 20 And _
       objstcPedido.Cod_Reg <> 107 And objstcPedido.Cod_Reg <> 129 And Pro_QtdParcelas = 0 Then
        Call M46V999.gpMenErro("FGE0075 - Quantidade de parcelas inválida. Não informada.")

    End If

    If Pro_QtdParcelas = 0 Then GoTo SAIDA
    If Pro_PagAntecipado = 1 And Pro_QtdParcelas = 1 Then GoTo SAIDA


'    '2018-08-01: Juliano - Removido para nao validar cod_segurado
'    If Pro_ForPagamento = 1 Then
'        Select Case Left(val(Pro_CodSegurado), 6)
'        Case 777701, 777702, 777704, 777705, 777706, 777707, 777709, 777710, 777711, 777712, 777713, 777714, 777716, 777733, 777787, 777788
'            If objstcPedido.ColecaoPedVenc.Count > 0 Then
'                For Each lObjstcVenc In objstcPedido.ColecaoPedVenc
'                    If Pro_PagAntecipado <> 9 Then
'                        If Right(lObjstcVenc.DAT_VENC, 2) <> 26 And lObjstcVenc.Num_Parc <> 1 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inválida (para este código de segurado " & _
'                                                   "a data da parcela deve ser agendada para o dia 26).")
'                        End If
'                        If Day(Date) >= 10 And val(Mid(lObjstcVenc.DAT_VENC, 5, 2)) = val(Month(Date)) And lObjstcVenc.Num_Parc = 2 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inválida (para este código de segurado " & _
'                                                   "a data da parcela 2 deve ser agendada para o mês segunte.")
'                        End If
'                    Else
'                        If Right(lObjstcVenc.DAT_VENC, 2) <> 26 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inválida (para este código de segurado " & _
'                                                   "a data da parcela deve ser agendada para o dia 26).")
'                        End If
'                        If Day(Date) >= 10 And val(Mid(lObjstcVenc.DAT_VENC, 5, 2)) = val(Month(Date)) And lObjstcVenc.Num_Parc = 1 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inválida (para este código de segurado " & _
'                                                   "a data da parcela 1 deve ser agendada para o mês segunte.")
'                        End If
'                    End If
'                Next lObjstcVenc
'            Else
'                If Right(Pro_CarDatPriParcela, 2) <> 26 Then
'                    Call M46V999.gpMenErro("FGE0099 - Data da primeira parcela do carnê inválida (para este código de segurado " & _
'                                           "a data da parcela deve ser agendada para o dia 26).")
'                Else
'                    If Day(Date) >= 10 And val(Mid(Pro_CarDatPriParcela, 5, 2)) = val(Month(Date)) Then
'                        Call M46V999.gpMenErro("FGE0099 - Data da parcela do carnê inválida (para este código de segurado " & _
'                                               "a data da parcela deve ser agendada para o mês segunte.")
'                    End If
'                End If
'            End If
'        End Select
'    Else
'        Select Case Left(val(Pro_CodSegurado), 6)
'        Case 777701, 777702, 777704, 777705, 777706, 777707, 777709, 777710, 777711, 777712, 777713, 777714, 777716, 777733, 777787, 777788
'            If objstcPedido.ColecaoPedVenc.Count > 0 Then
'                For Each lObjstcVenc In objstcPedido.ColecaoPedVenc
'                    If Pro_PagAntecipado <> 9 Then
'                        If Right(lObjstcVenc.DAT_VENC, 2) <> 26 And lObjstcVenc.Num_Parc <> 1 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inválida (para este código de segurado " & _
'                                                   "a data da parcela deve ser agendada para o dia 26).")
'                        End If
'                        If Day(Date) >= 10 And val(Mid(lObjstcVenc.DAT_VENC, 5, 2)) = val(Month(Date)) And lObjstcVenc.Num_Parc = 2 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inválida (para este código de segurado " & _
'                                                   "a data da parcela 2 deve ser agendada para o mês segunte.")
'                        End If
'                    Else
'                        If Right(lObjstcVenc.DAT_VENC, 2) <> 26 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inválida (para este código de segurado " & _
'                                                   "a data da parcela deve ser agendada para o dia 26).")
'                        End If
'                        If Day(Date) >= 10 And val(Mid(lObjstcVenc.DAT_VENC, 5, 2)) = val(Month(Date)) And lObjstcVenc.Num_Parc = 1 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inválida (para este código de segurado " & _
'                                                   "a data da parcela 1 deve ser agendada para o mês segunte.")
'                        End If
'                    End If
'                Next lObjstcVenc
'            Else
'                Call M46V999.gpMenErro("FGE0099 - Data da parcela inválida (para este código de segurado " & _
'                                       "a data da parcela deve ser agendada para o dia 26).")
'            End If
'        End Select
'    End If

SAIDA:
    Exit Sub
    GoSub FechaObj
FechaObj:
    Set lObjstcVenc = Nothing
    Return
End Sub

'Projeto 11080 - Debito automatico - Banco do Brasil
Private Function mfVerBancoBrasil(ByVal pAgencia As Long, ByVal pConta As Double, _
                                  ByRef pMensagem As String, Optional blnVerificarAgencia As Boolean = False) As Boolean
'Função: verifica informações de conta Sudameris.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetro de saída......pMensagem...Mensagem.

    Dim lAgencia As String    'Código da agência.
    Dim lConta As String  'Número da conta.
    Dim lDigito As String   'Dígito.

    mfVerBancoBrasil = False
    
    If blnVerificarAgencia Then
        '1. Verifica agência.
        If pAgencia > 99999 Then
            pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                      " - Banco do Brasil - mais que 5 posições)."
            Exit Function
        End If
    
        lAgencia = Format$(pAgencia, "000000000")
    
        '3. Calcula e verifica dígito.
        Set clsA99V001 = New clsA99V001
        Set clsA99V001.M46V999 = M46V999
    
        lDigito = ""
        If clsA99V001.DIG_BANCOBRASIL(lAgencia, lDigito, pMensagem) = False Then
            Exit Function
        End If
        Set clsA99V001 = Nothing
    
        If lDigito = "X" Then lDigito = "0"
        If lDigito <> Mid$(lAgencia, 9, 1) Then
            pMensagem = "FGE2077 - Dígito da agência inválido (Banco do Brasil- " & pAgencia & ")."
            'Ficha 900591
'            If M46V999.gTipoAcesso = DIGITACAO Then
'                MsgBox pMensagem, vbCritical
'                M46V999.gPodeEmitir = False
'                mfVerBancoBrasil = True
'                Exit Function
'            Else
                Exit Function
'            End If
        End If
    End If
    
    If pConta > 9999999999# Then
        pMensagem = "FGE2076 - Número de conta corrente inválido (" & pConta & _
                  " - Banco do Brasil - mais que 10 posições)."
        Exit Function
    End If

    lConta = Format$(pConta, "000000000")


    '3. Calcula e verifica dígito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_BANCOBRASIL(Left$(lConta, 8), lDigito, pMensagem) = False Then
        Exit Function
    End If
    Set clsA99V001 = Nothing

    If lDigito = "X" Then lDigito = "0"
    If lDigito <> Mid$(lConta, 9, 1) Then
        pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco do Brasil- " & pConta & ")."
'        'Ficha 900591
'        If M46V999.gTipoAcesso = DIGITACAO Then
'            MsgBox pMensagem, vbCritical
'            M46V999.gPodeEmitir = False
'            mfVerBancoBrasil = True
'            Exit Function
'        Else
            Exit Function
'        End If
    End If

    mfVerBancoBrasil = True
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

'CyberTools - Ari - Task 3795:Implementar função da validação do CNH, Renavam, Safra e Bancoob - Setembro\2014
Private Function mfVerBANCOOB(ByVal pAgencia As Long, ByVal pConta As Double, _
                                ByRef pMensagem As String) As Boolean
    'Função: verifica informações de conta Bancoob.

    'Parâmetros de entrada...pAgencia....Código da agência.
    '                        pConta......Número da conta.
    'Parâmetro de saída......pMensagem...Mensagem.

    Dim lAgencia    As String       'Código da agência.
    Dim lConta      As String       'Número da conta.
    Dim lDigito     As String       'Dígito.

    mfVerBANCOOB = False

    '1. Verifica agência.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                    " - Banco BANCOOB - mais que 4 posições)."
        Exit Function
    End If
    
    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    'If pConta > 9999999999# Then
    If Len(pConta) > 10 Then
        pMensagem = "FGE2076 - Número de conta corrente inválido (" & pConta & _
                    " - Banco BANCOOB - mais que 10 posições)."
        Exit Function
    End If
    
    lConta = Format$(pConta, "0000000000")

    '3. Calcula e verifica dígito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999
    
    If clsA99V001.DIG_BANCOOB(lAgencia, Left$(lConta, 9), lDigito, pMensagem) = False Then
        Exit Function
    End If
    
    Set clsA99V001 = Nothing
    
    If lDigito <> Mid$(lConta, 10, 1) Then
        pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco BANCOOB - " & pConta & ")."
    'Ficha 900591
'        If M46V999.gTipoAcesso = DIGITACAO Then
'            MsgBox pMensagem, vbCritical
'            M46V999.gPodeEmitir = False
'            mfVerBANCOOB = True
'            Exit Function
'        Else
            Exit Function
'        End If
    End If
    
    mfVerBANCOOB = True
End Function

Private Function mfVerSAFRA(ByVal pAgencia As Long, ByVal pConta As Double, _
                                ByRef pMensagem As String) As Boolean
    'Função: verifica informações de conta SAFRA.

    'Parâmetros de entrada...pAgencia....Código da agência.
    '                        pConta......Número da conta.
    'Parâmetro de saída......pMensagem...Mensagem.

    Dim lAgencia    As String       'Código da agência.
    Dim lConta      As String       'Número da conta.
    Dim lDigito     As String       'Dígito.

    mfVerSAFRA = False

    '1. Verifica agência.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                    " - Banco SAFRA - mais que 4 posições)."
        Exit Function
    End If
    
    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    If pConta > 9999999 Then
        pMensagem = "FGE2076 - Número de conta corrente inválido (" & pConta & _
                    " - Banco SAFRA - mais que 7 posições)."
        Exit Function
    End If
    
    lConta = Format$(pConta, "0000000")

    '3. Calcula e verifica dígito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999
    
    If clsA99V001.DIG_SAFRA(lAgencia, Left$(lConta, 6), lDigito, pMensagem) = False Then
        Exit Function
    End If
    
    Set clsA99V001 = Nothing
    
    If lDigito <> Right(lConta, 1) Then
        pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco SAFRA - " & pConta & ")."
    'Ficha 900591
'        If M46V999.gTipoAcesso = DIGITACAO Then
'            MsgBox pMensagem, vbCritical
'            M46V999.gPodeEmitir = False
'            mfVerSAFRA = True
'            Exit Function
'        Else
            Exit Function
'        End If
    End If
    
    mfVerSAFRA = True
End Function
'Novo Convênio banco BRB
Private Function mfVerBANCOBRB(ByVal pAgencia As Long, ByVal pConta As Double, _
                                ByRef pMensagem As String) As Boolean
    'Função: verifica informações de conta BANCO BRB.

    'Parâmetros de entrada...pAgencia....Código da agência.
    '                        pConta......Número da conta.
    'Parâmetro de saída......pMensagem...Mensagem.

    Dim lAgencia    As String       'Código da agência.
    Dim lConta      As String       'Número da conta.
    Dim lDigito     As String       'Dígito.

    mfVerBANCOBRB = False

    '1. Verifica agência.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - Código de agência inválido (" & pAgencia & _
                    " - Banco BRB - mais que 4 posições)."
        Exit Function
    End If
    
    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    If Len(pConta) > 10 Then
        pMensagem = "FGE2076 - Número de conta corrente inválido (" & pConta & _
                    " - Banco BRB - mais que 10 posições)."
        Exit Function
    End If
    
    lConta = Format$(pConta, "0000000")

    '3. Calcula e verifica dígito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999
    
    If clsA99V001.DIG_BANCOBRB(lAgencia, Left$(lConta, 9), lDigito, pMensagem) = False Then
        Exit Function
    End If
    
    Set clsA99V001 = Nothing
    
    If lDigito <> Right(lConta, 1) Then
        pMensagem = "FGE2077 - Dígito de conta corrente inválido (Banco BRB - " & pConta & ")."
    'Ficha 900591
'        If M46V999.gTipoAcesso = DIGITACAO Then
'            MsgBox pMensagem, vbCritical
'            M46V999.gPodeEmitir = False
'            mfVerSAFRA = True
'            Exit Function
'        Else
            Exit Function
'        End If
    End If
    
    mfVerBANCOBRB = True
End Function


