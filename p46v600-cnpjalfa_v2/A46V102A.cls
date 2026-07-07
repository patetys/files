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
Private clsA99V001 As clsA99V001    'C�lculo de d�gitos.
Public Pro_CarBanco As Integer    'Pagamento com carn�: c�digo do banco.
Public Pro_CarDatPriParcela As Double   'Pagamento com carn�: data de vencimento da primeira
'parcela do carn� - formato aaaammdd.
Public Pro_Cob_Bairro As String    'Proposta: endere�o de cobran�a - bairro.
Public Pro_Cob_CEP As String    'Proposta: endere�o de cobran�a - CEP.
Public Pro_Cob_Cidade As String    'Proposta: endere�o de cobran�a - cidade.
Public Pro_Cob_Complemento As String  'Proposta: endere�o de cobran�a - complemento.
Public Pro_Cob_NomLogradouro As String    'Proposta: endere�o de cobran�a - nome do logradouro.
Public Pro_Cob_TipLogradouro As String    'Proposta: endere�o de cobran�a - tipo do logradouro.
Public Pro_Cob_UF As String    'Proposta: endere�o de cobran�a - unidade da federa��o.
Public Pro_CodCorretor As Long    'Proposta: c�digo do corretor.
Public Pro_Cosseguros As Variant    'Proposta: cosseguros.
Public Pro_DebAgencia As Long    'Proposta: d�bito em conta - ag�ncia.
Public Pro_DebBanco As Integer    'Proposta: d�bito em conta - banco.
Public Pro_DebConta As Double    'Proposta: d�bito em conta - conta.
Public Pro_End_Bairro As String    'Proposta: endere�o do segurado - bairro.
Public Pro_End_CEP As String    'Proposta: endere�o do segurado - CEP.
Public Pro_End_Cidade As String    'Proposta: endere�o do segurado - cidade.
Public Pro_End_Complemento As String  'Proposta: endere�o do segurado - complemento.
Public Pro_End_NomLogradouro As String    'Proposta: endere�o do segurado - nome do logradouro.
Public Pro_End_TipLogradouro As String    'Proposta: endere�o do segurado - tipo do logradouro.
Public Pro_End_UF As String    'Proposta: endere�o do segurado - unidade da federa��o.
Public Pro_ForPagamento As Byte    'Proposta: forma de pagamento.
'1 = carn� e 2 = d�bito em conta.
Public Pro_IndAgencia As Long    'Proposta: indicador do seguro - ag�ncia.
Public Pro_IndBanco As Integer    'Proposta: indicador do seguro - banco.
Public Pro_IndConta As Double    'Proposta: indicador do seguro - conta.
Public Pro_IndCobranca As Byte    'Proposta: indicador do endere�o de cobran�a diferente.
Public Pro_IndCosseguros As Byte  'Proposta: indicador de cess�o de cosseguro.
Public Pro_NomSegurado As String    'Proposta: nome do segurado.
Public Pro_PagAntBloqueto As Double    'Proposta: pagamento antecipado - bloqueto.
Public Pro_PagAntecipado As Byte  'Proposta: indicador de pagamento antecipado.
Public Pro_UniProdutiva As Integer    'Proposta: unidade produtiva.
Public Pro_VigInicio As Double    'Proposta: vig�ncia - in�cio - formato aaaammdd.

Public Pro_TipPessoa As Byte    'Proposta: tipo de pessoa.
Public Pro_CNPJ_CPF As String    'Proposta: CNPJ/CPF.
Public Pro_CodSegurado As Long
Public Pro_QtdParcelas As Integer
Public Pro_Cod_Reg As Integer
Public Pro_CreAgencia As Long    'Proposta: d�bito em conta - ag�ncia.
Public Pro_CreBanco As Integer    'Proposta: d�bito em conta - banco.
Public Pro_CreConta As Double    'Proposta: d�bito em conta - conta.
Public Pro_Tip_Res_Cre As Integer

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V102A"
End Property

Public Sub VER_COMUM_NIVEL_2()
'Procedure: consiste dados comuns n�vel 2.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lCodSegurado As String    'C�digo do segurado.
    Dim lDigito As String    'D�gito.
    Dim lEncontrou As Boolean    'Indica se encontrou registro.
    Dim lEndereco As String    'Endere�o completo.
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
                Call M46V999.gpMenErro("FGE0036 - Por motivos t�cnicos n�o � poss�vel prosseguir com este calculo.")
            End If
        End If
    End If
    'FIM - 11.32 Verifica Black list
    
'3. Endere�o do segurado.
    If Not M46V999.gfPreenchido(Pro_End_TipLogradouro) Then
        Call M46V999.gpMenErro("FGE0037 - Endere�o do segurado inv�lido (tipo do logradouro n�o preenchido).")
    End If
    If Not M46V999.gfPreenchido(Pro_End_NomLogradouro) Then
        Call M46V999.gpMenErro("FGE0038 - Endere�o do segurado inv�lido (nome do logradouro n�o preenchido).")
    End If
    If Not M46V999.gfPreenchido(Pro_End_Complemento) Then
        Call M46V999.gpMenErro("FGE0108 - Complemento do endere�o do segurado inv�lido (Complemento do logradouro n�o preenchido).")
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
        Call M46V999.gpMenErro("FGE0039 - Endere�o do segurado inv�lido (endere�o completo com " & _
                               lTamanho & _
                             " posi��es - tamanho m�ximo permitido no sistema de 50 posi��es)." & _
                               vbLf & vbLf & UCase(lEndereco))
    End If
    If Not M46V999.gfPreenchido(Pro_End_Bairro) Then
        Call M46V999.gpMenErro("FGE0040 - Endere�o do segurado inv�lido (bairro n�o preenchido).")
    End If
    If Not M46V999.gfPreenchido(Pro_End_Cidade) Then
        Call M46V999.gpMenErro("FGE0041 - Endere�o do segurado inv�lido (cidade n�o " & _
                               "preenchida).")
    End If
    If Not M46V999.gfPreenchido(Pro_End_UF) Then
        Call M46V999.gpMenErro("FGE0042 - Endere�o do segurado inv�lido (unidade da " & _
                               "federa��o n�o preenchida).")
    End If
    If Not M46V999.gfPreenchido(Pro_End_CEP) Then
        Call M46V999.gpMenErro("FGE0043 - Endere�o do segurado inv�lido (CEP n�o " & _
                               "preenchido).")
    End If
    If val(Pro_End_CEP) = 0 Then
        Call M46V999.gpMenErro("FGE0044 - Endere�o do segurado inv�lido (CEP " & _
                               "zerado).")
    End If
    If Len(Trim(Pro_End_CEP)) <> 8 Then
        Call M46V999.gpMenErro("FGE0045 - Endere�o do segurado inv�lido " & _
                               "(CEP com tamanho diferente de 8 posi��es).")
    End If

    Pro_End_Bairro = Trim$(Pro_End_Bairro)
    Pro_End_Cidade = Trim$(Pro_End_Cidade)

    '4. Endere�o de cobran�a.
    If Pro_IndCobranca = 1 Then     'Indica endere�o de cobran�a diferente posicionado.
        If Not M46V999.gfPreenchido(Pro_Cob_TipLogradouro) And Mid(Pro_Cob_NomLogradouro, 1, 4) <> "FIN-" Then
            Call M46V999.gpMenErro("FGE0046 - Endere�o de cobran�a inv�lido (tipo do logradouro n�o preenchido).")
        End If
        If Not M46V999.gfPreenchido(Pro_Cob_NomLogradouro) Then
            Call M46V999.gpMenErro("FGE0047 - Endere�o de cobran�a inv�lido (nome do logradouro n�o " & _
                                   "preenchido).")
        End If
        If Not M46V999.gfPreenchido(Pro_End_Complemento) Then
            Call M46V999.gpMenErro("FGE0109 - Complemento do endere�o de cobran�a inv�lido (Complemento do logradouro n�o preenchido).")
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
            Call M46V999.gpMenErro("FGE0048 - Endere�o de cobran�a inv�lido (endere�o completo " & _
                                   "com " & lTamanho & " posi��es - tamanho m�ximo permitido no " & _
                                   "sistema de 50 posi��es)." & vbLf & vbLf & UCase(lEndereco))
        End If
        If Not M46V999.gfPreenchido(Pro_Cob_Bairro) Then
            Call M46V999.gpMenErro("FGE0049 - Endere�o de cobran�a inv�lido (bairro n�o " & _
                                   "preenchido).")
        End If
        If Not M46V999.gfPreenchido(Pro_Cob_Cidade) Then
            Call M46V999.gpMenErro("FGE0050 - Endere�o de cobran�a inv�lido (cidade n�o " & _
                                   "preenchida).")
        End If
        If Not M46V999.gfPreenchido(Pro_Cob_UF) Then
            Call M46V999.gpMenErro("FGE0051 - Endere�o de cobran�a inv�lido (unidade " & _
                                   "da federa��o n�o preenchida).")
        End If
        If Not M46V999.gfPreenchido(Pro_Cob_CEP) Then
            Call M46V999.gpMenErro("FGE0056 - Endere�o de cobran�a inv�lido (CEP " & _
                                   "n�o preenchido).")
        End If
        If val(Pro_Cob_CEP) = 0 Then
            Call M46V999.gpMenErro("FGE0057 - Endere�o de cobran�a inv�lido " & _
                                   "(CEP zerado).")
        End If
        If Len(Trim(Pro_Cob_CEP)) <> 8 Then
            Call M46V999.gpMenErro("FGE0058 - Endere�o de cobran�a " & _
                                   "inv�lido (CEP com tamanho diferente " & _
                                   "de 8 posi��es).")
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
    
        If Pro_ForPagamento > 1 And Pro_ForPagamento <> 4 Then '2 = d�bito em conta. 3 = d�bito com primeira em boleto
            If lBanco <> 0 And Pro_Cod_Reg <> 129 Then
                    If mfVerConta(1, lBanco, lAgencia, lConta, lMensagem) = False Then
                        Call M46V999.gpMenErro(lMensagem)
                    End If
            Else
                If lBanco = 0 Then  'DebConta
                    Call M46V999.gpMenErro("FGE0073 - � necess�rio informar os dados de Debito em conta.")
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
            Call M46V999.gpMenErro("FGE0073 - � necess�rio informar os dados conta de restitui��o.")
        End If
    End If

    '1001266 - Validar CPF e CNPJ na Cota��o e Proposta
    '1. e 2. CNPJ/CPF.
    '    Select Case Pro_TipPessoa
    '        Case 1      'Pessoa jur�dica - trata-se de CNPJ.
    '            If IsNull(Pro_CNPJ_CPF) Then
    '                Call M46V999.gpMenErro("FNP0040 - CNPJ do segurado n�o preenchido.")
    '            Else
    '                If Pro_CNPJ_CPF = 0 Then
    '                    Call M46V999.gpMenErro("FNP0040 - CNPJ do segurado n�o preenchido.")
    '                Else
    '                    lAuxiliar = Format$(Pro_CNPJ_CPF, "00000000000000")
    '                    lAuxiliar = Mid$(lAuxiliar, 1, 2) & "." & Mid$(lAuxiliar, 3, 3) & "." & _
                         '                                Mid$(lAuxiliar, 6, 3) & "/" & Mid$(lAuxiliar, 9, 4) & "-" & _
                         '                                Mid$(lAuxiliar, 13, 2)
    '                    Set clsA99V001 = New clsA99V001     'Calcula d�gitos.
    '                    Set clsA99V001.M46V999 = M46V999
    '                    If clsA99V001.DIG_CNPJ(lAuxiliar, lDigito, lMensagem) = False Then
    '                        Call M46V999.gpMenErro(lMensagem)
    '                    End If
    '                    Set clsA99V001 = Nothing
    '                End If
    '            End If
    '        Case 2      'Pessoa f�sica - trata-se de CPF.
    '            If IsNull(Pro_CNPJ_CPF) Then
    '                Call M46V999.gpMenErro("FNP0034 - CPF do segurado n�o preenchido.")
    '            Else
    '                If Pro_CNPJ_CPF = 0 Then
    '                    Call M46V999.gpMenErro("FNP0034 - CPF do segurado n�o preenchido.")
    '                Else
    '                    lAuxiliar = Format$(Pro_CNPJ_CPF, "00000000000")
    '                    lAuxiliar = Mid$(lAuxiliar, 1, 3) & "." & Mid$(lAuxiliar, 4, 3) & "." & _
                         '                                Mid$(lAuxiliar, 7, 3) & "-" & Mid$(lAuxiliar, 10, 2)
    '                    Set clsA99V001 = New clsA99V001     'Calcula d�gitos.
    '                    Set clsA99V001.M46V999 = M46V999
    '                    If clsA99V001.DIG_CPF(lAuxiliar, lDigito, lMensagem) = False Then
    '                        Call M46V999.gpMenErro(lMensagem)
    '                    End If
    '                    Set clsA99V001 = Nothing
    '                End If
    '            End If
    '    End Select

    If (Pro_CarBanco = 409 Or Pro_DebBanco = 409) And M46V999.gTipoAcesso = CALCULO Then
        Call M46V999.gpMenErro("FGE0144 - N�o ser� poss�vel efetuar pagamento com a op��o de banco Unibanco (Cod. 409). Favor selecionar outro banco.")
    End If

    '1201070 - Tarifa - 1050 e 1051 - Item 1 - Planos TGP
    If M46V999.gCod_Ramo = e_CodRamo_Residencial And _
       M46V999.M46V111.gCod_plano >= 2278 And _
       M46V999.M46V111.gCod_plano <= 2297 Then
        If Pro_CarBanco = e_Banco_ND Or Pro_DebAgencia = 0 Or Pro_DebConta = 0 Then
            Call M46V999.gpMenErro("Para Plano TGP, o pagamento s� pode ser D�bito em Conta Corrente.")
        End If
    End If

End Sub

Private Function mfVerBoston(ByVal pAgencia As Long, ByVal pConta As Double, ByRef pMensagem As String) _
        As Boolean
'Fun��o: verifica informa��es de conta Boston.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metro de sa�da......pMensagem...Mensagem.

    Dim lAgencia As String    'C�digo da ag�ncia.
    Dim lConta As String  'N�mero da conta.
    Dim lDigito As String   'D�gito.

    mfVerBoston = False

    '1. Verifica ag�ncia.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                  " - Banco de Boston - mais que 4 posi��es)."
        Exit Function
    End If
    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    If pConta > 99999999 Then
        pMensagem = "FGE2076 - N�mero de conta corrente inv�lido (" & pConta & _
                  " - Banco de Boston - mais que 8 posi��es)."
        Exit Function
    End If
    lConta = Format(pConta, "00000000")

    '3. Calcula e verifica d�gito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_BOSTON(Left$(lConta, 7), lDigito, pMensagem) = False Then
        Exit Function
    End If
    Set clsA99V001 = Nothing
    If lDigito <> Mid$(lConta, 8, 1) Then
        pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco de Boston - " & pConta & ")."
        'Ficha 900591
        '1001180 - Valida��o de DV na Triagem
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
'Fun��o: verifica informa��es de conta Bradesco.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metro de sa�da......pMensagem...Mensagem.

    Dim lAgencia As String    'C�digo da ag�ncia.
    Dim lConta As String  'N�mero da conta.
    Dim lDigito As String   'D�gito.

    mfVerBradesco = False
    
    If blnVerificarAgencia Then
        '1. Verifica ag�ncia.
        If pAgencia > 99999 Then
            pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                      " - Banco Bradescon - mais que 5 posi��es)."
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
            pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                      " - Banco Bradesco - d�gito inv�lido)."
            Exit Function
        End If
    End If

    '2. Verifica conta.
    If pConta > 99999999 Then
        pMensagem = "FGE2076 - N�mero de conta corrente inv�lido (" & pConta & _
                  " - Banco Bradesco - mais que 8 posi��es)."
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
        pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco Bradesco - " & pConta & ")."
        'Ficha 900591
        '1001180 - Valida��o de DV na Triagem
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
'Fun��o: verifica informa��es de conta corrente.

'Par�metros de entrada...pOpcao......1 = Conta para d�bito.
'                                    2 = Indicador de seguro.
'                                    3 = Conta de restitui��o
'                        pBanco......C�digo do banco.
'                        pAgencia....C�digo da ag�ncia.
'                        pConta......C�digo da conta.
'Par�metro de sa�da......pMensagem...Mensagem.

    mfVerConta = False

    '1. Verifica op��o.
    If Not M46V999.gfPreenchido(pOpcao) Then
        pMensagem = "FTE0006 - Op��o inv�lida (verifica conta corrente - op��o n�o preenchida)."
        Exit Function
    End If
    If Not IsNumeric(pOpcao) Then
        pMensagem = "FTE0006 - Op��o inv�lida (verifica conta corrente - op��o n�o num�rica)."
        Exit Function
    End If
    If pOpcao <> 1 And pOpcao <> 2 And pOpcao <> 3 Then
        pMensagem = "FTE0006 - Op��o inv�lida (verifica conta corrente - op��o diferente de 1, 2 e 3: " & _
                    pOpcao & ")."
        Exit Function
    End If

    '2. Verifica banco.
    If pBanco = 0 Then
        Select Case pOpcao
        Case 1  'Conta para d�bito.
            pMensagem = "FGE0095 - C�digo do banco inv�lido (conta para d�bito - banco n�o " & _
                        "preenchido)."
        Case 2  'Indicador do seguro.
            pMensagem = "FGE0095 - C�digo do banco inv�lido (conta do indicador de seguro - " & _
                        "banco n�o preenchido )."
        Case 3  'conta de restitui��o.
            pMensagem = "FGE0095 - C�digo do banco inv�lido (conta de restitui��o - " & _
                        "banco n�o preenchido )."
        End Select
        Exit Function
    End If
    
    Dim codigo As Integer, Nome As String, nomeAbreviado As String, isCarne As Boolean, isDebito As Boolean, isRestituicao As Boolean
    pMensagem = Empty
    
    If (M46V999.GetBancoByID(pBanco, codigo, Nome, nomeAbreviado, isCarne, isDebito, isRestituicao, pMensagem) = False Or codigo = 0) Then
        pMensagem = "FGE0095 - C�digo do banco inv�lido. " & pMensagem
        Exit Function
    End If
    
    If pOpcao = 1 Then  'Conta para d�bito.
        ''20170116: original
        'Select Case pBanco
        '    'Projeto 11080 - Debito automatico - Banco do Brasil
        'Case 1, 237, 341, 409, 479, 356, 399, 33, 47, 422, 756    '09029
        '    'Nada a fazer.
        'Case Else
        '    pMensagem = "FGE0095 - C�digo do banco inv�lido (conta para d�bito - " & pBanco & ")."
        '    Exit Function
        'End Select
        ''20170116: fim
        If codigo = 0 Or isDebito = False Then
            pMensagem = "FGE0095 - C�digo do banco inv�lido para conta d�bito (" & pBanco & ")."
            Exit Function
        End If
    End If

    If pOpcao = 3 Then  'Conta para restitui��o.
        ''20170116: original j� comentado
        '        Select Case pBanco
        '        Case 1, 104, 748, 399, 756, 3, 394, 320, 47, 289, 634, 21, 604, 477, 376, 263, 745, 641, 29, 422, 479, 456, 464, 341, 237, 653, 370, 611, 707, 487, 610, 33, 41
        '            'Nada a fazer.
        '        Case Else
        '            pMensagem = "FGE0095 - C�digo do banco inv�lido (conta de restitui��o - " & pBanco & ")."
        '            Exit Function
        '        End Select
        ''20170116: fim
        If codigo = 0 Or isRestituicao = False Then
            pMensagem = "FGE0095 - C�digo do banco inv�lido para conta restitui��o (" & pBanco & ")."
            Exit Function
        End If
    End If


    '3. Verifica ag�ncia.
    If IsNull(pAgencia) Then
        Select Case pOpcao
        Case 1      'Conta para d�bito.
            pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (conta para d�bito - n�o preenchido)."
        Case 2      'Indicador de seguro.
            pMensagem = "FNP0073 - C�digo de ag�ncia inv�lido (conta do indicador do seguro - n�o preenchido)."
        Case 3      'conta de restitui��o.
            pMensagem = "FNP0073 - C�digo de ag�ncia inv�lido (conta de restitui��o - n�o preenchido)."
        End Select
        Exit Function
    End If
    
    If pAgencia = 0 Then
        Select Case pOpcao
        Case 1      'Conta para d�bito.
            pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (conta para d�bito - n�o preenchido)."
        Case 2      'Indicador de seguro.
            pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (conta do indicador do seguro - n�o preenchido)."
        Case 3      'conta de restitui��o.
            pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (conta de restitui��o - n�o preenchido)."
        End Select
        Exit Function
    End If

    '3. Verifica conta.
    If IsNull(pConta) Then
        Select Case pOpcao
        Case 1      'Conta para d�bito.
            pMensagem = "FNP2002 - N�mero da conta corrente n�o preenchido (conta para d�bito)."
        Case 2      'Indicador de seguro.
            pMensagem = "FNP2002 - N�mero da conta corrente n�o preenchido (conta do indicador do seguro)."
        Case 3      'conta de restitui��o.
            pMensagem = "FNP2002 - N�mero da conta corrente n�o preenchido (conta de restitui��o)."
        End Select
        Exit Function
    End If
    If pConta = 0 Then
        Select Case pOpcao
        Case 1      'Conta para d�bito.
            pMensagem = "FNP2002 - N�mero da conta corrente n�o preenchido (conta para d�bito)."
        Case 2      'Indicador de seguro.
            pMensagem = "FNP2002 - N�mero da conta corrente n�o preenchido (conta do indicador do seguro)."
        Case 3      'conta de restitui��o.
            pMensagem = "FNP2002 - N�mero da conta corrente n�o preenchido (conta de restitui��o)."
        End Select
        Exit Function
    End If

    '4. Verifica��es espec�ficas por banco.
    Select Case pBanco
    Case 1  'Projeto 11080 - Debito automatico - Banco do Brasil
        If mfVerBancoBrasil(pAgencia, pConta, pMensagem, pOpcao <> 3) = False Then
            Exit Function
        End If
    Case 237    'Bradesco.
        If mfVerBradesco(pAgencia, pConta, pMensagem, pOpcao <> 3) = False Then
            Exit Function
        End If
    Case 341    'Ita�.
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
    'Nova Tarifa��o
    Case 422    'Safra
        If mfVerSAFRA(pAgencia, pConta, pMensagem) = False Then
            Exit Function
        End If
    Case 756    'BANCOOB
        If mfVerBANCOOB(pAgencia, pConta, pMensagem) = False Then
            Exit Function
        End If
    'Novo Conv�nio banco BRB
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
'Fun��o: verifica informa��es de conta HSBC

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metro de sa�da......pMensagem...Mensagem.

    Dim lAgencia As String    'C�digo da ag�ncia.
    Dim lConta As String  'N�mero da conta.
    Dim lDigito As String   'D�gito.

    mfVerHSBC = False

    '1. Verifica ag�ncia.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                  " - Banco HSBC- mais que 4 posi��es)."
        Exit Function
    End If

    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    If pConta > 99999999999# Then
        pMensagem = "FGE2076 - N�mero de conta corrente inv�lido (" & pConta & _
                  " - Banco HSBC - mais que 11 posi��es)."
        Exit Function
    End If

    lConta = Format(pConta, "00000000000")

    '3. Calcula e verifica d�gito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_HSBC(lAgencia, Left$(lConta, 10), lDigito, pMensagem) = False Then
        Exit Function
    End If

    Set clsA99V001 = Nothing

    If lDigito <> Mid$(lConta, 11, 1) Then
        pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco HSBC- " & pConta & ")."
        'Ficha 900591
        '1001180 - Valida��o de DV na Triagem
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
'Fun��o: verifica informa��es de conta Ita�.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metro de sa�da......pMensagem...Mensagem.

    Dim lAgencia As String    'C�digo da ag�ncia.
    Dim lConta As String  'N�mero da conta.
    Dim lDigito As String   'D�gito.

    mfVerItau = False

    '1. Verifica ag�ncia.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                  " - Banco Ita� - mais que 4 posi��es)."
        Exit Function
    End If
    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    If pConta > 999999 Then
        pMensagem = "FGE2076 - N�mero de conta corrente inv�lido (" & pConta & _
                  " - Banco Ita� - mais que 6 posi��es)."
        Exit Function
    End If
    lConta = Format(pConta, "000000")

    '3. Calcula e verifica d�gito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_ITAU(lAgencia, Left$(lConta, 5), lDigito, pMensagem) = False Then
        Exit Function
    End If
    Set clsA99V001 = Nothing
    If lDigito <> Mid$(lConta, 6, 1) Then
        pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco Ita� - " & pConta & ")."
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
'Fun��o: verifica informa��es de conta Sudameris.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metro de sa�da......pMensagem...Mensagem.

    Dim lAgencia As String    'C�digo da ag�ncia.
    Dim lConta As String  'N�mero da conta.
    Dim lDigito As String   'D�gito.

    mfVerSudameris = False

    '1. Verifica ag�ncia.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                  " - Banco Sudameris - mais que 4 posi��es)."
        Exit Function
    End If
    If pAgencia >= 1000 Then
        lAgencia = Format$(pAgencia, "0000")
    Else
        lAgencia = Format$(pAgencia, "000")
    End If

    '2. Verifica conta.
    If pConta > 9999999999# Then
        pMensagem = "FGE2076 - N�mero de conta corrente inv�lido (" & pConta & _
                  " - Banco Sudameris - mais que 10 posi��es)."
        Exit Function
    End If
    If pAgencia >= 1000 Then
        lConta = Format$(pConta, "00000000")
    Else
        lConta = Format$(pConta, "0000000000")
    End If

    '3. Calcula e verifica d�gito.
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
            pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco Sudameris - " & pConta & ")."
            'Ficha 900591
            '1001180 - Valida��o de DV na Triagem
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
            pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco Sudameris - " & pConta & ")."
            'Ficha 900591
            '1001180 - Valida��o de DV na Triagem
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
'Fun��o: verifica informa��es de conta Real.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metro de sa�da......pMensagem...Mensagem.

    Dim lAgencia As String    'C�digo da ag�ncia.
    Dim lConta As String  'N�mero da conta.
    Dim lDigito As String   'D�gito.

    mfVerReal = False

    '1. Verifica ag�ncia.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                  " - Banco Real - mais que 4 posi��es)."
        Exit Function
    End If
    lAgencia = Format$(pAgencia, "0000")


    '2. Verifica conta.
    If pConta > 9999999999# Then
        pMensagem = "FGE2076 - N�mero de conta corrente inv�lido (" & pConta & _
                  " - Banco Real - mais que 10 posi��es)."
        Exit Function
    End If
    lConta = Format$(pConta, "00000000")


    '3. Calcula e verifica d�gito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_REAL(lAgencia, Left$(lConta, 7), lDigito, pMensagem) = False Then
        Exit Function
    End If

    Set clsA99V001 = Nothing
    If lDigito <> Mid$(lConta, 8, 1) Then
        pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco Real - " & pConta & ")."
        'Ficha 900591
        '1001180 - Valida��o de DV na Triagem
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
'Fun��o: verifica informa��es de conta Unibanco.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metro de sa�da......pMensagem...Mensagem.

    Dim lAgencia As String    'C�digo da ag�ncia.
    Dim lConta As String  'N�mero da conta.
    Dim lDigito As String   'D�gito.

    mfVerUnibanco = False

    '1. Verifica ag�ncia.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                  " - Banco Unibanco - mais que 4 posi��es)."
        Exit Function
    End If
    lAgencia = Format(pAgencia, "0000")
    '2. Verifica conta.
    If pConta > 9999999 Then
        pMensagem = "FGE2076 - N�mero de conta corrente inv�lido (" & pConta & _
                  " - Banco Unibanco - mais que 7 posi��es)."
        Exit Function
    End If
    lConta = Format(pConta, "0000000")

    '3. Calcula e verifica d�gito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_UNIBANCO(lAgencia, Left$(lConta, 6), lDigito, pMensagem) = False Then
        Exit Function
    End If
    Set clsA99V001 = Nothing
    If lDigito <> Mid$(lConta, 7, 1) Then
        pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco Unibanco - " & pConta & ")."
        'Ficha 900591
        '1001180 - Valida��o de DV na Triagem
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
'Fun��o: verifica informa��es de conta Sudameris.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metro de sa�da......pMensagem...Mensagem.

    Dim lAgencia As String    'C�digo da ag�ncia.
    Dim lConta As String  'N�mero da conta.
    Dim lDigito As String   'D�gito.

    mfVerSantander = False

    '1. Verifica ag�ncia.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                  " - Banco Santander - mais que 4 posi��es)."
        Exit Function
    End If

    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    If pConta > 999999999 Then
        pMensagem = "FGE2076 - N�mero de conta corrente inv�lido (" & pConta & _
                  " - Banco Santander - mais que 8 posi��es)."
        Exit Function
    End If

    lConta = Format$(pConta, "000000000")


    '3. Calcula e verifica d�gito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_SANTANDER(lAgencia, Left$(lConta, 8), lDigito, pMensagem) = False Then
        Exit Function
    End If

    Set clsA99V001 = Nothing


    If lDigito <> Mid$(lConta, 9, 1) Then
        pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco Santander - " & pConta & ")."
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
'Fun��o: verifica informa��es de conta Sudameris.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metro de sa�da......pMensagem...Mensagem.

    Dim lAgencia As String    'C�digo da ag�ncia.
    Dim lConta As String  'N�mero da conta.
    Dim lDigito As String   'D�gito.

    mfVerBanese = False

    '1. Verifica ag�ncia.
    If pAgencia > 999 Then
        pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                  " - Banco Banese - mais que 3 posi��es)."
        Exit Function
    End If

    lAgencia = Format$(pAgencia, "000")

    '2. Verifica conta.
    If pConta > 999999999 Then
        pMensagem = "FGE2076 - N�mero de conta corrente inv�lido (" & pConta & _
                  " - Banco Banese - mais que 9 posi��es)."
        Exit Function
    End If

    lConta = Format$(pConta, "000000000")


    '3. Calcula e verifica d�gito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_BANESE(lAgencia, Left$(lConta, 8), lDigito, pMensagem) = False Then
        Exit Function
    End If

    Set clsA99V001 = Nothing


    If lDigito <> Mid$(lConta, 9, 1) Then
        pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco Banese - " & pConta & ")."
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


Public Sub VER_COMUM_NIVEL_3(ByRef objstcPedido As stcA46V702B)   'Procedure: consiste dados comuns n�vel 3.

    Dim lObjstcVenc As stcA46V713B  'Objeto da classe de estrutura stcA46V713A (TAB_PED_VENC)



    If objstcPedido.Cod_Reg <> 120 And objstcPedido.Cod_Reg <> 106 And objstcPedido.Cod_Reg <> 20 And _
       objstcPedido.Cod_Reg <> 107 And objstcPedido.Cod_Reg <> 129 And Pro_QtdParcelas = 0 Then
        Call M46V999.gpMenErro("FGE0075 - Quantidade de parcelas inv�lida. N�o informada.")

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
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inv�lida (para este c�digo de segurado " & _
'                                                   "a data da parcela deve ser agendada para o dia 26).")
'                        End If
'                        If Day(Date) >= 10 And val(Mid(lObjstcVenc.DAT_VENC, 5, 2)) = val(Month(Date)) And lObjstcVenc.Num_Parc = 2 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inv�lida (para este c�digo de segurado " & _
'                                                   "a data da parcela 2 deve ser agendada para o m�s segunte.")
'                        End If
'                    Else
'                        If Right(lObjstcVenc.DAT_VENC, 2) <> 26 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inv�lida (para este c�digo de segurado " & _
'                                                   "a data da parcela deve ser agendada para o dia 26).")
'                        End If
'                        If Day(Date) >= 10 And val(Mid(lObjstcVenc.DAT_VENC, 5, 2)) = val(Month(Date)) And lObjstcVenc.Num_Parc = 1 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inv�lida (para este c�digo de segurado " & _
'                                                   "a data da parcela 1 deve ser agendada para o m�s segunte.")
'                        End If
'                    End If
'                Next lObjstcVenc
'            Else
'                If Right(Pro_CarDatPriParcela, 2) <> 26 Then
'                    Call M46V999.gpMenErro("FGE0099 - Data da primeira parcela do carn� inv�lida (para este c�digo de segurado " & _
'                                           "a data da parcela deve ser agendada para o dia 26).")
'                Else
'                    If Day(Date) >= 10 And val(Mid(Pro_CarDatPriParcela, 5, 2)) = val(Month(Date)) Then
'                        Call M46V999.gpMenErro("FGE0099 - Data da parcela do carn� inv�lida (para este c�digo de segurado " & _
'                                               "a data da parcela deve ser agendada para o m�s segunte.")
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
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inv�lida (para este c�digo de segurado " & _
'                                                   "a data da parcela deve ser agendada para o dia 26).")
'                        End If
'                        If Day(Date) >= 10 And val(Mid(lObjstcVenc.DAT_VENC, 5, 2)) = val(Month(Date)) And lObjstcVenc.Num_Parc = 2 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inv�lida (para este c�digo de segurado " & _
'                                                   "a data da parcela 2 deve ser agendada para o m�s segunte.")
'                        End If
'                    Else
'                        If Right(lObjstcVenc.DAT_VENC, 2) <> 26 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inv�lida (para este c�digo de segurado " & _
'                                                   "a data da parcela deve ser agendada para o dia 26).")
'                        End If
'                        If Day(Date) >= 10 And val(Mid(lObjstcVenc.DAT_VENC, 5, 2)) = val(Month(Date)) And lObjstcVenc.Num_Parc = 1 Then
'                            Call M46V999.gpMenErro("FGE0099 - Data da parcela " & lObjstcVenc.Num_Parc & " inv�lida (para este c�digo de segurado " & _
'                                                   "a data da parcela 1 deve ser agendada para o m�s segunte.")
'                        End If
'                    End If
'                Next lObjstcVenc
'            Else
'                Call M46V999.gpMenErro("FGE0099 - Data da parcela inv�lida (para este c�digo de segurado " & _
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
'Fun��o: verifica informa��es de conta Sudameris.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metro de sa�da......pMensagem...Mensagem.

    Dim lAgencia As String    'C�digo da ag�ncia.
    Dim lConta As String  'N�mero da conta.
    Dim lDigito As String   'D�gito.

    mfVerBancoBrasil = False
    
    If blnVerificarAgencia Then
        '1. Verifica ag�ncia.
        If pAgencia > 99999 Then
            pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                      " - Banco do Brasil - mais que 5 posi��es)."
            Exit Function
        End If
    
        lAgencia = Format$(pAgencia, "000000000")
    
        '3. Calcula e verifica d�gito.
        Set clsA99V001 = New clsA99V001
        Set clsA99V001.M46V999 = M46V999
    
        lDigito = ""
        If clsA99V001.DIG_BANCOBRASIL(lAgencia, lDigito, pMensagem) = False Then
            Exit Function
        End If
        Set clsA99V001 = Nothing
    
        If lDigito = "X" Then lDigito = "0"
        If lDigito <> Mid$(lAgencia, 9, 1) Then
            pMensagem = "FGE2077 - D�gito da ag�ncia inv�lido (Banco do Brasil- " & pAgencia & ")."
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
        pMensagem = "FGE2076 - N�mero de conta corrente inv�lido (" & pConta & _
                  " - Banco do Brasil - mais que 10 posi��es)."
        Exit Function
    End If

    lConta = Format$(pConta, "000000000")


    '3. Calcula e verifica d�gito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    If clsA99V001.DIG_BANCOBRASIL(Left$(lConta, 8), lDigito, pMensagem) = False Then
        Exit Function
    End If
    Set clsA99V001 = Nothing

    If lDigito = "X" Then lDigito = "0"
    If lDigito <> Mid$(lConta, 9, 1) Then
        pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco do Brasil- " & pConta & ")."
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
        Err.Raise 91, Me.Class_Name, "Objeto M46V999 n�o atribu�do para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V999 j�
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

'CyberTools - Ari - Task 3795:Implementar fun��o da valida��o do CNH, Renavam, Safra e Bancoob - Setembro\2014
Private Function mfVerBANCOOB(ByVal pAgencia As Long, ByVal pConta As Double, _
                                ByRef pMensagem As String) As Boolean
    'Fun��o: verifica informa��es de conta Bancoob.

    'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
    '                        pConta......N�mero da conta.
    'Par�metro de sa�da......pMensagem...Mensagem.

    Dim lAgencia    As String       'C�digo da ag�ncia.
    Dim lConta      As String       'N�mero da conta.
    Dim lDigito     As String       'D�gito.

    mfVerBANCOOB = False

    '1. Verifica ag�ncia.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                    " - Banco BANCOOB - mais que 4 posi��es)."
        Exit Function
    End If
    
    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    'If pConta > 9999999999# Then
    If Len(pConta) > 10 Then
        pMensagem = "FGE2076 - N�mero de conta corrente inv�lido (" & pConta & _
                    " - Banco BANCOOB - mais que 10 posi��es)."
        Exit Function
    End If
    
    lConta = Format$(pConta, "0000000000")

    '3. Calcula e verifica d�gito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999
    
    If clsA99V001.DIG_BANCOOB(lAgencia, Left$(lConta, 9), lDigito, pMensagem) = False Then
        Exit Function
    End If
    
    Set clsA99V001 = Nothing
    
    If lDigito <> Mid$(lConta, 10, 1) Then
        pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco BANCOOB - " & pConta & ")."
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
    'Fun��o: verifica informa��es de conta SAFRA.

    'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
    '                        pConta......N�mero da conta.
    'Par�metro de sa�da......pMensagem...Mensagem.

    Dim lAgencia    As String       'C�digo da ag�ncia.
    Dim lConta      As String       'N�mero da conta.
    Dim lDigito     As String       'D�gito.

    mfVerSAFRA = False

    '1. Verifica ag�ncia.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                    " - Banco SAFRA - mais que 4 posi��es)."
        Exit Function
    End If
    
    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    If pConta > 9999999 Then
        pMensagem = "FGE2076 - N�mero de conta corrente inv�lido (" & pConta & _
                    " - Banco SAFRA - mais que 7 posi��es)."
        Exit Function
    End If
    
    lConta = Format$(pConta, "0000000")

    '3. Calcula e verifica d�gito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999
    
    If clsA99V001.DIG_SAFRA(lAgencia, Left$(lConta, 6), lDigito, pMensagem) = False Then
        Exit Function
    End If
    
    Set clsA99V001 = Nothing
    
    If lDigito <> Right(lConta, 1) Then
        pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco SAFRA - " & pConta & ")."
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
'Novo Conv�nio banco BRB
Private Function mfVerBANCOBRB(ByVal pAgencia As Long, ByVal pConta As Double, _
                                ByRef pMensagem As String) As Boolean
    'Fun��o: verifica informa��es de conta BANCO BRB.

    'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
    '                        pConta......N�mero da conta.
    'Par�metro de sa�da......pMensagem...Mensagem.

    Dim lAgencia    As String       'C�digo da ag�ncia.
    Dim lConta      As String       'N�mero da conta.
    Dim lDigito     As String       'D�gito.

    mfVerBANCOBRB = False

    '1. Verifica ag�ncia.
    If pAgencia > 9999 Then
        pMensagem = "FGE0073 - C�digo de ag�ncia inv�lido (" & pAgencia & _
                    " - Banco BRB - mais que 4 posi��es)."
        Exit Function
    End If
    
    lAgencia = Format$(pAgencia, "0000")

    '2. Verifica conta.
    If Len(pConta) > 10 Then
        pMensagem = "FGE2076 - N�mero de conta corrente inv�lido (" & pConta & _
                    " - Banco BRB - mais que 10 posi��es)."
        Exit Function
    End If
    
    lConta = Format$(pConta, "0000000")

    '3. Calcula e verifica d�gito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999
    
    If clsA99V001.DIG_BANCOBRB(lAgencia, Left$(lConta, 9), lDigito, pMensagem) = False Then
        Exit Function
    End If
    
    Set clsA99V001 = Nothing
    
    If lDigito <> Right(lConta, 1) Then
        pMensagem = "FGE2077 - D�gito de conta corrente inv�lido (Banco BRB - " & pConta & ")."
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


