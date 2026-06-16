VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA99V009A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public UltimoErro As String

Private Function mfBuscarDadosTabela(ByRef pobjstc As stcA99V009B) As Boolean
    Dim lSql As String
    Dim lrst As clsYasRecordSet


    On Error GoTo ControleErro

    lSql = " SELECT  * FROM " & pobjstc.NomeTabela & " WITH (NOLOCK)"
    lSql = lSql & " Where IOF = 7.38 "
    lSql = lSql & " and TIPO_CONTA = 'ANT'"
    lSql = lSql & " and COD_BANCO = " & pobjstc.numeroBanco

    If pobjstc.M46V999.gfObtRegistro(pobjstc.ConexaoBD, lSql, lrst, UltimoErro) = False Then
        GoTo ControleErro
    End If

    If lrst.EOF Then
        UltimoErro = "Registro com dados da conta do banco " & pobjstc.numeroBanco & _
                   " não foram localizados."
        Exit Function
    End If

    With pobjstc
        .nomeBanco = lrst("NOM_BANCO")
        .numeroBanco = lrst("COD_BANCO")
        .CodConv = lrst("COD_CONV")
        .carteira = lrst("CART")
        .AgenciaCedente = lrst("AGENCIA")
        .ContaCedente = lrst("CONTA_COR")
        .IdentifCIA = lrst("COD_CIA")
        .AgenciaContaMask = lrst("CONTA_COR_MASK")
    End With
    mfBuscarDadosTabela = True
    Set lrst = Nothing
    Exit Function

ControleErro:
    Set lrst = Nothing
    If (UCase(Err.Description) Like "*TIMEOUT EXPIRED*") Or (UCase(UltimoErro) Like "*TIMEOUT EXPIRED*") Then
        UltimoErro = "mfBuscarDadosTabela - Tempo de execução da consulta expirado."
    Else
        UltimoErro = "mfBuscarDadosTabela - Erro ao obter registro (Erro = " & Err & " - " & ERROR & _
                   " - comando: " & lSql & ")." & IIf(UltimoErro = "", "", " " & UltimoErro)
    End If
    mfBuscarDadosTabela = False
End Function

Private Function mfLinhaDigitavel(pobjstc As stcA99V009B) As Boolean

    On Error GoTo ControleErros

    With pobjstc
        .linhaDigitavel = Format(.numeroBanco, "000") & "9" & Mid(.CodigoBarra, 20, 5) & Mid(.CodigoBarra, 25, 10) & Mid(.CodigoBarra, 35, 10) & .DigitoVerificador & Mid(.CodigoBarra, 6, 4) & Mid(.CodigoBarra, 10, 10)

        Call mpCalcularDVLinhaDigitavel(pobjstc)

        .linhaDigitavel = Left(.linhaDigitavel, 5) & "." & Mid(.linhaDigitavel, 6, 4) & .DV1 & " " & Mid(.linhaDigitavel, 10, 5) & "." & Mid(.linhaDigitavel, 15, 5) & .DV2 & " " & Mid(.linhaDigitavel, 20, 5) & "." & Mid(.linhaDigitavel, 25, 5) & .DV3 & " " & Mid(.linhaDigitavel, 30, 1) & " " & Mid(.linhaDigitavel, 31)
    End With

    mfLinhaDigitavel = True
    Exit Function
ControleErros:
    mfLinhaDigitavel = False
    UltimoErro = Err.Description
End Function

Private Function mfCodigoBarra(pobjstc As stcA99V009B) As Boolean

    On Error GoTo ControleErros
    With pobjstc

        .valorDocumento = mfTrataValor(.valorDocumento)

        .CodigoBarra = (Format(.numeroBanco, "000") & "9" & Format((DateDiff("d", Format("07/10/1997", "dd/mm/yyyy"), .getVencimento)), "0000") & Format(Replace(FormatNumber(.valorDocumento, 2), ",", ""), "0000000000")) & mfMontarPadraoEspecifico(pobjstc)

        Call mfCalcularDV(pobjstc)

        .CodigoBarra = Left(.CodigoBarra, 4) & .DigitoVerificador & Mid(.CodigoBarra, 5)
    End With

    mfCodigoBarra = True
    Exit Function

ControleErros:
    mfCodigoBarra = False
    UltimoErro = Err.Description
End Function
Private Function mfTrataValor(Valor As String) As Currency
    If InStr(Valor, ",") > 0 Then
        mfTrataValor = val(Str(Valor))
    Else
        mfTrataValor = val(Valor)
    End If
End Function

Private Function mfMontarPadraoEspecifico(pobjstc As stcA99V009B) As String
    Dim padrao As String
    Dim lAuxiliar As String

    'monta um string de acordo com o padrao do banco (bradesco ou Itau)
    lAuxiliar = ""
    padrao = ""
    Select Case pobjstc.numeroBanco
    Case "237"
        ''*******************BRADESCO**************************
        '20-23   4   Código da Agência (sem dígito)
        lAuxiliar = Format(Mid(pobjstc.AgenciaCedente, 1, 4), "0000")
        '24-25   2   Código da Carteira
        lAuxiliar = lAuxiliar & Format(pobjstc.carteira, "00")
        '26-36   11  Nosso Número (sem dígito)
        lAuxiliar = lAuxiliar & Format(Left(pobjstc.NossoNumero, 11), "00000000000")
        '37-43   7   Conta do Cedente (sem dígito)
        lAuxiliar = lAuxiliar & Right(Format(pobjstc.ContaCedente, "00000000"), 7)
        '44-44   1   Zero Fixo
        lAuxiliar = lAuxiliar & "0"
        padrao = lAuxiliar

    Case "341"
        ''*******************ITAÚ******************************
        ''''''Format(pobjstc.Carteira, "000") & Format(Right(pobjstc.NumeroDocumento, 9), "000000000") & Format(Mid(pobjstc.AgenciaCedente, 1, 4), "0000") & Left(Format(pobjstc.ContaCedente, "000000"), 5) & Right(pobjstc.ContaCedente, 1) & "000"
        '20 a 22 03  Carteira
        lAuxiliar = Format(pobjstc.carteira, "000")
        '23 a 30 08  Nosso Número
        lAuxiliar = lAuxiliar & Format(Right(pobjstc.NossoNumero, 8), "00000000")
        '31 a 31 01  DAC [Agência /Conta/Carteira/Nosso Número] (Anexo 4)
        pobjstc.DigBanco = mfRetornarDivNumItau(pobjstc)
        lAuxiliar = lAuxiliar & pobjstc.DigBanco
        '32 a 35 04  N.º da Agência cedente
        lAuxiliar = lAuxiliar & Format(Mid(pobjstc.AgenciaCedente, 1, 4), "0000")
        '36 a 40 05  N.º da Conta Corrente
        lAuxiliar = lAuxiliar & Left(Format(pobjstc.ContaCedente, "000000"), 5)
        '41 a 41 01  DAC [Agência/Conta Corrente] (Anexo 3)
        lAuxiliar = lAuxiliar & Right(pobjstc.ContaCedente, 1)
        '42 a 44 03  Zeros
        lAuxiliar = lAuxiliar & "000"

    Case "347"
        ''*******************SUDAMERIS*************************
        '20 a 20 01  Ident express (fixo 9)
        lAuxiliar = "9"
        '21 a 22 02 Iof sudameris
        lAuxiliar = lAuxiliar & "47"
        '23 a 29 07  conta corrente do cedente
        lAuxiliar = lAuxiliar & Left(Format(pobjstc.ContaCedente, "00000000"), 7)
        '30 a 37 08  Nosso número
        lAuxiliar = lAuxiliar & Format(Right(pobjstc.NossoNumero, 8), "00000000")
        '38 a 38 01  Dígito do nosso número
        lAuxiliar = lAuxiliar & NossoDigYasAnt(Format(Right(pobjstc.NossoNumero, 7), "0000000"))
        '39 a 43 05 Identificação da CIA
        lAuxiliar = lAuxiliar & Format(pobjstc.IdentifCIA, "00000")
        '44 a 44 01 Código Express
        lAuxiliar = lAuxiliar & "7"

    Case "409"
        ''*******************UNIBANCO**************************
        '20 a 20 01  Trans Unibanco
        lAuxiliar = "5"
        '21 a 27 07  Num Conv Unibanco
        'lAuxiliar = lAuxiliar & "1688758"
        'lAuxiliar = lAuxiliar & "2654121"
        lAuxiliar = lAuxiliar & pobjstc.CodConv
        '28 a 29 02  Fixo unibanco
        lAuxiliar = lAuxiliar & "00"
        '30 a 43 14  Nosso número
        lAuxiliar = lAuxiliar & Format(pobjstc.NossoNumero, "00000000000000")
        '44 a 44 01 Digito do nosso número
        lAuxiliar = lAuxiliar & NossoDigYasAnt(Format(pobjstc.NossoNumero, "0000000"))
    Case "356"
        ''*******************REAL******************************
        '20 a 23 04 Agência
        lAuxiliar = Format(pobjstc.AgenciaCedente, "0000")
        '24 a 30 07 conta
        lAuxiliar = lAuxiliar & Format(pobjstc.ContaCedente, "0000000")
        '30 a 31 01 dígito
        pobjstc.DigBanco = mfRetornarDivNumReal(pobjstc)
        lAuxiliar = lAuxiliar & Format(pobjstc.DigBanco, "0")
        '32 a 33 02 Aliquota
        lAuxiliar = lAuxiliar & "18"
        '34 a 44 11 nosso numero
        lAuxiliar = lAuxiliar & Format(pobjstc.NossoNumero, "00000000000")
    Case 399
        ''*******************HSBC******************************
        '20 26 07 Código do Cedente
        lAuxiliar = Format(pobjstc.ContaCedente, "0000000")
        '27 39 13 Número Bancário ou Código do Documento (Nosso Número,sem os dígitos verificadores e o tipo identificador.
        lAuxiliar = lAuxiliar & Format(pobjstc.NossoNumero, "0000000000000")
        '40 43 04 Data de Vencimento no Formato Juliano.
        lAuxiliar = lAuxiliar & Format((DateDiff("d", "01/01/" & Left(pobjstc.vencimento, 4), Right(pobjstc.vencimento, 2) & "/" & Mid(pobjstc.vencimento, 5, 2) & "/" & Left(pobjstc.vencimento, 4)) + 1) & Mid(pobjstc.vencimento, 4, 1), "0000")
        '44 44 01 Código do Produto CNR, deve preencher o número 2.
        lAuxiliar = lAuxiliar & "2"
    Case "033", "33"
        ''*******************SANTANDER *************************yasudanet.dbo.Tab_Contas_Carne-P0045300
        '20-20 1 9 (01) Fixo “9”
        '21-27 7 9 (07) Código do cedente padrão Santander
        '28-32 5 9 (05) Fixo “00000”
        '33-40 8 9 (08) Nosso Número
        '41-41 1 9 (01) IOF – Seguradoras (Se 7% informar 7. Limitado a 9%) Demais clientes usar 0 (zero)
        '42-44 3 9 (03) 101-Cobrança Simples Rápida COM Registro 102- Cobrança simples – SEM Registro '201- Penhor Rápida com Registro
      
        '20 a 20 01  Ident express (fixo 9)
        lAuxiliar = "9"
        '21 a 27 07  conta corrente do cedente
        lAuxiliar = lAuxiliar & Right(Format(pobjstc.CodConv, "00000000"), 7)
        '28 a 32 05
        lAuxiliar = lAuxiliar & "00000"
        '33 a 40 08 Nosso Número
        lAuxiliar = lAuxiliar & Format(Right(pobjstc.NossoNumero, 8), "00000000")
        '41 a 41 01 IOF – Seguradoras
        lAuxiliar = lAuxiliar & "0"
        '42 a 44 03 Cobrança
        lAuxiliar = lAuxiliar & Format(pobjstc.carteira, "000")
    End Select
    padrao = lAuxiliar
    mfMontarPadraoEspecifico = padrao

End Function
Private Function mfPreenchido(ByVal pCampo)
'Função #029: verifica se campo preenchido.

    If IsNull(pCampo) Then
        mfPreenchido = False
    Else
        If Len(Trim(pCampo)) = 0 Then
            mfPreenchido = False
        Else
            mfPreenchido = True
        End If
    End If

End Function

Private Function mfConsistirDados(pobjstc As stcA99V009B) As Boolean

    With pobjstc
        If .vencimento = 0 Then
            UltimoErro = "Vencimento inválido (Não informado)."
            Exit Function
        End If
        If Not mfPreenchido(.NossoNumero) Then
            UltimoErro = "Número do documento inválido (não informado)."
            Exit Function
        End If
        If Not mfPreenchido(.numeroBanco) Then
            UltimoErro = "Número do banco inválido (não informado)."
            Exit Function
        End If
    End With
    mfConsistirDados = True
End Function
Private Function mfConsistirDadosPagina(pobjstc As stcA99V009B) As Boolean

    With pobjstc
        .localPagamento = .getLocalPagamento
        .Instrucao = .getInstrucao
        .usoBanco = .getUsoBanco
        .EspecieDocumento = .getEspecieDocumento
        .aceite = .getAceite
        .Quantidade = .getQuantidade
        .NomeCedente = .getNomeCedente
        .cnpjCedente = .getCNPJCedente
        .Apolice = .getApolice
        .endosso = .getEndosso
        .carteira = .getCarteira
        'Informações obrigatórias
        If .vencimento = 0 Then
            UltimoErro = "Vencimento inválido (Não informado)."
            Exit Function
        End If
        If Not mfPreenchido(.NossoNumero) Then
            UltimoErro = "Número do documento inválido (não informado)."
            Exit Function
        End If
        If Not mfPreenchido(.carteira) Then
            UltimoErro = "Carteira inválida (não informad)."
            Exit Function
        End If
        If Not mfPreenchido(.NomeSacado) Then
            UltimoErro = "Nome do sacado inválido (não informado)."
            Exit Function
        End If
        If Not mfPreenchido(.CPF_CNPJ_Sacado) Then
            UltimoErro = "CPF/CNPJ do sacado inválido (não informado)."
            Exit Function
        End If
        If Not mfPreenchido(.AgenciaContaMask) Then
            UltimoErro = "Agencia/Conta do cedente inválida (não informada)."
            Exit Function
        End If
        If Not mfPreenchido(.numeroBanco) Then
            UltimoErro = "Número do banco inválido (não informado)."
            Exit Function
        End If
        If Not mfPreenchido(.nomeBanco) Then
            UltimoErro = "Nome do banco inválido (não informado)."
            Exit Function
        End If
        If Not mfPreenchido(.Vigencia) Then
            UltimoErro = "Vigência inválida (não informada)."
            Exit Function
        End If
    End With
    mfConsistirDadosPagina = True
End Function

Public Function mfMontarBoleto(pobjstc As stcA99V009B) As Boolean

    If Not mfConsistirDados(pobjstc) Then GoTo ControleErros

    If pobjstc.CarregarDados = True Then
        If Not mfBuscarDadosTabela(pobjstc) Then GoTo ControleErros
    End If
    If Not mfCodigoBarra(pobjstc) Then GoTo ControleErros
    If Not mfMontarCodigoBarra(pobjstc) Then GoTo ControleErros
    If Not mfLinhaDigitavel(pobjstc) Then GoTo ControleErros

    mfMontarBoleto = True
    Exit Function
ControleErros:
    If mfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
    End If
End Function

Public Function mfRetornarPaginaMontada(pobjstc As stcA99V009B, pTemplate As String, pNovoArquivo As String) As Boolean
    Dim lArquivo As Scripting.FileSystemObject
    Dim lArqLido As Scripting.TextStream
    Dim lRegistro As String
    Dim lPosicao As Long
    Dim lPosInicial As Long
    Dim lPosFinal As Long
    Dim lPosInfo As Long
    Dim lInformacao As String
    Dim lUltPos As Long
    Dim lTamanho As Long
    Dim lAuxiliar As String
    Dim lNovoArq As Scripting.TextStream
    Dim ne As String
    Dim nl As String
    Dim we As String
    Dim wl As String
    Dim RecortarWN As String
    Dim lDig As String
    Dim lSomaAux As Double
    ne = "<IMG SRC=" & Chr(34) & pobjstc.LocalImagensBarra & "\NB.gif" & Chr(34) & "><IMG SRC=" & Chr(34) & pobjstc.LocalImagensBarra & "/NW.gif" & Chr(34) & ">"
    nl = "<IMG SRC=" & Chr(34) & pobjstc.LocalImagensBarra & "\NB.gif" & Chr(34) & "><IMG SRC=" & Chr(34) & pobjstc.LocalImagensBarra & "/WW.gif" & Chr(34) & ">"
    we = "<IMG SRC=" & Chr(34) & pobjstc.LocalImagensBarra & "\WB.gif" & Chr(34) & "><IMG SRC=" & Chr(34) & pobjstc.LocalImagensBarra & "/NW.gif" & Chr(34) & ">"
    wl = "<IMG SRC=" & Chr(34) & pobjstc.LocalImagensBarra & "\WB.gif" & Chr(34) & "><IMG SRC=" & Chr(34) & pobjstc.LocalImagensBarra & "/WW.gif" & Chr(34) & ">"

    If Not mfConsistirDadosPagina(pobjstc) Then GoTo ControleErro

    pobjstc.NossoNumeroMask = pobjstc.getNossoNumeroMask


    Set lArquivo = New Scripting.FileSystemObject
    If Not lArquivo.FileExists(pTemplate) Then
        UltimoErro = "Arquivo " & pTemplate & " não localizado."
        Exit Function
    End If


    Set lNovoArq = lArquivo.CreateTextFile(pNovoArquivo, True)
    FileCopy pTemplate, Replace(UCase(pTemplate), ".EXE", ".TXT")
    pTemplate = Replace(UCase(pTemplate), ".EXE", ".TXT")
    Set lArqLido = lArquivo.OpenTextFile(pTemplate)

    Do
        If lArqLido.AtEndOfStream Then Exit Do
        lRegistro = lArqLido.ReadLine
        lPosInicial = 1
        Do
            lPosInicial = InStr(lPosInicial, lRegistro, "@@")
            If lPosInicial = 0 Then Exit Do

            lPosFinal = InStr(lPosInicial + 2, lRegistro, "@@")

            lTamanho = lPosFinal + 2 - lPosInicial
            lInformacao = Mid(lRegistro, lPosInicial, lTamanho)
            lAuxiliar = ""
            Select Case UCase(lInformacao)
            Case "@@ACEITE@@"
                lAuxiliar = pobjstc.aceite
            Case "@@AGECODCEDENTE@@"
                lAuxiliar = pobjstc.AgenciaContaMask
            Case "@@APOLICE@@"
                lAuxiliar = pobjstc.Apolice
            Case "@@BANCO@@"
                lAuxiliar = pobjstc.numeroBanco
            Case "@@CARTEIRA@@"
                lAuxiliar = pobjstc.carteira
            Case "@@CEDENTE@@"
                lAuxiliar = pobjstc.NomeCedente
            Case "@@CEDENTECNPJ@@"
                lAuxiliar = pobjstc.cnpjCedente
            Case "@@CODIGOBARRA@@"
                'Tratar
                lAuxiliar = "<IMG SRC=" & Chr(34) & pobjstc.LocalImagensBarra & "/NB.gif" & Chr(34) & ">"
                lAuxiliar = lAuxiliar & "<IMG src=" & Chr(34) & pobjstc.LocalImagensBarra & "/NW.gif" & Chr(34) & ">"
                lAuxiliar = lAuxiliar & "<IMG src=" & Chr(34) & pobjstc.LocalImagensBarra & "/NB.gif" & Chr(34) & ">"
                lAuxiliar = lAuxiliar & "<IMG src=" & Chr(34) & pobjstc.LocalImagensBarra & "/NW.gif" & Chr(34) & ">"
                For lPosicao = 1 To Len(pobjstc.LinhaWN)
                    RecortarWN = Mid(pobjstc.LinhaWN, lPosicao, 2)
                    Select Case UCase(RecortarWN)
                    Case "NE"
                        lAuxiliar = lAuxiliar + ne
                    Case "NL"
                        lAuxiliar = lAuxiliar + nl
                    Case "WE"
                        lAuxiliar = lAuxiliar + we
                    Case "WL"
                        lAuxiliar = lAuxiliar + wl
                    End Select
                    lPosicao = lPosicao + 1
                Next
                lAuxiliar = lAuxiliar & "<IMG src=" & Chr(34) & pobjstc.LocalImagensBarra & "/WB.gif" & Chr(34) & ">"
                lAuxiliar = lAuxiliar & "<IMG src=" & Chr(34) & pobjstc.LocalImagensBarra & "/NW.gif" & Chr(34) & ">"
                lAuxiliar = lAuxiliar & "<IMG src=" & Chr(34) & pobjstc.LocalImagensBarra & "/NB.gif" & Chr(34) & ">"
                lAuxiliar = lAuxiliar & "<IMG src=" & Chr(34) & pobjstc.LocalImagensBarra & "/NW.gif" & Chr(34) & ">"
            Case "@@CORRETOR@@"
                lAuxiliar = pobjstc.Corretor
            Case "@@DATASOLICITACAO@@"
                lAuxiliar = pobjstc.getDataDocumento
            Case "@@DATAVENCIMENTO@@"
                lAuxiliar = pobjstc.getVencimento
            Case "@@DATDOCUMENTO@@"
                lAuxiliar = pobjstc.getDataDocumento
            Case "@@DATPROCESS@@"
                lAuxiliar = pobjstc.getDataProcessamento
            Case "@@EMISSAO@@"
                lAuxiliar = pobjstc.getDataDocumento
            Case "@@ENDOSSO@@"
                lAuxiliar = pobjstc.endosso
            Case "@@ESPECIEDOC@@"
                lAuxiliar = pobjstc.EspecieDocumento
            Case "@@IMAGEMBANCO@@"
                lAuxiliar = pobjstc.LocalImagensBanco
            Case "@@INSTRUCOES@@"
                lAuxiliar = pobjstc.Instrucao
            Case "@@LINHADIGITAVEL@@"
                lAuxiliar = pobjstc.linhaDigitavel
            Case "@@LOCALPAGAMENTO@@"
                lAuxiliar = pobjstc.localPagamento
            Case "@@MORAMULTA@@"
                lAuxiliar = ""
            Case "@@NOSSONUMERO@@"
                lAuxiliar = pobjstc.NossoNumeroMask
            Case "@@NUMDOCUMENTO@@"
                lAuxiliar = pobjstc.NumeroDoc
            Case "@@QTDEPARCELAS@@"
                lAuxiliar = pobjstc.Quantidade
            Case "@@PAGAMENTO@@"
                lAuxiliar = pobjstc.MensagemPgto
            Case "@@PARCELAPAGAR@@"
                lAuxiliar = pobjstc.Quantidade
            Case "@@QUANTIDADE@@"
                lAuxiliar = pobjstc.Quantidade
            Case "@@RAMO@@"
                lAuxiliar = pobjstc.ramo
            Case "@@SEGURADO@@", "@@SACADO@@"
                lAuxiliar = pobjstc.NomeSacado
            Case "@@SACADOCPFCNPJ@@"
                lAuxiliar = pobjstc.CPF_CNPJ_Sacado
            Case "@@USOBANCO@@"
                lAuxiliar = pobjstc.usoBanco
            Case "@@VALDOC@@", "@@VALPARCELA@@"
                lAuxiliar = Format(pobjstc.valorDocumento, "standard")
            Case "@@VIGENCIA@@"
                lAuxiliar = pobjstc.Vigencia
            Case Else
                lAuxiliar = ""
            End Select
            If mfPreenchido(lAuxiliar) = False Then
                lAuxiliar = "&nbsp;"
            End If
            lRegistro = Replace(lRegistro, lInformacao, lAuxiliar)
        Loop
        lNovoArq.Write lRegistro & vbCrLf

    Loop
    lArqLido.Close
    lNovoArq.Close
    Set lArquivo = Nothing
    Set lArqLido = Nothing
    Set lNovoArq = Nothing

    mfRetornarPaginaMontada = True
    Exit Function
ControleErro:
    Set lArquivo = Nothing
    Set lArqLido = Nothing
    Set lNovoArq = Nothing
    If mfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
    End If
End Function

Private Function mfCalcularDV(pobjstc As stcA99V009B)

    Dim lPosicao As Long
    Dim DV As Long    ' digito verificador
    Dim NumMultiplicar    '23456789
    Dim CalculoDV As Long
    Dim CaracterDV As Long

    NumMultiplicar = 2
    DV = 0

    For lPosicao = 0 To Len(pobjstc.CodigoBarra) - 1
        CaracterDV = CLng(Mid(pobjstc.CodigoBarra, Len(pobjstc.CodigoBarra) - lPosicao, 1))

        CalculoDV = CaracterDV * NumMultiplicar

        DV = DV + CalculoDV

        NumMultiplicar = NumMultiplicar + 1

        If NumMultiplicar = 10 Then
            NumMultiplicar = 2
        End If

    Next

    DV = DV Mod 11

    If ((DV = 0) Or (DV = 10) Or (DV = 1)) Then    ' Se o resto for 0,10 ou 1 o DigitoVerificador é 1
        pobjstc.DigitoVerificador = 1
    Else
        pobjstc.DigitoVerificador = 11 - DV
    End If

End Function

Private Function mpCalcularDVLinhaDigitavel(pobjstc As stcA99V009B)
    Dim lPosicao As Long
    Dim DV As Long    ' digito verificador
    Dim NumMultiplicar    '2 e 1
    Dim CalculoDV As Long
    Dim CaracterDV As Long
    Dim temp As String


    'calcular DV1
    pobjstc.DV1 = Modulo_10(Mid(pobjstc.linhaDigitavel, 1, 9))

    'Calcular DV2
    pobjstc.DV2 = Modulo_10(Mid(pobjstc.linhaDigitavel, 10, 10))

    'Calcular DV3
    pobjstc.DV3 = Modulo_10(Mid(pobjstc.linhaDigitavel, 20, 10))



End Function


Private Function mfMontarCodigoBarra(pobjstc As stcA99V009B) As Boolean
    Dim Sequencia01 As String
    Dim SequenciaNW As String
    Dim CodRepresentacaoA As String
    Dim CodRepresentacaoB As String
    Dim CodNW As String
    Dim CodAux As String
    Dim lLocal As Long
    Dim lPosicao As Long
    Dim a As String
    Dim b As String

    On Error GoTo ControleErros

    If (Len(pobjstc.CodigoBarra) Mod 2) <> 0 Then
        pobjstc.CodigoBarra = "0" & pobjstc.CodigoBarra
    End If
    Sequencia01 = ""
    For lLocal = 1 To Len(pobjstc.CodigoBarra) - 1
        ' pega sequencia de dois números
        a = Mid(pobjstc.CodigoBarra, lLocal, 1)
        b = Mid(pobjstc.CodigoBarra, lLocal + 1, 1)

        CodRepresentacaoA = mfCodigoRepresentacao(a)

        CodRepresentacaoB = mfCodigoRepresentacao(b)

        'gerar a sequencia de 0 e 1
        For lPosicao = 1 To 5
            Sequencia01 = Sequencia01 & Mid(CodRepresentacaoA, lPosicao, 1) & Mid(CodRepresentacaoB, lPosicao, 1)
        Next
        lLocal = lLocal + 1
    Next
    pobjstc.CodigoBarraBinario = "0000" & Sequencia01 & "1000"

    SequenciaNW = ""
    'gerar a sequencia de ne, nl, we, wl
    For lLocal = 1 To (Len(Sequencia01))

        If (Mid(Sequencia01, lLocal, 2) = "01") Then
            CodNW = "NL"
            CodAux = CodAux & "N"
        ElseIf (Mid(Sequencia01, lLocal, 2) = "00") Then
            CodNW = "NE"
            CodAux = CodAux & "n"
        ElseIf (Mid(Sequencia01, lLocal, 2) = "10") Then
            CodNW = "WE"
            CodAux = CodAux & "w"
        ElseIf (Mid(Sequencia01, lLocal, 2) = "11") Then
            CodNW = "WL"
            CodAux = CodAux & "W"
        End If
        SequenciaNW = SequenciaNW & CodNW
        lLocal = lLocal + 1
    Next
    pobjstc.LinhaWN = SequenciaNW
    pobjstc.LinhaWN_Fonte = CodAux

    mfMontarCodigoBarra = True

    Exit Function

ControleErros:
    mfMontarCodigoBarra = False
    UltimoErro = Err.Description
End Function

Private Function mfCodigoRepresentacao(pcodigo As String) As String
    Dim CodBi As String

    If pcodigo = 0 Then
        CodBi = "00110"
    ElseIf pcodigo = 1 Then
        CodBi = "10001"
    ElseIf pcodigo = 2 Then
        CodBi = "01001"
    ElseIf pcodigo = 3 Then
        CodBi = "11000"
    ElseIf pcodigo = 4 Then
        CodBi = "00101"
    ElseIf pcodigo = 5 Then
        CodBi = "10100"
    ElseIf pcodigo = 6 Then
        CodBi = "01100"
    ElseIf pcodigo = 7 Then
        CodBi = "00011"
    ElseIf pcodigo = 8 Then
        CodBi = "10010"
    ElseIf pcodigo = 9 Then
        CodBi = "01010"
    End If

    mfCodigoRepresentacao = CodBi

End Function

Private Function mfRetornarDivNumItau(pobjstc As stcA99V009B) As String
    Dim lAuxiliar As String

    lAuxiliar = Format(pobjstc.AgenciaCedente, "0000")
    lAuxiliar = lAuxiliar & Left(Format(pobjstc.ContaCedente, "00000"), 5)
    lAuxiliar = lAuxiliar & Format(pobjstc.carteira, "000")
    lAuxiliar = lAuxiliar & Format(pobjstc.NossoNumero, "00000000")
    mfRetornarDivNumItau = Modulo_10(lAuxiliar)

End Function
Private Function mfRetornarDivNumReal(pobjstc As stcA99V009B) As String
    Dim lAuxiliar As String

    '11 00,38
    '12 01,38
    '13 02,38
    '14 03,38
    '15 04,38
    '16 05,38
    '17 06,38
    '18 07,38   <==
    '19 08,38
    '20 09,38
    '21 10,38

    lAuxiliar = "18"
    lAuxiliar = lAuxiliar & Format(pobjstc.NossoNumero, "00000000000")
    lAuxiliar = lAuxiliar & Format(pobjstc.AgenciaCedente, "0000")
    lAuxiliar = lAuxiliar & Format(pobjstc.ContaCedente, "0000000")
    mfRetornarDivNumReal = Modulo_10(lAuxiliar)

End Function
Private Function Modulo_10(ByVal pEntrada As String) As Integer
    Dim lAuxiliar As String     'Auxiliar.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lResto As Integer    'Resto.
    Dim lResultado As String      'Resultado.
    Dim lPesos As String  'Tabela de pesos.
    Dim lTotal As Double  'Total.

'2. Posiciona controles.
    lAuxiliar = pEntrada
    lPesos = 2
    lTotal = 0

    '3. Aplica peso.
    For lI = Len(pEntrada) To 1 Step -1
        lResultado = Format(CByte(Mid(lAuxiliar, lI, 1)) * lPesos, "00")
        lTotal = lTotal + CLng(Mid(lResultado, 1, 1)) + CLng(Mid(lResultado, 2, 1))

        lPesos = lPesos - 1

        If lPesos = 0 Then
            lPesos = 2
        End If

    Next lI

    '4. Divide por 10 e obtém resto.
    lResto = lTotal Mod 10

    '5. Posiciona dígito.
    Select Case lResto
    Case 0
        Modulo_10 = 0
    Case Else
        Modulo_10 = 10 - lResto
    End Select

End Function
Private Function NossoDigYasAnt(pNumero As String) As String
    Dim lPeso As Integer
    Dim lTotal As Integer
    Dim lContador As Integer
    Dim lModulo As Integer
    Dim lDigito As Integer
    Dim lResu As Integer
    Dim lI As Integer

    lDigito = 0
    lPeso = 9
    lTotal = 108
    lContador = Len(pNumero)
    lI = 1
    While (lI <= lContador)
        lResu = Mid(pNumero, lI, 1) * lPeso
        lI = lI + 1
        lTotal = lTotal + lResu
        Select Case lPeso
        Case 9
            lPeso = 7
        Case 7
            lPeso = 3
        Case 3
            lPeso = 1
        Case 1
            lPeso = 9
        End Select
    Wend
    lModulo = lTotal Mod 10
    If lModulo = 0 Then
        lDigito = 0
    Else
        lDigito = 10 - lModulo
    End If
    NossoDigYasAnt = lDigito
End Function

