VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA99V001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A99V001"
End Property

Public Function DIG_BLOQUETO(ByVal pBloqueto As String, ByRef pDigito As String, _
                             ByRef pMensagem As String) As Boolean
'Função: calcula dígito do bloqueto (copiado do programa P16C031).

'Parâmetro de entrada...pBloqueto...Campo com número do bloqueto.
'Parâmetros de saída....pDigito.....Dígito calculado.
'                       pMensagem...Mensagem.

    Dim lAuxiliar As String     'Auxiliar.
    Dim lPeso As String    'Peso.
    Dim lPosicao As Integer   'Posição.
    Dim lResto As Integer    'Resto.
    Dim lSoma As Integer    'Soma.

    DIG_BLOQUETO = False

    '1. Multiplica dígito pelo peso e acumula.
    lAuxiliar = Mid$(pBloqueto, 2, 7)
    lPeso = "9731973"
    lSoma = 0
    For lPosicao = 1 To 7
        lSoma = lSoma + val(Mid$(lAuxiliar, lPosicao, 1)) * val(Mid$(lPeso, lPosicao, 1))
    Next lPosicao
    lSoma = lSoma + 108
    lResto = lSoma Mod 10
    Select Case lResto
    Case 0
        pDigito = "0"
    Case Else
        pDigito = Format$(10 - lResto, "0")
    End Select

    DIG_BLOQUETO = True
End Function
Public Function DIG_BOSTON(ByVal pEntrada As String, ByRef pDigito As String, _
                           ByRef pMensagem As String) As Boolean
'Função: calcula dígito de conta do Banco de Boston.

'Parâmetro de entrada...pEntrada....Campo a ser calculado dígito verificador.
'Parâmetros de saída....pDigito.....Dígito calculado.
'                       pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_BOSTON = False

    '1. Verifica parâmetro de entrada.
    If Not M46V999.gfPreenchido(pEntrada) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco de " & _
                    "Boston não preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pEntrada) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco de " & _
                    "Boston não numérico - " & pEntrada & ")."
        Exit Function
    End If

    If Len(pEntrada) <> 7 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco de " & _
                    "Boston com tamanho diferente de 7 posições - " & pEntrada & ")."
        Exit Function
    End If

    '2. Posiciona controles.
    lAuxiliar = Trim$(pEntrada)
    lTabPesos = "4758243"
    lTotal = 0

    '3. Aplica peso.
    For lI = 1 To 7
        lTotal = lTotal + (CByte(Mid$(lAuxiliar, lI, 1)) * CByte(Mid$(lTabPesos, lI, 1)))
    Next lI

    '4. Divide por 11 e obtém resto.
    lResto = lTotal Mod 11

    '5. Posiciona dígito.
    Select Case lResto
    Case 0, 1
        pDigito = 0
    Case Else
        pDigito = 11 - lResto
    End Select

    DIG_BOSTON = True
End Function
Public Function DIG_BRADESCO(ByVal pEntrada As String, ByRef pDigito As String, _
                             ByRef pMensagem As String) As Boolean
'Função: calcula dígito módulo 11 - BRADESCO.

'Parâmetro de entrada...pEntrada....Número a ser calculado dígito verificador (agência ou conta).
'Parâmetros de saída....pDigito.....Dígito calculado.
'                       pMensagem...Mensagem.

    Dim lAuxCalculo As Byte         'Auxiliar de cálculo.
    Dim lAuxiliar As String     'Auxiliar.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lPeso As Byte   'Peso.
    Dim lPosicao As Byte      'Posição.
    Dim lResto As Byte    'Resto.
    Dim lTabPesos As String     'Tabela de pesos.
    Dim lTamanho As Byte      'Tamanho do campo.
    Dim lTotal As Double  'Total.

    DIG_BRADESCO = False

    '1. Verifica parâmetro de entrada.
    If Not M46V999.gfPreenchido(pEntrada) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência ou número da conta " & _
                    "corrente do Banco Bradesco não preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pEntrada) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência ou número da conta " & _
                    "corrente do Banco Bradesco não numérico - " & pEntrada & ")."
        Exit Function
    End If

    '2. Posiciona controle.
    lAuxiliar = Trim$(pEntrada)
    lTamanho = Len(lAuxiliar)
    lTabPesos = "234567"
    lTotal = 0

    '3. Aplica peso.
    For lI = 1 To lTamanho
        lAuxCalculo = CByte(Mid(lAuxiliar, lTamanho - lI + 1, 1))
        If lI > 6 Then
            lPosicao = lI Mod 6
        Else
            lPosicao = lI
        End If
        lPeso = CByte(Mid(lTabPesos, lPosicao, 1))
        lTotal = lTotal + (lAuxCalculo * lPeso)
    Next lI

    '4. Multiplica resultado por 10, divide por 11 e obtém resto.
    lTotal = lTotal * 10
    lResto = lTotal Mod 11
    If lResto = 10 Then
        lResto = 0
    End If

    '5. Posiciona dígito.
    pDigito = lResto

    DIG_BRADESCO = True
End Function
Public Function DIG_CNPJ(ByVal pCNPJ As String, ByRef pDigito As String, ByRef pMensagem As String) _
       As Boolean
'Função: calcula dígito do CNPJ (cópia da função QSf_iCGC).

'Parâmetro de entrada...pCNPJ.......Campo com CNPJ.
'Parâmetros de saída....pDigito.....Dígito calculado.
'                       pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lDigCalculado As Integer    'Dígito calculado.
    Dim lDigInformado As String    'Dígito informado.
    Dim lPeso As String    'Peso.

    DIG_CNPJ = False

    '1. Verifica CNPJ.
    pCNPJ = Trim$(pCNPJ)
    Select Case Left$(pCNPJ, 10)
    Case "12.345.678", "11.111.111", "22.222.222", "33.333.333", "44.444.444", _
         "55.555.555", "66.666.666", "77.777.777", "88.888.888", "99.999.999"
        pMensagem = "FGE0054 - CNPJ inválido: " & pCNPJ & "."
        Exit Function
    End Select

    '2. Posiciona campos.
    lAuxiliar = Left(pCNPJ, 16)
    lDigInformado = Right$(pCNPJ, 2)

    '3. Calcula primeiro dígito do CNPJ.
    lPeso = "05040003020900080706000504030200"
    If DIG_MODULO11(lAuxiliar, lPeso, lDigCalculado, pMensagem) = False Then
        Exit Function
    End If
    lAuxiliar = lAuxiliar & Format(lDigCalculado, "0")

    '4. Calcula segundo dígito do CNPJ.
    lPeso = "0605000403020009080700060504030002"
    If DIG_MODULO11(lAuxiliar, lPeso, lDigCalculado, pMensagem) = False Then
        Exit Function
    End If
    lAuxiliar = lAuxiliar & Format(lDigCalculado, "0")

    '5. Posiciona dígito do CNPJ.
    pDigito = Right$(lAuxiliar, 2)

    '6. Verifica dígito.
    If pDigito <> lDigInformado Then
        pMensagem = "FGE0055 - CNPJ com dígito inválido: " & pCNPJ & "."
        Exit Function
    End If

    DIG_CNPJ = True
End Function

Public Function DIG_CodIdeApolice(ByVal pCodIdeApolice As String, ByRef pNumApolice As String, _
                                  ByRef pNumItem As String, ByRef pDigito As String, _
                                  ByRef pMensagem As String) As Boolean
'Função: calcula dígito do código de identificação da apólice.

'Parâmetro de entrada...pCodIdeApolice...Código de identificação da apólice
'Parâmetros de saída....pNumApolice......Número da apólice.
'                       pNumItem.........Número do item.
'                       pDigito..........Dígito calculado.
'                       pMensagem........Mensagem.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lCaracter As String    'Caracter.
    Dim lCodSeguradora As String  'Código da seguradora.
    Dim lI As Integer             'Utilizado com For...Next.
    Dim lParcela As Integer    'Valor da parcela.
    Dim lPosicao As Byte    'Utilizado com InStr.
    Dim lTotal As Integer    'Total.
    Dim lPeso As Integer    'Peso a ser multiplicado.
    Dim lResto As Integer    'Resto da divisão do módulo 11.

    DIG_CodIdeApolice = False

    '1. Verifica campo de entrada - preenchimento e tamanho.
    If Not M46V999.gfPreenchido(pCodIdeApolice) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código de identificação da apólice não " & _
                    "preenchido)."
        Exit Function
    End If
    lAuxiliar = Trim$(pCodIdeApolice)
    Do
        lPosicao = InStr(lAuxiliar, " ")
        If lPosicao = 0 Then
            Exit Do
        End If
        lAuxiliar = Left$(lAuxiliar, lPosicao - 1) & Mid$(lAuxiliar, lPosicao + 1)
    Loop
    If Len(lAuxiliar) <> 14 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código de identificação da apólice com " & _
                    "tamanho diferente de 14 posições - " & pCodIdeApolice & ")."
        Exit Function
    End If

    '2. Verifica código de identificação da apólice - código da seguradora.
    lCodSeguradora = Left$(lAuxiliar, 3)
    If Not IsNumeric(lCodSeguradora) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código de identificação da apólice com " & _
                    "código da seguradora não numérico - " & pCodIdeApolice & ")."
        Exit Function
    End If
    If val(lCodSeguradora) = 0 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código de identificação da apólice com " & _
                    "código da seguradora zerado - " & pCodIdeApolice & ")."
        Exit Function
    End If

    '3. Calcula dígito novo (Módulo 11).
    'a) Inicializa campo de total.
    lTotal = 0
    lPeso = 6

    'b) Para cada caracter, obtém valor da parcela (código ASCII), multiplica pelo número de ordem e acumula.

    For lI = 1 To 13
        lCaracter = UCase$(Mid$(lAuxiliar, lI, 1))
        Select Case lCaracter
        Case "0"
            lParcela = 48
        Case "1"
            lParcela = 49
        Case "2"
            lParcela = 50
        Case "3"
            lParcela = 51
        Case "4"
            lParcela = 52
        Case "5"
            lParcela = 53
        Case "6"
            lParcela = 54
        Case "7"
            lParcela = 55
        Case "8"
            lParcela = 56
        Case "9"
            lParcela = 57
        Case "A"
            lParcela = 65
        Case "B"
            lParcela = 66
        Case "C"
            lParcela = 67
        Case "D"
            lParcela = 68
        Case "E"
            lParcela = 69
        Case "F"
            lParcela = 70
        Case "G"
            lParcela = 71
        Case "H"
            lParcela = 72
        Case "I"
            lParcela = 73
        Case "J"
            lParcela = 74
        Case "K"
            lParcela = 75
        Case "L"
            lParcela = 76
        Case "M"
            lParcela = 77
        Case "N"
            lParcela = 78
        Case "O"
            lParcela = 79
        Case "P"
            lParcela = 80
        Case "Q"
            lParcela = 81
        Case "R"
            lParcela = 82
        Case "S"
            lParcela = 83
        Case "T"
            lParcela = 84
        Case "U"
            lParcela = 85
        Case "V"
            lParcela = 86
        Case "W"
            lParcela = 87
        Case "X"
            lParcela = 88
        Case "Y"
            lParcela = 89
        Case "Z"
            lParcela = 90
        Case Else
            pMensagem = "FTE0047 - Erro no cálculo de dígito (código de identificação da apólice " & _
                        "com caracter (" & lCaracter & ") inválido - " & pCodIdeApolice & ")."
            Exit Function
        End Select
        lTotal = lTotal + (lParcela * lPeso)
        lPeso = lPeso - 1
        If lPeso = 1 Then
            lPeso = 9
        End If
    Next lI

    lResto = lTotal Mod 11

    Select Case lResto
    Case 0, 1       'Resto igual a zero ou um: dígito = zero.
        pDigito = 0
    Case Else       'Outro resto: dígito = diferença entre 10 e resto da divisão.
        pDigito = 11 - (lResto)
    End Select

    If pDigito = Mid$(lAuxiliar, 14, 1) Then
        DIG_CodIdeApolice = True
        Exit Function
    End If

    '4. Calcula dígito velho.
    'a) Inicializa campo de total e parcela.
    lTotal = 0
    lParcela = 0
    'b) Para cada caracter, obtém valor da parcela (código ASCII), multiplica pelo número de ordem e
    '   acumula.
    For lI = 1 To 13
        lCaracter = UCase$(Mid$(lAuxiliar, lI, 1))
        Select Case lCaracter
        Case "0"
            lParcela = 48
        Case "1"
            lParcela = 49
        Case "2"
            lParcela = 50
        Case "3"
            lParcela = 51
        Case "4"
            lParcela = 52
        Case "5"
            lParcela = 53
        Case "6"
            lParcela = 54
        Case "7"
            lParcela = 55
        Case "8"
            lParcela = 56
        Case "9"
            lParcela = 57
        Case "A"
            lParcela = 65
        Case "B"
            lParcela = 66
        Case "C"
            lParcela = 67
        Case "D"
            lParcela = 68
        Case "E"
            lParcela = 69
        Case "F"
            lParcela = 70
        Case "G"
            lParcela = 71
        Case "H"
            lParcela = 72
        Case "I"
            lParcela = 73
        Case "J"
            lParcela = 74
        Case "K"
            lParcela = 75
        Case "L"
            lParcela = 76
        Case "M"
            lParcela = 77
        Case "N"
            lParcela = 78
        Case "O"
            lParcela = 79
        Case "P"
            lParcela = 80
        Case "Q"
            lParcela = 81
        Case "R"
            lParcela = 82
        Case "S"
            lParcela = 83
        Case "T"
            lParcela = 84
        Case "U"
            lParcela = 85
        Case "V"
            lParcela = 86
        Case "W"
            lParcela = 87
        Case "X"
            lParcela = 88
        Case "Y"
            lParcela = 89
        Case "Z"
            lParcela = 90
        Case Else
            pMensagem = "FTE0047 - Erro no cálculo de dígito (código de identificação da apólice " & _
                        "com caracter (" & lCaracter & ") inválido - " & pCodIdeApolice & ")."
            Exit Function
        End Select
        lTotal = lTotal + (lParcela * lI)
    Next lI
    'c) Se valor acumulado > 10, soma caracter a caracter do valor acumulado até chegar num valor menor
    '   que 10.
    Do
        If lTotal < 10 Then
            Exit Do
        End If
        lAuxiliar = lTotal
        lTotal = 0
        For lI = 1 To Len(lAuxiliar)
            lTotal = lTotal + val(Mid$(lAuxiliar, lI, 1))
        Next lI
    Loop
    'd) Posiciona valor do dígito.
    pDigito = lTotal

    '5. Obtém número de apólice e número do item (apenas para seguradora 641 - Yasuda).
    'a) Inicializa campos.
    pNumApolice = "0000000000"
    pNumItem = "0000"
    'b) Se seguradora diferente de 641, encerra processo.
    If val(lCodSeguradora) <> 641 Then
        DIG_CodIdeApolice = True
        Exit Function
    End If
    'c) Posiciona número da apólice e número do item.
    If IsNumeric(Mid$(pCodIdeApolice, 4, 1)) Then
        If val(Mid$(pCodIdeApolice, 4, 6)) > 980000 Then
            pNumApolice = "0032" & Mid$(pCodIdeApolice, 4, 6)
        Else
            pNumApolice = "0031" & Mid$(pCodIdeApolice, 4, 6)
        End If
        pNumItem = Mid$(pCodIdeApolice, 10, 4)
    Else
        Select Case Mid$(pCodIdeApolice, 4, 1)
        Case "A"
            pNumApolice = "40310" & Mid$(pCodIdeApolice, 5, 5)
            pNumItem = Mid$(pCodIdeApolice, 10, 4)
        Case "B"
            pNumApolice = "40311" & Mid$(pCodIdeApolice, 5, 5)
            pNumItem = Mid$(pCodIdeApolice, 10, 4)
        Case "C"
            pNumApolice = "40312" & Mid$(pCodIdeApolice, 5, 5)
            pNumItem = Mid$(pCodIdeApolice, 10, 4)
        Case "X"
            pNumApolice = "50310" & Mid$(pCodIdeApolice, 5, 5)
            pNumItem = Mid$(pCodIdeApolice, 10, 4)
        Case "Y"
            pNumApolice = "50311" & Mid$(pCodIdeApolice, 5, 5)
            pNumItem = Mid$(pCodIdeApolice, 10, 4)
        Case "Z"
            pNumApolice = "50312" & Mid$(pCodIdeApolice, 5, 5)
            pNumItem = Mid$(pCodIdeApolice, 10, 4)
        End Select
    End If

    DIG_CodIdeApolice = True

End Function
Public Function DIG_CPF(ByVal pCPF As String, ByRef pDigito As String, ByRef pMensagem As String) _
       As Boolean
'Função: calcula dígito do CPF (cópia da função QSf-iCPF).

'Parâmetro de entrada...pCPF........Campo com CPF no formato numérico.
'Parâmetros de saída....pDigito.....Dígito calculado.
'                       pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lDigCalculado As Integer    'Dígito calculado.
    Dim lDigInformado As String    'Dígito informado.
    Dim lPeso As String    'Peso.

    DIG_CPF = False

    '1. Verifica CPF.
    pCPF = Trim$(pCPF)
    Select Case Left$(pCPF, 11)
    Case "123.456.789", "111.111.111", "222.222.222", "333.333.333", "444.444.444", _
         "555.555.555", "666.666.666", "777.777.777", "888.888.888", "999.999.999"
        pMensagem = "FGE0052 - CPF inválido: " & pCPF & "."
        Exit Function
    End Select

    '2. Posiciona campos auxiliar.
    lAuxiliar = Left(pCPF, 12)
    lDigInformado = Right$(pCPF, 2)

    '3. Calcula primeiro dígito do CPF.
    lPeso = "100908000706050004030200"
    If DIG_MODULO11(lAuxiliar, lPeso, lDigCalculado, pMensagem) = False Then
        Exit Function
    End If
    lAuxiliar = lAuxiliar & Format(lDigCalculado, "0")

    '4. Calcula segundo dígito do CPF.
    lPeso = "11100900080706000504030002"
    If DIG_MODULO11(lAuxiliar, lPeso, lDigCalculado, pMensagem) = False Then
        Exit Function
    End If
    lAuxiliar = lAuxiliar & Format(lDigCalculado, "0")

    '5. Posiciona dígito do CPF.
    pDigito = Right$(lAuxiliar, 2)

    '6. Verifica dígito.
    If pDigito <> lDigInformado Then
        pMensagem = "FGE0053 - CPF com dígito inválido: " & pCPF & "."
        Exit Function
    End If

    DIG_CPF = True
End Function
Public Function DIG_ITAU(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                         ByRef pMensagem As String) As Boolean
'Função: calcula dígito módulo 10 - Itaú.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta corrente.
'Parâmetros de saída.....pDigito.....Dígito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String     'Auxiliar.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lResultado As String      'Resultado.
    Dim lTabPesos As String     'Tabela de pesos.
    Dim lTotal As Double  'Total.

    DIG_ITAU = False

    '1. Verifica se parâmetros de entrada estão preenchidos e válidos.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Itaú não " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Itaú não " & _
                    "numérico - " & pAgencia & ")."
        Exit Function
    End If

    If Len(pAgencia) <> 4 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Itaú com " & _
                    "tamanho diferente de 4 posições - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco Itaú não " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco Itaú não " & _
                    "numérico - " & pConta & ")."
        Exit Function
    End If

    If Len(pConta) <> 5 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco Itaú com " & _
                    "tamanho diferente de 5 posições - " & pConta & ")."
        Exit Function
    End If

    '2. Posiciona controles.
    lAuxiliar = pAgencia & pConta
    lTabPesos = "212121212"
    lTotal = 0

    '3. Aplica peso.
    For lI = 1 To 9
        lResultado = Format$(CByte(Mid(lAuxiliar, lI, 1)) * CByte(Mid(lTabPesos, lI, 1)), "00")
        lTotal = lTotal + CByte(Mid(lResultado, 1, 1)) + CByte(Mid(lResultado, 2, 1))
    Next lI

    '4. Divide por 10 e obtém resto.
    lResto = lTotal Mod 10

    '5. Posiciona dígito.
    Select Case lResto
    Case 0
        pDigito = 0
    Case Else
        pDigito = 10 - lResto
    End Select

    DIG_ITAU = True
End Function
Public Function DIG_HSBC(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                         ByRef pMensagem As String) As Boolean
'Função: calcula dígito módulo 10 - HSBC.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta corrente.
'Parâmetros de saída.....pDigito.....Dígito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String     'Auxiliar.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lResultado As String      'Resultado.
    Dim lTabPesos As String     'Tabela de pesos.
    Dim lTotal As Double  'Total.

    DIG_HSBC = False

    '1. Verifica se parâmetros de entrada estão preenchidos e válidos.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco HSBC não " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco HSBC não " & _
                    "numérico - " & pAgencia & ")."
        Exit Function
    End If

    If Len(pAgencia) <> 4 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco HSBC com " & _
                    "tamanho diferente de 4 posições - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco HSBC não " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco HSBC não " & _
                    "numérico - " & pConta & ")."
        Exit Function
    End If

    If Len(pConta) <> 10 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco HSBC com " & _
                    "tamanho diferente de 10 posições - " & pConta & ")."
        Exit Function
    End If

    'If Left(pConta, 4) <> pAgencia Then
    '    pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco HSBC não corresponde a agência informada)."
    '    Exit Function
    'End If

    '2. Posiciona controles.
    lAuxiliar = pAgencia & Right(pConta, 6)
    lTabPesos = "8923456789"
    lTotal = 0

    '3. Aplica peso.
    For lI = 1 To 10
        lResultado = Format$(CByte(Mid(lAuxiliar, lI, 1)) * CByte(Mid(lTabPesos, lI, 1)), "00")
        lTotal = lTotal + lResultado
    Next lI

    '4. Divide por 10 e obtém resto.
    lResto = lTotal Mod 11

    '5. Posiciona dígito.
    pDigito = lResto
    Select Case lResto
    Case 0, 10
        pDigito = 0
    End Select

    DIG_HSBC = True
End Function

Public Function DIG_MODULO11(ByVal pCampo As String, ByVal pPesos As String, _
                             ByRef pDigito As Integer, ByRef pMensagem As String) As Boolean
'Função: calcula dígito utilizando peso fornecido e algorítimo para módulo 11.

'Parâmetros de entrada...pCampo......Campo a calcular dígito.
'                        pPesos......Pesos para efetuar cálculo.
'Parâmetros de saída.....pDigito.....Dígito calculado.
'                        pMensagem...Mensagem.

'Observação: para cada dígito de pCampo existirão dois em pPesos.
'Exemplo: se peso é 1, 2, 3, 4, informar 01020304.

'Exemplos da função: DIG_MODULO11("098031208","100908070605040302") = 6.
'                    DIG_MODULO11("0980312086","01100908070605040302") = 7.

'Se a posição do peso a ser aplicado é igual a zero, o sistema não considera a posição do campo a ser
'calculado dígito correspondente (mesmo que não seja numérico).

    Dim lPosicao As Integer    'Posição.
    Dim lResto As Integer    'Resto.
    Dim lSoma As Integer    'Soma.
    Dim lTamCampo As Integer    'Tamanho do campo de entrada.
    Dim lTamPesos As Integer    'Tamanho do campo de pesos.

    DIG_MODULO11 = False

    '1. Verifica se campos de entrada são numéricos.
    If Not M46V999.gfPreenchido(pCampo) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (módulo 11 - campo a ser calculado dígito " & _
                    "não preenchido)."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pPesos) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (módulo 11 - campo de pesos não preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pPesos) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (módulo 11 - campo de pesos não numérico)."
        Exit Function
    End If

    pCampo = Trim$(pCampo)
    pPesos = Trim$(pPesos)
    lTamCampo = Len(pCampo)
    lTamPesos = Len(pPesos)

    If lTamPesos < (lTamCampo * 2) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (módulo 11 - campo de pesos (" & pPesos & _
                    ") incompatível com campo a ser calculado dígito (" & pCampo & "))."
        Exit Function
    End If

    '2. Multiplica dígito pelo peso e acumula.
    lSoma = 0
    For lPosicao = 1 To Len(pCampo)
        If val(Mid(pPesos, 1 + ((lPosicao - 1) * 2), 2)) <> 0 Then
            If Not IsNumeric(Mid(pCampo, lPosicao, 1)) Then
                pMensagem = "FTE0047 - Erro no cálculo de dígito (módulo 11 - campo a ser calculado " & _
                            "dígito não numérico - " & pCampo & ")."
                Exit Function
            End If
        End If
        lSoma = lSoma + (val(Mid(pCampo, lPosicao, 1)) * val(Mid(pPesos, 1 + ((lPosicao - 1) * 2), 2)))
    Next lPosicao

    '3. Retorno da função
    lResto = lSoma Mod 11
    Select Case lResto
    Case 0, 1       'Resto igual a zero ou um: dígito = zero.
        pDigito = 0
    Case Else       'Outro resto: dígito = diferença entre 10 e resto da divisão.
        pDigito = 11 - (lResto)
    End Select

    DIG_MODULO11 = True
End Function
Public Function DIG_SEGSUL(ByVal pCampo As String, ByVal pPesos As String, _
                           ByRef pDigito As Integer, ByRef pMensagem As String) As Boolean
'Função: calcula dígito utilizando peso fornecido e algorítimo para Segsul.

'Parâmetros de entrada...pCampo......Campo a calcular dígito.
'                        pPesos......Pesos para efetuar cálculo.
'Parâmetros de saída.....pDigito.....Dígito calculado.
'                        pMensagem...Mensagem.

'Observação: para cada dígito de pCampo existirão dois em pPesos.
'Exemplo: se peso é 1, 2, 3, 4, informar 01020304.

'Exemplos da função: DIG_SEGSUL("098031208","100908070605040302") = 6.
'                    DIG_SEGSUL("0980312086","01100908070605040302") = 7.

    Dim lAcum As Integer    'Acumulador.
    Dim lPosicao As Integer    'Posição.
    Dim lResto As Integer    'Resto.
    Dim lSoma As Integer    'Soma.
    Dim lTamCampo As Integer    'Tamanho do campo de entrada.
    Dim lTamPesos As Integer    'Tamanho do campo de pesos.

    DIG_SEGSUL = False

    '1. Verifica se campos de entrada são numéricos.
    If Not M46V999.gfPreenchido(pCampo) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código do Segsul não preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pCampo) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código do Segsul não numérico - " & pCampo & _
                    ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pPesos) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (Segsul - campo de pesos não preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pPesos) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (Segsul - campo de pesos não numérico - " & _
                    pPesos & ")."
        Exit Function
    End If

    pCampo = Trim$(pCampo)
    pPesos = Trim$(pPesos)
    lTamCampo = Len(pCampo)
    lTamPesos = Len(pPesos)

    '2. Multiplica dígito pelo peso e acumula.
    lSoma = 0
    lAcum = 0
    For lPosicao = 1 To Len(pCampo)
        lSoma = Format((val(Mid(pCampo, lPosicao, 1)) * val(Mid(pPesos, lPosicao, 1))), "00")
        lAcum = lAcum + (val(Mid(lSoma, 1, 1)) + val(Mid(lSoma, 2, 1)))
    Next lPosicao

    '3. Retorno da função
    lResto = lAcum Mod 11
    Select Case lResto
    Case 0, 1       'Resto igual a zero ou um: dígito = zero.
        pDigito = 0
    Case Else       'Outro resto: dígito = diferença entre 10 e resto da divisão.
        pDigito = 11 - (lResto)
    End Select

    DIG_SEGSUL = True
End Function
Public Function DIG_SEGSUL1(ByVal pCampo As String, ByVal pPesos As String, _
                            ByRef pDigito As String, ByRef pMensagem As String) As Boolean
'Função: calcula dígito utilizando peso fornecido e algorítimo para Segsul no P28V200.

'Parâmetros de entrada...pCampo......Campo a calcular dígito.
'                        pPesos......Pesos para efetuar cálculo.
'Parâmetros de saída.....pDigito.....Dígito calculado.
'                        pMensagem...Mensagem.

'Observação: para cada dígito de pCampo existirão dois em pPesos.
'Exemplo: se peso é 1, 2, 3, 4, informar 01020304.

'Exemplos da função: DIG_SEGSUL1("098031208","100908070605040302") = 6.
'                    DIG_SEGSUL1("0980312086","01100908070605040302") = 7.

    Dim lAcum As Integer    'Acumulador.
    Dim lPosicao As Integer    'Posição.
    Dim lResto As Integer    'Resto.
    Dim lSoma As Integer    'Soma.
    Dim lTamCampo As Integer    'Tamanho do campo de entrada.
    Dim lTamPesos As Integer    'Tamanho do campo de pesos.

    DIG_SEGSUL1 = False

    '1. Verifica se campos de entrada são numéricos.
    If Not M46V999.gfPreenchido(pCampo) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código do Segsul não preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pCampo) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código do Segsul não numérico - " & pCampo & _
                    ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pPesos) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (Segsul - campo de pesos não preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pPesos) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (Segsul - campo de pesos não numérico - " & _
                    pPesos & ")."
        Exit Function
    End If

    pCampo = Trim$(pCampo)
    pPesos = Trim$(pPesos)
    lTamCampo = Len(pCampo)
    lTamPesos = Len(pPesos)

    '2. Multiplica dígito pelo peso e acumula.
    lSoma = 0
    lAcum = 0
    For lPosicao = 1 To Len(pCampo)
        lSoma = Format((val(Mid(pCampo, lPosicao, 1)) * val(Mid(pPesos, lPosicao, 1))), "00")
        lAcum = lAcum + (val(Mid(lSoma, 1, 1)) + val(Mid(lSoma, 2, 1)))
    Next lPosicao

    '3. Retorno da função
    lResto = lAcum Mod 11
    Select Case lResto
    Case 0          'Resto igual a zero : digito = "A".
        pDigito = "A"
    Case 1          'Resto igual a um: dígito = zero.
        pDigito = 0
    Case Else       'Outro resto: dígito = diferença entre 10 e resto da divisão.
        pDigito = val(11 - (lResto))
    End Select

    DIG_SEGSUL1 = True
End Function
Public Function DIG_SUDAMERIS(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                              ByRef pMensagem As String) As Boolean
'Função: calcula dígito módulo 11 - conta corrente.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetros de saída.....pDigito.....Dígito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar...de 01 a 03 (03 posições)...Código do banco.
'           de 04 a 06 (03 posições)...Código da agência.
'           de 07 a 17 (11 posições)...Valor fixo.
'           de 18 a 26 (09 posições)...Número da conta.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_SUDAMERIS = False

    '1. Verifica parâmetros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Sudameris não " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Sudameris não " & _
                    "numérico - " & pAgencia & ")."
        Exit Function
    End If

    If Len(pAgencia) <> 3 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Sudameris com " & _
                    "tamanho diferente de 3 posições - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco " & _
                    "Sudameris não preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco " & _
                    "Sudameris não numérico - " & pConta & ")."
        Exit Function
    End If

    If Len(pConta) <> 9 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco " & _
                    "Sudameris com tamanho diferente de 9 posições - " & pConta & ")."
        Exit Function
    End If

    '2. Posiciona controles.
    lAuxiliar = "347" & pAgencia & "10000000000" & pConta
    lTabPesos = "32987654329876543298765432"
    lTotal = 0

    '3. Aplica peso.
    For lI = 1 To 26
        lTotal = lTotal + (CByte(Mid$(lAuxiliar, lI, 1)) * CByte(Mid$(lTabPesos, lI, 1)))
    Next lI

    '4. Divide por 11 e obtém resto.
    lResto = lTotal Mod 11

    '5. Posiciona dígito.
    Select Case lResto
    Case 0
        pDigito = 1
    Case 1
        pDigito = 0
    Case Else
        pDigito = 11 - lResto
    End Select

    DIG_SUDAMERIS = True
End Function
Public Function DIG_SUDAMERIS_NOVO(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                                   ByRef pMensagem As String) As Boolean
'Função: calcula dígito módulo 11 - conta corrente.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetros de saída.....pDigito.....Dígito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar...de 01 a 03 (03 posições)...Código do banco.
'           de 04 a 07 (04 posições)...Código da agência.
'           de 08 a 14 (07 posições)...Número da conta.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_SUDAMERIS_NOVO = False

    '1. Verifica parâmetros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Sudameris não " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Sudameris não " & _
                    "numérico - " & pAgencia & ")."
        Exit Function
    End If

    If Len(pAgencia) <> 4 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Sudameris com " & _
                    "tamanho diferente de 3 posições - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco " & _
                    "Sudameris não preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco " & _
                    "Sudameris não numérico - " & pConta & ")."
        Exit Function
    End If

    If Len(pConta) <> 7 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco " & _
                    "Sudameris com tamanho diferente de 7 posições - " & pConta & ")."
        Exit Function
    End If

    '2. Posiciona controles.
    lAuxiliar = pAgencia & pConta
    lTabPesos = "81472259395"
    lTotal = 0

    '3. Aplica peso.
    For lI = 1 To 11
        lTotal = lTotal + (CByte(Mid$(lAuxiliar, lI, 1)) * CByte(Mid$(lTabPesos, lI, 1)))
    Next lI

    '4. Divide por 11 e obtém resto.
    lResto = lTotal Mod 11

    '5. Posiciona dígito.
    Select Case lResto
    Case 0
        pDigito = 1
    Case 1
        pDigito = 0
    Case Else
        pDigito = 11 - lResto
    End Select

    DIG_SUDAMERIS_NOVO = True
End Function
Public Function DIG_REAL(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                         ByRef pMensagem As String) As Boolean
'Função: calcula dígito módulo 11 - conta corrente.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetros de saída.....pDigito.....Dígito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar...de 01 a 03 (03 posições)...Código do banco.
'           de 04 a 07 (04 posições)...Código da agência.
'           de 08 a 14 (07 posições)...Número da conta.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_REAL = False

    '1. Verifica parâmetros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Real não " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Real não " & _
                    "numérico - " & pAgencia & ")."
        Exit Function
    End If

    If Len(pAgencia) <> 4 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Real com " & _
                    "tamanho diferente de 3 posições - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco " & _
                    "Real não preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco " & _
                    "Real não numérico - " & pConta & ")."
        Exit Function
    End If

    If Len(pConta) <> 7 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Banco " & _
                    "Real com tamanho diferente de 7 posições - " & pConta & ")."
        Exit Function
    End If

    '2. Posiciona controles.
    lAuxiliar = pAgencia & pConta
    lTabPesos = "81472259395"
    lTotal = 0

    '3. Aplica peso.
    For lI = 1 To 11
        lTotal = lTotal + (CByte(Mid$(lAuxiliar, lI, 1)) * CByte(Mid$(lTabPesos, lI, 1)))
    Next lI

    '4. Divide por 11 e obtém resto.
    lResto = lTotal Mod 11

    '5. Posiciona dígito.
    Select Case lResto
    Case 0
        pDigito = 1
    Case 1
        pDigito = 0
    Case Else
        pDigito = 11 - lResto
    End Select

    DIG_REAL = True
End Function
Public Function DIG_BANCOBRASIL(ByVal pInfo As String, ByRef pDigito As String, _
                                ByRef pMensagem As String) As Boolean
'Função: calcula dígito módulo 11 - conta corrente.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetros de saída.....pDigito.....Dígito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar...de 01 a 03 (03 posições)...Código do banco.
'           de 04 a 07 (04 posições)...Código da agência.
'           de 08 a 14 (07 posições)...Número da conta.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As String    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_BANCOBRASIL = False

    '1. Verifica parâmetros de entrada.
    If Not M46V999.gfPreenchido(pInfo) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência ou conta do Banco Real não " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pInfo) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência ou conta do Banco Real não " & _
                    "numérico - " & pInfo & ")."
        Exit Function
    End If

    '2. Posiciona controles.
    lAuxiliar = pInfo
    lTabPesos = "23456789"
    lTotal = 0

    '3. Aplica peso.
    For lI = 1 To 8
        lTotal = lTotal + (CByte(Mid$(lAuxiliar, lI, 1)) * CByte(Mid$(lTabPesos, lI, 1)))
    Next lI

    '4. Divide por 11 e obtém resto.
    lResto = lTotal Mod 11

    '5. Posiciona dígito.
    If lResto = 10 Then lResto = "X"

    pDigito = lResto

    DIG_BANCOBRASIL = True
End Function

Public Function DIG_UNIBANCO(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                             ByRef pMensagem As String) As Boolean
'Função: calcula dígito módulo 10 - Unibanco.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetros de saída.....pDigito.....Dígito calculado.
'                        pMensagem...Mensagem.

'Exemplo para Banco Unibanco:
'
'3   9   8   1   0   3   5  1  4    ===>>> Agência 3 posições e conta 6 posições.
'2   1   2   1   2   1   2  1  2    ===>>> Seqüência de 2 e 1.
'6   9  16   1   0   3  10  1  8    ===>>> Total da multiplicação.
'
'Para consistir todos os números alteradas a seqüência de 2 e 1.
'
'0  3   9   8   0  1   0   3   5  1  4  ===>>> Agência 4 posições e conta 7 posições.
'   2   1   2      1   2   1   2  1  2  ===>>> Seqüência de 2 e 1.
'   6   9  16      1   0   3  10  1  8  ===>>> Total da multiplicação.

    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lResultado As String  'Resultado.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_UNIBANCO = False

    '1. Verifica se parâmetros de entrada estão preenchidos e válidos.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Unibanco não preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Unibanco não numérico " & _
                    "- " & pAgencia & ")."
        Exit Function
    End If

    If Len(pAgencia) <> 4 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Unibanco com tamanho " & _
                    "diferente de 4 posições - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Unibanco não " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Unibanco não " & _
                    "numérico - " & pConta & ")."
        Exit Function
    End If

    If Len(pConta) <> 6 Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da conta corrente do Unibanco com " & _
                    "tamanho diferente de 6 posições - " & pConta & ")."
        Exit Function
    End If

    '2. Posiciona controles.
    lTabPesos = "1212121212121"
    lTotal = 0

    '3. Aplica peso.
    'a) Para agência.
    For lI = 1 To 4
        lResultado = Format$(CByte(Mid(pAgencia, lI, 1)) * CByte(Mid(lTabPesos, lI, 1)), "00")
        lTotal = lTotal + CByte(Mid(lResultado, 1, 1)) + CByte(Mid(lResultado, 2, 1))
    Next lI
    'b) Para conta.
    For lI = 1 To 6
        lResultado = Format$(CByte(Mid(pConta, lI, 1)) * CByte(Mid(lTabPesos, lI, 1)), "00")
        lTotal = lTotal + CByte(Mid(lResultado, 1, 1)) + CByte(Mid(lResultado, 2, 1))
    Next lI

    '4. Divide por 10 e obtém resto.
    lResto = lTotal Mod 10

    '5. Posiciona dígito.
    Select Case lResto
    Case 0
        pDigito = 0
    Case Else
        pDigito = 10 - lResto
    End Select

    DIG_UNIBANCO = True
End Function
Public Function DIG_YASUDA(ByVal pEntrada As String, ByRef pDigito As String, _
                           ByRef pMensagem As String) As Boolean
'Função: calcula dígito módulo 11 - Yasuda.

'Parâmetro de entrada...pEntrada.....Número a ser calculado dígito verificador.
'Parâmetros de saída....pDigito.....Dígito calculado.
'                       pMensagem...Mensagem.

    Dim lAuxCalculo As Byte     'Auxiliar de cálculo.
    Dim lAuxiliar As String    'Auxiliar.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lPeso As Byte    'Peso.
    Dim lPosicao As Byte  'Posição.
    Dim lResto As Byte    'Resto.
    Dim lResultado As String  'Resultado.
    Dim lTamanho As Byte  'Tamanho do campo.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_YASUDA = False

    '1. Verifica se parâmetros de entrada estão preenchidos e válidos.
    If Not M46V999.gfPreenchido(pEntrada) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da Yasuda não preenchido)."
        Exit Function
    End If
    If Not IsNumeric(pEntrada) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (número da Yasuda não não numérico - " & _
                    pEntrada & ")."
        Exit Function
    End If

    '2. Posiciona controle.
    lAuxiliar = Trim$(pEntrada)
    lTamanho = Len(lAuxiliar)
    lTabPesos = "234567"
    lTotal = 0

    '3. Aplica pesos.
    For lI = 1 To lTamanho
        lAuxCalculo = CByte(Mid(lAuxiliar, lTamanho - lI + 1, 1))
        If lI > 6 Then
            lPosicao = lI Mod 6
        Else
            lPosicao = lI
        End If
        lPeso = CByte(Mid(lTabPesos, lPosicao, 1))
        lResultado = Format$(lAuxCalculo * lPeso, "00")
        lTotal = lTotal + CByte(Mid(lResultado, 1, 1)) + CByte(Mid(lResultado, 2, 1))
    Next lI

    '4. Divide resultado por 11 e obtém resto.
    lResto = lTotal Mod 11

    '5. Posiciona dígito.
    Select Case lResto
    Case 0
        pDigito = "A"
    Case 1
        pDigito = "0"
    Case Else
        pDigito = Format$(11 - lResto, "0")
    End Select

    DIG_YASUDA = True
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
Public Function DIG_BANESE(ByVal pAgencia As String, pConta As String, ByRef pDigito As String, _
                           ByRef pMensagem As String) As Boolean
'Função: calcula dígito módulo 11 - conta corrente.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetros de saída.....pDigito.....Dígito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar...de 01 a 03 (03 posições)...Código do banco.
'           de 04 a 07 (04 posições)...Código da agência.
'           de 08 a 14 (07 posições)...Número da conta.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As String    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.
    Dim lAgenciaC As String

    DIG_BANESE = False

    '1. Verifica parâmetros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Banese não " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Banese não " & _
                    "numérico - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da conta do Banco Banese não " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da conta do Banco Banese não " & _
                    "numérico - " & pConta & ")."
        Exit Function
    End If

    If Len(pAgencia) > 3 Then
        pAgencia = Right(pAgencia, 3)
    Else
        pAgencia = Format(pAgencia, "000")
    End If
    lAgenciaC = pAgencia
    Select Case pAgencia
    Case "002"
        lAgenciaC = "006"
    Case "003"
        lAgenciaC = "005"
    Case "004"
        lAgenciaC = "011"
    Case "005"
        lAgenciaC = "007"
    Case "006"
        lAgenciaC = "008"
    Case "007"
        lAgenciaC = "013"
    Case "008"
        lAgenciaC = "004"
    Case "009"
        lAgenciaC = "003"
    Case "010"
        lAgenciaC = "010"
    Case "011"
        lAgenciaC = "002"
    Case "012"
        lAgenciaC = "012"
    Case "013"
        lAgenciaC = "009"
    Case "014"
        lAgenciaC = "001"

    End Select
    '2. Posiciona controles.
    lAuxiliar = Format(val(lAgenciaC), "00") & Format(pConta, "00000000")
    lTabPesos = "3298765432"
    lTotal = 0

    '3. Aplica peso.
    For lI = 1 To 10
        lTotal = lTotal + (CByte(Mid$(lAuxiliar, lI, 1)) * CByte(Mid$(lTabPesos, lI, 1)))
    Next lI

    '4. Divide por 11 e obtém resto.
    lResto = lTotal Mod 11

    '5. Posiciona dígito.
    lResto = 11 - lResto
    If lResto = "10" Then lResto = "0"
    pDigito = lResto

    DIG_BANESE = True
End Function

Public Function DIG_SANTANDER(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                              ByRef pMensagem As String) As Boolean
'Função: calcula dígito conta corrente.

'Parâmetros de entrada...pAgencia....Código da agência.
'                        pConta......Número da conta.
'Parâmetros de saída.....pDigito.....Dígito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar...de 01 a 03 (03 posições)...Código do banco.
'           de 04 a 07 (04 posições)...Código da agência.
'           de 08 a 14 (07 posições)...Número da conta.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As String    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_SANTANDER = False

    '1. Verifica parâmetros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Santander não " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Santander não " & _
                    "numérico - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da conta do Banco Santander não " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da conta do Banco Santander não " & _
                    "numérico - " & pConta & ")."
        Exit Function
    End If

    '2. Posiciona controles.
    lAuxiliar = Format(pAgencia, "0000") & "00" & Format(pConta, "00000000")
    lTabPesos = "97310097131973"
    lTotal = 0

    '3. Aplica peso.
    For lI = 1 To 14
        lTotal = lTotal + Right((CByte(Mid$(lAuxiliar, lI, 1)) * CByte(Mid$(lTabPesos, lI, 1))), 1)
    Next lI

    '4. Divide por 11 e obtém resto.
    lResto = 10 - Right(lTotal, 1)

    '5. Posiciona dígito.
    If lResto = 10 Then lResto = "0"

    pDigito = lResto

    DIG_SANTANDER = True
End Function

'CyberTools - Ari - Task 3795:Implementar função da validação do CNH, Renavam, Safra e Bancoob - Setembro\2014
Public Function DIG_BANCOOB(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                             ByRef pMensagem As String) As Boolean
                             
'*M O D U L O   11  -  N A O  CONSIDERAR AGENCIA
'*ALGORITIMO E REGRA OBTIDOS DO PROG MARITIMA RER01S200
'*A CONTA BANCOOB[756]
'*TEM 9 DIGITOS + 1 DIG VERIFICADOR
'*5 7 6 1 0 0 0 0 1 - 6 Ex CONTA+DIGITO BANCOOB
'*10 9 8 7 6 5 4 3 2 = MULTIPLICADOR BANCOOB BYTE A BYTE
'*------------------------------------------------
'*(5*10=50) + (7*9=63) + (6*8=48) + (1*7=7 ) + (0*6=0 ) +
'*(0*5 =0 ) + (0*4=0 ) + (0*3=0 ) + (1*2=2 )  = 170
'*       170 MOD 11  = 5 OU
'*       OU|170  / 11   = 15,4545
'*       |15   * 11   = 165     170 - 165 = 5
'*       11  -  5   = 6 <<<<<DIGITO DA CONTA

'BANCOOB - Desconsiderado um zero à esquerda '*TEM 9 DIGITOS + 1 DIG VERIFICADOR
'756         0001      0070000602        8 - OK
'756         4027      0002276001        6 - OK
'756         4027      0018257001        0 - OK
'756         4027      0028884320        7 - OK
'756         4033      0000002510        0 - OK
'756         4037      0000001138        0 - OK
'756         4321      0002002095        3 - OK
'756         4312      0002570001        4 - OK
'756         4487      0000000073        6 - OK
'756         5311      0000000345        0 - OK

                             
Dim iContadorFor As Integer
Dim iSoma As Integer
Dim iMultiplicador(9) As Integer
Dim iResto As Integer
Dim div11 As Integer
Dim mult11 As Integer

    DIG_BANCOOB = False

    '1. Verifica parâmetros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Bancoob não " & _
                    "preenchido)."
        Exit Function
    End If
    
    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Bancoob não " & _
                    "numérico - " & pAgencia & ")."
        Exit Function
    End If
        
    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da conta do Banco Bancoob não " & _
                    "preenchido)."
        Exit Function
    End If
    
    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da conta do Banco Bancoob não " & _
                    "numérico - " & pConta & ")."
        Exit Function
    End If

    iMultiplicador(0) = 10
    iMultiplicador(1) = 9
    iMultiplicador(2) = 8
    iMultiplicador(3) = 7
    iMultiplicador(4) = 6
    iMultiplicador(5) = 5
    iMultiplicador(6) = 4
    iMultiplicador(7) = 3
    iMultiplicador(8) = 2
                             
    iSoma = 0
    iResto = 0
    pDigito = ""
    div11 = 0
    mult11 = 0
    
    For iContadorFor = 0 To 8
        iSoma = iSoma + iMultiplicador(iContadorFor) * Mid(pConta, iContadorFor + 1, 1)
    Next

    div11 = iSoma Mod 11
    pDigito = 11 - div11
                   
    If Not pDigito < 10 Then
         pDigito = 0
    End If

    DIG_BANCOOB = True

End Function

Public Function DIG_SAFRA(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                             ByRef pMensagem As String) As Boolean
                             
'*    M O D U L O   11  -  DEVE CONSIDERAR AGENCIA
'*     ALGORITIMO E REGRA OBTIDOS DO PROG MARITIMA RER01S160
'*     A CONTA SAFRA[422] TEM 6 DIGITOS + 1 DIG VERIFICADOR
'*     0140 015033-4 AGENCIA+CONTA+DIGITO 10 DIGITOS
'*     3298 765432 = MULTIPLICADOR SAFRA BYTE A BYTE
'*    ------------------------------------------------
'*    (0*3=0) + (1*2=2) + (4*9=36) + (0*8=0) +
'*    (0*7=0) + (1*6=6) + (5*5=25) + (0*4=0) +
'*    (3*3=9) + (3*2=6) = 84 MOD 11 = 7 | 11 - 7 = 4
'*    OU |84 / 11 = 7,63
'*       | 7 * 11 = 77    84 - 77 = 7
'*       |11 -  7 = 4 <<<<<DIGITO DA CONTA

'Contas para testar
'00021702012 - OK
'00190223898 - OK
'00305013505 - OK
'00540029572 - OK
'00830073205 - OK
'01158004106 - OK
'01260030269 - OK
'01470005643 - OK
'01540000483 - OK
'01970005669 - OK
                             
Dim iContadorFor As Integer
Dim iSoma As Integer
Dim iMultiplicador(10) As Integer
Dim iResto As Integer
Dim div11 As Integer
Dim mult11 As Integer
Dim sAgenciaConta As String

    
    '1. Verifica parâmetros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Safra não " & _
                    "preenchido)."
        Exit Function
    End If
    
    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Safra não " & _
                    "numérico - " & pAgencia & ")."
        Exit Function
    End If
        
    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da conta do Banco Safra não " & _
                    "preenchido)."
        Exit Function
    End If
    
    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da conta do Banco Safra não " & _
                    "numérico - " & pConta & ")."
        Exit Function
    End If

    iMultiplicador(0) = 3
    iMultiplicador(1) = 2
    iMultiplicador(2) = 9
    iMultiplicador(3) = 8
    iMultiplicador(4) = 7
    iMultiplicador(5) = 6
    iMultiplicador(6) = 5
    iMultiplicador(7) = 4
    iMultiplicador(8) = 3
    iMultiplicador(9) = 2
                             
    DIG_SAFRA = False
    iSoma = 0
    iResto = 0
    pDigito = ""
    div11 = 0
    mult11 = 0
    sAgenciaConta = Replace(Format$(val(pAgencia), "0000") & Format$(val(pConta), "000000"), "-", "")

    For iContadorFor = 0 To 9
        iSoma = iSoma + iMultiplicador(iContadorFor) * Mid(sAgenciaConta, iContadorFor + 1, 1)
    Next

    div11 = iSoma Mod 11
                   
    If div11 = 1 Then
         pDigito = 0
    Else
        If div11 = 0 Then
        pDigito = 1
    Else
            pDigito = 11 - div11
        End If
    End If

    DIG_SAFRA = True

End Function

'Novo Convênio banco BRB
Public Function DIG_BANCOBRB(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                             ByRef pMensagem As String) As Boolean
    Dim lAuxiliar       As String
    Dim lI              As Integer
    Dim lResto          As String
    Dim lTotal          As Long
    Dim lMultiplicador  As Integer
    Dim sContaSemDigito As String
    
    DIG_BANCOBRB = False
    
    '1. Verifica parâmetros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Bancoob não " & _
                    "preenchido)."
        Exit Function
    End If
    
    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da agência do Banco Bancoob não " & _
                    "numérico - " & pAgencia & ")."
        Exit Function
    End If
        
    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da conta do Banco Bancoob não " & _
                    "preenchido)."
        Exit Function
    End If
    
    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no cálculo de dígito (código da conta do Banco Bancoob não " & _
                    "numérico - " & pConta & ")."
        Exit Function
    End If
    
    ' Validação de conta do Banco BRB funciona da seguinte forma:
    ' Multiplicar cada número pela posição dele, da direita para a esquerda.
    ' A excessão é a posição 10. Nesta posição, multiplicar o número por 2.
    ' Depois somar os produtos e obter o resto da divisão desta soma por 11.
    ' Se o resto for 0 ou 1, o digito é 0, senão, o digito é 11 - resto.
    
    ' Exemplo: Conta 167121301-4. Posição 1 = 4, Posição 2 = 1, Posição 3 = 0, Posição 4 = 3, Posição 5 = 1... E assim por diante.
    ' A primeira posição é o dígito, então descartar (afinal você está calculando o dígito). Começar a partir da segunda posição.
    ' SomaDosProdutos = (2 * 1) + (3 * 0) + (4 * 3) + (5 * 1) + (6 * 2) + (7 * 1) + (8 * 7) + (9 * 6) + (2 * 1) => 150
    ' Resto = SomaDosProdutos Mod 11 => 150 Mod 11 => 7
    ' Digito = 11 - Resto => 11 - 7 => 4
    
    lMultiplicador = 2 ' Sempre começa com 2
    lTotal = 0
    For lI = Len(pConta) To 1 Step -1 ' Loop para cada digito da direita para a esquerda
        If lMultiplicador = 10 Then lMultiplicador = 2
        lAuxiliar = Mid(sContaSemDigito, lI, 1)
        lTotal = lTotal + (val(lAuxiliar) * lMultiplicador)
        lMultiplicador = lMultiplicador + 1
    Next
    
    lResto = lTotal Mod 11
    
    If lResto = 0 Or lResto = 1 Then
        pDigito = 0
    Else
        pDigito = 11 - lResto
    End If
    
    DIG_BANCOBRB = True
End Function
