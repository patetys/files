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
'Fun��o: calcula d�gito do bloqueto (copiado do programa P16C031).

'Par�metro de entrada...pBloqueto...Campo com n�mero do bloqueto.
'Par�metros de sa�da....pDigito.....D�gito calculado.
'                       pMensagem...Mensagem.

    Dim lAuxiliar As String     'Auxiliar.
    Dim lPeso As String    'Peso.
    Dim lPosicao As Integer   'Posi��o.
    Dim lResto As Integer    'Resto.
    Dim lSoma As Integer    'Soma.

    DIG_BLOQUETO = False

    '1. Multiplica d�gito pelo peso e acumula.
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
'Fun��o: calcula d�gito de conta do Banco de Boston.

'Par�metro de entrada...pEntrada....Campo a ser calculado d�gito verificador.
'Par�metros de sa�da....pDigito.....D�gito calculado.
'                       pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_BOSTON = False

    '1. Verifica par�metro de entrada.
    If Not M46V999.gfPreenchido(pEntrada) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco de " & _
                    "Boston n�o preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pEntrada) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco de " & _
                    "Boston n�o num�rico - " & pEntrada & ")."
        Exit Function
    End If

    If Len(pEntrada) <> 7 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco de " & _
                    "Boston com tamanho diferente de 7 posi��es - " & pEntrada & ")."
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

    '4. Divide por 11 e obt�m resto.
    lResto = lTotal Mod 11

    '5. Posiciona d�gito.
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
'Fun��o: calcula d�gito m�dulo 11 - BRADESCO.

'Par�metro de entrada...pEntrada....N�mero a ser calculado d�gito verificador (ag�ncia ou conta).
'Par�metros de sa�da....pDigito.....D�gito calculado.
'                       pMensagem...Mensagem.

    Dim lAuxCalculo As Byte         'Auxiliar de c�lculo.
    Dim lAuxiliar As String     'Auxiliar.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lPeso As Byte   'Peso.
    Dim lPosicao As Byte      'Posi��o.
    Dim lResto As Byte    'Resto.
    Dim lTabPesos As String     'Tabela de pesos.
    Dim lTamanho As Byte      'Tamanho do campo.
    Dim lTotal As Double  'Total.

    DIG_BRADESCO = False

    '1. Verifica par�metro de entrada.
    If Not M46V999.gfPreenchido(pEntrada) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia ou n�mero da conta " & _
                    "corrente do Banco Bradesco n�o preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pEntrada) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia ou n�mero da conta " & _
                    "corrente do Banco Bradesco n�o num�rico - " & pEntrada & ")."
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

    '4. Multiplica resultado por 10, divide por 11 e obt�m resto.
    lTotal = lTotal * 10
    lResto = lTotal Mod 11
    If lResto = 10 Then
        lResto = 0
    End If

    '5. Posiciona d�gito.
    pDigito = lResto

    DIG_BRADESCO = True
End Function
Public Function DIG_CNPJ(ByVal pCNPJ As String, ByRef pDigito As String, ByRef pMensagem As String) _
       As Boolean
'Fun��o: calcula d�gito do CNPJ (c�pia da fun��o QSf_iCGC).

'Par�metro de entrada...pCNPJ.......Campo com CNPJ.
'Par�metros de sa�da....pDigito.....D�gito calculado.
'                       pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lDigCalculado As Integer    'D�gito calculado.
    Dim lDigInformado As String    'D�gito informado.
    Dim lPeso As String    'Peso.

    DIG_CNPJ = False

    '1. Verifica CNPJ.
    pCNPJ = Trim$(pCNPJ)
    Select Case Left$(pCNPJ, 10)
    Case "12.345.678", "11.111.111", "22.222.222", "33.333.333", "44.444.444", _
         "55.555.555", "66.666.666", "77.777.777", "88.888.888", "99.999.999"
        pMensagem = "FGE0054 - CNPJ inv�lido: " & pCNPJ & "."
        Exit Function
    End Select

    '2. Posiciona campos.
    lAuxiliar = Left(pCNPJ, 16)
    lDigInformado = Right$(pCNPJ, 2)

    '3. Calcula primeiro d�gito do CNPJ.
    lPeso = "05040003020900080706000504030200"
    If DIG_MODULO11(lAuxiliar, lPeso, lDigCalculado, pMensagem) = False Then
        Exit Function
    End If
    lAuxiliar = lAuxiliar & Format(lDigCalculado, "0")

    '4. Calcula segundo d�gito do CNPJ.
    lPeso = "0605000403020009080700060504030002"
    If DIG_MODULO11(lAuxiliar, lPeso, lDigCalculado, pMensagem) = False Then
        Exit Function
    End If
    lAuxiliar = lAuxiliar & Format(lDigCalculado, "0")

    '5. Posiciona d�gito do CNPJ.
    pDigito = Right$(lAuxiliar, 2)

    '6. Verifica d�gito.
    If pDigito <> lDigInformado Then
        pMensagem = "FGE0055 - CNPJ com d�gito inv�lido: " & pCNPJ & "."
        Exit Function
    End If

    DIG_CNPJ = True
End Function

Public Function DIG_CodIdeApolice(ByVal pCodIdeApolice As String, ByRef pNumApolice As String, _
                                  ByRef pNumItem As String, ByRef pDigito As String, _
                                  ByRef pMensagem As String) As Boolean
'Fun��o: calcula d�gito do c�digo de identifica��o da ap�lice.

'Par�metro de entrada...pCodIdeApolice...C�digo de identifica��o da ap�lice
'Par�metros de sa�da....pNumApolice......N�mero da ap�lice.
'                       pNumItem.........N�mero do item.
'                       pDigito..........D�gito calculado.
'                       pMensagem........Mensagem.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lCaracter As String    'Caracter.
    Dim lCodSeguradora As String  'C�digo da seguradora.
    Dim lI As Integer             'Utilizado com For...Next.
    Dim lParcela As Integer    'Valor da parcela.
    Dim lPosicao As Byte    'Utilizado com InStr.
    Dim lTotal As Integer    'Total.
    Dim lPeso As Integer    'Peso a ser multiplicado.
    Dim lResto As Integer    'Resto da divis�o do m�dulo 11.

    DIG_CodIdeApolice = False

    '1. Verifica campo de entrada - preenchimento e tamanho.
    If Not M46V999.gfPreenchido(pCodIdeApolice) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo de identifica��o da ap�lice n�o " & _
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
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo de identifica��o da ap�lice com " & _
                    "tamanho diferente de 14 posi��es - " & pCodIdeApolice & ")."
        Exit Function
    End If

    '2. Verifica c�digo de identifica��o da ap�lice - c�digo da seguradora.
    lCodSeguradora = Left$(lAuxiliar, 3)
    If Not IsNumeric(lCodSeguradora) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo de identifica��o da ap�lice com " & _
                    "c�digo da seguradora n�o num�rico - " & pCodIdeApolice & ")."
        Exit Function
    End If
    If val(lCodSeguradora) = 0 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo de identifica��o da ap�lice com " & _
                    "c�digo da seguradora zerado - " & pCodIdeApolice & ")."
        Exit Function
    End If

    '3. Calcula d�gito novo (M�dulo 11).
    'a) Inicializa campo de total.
    lTotal = 0
    lPeso = 6

    'b) Para cada caracter, obt�m valor da parcela (c�digo ASCII), multiplica pelo n�mero de ordem e acumula.

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
            pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo de identifica��o da ap�lice " & _
                        "com caracter (" & lCaracter & ") inv�lido - " & pCodIdeApolice & ")."
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
    Case 0, 1       'Resto igual a zero ou um: d�gito = zero.
        pDigito = 0
    Case Else       'Outro resto: d�gito = diferen�a entre 10 e resto da divis�o.
        pDigito = 11 - (lResto)
    End Select

    If pDigito = Mid$(lAuxiliar, 14, 1) Then
        DIG_CodIdeApolice = True
        Exit Function
    End If

    '4. Calcula d�gito velho.
    'a) Inicializa campo de total e parcela.
    lTotal = 0
    lParcela = 0
    'b) Para cada caracter, obt�m valor da parcela (c�digo ASCII), multiplica pelo n�mero de ordem e
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
            pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo de identifica��o da ap�lice " & _
                        "com caracter (" & lCaracter & ") inv�lido - " & pCodIdeApolice & ")."
            Exit Function
        End Select
        lTotal = lTotal + (lParcela * lI)
    Next lI
    'c) Se valor acumulado > 10, soma caracter a caracter do valor acumulado at� chegar num valor menor
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
    'd) Posiciona valor do d�gito.
    pDigito = lTotal

    '5. Obt�m n�mero de ap�lice e n�mero do item (apenas para seguradora 641 - Yasuda).
    'a) Inicializa campos.
    pNumApolice = "0000000000"
    pNumItem = "0000"
    'b) Se seguradora diferente de 641, encerra processo.
    If val(lCodSeguradora) <> 641 Then
        DIG_CodIdeApolice = True
        Exit Function
    End If
    'c) Posiciona n�mero da ap�lice e n�mero do item.
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
'Fun��o: calcula d�gito do CPF (c�pia da fun��o QSf-iCPF).

'Par�metro de entrada...pCPF........Campo com CPF no formato num�rico.
'Par�metros de sa�da....pDigito.....D�gito calculado.
'                       pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lDigCalculado As Integer    'D�gito calculado.
    Dim lDigInformado As String    'D�gito informado.
    Dim lPeso As String    'Peso.

    DIG_CPF = False

    '1. Verifica CPF.
    pCPF = Trim$(pCPF)
    Select Case Left$(pCPF, 11)
    Case "123.456.789", "111.111.111", "222.222.222", "333.333.333", "444.444.444", _
         "555.555.555", "666.666.666", "777.777.777", "888.888.888", "999.999.999"
        pMensagem = "FGE0052 - CPF inv�lido: " & pCPF & "."
        Exit Function
    End Select

    '2. Posiciona campos auxiliar.
    lAuxiliar = Left(pCPF, 12)
    lDigInformado = Right$(pCPF, 2)

    '3. Calcula primeiro d�gito do CPF.
    lPeso = "100908000706050004030200"
    If DIG_MODULO11(lAuxiliar, lPeso, lDigCalculado, pMensagem) = False Then
        Exit Function
    End If
    lAuxiliar = lAuxiliar & Format(lDigCalculado, "0")

    '4. Calcula segundo d�gito do CPF.
    lPeso = "11100900080706000504030002"
    If DIG_MODULO11(lAuxiliar, lPeso, lDigCalculado, pMensagem) = False Then
        Exit Function
    End If
    lAuxiliar = lAuxiliar & Format(lDigCalculado, "0")

    '5. Posiciona d�gito do CPF.
    pDigito = Right$(lAuxiliar, 2)

    '6. Verifica d�gito.
    If pDigito <> lDigInformado Then
        pMensagem = "FGE0053 - CPF com d�gito inv�lido: " & pCPF & "."
        Exit Function
    End If

    DIG_CPF = True
End Function
Public Function DIG_ITAU(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                         ByRef pMensagem As String) As Boolean
'Fun��o: calcula d�gito m�dulo 10 - Ita�.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta corrente.
'Par�metros de sa�da.....pDigito.....D�gito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String     'Auxiliar.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lResultado As String      'Resultado.
    Dim lTabPesos As String     'Tabela de pesos.
    Dim lTotal As Double  'Total.

    DIG_ITAU = False

    '1. Verifica se par�metros de entrada est�o preenchidos e v�lidos.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Ita� n�o " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Ita� n�o " & _
                    "num�rico - " & pAgencia & ")."
        Exit Function
    End If

    If Len(pAgencia) <> 4 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Ita� com " & _
                    "tamanho diferente de 4 posi��es - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco Ita� n�o " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco Ita� n�o " & _
                    "num�rico - " & pConta & ")."
        Exit Function
    End If

    If Len(pConta) <> 5 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco Ita� com " & _
                    "tamanho diferente de 5 posi��es - " & pConta & ")."
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

    '4. Divide por 10 e obt�m resto.
    lResto = lTotal Mod 10

    '5. Posiciona d�gito.
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
'Fun��o: calcula d�gito m�dulo 10 - HSBC.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta corrente.
'Par�metros de sa�da.....pDigito.....D�gito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String     'Auxiliar.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lResultado As String      'Resultado.
    Dim lTabPesos As String     'Tabela de pesos.
    Dim lTotal As Double  'Total.

    DIG_HSBC = False

    '1. Verifica se par�metros de entrada est�o preenchidos e v�lidos.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco HSBC n�o " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco HSBC n�o " & _
                    "num�rico - " & pAgencia & ")."
        Exit Function
    End If

    If Len(pAgencia) <> 4 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco HSBC com " & _
                    "tamanho diferente de 4 posi��es - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco HSBC n�o " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco HSBC n�o " & _
                    "num�rico - " & pConta & ")."
        Exit Function
    End If

    If Len(pConta) <> 10 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco HSBC com " & _
                    "tamanho diferente de 10 posi��es - " & pConta & ")."
        Exit Function
    End If

    'If Left(pConta, 4) <> pAgencia Then
    '    pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco HSBC n�o corresponde a ag�ncia informada)."
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

    '4. Divide por 10 e obt�m resto.
    lResto = lTotal Mod 11

    '5. Posiciona d�gito.
    pDigito = lResto
    Select Case lResto
    Case 0, 10
        pDigito = 0
    End Select

    DIG_HSBC = True
End Function

Public Function DIG_MODULO11(ByVal pCampo As String, ByVal pPesos As String, _
                             ByRef pDigito As Integer, ByRef pMensagem As String) As Boolean
'Fun��o: calcula d�gito utilizando peso fornecido e algor�timo para m�dulo 11.

'Par�metros de entrada...pCampo......Campo a calcular d�gito.
'                        pPesos......Pesos para efetuar c�lculo.
'Par�metros de sa�da.....pDigito.....D�gito calculado.
'                        pMensagem...Mensagem.

'Observa��o: para cada d�gito de pCampo existir�o dois em pPesos.
'Exemplo: se peso � 1, 2, 3, 4, informar 01020304.

'Exemplos da fun��o: DIG_MODULO11("098031208","100908070605040302") = 6.
'                    DIG_MODULO11("0980312086","01100908070605040302") = 7.
'Obs: aceita caracteres alfanum�ricos (A-Z) em pCampo, conforme algoritmo de d�gito verificador do
'     CNPJ alfanum�rico da Receita Federal (valor do caracter = Asc(caracter) - 48).

'Se a posi��o do peso a ser aplicado � igual a zero, o sistema n�o considera a posi��o do campo a ser
'calculado d�gito correspondente (mesmo que n�o seja num�rico).

    Dim lPosicao As Integer    'Posi��o.
    Dim lResto As Integer    'Resto.
    Dim lSoma As Integer    'Soma.
    Dim lTamCampo As Integer    'Tamanho do campo de entrada.
    Dim lTamPesos As Integer    'Tamanho do campo de pesos.
    Dim lCaracter As String    'Caracter sendo processado (suporta CNPJ alfanum�rico).

    DIG_MODULO11 = False

    '1. Verifica se campos de entrada s�o num�ricos.
    If Not M46V999.gfPreenchido(pCampo) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (m�dulo 11 - campo a ser calculado d�gito " & _
                    "n�o preenchido)."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pPesos) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (m�dulo 11 - campo de pesos n�o preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pPesos) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (m�dulo 11 - campo de pesos n�o num�rico)."
        Exit Function
    End If

    pCampo = Trim$(pCampo)
    pPesos = Trim$(pPesos)
    lTamCampo = Len(pCampo)
    lTamPesos = Len(pPesos)

    If lTamPesos < (lTamCampo * 2) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (m�dulo 11 - campo de pesos (" & pPesos & _
                    ") incompat�vel com campo a ser calculado d�gito (" & pCampo & "))."
        Exit Function
    End If

    '2. Multiplica d�gito pelo peso e acumula.
    lSoma = 0
    For lPosicao = 1 To Len(pCampo)
        lCaracter = UCase(Mid(pCampo, lPosicao, 1))
        If val(Mid(pPesos, 1 + ((lPosicao - 1) * 2), 2)) <> 0 Then
            If Not ((lCaracter >= "0" And lCaracter <= "9") Or (lCaracter >= "A" And lCaracter <= "Z")) Then
                pMensagem = "FTE0047 - Erro no c�lculo de d�gito (m�dulo 11 - campo a ser calculado " & _
                            "d�gito n�o num�rico/alfanum�rico v�lido - " & pCampo & ")."
                Exit Function
            End If
        End If
        'Valor do caracter = Asc(caracter) - 48 (id�ntico ao valor num�rico para '0'-'9'; para 'A'-'Z'
        'resulta em 17-42, conforme algoritmo de d�gito verificador do CNPJ alfanum�rico).
        lSoma = lSoma + ((Asc(lCaracter) - 48) * val(Mid(pPesos, 1 + ((lPosicao - 1) * 2), 2)))
    Next lPosicao

    '3. Retorno da fun��o
    lResto = lSoma Mod 11
    Select Case lResto
    Case 0, 1       'Resto igual a zero ou um: d�gito = zero.
        pDigito = 0
    Case Else       'Outro resto: d�gito = diferen�a entre 10 e resto da divis�o.
        pDigito = 11 - (lResto)
    End Select

    DIG_MODULO11 = True
End Function
Public Function DIG_SEGSUL(ByVal pCampo As String, ByVal pPesos As String, _
                           ByRef pDigito As Integer, ByRef pMensagem As String) As Boolean
'Fun��o: calcula d�gito utilizando peso fornecido e algor�timo para Segsul.

'Par�metros de entrada...pCampo......Campo a calcular d�gito.
'                        pPesos......Pesos para efetuar c�lculo.
'Par�metros de sa�da.....pDigito.....D�gito calculado.
'                        pMensagem...Mensagem.

'Observa��o: para cada d�gito de pCampo existir�o dois em pPesos.
'Exemplo: se peso � 1, 2, 3, 4, informar 01020304.

'Exemplos da fun��o: DIG_SEGSUL("098031208","100908070605040302") = 6.
'                    DIG_SEGSUL("0980312086","01100908070605040302") = 7.

    Dim lAcum As Integer    'Acumulador.
    Dim lPosicao As Integer    'Posi��o.
    Dim lResto As Integer    'Resto.
    Dim lSoma As Integer    'Soma.
    Dim lTamCampo As Integer    'Tamanho do campo de entrada.
    Dim lTamPesos As Integer    'Tamanho do campo de pesos.

    DIG_SEGSUL = False

    '1. Verifica se campos de entrada s�o num�ricos.
    If Not M46V999.gfPreenchido(pCampo) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo do Segsul n�o preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pCampo) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo do Segsul n�o num�rico - " & pCampo & _
                    ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pPesos) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (Segsul - campo de pesos n�o preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pPesos) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (Segsul - campo de pesos n�o num�rico - " & _
                    pPesos & ")."
        Exit Function
    End If

    pCampo = Trim$(pCampo)
    pPesos = Trim$(pPesos)
    lTamCampo = Len(pCampo)
    lTamPesos = Len(pPesos)

    '2. Multiplica d�gito pelo peso e acumula.
    lSoma = 0
    lAcum = 0
    For lPosicao = 1 To Len(pCampo)
        lSoma = Format((val(Mid(pCampo, lPosicao, 1)) * val(Mid(pPesos, lPosicao, 1))), "00")
        lAcum = lAcum + (val(Mid(lSoma, 1, 1)) + val(Mid(lSoma, 2, 1)))
    Next lPosicao

    '3. Retorno da fun��o
    lResto = lAcum Mod 11
    Select Case lResto
    Case 0, 1       'Resto igual a zero ou um: d�gito = zero.
        pDigito = 0
    Case Else       'Outro resto: d�gito = diferen�a entre 10 e resto da divis�o.
        pDigito = 11 - (lResto)
    End Select

    DIG_SEGSUL = True
End Function
Public Function DIG_SEGSUL1(ByVal pCampo As String, ByVal pPesos As String, _
                            ByRef pDigito As String, ByRef pMensagem As String) As Boolean
'Fun��o: calcula d�gito utilizando peso fornecido e algor�timo para Segsul no P28V200.

'Par�metros de entrada...pCampo......Campo a calcular d�gito.
'                        pPesos......Pesos para efetuar c�lculo.
'Par�metros de sa�da.....pDigito.....D�gito calculado.
'                        pMensagem...Mensagem.

'Observa��o: para cada d�gito de pCampo existir�o dois em pPesos.
'Exemplo: se peso � 1, 2, 3, 4, informar 01020304.

'Exemplos da fun��o: DIG_SEGSUL1("098031208","100908070605040302") = 6.
'                    DIG_SEGSUL1("0980312086","01100908070605040302") = 7.

    Dim lAcum As Integer    'Acumulador.
    Dim lPosicao As Integer    'Posi��o.
    Dim lResto As Integer    'Resto.
    Dim lSoma As Integer    'Soma.
    Dim lTamCampo As Integer    'Tamanho do campo de entrada.
    Dim lTamPesos As Integer    'Tamanho do campo de pesos.

    DIG_SEGSUL1 = False

    '1. Verifica se campos de entrada s�o num�ricos.
    If Not M46V999.gfPreenchido(pCampo) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo do Segsul n�o preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pCampo) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo do Segsul n�o num�rico - " & pCampo & _
                    ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pPesos) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (Segsul - campo de pesos n�o preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pPesos) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (Segsul - campo de pesos n�o num�rico - " & _
                    pPesos & ")."
        Exit Function
    End If

    pCampo = Trim$(pCampo)
    pPesos = Trim$(pPesos)
    lTamCampo = Len(pCampo)
    lTamPesos = Len(pPesos)

    '2. Multiplica d�gito pelo peso e acumula.
    lSoma = 0
    lAcum = 0
    For lPosicao = 1 To Len(pCampo)
        lSoma = Format((val(Mid(pCampo, lPosicao, 1)) * val(Mid(pPesos, lPosicao, 1))), "00")
        lAcum = lAcum + (val(Mid(lSoma, 1, 1)) + val(Mid(lSoma, 2, 1)))
    Next lPosicao

    '3. Retorno da fun��o
    lResto = lAcum Mod 11
    Select Case lResto
    Case 0          'Resto igual a zero : digito = "A".
        pDigito = "A"
    Case 1          'Resto igual a um: d�gito = zero.
        pDigito = 0
    Case Else       'Outro resto: d�gito = diferen�a entre 10 e resto da divis�o.
        pDigito = val(11 - (lResto))
    End Select

    DIG_SEGSUL1 = True
End Function
Public Function DIG_SUDAMERIS(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                              ByRef pMensagem As String) As Boolean
'Fun��o: calcula d�gito m�dulo 11 - conta corrente.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metros de sa�da.....pDigito.....D�gito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar...de 01 a 03 (03 posi��es)...C�digo do banco.
'           de 04 a 06 (03 posi��es)...C�digo da ag�ncia.
'           de 07 a 17 (11 posi��es)...Valor fixo.
'           de 18 a 26 (09 posi��es)...N�mero da conta.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_SUDAMERIS = False

    '1. Verifica par�metros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Sudameris n�o " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Sudameris n�o " & _
                    "num�rico - " & pAgencia & ")."
        Exit Function
    End If

    If Len(pAgencia) <> 3 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Sudameris com " & _
                    "tamanho diferente de 3 posi��es - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco " & _
                    "Sudameris n�o preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco " & _
                    "Sudameris n�o num�rico - " & pConta & ")."
        Exit Function
    End If

    If Len(pConta) <> 9 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco " & _
                    "Sudameris com tamanho diferente de 9 posi��es - " & pConta & ")."
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

    '4. Divide por 11 e obt�m resto.
    lResto = lTotal Mod 11

    '5. Posiciona d�gito.
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
'Fun��o: calcula d�gito m�dulo 11 - conta corrente.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metros de sa�da.....pDigito.....D�gito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar...de 01 a 03 (03 posi��es)...C�digo do banco.
'           de 04 a 07 (04 posi��es)...C�digo da ag�ncia.
'           de 08 a 14 (07 posi��es)...N�mero da conta.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_SUDAMERIS_NOVO = False

    '1. Verifica par�metros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Sudameris n�o " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Sudameris n�o " & _
                    "num�rico - " & pAgencia & ")."
        Exit Function
    End If

    If Len(pAgencia) <> 4 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Sudameris com " & _
                    "tamanho diferente de 3 posi��es - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco " & _
                    "Sudameris n�o preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco " & _
                    "Sudameris n�o num�rico - " & pConta & ")."
        Exit Function
    End If

    If Len(pConta) <> 7 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco " & _
                    "Sudameris com tamanho diferente de 7 posi��es - " & pConta & ")."
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

    '4. Divide por 11 e obt�m resto.
    lResto = lTotal Mod 11

    '5. Posiciona d�gito.
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
'Fun��o: calcula d�gito m�dulo 11 - conta corrente.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metros de sa�da.....pDigito.....D�gito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar...de 01 a 03 (03 posi��es)...C�digo do banco.
'           de 04 a 07 (04 posi��es)...C�digo da ag�ncia.
'           de 08 a 14 (07 posi��es)...N�mero da conta.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_REAL = False

    '1. Verifica par�metros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Real n�o " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Real n�o " & _
                    "num�rico - " & pAgencia & ")."
        Exit Function
    End If

    If Len(pAgencia) <> 4 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Real com " & _
                    "tamanho diferente de 3 posi��es - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco " & _
                    "Real n�o preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco " & _
                    "Real n�o num�rico - " & pConta & ")."
        Exit Function
    End If

    If Len(pConta) <> 7 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Banco " & _
                    "Real com tamanho diferente de 7 posi��es - " & pConta & ")."
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

    '4. Divide por 11 e obt�m resto.
    lResto = lTotal Mod 11

    '5. Posiciona d�gito.
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
'Fun��o: calcula d�gito m�dulo 11 - conta corrente.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metros de sa�da.....pDigito.....D�gito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar...de 01 a 03 (03 posi��es)...C�digo do banco.
'           de 04 a 07 (04 posi��es)...C�digo da ag�ncia.
'           de 08 a 14 (07 posi��es)...N�mero da conta.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As String    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_BANCOBRASIL = False

    '1. Verifica par�metros de entrada.
    If Not M46V999.gfPreenchido(pInfo) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia ou conta do Banco Real n�o " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pInfo) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia ou conta do Banco Real n�o " & _
                    "num�rico - " & pInfo & ")."
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

    '4. Divide por 11 e obt�m resto.
    lResto = lTotal Mod 11

    '5. Posiciona d�gito.
    If lResto = 10 Then lResto = "X"

    pDigito = lResto

    DIG_BANCOBRASIL = True
End Function

Public Function DIG_UNIBANCO(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                             ByRef pMensagem As String) As Boolean
'Fun��o: calcula d�gito m�dulo 10 - Unibanco.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metros de sa�da.....pDigito.....D�gito calculado.
'                        pMensagem...Mensagem.

'Exemplo para Banco Unibanco:
'
'3   9   8   1   0   3   5  1  4    ===>>> Ag�ncia 3 posi��es e conta 6 posi��es.
'2   1   2   1   2   1   2  1  2    ===>>> Seq��ncia de 2 e 1.
'6   9  16   1   0   3  10  1  8    ===>>> Total da multiplica��o.
'
'Para consistir todos os n�meros alteradas a seq��ncia de 2 e 1.
'
'0  3   9   8   0  1   0   3   5  1  4  ===>>> Ag�ncia 4 posi��es e conta 7 posi��es.
'   2   1   2      1   2   1   2  1  2  ===>>> Seq��ncia de 2 e 1.
'   6   9  16      1   0   3  10  1  8  ===>>> Total da multiplica��o.

    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As Byte    'Resto.
    Dim lResultado As String  'Resultado.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_UNIBANCO = False

    '1. Verifica se par�metros de entrada est�o preenchidos e v�lidos.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Unibanco n�o preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Unibanco n�o num�rico " & _
                    "- " & pAgencia & ")."
        Exit Function
    End If

    If Len(pAgencia) <> 4 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Unibanco com tamanho " & _
                    "diferente de 4 posi��es - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Unibanco n�o " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Unibanco n�o " & _
                    "num�rico - " & pConta & ")."
        Exit Function
    End If

    If Len(pConta) <> 6 Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da conta corrente do Unibanco com " & _
                    "tamanho diferente de 6 posi��es - " & pConta & ")."
        Exit Function
    End If

    '2. Posiciona controles.
    lTabPesos = "1212121212121"
    lTotal = 0

    '3. Aplica peso.
    'a) Para ag�ncia.
    For lI = 1 To 4
        lResultado = Format$(CByte(Mid(pAgencia, lI, 1)) * CByte(Mid(lTabPesos, lI, 1)), "00")
        lTotal = lTotal + CByte(Mid(lResultado, 1, 1)) + CByte(Mid(lResultado, 2, 1))
    Next lI
    'b) Para conta.
    For lI = 1 To 6
        lResultado = Format$(CByte(Mid(pConta, lI, 1)) * CByte(Mid(lTabPesos, lI, 1)), "00")
        lTotal = lTotal + CByte(Mid(lResultado, 1, 1)) + CByte(Mid(lResultado, 2, 1))
    Next lI

    '4. Divide por 10 e obt�m resto.
    lResto = lTotal Mod 10

    '5. Posiciona d�gito.
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
'Fun��o: calcula d�gito m�dulo 11 - Yasuda.

'Par�metro de entrada...pEntrada.....N�mero a ser calculado d�gito verificador.
'Par�metros de sa�da....pDigito.....D�gito calculado.
'                       pMensagem...Mensagem.

    Dim lAuxCalculo As Byte     'Auxiliar de c�lculo.
    Dim lAuxiliar As String    'Auxiliar.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lPeso As Byte    'Peso.
    Dim lPosicao As Byte  'Posi��o.
    Dim lResto As Byte    'Resto.
    Dim lResultado As String  'Resultado.
    Dim lTamanho As Byte  'Tamanho do campo.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_YASUDA = False

    '1. Verifica se par�metros de entrada est�o preenchidos e v�lidos.
    If Not M46V999.gfPreenchido(pEntrada) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da Yasuda n�o preenchido)."
        Exit Function
    End If
    If Not IsNumeric(pEntrada) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (n�mero da Yasuda n�o n�o num�rico - " & _
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

    '4. Divide resultado por 11 e obt�m resto.
    lResto = lTotal Mod 11

    '5. Posiciona d�gito.
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
Public Function DIG_BANESE(ByVal pAgencia As String, pConta As String, ByRef pDigito As String, _
                           ByRef pMensagem As String) As Boolean
'Fun��o: calcula d�gito m�dulo 11 - conta corrente.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metros de sa�da.....pDigito.....D�gito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar...de 01 a 03 (03 posi��es)...C�digo do banco.
'           de 04 a 07 (04 posi��es)...C�digo da ag�ncia.
'           de 08 a 14 (07 posi��es)...N�mero da conta.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As String    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.
    Dim lAgenciaC As String

    DIG_BANESE = False

    '1. Verifica par�metros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Banese n�o " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Banese n�o " & _
                    "num�rico - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da conta do Banco Banese n�o " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da conta do Banco Banese n�o " & _
                    "num�rico - " & pConta & ")."
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

    '4. Divide por 11 e obt�m resto.
    lResto = lTotal Mod 11

    '5. Posiciona d�gito.
    lResto = 11 - lResto
    If lResto = "10" Then lResto = "0"
    pDigito = lResto

    DIG_BANESE = True
End Function

Public Function DIG_SANTANDER(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                              ByRef pMensagem As String) As Boolean
'Fun��o: calcula d�gito conta corrente.

'Par�metros de entrada...pAgencia....C�digo da ag�ncia.
'                        pConta......N�mero da conta.
'Par�metros de sa�da.....pDigito.....D�gito calculado.
'                        pMensagem...Mensagem.

    Dim lAuxiliar As String    'Auxiliar...de 01 a 03 (03 posi��es)...C�digo do banco.
'           de 04 a 07 (04 posi��es)...C�digo da ag�ncia.
'           de 08 a 14 (07 posi��es)...N�mero da conta.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lResto As String    'Resto.
    Dim lTabPesos As String    'Tabela de pesos.
    Dim lTotal As Double    'Total.

    DIG_SANTANDER = False

    '1. Verifica par�metros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Santander n�o " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Santander n�o " & _
                    "num�rico - " & pAgencia & ")."
        Exit Function
    End If

    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da conta do Banco Santander n�o " & _
                    "preenchido)."
        Exit Function
    End If

    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da conta do Banco Santander n�o " & _
                    "num�rico - " & pConta & ")."
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

    '4. Divide por 11 e obt�m resto.
    lResto = 10 - Right(lTotal, 1)

    '5. Posiciona d�gito.
    If lResto = 10 Then lResto = "0"

    pDigito = lResto

    DIG_SANTANDER = True
End Function

'CyberTools - Ari - Task 3795:Implementar fun��o da valida��o do CNH, Renavam, Safra e Bancoob - Setembro\2014
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

'BANCOOB - Desconsiderado um zero � esquerda '*TEM 9 DIGITOS + 1 DIG VERIFICADOR
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

    '1. Verifica par�metros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Bancoob n�o " & _
                    "preenchido)."
        Exit Function
    End If
    
    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Bancoob n�o " & _
                    "num�rico - " & pAgencia & ")."
        Exit Function
    End If
        
    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da conta do Banco Bancoob n�o " & _
                    "preenchido)."
        Exit Function
    End If
    
    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da conta do Banco Bancoob n�o " & _
                    "num�rico - " & pConta & ")."
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

    
    '1. Verifica par�metros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Safra n�o " & _
                    "preenchido)."
        Exit Function
    End If
    
    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Safra n�o " & _
                    "num�rico - " & pAgencia & ")."
        Exit Function
    End If
        
    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da conta do Banco Safra n�o " & _
                    "preenchido)."
        Exit Function
    End If
    
    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da conta do Banco Safra n�o " & _
                    "num�rico - " & pConta & ")."
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

'Novo Conv�nio banco BRB
Public Function DIG_BANCOBRB(ByVal pAgencia As String, ByVal pConta As String, ByRef pDigito As String, _
                             ByRef pMensagem As String) As Boolean
    Dim lAuxiliar       As String
    Dim lI              As Integer
    Dim lResto          As String
    Dim lTotal          As Long
    Dim lMultiplicador  As Integer
    Dim sContaSemDigito As String
    
    DIG_BANCOBRB = False
    
    '1. Verifica par�metros de entrada.
    If Not M46V999.gfPreenchido(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Bancoob n�o " & _
                    "preenchido)."
        Exit Function
    End If
    
    If Not IsNumeric(pAgencia) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da ag�ncia do Banco Bancoob n�o " & _
                    "num�rico - " & pAgencia & ")."
        Exit Function
    End If
        
    If Not M46V999.gfPreenchido(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da conta do Banco Bancoob n�o " & _
                    "preenchido)."
        Exit Function
    End If
    
    If Not IsNumeric(pConta) Then
        pMensagem = "FTE0047 - Erro no c�lculo de d�gito (c�digo da conta do Banco Bancoob n�o " & _
                    "num�rico - " & pConta & ")."
        Exit Function
    End If
    
    ' Valida��o de conta do Banco BRB funciona da seguinte forma:
    ' Multiplicar cada n�mero pela posi��o dele, da direita para a esquerda.
    ' A excess�o � a posi��o 10. Nesta posi��o, multiplicar o n�mero por 2.
    ' Depois somar os produtos e obter o resto da divis�o desta soma por 11.
    ' Se o resto for 0 ou 1, o digito � 0, sen�o, o digito � 11 - resto.
    
    ' Exemplo: Conta 167121301-4. Posi��o 1 = 4, Posi��o 2 = 1, Posi��o 3 = 0, Posi��o 4 = 3, Posi��o 5 = 1... E assim por diante.
    ' A primeira posi��o � o d�gito, ent�o descartar (afinal voc� est� calculando o d�gito). Come�ar a partir da segunda posi��o.
    ' SomaDosProdutos = (2 * 1) + (3 * 0) + (4 * 3) + (5 * 1) + (6 * 2) + (7 * 1) + (8 * 7) + (9 * 6) + (2 * 1) => 150
    ' Resto = SomaDosProdutos Mod 11 => 150 Mod 11 => 7
    ' Digito = 11 - Resto => 11 - 7 => 4
    
    lMultiplicador = 2 ' Sempre come�a com 2
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
