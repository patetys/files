Attribute VB_Name = "basM28V999"
'MÓDULO COM FUNÇÕES E PROCEDURES COMUNS DO SYAS.

'AO ALTERAR ESTE MÓDULO, COMPILAR E CRIAR EXECUTÁVEL DE TODOS OS MÓDULOS QUE O REFERENCIAM.

Option Explicit

Public gPodeEmitir          As Boolean
'[FRH.EX02]->
Public gModNome          As String
Public gModTipo          As String
Public gModBase          As String
Public gModNomeTab       As String
Public gModTabela        As String

Public gMCmpObg()        As String
Public gMCmpVal()        As String
Public gMCmpUnida()      As String

Public gMFlxObg()        As String
Public gMFlxVal()        As String
Public gMFlxUnida()      As String

Public gMPosObg()        As Long
Public gMPosVal()        As Long

Public gMTipUnida()      As String
Public gMTipObg()        As String
Public gMTipVal()        As String

Public mTipoDadoChave    As Integer
Public rsTbm_Alvo        As ADODB.Recordset
Public lRst_temp         As ADODB.Recordset

Public gModSQL_Incluir   As String
Public gModSQL_Excluir   As String
Public gModSQL_Alterar   As String
Public gModSQL_Carregar  As String

Public bdBase            As ADODB.Connection

'Adicionais Específicos na consulta SQL Genérica
'Public gSQLSelect_01     As Long
Public gSQLSelCarga      As String
'<-[FRH.EX02]

Public gTipEmissao          As e_TipEmissao     'Tipo de emissao da apólice
Public Enum e_TipEmissao
    e_TipEmissao_SegNovo = 100
    e_TipEmissao_RenCong = 101
    e_TipEmissao_RenYas = 200
    e_TipEmissao_EndAlt = 300
    e_TipEmissao_EndCanc = 301
    e_TipEmissao_EndInc = 302
    e_TipEmissao_EndExc = 303
End Enum
Private mQuaDias            As Integer          'Quantidade de dias.
Private mEndMinimo          As Boolean          'Indicador de endosso mínimo.
Private mRestituicao        As Boolean          'Indicador de restituição.
Private mNumProposta        As Long             'Número da proposta para DT1101D.
Private bdSYASEMIS          As ADODB.Connection
Public gNumSusep            As String
'Public gCod_Corretor        As Long             'Cod do Corretor.
Public gNumPI               As String
Public gProComLimite        As Double           'Produto: comissão básica.
Public gProComBasica        As Double           'Produto: comissão básica.
Public gProCusApolice       As Double           'Produto: custo de apólice.
Public gProJurIOF           As Single           'Produto: juros de IOF.
Public gProJurMinAnt01      As Single           'Produto: limite de juros.
Public gProJurMinAnt02      As Single           'Produto: limite de juros.
Public gProJurMinAnt03      As Single           'Produto: limite de juros.
Public gProJurMinAnt04      As Single           'Produto: limite de juros.
Public gProJurMinAnt05      As Single           'Produto: limite de juros.
Public gProJurMinAnt06      As Single           'Produto: limite de juros.
Public gProJurMinAnt07      As Single           'Produto: limite de juros.
Public gProJurMinAnt08      As Single           'Produto: limite de juros.
Public gProJurMinAnt09      As Single           'Produto: limite de juros.
Public gProJurMinAnt10      As Single           'Produto: limite de juros.
Public gProJurMinAnt11      As Single           'Produto: limite de juros.
Public gProJurMinAnt12      As Single           'Produto: limite de juros.
Public gProJurMinPos01      As Single           'Produto: limite de juros.
Public gProJurMinPos02      As Single           'Produto: limite de juros.
Public gProJurMinPos03      As Single           'Produto: limite de juros.
Public gProJurMinPos04      As Single           'Produto: limite de juros.
Public gProJurMinPos05      As Single           'Produto: limite de juros.
Public gProJurMinPos06      As Single           'Produto: limite de juros.
Public gProJurMinPos07      As Single           'Produto: limite de juros.
Public gProJurMinPos08      As Single           'Produto: limite de juros.
Public gProJurMinPos09      As Single           'Produto: limite de juros.
Public gProJurMinPos10      As Single           'Produto: limite de juros.
Public gProJurMinPos11      As Single           'Produto: limite de juros.
Public gProJurMinPos12      As Single           'Produto: limite de juros.
Public gProLimComissao      As Double           'Produto: limite da comissão.
Public gProLimMarComercial  As Single           'Produto: limite da margem comercial.
Public gProNome             As String           'Produto: nome.
Public gProPreMinParcela    As Double           'Produto: prêmio mínimo por parcela.
Public gProPreMinEndosso    As Double           'Produto: prêmio mínimo por parcela.
Public gProQuaMaxParcelas   As Byte             'Produto: quantidade máxima de parcelas.
Public gProQuaMaxParcelasAnt As Byte            'Produto: quantidade máxima de parcelas. (Anterior)
Public gProTipProtocolo     As Byte             'Produto: tipo do protocolo.
Public gProVigInicio        As Long
Public gProVigTermino       As Long
Public gCodProduto          As Long
Public gCodProtocolo        As Long
Public gCodSequencia        As Long
Public gNovaTransmissao     As Boolean

Public Function gfConverterADODB(Valor As String) As TipoDados

    Select Case CInt(Valor)
        'Caso CARACTER:
        Case adChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar, adWChar
            gfConverterADODB = tpoString
            
        'Caso NUMÉRICO:
        Case adDecimal, adSingle, adNumeric, adDouble, adInteger, adSmallInt, adBigInt, adTinyInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedBigInt, adUnsignedTinyInt
            gfConverterADODB = TpoNumerico
            
        Case Else
            gfConverterADODB = tpoString

    End Select

End Function

Public Function mfFormatarCampos(ByRef cmp01 As Variant, ByRef cmp02 As Variant)
   
Dim lCt_cmp As Long
Dim lchar As Variant

Dim lSinal_STR As Boolean

    Select Case VarType(cmp01)
        Case vbString
            lSinal_STR = False
            
            '[FRH.EX02]->Confere se o campo "cmp01" possui alguma letra
            For lCt_cmp = 1 To Len(cmp01)
                lchar = Asc(Mid(cmp01, lCt_cmp, 1))
                If (lchar >= 97 And lchar <= 122) Or (lchar >= 65 And lchar <= 90) Then
                    lSinal_STR = True
                    Exit For
                End If
            Next lCt_cmp
            
            If lSinal_STR Then
                cmp01 = CStr(Trim(UCase(cmp01)))
            Else
                cmp01 = CStr(Val(cmp01))
            End If
            
        Case vbLong, vbInteger: cmp01 = CStr(CDbl(Val(cmp01)))
        Case vbDouble: cmp01 = CStr(Val(cmp01))
    End Select
    
    Select Case VarType(cmp02)
        Case vbString
            lSinal_STR = False
        
            '[FRH.EX02]->Confere se o campo "cmp02" possui alguma letra
            For lCt_cmp = 1 To Len(cmp02)
                lchar = Asc(Mid(cmp02, lCt_cmp, 1))
                If (lchar >= 97 And lchar <= 122) Or (lchar >= 65 And lchar <= 90) Then
                    lSinal_STR = True
                    Exit For
                Else
                    lSinal_STR = False
                End If
            Next lCt_cmp
            
            If lSinal_STR Then
                cmp02 = CStr(Trim(UCase(cmp02)))
            Else
                cmp02 = CStr(Val(cmp02))
            End If
            
        Case vbLong, vbInteger: cmp02 = CStr(CDbl(Val(cmp02)))
        Case vbDouble: cmp02 = CStr(Val(cmp02))
    End Select
    
End Function

Public Function gfDifDatas(ByVal pInicio As Long, ByVal pTermino As Long) As Single
    'Função #014: calcula diferença entre datas (retorno em dias).

    'Parâmetros de entrada...pInicio....data de início.
    '                        pTermino...data de término.

    Dim lAuxIniDias         As String   'Auxiliar: valores para tabela para cálculo da diferença de dias.
    Dim lAuxTerDias         As String   'Auxiliar: valores para tabela para cálculo da diferença de dias.
    Dim lDia                As Single   'Índice auxiliar: dia.
    Dim lI                  As Integer  'Índice.
    Dim lIniDia             As Single   'Dia do início de vigência.
    Dim lIniMes             As Single   'Mês do início de vigência.
    Dim lMes                As Single   'Índice auxiliar: mês.
    Dim lPrazoLongo         As Integer  'Quantidade de dias para prazo longo.
    Dim lQuaIniDias         As Single   'Quantidade de dias (início de vigência).
    Dim lQuaTerDias         As Single   'Quantidade de dias (término de vigência).
    Dim lTabIniDias(31, 12) As String   'Tabela para cálculo de dias (data de início).
    Dim lTabTerDias(31, 12) As String   'Tabela para cálculo de dias (data de término).
    Dim lTerDia             As Single   'Dia do término de vigência.
    Dim lTerMes             As Single   'Mês do término de vigência.


    If Right(pInicio, 4) = "0229" Then
        pInicio = Left(pInicio, 4) & "0228"
    End If
    If Right(pTermino, 4) = "0229" Then
        pTermino = Left(pTermino, 4) & "0228"
    End If
    '1. Inicializa valores auxiliares para cálculo da diferença de dias.
    lAuxIniDias = "001032060091121152182213244274305335" & _
                  "002033061092122153183214245275306336" & _
                  "003034062093123154184215246276307337" & _
                  "004035063094124155185216247277308338" & _
                  "005036064095125156186217248278309339" & _
                  "006037065096126157187218249279310340" & _
                  "007038066097127158188219250280311341" & _
                  "008039067098128159189220251281312342" & _
                  "009040068099129160190221252282313343" & _
                  "010041069100130161191222253283314344" & _
                  "011042070101131162192223254284315345" & _
                  "012043071102132163193224255285316346" & _
                  "013044072103133164194225256286317347" & _
                  "014045073104134165195226257287318348" & _
                  "015046074105135166196227258288319349" & _
                  "016047075106136167197228259289320350" & _
                  "017048076107137168198229260290321351" & _
                  "018049077108138169199230261291322352" & _
                  "019050078109139170200231262292323353" & _
                  "020051079110140171201232263293324354" & _
                  "021052080111141172202233264294325355" & _
                  "022053081112142173203234265295326356" & _
                  "023054082113143174204235266296327357" & _
                  "024055083114144175205236267297328358" & _
                  "025056084115145176206237268298329359"
    lAuxIniDias = lAuxIniDias & _
                  "026057085116146177207238269299330360" & _
                  "027058086117147178208239270300331361" & _
                  "028059087118148179209240271301332362" & _
                  "029000088119149180210241272302333363" & _
                  "030000089120150181211242273303334364" & _
                  "031000090000151000212243000304000365"

    lAuxTerDias = "366397425456486517547578609639670700" & _
                  "367398426457487518548579610640671701" & _
                  "368399427458488519549580611641672702" & _
                  "369400428459489520550581612642673703" & _
                  "370401429460490521551582613643674704" & _
                  "371402430461491522552583614644675705" & _
                  "372403431462492523553584615645676706" & _
                  "373404432463493524554585616646677707" & _
                  "374405433464494525555586617647678708" & _
                  "375406434465495526556587618648679709" & _
                  "376407435466496527557588619649680710" & _
                  "377408436467497528558589620650681711" & _
                  "378409437468498529559590621651682712" & _
                  "379410438469499530560591622652683713" & _
                  "380411439470500531561592623653684714" & _
                  "381412440471501532562593624654685715" & _
                  "382413441472502533563594625655686716" & _
                  "383414442473503534564595626656687717" & _
                  "384415443474504535565596627657688718" & _
                  "385416444475505536566597628658689719" & _
                  "386417445476506537567598629659690720" & _
                  "387418446477507538568599630660691721" & _
                  "388419447478508539569600631661692722" & _
                  "389420448479509540570601632662693723" & _
                  "390421449480510541571602633663694724"
    lAuxTerDias = lAuxTerDias & _
                  "391422450481511542572603634664695725" & _
                  "392423451482512543573604635665696726" & _
                  "393424452483513544574605636666697727" & _
                  "394000453484514545575606637667698728" & _
                  "395000454485515546576607638668699729" & _
                  "396000455000516000577608000669000730"
    lPrazoLongo = 0

    '2. Posiciona dados da tabela de data de início.
    lI = 1
    lDia = 1
    Do Until lDia > 31
        lMes = 1
        Do Until lMes > 12
            lTabIniDias(lDia, lMes) = Val(Mid(lAuxIniDias, lI, 3))
            lMes = lMes + 1
            lI = lI + 3
        Loop
        lDia = lDia + 1
    Loop

    '3. Posiciona dados da tabela de data de término.
    lDia = 1
    lI = 1
    Do Until lDia > 31
        lMes = 1
        Do Until lMes > 12
            lTabTerDias(lDia, lMes) = Val(Mid(lAuxTerDias, lI, 3))
            lMes = lMes + 1
            lI = lI + 3
        Loop
        lDia = lDia + 1
    Loop

    '4. Apropria informação de dia e mês de início e término.
    lIniDia = Val(Mid(pInicio, 7, 2))
    lIniMes = Val(Mid(pInicio, 5, 2))
    lTerDia = Val(Mid(pTermino, 7, 2))
    lTerMes = Val(Mid(pTermino, 5, 2))

    If Val(Mid(pTermino, 1, 4)) - Val(Mid(pInicio, 1, 4)) > 2 Then
        lPrazoLongo = 730
    ElseIf Val(Mid(pTermino, 1, 4)) - Val(Mid(pInicio, 1, 4)) > 1 Then
        lPrazoLongo = 365
    End If

    '5. Posiciona quantidade de dias.
    lQuaIniDias = Val(lTabIniDias(lIniDia, lIniMes))
    If Val(Mid(pInicio, 1, 4)) <> Val(Mid(pTermino, 1, 4)) Then
        lQuaTerDias = lTabTerDias(lTerDia, lTerMes) + lPrazoLongo
    Else
        lQuaTerDias = lTabIniDias(lTerDia, lTerMes) + lPrazoLongo
    End If

    '6. Calcula diferença de dias.
    gfDifDatas = lQuaTerDias - lQuaIniDias
        
End Function


Public Function gfBusProduto(ByVal pVigInicio As Double, ByRef pCodProduto As Integer, _
                             ByRef pCodProtocolo As Long, ByRef pCodSequencia As Long, _
                             ByRef pDescProduto As String, ByRef pMensagem As String) As Boolean
    'Procedure: busca produto de acordo com início de vigência.

    'Parâmetro de entrada...pVigInicio......Início de vigência da cotação/proposta.
    'Parâmetros de saída....pCodProduto.....Código do produto.
    '                       pCodProtocolo...Código do protocolo.
    '                       pCodSequencia...Código da seqüência.
    '                       pMensagem.......Mensagem.

    Dim bdP0043700      As ADODB.Connection 'Arquivo - P0043700.
    Dim lAuxAno         As Integer          'Auxiliar: ano.
    Dim lAuxiliar       As String           'Auxiliar.
    Dim lAuxMes         As Integer          'Auxiliar: mês.
    Dim lDatLimite      As Double           'Data limite.
    Dim lAuxData        As Double           'Data auxiliar
    Dim lAuxPreData     As String           'Auxiliar.
    Dim lDias           As Integer          'Dias.
    Dim lDifDias        As Integer          'Dif dias.
    Dim rsP0043700      As ADODB.Recordset  'P0043700: tabela.
    Dim lGrupo          As String
    Dim lSql            As String
    gfBusProduto = False

    '1. Inicializa variáveis.
    pCodProduto = 0
    pCodProtocolo = 0
    pCodSequencia = 0

    '2. Abre arquivo.
    If gfAbrBasDados("P0043700", gTipP0043700, gEndP0043700, bdP0043700, pMensagem) = False Then
        Exit Function
    End If
    '3. Obtém produto sem término de vigência ou com término de vigência superior à data do dia.
    lSql = "SELECT * FROM P0043700 WHERE Tip_Prot = 1 AND Dat_Vig_Ini <= " & _
              pVigInicio
    If gCod_Corretor <> 0 Then
        Select Case gCod_Corretor
            Case 8001, 8002, 8003, 8004, 8005, 8006, 8008, 8011, 8014, 8028
                lGrupo = 8001
            Case 699499
                lGrupo = 329006
            Case Else
                lGrupo = gCod_Corretor
        End Select
        If gConcessionaria Then
            Select Case gCodConcessionaria
                Case e_CodConcessionaria_Honda
                    lSql = lSql & " and cod_protocolo in (31040)"
                Case e_CodConcessionaria_KIA
                    lSql = lSql & " and cod_protocolo in (31041)"
                Case e_CodConcessionaria_Cherry
                    lSql = lSql & " and cod_protocolo in (31042)"
                Case Else
                    lSql = lSql & " and cod_protocolo in (31040)"
            End Select
        Else
            lSql = lSql & " and cod_protocolo in (0, " & lGrupo & ")"
        End If
    End If
    lSql = lSql & " ORDER BY Cod_Produto DESC , "
    If gCod_Corretor <> 0 Then
        lSql = lSql & " cod_protocolo desc, "
    Else
        lSql = lSql & " cod_protocolo asc, "
    End If
    lSql = lSql & " Cod_Sequencia desc"
    If gfObtRegistro(bdP0043700, lSql, rsP0043700, pMensagem) = False Then
        Call gpFecha2(bdP0043700)
        Exit Function
    End If
    
    If rsP0043700.EOF Then
        lSql = "SELECT * FROM P0043700 WHERE Dat_Vig_Ini <= " & pVigInicio & _
                  " ORDER BY Cod_Produto DESC"
        If gfObtRegistro(bdP0043700, lSql, rsP0043700, pMensagem) = False Then
            Call gpFecha2(bdP0043700)
            Exit Function
        End If
        If rsP0043700.EOF Then
            Call gpFecha3(rsP0043700, bdP0043700)
            gfBusProduto = True
            Exit Function
        End If
    End If
    
    '4. Verifica produtos disponíveis.
    If pVigInicio > rsP0043700!Dat_Vig_Ter Then
        lAuxiliar = Format$(rsP0043700!Dat_Vig_Ter, "00000000")
        lAuxMes = CInt(Mid$(lAuxiliar, 5, 2)) + 1
        lAuxAno = CInt(Mid$(lAuxiliar, 1, 4))
        If lAuxMes = 13 Then
            lAuxMes = 1
            lAuxAno = lAuxAno + 1
        End If
        lDias = -1
        On Error Resume Next
        lDias = rsP0043700!Dias_Expiracao
        If lDias = -1 Then
            lDias = 15
        End If
        On Error GoTo 0
        lAuxPreData = CDate(Mid(rsP0043700!Dat_Vig_Ter, 7, 2) & "/" & Mid(rsP0043700!Dat_Vig_Ter, 5, 2) & "/" & Mid(rsP0043700!Dat_Vig_Ter, 1, 4)) + lDias
        lAuxData = Format(lAuxPreData, "YYYYMMDD")
        
        'Documento dentro do limite do produto (até x dias após término da
        'vigência) - sem problema.
        If Not (pVigInicio >= rsP0043700!Dat_Vig_Ini And pVigInicio <= lAuxData) Then
            pMensagem = "FGE0123 - Não há produto (tarifa) disponível para este cálculo, de acordo com a data de início de vigência (dd/mm/aaaa). Aguardar novo produto (tarifa). Caso seja necessário, o cálculo deve ser feito pela Yasuda."
            Call gpFecha3(rsP0043700, bdP0043700)
            gfBusProduto = True
            Exit Function
        End If
    End If

    '5. Posiciona dados de saída.
    pCodProduto = rsP0043700!Cod_Produto
    pCodProtocolo = rsP0043700!Cod_Protocolo
    pCodSequencia = rsP0043700!Cod_Sequencia
    pDescProduto = rsP0043700("Nom_Prot")
    Call gpFecha3(rsP0043700, bdP0043700)

    gfBusProduto = True
End Function

Public Function gfCalNosNumero(ByRef pNosNumero As String, ByRef pMensagem As String) As Boolean
    'Função: gera nosso número a partir de usuário, ano, dia juliano e número de segundos decorridos no
    '        dia.
    
    'Parâmetros de saída...pNosNumero...Nosso número.
    '                      pMensagem....Mensagem.

    'Formato do nosso número:
    '   AAADDDSSSSSUUUUUUUUD...AAA........ano.
    '                          DDD........dia juliano (dias decorridos no ano).
    '                          SSSSS......segundos decorridos no dia.
    '                          UUUUUUUU...usuário.
    '                          D..........dígito verificador (pela Prova 11).

    Dim bdP0044700      As ADODB.Connection 'Arquivo - P0044700.
    Dim clsA99V001      As clsA99V001       'Calcula dígitos.
    Dim lAuxAno         As String           'Ano no formato string.
    Dim lAuxDia         As String           'Dia no formato string.
    Dim lAuxSegundo     As String           'Segundo no formato string.
    Dim lCodUsuario     As String           'Código do usuário.
    Dim lDia            As Integer          'Dia.
    Dim lDigito         As Integer          'Dígito.
    Dim lMesBissexto    As String           'Controle de dias do ano bissexto.
    Dim lMesNormal      As String           'Controle de dias do ano normal.
    Dim lNosNumero      As String           'Nosso número.
    Dim lSegundo        As Double           'Segundo.
    Dim lSelect         As String           'Comando SELECT.
    Dim rsUSER          As ADODB.Recordset  'Registros de usuários.

    gfCalNosNumero = False

    '1. Posiciona dados para nosso número.
    'a) Ano.
    lAuxAno = Right$(Format$(gHojAno, "0000"), 3)
    'b) Dia juliano.
    lMesNormal = "000031059090120150181212243273303334365"
    lMesBissexto = "000031060091121151182213244274304335366"
    lDia = gHojDia
    Select Case gHojAno Mod 4
        Case 0      'Ano bissexto.
            lDia = lDia + Val(Mid(lMesBissexto, gHojMes * 3 - 2, 3))
        Case Else   'Ano NÃO bissexto.
            lDia = lDia + Val(Mid(lMesNormal, gHojMes * 3 - 2, 3))
    End Select
    lAuxDia = Right("000" & Trim(Str(lDia)), 3)
    'c) Segundos decorridos do dia.
    lSegundo = (Hour(Now) * 3600) + (Minute(Now) * 60) + Second(Now)
    lAuxSegundo = Right(Format$(lSegundo, "00000"), 5)
    'd) Código do usuário.
    Select Case gTipUsuario
        Case 1      'Yasuda.
            If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
                Exit Function
            End If
            lSelect = "SELECT COD_USER FROM YTAB_USER WHERE COD_LOGIN_USER = '" & UCase(gNomUsuario) & "'"
            If gfObtRegistro(bdP0044700, lSelect, rsUSER, pMensagem) = False Then
                Call gpFecha2(bdP0044700)
                Exit Function
            End If
            If rsUSER.EOF = True Then
                lCodUsuario = "00000000"
            Else
                lCodUsuario = Format$(rsUSER!Cod_User, "00000000")
            End If
            Call gpFecha3(rsUSER, bdP0044700)
        Case 2      'Corretor.
            lCodUsuario = gNomUsuario
        Case 3      'Concessionária/Funcionário
            lCodUsuario = gCodSeguradoSYAS
    End Select
    If IsNumeric(lCodUsuario) Then
        lCodUsuario = Trim$(lCodUsuario)
        If Len(lCodUsuario) < 8 Then
            lCodUsuario = String(8 - Len(lCodUsuario), "0") & lCodUsuario
        End If
    Else
        lCodUsuario = "00000000"
    End If
    If lCodUsuario = "00000000" Then
        Randomize: lCodUsuario = Left$(Format$(CDbl((10000000 * Rnd) + 1), "00000000"), 8)
    End If
    
    'e) Monta nosso número sem dígito.
    lNosNumero = lAuxAno & lAuxDia & lAuxSegundo & lCodUsuario

    '2. Calcula dígito.
    Set clsA99V001 = New clsA99V001
    If clsA99V001.DIG_MODULO11(lNosNumero, "20191817161514131211100908070605040302", lDigito, _
       pMensagem) = False Then
        Exit Function
    End If
    Set clsA99V001 = Nothing

    '3. Completa nosso número.
    pNosNumero = lNosNumero & Format$(lDigito, "0")
    gfCalNosNumero = True
End Function
Public Function gfDisponivel(ByVal pNome As String) As Boolean
    'Função: indica se form está disponível.

    Dim lI  As Byte     'Utilizado com For...Next.

    gfDisponivel = False
    For lI = 0 To (Forms.Count - 1)
        If UCase(Forms(lI).Name) = UCase(pNome) Then
            gfDisponivel = True
            Exit For
        End If
    Next lI
End Function
Public Function gfForAno(pData As Integer) As String
    If Left$(pData, 2) > 50 Then    'De 1951 a 1999.
        gfForAno = "19" & Format$(pData, "00")
    Else                            'De 2000 a 2050.
        gfForAno = "20" & Format$(pData, "00")
    End If
End Function

Public Function gfForAspas(ByVal pCampo As String) As String
    'Função: formata campo texto com aspas para ser utilizado no comando SQL.

    Dim lAuxiliar   As String       'Utilizado com InStr.
    Dim lPosicao    As Integer      'Utilizado com InStr.
    Dim lTemAspa    As Boolean      'Indicador de que tem aspa.
    
    '1. Se campo não preenchido, retorna campo nulo.
    If Not gfPreenchido(pCampo) Then
        gfForAspas = ""
        Exit Function
    End If

    '2. Posiciona controle.
    lTemAspa = True

    '3. Verifica se campo tem aspa.
    lPosicao = InStr(pCampo, "'")
    
    '4. Verifica se campo tem "|".
    If lPosicao = 0 Then
        gfForAspas = pCampo
        lTemAspa = False
        lPosicao = InStr(pCampo, "|")
        If lPosicao = 0 Then
            gfForAspas = pCampo
            Exit Function
        End If
    End If

    '5. Se campo tem aspa, duplica campo de aspa para comando SQL.
    lAuxiliar = pCampo
    If lTemAspa = True Then
        Do
            If lPosicao = Len(lAuxiliar) Then
                gfForAspas = lAuxiliar & "'"
                Exit Function
            End If
            lAuxiliar = Left$(lAuxiliar, lPosicao) & "'" & Mid$(lAuxiliar, lPosicao + 1)
            lPosicao = InStr(lPosicao + 2, lAuxiliar, "'")
            If lPosicao = 0 Then
                lPosicao = InStr(pCampo, "|")
                If lPosicao = 0 Then
                    gfForAspas = lAuxiliar
                    Exit Function
                Else
                    Exit Do
                End If
            End If
        Loop
    End If

    '6. Retira "|" do campo.
    Do
        If lPosicao = Len(lAuxiliar) Then
            gfForAspas = Left$(lAuxiliar, lPosicao - 1)
            Exit Function
        End If
        lAuxiliar = Left$(lAuxiliar, lPosicao - 1) & Mid$(lAuxiliar, lPosicao + 1)
        lPosicao = InStr(lPosicao + 1, lAuxiliar, "|")
        If lPosicao = 0 Then
            gfForAspas = lAuxiliar
            Exit Function
        End If
    Loop
End Function
Public Function gfForData(ByVal pOpcao As Byte, Optional ByVal pData, Optional ByVal pTipArquivo) _
                          As String
    'Função: formata data.

    'Formato dos campos de data (numérico duplo): aaaammdd.

    'Parâmetros...pOpcao.........01 = formata campo de "dd/mm/aaaa" para "aaaammdd" para comandos SQL.
    '                            02 = formata campo de "aaaammdd" para "dd/mm/aaaa".
    '                            03 = fornece data atual no formato aaaammddhhmmss.
    '                            04 = formata campo de "aaaammddhhmmss" para "dd/mm/aaaa hh:mm:ss".
    '                            05 = formata campo de "aaaammdd" para formato conveniente
    '                                 ('mm/dd/aaaa' ou #mm/dd/aaaa#).
    '                            06 = formata campo de "dd/mm/aaaa" para formato conveniente
    '                                 ('mm/dd/aaaa' ou #mm/dd/aaaa#).
    '                            07 = formata campo de "aammdd" para "dd/mm/aaaa".
    '                            08 = formata campo de ano para "aaaa".
    '                            09 = trata campo de data na digitação (utilizado nos eventos LostFocus).
    '                            10 = formata campo de "aammdd" para "aaaammdd".
    '             pData..........Campo de data a ser tratado (opcional).
    '             pTipArquivo...Tipo do arquivo (opcional - utilizado para opção 5).

    Dim lAno        As String   'Ano.
    Dim lAuxData    As String   'Auxiliar: data.
    Dim lDia        As String   'Dia.
    Dim lMes        As String   'Mês.
    Dim lPosicao    As Integer  'Utilizado com InStr.

    '1. Verifica opção.
    Select Case pOpcao
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
            'Nada a fazer - opções válidas.
        Case Else
            Call gpGraLog(2, "FTE0006 - Opção inválida (formata data - opção desconhecida: " & pOpcao & ").")
            Exit Function
    End Select

    '2. Processa formatação de data.
    Select Case pOpcao
        Case 1
            If gfPreenchido(pData) Then
                If Len(pData) = 10 Then
                    gfForData = Mid$(pData, 7, 4) & Mid$(pData, 4, 2) & Mid$(pData, 1, 2)
                Else
                    gfForData = "00000000"
                End If
            Else
                gfForData = "00000000"
            End If
        Case 2
            If gfPreenchido(pData) Then
                lAuxData = Format$(pData, "00000000")
            Else
                lAuxData = "00000000"
            End If
            gfForData = Mid$(lAuxData, 7, 2) & "/" & Mid$(lAuxData, 5, 2) & "/" & Mid$(lAuxData, 1, 4)
        Case 3
            gfForData = Format$(Now, "yyyymmddhhnnss")
        Case 4
            If gfPreenchido(pData) Then
                lAuxData = Format$(pData, "00000000000000")
            Else
                lAuxData = "00000000000000"
            End If
            gfForData = Mid$(lAuxData, 7, 2) & "/" & _
                        Mid$(lAuxData, 5, 2) & "/" & _
                        Mid$(lAuxData, 1, 4) & " " & _
                        Mid$(lAuxData, 9, 2) & ":" & _
                        Mid$(lAuxData, 11, 2) & ":" & _
                        Mid$(lAuxData, 13, 4)
        Case 5
            lAuxData = Format$(pData, "00000000")
            Select Case pTipArquivo
                Case "1"        'Access.
                    gfForData = "#" & Mid$(lAuxData, 5, 2) & "/" & Mid$(lAuxData, 7, 2) & "/" & _
                                Mid$(lAuxData, 1, 4) & "#"
                Case "2"        'SQL-Server
                    gfForData = "'" & Mid$(lAuxData, 5, 2) & "/" & Mid$(lAuxData, 7, 2) & "/" & _
                                Mid$(lAuxData, 1, 4) & "'"
            End Select
        Case 6
            Select Case pTipArquivo
                Case "1"        'Access.
                    gfForData = "#" & Mid$(pData, 4, 2) & "/" & Mid$(pData, 1, 2) & "/" & _
                                Mid$(pData, 7, 4) & "#"
                Case "2"        'SQL-Server
                    gfForData = "'" & Mid$(pData, 4, 2) & "/" & Mid$(pData, 1, 2) & "/" & _
                                Mid$(pData, 7, 4) & "'"
            End Select
        Case 7
            If gfPreenchido(pData) Then
                lAuxData = Format$(pData, "000000")
            Else
                lAuxData = "000000"
            End If
            If Left$(lAuxData, 2) > 50 Then 'De 1951 a 1999.
                gfForData = Mid$(lAuxData, 5, 2) & "/" & _
                            Mid$(lAuxData, 3, 2) & "/19" & _
                            Left$(lAuxData, 2)
            Else                            'De 2000 a 2050.
                gfForData = Mid$(lAuxData, 5, 2) & "/" & _
                            Mid$(lAuxData, 3, 2) & "/20" & _
                            Left$(lAuxData, 2)
            End If
        Case 8
            If gfPreenchido(pData) Then
                lAuxData = Format$(pData, "0000")
            Else
                lAuxData = "0000"
            End If
            If Left$(lAuxData, 2) = "00" Then
                If Right$(lAuxData, 2) > 50 Then 'De 1951 a 1999.
                    gfForData = "19" & Right$(lAuxData, 2)
                Else                            'De 2000 a 2050.
                    gfForData = "20" & Right$(lAuxData, 2)
                End If
            Else
                gfForData = lAuxData
            End If
        Case 9
            If IsMissing(pData) Then
                gfForData = ""
            Else
                If Not gfPreenchido(pData) Then
                    gfForData = ""
                Else
                    gfForData = lAuxData
                    If IsNumeric(pData) Then
                        Select Case Len(Trim(pData))
                            Case 6      'Provável digitação: ddmmaa.
                                lDia = Mid$(pData, 1, 2)
                                lMes = Mid$(pData, 3, 2)
                                lAno = Mid$(pData, 5, 2)
                                If Val(lAno) > 50 Then  'De 1951 a 1999.
                                    lAno = "19" & lAno
                                Else                    'De 2000 a 2050.
                                    lAno = "20" & lAno
                                End If
                                gfForData = lDia & "/" & lMes & "/" & lAno
                            Case 8      'Provável digitação: ddmmaaaa.
                                gfForData = Mid$(pData, 1, 2) & "/" & _
                                            Mid$(pData, 3, 2) & "/" & _
                                            Mid$(pData, 5, 4)
                            Case Else
                                gfForData = pData
                        End Select
                    Else
                        lAuxData = Trim$(pData)
                        'Dia.
                        lPosicao = InStr(lAuxData, "/")
                        Select Case lPosicao
                            Case 2
                                lDia = "0" & Mid$(lAuxData, 1, 1)
                            Case 3
                                lDia = Mid$(lAuxData, 1, 2)
                            Case Else
                                Exit Function
                        End Select
                        If Not IsNumeric(lDia) Then
                            Exit Function
                        End If
                        If Len(lAuxData) = lPosicao Then
                            Exit Function
                        End If
                        lAuxData = Mid$(lAuxData, lPosicao + 1)
                        'Mês.
                        lPosicao = InStr(lAuxData, "/")
                        Select Case lPosicao
                            Case 2
                                lMes = "0" & Mid$(lAuxData, 1, 1)
                            Case 3
                                lMes = Mid$(lAuxData, 1, 2)
                            Case Else
                                Exit Function
                        End Select
                        If Not IsNumeric(lMes) Then
                            Exit Function
                        End If
                        If Len(lAuxData) = lPosicao Then
                            Exit Function
                        End If
                        'Ano.
                        lAno = Mid$(lAuxData, lPosicao + 1)
                        If Not IsNumeric(lAno) Then
                            Exit Function
                        End If
                        If Len(lAno) = 1 Then
                            lAno = "0" & lAno
                        End If
                        If Len(lAno) = 2 Then
                            If Val(lAno) > 50 Then  'De 1951 a 1999.
                                lAno = "19" & lAno
                            Else                    'De 2000 a 2050.
                                lAno = "20" & lAno
                            End If
                        End If
                        If Len(lAno) = 3 Then
                            lAno = "0" & lAno
                        End If
                        If Len(lAno) > 4 Then
                            lAno = Left$(lAno, 4)
                        End If
                        gfForData = lDia & "/" & lMes & "/" & lAno
                    End If
                    If Not IsDate(gfForData) Then
                        gfForData = ""
                    End If
                End If
            End If
        Case 10
            If gfPreenchido(pData) Then
                lAuxData = Format$(pData, "000000")
                If Left$(lAuxData, 2) > 50 Then 'De 1951 a 1999.
                    gfForData = "19" & lAuxData
                Else                            'De 2000 a 2050.
                    gfForData = "20" & lAuxData
                End If
            Else
                lAuxData = "00000000"
            End If
    End Select
End Function
Public Function gfForValor(ByVal pOpcao As Byte, ByVal pCampo As String, Optional pPadrao As String) _
                           As String
    'Função: formata valor.
    
    'Parâmetros...pOpcao....1 = troca vírgula por ponto e tira ponto, para ser utilizado no comando SQL.
    '                       2 = formata no padrão standard e coloca brancos à esquerda para ser utilizado
    '                           com os formulários.
    '                       3 = coloca brancos à esquerda para ser utilizado com os formulários.
    '             pCampo....campo a ser formatado.
    '             pPadrao...para opção 1, informa padrão a ser utilizado caso campo não esteja preenchido.
    '                       para opção 2 e 3, informa tamanho do campo (se omitido, assume 15 posições).

    Dim lAuxiliar   As String   'Auxiliar.
    Dim lI          As Byte     'Utilizado com For...Next.
    Dim lTamanho    As Byte     'Tamanho.
    Dim lTexto      As String   'Auxiliar para tirar ponto.

    '1. Verifica opção.
    If pOpcao <> 1 And pOpcao <> 2 And pOpcao <> 3 Then
        Call gpGraLog(2, "FTE0006 - Opção inválida (formata valor - opção diferente de 1, 2 e 3: " & _
                         pOpcao & ").")
        Exit Function
    End If

    '2. Processa formatação de data.
    Select Case pOpcao
        Case 1
            If gfPreenchido(pCampo) Then
                If IsNumeric(pCampo) Then
                    lAuxiliar = Trim$(pCampo)
                    lTexto = ""
                    For lI = 1 To Len(lAuxiliar)
                        Select Case Mid(lAuxiliar, lI, 1)
                            Case ","
                                lTexto = lTexto & "."
                            Case "."
                                'lTexto = lTexto
                            Case Else
                                lTexto = lTexto & Mid(lAuxiliar, lI, 1)
                        End Select
                    Next lI
                    If lTexto = "" Then
                        lTexto = "0"
                    End If
                    gfForValor = lTexto
                    Exit Function
                End If
            End If
            If IsMissing(pPadrao) Then
                gfForValor = "0"
            Else
                If gfPreenchido(pPadrao) Then
                    gfForValor = pPadrao
                Else
                    gfForValor = "0"
                End If
            End If
        Case 2
            If gfPreenchido(pCampo) Then
                If IsNumeric(pCampo) Then
                    lAuxiliar = Trim$(Format$(pCampo, "standard"))
                Else
                    lAuxiliar = "0,00"
                End If
            Else
                lAuxiliar = "0,00"
            End If
            If IsMissing(pPadrao) Then
                lTamanho = 15
            Else
                If Not IsNumeric(pPadrao) Then
                    lTamanho = 15
                Else
                    lTamanho = pPadrao
                End If
            End If
            If Len(lAuxiliar) > lTamanho Then
                gfForValor = lAuxiliar
            Else
                gfForValor = Space(lTamanho - Len(lAuxiliar)) & lAuxiliar
            End If
        Case 3
            If IsMissing(pPadrao) Then
                lTamanho = 5
            Else
                If Not IsNumeric(pPadrao) Then
                    lTamanho = 5
                Else
                    lTamanho = pPadrao
                End If
            End If
            If gfPreenchido(pCampo) Then
                lAuxiliar = Trim$(pCampo)
                If Len(lAuxiliar) > lTamanho Then
                    gfForValor = lAuxiliar
                Else
                    gfForValor = Space(lTamanho - Len(lAuxiliar)) & lAuxiliar
                End If
            Else
                gfForValor = Space(lTamanho)
            End If
    End Select
End Function


Public Function gfNomCorretor(ByVal pCodCorretor As String, ByRef pNomCorretor As String, _
                              ByRef pMensagem As String) As Boolean
    'Função: obtém nome do corretor.

    'Parâmetro de entrada...pCodCorretor...Código do corretor.
    'Parâmetros de saída....pNomCorretor...Nome do corretor.
    '                       pMensagem......Mensagem.

    Dim bdP0044700  As ADODB.Connection 'Arquivo - P0044700.
    Dim lSelect     As String           'Comando SELECT.
    Dim rsCORR      As ADODB.Recordset  'YTAB_CORR: tabela.

    gfNomCorretor = False

    '1. Verifica preenchimento.
    If Not gfPreenchido(pCodCorretor) Then
        pNomCorretor = "Código do corretor não preenchido."
        gfNomCorretor = True
        Exit Function
    End If
    If Val(pCodCorretor) = 0 Then
        pNomCorretor = "Código do corretor zerado."
        gfNomCorretor = True
        Exit Function
    End If

    '2. Obtém nome do corretor.
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    lSelect = "SELECT NOM_ABRV_CORR FROM YTAB_CORR WHERE COD_CORR = " & pCodCorretor
    If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
        Call gpFecha2(bdP0044700)
        Exit Function
    End If
    If rsCORR.EOF = True Then
        Call gpFecha3(rsCORR, bdP0044700)
        pMensagem = "FGE0044 - " & gPrefixo & "Corretor inválido (código de corretor não localizado " & _
                    "- " & pCodCorretor & ")."
        Exit Function
    End If
    pNomCorretor = Trim$(rsCORR!NOM_ABRV_CORR)
    Call gpFecha3(rsCORR, bdP0044700)

    gfNomCorretor = True
End Function
Public Function gfBuscaSusep(ByVal pCodCorretor As String, ByRef pSusep As String, _
                              ByRef pMensagem As String) As Boolean
    'Função: obtém nome do corretor.

    'Parâmetro de entrada...pCodCorretor...Código do corretor.
    'Parâmetros de saída....pNomCorretor...Nome do corretor.
    '                       pMensagem......Mensagem.

    Dim bdP0044700  As ADODB.Connection 'Arquivo - P0044700.
    Dim lSelect     As String           'Comando SELECT.
    Dim rsCORR      As ADODB.Recordset  'YTAB_CORR: tabela.

    gfBuscaSusep = False

    '1. Verifica preenchimento.
    If Not gfPreenchido(pCodCorretor) Then
        gfBuscaSusep = ""
        gfBuscaSusep = True
        Exit Function
    End If
    If Val(pCodCorretor) = 0 Then
        gfBuscaSusep = ""
        gfBuscaSusep = True
        Exit Function
    End If

    '2. Obtém nome do corretor.
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    lSelect = "SELECT NUM_SUSEP FROM YTAB_CORR WHERE COD_CORR = " & pCodCorretor
    If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
        Call gpFecha2(bdP0044700)
        Exit Function
    End If
    If rsCORR.EOF = True Then
        Call gpFecha3(rsCORR, bdP0044700)
        pMensagem = "FGE0044 - " & gPrefixo & "Corretor inválido (código de corretor não localizado " & _
                    "- " & pCodCorretor & ")."
        Exit Function
    End If
    pSusep = Trim$(rsCORR!NUM_SUSEP)
    Call gpFecha3(rsCORR, bdP0044700)

    gfBuscaSusep = True
End Function

Public Function gfNomSeguradora(ByRef pCodSeguradora As String, ByRef pNomSeguradora As String, _
                                ByRef pMensagem As String) As Boolean
    'Função: obtém nome da seguradora

    'Parâmetro de entrada...pCodSeguradora...Código da seguradora.
    'Parâmetros de saída....pNomSeguradora...Nome da seguradora.
    '                       pMensagem........Mensagem.

    Dim bdP0044700  As ADODB.Connection 'Arquivo - P0044700.
    Dim lSelect     As String           'Comando SELECT.
    Dim rsP0044700  As ADODB.Recordset  'P0044700: tabela.

    gfNomSeguradora = False

    '1. Verifica preenchimento.
    If Not gfPreenchido(pCodSeguradora) Then
        pMensagem = "FTE0039 - " & gPrefixo & _
                    "Erro ao obter nome da seguradora (código da seguradora não preenchido)."
        Exit Function
    End If
    If Not IsNumeric(pCodSeguradora) Then
        pMensagem = "FTE0039 - " & gPrefixo & _
                    "Erro ao obter nome da seguradora (código da seguradora não numérico)."
        Exit Function
    End If
    If Val(pCodSeguradora) = 0 Then
        pMensagem = "FTE0039 - " & gPrefixo & _
                    "Erro ao obter nome da seguradora (código da seguradora zerado)."
        Exit Function
    End If

    '2. Obtém nome da seguradora.
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    lSelect = "SELECT Seguradora FROM YTAB_SEGURADORA WHERE CSeguradora = '" & Trim(pCodSeguradora) & "'"
    If gfObtRegistro(bdP0044700, lSelect, rsP0044700, pMensagem) = False Then
        Call gpFecha2(bdP0044700)
        Exit Function
    End If
    If rsP0044700.EOF = True Then
        pMensagem = "FTE0039 - " & gPrefixo & _
                    "Erro ao obter nome da seguradora (código da seguradora não localizado - " & _
                    pCodSeguradora & ")."
        Exit Function
    End If
    pNomSeguradora = rsP0044700!Seguradora
    Call gpFecha3(rsP0044700, bdP0044700)

    gfNomSeguradora = True
End Function
Public Function gfVerCorrPermitido(ByVal pCorrInstalacao As Long, ByVal pCorrSelecionado As Long) As Boolean
Dim lCorrInstacao       As Long                 'Código do corretor - instalação
Dim lCorrSelecionado    As Long                 'Código do corretor - selecionado
Dim lEncArquivo         As Boolean              'Identifica se a verificação foi feita através do arquivo, caso contrário deverá ser buscado do código
Dim lEncontrou          As Boolean              'Identifica se foi encontrato o [grupos] no arquivo
Dim lEndP366CORR        As String               'Endereço do arquivo P366CORR.EXE.
Dim lNumP366CORR        As Integer              'Utilizado com FreeFile para o arquivo P366CORR.EXE.
Dim lRegistro           As String               'Registro.
    
    'A comparação no arquivo é feita com o número do corretor sem os dois ultimos digitos,
    'os dois ultimos digitos variam mas dentro de um mesmo grupo de corretor.
    'No arquivo P366CORR.EXE não tem o número completo do corretor, está sem os dois últimos
    'dígitos.
    lCorrInstacao = Left(Format(pCorrInstalacao, "0000000"), 5)
    lCorrSelecionado = Left(Format(pCorrSelecionado, "0000000"), 5)
    lEndP366CORR = gAppPath & "P366CORR.EXE"
    If Not gfPreenchido(Dir(lEndP366CORR)) Then
        lEncArquivo = False
    Else
        lEncArquivo = True
    End If
    If lEncArquivo = True Then
        lNumP366CORR = FreeFile
        On Error Resume Next
        Open lEndP366CORR For Input Access Read As lNumP366CORR
        If Err <> 0 Then
            lEncArquivo = False
        End If
        If lEncArquivo = True Then
            On Error GoTo 0
            '4. Trata registros do arquivo de inicialização.
            lEncontrou = False
            lEncArquivo = False
            Do While Not EOF(lNumP366CORR)
                Line Input #lNumP366CORR, lRegistro
                If gfPreenchido(lRegistro) Then
                    lRegistro = Trim$(lRegistro)
                    If UCase(lRegistro) = "[GRUPOS]" Then
                        lEncontrou = True
                    End If
                    If UCase(lRegistro) = "[SOMENTE MULTI]" Then
                        Exit Do
                    End If
                    If UCase(lRegistro) = "[INTEGRACAO]" Then
                        Exit Do
                    End If
                    If lEncontrou = True Then
                        If InStr(1, lRegistro, ";" & lCorrInstacao & ";") > 0 Then
                            If InStr(1, lRegistro, ";" & lCorrSelecionado & ";") > 0 Then
                                gfVerCorrPermitido = True
                            Else
                                gfVerCorrPermitido = False
                            End If
                            lEncArquivo = True
                            Exit Do
                        End If
                        If lCorrInstacao <> lCorrSelecionado Then
                            lEncArquivo = False
                            gfVerCorrPermitido = False
                        Else
                            lEncArquivo = True
                            gfVerCorrPermitido = True
                            Exit Do
                        End If
                    End If
                End If
            Loop
            Close lNumP366CORR
        End If
    End If
    If lEncArquivo = False Then
        Select Case lCorrInstacao
                Case 809, 2387
                    Select Case lCorrSelecionado
                        Case 809, 2387
                            'Seleção OK
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 376, 3813
                    Select Case lCorrSelecionado
                        Case 376, 3813
                            gfVerCorrPermitido = True
                        Case Else
                                'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2234, 904
                    Select Case lCorrSelecionado
                        Case 2234, 904
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 183, 4148
                    Select Case lCorrSelecionado
                        Case 183, 4148
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 77, 4246, 4708, 4709, 4710
                    Select Case lCorrSelecionado
                        Case 77, 4246, 4708, 4709, 4710
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4618, 1102
                    Select Case lCorrSelecionado
                        Case 4618, 1102
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 843, 1172, 845, 377
                    Select Case lCorrSelecionado
                        Case 843, 1172, 845, 377
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6449, 1749
                    Select Case lCorrSelecionado
                        Case 6449, 1749
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1889, 495
                    Select Case lCorrSelecionado
                        Case 1889, 495
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6106, 4978, 4800, 4792, 4846, 2769, 3223
                    Select Case lCorrSelecionado
                        Case 6106, 4978, 4800, 4792, 4846, 2769, 3223
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3272, 34, 3335, 393, 3070, 73
                    Select Case lCorrSelecionado
                        Case 3272, 34, 3335, 393, 3070, 73
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2577, 2101
                    Select Case lCorrSelecionado
                        Case 2577, 2101
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2422, 6145
                    Select Case lCorrSelecionado
                        Case 2422, 6145
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 172, 3407
                    Select Case lCorrSelecionado
                        Case 172, 3407
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 855, 1170, 1171, 2582, 6659
                    Select Case lCorrSelecionado
                        Case 855, 1170, 1171, 2582, 6659
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 60, 3301, 3302
                    Select Case lCorrSelecionado
                        Case 60, 3301, 3302
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1184, 709
                    Select Case lCorrSelecionado
                        Case 1184, 709
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 389, 4252
                    Select Case lCorrSelecionado
                        Case 389, 4252
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2505, 4162
                    Select Case lCorrSelecionado
                        Case 2505, 4162
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 96, 534, 1454
                    Select Case lCorrSelecionado
                        Case 96, 534, 1454
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 822, 406
                    Select Case lCorrSelecionado
                        Case 822, 406
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 530, 2171
                    Select Case lCorrSelecionado
                        Case 530, 2171
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6369, 2994, 6544
                    Select Case lCorrSelecionado
                        Case 6369, 2994, 6544
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3668, 1496
                    Select Case lCorrSelecionado
                        Case 3668, 1496
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4675, 6446
                    Select Case lCorrSelecionado
                        Case 4675, 6446
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1491, 3821
                    Select Case lCorrSelecionado
                        Case 1491, 3821
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3272, 34, 4726, 4431, 3438, 3302, 73, 3456, 4723, 3459, 3070, 3392
                    Select Case lCorrSelecionado
                        Case 3272, 34, 4726, 4431, 3438, 3302, 73, 3456, 4723, 3459, 3070, 3392
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2360, 2888
                    Select Case lCorrSelecionado
                        Case 2360, 2888
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4565, 4669
                    Select Case lCorrSelecionado
                        Case 4565, 4669
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 748, 1069
                    Select Case lCorrSelecionado
                        Case 748, 1069
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1155, 2501, 1255, 6205
                    Select Case lCorrSelecionado
                        Case 1155, 2501, 1255, 6205
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3494, 3189, 3583, 3509, 356, 2956
                    Select Case lCorrSelecionado
                        Case 3494, 3189, 3583, 3509, 356, 2956
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3075, 3367
                    Select Case lCorrSelecionado
                        Case 3075, 3367
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1863, 6238, 4621, 3311, 6237, 6244, 6242, 6248, 6231, 1860, 6191, 6243, 4622, 6658
                    Select Case lCorrSelecionado
                        Case 1863, 6238, 4621, 3311, 6237, 6244, 6242, 6248, 6231, 1860, 6191, 6243, 4622, 6658
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 322, 6361, 263, 2311, 3865, 4866, 6439
                    Select Case lCorrSelecionado
                        Case 322, 6361, 263, 2311, 3865, 4866, 6439
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4267, 6678
                    Select Case lCorrSelecionado
                        Case 4267, 6678
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 214, 6125
                    Select Case lCorrSelecionado
                        Case 214, 6125
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2841, 2842
                    Select Case lCorrSelecionado
                        Case 2841, 2842
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6593, 1465
                    Select Case lCorrSelecionado
                        Case 6593, 1465
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2453, 4485
                    Select Case lCorrSelecionado
                        Case 2453, 4485
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4503, 4353, 3416
                    Select Case lCorrSelecionado
                        Case 4503, 4353, 3416
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3111, 1614
                    Select Case lCorrSelecionado
                        Case 3111, 1614
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2355, 2344
                    Select Case lCorrSelecionado
                        Case 2355, 2344
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1116, 1901, 1321, 1890
                    Select Case lCorrSelecionado
                        Case 1116, 1901, 1321, 1890
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2451, 1148
                    Select Case lCorrSelecionado
                        Case 2451, 1148
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3043, 6515
                    Select Case lCorrSelecionado
                        Case 3043, 6515
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2142, 133
                    Select Case lCorrSelecionado
                        Case 2142, 133
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3775, 3098, 1807, 1769, 3082, 6621, 768, 4213, 3081
                    Select Case lCorrSelecionado
                        Case 3775, 3098, 1807, 1769, 3082, 6621, 768, 4213, 3081
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6435, 1688
                    Select Case lCorrSelecionado
                        Case 6435, 1688
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3379, 6625
                    Select Case lCorrSelecionado
                        Case 3379, 6625
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2145, 871
                    Select Case lCorrSelecionado
                        Case 2145, 871
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2369, 162, 469, 483
                    Select Case lCorrSelecionado
                        Case 2369, 162, 469, 483
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2912, 2911
                    Select Case lCorrSelecionado
                        Case 2912, 2911
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2479, 2889, 4352
                    Select Case lCorrSelecionado
                        Case 2479, 2889, 4352
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3610, 6592, 4561, 3642, 3271, 4245, 3742, 1640, 1923, 2859, 3667
                    Select Case lCorrSelecionado
                        Case 3610, 6592, 4561, 3642, 3271, 4245, 3742, 1640, 1923, 2859, 3667
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3544, 3543, 3547, 3541, 3542, 3545, 3546, 3548, 2740
                    Select Case lCorrSelecionado
                        Case 3544, 3543, 3547, 3541, 3542, 3545, 3546, 3548, 2740
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2734, 4886
                    Select Case lCorrSelecionado
                        Case 2734, 4886
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3457, 3458, 3459, 3456
                    Select Case lCorrSelecionado
                        Case 3457, 3458, 3459, 3456
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4573, 6008
                    Select Case lCorrSelecionado
                        Case 4573, 6008
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1060, 2702
                    Select Case lCorrSelecionado
                        Case 1060, 2702
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 58, 6487, 4037
                    Select Case lCorrSelecionado
                        Case 58, 6487, 4037
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3982, 3844, 3119, 6080, 3149, 6378
                    Select Case lCorrSelecionado
                        Case 3982, 3844, 3119, 6080, 3149, 6378
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2909, 2182
                    Select Case lCorrSelecionado
                        Case 2909, 2182
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 62, 6616
                    Select Case lCorrSelecionado
                        Case 62, 6616
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2621, 179
                    Select Case lCorrSelecionado
                        Case 2621, 179
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3438, 3392
                    Select Case lCorrSelecionado
                        Case 3438, 3392
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3319, 6131
                    Select Case lCorrSelecionado
                        Case 3319, 6131
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 827, 337
                    Select Case lCorrSelecionado
                        Case 827, 337
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4481, 2147, 4589
                    Select Case lCorrSelecionado
                        Case 4481, 2147, 4589
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 80, 4086
                    Select Case lCorrSelecionado
                        Case 80, 4086
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6240, 6198, 6195, 3397, 6210, 6256, 6298, 6219, 6209, 6442, 6420, 6172, 6297, 6301, 6193, 6226, 6413, 6218, 1895, 6228, 1864, 6419, 6588, 6208, 6224, 1435, 6261, 6175, 6596, 6184
                    Select Case lCorrSelecionado
                        Case 6240, 6198, 6195, 3397, 6210, 6256, 6298, 6219, 6209, 6442, 6420, 6172, 6297, 6301, 6193, 6226, 6413, 6218, 1895, 6228, 1864, 6419, 6588, 6208, 6224, 1435, 6261, 6175, 6596, 6184
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6128, 4700, 6217, 1771, 4661, 1872, 6372
                    Select Case lCorrSelecionado
                        Case 6128, 4700, 6217, 1771, 4661, 1872, 6372
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1825, 4118, 4389, 6611, 6404, 4724, 4391, 1705, 4775, 6293, 4393, 4117, 4619, 4394, 6606, 4463, 6294, 4556, 6580, 4396, 4576, 4916, 4398, 4457, 4482, 6603, 4458, 4399, 6291, 1893, 4735, 6609, 4693, 1707, 6394, 6636, 1871, 1761, 6009, 4538, 6655, 6098, 6406, 6351, 1828, 4405, 6255, 4640, 4407, 4408, 4702, 6352, 6643, 4483, 4410, 1823, 1715, 6440, 1827, 4413, 4429, 4414, 4416, 1791, 6232, 6299, 6485, 4587, 6608, 4419, 6626, 4448, 6418, 6019, 6182, 4518, 4690, 4633, 4818, 6190, 6513, 1911
                    Select Case lCorrSelecionado
                        Case 1825, 4118, 4389, 6611, 6404, 4724, 4391, 1705, 4775, 6293, 4393, 4117, 4619, 4394, 6606, 4463, 6294, 4556, 6580, 4396, 4576, 4916, 4398, 4457, 4482, 6603, 4458, 4399, 6291, 1893, 4735, 6609, 4693, 1707, 6394, 6636, 1871, 1761, 6009, 4538, 6655, 6098, 6406, 6351, 1828, 4405, 6255, 4640, 4407, 4408, 4702, 6352, 6643, 4483, 4410, 1823, 1715, 6440, 1827, 4413, 4429, 4414, 4416, 1791, 6232, 6299, 6485, 4587, 6608, 4419, 6626, 4448, 6418, 6019, 6182, 4518, 4690, 4633, 4818, 6190, 6513, 1911
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4966, 4875, 4900, 4830, 4898
                    Select Case lCorrSelecionado
                        Case 4966, 4875, 4900, 4830, 4898
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3611, 3708, 1448, 1468, 1464, 1460, 3485, 3486, 2747, 3100, 3612, 3484, 3511, 3482, 3483, 3585, 1449, 3502, 3892, 3531, 1474, 1484, 3569, 1451, 3117, 1485, 3481, 3636, 3581, 3638, 3480, 3568, 3732, 3637, 3487, 2585, 3627, 3385, 3493, 3665, 3273, 3654, 3499
                    Select Case lCorrSelecionado
                        Case 3611, 3708, 1448, 1468, 1464, 1460, 3485, 3486, 2747, 3100, 3612, 3484, 3511, 3482, 3483, 3585, 1449, 3502, 3892, 3531, 1474, 1484, 3569, 1451, 3117, 1485, 3481, 3636, 3581, 3638, 3480, 3568, 3732, 3637, 3487, 2585, 3627, 3385, 3493, 3665, 3273, 3654, 3499
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case Else
                    If lCorrInstacao <> lCorrSelecionado Then
                        gfVerCorrPermitido = False
                    Else
                        gfVerCorrPermitido = True
                    End If
        End Select
    End If

End Function
Public Function gfVerCorrIntegracao() As Boolean
Dim lCodCorretor2       As Long                 'Código do corretor - instalação
Dim lEncArquivo         As Boolean              'Identifica se a verificação foi feita através do arquivo, caso contrário deverá ser buscado do código
Dim lIndTransmissao     As String
Dim lEncontrou          As Boolean              'Identifica se foi encontrato o [grupos] no arquivo
Dim lEndP366CORR        As String               'Endereço do arquivo P366CORR.EXE.
Dim lNumP366CORR        As Integer              'Utilizado com FreeFile para o arquivo P366CORR.EXE.
Dim lRegistro           As String               'Registro.
    
    'A comparação no arquivo é feita com o número do corretor sem os dois ultimos digitos,
    'os dois ultimos digitos variam mas dentro de um mesmo grupo de corretor.
    'No arquivo P366CORR.EXE não tem o número completo do corretor, está sem os dois últimos
    'dígitos.
    
    If gTipUsuario = "1" Then
        gfVerCorrIntegracao = False
        Exit Function
    End If
    lEndP366CORR = gAppPath & "P366CORR.EXE"
    If Not gfPreenchido(Dir(lEndP366CORR)) Then
        lEncArquivo = False
    Else
        lEncArquivo = True
    End If
    If lEncArquivo = True Then
        lNumP366CORR = FreeFile
        On Error Resume Next
        Open lEndP366CORR For Input Access Read As lNumP366CORR
        If Err <> 0 Then
            lEncArquivo = False
        End If
        If lEncArquivo = True Then
            On Error GoTo 0
            '4. Trata registros do arquivo de inicialização.
            lEncontrou = False
            lEncArquivo = False
            Do While Not EOF(lNumP366CORR)
                Line Input #lNumP366CORR, lRegistro
                If gfPreenchido(lRegistro) Then
                    lRegistro = Trim$(lRegistro)
                    If UCase(lRegistro) = "[INTEGRACAO]" Then
                        lEncontrou = True
                    End If
                    
                    If lEncontrou = True Then
                        Select Case Len(lRegistro)
                            Case 8
                                lCodCorretor2 = Left$(lRegistro, 7)
                                lIndTransmissao = "N"
                            Case 9
                                lCodCorretor2 = Left$(lRegistro, 7)
                                lIndTransmissao = Mid$(lRegistro, 9, 1)
                            Case Else
                                lCodCorretor2 = "0000000"
                                lIndTransmissao = "N"
                        End Select
                    End If
                    If lCodCorretor2 = gCod_Corretor Then
                        Exit Do
                    Else
                        lIndTransmissao = "N"
                    End If
                End If
            Loop
            Close lNumP366CORR
        End If
    End If
    
    If lIndTransmissao = "S" Then
        gfVerCorrIntegracao = True
    Else
        gfVerCorrIntegracao = False
    End If
    
End Function

Public Function gfObtCorretor(ByVal pOpcao As Byte, ByVal pCorretor As String, _
                              ByRef pSaida As Variant, ByRef pMensagem As String, _
                              Optional pTipEmissao As Long)

    'Parâmetros de entrada...pOpcao......1 = Código do corretor e 2 = Nome do corretor.
    '                        pCorretor...Código ou nome do corretor, de acordo com pOpcao.
    'Parâmetro de saída......pSaida......Array contendo resultados.
    '                                    (x, 1) Código do corretor.
    '                                    (x, 2) Nome do corretor.
    '                                    (x, 3) Nome abreviado do corretor.
    Dim bdP0044700  As ADODB.Connection 'Arquivo - P0044700.
    Dim lI          As Double           'Contador.
    Dim lSelect     As String           'Comando SELECT.
    Dim rsCORR      As ADODB.Recordset  'Tabela: YTAB_CORR.

    gfObtCorretor = False

    '1. Verifica parâmetros.
    'a) Opção.
    If pOpcao <> 1 And pOpcao <> 2 Then
        Call gpGraLog(2, "FTE0006 - Opção inválida (obter corretor - opção diferente de 1 e 2).")
        Exit Function
    End If
    'b) Corretor.
    If Not gfPreenchido(pCorretor) And gPrograma <> "D028O003" Then
        Call gpGraLog(2, "FTE0040 - Erro ao obter corretor (código do corretor não preenchido).")
        Exit Function
    End If

    '2. Busca informação.
    Select Case pOpcao
        Case 1      'Código do corretor.
            If Not IsNumeric(pCorretor) Then
                Call gpGraLog(2, "FTE0040 - Erro ao obter corretor (código do corretor não numérico)." & _
                                 vbLf & vbLf & "O sistema será encerrado.")
                gFinalizarPrograma = True
                Exit Function
            End If
            lSelect = "SELECT COD_CORR, NOM_CORR, NOM_ABRV_CORR FROM YTAB_CORR WHERE COD_CORR = " & _
                      pCorretor
            Select Case gTipoAcesso
                Case CALCULO, 0
                    lSelect = lSelect & " AND COD_ATIVO = 1"
                Case DIGITACAO
                    Select Case pTipEmissao
                        Case 300, 301, 303
                            'Pode ser emitido com corretor desativado
                        Case Else
                            lSelect = lSelect & " AND COD_ATIVO = 1"
                    End Select
            End Select
                
            If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
                Exit Function
            End If
            If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
                Exit Function
            End If
            If rsCORR.RecordCount = 0 Then
                Call gpFecha3(rsCORR, bdP0044700)
                pMensagem = "FGE0044 - " & gPrefixo & "Corretor inválido (código de corretor não " & _
                            "localizado - " & pCorretor & ")."
                Exit Function
            End If
            ReDim pSaida(1, 3)
            pSaida(1, 1) = rsCORR!Cod_Corr
            pSaida(1, 2) = rsCORR!NOM_CORR
            pSaida(1, 3) = rsCORR!NOM_ABRV_CORR
            Call gpFecha3(rsCORR, bdP0044700)
            gfObtCorretor = True
            Exit Function
        Case 2      'Nome do corretor.
            lSelect = "SELECT COD_CORR, NOM_CORR, NOM_ABRV_CORR FROM YTAB_CORR " & _
                      "WHERE (NOM_CORR LIKE '%" & pCorretor & "%' " & _
                      "OR     NOM_ABRV_CORR LIKE '%" & pCorretor & "%') "
            Select Case gTipoAcesso
                Case CALCULO, 0
                    lSelect = lSelect & " AND COD_ATIVO = 1"
                Case DIGITACAO
                    Select Case pTipEmissao
                        Case 300, 301, 303
                            'Pode ser emitido com corretor desativado
                        Case Else
                            lSelect = lSelect & " AND COD_ATIVO = 1"
                    End Select
            End Select
                                            
            If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
                Exit Function
            End If
            If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
                Exit Function
            End If
            If rsCORR.RecordCount = 0 Then
                Call gpFecha3(rsCORR, bdP0044700)
                pMensagem = "FGE0044 - Corretor inválido (corretor não localizado com nome ou parte " & _
                            "do nome = " & pCorretor & ")."
                Exit Function
            End If
            rsCORR.MoveLast
            ReDim Preserve pSaida(rsCORR.RecordCount, 3)
            lI = 0
            rsCORR.MoveFirst
            While Not rsCORR.EOF
                lI = lI + 1
                pSaida(lI, 1) = rsCORR!Cod_Corr
                pSaida(lI, 2) = rsCORR!NOM_CORR
                pSaida(lI, 3) = rsCORR!NOM_ABRV_CORR
                rsCORR.MoveNext
            Wend
            Call gpFecha3(rsCORR, bdP0044700)
            gfObtCorretor = True
            Exit Function
    End Select
End Function
Public Function gfObtProduto(ByVal pCodProduto As Integer, ByVal pCodProtocolo As Long, _
                             ByVal pCodSequencia As Long, ByRef pSaida As Variant, _
                             ByRef pMensagem As String) As Boolean
    'Função: obtém dados do produto - utilizada apenas no módulo P28V300.

    'Parâmetros de entrada...pCodProduto.....Código do produto.
    '                        pCodProtocolo...Código do protocolo.
    '                        pCodSequencia...Código da seqüência.
    'Parâmetros de saída.....pSaida..........Matriz com dados do produto.
    '                                        (0) Nome do produto.
    '                                        (1) Limite da comissão.
    '                                        (2) Comissão básica.
    '                                        (3) Juros de IOF.
    '                        pMensagem.......Mensagem.

    Dim bdP0043700      As ADODB.Connection 'Arquivo - P0043700.
    Dim lJurIOF         As Double           'Juros de IOF.
    Dim lNomProduto     As String           'Nome do produto.
    Dim lPerComBasica   As Double           'Percentual de comissão básica.
    Dim lPerComLimite   As Double           'Percentual de comissão: limite.
    Dim lSelect         As String           'Comando SELECT.
    Dim rsP0043700      As ADODB.Recordset  'P0043700: tabela.

    gfObtProduto = False

    '1. Abre arquivo.
    If gfAbrBasDados("P0043700", gTipP0043700, gEndP0043700, bdP0043700, pMensagem) = False Then
        Exit Function
    End If

    '2. Obtém registro.
    lSelect = "SELECT Nom_Prot, Per_Cms_Limite, Per_Cms_Basica, Jur_IOF FROM P0043700 WHERE " & _
              "Cod_Produto = " & pCodProduto & " AND Cod_Protocolo = " & pCodProtocolo & _
              " AND Cod_Sequencia = " & pCodSequencia
    If gfObtRegistro(bdP0043700, lSelect, rsP0043700, pMensagem) = False Then
        Call gpFecha2(bdP0043700)
        Exit Function
    End If
    If rsP0043700.EOF = True Then
        Call gpFecha3(rsP0043700, bdP0043700)
        pMensagem = "FNL0032 - " & gPrefixo & "Produto não localizado: " & _
                    Format$(pCodProduto, "0000") & "/" & Format$(pCodProtocolo, "0000") & "/" & _
                    Format$(pCodSequencia, "0000") & "."
        Exit Function
    End If
    lNomProduto = rsP0043700!Nom_Prot
    lPerComLimite = rsP0043700!Per_Cms_Limite
    lPerComBasica = rsP0043700!Per_Cms_Basica
    lJurIOF = rsP0043700!Jur_IOF
    Call gpFecha3(rsP0043700, bdP0043700)
    pSaida = Array(lNomProduto, lPerComLimite, lPerComBasica, lJurIOF)

    gfObtProduto = True
End Function
Public Function gfObtUniProdutiva(ByVal pCodCorretor As String, ByRef pCodUniNegocios As Integer, _
                                  ByRef pCodProdutor As Integer, ByRef pMensagem As String, _
                                Optional ByRef pTotalRegistros As Integer) As Boolean
    'Função: obtém unidade de negócios e produtor do corretor.

    'Parâmetro de entrada...pCodCorretor......Código do corretor.
    'Parâmetros de saída....pCodUniNegocios...Código da unidade de negócios.
    '                       pCodProdutor......Código do produtor.
    '                       pMensagem.........Mensagem.

    Dim bdP0044700      As ADODB.Connection 'Arquivo - P0044700.
    Dim lI              As Double           'Contador.
    Dim lNomCorretor    As String           'Nome do corretor.
    Dim lSelect         As String           'Comando SELECT.
    Dim rsUNIDCORRPROD  As ADODB.Recordset  'Registro - YTAB_UNID_CORR_PROD.

    gfObtUniProdutiva = False

    '1. Inicializa saída.
    pCodUniNegocios = 0
    pCodProdutor = 0

    '2. Verifica parâmetro.
    If Not gfPreenchido(pCodCorretor) Then
        pMensagem = "FTE0041 - Erro ao obter unidade de negócios e produtor (código do corretor não " & _
                    "preenchido)."
        Exit Function
    End If

    '3. Busca informação.
    If Not IsNumeric(pCodCorretor) Then
        pMensagem = "FTE0041 - Erro ao obter unidade negócios e produtor (código do corretor não " & _
                    "numérico)."
        Exit Function
    End If
    lSelect = "SELECT COD_UNID, COD_PROD FROM YTAB_UNID_CORR_PROD WHERE COD_CORR = " & pCodCorretor & _
              " ORDER BY COD_CORR, COD_UNID"
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    If gfObtRegistro(bdP0044700, lSelect, rsUNIDCORRPROD, pMensagem) = False Then
        Exit Function
    End If
    If rsUNIDCORRPROD.RecordCount = 0 Then
        Call gpFecha3(rsUNIDCORRPROD, bdP0044700)
        If gfNomCorretor(pCodCorretor, lNomCorretor, pMensagem) = False Then
            Exit Function
        End If
        pMensagem = "FTE0041 - Erro ao obter unidade de negócios e produtor (corretor " & pCodCorretor & _
                    " - " & lNomCorretor & " sem informação de unidade de negócios e produtor)."
        Exit Function
    Else
        pTotalRegistros = rsUNIDCORRPROD.RecordCount
    End If

    '4. Encerra processo.
    pCodUniNegocios = rsUNIDCORRPROD!COD_UNID
    pCodProdutor = rsUNIDCORRPROD!COD_PROD
    Call gpFecha3(rsUNIDCORRPROD, bdP0044700)

    gfObtUniProdutiva = True
End Function
Public Function gfObtListaUniProd(ByVal pCodCorretor As String, ByRef pMensagem As String, _
                                ByRef pLista() As String) As Boolean
    'Função: obtém unidade de negócios e produtor do corretor.

    'Parâmetro de entrada...pCodCorretor......Código do corretor.
    'Parâmetros de saída....pCodUniNegocios...Código da unidade de negócios.
    '                       pCodProdutor......Código do produtor.
    '                       pMensagem.........Mensagem.

    Dim bdP0044700      As ADODB.Connection 'Arquivo - P0044700.
    Dim lI              As Double           'Contador.
    Dim lNomCorretor    As String           'Nome do corretor.
    Dim lSelect         As String           'Comando SELECT.
    Dim rsUNIDCORRPROD  As ADODB.Recordset  'Registro - YTAB_UNID_CORR_PROD.
    

    gfObtListaUniProd = False

    '1. Inicializa saída.

    '2. Verifica parâmetro.
    If Not gfPreenchido(pCodCorretor) Then
        pMensagem = "FTE0041 - Erro ao obter unidade de negócios e produtor (código do corretor não " & _
                    "preenchido)."
        Exit Function
    End If

    '3. Busca informação.
    If Not IsNumeric(pCodCorretor) Then
        pMensagem = "FTE0041 - Erro ao obter unidade negócios e produtor (código do corretor não " & _
                    "numérico)."
        Exit Function
    End If
    lSelect = " SELECT COD_UNID, NOM_DEPT  as nom_unid, c.COD_PROD , nom_prod FROM YTAB_UNID_CORR_PROD as c inner join "
    lSelect = lSelect & " ytab_produtor as p on c.cod_prod =  p.cod_prod inner join YTAB_DEPT as u on c.cod_unid = u.num_dept"
    lSelect = lSelect & " WHERE COD_CORR = " & pCodCorretor & _
              " ORDER BY COD_CORR, COD_UNID"
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    If gfObtRegistro(bdP0044700, lSelect, rsUNIDCORRPROD, pMensagem) = False Then
        Exit Function
    End If
    If rsUNIDCORRPROD.RecordCount = 0 Then
        Call gpFecha3(rsUNIDCORRPROD, bdP0044700)
        If gfNomCorretor(pCodCorretor, lNomCorretor, pMensagem) = False Then
            Exit Function
        End If
        pMensagem = "FTE0041 - Erro ao obter unidade de negócios e produtor (corretor " & pCodCorretor & _
                    " - " & lNomCorretor & " sem informação de unidade de negócios e produtor)."
        Exit Function
    
    End If
    lI = 0
    Do While Not rsUNIDCORRPROD.EOF
        lI = lI + 1
        ReDim Preserve pLista(4, lI)
        pLista(1, lI) = rsUNIDCORRPROD!COD_UNID
        pLista(2, lI) = rsUNIDCORRPROD!nom_unid
        pLista(3, lI) = rsUNIDCORRPROD!COD_PROD
        pLista(4, lI) = rsUNIDCORRPROD!nom_PROD
        rsUNIDCORRPROD.MoveNext
    Loop
    Call gpFecha3(rsUNIDCORRPROD, bdP0044700)

    gfObtListaUniProd = True
End Function

Public Function gfObtUnidCorr(ByVal pCodCorretor As String, ByRef pCodUniNegocios As Integer, ByRef pMensagem As String, _
                                  ByRef pTotalRegistros As Integer) As Boolean
    'Função: obtém unidade de negócios e produtor do corretor.

    'Parâmetro de entrada...pCodCorretor......Código do corretor.
    'Parâmetros de saída....pCodUniNegocios...Código da unidade de negócios.
    '                       pCodProdutor......Código do produtor.
    '                       pMensagem.........Mensagem.

    Dim bdP0044700      As ADODB.Connection 'Arquivo - P0044700.
    Dim lI              As Double           'Contador.
    Dim lNomCorretor    As String           'Nome do corretor.
    Dim lSelect         As String           'Comando SELECT.
    Dim rsUNIDCORRPROD  As ADODB.Recordset  'Registro - YTAB_UNID_CORR_PROD.

    gfObtUnidCorr = False

    '2. Verifica parâmetro.
    If Not gfPreenchido(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade de negócios e produtor (código do corretor não " & _
                    "preenchido)."
        Exit Function
    End If

    '3. Busca informação.
    If Not IsNumeric(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade negócios e produtor (código do corretor não " & _
                    "numérico)."
        Exit Function
    End If
    lSelect = "SELECT COD_PROD FROM YTAB_UNID_CORR_PROD WHERE " & _
                " COD_CORR = " & pCodCorretor & _
                " and cod_unid = " & pCodUniNegocios
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    If gfObtRegistro(bdP0044700, lSelect, rsUNIDCORRPROD, pMensagem) = False Then
        Exit Function
    End If
    If rsUNIDCORRPROD.RecordCount = 0 Then
        pTotalRegistros = 0
    Else
        pTotalRegistros = rsUNIDCORRPROD.RecordCount
    End If

    '4. Encerra processo.
    Call gpFecha3(rsUNIDCORRPROD, bdP0044700)

    gfObtUnidCorr = True
End Function

Public Function gfObtValMinCoberturas(ByVal pCodProduto As Integer, ByVal pCodProtocolo As Long, _
                                      ByVal pCodSequencia As Long, ByRef pSaida As Variant, _
                                      ByRef pMensagem As String) As Boolean
    'Função: obtém valores mínimos das coberturas.

    'Parâmetros de entrada...pCodProduto......Código do produto.
    '                        pCodProtocolo....Código do protocolo.
    '                        pCodSequencia....Código da seqüência.
    'Parâmetros de saída.....pSaida...........Matriz com valores iniciais.
    '                        - pSaida(1, n)...Código da cobertura.
    '                        - pSaida(2, n)...Valor inicial.
    '                        pMensagem........Mensagem.

    Dim bdP0043200      As ADODB.Connection     'Arquivo - P0043200.
    Dim lI              As Byte                 'Contador.
    Dim lSelect         As String               'Comando SELECT.
    Dim rsP0043200      As ADODB.Recordset      'Registro - P0043200.
    Dim lCodProtocolo   As Long                 'Cod Protocolo.
    Dim lCodSequencia   As Long                 'Cod Sequencia.
    
    lCodProtocolo = pCodProtocolo
    lCodSequencia = pCodSequencia
    
    gfObtValMinCoberturas = False

    '1. Obtém registros de coberturas do produto.
    If gfAbrBasDados("P0043200", 1, gEndP0043200, bdP0043200, pMensagem) = False Then
        Exit Function
    End If
    Do
        lSelect = "SELECT * FROM P0043200 WHERE Cod_Produto = " & pCodProduto & " AND Cod_Protocolo = " & _
                  lCodProtocolo & " AND Cod_Sequencia = " & lCodSequencia & " AND Cod_Categ = 0"
        If gfObtRegistro(bdP0043200, lSelect, rsP0043200, pMensagem) = False Then
            Call gpFecha2(bdP0043200)
            Exit Function
        End If

        '3. Verifica registros.
        lI = 0
        If rsP0043200.EOF = False Then
            lI = 0
            rsP0043200.MoveFirst
            While Not rsP0043200.EOF
                If rsP0043200!Val_IS_Inicial <> 0 Then
                    lI = lI + 1
                    If lI = 1 Then
                        ReDim pSaida(2, 1)
                    Else
                        ReDim Preserve pSaida(2, lI)
                    End If
                    pSaida(1, lI) = rsP0043200!Cod_Cobert
                    pSaida(2, lI) = rsP0043200!Val_IS_Inicial
                End If
                rsP0043200.MoveNext
            Wend
            Exit Do
        Else
            If lCodSequencia <> 0 Then
                lCodSequencia = lCodSequencia - 1
            Else
                If lCodProtocolo <> 0 Then
                    lCodProtocolo = 0
                    lCodSequencia = pCodSequencia
                Else
                    Exit Do
                End If
            End If
        End If
    Loop
    If lI = 0 Then
        ReDim pSaida(2, 1)
        pSaida(1, 1) = 0
        pSaida(2, 1) = 0
    End If
    Call gpFecha3(rsP0043200, bdP0043200)

    gfObtValMinCoberturas = True
End Function
Public Function gfRetornarListIndex(pCombo As Object, pItemData As String) As Long
    'Função:

    Dim lCont   As Long

    gfRetornarListIndex = -1
    For lCont = 0 To pCombo.ListCount - 1
        If pCombo.ItemData(lCont) = pItemData Then
            gfRetornarListIndex = lCont
            Exit Function
        End If
    Next
End Function
Public Function gfVerData(ByVal pOpcao As Byte, ByVal pData As String) As Boolean
    'Função: verifica data.

    'Parâmetros de entrada...pOpcao...Opção 1 = data no formato aaaammdd.
    '                                 Opção 2 = data no formato dd/mm/aaaa.
    '                        pData....Data no formato indicado pela opção.

    Dim lAno        As Integer  'Ano.
    Dim lAuxiliar   As String   'Auxiliar.
    Dim lDia        As Integer  'Dia.
    Dim lMes        As Integer  'Mês.
    Dim lResto      As Integer  'Resto.

    gfVerData = False

    '1. Verifica opção.
    If pOpcao <> 1 And pOpcao <> 2 Then
        Call gpGraLog(2, "FTE0006 - Opção inválida (verifica data - diferente de 1 e 2: " & pOpcao & _
                         ").")
        Exit Function
    End If

    '2. Separa dia, mês e ano.
    Select Case pOpcao
        Case 1      'Data no formato aaaammdd.
            lAuxiliar = Format$(pData, "00000000")
            lDia = Mid$(lAuxiliar, 7, 2)
            lMes = Mid$(lAuxiliar, 5, 2)
            lAno = Mid$(lAuxiliar, 1, 4)
        Case 2
            lDia = Mid$(pData, 1, 2)
            lMes = Mid$(pData, 4, 2)
            lAno = Mid$(pData, 7, 4)
    End Select

    '3. Verifica mês.
    If lMes < 1 Or lMes > 12 Then
        Exit Function
    End If

    '4. Verifica dia.
    If lDia < 1 Then
        Exit Function
    End If
    Select Case lMes
        Case 1, 3, 5, 7, 8, 10, 12
            If lDia > 31 Then
                Exit Function
            End If
        Case 4, 6, 9, 11
            If lDia > 30 Then
                Exit Function
            End If
        Case 2
            Select Case (lAno Mod 4)
                Case 0      'Ano bissexto.
                    If lDia > 29 Then
                        Exit Function
                    End If
                Case Else   'Ano NÃO bissexto.
                    If lDia > 28 Then
                        Exit Function
                    End If
            End Select
    End Select

    gfVerData = True
End Function
Public Function gfVerTransmissao() As Boolean
    'Função: verifica se corretor pode realizar transmissão de propostas.

    Dim lCodCorretor1       As String       'Código do corretor no arquivo P2842400.EXE.
    Dim lCodCorretor2       As String       'Código do corretor no arquivo P366CORR.EXE.
    Dim lEncontrou          As Boolean      'Indica que encontrou o registro de controle.
    Dim lEndP2842400        As String       'Endereço do arquivo P2842400.EXE.
    Dim lEndP366CORR        As String       'Endereço do arquivo P366CORR.EXE.
    Dim lIndTransmissao     As String       'Indicador de transmissão: S ou N.
    Dim lNumP2842400        As Integer      'Utilizado com FreeFile para o arquivo P2842400.EXE.
    Dim lNumP366CORR        As Integer      'Utilizado com FreeFile para o arquivo P366CORR.EXE.
    Dim lRegistro           As String       'Registro.

    gfVerTransmissao = False

    'Ficha nº 0901957
    'Liberar a transmissão para todos os corretores
    gfVerTransmissao = True
    Exit Function
    

    '1. Obtém código do corretor.
    lEndP2842400 = gAppPath & "P2842400.EXE"
    lNumP2842400 = FreeFile
    On Error Resume Next
    Open lEndP2842400 For Input Access Read As lNumP2842400
    If Err = 0 Then     'Encontrou arquivo P2842400.EXE.
        Line Input #lNumP2842400, lRegistro
        lCodCorretor1 = Format(lRegistro, "00000000")
    Else
        lCodCorretor1 = "00000000"
    End If
    On Error GoTo 0
    Close lNumP2842400

    '2. Verifica preenchimento.
    If Not gfPreenchido(lCodCorretor1) Then
        Call gpGraLog(1, "FGE0056 - Para transmissão de propostas para a Yasuda via Internet, " & _
                         "solicite o termo de ""ADESÃO AO SISTEMA DE TRANSMISSÃO DE PROPOSTAS DE " & _
                         "AUTOMÓVEL"" com o produtor (código do corretor não preenchido).")
        Exit Function
    End If
    lCodCorretor1 = Trim$(lCodCorretor1)
    If Not IsNumeric(lCodCorretor1) Then
        Call gpGraLog(1, "FGE0056 - Para transmissão de propostas para a Yasuda via Internet, " & _
                         "solicite o termo de ""ADESÃO AO SISTEMA DE TRANSMISSÃO DE PROPOSTAS DE " & _
                         "AUTOMÓVEL"" com o produtor (código do corretor não numérico).")
        Exit Function
    End If
    If Val(lCodCorretor1) = 0 Then
        Call gpGraLog(1, "FGE0056 - Para transmissão de propostas para a Yasuda via Internet, " & _
                         "solicite o termo de ""ADESÃO AO SISTEMA DE TRANSMISSÃO DE PROPOSTAS DE " & _
                         "AUTOMÓVEL"" com o produtor (código do corretor zerado).")
        Exit Function
    End If
    Select Case Len(lCodCorretor1)
        Case Is < 7
            lCodCorretor1 = String(7 - Len(lCodCorretor1), "0") & lCodCorretor1
        Case Is > 7
            lCodCorretor1 = Right$(lCodCorretor1, 7)
    End Select

    '3. Abre aquivo de controle de transmissão (temporário).
    lEndP366CORR = gAppPath & "P366CORR.EXE"
    If Not gfPreenchido(Dir(lEndP366CORR)) Then
        Call gpGraLog(1, "FTE0025 - Arquivo não localizado: controle do termo de adesão.")
        Exit Function
    End If
    lNumP366CORR = FreeFile
    On Error Resume Next
    Open lEndP366CORR For Input Access Read As lNumP366CORR
    If Err <> 0 Then
        Call gpGraLog(1, "FTE0026 - Erro ao abrir arquivo de controle do termo de adesão (" & Err & _
                         " - " & Error & ").")
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    '4. Trata registros do arquivo de inicialização.
    lEncontrou = False
    Do While Not EOF(lNumP366CORR)
        Line Input #lNumP366CORR, lRegistro
        If gfPreenchido(lRegistro) Then
            lRegistro = Trim$(lRegistro)
            If UCase(lRegistro) = "[GRUPOS]" Then
                Exit Do
            End If
            If UCase(lRegistro) = "[SOMENTE MULTI]" Then
                Exit Do
            End If
            If UCase(lRegistro) = "[INTEGRACAO]" Then
                Exit Do
            End If
            
            Select Case Len(lRegistro)
                Case 8
                    lCodCorretor2 = Left$(lRegistro, 7)
                    lIndTransmissao = "N"
                Case 9
                    lCodCorretor2 = Left$(lRegistro, 7)
                    lIndTransmissao = Mid$(lRegistro, 9, 1)
                Case Else
                    lCodCorretor2 = "0000000"
                    lIndTransmissao = "N"
            End Select
            If lCodCorretor1 = lCodCorretor2 Then
                lEncontrou = True
                Exit Do
            End If
        End If
    Loop
    Close lNumP366CORR
    If lEncontrou = False Or UCase(lIndTransmissao) <> "S" Then
        Call gpGraLog(1, "FGE0056 - Corretor não autorizado a efetuar transmissão via Internet. " & _
                         "Para maiores informações, entre em contato com seu executivo de contas). ")
                      Exit Function
    End If
    gfVerTransmissao = True
End Function

Public Sub gpPreencherCombo(ByRef pComboListBox As ComboBox, prst As ADODB.Recordset, _
                            Optional pCampoDescricao As String, Optional pCampoID As String, _
                            Optional pValorDefault As String, Optional pPosicaoDefault As Integer = -1, _
                            Optional pForm As Form = Nothing)
    'Procedure: preencher um listbox ou combobox a partir de uma coleção de objetos de estrutura.

    'Parâmetros...pComboListBox...combobox ou listbox que será preenchido.
    '             pColecao........coleção de objetos de estrutura que implementam a classe stcGenerica.
    '             pValorDefault...item que será posicionado o combo ou listbox.

    Dim bErroValor  As Boolean
    Dim bFoco       As Boolean

    On Error GoTo ControleErro

    bFoco = False
    If Not pForm Is Nothing Then
        If Not pForm.ActiveControl Is Nothing Then
            If UCase(TypeName(pForm.ActiveControl)) = "COMBOBOX" Then
                bFoco = (pForm.ActiveControl.Name = pComboListBox.Name And pForm.ActiveControl.Index = pComboListBox.Index)
            End If
        End If
    End If

    bErroValor = False
    If Not bFoco Then pComboListBox.Visible = False
    pComboListBox.Clear
    If Not (prst.BOF And prst.EOF) Then
        Do While Not prst.EOF
            pComboListBox.AddItem prst("" & pCampoDescricao & "")
            pComboListBox.ItemData(pComboListBox.NewIndex) = prst("" & pCampoID & "")
            prst.MoveNext
        Loop
        
        If pValorDefault <> "" Then
            bErroValor = True
            pComboListBox.Text = pValorDefault
        Else
            pComboListBox.ListIndex = pPosicaoDefault
        End If
    End If
    If Not bFoco Then pComboListBox.Visible = True
    Exit Sub

ControleErro:
    If bErroValor Then pComboListBox.ListIndex = pPosicaoDefault
    pComboListBox.Visible = True
End Sub
Public Function gfObtProdutoVar(ByVal pCodProduto As Integer, ByVal pCodProtocolo As Long, _
                         ByVal pCodSequencia As Long, pMensagem As String) As Boolean
    'Procedure: obtém produto e informações dependentes de produto.

    'Parâmetros de entrada...pCodProduto.....código do produto.
    '                        pCodProtocolo...código do protocolo.
    '                        pCodSequencia...código da seqüência.

    Dim bdP0043500      As ADODB.Connection 'Base de dados - P0043500.
    Dim bdP0043600      As ADODB.Connection 'Base de dados - P0043600.
    Dim bdP0043700      As ADODB.Connection 'Base de dados - P0043700.
    Dim bdP0045000      As ADODB.Connection 'Base de dados - P0045000.
    Dim lAntBonus       As String           'Anterior: bônus.
    Dim lAntFidelidade  As String           'Anterior: fidelidade.
    Dim lCodProtocolo   As Long             'Código do protocolo.
    Dim lCodSequencia   As Long             'Código da seqüência.
    Dim lI              As Byte             'Utilizado com For...Next.
    Dim lPosicao        As Integer          'Utilizado com InStr.
    Dim rsP0043500      As ADODB.Recordset  'Registro - P0043500.
    Dim rsP0043600      As ADODB.Recordset  'Registro - P0043600.
    Dim rsP0043700      As ADODB.Recordset  'Registro - P0043700.
    Dim rsP0045000      As ADODB.Recordset  'Registro - P0045000.
    Dim lCarga          As Boolean
    Dim lRetorno        As String
    Dim lSql            As String
    '1. Inicializa variáveis globais - dados do produto.
    gProTipProtocolo = 1
    gProNome = " "
    gProLimComissao = 0
    gProComLimite = 0
    gProComBasica = 0
    gProCusApolice = 0
    gProPreMinParcela = 0
    gProPreMinEndosso = 0
    gProLimMarComercial = 0
    gProQuaMaxParcelas = 0
    gProQuaMaxParcelasAnt = 0
    gProJurMinAnt01 = 3
    gProJurMinAnt02 = 3
    gProJurMinAnt03 = 3
    gProJurMinAnt04 = 3
    gProJurMinAnt05 = 3
    gProJurMinAnt06 = 3
    gProJurMinAnt07 = 3
    gProJurMinAnt08 = 3
    gProJurMinAnt09 = 3
    gProJurMinAnt10 = 3
    gProJurMinAnt11 = 3
    gProJurMinAnt12 = 3
    gProJurMinPos01 = 3
    gProJurMinPos02 = 3
    gProJurMinPos03 = 3
    gProJurMinPos04 = 3
    gProJurMinPos05 = 3
    gProJurMinPos06 = 3
    gProJurMinPos07 = 3
    gProJurMinPos08 = 3
    gProJurMinPos09 = 3
    gProJurMinPos10 = 3
    gProJurMinPos11 = 3
    gProJurMinPos12 = 3
    gProJurIOF = 0

    '2. Se produto zerado, encerra processo.
    If pCodProduto = 0 And pCodProtocolo = 0 And pCodSequencia = 0 Then
        Exit Function
    End If

    '3. Abre arquivo.
    If gfAbrBasDados("P0043700", gTipP0043700, gEndP0043700, bdP0043700, pMensagem) = False Then
        gCodProduto = 0
        gCodProtocolo = 0
        gCodSequencia = 0
        Exit Function
    End If

    '4. Obtém produto informado.
    lSql = "SELECT * FROM P0043700 WHERE Cod_Produto = " & pCodProduto & " AND Cod_Protocolo = " & _
              pCodProtocolo & " AND Cod_Sequencia = " & pCodSequencia
    If gfObtRegistro(bdP0043700, lSql, rsP0043700, pMensagem) = False Then
        Call gpFecha2(bdP0043700)
        gCodProduto = 0
        gCodProtocolo = 0
        gCodSequencia = 0
        Exit Function
    End If
    If rsP0043700.EOF = True Then
        Call gpFecha3(rsP0043700, bdP0043700)
        pMensagem = "FNL0032 - " & gPrefixo & "Produto não localizado (" & _
                         Format$(pCodProduto, "000") & "/" & Format$(pCodProtocolo, "000") & "/" & _
                         Format$(pCodSequencia, "000") & ")."
        gCodProduto = 0
        gCodProtocolo = 0
        gCodSequencia = 0
        Exit Function
    End If

    '.Verifica se usuário pode selecionar versão Interna (a partir da tarifa 9.00).
    Select Case gCodProduto
        Case Is >= 900
            If rsP0043700!Tip_Prot = 2 Then
                Select Case gPerfilAcesso
                    Case 1, 2, 3, 5, 7
                        'Acesso a versão Interna liberada.
                    Case Else
                        pMensagem = "FGE00XX - Versão Interna permitida somente para unidade de produtos Automóvel."
                        gCodProduto = 0
                        gCodProtocolo = 0
                        gCodSequencia = 0
                        Exit Function
                End Select
            End If
    End Select

    '5. Carrega variáveis globais.
    gCodProduto = rsP0043700!Cod_Produto
    gCodProtocolo = rsP0043700!Cod_Protocolo
    gCodSequencia = rsP0043700!Cod_Sequencia
    gProTipProtocolo = rsP0043700!Tip_Prot
    gProNome = rsP0043700!Nom_Prot
    gProVigInicio = rsP0043700!Dat_Vig_Ini
    gProVigTermino = rsP0043700!Dat_Vig_Ter
    gProLimComissao = rsP0043700!Per_Cms_Limite
    gProComLimite = rsP0043700!Per_Cms_Limite
    gProComBasica = rsP0043700!Per_Cms_Basica
    If gTipEmissao = 300 Or gTipEmissao = 301 Or gTipEmissao = 303 Or gTipEmissao = 302 Then
        gProCusApolice = rsP0043700!Val_Cus_Emis_Endo
    Else
        gProCusApolice = rsP0043700!Val_Cus_Emis_Apol
    End If
    gProPreMinParcela = rsP0043700!Pre_Min_Parcela
    gProPreMinEndosso = rsP0043700!Pre_Min_Endosso
    gProLimMarComercial = rsP0043700!Lim_Marg_Coml
    gProQuaMaxParcelas = rsP0043700!Qua_Max_Parcelas
    If gTipUsuario = "3" And gPlanoFuncionario = True Then
        gProQuaMaxParcelas = 11
    End If
    gProQuaMaxParcelasAnt = gProQuaMaxParcelas
    gProJurMinAnt01 = rsP0043700!Jur_Min_Ant_01
    gProJurMinAnt02 = rsP0043700!Jur_Min_Ant_02
    gProJurMinAnt03 = rsP0043700!Jur_Min_Ant_03
    gProJurMinAnt04 = rsP0043700!Jur_Min_Ant_04
    gProJurMinAnt05 = rsP0043700!Jur_Min_Ant_05
    gProJurMinAnt06 = rsP0043700!Jur_Min_Ant_06
    gProJurMinAnt07 = rsP0043700!Jur_Min_Ant_07
    gProJurMinAnt08 = rsP0043700!Jur_Min_Ant_08
    gProJurMinAnt09 = rsP0043700!Jur_Min_Ant_09
    gProJurMinAnt10 = rsP0043700!Jur_Min_Ant_10
    gProJurMinAnt11 = rsP0043700!Jur_Min_Ant_11
    gProJurMinAnt12 = rsP0043700!Jur_Min_Ant_12
    gProJurMinPos01 = rsP0043700!Jur_Min_Pos_01
    gProJurMinPos02 = rsP0043700!Jur_Min_Pos_02
    gProJurMinPos03 = rsP0043700!Jur_Min_Pos_03
    gProJurMinPos04 = rsP0043700!Jur_Min_Pos_04
    gProJurMinPos05 = rsP0043700!Jur_Min_Pos_05
    gProJurMinPos06 = rsP0043700!Jur_Min_Pos_06
    gProJurMinPos07 = rsP0043700!Jur_Min_Pos_07
    gProJurMinPos08 = rsP0043700!Jur_Min_Pos_08
    gProJurMinPos09 = rsP0043700!Jur_Min_Pos_09
    gProJurMinPos10 = rsP0043700!Jur_Min_Pos_10
    gProJurMinPos11 = rsP0043700!Jur_Min_Pos_11
    gProJurMinPos12 = rsP0043700!Jur_Min_Pos_12
    gProJurIOF = rsP0043700!Jur_IOF
    Call gpFecha3(rsP0043700, bdP0043700)
    
    gfObtProdutoVar = True
    
End Function
'Função para verificar grupo de corretor na base de dados
'Ficha 1000153 - Automatização do agrupamento de corretores no Syas após o cadastro realizado pelo help-desk.
Public Function gfVerCorrPermitido_TBS_Corr_Grupo(ByVal pCorrInstalacao As Long, ByVal pCorrSelecionado As Long, ByRef pMensagem As String) As Boolean
    Dim bdP0044700  As ADODB.Connection 'Arquivo - P0044700.
    Dim lSelect     As String           'Comando SELECT.
    Dim rsCORR      As ADODB.Recordset  'Tabela: YTAB_CORR.
    Dim lCorrInstacao As String
    Dim lCorrSelecionado As String

    gfVerCorrPermitido_TBS_Corr_Grupo = False

    lCorrInstacao = Format(pCorrInstalacao, "0000000")
    lCorrSelecionado = Format(pCorrSelecionado, "0000000")

    '1. Verifica parâmetros.
    'a) Opção.
    If Not gfPreenchido(pCorrInstalacao) Then
        Call gpGraLog(2, "FTE0067 - Opção inválida (obter corretor - opção não preenchida)." & vbLf & _
                         vbLf & "O sistema será encerrado.")
        Exit Function
    End If
    'b) Corretor.
    If Not gfPreenchido(pCorrSelecionado) Then
        Call gpGraLog(2, "FTE0068 - Erro ao obter corretor (código do corretor não preenchido)." & _
                         vbLf & vbLf & "O sistema será encerrado.")
        Exit Function
    End If

    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    'instalação  = Código aliado  e informado = cód líder
    lSelect = " SELECT * FROM TBS_Corr_Grupo " & _
              "WHERE CodAliado = '" & lCorrInstacao & _
              "' AND CodCorr = '" & lCorrSelecionado & _
              "' AND CodRamo = 310 "
    If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
        Exit Function
    End If
    If rsCORR.RecordCount = 0 Then
        'instalação  = cód líder  e  informado = Código aliado
        lSelect = " SELECT * FROM TBS_Corr_Grupo " & _
                  "WHERE CodCorr = '" & lCorrInstacao & _
                  "' AND CodAliado = '" & lCorrSelecionado & _
                  "' AND CodRamo = 310 "
        If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
            Exit Function
        End If
    Else
        gfVerCorrPermitido_TBS_Corr_Grupo = True
        Exit Function
    End If
    If rsCORR.RecordCount = 0 Then
        'instalação =  código aliado e informado = código aliado só que ambos são aliados do mesmo líder.
        lSelect = "SELECT A.CodCorr" & _
                  "  FROM TBS_Corr_Grupo AS A " & _
                  " INNER JOIN TBS_Corr_Grupo AS B " & _
                  "    ON A.CodCorr = B.CodCorr " & _
                  " WHERE (((A.CodAliado)= '" & lCorrInstacao & "' ) " & _
                  "   AND ((B.CodAliado)='" & lCorrSelecionado & "') " & _
                  "   AND ((A.CodRamo)=310)) "
        If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
            Exit Function
        End If
        If rsCORR.RecordCount = 0 Then
            Call gpFecha3(rsCORR, bdP0044700)
            pMensagem = "FGE0003 - " & gPrefixo & "Corretor inválido (código de corretor não " & _
                        "localizado - " & lCorrInstacao & ")."
            Exit Function
        Else
            gfVerCorrPermitido_TBS_Corr_Grupo = True
            Exit Function
        End If
    Else
        gfVerCorrPermitido_TBS_Corr_Grupo = True
        Exit Function
    End If
    
End Function

Public Function gfCopiaProposta(pAntNosNumero As String, ByRef lNovNosNumero As String, _
                                ByRef lNovReferencia As String, pMensagem As String, _
                                Optional pNaoRepProduto As Boolean, Optional pReentrada As Integer) As Boolean

Dim lSaida As String
Dim lCodProduto As Integer
Dim lCodProtocolo As Long
Dim lCodSequencia As Long

    '4. Calcula nosso número e referência.
    If gfCalNosNumero(lNovNosNumero, pMensagem) = False Then
        Exit Function
    End If
    If gfCalReferencia(lNovReferencia) = False Then
        'If gFinalizarPrograma = True Then End
        Exit Function
    End If
    gCodProduto = 0
    gCodProtocolo = 0
    gCodSequencia = 0
    If gTipUsuario <> "1" Then
'Ficha 901717
        If pNaoRepProduto = False Then 'Reposicionar na última tarifa
            If gfBusProduto(gHoje, lCodProduto, lCodProtocolo, lCodSequencia, lSaida, pMensagem) = False Then
                'If gFinalizarPrograma = True Then End
                Exit Function
            End If
            gCodProduto = lCodProduto
            gCodProtocolo = lCodProtocolo
            gCodSequencia = lCodSequencia
        End If
    End If

    '5. Copia documento.
    'a) Grava registro de dados comuns.
    If mfCopProposta("P0042100", gTipP0042100, gEndP0042100, "P0042100", pAntNosNumero, _
                     lNovNosNumero, lNovReferencia, pMensagem, pReentrada) = False Then
        Exit Function
    End If
    'b) Grava registros de itens.
    If mfCopProposta("P0042200", gTipP0042200, gEndP0042200, "P0042200", pAntNosNumero, _
                     lNovNosNumero, lNovReferencia, pMensagem, pReentrada) = False Then
        Exit Function
    End If
    'c) Grava registros de coberturas dos itens.
    If mfCopProposta("P0042600", gTipP0042600, gEndP0042600, "P0042600", pAntNosNumero, _
                     lNovNosNumero, lNovReferencia, pMensagem, pReentrada) = False Then
        Exit Function
    End If
    'd) Grava registros de descontos dos itens.
    If mfCopProposta("P0042700", gTipP0042700, gEndP0042700, "P0042700", pAntNosNumero, _
                     lNovNosNumero, lNovReferencia, pMensagem, pReentrada) = False Then
        Exit Function
    End If
    'e) Grava registros de observações dos itens.
    If mfCopProposta("P0042300", gTipP0042300, gEndP0042300, "P0042300", pAntNosNumero, _
                     lNovNosNumero, lNovReferencia, pMensagem, pReentrada) = False Then
        Exit Function
    End If
    'f) Grava registros de parcelamento.
    If mfCopProposta("P0042800", gTipP0042800, gEndP0042800, "P0042800", pAntNosNumero, _
                     lNovNosNumero, lNovReferencia, pMensagem, pReentrada) = False Then
        Exit Function
    End If
    
    If mfCopProposta("P0042400", gTipP0042400, gEndP0042400, "P0042400", pAntNosNumero, _
                     lNovNosNumero, lNovReferencia, pMensagem, pReentrada) = False Then
        Exit Function
    End If
    
    If gTipUsuario = "1" Then
        If mfCopProposta("TAB_AJUSTE_PREMIO", gTipP0042200, gEndP0042200, "TAB_AJUSTE_PREMIO", pAntNosNumero, _
                     lNovNosNumero, lNovReferencia, pMensagem, pReentrada) = False Then
            Exit Function
        End If
    End If
    
    gfCopiaProposta = True
End Function

Private Function mfCopProposta(ByVal pIndBasDados As String, ByVal pTipBasDados As String, _
                               ByVal pEndBasDados As String, ByVal pIdeTabela As String, _
                               ByVal pAntNosNumero As String, ByVal pNovNosNumero As String, _
                               ByVal pNovReferencia As String, ByRef pMensagem As String, _
                               pReentrada As Integer) As Boolean
    'Função: monta comando INSERT genérico para dados de proposta.

    'Parâmetros de entrada...pIndBasDados.....Indicador da base de dados.
    '                        pTipBasDados.....Tipo da base de dados.
    '                        pEndBasDados.....Endereço da base de dados.
    '                        pIdeTabela.......Identificação da tabela.
    '                        pAntNosNumero....Anterior: nosso número.
    '                        pNovNosNumero....Novo: nosso número.
    '                        pNovReferencia...Novo: número de referência.
    'Parâmetro de saída......pMensagem........Mensagem.

    Dim bdBasDados  As ADODB.Connection 'Base de dados.
    Dim flTabela    As ADODB.Field      'Campos da tabela.
    Dim rsTabela    As ADODB.Recordset  'Registro da tabela.
    Dim lSql        As String
    mfCopProposta = False

    '1.Abre base de dados.
    If gfAbrBasDados(pIndBasDados, pTipBasDados, pEndBasDados, bdBasDados, pMensagem) = False Then
        Exit Function
    End If

    '2. Obtém registros da tabela.
    If pIdeTabela = "TAB_AJUSTE_PREMIO" Then
        lSql = "SELECT * FROM " & gPropTabela & pIdeTabela & " WHERE Num_Proposta = '" & pAntNosNumero & "'"
    Else
        lSql = "SELECT * FROM " & gPropTabela & pIdeTabela & " WHERE Nosso_Numero = '" & pAntNosNumero & "'"
    End If
    If gfObtRegistro(bdBasDados, lSql, rsTabela, pMensagem) = False Then
        Call gpFecha2(bdBasDados)
        Exit Function
    End If

    '3. Grava nova cotação/proposta.
    While Not rsTabela.EOF
        gInsert = "INSERT INTO " & gPropTabela & pIndBasDados & " ("
        For Each flTabela In rsTabela.Fields
            gInsert = gInsert & flTabela.Name & ", "
        Next flTabela
        gInsert = Left$(gInsert, Len(gInsert) - 2) & ") VALUES ("
        For Each flTabela In rsTabela.Fields
            Select Case UCase(flTabela.Name)
                Case "NOSSO_NUMERO", "NUM_PROPOSTA"
                    gInsert = gInsert & "'" & pNovNosNumero & "', "                 'Nosso_Numero.
                Case "NUM_REFERENCIA"
                    gInsert = gInsert & pNovReferencia & ", "                       'Num_Referencia.
                Case "SIT_DOCUMENTO"
                    gInsert = gInsert & "0, "                                       'Sit_Documento.
                Case "CUS_EMISSAO"
                    gInsert = gInsert & "0, "                                       'CUS_EMISSAO.
                Case "DAT_ALT"
                    gInsert = gInsert & "0, "                                       'Dat_Alt.
                Case "COD_USER_ALT"
                    gInsert = gInsert & "'COPIA', "                                 'Cod_User_Alt.
                Case "SIT_ITEM"
                    gInsert = gInsert & "1, "                                       'SIT_ITEM
                Case "PAG_ANT_BLOQUETO"
                    If pReentrada = 1 Then
                        gInsert = gInsert & flTabela.Value & ", "                   'PAG_ANT_BLOQUETO
                    Else
                        gInsert = gInsert & "0, "                                    'PAG_ANT_BLOQUETO
                    End If
                Case "PAG_ANT_CHQ_BANCO"
                    gInsert = gInsert & "0, "                                       'Pag_Ant_Chq_Banco
                Case "QTD_RENOVACOES"
                    If gTipUsuario = "2" Or gTipUsuario = "3" Then
                        gInsert = gInsert & "0, "                                   'Pag_Ant_Chq_Banco
                    Else
                        gInsert = gInsert & flTabela.Value & ", "                   'Pag_Ant_Chq_Banco
                    End If
                 Case "COD_PRODUTO"
                    If gCodProduto <> 0 Then
                        gInsert = gInsert & gCodProduto & ", "
                    Else
                        gInsert = gInsert & flTabela.Value & ", "
                    End If
                Case "COD_PROTOCOLO"
                    If gCodProduto <> 0 Then
                        gInsert = gInsert & gCodProtocolo & ", "
                    Else
                        gInsert = gInsert & flTabela.Value & ", "
                    End If
                Case "COD_SEQUENCIA"
                    If gCodProduto <> 0 Then
                        gInsert = gInsert & gCodSequencia & ", "
                    Else
                        gInsert = gInsert & flTabela.Value & ", "
                    End If
                Case "DAT_EMISSAO"
                    If pIndBasDados = "P0042100" Then
                        gInsert = gInsert & "0, "
                    Else
                        gInsert = gInsert & flTabela.Value & ", "
                    End If
                Case "NOSSO_NUMERO_ORIGINAL"
                    gInsert = gInsert & "'', "
                Case "IND_REENTRADA"
                    gInsert = gInsert & "0, "
                Case Null
                    gInsert = gInsert & "0, "
                Case Else
                    Select Case flTabela.Type
                        Case adBigInt, adBinary, adInteger, adLongVarBinary, adNumeric, adSmallInt, _
                             adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, _
                             adUnsignedTinyInt                          'Numérico - inteiro.
                            If IsNull(flTabela.Value) Then
                                gInsert = gInsert & "0, "
                            Else
                                gInsert = gInsert & flTabela.Value & ", "
                            End If
                        Case adCurrency, adSingle, adDouble             'Numérico - com vírgula.
                            If IsNull(flTabela.Value) Then
                                gInsert = gInsert & "0, "
                            Else
                                gInsert = gInsert & gfForValor(1, flTabela.Value) & ", "
                            End If
                        Case adChar, adVarChar                          'Texto.
                            If gfPreenchido(flTabela.Value) Then
                                gInsert = gInsert & "'" & gfForAspas(flTabela.Value) & "', "
                            Else
                                gInsert = gInsert & "' ', "
                            End If
                        Case adDate                                     'Data.
                            If IsNull(flTabela.Value) Then
                                gInsert = gInsert & "0, "
                            Else
                                gInsert = gInsert & gfForData(1, flTabela.Value) & ", "
                            End If
                        Case adBoolean                                  'Sim/Não.
                            If IsNull(flTabela.Value) Then
                                gInsert = gInsert & "NULL, "
                            Else
                                gInsert = gInsert & flTabela.Value & ", "
                            End If
                        Case Else
                            Call gpFecha3(rsTabela, bdBasDados)
                            Call gpGraLog(2, "FTE0057 - Tabela " & pIdeTabela & " com tipo do campo " & _
                                             flTabela.Name & " diferente dos previstos.")
                            'End
                    End Select
            End Select
        Next flTabela
        gInsert = Left$(gInsert, Len(gInsert) - 2) & ")"
        If gfExeSQL(bdBasDados, gInsert, pMensagem) <> 0 Then
            Call gpFecha3(rsTabela, bdBasDados)
            Exit Function
        End If
        rsTabela.MoveNext
    Wend
    Call gpFecha3(rsTabela, bdBasDados)

    mfCopProposta = True
End Function

''--------------
Public Function gfCalReferencia(ByRef pReferencia As String) As Boolean
    'Função: calcula referência.

    'Parâmetro de saída...pReferencia...Número de referência.
    
    Dim objP0042500                     As clsA28V711A      'Objeto da classe clsA28V711A (P0042500)
    Dim lAuxiliar                       As Long             'Auxiliar.
    Dim lMensagem                       As String           'Mensagem.
    Dim objstcP0042500                  As stcA28V711B   'Objeto da classe de estrutura stcA28V711B (P0042500)

    
    gfCalReferencia = False

    '1. Instancio os objetos.
    Set objP0042500 = New clsA28V711A
    Set objstcP0042500 = New stcA28V711B
    
    '2. Informando Ult_Num_Referencia = -1 pois a query de consulta está preparada
    'para que se ela receber -1, busque na tabela o registro com valor menor que 1000000
    objstcP0042500.Ult_Num_Referencia = -1
    
    '3. Chamo a função que irá carregar a coleção com a referencia menor que 1000000
    If Not objP0042500.mfCarregarColecao(objstcP0042500) Then
        lMensagem = objP0042500.UltimoErro
        Call gpGraLog(2, lMensagem)
        gFinalizarPrograma = True
        Exit Function
    End If
    
    '3. Calcula referência e atualiza arquivo.
    If objP0042500.Colecao.Count = 0 Then
        lAuxiliar = 1
        objstcP0042500.Ult_Num_Referencia = lAuxiliar
        objstcP0042500.Dat_Alt = gfForData(3)
        objstcP0042500.Cod_User_Alt = gNomUsuario8
        'Não existe o registro, então eu incluo.
        If Not objP0042500.mfIncluir(objstcP0042500) Then
            lMensagem = objP0042500.UltimoErro
            Call gpGraLog(1, lMensagem)
            GoTo Finalizar
        End If
    Else
        'Carrego o stc com o item da coleção (só irá existir um)
        Set objstcP0042500 = objP0042500.Colecao.Item(1)
        lAuxiliar = objstcP0042500.Ult_Num_Referencia + 1
        If lAuxiliar > 999999 Then
            lAuxiliar = 1
        End If
        objstcP0042500.Ult_Num_Referencia = lAuxiliar
        objstcP0042500.Dat_Alt = gfForData(3)
        objstcP0042500.Cod_User_Alt = gNomUsuario8
        'Altero para a referência calculada.
        If Not objP0042500.mfAlterar(objstcP0042500) Then
            lMensagem = objP0042500.UltimoErro
            Call gpGraLog(1, lMensagem)
            GoTo Finalizar
        End If
    End If
    
Finalizar:
    Set objP0042500 = Nothing
    Set objstcP0042500 = Nothing
    
    'Confere a alteração do registro na base
    If gTipUsuario = 2 Then
        If mfConfirmaUpdateP00425(lAuxiliar) = False Then
            'Entra na rotina de reparação da base P0042500
                '*** OBS.:     Os comandos
                '        Set objP0042500 = Nothing
                '        Set objstcP0042500 = Nothing
                'Precisam ser executados antes de chamar a procedure abaixo
            Call mpSubstituirP00425(lAuxiliar)
        End If
    End If
    
    pReferencia = Format$(lAuxiliar, "000000")
    gfCalReferencia = True
    
End Function

Private Function mfConfirmaUpdateP00425(ByVal pNovaRef As Long) As Boolean
Dim lMsgErr             As String
Dim objP0042500         As clsA28V711A      'Objeto da classe clsA28V711A (P0042500)
Dim lAuxiliar           As Long             'Auxiliar.
Dim objstcP0042500      As stcA28V711B   'Objeto da classe de estrutura stcA28V711B (P0042500)

Err = 0
On Error GoTo TRATAMENTO_ERRO
    mfConfirmaUpdateP00425 = False

    '1. Instancio os objetos.
    Set objP0042500 = New clsA28V711A
    Set objstcP0042500 = New stcA28V711B
    
    '2. Informando Ult_Num_Referencia = -1 pois a query de consulta está preparada
    'para que se ela receber -1, busque na tabela o registro com valor menor que 1000000
    objstcP0042500.Ult_Num_Referencia = -1
    
    '3. Chamo a função que irá carregar a coleção com a referencia menor que 1000000
    If Not objP0042500.mfCarregarColecao(objstcP0042500) Then
        lMsgErr = objP0042500.UltimoErro
        GoTo TRATAMENTO_ERRO
    End If
    
    '4. obtem a referencia atual.
    If objP0042500.Colecao.Count = 0 Then
        lAuxiliar = 0
    Else
        'Carrego o stc com o item da coleção (só irá existir um)
        Set objstcP0042500 = objP0042500.Colecao.Item(1)
        lAuxiliar = objstcP0042500.Ult_Num_Referencia
    End If
    
    '5. Faz a comparação
    If lAuxiliar = 999999 And pNovaRef = 1 Then
        mfConfirmaUpdateP00425 = False
    Else
        If pNovaRef <= lAuxiliar Then
            mfConfirmaUpdateP00425 = True
        Else
            mfConfirmaUpdateP00425 = False
        End If
    End If
    
    
SAIDA:
    Set objP0042500 = Nothing
    Set objstcP0042500 = Nothing
    Err = 0
    On Error GoTo 0
    Exit Function
    
TRATAMENTO_ERRO:
    If gfPreenchido(lMsgErr) = False Then
        lMsgErr = Err & " - " & Error
    End If
    lMsgErr = "mfConfirmaUpdateP00425 Erro " & lMsgErr
    Call gpGraLog(2, lMsgErr)
    GoTo SAIDA
    
End Function


Private Sub mpSubstituirP00425(ByVal pNovoNumRef As Long)
Dim lMsgErr             As String
Dim objP0042500         As clsA28V711A      'Objeto da classe clsA28V711A (P0042500)
Dim objstcP0042500      As stcA28V711B   'Objeto da classe de estrutura stcA28V711B (P0042500)
Dim lEndP0042500_LDB    As String
Err = 0
On Error GoTo TRATAMENTO_ERRO
    'VERIFICA SE A BASE ESTA ABERTA
    lEndP0042500_LDB = LCase(gEndP0042500)
    lEndP0042500_LDB = Replace(lEndP0042500_LDB, ".mdb", ".ldb")

    If gfPreenchido(Dir(lEndP0042500_LDB)) = True Then
        'Base esta aberta, não é possivel substituir
        GoTo SAIDA
    End If
    
'Copia a nova base
    If gfPreenchido(Dir(gAppPath & "P0042500.NEW")) = True Then
        'Sobrescreve a base atual, por uma base nova
        Call FileCopy(gAppPath & "P0042500.NEW", gEndP0042500)
    Else
        GoTo SAIDA
    End If
    

'--- Atualiza com o numero novo
    '1. Instancio os objetos.
    Set objP0042500 = New clsA28V711A
    Set objstcP0042500 = New stcA28V711B
    
    '2. Informando Ult_Num_Referencia = -1 pois a query de consulta está preparada
    'para que se ela receber -1, busque na tabela o registro com valor menor que 1000000
    objstcP0042500.Ult_Num_Referencia = -1
    
    '3. Chamo a função que irá carregar a coleção com a referencia menor que 1000000
    If Not objP0042500.mfCarregarColecao(objstcP0042500) Then
        lMsgErr = objP0042500.UltimoErro
        GoTo TRATAMENTO_ERRO
    End If
    
    '3. Calcula referência e atualiza arquivo.
    If objP0042500.Colecao.Count = 0 Then
        objstcP0042500.Ult_Num_Referencia = pNovoNumRef
        objstcP0042500.Dat_Alt = gfForData(3)
        objstcP0042500.Cod_User_Alt = gNomUsuario8
        'Não existe o registro, então eu incluo.
        If Not objP0042500.mfIncluir(objstcP0042500) Then
            lMsgErr = objP0042500.UltimoErro
            GoTo TRATAMENTO_ERRO
        End If
    Else
        'Carrego o stc com o item da coleção (só irá existir um)
        Set objstcP0042500 = objP0042500.Colecao.Item(1)
        objstcP0042500.Ult_Num_Referencia = pNovoNumRef
        objstcP0042500.Dat_Alt = gfForData(3)
        objstcP0042500.Cod_User_Alt = gNomUsuario8
        'Altero para a referência calculada.
        If Not objP0042500.mfAlterar(objstcP0042500) Then
            lMsgErr = objP0042500.UltimoErro
            GoTo TRATAMENTO_ERRO
        End If
    End If

SAIDA:
    Set objP0042500 = Nothing
    Set objstcP0042500 = Nothing
    Err = 0
    On Error GoTo 0
    Exit Sub
    
TRATAMENTO_ERRO:
    If gfPreenchido(lMsgErr) = False Then
        lMsgErr = Err & " - " & Error
    End If
    lMsgErr = "mpSubstituirP00425 Erro " & lMsgErr
    Call gpGraLog(2, lMsgErr)
    GoTo SAIDA
    
End Sub

