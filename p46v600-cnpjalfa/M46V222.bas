Attribute VB_Name = "basM46V222"
Option Explicit

Public gConsisteComum    As Integer 'Mensagens de consistência do Comum
    '1  - 'PROCESSO INTERNO SEM DADOS DE COMUM'
    '2  - 'PROCESSO INTERNO INEXISTENTE'
    '3  - 'PROCESSO INTERNO EMITIDO'
    '4  - 'PROCESSO INTERNO CANCELADO'
    '5  - 'APOLICE ANTERIOR DIFERE DA TRIADA
    '6  - 'APOLICE JA EMITIDA
    '7  - 'CODIGO DO RAMO NAO CONFERE COM A APOLICE
    '8  - 'COD SEGURADO NAO PODE SER SER ALTERNATIVO
    '9  - 'CODIGO DE REGISTRO NAO PERMITE CUSTO
    '10 - 'OBRIGATORIO CUSTO PARCELADO
    '11 - 'OBRIGATORIO CUSTO ZERADO
    '12 - 'TERMINO INFERIOR A INICIO DE VIGENCIA
    '13 - 'TERMINO IGUAL AO INICIO DE VIGENCIA
    '14 - 'REFERENCIA NÃO PREENCHIDA
    '15 - 'DIGITAR CORRETOR-1 MARCH E CORRETOR-2 NOMMINAL'
    '16 - 'PARA CORRETOR NOMMINAL DEVE EXISTIR CORRETOR MARCH
    '17 - 'DEVE SER DIGITADO O CODIGO DA SEGURATEC'
    '18 - 'PARA BASCOL, DEVE SER DIGITADO O CODIGO DA SEGURATEC
    '19 - 'PARA PGTO.ANTEC. A VISTA, JURO DEVE SER ZEROS'
    '20 - 'CORRETOR FOI DIGITADO MAIS DE UMA VEZ'
    '21 - 'QDO CORRETOR FOR PTMM EH OBRIGATORIO DIG TMM '
    '22 - 'QDO CORRETOR FOR TMM EH OBRIGATORIO DIG PTMM '
    '23 - '" YASUDA  " NAO PODE SER COSSEGURADORA'
    '24 - 'COSSEGURADORA 955 PERMITIDO SOMENTE PARA RAMO VIDA'
    '25 - 'NAO E MAIS PERMITIDO COSSEGURO SEASUL'
    '26 - 'TIPO DE ENVIO DEVE SER "S" PARA SEGURADO E "C" PARA CORRETOR'
    '27 - 'TIPO DE COTACAO DEVE SER 0, 1, 2 OU 3 '
    '28 - 'NUMERO DA COTACAO DEVE SER NUMERICO '
    '29 - 'CORRETAGEM VVD/PALLAS,DIGITAR PRIMEIRO PALLAS'
    '30 - 'QUANDO FOR CORRETOR PALLAS UTILIZAR TAMBEM VVD'
    '31 - 'LOCALIDADE DA PALLAS E VVD DEVE SER IGUAIS'
    '32 - 'VVD DEVE TER 70 % DA COMISSAO'
    '33 - 'PALLAS DEVE TER 30 % DA COMISSAO'
    '34 - 'PAGAMENTO ANTECIPADO NAO CADASTRADO'
    '35 - 'FOI INFORMADO DEBITO DA 1A PARCELA, POREM EXISTE PAGAMENTO ANTECIPADO CADASTRADO'
    '36 - 'PREMIO ANTECIPADO DEVE SER MAIOR QUE CUSTO + IOF'
    '37 - 'PARA ESTE CODIGO DE REGISTRO JURO NAO DEVE SER PREENCHIDO'

                                    
                                    
Public gConsisteItem    As Integer 'Mensagens de consistência dos Itens
    '1  - 'PROCESSO INTERNO SEM ITENS CADASTRADOS'
    '2  - 'ITEM DA APOLICE JA EMITIDO
    '3  - 'INSPECAO 9999999 INVALIDA'
    '4  - 'CODIGO DA CLASSIFICACAO NAO ESTA PREENCHIDO'
    '5  - 'COMPLEMENTO DA CLASSIFICACAO NAO ESTA PREENCHIDO'
    '6  - 'PERCENTUAL DE BONUS INVALIDO PARA TELEMAR'
    '7  - 'PERCENTUAL DE DESCONTO PROGRESSIVO INVALIDO PARA TELEMAR'
    '8  - 'CODIGO DE CONSTRUCAO INVALIDA PARA ESTA CLASSIFICACAO
    '9  - 'PERCENTUAL DESC. COMERCIAL DA COBERTURA 9999 INVALIDO PARA TELEMAR'
    '10 - 'QUANTIDADE DE EMPREGADOS DEVE SER ZEROS'
    '11 - 'INFORME A QUANTIDADE DE EMPREGADOS
    '12 - 'FALTA CLAUSULA 0001 OU 0002'
    '13 - 'DIGITE SOMENTE CLAUSULA 0001 OU 0002'
    '14 - 'PARA CLASSIFICACAO 197 E 134, A CLAUSULA DEVE SER 0001'
 
Private Type ty_Parm
    Cod_User    As String * 6
    Cod_Ramo As Integer
    NUM_PI  As Long
    Num_Ped  As Long
    Num_Item As Long
    Num_Apol_Ant  As Double
    Num_Item_Ant  As Long
    Tip_Proc As Long
    flag_debug As Integer
    Cod_Dizer As Integer
    Versao_Corr As Integer
    Cod_Retorno As Integer
    Cod_Err_Nor_1 As Long
    Mens_Err_1 As String * 80
    Cod_Err_Nor_2 As Long
    Mens_Err_2 As String * 80
    Cod_Err_Nor_3 As Long
    Mens_Err_3 As String * 80
    Cod_Err_Nor_4 As Long
    Mens_Err_4 As String * 80
    Cod_Err_Nor_5 As Long
    Mens_Err_5 As String * 80
    Cod_Err_Nor_6 As Long
    Mens_Err_6 As String * 80
    Cod_Err_Nor_7 As Long
    Mens_Err_7 As String * 80
    Cod_Err_Nor_8 As Long
    Mens_Err_8 As String * 80
    Cod_Err_Nor_9 As Long
    Mens_Err_9 As String * 80
    Cod_Err_Nor_10 As Long
    Mens_Err_10 As String * 80
    Cod_Err_Nor_11 As Long
    Mens_Err_11 As String * 80
    Cod_Err_Nor_12 As Long
    Mens_Err_12 As String * 80
    Cod_Err_Nor_13 As Long
    Mens_Err_13 As String * 80
    Cod_Err_Nor_14 As Long
    Mens_Err_14 As String * 80
    Cod_Err_Nor_15 As Long
    Mens_Err_15 As String * 80
    Cod_Err_Nor_16 As Long
    Mens_Err_16 As String * 80
    Cod_Err_Nor_17 As Long
    Mens_Err_17 As String * 80
    Cod_Err_Nor_18 As Long
    Mens_Err_18 As String * 80
    Cod_Err_Nor_19 As Long
    Mens_Err_19 As String * 80
    Cod_Err_Nor_20 As Long
    Mens_Err_20 As String * 80
 End Type
 
 
Public Function gfPesqDadosComum(ByRef pobjstcPed As stcA46V702B, ByRef pobjstcPedLoc As stcA46V708B, _
                        ByVal pTipoProc As e_TipProc, ByVal pItem As Long, _
                        ByRef pCod_Retorno As Integer, pTipHomologacao As String) As String
    ' 1 - para consistir comum
    ' 2 - para consistir o item
    ' 3 - para normas

Dim bdP00APOL              As ADODB.Connection 'Base de dados: P00Apol.
Dim rsP00APOL              As clsYasRecordSet  'Registro: P00Apol.

Dim lParm                  As ty_Parm
Dim lretorno               As String
Dim lNomeArq               As String
Dim lMensagem              As String
Dim lAchouComum            As Boolean
Dim lColigada              As Boolean
Dim lSeasul                As Boolean
Dim lPerCorrPallas         As Double
Dim lPerCorrVVD            As Double
Dim lPerCorrPallas_calc    As Double
Dim lPerCorrVVD_calc       As Double
Dim lCodCorrPallas         As Long
Dim lCodCorrVVD            As Long
Dim lSelect     As String
    On Error GoTo ControleErro
    gfPesqDadosComum = False

    'vou criar um arquivo somente para validar se o usuário é único no processamento.
    If mpVerCalculoParalelo() Then
        Call gM46V999.gpGraLog(1, lMensagem)
    End If

    'Verifica Comum
    lAchouComum = False
    
    lParm.Cod_Ramo = pobjstcPed.Cod_Ramo
    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        lParm.Cod_User = "SYAS98"
    Else
        lParm.Cod_User = "SYAS99"
    End If
    lParm.NUM_PI = gM46V111.gNum_Pi
    lParm.Num_Ped = pobjstcPed.Num_Ped
    lParm.Num_Item = pItem
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Or pobjstcPedLoc.Tip_Emissao = e_TipEmissao_Endosso Then
        lParm.Num_Apol_Ant = pobjstcPedLoc.Num_Apol_Anterior
        lParm.Num_Item_Ant = val(pItem)
    Else
        lParm.Num_Apol_Ant = 0
        lParm.Num_Item_Ant = 0
    End If
    lParm.flag_debug = 1
    lParm.flag_debug = 0
    lParm.Tip_Proc = pTipoProc
    'se pTipoProc = 3, = Vernormas - Passar Tarifa no número do PI
    If pTipoProc = 3 Then
        lParm.NUM_PI = pobjstcPed.Cod_Produto
    End If
    lParm.Cod_Dizer = pobjstcPed.Cod_Dizer
    If pTipHomologacao = "1" Then
        lParm.Versao_Corr = 2   'versão homologação
    Else
        lParm.Versao_Corr = 1
    End If

   '============================================================================================'
   ' TAB_PED
   '============================================================================================'

    'VERIFICAR TAB_APOL / RAMOSDIVERSOS..TAB_PED
    If gM46V999.gBaseLocal = True Then
        lNomeArq = "TAB_APOL"
    Else
        lNomeArq = "RAMOSDIVERSOS.DBO.TAB_PED"
    End If
   
    '1. Obtém registro da apólice.
    'a) Monta comando SELECT.
    lSelect = "SELECT * FROM " & lNomeArq
    lSelect = lSelect & " Where Num_Ped  =  " & lParm.Num_Ped
    lSelect = lSelect & " AND   Cod_Ramo =  " & lParm.Cod_Ramo
    lSelect = lSelect & " AND   Cod_Org  = 0 AND  Num_Ped <> 0"
    
    'b) Abre arquivo.
    If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, bdP00APOL, lMensagem) = False Then
        Exit Function
    End If
    'c) Obtém registro.
    If gM46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, lMensagem) = False Then
        Call gM46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If
    If rsP00APOL.EOF = True Then
        lAchouComum = False
        lMensagem = "Apólice " & lParm.Num_Apol_Ant & " não localizada."
    Else
        lAchouComum = True
    End If
    
    '--------------------APOLICE----------------------------------------
    If pobjstcPed.Num_Apol_An <> lParm.Num_Apol_Ant Then
        'APOLICE ANTERIOR DIFERE DA TRIADA
        pCod_Retorno = 5
        Exit Function
    End If
    If lAchouComum = False Then
        'PROCESSO INTERNO SEM DADOS DE COMUM'
        pCod_Retorno = 1
        Exit Function
    End If
    If rsP00APOL("cod_sit_ped") = 5 Then
        'APOLICE JA EMITIDA
        pCod_Retorno = 6
        Exit Function
    End If
    
    '--------------------SEGURADO----------------------------------------
    If Left(rsP00APOL("cod_segurado"), 2) = 99 Then
        'COD SEGURADO NAO PODE SER SER ALTERNATIVO
        pCod_Retorno = 8
        Exit Function
    End If
    If Left(rsP00APOL("cod_segurado"), 4) = 7777 And Right(rsP00APOL("cod_segurado"), 2) = 98 Then
        lColigada = True
    Else
        lColigada = False
    End If
    
    '--------------------PAGAMENTO ANTECIPADO----------------------------
    If rsP00APOL("val_pr_antec") > 0 And (rsP00APOL("tip_cust_emis") <> 4 And rsP00APOL("tip_cust_emis") <> 5) Then
        'PAGAMENTO ANTECIPADO NAO CADASTRADO'
        pCod_Retorno = 34
        Exit Function
    End If
'    If pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim And (rsP00APOL("tip_cust_emis") = 2 Or rsP00APOL("tip_cust_emis") = 3) Then
'        'FOI INFORMADO DEBITO DA 1A PARCELA, POREM EXISTE PAGAMENTO ANTECIPADO CADASTRADO'
'        pCod_Retorno = 35
'        Exit Function
'    End If

   
    '--------------------CUSTO----------------------------------------
    If rsP00APOL("val_cust_emis") > 0 Then
        If rsP00APOL("cod_reg") <> 1 And rsP00APOL("cod_reg") <> 3 And rsP00APOL("cod_reg") <> 101 And rsP00APOL("cod_reg") <> 103 And rsP00APOL("cod_reg") <> 20 Then
            'CODIGO DE REGISTRO NAO PERMITE CUSTO
            pCod_Retorno = 9
            Exit Function
        End If
    End If
    If rsP00APOL("COD_RAMO") = 113 Then
        If (rsP00APOL("cod_prod") = 1001 Or rsP00APOL("cod_prod") = 1002 Or rsP00APOL("cod_prod") = 1003 Or _
            rsP00APOL("cod_prod") = 1004 Or rsP00APOL("cod_prod") = 1005 Or rsP00APOL("cod_prod") = 1006 Or _
            rsP00APOL("cod_prod") = 1007 Or rsP00APOL("cod_prod") = 1008 Or rsP00APOL("cod_prod") = 1009 Or _
            rsP00APOL("cod_prod") = 1010 Or rsP00APOL("cod_prod") = 1011 Or rsP00APOL("cod_prod") = 1012 Or _
            rsP00APOL("cod_prod") = 1013 Or rsP00APOL("cod_prod") = 1014 Or rsP00APOL("cod_prod") = 1015 Or _
            rsP00APOL("cod_prod") = 1016 Or rsP00APOL("cod_prod") = 1017 Or rsP00APOL("cod_prod") = 1018 Or _
            rsP00APOL("cod_prod") = 1019 Or rsP00APOL("cod_prod") = 1020) Then
                If rsP00APOL("cod_tip_doc") = e_TipEmissao_SeguroNovo Or rsP00APOL("cod_tip_doc") = e_TipEmissao_RenYasuda Then
                        If rsP00APOL("tip_cust_emis") <> 1 And rsP00APOL("tip_cust_emis") <> 3 And rsP00APOL("tip_cust_emis") <> 5 Then
                        'Para este produto o tipo de custo deve ser parcelado.................'
                        'OBRIGATORIO CUSTO PARCELADO
                            pCod_Retorno = 10
                            Exit Function
                        End If
                End If
        End If
    End If
    If rsP00APOL("qtd_parc") = 0 Then
        If rsP00APOL("tip_cust_emis") <> 0 And rsP00APOL("tip_cust_emis") <> 2 And rsP00APOL("tip_cust_emis") <> 4 Then
            'QUANDO PARCELAMENTO = ZEROS, TIPO DE CUSTO DEVE SER ZEROS....................'
            'OBRIGATORIO CUSTO ZERADO
            pCod_Retorno = 11
            Exit Function
        End If
    End If
    
    '--------------------VIGÊNCIA----------------------------------------
    If rsP00APOL("dat_ter_vig") < rsP00APOL("dat_ini_vig") Then
        'TERMINO INFERIOR AO INICIO DE VIGENCIA'
        pCod_Retorno = 12
        Exit Function
    End If
    If rsP00APOL("dat_ter_vig") = rsP00APOL("dat_ini_vig") Then
        'TERMINO IGUAL AO INICIO DE VIGENCIA'
        pCod_Retorno = 13
        Exit Function
    End If
    
    '--------------------JUROS----------------------------------------
    If rsP00APOL("qtd_parc") = 1 And rsP00APOL("val_pr_antec") > 0 Then
        If rsP00APOL("per_juro_frac") > 0 Then
            'PARA PGTO.ANTEC. A VISTA, JURO DEVE SER ZEROS'
            pCod_Retorno = 19
            Exit Function
        End If
    End If
    If rsP00APOL("cod_reg") = 5 And rsP00APOL("qtd_parc") = 1 Then
        If rsP00APOL("per_juro_frac") > 0 Then
            'PARA ESTE CODIGO DE REGISTRO JURO NAO DEVE SER PREENCHIDO'
            pCod_Retorno = 37
            Exit Function
        End If
    End If
    If rsP00APOL("cod_reg") = 105 And (rsP00APOL("qtd_parc") = 1 Or rsP00APOL("qtd_parc") = 31) Then
        If rsP00APOL("per_juro_frac") > 0 Then
            'PARA ESTE CODIGO DE REGISTRO JURO NAO DEVE SER PREENCHIDO'
            pCod_Retorno = 37
            Exit Function
        End If
    End If
    If rsP00APOL("cod_reg") = 20 Or rsP00APOL("cod_reg") = 106 Or rsP00APOL("cod_reg") = 107 Or rsP00APOL("cod_reg") = 120 Or _
       rsP00APOL("cod_reg") = 121 Or rsP00APOL("cod_reg") = 122 Or rsP00APOL("cod_reg") = 129 Then
       If rsP00APOL("per_juro_frac") > 0 Then
            'PARA ESTE CODIGO DE REGISTRO JURO NAO DEVE SER PREENCHIDO'
            pCod_Retorno = 37
            Exit Function
       End If
    End If

    '--------------------VALOR----------------------------------------
'????????????????????????????????????
    
   '============================================================================================'
   ' TAB_CTRL_EMIS
   '============================================================================================'

    'VERIFICAR TAB_APOL / RAMOSDIVERSOS..TAB_CTRL_EMIS
    If gM46V999.gBaseLocal = True Then
        lNomeArq = "TAB_CTRL_EMIS"
    Else
        lNomeArq = "RAMOSDIVERSOS.DBO.TAB_CTRL_EMIS"
    End If
   
    '1. Obtém registro da apólice.
    'a) Monta comando SELECT.
    lSelect = "SELECT * FROM " & lNomeArq
    lSelect = lSelect & " Where Num_PI  =  " & lParm.NUM_PI
    
    'b) Abre arquivo.
    If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, bdP00APOL, lMensagem) = False Then
        Exit Function
    End If
    'c) Obtém registro.
    If gM46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, lMensagem) = False Then
        Call gM46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If
    
    
    '--------------------PI----------------------------------------
    If rsP00APOL.EOF = True Then
        'PROCESSO INTERNO INEXISTENTE
        pCod_Retorno = 2
        Exit Function
    Else
        pCod_Retorno = 0
    End If
    
    '--------------------REFERÊNCIA----------------------------------------
    If gM46V999.gfPreenchido(rsP00APOL("Num_Ref_Corr")) = False Then
        'REFERENCIA NÃO PREENCHIDA
        pCod_Retorno = 14
        Exit Function
    End If
    If gM46V999.gfPreenchido(rsP00APOL("Cod_User_Digit")) And rsP00APOL("Cod_User_Digit") <> "SYAS98" Then
        If Left(rsP00APOL("Num_Ref_Corr"), 1) = " " Then
            'DIGITE A REFERENCIA DO CORRETOR A PARTIR DA 1A POSICAO
            pCod_Retorno = 14
            Exit Function
        End If
    End If

    
    'VERIFICAR TAB_APOL / GED..TAB_CTRL_EMIS
    If gM46V999.gBaseLocal = True Then
        lNomeArq = "TAB_CTRL_EMIS"
    Else
        lNomeArq = "GED.DBO.TAB_CTRL_EMIS"
    End If
   
    '1. Obtém registro da apólice.
    'a) Monta comando SELECT.
    lSelect = "SELECT * FROM " & lNomeArq
    lSelect = lSelect & " Where Num_PI  =  " & lParm.NUM_PI
    
    'b) Abre arquivo.
    If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, bdP00APOL, lMensagem) = False Then
        Exit Function
    End If
    'c) Obtém registro.
    If gM46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, lMensagem) = False Then
        Call gM46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If
    If Not rsP00APOL.EOF Then
     
     '--------------------PI----------------------------------------
        If rsP00APOL("cod_iden_stat_acao") = 90 Then
            'PROCESSO INTERNO EMITIDO
            pCod_Retorno = 3
            Exit Function
        End If
        If rsP00APOL("cod_iden_stat_acao") = 80 Or rsP00APOL("cod_iden_stat_acao") = 81 Then
            'PROCESSO INTERNO CANCELADO
            pCod_Retorno = 4
            Exit Function
        End If

    '--------------------RAMO----------------------------------------
        If gM46V999.gCod_Ramo = 113 And (Mid(rsP00APOL("num_apol_ant"), 3, 3) <> 118 And Mid(rsP00APOL("num_apol_ant"), 3, 3) <> 119) Then
            'CODIGO DO RAMO NAO CONFERE COM A APOLICE
            pCod_Retorno = 7
            Exit Function
        End If
        If gM46V999.gCod_Ramo = 112 And Mid(rsP00APOL("num_apol_ant"), 3, 3) <> 117 Then
            'CODIGO DO RAMO NAO CONFERE COM A APOLICE
            pCod_Retorno = 7
            Exit Function
        End If
        
    '--------------------TIPO DE ENVIO----------------------------------------
        If rsP00APOL("cod_envi_doc") <> "S" And rsP00APOL("cod_envi_doc") <> "C" And rsP00APOL("cod_envi_doc") <> "T" Then
            'TIPO DE ENVIO DEVE SER "S" PARA SEGURADO E "C" PARA CORRETOR'
            pCod_Retorno = 26
            Exit Function
        End If
    
    '--------------------COTACAO----------------------------------------
        If rsP00APOL("tip_cotac") <> "0" And rsP00APOL("tip_cotac") <> "1" And rsP00APOL("tip_cotac") <> "2" And rsP00APOL("tip_cotac") <> "3" Then
            'TIPO DE COTACAO DEVE SER 0, 1, 2 OU 3 '
            pCod_Retorno = 27
            Exit Function
        End If
        If Not IsNumeric(rsP00APOL("num_cotac")) Then
            'NUMERO DA COTACAO DEVE SER NUMERICO '
            pCod_Retorno = 28
            Exit Function
        End If
    
        
    '--------------------CORRETOR----------------------------------------
        If rsP00APOL("Cod_Corr") = 333501 Or rsP00APOL("Cod_Corr_3") = 333501 Then
            'DIGITAR CORRETOR-1 MARCH E CORRETOR-2 NOMMINAL'
            pCod_Retorno = 15
            Exit Function
        Else
            If rsP00APOL("Cod_Corr_2") = 333501 Then
                If rsP00APOL("Cod_Corr") = 3401 Or 646401 Then
                    'PARA CORRETOR NOMMINAL DEVE EXISTIR CORRETOR MARCH
                    pCod_Retorno = 16
                    Exit Function
            End If
        End If
        If gM46V999.gfPreenchido(rsP00APOL("Cod_User_Digit")) And Left(rsP00APOL("Cod_User_Digit"), 4) = "STEC" And rsP00APOL("Cod_Corr") <> 1601 Then
            'DEVE SER DIGITADO O CODIGO DA SEGURATEC'
            pCod_Retorno = 17
            Exit Function
        End If
        If lColigada = True And rsP00APOL("Cod_Corr") <> 1601 Then
            'PARA BASCOL, DEVE SER DIGITADO O CODIGO DA SEGURATEC
            pCod_Retorno = 18
            Exit Function
        End If
        If (Left(rsP00APOL("Cod_Corr"), 5) = Left(rsP00APOL("Cod_Corr_2"), 5)) And rsP00APOL("Cod_Corr") <> 0 Then
            'CORRETOR FOI DIGITADO MAIS DE UMA VEZ'
            pCod_Retorno = 20
            Exit Function
        End If
        If (Left(rsP00APOL("Cod_Corr_2"), 5) = Left(rsP00APOL("Cod_Corr_3"), 5)) And rsP00APOL("Cod_Corr_2") <> 0 Then
            'CORRETOR FOI DIGITADO MAIS DE UMA VEZ'
            pCod_Retorno = 20
            Exit Function
        End If
        If Left(rsP00APOL("Cod_Corr"), 5) = "03301" Or Left(rsP00APOL("Cod_Corr_2"), 5) = "03301" Or Left(rsP00APOL("Cod_Corr_3"), 5) = "03301" Then
            If Left(rsP00APOL("Cod_Corr"), 5) = "03302" Or Left(rsP00APOL("Cod_Corr_2"), 5) = "03302" Or Left(rsP00APOL("Cod_Corr_3"), 5) = "03302" Then
                'não faz nada
            Else
                'QDO CORRETOR FOR PTMM EH OBRIGATORIO DIG TMM '
                pCod_Retorno = 21
                Exit Function
            End If
        Else
            If Left(rsP00APOL("Cod_Corr"), 5) = "03302" Or Left(rsP00APOL("Cod_Corr_2"), 5) = "03302" Or Left(rsP00APOL("Cod_Corr_3"), 5) = "03302" Then
                'QDO CORRETOR FOR TMM EH OBRIGATORIO DIG PTMM '
                pCod_Retorno = 21
                Exit Function
            End If
        End If
        If Left(rsP00APOL("Cod_Corr_2"), 5) = "00058" Or Left(rsP00APOL("Cod_Corr_3"), 5) = "00058" Then
            If Left(rsP00APOL("Cod_Corr"), 5) = "00080" Then
                'CORRETAGEM VVD/PALLAS,DIGITAR PRIMEIRO PALLAS'
                pCod_Retorno = 29
                Exit Function
            End If
        End If
        
        If Left(rsP00APOL("Cod_Corr"), 5) = "00058" Then
            lCodCorrPallas = rsP00APOL("Cod_Corr")
            If rsP00APOL("Per_Corr") > 0 Then
                lPerCorrPallas = rsP00APOL("Per_Corr") / 100
            End If
        End If
        If Left(rsP00APOL("Cod_Corr_2"), 5) = "00058" Then
            lCodCorrPallas = rsP00APOL("Cod_Corr_2")
            If rsP00APOL("Per_Corr_2") > 0 Then
                lPerCorrPallas = rsP00APOL("Per_Corr_2") / 100
            End If
        End If
        If Left(rsP00APOL("Cod_Corr_3"), 5) = "00058" Then
            lCodCorrPallas = rsP00APOL("Cod_Corr_3")
            If rsP00APOL("Per_Corr_3") > 0 Then
                lPerCorrPallas = rsP00APOL("Per_Corr_3") / 100
            End If
        End If
        If Left(rsP00APOL("Cod_Corr"), 5) = "00080" Then
            lCodCorrVVD = rsP00APOL("Cod_Corr")
            If rsP00APOL("Per_Corr") > 0 Then
                lPerCorrVVD = rsP00APOL("Per_Corr") / 100
            End If
        End If
        If Left(rsP00APOL("Cod_Corr_2"), 5) = "00080" Then
            lCodCorrVVD = rsP00APOL("Cod_Corr_2")
            If rsP00APOL("Per_Corr_2") > 0 Then
                lPerCorrVVD = rsP00APOL("Per_Corr_2") / 100
            End If
        End If
        If Left(rsP00APOL("Cod_Corr_3"), 5) = "00080" Then
            lCodCorrVVD = rsP00APOL("Cod_Corr_3")
            If rsP00APOL("Per_Corr_3") > 0 Then
                lPerCorrVVD = rsP00APOL("Per_Corr_3") / 100
            End If
        End If
        If lCodCorrPallas = 0 Then
            GoTo Fim_Consiste_Pallas_VVD
        End If
        If lCodCorrVVD = 0 Then
            'QUANDO FOR CORRETOR PALLAS UTILIZAR TAMBEM VVD'
            pCod_Retorno = 31
            Exit Function
        End If
        If Mid(lCodCorrVVD, 6, 2) <> Mid(lCodCorrPallas, 6, 2) Then
            'LOCALIDADE DA PALLAS E VVD DEVE SER IGUAIS'
            pCod_Retorno = 31
            Exit Function
        End If
        If (lPerCorrVVD + lPerCorrPallas) = 0 Then
            GoTo Fim_Consiste_Pallas_VVD
        End If
        If lPerCorrVVD = "0.01" And lPerCorrPallas = "0.01" Then
            GoTo Fim_Consiste_Pallas_VVD
        End If
        
        lPerCorrVVD_calc = ((lPerCorrVVD * 100) / (lPerCorrVVD + lPerCorrPallas))

        If lPerCorrVVD_calc < 69 Or lPerCorrVVD_calc > 71 Then
            'VVD DEVE TER 70 % DA COMISSAO'
            pCod_Retorno = 32
            Exit Function
        End If
        If lPerCorrPallas_calc < 29 Or lPerCorrPallas_calc > 31 Then
            'PALLAS DEVE TER 30 % DA COMISSAO'
            pCod_Retorno = 33
            Exit Function
        End If
    End If
Fim_Consiste_Pallas_VVD:
 
    
    '--------------------PREMIOS-------------------------------------------
        Dim lValCustEmis    As Double
        
'        If rsP00APOL("cod_reg") <> 4 And 104 Then
'            If rsP00APOL("val_pr_antec") > 0 Then
'                If rsP00APOL("qtd_parc") > 1 Then
'                    If rsP00APOL("val_cust_emis") > 0 Then
'                        If rsP00APOL("tip_cust_emis") = 1 Or 3 Or 5 Then
'                            lValCustEmis = (rsP00APOL("val_cust_emis") / rsP00APOL("qtd_parc")) + rsP00APOL("val_iof")
'                        Else
'                            lValCustEmis = rsP00APOL("val_cust_emis") + rsP00APOL("val_iof")
'                        End If
'                        If rsP00APOL("val_pr_antec") <= lValCustEmis Then
'                            'PREMIO ANTECIPADO DEVE SER MAIOR QUE CUSTO + IOF'
'                            pCod_Retorno = e_ConsisteComum_PrAntecMenorCustoIOF
'                            Exit Function
'                        End If
'                    End If
'                End If
'            End If
'        End If
    End If
        
    '--------------------COSSEGURO----------------------------------------
    'VERIFICAR TAB_PED_COSS
    If gM46V999.gBaseLocal = True Then
        lNomeArq = "TAB_PED_COSS"
    Else
        lNomeArq = "GED.DBO.TAB_PED_COSS"
    End If

    '1. Obtém registro da apólice.
    'a) Monta comando SELECT.
    lSelect = "SELECT * FROM " & lNomeArq
    lSelect = lSelect & " Where Num_PI  =  " & lParm.NUM_PI

    'b) Abre arquivo.
    If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, bdP00APOL, lMensagem) = False Then
        Exit Function
    End If
    'c) Obtém registro.
    If gM46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, lMensagem) = False Then
        Call gM46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If
    lSeasul = False
    Do While Not rsP00APOL.EOF
        If rsP00APOL("Cod_Coss") = "641" Then
            '" YASUDA  " NAO PODE SER COSSEGURADORA'
            pCod_Retorno = 23
            Exit Function
        End If
        If rsP00APOL("Cod_Coss") = "955" Then
            'COSSEGURADORA 955 PERMITIDO SOMENTE PARA RAMO VIDA'
            pCod_Retorno = 24
            Exit Function
        End If
        If rsP00APOL("Cod_Coss") = "693" Then
            lSeasul = True
        End If

        If pobjstcPedLoc.Tip_Emissao <> e_TipEmissao_Endosso Then
            If Left(pobjstcPed.Cod_Segurado, 4) = 7777 And Right(pobjstcPed.Cod_Segurado, 2) <> 98 Then
                If lSeasul = True Then
                    'NAO E MAIS PERMITIDO COSSEGURO SEASUL'
                    pCod_Retorno = 25
                    Exit Function
                End If
            End If
        End If
        rsP00APOL.MoveNext
    Loop
    
    gfPesqDadosComum = True
    
    Call mpDeletarArqCalculo
    
    Exit Function
    
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "gfPesqDadosComum ERRO - " & Err & " " & Err.Description & " " & lMensagem)

End Function

Public Function gfPesqDadosItem(ByRef pobjstcPed As stcA46V702B, ByRef pobjstcPedLoc As stcA46V708B, _
                        ByVal pTipoProc As e_TipProc, ByVal pItem As Long, _
                        ByRef pCod_Retorno As Integer, pTipHomologacao As String) As String
    ' 1 - para consistir comum
    ' 2 - para consistir o item
    ' 3 - para normas

Dim bdP00APOL   As ADODB.Connection 'Base de dados: P00Apol.
Dim rsP00APOL   As clsYasRecordSet  'Registro: P00Apol.
Dim lSelect     As String
Dim lParm           As ty_Parm
Dim lretorno        As String
Dim lNomeArq        As String
Dim lMensagem       As String
Dim lAchouItem      As Boolean
Dim lAchouCobert    As Boolean
Dim lCodIdenTelemar As Boolean
Dim lPermiteEmpreg  As Boolean
Dim lQtdeEmpreg     As Long

    On Error GoTo ControleErro
    gfPesqDadosItem = False

    'vou criar um arquivo somente para validar se o usuário é único no processamento.
    If mpVerCalculoParalelo() Then
        Call gM46V999.gpGraLog(1, lMensagem)
    End If

    'Verifica Comum
    lAchouItem = False
    
    lParm.Cod_Ramo = pobjstcPed.Cod_Ramo
    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        lParm.Cod_User = "SYAS98"
    Else
        lParm.Cod_User = "SYAS99"
    End If
    lParm.NUM_PI = gM46V111.gNum_Pi
    lParm.Num_Ped = pobjstcPed.Num_Ped
    lParm.Num_Item = pItem
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Or pobjstcPedLoc.Tip_Emissao = e_TipEmissao_Endosso Then
        lParm.Num_Apol_Ant = pobjstcPedLoc.Num_Apol_Anterior
        lParm.Num_Item_Ant = val(pItem)
    Else
        lParm.Num_Apol_Ant = 0
        lParm.Num_Item_Ant = 0
    End If
    lParm.flag_debug = 1
    lParm.flag_debug = 0
    lParm.Tip_Proc = pTipoProc
    'se pTipoProc = 3, = Vernormas - Passar Tarifa no número do PI
    If pTipoProc = 3 Then
        lParm.NUM_PI = pobjstcPed.Cod_Produto
    End If
    lParm.Cod_Dizer = pobjstcPed.Cod_Dizer
    If pTipHomologacao = "1" Then
        lParm.Versao_Corr = 2   'versão homologação
    Else
        lParm.Versao_Corr = 1
    End If

    'VERIFICAR TAB_APOL / RAMOSDIVERSOS..TAB_PED_LOC
    If gM46V999.gBaseLocal = True Then
        lNomeArq = "TAB_APOL"
    Else
        lNomeArq = "RAMOSDIVERSOS.DBO.TAB_PED_LOC"
    End If
   
    '1. Obtém registro da apólice.
    'a) Monta comando SELECT.
    lSelect = "SELECT * FROM " & lNomeArq
    lSelect = lSelect & " Where Num_Ped  =  " & lParm.Num_Ped
    lSelect = lSelect & " AND   Cod_Ramo =  " & lParm.Cod_Ramo
    lSelect = lSelect & " AND   Cod_Org  = 0 AND  Num_Ped <> 0"
    
    'b) Abre arquivo.
    If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, bdP00APOL, lMensagem) = False Then
        Exit Function
    End If
    'c) Obtém registro.
    If gM46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, lMensagem) = False Then
        Call gM46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If
    If rsP00APOL.EOF = True Then
        lAchouItem = False
        lMensagem = "Itens da apólice " & lParm.Num_Apol_Ant & " não localizados."
    Else
        lAchouItem = True
    End If
    
    '--------------------PI----------------------------------------
    If lAchouItem = False Then
        'PROCESSO INTERNO SEM ITENS CADASTRADOS'
        pCod_Retorno = 1
        Exit Function
    End If

    '--------------------APOLICE----------------------------------------
    If lAchouItem = True And rsP00APOL("cod_sit_item") = 5 Then
        'ITEM DA APOLICE JA EMITIDO
        pCod_Retorno = 2
        Exit Function
    End If

    '--------------------INSPECAO----------------------------------------
    If rsP00APOL("num_insp") = 9999999 Then
        'INSPECAO 9999999 INVALIDA'
        pCod_Retorno = 3
        Exit Function
    End If

    '--------------------CLASSIFICACAO----------------------------------------
    If rsP00APOL("cod_ramo") <> 710 Then
        If rsP00APOL("cod_clasf") = "" Then
            'CODIGO DA CLASSIFICACAO NAO ESTA PREENCHIDO'
            pCod_Retorno = 4
            Exit Function
        End If
        If rsP00APOL("cod_compl_clasf") = "" Then
            'COMPLEMENTO DA CLASSIFICACAO NAO ESTA PREENCHIDO'
            pCod_Retorno = 5
            Exit Function
        End If
    End If
    
    '--------------------BONUS------------------------------------------------
    If Mid(pobjstcPed.Cod_Segurado, 4, 6) = 950728 Then
        lCodIdenTelemar = True
    Else
        lCodIdenTelemar = False
    End If
    If rsP00APOL("cod_ramo") = 113 And lCodIdenTelemar = True Then
        If rsP00APOL("per_bonus") <> 25 Then
            'PERCENTUAL DE BONUS INVALIDO PARA TELEMAR'
            pCod_Retorno = 6
            Exit Function
        End If
    End If
    
    '--------------------DESCONTO PROGRESSIVO---------------------------------
    If rsP00APOL("cod_ramo") = 113 And lCodIdenTelemar = True Then
        If rsP00APOL("per_desc_progres") <> 5 Then
            ''PERCENTUAL DE DESCONTO PROGRESSIVO INVALIDO PARA TELEMAR'
            pCod_Retorno = 7
            Exit Function
        End If
    End If


    '--------------------CONSTRUCAO--------------------------------------------
    If (rsP00APOL("cod_ramo") = 113) And (rsP00APOL("cod_clasf") = "001 ") And _
       (rsP00APOL("cod_compl_clasf") = "21" Or "22") And (rsP00APOL("cod_constr") <> 1) Then
        'CODIGO DE CONSTRUCAO INVALIDA PARA ESTA CLASSIFICACAO
        pCod_Retorno = 8
        Exit Function
    End If

    '--------------------QTDE DE EMPREGADOS------------------------------------
    lQtdeEmpreg = rsP00APOL("qtd_empreg")
    
    
    '--------------------COBERTURAS--------------------------------------------
    'VERIFICAR TAB_PED_COBERT / RAMOSDIVERSOS..TAB_PED_COBERT
    If gM46V999.gBaseLocal = True Then
        lNomeArq = "TAB_PED_COBERT"
    Else
        lNomeArq = "RAMOSDIVERSOS.DBO.TAB_PED_COBERT"
    End If
   
    '1. Obtém registro da apólice.
    'a) Monta comando SELECT.
    lSelect = "SELECT * FROM " & lNomeArq
    lSelect = lSelect & " Where Num_Ped  =  " & lParm.Num_Ped
    lSelect = lSelect & " AND   Num_Item =  " & lParm.Num_Item
    lSelect = lSelect & " AND   Cod_Ramo =  " & lParm.Cod_Ramo
    lSelect = lSelect & " AND   Cod_Org  = 0 AND  Num_Ped <> 0"
    
    'b) Abre arquivo.
    If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, bdP00APOL, lMensagem) = False Then
        Exit Function
    End If
    'c) Obtém registro.
    If gM46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, lMensagem) = False Then
        Call gM46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If
    If rsP00APOL.EOF = True Then
        lAchouCobert = False
        lMensagem = "Coberturas do item " & lParm.Num_Item & " não localizadas."
    Else
        lAchouCobert = True
    End If
    
    
    If (rsP00APOL("cod_ramo") = 113) And (rsP00APOL("cod_cobert") = "0021") Then
        lPermiteEmpreg = True
    Else
        lPermiteEmpreg = False
    End If
    
    '--------------------QTDE DE EMPREGADOS------------------------------------
    If lPermiteEmpreg = False Then
        If lQtdeEmpreg <> 0 Then
            'QUANTIDADE DE EMPREGADOS DEVE SER ZEROS'
            pCod_Retorno = 10
            Exit Function
        End If
    Else
        If lQtdeEmpreg = 0 Then
            'INFORME A QUANTIDADE DE EMPREGADOS
            pCod_Retorno = 11
            Exit Function
        End If
    End If

    '--------------------IS POR COBERTURA--------------------------------------
    If rsP00APOL("cod_ramo") = 112 Or 113 Or 114 Then

    End If
 
 
     '--------------------------TAXAS POR COBERTURA-----------------------------



    '--------------------DESCONTO COMERCIAL POR COBERTURA----------------------
    If (rsP00APOL("cod_ramo") = 113) And lCodIdenTelemar = True Then
        'PERCENTUAL DESC. COMERCIAL DA COBERTURA 9999 INVALIDO PARA TELEMAR'
        pCod_Retorno = 9
        Exit Function
    End If

    
    '--------------------DESCONTO DE EXPERIENCIA POR COBERTURA----------------


    
    
    '--------------------CLAUSULAS---------------------------------------------
    
    'VERIFICAR TAB_APOL_CLAU / GED.DBO.TAB_PED_CLAU
    If gM46V999.gBaseLocal = True Then
        lNomeArq = "TAB_APOL_CLAU"
    Else
        lNomeArq = "RAMOSDIVERSOS.DBO.TAB_PED_CLAU"
    End If
   
    '1. Obtém registro da apólice.
    'a) Monta comando SELECT.
    lSelect = "SELECT * FROM " & lNomeArq
    lSelect = lSelect & " Where Num_Pi  =  " & pobjstcPedLoc.NUM_PI
    lSelect = lSelect & " AND   Num_Item =  " & pobjstcPedLoc.Num_Item
    
    'b) Abre arquivo.
    If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, bdP00APOL, lMensagem) = False Then
        Exit Function
    End If
    'c) Obtém registro.
    If gM46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, lMensagem) = False Then
        Call gM46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If
 
    'Clausula do beneficiário.
    Dim lClausula1  As Boolean
    Dim lClausula2  As Boolean
    
    lClausula1 = False
    lClausula2 = False
    Do While Not rsP00APOL.EOF
        If rsP00APOL("Cod_Clau") = 1 Then lClausula1 = True
        If rsP00APOL("Cod_Clau") = 2 Then lClausula2 = True
    Loop
    
    If pobjstcPedLoc.Cod_Ramo = 113 Then
        If lClausula1 = False And lClausula2 = False Then
            'FALTA CLAUSULA 0001 OU 0002'
            pCod_Retorno = 12
            Exit Function
        Else
            If lClausula1 = True And lClausula2 = True Then
                'DIGITE SOMENTE CLAUSULA 0001 OU 0002'
                pCod_Retorno = 13
                Exit Function
            End If
        End If
    End If
    If pobjstcPedLoc.Cod_Ramo = 112 Then
        If lClausula1 = False And (pobjstcPedLoc.Cod_Clasf = "197 " Or "134 ") Then
            'PARA CLASSIFICACAO 197 E 134, A CLAUSULA DEVE SER 0001'
            pCod_Retorno = 14
            Exit Function
        End If
    End If
    
    gfPesqDadosItem = True
    
    Call mpDeletarArqCalculo
    
    Exit Function
    
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "gfPesqDadosItem ERRO - " & Err & " " & Err.Description & " " & lMensagem)

End Function

Private Sub mpDeletarArqCalculo()
Dim lDiretorio  As String
Dim lNomeArquivo    As String
Dim lNomeComputador As String

    On Error Resume Next
    lDiretorio = Replace(UCase(gM46V999.gEndP00MULTI), "P00MULTI.MDB", "")
    lNomeComputador = gM99V001.GetLoginComputerName()
    lNomeArquivo = lNomeComputador & ".PROCESSANDO"
    Kill lDiretorio & lNomeArquivo
    Err.Clear
    On Error GoTo 0
End Sub

Private Function mpVerCalculoParalelo() As Boolean
Dim lDiretorio  As String
Dim lOutroArq   As String
Dim lArquivo    As Integer
Dim lNomeArquivo    As String
Dim lNomeComputador As String
    On Error GoTo ControleErro
    lDiretorio = Replace(UCase(gM46V999.gEndP00MULTI), "P00MULTI.MDB", "")
    lNomeComputador = gM99V001.GetLoginComputerName()
    lNomeArquivo = lNomeComputador & ".PROCESSANDO"
Retorno:
    lOutroArq = Dir(lDiretorio & "*.PROCESSANDO")
    If gM46V999.gfPreenchido(lOutroArq) Then
        If FileDateTime(lDiretorio & lOutroArq) < (DateAdd("s", -20, Now)) Then
            Kill lDiretorio & lOutroArq
            GoTo Retorno
        Else
            'If lNomeMaq <> Left(lOutroArq, InStr(1, lOutroArq, ".") - 1) Then
                'Call  gM46V999.gpMenErro("Existe outra estação efetuando um cálculo simultâneo.(" & lNomeComputador & "," & Left(lOutroArq, InStr(1, lOutroArq, ".") - 1) & ")")
                mpVerCalculoParalelo = True
            'End If
        End If
    Else
        'crio o arquivo
        lArquivo = FreeFile
        Open lDiretorio & lNomeArquivo For Output As #lArquivo
            Print #lArquivo, "Processando " & Now
         Close #lArquivo
    End If
    
    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(0, "mpVerCalculoParalelo - " & Err.Description)
End Function


