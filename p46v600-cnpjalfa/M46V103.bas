Attribute VB_Name = "basM46V103"
Option Explicit
'---   Dados fixos da proposta
Dim lCod_Ramo                 As Integer
Dim lCod_Modalidade           As Integer
Dim lCod_Tarifa               As Integer
Dim lCod_Sequencia            As Integer
Dim lCod_Corretor             As Long
Dim lCod_Clasf                As String
Dim lCod_Compl_Clasf          As String
Dim lCod_Construcao           As Integer
Dim lNum_Cep                  As Long
Dim lCod_Cobert               As Integer
Dim lVal_IS                   As Double
Dim lVal_IS_Basica            As Double
Dim lVal_VR                   As Double
Dim lVal_Franquia             As Double
Dim lTip_Emissao              As e_TipEmissao
Dim lCoePrazo                 As Single
Dim lTip_Cobert            As Integer

'1100192 - Tarifa 1006 e 1007 Marco 2011
Dim lVal_VR_Item              As Double


'---  Dados para Identificação dos Coeficientes
Dim lCod_Classe          As Integer
Dim lClasseLMG           As Integer
Dim lCodFaixa_Reg        As Integer
Dim lCoef_Regiao         As Double 'Coeficiente de Região
Dim lCoef_LMG            As Double 'Coeficiente de LMG
Dim lCoef_Franquia       As Double 'Coeficiente de Franquia
Dim lCoef_Protecao       As Double 'Coeficiente de Proteção
Dim lCoef_ProtecaoResseg As Double 'Coeficiente de Proteção somente para Resseguro
Dim lPer_Coml            As Double 'Percentual de Comissão do Corretor
Dim lCoef_PI             As Double 'Coeficiente Período Indenitário
Dim lCoef_Desc_Prog      As Double 'Desconto Progressivo (Tab_Ped_Loc)
Dim lCoef_Bonus          As Double 'Desconto de Experiência ou Bônus
Dim lCoef_Desc_Coml      As Double 'Desconto Comercial
Dim lCoefIs As Double
Dim lClasseAPont         As Integer
Dim lTaxaAPont          As Double

'1001146 - Tarifa 1000 e 1001 Dezembro 2010
'Dim lIndLMI              As Integer
Dim lCodRegiao           As Long

Public Enum e_Tip_Valor
   e_Tip_Valor_DescComl = 1             'Utilizado na precificação
   e_Tip_Valor_ClasAceit = 2            'Utilizado antes de consistir cobol
   e_Tip_Valor_Comissao = 3             'Utilizado na consistência de comissão (limite máximo)
   e_Tip_Valor_CoefRegiao = 4           'Utilizado na precificação
   e_Tip_Valor_CoefLMG = 5              'Utilizado na precificação
   e_Tip_Valor_CoefFranq = 6            'Utilizado na precificação
   e_Tip_Valor_CoefProt = 7             'Utilizado na precificação
   e_Tip_Valor_PerNovo = 8              'Utilizado antes de consistir cobol
   e_Tip_Valor_PerRenov = 9             'Utilizado antes de consistir cobol
   e_Tip_Valor_PerEndosso = 10          'Utilizado antes de consistir cobol
   e_Tip_Valor_PerCobert = 11           'Utilizado antes de consistir cobol
   e_Tip_Valor_Rubrica = 12             'Utilizado antes de consistir cobol
   e_Tip_Valor_ = 13             'Utilizado antes de consistir cobol
   e_Tip_Valor_TaxaClasse = 14          'Utilizado na precificação
   e_Tip_Valor_Classe = 15          'Utilizado na precificação
   e_Tip_Valor_PercFranq = 16          'Utilizado na precificação
   e_Tip_Valor_ValMinFranq = 17         'Utilizado na precificação
   e_Tip_Valor_Segurado = 18            'Utilizado na precificação
   e_Tip_Valor_Regiao = 19              'Utilizado na franquia diferenciada
   'Ficha 1000970 - Tarifa 996 e 997 Outubro 2010
   e_Tip_Valor_CoeficienteTaxa = 20     'Utilizado na taxa para agravar ou descontar
   '1100864 - Tarifa 1018 e 1019 Setembro 2011
   e_Tip_Valor_CoeficienteEspecialTaxa = 21     'Utilizado na taxa para agravar ou descontar
End Enum




Function BuscaCoef_IS_VR(ByVal pPropTabela As String, _
                         ByVal pNomeTabImport As String, _
                         ByRef pobjstcItem As stcA46V708B, _
                         ByVal pCod_Grupo_Cobert As Integer, _
                         ByVal pCod_Cobert As Integer, _
                         ByVal pVal_IS As Double, _
                         ByVal pVal_VR As Double, _
                         ByRef pMensagem As String) As Boolean
Dim lbdTbm_Coef_LMG_VR    As ADODB.Connection         'Arquivo: P00Tarifa.
Dim lrsTbm_Coef_LMG_VR    As ADODB.Recordset          'Registro : Tbm_Fator_LMG.
Dim lCodClau               As Integer
Dim lSelect              As String

    BuscaCoef_IS_VR = True

    If lCod_Ramo = 112 Then
        If gCodProduto >= 1008 Then
            Select Case pobjstcItem.Cod_Produto
                Case 1020, 1030, 1040, 1050, 1060, 1010
                    lCoefIs = 1
                    lCodClau = 1
                    GoSub GravaCodClau
                    Exit Function
            End Select
        ElseIf gCodProduto >= 1004 Then
            If InStr(1, "1020,  1030, 1040,1050, 1060 ,1010", pobjstcItem.Cod_Produto) <> 0 Then
                lCoefIs = 1
                lCodClau = 1
                GoSub GravaCodClau
                Exit Function
            End If
        End If
    End If
    If lCod_Ramo = 112 Then
        If lCod_Clasf = "197" Or lCod_Clasf = "134" Then
             lCoefIs = 1
             lCodClau = 1
             GoSub GravaCodClau
             Exit Function
       End If
    End If
    If lCod_Ramo = 113 Then
        lCoefIs = 1
        Exit Function
    End If
    If lCod_Ramo = 114 Then
        If lCod_Compl_Clasf = "11" Or lCod_Compl_Clasf = "11" Or lCod_Compl_Clasf = "14" Or lCod_Compl_Clasf = "15" Then
           lCoefIs = 1
           Exit Function
        End If
        If gCodProduto >= 1012 Then
           lCoefIs = 1
           Exit Function
        End If
    End If

     If pVal_VR = 0 Then
          If gCod_Dizer <> 0 Then
            lCoefIs = 1
            lCodClau = 1
            GoSub GravaCodClau
            Exit Function
          End If
          pMensagem = "Cobertura:" & pCod_Cobert & " - Valor do Risco deve estar preenchido!"
          BuscaCoef_IS_VR = False
          Exit Function
     End If


    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTbm_Coef_LMG_VR, pMensagem) = False Then
        BuscaCoef_IS_VR = False
        Exit Function
    End If

    If pCod_Grupo_Cobert = 1 Then
         If pVal_VR <= 5000000# Then
              lCoefIs = 1
              lCodClau = 1
              GoSub GravaCodClau
              Exit Function
        End If
    Else
        lSelect = " Select * "
        lSelect = lSelect & " from Tbm_Cob_LucroCess "
        lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
        lSelect = lSelect & " and Cod_Cobert = " & pCod_Cobert
        lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
        lSelect = lSelect & " and Cod_Sequencia <= " & lCod_Sequencia

        If gfObtRegistro(lbdTbm_Coef_LMG_VR, lSelect, lrsTbm_Coef_LMG_VR, pMensagem) = False Then
            Call gpFecha2(lbdTbm_Coef_LMG_VR)
            BuscaCoef_IS_VR = False
            Exit Function
        End If
        
        If lrsTbm_Coef_LMG_VR.EOF Then
            lCoefIs = 1
            lrsTbm_Coef_LMG_VR.Close
            Call gpFecha2(lbdTbm_Coef_LMG_VR)
            Exit Function
        End If
        lrsTbm_Coef_LMG_VR.Close
    End If

    lCodClau = 2
    GoSub GravaCodClau
    
    Dim Per_IS_Vr As Double

    Per_IS_Vr = (pVal_IS / pVal_VR) * 100

    If Per_IS_Vr > 100 Then
       pMensagem = "Cobertura: " & pCod_Cobert & " - Valor do LMG não pode ser maior que Valor do Risco!"
       BuscaCoef_IS_VR = False
       Exit Function
    ElseIf Per_IS_Vr < 40 And pCod_Grupo_Cobert = 1 Then
       pMensagem = "Cobertura: " & pCod_Cobert & " - Valor do LMG não pode ser menor que 40% do Valor do Risco!"
       BuscaCoef_IS_VR = False
       Exit Function
    ElseIf Per_IS_Vr < 5 And pCod_Grupo_Cobert <> 1 Then
       pMensagem = "Cobertura: " & pCod_Cobert & " - Valor do LMG não pode ser menor que 5% do Valor de Lucros Cessantes!"
       BuscaCoef_IS_VR = False
       Exit Function
    End If
     
    lSelect = " Select * "
    lSelect = lSelect & " from Tbm_Coef_LMG_VR "
    lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
    lSelect = lSelect & " and Cod_grupo_cobert = " & pCod_Grupo_Cobert
    lSelect = lSelect & " and Perc_LMI_VR <= " & Replace(Format(Per_IS_Vr, "00.00"), ",", ".")
    lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
    lSelect = lSelect & " and Cod_Sequencia <= " & lCod_Sequencia
    lSelect = lSelect & " order by Perc_LMI_VR desc"

    If gfObtRegistro(lbdTbm_Coef_LMG_VR, lSelect, lrsTbm_Coef_LMG_VR, pMensagem) = False Then
        Call gpFecha2(lbdTbm_Coef_LMG_VR)
        Exit Function
    End If
    
    If Not lrsTbm_Coef_LMG_VR.EOF Then
       lCoefIs = Check(lrsTbm_Coef_LMG_VR("Coef_Agrav"))
    Else
       pMensagem = "Cobertura:" & pCod_Cobert & " - Valor em Risco Inválido!"
       BuscaCoef_IS_VR = False
    End If

    lrsTbm_Coef_LMG_VR.Close
    Call gpFecha2(lbdTbm_Coef_LMG_VR)
    Exit Function

GravaCodClau:

    If pCod_Grupo_Cobert = 1 Then
        '801511 - Favor providenciar a desativação das cláusulas 001 e 002  dos Seguros Empresariais para que não sejam mais incluídas/informadas nas emissões de apólices.
        If lCod_Tarifa >= 948 And lCod_Ramo = 112 Then GoTo NaoIncluir
        If Not mfIncluirClausula(pPropTabela, pNomeTabImport, lCodClau, pobjstcItem, pMensagem) Then
            Exit Function
        End If
NaoIncluir:
    End If
    Return
End Function
Private Function fAchouShopping(ByRef pStcItem As stcA46V708B) As Boolean
    Dim lobjstcFator As stcA46V720B
    fAchouShopping = False
    If pStcItem.Cod_Ramo = e_CodRamo_Empresarial Then
        For Each lobjstcFator In pStcItem.ColecaoPedFator
            If lobjstcFator.Cod_Fator = 9 Then 'RISCO EM SHOPPING CENTER
                fAchouShopping = True
                Exit For
            End If
        Next lobjstcFator
    End If
End Function

Private Function mfCalcPrazo(ByVal pDat_Ini_Vig As Long, ByVal pDat_Ter_Vig As Long, ByVal pCod_Iden_Praz As Integer, pMensagem As String) As Boolean
    Dim lQtdiasCorridos         As Integer
    Dim lVigMeses               As Integer  'Quantidade de meses de vigência.
    Dim lAuxiliar               As String   'Auxiliar
    Dim lSelect
     
    Dim mVigIniAno              As Integer  'Data de início de vigência: ano.
    Dim mVigIniAnoMes           As Long     'Data de início de vigência: ano/mes no formato = aaaamm.
    Dim mVigIniDia              As Integer  'Data de início de vigência: dia.
    Dim mVigIniMes              As Integer  'Data de início de vigência: mês.
    Dim mVigTerAno              As Integer  'Data de término de vigência: ano.
    Dim mVigTerAnoMes           As Long     'Data de término de vigência: ano/mes no formato = aaaamm.
    Dim mVigTerDia              As Long     'Data de término de vigência: dia.
    Dim mVigTerMes              As Integer  'Data de término de vigência: mês.
     
    Dim lbdTab_Prazo            As ADODB.Connection         'Arquivo: P00Tarifa.
    Dim lrsTab_Prazo            As ADODB.Recordset          'Registro : Tab_Prazo.
     
    mfCalcPrazo = False
     
    'Calcula coeficiente de prazo.
    If pDat_Ini_Vig = 0 Or pDat_Ter_Vig = 0 Or pDat_Ter_Vig < pDat_Ini_Vig Then
        lQtdiasCorridos = 0
    Else
        lAuxiliar = Format$(pDat_Ini_Vig, "00000000")    'Formato: aaaammdd.
        mVigIniAno = CInt(Mid(lAuxiliar, 1, 4))
        mVigIniMes = CInt(Mid(lAuxiliar, 5, 2))
        mVigIniDia = CInt(Mid(lAuxiliar, 7, 2))
        mVigIniAnoMes = CLng(Mid(lAuxiliar, 1, 6))
    
        lAuxiliar = Format$(pDat_Ter_Vig, "00000000")    'Formato: aaaammdd.
        mVigTerAno = CInt(Mid(lAuxiliar, 1, 4))
        mVigTerMes = CInt(Mid(lAuxiliar, 5, 2))
        mVigTerDia = CInt(Mid(lAuxiliar, 7, 2))
        mVigTerAnoMes = CLng(Mid(lAuxiliar, 1, 6))
    
        'a) Calcula quantidade de meses de vigência.
        Do Until Val(Mid(mVigIniAnoMes, 1, 4)) = Val(Mid(mVigTerAnoMes, 1, 4))
            mVigTerAnoMes = mVigTerAnoMes - 100
            mVigTerAnoMes = mVigTerAnoMes + 12
        Loop
        lVigMeses = mVigTerAnoMes - mVigIniAnoMes
        lQtdiasCorridos = mVigTerDia - mVigIniDia
        If lQtdiasCorridos < -15 Then
            lVigMeses = lVigMeses - 1
        Else
            If lQtdiasCorridos > 15 Then
                lVigMeses = lVigMeses + 1
            End If
        End If
        'Call gpGraLog(0, "#035/004 - Calcula quantidade de meses de vigência (mfCalcPrazo)." & vbLf & _
        '                vbLf & _
        '                "Início de vigência = " & pDat_Ini_Vig & vbLf & _
        '                "Término de vigência = " & pDat_Ter_Vig & vbLf & _
        '                "Quantidade calculada de meses = " & lVigMeses & vbLf & _
        '                "Indicador de prazo = " & pCod_Iden_Praz)
        If lVigMeses < 12 Then
            lVigMeses = 12
        End If
        If lVigMeses > 12 And pCod_Iden_Praz <> 3 Then
            pCod_Iden_Praz = 0
            Call gpGraLog(2, "FGE0009 - Prazo longo não selecionado para período de vigência maior que 1 " & _
                        "ano (" & lVigMeses & " meses).")
        End If
        lQtdiasCorridos = mfDifDatas(pDat_Ini_Vig, pDat_Ter_Vig)
        'c) Obtém coeficiente de prazo.
        Select Case pCod_Iden_Praz
            Case 1
               ' If lTip_Emissao = 300 Or lTip_Emissao = 301 Or lTip_Emissao = 303 Then
               '     lCoePrazo = 1
               ' Else
                    lCoePrazo = gfTrunca(lQtdiasCorridos / 365, 5)
               ' End If
            Case 2
                pMensagem = ""
Voltar:
                lSelect = "SELECT Val_Coef FROM TAB_PRAZO_CURTO WHERE Qtd_Dias_Ini <= " & lQtdiasCorridos & _
                          " AND Qtd_Dias_Fim >= " & lQtdiasCorridos
                  
                If pMensagem = "" Then
                    lSelect = lSelect & " and Posicao = 1"
                End If
                
                If gfAbrBasDados("P00MULTI2", gTipP00MULTI, gEndP00MULTI, lbdTab_Prazo, pMensagem) = False Then
                    Call gpGraLog(2, pMensagem)
                    Exit Function
                End If
                
                If gfObtRegistro(lbdTab_Prazo, lSelect, lrsTab_Prazo, pMensagem) = False Then
                    If InStr(1, pMensagem, "-2147217904") > 0 Then
                        GoTo Voltar
                    Else
                        Call gpGraLog(2, pMensagem)
                        lCoePrazo = 0
                    End If
                Else
                    If lrsTab_Prazo.EOF = True Then
                        Call gpGraLog(2, "FNL0019 - Coeficiente de prazo curto não localizado (quantidade " & _
                                    "de dias de vigência = " & lQtdiasCorridos & ").")
                        lCoePrazo = 0
                    Else
                        lCoePrazo = lrsTab_Prazo!Val_Coef / 100
                    End If
                    Call gpFecha2(lbdTab_Prazo)
                End If
            Case 3
                If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTab_Prazo, pMensagem) = False Then
                    Call gpGraLog(2, pMensagem)
                    Exit Function
                End If
            
                lSelect = "SELECT Val_Coef FROM TAB_PRAZO_LONGO WHERE Qtd_Meses = " & lVigMeses
                If gfObtRegistro(lbdTab_Prazo, lSelect, lrsTab_Prazo, pMensagem) = False Then
                    Call gpGraLog(2, pMensagem)
                    lCoePrazo = 0
                Else
                    If lrsTab_Prazo.EOF = True Then
                        Call gpGraLog(2, "FNL0025 - Coeficiente de prazo longo não localizado (quantidade " & _
                                    "de meses = " & lVigMeses & ").")
                        lCoePrazo = 0
                    Else
                        lCoePrazo = lrsTab_Prazo!Val_Coef
                    End If
                    Call gpFecha2(lbdTab_Prazo)
                End If
            Case Else
                lCoePrazo = 0
        End Select
        'Call gpGraLog(0, "#035/005 - Obtém coeficiente de prazo (mpCalculo)." & vbLf & vbLf & _
        '                "Início de vigência = " & pDat_Ini_Vig & vbLf & _
        '                "Término de vigência = " & pDat_Ter_Vig & vbLf & _
        '                "Quantidade calculada de meses = " & lVigMeses & vbLf & _
        '                "Quantidade calculada de dias = " & lQtdiasCorridos & vbLf & _
        '                "Indicador de prazo = " & pCod_Iden_Praz & vbLf & _
        '                "Coeficiente de prazo = " & lCoePrazo)
    End If
    mfCalcPrazo = True
End Function


Public Function mfClasseAceit(ByRef pobjstcItem As stcA46V708B, ByRef pobjstcPedido As stcA46V702B, Optional ByRef pMensagem As String) As Integer
'ByVal pObjCobert As clsA46V704A,
    Dim lobjstcCobertura     As stcA46V704B
    Dim lobjstcCorretor      As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    Dim lCod_Aceitacao       As Integer          'Código de Aceitação
    Dim lCod_Aceitacao_Aux   As Integer

    lCod_Aceitacao = 0

    '1. Armazenar variáveis da Proposta
    lCod_Ramo = pobjstcPedido.Cod_Ramo
    lCod_Modalidade = pobjstcItem.Cod_Produto
    lCod_Tarifa = pobjstcItem.Cod_Iden_Tarif
    lCod_Sequencia = pobjstcPedido.Cod_Sequencia
    lTip_Emissao = pobjstcItem.Tip_Emissao
    lCod_Clasf = pobjstcItem.Cod_Clasf
    lCod_Compl_Clasf = pobjstcItem.Cod_Compl_Clasf
    lCod_Construcao = pobjstcItem.Cod_Constr
    lNum_Cep = pobjstcItem.Num_Cep_Risco
    lVal_VR = pobjstcItem.Val_Vr
       
    For Each lobjstcCorretor In pobjstcPedido.ColecaoPedCorr
        lCod_Corretor = lobjstcCorretor.Cod_Corr
        lPer_Coml = lPer_Coml + lobjstcCorretor.Per_Corr
        For Each lobjstcCobertura In pobjstcItem.ColecaoPedCobert
            lVal_IS = lobjstcCobertura.Val_IS
            lCod_Cobert = lobjstcCobertura.Cod_Cobert
            lVal_Franquia = lobjstcCobertura.Val_Franq_Min
            If mfAjustesPontuais(e_Tip_Valor_ClasAceit, TpoBoolean, lCod_Aceitacao_Aux, pMensagem) = False Then
               Call gpGraLog(2, pMensagem)
               Exit Function
            End If
            If lCod_Aceitacao_Aux > lCod_Aceitacao Then
                lCod_Aceitacao = lCod_Aceitacao_Aux
            End If
        Next lobjstcCobertura
    Next lobjstcCorretor
    mfClasseAceit = lCod_Aceitacao
'Ficha 1000305
    lTip_Cobert = pobjstcItem.Tip_Cobert
    gTipCobertura = pobjstcItem.Tip_Cobert
    
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcCobertura = Nothing
    Set lobjstcCorretor = Nothing
    Return
End Function

Public Function mfConsisteComissao(ByVal pobjstcItem As stcA46V708B, ByVal pobjstcPedido As stcA46V702B) As Boolean

    Dim lobjstcCobertura     As stcA46V704B
    Dim lobjstcCorretor      As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    Dim lMensagem            As String
    Dim lCod_Aceitacao       As Integer          'Código de Aceitação
    Dim lCod_Aceitacao_Aux   As Integer
    
    Dim lCod_Corretor1             As Long
    Dim lPer_Coml1                 As Double
    Dim lCod_Corretor2             As Long
    Dim lPer_Coml2                 As Double
    Dim lCod_Corretor3             As Long
    Dim lPer_Coml3                 As Double
    Dim lPer_ComlTot               As Double
    
    Dim lCod_Comissao_Aux
    
    'aqui

    mfConsisteComissao = True
    Exit Function
    mfConsisteComissao = False
    
    '1. Armazenar variáveis da Proposta
    lCod_Ramo = pobjstcPedido.Cod_Ramo
    lCod_Modalidade = pobjstcItem.Cod_Produto
    lCod_Tarifa = pobjstcItem.Cod_Iden_Tarif
    lCod_Sequencia = pobjstcPedido.Cod_Sequencia
    lTip_Emissao = pobjstcItem.Tip_Emissao
    lCod_Clasf = pobjstcItem.Cod_Clasf
    lCod_Compl_Clasf = pobjstcItem.Cod_Compl_Clasf
    lCod_Construcao = pobjstcItem.Cod_Constr
    lNum_Cep = pobjstcItem.Num_Cep_Risco
    lVal_VR = pobjstcItem.Val_Vr
    
    For Each lobjstcCorretor In pobjstcPedido.ColecaoPedCorr
        For Each lobjstcCobertura In pobjstcItem.ColecaoPedCobert
            lVal_IS = lobjstcCobertura.Val_IS
            lCod_Cobert = lobjstcCobertura.Cod_Cobert
            lVal_Franquia = lobjstcCobertura.Val_Franq_Min
            
            If mfAjustesPontuais(e_Tip_Valor_Comissao, TpoValor, lCod_Comissao_Aux, lMensagem) = False Then
               Call gpGraLog(2, lMensagem)
               Exit Function
            End If
            If lCod_Aceitacao_Aux > lCod_Aceitacao Then
                lCod_Aceitacao = lCod_Aceitacao_Aux
            End If
        Next lobjstcCobertura
    Next lobjstcCorretor
    
    lCod_Corretor = lobjstcCorretor.Cod_Corr
    lPer_Coml = lPer_Coml + lobjstcCorretor.Per_Corr
'Ficha 1000305
    lTip_Cobert = pobjstcItem.Tip_Cobert
    gTipCobertura = pobjstcItem.Tip_Cobert
    
    mfConsisteComissao = True

End Function


Private Function mfDifDatas(ByVal pInicio As Long, ByVal pTermino As Long) As Single
    'Função #014: calcula diferença entre datas (retorno em dias).

    'Parâmetros de entrada...pInicio....data de início.
    '                        pTermino...data de término.

    Dim lAuxIniDias         As String   'Auxiliar: valores para tabela para cálculo da diferença de dias.
    Dim lAuxTerDias         As String   'Auxiliar: valores para tabela para cálculo da diferença de dias.
    Dim lDia                As Single   'Índice auxiliar: dia.
    Dim li                  As Integer  'Índice.
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
    li = 1
    lDia = 1
    Do Until lDia > 31
        lMes = 1
        Do Until lMes > 12
            lTabIniDias(lDia, lMes) = Val(Mid(lAuxIniDias, li, 3))
            lMes = lMes + 1
            li = li + 3
        Loop
        lDia = lDia + 1
    Loop

    '3. Posiciona dados da tabela de data de término.
    lDia = 1
    li = 1
    Do Until lDia > 31
        lMes = 1
        Do Until lMes > 12
            lTabTerDias(lDia, lMes) = Val(Mid(lAuxTerDias, li, 3))
            lMes = lMes + 1
            li = li + 3
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
        lQuaTerDias = Val(lTabTerDias(lTerDia, lTerMes)) + lPrazoLongo
    Else
        lQuaTerDias = Val(lTabIniDias(lTerDia, lTerMes)) + lPrazoLongo
    End If

    '6. Calcula diferença de dias.
    mfDifDatas = lQuaTerDias - lQuaIniDias
    
    'Call gpGraLog(0, "#014/001 - Calcula diferença de datas (mfDifDatas)." & vbLf & vbLf & _
    '                "Data de início = " & pInicio & vbLf & _
    '                "Data de término = " & pTermino & vbLf & _
    '                "Quantidade de dias da data de início (lQuaIniDias) = " & lQuaIniDias & vbLf & _
    '                "Quantidade de dias da data de término (lQuaTerDias) = " & lQuaTerDias & vbLf & _
    '                "Quantidade calculada de dias = " & mfDifDatas)
End Function

Sub mpCalcula_Endosso(pNum_Apol_Anterior As Long, pNum_Item_Anterior As Integer, pCod_Tip_Calc_Ws As String, ByRef pVal_pr_ws As Single)
    Dim lSql        As String
    Dim lrst        As ADODB.Recordset
    Dim lbdP00APOL  As ADODB.Connection
    Dim lNomeTab    As String
    Dim li          As Integer
    Dim lUltimoErro As String

    If gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
    Else
        lNomeTab = "GED.DBO.TAB_PED"
    End If

    If gCod_Dizer = 4 Or gCod_Dizer = 10 Then
        pVal_pr_ws = 0
    End If

    If gfAbrBasDados("P00APOL", gTipP00APOL, gEndP00APOL, lbdP00APOL, lUltimoErro) Then
        lSql = " Select  Val_pr_anu, Val_pr_tarif_anu,  Val_Pr_Resg_Anu from " & lNomeTab & "_Cobert where num_apol = " & pNum_Apol_Anterior
        lSql = lSql & " and num_item = " & pNum_Item_Anterior
        lSql = lSql & " and cod_cobert = " & lCod_Cobert
        If gfObtRegistro(lbdP00APOL, lSql, lrst, lUltimoErro) Then
            If lrst.EOF Then
                Call gpGraLog(0, "O registro de Bonus em 'mfRetornaBonusApolice' para apólice: " & pNum_Apol_Anterior & " não foi encontrado-" & lUltimoErro)
                Exit Sub
            End If
        End If
    Else
        Call gpGraLog(0, "Erro - Abertura do Banco P00APOL em 'mfRetornaBonusApolice' para apólice: " & pNum_Apol_Anterior & "-" & lUltimoErro)
        Exit Sub
    End If

                                                                  
    If pCod_Tip_Calc_Ws = "COML" Then
        pVal_pr_ws = pVal_pr_ws - Check(lrst("Val_pr_anu"))
    Else
        If pCod_Tip_Calc_Ws = "TARI" Then
               pVal_pr_ws = pVal_pr_ws - Check(lrst("Val_pr_tarif_anu"))
        Else
            If pCod_Tip_Calc_Ws = "RESG" Then
               pVal_pr_ws = pVal_pr_ws - Check(lrst("Val_pr_resg_anu"))
            End If
        End If
    End If
    
    Call gpFecha3(lrst, lbdP00APOL)
    
    If pVal_pr_ws > -0.1 And pVal_pr_ws < 8 Then
        pVal_pr_ws = 0
    End If
    If gCod_Dizer = 6 Then
        pVal_pr_ws = 0
    End If
 '''   pVal_pr_ws = pVal_pr_ws * lCoePrazo
                                                                  
End Sub
Public Function mfCalNewPrecif(ByRef pStcItem As stcA46V708B, _
                               ByVal pReCalcProg As Boolean, _
                               ByRef pMensagem As String, _
                               ByVal pPropTabela As String, _
                               ByVal pNomeTabImport As String) As Boolean
    
    mfCalNewPrecif = False
    
    Dim lArrayTaxas()        As Variant
    Dim lQtdeTaxas           As Integer
    Dim TxAlterada           As Boolean
    Dim lQtdeCobert          As Integer
    
    Dim lobjstcCobertura     As stcA46V704B
    Dim lobjstcCorretor      As stcA46V705B       'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    
    Dim lbdTab_Tarifa        As ADODB.Connection  'Arquivo: P00Tarifa.
    Dim lrsTab_Tarifa        As ADODB.Recordset   'Registro : TAB_Tarifa.
    Dim I                    As Integer
    Dim lObjPedLoc           As clsA46V708A      'Objeto da classe de itens.
    
'--- Taxas Intermediárias
    Dim lTaxaNet             As Single
    Dim lTaxaComl            As Single
    Dim lTaxaFinal           As Single
    
    Dim lPremio              As Single
    Dim lTaxaMin             As Single
    Dim lTaxaReser           As Single
    Dim lTaxaRess            As Single
    Dim lobjCobert           As clsA46V704A
    Dim lRetorno             As String
    
    Dim mPer_Coml            As Single
    Dim lPer_Coml_Cota       As Single
    Dim lVal_IS_Minima_ws    As Single
    Dim lVal_is_basica_ws    As Single

    Dim lColecaoTemp         As Collection
    Dim lComissaoBasica      As Double
    Dim lRegiao              As Integer
    
    Dim lSelect             As String
    Dim bdPremio            As ADODB.Connection
    Dim lrst                As ADODB.Recordset
    Dim lPremioAnt          As Single
    On Error GoTo 0
    
    
    
    
    
    '''Call mpVerNormas(pStcItem.ObjPai, pStcItem, pMensagem)
    Set lobjCobert = New clsA46V704A
    '1. Armazenar variáveis da Proposta
    lCod_Ramo = pStcItem.ObjPai.Cod_Ramo
    lCod_Modalidade = pStcItem.Cod_Produto
    lCod_Tarifa = pStcItem.Cod_Iden_Tarif
    lCod_Sequencia = pStcItem.ObjPai.Cod_Sequencia
    lTip_Emissao = pStcItem.Tip_Emissao
    lCod_Clasf = pStcItem.Cod_Clasf
    lCod_Compl_Clasf = pStcItem.Cod_Compl_Clasf
    lCod_Construcao = pStcItem.Cod_Constr
    lNum_Cep = pStcItem.Num_Cep_Risco
    pStcItem.Per_Tot_Protec = 0
    'Ficha 1000305
    lTip_Cobert = pStcItem.Tip_Cobert
    gTipCobertura = pStcItem.Tip_Cobert
    
    '1100192 - Tarifa 1006 e 1007 Marco 2011
    lVal_VR_Item = pStcItem.Val_Vr
    
    '1001146 - Tarifa 1000 e 1001 Dezembro 2010
'    lIndLMI = pStcItem.ObjPai.LMI_Unico
    lCodRegiao = pStcItem.gCod_RegiaoNova
    
    'Resseguro facultativo
    'Call gpVerNormasParametrizada(pStcItem.ObjPai, pStcItem, pMensagem, 1)
    
    '1200029 - Clausula 315 e Texto LMI
    If pStcItem.ObjPai.LMI_Unico = 1 And pStcItem.Cod_Ramo = e_CodRamo_Empresarial Then
        If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 315, pStcItem, pMensagem) Then
            Call gpGraLog(1, pMensagem)
            GoTo SAIDA
        End If
    End If
    
    If gCodProduto >= 942 Then
        If Val(pStcItem.Cod_Clasf) = 190 And Val(pStcItem.Cod_Compl_Clasf) = 13 Then
            If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 312, pStcItem, pMensagem) Then
                Call gpGraLog(1, pMensagem)
                GoTo SAIDA
            End If
        End If
        If Val(pStcItem.Cod_Clasf) = 190 And Val(pStcItem.Cod_Compl_Clasf) = 14 Then
            If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 313, pStcItem, pMensagem) Then
                Call gpGraLog(1, pMensagem)
                GoTo SAIDA
            End If
        End If
    End If
    
    If mfCalcPrazo(pStcItem.ObjPai.Dat_Ini_Vig, pStcItem.ObjPai.Dat_Ter_Vig, pStcItem.ObjPai.Cod_Iden_Praz, pMensagem) = False Then
        Call gpGraLog(0, pMensagem)
        GoTo SAIDA
    End If
    
    'Encontrar Corretor e Percentual de Comissão
    lPer_Coml = 0
    lCod_Corretor = 0
    
    If pStcItem.ObjPai.ColecaoPedCorr.Count > 0 Then
        For Each lobjstcCorretor In pStcItem.ObjPai.ColecaoPedCorr
            If lobjstcCorretor.Num_Ordem_Corr = 1 Then
                lCod_Corretor = lobjstcCorretor.Cod_Corr
            End If
            lPer_Coml = lPer_Coml + lobjstcCorretor.Per_Corr
        Next lobjstcCorretor
    End If

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTab_Tarifa, pMensagem) = False Then
        Call gpGraLog(0, pMensagem)
        GoTo SAIDA
    End If
    
    '-------------------------------------------------------------
    '2. Encontrar as Classes de Cada Cobertura e
       'Guardar numa array
    '-------------------------------------------------------------
    
    lQtdeTaxas = 0
    lQtdeCobert = 0
    lVal_IS_Minima_ws = 0
    lVal_is_basica_ws = 0
    If pReCalcProg Then
        If lCod_Ramo = 112 Or lCod_Ramo = 113 Or lCod_Ramo = 114 Then
            For Each lobjstcCobertura In pStcItem.ColecaoPedCobert
                If lobjstcCobertura.gCod_Grupo_Cobert = "1" Then
                    lVal_is_basica_ws = lobjstcCobertura.Val_IS
                    If lCod_Ramo = 112 Then
                         lVal_IS_Minima_ws = lVal_is_basica_ws * 0.05
                         If lVal_IS_Minima_ws > 10000 Then
                            lVal_IS_Minima_ws = 10000
                         End If
                    ElseIf lCod_Ramo = 113 Then
                         lVal_IS_Minima_ws = lVal_is_basica_ws * 0.05
                         If lVal_IS_Minima_ws > 5000 Then
                            lVal_IS_Minima_ws = 5000
                         End If
                    ElseIf lCod_Ramo = 114 Then '1200004 - Tarifa 1028 e 1029 Janeiro 2012
                         If lCod_Tarifa >= 1028 Then
                            lVal_IS_Minima_ws = lVal_is_basica_ws * 0.05
                            If lVal_IS_Minima_ws > 5000 Then
                               lVal_IS_Minima_ws = 5000
                            End If
                         End If
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    
    '1100864 - Tarifa 1018 e 1019 Setembro 2011
    If (gCodProduto >= 1018) And (Val(lCod_Clasf) = 282) And (Val(lCod_Compl_Clasf) = 10) And (pStcItem.Cod_Ramo = e_CodRamo_Empresarial) Then
        'CLAUSULA PARTICULAR - DANOS A BENS DE HOSPEDES ATÉ R$ 20 MIL
        If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 320, pStcItem, pMensagem) Then
            Call gpGraLog(1, pMensagem)
            GoTo SAIDA
        End If
    End If
    
    For Each lobjstcCobertura In pStcItem.ColecaoPedCobert
        lQtdeTaxas = lQtdeTaxas + 1
        ReDim Preserve lArrayTaxas(7, lQtdeTaxas)
        
        lCod_Cobert = lobjstcCobertura.Cod_Cobert
        lVal_IS = lobjstcCobertura.Val_IS
        
        'If lobjstcCobertura.gCod_Grupo_Cobert > "1" Then '2 = Acessórias
        '    lQtdeCobert = lQtdeCobert + 1                '3 = Adicionais
        'End If
        If lobjstcCobertura.Cod_Cobert = 22 And gCodProduto >= 948 And lobjstcCobertura.Cod_Ramo = e_CodRamo_Residencial Then
            If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 304, pStcItem, pMensagem) Then
                Call gpGraLog(1, pMensagem)
                GoTo SAIDA
            End If
        End If


        '1100589 - Tarifa 1014 e 1015 Julho 2011
        If gCodProduto >= 1014 Then
            If lobjstcCobertura.Cod_Ramo = e_CodRamo_Empresarial Then
                Select Case lobjstcCobertura.Cod_Cobert
                    Case 56, 122, 123
                        'CLAUSULA PARTICULAR - COBERT. PATIOS - SUBTRACAO DE VEICULOS
                        If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 318, pStcItem, pMensagem) Then
                            Call gpGraLog(1, pMensagem)
                            GoTo SAIDA
                        End If
                    Case 57, 124, 125
                        If gCodProduto <= 1015 Then '1100705 - Tarifa 1016 e 1017 Agosto 2011
                            'CLAUSULA PARTICULAR - COBERT.RC.CONCE. SUBTRACAO DE VEICULOS
                            If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 319, pStcItem, pMensagem) Then
                                Call gpGraLog(1, pMensagem)
                                GoTo SAIDA
                            End If
                        End If
                    Case 23
                        '1100864 - Tarifa 1018 e 1019 Setembro 2011
                        If (gCodProduto >= 1018) And Val(lCod_Clasf) = 282 And Val(lCod_Compl_Clasf) = 10 Then
                            'CLAUSULA PARTICULAR - ROUBO/FURTO C/ VEST. BENS DE HOSPEDES
                            If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 321, pStcItem, pMensagem) Then
                                Call gpGraLog(1, pMensagem)
                                GoTo SAIDA
                            End If
                        End If
                    Case 24
                        '1100864 - Tarifa 1018 e 1019 Setembro 2011
                        If (gCodProduto >= 1018) And Val(lCod_Clasf) = 282 And Val(lCod_Compl_Clasf) = 10 Then
                            'CLAUSULA PARTICULAR - ROUBO DE VALORES INTERIOR DO ESTAB.
                            If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 322, pStcItem, pMensagem) Then
                                Call gpGraLog(1, pMensagem)
                                GoTo SAIDA
                            End If
                        End If
                End Select
            End If
        End If
         
         
         'Calcula qtde de cobertura para o desconto progressivo -> Per_Desc_Progres
        If lCod_Tarifa >= 1028 Then '1200004 - Tarifa 1028 e 1029 Janeiro 2012
             If pReCalcProg Then
                If lCod_Ramo = 112 And _
                    lobjstcCobertura.gCod_Grupo_Cobert > 1 And _
                    lVal_IS >= lVal_IS_Minima_ws Then
                    lQtdeCobert = lQtdeCobert + 1
                ElseIf lCod_Ramo = 113 And _
                    lobjstcCobertura.gCod_Grupo_Cobert > 1 And _
                    lVal_IS >= lVal_IS_Minima_ws Then
                    lQtdeCobert = lQtdeCobert + 1
                ElseIf lCod_Ramo = 114 And _
                    lobjstcCobertura.gCod_Grupo_Cobert > 1 And _
                    lVal_IS >= lVal_IS_Minima_ws Then
                    lQtdeCobert = lQtdeCobert + 1
                End If
            End If
        Else
             If pReCalcProg Then
                 If lCod_Ramo = 112 And _
                    lobjstcCobertura.gCod_Grupo_Cobert > 1 And _
                    lCod_Cobert <> 23 And _
                    lCod_Cobert <> 24 And _
                    lCod_Cobert <> 25 And _
                    lCod_Cobert <> 26 And _
                    lCod_Cobert <> 63 And _
                    lVal_IS >= lVal_IS_Minima_ws Then
                    lQtdeCobert = lQtdeCobert + 1
                ElseIf lCod_Ramo = 113 And lCod_Tarifa < 946 And _
                    lobjstcCobertura.gCod_Grupo_Cobert > 1 And _
                    lCod_Cobert <> 16 And _
                    lCod_Cobert <> 5 And _
                    lVal_IS >= lVal_IS_Minima_ws Then
                    lQtdeCobert = lQtdeCobert + 1
                ElseIf lCod_Ramo = 113 And lCod_Tarifa >= 946 And _
                    lobjstcCobertura.gCod_Grupo_Cobert > 1 And _
                    lCod_Cobert <> 16 And _
                    lVal_IS >= lVal_IS_Minima_ws Then
                    lQtdeCobert = lQtdeCobert + 1
                ElseIf lCod_Ramo = 114 And _
                    lobjstcCobertura.gCod_Grupo_Cobert > 1 Then
                    lQtdeCobert = lQtdeCobert + 1
                End If
            End If
        End If
        
        lClasseLMG = 0
        lCodFaixa_Reg = 0
        lCod_Classe = 0
        lTaxaMin = 0
        lTaxaReser = 0
        lTaxaRess = 0
        lRegiao = 0
        
        '1. Encontrar a Faixa da Regiao
        
        'Ficha 901549 - Faixa de região dos planinhos
        If lCod_Modalidade > 2000 Then
            Select Case lCod_Modalidade
                Case 2051 To 2070
                    lRegiao = 114 'ok
                'Ficha 901673 e 901772
                Case 2111 To 2130
                    lRegiao = 114 'ok
                Case 2042 To 2049
                    lRegiao = 310 'ok
'                Case Else 'Comentado porque foi criada funcionalidade de cadastro/manutenção de plano no gerenciador que gera todo cadastro automaticamente
'                    lCodFaixa_Reg = 1
'                    GoTo Achei_Faixa
            End Select
        End If
        If lRegiao > 0 Then
            lSelect = "Select Cod_Faixa_Reg from Tbm_Faixa_Regiao "
            lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
            lSelect = lSelect & "    and Cod_Modalidade = " & lCod_Modalidade
            lSelect = lSelect & "    and Cod_Reg = " & lRegiao
            lSelect = lSelect & "    and Cod_Tarif = " & lCod_Tarifa
            lSelect = lSelect & "    and Cod_Sequencia  <= " & lCod_Sequencia
            lSelect = lSelect & "    order by Cod_Sequencia desc "

        Else
            lSelect = "Select Cod_Faixa_Reg from Tbm_Faixa_Regiao "
            lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
            lSelect = lSelect & "    and Cod_Modalidade = " & lCod_Modalidade
            lSelect = lSelect & "    and Cep_Ini <= " & lNum_Cep
            lSelect = lSelect & "    and Cep_Fim >= " & lNum_Cep
            lSelect = lSelect & "    and Cod_Tarif = " & lCod_Tarifa
            lSelect = lSelect & "    and Cod_Sequencia  <= " & lCod_Sequencia
            lSelect = lSelect & "    order by Cod_Sequencia desc "
        End If
        If gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, pMensagem) = False Then
            Call gpGraLog(0, pMensagem)
            Call gpFecha2(lbdTab_Tarifa)
            GoTo SAIDA
        End If

        If Not lrsTab_Tarifa.EOF Then
            lCodFaixa_Reg = Check(lrsTab_Tarifa("Cod_Faixa_Reg"))
        End If
        lrsTab_Tarifa.Close
        
Achei_Faixa:
        '2. Encontrar a Classe
        lSelect = "Select Cod_Classe from Tbm_Classe "
        lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
        lSelect = lSelect & "    and Cod_Modalidade = " & lCod_Modalidade
        lSelect = lSelect & "    and Cod_Clasf = '" & lCod_Clasf & "'"
        lSelect = lSelect & "    and Cod_Compl_Clasf = '" & lCod_Compl_Clasf & "'"
        lSelect = lSelect & "    and Cod_Cobert = " & lCod_Cobert
        lSelect = lSelect & "    and Cod_Tarif = " & lCod_Tarifa
        lSelect = lSelect & "    and Cod_Sequencia  <= " & lCod_Sequencia
        lSelect = lSelect & "    order by Cod_Sequencia desc "
        
        If gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, pMensagem) = False Then
            Call gpFecha2(lbdTab_Tarifa)
            Call gpGraLog(0, pMensagem)
            GoTo SAIDA
        End If

        If lrsTab_Tarifa.EOF Then
            lrsTab_Tarifa.Close
            GoTo N_loc_Taxa
        End If
        lCod_Classe = Check(lrsTab_Tarifa("Cod_Classe"))
        lrsTab_Tarifa.Close
        
        
        'Verifico se tem ajustes pontuais para encontrar a classe
        lClasseAPont = 0
        If mfAjustesPontuais(e_Tip_Valor_Classe, tpoNumerico, lClasseAPont, pMensagem) = False Then
           Call gpGraLog(0, pMensagem)
           GoTo SAIDA
        End If
        If lClasseAPont <> 0 Then
            lCod_Classe = lClasseAPont
        End If
        
        '3. Encontrar a Classe para os Vultosos Específicos
       
        lSelect = "Select Cod_Classe_LMG from Tbm_Classe_Vultosos  "
        lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
        lSelect = lSelect & "    and Cod_Modalidade = " & lCod_Modalidade
        lSelect = lSelect & "    and Cod_Tip_Reg = 'E'"
        lSelect = lSelect & "    and Cod_Clasf = '" & lCod_Clasf & "'"
        lSelect = lSelect & "    and Cod_Compl_Clasf = '" & lCod_Compl_Clasf & "'"
        lSelect = lSelect & "    and Cod_Construcao = " & lCod_Construcao
        lSelect = lSelect & "    and Val_LMG_Min <= " & gfForValor(1, lVal_IS)
        lSelect = lSelect & "    and Val_LMG_Max >= " & gfForValor(1, lVal_IS)
        lSelect = lSelect & "    and Cod_Cobert = " & lCod_Cobert
        lSelect = lSelect & "    and Cod_Tarif = " & lCod_Tarifa
        lSelect = lSelect & "    and Cod_Sequencia  <= " & lCod_Sequencia
        lSelect = lSelect & "    order by Cod_Sequencia desc "
        
        If gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, pMensagem) = False Then
            Call gpFecha2(lbdTab_Tarifa)
            Call gpGraLog(0, pMensagem)
            GoTo SAIDA
        End If

        If Not lrsTab_Tarifa.EOF Then
            lClasseLMG = Check(lrsTab_Tarifa("Cod_Classe_LMG"))
        Else
            lrsTab_Tarifa.Close
            '3. Encontrar a Classe para os Vultosos Default
            lSelect = "Select Cod_Classe_LMG from Tbm_Classe_Vultosos  "
            lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
            lSelect = lSelect & "    and Cod_Modalidade = " & lCod_Modalidade
            lSelect = lSelect & "    and Cod_Tip_Reg = 'D'"
            lSelect = lSelect & "    and Val_LMG_Min <= " & gfForValor(1, lVal_IS)
            lSelect = lSelect & "    and Val_LMG_Max >= " & gfForValor(1, lVal_IS)
            lSelect = lSelect & "    and Cod_Cobert = " & lCod_Cobert
            lSelect = lSelect & "    and Cod_Tarif = " & lCod_Tarifa
            lSelect = lSelect & "    and Cod_Sequencia  <= " & lCod_Sequencia
            lSelect = lSelect & "    order by Cod_Sequencia desc "

            If gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, pMensagem) = False Then
                Call gpFecha2(lbdTab_Tarifa)
                Call gpGraLog(0, pMensagem)
                GoTo SAIDA
           End If
            If Not lrsTab_Tarifa.EOF Then
                lClasseLMG = Check(lrsTab_Tarifa("Cod_Classe_LMG"))
            End If
        End If
        lrsTab_Tarifa.Close
                

        '4. Encontrar a Taxa
        lSelect = "Select Per_Taxa_Min,  Per_Taxa_Reser, Per_Taxa_Ress from Tbm_Taxa_Classe "
        lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
        lSelect = lSelect & "    and Cod_Modalidade = " & lCod_Modalidade
        lSelect = lSelect & "    and Cod_Cobert = " & lCod_Cobert
        lSelect = lSelect & "    and Cod_Classe = " & lCod_Classe
        lSelect = lSelect & "    and Cod_Faixa_Reg in (0," & lCodFaixa_Reg & ")"
        lSelect = lSelect & "    and Cod_Classe_LMG in (0," & lClasseLMG & ")"
        lSelect = lSelect & "    and Cod_Construcao in (0," & lCod_Construcao & ")"
    'Ficha 1000305
        lSelect = lSelect & "    and Tip_Cobert in (0," & lTip_Cobert & ")"
        
        lSelect = lSelect & "    and Cod_Tarif = " & lCod_Tarifa
        lSelect = lSelect & "    and Cod_Sequencia  <= " & lCod_Sequencia
        lSelect = lSelect & "    order by Cod_Sequencia desc,Cod_Faixa_Reg desc , Cod_Classe_LMG desc, Cod_Construcao desc, Tip_Cobert desc "
        
        If gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, pMensagem) = False Then
            Call gpFecha2(lbdTab_Tarifa)
            Call gpGraLog(0, pMensagem)
            GoTo SAIDA
        End If

        If lrsTab_Tarifa.EOF Then
            lrsTab_Tarifa.Close
            GoTo Ver_Ajuste
        End If
        If lrsTab_Tarifa.RecordCount = 1 Then
            lTaxaMin = Check(lrsTab_Tarifa("Per_Taxa_Min"))
            lTaxaReser = Check(lrsTab_Tarifa("Per_Taxa_Reser"))
            lTaxaRess = Check(lrsTab_Tarifa("Per_Taxa_Ress"))
        Else
            'Fazer looping  aqui
            'MsgBox "O Sistema encontrou mais de um registro de taxa! Não será possivel o cálculo"
            'Verificar com Produto
            lTaxaMin = Check(lrsTab_Tarifa("Per_Taxa_Min"))
            lTaxaReser = Check(lrsTab_Tarifa("Per_Taxa_Reser"))
            lTaxaRess = Check(lrsTab_Tarifa("Per_Taxa_Ress"))
        End If
        lrsTab_Tarifa.Close
    
Ver_Ajuste:
        'Verifico se tem ajustes pontuais para encontrar a taxa
        lTaxaAPont = 0
        If mfAjustesPontuais(e_Tip_Valor_TaxaClasse, tpoNumerico, lTaxaAPont, pMensagem) = False Then
           Call gpGraLog(0, pMensagem)
           GoTo SAIDA
        End If
        If lTaxaAPont <> 0 Then
            lTaxaMin = lTaxaAPont
            lTaxaReser = lTaxaAPont
            lTaxaRess = lTaxaAPont
        End If

N_loc_Taxa:
        'Ficha 1000970 - Tarifa 996 e 997 Outubro 2010
        lTaxaAPont = 0
        If mfAjustesPontuais(e_Tip_Valor_CoeficienteTaxa, tpoNumerico, lTaxaAPont, pMensagem) = False Then
           Call gpGraLog(0, pMensagem)
           GoTo SAIDA
        End If
        '1100864 - Tarifa 1018 e 1019 Setembro 2011 (Item 3 Não aplicar o coeficiente p/ Renovações Yasuda Sem Sinistro)
        If lTaxaAPont <> 0 Then
            If (lCod_Tarifa < 1018) Or _
               (lCod_Tarifa >= 1018 And pStcItem.Tip_Emissao <> e_TipEmissao_RenYasuda Or pStcItem.Ind_Sinistro <> 2) Then
                    lTaxaMin = lTaxaMin * lTaxaAPont
                    lTaxaReser = lTaxaMin * lTaxaAPont
                    lTaxaRess = lTaxaMin * lTaxaAPont
            End If
        End If


        '1000970 - Item 1 - Empresarial/Residencial/Condominio - para corretor 0663299 Legend agravar em 30% todas as taxas de todas as coberturas dos produtos.
        lTaxaAPont = 0
        If mfAjustesPontuais(e_Tip_Valor_CoeficienteEspecialTaxa, tpoNumerico, lTaxaAPont, pMensagem) = False Then
           Call gpGraLog(0, pMensagem)
           GoTo SAIDA
        End If
        If lTaxaAPont <> 0 Then
            lTaxaMin = lTaxaMin * lTaxaAPont
            lTaxaReser = lTaxaMin * lTaxaAPont
            lTaxaRess = lTaxaMin * lTaxaAPont
        End If


        lArrayTaxas(0, lQtdeTaxas) = lCod_Cobert
        lArrayTaxas(1, lQtdeTaxas) = lClasseLMG
        lArrayTaxas(2, lQtdeTaxas) = lCodFaixa_Reg
        lArrayTaxas(3, lQtdeTaxas) = lCod_Classe
        lArrayTaxas(4, lQtdeTaxas) = lTaxaMin
        lArrayTaxas(5, lQtdeTaxas) = lTaxaReser
        lArrayTaxas(6, lQtdeTaxas) = lTaxaRess
    
    Next lobjstcCobertura
    
    
    '[FRH] F:900097
    '------------------------------------------------------
    Dim val86 As Long
    Dim val64 As Long
    Dim val66 As Long
    Dim Y As Long
    If gCodProduto >= 956 Then
        val86 = 9999
        val64 = 9999
        val66 = 9999
            
        For Y = 0 To UBound(lArrayTaxas, 2)
            Select Case Val(lArrayTaxas(0, Y))
                Case 86: val86 = Y
                Case 64: val64 = Y
                Case 66: val66 = Y
            End Select
        Next Y
        
        If val86 < 9999 Then
            If val64 < 9999 Then
                lArrayTaxas(4, val86) = lArrayTaxas(4, val64)
                lArrayTaxas(5, val86) = lArrayTaxas(5, val64)
                lArrayTaxas(6, val86) = lArrayTaxas(6, val64)
            ElseIf val66 < 9999 Then
                lArrayTaxas(4, val86) = lArrayTaxas(4, val66)
                lArrayTaxas(5, val86) = lArrayTaxas(5, val66)
                lArrayTaxas(6, val86) = lArrayTaxas(6, val66)
            End If
        End If
    End If
    '------------------------------------------------------
    
    lCoef_Desc_Prog = 0
    
   'Buscar Desconto Progressivo
    If (pStcItem.ObjPai.getPlano.EhPlaninho = True) And _
       (pStcItem.ObjPai.getPlano.TemValorRegra(eRegraPlano_DescontoProgressivo) = True) Then
        lCoef_Desc_Prog = pStcItem.ObjPai.getPlano.ValorRegra(eRegraPlano_DescontoProgressivo)
    Else
        If pReCalcProg Then
            If lQtdeCobert = 3 Then
                lCoef_Desc_Prog = 5
            ElseIf lQtdeCobert = 4 Then
                lCoef_Desc_Prog = 7
            ElseIf lQtdeCobert >= 5 Then
                lCoef_Desc_Prog = 10
            End If
        Else
            lCoef_Desc_Prog = pStcItem.Per_Desc_Progres
        End If
    End If
     
    If lCod_Ramo = 113 Then
        If lCod_Modalidade >= 2011 And lCod_Modalidade <= 2020 Then
            lCoef_Desc_Prog = 7
        End If
    
    'Ficha 900740 - o Produto calculou o plano 06 com 7% de Progressivo e não 10%
        If lCod_Modalidade >= 2021 And lCod_Modalidade <= 2031 Then
            lCoef_Desc_Prog = 10
        End If
'1000744 - Item 4 - Plano Mercedes
        If (lCod_Modalidade = 2026) And (gCodProduto < 992) Then
            lCoef_Desc_Prog = 7
        End If
        
    'Ficha 901346 - o Produto calculou o plano 06 com 7% de Progressivo e não 10%
    'Ficha 901673 e 901772
        If (lCod_Modalidade >= 2061 And lCod_Modalidade <= 2100) Or (lCod_Modalidade >= 2111 And lCod_Modalidade <= 2130) Then
            lCoef_Desc_Prog = 10
        End If
    
    End If
    
        
   '-------------------------------------------------------------
   'Nesse momento vai realmente fazer os cálculos
   '-------------------------------------------------------------
   
    Set lColecaoTemp = New Collection
    
    For Each lobjstcCobertura In pStcItem.ColecaoPedCobert
        lTaxaNet = 0
        lTaxaComl = 0
        lTaxaFinal = 0
        lTaxaMin = 0
        
        lVal_IS = lobjstcCobertura.Val_IS
        lVal_VR = lobjstcCobertura.Val_Risco_Cob
        lCod_Cobert = lobjstcCobertura.Cod_Cobert
        lVal_Franquia = lobjstcCobertura.Val_Franq_Min
        
        'Ficha 901673 e 901772
'        If (lCod_Modalidade >= 2001 And lCod_Modalidade <= 2031) Or (lCod_Modalidade >= 2051 And lCod_Modalidade <= 2100) Or (lCod_Modalidade >= 2111 And lCod_Modalidade <= 2130) Then
        'If gEplano = True Then
            'lCoef_Desc_Coml = gDesc_Coml
        If pStcItem.ObjPai.getPlano.ehPlanoEngessado = True Then
            lCoef_Desc_Coml = pStcItem.getPlanoItem.ValorRegra(eRegraPlano_DescontoComercial)
        Else
            lCoef_Desc_Coml = lobjstcCobertura.Per_Desc_Coml
        End If
        
        
        
        'Percentual de Comissão
        Select Case lCod_Ramo
            Case 112
                mPer_Coml = gfTrunca(0.8 / (1 - (lPer_Coml / 100)), 5)
            Case 113, 114
                mPer_Coml = gfTrunca(0.75 / (1 - (lPer_Coml / 100)), 5)
            Case 150
                mPer_Coml = gfTrunca(0.85 / (1 - (lPer_Coml / 100)), 5)
            Case Else
                mPer_Coml = gfTrunca(0.8 / (1 - (lPer_Coml / 100)), 5)
        End Select
        
        

        
        'Caso o digitador tenha feito alteracao manual da taxa ela vai entrar aqui
        TxAlterada = False
        If Trim(UCase(lobjstcCobertura.Alterar_Cobert)) = "ALTERADO" And lobjstcCobertura.Per_Taxa <> 0 Then
            TxAlterada = True
            'Calcular Prêmio Comercial
            'lTaxaMin = lobjstcCobertura.Per_Taxa
            'lTaxaFinal = gfTrunca(lTaxaMin * mPer_Coml, 5)
            'lTaxaReser = lTaxaFinal
            'lTaxaRess = lTaxaFinal
            'GoTo CalculaPremio
        End If
        
        For I = 1 To lQtdeTaxas
            If lArrayTaxas(0, I) = lCod_Cobert Then
                lClasseLMG = lArrayTaxas(1, I)
                lCodFaixa_Reg = lArrayTaxas(2, I)
                lCod_Classe = lArrayTaxas(3, I)
                lTaxaMin = lArrayTaxas(4, I)
                lTaxaReser = lArrayTaxas(5, I)
                lTaxaRess = lArrayTaxas(6, I)
                Exit For
            End If
        Next I
        
        'Verifica se existe taxa para essa cobertura
        'Caso esteja zerada vai procurar a cobertura principal
        If lTaxaMin = 0 Then
            lSelect = " Select *  from Tbm_Cob_Dep_Taxa "
            lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
            lSelect = lSelect & "  and Cod_Modalidade = " & lCod_Modalidade
            lSelect = lSelect & " and Cod_Dep_Principal = " & lCod_Cobert
            lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
            lSelect = lSelect & " and Cod_Sequencia <= " & lCod_Sequencia
        
            If gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, pMensagem) = False Then
                Call gpFecha2(lbdTab_Tarifa)
                Call gpGraLog(0, pMensagem)
                GoTo ProximaCobertura
            End If
    
            If Not lrsTab_Tarifa.EOF Then
                For I = 1 To lQtdeTaxas
                    If lArrayTaxas(0, I) = lrsTab_Tarifa("Cod_Principal") Then
                        lClasseLMG = lArrayTaxas(1, I)
                        lCodFaixa_Reg = lArrayTaxas(2, I)
                        lCod_Classe = lArrayTaxas(3, I)
                        lTaxaMin = lArrayTaxas(4, I)
                        lTaxaReser = lArrayTaxas(5, I)
                        lTaxaRess = lArrayTaxas(6, I)
                        Exit For
                    End If
                Next I
            End If
            lrsTab_Tarifa.Close
        End If
        
        'Coeficiente de IS VR
        If Not BuscaCoef_IS_VR(pPropTabela, pNomeTabImport, pStcItem, lobjstcCobertura.gCod_Grupo_Cobert, lCod_Cobert, lVal_IS, lVal_VR, pMensagem) Then
           Call gpGraLog(0, pMensagem)
           GoTo SAIDA
        End If
        
        'Vai procurar Desconto Comercial na tabela de Ajustes pontuais
        'If mf  AjustesPontuais(e_Tip_Valor_DescComl, TpoValor, lCoef_Desc_Coml, pMensagem) = False Then
        '   Call gpGraLog(0, pMensagem)
        '   GoTo SAIDA
        'End If
        
        'Encontrar Coeficiente de Franquia
        If lobjstcCobertura.Val_Fator_Franq > 0 Then
            lCoef_Franquia = lobjstcCobertura.Val_Fator_Franq
        Else
            lCoef_Franquia = 1
        End If
        
        'Vai procurar Coeficiente de Franquia na tabela de Ajustes pontuais
        If mfAjustesPontuais(e_Tip_Valor_CoefFranq, TpoValor, lCoef_Franquia, pMensagem) = False Then
           Call gpGraLog(0, pMensagem)
           GoTo SAIDA
        End If
        
        'Calcular Prêmio Net
        lTaxaNet = lTaxaMin * lCoef_Franquia * lCoefIs
        lTaxaReser = lTaxaReser * lCoef_Franquia * lCoefIs
        lTaxaRess = lTaxaRess * lCoef_Franquia * lCoefIs
                
        'Encontrar Coeficiente do Período Indenitário
        If mpFindCoefPI(lCoef_PI, lobjstcCobertura.Num_Meses_PI, pMensagem) = False Then
            Call gpGraLog(0, pMensagem)
        End If

        '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 11
        If (pStcItem.ObjPai.Cod_Produto >= 1038) And _
           (pStcItem.ObjPai.Cod_Ramo = e_CodRamo_Empresarial) And _
           (pStcItem.ObjPai.Cod_Estip >= 2170 And pStcItem.ObjPai.Cod_Estip <= 2177) And _
           (lobjstcCobertura.Cod_Cobert = 39) Then
                lCoef_PI = 1
        End If

        'Calcular o Prêmio Comercial com o Período Indenitário
        lTaxaNet = lTaxaNet * lCoef_PI
        lTaxaReser = lTaxaReser * lCoef_PI
        lTaxaRess = lTaxaRess * lCoef_PI
        
        'Calcular Prêmio com Descontro Progressivo
        'Exceto para as coberturas básicas
        If lobjstcCobertura.Cod_Grupo_Cobertura <> 1 Then
            If lCoef_Desc_Prog <> 0 Then
                lTaxaNet = lTaxaNet * (1 - (lCoef_Desc_Prog / 100))
                lTaxaReser = lTaxaReser * (1 - (lCoef_Desc_Prog / 100))
                lTaxaRess = lTaxaRess * (1 - (lCoef_Desc_Prog / 100))
            End If
        End If
        
        'Buscar Bônus / Desconto Experiência
        '<POG> ESSE IF EXISTE SOMENTE PARA MANTER OS DADOS QUE FORAM GERADOS COM O FONTE ERRADO. Remover tudo que esta dentro da tag POG no dia 27/07
        If (UCase(App.EXEName) = "D46O003" _
            And pStcItem.ObjPai.Dat_Pri_Calculo >= 20120625 And pStcItem.ObjPai.Dat_Pri_Calculo <= 20120627) Then
            
                If (lCod_Modalidade >= 2001 And lCod_Modalidade <= 2031) Or (lCod_Modalidade >= 2061 And lCod_Modalidade <= 2100) Or (lCod_Modalidade >= 2111 And lCod_Modalidade <= 2130) Then
                    lCoef_Desc_Coml = pStcItem.getPlanoItem.ValorRegra(eRegraPlano_DescontoComercial)
                Else
                    lCoef_Desc_Coml = lobjstcCobertura.Per_Desc_Coml
                End If
                
        Else
        '</POG>
            lCoef_Bonus = pStcItem.Per_Bonus
        '<POG>
        End If
        '</POG>
        
        'Calcular Prêmio com Bônus
        lTaxaNet = lTaxaNet * (1 - (lCoef_Bonus / 100))
        lTaxaReser = lTaxaReser * (1 - (lCoef_Bonus / 100))
        lTaxaRess = lTaxaRess * (1 - (lCoef_Bonus / 100))
       
        'Calcular Prêmio Comercial
        'Resseguro e Tarifário não faz
        lTaxaComl = gfTrunca(lTaxaNet * mPer_Coml, 5)
        
        'Calcular Prêmio Final
        lTaxaFinal = gfTrunca(lTaxaComl * (1 - (lCoef_Desc_Coml / 100)), 5)
        lCoef_Protecao = 0
        'Encontrar Coeficiente de Fator de Proteção
        If mpFindCoefProtecao(lCoef_Protecao, lCoef_ProtecaoResseg, pStcItem.ColecaoPedFator, pMensagem) = False Then
            Call gpGraLog(0, pMensagem)
            
        End If
        
        If gCodProduto >= 934 Then
            If Not gfVerParametro(8, pStcItem.ObjPai, pStcItem, lobjstcCobertura, lRetorno, pMensagem) Then
                Call gpGraLog(1, pMensagem)
                GoTo SAIDA
            End If
            If gfPreenchido(lRetorno) Then
                If lCoef_Protecao > Val(lRetorno) Then
                    lCoef_Protecao = lRetorno
                End If
            End If
            If gCodProduto >= 942 And pStcItem.Cod_Ramo = e_CodRamo_Empresarial Then
                If lCoef_Protecao > 0 And lobjstcCobertura.Cod_Cobert = 1 Then
                    If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 310, pStcItem, pMensagem) Then
                        Call gpGraLog(1, pMensagem)
                        GoTo SAIDA
                    End If
                End If
                If lCoef_Protecao > 0 And lobjstcCobertura.Cod_Cobert = 23 Then
                    If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 311, pStcItem, pMensagem) Then
                        Call gpGraLog(1, pMensagem)
                        GoTo SAIDA
                    End If
                End If
            End If
        Else
            If gCodProduto >= 932 And lCoef_Protecao > 15 Then
                lCoef_Protecao = 15
            End If
        End If
        
'Ficha 1000304
        If (lobjstcCobertura.Cod_Cobert = 23 Or lobjstcCobertura.Cod_Cobert = 63) Then
            If fAchouShopping(pStcItem) = True Then
                lCoef_Protecao = 30
            End If
        End If
      
        '-->Ficha 1000888
        If lCod_Ramo = 112 Then
            Select Case lobjstcCobertura.Cod_Cobert
                Case 1, 2, 32, 39 Or 42
                    pStcItem.Per_Tot_Protec = lCoef_Protecao
            End Select
        End If
        If lCod_Ramo = 114 And lobjstcCobertura.Cod_Cobert = 1 Then
            pStcItem.Per_Tot_Protec = lCoef_Protecao
        End If
        
        If (lCod_Ramo = 112 And (lobjstcCobertura.Cod_Cobert = 23 Or lobjstcCobertura.Cod_Cobert = 63)) Or _
            (lCod_Ramo = 113 And (lobjstcCobertura.Cod_Cobert = 5 Or lobjstcCobertura.Cod_Cobert = 16)) Or _
            (lCod_Ramo = 114 And lobjstcCobertura.Cod_Cobert = 5) Then
            pStcItem.Per_Tot_Segur = lCoef_Protecao
        End If
        
        pStcItem.Cod_Regiao = pStcItem.gCod_RegiaoNova
        'Ficha 1000888 <--
        
        
        lTaxaFinal = lTaxaFinal * (1 - (lCoef_Protecao / 100))
        lTaxaReser = lTaxaReser * (1 - (lCoef_Protecao / 100))
        lTaxaRess = lTaxaRess * (1 - (lCoef_ProtecaoResseg / 100))
         
        lobjstcCobertura.Per_Desc_Coml = lCoef_Desc_Coml
        lobjstcCobertura.Cod_Classe = lCod_Classe
        
CalculaPremio:
        lPremioAnt = 0
        If TxAlterada Or gfPreenchido(pStcItem.ObjPai.Num_Cotac_LMI) Then
            'Calcular Prêmio Comercial qdo a taxa foi alterada
            If gCod_Ramo = e_CodRamo_Roubos Then
                'Ficha 0901367, não permitir que taxa alterada seja menor que TaxaCalculada
                If lobjstcCobertura.Per_Taxa < lTaxaFinal Then
                    TxAlterada = False
                    lobjstcCobertura.Alterar_Cobert = ""
                Else
                    lTaxaMin = lobjstcCobertura.Per_Taxa
                    lTaxaComl = gfTrunca(lTaxaMin * mPer_Coml, 5)
                    lTaxaFinal = gfTrunca(lTaxaComl * (1 - (lCoef_Desc_Coml / 100)), 5)
                End If
            Else
                If gfPreenchido(pStcItem.ObjPai.Num_Cotac_LMI) Then
                                        
                    
                    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdPremio, pMensagem) = False Then
                        GoTo SAIDA
                    End If
                    lSelect = "select sum(val_pr) as soma , max(per_DESC_COML) AS per_DESC, "
                    lSelect = lSelect & " (select sum(per_corr) from tab_ped_corr where nosso_numero = '" & pStcItem.ObjPai.Num_Cotac_LMI & "')  AS COMIS "
                    lSelect = lSelect & " from tab_ped_cobert where nosso_numero = '" & pStcItem.ObjPai.Num_Cotac_LMI & "'"
                    lSelect = lSelect & " and cod_cobert = " & lobjstcCobertura.Cod_Cobert
                    If Not gfObtRegistro(bdPremio, lSelect, lrst, pMensagem) Then
                        Call gpGraLog(1, pMensagem)
                        GoTo SAIDA
                    End If
                    lPremioAnt = 0
                    If Not lrst.EOF Then
                        If IsNull(lrst("soma")) Then
                            pMensagem = "AGE0022 - Cotação LMI não localizada."
                            GoTo SAIDA
                        End If
                        lPremioAnt = gfTratarNulo(lrst("soma"), tpoNumerico)
                        'Percentual de Comissão
                        Select Case lCod_Ramo
                            Case 112
                                lPer_Coml_Cota = gfTrunca(0.8 / (1 - (lrst("Comis") / 100)), 5)
                            Case 113, 114
                                lPer_Coml_Cota = gfTrunca(0.75 / (1 - (lrst("Comis") / 100)), 5)
                            Case 150
                                lPer_Coml_Cota = gfTrunca(0.85 / (1 - (lrst("Comis") / 100)), 5)
                            Case Else
                                lPer_Coml_Cota = gfTrunca(0.8 / (1 - (lrst("Comis") / 100)), 5)
                        End Select
                        
                        lobjstcCobertura.Per_Taxa = gfTrunca(lPremioAnt / lVal_IS * 100 / lPer_Coml_Cota / (1 - (lrst("Per_DEsc") / 100)), 5)
                        If lrst("Comis") <> lPer_Coml Or lrst("Per_DEsc") <> lCoef_Desc_Coml Then
                            lPremioAnt = 0

                        End If
                        'lobjstcCobertura.Per_Taxa = gfTrunca(lPremioAnt / lVal_IS * 100 / lPer_Coml_Cota, 5)
                        'If lrst("Comis") <> lPer_Coml Then 'Or lrst("Per_DEsc") <> lCoef_Desc_Coml Then
                        '    lobjstcCobertura.Per_Taxa = gfTrunca(lobjstcCobertura.Per_Taxa * mPer_Coml, 5)
                        '    lPremioAnt = gfTrunca(lVal_IS * (lobjstcCobertura.Per_Taxa / 100), 2)
                        'End If
                        
                        TxAlterada = True
                    End If
                End If
                lTaxaMin = lobjstcCobertura.Per_Taxa
                lTaxaComl = gfTrunca(lTaxaMin * mPer_Coml, 5)
                lTaxaFinal = gfTrunca(lTaxaComl * (1 - (lCoef_Desc_Coml / 100)), 5)
            End If
        End If
        
        'Consiste Taxa Final
        If lTaxaFinal = 0 Then
            pMensagem = "AGE0022 - Cobertura " & lCod_Cobert & " taxa não encontrada ou taxa zerada! Solicitar Cotação."
            GoTo SAIDA
        End If
        
        
        'Calcular Prêmio Anual
        lPremio = 0
        lPremio = gfTrunca(lVal_IS * (lTaxaFinal / 100), 2)
        If lPremioAnt <> 0 Then
            lPremio = lPremioAnt
        End If
        'Aplica desconto de LMI
        If gNovaPropostaLMI Then
            If pStcItem.ObjPai.Per_Desc_LMI > 0 Then
                'Debug.Print pStcItem.Num_Item & "-" & lCod_Cobert & "-" & lPremio
                lPremio = lPremio * ((100 - pStcItem.ObjPai.Per_Desc_LMI) / 100)
'                Debug.Print lPremio
            End If
        End If
        
        
        lobjstcCobertura.Val_Pr_Anu = lPremio
        If gCod_Dizer <> 0 Then
            Call mpCalcula_Endosso(pStcItem.ObjPai.Num_Apol_An, pStcItem.Num_Item_Ant, "COML", lPremio)
        End If
        
        If Not gfPreenchido(pStcItem.ObjPai.Num_Cotac_LMI) Then
            lPremio = lPremio * lCoePrazo
        End If
        lobjstcCobertura.VAL_PR = lPremio
        If gCod_Dizer = 0 Then
            If lPremio = 0 Then
                pMensagem = "AGE0051 -  Erro no cálculo do prêmio da cobertura " & lCod_Cobert & ". Solicitar cotação."
                GoTo SAIDA
            End If
        End If
        
        'Calcular Prêmio Tarifário
        'Resseguro-sandra
        If pStcItem.ObjPai.Dat_Ini_Vig >= 20090501 Then
            lobjstcCobertura.Val_Pr_Tarif_Anu = lobjstcCobertura.Val_Pr_Anu
        Else
            lPremio = gfTrunca(lVal_IS * (lTaxaReser / 100), 2)
            lobjstcCobertura.Val_Pr_Tarif_Anu = lPremio
        End If
        If gCod_Dizer <> 0 Then
            Call mpCalcula_Endosso(pStcItem.ObjPai.Num_Apol_An, pStcItem.Num_Item_Ant, "TARI", lPremio)
        End If
        
        'Resseguro-sandra
        If pStcItem.ObjPai.Dat_Ini_Vig >= 20090501 Then
            lobjstcCobertura.Val_Pr_Tarif = lobjstcCobertura.VAL_PR
        Else
            lPremio = lPremio * lCoePrazo
            lobjstcCobertura.Val_Pr_Tarif = lPremio
        End If
        
        'Calcular Prêmio Resseguro
        'Resseguro-sandra
        If pStcItem.ObjPai.Dat_Ini_Vig >= 20090501 Then
            lobjstcCobertura.Val_Pr_Resg_Anu = lobjstcCobertura.Val_Pr_Anu
        Else
            lPremio = gfTrunca(lVal_IS * (lTaxaRess / 100), 2)
            lobjstcCobertura.Val_Pr_Resg_Anu = lPremio
        End If
        If gCod_Dizer <> 0 Then
            Call mpCalcula_Endosso(pStcItem.ObjPai.Num_Apol_An, pStcItem.Num_Item_Ant, "RESG", lPremio)
        End If
        
        'Resseguro-sandra
        If pStcItem.ObjPai.Dat_Ini_Vig >= 20090501 Then
            lobjstcCobertura.Val_Pr_Resg = lobjstcCobertura.VAL_PR
        Else
            lPremio = lPremio * lCoePrazo
            lobjstcCobertura.Val_Pr_Resg = lPremio
        End If
        
        If Not TxAlterada Then
            lobjstcCobertura.Per_Taxa = lTaxaMin
            lobjstcCobertura.Formula = "CL" & lCod_Classe & ";Tx=" & lTaxaMin & ";Frq=" & lCoef_Franquia & ";Prot=" & lCoef_Protecao & ";Bon=" & lCoef_Bonus & ";PI=" & lCoef_PI & ";Cms=" & mPer_Coml & ";Lmg/VR=" & lCoefIs
            If lobjstcCobertura.Cod_Grupo_Cobertura <> 1 Then
                lobjstcCobertura.Formula = lobjstcCobertura.Formula & ";PROGR=" & lCoef_Desc_Prog
            End If
            lobjstcCobertura.Formula = lobjstcCobertura.Formula & ";DCOML=" & lCoef_Desc_Coml & ";TxFim=" & lTaxaFinal
            If gCod_Dizer <> 0 Then
                lobjstcCobertura.Formula = lobjstcCobertura.Formula & ";CoefPrazo=" & lCoePrazo
            End If
        Else
            lobjstcCobertura.Per_Taxa = lTaxaMin
            lobjstcCobertura.Formula = "Taxa Alterada Manualmente=" & lTaxaMin & "; Cms=" & mPer_Coml & ";DCOML=" & lCoef_Desc_Coml & "; TxFim=" & lTaxaFinal
        End If
        If lobjstcCobertura.Per_Taxa > 99 Then lobjstcCobertura.Per_Taxa = 99
        If Not lobjCobert.mfAlterarValoresPrecificacao(lobjstcCobertura, pPropTabela, pNomeTabImport) Then
            pMensagem = lobjCobert.UltimoErro
            GoTo SAIDA
        End If
        
ProximaCobertura:
        'Adiciono a cobertura em uma nova coleção
        Set lobjstcCobertura.ObjPai = pStcItem
        lColecaoTemp.Add lobjstcCobertura, lobjstcCobertura.Class_ID
    
    Next lobjstcCobertura
    
    mfCalNewPrecif = True
    Set pStcItem.ColecaoPedCobert = Nothing
    Set pStcItem.ColecaoPedCobert = lColecaoTemp
    
    pStcItem.Per_Desc_Progres = lCoef_Desc_Prog
    If pStcItem.DAT_DIGIT = 0 Then
        pStcItem.DAT_DIGIT = Left(gfForData(3), 8)
        pStcItem.Cod_User_Liber = gNomUsuario6
    End If
    
    Set lObjPedLoc = New clsA46V708A
   
    If Not lObjPedLoc.mfAlterar(pStcItem, pPropTabela, pNomeTabImport) Then
        pMensagem = lObjPedLoc.UltimoErro
        GoTo SAIDA
    End If
   
SAIDA:
    GoSub FechaObj
    Exit Function
   
FechaObj:
    Set lobjstcCobertura = Nothing
    Set lobjstcCorretor = Nothing
    Set lbdTab_Tarifa = Nothing
    Set lrsTab_Tarifa = Nothing
    Set lObjPedLoc = Nothing
    Set lobjCobert = Nothing
    Set lColecaoTemp = Nothing
    Set bdPremio = Nothing
    Set lrst = Nothing
    Return
End Function

Private Function mpFindCoefLMG(ByRef pCoef_LMG As Double, ByRef pMensagem As String) As Boolean
    Dim lbdTbm_Fator_LMG    As ADODB.Connection         'Arquivo: P00Tarifa.
    Dim lrsTbm_Fator_LMG    As ADODB.Recordset          'Registro : Tbm_Fator_LMG.
    Dim mCodSequencia As Integer
    Dim lSelect             As String
    
    
    pMensagem = ""
    pCoef_LMG = 0
    mpFindCoefLMG = False
    
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTbm_Fator_LMG, pMensagem) = False Then
        Exit Function
    End If
    
    mCodSequencia = lCod_Sequencia
    
NovaSequencia:
    
    '1 - Procura Registro na Tabela de Fator de Ajuste por Limite Máximo de Garantia
    ' Exceção: Faz busca exata
    
    lSelect = " Select Fator_Ajuste as Fator, Lmg_minimo, Lmg_maximo"
    lSelect = lSelect & " from Tbm_Fator_LMG "
    lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & lCod_Modalidade
    lSelect = lSelect & " and Cod_Tip_Reg = 'E' "
    lSelect = lSelect & " and Cod_Clasf = '" & lCod_Clasf & "'"
    lSelect = lSelect & " and Cod_Compl_Clasf = '" & lCod_Compl_Clasf & "'"
    lSelect = lSelect & " and Cod_Cobert = " & lCod_Cobert
    lSelect = lSelect & " and Cod_Construcao = " & lCod_Construcao
    lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
    lSelect = lSelect & " and Cod_Sequencia <= " & mCodSequencia
    lSelect = lSelect & " order by Lmg_minimo desc "

    If gfObtRegistro(lbdTbm_Fator_LMG, lSelect, lrsTbm_Fator_LMG, pMensagem) = False Then
        Call gpFecha2(lbdTbm_Fator_LMG)
        Exit Function
    End If
    
    '2 - Se não encontrar vai procurar por Padrão na mesma Sequencia
    If lrsTbm_Fator_LMG.EOF Then
       lrsTbm_Fator_LMG.Close
       GoTo BuscaCoefPadraoLMG
    Else
        Do While Not lrsTbm_Fator_LMG.EOF
            If lVal_IS >= Check(lrsTbm_Fator_LMG("Lmg_minimo")) And _
                lVal_IS <= Check(lrsTbm_Fator_LMG("Lmg_maximo")) Then
                pCoef_LMG = Check(lrsTbm_Fator_LMG("Fator"))
                Exit Do
            End If
            lrsTbm_Fator_LMG.MoveNext
        Loop
        lrsTbm_Fator_LMG.Close
        GoTo FimCoef
    End If
    
BuscaCoefPadraoLMG:
    
    '3 - Procura Registro na Tabela de Fator de Ajuste por Limite Máximo de Garantia
    ' Padrão : Busca pela Cobertura
    
    lSelect = " Select Fator_Ajuste as Fator, Lmg_minimo, Lmg_maximo"
    lSelect = lSelect & " from Tbm_Fator_LMG "
    lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & lCod_Modalidade
    lSelect = lSelect & " and Cod_Tip_Reg = 'D' "                  'Default
    lSelect = lSelect & " and Cod_Cobert = " & lCod_Cobert
    lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
    lSelect = lSelect & " and Cod_Sequencia <= " & mCodSequencia
    lSelect = lSelect & " order by Lmg_minimo desc "

    If gfObtRegistro(lbdTbm_Fator_LMG, lSelect, lrsTbm_Fator_LMG, pMensagem) = False Then
        Call gpFecha2(lbdTbm_Fator_LMG)
        Exit Function
    End If

    '4 - Se não encontrar a sequencia vai procurar sequencia anterior
    If lrsTbm_Fator_LMG.EOF Then
        If mCodSequencia > 0 Then
            mCodSequencia = mCodSequencia - 1
            lrsTbm_Fator_LMG.Close
            GoTo NovaSequencia
        End If
    Else
        Do While Not lrsTbm_Fator_LMG.EOF
            If lVal_IS >= Check(lrsTbm_Fator_LMG("Lmg_minimo")) And _
                lVal_IS <= Check(lrsTbm_Fator_LMG("Lmg_maximo")) Then
                pCoef_LMG = Check(lrsTbm_Fator_LMG("Fator"))
                Exit Do
            End If
            lrsTbm_Fator_LMG.MoveNext
        Loop
        lrsTbm_Fator_LMG.Close
    End If
    
FimCoef:
    
    Call gpFecha2(lbdTbm_Fator_LMG)
    mpFindCoefLMG = True
    '3 - Vai procurar na tabela de Ajustes pontuais
    If mfAjustesPontuais(e_Tip_Valor_CoefLMG, TpoValor, pCoef_LMG, pMensagem) = False Then
       Exit Function
    End If

End Function

Public Function mfAjustesPontuais(ByVal pCod_Tip_Valor As Integer, ByVal pTip_Dado As TipoDados, ByRef ValorAjustavel As Variant, ByVal pMensagem As String, Optional ByRef pobjstcPedLoc As stcA46V708B = Nothing) As Boolean
    Dim lbdTbm_Parameter    As ADODB.Connection         'Arquivo: P00Tarifa.
    Dim lrsTbm_Parameter    As ADODB.Recordset          'Registro : Tbm_Parameter.
    Dim mCodSequencia As Integer
    Dim lSelect             As String
    
    '1100192 - Tarifa 1006 e 1007 Marco 2011
    Dim lValCoef            As Double
    
    On Error GoTo ControleErros
    
    pMensagem = ""
    mfAjustesPontuais = False
    
    '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 1
    If Not pobjstcPedLoc Is Nothing Then
        lCod_Ramo = pobjstcPedLoc.Cod_Ramo
        lCod_Modalidade = pobjstcPedLoc.Cod_Produto
        If Not pobjstcPedLoc.ObjPai Is Nothing Then
            lCod_Corretor = pobjstcPedLoc.ObjPai.objstcCtrlEmis.Cod_Corr
            mCodSequencia = pobjstcPedLoc.ObjPai.Cod_Sequencia
            lCod_Tarifa = pobjstcPedLoc.ObjPai.Cod_Produto
        Else
            lCod_Tarifa = pobjstcPedLoc.Cod_Iden_Tarif
        End If
        lCod_Clasf = pobjstcPedLoc.Cod_Clasf
        lCod_Compl_Clasf = pobjstcPedLoc.Cod_Compl_Clasf
        lCod_Construcao = pobjstcPedLoc.Cod_Constr
    End If
    
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTbm_Parameter, pMensagem) = False Then
        Exit Function
    End If
    
    mCodSequencia = lCod_Sequencia
    
NovaSequencia:
    
    If mCodSequencia > 0 Then
        'Procurar se existem registros para a atual sequencia
        'Se existir pelo menos 1, essa é a sequencia válida
        
        lSelect = " Select count(*) as Total  from Tbm_Parameter "
        lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & lCod_Modalidade
        lSelect = lSelect & " and Cod_Sequencia <= " & mCodSequencia
        
        If gfObtRegistro(lbdTbm_Parameter, lSelect, lrsTbm_Parameter, pMensagem) = False Then
            If mCodSequencia > 0 Then
                mCodSequencia = mCodSequencia - 1
                lrsTbm_Parameter.Close
                GoTo NovaSequencia
            End If
        End If
        
        If Check(lrsTbm_Parameter("Total")) = 0 Then
            If mCodSequencia > 0 Then
                mCodSequencia = mCodSequencia - 1
                lrsTbm_Parameter.Close
                GoTo NovaSequencia
            End If
        End If
        
        lrsTbm_Parameter.Close
    End If
    
    lSelect = " Select Cod_Cobert, Cod_Tip_Chave, Chave_Ini, Chave_Ter, Valor from Tbm_Parameter "
    lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & lCod_Modalidade
    lSelect = lSelect & " and Cod_Corr in (0, " & lCod_Corretor & ")"
    lSelect = lSelect & " and ((rtrim(ltrim(Cod_Clasf)) = '" & Trim(lCod_Clasf) & "'"
    lSelect = lSelect & " and Cod_Compl_Clasf = '" & lCod_Compl_Clasf & "') or rtrim(ltrim(Cod_Clasf)) = '0' )"
    lSelect = lSelect & " and Cod_Cobert in (0," & lCod_Cobert & ")"
    lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
    lSelect = lSelect & " and Cod_Sequencia <= " & mCodSequencia
    lSelect = lSelect & " and Cod_Tip_Valor = " & pCod_Tip_Valor
    lSelect = lSelect & " and Cod_Constr in (0," & lCod_Construcao & ")"
    lSelect = lSelect & " order by Cod_Cobert, Cod_Prioridade, Cod_Constr desc "

    If gfObtRegistro(lbdTbm_Parameter, lSelect, lrsTbm_Parameter, pMensagem) = False Then
        Call gpFecha2(lbdTbm_Parameter)
        Exit Function
    End If
    
    '1100192 - Tarifa 1006 e 1007 Marco 2011
    lValCoef = 1
    Do While Not lrsTbm_Parameter.EOF
        If CDblx(Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 1 Then
            If CDblx(Check(lrsTbm_Parameter("Chave_Ini"))) <= lVal_IS And CDblx(Check(lrsTbm_Parameter("Chave_Ter"))) >= lVal_IS Then
                ValorAjustavel = gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        If CDblx(Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 2 Then
            If CDblx(Check(lrsTbm_Parameter("Chave_Ini"))) <= lVal_Franquia And CDblx(Check(lrsTbm_Parameter("Chave_Ter"))) >= lVal_Franquia Then
                ValorAjustavel = gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        If CDblx(Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 3 Then
            If CDblx(Check(lrsTbm_Parameter("Chave_Ini"))) <= lNum_Cep And CDblx(Check(lrsTbm_Parameter("Chave_Ter"))) >= lNum_Cep Then
                ValorAjustavel = gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
                If (pCod_Tip_Valor = e_Tip_Valor_ClasAceit) And _
                   (lCod_Cobert = 33 Or lCod_Cobert = 55) And _
                   (Val(ValorAjustavel) = 3) Then
                        pMensagem = "Região restrita para Alagamento."
                End If
            End If
        End If
        If CDblx(Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 4 Then
            If Check(lrsTbm_Parameter("Chave_Ini")) = lTip_Emissao And CDblx(Check(lrsTbm_Parameter("Chave_Ter"))) = lTip_Emissao Then
                ValorAjustavel = gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        If CDblx(Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 5 Then
            If CDblx(Check(lrsTbm_Parameter("Chave_Ini"))) <= lPer_Coml And CDblx(Check(lrsTbm_Parameter("Chave_Ter"))) >= lPer_Coml Then
                ValorAjustavel = gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        If CDblx(Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 6 Then
            If CDblx(Check(lrsTbm_Parameter("Chave_Ini"))) <= lCoef_Desc_Coml And CDblx(Check(lrsTbm_Parameter("Chave_Ter"))) >= lCoef_Desc_Coml Then
                ValorAjustavel = gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        If CDblx(Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 7 Then
            '1100192 - Tarifa 1006 e 1007 Marco 2011
            If CDblx(Check(lrsTbm_Parameter("Chave_Ini"))) <= lVal_VR_Item And CDblx(Check(lrsTbm_Parameter("Chave_Ter"))) >= lVal_VR_Item Then
                ValorAjustavel = gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        '1001146 - Tarifa 1000 e 1001 Dezembro 2010
        If CDblx(Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 8 Then 'And lIndLMI = 1 Then
            If CDblx(Check(lrsTbm_Parameter("Chave_Ini"))) <= lCodRegiao And CDblx(Check(lrsTbm_Parameter("Chave_Ter"))) >= lCodRegiao Then
                ValorAjustavel = gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If

        '1200539 - Tarifa 1038 e 1039 Junho 2012
        If CDblx(Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 9 Then
            ValorAjustavel = gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
        End If

        '1100192 - Tarifa 1006 e 1007 Marco 2011
        If pCod_Tip_Valor = e_Tip_Valor_CoeficienteTaxa And ValorAjustavel <> 0 Then
            lValCoef = lValCoef * ValorAjustavel
        End If
        lrsTbm_Parameter.MoveNext
    Loop
    lrsTbm_Parameter.Close
    '1100192 - Tarifa 1006 e 1007 Marco 2011
    If pCod_Tip_Valor = e_Tip_Valor_CoeficienteTaxa Then
        ValorAjustavel = lValCoef
    End If
    mfAjustesPontuais = True
    Call gpFecha2(lbdTbm_Parameter)
    
    Exit Function
ControleErros:
    mfAjustesPontuais = False
    pMensagem = "mfAjustesPontuais ERRO - " & Err & " " & Err.Description & _
        vbCrLf & " SQL: " & lSelect
    Call gpGraLog(0, pMensagem)

End Function


Private Function mpFindCoefPI(ByRef pCoef_PI As Double, ByVal pNum_Meses_PI, ByRef pMensagem As String) As Boolean
Dim lbdTbm_Coef_PI      As ADODB.Connection         'Arquivo: P00Multi.
Dim lMensagem           As String                   'Mensagem.
Dim lrsTbm_Coef_PI      As ADODB.Recordset          'Registro : Tbm_Coef_PI.
Dim lSelect             As String

    pCoef_PI = 1
    mpFindCoefPI = False
    
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTbm_Coef_PI, lMensagem) = False Then
        Exit Function
    End If
    
    '1 - Procura pela Cobertura na Tabela de Coeficiente
    lSelect = " Select *  from Tbm_Coef_PI "
    lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & lCod_Cobert
    lSelect = lSelect & " and Num_Meses = " & pNum_Meses_PI
    lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
    lSelect = lSelect & " and Cod_Sequencia <= " & lCod_Sequencia

    If gfObtRegistro(lbdTbm_Coef_PI, lSelect, lrsTbm_Coef_PI, lMensagem) = False Then
        Call gpFecha2(lbdTbm_Coef_PI)
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
     mpFindCoefPI = True
    '2 - Se encontrar permite cadastro de Período Indenitário
    If Not lrsTbm_Coef_PI.EOF Then
        If lrsTbm_Coef_PI("Tip_Agravacao") = "S" Then
            pCoef_PI = Check(lrsTbm_Coef_PI("TAXA_COEF_PI"))
        End If
    End If
    
    Call gpFecha2(lbdTbm_Coef_PI)
    
End Function

Private Function mpFindCoefProtecao(ByRef pCoef_Protecao As Double, ByRef pCoef_ProtecaoResseg As Double, ByVal pColecaoPedFator As Collection, ByRef pMensagem As String) As Boolean
    Dim lbdTbm_Fator_Protecao    As ADODB.Connection         'Arquivo: P00Tarifa.
    Dim lrsTbm_Fator_Protecao    As ADODB.Recordset          'Registro : Tbm_Fator_Protecao.
    Dim lobjstcFator        As stcA46V720B       'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lSelect                 As String
    
    Dim mCodSequencia As Integer
    pMensagem = ""
    pCoef_Protecao = 0
    pCoef_ProtecaoResseg = 0
    mpFindCoefProtecao = False
        
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTbm_Fator_Protecao, pMensagem) = False Then
        Exit Function
    End If
            
    'Vai procurar para todos os fatores cadastrados se existem Agravação para a cobertura
    mCodSequencia = lCod_Sequencia
    For Each lobjstcFator In pColecaoPedFator
NovaSequencia:
        '1 - Procura Registro na Tabela de Fator de Ajuste por Protecao
        ' Exceção: Faz busca exata
    
        lSelect = " Select Fator_Ajuste as Fator, Tipo_Resseguro from Tbm_Fator_Protecao "
        lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & lCod_Modalidade
        lSelect = lSelect & " and Cod_Tip_Reg = 'E' "                      'Exceção
        lSelect = lSelect & " and Cod_Clasf = '" & lCod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & lCod_Compl_Clasf & "'"
        lSelect = lSelect & " and Cod_Construcao = " & lCod_Construcao
        lSelect = lSelect & " and Cod_Cobert = " & lCod_Cobert
        lSelect = lSelect & " and Cod_Fator = " & lobjstcFator.Cod_Fator
        lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
        lSelect = lSelect & " and Cod_Sequencia <= " & mCodSequencia

        If gfObtRegistro(lbdTbm_Fator_Protecao, lSelect, lrsTbm_Fator_Protecao, pMensagem) = False Then
            Call gpFecha2(lbdTbm_Fator_Protecao)
            Exit Function
        End If
        
        '2 - Se não encontrar vai procurar por Padrão na mesma Sequencia
        If lrsTbm_Fator_Protecao.EOF Then
            lrsTbm_Fator_Protecao.Close
            GoTo BuscaCoefPadrao
        Else
            pCoef_Protecao = pCoef_Protecao + Check(lrsTbm_Fator_Protecao("Fator"))
            If Check(lrsTbm_Fator_Protecao("Tipo_Resseguro")) = "S" Then
                pCoef_ProtecaoResseg = pCoef_ProtecaoResseg + Check(lrsTbm_Fator_Protecao("Fator"))
            End If
            lrsTbm_Fator_Protecao.Close
            GoTo ProximaProtecao
        End If

BuscaCoefPadrao:
        '3 - Procura Registro na Tabela de Fator de Ajuste por Protecao
        ' Padrão : Busca pela Fator e Cobertura
    
        lSelect = " Select Fator_Ajuste as Fator, Tipo_Resseguro from Tbm_Fator_Protecao "
        lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & lCod_Modalidade
        lSelect = lSelect & " and Cod_Tip_Reg = 'D' "                  'Default
        lSelect = lSelect & " and Cod_Cobert = " & lCod_Cobert
        lSelect = lSelect & " and Cod_Fator = " & lobjstcFator.Cod_Fator
        lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
        lSelect = lSelect & " and Cod_Sequencia <= " & mCodSequencia

        If gfObtRegistro(lbdTbm_Fator_Protecao, lSelect, lrsTbm_Fator_Protecao, pMensagem) = False Then
            Call gpFecha2(lbdTbm_Fator_Protecao)
            Exit Function
        End If
        
        '4 - Se não encontrar, deverá buscar novamente a exceção com
        'sequencia anterior
        If lrsTbm_Fator_Protecao.EOF Then
            If mCodSequencia > 0 Then
                mCodSequencia = mCodSequencia - 1
                lrsTbm_Fator_Protecao.Close
                GoTo NovaSequencia
            End If
        Else
            pCoef_Protecao = pCoef_Protecao + Check(lrsTbm_Fator_Protecao("Fator"))
            If Check(lrsTbm_Fator_Protecao("Tipo_Resseguro")) = "S" Then
                pCoef_ProtecaoResseg = pCoef_ProtecaoResseg + Check(lrsTbm_Fator_Protecao("Fator"))
            End If
        End If
        lrsTbm_Fator_Protecao.Close
        
ProximaProtecao:
    Next lobjstcFator
    

    Call gpFecha2(lbdTbm_Fator_Protecao)
    mpFindCoefProtecao = True
    
    '5 - Vai procurar na tabela de Ajustes pontuais
    If mfAjustesPontuais(e_Tip_Valor_CoefProt, TpoValor, pCoef_Protecao, pMensagem) = False Then
       Exit Function
    End If

End Function
Private Function mpFindCoefRegiao(ByRef pCoef_Regiao As Double, ByRef pMensagem As String) As Boolean
    Dim lbdTbm_Fator_Cep    As ADODB.Connection         'Arquivo: P00Tarifa.
    Dim lrsTbm_Fator_Cep    As ADODB.Recordset          'Registro : Tbm_Fator_Cep.
    Dim mCodSequencia As Integer
    Dim lSelect                 As String
    
    pMensagem = ""
    pCoef_Regiao = 0
    mpFindCoefRegiao = False
    
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTbm_Fator_Cep, pMensagem) = False Then
        Exit Function
    End If
    
    mCodSequencia = lCod_Sequencia
    
NovaSequencia:
    
    '1 - Procura Registro na Tabela de Fator de Ajuste por CEP

    lSelect = " Select Fator_Ajuste as Fator, Num_Cep_Inicial, Num_Cep_Final"
    lSelect = lSelect & " from Tbm_Fator_Cep "
    lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & lCod_Modalidade
    lSelect = lSelect & " and Cod_Cobert = " & lCod_Cobert
    lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
    lSelect = lSelect & " and Cod_Sequencia <= " & mCodSequencia
    lSelect = lSelect & " order by Num_Cep_Inicial "

    If gfObtRegistro(lbdTbm_Fator_Cep, lSelect, lrsTbm_Fator_Cep, pMensagem) = False Then
        Call gpFecha2(lbdTbm_Fator_Cep)
        Exit Function
    End If
    
    '2 - Se não encontrar a sequencia vai procurar sequencia anterior
    If lrsTbm_Fator_Cep.EOF Then
        If mCodSequencia > 0 Then
            mCodSequencia = mCodSequencia - 1
            lrsTbm_Fator_Cep.Close
            GoTo NovaSequencia
        End If
    Else
    '3 - Se encontrar Coeficiente para essa cobertura vai verificar se está dentro da faixa
        Do While Not lrsTbm_Fator_Cep.EOF
            If lNum_Cep >= Check(lrsTbm_Fator_Cep("Num_Cep_Inicial")) And _
               lNum_Cep <= Check(lrsTbm_Fator_Cep("Num_Cep_Final")) Then
                pCoef_Regiao = Check(lrsTbm_Fator_Cep("Fator"))
                Exit Do
            End If
            lrsTbm_Fator_Cep.MoveNext
        Loop
    End If

    lrsTbm_Fator_Cep.Close
     
    '4 - Vai procurar na tabela de Ajustes pontuais
    If mfAjustesPontuais(e_Tip_Valor_CoefRegiao, TpoValor, pCoef_Regiao, pMensagem) = False Then
       Exit Function
    End If
    
    
    mpFindCoefRegiao = True
    Call gpFecha2(lbdTbm_Fator_Cep)
    
End Function

Public Function gfVerParametro(ByVal pTipoValor As Integer, _
                                ByRef pobjstcPed As stcA46V702B, ByRef pobjstcLoc As stcA46V708B, _
                                ByRef pobjstcCobert As stcA46V704B, _
                                ByRef pRetorno As String, _
                               ByRef pMensagem As String) As Boolean
    'Função #071: verifica parametrizações do produto.

    'Parâmetros de saída.....pRetorno...........retorno.
    '                        pMensagem..........mensagem.

    Dim bdParametro         As ADODB.Connection     'P0044800: base de dados.
    Dim lSelect             As String               'Comando SELECT.
    Dim rsParm          As ADODB.Recordset      'P0044800: tabela.
    Dim lParametro          As Variant              'Array com os valores
    Dim lBlnContinuar       As Boolean              'Controla a validação
    Dim lCodProtocolo       As Long                 'Código do protocolo
    Dim lCodSequencia       As Long                 'Sequencia.
    Dim lProxSequencia      As Long
        

    gfVerParametro = False
    
    
    
    '1. Posiciona controles.
    pRetorno = " "
    '2. Carrego os parametros
    If Not mfCarregarParametros(lParametro, pMensagem, pobjstcPed, pobjstcLoc, pobjstcCobert) Then Exit Function
    '0  -   Tarifa Básica
    '1  -   Código da unidade
    '2  -   Código do corretor
    '3  -   Parcela
    '4  -   Pagamento antecipado?
    '5  -   Cobertura
    '6  -   Contém proteção Sprinklers ?
    '7  -   Modalidade
    
    'Estas abaixo não estarão carregadas no lParametro
    '5  -   Maior IS entre coberturas
    '1XX  - IS da cobertura XX
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdParametro, pMensagem) = False Then
        Exit Function
    End If

    lCodSequencia = gCodSequencia
    lCodProtocolo = gCodProtocolo
    If lCodSequencia > 0 Then
Voltar:
        lSelect = "SELECT top 1 * from TAB_PARM_DIVERSOS where " & _
              "Cod_Produto = " & gCodProduto & " " & _
              "AND Cod_Protocolo = " & lCodProtocolo & " " & _
              "AND Cod_Sequencia <= " & lCodSequencia
        If gfObtRegistro(bdParametro, lSelect, rsParm, pMensagem) = False Then
            Call gpFecha2(bdParametro)
            Exit Function
        End If
        If rsParm.EOF Then
            If lCodSequencia > 0 Then
                lCodSequencia = lCodSequencia - 1
                GoTo Voltar
            Else
                If lCodProtocolo > 0 Then
                    If lCodProtocolo > 10 Then
                        lCodProtocolo = 0
                    Else
                        lCodProtocolo = lCodProtocolo - 1
                    End If
                    lCodSequencia = gCodSequencia
                    GoTo Voltar
                End If
            End If
        End If
    End If

    lSelect = "SELECT * " & _
              "FROM TAB_PARM_DIVERSOS inner join TAB_TIP_CHAVE " & _
              "On TAB_PARM_DIVERSOS.Tipo_Chave = TAB_TIP_CHAVE.Tip_Chave " & _
              "WHERE Cod_Produto = " & gCodProduto & " " & _
              "AND Cod_Protocolo = " & lCodProtocolo & " " & _
              "AND Cod_Sequencia <= " & lCodSequencia & _
              " AND Tipo_Valor = " & pTipoValor & " "
              
    lSelect = lSelect & " and cod_ramo = " & gCod_Ramo
    If pTipoValor = 7 Then
        lSelect = lSelect & " ORDER BY Nivel_Prioridade asc"
    Else
        lSelect = lSelect & " ORDER BY Nivel_Prioridade desc"
    End If
    
    If gfObtRegistro(bdParametro, lSelect, rsParm, pMensagem) = False Then
        Call gpFecha2(bdParametro)
        Exit Function
    End If
    '4. Verifica registros.
    While Not rsParm.EOF
        lBlnContinuar = False
        
        
        If rsParm("Ope_Valor1") = 1 Then        '=
            If rsParm("Chave_1") = CStr(lParametro(rsParm("Cod_Valor1"))) Then
                lBlnContinuar = True
            Else
                lBlnContinuar = False
            End If
        ElseIf rsParm("Ope_Valor1") = 2 Then    '>
            If CSng(lParametro(rsParm("Cod_Valor1"))) > CSng(rsParm("Chave_1")) Then
                lBlnContinuar = True
            Else
                lBlnContinuar = False
            End If
        ElseIf rsParm("Ope_Valor1") = 3 Then    '<
            If CSng(lParametro(rsParm("Cod_Valor1"))) < CSng(rsParm("Chave_1")) Then
                lBlnContinuar = True
            Else
                lBlnContinuar = False
            End If
        ElseIf rsParm("Ope_Valor1") = 4 Then    '>=
            If CSng(lParametro(rsParm("Cod_Valor1"))) >= CSng(rsParm("Chave_1")) Then
                lBlnContinuar = True
            Else
                lBlnContinuar = False
            End If
        ElseIf rsParm("Ope_Valor1") = 5 Then    '<=
            If CSng(lParametro(rsParm("Cod_Valor1"))) <= CSng(rsParm("Chave_1")) Then
                lBlnContinuar = True
            Else
                lBlnContinuar = False
            End If
        ElseIf rsParm("Ope_Valor1") = 6 Then    '<>
            If rsParm("Chave_1") <> CStr(lParametro(rsParm("Cod_Valor1"))) Then
                lBlnContinuar = True
            Else
                lBlnContinuar = False
            End If
        End If
        
        If lBlnContinuar = True Then
            If rsParm("Ope_Valor2") = 1 Then        '=
                If rsParm("Chave_2") = CStr(lParametro(rsParm("Cod_Valor2"))) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor2") = 2 Then    '>
                If CSng(lParametro(rsParm("Cod_Valor2"))) > CSng(rsParm("Chave_2")) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor2") = 3 Then    '<
                If CSng(lParametro(rsParm("Cod_Valor2"))) < CSng(rsParm("Chave_2")) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor2") = 4 Then    '>=
                If CSng(lParametro(rsParm("Cod_Valor2"))) >= CSng(rsParm("Chave_2")) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor2") = 5 Then    '<=
                If CSng(lParametro(rsParm("Cod_Valor2"))) <= CSng(rsParm("Chave_2")) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor2") = 6 Then    '<>                                        '<>
                If rsParm("Chave_2") <> CStr(lParametro(rsParm("Cod_Valor2"))) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            End If
        End If

        If lBlnContinuar = True Then
            If rsParm("Ope_Valor3") = 1 Then        '=
                If rsParm("Chave_3") = CStr(lParametro(rsParm("Cod_Valor3"))) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor3") = 2 Then    '>
                If CSng(lParametro(rsParm("Cod_Valor3"))) > CSng(rsParm("Chave_3")) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor3") = 3 Then    '<
                If CSng(lParametro(rsParm("Cod_Valor3"))) < CSng(rsParm("Chave_3")) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor3") = 4 Then    '>=
                If CSng(lParametro(rsParm("Cod_Valor3"))) >= CSng(rsParm("Chave_3")) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor3") = 5 Then    '<=
                If CSng(lParametro(rsParm("Cod_Valor3"))) <= CSng(rsParm("Chave_3")) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor3") = 6 Then    '<>                                        '<>
                If rsParm("Chave_3") <> CStr(lParametro(rsParm("Cod_Valor3"))) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            End If
        End If

        If lBlnContinuar = True Then
            If rsParm("Ope_Valor4") = 1 Then        '=
                If rsParm("Chave_4") = CStr(lParametro(rsParm("Cod_Valor4"))) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor4") = 2 Then    '>
                If CSng(lParametro(rsParm("Cod_Valor4"))) > CSng(rsParm("Chave_4")) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor4") = 3 Then    '<
                If CSng(lParametro(rsParm("Cod_Valor4"))) < CSng(rsParm("Chave_4")) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor4") = 4 Then    '>=
                If CSng(lParametro(rsParm("Cod_Valor4"))) >= CSng(rsParm("Chave_4")) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor4") = 5 Then    '<=
                If CSng(lParametro(rsParm("Cod_Valor4"))) <= CSng(rsParm("Chave_4")) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            ElseIf rsParm("Ope_Valor4") = 6 Then    '<>                                        '<>
                If rsParm("Chave_4") <> lParametro(rsParm("Cod_Valor4")) Then
                    lBlnContinuar = True
                Else
                    lBlnContinuar = False
                End If
            End If
        End If
        
        If lBlnContinuar = True Then
            pRetorno = rsParm!Retorno
            '7-(Limite de Juros) 10-(Desconto)
            If (pTipoValor <> 7) And pTipoValor <> 10 Then  '1100705 - Tarifa 1016 e 1017 Agosto 2011
                GoTo Fim
            End If
        End If
        rsParm.MoveNext
    Wend
Fim:
    Call gpFecha3(rsParm, bdParametro)
    gfVerParametro = True
End Function

Private Function mfCarregarParametros(ByRef pParametros As Variant, pMensagem As String, _
                    pobjstcPed As stcA46V702B, pobjstcLoc As stcA46V708B, _
                    ByRef pobjstcCobert As stcA46V704B) As Boolean
'Procedure #082: Carrega os parâmetros para a função VER_PARAMETRO
        
        ' pParametros  Matriz com os parametros
                '0  -   Tarifa Básica
                '1  -   Código da unidade
                '2  -   Código do corretor
                '3  -   Parcela
                '4  -   Pagamento antecipado?
                '5  -   Cobertura
                '6  -   Contém proteção Sprinklers ?
                '7  -   Modalidade

Dim lTarifaBasica       As Long     'Tarifa Básica.
Dim lCodUnidade         As Long     'Código da Unidade.
Dim lCodCorretora       As Long     'Código do corretor.
Dim lNumeroParcela      As Integer  'Número de parcelas.
Dim lPagAntecipado      As Integer  'Indicador de pagamento antecipado.
Dim lCodCobertura       As Integer  'Código da cobertura
Dim lSprinklers         As Integer  'Contém sprinklers.
Dim lModalidade         As Long     'Modalidade
Dim lValLMG             As Double   'Valor LMG
Dim lRubrica            As String   'Classificação
Dim lCodSegurado        As String
Dim lCodConstr          As e_CodContr
Dim lobjstcFator        As stcA46V720B          'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
Dim lobjstcCobertura    As stcA46V704B
Dim lSinistro           As String '1100705 - Tarifa 1016 e 1017 Agosto 2011

'1001224 - Condições Diferenciada - Seguratec
Dim lDesconto           As Double
Dim lobjstcCorretor     As stcA46V705B
   
   
    mfCarregarParametros = False
    '0. Tarifa Basica
    lTarifaBasica = "0"
    
    '1. Código da Unidade
    lCodUnidade = pobjstcPed.Cod_Unid_Prod
        
    '2. Código do Corretor
    lCodCorretora = pobjstcPed.objstcCtrlEmis.Cod_Corr
        
    '3. Número de parcelas
    lNumeroParcela = pobjstcPed.Qtd_Parc
        
    '4 - PagamentoAntecipado?
    lPagAntecipado = pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec
    
    '5 - Cobertura
    lCodCobertura = pobjstcCobert.Cod_Cobert
    
    '6 - Contém Sprinklers?
    lSprinklers = 2
    For Each lobjstcFator In pobjstcLoc.ColecaoPedFator
        If gCod_Ramo = e_CodRamo_Empresarial Or gCod_Ramo = e_CodRamo_Roubos Or gCod_Ramo = e_CodRamo_Riscos_diversos Then
            If lobjstcFator.Cod_Fator = 4 Then
                lSprinklers = 1
                Exit For
            End If
        ElseIf gCod_Ramo = e_CodRamo_Condominio Then
            If lobjstcFator.Cod_Fator = 8 Then
                lSprinklers = 1
                Exit For
            End If
        End If
    Next lobjstcFator
           
    '7  -   Modalidade
    lModalidade = pobjstcLoc.Cod_Plano
    
    '8  -   Rubrica
    lRubrica = Val(Replace(pobjstcLoc.Cod_Clasf, "A", "999") & pobjstcLoc.Cod_Compl_Clasf)
    
    '9  -   LMG
    lValLMG = pobjstcCobert.Val_IS
    
    '10 - Codigo de construção
    lCodConstr = pobjstcLoc.Cod_Constr
    
    '11 - Código do segurado - (1001224 - Condições Diferenciada - Seguratec)
    lCodSegurado = Val(pobjstcPed.Cod_Segurado)
    
    '12 - Desconto - (1001224 - Condições Diferenciada - Seguratec)
    lDesconto = pobjstcCobert.Per_Desc_Coml
    
    '13 -   Número da Parcela
    lNumeroParcela = pobjstcPed.Qtd_Parc
    
    '18 - cod segurado
    lCodSegurado = pobjstcPed.Cod_Segurado
    
    '22 -   Pagamento antecipado?
    lPagAntecipado = pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec
    
    '23 - Sinistro?
    lSinistro = pobjstcLoc.Tip_Emissao & pobjstcLoc.Ind_Sinistro '1100705 - Tarifa 1016 e 1017 Agosto 2011
    
    '26. Carrego o array
    pParametros = Array(lTarifaBasica, lCodUnidade, lCodCorretora, _
                        lNumeroParcela, lPagAntecipado, lCodCobertura, _
                        lSprinklers, lModalidade, lRubrica, lValLMG, _
                        lCodConstr, lCodSegurado, lDesconto, _
                        lNumeroParcela, 0, 0, 0, 0, lCodSegurado, 0, 0, 0, _
                        lPagAntecipado, lSinistro)  '1100705 - Tarifa 1016 e 1017 Agosto 2011
                        '0,1,2
                        '3,4,5
                        '6,7,8,9
                        '10,11,12
                        '13,14,15,16,17,18,19,20,21
                        '22
    mfCarregarParametros = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcFator = Nothing
    Set lobjstcCobertura = Nothing
    Set lobjstcCorretor = Nothing
    Return
End Function

Private Function mfIncluirClausula(ByVal pPropTabela As String, _
                                   ByVal pNomeTabImport As String, _
                                   ByRef pClausula As Integer, _
                                   ByRef pobjstcItem As stcA46V708B, _
                                   ByRef pMensagem As String) As Boolean
Dim lobjClausula    As clsA46V703A
Dim lobjstcClau     As stcA46V703B
        
    Set lobjClausula = New clsA46V703A
    Set lobjstcClau = New stcA46V703B
    
    With lobjstcClau
           .NOSSO_NUMERO = pobjstcItem.NOSSO_NUMERO
           .Cod_Clau = pClausula
           .Cod_Ramo = pobjstcItem.Cod_Ramo
           .Num_Ped = pobjstcItem.Num_Ped
           .Num_Item = pobjstcItem.Num_Item
           .Descricao = " "
           .Num_PI = pobjstcItem.Num_PI
    End With
    If Not lobjClausula.mfIncluir(lobjstcClau, pPropTabela, pNomeTabImport) Then
        pMensagem = lobjClausula.UltimoErro
        GoTo SAIDA
    End If
    On Error Resume Next
    pobjstcItem.ColecaoPedClau.Remove lobjstcClau.Class_ID
    On Error GoTo 0
    Set lobjstcClau.ObjPai = pobjstcItem
    pobjstcItem.ColecaoPedClau.Add lobjstcClau, lobjstcClau.Class_ID
    
    mfIncluirClausula = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjClausula = Nothing
    Set lobjstcClau = Nothing
    Return
End Function
Public Sub gpVerNormasParametrizada(pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B, pMensagem As String, pTipoConsistencia As Integer)

    'Procedure #: verifica normas de aceitação
Dim lMensagem           As String           'Mensagem.
Dim lNormaRes           As Boolean          'Resposta da norma.
Dim lPrefNorma          As String           'Prefixo na mensagem.
Dim lResultado          As String           'Resultado.
Dim lRstNorma           As ADODB.Recordset  'Registros (normas).
Dim lSelect             As String           'Select.
Dim lNivel()            As String           'Array de nível.
Dim lCodSequencia       As Long             'Código da sequencia.
Dim lCodProtocolo       As Long             'Código da sequencia.
Dim BDP00Tab            As ADODB.Connection 'Conexão com a tabela.
Dim lListaRegra         As String
Dim lBDP00Multi         As ADODB.Connection
Dim lrst                As ADODB.Recordset
Dim lInsert             As String

    If Not UCase(gSerDIVERSOS) Like "*MIX*" Or gTipDIVERSOS = "2" Then GoTo SAIDA
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, BDP00Tab, pMensagem) = False Then
        GoTo SAIDA
    End If
    
    lCodSequencia = lCodSequencia
    lCodProtocolo = lCodProtocolo
    Do
        lSelect = "Select distinct Tbm_Parm_Normas.cod_grupo_Parm,  "
        lSelect = lSelect & " Tbm_Parm_Normas.cod_clasf_Parm, "
        lSelect = lSelect & " Tbm_Parm_Normas.cod_Parm , "
        lSelect = lSelect & " Tbm_Parm_Normas.Cod_Parm_Seq , "
        lSelect = lSelect & " Tbm_Parm_Normas.Resultado ,"
        lSelect = lSelect & " Tbm_Parm_Normas.Desc_Parm "
        lSelect = lSelect & " as Descricao from"
        lSelect = lSelect & " Tbm_Parm_Normas "
        lSelect = lSelect & " where "
        lSelect = lSelect & " Tbm_Parm_Normas.cod_ramo = " & lCod_Ramo
        lSelect = lSelect & " and tip_parm = " & pTipoConsistencia
        lSelect = lSelect & " and Tbm_Parm_Normas.cod_produto = " & lCod_Tarifa
        lSelect = lSelect & " and Tbm_Parm_Normas.cod_protocolo in (0, " & lCodProtocolo & ")"
        lSelect = lSelect & " and Tbm_Parm_Normas.cod_sequencia = " & lCodSequencia
                
        If gfObtRegistro(BDP00Tab, lSelect, lRstNorma, lMensagem) = False Then
            Call gpFecha2(BDP00Tab)
            Call gpMenErro(lMensagem)
            GoTo SAIDA
        End If
        
        If lRstNorma.EOF = True Then
            If lCodSequencia > 0 Then
                lCodSequencia = lCodSequencia - 1
            Else
                If lCodProtocolo > 0 Then
                    If lCodProtocolo > 10 Then
                        lCodProtocolo = 0
                    Else
                        lCodProtocolo = lCodProtocolo - 1
                    End If
                    lCodSequencia = lCodSequencia
                Else
                    Call gpFecha3(lRstNorma, BDP00Tab)
                    Call gpMenErro("Nenhuma norma foi localizada!")
                    GoTo SAIDA
                End If
            End If
        Else
            Exit Do
        End If
    Loop
    'Fim A
    
    Do While Not lRstNorma.EOF
        If Not mfParmCalculo(BDP00Tab, lRstNorma, lResultado, lCodProtocolo, _
            lCodSequencia, 1, pobjstcPed, pobjstcPedLoc, lMensagem) Then
            Call gpMenErro(lMensagem)
            GoTo SAIDA
        End If
        lNormaRes = (lResultado = "1")
        If lNormaRes = True Then
            lPrefNorma = "Norma " & lRstNorma("Cod_Grupo_Parm") & "." & lRstNorma("Cod_Clasf_Parm") & "." & lRstNorma("Cod_Parm")
            If lRstNorma("Cod_Parm_Seq") <> 0 Then
                lPrefNorma = lPrefNorma & "." & lRstNorma("Cod_Parm_Seq")
            End If
            lPrefNorma = lPrefNorma & " - "
            If IsNull(lRstNorma("Descricao")) Then
                lMensagem = ""
            Else
                lMensagem = lRstNorma("Descricao")
            End If
            If VerificarNormas = True Then
                Call gpMenErro(lPrefNorma & lMensagem)
            End If
            lListaRegra = lListaRegra & lRstNorma("Cod_Grupo_Parm") & "." & lRstNorma("Cod_Clasf_Parm") & "." & lRstNorma("Cod_Parm") & "/"
            
            If gTipUsuario = e_Tipo_Usu_Yasuda And gTipoAcesso = CALCULO Then
                If pobjstcPed.objstcCtrlEmis.Tip_Cotac = "2" And pobjstcPed.objstcCtrlEmis.Num_Cotac <> "" Then
                    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
                        Call gpGraLog(2, lMensagem)
                        GoTo SAIDA
                    End If
                    lInsert = " Insert into siscota.dbo.Tab_Cotacao_CheckList_Resseguro (CCotacao, Cod_ItemCheckList, Cod_Ramo, Num_Item ) values " & _
                            "('" & pobjstcPed.objstcCtrlEmis.Num_Cotac & "'," & lRstNorma("Cod_Grupo_Parm") & lRstNorma("Cod_Clasf_Parm") & Format(lRstNorma("Cod_Parm"), "00") & "," & _
                            pobjstcPed.Cod_Ramo & "," & pobjstcPedLoc.Num_Item & ")"
                    If gfExeSQL(lBDP00Multi, lInsert, lMensagem) <> 0 Then
                        Call gpGraLog(2, lMensagem)
                        GoTo SAIDA
                    End If
                End If
            End If
        End If
        lRstNorma.MoveNext
    Loop
    
    'Thais - resseguro
    Call gpFecha3(lRstNorma, BDP00Tab)
    If gfPreenchido(lListaRegra) Then
        If gTipUsuario = e_Tipo_Usu_Yasuda And gTipoAcesso = DIGITACAO Then
            If gfAbrBasDados("GED", gTipGed, gEndGed, lBDP00Multi, lMensagem) = False Then
                Call gpGraLog(2, lMensagem)
                GoTo SAIDA
            End If
            lSelect = " Select * from ged.dbo.tab_prop_mensagem where num_pi = " & pobjstcPed.Num_PI & " and cod_mensagem = 8000 and item_pi = " & pobjstcPedLoc.Num_Item
            If Not gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) Then
                Call gpGraLog(2, lMensagem)
                GoTo SAIDA
            End If
            If lrst.EOF Then
                lInsert = " INSERT INTO [GED].[dbo].[tab_Prop_Mensagem]([num_pi], [item_pi], [cod_mensagem], [cod_tip_mensagem], [motivo], [dat_digit], [dat_liber], [hor_liber]) "
                lInsert = lInsert & " VALUES(" & pobjstcPed.Num_PI & "," & pobjstcPedLoc.Num_Item & ",8000,99,'" & lListaRegra & "'," & Format(Date, "yyyymmdd") & ",0,0)"
                If gfExeSQL(lBDP00Multi, lInsert, lMensagem) <> 0 Then
                    Call gpGraLog(2, lMensagem)
                    GoTo SAIDA
                End If
            End If
        End If
    Else
        If gTipUsuario = e_Tipo_Usu_Yasuda And gTipoAcesso = DIGITACAO Then
            If gfAbrBasDados("GED", gTipGed, gEndGed, lBDP00Multi, lMensagem) = False Then
                Call gpGraLog(2, lMensagem)
                GoTo SAIDA
            End If
            lSelect = " Delete from ged.dbo.tab_prop_mensagem where num_pi = " & pobjstcPed.Num_PI & " and cod_mensagem = 8000 and dat_liber = 0  and item_PI = " & pobjstcPedLoc.Num_Item
            If Not gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) Then
                Call gpGraLog(2, lMensagem)
                GoTo SAIDA
            End If
        End If
    End If
    
SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lRstNorma = Nothing
    Set BDP00Tab = Nothing
    Set lBDP00Multi = Nothing
    Set lrst = Nothing
    Return
End Sub
Public Function gpInicializarIdentificacaoResseguro(pobjstcPed As stcA46V702B)
Dim lBDP00Multi     As ADODB.Connection
Dim lMensagem       As String

    If gTipUsuario = e_Tipo_Usu_Yasuda And gTipoAcesso = CALCULO Then
        If pobjstcPed.objstcCtrlEmis.Tip_Cotac = "2" And pobjstcPed.objstcCtrlEmis.Num_Cotac <> "" Then
            If Not UCase(gSerDIVERSOS) Like "*MIX*" Then GoTo SAIDA
            If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
                Call gpGraLog(2, lMensagem)
                GoTo SAIDA
            End If
            gDelete = " Delete from siscota.dbo.Tab_Cotacao_CheckList_Resseguro where ccotacao = '" & pobjstcPed.objstcCtrlEmis.Num_Cotac & "'"
            If gfExeSQL(lBDP00Multi, gDelete, lMensagem) <> 0 Then
                Call gpGraLog(2, lMensagem)
                GoTo SAIDA
            End If
        End If
    End If
    
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lBDP00Multi = Nothing
    Return
End Function
Private Function mfParmCalculo(pBDConexao As ADODB.Connection, pRstParm As ADODB.Recordset, _
                            pResultado As String, pCodProtocolo As Long, _
                            pCodSequencia As Long, pTip_Parm As Integer, _
                            pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B, pMensagem As String) As Boolean
                            
                            
Dim lAux1Num            As Double           'Auxiliar 1 numérico.
Dim lAux2Num            As Double           'Auxiliar 2 numérico.
Dim lCondicao           As Byte             '1 = and; 2= or.
Dim lCondNivel          As Byte             'Condição do Nível.
Dim lCount              As Integer          'Utilizado no for...loop.
Dim lNivelRes           As Boolean          'Resposta no nível.
Dim lNivelAnt           As Integer          'Nível anterior.
Dim lNormaRes           As Boolean          'Resposta da norma.
Dim lPriVez             As Boolean          'Indicador de primeira vez.
Dim lRegraRes           As Boolean          'Resposta da regra.
Dim lRstRegra           As ADODB.Recordset  'Registros (regras).
Dim lSaida              As Variant          'Saída da função.
Dim lSelect             As String           'Select.
Dim lValorNum           As Double           'Valor de retorno numérico.
Dim lValorRegraNum      As Double           'Valor da regra (numérico).
Dim lValorRegraStr      As String           'Valor da regra (string).
Dim lValorStr           As String           'Valor de retorno string.
Dim lNivel()            As String           'Array de nível.
Dim lUltCodProtocolo    As Long             'Ultimo codigo do protocolo
Dim lRetorno            As String           'Retorno
Dim lObjstcCorr         As stcA46V705B
Dim lobjstcCobert       As stcA46V704B
Dim lobjstcPedLoc       As stcA46V708B
Dim lValorIS            As Double
Dim lValorSoma          As Double
Dim lISCob1             As Double
Dim lISCob2             As Double
Dim lAuxiliar           As String
Dim lVigIniAno          As Long
Dim lVigIniMes          As Long
Dim lVigIniDia          As Long
Dim lVigIniAnoMes       As Long
Dim lVigTerAno          As Long
Dim lVigTerMes          As Long
Dim lVigTerDia          As Long
Dim lVigTerAnoMes       As Long
                

    pResultado = ""
    mfParmCalculo = False
      
        
    
    'início C
    lNivelAnt = 1
    'Fim C
    'Início D
    lSelect = " Select Tbm_Parm_Normas.cod_protocolo , Tbm_Parm_Normas.cod_regra, Desc_Parm, Nivel, Cond_nivel,"
    lSelect = lSelect & " Resultado, Tip_Regra,Item_Parm, Val_Regra, Operador, "
    lSelect = lSelect & " Condicao , Resultado FROM Tbm_Regra_Norma INNER JOIN Tbm_Parm_Normas "
    lSelect = lSelect & " ON Tbm_Regra_Norma.Cod_Regra = Tbm_Parm_Normas.Cod_Regra"
    lSelect = lSelect & " where Tbm_Parm_Normas.cod_parm = " & pRstParm("Cod_Parm")
    lSelect = lSelect & " and cod_clasf_parm = " & pRstParm("Cod_Clasf_Parm")
    lSelect = lSelect & " and Cod_Grupo_Parm = " & pRstParm("Cod_Grupo_Parm")
    lSelect = lSelect & " and Cod_Parm_Seq = " & pRstParm("Cod_Parm_Seq")
    lSelect = lSelect & " and Tip_Parm = " & pTip_Parm
    lSelect = lSelect & " and Cod_produto = " & lCod_Tarifa
    lSelect = lSelect & " and cod_protocolo in (0, " & pCodProtocolo & ")"
    lSelect = lSelect & " and cod_sequencia = " & pCodSequencia
    lSelect = lSelect & " and Tbm_Parm_Normas.cod_ramo = " & lCod_Ramo
    lSelect = lSelect & " order by cod_protocolo desc , Item_Parm "
    
    If gfObtRegistro(pBDConexao, lSelect, lRstRegra, pMensagem) = False Then
        Call gpFecha2(pBDConexao)
        Call gpMenErro(pMensagem)
        GoTo SAIDA
    End If
    
    If lRstRegra.EOF = True Then
        Call gpFecha3(lRstRegra, pBDConexao)
        Call gpMenErro("Nenhuma regra foi localizada!")
        GoTo SAIDA
    End If
    lPriVez = True
    lUltCodProtocolo = -1
    Do While Not lRstRegra.EOF
        If lUltCodProtocolo = -1 Then
            lUltCodProtocolo = lRstRegra("cod_protocolo")
        End If
        If lUltCodProtocolo <> lRstRegra("cod_protocolo") Then
            Exit Do
        End If
        'início E
        If lRstRegra("Nivel") <> lNivelAnt And lRstRegra("Nivel") > 1 Then
            ReDim Preserve lNivel(3, lNivelAnt)
            lNivel(1, lNivelAnt) = lNivelAnt
            lNivel(2, lNivelAnt) = lNivelRes
            lNivel(3, lNivelAnt) = lCondNivel
            lNivelAnt = lRstRegra("Nivel")
            lPriVez = True
        End If
        'Fim E
        
        'Início G
        lCondNivel = lRstRegra("Cond_Nivel")
        '1- Verifica se é a primeira regra do parametro que esta sendo avaliada.
        If lPriVez = True Then
           lPriVez = False
           lNivelRes = True
           lCondicao = 1      'AND
        End If
    
        '2- Atribui o valor da regra para uma variavel de acordo com o tipo de valor.
        Select Case lRstRegra("Tip_Regra")
            Case 1    '' texto
                lValorRegraStr = lRstRegra("Val_Regra")
            Case Else
                lValorRegraNum = lRstRegra("Val_Regra")
        End Select
        lValorNum = 0
        lValorStr = ""
        '3- Atribui o valor informado para validação para uma variavel de acordo com o tipo de valor.**
        Select Case lRstRegra("Cod_Regra")
            Case 1   'Comissão de corretagem
                lValorNum = 0
                For Each lObjstcCorr In pobjstcPed.ColecaoPedCorr
                    lValorNum = lValorNum + lObjstcCorr.Per_Corr
                Next lObjstcCorr
            Case 2   'Código da classificação
                lValorNum = Val(pobjstcPedLoc.Cod_Clasf)
            Case 3   'Código do complemento da classificação
                lValorNum = Val(pobjstcPedLoc.Cod_Compl_Clasf)
            Case 4   'Tipo de emissão
                lValorNum = pobjstcPedLoc.Tip_Emissao
            Case 5   'Tem Habitual?
                lValorNum = 2
                For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
                    If lobjstcPedLoc.Cod_Compl_Clasf = "011" Or lobjstcPedLoc.Cod_Compl_Clasf = "021" Then
                        lValorNum = 1
                    End If
                Next lobjstcPedLoc
            Case 6   'Código da cobertura
                'Buscar a cobertura
                lValorNum = 0
                For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
                    If lobjstcCobert.Cod_Cobert = lValorRegraNum Then
                        lValorNum = lobjstcCobert.Cod_Cobert
                        lValorIS = lobjstcCobert.Val_IS
                        Exit For
                    End If
                Next lobjstcCobert
                 
            Case 7   'IS da cobertura
                lValorNum = lValorIS
            Case 8   'Classe de construção
                lValorNum = pobjstcPedLoc.Cod_Constr
            Case 9   'CEP de risco
                lValorNum = pobjstcPedLoc.Num_Cep_Risco
            Case 10  'Código do corretor
                lValorNum = pobjstcPed.objstcCtrlEmis.Cod_Corr
            Case 11 '                11  (Somatoria) IS da cobertura
                lValorNum = 0
                For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
                    If lobjstcCobert.Cod_Cobert = lValorRegraNum Then
                        lValorSoma = lValorSoma + lobjstcCobert.Val_IS
                        lValorNum = lRstRegra("Val_Regra")
                        Exit For
                    End If
                Next lobjstcCobert
                If lRstRegra("Val_Regra") = 1 And lValorNum = 0 Then
                    lValorSoma = pobjstcPedLoc.Val_Vr
                    lValorNum = 1
                End If
            Case 12
                lValorNum = lValorSoma
            Case 13 'IS Cob 1 - comparação %
                 lValorNum = 0
                 lISCob1 = 0
                For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
                    If lobjstcCobert.Cod_Cobert = lValorRegraNum Then
                        lValorNum = lRstRegra("Val_Regra")
                        lISCob1 = lobjstcCobert.Val_IS
                        Exit For
                    End If
                Next lobjstcCobert
            Case 14 'IS Cob 2 - comparação %
                lValorNum = 0
                lISCob2 = 0
                For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
                    If lobjstcCobert.Cod_Cobert = lValorRegraNum Then
                        lValorNum = lRstRegra("Val_Regra")
                        lISCob2 = lobjstcCobert.Val_IS
                        Exit For
                    End If
                Next lobjstcCobert
            Case 15 'Is cob 1 / Is cob 2 %
                lValorNum = 0
                If lISCob2 <> 0 Then
                    lValorNum = (lISCob1 / lISCob2) * 100
                End If
            Case 16 'Meses de vigência
                'a) Calcula quantidade de meses de vigência.
                lAuxiliar = Format$(pobjstcPed.Dat_Ini_Vig, "00000000")    'Formato: aaaammdd.
                lVigIniAno = CInt(Mid(lAuxiliar, 1, 4))
                lVigIniMes = CInt(Mid(lAuxiliar, 5, 2))
                lVigIniDia = CInt(Mid(lAuxiliar, 7, 2))
                lVigIniAnoMes = CLng(Mid(lAuxiliar, 1, 6))
        
                lAuxiliar = Format$(pobjstcPed.Dat_Ter_Vig, "00000000")    'Formato: aaaammdd.
                lVigTerAno = CInt(Mid(lAuxiliar, 1, 4))
                lVigTerMes = CInt(Mid(lAuxiliar, 5, 2))
                lVigTerDia = CInt(Mid(lAuxiliar, 7, 2))
                lVigTerAnoMes = CLng(Mid(lAuxiliar, 1, 6))
                
                Do Until Val(Mid(lVigIniAnoMes, 1, 4)) = Val(Mid(lVigTerAnoMes, 1, 4))
                    lVigTerAnoMes = lVigTerAnoMes - 100
                    lVigTerAnoMes = lVigTerAnoMes + 12
                Loop
                lValorNum = lVigTerAnoMes - lVigIniAnoMes
            
            Case Else
                gpMenErro ("FDE0023 - Tipo de regra desconhecido - Verificação das normas.")
        End Select

        '4- Faz a validação o item da regra.
        Select Case lRstRegra("Tip_Regra")
            Case 1      'texto
                Select Case lRstRegra("Operador")
                    Case 1     '=
                        If UCase(lValorStr) = UCase(lValorRegraStr) Then
                            lRegraRes = True
                        Else
                            lRegraRes = False
                        End If
                    Case 6      '<>
                        If UCase(lValorStr) <> UCase(lValorRegraStr) Then
                            lRegraRes = True
                        Else
                            lRegraRes = False
                        End If
                    Case Else
                        pMensagem = "erro"
                End Select
            Case Else     'Numerico
                Select Case lRstRegra("Operador")
                    Case 1      '=
                        If lValorNum = lValorRegraNum Then
                            lRegraRes = True
                        Else
                            lRegraRes = False
                        End If
                    Case 2      '>
                        If lValorNum > lValorRegraNum Then
                            lRegraRes = True
                        Else
                            lRegraRes = False
                        End If
                    Case 3     ' <
                        If lValorNum < lValorRegraNum Then
                            lRegraRes = True
                        Else
                            lRegraRes = False
                        End If
                    Case 4     ' >=
                        If lValorNum >= lValorRegraNum Then
                            lRegraRes = True
                        Else
                            lRegraRes = False
                        End If
                    Case 5     ' <=
                        If lValorNum <= lValorRegraNum Then
                            lRegraRes = True
                        Else
                            lRegraRes = False
                        End If
                    Case 6      ' <>
                        If lValorNum <> lValorRegraNum Then
                            lRegraRes = True
                        Else
                            lRegraRes = False
                        End If
                End Select
        End Select
    
        Select Case lCondicao
            Case 1    'AND"
                If lNivelRes = True And lRegraRes = True Then
                    lNivelRes = True
                Else
                    lNivelRes = False
                End If
            Case 2   '"OR"
                If lNivelRes = True Or lRegraRes = True Then
                    lNivelRes = True
                Else
                    lNivelRes = False
                End If
        End Select
        lCondicao = lRstRegra("Condicao")
        lRetorno = lRstRegra("Resultado")
        'Fim G
        lRstRegra.MoveNext
    Loop
    'Adiciono na coleção o último nível
    ReDim Preserve lNivel(3, lNivelAnt)
    lNivel(1, lNivelAnt) = lNivelAnt
    lNivel(2, lNivelAnt) = lNivelRes
    lNivel(3, lNivelAnt) = lCondNivel
            
    'início H
    'Este procedimento faz a validação final da regra, verificando o retorno dos níveis do parâmetro atribuindo seu resultado final.
    pResultado = ""
    For lCount = 1 To lNivelAnt
        If lCount = 1 Then
            lNormaRes = True
            If lNormaRes = True And lNivel(2, lCount) = True Then
                pResultado = lRetorno
                lNormaRes = True
            Else
                pResultado = ""
                lNormaRes = False
            End If
            lCondNivel = lNivel(3, lCount)
        Else
            Select Case lCondNivel
                Case 1      'AND
                    If lNormaRes = True And lNivel(2, lCount) = True Then
                        pResultado = lRetorno
                        lNormaRes = True
                    Else
                        pResultado = ""
                        lNormaRes = False
                    End If
                Case 2      'OR
                    If lNormaRes = True Or lNivel(2, lCount) = True Then
                        pResultado = lRetorno
                        lNormaRes = True
                    Else
                        pResultado = ""
                        lNormaRes = False
                    End If
            End Select
            lCondNivel = lNivel(3, lCount)
        End If
    Next lCount
    
    Call gpFecha1(lRstRegra)
    
    mfParmCalculo = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lRstRegra = Nothing
    Set lObjstcCorr = Nothing
    Set lobjstcCobert = Nothing
    Set lobjstcPedLoc = Nothing
    Return
End Function


'1200429 - Tarifa 1036 e 1037 Maio 2012 - Item - 9
Public Function mfAdicionaTextoEspecificacao(ByVal pCodEspecificacao As Long, ByRef pObjTextoEspec As clsA46V726A, ByRef pStcPedido As stcA46V702B, ByVal pPropTabela As String, ByRef pMensagem As String) As Boolean
Dim lobjstcTextoEspec     As stcA46V726B
Dim lbdTab_Ped_Dizer_Pad  As ADODB.Connection
Dim lSelect               As String
Dim lrsTab_Ped_Dizer_Pad  As ADODB.Recordset

    mfAdicionaTextoEspecificacao = False

    Set lobjstcTextoEspec = New stcA46V726B
    
    With lobjstcTextoEspec
        .NOSSO_NUMERO = pStcPedido.NOSSO_NUMERO
        .Cod_Org = pStcPedido.Cod_Org
        .Cod_Ramo = pStcPedido.Cod_Ramo
        .Num_Ped = pStcPedido.Num_Ped
        .Num_Apol = pStcPedido.Num_Apol
        .Num_Endo = pStcPedido.Num_Endo
        .Cod_Espec = pCodEspecificacao   'cadastrado no main-frame
        .Num_pag = pStcPedido.ColecaoTextoEspec.Count + 1
        .Qtd_vias = 3
'        'abre base de dados
        If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTab_Ped_Dizer_Pad, pMensagem) = False Then
            GoTo SAIDA
        End If
        'consulta mensagem
        lSelect = "SELECT * FROM TAB_PED_DIZER_PAD WHERE COD_DIZER_PAD = " & .Cod_Espec
        If gfObtRegistro(lbdTab_Ped_Dizer_Pad, lSelect, lrsTab_Ped_Dizer_Pad, pMensagem) = False Then
            Call gpFecha2(lbdTab_Ped_Dizer_Pad)
            GoTo SAIDA
        End If
        If lrsTab_Ped_Dizer_Pad.RecordCount = 0 Then pMensagem = "Mensagem Padrão " & .Cod_Espec & " não encontrada!": GoTo SAIDA
        .Espec_lin1 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString)
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin2 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin3 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin4 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin5 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin6 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin7 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin8 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin9 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin10 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin11 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin12 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin13 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin14 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin15 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin16 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin17 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin18 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin19 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin20 = gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
FimDizer:
        Call gpFecha3(lrsTab_Ped_Dizer_Pad, lbdTab_Ped_Dizer_Pad)
        .Num_PI = pStcPedido.Num_PI '  gNum_Pi
    End With
    If Not pStcPedido.FindTextoEspec(lobjstcTextoEspec) Then
        If Not pObjTextoEspec.mfIncluir(lobjstcTextoEspec, pPropTabela, "") Then
            pMensagem = pObjTextoEspec.UltimoErro
            GoTo SAIDA
        End If
        pStcPedido.ColecaoTextoEspec.Add lobjstcTextoEspec, lobjstcTextoEspec.Class_ID
    End If
    
    mfAdicionaTextoEspecificacao = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcTextoEspec = Nothing
    Set lbdTab_Ped_Dizer_Pad = Nothing
    Set lrsTab_Ped_Dizer_Pad = Nothing
    Return
End Function


