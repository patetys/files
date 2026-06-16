VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsM46V998"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================================================================================
' Class Name: clsM46V998
' Sistema: SyasMult
' Analista: Marcio Oliveira
' Objetivo: Substitui o antigo módulo bas Adaptado para ser utilizado como classe gerenciavel.
' ======================================================================================================================================================
'Ficha 1000495 - Plano Belacity
Public gPlanoImobiliaria As Boolean
'Ficha 1000601

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "M46V998"
End Property

Private Function mfCoeJuros(ByVal pTipPagamento As Byte, ByVal pQuaParcelas As Single, _
                            ByVal pJuros As Single) As Single
'Função #007: calcula coeficiente de juros.

'Parâmetros de entrada...pTipPagamento...0 = pagamento postecipado e 1 = pagamento antecipado.
'                        pQuaParcelas....quantidade de parcelas.
'                        pJuros..........juros.

    Dim lDesPagamento As String    'Descrição do pagamento.

    If pJuros = 0 Then
        If pQuaParcelas = 0 Then
            mfCoeJuros = 1
        Else
            mfCoeJuros = 1 / pQuaParcelas
        End If
    Else
        Select Case pTipPagamento
        Case 0  'Pagamento postecipado.
            mfCoeJuros = M46V999.gfTrunca(((pJuros / 100) / (1 - (1 / (1 + (pJuros / 100)) ^ _
                                                                  pQuaParcelas))), 4)
        Case 1  'Pagamento antecipado.
            mfCoeJuros = M46V999.gfTrunca(((pJuros / 100) / (1 - (1 / (1 + (pJuros / 100)) ^ _
                                                                  pQuaParcelas))) * (1 / (1 + (pJuros / 100))), 4)
        End Select
    End If

    If pTipPagamento = 0 Then
        lDesPagamento = "0 (pagamento postecipado)"
    Else
        lDesPagamento = "1 (pagamento antecipado)"
    End If

End Function

Public Function gfCalFracionamentoOld2(ByRef pStcPed As stcA46V702B, _
                                       ByVal pTip_Emissao As Long, _
                                       ByRef pParcela As Variant, ByRef pMensagem As String, _
                                       ByVal pProJurIOF As Single, Optional ByVal pCodReg As Integer, _
                                       Optional ByRef pIOF As Single, Optional ByVal pParcMin As Single)

'1100303 - Arredondamento no Parcelamento - 'Foi realizado troca de algumas variáveis de Double para Single e tambem substituido a função M46V999.gfTrunca() pela Round() - caso queira comparar historico de alteração no SouceSafe

'Função: calcula fracionamento.

'Parâmetros de entrada...pStcPed ..........Stc do pedido
'                        pTip_Emissao......Tipo de Emissão
'                        pCod_Registro.....Código do Registro
'                        pProJurIOF........ Juros IOF
'                        pCodReg........... Codigo de registro (Opcional)
'                        pParcMin.......... Valor minimo de parcela
'Parâmetro de saída......pParcela.........matriz com resultados das parcelas.
'                        pMensagem........mensagem.
'                        pErro............lista de erros.
'                        pIOF.............Valor do IOF para valor determinado (opcional).
    Dim lI1 As Byte    'Utilizado com For...Next.
    Dim lI2 As Byte    'Utilizado com For...Next.
    Dim lValorMin As Single    'Valor mínimo da parcela ou endosso.
    Dim lPremioMin As Single  'Valor mínimo do prêmio de parcela ou endosso.
    Dim lParcVista As Integer    'Quantidade de parcelas máxima possível em caso de pagamento antecipado.
    Dim lParcs30dia As Integer  'Quandidade de parcelas máxima possível em caso de paramento não antecipado.
    Dim lPreTotal As Single    'Prêmio total.
    Dim lValor As Single    'Prêmio Liquido.
    Dim lIOF As Single    'Valor IOF.
    Dim lPer_Juros_1 As Double    'Juros da 1ª parcela.
    Dim lPer_Juros_2 As Double    'Juros da 2ª parcela.
    Dim lPer_Juros_3 As Double    'Juros da 3ª parcela.
    Dim lPer_Juros_4 As Double    'Juros da 4ª parcela.
    Dim lPer_Juros_5 As Double    'Juros da 5ª parcela.
    Dim lPer_Juros_6 As Double    'Juros da 6ª parcela.
    Dim lPer_Juros_7 As Double    'Juros da 7ª parcela.
    Dim lPer_Juros_8 As Double    'Juros da 8ª parcela.
    Dim lPer_Juros_9 As Double    'Juros da 9ª parcela.
    Dim lPer_Juros_10 As Double     'Juros da 10ª parcela.
    Dim lPer_Juros_11 As Double     'Juros da 11ª parcela.
    Dim lPer_Juros_12 As Double     'Juros da 12ª parcela.
    Dim lPer_Juros_0_1 As Double      'Juros da 1ª parcela antecipada.
    Dim lPer_Juros_0_2 As Double      'Juros da 2ª parcela antecipada.
    Dim lPer_Juros_0_3 As Double      'Juros da 3ª parcela antecipada.
    Dim lPer_Juros_0_4 As Double      'Juros da 4ª parcela antecipada.
    Dim lPer_Juros_0_5 As Double      'Juros da 5ª parcela antecipada.
    Dim lPer_Juros_0_6 As Double      'Juros da 6ª parcela antecipada.
    Dim lPer_Juros_0_7 As Double      'Juros da 7ª parcela antecipada.
    Dim lPer_Juros_0_8 As Double      'Juros da 8ª parcela antecipada.
    Dim lPer_Juros_0_9 As Double      'Juros da 9ª parcela antecipada.
    Dim lPer_Juros_0_10 As Double       'Juros da 10ª parcela antecipada.
    Dim lPer_Juros_0_11 As Double       'Juros da 11ª parcela antecipada.
    Dim lPer_Juros_0_12 As Double       'Juros da 12ª parcela antecipada.
    Dim lobjstcJuros As stcA46V711B
    Dim lTemAjusteVenc As Boolean
    Dim lSelect As String

    Dim lIniVigencia As String
    Dim lTerVigencia As String
    Dim lCusEmissao As Single
    Dim lCusParcelado As String
    Dim lPreLiquido As Single
    Dim lCod_Iden_Tarf As Long

    Dim objStcParcela As stcParcelamento
    '001 - Apólice com cobrança de emolumentos.
    '002 - Apólice sem cobrança de emolumentos.
    '003 - Apólice sem cobrança de IOF.
    '004 - Apólice sem cobrança de custo de apólice.
    '005 - Apólice com moeda estrangeira (ramo transporte/crédito de exportação).
    '006 - Apólice coletiva de acidentes pessoais - pagamento mensal.
    '007 - Apólice coletiva - pagamento trimestral.
    '008 - Apólice coletiva - pagamento semestral.
    '020 - Apólice a averbar (sem cobrança de prêmio depósito).
    '101 - Endosso com cobrança de emolumentos.
    '102 - Endosso sem cobrança de emolumentos.
    '103 - Endosso sem cobrança de IOF.
    '104 - Endosso sem cobrança de custo de apólice.
    '105 - Endosso com moeda estrangeira (ramo transporte/crédito de exportação).
    '106 - Endosso de restituição de prêmio sem cancelamento da apólice.
    '107 - Endosso de restituição de prêmio com cancelamento da apólice.
    '120 - Endosso sem movimento de prêmio.
    '121 - Endosso sem movimento com prêmio negativo.
    '122 - Endosso sem movimento com prêmio positivo.
    '129 - Cancelamento da apólice sem estorno de prêmio.


    On Error GoTo TrataErro

    gfCalFracionamentoOld2 = False



    'Inicio os valores

    lIniVigencia = M46V999.gfForData(2, pStcPed.Dat_Ini_Vig)
    lTerVigencia = M46V999.gfForData(2, pStcPed.Dat_Ter_Vig)

    'Simony
    M46V999.gProQuaMaxParcelas = pStcPed.getProduto.ProQuaMaxParcelas


    Call mpObtemMaxParcelas(M46V999.gProQuaMaxParcelas, lParcVista, lParcs30dia, pStcPed.Dat_Ini_Vig, pStcPed.Dat_Ter_Vig)


    lCusEmissao = pStcPed.Val_Cust_Emis
    lCusParcelado = IIf(pStcPed.gCustoParcelado = True, "S", "N")
    lPreLiquido = pStcPed.VAL_PR
    lCod_Iden_Tarf = pStcPed.Cod_Produto

    If lPreLiquido = 0 Then
        If pStcPed.Cod_Dizer <> 0 Then
            pStcPed.Cod_Reg = 120
        Else
            If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
                pStcPed.Cod_Reg = 4
            Else
                pStcPed.Cod_Reg = 20
            End If
        End If
        pStcPed.Val_Cust_Emis = 0
    ElseIf lPreLiquido > 0 Then
        If pStcPed.Cod_Dizer = 0 Then
            If pStcPed.Cod_Reg > 100 Or pStcPed.Cod_Reg = 0 Then
                pStcPed.Cod_Reg = 1
            End If
        Else
            If pStcPed.Cod_Reg < 100 Or pStcPed.Cod_Reg > 105 Then
                pStcPed.Cod_Reg = 101
            End If
        End If
    Else
        'Valor negativo
        If pStcPed.Cod_Dizer = 10 Then
            pStcPed.Cod_Reg = 107
        Else
            pStcPed.Cod_Reg = 106
        End If
        pStcPed.Val_Cust_Emis = 0
    End If

    If pStcPed.Dat_Ini_Vig >= 20130101 Then
        pStcPed.Val_Cust_Emis = 0
        If pStcPed.Cod_Reg = 1 Then
            pStcPed.Cod_Reg = 4
        ElseIf pStcPed.Cod_Reg = 3 Then
            pStcPed.Cod_Reg = 2
        ElseIf pStcPed.Cod_Reg = 101 Then
            pStcPed.Cod_Reg = 104
        ElseIf pStcPed.Cod_Reg = 103 Then
            pStcPed.Cod_Reg = 102
        End If
    Else
        If pStcPed.Cod_Reg = 4 Then
            pStcPed.Cod_Reg = 1
        ElseIf pStcPed.Cod_Reg = 2 Then
            pStcPed.Cod_Reg = 3
        ElseIf pStcPed.Cod_Reg = 104 Then
            pStcPed.Cod_Reg = 101
        ElseIf pStcPed.Cod_Reg = 102 Then
            pStcPed.Cod_Reg = 103
        End If
    End If
    lCusEmissao = pStcPed.Val_Cust_Emis
    lPer_Juros_1 = 3
    lPer_Juros_2 = 3
    lPer_Juros_3 = 3
    lPer_Juros_4 = 3
    lPer_Juros_5 = 3
    lPer_Juros_6 = 3
    lPer_Juros_7 = 3
    lPer_Juros_8 = 3
    lPer_Juros_9 = 3
    lPer_Juros_10 = 3
    lPer_Juros_11 = 3
    lPer_Juros_12 = 3
    lPer_Juros_0_1 = 3
    lPer_Juros_0_2 = 3
    lPer_Juros_0_3 = 3
    lPer_Juros_0_4 = 3
    lPer_Juros_0_5 = 3
    lPer_Juros_0_6 = 3
    lPer_Juros_0_7 = 3
    lPer_Juros_0_8 = 3
    lPer_Juros_0_9 = 3
    lPer_Juros_0_10 = 3
    lPer_Juros_0_11 = 3
    lPer_Juros_0_12 = 3

    'Incluido Simony
    If pStcPed.Cod_Reg = 2 Or _
       pStcPed.Cod_Reg = 3 Or _
       pStcPed.Cod_Reg = 102 Or _
       pStcPed.Cod_Reg = 103 Or _
       pStcPed.Cod_Reg = 105 Or _
       pStcPed.Cod_Reg = 106 Or _
       pStcPed.Cod_Reg = 107 Or _
       pStcPed.Cod_Reg = 120 Or _
       pStcPed.Cod_Reg = 202 Or _
       pStcPed.Cod_Reg = 203 Then
        pProJurIOF = 0
        pIOF = 0
    End If
    'Fim


    For Each lobjstcJuros In pStcPed.ColecaoPedJuros
        Select Case lobjstcJuros.Num_Parcela
        Case 1
            lPer_Juros_0_1 = lobjstcJuros.Val_Pos
            lPer_Juros_1 = lobjstcJuros.Val_Ant
        Case 2
            lPer_Juros_0_2 = lobjstcJuros.Val_Pos
            lPer_Juros_2 = lobjstcJuros.Val_Ant
        Case 3
            lPer_Juros_0_3 = lobjstcJuros.Val_Pos
            lPer_Juros_3 = lobjstcJuros.Val_Ant
        Case 4
            lPer_Juros_0_4 = lobjstcJuros.Val_Pos
            lPer_Juros_4 = lobjstcJuros.Val_Ant
        Case 5
            lPer_Juros_0_5 = lobjstcJuros.Val_Pos
            lPer_Juros_5 = lobjstcJuros.Val_Ant
        Case 6
            lPer_Juros_0_6 = lobjstcJuros.Val_Pos
            lPer_Juros_6 = lobjstcJuros.Val_Ant
        Case 7
            lPer_Juros_0_7 = lobjstcJuros.Val_Pos
            lPer_Juros_7 = lobjstcJuros.Val_Ant
        Case 8
            lPer_Juros_0_8 = lobjstcJuros.Val_Pos
            lPer_Juros_8 = lobjstcJuros.Val_Ant
        Case 9
            lPer_Juros_0_9 = lobjstcJuros.Val_Pos
            lPer_Juros_9 = lobjstcJuros.Val_Ant
        Case 10
            lPer_Juros_0_10 = lobjstcJuros.Val_Pos
            lPer_Juros_10 = lobjstcJuros.Val_Ant
        Case 11
            lPer_Juros_0_11 = lobjstcJuros.Val_Pos
            lPer_Juros_11 = lobjstcJuros.Val_Ant
        Case 12
            lPer_Juros_0_12 = lobjstcJuros.Val_Pos
            lPer_Juros_12 = lobjstcJuros.Val_Ant
        End Select
    Next lobjstcJuros

    '2. Prepara planilha de resultados.
    ReDim pParcela(12, 4)
    For lI1 = 1 To 12
        For lI2 = 1 To 4
            pParcela(lI1, lI2) = 0
        Next lI2
    Next lI1

    '3. Calcula fracionamento e posiciona resultados na planilha.
    lValor = lPreLiquido
    If lValor <> 0 Then
        '****************************************

        If pStcPed.Cod_Reg < 100 Then
            lPremioMin = M46V999.gProPremioMinimo
        Else
            lPremioMin = M46V999.gProPreMinEndosso
        End If
        If M46V999.gCodProduto >= 976 And lValor < lPremioMin Then
            If pStcPed.Cod_Reg = 106 Or pStcPed.Cod_Reg = 107 Then
                If lValor > lPremioMin Then
                    lValor = 0
                End If
            ElseIf pStcPed.Cod_Reg = 101 Or pStcPed.Cod_Reg = 102 Or pStcPed.Cod_Reg = 103 Or pStcPed.Cod_Reg = 104 Then
                If lValor < lPremioMin Then
                    lValor = 0
                End If
            Else
                lValor = lPremioMin    'prêmio mínimo
            End If
        End If

        lIOF = Round((lValor + lCusEmissao) * (pProJurIOF / 100), 2)
        lPreTotal = lValor + lCusEmissao + lIOF

        lValorMin = pParcMin

        If lValor < lPremioMin And lValor > 0 Then
            lValor = 0
            lPreTotal = 0
            lCusEmissao = 0
            lIOF = 0
        End If
        '2. Calcula valores das parcelas.
        'c) Posiciona resultados.
        '===>>> Parcelas.
        If lCusParcelado = "S" Then
            'a) Considerando custo de emissão parcelado.
            pParcela(1, 1) = Round((lValor + lCusEmissao) * (1 + pProJurIOF / 100), 2)

            pParcela(2, 1) = Round((lValor * mfCoeJuros(1, 2, lPer_Juros_2) + _
                                    (lCusEmissao / 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 2) = pParcela(2, 1)

            pParcela(3, 1) = Round((lValor * mfCoeJuros(1, 3, lPer_Juros_3) + _
                                    (lCusEmissao / 3)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 2) = pParcela(3, 1)

            pParcela(4, 1) = Round((lValor * mfCoeJuros(1, 4, lPer_Juros_4) + _
                                    (lCusEmissao / 4)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 2) = pParcela(4, 1)
            pParcela(5, 1) = Round((lValor * mfCoeJuros(1, 5, lPer_Juros_5) + _
                                    (lCusEmissao / 5)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 2) = pParcela(5, 1)
            pParcela(6, 1) = Round((lValor * mfCoeJuros(1, 6, lPer_Juros_6) + _
                                    (lCusEmissao / 6)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 2) = pParcela(6, 1)
            pParcela(7, 1) = Round((lValor * mfCoeJuros(1, 7, lPer_Juros_7) + _
                                    (lCusEmissao / 7)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 2) = pParcela(7, 1)
            pParcela(8, 1) = Round((lValor * mfCoeJuros(1, 8, lPer_Juros_8) + _
                                    (lCusEmissao / 8)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 2) = pParcela(8, 1)
            pParcela(9, 1) = Round((lValor * mfCoeJuros(1, 9, lPer_Juros_9) + _
                                    (lCusEmissao / 9)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 2) = pParcela(9, 1)
            pParcela(10, 1) = Round((lValor * mfCoeJuros(1, 10, lPer_Juros_10) + _
                                     (lCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 2) = pParcela(10, 1)
            pParcela(11, 1) = Round((lValor * mfCoeJuros(1, 11, lPer_Juros_11) + _
                                     (lCusEmissao / 11)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 2) = pParcela(11, 1)
            pParcela(12, 1) = Round((lValor * mfCoeJuros(1, 12, lPer_Juros_12) + _
                                     (lCusEmissao / 12)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 2) = pParcela(12, 1)


            pParcela(1, 3) = Round((lValor * mfCoeJuros(0, 1, lPer_Juros_0_1) + _
                                    (lCusEmissao / 1)) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 3) = Round((lValor * mfCoeJuros(0, 2, lPer_Juros_0_2) + _
                                    (lCusEmissao / 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 4) = pParcela(2, 3)
            pParcela(3, 3) = Round((lValor * mfCoeJuros(0, 3, lPer_Juros_0_3) + _
                                    (lCusEmissao / 3)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 4) = pParcela(3, 3)
            pParcela(4, 3) = Round((lValor * mfCoeJuros(0, 4, lPer_Juros_0_4) + _
                                    (lCusEmissao / 4)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 4) = pParcela(4, 3)
            pParcela(5, 3) = Round((lValor * mfCoeJuros(0, 5, lPer_Juros_0_5) + _
                                    (lCusEmissao / 5)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 4) = pParcela(5, 3)
            pParcela(6, 3) = Round((lValor * mfCoeJuros(0, 6, lPer_Juros_0_6) + _
                                    (lCusEmissao / 6)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 4) = pParcela(6, 3)
            pParcela(7, 3) = Round((lValor * mfCoeJuros(0, 7, lPer_Juros_0_7) + _
                                    (lCusEmissao / 7)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 4) = pParcela(7, 3)
            pParcela(8, 3) = Round((lValor * mfCoeJuros(0, 8, lPer_Juros_0_8) + _
                                    (lCusEmissao / 8)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 4) = pParcela(8, 3)
            pParcela(9, 3) = Round((lValor * mfCoeJuros(0, 9, lPer_Juros_0_9) + _
                                    (lCusEmissao / 9)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 4) = pParcela(9, 3)
            pParcela(10, 3) = Round((lValor * mfCoeJuros(0, 10, lPer_Juros_0_10) + _
                                     (lCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 4) = pParcela(10, 3)

            pParcela(11, 3) = Round((lValor * mfCoeJuros(0, 11, lPer_Juros_0_11) + _
                                     (lCusEmissao / 11)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 4) = pParcela(11, 3)
            pParcela(12, 3) = Round((lValor * mfCoeJuros(0, 12, lPer_Juros_0_12) + _
                                     (lCusEmissao / 12)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 4) = pParcela(12, 3)
        Else
            'b) Considerando custo de emissão NÃO parcelado.
            pParcela(1, 1) = Round((lValor + lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 1) = Round((lValor * mfCoeJuros(1, 2, lPer_Juros_2) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 2) = Round((lValor * _
                                    mfCoeJuros(1, 2, lPer_Juros_2)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 1) = Round((lValor * mfCoeJuros(1, 3, lPer_Juros_3) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 2) = Round((lValor * _
                                    mfCoeJuros(1, 3, lPer_Juros_3)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 1) = Round((lValor * mfCoeJuros(1, 4, lPer_Juros_4) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 2) = Round((lValor * _
                                    mfCoeJuros(1, 4, lPer_Juros_4)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 1) = Round((lValor * mfCoeJuros(1, 5, lPer_Juros_5) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 2) = Round((lValor * _
                                    mfCoeJuros(1, 5, lPer_Juros_5)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 1) = Round((lValor * mfCoeJuros(1, 6, lPer_Juros_6) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 2) = Round((lValor * _
                                    mfCoeJuros(1, 6, lPer_Juros_6)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 1) = Round((lValor * mfCoeJuros(1, 7, lPer_Juros_7) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 2) = Round((lValor * _
                                    mfCoeJuros(1, 7, lPer_Juros_7)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 1) = Round((lValor * mfCoeJuros(1, 8, lPer_Juros_8) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 2) = Round((lValor * _
                                    mfCoeJuros(1, 8, lPer_Juros_8)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 1) = Round((lValor * mfCoeJuros(1, 9, lPer_Juros_9) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 2) = Round((lValor * _
                                    mfCoeJuros(1, 9, lPer_Juros_9)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 1) = Round((lValor * mfCoeJuros(1, 10, lPer_Juros_10) + _
                                     lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 2) = Round((lValor * _
                                     mfCoeJuros(1, 10, lPer_Juros_10)) * (1 + pProJurIOF / 100), 2)

            pParcela(11, 1) = Round((lValor * mfCoeJuros(1, 11, lPer_Juros_11) + _
                                     lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 2) = Round((lValor * _
                                     mfCoeJuros(1, 11, lPer_Juros_11)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 1) = Round((lValor * mfCoeJuros(1, 12, lPer_Juros_12) + _
                                     lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 2) = Round((lValor * _
                                     mfCoeJuros(1, 12, lPer_Juros_12)) * (1 + pProJurIOF / 100), 2)

            pParcela(1, 3) = Round((lValor * mfCoeJuros(0, 1, lPer_Juros_0_1) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 3) = Round((lValor * mfCoeJuros(0, 2, lPer_Juros_0_2) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 4) = Round((lValor * _
                                    mfCoeJuros(0, 2, lPer_Juros_0_2)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 3) = Round((lValor * mfCoeJuros(0, 3, lPer_Juros_0_3) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 4) = Round((lValor * _
                                    mfCoeJuros(0, 3, lPer_Juros_0_3)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 3) = Round((lValor * mfCoeJuros(0, 4, lPer_Juros_0_4) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 4) = Round((lValor * _
                                    mfCoeJuros(0, 4, lPer_Juros_0_4)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 3) = Round((lValor * mfCoeJuros(0, 5, lPer_Juros_0_5) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 4) = Round((lValor * _
                                    mfCoeJuros(0, 5, lPer_Juros_0_5)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 3) = Round((lValor * mfCoeJuros(0, 6, lPer_Juros_0_6) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 4) = Round((lValor * _
                                    mfCoeJuros(0, 6, lPer_Juros_0_6)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 3) = Round((lValor * mfCoeJuros(0, 7, lPer_Juros_0_7) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 4) = Round((lValor * _
                                    mfCoeJuros(0, 7, lPer_Juros_0_7)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 3) = Round((lValor * mfCoeJuros(0, 8, lPer_Juros_0_8) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 4) = Round((lValor * _
                                    mfCoeJuros(0, 8, lPer_Juros_0_8)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 3) = Round((lValor * mfCoeJuros(0, 9, lPer_Juros_0_9) + _
                                    lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 4) = Round((lValor * _
                                    mfCoeJuros(0, 9, lPer_Juros_0_9)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 3) = Round((lValor * mfCoeJuros(0, 10, lPer_Juros_0_10) + _
                                     lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 4) = Round((lValor * _
                                     mfCoeJuros(0, 10, lPer_Juros_0_10)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 3) = Round((lValor * mfCoeJuros(0, 11, lPer_Juros_0_11) + _
                                     lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 4) = Round((lValor * _
                                     mfCoeJuros(0, 11, lPer_Juros_0_11)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 3) = Round((lValor * mfCoeJuros(0, 12, lPer_Juros_0_12) + _
                                     lCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 4) = Round((lValor * _
                                     mfCoeJuros(0, 12, lPer_Juros_0_12)) * (1 + pProJurIOF / 100), 2)
        End If
        '===>>> IOF.
        If Not IsMissing(pIOF) Then
            pIOF = lIOF
        End If

        Dim lP00Multi_Tplano As ADODB.Connection
        Dim lrstRegraPlano As ADODB.Recordset
        Dim lMensagem As String

        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lP00Multi_Tplano, lMensagem) = False Then Call M46V999.gpGraLog(2, lMensagem): GoTo SAIDA
        lSelect = "SELECT COD_REGRA,VAL_REGRA FROM TMB_REGRA_PLANO "
        lSelect = lSelect & " Where COD_RAMO = " & pStcPed.Cod_Ramo
        lSelect = lSelect & " AND COD_MODALIDADE = " & pStcPed.Cod_Estip
        lSelect = lSelect & " AND COD_SEQUENCIA <=" & pStcPed.Cod_Sequencia
        lSelect = lSelect & " AND COD_TARIF=" & pStcPed.Cod_Produto
        lSelect = lSelect & " Order By Cod_Regra"

        Set lrstRegraPlano = New ADODB.Recordset
        lrstRegraPlano.CursorLocation = adUseClient
        Call lrstRegraPlano.Open(lSelect, lP00Multi_Tplano, adOpenForwardOnly, adLockReadOnly)    'Este recordSet esta usando a propriedade Filter (nao implementado no YasRecordSet)
        Set lrstRegraPlano.ActiveConnection = Nothing

        'forma de pagamento de acordo com o plano se estiver cadastrado
        If lrstRegraPlano.RecordCount > 0 Then
            lI2 = 1
            For lI1 = 20 To 42 Step 2
                'Parcela à Vista
                lrstRegraPlano.Filter = "Cod_Regra = " & lI1
                If lrstRegraPlano.RecordCount = 0 Then
                    pParcela(lI2, 1) = 0
                    pParcela(lI2, 2) = 0
                End If
                lrstRegraPlano.Filter = adFilterNone

                'Parcela à 30 Dias
                lrstRegraPlano.Filter = "Cod_Regra = " & lI1 + 1
                If lrstRegraPlano.RecordCount = 0 Then
                    pParcela(lI2, 3) = 0
                    pParcela(lI2, 4) = 0
                End If
                lrstRegraPlano.Filter = adFilterNone
                lI2 = lI2 + 1
            Next lI1
        End If
        Call M46V999.gpFecha2(lP00Multi_Tplano)

        'Correta e National
        '        If (M46V111.gCod_Plano >= 2001 And M46V111.gCod_Plano <= 2010) Or (M46V111.gCod_Plano >= 2032 And M46V111.gCod_Plano <= 2041) Then
        '            pParcela(2, 1) = 0
        '            pParcela(2, 2) = 0
        '            pParcela(4, 1) = 0
        '            pParcela(4, 2) = 0
        '            pParcela(6, 1) = 0
        '            pParcela(6, 2) = 0
        '            pParcela(7, 1) = 0
        '            pParcela(7, 2) = 0
        '            pParcela(8, 1) = 0
        '            pParcela(8, 2) = 0
        '            pParcela(9, 1) = 0
        '            pParcela(9, 2) = 0
        '            pParcela(10, 1) = 0
        '            pParcela(10, 2) = 0
        '            pParcela(11, 1) = 0
        '            pParcela(11, 2) = 0
        '            pParcela(12, 1) = 0
        '            pParcela(12, 2) = 0
        '        End If
        'GalCorr
        '        If M46V111.gCod_Plano >= 2011 And M46V111.gCod_Plano <= 2020 Then
        '            pParcela(1, 1) = 0
        '            pParcela(2, 1) = 0
        '            pParcela(2, 2) = 0
        '            pParcela(3, 1) = 0
        '            pParcela(3, 2) = 0
        '            pParcela(4, 1) = 0
        '            pParcela(4, 2) = 0
        '            pParcela(5, 1) = 0
        '            pParcela(5, 2) = 0
        '            pParcela(6, 1) = 0
        '            pParcela(6, 2) = 0
        '            pParcela(7, 1) = 0
        '            pParcela(7, 2) = 0
        '            pParcela(8, 1) = 0
        '            pParcela(8, 2) = 0
        '            pParcela(9, 1) = 0
        '            pParcela(9, 2) = 0
        '            pParcela(10, 1) = 0
        '            pParcela(10, 2) = 0
        '            pParcela(11, 1) = 0
        '            pParcela(11, 2) = 0
        '            pParcela(12, 1) = 0
        '            pParcela(12, 2) = 0
        '        End If

        'Mercedes
        '        If M46V111.gCod_Plano >= 2021 And M46V111.gCod_Plano <= 2031 Then
        '            pParcela(2, 1) = 0
        '            pParcela(2, 2) = 0
        '            pParcela(3, 1) = 0
        '            pParcela(3, 2) = 0
        '            pParcela(4, 1) = 0
        '            pParcela(4, 2) = 0
        '            pParcela(6, 1) = 0
        '            pParcela(6, 2) = 0
        '            pParcela(7, 1) = 0
        '            pParcela(7, 2) = 0
        '            pParcela(8, 1) = 0
        '            pParcela(8, 2) = 0
        '            pParcela(9, 1) = 0
        '            pParcela(9, 2) = 0
        '            pParcela(10, 1) = 0
        '            pParcela(10, 2) = 0
        '            pParcela(12, 1) = 0
        '            pParcela(12, 2) = 0
        '        End If

        'Marsh - WorkSite
        '        If M46V111.gCod_Plano >= 2051 And M46V111.gCod_Plano <= 2060 Then
        '            pParcela(3, 1) = 0
        '            pParcela(3, 2) = 0
        '            pParcela(5, 1) = 0
        '            pParcela(5, 2) = 0
        '            pParcela(6, 1) = 0
        '            pParcela(6, 2) = 0
        '            pParcela(7, 1) = 0
        '            pParcela(7, 2) = 0
        '            pParcela(8, 1) = 0
        '            pParcela(8, 2) = 0
        '            pParcela(9, 1) = 0
        '            pParcela(9, 2) = 0
        '            pParcela(10, 1) = 0
        '            pParcela(10, 2) = 0
        '            pParcela(11, 1) = 0
        '            pParcela(11, 2) = 0
        '            pParcela(12, 1) = 0
        '            pParcela(12, 2) = 0
        '        End If

        'Fichas 901282 e 901346 e 901673 e 901772
        '        If (M46V111.gCod_Plano >= 2061 And M46V111.gCod_Plano <= 2100) Or (M46V111.gCod_Plano >= 2111 And M46V111.gCod_Plano <= 2130) Then
        '            pParcela(3, 1) = 0
        '            pParcela(3, 2) = 0
        '            pParcela(5, 1) = 0
        '            pParcela(5, 2) = 0
        '            pParcela(6, 1) = 0
        '            pParcela(6, 2) = 0
        '            pParcela(7, 1) = 0
        '            pParcela(7, 2) = 0
        '            pParcela(8, 1) = 0
        '            pParcela(8, 2) = 0
        '            pParcela(9, 1) = 0
        '            pParcela(9, 2) = 0
        '            pParcela(10, 1) = 0
        '            pParcela(10, 2) = 0
        '            pParcela(11, 1) = 0
        '            pParcela(11, 2) = 0
        '            pParcela(12, 1) = 0
        '            pParcela(12, 2) = 0
        '        End If

        'Correta e National
        '        If (M46V111.gCod_Plano >= 2001 And M46V111.gCod_Plano <= 2010) Or (M46V111.gCod_Plano >= 2032 And M46V111.gCod_Plano <= 2041) Then
        '            pParcela(1, 3) = 0
        '            pParcela(2, 3) = 0
        '            pParcela(2, 4) = 0
        '            pParcela(3, 3) = 0
        '            pParcela(3, 4) = 0
        '            pParcela(4, 3) = 0
        '            pParcela(4, 4) = 0
        '            pParcela(5, 3) = 0
        '            pParcela(5, 4) = 0
        '            pParcela(6, 3) = 0
        '            pParcela(6, 4) = 0
        '            pParcela(7, 3) = 0
        '            pParcela(7, 4) = 0
        '            pParcela(8, 3) = 0
        '            pParcela(8, 4) = 0
        '            pParcela(9, 3) = 0
        '            pParcela(9, 4) = 0
        '            pParcela(10, 3) = 0
        '            pParcela(10, 4) = 0
        '            pParcela(11, 3) = 0
        '            pParcela(11, 4) = 0
        '            pParcela(12, 3) = 0
        '            pParcela(12, 4) = 0
        '        End If
        'GalCorr
        '        If M46V111.gCod_Plano >= 2011 And M46V111.gCod_Plano <= 2020 Then
        '            pParcela(2, 3) = 0
        '            pParcela(2, 4) = 0
        '            pParcela(4, 3) = 0
        '            pParcela(4, 4) = 0
        '            pParcela(5, 3) = 0
        '            pParcela(5, 4) = 0
        '            pParcela(7, 3) = 0
        '            pParcela(7, 4) = 0
        '            pParcela(8, 3) = 0
        '            pParcela(8, 4) = 0
        '            pParcela(9, 3) = 0
        '            pParcela(9, 4) = 0
        '            pParcela(10, 3) = 0
        '            pParcela(10, 4) = 0
        '            pParcela(11, 3) = 0
        '            pParcela(11, 4) = 0
        '            pParcela(12, 3) = 0
        '            pParcela(12, 4) = 0
        '        End If

        'Mercedes
        '        If M46V111.gCod_Plano >= 2021 And M46V111.gCod_Plano <= 2031 Then
        '            pParcela(2, 3) = 0
        '            pParcela(2, 4) = 0
        '            pParcela(3, 3) = 0
        '            pParcela(3, 4) = 0
        '            pParcela(4, 3) = 0
        '            pParcela(4, 4) = 0
        '            pParcela(6, 3) = 0
        '            pParcela(6, 4) = 0
        '            pParcela(7, 3) = 0
        '            pParcela(7, 4) = 0
        '            pParcela(8, 3) = 0
        '            pParcela(8, 4) = 0
        '            pParcela(9, 3) = 0
        '            pParcela(9, 4) = 0
        '            pParcela(10, 3) = 0
        '            pParcela(10, 4) = 0
        '
        '            lUltParcela = 0
        '
        '            If IsDate(lIniVigencia) Then
        '                lIniVigencia = M46V999.gfForData(1, lIniVigencia)
        '            End If
        '            If IsDate(lTerVigencia) Then
        '                lTerVigencia = M46V999.gfForData(1, lTerVigencia)
        '            End If
        '            Call mpObtemMaxParcelas(M46V999.gProQuaMaxParcelas, lParcVista, lParcs30dia, lIniVigencia, lTerVigencia, lUltParcela)
        '
        '            If M46V999.gfPreenchido(M46V999.gProDatUltParcela) = True Then
        '               If M46V999.gProDatUltParcela > 0 Then
        '                  If M46V999.gProDatUltParcela > lTerVigencia Then
        '                     pParcela(11, 3) = 0
        '                     pParcela(11, 4) = 0
        '                  End If
        '                End If
        '            Else
        '               If lUltParcela > CLng(lTerVigencia) Then
        '                  pParcela(11, 3) = 0
        '                  pParcela(11, 4) = 0
        '               End If
        '            End If
        '
        '            pParcela(12, 3) = 0
        '            pParcela(12, 4) = 0
        '        End If

        'Marsh - WorkSite
        '        If M46V111.gCod_Plano >= 2051 And M46V111.gCod_Plano <= 2060 Then
        '            pParcela(3, 3) = 0
        '            pParcela(3, 4) = 0
        '            pParcela(5, 3) = 0
        '            pParcela(5, 4) = 0
        '            pParcela(6, 3) = 0
        '            pParcela(6, 4) = 0
        '            pParcela(7, 3) = 0
        '            pParcela(7, 4) = 0
        '            pParcela(8, 3) = 0
        '            pParcela(8, 4) = 0
        '            pParcela(9, 3) = 0
        '            pParcela(9, 4) = 0
        '            pParcela(10, 3) = 0
        '            pParcela(10, 4) = 0
        '            pParcela(11, 3) = 0
        '            pParcela(11, 4) = 0
        '            pParcela(12, 3) = 0
        '            pParcela(12, 4) = 0
        '        End If

        'Fichas 901282 e 901346 e 901673 e 901172
        '        If (M46V111.gCod_Plano >= 2061 And M46V111.gCod_Plano <= 2100) Or (M46V111.gCod_Plano >= 2111 And M46V111.gCod_Plano <= 2130) Then
        '            pParcela(3, 3) = 0
        '            pParcela(3, 4) = 0
        '            pParcela(5, 3) = 0
        '            pParcela(5, 4) = 0
        '            pParcela(6, 3) = 0
        '            pParcela(6, 4) = 0
        '            pParcela(7, 3) = 0
        '            pParcela(7, 4) = 0
        '            pParcela(8, 3) = 0
        '            pParcela(8, 4) = 0
        '            pParcela(9, 3) = 0
        '            pParcela(9, 4) = 0
        '            pParcela(10, 3) = 0
        '            pParcela(10, 4) = 0
        '            pParcela(11, 3) = 0
        '            pParcela(11, 4) = 0
        '            pParcela(12, 3) = 0
        '            pParcela(12, 4) = 0
        '        End If

        'Marsh - Manaus
        '        If M46V111.gCod_Plano >= 2042 And M46V111.gCod_Plano <= 2049 Then
        ''Ficha 900992
        ''            pParcela(1, 1) = 0
        ''            pParcela(1, 2) = 0
        ''            pParcela(2, 1) = 0
        ''            pParcela(2, 2) = 0
        ''            pParcela(3, 1) = 0
        ''            pParcela(3, 2) = 0
        '
        '            pParcela(4, 1) = 0
        '            pParcela(4, 2) = 0
        '            pParcela(5, 1) = 0
        '            pParcela(5, 2) = 0
        '            pParcela(6, 1) = 0
        '            pParcela(6, 2) = 0
        '            pParcela(7, 1) = 0
        '            pParcela(7, 2) = 0
        '            pParcela(8, 1) = 0
        '            pParcela(8, 2) = 0
        '            pParcela(9, 1) = 0
        '            pParcela(9, 2) = 0
        '            pParcela(10, 1) = 0
        '            pParcela(10, 2) = 0
        '            pParcela(11, 1) = 0
        '            pParcela(11, 2) = 0
        '            pParcela(12, 1) = 0
        '            pParcela(12, 2) = 0
        '        End If


        '===>>> IOF.
        If pStcPed.ColecaoPedVenc.Count > 0 Then
            lTemAjusteVenc = True
        End If

        If (lParcVista = 0 And lParcs30dia = 0) Or (M46V999.gTipoAcesso = DIGITACAO And lTemAjusteVenc = True) Then
            lParcVista = 12
            lParcs30dia = 12
        End If
        If pParcela(1, 1) < 0 Then
            lParcVista = 1
            lParcs30dia = 0
        End If
        If pParcela(1, 1) < lValorMin And pParcela(1, 1) > 0 Then
            If pParcela(1, 1) < lPremioMin Then
                pParcela(1, 1) = 0
            End If
        End If
        For lI1 = 2 To 12
            If pParcela(lI1, 2) < lValorMin Or lParcVista < lI1 Then
                pParcela(lI1, 2) = 0
                pParcela(lI1, 1) = 0
            End If
            If pParcela(lI1, 4) < lValorMin Or lParcs30dia < lI1 Then
                pParcela(lI1, 4) = 0
                pParcela(lI1, 3) = 0
            End If
        Next

        'Merge com berçot
        'Exceção: SALÃO DE BELEZA
        'OBTEM O CORRETOR DA TELA POIS QUANDO O ACESSO VEM DO KIT A VARIAVEL ESTA ZERADA
        If UCase(App.EXEName) = "P46V101" Then
            Dim lI As Integer
            'Utilizado latebinding para não ter problema de compilação na DLL
            For lI = 0 To Forms.Count - 1
                If UCase(Forms(lI).Name) = UCase("frmT46V102A") Then
                    M46V999.gCod_Corretor = Forms(lI).txtCorretor(1).Tag
                    Exit For
                End If
            Next lI
        End If

        'não exibe o parcelamento de 30 dias
        If pStcPed.Cod_Segmentacao <> 13 And _
           M46V999.M46V111.gCod_plano < 1000 And _
           M46V999.gCodProduto >= 1072 And _
           pStcPed.Cod_Dizer = 0 Then

            If M46V999.gCodProduto < 1074 Then
                For lI1 = 1 To 12
                    pParcela(lI1, 4) = 0
                    pParcela(lI1, 3) = 0
                Next
            Else
                If ((M46V999.gCod_Corretor <> 1601) Or (M46V999.gCod_Corretor = 1601 And M46V999.gCod_Ramo <> e_CodRamo_Residencial)) And _
                   ((M46V999.gCod_Corretor <> 383111) Or (M46V999.gCod_Corretor = 383111 And M46V999.gCod_Ramo = e_CodRamo_Condominio)) Then
                    For lI1 = 1 To 12
                        pParcela(lI1, 4) = 0
                        pParcela(lI1, 3) = 0
                    Next
                End If
            End If
        End If

    End If



    If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
        '022. Forma de pagamento: primeira parcela.
        If pStcPed.Qtd_Parc > lParcVista Then
            pStcPed.Qtd_Parc = lParcVista
        End If

        pStcPed.Val_PriParcela = pParcela(pStcPed.Qtd_Parc, 1)
        '022. Forma de pagamento: demais parcelas.
        pStcPed.val_DemParcelas = pParcela(pStcPed.Qtd_Parc, 2)
    Else
        If pStcPed.Qtd_Parc > lParcs30dia Then
            pStcPed.Qtd_Parc = lParcs30dia
        End If
        '023. Forma de pagamento: primeira parcela.
        pStcPed.Val_PriParcela = pParcela(pStcPed.Qtd_Parc, 3)
        '023. Forma de pagamento: demais parcelas.
        pStcPed.val_DemParcelas = pParcela(pStcPed.Qtd_Parc, 4)
    End If

    For lI1 = 1 To pStcPed.getProduto.ProQuaMaxParcelas
        '        If pParcela(lI1, 1) = 0 Then Exit For
        Set objStcParcela = New stcParcelamento
        objStcParcela.nosso_numero = pStcPed.nosso_numero
        objStcParcela.Num_Parcela = lI1
        objStcParcela.Val_Primeira_Vista = CDbl(Format(pParcela(lI1, 1), "standard"))
        objStcParcela.Val_Primeira_30dias = CDbl(Format(pParcela(lI1, 3), "standard"))
        objStcParcela.Val_Demais_Vista = CDbl(Format(pParcela(lI1, 2), "standard"))
        objStcParcela.Val_Demais_30dias = CDbl(Format(pParcela(lI1, 4), "standard"))
        If lI1 < pStcPed.ColecaoPedJuros.Count Then
            Set lobjstcJuros = pStcPed.ColecaoPedJuros(lI1)
            objStcParcela.Val_Juros_Vista = CDbl(Format(lobjstcJuros.Val_Ant, "standard"))
            objStcParcela.Val_Juros_30dias = CDbl(Format(lobjstcJuros.Val_Pos, "standard"))
        End If
        Set objStcParcela.ObjPai = pStcPed
        pStcPed.ColecaoParcelamento.Add objStcParcela
    Next lI1

    gfCalFracionamentoOld2 = True

SAIDA:
    GoSub FechaObj
    Exit Function

TrataErro:
    If M46V999.gfPreenchido(pMensagem) Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "gfCalFracionamento Erro:" & pMensagem
    GoTo SAIDA

FechaObj:
    Set lobjstcJuros = Nothing
    Set objStcParcela = Nothing
    Set lP00Multi_Tplano = Nothing
    Set lrstRegraPlano = Nothing
    Return
End Function


Public Function gfCalFracionamento(ByRef pStcPed As stcA46V702B, _
                                   ByVal pTip_Emissao As Long, _
                                   ByRef pParcela As Variant, ByRef pMensagem As String, _
                                   ByVal pProJurIOF As Single, Optional ByVal pCodReg As Integer, _
                                   Optional ByRef pIOF As Single, Optional ByVal pParcMin As Single)

'Novo retornando juros com formas de pagamento

'1100303 - Arredondamento no Parcelamento - 'Foi realizado troca de algumas variáveis de Double para Single e tambem substituido a função M46V999.gfTrunca() pela Round() - caso queira comparar historico de alteração no SouceSafe

'Função: calcula fracionamento.

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "gfCalFracionamento"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(M46V999.tagDynaTrace, methodName, Class_Name)
    
    On Error GoTo TrataErro

    gfCalFracionamento = False

'Parâmetros de entrada...pStcPed ..........Stc do pedido
'                        pTip_Emissao......Tipo de Emissão
'                        pCod_Registro.....Código do Registro
'                        pProJurIOF........ Juros IOF
'                        pCodReg........... Codigo de registro (Opcional)
'                        pParcMin.......... Valor minimo de parcela
'Parâmetro de saída......pParcela.........matriz com resultados das parcelas.
'                        pMensagem........mensagem.
'                        pErro............lista de erros.
'                        pIOF.............Valor do IOF para valor determinado (opcional).
    Dim lI1 As Byte    'Utilizado com For...Next.
    Dim lI2 As Byte    'Utilizado com For...Next.
    Dim lI3 As Byte    'Utilizado com For...Next.
    Dim lTip As Byte    'Utilizado para Forma de Pagamento
    Dim lValorMin As Single    'Valor mínimo da parcela ou endosso.
    Dim lPremioMin As Single  'Valor mínimo do prêmio de parcela ou endosso.
    Dim lParcVista As Integer    'Quantidade de parcelas máxima possível em caso de pagamento antecipado.
    Dim lParcs30dia As Integer  'Quandidade de parcelas máxima possível em caso de paramento não antecipado.
    Dim lParcsCartao As Integer  'Quandidade de parcelas máxima possível em caso de paramento cartao de credito.
    Dim lPreTotal As Single    'Prêmio total.
    Dim lValor As Single    'Prêmio Liquido.
    Dim lIOF As Single    'Valor IOF.
    Dim lPer_Juros_1 As Variant   'Juros da 1ª parcela.
    Dim lPer_Juros_2 As Variant    'Juros da 2ª parcela.
    Dim lPer_Juros_3 As Variant    'Juros da 3ª parcela.
    Dim lPer_Juros_4 As Variant    'Juros da 4ª parcela.
    Dim lPer_Juros_5 As Variant    'Juros da 5ª parcela.
    Dim lPer_Juros_6 As Variant    'Juros da 6ª parcela.
    Dim lPer_Juros_7 As Variant    'Juros da 7ª parcela.
    Dim lPer_Juros_8 As Variant    'Juros da 8ª parcela.
    Dim lPer_Juros_9 As Variant    'Juros da 9ª parcela.
    Dim lPer_Juros_10 As Variant     'Juros da 10ª parcela.
    Dim lPer_Juros_11 As Variant     'Juros da 11ª parcela.
    Dim lPer_Juros_12 As Variant     'Juros da 12ª parcela.
    Dim lPer_Juros_0_1 As Variant      'Juros da 1ª parcela antecipada.
    Dim lPer_Juros_0_2 As Variant      'Juros da 2ª parcela antecipada.
    Dim lPer_Juros_0_3 As Variant      'Juros da 3ª parcela antecipada.
    Dim lPer_Juros_0_4 As Variant      'Juros da 4ª parcela antecipada.
    Dim lPer_Juros_0_5 As Variant      'Juros da 5ª parcela antecipada.
    Dim lPer_Juros_0_6 As Variant      'Juros da 6ª parcela antecipada.
    Dim lPer_Juros_0_7 As Variant      'Juros da 7ª parcela antecipada.
    Dim lPer_Juros_0_8 As Variant      'Juros da 8ª parcela antecipada.
    Dim lPer_Juros_0_9 As Variant      'Juros da 9ª parcela antecipada.
    Dim lPer_Juros_0_10 As Variant       'Juros da 10ª parcela antecipada.
    Dim lPer_Juros_0_11 As Variant       'Juros da 11ª parcela antecipada.
    Dim lPer_Juros_0_12 As Variant       'Juros da 12ª parcela antecipada.
    Dim lobjstcJuros As stcA46V711B
    Dim lTemAjusteVenc As Boolean
    Dim lSelect As String
    Dim lMensagem As String

    Dim lIniVigencia As String
    Dim lTerVigencia As String
    Dim lCusEmissao As Single
    Dim lCusParcelado As String
    Dim lPreLiquido As Single
    Dim lCod_Iden_Tarf As Long

    Dim objStcParcela As stcParcelamento

    Dim xAuxErr As String
    Dim stcProduto As stcProduto
    Dim lJurosSegmentados As String
    Dim lobjstcPedLoc As stcA46V708B
    Dim lArrayRetricaoCorr As Variant
    Dim i As Integer
    Dim lMaxParcCorr As Integer
    Dim lMaxParcSeg As Integer
    Dim parcelaCartao As Integer
    Dim QtdeParcCartao As Integer
    
    Set stcProduto = New stcProduto
    Set lobjstcPedLoc = New stcA46V708B

    Set lobjstcPedLoc.ObjPai = pStcPed
    
    lJurosSegmentados = ""
    pStcPed.Per_Juro_Frac = 0
    lMaxParcCorr = 0
    lMaxParcSeg = 7

    'VERIFICA OS ITENS CASO TENHA ALGUM QUE NÃO SEJA SEGMENTADO NÃO UTILIZAR OS JUROS PARAMETRIZADOS.
    For Each lobjstcPedLoc In pStcPed.ColecaoPedLoc
        M46V999.gCod_Segmentacao = 0
        If Trim(lobjstcPedLoc.Cod_Clasf) <> "" Then
                M46V999.gCod_Segmentacao = M46V999.gfObterCodSegmentado(lobjstcPedLoc.Cod_Clasf, lobjstcPedLoc.Cod_Compl_Clasf)

                If M46V999.M46V103.mfAjustesPontuais(e_Tip_Valor_JurosSegmentados, TpoValor, lJurosSegmentados, pMensagem, lobjstcPedLoc, 0, M46V999.tagDynaTrace) = False Then
                        Call M46V999.gpGraLog(0, pMensagem)
                End If
        End If

        If Trim(lJurosSegmentados) = "" Then
            Exit For
        End If
    Next lobjstcPedLoc

    'Obtem o numero de parcelas maxima cadastrado para o corretor apenas web
    If pStcPed.Cod_Produto >= 1086 Then
        If M46V999.gPrograma = "D046O003" And M46V999.gOrigemChamada = e_Origem_Web And pStcPed.Cod_Produto <> 1090 Then
            If M46V999.M46V103.mfAjustesPontuais(e_Tip_Valor_ParcMaxCorretor, TpoValor, lMaxParcCorr, pMensagem, lobjstcPedLoc, 0, M46V999.tagDynaTrace) = False Then
                Call M46V999.gpGraLog(0, pMensagem)
            End If
        Else
            If (pStcPed.Cod_Estip >= 2350 And pStcPed.Cod_Estip <= 2373) Or _
               (pStcPed.Cod_Estip = 2303 And pStcPed.Cod_Ramo = e_CodRamo_Empresarial) Or _
               (pStcPed.Cod_Estip = 2374 And pStcPed.Cod_Ramo = e_CodRamo_Residencial) Then
                lMaxParcCorr = 11
            ElseIf pStcPed.Cod_Estip = 20 Then
                lMaxParcCorr = 12
            Else
                lMaxParcCorr = 11 'Kit liberado o parcelamento padrão em 11x
            End If
        End If
    End If

    '001 - Apólice com cobrança de emolumentos.
    '002 - Apólice sem cobrança de emolumentos.
    '003 - Apólice sem cobrança de IOF.
    '004 - Apólice sem cobrança de custo de apólice.
    '005 - Apólice com moeda estrangeira (ramo transporte/crédito de exportação).
    '006 - Apólice coletiva de acidentes pessoais - pagamento mensal.
    '007 - Apólice coletiva - pagamento trimestral.
    '008 - Apólice coletiva - pagamento semestral.
    '020 - Apólice a averbar (sem cobrança de prêmio depósito).
    '101 - Endosso com cobrança de emolumentos.
    '102 - Endosso sem cobrança de emolumentos.
    '103 - Endosso sem cobrança de IOF.
    '104 - Endosso sem cobrança de custo de apólice.
    '105 - Endosso com moeda estrangeira (ramo transporte/crédito de exportação).
    '106 - Endosso de restituição de prêmio sem cancelamento da apólice.
    '107 - Endosso de restituição de prêmio com cancelamento da apólice.
    '120 - Endosso sem movimento de prêmio.
    '121 - Endosso sem movimento com prêmio negativo.
    '122 - Endosso sem movimento com prêmio positivo.
    '129 - Cancelamento da apólice sem estorno de prêmio.

    ReDim lPer_Juros_1(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_2(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_3(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_4(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_5(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_6(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_7(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_8(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_9(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_10(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_11(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_12(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_0_1(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_0_2(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_0_3(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_0_4(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_0_5(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_0_6(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_0_7(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_0_8(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_0_9(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_0_10(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_0_11(M46V999.gQtdMaxComp)
    ReDim lPer_Juros_0_12(M46V999.gQtdMaxComp)



    'Inicio os valores

    lIniVigencia = M46V999.gfForData(2, pStcPed.Dat_Ini_Vig)
    lTerVigencia = M46V999.gfForData(2, pStcPed.Dat_Ter_Vig)

    M46V999.gParcMaxCorr = 0
    If lMaxParcCorr <> 0 And pStcPed.Cod_Produto >= 1086 Then
        M46V999.gProQuaMaxParcelas = lMaxParcCorr
        M46V999.gParcMaxCorr = lMaxParcCorr
    Else
    M46V999.gProQuaMaxParcelas = pStcPed.getProduto.ProQuaMaxParcelas
    End If

    '8000046529, Ajuste da forma de pagto
    If M46V999.gEhRetornoSiscota(pStcPed.nosso_numero) Then
    
        Dim lsSQl As String
        Dim strParcelas_Impressao_30dias As String
        Dim strParcelas_Impressao_AVista As String
        
        'Verificar se já existe alguma opção de impressão na base.
        lsSQl = "SELECT Parcelas_Impressao_30dias,Parcelas_Impressao_AVista FROM Tab_Ctrl_Emis_Compl (nolock) "
        lsSQl = lsSQl & " WHERE"
        lsSQl = lsSQl & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pStcPed.nosso_numero)
    
        Dim lbdTabCtrl As ADODB.Connection
        Dim lrsTabCtrl As clsYasRecordSet  'Registro : Tab_Ctrl_Emis_Compl
    
        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTabCtrl, lMensagem) = False Then
            Call M46V999.gpGraLog(2, lMensagem)
        End If
    
        If Not M46V999.gfObtRegistro(lbdTabCtrl, lsSQl, lrsTabCtrl, lMensagem) Then
            Call M46V999.gpGraLog(1, lMensagem)
        End If
        
        'Recupera valor.
        If Not lrsTabCtrl.EOF Then
            strParcelas_Impressao_30dias = lrsTabCtrl("Parcelas_Impressao_30dias")
            strParcelas_Impressao_AVista = lrsTabCtrl("Parcelas_Impressao_AVista")
        End If
        
        Debug.Print "strParcelas_Impressao_30dias: " + CStr(strParcelas_Impressao_30dias)
        Debug.Print "strParcelas_Impressao_AVista: " + CStr(strParcelas_Impressao_AVista)
        
    End If
       
    'Debug.Print "mpObtemMaxParcelas: MaxParcelas -> " + CStr(IIf(M46V999.gEhRetornoSiscota(pStcPed.NOSSO_NUMERO), 12, M46V999.gProQuaMaxParcelas))
           
    Call mpObtemMaxParcelas(IIf(M46V999.gEhRetornoSiscota(pStcPed.nosso_numero), 12, M46V999.gProQuaMaxParcelas), lParcVista, lParcs30dia, pStcPed.Dat_Ini_Vig, pStcPed.Dat_Ter_Vig)

        'Obtem quantidade de parcelas para cartão de crédito
    lParcsCartao = mfValidaParcelamentoSemJurosCartaoCredito(pStcPed, pMensagem)

    lCusEmissao = pStcPed.Val_Cust_Emis
    lCusParcelado = IIf(pStcPed.gCustoParcelado = True, "S", "N")
    lPreLiquido = pStcPed.VAL_PR
    lCod_Iden_Tarf = pStcPed.Cod_Produto

    If lPreLiquido = 0 Then
        If pStcPed.Cod_Dizer <> 0 Then
            If pStcPed.Cod_Dizer = 10 Then 'cancelamento
                pStcPed.Cod_Reg = 129
            Else
                pStcPed.Cod_Reg = 120
            End If
        Else
            If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
                pStcPed.Cod_Reg = 4
            Else
                pStcPed.Cod_Reg = 20
            End If
        End If
        pStcPed.Val_Cust_Emis = 0
    ElseIf lPreLiquido > 0 Then
        If pStcPed.Cod_Dizer = 0 Then
            If pStcPed.Cod_Reg > 100 Or pStcPed.Cod_Reg = 0 Then
                pStcPed.Cod_Reg = 1
            End If
        Else
            If pStcPed.Cod_Reg < 100 Or pStcPed.Cod_Reg > 105 Then
                pStcPed.Cod_Reg = 101
            End If
        End If
    Else
        'Valor negativo
        If pStcPed.Cod_Dizer = 10 Then
            pStcPed.Cod_Reg = 107
        Else
            pStcPed.Cod_Reg = 106
        End If
        pStcPed.Val_Cust_Emis = 0
    End If

    If pStcPed.Dat_Ini_Vig >= 20130101 Then
        pStcPed.Val_Cust_Emis = 0
        If pStcPed.Cod_Reg = 1 Then
            pStcPed.Cod_Reg = 4
        ElseIf pStcPed.Cod_Reg = 3 Then
            pStcPed.Cod_Reg = 2
        ElseIf pStcPed.Cod_Reg = 101 Then
            pStcPed.Cod_Reg = 104
        ElseIf pStcPed.Cod_Reg = 103 Then
            pStcPed.Cod_Reg = 102
        End If
    Else
        If pStcPed.Cod_Reg = 4 Then
            pStcPed.Cod_Reg = 1
        ElseIf pStcPed.Cod_Reg = 2 Then
            pStcPed.Cod_Reg = 3
        ElseIf pStcPed.Cod_Reg = 104 Then
            pStcPed.Cod_Reg = 101
        ElseIf pStcPed.Cod_Reg = 102 Then
            pStcPed.Cod_Reg = 103
        End If
    End If
    lCusEmissao = pStcPed.Val_Cust_Emis

    'Retirar regras do segmentado
    'Não está ohando tipo de compensação
'    If lJurosSegmentados <> "" And pStcPed.Cod_Produto >= 1082 And M46V999.gOrigemChamada = e_Origem_Kit Then
'        For Each lobjstcJuros In pStcPed.ColecaoPedJuros
'            If lobjstcJuros.Num_Parcela <= lMaxParcSeg Then
'                lobjstcJuros.Val_Ant = lJurosSegmentados
'            End If
'        Next lobjstcJuros
'    Else
'        For lI1 = 0 To M46V999.gQtdMaxComp - 1
'            lPer_Juros_1(lI1) = 3
'            lPer_Juros_2(lI1) = 3
'            lPer_Juros_3(lI1) = 3
'            lPer_Juros_4(lI1) = 3
'            lPer_Juros_5(lI1) = 3
'            lPer_Juros_6(lI1) = 3
'            lPer_Juros_7(lI1) = 3
'            lPer_Juros_8(lI1) = 3
'            lPer_Juros_9(lI1) = 3
'            lPer_Juros_10(lI1) = 3
'            lPer_Juros_11(lI1) = 3
'            lPer_Juros_12(lI1) = 3
'            lPer_Juros_0_1(lI1) = 3
'            lPer_Juros_0_2(lI1) = 3
'            lPer_Juros_0_3(lI1) = 3
'            lPer_Juros_0_4(lI1) = 3
'            lPer_Juros_0_5(lI1) = 3
'            lPer_Juros_0_6(lI1) = 3
'            lPer_Juros_0_7(lI1) = 3
'            lPer_Juros_0_8(lI1) = 3
'            lPer_Juros_0_9(lI1) = 3
'            lPer_Juros_0_10(lI1) = 3
'            lPer_Juros_0_11(lI1) = 3
'            lPer_Juros_0_12(lI1) = 3
'        Next lI1
'    End If


    'Incluido Simony
    If pStcPed.Cod_Reg = 2 Or _
       pStcPed.Cod_Reg = 3 Or _
       pStcPed.Cod_Reg = 102 Or _
       pStcPed.Cod_Reg = 103 Or _
       pStcPed.Cod_Reg = 105 Or _
       pStcPed.Cod_Reg = 106 Or _
       pStcPed.Cod_Reg = 107 Or _
       pStcPed.Cod_Reg = 120 Or _
       pStcPed.Cod_Reg = 202 Or _
       pStcPed.Cod_Reg = 203 Or _
       pStcPed.Cod_Reg = 129 Then
        pProJurIOF = 0
        pIOF = 0
    End If
    'Fim


    For Each lobjstcJuros In pStcPed.ColecaoPedJuros
        'Vai incluir no tipo de compensação correspondente ao objeto
        lTip = lobjstcJuros.Tip_Compensacao
        If lTip = 0 Then lTip = e_FormaPagamento_Carne
        Select Case lobjstcJuros.Num_Parcela
        Case 1
            lPer_Juros_0_1(lTip) = lobjstcJuros.Val_Pos
            lPer_Juros_1(lTip) = lobjstcJuros.Val_Ant
        Case 2
            lPer_Juros_0_2(lTip) = lobjstcJuros.Val_Pos
            lPer_Juros_2(lTip) = lobjstcJuros.Val_Ant
        Case 3
            lPer_Juros_0_3(lTip) = lobjstcJuros.Val_Pos
            lPer_Juros_3(lTip) = lobjstcJuros.Val_Ant
        Case 4
            lPer_Juros_0_4(lTip) = lobjstcJuros.Val_Pos
            lPer_Juros_4(lTip) = lobjstcJuros.Val_Ant
        Case 5
            lPer_Juros_0_5(lTip) = lobjstcJuros.Val_Pos
            lPer_Juros_5(lTip) = lobjstcJuros.Val_Ant
        Case 6
            lPer_Juros_0_6(lTip) = lobjstcJuros.Val_Pos
            lPer_Juros_6(lTip) = lobjstcJuros.Val_Ant
        Case 7
            lPer_Juros_0_7(lTip) = lobjstcJuros.Val_Pos
            lPer_Juros_7(lTip) = lobjstcJuros.Val_Ant
        Case 8
            lPer_Juros_0_8(lTip) = lobjstcJuros.Val_Pos
            lPer_Juros_8(lTip) = lobjstcJuros.Val_Ant
        Case 9
            lPer_Juros_0_9(lTip) = lobjstcJuros.Val_Pos
            lPer_Juros_9(lTip) = lobjstcJuros.Val_Ant
        Case 10
            lPer_Juros_0_10(lTip) = lobjstcJuros.Val_Pos
            lPer_Juros_10(lTip) = lobjstcJuros.Val_Ant
        Case 11
            lPer_Juros_0_11(lTip) = lobjstcJuros.Val_Pos
            lPer_Juros_11(lTip) = lobjstcJuros.Val_Ant
        Case 12
            lPer_Juros_0_12(lTip) = lobjstcJuros.Val_Pos
            lPer_Juros_12(lTip) = lobjstcJuros.Val_Ant
        End Select
    Next lobjstcJuros

    '11.38 - Forma de pagamento 0+4 para renovações Empresariais porto
    If mfValidaCongenerePortoPostecipadoSemJuros(pStcPed, pTip_Emissao, pMensagem) Then
        For Each lobjstcJuros In pStcPed.ColecaoPedJuros
            lTip = lobjstcJuros.Tip_Compensacao
            If lTip = 0 Then lTip = e_FormaPagamento_Carne
            Select Case lobjstcJuros.Num_Parcela
            Case 1
                lobjstcJuros.Val_Pos = 0
                lPer_Juros_0_1(lTip) = 0
            Case 2
                lobjstcJuros.Val_Pos = 0
                lPer_Juros_0_2(lTip) = 0
            Case 3
                lobjstcJuros.Val_Pos = 0
                lPer_Juros_0_3(lTip) = 0
            Case 4
                lobjstcJuros.Val_Pos = 0
                lPer_Juros_0_4(lTip) = 0
            End Select
        Next lobjstcJuros
    End If
    
    If M46V999.gfPreenchido(pMensagem) Then
        GoTo TrataErro
    End If
    
'    e_FormaPagamento_ND = 0
'    e_FormaPagamento_Carne = 1
'    e_FormaPagamento_Debito = 2
'    e_FormaPagamento_Credito = 3

    '11.72 - Forma de pagamento 1+9 para corretores ATHINA A C SEG
    If mfAtribuiUmaMaisNoveSemJurosCorretoraAthina(pStcPed, pMensagem) Then
        For Each lobjstcJuros In pStcPed.ColecaoPedJuros
            lTip = lobjstcJuros.Tip_Compensacao
            If lTip = 0 Then lTip = e_FormaPagamento_Carne
            
            If lTip = e_FormaPagamento_Carne Then
            
                Select Case lobjstcJuros.Num_Parcela
                Case 1
                    lobjstcJuros.Val_Ant = 0
                    lPer_Juros_0_1(lTip) = 0
                    
                    lobjstcJuros.Val_Pos = 0
                    lPer_Juros_1(lTip) = 0
                Case 2
                    lobjstcJuros.Val_Ant = 0
                    lPer_Juros_0_2(lTip) = 0
                    
                    lobjstcJuros.Val_Pos = 0
                    lPer_Juros_2(lTip) = 0
                Case 3
                    lobjstcJuros.Val_Ant = 0
                    lPer_Juros_0_3(lTip) = 0
                    
                    lobjstcJuros.Val_Pos = 0
                    lPer_Juros_3(lTip) = 0
                Case 4
                    lobjstcJuros.Val_Ant = 0
                    lPer_Juros_0_4(lTip) = 0
                    
                    lobjstcJuros.Val_Pos = 0
                    lPer_Juros_4(lTip) = 0
                Case 5
                    lobjstcJuros.Val_Ant = 0
                    lPer_Juros_0_5(lTip) = 0
                    
                    lobjstcJuros.Val_Pos = 0
                    lPer_Juros_5(lTip) = 0
                Case 6
                    lobjstcJuros.Val_Ant = 0
                    lPer_Juros_0_6(lTip) = 0
                    
                    lobjstcJuros.Val_Pos = 0
                    lPer_Juros_6(lTip) = 0
                Case 7
                    lobjstcJuros.Val_Ant = 0
                    lPer_Juros_0_7(lTip) = 0
                    
                    lobjstcJuros.Val_Pos = 0
                    lPer_Juros_7(lTip) = 0
                Case 8
                    lobjstcJuros.Val_Ant = 0
                    lPer_Juros_0_8(lTip) = 0
                    
                    lobjstcJuros.Val_Pos = 0
                    lPer_Juros_8(lTip) = 0
                Case 9
                    lobjstcJuros.Val_Ant = 0
                    lPer_Juros_0_9(lTip) = 0
                    
                    lobjstcJuros.Val_Pos = 0
                    lPer_Juros_9(lTip) = 0
                Case 10
                    lobjstcJuros.Val_Ant = 0
                    lPer_Juros_0_10(lTip) = 0
                    
                    lobjstcJuros.Val_Pos = 0
                    lPer_Juros_10(lTip) = 0
                End Select
                
            End If
        Next lobjstcJuros
    End If
    
    If M46V999.gfPreenchido(pMensagem) Then
        GoTo TrataErro
    End If
    
    '2. Prepara planilha de resultados.
    ReDim pParcela(12, 4, M46V999.gQtdMaxComp)
    For lI1 = 1 To 12
        For lI2 = 1 To 4
            For lI3 = 1 To M46V999.gQtdMaxComp
                pParcela(lI1, lI2, lI3) = 0
            Next lI3
        Next lI2
    Next lI1

    '3. Calcula fracionamento e posiciona resultados na planilha.
    lValor = lPreLiquido
    If lValor <> 0 Then
        '****************************************
        'Ajute retirar premio minimo do imobiliario
        If pStcPed.Cod_Estip < 2000 Then
            If pStcPed.Cod_Reg < 100 And pStcPed.Cod_Estip < 2000 Then
                lPremioMin = M46V999.gProPremioMinimo
            Else
                If M46V999.gPermitePrInformado Or pStcPed.Cod_Estip = 20 Then
                    lPremioMin = lValor
                Else
                    If pStcPed.Cod_Estip <> 20 Then
                                lPremioMin = M46V999.gProPreMinEndosso
                    Else
                        lPremioMin = lValor
                    End If
                End If
            End If
        End If
        If M46V999.gCodProduto >= 976 And lValor < lPremioMin Then
            If pStcPed.Cod_Reg = 106 Or pStcPed.Cod_Reg = 107 Then
                If lValor > lPremioMin Then
                    lValor = 0
                End If
            ElseIf pStcPed.Cod_Reg = 101 Or pStcPed.Cod_Reg = 102 Or pStcPed.Cod_Reg = 103 Or pStcPed.Cod_Reg = 104 Then
                If lValor < lPremioMin Then
                    lValor = 0
                End If
            Else
                lValor = lPremioMin    'prêmio mínimo
            End If
        End If

        lIOF = Round((lValor + lCusEmissao) * (pProJurIOF / 100), 2)
        lPreTotal = lValor + lCusEmissao + lIOF

        lValorMin = pParcMin

        If lValor < lPremioMin And lValor > 0 Then
            lValor = 0
            lPreTotal = 0
            lCusEmissao = 0
            lIOF = 0
        End If
        '2. Calcula valores das parcelas.
        'c) Posiciona resultados.
        '===>>> Parcelas.
        For lI3 = 1 To M46V999.gQtdMaxComp
            If lCusParcelado = "S" Then
                'a) Considerando custo de emissão parcelado.
                'Adicionado juros na 1º parcela novo produto
                pParcela(1, 1, lI3) = Round((lValor * mfCoeJuros(1, 1, lPer_Juros_1(lI3)) + _
                                             (lCusEmissao / 2)) * (1 + pProJurIOF / 100), 2)

                pParcela(2, 1, lI3) = Round((lValor * mfCoeJuros(1, 2, lPer_Juros_2(lI3)) + _
                                             (lCusEmissao / 2)) * (1 + pProJurIOF / 100), 2)
                pParcela(2, 2, lI3) = pParcela(2, 1, lI3)

                pParcela(3, 1, lI3) = Round((lValor * mfCoeJuros(1, 3, lPer_Juros_3(lI3)) + _
                                             (lCusEmissao / 3)) * (1 + pProJurIOF / 100), 2)
                pParcela(3, 2, lI3) = pParcela(3, 1, lI3)

                pParcela(4, 1, lI3) = Round((lValor * mfCoeJuros(1, 4, lPer_Juros_4(lI3)) + _
                                             (lCusEmissao / 4)) * (1 + pProJurIOF / 100), 2)
                pParcela(4, 2, lI3) = pParcela(4, 1, lI3)
                pParcela(5, 1, lI3) = Round((lValor * mfCoeJuros(1, 5, lPer_Juros_5(lI3)) + _
                                             (lCusEmissao / 5)) * (1 + pProJurIOF / 100), 2)
                pParcela(5, 2, lI3) = pParcela(5, 1, lI3)
                pParcela(6, 1, lI3) = Round((lValor * mfCoeJuros(1, 6, lPer_Juros_6(lI3)) + _
                                             (lCusEmissao / 6)) * (1 + pProJurIOF / 100), 2)
                pParcela(6, 2, lI3) = pParcela(6, 1, lI3)
                pParcela(7, 1, lI3) = Round((lValor * mfCoeJuros(1, 7, lPer_Juros_7(lI3)) + _
                                             (lCusEmissao / 7)) * (1 + pProJurIOF / 100), 2)
                pParcela(7, 2, lI3) = pParcela(7, 1, lI3)
                pParcela(8, 1, lI3) = Round((lValor * mfCoeJuros(1, 8, lPer_Juros_8(lI3)) + _
                                             (lCusEmissao / 8)) * (1 + pProJurIOF / 100), 2)
                pParcela(8, 2, lI3) = pParcela(8, 1, lI3)
                pParcela(9, 1, lI3) = Round((lValor * mfCoeJuros(1, 9, lPer_Juros_9(lI3)) + _
                                             (lCusEmissao / 9)) * (1 + pProJurIOF / 100), 2)
                pParcela(9, 2, lI3) = pParcela(9, 1, lI3)
                pParcela(10, 1, lI3) = Round((lValor * mfCoeJuros(1, 10, lPer_Juros_10(lI3)) + _
                                              (lCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
                pParcela(10, 2, lI3) = pParcela(10, 1, lI3)
                pParcela(11, 1, lI3) = Round((lValor * mfCoeJuros(1, 11, lPer_Juros_11(lI3)) + _
                                              (lCusEmissao / 11)) * (1 + pProJurIOF / 100), 2)
                pParcela(11, 2, lI3) = pParcela(11, 1, lI3)
                pParcela(12, 1, lI3) = Round((lValor * mfCoeJuros(1, 12, lPer_Juros_12(lI3)) + _
                                              (lCusEmissao / 12)) * (1 + pProJurIOF / 100), 2)
                pParcela(12, 2, lI3) = pParcela(12, 1, lI3)


                pParcela(1, 3, lI3) = Round((lValor * mfCoeJuros(0, 1, lPer_Juros_0_1(lI3)) + _
                                             (lCusEmissao / 1)) * (1 + pProJurIOF / 100), 2)
                pParcela(2, 3, lI3) = Round((lValor * mfCoeJuros(0, 2, lPer_Juros_0_2(lI3)) + _
                                             (lCusEmissao / 2)) * (1 + pProJurIOF / 100), 2)
                pParcela(2, 4, lI3) = pParcela(2, 3, lI3)
                pParcela(3, 3, lI3) = Round((lValor * mfCoeJuros(0, 3, lPer_Juros_0_3(lI3)) + _
                                             (lCusEmissao / 3)) * (1 + pProJurIOF / 100), 2)
                pParcela(3, 4, lI3) = pParcela(3, 3, lI3)
                pParcela(4, 3, lI3) = Round((lValor * mfCoeJuros(0, 4, lPer_Juros_0_4(lI3)) + _
                                             (lCusEmissao / 4)) * (1 + pProJurIOF / 100), 2)
                pParcela(4, 4, lI3) = pParcela(4, 3, lI3)
                pParcela(5, 3, lI3) = Round((lValor * mfCoeJuros(0, 5, lPer_Juros_0_5(lI3)) + _
                                             (lCusEmissao / 5)) * (1 + pProJurIOF / 100), 2)
                pParcela(5, 4, lI3) = pParcela(5, 3, lI3)
                pParcela(6, 3, lI3) = Round((lValor * mfCoeJuros(0, 6, lPer_Juros_0_6(lI3)) + _
                                             (lCusEmissao / 6)) * (1 + pProJurIOF / 100), 2)
                pParcela(6, 4, lI3) = pParcela(6, 3, lI3)
                pParcela(7, 3, lI3) = Round((lValor * mfCoeJuros(0, 7, lPer_Juros_0_7(lI3)) + _
                                             (lCusEmissao / 7)) * (1 + pProJurIOF / 100), 2)
                pParcela(7, 4, lI3) = pParcela(7, 3, lI3)
                pParcela(8, 3, lI3) = Round((lValor * mfCoeJuros(0, 8, lPer_Juros_0_8(lI3)) + _
                                             (lCusEmissao / 8)) * (1 + pProJurIOF / 100), 2)
                pParcela(8, 4, lI3) = pParcela(8, 3, lI3)
                pParcela(9, 3, lI3) = Round((lValor * mfCoeJuros(0, 9, lPer_Juros_0_9(lI3)) + _
                                             (lCusEmissao / 9)) * (1 + pProJurIOF / 100), 2)
                pParcela(9, 4, lI3) = pParcela(9, 3, lI3)
                pParcela(10, 3, lI3) = Round((lValor * mfCoeJuros(0, 10, lPer_Juros_0_10(lI3)) + _
                                              (lCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
                pParcela(10, 4, lI3) = pParcela(10, 3, lI3)

                pParcela(11, 3, lI3) = Round((lValor * mfCoeJuros(0, 11, lPer_Juros_0_11(lI3)) + _
                                              (lCusEmissao / 11)) * (1 + pProJurIOF / 100), 2)
                pParcela(11, 4, lI3) = pParcela(11, 3, lI3)
                pParcela(12, 3, lI3) = Round((lValor * mfCoeJuros(0, 12, lPer_Juros_0_12(lI3)) + _
                                              (lCusEmissao / 12)) * (1 + pProJurIOF / 100), 2)
                pParcela(12, 4, lI3) = pParcela(12, 3, lI3)
            Else
                'b) Considerando custo de emissão NÃO parcelado.
                'Adicionado juros na 1º parcela novo produto
                If pStcPed.Cod_Produto >= 1080 Then
                    pParcela(1, 1, lI3) = Round((lValor * mfCoeJuros(1, 1, lPer_Juros_1(lI3)) + _
                                                 lCusEmissao) * (1 + pProJurIOF / 100), 2)
                Else
                        pParcela(1, 1, lI3) = Round((lValor + lCusEmissao) * (1 + pProJurIOF / 100), 2)
                End If
                
                pParcela(2, 1, lI3) = Round((lValor * mfCoeJuros(1, 2, lPer_Juros_2(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(2, 2, lI3) = Round((lValor * _
                                             mfCoeJuros(1, 2, lPer_Juros_2(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(3, 1, lI3) = Round((lValor * mfCoeJuros(1, 3, lPer_Juros_3(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(3, 2, lI3) = Round((lValor * _
                                             mfCoeJuros(1, 3, lPer_Juros_3(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(4, 1, lI3) = Round((lValor * mfCoeJuros(1, 4, lPer_Juros_4(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(4, 2, lI3) = Round((lValor * _
                                             mfCoeJuros(1, 4, lPer_Juros_4(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(5, 1, lI3) = Round((lValor * mfCoeJuros(1, 5, lPer_Juros_5(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(5, 2, lI3) = Round((lValor * _
                                             mfCoeJuros(1, 5, lPer_Juros_5(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(6, 1, lI3) = Round((lValor * mfCoeJuros(1, 6, lPer_Juros_6(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(6, 2, lI3) = Round((lValor * _
                                             mfCoeJuros(1, 6, lPer_Juros_6(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(7, 1, lI3) = Round((lValor * mfCoeJuros(1, 7, lPer_Juros_7(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(7, 2, lI3) = Round((lValor * _
                                             mfCoeJuros(1, 7, lPer_Juros_7(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(8, 1, lI3) = Round((lValor * mfCoeJuros(1, 8, lPer_Juros_8(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(8, 2, lI3) = Round((lValor * _
                                             mfCoeJuros(1, 8, lPer_Juros_8(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(9, 1, lI3) = Round((lValor * mfCoeJuros(1, 9, lPer_Juros_9(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(9, 2, lI3) = Round((lValor * _
                                             mfCoeJuros(1, 9, lPer_Juros_9(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(10, 1, lI3) = Round((lValor * mfCoeJuros(1, 10, lPer_Juros_10(lI3)) + _
                                              lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(10, 2, lI3) = Round((lValor * _
                                              mfCoeJuros(1, 10, lPer_Juros_10(lI3))) * (1 + pProJurIOF / 100), 2)

                pParcela(11, 1, lI3) = Round((lValor * mfCoeJuros(1, 11, lPer_Juros_11(lI3)) + _
                                              lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(11, 2, lI3) = Round((lValor * _
                                              mfCoeJuros(1, 11, lPer_Juros_11(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(12, 1, lI3) = Round((lValor * mfCoeJuros(1, 12, lPer_Juros_12(lI3)) + _
                                              lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(12, 2, lI3) = Round((lValor * _
                                              mfCoeJuros(1, 12, lPer_Juros_12(lI3))) * (1 + pProJurIOF / 100), 2)

                pParcela(1, 3, lI3) = Round((lValor * mfCoeJuros(0, 1, lPer_Juros_0_1(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(2, 3, lI3) = Round((lValor * mfCoeJuros(0, 2, lPer_Juros_0_2(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(2, 4, lI3) = Round((lValor * _
                                             mfCoeJuros(0, 2, lPer_Juros_0_2(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(3, 3, lI3) = Round((lValor * mfCoeJuros(0, 3, lPer_Juros_0_3(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(3, 4, lI3) = Round((lValor * _
                                             mfCoeJuros(0, 3, lPer_Juros_0_3(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(4, 3, lI3) = Round((lValor * mfCoeJuros(0, 4, lPer_Juros_0_4(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(4, 4, lI3) = Round((lValor * _
                                             mfCoeJuros(0, 4, lPer_Juros_0_4(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(5, 3, lI3) = Round((lValor * mfCoeJuros(0, 5, lPer_Juros_0_5(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(5, 4, lI3) = Round((lValor * _
                                             mfCoeJuros(0, 5, lPer_Juros_0_5(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(6, 3, lI3) = Round((lValor * mfCoeJuros(0, 6, lPer_Juros_0_6(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(6, 4, lI3) = Round((lValor * _
                                             mfCoeJuros(0, 6, lPer_Juros_0_6(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(7, 3, lI3) = Round((lValor * mfCoeJuros(0, 7, lPer_Juros_0_7(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(7, 4, lI3) = Round((lValor * _
                                             mfCoeJuros(0, 7, lPer_Juros_0_7(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(8, 3, lI3) = Round((lValor * mfCoeJuros(0, 8, lPer_Juros_0_8(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(8, 4, lI3) = Round((lValor * _
                                             mfCoeJuros(0, 8, lPer_Juros_0_8(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(9, 3, lI3) = Round((lValor * mfCoeJuros(0, 9, lPer_Juros_0_9(lI3)) + _
                                             lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(9, 4, lI3) = Round((lValor * _
                                             mfCoeJuros(0, 9, lPer_Juros_0_9(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(10, 3, lI3) = Round((lValor * mfCoeJuros(0, 10, lPer_Juros_0_10(lI3)) + _
                                              lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(10, 4, lI3) = Round((lValor * _
                                              mfCoeJuros(0, 10, lPer_Juros_0_10(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(11, 3, lI3) = Round((lValor * mfCoeJuros(0, 11, lPer_Juros_0_11(lI3)) + _
                                              lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(11, 4, lI3) = Round((lValor * _
                                              mfCoeJuros(0, 11, lPer_Juros_0_11(lI3))) * (1 + pProJurIOF / 100), 2)
                pParcela(12, 3, lI3) = Round((lValor * mfCoeJuros(0, 12, lPer_Juros_0_12(lI3)) + _
                                              lCusEmissao) * (1 + pProJurIOF / 100), 2)
                pParcela(12, 4, lI3) = Round((lValor * _
                                              mfCoeJuros(0, 12, lPer_Juros_0_12(lI3))) * (1 + pProJurIOF / 100), 2)

            End If
        Next
        '===>>> IOF.
        If Not IsMissing(pIOF) Then
            pIOF = lIOF
        End If

        Dim lP00Multi_Tplano As ADODB.Connection
        Dim lrstRegraPlano As ADODB.Recordset

        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lP00Multi_Tplano, lMensagem) = False Then
            Call M46V999.gpGraLog(2, lMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        lSelect = "SELECT COD_REGRA,VAL_REGRA FROM TMB_REGRA_PLANO "
        lSelect = lSelect & " Where COD_RAMO = " & pStcPed.Cod_Ramo
        lSelect = lSelect & " AND COD_MODALIDADE = " & pStcPed.Cod_Estip
        lSelect = lSelect & " AND COD_SEQUENCIA <=" & pStcPed.Cod_Sequencia
        lSelect = lSelect & " AND COD_TARIF=" & pStcPed.Cod_Produto
        lSelect = lSelect & " Order By Cod_Regra"

        Set lrstRegraPlano = New ADODB.Recordset
        lrstRegraPlano.CursorLocation = adUseClient
        Call lrstRegraPlano.Open(lSelect, lP00Multi_Tplano, adOpenForwardOnly, adLockReadOnly)    'Este recordSet esta usando a propriedade Filter (nao implementado no YasRecordSet)
        Set lrstRegraPlano.ActiveConnection = Nothing

        Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "INFO: Abriu e desconectou RecordSet lrstRegraPlano", 0, False)
            
        'forma de pagamento de acordo com o plano se estiver cadastrado
        If lrstRegraPlano.RecordCount > 0 Then
            For lI3 = 1 To M46V999.gQtdMaxComp
                lI2 = 1
                For lI1 = 20 To 42 Step 2
                    'Parcela à Vista
                    lrstRegraPlano.Filter = "Cod_Regra = " & lI1
                    If lrstRegraPlano.RecordCount = 0 Then
'                        If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim And lI2 <= pStcPed.getProduto.ProQuaMaxParcelas Then
'                            pMensagem = "Forma de pagamento (Antecipada) não permitida para este planinho."
'                            Exit Function
'                        End If
                        pParcela(lI2, 1, lI3) = 0
                        pParcela(lI2, 2, lI3) = 0
                    End If
                    lrstRegraPlano.Filter = adFilterNone

                    'Parcela à 30 Dias
                    lrstRegraPlano.Filter = "Cod_Regra = " & lI1 + 1
                    If lrstRegraPlano.RecordCount = 0 Then
'                        If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Nao And lI2 <= pStcPed.getProduto.ProQuaMaxParcelas Then
'                            pMensagem = "Forma de pagamento (Postecipada) não permitida para este planinho."
'                            Exit Function
'                        End If
                        pParcela(lI2, 3, lI3) = 0
                        pParcela(lI2, 4, lI3) = 0
                    End If
                    lrstRegraPlano.Filter = adFilterNone
                    lI2 = lI2 + 1
                Next lI1
            Next lI3
        End If
        Call M46V999.gpFecha2(lP00Multi_Tplano)

        '8000046529, Ajuste da forma de pagto
        If Not (strParcelas_Impressao_30dias = Empty) Then
            'Zera as parcelas que estão acima do parametrizado
            Debug.Print "Zera as parcelas que estão acima do parametrizado: 30dias " + strParcelas_Impressao_30dias
            For lI3 = 1 To M46V999.gQtdMaxComp
                For lI1 = CInt(Split(strParcelas_Impressao_30dias, ",")(UBound(Split(strParcelas_Impressao_30dias, ",")))) + 1 To 12
                    pParcela(lI1, 4, lI3) = 0
                    pParcela(lI1, 3, lI3) = 0
                    Debug.Print "Zerou parcela: " + CStr(lI1)
                Next
            Next
        End If
        
        If Not (strParcelas_Impressao_AVista = Empty) Then
            'Zera as parcelas que estão acima do parametrizado
            Debug.Print "Zera as parcelas que estão acima do parametrizado: AVista " + strParcelas_Impressao_AVista
            For lI3 = 1 To M46V999.gQtdMaxComp
                For lI1 = CInt(Split(strParcelas_Impressao_AVista, ",")(UBound(Split(strParcelas_Impressao_AVista, ",")))) + 1 To 12
                    pParcela(lI1, 1, lI3) = 0
                    pParcela(lI1, 2, lI3) = 0
                    Debug.Print "Zerou parcela: " + CStr(lI1)
                Next
            Next
        End If
        
        '11.72 - Forma de pagamento 1+9 para corretores ATHINA A C SEG
        If mfAtribuiUmaMaisNoveSemJurosCorretoraAthina(pStcPed, pMensagem) Then
            'For lI3 = 1 To M46V999.gQtdMaxComp
            For lI1 = 11 To 12
                pParcela(lI1, 1, 1) = 0
                pParcela(lI1, 2, 1) = 0
                pParcela(lI1, 3, 1) = 0
                pParcela(lI1, 4, 1) = 0
            Next
            'Next
        End If

        '===>>> IOF.
        If pStcPed.ColecaoPedVenc.Count > 0 Then
            lTemAjusteVenc = True
        End If

        If (lParcVista = 0 And lParcs30dia = 0) Or (M46V999.gTipoAcesso = DIGITACAO And lTemAjusteVenc = True) Then
            lParcVista = 12
            lParcs30dia = 12
        End If
        If pParcela(1, 1, 1) < 0 Or pStcPed.Cod_Reg = 106 Or pStcPed.Cod_Reg = 107 Then
            lParcVista = 1
            lParcs30dia = 0
        End If
                
        ' Aplicar premio minimo por parcela
        '8000045890, Erro Syas - Restituição com parcelamento
        If pStcPed.Cod_Estip < 2000 And (pStcPed.Cod_Estip <> 20 Or lPremioMin < 0) Then
            For lI3 = 1 To M46V999.gQtdMaxComp
                If pParcela(1, 1, lI3) < lValorMin And pParcela(1, 1, lI3) > 0 Then
                    If pParcela(1, 1, lI3) < lPremioMin Then
                        pParcela(1, 1, lI3) = 0
                    End If
                End If
                For lI1 = 2 To 12
                    If pParcela(lI1, 2, lI3) < lValorMin Or lParcVista < lI1 Then
                        pParcela(lI1, 2, lI3) = 0
                        pParcela(lI1, 1, lI3) = 0
                    End If
                    If pParcela(lI1, 4, lI3) < lValorMin Or lParcs30dia < lI1 Then
                        pParcela(lI1, 4, lI3) = 0
                        pParcela(lI1, 3, lI3) = 0
                    End If
                Next
            Next
        End If
    End If

    Dim ZerarParcela30dias As Boolean
    ZerarParcela30dias = False

    Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "INFO: ZerarParcela30dias", 0, False)

    'Exceção: Cod_Segmentacao: 9 - Outros Empresarias / 10 - Residencial / 12 - Condominio e endosso
    'If (pStcPed.Cod_Segmentacao = 0 Or pStcPed.Cod_Segmentacao = 9 Or pStcPed.Cod_Segmentacao = 10 Or pStcPed.Cod_Segmentacao = 12) _
     '        And M46V999.M46V111.gCod_Plano < 1000 And M46V999.gCodProduto >= 1072 And pStcPed.Cod_Dizer = 0 Then  'não exibe o parcelamento de 30 dias
    'não exibe o parcelamento de 30 dias

    'Merge com berçot
    'Exceção: SALÃO DE BELEZA
    'OBTEM O CORRETOR DA TELA POIS QUANDO O ACESSO VEM DO KIT A VARIAVEL ESTA ZERADA
    If UCase(App.EXEName) = "P46V101" Then
        Dim lI As Integer
        'Utilizado latebinding para não ter problema de compilação na DLL
        For lI = 0 To Forms.Count - 1
            If UCase(Forms(lI).Name) = UCase("frmT46V102A") Then
                M46V999.gCod_Corretor = Forms(lI).txtCorretor(1).Tag
                Exit For
            End If
        Next lI
    End If

    If pStcPed.Cod_Estip < 2000 And _
       M46V999.gCodProduto >= 1072 And _
       ( _
        (pStcPed.Cod_Dizer = 0 And M46V999.gCodProduto < 1080) Or _
        ( _
            pTip_Emissao <> 200 And _
            pStcPed.objstcCtrlEmis.cod_origem <> 2 And _
            pStcPed.Cod_Dizer >= 0 And _
            pStcPed.Cod_Ramo <> e_CodRamo_Condominio And _
            M46V999.gCodProduto >= 1080 _
         ) _
       ) Then

        If M46V999.gPermitePostecipada = False Then
                If M46V999.gCodProduto < 1074 Then
                    ZerarParcela30dias = True
                    'For lI1 = 1 To 12
                    'pParcela(lI1, 4) = 0
                    'pParcela(lI1, 3) = 0
                    'Next
                Else
                    If ((M46V999.gCod_Corretor <> 1601) Or (M46V999.gCod_Corretor = 1601 And M46V999.gCod_Ramo <> e_CodRamo_Residencial)) And _
                       ((M46V999.gCod_Corretor <> 383111) Or (M46V999.gCod_Corretor = 383111 And M46V999.gCod_Ramo = e_CodRamo_Condominio)) Then
                        ZerarParcela30dias = True
                        'For lI1 = 1 To 12
                        '    pParcela(lI1, 4) = 0
                        '    pParcela(lI1, 3) = 0
                        'Next
                    End If
                End If
                
        End If
    End If
    
    'Tarifa 11.68 - Item 10 - Condição de 0+ Seguros Novos e Congêneres Corretora Félix
    If (M46V999.gCod_Corretor = 933311 Or M46V999.gCod_Corretor = 723911) And (pStcPed.Cod_Estip = 10 Or pStcPed.Cod_Estip = 11) And pStcPed.Cod_Ramo = 112 Then
        ZerarParcela30dias = False
    End If

    If pStcPed.Cod_Produto >= 1082 And pStcPed.Cod_Estip = 2300 Then 'ipiranga
        ZerarParcela30dias = True
    End If
        
    'Item 1 Nova tarifa 1086 / 1087
    'Incluido 0 + para o corretor 919096
    If pStcPed.Cod_Produto >= 1086 Then
        If pStcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            lArrayRetricaoCorr = Split("1675;17048;723911;860301;919096", ";")
        Else
            lArrayRetricaoCorr = Split("860301;919096", ";")
        End If
        
        For i = 0 To UBound(lArrayRetricaoCorr)
            If M46V999.gCod_Corretor = lArrayRetricaoCorr(i) Then
                ZerarParcela30dias = False
            End If
        Next
        
    End If
    
    'Item 4 Nova tarifa 1086 / 1087 parcelamento maximo por corretor
    'Liberar 0 + para digitação manual maritima
    If lMaxParcCorr <> 0 Or pStcPed.ColecaoPedLoc(1).Ren_Con_Seguradora = 572 Then
        ZerarParcela30dias = False
    End If
    
    'Permitir postecipado para o segmentado quando renovação
    If pStcPed.getPlano.EhPlanoSegmentado Then
        If pTip_Emissao = 200 Then
            ZerarParcela30dias = False
        Else
            ZerarParcela30dias = True
        End If
    End If
    
    '11.38 - Forma de pagamento 0+4 para renovações Empresariais porto
    Call mfAtribuiCongenerePortoPostecipadoSemJuros(pStcPed, pTip_Emissao, pParcela, ZerarParcela30dias, pMensagem)
    If M46V999.gfPreenchido(pMensagem) Then
        GoTo TrataErro
    End If
    
    If ZerarParcela30dias Then
        For lI3 = 1 To M46V999.gQtdMaxComp
            For lI1 = 1 To 12
                pParcela(lI1, 4, lI3) = 0
                pParcela(lI1, 3, lI3) = 0
            Next
        Next
    Else
        'parcelamento apenas para renovação tarifa 1080
        If (pStcPed.Num_Apol_An > 0 Or pStcPed.objstcCtrlEmis.cod_origem = 2) And _
           pStcPed.Cod_Dizer = 0 And _
           M46V999.gCodProduto >= 1080 And _
           pStcPed.Cod_Estip < 2000 And _
           lMaxParcCorr = 0 Then
            Select Case pStcPed.Cod_Ramo
                Case e_CodRamo_Empresarial
                    If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 2 Then
                        If pStcPed.Qtd_Parc > 11 Then
                                pStcPed.Qtd_Parc = 11
                        End If
                    Else
                        If pStcPed.Qtd_Parc > pStcPed.getProduto.ProQuaMaxParcelas Then
                                pStcPed.Qtd_Parc = pStcPed.getProduto.ProQuaMaxParcelas
                        End If
                    End If
                    
                    For lI3 = 1 To M46V999.gQtdMaxComp
                        For lI1 = 11 To 12
                            pParcela(lI1, 4, lI3) = 0
                            pParcela(lI1, 3, lI3) = 0
                        Next
                    Next
                Case e_CodRamo_Condominio
                    If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 2 Then
                        If pStcPed.Qtd_Parc > 11 Then
                                pStcPed.Qtd_Parc = 11
                        End If
                    Else
                       If pStcPed.Qtd_Parc > pStcPed.getProduto.ProQuaMaxParcelas Then
                                pStcPed.Qtd_Parc = pStcPed.getProduto.ProQuaMaxParcelas
                        End If
                    End If
                    
                    For lI3 = 1 To M46V999.gQtdMaxComp
                        For lI1 = 11 To 12
                            pParcela(lI1, 4, lI3) = 0
                            pParcela(lI1, 3, lI3) = 0
                        Next
                    Next
                Case e_CodRamo_Residencial
                    If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 2 Then
                        If pStcPed.Qtd_Parc > 4 Then
                                pStcPed.Qtd_Parc = 4
                        End If
                    Else
                        If pStcPed.Qtd_Parc > pStcPed.getProduto.ProQuaMaxParcelas Then
                                pStcPed.Qtd_Parc = pStcPed.getProduto.ProQuaMaxParcelas
                        End If
                    End If
                    
                    For lI3 = 1 To M46V999.gQtdMaxComp
                        For lI1 = 5 To 12
                            pParcela(lI1, 4, lI3) = 0
                            pParcela(lI1, 3, lI3) = 0
                        Next
                    Next
                Case Else
                    For lI3 = 1 To M46V999.gQtdMaxComp
                        For lI1 = 1 To 12
                            pParcela(lI1, 4, lI3) = 0
                            pParcela(lI1, 3, lI3) = 0
                        Next
                    Next
                
            End Select
        End If
    End If

    lTip = pStcPed.gTipPagamento

    If lTip = 0 Then lTip = e_FormaPagamento_Carne

    If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
        '022. Forma de pagamento: primeira parcela.
        If pStcPed.Qtd_Parc > lParcVista Then
            pStcPed.Qtd_Parc = lParcVista
        End If

        If lMaxParcCorr <> 0 Then
            For i = 2 To 12
                If pParcela(i, 2, lTip) = 0 And pStcPed.Cod_Tip_Doc <> e_TipDoc_Proposta Then
                    pStcPed.Qtd_Parc = i - 1
                    Exit For
                End If
            Next i
        End If

        pStcPed.Val_PriParcela = pParcela(pStcPed.Qtd_Parc, 1, lTip)
        '022. Forma de pagamento: demais parcelas.
        pStcPed.val_DemParcelas = pParcela(pStcPed.Qtd_Parc, 2, lTip)
    Else
        If pStcPed.Qtd_Parc > lParcs30dia Then
            pStcPed.Qtd_Parc = lParcs30dia
        End If
        
        If lMaxParcCorr <> 0 Then
            For i = 2 To 12
                If pParcela(i, 2, lTip) = 0 And pStcPed.Cod_Tip_Doc <> e_TipDoc_Proposta Then
                    pStcPed.Qtd_Parc = i - 1
                    Exit For
                End If
            Next i
        End If
        
        '023. Forma de pagamento: primeira parcela.
        pStcPed.Val_PriParcela = pParcela(pStcPed.Qtd_Parc, 3, lTip)
        '023. Forma de pagamento: demais parcelas.
        pStcPed.val_DemParcelas = pParcela(pStcPed.Qtd_Parc, 4, lTip)
    End If
       
    'Cartao de Credito - Atribui parcelamento
    Call mfAtribuiParcelamentoCartaoCredito(pParcela, lParcsCartao, True, pMensagem)
    If M46V999.gfPreenchido(pMensagem) Then
        GoTo TrataErro
    End If

    For lI1 = 1 To IIf(M46V999.gEhRetornoSiscota(pStcPed.nosso_numero), 12, pStcPed.getProduto.ProQuaMaxParcelas) 'pStcPed.getProduto.ProQuaMaxParcelas
        '        If pParcela(lI1, 1) = 0 Then Exit For
        For lI2 = 1 To M46V999.gQtdMaxComp
            Set objStcParcela = New stcParcelamento
            objStcParcela.nosso_numero = pStcPed.nosso_numero
            objStcParcela.Num_Parcela = lI1
            objStcParcela.TipPagamento = lI2

            objStcParcela.Val_Primeira_Vista = CDbl(Format(pParcela(lI1, 1, lI2), "standard"))
            objStcParcela.Val_Primeira_30dias = CDbl(Format(pParcela(lI1, 3, lI2), "standard"))
            objStcParcela.Val_Demais_Vista = CDbl(Format(pParcela(lI1, 2, lI2), "standard"))
            objStcParcela.Val_Demais_30dias = CDbl(Format(pParcela(lI1, 4, lI2), "standard"))
            If lI1 < pStcPed.ColecaoPedJuros.Count Then
                For lI3 = 1 To pStcPed.ColecaoPedJuros.Count
                    Set lobjstcJuros = pStcPed.ColecaoPedJuros(lI3)
                    If lobjstcJuros.Num_Parcela = lI1 And lobjstcJuros.Tip_Compensacao = lI2 Then
                        objStcParcela.Val_Juros_Vista = CDbl(Format(lobjstcJuros.Val_Ant, "standard"))
                        objStcParcela.Val_Juros_30dias = CDbl(Format(lobjstcJuros.Val_Pos, "standard"))
                        Exit For
                    End If
                Next lI3
            End If
            Set objStcParcela.ObjPai = pStcPed
            pStcPed.ColecaoParcelamento.Add objStcParcela
        Next lI2
    Next lI1
    
    pStcPed.Per_Juro_Frac = CDbl(Replace(pStcPed.getRpt.Per_Juros_Neg, "%", ""))

    gfCalFracionamento = True

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    GoSub FechaObj
    Exit Function

TrataErro:

    xAuxErr = "gfCalFracionamento " & Err & " " & ERROR
    If M46V999.gfPreenchido(pMensagem) Then
        pMensagem = xAuxErr
    End If
    pMensagem = "gfCalFracionamento Erro:" & pMensagem
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
    GoTo SAIDA
    Resume

FechaObj:
    Set lobjstcJuros = Nothing
    Set objStcParcela = Nothing
    Set lP00Multi_Tplano = Nothing
    Set lrstRegraPlano = Nothing
    Return
End Function


Public Function gfOldCalFracionamento(ByVal pQuaParcelas As Byte, ByVal pIniVigencia As String, _
                                      ByVal pCusEmissao As Single, ByVal pCusParcelado As String, _
                                      ByVal pPreLiquido As Single, pColecaoJuros As Collection, _
                                      ByVal pTip_Emissao As Long, ByVal pCod_Iden_Tarf As Long, _
                                      ByRef pParcela As Variant, ByRef pMensagem As String, _
                                      ByVal pProJurIOF As Double, Optional ByRef pIOF As Single) As Boolean

'Função: calcula fracionamento.

'Parâmetros de entrada...pQuaParcelas.....quantidade de parcelas.
'                        pIniVigencia.....início de vigência.
'                        pCusEmissao......custo de emissão.
'                        pCusParcelado....indicador de custo parcelado: S = Sim ou N = Não.
'                        pVD_PreLiquido...valor determinado - valor do prêmio líquido.
'                        pVM_PreLiquido...valor de mercado - valor do prêmio líquido.
'                        pJurAntxx........juros para pagamento antecipado - xx parcelas.
'                        pJurPosxx........juros para pagamento postecipado - xx parcelas.
'                        pCod_Registro.....Código do Registro
'Parâmetro de saída......pParcela.........matriz com resultados das parcelas.
'                        pMensagem........mensagem.
'                        pErro............lista de erros.
'                        pVD-IOF..........Valor do IOF para valor determinado (opcional).
'                        pVM-IOF..........Valor do IOF para valor de mercado (opcional).
    Dim lI1 As Byte    'Utilizado com For...Next.
    Dim lI2 As Byte    'Utilizado com For...Next.
    Dim lValorMin As Double    'Valor mínimo da parcela ou endosso.
    Dim lPremioMin As Double  'Valor mínimo do prêmio de parcela ou endosso.
    Dim lParcVista As Integer    'Quantidade de parcelas máxima possível em caso de pagamento antecipado.
    Dim lParcs30dia As Integer  'Quandidade de parcelas máxima possível em caso de paramento não antecipado.
    Dim lPreTotal As Double    'Prêmio total.
    Dim lDivCust As Integer    'Número pelo qual será dividido o custo
    Dim lValor As Double    'Prêmio Liquido.
    Dim lIOF As Double    'Valor IOF.
    Dim lPer_Juros_1 As Double    'Juros da 1ª parcela.
    Dim lPer_Juros_2 As Double    'Juros da 2ª parcela.
    Dim lPer_Juros_3 As Double    'Juros da 3ª parcela.
    Dim lPer_Juros_4 As Double    'Juros da 4ª parcela.
    Dim lPer_Juros_5 As Double    'Juros da 5ª parcela.
    Dim lPer_Juros_6 As Double    'Juros da 6ª parcela.
    Dim lPer_Juros_7 As Double    'Juros da 7ª parcela.
    Dim lPer_Juros_8 As Double    'Juros da 8ª parcela.
    Dim lPer_Juros_9 As Double    'Juros da 9ª parcela.
    Dim lPer_Juros_10 As Double     'Juros da 10ª parcela.
    Dim lPer_Juros_11 As Double     'Juros da 11ª parcela.
    Dim lPer_Juros_12 As Double     'Juros da 12ª parcela.
    Dim lPer_Juros_0_1 As Double      'Juros da 1ª parcela antecipada.
    Dim lPer_Juros_0_2 As Double      'Juros da 2ª parcela antecipada.
    Dim lPer_Juros_0_3 As Double      'Juros da 3ª parcela antecipada.
    Dim lPer_Juros_0_4 As Double      'Juros da 4ª parcela antecipada.
    Dim lPer_Juros_0_5 As Double      'Juros da 5ª parcela antecipada.
    Dim lPer_Juros_0_6 As Double      'Juros da 6ª parcela antecipada.
    Dim lPer_Juros_0_7 As Double      'Juros da 7ª parcela antecipada.
    Dim lPer_Juros_0_8 As Double      'Juros da 8ª parcela antecipada.
    Dim lPer_Juros_0_9 As Double      'Juros da 9ª parcela antecipada.
    Dim lPer_Juros_0_10 As Double       'Juros da 10ª parcela antecipada.
    Dim lPer_Juros_0_11 As Double       'Juros da 11ª parcela antecipada.
    Dim lPer_Juros_0_12 As Double       'Juros da 12ª parcela antecipada.
    Dim lobjstcJuros As stcA46V711B

    gfOldCalFracionamento = False

    'Inicio os valores
    lPer_Juros_1 = 3
    lPer_Juros_2 = 3
    lPer_Juros_3 = 3
    lPer_Juros_4 = 3
    lPer_Juros_5 = 3
    lPer_Juros_6 = 3
    lPer_Juros_7 = 3
    lPer_Juros_8 = 3
    lPer_Juros_9 = 3
    lPer_Juros_10 = 3
    lPer_Juros_11 = 3
    lPer_Juros_12 = 3
    lPer_Juros_0_1 = 3
    lPer_Juros_0_2 = 3
    lPer_Juros_0_3 = 3
    lPer_Juros_0_4 = 3
    lPer_Juros_0_5 = 3
    lPer_Juros_0_6 = 3
    lPer_Juros_0_7 = 3
    lPer_Juros_0_8 = 3
    lPer_Juros_0_9 = 3
    lPer_Juros_0_10 = 3
    lPer_Juros_0_11 = 3
    lPer_Juros_0_12 = 3

    For Each lobjstcJuros In pColecaoJuros
        Select Case lobjstcJuros.Num_Parcela
        Case 1
            lPer_Juros_0_1 = lobjstcJuros.Val_Pos
            lPer_Juros_1 = lobjstcJuros.Val_Ant
        Case 2
            lPer_Juros_0_2 = lobjstcJuros.Val_Pos
            lPer_Juros_2 = lobjstcJuros.Val_Ant
        Case 3
            lPer_Juros_0_3 = lobjstcJuros.Val_Pos
            lPer_Juros_3 = lobjstcJuros.Val_Ant
        Case 4
            lPer_Juros_0_4 = lobjstcJuros.Val_Pos
            lPer_Juros_4 = lobjstcJuros.Val_Ant
        Case 5
            lPer_Juros_0_5 = lobjstcJuros.Val_Pos
            lPer_Juros_5 = lobjstcJuros.Val_Ant
        Case 6
            lPer_Juros_0_6 = lobjstcJuros.Val_Pos
            lPer_Juros_6 = lobjstcJuros.Val_Ant
        Case 7
            lPer_Juros_0_7 = lobjstcJuros.Val_Pos
            lPer_Juros_7 = lobjstcJuros.Val_Ant
        Case 8
            lPer_Juros_0_8 = lobjstcJuros.Val_Pos
            lPer_Juros_8 = lobjstcJuros.Val_Ant
        Case 9
            lPer_Juros_0_9 = lobjstcJuros.Val_Pos
            lPer_Juros_9 = lobjstcJuros.Val_Ant
        Case 10
            lPer_Juros_0_10 = lobjstcJuros.Val_Pos
            lPer_Juros_10 = lobjstcJuros.Val_Ant
        Case 11
            lPer_Juros_0_11 = lobjstcJuros.Val_Pos
            lPer_Juros_11 = lobjstcJuros.Val_Ant
        Case 12
            lPer_Juros_0_12 = lobjstcJuros.Val_Pos
            lPer_Juros_12 = lobjstcJuros.Val_Ant
        End Select
    Next lobjstcJuros

    '2. Prepara planilha de resultados.
    ReDim pParcela(12, 4)
    For lI1 = 1 To 12
        For lI2 = 1 To 4
            pParcela(lI1, lI2) = 0
        Next lI2
    Next lI1

    '3. Calcula fracionamento e posiciona resultados na planilha.
    lValor = pPreLiquido
    If lValor <> 0 Then
        '****************************************
        lIOF = Round((lValor + pCusEmissao) * (pProJurIOF / 100), 2)
        lPreTotal = lValor + pCusEmissao + lIOF
        lValorMin = 30
        If lValor < lPremioMin Then
            lValor = 0
            lPreTotal = 0
            pCusEmissao = 0
            lIOF = 0
        End If
        '2. Calcula valores das parcelas.
        'c) Posiciona resultados.
        '===>>> Parcelas.
        If pCusParcelado = "S" Then
            'a) Considerando custo de emissão parcelado.
            pParcela(1, 1) = M46V999.gfTrunca((lValor + pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 2, lPer_Juros_2), 2)) + _
                                               (pCusEmissao / 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 2) = pParcela(2, 1)
            pParcela(3, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 3, lPer_Juros_3), 2)) + _
                                               (pCusEmissao / 3)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 2) = pParcela(3, 1)
            pParcela(4, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 4, lPer_Juros_4), 2)) + _
                                               (pCusEmissao / 4)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 2) = pParcela(4, 1)
            pParcela(5, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 5, lPer_Juros_5), 2)) + _
                                               (pCusEmissao / 5)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 2) = pParcela(5, 1)
            pParcela(6, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 6, lPer_Juros_6), 2)) + _
                                               (pCusEmissao / 6)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 2) = pParcela(6, 1)
            pParcela(7, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 7, lPer_Juros_7), 2)) + _
                                               (pCusEmissao / 7)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 2) = pParcela(7, 1)
            pParcela(8, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 8, lPer_Juros_8), 2)) + _
                                               (pCusEmissao / 8)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 2) = pParcela(8, 1)
            pParcela(9, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 9, lPer_Juros_9), 2)) + _
                                               (pCusEmissao / 9)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 2) = pParcela(9, 1)
            pParcela(10, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 10, lPer_Juros_10), 2)) + _
                                                (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 2) = pParcela(10, 1)
            pParcela(11, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 11, lPer_Juros_11), 2)) + _
                                                (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 2) = pParcela(11, 1)
            pParcela(12, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 12, lPer_Juros_12), 2)) + _
                                                (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 2) = pParcela(12, 1)
            pParcela(1, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 1, lPer_Juros_0_1), 2)) + _
                                               (pCusEmissao / 1)) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 2, lPer_Juros_0_2), 2)) + _
                                               (pCusEmissao / 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 4) = pParcela(2, 3)
            pParcela(3, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 3, lPer_Juros_0_3), 2)) + _
                                               (pCusEmissao / 3)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 4) = pParcela(3, 3)
            pParcela(4, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 4, lPer_Juros_0_4), 2)) + _
                                               (pCusEmissao / 4)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 4) = pParcela(4, 3)
            pParcela(5, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 5, lPer_Juros_0_5), 2)) + _
                                               (pCusEmissao / 5)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 4) = pParcela(5, 3)
            pParcela(6, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 6, lPer_Juros_0_6), 2)) + _
                                               (pCusEmissao / 6)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 4) = pParcela(6, 3)
            pParcela(7, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 7, lPer_Juros_0_7), 2)) + _
                                               (pCusEmissao / 7)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 4) = pParcela(7, 3)
            pParcela(8, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 8, lPer_Juros_0_8), 2)) + _
                                               (pCusEmissao / 8)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 4) = pParcela(8, 3)
            pParcela(9, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 9, lPer_Juros_0_9), 2)) + _
                                               (pCusEmissao / 9)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 4) = pParcela(9, 3)
            pParcela(10, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 10, lPer_Juros_0_10), 2)) + _
                                                (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 4) = pParcela(10, 3)
            pParcela(11, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 11, lPer_Juros_0_11), 2)) + _
                                                (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 4) = pParcela(11, 3)
            pParcela(12, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 12, lPer_Juros_0_12), 2)) + _
                                                (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 4) = pParcela(12, 3)
        Else
            'b) Considerando custo de emissão NÃO parcelado.
            pParcela(1, 1) = M46V999.gfTrunca((lValor + pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 2, lPer_Juros_2), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 2) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(1, 2, lPer_Juros_2), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 3, lPer_Juros_3), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 2) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(1, 3, lPer_Juros_3), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 4, lPer_Juros_4), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 2) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(1, 4, lPer_Juros_4), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 5, lPer_Juros_5), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 2) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(1, 5, lPer_Juros_5), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 6, lPer_Juros_6), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 2) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(1, 6, lPer_Juros_6), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 7, lPer_Juros_7), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 2) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(1, 7, lPer_Juros_7), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 8, lPer_Juros_8), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 2) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(1, 8, lPer_Juros_8), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 9, lPer_Juros_9), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 2) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(1, 9, lPer_Juros_9), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 10, lPer_Juros_10), 2)) + _
                                                pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 2) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                 mfCoeJuros(1, 10, lPer_Juros_10), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 11, lPer_Juros_11), 2)) + _
                                                pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 2) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                 mfCoeJuros(1, 10, lPer_Juros_11), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 1) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(1, 12, lPer_Juros_12), 2)) + _
                                                pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 2) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                 mfCoeJuros(1, 12, lPer_Juros_12), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(1, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 1, lPer_Juros_0_1), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 2, lPer_Juros_0_2), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 4) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(0, 2, lPer_Juros_0_2), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 3, lPer_Juros_0_3), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 4) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(0, 3, lPer_Juros_0_3), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 4, lPer_Juros_0_4), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 4) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(0, 4, lPer_Juros_0_4), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 5, lPer_Juros_0_5), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 4) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(0, 5, lPer_Juros_0_5), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 6, lPer_Juros_0_6), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 4) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(0, 6, lPer_Juros_0_6), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 7, lPer_Juros_0_7), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 4) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(0, 7, lPer_Juros_0_7), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 8, lPer_Juros_0_8), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 4) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(0, 8, lPer_Juros_0_8), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 9, lPer_Juros_0_9), 2)) + _
                                               pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 4) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                mfCoeJuros(0, 9, lPer_Juros_0_9), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 10, lPer_Juros_0_10), 2)) + _
                                                pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 4) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                 mfCoeJuros(0, 10, lPer_Juros_0_10), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 11, lPer_Juros_0_11), 2)) + _
                                                pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 4) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                 mfCoeJuros(0, 11, lPer_Juros_0_11), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 3) = M46V999.gfTrunca(((M46V999.gfTrunca(lValor * mfCoeJuros(0, 12, lPer_Juros_0_12), 2)) + _
                                                pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 4) = M46V999.gfTrunca((M46V999.gfTrunca(lValor * _
                                                                 mfCoeJuros(0, 12, lPer_Juros_0_12), 2)) * (1 + pProJurIOF / 100), 2)
        End If
        '===>>> IOF.
        If Not IsMissing(pIOF) Then
            pIOF = lIOF
        End If


        '===>>> IOF.
        If lParcVista = 0 And lParcs30dia = 0 Then
            lParcVista = 12
            lParcs30dia = 12
        End If

        If pParcela(1, 1) < lValorMin Then
            If pParcela(1, 1) < lPremioMin Then
                pParcela(1, 1) = 0
            End If
        End If
        For lI1 = 2 To 12
            If pParcela(lI1, 2) < lValorMin Or lParcs30dia < lI1 Then
                pParcela(lI1, 2) = 0
                pParcela(lI1, 1) = 0
            End If
            If pParcela(lI1, 4) < lValorMin Or lParcs30dia < lI1 Then
                pParcela(lI1, 4) = 0
                pParcela(lI1, 3) = 0
            End If
        Next lI1
    End If
    gfOldCalFracionamento = True
End Function
Public Function gfCalcData(pData As Long, pDias As Integer) As Long
    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "gfCalcData"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(M46V999.tagDynaTrace, methodName, Class_Name)
    
    Dim lAuxPreData As Date
    If pData <> 0 Then

        If Right(pData, 4) = "0229" Then
            pData = Left(pData, 4) & "0228"
        End If

        If Mid(pData, 7, 2) & "/" & Mid(pData, 5, 2) & "/" & Mid(pData, 1, 4) <> Format(Mid(pData, 7, 2) & "/" & Mid(pData, 5, 2) & "/" & Mid(pData, 1, 4), "dd/mm/yyyy") Then
            lAuxPreData = CDate(Mid(pData, 5, 2) & "/" & Mid(pData, 7, 2) & "/" & Mid(pData, 1, 4))
        Else
            lAuxPreData = CDate(Mid(pData, 7, 2) & "/" & Mid(pData, 5, 2) & "/" & Mid(pData, 1, 4))
        End If

        gfCalcData = Format(CDate(lAuxPreData) + pDias, "yyyymmdd")
    End If

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

End Function
Public Sub mpObtemMaxParcelas(ByVal pMaxParcPerm As Integer, _
                              ByRef lParcsVista As Integer, _
                              ByRef lParcs30dia As Integer, _
                              ByRef pIniVigencia As String, _
                              ByRef pTerVigencia As String, _
                              Optional lUltParcela As Double)

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mpObtemMaxParcelas"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(M46V999.tagDynaTrace, methodName, Class_Name)

    On Error GoTo TrataErro
    
'Determina quantide máxima de parcelas
'Parametros de entrada/saida      lParcsVista ... parcelas com pagamento antecipado
'                                 lParcs30dia ... parcelas com pgto Postecipado
'                                 lParcsCartao ... parcelas com pgto Cartão de Crédito

    Dim lAchouNParc As Boolean
    Dim lAuxData As Long
    Dim lTotalDIAS As Integer
    Dim lDataVen As Variant    'Retorna todas as datas.
    Dim lDataUltParc As Double    'Retorna data da ultima parcela possivel
    Dim lIrregular As Boolean    'Indica se há data irregula.
    Dim lHoje As Long    'Dia de hj.
    Dim lDiasAnteTerVig As Integer      'Numero de dias antes do término de vigencia
    'para determinar o vencimento da ultima parcela
    Dim lretorno As String    'Retorno da DLL.
    Dim lMensagem As String    'Mensagem.
    Dim lDataFinal As Double
    '1 inicia o valor das variaveis
    lParcs30dia = 0
    lParcsVista = 0
    If pMaxParcPerm = 0 Then pMaxParcPerm = 12

    '2- Determina o numero máximo de dias antes do término de vigencia para a
    'ultima parcela.
    lDiasAnteTerVig = 30

    lHoje = Format(Date, "yyyymmdd")

    '3- Obtem o numero de parcelas  a vista

    lDataFinal = gfCalcData(CLng(pTerVigencia), (lDiasAnteTerVig * -1))
    If Right(lDataFinal, 4) = "0229" Then
        lDataFinal = Left(lDataFinal, 4) & "0228"
    End If
    If CLng(pIniVigencia) <> 0 And CLng(pTerVigencia) <> 0 Then
        If lHoje < pIniVigencia Then
            'Inicio de vigencia é a referencia p/ a qtd de parcelas
            lTotalDIAS = val(M46V999.gfDifDatas(pIniVigencia, lDataFinal)) - 10
        Else
            'Inicio de vigencia menor que a data atual
            'Data atual é referencia p/ a qtd de parcelas
            lTotalDIAS = val(M46V999.gfDifDatas(lHoje, lDataFinal)) - 10

        End If
        If lHoje > lDataFinal Then
            lTotalDIAS = 1
        End If
    End If

    If lTotalDIAS > 30 Then
        lParcsVista = lTotalDIAS / 30
    Else
        If lTotalDIAS > 0 Then
            lParcsVista = 1
        Else
            lParcsVista = 0
        End If
    End If

    If lParcsVista > 0 Then
        lAchouNParc = False
        If lTotalDIAS > 30 Then
            lParcsVista = Int(lParcsVista) + 1
        End If
    Else
        ' Vai retornar 0
        lAchouNParc = True
    End If

    '4- Obtem maximo de parcelas com a 1ª em 30 Dias
    lDataFinal = gfCalcData(CLng(pTerVigencia), (lDiasAnteTerVig * -1))

    If Right(lDataFinal, 4) = "0229" Then
        lDataFinal = Left(lDataFinal, 4) & "0228"
    End If
    If CLng(pIniVigencia) <> 0 And CLng(pTerVigencia) <> 0 Then
        If lHoje < CLng(pIniVigencia) Then
            'Inicio de vigencia é a referencia p/ a qtd de parcelas
            lTotalDIAS = val(M46V999.gfDifDatas(pIniVigencia, lDataFinal)) - 10
        Else
            'Inicio de vigencia menor que a data atual
            'Data atual é referencia p/ a qtd de parcelas
            lTotalDIAS = val(M46V999.gfDifDatas(lHoje, lDataFinal)) - 10
        End If
        If lHoje > lDataFinal Then
            lTotalDIAS = 1
        End If
    End If

    If lTotalDIAS > 30 Then
        lParcs30dia = lTotalDIAS / 30
    Else
        If lTotalDIAS > 30 Then
            lParcs30dia = 1
        Else
            lParcs30dia = 0
        End If
    End If

    If lParcs30dia > 0 Then
        lAchouNParc = False
        lParcs30dia = Int(lParcs30dia)
    Else
        ' Vai retornar 0
        lAchouNParc = True
    End If

    '5- Limita as quantidades de parcela ao limite máximo permitido
    If lParcsVista > pMaxParcPerm Then
        lParcsVista = pMaxParcPerm
    End If

    If lParcs30dia > pMaxParcPerm Then
        lParcs30dia = pMaxParcPerm
    End If


    '7- Ajusta numero de parcelas para M99V001.Perfil 1
    If M46V999.PerfilAcesso = e_Perfil_DIG_Gestor Then
        'lParcsVista = pMaxParcPerm         'Comentada devida a ficha 0600403
        'lParcs30dia = pMaxParcPerm         'Comentada devida a ficha 0600403
        'Devida a ficha nº 0600403 deverá ser consistido para o usuário digitação gestor
        'que o número de parcelas deverá ser a quantidade permitida e não a máxima do produto
        If lParcsVista = 0 Then
            lParcsVista = 1
        End If
        If lParcs30dia = 0 Then
            lParcs30dia = 1
        End If
    End If

    '8- Grava na variavel opcional de retorno a data limite da ultima parcela
    If CLng(pTerVigencia) <> 0 Then
        lUltParcela = lDataFinal
    End If

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Sub
    
TrataErro:
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, Err.Description, 2, Class_Name, methodName)
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Sub


Private Function mfCalcDataParc(pData As String, pDias As Integer) As Long
    Dim lAuxPreData As Date
    If CDate(pData) <> 0 Then

        If Left(CDate(pData), 2) & Mid(CDate(pData), 4, 2) = "0229" Then
            pData = Right(CDate(pData), 4) & "0228"
        End If

        lAuxPreData = CDate(pData)

        mfCalcDataParc = Format(CDate(lAuxPreData) + pDias, "yyyymmdd")
    End If

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

Private Function mfValidaCongenerePortoPostecipadoSemJuros(ByRef pStcPed As stcA46V702B, _
                                                           ByVal pTip_Emissao As Long, _
                                                           ByRef pMensagem As String) As Boolean
    
    Dim lConn As ADODB.Connection
    Dim lRsQuery As clsYasRecordSet
    Dim lSelect As String
    
    mfValidaCongenerePortoPostecipadoSemJuros = False
    On Error GoTo Erro
    
    If M46V999.gfAbrBasDados("RAMOSDIVERSOS", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, lConn, pMensagem) = False Then
        GoTo Erro
    End If
    
    lSelect = "SELECT Cod_Tarifa, Cod_Corretor From tbm_juros_congenere (nolock) Where Cod_Tarifa = " & pStcPed.Cod_Produto
    lSelect = lSelect & " AND Cod_Corretor = " & pStcPed.objstcCtrlEmis.Cod_Corr
    
    If M46V999.gfObtRegistro(lConn, lSelect, lRsQuery, pMensagem) = False Then
        GoTo Erro
    End If
    
    If lRsQuery.RecordCount > 0 Then
        'Renovações Porto Seguro
        If pStcPed.Cod_Produto >= 1138 And pTip_Emissao = 101 And pStcPed.Cod_Ramo <> e_CodRamo_Condominio And _
           (pStcPed.Cod_Estip = 10 Or pStcPed.Cod_Estip = 11 Or pStcPed.Cod_Estip = 20) And _
           pStcPed.ColecaoPedLoc(1).Ren_Con_Seguradora = 588 Then
    
            mfValidaCongenerePortoPostecipadoSemJuros = True
        End If
    End If
    
    GoTo Fim
Erro:
    If Err > 0 Then
        pMensagem = pMensagem & "Erro: " & Err.Description
    End If
    
Fim:
    Set lRsQuery = Nothing
    
    If Not lConn Is Nothing Then
        lConn.Close
    End If
    Set lConn = Nothing
End Function

Private Sub mfAtribuiCongenerePortoPostecipadoSemJuros(ByRef pStcPed As stcA46V702B, _
                                                       ByVal pTip_Emissao As Long, _
                                                       ByRef pParcela As Variant, _
                                                       ByRef ZerarParcela30dias As Boolean, _
                                                       ByRef pMensagem As String)
    
    Dim lConn As ADODB.Connection
    Dim lRsQuery As clsYasRecordSet
    Dim lSelect As String
    Dim lI1 As Integer
    Dim lI3 As Integer
    
    On Error GoTo Erro
    
    If mfValidaCongenerePortoPostecipadoSemJuros(pStcPed, pTip_Emissao, pMensagem) Then
            For lI3 = 1 To M46V999.gQtdMaxComp
                For lI1 = 5 To 12
                    pParcela(lI1, 4, lI3) = 0
                    pParcela(lI1, 3, lI3) = 0
                Next
            Next
            ZerarParcela30dias = False
        End If
    
    GoTo Fim
Erro:
    If Err > 0 Then
        pMensagem = pMensagem & "Erro: " & Err.Description
    End If
    
Fim:
    Set lRsQuery = Nothing
    
    If Not lConn Is Nothing Then
        lConn.Close
    End If
    Set lConn = Nothing
End Sub

Private Function mfAtribuiUmaMaisNoveSemJurosCorretoraAthina(pStcPed As stcA46V702B, _
                                                             pMensagem As String) As Boolean
    Dim lComissao As Double
    On Error GoTo Erro
    mfAtribuiUmaMaisNoveSemJurosCorretoraAthina = False
    
    lComissao = (pStcPed.objstcCtrlEmis.Per_Corr + _
                 pStcPed.objstcCtrlEmis.Per_Corr_2 + _
                 pStcPed.objstcCtrlEmis.Per_Corr_3)
                 
    If M46V999.gCodProduto < 1172 Then Exit Function
    If pStcPed.Cod_Estip = 30 Then Exit Function 'Supremo
    If pStcPed.ColecaoPedLoc(1).Cod_Clasf = "914" Or pStcPed.ColecaoPedLoc(1).Cod_Clasf = "922" Then Exit Function
   
    If lComissao > 35 Then Exit Function

    If (M46V999.gCod_Corretor = 907677 Or M46V999.gCod_Corretor = 929745 Or M46V999.gCod_Corretor = 931081 Or M46V999.gCod_Corretor = 932688 Or M46V999.gCod_Corretor = 932732 Or M46V999.gCod_Corretor = 932733 Or M46V999.gCod_Corretor = 930550) Then
        mfAtribuiUmaMaisNoveSemJurosCorretoraAthina = True
        Exit Function
    End If

Erro:
    If Err > 0 Then
        pMensagem = pMensagem & "Erro: " & Err.Description
    End If

End Function

Private Function mfValidaParcelamentoSemJurosCartaoCredito(ByRef pStcPed As stcA46V702B, _
                                                           ByRef pMensagem As String) As Integer
    
    Dim lConn As ADODB.Connection
    Dim lRsQuery As clsYasRecordSet
    Dim lSelect As String
    
    mfValidaParcelamentoSemJurosCartaoCredito = False
    On Error GoTo Erro
    
    If M46V999.gfAbrBasDados("RAMOSDIVERSOS", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, lConn, pMensagem) = False Then
        GoTo Erro
    End If
    
    lSelect = "SELECT DISTINCT MAX(NUM_PARCELA) AS QTDE_PARC From tbm_juros Where Cod_Tarifa = " & pStcPed.Cod_Produto
    lSelect = lSelect & " AND COD_RAMO = " & pStcPed.Cod_Ramo & " AND (COD_MODALIDADE = " & pStcPed.Cod_Estip & " OR COD_MODALIDADE = 0) AND Tip_Compensacao = 4 "
    
    If M46V999.gfObtRegistro(lConn, lSelect, lRsQuery, pMensagem) = False Then
        GoTo Erro
    End If
    
    If lRsQuery.RecordCount > 0 Then
        If IsNull(lRsQuery("QTDE_PARC")) Then
            mfValidaParcelamentoSemJurosCartaoCredito = 0
        Else
            mfValidaParcelamentoSemJurosCartaoCredito = lRsQuery("QTDE_PARC")
        End If
    Else
        mfValidaParcelamentoSemJurosCartaoCredito = 0
    End If
    
    GoTo Fim
Erro:
    If Err > 0 Then
        pMensagem = pMensagem & "Erro: " & Err.Description
    End If
    
Fim:
    Set lRsQuery = Nothing
    
    If Not lConn Is Nothing Then
        lConn.Close
    End If
    Set lConn = Nothing
End Function

Private Sub mfAtribuiParcelamentoCartaoCredito(ByRef pParcela As Variant, _
                                               ByRef QtdeParcCartao As Integer, _
                                               ByVal ZerarParc30dias As Boolean, _
                                               ByRef pMensagem As String)
    
    Dim lI1 As Integer
    Dim lI3 As Integer
    
    On Error GoTo Erro
    
    lI3 = 4
    
    If QtdeParcCartao > 0 Then
        For lI1 = (QtdeParcCartao + 1) To 12
            pParcela(lI1, 1, lI3) = 0
            pParcela(lI1, 2, lI3) = 0
        Next
        
        For lI1 = 1 To 12
            pParcela(lI1, 4, lI3) = 0
            pParcela(lI1, 3, lI3) = 0
        Next
    Else
        For lI1 = 1 To 12
            pParcela(lI1, 1, lI3) = 0
            pParcela(lI1, 2, lI3) = 0
            pParcela(lI1, 4, lI3) = 0
            pParcela(lI1, 3, lI3) = 0
        Next
    End If
    
Erro:
    If Err > 0 Then
        pMensagem = pMensagem & "Erro: " & Err.Description
    End If
    
End Sub
