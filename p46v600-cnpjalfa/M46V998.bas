Attribute VB_Name = "basM46V998"
'MÓDULO COM FUNÇÕES E PROCEDURES COMUNS DO SYAS QUE UTILIZAM DLL.

'AO ALTERAR ESTE MÓDULO, COMPILAR E CRIAR EXECUTÁVEL DE TODOS OS MÓDULOS QUE O REFERENCIAM.

Option Explicit
Private mobjstcPed As stcA46V702B
Public gRst_RegrasPlano         As ADODB.Recordset
'Ficha 1000495 - Plano Belacity
Public gPlanoImobiliaria        As Boolean
'Ficha 1000601
Public gLiberaDescontoEspecial  As Boolean  'Liberar item Desconto Especial na tela de Proteção

Private Function mfCoeJuros(ByVal pTipPagamento As Byte, ByVal pQuaParcelas As Single, _
                            ByVal pJuros As Single) As Single
    'Função #007: calcula coeficiente de juros.

    'Parâmetros de entrada...pTipPagamento...0 = pagamento postecipado e 1 = pagamento antecipado.
    '                        pQuaParcelas....quantidade de parcelas.
    '                        pJuros..........juros.

    Dim lDesPagamento   As String   'Descrição do pagamento.

    If pJuros = 0 Then
        If pQuaParcelas = 0 Then
            mfCoeJuros = 1
        Else
            mfCoeJuros = 1 / pQuaParcelas
        End If
    Else
        Select Case pTipPagamento
            Case 0  'Pagamento postecipado.
                mfCoeJuros = gfTrunca(((pJuros / 100) / (1 - (1 / (1 + (pJuros / 100)) ^ _
                                        pQuaParcelas))), 4)
            Case 1  'Pagamento antecipado.
                mfCoeJuros = gfTrunca(((pJuros / 100) / (1 - (1 / (1 + (pJuros / 100)) ^ _
                                        pQuaParcelas))) * (1 / (1 + (pJuros / 100))), 4)
        End Select
    End If

    If pTipPagamento = 0 Then
        lDesPagamento = "0 (pagamento postecipado)"
    Else
        lDesPagamento = "1 (pagamento antecipado)"
    End If
    
End Function

Public Function gfCalFracionamento(ByRef pStcPed As stcA46V702B, _
                                   ByVal pTip_Emissao As Long, _
                                   ByRef pParcela As Variant, ByRef pMensagem As String, _
                                   ByVal pProJurIOF As Single, Optional ByVal pCodReg As Integer, _
                                   Optional ByRef pIOF As Single, Optional ByVal pParcMin As Single)
    
    '1100303 - Arredondamento no Parcelamento - 'Foi realizado troca de algumas variáveis de Double para Single e tambem substituido a função gfTrunca() pela Round() - caso queira comparar historico de alteração no SouceSafe
        
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
    Dim lI1             As Byte         'Utilizado com For...Next.
    Dim lI2             As Byte         'Utilizado com For...Next.
    Dim lValorMin       As Single       'Valor mínimo da parcela ou endosso.
    Dim lPremioMin      As Single       'Valor mínimo do prêmio de parcela ou endosso.
    Dim lParcVista      As Integer      'Quantidade de parcelas máxima possível em caso de pagamento antecipado.
    Dim lParcs30dia     As Integer      'Quandidade de parcelas máxima possível em caso de paramento não antecipado.
    Dim lPreTotal       As Single       'Prêmio total.
    Dim lValor          As Single       'Prêmio Liquido.
    Dim lIOF            As Single       'Valor IOF.
    Dim lPer_Juros_1    As Double       'Juros da 1ª parcela.
    Dim lPer_Juros_2    As Double       'Juros da 2ª parcela.
    Dim lPer_Juros_3    As Double       'Juros da 3ª parcela.
    Dim lPer_Juros_4    As Double       'Juros da 4ª parcela.
    Dim lPer_Juros_5    As Double       'Juros da 5ª parcela.
    Dim lPer_Juros_6    As Double       'Juros da 6ª parcela.
    Dim lPer_Juros_7    As Double       'Juros da 7ª parcela.
    Dim lPer_Juros_8    As Double       'Juros da 8ª parcela.
    Dim lPer_Juros_9    As Double       'Juros da 9ª parcela.
    Dim lPer_Juros_10   As Double       'Juros da 10ª parcela.
    Dim lPer_Juros_11   As Double       'Juros da 11ª parcela.
    Dim lPer_Juros_12   As Double       'Juros da 12ª parcela.
    Dim lPer_Juros_0_1  As Double       'Juros da 1ª parcela antecipada.
    Dim lPer_Juros_0_2  As Double       'Juros da 2ª parcela antecipada.
    Dim lPer_Juros_0_3  As Double       'Juros da 3ª parcela antecipada.
    Dim lPer_Juros_0_4  As Double       'Juros da 4ª parcela antecipada.
    Dim lPer_Juros_0_5  As Double       'Juros da 5ª parcela antecipada.
    Dim lPer_Juros_0_6  As Double       'Juros da 6ª parcela antecipada.
    Dim lPer_Juros_0_7  As Double       'Juros da 7ª parcela antecipada.
    Dim lPer_Juros_0_8  As Double       'Juros da 8ª parcela antecipada.
    Dim lPer_Juros_0_9  As Double       'Juros da 9ª parcela antecipada.
    Dim lPer_Juros_0_10 As Double       'Juros da 10ª parcela antecipada.
    Dim lPer_Juros_0_11 As Double       'Juros da 11ª parcela antecipada.
    Dim lPer_Juros_0_12 As Double       'Juros da 12ª parcela antecipada.
    Dim lobjstcJuros    As stcA46V711B
    Dim lTemAjusteVenc  As Boolean
    Dim lSelect         As String
    
Dim lIniVigencia As String
Dim lTerVigencia As String
Dim lCusEmissao As Single
Dim lCusParcelado As String
Dim lPreLiquido As Single
Dim lCod_Iden_Tarf As Long
    
Dim objStcParcela As stcParcelamento
    
On Error GoTo TrataErro
    
    gfCalFracionamento = False
    
    'Inicio os valores
    
    lIniVigencia = gfForData(2, pStcPed.Dat_Ini_Vig)
    lTerVigencia = gfForData(2, pStcPed.Dat_Ter_Vig)
    lCusEmissao = pStcPed.Val_Cust_Emis
    lCusParcelado = IIf(pStcPed.gCustoParcelado = True, "S", "N")
    lPreLiquido = pStcPed.VAL_PR
    lCod_Iden_Tarf = pStcPed.Cod_Produto
    
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
    If pCodReg = 2 Or _
       pCodReg = 3 Or _
       pCodReg = 102 Or _
       pCodReg = 103 Or _
       pCodReg = 202 Or _
       pCodReg = 203 Or _
       pStcPed.Qtd_Parc = 0 Then
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
        If gCodProduto >= 976 And lValor < gProPremioMinimo Then
            lValor = gProPremioMinimo 'prêmio mínimo
        End If
        
        lIOF = Round((lValor + lCusEmissao) * (pProJurIOF / 100), 2)
        lPreTotal = lValor + lCusEmissao + lIOF
        
        lValorMin = pParcMin

        If lValor < lPremioMin Then
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
    
        Dim lP00Multi_Tplano  As ADODB.Connection
        Dim lrstRegraPlano  As ADODB.Recordset
        Dim lMensagem       As String
        
        If gfAbrBasDados("P00MULTI2", gTipP00MULTI, gEndP00MULTI, lP00Multi_Tplano, lMensagem) = False Then Call gpGraLog(2, lMensagem): GoTo SAIDA
        lSelect = "SELECT COD_REGRA,VAL_REGRA FROM TMB_REGRA_PLANO "
        lSelect = lSelect & " Where COD_RAMO = " & pStcPed.Cod_Ramo
        lSelect = lSelect & " AND COD_MODALIDADE = " & pStcPed.Cod_Estip
        lSelect = lSelect & " AND COD_SEQUENCIA <=" & pStcPed.Cod_Sequencia
        lSelect = lSelect & " AND COD_TARIF=" & pStcPed.Cod_Produto
        lSelect = lSelect & " Order By Cod_Regra"
        If gfObtRegistro(lP00Multi_Tplano, lSelect, lrstRegraPlano, lMensagem) = False Then Call gpGraLog(2, lMensagem): GoTo SAIDA
    
        'forma de pagamento de acordo com o plano se estiver cadastrado
        If lrstRegraPlano.RecordCount > 0 Then
            lI2 = 1
            For lI1 = 20 To 40 Step 2
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
        Call gpFecha3(lrstRegraPlano, lP00Multi_Tplano)
        
        'Correta e National
'        If (gCod_Plano >= 2001 And gCod_Plano <= 2010) Or (gCod_Plano >= 2032 And gCod_Plano <= 2041) Then
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
'        If gCod_Plano >= 2011 And gCod_Plano <= 2020 Then
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
'        If gCod_Plano >= 2021 And gCod_Plano <= 2031 Then
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
'        If gCod_Plano >= 2051 And gCod_Plano <= 2060 Then
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
'        If (gCod_Plano >= 2061 And gCod_Plano <= 2100) Or (gCod_Plano >= 2111 And gCod_Plano <= 2130) Then
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
'        If (gCod_Plano >= 2001 And gCod_Plano <= 2010) Or (gCod_Plano >= 2032 And gCod_Plano <= 2041) Then
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
'        If gCod_Plano >= 2011 And gCod_Plano <= 2020 Then
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
'        If gCod_Plano >= 2021 And gCod_Plano <= 2031 Then
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
'                lIniVigencia = gfForData(1, lIniVigencia)
'            End If
'            If IsDate(lTerVigencia) Then
'                lTerVigencia = gfForData(1, lTerVigencia)
'            End If
'            Call mpObtemMaxParcelas(gProQuaMaxParcelas, lParcVista, lParcs30dia, lIniVigencia, lTerVigencia, lUltParcela)
'
'            If gfPreenchido(gProDatUltParcela) = True Then
'               If gProDatUltParcela > 0 Then
'                  If gProDatUltParcela > lTerVigencia Then
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
'        If gCod_Plano >= 2051 And gCod_Plano <= 2060 Then
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
'        If (gCod_Plano >= 2061 And gCod_Plano <= 2100) Or (gCod_Plano >= 2111 And gCod_Plano <= 2130) Then
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
'        If gCod_Plano >= 2042 And gCod_Plano <= 2049 Then
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
        
        If (lParcVista = 0 And lParcs30dia = 0) Or (gTipoAcesso = DIGITACAO And lTemAjusteVenc = True) Then
            lParcVista = 12
            lParcs30dia = 12
        End If
        
        If pParcela(1, 1) < lValorMin Then
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
        
    End If
        
    If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
        '022. Forma de pagamento: primeira parcela.
        pStcPed.Val_PriParcela = pParcela(pStcPed.Qtd_Parc, 1)
        '022. Forma de pagamento: demais parcelas.
        pStcPed.val_DemParcelas = pParcela(pStcPed.Qtd_Parc, 2)
    Else
        '023. Forma de pagamento: primeira parcela.
        pStcPed.Val_PriParcela = pParcela(pStcPed.Qtd_Parc, 3)
        '023. Forma de pagamento: demais parcelas.
        pStcPed.val_DemParcelas = pParcela(pStcPed.Qtd_Parc, 4)
    End If
        
    For lI1 = 1 To pStcPed.getProduto.ProQuaMaxParcelas
'        If pParcela(lI1, 1) = 0 Then Exit For
        Set objStcParcela = New stcParcelamento
        objStcParcela.Nosso_Numero = pStcPed.Nosso_Numero
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
    
    gfCalFracionamento = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If gfPreenchido(pMensagem) Then
        pMensagem = Err & " - " & Error
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
    Dim lI1             As Byte         'Utilizado com For...Next.
    Dim lI2             As Byte         'Utilizado com For...Next.
    Dim lValorMin       As Double       'Valor mínimo da parcela ou endosso.
    Dim lPremioMin      As Double       'Valor mínimo do prêmio de parcela ou endosso.
    Dim lParcVista      As Integer      'Quantidade de parcelas máxima possível em caso de pagamento antecipado.
    Dim lParcs30dia     As Integer      'Quandidade de parcelas máxima possível em caso de paramento não antecipado.
    Dim lPreTotal       As Double       'Prêmio total.
    Dim lDivCust        As Integer      'Número pelo qual será dividido o custo
    Dim lValor          As Double       'Prêmio Liquido.
    Dim lIOF            As Double       'Valor IOF.
    Dim lPer_Juros_1    As Double       'Juros da 1ª parcela.
    Dim lPer_Juros_2    As Double       'Juros da 2ª parcela.
    Dim lPer_Juros_3    As Double       'Juros da 3ª parcela.
    Dim lPer_Juros_4    As Double       'Juros da 4ª parcela.
    Dim lPer_Juros_5    As Double       'Juros da 5ª parcela.
    Dim lPer_Juros_6    As Double       'Juros da 6ª parcela.
    Dim lPer_Juros_7    As Double       'Juros da 7ª parcela.
    Dim lPer_Juros_8    As Double       'Juros da 8ª parcela.
    Dim lPer_Juros_9    As Double       'Juros da 9ª parcela.
    Dim lPer_Juros_10   As Double       'Juros da 10ª parcela.
    Dim lPer_Juros_11   As Double       'Juros da 11ª parcela.
    Dim lPer_Juros_12   As Double       'Juros da 12ª parcela.
    Dim lPer_Juros_0_1  As Double       'Juros da 1ª parcela antecipada.
    Dim lPer_Juros_0_2  As Double       'Juros da 2ª parcela antecipada.
    Dim lPer_Juros_0_3  As Double       'Juros da 3ª parcela antecipada.
    Dim lPer_Juros_0_4  As Double       'Juros da 4ª parcela antecipada.
    Dim lPer_Juros_0_5  As Double       'Juros da 5ª parcela antecipada.
    Dim lPer_Juros_0_6  As Double       'Juros da 6ª parcela antecipada.
    Dim lPer_Juros_0_7  As Double       'Juros da 7ª parcela antecipada.
    Dim lPer_Juros_0_8  As Double       'Juros da 8ª parcela antecipada.
    Dim lPer_Juros_0_9  As Double       'Juros da 9ª parcela antecipada.
    Dim lPer_Juros_0_10 As Double       'Juros da 10ª parcela antecipada.
    Dim lPer_Juros_0_11 As Double       'Juros da 11ª parcela antecipada.
    Dim lPer_Juros_0_12 As Double       'Juros da 12ª parcela antecipada.
    Dim lobjstcJuros    As stcA46V711B
    
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
            pParcela(1, 1) = gfTrunca((lValor + pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 2, lPer_Juros_2), 2)) + _
                                       (pCusEmissao / 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 2) = pParcela(2, 1)
            pParcela(3, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 3, lPer_Juros_3), 2)) + _
                                       (pCusEmissao / 3)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 2) = pParcela(3, 1)
            pParcela(4, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 4, lPer_Juros_4), 2)) + _
                                       (pCusEmissao / 4)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 2) = pParcela(4, 1)
            pParcela(5, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 5, lPer_Juros_5), 2)) + _
                                       (pCusEmissao / 5)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 2) = pParcela(5, 1)
            pParcela(6, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 6, lPer_Juros_6), 2)) + _
                                       (pCusEmissao / 6)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 2) = pParcela(6, 1)
            pParcela(7, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 7, lPer_Juros_7), 2)) + _
                                       (pCusEmissao / 7)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 2) = pParcela(7, 1)
            pParcela(8, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 8, lPer_Juros_8), 2)) + _
                                       (pCusEmissao / 8)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 2) = pParcela(8, 1)
            pParcela(9, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 9, lPer_Juros_9), 2)) + _
                                       (pCusEmissao / 9)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 2) = pParcela(9, 1)
            pParcela(10, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 10, lPer_Juros_10), 2)) + _
                                        (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 2) = pParcela(10, 1)
            pParcela(11, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 11, lPer_Juros_11), 2)) + _
                                        (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 2) = pParcela(11, 1)
            pParcela(12, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 12, lPer_Juros_12), 2)) + _
                                        (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 2) = pParcela(12, 1)
            pParcela(1, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 1, lPer_Juros_0_1), 2)) + _
                                       (pCusEmissao / 1)) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 2, lPer_Juros_0_2), 2)) + _
                                       (pCusEmissao / 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 4) = pParcela(2, 3)
            pParcela(3, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 3, lPer_Juros_0_3), 2)) + _
                                       (pCusEmissao / 3)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 4) = pParcela(3, 3)
            pParcela(4, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 4, lPer_Juros_0_4), 2)) + _
                                       (pCusEmissao / 4)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 4) = pParcela(4, 3)
            pParcela(5, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 5, lPer_Juros_0_5), 2)) + _
                                       (pCusEmissao / 5)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 4) = pParcela(5, 3)
            pParcela(6, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 6, lPer_Juros_0_6), 2)) + _
                                       (pCusEmissao / 6)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 4) = pParcela(6, 3)
            pParcela(7, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 7, lPer_Juros_0_7), 2)) + _
                                       (pCusEmissao / 7)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 4) = pParcela(7, 3)
            pParcela(8, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 8, lPer_Juros_0_8), 2)) + _
                                       (pCusEmissao / 8)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 4) = pParcela(8, 3)
            pParcela(9, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 9, lPer_Juros_0_9), 2)) + _
                                       (pCusEmissao / 9)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 4) = pParcela(9, 3)
            pParcela(10, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 10, lPer_Juros_0_10), 2)) + _
                                        (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 4) = pParcela(10, 3)
            pParcela(11, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 11, lPer_Juros_0_11), 2)) + _
                                        (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 4) = pParcela(11, 3)
            pParcela(12, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 12, lPer_Juros_0_12), 2)) + _
                                        (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 4) = pParcela(12, 3)
        Else
            'b) Considerando custo de emissão NÃO parcelado.
            pParcela(1, 1) = gfTrunca((lValor + pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 2, lPer_Juros_2), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 2) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(1, 2, lPer_Juros_2), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 3, lPer_Juros_3), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 2) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(1, 3, lPer_Juros_3), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 4, lPer_Juros_4), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 2) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(1, 4, lPer_Juros_4), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 5, lPer_Juros_5), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 2) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(1, 5, lPer_Juros_5), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 6, lPer_Juros_6), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 2) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(1, 6, lPer_Juros_6), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 7, lPer_Juros_7), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 2) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(1, 7, lPer_Juros_7), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 8, lPer_Juros_8), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 2) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(1, 8, lPer_Juros_8), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 9, lPer_Juros_9), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 2) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(1, 9, lPer_Juros_9), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 10, lPer_Juros_10), 2)) + _
                                         pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 2) = gfTrunca((gfTrunca(lValor * _
                                               mfCoeJuros(1, 10, lPer_Juros_10), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 11, lPer_Juros_11), 2)) + _
                                         pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 2) = gfTrunca((gfTrunca(lValor * _
                                               mfCoeJuros(1, 10, lPer_Juros_11), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 1) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(1, 12, lPer_Juros_12), 2)) + _
                                         pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 2) = gfTrunca((gfTrunca(lValor * _
                                               mfCoeJuros(1, 12, lPer_Juros_12), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(1, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 1, lPer_Juros_0_1), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 2, lPer_Juros_0_2), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 4) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(0, 2, lPer_Juros_0_2), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 3, lPer_Juros_0_3), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 4) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(0, 3, lPer_Juros_0_3), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 4, lPer_Juros_0_4), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 4) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(0, 4, lPer_Juros_0_4), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 5, lPer_Juros_0_5), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 4) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(0, 5, lPer_Juros_0_5), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 6, lPer_Juros_0_6), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 4) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(0, 6, lPer_Juros_0_6), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 7, lPer_Juros_0_7), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 4) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(0, 7, lPer_Juros_0_7), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 8, lPer_Juros_0_8), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 4) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(0, 8, lPer_Juros_0_8), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 9, lPer_Juros_0_9), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 4) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(0, 9, lPer_Juros_0_9), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 10, lPer_Juros_0_10), 2)) + _
                                         pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 4) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(0, 10, lPer_Juros_0_10), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 11, lPer_Juros_0_11), 2)) + _
                                         pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 4) = gfTrunca((gfTrunca(lValor * _
                                              mfCoeJuros(0, 11, lPer_Juros_0_11), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 3) = gfTrunca(((gfTrunca(lValor * mfCoeJuros(0, 12, lPer_Juros_0_12), 2)) + _
                                         pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 4) = gfTrunca((gfTrunca(lValor * _
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
    
End Function
Public Sub mpObtemMaxParcelas(ByVal pMaxParcPerm As Integer, _
                              ByRef lParcsVista As Integer, _
                              ByRef lParcs30dia As Integer, _
                              ByRef pIniVigencia As String, _
                              ByRef pTerVigencia As String, _
                              Optional lUltParcela As Double)
'Determina quantide máxima de parcelas
    'Parametros de entrada/saida      lParcsVista ... parcelas com pagamento antecipado
    '                                 lParcs30dia ... parcelas com pgto Postecipado
    
    Dim lAchouNParc         As Boolean
    Dim lAuxData            As Long
    Dim lTotalDIAS          As Integer
    Dim lDataVen            As Variant          'Retorna todas as datas.
    Dim lDataUltParc        As Double           'Retorna data da ultima parcela possivel
    Dim lIrregular          As Boolean          'Indica se há data irregula.
    Dim lHoje               As Long             'Dia de hj.
    Dim lDiasAnteTerVig     As Integer          'Numero de dias antes do término de vigencia
                                                    'para determinar o vencimento da ultima parcela
    Dim lRetorno            As String           'Retorno da DLL.
    Dim lMensagem           As String           'Mensagem.
    Dim lDataFinal          As Double
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
              lTotalDIAS = Val(gfDifDatas(pIniVigencia, lDataFinal)) - 10
        Else
              'Inicio de vigencia menor que a data atual
              'Data atual é referencia p/ a qtd de parcelas
              lTotalDIAS = Val(gfDifDatas(lHoje, lDataFinal)) - 10
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
        lParcsVista = Int(lParcsVista) + 1
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
              lTotalDIAS = Val(gfDifDatas(pIniVigencia, lDataFinal)) - 10
        Else
              'Inicio de vigencia menor que a data atual
              'Data atual é referencia p/ a qtd de parcelas
              lTotalDIAS = Val(gfDifDatas(lHoje, lDataFinal)) - 10
        End If
    End If

    If lTotalDIAS > 30 Then
          lParcs30dia = lTotalDIAS / 30
    Else
          If lTotalDIAS > 0 Then
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
   
        
'7- Ajusta numero de parcelas para perfil 1
    If gPerfilAcesso = e_Perfil_DIG_Gestor Then
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
