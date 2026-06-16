VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V732rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStcA46V732B As stcA46V732B

Public Property Set stcOrigem(ByRef pStcA46V732B As stcA46V732B)
    Set mStcA46V732B = pStcA46V732B
End Property


Public Property Get CodCobert() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format$(mStcA46V732B.Cod_Cobert, "0000")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    CodCobert = lretorno
End Property
Public Property Get nomCobert() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Left(mStcA46V732B.gNomeCobertura, 80)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    nomCobert = lretorno
End Property
Public Property Get limMaxGarantia() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Trim(Format$(mStcA46V732B.Val_IS, "standard"))
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    limMaxGarantia = lretorno
End Property

Public Property Get temFranquia() As Boolean
    Dim lretorno As Boolean
    'lretorno = (mStcA46V732B.Per_Franq <> 0 Or mStcA46V732B.Val_Franq_Min <> 0 Or mStcA46V732B.Tem_VideObs = "S")
    lretorno = ((mStcA46V732B.Num_Meses_PI <> 0 And mStcA46V732B.M46V999.gCodProduto < 1080) Or mStcA46V732B.Per_Franq <> 0 Or mStcA46V732B.Val_Franq_Min <> 0 Or mStcA46V732B.gPrazo <> 0 Or mStcA46V732B.TextoFranquia <> "")
    temFranquia = lretorno
End Property
Public Property Get codFranquia() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format$(mStcA46V732B.Cod_Cobert, "000")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    codFranquia = lretorno
End Property
Public Property Get nomFranquia() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Left(mStcA46V732B.gNomeCobertura, 40)
    '1100864 - Tarifa 1018 e 1019 Setembro 2011
    '            If (mStcA46V732B.ObjPai.Cod_Produto >= 1018) And (mStcA46V732B.Cod_Ramo = e_CodRamo_Empresarial) And _
                 '               (Val(mStcA46V732B.ObjPai.Cod_Clasf) = 282) And (Val(mStcA46V732B.ObjPai.Cod_Compl_Clasf) = 10) And _
                 '               mStcA46V732B.Cod_Cobert = 23 Then
    '                lRetorno = lRetorno & " (DE HÓSPEDES)"
    '            End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    nomFranquia = lretorno
End Property

Private Function mpGetObservacaoFranquia(ByVal pStcA46V732B As stcA46V732B) As String
    Dim lrsFranquia As clsYasRecordSet
    Dim objStcTabPedLoc As stcA46V708B
    
    Set objStcTabPedLoc = pStcA46V732B.ObjPai.ColecaoPedLoc(1)
    
    Call mStcA46V732B.M46V999.mfFranqNovoLim(pStcA46V732B.Cod_Cobert, _
                                             pStcA46V732B.Val_IS, _
                                             objStcTabPedLoc.Cod_Clasf, _
                                             objStcTabPedLoc.Cod_Compl_Clasf, _
                                             objStcTabPedLoc.Cod_Constr, _
                                             Format(objStcTabPedLoc.Num_Cep_Risco, "00000000"), _
                                             objStcTabPedLoc.Tip_Cobert, _
                                             pStcA46V732B.Num_Meses_PI, _
                                             lrsFranquia)
    If lrsFranquia.RecordCount > 0 Then
        mpGetObservacaoFranquia = mStcA46V732B.M46V999.gfTratarNulo(lrsFranquia("Observacao_Franq"), tpoString)
    End If
End Function

Public Property Get descFranquia() As String
    Dim lretorno As String
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer
    Dim lPrazo As Integer
    Dim lobjstcItem As stcA46V708B
    Dim lTextoFranquia As String
    Dim lTextoObservacao As String
    Dim lobjstcCobert As stcA46V704B
    
    'Verificação de corretude do Texto de Franquia para Tarifa igual ou superior 1080
    If mStcA46V732B.M46V999.gCodProduto >= 1080 Then

        
        For Each lobjstcItem In mStcA46V732B.ObjPai.ColecaoPedLoc
            For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
                If lobjstcCobert.Cod_Cobert = mStcA46V732B.Cod_Cobert Then
                    Exit For
                End If
            Next
            If Not lobjstcCobert Is Nothing Then
                If lobjstcCobert.Cod_Cobert = mStcA46V732B.Cod_Cobert Then
                    Exit For
                End If
            End If
        Next
        
        lTextoObservacao = mpGetObservacaoFranquia(mStcA46V732B)
        lTextoFranquia = mStcA46V732B.M46V999.gfTextoFranquia(mStcA46V732B.Val_Franq_Min, _
                                                              mStcA46V732B.Per_Franq, _
                                                              Left(lTextoObservacao, 1), _
                                                              mStcA46V732B.Cod_Cobert, _
                                                              mStcA46V732B.Num_Meses_PI, _
                                                              mStcA46V732B.Val_IS, _
                                                              lobjstcItem)
        'Retira caracteres especiais
        lTextoFranquia = Trim(mStcA46V732B.M46V999.gfRetirarCaracterEspecial(lTextoFranquia))
        descFranquia = lTextoFranquia
        Exit Property
    End If

    If mStcA46V732B.Per_Franq <> 0 Then
        '%Sobre Prejuízos
        lAuxiliar1 = mStcA46V732B.Per_Franq
        lAuxiliar3 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar3)
        If lTamanho < 10 Then
            lAuxiliar3 = Space(10 - lTamanho) & lAuxiliar3
        End If
        lAuxiliar2 = "- " & lAuxiliar3 & " % sobre prej.indeniz"
        'Limitado ao Mínimo
        lAuxiliar1 = mStcA46V732B.Val_Franq_Min
        lAuxiliar3 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar3)
        If lTamanho < 5 Then
            lAuxiliar3 = Space(5 - lTamanho) & lAuxiliar3
        End If
        If (mStcA46V732B.Val_Franq_Max > 0 Or mStcA46V732B.Val_Franq_Min > 0) Then
            lAuxiliar2 = lAuxiliar2 & ", limit ao mín de " & IIf(lAuxiliar3 = "", " ", lAuxiliar3)

            'Limitado ao Máximo
            lAuxiliar1 = mStcA46V732B.Val_Franq_Max
            lAuxiliar3 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar3)
            If lTamanho < 5 Then
                lAuxiliar3 = Space(5 - lTamanho) & lAuxiliar3
            End If
            
            If mStcA46V732B.Val_Franq_Max > 0 Then
                lAuxiliar2 = lAuxiliar2 & " e max de " & IIf(lAuxiliar3 = "", " ", lAuxiliar3)
            End If
        Else
'            If mStcA46V732B.Val_Franq_Min > 0 Then
'                lAuxiliar2 = lAuxiliar2 & ", limitado ao mínimo de " & IIf(lAuxiliar3 = "", " ", lAuxiliar3)
'            Else
                lAuxiliar2 = lAuxiliar2 & "."
'            End If
        End If
        '1100589 - Tarifa 1014 e 1015 Julho 2011
        If mStcA46V732B.ObjPai.Cod_Produto >= 1014 And mStcA46V732B.Cod_Ramo = e_CodRamo_Condominio Then
            'cobertura basica
            If mStcA46V732B.Cod_Cobert = 1 Then
                If mStcA46V732B.ObjPai.Cod_Estip = mStcA46V732B.M46V999.gfPlanoBasico(mStcA46V732B.Cod_Ramo) Then
                    'se plano simples
                    lAuxiliar2 = Format$(mStcA46V732B.Per_Franq, "standard") & " % prej.indeniz, limitado ao mínimo " & Format$(mStcA46V732B.Val_Franq_Min, "standard") & " p/ queda de raio"
                Else
                    'se plano amplo
                    lAuxiliar2 = Format$(mStcA46V732B.Per_Franq, "standard") & "% do LMG da Cobertura Básica Ampla."
                End If
            End If
        End If
        '1100705 - Tarifa 1016 e 1017 Agosto 2011
        lretorno = lAuxiliar2
        
'        For Each lobjstcItem In mStcA46V732B.ObjPai.ColecaoPedLoc
'            If (mStcA46V732B.ObjPai.Cod_Produto >= 1016) And (mStcA46V732B.Cod_Ramo = e_CodRamo_Empresarial) Then
'                If (Val(lobjstcItem.Cod_Clasf) = mStcA46V732B.M46V999.gfBuscarDePara(22, 2) And Val(lobjstcItem.Cod_Compl_Clasf) = mStcA46V732B.M46V999.gfBuscarDePara(22, 2, 13)) And _
'                   (mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(56, 1) Or mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(57, 1) Or _
'                    mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(122, 1) Or mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(123, 1) Or _
'                    mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(124, 1) Or mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(125, 1)) Then
'                    lAuxiliar2 = Format$(mStcA46V732B.Val_Franq_Min, "standard")
'                    lTamanho = Len(lAuxiliar2)
'                    If lTamanho < 10 Then
'                        lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar3
'                    End If
'                    lAuxiliar2 = "- " & lAuxiliar2
'                    lAuxiliar2 = lAuxiliar2 & " - Por Veículo (vide Observações)"
'                    Exit For
'                End If
'            End If
'        Next lobjstcItem
'        '1100864 - Tarifa 1018 e 1019 Setembro 2011
'        For Each lobjstcItem In mStcA46V732B.ObjPai.ColecaoPedLoc
'            If (mStcA46V732B.ObjPai.Cod_Produto >= 1018) And (mStcA46V732B.Cod_Ramo = e_CodRamo_Empresarial) And _
'               (Val(lobjstcItem.Cod_Clasf) = mStcA46V732B.M46V999.gfBuscarDePara(282, 2)) And (Val(lobjstcItem.Cod_Compl_Clasf) = mStcA46V732B.M46V999.gfBuscarDePara(282, 2, 10)) And _
'               (mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(23, 1) Or mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(38, 1)) Then
'                lAuxiliar2 = Format$(mStcA46V732B.Per_Franq, "standard") & "% s/ prej. indeniz, lim. mín. " & Format$(mStcA46V732B.Val_Franq_Min, "standard")
'                If (mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(38, 1)) Or (mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(23, 1) And InStr(1, "AC;AM;AP;RO;RR", lobjstcItem.Sig_UF_Risco, vbTextCompare)) Then
'                    lAuxiliar2 = lAuxiliar2 & " (Vide Observações)"
'                    Exit For
'                End If
'            End If
'        Next lobjstcItem
'        lretorno = lAuxiliar2
        '                GeraCYX_PROP2 pLinhaCXY, "franquia " & lCountLine, pPosY, 10, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        '                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ElseIf mStcA46V732B.Val_Franq_Min <> 0 Then
        'Limitado ao Mínimo
        lAuxiliar1 = mStcA46V732B.Val_Franq_Min
        lAuxiliar3 = "R$ " & Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar3)
        If lTamanho < 10 Then
            lAuxiliar3 = Space(10 - lTamanho) & lAuxiliar3
        End If
        lAuxiliar3 = "- " & lAuxiliar3
        '1100705 - Tarifa 1016 e 1017 Agosto 2011
        For Each lobjstcItem In mStcA46V732B.ObjPai.ColecaoPedLoc
            If (mStcA46V732B.ObjPai.Cod_Produto >= 1016) And (mStcA46V732B.Cod_Ramo = e_CodRamo_Empresarial) Then
                If (val(lobjstcItem.Cod_Clasf) = mStcA46V732B.M46V999.gfBuscarDePara(22, 2) And val(lobjstcItem.Cod_Compl_Clasf) = mStcA46V732B.M46V999.gfBuscarDePara(22, 2, 13)) And _
                   (mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(56, 1) Or mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(57, 1) Or _
                    mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(122, 1) Or mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(123, 1) Or _
                    mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(124, 1) Or mStcA46V732B.Cod_Cobert = mStcA46V732B.M46V999.gfBuscarDePara(125, 1)) Then
                    lAuxiliar3 = lAuxiliar3 & " - Por Veículo (vide Observações)"
                    Exit For
                End If
            End If
        Next lobjstcItem
        lretorno = lAuxiliar3
        '                GeraCYX_PROP2 pLinhaCXY, "franquia " & lCountLine, pPosY, 10, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        '                Print #pNumArqDados, IIf(lAuxiliar3 = "", " ", lAuxiliar3)
    ElseIf mStcA46V732B.Num_Meses_PI > 0 Then  'Pi indenitário
        lAuxiliar2 = "PI = " & mStcA46V732B.Num_Meses_PI & " meses"
        Call mpPrazoTabCoefPI(mStcA46V732B.Num_Meses_PI, lPrazo)
        If lPrazo <> 0 Then
            lAuxiliar2 = lAuxiliar2 & ": " & lPrazo & " dias"
        End If
        lretorno = lAuxiliar2
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    
    If Left(lretorno, 1) = "-" Then
        lretorno = Trim(Right(lretorno, Len(lretorno) - 1))
    End If
    
    lretorno = IIf(lretorno = "", " ", lretorno)
    Set lobjstcItem = Nothing
    Set lobjstcCobert = Nothing
    descFranquia = lretorno
End Property

Public Property Get descFranquiaPrazo() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V732B.gPrazo <> 0 Then
        lretorno = mStcA46V732B.gPrazo & " dias."
    Else
        lretorno = ""
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    '    lRetorno = IIf(lRetorno = "", " ", lRetorno)
    descFranquiaPrazo = lretorno
End Property

Public Property Get periodoIndenitario() As String
    Dim lretorno As String
    If mStcA46V732B.Num_Meses_PI <> 0 Then
        lretorno = Me.Num_Meses_PI
    Else
        lretorno = ""
    End If
    periodoIndenitario = lretorno
End Property


'Public Property Get xxx() As String
'Dim lretorno As String
'    Select Case tipoRelatorio
'        case eTipCalcRPT_CotacIndividual
'        case eTipCalcRPT_CotacColetiva
'        case eTipCalcRPT_PropIndividual
'        case eTipCalcRPT_PropColetiva
'    End Select
'    lRetorno = IIf(lRetorno = "", " ", lRetorno)
'    xxx = lretorno
'End Property
'=========================Funções Privadas na classe
Private Property Get tipoRelatorio() As eTipCalcRPT
    Static lretorno As eTipCalcRPT
    Dim lstcTabPed As stcA46V702B
    If lretorno = eTipCalcRPT_ND Then
        Set lstcTabPed = mStcA46V732B.ObjPai
        lretorno = lstcTabPed.getRpt.tipoRelatorio
    End If
    tipoRelatorio = lretorno
End Property

Public Property Get Num_Meses_PI() As String
    Dim lretorno As String

    'pi mes
    lretorno = Format$(mStcA46V732B.Num_Meses_PI, "00") & " meses "

    lretorno = IIf(lretorno = "", " ", lretorno)
    Num_Meses_PI = lretorno
End Property

Private Sub mpPrazoTabCoefPI(pNumMeses As Integer, pPrazo As Integer)
    Dim lBDP00Multi As ADODB.Connection
    Dim lMensagem As String
    Dim lrst As clsYasRecordSet
    Dim lSelect As String

    If mStcA46V732B.M46V999.gfAbrBasDados("P00MULTI", mStcA46V732B.M46V999.gTipP00MULTI, mStcA46V732B.M46V999.gEndP00MULTI, _
                                          lBDP00Multi, lMensagem) = False Then
        Call mStcA46V732B.M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If
    lSelect = " Select Prazo from Tbm_COEF_PI Where"
    lSelect = lSelect & " Cod_Ramo = " & mStcA46V732B.ObjPai.Cod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & mStcA46V732B.Cod_Cobert
    lSelect = lSelect & " and Num_meses = " & pNumMeses
    lSelect = lSelect & " and Cod_Tarif = " & mStcA46V732B.ObjPai.Cod_Produto
    lSelect = lSelect & " and Cod_Sequencia <= " & mStcA46V732B.ObjPai.Cod_Sequencia

    If mStcA46V732B.M46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
        Call mStcA46V732B.M46V999.gpGraLog(1, lMensagem)
        pPrazo = 0
        GoTo SAIDA
    End If
    pPrazo = 0
    If Not lrst.EOF Then
        pPrazo = IIf(IsNull(lrst("Prazo")), 0, lrst("Prazo"))
    End If

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Call mStcA46V732B.M46V999.gpFecha2(lBDP00Multi)
    Set lBDP00Multi = Nothing
    Set lrst = Nothing
    Return
End Sub

Private Sub Class_Initialize()
    Set mStcA46V732B = Nothing
End Sub

Private Sub Class_Terminate()
    Set mStcA46V732B = Nothing
End Sub




