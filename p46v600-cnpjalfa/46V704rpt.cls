VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V704rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStcA46V704B As stcA46V704B

Public Property Set stcOrigem(ByRef pStcA46V704B As stcA46V704B)
    Set mStcA46V704B = pStcA46V704B
End Property

Public Property Get NumItem() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format$(mStcA46V704B.Num_Item, "0000")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    NumItem = lretorno
End Property

Public Property Get CodCobert() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format$(mStcA46V704B.Cod_Cobert, "0000")
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
    'lretorno = Left(Trim(mStcA46V704B.gNomeCobertura) & " " & Trim(mStcA46V704B.Compl_Desc_Cobertura), 80)
    If mStcA46V704B.Cod_Cobert = 1103 Then
        lretorno = IIf(Len(Trim(mStcA46V704B.Compl_Desc_Cobertura)) = 0, Left(Trim(mStcA46V704B.gNomeCobertura), 80), Left(Trim(mStcA46V704B.gNomeCobertura), 80) & " - " & Trim(mStcA46V704B.Compl_Desc_Cobertura))
    Else
        lretorno = IIf(Len(Trim(mStcA46V704B.Compl_Desc_Cobertura)) = 0, Left(Trim(mStcA46V704B.gNomeCobertura), 80), "(" & Left(Trim(mStcA46V704B.Compl_Desc_Cobertura), 1) & ") " & Left(Trim(mStcA46V704B.gNomeCobertura), 80))
    End If
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
    lretorno = Trim(Format$(mStcA46V704B.Val_IS, "standard"))
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    limMaxGarantia = lretorno
End Property
Public Property Get premioLiq() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Trim(Format$(mStcA46V704B.VAL_PR, "standard"))
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    premioLiq = lretorno
End Property

Public Property Get temFranquia() As Boolean
    Dim lretorno As Boolean
    lretorno = ((mStcA46V704B.Num_Meses_PI <> 0 And mStcA46V704B.M46V999.gCodProduto < 1080) And mStcA46V704B.Per_Franq <> 0 Or mStcA46V704B.Val_Franq_Min <> 0 Or mStcA46V704B.gPrazo <> 0 Or mStcA46V704B.TextoFranquia <> "")
    temFranquia = lretorno
End Property
Public Property Get codFranquia() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format$(mStcA46V704B.Cod_Cobert, "0000")
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
    'lRetorno = Left(mStcA46V704B.gNomeCobertura, 40)
    lretorno = Left(mStcA46V704B.gNomeFranquia, 80)
    '1100864 - Tarifa 1018 e 1019 Setembro 2011
    If (mStcA46V704B.ObjPai.Cod_Iden_Tarif >= 1018) And (mStcA46V704B.Cod_Ramo = e_CodRamo_Empresarial) And _
       (val(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(282, 2)) And (val(mStcA46V704B.ObjPai.Cod_Compl_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(282, 2, 10)) And _
       mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(23, 1) Then
        lretorno = lretorno & " (DE HÓSPEDES)"
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    nomFranquia = lretorno
End Property

Public Property Get descFranquia_old() As String
    Dim lretorno As String
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer

    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V704B.Per_Franq <> 0 Then
        '%Sobre Prejuízos
        lAuxiliar1 = mStcA46V704B.Per_Franq
        lAuxiliar3 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar3)
        If lTamanho < 10 Then
            lAuxiliar3 = Space(10 - lTamanho) & lAuxiliar3
        End If
        lAuxiliar2 = "- " & lAuxiliar3 & " % sobre prej.indeniz"
        'Limitado ao Mínimo
        lAuxiliar1 = mStcA46V704B.Val_Franq_Min
        lAuxiliar3 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar3)
        If lTamanho < 5 Then
            lAuxiliar3 = Space(5 - lTamanho) & lAuxiliar3
        End If
'        If mStcA46V704B.Val_Franq_Min > 0 Then
'            lAuxiliar2 = lAuxiliar2 & ", limitado ao mínimo de " & IIf(lAuxiliar3 = "", " ", lAuxiliar3)
'        Else
            lAuxiliar2 = lAuxiliar2 & "."
'        End If
        '1100589 - Tarifa 1014 e 1015 Julho 2011
        If mStcA46V704B.ObjPai.Cod_Iden_Tarif >= 1014 And mStcA46V704B.Cod_Ramo = e_CodRamo_Condominio Then
            'cobertura basica
            If mStcA46V704B.Cod_Cobert = 1 Then
                If mStcA46V704B.ObjPai.Cod_Plano = 1 Then
                    'se plano simples
                    lAuxiliar2 = Format$(mStcA46V704B.Per_Franq, "standard") & " % prej.indeniz, limitado ao mínimo " & Format$(mStcA46V704B.Val_Franq_Min, "standard") & " p/ queda de raio"
                Else
                    'se plano amplo
                    lAuxiliar2 = Format$(mStcA46V704B.Per_Franq, "standard") & "% do LMG da Cobertura Básica Ampla."
                End If
            End If
        End If
        '1100705 - Tarifa 1016 e 1017 Agosto 2011
        If (mStcA46V704B.ObjPai.Cod_Iden_Tarif >= 1016) And (mStcA46V704B.Cod_Ramo = e_CodRamo_Empresarial) Then
            If (val(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(22, 2) And val(mStcA46V704B.ObjPai.Cod_Compl_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(22, 2, 13)) And _
               (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(56, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(57, 1) Or _
                mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(122, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(123, 1) Or _
                mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(124, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(125, 1)) Then
                lAuxiliar2 = Format$(mStcA46V704B.Val_Franq_Min, "standard")
                lTamanho = Len(lAuxiliar2)
                If lTamanho < 10 Then
                    lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar3
                End If
                lAuxiliar2 = "- " & lAuxiliar2
                lAuxiliar2 = lAuxiliar2 & " - Por Veículo (vide Observações)"
            End If
        End If
        '1100864 - Tarifa 1018 e 1019 Setembro 2011
        If (mStcA46V704B.ObjPai.Cod_Iden_Tarif >= 1018) And (mStcA46V704B.Cod_Ramo = e_CodRamo_Empresarial) And _
           (val(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(282, 2)) And (val(mStcA46V704B.ObjPai.Cod_Compl_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(22, 2, 10)) And _
           (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(23, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(38, 1)) Then
            lAuxiliar2 = Format$(mStcA46V704B.Per_Franq, "standard") & "% s/ prej. indeniz, lim. mín. " & Format$(mStcA46V704B.Val_Franq_Min, "standard")
            If (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(38, 1)) Or (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(23, 1) And InStr(1, "AC;AM;AP;RO;RR", mStcA46V704B.ObjPai.Sig_UF_Risco, vbTextCompare)) Then
                lAuxiliar2 = lAuxiliar2 & " (Vide Observações)"
            End If
        End If
        lretorno = lAuxiliar2
        '                GeraCYX_PROP2 pLinhaCXY, "franquia " & lCountLine, pPosY, 10, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        '                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ElseIf mStcA46V704B.Val_Franq_Min <> 0 Then
        'Limitado ao Mínimo
        lAuxiliar1 = mStcA46V704B.Val_Franq_Min
        lAuxiliar3 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar3)
        If lTamanho < 10 Then
            lAuxiliar3 = Space(10 - lTamanho) & lAuxiliar3
        End If
        lAuxiliar3 = "- " & lAuxiliar3
        '1100705 - Tarifa 1016 e 1017 Agosto 2011
        If (mStcA46V704B.ObjPai.Cod_Iden_Tarif >= 1016) And (mStcA46V704B.Cod_Ramo = e_CodRamo_Empresarial) Then
            If (val(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(22, 2) And val(mStcA46V704B.ObjPai.Cod_Compl_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(22, 2, 13)) And _
               (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(56, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(57, 1) Or _
                mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(122, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(123, 1) Or _
                mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(124, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(125, 1)) Then
                lAuxiliar3 = lAuxiliar3 & " - Por Veículo (vide Observações)"
            End If
        End If
        lretorno = lAuxiliar3
        '                GeraCYX_PROP2 pLinhaCXY, "franquia " & lCountLine, pPosY, 10, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        '                Print #pNumArqDados, IIf(lAuxiliar3 = "", " ", lAuxiliar3)
    Else
        lretorno = "-      Vide Observações"
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select


    '1200903 - Tarifa - 1046 e 1047 - Item 15, 16 e 17
    Dim blnFranqEspec_59_e_60 As Boolean
    blnFranqEspec_59_e_60 = False
    Dim blnFranqEspec_22_59_e_60 As Boolean
    blnFranqEspec_22_59_e_60 = False
    If mStcA46V704B.ObjPai.Cod_Iden_Tarif >= 1046 Then
        If Trim(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("420", 2) Then   'Item 15 - Padarias
            blnFranqEspec_59_e_60 = True
        ElseIf Trim(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("061", 2) Then   'Item 16 - Bares e Restaurantes
            blnFranqEspec_59_e_60 = True
        ElseIf (Trim(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("282", 2) And _
                Trim(mStcA46V704B.ObjPai.Cod_Compl_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("282", 2, "10")) Then    'Item 16 - Hoteis e Pousadas
            blnFranqEspec_59_e_60 = True
        ElseIf (Trim(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("282A", 2) And _
                Trim(mStcA46V704B.ObjPai.Cod_Compl_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("282A", 2, "10")) Then  'Item 16 - Hoteis e Pousadas
            blnFranqEspec_59_e_60 = True
        ElseIf (Trim(mStcA46V704B.ObjPai.Cod_Clasf) = "496" And _
                Trim(mStcA46V704B.ObjPai.Cod_Compl_Clasf) = "10") Then  'Item 17 - SuperMercados
            blnFranqEspec_22_59_e_60 = True
        Else
            blnFranqEspec_59_e_60 = False
            blnFranqEspec_22_59_e_60 = False
        End If
    End If

    If (blnFranqEspec_59_e_60 = True And (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(59, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(60, 1))) Or _
       (blnFranqEspec_22_59_e_60 = True And (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(22, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(59, 1) Or _
       mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(60, 1))) Then   '1200903 - Tarifa - 1046 e 1047 - Item 15, 16 e 17
        lretorno = "-      Vide Observações"
    End If


    lretorno = IIf(lretorno = "", " ", lretorno)
    descFranquia_old = lretorno
End Property


Public Property Get descFranquia() As String
    Dim lretorno As String
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer


    If Trim(mStcA46V704B.TextoFranquia) <> "" Then
        descFranquia = mStcA46V704B.TextoFranquia
        Exit Property
    End If

    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V704B.Per_Franq <> 0 Then
        '%Sobre Prejuízos
        lAuxiliar1 = mStcA46V704B.Per_Franq
        lAuxiliar3 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar3)
        If lTamanho < 10 Then
            lAuxiliar3 = Space(10 - lTamanho) & lAuxiliar3
        End If
        lAuxiliar2 = "- " & lAuxiliar3 & " % sobre prej.indeniz"
        'Limitado ao Mínimo
        lAuxiliar1 = mStcA46V704B.Val_Franq_Min
        lAuxiliar3 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar3)
        If lTamanho < 5 Then
            lAuxiliar3 = Space(5 - lTamanho) & lAuxiliar3
        End If
'        If mStcA46V704B.Val_Franq_Min > 0 Then
'            lAuxiliar2 = lAuxiliar2 & ", limitado ao mínimo de " & IIf(lAuxiliar3 = "", " ", lAuxiliar3)
'        Else
            lAuxiliar2 = lAuxiliar2 & "."
'        End If
        '1100589 - Tarifa 1014 e 1015 Julho 2011
        If mStcA46V704B.ObjPai.Cod_Iden_Tarif >= 1014 And mStcA46V704B.Cod_Ramo = e_CodRamo_Condominio Then
            'cobertura basica
            If mStcA46V704B.Cod_Cobert = 1 Then
                If mStcA46V704B.ObjPai.Cod_Plano = 1 Then
                    'se plano simples
                    lAuxiliar2 = Format$(mStcA46V704B.Per_Franq, "standard") & " % prej.indeniz, limitado ao mínimo " & Format$(mStcA46V704B.Val_Franq_Min, "standard") & " p/ queda de raio"
                Else
                    'se plano amplo
                    lAuxiliar2 = Format$(mStcA46V704B.Per_Franq, "standard") & "% do LMG da Cobertura Básica Ampla."
                End If
            End If
        End If
        '1100705 - Tarifa 1016 e 1017 Agosto 2011
        If (mStcA46V704B.ObjPai.Cod_Iden_Tarif >= 1016) And (mStcA46V704B.Cod_Ramo = e_CodRamo_Empresarial) Then
            If (val(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(22, 2) And val(mStcA46V704B.ObjPai.Cod_Compl_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(22, 2, 13)) And _
               (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(56, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(57, 1) Or _
                mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(122, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(123, 1) Or _
                mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(124, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(125, 1)) Then
                lAuxiliar2 = Format$(mStcA46V704B.Val_Franq_Min, "standard")
                lTamanho = Len(lAuxiliar2)
                If lTamanho < 10 Then
                    lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar3
                End If
                lAuxiliar2 = "- " & lAuxiliar2
                If mStcA46V704B.ObjPai.Cod_Iden_Tarif <= 1057 Then
                    lAuxiliar2 = lAuxiliar2 & " - Por Veículo (vide Observações)"
                End If
            End If
        End If
        '1100864 - Tarifa 1018 e 1019 Setembro 2011
        If (mStcA46V704B.ObjPai.Cod_Iden_Tarif >= 1018) And (mStcA46V704B.Cod_Ramo = e_CodRamo_Empresarial) And _
           (val(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(282, 2)) And (val(mStcA46V704B.ObjPai.Cod_Compl_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(282, 2, 10)) And _
           (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(23, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(38, 1)) Then
            lAuxiliar2 = Format$(mStcA46V704B.Per_Franq, "standard") & "% s/ prej. indeniz, lim. mín. " & Format$(mStcA46V704B.Val_Franq_Min, "standard")
            If (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(38, 1)) Or (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(23, 1) And InStr(1, "AC;AM;AP;RO;RR", mStcA46V704B.ObjPai.Sig_UF_Risco, vbTextCompare)) Then
                If mStcA46V704B.ObjPai.Cod_Iden_Tarif <= 1057 Then
                    lAuxiliar2 = lAuxiliar2 & " (Vide Observações)"
                End If
            End If
        End If
        lretorno = lAuxiliar2
        '                GeraCYX_PROP2 pLinhaCXY, "franquia " & lCountLine, pPosY, 10, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        '                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ElseIf mStcA46V704B.Val_Franq_Min <> 0 Then
        'Limitado ao Mínimo
        lAuxiliar1 = mStcA46V704B.Val_Franq_Min
        lAuxiliar3 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar3)
        If lTamanho < 10 Then
            lAuxiliar3 = Space(10 - lTamanho) & lAuxiliar3
        End If
        lAuxiliar3 = "- " & lAuxiliar3
        '1100705 - Tarifa 1016 e 1017 Agosto 2011
        If (mStcA46V704B.ObjPai.Cod_Iden_Tarif >= 1016) And (mStcA46V704B.Cod_Ramo = e_CodRamo_Empresarial) Then
            If (val(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(22, 2) And val(mStcA46V704B.ObjPai.Cod_Compl_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara(22, 2, 13)) And _
               (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(56, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(57, 1) Or _
                mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(122, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(123, 1) Or _
                mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(124, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(125, 1)) Then
                If mStcA46V704B.ObjPai.Cod_Iden_Tarif <= 1057 Then
                    lAuxiliar3 = lAuxiliar3 & " - Por Veículo (vide Observações)"
                End If
            End If
        End If
        lretorno = lAuxiliar3
        '                GeraCYX_PROP2 pLinhaCXY, "franquia " & lCountLine, pPosY, 10, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        '                Print #pNumArqDados, IIf(lAuxiliar3 = "", " ", lAuxiliar3)
    Else
        If mStcA46V704B.ObjPai.Cod_Iden_Tarif <= 1057 Then
            lretorno = "-      Vide Observações"
        End If
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select


    '1200903 - Tarifa - 1046 e 1047 - Item 15, 16 e 17
    Dim blnFranqEspec_59_e_60 As Boolean
    blnFranqEspec_59_e_60 = False
    Dim blnFranqEspec_22_59_e_60 As Boolean
    blnFranqEspec_22_59_e_60 = False
    If mStcA46V704B.ObjPai.Cod_Iden_Tarif >= 1046 Then
        If Trim(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("420", 2) Then    'Item 15 - Padarias
            blnFranqEspec_59_e_60 = True
        ElseIf Trim(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("061", 2) Then    'Item 16 - Bares e Restaurantes
            blnFranqEspec_59_e_60 = True
        ElseIf (Trim(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("282", 2) And _
                Trim(mStcA46V704B.ObjPai.Cod_Compl_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("282", 2, "10")) Then    'Item 16 - Hoteis e Pousadas
            blnFranqEspec_59_e_60 = True
        ElseIf (Trim(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("282A", 2) And _
                Trim(mStcA46V704B.ObjPai.Cod_Compl_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("282", 2, "10")) Then  'Item 16 - Hoteis e Pousadas
            blnFranqEspec_59_e_60 = True
        ElseIf (Trim(mStcA46V704B.ObjPai.Cod_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("496", 2) And _
                Trim(mStcA46V704B.ObjPai.Cod_Compl_Clasf) = mStcA46V704B.M46V999.gfBuscarDePara("496", 2, "10")) Then  'Item 17 - SuperMercados
            blnFranqEspec_22_59_e_60 = True
        Else
            blnFranqEspec_59_e_60 = False
            blnFranqEspec_22_59_e_60 = False
        End If
    End If

    If (blnFranqEspec_59_e_60 = True And (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(59, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(60, 1))) Or _
       (blnFranqEspec_22_59_e_60 = True And (mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(22, 1) Or mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(59, 1) Or _
       mStcA46V704B.Cod_Cobert = mStcA46V704B.M46V999.gfBuscarDePara(60, 1))) Then   '1200903 - Tarifa - 1046 e 1047 - Item 15, 16 e 17
        If mStcA46V704B.ObjPai.Cod_Iden_Tarif <= 1057 Then
            lretorno = "-      Vide Observações"
        End If
    End If


    lretorno = IIf(lretorno = "", " ", lretorno)
    descFranquia = lretorno
End Property


Public Property Get descFranquiaPrazo() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V704B.gPrazo <> 0 Then
        lretorno = mStcA46V704B.gPrazo & " dias."
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
    If mStcA46V704B.Num_Meses_PI <> 0 Then
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
        Set lstcTabPed = mStcA46V704B.ObjPai.ObjPai
        lretorno = lstcTabPed.getRpt.tipoRelatorio
    End If
    tipoRelatorio = lretorno
End Property

Public Property Get Num_Meses_PI() As String
    Dim lretorno As String

    'pi mes
    lretorno = Format$(mStcA46V704B.Num_Meses_PI, "00") & " meses "

    lretorno = IIf(lretorno = "", " ", lretorno)
    Num_Meses_PI = lretorno
End Property

Private Sub mpPrazoTabCoefPI(pNumMeses As Integer, pPrazo As Integer)
    Dim lBDP00Multi As ADODB.Connection
    Dim lMensagem As String
    Dim lrst As clsYasRecordSet
    Dim lSelect As String

    If mStcA46V704B.M46V999.gfAbrBasDados("P00MULTI", mStcA46V704B.M46V999.gTipP00MULTI, mStcA46V704B.M46V999.gEndP00MULTI, _
                                          lBDP00Multi, lMensagem) = False Then
        Call mStcA46V704B.M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If
    lSelect = " Select Prazo from Tbm_COEF_PI Where"
    lSelect = lSelect & " Cod_Ramo = " & mStcA46V704B.ObjPai.Cod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & mStcA46V704B.Cod_Cobert
    lSelect = lSelect & " and Num_meses = " & pNumMeses
    lSelect = lSelect & " and Cod_Tarif = " & mStcA46V704B.ObjPai.Cod_Iden_Tarif
    lSelect = lSelect & " and Cod_Sequencia <= " & mStcA46V704B.ObjPai.ObjPai.Cod_Sequencia

    If mStcA46V704B.M46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
        Call mStcA46V704B.M46V999.gpGraLog(1, lMensagem)
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
    Call mStcA46V704B.M46V999.gpFecha2(lBDP00Multi)
    Set lBDP00Multi = Nothing
    Set lrst = Nothing
    Return
End Sub

Private Sub Class_Initialize()
    Set mStcA46V704B = Nothing
End Sub

Private Sub Class_Terminate()
    Set mStcA46V704B = Nothing
End Sub


Public Property Get Observacao_Franquia() As String
    Dim lretorno As String
    Dim lobjstcPedFranq As stcA46V738B

    For Each lobjstcPedFranq In mStcA46V704B.ColecaoNovaFranquia
        lretorno = lobjstcPedFranq.Texto_Observacoes
    Next lobjstcPedFranq
    
    lretorno = IIf(lretorno = "", " ", lretorno)
    Observacao_Franquia = lretorno
End Property
