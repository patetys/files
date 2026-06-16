VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V708rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStcA46V708B As stcA46V708B
Private strTextoAdicional As String

Private Sub Class_Initialize()
    Set mStcA46V708B = Nothing
End Sub

Private Sub Class_Terminate()
    Set mStcA46V708B = Nothing
End Sub

Public Property Set stcOrigem(ByRef pStcA46V708B As stcA46V708B)
    Set mStcA46V708B = pStcA46V708B
End Property

Public Property Get NumItem() As String
    Dim lretorno As String
    lretorno = Format$(mStcA46V708B.Num_Item, "0000")
    Select Case tipoRelatorio
    Case eTipCalcRPT_PropColetivaLMI, eTipCalcRPT_CotacColetivaLMI
        lretorno = lretorno & "          " & IIf(mStcA46V708B.Insp_Obr, "INSP", "")
    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    NumItem = lretorno
End Property

Public Property Get LogrRisco() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.Cod_Tip_Logr_Risco & " " & mStcA46V708B.Nom_Logr_Risco
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    LogrRisco = lretorno
End Property

Public Property Get NumLogrRisco() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.objstcPedInsp.Num_Logr_Risco
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    NumLogrRisco = lretorno
End Property

Public Property Get DscComplRisco() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.Dsc_Compl_Risco
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    DscComplRisco = lretorno
End Property


Public Property Get Val_Vr_Lc() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = "R$ " & Format(mStcA46V708B.Val_Vr_Lc, "standard")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Val_Vr_Lc = lretorno
End Property

Public Property Get NomBairroRisco() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.Nom_Bairro_Risco
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select

    'CR29035 - Ajuste para abreviar a descrição do bairro pois o mesmo quando era maior que 24 posições estava sobrescrevendo o cep.
    'Abreviar se tiver a palavra PARQUE no bairro
    If Len(lretorno) > 24 Then
        If InStr(1, lretorno, "PARQUE", vbTextCompare) > 0 Then
            lretorno = Replace(lretorno, "PARQUE", "PQ.", , , vbTextCompare)
        End If
    End If

    'Abreviar se tiver a palavra RESIDENCIAL no bairro
    If Len(lretorno) > 24 Then
        If InStr(1, lretorno, "RESIDENCIAL", vbTextCompare) > 0 Then
            lretorno = Replace(lretorno, "RESIDENCIAL", "RES.", , , vbTextCompare)
        End If
    End If

    'Se mesmo com as abreviações ainda for maior, truncar em 24 posições.
    If Len(lretorno) > 24 Then
        lretorno = Mid(lretorno, 1, 24)
    End If

    lretorno = IIf(lretorno = "", " ", lretorno)
    NomBairroRisco = lretorno
End Property

Public Property Get CepRisco() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.gCEP_Formatado
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    CepRisco = lretorno
End Property

Public Property Get CidadeRisco() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.Nom_Cid_Risco
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    CidadeRisco = lretorno
End Property

Public Property Get UfRisco() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.Sig_UF_Risco
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    UfRisco = lretorno
End Property

Public Property Get TipoConstrucao() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = " "
    If codRamo = e_CodRamo_Empresarial Then
        Select Case mStcA46V708B.Cod_Constr
        Case IIf(mStcA46V708B.M46V999.gCodProduto >= 1080, e_CodContr_NvTarifa_Superior, e_CodContr_Superior)
            If mStcA46V708B.M46V999.gCodProduto >= 1080 Then
                lretorno = "Superior"
            Else
            lretorno = "Superior/sólida"
            End If
        Case IIf(mStcA46V708B.M46V999.gCodProduto >= 1080, e_CodContr_NvTarifa_Inferior, e_CodContr_Inferior)
            lretorno = "Inferior"
        Case IIf(mStcA46V708B.M46V999.gCodProduto >= 1080, e_CodContr_NvTarifa_Mista, e_CodContr_Mista)
            lretorno = "Mista"
        Case Else
            If mStcA46V708B.M46V999.gCodProduto >= 1080 Then
                lretorno = "Sólida"
            Else
            lretorno = "Superior/sólida"
            End If
        End Select
    Else
        Select Case mStcA46V708B.Cod_Constr
        Case IIf(mStcA46V708B.M46V999.gCodProduto >= 1080, e_CodContr_NvTarifa_Superior, e_CodContr_Superior)
            If mStcA46V708B.M46V999.gCodProduto >= 1080 Then
                lretorno = "Superior"
            Else
            lretorno = "Superior/sólida"
            End If
        Case IIf(mStcA46V708B.M46V999.gCodProduto >= 1080, e_CodContr_NvTarifa_Inferior, e_CodContr_Inferior)
            If mStcA46V708B.M46V999.gCodProduto >= 1080 Then
                lretorno = "Inferior"
            Else
            lretorno = "Inferior/mista"
            End If
        Case IIf(mStcA46V708B.M46V999.gCodProduto >= 1080, e_CodContr_NvTarifa_Mista, e_CodContr_Mista)
            lretorno = "Mista"
        Case Else
            If mStcA46V708B.M46V999.gCodProduto >= 1080 Then
                lretorno = "Sólida"
            Else
            lretorno = "Superior/sólida"
            End If
        End Select
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    TipoConstrucao = lretorno
End Property

Public Property Get BensCompreendidos() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    'If codRamo = e_CodRamo_Empresarial Or codRamo = e_CodRamo_Residencial Then
    Select Case mStcA46V708B.Tip_Cobert

    Case e_TipCobert_PredioeConteudo
        lretorno = "   Prédio/Conteúdo"

    Case e_TipCobert_Predio
        lretorno = "        Prédio    "

    Case e_TipCobert_Conteudo
        lretorno = "       Conteúdo   "

    Case Else
        lretorno = "   Prédio/Conteúdo"

    End Select
    'Else
    '    If mStcA46V708B.Cod_Plano = 2168 Then
    '        lretorno = "        Prédio    "
    '    Else
    '        lretorno = "   Prédio/Conteúdo"
    '    End If
    'End If
    'lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    'lretorno = IIf(lretorno = "", " ", lretorno)

    BensCompreendidos = lretorno
End Property

Public Property Get Classificacao() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.Cod_Clasf & "-" & mStcA46V708B.Cod_Compl_Clasf & "- " & mStcA46V708B.Dsc_Clasf
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Classificacao = lretorno
End Property

Public Property Get ValorRisco() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = "R$ " & Format(mStcA46V708B.Val_Vr, "standard")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    ValorRisco = lretorno
End Property

Public Property Get PlanoProduto() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = " PLANO: " & descPlano
    lretorno = lretorno & " - PRODUTO: " & Format(mStcA46V708B.Cod_Produto, "0000")
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    PlanoProduto = lretorno
End Property

Public Property Get Regiao() As String
    Dim lbdTbm_Regiao As ADODB.Connection
    Dim lrsTbm_Regiao As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String

    Static lsCodRamo As e_CodRamo
    Static lsCodModalidade As Long
    Static lsTarifa As Integer
    Static lsCodSequencia As Integer
    Static lsNum_Cep_Risco As Long
    Static lsRetorno As String

    If mStcA46V708B.Cod_Iden_Tarif <= mStcA46V708B.M46V999.TarifaAntiga Then
        lsRetorno = mStcA46V708B.Cod_Regiao & " - " & Left(mStcA46V708B.gDesc_Regiao, 38)
        lsRetorno = IIf(lsRetorno = "", " ", lsRetorno)
    Else
        'Nova Região
        If (lsCodRamo <> codRamo) Or _
           (lsCodModalidade <> mStcA46V708B.Cod_Produto) Or _
           (lsTarifa <> mStcA46V708B.Cod_Iden_Tarif) Or _
           (lsCodSequencia <> CodSequencia) Or _
           (lsNum_Cep_Risco <> mStcA46V708B.Num_Cep_Risco) Then

            lsCodRamo = codRamo
            lsCodModalidade = mStcA46V708B.Cod_Produto
            lsTarifa = mStcA46V708B.Cod_Iden_Tarif
            lsCodSequencia = CodSequencia
            lsNum_Cep_Risco = mStcA46V708B.Num_Cep_Risco

            If mStcA46V708B.M46V999.gfAbrBasDados("P00MULTI", mStcA46V708B.M46V999.gTipP00MULTI, mStcA46V708B.M46V999.gEndP00MULTI, lbdTbm_Regiao, lMensagem) = False Then
                Call mStcA46V708B.M46V999.gpGraLog(0, lMensagem)
                lsRetorno = ""
                GoTo Retorno
            End If

            lSelect = "Select Cod_Faixa_Reg, DSC_REG from Tbm_Faixa_Regiao "
            lSelect = lSelect & " Where cod_Ramo = " & lsCodRamo
            lSelect = lSelect & "    and Cod_Modalidade = " & lsCodModalidade
            lSelect = lSelect & "    and Cod_Tarif = " & lsTarifa
            lSelect = lSelect & "    and COD_Sequencia <= " & lsCodSequencia
            lSelect = lSelect & "    and Cep_Ini <= " & lsNum_Cep_Risco
            lSelect = lSelect & "    and Cep_Fim >= " & lsNum_Cep_Risco
            lSelect = lSelect & "    order by Cod_Sequencia desc "

            If mStcA46V708B.M46V999.gfObtRegistro(lbdTbm_Regiao, lSelect, lrsTbm_Regiao, lMensagem) = False Then
                Call mStcA46V708B.M46V999.gpGraLog(0, lMensagem)
                Call mStcA46V708B.M46V999.gpFecha2(lbdTbm_Regiao)
                GoTo Retorno
            End If

            If Not lrsTbm_Regiao.EOF Then
                If mStcA46V708B.Cod_Iden_Tarif >= 1086 Then
                    lsRetorno = mStcA46V708B.M46V999.Check(lrsTbm_Regiao("Cod_Faixa_Reg")) '& " - " & Left(mStcA46V708B.M46V999.Check(lrsTbm_Regiao("DSC_REG")), 38)
                Else
                lsRetorno = mStcA46V708B.M46V999.Check(lrsTbm_Regiao("Cod_Faixa_Reg")) & " - " & Left(mStcA46V708B.M46V999.Check(lrsTbm_Regiao("DSC_REG")), 38)
                End If
            Else
                lsRetorno = "Não cadastrada."
            End If
            lsRetorno = IIf(lsRetorno = "", " ", lsRetorno)

        End If
    End If

Retorno:
    Regiao = lsRetorno
End Property

Public Property Get TipoEmissao() As String
    Dim lretorno As String
    lretorno = mStcA46V708B.gDesc_TipEmissao
    Select Case mStcA46V708B.Tip_Emissao
    Case e_TipEmissao_RenYasuda, e_TipEmissao_Endosso
        lretorno = lretorno & " Apólice: " & mStcA46V708B.Num_Apol_Anterior & " Item: " & IIf(mStcA46V708B.Num_Item_Anterior = 0, mStcA46V708B.Num_Item, mStcA46V708B.Num_Item_Anterior) & "  Sinistro: " & IIf(mStcA46V708B.Ind_Sinistro = e_SimNao_Sim, "Sim", "Não")
    Case e_TipEmissao_RenCongenere
        lretorno = Trim(lretorno)
    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    TipoEmissao = lretorno
End Property

Public Property Get dadosCongenere1() As String
    Dim lretorno As String
    Dim lMensagem As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    Select Case mStcA46V708B.Tip_Emissao
    Case e_TipEmissao_RenCongenere
        '1000347 - Inclusão de informações da congênere em renovações
        If mStcA46V708B.Ren_Con_Seguradora > 0 Then
            If Not mStcA46V708B.M46V999.gfNomSeguradora(mStcA46V708B.Ren_Con_Seguradora, lretorno, lMensagem) Then
                Call mStcA46V708B.M46V999.gpGraLog(0, lMensagem)
                GoTo Retorno
            End If
        Else
            lretorno = ""
        End If
        lretorno = IIf(lretorno = "", " ", Trim(lretorno))
    Case Else
        lretorno = ""
    End Select
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
Retorno:
    lretorno = IIf(lretorno = "", " ", lretorno)
    dadosCongenere1 = lretorno
End Property

Public Property Get dadosCongenere2() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    Select Case mStcA46V708B.Tip_Emissao
    Case e_TipEmissao_RenCongenere
        '1000347 - Inclusão de informações da congênere em renovações
        lretorno = "Apólice: " & mStcA46V708B.Ren_Con_Apolice & "  Item: " & mStcA46V708B.Ren_Con_Item & "  Sinistro: " & IIf(mStcA46V708B.Ind_Sinistro = e_SimNao_Sim, "Sim", "Não")
    Case Else
        lretorno = ""
    End Select
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
Retorno:
    lretorno = IIf(lretorno = "", " ", lretorno)
    dadosCongenere2 = lretorno
End Property

Public Property Get InspPessoaContato() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.objstcPedInsp.Pess_Cont
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    InspPessoaContato = lretorno
End Property

Public Property Get InspNumTel() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.objstcPedInsp.Num_Tel
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    InspNumTel = lretorno
End Property

Public Property Get ValPrTotCober() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format$(mStcA46V708B.ValPrTotCober, "standard")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    ValPrTotCober = lretorno
End Property

Public Property Get perDescComlUltCob() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format(mStcA46V708B.perDescComlUltCob * 100, "0000")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    perDescComlUltCob = lretorno
End Property

Public Property Get CodCoenferencia_old() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    '   B: 25 / A: 50 / P: 10 / CO: 99999 / C: 99999 Onde:
    '   B = bônus
    '   A = aparelho
    '   P = progressivo
    '   CO = coeficiente de comissão
    '   C = coeficiente LMG/VR
    '   DC = Desconto Comercial

    '2001 a 2010 - Correta
    '2011 a 2020 - GalCorr
    '2021 a 2031 - Mercedes
    '2032 a 2041 - National
    '2042 a 2049 - Manaus
    '2050 - Ballarin
    '2051 a 2060 - Marsh
    '2061 a 2070 - NSK
    '2071 a 2080 - Ajinomoto
    '2081 a 2090 - Nissin
    '2091 a 2100 - Panasonic
    '2101 - Fanel
    '2102 - JLA
    '2111 a 2120 - Embraer
    '2121 a 2130 - Bardella

    '=====B = bônus
    If (CodPlano >= 2001 And CodPlano <= 2049) Or (CodPlano >= 2051 And CodPlano <= 2100) Or (CodPlano >= 2111 And CodPlano <= 2130) Then
        lretorno = "B:00"
    Else
        lretorno = "B:" & Format(mStcA46V708B.Per_Bonus, "00")
    End If

    '=====P = progressivo
    'Ficha 900740
    If CodPlano >= 2021 And CodPlano <= 2031 Then
        lretorno = lretorno & "/P:00"
    Else
        lretorno = lretorno & "/P:" & Format(mStcA46V708B.Per_Desc_Progres, "00")
    End If

    '=====CO = coeficiente de comissão
    lretorno = lretorno & "/CO:" & Format(mStcA46V708B.ObjPai.ValTotalComissao * 100, "0000")

    '=====DC = Desconto Comercial
    'Ficha 900740
    If CodPlano >= 2021 And CodPlano <= 2031 Then
        lretorno = lretorno & "/DC:00"
    Else
        lretorno = lretorno & "/DC:" & mStcA46V708B.getRpt.perDescComlUltCob
    End If

    'Ficha 900550
    '/V1.0 -> versão após ajustar o bônus das coberturas
    lretorno = lretorno & "/V1.0"

    If mStcA46V708B.M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or mStcA46V708B.M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        If mStcA46V708B.Insp_Obr = True Then
            If mStcA46V708B.Insp_Agend = True Then
                lretorno = lretorno & "/INSPS"
            Else
                lretorno = lretorno & "/INSPN"
            End If
        End If
    End If

    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select

    lretorno = IIf(lretorno = "", " ", lretorno)
    CodCoenferencia_old = lretorno

End Property

Public Property Get CodCoenferencia() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    '   B: 25 / A: 50 / P: 10 / CO: 99999 / C: 99999 Onde:
    '   B = bônus
    '   A = aparelho
    '   P = progressivo
    '   CO = coeficiente de comissão
    '   C = coeficiente LMG/VR
    '   DC = Desconto Comercial

    '2001 a 2010 - Correta
    '2011 a 2020 - GalCorr
    '2021 a 2031 - Mercedes
    '2032 a 2041 - National
    '2042 a 2049 - Manaus
    '2050 - Ballarin
    '2051 a 2060 - Marsh
    '2061 a 2070 - NSK
    '2071 a 2080 - Ajinomoto
    '2081 a 2090 - Nissin
    '2091 a 2100 - Panasonic
    '2101 - Fanel
    '2102 - JLA
    '2111 a 2120 - Embraer
    '2121 a 2130 - Bardella

    '=====B = bônus
    If (CodPlano >= 2001 And CodPlano <= 2049) Or (CodPlano >= 2051 And CodPlano <= 2100) Or (CodPlano >= 2111 And CodPlano <= 2130) Then
        lretorno = ""
    Else
        lretorno = Format(mStcA46V708B.Per_Bonus, "00")
    End If

    '=====P = progressivo
    'Ficha 900740
    If CodPlano >= 2021 And CodPlano <= 2031 Then
        lretorno = lretorno & "/00"
    Else
        lretorno = lretorno & "/" & Format(mStcA46V708B.Per_Desc_Progres, "00")
    End If

    '=====CO = coeficiente de comissão
    lretorno = lretorno & "/" & Format(mStcA46V708B.ObjPai.ValTotalComissao * 100, "0000")

    '=====DC = Desconto Comercial
    'Ficha 900740
    If CodPlano >= 2021 And CodPlano <= 2031 Then
        lretorno = lretorno & "/00"
    Else
        lretorno = lretorno & "/" & mStcA46V708B.getRpt.perDescComlUltCob
    End If

    'Ficha 900550
    '/V1.0 -> versão após ajustar o bônus das coberturas
    lretorno = lretorno & "/V1.0"

    ''2016-10-19~2018-02-28: DescontoComercial
    Dim lCobertTemp As stcA46V704B
    Dim perDescComl As Double
    perDescComl = 0
    
    For Each lCobertTemp In mStcA46V708B.ColecaoPedCobert
        If (perDescComl < lCobertTemp.Per_Desc_Coml) Then
            perDescComl = lCobertTemp.Per_Desc_Coml
        End If
    Next
    
    ''2016-12-13: Solicitado para remover
    'If perDescComl > 0 Then
    '    lretorno = lretorno & "/Dc" & Format(perDescComl, "standard")
    'End If
        
    ''2016-10-19: VerbaCom
    If Not IsNull(mStcA46V708B.ObjPai.objstcCtrlEmis.Val_VerbaCom) Then
        If mStcA46V708B.ObjPai.objstcCtrlEmis.Val_VerbaCom > 0 Then
            lretorno = lretorno & "/Vc" & Format(mStcA46V708B.ObjPai.objstcCtrlEmis.Val_VerbaCom, "standard")
        End If
    End If

    If mStcA46V708B.M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or mStcA46V708B.M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        If mStcA46V708B.Insp_Obr = True Then
            If mStcA46V708B.Insp_Agend = True Then
                lretorno = lretorno & "/INSPS"
            Else
                lretorno = lretorno & "/INSPN"
            End If
        End If
    End If

    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select

    '2018-07-25: Verifica se ha Repique
    If (mStcA46V708B.M46V999.GetRepique(mStcA46V708B.ObjPai.nosso_numero)) Then
        lretorno = lretorno & "/REP"
    End If

    lretorno = IIf(lretorno = "", " ", lretorno)
    CodCoenferencia = lretorno

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
    If lretorno = eTipCalcRPT_ND Then
        lretorno = mStcA46V708B.ObjPai.getRpt.tipoRelatorio
    End If
    tipoRelatorio = lretorno
End Property

Private Property Get codRamo() As e_CodRamo
    Static lretorno As e_CodRamo
    If lretorno = e_CodRamo_NaoDefinido Then
        lretorno = mStcA46V708B.ObjPai.Cod_Ramo
    End If
    codRamo = lretorno
End Property

Private Property Get CodProduto() As Integer
    Static lretorno As Integer
    If lretorno = 0 Then
        lretorno = mStcA46V708B.ObjPai.Cod_Produto
    End If
    CodProduto = lretorno
End Property

Private Property Get CodSequencia() As Integer
    Static lretorno As Integer
    If lretorno = 0 Then
        lretorno = mStcA46V708B.ObjPai.Cod_Sequencia
    End If
    CodSequencia = lretorno
End Property

Private Property Get descPlano() As String
    descPlano = mStcA46V708B.ObjPai.Desc_Plano
End Property

'' 20170524: Indicador Oculta Web
Private Property Get IndOcultaWeb() As Integer
    IndOcultaWeb = mStcA46V708B.ObjPai.Ind_Oculta_Web
End Property

Private Property Get CodPlano() As Double
    Static lretorno As Double
    If Not mStcA46V708B.M46V999.gfPreenchido(lretorno) Then
        lretorno = mStcA46V708B.ObjPai.Cod_Estip
    End If
    CodPlano = lretorno
End Property

Public Property Get Observacoes() As Collection
    Dim lCollection As Collection
    Dim lObjstcCorr As stcA46V705B
    Dim lobjstcCobert As stcA46V704B
    Dim lobjstcClau As stcA46V703B
    Dim lObs As Variant
    Dim lCount As Integer
    Dim lPos As Integer
    Dim lAuxiliar As String
    Dim lobjstcTextoEspec As stcA46V726B      'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).
    Dim lobjstcPedDizer As stcA46V722B
    Dim lTemAssistencia24hr As Boolean
    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If mStcA46V708B.M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(mStcA46V708B.M46V999.gDataCorteSompo) Then
            If CLng(mStcA46V708B.M46V999.gDataCorteSompo) <= mStcA46V708B.M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If

    Set lCollection = New Collection
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        Call addCollectionObs(lCollection, "Cotação e L.M.I. para simples consulta, não implica na aceitação automática.")
    End If
    lTemAssistencia24hr = False
    For Each lobjstcClau In mStcA46V708B.ColecaoPedClau
        If lobjstcClau.Cod_Clau = 400 Or lobjstcClau.Cod_Clau = 401 Or lobjstcClau.Cod_Clau = 402 Then
            lTemAssistencia24hr = True
        End If
    Next lobjstcClau
    '1101045 - Tarifa 1022 e 1023 Novembro 2011
    If lTemAssistencia24hr = False Then
        '            If mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        If mStcA46V708B.Cod_Ramo = e_CodRamo_Residencial Then
            If (val(mStcA46V708B.Cod_Clasf) = mStcA46V708B.M46V999.gfBuscarDePara(1, 2)) And _
               (val(mStcA46V708B.Cod_Compl_Clasf) = mStcA46V708B.M46V999.gfBuscarDePara(1, 2, 13) Or val(mStcA46V708B.Cod_Compl_Clasf) = mStcA46V708B.M46V999.gfBuscarDePara(1, 2, 23)) Then
                'nada
            ElseIf mStcA46V708B.Cod_Iden_Tarif < 1080 Then
                Call addCollectionObs(lCollection, "Assistência Domiciliar 24 Horas Gratuita (Exceto para Residências Eventuais).")
            End If
        ElseIf mStcA46V708B.Cod_Ramo = e_CodRamo_Empresarial And mStcA46V708B.M46V999.gCodProduto >= 950 And mStcA46V708B.Cod_Iden_Tarif < 1080 Then
            If mStcA46V708B.Cod_Plano <> 2218 And mStcA46V708B.Cod_Plano <> 2219 And _
               (mStcA46V708B.Cod_Plano < 2228 And mStcA46V708B.Cod_Plano > 2271) Then
                Call addCollectionObs(lCollection, "Assistência 24 Horas Gratuita.")
            End If
        ElseIf mStcA46V708B.Cod_Ramo <> e_CodRamo_Roubos Then
            If mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao And mStcA46V708B.Cod_Iden_Tarif < 1080 Then
                Call addCollectionObs(lCollection, "Assistência 24 Horas Gratuita, restrita para apólice individual.")
            End If
        End If
        '            End If
    End If
    If mStcA46V708B.Cod_Ramo = e_CodRamo_Roubos Then
        Call addCollectionObs(lCollection, "Nos termos do art. 782 do Código Civil Brasileiro, o Segurado obriga-se a declarar à " & lNomeSeguradora & " Seguros, no momento da apresentação da Proposta ")
        Call addCollectionObs(lCollection, "de Seguro, a existência de quaisquer outros seguros contratados nessa ou em outra seguradora que garantam, contra os mesmos riscos, os bens")
        Call addCollectionObs(lCollection, "e interesses segurados. O Segurado deve também comunicar à " & lNomeSeguradora & " Seguros a efetivação posterior de outros seguros conforme acima descrito.")
        Call addCollectionObs(lCollection, "A presente cotação é apresentada apenas para análise de parâmetro de custo não implicando no reconhecimento da aceitação do seguro a qual ")
        Call addCollectionObs(lCollection, "ficará condicionada a verificação das condições favoráveis do(s) risco(s) a ser(em) constatadas mediante inspeção(ões) no(s) local(is). ")
        Call addCollectionObs(lCollection, "A " & lNomeSeguradora & " Seguros se reserva ainda ao direito de alterar a presente cotação caso o(s) risco(s) em questão não apresente(m) condições favoráveis ")
        Call addCollectionObs(lCollection, "e/ou sejam constatadas omissões/divergências nas informações recebidas.")
        Call addCollectionObs(lCollection, "Cotação válida por 15 dias.")
        If mStcA46V708B.ColecaoPedLocBem.Count > 0 Then
            Call addCollectionObs(lCollection, "A soma dos valores dos Bens não poderá ser superior ao valor do LMG Total.")
        End If
        If mStcA46V708B.ObjPai.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = mStcA46V708B.ObjPai.ColecaoPedCorr.Item(1)
            If lObjstcCorr.Per_Corr = 0 Then
                Call addCollectionObs(lCollection, "Prêmio apresentado na forma NET, ou seja, sem carregamento de Comissão de Corretagem.")
            Else
                Call addCollectionObs(lCollection, "Prêmio apresentado na forma Líquida, ou seja, com carregamento de Comissão de Corretagem.")
            End If
        End If
    End If
'    If mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao Then
'        Call addCollectionObs(lCollection, "Na Razão Social não deverá constar a expressão PCPT, conforme informativo Yasuda  01/2004 de 13/04/2004.")
'    End If
    If mStcA46V708B.ObjPai.objstcCtrlEmis.Tip_Cotac = "2" And mStcA46V708B.M46V999.gfPreenchido(mStcA46V708B.ObjPai.objstcCtrlEmis.Num_Cotac) Then
        Call addCollectionObs(lCollection, "Solicitação de Análise:" & mStcA46V708B.ObjPai.objstcCtrlEmis.Num_Cotac)
    End If
    '### INSPECAO
    If mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Proposta Then
        '                Call addCollectionObs(lCollection, "A aceitação desta proposta está condicionada a análise da Yasuda Maritima Seguros até 15 dias de seu recebimento. Este prazo ficará ")
        '                Call addCollectionObs(lCollection, "suspenso até a entrega de documentos complementares e/ou data da inspeção do risco.")
    Else
        If mStcA46V708B.Insp_Obr = True Then
            'Call addCollectionObs(lCollection, "A aceitação deste seguro está condicionada a solicitação e análise da inspeção do risco, antes da entrada da proposta na cia.")
            Call addCollectionObs(lCollection, "A aceitação deste seguro está condicionada a análise da " & lNomeSeguradora & " Seguros.")
        Else
            Call addCollectionObs(lCollection, "A aceitação deste seguro está condicionada a análise da " & lNomeSeguradora & " Seguros.")
        End If
    End If
    'Ficha 901562
    If mStcA46V708B.M46V999.gCodProduto >= 958 And (mStcA46V708B.Cod_Plano = 2050 Or _
                                                    mStcA46V708B.Cod_Plano = 2101 Or _
                                                    mStcA46V708B.Cod_Plano = 2102 Or _
                                                    mStcA46V708B.Cod_Plano = 2142) Then    'Ficha 1000495 - Plano Belacity
        Call addCollectionObs(lCollection, "O presente seguro destina-se a cobrir exclusivamente o prédio (imóvel).")
    End If
    'Ficha 1001070
    If mStcA46V708B.Cod_Plano = 2157 Or mStcA46V708B.Cod_Plano = 2168 Or mStcA46V708B.Cod_Plano = 2169 Then
        Call addCollectionObs(lCollection, "Cláusula Beneficiária a favor do proprietário do imóvel aplicável exclusivamente ao prédio.")
    End If
    If mStcA46V708B.Cod_Plano = 2168 Or mStcA46V708B.Cod_Plano = 2169 Then
        Call addCollectionObs(lCollection, "Seguro exclusivamente para o prédio.")
        For Each lobjstcCobert In mStcA46V708B.ColecaoPedCobert
            If lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(5, 1) Then
                Call addCollectionObs(lCollection, "A cobertura 05 Roubo ou Furto Com Vestígios garante exclusivamente os danos materiais diretamente causados ao prédio ")
                Call addCollectionObs(lCollection, "segurado durante a prática ou tentativa de roubo e/ou furto com vestígios.")
            End If
            If lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(23, 1) Then
                Call addCollectionObs(lCollection, "A cobertura 23 Roubo de Bens garante exclusivamente os danos materiais diretamente causados ao prédio segurado ")
                Call addCollectionObs(lCollection, "durante a prática de roubo ou furto com vestígios, quer o evento se tenha consumado, ou quer se tenha caracterizado")
                Call addCollectionObs(lCollection, "a simples tentativa.")
            End If
        Next lobjstcCobert
    End If
    'Ficha 901378
    If mStcA46V708B.ObjPai.Cod_Tip_Doc <> e_TipDoc_Proposta Then
        Call addCollectionObs(lCollection, "O registro deste plano na SUSEP não implica, por parte da Autarquia, incentivo ou recomendação a sua comercialização.")
        Call addCollectionObs(lCollection, "O segurado poderá consultar a situação cadastral de seu corretor de seguros, no site www.susep.gov.br, por meio de seu ")
        Call addCollectionObs(lCollection, "registro na SUSEP, nome completo, CNPJ ou CPF.")
    End If
    If mStcA46V708B.ObjPai.M46V111.gCod_plano = 2139 Then
        Call addCollectionObs(lCollection, " Este seguro destina-se a cobrir exclusivamente o conteúdo do imóvel.")
    End If

    If mStcA46V708B.M46V999.gfPreenchido(mStcA46V708B.objstcPedObs.Observacao) And Right(mStcA46V708B.objstcPedObs.Observacao, 2) <> "&%" Then
        mStcA46V708B.objstcPedObs.Observacao = mStcA46V708B.objstcPedObs.Observacao & "&%"
    End If

    lObs = Split(mStcA46V708B.objstcPedObs.Observacao, "&%")
    For lCount = 0 To UBound(lObs)
        lAuxiliar = lObs(lCount)
        If mStcA46V708B.M46V999.gfPreenchido(lAuxiliar) Then
            If Len(lAuxiliar) < 100 Then
                Call addCollectionObs(lCollection, lAuxiliar)
            Else
                lPos = 1
                Do
                    Call addCollectionObs(lCollection, Mid(lAuxiliar, lPos, 100))
                    If Len(Mid(lAuxiliar, lPos, 100)) < 100 Then Exit Do
                    lPos = lPos + 100
                Loop
            End If
        End If
    Next lCount

    'Impressão de texto(s) especiais por cobertura(s) e demais condições
    For Each lobjstcCobert In mStcA46V708B.ColecaoPedCobert
        If (mStcA46V708B.ColecaoPedLocBem.Count = 0) Then
            If (mStcA46V708B.Cod_Ramo = e_CodRamo_Empresarial And lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(94, 1)) Or _
               (mStcA46V708B.Cod_Ramo = e_CodRamo_Residencial And (lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(27, 1) Or _
               lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(29, 1) Or lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(30, 1) Or _
               lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(31, 1))) Then
                lAuxiliar = "Para a cobertura " & lobjstcCobert.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores. Opção escolhida pelo corretor: "
                Call addCollectionObs(lCollection, lAuxiliar)
                lAuxiliar = "Entrega da relação juntamente com a proposta de seguro."
                Call addCollectionObs(lCollection, lAuxiliar)
            End If
        End If
        '1101045 - Tarifa 1022 e 1023 Novembro 2011
        If (mStcA46V708B.Cod_Iden_Tarif >= 1022) And _
           (mStcA46V708B.Cod_Ramo = e_CodRamo_Residencial) And _
           (mStcA46V708B.Cod_Compl_Clasf <> 13 And mStcA46V708B.Cod_Compl_Clasf <> 23) And _
           (lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(5, 1) Or lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(16, 1)) Then
            lAuxiliar = "RELAÇÃO DE BENS SEGURADOS: Objetos com valores acima dos Limites de Indenização"
            Call addCollectionObs(lCollection, lAuxiliar)
            'CR27808 Ajuste a referencia nas condições gerais versão 02/2014
            'lAuxiliar = "estabelecidos na Cláusula " & IIf(lobjstcCobert.Cod_Cobert = 5, "4.2.15", "4.2.16") & " das Condições Gerais do Produto Yasuda Residencial, no "
            lAuxiliar = "estabelecidos na Cláusula " & IIf(lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(5, 1), "4.2.13", "4.2.16") & " das Condições Gerais do Produto " & lNomeSeguradora & " Residencial, no "
            Call addCollectionObs(lCollection, lAuxiliar)
            lAuxiliar = "que tange a Cobertura de " & lobjstcCobert.getRpt.nomCobert & ", deverão estar discriminados na "
            Call addCollectionObs(lCollection, lAuxiliar)
            lAuxiliar = "proposta de seguro ( marca e modelo ) e estarão sujeitos a análise de aceitação por parte da " & lNomeSeguradora & " Seguros."
            Call addCollectionObs(lCollection, lAuxiliar)
        End If
        '1101181 - Tarifa 1022 e 1023 Dezembro 2011
        ' Comentado a pedido do produto - Tarifa 1102
        If (mStcA46V708B.Cod_Iden_Tarif <= 1100) And _
           (mStcA46V708B.Cod_Ramo = e_CodRamo_Residencial) And _
           (lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(5, 1)) And _
           (mStcA46V708B.ObjPai.objstcCtrlEmis.Cod_Corr = 309801) Then    '0309801 - HELLNER CORRETAGEM DE SEGUROS LTDA
            'CR27808 Ajuste a referencia nas condições gerais versão 02/2014
            'lAuxiliar = "Fica entendido e acordado que, ao contrário do que consta no item 4.2.15. ROUBO E FURTO COM VESTÍGIOS"
            lAuxiliar = "Fica entendido e acordado que, ao contrário do que consta no item 4.2.13. ROUBO E FURTO COM VESTÍGIOS"
            Call addCollectionObs(lCollection, lAuxiliar)
            lAuxiliar = "das Condições Gerais do produto " & lNomeSeguradora & " Residencial, em caso de sinistro coberto poderão ser indenizados valores"
            Call addCollectionObs(lCollection, lAuxiliar)
            lAuxiliar = "superiores aos limites de indenização previstos para roupas, calçados, cama, mesa e banho, fitas de vídeo e áudio, DVDs, CDs,"
            Call addCollectionObs(lCollection, lAuxiliar)
            lAuxiliar = "discos de vinil, mídias de gravação e livros, desde que conste a discriminação com respectivo valor individual em boletim de"
            Call addCollectionObs(lCollection, lAuxiliar)
            lAuxiliar = "ocorrência, e que o somatório dos prejuízos indenizáveis não ultrapasse o Limite Máximo de Garantia contratado para a cobertura."
            Call addCollectionObs(lCollection, lAuxiliar)
        End If
    Next lobjstcCobert


    If mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao And (mStcA46V708B.Tip_Cobert <> e_TipCobert_Predio And _
                                                               mStcA46V708B.gTemRegProfissoes = True) Then
        If mStcA46V708B.ObjPai.Cod_Ramo = e_CodRamo_Empresarial And mStcA46V708B.ObjPai.Tip_Pessoa = e_TipPessoa_Fisica And mStcA46V708B.Tip_Cobert <> e_TipCobert_Predio Then
            lAuxiliar = "Seguro de Pessoa Física: Na proposta será obrigatório informar o número do registro profissional no respectivo Conselho "
            Call addCollectionObs(lCollection, lAuxiliar)
            lAuxiliar = "da Categoria. Ex.: OAB, CRM, CRO etc. "
            Call addCollectionObs(lCollection, lAuxiliar)
        End If
    End If

    'Ficha 1000305
    If mStcA46V708B.Tip_Cobert = e_TipCobert_Conteudo Then
        lAuxiliar = "O presente seguro destina-se a cobrir exclusivamente o conteúdo do imóvel."
        Call addCollectionObs(lCollection, lAuxiliar)
    ElseIf mStcA46V708B.Tip_Cobert = e_TipCobert_Predio Then
        lAuxiliar = "O presente seguro destina-se a cobrir exclusivamente o prédio (imóvel)."
        Call addCollectionObs(lCollection, lAuxiliar)
    End If

    'Ficha 1100105
    If mStcA46V708B.ObjPai.ColecaoTextoEspec.Count > 0 Then
        For Each lobjstcTextoEspec In mStcA46V708B.ObjPai.ColecaoTextoEspec
            If lobjstcTextoEspec.Cod_Espec <> 8500 Then
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin1) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin1)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin2) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin2)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin3) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin3)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin4) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin4)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin5) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin5)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin6) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin6)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin7) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin7)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin8) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin8)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin9) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin9)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin10) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin10)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin11) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin11)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin12) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin12)
                End If
            End If
        Next lobjstcTextoEspec
    End If

    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select

    '1200004 - Tarifa 1028 e 1029 Janeiro 2012
    Select Case mStcA46V708B.Cod_Plano
    Case 2229, 2230, 2231, 2232, 2233, 2234, 2235, 2236, 2237, 2238
        Call addCollectionObs(lCollection, "Seguro exclusivo para funcionários da Rinnai Brasil Tecnologia de Aquecimento Ltda.")
    End Select

    Call addCollectionObs(lCollection, " ")
    Call addCollectionObs(lCollection, " ")
    For Each lobjstcPedDizer In mStcA46V708B.ObjPai.ColecaoTextoEndo
        Call addCollectionObs(lCollection, lobjstcPedDizer.Dsc_Dizer_Endo)
    Next lobjstcPedDizer

    Set lObjstcCorr = Nothing
    Set lobjstcPedDizer = Nothing
    Set lobjstcCobert = Nothing
    Set lobjstcTextoEspec = Nothing
    Set lobjstcCobert = Nothing
    Set Observacoes = lCollection
    Set lobjstcClau = Nothing

End Property

'JIRA-25945 - Opção "Sem Depreciação" no Condomínio
Public Function BuscaTextoLivre(ByVal codigoTipoTexto As Integer, _
                                ByRef tipoText As String, _
                                ByRef conteudoTexto As String, _
                                ByRef CodigoTexto As Integer, _
                                ByRef mensagemErro As String) As Boolean
                           
    On Error GoTo TRATA_ERRO
    
    Dim db As ADODB.Connection
    Dim rsTexto As ADODB.Recordset
    BuscaTextoLivre = True
    Dim encontrou As Boolean
        
    If (mStcA46V708B.M46V999.gfPreenchido(mensagemErro)) Then GoTo FIM_FUNCTION

    If (Not mStcA46V708B.M46V999.gfAbrBasDados("RAMOSDIVERSOS", mStcA46V708B.M46V999.gTipDIVERSOS, mStcA46V708B.M46V999.gEndDIVERSOS, db, mensagemErro)) Then GoTo FIM_FUNCTION

    Set rsTexto = New ADODB.Recordset
    
    Call rsTexto.Open("EXEC dbo.PR_S_CONFIGURACAO_TEXTO_LIVRE " & CStr(codigoTipoTexto), db, adOpenStatic)
    
    If (Not rsTexto Is Nothing) Then
        If (Not rsTexto.EOF) Then
            encontrou = True
            Call rsTexto.MoveFirst
            'nessa PROC, só volta um registro
            tipoText = rsTexto!APELIDO_TIPO_TEXTO
            conteudoTexto = rsTexto!DESCRICAO_TEXTO
            CodigoTexto = rsTexto!Codigo_texto
        End If
    End If

    If (Not encontrou) Then
        Call Err.Raise(99, , "Não encontrou texto ativo com o tipo " & CStr(codigoTipoTexto))
    End If

FIM_FUNCTION:
    On Error Resume Next
    If (rsTexto.State = 1) Then Call rsTexto.Close
    Set rsTexto = Nothing
    If (db.State = 1) Then Call db.Close
    Set db = Nothing
    On Error GoTo 0
    Exit Function

TRATA_ERRO:
    BuscaTextoLivre = False
    
    mensagemErro = " Function BuscaTextoLivre (Error) - (" & Err.Number & ")-" & Err.Description
    
    GoTo FIM_FUNCTION

End Function

Public Property Get ObservacoesNew() As String
    Dim lCollection As Collection
    Dim lObjstcCorr As stcA46V705B
    Dim lobjstcCobert As stcA46V704B
    Dim lobjstcClau As stcA46V703B

    Dim lObs As Variant
    Dim lCount As Integer
    Dim lPos As Integer
    Dim lAuxiliar As String
    Dim lobjstcTextoEspec As stcA46V726B      'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).
    Dim lobjstcPedDizer As stcA46V722B
    Dim lTemAssistencia24hr As Boolean
    
    Dim lobjstcNovoTexto As stcA46V739B
    Dim lPulaLinha As Boolean
    
    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If mStcA46V708B.M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(mStcA46V708B.M46V999.gDataCorteSompo) Then
            If CLng(mStcA46V708B.M46V999.gDataCorteSompo) <= mStcA46V708B.M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If

    Set lCollection = New Collection
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        Call addCollectionObs(lCollection, "Cotação e L.M.I. para simples consulta, não implica na aceitação automática.", True)
    End If
    
    '*************************************************************************************
    'Tarifa 1162 - 8000046584, Co-corretagem na impressão
    '*************************************************************************************
    Dim lobjstcCoCorr As stcA46V705B
    
    Dim strCorretores As String
    Dim intControleConcat As Integer
    Dim strControleConcat As String
    
    If (mStcA46V708B.ObjPai.ColecaoPedCorr.Count > 1) And (mStcA46V708B.Cod_Iden_Tarif >= 1162) Then
        
        intControleConcat = 1
        Call addCollectionObs(lCollection, "CO-CORRETAGEM: ", False)

        For Each lobjstcCoCorr In mStcA46V708B.ObjPai.ColecaoPedCorr
            
            If mStcA46V708B.ObjPai.ColecaoPedCorr.Count > 2 Then
                strControleConcat = ","
                
                If mStcA46V708B.ObjPai.ColecaoPedCorr.Count - 1 = intControleConcat Then
                    strControleConcat = " e "
                End If
            Else
                    strControleConcat = " e "
            End If
            
            If mStcA46V708B.ObjPai.ColecaoPedCorr.Count = intControleConcat Then
                strControleConcat = "."
            End If
        
            Call addCollectionObs(lCollection, CStr(lobjstcCoCorr.Cod_Corr) + " - " + lobjstcCoCorr.gNom_Corretor + strControleConcat, IIf(mStcA46V708B.ObjPai.ColecaoPedCorr.Count = intControleConcat, True, False))
            
            intControleConcat = intControleConcat + 1
            
        Next lobjstcCoCorr
        
    End If
    
    '*************************************************************************************
    'Tarifa 1162 - 8000046584, Co-corretagem na impressão
    '*************************************************************************************
       
    'JIRA-25945 - Opção "Sem Depreciação" no Condomínio
    'Verificação clausula de não pobjstcPedLoc.ClausulaNaoDepreciacao.
    ' O texto Livre, com o código tipo 1, já foi incluido nesse item?
    ' Se já foi incluido, não vai incluir novamente
    
    Dim pMensagem As String
    
    If (mStcA46V708B.ClausulaNaoDepreciacao = 1) Then
        Dim tipotexto As String
        Dim conteudoTexto As String
        Dim CodigoTexto As Integer
        If (Not BuscaTextoLivre(1, tipotexto, conteudoTexto, CodigoTexto, pMensagem)) Then
            GoTo FechaObj
        End If
    
        Call addCollectionObs(lCollection, conteudoTexto)
        
    End If

    '*************************************************************************************
    'Tarifa 1160 - Item (7) Inclusão de desconto por cobertura Corretora Volkswagen
    '*************************************************************************************
    Dim CobKM As Boolean
    Dim CorrVolks As Boolean
    
    CobKM = False
    CorrVolks = False
    
    'Somente quando o PI está na situação de cotação.
    If (mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao) And (mStcA46V708B.Cod_Iden_Tarif >= 1160) And (mStcA46V708B.ObjPai.M46V111.gCod_plano = 10 Or mStcA46V708B.ObjPai.M46V111.gCod_plano = 11) And (mStcA46V708B.Cod_Clasf = "543" And mStcA46V708B.Cod_Compl_Clasf = "01") Then
        
        'Contratou?
        '3057 - RC Concessionária até 100 Km - Aplicar o desconto de 80% no valor da cobertura
        '3124 - RC Concessionária até 200 Km - Aplicar o desconto de 70% no valor da cobertura
        '3125 - RC Concessionária até 300 KM - Aplicar o desconto de 60% no valor da cobertura
        For Each lobjstcCobert In mStcA46V708B.ColecaoPedCobert
            If lobjstcCobert.Cod_Cobert = "3057" Or lobjstcCobert.Cod_Cobert = "3124" Or lobjstcCobert.Cod_Cobert = "3125" Then
                CobKM = True
            End If
        Next lobjstcCobert
        
        'Corretores.
        'VOLKSWAGEN CORRETORA SEGUROS LTDA   8001
        'VOLKSWAGEN CORRETORA SEGUROS LTDA   8002
        'VOLKSWAGEN CORRETORA SEGUROS LTDA   8003
        'VOLKSWAGEN CORRETORA SEGUROS LTDA   8005
        'VOLKSWAGEN CORRETORA SEGUROS LTDA   8006
        'VOLKSWAGEN CORRETORA SEGUROS LTDA   8008
        'VOLKSWAGEN CORRETORA SEGUROS LTDA   8011
        
        If (mStcA46V708B.ObjPai.objstcCtrlEmis.Cod_Corr = 8001 Or mStcA46V708B.ObjPai.objstcCtrlEmis.Cod_Corr = 8002 Or mStcA46V708B.ObjPai.objstcCtrlEmis.Cod_Corr = 8003 Or mStcA46V708B.ObjPai.objstcCtrlEmis.Cod_Corr = 8005 Or mStcA46V708B.ObjPai.objstcCtrlEmis.Cod_Corr = 8006 Or mStcA46V708B.ObjPai.objstcCtrlEmis.Cod_Corr = 8008 Or mStcA46V708B.ObjPai.objstcCtrlEmis.Cod_Corr = 8011) Then
            CorrVolks = True
        End If
        
        'Inclusão dos textos especificados.
        If CobKM And CorrVolks Then
            Call addCollectionObs(lCollection, "* Esta cotação apenas destaca os aspectos de maior relevância de cada produto, porém não suprime de maneira alguma a leitura das Condições Gerais da Seguradora, a qual deve ser feita integral e atentamente, pois nestas constam as particularidades, os riscos cobertos e os riscos excluídos do seguro.", True)
            Call addCollectionObs(lCollection, "* Esta cotação em hipótese alguma significa a cobertura do risco, a qual somente será concedida de maneira definitiva após a análise da proposta e inspeção favorável dos locais de risco pela Seguradora. A aceitação está vinculada à análise da inspeção do risco, e, de acordo com o resultado do laudo de inspeção, a Seguradora poderá modificar a proposta ora apresentada, como por exemplo: alterar o custo, alterar as franquias, exigir melhorias no risco, ou, até mesmo, recusar a proposta.", True)
        End If
    End If
    '*************************************************************************************
    'Tarifa 1160 - Item (7) Inclusão de desconto por cobertura Corretora Volkswagen
    '*************************************************************************************
    
    lTemAssistencia24hr = False
    For Each lobjstcClau In mStcA46V708B.ColecaoPedClau
        If lobjstcClau.Cod_Clau = 400 Or lobjstcClau.Cod_Clau = 401 Or lobjstcClau.Cod_Clau = 402 Then
            lTemAssistencia24hr = True
        End If
    Next lobjstcClau
    '1101045 - Tarifa 1022 e 1023 Novembro 2011
    If lTemAssistencia24hr = False Then
        '            If mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        If mStcA46V708B.Cod_Ramo = e_CodRamo_Residencial Then
            If (val(mStcA46V708B.Cod_Clasf) = mStcA46V708B.M46V999.gfBuscarDePara(1, 2)) And _
               (val(mStcA46V708B.Cod_Compl_Clasf) = mStcA46V708B.M46V999.gfBuscarDePara(1, 2, 13) Or val(mStcA46V708B.Cod_Compl_Clasf) = mStcA46V708B.M46V999.gfBuscarDePara(22, 2, 23)) Then
                'nada
            ElseIf mStcA46V708B.Cod_Iden_Tarif < 1080 Then
                Call addCollectionObs(lCollection, "Assistência Domiciliar 24 Horas Gratuita (Exceto para Residências Eventuais).", True)
            End If
        ElseIf mStcA46V708B.Cod_Ramo = e_CodRamo_Empresarial And mStcA46V708B.M46V999.gCodProduto >= 950 And mStcA46V708B.Cod_Iden_Tarif < 1080 Then
            If mStcA46V708B.Cod_Plano <> 2218 And mStcA46V708B.Cod_Plano <> 2219 And _
               (mStcA46V708B.Cod_Plano < 2228 And mStcA46V708B.Cod_Plano > 2271) Then
                Call addCollectionObs(lCollection, "Assistência 24 Horas Gratuita.", True)
            End If
        ElseIf mStcA46V708B.Cod_Ramo <> e_CodRamo_Roubos Then
            If mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao And mStcA46V708B.Cod_Iden_Tarif < 1080 Then
                Call addCollectionObs(lCollection, "Assistência 24 Horas Gratuita, restrita para apólice individual.", True)
            End If
        End If
        '            End If
    End If
    If mStcA46V708B.Cod_Ramo = e_CodRamo_Roubos Then
        Call addCollectionObs(lCollection, "Nos termos do art. 782 do Código Civil Brasileiro, o Segurado obriga-se a declarar à " & lNomeSeguradora & " Seguros, no momento da apresentação da Proposta ")
        Call addCollectionObs(lCollection, "de Seguro, a existência de quaisquer outros seguros contratados nessa ou em outra seguradora que garantam, contra os mesmos riscos, os bens")
        Call addCollectionObs(lCollection, "e interesses segurados. O Segurado deve também comunicar à " & lNomeSeguradora & " Seguros a efetivação posterior de outros seguros conforme acima descrito.", True)
        Call addCollectionObs(lCollection, "A presente cotação é apresentada apenas para análise de parâmetro de custo não implicando no reconhecimento da aceitação do seguro a qual ")
        Call addCollectionObs(lCollection, "ficará condicionada a verificação das condições favoráveis do(s) risco(s) a ser(em) constatadas mediante inspeção(ões) no(s) local(is). ", True)
        Call addCollectionObs(lCollection, "A " & lNomeSeguradora & " Seguros se reserva ainda ao direito de alterar a presente cotação caso o(s) risco(s) em questão não apresente(m) condições favoráveis ")
        Call addCollectionObs(lCollection, "e/ou sejam constatadas omissões/divergências nas informações recebidas.", True)
        Call addCollectionObs(lCollection, "Cotação válida por 15 dias.", True)
        If mStcA46V708B.ColecaoPedLocBem.Count > 0 Then
            Call addCollectionObs(lCollection, "A soma dos valores dos Bens não poderá ser superior ao valor do LMG Total.", True)
        End If
        If mStcA46V708B.ObjPai.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = mStcA46V708B.ObjPai.ColecaoPedCorr.Item(1)
            If lObjstcCorr.Per_Corr = 0 Then
                Call addCollectionObs(lCollection, "Prêmio apresentado na forma NET, ou seja, sem carregamento de Comissão de Corretagem.", True)
            Else
                Call addCollectionObs(lCollection, "Prêmio apresentado na forma Líquida, ou seja, com carregamento de Comissão de Corretagem.", True)
            End If
        End If
    End If
    'If mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao Then
    '    Call addCollectionObs(lCollection, "Na Razão Social não deverá constar a expressão PCPT, conforme informativo Yasuda  01/2004 de 13/04/2004.", True)
    'End If
    If mStcA46V708B.ObjPai.objstcCtrlEmis.Tip_Cotac = "2" And mStcA46V708B.M46V999.gfPreenchido(mStcA46V708B.ObjPai.objstcCtrlEmis.Num_Cotac) Then
        Call addCollectionObs(lCollection, "Solicitação de Análise:" & mStcA46V708B.ObjPai.objstcCtrlEmis.Num_Cotac, True)
    End If
    '### INSPECAO
    If mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Proposta Then
        '                Call addCollectionObs(lCollection, "A aceitação desta proposta está condicionada a análise da Yasuda Maritima Seguros até 15 dias de seu recebimento. Este prazo ficará ")
        '                Call addCollectionObs(lCollection, "suspenso até a entrega de documentos complementares e/ou data da inspeção do risco.")
    Else
        If mStcA46V708B.Insp_Obr = True Then
            'Call addCollectionObs(lCollection, "A aceitação deste seguro está condicionada a solicitação e análise da inspeção do risco, antes da entrada da proposta na cia.", True)
            Call addCollectionObs(lCollection, "A aceitação deste seguro está condicionada a análise da " & lNomeSeguradora & " Seguros.", True)
        Else
            Call addCollectionObs(lCollection, "A aceitação deste seguro está condicionada a análise da " & lNomeSeguradora & " Seguros.", True)
        End If
    End If
    'Ficha 901562
    If mStcA46V708B.M46V999.gCodProduto >= 958 And (mStcA46V708B.Cod_Plano = 2050 Or _
                                                    mStcA46V708B.Cod_Plano = 2101 Or _
                                                    mStcA46V708B.Cod_Plano = 2102 Or _
                                                    mStcA46V708B.Cod_Plano = 2142) Then    'Ficha 1000495 - Plano Belacity
        Call addCollectionObs(lCollection, "O presente seguro destina-se a cobrir exclusivamente o prédio (imóvel).", True)
    End If
    'Ficha 1001070
    If mStcA46V708B.Cod_Plano = 2157 Or mStcA46V708B.Cod_Plano = 2168 Or mStcA46V708B.Cod_Plano = 2169 Then
        Call addCollectionObs(lCollection, "Cláusula Beneficiária a favor do proprietário do imóvel aplicável exclusivamente ao prédio.", True)
    End If
    If mStcA46V708B.Cod_Plano = 2168 Or mStcA46V708B.Cod_Plano = 2169 Then
        Call addCollectionObs(lCollection, "Seguro exclusivamente para o prédio.", True)
        For Each lobjstcCobert In mStcA46V708B.ColecaoPedCobert
            If lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(5, 1) Then
                Call addCollectionObs(lCollection, "A cobertura 05 Roubo ou Furto Com Vestígios garante exclusivamente os danos materiais diretamente causados ao prédio ")
                Call addCollectionObs(lCollection, "segurado durante a prática ou tentativa de roubo e/ou furto com vestígios.", True)
            End If
            If lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(23, 1) Then
                Call addCollectionObs(lCollection, "A cobertura 23 Roubo de Bens garante exclusivamente os danos materiais diretamente causados ao prédio segurado ")
                Call addCollectionObs(lCollection, "durante a prática de roubo ou furto com vestígios, quer o evento se tenha consumado, ou quer se tenha caracterizado")
                Call addCollectionObs(lCollection, "a simples tentativa.", True)
            End If
        Next lobjstcCobert
    End If
    'Ficha 901378
    If mStcA46V708B.ObjPai.Cod_Tip_Doc <> e_TipDoc_Proposta Then
        Call addCollectionObs(lCollection, "O registro deste plano na SUSEP não implica, por parte da Autarquia, incentivo ou recomendação a sua comercialização.", True)
        Call addCollectionObs(lCollection, "O segurado poderá consultar a situação cadastral de seu corretor de seguros, no site www.susep.gov.br, por meio de seu ")
        Call addCollectionObs(lCollection, "registro na SUSEP, nome completo, CNPJ ou CPF.", True)
    End If
    If mStcA46V708B.ObjPai.M46V111.gCod_plano = 2139 Then
        Call addCollectionObs(lCollection, " Este seguro destina-se a cobrir exclusivamente o conteúdo do imóvel.", True)
    End If

    If mStcA46V708B.M46V999.gfPreenchido(mStcA46V708B.objstcPedObs.Observacao) And Right(mStcA46V708B.objstcPedObs.Observacao, 2) <> "&%" Then
        mStcA46V708B.objstcPedObs.Observacao = mStcA46V708B.objstcPedObs.Observacao & "&%"
    End If

    lObs = Split(mStcA46V708B.objstcPedObs.Observacao, "&%")
    For lCount = 0 To UBound(lObs)
        lAuxiliar = lObs(lCount)
        Call addCollectionObs(lCollection, lAuxiliar, True)
    Next lCount
    

'    lObs = Split(mStcA46V708B.objstcPedObs.Observacao, "&%")
'    For lCount = 0 To UBound(lObs)
'        lAuxiliar = lObs(lCount)
'        If mStcA46V708B.M46V999.gfPreenchido(lAuxiliar) Then
'            If Len(lAuxiliar) < 100 Then
'                Call addCollectionObs(lCollection, lAuxiliar)
'            Else
'                lPos = 1
'                Do
'                    Call addCollectionObs(lCollection, Mid(lAuxiliar, lPos, 100))
'                    If Len(Mid(lAuxiliar, lPos, 100)) < 100 Then Exit Do
'                    lPos = lPos + 100
'                Loop
'            End If
'        End If
'    Next lCount
'
'    If UBound(lObs) > 0 Then
'        Call addCollectionObs(lCollection, "", True)
'    End If

    'Impressão de texto(s) especiais por cobertura(s) e demais condições
    For Each lobjstcCobert In mStcA46V708B.ColecaoPedCobert
        If (mStcA46V708B.ColecaoPedLocBem.Count = 0) Then
            'RETIRADO A CONSISTENCIA NA NOVA TARIFA 1080 / 1081
            If mStcA46V708B.Cod_Iden_Tarif < 1080 Then
                If (mStcA46V708B.Cod_Ramo = e_CodRamo_Empresarial And lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(94, 1)) Or _
                   (mStcA46V708B.Cod_Ramo = e_CodRamo_Residencial And (lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(27, 1) Or _
                   lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(29, 1) Or lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(30, 1) Or _
                   lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(31, 1))) Then
                    lAuxiliar = "Para a cobertura " & lobjstcCobert.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores. Opção escolhida pelo corretor: "
                    Call addCollectionObs(lCollection, lAuxiliar, True)
                    lAuxiliar = "Entrega da relação juntamente com a proposta de seguro."
                    Call addCollectionObs(lCollection, lAuxiliar, True)
                End If
            End If
        End If
        '1101045 - Tarifa 1022 e 1023 Novembro 2011
'piloto
        If (mStcA46V708B.Cod_Iden_Tarif < 1080) Then
            If (mStcA46V708B.Cod_Iden_Tarif >= 1022) And _
               (mStcA46V708B.Cod_Ramo = e_CodRamo_Residencial) And _
               (mStcA46V708B.Cod_Compl_Clasf <> 13 And mStcA46V708B.Cod_Compl_Clasf <> 23) And _
           (lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(5, 1) Or lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(16, 1)) Then
                lAuxiliar = "RELAÇÃO DE BENS SEGURADOS: Objetos com valores acima dos Limites de Indenização"
                Call addCollectionObs(lCollection, lAuxiliar)
                'CR27808 Ajuste a referencia nas condições gerais versão 02/2014
                'lAuxiliar = "estabelecidos na Cláusula " & IIf(lobjstcCobert.Cod_Cobert = 5, "4.2.15", "4.2.16") & " das Condições Gerais do Produto Yasuda Residencial, no "
            lAuxiliar = "estabelecidos na Cláusula " & IIf(lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(5, 1), "4.2.13", "4.2.16") & " das Condições Gerais do Produto " & lNomeSeguradora & " Residencial, no "
                Call addCollectionObs(lCollection, lAuxiliar)
                lAuxiliar = "que tange a Cobertura de " & lobjstcCobert.getRpt.nomCobert & ", deverão estar discriminados na "
                Call addCollectionObs(lCollection, lAuxiliar)
                lAuxiliar = "proposta de seguro ( marca e modelo ) e estarão sujeitos a análise de aceitação por parte da " & lNomeSeguradora & " Seguros."
                Call addCollectionObs(lCollection, lAuxiliar, True)
            End If
        End If
        '1101181 - Tarifa 1022 e 1023 Dezembro 2011
        If (mStcA46V708B.Cod_Iden_Tarif <= 1100) And _
           (mStcA46V708B.Cod_Ramo = e_CodRamo_Residencial) And _
           (lobjstcCobert.Cod_Cobert = mStcA46V708B.M46V999.gfBuscarDePara(5, 1)) And _
           (mStcA46V708B.ObjPai.objstcCtrlEmis.Cod_Corr = 309801) Then    '0309801 - HELLNER CORRETAGEM DE SEGUROS LTDA
            'CR27808 Ajuste a referencia nas condições gerais versão 02/2014
            'lAuxiliar = "Fica entendido e acordado que, ao contrário do que consta no item 4.2.15. ROUBO E FURTO COM VESTÍGIOS"
            lAuxiliar = "Fica entendido e acordado que, ao contrário do que consta no item 4.2.13. ROUBO E FURTO COM VESTÍGIOS"
            Call addCollectionObs(lCollection, lAuxiliar)
            lAuxiliar = "das Condições Gerais do produto " & lNomeSeguradora & " Residencial, em caso de sinistro coberto poderão ser indenizados valores"
            Call addCollectionObs(lCollection, lAuxiliar)
            lAuxiliar = "superiores aos limites de indenização previstos para roupas, calçados, cama, mesa e banho, fitas de vídeo e áudio, DVDs, CDs,"
            Call addCollectionObs(lCollection, lAuxiliar)
            lAuxiliar = "discos de vinil, mídias de gravação e livros, desde que conste a discriminação com respectivo valor individual em boletim de"
            Call addCollectionObs(lCollection, lAuxiliar)
            lAuxiliar = "ocorrência, e que o somatório dos prejuízos indenizáveis não ultrapasse o Limite Máximo de Garantia contratado para a cobertura."
            Call addCollectionObs(lCollection, lAuxiliar, True)
        End If
    Next lobjstcCobert


'Piloto
    If (mStcA46V708B.Cod_Iden_Tarif < 1080) Then
        If mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao And (mStcA46V708B.Tip_Cobert <> e_TipCobert_Predio And _
                                                               mStcA46V708B.gTemRegProfissoes = True) Then
                If mStcA46V708B.ObjPai.Cod_Ramo = e_CodRamo_Empresarial And mStcA46V708B.ObjPai.Tip_Pessoa = e_TipPessoa_Fisica And mStcA46V708B.Tip_Cobert <> e_TipCobert_Predio Then
                lAuxiliar = "Seguro de Pessoa Física: Na proposta será obrigatório informar o número do registro profissional no respectivo Conselho "
                Call addCollectionObs(lCollection, lAuxiliar)
                lAuxiliar = "da Categoria. Ex.: OAB, CRM, CRO etc. "
                Call addCollectionObs(lCollection, lAuxiliar, True)
                End If
        End If
    End If

    'Ficha 1000305
    If mStcA46V708B.Tip_Cobert = e_TipCobert_Conteudo Then
        lAuxiliar = "O presente seguro destina-se a cobrir exclusivamente o conteúdo do imóvel."
        Call addCollectionObs(lCollection, lAuxiliar, True)
    ElseIf mStcA46V708B.Tip_Cobert = e_TipCobert_Predio Then
        lAuxiliar = "O presente seguro destina-se a cobrir exclusivamente o prédio (imóvel)."
        Call addCollectionObs(lCollection, lAuxiliar, True)
    End If

    'Ficha 1100105
    If mStcA46V708B.ObjPai.ColecaoTextoEspec.Count > 0 Then
        For Each lobjstcTextoEspec In mStcA46V708B.ObjPai.ColecaoTextoEspec
            lPulaLinha = False
            If lobjstcTextoEspec.Cod_Espec = 6500 Then
                lPulaLinha = True
            End If
            
            If lobjstcTextoEspec.Cod_Espec <> 8500 Then
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin1) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin1)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin2) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin2)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin3) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin3, lPulaLinha)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin4) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin4)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin5) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin5)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin6) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin6)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin7) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin7)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin8) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin8, lPulaLinha)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin9) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin9, lPulaLinha)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin10) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin10, lPulaLinha)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin11) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin11)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin12) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin12, lPulaLinha)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin13) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin13)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin14) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin14, lPulaLinha)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin15) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin15)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin16) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin16, lPulaLinha)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin17) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin17)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin18) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin18, lPulaLinha)
                End If
                If mStcA46V708B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin19) Then
                    Call addCollectionObs(lCollection, lobjstcTextoEspec.Espec_lin19)
                End If
                
                Call addCollectionObs(lCollection, "", True)
            End If
        Next lobjstcTextoEspec
    End If

    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select

    '1200004 - Tarifa 1028 e 1029 Janeiro 2012
    Select Case mStcA46V708B.Cod_Plano
    Case 2229, 2230, 2231, 2232, 2233, 2234, 2235, 2236, 2237, 2238
        Call addCollectionObs(lCollection, "Seguro exclusivo para funcionários da Rinnai Brasil Tecnologia de Aquecimento Ltda.", True)
        Call addCollectionObs(lCollection, " ", True)
        Call addCollectionObs(lCollection, " ", True)
    End Select

    For Each lobjstcPedDizer In mStcA46V708B.ObjPai.ColecaoTextoEndo
        Call addCollectionObs(lCollection, lobjstcPedDizer.Dsc_Dizer_Endo)
    Next lobjstcPedDizer

    If mStcA46V708B.ObjPai.ColecaoTextoEndo.Count > 0 Then
        Call addCollectionObs(lCollection, " ", True)
    End If
    
    '11.44 Clausula de bicicletas quando Residencial Supremo e Cobertura de bicicleta
    If mStcA46V708B.Cod_Ramo = 113 And mStcA46V708B.Cod_Plano = 30 Then
        For Each lobjstcCobert In mStcA46V708B.ColecaoPedCobert
            If lobjstcCobert.Cod_Cobert = 1215 Then
                Call addCollectionObs(lCollection, "COBERTURA ADICIONAL BICICLETAS", True)
                Call addCollectionObs(lCollection, "1.  Esta cobertura garante, até o Limite Máximo de Indenização estabelecido nas Condições Particulares da apólice, a indenização por perdas e danos materiais diretamente causados a bicicleta descrita na apólice por: a) roubo; b) furto qualificado; e c) acidente ocorrido com o veículo transportador.", True)
                Call addCollectionObs(lCollection, "2.  Para efeito desta cobertura, entende-se por: Roubo: subtração do bem segurado (bicicleta) mediante emprego ou ameaça de violência, seja pela ação física ou assalto à mão armada, contra o Segurado. Esta cobertura somente será válida dentro do território nacional e quando a bicicleta estiver em poder do Segurado. Furto Qualificado: Subtração do bem segurado (bicicleta) mediante rompimento, destruição de obstáculos ou arrombamento, desde que qualquer desses meios tenha deixado vestígio. Esta cobertura somente será válida quando a bicicleta estiver dentro do local do risco mencionado na apólice (residência habitual do segurado). Acidente com o veículo Transportador: Dano parcial ou total causado ao objeto segurado (bicicleta) decorrente de um acidente com o veículo transportador. Esta cobertura somente será válida dentro do território nacional e quando o Segurado estiver neste veículo.", True)
                Call addCollectionObs(lCollection, "3.  Além das exclusões previstas na Cláusula 5ª – “Riscos Excluídos” das Condições Gerais, não estão cobertos os prejuízos decorrentes de: a) furto simples, ou seja, o desaparecimento inexplicável do bem segurado; b) infidelidade dos empregados ou prepostos do segurado.", True)
                Call addCollectionObs(lCollection, "4.  Além dos bens / interesses relacionados na Cláusula 6ª – “Bens / Interesses Não Garantidos” das Condições Gerais, não estão cobertos: a) o furto qualificado de bicicleta que esteja ao ar livre ou em edificações abertas ou semi-abertas; b) bicicletas quaisquer que não sejam de propriedade do Segurado, seu cônjuge ou dependentes legais; c) bicicletas de empregados; d) danos a quaisquer outros veículos que não sejam bicicletas.", True)
                Call addCollectionObs(lCollection, "5.  Em cada sinistro, o Segurado participará com o valor da franquia abaixo: ", True)
                Call addCollectionObs(lCollection, "    – Franquia para as Bicicletas:", True)
                Call addCollectionObs(lCollection, "      Bicicletas com Limite Máximo de Indenização até R$ 5.000,00", True)
                Call addCollectionObs(lCollection, "      Franquia de 10% dos prejuízos indenizáveis com mínimo de R$ 400,00", True)
                Call addCollectionObs(lCollection, "      Bicicletas com Limite Máximo de Indenização superior a R$ 5.000,00", True)
                Call addCollectionObs(lCollection, "      Franquia de 10% dos prejuízos indenizáveis com mínimo de R$ 800,00", True)
                Exit For
            End If
        Next lobjstcCobert
    End If
    
    '11.40 Observação especial Residencial Supremo - Assistências
    If mStcA46V708B.Cod_Ramo = 113 And mStcA46V708B.Cod_Plano = 30 Then
        Call addCollectionObs(lCollection, "Serviço de Proteção a Garagem (Central de Atendimento 0800 2839919)", True)
        Call addCollectionObs(lCollection, "Serviço de Descarte Ecológico (Central de Atendimento 0800 3262002)", True)
    End If
    
    Set lObjstcCorr = Nothing
    Set lobjstcPedDizer = Nothing
    Set lobjstcCobert = Nothing
    Set lobjstcTextoEspec = Nothing
    Set lobjstcCobert = Nothing
    
    Dim lobjstcV714 As stcA46V714B
    Dim strRetorno As String
    For Each lobjstcV714 In lCollection
        strRetorno = strRetorno + " " + lobjstcV714.Observacao + vbCrLf + vbCrLf
    Next lobjstcV714
    
    Dim lChave As String
    Dim lChaveAnt As String
    lChaveAnt = Empty
    
    For Each lobjstcNovoTexto In mStcA46V708B.ColecaoTextos
        lChave = CStr(lobjstcNovoTexto.nosso_numero) & _
                 CStr(Format(lobjstcNovoTexto.Num_Item, "00000")) & _
                 CStr(Format(lobjstcNovoTexto.Cod_Cobert, "0000"))
        
        If lChave = lChaveAnt Then
            strRetorno = strRetorno + " " + lobjstcNovoTexto.Desc_Texto + vbCrLf
        Else
        strRetorno = strRetorno + " " + lobjstcNovoTexto.Desc_Texto + vbCrLf + vbCrLf
        End If
    Next
    
    ObservacoesNew = strRetorno

FechaObj:
    Set lObjstcCorr = Nothing
    Set lobjstcCobert = Nothing
    Set lobjstcClau = Nothing
    Set lobjstcTextoEspec = Nothing
    Set lobjstcPedDizer = Nothing
    Set lobjstcNovoTexto = Nothing

End Property


Private Sub addCollectionObs(ByRef pCollection As Collection, ByVal Texto As String, Optional ByVal PulaLinha As Boolean = False)
    strTextoAdicional = strTextoAdicional + " " + Texto
    
    If PulaLinha Then
        Dim lobjstcV714 As stcA46V714B
        Set lobjstcV714 = New stcA46V714B
        Set lobjstcV714.ObjPai = mStcA46V708B
        lobjstcV714.Observacao = strTextoAdicional
        pCollection.Add lobjstcV714, mStcA46V708B.nosso_numero & pCollection.Count + 1
        strTextoAdicional = ""
    End If
    
End Sub


Public Property Get gFatoresDeProtecao() As Collection
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer
    Dim lCountLine As Integer
    Dim lbdTab_Tipo_Fator As ADODB.Connection
    Dim lMensagem As String
    Dim lrsTab_Tipo_Fator As clsYasRecordSet
    Dim lblnAchou As Boolean
    Dim lTemInfo As Boolean
    Dim lobjstcFatorRead As stcA46V720B       'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lobjstcFatorWrite As stcA46V720B        'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lCollection As Collection
    Dim lSelect As String

    Set lCollection = New Collection

    'Ficha 1000601
    If mStcA46V708B.Cod_Ramo = e_CodRamo_Residencial Then
        '1101045 - Tarifa 1022 e 1023 Novembro 2011
        '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 8
        'Retirado na ficha 1200733
        '        If (mStcA46V708B.Cod_Iden_Tarif >= 1022 And mStcA46V708B.Cod_Iden_Tarif <= 1033) Or _
                 '           (mStcA46V708B.Cod_Iden_Tarif >= 1038 And mStcA46V708B.Cod_Iden_Tarif <= 1043) Then
        '1201062 - Tarifa 1048 e 1049 - Item 1
        If mStcA46V708B.gLiberaDescontoEspecial = True Then
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & mStcA46V708B.Cod_Ramo
            lSelect = lSelect & " and Cod_Modalidade = " & mStcA46V708B.Cod_Produto
            lSelect = lSelect & " and COD_TARIF = " & mStcA46V708B.Cod_Iden_Tarif
            lSelect = lSelect & " and cod_Sequencia  <= " & mStcA46V708B.ObjPai.Cod_Sequencia
            lSelect = lSelect & " order by DSC_FATOR"
        Else
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & mStcA46V708B.Cod_Ramo
            lSelect = lSelect & " and Cod_Modalidade = " & mStcA46V708B.Cod_Produto
            lSelect = lSelect & " and COD_TARIF = " & mStcA46V708B.Cod_Iden_Tarif
            lSelect = lSelect & " and cod_Sequencia  <= " & mStcA46V708B.ObjPai.Cod_Sequencia
            lSelect = lSelect & " and Cod_Fator <> 7 "    'Não carregar o item Desconto Especial
            lSelect = lSelect & " order by DSC_FATOR"
        End If
    Else
        lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
        lSelect = lSelect & " Where cod_Ramo = " & mStcA46V708B.Cod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & mStcA46V708B.Cod_Produto
        lSelect = lSelect & " and COD_TARIF = " & mStcA46V708B.Cod_Iden_Tarif
        lSelect = lSelect & " and cod_Sequencia  <= " & mStcA46V708B.ObjPai.Cod_Sequencia
        lSelect = lSelect & " order by DSC_FATOR"
    End If

    If mStcA46V708B.M46V999.gfAbrBasDados("P00MULTI", mStcA46V708B.M46V999.gTipP00MULTI, mStcA46V708B.M46V999.gEndP00MULTI, lbdTab_Tipo_Fator, lMensagem) = False Then
        Call mStcA46V708B.M46V999.gpGraLog(2, lMensagem)
        Exit Property
    End If

    If mStcA46V708B.M46V999.gfObtRegistro(lbdTab_Tipo_Fator, lSelect, lrsTab_Tipo_Fator, lMensagem) = False Then
        Call mStcA46V708B.M46V999.gpFecha2(lbdTab_Tipo_Fator)
        Call mStcA46V708B.M46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Property
    End If

    lCountLine = 0
    Do While Not lrsTab_Tipo_Fator.EOF
        lblnAchou = False
        lCountLine = lCountLine + 1
        Set lobjstcFatorWrite = New stcA46V720B
        Set lobjstcFatorWrite.ObjPai = mStcA46V708B
        lobjstcFatorWrite.Cod_Fator = lrsTab_Tipo_Fator("Cod_Fator")
        lobjstcFatorWrite.Cod_Ramo = mStcA46V708B.Cod_Ramo
        For Each lobjstcFatorRead In mStcA46V708B.ColecaoPedFator
            If lobjstcFatorRead.Cod_Fator = lrsTab_Tipo_Fator("Cod_Fator") Then
                lTemInfo = True
                lblnAchou = True
                lAuxiliar2 = Format(lCountLine, "00") & " ( X ) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
                If Trim(lobjstcFatorRead.Dsc_Compl_Fator) <> "" Then
                    lAuxiliar2 = lAuxiliar2 & " - " & Trim(lobjstcFatorRead.Dsc_Compl_Fator)
                End If
                lobjstcFatorWrite.Desc_Fator = lAuxiliar2
                lCollection.Add lobjstcFatorWrite, mStcA46V708B.ObjPai.nosso_numero & lobjstcFatorWrite.Cod_Fator
                Exit For
            End If
        Next lobjstcFatorRead
        If lblnAchou = False Then
            lAuxiliar2 = Format(lCountLine, "00") & " (  ) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
            lobjstcFatorWrite.Desc_Fator = lAuxiliar2
            lCollection.Add lobjstcFatorWrite, mStcA46V708B.ObjPai.nosso_numero & lobjstcFatorWrite.Cod_Fator
        End If
        lrsTab_Tipo_Fator.MoveNext
    Loop
    Set lobjstcFatorRead = Nothing
    Set lobjstcFatorWrite = Nothing
    Call mStcA46V708B.M46V999.gpFecha2(lbdTab_Tipo_Fator)
    Set gFatoresDeProtecao = lCollection
End Property

Public Property Get gFatoresDeProtecaoNew() As String
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer
    Dim lCountLine As Integer
    Dim lbdTab_Tipo_Fator As ADODB.Connection
    Dim lMensagem As String
    Dim lrsTab_Tipo_Fator As clsYasRecordSet
    Dim lblnAchou As Boolean
    Dim lTemInfo As Boolean
    Dim lobjstcFatorRead As stcA46V720B       'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lobjstcFatorWrite As stcA46V720B        'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lCollection As Collection
    Dim lSelect As String

    Set lCollection = New Collection

    'Ficha 1000601
    If mStcA46V708B.Cod_Ramo = e_CodRamo_Residencial Then
        '1101045 - Tarifa 1022 e 1023 Novembro 2011
        '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 8
        'Retirado na ficha 1200733
        '        If (mStcA46V708B.Cod_Iden_Tarif >= 1022 And mStcA46V708B.Cod_Iden_Tarif <= 1033) Or _
                 '           (mStcA46V708B.Cod_Iden_Tarif >= 1038 And mStcA46V708B.Cod_Iden_Tarif <= 1043) Then
        '1201062 - Tarifa 1048 e 1049 - Item 1
        If mStcA46V708B.gLiberaDescontoEspecial = True Then
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & mStcA46V708B.Cod_Ramo
            lSelect = lSelect & " and Cod_Modalidade = " & mStcA46V708B.Cod_Produto
            lSelect = lSelect & " and COD_TARIF = " & mStcA46V708B.Cod_Iden_Tarif
            lSelect = lSelect & " and cod_Sequencia  <= " & mStcA46V708B.ObjPai.Cod_Sequencia
            lSelect = lSelect & " order by DSC_FATOR"
        Else
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & mStcA46V708B.Cod_Ramo
            lSelect = lSelect & " and Cod_Modalidade = " & mStcA46V708B.Cod_Produto
            lSelect = lSelect & " and COD_TARIF = " & mStcA46V708B.Cod_Iden_Tarif
            lSelect = lSelect & " and cod_Sequencia  <= " & mStcA46V708B.ObjPai.Cod_Sequencia
            lSelect = lSelect & " and Cod_Fator <> 7 "    'Não carregar o item Desconto Especial
            lSelect = lSelect & " order by DSC_FATOR"
        End If
    Else
        lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
        lSelect = lSelect & " Where cod_Ramo = " & mStcA46V708B.Cod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & mStcA46V708B.Cod_Produto
        lSelect = lSelect & " and COD_TARIF = " & mStcA46V708B.Cod_Iden_Tarif
        lSelect = lSelect & " and cod_Sequencia  <= " & mStcA46V708B.ObjPai.Cod_Sequencia
        lSelect = lSelect & " order by DSC_FATOR"
    End If

    If mStcA46V708B.M46V999.gfAbrBasDados("P00MULTI", mStcA46V708B.M46V999.gTipP00MULTI, mStcA46V708B.M46V999.gEndP00MULTI, lbdTab_Tipo_Fator, lMensagem) = False Then
        Call mStcA46V708B.M46V999.gpGraLog(2, lMensagem)
        Exit Property
    End If

    If mStcA46V708B.M46V999.gfObtRegistro(lbdTab_Tipo_Fator, lSelect, lrsTab_Tipo_Fator, lMensagem) = False Then
        Call mStcA46V708B.M46V999.gpFecha2(lbdTab_Tipo_Fator)
        Call mStcA46V708B.M46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Property
    End If

    lCountLine = 0
    Do While Not lrsTab_Tipo_Fator.EOF
        lblnAchou = False
        lCountLine = lCountLine + 1
        Set lobjstcFatorWrite = New stcA46V720B
        Set lobjstcFatorWrite.ObjPai = mStcA46V708B
        lobjstcFatorWrite.Cod_Fator = lrsTab_Tipo_Fator("Cod_Fator")
        lobjstcFatorWrite.Cod_Ramo = mStcA46V708B.Cod_Ramo
        For Each lobjstcFatorRead In mStcA46V708B.ColecaoPedFator
            If lobjstcFatorRead.Cod_Fator = lrsTab_Tipo_Fator("Cod_Fator") Then
                lTemInfo = True
                lblnAchou = True
                lAuxiliar2 = Format(lCountLine, "00") & " ( X ) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
                If Trim(lobjstcFatorRead.Dsc_Compl_Fator) <> "" Then
                    lAuxiliar2 = lAuxiliar2 & " - " & Trim(lobjstcFatorRead.Dsc_Compl_Fator)
                End If
                lobjstcFatorWrite.Desc_Fator = lAuxiliar2
                lCollection.Add lobjstcFatorWrite, mStcA46V708B.ObjPai.nosso_numero & lobjstcFatorWrite.Cod_Fator
                Exit For
            End If
        Next lobjstcFatorRead
        If lblnAchou = False Then
            lAuxiliar2 = Format(lCountLine, "00") & " (  ) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
            lobjstcFatorWrite.Desc_Fator = lAuxiliar2
            lCollection.Add lobjstcFatorWrite, mStcA46V708B.ObjPai.nosso_numero & lobjstcFatorWrite.Cod_Fator
        End If
        lrsTab_Tipo_Fator.MoveNext
    Loop
    Set lobjstcFatorRead = Nothing
    Set lobjstcFatorWrite = Nothing
    Call mStcA46V708B.M46V999.gpFecha2(lbdTab_Tipo_Fator)
    'Set gFatoresDeProtecao = lCollection
    Dim strRetorno As String
    
    For Each lobjstcFatorRead In lCollection
        strRetorno = strRetorno + lobjstcFatorRead.Desc_Fator + vbCrLf
    Next lobjstcFatorRead
    
    gFatoresDeProtecaoNew = strRetorno
    
End Property


Public Property Get Declaracoes() As Collection
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer
    Dim lColecao As Collection

    Set lColecao = New Collection

    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If mStcA46V708B.M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(mStcA46V708B.M46V999.gDataCorteSompo) Then
            If CLng(mStcA46V708B.M46V999.gDataCorteSompo) <= mStcA46V708B.M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If
    
    If mStcA46V708B.Cod_Ramo = e_CodRamo_Empresarial Or mStcA46V708B.Cod_Ramo = e_CodRamo_Condominio Then
        lAuxiliar2 = "Valor em risco declarado: R$ " & Format(mStcA46V708B.Val_Vr, "standard")
        Call addCollectionImp(lColecao, lAuxiliar2)
    End If

    lAuxiliar2 = "-Existem outros seguros cobrindo os mesmos bens? " & IIf(mStcA46V708B.ColecaoPedSeguro.Count > 0, "SIM", "NÃO")
    Call addCollectionImp(lColecao, lAuxiliar2)

    If mStcA46V708B.Cod_Ramo <> e_CodRamo_Condominio And mStcA46V708B.Cod_Iden_Tarif < 1102 Then
        lAuxiliar2 = "-Deseja a inclusão de Cláusula Beneficiária? "
        If mStcA46V708B.ObjPai.getRpt.Beneficiarios.Count > 0 Then
            lAuxiliar2 = lAuxiliar2 & "SIM"
        Else
            lAuxiliar2 = lAuxiliar2 & "NÃO"
        End If
        Call addCollectionImp(lColecao, lAuxiliar2)
    End If

    If (mStcA46V708B.Cod_Ramo = e_CodRamo_Residencial) And (mStcA46V708B.Cod_Iden_Tarif < 1022) Then
        'lAuxiliar2 = "Existem bens no interior da residência com valor unitário superior a R$ 5.000,00? "
        lAuxiliar2 = "-RELAÇÃO DE BENS SEGURADOS : Objetos com valores acima dos Limites de Indenização estabelecidos na"
        Call addCollectionImp(lColecao, lAuxiliar2)

        lAuxiliar2 = "Cláusula 4 das Condições Gerais do produto " & lNomeSeguradora & " Residencial, no que tange a Cobertura de Roubo e Furto"
        Call addCollectionImp(lColecao, lAuxiliar2)

        lAuxiliar2 = "Com Vestígios, deverão estar discriminados na proposta de seguro ( marca e modelo ) e estarão sujeitos a"
        Call addCollectionImp(lColecao, lAuxiliar2)

        lAuxiliar2 = "análise de aceitação por parte da " & lNomeSeguradora & " Marítima Seguros."
        Call addCollectionImp(lColecao, lAuxiliar2)
    End If

    If (mStcA46V708B.Cod_Ramo <> 112 And mStcA46V708B.Cod_Iden_Tarif < 1080) Or (mStcA46V708B.Cod_Ramo = 113 And mStcA46V708B.Cod_Iden_Tarif >= 1080) Then
        lAuxiliar2 = "-Número de empregados na residência: " & mStcA46V708B.Qtd_Empreg
        Call addCollectionImp(lColecao, lAuxiliar2)
    End If

    Set Declaracoes = lColecao
SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lColecao = Nothing
    Return
End Property

Private Sub addCollectionImp(ByRef pCollection As Collection, ByVal Texto As String)
    Dim lobjstcTextoImpressao As stcTextoImpressao
    Set lobjstcTextoImpressao = New stcTextoImpressao
    lobjstcTextoImpressao.Sequencia = pCollection.Count + 1
    lobjstcTextoImpressao.Descricao = Texto
    Call pCollection.Add(lobjstcTextoImpressao, mStcA46V708B.nosso_numero & pCollection.Count + 1)
End Sub

Public Property Get LocalRisco() As String
    Dim lretorno As String
    lretorno = lretorno & Trim(mStcA46V708B.Cod_Tip_Logr_Risco) & " "
    lretorno = lretorno & Trim(Left(mStcA46V708B.Nom_Logr_Risco, 30)) & ",  "
    lretorno = lretorno & Trim(mStcA46V708B.objstcPedInsp.Num_Logr_Risco)
    lretorno = lretorno & IIf(Trim(mStcA46V708B.Dsc_Compl_Risco) = "", "", " - " & Trim(mStcA46V708B.Dsc_Compl_Risco)) & " "
    lretorno = lretorno & Mid(Trim(mStcA46V708B.Nom_Bairro_Risco), 1, 15) & " - "
    lretorno = lretorno & Trim(mStcA46V708B.Nom_Cid_Risco) & " "
    lretorno = lretorno & Trim(mStcA46V708B.Sig_UF_Risco) & " "
    lretorno = lretorno & Format(mStcA46V708B.Num_Cep_Risco, "00000-000")

    LocalRisco = lretorno
End Property

Public Property Get Rubrica() As String
    Dim lretorno As String
    'lRetorno = Val(Replace(mStcA46V708B.Cod_Clasf, "A", "999") & mStcA46V708B.Cod_Compl_Clasf)
    lretorno = mStcA46V708B.Cod_Clasf & "-" & mStcA46V708B.Cod_Compl_Clasf & "-" & mStcA46V708B.gNom_Class
    Rubrica = lretorno
End Property


Public Property Get descricao_Ass24h() As String
    Dim lretorno As String
    If Len(Trim(mStcA46V708B.gDescricao_Ass24h)) > 0 Then
        lretorno = "Assistência 24hs: " & mStcA46V708B.gDescricao_Ass24h
    Else
        If mStcA46V708B.ObjPai.Cod_Produto >= 1080 Then
            If mStcA46V708B.ColecaoPedClau.Count > 0 Then
                lretorno = "Assistência 24hs: " & mStcA46V708B.ColecaoPedClau(1).Descricao
            Else
                lretorno = "Assistência 24hs: Não se aplica"
            End If
        Else
            lretorno = "Assistência 24hs: Não se aplica"
        End If
    End If
    descricao_Ass24h = lretorno
End Property

Public Property Get InspNumTel2() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.objstcPedInsp.NUM_TEL_2
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    InspNumTel2 = lretorno
End Property
Public Property Get InspNumTel3() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.objstcPedInsp.NUM_TEL_3
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    InspNumTel3 = lretorno
End Property

Public Property Get InspPessoaContato2() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.objstcPedInsp.PESS_CONTATO_2
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    InspPessoaContato2 = lretorno
End Property

Public Property Get InspPessoaContato3() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.objstcPedInsp.PESS_CONTATO_3
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    InspPessoaContato3 = lretorno
End Property

Public Property Get InspDscEmail() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.objstcPedInsp.DSC_EMAIL
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    InspDscEmail = lretorno
End Property


Public Property Get InspDscPontoRef() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V708B.objstcPedInsp.DSC_PONTO_REF
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    InspDscPontoRef = lretorno
End Property
