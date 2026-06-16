VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V702rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStcA46V702B As stcA46V702B

Public Enum eTipCalcRPT
    eTipCalcRPT_ND = 0
    eTipCalcRPT_CotacIndividual = 1
    eTipCalcRPT_CotacColetiva = 2
    eTipCalcRPT_PropIndividual = 3
    eTipCalcRPT_PropColetiva = 4
    eTipCalcRPT_CotacColetivaLMI = 5
    eTipCalcRPT_PropColetivaLMI = 6
End Enum

Private Enum eTipoRetornoTxtExpec
    eTipoRetornoTxtExpec_Beneficiarios = 1
    eTipoRetornoTxtExpec_Franquia = 2
End Enum

Private Sub Class_Initialize()
    Set mStcA46V702B = Nothing
End Sub

Private Sub Class_Terminate()
    Set mStcA46V702B = Nothing
End Sub

Public Property Set stcOrigem(ByRef pStcA46V702B As stcA46V702B)
    Set mStcA46V702B = pStcA46V702B
End Property

Public Property Get DescRamo() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    lretorno = UCase(mStcA46V702B.gDesc_Ramo)
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    DescRamo = lretorno
End Property

Public Property Get TituloSeguro() As String
    Dim lretorno As String
    Dim lNomeSeguradora As String
    
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado YASUDA RESIDENCIAL, YASUDA CONDOMÍNIO ou YASUDA EMPRESARIAL.
    If mStcA46V702B.getPlano.EhPlanoSegmentado Then
        Select Case mStcA46V702B.Cod_Ramo
            Case 112
                lretorno = "Empresarial " & mStcA46V702B.getPlano.NomPlano
            Case 113
                lretorno = "Residencial " & mStcA46V702B.getPlano.NomPlano
            Case 114
                lretorno = "Condomínio " & mStcA46V702B.getPlano.NomPlano
        End Select
    Else
        ' Obter o nome da Seguradora
        lNomeSeguradora = "YASUDA"
        If mStcA46V702B.M46V999.gDataCorteSompo <> "" Then
            If IsNumeric(mStcA46V702B.M46V999.gDataCorteSompo) Then
                If CLng(mStcA46V702B.M46V999.gDataCorteSompo) <= mStcA46V702B.M46V999.gHoje Then
                    lNomeSeguradora = "SOMPO"
                End If
            End If
        End If
    
    
        lretorno = lNomeSeguradora & " " & UCase(mStcA46V702B.gDesc_Ramo)
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    TituloSeguro = lretorno
End Property

Public Property Get ProcessoSusep() As String
    Dim lretorno As String
    Dim ProcessoBD As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    '[FRH] - FICHA 901116
    '001B. Código do Processo Susep.
    '     Observar que este número deve ser de acordo com o título do seguro:
    '     RESIDENCIAL, CONDOMÍNIO ou EMPRESARIAL.
    Select Case mStcA46V702B.Cod_Ramo
    Case e_CodRamo_Condominio
        lretorno = "Processo Susep 15414.004604/2004-72"
    Case e_CodRamo_Empresarial
        lretorno = "Processo Susep 15414.004605/2004-17"
    Case e_CodRamo_Residencial
        lretorno = "Processo Susep 15414.003659/2004-65"
    Case Else
        lretorno = " "
    End Select
    
    ProcessoBD = mStcA46V702B.gCodSusep
    
    If mStcA46V702B.M46V999.gfPreenchido(ProcessoBD) Then
        lretorno = "Processo Susep " & ProcessoBD
    End If
    
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    ProcessoSusep = lretorno

End Property

Public Property Get Emissao() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format$(Now, "dd/mm/yyyy hh:nn:ss")
    If mStcA46V702B.Cod_Tip_Doc = e_TipDoc_Proposta Then
        If mStcA46V702B.Num_PI_Emissao <> 0 Then
            lretorno = lretorno & Space(20) & "PI: " & mStcA46V702B.Num_PI_Emissao & "-" & mStcA46V702B.Dig_PI_Emissao
        Else
            lretorno = lretorno & Space(20) & "PROPOSTA NÃO TRANSMITIDA"
        End If
    End If
    If mStcA46V702B.Ind_Reentrada = 1 Then
        lretorno = lretorno & "**REENTRADA**"
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Emissao = lretorno
End Property

Public Property Get Validade() As String
    Dim lretorno As String
    Dim lPedLoc As stcA46V708B
    Dim lValidade As Integer
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    Set lPedLoc = mStcA46V702B.ColecaoPedLoc(1)
    If lPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Or lPedLoc.Ren_Con_Seguradora = 572 Then 'renovação yasuda e maritima
        If lPedLoc.Cod_Iden_Tarif >= 978 And lPedLoc.Cod_Iden_Tarif <= 1047 Then
            lValidade = 15
            'lretorno = Format$(CDate(mStcA46V702B.M46V999.gfForData(2, mStcA46V702B.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
        Else
            lValidade = 30
            'lretorno = Format$(CDate(mStcA46V702B.M46V999.gfForData(2, mStcA46V702B.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
        End If
    Else
        lValidade = 15
        'lretorno = Format$(CDate(mStcA46V702B.M46V999.gfForData(2, mStcA46V702B.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
    End If
    
    If mStcA46V702B.M46V999.gValidadeCotCorr > 0 Then
        lValidade = mStcA46V702B.M46V999.gValidadeCotCorr
    End If
    
    '11.46 - Empresarial item 4
    'Tranferencia da regra de 917581 - Triunfo para 939571 - BR Insurance - 1174
    If lPedLoc.Cod_Ramo = e_CodRamo_Condominio And mStcA46V702B.M46V999.gCod_Corretor = 939571 And (lPedLoc.Tip_Emissao = e_TipEmissao_RenCongenere Or lPedLoc.Tip_Emissao = e_TipEmissao_SeguroNovo) Then
        lValidade = 30
    End If
    
    lretorno = Format$(CDate(mStcA46V702B.M46V999.gfForData(2, mStcA46V702B.Dat_Pri_Calculo)) + lValidade, "dd/mm/yyyy")
    
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Set lPedLoc = Nothing
    Validade = lretorno
End Property

Public Property Get Siscota() As String
    Dim lretorno As String

    If val(mStcA46V702B.objstcCtrlEmis.Tip_Cotac) = 2 And val(mStcA46V702B.objstcCtrlEmis.Num_Cotac) <> 0 Then
        lretorno = mStcA46V702B.objstcCtrlEmis.Num_Cotac
    Else
        lretorno = " "
    End If
    lretorno = IIf(lretorno = "", " ", lretorno)
    Siscota = lretorno
End Property

Public Property Get CorretorCod() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V702B.ColecaoPedCorr.Count > 0 Then
        lretorno = mStcA46V702B.ColecaoPedCorr.Item(1).Cod_Corr
    Else
        lretorno = 0
    End If
    lretorno = Format$(lretorno, "0000")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    CorretorCod = lretorno
End Property

Public Property Get CorretorNome() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V702B.ColecaoPedCorr.Count > 0 Then
        lretorno = mStcA46V702B.ColecaoPedCorr.Item(1).gNom_Corretor
    Else
        lretorno = " "
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    CorretorNome = lretorno
End Property


Public Property Get CorretorEmail() As String
    Dim lBDP00Multi As ADODB.Connection
    Dim lMensagem As String
    Dim lrst As clsYasRecordSet
    Dim lSelect As String

    If mStcA46V702B.M46V999.gfAbrBasDados("P00MULTI", mStcA46V702B.M46V999.gTipP00MULTI, mStcA46V702B.M46V999.gEndP00MULTI, _
                                          lBDP00Multi, lMensagem) = False Then
        Call mStcA46V702B.M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If
    
    lSelect = "SELECT  email_corretor "
    lSelect = lSelect & "FROM ged..tbs_corretor "
    lSelect = lSelect & "WHERE cod_corretor = " & Me.CorretorCod
    
    lMensagem = Empty
    If mStcA46V702B.M46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
        GoTo SAIDA
    End If
    
    If Not lrst.EOF Then
        CorretorEmail = "" & lrst!email_corretor
    End If

SAIDA:
    GoSub FechaObj
    If lMensagem <> Empty Then
        Err.Raise 999, , lMensagem
    End If
    Exit Property
FechaObj:
    Call mStcA46V702B.M46V999.gpFecha2(lBDP00Multi)
    Set lBDP00Multi = Nothing
    Set lrst = Nothing
    Return
End Property

Public Property Get CorretorTelefone() As String
    Dim lBDP00Multi As ADODB.Connection
    Dim lMensagem As String
    Dim lrst As clsYasRecordSet
    Dim lSelect As String

    If mStcA46V702B.M46V999.gfAbrBasDados("P00MULTI", mStcA46V702B.M46V999.gTipP00MULTI, mStcA46V702B.M46V999.gEndP00MULTI, _
                                          lBDP00Multi, lMensagem) = False Then
        Call mStcA46V702B.M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If
    
    lSelect = "SELECT  num_telefone "
    lSelect = lSelect & "FROM ged..tbs_corretor "
    lSelect = lSelect & "WHERE cod_corretor = " & Me.CorretorCod
    
    lMensagem = Empty
    If mStcA46V702B.M46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
        GoTo SAIDA
    End If
    
    If Not lrst.EOF Then
        CorretorTelefone = "" & lrst!num_telefone
    End If

SAIDA:
    GoSub FechaObj
    If lMensagem <> Empty Then
        Err.Raise 999, , lMensagem
    End If
    Exit Property
FechaObj:
    Call mStcA46V702B.M46V999.gpFecha2(lBDP00Multi)
    Set lBDP00Multi = Nothing
    Set lrst = Nothing
    Return
End Property

Public Property Get ComissaoEDescontoLMI() As String
    Dim lretorno As String
    Dim lObjstcCorr As stcA46V705B
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    
    If mStcA46V702B.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = mStcA46V702B.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Per_Corr
        If mStcA46V702B.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = mStcA46V702B.ColecaoPedCorr.Item(2)
            lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            If mStcA46V702B.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = mStcA46V702B.ColecaoPedCorr.Item(3)
                lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            End If
        End If
    End If
    lAuxiliar2 = lAuxiliar2 & Format(lAuxiliar1 * 100, "0000")
    ComissaoEDescontoLMI = lAuxiliar2 & Format(mStcA46V702B.objstcCtrlEmis.Per_Desc_Coml * 100, "0000")
End Property

Public Property Get Referencia() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format$(mStcA46V702B.Num_Ped, "000000")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Referencia = lretorno
End Property

Public Property Get Tarifa() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.gDescTarifa
    If mStcA46V702B.Cod_Sequencia = 1 Then
        lretorno = lretorno & " (A)"
    End If
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Tarifa = lretorno
End Property

Public Property Get NossoNumero() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.nosso_numero
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    NossoNumero = lretorno
End Property


Public Property Get DV() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = "DV: " & mStcA46V702B.Cod_Unicidade & " - " & mStcA46V702B.Cod_Ver_Unici
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    DV = lretorno
End Property


Public Property Get NomSegurado() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.Nom_Segurado
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    NomSegurado = lretorno
End Property

Public Property Get CodSegurado() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format$(mStcA46V702B.Cod_Segurado, "0000000")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    CodSegurado = lretorno
End Property

Public Property Get TipoPessoa() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    Select Case mStcA46V702B.Tip_Pessoa
    Case e_TipPessoa_Juridica
        lretorno = "Jurídica"
    Case e_TipPessoa_Fisica
        lretorno = "Física"
    End Select
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    TipoPessoa = lretorno
End Property

Public Property Get CpfCnpj() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.gCNPJ_CPF_Formatado
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    CpfCnpj = lretorno
End Property

Public Property Get Identificacao1() As String
    Dim lretorno As String
    Select Case tipoRelatorio
    Case eTipCalcRPT_CotacIndividual
        Select Case mStcA46V702B.Tip_Pessoa
        Case e_TipPessoa_Juridica
            lretorno = " "
        Case e_TipPessoa_Fisica
            lretorno = ""
            'Ficha 1000307
            If mStcA46V702B.M46V999.gfPreenchido(mStcA46V702B.objstcPedEndereco.gDesc_Reg) Then
                lretorno = lretorno & "       Cat.Registro: " & mStcA46V702B.objstcPedEndereco.gDesc_Reg
                lretorno = lretorno & "        " & "Número Registro: " & mStcA46V702B.objstcPedEndereco.Num_Reg_Cons
            End If
        End Select
    Case eTipCalcRPT_PropIndividual
        Select Case mStcA46V702B.Tip_Pessoa
        Case e_TipPessoa_Juridica
            lretorno = "Ramo de atividade: " & IIf(mStcA46V702B.gDesc_RamoAtiv = "", "Não Informado", mStcA46V702B.gDesc_RamoAtiv)                                '1100705 - Tarifa 1016 e 1017 Agosto 2011 - Item 1 não obrigar informar Ramo de Atividade
        Case e_TipPessoa_Fisica
            lretorno = mStcA46V702B.gCod_Doc_Segur & ": " & Trim(mStcA46V702B.Num_Doc) & "   " & _
                       "Órgão emissor: " & Trim(mStcA46V702B.Org_Exp)
        End Select
        '        Case eTipCalcRPT_CotacColetiva
        '        Case eTipCalcRPT_PropColetiva
    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Identificacao1 = lretorno
End Property

Public Property Get Identificacao2() As String
    Dim lretorno As String
    Select Case tipoRelatorio
    Case eTipCalcRPT_CotacIndividual
        Select Case mStcA46V702B.Tip_Pessoa
        Case e_TipPessoa_Juridica
            lretorno = "Ramo de atividade: " & mStcA46V702B.gDesc_RamoAtiv
        Case e_TipPessoa_Fisica
            lretorno = mStcA46V702B.gCod_Doc_Segur & ": " & mStcA46V702B.Num_Doc & "   " & _
                       "Órgão emissor: " & mStcA46V702B.Org_Exp & "   " & _
                       "Emissão: " & mStcA46V702B.M46V999.gfForData(2, mStcA46V702B.Dat_Emis_Doc)
        End Select
    Case eTipCalcRPT_PropIndividual
        Select Case mStcA46V702B.Tip_Pessoa
        Case e_TipPessoa_Juridica
            If mStcA46V702B.Val_RendaFatu > 0 Then
                lretorno = " Faturamento: " & Format(mStcA46V702B.Val_RendaFatu, "standard")
            Else
                lretorno = " Faturamento: não informado"
            End If
            'Fichas 1200168 e 1200169 Inclusão dos campos de listagem Renda e Faixa de Faturamento Pessoa física e jurídica
            lretorno = " Faturamento: " & mStcA46V702B.gDesc_Renda_Faturamento
        Case e_TipPessoa_Fisica
            lretorno = "Emissão: " & mStcA46V702B.M46V999.gfForData(2, mStcA46V702B.Dat_Emis_Doc)
            '                    If mStcA46V702B.Val_RendaFatu > 0 Then
            '                        lRetorno = lRetorno & " Renda: " & Format(mStcA46V702B.Val_RendaFatu, "standard")
            '                    Else
            '                        lRetorno = lRetorno & " Renda: não informada "
            '                    End If
            'Fichas 1200168 e 1200169 Inclusão dos campos de listagem Renda e Faixa de Faturamento Pessoa física e jurídica
            lretorno = lretorno & " Renda: " & mStcA46V702B.gDesc_Renda_Faturamento
            'Ficha 1000307
            If mStcA46V702B.M46V999.gfPreenchido(mStcA46V702B.objstcPedEndereco.gDesc_Reg) Then
                lretorno = lretorno & "       Cat.Registro: " & mStcA46V702B.objstcPedEndereco.gDesc_Reg
                lretorno = lretorno & "        " & "Número Registro: " & mStcA46V702B.objstcPedEndereco.Num_Reg_Cons
            End If
        End Select
    Case eTipCalcRPT_CotacColetiva
    Case eTipCalcRPT_PropColetiva
    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Identificacao2 = lretorno
End Property

Public Property Get Vigencia() As String
    Dim lretorno As String
    Dim lQuaDias As Integer
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = "de " & mStcA46V702B.M46V999.gfForData(2, mStcA46V702B.Dat_Ini_Vig) & _
             " a " & mStcA46V702B.M46V999.gfForData(2, mStcA46V702B.Dat_Ter_Vig)
    'Se for M99V001.Perfil Calculo não será impresso conforme ficha 0701831
    If mStcA46V702B.M46V999.gTipUsuario = e_Tipo_Usu_Yasuda And _
       (mStcA46V702B.M46V999.PerfilAcesso = e_Perfil_CALC Or mStcA46V702B.M46V999.PerfilAcesso = e_Perfil_CALC_Gestor) Then
        lQuaDias = mStcA46V702B.M46V999.gfDifDatas(mStcA46V702B.Dat_Ini_Vig, mStcA46V702B.Dat_Ter_Vig)
        If lQuaDias >= 365 Then
            lretorno = ""
        End If
    End If
    lretorno = IIf(lretorno = "", " ", lretorno)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Vigencia = lretorno
End Property

Public Property Get UsuarioLiberacao() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.Cod_User_Liber & " - " & mStcA46V702B.M46V999.gfForData(2, mStcA46V702B.Dat_Liber) & " - " & Mid(Format(mStcA46V702B.Hor_Liber, "0000"), 1, 2) & ":" & Mid(Format(mStcA46V702B.Hor_Liber, "0000"), 3, 2)
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    UsuarioLiberacao = lretorno
End Property

Public Property Get Qtd_Parc() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.Qtd_Parc
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Qtd_Parc = lretorno
End Property

Public Property Get Cod_Iden_Pag_Antec() As String
    Dim lretorno As String
    lretorno = IIf(mStcA46V702B.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1, "SIM", "NÃO")

    If (lretorno = "NÃO") Then
        If mStcA46V702B.Cod_Sit_Ped = e_SitPedido_Transmitido Then
            lretorno = IIf((mStcA46V702B.objstcCtrlEmis.Cod_Bco_Boleto <> 0 And mStcA46V702B.Cod_Bco_Cobr <> 0), "SIM", "NÃO")
        End If
    End If

    Cod_Iden_Pag_Antec = lretorno
End Property

Public Property Get Num_Bloq() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V702B.objstcCtrlEmis.Num_Bloq <> 0 Then
        lretorno = mStcA46V702B.objstcCtrlEmis.Num_Bloq
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Num_Bloq = lretorno
End Property

Public Property Get DIG_BLOQUETO() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V702B.objstcCtrlEmis.Num_Bloq <> 0 Then
        lretorno = mStcA46V702B.objstcCtrlEmis.DIG_BLOQUETO
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    DIG_BLOQUETO = lretorno
End Property
Public Property Get DescricaoPremio() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V702B.ValTotalComissao = 0 Then
        lretorno = "Prêmio NET R$ :"
    Else

        If mStcA46V702B.Cod_Reg > 100 Then
            If mStcA46V702B.VAL_PR = mStcA46V702B.M46V999.gProPreMinEndosso Then
                lretorno = "Prêmio Mínimo R$ :"
            Else
                lretorno = "Prêmio líquido R$ :"
            End If
        Else
'            ' Daniel Ribeiro - Verificar se o valor do Premio Liquido total é maior
'            ' do que a soma dos premios liquidos de cada item. Se for, é porque caiu na regra
'            ' de premio mínimo e a mensagem será diferente.
'            Dim vlPremioLiquidoTotal As Double
'            Dim vlPremioLiquidoItens As Double
'            Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
'            Dim lobjstcCob As stcA46V704B
'
'            vlPremioLiquidoTotal = mStcA46V702B.gValPrLiqTotal
'            vlPremioLiquidoItens = 0
'
'            For Each lobjstcPedLoc In mStcA46V702B.ColecaoPedLoc
'                For Each lobjstcCob In lobjstcPedLoc.ColecaoPedCobert
'                    vlPremioLiquidoItens = vlPremioLiquidoItens + lobjstcCob.VAL_PR
'                Next
'            Next
'
'            If vlPremioLiquidoTotal > vlPremioLiquidoItens Then
'                lretorno = "Prêmio Mínimo   R$ :"
'            Else
                lretorno = "Prêmio Líquido  R$ :"
'            End If
        End If
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    DescricaoPremio = lretorno
End Property

Public Property Get ValPrLiqTotal() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format$(mStcA46V702B.gValPrLiqTotal, "standard")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    ValPrLiqTotal = lretorno
End Property

Public Property Get ValCusEmis() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format$(mStcA46V702B.Val_Cust_Emis, "standard")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    ValCusEmis = lretorno
End Property

Public Property Get ValIOF() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format$(mStcA46V702B.Val_IOF, "standard")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    ValIOF = lretorno
End Property

Public Property Get ValPremioTotal() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = Format$(mStcA46V702B.gValorTotal, "standard")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    ValPremioTotal = lretorno
End Property


Public Property Get Cosseguro() As String
    Dim lretorno As String
    Dim lobjstcCoss As stcA46V706B    'Classe (stc) de cosseguro.
    Dim gM46V999 As clsM46V999
    
    Dim ConexaoBD As ADODB.Connection
    Dim lMensagem As String
    Dim lrst As clsYasRecordSet
    Dim lSelect As String
    Dim UltimoErro As String

    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    
    '8000046585, Cosseguradora na impressão
    Set gM46V999 = mStcA46V702B.M46V999
       
    Dim lPerLider As Double

    lPerLider = 0
 
    lretorno = ""
    For Each lobjstcCoss In mStcA46V702B.ColecaoPedCoss
        
        If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, ConexaoBD, UltimoErro) = False Then
            Call gM46V999.gpGraLog(2, UltimoErro)
            Exit Sub
        End If
        
        lSelect = "SELECT CSeguradora, Seguradora FROM GED.DBO.TAB_SEGURADORA A WHERE A.CSeguradora = " + CStr(lobjstcCoss.Cod_Coss)
        If gM46V999.gfObtRegistro(ConexaoBD, lSelect, lrst, lMensagem) = False Then
            Call gM46V999.gpFecha2(ConexaoBD)
            Exit Sub
        End If
        
        Dim strNomeSeguradora As String
        
        If Not lrst.EOF Then
            strNomeSeguradora = Trim(lrst("Seguradora"))
        End If
        
        Dim strSaidaCnpj As String
        
        Call gM46V999.gpForCNPJ(1, (gM46V999.gfObterCnpjCoss(lobjstcCoss.Cod_Coss)), strSaidaCnpj)
        
        lretorno = lretorno & lobjstcCoss.Cod_Coss & " - " & strNomeSeguradora & " - " & IIf(Trim(gM46V999.gfObterCnpjCoss(lobjstcCoss.Cod_Coss)) = Empty, "", strSaidaCnpj & " - ") & lobjstcCoss.Per_Coss & "%" & vbCrLf
        
        lPerLider = lPerLider + lobjstcCoss.Per_Coss
        
    Next lobjstcCoss
    
    Dim strlretornoLider As String
    
    If mStcA46V702B.ColecaoPedCoss.Count > 0 Then
    
        strlretornoLider = "572 - SOMPO SEGUROS S/A (LÍDER) - 61.383.493/0001-80 - " & CStr(100 - lPerLider) & "%"
        
        lretorno = strlretornoLider + vbCrLf + lretorno
        
        '        Case eTipCalcRPT_PropColetiva
        '    End Select
    
    End If
    
    lretorno = IIf(lretorno = "", " ", lretorno)
    Cosseguro = lretorno
End Property

Public Property Get UsoInterno_Unidade() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V702B.Cod_Unid_Prod < 1000 Then
        lretorno = "Unidade: " & "MATRIZ" & Space(16)
    Else
        lretorno = "Unidade: " & mStcA46V702B.Cod_Unid_Prod & Space(16)
    End If

    'Ficha: 1200736
    '039. Corretor .
    Dim lObjstcCorr As stcA46V705B    'Classe (stc) de corretor.
    Dim lAuxiliar_01 As String
    Dim lAuxiliar_02 As String
    Dim lAuxiliar_03 As String
    lAuxiliar_01 = "0000"
    lAuxiliar_02 = "0000"
    lAuxiliar_03 = "0000"

    If mStcA46V702B.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = mStcA46V702B.ColecaoPedCorr.Item(1)
        lAuxiliar_01 = Format(lObjstcCorr.Per_Corr * 100, "0000")
        If mStcA46V702B.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = mStcA46V702B.ColecaoPedCorr.Item(2)
            lAuxiliar_02 = Format(lObjstcCorr.Per_Corr * 100, "0000")
            If mStcA46V702B.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = mStcA46V702B.ColecaoPedCorr.Item(3)
                lAuxiliar_03 = Format(lObjstcCorr.Per_Corr * 100, "0000")
            End If
        End If
    End If
    lretorno = lretorno & lAuxiliar_01 & "." & lAuxiliar_02 & "." & lAuxiliar_03

    '038. Produtor .
    If mStcA46V702B.Cod_Unid_Prod < 1000 Then
        lretorno = lretorno & " Produtor: "
    Else
        lretorno = lretorno & " Produtor: " & mStcA46V702B.Cod_Prod
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    Set lObjstcCorr = Nothing
    lretorno = IIf(lretorno = "", " ", lretorno)
    UsoInterno_Unidade = lretorno
End Property

Public Property Get UsoInterno_Corretor() As String
    Dim lretorno As String
    Dim lObjstcCorr As stcA46V705B    'Classe (stc) de corretor.
    Dim lAuxiliar_01 As String
    Dim lAuxiliar_02 As String
    Dim lAuxiliar_03 As String

    Exit Property
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V702B.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = mStcA46V702B.ColecaoPedCorr.Item(1)
        lAuxiliar_01 = Format(lObjstcCorr.Cod_Corr, "0000000")
        lAuxiliar_01 = lAuxiliar_01 & "/" & Format(lObjstcCorr.Per_Corr * 100, "0000")
        If mStcA46V702B.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = mStcA46V702B.ColecaoPedCorr.Item(2)
            lAuxiliar_02 = Format(lObjstcCorr.Cod_Corr, "0000000")
            lAuxiliar_02 = lAuxiliar_02 & "/" & Format(lObjstcCorr.Per_Corr * 100, "0000")
            If mStcA46V702B.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = mStcA46V702B.ColecaoPedCorr.Item(3)
                lAuxiliar_03 = Format(lObjstcCorr.Cod_Corr, "0000000")
                lAuxiliar_03 = lAuxiliar_03 & "/" & Format(lObjstcCorr.Per_Corr * 100, "0000")
            Else
                lAuxiliar_03 = " "
            End If
        Else
            lAuxiliar_02 = " "
            lAuxiliar_03 = " "
        End If
    Else
        lAuxiliar_01 = " "
        lAuxiliar_02 = " "
        lAuxiliar_03 = " "
    End If
    lretorno = "Corretor: " & lAuxiliar_01 & Space(2) & lAuxiliar_02 & Space(2) & lAuxiliar_03
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    Set lObjstcCorr = Nothing
    lretorno = IIf(lretorno = "", " ", lretorno)
    UsoInterno_Corretor = lretorno
End Property

Public Property Get Ped_Endereco() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.objstcPedEndereco.Cod_Tip_Logr & " " & _
               mStcA46V702B.objstcPedEndereco.Nom_Logr & ", " & _
               mStcA46V702B.objstcPedEndereco.Dsc_Compl_Logr
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Ped_Endereco = lretorno
End Property

Public Property Get Ped_CEP() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.objstcPedEndereco.gCEP_Formatado
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Ped_CEP = lretorno
End Property

Public Property Get Ped_Fone() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.objstcPedEndereco.Num_Fone
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Ped_Fone = lretorno
End Property

Public Property Get Ped_Bairro() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.objstcPedEndereco.Nom_Bairro
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Ped_Bairro = lretorno
End Property

Public Property Get Ped_Cidade() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.objstcPedEndereco.Nom_Cid
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Ped_Cidade = lretorno
End Property

Public Property Get Ped_UF() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.objstcPedEndereco.Sig_UF
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Ped_UF = lretorno
End Property

Public Property Get Profissao() As String
    Dim lretorno As String
    If Not (mStcA46V702B.M46V999.gPrograma = "D046O003") Then
        lretorno = "Profissão: "
    End If
    lretorno = lretorno & mStcA46V702B.gDescProfissao

    lretorno = IIf(lretorno = "", " ", lretorno)
    Profissao = lretorno
End Property

Public Property Get Per_Juros_Neg() As String
    Dim lretorno As String
    Dim lobjstcPedJur As stcA46V711B
    Dim Qtd_Parc_Ant As Integer
    Dim Qtd_Parc_Pos As Integer
    Dim Qtd_Parc As Integer

    If mStcA46V702B.Cod_Produto < mStcA46V702B.M46V999.TarifaPrecificacaoMaritima Then
        For Each lobjstcPedJur In mStcA46V702B.ColecaoPedJuros
            If lobjstcPedJur.Num_Parcela = mStcA46V702B.Qtd_Parc Then
                If mStcA46V702B.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
                    mStcA46V702B.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Ant
                Else
                    mStcA46V702B.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Pos
                End If
                Exit For
            End If
        Next lobjstcPedJur
    Else

'        If Trim(mStcA46V702B.objstcCtrlEmis.Parcelas_Impressao_Avista) <> "0" And Trim(mStcA46V702B.objstcCtrlEmis.Parcelas_Impressao_Avista) <> "" Then
'            If Trim(mStcA46V702B.objstcCtrlEmis.Parcelas_Impressao_Avista) = "1" Then
'                Qtd_Parc_Ant = Trim(mStcA46V702B.objstcCtrlEmis.Parcelas_Impressao_Avista)
'            Else
'                Qtd_Parc_Ant = Replace(Mid(mStcA46V702B.objstcCtrlEmis.Parcelas_Impressao_Avista, Len(mStcA46V702B.objstcCtrlEmis.Parcelas_Impressao_Avista) - 1, 2), ",", "")
'            End If
'            Qtd_Parc = Qtd_Parc_Ant
'        End If
'
'        If Trim(mStcA46V702B.objstcCtrlEmis.Parcelas_Impressao_30dias) <> "0" And Trim(mStcA46V702B.objstcCtrlEmis.Parcelas_Impressao_30dias) <> "" Then
'            If Trim(mStcA46V702B.objstcCtrlEmis.Parcelas_Impressao_30dias) = "1" Then
'                Qtd_Parc_Pos = Trim(mStcA46V702B.objstcCtrlEmis.Parcelas_Impressao_30dias)
'            Else
'                Qtd_Parc_Pos = Replace(Mid(mStcA46V702B.objstcCtrlEmis.Parcelas_Impressao_30dias, Len(mStcA46V702B.objstcCtrlEmis.Parcelas_Impressao_30dias) - 1, 2), ",", "")
'            End If
'
'            Qtd_Parc = Qtd_Parc_Pos
'        End If
    
        For Each lobjstcPedJur In mStcA46V702B.ColecaoPedJuros
            If lobjstcPedJur.Num_Parcela = mStcA46V702B.Qtd_Parc Then
            'If lobjstcPedJur.Num_Parcela = Qtd_Parc Then
                If mStcA46V702B.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
                    If mStcA46V702B.objstcCtrlEmis.Cod_Bco_Boleto > 0 And mStcA46V702B.Cod_Bco_Cobr > 0 And lobjstcPedJur.Tip_Compensacao = 2 Then '1 carne com demais em debito
                        mStcA46V702B.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Ant
                        Exit For
                    ElseIf lobjstcPedJur.Tip_Compensacao = 1 And mStcA46V702B.gTipPagamento = e_FormaPagamento_Carne Then 'carne
                        mStcA46V702B.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Ant
                        Exit For
                    ElseIf mStcA46V702B.objstcCtrlEmis.Cod_Bco_Boleto = 0 And mStcA46V702B.Cod_Bco_Cobr > 0 And lobjstcPedJur.Tip_Compensacao = 2 Then 'Debito
                        mStcA46V702B.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Ant
                        Exit For
                    ElseIf lobjstcPedJur.Tip_Compensacao = 1 And mStcA46V702B.gTipPagamento = e_FormaPagamento_Credito Then 'Credito
                        mStcA46V702B.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Ant
                        Exit For
                    ElseIf lobjstcPedJur.Tip_Compensacao = 4 And mStcA46V702B.gTipPagamento = e_FormaPagamento_Credito Then 'Credito
                        mStcA46V702B.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Ant
                        Exit For
                    End If
                Else
                    If mStcA46V702B.objstcCtrlEmis.Cod_Bco_Boleto > 0 And mStcA46V702B.Cod_Bco_Cobr > 0 And lobjstcPedJur.Tip_Compensacao = 2 Then '1 carne com demais em debito
                        mStcA46V702B.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Pos
                        Exit For
                    ElseIf lobjstcPedJur.Tip_Compensacao = 1 And mStcA46V702B.gTipPagamento = e_FormaPagamento_Carne Then 'carne
                        mStcA46V702B.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Pos
                        Exit For
                    ElseIf mStcA46V702B.objstcCtrlEmis.Cod_Bco_Boleto = 0 And mStcA46V702B.Cod_Bco_Cobr > 0 And lobjstcPedJur.Tip_Compensacao = 2 Then 'Debito
                        mStcA46V702B.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Pos
                    ElseIf lobjstcPedJur.Tip_Compensacao = 1 And mStcA46V702B.gTipPagamento = e_FormaPagamento_Credito Then 'Credito
                        mStcA46V702B.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Pos
                        Exit For
                    ElseIf lobjstcPedJur.Tip_Compensacao = 4 And mStcA46V702B.gTipPagamento = e_FormaPagamento_Credito Then 'Credito
                        mStcA46V702B.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Pos
                        Exit For
                Exit For
            End If
                End If
            End If
        Next lobjstcPedJur
    End If
    
    lretorno = Format$(mStcA46V702B.objstcCtrlEmis.Per_Juros_Neg, "standard")
    If Len(lretorno) < 5 Then
        lretorno = Space(5 - Len(lretorno)) & lretorno
    End If
    lretorno = IIf(lretorno = "", " ", lretorno) & "%"
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Per_Juros_Neg = lretorno
End Property

Public Property Get Per_Juros_Mora() As String
    Dim lretorno As String
    Dim lobjstcPedJur As stcA46V711B
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V702B.objstcCtrlEmis.Qtd_Parc = 1 And mStcA46V702B.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
        lretorno = 0
    Else
        lretorno = 4.5
    End If
    lretorno = Format$(4.5, "standard")
    If Len(lretorno) < 5 Then
        lretorno = Space(5 - Len(lretorno)) & lretorno
    End If
    lretorno = IIf(lretorno = "", " ", lretorno) & "%"
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Per_Juros_Mora = lretorno
End Property

Public Property Get Banco() As String
    Dim lretorno As String
    Dim lobjstcPedJur As stcA46V711B
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V702B.Cod_Bco_Cobr <> 0 Then
        lretorno = mStcA46V702B.Cod_Bco_Cobr & " - " & mStcA46V702B.gDesc_Banco
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Banco = lretorno
End Property


Public Property Get Val_PriParcela() As String
    Dim lretorno As String
    Dim lobjstcPedJur As stcA46V711B
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = "R$ " & Format(mStcA46V702B.Val_PriParcela, "standard")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Val_PriParcela = lretorno
End Property


Public Property Get val_DemParcelas() As String
    Dim lretorno As String
    Dim lobjstcPedJur As stcA46V711B
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = "R$ " & Format(mStcA46V702B.val_DemParcelas, "standard")
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    val_DemParcelas = lretorno
End Property

Public Property Get AgenciaDig() As String
    Dim lretorno As String
    Dim clsDigito As clsA99V001
    Dim lDigito As String
    Dim lMensagem As String

    If mStcA46V702B.Cod_Bco_Cobr <> 0 And mStcA46V702B.Cod_Agen_Bas_Cor <> 0 Then
        lretorno = mStcA46V702B.gAgenciaDig
        If mStcA46V702B.Cod_Bco_Cobr = 237 Then
            lretorno = Left(Format(lretorno, "00000"), 4) & "-" & Right(lretorno, 1)
        End If
        If mStcA46V702B.Cod_Bco_Cobr = 1 Then    'Projeto 11080 - Debito automatico - Banco do Brasil
            If Right(lretorno, 1) = "0" Then
                Set clsDigito = New clsA99V001
                Set clsDigito.M46V999 = mStcA46V702B.M46V999

                Call clsDigito.DIG_BANCOBRASIL(Format(lretorno, "00000000"), lDigito, lMensagem)
                If lDigito = "X" Then lretorno = Left(Format(lretorno, "00000"), 4) & "X"
            End If
            lretorno = Left(Format(lretorno, "00000"), 4) & "-" & Right(lretorno, 1)
        End If
    End If
    lretorno = IIf(lretorno = "", " ", lretorno)
    AgenciaDig = lretorno

End Property

Public Property Get ContaCorrente() As String
    Dim lretorno As String

    If mStcA46V702B.Cod_Bco_Cobr <> 0 And mStcA46V702B.Cod_Agen_Bas_Cor <> 0 Then
        If mStcA46V702B.mNum_Cta_Bas_Cor_Compl = 0 Then
            mStcA46V702B.mNum_Cta_Bas_Cor_Compl = mStcA46V702B.Num_Cta_Bas_Cor
        End If
        Call mpForConta1(mStcA46V702B.Cod_Bco_Cobr, mStcA46V702B.mNum_Cta_Bas_Cor_Compl, lretorno)
    End If
    lretorno = IIf(lretorno = "", " ", lretorno)
    ContaCorrente = lretorno

End Property

Public Property Get Beneficiarios() As Collection
    Dim lColecao As Collection

    Set lColecao = mfDadosTextoEspecificacao(eTipoRetornoTxtExpec_Beneficiarios)

    Set Beneficiarios = lColecao

SAIDA:
    GoSub FechaObj
    Exit Property

FechaObj:
    Set lColecao = Nothing
    Return
End Property

Public Property Get TextoFranquia() As Collection
    Dim lColecao As Collection

    Set lColecao = mfDadosTextoEspecificacao(eTipoRetornoTxtExpec_Franquia)

    Set TextoFranquia = lColecao

SAIDA:
    GoSub FechaObj
    Exit Property

FechaObj:
    Set lColecao = Nothing
    Return
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

Private Function mfDadosTextoEspecificacao(ByVal TipoRetorno As eTipoRetornoTxtExpec) As Collection
    Dim lCountLine As Integer
    Dim lobjstcTextoEspec As stcA46V726B    'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).
    Dim mColection As Collection
    Dim lobjstcPedDizer As stcA46V722B

    Set mColection = New Collection

    If TipoRetorno = eTipoRetornoTxtExpec_Beneficiarios Then
        For Each lobjstcTextoEspec In mStcA46V702B.ColecaoTextoEspec
            'caso for beneficiario
            If lobjstcTextoEspec.Cod_Espec = 8500 Then
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin1) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin1)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin2) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin2)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin3) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin3)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin4) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin4)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin5) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin5)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin6) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin6)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin7) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin7)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin8) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin8)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin9) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin9)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin10) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin10)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin11) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin11)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin12) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin12)
                End If
            End If
        Next lobjstcTextoEspec
    Else
        For Each lobjstcTextoEspec In mStcA46V702B.ColecaoTextoEspec
            'caso não for beneficiario
            If lobjstcTextoEspec.Cod_Espec <> 8500 Then
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin1) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin1)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin2) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin2)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin3) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin3)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin4) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin4)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin5) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin5)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin6) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin6)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin7) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin7)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin8) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin8)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin9) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin9)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin10) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin10)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin11) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin11)
                End If
                If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin12) Then
                    Call addCollectionImp(mColection, lobjstcTextoEspec.Espec_lin12)
                End If
            End If
        Next lobjstcTextoEspec
    End If


    Set mfDadosTextoEspecificacao = mColection

End Function

Private Sub addCollectionImp(ByRef pCollection As Collection, ByVal Texto As String)
    Dim lobjstcTextoImpressao As stcTextoImpressao
    Set lobjstcTextoImpressao = New stcTextoImpressao
    lobjstcTextoImpressao.Sequencia = pCollection.Count + 1
    lobjstcTextoImpressao.Descricao = Texto
    Call pCollection.Add(lobjstcTextoImpressao, mStcA46V702B.nosso_numero & pCollection.Count + 1)
End Sub

Public Property Get tipoRelatorio() As eTipCalcRPT
    Dim lretorno As eTipCalcRPT
    
    If mStcA46V702B.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        If mStcA46V702B.LMI_Unico = 1 Then
            lretorno = eTipCalcRPT_CotacColetivaLMI
        Else
            If mStcA46V702B.ColecaoPedLoc.Count > 1 Then
                lretorno = eTipCalcRPT_CotacColetiva
                Else
                    lretorno = eTipCalcRPT_CotacIndividual
                End If
        End If
    Else
        If mStcA46V702B.LMI_Unico = 1 Then
            lretorno = eTipCalcRPT_PropColetivaLMI
        Else
            If mStcA46V702B.ColecaoPedLoc.Count > 1 Then
                    lretorno = eTipCalcRPT_PropColetiva
                    Else
                        lretorno = eTipCalcRPT_PropIndividual
                    End If
        End If
    End If
    tipoRelatorio = lretorno
End Property

Private Sub mpForConta1(ByVal pCodBanco As Integer, ByVal pNumConta As Double, ByRef pSaida As String)
'Procedure: formata campo de conta corrente de numérico para caracter.

'Parâmetros de entrada...pBanco......Código do banco.
'                        pNumConta...Número da conta corrente.
'Parâmetro de saída......pSaida......Saída - número da conta corrente.

    Dim lAuxiliar As String    'Auxiliar.
    Dim clsDigito As clsA99V001
    Dim lDigito As String
    Dim lMensagem As String
    Dim lConta As String

    Select Case pCodBanco
    Case 1    'Banco do Brasil. 'Projeto 11080 - Debito automatico - Banco do Brasil
        lAuxiliar = Format$(pNumConta, "000000000")
        If Mid$(lAuxiliar, 9, 1) = "0" Then
            Set clsDigito = New clsA99V001
            Set clsDigito.M46V999 = mStcA46V702B.M46V999

            Call clsDigito.DIG_BANCOBRASIL(lAuxiliar, lDigito, lMensagem)
            If lDigito = "X" Then lAuxiliar = Mid$(lAuxiliar, 1, 8) & "X"
        End If
        pSaida = Mid$(lAuxiliar, 1, 8) & "-" & Mid$(lAuxiliar, 9, 1)
    Case 237    'Bradesco.
        lAuxiliar = Format$(pNumConta, "00000000")
        pSaida = Mid$(lAuxiliar, 1, 7) & "-" & Mid$(lAuxiliar, 8, 1)
    Case 409    'Unibanco.
        lAuxiliar = Format$(pNumConta, "0000000")
        pSaida = Mid$(lAuxiliar, 1, 6) & "-" & Mid$(lAuxiliar, 7, 1)
    Case 341    'Itaú.
        lAuxiliar = Format$(pNumConta, "000000")
        pSaida = Mid$(lAuxiliar, 1, 5) & "-" & Mid$(lAuxiliar, 6, 1)
    Case 347    'Sudameris.
        lAuxiliar = Format$(pNumConta, "0000000000")
        pSaida = Mid$(lAuxiliar, 1, 9) & "-" & Mid$(lAuxiliar, 10, 1)
    Case 479    'Boston.
        lAuxiliar = Format$(pNumConta, "00000000")
        pSaida = Mid$(lAuxiliar, 1, 7) & "-" & Mid$(lAuxiliar, 8, 1)
    Case 356    'Real
        lAuxiliar = Format$(pNumConta, "0000000000")
        pSaida = Mid$(lAuxiliar, 1, 9) & "-" & Mid$(lAuxiliar, 10, 1)
    Case 399
        lAuxiliar = Format$(pNumConta, "00000000000")
        pSaida = Mid$(lAuxiliar, 1, 9) & "-" & Mid$(lAuxiliar, 10, 2)
    Case 47     'Banese
        lAuxiliar = Format$(pNumConta, "000000000")
        pSaida = Mid$(lAuxiliar, 1, 8) & "-" & Mid$(lAuxiliar, 9, 1)
    Case 33     'Santander
        lAuxiliar = Format$(pNumConta, "000000000")
        pSaida = Mid$(lAuxiliar, 1, 8) & "-" & Mid$(lAuxiliar, 9, 1)
    'inclusão novos bancos
    Case 422    'Safra
        lAuxiliar = Format$(pNumConta, "0000000")
        pSaida = Mid$(lAuxiliar, 1, 6) & "-" & Mid$(lAuxiliar, 7, 1)
    Case 756    'BANCOOB
        lAuxiliar = Format$(pNumConta, "00000000000")
        pSaida = Mid$(lAuxiliar, 1, 10) & "-" & Mid$(lAuxiliar, 11, 1)
    Case Else       'Outros bancos.
        pSaida = Format$(pNumConta, "0000000000")
    End Select
End Sub

Public Property Get TextoProposta() As Collection
    Dim lAuxiliar2 As String
    Dim lColecao As Collection

    Set lColecao = New Collection

    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If mStcA46V702B.M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(mStcA46V702B.M46V999.gDataCorteSompo) Then
            If CLng(mStcA46V702B.M46V999.gDataCorteSompo) <= mStcA46V702B.M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If

    ''a) Linha 1.
    lAuxiliar2 = "A aceitação do seguro estará sujeita à análise do risco. O registro deste plano na SUSEP não implica,  por  parte da"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''b) Linha 2.
    lAuxiliar2 = "Autarquia, incentivo ou recomendação à sua comercialização. O segurado poderá consultar a situação cadastral"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''c) Linha 3.
    lAuxiliar2 = "do seu corretor de seguros no site www.susep.gov.br, por meio do número de seu registro na SUSEP,"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''d) Linha 4.
    lAuxiliar2 = "nome  completo, CNPJ ou CPF."
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''036 a 047. Observação 2.
    ''a) Linha 1.
    lAuxiliar2 = "Propõe(m) à " & UCase(lNomeSeguradora) & " SEGUROS S.A. a realização do seguro, declarando conhecer previamente e  aceitar  as  Condições"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''b) Linha 2.
    lAuxiliar2 = "Gerais do(s) anexo(s) que servem de base para esta proposta. Também presta(m)  as  informações  solicitadas  pela"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''c) Linha 3.
    lAuxiliar2 = UCase(lNomeSeguradora) & " SEGUROS S.A., constantes do(s) anexo(s) a esta, para  fins  de  estudos  da  concretização  do  respectivo"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''d) Linha 4.
    lAuxiliar2 = "contrato.  Assume(m)  o(s)  proponente(s),  nesta  oportunidade,  integral  responsabilidade  pela  exatidão  das"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''e) Linha 5.
    lAuxiliar2 = "informações ora prestadas, autorizando a " & UCase(lNomeSeguradora) & " SEGUROS S.A., caso aceite esta proposta, a  emitir  a  respectiva"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''f) Linha 6.
    lAuxiliar2 = "apólice, cujo prêmio e eventuais despesas o(s) proponente(s) compromete(m)-se  a  pagar,  tão  logo  lhe(s)  seja"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''g) Linha 7.
    lAuxiliar2 = "exigido. Em caso de recusa da proposta de seguro, os valores eventualmente adiantados para  futuro  pagamento  de"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''h) Linha 8.
    lAuxiliar2 = "prêmio pelo(s) proponente(s) serão devolvidos integralmente, no prazo de 10 dias corridos da data de sua  recusa."
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''i) Linha 9.
    lAuxiliar2 = "Ultrapassado este prazo, os valores sujeitam-se a atualização monetária pelo IPC/FIPE, contados a partir da  data"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''j) Linha 10.
    lAuxiliar2 = "em que se tornarem exigíveis e juros moratórios de 0,5% (cinco décimos por cento) ao mês, contados  a  partir  do"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''k) Linha 11.
    lAuxiliar2 = "primeiro dia posterior ao término do prazo fixado acima. Ambos calculados pro-rata-die  até  a  data  da  efetiva restituição."
    Call addCollectionImp(lColecao, lAuxiliar2)

    'CR20834
    ''l) Linha 12.
    '    lAuxiliar2 = "AS CONDIÇÕES CONTRATUAIS/REGULAMENTOS DESTE PRODUTO ENCONTRAM-SE REGISTRADAS NA SUSEP"
    lAuxiliar2 = "As condições contratuais/regulamentos deste produto protocolizadas pela sociedade/entidade"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''m) Linha 13.
    '    lAuxiliar2 = "DE ACORDO COM O NÚMERO DE PROCESSO CONSTANTE DA APÓLICE/PROPOSTA E PODERÃO SER CONSULTADAS"
    lAuxiliar2 = "junto à susep poderão ser consultadas no endereço eletrônico www.susep.gov.br, de acordo com o"
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''n) Linha 14.
    '    lAuxiliar2 = "NO ENDEREÇO ELETRÔNICO WWW.SUSEP.GOV.BR <HTTP://WWW.SUSEP.GOV.BR>."
    lAuxiliar2 = "número de processo constante da apólice/proposta."
    Call addCollectionImp(lColecao, lAuxiliar2)

    ''o) Linha 15.
    'Nova tarifa 1100
'    lAuxiliar2 = "O REGISTRO DESTE PLANO DA SUSEP NÃO IMPLICA, POR PARTE DA AUTARQUIA, INCENTIVO"
'    Call addCollectionImp(lColecao, lAuxiliar2)

    ''p) Linha 16.
'    lAuxiliar2 = "OU RECOMENDAÇÃO A SUA COMERCIALIZAÇÃO."
'    Call addCollectionImp(lColecao, lAuxiliar2)

    Set TextoProposta = lColecao
SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lColecao = Nothing
    Return
End Property

Public Property Get TextoPropostaNew() As String
    Dim lAuxiliar2 As String

    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If mStcA46V702B.M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(mStcA46V702B.M46V999.gDataCorteSompo) Then
            If CLng(mStcA46V702B.M46V999.gDataCorteSompo) <= mStcA46V702B.M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If
    
    ''a) Linha 1.
    lAuxiliar2 = lAuxiliar2 + "A aceitação do seguro estará sujeita à análise do risco. O registro deste plano na SUSEP não implica,  por  parte da" '+ vbCrLf

    ''b) Linha 2.
    lAuxiliar2 = lAuxiliar2 + " Autarquia, incentivo ou recomendação à sua comercialização. O segurado poderá consultar a situação cadastral" '+ vbCrLf

    ''c) Linha 3.
    lAuxiliar2 = lAuxiliar2 + " do seu corretor de seguros no site www.susep.gov.br, por meio do número de seu registro na SUSEP," '+ vbCrLf

    ''d) Linha 4.
    lAuxiliar2 = lAuxiliar2 + " nome  completo, CNPJ ou CPF." '+ vbCrLf

    ''036 a 047. Observação 2.
    ''a) Linha 1.
    lAuxiliar2 = lAuxiliar2 + "Propõe(m) à " & UCase(lNomeSeguradora) & " SEGUROS S.A. a realização do seguro, declarando conhecer previamente e  aceitar  as  Condições" '+ vbCrLf

    ''b) Linha 2.
    lAuxiliar2 = lAuxiliar2 + " Gerais do(s) anexo(s) que servem de base para esta proposta. Também presta(m)  as  informações  solicitadas  pela" '+ vbCrLf

    ''c) Linha 3.
    lAuxiliar2 = lAuxiliar2 + " " & UCase(lNomeSeguradora) & " SEGUROS S.A., constantes do(s) anexo(s) a esta, para  fins  de  estudos  da  concretização  do  respectivo" '+ vbCrLf

    ''d) Linha 4.
    lAuxiliar2 = lAuxiliar2 + " contrato.  Assume(m)  o(s)  proponente(s),  nesta  oportunidade,  integral  responsabilidade  pela  exatidão  das" '+ vbCrLf

    ''e) Linha 5.
    lAuxiliar2 = lAuxiliar2 + " informações ora prestadas, autorizando a " & UCase(lNomeSeguradora) & " SEGUROS S.A., caso aceite esta proposta, a  emitir  a  respectiva" '+ vbCrLf

    ''f) Linha 6.
    lAuxiliar2 = lAuxiliar2 + " apólice, cujo prêmio e eventuais despesas o(s) proponente(s) compromete(m)-se  a  pagar,  tão  logo  lhe(s)  seja" '+ vbCrLf

    ''g) Linha 7.
    lAuxiliar2 = lAuxiliar2 + " exigido. Em caso de recusa da proposta de seguro, os valores eventualmente adiantados para  futuro  pagamento  de" '+ vbCrLf
    
    ''h) Linha 8.
    lAuxiliar2 = lAuxiliar2 + " prêmio pelo(s) proponente(s) serão devolvidos integralmente, no prazo de 10 dias corridos da data de sua  recusa." '+ vbCrLf
    
    ''i) Linha 9.
    lAuxiliar2 = lAuxiliar2 + "Ultrapassado este prazo, os valores sujeitam-se a atualização monetária pelo IPC/FIPE, contados a partir da  data" '+ vbCrLf

    ''j) Linha 10.
    lAuxiliar2 = lAuxiliar2 + " em que se tornarem exigíveis e juros moratórios de 0,5% (cinco décimos por cento) ao mês, contados  a  partir  do" '+ vbCrLf
    
    ''k) Linha 11.
    lAuxiliar2 = lAuxiliar2 + " primeiro dia posterior ao término do prazo fixado acima. Ambos calculados pro-rata-die  até  a  data  da  efetiva restituição." '+ vbCrLf

    'CR20834
    ''l) Linha 12.
    '    lAuxiliar2 = "AS CONDIÇÕES CONTRATUAIS/REGULAMENTOS DESTE PRODUTO ENCONTRAM-SE REGISTRADAS NA SUSEP"
    lAuxiliar2 = lAuxiliar2 + "As condições contratuais/regulamentos deste produto protocolizadas pela sociedade/entidade" '+ vbCrLf

    ''m) Linha 13.
    '    lAuxiliar2 = "DE ACORDO COM O NÚMERO DE PROCESSO CONSTANTE DA APÓLICE/PROPOSTA E PODERÃO SER CONSULTADAS"
    lAuxiliar2 = lAuxiliar2 + " junto à susep poderão ser consultadas no endereço eletrônico www.susep.gov.br, de acordo com o" '+ vbCrLf

    ''n) Linha 14.
    '    lAuxiliar2 = "NO ENDEREÇO ELETRÔNICO WWW.SUSEP.GOV.BR <HTTP://WWW.SUSEP.GOV.BR>."
    lAuxiliar2 = lAuxiliar2 + " número de processo constante da apólice/proposta." '+ vbCrLf

    ''o) Linha 15.
    'Nova tarifa 1100
'    lAuxiliar2 = lAuxiliar2 + "O REGISTRO DESTE PLANO DA SUSEP NÃO IMPLICA, POR PARTE DA AUTARQUIA, INCENTIVO" + vbCrLf
'
'    ''p) Linha 16.
'    lAuxiliar2 = lAuxiliar2 + "OU RECOMENDAÇÃO A SUA COMERCIALIZAÇÃO." + vbCrLf
    
    TextoPropostaNew = lAuxiliar2
SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Return
End Property

Public Property Get End_Corresp_Logradouro() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.Cod_Tip_Logr & " " & _
               mStcA46V702B.Nom_Logr & ", " & _
               mStcA46V702B.Dsc_Compl_Logr
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    End_Corresp_Logradouro = lretorno
End Property

Public Property Get End_Corresp_Bairro() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.Nom_Bairro
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    'CR29035 - Ajuste para abreviar a descrição do bairro pois o mesmo quando era maior que 23 posições estava sobrescrevendo o cep.
    'Abreviar se tiver a palavra PARQUE no bairro
    If Len(lretorno) > 23 Then
        If InStr(1, lretorno, "PARQUE", vbTextCompare) > 0 Then
            lretorno = Replace(lretorno, "PARQUE", "PQ.", , , vbTextCompare)
        End If
    End If

    'Abreviar se tiver a palavra RESIDENCIAL no bairro
    If Len(lretorno) > 23 Then
        If InStr(1, lretorno, "RESIDENCIAL", vbTextCompare) > 0 Then
            lretorno = Replace(lretorno, "RESIDENCIAL", "RES.", , , vbTextCompare)
        End If
    End If

    'Se mesmo com as abreviações ainda for maior, truncar em 24 posições.
    If Len(lretorno) > 23 Then
        lretorno = Mid(lretorno, 1, 23)
    End If

    lretorno = IIf(lretorno = "", " ", lretorno)
    End_Corresp_Bairro = lretorno
End Property

Public Property Get End_Corresp_Cidade() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.Nom_Cid
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    End_Corresp_Cidade = lretorno
End Property

Public Property Get End_Corresp_UF() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.Sig_UF
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    End_Corresp_UF = lretorno
End Property

Public Property Get End_Corresp_CEP() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V702B.gCEP_Formatado
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    End_Corresp_CEP = lretorno
End Property

Public Property Get CodUnidade() As String
    Dim lretorno As String
    lretorno = mStcA46V702B.Cod_Unid_Prod
    lretorno = IIf(lretorno = "", " ", lretorno)
    CodUnidade = lretorno
End Property

Public Property Get CodProdutor() As String
    Dim lretorno As String
    lretorno = mStcA46V702B.Cod_Prod
    lretorno = IIf(lretorno = "", " ", lretorno)
    CodProdutor = lretorno
End Property

Public Property Get LinhaParcela() As String
    Dim lretorno As String

    'If Val(mStcA46V702B.objstcCtrlEmis.Num_Bloq & mStcA46V702B.objstcCtrlEmis.DIG_BLOQUETO) = 0 Then
    'comentado aqui p/ baixo
    'If mStcA46V702B.objstcCtrlEmis.Cod_Iden_Pag_Antec <> 1 Then
    '    lretorno = Format(mStcA46V702B.Qtd_Parc - IIf((mStcA46V702B.objstcCtrlEmis.Cod_Bco_Boleto <> 0 _
         '                                                   And mStcA46V702B.Cod_Bco_Cobr <> 0), 1, 0), "00") & _
         '                        " parcelas de R$ " & _
         '                        Format(mStcA46V702B.val_DemParcelas, "Standard") & " cada."
    'Else
    '    lretorno = "1ª Parcela de R$ " & Format(mStcA46V702B.Val_PriParcela, "Standard") & " + " & _
         '                        Format(mStcA46V702B.Qtd_Parc, "00") - 1 & " parcelas de R$ " & _
         '                        Format(mStcA46V702B.val_DemParcelas, "Standard") & " cada."
    'End If

    If mStcA46V702B.objstcCtrlEmis.Cod_Iden_Pag_Antec <> 1 Then
        If mStcA46V702B.Qtd_Parc = 1 Then
        lretorno = Format(mStcA46V702B.Qtd_Parc, "00") & _
                 " parcelas de R$ " & _
                   Format(mStcA46V702B.gValorTotal, "Standard") & " cada."
        ElseIf mStcA46V702B.objstcCtrlEmis.Cod_Bco_Boleto > 0 And mStcA46V702B.Num_Cta_Bas_Cor > 0 Then
            lretorno = Format(mStcA46V702B.Qtd_Parc, "00") - 1 & " parcela(s) de R$ " & _
                       Format(mStcA46V702B.val_DemParcelas, "Standard") & " cada."
        Else
            lretorno = "1ª Parcela de R$ " & Format(mStcA46V702B.Val_PriParcela, "Standard") & " + " & _
                       Format(mStcA46V702B.Qtd_Parc, "00") - 1 & " parcelas de R$ " & _
                       Format(mStcA46V702B.val_DemParcelas, "Standard") & " cada."
        End If

    Else
        If mStcA46V702B.Cod_Bco_Cobr <> 0 And mStcA46V702B.Qtd_Parc = 1 Then
            lretorno = "1 parcela de R$ " & _
                       Format(mStcA46V702B.gValorTotal, "Standard") & "."
        ElseIf mStcA46V702B.objstcCtrlEmis.Cod_Bco_Boleto > 0 And mStcA46V702B.Num_Cta_Bas_Cor > 0 Then
            lretorno = Format(mStcA46V702B.Qtd_Parc, "00") - 1 & " parcelas de R$ " & _
                       Format(mStcA46V702B.val_DemParcelas, "Standard") & " cada."
        Else
            lretorno = "1ª Parcela de R$ " & Format(mStcA46V702B.Val_PriParcela, "Standard") & " + " & _
                       Format(mStcA46V702B.Qtd_Parc, "00") - 1 & " parcelas de R$ " & _
                       Format(mStcA46V702B.val_DemParcelas, "Standard") & " cada."
        End If
    End If

    lretorno = IIf(lretorno = "", " ", lretorno)
    LinhaParcela = lretorno
End Property

Public Property Get VencimentoParcelas() As String
    Dim lObjstcVenc As stcA46V713B
    Dim lretorno As String
    Dim lparc As Integer
    'Ocultar a primeira se for Pagto Antec

    If mStcA46V702B.ColecaoPedVenc.Count > 0 Then
        lretorno = ""
        lparc = 1
        For Each lObjstcVenc In mStcA46V702B.ColecaoPedVenc
            If Not (lObjstcVenc.Num_Parc = 1 And mStcA46V702B.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim) Then
                If lObjstcVenc.Num_Parc <= mStcA46V702B.Qtd_Parc Then
                    If lparc < 6 Then
                        lretorno = lretorno & "(" & lObjstcVenc.Num_Parc & "ª-" & mStcA46V702B.M46V999.gfForData(2, lObjstcVenc.DAT_VENC) & ") "
                    Else
                        If lparc = 6 Then
                            lretorno = lretorno & vbCrLf
                        End If
                        lretorno = lretorno & "(" & lObjstcVenc.Num_Parc & "ª-" & mStcA46V702B.M46V999.gfForData(2, lObjstcVenc.DAT_VENC) & ") "
                    End If
                End If
                lparc = lparc + 1
            End If
        Next lObjstcVenc
    End If
    lretorno = IIf(lretorno = "", " ", lretorno)

    VencimentoParcelas = lretorno

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lObjstcVenc = Nothing
    Return
End Property

Public Property Get LocalData() As String
    Dim lretorno As String
    lretorno = "São Paulo, " & Format(Now, "DD") & _
             " de " & Format(Now, "MMMM") & _
             " de " & Format(Now, "yyyy")

    LocalData = lretorno
End Property

Public Property Get ProcessoInterno() As String
    ProcessoInterno = mStcA46V702B.Class_ID_PI
End Property

Public Property Get DatIniVig() As String
    Dim lretorno As String

    lretorno = mStcA46V702B.M46V999.gfForData(2, mStcA46V702B.Dat_Ini_Vig)

    DatIniVig = lretorno
End Property

Public Property Get DatTerVig() As String
    Dim lretorno As String

    lretorno = mStcA46V702B.M46V999.gfForData(2, mStcA46V702B.Dat_Ter_Vig)

    DatTerVig = lretorno
End Property

Public Property Get ClausulasLMI() As Collection
    Dim lAuxiliar2 As String
    Dim lColecao As Collection
    Dim lI As Integer             'Utilizado com For...Next.
    Dim lListClausula As String
    Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcClau As stcA46V703B    'Objeto da classe clsA46V703A (TAB_PED_CLAU).

    Set lColecao = New Collection

    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If mStcA46V702B.M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(mStcA46V702B.M46V999.gDataCorteSompo) Then
            If CLng(mStcA46V702B.M46V999.gDataCorteSompo) <= mStcA46V702B.M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If
    
    ''a) Linha 1.
    Call addCollectionImp(lColecao, "Ratificam-se as seguintes Cláusulas e Condições:")

    lAuxiliar2 = " "
    Select Case mStcA46V702B.Cod_Ramo
    Case e_CodRamo_Condominio
        If mStcA46V702B.Cod_Produto >= 1080 Then
            If mStcA46V702B.Cod_Produto >= 1100 Then
                lAuxiliar2 = "a) Condições Gerais do Produto Condomínio"
            Else
            lAuxiliar2 = "a) Condições Gerais do Produto Condomínio (versão 10/2014)"
            End If
        Else
            lAuxiliar2 = "a) Condições Gerais do Produto " & lNomeSeguradora & " Condomínio (versão 09/2008)"
        End If
            
        Call addCollectionImp(lColecao, lAuxiliar2)
        lI = 98
    Case e_CodRamo_Empresarial
        If mStcA46V702B.Cod_Produto >= 1080 Then
            If mStcA46V702B.Cod_Produto >= 1100 Then
                lAuxiliar2 = "a) Condições Gerais do Produto Empresarial"
            Else
            lAuxiliar2 = "a) Condições Gerais do Produto Empresarial (versão 10/2014)"
            End If
        Else
            lAuxiliar2 = "a) Condições Gerais do Produto " & lNomeSeguradora & " Empresarial (versão 06/2009)"
        End If
            
        Call addCollectionImp(lColecao, lAuxiliar2)
        '1200029 - Clausula 315 e Texto LMI
        '''            lAuxiliar2 = "b) Cláusula 315- LMI Único "
        '''            Call addCollectionImp(lColecao, lAuxiliar2
        lI = 98
    Case e_CodRamo_Residencial
        If mStcA46V702B.Cod_Produto >= 1080 Then
            If mStcA46V702B.Cod_Produto >= 1100 Then
                lAuxiliar2 = "a) Condições Gerais do Produto Residencial"
            Else
            lAuxiliar2 = "a) Condições Gerais do Produto Residencial (versão 10/2014)"
            End If
        Else
            lAuxiliar2 = "a) Condições Gerais do Produto " & lNomeSeguradora & " Residencial (versão 09/2008)"
        End If
            
        Call addCollectionImp(lColecao, lAuxiliar2)
        lI = 98
    End Select

    If mStcA46V702B.Cod_Ramo = e_CodRamo_Residencial Or mStcA46V702B.Cod_Ramo = e_CodRamo_Condominio Then
        lAuxiliar2 = Chr(lI) & ") CLÁUSULA PARTICULAR L.M.I. Único - Declara-se para os devidos fins e efeitos que, este seguro foi contratado sob "
        lI = lI + 1
        Call addCollectionImp(lColecao, lAuxiliar2)
        lAuxiliar2 = "a forma de LMI Único para todos os locais segurados especificados na presente apólice. Ao contrário do que possa constar "
        Call addCollectionImp(lColecao, lAuxiliar2)
        lAuxiliar2 = "nas condições gerais do presente contrato de seguro, fica entendido e acordado que em caso de eventual sinistro, a "
        Call addCollectionImp(lColecao, lAuxiliar2)
        lAuxiliar2 = "indenização não poderá exceder ao LMI de cada cobertura contratada, limitado ao VR declarado para cada local."
        Call addCollectionImp(lColecao, lAuxiliar2)
    End If

    lListClausula = ""
    For Each lobjstcPedLoc In mStcA46V702B.ColecaoPedLoc
        For Each lobjstcClau In lobjstcPedLoc.ColecaoPedClau
            If lobjstcClau.Cod_Clau >= 310 Or (lobjstcClau.Cod_Ramo = e_CodRamo_Residencial And lobjstcClau.Cod_Clau > 300) Or mStcA46V702B.LMI_Unico = 1 Then
                If Not lListClausula Like "*;" & lobjstcClau.Cod_Clau & ";*" Then
                    lListClausula = lListClausula & ";" & lobjstcClau.Cod_Clau & ";"
                    If lobjstcClau.gEhAss24h Then
                        lAuxiliar2 = Chr(lI) & ") Cláusula " & lobjstcClau.Cod_Clau & " - " & "Assistência 24 horas" & " - " & lobjstcClau.gDescClausula
                    Else
                    lAuxiliar2 = Chr(lI) & ") Cláusula " & lobjstcClau.Cod_Clau & " - " & lobjstcClau.gDescClausula
                    End If
                    Call addCollectionImp(lColecao, lAuxiliar2)
                    lI = lI + 1
                End If
            End If
        Next lobjstcClau
    Next

    Set ClausulasLMI = lColecao
SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lColecao = Nothing
    Set lobjstcPedLoc = Nothing
    Set lobjstcClau = Nothing
    Return
End Property

Public Property Get ObservacoesLMI() As Collection
    Dim lAuxiliar2 As String
    Dim lColecao As Collection
    Dim lI As Integer             'Utilizado com For...Next.
    Dim lTam As Integer
    Dim lTemObs As Boolean
    Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lObs As Variant
    Dim lCount As Integer
    Dim lMsgInspecao As String
    Dim lTextoCong As String
    Dim lDescImportante() As String
    Dim lPos As Long
    Dim Aux As Integer
    Dim blnCorretorVolks As Boolean
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lConcessionaria As Boolean
    Dim lValorComis As Single
    Dim lTemDesc As Boolean
    Dim Linha As Variant
    Dim lobjstcCob As stcA46V704B
    Dim lArrayAddTextoCodCobert() As Integer
    Dim i As Integer

    Set lColecao = New Collection
    ReDim Preserve lArrayAddTextoCodCobert(0)
    lI = 0
    lTam = 95
    lTemObs = False
    
    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If mStcA46V702B.M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(mStcA46V702B.M46V999.gDataCorteSompo) Then
            If CLng(mStcA46V702B.M46V999.gDataCorteSompo) <= mStcA46V702B.M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If
    
    'Soma da Comissão para saber se o prêmio é NET
    If mStcA46V702B.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = mStcA46V702B.ColecaoPedCorr.Item(1)
        lValorComis = lObjstcCorr.Per_Corr
        If mStcA46V702B.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = mStcA46V702B.ColecaoPedCorr.Item(2)
            lValorComis = lValorComis + lObjstcCorr.Per_Corr
            If mStcA46V702B.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = mStcA46V702B.ColecaoPedCorr.Item(3)
                lValorComis = lValorComis + lObjstcCorr.Per_Corr
            End If
        End If
    End If
    
    If mStcA46V702B.M46V111.gNum_Pi = mStcA46V702B.NUM_PI Then    'Se estiver imprimindo o mesmo que está sendo calculado
        For Each lobjstcPedLoc In mStcA46V702B.ColecaoPedLoc

            'Set lobjstcPedLoc = mStcA46V702B.ColecaoPedLoc.Item(1)
            lAuxiliar2 = lobjstcPedLoc.objstcPedObs.Observacao
            If mStcA46V702B.M46V999.gfPreenchido(lAuxiliar2) Then
                lAuxiliar2 = lAuxiliar2 & "&%"
            End If

            lObs = Split(lAuxiliar2, "&%")
            For lCount = 0 To UBound(lObs)
                lAuxiliar2 = lObs(lCount)
                If lAuxiliar2 Like "*OBRIGATORIO A REALIZACAO DE INSPECAO DO RISCO*" Then
                    lMsgInspecao = lAuxiliar2
                    lAuxiliar2 = ""
                End If
                If lAuxiliar2 Like "*ENVIO DE COPIA DA APOLICE DA CONGENERE*" Then
                    lTextoCong = lAuxiliar2
                    lAuxiliar2 = ""
                End If


                If mStcA46V702B.M46V999.gfPreenchido(lAuxiliar2) Then
                    If Len(lAuxiliar2) < lTam Then
                        ReDim Preserve lDescImportante(lI)
                        lDescImportante(lI) = lAuxiliar2
                        lTemObs = True
                        lI = lI + 1
                    Else
                        lPos = 1
                        Do
                            ReDim Preserve lDescImportante(lI)
                            lTemObs = True
                            Aux = 0
                            If Len(Mid(lAuxiliar2, lPos)) > lTam Then
                                Aux = InStrRev(Mid(lAuxiliar2, lPos, lTam), " ")
                                'Sandra
                                If Aux > 0 And Aux < lTam Then
                                    Aux = lTam - Aux
                                Else
                                    Aux = 0
                                End If
                            End If
                            lDescImportante(lI) = Mid(lAuxiliar2, lPos, lTam - Aux)
                            lI = lI + 1
                            'Sandra
                            If (Len(Mid(lAuxiliar2, lPos, lTam)) <= lTam) And Aux = 0 Then Exit Do
                            lPos = lPos + (lTam - Aux)
                        Loop
                    End If
                End If
            Next lCount
        Next lobjstcPedLoc
    End If


    Call addCollectionImp(lColecao, "Observações:")
    lCount = 1

    lAuxiliar2 = lCount & "-Prêmios, taxas, franquias e demais condições sujeitas à alteração mediante inspeção no local segurado, e/ou caso sejam"
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))
    lCount = lCount + 1

    lAuxiliar2 = "constatados fatores e/ou informações divergentes daquelas fornecidas e/ou atividades agravantes provisórias e/ou definitivas."
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))

    'Ficha 1000742 - Não exibir especificação Item 2 para Corretor Volkswagem
    blnCorretorVolks = False
    For Each lObjstcCorr In mStcA46V702B.ColecaoPedCorr
        If InStr(1, "0008001, 0008002, 0008003, 0008004, 0008005, 0008006, 0008008, 0008011, 0008014, 0008028", lObjstcCorr.Cod_Corr, vbTextCompare) > 0 Then
            blnCorretorVolks = True
            If lConcessionaria = False Then blnCorretorVolks = False
        End If
    Next

    'Ficha 1000742 - Não exibir especificação Item 2 para Corretor Volkswagem
    If blnCorretorVolks = False Then
        lAuxiliar2 = lCount & "-Acham-se excluídas da presente cotação as construções mistas, inferiores, armazéns infláveis, galpão de vinilona e"
        Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))

        lAuxiliar2 = "assemelhados e seus respectivos conteúdos, porventura existentes no local segurado."
        Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))
    End If
    lCount = lCount + 1

    If lValorComis = 0 Then
        lAuxiliar2 = lCount & "-Prêmio acima desprovido de Comissão de Corretagem"
    Else
        lAuxiliar2 = lCount & "-Prêmio acima com carregamento de Comissão de Corretagem"
    End If
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))
    lCount = lCount + 1

    lAuxiliar2 = lCount & "-A presente cotação é apresentada apenas como parâmetro de custos, e não implica no reconhecimento da aceitação do risco,"
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))
    lCount = lCount + 1

    lAuxiliar2 = "a qual dependerá da análise do mesmo,após a respectiva inspeção,bem como as taxas,franquias e condições."
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))

    lAuxiliar2 = lCount & "-O registro deste plano na SUSEP não implica, por parte da Autarquia, incentivo ou recomendação a sua comercialização."
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))
    lCount = lCount + 1

    lAuxiliar2 = lCount & "-O segurado poderá consultar a situação cadastral de seu corretor de seguros, no site www.susep.gov.br, por meio de seu"
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))
    lCount = lCount + 1

    lAuxiliar2 = "registro na SUSEP, nome completo, CNPJ ou CPF."
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))

    lAuxiliar2 = lCount & "-Em havendo outros seguros efetuados do mesmo local e mesmo segurado à " & lNomeSeguradora & " Seguros deverá ser comunicado,"
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))
    lCount = lCount + 1

    lAuxiliar2 = "tendo em vista acumulo de risco."
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))

    lAuxiliar2 = lCount & "-Cotação válida por 15 dias."
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))
    lCount = lCount + 1

    If mStcA46V702B.Cod_Produto < 1080 Then
        Select Case mStcA46V702B.Cod_Ramo
        Case e_CodRamo_Condominio
                lAuxiliar2 = lCount & "-Assistência 24 Horas Gratuita."
        Case e_CodRamo_Empresarial
                lAuxiliar2 = lCount & "-Assistência 24 Horas Gratuita."
        Case e_CodRamo_Residencial
                lAuxiliar2 = lCount & "-Assistência Domiciliar 24 Horas Gratuita (exceto para residências eventuais)."
        End Select
        Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))
        lCount = lCount + 1
    End If

    '1101065 - Mensagem LMI
    lAuxiliar2 = lCount & "-Esta cotação foi elaborada sob a forma de LMI Único. Verifique os sub-limites estabelecidos na relação dos locais de risco"
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))
    lCount = lCount + 1

    lAuxiliar2 = "para cada endereço segurado e, caso não localize ou ainda, haja alguma divergência quanto aos valores e/ou locais segurados,"
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))

    lAuxiliar2 = "entre em contato com seu corretor de seguros e solicite uma nova cotação."
    Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))

    Dim lobjstcTextoEspec
    If mStcA46V702B.ColecaoTextoEspec.Count > 0 Then
        For Each lobjstcTextoEspec In mStcA46V702B.ColecaoTextoEspec
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin1) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin1)
            End If
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin2) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin2)
            End If
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin3) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin3)
            End If
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin4) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin4)
            End If
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin5) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin5)
            End If
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin6) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin6)
            End If
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin7) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin7)
            End If
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin8) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin8)
            End If
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin9) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin9)
            End If
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin10) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin10)
            End If
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin11) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin11)
            End If
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin12) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin12)
            End If
            
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin13) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin13)
            End If
            
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin14) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin14)
            End If
            
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin15) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin15)
            End If

            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin16) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin16)
            End If
            
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin17) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin17)
            End If

            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin18) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin18)
            End If
            
            If mStcA46V702B.M46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin19) Then
                Call addCollectionImp(lColecao, lobjstcTextoEspec.Espec_lin19)
            End If
            
            'End If
        Next lobjstcTextoEspec
    End If

    If mStcA46V702B.Cod_Ramo = e_CodRamo_Empresarial And lConcessionaria = True Then
        Select Case mStcA46V702B.objstcCtrlEmis.Cod_Corr
        Case 8001, 8002, 8003, 8004, 8005, 8006, 8008, 8011, 8014, 8028
            lAuxiliar2 = "Mercadorias e matérias-primas de terceiros em poder do segurado: seus valores deverão ser considerados na composição "
            Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))

            lAuxiliar2 = "do Valor Risco Total Declarado, devendo ser observado os bens excluídos das Condições Gerais do Produto Multi"
            Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))

            lAuxiliar2 = "Seguro Empresarial;"
            Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))
        End Select
    End If

    'Ficha 1000307
    If mStcA46V702B.gObrRegistro = True Then
        If mStcA46V702B.Cod_Ramo = e_CodRamo_Empresarial And mStcA46V702B.Tip_Pessoa = e_TipPessoa_Fisica And mStcA46V702B.M46V999.gTipCobertura <> e_TipCobert_Predio Then

            lAuxiliar2 = "Seguro de Pessoa Física: Na proposta será obrigatório informar o número do registro profissional no respectivo Conselho "
            Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))

            lAuxiliar2 = "da Categoria. Ex.: OAB, CRM, CRO etc. "
            Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))
        End If
    End If

    If mStcA46V702B.M46V111.gNum_Pi = mStcA46V702B.NUM_PI Then    'Se estiver imprimindo o mesmo que está sendo calculado
        If lTemObs = True Then
            For lI = 0 To UBound(lDescImportante)
                lAuxiliar2 = lDescImportante(lI)
                Call addCollectionImp(lColecao, IIf(lAuxiliar2 = "", " ", lAuxiliar2))
            Next lI
        End If
    End If

    If mStcA46V702B.M46V999.gfPreenchido(lMsgInspecao) Then
        lAuxiliar2 = lMsgInspecao
        If Len(lAuxiliar2) < lTam Then
            Call addCollectionImp(lColecao, lMsgInspecao)
        Else
            lPos = 1
            Do
                Aux = 0
                If Len(Mid(lAuxiliar2, lPos)) > lTam Then
                    Aux = InStrRev(Mid(lAuxiliar2, lPos, lTam), " ")
                    If Aux < lTam Then
                        Aux = lTam - Aux
                    Else
                        Aux = 0
                    End If
                End If
                Call addCollectionImp(lColecao, Mid(lAuxiliar2, lPos, lTam - Aux))
                If (Len(Mid(lAuxiliar2, lPos, lTam)) < lTam) And Aux = 0 Then Exit Do
                lPos = lPos + (lTam - Aux)
            Loop
        End If
    End If
    If mStcA46V702B.M46V999.gfPreenchido(lTextoCong) Then
        lAuxiliar2 = lTextoCong
        If Len(lAuxiliar2) < lTam Then
            Call addCollectionImp(lColecao, lTextoCong)
        Else
            lPos = 1
            Do
                Aux = 0
                If Len(Mid(lAuxiliar2, lPos)) > lTam Then
                    Aux = InStrRev(Mid(lAuxiliar2, lPos, lTam), " ")
                    If Aux < lTam Then
                        Aux = lTam - Aux
                    Else
                        Aux = 0
                    End If
                End If
                Call addCollectionImp(lColecao, Mid(lAuxiliar2, lPos, lTam - Aux))
                If (Len(Mid(lAuxiliar2, lPos, lTam)) < lTam) And Aux = 0 Then Exit Do
                lPos = lPos + (lTam - Aux)
            Loop
        End If
    End If

    If lTemDesc = True Then
        Call addCollectionImp(lColecao, "PRÊMIO CALCULADO COM DESCONTO COMERCIAL")
    End If

    'Texto Observação de Franquia
    For Each lobjstcPedLoc In mStcA46V702B.ColecaoPedLoc
        For Each lobjstcCob In lobjstcPedLoc.ColecaoPedCobert
            'Verifica se a cobertura já passou pela verificação de textos para não gravar duplicado
            For i = 0 To UBound(lArrayAddTextoCodCobert)
                If lobjstcCob.Cod_Cobert = lArrayAddTextoCodCobert(i) Then
                    GoTo ProximaCobert
                End If
            Next i
        
            If lobjstcCob.ColecaoNovaFranquia.Count > 0 Then
                If lobjstcCob.ColecaoNovaFranquia.Item(1).Texto_Observacoes <> Empty Then
                    If InStr(lobjstcCob.ColecaoNovaFranquia.Item(1).Texto_Observacoes, Chr(10)) > 0 Then
                        Linha = lobjstcCob.ColecaoNovaFranquia.Item(1).Texto_Observacoes

                        Call addCollectionImp(lColecao, Linha)
                    Else
                        Linha = mfSplitEX(lobjstcCob.ColecaoNovaFranquia.Item(1).Texto_Observacoes, Asc(vbCrLf), 110)

                        For lI = 0 To UBound(Linha)
                            Call addCollectionImp(lColecao, Linha(lI))
                        Next lI
                    End If
                End If
            End If
            
            ReDim Preserve lArrayAddTextoCodCobert(i)
            lArrayAddTextoCodCobert(i) = lobjstcCob.Cod_Cobert
ProximaCobert:
        Next
    Next

    'thais - lmi
    Dim lobjstcPedDizer As stcA46V722B
    
    For Each lobjstcPedDizer In mStcA46V702B.ColecaoTextoEndo
        Call addCollectionImp(lColecao, lobjstcPedDizer.Dsc_Dizer_Endo)
    Next lobjstcPedDizer
    

    Set ObservacoesLMI = lColecao
SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lColecao = Nothing
    Set lobjstcPedLoc = Nothing
    Set lObjstcCorr = Nothing
    Set lobjstcCob = Nothing
    Return
End Property

Private Function mfSplitEX(ByVal pTexto As String, ByVal pDelimitador As Byte, ByVal pMaximo As Integer) As String()
    Dim Retorno() As String
    Dim i As Integer
    Dim pos As Integer

    ReDim Retorno(1000)

    i = 0
    Do While Len(pTexto) > 0

        If InStr(pTexto, Chr(pDelimitador)) <= pMaximo And InStr(pTexto, Chr(pDelimitador)) > 0 Then
            Retorno(i) = Left(pTexto, InStr(pTexto, Chr(pDelimitador)))
            pTexto = Mid(pTexto, Len(Retorno(i)) + 2, Len(pTexto))
        Else
            If InStr(pTexto, Chr(pDelimitador)) = 0 Then
                pos = InStrRev(Left(pTexto, pMaximo), " ")
                If pos = 0 Then
                    pos = pMaximo
                End If
                Retorno(i) = Left(pTexto, pos)
                pTexto = Mid(pTexto, Len(Retorno(i)) + 1, Len(pTexto))
            Else
                Retorno(i) = Left(pTexto, pMaximo)
                pTexto = Mid(pTexto, Len(Retorno(i)) + 1, Len(pTexto))
    End If
    End If

        i = i + 1
    Loop

    If i > 1 Then
        ReDim Preserve Retorno(i - 1)
    Else
        ReDim Preserve Retorno(1)
    End If

    mfSplitEX = Retorno
End Function

Public Property Get PremioLiquidoLMI() As String
    Dim lretorno As String
    lretorno = Format(mStcA46V702B.VAL_PR, "standard")
    PremioLiquidoLMI = lretorno
End Property
Public Property Get RiscosCobertos() As String
    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If mStcA46V702B.M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(mStcA46V702B.M46V999.gDataCorteSompo) Then
            If CLng(mStcA46V702B.M46V999.gDataCorteSompo) <= mStcA46V702B.M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If
    
    Dim lretorno As String
    Select Case mStcA46V702B.Cod_Ramo
    Case e_CodRamo_Condominio
        lretorno = "Conforme Condições Gerais do Produto " & lNomeSeguradora & " Condomínio"
    Case e_CodRamo_Empresarial
        lretorno = "Conforme Condições Gerais do Produto " & lNomeSeguradora & " Empresarial"
    Case e_CodRamo_Residencial
        lretorno = "Conforme Condições Gerais do Produto " & lNomeSeguradora & " Residencial"
    End Select
    RiscosCobertos = lretorno
End Property
Public Property Get LocaisDeRisco() As String
    Dim lretorno As String
    lretorno = "(" & mStcA46V702B.ColecaoPedLoc.Count & " locais) conforme especificação anexa"
    LocaisDeRisco = lretorno
End Property
Public Property Get ValorEmRisco() As String
    Dim lretorno As String
    lretorno = " Conforme especificação anexa"
    ValorEmRisco = lretorno
End Property
Public Property Get Ocupacao() As String
    Dim lretorno As String
    lretorno = " Conforme especificação anexa"
    Ocupacao = lretorno
End Property

Public Property Get TipoEmissao() As String
    Dim lretorno As String
    Dim lPedLoc As stcA46V708B

    Set lPedLoc = mStcA46V702B.ColecaoPedLoc(1)
    lretorno = lPedLoc.getRpt.TipoEmissao

    lretorno = IIf(lretorno = "", " ", lretorno)
    Set lPedLoc = Nothing
    TipoEmissao = lretorno
End Property

Public Property Get BancoRestituicao() As String
    Dim lretorno As String

    If mStcA46V702B.M46V999.gfPreenchido(mStcA46V702B.gDesc_BancoRestituicao) Then
        lretorno = mStcA46V702B.gDesc_BancoRestituicao
    Else
        Dim codigo As Integer, Nome As String, nomeAbreviado As String, isCarne As Boolean, isDebito As Boolean, isRestituicao As Boolean, pMensagem As String
        lretorno = Empty
        
        If (mStcA46V702B.M46V999.GetBancoByID(mStcA46V702B.objstcCtrlEmis.Cod_Banco_Cre, codigo, Nome, nomeAbreviado, isCarne, isDebito, isRestituicao, pMensagem) = False) Then
            lretorno = ""
        Else
            mStcA46V702B.gDesc_BancoRestituicao = nomeAbreviado
            lretorno = nomeAbreviado
        End If
    
        'Select Case mStcA46V702B.objstcCtrlEmis.Cod_Banco_Cre
        '    Case 3:
        '        lretorno = "AMAZONIA S/A"
        '    Case 1:
        '        lretorno = "BANCO DO BRASIL"
        '    Case 634:
        '        lretorno = "BANCO TRIANGULO S/A"
        '    Case 610:
        '        lretorno = "BANCO VR"
        '    Case 756:
        '        lretorno = "BANCOOB"
        '    Case 29:
        '        lretorno = "BANERJ S/A"
        '    Case 47:
        '        lretorno = "BANESE"
        '    Case 320:
        '        lretorno = "BIC"
        '    Case 641:
        '        lretorno = "BILBAO VIZCAYA"
        '    Case 394:
        '        lretorno = "BMC"
        '    Case 479:
        '        lretorno = "BOSTON"
        '    Case 237:
        '        lretorno = "BRADESCO SA"
        '    Case 263:
        '        lretorno = "CACIQUE"
        '    Case 104:
        '        lretorno = "CAIXA ECONOMICA"
        '    Case 745:
        '        lretorno = "CITIBANK"
        '    Case 477:
        '        lretorno = "CITY BANK"
        '    Case 707:
        '        lretorno = "DAYCOVAL"
        '    Case 487:
        '        lretorno = "DEUTSCHE BANK SA"
        '    Case 21:
        '        lretorno = "ESPIRITO SANTO"
        '    Case 41:
        '        lretorno = "ESTADO DO R.G.SUL"
        '    Case 370:
        '        lretorno = "EUROPEU P/ AM. LATINA"
        '    Case 399:
        '        lretorno = "HSBC"
        '    Case 604:
        '        lretorno = "INDUSTRIAL DO BRASIL"
        '    Case 653:
        '        lretorno = "INDUSVAL"
        '    Case 341:
        '        lretorno = "ITAU"
        '    Case 376:
        '        lretorno = "LAR BRASILEIRO"
        '    Case 389:
        '        lretorno = "MERCANTIL DO BRASIL"
        '    Case 611:
        '        lretorno = "PAULISTA SA"
        '    Case 422:
        '        lretorno = "SAFRA S.A."
        '    Case 33:
        '        lretorno = "SANTANDER"
        '    Case 748:
        '        lretorno = "SICREDI"
        '    Case 464:
        '        lretorno = "SUMITOMO BRAS."
        '    Case 456:
        '        lretorno = "TOKYO-MITSUBISHI BR"
        'End Select
    End If

    If mStcA46V702B.objstcCtrlEmis.Cod_Banco_Cre <> 0 Then
        lretorno = mStcA46V702B.objstcCtrlEmis.Cod_Banco_Cre & " - " & lretorno
    End If

    lretorno = IIf(lretorno = "", " ", lretorno)
    BancoRestituicao = lretorno
End Property

Public Property Get AgenciaDigRestituicao() As String
    Dim lretorno As String
    Dim clsDigito As clsA99V001
    Dim lDigito As String
    Dim lMensagem As String

    If mStcA46V702B.objstcCtrlEmis.Cod_Banco_Cre <> 0 And mStcA46V702B.objstcCtrlEmis.Cod_Agenc_Cre <> 0 Then
        lretorno = mStcA46V702B.objstcCtrlEmis.AgenciaDig_Cre
    End If
    lretorno = IIf(lretorno = "", " ", lretorno)
    AgenciaDigRestituicao = lretorno

End Property

Public Property Get ContaCorrenteRestituicao() As String
    Dim lretorno As String

    If mStcA46V702B.objstcCtrlEmis.Cod_Banco_Cre <> 0 And mStcA46V702B.objstcCtrlEmis.Cod_Agenc_Cre <> 0 Then
        Call mpForConta1(mStcA46V702B.objstcCtrlEmis.Cod_Banco_Cre, mStcA46V702B.objstcCtrlEmis.Cod_Conta_Cre & IIf(UCase(mStcA46V702B.objstcCtrlEmis.Cod_DIG_Conta) = "X", "0", mStcA46V702B.objstcCtrlEmis.Cod_DIG_Conta), lretorno)
    End If
    lretorno = IIf(lretorno = "", " ", lretorno)
    ContaCorrenteRestituicao = lretorno

End Property
