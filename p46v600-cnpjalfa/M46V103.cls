VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsM46V103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================================================================================
' Class Name: clsM46V103
' Sistema: SyasMult
' Analista: Marcio Oliveira
' Objetivo: Substitui o antigo módulo bas Adaptado para ser utilizado como classe gerenciavel.
' ======================================================================================================================================================

'---   Dados fixos da proposta
Private lCod_Ramo As Integer
Private lCod_Modalidade As Integer
Private lCod_Tarifa As Integer
Private lCod_Sequencia As Integer
Private lCod_Corretor As Long
Private lCod_Clasf As String
Private lCod_Compl_Clasf As String
Private lCod_Construcao As Integer
Private lTip_Atividade As Integer
Private lNum_Cep As Long
Public lCod_Cobert As Integer
Private lVal_IS As Double
Private lVal_IS_Basica As Double
Private lVal_VR As Double
Private lVal_Franquia As Double
Private lTip_Emissao As e_TipEmissao
Private lCoePrazo As Single
Private lTip_Cobert As Integer

'1100192 - Tarifa 1006 e 1007 Marco 2011
Public lVal_VR_Item As Double


'---  Dados para Identificação dos Coeficientes
Private lCod_Classe As Integer
Private lClasseLMG As Integer
Private lCodFaixa_Reg As Integer
Private lCodClasse_Reg As Integer
Private lCoef_Regiao As Double                  'Coeficiente de Região
Private lCoef_LMG As Double                             'Coeficiente de LMG
Private lCoef_Franquia As Double            'Coeficiente de Franquia
Private lCoef_Protecao As Double            'Coeficiente de Proteção
Private lCoef_ProtecaoResseg As Double  'Coeficiente de Proteção somente para Resseguro
Private lPer_Coml As Double             'Percentual de Comissão do Corretor
Private lCoef_PI As Double              'Coeficiente Período Indenitário
Private lCoef_Desc_Prog As Double       'Desconto Progressivo (Tab_Ped_Loc)
Private lCoef_Bonus As Double           'Desconto de Experiência ou Bônus
Private lCoef_Desc_Coml As Double       'Desconto Comercial
Private lCoef_Desc_Coml_Congenere As Double       'Desconto Comercial Congenere
Private lCoefIs As Double
Private lClasseAPont As Integer
Private lTaxaAPont As Double

'1001146 - Tarifa 1000 e 1001 Dezembro 2010
'private lIndLMI              As Integer
Private lCodRegiao As Long

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
    e_Tip_Valor_ = 13                    'Utilizado antes de consistir cobol
    e_Tip_Valor_TaxaClasse = 14          'Utilizado na precificação
    e_Tip_Valor_Classe = 15              'Utilizado na precificação
    e_Tip_Valor_PercFranq = 16           'não utilizado
    e_Tip_Valor_ValMinFranq = 17         'não utilizado
    e_Tip_Valor_Segurado = 18            'Utilizado na precificação
    e_Tip_Valor_Regiao = 19              'Utilizado na franquia diferenciada
    e_Tip_Valor_CoeficienteTaxa = 20     'Utilizado na taxa para agravar ou descontar
    e_Tip_Valor_CoeficienteEspecialTaxa = 21     'Utilizado na taxa para agravar ou descontar
    e_Tip_Valor_FatorProtecao = 22               'Carrega automaticamente um Fator de Protecao
    e_Tip_Restricao_Corretor = 23                'Restrição de corretor
    e_Tip_Valor_DescComl_TipEmis = 24            'Desconto comercial diferente por Tipo de emissão
    e_Tip_Valor_Comissao_TipEmis = 25            'Comissão diferente por Tipo de Emissão
    e_Tip_Valor_ISFixa_Cobertura = 26            ' IS Fixa da Cobertura
    e_Tip_Valor_PremioMinimo = 27                ' Premio Minimo
    e_Tip_Valor_CoberturaLiberada = 28           ' CoberturaRestrita
    e_Tip_Valor_DescontoAgrupamento = 29         ' Desconto de Agrupamento
    e_Tip_Valor_ISLimite_Cobertura = 30          ' IS Máxima que pode ser contratada
    e_Tip_Valor_LimPercCobert = 31               ' Percentual de limite da cobertura
    e_Tip_Valor_JurosSegmentados = 32            ' Utilizada para parametrizar juros por segmentados
    e_Tip_Valor_ParcMaxCorretor = 33             ' Parametrização de parcelamento maximo por corretor
    e_Tip_Valor_ValidaDiasCorr = 36             ' Parametrização de validade de cotação / Proposta por corretor
    e_Tip_Valor_DiasParaVencimentoBoleto = 37    ' Utilizada para modificar o prazo de vencimento dos boletos
    e_Tip_Valor_ComissaoAss24hrs = 38            ' Comissão Assistência 24hrs
    e_Tip_Valor_VigenciaApolice = 39             ' Parametrização da vigência da apólice minima 1 mes; maxima 1 ano.
    e_Tip_Valor_CoeficienteDivisaoTaxa = 41      ' Utilizado na taxa para agravar ou descontar
End Enum

Public Enum e_Cod_Tip_Chave
    e_Cod_Tip_Chave_Val_IS = 1
    e_Cod_Tip_Chave_Val_Franquia = 2
    e_Cod_Tip_Chave_Num_CEP = 3
    e_Cod_Tip_Chave_Tip_Emissao = 4
    e_Cod_Tip_Chave_Comissao = 5
    e_Cod_Tip_Chave_Desc_Coml = 6
    e_Cod_Tip_Chave_Val_VR = 7
    e_Cod_Tip_Chave_Cod_Regiao = 8
    e_Cod_Tip_Chave_todos = 9  '??
    e_Cod_Tip_Chave_Fator_Ajuste = 10  'Desconto / Fator de Proteção
    e_Cod_Tip_Chave_Val_IS_Basica = 11
    e_Cod_Tip_Chave_RestricaoCorretor = 12
    e_Cod_Tip_Chave_Tip_EmissaoInicial = 13
    e_Cod_Tip_Chave_Qtde_Cob_Agrup = 14
    e_Cod_Tip_Chave_Juros_Segmentados = 15      ' Juros por segmentados
    e_Cod_Tip_ParcMaxCorretor = 16              ' Parametrização de parcelamento maximo por corretor
    e_Cod_Tip_ValidaDiasCorr = 17               ' Parametrização de validade de cotação / Proposta por corretor
    e_Cod_Tip_Vencimento_De_Boleto = 23         ' Muda vencimento dos boletos
        e_Cod_Tip_ClasseOcupacao = 24               ' Classe de ocupação
End Enum

Private mClsM46V111 As clsM46V111



'Public Sub BuscaItemEndosso(txtRenApolice As String, txtRenItem As String)
'    Dim lbdP00APOL          As ADODB.Connection 'Arquivo - P00APOL.
'    Dim lBonus              As Long
'    Dim lCodSeguradora      As Integer          'Código da seguradora.
'    Dim lIndSinistro        As Byte             'Indicador de sinistro.
'    Dim lNumItem            As Integer          'Número do item.
'    Dim lMensagem           As String           'Mensagem.
'    Dim rsAPOL              As clsYasRecordSet  'Registro: P00APOL.
'    Dim lI                  As Long             'Auxiliar.
'    Dim lobjstcClau         As stcA46V703B      'Objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU)
'    Dim ColecaoClau         As Collection       'Coleção de cláusulas.
'    Dim lobjstcCobert       As stcA46V704B      'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
'    Dim ColecaoCobert       As Collection       'Coleção de coberturas.
'    Dim lRegistro           As String           'Registro.
'    Dim lIdeTabela          As String
'    Dim lSelect             As String
'
'    Screen.MousePointer = vbHourglass
'
'
'    '2. Consiste número de apólice anterior.
'    If Not M46V999.gfPreenchido(txtRenApolice) Then
'        Call M46V999.gpGraLog(1, "FGE0069 - " & M46V999.gPrefixo & _
         '                         "Número da apólice inválido (não preenchido).")
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    If Not IsNumeric(txtRenApolice) Then
'        Call M46V999.gpGraLog(1, "FGE0069 - " & M46V999.gPrefixo & _
         '                         "Número da apólice inválido (não numérico).")
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    If Val(txtRenApolice) = 0 Then
'        Call M46V999.gpGraLog(1, "FGE0069 - " & M46V999.gPrefixo & _
         '                         "Número da apólice inválido (zerado).")
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    If Len(Trim(txtRenApolice)) < 10 Then
'        txtRenApolice = Trim$(txtRenApolice)
'        If Len(txtRenApolice) < 10 Then
'            txtRenApolice = String(10 - Len(txtRenApolice), "0") & txtRenApolice
'        End If
'    End If
'
'    '3. Consiste número do item da apólice.
'    If Not M46V999.gfPreenchido(txtRenItem) Then
'        Call M46V999.gpGraLog(1, "FGE0070 - " & M46V999.gPrefixo & _
         '                         "Número do item da apólice inválido ( não preenchido).")
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    If Not IsNumeric(txtRenItem) Then
'        Call M46V999.gpGraLog(1, "FGE0070 - " & M46V999.gPrefixo & _
         '                         "Número do item da apólice inválido (não numérico).")
''        If txtRenItem.Visible = True And txtRenItem.Enabled = True Then
''            txtRenItem.SetFocus
''        End If
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    If Val(txtRenItem) = 0 Then
'        Call M46V999.gpGraLog(1, "FGE0070 - " & M46V999.gPrefixo & _
         '                         "Número do item da apólice inválido (zerado).")
''        If txtRenItem.Visible = True And txtRenItem.Enabled = True Then
''            txtRenItem.SetFocus
''        End If
''        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    If Len(Trim(txtRenItem)) < 5 Then
'        txtRenItem = Trim$(txtRenItem)
'        If Len(txtRenItem) < 5 Then
'            txtRenItem = String(5 - Len(txtRenItem), "0") & txtRenItem
'        End If
'    End If
'
'    '4. Obtém apólice/item.
'    If M46V999.gfAbrBasDados("P00APOL", M46V999.gTipP00APOL, M46V999.gEndP00APOL, lbdP00APOL, lMensagem) = False Then
'        Call M46V999.gpGraLog(1, lMensagem)
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    If M46V999.gBaseLocal = True Then
'        lIdeTabela = "TAB_APOL_LOC"
'    Else
'        lIdeTabela = "GED.DBO.TAB_PED_LOC"
'    End If
'
'    lSelect = "SELECT * FROM " & lIdeTabela & _
     '            " WHERE " & lIdeTabela & ".Num_Apol = " & txtRenApolice & _
     '                " AND " & lIdeTabela & ".Num_Item = " & txtRenItem
'
'    If M46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) = False Then
'        Call M46V999.gpFecha2(lbdP00APOL)
'        Call M46V999.gpGraLog(1, lMensagem)
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    If rsAPOL.EOF = True Then
'        If M46V999.gBaseLocal = True Then
'            lIdeTabela = "TAB_APOL"
'        Else
'            lIdeTabela = "GED.DBO.TAB_PED"
'        End If
'
'        lSelect = "SELECT * from " & lIdeTabela & _
         '            " WHERE " & lIdeTabela & ".Num_Apol = " & txtRenApolice
'        If M46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) = False Then
'            Call M46V999.gpFecha2(lbdP00APOL)
'            Call M46V999.gpGraLog(1, lMensagem)
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
'        If rsAPOL.EOF = True Then
'            lMensagem = "FNL0008 - Apólice " & txtRenApolice & " não localizada."
'        Else
'            lMensagem = "FNL0009 - Item  " & txtRenItem & " não localizado na apólice " & _
             '                        txtRenApolice & "."
'        End If
'
'        'proposta eletrônica
'        mAntCNPJ = Format(Val(rsAPOL("NUM_CGC_CPF")), "00000000000000")
'        mAntSeg = rsAPOL("NOM_SEGURADO")
'
'        Call M46V999.gpFecha2(lbdP00APOL)
'        Call M46V999.gpGraLog(1, lMensagem)
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    gCarga = True
'    frmT46V101A.txtRisTipLogradouro = M46V999.gfTratarNulo(rsAPOL("Cod_Tip_Logr_Risco"), tpoString)
'    frmT46V101A.txtRisNomLogradouro = M46V999.gfTratarNulo(rsAPOL("Nom_Logr_Risco"), tpoString)
'    frmT46V101A.txtRisComplemento = M46V999.gfTratarNulo(rsAPOL("Dsc_Compl_Risco"), tpoString)
'    frmT46V101A.txtRisBairro = M46V999.gfTratarNulo(rsAPOL("Nom_Bairro_Risco"), tpoString)
'    frmT46V101A.txtRisCidade = M46V999.gfTratarNulo(rsAPOL("Nom_Cid_Risco"), tpoString)
'    frmT46V101A.cboRisUF.ListIndex = -1
'    If M46V999.gfPreenchido(M46V999.gfTratarNulo(rsAPOL("Sig_UF_Risco"), tpoString)) Then
'        For lI = 0 To (frmT46V101A.cboRisUF.ListCount - 1)
'            If M46V999.gfTratarNulo(rsAPOL("Sig_UF_Risco"), tpoString) = frmT46V101A.cboRisUF.List(lI) Then
'                frmT46V101A.cboRisUF.ListIndex = lI
'                Exit For
'            End If
'        Next lI
'    End If
'    frmT46V101A.txtRisCEP = M46V999.gfTratarNulo(rsAPOL("Num_CEP_Risco"), tpoNumerico)
'    frmT46V101A.txtDescProgr = Format(M46V999.gfTratarNulo(rsAPOL("Per_Desc_Progres"), TpoValor), "standard")
'    frmT46V101A.txtISBas = Format(M46V999.gfTratarNulo(rsAPOL("Per_IS_Vr"), TpoValor), "standard")
'    frmT46V101A.txtCoefISVR = Format(M46V999.gfTratarNulo(rsAPOL("Val_Coef_IS_Vr"), TpoValor), "standard")
'    frmT46V101A.txtValoRisco = Format(M46V999.gfTratarNulo(rsAPOL("Val_Vr"), TpoValor), "standard")
'    frmT46V101A.txtClassificacao.Tag = M46V999.gfTratarNulo(rsAPOL("Cod_Clasf"), tpoNumerico) & M46V999.gfTratarNulo(rsAPOL("Cod_Compl_Clasf"), tpoNumerico) & Format(M46V999.gfTratarNulo(rsAPOL("Tab_Clasf"), tpoNumerico), "0000")
'    frmT46V101A.txtClassificacao = M46V999.gfTratarNulo(rsAPOL("Cod_Clasf"), tpoNumerico) & "-" & M46V999.gfTratarNulo(rsAPOL("Dsc_Clasf"), tpoString)
''Ficha 1000305
'    If M46V999.gCod_Ramo = e_CodRamo_Empresarial Then
'        Select Case M46V999.gfTratarNulo(rsAPOL("Cod_Constr"), tpoNumerico)
'            Case e_CodContr_Superior
'                frmT46V101A.cboConstrucao.ListIndex = 1
'            Case e_CodContr_Inferior
'                frmT46V101A.cboConstrucao.ListIndex = 2
'            Case e_CodContr_Mista
'                frmT46V101A.cboConstrucao.ListIndex = 3
'            Case Else
'                frmT46V101A.cboConstrucao.ListIndex = 1
'        End Select
'    Else
'        If M46V999.gfTratarNulo(rsAPOL("Cod_Constr"), tpoNumerico) = e_CodContr_Superior Then
'            frmT46V101A.optContrucao(0).Value = True
'        Else
'            frmT46V101A.optContrucao(1).Value = True
'        End If
'    End If
'
''Ficha 1000305
''== Tipo Cobertura
'    If M46V999.gCod_Ramo = e_CodRamo_Empresarial Then
'        Select Case M46V999.gfTratarNulo(rsAPOL("Tip_Cobert"), tpoNumerico)
'            Case e_TipCobert_PredioeConteudo
'                frmT46V101A.cboTipCobert.ListIndex = 0
'            Case e_TipCobert_Predio
'                frmT46V101A.cboTipCobert.ListIndex = 1
'            Case e_TipCobert_Conteudo
'                frmT46V101A.cboTipCobert.ListIndex = 2
'            Case Else
'                frmT46V101A.cboTipCobert.ListIndex = 0
'        End Select
'    Else
'        frmT46V101A.cboTipCobert.ListIndex = 0
'    End If
'    lBonus = M46V999.gfTratarNulo(rsAPOL("Per_Bonus"), tpoNumerico)
'
'    If optSinistroYas(0).Value = True Then 'Com sinistro
'        lBonus = lBonus - 5
'    Else
'        lBonus = lBonus + 5
'    End If
'    If lBonus < 0 Then lBonus = 0
'
'    If pTipEmissao <> 300 Then
'        If M46V999.gCod_Ramo = e_CodRamo_Empresarial Then
'            If M46V999.gCodProduto >= 942 Then
'                'Soma mais 5 da fidelidade.
'                lBonus = lBonus + 5
'            End If
'            If lBonus > 30 Then lBonus = 30
'        ElseIf M46V999.gCod_Ramo = e_CodRamo_Residencial Then
'            If M46V999.gCodProduto >= 946 Then
'                'Soma mais 5 da fidelidade.
'                lBonus = lBonus + 5
'                If lBonus > 30 Then lBonus = 30
'            Else
'                If lBonus > 25 Then lBonus = 25
'            End If
'        Else
'            If lBonus > 25 Then lBonus = 25
'        End If
'
'        'Ficha 901673 e 901772  'Ficha 1000099 - Planos 2131 a 2138 Corretor Schroeder
'        If (gM46V111.gCod_Plano >= 2001 And gM46V111.gCod_Plano <= 2031) Or (gM46V111.gCod_Plano >= 2051 And gM46V111.gCod_Plano <= 2100) Or (gM46V111.gCod_Plano >= 2111 And gM46V111.gCod_Plano <= 2138) Then
'            lBonus = 0
'        End If
'    End If
'
'    'Será carregado no cálculo
'    frmT46V101A.txtDescBonus = lBonus
'
'    frmT46V101A.txtQuaEmpregados = M46V999.gfTratarNulo(rsAPOL("Qtd_Empreg"), tpoNumerico)
'
'            'Segurança
'    If M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_01"), tpoString) = "S" Or _
     '       M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_02"), tpoString) = "S" Or _
     '       M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_03"), tpoString) = "S" Or _
     '       M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_04"), tpoString) = "S" Or _
     '       M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_05"), tpoString) = "S" Then
'        frmT46V101A.optSeguranca(0).Value = True
'        If M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_01"), tpoString) = "S" Then
'            frmT46V128A.chkSeguranca(0).Value = vbChecked
'        Else
'            frmT46V128A.chkSeguranca(0).Value = vbUnchecked
'        End If
'        If M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_02"), tpoString) = "S" Then
'            frmT46V128A.chkSeguranca(1).Value = vbChecked
'        Else
'            frmT46V128A.chkSeguranca(1).Value = vbUnchecked
'        End If
'        If M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_03"), tpoString) = "S" Then
'            frmT46V128A.chkSeguranca(2).Value = vbChecked
'        Else
'            frmT46V128A.chkSeguranca(2).Value = vbUnchecked
'        End If
'        If M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_04"), tpoString) = "S" Then
'            frmT46V128A.chkSeguranca(3).Value = vbChecked
'        Else
'            frmT46V128A.chkSeguranca(3).Value = vbUnchecked
'        End If
'        If M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_05"), tpoString) = "S" Then
'            frmT46V128A.chkSeguranca(4).Value = vbChecked
'        Else
'            frmT46V128A.chkSeguranca(4).Value = vbUnchecked
'        End If
'    Else
'        frmT46V101A.optSeguranca(1).Value = True
'        Unload frmT46V128A
'    End If
'
'    If M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_01"), tpoString) = "S" Or _
     '       M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_02"), tpoString) = "S" Or _
     '       M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_03"), tpoString) = "S" Or _
     '       M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_04"), tpoString) = "S" Then
'        frmT46V101A.optProtecao(0).Value = True
'        If M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_01"), tpoString) = "S" Then
'            frmT46V126A.chkProtecao(0).Value = vbChecked
'        Else
'            frmT46V126A.chkProtecao(0).Value = vbUnchecked
'        End If
'        If M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_02"), tpoString) = "S" Then
'            frmT46V126A.chkProtecao(1).Value = vbChecked
'        Else
'            frmT46V126A.chkProtecao(1).Value = vbUnchecked
'        End If
'        If M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_03"), tpoString) = "S" Then
'            frmT46V126A.chkProtecao(2).Value = vbChecked
'        Else
'            frmT46V126A.chkProtecao(2).Value = vbUnchecked
'        End If
'        If M46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_04"), tpoString) = "S" Then
'            frmT46V126A.chkProtecao(3).Value = vbChecked
'        Else
'            frmT46V126A.chkProtecao(3).Value = vbUnchecked
'        End If
'    Else
'        frmT46V101A.optProtecao(1).Value = True
'        Unload frmT46V126A
'    End If
'
'    frmT46V101A.txtRisRegiao.Tag = Format(M46V999.gfTratarNulo(rsAPOL("Cod_Regiao"), tpoNumerico), "00000") & Format(M46V999.gfTratarNulo(rsAPOL("Cod_Produto"), tpoNumerico), "00000")
'    gM46V111.gCod_Regiao = Val(Right(frmT46V101A.txtRisRegiao.Tag, 4))
'    frmT46V101A.objstcPedLoc.Cod_Regiao = M46V999.gfTratarNulo(rsAPOL("Cod_Regiao"), tpoNumerico)
'    frmT46V101A.objstcPedLoc.Cod_Produto = M46V999.gfTratarNulo(rsAPOL("Cod_Produto"), tpoNumerico)
'    frmT46V101A.txtRisRegiao = frmT46V101A.objstcPedLoc.gDesc_Regiao
'
'    'Clausulas
'    If M46V999.gBaseLocal = True Then
'        lIdeTabela = "TAB_APOL_CLAU"
'    Else
'        lIdeTabela = "GED.DBO.TAB_PED_CLAU"
'    End If
'
'    lSelect = " Select * from " & lIdeTabela & " Where Num_Apol = " & rsAPOL("Num_Apol")
'    lSelect = lSelect & " and num_item = " & rsAPOL("Num_Item")
'
'    If Not M46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) Then
'        Exit Sub
'    End If
'    'Clausula do beneficiário.
'    If Not rsAPOL.EOF Then
'        If rsAPOL("Cod_Clau") = 4 Or rsAPOL("Cod_Clau") = 1 Then
'            frmT46V101A.optBeneficiario(0).Value = True
'        Else
'            frmT46V101A.optBeneficiario(1).Value = True
'        End If
'    End If
'
'    If M46V999.gBaseLocal = True Then
'        lIdeTabela = "TAB_APOL_COBERT"
'    Else
'        lIdeTabela = "GED.DBO.TAB_PED_COBERT"
'    End If
'
'    'Coberturas
'    lSelect = " Select * from " & lIdeTabela & " Where Num_Apol = " & txtRenApolice
'    lSelect = lSelect & " and num_item = " & txtRenItem
'
'    If Not M46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) Then
'        Exit Sub
'    End If
'    Set ColecaoCobert = New Collection
'    Do While Not rsAPOL.EOF
'        Set lobjstcCobert = New stcA46V704B
'        Set lobjstcCobert.M46V999 = M46V999
'
'        With lobjstcCobert
'            .Cod_Cobert = M46V999.gfTratarNulo(rsAPOL("Cod_Cobert"), tpoNumerico)
'            .Per_Desc_Coml = M46V999.gfTratarNulo(rsAPOL("Per_Desc_Coml"), tpoNumerico)
'            .Per_Desc_Exp = M46V999.gfTratarNulo(rsAPOL("Per_Desc_Exp"), tpoNumerico)
'        End With
'        ColecaoCobert.Add lobjstcCobert, lobjstcCobert.Class_ID
'        rsAPOL.MoveNext
'    Loop
'    'Cobertura
'    Unload frmT46V117A
'    If ColecaoCobert.Count > 0 Then
'        Call frmT46V117A.mpSelecionarCoberturas(ColecaoCobert)
'    End If
'
'    Set ColecaoClau = Nothing
'    Set ColecaoCobert = Nothing
'    gCarga = False
'    Call gpHabSalvar
'    gAltT46V112A = True
'
'    '7. Reposiciona informações do item e retorna ao form de item.
'    Me.Hide
'    Screen.MousePointer = vbDefault
'End Sub


Public Function BuscaCoef_IS_VR_LC(ByVal pobjstcItem As stcA46V708B, _
                                   ByRef pCoef_LMG_LC As Double, _
                                   ByRef pMensagem As String) As Boolean

    Dim lbdTbm_Coef_LMG_VR As ADODB.Connection         'Arquivo: P00Tarifa.
    Dim lrsTbm_Coef_LMG_VR As clsYasRecordSet          'Registro : Tbm_Fator_LMG.
    Dim lSelect As String

    Dim Per_IS_Vr As Double



    BuscaCoef_IS_VR_LC = True

    If lCod_Ramo <> 112 Then
        pCoef_LMG_LC = 1
        Exit Function
    End If
    
    
    If pobjstcItem.Val_Vr_Lc <= CDbl(5000000) And pobjstcItem.gISTotal_LC <= CDbl(5000000) Then
        pCoef_LMG_LC = 1
        Exit Function
    End If
    
    If Trim(lCod_Clasf) = "042" Or _
         Trim(lCod_Clasf) = "043" Or _
         Trim(lCod_Clasf) = "045" Or _
         Trim(lCod_Clasf) = "510" Or _
         Trim(lCod_Clasf) = "530" Or _
         Trim(lCod_Clasf) = "531" Then
         
             lCoefIs = 1
         Exit Function
     End If

    If pobjstcItem.gISTotal_LC = 0 And pobjstcItem.Val_Vr_Lc = 0 Then
        pCoef_LMG_LC = 1
        Exit Function
    End If

    If pobjstcItem.Val_Vr_Lc = 0 Then
        pMensagem = "Valor do Risco não pode ser 0!"
        BuscaCoef_IS_VR_LC = False
        Exit Function
    End If

    Per_IS_Vr = pobjstcItem.gISTotal_LC / pobjstcItem.Val_Vr_Lc

    If Per_IS_Vr > 1 Then
        pMensagem = "Valor do LMG de Lucros Cessantes não pode ser maior que Valor do Risco de Lucros Cessantes!"
        BuscaCoef_IS_VR_LC = False
        Exit Function
    ElseIf Per_IS_Vr < 0.05 Then
        pMensagem = "Valor do LMG não pode ser menor que 5% do Valor de Lucros Cessantes!"
        BuscaCoef_IS_VR_LC = False
        Exit Function
    End If

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Coef_LMG_VR, pMensagem) = False Then
        BuscaCoef_IS_VR_LC = False
        Exit Function
    End If

    Per_IS_Vr = Per_IS_Vr * 100

    lSelect = " Select top 1 * "
    lSelect = lSelect & " from Tbm_Coef_LMG_VR WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
    lSelect = lSelect & " and Cod_Grupo_Cobert = " & 3
    lSelect = lSelect & " and Perc_LMI_VR <= " & Replace(Format(Per_IS_Vr, "#0.00"), ",", ".")
    lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
    lSelect = lSelect & " and Cod_Sequencia <= " & lCod_Sequencia
    lSelect = lSelect & " order by Perc_LMI_VR desc "


    If M46V999.gfObtRegistro(lbdTbm_Coef_LMG_VR, lSelect, lrsTbm_Coef_LMG_VR, pMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Coef_LMG_VR)
        Exit Function
    End If


    If lrsTbm_Coef_LMG_VR.EOF Then
        pCoef_LMG_LC = 1
        Call M46V999.gpFecha2(lbdTbm_Coef_LMG_VR)
        Exit Function
    End If


    If Not lrsTbm_Coef_LMG_VR.EOF Then
        pCoef_LMG_LC = M46V999.Check(lrsTbm_Coef_LMG_VR("Coef_Agrav"))
    Else
        pMensagem = "Não encontrado o Coeficiente do IS/VR Lucros Cessantes"
        BuscaCoef_IS_VR_LC = False
    End If


    Call M46V999.gpFecha2(lbdTbm_Coef_LMG_VR)
    Set lrsTbm_Coef_LMG_VR = Nothing
    Exit Function
End Function


Public Property Get Class_Name() As String
    Class_Name = "M46V103"
End Property

Public Function BuscaCoef_IS_VR(ByVal pPropTabela As String, _
                                ByVal pNomeTabImport As String, _
                                ByRef pobjstcItem As stcA46V708B, _
                                ByVal pCod_Grupo_Cobert As Integer, _
                                ByVal pCod_Cobert As Integer, _
                                ByVal pVal_IS As Double, _
                                ByVal pVal_Vr As Double, _
                                ByRef pMensagem As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "BuscaCoef_IS_VR"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lbdTbm_Coef_LMG_VR As ADODB.Connection         'Arquivo: P00Tarifa.
    Dim lrsTbm_Coef_LMG_VR As clsYasRecordSet          'Registro : Tbm_Fator_LMG.
    Dim lCodClau As Integer
    Dim lSelect As String

    BuscaCoef_IS_VR = True

    If lCod_Ramo = 112 Then
        If M46V999.gCodProduto >= 1008 Then
            Select Case pobjstcItem.Cod_Produto
            Case 1020, 1030, 1040, 1050, 1060, 1010
                lCoefIs = 1
                lCodClau = 1
                GoSub GravaCodClau
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
                Exit Function
            End Select
        ElseIf M46V999.gCodProduto >= 1004 Then
            If InStr(1, "1020,  1030, 1040,1050, 1060 ,1010", pobjstcItem.Cod_Produto) <> 0 Then
                lCoefIs = 1
                lCodClau = 1
                GoSub GravaCodClau
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
                Exit Function
            End If
        End If
    End If
    If lCod_Ramo = 112 Then
        If M46V999.gCodProduto >= 1080 Then
            If Trim(lCod_Clasf) = "042" Or _
                Trim(lCod_Clasf) = "043" Or _
                Trim(lCod_Clasf) = "045" Or _
                Trim(lCod_Clasf) = "510" Or _
                Trim(lCod_Clasf) = "530" Or _
                Trim(lCod_Clasf) = "531" Then
                
                    lCoefIs = 1
                    lCodClau = 1
                    GoSub GravaCodClau
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
                Exit Function
            End If
        Else
            If lCod_Clasf = M46V999.gfBuscarDePara("197", 2) Or lCod_Clasf = M46V999.gfBuscarDePara("134", 2) Then
                lCoefIs = 1
                lCodClau = 1
                GoSub GravaCodClau
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
                Exit Function
                End If
        End If
    End If
    If lCod_Ramo = 113 Then
        lCoefIs = 1
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
        Exit Function
    End If
    If lCod_Ramo = 114 Then
        If lCod_Compl_Clasf = "11" Or lCod_Compl_Clasf = "11" Or lCod_Compl_Clasf = "14" Or lCod_Compl_Clasf = "15" Then
            lCoefIs = 1
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
            Exit Function
        End If
        If M46V999.gCodProduto >= 1012 Then
            lCoefIs = 1
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
            Exit Function
        End If
    End If

    If pVal_Vr = 0 Then
        If M46V111.gCod_Dizer <> 0 Then
            lCoefIs = 1
            lCodClau = 1
            GoSub GravaCodClau
            Exit Function
        End If
        pMensagem = "Cobertura:" & pCod_Cobert & " - Valor do Risco deve estar preenchido!"
        BuscaCoef_IS_VR = False
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If


    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Coef_LMG_VR, pMensagem) = False Then
        BuscaCoef_IS_VR = False
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    If pCod_Grupo_Cobert = 1 Then
        'BRFoods
        If lCod_Modalidade = 2218 Or lCod_Modalidade = 2219 Or (lCod_Modalidade >= 2228 And lCod_Modalidade <= 2271) Then
            ' Não fazer nada
        Else
            If pVal_Vr < CDbl(2000000) Then
                lCoefIs = 1
                lCodClau = 1
                GoSub GravaCodClau
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, "Val_VR < 2.000.000", 2, Class_Name, methodName)
                Exit Function
            End If
        End If
    Else
        lSelect = " Select * "
        lSelect = lSelect & " from Tbm_Cob_LucroCess WITH (NOLOCK) "
        lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
        lSelect = lSelect & " and Cod_Cobert = " & pCod_Cobert
        lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
        lSelect = lSelect & " and Cod_Sequencia <= " & lCod_Sequencia

        If M46V999.gfObtRegistro(lbdTbm_Coef_LMG_VR, lSelect, lrsTbm_Coef_LMG_VR, pMensagem) = False Then
            Call M46V999.gpFecha2(lbdTbm_Coef_LMG_VR)
            BuscaCoef_IS_VR = False
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            Exit Function
        End If

        If lrsTbm_Coef_LMG_VR.EOF Then
            lCoefIs = 1

            Call M46V999.gpFecha2(lbdTbm_Coef_LMG_VR)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
            Exit Function
        End If

    End If

    lCodClau = 2
    GoSub GravaCodClau

    Dim Per_IS_Vr As Double

    Per_IS_Vr = (pVal_IS / pVal_Vr) * 100

    If Per_IS_Vr > 100 Then
        pMensagem = "Item " & pobjstcItem.Num_Item & " Cobertura: " & pCod_Cobert & " - Valor do LMG não pode ser maior que Valor do Risco!"
        BuscaCoef_IS_VR = False
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    ElseIf Per_IS_Vr < 40 And pCod_Grupo_Cobert = 1 Then
        pMensagem = "Item " & pobjstcItem.Num_Item & " Cobertura: " & pCod_Cobert & " - Valor do LMG não pode ser menor que 40% do Valor do Risco!"
        BuscaCoef_IS_VR = False
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    ElseIf Per_IS_Vr < 5 And pCod_Grupo_Cobert <> 1 Then
        pMensagem = "Item " & pobjstcItem.Num_Item & " Cobertura: " & pCod_Cobert & " - Valor do LMG não pode ser menor que 5% do Valor de Lucros Cessantes!"
        BuscaCoef_IS_VR = False
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    lSelect = " Select * "
    lSelect = lSelect & " from Tbm_Coef_LMG_VR WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
    lSelect = lSelect & " and Cod_grupo_cobert = " & pCod_Grupo_Cobert
    lSelect = lSelect & " and Perc_LMI_VR <= " & Replace(Format(Per_IS_Vr, "00.00"), ",", ".")
    lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
    lSelect = lSelect & " and Cod_Sequencia <= " & lCod_Sequencia
    lSelect = lSelect & " order by Perc_LMI_VR desc"

    If M46V999.gfObtRegistro(lbdTbm_Coef_LMG_VR, lSelect, lrsTbm_Coef_LMG_VR, pMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Coef_LMG_VR)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    If Not lrsTbm_Coef_LMG_VR.EOF Then
        lCoefIs = M46V999.Check(lrsTbm_Coef_LMG_VR("Coef_Agrav"))
    Else
        pMensagem = "Cobertura:" & pCod_Cobert & " - Valor em Risco Inválido!"
        BuscaCoef_IS_VR = False
    End If


    Call M46V999.gpFecha2(lbdTbm_Coef_LMG_VR)
    Set lrsTbm_Coef_LMG_VR = Nothing
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function

GravaCodClau:

    If pCod_Grupo_Cobert = 1 Then
        '801511 - Favor providenciar a desativação das cláusulas 001 e 002  dos Seguros Empresariais para que não sejam mais incluídas/informadas nas emissões de apólices.
        If lCod_Tarifa >= 948 And lCod_Ramo = 112 Then GoTo NaoIncluir
        If Not mfIncluirClausula(pPropTabela, pNomeTabImport, lCodClau, pobjstcItem, pMensagem, tagDynaTrace) Then
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
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
            If lobjstcFator.Cod_Fator = 9 Then    'RISCO EM SHOPPING CENTER
                fAchouShopping = True
                Exit For
            End If
        Next lobjstcFator
        Set lobjstcFator = Nothing
    End If
End Function

Public Function mfCalcPrazo(ByVal pDat_Ini_Vig As Long, ByVal pDat_Ter_Vig As Long, ByVal pCod_Iden_Praz As Integer, ByRef pCoePrazo As Single, pMensagem As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfCalcPrazo"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lQtdiasCorridos As Integer
    Dim lVigMeses As Integer  'Quantidade de meses de vigência.
    Dim lAuxiliar As String   'Auxiliar
    Dim lSelect

    Dim mVigIniAno As Integer                   'Data de início de vigência: ano.
    Dim mVigIniAnoMes As Long                   'Data de início de vigência: ano/mes no formato = aaaamm.
    Dim mVigIniDia As Integer                   'Data de início de vigência: dia.
    Dim mVigIniMes As Integer                   'Data de início de vigência: mês.
    Dim mVigTerAno As Integer                   'Data de término de vigência: ano.
    Dim mVigTerAnoMes As Long                   'Data de término de vigência: ano/mes no formato = aaaamm.
    Dim mVigTerDia As Long                              'Data de término de vigência: dia.
    Dim mVigTerMes As Integer                   'Data de término de vigência: mês.

    Dim lbdTab_Prazo As ADODB.Connection         'Arquivo: P00Tarifa.
    Dim lrsTab_Prazo As clsYasRecordSet          'Registro : Tab_Prazo.

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
        Do Until val(Mid(mVigIniAnoMes, 1, 4)) = val(Mid(mVigTerAnoMes, 1, 4))
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
        'Call M46V999.gpGraLog(0, "#035/004 - Calcula quantidade de meses de vigência (mfCalcPrazo)." & vbLf & _
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
            Call M46V999.gpGraLog(2, "FGE0009 - Prazo longo não selecionado para período de vigência maior que 1 " & _
                                     "ano (" & lVigMeses & " meses).")
        End If
        lQtdiasCorridos = mfDifDatas(pDat_Ini_Vig, pDat_Ter_Vig)
        'c) Obtém coeficiente de prazo.
        Select Case pCod_Iden_Praz
        Case 1
            ' If lTip_Emissao = 300 Or lTip_Emissao = 301 Or lTip_Emissao = 303 Then
            '     pCoePrazo = 1
            ' Else
            pCoePrazo = M46V999.gfTrunca(lQtdiasCorridos / 365, 5)
            ' End If
        Case 2
            pMensagem = ""
Voltar:
            lSelect = "SELECT Val_Coef FROM TAB_PRAZO_CURTO WITH (NOLOCK) WHERE Qtd_Dias_Ini <= " & lQtdiasCorridos & _
                    " AND Qtd_Dias_Fim >= " & lQtdiasCorridos

            If pMensagem = "" Then
                lSelect = lSelect & " and Posicao = 1"
            End If

            If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTab_Prazo, pMensagem) = False Then
                Call M46V999.gpGraLog(2, pMensagem)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                Exit Function
            End If

            If M46V999.gfObtRegistro(lbdTab_Prazo, lSelect, lrsTab_Prazo, pMensagem) = False Then
                If InStr(1, pMensagem, "-2147217904") > 0 Then
                    GoTo Voltar
                Else
                    Call M46V999.gpGraLog(2, pMensagem)
                    pCoePrazo = 0
                End If
            Else
                If lrsTab_Prazo.EOF = True Then
                    Call M46V999.gpGraLog(2, "FNL0019 - Coeficiente de prazo curto não localizado (quantidade " & _
                                             "de dias de vigência = " & lQtdiasCorridos & ").")
                    pCoePrazo = 0
                Else
                    pCoePrazo = lrsTab_Prazo!Val_Coef / 100
                End If
                Call M46V999.gpFecha2(lbdTab_Prazo)
            End If
        Case 3
            If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTab_Prazo, pMensagem) = False Then
                Call M46V999.gpGraLog(2, pMensagem)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                Exit Function
            End If

            lSelect = "SELECT Val_Coef FROM TAB_PRAZO_LONGO WITH (NOLOCK) WHERE Qtd_Meses = " & lVigMeses
            If M46V999.gfObtRegistro(lbdTab_Prazo, lSelect, lrsTab_Prazo, pMensagem) = False Then
                Call M46V999.gpGraLog(2, pMensagem)
                pCoePrazo = 0
            Else
                If lrsTab_Prazo.EOF = True Then
                    Call M46V999.gpGraLog(2, "FNL0025 - Coeficiente de prazo longo não localizado (quantidade " & _
                                             "de meses = " & lVigMeses & ").")
                    pCoePrazo = 0
                Else
                    pCoePrazo = lrsTab_Prazo!Val_Coef
                End If
                Call M46V999.gpFecha2(lbdTab_Prazo)
            End If
        Case Else
            pCoePrazo = 0
        End Select
        'Call M46V999.gpGraLog(0, "#035/005 - Obtém coeficiente de prazo (mpCalculo)." & vbLf & vbLf & _
         '                "Início de vigência = " & pDat_Ini_Vig & vbLf & _
         '                "Término de vigência = " & pDat_Ter_Vig & vbLf & _
         '                "Quantidade calculada de meses = " & lVigMeses & vbLf & _
         '                "Quantidade calculada de dias = " & lQtdiasCorridos & vbLf & _
         '                "Indicador de prazo = " & pCod_Iden_Praz & vbLf & _
         '                "Coeficiente de prazo = " & pCoePrazo)
    End If
    Set lbdTab_Prazo = Nothing
    Set lrsTab_Prazo = Nothing
    mfCalcPrazo = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
End Function


Public Function mfClasseAceit(ByRef pobjstcItem As stcA46V708B, _
                            ByRef pobjstcPedido As stcA46V702B, _
                            ByRef pMensagem As String, tagDynaTrace As String) As Integer
                                                          'ByVal pObjCobert As clsA46V704A,
    
    ' Valor padrão:
    ' pMensagem = ""
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfClasseAceit"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lobjstcCobertura As stcA46V704B
    Dim lobjstcCorretor As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    Dim lCod_Aceitacao As Integer                       'Código de Aceitação
    Dim lCod_Aceitacao_Aux As Integer

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
            If mfAjustesPontuais(e_Tip_Valor_ClasAceit, TpoBoolean, lCod_Aceitacao_Aux, pMensagem, Nothing, 0, tagDynaTrace) = False Then
                Call M46V999.gpGraLog(2, pMensagem)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
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
    M46V999.gTipCobertura = pobjstcItem.Tip_Cobert
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcCobertura = Nothing
    Set lobjstcCorretor = Nothing
    Return
End Function

Public Function mfConsisteComissao(ByVal pobjstcItem As stcA46V708B, _
                                    ByVal pobjstcPedido As stcA46V702B, tagDynaTrace As String) As Boolean

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfConsisteComissao"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lobjstcCobertura As stcA46V704B
    Dim lobjstcCorretor As stcA46V705B          'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    Dim lMensagem As String
    Dim lCod_Aceitacao As Integer                       'Código de Aceitação
    Dim lCod_Aceitacao_Aux As Integer

    Dim lCod_Corretor1 As Long
    Dim lPer_Coml1 As Double
    Dim lCod_Corretor2 As Long
    Dim lPer_Coml2 As Double
    Dim lCod_Corretor3 As Long
    Dim lPer_Coml3 As Double
    Dim lPer_ComlTot As Double

    Dim lCod_Comissao_Aux

    'aqui

    mfConsisteComissao = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
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

            If mfAjustesPontuais(e_Tip_Valor_Comissao, TpoValor, lCod_Comissao_Aux, lMensagem, Nothing, 0, tagDynaTrace) = False Then
                Call M46V999.gpGraLog(2, lMensagem)
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
    M46V999.gTipCobertura = pobjstcItem.Tip_Cobert

    Set lobjstcCobertura = Nothing
    Set lobjstcCorretor = Nothing
    mfConsisteComissao = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

End Function


Public Function mfDifDatas(ByVal pInicio As Long, ByVal pTermino As Long) As Single
'Função #014: calcula diferença entre datas (retorno em dias).

'Parâmetros de entrada...pInicio....data de início.
'                        pTermino...data de término.

    Dim lAuxIniDias As String   'Auxiliar: valores para tabela para cálculo da diferença de dias.
    Dim lAuxTerDias As String   'Auxiliar: valores para tabela para cálculo da diferença de dias.
    Dim lDia As Single   'Índice auxiliar: dia.
    Dim lI As Integer  'Índice.
    Dim lIniDia As Single   'Dia do início de vigência.
    Dim lIniMes As Single   'Mês do início de vigência.
    Dim lMes As Single   'Índice auxiliar: mês.
    Dim lPrazoLongo As Integer  'Quantidade de dias para prazo longo.
    Dim lQuaIniDias As Single   'Quantidade de dias (início de vigência).
    Dim lQuaTerDias As Single   'Quantidade de dias (término de vigência).
    Dim lTabIniDias(31, 12) As String   'Tabela para cálculo de dias (data de início).
    Dim lTabTerDias(31, 12) As String   'Tabela para cálculo de dias (data de término).
    Dim lTerDia As Single             'Dia do término de vigência.
    Dim lTerMes As Single             'Mês do término de vigência.

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
            lTabIniDias(lDia, lMes) = val(Mid(lAuxIniDias, lI, 3))
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
            lTabTerDias(lDia, lMes) = val(Mid(lAuxTerDias, lI, 3))
            lMes = lMes + 1
            lI = lI + 3
        Loop
        lDia = lDia + 1
    Loop

    '4. Apropria informação de dia e mês de início e término.
    lIniDia = val(Mid(pInicio, 7, 2))
    lIniMes = val(Mid(pInicio, 5, 2))
    lTerDia = val(Mid(pTermino, 7, 2))
    lTerMes = val(Mid(pTermino, 5, 2))

    If val(Mid(pTermino, 1, 4)) - val(Mid(pInicio, 1, 4)) > 2 Then
        lPrazoLongo = 730
    ElseIf val(Mid(pTermino, 1, 4)) - val(Mid(pInicio, 1, 4)) > 1 Then
        lPrazoLongo = 365
    End If

    '5. Posiciona quantidade de dias.
    lQuaIniDias = val(lTabIniDias(lIniDia, lIniMes))
    If val(Mid(pInicio, 1, 4)) <> val(Mid(pTermino, 1, 4)) Then
        lQuaTerDias = val(lTabTerDias(lTerDia, lTerMes)) + lPrazoLongo
    Else
        lQuaTerDias = val(lTabIniDias(lTerDia, lTerMes)) + lPrazoLongo
    End If

    '6. Calcula diferença de dias.
    mfDifDatas = lQuaTerDias - lQuaIniDias

    'Call M46V999.gpGraLog(0, "#014/001 - Calcula diferença de datas (mfDifDatas)." & vbLf & vbLf & _
     '                "Data de início = " & pInicio & vbLf & _
     '                "Data de término = " & pTermino & vbLf & _
     '                "Quantidade de dias da data de início (lQuaIniDias) = " & lQuaIniDias & vbLf & _
     '                "Quantidade de dias da data de término (lQuaTerDias) = " & lQuaTerDias & vbLf & _
     '                "Quantidade calculada de dias = " & mfDifDatas)
End Function

Public Sub mpCalcula_Endosso(pNum_Apol_Anterior As Double, pNum_Item_Anterior As Integer, pCod_Tip_Calc_Ws As String, ByRef pVal_pr_ws As Single)
    Dim lSql As String
    Dim lrst As clsYasRecordSet
    Dim lbdP00APOL As ADODB.Connection
    Dim lNomeTab As String
    Dim lI As Integer
    Dim lUltimoErro As String

    If M46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
    Else
        lNomeTab = "GED.DBO.TAB_PED"
    End If

    If M46V111.gCod_Dizer = 4 Or M46V111.gCod_Dizer = 10 Then
        pVal_pr_ws = 0
    End If

    If M46V999.gfAbrBasDados("P00APOL", M46V999.gTipP00APOL, M46V999.gEndP00APOL, lbdP00APOL, lUltimoErro) Then
        lSql = " Select  Val_pr_anu, Val_pr_tarif_anu,  Val_Pr_Resg_Anu from " & lNomeTab & "_Cobert WITH (NOLOCK) where num_apol = " & pNum_Apol_Anterior
        lSql = lSql & " and num_item = " & pNum_Item_Anterior
        lSql = lSql & " and cod_cobert = " & lCod_Cobert
        If M46V999.gfObtRegistro(lbdP00APOL, lSql, lrst, lUltimoErro) Then
            If lrst.EOF Then
                Call M46V999.gpGraLog(0, "O registro de Bonus em 'mfRetornaBonusApolice' para apólice: " & pNum_Apol_Anterior & " não foi encontrado-" & lUltimoErro)
                Exit Sub
            End If
        End If
    Else
        Call M46V999.gpGraLog(0, "Erro - Abertura do Banco P00APOL em 'mfRetornaBonusApolice' para apólice: " & pNum_Apol_Anterior & "-" & lUltimoErro)
        Exit Sub
    End If


    If pCod_Tip_Calc_Ws = "COML" Then
        pVal_pr_ws = pVal_pr_ws - M46V999.Check(lrst("Val_pr_anu"))
    Else
        If pCod_Tip_Calc_Ws = "TARI" Then
            pVal_pr_ws = pVal_pr_ws - M46V999.Check(lrst("Val_pr_tarif_anu"))
        Else
            If pCod_Tip_Calc_Ws = "RESG" Then
                pVal_pr_ws = pVal_pr_ws - M46V999.Check(lrst("Val_pr_resg_anu"))
            End If
        End If
    End If

    Call M46V999.gpFecha2(lbdP00APOL)

    If pVal_pr_ws > -0.1 And pVal_pr_ws < 8 Then
        pVal_pr_ws = 0
    End If
    If M46V111.gCod_Dizer = 6 Then
        pVal_pr_ws = 0
    End If
    '''   pVal_pr_ws = pVal_pr_ws * lCoePrazo

    Set lrst = Nothing
    Set lbdP00APOL = Nothing

End Sub

Public Function mfAtualizarFranquiasLMI(ByRef objstcPedido As stcA46V702B)
    Dim lColecaoTemp As Collection
    Dim lobjstcItem As stcA46V708B              'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcCobert As stcA46V704B    'Objeto da classe (stc) de coberturas.
    Dim TemVideObs As String

    Set lColecaoTemp = objstcPedido.ColecaoPedLoc

    ' Acumular em arrays os maiores valores de franquia e o * do Texto Franquia.
    Dim arrPerFranq() As Double
    Dim arrValFranqMin() As Double
    Dim arrTextoFranquiaAst() As Boolean
    Dim arrCodCobertFranq() As Integer
    Dim incr As Integer
    Dim achouCobert As Boolean

    ReDim arrPerFranq(0)
    ReDim arrValFranqMin(0)
    ReDim arrTextoFranquiaAst(0)
    ReDim arrCodCobertFranq(0)

    If M46V999.gCodProduto >= 1060 And objstcPedido.LMI_Unico = 1 Then
        For Each lobjstcItem In lColecaoTemp
            For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
                ' Para cada cobertura de cada item, verificar se o array já possui esta cobertura
                achouCobert = False
                For incr = 0 To UBound(arrCodCobertFranq)
                    If arrCodCobertFranq(incr) = lobjstcCobert.Cod_Cobert Then
                        achouCobert = True
                        Exit For
                    End If
                Next incr

                ' Se não achou, incluir esta cobertura no array
                If achouCobert = False Then
                    'incr = incr + 1
                    ReDim Preserve arrPerFranq(incr)
                    ReDim Preserve arrValFranqMin(incr)
                    ReDim Preserve arrTextoFranquiaAst(incr)
                    ReDim Preserve arrCodCobertFranq(incr)

                    arrPerFranq(incr) = lobjstcCobert.Per_Franq
                    arrValFranqMin(incr) = lobjstcCobert.Val_Franq_Min
                    arrTextoFranquiaAst(incr) = IIf(Left(lobjstcCobert.TextoFranquia, 1) = "*", True, False)
                    arrCodCobertFranq(incr) = lobjstcCobert.Cod_Cobert
                Else    ' Se achou, então definir os maiores valores para esta cobertura
                    If arrPerFranq(incr) < lobjstcCobert.Per_Franq Then _
                       arrPerFranq(incr) = lobjstcCobert.Per_Franq

                    If arrValFranqMin(incr) < lobjstcCobert.Val_Franq_Min Then _
                       arrValFranqMin(incr) = lobjstcCobert.Val_Franq_Min

                    If arrTextoFranquiaAst(incr) = False Then _
                       arrTextoFranquiaAst(incr) = IIf(Left(lobjstcCobert.TextoFranquia, 1) = "*", True, False)

                End If
            Next lobjstcCobert
        Next lobjstcItem

        ' Agora que já tem o array populado, definir os valores em todas as coberturas de todos os itens.
        For Each lobjstcItem In lColecaoTemp
            For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
                ' Para cada cobertura de cada item, localizar a cobertura no array e definir os valores
                For incr = 0 To UBound(arrCodCobertFranq)
                    If arrCodCobertFranq(incr) = lobjstcCobert.Cod_Cobert Then

                        If lobjstcCobert.Per_Franq < arrPerFranq(incr) _
                           Or lobjstcCobert.Val_Franq_Min < arrValFranqMin(incr) Then

                            lobjstcCobert.Per_Franq = arrPerFranq(incr)
                            lobjstcCobert.Val_Franq_Min = arrValFranqMin(incr)

                            If Left(lobjstcCobert.TextoFranquia, 1) <> "*" Then
                                TemVideObs = "*"
                            Else
                                TemVideObs = ""
                            End If

                            lobjstcCobert.TextoFranquia = M46V999.gfTextoFranquia( _
                                                          lobjstcCobert.Val_Franq_Min, _
                                                          lobjstcCobert.Per_Franq, _
                                                          TemVideObs, _
                                                          lobjstcCobert.Cod_Cobert, _
                                                          lobjstcCobert.Num_Meses_PI, _
                                                          lobjstcCobert.Val_IS, _
                                                          lobjstcItem)
                        End If

                        If Left(lobjstcCobert.TextoFranquia, 1) <> "*" And arrTextoFranquiaAst(incr) = True Then _
                           lobjstcCobert.TextoFranquia = "*" & lobjstcCobert.TextoFranquia

                        Exit For
                    End If
                Next incr
            Next lobjstcCobert
        Next lobjstcItem
    End If

    Set objstcPedido.ColecaoPedLoc = lColecaoTemp

FechaObj:
    Set lColecaoTemp = Nothing
    Set lobjstcItem = Nothing
    Set lobjstcCobert = Nothing

End Function

Public Function mfCalNewPrecif(ByRef pStcItem As stcA46V708B, _
                               ByVal pReCalcProg As Boolean, _
                               ByRef pMensagem As String, _
                               ByVal pPropTabela As String, _
                               ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfCalNewPrecif"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    mfCalNewPrecif = False

    Dim lArrayTaxas() As Variant
    Dim lQtdeTaxas As Integer
    Dim TxAlterada As Boolean
    Dim lQtdeCobert As Integer

    Dim lobjstcCobertura As stcA46V704B
    Dim lobjstcCorretor As stcA46V705B  'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)

    Dim lobjstcProlabore As stcA46V733B

    Dim lbdTab_Tarifa As ADODB.Connection  'Arquivo: P00Tarifa.
    Dim lrsTab_Tarifa As clsYasRecordSet   'Registro : TAB_Tarifa.
    Dim i As Integer
    Dim lobjPedLoc As clsA46V708A      'Objeto da classe de itens.

    '--- Taxas Intermediárias
    Dim lTaxaNet As Single
    Dim lTaxaComl As Single
    Dim lTaxaFinal As Single

    Dim lPremio As Double
    Dim lTaxaMin As Single
    Dim lTaxaMinFormula As Single
    Dim lAgravoDesconto As Single
    Dim lTaxaReser As Single
    Dim lTaxaRess As Single
    Dim lobjCobert As clsA46V704A
    Dim lretorno As String

    Dim mPer_Coml As Single
    Dim mPer_ComlAss24hrs As Single
    Dim lPer_Coml_Cota As Single
    Dim lVal_IS_Minima_ws As Single
    Dim lVal_is_basica_ws As Single

    Dim lColecaoTemp As Collection
    Dim lComissaoBasica As Double
    Dim lRegiao As Integer

    Dim lSelect As String
    Dim bdPremio As ADODB.Connection
    Dim lrst As clsYasRecordSet
    Dim lPremioAnt As Single
    Dim EhLmi As Boolean
    Dim UtilPremioMinino As Boolean
    Dim lPremioPont As Single
    On Error GoTo TrataErro

    Dim lVal_PrAnterior As Double
    Dim lVal_Pr_Orig As Double

    Dim lrsTab_Ped_Cobert As ADODB.Recordset
    Dim Vlr_Pr_Ant As Double
    Dim Vlr_Pr_Atu As Double

    Dim lCoef_LMG_LC As Double
    Dim lPer_Desc_Agrav As Double
    Dim lPossuiDePara As Boolean

    Dim lzeraDescAgravCober As Boolean
    Dim lcodcobertEndoVida As String
    Dim lNumItem As Integer
    Dim lCustoInspecao As Double
    Dim lCustoInspecaoFinal As Double
    Dim lValidaTarifa As Integer

    ''2016-10-19~2018-02-28: VerbaCom
    Dim lobjstcCoberturaTemp As stcA46V704B
    Dim valVerbaComTemp As Double
    Dim valVerbaComProp As Double
    Dim valVerbaComDif As Double
    Dim valPrTemp As Double
    valPrTemp = 0
    
    '''Call mpVerNormas(pStcItem.ObjPai, pStcItem, pMensagem)
    Set lobjCobert = New clsA46V704A
    Set lobjCobert.M46V999 = M46V999

    Set lrsTab_Ped_Cobert = New ADODB.Recordset
    lrsTab_Ped_Cobert.CursorLocation = adUseClient

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
    lTip_Atividade = pStcItem.gTip_Atividade
    pStcItem.Per_Tot_Protec = 0
    'Ficha 1000305
    lTip_Cobert = pStcItem.Tip_Cobert
    M46V999.gTipCobertura = pStcItem.Tip_Cobert

    '1100192 - Tarifa 1006 e 1007 Marco 2011
    lVal_VR_Item = pStcItem.Val_Vr

    '1001146 - Tarifa 1000 e 1001 Dezembro 2010
    '    lIndLMI = pStcItem.ObjPai.LMI_Unico
    lCodRegiao = pStcItem.gCod_RegiaoNova

    'Resseguro facultativo
    'Call gpVerNormasParametrizada(pStcItem.ObjPai, pStcItem, pMensagem, 1)

    lPer_Desc_Agrav = pStcItem.ObjPai.objstcCtrlEmis.Per_Desc_Agrav

    If pStcItem.ObjPai.Dat_Ini_Vig > pStcItem.ObjPai.Dat_Ter_Vig Then
        pMensagem = "AGE0022 - Vigência inválida. Início maior que o término."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        GoTo SAIDA
    End If
    '1200029 - Clausula 315 e Texto LMI
    If pStcItem.ObjPai.LMI_Unico = 1 And pStcItem.Cod_Ramo = e_CodRamo_Empresarial Then
        If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 315, pStcItem, pMensagem, tagDynaTrace) Then
            Call M46V999.gpGraLog(1, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If

    'CR7348 - Plano BRFoods - INICIO
    If lCod_Ramo = 112 And (lCod_Modalidade = 2218 Or lCod_Modalidade = 2219 Or (lCod_Modalidade >= 2228 And lCod_Modalidade <= 2271)) _
       And lCod_Tarifa >= 1060 Then
        If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 202, pStcItem, pMensagem, tagDynaTrace) Then
            Call M46V999.gpGraLog(1, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If
    'CR7348 - Plano BRFoods - FIM

    If M46V999.gCodProduto >= 942 Then
        If val(pStcItem.Cod_Clasf) = M46V999.gfBuscarDePara(190, 2) And val(pStcItem.Cod_Compl_Clasf) = M46V999.gfBuscarDePara(190, 2, 13) Then
            If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 312, pStcItem, pMensagem, tagDynaTrace) Then
                Call M46V999.gpGraLog(1, pMensagem)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If
        End If
        If val(pStcItem.Cod_Clasf) = M46V999.gfBuscarDePara(190, 2) And val(pStcItem.Cod_Compl_Clasf) = M46V999.gfBuscarDePara(190, 2, 14) Then
            If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 313, pStcItem, pMensagem, tagDynaTrace) Then
                Call M46V999.gpGraLog(1, pMensagem)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If
        End If
    End If

    If mfCalcPrazo(pStcItem.ObjPai.Dat_Ini_Vig, pStcItem.ObjPai.Dat_Ter_Vig, pStcItem.ObjPai.Cod_Iden_Praz, lCoePrazo, pMensagem, tagDynaTrace) = False Then
        Call M46V999.gpGraLog(0, pMensagem)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        GoTo SAIDA
    End If

    'Encontrar Corretor e Percentual de Comissão
    lPer_Coml = 0
    lCod_Corretor = 0

    pStcItem.Per_Tot_Protec = 0      'Incendio
    pStcItem.Cod_Iden_Protec_09 = 0
    pStcItem.Per_Tot_Segur = 0       'Roubo
    pStcItem.Cod_Iden_Protec_10 = 0


    If pStcItem.ObjPai.ColecaoPedCorr.Count > 0 Then
        For Each lobjstcCorretor In pStcItem.ObjPai.ColecaoPedCorr
            If lobjstcCorretor.Num_Ordem_Corr = 1 Then
                lCod_Corretor = lobjstcCorretor.Cod_Corr
            End If
            lPer_Coml = lPer_Coml + lobjstcCorretor.Per_Corr
        Next lobjstcCorretor
    End If

    'Somando os percentuais de prolabore
    If pStcItem.ObjPai.ColecaoProlabore.Count > 0 Then
        For Each lobjstcProlabore In pStcItem.ObjPai.ColecaoProlabore
            lPer_Coml = lPer_Coml + lobjstcProlabore.Perc_Canal
        Next lobjstcProlabore
    End If

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTab_Tarifa, pMensagem) = False Then
        Call M46V999.gpGraLog(0, pMensagem)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
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
                    'lVal_is_basica_ws = lobjstcCobertura.Val_IS
                    lVal_is_basica_ws = pStcItem.gISBasica
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
                    ElseIf lCod_Ramo = 114 Then    '1200004 - Tarifa 1028 e 1029 Janeiro 2012
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
    If (M46V999.gCodProduto >= 1018) And (val(lCod_Clasf) = M46V999.gfBuscarDePara(282, 2)) And (val(lCod_Compl_Clasf) = M46V999.gfBuscarDePara(282, 2, 10)) And (pStcItem.Cod_Ramo = e_CodRamo_Empresarial) Then
        'CLAUSULA PARTICULAR - DANOS A BENS DE HOSPEDES ATÉ R$ 20 MIL
        If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 320, pStcItem, pMensagem, tagDynaTrace) Then
            Call M46V999.gpGraLog(1, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If
        
    'Item nova tarifa 1090 / 1091 não permitir descontos para vida
    lzeraDescAgravCober = False
    lcodcobertEndoVida = ""
    lNumItem = 1 ' O objeto de apólice carrega apenas o item atual.
    For i = 1 To pStcItem.ColecaoPedCobert.Count
        If pStcItem.ObjPai.Cod_Dizer = 12 Then
            'zerar o desconto quando inclusão de cobertura de vida
            If pStcItem.Tip_Emissao <> e_TipEmissao_SeguroNovo Then
                If pStcItem.ObjPai.gStcPedidoApolice.ColecaoPedLoc.Count > 0 Then
                    With pStcItem.ObjPai.gStcPedidoApolice.ColecaoPedLoc(lNumItem)
                            If Not .FindCobertura(pStcItem.ColecaoPedCobert(i).Cod_Cobert) And pStcItem.ColecaoPedCobert(i).Cod_Grupo_Cobertura = 4 Then
                                lzeraDescAgravCober = True
                                Exit For
                            ElseIf pStcItem.ColecaoPedCobert(i).Cod_Grupo_Cobertura = 4 Then
                                lzeraDescAgravCober = True
                                Exit For
                            End If
                    End With
                Else
                    pMensagem = "AGE0060 - Processo de calculo abortado. Favor verificar se a apólice esta cancelada."
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If
            End If
        Else
            If pStcItem.ObjPai.Cod_Dizer = 1 Or pStcItem.ObjPai.Cod_Dizer = 0 Then
                lzeraDescAgravCober = True
                Exit For
            End If
        End If
        
        'Nova tarifa 1092 / 1093
'        If pStcItem.ColecaoPedCobert(i).Cod_Cobert = 1215 Then
'            lzeraDescAgravCober = True
'            Exit For
'        End If
        
        If lzeraDescAgravCober Then
            Exit For
        End If
    Next i



'    *******************************************************************************************************
'    *#20180829'    *Finalidade: Apura Desconto Comercial Congenere
'    * Controla o Acesso a Busca do Desconto Congenere
'    *******************************************************************************************************
     Dim ItemAtivo As Integer

    For Each lobjstcCobertura In pStcItem.ColecaoPedCobert
        lQtdeTaxas = lQtdeTaxas + 1
        ReDim Preserve lArrayTaxas(9, lQtdeTaxas)

        lCod_Cobert = lobjstcCobertura.Cod_Cobert
        lVal_IS = lobjstcCobertura.Val_IS

        'If lobjstcCobertura.gCod_Grupo_Cobert > "1" Then '2 = Acessórias
        '    lQtdeCobert = lQtdeCobert + 1                '3 = Adicionais
        'End If
        If lobjstcCobertura.Cod_Cobert = M46V999.gfBuscarDePara(22, 1) And M46V999.gCodProduto >= 948 And lobjstcCobertura.Cod_Ramo = e_CodRamo_Residencial Then
            If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 304, pStcItem, pMensagem, tagDynaTrace) Then
                Call M46V999.gpGraLog(1, pMensagem)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If
        End If


        '1100589 - Tarifa 1014 e 1015 Julho 2011
        If M46V999.gCodProduto >= 1014 Then
            If lobjstcCobertura.Cod_Ramo = e_CodRamo_Empresarial Then
                Select Case lobjstcCobertura.Cod_Cobert
                Case 56, 122, 123
                    'CLAUSULA PARTICULAR - COBERT. PATIOS - SUBTRACAO DE VEICULOS
                    If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 318, pStcItem, pMensagem, tagDynaTrace) Then
                        Call M46V999.gpGraLog(1, pMensagem)
                        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                        GoTo SAIDA
                    End If
                Case 57, 124, 125
                    If M46V999.gCodProduto <= 1015 Then    '1100705 - Tarifa 1016 e 1017 Agosto 2011
                        'CLAUSULA PARTICULAR - COBERT.RC.CONCE. SUBTRACAO DE VEICULOS
                        If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 319, pStcItem, pMensagem, tagDynaTrace) Then
                            Call M46V999.gpGraLog(1, pMensagem)
                            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                            GoTo SAIDA
                        End If
                    End If
                Case 23
                    '1100864 - Tarifa 1018 e 1019 Setembro 2011
                    If M46V999.gCodProduto < 1080 Then
                        If (M46V999.gCodProduto >= 1018) And val(lCod_Clasf) = M46V999.gfBuscarDePara(282, 2) And val(lCod_Compl_Clasf) = M46V999.gfBuscarDePara(282, 2, 10) Then
                        'CLAUSULA PARTICULAR - ROUBO/FURTO C/ VEST. BENS DE HOSPEDES
                        If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 321, pStcItem, pMensagem, tagDynaTrace) Then
                            Call M46V999.gpGraLog(1, pMensagem)
                            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                            GoTo SAIDA
                        End If
                    End If
                    End If
                Case 24
                    '1100864 - Tarifa 1018 e 1019 Setembro 2011
                    If (M46V999.gCodProduto >= 1018) And val(lCod_Clasf) = M46V999.gfBuscarDePara(282, 2) And val(lCod_Compl_Clasf) = M46V999.gfBuscarDePara(282, 2, 10) Then
                        'CLAUSULA PARTICULAR - ROUBO DE VALORES INTERIOR DO ESTAB.
                        If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 322, pStcItem, pMensagem, tagDynaTrace) Then
                            Call M46V999.gpGraLog(1, pMensagem)
                            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                            GoTo SAIDA
                        End If
                    End If
                End Select
            End If
        End If

        'Calcula qtde de cobertura para o desconto progressivo -> Per_Desc_Progres
        If lCod_Tarifa >= 1028 Then    '1200004 - Tarifa 1028 e 1029 Janeiro 2012
            If pReCalcProg Then
                If lCod_Tarifa >= 1080 Then
                        If lobjstcCobertura.gCod_Grupo_Cobert <> 4 Then
                            lQtdeCobert = lQtdeCobert + 1
                        End If
                    Else
                        If lCod_Ramo = 112 And _
                           lobjstcCobertura.gCod_Grupo_Cobert <> 1 And _
                           lobjstcCobertura.gCod_Grupo_Cobert <> 5 And _
                           lVal_IS >= lVal_IS_Minima_ws Then
                            lQtdeCobert = lQtdeCobert + 1
                        ElseIf lCod_Ramo = 113 And _
                               lobjstcCobertura.gCod_Grupo_Cobert <> 1 And _
                               lobjstcCobertura.gCod_Grupo_Cobert <> 5 And _
                               lVal_IS >= lVal_IS_Minima_ws Then
                            lQtdeCobert = lQtdeCobert + 1
                        ElseIf lCod_Ramo = 114 And _
                               lobjstcCobertura.gCod_Grupo_Cobert <> 1 And _
                               lobjstcCobertura.gCod_Grupo_Cobert <> 5 And _
                               lVal_IS >= lVal_IS_Minima_ws Then
                            lQtdeCobert = lQtdeCobert + 1
                        End If
                    End If
            End If
        Else
            If pReCalcProg Then
                If lCod_Ramo = 112 And _
                   lobjstcCobertura.gCod_Grupo_Cobert <> 1 And _
                   lobjstcCobertura.gCod_Grupo_Cobert <> 5 And _
                   lCod_Cobert <> M46V999.gfBuscarDePara(23, 1) And _
                   lCod_Cobert <> M46V999.gfBuscarDePara(24, 1) And _
                   lCod_Cobert <> M46V999.gfBuscarDePara(25, 1) And _
                   lCod_Cobert <> M46V999.gfBuscarDePara(26, 1) And _
                   lCod_Cobert <> M46V999.gfBuscarDePara(63, 1) And _
                   lVal_IS >= lVal_IS_Minima_ws Then
                    lQtdeCobert = lQtdeCobert + 1
                ElseIf lCod_Ramo = 113 And lCod_Tarifa < 946 And _
                       lobjstcCobertura.gCod_Grupo_Cobert <> 1 And _
                       lobjstcCobertura.gCod_Grupo_Cobert <> 5 And _
                       lCod_Cobert <> M46V999.gfBuscarDePara(16, 1) And _
                       lCod_Cobert <> M46V999.gfBuscarDePara(5, 1) And _
                       lVal_IS >= lVal_IS_Minima_ws Then
                    lQtdeCobert = lQtdeCobert + 1
                ElseIf lCod_Ramo = 113 And lCod_Tarifa >= 946 And _
                       lobjstcCobertura.gCod_Grupo_Cobert <> 1 And _
                       lobjstcCobertura.gCod_Grupo_Cobert <> 5 And _
                       lCod_Cobert <> M46V999.gfBuscarDePara(16, 1) And _
                       lVal_IS >= lVal_IS_Minima_ws Then
                    lQtdeCobert = lQtdeCobert + 1
                ElseIf lCod_Ramo = 114 And _
                       lobjstcCobertura.gCod_Grupo_Cobert <> 1 And _
                       lobjstcCobertura.gCod_Grupo_Cobert <> 5 Then
                    lQtdeCobert = lQtdeCobert + 1
                End If
            End If
        End If

        lClasseLMG = 0
        lCodFaixa_Reg = 0
        lCodClasse_Reg = 0
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
                lRegiao = 114    'ok
                'Ficha 901673 e 901772
            Case 2111 To 2130
                lRegiao = 114    'ok
            Case 2042 To 2049
                lRegiao = 310    'ok
                '                Case Else 'Comentado porque foi criada funcionalidade de cadastro/manutenção de plano no gerenciador que gera todo cadastro automaticamente
                '                    lCodFaixa_Reg = 1
                '                    GoTo Achei_Faixa
            End Select
        End If
        If lRegiao > 0 Then
            lSelect = "Select Cod_Faixa_Reg, Cod_Classe_Reg from Tbm_Faixa_Regiao WITH (NOLOCK) "
            lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
            lSelect = lSelect & "    and Cod_Modalidade = " & lCod_Modalidade
            lSelect = lSelect & "    and Cod_Reg = " & lRegiao
            lSelect = lSelect & "    and Cod_Tarif = " & lCod_Tarifa
            lSelect = lSelect & "    and Cod_Sequencia  <= " & lCod_Sequencia
            lSelect = lSelect & "    order by Cod_Sequencia desc "

        Else
            lSelect = "Select Cod_Faixa_Reg, Cod_Classe_Reg from Tbm_Faixa_Regiao WITH (NOLOCK) "
            lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
            lSelect = lSelect & "    and Cod_Modalidade = " & lCod_Modalidade
            lSelect = lSelect & "    and Cep_Ini <= " & lNum_Cep
            lSelect = lSelect & "    and Cep_Fim >= " & lNum_Cep
            lSelect = lSelect & "    and Cod_Tarif = " & lCod_Tarifa
            lSelect = lSelect & "    and Cod_Sequencia  <= " & lCod_Sequencia
            lSelect = lSelect & "    order by Cod_Sequencia desc "
        End If
        If M46V999.gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, pMensagem) = False Then
            Call M46V999.gpGraLog(0, pMensagem)
            Call M46V999.gpFecha2(lbdTab_Tarifa)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If

        If Not lrsTab_Tarifa.EOF Then
            lCodFaixa_Reg = M46V999.Check(lrsTab_Tarifa("Cod_Faixa_Reg"))
            lCodClasse_Reg = M46V999.Check(lrsTab_Tarifa("COD_Classe_Reg"))
        End If

        'ficha 1200806 - Item 13 Rubrica de Calcados
        'Verifica nova classe de pra Fabrica de Calçados
        'If M46V999.gCodProduto >= 1044 Then
        '   If Val(pStcItem.Cod_Clasf) = 104 And Val(pStcItem.Cod_Compl_Clasf) = 20 Then
        '       If lCodClasse_Reg = 5 Then
        '           lCodFaixa_Reg = 2
        '       ElseIf lCodClasse_Reg = 6 Then
        '           lCodFaixa_Reg = 3
        '       Else
        '           lCodFaixa_Reg = 1
        '       End If
        '   End If
        'End If


Achei_Faixa:
        '2. Encontrar a Classe
        lSelect = "Select Cod_Classe from Tbm_Classe WITH (NOLOCK) "
        lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
        lSelect = lSelect & "    and Cod_Modalidade = " & lCod_Modalidade
        lSelect = lSelect & "    and Cod_Clasf = '" & lCod_Clasf & "'"
        lSelect = lSelect & "    and Cod_Compl_Clasf = '" & lCod_Compl_Clasf & "'"
        lSelect = lSelect & "    and Cod_Cobert = " & lCod_Cobert
        lSelect = lSelect & "    and Cod_Tarif = " & lCod_Tarifa
        lSelect = lSelect & "    and Cod_Sequencia  <= " & lCod_Sequencia
        lSelect = lSelect & "    order by Cod_Sequencia desc "

        If M46V999.gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, pMensagem) = False Then
            Call M46V999.gpFecha2(lbdTab_Tarifa)
            Call M46V999.gpGraLog(0, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If

        If lrsTab_Tarifa.EOF Then

            GoTo N_loc_Taxa
        End If
        lCod_Classe = M46V999.Check(lrsTab_Tarifa("Cod_Classe"))



        'Verifico se tem ajustes pontuais para encontrar a classe
        lClasseAPont = 0
        If mfAjustesPontuais(e_Tip_Valor_Classe, tpoNumerico, lClasseAPont, pMensagem, pStcItem, 0, tagDynaTrace) = False Then
            Call M46V999.gpGraLog(0, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        If lClasseAPont <> 0 Then
            lCod_Classe = lClasseAPont
        End If

        ' 1300197 - Tarifa 1056 e 1057 - Inicio
        lClasseAPont = 0
        If mfAjustesPontuais(e_Tip_Restricao_Corretor, tpoNumerico, lClasseAPont, pMensagem, pStcItem, 0, tagDynaTrace) = False Then
            Call M46V999.gpGraLog(0, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        If lClasseAPont <> 0 Then
            lCod_Classe = lClasseAPont
        End If
        ' 1300197 - Tarifa 1056 e 1057 - Fim

        '3. Encontrar a Classe para os Vultosos Específicos
        'Desativado tarifas 1044 e 1045
        If M46V999.gCodProduto >= 1044 Then
            lClasseLMG = 1
        Else
            lSelect = "Select Cod_Classe_LMG from Tbm_Classe_Vultosos WITH (NOLOCK)  "
            lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
            lSelect = lSelect & "    and Cod_Modalidade = " & lCod_Modalidade
            lSelect = lSelect & "    and Cod_Tip_Reg = 'E'"
            lSelect = lSelect & "    and Cod_Clasf = '" & lCod_Clasf & "'"
            lSelect = lSelect & "    and Cod_Compl_Clasf = '" & lCod_Compl_Clasf & "'"
            lSelect = lSelect & "    and Cod_Construcao = " & lCod_Construcao
            lSelect = lSelect & "    and Val_LMG_Min <= " & M46V999.gfForValor(1, lVal_IS)
            lSelect = lSelect & "    and Val_LMG_Max >= " & M46V999.gfForValor(1, lVal_IS)
            lSelect = lSelect & "    and Cod_Cobert = " & lCod_Cobert
            lSelect = lSelect & "    and Cod_Tarif = " & lCod_Tarifa
            lSelect = lSelect & "    and Cod_Sequencia  <= " & lCod_Sequencia
            lSelect = lSelect & "    order by Cod_Sequencia desc "

            If M46V999.gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, pMensagem) = False Then
                Call M46V999.gpFecha2(lbdTab_Tarifa)
                Call M46V999.gpGraLog(0, pMensagem)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If

            If Not lrsTab_Tarifa.EOF Then
                lClasseLMG = M46V999.Check(lrsTab_Tarifa("Cod_Classe_LMG"))
            Else
                '3. Encontrar a Classe para os Vultosos Default
                lSelect = "Select Cod_Classe_LMG from Tbm_Classe_Vultosos WITH (NOLOCK)  "
                lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
                lSelect = lSelect & "    and Cod_Modalidade = " & lCod_Modalidade
                lSelect = lSelect & "    and Cod_Tip_Reg = 'D'"
                lSelect = lSelect & "    and Val_LMG_Min <= " & M46V999.gfForValor(1, lVal_IS)
                lSelect = lSelect & "    and Val_LMG_Max >= " & M46V999.gfForValor(1, lVal_IS)
                lSelect = lSelect & "    and Cod_Cobert = " & lCod_Cobert
                lSelect = lSelect & "    and Cod_Tarif = " & lCod_Tarifa
                lSelect = lSelect & "    and Cod_Sequencia  <= " & lCod_Sequencia
                lSelect = lSelect & "    order by Cod_Sequencia desc "
                If M46V999.gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, pMensagem) = False Then
                    Call M46V999.gpFecha2(lbdTab_Tarifa)
                    Call M46V999.gpGraLog(0, pMensagem)
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If
                If Not lrsTab_Tarifa.EOF Then
                    lClasseLMG = M46V999.Check(lrsTab_Tarifa("Cod_Classe_LMG"))
                End If
            End If
        End If


        '4. Encontrar a Taxa
        lSelect = "Select Per_Taxa_Min,  Per_Taxa_Reser, Per_Taxa_Ress from Tbm_Taxa_Classe WITH (NOLOCK) "
        lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
        lSelect = lSelect & "    and Cod_Modalidade = " & lCod_Modalidade
        lSelect = lSelect & "    and Cod_Cobert = " & lCod_Cobert
        
        'Se Danos Eletricos com elevador alterar a taxa
        If lCod_Ramo = e_CodRamo_Condominio And lCod_Cobert = 1103 Then
            If pStcItem.Possui_Elevador = 1 Then
                lCod_Construcao = 1
                
                'Atividades 809 se for com elevador seta classe 1
                If Trim(pStcItem.Cod_Clasf) = "809" And pStcItem.Cod_Iden_Tarif < 1102 Then
                    lCod_Classe = 1
                End If
            Else
                lCod_Construcao = 0
                
                'Atividades 801 se for sem elevador setar classe 2
                If Trim(pStcItem.Cod_Clasf) = "801" And pStcItem.Cod_Iden_Tarif < 1102 Then
                    lCod_Classe = 2
                End If
            End If
                        
        End If
        
        lSelect = lSelect & "    and Cod_Classe = " & lCod_Classe
        lSelect = lSelect & "    and Cod_Faixa_Reg in (0," & lCodFaixa_Reg & ")"
        lSelect = lSelect & "    and Cod_Classe_LMG in (0," & lClasseLMG & ")"
        lSelect = lSelect & "    and Cod_Construcao in (0," & lCod_Construcao & ")"
        
        lTip_Cobert = pStcItem.Tip_Cobert
        If lTip_Cobert = 0 And pStcItem.Ind_VerbaSep = 1 Then  'Se predio+conteudo e verbas separadas = sim
            If lobjstcCobertura.gCod_Grupo_Cobert = 1 And (lCod_Ramo <> 112 Or lCod_Tarifa < 1122) Then 'Cobertura Predio
                lTip_Cobert = 1
            End If
            If lobjstcCobertura.gCod_Grupo_Cobert = 5 Then   'Cobertura Conteudo
                lTip_Cobert = 2
            End If
        Else
            If lobjstcCobertura.gCod_Grupo_Cobert = 5 And pStcItem.Tip_Emissao = e_TipEmissao_Endosso Then
                lTip_Cobert = 2
            End If
        End If
        
        'Ficha 1000305
        lSelect = lSelect & "    and Tip_Cobert in (0," & lTip_Cobert & ")"
        lSelect = lSelect & "    and Cod_Tarif = " & lCod_Tarifa
        lSelect = lSelect & "    and Cod_Sequencia  <= " & lCod_Sequencia
        lSelect = lSelect & "    and Tip_Atividade in (0, " & lTip_Atividade & ")"
        lSelect = lSelect & "    order by Cod_Sequencia desc,Cod_Faixa_Reg desc , Cod_Classe_LMG desc, Cod_Construcao desc, Tip_Cobert desc, Tip_Atividade desc"

        If M46V999.gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, pMensagem) = False Then
            Call M46V999.gpFecha2(lbdTab_Tarifa)
            Call M46V999.gpGraLog(0, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        
        'Se o Imobiliario com premio informado não tiver taxa força 1
        If lrsTab_Tarifa.EOF And (lCod_Modalidade <> 20 Or Not lobjstcCobertura.PremioInformado) Then
            GoTo Ver_Ajuste
        Else
            If Not lrsTab_Tarifa.EOF Then
                If lrsTab_Tarifa.RecordCount = 1 Then
                    lTaxaMin = M46V999.Check(lrsTab_Tarifa("Per_Taxa_Min"))
                    lTaxaReser = M46V999.Check(lrsTab_Tarifa("Per_Taxa_Reser"))
                    lTaxaRess = M46V999.Check(lrsTab_Tarifa("Per_Taxa_Ress"))
                Else
                    'Fazer looping  aqui
                    'MsgBox "O Sistema encontrou mais de um registro de taxa! Não será possivel o cálculo"
                    'Verificar com Produto
                    lTaxaMin = M46V999.Check(lrsTab_Tarifa("Per_Taxa_Min"))
                    lTaxaReser = M46V999.Check(lrsTab_Tarifa("Per_Taxa_Reser"))
                    lTaxaRess = M46V999.Check(lrsTab_Tarifa("Per_Taxa_Ress"))
                End If
            Else
                'Se o Imobiliario com premio informado não tiver taxa força 1
                lTaxaMin = 1
                lTaxaReser = 1
                lTaxaRess = 1
            End If
        End If


Ver_Ajuste:
        'Verifico se tem ajustes pontuais para encontrar a taxa
        lTaxaAPont = 0
        If mfAjustesPontuais(e_Tip_Valor_TaxaClasse, tpoNumerico, lTaxaAPont, pMensagem, pStcItem, 0, tagDynaTrace) = False Then
            Call M46V999.gpGraLog(0, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
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
        If mfAjustesPontuais(e_Tip_Valor_CoeficienteTaxa, tpoNumerico, lTaxaAPont, pMensagem, pStcItem, 0, tagDynaTrace) = False Then
            Call M46V999.gpGraLog(0, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        '1100864 - Tarifa 1018 e 1019 Setembro 2011 (Item 3 Não aplicar o coeficiente p/ Renovações Yasuda Sem Sinistro)

        'Nova tarifa 1092 / 1093 inclusão de cobertura bicicleta na restrição
        If lzeraDescAgravCober And lobjstcCobertura.Cod_Grupo_Cobertura = 4 And lCod_Tarifa >= 1090 Then
            lTaxaAPont = 0
        End If

        ' Se for Imobiliário e Corretores especiais, zerar o desconto ou agravo
        'Tranferencia da regra de 917581 - Triunfo para 939571 - BR Insurance - 1174
        If lCod_Modalidade = 20 And _
              (lCod_Corretor = 18901 _
            Or lCod_Corretor = 366199 _
            Or lCod_Corretor = 415632 _
            Or lCod_Corretor = 682987 _
            Or lCod_Corretor = 939571) _
            Then lTaxaAPont = 0

        lTaxaMinFormula = lTaxaMin
        lAgravoDesconto = 1
        'If pStcItem.Num_Item = 2 Then
        'pStcItem.Ind_Sinistro = 0
        'End if
        If lTaxaAPont <> 0 Then
            If lCod_Tarifa < 1080 Then
                If (lCod_Modalidade > 1000) Or _
                        (lCod_Tarifa < 1018) Or _
                        (lCod_Tarifa >= 1018 And pStcItem.Tip_Emissao <> e_TipEmissao_RenYasuda Or pStcItem.Ind_Sinistro <> 2) Then
                                lAgravoDesconto = lTaxaAPont
                                lTaxaMin = lTaxaMin * lTaxaAPont
                                lTaxaReser = lTaxaMin * lTaxaAPont
                                lTaxaRess = lTaxaMin * lTaxaAPont
                End If
            Else
                    lAgravoDesconto = lTaxaAPont
                    lTaxaMin = lTaxaMin * lTaxaAPont
                    lTaxaReser = lTaxaMin * lTaxaAPont
                    lTaxaRess = lTaxaMin * lTaxaAPont
            End If
        End If


        '1000970 - Item 1 - Empresarial/Residencial/Condominio - para corretor 0663299 Legend agravar em 30% todas as taxas de todas as coberturas dos produtos.
        lTaxaAPont = 0
        If mfAjustesPontuais(e_Tip_Valor_CoeficienteEspecialTaxa, tpoNumerico, lTaxaAPont, pMensagem, pStcItem, 0, tagDynaTrace) = False Then
            Call M46V999.gpGraLog(0, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        'Tranferencia da regra de 917581 - Triunfo para 939571 - BR Insurance
        ' Se for Imobiliário e Corretores especiais, zerar o desconto ou agravo
        If lCod_Modalidade = 20 And _
              (lCod_Corretor = 18901 _
            Or lCod_Corretor = 366199 _
            Or lCod_Corretor = 415632 _
            Or lCod_Corretor = 682987 _
            Or lCod_Corretor = 939571) _
            Then lTaxaAPont = 0
            
        'nova tarifa 1092 / 1093
'        If lCod_Cobert = 1215 And lCod_Ramo = 113 Then
'            lTaxaAPont = 0
'        End If

        If lTaxaAPont <> 0 Then
            lAgravoDesconto = lAgravoDesconto * lTaxaAPont
            lTaxaMin = lTaxaMin * lTaxaAPont
            lTaxaReser = lTaxaMin * lTaxaAPont
            lTaxaRess = lTaxaMin * lTaxaAPont
        End If
        
        ' Coeficiente de Taxa por divisao (Agravo e desconto)
        lTaxaAPont = 0
        If mfAjustesPontuais(e_Tip_Valor_CoeficienteDivisaoTaxa, tpoNumerico, lTaxaAPont, pMensagem, pStcItem, 0, tagDynaTrace) = False Then
            Call M46V999.gpGraLog(0, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If

        'Tranferencia da regra de 917581 - Triunfo para 939571 - BR Insurance
        ' Se for Imobiliário e Corretores especiais, zerar o desconto ou agravo
        If lCod_Modalidade = 20 And _
              (lCod_Corretor = 18901 _
            Or lCod_Corretor = 366199 _
            Or lCod_Corretor = 415632 _
            Or lCod_Corretor = 682987 _
            Or lCod_Corretor = 939571) _
            Then lTaxaAPont = 0
        
        If lTaxaAPont <> 0 Then
            lAgravoDesconto = lAgravoDesconto / lTaxaAPont
            lTaxaMin = lTaxaMin / lTaxaAPont
            lTaxaReser = lTaxaMin / lTaxaAPont
            lTaxaRess = lTaxaMin / lTaxaAPont
        End If


        lArrayTaxas(0, lQtdeTaxas) = lCod_Cobert
        lArrayTaxas(1, lQtdeTaxas) = lClasseLMG
        lArrayTaxas(2, lQtdeTaxas) = lCodFaixa_Reg
        lArrayTaxas(3, lQtdeTaxas) = lCod_Classe
        lArrayTaxas(4, lQtdeTaxas) = lTaxaMin
        lArrayTaxas(5, lQtdeTaxas) = lTaxaReser
        lArrayTaxas(6, lQtdeTaxas) = lTaxaRess
        lArrayTaxas(7, lQtdeTaxas) = lTaxaMinFormula
        lArrayTaxas(8, lQtdeTaxas) = lAgravoDesconto
        
'    *******************************************************************************************************
'    *#20180829
'    *Finalidade: Apura Desconto Comercial Congenere
'    * Cod_Tip_Chave definido com o valor 76
'    *******************************************************************************************************
       Dim MyQuery As String
       Dim prst As clsYasRecordSet

       Dim Cong As stcA46V708B
       Set Cong = pStcItem.ObjPai.ColecaoPedLoc.Item(lobjstcCobertura.Num_Item)

        '[Chave_Ini] = Representa o Tipo de Emissão
        '[Chave_Ter] = Cod_Seguradora_Congenere

        If ItemAtivo <> lobjstcCobertura.Num_Item Then
        lCoef_Desc_Coml_Congenere = 0
        ItemAtivo = lobjstcCobertura.Num_Item
            If Cong.DescComlCongenereIncorporado = 0 Then
                MyQuery = "SELECT Cod_Tip_Chave,Cod_Ramo,Chave_Ini,Chave_Ter,Valor FROM [dbo].[Tbm_Parameter] WHERE [Cod_Tip_Chave]= 76 AND [Chave_Ini] = " & Cong.Tip_Emissao & " AND Chave_Ter = " & Cong.Ren_Con_Seguradora
                Set prst = New clsYasRecordSet
                M46V999.expObterDados "RamosDiversos", MyQuery, prst, pMensagem
                    If Len(pMensagem) > 0 Then
                        Call M46V999.gpGraLog(0, pMensagem)
                        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                        GoTo TrataErro
                    Else
                        If Not prst.EOF Then
                             lCoef_Desc_Coml_Congenere = CDbl(Trim(prst("Valor")))
                        End If
                    End If
                pStcItem.ObjPai.ColecaoPedLoc.Item(lobjstcCobertura.Num_Item).DescComlCongenereIncorporado = 1
                MyQuery = "UPDATE dbo.Tab_Ped_Loc_Compl SET DescComlCongenereIncorporado = 1 WHERE Nosso_numero = '" & Cong.nosso_numero & "' AND NUM_ITEM = " & lobjstcCobertura.Num_Item
                M46V999.expExeSQL IIf(UCase(App.EXEName) Like "D46O003*", "SyasPortal", "RamosDiversos"), MyQuery, pMensagem
                If pMensagem <> "" Then
                    Call M46V999.gpGraLog(2, pMensagem)
                    GoTo TrataErro
                End If
            End If
        End If

        
    
'    *******************************************************************************************************

    Next lobjstcCobertura


    '[FRH] F:900097
    '------------------------------------------------------
    Dim val86 As Long
    Dim val64 As Long
    Dim val66 As Long
    Dim y As Long
    If M46V999.gCodProduto >= 956 Then
        val86 = 9999
        val64 = 9999
        val66 = 9999

        For y = 0 To UBound(lArrayTaxas, 2)
            Select Case val(lArrayTaxas(0, y))
            Case 86: val86 = y
            Case 64: val64 = y
            Case 66: val66 = y
            End Select
        Next y

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
    If lCod_Tarifa >= 1080 And lCod_Modalidade < 1000 Then
        If mfAjustesPontuais(e_Tip_Valor_DescontoAgrupamento, tpoNumerico, lCoef_Desc_Prog, pMensagem, pStcItem, lQtdeCobert, tagDynaTrace) = False Then
            Call M46V999.gpGraLog(0, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    Else
        If (pStcItem.ObjPai.getPlano.EhPlaninho = True) And _
           (pStcItem.ObjPai.getPlano.TemValorRegra(eRegraPlano_DescontoProgressivo) = True) Then
            lCoef_Desc_Prog = pStcItem.ObjPai.getPlano.ValorRegra(eRegraPlano_DescontoProgressivo)
        Else
            If pReCalcProg And pStcItem.Cod_DizerItem <= 1 Then
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
        If (lCod_Modalidade = 2026) And (M46V999.gCodProduto < 992) Then
            lCoef_Desc_Prog = 7
        End If

        'Ficha 901346 - o Produto calculou o plano 06 com 7% de Progressivo e não 10%
        'Ficha 901673 e 901772
        If (lCod_Modalidade >= 2061 And lCod_Modalidade <= 2100) Or (lCod_Modalidade >= 2111 And lCod_Modalidade <= 2130) Then
            lCoef_Desc_Prog = 10
        End If

    End If

    'Coeficiente IS/VR Lucros Cessantes
    lCoef_LMG_LC = 1
    
    If M46V999.gCod_Ramo = e_CodRamo_Empresarial And M46V999.gCodProduto >= 1076 And pStcItem.Cod_DizerItem = 0 Then
        If Not BuscaCoef_IS_VR_LC(pStcItem, lCoef_LMG_LC, pMensagem) Then
            Call M46V999.gpGraLog(0, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If

    '-------------------------------------------------------------
    'Nesse momento vai realmente fazer os cálculos
    '-------------------------------------------------------------
    
    'Carrega o De->Para caso não tinha sido carregado
    lPossuiDePara = False
    If M46V999.gRst_Tbm_De_Para.State = 0 Then
        lPossuiDePara = True
    ElseIf M46V999.gRst_Tbm_De_Para.RecordCount = 0 Then
        lPossuiDePara = True
    End If
    
    If pStcItem.Cod_Iden_Tarif >= 1080 And lPossuiDePara = True And pStcItem.ObjPai.Cod_Estip < 2000 Then
        M46V999.gCodProduto = pStcItem.Cod_Iden_Tarif
        M46V999.gCod_Ramo = pStcItem.Cod_Ramo
        M46V999.M46V111.gCod_plano = pStcItem.ObjPai.Cod_Estip
        
        Call M46V999.mfCarregarDePara
    End If
    
    Dim lLogBatente As stcA46V745B
    Set lLogBatente = New stcA46V745B
    
    If pStcItem.ObjPai.CalBatente Then
        If Not ConsisteRegraBatente(pStcItem.ObjPai, lPer_Coml, False, lrsTab_Ped_Cobert, Vlr_Pr_Ant, Vlr_Pr_Atu, pStcItem) Then
            'CASO DE APOLICES COM MAIS DE UM ITEM SE UM DELE NÃO PASSAR NAS REGRAS DE CONSISTENCIA
            pStcItem.ObjPai.CalBatente = False
        End If
    End If
    
    '11.72 - Batente 0 Corretora Triunfo
    'Tranferencia da regra de 917581 - Triunfo para 939571 - BR Insurance
    If (pStcItem.ObjPai.Num_Apol_An > 0 And (pStcItem.ObjPai.Cod_Ramo = e_CodRamo_Residencial Or pStcItem.ObjPai.Cod_Ramo = e_CodRamo_Condominio) And pStcItem.ObjPai.Cod_Dizer = 0) _
    And (pStcItem.Tip_Emissao = e_TipEmissao_RenYasuda And (M46V999.gCod_Corretor = 939571 Or M46V999.gCod_Corretor = 930446) And pStcItem.ObjPai.Cod_Estip <> 20) Then
        If Not ConsisteRegraBatente(pStcItem.ObjPai, lPer_Coml, False, lrsTab_Ped_Cobert, Vlr_Pr_Ant, Vlr_Pr_Atu, pStcItem) Then
            'CASO DE APOLICES COM MAIS DE UM ITEM SE UM DELE NÃO PASSAR NAS REGRAS DE CONSISTENCIA
            pStcItem.ObjPai.CalBatente = False
        End If
    End If
    
    'limpar a coleção
    Set pStcItem.ColecaoRegraBatente = New Collection

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

        ' Se existe Prêmio informado, ele está na propriedade VAL_PR. É preciso guardar este valor
        Dim ValPremioInformado As Double
        If lobjstcCobertura.PremioInformado = True Then
            ValPremioInformado = lobjstcCobertura.VAL_PR
        End If

        'Ficha 901673 e 901772
        '        If (lCod_Modalidade >= 2001 And lCod_Modalidade <= 2031) Or (lCod_Modalidade >= 2051 And lCod_Modalidade <= 2100) Or (lCod_Modalidade >= 2111 And lCod_Modalidade <= 2130) Then
        'If M46V999.gEplano = True Then
        'lCoef_Desc_Coml = M46V111.gDesc_Coml

        If pStcItem.ObjPai.getPlano.ehPlanoEngessado = True Then
            lCoef_Desc_Coml = pStcItem.getPlanoItem.ValorRegra(eRegraPlano_DescontoComercial)
        Else
'            If pStcItem.ObjPai.Cod_Produto >= 1080 And pStcItem.Cod_Ramo = 114 Then
'                lCoef_Desc_Coml = 0
'            Else
                lCoef_Desc_Coml = lobjstcCobertura.Per_Desc_Coml
'            End If
        End If

        Dim lComissao As Double    'Comissão.

        'TARIFA 11.66/11.67 - 04/2018
        'Item 12 Inclusão do Desconto na Web para Corretora 915761 - PORTOALEGRENSE
        'Flexibilização do sistema para aplicação de desconto - Residencial
        If pStcItem.ObjPai.ColecaoPedCorr(1).Cod_Corr = 915761 And pStcItem.Cod_Ramo = 113 And pStcItem.Cod_Plano = 30 Then
            
            lCoef_Desc_Coml = 0
            
            lComissao = (pStcItem.ObjPai.objstcCtrlEmis.Per_Corr + pStcItem.ObjPai.objstcCtrlEmis.Per_Corr_2 + pStcItem.ObjPai.objstcCtrlEmis.Per_Corr_3)
                         
            'Para contratações com até 30% de comissão, o sistema aplicará 10% de desconto no prêmio.
            If lComissao > 0 And lComissao <= 25 Then
               lCoef_Desc_Coml = 15
            End If
            
            'Para contratações com até 25% de comissão, o sistema aplicará 15% de desconto no prêmio.
            If lComissao > 25 And lComissao <= 30 Then
               lCoef_Desc_Coml = 10
            End If

        End If

        'CR3939 - Pegar valor do desconto comercial
        If Not ((Not UCase(App.EXEName) Like "D46O003*") And lCod_Corretor = 632007 And pStcItem.Tab_Clasf < 1000) Then ' 11.12 Incluido para liberar Desc_Comercial para Corretor Aratu no Kit
                If pStcItem.ObjPai.Cod_Produto < 1080 Or (pStcItem.ObjPai.Cod_Produto >= 1100 And pStcItem.Tab_Clasf <= 2000) Then
                    If mfAjustesPontuais(e_Tip_Valor_DescComl_TipEmis, TpoValor, lCoef_Desc_Coml, pMensagem, pStcItem, 0, tagDynaTrace) = False Then
                    Call M46V999.gpGraLog(2, pMensagem)
                    'Exit Function
                    End If
                End If
        End If
        'CR3939 - FIM

        'Nova tarifa 1094 / 1095 desconto de 10% para os corretores Celisa e Pecus
'        If (pStcItem.ObjPai.ColecaoPedCorr(1).Cod_Corr = 921006 Or pStcItem.ObjPai.ColecaoPedCorr(1).Cod_Corr = 346302) And _
'            pStcItem.Cod_Ramo = 114 And _
'            pStcItem.Cod_Iden_Tarif >= 1094 Then
'            lCoef_Desc_Coml = 10
'        End If

        'Percentual de Comissão
        If (M46V999.gCodProduto < 1076 Or pStcItem.Cod_Plano > 1000 _
        Or (pStcItem.Cod_Ramo <> e_CodRamo_Residencial And M46V999.gCodProduto < 1080) _
        ) And pStcItem.ObjPai.getPlano.EhPlanoSegmentado = False Then
            Select Case lCod_Ramo
            Case 112
                mPer_Coml = M46V999.gfTrunca(0.8 / IIf(lPer_Coml = 100, 1, (1 - (lPer_Coml / 100))), 5)
            Case 113, 114
                mPer_Coml = M46V999.gfTrunca(0.75 / IIf(lPer_Coml = 100, 1, (1 - (lPer_Coml / 100))), 5)
            Case 150
                mPer_Coml = M46V999.gfTrunca(0.85 / IIf(lPer_Coml = 100, 1, (1 - (lPer_Coml / 100))), 5)
            Case Else
                mPer_Coml = M46V999.gfTrunca(0.8 / IIf(lPer_Coml = 100, 1, (1 - (lPer_Coml / 100))), 5)
            End Select
        Else
            mPer_Coml = mfRetornarCoeficiente(pStcItem, 1, lPer_Coml, "Ramosdiversos.dbo.")
            Dim mPer_Coml_Corr As Double
            mPer_Coml_Corr = mfRetornarCoeficienteCorretor(pStcItem, lPer_Coml, "Ramosdiversos.dbo.")
            If mPer_Coml_Corr > 0 Then mPer_Coml = mPer_Coml_Corr
        End If

        If pStcItem.Tab_Clasf >= 2178 And pStcItem.Tab_Clasf <= 2182 And pStcItem.Cod_Ramo = 112 And pStcItem.Tip_Emissao = e_TipEmissao_RenYasuda Then
            mPer_Coml = 1
        End If

        'Tarifa 11.12 31/03/2016 Item 2.3
        If UCase(App.EXEName) Like "D46O003*" And M46V999.gCod_Ramo = 114 And M46V999.gCod_Corretor = 177804 Then
            If lPer_Coml <= 40 Then
                lCoef_Desc_Coml = 20
            Else
                lCoef_Desc_Coml = 0
            End If
        End If
        'FIM Tarifa 11.12 31/03/2016 Item 2.3
        
        'Tarifa 11.14 29/04/2016
        
        'Buffalo
        If UCase(App.EXEName) Like "D46O003*" And M46V999.gCod_Ramo = 113 And _
           (M46V999.gCod_Corretor = 918390) Then
            If lPer_Coml <= 40 Then
                lCoef_Desc_Coml = 15
            Else
                lCoef_Desc_Coml = 0
            End If
        End If
        
        'Tarifa 1178 - 10 Inclusão do Desconto Residencial para Corretora 0889199 - PRISMINAS
        If M46V999.gCod_Ramo = 113 And _
            M46V999.gCodProduto >= 1178 And _
            M46V999.gCod_Corretor = 889199 And _
            (M46V111.gCod_plano = 10 Or M46V111.gCod_plano = 30) Then
            If lPer_Coml <= 25 Then
                lCoef_Desc_Coml = 15
            Else
                lCoef_Desc_Coml = 0
            End If
        End If
        
'        If UCase(App.EXEName) Like "D46O003*" And M46V999.gCod_Ramo = 113 And _
'           (M46V999.gCod_Corretor = 907677 Or M46V999.gCod_Corretor = 929745 Or M46V999.gCod_Corretor = 931081 Or _
'           M46V999.gCod_Corretor = 932688 Or M46V999.gCod_Corretor = 932732 Or M46V999.gCod_Corretor = 932733 Or _
'           M46V999.gCod_Corretor = 930550) Then
'            If lPer_Coml <= 35 Then
'                    lCoef_Desc_Coml = 10
'                Else
'                    lCoef_Desc_Coml = 0
'            End If
'        End If
        
        'ATHINA 'Tarifa 11.72 - Item 7 Condição de Desconto Comercial - Projeto Weg (Corretora Áthina)
        'Códigos         907677, 929745, 931081, 932688, 932732, 932733 e 930550
        'Descontos       10% Desconto para Apartamentos Habituais
        '                20% Desconto para Casas Habituais
        'Comissão        Limitado ao máximo de 35%
        'Região          Território Nacional
        'Parcelamento    10x sem Juros
        '17/07/2018 - Ajuste (Inclusão da Renovação Congenere) -
        '18/07/2018 - Ajuste (Inclusão de todos os tipos de emissão menos Endossos de Exclusão e Cancelmanento -
        'If pStcItem.Tip_Emissao = e Or pStcItem.Tip_Emissao = e_TipEmissao_RenYasuda Or pStcItem.Tip_Emissao = e_TipEmissao_RenCongenere Then
        
        If Not (M46V111.gCod_Dizer = 4 Or M46V111.gCod_Dizer = 10) Then 'Endosso de Exclusão e Cancelamento.
            If UCase(App.EXEName) Like "D46O003*" And (M46V999.gCod_Ramo = 113 And M46V111.gCod_plano = 10) And _
               (M46V999.gCod_Corretor = 907677 Or M46V999.gCod_Corretor = 929745 Or M46V999.gCod_Corretor = 931081 Or _
               M46V999.gCod_Corretor = 932688 Or M46V999.gCod_Corretor = 932732 Or M46V999.gCod_Corretor = 932733 Or _
               M46V999.gCod_Corretor = 930550) Then
               
                lCoef_Desc_Coml = 0
                
                If lPer_Coml <= 35 Then
                    Select Case lCod_Clasf
                        Case "910": 'CASA HABITUAL
                            lCoef_Desc_Coml = 20
                        Case "918": 'APARTAMENTO HABITUAL
                            lCoef_Desc_Coml = 10
                        Case Else
                            lCoef_Desc_Coml = 0
                    End Select
                End If
            End If
        End If

        
        'ARBOS
        If UCase(App.EXEName) Like "D46O003*" And M46V999.gCod_Ramo = 113 And M46V999.gCod_Corretor = 912725 And M46V111.gCod_plano = 10 Then
            If lPer_Coml <= 30 Then
                lCoef_Desc_Coml = 15
            Else
                lCoef_Desc_Coml = 0
            End If
        End If
        
        'RIO CLARO / GSP
        If UCase(App.EXEName) Like "D46O003*" And M46V999.gCod_Ramo = 113 And _
           (M46V999.gCod_Corretor = 909306 Or M46V999.gCod_Corretor = 930786) Then
            If lPer_Coml <= 35 Then
                lCoef_Desc_Coml = 15
            Else
                lCoef_Desc_Coml = 0
            End If
        End If
        
        'FIM Tarifa 11.14 29/04/2016
        
        '11.48 Tarifa Residencial 2.Parametrização na Web para a corretora Tresoli - Código 922048
        If UCase(App.EXEName) Like "D46O003*" And M46V999.gCod_Ramo = 113 And M46V111.gCod_plano = 10 And _
           M46V999.gCod_Corretor = 922048 Then
            If lPer_Coml <= 30 Then
                lCoef_Desc_Coml = 10
            Else
                lCoef_Desc_Coml = 0
            End If
        End If
        
        '11.32 6.  Aplicação de desconto para renovações SOMPO - Corretora Seguratec (Código 1601)
        If UCase(App.EXEName) Like "D46O003*" And M46V999.gCodProduto >= 1132 And M46V999.gCod_Ramo = 113 And _
           pStcItem.Tip_Emissao = e_TipEmissao_RenYasuda And M46V999.gCod_Corretor = 1601 Then
            lCoef_Desc_Coml = 15
        End If
        '11.32 6. FIM
        
        'Tarifa 11.20 01/07/2016
        If M46V999.gCodProduto >= 1120 Then
            'Mila Corretora de Seguros
            If UCase(App.EXEName) Like "D46O003*" And M46V999.gCod_Ramo = 113 And M46V999.gCod_Corretor = 437698 And _
               M46V111.gCod_plano < 1000 And (lCod_Clasf = "910" Or lCod_Clasf = "918") Then
                If lPer_Coml <= 30 Then
                    lCoef_Desc_Coml = 10
                Else
                    lCoef_Desc_Coml = 0
                End If
            End If
        
            'Diamond Corretora de Seguros
            If UCase(App.EXEName) Like "D46O003*" And M46V999.gCod_Ramo = 113 And M46V111.gCod_plano < 1000 And M46V999.gCod_Corretor = 918630 And M46V111.gCod_Dizer = 0 Then
                If lPer_Coml <= 30 Then
                    Select Case lCod_Clasf
                        Case "910": 'CASA HABITUAL
                            lCoef_Desc_Coml = 20
                        Case "918": 'APARTAMENTO HABITUAL
                            lCoef_Desc_Coml = 10
                        Case Else
                            lCoef_Desc_Coml = 0
                    End Select
                
                Else
                    lCoef_Desc_Coml = 0
                End If
            End If
            
            'DBR Corretora de Seguros
            If UCase(App.EXEName) Like "D46O003*" And M46V999.gCod_Ramo = 113 And M46V999.gCod_Corretor = 918658 And _
               M46V111.gCod_plano < 1000 And (lCod_Clasf = "910" Or lCod_Clasf = "918") Then
                If lPer_Coml <= 30 Then
                    lCoef_Desc_Coml = 10
                Else
                    lCoef_Desc_Coml = 0
                End If
            End If
        End If
        'FIM Tarifa 11.20 01/07/2016
        
        'Tarifa 11.52   3 - Aplicação de desconto corretora Solidariedade
        If M46V999.gCodProduto >= 1152 And UCase(App.EXEName) Like "D46O003*" And M46V999.gCod_Ramo = 113 And _
            M46V111.gCod_plano = 10 And M46V999.gCod_Corretor = 933832 And M46V111.gCod_Dizer = 0 Then
            If lPer_Coml <= 30 Then
                Select Case lCod_Clasf
                    Case "910": 'CASA HABITUAL
                        lCoef_Desc_Coml = 20
                    Case "918": 'APARTAMENTO HABITUAL
                        lCoef_Desc_Coml = 10
                    Case Else
                        lCoef_Desc_Coml = 0
                End Select
            
            Else
                lCoef_Desc_Coml = 0
            End If
        End If
        
        'Caso o digitador tenha feito alteracao manual da taxa ela vai entrar aqui
        TxAlterada = False
        If Trim(UCase(lobjstcCobertura.Alterar_Cobert)) = "ALTERADO" And lobjstcCobertura.Per_Taxa <> 0 Then
            TxAlterada = True
            'Calcular Prêmio Comercial
            'lTaxaMin = lobjstcCobertura.Per_Taxa
            'lTaxaFinal = M46V999.gfTrunca(lTaxaMin * mPer_Coml, 5)
            'lTaxaReser = lTaxaFinal
            'lTaxaRess = lTaxaFinal
            'GoTo CalculaPremio
        End If

        For i = 1 To lQtdeTaxas
            If lArrayTaxas(0, i) = lCod_Cobert Then
                lClasseLMG = lArrayTaxas(1, i)
                lCodFaixa_Reg = lArrayTaxas(2, i)
                lCod_Classe = lArrayTaxas(3, i)
                lTaxaMin = lArrayTaxas(4, i)
                lTaxaReser = lArrayTaxas(5, i)
                lTaxaRess = lArrayTaxas(6, i)
                lTaxaMinFormula = lArrayTaxas(7, i)
                lAgravoDesconto = lArrayTaxas(8, i)
                Exit For
            End If
        Next i

        'Verifica se existe taxa para essa cobertura
        'Caso esteja zerada vai procurar a cobertura principal
        If lTaxaMin = 0 Then
            lSelect = " Select *  from Tbm_Cob_Dep_Taxa WITH (NOLOCK) "
            lSelect = lSelect & "  where Cod_Ramo = " & lCod_Ramo
            lSelect = lSelect & "  and Cod_Modalidade = " & lCod_Modalidade
            lSelect = lSelect & " and Cod_Dep_Principal = " & lCod_Cobert
            lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
            lSelect = lSelect & " and Cod_Sequencia <= " & lCod_Sequencia

            If M46V999.gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, pMensagem) = False Then
                Call M46V999.gpFecha2(lbdTab_Tarifa)
                Call M46V999.gpGraLog(0, pMensagem)
                GoTo ProximaCobertura
            End If

            If Not lrsTab_Tarifa.EOF Then
                For i = 1 To lQtdeTaxas
                    If lArrayTaxas(0, i) = lrsTab_Tarifa("Cod_Principal") Then
                        lClasseLMG = lArrayTaxas(1, i)
                        lCodFaixa_Reg = lArrayTaxas(2, i)
                        lCod_Classe = lArrayTaxas(3, i)
                        lTaxaMin = lArrayTaxas(4, i)
                        lTaxaReser = lArrayTaxas(5, i)
                        lTaxaRess = lArrayTaxas(6, i)
                        lTaxaMinFormula = lArrayTaxas(7, i)
                        lAgravoDesconto = lArrayTaxas(8, i)
                        Exit For
                    End If
                Next i
            End If

        End If
        lCoefIs = 1
        'CR7348 - Para o plano BRFoods, todas as coberturas devem ter o mesmo coeficiente da básica
        If lCod_Ramo = 112 And (lCod_Modalidade = 2218 Or lCod_Modalidade = 2219 Or (lCod_Modalidade >= 2228 And lCod_Modalidade <= 2271)) And lCod_Cobert <> 1 Then
            ' Não atualizar coeficiente - ele já está conforme a cobertura básica
        Else
            'Coeficiente de IS VR
            If lobjstcCobertura.gCod_Grupo_Cobert = 1 Or lobjstcCobertura.gCod_Grupo_Cobert = 5 Then
                lVal_VR = pStcItem.Val_Vr
                If Not BuscaCoef_IS_VR(pPropTabela, pNomeTabImport, pStcItem, lobjstcCobertura.gCod_Grupo_Cobert, lCod_Cobert, pStcItem.gISBasica, lVal_VR, pMensagem, tagDynaTrace) Then
                    Call M46V999.gpGraLog(0, pMensagem)
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If
            End If
        End If
        'CR7348 - Fim

        'Vai procurar Desconto Comercial na tabela de Ajustes pontuais
        'If mfAjustesPontuais(e_Tip_Valor_DescComl, TpoValor, lCoef_Desc_Coml, pMensagem, pStcItem) = False Then
        '   Call M46V999.gpGraLog(0, pMensagem)
        '   GoTo SAIDA
        'End If

        'Encontrar Coeficiente de Franquia
        If lobjstcCobertura.Val_Fator_Franq > 0 Then
            lCoef_Franquia = lobjstcCobertura.Val_Fator_Franq
        Else
            lCoef_Franquia = 1
        End If


        'Calcular Prêmio Net
        lTaxaNet = lTaxaMin * lCoef_Franquia * lCoefIs
        lTaxaReser = lTaxaReser * lCoef_Franquia * lCoefIs
        lTaxaRess = lTaxaRess * lCoef_Franquia * lCoefIs


        'Encontrar Coeficiente do Período Indenitário
        If mpFindCoefPI(lCoef_PI, lobjstcCobertura.Num_Meses_PI, pMensagem) = False Then
            Call M46V999.gpGraLog(0, pMensagem)
        End If

        '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 11
        If (pStcItem.ObjPai.Cod_Produto >= 1038) And _
           (pStcItem.ObjPai.Cod_Ramo = e_CodRamo_Empresarial) And _
           (pStcItem.ObjPai.Cod_Estip >= 2170 And pStcItem.ObjPai.Cod_Estip <= 2177) And _
           (lobjstcCobertura.Cod_Cobert = M46V999.gfBuscarDePara(39, 1)) Then
            lCoef_PI = 1
        End If

        'Calcular o Prêmio Comercial com o Período Indenitário
        lTaxaNet = lTaxaNet * lCoef_PI
        lTaxaReser = lTaxaReser * lCoef_PI
        lTaxaRess = lTaxaRess * lCoef_PI

        'Calcular Prêmio com Descontro Progressivo
        If (M46V999.gCodProduto < 1080 And lobjstcCobertura.Cod_Grupo_Cobertura <> 1 And lobjstcCobertura.Cod_Grupo_Cobertura <> 5) Or _
            (M46V999.gCodProduto >= 1080 And lobjstcCobertura.Cod_Grupo_Cobertura <> 4) Then
            
            If lCoef_Desc_Prog <> 0 Then
                If lCoef_Desc_Prog <> 0 And pStcItem.ObjPai.Cod_Estip < 2000 Then
                    lTaxaNet = lTaxaNet * (1 - (lCoef_Desc_Prog / 100))
                    lTaxaReser = lTaxaReser * (1 - (lCoef_Desc_Prog / 100))
                    lTaxaRess = lTaxaRess * (1 - (lCoef_Desc_Prog / 100))
                ElseIf lCoef_Desc_Prog <> 0 And pStcItem.ObjPai.Cod_Estip > 2000 And lobjstcCobertura.Cod_Grupo_Cobertura <> 1 Then
                        lTaxaNet = lTaxaNet * (1 - (lCoef_Desc_Prog / 100))
                        lTaxaReser = lTaxaReser * (1 - (lCoef_Desc_Prog / 100))
                        lTaxaRess = lTaxaRess * (1 - (lCoef_Desc_Prog / 100))
                End If
            End If
            If lobjstcCobertura.gCod_Grupo_Cobert = 3 Then
                lTaxaNet = lTaxaNet * lCoef_LMG_LC
                lTaxaReser = lTaxaReser * lCoef_LMG_LC
                lTaxaRess = lTaxaRess * lCoef_LMG_LC
            End If
        End If

        'Buscar Bônus / Desconto Experiência
        '<POG> ESSE IF EXISTE SOMENTE PARA MANTER OS DADOS QUE FORAM GERADOS COM O FONTE ERRADO. Remover tudo que esta dentro da tag POG no dia 27/07
'            If (UCase(App.EXEName) = "D46O003" _
'                And pStcItem.ObjPai.Dat_Pri_Calculo >= 20120625 And pStcItem.ObjPai.Dat_Pri_Calculo <= 20120627) Then
'
'                If (lCod_Modalidade >= 2001 And lCod_Modalidade <= 2031) Or (lCod_Modalidade >= 2061 And lCod_Modalidade <= 2100) Or (lCod_Modalidade >= 2111 And lCod_Modalidade <= 2130) Then
'                    lCoef_Desc_Coml = pStcItem.getPlanoItem.ValorRegra(eRegraPlano_DescontoComercial)
'                Else
'                    lCoef_Desc_Coml = lobjstcCobertura.Per_Desc_Coml
'                End If
'
'            Else
                '</POG>
                lCoef_Bonus = pStcItem.Per_Bonus
                '<POG>
'            End If
            '</POG>
            
        If lzeraDescAgravCober And lobjstcCobertura.Cod_Grupo_Cobertura = 4 And lCod_Tarifa >= 1090 Then  'nova tarifa 1090 / 1091
            lCoef_Bonus = 0
            lCoef_Desc_Coml = 0
        End If

        'Calcular Prêmio com Bônus
        lTaxaNet = lTaxaNet * (1 - (lCoef_Bonus / 100))
        lTaxaReser = lTaxaReser * (1 - (lCoef_Bonus / 100))
        lTaxaRess = lTaxaRess * (1 - (lCoef_Bonus / 100))

        'Calcular Prêmio Comercial
        'Resseguro e Tarifário não faz
        lTaxaComl = M46V999.gfTrunca(lTaxaNet * mPer_Coml, 5)



'    *******************************************************************************************************************
'    *#20180829
'    *Finalidade: Apura Desconto Comercial Congenere
'     Conforme regra de negocio o Desconto Comercial do Congenre deve ser adicionado aos outros descontos concedidos
'    *******************************************************************************************************************
        lCoef_Desc_Coml = lCoef_Desc_Coml + lCoef_Desc_Coml_Congenere


        'Calcular Prêmio Final
        If lCod_Modalidade = 2218 Or lCod_Modalidade = 2219 Or (lCod_Modalidade >= 2228 And lCod_Modalidade <= 2271) Or lCod_Tarifa >= 1062 Then
            lTaxaFinal = M46V999.gfTrunca(Round(lTaxaComl * (1 - (lCoef_Desc_Coml / 100)), 5), 5)
        Else
            lTaxaFinal = M46V999.gfTrunca(lTaxaComl * (1 - (lCoef_Desc_Coml / 100)), 5)
        End If
        lCoef_Protecao = 0
        'Encontrar Coeficiente de Fator de Proteção
        If mpFindCoefProtecao(lCoef_Protecao, lCoef_ProtecaoResseg, pStcItem.ColecaoPedFator, pMensagem, tagDynaTrace, pStcItem.ObjPai.getPlano) = False Then
            Call M46V999.gpGraLog(0, pMensagem)
        End If

        If M46V999.gCodProduto >= 934 Then
            If Not gfVerParametro(8, pStcItem.ObjPai, pStcItem, lobjstcCobertura, lretorno, pMensagem) Then
                Call M46V999.gpGraLog(1, pMensagem)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If
            If M46V999.gfPreenchido(lretorno) Then
                If lCoef_Protecao > val(lretorno) Then
                    lCoef_Protecao = lretorno
                End If
            End If
            If M46V999.gCodProduto >= 942 And pStcItem.Cod_Ramo = e_CodRamo_Empresarial Then
                If lCoef_Protecao > 0 And lobjstcCobertura.Cod_Cobert = 1 Then
                    If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 310, pStcItem, pMensagem, tagDynaTrace) Then
                        Call M46V999.gpGraLog(1, pMensagem)
                        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                        GoTo SAIDA
                    End If
                End If
                If lCoef_Protecao > 0 And lobjstcCobertura.Cod_Cobert = M46V999.gfBuscarDePara(23, 1) Then
                    If Not mfIncluirClausula(pPropTabela, pNomeTabImport, 311, pStcItem, pMensagem, tagDynaTrace) Then
                        Call M46V999.gpGraLog(1, pMensagem)
                        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                        GoTo SAIDA
                    End If
                End If
            End If
        Else
            If M46V999.gCodProduto >= 932 And lCoef_Protecao > 15 Then
                lCoef_Protecao = 15
            End If
        End If

        'Ficha 1000304
        If (lobjstcCobertura.Cod_Cobert = M46V999.gfBuscarDePara(23, 1) Or lobjstcCobertura.Cod_Cobert = M46V999.gfBuscarDePara(63, 1)) Then
            If fAchouShopping(pStcItem) = True And (lCod_Modalidade < 2220 Or lCod_Modalidade > 2227) Then
                lCoef_Protecao = 30
            End If

            'Ajuste no coef de proteção pois no primeiro calculo o pStcItem.ColecaoPedFator esta vindo vazio.
            If pStcItem.ColecaoPedFator.Count = 0 Then
                If (lCod_Modalidade >= 2194 And lCod_Modalidade <= 2197 And pStcItem.Cod_Ramo = e_CodRamo_Empresarial) Then
                    lCoef_Protecao = 30
                End If
            End If
        End If

        '-->Ficha 1000888
        If lCod_Ramo = 112 Then
            Select Case lobjstcCobertura.Cod_Cobert
            Case 1, M46V999.gfBuscarDePara(2, 1), M46V999.gfBuscarDePara(32, 1), M46V999.gfBuscarDePara(39, 1) Or M46V999.gfBuscarDePara(42, 1)
                pStcItem.Per_Tot_Protec = lCoef_Protecao
            End Select
        End If
        If lCod_Ramo = 114 And lobjstcCobertura.Cod_Cobert = 1 Then
            pStcItem.Per_Tot_Protec = lCoef_Protecao
        End If

        'Cod_Iden_Protec_10
        'If (lCod_Ramo = 112 And (lobjstcCobertura.Cod_Cobert = 23 Or lobjstcCobertura.Cod_Cobert = 63)) Or _
         '    (lCod_Ramo = 113 And (lobjstcCobertura.Cod_Cobert = 5 Or lobjstcCobertura.Cod_Cobert = 16)) Or _
         '    (lCod_Ramo = 114 And lobjstcCobertura.Cod_Cobert = 5) Then
        '    pStcItem.Per_Tot_Segur = lCoef_Protecao
        '    pStcItem.Cod_Iden_Protec_10 = lCoef_Protecao
        'End If

        'Cod_Iden_Protec_09
        'If lobjstcCobertura.Cod_Cobert = 1 Then
        '    pStcItem.Cod_Iden_Protec_09 = lCoef_Protecao
        'End If


        pStcItem.Cod_Regiao = pStcItem.gCod_RegiaoNova
        'Ficha 1000888 <--


        lTaxaFinal = lTaxaFinal * (1 - (lCoef_Protecao / 100))
        lTaxaReser = lTaxaReser * (1 - (lCoef_Protecao / 100))
        lTaxaRess = lTaxaRess * (1 - (lCoef_ProtecaoResseg / 100))

        lobjstcCobertura.Per_Desc_Coml = lCoef_Desc_Coml
        lobjstcCobertura.COD_CLASSE = lCod_Classe

        If (M46V999.gCodProduto >= M46V999.TarifaMigracaoMaritima) Then
'            If lzeraDescAgravCober And lCod_Cobert = 1215 Then
'                lPer_Desc_Agrav = 0
'            End If

            If lPer_Desc_Agrav <> 0 Then    'Desconto/Agravo
                If lPer_Desc_Agrav > 0 Then    'Agravo
                    lTaxaFinal = lTaxaFinal * (1 + (lPer_Desc_Agrav / 100))
                    lTaxaReser = lTaxaReser * (1 + (lPer_Desc_Agrav / 100))
                    lTaxaRess = lTaxaRess * (1 + (lPer_Desc_Agrav / 100))
                Else
                    lTaxaFinal = lTaxaFinal * (1 - (lPer_Desc_Agrav / 100))
                    lTaxaReser = lTaxaReser * (1 - (lPer_Desc_Agrav / 100))
                    lTaxaRess = lTaxaRess * (1 - (lPer_Desc_Agrav / 100))
                End If
            End If

        End If

CalculaPremio:
        lPremioAnt = 0
        EhLmi = False
        If TxAlterada Or M46V999.gfPreenchido(pStcItem.ObjPai.Num_Cotac_LMI) And Not pStcItem.ObjPai.Cod_Estip = 20 Then
            'Calcular Prêmio Comercial qdo a taxa foi alterada
            If M46V999.gCod_Ramo = e_CodRamo_Roubos Then
                'Ficha 0901367, não permitir que taxa alterada seja menor que TaxaCalculada
                If lobjstcCobertura.Per_Taxa < lTaxaFinal Then
                    TxAlterada = False
                    lobjstcCobertura.Alterar_Cobert = ""
                Else
                    lTaxaMin = lobjstcCobertura.Per_Taxa
                    lTaxaComl = M46V999.gfTrunca(lTaxaMin * mPer_Coml, 5)
                    lTaxaFinal = M46V999.gfTrunca(lTaxaComl * (1 - (lCoef_Desc_Coml / 100)), 5)
                End If
            Else
                If M46V999.gfPreenchido(pStcItem.ObjPai.Num_Cotac_LMI) And pStcItem.ObjPai.Num_Cotac_LMI <> "00000000000000000000" Then
                    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, bdPremio, pMensagem) = False Then
                        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                        GoTo SAIDA
                    End If
                    lSelect = "select sum(val_pr) as soma , max(per_DESC_COML) AS per_DESC, "
                    lSelect = lSelect & " (select sum(per_corr) from tab_ped_corr WITH (NOLOCK) where nosso_numero = '" & pStcItem.ObjPai.Num_Cotac_LMI & "')  AS COMIS "
                    lSelect = lSelect & " from tab_ped_cobert WITH (NOLOCK) where nosso_numero = '" & pStcItem.ObjPai.Num_Cotac_LMI & "'"
                    lSelect = lSelect & " and cod_cobert = " & lobjstcCobertura.Cod_Cobert
                    If Not M46V999.gfObtRegistro(bdPremio, lSelect, lrst, pMensagem) Then
                        Call M46V999.gpGraLog(1, pMensagem)
                        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                        GoTo SAIDA
                    End If
                    lPremioAnt = 0
                    If Not lrst.EOF Then
                        EhLmi = True
                        If IsNull(lrst("soma")) Then
                            pMensagem = "AGE0022 - Cotação LMI não localizada."
                            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                            GoTo SAIDA
                        End If
                        lPremioAnt = M46V999.gfTratarNulo(lrst("soma"), tpoNumerico)
                        'Percentual de Comissão
                        If (M46V999.gCodProduto <= 1075) Or (pStcItem.Cod_Plano > 1000) Then
                            Select Case lCod_Ramo
                            Case 112
                                lPer_Coml_Cota = M46V999.gfTrunca(0.8 / (1 - (lrst("Comis") / 100)), 5)
                            Case 113, 114
                                lPer_Coml_Cota = M46V999.gfTrunca(0.75 / (1 - (lrst("Comis") / 100)), 5)
                            Case 150
                                lPer_Coml_Cota = M46V999.gfTrunca(0.85 / (1 - (lrst("Comis") / 100)), 5)
                            Case Else
                                lPer_Coml_Cota = M46V999.gfTrunca(0.8 / (1 - (lrst("Comis") / 100)), 5)
                            End Select
                        Else
                            lPer_Coml_Cota = mfRetornarCoeficiente(pStcItem, 1, lrst("Comis"), "Ramosdiversos.dbo.")
                        End If

                        lobjstcCobertura.Per_Taxa = M46V999.gfTrunca(lPremioAnt / lVal_IS * 100 / lPer_Coml_Cota / (1 - (lrst("Per_DEsc") / 100)), 5)
                        If lrst("Comis") <> lPer_Coml Or lrst("Per_DEsc") <> lCoef_Desc_Coml Then
                            lPremioAnt = 0
                        End If
                        'lobjstcCobertura.Per_Taxa = M46V999.gfTrunca(lPremioAnt / lVal_IS * 100 / lPer_Coml_Cota, 5)
                        'If lrst("Comis") <> lPer_Coml Then 'Or lrst("Per_DEsc") <> lCoef_Desc_Coml Then
                        '    lobjstcCobertura.Per_Taxa = M46V999.gfTrunca(lobjstcCobertura.Per_Taxa * mPer_Coml, 5)
                        '    lPremioAnt = M46V999.gfTrunca(lVal_IS * (lobjstcCobertura.Per_Taxa / 100), 2)
                        'End If

                        TxAlterada = True
                    End If
                End If
                lTaxaMin = lobjstcCobertura.Per_Taxa
                lTaxaComl = M46V999.gfTrunca(lTaxaMin * mPer_Coml, 5)
                lTaxaFinal = M46V999.gfTrunca(lTaxaComl * (1 - (lCoef_Desc_Coml / 100)), 5)
            End If
        End If

        'Consiste Taxa Final
        If lTaxaFinal = 0 Then
            If EhLmi = False Then
                pMensagem = "AGE0022 - Cobertura " & lCod_Cobert & " taxa não encontrada ou taxa zerada! Solicitar Cotação."
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If
        End If


        'Calcular Prêmio Anual
        lPremio = 0
        If lCod_Modalidade = 2218 Or lCod_Modalidade = 2219 Or (lCod_Modalidade >= 2228 And lCod_Modalidade <= 2271) Or lCod_Tarifa >= 1062 Then
            lPremio = M46V999.gfTrunca(Round(lVal_IS * (lTaxaFinal / 100), 2), 2)
        Else
            lPremio = M46V999.gfTrunca(lVal_IS * (lTaxaFinal / 100), 2)
        End If
        If lPremioAnt <> 0 Then
            lPremio = lPremioAnt
        End If

        'Aplica desconto de LMI

        If pStcItem.ObjPai.LMI_Unico = 1 Then
            'Busco o desconto LMI
            If pStcItem.ObjPai.Per_Desc_LMI <> 0 Then
                ' O desconto já está definido, então não buscar da tabela
            ElseIf pStcItem.ObjPai.Per_Desc_LMI = 0 Or (M46V999.gTipUsuario = e_Tipo_Usu_Corretor And pStcItem.ObjPai.mStatusDivSiscota <> enum_StatusTransm_Aprovada) Then
                If Not gfVerParametro(12, pStcItem.ObjPai, pStcItem, lobjstcCobertura, lretorno, pMensagem) Then
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                    Exit Function
                End If
                
                ' Procurar o desconto na nova tabela de descontos de LMI
                pStcItem.ObjPai.Per_Desc_LMI = mfRetornarDescontoLMI(pStcItem)
            End If
            M46V999.gNovaPropostaLMI = True
        End If

        If M46V999.gNovaPropostaLMI = True And M46V999.PerfilAcesso = e_Perfil_CALC Then
            'Consistências para permitir ser ou não ser LMI
            If pStcItem.ObjPai.ColecaoPedLoc.Count = 1 And pStcItem.Tip_Emissao <> e_TipEmissao_Endosso Then
                pMensagem = "AGE0122 - Não será permitido cotar como LMI pois este cálculo só possui um item."
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If
        End If

        If M46V999.gNovaPropostaLMI Then
            If pStcItem.ObjPai.Per_Desc_LMI > 0 And M46V999.gfPreenchido(pStcItem.ObjPai.Num_Cotac_LMI) = False Then
                'Debug.Print pStcItem.Num_Item & "-" & lCod_Cobert & "-" & lPremio
                lPremio = lPremio * ((100 - pStcItem.ObjPai.Per_Desc_LMI) / 100)
                'Debug.Print lPremio
            End If
        End If

        lobjstcCobertura.Val_Pr_Anu = lPremio

        'Projeto Endosso
        'If M46V111.gCod_Dizer <> 0 Then
        '    Call mpCalcula_Endosso(pStcItem.ObjPai.Num_Apol_An, pStcItem.Num_Item_Ant, "COML", lPremio)
        'End If
        
        'Debug.Assert lCod_Cobert <> 1
        If Not M46V999.gfPreenchido(pStcItem.ObjPai.Num_Cotac_LMI) Then
            lPremio = lPremio * lCoePrazo
        End If
        
        lCustoInspecao = 0
        lCustoInspecaoFinal = 0
        If pStcItem.ObjPai.Cod_Estip <> 20 Then ' Não aplicar para o imobiliario
            If (pStcItem.ObjPai.Cod_Dizer = 0 Or pStcItem.Cod_DizerItem = 1) And pStcItem.ObjPai.Cod_Estip < 2000 Then
                Select Case lCod_Ramo
                    Case 112
                        lCustoInspecao = 8.56
                    Case 113
                        '11.44
                        If pStcItem.ObjPai.Cod_Produto < 1144 Then
                            lCustoInspecao = 0.5
                        Else
                            If pStcItem.ObjPai.Cod_Estip = 30 Then
                                'Tarifa 11.68 - Item 13 Alteração no Custo Fixo de Inspeção do Produto Residencial
                                'lCustoInspecao = 65.8
                                lCustoInspecao = 62.3
                            Else
                                'Tarifa 11.68 - Item 13 Alteração no Custo Fixo de Inspeção do Produto Residencial
                                'lCustoInspecao = 7
                                lCustoInspecao = 3.5
                            End If
                        End If
                    Case 114
                        lCustoInspecao = 32.27
                End Select
            End If
        End If
        
        mPer_ComlAss24hrs = 0
        If mfAjustesPontuais(e_Tip_Valor_ComissaoAss24hrs, tpoNumerico, mPer_ComlAss24hrs, pMensagem, pStcItem, 0, tagDynaTrace) = False Then
            Call M46V999.gpGraLog(0, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        
        If lobjstcCobertura.Cod_Cobert = 1 Then
            If pStcItem.ObjPai.Cod_Produto >= 1098 Then
                lCustoInspecaoFinal = (lCustoInspecao * mPer_Coml)
            Else
                lCustoInspecaoFinal = 0
            End If
            
            If pStcItem.ObjPai.Cod_Estip <> 20 Then
                If mPer_ComlAss24hrs <> 0 Then
                    lPremio = lPremio + (pStcItem.gVal_PR_Ass24h * (mPer_ComlAss24hrs / 100)) + lCustoInspecaoFinal
                Else
                    lPremio = lPremio + (pStcItem.gVal_PR_Ass24h * mPer_Coml) + lCustoInspecaoFinal
                End If
            End If
        End If


        'Rotina Nova Premio Minimo

        UtilPremioMinino = False
        If mfAjustesPontuais(e_Tip_Valor_PremioMinimo, tpoNumerico, lPremioPont, pMensagem, pStcItem, 0, tagDynaTrace) = False Then
            Call M46V999.gpGraLog(0, pMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If

        If (M46V999.gCodProduto >= M46V999.TarifaMigracaoMaritima) Then

            'lobjstcCobertura.gVal_Prem_Min = Premio Minimo da cobertura
            If lPremioPont < lobjstcCobertura.gVal_Prem_Min Then
                lPremioPont = lobjstcCobertura.gVal_Prem_Min
            End If

            If lPremio < lPremioPont Then
                UtilPremioMinino = True
                lPremio = lPremioPont
            End If

        End If

        ' A esta altura, já temos o prêmio da cobertura Calculado.
        ' Se o premio for informado, então temos que substituir o valor do premio pelo valor informado
        ' Mas temos que guardar o valor calculado no campo Val_Pr_Orig
        If lobjstcCobertura.PremioInformado = True Then
            lobjstcCobertura.Val_Pr_Orig = lPremio
            lPremio = ValPremioInformado
            
            If pStcItem.ObjPai.Cod_Estip = 20 And lobjstcCobertura.Cod_Cobert = 1 Then
                lPremio = lPremio - pStcItem.gVal_PR_Ass24h
            End If
        End If

        lobjstcCobertura.VAL_PR = lPremio

        'REALIZA AS VERIFICAÇÕES E CALCULA OS DESCONTOS / AGRAVA NA FRANQUIA E PREMIO
        Call RetornaVlrMultFranquia(lobjstcCobertura)

        If pStcItem.Cod_DizerItem = 0 Then
            If lPremio = 0 Then
                pMensagem = "AGE0051 -  Erro no cálculo do prêmio da cobertura " & lCod_Cobert & ". Solicitar cotação."
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If
        End If

        'Calcular Prêmio Tarifário
        'Resseguro-sandra
        If pStcItem.ObjPai.Dat_Ini_Vig >= 20090501 Then
            lobjstcCobertura.Val_Pr_Tarif_Anu = lobjstcCobertura.Val_Pr_Anu
        Else
            lPremio = M46V999.gfTrunca(lVal_IS * (lTaxaReser / 100), 2)
            lobjstcCobertura.Val_Pr_Tarif_Anu = lPremio
        End If

        'Projeto Endosso
        'If M46V111.gCod_Dizer <> 0 Then
        '    Call mpCalcula_Endosso(pStcItem.ObjPai.Num_Apol_An, pStcItem.Num_Item_Ant, "TARI", lPremio)
        'End If

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
            lPremio = M46V999.gfTrunca(lVal_IS * (lTaxaRess / 100), 2)
            lobjstcCobertura.Val_Pr_Resg_Anu = lPremio
        End If

        'Projeto Endosso
        'If M46V111.gCod_Dizer <> 0 Then
        '    Call mpCalcula_Endosso(pStcItem.ObjPai.Num_Apol_An, pStcItem.Num_Item_Ant, "RESG", lPremio)
        'End If

        'Resseguro-sandra
        If pStcItem.ObjPai.Dat_Ini_Vig >= 20090501 Then
            lobjstcCobertura.Val_Pr_Resg = lobjstcCobertura.VAL_PR
        Else
            lPremio = lPremio * lCoePrazo
            lobjstcCobertura.Val_Pr_Resg = lPremio
        End If

        If Not TxAlterada Then
            lobjstcCobertura.Per_Taxa = lTaxaMin
            ''lobjstcCobertura.Formula = "CL" & lCod_Classe & ";Tx=" & lTaxaMin & ";Frq=" & lCoef_Franquia & ";Prot=" & lCoef_Protecao & ";Bon=" & lCoef_Bonus & ";PI=" & lCoef_PI & ";Cms=" & mPer_Coml & ";Lmg/VR=" & lCoefIs
            lobjstcCobertura.Formula = "CL" & lCod_Classe & ";Tx=" & lTaxaMinFormula & ";AD=" & lAgravoDesconto & ";Frq=" & lCoef_Franquia & ";Prot=" & lCoef_Protecao & ";Bon=" & lCoef_Bonus & ";PI=" & lCoef_PI & ";Cms=" & mPer_Coml & ";Lmg/VR=" & lCoefIs & ";Agrav=" & lPer_Desc_Agrav
            
            If (M46V999.gCodProduto < 1080 And lobjstcCobertura.Cod_Grupo_Cobertura <> 1 And lobjstcCobertura.Cod_Grupo_Cobertura <> 5) Or _
                (M46V999.gCodProduto >= 1080 And lobjstcCobertura.Cod_Grupo_Cobertura <> 4) Then
                lobjstcCobertura.Formula = lobjstcCobertura.Formula & ";PROGR=" & lCoef_Desc_Prog
                
                If lobjstcCobertura.gCod_Grupo_Cobert = 3 Then
                    lobjstcCobertura.Formula = lobjstcCobertura.Formula & ";ISVR=" & lCoef_LMG_LC
                End If
            End If
            lobjstcCobertura.Formula = lobjstcCobertura.Formula & ";DCOML=" & lCoef_Desc_Coml & ";TxFim=" & lTaxaFinal
            'If M46V111.gCod_Dizer <> 0 Then
            lobjstcCobertura.Formula = lobjstcCobertura.Formula & ";CoefPrazo=" & lCoePrazo
            'End If

            If UtilPremioMinino Then
                lobjstcCobertura.Formula = lobjstcCobertura.Formula & ";PremioMinimo"
            End If
        Else
            lobjstcCobertura.Per_Taxa = lTaxaMin
            lobjstcCobertura.Formula = "Taxa Alterada Manualmente=" & lTaxaMin & "; Cms=" & mPer_Coml & ";DCOML=" & lCoef_Desc_Coml & "; TxFim=" & lTaxaFinal
        End If

        lobjstcCobertura.Formula = lobjstcCobertura.Formula & ";Sin=" & pStcItem.Ind_Sinistro

        ' Se o premio é informado, precisa incluir esta informação na fórmula
        If lobjstcCobertura.PremioInformado = True Then
            lobjstcCobertura.Formula = lobjstcCobertura.Formula & ";PremioAjustado"
        End If
        
        If lCustoInspecaoFinal > 0 And lCod_Cobert = 1 Then
            lobjstcCobertura.Formula = lobjstcCobertura.Formula & ";CI=" & Round(lCustoInspecaoFinal, 2)
            pStcItem.ValCustoInspecao = Round(lCustoInspecaoFinal, 2)
        End If

        If mPer_ComlAss24hrs > 0 And pStcItem.gVal_PR_Ass24h > 0 Then
            lobjstcCobertura.Formula = lobjstcCobertura.Formula & ";Cms_Ass24hrs=" & mPer_ComlAss24hrs
        End If

        Vlr_Pr_Atu = Vlr_Pr_Atu + lobjstcCobertura.VAL_PR

        If lobjstcCobertura.Per_Taxa > 99 Then lobjstcCobertura.Per_Taxa = 99
        
        If Not lobjCobert.mfAlterarValoresPrecificacao(lobjstcCobertura, pPropTabela, pNomeTabImport) Then
            pMensagem = lobjCobert.UltimoErro
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If

ProximaCobertura:
        'Adiciono a cobertura em uma nova coleção
        Set lobjstcCobertura.ObjPai = pStcItem
        lColecaoTemp.Add lobjstcCobertura, lobjstcCobertura.Class_ID

        '11.72 - Batente 0 Corretora Triunfo
        'Tranferencia da regra de 917581 - Triunfo para 939571 - BR Insurance
        If (pStcItem.ObjPai.Num_Apol_An > 0 And (pStcItem.ObjPai.Cod_Ramo = e_CodRamo_Residencial Or pStcItem.ObjPai.Cod_Ramo = e_CodRamo_Condominio) And pStcItem.ObjPai.Cod_Dizer = 0) _
        And (pStcItem.Tip_Emissao = e_TipEmissao_RenYasuda And (M46V999.gCod_Corretor = 939571 Or M46V999.gCod_Corretor = 930446) And pStcItem.ObjPai.Cod_Estip <> 20) Then
            If ConsisteRegraBatente(pStcItem.ObjPai, lPer_Coml, False, lrsTab_Ped_Cobert, Vlr_Pr_Ant, Vlr_Pr_Atu, pStcItem) Then
                pStcItem.ObjPai.CalBatente = True
            Else
                pStcItem.ObjPai.CalBatente = False
            End If
        End If

        If pStcItem.ObjPai.CalBatente Then
            'Adiciona a coleção de coberturas para calculo do batente e gravação dos novos valores
            pStcItem.ColecaoRegraBatente.Add lobjstcCobertura, lobjstcCobertura.Class_ID
        End If

        ''2016-10-19~2018-02-28: VerbaCom
        valPrTemp = valPrTemp + lobjstcCobertura.VAL_PR

    Next lobjstcCobertura

    ''2016-10-19~2018-02-28: VerbaCom -> Fazer o calculo proporcional, cujo original está na EMIS_COMPL
    '' Vlr_Pr_Atu
    If pStcItem.ObjPai.Qtd_Anexo > 0 Then
        valVerbaComTemp = pStcItem.ObjPai.objstcCtrlEmis.Val_VerbaCom / CDbl(pStcItem.ObjPai.Qtd_Anexo)
    Else
        valVerbaComTemp = 0
    End If
        
    If valVerbaComTemp > 0 Then
        valVerbaComDif = 0
        
        '' Aplicar VerbaCom Proporcional a cada cobertura
        For Each lobjstcCoberturaTemp In pStcItem.ColecaoPedCobert
            '' Determinar o valor proporcional
            valVerbaComProp = (lobjstcCoberturaTemp.VAL_PR / valPrTemp) * valVerbaComTemp
            
            '' Se o valor ja e minimo, colocar a VerbaComProp na DIF (que depois ira p Basica)
            If InStr(1, lobjstcCoberturaTemp.Formula, "PremioMinimo") > 0 Then
                valVerbaComDif = valVerbaComDif + valVerbaComProp
                lobjstcCoberturaTemp.Val_VerbaComProp = 0
            Else
                lobjstcCoberturaTemp.Val_VerbaComProp = valVerbaComProp
                lobjstcCoberturaTemp.VAL_PR = lobjstcCoberturaTemp.VAL_PR - valVerbaComProp
                lobjstcCoberturaTemp.Formula = lobjstcCoberturaTemp.Formula & ";VerbaComProp=" & Round(valVerbaComProp, 2)
            End If
        Next lobjstcCoberturaTemp
        
        '' Se houver DIF, reajustar a basica
        If valVerbaComDif > 0 Then
            pStcItem.ColecaoPedCobert(1).Formula = pStcItem.ColecaoPedCobert(1).Formula & ";VerbaComPropMinimo=" & Round(valVerbaComDif, 2)
            pStcItem.ColecaoPedCobert(1).Val_VerbaComProp = pStcItem.ColecaoPedCobert(1).Val_VerbaComProp + valVerbaComDif
            pStcItem.ColecaoPedCobert(1).VAL_PR = pStcItem.ColecaoPedCobert(1).VAL_PR - valVerbaComDif
        End If
               
        valPrTemp = valPrTemp - valVerbaComTemp
    End If
    '' Fim VerbaCom
       
    '11.72 - Batente 0 Corretora Triunfo
    'Tranferencia da regra de 917581 - Triunfo para 939571 - BR Insurance
    If (pStcItem.ObjPai.Num_Apol_An > 0 And (pStcItem.ObjPai.Cod_Ramo = e_CodRamo_Residencial Or pStcItem.ObjPai.Cod_Ramo = e_CodRamo_Condominio) And pStcItem.ObjPai.Cod_Dizer = 0) _
    And (pStcItem.Tip_Emissao = e_TipEmissao_RenYasuda And (M46V999.gCod_Corretor = 939571 Or M46V999.gCod_Corretor = 930446) And pStcItem.ObjPai.Cod_Estip <> 20) Then
        If ConsisteRegraBatente(pStcItem.ObjPai, lPer_Coml, False, lrsTab_Ped_Cobert, Vlr_Pr_Ant, Vlr_Pr_Atu, pStcItem) Then
            pStcItem.ObjPai.CalBatente = True
        Else
            pStcItem.ObjPai.CalBatente = False
        End If
    End If
       
    If pStcItem.Tip_Emissao = e_TipEmissao_RenYasuda Then
        If pStcItem.ObjPai.CalBatente And pStcItem.Cod_Iden_Tarif < 1080 Then
            Call RetornaVlrRgrBatente(pStcItem, pStcItem.ColecaoRegraBatente, lrsTab_Ped_Cobert, Vlr_Pr_Ant, Vlr_Pr_Atu, pPropTabela, pNomeTabImport)
        ElseIf pStcItem.ObjPai.CalBatente Then
            Call RetornaVlrRgrBatente_NovoProduto(pStcItem, pStcItem.ColecaoRegraBatente, lrsTab_Ped_Cobert)
        End If
    End If
    
    mfCalNewPrecif = True
    Set pStcItem.ColecaoPedCobert = Nothing
    Set pStcItem.ColecaoPedCobert = lColecaoTemp

    pStcItem.Per_Desc_Progres = lCoef_Desc_Prog

    If pStcItem.DAT_DIGIT = 0 Then
        pStcItem.DAT_DIGIT = Left(M46V999.gfForData(3), 8)
        pStcItem.Cod_User_Liber = M46V999.gNomUsuario6
    End If

    '' 2017-12-29: Siscota End
    If pStcItem.Dat_Liber = 0 Then
        pStcItem.Dat_Liber = Left(M46V999.gfForData(3), 8)
        pStcItem.Hor_Liber = Mid(M46V999.gfForData(3), 9, 4)
    End If

    Set lobjPedLoc = New clsA46V708A
    Set lobjPedLoc.M46V999 = M46V999

    If Not lobjPedLoc.mfAlterar(pStcItem, pPropTabela, pNomeTabImport) Then
        pMensagem = lobjPedLoc.UltimoErro
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        GoTo SAIDA
    End If

    For Each lobjstcCobertura In pStcItem.ColecaoPedCobert
        If Not lobjstcCobertura.Log_Batente Is Nothing Then
            Dim pobjLogBatente As clsA46V745A
            Set pobjLogBatente = New clsA46V745A
            Set pobjLogBatente.M46V999 = pStcItem.M46V999
        
            With lobjstcCobertura.Log_Batente
                .nosso_numero = lobjstcCobertura.nosso_numero
                .Num_Item = lobjstcCobertura.Num_Item
                .Cod_Cobert = lobjstcCobertura.Cod_Cobert
                .Formula = lobjstcCobertura.Formula
                .Usuario = M46V999.gNomUsuario
            End With
            Call pobjLogBatente.mfIncluir(lobjstcCobertura.Log_Batente, pPropTabela, pNomeTabImport, tagDynaTrace)
        End If
    Next
    
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    GoSub FechaObj
    Exit Function

TrataErro:
    pMensagem = "Erro ao calcular. Favor contactar o suporte técnico."
    Call M46V999.gpMenErro(pMensagem)
    Call M46V999.gpGraLog(0, "ERR - " & methodName & " " & Err.Description & "(" & Err.Number & ")")
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, Err.Description, 2, Class_Name, methodName)
    mfCalNewPrecif = False
    
    GoTo SAIDA

FechaObj:
    Set lobjstcCobertura = Nothing
    Set lobjstcCorretor = Nothing
    Set lbdTab_Tarifa = Nothing
    Set lrsTab_Tarifa = Nothing
    Set lobjPedLoc = Nothing
    Set lobjCobert = Nothing
    Set lColecaoTemp = Nothing
    Set bdPremio = Nothing
    Set lrst = Nothing
    Set lobjstcProlabore = Nothing
    Return
End Function

Private Function mpFindCoefLMG(ByRef pCoef_LMG As Double, ByRef pMensagem As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mpFindCoefLMG"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lbdTbm_Fator_LMG As ADODB.Connection      'Arquivo: P00Tarifa.
    Dim lrsTbm_Fator_LMG As clsYasRecordSet       'Registro : Tbm_Fator_LMG.
    Dim mCodSequencia As Integer
    Dim lSelect As String


    pMensagem = ""
    pCoef_LMG = 0
    mpFindCoefLMG = False

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Fator_LMG, pMensagem) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    mCodSequencia = lCod_Sequencia

NovaSequencia:

    '1 - Procura Registro na Tabela de Fator de Ajuste por Limite Máximo de Garantia
    ' Exceção: Faz busca exata

    lSelect = " Select Fator_Ajuste as Fator, Lmg_minimo, Lmg_maximo"
    lSelect = lSelect & " from Tbm_Fator_LMG WITH (NOLOCK) "
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

    If M46V999.gfObtRegistro(lbdTbm_Fator_LMG, lSelect, lrsTbm_Fator_LMG, pMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Fator_LMG)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    '2 - Se não encontrar vai procurar por Padrão na mesma Sequencia
    If lrsTbm_Fator_LMG.EOF Then

        GoTo BuscaCoefPadraoLMG
    Else
        Do While Not lrsTbm_Fator_LMG.EOF
            If lVal_IS >= M46V999.Check(lrsTbm_Fator_LMG("Lmg_minimo")) And _
               lVal_IS <= M46V999.Check(lrsTbm_Fator_LMG("Lmg_maximo")) Then
                pCoef_LMG = M46V999.Check(lrsTbm_Fator_LMG("Fator"))
                Exit Do
            End If
            lrsTbm_Fator_LMG.MoveNext
        Loop

        GoTo FimCoef
    End If

BuscaCoefPadraoLMG:

    '3 - Procura Registro na Tabela de Fator de Ajuste por Limite Máximo de Garantia
    ' Padrão : Busca pela Cobertura

    lSelect = " Select Fator_Ajuste as Fator, Lmg_minimo, Lmg_maximo"
    lSelect = lSelect & " from Tbm_Fator_LMG WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & lCod_Modalidade
    lSelect = lSelect & " and Cod_Tip_Reg = 'D' "                  'Default
    lSelect = lSelect & " and Cod_Cobert = " & lCod_Cobert
    lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
    lSelect = lSelect & " and Cod_Sequencia <= " & mCodSequencia
    lSelect = lSelect & " order by Lmg_minimo desc "

    If M46V999.gfObtRegistro(lbdTbm_Fator_LMG, lSelect, lrsTbm_Fator_LMG, pMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Fator_LMG)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    '4 - Se não encontrar a sequencia vai procurar sequencia anterior
    If lrsTbm_Fator_LMG.EOF Then
        If mCodSequencia > 0 Then
            mCodSequencia = mCodSequencia - 1

            GoTo NovaSequencia
        End If
    Else
        Do While Not lrsTbm_Fator_LMG.EOF
            If lVal_IS >= M46V999.Check(lrsTbm_Fator_LMG("Lmg_minimo")) And _
               lVal_IS <= M46V999.Check(lrsTbm_Fator_LMG("Lmg_maximo")) Then
                pCoef_LMG = M46V999.Check(lrsTbm_Fator_LMG("Fator"))
                Exit Do
            End If
            lrsTbm_Fator_LMG.MoveNext
        Loop

    End If

FimCoef:

    Call M46V999.gpFecha2(lbdTbm_Fator_LMG)
    Set lbdTbm_Fator_LMG = Nothing
    Set lrsTbm_Fator_LMG = Nothing
    mpFindCoefLMG = True
    '3 - Vai procurar na tabela de Ajustes pontuais
    If mfAjustesPontuais(e_Tip_Valor_CoefLMG, TpoValor, pCoef_LMG, pMensagem, Nothing, 0, tagDynaTrace) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

End Function

Public Function mfAjustesPontuais(ByVal pCod_Tip_Valor As Integer, _
                                ByVal pTip_Dado As TipoDados, _
                                ByRef ValorAjustavel As Variant, _
                                ByRef pMensagem As String, _
                                ByRef pobjstcPedLoc As stcA46V708B, _
                                VarAux As Integer, tagDynaTrace As String) As Boolean
                                
    ' Valores padrão
    ' pobjstcPedLoc = Nothing
    ' VarAux = 0

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfAjustesPontuais"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lbdTbm_Parameter As ADODB.Connection      'Arquivo: P00Tarifa.
    Dim lrsTbm_Parameter As clsYasRecordSet       'Registro : Tbm_Parameter.
    Dim mCodSequencia As Integer
    Dim lSelect As String
    Dim ValorComparacao As Double
    Dim lQtdCoberturas As Integer
    Dim lCodSegmentado As Integer
    Dim lCodClasseOcupacao As Integer

    If pCod_Tip_Valor = e_Tip_Valor_DescontoAgrupamento Then
        lQtdCoberturas = VarAux
    End If

    '1100192 - Tarifa 1006 e 1007 Marco 2011
    Dim lValCoef As Double

    On Error GoTo ControleErros

    pMensagem = ""
    mfAjustesPontuais = False
    ValorComparacao = 0#

    '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 1
    If Not pobjstcPedLoc Is Nothing Then
        lCod_Ramo = pobjstcPedLoc.Cod_Ramo
        lCod_Modalidade = pobjstcPedLoc.Cod_Produto
        If Not pobjstcPedLoc.ObjPai Is Nothing Then
            lCodSegmentado = M46V999.gCod_Segmentacao
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

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Parameter, pMensagem) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    mCodSequencia = lCod_Sequencia

NovaSequencia:

    If mCodSequencia > 0 Then
        'Procurar se existem registros para a atual sequencia
        'Se existir pelo menos 1, essa é a sequencia válida

        lSelect = " Select count(*) as Total  from Tbm_Parameter WITH (NOLOCK) "
        lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & lCod_Modalidade
        lSelect = lSelect & " and Cod_Sequencia <= " & mCodSequencia

        If M46V999.gfObtRegistro(lbdTbm_Parameter, lSelect, lrsTbm_Parameter, pMensagem) = False Then
            If mCodSequencia > 0 Then
                mCodSequencia = mCodSequencia - 1

                GoTo NovaSequencia
            End If
        End If

        If M46V999.Check(lrsTbm_Parameter("Total")) = 0 Then
            If mCodSequencia > 0 Then
                mCodSequencia = mCodSequencia - 1

                GoTo NovaSequencia
            End If
        End If


    End If

    lSelect = " Select Cod_Cobert, Cod_Tip_Chave, Chave_Ini, Chave_Ter, Valor from Tbm_Parameter WITH (NOLOCK) "
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
    If M46V999.gCodProduto >= 1080 And M46V111.gCod_plano < 1000 Then
        lSelect = lSelect & " and Tip_Atividade in (0, " & val(lCod_Compl_Clasf) & ")"
    End If
    lSelect = lSelect & " order by Cod_Cobert, Cod_Prioridade, Cod_Constr desc "

    If M46V999.gfObtRegistro(lbdTbm_Parameter, lSelect, lrsTbm_Parameter, pMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Parameter)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    '1100192 - Tarifa 1006 e 1007 Marco 2011
    lValCoef = 1
    Do While Not lrsTbm_Parameter.EOF
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 1 Then  'Valor de IS
            If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) <= lVal_IS And M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ter"))) >= lVal_IS Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 2 Then
            If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) <= lVal_Franquia And M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ter"))) >= lVal_Franquia Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 3 Then
            If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) <= lNum_Cep And M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ter"))) >= lNum_Cep Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
                If (pCod_Tip_Valor = e_Tip_Valor_ClasAceit) And _
                   (lCod_Cobert = M46V999.gfBuscarDePara(33, 1) Or lCod_Cobert = M46V999.gfBuscarDePara(55, 1)) And _
                   (val(ValorAjustavel) = 3) Then
                    pMensagem = "Região restrita para Alagamento."
                End If
            End If
        End If
        
        'JIRA-25945 - Opção "Sem Depreciação" no Condomínio
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 25 Then    'Cláusula de Nao Depreciacao
            If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) <= pobjstcPedLoc.ClausulaNaoDepreciacao _
            And M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ter"))) >= pobjstcPedLoc.ClausulaNaoDepreciacao Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 4 Then
            If M46V999.Check(lrsTbm_Parameter("Chave_Ini")) = lTip_Emissao And M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ter"))) = lTip_Emissao Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 5 Then
            If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) <= lPer_Coml And M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ter"))) >= lPer_Coml Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 6 Then
            If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) <= lCoef_Desc_Coml And M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ter"))) >= lCoef_Desc_Coml Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 7 Then
            '1100192 - Tarifa 1006 e 1007 Marco 2011
            If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) <= lVal_VR_Item And M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ter"))) >= lVal_VR_Item Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        '1001146 - Tarifa 1000 e 1001 Dezembro 2010
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 8 Then    'And lIndLMI = 1 Then
            If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) <= lCodRegiao And M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ter"))) >= lCodRegiao Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If

        '1200539 - Tarifa 1038 e 1039 Junho 2012
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 9 Then
            ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
        End If

        '1100192 - Tarifa 1006 e 1007 Marco 2011
        If pCod_Tip_Valor = e_Tip_Valor_CoeficienteTaxa And ValorAjustavel <> 0 Then
            lValCoef = lValCoef * ValorAjustavel
            ValorAjustavel = 0
        End If

        '1200806 - Tarifa 1044 e 1045 Marco 2012
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 10 Then  'Fatores de Ajuste Automáticos
            ValorAjustavel = ValorAjustavel & M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado) & "@" & M46V999.gfTratarNulo(lrsTbm_Parameter("Chave_Ini"), tpoString) & ";"
        End If

        '1201070 - Tarifa - 1050 e 1051 - Item 2 Ajuste de Taxa
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 11 Then
            If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) <= pobjstcPedLoc.gISBasica And M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ter"))) >= pobjstcPedLoc.gISBasica Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If

        ' 1300197 - Tarifa 1056 e 1057 - Inicio
        ' Restrições
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 12 Then
            If pCod_Tip_Valor = e_Tip_Restricao_Corretor Then
                pMensagem = "Cálculo restrito por políticas internas."
                mfAjustesPontuais = False
                GoTo SairNormal
            End If
        End If
        ' 1300197 - Tarifa 1056 e 1057 - Fim


        'CR3939 - Inicio
        ' Tipo de valor 13 - Olhar apenas para o campo Chave_Ini é igual ao tipo de emissão
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 13 Then
            If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) = pobjstcPedLoc.Tip_Emissao Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        'CR3939 - Fim

        'Luiz Alberto - 08/10/2015 - Tarifa 1104
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 23 Then
            ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
        End If
        
        'Pega o menor valor que encontrar
        If pCod_Tip_Valor = e_Tip_Valor_PremioMinimo Then
            If ValorComparacao < ValorAjustavel Then
                ValorComparacao = ValorAjustavel
            End If
        End If

        If pCod_Tip_Valor = e_Tip_Valor_DescontoAgrupamento Then
            If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = e_Cod_Tip_Chave_Qtde_Cob_Agrup Then
                If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) <= lQtdCoberturas And M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ter"))) >= lQtdCoberturas Then
                    ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
                End If
            End If
        End If

        'Define os juros do parcelamento por segmentado
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = e_Cod_Tip_Chave_Juros_Segmentados Then
            If val(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) = lCodSegmentado Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If

        'Define o parcelamento maximo para o corretor
        If pCod_Tip_Valor = e_Tip_Valor_ParcMaxCorretor Then
            ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
        End If

        'Validade de Cotação / Proposta por corretor
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = e_Cod_Tip_ValidaDiasCorr Then
            If pCod_Tip_Valor = e_Tip_Valor_ValidaDiasCorr Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If

        ' Chave por classe de ocupação
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = e_Cod_Tip_ClasseOcupacao Then
            If val(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) = lCod_Classe Then
                If ValorAjustavel = 0 Then
                    ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
                Else
                    ValorAjustavel = ValorAjustavel * M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
                End If
            End If
        End If

        lrsTbm_Parameter.MoveNext
    Loop

    '1100192 - Tarifa 1006 e 1007 Marco 2011
    If pCod_Tip_Valor = e_Tip_Valor_CoeficienteTaxa Then
        ValorAjustavel = lValCoef
    End If

    If pCod_Tip_Valor = e_Tip_Valor_PremioMinimo Then
        ValorAjustavel = ValorComparacao
    End If

    mfAjustesPontuais = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SairNormal:
    Call M46V999.gpFecha2(lbdTbm_Parameter)
    Set lbdTbm_Parameter = Nothing
    Set lrsTbm_Parameter = Nothing

    Exit Function
ControleErros:
    Set lbdTbm_Parameter = Nothing
    Set lrsTbm_Parameter = Nothing
    mfAjustesPontuais = False
    pMensagem = "mfAjustesPontuais ERRO - " & Err & " " & Err.Description & _
                vbCrLf & " SQL: " & lSelect
    Call M46V999.gpGraLog(0, pMensagem)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)

End Function



Public Sub CarregaCoberturaVida(ByRef pobjstcItem As stcA46V708B, ByVal pMensagem As String)
    Dim lbdTab_Tarifa As ADODB.Connection       'Arquivo: P00Tarifa.
    Dim lrsTbm_PlanVida As clsYasRecordSet      'Registro : TAB_Tarifa.
    Dim lColecaoCobertura As Collection
    Dim lobjstcCobert As stcA46V704B
    Dim lMensagem As String
    Dim lSelect As String
    Dim Aux As Double

    On Error GoTo ControleErro

    Set lColecaoCobertura = New Collection

    For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert
        If lobjstcCobert.gCod_Grupo_Cobert <> 4 Then
            lColecaoCobertura.Add lobjstcCobert, lobjstcCobert.Class_ID
        End If
    Next


    'pobjstcItem.Tip_Plano_Vida = 1
    'pobjstcItem.Qtd_Vidas = 6
    'pobjstcItem.Val_LMG_Vida = 15000
    'pobjstcItem.Cod_Plano_Cobertura = 3
    'pobjstcItem.Ind_Aux_Funeral = 1
    'pobjstcItem.Val_LMG_Funeral = 30
    'pobjstcItem.Val_Aux_Cesta = 40


    If pobjstcItem.Tip_Plano_Vida = 0 Then
        'Se nenhum plano escolhido retorna as coleções sem o vida
        Set pobjstcItem.ColecaoPedCobert = lColecaoCobertura
        Exit Sub
    End If

    If pobjstcItem.Qtd_Vidas = 0 Then
        Call M46V999.gpMenErro("FGE2294 - Quantidade de Vidas não pode estar zerado.")
        Exit Sub
    End If

    If pobjstcItem.Val_LMG_Vida = 0 Then
        Call M46V999.gpMenErro("FGE2294 - Quantidade de Vidas não pode estar zerado.")
        Exit Sub
    End If
    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTab_Tarifa, lMensagem) = False Then
        Call M46V999.gpMenErro(lMensagem)
        GoTo SAIDA
    End If

    lSelect = " select * from Tbm_Plano_Vida as V WITH (NOLOCK) "
    lSelect = lSelect & " WHERE v.Cod_Ramo = " & pobjstcItem.Cod_Ramo
    lSelect = lSelect & " and v.Cod_Modalidade = " & pobjstcItem.Tab_Clasf
    lSelect = lSelect & " and v.Cod_Tarif = " & pobjstcItem.Cod_Iden_Tarif
    lSelect = lSelect & " and v.Cod_Plano_Vida = " & pobjstcItem.Cod_Plano_Cobertura
    lSelect = lSelect & " and v.Cod_Sequencia  <= " & M46V999.gCodSequencia

    If M46V999.gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTbm_PlanVida, lMensagem) = False Then
        Call M46V999.gpGraLog(0, lMensagem)
        Call M46V999.gpMenErro(lMensagem)
        Call M46V999.gpFecha2(lbdTab_Tarifa)
        GoTo SAIDA
    End If
    Aux = 0
    Do While Not lrsTbm_PlanVida.EOF
        Set lobjstcCobert = New stcA46V704B
        With lobjstcCobert
            .Num_Item = pobjstcItem.Num_Item
            .nosso_numero = pobjstcItem.nosso_numero
            .Cod_Org = pobjstcItem.Cod_Org
            .Cod_Ramo = pobjstcItem.Cod_Ramo
            .Num_Ped = pobjstcItem.Num_Ped
            .Tab_Cobert = pobjstcItem.Cod_Plano
            .Cod_Cobert = lrsTbm_PlanVida("Cod_Cobert")
            .Num_Apol = pobjstcItem.Num_Apol
            .Num_Endo = pobjstcItem.Num_Endo
            .NUM_PI = pobjstcItem.NUM_PI
            Aux = pobjstcItem.Qtd_Vidas * pobjstcItem.Val_LMG_Vida
            If pobjstcItem.Ind_Inc_Conjuge <> 0 And lrsTbm_PlanVida("Tip_Plano_Vida") = 1 Then
                Aux = Aux + (pobjstcItem.Qtd_Vidas * pobjstcItem.Val_LMG_Vida * 0.5)
            End If
            .Val_IS = Aux
            Set .ObjPai = pobjstcItem
            Set .M46V999 = M46V999
            lColecaoCobertura.Add lobjstcCobert, lobjstcCobert.Class_ID

        End With

        lrsTbm_PlanVida.MoveNext
    Loop

    ' 4 - AUXILIO FUNERAL

    If pobjstcItem.Ind_Aux_Funeral = 1 Then

        If pobjstcItem.Val_LMG_Funeral = 0 Then
            Call M46V999.gpMenErro("FGE2294 - Valor de LMG de Funeral não pode estar zerado.")
            Exit Sub
        End If

        lSelect = " select * from Tbm_Plano_Vida as V WITH (NOLOCK) "
        lSelect = lSelect & " WHERE v.Cod_Ramo = " & pobjstcItem.Cod_Ramo
        lSelect = lSelect & " and v.Cod_Modalidade = " & pobjstcItem.Tab_Clasf
        lSelect = lSelect & " and v.Cod_Tarif = " & pobjstcItem.Cod_Iden_Tarif
        lSelect = lSelect & " and v.Cod_Plano_Vida = " & 4
        lSelect = lSelect & " and v.Cod_Sequencia  <= " & M46V999.gCodSequencia

        If M46V999.gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTbm_PlanVida, lMensagem) = False Then
            Call M46V999.gpGraLog(0, lMensagem)
            Call M46V999.gpMenErro(lMensagem)
            Call M46V999.gpFecha2(lbdTab_Tarifa)
            GoTo SAIDA
        End If
        Do While Not lrsTbm_PlanVida.EOF
            Set lobjstcCobert = New stcA46V704B
            With lobjstcCobert
                .Num_Item = pobjstcItem.Num_Item
                .nosso_numero = pobjstcItem.nosso_numero
                .Cod_Org = pobjstcItem.Cod_Org
                .Cod_Ramo = pobjstcItem.Cod_Ramo
                .Num_Ped = pobjstcItem.Num_Ped
                .Tab_Cobert = pobjstcItem.Cod_Plano
                .Cod_Cobert = lrsTbm_PlanVida("Cod_Cobert")
                .Num_Apol = pobjstcItem.Num_Apol
                .Num_Endo = pobjstcItem.Num_Endo
                .NUM_PI = pobjstcItem.NUM_PI
                .Val_IS = pobjstcItem.Qtd_Vidas * pobjstcItem.Val_LMG_Funeral
                Set .ObjPai = pobjstcItem
                Set .M46V999 = M46V999

                lColecaoCobertura.Add lobjstcCobert, lobjstcCobert.Class_ID
            End With
            lrsTbm_PlanVida.MoveNext
        Loop
    End If

    ' 5 - AUXILIO CESTA BASICA

    If pobjstcItem.Val_Aux_Cesta <> 0 Then
        lSelect = " select * from Tbm_Plano_Vida as V WITH (NOLOCK) "
        lSelect = lSelect & " WHERE v.Cod_Ramo = " & pobjstcItem.Cod_Ramo
        lSelect = lSelect & " and v.Cod_Modalidade = " & pobjstcItem.Tab_Clasf
        lSelect = lSelect & " and v.Cod_Tarif = " & pobjstcItem.Cod_Iden_Tarif
        lSelect = lSelect & " and v.Cod_Plano_Vida = " & 5
        lSelect = lSelect & " and v.Cod_Sequencia  <= " & M46V999.gCodSequencia

        If M46V999.gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTbm_PlanVida, lMensagem) = False Then
            Call M46V999.gpGraLog(0, lMensagem)
            Call M46V999.gpMenErro(lMensagem)

            GoTo SAIDA
        End If
        Do While Not lrsTbm_PlanVida.EOF
            Set lobjstcCobert = New stcA46V704B
            With lobjstcCobert
                .Num_Item = pobjstcItem.Num_Item
                .nosso_numero = pobjstcItem.nosso_numero
                .Cod_Org = pobjstcItem.Cod_Org
                .Cod_Ramo = pobjstcItem.Cod_Ramo
                .Num_Ped = pobjstcItem.Num_Ped
                .Tab_Cobert = pobjstcItem.Cod_Plano
                .Cod_Cobert = lrsTbm_PlanVida("Cod_Cobert")
                .Num_Apol = pobjstcItem.Num_Apol
                .Num_Endo = pobjstcItem.Num_Endo
                .NUM_PI = pobjstcItem.NUM_PI
                .Val_IS = pobjstcItem.Val_Aux_Cesta
                Set .ObjPai = pobjstcItem
                Set .M46V999 = M46V999

                lColecaoCobertura.Add lobjstcCobert, lobjstcCobert.Class_ID
            End With
            lrsTbm_PlanVida.MoveNext
        Loop
    End If

    Set pobjstcItem.ColecaoPedCobert = lColecaoCobertura

    Exit Sub

SAIDA:
    Call M46V999.gpFecha2(lbdTab_Tarifa)
    Set lColecaoCobertura = Nothing
    Set lobjstcCobert = Nothing
    
    Exit Sub
ControleErro:
    pMensagem = "CarregaCoberturaVida ERRO - " & Err & " " & Err.Description & _
                vbCrLf & " SQL: " & lSelect
    Call M46V999.gpGraLog(0, pMensagem)

    GoTo SAIDA

End Sub


Private Function mpFindCoefPI(ByRef pCoef_PI As Double, ByVal pNum_Meses_PI, ByRef pMensagem As String) As Boolean
    Dim lbdTbm_Coef_PI As ADODB.Connection              'Arquivo: P00Multi.
    Dim lMensagem As String                             'Mensagem.
    Dim lrsTbm_Coef_PI As clsYasRecordSet               'Registro : Tbm_Coef_PI.
    Dim lSelect As String

    pCoef_PI = 1
    mpFindCoefPI = False

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Coef_PI, lMensagem) = False Then
        Exit Function
    End If

    '1 - Procura pela Cobertura na Tabela de Coeficiente
    lSelect = " Select *  from Tbm_Coef_PI WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & lCod_Cobert
    lSelect = lSelect & " and Num_Meses = " & pNum_Meses_PI
    lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
    lSelect = lSelect & " and Cod_Sequencia <= " & lCod_Sequencia

    If M46V999.gfObtRegistro(lbdTbm_Coef_PI, lSelect, lrsTbm_Coef_PI, lMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Coef_PI)
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    mpFindCoefPI = True
    '2 - Se encontrar permite cadastro de Período Indenitário
    If Not lrsTbm_Coef_PI.EOF Then
        If lrsTbm_Coef_PI("Tip_Agravacao") = "S" Then
            pCoef_PI = M46V999.Check(lrsTbm_Coef_PI("TAXA_COEF_PI"))
        End If
    End If

    Call M46V999.gpFecha2(lbdTbm_Coef_PI)
    Set lrsTbm_Coef_PI = Nothing

End Function

Private Function mpFindCoefProtecao(ByRef pCoef_Protecao As Double, _
                                    ByRef pCoef_ProtecaoResseg As Double, _
                                    ByVal pColecaoPedFator As Collection, _
                                    ByRef pMensagem As String, tagDynaTrace As String, objPlano As stcPlano) As Boolean
                                    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mpFindCoefProtecao"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lbdTbm_Fator_Protecao As ADODB.Connection         'Arquivo: P00Tarifa.
    Dim lrsTbm_Fator_Protecao As clsYasRecordSet          'Registro : Tbm_Fator_Protecao.
    Dim lobjstcFator As stcA46V720B       'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lSelect As String

    Dim mCodSequencia As Integer
    pMensagem = ""
    pCoef_Protecao = 0
    pCoef_ProtecaoResseg = 0
    mpFindCoefProtecao = False

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Fator_Protecao, pMensagem) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    'Vai procurar para todos os fatores cadastrados se existem Agravação para a cobertura
    mCodSequencia = lCod_Sequencia
    For Each lobjstcFator In pColecaoPedFator
NovaSequencia:
        '1 - Procura Registro na Tabela de Fator de Ajuste por Protecao
        ' Exceção: Faz busca exata

        lSelect = " Select Fator_Ajuste as Fator, Tipo_Resseguro from Tbm_Fator_Protecao WITH (NOLOCK) "
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

        If M46V999.gfObtRegistro(lbdTbm_Fator_Protecao, lSelect, lrsTbm_Fator_Protecao, pMensagem) = False Then
            Call M46V999.gpFecha2(lbdTbm_Fator_Protecao)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            Exit Function
        End If

        '2 - Se não encontrar vai procurar por Padrão na mesma Sequencia
        If lrsTbm_Fator_Protecao.EOF Then

            GoTo BuscaCoefPadrao
        Else
            If objPlano.EhPlanoSegmentado Then
                If pCoef_Protecao < M46V999.Check(lrsTbm_Fator_Protecao("Fator")) Then
                    pCoef_Protecao = M46V999.Check(lrsTbm_Fator_Protecao("Fator"))
                    If M46V999.Check(lrsTbm_Fator_Protecao("Tipo_Resseguro")) = "S" Then
                        pCoef_ProtecaoResseg = M46V999.Check(lrsTbm_Fator_Protecao("Fator"))
                    End If
                End If
            Else
                pCoef_Protecao = pCoef_Protecao + M46V999.Check(lrsTbm_Fator_Protecao("Fator"))
                                
                '8000045507, Desconto - Cobertura de Roubo
                'Para classificação diferente de casa habital, não se aplica o coeficiente.
                If lCod_Ramo = 113 And (Trim(lCod_Clasf) <> "910" And Trim(lCod_Clasf) <> "914") Then
                    pCoef_Protecao = 0
                End If
                
                If lCod_Ramo = 113 And pCoef_Protecao > 40 Then
                    pCoef_Protecao = 40
                End If
                                
                If M46V999.Check(lrsTbm_Fator_Protecao("Tipo_Resseguro")) = "S" Then
                    pCoef_ProtecaoResseg = pCoef_ProtecaoResseg + M46V999.Check(lrsTbm_Fator_Protecao("Fator"))
                End If
            End If
            GoTo ProximaProtecao
        End If

BuscaCoefPadrao:
        '3 - Procura Registro na Tabela de Fator de Ajuste por Protecao
        ' Padrão : Busca pela Fator e Cobertura

        lSelect = " Select Fator_Ajuste as Fator, Tipo_Resseguro from Tbm_Fator_Protecao WITH (NOLOCK) "
        lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & lCod_Modalidade
        lSelect = lSelect & " and Cod_Tip_Reg = 'D' "                  'Default
        lSelect = lSelect & " and Cod_Cobert = " & lCod_Cobert
        lSelect = lSelect & " and Cod_Fator = " & lobjstcFator.Cod_Fator
        lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
        lSelect = lSelect & " and Cod_Sequencia <= " & mCodSequencia

        If M46V999.gfObtRegistro(lbdTbm_Fator_Protecao, lSelect, lrsTbm_Fator_Protecao, pMensagem) = False Then
            Call M46V999.gpFecha2(lbdTbm_Fator_Protecao)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            Exit Function
        End If

        '4 - Se não encontrar, deverá buscar novamente a exceção com
        'sequencia anterior
        If lrsTbm_Fator_Protecao.EOF Then
            If mCodSequencia > 0 Then
                mCodSequencia = mCodSequencia - 1

                GoTo NovaSequencia
            End If
        Else
            If objPlano.EhPlanoSegmentado And 0 = 1 Then ' Desvio desabilitado por enquanto (Daniel Ribeiro 30/05/2016)
                If pCoef_Protecao < M46V999.Check(lrsTbm_Fator_Protecao("Fator")) Then
                    pCoef_Protecao = M46V999.Check(lrsTbm_Fator_Protecao("Fator"))

                    If M46V999.Check(lrsTbm_Fator_Protecao("Tipo_Resseguro")) = "S" Then
                        pCoef_ProtecaoResseg = M46V999.Check(lrsTbm_Fator_Protecao("Fator"))
                    End If
                End If
            Else
                pCoef_Protecao = pCoef_Protecao + M46V999.Check(lrsTbm_Fator_Protecao("Fator"))
                
                If lCod_Ramo = 113 And (Trim(lCod_Clasf) <> "910" And Trim(lCod_Clasf) <> "914") Then
                    pCoef_Protecao = 0
                End If
                
                If lCod_Ramo = 113 And pCoef_Protecao > 40 Then
                    pCoef_Protecao = 40
                End If
                
                If M46V999.Check(lrsTbm_Fator_Protecao("Tipo_Resseguro")) = "S" Then
                    pCoef_ProtecaoResseg = pCoef_ProtecaoResseg + M46V999.Check(lrsTbm_Fator_Protecao("Fator"))
                End If
            End If
        End If


ProximaProtecao:
    Next lobjstcFator
    Set lobjstcFator = Nothing
    Set lrsTbm_Fator_Protecao = Nothing
    Call M46V999.gpFecha2(lbdTbm_Fator_Protecao)
    mpFindCoefProtecao = True

    '5 - Vai procurar na tabela de Ajustes pontuais
    If mfAjustesPontuais(e_Tip_Valor_CoefProt, TpoValor, pCoef_Protecao, pMensagem, Nothing, 0, tagDynaTrace) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
End Function
Private Function mpFindCoefRegiao(ByRef pCoef_Regiao As Double, ByRef pMensagem As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mpFindCoefRegiao"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lbdTbm_Fator_Cep As ADODB.Connection      'Arquivo: P00Tarifa.
    Dim lrsTbm_Fator_Cep As clsYasRecordSet       'Registro : Tbm_Fator_Cep.
    Dim mCodSequencia As Integer
    Dim lSelect As String

    pMensagem = ""
    pCoef_Regiao = 0
    mpFindCoefRegiao = False

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Fator_Cep, pMensagem) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    mCodSequencia = lCod_Sequencia

NovaSequencia:

    '1 - Procura Registro na Tabela de Fator de Ajuste por CEP

    lSelect = " Select Fator_Ajuste as Fator, Num_Cep_Inicial, Num_Cep_Final"
    lSelect = lSelect & " from Tbm_Fator_Cep WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & lCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & lCod_Modalidade
    lSelect = lSelect & " and Cod_Cobert = " & lCod_Cobert
    lSelect = lSelect & " and Cod_Tarif = " & lCod_Tarifa
    lSelect = lSelect & " and Cod_Sequencia <= " & mCodSequencia
    lSelect = lSelect & " order by Num_Cep_Inicial "

    If M46V999.gfObtRegistro(lbdTbm_Fator_Cep, lSelect, lrsTbm_Fator_Cep, pMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Fator_Cep)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    '2 - Se não encontrar a sequencia vai procurar sequencia anterior
    If lrsTbm_Fator_Cep.EOF Then
        If mCodSequencia > 0 Then
            mCodSequencia = mCodSequencia - 1

            GoTo NovaSequencia
        End If
    Else
        '3 - Se encontrar Coeficiente para essa cobertura vai verificar se está dentro da faixa
        Do While Not lrsTbm_Fator_Cep.EOF
            If lNum_Cep >= M46V999.Check(lrsTbm_Fator_Cep("Num_Cep_Inicial")) And _
               lNum_Cep <= M46V999.Check(lrsTbm_Fator_Cep("Num_Cep_Final")) Then
                pCoef_Regiao = M46V999.Check(lrsTbm_Fator_Cep("Fator"))
                Exit Do
            End If
            lrsTbm_Fator_Cep.MoveNext
        Loop
    End If

    '4 - Vai procurar na tabela de Ajustes pontuais
    If mfAjustesPontuais(e_Tip_Valor_CoefRegiao, TpoValor, pCoef_Regiao, pMensagem, Nothing, 0, tagDynaTrace) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If


    mpFindCoefRegiao = True
    Call M46V999.gpFecha2(lbdTbm_Fator_Cep)
    Set lrsTbm_Fator_Cep = Nothing
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

End Function

Public Function gfVerParametro(ByVal pTipoValor As Integer, _
                               ByRef pobjstcPed As stcA46V702B, ByRef pobjstcLoc As stcA46V708B, _
                               ByRef pobjStcCobert As stcA46V704B, _
                               ByRef pRetorno As String, _
                               ByRef pMensagem As String) As Boolean
'Função #071: verifica parametrizações do produto.

'Parâmetros de saída.....pRetorno...........retorno.
'                        pMensagem..........mensagem.

    Dim bdParametro As ADODB.Connection         'P0044800: base de dados.
    Dim lSelect As String                                       'Comando SELECT.
    Dim rsParm As clsYasRecordSet                       'P0044800: tabela.
    Dim lParametro As Variant                           'Array com os valores
    Dim lBlnContinuar As Boolean                        'Controla a validação
    Dim lCodProtocolo As Long                           'Código do protocolo
    Dim lCodSequencia As Long                           'Sequencia.
    Dim lProxSequencia As Long


    gfVerParametro = False

    '1. Posiciona controles.
    pRetorno = " "
    '2. Carrego os parametros
    If Not mfCarregarParametros(lParametro, pMensagem, pobjstcPed, pobjstcLoc, pobjStcCobert) Then Exit Function
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
    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, bdParametro, pMensagem) = False Then
        Exit Function
    End If

    lCodSequencia = M46V999.gCodSequencia
    lCodProtocolo = M46V999.gCodProtocolo
    If lCodSequencia > 0 Then
Voltar:
        lSelect = "SELECT top 1 * from TAB_PARM_DIVERSOS WITH (NOLOCK) where " & _
                  "Cod_Produto = " & M46V999.gCodProduto & " " & _
                  "AND Cod_Protocolo = " & lCodProtocolo & " " & _
                  "AND Cod_Sequencia <= " & lCodSequencia
        If M46V999.gfObtRegistro(bdParametro, lSelect, rsParm, pMensagem) = False Then
            Call M46V999.gpFecha2(bdParametro)
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
                    lCodSequencia = M46V999.gCodSequencia
                    GoTo Voltar
                End If
            End If
        End If
    End If

    lSelect = "SELECT * " & _
              "FROM TAB_PARM_DIVERSOS inner join TAB_TIP_CHAVE WITH (NOLOCK) " & _
              "On TAB_PARM_DIVERSOS.Tipo_Chave = TAB_TIP_CHAVE.Tip_Chave " & _
              "WHERE Cod_Produto = " & M46V999.gCodProduto & " " & _
              "AND Cod_Protocolo = " & lCodProtocolo & " " & _
              "AND Cod_Sequencia <= " & lCodSequencia & _
            " AND Tipo_Valor = " & pTipoValor & " "

    lSelect = lSelect & " and cod_ramo = " & M46V999.gCod_Ramo
    If pTipoValor = 7 Then
        lSelect = lSelect & " ORDER BY Nivel_Prioridade asc"
    Else
        lSelect = lSelect & " ORDER BY Nivel_Prioridade desc"
    End If

    If M46V999.gfObtRegistro(bdParametro, lSelect, rsParm, pMensagem) = False Then
        Call M46V999.gpFecha2(bdParametro)
        Exit Function
    End If
    '4. Verifica registros.
    Do While Not rsParm.EOF
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
    Loop

Fim:
    Call M46V999.gpFecha2(bdParametro)
    Set rsParm = Nothing
    gfVerParametro = True
End Function

Private Function mfCarregarParametros(ByRef pParametros As Variant, pMensagem As String, _
                                      pobjstcPed As stcA46V702B, pobjstcLoc As stcA46V708B, _
                                      ByRef pobjStcCobert As stcA46V704B) As Boolean
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

    Dim lTarifaBasica As Long     'Tarifa Básica.
    Dim lCodUnidade As Long     'Código da Unidade.
    Dim lCodCorretora As Long     'Código do corretor.
    Dim lNumeroParcela As Integer  'Número de parcelas.
    Dim lPagAntecipado As Integer  'Indicador de pagamento antecipado.
    Dim lCodCobertura As Integer  'Código da cobertura
    Dim lSprinklers As Integer  'Contém sprinklers.
    Dim lModalidade As Long     'Modalidade
    Dim lValLMG As Double   'Valor LMG
    Dim lRubrica As String   'Classificação
    Dim lCodSegurado As String
    Dim lCodConstr As e_CodContr
    Dim lobjstcFator As stcA46V720B          'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lobjstcCobertura As stcA46V704B
    Dim lSinistro As String    '1100705 - Tarifa 1016 e 1017 Agosto 2011
    Dim lCepRisco As Long   '1201120 - Tarifa 1052 e 1053 - Item 2

    '1001224 - Condições Diferenciada - Seguratec
    Dim lDesconto As Double
    Dim lobjstcCorretor As stcA46V705B


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
    lCodCobertura = pobjStcCobert.Cod_Cobert

    '6 - Contém Sprinklers?
    lSprinklers = 2
    For Each lobjstcFator In pobjstcLoc.ColecaoPedFator
        If M46V999.gCod_Ramo = e_CodRamo_Empresarial Or M46V999.gCod_Ramo = e_CodRamo_Roubos Or M46V999.gCod_Ramo = e_CodRamo_Riscos_diversos Then
            If lobjstcFator.Cod_Fator = 4 Then
                lSprinklers = 1
                Exit For
            End If
        ElseIf M46V999.gCod_Ramo = e_CodRamo_Condominio Then
            If lobjstcFator.Cod_Fator = 8 Then
                lSprinklers = 1
                Exit For
            End If
        End If
    Next lobjstcFator

    '7  -   Modalidade
    lModalidade = pobjstcLoc.Cod_Plano

    '8  -   Rubrica
    lRubrica = val(Replace(pobjstcLoc.Cod_Clasf, "A", "999") & pobjstcLoc.Cod_Compl_Clasf)

    '9  -   LMG
    lValLMG = pobjStcCobert.Val_IS

    '10 - Codigo de construção
    lCodConstr = pobjstcLoc.Cod_Constr

    '11 - Código do segurado - (1001224 - Condições Diferenciada - Seguratec)
    lCodSegurado = val(pobjstcPed.Cod_Segurado)

    '12 - Desconto - (1001224 - Condições Diferenciada - Seguratec)
    lDesconto = pobjStcCobert.Per_Desc_Coml

    '13 -   Número da Parcela
    lNumeroParcela = pobjstcPed.Qtd_Parc

    '18 - cod segurado
    lCodSegurado = pobjstcPed.Cod_Segurado

    '22 -   Pagamento antecipado?
    lPagAntecipado = pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec

    '23 - Sinistro?
    lSinistro = pobjstcLoc.Tip_Emissao & pobjstcLoc.Ind_Sinistro    '1100705 - Tarifa 1016 e 1017 Agosto 2011

    '24 - Cep de risco
    lCepRisco = pobjstcLoc.Num_Cep_Risco  '1201120 - Tarifa 1052 e 1053 - Item 2


    pParametros = Array(lTarifaBasica, lCodUnidade, lCodCorretora, _
                        lNumeroParcela, lPagAntecipado, lCodCobertura, _
                        lSprinklers, lModalidade, lRubrica, lValLMG, _
                        lCodConstr, lCodSegurado, lDesconto, _
                        lNumeroParcela, 0, 0, 0, 0, lCodSegurado, 0, 0, 0, _
                        lPagAntecipado, lSinistro, lCepRisco)    '1100705 - Tarifa 1016 e 1017 Agosto 2011
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
                                   ByRef pMensagem As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfIncluirClausula"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lobjClausula As clsA46V703A
    Dim lobjstcClau As stcA46V703B

    Set lobjClausula = New clsA46V703A
    Set lobjClausula.M46V999 = M46V999

    Set lobjstcClau = New stcA46V703B
    Set lobjstcClau.M46V999 = M46V999

    With lobjstcClau
        .nosso_numero = pobjstcItem.nosso_numero
        .Cod_Clau = pClausula
        .Cod_Ramo = pobjstcItem.Cod_Ramo
        .Num_Ped = pobjstcItem.Num_Ped
        .Num_Item = pobjstcItem.Num_Item
        .Descricao = " "
        .NUM_PI = pobjstcItem.NUM_PI
    End With
    If Not lobjClausula.mfIncluir(lobjstcClau, pPropTabela, pNomeTabImport, tagDynaTrace) Then
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
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
FechaObj:
    Set lobjClausula = Nothing
    Set lobjstcClau = Nothing
    Return
End Function
Public Sub gpVerNormasParametrizada(pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B, pMensagem As String, pTipoConsistencia As Integer)

'Procedure #: verifica normas de aceitação
    Dim lMensagem As String                             'Mensagem.
    Dim lNormaRes As Boolean                    'Resposta da norma.
    Dim lPrefNorma As String                    'Prefixo na mensagem.
    Dim lResultado As String                    'Resultado.
    Dim lRstNorma As clsYasRecordSet    'Registros (normas).
    Dim lSelect As String                       'Select.
    Dim lNivel() As String                      'Array de nível.
    Dim lCodSequencia As Long           'Código da sequencia.
    Dim lCodProtocolo As Long           'Código da sequencia.
    Dim BDP00Tab As ADODB.Connection    'Conexão com a tabela.
    Dim lListaRegra As String
    Dim lBDP00Multi As ADODB.Connection
    Dim lrst As clsYasRecordSet
    Dim lInsert As String

    If Not UCase(M46V999.gSerDIVERSOS) Like "*MIX*" Or M46V999.gTipDIVERSOS = "2" Then GoTo SAIDA
    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, BDP00Tab, pMensagem) = False Then
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
        lSelect = lSelect & " Tbm_Parm_Normas WITH (NOLOCK) "
        lSelect = lSelect & " where "
        lSelect = lSelect & " Tbm_Parm_Normas.cod_ramo = " & lCod_Ramo
        lSelect = lSelect & " and tip_parm = " & pTipoConsistencia
        lSelect = lSelect & " and Tbm_Parm_Normas.cod_produto = " & lCod_Tarifa
        lSelect = lSelect & " and Tbm_Parm_Normas.cod_protocolo in (0, " & lCodProtocolo & ")"
        lSelect = lSelect & " and Tbm_Parm_Normas.cod_sequencia = " & lCodSequencia

        If M46V999.gfObtRegistro(BDP00Tab, lSelect, lRstNorma, lMensagem) = False Then
            Call M46V999.gpFecha2(BDP00Tab)
            Call M46V999.gpMenErro(lMensagem)
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
                    Call M46V999.gpFecha2(BDP00Tab)
                    Call M46V999.gpMenErro("Nenhuma norma foi localizada!")
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
            Call M46V999.gpMenErro(lMensagem)
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
            If M46V999.VerificarNormas = True Then
                Call M46V999.gpMenErro(lPrefNorma & lMensagem)
            End If
            lListaRegra = lListaRegra & lRstNorma("Cod_Grupo_Parm") & "." & lRstNorma("Cod_Clasf_Parm") & "." & lRstNorma("Cod_Parm") & "/"

            If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda And M46V999.gTipoAcesso = CALCULO Then
                If pobjstcPed.objstcCtrlEmis.Tip_Cotac = "2" And pobjstcPed.objstcCtrlEmis.Num_Cotac <> "" Then
                    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
                        Call M46V999.gpGraLog(2, lMensagem)
                        GoTo SAIDA
                    End If
                    lInsert = " Insert into siscota.dbo.Tab_Cotacao_CheckList_Resseguro (CCotacao, Cod_ItemCheckList, Cod_Ramo, Num_Item ) values " & _
                              "('" & pobjstcPed.objstcCtrlEmis.Num_Cotac & "'," & lRstNorma("Cod_Grupo_Parm") & lRstNorma("Cod_Clasf_Parm") & Format(lRstNorma("Cod_Parm"), "00") & "," & _
                              pobjstcPed.Cod_Ramo & "," & pobjstcPedLoc.Num_Item & ")"
                    If M46V999.gfExeSQL(lBDP00Multi, lInsert, lMensagem) <> 0 Then
                        Call M46V999.gpGraLog(2, lMensagem)
                        GoTo SAIDA
                    End If
                End If
            End If
        End If
        lRstNorma.MoveNext
    Loop

    'Thais - resseguro
    Call M46V999.gpFecha2(BDP00Tab)
    If M46V999.gfPreenchido(lListaRegra) Then
        If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda And M46V999.gTipoAcesso = DIGITACAO Then
            If M46V999.gfAbrBasDados("GED", M46V999.gTipGed, M46V999.gEndGed, lBDP00Multi, lMensagem) = False Then
                Call M46V999.gpGraLog(2, lMensagem)
                GoTo SAIDA
            End If
            lSelect = " Select * from ged.dbo.tab_prop_mensagem WITH (NOLOCK) where num_pi = " & pobjstcPed.NUM_PI & " and cod_mensagem = 8000 and item_pi = " & pobjstcPedLoc.Num_Item
            If Not M46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) Then
                Call M46V999.gpGraLog(2, lMensagem)
                GoTo SAIDA
            End If
            If lrst.EOF Then
                lInsert = " INSERT INTO [GED].[dbo].[tab_Prop_Mensagem]([num_pi], [item_pi], [cod_mensagem], [cod_tip_mensagem], [motivo], [dat_digit], [dat_liber], [hor_liber]) "
                lInsert = lInsert & " VALUES(" & pobjstcPed.NUM_PI & "," & pobjstcPedLoc.Num_Item & ",8000,99,'" & lListaRegra & "'," & Format(Date, "yyyymmdd") & ",0,0)"
                If M46V999.gfExeSQL(lBDP00Multi, lInsert, lMensagem) <> 0 Then
                    Call M46V999.gpGraLog(2, lMensagem)
                    GoTo SAIDA
                End If
            End If
        End If
    Else
        If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda And M46V999.gTipoAcesso = DIGITACAO Then
            If M46V999.gfAbrBasDados("GED", M46V999.gTipGed, M46V999.gEndGed, lBDP00Multi, lMensagem) = False Then
                Call M46V999.gpGraLog(2, lMensagem)
                GoTo SAIDA
            End If
            lSelect = " Delete from ged.dbo.tab_prop_mensagem where num_pi = " & pobjstcPed.NUM_PI & " and cod_mensagem = 8000 and dat_liber = 0  and item_PI = " & pobjstcPedLoc.Num_Item
            If Not M46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) Then
                Call M46V999.gpGraLog(2, lMensagem)
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
    Dim lBDP00Multi As ADODB.Connection
    Dim lMensagem As String
    Dim lDelete As String

    If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda And M46V999.gTipoAcesso = CALCULO Then
        If pobjstcPed.objstcCtrlEmis.Tip_Cotac = "2" And pobjstcPed.objstcCtrlEmis.Num_Cotac <> "" Then
            If Not UCase(M46V999.gSerDIVERSOS) Like "*MIX*" Then GoTo SAIDA
            If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
                Call M46V999.gpGraLog(2, lMensagem)
                GoTo SAIDA
            End If
            lDelete = " Delete from siscota.dbo.Tab_Cotacao_M46V999.CheckList_Resseguro where ccotacao = '" & pobjstcPed.objstcCtrlEmis.Num_Cotac & "'"
            If M46V999.gfExeSQL(lBDP00Multi, lDelete, lMensagem) <> 0 Then
                Call M46V999.gpGraLog(2, lMensagem)
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
Private Function mfParmCalculo(pBDConexao As ADODB.Connection, pRstParm As clsYasRecordSet, _
                               pResultado As String, pCodProtocolo As Long, _
                               pCodSequencia As Long, pTip_Parm As Integer, _
                               pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B, pMensagem As String) As Boolean


    Dim lAux1Num As Double                              'Auxiliar 1 numérico.
    Dim lAux2Num As Double                              'Auxiliar 2 numérico.
    Dim lCondicao As Byte                               '1 = and; 2= or.
    Dim lCondNivel As Byte                              'Condição do Nível.
    Dim lCount As Integer                               'Utilizado no for...loop.
    Dim lNivelRes As Boolean                    'Resposta no nível.
    Dim lNivelAnt As Integer                    'Nível anterior.
    Dim lNormaRes As Boolean                    'Resposta da norma.
    Dim lPriVez As Boolean                              'Indicador de primeira vez.
    Dim lRegraRes As Boolean                    'Resposta da regra.
    Dim lRstRegra As clsYasRecordSet    'Registros (regras).
    Dim lSaida As Variant                               'Saída da função.
    Dim lSelect As String                               'Select.
    Dim lValorNum As Double                             'Valor de retorno numérico.
    Dim lValorRegraNum As Double                'Valor da regra (numérico).
    Dim lValorRegraStr As String        'Valor da regra (string).
    Dim lValorStr As String                     'Valor de retorno string.
    Dim lNivel() As String                      'Array de nível.
    Dim lUltCodProtocolo As Long        'Ultimo codigo do protocolo
    Dim lretorno As String                      'Retorno
    Dim lObjstcCorr As stcA46V705B
    Dim lobjstcCobert As stcA46V704B
    Dim lobjstcPedLoc As stcA46V708B
    Dim lValorIS As Double
    Dim lValorSoma As Double
    Dim lISCob1 As Double
    Dim lISCob2 As Double
    Dim lAuxiliar As String
    Dim lVigIniAno As Long
    Dim lVigIniMes As Long
    Dim lVigIniDia As Long
    Dim lVigIniAnoMes As Long
    Dim lVigTerAno As Long
    Dim lVigTerMes As Long
    Dim lVigTerDia As Long
    Dim lVigTerAnoMes As Long


    pResultado = ""
    mfParmCalculo = False



    'início C
    lNivelAnt = 1
    'Fim C
    'Início D
    lSelect = " Select Tbm_Parm_Normas.cod_protocolo , Tbm_Parm_Normas.cod_regra, Desc_Parm, Nivel, Cond_nivel,"
    lSelect = lSelect & " Resultado, Tip_Regra,Item_Parm, Val_Regra, Operador, "
    lSelect = lSelect & " Condicao , Resultado FROM Tbm_Regra_Norma WITH (NOLOCK) INNER JOIN Tbm_Parm_Normas WITH (NOLOCK) "
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

    If M46V999.gfObtRegistro(pBDConexao, lSelect, lRstRegra, pMensagem) = False Then
        Call M46V999.gpFecha2(pBDConexao)
        Call M46V999.gpMenErro(pMensagem)
        GoTo SAIDA
    End If

    If lRstRegra.EOF = True Then
        Call M46V999.gpFecha2(pBDConexao)
        Call M46V999.gpMenErro("Nenhuma regra foi localizada!")
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
            lValorNum = val(pobjstcPedLoc.Cod_Clasf)
        Case 3   'Código do complemento da classificação
            lValorNum = val(pobjstcPedLoc.Cod_Compl_Clasf)
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
        Case 11    '                11  (Somatoria) IS da cobertura
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
        Case 13    'IS Cob 1 - comparação %
            lValorNum = 0
            lISCob1 = 0
            For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
                If lobjstcCobert.Cod_Cobert = lValorRegraNum Then
                    lValorNum = lRstRegra("Val_Regra")
                    lISCob1 = lobjstcCobert.Val_IS
                    Exit For
                End If
            Next lobjstcCobert
        Case 14    'IS Cob 2 - comparação %
            lValorNum = 0
            lISCob2 = 0
            For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
                If lobjstcCobert.Cod_Cobert = lValorRegraNum Then
                    lValorNum = lRstRegra("Val_Regra")
                    lISCob2 = lobjstcCobert.Val_IS
                    Exit For
                End If
            Next lobjstcCobert
        Case 15    'Is cob 1 / Is cob 2 %
            lValorNum = 0
            If lISCob2 <> 0 Then
                lValorNum = (lISCob1 / lISCob2) * 100
            End If
        Case 16    'Meses de vigência
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

            Do Until val(Mid(lVigIniAnoMes, 1, 4)) = val(Mid(lVigTerAnoMes, 1, 4))
                lVigTerAnoMes = lVigTerAnoMes - 100
                lVigTerAnoMes = lVigTerAnoMes + 12
            Loop
            lValorNum = lVigTerAnoMes - lVigIniAnoMes

        Case Else
            M46V999.gpMenErro ("FDE0023 - Tipo de regra desconhecido - Verificação das normas.")
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
        lretorno = lRstRegra("Resultado")
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
                pResultado = lretorno
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
                    pResultado = lretorno
                    lNormaRes = True
                Else
                    pResultado = ""
                    lNormaRes = False
                End If
            Case 2      'OR
                If lNormaRes = True Or lNivel(2, lCount) = True Then
                    pResultado = lretorno
                    lNormaRes = True
                Else
                    pResultado = ""
                    lNormaRes = False
                End If
            End Select
            lCondNivel = lNivel(3, lCount)
        End If
    Next lCount


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
Public Function mfAdicionaTextoEspecificacao(ByVal pCodEspecificacao As Long, _
                                            ByRef pObjTextoEspec As clsA46V726A, _
                                            ByRef pStcPedido As stcA46V702B, _
                                            ByVal pPropTabela As String, _
                                            ByRef pMensagem As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfAdicionaTextoEspecificacao"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lobjstcTextoEspec As stcA46V726B
    Dim lbdTab_Ped_Dizer_Pad As ADODB.Connection
    Dim lSelect As String
    Dim lrsTab_Ped_Dizer_Pad As clsYasRecordSet

    mfAdicionaTextoEspecificacao = False

    Set lobjstcTextoEspec = New stcA46V726B

    With lobjstcTextoEspec
        .nosso_numero = pStcPedido.nosso_numero
        .Cod_Org = pStcPedido.Cod_Org
        .Cod_Ramo = pStcPedido.Cod_Ramo
        .Num_Ped = pStcPedido.Num_Ped
        .Num_Apol = pStcPedido.Num_Apol
        .Num_Endo = pStcPedido.Num_Endo
        .Cod_Espec = pCodEspecificacao   'cadastrado no main-frame
        .Num_pag = pStcPedido.ColecaoTextoEspec.Count + 1
        .Qtd_vias = 3
        '        'abre base de dados
        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTab_Ped_Dizer_Pad, pMensagem) = False Then
            GoTo SAIDA
        End If
        'consulta mensagem
        lSelect = "SELECT * FROM TAB_PED_DIZER_PAD WITH (NOLOCK) WHERE COD_DIZER_PAD = " & .Cod_Espec
        lSelect = lSelect & " order by COD_DIZER_PAD , num_seq "
        If M46V999.gfObtRegistro(lbdTab_Ped_Dizer_Pad, lSelect, lrsTab_Ped_Dizer_Pad, pMensagem) = False Then
            GoTo SAIDA
        End If
        If lrsTab_Ped_Dizer_Pad.RecordCount = 0 Then pMensagem = "Mensagem Padrão " & .Cod_Espec & " não encontrada!": GoTo SAIDA
        .Espec_lin1 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString)
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin2 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin3 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin4 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin5 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin6 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin7 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin8 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin9 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin10 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin11 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin12 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin13 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin14 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin15 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin16 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin17 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin18 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin19 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
        lrsTab_Ped_Dizer_Pad.MoveNext: If Not lrsTab_Ped_Dizer_Pad.EOF Then .Espec_lin20 = M46V999.gfTratarNulo(lrsTab_Ped_Dizer_Pad("DSC_DIZER_ENDO"), tpoString) Else GoTo FimDizer
FimDizer:
        .NUM_PI = pStcPedido.NUM_PI    '  gNum_Pi
    End With
    If Not pStcPedido.FindTextoEspec(lobjstcTextoEspec) Then
        If Not pObjTextoEspec.mfIncluir(lobjstcTextoEspec, pPropTabela, "", tagDynaTrace) Then
            pMensagem = pObjTextoEspec.UltimoErro
            GoTo SAIDA
        End If
        pStcPedido.ColecaoTextoEspec.Add lobjstcTextoEspec, lobjstcTextoEspec.Class_ID
    End If

    mfAdicionaTextoEspecificacao = True

SAIDA:
    GoSub FechaObj
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
FechaObj:
    Set lobjstcTextoEspec = Nothing
    Call M46V999.gpFecha2(lbdTab_Ped_Dizer_Pad)
    Set lrsTab_Ped_Dizer_Pad = Nothing
    Return
End Function

Public Property Get M46V999() As clsM46V999
    Set M46V999 = M46V111.M46V999
End Property

Public Property Set M46V111(pClsM46V111 As clsM46V111)
    Set mClsM46V111 = pClsM46V111
End Property

Public Property Get M46V111() As clsM46V111
    If mClsM46V111 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V111 não atribuído para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V111 já
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46V111 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46V111 = mClsM46V111
End Property

Private Sub Class_Initialize()
    Set mClsM46V111 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V111 = Nothing
End Sub

Sub RetornaVlrRgrBatente(ByRef pStcItem As stcA46V708B, _
                         ByRef lobjstcCobert As Collection, _
                         lrsTab_Ped_Cobert As ADODB.Recordset, _
                         ByVal Vlr_Pr_Ant As Double, _
                         ByVal Vlr_Pr_Atu As Double, _
                         ByVal pPropTabela As String, _
                         ByVal pNomeTabImport As String)
    Dim lbdTab_Ped_Cobert As ADODB.Connection
    Dim lSelect As String
    '    Dim lrsTab_Ped_Cobert   As clsYasRecordSet
    Dim pMensagem As String
    Dim Vlr_Pr_Ant_ As Double
    Dim Vlr_Pr_Atu_ As Double
    Dim VlrLimBatInferior As Double
    Dim VlrLimBatSuperior As Double
    Dim PerBatente As Double
    Dim DifLimBatente As Double
    Dim lobjstcCobertAtu As stcA46V704B
    Dim PerAuxBatente As Double
    Dim PerDifBatVlrAtual As Double
    Dim lobjCobert As clsA46V704A
    Dim StrPerBat As String
    Dim VlrPrCoberAnt As Double
    Dim BlnVlrPrNegativo    As Boolean
    Dim DblVerificaPrNeg    As Double
    Dim StrFormula          As String
    Dim intPosicao          As Integer
    Dim lCobertDePara       As Integer

    Set lobjCobert = New clsA46V704A
    Set lobjCobert.M46V999 = M46V999

    BlnVlrPrNegativo = False
    DblVerificaPrNeg = 0
    StrFormula = ""
    intPosicao = 0
    VlrLimBatInferior = 0
    VlrLimBatSuperior = 0
    PerBatente = 0
    lCobertDePara = 0
    
        DifLimBatente = 0.1

    PerAuxBatente = 0

    VlrLimBatInferior = Vlr_Pr_Ant - (Vlr_Pr_Ant * DifLimBatente)
    VlrLimBatSuperior = Vlr_Pr_Ant + (Vlr_Pr_Ant * DifLimBatente)
    PerBatente = Vlr_Pr_Atu / Vlr_Pr_Ant - 1

    'VERIFICA SE O PERCENTUAL ESTA FORA DO BATENTE
    If (PerBatente > DifLimBatente) Or (PerBatente < DifLimBatente) Then
        If Vlr_Pr_Atu > VlrLimBatSuperior Then
            PerDifBatVlrAtual = VlrLimBatSuperior - Vlr_Pr_Atu
        ElseIf Vlr_Pr_Atu < VlrLimBatInferior Then
            PerDifBatVlrAtual = VlrLimBatInferior - Vlr_Pr_Atu
        Else
            PerDifBatVlrAtual = Vlr_Pr_Atu
        End If
        If lrsTab_Ped_Cobert.RecordCount > 0 Then

CalculaPremioNegativo:
            lrsTab_Ped_Cobert.MoveFirst

            Do Until lrsTab_Ped_Cobert.EOF
                PerAuxBatente = lrsTab_Ped_Cobert!VAL_PR / Vlr_Pr_Ant
                If lobjstcCobert.Count > 0 Then
                    For Each lobjstcCobertAtu In lobjstcCobert
                            lCobertDePara = lrsTab_Ped_Cobert!Cod_Cobert
                        
                        If lobjstcCobertAtu.Cod_Cobert = lCobertDePara Then
                            
                            If (Vlr_Pr_Atu - PerDifBatVlrAtual) > 0 Or (Vlr_Pr_Atu - PerDifBatVlrAtual) < 0 Then

                                'OBTEM A FORMULA ORIGINAL SEM O BATENTE POIS QUANDO PREMIO É NEGATIVO É NECESSARIO MONTAR NOVAMENTE
                                If InStr(1, lobjstcCobertAtu.Formula, ";BAT=") > 0 Then
                                    intPosicao = InStr(1, lobjstcCobertAtu.Formula, ";BAT=") - 1
                                Else
                                    intPosicao = Len(lobjstcCobertAtu.Formula)
                                End If

                                StrFormula = Mid(lobjstcCobertAtu.Formula, 1, intPosicao)

                                'TRATAMENTO DO PREMIO NEGATIVO CONFORME ALINHADO COM O PRODUTO (SE FOR NEGATIVO AGRAVAR TODAS OS PREMIOS EM 10 % UTILIZANDO O VALOR ANTERIOR)
                                If BlnVlrPrNegativo Then
                                    PerDifBatVlrAtual = DifLimBatente
                                    VlrPrCoberAnt = lrsTab_Ped_Cobert!VAL_PR
                                    lobjstcCobertAtu.Val_Pr_Tarif_Anu = lrsTab_Ped_Cobert!Val_Pr_Anu + (lrsTab_Ped_Cobert!Val_Pr_Anu * PerDifBatVlrAtual)
                                    lobjstcCobertAtu.Val_Pr_Tarif = lrsTab_Ped_Cobert!VAL_PR + (lrsTab_Ped_Cobert!VAL_PR * PerDifBatVlrAtual)
                                    lobjstcCobertAtu.Val_Pr_Resg_Anu = lrsTab_Ped_Cobert!Val_Pr_Anu + (lrsTab_Ped_Cobert!Val_Pr_Anu * PerDifBatVlrAtual)
                                    lobjstcCobertAtu.Val_Pr_Resg = lrsTab_Ped_Cobert!VAL_PR + (lrsTab_Ped_Cobert!VAL_PR * PerDifBatVlrAtual)
                                    lobjstcCobertAtu.VAL_PR = lrsTab_Ped_Cobert!VAL_PR + (lrsTab_Ped_Cobert!VAL_PR * PerDifBatVlrAtual)
                                    lobjstcCobertAtu.Val_Pr_Anu = lrsTab_Ped_Cobert!Val_Pr_Anu + (lrsTab_Ped_Cobert!Val_Pr_Anu * PerDifBatVlrAtual)
                                    lobjstcCobertAtu.VAL_PR_BATENTE = VlrPrCoberAnt

                                    If VlrPrCoberAnt < lobjstcCobertAtu.VAL_PR Then
                                        StrPerBat = "- " & Round(((lobjstcCobertAtu.VAL_PR - VlrPrCoberAnt) * 100) / lobjstcCobertAtu.VAL_PR, 4) & "%"
                                    Else
                                        StrPerBat = Round(((lobjstcCobertAtu.VAL_PR - VlrPrCoberAnt) * 100) / lobjstcCobertAtu.VAL_PR, 4) & "%"
                                    End If

                                Else
                                    VlrPrCoberAnt = lobjstcCobertAtu.VAL_PR
                                    lobjstcCobertAtu.Val_Pr_Tarif_Anu = lobjstcCobertAtu.Val_Pr_Anu + (PerAuxBatente * PerDifBatVlrAtual)
                                    lobjstcCobertAtu.Val_Pr_Tarif = lobjstcCobertAtu.VAL_PR + (PerAuxBatente * PerDifBatVlrAtual)
                                    lobjstcCobertAtu.Val_Pr_Resg_Anu = lobjstcCobertAtu.Val_Pr_Anu + (PerAuxBatente * PerDifBatVlrAtual)
                                    lobjstcCobertAtu.Val_Pr_Resg = lobjstcCobertAtu.VAL_PR + (PerAuxBatente * PerDifBatVlrAtual)
                                    lobjstcCobertAtu.VAL_PR = lobjstcCobertAtu.VAL_PR + (PerAuxBatente * PerDifBatVlrAtual)
                                    lobjstcCobertAtu.Val_Pr_Anu = lobjstcCobertAtu.Val_Pr_Anu + (PerAuxBatente * PerDifBatVlrAtual)
                                    lobjstcCobertAtu.VAL_PR_BATENTE = VlrPrCoberAnt

                                    If VlrPrCoberAnt < lobjstcCobertAtu.VAL_PR Then
                                        StrPerBat = "- " & Round(((PerAuxBatente * IIf(PerDifBatVlrAtual < 0, PerDifBatVlrAtual * -1, PerDifBatVlrAtual)) * 100) / lobjstcCobertAtu.VAL_PR, 4) & "%"
                                    Else
                                        StrPerBat = Round(((PerAuxBatente * IIf(PerDifBatVlrAtual < 0, PerDifBatVlrAtual * -1, PerDifBatVlrAtual)) * 100) / lobjstcCobertAtu.VAL_PR, 4) & "%"
                                    End If

                                End If

                                If lobjstcCobertAtu.VAL_PR < 0 Then
                                    BlnVlrPrNegativo = True
                                    GoTo CalculaPremioNegativo
                                End If

                                lobjstcCobertAtu.Formula = StrFormula & ";BAT=" & Round(lobjstcCobertAtu.VAL_PR, 2) & "/" & StrPerBat

                            End If

                        End If
                    Next lobjstcCobertAtu
                End If

                lrsTab_Ped_Cobert.MoveNext
            Loop
        End If
    End If

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lbdTab_Ped_Cobert = Nothing
    Set lobjstcCobertAtu = Nothing
    Set lobjCobert = Nothing

    Return

End Sub

Function ConsisteRegraBatente(mobjstcPed As stcA46V702B, _
                              lPer_Coml As Double, _
                              ConsisteApolice As Boolean, _
                              Optional ByRef lrsTab_Ped_Cobert As ADODB.Recordset, _
                              Optional ByRef Vlr_Pr_Ant As Double, _
                              Optional ByRef Vlr_Pr_Atu As Double, _
                              Optional pStcItem As stcA46V708B) As Boolean
    Dim lobjstcCobertura As stcA46V704B
    Dim lobjstcPedClau As stcA46V703B
    Dim lbdTab_Ped_Cobert As ADODB.Connection
    Dim lSelect As String
    Dim lrsTab_Ped_Corr As clsYasRecordSet
    Dim lrsNumEndo As clsYasRecordSet
    Dim lrsNumEndRisco As clsYasRecordSet
    Dim lrsTabPedLoc As clsYasRecordSet
    Dim lrsAssis24hrs As clsYasRecordSet
    Dim pMensagem As String
    Dim VlrLimBatenteDesconto As Double
    Dim VlrLimBatenteAgravo As Double
    Dim PerBatente As Double
    Dim DifLimBatente As Double
    Dim LngNumEndo As Long
    Dim lAuxCobert As Integer
    Dim lCobertNaoExiste As Boolean
    Dim lBookMarkAnt As Integer
    Dim lDeParaMultiplo As Boolean
    Dim lValIPC As Double
    Dim lobjstcLogBatente As stcA46V745B

    Set lrsTab_Ped_Cobert = New ADODB.Recordset
    Set lobjstcLogBatente = New stcA46V745B

    Vlr_Pr_Ant = 0
    lAuxCobert = 0
    lBookMarkAnt = 0

    ConsisteRegraBatente = True
    
    If mobjstcPed.Num_Apol_An = 0 Or mobjstcPed.Cod_Estip > 1000 Then
        ConsisteRegraBatente = False
        GoTo SAIDA
    End If

    If mobjstcPed.objstcCtrlEmis.Suprimir_Batente = True Then
        lobjstcLogBatente.Status_Batente = eStatusBatente_SuprimirBatente
        ConsisteRegraBatente = False
        GoTo SalvaLogBatente
    End If
    
    'abre base de dados
    If M46V999.gfAbrBasDados("GED", M46V999.gTipGed, M46V999.gEndGed, lbdTab_Ped_Cobert, pMensagem) = False Then
        Err.Raise vbObjectError + 513, , "ConsisteRegraBatente: " & pMensagem
        'GoTo SAIDA
    End If

    If ConsisteApolice = True Then

        'VERIFICA SE A APOLICE TEM ENDOSSO
        lSelect = "SELECT TOP 1 NUM_ENDO FROM TAB_APOL_COBERT WITH (NOLOCK) WHERE NUM_APOL = " & mobjstcPed.Num_Apol_An & " ORDER BY NUM_ENDO DESC"

        If M46V999.gfObtRegistro(lbdTab_Ped_Cobert, lSelect, lrsNumEndo, pMensagem) = False Then
            Err.Raise vbObjectError + 513, , "ConsisteRegraBatente: " & pMensagem
            'GoTo SAIDA
        End If

        If lrsNumEndo.RecordCount > 0 Then
            If lrsNumEndo!Num_Endo > 0 Then
                lobjstcLogBatente.Status_Batente = eStatusBatente_ApolicePossuiEndosso
                ConsisteRegraBatente = False
                GoTo SalvaLogBatente
            End If
        Else
            Err.Raise vbObjectError + 513, , "Não foi possivel localizar a apolice anterior número: " & mobjstcPed.Num_Apol_An
            'GoTo SAIDA
        End If

        'VERIFICA SE A QUANTIDADE DE ITENS DA APOLICE É A MESMA
        lSelect = "SELECT COUNT(NUM_ITEM) AS NRO_ITENS_APOL FROM TAB_APOL_LOC WITH (NOLOCK) WHERE NUM_APOL = " & mobjstcPed.Num_Apol_An & " And Num_Endo = 0"

        If M46V999.gfObtRegistro(lbdTab_Ped_Cobert, lSelect, lrsTabPedLoc, pMensagem) = False Then
            Err.Raise vbObjectError + 513, , "ConsisteRegraBatente: " & pMensagem
            'GoTo SAIDA
        End If
        
        If lrsTabPedLoc!NRO_ITENS_APOL = 0 Then
            Err.Raise vbObjectError + 513, , "Não foi possivel localizar locais de risco da apolice anterior número: " & mobjstcPed.Num_Apol_An
        End If

        If mobjstcPed.ColecaoPedLoc.Count <> lrsTabPedLoc!NRO_ITENS_APOL Then
            lobjstcLogBatente.Status_Batente = eStatusBatente_QtdItensDiferente
            ConsisteRegraBatente = False
            GoTo SalvaLogBatente
        End If
    Else
        'VERIFICA SE TEM COMISSÃO DIFERENTE
'        lSelect = "SELECT SUM(PER_CORR) Per_Corr FROM TAB_APOL_CORR WITH (NOLOCK) WHERE NUM_APOL = " & pStcItem.Num_Apol_Anterior
'
'        If M46V999.gfObtRegistro(lbdTab_Ped_Cobert, lSelect, lrsTab_Ped_Corr, pMensagem) = False Then
'            GoTo SAIDA
'        End If
'
'        If lrsTab_Ped_Corr.RecordCount = 0 Then
'            GoTo SAIDA
'        End If
'
'        If lPer_Coml <> lrsTab_Ped_Corr!Per_Corr Then
'            ConsisteRegraBatente = False
'            GoTo SAIDA
'        End If

        'Tarifa 11.18
        If pStcItem.Tip_Emissao <> e_TipEmissao_RenYasuda Then
            ConsisteRegraBatente = False
            GoTo SAIDA
        End If

        'VERIFICA SE O ENDEREÇO DE RISCO É O MESMO
        lSelect = "SELECT TOP 1 Cod_Clasf,Cod_Compl_Clasf,Num_Cep_Risco FROM TAB_APOL_LOC WITH (NOLOCK) WHERE NUM_APOL = " & pStcItem.Num_Apol_Anterior
        lSelect = lSelect & " AND NUM_ITEM = " & pStcItem.Num_Item

        If M46V999.gfObtRegistro(lbdTab_Ped_Cobert, lSelect, lrsNumEndRisco, pMensagem) = False Then
            Err.Raise vbObjectError + 513, , "ConsisteRegraBatente: " & pMensagem
            'GoTo SAIDA
        End If

        If lrsNumEndRisco.RecordCount > 0 Then
            If pStcItem.Num_Cep_Risco <> lrsNumEndRisco!Num_Cep_Risco Then
                lobjstcLogBatente.Status_Batente = eStatusBatente_CEPEnderecoRiscoDiferente
                ConsisteRegraBatente = False
                GoTo SalvaLogBatente
            End If
        Else
            Err.Raise vbObjectError + 513, , "Não foi possivel localizar cep do local de risco da apolice anterior número: " & mobjstcPed.Num_Apol_An
            'GoTo SAIDA
        End If

        'VERIFICA SE A ATIVIDADE É 061 (Bar, Botequim e Restaurantes) APENAS PARA EMPRESA
        If (M46V999.gfTratarNulo(lrsNumEndRisco!Cod_Clasf, tpoString) = "061" Or _
           M46V999.gfTratarNulo(pStcItem.Cod_Clasf, tpoString) = "061") And _
           M46V999.gfTratarNulo(pStcItem.Cod_Ramo, TpoNumInteiro) = 112 And _
           pStcItem.Cod_Iden_Tarif < 1084 Then
            ConsisteRegraBatente = False
            GoTo SAIDA
        End If

        'VERIFICAR A ALTERAÇÃO DE ATIVIDADE
        If Trim(M46V999.gfBuscarDePara(M46V999.gfTratarNulo(lrsNumEndRisco!Cod_Clasf, tpoString), 2, M46V999.gfTratarNulo(lrsNumEndRisco!Cod_Compl_Clasf, tpoString), False)) <> Trim(pStcItem.Cod_Clasf) And _
           Trim(M46V999.gfBuscarDePara(M46V999.gfTratarNulo(lrsNumEndRisco!Cod_Clasf, tpoString), 2, M46V999.gfTratarNulo(lrsNumEndRisco!Cod_Compl_Clasf, tpoString), True)) <> Trim(pStcItem.Cod_Compl_Clasf) Then
            lobjstcLogBatente.Status_Batente = eStatusBatente_AtividadeDiferente
            ConsisteRegraBatente = False
            GoTo SalvaLogBatente
        End If

        'VERIFICA SE O NUMERO DE ITEM É O MESMO
'        If pStcItem.Num_Item <> pStcItem.Num_Item_Anterior Then
'            ConsisteRegraBatente = False
'            GoTo SAIDA
'        End If

        'CONSULTA PREMIO LIQUIDO ANTERIOR
        'UTILIZADO PARA VERIFICAR DE->PARA DE COBERTURA  DE N PARA 1
        lSelect = "SELECT *, Cod_Org as FazBatente FROM TAB_APOL_COBERT WITH (NOLOCK) WHERE NUM_APOL = " & pStcItem.Num_Apol_Anterior

        lSelect = lSelect & " AND NUM_ITEM = " & pStcItem.Num_Item
        lSelect = lSelect & " ORDER BY COD_COBERT"

        
        lrsTab_Ped_Cobert.CursorLocation = adUseClient
        lrsTab_Ped_Cobert.LockType = adLockBatchOptimistic
        Call lrsTab_Ped_Cobert.Open(lSelect, lbdTab_Ped_Cobert, adOpenKeyset)
        lrsTab_Ped_Cobert.ActiveConnection = Nothing

'        If M46V999.gfObtRegistro(lbdTab_Ped_Cobert, lSelect, lrsTab_Ped_Cobert, pMensagem) = False Then
'            GoTo SAIDA
'        End If

        If lrsTab_Ped_Cobert.RecordCount = 0 Then
            Err.Raise vbObjectError + 513, , "Não foi possivel localizar a apolice anterior número: " & pStcItem.Num_Apol_Anterior
        End If

        'VERIFICAR SE A QUANTIDADE DE COBERTURAS SÃO IGUAIS
        If pStcItem.ColecaoPedCobert.Count = lrsTab_Ped_Cobert.RecordCount And pStcItem.Cod_Iden_Tarif < 1080 And pStcItem.Cod_Iden_Tarif >= 1076 Then
            For Each lobjstcCobertura In pStcItem.ColecaoPedCobert

                If lobjstcCobertura.Cod_Cobert = lrsTab_Ped_Cobert!Cod_Cobert Then
                    'VERIFICA SE TEM DIFERENÇA DE VALORES NO LMG DAS COBERTURAS
                    If lobjstcCobertura.Val_IS <> lrsTab_Ped_Cobert!Val_IS Then
                        ConsisteRegraBatente = False
                        GoTo SAIDA
                    End If

                    Vlr_Pr_Ant = Vlr_Pr_Ant + lrsTab_Ped_Cobert!VAL_PR
                    GoTo Proximo
                Else
                    ConsisteRegraBatente = False
                    GoTo SAIDA
                End If
Proximo:
                lrsTab_Ped_Cobert.MoveNext
            Next lobjstcCobertura
        Else
            If pStcItem.ObjPai.Cod_Produto >= 1080 Then
                Do Until lrsTab_Ped_Cobert.EOF
                    For Each lobjstcCobertura In pStcItem.ColecaoPedCobert
                        If M46V999.gfBuscarDePara(lrsTab_Ped_Cobert!Cod_Cobert, 1) = lobjstcCobertura.Cod_Cobert Then
                            'VERIFICA SE TEM DIFERENÇA DE VALORES NO LMG DAS COBERTURAS
                            lDeParaMultiplo = False
'                            If M46V999.gCorrecaoMonetariaIPC Then
'                                lValIPC = Round(lrsTab_Ped_Cobert!Val_IS * (1 + (mobjstcPed.objstcCtrlEmis.Indice_IPC / 100)), 2)
'                            Else
                                lValIPC = lrsTab_Ped_Cobert!Val_IS
'                            End If
                            
                            If lobjstcCobertura.Val_IS <> lValIPC Then
                                lBookMarkAnt = lrsTab_Ped_Cobert.AbsolutePosition
                                lrsTab_Ped_Cobert.MoveFirst
                                
                                'VERIFICAR SE TEM COBERTURAS COM O MESMO DE-> PARA E VERIFICA O VALOR DA IS
                                lDeParaMultiplo = False
                                Do Until lrsTab_Ped_Cobert.EOF
                                    If M46V999.gfBuscarDePara(lrsTab_Ped_Cobert!Cod_Cobert, 1) = lobjstcCobertura.Cod_Cobert Then
                                        If lobjstcCobertura.Val_IS = lrsTab_Ped_Cobert!Val_IS Then 'lValIPC Then
                                            lDeParaMultiplo = True
                                            lobjstcCobertura.RealizaCalBatente = True
                                            lrsTab_Ped_Cobert!FazBatente = 1
                                            Exit Do
                                        End If
                                    End If
                                
                                lrsTab_Ped_Cobert.MoveNext
                                Loop
                                
                                'VOLTA O RS PARA A POSIÇÃO ATUAL ANTES DA VERIFICAÇÃO ACIMA
                                lrsTab_Ped_Cobert.AbsolutePosition = lBookMarkAnt
                                
                                If lDeParaMultiplo = False Then
                                    lobjstcCobertura.RealizaCalBatente = False
                                    lrsTab_Ped_Cobert!FazBatente = 0
                                    Exit For
                                End If
                            End If
                                    
                            lobjstcCobertura.RealizaCalBatente = True
                            If lDeParaMultiplo = False Then
                                lrsTab_Ped_Cobert!FazBatente = 1
                            End If
                            Exit For
                        End If
                    Next lobjstcCobertura
                    
                    lrsTab_Ped_Cobert.MoveNext
                Loop
            Else
                ConsisteRegraBatente = False
                GoTo SAIDA
            End If
        End If

        'VERIFICA SE FOI ALTERADO A ASSISTENCIA 24HRS
'        lSelect = "SELECT * FROM TAB_APOL_CLAU WITH (NOLOCK) WHERE"
'        lSelect = lSelect & " NUM_APOL = " & pStcItem.Num_Apol_Anterior
'        lSelect = lSelect & " AND NUM_ITEM = " & pStcItem.Num_Item
'        lSelect = lSelect & " AND COD_CLAU IN(400,401,402,404,405)"
'
'        If M46V999.gfObtRegistro(lbdTab_Ped_Cobert, lSelect, lrsAssis24hrs, pMensagem) = False Then
'            GoTo SAIDA
'        End If
'
'        If lrsAssis24hrs.RecordCount > 0 Then
'            For Each lobjstcPedClau In pStcItem.ColecaoPedClau
'                If lobjstcPedClau.gEhAss24h Then
'                    If lrsAssis24hrs!Cod_Clau <> lobjstcPedClau.Cod_Clau Then
'                        ConsisteRegraBatente = False
'                        GoTo SAIDA
'                    End If
'                End If
'            Next lobjstcPedClau
'        Else
'            ConsisteRegraBatente = False
'            GoTo SAIDA
'        End If
        
    End If

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcCobertura = Nothing
    Set lbdTab_Ped_Cobert = Nothing
    Set lrsTab_Ped_Corr = Nothing
    Set lrsNumEndo = Nothing
    Set lrsNumEndRisco = Nothing
    Set lrsTabPedLoc = Nothing
    Set lrsAssis24hrs = Nothing
    Set lobjstcPedClau = Nothing

    Return
    Exit Function

SalvaLogBatente:
    Dim lCobert As stcA46V704B
    
    If Not pStcItem Is Nothing Then
        For Each lCobert In pStcItem.ColecaoPedCobert
            Set lCobert.Log_Batente = lobjstcLogBatente
        Next
    Else
        If Not mobjstcPed Is Nothing Then
            For Each pStcItem In mobjstcPed.ColecaoPedLoc
                For Each lCobert In pStcItem.ColecaoPedCobert
                    Set lCobert.Log_Batente = lobjstcLogBatente
                Next
            Next
        End If
    End If
    
    GoTo SAIDA:
End Function

Public Function mfRetornarCoeficiente(ByRef pobjstcPedLoc As stcA46V708B, Tip_Coeficiente As Integer, lPer_Coml As Double, ByVal pPropTabela As String) As Single
    Dim lrst As clsYasRecordSet    'Registro.
    Dim lSql As String        'SQL
    Dim lbdRetornaCoef As ADODB.Connection
    Dim UltimoErro As String

    On Error GoTo ControleErro

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdRetornaCoef, UltimoErro) = False Then
        Exit Function
    End If
    
    If M46V999.gCodProduto < 1080 Then
        If InStr(CStr(lPer_Coml), ".") > 0 Or InStr(CStr(lPer_Coml), ",") > 0 Then
            lPer_Coml = Fix(lPer_Coml) + 1
        End If
    End If
    
    lSql = "select " & pPropTabela & "Func_RetornarCoeficiente (" & pobjstcPedLoc.Cod_Ramo & ","
    lSql = lSql & IIf(M46V999.gCodProduto < 1080, "0", M46V111.gCod_plano) & ","  '0 = vale para todo mundo
    lSql = lSql & pobjstcPedLoc.Cod_Iden_Tarif & ","
    lSql = lSql & Tip_Coeficiente & ","    '1 = comissão ; 2 = cep
    lSql = lSql & IIf(Tip_Coeficiente = 1, 0, pobjstcPedLoc.M46V999.M46V103.lCod_Cobert) & ","    'para comissão este parametro é zerado
    lSql = lSql & IIf(Tip_Coeficiente = 1, 0, pobjstcPedLoc.Tip_Atividade) & ","    'para comissão este parametro é zerado
    lSql = lSql & Replace(IIf(M46V999.gCodProduto < 1080, lPer_Coml - 1, lPer_Coml), ",", ".") & ","   'Valor inicial
    lSql = lSql & Replace(lPer_Coml, ",", ".") & ")"       'Valor final

    If M46V999.gfObtRegistro(lbdRetornaCoef, lSql, _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    mfRetornarCoeficiente = IIf(IsNull(lrst.Fields(0).Value), 1, lrst.Fields(0).Value)
    Set lrst = Nothing
    Exit Function
ControleErro:
    Set lrst = Nothing
    mfRetornarCoeficiente = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "clsM46V103.mfRetornarCoeficiente ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lSql)
    Else
        Call M46V999.gpGraLog(0, "clsM46V103.mfRetornarCoeficiente ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lSql)
    End If
End Function

Public Function mfRetornarDescontoLMI(ByRef pobjstcPedLoc As stcA46V708B) As Double
    Dim lrst As clsYasRecordSet    'Registro.
    Dim lSql As String        'SQL
    Dim lbdRetornaCoef As ADODB.Connection
    Dim UltimoErro As String

    On Error GoTo ControleErro

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdRetornaCoef, UltimoErro) = False Then
        Exit Function
    End If
    
    lSql = "EXEC proc_RetornaDescontoLMIUnico " & pobjstcPedLoc.Cod_Ramo & ", " & _
                    pobjstcPedLoc.ObjPai.Cod_Estip & ", " & pobjstcPedLoc.ObjPai.Cod_Produto & ", " & _
                    pobjstcPedLoc.ObjPai.ColecaoPedLoc.Count

    If M46V999.gfObtRegistro(lbdRetornaCoef, lSql, _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    mfRetornarDescontoLMI = IIf(IsNull(lrst.Fields(0).Value), 1, lrst.Fields(0).Value)
    Set lrst = Nothing
    Exit Function
ControleErro:
    Set lrst = Nothing
    mfRetornarDescontoLMI = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "clsM46V103.mfRetornarDescontoLMI ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lSql)
    Else
        Call M46V999.gpGraLog(0, "clsM46V103.mfRetornarDescontoLMI ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lSql)
    End If
End Function

Public Function mfRetornarCoeficienteCorretor(ByRef pobjstcPedLoc As stcA46V708B, lPer_Coml As Double, ByVal pPropTabela As String) As Double
    Dim lrst As clsYasRecordSet    'Registro.
    Dim lSql As String        'SQL
    Dim lbdRetornaCoef As ADODB.Connection
    Dim UltimoErro As String

    On Error GoTo ControleErro

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdRetornaCoef, UltimoErro) = False Then
        Exit Function
    End If
    
    'Item 15 - Alterações Corretora Triunfo - Imobiliário
    Dim lPer_ComlAux As String
    lPer_ComlAux = Replace(lPer_Coml, ",", ".")
    lSql = "EXEC proc_RetornaCoeficienteComissaoCorretor " & pobjstcPedLoc.Cod_Ramo & ", " & M46V111.gCod_plano & ", " & pobjstcPedLoc.Cod_Iden_Tarif & ", " & pobjstcPedLoc.ObjPai.ColecaoPedCorr(1).Cod_Corr & ", " & lPer_ComlAux
    
    If M46V999.gfObtRegistro(lbdRetornaCoef, lSql, _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    If lrst.EOF Then
        mfRetornarCoeficienteCorretor = 0
    Else
        mfRetornarCoeficienteCorretor = IIf(IsNull(lrst.Fields(0).Value), 1, lrst.Fields(0).Value)
    End If
    Set lrst = Nothing
    Exit Function
ControleErro:
    Set lrst = Nothing
    mfRetornarCoeficienteCorretor = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "clsM46V103.mfRetornarCoeficienteCorretor ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lSql)
    Else
        Call M46V999.gpGraLog(0, "clsM46V103.mfRetornarCoeficienteCorretor ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lSql)
    End If
End Function

Public Function mpVerificaFranqGrdRisco(Cod_Ramo As Integer, Cod_Plano As Integer, Cod_Tarifa As Integer, _
                                        Val_IS As Double) As Boolean
                                        
    Dim lrst As clsYasRecordSet    'Registro.
    Dim lSql As String        'SQL
    Dim lbdRetornaFranquia As ADODB.Connection
    Dim UltimoErro As String

    On Error GoTo ControleErro
    
    mpVerificaFranqGrdRisco = False

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdRetornaFranquia, UltimoErro) = False Then
        Exit Function
    End If
        
    lSql = "Exec proc_RetornarFaixaFranquiaGrdRisco "
    lSql = lSql & Cod_Ramo & ","
    lSql = lSql & Cod_Plano & ","
    lSql = lSql & Cod_Tarifa & ","
    lSql = lSql & Replace(Val_IS, ",", ".")

    If M46V999.gfObtRegistro(lbdRetornaFranquia, lSql, _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    mpVerificaFranqGrdRisco = lrst.Fields(0).Value

    Exit Function
ControleErro:
    
    mpVerificaFranqGrdRisco = False
    
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "clsM46V103.mpVerificaFranqGrdRisco ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lSql)
    Else
        Call M46V999.gpGraLog(0, "clsM46V103.mpVerificaFranqGrdRisco ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lSql)
    End If

End Function

Sub RetornaVlrMultFranquia(ByRef pStcItem As stcA46V704B)

    Dim Vlr_Franquia As Double
    Dim lobjstcPedFranquia As stcA46V738B

    If pStcItem.Cod_Cobert = 1 Then
        pStcItem.EhGrandesRiscos = False
        'NOVO-COTADOR
        If mpVerificaFranqGrdRisco(pStcItem.Cod_Ramo, pStcItem.M46V111.gCod_plano, M46V999.gCodProduto, pStcItem.Val_IS) And _
           (pStcItem.ObjPai.ObjPai.StatusCotSiscota <> enum_StatusTransm_EnvioKit And pStcItem.ObjPai.ObjPai.StatusCotSiscota <> enum_StatusTransm_Aprovada) Then
                        Vlr_Franquia = (Round(pStcItem.VAL_PR, 5) / 5000)
            
            'ADICIONADO O TRATAMENTO PARA SEMPRE ARREDONDAR PARA CIMA DESPREZANDO AS CASAS DECIMAIS
            If InStr(CStr(Vlr_Franquia), ".") > 0 Or InStr(CStr(Vlr_Franquia), ",") > 0 Then
                Vlr_Franquia = Fix(Vlr_Franquia) + 1
            End If
            
                Vlr_Franquia = Vlr_Franquia * 5000

                For Each lobjstcPedFranquia In pStcItem.ColecaoNovaFranquia
                        lobjstcPedFranquia.Val_Min_Franq = Vlr_Franquia
                Next lobjstcPedFranquia
                pStcItem.EhGrandesRiscos = True
        End If
    End If

SAIDA:
    Exit Sub

End Sub

Sub RetornaVlrRgrBatente_NovoProduto(ByRef pStcItem As stcA46V708B, _
                         ByRef lobjstcCobert As Collection, _
                         lrsTab_Ped_Cobert As ADODB.Recordset)

    Dim Vlr_Pr_Ant As Double
    Dim Vlr_Pr_Atu As Double
    Dim PerBatente As Double
    Dim DifLimBatente As Double
    Dim lobjstcCobertAtu As stcA46V704B
    Dim PerDifBatVlrAtual As Double
    Dim lobjCobert As clsA46V704A
    Dim StrPerBat As String
    Dim VlrPrCoberAnt As Double
    Dim StrFormula          As String
    Dim intPosicao          As Integer
    Dim lCobertDePara       As Integer
    Dim lCoefComisCorrAtu   As Double
    Dim lCoefTaxaFinalAtu   As Double
    Dim lAgravAtu           As Double
    Dim lCoefComisCorrAnt   As Double
    Dim lCoefTaxaFinalAnt   As Double
    Dim lAgravAnt           As Double
    Dim lComisCorrAtu As Double
    Dim lComisCorrAnt As Double
    Dim i As Integer
    Dim j As Integer

    Set lobjCobert = New clsA46V704A
    Set lobjCobert.M46V999 = M46V999

    StrFormula = ""
    intPosicao = 0
    PerBatente = 0
    lCobertDePara = 0
    lCoefComisCorrAtu = 0
    lCoefTaxaFinalAtu = 0
    lCoefComisCorrAnt = 0
    lCoefTaxaFinalAnt = 0
    lComisCorrAnt = 0
    lComisCorrAtu = 0
    
    If pStcItem.Cod_Ramo = 112 Then
        If M46V999.gCodProduto >= 1122 Then '11.22
            DifLimBatente = 0.1
        Else
                DifLimBatente = 0.2
        End If
    Else
        DifLimBatente = 0.1
    End If

    If lrsTab_Ped_Cobert.RecordCount > 0 Then
        lrsTab_Ped_Cobert.MoveFirst

        Do Until lrsTab_Ped_Cobert.EOF
            
            If lobjstcCobert.Count > 0 Then
                For Each lobjstcCobertAtu In lobjstcCobert
                    lCobertDePara = M46V999.gfBuscarDePara(lrsTab_Ped_Cobert!Cod_Cobert, 1)

                    If lobjstcCobertAtu.Log_Batente Is Nothing Then
                        Set lobjstcCobertAtu.Log_Batente = New stcA46V745B
                    End If
                    
                    If lobjstcCobertAtu.Cod_Cobert = lCobertDePara And lobjstcCobertAtu.RealizaCalBatente And lrsTab_Ped_Cobert!FazBatente = 1 Then
                        'Desconsiderar se Premio informado
                        If lobjstcCobertAtu.PremioInformado = True Then
                            GoTo ProximoReg
                        End If
                        
                        'Obtem a comissao da formula de calculo
                        lCoefComisCorrAtu = Mid(lobjstcCobertAtu.Formula, InStr(1, UCase(lobjstcCobertAtu.Formula), "CMS=") + 4, InStr(InStr(1, UCase(lobjstcCobertAtu.Formula), "CMS="), UCase(lobjstcCobertAtu.Formula), ";") - InStr(1, UCase(lobjstcCobertAtu.Formula), "CMS=") - 4)
                        lCoefTaxaFinalAtu = Mid(lobjstcCobertAtu.Formula, InStr(1, UCase(lobjstcCobertAtu.Formula), "TXFIM=") + 6, InStr(InStr(1, UCase(lobjstcCobertAtu.Formula), "TXFIM="), UCase(lobjstcCobertAtu.Formula), ";") - InStr(1, UCase(lobjstcCobertAtu.Formula), "TXFIM=") - 6)
                        If lobjstcCobertAtu.Formula Like "AGRAV=" Then
                            lAgravAtu = Mid(lobjstcCobertAtu.Formula, InStr(1, UCase(lobjstcCobertAtu.Formula), "AGRAV=") + 6, InStr(InStr(1, UCase(lobjstcCobertAtu.Formula), "AGRAV="), UCase(lobjstcCobertAtu.Formula), ";") - InStr(1, UCase(lobjstcCobertAtu.Formula), "AGRAV=") - 6)
                        Else
                            lAgravAtu = 0
                        End If
                        For i = 1 To pStcItem.ObjPai.gStcPedidoApolice.ColecaoPedLoc.Count
                            If pStcItem.Num_Item = pStcItem.ObjPai.gStcPedidoApolice.ColecaoPedLoc(i).Num_Item Then
                                For j = 1 To pStcItem.ObjPai.gStcPedidoApolice.ColecaoPedLoc(i).ColecaoPedCobert.Count
                                    If lrsTab_Ped_Cobert!Cod_Cobert = pStcItem.ObjPai.gStcPedidoApolice.ColecaoPedLoc(i).ColecaoPedCobert(j).Cod_Cobert Then
                                    
                                        If Trim(pStcItem.ObjPai.gStcPedidoApolice.ColecaoPedLoc(i).ColecaoPedCobert(j).Formula) = "" Then
                                            GoTo ProximoReg
                                        End If
                                        
                                        With pStcItem.ObjPai.gStcPedidoApolice.ColecaoPedLoc(i).ColecaoPedCobert(j)
                                            lCoefComisCorrAnt = Mid(.Formula, InStr(1, UCase(.Formula), "CMS=") + 4, InStr(InStr(1, UCase(.Formula), "CMS="), UCase(.Formula), " ") - InStr(1, UCase(.Formula), "CMS=") - 4)
                                            lCoefTaxaFinalAnt = Mid(.Formula, InStr(1, UCase(.Formula), "TXFIM=") + 6, InStr(InStr(1, UCase(.Formula), "TXFIM="), UCase(.Formula), " ") - InStr(1, UCase(.Formula), "TXFIM=") - 6)
                                            If .Formula Like "AGRAV=" Then
                                                lAgravAnt = Mid(.Formula, InStr(1, UCase(.Formula), "AGRAV=") + 6, InStr(InStr(1, UCase(.Formula), "AGRAV="), UCase(.Formula), " ") - InStr(1, UCase(.Formula), "AGRAV=") - 6)
                                            Else
                                                lAgravAnt = 0
                                            End If
                                            lobjstcCobertAtu.Log_Batente.Formula_Apol = .Formula
                                        End With
                                        Exit For
                                    Else
                                        lCoefComisCorrAnt = 0
                                        lCoefTaxaFinalAnt = 0
                                    End If
                                Next
                        
                                If lCoefComisCorrAnt <> 0 Then
                                    Exit For
                                End If
                            End If
                        Next
                        
                        lobjstcCobertAtu.Log_Batente.nosso_numero = lobjstcCobertAtu.nosso_numero
                        lobjstcCobertAtu.Log_Batente.Num_Item = lobjstcCobertAtu.Num_Item
                        lobjstcCobertAtu.Log_Batente.Cod_Cobert = lobjstcCobertAtu.Cod_Cobert
                        lobjstcCobertAtu.Log_Batente.Usuario = M46V999.gNomUsuario
                        
'                        'Desconsiderar se houve mudança no coeficiente de comissão
'                        If Round(M46V999.CDblx(lCoefComisCorrAtu), 4) <> Round(M46V999.CDblx(lCoefComisCorrAnt), 4) Then
'                            lobjstcCobertAtu.Log_Batente.Status_Batente = eStatusBatente_ComissaoDiferente
'                            lobjstcCobertAtu.Log_Batente.Formula = lobjstcCobertAtu.Formula
'                            GoTo ProximoReg
'                        End If
                        
                        lobjstcCobertAtu.Log_Batente.Formula = lobjstcCobertAtu.Formula
                        'Vlr_Pr_Atu = (lobjstcCobertAtu.Val_IS * Round((lCoefTaxaFinalAtu / (1 + (lAgravAtu / 100))), 5)) / 100
                        Vlr_Pr_Atu = (lobjstcCobertAtu.Val_IS * Round((lCoefTaxaFinalAtu / lCoefComisCorrAtu), 5)) / 100
                        lobjstcCobertAtu.Log_Batente.Val_Pr_Calc = Vlr_Pr_Atu
                        
                        If pStcItem.Cod_Iden_Tarif >= 1118 Then
                            lCoefTaxaFinalAnt = Round(lrsTab_Ped_Cobert!VAL_PR / lrsTab_Ped_Cobert!Val_IS * 100, 5)
                        End If
                        
                        'Vlr_Pr_Ant = (lrsTab_Ped_Cobert!Val_IS * Round((lCoefTaxaFinalAnt / (1 + (lAgravAnt / 100))), 5)) / 100
                        Vlr_Pr_Ant = (lrsTab_Ped_Cobert!Val_IS * Round((lCoefTaxaFinalAnt / lCoefComisCorrAnt), 5)) / 100
                        
                        lobjstcCobertAtu.Log_Batente.Val_Pr_Base_Apol = Vlr_Pr_Ant
                        lobjstcCobertAtu.Log_Batente.Val_Pr_Pago_Apol = lrsTab_Ped_Cobert!VAL_PR
                        lobjstcCobertAtu.Log_Batente.Per_Dif_Batente = DifLimBatente
                        
                        PerBatente = Vlr_Pr_Atu / Vlr_Pr_Ant - 1
                                               
                        lobjstcCobertAtu.Log_Batente.Per_Dif_Renovacao = Round(PerBatente, 5)
                        
                        'Comparação de Valores do Batente
                        If (PerBatente > DifLimBatente) Then
                            PerDifBatVlrAtual = 0
                            
                            If Vlr_Pr_Atu > Vlr_Pr_Ant Then
                                PerDifBatVlrAtual = 1 + DifLimBatente
                            End If
                            
                            'Tarifa 11.72 - Regra de Renovação - Triunfo
                            Dim NaoLimitarBatente As Boolean
                            
                            'Tranferencia da regra de 917581 - Triunfo para 939571 - BR Insurance
                            If pStcItem.Tip_Emissao = e_TipEmissao_RenYasuda And (M46V999.gCod_Corretor = 939571 Or M46V999.gCod_Corretor = 930446) And (M46V999.gCod_Ramo = 113 Or M46V999.gCod_Ramo = 114) And pStcItem.ObjPai.Cod_Estip <> 20 Then
                                NaoLimitarBatente = True
                            End If
                                                                               
'                            If lobjstcCobertAtu.gVal_Prem_Min >= lobjstcCobertAtu.VAL_PR And lobjstcCobertAtu.VAL_PR > 0 Then
'                                GoTo ProximoReg
'                            End If
                            
                            intPosicao = Len(lobjstcCobertAtu.Formula)

                            StrFormula = Mid(lobjstcCobertAtu.Formula, 1, intPosicao)
                                                        
                            If lrsTab_Ped_Cobert!VAL_PR = 0 Then
                                GoTo ProximoReg
                            End If
                            VlrPrCoberAnt = lrsTab_Ped_Cobert!VAL_PR
                            
                            If lobjstcCobertAtu.Cod_Cobert = 1 Then
                                lobjstcCobertAtu.Val_Pr_Tarif_Anu = (lrsTab_Ped_Cobert!Val_Pr_Anu / lCoefComisCorrAnt) * PerDifBatVlrAtual * lCoefComisCorrAtu
                                lobjstcCobertAtu.Val_Pr_Tarif = (lrsTab_Ped_Cobert!VAL_PR / lCoefComisCorrAnt) * PerDifBatVlrAtual * lCoefComisCorrAtu + pStcItem.gVal_PR_Ass24h
                                lobjstcCobertAtu.Val_Pr_Resg_Anu = (lrsTab_Ped_Cobert!Val_Pr_Anu / lCoefComisCorrAnt) * PerDifBatVlrAtual * lCoefComisCorrAtu
                                lobjstcCobertAtu.Val_Pr_Resg = (lrsTab_Ped_Cobert!VAL_PR / lCoefComisCorrAnt) * PerDifBatVlrAtual * lCoefComisCorrAtu + pStcItem.gVal_PR_Ass24h
                                lobjstcCobertAtu.VAL_PR = (lrsTab_Ped_Cobert!VAL_PR / lCoefComisCorrAnt) * PerDifBatVlrAtual * lCoefComisCorrAtu + pStcItem.gVal_PR_Ass24h
                                lobjstcCobertAtu.Val_Pr_Anu = (lrsTab_Ped_Cobert!Val_Pr_Anu / lCoefComisCorrAnt) * PerDifBatVlrAtual * lCoefComisCorrAtu
                                lobjstcCobertAtu.VAL_PR_BATENTE = VlrPrCoberAnt
                            Else
                                lobjstcCobertAtu.Val_Pr_Tarif_Anu = (lrsTab_Ped_Cobert!Val_Pr_Anu / lCoefComisCorrAnt) * PerDifBatVlrAtual * lCoefComisCorrAtu
                                lobjstcCobertAtu.Val_Pr_Tarif = (lrsTab_Ped_Cobert!VAL_PR / lCoefComisCorrAnt) * PerDifBatVlrAtual * lCoefComisCorrAtu
                                lobjstcCobertAtu.Val_Pr_Resg_Anu = (lrsTab_Ped_Cobert!Val_Pr_Anu / lCoefComisCorrAnt) * PerDifBatVlrAtual * lCoefComisCorrAtu
                                lobjstcCobertAtu.Val_Pr_Resg = (lrsTab_Ped_Cobert!VAL_PR / lCoefComisCorrAnt) * PerDifBatVlrAtual * lCoefComisCorrAtu
                                lobjstcCobertAtu.VAL_PR = (lrsTab_Ped_Cobert!VAL_PR / lCoefComisCorrAnt) * PerDifBatVlrAtual * lCoefComisCorrAtu
                                lobjstcCobertAtu.Val_Pr_Anu = (lrsTab_Ped_Cobert!Val_Pr_Anu / lCoefComisCorrAnt) * PerDifBatVlrAtual * lCoefComisCorrAtu
                                lobjstcCobertAtu.VAL_PR_BATENTE = VlrPrCoberAnt
                            End If
                                
                            'Tarifa 11.72 - Regra de Renovação - Triunfo
                            If NaoLimitarBatente Then
                                
                                lobjstcCobertAtu.Val_Pr_Tarif_Anu = IIf(lobjstcCobertAtu.Val_Pr_Tarif_Anu > lrsTab_Ped_Cobert!Val_Pr_Tarif_Anu, lrsTab_Ped_Cobert!Val_Pr_Anu, lobjstcCobertAtu.Val_Pr_Tarif_Anu)
                                lobjstcCobertAtu.Val_Pr_Tarif = IIf(lobjstcCobertAtu.Val_Pr_Tarif > lrsTab_Ped_Cobert!Val_Pr_Tarif, lrsTab_Ped_Cobert!Val_Pr_Tarif, lobjstcCobertAtu.Val_Pr_Tarif)
                                lobjstcCobertAtu.Val_Pr_Resg_Anu = IIf(lobjstcCobertAtu.Val_Pr_Resg_Anu > lrsTab_Ped_Cobert!Val_Pr_Resg_Anu, lrsTab_Ped_Cobert!Val_Pr_Resg_Anu, lobjstcCobertAtu.Val_Pr_Resg_Anu)
                                lobjstcCobertAtu.Val_Pr_Resg = IIf(lobjstcCobertAtu.Val_Pr_Resg > lrsTab_Ped_Cobert!Val_Pr_Resg, lrsTab_Ped_Cobert!Val_Pr_Resg, lobjstcCobertAtu.Val_Pr_Resg)
                                lobjstcCobertAtu.VAL_PR = IIf(lobjstcCobertAtu.VAL_PR > lrsTab_Ped_Cobert!VAL_PR, lrsTab_Ped_Cobert!VAL_PR, lobjstcCobertAtu.Val_Pr_Tarif_Anu)
                                lobjstcCobertAtu.Val_Pr_Anu = IIf(lobjstcCobertAtu.Val_Pr_Anu > lrsTab_Ped_Cobert!Val_Pr_Anu, lrsTab_Ped_Cobert!Val_Pr_Anu, lobjstcCobertAtu.Val_Pr_Anu)
                                
                                Debug.Print ("PREMIO ASSUMIDO: " + CStr(CDec(lobjstcCobertAtu.VAL_PR)))
                                
                            End If

                            If VlrPrCoberAnt < lobjstcCobertAtu.VAL_PR Then
                                StrPerBat = "- " & Round(((lobjstcCobertAtu.VAL_PR - VlrPrCoberAnt) * 100) / IIf(VlrPrCoberAnt = 0, 1, VlrPrCoberAnt), 4) & "%"
                            Else
                                StrPerBat = Round(((lobjstcCobertAtu.VAL_PR - VlrPrCoberAnt) * 100) / IIf(VlrPrCoberAnt = 0, 1, VlrPrCoberAnt), 4) & "%"
                            End If

                            lobjstcCobertAtu.Formula = StrFormula & ";BAT=" & Round(lobjstcCobertAtu.VAL_PR, 2) & "/" & StrPerBat
                            lobjstcCobertAtu.Log_Batente.Formula = lobjstcCobertAtu.Formula
                            lobjstcCobertAtu.Log_Batente.Status_Batente = eStatusBatente_Aplicado
                        Else
                            lobjstcCobertAtu.Log_Batente.Status_Batente = eStatusBatente_LimiteInferior
                        End If

                    End If
ProximoReg:
                Next lobjstcCobertAtu
                
            End If

            lrsTab_Ped_Cobert.MoveNext
        Loop
        
    End If

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lobjstcCobertAtu = Nothing
    Set lobjCobert = Nothing

    Return

End Sub


