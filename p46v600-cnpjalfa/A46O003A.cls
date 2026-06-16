VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46O003A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ======================================================================================================================================================
' Class Name: clsA46O003 --> Classe do
' Data/Hora: 01/06/2011
' Sistema: SyasMultWeb
' Simony Sales Andrade
' ======================================================================================================================================================

Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long

Public Enum Enum_TipDocumento
    Enum_TipDocumentoCot = 1        'Cota��es e propostas.
    Enum_TipDocumentoApol = 2       'Apolice Vigentes
    Enum_TipDocumentoProt = 3       'Protocolos
End Enum

Public Enum Enum_SitDocumento
    Enum_SitDocumentoTodas = 0         'Todas
    Enum_SitDocumentoNaoTran = 1       'N�o Transmitidas
    Enum_SitDocumentoTran = 2          'Transmitidas
    Enum_SitDocumentoInspec = 3        'Inspecao n�o solicitadas
End Enum

Public Enum E_VerifBoleto
    E_VerifBoleto_Nao = 0
    E_VerifBoleto_Questiona = 1
    E_VerifBoleto_Alerta = 2
End Enum
Public MensagemAlerta As String


Public UltimoErro As String
Public objstcUsuario As stcA46V725B
'-------------
Private mPropTabela As String

Const Class_Name As String = "clsA46O003A"

Private mClsM46V999 As clsM46V999
Private mClsM46O001 As clsM46O001
Private mClsM46V103 As clsM46V103
Private mClsM46V111 As clsM46V111
Private mClsM46V998 As clsM46V998
Private mClsM99V001 As clsM99V001
Private mClsM46V001 As clsM46V001

Public Enum Enum_Tipo_LogEmissao
    Enum_Tipo_LogEmissao_Inicio = 1
    Enum_Tipo_LogEmissao_Fim = 10
    Enum_Tipo_LogEmissao_Erro = 3
    Enum_Tipo_LogEmissao_Abortado = 4
End Enum

Public Function gfSalvarParcelasParcelamento( _
                                pNossoNumero As String, _
                                pParcelas_Impressao_Avista As String, _
                                pParcelas_Impressao_30dias As String, _
                                pParcelas_Impressao_AvistaDebito As String, _
                                pParcelas_Impressao_30diasDebito As String, _
                                pParcelas_Impressao_Cartao As String, _
                                ByRef pMensagem As String, tagDynaTrace As String) As Boolean


    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfSalvarParcelasParcelamento"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim objCtrl_Emis_compl As stcA46V712B
    Dim clsCtrl_Emis_compl As clsA46V712A

    Set objCtrl_Emis_compl = New stcA46V712B
    Set clsCtrl_Emis_compl = New clsA46V712A

    objCtrl_Emis_compl.NOSSO_NUMERO = pNossoNumero
    objCtrl_Emis_compl.Parcelas_Impressao_Avista = pParcelas_Impressao_Avista
    objCtrl_Emis_compl.Parcelas_Impressao_30dias = pParcelas_Impressao_30dias
    objCtrl_Emis_compl.Parcelas_Impressao_AvistaDebito = pParcelas_Impressao_AvistaDebito
    objCtrl_Emis_compl.Parcelas_Impressao_30diasDebito = pParcelas_Impressao_30diasDebito
    objCtrl_Emis_compl.Parcelas_Impressao_Cartao = pParcelas_Impressao_Cartao

    Set clsCtrl_Emis_compl.M46V999 = Me.M46V999

    If Not clsCtrl_Emis_compl.mfAlterarStringParcelamento(objCtrl_Emis_compl, objstcUsuario.PropTabela, "", tagDynaTrace) Then
        pMensagem = clsCtrl_Emis_compl.UltimoErro
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, eSeveridadeErro_Medio, Class_Name, methodName)
        Exit Function
    End If

    gfSalvarParcelasParcelamento = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

    Exit Function
ControleErros:
    gfSalvarParcelasParcelamento = False
    pMensagem = " gfSalvarParcelasParcelamento ERRO - " & Err & " " & Err.Description
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, eSeveridadeErro_Medio, Class_Name, methodName)

End Function

Private Sub Class_Initialize()

    Set mClsM46V999 = Nothing
    Set mClsM46O001 = Nothing
    Set mClsM46V103 = Nothing
    Set mClsM46V111 = Nothing
    Set mClsM46V998 = Nothing
    Set mClsM99V001 = Nothing
    Set mClsM46V001 = Nothing

    UltimoErro = ""
    M46V999.gBaseLocal = False
    M46V999.gPrograma = "D046O003"
    M46V999.gProibidoMsgBox = True
    M46V999.VerificarNormas = True
    Set objstcUsuario = New stcA46V725B
    Call M46V999.gpInicia(mPropTabela)
    

End Sub

Private Sub Class_Terminate()
    Set objstcUsuario = Nothing
    Set mClsM46V999 = Nothing
    Set mClsM46O001 = Nothing
    Set mClsM46V103 = Nothing
    Set mClsM46V111 = Nothing
    Set mClsM46V998 = Nothing
    Set mClsM99V001 = Nothing
    Set mClsM46V001 = Nothing
End Sub

Public Property Get versaoDLL() As String
    Err.Clear
    On Error Resume Next
    If M46V999.gAppPath = "" Then Call M46V999.gpInicia(mPropTabela)
    versaoDLL = Format(FileDateTime(M46V999.gAppPath & App.EXEName & ".dll"), "DD/MM/YYYY Hh:Nn:Ss")
    Err.Clear
    On Error GoTo 0
End Property

Public Function gfPesCotInspecao(tagDynaTrace As String) As Collection

'Procedure: pesquisa cota��es e propostas.

    Dim lListaNossoNumero As String         'Lista de nosso_numero.
    Dim lCount As Double    'Contador.
    Dim lI As Double    'Auxiliar.
    Dim lColuna As Integer    'Coluna do grid.
    Dim lObjPed As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED).
    Dim lObjstcPed As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED).

    Dim lColecaoTemp As Collection
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfPesCotInspecao"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro
    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '3. Instancio os objetos
    Set lObjPed = New clsA46V702A
    Set lObjPed.M46V999 = M46V999

    Set lObjstcPed = New stcA46V702B
    Set lObjstcPed.M46V999 = M46V999


    '4. Busco os dados com os par�metros de pesquisa informados.
    If Not lObjPed.mfCarregarColecaoInspecaoWeb(objstcUsuario.codCorretor, objstcUsuario.PropTabela, "") Then

        UltimoErro = lObjPed.UltimoErro
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '5. Se n�o encontrou registro, informa usu�rio.
    If lObjPed.Colecao.Count = 0 Then
        UltimoErro = "INSP0174 - N�o h� cota��o dentro do prazo de solicita��o de inspe��o."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    Set gfPesCotInspecao = lObjPed.Colecao

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    UltimoErro = "[ERR] - gfPesCotInspecao " & Err & " " & Err.Description & UltimoErro
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lObjPed = Nothing
    Set lObjstcPed = Nothing
    Return

End Function


Private Function mfConsistePreCalc(ByRef pStcPed As stcA46V702B, ByRef pColAlertas As Collection, ByRef pColErros As Collection, tagDynaTrace As String) As Boolean
    Dim lMsgErr As String
    Dim lstcPedLoc As stcA46V708B

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfConsistePreCalc"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo TrataErro: mfConsistePreCalc = False

    Set pColErros = New Collection
    Call M46V999.gpInicializar_gErros

    '*** Campos minimos de preenchimento do Pedido
    With pStcPed
        If M46V999.gfPreenchido(.Nom_Segurado) = False Then
            lMsgErr = "FNP9999 - Nome do Segurado N�o preenchido."
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If .Cod_Unid_Prod = 0 Then
            lMsgErr = "FNP9999 - Unidade Produtiva n�o preenchida."
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If .Cod_Prod = 0 Then
            lMsgErr = "FNP9999 - Produtor n�o preenchido."
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If .Cod_Estip = 0 Then
            lMsgErr = "FNP9999 - Plano n�o preenchido."
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If .Cod_Ramo = e_CodRamo_NaoDefinido Then
            lMsgErr = "FNP9999 - Ramo n�o preenchido."
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If .Dat_Ini_Vig = 0 Then
            lMsgErr = "FNP9999 - Data de in�cio de vig�ncia n�o preenchida."
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If .Dat_Ter_Vig = 0 Then
            lMsgErr = "FNP9999 - Data de t�rmino de vig�ncia n�o preenchida."
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If .Dat_Ter_Vig < .Dat_Ini_Vig Then
            lMsgErr = "FNP9999 - T�rmino inferior ao in�cio de vig�ncia."
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If .Dat_Ter_Vig < .Dat_Ini_Vig Then
            lMsgErr = "FNP9999 - T�rmino igual ao in�cio de vig�ncia."
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If .Cod_Iden_Praz = 0 Then
            lMsgErr = "FNP9999 - Tipo de Prazo n�o preenchido."
            Call M46V999.gpMenErro(lMsgErr)
        End If
    End With

    '******

    'Verificar se � PJ quando ramo Condom�nio
    If Not mfConsPessoa(pStcPed, lMsgErr) Then
        Call M46V999.gpMenErro(lMsgErr)
    End If

    'Resolu��o CNSP 218/10.
    If pStcPed.Cod_Ramo = e_CodRamo_Condominio And CVDate(M46V999.gfForData(2, pStcPed.Dat_Ini_Vig)) >= CVDate("01/07/2011") And _
       pStcPed.Cod_Produto < 1014 Then

        lMsgErr = "FGE1094 - Cota��o n�o pode ser realizada devido o disposto na Resolu��o CNSP 218/10."
        Call M46V999.gpMenErro(lMsgErr)
    End If


    'Corretor permitido!
    If pStcPed.ColecaoPedCorr.Count = 0 Then
        Call M46V999.gpMenErro("FNP9999 - Corretor n�o preenchido.")
    ElseIf Not M46V999.gfPreenchido(pStcPed.ColecaoPedCorr(1).gNom_Corretor) Then
        Call M46V999.gpMenErro("FNP9999 - Corretor n�o preenchido.")
    ElseIf objstcUsuario.TipoUsuario <> CStr(e_Tipo_Usu_Yasuda) Then
        'Ficha 1000153 - Automatiza��o do agrupamento de corretores no Syas ap�s o cadastro realizado pelo help-desk.
        If M46V999.gfVerCorrPermitido_TBS_Corr_Grupo(objstcUsuario.codCorretor, pStcPed.ColecaoPedCorr(1).Cod_Corr, CInt(pStcPed.Cod_Ramo), lMsgErr) = False Then
            Call M46V999.gpMenErro(lMsgErr)
        End If
    End If

    If objstcUsuario.TipoAcesso = DIGITACAO Then
        If pStcPed.gPagAntecipado = True Then
            If pStcPed.Val_Pr_Antec = 0 Then
                lMsgErr = "FGE0112 - Valor do pagamento antecipado inv�lido (zerado - campo obrigat�rio quando pagamento antecipado)."
                Call M46V999.gpMenErro(lMsgErr)
            End If
        End If
    End If

    '***(Questionamento!)
    'Desconto de LMI n�o informado
    '    If M46V999.gNovaPropostaLMI And pStcPed.LMI_Unico = 1 And pStcPed.Per_Desc_LMI = 0 And objstcUsuario.TipoAcesso = CALCULO Then
    '        If mfRespostaUsuario(E_IDConsist_DescontoLMI, pColAlertas) <> E_RespQuest_Sim Then
    '            Call mpAddAlerta(E_IDConsist_DescontoLMI, "Desconto de LMI n�o informado. Deseja continuar?", E_TipConsist_Quest, pColAlertas)
    '         End If
    '    End If
    
            
    If (pStcPed.LMI_Unico = 1 And pStcPed.Cod_Dizer <> 0) Then
        If (pStcPed.StatusCotSiscota = enum_StatusTransm_EnvioKit Or pStcPed.StatusCotSiscota = enum_StatusTransm_Aprovada) Then
            'Por hora aprovado, nao fazer nada
            
        Else
            lMsgErr = "N�o � permitido endosso para cota��es LMI, solicitar cota��o."
            Call M46V999.gpMenErro(lMsgErr)
        End If
    End If
    
    'Para cota��es LMI
    If pStcPed.LMI_Unico = 1 And objstcUsuario.TipoAcesso = CALCULO Then
        If pStcPed.ColecaoPedLoc.Count = 1 Then
            If pStcPed.ColecaoPedLoc(1).Tip_Emissao <> e_TipEmissao_Endosso Then
            lMsgErr = "Para cota��es LMI, � necess�rio ter mais de um item"
            Call M46V999.gpMenErro(lMsgErr)
        End If
    End If
    End If

    'N�o localizado registro de item
    If pStcPed.ColecaoPedLoc.Count = 0 Then
        lMsgErr = "FTE0011 - " & M46V999.gPrefixo & "N�o localizado registro de item (c�lculo de pr�mios)."
        Call M46V999.gpMenErro(lMsgErr)
    End If

    'Posiciona os Juros
    If pStcPed.StatusCotSiscota <> enum_StatusTransm_Aprovada And pStcPed.StatusCotSiscota <> enum_StatusTransm_EnvioKit Then
        Call mpAtualizaJurosNovo(pStcPed, tagDynaTrace)
    End If

    '***Consistencia
    Call mpVerComum(pStcPed, pColAlertas, pColErros)

    '--- Consiste os Itens
    For Each lstcPedLoc In pStcPed.ColecaoPedLoc
        If mfConsistePreCalcItem(lstcPedLoc, pColAlertas, pColErros, pStcPed.objstcCtrlEmis.Ind_Bloq_Web, tagDynaTrace) = False Then
            GoTo TrataErro
        End If
    Next

    '***************** Retorno com Sucesso
    If UBound(M46V999.gerros) > 0 Then
        Set pColErros = mfArrayToColecao(M46V999.gerros)
    End If
    mfConsistePreCalc = True
    '*****************
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    Err.Clear
    GoSub FechaObj
    Exit Function

TrataErro:
    mfConsistePreCalc = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err & " - " & ERROR
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lstcPedLoc = Nothing
    Return

End Function

Private Function mfConsistePreCalcItem(ByRef pStcPedLoc As stcA46V708B, _
                                       ByRef pColAlertas As Collection, _
                                       ByRef pColErros As Collection, _
                                       ByRef pIndBloqWeb As String, _
                                       tagDynaTrace As String) As Boolean

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfConsistePreCalcItem"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lStrIdItem As String
    Dim lMsgErr As String
    Dim mDatTerVig As Date
    Dim lMensagem As String

    Dim lObjPed As clsA46V702A
    Dim lstcPedLocBem As stcA46V715B

    Dim msg_disp As String
    Dim msg_01 As String
    Dim msg_02 As String
    Dim sinal_cob20 As Boolean
    Dim sinal_cob86 As Boolean
    Dim sinal_cob64 As Boolean
    Dim sinal_cob66 As Boolean
    Dim mat_cobs
    Dim mat_rub
    Dim X As Long
    Dim lstcCobert As stcA46V704B
    Dim lstcCobert_Acess As stcA46V704B
    Dim lstcCobertLimite As stcA46V704B

    Dim str_sql As String
    Dim cxn As ADODB.Connection
    Dim obj_rs As clsYasRecordSet
    Dim msg_erro As String

    Dim mat_cobs_conce

    Dim lstcCobert_aux As stcA46V704B
    Dim BOL_sinal_cob57 As Boolean
    Dim BOL_sinal_cob124 As Boolean
    Dim BOL_sinal_cob125 As Boolean

    Dim lDivEndereco As Boolean
    Dim lAux As String

    Dim lPerguntaQuest As stcConsistPreCalc
    Dim Per_IS_Vr As Double
    Dim lPercCobertDependencia As Double
    Dim i As Integer
    Dim lValTotalIsVerbasSeparadas As Double
    Dim lValorLimPercCobert As Double
    Dim lExibeMsgLimiteCobert As Boolean
    Dim bValLimCobert As Boolean
'    Dim bValIsOrig As Boolean
    Dim lExibeMsgLimiteMinimoCobert As Boolean
    Dim lExibeMsgLimiteCobertPercentual As Boolean
    Dim lExibeMsgLimiteMaximoCobert As Boolean

    Err.Clear
    On Error GoTo TrataErro: mfConsistePreCalcItem = False
    lStrIdItem = "(Item:" & Format(pStcPedLoc.Num_Item, "000") & ")"
    On Error GoTo 0


    '*** Campos minimos de preenchimento do Item
    With pStcPedLoc
        If .Val_Vr = 0 Then
            lMsgErr = "FNP9999 - Valor do risco n�o preenchido." & lStrIdItem
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If .Num_Cep_Risco = 0 Then
            lMsgErr = "FNP9999 - CEP do risco n�o preenchido." & lStrIdItem
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If Not (M46V999.gfPreenchido(.Cod_Clasf) Or M46V999.gfPreenchido(.Cod_Compl_Clasf) Or M46V999.gfPreenchido(.Dsc_Clasf)) Then
            lMsgErr = "FNP9999 - Classifica��o n�o preenchida." & lStrIdItem
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If .Cod_Constr = 0 Then
            lMsgErr = "FNP9999 - Tipo de constru��o n�o preenchido."
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If .Num_Apol_Anterior <> 0 And .Num_Item_Anterior = 0 And .Cod_DizerItem = 0 Then
            lMsgErr = "FNP9999 - Renova��o da Ap�lice " & .Num_Apol_Anterior & ". N�mero do item n�o preenchido." & lStrIdItem
            Call M46V999.gpMenErro(lMsgErr)
        End If
        If .gISTotal_LC <> 0 And .Cod_DizerItem = 0 And .ObjPai.Cod_Produto >= 1076 Then
            If .Val_Vr_Lc = 0 Then
                lMsgErr = "FNP9999 - Valor do risco Lucros Cessantes B�sica ou Despesas Fixas B�sica n�o preenchido." & lStrIdItem
                Call M46V999.gpMenErro(lMsgErr)
            Else
                Per_IS_Vr = .gISTotal_LC / .Val_Vr_Lc
                If Per_IS_Vr > 1 Then
                    lMsgErr = "Valor do LMG de Lucros Cessantes n�o pode ser maior que Valor do Risco de Lucros Cessantes!" & lStrIdItem
                    Call M46V999.gpMenErro(lMsgErr)
                ElseIf Per_IS_Vr < 0.05 Then
                    lMsgErr = "Valor do LMG/VR  n�o pode ser menor que 5% do Valor de Lucros Cessantes!" & lStrIdItem
                    Call M46V999.gpMenErro(lMsgErr)
                End If
            End If
        End If
    End With

    '***

    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(M46V999.gDataCorteSompo) Then
            If CLng(M46V999.gDataCorteSompo) <= M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If

    '1100589 - Tarifa 1014 e 1015 Julho 2011
    If pStcPedLoc.ObjPai.Cod_Produto >= 1014 And pStcPedLoc.ObjPai.Cod_Ramo = e_CodRamo_Condominio Then
        If val(pStcPedLoc.Cod_Clasf) = 1 And val(pStcPedLoc.Cod_Compl_Clasf) = 17 Then
            lMsgErr = "C�lculo n�o dispon�vel para este tipo de condom�nio. Favor solicitar cota��o � " & lNomeSeguradora & " Seguros." & lStrIdItem
            Call M46V999.gpMenErro(lMsgErr)
        End If
    End If


    'Ficha 1000681
    If pStcPedLoc.Cod_Plano <> 2218 And pStcPedLoc.Cod_Plano <> 2219 And (pStcPedLoc.Cod_Plano < 2228 And pStcPedLoc.Cod_Plano > 2271) Then
        If (pStcPedLoc.Cod_DizerItem <> 4 And pStcPedLoc.Cod_DizerItem <> 10) And _
           (pStcPedLoc.Tip_Cobert <> e_TipCobert_Predio) And _
           ((pStcPedLoc.gTemRegProfissoes = True) And (pStcPedLoc.ObjPai.Tip_Pessoa = e_TipPessoa_Fisica) And (pStcPedLoc.ObjPai.Cod_Ramo <> e_CodRamo_Empresarial)) Then
            lMsgErr = "FGE0085 - Seguro Pessoa F�sica: Aceita��o somente para Pr�dio." & lStrIdItem
            Call M46V999.gpMenErro(lMsgErr)
        End If

        If (pStcPedLoc.Cod_DizerItem <> 4 And pStcPedLoc.Cod_DizerItem <> 10) And (pStcPedLoc.Tip_Cobert <> e_TipCobert_Predio) And _
           ((pStcPedLoc.gTemRegProfissoes = False) And (pStcPedLoc.ObjPai.Tip_Pessoa = e_TipPessoa_Fisica) And (pStcPedLoc.ObjPai.Cod_Ramo = e_CodRamo_Empresarial) And (M46V999.gTipCobertura <> e_TipCobert_Predio)) Then
            lMsgErr = "FGE0085 - Seguro Pessoa F�sica: Aceita��o somente para Pr�dio." & lStrIdItem
            Call M46V999.gpMenErro(lMsgErr)
        End If
    End If

    'Ficha 1000307
    If (pStcPedLoc.Tip_Cobert <> e_TipCobert_Predio) And _
       (pStcPedLoc.ObjPai.Cod_Tip_Doc = e_TipDoc_Proposta And pStcPedLoc.ObjPai.Cod_Produto >= 988 And M46V999.gTipCobertura <> e_TipCobert_Predio) Then
        If pStcPedLoc.gTemRegProfissoes = True And pStcPedLoc.ObjPai.Tip_Pessoa = e_TipPessoa_Fisica And (pStcPedLoc.ObjPai.objstcPedEndereco.Cod_Reg_Cons = -1) Then
            lMsgErr = "FGE0086 - C�digo de Registro da Categoria inv�lido (n�o selecionado). Verifique " & lStrIdItem
            Call M46V999.gpMenErro(lMsgErr)
        End If
    End If

'    If (pStcPedLoc.Tip_Cobert <> e_TipCobert_Predio) And _
'       (pStcPedLoc.ObjPai.Cod_Tip_Doc = e_TipDoc_Proposta And pStcPedLoc.ObjPai.Cod_Produto >= 988) Then
'        If pStcPedLoc.gTemRegProfissoes = True And pStcPedLoc.ObjPai.Tip_Pessoa = e_TipPessoa_Fisica And (Not M46V999.gfPreenchido(pStcPedLoc.ObjPai.objstcPedEndereco.Num_Reg_Cons)) Then
'            lMsgErr = "FGE0087 - N�mero do Registro da Categoria inv�lido (n�o informado). Verifique " & lStrIdItem
'            Call M46V999.gpMenErro(lMsgErr)
'        End If
'    End If


    'Ficha 1000307
    'n�mero do registro profissional no respectivo Conselho da Categoria
    
    'piloto
    
    If (pStcPedLoc.Tip_Cobert <> e_TipCobert_Predio) And _
       (pStcPedLoc.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao And pStcPedLoc.ObjPai.Cod_Produto >= 988 And pStcPedLoc.ObjPai.Cod_Produto <= 1077) Then
        If pStcPedLoc.gTemRegProfissoes = True And pStcPedLoc.ObjPai.Tip_Pessoa = e_TipPessoa_Fisica And ((Not M46V999.gfPreenchido(pStcPedLoc.ObjPai.objstcPedEndereco.Num_Reg_Cons)) Or pStcPedLoc.ObjPai.objstcPedEndereco.Cod_Reg_Cons = 0) Then
            lMsgErr = "Seguro de Pessoa F�sica: Na proposta ser� obrigat�rio informar o n�mero do registro profissional no respectivo Conselho da Categoria." & lStrIdItem & vbCrLf & "Ex.: OAB, CRM, CRO etc."
            Call M46V999.gpMenErro(lMsgErr)
        End If
    End If
    'Consiste vig�ncia e renova��o
    Select Case pStcPedLoc.Tip_Emissao
    Case e_TipEmissao_RenYasuda
        Set lObjPed = New clsA46V702A
        Set lObjPed.M46V999 = M46V999
        
        If lObjPed.mfRetornaTerminoVigencia(pStcPedLoc.ObjPai) <> 0 Then
            mDatTerVig = M46V999.gfForData(2, lObjPed.mfRetornaTerminoVigencia(pStcPedLoc.ObjPai))
            If DateAdd("m", 1, mDatTerVig) < CDate(M46V999.gfForData(2, pStcPedLoc.ObjPai.Dat_Ini_Vig)) Then
                lMsgErr = "FGE0111 - In�cio de vig�ncia superior a 1 m�s " & _
                          "do final de vig�ncia da ap�lice anterior! Favor calcular seguro novo." & lStrIdItem
                Call M46V999.gpMenErro(lMsgErr)
            End If
        End If
    End Select

    'Ficha:900097

    sinal_cob86 = False
    sinal_cob64 = False
    sinal_cob66 = False
    sinal_cob20 = False

    mat_rub = Split("022-13", "-")             '<-- Rubrica a ser considerada: 22-13 Concession�ria

    'Ficha 1000268 - permitir a contrata��o da cobertura 22
    'mat_cobs = Split("22,23,28,33,41,44", ",") '<-- Coberturas n�o permitidas

    mat_cobs = Split("23,28,33,41,44", ",")    '<-- Coberturas n�o permitidas

    msg_01 = "Para cobertura XXX � necess�rio a cobertura YYY ou ZZZ."
    msg_02 = "As coberturas XXX e YYY n�o podem ser contratadas simult�neamente."

    'Ficha :0801779
    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, cxn, lMensagem) = False Then
        Call M46V999.gpMenErro(lMensagem)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If
    '===============



    For Each lstcCobert In pStcPedLoc.ColecaoPedCobert
        If lstcCobert.Val_IS = 0 Then
            If lstcCobert.Cod_Cobert = 1 Then
                lMsgErr = "FNP9999 - Valor da Cobertura B�sica n�o preenchido." & lStrIdItem
                Call M46V999.gpMenErro(lMsgErr)
            Else
                If pStcPedLoc.Cod_DizerItem <> 12 Then
                    lMsgErr = "FNP9999 - Valor da Cobertura """ & lstcCobert.gNomeCobertura & """ n�o preenchido." & lStrIdItem
                    Call M46V999.gpMenErro(lMsgErr)
                End If
            End If
        End If

        'Consiste valor do LMG/VR
        '        If lstcCobert.gCod_Grupo_Cobert = "1" Then
        '            If Not (lstcCobert.Cod_Ramo = e_CodRamo_Condominio And pstcPedLoc.ObjPai.Cod_Produto >= 1012) Then
        '                If lstcCobert.Val_IS > pstcPedLoc.Val_Vr Then
        '                    lMsgErr = "FGE0114 - Valor LMG n�o pode ser maior do que o Valor do Risco." & lStrIdItem
        '                    Call M46V999.gpMenErro(lMsgErr)
        '                End If
        '            End If
        '        End If

        'Ficha :0801779
        'Colocado a consulta fora do If para valida��es de limites na cobertura basica
        str_sql = "SELECT COD_LIM_COBERT,PER_LIM_MAX_COBERT,VAL_LIM_MIN_COBERT,VAL_LIM_MAX_COBERT FROM TAB_COBERT " & _
                  "WHERE COD_RAMO = " & pStcPedLoc.ObjPai.Cod_Ramo & _
                  " AND COD_TARIF_INI <= " & pStcPedLoc.ObjPai.Cod_Produto & _
                  " AND COD_TARIF_TER >= " & pStcPedLoc.ObjPai.Cod_Produto & _
                  " AND TAB_COBERT = " & pStcPedLoc.Cod_Produto & _
                  " AND COD_COBERT = " & lstcCobert.Cod_Cobert
        '==============
        '***gCodPlano = pstcPedLoc.Cod_Produto
        '==============

        If M46V999.gfObtRegistro(cxn, str_sql, obj_rs, lMensagem) = False Then
            Call M46V999.gpMenErro(lMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
        
        If Not (obj_rs.BOF And obj_rs.EOF) And (pStcPedLoc.Cod_DizerItem <> 4 And pStcPedLoc.Cod_DizerItem <> 10) Then
            If pStcPedLoc.ColecaoPedCobert.Count > 0 And lstcCobert.ObjPai.Ind_VerbaSep = 1 Then
                lValTotalIsVerbasSeparadas = 0
                For i = 1 To pStcPedLoc.ColecaoPedCobert.Count
                    If pStcPedLoc.ColecaoPedCobert(i).Cod_Grupo_Cobertura = 5 Or pStcPedLoc.ColecaoPedCobert(i).Cod_Grupo_Cobertura = 1 Then
                        lValTotalIsVerbasSeparadas = lValTotalIsVerbasSeparadas + pStcPedLoc.ColecaoPedCobert(i).Val_IS
                    End If
                Next
            End If
                       
            bValLimCobert = False
        
            For Each lstcCobert_Acess In pStcPedLoc.ColecaoPedCobert
                lPercCobertDependencia = 0
                lValorLimPercCobert = 0
                lExibeMsgLimiteCobert = False
                lExibeMsgLimiteMinimoCobert = False
                lExibeMsgLimiteCobertPercentual = False
                
                'n�o validar coberturas excluidas em endossos
                If pStcPedLoc.Cod_DizerItem > 0 And pStcPedLoc.Tip_Emissao = e_TipEmissao_Endosso And lstcCobert.Val_IS = 0 Then
                    Exit For
                End If

                If pIndBloqWeb = 1 Then
                    bValLimCobert = False
                Else
                    bValLimCobert = True
                End If
                
                If bValLimCobert Then
                    If Trim(val(lstcCobert_Acess.Cod_Cobert)) = Trim(val(obj_rs("COD_LIM_COBERT"))) And lstcCobert.Tab_Cobert < 2000 Then 'N�o faz valida��o para planinho
                        If lstcCobert.ObjPai.Ind_VerbaSep = 1 Then
                            lPercCobertDependencia = (lstcCobert.Val_IS / lValTotalIsVerbasSeparadas) * 100
                            lValorLimPercCobert = lstcCobert_Acess.Val_IS * (obj_rs("PER_LIM_MAX_COBERT") / 100)
                        Else
                            lPercCobertDependencia = (lstcCobert.Val_IS / lstcCobert_Acess.Val_IS) * 100
                            lValorLimPercCobert = lstcCobert_Acess.Val_IS * (obj_rs("PER_LIM_MAX_COBERT") / 100)
                        End If
                
                        If (lPercCobertDependencia > obj_rs("PER_LIM_MAX_COBERT")) And (lstcCobert.Val_IS <> obj_rs("VAL_LIM_MIN_COBERT")) Then
                            lExibeMsgLimiteCobertPercentual = True
                            lExibeMsgLimiteCobert = True
                        End If
                
                        'N�o validar limite PERCENTUAL se for LMI Unico
                        If lExibeMsgLimiteCobertPercentual = True And pStcPedLoc.ObjPai.LMI_Unico = 1 Then
                            Exit For
                        End If
                        
                        If lstcCobert.ObjPai.Ind_VerbaSep = 1 Then
                            If lstcCobert.Val_IS > lValTotalIsVerbasSeparadas Then
                                msg_erro = " FGE-LIM_COBERT - Consist�ncia do Item " & CStr(lstcCobert.Num_Item) & " - COBERTURA " & CStr(lstcCobert.Cod_Cobert) & " COM IS SUPERIOR AO DA COBERTURA " & lstcCobert_Acess.Cod_Cobert & "."
                                Call M46V999.gpMenErro(msg_erro)
                            End If
                        Else
                            If CDbl(lstcCobert.Val_IS) > CDbl(lstcCobert_Acess.Val_IS) And lstcCobert.Cod_Cobert <> 1 Then
                                msg_erro = " FGE-LIM_COBERT - Consist�ncia do Item " & CStr(lstcCobert.Num_Item) & " - COBERTURA " & CStr(lstcCobert.Cod_Cobert) & " COM IS SUPERIOR AO DA COBERTURA " & lstcCobert_Acess.Cod_Cobert & "."
                                Call M46V999.gpMenErro(msg_erro)
                            End If
                        End If
                
                        If pStcPedLoc.Cod_Iden_Tarif >= 1080 Then
                            If pStcPedLoc.Cod_DizerItem = 0 Or (pStcPedLoc.Cod_DizerItem <> 4 And pStcPedLoc.Cod_DizerItem <> 10) Then
                                If pStcPedLoc.Tip_Emissao <> e_TipEmissao_RenYasuda Then
                                    'NOVA TARIFA��O YASUDA MARITIMA
                                    If (lstcCobert.Val_IS < obj_rs("VAL_LIM_MIN_COBERT") And obj_rs("VAL_LIM_MIN_COBERT") > 0) Then
                                        If lstcCobert.ObjPai.Ind_VerbaSep = 1 And (lstcCobert.Cod_Grupo_Cobertura = 1 Or lstcCobert.Cod_Grupo_Cobertura = 5) Then
                                            Exit For
                                        End If
                                        
                                        lExibeMsgLimiteCobert = True
                                        lExibeMsgLimiteMinimoCobert = True
                                        
                                        If pStcPedLoc.ObjPai.StatusCotSiscota = enum_StatusTransm_EnvioKit Or pStcPedLoc.ObjPai.StatusCotSiscota = enum_StatusTransm_Aprovada Then
                                            If pStcPedLoc.ITEM_ORIGEM_C4 <> 0 And (lstcCobert.Tip_Origem_C4 = 2 Or lstcCobert.Tip_Origem_C4 = 3) And lstcCobert.Val_IS = lstcCobert.Val_IS_Orig Then
                                                lExibeMsgLimiteCobert = False
                                                lExibeMsgLimiteMinimoCobert = False
                                            End If
                                        End If
                                    End If
                                    
                                    If (lstcCobert.Val_IS > obj_rs("VAL_LIM_MAX_COBERT") And obj_rs("VAL_LIM_MAX_COBERT") > 0) Then
                                        If lstcCobert.ObjPai.Ind_VerbaSep = 1 And (lstcCobert.Cod_Grupo_Cobertura = 1 Or lstcCobert.Cod_Grupo_Cobertura = 5) Then
                                            Exit For
                                        End If

                                        lExibeMsgLimiteCobert = True
                                        lExibeMsgLimiteMaximoCobert = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                If (pStcPedLoc.ObjPai.StatusCotSiscota = enum_StatusTransm_EnvioKit Or pStcPedLoc.ObjPai.StatusCotSiscota = enum_StatusTransm_Aprovada) Then
                    If pStcPedLoc.ITEM_ORIGEM_C4 <> 0 Then
                        If lExibeMsgLimiteCobert Then
                            If lstcCobert.Tip_Origem_C4 = 2 Or lstcCobert.Tip_Origem_C4 = 3 Then
                                If lstcCobert.Val_IS > lstcCobert.Val_IS_Orig Then
                                    msg_erro = " FGE-LIM_COBERT - Item " & CStr(lstcCobert.Num_Item) & " - Cobertura " & CStr(lstcCobert.Cod_Cobert) & " - O LMI da cobertura ultrapassa o limite permitido. Necess�ria an�lise pr�via da Seguradora."
                                    Call M46V999.gpMenErro(msg_erro)
                                    lExibeMsgLimiteMinimoCobert = False
                                End If
                                
                                If lExibeMsgLimiteMinimoCobert Then
                                    lExibeMsgLimiteCobert = True
                                Else
                                    lExibeMsgLimiteCobert = False
                                End If
                            Else
                                If lstcCobert.Val_IS > lstcCobert.Val_IS_Orig Then
                                    msg_erro = " FGE-LIM_COBERT - Item " & CStr(lstcCobert.Num_Item) & " - Cobertura " & CStr(lstcCobert.Cod_Cobert) & " - O LMI da cobertura ultrapassa o limite permitido. Necess�ria an�lise pr�via da Seguradora."
                                    Call M46V999.gpMenErro(msg_erro)
                                    lExibeMsgLimiteMinimoCobert = False
                                End If
                                
                                If lExibeMsgLimiteMinimoCobert Then
                                    lExibeMsgLimiteCobert = True
                                Else
                                    lExibeMsgLimiteCobert = False
                                End If
                            End If
                        End If
                    End If
                Else
                    If lExibeMsgLimiteMaximoCobert Then
                        lExibeMsgLimiteCobert = False
                    End If
                End If
                
                If lExibeMsgLimiteCobert Then
                    msg_erro = " FGE-LIM_COBERT - Consist�ncia do Item " & CStr(lstcCobert.Num_Item)
                    msg_erro = msg_erro & " - Cobertura " & CStr(lstcCobert.Cod_Cobert) & " - " & lstcCobert.Nom_Cobertura & ":"
                    msg_erro = msg_erro & " O LMI m�ximo permitido � " & CInt(obj_rs("PER_LIM_MAX_COBERT")) & "% da cobertura " & lstcCobert_Acess.Cod_Cobert
                    msg_erro = msg_erro & ", limitado ao minimo de R$ " & Format(obj_rs("VAL_LIM_MIN_COBERT"), "standard")
                    msg_erro = msg_erro & " e ao m�ximo de R$ " & Format(obj_rs("VAL_LIM_MAX_COBERT"), "standard") & "."
                    
                    If (pStcPedLoc.ObjPai.StatusCotSiscota = enum_StatusTransm_EnvioKit Or pStcPedLoc.ObjPai.StatusCotSiscota = enum_StatusTransm_Aprovada) Then 'Se retorno de classe 4 n�o validar
                        msg_erro = msg_erro & " Solicitar An�lise."
                    End If
                    
                    Call M46V999.gpMenErro(msg_erro)
                    Exit For
                End If
            Next
        End If

        If lstcCobert.Cod_Cobert > 1 Then
            'Se for a rubrica:
            If Trim(pStcPedLoc.Cod_Clasf) = val(mat_rub(0)) And Trim(pStcPedLoc.Cod_Compl_Clasf) = val(mat_rub(1)) And pStcPedLoc.ObjPai.Cod_Produto >= 956 Then  '1000871 - Inclus�o de Org�o Categoria Conselho
                'Ficha:900097
                'Conferir se a cob 20,64 ou 66 foram contratadas:
                If lstcCobert.Cod_Cobert = 86 Then sinal_cob86 = True
                If lstcCobert.Cod_Cobert = 20 Then sinal_cob20 = True
                If lstcCobert.Cod_Cobert = 64 Then sinal_cob64 = True
                If lstcCobert.Cod_Cobert = 66 Then sinal_cob66 = True

                'Verificar todas as coberturas se:
                For X = 0 To UBound(mat_cobs)
                    'Est�o dentro da matriz de cobs n�o permitidas;
                    If StrComp(lstcCobert.Cod_Cobert, mat_cobs(X)) = 0 Then
                        If objstcUsuario.TipoAcesso = DIGITACAO And pStcPedLoc.ObjPai.LMI_Unico = 1 Then Exit For
                        Call M46V999.gpMenErro("A cobertura " & lstcCobert.Cod_Cobert & " n�o se aplica para a rubrica " & mat_rub(0) & "-" & mat_rub(1) & "." & lStrIdItem)
                        Exit For
                    End If
                Next X
            End If

            '1200526 - Valida��es Rubrica Concessionarias
            If (Trim(pStcPedLoc.Cod_Clasf) <> mat_rub(0) Or Trim(pStcPedLoc.Cod_Compl_Clasf) <> mat_rub(1)) And (pStcPedLoc.ObjPai.Cod_Produto >= 1036) Then
                mat_cobs_conce = Split("0056,0057,0063,0064,0066,0068,0070,0072,0074,0122,0123,0124,0125,0098,0126,0127,0137,0138,0152,0153,0154", ",")    '<-- Coberturas n�o permitidas
                For X = 0 To UBound(mat_cobs_conce)
                    'Est�o dentro da matriz de cobs n�o permitidas;
                    If StrComp(Format(lstcCobert.Cod_Cobert, "0000"), mat_cobs_conce(X)) = 0 Then
                        ' 1300395 - Plano Hertz - Cobertura 57 e 66 pode ser contratada - Inicio
                        If (lstcCobert.Cod_Cobert = 57 Or lstcCobert.Cod_Cobert = 66) _
                           And pStcPedLoc.Cod_Clasf = 22 And pStcPedLoc.Cod_Compl_Clasf = 11 Then
                            ' Nada a fazer - pode contratar
                        Else
                            ' 1300395 - Plano Hertz - Cobertura 57 e 66 pode ser contratada - Fim
                            Call M46V999.gpMenErro("A cobertura " & lstcCobert.Cod_Cobert & " n�o se aplica para a rubrica diferente de 022-13 - Concession�rias." & lStrIdItem)
                            Exit For
                        End If
                    End If
                Next X
            End If
        End If
        '1001146 - Tarifa 1000 e 1001 Dezembro 2010
        If pStcPedLoc.ObjPai.Cod_Produto > 999 Then
            If (pStcPedLoc.ObjPai.Cod_Ramo = e_CodRamo_Empresarial And lstcCobert.Cod_Cobert = 94) Or _
               (pStcPedLoc.ObjPai.Cod_Ramo = e_CodRamo_Residencial And (lstcCobert.Cod_Cobert = 27 Or _
                                                                        lstcCobert.Cod_Cobert = 29 Or _
                                                                        lstcCobert.Cod_Cobert = 30 Or _
                                                                        lstcCobert.Cod_Cobert = 31)) Then
                'Comparar soma dos Bens com o LMG da cobertura 27 ou 29 ou 30 ou 31 ou 94
                If pStcPedLoc.ColecaoPedLocBem.Count > 0 Then
                    Call mfConsTotBens(Format(lstcCobert.Val_IS, "standard"), lstcCobert.Cod_Cobert, pStcPedLoc, tagDynaTrace)
                End If
            End If
        End If

        '1100589 - Tarifa 1014 e 1015 Julho 2011
        If pStcPedLoc.ObjPai.Cod_Produto >= 1014 Then
            'Valida��o Empresarial - Somat�rio RC maior 2,5Mi
            '            If M46V999.gCod_Ramo = e_CodRamo_Empresarial Then
            '                Select Case lstcCobert.Cod_Cobert
            '                    Case 22, 57, 58, 59, 60, 61, 62, 124, 125
            '                        dblTotalIsRC = dblTotalIsRC + lstcCobert.Val_IS
            '                End Select
            '                If dblTotalIsRC > 2500000 Then
            '                    If objstcUsuario.TipoUsuario = e_Tipo_Usu_Yasuda Then
            '                        If MsgBox("Somat�rio RC maior que o limite permitido." & vbCrLf & "Deseja desconsiderar esta consist�ncia?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            '                            Call M46V999.gpMenErro("- Somat�rio RC maior que o limite permitido.")
            '                        End If
            '                    End If
            '                End If
            '            End If
            'Valida��o Condom�nio - Cobertura 26 - Equipamentos de Inform�tica
            If (pStcPedLoc.ObjPai.Cod_Ramo = e_CodRamo_Condominio) And _
               (lstcCobert.Cod_Cobert = 26) And _
               (val(pStcPedLoc.Cod_Clasf) <> 1 Or val(pStcPedLoc.Cod_Compl_Clasf) <> 17) Then
                Call M46V999.gpMenErro("- Cobertura 26 - Equipamentos de Inform�tica do Condom�nico � v�lida somente para Condom�nio Horizontal Residencial.")
            End If
        End If
    Next

    'Se for a rubrica:
    If Trim(pStcPedLoc.Cod_Clasf) = val(mat_rub(0)) And Trim(pStcPedLoc.Cod_Compl_Clasf) = val(mat_rub(1)) And pStcPedLoc.ObjPai.Cod_Produto >= 956 Then  '1000871 - Inclus�o de Org�o Categoria Conselho
        'Se cob 86, necess�rio a 64 ou 66:
        If sinal_cob86 And Not (sinal_cob64 Or sinal_cob66) Then
            msg_disp = msg_01
            M46V999.gpMenErro (Replace(Replace(Replace(msg_disp, "XXX", "86"), "YYY", "64"), "ZZZ", "66") & lStrIdItem)
        End If
        'Se existem duas que n�o podem ser contratadas ao mesmo tempo;
        If sinal_cob20 And sinal_cob64 Then
            msg_disp = msg_02
            M46V999.gpMenErro (Replace(Replace(msg_disp, "XXX", "20"), "YYY", "64") & lStrIdItem)
        End If
        If sinal_cob20 And sinal_cob66 Then
            msg_disp = msg_02
            M46V999.gpMenErro (Replace(Replace(msg_disp, "XXX", "20"), "YYY", "66") & lStrIdItem)
        End If
        '        If sinal_cob64 And sinal_cob66 Then
        '            msg_disp = msg_02
        '            M46V999.gpMenErro (Replace(Replace(msg_disp, "XXX", "64"), "YYY", "66"))
        '        End If
    End If
    'Ficha 0801749
    '[Verifico se existe mais de um item renovando o mesmo item da ap�lice.]
    '    If pStcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
    '        If InStr(1, lItemRenova, ";" & pStcPedLoc.Num_Apol_Anterior & ";" & pStcPedLoc.Num_Item_Anterior & ";") > 0 Then
    '            Call M46V999.gpMenErro("AGE0154 - Os itens " & Mid(lItemRenova, InStr(1, lItemRenova, ";" & pStcPedLoc.Num_Apol_Anterior & ";" & pStcPedLoc.Num_Item_Anterior & ";") - 4, 4) & " e " & Format(pStcPedLoc.Num_Item, "0000") & " est�o renovando o mesmo item (" & Format(pStcPedLoc.Num_Item_Anterior, "0000") & ") da ap�lice " & Format(pStcPedLoc.Num_Apol_Anterior, "0000000000") & ".")
    '        End If
    '        lItemRenova = lItemRenova & ";" & Format(pStcPedLoc.Num_Item, "0000") & ";" & pStcPedLoc.Num_Apol_Anterior & ";" & pStcPedLoc.Num_Item_Anterior & ";"
    '    End If


    'Ficha 901419
    'A cobertura 62 n�o est� relacionada com a cobertura 57
    'Ao contratar a 58 tem que contratar a 57 ou 124 ou 125
    M46V999.GLB_sinal_concess = False

    If val(pStcPedLoc.Cod_Clasf) = 22 And val(pStcPedLoc.Cod_Compl_Clasf) = 13 And pStcPedLoc.ObjPai.Cod_Produto >= 936 Then
        M46V999.GLB_sinal_concess = True

        For Each lstcCobert In pStcPedLoc.ColecaoPedCobert

            'Ficha 1000268 - Quando concession�ria, permitir cobertura 22 no Syas interno
            '            If lstcCobert.Cod_Cobert = 22 Then
            '               If objstcUsuario.TipoUsuario = e_Tipo_Usu_Yasuda And pStcPedLoc.ObjPai.Cod_Ramo = e_CodRamo_Empresarial Then
            '
            '                    If mfRespostaUsuario(E_IDConsist_RcOperacoes, pColAlertas) = E_RespQuest_ND Then
            '                        Call mpAddAlerta(E_IDConsist_RcOperacoes, "Cobertura 22 - RC OPERA��ES n�o se aplica para a classifica��o 022-13 (Concession�ria)." & vbCrLf & "Deseja desconsiderar esta consist�ncia?" & lStrIdItem, E_TipConsist_Quest, pColErros)
            '                    End If
            '                Else
            '                    Call M46V999.gpMenErro(" Consist�ncia do Item " & CStr(lstcCobert.Num_Item) & " -PARA ESTA CLASSIFICA��O, A COBERTURA 00" & CStr(lstcCobert.Cod_Cobert) & " N�O � V�LIDA." & lStrIdItem)
            '               End If
            '            End If

            If lstcCobert.Cod_Cobert = 58 Then
                BOL_sinal_cob57 = False
                BOL_sinal_cob124 = False
                BOL_sinal_cob125 = False

                For Each lstcCobert_aux In pStcPedLoc.ColecaoPedCobert
                    If lstcCobert_aux.Cod_Cobert = 57 Then BOL_sinal_cob57 = True: Exit For
                    If lstcCobert_aux.Cod_Cobert = 124 Then BOL_sinal_cob124 = True: Exit For
                    If lstcCobert_aux.Cod_Cobert = 125 Then BOL_sinal_cob125 = True: Exit For
                Next

                If Not BOL_sinal_cob57 And Not BOL_sinal_cob124 And Not BOL_sinal_cob125 Then
                    Call M46V999.gpMenErro(" Consist�ncia do Item " & CStr(lstcCobert.Num_Item) & " - PARA COBERTURA 00" & CStr(lstcCobert.Cod_Cobert) & ", FALTA A COBERTURA 0057 OU 0124 OU 0125." & lStrIdItem)
                End If
            End If

            GoTo LBL_FIM:
LBL_GRAVAR:
            Call M46V999.gpMenErro(" Consist�ncia do Item " & CStr(lstcCobert.Num_Item) & " -PARA ESTA CLASSIFICA��O, A COBERTURA 00" & CStr(lstcCobert.Cod_Cobert) & " N�O � V�LIDA." & lStrIdItem)
LBL_FIM:
        Next
    End If

    '0800635
    If pStcPedLoc.Num_Cep_Risco = pStcPedLoc.ObjPai.Num_CEP And _
       pStcPedLoc.ObjPai.Cod_Tip_Doc = e_TipDoc_Proposta And _
       (objstcUsuario.TipoAcesso = e_Perfil_CALC Or objstcUsuario.TipoAcesso = e_Perfil_DIG_IST Or _
        objstcUsuario.TipoAcesso = e_Perfil_DIG_Yas) Then
        'Verifico se os demais campos s�o iguais
        lMensagem = "Dados do(s) campo(s) (<Campo>) divergente entre Proponente e Local em Risco."
        lDivEndereco = False
        If pStcPedLoc.Cod_Tip_Logr_Risco <> pStcPedLoc.ObjPai.Cod_Tip_Logr Then
            lAux = "Tipo de logradouro;"
            lDivEndereco = True
        End If
        If pStcPedLoc.Nom_Logr_Risco <> pStcPedLoc.ObjPai.Nom_Logr Then
            lAux = lAux & "Logradouro;"
            lDivEndereco = True
        End If
        If pStcPedLoc.Nom_Bairro_Risco <> pStcPedLoc.ObjPai.Nom_Bairro Then
            lAux = lAux & "Bairro;"
            lDivEndereco = True
        End If
        If Trim(Replace(pStcPedLoc.objstcPedInsp.Num_Logr_Risco & " " & pStcPedLoc.Dsc_Compl_Risco, ".", "")) <> Trim(pStcPedLoc.ObjPai.Dsc_Compl_Logr) Then
            lAux = lAux & "Complemento;"
            lDivEndereco = True
        End If
        If pStcPedLoc.Sig_UF_Risco <> pStcPedLoc.ObjPai.Sig_UF Then
            lAux = lAux & "UF;"
            lDivEndereco = True
        End If
        If lDivEndereco = True Then
            lMensagem = Replace(lMensagem, "<Campo>", lAux)
            Set lPerguntaQuest = PerguntaQuestionario(pColAlertas, E_IDConsist_EndDivergente, pStcPedLoc.Num_Item)
            If lPerguntaQuest Is Nothing Then
                Call mpAddAlerta(E_IDConsist_EndDivergente, "Item " & pStcPedLoc.Num_Item & " -" & lMensagem & " Deseja continuar desta forma?", E_TipConsist_Quest, pColAlertas, pStcPedLoc.Num_Item)
            Else
                Select Case lPerguntaQuest.RespostaQuest
                Case E_RespQuest_ND
                    lPerguntaQuest.RespostaValidada = False
                Case E_RespQuest_Nao
                    lPerguntaQuest.RespostaValidada = False
                    Call M46V999.gpMenErro(lMensagem)
                Case E_RespQuest_Sim
                    lPerguntaQuest.RespostaValidada = True
                End Select
            End If
        Else
            Call mfRemoveQuestionario(E_IDConsist_EndDivergente, pColAlertas, pStcPedLoc.Num_Item)
        End If
    End If

    If pStcPedLoc.Cod_DizerItem <> 4 And pStcPedLoc.Cod_DizerItem <> 10 Then
        For Each lstcPedLocBem In pStcPedLoc.ColecaoPedLocBem
            If lstcPedLocBem.Cod_Bem = 0 Then
                lMsgErr = "FNP9999 - Rela��o de bens: O bem n�mero " & lstcPedLocBem.Cod_Seq & " n�o est� preenchido." & lStrIdItem
                Call M46V999.gpMenErro(lMsgErr)
            End If
            If Trim(lstcPedLocBem.Dsc_Bem) = "" Or lstcPedLocBem.Val_Bem = 0 Then
                If Trim(lstcPedLocBem.Dsc_Bem) = "" Then
                    lMsgErr = "FNP9999 - Rela��o de bens: Favor infomar Descri��o do bem n�mero " & lstcPedLocBem.Cod_Seq & lStrIdItem
                    Call M46V999.gpMenErro(lMsgErr)
                End If
            End If
        Next lstcPedLocBem
    End If

    '***************** Retorno com Sucesso
    mfConsistePreCalcItem = True
    '*****************

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

TrataErro:
    mfConsistePreCalcItem = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err & " - " & ERROR
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lObjPed = Nothing
    Set lstcPedLocBem = Nothing
    Set lstcCobert = Nothing
    Set lstcCobert_Acess = Nothing
    Set lstcCobert_aux = Nothing
    Set cxn = Nothing
    Set obj_rs = Nothing
    Set lstcCobert_aux = Nothing
    Set lPerguntaQuest = Nothing
    Return
End Function

Private Function mfRespostaUsuario(ByVal IDQuest As E_IDConsist, ByRef pColAlertas As Collection, Optional ByVal pNumItem As Long = 0) As E_RespQuest
    Dim lstrIDQuest As String
    Dim lstcConsist As stcConsistPreCalc
    Dim lretorno As E_RespQuest
    Dim i As Integer
    Err.Clear
    On Error GoTo TrataErro
    lstrIDQuest = Format(pNumItem, "000") & "-" & Format(IDQuest, "000")

    If Not pColAlertas Is Nothing Then
        For i = 1 To pColAlertas.Count
            Set lstcConsist = pColAlertas(i)
            If lstrIDQuest = lstcConsist.Class_ID Then
                lretorno = lstcConsist.RespostaQuest
                Call pColAlertas.Remove(i)
                Exit For
            End If
        Next
    Else
        lretorno = E_RespQuest_ND
    End If

    mfRespostaUsuario = lretorno

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

TrataErro:
    lretorno = E_RespQuest_ND
    GoTo SAIDA

FechaObj:
    Set lstcConsist = Nothing
    Return
End Function

Public Function gfCalcular(ByVal pNossoNumero As String, _
                           ByVal pColNormas As Collection, _
                           ByRef pColAlertas As Collection, _
                           ByRef pColErros As Collection, _
                           ByRef pClasse4 As Boolean, _
                           ByVal tagDynaTrace As String) As stcA46V702B
    Dim lStcPed As stcA46V702B
    Dim i As Integer
    Dim pcodigo As String
    Dim pCodigoCadin As String
    Dim lcpf As String
    Dim lobjstcPedLoc As stcA46V708B
    Dim lobjstcPedCobert As stcA46V704B

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfCalcular"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    Call M46V999.gpGraLog(0, "Tag DynaTrace: " & tagDynaTrace & " | KeyDynaTrace " & CStr(lKeyDynaTrace))
    
    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(M46V999.gDataCorteSompo) Then
            If CLng(M46V999.gDataCorteSompo) <= M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If
    
    
    Err.Clear: UltimoErro = ""
    On Error GoTo TrataErro
    Set pColErros = New Collection
    Call M46V999.gpInicializar_gErros

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        Exit Function
    End If

    'Busca o pedido persistido na Base.
    Set lStcPed = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)
    If M46V999.gfPreenchido(UltimoErro) = True Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    'Executa pr�-Consistencias minimas p/ seguir no calculo
    If mfConsistePreCalc(lStcPed, pColAlertas, pColErros, tagDynaTrace) = True Then
        If pColErros.Count > 0 Then
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
            GoTo SAIDA
        Else
            For i = 1 To pColAlertas.Count
                If pColAlertas(i).TipoConsistencia = E_TipConsist_Quest And _
                   pColAlertas(i).RespostaValidada = False Then
                    'H� Questionamentos n�o respondidos
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
                    GoTo SAIDA
                End If
            Next i
        End If
    Else
        GoTo TrataErro
    End If

    If lStcPed.Cod_Dizer = 10 Then
        If lStcPed.gApoliceQuitada = False Then
            If lStcPed.Ind_CotacaoSolic = False Then
                M46V999.gpMenErro ("C�lculo de endosso de cancelamento ser� realizado pela " & lNomeSeguradora & ". Solicitar a cota��o.")
                pClasse4 = True
                GoTo PularPremio
            End If
        End If
    End If

    'Residencial Supremo
    M46V999.gCodProduto = lStcPed.Cod_Produto
    If lStcPed.Ind_CotacaoSolic = False And lStcPed.Cod_Ramo = 113 And lStcPed.Cod_Estip = 30 And M46V999.LiberaPlanoResidencialSupremo(lStcPed.Cod_Estip, lStcPed.ColecaoPedLoc(1).Tip_Emissao) Then
        If lStcPed.ColecaoPedLoc.Count > 20 Then
            M46V999.gpMenErro ("Quantidade de Locais maior que o permitido para o produto. Solicitar a cota��o.")
            pClasse4 = True
            GoTo PularPremio
        End If
    End If

    ' LMI Unico - In�cio
    If lStcPed.Ind_CotacaoSolic = False And lStcPed.Cod_Ramo = 112 And _
        lStcPed.LMI_Unico = 1 And (lStcPed.Cod_Estip = 10 Or lStcPed.Cod_Estip = 11) Then
        
        If lStcPed.ColecaoPedLoc.Count > 50 Then
            M46V999.gpMenErro ("O n�mero m�ximo de itens permitidos para o cotador on-line � de 50 itens. As informa��es digitadas ser�o enviadas para an�lise interna da Cia. Clique em solicitar An�lise.")
            pClasse4 = True
            GoTo PularPremio
        End If
        
        If lStcPed.ColecaoPedLoc.Count > 0 Then
            If lStcPed.ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_RenYasuda Then
                M46V999.gpMenErro ("Renova��es de ap�lices LMI �nico devem ser analisadas internamente. As informa��es digitadas ser�o enviadas para an�lise. Clique em Solicitar An�lise.")
                pClasse4 = True
                GoTo PularPremio
            End If
            If lStcPed.ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_Endosso Then
                M46V999.gpMenErro ("Endosso de ap�lice LMI �nico deve ser analisado internamente. As informa��es digitadas ser�o enviadas para an�lise. Clique em Solicitar An�lise.")
                pClasse4 = True
                GoTo PularPremio
            End If
        End If
    End If
    ' LMI Unico - Fim

    'Anexa as normas recebidas em seus respectivos itens
    Call M46V999.gpGraLog(0, "Quantidade de normas : " & pColNormas.Count)
    If mfCarregaNormas(lStcPed, pColNormas, tagDynaTrace) = False Then GoTo TrataErro

    '-Calcula o Premio
    If mfCalPremios(lStcPed, pColAlertas, pColErros, tagDynaTrace) = False Then
        GoTo TrataErro
    End If


    If M46V999.gQuaErrosDemais = 0 And M46V999.gQuaErrosNormas > 0 And _
       (objstcUsuario.TipoUsuario = e_Tipo_Usu_Corretor Or objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda) And _
       lStcPed.Ind_CotacaoSolic = False Then
        pClasse4 = True
    Else
        pClasse4 = False
        ' CR3923 - Item 8 - Permitir solicitar cota��o sempre que houver algum erro com o texto "Solicitar cota��o"
        For i = 0 To UBound(M46V999.gerros)
            If (objstcUsuario.TipoUsuario = e_Tipo_Usu_Corretor Or objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda) And lStcPed.Ind_CotacaoSolic = False Then
                If InStr(UCase(M46V999.gerros(i)), "SOLICITAR COTA��O") Or InStr(UCase(M46V999.gerros(i)), "SEM ACEITA��O") Or InStr(UCase(M46V999.gerros(i)), "SOLICITAR AN�LISE") Then
                    pClasse4 = True
                End If
                'If InStr(UCase(M46V999.gerros(i)), "SEM ACEITA��O") Then
                '    pClasse4 = False
                '    Exit For
                'End If
            Else
                If (objstcUsuario.TipoUsuario = e_Tipo_Usu_Corretor Or objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda) And lStcPed.Ind_CotacaoSolic = True And M46V999.gEhRetornoSiscota(lStcPed.NOSSO_NUMERO) Then
                    For Each lobjstcPedLoc In lStcPed.ColecaoPedLoc
                        If mfValidaCoberturaClasse4(lobjstcPedLoc) Then
                            If InStr(UCase(M46V999.gerros(i)), "SOLICITAR COTA��O") Or InStr(UCase(M46V999.gerros(i)), "SEM ACEITA��O") Or InStr(UCase(M46V999.gerros(i)), "SOLICITAR AN�LISE") Then
                                pClasse4 = True
                            End If
                            Exit For
                        End If
                    Next
                End If
            End If
        Next i
        ' CR3923 - Item 8 - Fim
    End If
    
    If M46V999.gQuaErrosDemais = 0 And M46V999.gQuaErrosNormas = 0 And lStcPed.Ind_CotacaoSolic = False Then
        For Each lobjstcPedLoc In lStcPed.ColecaoPedLoc
            If lobjstcPedLoc.Tip_Emissao < 200 Then
                For Each lobjstcPedCobert In lobjstcPedLoc.ColecaoPedCobert
                    If lobjstcPedCobert.VAL_PR = 0 Then
                        M46V999.gpMenErro ("Ocorreu uma falha no c�lculo, favor calcular novamente.")
                        GoTo PularPremio
                    End If
                Next
            End If
        Next
    End If
    
PularPremio:
    '***************** Retorno
    If UBound(M46V999.gerros) = 0 Then
        '- Salva os Dados Comuns (incluindo todos os dependentes do pedido)
        If mfSalvarPedido(lStcPed, UltimoErro, tagDynaTrace) = False Then
            GoTo TrataErro
        Else
            i = 1
            Do Until i > pColAlertas.Count
                If pColAlertas(i).TipoConsistencia = E_TipConsist_Quest Then
                    Call pColAlertas.Remove(i)
                Else
                    i = i + 1
                End If
            Loop
        End If
    Else
        Set pColErros = mfArrayToColecao(M46V999.gerros)
    End If
    Set gfCalcular = lStcPed
    '*****************

    pCodigoCadin = ""
    lcpf = ""
    
    'Tarifa 11.38 N�o Validar CADIN em retorno de classe 4
    If Not (lStcPed.Cod_Produto >= 1138 And M46V999.gEhRetornoSiscota(lStcPed.NOSSO_NUMERO) And _
       (lStcPed.gSegmentado Or lStcPed.Cod_Estip = 10 Or lStcPed.Cod_Estip = 11 Or lStcPed.Cod_Estip = 20)) Then
            If lStcPed.Cod_Tip_Doc = e_TipDoc_Proposta And lStcPed.Cod_Produto >= 1082 Then
                lcpf = lStcPed.Num_Cgc_Cpf
                
                If lStcPed.objstcCtrlEmis.Cod_Iden_Pessoa = e_TipPessoa_Fisica Then
                    lcpf = Mid(lcpf, 1, Len(Trim(lcpf)) - 2)
                Else
                    lcpf = Mid(lcpf, 1, Len(Trim(lcpf)) - 6)
                End If
                
                If M46V999.gfObtPesquisaCadin(lcpf, lStcPed.objstcCtrlEmis.Cod_Iden_Pessoa, pCodigoCadin, UltimoErro) = False Then
                    If UltimoErro <> "" Then
                        UltimoErro = "N�o foi possivel prosseguir com o c�lculo. Por favor tente mais tarde!"
                        GoTo TrataErro
                    End If
                    
                                'tarifa 1094
                    UltimoErro = "Em aten��o ao pedido de cota��o, vimos agradecer a oportunidade, mas ficamos impossibilitados de apresentar uma proposta considerando nossa atual pol�tica de aceita��o."
                    GoTo TrataCadin
                End If
            End If
    End If

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    Err.Clear
    GoSub FechaObj
    Exit Function

TrataErro:
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err & " - " & ERROR
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA

TrataCadin:
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lStcPed = Nothing
    Set lobjstcPedLoc = Nothing
    Set lobjstcPedCobert = Nothing
    Return
End Function

Private Function mfCalPremios(ByRef pStcPed As stcA46V702B, _
                                ByRef pColAlertas As Collection, _
                                ByRef pColErros As Collection, _
                                tagDynaTrace As String) As Boolean
'Fun��o: calcula pr�mios.
'********************Vari�veis padr�o*************************************
    Dim lCkPoint As String    'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lCkStr As String    'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
'lCkPoint. (posi��o das variaveis em ordem de declara��o)
'********************Vari�veis Declaradas para o processamento************
    Dim lI As Long    'Utilizada no for...loop
    Dim lJuros As String    'Juros.
    Dim lMensagem As String    'Mensagem.
    Dim lParcela1 As Double    'Quantidade de parcelas para antecipada.
    Dim lParcelaN As Double    'Quantidade de parcelas para postecipada.
    Dim lPrefixo As String    'Prefixo de mensagem.
    Dim lUltData As Double    'Data da �ltima atualiza��o.
    Dim lQutdeParcelas As Integer     'Quantidade de parcelas.
    Dim lobjstcItem As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)

    Dim clsA46V102 As clsA46V102A    'DLL : consiste dados comuns (n�vel 2).
    Dim lObjstcVenc As stcA46V713B    'Objeto da classe de estrutura stcA46V713B (TAB_PED_Venc)

    Dim lParcelasGrid As Integer    ' Numero de parcelas exibidas no grid
    Dim lColecaoTemp As Collection    'Cole��o tempor�ria de itens.
    Dim lretorno As Integer    'Retorno da chamada da DLL.

    Dim lTotPr As Single
    Dim lTotPr_Resg As Single
    Dim lTotPr_Tarif As Single
    Dim lobjstcCobert As stcA46V704B
    Dim mQtdInsp As Integer
    Dim mComUltData As Double    'Dados comuns - �ltima atualiza��o: data.
    Dim mComUltUsuario As String  'Dados comuns - �ltima atualiza��o: usu�rio.
    Dim objPedido As clsA46V702A    'Objeto da classe clsA46V702A  (TAB_PED)

    Dim lMsgErr As String
    Dim lobjPedLMI As clsA46V732A
    Dim lobjCobert As clsA46V704A

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    Dim lExibeMsgSupremo As Boolean
    Dim blnExisteSupremo As Boolean
    Dim blnValoresSupremo As Boolean
    
    methodName = "mfCalPremios"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    mfCalPremios = True    '***** So sai False se der Erro de RunTime!!!

    '1. Limpa grade de mensagens e par�metros de sa�da.
    Err = 0
    lCkPoint = ""
    lCkStr = ""
    On Error GoTo ControleErro

    Set lobjCobert = New clsA46V704A
    Set lobjCobert.M46V999 = M46V999

    'Zero os contadores de erros
    M46V999.gQuaErrosNormas = 0         '***VarGlobal
    M46V999.gQuaErrosDemais = 0         '***VarGlobal

    M46V999.gPodeEmitir = True          '***VarGlobal

    lCkPoint = "1"

    '- Verifica quantidade de parcelas.
    If pStcPed.Qtd_Parc = 0 And pStcPed.Cod_Dizer = 0 Then
        pStcPed.Qtd_Parc = 1
    End If

    If objstcUsuario.TipoUsuario = e_Tipo_Usu_Corretor Or objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda Then
        pStcPed.Val_Cust_Emis = pStcPed.getProduto.ProCusApolice
    End If
    '==============

    lCkPoint = "2"

    '- "C�lculo de pr�mios. Consistindo os dados de comum..."
    '5. COnsit�ncia dos dados de comum
    'Posicionar juros

    'RETIREI ISSO POIS ENTENDO QUE J� FOI FEITO
    'VERIFICAR E SE NECESSARIO VOLTAR

    'If pStcPed.StatusCotSiscota <> enum_StatusTransm_Aprovada Then
    '    Call mpAtualizaJurosNovo(pStcPed)
    'End If

    lCkPoint = "5"
    If UBound(M46V999.gerros) > 0 Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erros anteriores Checkpoint 5", 0, Class_Name, methodName)
        GoTo SAIDA
    End If

    'CR3334 - Replicar LMI
    Call M46V103.mfAtualizarFranquiasLMI(pStcPed)
    'CR3334 - Fim

    '7. Processa item.
    'Vou copiar a cole��o porque durante o c�lculo os itens sofrem altera��o
    'e s�o removidos e adicionados novamente na cole��o, ent�o para que n�o aconte�a
    'de reprocessar o mesmo item, ent�o � feita a c�pia da cole��o.

    'VERIFICAR AS REGRAS DE CONSISTENCIAS PARA CALCULAR O BATENTE
    If pStcPed.Num_Apol_An > 0 And (pStcPed.Cod_Ramo = e_CodRamo_Residencial Or pStcPed.Cod_Ramo = e_CodRamo_Empresarial) _
        And pStcPed.Cod_Dizer = 0 Then
        If M46V103.ConsisteRegraBatente(pStcPed, 0, True) Then
            pStcPed.CalBatente = True
        Else
            pStcPed.CalBatente = False
        End If
    End If

    lCkPoint = "6"
    'Ficha 1000304
    mQtdInsp = 0
    M46V111.gCod_DizerTexto = "ALTERA��ES QUE GERARAM ENDOSSO:" & vbCrLf & vbCrLf
    Set lColecaoTemp = pStcPed.ColecaoPedLoc
    For Each lobjstcItem In lColecaoTemp
        'Ficha 1000304
        mQtdInsp = mQtdInsp + 1

        lCkPoint = "7"
        If lobjstcItem.Cod_Sit_item = e_SitItem_NaoCalc Then
            If mfCalPreItem(lobjstcItem, pColAlertas, tagDynaTrace) = False Then
                If UBound(M46V999.gerros()) = 0 Then
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erros anteriores Checkpoint 7", 0, Class_Name, methodName)
                    GoTo SAIDA
                End If
            End If
        End If
        'i) Posiciona no pr�ximo item.
    Next lobjstcItem
    lCkPoint = "11"





'    If pStcped.Cod_Dizer <> 0 Then
'        Dim clsA46V101A As clsA46V101A
'        Set clsA46V101A = New clsA46V101A
'        Set clsA46V101A.M46V999 = M46V999
'        If pStcped.Cod_Dizer = 10 Then
'            M46V111.gCod_DizerTexto = M46V111.gCod_DizerTexto & "CANCELAMENTO DA APOLICE"
'        End If
'
'        If M46V111.gCod_DizerTexto = ("ALTERA��ES QUE GERARAM ENDOSSO:" & vbCrLf & vbCrLf) Then
'            M46V111.gCod_DizerTexto = "NENHUMA ALTERACAO REALIZADA"
'        End If
'
'        If Not clsA46V101A.mfSalvarTextoEndosso(pStcped, M46V111.gCod_DizerTexto, lMensagem, objstcUsuario.PropTabela, tagDynaTrace) Then
'            Call M46V999.gpMenErro(lMensagem)
'            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
'            Exit Function
'        End If
'    End If


    '9. Se houve erro, encerra processo.
    If UBound(M46V999.gerros()) > 0 Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erros anteriores Checkpoint 9", 0, Class_Name, methodName)
        GoTo SAIDA
    End If


    '10. Efetua totaliza��o.
    '- "Aguarde... Efetuando totaliza��o..."
    'a) Abre o form T28V123A para c�lculo do fracionamento.

    lCkPoint = "13"
    '11. Efetua fracionamento.
    '-"Aguarde... Efetuando fracionamento..."

    'Demostrativo de Pr�mio (Call frmT46V124A.gpInstanciarForm(pstcPed))
    'PArcelamento e Juros!!!!
    lCkPoint = "14"

    '13. Atualiza registro de dados comuns com dados calculados.
    'a) Posiciona dados.
    lUltData = M46V999.gfForData(3)

    lCkPoint = "15"
    'c) atualiza��o dos dados
    '===1) Situa��o do documento
    pStcPed.Cod_Sit_Ped = e_SitPedido_ComCalculo
    'AQUI
    '===2) Custo de emiss�o.
    'Estrutura��o tarifaria
    pStcPed.Val_Cust_Emis = 0 'pStcPed.getProduto.ProCusApolice

    'Ficha 900046
    If M46V999.gfVerificaParcelaAntec(pStcPed) = False Then 'Ajuste realizado pois estava zerando o valor antecipado indevidamente
    pStcPed.Val_Pr_Antec = 0    'Ajuste solicitado por Thais para transmiss�o Web - Off-Line - 19/10/2012
    End If
    If (pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1) And (pStcPed.Num_Cta_Bas_Cor = 0 And pStcPed.Cod_Agen_Bas_Cor = 0) Then
        If objstcUsuario.TipoAcesso <> DIGITACAO Then
            pStcPed.Val_Pr_Antec = lParcela1
        End If
    End If

    '===16) Data e usu�rio da �ltima altera��o.
    'pstcPed.VAL_PR = frmT46V123A.flxResValores.TextMatrix(0, 1)

    'ATUALIZA OS PREMIOS DE ACORDO COM A REGRA DE BATENTE
    'Residencial Supremo
        lExibeMsgSupremo = False
        blnExisteSupremo = False
        blnValoresSupremo = False
        
        For Each lobjstcItem In pStcPed.ColecaoPedLoc
            If M46V999.ContemResidencialSupremo(, lobjstcItem) And blnExisteSupremo = False Then
                blnExisteSupremo = True
                Exit For
            End If
        Next lobjstcItem
                
        For Each lobjstcItem In pStcPed.ColecaoPedLoc
            If lobjstcItem.Val_Vr >= 2000000 And lobjstcItem.Tip_Emissao = e_TipEmissao_Endosso Then
                blnValoresSupremo = True
                Exit For
            End If
        Next lobjstcItem
                
        For Each lobjstcItem In pStcPed.ColecaoPedLoc
        If pStcPed.CalBatente = True And lobjstcItem.ColecaoRegraBatente.Count > 0 Then
            For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
                If lobjstcCobert.VAL_PR = 0 Then
                    lMensagem = "Cobertura " & lobjstcCobert.Cod_Cobert & " - " & lobjstcCobert.Nom_Cobertura & " com pr�mio zerado."
                    Call M46V999.gpMenErro(lMensagem)
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                    Exit Function
                End If
                
                If Not lobjCobert.mfAlterarValoresPrecificacao(lobjstcCobert, "", "") Then
                    GoTo ControleErro
                End If
            Next lobjstcCobert
        End If
                                            
            M46V999.gCodProduto = pStcPed.Cod_Produto
        
        'Rsidencial Supremo
        If blnValoresSupremo = False And lobjstcItem.Tip_Emissao <> e_TipEmissao_Endosso Then
            If lobjstcItem.Cod_Ramo = 113 And M46V999.LiberaPlanoResidencialSupremo(lobjstcItem.Cod_Plano, lobjstcItem.Tip_Emissao) Then
                For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
                    If lobjstcCobert.Cod_Cobert = 1 And lobjstcItem.Cod_Ramo = 113 And lobjstcItem.Tab_Clasf = 10 And blnExisteSupremo = False And lobjstcItem.Tip_Emissao <> e_TipEmissao_Endosso Then
                        If lobjstcItem.Ind_VerbaSep Then
                            If M46V999.mfRetorna_ValIS_VerbasSeparadas(lobjstcCobert) >= 2000000 Then
                                lExibeMsgSupremo = True
                            End If
                        Else
                            If lobjstcCobert.Val_IS >= 2000000 Then
                                lExibeMsgSupremo = True
                            End If
                        End If
                        
                        If lExibeMsgSupremo Then
                            lMensagem = lMensagem & "Item - " & lobjstcItem.Num_Item & " - Limite permitido para o produto Residencial valor menor que R$ 2.000.000,00." & "\n"
                        End If
                    End If
                    
                    If lobjstcCobert.Cod_Cobert = 1 And lobjstcItem.Cod_Ramo = 113 And lobjstcItem.Tab_Clasf = 30 And blnExisteSupremo = False Then
                        If lobjstcItem.Ind_VerbaSep Then
                            If M46V999.mfRetorna_ValIS_VerbasSeparadas(lobjstcCobert) < 2000000 Then
                                lMensagem = lMensagem & "Item - " & lobjstcItem.Num_Item & " - Limite permitido para o produto Residencial Supremo valor maior ou igual a R$ 2.000.000,00 at� R$ 80.000.000,00." & "\n"
                            End If
                            
                            If M46V999.mfRetorna_ValIS_VerbasSeparadas(lobjstcCobert) > 80000000 Then
                                lMensagem = lMensagem & "Item - " & lobjstcItem.Num_Item & " - Limite acima do permitido para este produto." & "\n"
                            End If
                        Else
                            If lobjstcCobert.Val_IS < 2000000 Then
                                lMensagem = lMensagem & "Item - " & lobjstcItem.Num_Item & " - Limite permitido para o produto Residencial Supremo valor maior ou igual a R$ 2.000.000,00 at� R$ 80.000.000,00." & "\n"
                            End If
                            
                            If lobjstcCobert.Val_IS > 80000000 Then
                                lMensagem = lMensagem & "Item - " & lobjstcItem.Num_Item & " - Limite acima do permitido para este produto." & "\n"
                            End If
                        End If
                    End If
                Next lobjstcCobert
            End If
        End If
        Next lobjstcItem

    If Trim(lMensagem) <> "" Then
        Call M46V999.gpMenErro(lMensagem)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    lTotPr = pStcPed.gValPrLiqTotal
    lTotPr_Resg = 0
    lTotPr_Tarif = 0

    'RESSEGURO
    If pStcPed.Dat_Ini_Vig >= 20090501 Then
        lTotPr_Resg = lTotPr
        lTotPr_Tarif = lTotPr
    Else
        For Each lobjstcItem In pStcPed.ColecaoPedLoc
            For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
                lTotPr_Resg = lTotPr_Resg + lobjstcCobert.Val_Pr_Resg
                lTotPr_Tarif = lTotPr_Tarif + lobjstcCobert.Val_Pr_Tarif
            Next lobjstcCobert
        Next lobjstcItem
        
        If M46V999.gPermitePrInformado = False Then
        If lTotPr < pStcPed.getProduto.ProPremioMinimo Then
            lTotPr = pStcPed.getProduto.ProPremioMinimo
        End If
        If lTotPr_Resg < pStcPed.getProduto.ProPremioMinimo Then
            lTotPr_Resg = pStcPed.getProduto.ProPremioMinimo
        End If
        If lTotPr_Tarif < pStcPed.getProduto.ProPremioMinimo Then
            lTotPr_Tarif = pStcPed.getProduto.ProPremioMinimo
        End If
    End If
    End If
    
    'RESSEGURO
    If pStcPed.Dat_Ini_Vig >= 20090501 Then
        pStcPed.VAL_PR = lTotPr
        pStcPed.Val_Pr_Ref_Resg = lTotPr
        pStcPed.Val_Pr_Ref_Tarif = lTotPr
    Else
        pStcPed.VAL_PR = lTotPr
        pStcPed.Val_Pr_Ref_Resg = lTotPr_Resg
        pStcPed.Val_Pr_Ref_Tarif = lTotPr_Tarif
    End If

    mComUltData = pStcPed.Dat_Liber
    mComUltUsuario = pStcPed.Cod_User_Liber

    pStcPed.Dat_Liber = Left(lUltData, 8)
    pStcPed.Hor_Liber = Mid(lUltData, 9, 4)

    'Campo inclu�do por solicita��o da Sandra 03/12/2010
    pStcPed.Ind_CondGerais = 1

    lCkPoint = "15A"
    
    If mfSalvarTextoEspecPadrao(pStcPed, lMsgErr, True, tagDynaTrace) = False Then
        GoTo ControleErro
    End If
    
    'Alterado o lugar de salvar o texto de endosso para beneficiario pois depende de dados na tab_ped_espec
    If pStcPed.Cod_Dizer <> 0 Then
        Dim clsA46V101A As clsA46V101A
        Set clsA46V101A = New clsA46V101A
        Set clsA46V101A.M46V999 = M46V999
        If pStcPed.Cod_Dizer = 10 Then
            M46V111.gCod_DizerTexto = M46V111.gCod_DizerTexto & "CANCELAMENTO DA APOLICE"
        End If

        If M46V111.gCod_DizerTexto = ("ALTERA��ES QUE GERARAM ENDOSSO:" & vbCrLf & vbCrLf) Then
            M46V111.gCod_DizerTexto = "NENHUMA ALTERACAO REALIZADA"
        End If

        If Not clsA46V101A.mfSalvarTextoEndosso(pStcPed, M46V111.gCod_DizerTexto, lMensagem, objstcUsuario.PropTabela, tagDynaTrace) Then
            Call M46V999.gpMenErro(lMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
            Exit Function
        End If
    End If
        
    If M46V999.mfSalvarNovosTextoCoberturas(pStcPed, lMsgErr, objstcUsuario.PropTabela) = False Then
        GoTo ControleErro
    End If
    
    If mfCalcFracionamento(pStcPed, lMsgErr, tagDynaTrace) = False Then
        GoTo ControleErro
    End If

    '12. Consiste dados comuns - n�vel 3.
    If pStcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        If pStcPed.Cod_Reg = 120 Or pStcPed.Cod_Reg = 107 Or pStcPed.Cod_Reg = 106 Or pStcPed.Cod_Reg = 129 Then
            If pStcPed.ColecaoPedVenc.Count > 0 Then
                Dim lobjVenc As clsA46V713A    'Objeto da classe clsA46V713A (TAB_PED_VENC)
                Set lobjVenc = New clsA46V713A
                Set lobjVenc.M46V999 = M46V999
                Set lObjstcVenc = New stcA46V713B
                lObjstcVenc.NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
                If lobjVenc.mfExcluir(lObjstcVenc, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                    lMsgErr = lobjVenc.UltimoErro
                    GoTo ControleErro
                End If
                Set pStcPed.ColecaoPedVenc = lobjVenc.Colecao
            End If
            pStcPed.Qtd_Parc = 0
            pStcPed.Cod_Agen_Bas_Cor = 0
            pStcPed.Num_Cta_Bas_Cor = 0
            pStcPed.Cod_Bco_Cobr = 215
            pStcPed.objstcCtrlEmis.Num_Bloq = 0
            pStcPed.Val_Pr_Antec = 0
            pStcPed.Tip_Cust_Emis = 1
        Else
            Set clsA46V102 = New clsA46V102A    'DLL : consiste dados comuns (n�vel 3).
            Set clsA46V102.M46V999 = M46V999

            With clsA46V102
                'a) Prepara dados para consist�ncia de comum n�vel 3.
                If pStcPed.ColecaoPedVenc.Count > 0 Then
                    Set lObjstcVenc = pStcPed.ColecaoPedVenc.Item(1)
                    .Pro_CarDatPriParcela = lObjstcVenc.DAT_VENC
                End If

                'Proposta: forma de pagamento.
                '1 = carn� e 2 = d�bito em conta.
                If pStcPed.Num_Cta_Bas_Cor <> 0 Then
                    If pStcPed.Num_Cta_Bas_Cor = 9999999 And pStcPed.Cod_Agen_Bas_Cor = 9999 And pStcPed.Cod_Bco_Cobr = 999 Then
                        .Pro_ForPagamento = 4
                    Else
                        .Pro_ForPagamento = 2
                    End If
                Else
                    .Pro_ForPagamento = 1
                End If

                'Ficha 900046
                If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
                    If pStcPed.objstcCtrlEmis.Num_Bloq <> 0 Then
                        .Pro_PagAntecipado = 1
                    Else
                        .Pro_PagAntecipado = 2
                    End If
                Else
                    .Pro_PagAntecipado = 0
                End If
                .Pro_QtdParcelas = pStcPed.Qtd_Parc
                .Pro_CodSegurado = pStcPed.Cod_Segurado

                'b) Executa consist�ncia.
                On Error Resume Next

                Call .VER_COMUM_NIVEL_3(pStcPed)

                'c) Se houve erro, encerra processo.
                If UBound(M46V999.gerros()) > 0 Then
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erros anteriores VerComum Nivel 3", 0, Class_Name, methodName)
                    GoTo SAIDA
                End If
            End With
        End If
    End If

    '2 - Valida as parcelas
    Call mpValidaParcela(pStcPed, pColAlertas, pColErros)

    '===17) Altera o registro.
    'a)
    Set objPedido = New clsA46V702A
    Set objPedido.M46V999 = M46V999



    If Not objPedido.mfAlterar(pStcPed, objstcUsuario.PropTabela, "", mComUltData, mComUltUsuario) Then
        Call M46V999.gpMenErro(objPedido.UltimoErro)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erros anteriores mfAlterar objPedido", 0, Class_Name, methodName)
        GoTo SAIDA
    End If

    'b) Posiciona informa��es de data e usu�rio da �ltima atualiza��o.
    mComUltData = Left(lUltData, 8)
    mComUltUsuario = pStcPed.Cod_User_Liber
    lCkPoint = "16"
    '13. Calcula e grava c�digo de verifica��o da unicidade da proposta.
    'a) Calcula c�digo de verifica��o.
    If pStcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        If M46V999.gfCalCodUnicidade_V1(pStcPed, lMensagem) = False Then
            M46V999.gpMenErro (lMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 0, Class_Name, methodName)
            GoTo SAIDA
        End If
        pStcPed.Dat_Liber = mComUltData
        If Not objPedido.mfAlterar(pStcPed, objstcUsuario.PropTabela, "", mComUltData, mComUltUsuario) Then
            lMensagem = objPedido.UltimoErro
            Call M46V999.gpMenErro(lMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 0, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If

    If pStcPed.LMI_Unico = 1 Then
        Set lobjPedLMI = New clsA46V732A
        Set lobjPedLMI.M46V999 = M46V999
        If Not lobjPedLMI.mfGerarDadosLMI(pStcPed, objstcUsuario.PropTabela, "") Then
            GoTo ControleErro
        End If
    End If

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    'Obtem o inicio da mensagem de erro
    lMensagem = "mfCalPremios ERRO - " & Err & " " & ERROR & _
                Chr(13) & "CheckPoint " & lCkPoint & " - "
    lCkStr = " | " & ""
    'Completa a mensagem de erro
    lMensagem = lMensagem & lCkStr

    UltimoErro = lMensagem
    
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    
    mfCalPremios = False
    GoTo SAIDA

FechaObj:
    Set lobjstcItem = Nothing
    Set clsA46V102 = Nothing
    Set lObjstcVenc = Nothing
    Set lColecaoTemp = Nothing
    Set lobjstcCobert = Nothing
    Set objPedido = Nothing
    Set lobjPedLMI = Nothing
    Set lobjCobert = Nothing
    Return

End Function

Private Function mfCalPreItem(pStcPedLoc As stcA46V708B, ByRef pColAlertas As Collection, tagDynaTrace As String) As Boolean
'Fun��o: Calcula pr�mios do item.N
'Par�metros de entrada...
'                        pAtuStsItem.........Indica se o status do pr�mio pode ou n�o ser atualizado.
'Fun��o: calcula pr�mios.
'********************Vari�veis padr�o*************************************
    Dim lCkPoint As String    'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lCkStr As String    'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
'lCkPoint. (posi��o das variaveis em ordem de declara��o)
'********************Vari�veis Declaradas para o processamento************
    Dim lAguarde As String    'Mensagem de aguarde.
    Dim lbdMulti As ADODB.Connection    'Conex�o do Banco local das TAB_PED tempor�rias
    Dim lPosicaoItem As Integer      'Posi��o na cole��o.
    Dim lErro As Variant    'Erros da DLL de c�lculo de pr�mio.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lMensagem As String    'Mensagem.
    Dim lMsgInsp As String    'Mensagem de inspe��o
    Dim lPrefixo As String    'Prefixo de mensagem.
    Dim lretorno As Integer    'Retorno.
    Dim lrst As clsYasRecordSet    'Registro.
    Dim lSelect As String    'Select
    Dim lUltData As Double    'Data da �ltima atualiza��o.
    Dim lstcCobert As stcA46V704B    'Objeto da classe (stc) de coberturas.
    Dim lstcCobert_Acess As stcA46V704B      'Objeto da classe (stc) de coberturas.
    Dim lobjCobert As clsA46V704A    'Objeto da classe de coberturas.
    Dim lobjPedObs As clsA46V714A    'Objeto da classe de observa��o.
    Dim lobjPedLoc As clsA46V708A    'Objeto da classe de itens.
    Dim lBonus As Long
    Dim lColecaoTemp As Collection
    Dim AtualizaProgres As Boolean
    Dim lCod_Aceit As Integer

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfCalPreItem"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

    '1100589 - Tarifa 1014 e 1015 Julho 2011
    Dim dblTotalIsRC As Double: dblTotalIsRC = 0
    Dim lLimComissMax As Double
    Dim lTemCobertBicicleta As Boolean
    Dim i As Integer
    Dim X As Integer

    On Error GoTo ControleErro

    'Se selecionada cobs. 58 ou 62 E classif. "Concession�ria", necessaria cobert. 57.
    '    'NOVA ADI��O: Quando concession�ria, impedir cobertura 22.

    '1001146 - Tarifa 1000 e 1001 Dezembro 2010
    For Each lstcCobert In pStcPedLoc.ColecaoPedCobert
        If pStcPedLoc.ObjPai.Cod_Produto > 999 Then
            If (pStcPedLoc.ObjPai.Cod_Ramo = e_CodRamo_Empresarial And lstcCobert.Cod_Cobert = 94) Or _
               (pStcPedLoc.ObjPai.Cod_Ramo = e_CodRamo_Residencial And (lstcCobert.Cod_Cobert = 27 Or _
                                                                        lstcCobert.Cod_Cobert = 29 Or _
                                                                        lstcCobert.Cod_Cobert = 30 Or _
                                                                        lstcCobert.Cod_Cobert = 31)) Then
                pStcPedLoc.Cob_Bens = True
                pStcPedLoc.Dig_Bens = (pStcPedLoc.ColecaoPedLocBem.Count > 0)
            End If
        End If

        'valida obrigatoriedade de informar quantidade de empregados
        If pStcPedLoc.ObjPai.Cod_Produto > 1027 And _
           pStcPedLoc.ObjPai.Cod_Ramo = e_CodRamo_Residencial And _
           pStcPedLoc.Qtd_Empreg = 0 And lstcCobert.Cod_Cobert = 21 Then
            Call M46V999.gpMenErro(" Consist�ncia do Item " & CStr(pStcPedLoc.Num_Item) & " Obrigat�rio informar a Quantidade de Empregados.")
        End If
    Next

    'Comparar soma dos Bens com o LMG - Syas Roubo para plano AllRisks
    If pStcPedLoc.ObjPai.Cod_Ramo = e_CodRamo_Roubos Then

        'Ficha 1000308 - Permitir itens com planos diferentes no produto Roubo (c�lculo coletivo)
        Call mfConsTotBens(Format(pStcPedLoc.Val_Vr, "standard"), 0, pStcPedLoc, tagDynaTrace)
        If pStcPedLoc.ColecaoPedLocBem.Count = 0 Then

            If pStcPedLoc.Cod_Plano = 2 Then
                For Each lstcCobert In pStcPedLoc.ColecaoPedCobert
                    If lstcCobert.Cod_Cobert = 5 Then
                        Call M46V999.gpMenErro(" Consist�ncia do Item " & CStr(lstcCobert.Num_Item) & " Obrigat�rio incluir a rela��o de objetos com seus respectivos valores unit�rios.")
                        Exit For
                    End If
                Next lstcCobert
            End If
        End If

    End If

    ' CR3923 - Item 6 - Bloquear constru��o mista/inferior - Inicio
    If pStcPedLoc.ObjPai.Cod_Produto < 1080 Then
    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        If pStcPedLoc.Cod_Constr = e_CodContr_Inferior Or pStcPedLoc.Cod_Constr = e_CodContr_Mista Then
            If pStcPedLoc.Cod_Ramo = e_CodRamo_Empresarial Then
                Call M46V999.gpMenErro("Tipo de constru��o Mista/Inferior sem aceita��o.")
            End If
            If pStcPedLoc.Cod_Ramo = e_CodRamo_Residencial Then
                If (pStcPedLoc.Cod_Plano >= 2001 And pStcPedLoc.Cod_Plano <= 2020) Or _
                   (pStcPedLoc.Cod_Plano >= 2042 And pStcPedLoc.Cod_Plano <= 2049) Or _
                   (pStcPedLoc.Cod_Plano >= 2051 And pStcPedLoc.Cod_Plano <= 2060) Or _
                   (pStcPedLoc.Cod_Plano >= 2121 And pStcPedLoc.Cod_Plano <= 2141) Or _
                   (pStcPedLoc.Cod_Plano >= 2179 And pStcPedLoc.Cod_Plano <= 2188) Or _
                   (pStcPedLoc.Cod_Plano >= 2157 And pStcPedLoc.Cod_Plano <= 2167) Or _
                   (pStcPedLoc.Cod_Plano >= 2169 And pStcPedLoc.Cod_Plano <= 2178) Or _
                   (pStcPedLoc.Cod_Plano >= 2199 And pStcPedLoc.Cod_Plano <= 2208) Or _
                   (pStcPedLoc.Cod_Plano >= 2219 And pStcPedLoc.Cod_Plano <= 2238) Or _
                   (pStcPedLoc.Cod_Plano >= 2240 And pStcPedLoc.Cod_Plano <= 2245) Or _
                   (pStcPedLoc.Cod_Plano >= 2250 And pStcPedLoc.Cod_Plano <= 2263) Or _
                   (pStcPedLoc.Cod_Plano >= 2265 And pStcPedLoc.Cod_Plano <= 2271) Or _
                   (pStcPedLoc.Cod_Plano >= 2316 And pStcPedLoc.Cod_Plano <= 2325) Or _
                   (pStcPedLoc.Cod_Plano >= 2326 And pStcPedLoc.Cod_Plano <= 2349) Or _
                   (pStcPedLoc.Cod_Plano >= 2350 And pStcPedLoc.Cod_Plano <= 2373) Or _
                   (pStcPedLoc.Cod_Plano >= 2298 And pStcPedLoc.Cod_Plano <= 2303) Then
                    Call M46V999.gpMenErro("Tipo de constru��o Mista/Inferior sem aceita��o.")
                End If
            End If

        End If
    End If
    End If
    ' CR3923 - Item 6 - Bloquear constru��o mista/inferior - Fim
    mfCalPreItem = False
    'Calcula
    'a) Prepara controles de mensagens.
    lAguarde = "C�lculo de pr�mios. Item " & pStcPedLoc.Num_Item & ". " & vbLf
    lPrefixo = "Item " & pStcPedLoc.Num_Item & " - "


    'b) Calcula pr�mios e salva registro com resultado do c�lculo.
    Set lobjPedLoc = New clsA46V708A
    Set lobjPedLoc.M46V999 = M46V999

    ''Resseguro facultativo
    lCkPoint = "8"
    'c) Consiste item - n�vel 2.
    If pStcPedLoc.Cod_DizerItem <> 4 And pStcPedLoc.Cod_DizerItem <> 10 Then
        Call mpVerIteNiv_2(pStcPedLoc, tagDynaTrace)
    End If
    lCkPoint = "9"
    'd) Consiste item - normas
    lCkPoint = "10"

    '2018-07-25: Verifica se ha Repique
    If (M46V999.GetRepique(pStcPedLoc.NOSSO_NUMERO)) Then
        M46V999.VerificarNormas = False
    End If

    'Verifica Regras de Aceita��o Pontuais
    If M46V999.VerificarNormas Then
        lCod_Aceit = 0
        lMensagem = ""
        lCod_Aceit = M46V103.mfClasseAceit(pStcPedLoc, pStcPedLoc.ObjPai, lMensagem, tagDynaTrace)
        If lCod_Aceit = 3 Then
            ' M46V999.gpMenErro (IIf(lMensagem <> "", lMensagem, "FGE0156 - Proposta sem aceita��o. "))
            M46V999.gpMenErro (IIf(Len(lMensagem) > 5, lMensagem, "FGE0156 - Proposta sem aceita��o. "))
        ElseIf lCod_Aceit = 4 Then
            M46V999.gpMenErro ("FGE0157 - Para essa proposta solicitar cota��o a Unidade de Produtos Diversos.")
        End If
        lCkPoint = "11"
    End If

    '==========--------- Ser� validado pelo COBOL.NEt
    'Aqui Apenas interpreta as mensagens caso existam.
    Call mpVerNormas(pStcPedLoc, lCod_Aceit, lMsgInsp)
    '==========---------

    Call M46V999.gpGraLog(0, "Simony2 Item: " & pStcPedLoc.Num_Item & " Total sinistro " & pStcPedLoc.Ind_Sin_Inspec)

    If lMsgInsp = "" And pStcPedLoc.Ind_Sin_Inspec > 0 Then
        lMsgInsp = "OBRIGATORIO A REALIZACAO DE INSPECAO DO RISCO"
    End If
PulaNorma:

    'NOVO_COTADOR Apenas o IF .StatusCotSiscota <> enum_StatusTransm_EnvioKit ...
    If pStcPedLoc.ObjPai.StatusCotSiscota <> enum_StatusTransm_EnvioKit And pStcPedLoc.ObjPai.StatusCotSiscota <> enum_StatusTransm_Aprovada Then
        '############## PROJETO 10018 - INSPECAO ####################
        pStcPedLoc.Insp_Obr = False
    End If

    If (pStcPedLoc.Insp_Obr Or M46V999.gfPreenchido(lMsgInsp)) And pStcPedLoc.Cod_DizerItem <> 4 And pStcPedLoc.Cod_DizerItem <> 10 Then    'INSPE��O OBRIGAT�RIA
        
        pStcPedLoc.Insp_Obr = True
        
        '*Adicionado por Luiz Alberto em 22/02/2016
        '******************************************
        If pStcPedLoc.ObjPai.Cod_Tip_Doc = e_TipDoc_Proposta And pStcPedLoc.Ind_StatusInspecao <> 1 And pStcPedLoc.Ind_StatusInspecao <> 3 Then

            If M46V999.gfPreenchido(pStcPedLoc.objstcPedInsp.Num_Tel) = False Then
                M46V999.gpMenErro ("FGE0105 -" & " Item " & pStcPedLoc.Num_Item & " - Telefone de contato 1 para inspe��o inv�lido (n�o preenchido).")
            Else
                '1000477 - Telefone de Contato Formatado
                '1 - valida se esta totalmente preenchido
                If Len(pStcPedLoc.objstcPedInsp.Num_Tel) < 13 Then
                    M46V999.gpMenErro ("FGE0105 -" & " Item " & pStcPedLoc.Num_Item & " - Telefone de contato 1 para inspe��o inv�lido (Favor informar o DDD e o n�mero do telefone com 8 d�gitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                Else
                    '2 - valida formado (99)9999-9999
                    If Mid(pStcPedLoc.objstcPedInsp.Num_Tel, 1, 1) <> "(" Or Mid(pStcPedLoc.objstcPedInsp.Num_Tel, 4, 1) <> ")" Or Mid(pStcPedLoc.objstcPedInsp.Num_Tel, 9, 1) <> "-" And Mid(pStcPedLoc.objstcPedInsp.Num_Tel, 10, 1) <> "-" Then
                        M46V999.gpMenErro ("FGE0105 -" & " Item " & pStcPedLoc.Num_Item & " - Telefone de contato 1 para inspe��o inv�lido (Favor informar o DDD e o n�mero do telefone com 8 d�gitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                    Else
                        '3 - n�o permite todos os numeros do telefone iguais
                        If InStr(1, "0000-0000 1111-1111 2222-2222 3333-3333 4444-4444 5555-5555 6666-6666 7777-7777 8888-8888 9999-9999", Mid(pStcPedLoc.objstcPedInsp.Num_Tel, 5, Len(pStcPedLoc.objstcPedInsp.Num_Tel)), vbTextCompare) > 0 Or _
                           (InStr(1, "00000-0000 11111-1111 22222-2222 33333-3333 44444-4444 55555-5555 66666-6666 77777-7777 88888-8888 99999-9999", Mid(pStcPedLoc.objstcPedInsp.Num_Tel, 5, Len(pStcPedLoc.objstcPedInsp.Num_Tel)), vbTextCompare) > 0) Then
                            Call M46V999.gpMenErro("FGE0105 -" & " Item " & pStcPedLoc.Num_Item & " - Telefone de contato 1 para inspe��o inv�lido (Favor informar o DDD e o n�mero do telefone com 8 d�gitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                        End If
                    End If
                End If
            End If

            If M46V999.gfPreenchido(pStcPedLoc.objstcPedInsp.Pess_Cont) = False Or Len(pStcPedLoc.objstcPedInsp.Pess_Cont) <= 3 Then
                M46V999.gpMenErro ("FGE0106 -" & " Item " & pStcPedLoc.Num_Item & " - Contato 1 para inspe��o inv�lido (informar no m�nimo 4 caracteres).")
            End If
            
            If M46V999.gfPreenchido(pStcPedLoc.objstcPedInsp.NUM_TEL_2) = False Then
                M46V999.gpMenErro ("FGE0119 -" & " Item " & pStcPedLoc.Num_Item & " - Telefone de contato 2 para inspe��o inv�lido (n�o preenchido).")
            Else
                '1000477 - Telefone de Contato Formatado
                '1 - valida se esta totalmente preenchido
                If Len(pStcPedLoc.objstcPedInsp.NUM_TEL_2) < 13 Then
                    M46V999.gpMenErro ("FGE0120 -" & " Item " & pStcPedLoc.Num_Item & " - Telefone de contato 2 para inspe��o inv�lido (Favor informar o DDD e o n�mero do telefone com 8 d�gitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                Else
                    '2 - valida formado (99)9999-9999
                    If Mid(pStcPedLoc.objstcPedInsp.NUM_TEL_2, 1, 1) <> "(" Or Mid(pStcPedLoc.objstcPedInsp.NUM_TEL_2, 4, 1) <> ")" Or Mid(pStcPedLoc.objstcPedInsp.NUM_TEL_2, 9, 1) <> "-" And Mid(pStcPedLoc.objstcPedInsp.NUM_TEL_2, 10, 1) <> "-" Then
                        M46V999.gpMenErro ("FGE0121 -" & " Item " & pStcPedLoc.Num_Item & " - Telefone de contato 2 para inspe��o inv�lido (Favor informar o DDD e o n�mero do telefone com 8 d�gitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                    Else
                        '3 - n�o permite todos os numeros do telefone iguais
                        If InStr(1, "0000-0000 1111-1111 2222-2222 3333-3333 4444-4444 5555-5555 6666-6666 7777-7777 8888-8888 9999-9999", Mid(pStcPedLoc.objstcPedInsp.NUM_TEL_2, 5, Len(pStcPedLoc.objstcPedInsp.NUM_TEL_2)), vbTextCompare) > 0 Or _
                           (InStr(1, "00000-0000 11111-1111 22222-2222 33333-3333 44444-4444 55555-5555 66666-6666 77777-7777 88888-8888 99999-9999", Mid(pStcPedLoc.objstcPedInsp.NUM_TEL_2, 5, Len(pStcPedLoc.objstcPedInsp.NUM_TEL_2)), vbTextCompare) > 0) Then
                            Call M46V999.gpMenErro("FGE0122 -" & " Item " & pStcPedLoc.Num_Item & " - Telefone de contato para inspe��o inv�lido (Favor informar o DDD e o n�mero do telefone com 8 d�gitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                        End If
                    End If
                End If
            End If
            
            If M46V999.gfPreenchido(pStcPedLoc.objstcPedInsp.PESS_CONTATO_2) = False Or Len(pStcPedLoc.objstcPedInsp.PESS_CONTATO_2) <= 3 Then
                M46V999.gpMenErro ("FGE0118 -" & " Item " & pStcPedLoc.Num_Item & " - Contato 2 para inspe��o inv�lido (informar no m�nimo 4 caracteres).")
            End If

            If objstcUsuario.TipoUsuario = e_Tipo_Usu_Yasuda Then
                If pStcPedLoc.gNormaInspecao = True Then
                    If mfRespostaUsuario(E_IDConsist_Aviso_InspRisco, pColAlertas, pStcPedLoc.Num_Item) <> E_RespQuest_OK Then
                        Call mpAddAlerta(E_IDConsist_Aviso_InspRisco, "A aceita��o deste seguro est� condicionada a an�lise da inspe��o do risco.", E_TipConsist_Aviso, pColAlertas, pStcPedLoc.Num_Item)
                    End If
                End If
            Else
                If mfRespostaUsuario(E_IDConsist_Aviso_InspRisco, pColAlertas, pStcPedLoc.Num_Item) <> E_RespQuest_OK Then
                    Call mpAddAlerta(E_IDConsist_Aviso_InspRisco, "A aceita��o deste seguro est� condicionada a an�lise da inspe��o do risco - Item:" & pStcPedLoc.Num_Item & ".", E_TipConsist_Aviso, pColAlertas, pStcPedLoc.Num_Item)
                End If
            End If
            Call M46V999.gpGraLog(0, lMensagem & " SIM.")

        '******************************************Fim
        ElseIf pStcPedLoc.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao Then
            pStcPedLoc.Insp_Agend = False
            Call M46V999.gpGraLog(0, lMensagem & " N�O.")
        End If

        lretorno = 0

    End If

    If pStcPedLoc.ColecaoPedCobert.Count = 0 Then
        Call M46V999.gpMenErro("AGE0047 - Nenhuma cobertura foi selecionada.")
    End If

    'Se houve erro nos tipos de produto selecionados, encerra processo.
    If M46V999.gQuaErros <> 0 Or lretorno <> 0 Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "qQuaErros <> 0", 0, Class_Name, methodName)
        GoTo SAIDA
    End If

    If Not M46V999.gfAbrBasDados("P00MULTI2", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdMulti, lMensagem) Then
        Call M46V999.gpMenErro(lMensagem)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 0, Class_Name, methodName)
        GoTo SAIDA
    End If

    AtualizaProgres = True
    If objstcUsuario.TipoUsuario = e_Tipo_Usu_Corretor Or objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda Or _
       (objstcUsuario.TipoUsuario = e_Tipo_Usu_Yasuda) Then
        'And (pstcPedLoc.Per_Desc_Progres = 99)) Then
        lSelect = "Select * from " & objstcUsuario.PropTabela & "TAB_PED_LOC WITH (NOLOCK)"
        lSelect = lSelect & " where Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pStcPedLoc.NOSSO_NUMERO)
        lSelect = lSelect & " and num_item = " & pStcPedLoc.Num_Item
        lSelect = lSelect & " and num_apol = 0"

        'Busca os pr�mios de coberturas
        If Not M46V999.gfObtRegistro(lbdMulti, lSelect, lrst, lMensagem) Then
            Call M46V999.gpMenErro(lMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 0, Class_Name, methodName)
            GoTo SAIDA
        End If

        'Ficha 901673 e 901772
        If pStcPedLoc.Cod_Plano < 2000 Or _
           (pStcPedLoc.Cod_Plano >= 2021 And pStcPedLoc.Cod_Plano <= 2031) Or _
           (pStcPedLoc.Cod_Plano >= 2061 And pStcPedLoc.Cod_Plano <= 2100) Or _
           (pStcPedLoc.Cod_Plano >= 2111 And pStcPedLoc.Cod_Plano <= 2130) Then
            pStcPedLoc.Per_Desc_Progres = lrst("Per_Desc_Progres")
        End If

        If Not lobjPedLoc.mfAlterar(pStcPedLoc, objstcUsuario.PropTabela, "") Then
            lMensagem = lobjPedLoc.UltimoErro
            Call M46V999.gpMenErro(lMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 0, Class_Name, methodName)
            GoTo SAIDA
        End If
    Else
        AtualizaProgres = False
    End If

    If pStcPedLoc.ObjPai.Cod_Produto <= M46V999.TarifaAntiga Then
        lSelect = "Select * from " & objstcUsuario.PropTabela & "Tab_Ped_Cobert WITH (NOLOCK)"
        lSelect = lSelect & " where Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pStcPedLoc.NOSSO_NUMERO)
        lSelect = lSelect & " and num_item = " & pStcPedLoc.Num_Item
        lSelect = lSelect & " and num_apol = 0"
        lSelect = lSelect & " and Val_IS <> 0"

        'Busca os pr�mios de coberturas
        If Not M46V999.gfObtRegistro(lbdMulti, lSelect, lrst, lMensagem) Then
            Call M46V999.gpMenErro(lMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 0, Class_Name, methodName)
            GoTo SAIDA
        End If

TentadeNovo:
        Set pStcPedLoc.ColecaoPedCobert = New Collection

        Set lstcCobert = New stcA46V704B
        Set lstcCobert.M46V999 = M46V999

        Set lobjCobert = New clsA46V704A
        Set lobjCobert.M46V999 = M46V999

        'Set lObjCobert.ConexaoBD = objPedido.ConexaoBD
        lstcCobert.NOSSO_NUMERO = pStcPedLoc.ObjPai.NOSSO_NUMERO
        lstcCobert.Num_Item = pStcPedLoc.Num_Item
        If Not lobjCobert.mfExcluir(lstcCobert, objstcUsuario.PropTabela, "", True) Then
            Call M46V999.gpMenErro(lobjCobert.UltimoErro)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lobjCobert.UltimoErro, 0, Class_Name, methodName)
            GoTo SAIDA
        End If
        Do While Not lrst.EOF
            Set lstcCobert = New stcA46V704B
            Set lstcCobert.M46V999 = M46V999

            lstcCobert.Cod_Cobert = lrst("Cod_Cobert")
            lstcCobert.Cod_Org = lrst("Cod_Org")
            lstcCobert.Cod_Ramo = pStcPedLoc.ObjPai.Cod_Ramo
            lstcCobert.NOSSO_NUMERO = pStcPedLoc.ObjPai.NOSSO_NUMERO
            lstcCobert.Num_Item = pStcPedLoc.Num_Item
            lstcCobert.Num_Ped = pStcPedLoc.Num_Ped
            lstcCobert.Per_Desc_Coml = lrst("Per_Desc_Coml")
            lstcCobert.Per_Desc_Exp = lrst("Per_Desc_Exp")
            lstcCobert.Per_Franq = lrst("Per_Franq")
            lstcCobert.Per_Taxa = lrst("Per_Taxa")
            lstcCobert.Tab_Cobert = lrst("Tab_Cobert")
            lstcCobert.Val_Franq_Max = lrst("Val_Franq_Max")
            lstcCobert.Val_Franq_Min = lrst("Val_Franq_Min")
            lstcCobert.Val_IS = lrst("Val_IS")
            lstcCobert.VAL_PR = lrst("Val_Pr")
            lstcCobert.Val_Pr_Anu = lrst("Val_Pr_Anu")
            lstcCobert.Val_Pr_Resg = lrst("Val_Pr_Resg")
            lstcCobert.Val_Pr_Resg_Anu = lrst("Val_Pr_Resg_Anu")
            lstcCobert.Val_Pr_Tarif = lrst("Val_Pr_Tarif")
            lstcCobert.Val_Pr_Tarif_Anu = lrst("Val_Pr_Tarif_Anu")
            If Not lobjCobert.mfIncluir(lstcCobert, objstcUsuario.PropTabela, "", True, False, tagDynaTrace) Then
                Call M46V999.gpMenErro(lobjCobert.UltimoErro)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lobjCobert.UltimoErro, 0, Class_Name, methodName)
                GoTo SAIDA
            End If
            pStcPedLoc.ColecaoPedCobert.Add lstcCobert, lstcCobert.Class_ID
            lrst.MoveNext
        Loop

    End If
    
    ' Plano Ipiranga - Inicio
    Dim dValISLimite As Double
    Dim codCobertAnt As Integer
    ' Validar se existem valores limite de IS e impedir o calculo se a IS for maior que o valor do limite
    For Each lstcCobert In pStcPedLoc.ColecaoPedCobert
        dValISLimite = 0
        codCobertAnt = M46V103.lCod_Cobert
        M46V103.lCod_Cobert = lstcCobert.Cod_Cobert
        If M46V103.mfAjustesPontuais(e_Tip_Valor_ISLimite_Cobertura, TpoValor, dValISLimite, lMensagem, pStcPedLoc, 0, tagDynaTrace) = False Then
            Call M46V999.gpGraLog(2, lMensagem)
            'Exit Function
        End If
        M46V103.lCod_Cobert = codCobertAnt
        If dValISLimite <> 0 Then
            If lstcCobert.Val_IS > dValISLimite Then
                Call M46V999.gpMenErro("IS da cobertura " & lstcCobert.Cod_Cobert & " � superior ao limite.")
            End If
        End If
    Next
    
    If M46V999.gQuaErros <> 0 Or lretorno <> 0 Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erros de valida��o no mfCalPreItem. Verificar mensagens anteriores", 1, Class_Name, methodName)
        Exit Function
    End If
    ' Plano Ipiranga - Fim
    
    lCkPoint = "50"
    'Antes de Carregar objeto vai calcular novamente
    If pStcPedLoc.ObjPai.Cod_Produto > M46V999.TarifaAntiga Then
        If M46V111.gTipoGrupMulti = e_GrupoRamo_Multi Then
            
            '11.42
            If M46V999.gEhRetornoSiscota(pStcPedLoc.NOSSO_NUMERO) And pStcPedLoc.ObjPai.objstcCtrlEmis.Per_Max_Comis_Corr > 0 Then
                lLimComissMax = pStcPedLoc.ObjPai.objstcCtrlEmis.Per_Max_Comis_Corr
            Else
                lLimComissMax = M46V999.gfObterLimComissMax(pStcPedLoc.Cod_Ramo, pStcPedLoc.ObjPai.Val_IS)
            End If
                
            If pStcPedLoc.ObjPai.ColecaoPedCorr(1).Per_Corr > lLimComissMax And lLimComissMax > 0 Then
                If M46V999.gEhRetornoSiscota(pStcPedLoc.NOSSO_NUMERO) Then
                    lMensagem = "Comiss�o acima do limite. Solicitar Cota��o."
                Else
                    lMensagem = "Comiss�o acima do limite. Favor entrar em contato com seu gerente de produ��o."
                End If
                Call M46V999.gpMenErro(lMensagem)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, lMensagem, 0, Class_Name, methodName)
                GoTo SAIDA
            End If
   
            If Not M46V103.mfCalNewPrecif(pStcPedLoc, AtualizaProgres, lMensagem, objstcUsuario.PropTabela, "", tagDynaTrace) Then
                Call M46V999.gpMenErro(lMensagem)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 0, Class_Name, methodName)
                GoTo SAIDA
            End If
        End If
    End If

    Dim lobjstcPedidoApolice As stcA46V702B
    Dim clsA46V101A As clsA46V101A
    Set clsA46V101A = New clsA46V101A
    Set clsA46V101A.M46V999 = M46V999

    'Nova tarifa 1094 / 1095
    lTemCobertBicicleta = False
'    If pStcPedLoc.Cod_Iden_Tarif >= 1094 Then
'        For i = 1 To pStcPedLoc.ColecaoPedCobert.Count
'            If pStcPedLoc.ColecaoPedCobert(i).Cod_Cobert = 1215 And pStcPedLoc.ColecaoPedCobert(i).Cod_Ramo = 113 Then
'                For x = 1 To pStcPedLoc.ColecaoPedLocBem.Count
'                    If pStcPedLoc.ColecaoPedLocBem(x).Cod_Bem = 21 Then
'                        lTemCobertBicicleta = True
'                        Exit For
'                    End If
'                Next
'
'                If lTemCobertBicicleta Then
'                    If Trim(pStcPedLoc.ColecaoPedLocBem(x).Dsc_Bem) = "" Then
'                        Call M46V999.gpMenErro(" Consist�ncia do Item " & CStr(pStcPedLoc.Num_Item) & ": Favor relacionar a Bicicleta no campo ""Rela��o de Bens""")
'                        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Favor relacionar a Bicicleta no campo ""Rela��o de Bens""", 2, Class_Name, methodName)
'                        GoTo SAIDA
'                    End If
'                End If
'
'                If Not lTemCobertBicicleta Or pStcPedLoc.ColecaoPedLocBem.Count = 0 Then
'                    Call M46V999.gpMenErro(" Consist�ncia do Item " & CStr(pStcPedLoc.Num_Item) & ": Favor relacionar a Bicicleta no campo ""Rela��o de Bens""")
'                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Favor relacionar a Bicicleta no campo ""Rela��o de Bens""", 2, Class_Name, methodName)
'                    GoTo SAIDA
'                End If
'            End If
'        Next i
'    End If


    If pStcPedLoc.Cod_DizerItem <> 0 Then
        Call clsA46V101A.mfCarregarDadosOriginais(pStcPedLoc.ObjPai, lobjstcPedidoApolice)
        If Not clsA46V101A.mfComparareCalcular(pStcPedLoc, pStcPedLoc.ObjPai, lobjstcPedidoApolice, lMensagem, objstcUsuario.PropTabela, tagDynaTrace) Then
            Call M46V999.gpMenErro(lMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 0, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If

    Set clsA46V101A = Nothing

    'Salva resultado do c�lculo.
    'a) Atualiza registro do item.
    lUltData = CDbl(M46V999.gfForData(3))
    'Ficha 0901834 Corre��o de Status (Para Emiss�o Mainframe n�o pode ir 2)
    '    pstcPedLoc.Cod_Sit_item = e_SitItem_Calculado

    '========================== iNUTIL... mosilva
    '    pstcPedLoc.ObjPai.ColecaoPedLoc.Remove pstcPedLoc.Class_ID
    '========================== iNUTIL... mosilva

    'Verifico se deve ser gravada a observa��o de inspe��o.
    Select Case pStcPedLoc.Ind_StatusInspecao
        Case 0
            If M46V999.gfPreenchido(lMsgInsp) = True Then
                If InStr(1, M46V999.gfRetirarCaracterEspecial(pStcPedLoc.objstcPedObs.Observacao, True), M46V999.gfRetirarCaracterEspecial(lMsgInsp, True)) = 0 Then
                    If M46V999.gfPreenchido(pStcPedLoc.objstcPedObs.Observacao) Then
                        If M46V999.gNovaPropostaLMI Then
                            pStcPedLoc.objstcPedObs.Observacao = pStcPedLoc.objstcPedObs.Observacao & vbCrLf & "Item " & Format(pStcPedLoc.Num_Item, "00000") & ": " & M46V999.gfRetirarCaracterEspecial(lMsgInsp, True)
                        Else
                            pStcPedLoc.objstcPedObs.Observacao = M46V999.gfRetirarCaracterEspecial(lMsgInsp, True) & vbCrLf & pStcPedLoc.objstcPedObs.Observacao
                        End If
                    Else
                        If M46V999.gNovaPropostaLMI Then
                            pStcPedLoc.objstcPedObs.Observacao = "Item " & Format(pStcPedLoc.Num_Item, "00000") & ": " & M46V999.gfRetirarCaracterEspecial(lMsgInsp, True)
                        Else
                            pStcPedLoc.objstcPedObs.Observacao = M46V999.gfRetirarCaracterEspecial(lMsgInsp, True)
                        End If
                    End If
                    
                    pStcPedLoc.objstcPedObs.Observacao = Replace(pStcPedLoc.objstcPedObs.Observacao, vbCrLf, "&%")
                End If
            End If
            
        Case 1
            pStcPedLoc.objstcPedObs.Observacao = "Inspecao dispensada"
        Case 2
            pStcPedLoc.objstcPedObs.Observacao = "Inspecao obrigatoria"
        Case 3
            pStcPedLoc.objstcPedObs.Observacao = "Inspecao programada n� " & pStcPedLoc.Num_Insp
    End Select
    
    If M46V999.gfPreenchido(lMsgInsp) = True Then
        Set lobjPedObs = New clsA46V714A
        Set lobjPedObs.M46V999 = M46V999

        If Not lobjPedObs.mfExcluir(pStcPedLoc.objstcPedObs, objstcUsuario.PropTabela, "") Then
            lMensagem = lobjPedObs.UltimoErro
            Call M46V999.gpMenErro(lMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 0, Class_Name, methodName)
            GoTo SAIDA
        End If
        With pStcPedLoc.objstcPedObs
            .NOSSO_NUMERO = pStcPedLoc.NOSSO_NUMERO
            .Num_Item = pStcPedLoc.Num_Item
        End With
        If Not lobjPedObs.mfIncluir(pStcPedLoc.objstcPedObs, objstcUsuario.PropTabela, "", tagDynaTrace) Then
            lMensagem = lobjPedObs.UltimoErro
            Call M46V999.gpMenErro(lMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 0, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If

    '##############
    '  *** S� agendar� a inspe��o ao transmitir a proposta.
    '    '############## PROJETO 10018 - INSPECAO ####################
    '    If M46V999.gfPreenchido(lMsgInsp) = True Then
    '        If objstcUsuario.TipoUsuario = e_Tipo_Usu_Corretor Or objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda Then
    '            'N�o vai agendar inspe��o qdo for CLASSE 4, o c�lculo do Produto j� vai ter gravado o agendamento
    '            'O corretor ir� acessar pelo YasudaNet e confirmar o agendamento
    '            'If pstcPedLoc.ObjPai.Ind_CotacaoSolic = True And _
                 '            '    pstcPedLoc.ObjPai.StatusCotSiscota = enum_StatusTransm_Aprovada Then
    '            '        GoTo NaoAgendar
    '            'End If
    '
    '            If Not M46V111.gfGravarInspecaoProduto(lObjstcPed, lobjstcPedLoc, UltimoErro, 1) Then
    '                If pstcPedLoc.Insp_Agend = True Then
    '                    If mfRespostaUsuario(E_IDConsist_Aviso_ArqTransmInsp, pColAlertas, pstcPedLoc.Num_Item) <> E_RespQuest_OK Then
    '                        Call mpAddAlerta(E_IDConsist_Aviso_ArqTransmInsp, "N�o gerou arquivo de transmiss�o da inspe��o " & pstcPedLoc.ObjPai.Nosso_Numero, E_TipConsist_Aviso, pColAlertas, pstcPedLoc.Num_Item)
    '                    End If
    '                End If
    '            End If
    '        End If
    '    End If

NaoAgendar:

    '========================== iNUTIL... mosilva
    '    'Verifico em que lugar da cole��o ser� adicionado.
    '    '� feito isso para que n�o seja exibido com os n�mero fora de ordem
    '    lPosicaoItem = pstcPedLoc.ObjPai.gfPosicaoColecao(pstcPedLoc.Num_Item)
    '    If pstcPedLoc.ObjPai.ColecaoPedLoc.Count >= lPosicaoItem Then
    '        pstcPedLoc.ObjPai.ColecaoPedLoc.Add pstcPedLoc, pstcPedLoc.Class_ID, lPosicaoItem
    '    Else
    '        pstcPedLoc.ObjPai.ColecaoPedLoc.Add pstcPedLoc, pstcPedLoc.Class_ID
    '    End If
    '==========================

    mfCalPreItem = True

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, Err.Description, eSeveridadeErro_Medio, Class_Name, methodName)
    Call M46V999.gpMenErro(Err.Description)
    GoTo SAIDA

FechaObj:
    Set lbdMulti = Nothing
    Set lrst = Nothing
    Set lstcCobert = Nothing
    Set lstcCobert_Acess = Nothing
    Set lobjCobert = Nothing
    Set lobjPedObs = Nothing
    Set lobjPedLoc = Nothing
    Set lColecaoTemp = Nothing
    Return
End Function

Private Function mfCalcFracionamento(ByRef pStcPed As stcA46V702B, ByRef pMsgErr As String, tagDynaTrace As String) As Boolean
    Dim lobjstcLoc As stcA46V708B
    Dim lIOF As Single    'VD - IOF.
    Dim lParcela As Variant   'Resultado do fracionamento.
    Dim lMsgErr As String

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfCalcFracionamento"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    On Error GoTo TrataErro
    Set lobjstcLoc = pStcPed.ColecaoPedLoc.Item(1)

    lIOF = pStcPed.Val_IOF
    pStcPed.Val_Cust_Emis = pStcPed.getProduto.ProCusApolice
    pStcPed.Val_Parc_Min = pStcPed.getProduto.ProPreMinParcela

    Set pStcPed.ColecaoParcelamento = New Collection
    If M46V998.gfCalFracionamento(pStcPed, pStcPed.ColecaoPedLoc(1).Tip_Emissao, _
                                  lParcela, lMsgErr, pStcPed.getProduto.ProJurIOF, pStcPed.Cod_Reg, lIOF, pStcPed.Val_Parc_Min) = False Then
        GoTo TrataErro
    End If

    If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 Then
        pStcPed.Val_PriParcela = lParcela(IIf(pStcPed.Qtd_Parc = 0, 1, pStcPed.Qtd_Parc), 1, pStcPed.gTipPagamento)
        pStcPed.val_DemParcelas = lParcela(IIf(pStcPed.Qtd_Parc = 0, 1, pStcPed.Qtd_Parc), 2, pStcPed.gTipPagamento)
    Else
        pStcPed.Val_PriParcela = lParcela(IIf(pStcPed.Qtd_Parc = 0, 1, pStcPed.Qtd_Parc), 3, pStcPed.gTipPagamento)
        pStcPed.val_DemParcelas = lParcela(IIf(pStcPed.Qtd_Parc = 0, 1, pStcPed.Qtd_Parc), 4, pStcPed.gTipPagamento)
    End If
    pStcPed.Val_IOF = lIOF

    If M46V999.gfVerificaParcelaAntec(pStcPed) Then
        pStcPed.Val_Pr_Antec = pStcPed.Val_PriParcela
    Else
        pStcPed.Val_Pr_Antec = 0
    End If

    mfCalcFracionamento = True

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

TrataErro:
    If M46V999.gfPreenchido(lMsgErr) = False Then
        lMsgErr = Err & " - " & ERROR
    End If
    pMsgErr = lMsgErr
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMsgErr, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lobjstcLoc = Nothing
    Return

End Function

Public Function gfSalvar(ByRef pStcPed As stcA46V702B, tagDynaTrace As String) As Boolean
    Dim lMsrErr As String

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfSalvar"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo TrataErro: gfSalvar = False

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "mfPosicionaUsuario retornou false", eSeveridadeAlerta, Class_Name, methodName)
        Exit Function
    End If

    If mfNormalizaPedido(pStcPed, lMsrErr, tagDynaTrace) = False Then
        GoTo TrataErro
    End If

    '
    Call M46V999.gpGraLog(0, "Trace - GFSalvar - Processou mfNormalizaPedido.")

    M46V999.gBufferSql = False

    '- Salva os Dados Comuns (incluindo todos os dependentes do pedido)
    If mfSalvarPedido(pStcPed, lMsrErr, tagDynaTrace) = False Then
        GoTo TrataErro
    End If

    If M46V999.gBufferSql Then
        M46V999.gBufferSql = False
    End If
    '***************** Retorno com Sucesso
    gfSalvar = True
    '*****************
    Call M46V999.gpGraLog(0, "Trace - mfSalvarPedido - Processou GFSalvar.")
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear
    Exit Function
TrataErro:
    gfSalvar = False
    If M46V999.gfPreenchido(lMsrErr) = False Then
        lMsrErr = Err & " - " & ERROR
    End If
    UltimoErro = lMsrErr
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA
End Function

Private Function mfSalvarPedido(ByRef pStcPed As stcA46V702B, ByRef pMsgErr As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfSalvarPedido"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lMsrErr As String
    Dim lClsPed As clsA46V702A
    Dim lStcPed As stcA46V702B
    Dim lobjCtrlEmis As clsA46V712A    'Objeto da classe clsA46V712A (TAB_CTRL_EMIS)
    Dim lobjEndereco As clsA46V707A    'Objeto da classe clsA46V707A (TAB_PED_ENDER)
    Dim lobjAuto As clsA46V719A    'Objeto da classe clsA46V719A (TAB_PED_AUTO)
    Dim lobjCorretor As clsA46V705A    'Objeto da classe clsA46V705A (TAB_PED_CORR)
    Dim lobjstcCorretor As stcA46V705B  'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    Dim lobjCosseguro As clsA46V706A    'Objeto da classe clsA46V706A (TAB_PED_COSS)
    Dim lobjstcCosseguro As stcA46V706B    'Objeto da classe de estrutura stcA46V706B (TAB_PED_COSS)
    Dim lobjJuros As clsA46V711A    'Objeto da classe clsA46V711A (TAB_PED_JUROS)
    Dim lobjstcJuros As stcA46V711B    'Objeto da classe de estrutura stcA46V711B (TAB_PED_JUROS)
    Dim lobjTextoEndo As clsA46V722A    'Objeto da classe clsA46V722A (TAB_PED_DIZER)
    Dim lobjstcTextoEndo As stcA46V722B   'Objeto da classe de estrutura stcA46V722B (TAB_PED_DIZER)
    Dim lobjVenc As clsA46V713A    'Objeto da classe clsA46V713A (TAB_PED_VENC)
    Dim lObjstcVenc As stcA46V713B    'Objeto da classe de estrutura stcA46V713A (TAB_PED_VENC)
    Dim lobjTextoEspec As clsA46V726A
    Dim lobjstcTextoEspec As stcA46V726B

    Dim lobjProlabore As clsA46V733A
    Dim lobjstcProlabore As stcA46V733B

    Dim lobjTelefones As clsA46V736A
    Dim lobjstcTelefones As stcA46V736B

    Dim lobjTextos As clsA46V739A
    Dim lobjstcTextos As stcA46V739B

    Err.Clear
    On Error GoTo TrataErro: mfSalvarPedido = False

    'Exclui os Dados
    M46V999.gBufferSql = True
    
    Set lClsPed = New clsA46V702A
    Set lClsPed.M46V999 = M46V999

    Set lStcPed = New stcA46V702B
    Set lStcPed.M46V999 = M46V999

    lStcPed.NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
    lStcPed.Cod_Unicidade = pStcPed.Cod_Unicidade
    lStcPed.NUM_PI = pStcPed.NUM_PI
    lStcPed.Ind_CotacaoSolic = pStcPed.Ind_CotacaoSolic
    lStcPed.objstcCtrlEmis.Val_VerbaCom = pStcPed.objstcCtrlEmis.Val_VerbaCom
    
    If lClsPed.mfExcluir(lStcPed, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
        lMsrErr = lClsPed.UltimoErro
        GoTo TrataErro
    End If

    
    M46V999.gBufferSql = True

    '- Inclui o Pedido
    If lClsPed.mfIncluir(pStcPed, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
        lMsrErr = lClsPed.UltimoErro
        GoTo TrataErro
    End If

    

    '- Inclui os Dependentes do Pedido
    Set lobjCtrlEmis = New clsA46V712A
    Set lobjCtrlEmis.M46V999 = M46V999

    'Ajuste forma de pagamento - Carn� + pagto � 30 dias n�o gera boleto
    If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec <> 1 Then
        pStcPed.objstcCtrlEmis.Cod_Bco_Boleto = 0
    Else
        'Forma de pagamento 1 = carn�
        If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 And pStcPed.gTipPagamento = 1 Then
            pStcPed.objstcCtrlEmis.Cod_Bco_Boleto = 0
        End If
    End If

    If lobjCtrlEmis.mfIncluir(pStcPed.objstcCtrlEmis, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
        lMsrErr = lobjCtrlEmis.UltimoErro
        GoTo TrataErro
    End If

    

    Set lobjEndereco = New clsA46V707A
    Set lobjEndereco.M46V999 = M46V999

    If lobjEndereco.mfIncluir(pStcPed.objstcPedEndereco, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
        lMsrErr = lobjEndereco.UltimoErro
        GoTo TrataErro
    End If

    

    Set lobjAuto = New clsA46V719A
    Set lobjAuto.M46V999 = M46V999

    If pStcPed.Cod_Ramo = e_CodRamo_Automoveis Then
        If lobjAuto.mfIncluir(pStcPed.objstcPedAuto, objstcUsuario.PropTabela, "") = False Then
            lMsrErr = lobjAuto.UltimoErro
            GoTo TrataErro
        End If
    End If

    

    Set lobjCorretor = New clsA46V705A
    Set lobjCorretor.M46V999 = M46V999

    For Each lobjstcCorretor In pStcPed.ColecaoPedCorr
        If lobjCorretor.mfIncluir(lobjstcCorretor, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
            lMsrErr = lobjCorretor.UltimoErro
            GoTo TrataErro
        End If
    Next lobjstcCorretor

    

    Set lobjCosseguro = New clsA46V706A
    Set lobjCosseguro.M46V999 = M46V999

    For Each lobjstcCosseguro In pStcPed.ColecaoPedCoss
        If lobjCosseguro.mfIncluir(lobjstcCosseguro, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
            lMsrErr = lobjCosseguro.UltimoErro
            GoTo TrataErro
        End If
    Next lobjstcCosseguro

    

    Set lobjJuros = New clsA46V711A
    Set lobjJuros.M46V999 = M46V999

    For Each lobjstcJuros In pStcPed.ColecaoPedJuros
        If lobjJuros.mfIncluir(lobjstcJuros, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
            lMsrErr = lobjJuros.UltimoErro
            GoTo TrataErro
        End If
    Next lobjstcJuros

    

    Set lobjTextoEndo = New clsA46V722A
    Set lobjTextoEndo.M46V999 = M46V999

    For Each lobjstcTextoEndo In pStcPed.ColecaoTextoEndo
        If lobjTextoEndo.mfIncluir(lobjstcTextoEndo, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
            lMsrErr = lobjTextoEndo.UltimoErro
            GoTo TrataErro
        End If
    Next lobjstcTextoEndo

    

    Set lobjVenc = New clsA46V713A
    Set lobjVenc.M46V999 = M46V999

    For Each lObjstcVenc In pStcPed.ColecaoPedVenc
        If lobjVenc.mfIncluir(lObjstcVenc, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
            lMsrErr = lobjVenc.UltimoErro
            GoTo TrataErro
        End If
    Next lObjstcVenc

    

    '- Salva os Itens (incluindo dependentes dos itens)
    If mfSalvarItens(pStcPed, lMsrErr, tagDynaTrace) = False Then
        GoTo TrataErro
    End If

    

    Set lobjTextoEspec = New clsA46V726A
    Set lobjTextoEspec.M46V999 = M46V999

    For Each lobjstcTextoEspec In pStcPed.ColecaoTextoEspec
        If lobjTextoEspec.mfIncluir(lobjstcTextoEspec, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
            lMsrErr = lobjTextoEspec.UltimoErro
            GoTo TrataErro
        End If
    Next lobjstcTextoEspec

    

    Set lobjProlabore = New clsA46V733A
    Set lobjProlabore.M46V999 = M46V999

    For Each lobjstcProlabore In pStcPed.ColecaoProlabore
        Set lobjstcProlabore.ObjPai = pStcPed
        lobjstcProlabore.Tip_Canal = 2
        lobjstcProlabore.Qtd_Parc_Apol = pStcPed.Qtd_Parc
        lobjstcProlabore.Qtd_Parc_Pag = pStcPed.Qtd_Parc
        If lobjProlabore.mfIncluir(lobjstcProlabore, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
            lMsrErr = lobjProlabore.UltimoErro
            GoTo TrataErro
        End If
    Next lobjstcProlabore

    

    'Agenciamento
    Set lobjProlabore = New clsA46V733A
    Set lobjProlabore.M46V999 = M46V999

    For Each lobjstcProlabore In pStcPed.ColecaoAgenciador
        lobjstcProlabore.Tip_Canal = 3
        Set lobjstcProlabore.ObjPai = pStcPed
        If lobjProlabore.mfIncluir(lobjstcProlabore, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
            lMsrErr = lobjProlabore.UltimoErro
            GoTo TrataErro
        End If
    Next lobjstcProlabore

    

    'Telefones
    Set lobjTelefones = New clsA46V736A
    Set lobjTelefones.M46V999 = M46V999

    For Each lobjstcTelefones In pStcPed.ColecaoTelefones
        Set lobjstcTelefones.ObjPai = pStcPed
        If lobjTelefones.mfIncluir(lobjstcTelefones, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
            lMsrErr = lobjTelefones.UltimoErro
            GoTo TrataErro
        End If
    Next lobjstcTelefones

    If M46V999.gBufferSql And Len(M46V999.gScriptSql) > 0 Then
        M46V999.gBufferSql = False
        If M46V999.gfExeSQL(M46V999.mConexaoSQL, M46V999.gScriptSql, lMsrErr) <> 0 Then
            Call M46V999.gpGraLog(2, lMsrErr)
        End If
    
        M46V999.gBufferSql = False
        M46V999.gScriptSql = ""
    End If

    '***************** Retorno com Sucesso
    mfSalvarPedido = True
    '*****************
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    GoSub FechaObj
    Err.Clear
    Exit Function

TrataErro:
    mfSalvarPedido = False
    If M46V999.gfPreenchido(lMsrErr) = False Then
        lMsrErr = Err & " - " & ERROR
    End If
    pMsgErr = lMsrErr
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMsrErr, 2, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lStcPed = Nothing
    Set lClsPed = Nothing
    Set lobjCtrlEmis = Nothing
    Set lobjEndereco = Nothing
    Set lobjAuto = Nothing
    Set lobjCorretor = Nothing
    Set lobjstcCorretor = Nothing
    Set lobjCosseguro = Nothing
    Set lobjstcCosseguro = Nothing
    Set lobjJuros = Nothing
    Set lobjstcJuros = Nothing
    Set lobjTextoEndo = Nothing
    Set lobjstcTextoEndo = Nothing
    Set lobjTextoEspec = Nothing
    Set lobjstcTextoEspec = Nothing
    Set lobjVenc = Nothing
    Set lObjstcVenc = Nothing
    Return
End Function

Private Function mfZerarDescontoLMI(ByRef pStcPed As stcA46V702B)
    Dim lUpdate As String
    Dim lCnSyasPortal2008 As ADODB.Connection
    
    If M46V999.gfAbrBasDados("SYASPORTAL", M46V999.gTipSyasPortal, M46V999.gEndSyasPortal, lCnSyasPortal2008, UltimoErro) = False Then
        Exit Function
    End If

    lUpdate = "UPDATE SyasPortal.dbo.Tab_Ped SET Per_Desc_LMI = 0"
    lUpdate = lUpdate & " WHERE NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pStcPed.NOSSO_NUMERO)
    
    If M46V999.gfExeSQL(lCnSyasPortal2008, lUpdate, UltimoErro) <> 0 Then
        Exit Function
    End If

End Function

Private Function mfSalvarItens(ByRef pStcPed As stcA46V702B, ByRef pMsgErr As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfSalvarItens"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lMsrErr As String
    Dim lobjItem As clsA46V708A    'Objeto da classe clsA46V708A (TAB_PED_LOC)
    Dim lobjstcItem As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjCobertura As clsA46V704A    'objeto da classe clsA46V704A (TAB_PED_COBERT)
    Dim lobjstcCobertura As stcA46V704B    'objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
    Dim lobjClausula As clsA46V703A    'objeto da classe clsA46V703A (TAB_PED_CLAU)
    Dim lobjStcClausula As stcA46V703B    'objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU)
    Dim lobjObservacao As clsA46V714A    'objeto da classe clsA46V714A (TAB_PED_OBS)
    Dim lobjBem As clsA46V715A                 'objeto da classe clsA46V714A (TAB_PED_OBS)
    Dim lobjstcBem As stcA46V715B    'objeto da classe de estrutura stcA46V714B (TAB_PED_OBS)
    Dim lobjSeguro As clsA46V716A    'objeto da classe clsA46V716A (TAB_PED_LOC_SEGURO)
    Dim lobjstcSeguro As stcA46V716B    'objeto da classe de estrutura stcA46V716B (TAB_PED_LOC_SEGURO)
    Dim lobjInsp As clsA46V717A    'objeto da classe clsA46V717A (TAB_PED_INSP)
    Dim lobjFator As clsA46V720A    'objeto da classe clsA46V720A (Tab_Ped_Fat_Risco)
    Dim lobjstcFator As stcA46V720B    'objeto da classe de estrutura stcA46V720B (Tab_Ped_Fat_Risco)
    Dim lobjLMI As clsA46V728A    'objeto da classe clsA46V728A (Tab_Ped_Loc_LMI)
    Dim lobjstcLMI As stcA46V728B    'objeto da classe de estrutura stcA46V728B (Tab_Ped_Loc_LMI)

    Dim lobjPessoas As clsA46V735A    'objeto da classe clsA46V735A (TAB_PED_PESSOA)
    Dim lobjstcPessoas As stcA46V735B    'objeto da classe de estrutura stcA46V735B (TAB_PED_PESSOA)
    
    Dim lobjCobertMultFraq As clsA46V738A    'objeto da classe clsA46V738A (TAB_PED_COBERT_FRANQUIA)
    Dim lobjstcCobertMultFraq As stcA46V738B    'objeto da classe de estrutura stcA46V738B (TAB_PED_COBERT_FRANQUIA)
    
    Dim lobjTextos As clsA46V739A       'objeto da classe clsA46V738A (TAB_PED_Textos)
    Dim lobjstcTextos As stcA46V739B    'objeto da classe de estrutura stcA46V739B (TAB_PED_Textos)

    Dim lobjPedLocBicicleta As clsA46V744A  'objeto da classe clsA46V738A (TAB_PED_LOC_BICICLETA)
    Dim lobjstcPedLocBicicleta As stcA46V744B 'objeto da classe de estrutura stcA46V739B (TAB_PED_LOC_BICICLETA)

    Dim lobjPedAllRisks As clsA46V747A  'objeto da classe clsA46V747A (Tab_Ped_All_Risks)
    Dim lobjstcPedAllRisks As stcA46V747B 'objeto da classe de estrutura stcA46V747B (Tab_Ped_All_Risks)
    
    Dim lobjPedObraArte As clsA46V748A  'objeto da classe clsA46V748A (tab_ped_obras_arte)
    Dim lobjstcPedObraArte As stcA46V748B 'objeto da classe de estrutura stcA46V748B (tab_ped_obras_arte)

    Err.Clear
    On Error GoTo TrataErro: mfSalvarItens = False

    Set lobjItem = New clsA46V708A
    Set lobjItem.M46V999 = M46V999

    Set lobjCobertura = New clsA46V704A
    Set lobjCobertura.M46V999 = M46V999

    Set lobjClausula = New clsA46V703A
    Set lobjClausula.M46V999 = M46V999

    Set lobjObservacao = New clsA46V714A
    Set lobjObservacao.M46V999 = M46V999

    Set lobjBem = New clsA46V715A
    Set lobjBem.M46V999 = M46V999

    Set lobjSeguro = New clsA46V716A
    Set lobjSeguro.M46V999 = M46V999

    Set lobjInsp = New clsA46V717A
    Set lobjInsp.M46V999 = M46V999

    Set lobjFator = New clsA46V720A
    Set lobjFator.M46V999 = M46V999

    Set lobjLMI = New clsA46V728A
    Set lobjLMI.M46V999 = M46V999


    Set lobjPessoas = New clsA46V735A
    Set lobjPessoas.M46V999 = M46V999
    
    Set lobjCobertMultFraq = New clsA46V738A
    Set lobjCobertMultFraq.M46V999 = M46V999
    
    Set lobjstcCobertMultFraq = New stcA46V738B
    Set lobjstcCobertMultFraq.M46V999 = M46V999
    
    Set lobjTextos = New clsA46V739A
    Set lobjTextos.M46V999 = M46V999
    
    Set lobjstcTextos = New stcA46V739B
    Set lobjstcTextos.M46V999 = M46V999

    Set lobjPedLocBicicleta = New clsA46V744A
    Set lobjPedLocBicicleta.M46V999 = M46V999

    Set lobjPedAllRisks = New clsA46V747A
    Set lobjPedAllRisks.M46V999 = M46V999
    
    Set lobjPedObraArte = New clsA46V748A
    Set lobjPedObraArte.M46V999 = M46V999

    For Each lobjstcItem In pStcPed.ColecaoPedLoc
        '- Inclui o Item
        If lobjItem.mfIncluir(lobjstcItem, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
            lMsrErr = lobjItem.UltimoErro
            GoTo TrataErro
        End If

        'Verifica se tem coberturas de vida se estiver inclui na cole��o de cobertura antes de salvar
        'lMsrErr = ""
        'Call M46V103.CarregaCoberturaVida(lobjstcItem, lMsrErr)
        'If lMsrErr <> "" Then
        '    GoTo TrataErro
        'End If

        '- Inclui os dependentes do Item
        For Each lobjstcCobertura In lobjstcItem.ColecaoPedCobert
            If lobjCobertura.mfIncluir(lobjstcCobertura, objstcUsuario.PropTabela, "", False, False, tagDynaTrace) = False Then
                lMsrErr = lobjCobertura.UltimoErro
                GoTo TrataErro
            End If
            
            'Projeto Multiplo Franquia
            For Each lobjstcCobertMultFraq In lobjstcCobertura.ColecaoNovaFranquia
                If lobjCobertMultFraq.mfIncluir(lobjstcCobertMultFraq, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                    lMsrErr = lobjCobertura.UltimoErro
                    GoTo TrataErro
                End If
            Next
        Next

        For Each lobjstcTextos In lobjstcItem.ColecaoTextos
            If lobjTextos.mfIncluir(lobjstcTextos, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                lMsrErr = lobjTextos.UltimoErro
                GoTo TrataErro
            End If
        Next

        For Each lobjStcClausula In lobjstcItem.ColecaoPedClau
            If lobjClausula.mfIncluir(lobjStcClausula, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                lMsrErr = lobjClausula.UltimoErro
                GoTo TrataErro
            End If
        Next
        
        If M46V999.gfPreenchido(lobjstcItem.objstcPedObs.Observacao) And lobjstcItem.ObjPai.Ind_CotacaoSolic = False Then
            If lobjObservacao.mfIncluir(lobjstcItem.objstcPedObs, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                lMsrErr = lobjObservacao.UltimoErro
                GoTo TrataErro
            End If
        End If
        
        For Each lobjstcBem In lobjstcItem.ColecaoPedLocBem
            If lobjBem.mfIncluir(lobjstcBem, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                lMsrErr = lobjBem.UltimoErro
                GoTo TrataErro
            End If
        Next

        For Each lobjstcPedLocBicicleta In lobjstcItem.ColecaoPedLocBicicleta
            If lobjPedLocBicicleta.mfIncluir(lobjstcPedLocBicicleta, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                lMsrErr = lobjPedLocBicicleta.UltimoErro
                GoTo TrataErro
            End If
        Next

        For Each lobjstcPedAllRisks In lobjstcItem.ColecaoAllRisks
            If lobjPedAllRisks.mfIncluir(lobjstcPedAllRisks, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                lMsrErr = lobjPedAllRisks.UltimoErro
                GoTo TrataErro
            End If
        Next
        
        For Each lobjstcPedObraArte In lobjstcItem.ColecaoObrasArte
            If lobjPedObraArte.mfIncluir(lobjstcPedObraArte, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                lMsrErr = lobjPedObraArte.UltimoErro
                GoTo TrataErro
            End If
        Next

        For Each lobjstcSeguro In lobjstcItem.ColecaoPedSeguro
            If lobjSeguro.mfIncluir(lobjstcSeguro, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                lMsrErr = lobjSeguro.UltimoErro
                GoTo TrataErro
            End If
        Next

        If lobjInsp.mfIncluir(lobjstcItem.objstcPedInsp, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
            lMsrErr = lobjInsp.UltimoErro
            GoTo TrataErro
        End If

        For Each lobjstcFator In lobjstcItem.ColecaoPedFator
            If lobjFator.mfIncluir(lobjstcFator, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                lMsrErr = lobjFator.UltimoErro
                GoTo TrataErro
            End If
        Next

        For Each lobjstcLMI In lobjstcItem.ColecaoLMI
            If lobjLMI.mfIncluir(lobjstcLMI, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                lMsrErr = lobjLMI.UltimoErro
                GoTo TrataErro
            End If
        Next

        'ProjetoFormularios
        For Each lobjstcPessoas In lobjstcItem.ColecaoBeneficiario
            lobjstcPessoas.Tip_Loc_Pessoa = e_TipLocPessoa_Beneficiario
            If lobjPessoas.mfIncluir(lobjstcPessoas, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                lMsrErr = lobjPessoas.UltimoErro
                GoTo TrataErro
            End If
        Next


        'ProjetoFormularios
        If M46V999.gfPreenchido(lobjstcItem.objstcPedInquilino.Num_CPF_CNPJ) And Trim(lobjstcItem.objstcPedInquilino.Num_CPF_CNPJ) <> "0" Then
            Set lobjstcPessoas = lobjstcItem.objstcPedInquilino
            lobjstcPessoas.Tip_Loc_Pessoa = e_TipLocPessoa_Inquilino
            If lobjPessoas.mfIncluir(lobjstcPessoas, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                lMsrErr = lobjPessoas.UltimoErro
                GoTo TrataErro
            End If
        End If


        'ProjetoFormularios
        If M46V999.gfPreenchido(lobjstcItem.objstcPedProprietario.Num_CPF_CNPJ) And Trim(lobjstcItem.objstcPedProprietario.Num_CPF_CNPJ) <> "0" Then
            Set lobjstcPessoas = lobjstcItem.objstcPedProprietario
            lobjstcPessoas.Tip_Loc_Pessoa = e_TipLocPessoa_Proprietario
            If lobjPessoas.mfIncluir(lobjstcPessoas, objstcUsuario.PropTabela, "", tagDynaTrace) = False Then
                lMsrErr = lobjPessoas.UltimoErro
                GoTo TrataErro
            End If
        End If
    Next lobjstcItem


    '***************** Retorno com Sucesso
    mfSalvarItens = True
    '*****************
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    GoSub FechaObj
    Err.Clear
    Exit Function

TrataErro:
    mfSalvarItens = False
    If M46V999.gfPreenchido(lMsrErr) = False Then
        lMsrErr = Err & " - " & ERROR
    End If
    pMsgErr = lMsrErr
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMsrErr, 2, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lobjItem = Nothing
    Set lobjstcItem = Nothing
    Set lobjCobertura = Nothing
    Set lobjstcCobertura = Nothing
    Set lobjClausula = Nothing
    Set lobjStcClausula = Nothing
    Set lobjObservacao = Nothing
    Set lobjBem = Nothing
    Set lobjstcBem = Nothing
    Set lobjSeguro = Nothing
    Set lobjstcSeguro = Nothing
    Set lobjInsp = Nothing
    Set lobjFator = Nothing
    Set lobjstcFator = Nothing
    Set lobjLMI = Nothing
    Set lobjstcLMI = Nothing
    Set lobjstcPessoas = Nothing
    Set lobjPessoas = Nothing
    Set lobjCobertMultFraq = Nothing
    Set lobjstcCobertMultFraq = Nothing
    Set lobjTextos = Nothing
    Set lobjstcTextos = Nothing
    Set lobjstcPedAllRisks = Nothing
    Set lobjstcPedObraArte = Nothing

    Return

End Function

Public Function gfCopiarProposta(ByVal pAntNosNumero As String, _
                                 ByVal pReentrada As Integer, tagDynaTrace As String) As stcA46V702B
    Dim lAntNosNumero As String     'Anterior: nosso n�mero.
    Dim lNovNosNumero As String     'Novo: nosso n�mero.
    Dim lNovNumPedido As String     'Novo: n�mero do pediro.
    Dim lNumApolice As String   'N�mero da ap�lice.
    Dim lReentrada As Integer
    Dim objPed As clsA46V702A
    Dim objstcPed As stcA46V702B
    Dim lStcPed As stcA46V702B

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfCopiarProposta"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    Set gfCopiarProposta = Nothing

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "mfPosicionaUsuario retornou false", eSeveridadeAlerta, Class_Name, methodName)
        GoTo SAIDA
    End If

    '1. Verifica se tem documento selecionado.
    'If .Row = 0 Then
    If M46V999.CDblx(pAntNosNumero) = 0 Then
        UltimoErro = "FNP0014 - Documento n�o selecionado."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    If val(pReentrada) = 0 Then
        UltimoErro = "FNP0014 - Preencher Parm de Reentrada."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '2. Posiciona controles.
    M46V999.gPrefixo = " Pedido " & pAntNosNumero & ". "
    lAntNosNumero = pAntNosNumero  'Nosso n�mero.

    lReentrada = pReentrada

    '4. Calcula nosso n�mero e n�mero do pedido.
    If M46V999.gfCalNosNumero(lNovNosNumero, UltimoErro) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfCalNosNumero retornou false", eSeveridadeAlerta, Class_Name, methodName)
        GoTo SAIDA
    End If

    If M46V999.gfCalNumPedido(lNovNumPedido, UltimoErro) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfCalNumPedido retornou false", eSeveridadeAlerta, Class_Name, methodName)
        GoTo SAIDA
    End If

    '5. Copia documento.
    If M46V999.gfCopProposta(lAntNosNumero, lNovNosNumero, lNovNumPedido, UltimoErro, 0, 0, 0, , , lReentrada) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfCopProposta retornou false", eSeveridadeAlerta, Class_Name, methodName)
        GoTo SAIDA
    End If
    If lReentrada = 1 Then
        Set objPed = New clsA46V702A
        Set objPed.M46V999 = M46V999

        Set objstcPed = New stcA46V702B
        Set objstcPed.M46V999 = M46V999

        objstcPed.NOSSO_NUMERO = lNovNosNumero
        If Not objPed.mfAlterarReentrada(objstcPed, objstcUsuario.PropTabela, lReentrada, lAntNosNumero, tagDynaTrace) Then
            UltimoErro = objPed.UltimoErro
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeAlerta, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If

    'g) Conclus�o.
    Call M46V999.gpGraLog(0, "AGE0002 - Processo conclu�do: c�pia realizada para pedido " & lNovNumPedido & _
                             " (nosso n�mero " & lNovNosNumero & ").")

    '6. Abre novo documento.

    M46V999.gPrefixo = M46V999.gPrefixo & " " & lNovNumPedido & ". "
    'If M46V111.gfGraRegComunicacao(lNovNosNumero, UltimoErro) = False Then
    '    Call M46V999.gpGraLog(0,, UltimoErro)
    '    Exit Sub
    'End If

    Set lStcPed = gfProcuraNossoNumero(lNovNosNumero, False, "", tagDynaTrace)

    Set gfCopiarProposta = lStcPed

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    UltimoErro = "[ERR] - gfCopiarProposta " & Err & " " & Err.Description
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    
    GoTo SAIDA

FechaObj:
    Set objPed = Nothing
    Set objstcPed = Nothing
    Set lStcPed = Nothing
    Return

End Function

Public Function gfCopiarPropostaClasse4(ByVal pNossoNumeroAnt As String, _
                                        ByVal pNumPedAnt As String, _
                                        ByVal pNumCotacAnt As String, _
                                        ByVal pReentrada As Integer, tagDynaTrace As String) As stcA46V702B
    Dim lAntNosNumero As String     'Anterior: nosso n�mero.
    Dim lNovNosNumero As String     'Novo: nosso n�mero.
    Dim lNovNumPedido As String     'Novo: n�mero do pedido.
    
    
    Dim lNumApolice As String       'N�mero da ap�lice.
    Dim lReentrada As Integer
    Dim objPed As clsA46V702A
    Dim objstcPed As stcA46V702B
    Dim lStcPed As stcA46V702B

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    
    methodName = "gfCopiarPropostaClasse4"
    M46V999.tagDynaTrace = tagDynaTrace
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    Set gfCopiarPropostaClasse4 = Nothing

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "mfPosicionaUsuario retornou false", eSeveridadeAlerta, Class_Name, methodName)
        GoTo SAIDA
    End If

    '1. Verifica se tem documento selecionado.
    If M46V999.CDblx(pNossoNumeroAnt) = 0 Then
        UltimoErro = "FNP0014 - Cota��o n�o existente."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    If val(pReentrada) = 0 Then
        UltimoErro = "FNP0014 - Preencher Parm de Reentrada."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '2. Posiciona controles.
    M46V999.gPrefixo = " Pedido " & pNossoNumeroAnt & ". "
    lAntNosNumero = pNossoNumeroAnt  'Nosso n�mero.

    lReentrada = pReentrada

    '4. Calcula nosso n�mero e n�mero do pedido.
    If M46V999.gfCalNosNumero(lNovNosNumero, UltimoErro) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfCalNosNumero retornou false", eSeveridadeAlerta, Class_Name, methodName)
        GoTo SAIDA
    End If

    If M46V999.gfCalNumPedido(lNovNumPedido, UltimoErro) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfCalNumPedido retornou false", eSeveridadeAlerta, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    If M46V999.gfCopPropostaClasse4(lAntNosNumero, pNumPedAnt, pNumCotacAnt, lNovNosNumero, lNovNumPedido, UltimoErro, 0, 0, 0, , , lReentrada) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfCopProposta retornou false", eSeveridadeAlerta, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    If lReentrada = 1 Then
        Set objPed = New clsA46V702A
        Set objPed.M46V999 = M46V999

        Set objstcPed = New stcA46V702B
        Set objstcPed.M46V999 = M46V999

        objstcPed.NOSSO_NUMERO = lNovNosNumero
        If Not objPed.mfAlterarReentrada(objstcPed, objstcUsuario.PropTabela, lReentrada, lAntNosNumero, tagDynaTrace) Then
            UltimoErro = objPed.UltimoErro
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeAlerta, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If

    'g) Conclus�o.
    Call M46V999.gpGraLog(0, "AGE0002 - Processo conclu�do: Nova cota��o gerada: Refer�ncia: " & lNovNumPedido & _
                             " (Nosso N�mero: " & lNovNosNumero & ").")

    '6. Abre novo documento.

    M46V999.gPrefixo = M46V999.gPrefixo & " " & lNovNumPedido & ". "
    'If M46V111.gfGraRegComunicacao(lNovNosNumero, UltimoErro) = False Then
    '    Call M46V999.gpGraLog(0,, UltimoErro)
    '    Exit Sub
    'End If

    Set lStcPed = gfProcuraNossoNumero(lNovNosNumero, False, "", tagDynaTrace)

    Set gfCopiarPropostaClasse4 = lStcPed

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    UltimoErro = "[ERR] - gfCopiarPropostaClasse4 " & Err & " " & Err.Description
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    
    GoTo SAIDA

FechaObj:
    Set objPed = Nothing
    Set objstcPed = Nothing
    Set lStcPed = Nothing
    Return

End Function

Public Function gfExcluirProposta(ByVal pNosNumero As String, tagDynaTrace As String) As Boolean
'Procedure: exclui documento selecionada (grava N no campo Cod_Ativo).

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfExcluirProposta"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    gfExcluirProposta = False
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    If UCase(Mid(pNosNumero, 1, 6)) <> "PADRAO" And val(pNosNumero) = 0 Then
        UltimoErro = "FNP0014 - Documento n�o selecionado."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    'b) Exclui registro de dados.
    Call M46V999.gpExcluiRegistros(pNosNumero, 0, objstcUsuario.PropTabela, "", UltimoErro)
    Call M46V999.gpGraLog(0, "AGE0002 - Processo conclu�do: documento " & pNosNumero & _
                             " exclu�do com sucesso.")

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    
    gfExcluirProposta = True
SAIDA:
    Exit Function

ControleErro:
    UltimoErro = "[ERR] - gfExcluirProposta ERRO - " & Err & " " & Err.Description
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA
End Function

Public Function gfProtocoloTransmissao(ByVal pNossoNumero As String, _
                                       ByVal pPropTabela As String, _
                                       tagDynaTrace As String) As String

    Dim lDir As String    'Utilizado com a fun��o DIR.
    Dim lMensagem As String     'Mensagem.
    Dim lNosNumero As String      'Nosso numero
    Dim lNumArquivo As Integer      'Utilizado com FreeFile.
    Dim lPriProposta As Boolean       'Primeira proposta.
    Dim lRegistro As String     'Registro.
    Dim lobjTransm As clsA46V718A    'Objeto da classe clsA46V718A (TAB_PED_TRANSM)
    Dim lobjstcTransm As stcA46V718B    'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM)
    Dim lCount As Long    'Contador.
    Dim clsA99V001 As clsA99V001
    Dim lDigito As String
    Dim mQuaLinhas As Long
    Dim lProtocolo As String
    Dim Ind_Cancel As String

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfProtocoloTransmissao"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    Call M46V999.gpGraLog(0, "Tag DynaTrace: " & tagDynaTrace & " | KeyDynaTrace " & CStr(lKeyDynaTrace))
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(M46V999.gDataCorteSompo) Then
            If CLng(M46V999.gDataCorteSompo) <= M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If

    'Obt�m registros de protocolo referente ao nosso n�mero.
    Set lobjTransm = New clsA46V718A
    Set lobjTransm.M46V999 = M46V999

    Set lobjstcTransm = New stcA46V718B
    lobjstcTransm.NOSSO_NUMERO = pNossoNumero

    If (pPropTabela <> "") Then
        objstcUsuario.PropTabela = pPropTabela + ".dbo."
    End If

    If Not lobjTransm.mfCarregarColecao(lobjstcTransm, objstcUsuario.PropTabela, "") Then
        UltimoErro = lobjTransm.UltimoErro
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If
    If lobjTransm.Colecao.Count = 0 Then
        UltimoErro = "FNL0014 - Registro de protocolo de transmiss�o n�o localizado (pedido " & _
                     M46V111.gNumPedido & " - nosso n�mero " & pNossoNumero & ")."
        Set lobjTransm = Nothing
        Set lobjstcTransm = Nothing
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    'Obt�m dados da proposta transmitida (detalhes - erro).
    '*************************Ver este ponto

    'Processa registro de protocolo.
    lPriProposta = True
    '6. Inicia processo.
    mQuaLinhas = 0

    Ind_Cancel = "0"

    On Error Resume Next
    Set lobjstcTransm = lobjTransm.Colecao(1)
    Ind_Cancel = lobjstcTransm.Ind_Cancel
    On Error GoTo ControleErro

    If val(Ind_Cancel) <> 0 Then
        Call mpPosLinha(lProtocolo, mQuaLinhas, UCase(lNomeSeguradora) & " - PROTOCOLO DE CANCELAMENTO DE TRANSMISS�O - " & Format$(Now, "dd/mm/yyyy hh:nn:ss") & ".")
    Else
        Call mpPosLinha(lProtocolo, mQuaLinhas, UCase(lNomeSeguradora) & " - PROTOCOLO DE TRANSMISS�O - " & Format$(Now, "dd/mm/yyyy hh:nn:ss") & ".")
    End If

    Call mpPosLinha(lProtocolo, mQuaLinhas, " ")

    For Each lobjstcTransm In lobjTransm.Colecao
        'a) Exibe dados
        'Call mpPosLinha(lProtocolo, mQuaLinhas,"Transmiss�o da Yasuda: " & lobjstcTransm.Nom_Arq_Retorno & ".")
        If lNosNumero <> lobjstcTransm.NOSSO_NUMERO Then
            lNosNumero = lobjstcTransm.NOSSO_NUMERO
            If M46V999.gfPreenchido(lobjstcTransm.NOSSO_NUMERO) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Proposta recebida....." & lobjstcTransm.NOSSO_NUMERO)
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Proposta recebida.....n�o informado")
            End If
            If M46V999.gfPreenchido(lobjstcTransm.Num_Ped) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Refer�ncia............" & lobjstcTransm.Num_Ped)
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Refer�ncia............n�o informado")
            End If
            If M46V999.gfPreenchido(lobjstcTransm.Cod_Ramo) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Ramo.................." & lobjstcTransm.Cod_Ramo)
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Ramo..................n�o informado")
            End If
            If M46V999.gfPreenchido(lobjstcTransm.Nom_Segurado) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Segurado.............." & Trim$(lobjstcTransm.Nom_Segurado))
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Segurado..............n�o informado.")
            End If
            If M46V999.gfPreenchido(lobjstcTransm.Dat_Ini_Vig) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "In�cio de vig�ncia...." & M46V999.gfForData(2, lobjstcTransm.Dat_Ini_Vig))
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "In�cio de vig�ncia....n�o informado")
            End If
            If M46V999.gfPreenchido(lobjstcTransm.Dat_Ter_Vig) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "T�rmino de vig�ncia..." & M46V999.gfForData(2, lobjstcTransm.Dat_Ter_Vig))
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "T�rmino de vig�ncia...n�o informado")
            End If
            Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Itens da proposta:")
        End If
        '***==== DADOS ITENS
        Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
        If M46V999.gfPreenchido(lobjstcTransm.Num_Item) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Item.................." & lobjstcTransm.Num_Item)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Item..................n�o informado")
        End If
        Call mpPosLinha(lProtocolo, mQuaLinhas, "Logradouro............" & lobjstcTransm.Cod_Tip_Logr_Risco & _
                                                " " & lobjstcTransm.Nom_Logr_Risco & _
                                                ", " & lobjstcTransm.Num_Risco)

        If M46V999.gfPreenchido(lobjstcTransm.Dsc_Compl_Risco) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Complemento..........." & lobjstcTransm.Dsc_Compl_Risco)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Complemento...........n�o informado")
        End If
        If M46V999.gfPreenchido(lobjstcTransm.Num_Cep_Risco) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "CEP..................." & Format(lobjstcTransm.Num_Cep_Risco, "00000-000"))
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "CEP...................n�o informado")
        End If
        If M46V999.gfPreenchido(lobjstcTransm.Nom_Cid_Risco) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Cidade................" & lobjstcTransm.Nom_Cid_Risco)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Cidade................n�o informada")
        End If
        If M46V999.gfPreenchido(lobjstcTransm.Sig_UF_Risco) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "UF...................." & lobjstcTransm.Sig_UF_Risco)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "UF....................n�o informada")
        End If
        Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
        If val(lobjstcTransm.NUM_PI) <> 0 Then
            If val(lobjstcTransm.Ind_Cancel) <> 0 Then
                'Transmiss�o Cancelada!
                Select Case lobjstcTransm.Ind_Cancel
                Case 1
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "Emiss�o do PI " & lobjstcTransm.NUM_PI & " cancelada com sucesso.")
                    Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "***Desprezar proposta f�sica e boletos gerados.***")
                Case 3
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "Emiss�o do PI " & lobjstcTransm.NUM_PI & " n�o pode ser cancelada. PI j� recusado.")
                Case 2
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "Emiss�o do PI " & lobjstcTransm.NUM_PI & " n�o pode ser cancelada. PI j� cancelado.")
                Case 4
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "Emiss�o do PI " & lobjstcTransm.NUM_PI & " n�o pode ser cancelada. PI j� emitido.")
                End Select
            Else
                Set clsA99V001 = New clsA99V001
                Set clsA99V001.M46V999 = M46V999

                If clsA99V001.DIG_YASUDA(lobjstcTransm.NUM_PI, lDigito, lMensagem) = False Then
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, eSeveridadeErro_Medio, Class_Name, methodName)
                    GoTo SAIDA
                End If
                
                Set clsA99V001 = Nothing
                If lDigito = "A" Then lDigito = "0"
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Protocolo Interno....." & lobjstcTransm.NUM_PI & "-" & lDigito)
                
                Dim lObjPed As clsA46V702A
                Set lObjPed = New clsA46V702A
                Set lObjPed.M46V999 = M46V999

                lObjPed.UltimoErro = Empty
                If Not lObjPed.mfInspecaoObrigatoria(pNossoNumero, "SYASPORTAL", "") Then
                    If lObjPed.UltimoErro <> Empty Then
                        UltimoErro = lObjPed.UltimoErro
                        GoTo ControleErro
                    End If
                Else
                    Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "Pend�ncia de aceita��o tendo em vista a necessidade de inspe��o conforme")
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "j� informado na cota��o.")
                End If
                
                Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Os anexos, quando necess�rios, poder�o ser enviados das seguintes formas:")
                Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "1) VIA FAX: Atrav�s  do  n�mero (0xx11) 3305-5664,  informando  o  n�mero")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "do protocolo interno " & lobjstcTransm.NUM_PI & lDigito & " no atendimento eletr�nico ou;")
                Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "2) VIA " & lNomeSeguradora & "Net, '�REA DO CORRETOR': Carregar o(s) documento(s)")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "digitalizado(s) atrav�s  dos  links  Seguro  Demais Ramos  -->  Emiss�o.")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Filtrar pelo N�mero do Protocolo (PI).")
                Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
            End If
        End If
        Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
        Select Case lobjstcTransm.Ind_Proposta
        Case "0"
            Call mpPosLinha(lProtocolo, mQuaLinhas, "STATUS................Proposta n�o transmitida.")
        Case 1

        Case 2
            Call mpPosLinha(lProtocolo, mQuaLinhas, "STATUS................Proposta apresentou erro.")
            Call mpPosLinha(lProtocolo, mQuaLinhas, "� necess�rio retransmitir esta proposta.")
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Descri��o do erro....." & lobjstcTransm.Tip_Emissao)    'Cont�m observa��o do erro
        Case Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "STATUS................Proposta apresentou erro.")
            Call mpPosLinha(lProtocolo, mQuaLinhas, "� necess�rio retransmitir esta proposta.")
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Descri��o do erro.....Indicador de proposta desconhecido - comunicar Central de Atendimento.")
        End Select
    Next lobjstcTransm


    gfProtocoloTransmissao = lProtocolo

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    If Err <> 0 Then
    UltimoErro = "[ERR] - gfProtocoloTransmissao ERRO - " & Err & " " & Err.Description & UltimoErro
    End If
    UltimoErro = UltimoErro & "Linha: " & Erl
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lobjstcTransm = Nothing
    Set lobjTransm = Nothing
    Set clsA99V001 = Nothing
    Return

End Function


Public Function gfCancelarEmissao(ByVal pNossoNumero As String, _
                                  ByRef pMsgRetorno As String, tagDynaTrace As String) As Boolean
    Dim lMensagem As String             'Mensagem.
    Dim lPoint As String
    '================================================================================================
    Dim lHora As Date
    Dim lSelect As String
    Dim lUpdate As String
    Dim lrst As clsYasRecordSet
    Dim lCnRamosDiv As ADODB.Connection
    Dim lCnSyasPortal2008 As ADODB.Connection
    Dim lUsuario As String
    Dim lUsuarioHora As String
    Dim lObjstcPed As stcA46V702B


    Dim lIndCancel As Integer

    Dim lMsgRetorno As String
    Dim lretorno As Boolean
    Dim lAlertas As String

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfCancelarEmissao"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro: lretorno = False

    If M46V999.gTipoTransmDLL = ETransmDLL_OFFLine Then
        gfCancelarEmissao = mfCancelarEmissaoOFFLine(pNossoNumero, pMsgRetorno, tagDynaTrace)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Transmissao Offline", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Falha ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        Exit Function
    End If

    Set lObjstcPed = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)

    If lObjstcPed.Cod_Sit_Ped = e_SitPedido_Transmitido Then
        If lObjstcPed.Num_PI_Emissao = 0 Then
            UltimoErro = "FGE2062 - Registro inconsistente: Proposta transmitida sem n�mero de PI."
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
    Else
        UltimoErro = "AGE0086 - Documento n�o transmitido"
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    lUsuario = Me.objstcUsuario.login_user_rede

    lPoint = "01"
    lHora = Time
    If lHora > "19:00" Or lHora < "06:00" Then
        lMensagem = "O cancelamento permitido apenas em hor�rio comercial (06:00 �s 19:00)."
    End If

    lUsuarioHora = "Usu�rio: " & Mid(lUsuario, 1, 6) & " " & Format(Now, "DD/MM/YYYY Hh:Nn:Ss")
    lPoint = "04"

    If M46V999.gfPreenchido(lMensagem) Then
        'Erro
        lMsgRetorno = "CANCELAMENTO N�O EFETUADO. " & vbCrLf & vbCrLf _
                    & "Nosso N�mero .: " & pNossoNumero & vbCrLf _
                    & "Num�ro PI ....: " & lObjstcPed.Num_PI_Emissao & vbCrLf _
                    & "01 ...........: " & lMensagem
    ElseIf M46V999.gfPreenchido(pNossoNumero) = False Or val(Trim$(lObjstcPed.Num_PI_Emissao)) = 0 Then
        lMsgRetorno = "CANCELAMENTO N�O EFETUADO. " & vbCrLf & vbCrLf _
                    & "Nosso N�mero .: " & pNossoNumero & vbCrLf _
                    & "Num�ro PI ....: " & lObjstcPed.Num_PI_Emissao & vbCrLf _
                    & "02 ...........: Par�metros inv�lidos"
    Else
        lPoint = "02"
        If Not M46V999.gfAbrBasDados("RAMOSDIVERSOS", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, lCnRamosDiv, lMensagem) Then
            If Not UCase(M46V999.gEndDIVERSOS) Like "*MIX*" Then
                Shell "net send tjgomes " & M99V001.GetLoginComputerName & " com falha na conex�o SQL.  Verificar com urg�ncia. (SyasWeb)"
                Shell "net send hideo " & M99V001.GetLoginComputerName & " com falha na conex�o SQL.  Verificar com urg�ncia. (SyasWeb)"
                Shell "net send newton " & M99V001.GetLoginComputerName & " com falha na conex�o SQL.  Verificar com urg�ncia. (SyasWeb)"
                Shell "net send oscar " & M99V001.GetLoginComputerName & " com falha na conex�o SQL.  Verificar com urg�ncia. (SyasWeb)"
            End If
        End If
        'Executa a procedure
        lPoint = "05"
        lSelect = "Exec ged.dbo.proc_CancelaEmissaoSyas " & val(Trim$(lObjstcPed.Num_PI_Emissao)) & ",'" & lUsuarioHora & "'"
        If Not M46V999.gfObtRegistro(lCnRamosDiv, lSelect, lrst, UltimoErro) Then
            UltimoErro = App.EXEName & " - ERRO ao executar procedure de cancelamento. (SyasWeb) PI:" & lObjstcPed.Num_PI_Emissao & "||" & UltimoErro
            Shell "net send tjgomes " & UltimoErro
            Shell "net send mosilva " & UltimoErro
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
        lPoint = "06"
        If lrst("Acao") = 0 Then
            'Cancelamento executado com sucesso
            lIndCancel = 1
        Else
            'Cancelamento n�o executado
            Select Case lrst("Acao")
            Case 70
                lIndCancel = 2
            Case 80, 81
                lIndCancel = 3
            Case 90
                lIndCancel = 4
            End Select
        End If
        If lIndCancel <> 0 Then
            If Not M46V999.gfPreenchido(M46V999.gEndSyasPortal) Then Call M46V999.gpInicia(mPropTabela)
            If M46V999.gfAbrBasDados("SYASPORTAL", M46V999.gTipSyasPortal, M46V999.gEndSyasPortal, lCnSyasPortal2008, UltimoErro) = False Then
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
                GoTo SAIDA
            End If
            lUpdate = "Update SyasPortal.dbo.Tab_Ped_Transm set Ind_Cancel = " & lIndCancel
            lUpdate = lUpdate & " , Cod_Acao = " & IIf(lIndCancel = 1, 80, lrst("acao"))
            lUpdate = lUpdate & " , Dat_Emissao = " & M46V999.gfFormatarCamposSQL(Format(Now, "YYYYMMDD"))
            lUpdate = lUpdate & " where nosso_Numero = " & M46V999.gfFormatarCamposSQL(lObjstcPed.NOSSO_NUMERO)
            lUpdate = lUpdate & " AND ([COD_ACAO] IS NULL OR [COD_ACAO] = 0 OR COD_ACAO = 55)"
            If M46V999.gfExeSQL(lCnSyasPortal2008, lUpdate, UltimoErro) <> 0 Then
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
                GoTo SAIDA
            End If
        End If
        lMsgRetorno = gfProtocoloTransmissao(pNossoNumero, "", tagDynaTrace)
        lretorno = True
    End If
    lPoint = "07"

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

Retorno:
    pMsgRetorno = lMsgRetorno
    gfCancelarEmissao = lretorno

SAIDA:
    GoSub FechaObj
    Err.Clear: On Error GoTo 0
    Exit Function

ControleErro:
    lretorno = False
    lMsgRetorno = "gpCarregaCancelamentoEmissao - Erro - " & lPoint & " - " & Err.Description
    
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMsgRetorno, eSeveridadeErro_Medio, Class_Name, methodName)
    
    GoTo Retorno
FechaObj:
    On Error Resume Next
    Set lrst = Nothing
    Set lObjstcPed = Nothing
    If Not lCnRamosDiv Is Nothing Then
        If lCnRamosDiv.State = adStateOpen Then
            lCnRamosDiv.Close
        End If
    End If: Set lCnRamosDiv = Nothing
    If Not lCnSyasPortal2008 Is Nothing Then
        If lCnSyasPortal2008.State = adStateOpen Then
            lCnSyasPortal2008.Close
        End If
    End If: Set lCnSyasPortal2008 = Nothing
    Return
End Function

Private Function mfCancelarEmissaoOFFLine(ByVal pNossoNumero As String, _
                                          ByRef pMsgRetorno As String, tagDynaTrace As String) As Boolean
    Dim lMensagem As String             'Mensagem.
    Dim lPoint As String
    '================================================================================================
    Dim lHora As Date
    Dim lSelect As String
    Dim lUpdate As String
    Dim lrst As clsYasRecordSet
    Dim lCnSyasPortal2008 As ADODB.Connection
    Dim lUsuario As String
    Dim lUsuarioHora As String
    Dim lObjstcPed As stcA46V702B


    Dim lIndCancel As Integer

    Dim lMsgRetorno As String
    Dim lretorno As Boolean
    Dim lAlertas As String
    Dim lObjA46V101A As clsA46V101A

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfCancelarEmissaoOFFLine"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro: lretorno = False

    If Not mfPosicionaUsuario(tagDynaTrace) Then Exit Function

    Set lObjstcPed = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)

    If lObjstcPed.Cod_Sit_Ped = e_SitPedido_Transmitido Then
        If lObjstcPed.Num_PI_Emissao = 0 Then
            UltimoErro = "FGE2062 - Registro inconsistente: Proposta transmitida sem n�mero de PI."
            GoTo SAIDA
        End If
    Else
        UltimoErro = "AGE0086 - Documento n�o transmitido"
        GoTo SAIDA
    End If

    lUsuario = Me.objstcUsuario.login_user_rede

    lPoint = "01"
    lHora = Time
    If lHora > "19:00" Or lHora < "06:00" Then
        lMensagem = "O cancelamento permitido apenas em hor�rio comercial (06:00 �s 19:00)."
    End If

    lUsuarioHora = "Usu�rio: " & Mid(lUsuario, 1, 6) & " " & Format(Now, "DD/MM/YYYY Hh:Nn:Ss")
    lPoint = "04"

    If M46V999.gfPreenchido(lMensagem) Then
        'Erro
        lMsgRetorno = "CANCELAMENTO N�O EFETUADO. " & vbCrLf & vbCrLf _
                    & "Nosso N�mero .: " & pNossoNumero & vbCrLf _
                    & "Num�ro PI ....: " & lObjstcPed.Num_PI_Emissao & vbCrLf _
                    & "01 ...........: " & lMensagem
    ElseIf M46V999.gfPreenchido(pNossoNumero) = False Or val(Trim$(lObjstcPed.Num_PI_Emissao)) = 0 Then
        lMsgRetorno = "CANCELAMENTO N�O EFETUADO. " & vbCrLf & vbCrLf _
                    & "Nosso N�mero .: " & pNossoNumero & vbCrLf _
                    & "Num�ro PI ....: " & lObjstcPed.Num_PI_Emissao & vbCrLf _
                    & "02 ...........: Par�metros inv�lidos"
    Else
        lPoint = "02"
        If Not M46V999.gfPreenchido(M46V999.gEndSyasPortal) Then Call M46V999.gpInicia(mPropTabela)
        If M46V999.gfAbrBasDados("SYASPORTAL", M46V999.gTipSyasPortal, M46V999.gEndSyasPortal, lCnSyasPortal2008, UltimoErro) = False Then
            GoTo SAIDA
        End If
        'Consulta se pode cancelar
        lPoint = "05"

        lSelect = "Select Cod_Acao from SyasPortal.dbo.Tab_Ped_Transm WITH (NOLOCK)"
        lSelect = lSelect & " where nosso_Numero = " & M46V999.gfFormatarCamposSQL(lObjstcPed.NOSSO_NUMERO)

        If Not M46V999.gfObtRegistro(lCnSyasPortal2008, lSelect, lrst, UltimoErro) Then
            UltimoErro = App.EXEName & " - ERRO ao executar procedure de cancelamento. (SyasWeb) PI:" & lObjstcPed.Num_PI_Emissao & "||" & UltimoErro
            GoTo SAIDA
        End If

        lPoint = "06"
        If Not lrst.EOF Then
            Select Case M46V999.gfTratarNulo(lrst("Cod_Acao"), tpoNumerico)
            Case 70, 80, 81, 90
                'Nao pode Cancelar
                DoEvents
            Case Else
                'Gera o Arquivo de Cancelamento
                lIndCancel = 1

                lUpdate = "Update SyasPortal.dbo.Tab_Ped_Transm set Ind_Cancel = " & lIndCancel
                lUpdate = lUpdate & " , Cod_Acao = " & IIf(lIndCancel = 1, 80, lrst("Cod_Acao"))
                lUpdate = lUpdate & " , Dat_Emissao = " & M46V999.gfFormatarCamposSQL(Format(Now, "YYYYMMDD"))
                lUpdate = lUpdate & " where nosso_Numero = " & M46V999.gfFormatarCamposSQL(lObjstcPed.NOSSO_NUMERO)
                lUpdate = lUpdate & " AND ([COD_ACAO] IS NULL OR [COD_ACAO] = 0 OR COD_ACAO = 55)"
                If M46V999.gfExeSQL(lCnSyasPortal2008, lUpdate, UltimoErro) <> 0 Then
                    GoTo SAIDA
                End If

                Set lObjA46V101A = New clsA46V101A
                Set lObjA46V101A.M46V999 = M46V999

                If lObjA46V101A.gfCriArqCancelamentoEmissao(Me.objstcUsuario.codCorretor, mUsuarioTransmissao, pNossoNumero, lObjstcPed.Num_PI_Emissao, UltimoErro) = False Then
                    GoTo SAIDA
                End If
            End Select
        End If

        lMsgRetorno = gfProtocoloTransmissao(pNossoNumero, "", tagDynaTrace)
        lretorno = True
    End If
    lPoint = "07"
    
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

Retorno:
    pMsgRetorno = lMsgRetorno
    mfCancelarEmissaoOFFLine = lretorno

SAIDA:
    GoSub FechaObj
    Err.Clear: On Error GoTo 0
    Exit Function

ControleErro:
    lretorno = False
    lMsgRetorno = "mfCancelarEmissaoOFFLine - Erro - " & lPoint & " - " & Err.Description
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMsgRetorno, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo Retorno
FechaObj:
    On Error Resume Next
    Set lrst = Nothing
    Set lObjstcPed = Nothing
    If Not lCnSyasPortal2008 Is Nothing Then
        If lCnSyasPortal2008.State = adStateOpen Then
            lCnSyasPortal2008.Close
        End If
    End If: Set lCnSyasPortal2008 = Nothing
    Set lObjA46V101A = Nothing
    Return
End Function

Public Function gfProcuraNossoNumero(ByVal pNossoNumero As String, _
                                     ByVal pEhImpressao As Boolean, _
                                     ByVal pPropTabela As String, _
                                     ByVal tagDynaTrace As String, _
                                     Optional ByVal pOrigem As Integer = 0) As stcA46V702B
    ' Parametros padr�o:
    ' pEhImpressao = false
    ' pPropTabela = ""
                                     
    Dim lclsPedido As clsA46V702A
    Dim lStcPed As stcA46V702B
    Dim lstcBoleto As stcA99V009B
    Dim lDiasValidade As Integer
    Dim lstcPedLoc As stcA46V708B
    Dim RelUsuPrInformado() As String
    Dim lInd As Integer

    '2016-07-11
    Dim pTablePrefixo As String
    pTablePrefixo = ""
    

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfProcuraNossoNumero"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    Call M46V999.gpGraLog(0, "Tag DynaTrace: " & tagDynaTrace & " | KeyDynaTrace " & CStr(lKeyDynaTrace))

    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    M46V111.gNum_Pi = 0

    Set lStcPed = New stcA46V702B
    Set lStcPed.M46V999 = M46V999
    Set lstcPedLoc = New stcA46V708B
    Set lstcPedLoc.M46V999 = M46V999

    lStcPed.NOSSO_NUMERO = pNossoNumero
    Set lclsPedido = New clsA46V702A
    Set lclsPedido.M46V999 = M46V999
    If UCase(pNossoNumero) Like "*PADRAO*" Then
        Call M46V999.mfVerRegPadrao(UltimoErro)
    End If

    If (pPropTabela <> "") Then
        objstcUsuario.PropTabela = pPropTabela + ".dbo."
    End If

    ''2016-07-11: Origem, 0:Nao informado, 1:Kit, 2:Web, 3:FTP
    If pOrigem = 3 Then
        lStcPed.M46V999.gOrigemChamada = pOrigem
        pTablePrefixo = "FTP_"
    Else
            If UCase(objstcUsuario.PropTabela) Like "*SYASPORTAL*" Then
                lStcPed.M46V999.gOrigemChamada = e_Origem_Web
            Else
                 lStcPed.M46V999.gOrigemChamada = e_Origem_Kit
            End If
        pTablePrefixo = ""
    End If

    If Not lclsPedido.mfCarregarColecao(lStcPed, objstcUsuario.PropTabela, pTablePrefixo, True) Then
        UltimoErro = lclsPedido.UltimoErro
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    If lclsPedido.Colecao.Count = 0 Then
        UltimoErro = "FNL0003 - Pedido n�o localizado com nosso n�mero " & pNossoNumero & "."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    Set lStcPed = lclsPedido.Colecao.Item(1)
    Set lstcPedLoc = lStcPed.ColecaoPedLoc(1)

    'OBTEM OS USUARIOS QUE S�O LIBERADOS PARA EMITIR COM PREMIO INFORMADO ABAIXO DO PREMIO MINIMO
    RelUsuPrInformado = Split(M46V999.gRelUsuPrInformado, ";")

    M46V999.gPermitePrInformado = False
    
    For lInd = 0 To UBound(RelUsuPrInformado())
        If RelUsuPrInformado(lInd) <> "" Then
            If UCase(RelUsuPrInformado(lInd)) = UCase(lStcPed.Cod_User_Liber) Then
                M46V999.gPermitePrInformado = True
                Exit For
            End If
        End If
    Next

    lDiasValidade = 0
    'Vila Velha - Nova tarifa 1104
    If lStcPed.objstcCtrlEmis.Cod_Corr = 682987 Then
        If M46V103.mfAjustesPontuais(e_Tip_Valor_ValidaDiasCorr, TpoNumInteiro, lDiasValidade, UltimoErro, lstcPedLoc, 0, tagDynaTrace) = False Then
            GoTo SAIDA
        End If
        lStcPed.M46V999.gValidadeCotCorr = lDiasValidade
    End If


    If pEhImpressao = True Then
        Call lStcPed.gpAtualizaDatPriCal
        If lStcPed.Cod_Sit_Ped = e_SitPedido_Transmitido Then
            Call mfCalcFracionamento(lStcPed, UltimoErro, tagDynaTrace)
            If lStcPed.BoletoGerado = False Or lStcPed.ValidarValorBoletoGerado = False Then
                'Set lstcBoleto = gfRetornaBoleto(lStcPed.Nosso_Numero, lStcPed.ehReentrada, tagDynaTrace)
                Set lstcBoleto = gfRetornaBoletoWS(lStcPed.NOSSO_NUMERO, lStcPed.ehReentrada, tagDynaTrace, lStcPed.Num_PI_Emissao)
                
                If Not lstcBoleto Is Nothing Then
                    'Recarrega o Pedido
                    Set lStcPed = gfProcuraNossoNumero(pNossoNumero, False, pTablePrefixo, tagDynaTrace, pOrigem)
                    If M46V999.gfPreenchido(UltimoErro) Then
                        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
                        GoTo SAIDA
                    End If
                    
                    ' Colocando o boleto no Pedido
                    lStcPed.objstcBoleto.COD_BANCO = 33
                    lStcPed.objstcBoleto.Cod_Corr = lStcPed.ColecaoPedCorr(1).Cod_Corr
                    lStcPed.objstcBoleto.Cod_Ramo = lStcPed.Cod_Ramo
                    lStcPed.objstcBoleto.COD_UNID = lStcPed.Cod_Unid_Prod
                    lStcPed.objstcBoleto.Cod_User = "test11"
                    lStcPed.objstcBoleto.CPF_CNPJ_Segurado = lStcPed.Num_Cgc_Cpf
                    lStcPed.objstcBoleto.DAT_DIGIT = lStcPed.DAT_DIGIT
                    lStcPed.objstcBoleto.DAT_VENC = lstcBoleto.vencimento
                    lStcPed.objstcBoleto.DigBoleto = lStcPed.objstcCtrlEmis.DIG_BLOQUETO
                    lStcPed.objstcBoleto.Nom_Segurado = lstcBoleto.NomeSacado
                    lStcPed.objstcBoleto.NOSSO_NUMERO = lStcPed.NOSSO_NUMERO
                    lStcPed.objstcBoleto.NUM_BOLETO = lStcPed.objstcCtrlEmis.Num_Bloq
                    lStcPed.objstcBoleto.NUM_LINHA_DIGIT = lstcBoleto.linhaDigitavel
                    lStcPed.objstcBoleto.STATUS = 3
                    lStcPed.objstcBoleto.VAL_PR = lstcBoleto.valorDocumento
                    lStcPed.objstcBoleto.Vigencia = lstcBoleto.Vigencia
                Else
                    UltimoErro = "Falha ao carregar boleto(3):NN=" & pNossoNumero & "|Prefixo=" & pTablePrefixo & "|Origem=" & CStr(pOrigem) & "|" & UltimoErro
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
                    GoTo SAIDA
                End If
            End If
        End If
    End If

    If lStcPed.ColecaoPedLoc.Count = 0 Then
        UltimoErro = "FNL0004 - Item do pedido n�o localizado com nosso n�mero " & pNossoNumero & "."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If


    M46V999.gNovaPropostaLMI = False
    M46V111.gNumPedido = lStcPed.Num_Ped
    M46V111.gNum_Pi = lStcPed.NUM_PI

    'c) Prefixo das mensagens.
    Select Case lStcPed.Cod_Tip_Doc
    Case 1
        M46V999.gPrefixo = "Cota��o "
    Case 2
        M46V999.gPrefixo = "Proposta "
    End Select
    M46V999.gPrefixo = M46V999.gPrefixo & Format$(M46V111.gNumPedido, "000000") & " - "

    'd)  C�digo do Dizer
    M46V111.gCod_Dizer = lStcPed.Cod_Dizer

    M46V999.gCod_Ramo = lStcPed.Cod_Ramo
    M46V111.gTipoGrupMulti = e_GrupoRamo_NDefinido

    If M46V999.gCod_Ramo = e_CodRamo_Empresarial Or _
       M46V999.gCod_Ramo = e_CodRamo_Condominio Or _
       M46V999.gCod_Ramo = e_CodRamo_Residencial Or _
       M46V999.gCod_Ramo = e_CodRamo_Riscos_diversos Or _
       M46V999.gCod_Ramo = e_CodRamo_Roubos Then
        M46V111.gTipoGrupMulti = e_GrupoRamo_Multi
    ElseIf M46V999.gCod_Ramo = e_CodRamo_D_P_E_M Or _
           M46V999.gCod_Ramo = e_CodRamo_Ap_bilhete Then
        M46V111.gTipoGrupMulti = e_GrupoRamo_570_820
    ElseIf M46V999.gCod_Ramo = e_CodRamo_Automoveis Then
        M46V111.gTipoGrupMulti = e_GrupoRamo_Auto
    ElseIf M46V999.gCod_Ramo = e_CodRamo_Acidentes_pessoais Then
        M46V111.gTipoGrupMulti = e_GrupoRamo_Acidentes_pessoais
    ElseIf M46V999.gCod_Ramo = e_CodRamo_Transp_internacional Or _
           M46V999.gCod_Ramo = e_CodRamo_Transp_nacional Or _
           M46V999.gCod_Ramo = e_CodRamo_Rc_Transportador Or _
           M46V999.gCod_Ramo = e_CodRamo_Rcf_dc Then
        M46V111.gTipoGrupMulti = e_GrupoRamo_Transp_Diversos
    ElseIf M46V999.gCod_Ramo <> e_CodRamo_NaoDefinido Then
        M46V111.gTipoGrupMulti = e_GrupoRamo_Generico
    End If


    If mfCalcFracionamento(lStcPed, UltimoErro, tagDynaTrace) = False Then
        GoTo ControleErro
    End If

    'Normaliza a estrutura do Pedido
    If mfNormalizaPedido(lStcPed, UltimoErro, tagDynaTrace) = False Then
        GoTo ControleErro
    End If
    
    'Obtem dados de transa��o do cart�o, caso cart�o de cr�dito
    If lStcPed.Cod_Bco_Cobr = 999 And lStcPed.Cod_Agen_Bas_Cor = 9999 And lStcPed.Num_Cta_Bas_Cor = 9999999 Then
        lStcPed.TransactionId = gfObterDadosTransacaoCartao(lStcPed, tagDynaTrace)
    Else
        lStcPed.TransactionId = ""
    End If

    Set gfProcuraNossoNumero = lStcPed

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    UltimoErro = "[ERR] - gfProcuraNossoNumero ERRO - " & Err & " " & Err.Description & UltimoErro
    
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    
    GoTo SAIDA

FechaObj:
    Set lclsPedido = Nothing
    Set lStcPed = Nothing
    Set lstcBoleto = Nothing
    Return
End Function

Public Function gfObterDadosTransacaoCartao(ByRef lobjstcPedido As stcA46V702B, tagDynaTrace As String) As String
    
    Dim lConexaoDiv As ADODB.Connection
    Dim lSelect As String
    Dim lrst As clsYasRecordSet
    Dim lTransactionId As String

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfObterDadosTransacaoCartao"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""

    On Error GoTo ControleErro

    '- Abre as bases que ser�o utilizadas
    If Not M46V999.gfAbrBasDados("RamosDiversos", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, lConexaoDiv, UltimoErro) Then
        GoTo ControleErro
    End If

    lSelect = "SELECT ISNULL(TransactionId, '') AS TransactionId FROM Tab_Controle_Cartao "
    lSelect = lSelect & "WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO)
    
    If Not M46V999.gfObtRegistro(lConexaoDiv, lSelect, lrst, UltimoErro) Then
        GoTo ControleErro
    End If
    
    If Not lrst.EOF Then
        lTransactionId = lrst("TransactionId")
    Else
        lTransactionId = ""
    End If
    
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    
    gfObterDadosTransacaoCartao = lTransactionId

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    If Err <> 0 Then
        UltimoErro = UltimoErro & " - " & "Ln: " & Erl & " Erro: " & Err & " " & Err.Description
    End If

    '###### 11.44 TRATA ERRO PARA O USUARIO
    If InStr(UltimoErro, "<") > 0 Then
        UltimoErro = Left(UltimoErro, InStr(UltimoErro, "<") - 1)
        If Len(UltimoErro) - (InStr(UltimoErro, "Mensagem de retorno:") + 20) > 0 Then
            UltimoErro = Right(UltimoErro, Len(UltimoErro) - (InStr(UltimoErro, "Mensagem de retorno:") + 20))
            UltimoErro = Replace(UltimoErro, "##1", Empty)
            UltimoErro = Replace(UltimoErro, "Input:", Empty)
        End If
    End If

    If Err > 0 Then
        UltimoErro = UltimoErro & Err.Description & "Ln: " & Erl
    End If
    '###### FIM 11.44 TRATA ERRO PARA O USUARIO

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)

    GoTo SAIDA

FechaObj:
    Set lConexaoDiv = Nothing
    Set lrst = Nothing
    
    Return
    
End Function

Private Function mfPosicionaUsuario(tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfPosicionaUsuario"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    On Error GoTo ControleErros
    mfPosicionaUsuario = False
    UltimoErro = ""

    If Me.objstcUsuario Is Nothing Then
        UltimoErro = "[AUT] - Usu�rio n�o logado."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
        Exit Function
    End If

    If Not M46V999.gfPreenchido(Me.objstcUsuario.login_user_rede) Then
        UltimoErro = "[AUT] - Usu�rio n�o logado."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
        Exit Function
    End If

    Call M46V999.gpGraLog(0, "*****Usu�rio logado***** " & Me.objstcUsuario.login_user_rede)

    Call M46V999.gpGraLog(0, Me.objstcUsuario.TipoUsuario & "/" & Me.objstcUsuario.TipoAcesso & "/" & "/" & Me.objstcUsuario.codCorretor & _
                             "/" & Me.objstcUsuario.login_user_rede & "/" & Me.objstcUsuario.cod_user_msp & "/" & Me.objstcUsuario.prior_multi)


    M46V999.gTipUsuario = Me.objstcUsuario.TipoUsuario
    M46V999.gTipoAcesso = Me.objstcUsuario.TipoAcesso
    M46V999.gCod_Corretor = Me.objstcUsuario.codCorretor
    M46V999.gNomUsuario = Me.objstcUsuario.login_user_rede

    M46V999.PerfilAcesso = objstcUsuario.prior_multi
    M46O001.PermiteAlterarTaxaFinal = IIf(objstcUsuario.prior_m_taxa = 1, True, False)
    M46O001.PermiteQQDescontoComercial = IIf(objstcUsuario.prior_m_dcoml = 1, True, False)
    M46O001.PermiteDigitarSemNormas = IIf(objstcUsuario.prior_m_normas = 1, True, False)
    M46O001.PermiteAlterarFranquia = IIf(objstcUsuario.prior_m_franquia = 1, True, False)
    M46O001.PermiteSemSiscota = IIf(objstcUsuario.prior_m_siscota = 1, True, False)

    If M46O001.PermiteAlterarFranquia = True Then
        M46O001.gAlterarFranquia = True
    Else
        M46O001.gAlterarFranquia = False
    End If

    If M46O001.PermiteSemSiscota = True Then
        M46O001.gPermiteSemSiscota = True
    Else
        M46O001.gPermiteSemSiscota = False
    End If

    M46V999.gNomUsuario = Me.objstcUsuario.login_user_rede
    M46V999.gNomUsuMainFrame = Me.objstcUsuario.cod_user_msp
    M46V999.gNomUsuario6 = Left$(Me.objstcUsuario.login_user_rede, 6)

    M46V111.gNum_Pi = 0

    Call M46V999.gpGraLog(0, "TipoUsuario=" & objstcUsuario.TipoUsuario & "/TipoAcesso=" & objstcUsuario.TipoAcesso & _
                             "/Corretor=" & M46V999.gCod_Corretor & "/CodUserMsp=" & M46V999.gNomUsuMainFrame & "/prior_multi=" & Me.objstcUsuario.prior_multi & "/" & objstcUsuario.PropTabela)

    mfPosicionaUsuario = True

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    
    Exit Function
ControleErros:

    UltimoErro = "[ERR] - mfPosicionaUsuario " & Err & " " & Err.Description & UltimoErro
    Call M46V999.gpGraLog(0, UltimoErro)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function


Public Function gfProRenYasuda(ByVal pNumRenApolice As String, ByVal pNumRenItem As Integer, _
                               ByVal pItemRenovado As Boolean, ByRef pMsgAlerta As String, tagDynaTrace As String) As stcA46V702B
'Fun��o: processa op��o de renova��o de seguro Yasuda.

    Dim lIndSinistro As e_SimNao
    Dim lNossoNumero As String
    Dim lStcPed As stcA46V702B
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfProRenYasuda"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(M46V999.gDataCorteSompo) Then
            If CLng(M46V999.gDataCorteSompo) <= M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If
    
    Set gfProRenYasuda = Nothing

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    If UCase(objstcUsuario.PropTabela) Like "*SYASPORTAL*" Then
        M46V999.gOrigemChamada = e_Origem_Web
    Else
         M46V999.gOrigemChamada = e_Origem_Kit
    End If

    '1. Consiste n�mero de ap�lice anterior.
    If Not M46V999.gfPreenchido(pNumRenApolice) Then
        UltimoErro = "FGE0069 - N�mero da ap�lice inv�lido (renova��o " & lNomeSeguradora & " - n�o preenchido)."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If
    If Not IsNumeric(pNumRenApolice) Then
        UltimoErro = "FGE0069 - N�mero da ap�lice inv�lido (renova��o " & lNomeSeguradora & " - n�o num�rico)."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If
    If Len(Trim(pNumRenApolice)) < 10 Then
        pNumRenApolice = Trim$(pNumRenApolice)
        pNumRenApolice = String(10 - Len(pNumRenApolice), "0") & pNumRenApolice
    End If

    '2. Consiste n�mero do item da ap�lice.
    If pItemRenovado = False Then  'Apenas um item.
        If Not M46V999.gfPreenchido(pNumRenItem) Then
            UltimoErro = "FGE0070 - N�mero do item da ap�lice inv�lido (renova��o " & lNomeSeguradora & " - n�o preenchido)."
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
        If Not IsNumeric(pNumRenItem) Then
            UltimoErro = "FGE0070 - N�mero do item da ap�lice inv�lido (renova��o " & lNomeSeguradora & " - n�o num�rico)."
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
        If Len(Trim(pNumRenItem)) < 4 Then
            pNumRenItem = Trim$(pNumRenItem)
            pNumRenItem = String(4 - Len(pNumRenItem), "0") & pNumRenItem
        End If
    End If
    'lIndSinistro = pIndSinistro
    lIndSinistro = 0
    '3. Grava registros de cota��o/proposta para renova��o Yasuda.
    If M46V001.gfGraRenYasuda(pNumRenApolice, pNumRenItem, lNossoNumero, _
                              UltimoErro, pMsgAlerta, tagDynaTrace) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    Set lStcPed = gfProcuraNossoNumero(lNossoNumero, False, "", tagDynaTrace)

    Set gfProRenYasuda = lStcPed

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    UltimoErro = "[ERR] - gfProRenYasuda ERRO - " & Err & " " & Err.Description & UltimoErro
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA
FechaObj:
    Set lStcPed = Nothing
    Return
End Function

Private Sub mpPosLinha(ByRef rtbTexto As String, ByRef pQuaLinhas As Long, ByVal pMensagem As String)
'Procedure: posiciona linha de mensagem.

    Dim lAuxiliar As String    'Auxiliar.

    lAuxiliar = pMensagem
    If Len(lAuxiliar) < 81 Then
        If M46V999.gfPreenchido(rtbTexto) Then
            rtbTexto = rtbTexto & vbLf & lAuxiliar
        Else
            rtbTexto = lAuxiliar
        End If
        pQuaLinhas = pQuaLinhas + 1
    Else
        If M46V999.gfPreenchido(rtbTexto) Then
            rtbTexto = rtbTexto & vbLf & Left$(lAuxiliar, 77) & "..."
        Else
            rtbTexto = Left$(lAuxiliar, 77) & "..."
        End If
        pQuaLinhas = pQuaLinhas + 1
        lAuxiliar = Mid$(lAuxiliar, 78)
        Do
            If Len(lAuxiliar) < 77 Then
                rtbTexto = rtbTexto & vbLf & "..." & lAuxiliar
                pQuaLinhas = pQuaLinhas + 1
                Exit Do
            End If
            rtbTexto = rtbTexto & vbLf & "..." & Left$(lAuxiliar, 74) & "..."
            pQuaLinhas = pQuaLinhas + 1
            lAuxiliar = Mid$(lAuxiliar, 5)
        Loop
        rtbTexto = rtbTexto & vbLf & " "
        pQuaLinhas = pQuaLinhas + 1
    End If

End Sub

Public Function gfPesProTransmissao(pNumPedido As String, pSegurado As String, _
                                    pIniVigencia As String, pTerVigencia As String, tagDynaTrace As String) As Collection
'Procedure: pesquisa protocolo de transmiss�o.

    Dim lData As String    'Campo auxiliar: data.
    Dim lDir As String    'Utilizado com a fun��o DIR.
    Dim lobjTransm As clsA46V718A    'Objeto da classe clsA46V718A (TAB_PED_TRANSM)
    Dim lobjstcTransm As stcA46V718B    'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM)

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfPesProTransmissao"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '1. Instancio os objetos
    Set lobjTransm = New clsA46V718A
    Set lobjTransm.M46V999 = M46V999

    Set lobjstcTransm = New stcA46V718B


    '2. Carrego os par�metro de pesquisa.
    'Obs.: Se n�mero do pedido est� preenchido, utiliza apenas esta informa��o no SELECT.
    If M46V999.gfPreenchido(pNumPedido) Then
        lobjstcTransm.Num_Ped = Format$(pNumPedido, "000000")
    Else
        If M46V999.gfPreenchido(pSegurado) Then
            lobjstcTransm.Nom_Segurado = M46V999.gfForAspas(UCase(pSegurado))
        End If
        'b) Vig�ncia.
        If M46V999.gfPreenchido(pIniVigencia) Then
            lData = M46V999.gfForData(1, pIniVigencia)
            lobjstcTransm.Dat_Ini_Vig = lData
        End If
        If M46V999.gfPreenchido(pTerVigencia) Then
            lData = M46V999.gfForData(1, pTerVigencia)
            lobjstcTransm.Dat_Ter_Vig = lData
        End If
    End If

    '3. Obt�m registros de protocolo.
    If Not lobjTransm.mfCarregarColecao(lobjstcTransm, objstcUsuario.PropTabela, "") Then
        UltimoErro = lobjTransm.UltimoErro
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '4. Se n�o encontrou registro, informa usu�rio.
    If lobjTransm.Colecao.Count = 0 Then
        lDir = Dir("*.Y01")
        If M46V999.gfPreenchido(lDir) Then
            'Call mnuExiProAntigos_Click
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Registro n�o encontrado", eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        Else
            Call mpMenNaoEncontrado("", pNumPedido, pSegurado, pIniVigencia, pTerVigencia, "", 3)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Registro n�o encontrado", eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If

    Set gfPesProTransmissao = lobjTransm.Colecao
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    UltimoErro = "[ERR] - mpPesProTransmissao " & Err & " " & Err.Description & UltimoErro
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lobjstcTransm = Nothing
    Set lobjTransm = Nothing
    Return
End Function



Private Sub mpMenNaoEncontrado(pCodCorretor As String, pNumPedido As String, _
                               pSegurado As String, pIniVigencia As String, _
                               pTerVigencia As String, pCodProdutor As String, _
                               pTipDocumento As Integer)

    UltimoErro = "FNL0011 - N�o localizado registro que satisfa�a os crit�rios selecionados." & vbLf
    If M46V999.gfPreenchido(pNumPedido) Then
        UltimoErro = UltimoErro & vbLf & "N�mero do pedido = " & pNumPedido & "."
    Else
        If pTipDocumento = Enum_TipDocumentoCot Then
            UltimoErro = UltimoErro & vbLf & "Tipo de documento = cota��es e propostas"
            UltimoErro = UltimoErro & "."
        Else
            If pTipDocumento = Enum_TipDocumentoApol Then
                UltimoErro = UltimoErro & vbLf & "Tipo de documento = ap�lices vigentes."
            Else
                UltimoErro = UltimoErro & vbLf & "Tipo de documento = protocolo de transmiss�o."
            End If
        End If
        If M46V999.gfPreenchido(pSegurado) Then
            UltimoErro = UltimoErro & vbLf & "Nome do segurado que contenha """ & pSegurado & """."
        End If
        If M46V999.gfPreenchido(pCodProdutor) Then
            UltimoErro = UltimoErro & vbLf & "Nome do produtor que contenha """ & pCodProdutor & """."
        End If
        If M46V999.gfPreenchido(pCodCorretor) Then
            UltimoErro = UltimoErro & vbLf & "Corretor = " & pCodCorretor & "."
        End If
        If M46V999.gfPreenchido(pIniVigencia) And M46V999.gfPreenchido(pTerVigencia) Then
            UltimoErro = UltimoErro & vbLf & "Vig�ncia: de " & pIniVigencia & " a " & _
                         pTerVigencia & "."
        Else
            If M46V999.gfPreenchido(pIniVigencia) Then
                UltimoErro = UltimoErro & vbLf & "Vig�ncia em " & pIniVigencia & "."
            Else
                If M46V999.gfPreenchido(pTerVigencia) Then
                    UltimoErro = UltimoErro & vbLf & "Vig�ncia em " & pTerVigencia & "."
                End If
            End If
        End If

        'If pCodRamo <> e_CodRamo_NaoDefinido Then
        '    UltimoErro = UltimoErro & vbLf & "Ramo =" & lCodRamo & "."
        'End If
    End If
End Sub

Public Function gfPesApolices(ByVal pCodCorretor As String, _
                              ByVal pSegurado As String, _
                              ByVal pCodProdutor As String, ByVal _
                              pCodRamo As Integer, tagDynaTrace As String) As Collection
'Procedure: pesquisa cota��es e propostas.
    Dim lListaNossoNumero As String         'Lista de nosso_numero.
    Dim lrst As clsYasRecordSet    'Registro de ap�lice.
    Dim bdP00APOL As ADODB.Connection    'Arquivo de ap�lice (Conex�o).
    Dim lCpfCnpj As String    'Retorno da fun��o de forma��o de cpf e cnpj.
    Dim lCod_Corr As String    'Codigo do corretor.
    Dim lRegistro As String    'Registro.
    Dim lNomeTab As String
    Dim lObjstcA46O001B As stcA46O001B
    Dim lCollection As New Collection
    Dim lSelect As String
    Dim i As Integer
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfPesApolices"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", 0, Class_Name, methodName)
        GoTo SAIDA
    End If


    '2. Abre base de dados
    If M46V999.gfAbrBasDados("P00APOL", M46V999.gTipP00APOL, M46V999.gEndP00APOL, bdP00APOL, UltimoErro) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    'Busco os dados da ap�lice

    'Ficha 1000782 - N�o permitir Renova��o LMI para Corretores
    lSelect = "select p.* , NOM_LOGR_RISCO, NUM_ITEM From"
    lSelect = lSelect & " ged.dbo.tab_ped_loc_atual as l WITH (NOLOCK) inner join ged.dbo.tab_ped_Atual as p WITH (NOLOCK)"
    lSelect = lSelect & "     on l.num_apol = p.num_apol and l.num_endo = p.num_endo"
    lSelect = lSelect & "     inner join ged.dbo.tab_ped_corr_atual as c WITH (NOLOCK)"
    lSelect = lSelect & "     on c.num_apol = p.num_apol and c.num_endo = p.num_endo"


    'Nome do segurado.
    If M46V999.gfPreenchido(pSegurado) Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Nom_Segurado like '%" & UCase(pSegurado) & "%'"
    End If

    'Produtor.
    If M46V999.gfPreenchido(pCodProdutor) Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Prod = " & pCodProdutor
    End If

    If pCodRamo <> e_CodRamo_NaoDefinido Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " p.Cod_Ramo = " & pCodRamo
    End If

    If val(pCodCorretor) <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " c.Cod_Corr = " & pCodCorretor
    End If

    If objstcUsuario.TipoUsuario <> e_Tipo_Usu_Yasuda Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_corr in (SELECT * FROM [Ged].[dbo].[Func_BuscaCorrGrupoAtivo] ('" & Format(Me.objstcUsuario.codCorretor, "0000000") & "', 112))"
    End If
    'Ficha 1000782 - N�o permitir Renova��o LMI para Corretores
    If objstcUsuario.TipoUsuario <> e_Tipo_Usu_Yasuda Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " NOM_LOGR_RISCO not like '%CONFORME ESPEC%'"
    End If

    lSelect = M46V999.gfWhereAnd(lSelect) & " cod_dizer not in (4,10) "
    lSelect = lSelect & " ORDER BY p.num_apol"


    '4. obtenho o registro.
    If Not M46V999.gfObtRegistro(bdP00APOL, lSelect, lrst, UltimoErro) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '5. Se n�o encontrou registro, informa usu�rio.
    If lrst.EOF Then
        Call mpMenNaoEncontrado(pCodCorretor, "", pSegurado, "", "", pCodProdutor, 2)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, "N�o retornou nenhum registro", 0, Class_Name, methodName)
        GoTo SAIDA
    End If

    Set lCollection = New Collection
    Do While Not lrst.EOF
        'b) Coluna 0 - Excluir - M46V999.CheckBox para marcar itens a serem exclu�dos.
        Set lObjstcA46O001B = New stcA46O001B
        If lrst("Cod_Ramo") = e_CodRamo_Condominio Then
            lObjstcA46O001B.Dsc_Ramo = "Condom�nio"
        ElseIf lrst("Cod_Ramo") = e_CodRamo_Empresarial Then
            lObjstcA46O001B.Dsc_Ramo = "Empresarial"
        ElseIf lrst("Cod_Ramo") = e_CodRamo_Residencial Then
            lObjstcA46O001B.Dsc_Ramo = "Residencial"
        Else
            lObjstcA46O001B.Dsc_Ramo = "N�o definido"
        End If
        lObjstcA46O001B.Num_Apol = lrst("Num_Apol")
        lObjstcA46O001B.Num_Item = lrst("Num_Item")
        lObjstcA46O001B.NomeSegurado = lrst("Nom_Segurado")
        Select Case Len(CStr(lrst("Num_Cgc_Cpf")))
        Case Is > 11
            If IsNumeric(lrst("Num_Cgc_Cpf")) Then
                Call M46V999.gpForCNPJ(1, Format$(lrst("Num_Cgc_Cpf"), "00000000000000"), lCpfCnpj)
            Else
                Call M46V999.gpForCNPJ(1, Right$(String(14, "0") & Trim(lrst("Num_Cgc_Cpf")), 14), lCpfCnpj)
            End If
        Case Else
            Call M46V999.gpForCPF(1, Format$(lrst("Num_Cgc_Cpf"), "00000000000"), lCpfCnpj)
        End Select
        lObjstcA46O001B.NumCNPJCPF = lCpfCnpj
        lObjstcA46O001B.IniVigencia = lrst("Dat_Ini_Vig")
        lObjstcA46O001B.TerVigencia = lrst("Dat_Ter_Vig")
        lObjstcA46O001B.CodModalidade = lrst("Cod_Estip")
        lObjstcA46O001B.Cod_Ramo = lrst("Cod_Ramo")
        
        For i = 1 To lCollection.Count
            If lCollection(i).Class_ID = lObjstcA46O001B.Class_ID Then
                GoTo Proximo
            End If
        Next
        
        lCollection.Add lObjstcA46O001B, lObjstcA46O001B.Class_ID
        
Proximo:
        lrst.MoveNext
    Loop

    Set gfPesApolices = lCollection

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    UltimoErro = "[ERR] - mpPesApolice " & Err & " " & Err.Description & UltimoErro
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lrst = Nothing
    Set bdP00APOL = Nothing
    Set lObjstcA46O001B = Nothing
    Set lCollection = Nothing
    Return

End Function

Public Function gfPesCotPropostas(ByVal pCodCorretor As String, _
                                  ByVal pNumPedido As String, _
                                  ByVal pSegurado As String, _
                                  ByVal pIniVigencia As String, _
                                  ByVal pTerVigencia As String, _
                                  ByVal pCodProdutor As String, _
                                  ByVal pSitDocumento As Integer, _
                                  ByVal pCodRamo As Integer, tagDynaTrace As String) As Collection
'Procedure: pesquisa cota��es e propostas.

    Dim lListaNossoNumero As String         'Lista de nosso_numero.
    Dim lCount As Double    'Contador.
    Dim lI As Double    'Auxiliar.
    Dim lColuna As Integer    'Coluna do grid.
    Dim lObjPed As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED).
    Dim lObjstcPed As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED).
    Dim lObjPedCorr As clsA46V705A    'Objeto da classe clsA46V705A (TAB_PED_CORR).
    Dim lobjstcPedCorr As stcA46V705B    'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).
    Dim lobjstcPedLoc As stcA46V708B

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfPesCotPropostas"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro
    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '2. Busco os registro com o corretor informado  (se informado).
    lListaNossoNumero = ""
    If M46V999.gfPreenchido(pCodCorretor) Then
        Set lObjPedCorr = New clsA46V705A
        Set lObjPedCorr.M46V999 = M46V999

        Set lobjstcPedCorr = New stcA46V705B
        Set lobjstcPedCorr.M46V999 = M46V999

        lobjstcPedCorr.Cod_Corr = pCodCorretor
        If Not lObjPedCorr.mfCarregarColecao(lobjstcPedCorr, objstcUsuario.PropTabela, "", lListaNossoNumero) Then
            UltimoErro = lObjPedCorr.UltimoErro
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
    ElseIf objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda Then
        Set lObjPedCorr = New clsA46V705A
        Set lObjPedCorr.M46V999 = M46V999

        Set lobjstcPedCorr = New stcA46V705B
        Set lobjstcPedCorr.M46V999 = M46V999

        lobjstcPedCorr.Cod_Corr = M46V999.gCod_Corretor
        If Not lObjPedCorr.mfCarregarColecao(lobjstcPedCorr, objstcUsuario.PropTabela, "", lListaNossoNumero) Then
            UltimoErro = lObjPedCorr.UltimoErro
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If

    '3. Instancio os objetos
    Set lObjPed = New clsA46V702A
    Set lObjPed.M46V999 = M46V999

    Set lObjstcPed = New stcA46V702B
    Set lObjstcPed.M46V999 = M46V999

    If pCodRamo <> e_CodRamo_NaoDefinido Then
        lObjstcPed.Cod_Ramo = pCodRamo
    End If

    'N�mero do pedido.
    If M46V999.gfPreenchido(pNumPedido) Then
        lObjstcPed.Num_Ped = pNumPedido
    End If
    'Situa��o do documento.
    If pSitDocumento = Enum_SitDocumentoNaoTran Then       'Propostas n�o transmitidas.
        lObjstcPed.Cod_Sit_Ped = e_SitPedido_Nenhum
    Else
        If pSitDocumento = Enum_SitDocumentoTran Then        'Propostas transmitidas.
            lObjstcPed.Cod_Sit_Ped = e_SitPedido_Transmitido
        End If
    End If
    'Nome do segurado.
    If M46V999.gfPreenchido(pSegurado) Then
        lObjstcPed.Nom_Segurado = pSegurado
    End If
    'Vig�ncia.
    If M46V999.gfPreenchido(pIniVigencia) And pIniVigencia <> "0" Then
        lObjstcPed.Dat_Ini_Vig = M46V999.gfForData(1, pIniVigencia)
    End If
    If M46V999.gfPreenchido(pTerVigencia) And pTerVigencia <> "0" Then
        lObjstcPed.Dat_Ter_Vig = M46V999.gfForData(1, pTerVigencia)
    End If
    'Produtor.
    If M46V999.gfPreenchido(pCodProdutor) Then
        lObjstcPed.Cod_Prod = pCodProdutor
    End If
    lObjstcPed.Num_Apol = 9999999999#
    lObjstcPed.NossoNumeroIn = lListaNossoNumero

    If objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda Then
        lObjstcPed.Cod_Segurado = M46V999.gCodSeguradoYas
    End If

    '4. Busco os dados com os par�metros de pesquisa informados.
    If Not lObjPed.mfCarregarColecao(lObjstcPed, objstcUsuario.PropTabela, "", False) Then
        UltimoErro = lObjPed.UltimoErro
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '5. Se n�o encontrou registro, informa usu�rio.
    If lObjPed.Colecao.Count = 0 Then
        Call mpMenNaoEncontrado(pCodCorretor, pNumPedido, pSegurado, pIniVigencia, pTerVigencia, pCodProdutor, 1)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Nenhum registro encontrado", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    Set gfPesCotPropostas = lObjPed.Colecao

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    GoSub FechaObj
    Err.Clear: On Error GoTo 0
    Exit Function

ControleErro:
    UltimoErro = "[ERR] - mpPesCotProposta " & Err & " " & Err.Description & UltimoErro
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lObjPed = Nothing
    Set lObjstcPed = Nothing
    Set lObjPedCorr = Nothing
    Set lobjstcPedCorr = Nothing
    Set lobjstcPedLoc = Nothing
    Return
End Function

Public Sub gfAutenticarUsuario(pUsuario As String, tagDynaTrace As String)
    Dim objTabUser As clsA46V725A
    Dim objstcTabUser As stcA46V725B

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfAutenticarUsuario"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    M46V999.gNomUsuario = pUsuario

    If Trim(pUsuario) = "" Then
        UltimoErro = "[AUT] - Usu�rio desconhecido"
        Set Me.objstcUsuario = Nothing
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Baixo, Class_Name, methodName)
        GoTo SAIDA
    End If

    'Verifico se � usu�rio interno.
    Set objTabUser = New clsA46V725A
    Set objTabUser.M46V999 = M46V999

    Set objstcTabUser = New stcA46V725B
    objstcTabUser.login_user_rede = UCase(pUsuario)

    If Not objTabUser.mfCarregarColecao(objstcTabUser) Then
        If M46V999.gfPreenchido(objTabUser.UltimoErro) Then
            UltimoErro = objTabUser.UltimoErro
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Baixo, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If

    If objTabUser.Colecao.Count <> 0 Then
        Set objstcTabUser = objTabUser.Colecao.Item(1)
        objstcTabUser.PropTabela = mPropTabela
        M46V999.gTipUsuario = e_Tipo_Usu_Yasuda
        objstcTabUser.TipoUsuario = M46V999.gTipUsuario
    Else
        If Not mfCarregaUsuarioPortal(UltimoErro, tagDynaTrace) Then
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
            GoTo SAIDA
        End If
        Set objstcTabUser = New stcA46V725B
        objstcTabUser.login_user_rede = pUsuario
        objstcTabUser.codCorretor = M46V999.gCod_Corretor
        M46V999.gTipUsuario = e_Tipo_Usu_Corretor
        objstcTabUser.TipoUsuario = M46V999.gTipUsuario
        objstcTabUser.TipoAcesso = 2
        objstcTabUser.PropTabela = mPropTabela
    End If

    Set Me.objstcUsuario = objstcTabUser

    'Call M46V999.mfVerRegPadrao(UltimoErro)

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Sub
ControleErro:

    UltimoErro = "[ERR] - gfAutenticarUsuario " & Err & " " & Err.Description & UltimoErro
    
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    
    GoTo SAIDA

FechaObj:
    Set objTabUser = Nothing
    Set objstcTabUser = Nothing
    Return
End Sub

Public Function gfSeguroNovo(ByVal pCod_Ramo As Integer, _
                            pCodSegmentacao As Integer, _
                            pCodPlano As Long, tagDynaTrace As String) As stcA46V702B

    Dim lStcPed As stcA46V702B
    Dim lMsgErr As String

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfSeguroNovo"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    Set gfSeguroNovo = Nothing

    Dim lNossoNumero As String

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    If Not IsNumeric(pCod_Ramo) Then
        lMsgErr = "C�digo de Ramo n�o preenchido."
        GoTo ControleErro
    End If

    If pCod_Ramo = 0 Then
        lMsgErr = "C�digo de Ramo zerado."
        GoTo ControleErro
    End If

    If M46V001.gfSegNovo(lNossoNumero, UltimoErro, pCod_Ramo, pCodSegmentacao, pCodPlano, tagDynaTrace) = True Then
        Set lStcPed = gfProcuraNossoNumero(lNossoNumero, False, "", tagDynaTrace)
        
        If lStcPed.Cod_Unid_Prod = 0 Then
            If Not mfPosUniProProdutor(lStcPed, lMsgErr) Then
                GoTo ControleErro
                End If
        End If
    End If
    
    Set gfSeguroNovo = lStcPed
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    If M46V999.gfPreenchido(lMsgErr) = False Then
        lMsgErr = "ERRO - " & Err & " " & Err.Description
    End If
    UltimoErro = "[ERR] - M46V001.gfSegNovo - " & lMsgErr
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lStcPed = Nothing
    Return

End Function


Public Function gfProRenCongenere(ByVal pCodSeguradora As Integer, _
                                  ByVal pIndSinistro As Boolean, _
                                  ByVal pNumApolice As String, _
                                  ByVal pNumItem As String, _
                                  ByVal pCodRamo As e_CodRamo, _
                                  pCodSegmentacao As Integer, _
                                  pCodPlano As Long, tagDynaTrace As String) As stcA46V702B
'Fun��o: processa op��o de renova��o de cong�nere.

    Dim lCodIdeApolice As String          'C�digo de identifica��o da ap�lice.
    Dim lCodProduto As Integer      'C�digo do produto.
    Dim lCodProtocolo As Long           'C�digo do protocolo.
    Dim lCodSequencia As Long           'C�digo da seq��ncia.
    Dim lCodSeguradora As Integer         'C�digo da seguradora.
    Dim lDescTarifa As String       'Descri��o da tarifa.
    Dim lDigito As String   'D�gito.
    Dim lIndSinistro As Byte          'Indicador de sinistro.
    Dim lNosNumero As String      'Nosso n�mero.
    Dim lNumApolice As String       'N�mero da ap�lice.
    Dim lNumItem As String    'N�mero do item.
    Dim lNumPedido As String      'N�mero do pedido.
    Dim lPosicao As Byte      'Utilizado com InStr.
    Dim lVigDatIni As Double      'Data de vig�ncia: in�cio.
    Dim lVigDatTer As Double      'Data de vig�ncia: t�rmino.
    Dim lI As Integer    'Utilizafo no for...next.
    Dim bdP00MPED As ADODB.Connection    'Conex�o.
    Dim lStcPed As stcA46V702B

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfProRenCongenere"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    Set gfProRenCongenere = Nothing

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '1000347 - Inclus�o de informa��es da cong�nere em renova��es --Inicio
    'Posiciona seguradora.
    'If cboSeguradora.ListIndex = -1 Then
    '    Call M46V999.gpGraLog(0, "FGE0094 - Seguradora inv�lida (n�o selecionada).")
    '    If cboSeguradora.Visible = True And cboSeguradora.Enabled = True Then
    '        cboSeguradora.SetFocus
    '    End If
    '    GoTo SAIDA
    'End If
    'lCodSeguradora = cboSeguradora.ItemData(cboSeguradora.ListIndex)

    If pCodSeguradora = 0 Then
        UltimoErro = "FGE0094 - Seguradora inv�lida (n�o selecionada)."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
        GoTo SAIDA
    End If

    lCodSeguradora = pCodSeguradora


    'Posiciona indicador de sinistro.
    If pIndSinistro = True Then
        lIndSinistro = 1
    Else
        lIndSinistro = 0
    End If

    If M46V999.gfPreenchido(pNumApolice) Then
        lNumApolice = pNumApolice
    Else
        lNumApolice = " "
    End If
    If M46V999.gfPreenchido(pNumItem) Then
        lNumItem = pNumItem
    Else
        lNumItem = " "
    End If
    '1000347 - Inclus�o de informa��es da cong�nere em renova��es --Fim

    '1. Calcula nosso n�mero.

    If M46V999.gfCalNosNumero(lNosNumero, UltimoErro) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
        GoTo SAIDA
    End If

    '2. Abre base de dados.
    If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, bdP00MPED, UltimoErro) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '5. Calcula n�mero do pedido.
    If M46V999.gfCalNumPedido(lNumPedido, UltimoErro, bdP00MPED) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
        GoTo SAIDA
    End If
    M46V999.gPrefixo = "Cota��o " & lNumPedido & ". "

    '6. Posiciona per�odo de vig�ncia.
    lVigDatIni = CDbl(Format$(Now, "yyyymmdd"))
    lVigDatTer = lVigDatIni + 10000

    '7. Busca produto que atenda � vig�ncia.
    If M46V999.gfBusProduto(lVigDatIni, lCodProduto, lCodProtocolo, lCodSequencia, lDescTarifa, UltimoErro, 0) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
        GoTo SAIDA
    End If

    '8. Grava registro de dados comuns.
    If M46V001.gfGraNovPedido(bdP00MPED, lNosNumero, lNumPedido, lVigDatIni, lVigDatTer, lCodProduto, lCodProtocolo, _
                              lCodSequencia, UltimoErro, CInt(pCodRamo), pCodPlano, pCodSegmentacao, lStcPed, tagDynaTrace) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '1000347 - Inclus�o de informa��es da cong�nere em renova��es
    '09. Grava o item
    If M46V999.gfGraItem(bdP00MPED, lNosNumero, lNumPedido, "0001", UltimoErro, _
                         e_TipEmissao_RenCongenere, , lCodSeguradora, lNumApolice, lNumItem, lIndSinistro, _
                         CInt(pCodRamo), lCodProduto, pCodPlano, lStcPed) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '10. Cria arquivo de comunica��o.
    'If M46V111.gfGraRegComunicacao(lNosNumero, UltimoErro) = False Then
    '    Call M46V999.gpGraLog(0,, UltimoErro)
    '    GoTo SAIDA
    'End If

    Set lStcPed = gfProcuraNossoNumero(lNosNumero, False, "", tagDynaTrace)
    If mfPosUniProProdutor(lStcPed, UltimoErro) = False Then
        GoTo ControleErro
    End If

Retorno:
    Set gfProRenCongenere = lStcPed

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    UltimoErro = "[ERR] - gfProRenCongenere ERRO - " & Err & " " & Err.Description & UltimoErro
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA
FechaObj:
    Set bdP00MPED = Nothing
    Set lStcPed = Nothing
    Return

End Function

Private Function mfCarregaUsuarioPortal(pMensagem As String, tagDynaTrace As String) As Boolean
    Dim lSql As String
    Dim lrst As clsYasRecordSet

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfCarregaUsuarioPortal"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    On Error GoTo ControleErro

    If Not M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, M46V999.mConexaoSQL, pMensagem) Then
        GoTo SAIDA
    End If

    lSql = " select cod_corretor, c.cod_corr From yasudanet.dbo.Tab_Grupo_Corr as c WITH (NOLOCK) inner join yasudanet.dbo.Tab_Portal_Acesso as a WITH (NOLOCK) on c.cod_user = a.cod_user  "
    lSql = lSql & " inner join ged.dbo.tbs_corretor as co WITH (NOLOCK) on co.num_susep = c.num_susep "
    lSql = lSql & " where c.Cod_User = '" & M46V999.gNomUsuario & "' And cod_portal = 1"
    lSql = lSql & " and cod_ativo = 1"

    If Not M46V999.gfObtRegistro(M46V999.mConexaoSQL, lSql, lrst, pMensagem) Then
        Exit Function
    End If
    If Not lrst.EOF Then
        If IsNull(lrst("cod_corr")) Then
            M46V999.gCod_Corretor = lrst("cod_corretor")
        Else
            If lrst("cod_corr") = 0 Then
                M46V999.gCod_Corretor = lrst("cod_corretor")
            Else
                M46V999.gCod_Corretor = lrst("cod_corr")
            End If
        End If
    Else
        pMensagem = "AGE0178 - Usu�rio n�o cadastrado. Acesso negado."
        GoTo ControleErro
    End If


    mfCarregaUsuarioPortal = True

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    UltimoErro = "[AUT] - mfCarregaUsuarioPortal " & Err.Number & " " & Err.Description & pMensagem
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    pMensagem = UltimoErro
    GoTo SAIDA
FechaObj:
    Set lrst = Nothing
    Return
End Function

Private Function mfConsPessoa(ByRef pStcPed As stcA46V702B, ByRef pMsgErr As String)
'Verificar se � PJ quando ramo Condom�nio

    If pStcPed.Cod_Produto >= 946 Then
        If pStcPed.Cod_Ramo = e_CodRamo_Condominio And pStcPed.Tip_Pessoa = e_TipPessoa_Fisica Then
            pMsgErr = "FNP0074 - " & M46V999.gPrefixo & "Tipo de pessoa inv�lido para Condom�nio."
            Exit Function
        End If
    End If

    mfConsPessoa = True
    Exit Function

ControleErro:
    mfConsPessoa = False

End Function

Private Sub mpVerComum(ByRef pStcPed As stcA46V702B, ByRef pColAlertas As Collection, ByRef pColErros As Collection)
'Procedure: consiste dados comuns

    Dim lTipPessoa As Byte    'Tipo da pessoa: 1 = pessoa jur�dica e 2 = pessoa f�sica.
    Dim lQutdeParcelas As Integer     'Quantidade de parcelas
    Dim lI As Integer    'Utilizado no for...next
    Dim lMaxParcela As Integer  'N�mero de parcelas m�ximo permitido.
    Dim clsA46V102 As clsA46V102A    'DLL : consiste dados comuns (n�vel 2).
    Dim lObjstcVenc As stcA46V713B    'Objeto da classe de estrutura stcA46V713B (TAB_PED_Venc)
    Dim lobjstcCobert As stcA46V704B
    Dim lObjstcCorr As stcA46V705B    'Objeto da classe de estrutura stcA46V705B (TAB_PED_Corr)
    Dim lobjstcCoss As stcA46V706B    'Objeto da classe de estrutura stcA46V706B (TAB_PED_Coss)
    Dim lobjstcItem As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcEspec As stcA46V726B
    Dim lCosseguros(5, 2) As String         'Cosseguros.
    Dim lCod_Retorno As Integer
    Dim clsA99V001 As clsA99V001
    Dim lMensagem As String
    Dim lDigito As String
    Dim lAuxiliar As String
    Dim lPercAux As Single
    Dim lComissao As Single
    Dim lValLMG As Long
    Dim lCodCorr As Long
    Dim lCalcDat1 As Date
    Dim lCalcDat2 As Long
    Dim lBlnSiscad As Boolean    'Indicador se existe dados no Siscad
    Dim lValidade As Long
    Dim txtobs As String
    Dim lStatusSiscota As Integer
    Dim lRamoSiscota As Long
    Dim lClassificacao As String
    Dim lClassif As String
    Dim lMsgErr As String
    Dim lAuxMsg As String
    Dim lobjPedObs As clsA46V714A

    Dim lString1, lString2, lPosicao

    Dim X As Long
    Dim DadosChave
    Dim blnAchouW As Boolean

    Dim lPerguntaQuest As stcConsistPreCalc

    Dim msgBens As String
    Dim strItens As String
    Dim strCob As String

    Dim lLimDias As Integer

    Dim lobjstcJuros As stcA46V711B

    Dim lJuros As Single

    Dim lQuaDias As Long
    Dim lQuaDias_Orig As Long

    Dim lTip As Integer
    Dim Rst_Unid_Ativa As clsYasRecordSet
    Dim lUnidadeAtiva As Boolean
    Dim lArrayCodDoc As Variant
    Dim lValidaDoc As Boolean
    Dim i As Integer
    
    lPercAux = 0
    lCodCorr = 0
    lComissao = 0

    'lClassif = Val(Trim(Left(pStcPed.ColecaoPedLoc(1).Dsc_Clasf, 3)) + Trim(Mid(pStcPed.ColecaoPedLoc(1).Dsc_Clasf, 6, 2)))
    lClassif = pStcPed.ColecaoPedLoc(1).Cod_Clasf & pStcPed.ColecaoPedLoc(1).Cod_Compl_Clasf

    '1100237 - Tarifa 1008 e 1009 Abril 2011
    'lClassificacao = Trim(Left(pStcPed.ColecaoPedLoc(1).Dsc_Clasf, 4))
    lClassificacao = pStcPed.ColecaoPedLoc(1).Cod_Clasf

    If pStcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pStcPed.ColecaoPedCorr.Item(1)
        lPercAux = lObjstcCorr.Per_Corr
        lCodCorr = lObjstcCorr.Cod_Corr
    End If

    lValLMG = pStcPed.gValLMG

    'Ficha 1000308 - Permitir itens com planos diferentes no produto Roubo (c�lculo coletivo)
    'Roubo n�o permite c�lculo de itens (Riscos Comerciais e/ou Industriais) com os demais

    If pStcPed.gTemPlanoComercial = True And pStcPed.gTemPlanoNaoComercial = True Then
        lMsgErr = "FGE0050 - N�o � permitido c�lculo de itens (Riscos Comerciais e/ou Industriais) com os demais."
        Call M46V999.gpMenErro(lMsgErr)
    End If


    'Ficha 900653
    lString1 = pStcPed.Nom_Segurado
    lString2 = "ESPOLIO"

    lPosicao = InStr(lString1, lString2)

    If (objstcUsuario.TipoUsuario = e_Tipo_Usu_Corretor Or objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda) And lPosicao > 0 Then
        lMsgErr = "FGE0050 - Palavra 'ESPOLIO' n�o permitida no nome do Segurado"
        Call M46V999.gpMenErro(lMsgErr)
    End If

    'Ficha 901224 e 901732
    If pStcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        Set lobjstcItem = pStcPed.ColecaoPedLoc(1)
        If UCase(lobjstcItem.objstcPedObs.Observacao) Like "*VENCTO*" Or _
           UCase(lobjstcItem.objstcPedObs.Observacao) Like "*VECTO*" Or _
           UCase(lobjstcItem.objstcPedObs.Observacao) Like "*VENCIMENTO*" Or _
           UCase(lobjstcItem.objstcPedObs.Observacao) Like "*PAGAMENTOS*" Then
            lMsgErr = "FGE0051 - Preencher as informa��es de vencimento no campo correto (Ajustar data das parcelas)"
            Call M46V999.gpMenErro(lMsgErr)
        End If
    End If

    'Ficha 901782
    If pStcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        If UCase(pStcPed.Nom_Segurado) Like "*E/OU*" Then
            lMsgErr = "FGE0052 - Nome inv�lido (n�o utilizar E/OU)."
            Call M46V999.gpMenErro(lMsgErr)
        End If
    End If


    '1001224 - Condi��es Diferenciada - Seguratec
    Set lobjstcItem = pStcPed.ColecaoPedLoc(1)
    If lobjstcItem.ColecaoPedCobert.Count > 0 Then
        Call M46V111.gfBuscaComissaoParm(lCodCorr, pStcPed.Cod_Unid_Prod, lComissao, lValLMG, lClassif, lMensagem, pStcPed, lobjstcItem)
    Else
        Call M46V999.gpMenErro("FNP9999 - Nenhuma cobertura est� selecionada.(Item:001)")
    End If

    If lComissao > 0 And lPercAux > lComissao Then
        lMsgErr = "FGE0050 - Comiss�o Inv�lida para esta rubrica/LMG" & lMensagem
        Call M46V999.gpMenErro(lMsgErr)
    End If

    If objstcUsuario.TipoUsuario = e_Tipo_Usu_Corretor Or objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda Then
        'Ajuste para n�o validar qdo retorno de classe com premio net quando cota��o
        If lPercAux = 0 And pStcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
            lMsgErr = "FNP1010 - Comiss�o deve estar preenchida!"
            Call M46V999.gpMenErro(lMsgErr)
        End If
        '    Else
        '**** Consistencia p/ usuario interno
        '       ' If M46V999.gfVerEtiqueta(lCodCorr,   Mid(frmT46V112A.cboCodReg.Text, 1, 3), lTipPessoa, Mid(frmT46V112A.cboTipEnvio.Text, 1, 1), lMensagem) Then
        '        If M46V999.gfVerEtiqueta(lCodCorr, pStcPed.Cod_Reg, pStcPed.Tip_Pessoa, pStcPed.objstcCtrlEmis.Cod_Envi_Doc, lMensagem) Then
        '            Call M46V999.gpMenErro("FGE0019 - " & lMensagem)
        '        Else
        '            'Ficha:0801295
        '            'Novas op��es "Tipo de Emiss�o=Conforme cadastro":
        '            '"O" -> Conforme Cadastro Corretor e "E" -> Conforme Cadastro Segurado
        '            'a defini��o se "O" ou "E" � feita dentro da M46V999.gfVerEtiqueta
        '            If lMensagem = "O" Then '<- digitacao
        '                pStcPed.objstcCtrlEmis.Cod_Envi_Doc = "O"
        '            ElseIf lMensagem = "E" Then
        '                pStcPed.objstcCtrlEmis.Cod_Envi_Doc = "E"
        '            End If
        '        End If
    End If


    'Ficha:900048 Isen��o de IOF para segurados:
    '-Contendo as palavras-chaves no nome;
    '-Ramo de atividade -> Administra��o publica, defesa e seguridade social;
    '-C�digo de registro 1.
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    '11.52
    'Assunto: Fwd: Regra de isen��o de IOF
    'Luana,
    'Bom Dia!
    '
    'Solicito inibir a regra mencionada de isen��o de IOF mencionada em seu e-mail,  para todas as atividades, ou seja,  atividade 12 - Atividade Imobili�ria ou 15 - Administra��o P�blica.
    'Caber� a �rea de subscri��o ap�s solicita��o de documentos espec�ficos liberar ou n�o o IOF.
    '
    'Atenciosamente,
    '
    'Margarida Maria Pinheiro de Oliveira
    'Solu��es Commercial Lines
'    If pStcPed.Cod_Produto <= 1152 Then
'        blnAchouW = False
'        DadosChave = Split(UCase(Trim("Estado,Distrito,Federal,Estadual,Municipal,Munic�pio,Municipio,Prefeitura,Governo,P�blic,Public,Secretaria,Minist�rio,Ministerio")), ",")
'        For X = 0 To UBound(DadosChave)
'            If InStr(1, UCase(Trim(pStcPed.Nom_Segurado)), DadosChave(X)) > 0 And pStcPed.Cod_Ramo_Ativ = 12 Then
'                blnAchouW = True
'                Exit For
'            End If
'        Next X
'        If blnAchouW = True Then
'            If pStcPed.Cod_Reg = 1 Or pStcPed.Cod_Reg = 4 Or pStcPed.Cod_Reg = 101 Or pStcPed.Cod_Reg = 104 Then
'                Set lPerguntaQuest = PerguntaQuestionario(pColAlertas, E_IDConsist_IsencaoIOF)
'                If lPerguntaQuest Is Nothing Then
'                    Call mpAddAlerta(E_IDConsist_IsencaoIOF, "Emitir com isen��o de IOF?", E_TipConsist_Quest, pColAlertas)
'                Else
'                    Select Case lPerguntaQuest.RespostaQuest
'                    Case E_RespQuest_ND
'                        lPerguntaQuest.RespostaValidada = False
'                    Case E_RespQuest_Nao
'                        lPerguntaQuest.RespostaValidada = True
'                    Case E_RespQuest_Sim
'                        If pStcPed.ColecaoPedLoc(1).Tip_Emissao >= 300 Then
'                            If pStcPed.Dat_Ini_Vig >= 20130101 Then
'                                pStcPed.Cod_Reg = 102
'                            Else
'                                pStcPed.Cod_Reg = 103
'                            End If
'                        Else
'                            If pStcPed.Dat_Ini_Vig >= 20130101 Then
'                                pStcPed.Cod_Reg = 2
'                            Else
'                                pStcPed.Cod_Reg = 3
'                            End If
'                        End If
'                        lPerguntaQuest.RespostaValidada = True
'                    End Select
'                End If
'            End If
'        Else
'            Call mfRemoveQuestionario(E_IDConsist_IsencaoIOF, pColAlertas)
'            If pStcPed.Cod_Reg = 3 Or pStcPed.Cod_Reg = 103 Then
'                If pStcPed.ColecaoPedLoc(1).Tip_Emissao >= 300 Then
'                    If pStcPed.Dat_Ini_Vig >= 20130101 Then
'                        pStcPed.Cod_Reg = 104
'                    Else
'                        pStcPed.Cod_Reg = 101
'                    End If
'                Else
'                    If pStcPed.Dat_Ini_Vig >= 20130101 Then
'                        pStcPed.Cod_Reg = 4
'                    Else
'                        pStcPed.Cod_Reg = 1
'                    End If
'                End If
'            End If
'        End If
'    End If

    '--------------------------------------------------------------------------------------------------------------------------------------------



    '1001146 - Tarifa 1000 e 1001 Dezembro 2010
    '---
    strItens = ""
    If pStcPed.Cod_Produto > 999 Then
        For Each lobjstcItem In pStcPed.ColecaoPedLoc
            '            If lobjstcItem.Cob_Bens = True Then
            For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
                If (pStcPed.Cod_Ramo = e_CodRamo_Empresarial And lobjstcCobert.Cod_Cobert = 94) Or _
                   (pStcPed.Cod_Ramo = e_CodRamo_Residencial And (lobjstcCobert.Cod_Cobert = 27 Or _
                                                                  lobjstcCobert.Cod_Cobert = 29 Or _
                                                                  lobjstcCobert.Cod_Cobert = 30 Or _
                                                                  lobjstcCobert.Cod_Cobert = 31)) Then
                    'se nao existir rela��o de bens
                    'If lobjstcItem.Dig_Bens = False Then
                    If lobjstcItem.ColecaoPedLocBem.Count = 0 Then
                        'guarda o item somente 1 vez na vari�vel
                        If InStr(1, strItens, lobjstcItem.Num_Item, vbTextCompare) = 0 Then
                            strItens = strItens & lobjstcItem.Num_Item & ","
                        End If
                        'guarda a cobertura somente 1 vez na vari�vel
                        If InStr(1, strCob, lobjstcCobert.Cod_Cobert, vbTextCompare) = 0 Then
                            strCob = strCob & lobjstcCobert.Cod_Cobert & ","
                        End If
                    End If
                End If
            Next lobjstcCobert
            '            End If
        Next lobjstcItem
        '1001146 - Tarifa 1000 e 1001 Dezembro 2010
        If Len(strItens) > 0 Then strItens = IIf(Len(strItens) > 2, "Itens(", "Item(") & Left(strItens, Len(strItens) - 1) & ")"
        If Len(strCob) > 0 Then strCob = Left(strCob, Len(strCob) - 1)
    End If

    If strItens <> "" Then
        'se existir pelo menos um item com cobertura de bens sem digita��o de rela��o de bens
        'CONSISTENCIA RETIRADA NA NOVA TARIFA 1080 / 1081
'        strItens = strItens & vbCrLf & vbCrLf
'        msgBens = strItens & "Para a cobertura " & strCob & " � obrigat�rio relacionar os bens e respectivos valores." & vbLf & vbLf & _
'                  "A rela��o poder� ser preenchida diretamente no SYAS ou ser entregue juntamente com a proposta de seguro." & vbLf & vbLf & _
'                  "Deseja preencher a rela��o de bens agora?"
'
'        Set lPerguntaQuest = PerguntaQuestionario(pColAlertas, E_IDConsist_RelacaoBens)
'        If lPerguntaQuest Is Nothing Then
'            Call mpAddAlerta(E_IDConsist_RelacaoBens, msgBens, E_TipConsist_Quest, pColAlertas)
'        Else
'            Select Case lPerguntaQuest.RespostaQuest
'            Case E_RespQuest_ND
'                lPerguntaQuest.RespostaValidada = False
'            Case E_RespQuest_Nao
'                lPerguntaQuest.RespostaValidada = True
'            Case E_RespQuest_Sim
'                lPerguntaQuest.RespostaValidada = False
'            End Select
'        End If
    Else
        Call mfRemoveQuestionario(E_IDConsist_RelacaoBens, pColAlertas)
    End If


    '1200004 - Tarifa 1028 e 1029 Janeiro 2012 - Item 2
    strItens = ""
    If pStcPed.Cod_Produto >= 1028 Then
        For Each lobjstcItem In pStcPed.ColecaoPedLoc
            If lobjstcItem.Cod_DizerItem <> 4 And lobjstcItem.Cod_DizerItem <> 10 Then
                For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
                    If (pStcPed.Cod_Ramo = e_CodRamo_Empresarial And lobjstcCobert.Cod_Cobert = 28) Then
                        'guarda o item somente 1 vez na vari�vel
                        If InStr(1, strItens, lobjstcItem.Num_Item, vbTextCompare) = 0 Then
                            strItens = strItens & lobjstcItem.Num_Item & ","
                        End If
                    End If
                Next lobjstcCobert
            End If
        Next lobjstcItem
        If Len(strItens) > 0 Then strItens = IIf(Len(strItens) > 2, "Itens(", "Item(") & Left(strItens, Len(strItens) - 1) & ")"
    End If

    If strItens <> "" Then
        Set lPerguntaQuest = PerguntaQuestionario(pColAlertas, E_IDConsist_Cob28_Empresarial)
        If lPerguntaQuest Is Nothing Then
            msgBens = strItens & "A cobertura 28 - Vendaval (Exceto Bens ao Ar Livre) n�o garante bens que estiverem ao ar livre, " & vbCrLf & _
                      "em terra�o, edifica��es abertas ou semiabertas, galp�es, alpendres ou semelhantes. " & vbCrLf & _
                      "Caso deseje cobertura para estes bens, " & _
                      "contrate a cobertura 35 - Vendaval (Inclusive Bens ao Ar Livre)." & vbCrLf & vbCrLf & _
                      "Deseja Prosseguir?"
            Call mpAddAlerta(E_IDConsist_Cob28_Empresarial, msgBens, E_TipConsist_Quest, pColAlertas)
        Else
            Select Case lPerguntaQuest.RespostaQuest
            Case E_RespQuest_ND
                lPerguntaQuest.RespostaValidada = False
            Case E_RespQuest_Nao
                lPerguntaQuest.RespostaValidada = False
            Case E_RespQuest_Sim
                lPerguntaQuest.RespostaValidada = True
                msgBens = "Voc� optou por contratar a cobertura 28 - Vendaval (Exceto Bens ao Ar Livre)"
                Call mpAddAlerta(E_IDConsist_Cob28_Aviso_Empresarial, msgBens, E_TipConsist_Aviso, pColAlertas)
            End Select
        End If
    Else
        Call mfRemoveQuestionario(E_IDConsist_Cob28_Empresarial, pColAlertas)
        Call mfRemoveQuestionario(E_IDConsist_Cob28_Aviso_Empresarial, pColAlertas)
    End If


    '**************Ser� realizada pelo COBOL.Net
    '1. Consiste dados comuns - n�vel 1 (dados para c�lculo).
    '    Call mpConsisteComum(lCod_Retorno)         'DLL
    '**************

    If Right(pStcPed.Dat_Ini_Vig, 4) = "0229" Then
        lMsgErr = "FGE0046 - Vig�ncia inv�lida (data de in�cio com ano " & _
                  "bissexto - utilizar 28/02 ou 01/03)."
        Call M46V999.gpMenErro(lMsgErr)
    End If

    ' Consist�ncia ser� retirada do Syas para incluir no Batch
    If M46V999.gCodProduto >= 1082 Then
        If pStcPed.ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_RenYasuda Then
            lCalcDat1 = CDate(M46V999.gfForData(2, pStcPed.Dat_Ini_Vig)) + 3
        Else
            lCalcDat1 = CDate(M46V999.gfForData(2, pStcPed.Dat_Ini_Vig))
        End If
    Else
    lCalcDat1 = CDate(M46V999.gfForData(2, pStcPed.Dat_Ini_Vig)) + 7
    End If
    
    lCalcDat2 = M46V999.gfForData(1, Format(lCalcDat1, "dd/mm/yyyy"))

    If Format(Date, "yyyymmdd") > lCalcDat2 Then
        If lCalcDat2 < Format(Date, "yyyymmdd") And pStcPed.Ind_Reentrada <> 1 And objstcUsuario.TipoUsuario = e_Tipo_Usu_Corretor Then
            If M46V999.gCodProduto >= 1082 Then
                If pStcPed.ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_RenYasuda Then
                    lMsgErr = "FGE0150 - In�cio de vig�ncia inv�lido (superior a 3 dias)"
                Else
                    'Tarifa 11.32 Residencial corretor ESTAR.
                    'Aumentar inicio de vig�ncia para at� 5 dias anteriores a data atual
                    If pStcPed.Cod_Produto >= 1132 And pStcPed.Cod_Ramo = e_CodRamo_Residencial And pStcPed.objstcCtrlEmis.Cod_Corr = 908624 Then
                        If CDate(M46V999.gfForData(2, pStcPed.Dat_Ini_Vig)) + 5 < Date Then
                            lMsgErr = "FGE0150 - In�cio de vig�ncia inv�lido para este corretor. Data in�cio de Vig�ncia pode ser at� 5 dias retroativa � data atual."
                        End If
                    Else
                                                'CR93893 Ajuste para n�o gerar erro quando for retorno de classe 4 (rsilva)
                            If Not pStcPed.Ind_CotacaoSolic Then
                                lMsgErr = "FGE0150 - In�cio de vig�ncia inv�lido (inferior a data atual)"
                                                End If
                        End If
                End If
            Else
                lMsgErr = "FGE0150 - In�cio de vig�ncia inv�lido (superior a 7 dias)"
            End If
            'CR93893 Ajuste para n�o gerar erro quando for retorno de classe 4 (rsilva)
            If Len(lMsgErr) > 0 Then
                Call M46V999.gpMenErro(lMsgErr)
                End If
        End If
    End If
    
    '8000048167, Flexibiliza��o de Renova��es
    '1200708 - Limite Inicio de Vigencia
    If Not (pStcPed.StatusCotSiscota = enum_StatusTransm_Aprovada Or pStcPed.StatusCotSiscota = enum_StatusTransm_EnvioKit) Then
    
        If pStcPed.Dat_Ini_Vig > 0 Then
            lCalcDat1 = DateSerial(Mid(pStcPed.Dat_Ini_Vig, 1, 4), Mid(pStcPed.Dat_Ini_Vig, 5, 2), Mid(pStcPed.Dat_Ini_Vig, 7, 2))
            If lCalcDat1 > DateAdd("m", 2, Date) Then
                Call M46V999.gpMenErro("FGE0150 - Inicio de Vig�ncia inv�lida (superior a 2 meses)")
            End If
        End If
    
    End If

    If pStcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        If (pStcPed.Cod_Estip >= 2001 And pStcPed.Cod_Estip <= 2010) Or (pStcPed.Cod_Estip >= 2032 And pStcPed.Cod_Estip <= 2041) Then
            If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Nao Then
                lMsgErr = "AGE0111 - Para este plano, o pagamento s� pode ser antecipado."
                Call M46V999.gpMenErro(lMsgErr)
            End If
        End If
    End If

    If pStcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        '2. Consiste dados comuns - n�vel 2.
        Set clsA46V102 = New clsA46V102A    'DLL : consiste dados comuns (n�vel 2).
        Set clsA46V102.M46V999 = M46V999

        With clsA46V102
            'a) Prepara dados para consist�ncia de comum n�vel 2.
            .Pro_CarBanco = pStcPed.Cod_Bco_Cobr
            'Obs.: O campo Ven_Parcela1_VM cont�m a data de vencimento da primeira parcela do carn�.
            If pStcPed.ColecaoPedVenc.Count > 0 Then
                Set lObjstcVenc = pStcPed.ColecaoPedVenc.Item(1)
                .Pro_CarDatPriParcela = lObjstcVenc.DAT_VENC
            End If
            .Pro_Cob_Bairro = pStcPed.Nom_Bairro
            .Pro_Cob_CEP = Format(pStcPed.Num_CEP, "00000000")
            .Pro_Cob_Cidade = pStcPed.Nom_Cid
            If M46V999.gfPreenchido(pStcPed.Dsc_Compl_Logr) Then
                .Pro_Cob_Complemento = pStcPed.Dsc_Compl_Logr
            Else
                .Pro_Cob_Complemento = ""
            End If
            .Pro_Cob_NomLogradouro = pStcPed.Nom_Logr
            .Pro_Cob_TipLogradouro = pStcPed.Cod_Tip_Logr
            .Pro_Cob_UF = pStcPed.Sig_UF
            If pStcPed.ColecaoPedCorr.Count > 0 Then
                Set lObjstcCorr = pStcPed.ColecaoPedCorr.Item(1)
                .Pro_CodCorretor = lObjstcCorr.Cod_Corr
            Else
                .Pro_CodCorretor = 0
            End If
            lCosseguros(1, 1) = 0
            lCosseguros(1, 2) = 0
            lCosseguros(2, 1) = 0
            lCosseguros(2, 2) = 0
            lCosseguros(3, 1) = 0
            lCosseguros(3, 2) = 0
            lCosseguros(4, 1) = 0
            lCosseguros(4, 2) = 0
            lCosseguros(5, 1) = 0
            lCosseguros(5, 2) = 0
            lI = 0
            For Each lobjstcCoss In pStcPed.ColecaoPedCoss
                lI = lI + 1
                lCosseguros(lI, 1) = lobjstcCoss.Cod_Coss
                lCosseguros(lI, 2) = lobjstcCoss.Per_Coss
            Next
            .Pro_Cosseguros = lCosseguros
            .Pro_DebAgencia = pStcPed.gAgenciaDig
            .Pro_DebBanco = pStcPed.Cod_Bco_Cobr
            .Pro_DebConta = pStcPed.Num_Cta_Bas_Cor
            'Verificar se consegue fazer a consist�ncia,
            'pois para Yasuda foi gravado sem digito
            If objstcUsuario.TipoUsuario = e_Tipo_Usu_Yasuda Then
                .Pro_DebConta = .Pro_DebConta & Right(pStcPed.Num_Cta_Bas_Cor, 1)
            End If
            .Pro_End_Bairro = pStcPed.objstcPedEndereco.Nom_Bairro
            .Pro_End_CEP = Format(pStcPed.objstcPedEndereco.Num_CEP, "00000000")
            .Pro_End_Cidade = pStcPed.objstcPedEndereco.Nom_Cid
            If M46V999.gfPreenchido(pStcPed.objstcPedEndereco.Dsc_Compl_Logr) Then
                .Pro_End_Complemento = pStcPed.objstcPedEndereco.Dsc_Compl_Logr
            Else
                .Pro_End_Complemento = ""
            End If
            .Pro_End_NomLogradouro = pStcPed.objstcPedEndereco.Nom_Logr
            .Pro_End_TipLogradouro = pStcPed.objstcPedEndereco.Cod_Tip_Logr
            .Pro_End_UF = pStcPed.objstcPedEndereco.Sig_UF
            'Proposta: forma de pagamento.
            '1 = carn� e 2 = d�bito em conta.
            .Pro_ForPagamento = pStcPed.gTipPagamento
            .Pro_IndAgencia = pStcPed.gAgenciaDig
            .Pro_IndBanco = pStcPed.Cod_Bco_Cobr
            .Pro_IndConta = pStcPed.Num_Cta_Bas_Cor
            .Pro_IndCobranca = IIf(pStcPed.gEnd_Cob_Diferente = True, 1, 2)
            .Pro_IndCosseguros = IIf(pStcPed.ColecaoPedCoss.Count > 0, 1, 2)
            .Pro_NomSegurado = pStcPed.Nom_Segurado
            .Pro_PagAntBloqueto = pStcPed.objstcCtrlEmis.Num_Bloq
            .Pro_PagAntecipado = pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec
            .Pro_QtdParcelas = pStcPed.Qtd_Parc
            .Pro_UniProdutiva = pStcPed.Cod_Unid_Prod
            .Pro_VigInicio = pStcPed.Dat_Ini_Vig
            .Pro_TipPessoa = pStcPed.Tip_Pessoa
            .Pro_CNPJ_CPF = pStcPed.Num_Cgc_Cpf
            .Pro_CodSegurado = pStcPed.Cod_Segurado
            .Pro_Cod_Reg = pStcPed.Cod_Reg
            .Pro_CreBanco = pStcPed.objstcCtrlEmis.Cod_Banco_Cre
            .Pro_CreAgencia = pStcPed.objstcCtrlEmis.Cod_Agenc_Cre
            .Pro_CreConta = pStcPed.objstcCtrlEmis.Cod_Conta_Cre & IIf(UCase(pStcPed.objstcCtrlEmis.Cod_DIG_Conta) = "X", "0", pStcPed.objstcCtrlEmis.Cod_DIG_Conta)
            'b) Executa consist�ncia.
            On Error Resume Next
            Call .VER_COMUM_NIVEL_2
            If Err <> 0 Then
                Call M46V999.gpMenErro("FTE0009 - " & M46V999.gPrefixo & "Erro ao consistir dados comuns n�vel 2 (" & _
                                       Err & " - " & ERROR & ").")
                On Error GoTo 0
                GoTo SAIDA
            End If
            On Error GoTo 0
        End With
        Set clsA46V102 = Nothing

        '3. Consiste dados de identifica��o do proponente.
        If pStcPed.Tip_Pessoa = e_TipPessoa_Fisica Then
            lTipPessoa = 2  'Pessoa f�sica.
        Else
            lTipPessoa = 1  'Pessoa jur�dica.
        End If
        'PE 0401710 - Mensagem de erro quando n�o informado dados SUSEP.
        Select Case lTipPessoa
        Case 1  'Pessoa jur�dica.
            '1100705 - Tarifa 1016 e 1017 Agosto 2011 - Item 1 n�o obrigar informar Ramo de Atividade
            If pStcPed.Cod_Produto >= 1016 Then
                If (pStcPed.Cod_Ramo_Ativ = 0) And (pStcPed.Cod_Ramo <> e_CodRamo_Condominio) Then
                    '                       CR31199: Syas Web Multi - Ajuste da Descri��o de Texto
                    'Call M46V999.gpMenErro("FNP0001 - " & M46V999.gPrefixo & "Ramo de atividade do proponente n�o selecionado (F5 - dados proposta).")
                    Call M46V999.gpMenErro("FNP0001 - " & M46V999.gPrefixo & "Ramo de atividade do proponente n�o selecionado.")
                End If
            Else
                If (pStcPed.Cod_Ramo_Ativ = 0) Then
                    '                       CR31199: Syas Web Multi - Ajuste da Descri��o de Texto
                    'Call M46V999.gpMenErro("FNP0001 - " & M46V999.gPrefixo & "Ramo de atividade do proponente n�o selecionado (F5 - dados proposta).")
                    Call M46V999.gpMenErro("FNP0001 - " & M46V999.gPrefixo & "Ramo de atividade do proponente n�o selecionado.")
                End If
            End If
            If Not M46V999.gfPreenchido(pStcPed.gCNPJ_CPF_Formatado) Then
                '                       CR31199: Syas Web Multi - Ajuste da Descri��o de Texto
                '                        Call M46V999.gpMenErro("FNP0002 - CNPJ do segurado n�o preenchido." & _
                                         '                                       "(F5 - dados proposta).")
                Call M46V999.gpMenErro("FNP0002 - CNPJ do segurado n�o preenchido.")
            End If
            'Fichas 1200168 e 1200169 Inclus�o dos campos de listagem Renda e Faixa de Faturamento Pessoa f�sica e jur�dica
            If (pStcPed.Val_RendaFatu = -1) Then
                '                   CR31199: Syas Web Multi - Ajuste da Descri��o de Texto
                '                   Call M46V999.gpMenErro("FNP0001 - " & M46V999.gPrefixo & "Faturamento do proponente n�o selecionado (F5 - dados proposta).")
                Call M46V999.gpMenErro("FNP0001 - " & M46V999.gPrefixo & "Faturamento do proponente n�o selecionado.")
            End If
        Case 2  'Pessoa f�sica.
            'a) Tipo de documento de identifica��o.
            
            If pStcPed.Cod_Doc <> 0 And pStcPed.ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_Endosso Then
                lArrayCodDoc = Split("1;2;3;7;8;9;10", ";")
                
                lValidaDoc = True
                
                For i = 0 To UBound(lArrayCodDoc)
                    If pStcPed.Cod_Doc = lArrayCodDoc(i) Then
                        lValidaDoc = False
                        Exit For
                    End If
                Next
            End If
            
            If pStcPed.Cod_Doc = 0 Or lValidaDoc Then
                '                    Call M46V999.gpMenErro("FNP0003 - Tipo do documento do proponente n�o selecionado " & _
                                     '                                   "(F5 - dados proposta).")
                Call M46V999.gpMenErro("FNP0003 - Tipo do documento do proponente n�o selecionado.")
            End If
            'b) N�mero do documento.
            If Not M46V999.gfPreenchido(pStcPed.Num_Doc) Then
                '                    Call M46V999.gpMenErro("FNP0004 - N�mero do documento do proponente n�o preenchido " & _
                                     '                                   "(F5 - dados proposta).")
                Call M46V999.gpMenErro("FNP0004 - N�mero do documento do proponente n�o preenchido.")
            End If

            If Not M46V999.gfPreenchido(pStcPed.gCNPJ_CPF_Formatado) Then
                '                    Call M46V999.gpMenErro("FNP0005 - CPF do segurado n�o preenchido." & _
                                     '                                   "(F5 - dados proposta).")
                Call M46V999.gpMenErro("FNP0005 - CPF do segurado n�o preenchido.")
            End If

            If Not M46V999.gfPreenchido(pStcPed.Sexo) Then
                '                    Call M46V999.gpMenErro("FNP0006 - Indicador de sexo do segurado n�o selecionado  (F5 - dados proposta).")
                Call M46V999.gpMenErro("FNP0006 - Indicador de sexo do segurado n�o selecionado.")

            End If
            'Fichas 1200168 e 1200169 Inclus�o dos campos de listagem Renda e Faixa de Faturamento Pessoa f�sica e jur�dica
            If (pStcPed.Val_RendaFatu = -1) Then
                '                    Call M46V999.gpMenErro("FNP0001 - Renda do proponente n�o selecionada (F5 - dados proposta).")
                Call M46V999.gpMenErro("FNP0001 - Renda do proponente n�o selecionada.")
            End If
        End Select
        If pStcPed.Qtd_Parc = 0 And pStcPed.Cod_Reg < 100 Then
            Call M46V999.gpMenErro("FGE0111 - Quantidade de parcelas inv�lida (n�o informada)")
        End If


        If pStcPed.objstcCtrlEmis.Num_Bloq <> 0 Then
            lAuxiliar = pStcPed.objstcCtrlEmis.Num_Bloq & pStcPed.objstcCtrlEmis.DIG_BLOQUETO
            Set clsA99V001 = New clsA99V001
            Set clsA99V001.M46V999 = M46V999

            Call clsA99V001.DIG_BLOQUETO(Format(lAuxiliar, "000000000"), lDigito, lMensagem)
' Comentado em 10/2/2017 por Daniel Ribeiro devido ao projeto boleto com registro
'            If Right(lAuxiliar, 1) <> lDigito Then
'                Call M46V999.gpMenErro("FGE0059 - Pagamento antecipado com bloqueto - n�mero do bloqueto com d�gito inv�lido.")
'            End If
        End If

        lLimDias = 31

        '1200873 - Agendamento da primeira parcela para at� 60 dias
        Select Case pStcPed.Cod_Estip
        Case 2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030, 2031
            lLimDias = 60
        End Select

        Select Case Left(CStr(pStcPed.Cod_Segurado), 6)
        Case 777701, 777702, 777704, 777705, 777706, 777707, 777709, 777710, 777711, 777712, 777713, 777714, 777716, 777733, 777787, 777788
            lLimDias = 60
        End Select

        'Ficha 901170
        Select Case pStcPed.Cod_Segurado
        Case 126, 503863
            '               lLimDias = 60
            'Ficha 901651
            lLimDias = 90
        End Select

        If objstcUsuario.TipoUsuario = e_Tipo_Usu_Yasuda Then
            'Ficha 900745
            lLimDias = 99
        End If

        If pStcPed.gTemAjusVenc = True Then
            If pStcPed.ColecaoPedVenc.Count < pStcPed.Qtd_Parc Then
                Call M46V999.gpMenErro("FGE0111 - Favor informar o vencimento da(s) " & pStcPed.Qtd_Parc & " parcela(s).")
                GoTo SAIDA
            End If
            If pStcPed.ColecaoPedVenc(1).DAT_VENC <> 0 Then
                For lI = 1 To val(pStcPed.Qtd_Parc)
                    If lI = 1 Then
                        If CDate(pStcPed.ColecaoPedVenc(lI).gDataVenc_Formatada) > CDate(Date + lLimDias) Then
                            Call M46V999.gpMenErro("FGE0111 - Data de vencimento da parcela " & lI & " n�o pode ser superior � " & lLimDias & " dias (" & pStcPed.ColecaoPedVenc(lI).gDataVenc_Formatada & ").")
                        End If
                    End If
                    If Not M46V999.gfPreenchido(pStcPed.ColecaoPedVenc(lI).gDataVenc_Formatada) Then
                        Call M46V999.gpMenErro("FGE0112 - Data de vencimento da parcela " & lI & " n�o preenchida.")
                    Else
                        If Not IsDate(pStcPed.ColecaoPedVenc(lI).gDataVenc_Formatada) Then
                            Call M46V999.gpMenErro("FGE0006 - Data de vencimento da parcela " & lI & " inv�lida (" & pStcPed.ColecaoPedVenc(lI).gDataVenc_Formatada & ").")
                        End If
                        If CDate(pStcPed.ColecaoPedVenc(lI).gDataVenc_Formatada) < CDate(Date + 8) And pStcPed.Cod_Reg <> 106 And pStcPed.Cod_Reg <> 107 Then
                            Call M46V999.gpMenErro("FGE0007 - Data de vencimento da parcela " & lI & " inv�lida (inferior a " & CDate(Date + 8) & ").")
                        End If
                        If CDate(pStcPed.ColecaoPedVenc(lI).gDataVenc_Formatada) > CDate(M46V999.gfForData(2, pStcPed.Dat_Ter_Vig)) - 10 Then
                            Call M46V999.gpMenErro("FGE0110 - Data de vencimento da parcela " & lI & " inv�lida (Superior ao t�rmino da vig�ncia).")
                        End If
                        'Ficha 900046
                        If lI = 1 And pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec <> e_SimNao_Sim And CDate(pStcPed.ColecaoPedVenc(lI).gDataVenc_Formatada) < CDate(Date + 20) Then
                            Call M46V999.gpMenErro("FGE0009 - Data de vencimento da parcela " & lI & " inv�lida (inferior a " & CDate(Date + 20) & ").")
                        End If
                        If lI = 2 And pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim And CDate(pStcPed.ColecaoPedVenc(lI).gDataVenc_Formatada) < CDate(Date + 20) Then
                            Call M46V999.gpMenErro("FGE0009 - Data de vencimento da parcela " & lI & " inv�lida (inferior a " & CDate(Date + 20) & ").")
                        End If
                        If lI > 1 Then
                            If CDate(pStcPed.ColecaoPedVenc(lI).gDataVenc_Formatada) > CDate(pStcPed.ColecaoPedVenc(lI - 1).gDataVenc_Formatada) + lLimDias Then
                                Call M46V999.gpMenErro("FGE0109 - Data de vencimento da parcela " & lI & " inv�lida (Superior a " & lLimDias & " da parcela anterior).")
                            End If
                            'Ficha 901136
                            If DateDiff("d", CDate(pStcPed.ColecaoPedVenc(lI - 1).gDataVenc_Formatada), CDate(pStcPed.ColecaoPedVenc(lI).gDataVenc_Formatada)) < 9 Then
                                Call M46V999.gpMenErro("FGE0109 - Data de vencimento da parcela " & lI & " inv�lida (um vencimento deve ter no m�nimo 9 dias de diferen�a do outro vencimento).")
                            End If
                        End If
                        If lI < pStcPed.Qtd_Parc Then
                            If M46V999.gfPreenchido(pStcPed.ColecaoPedVenc(lI + 1).gDataVenc_Formatada) Then
                                If CDate(pStcPed.ColecaoPedVenc(lI).gDataVenc_Formatada) > CDate(pStcPed.ColecaoPedVenc(lI + 1).gDataVenc_Formatada) Then
                                    Call M46V999.gpMenErro("FGE0009 - Data de vencimento da parcela " & lI & " maior que a data de vencimento da parcela " & lI + 1)
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
        If objstcUsuario.TipoUsuario = e_Tipo_Usu_Corretor Or objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda Then
            If pStcPed.Cod_Prod = 0 Then
                Call M46V999.gpMenErro("FGE0015 - Produtor inv�lido (N�o informado).")
            End If
            If pStcPed.Cod_Unid_Prod = 0 Then
                Call M46V999.gpMenErro("FGE0016 - Unidade inv�lida (N�o informada).")
            End If
            'Ficha 1200066 - Somente permitir unidade ativa
            Set Rst_Unid_Ativa = New clsYasRecordSet
            Call M46V999.mfCarregarUnidAtivas(Rst_Unid_Ativa)
            
            lUnidadeAtiva = False
            If Rst_Unid_Ativa.RecordCount > 0 Then
                Do Until Rst_Unid_Ativa.EOF()
                    If pStcPed.Cod_Unid_Prod = Rst_Unid_Ativa!cod_Depto Then
                        lUnidadeAtiva = True
                        Exit Do
                    End If
                    
                    Rst_Unid_Ativa.MoveNext
                Loop
            End If
            If lUnidadeAtiva = False Then
                Call M46V999.gpMenErro("FGE0016 - Unidade inativa.")
            End If
        End If
    End If


    '1001266 - Validar CPF e CNPJ na Cota��o e Proposta
    Set clsA99V001 = New clsA99V001     'Calcula d�gitos.
    Set clsA99V001.M46V999 = M46V999

    If pStcPed.Tip_Pessoa = e_TipPessoa_Fisica Then
        'Pessoa f�sica.
        lAuxiliar = pStcPed.gCNPJ_CPF_Formatado
        If M46V999.gfPreenchido(lAuxiliar) Then
            If clsA99V001.DIG_CPF(lAuxiliar, lDigito, lMensagem) = False Then
                Call M46V999.gpMenErro(lMensagem)
            End If
        End If
    ElseIf pStcPed.Tip_Pessoa = e_TipPessoa_Juridica Then
        'Pessoa jur�dica.
        lAuxiliar = pStcPed.gCNPJ_CPF_Formatado
        If M46V999.gfPreenchido(lAuxiliar) Then
            If clsA99V001.DIG_CNPJ(lAuxiliar, lDigito, lMensagem) = False Then
                Call M46V999.gpMenErro(lMensagem)
            End If
        Else
            'ITEM 7 NOVA TARIFA 1074 / 1075 - CNPJ OBRIGATORIO NA COTA��O PARA CONDOMINIO
            If M46V999.gCod_Ramo = e_CodRamo_Condominio Then
                Call M46V999.gpMenErro("FNP0002 - CNPJ do segurado n�o preenchido.")
            End If
        End If
    Else
        lMensagem = "FNP9999 - Tipo de pessoa n�o informado."
        Call M46V999.gpMenErro(lMensagem)
    End If
    Set clsA99V001 = Nothing


    'If pstcPed.Cod_Bco_Cobr <> 341 And pstcPed.Num_Cta_Bas_Cor = 0 And pstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
    '    If pstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim And pstcPed.Qtd_Parc = 1 Then
    '        'Ignorar regra do banco
    '        'antecipado com uma parcela
    '    Else
    '        Call M46V999.gpMenErro("Devido ao Decreto Lei n� 6.335 de 03/01/08 que altera a aliquota de IOF para 7,38%, no momento est� dispon�vel apenas o Banco Ita� ( banco 341 ).")
    '    End If
    'End If
    'ARRUMAR AQUI
    If (objstcUsuario.TipoUsuario = e_Tipo_Usu_Corretor Or objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda) And pStcPed.StatusCotSiscota <> enum_StatusTransm_Aprovada And pStcPed.StatusCotSiscota <> enum_StatusTransm_EnvioKit Then
        For Each lobjstcJuros In pStcPed.ColecaoPedJuros
            lTip = lobjstcJuros.Tip_Compensacao
            Select Case lobjstcJuros.Num_Parcela
            Case 1
                If lobjstcJuros.Val_Ant < pStcPed.getProduto.ProJurMinAnt01(lTip) Then
                    GoSub ComparaJurosAnt
                End If
                If lobjstcJuros.Val_Pos < pStcPed.getProduto.ProJurMinPos01(lTip) Then
                    GoSub ComparaJurosPos
                End If
            Case 2
                If lobjstcJuros.Val_Ant < pStcPed.getProduto.ProJurMinAnt02(lTip) Then
                    GoSub ComparaJurosAnt
                End If
                If lobjstcJuros.Val_Pos < pStcPed.getProduto.ProJurMinPos02(lTip) Then
                    GoSub ComparaJurosPos
                End If
            Case 3
                If lobjstcJuros.Val_Ant < pStcPed.getProduto.ProJurMinAnt03(lTip) Then
                    GoSub ComparaJurosAnt
                End If
                If lobjstcJuros.Val_Pos < pStcPed.getProduto.ProJurMinPos03(lTip) Then
                    GoSub ComparaJurosPos
                End If
            Case 4
                If lobjstcJuros.Val_Ant < pStcPed.getProduto.ProJurMinAnt04(lTip) Then
                    GoSub ComparaJurosAnt
                End If
                If lobjstcJuros.Val_Pos < pStcPed.getProduto.ProJurMinPos04(lTip) Then
                    GoSub ComparaJurosPos
                End If
            Case 5
                If lobjstcJuros.Val_Ant < pStcPed.getProduto.ProJurMinAnt05(lTip) Then
                    GoSub ComparaJurosAnt
                End If
                If lobjstcJuros.Val_Pos < pStcPed.getProduto.ProJurMinPos05(lTip) Then
                    GoSub ComparaJurosPos
                End If
            Case 6
                If lobjstcJuros.Val_Ant < pStcPed.getProduto.ProJurMinAnt06(lTip) Then
                    GoSub ComparaJurosAnt
                End If
                If lobjstcJuros.Val_Pos < pStcPed.getProduto.ProJurMinPos06(lTip) Then
                    GoSub ComparaJurosPos
                End If
            Case 7
                If lobjstcJuros.Val_Ant < pStcPed.getProduto.ProJurMinAnt07(lTip) Then
                    GoSub ComparaJurosAnt
                End If
                If lobjstcJuros.Val_Pos < pStcPed.getProduto.ProJurMinPos07(lTip) Then
                    GoSub ComparaJurosPos
                End If
            Case 8
                If lobjstcJuros.Val_Ant < pStcPed.getProduto.ProJurMinAnt08(lTip) Then
                    GoSub ComparaJurosAnt
                End If
                If lobjstcJuros.Val_Pos < pStcPed.getProduto.ProJurMinPos08(lTip) Then
                    GoSub ComparaJurosPos
                End If
            Case 9
                If lobjstcJuros.Val_Ant < pStcPed.getProduto.ProJurMinAnt09(lTip) Then
                    GoSub ComparaJurosAnt
                End If
                If lobjstcJuros.Val_Pos < pStcPed.getProduto.ProJurMinPos09(lTip) Then
                    GoSub ComparaJurosPos
                End If
            Case 10
                If lobjstcJuros.Val_Ant < pStcPed.getProduto.ProJurMinAnt10(lTip) Then
                    GoSub ComparaJurosAnt
                End If
                If lobjstcJuros.Val_Pos < pStcPed.getProduto.ProJurMinPos10(lTip) Then
                    GoSub ComparaJurosPos
                End If
            Case 11
                If lobjstcJuros.Val_Ant < pStcPed.getProduto.ProJurMinAnt11(lTip) Then
                    GoSub ComparaJurosAnt
                End If
                If lobjstcJuros.Val_Pos < pStcPed.getProduto.ProJurMinPos11(lTip) Then
                    GoSub ComparaJurosPos
                End If
            Case 12
                If lobjstcJuros.Val_Ant < pStcPed.getProduto.ProJurMinAnt12(lTip) Then
                    GoSub ComparaJurosAnt
                End If
                If lobjstcJuros.Val_Pos < pStcPed.getProduto.ProJurMinPos12(lTip) Then
                    GoSub ComparaJurosPos
                End If
            End Select
        Next
    End If

    'Solicita��o n� 0801751

    'If pStcPed.Dat_Ini_Vig <> 0 And pStcPed.Dat_Ter_Vig <> 0 And (objstcUsuario.TipoUsuario = e_Tipo_Usu_Corretor Or objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda) And pStcPed.Ind_CotacaoSolic = True Then
    'MOSILVA: Ver com Thais!!! (Consistencia estava sendo feita s� p/ recota��o!)
    If pStcPed.Dat_Ini_Vig <> 0 And pStcPed.Dat_Ter_Vig <> 0 And (objstcUsuario.TipoUsuario = e_Tipo_Usu_Corretor Or objstcUsuario.TipoUsuario = e_Tipo_Usu_CorretorYasuda) Then
        lQuaDias = M46V999.gfDifDatas(pStcPed.Dat_Ini_Vig, pStcPed.Dat_Ter_Vig)
        If lQuaDias = 366 And Mid(pStcPed.Dat_Ini_Vig, 5, 4) = Mid(pStcPed.Dat_Ter_Vig, 5, 4) Then
            lQuaDias = 365
        End If
        
        If Not (pStcPed.objstcCtrlEmis.Dat_Ini_Vig_Orig = 0 Or pStcPed.objstcCtrlEmis.Dat_Ter_Vig_Orig = 0) Then
            lQuaDias_Orig = M46V999.gfDifDatas(pStcPed.objstcCtrlEmis.Dat_Ini_Vig_Orig, pStcPed.objstcCtrlEmis.Dat_Ter_Vig_Orig)
            If lQuaDias_Orig = 366 And Mid(pStcPed.objstcCtrlEmis.Dat_Ini_Vig_Orig, 5, 4) = Mid(pStcPed.objstcCtrlEmis.Dat_Ter_Vig_Orig, 5, 4) Then
                lQuaDias_Orig = 365
            End If
        End If
    
        'Tarifa 11.38 Vig�ncia Pr�-Rata Flexibiliza��o
        'Tarifa 11.40 Endossos n�o entram nesta condi��o
        If pStcPed.Cod_Produto >= 1138 And M46V999.gEhRetornoSiscota(pStcPed.NOSSO_NUMERO) And pStcPed.Cod_Dizer = 0 And pStcPed.Cod_Ramo <> e_CodRamo_Residencial And _
           (pStcPed.gSegmentado Or pStcPed.Cod_Estip = 10 Or pStcPed.Cod_Estip = 11 Or pStcPed.Cod_Estip = 20) And (pStcPed.objstcCtrlEmis.Dat_Ini_Vig_Orig <> 0 And pStcPed.objstcCtrlEmis.Dat_Ter_Vig_Orig <> 0) Then
            If lQuaDias_Orig <> lQuaDias Then
                Call M46V999.gpMenErro("FNP0011 - Vig�ncia inv�lida. Per�odo diferente do acordado.")
            End If
        Else
            If lQuaDias <> 365 And pStcPed.Cod_Dizer = 0 Then
                Call M46V999.gpMenErro("FNP0011 - Vig�ncia inv�lida. Deve ser anual.")
            End If
        End If

        '11.46 - Empresarial item 4
        If pStcPed.ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_RenYasuda Then
            If pStcPed.Cod_Produto >= 978 Then
                lValidade = M46V998.gfCalcData(pStcPed.Dat_Pri_Calculo, 15)
            Else
                lValidade = M46V998.gfCalcData(pStcPed.Dat_Pri_Calculo, 30)
            End If
        Else
			'Tranferencia da regra de 917581 - Triunfo para 939571 - BR Insurance - 1174
            If pStcPed.Cod_Ramo = e_CodRamo_Condominio And M46V999.gCod_Corretor = 939571 And (pStcPed.ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_RenCongenere Or pStcPed.ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_SeguroNovo) Then
                lValidade = M46V998.gfCalcData(pStcPed.Dat_Pri_Calculo, 30)
            Else
                lValidade = M46V998.gfCalcData(pStcPed.Dat_Pri_Calculo, 15)
            End If
        End If
        
        If pStcPed.Dat_Ini_Vig > lValidade And lValidade <> 0 Then
            'Call M46V999.gpMenErro("FNP0011 - Vig�ncia inv�lida. Deve ser menor que o prazo da cota��o.")
        End If
    End If

    If pStcPed.Dat_Emis > 0 Then
        If M46V999.gfDifDatas(M46V999.gHoje, pStcPed.Dat_Ini_Vig) > 45 Then
            Call M46V999.gpMenErro("FGE0135 - S� � permitida agendar uma emiss�o emiss�o com at� 45 dias de anteced�ncia..")
        End If
        If M46V999.gfDifDatas(pStcPed.Dat_Ini_Vig, pStcPed.Dat_Emis) > 5 Then
            Call M46V999.gpMenErro("FGE0135 - Data da emiss�o agendada acima de cinco dias do in�cio da vig�ncia.")
        End If
    End If
    lI = 0

    Dim lB As Integer
    lB = 0

    Dim lstcPedBenefiario As stcA46V735B

    For Each lobjstcItem In pStcPed.ColecaoPedLoc
        If lobjstcItem.ColecaoBeneficiario.Count > 0 Then
            lB = 1
            Exit For
        End If
    Next lobjstcItem

    If pStcPed.Cod_Tip_Doc = e_TipDoc_Proposta And (pStcPed.Cod_Estip = 2168 Or pStcPed.Cod_Estip = 2169 Or pStcPed.Cod_Estip = 2157) Then
        If pStcPed.Cod_Dizer <> 4 And pStcPed.Cod_Dizer <> 10 Then
            For Each lobjstcEspec In pStcPed.ColecaoTextoEspec
                If lobjstcEspec.Cod_Espec = 8500 Then
                    lI = 1
                    Exit For
                End If
            Next lobjstcEspec
            If lI = 0 And lB = 0 Then
                Call M46V999.gpMenErro("FGE0135 - Benefici�rio n�o informado.")
            End If
        End If

    End If

    'Obrigatoriedade do campo data de emiss�o do documento segurado
    If pStcPed.Dat_Emis_Doc = 0 And pStcPed.Cod_Tip_Doc = e_TipDoc_Proposta And pStcPed.Tip_Pessoa = e_TipPessoa_Fisica Then
        lMensagem = "FGE0136 - Data de Emiss�o do Documento n�o preenchida. "
        Call M46V999.gpMenErro(lMensagem)
    End If

    For i = 1 To pStcPed.ColecaoPedLoc.Count
        If pStcPed.Cod_Dizer = 4 Then
            If pStcPed.ColecaoPedLoc(i).Cod_DizerItem = 1 Then
                lMensagem = "FGE0137 - N�o � permitido Exclus�o ou Inclus�o de Itens no mesmo endosso - Favor realizar as opera��es serparadas. "
                Call M46V999.gpMenErro(lMensagem)
            End If
        ElseIf pStcPed.Cod_Dizer = 1 Then
            If pStcPed.ColecaoPedLoc(i).Cod_DizerItem = 4 Then
                lMensagem = "FGE0137 - N�o � permitido Exclus�o ou Inclus�o de Itens no mesmo endosso - Favor realizar as opera��es serparadas. "
                Call M46V999.gpMenErro(lMensagem)
            End If
        End If
        
        'Conge 11.22
        If pStcPed.ColecaoPedLoc(i).Tip_Emissao = e_TipEmissao_RenCongenere Then

            'Tarifa 1176 - 8 Inclus�o de Cl�usula de Renova��o Cong�nere no Produto Condom�nio
            If M46V999.gCodProduto >= 1122 And (pStcPed.Cod_Ramo = e_CodRamo_Residencial Or pStcPed.Cod_Ramo = e_CodRamo_Condominio) Then '11.22
                lAuxMsg = "As informa��es apresentadas nesta proposta s�o fatores determinantes para aceita��o do risco,"
                lAuxMsg = lAuxMsg & " podendo conduzir a redu��o do custo do pr�mio de seguro. Se constatado que as respostas"
                lAuxMsg = lAuxMsg & " n�o foram verdadeiras o segurado participar� dos preju�zos,"
                lAuxMsg = lAuxMsg & " ""aplicando a proporcionalidade no valor da indeniza��o"", calculada em fun��o do"
                lAuxMsg = lAuxMsg & " coeficiente obtido entre pr�mio pago e o pr�mio devido, conforme os artigos 765 e 766"
                lAuxMsg = lAuxMsg & " do C�digo Civil Brasileiro. Em caso de dolo ou m� f� o segurado poder� perder direito"
                lAuxMsg = lAuxMsg & " a indeniza��o de seguro"
            End If
    
            lAuxMsg = M46V999.gfRetirarCaracterEspecial(lAuxMsg, True)
            
            txtobs = UCase(pStcPed.ColecaoPedLoc(i).objstcPedObs.Observacao)
            If InStr(txtobs, UCase(lAuxMsg)) > 0 Then
                lAuxMsg = txtobs
            Else
                lAuxMsg = UCase(lAuxMsg) & vbCrLf & txtobs
            End If
            pStcPed.ColecaoPedLoc(i).objstcPedObs.Observacao = lAuxMsg
    
            'Verifico se deve ser gravada a observa��o de renova��o.
            pStcPed.ColecaoPedLoc(i).objstcPedObs.Observacao = Replace(lAuxMsg, vbCrLf, "&%")
            Set lobjPedObs = New clsA46V714A
            Set lobjPedObs.M46V999 = M46V999
    
            Set lobjPedObs.ConexaoBD = M46V999.mConexaoSQL
            If Not lobjPedObs.mfExcluir(pStcPed.ColecaoPedLoc(i).objstcPedObs, objstcUsuario.PropTabela, "") Then
                lMensagem = lobjPedObs.UltimoErro
                Call M46V999.gpMenErro("FNP1002 - " & lMensagem)
            End If
            With pStcPed.ColecaoPedLoc(i).objstcPedObs
                .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(i).NOSSO_NUMERO
                .Num_Item = pStcPed.ColecaoPedLoc(i).Num_Item
            End With
            If Not lobjPedObs.mfIncluir(pStcPed.ColecaoPedLoc(i).objstcPedObs, objstcUsuario.PropTabela, "", "") Then
                lMensagem = lobjPedObs.UltimoErro
                Call M46V999.gpMenErro("FNP1001 - " & lMensagem)
            End If
        End If
        
    Next

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Sub

ComparaJurosAnt:
    'CRQ000000458062 - N�o validar juros em volta de Classe3/4
    If Not (pStcPed.StatusCotSiscota = enum_StatusTransm_Aprovada Or pStcPed.StatusCotSiscota = enum_StatusTransm_EnvioKit) Then
    '1001224 - Condi��es Diferenciada - Seguratec
    If Not M46V111.gfBuscaJurosParm(pStcPed.ColecaoPedLoc(1), lobjstcJuros.Num_Parcela, 1, lobjstcJuros.Tip_Compensacao, lJuros, lMensagem) Then
        Call M46V999.gpMenErro("FGE0010 - Juros Inv�lido para " & lobjstcJuros.Num_Parcela & " parcelas (Antecipado)." & lMensagem)
    End If
                        
    If lobjstcJuros.Val_Ant < lJuros Then
        Call M46V999.gpMenErro("FGE0011 - Juros abaixo do permitido para " & lobjstcJuros.Num_Parcela & " parcelas (Antecipado)." & lMensagem)
    End If
    Return
    End If

ComparaJurosPos:
    'CRQ000000458062 - N�o validar juros em volta de Classe3/4
    If Not (pStcPed.StatusCotSiscota = enum_StatusTransm_Aprovada Or pStcPed.StatusCotSiscota = enum_StatusTransm_EnvioKit) Then
    '1001224 - Condi��es Diferenciada - Seguratec
    If Not M46V111.gfBuscaJurosParm(pStcPed.ColecaoPedLoc(1), lobjstcJuros.Num_Parcela, 2, lobjstcJuros.Tip_Compensacao, lJuros, lMensagem) Then
        Call M46V999.gpMenErro("FGE0010 - Juros Inv�lido para " & lobjstcJuros.Num_Parcela & " parcelas (Postecipado)." & lMensagem)
    End If
        
    If lobjstcJuros.Val_Pos < lJuros Then
        Call M46V999.gpMenErro("FGE0011 - Juros abaixo do permitido para " & lobjstcJuros.Num_Parcela & " parcelas (Postecipado)." & lMensagem)
    End If
    Return
    End If

FechaObj:
    Set clsA46V102 = Nothing
    Set lObjstcVenc = Nothing
    Set lobjstcCobert = Nothing
    Set lObjstcCorr = Nothing
    Set lobjstcCoss = Nothing
    Set lobjstcItem = Nothing
    Set lobjstcEspec = Nothing
    Set clsA99V001 = Nothing
    Set lPerguntaQuest = Nothing
    Set lobjstcJuros = Nothing
    Return

End Sub

Private Sub mpAtualizaJurosNovo(ByRef pStcPed As stcA46V702B, tagDynaTrace As String)
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mpAtualizaJurosNovo"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lPerJuros As Single
    Dim lI As Integer
    Dim lI2 As Integer
    Dim lobjstcJuros As stcA46V711B    'Objeto da classe de estrutura stcA46V711B (TAB_PED_JUROS).
    Dim lobjPedJuros As clsA46V711A

    'Juros por parcela.
    Dim MaxAntec As Integer, MaxPosPg As Integer
    MaxAntec = 4  'M�ximo permitido de Parcelas sem juros
    MaxPosPg = 3

    Set lobjPedJuros = New clsA46V711A
    Set lobjPedJuros.M46V999 = M46V999

    If objstcUsuario.codCorretor = 1601 Or objstcUsuario.codCorretor = 0 Then
        'Diferenciado para Seguratec - 15/02/2007 - Fernando
        MaxAntec = pStcPed.getProduto.ProQuaMaxParcelas
        MaxPosPg = pStcPed.getProduto.ProQuaMaxParcelas
    ElseIf objstcUsuario.codCorretor = 435241 Or objstcUsuario.codCorretor = 288911 Then
        'Diferenciado para Avant Marine 435241 e Marine 288911 - 15/02/2007 - Fernando
        MaxAntec = 5
        MaxPosPg = 4
        'Ficha 901507
    ElseIf (objstcUsuario.codCorretor = 8001 Or objstcUsuario.codCorretor = 8002 Or objstcUsuario.codCorretor = 8003 Or objstcUsuario.codCorretor = 8004 Or _
            objstcUsuario.codCorretor = 8005 Or objstcUsuario.codCorretor = 8006 Or objstcUsuario.codCorretor = 8008 Or objstcUsuario.codCorretor = 8011 Or _
            objstcUsuario.codCorretor = 8014 Or objstcUsuario.codCorretor = 8028) And pStcPed.Cod_Ramo = e_CodRamo_Residencial Then
        'Diferenciado para Volks - 15/09/2009
        MaxAntec = 9
        MaxPosPg = 9
    End If

    'Ficha 901772

    If pStcPed.Cod_Ramo = e_CodRamo_Residencial And pStcPed.Cod_Estip = M46V999.gfPlanoBasico(pStcPed.Cod_Ramo) And objstcUsuario.codCorretor = 1601 Then
        If (pStcPed.ColecaoPedLoc(1).Cod_Clasf = "001" And pStcPed.ColecaoPedLoc(1).Cod_Compl_Clasf = "11") Or _
           (pStcPed.ColecaoPedLoc(1).Cod_Clasf = "001" And pStcPed.ColecaoPedLoc(1).Cod_Compl_Clasf = "21") Then
            'Diferenciado para Seguratec - 23/11/2009
            MaxAntec = 5
            MaxPosPg = 4
        End If
    End If

    If objstcUsuario.TipoAcesso = DIGITACAO Then
        lPerJuros = pStcPed.ColecaoPedLoc(1).Per_Juro_Frac
    Else
        lPerJuros = pStcPed.getProduto.ProJurMinPadrao
    End If

    If pStcPed.ColecaoPedJuros.Count = 0 Then
        If Not mfPosicionaJuros(pStcPed, tagDynaTrace) Then
            GoTo SAIDA
        End If
    Else
        lI = 1
        'Quando for digita��o ent�o n�o reposiciono o juros.
        'CRQ000000458062 - N�o reposiciona juros em volta de Classe3/4
        If objstcUsuario.TipoAcesso <> DIGITACAO And Not (pStcPed.StatusCotSiscota = enum_StatusTransm_Aprovada Or pStcPed.StatusCotSiscota = enum_StatusTransm_EnvioKit) Then
            If Not mfPosicionaJuros(pStcPed, tagDynaTrace) Then
                GoTo SAIDA
            End If
        Else
            If Not (pStcPed.StatusCotSiscota = enum_StatusTransm_Aprovada Or pStcPed.StatusCotSiscota = enum_StatusTransm_EnvioKit) Then
            Set pStcPed.ColecaoPedJuros = New Collection
            For lI = 1 To 11    '<- originalmente 10
                Set lobjstcJuros = New stcA46V711B
                With lobjstcJuros
                    .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
                    .Num_Parcela = lI
                    If lI > 1 Then
                        .Val_Ant = CDbl(lPerJuros)
                        .Val_Pos = CDbl(lPerJuros)
                    Else
                        .Val_Ant = 0
                        .Val_Pos = CDbl(lPerJuros)
                    End If
                End With
                If Not lobjPedJuros.mfIncluir(lobjstcJuros, objstcUsuario.PropTabela, "", tagDynaTrace) Then
                    Call M46V999.gpMenErro(lobjPedJuros.UltimoErro)
                    GoTo SAIDA
                End If
                pStcPed.ColecaoPedJuros.Add lobjstcJuros, lobjstcJuros.Class_ID
            Next
        End If
    End If
    End If


    'Ficha 901673 e 901793
    If pStcPed.Cod_Ramo = e_CodRamo_Residencial And _
       (pStcPed.Cod_Produto >= 974 And pStcPed.Cod_Produto <= 979) And _
       pStcPed.Cod_Estip = M46V999.gfPlanoBasico(pStcPed.Cod_Ramo) And objstcUsuario.TipoAcesso = CALCULO Then
        lPerJuros = 0
        Set pStcPed.ColecaoPedJuros = New Collection
        For lI2 = 1 To M46V999.gQtdMaxComp
            For lI = 1 To 11    '<- originalmente 10
                Set lobjstcJuros = New stcA46V711B
                With lobjstcJuros
                    .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
                    .Num_Parcela = lI
                    .Tip_Compensacao = lI2
                    .Val_Ant = CDbl(lPerJuros)
                    .Val_Pos = CDbl(lPerJuros)
                End With
                If Not lobjPedJuros.mfIncluir(lobjstcJuros, objstcUsuario.PropTabela, "", tagDynaTrace) Then
                    Call M46V999.gpMenErro(lobjPedJuros.UltimoErro)
                    GoTo SAIDA
                End If
                pStcPed.ColecaoPedJuros.Add lobjstcJuros, lobjstcJuros.Class_ID
            Next
        Next
    End If

SAIDA:
    Err.Clear: On Error GoTo 0
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    GoSub FechaObj
    Exit Sub

FechaObj:
    Set lobjstcJuros = Nothing
    Set lobjPedJuros = Nothing
    Return
End Sub

Private Function mfConsTotBens(ByVal pValLMG As Double, ByVal pCodCobert As Long, pStcPedLoc As stcA46V708B, tagDynaTrace As String)
    Dim pValBens As Double
    Dim lI As Integer    'Utilizado no for...loop.
    Dim lObjstcBens As stcA46V715B

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfConsTotBens"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    'Verificar se a soma dos Bens � igual ao LMG - para plano All Risks
    pValBens = 0

    If pStcPedLoc.ObjPai.Cod_Produto >= 947 And pStcPedLoc.ObjPai.Cod_Ramo = e_CodRamo_Roubos Then

        For Each lObjstcBens In pStcPedLoc.ColecaoPedLocBem
            pValBens = pValBens + lObjstcBens.Val_Bem
        Next lObjstcBens

        'Ficha 1000308 - Permitir itens com planos diferentes no produto Roubo (c�lculo coletivo)
        If pStcPedLoc.Cod_Plano = 1 Then
            If pValBens <> pValLMG Then
                Call M46V999.gpMenErro("FNP0074 - " & M46V999.gPrefixo & "Somat�ria dos bens deve ser igual ao LMG.")
            End If
        Else
            If pValBens > pValLMG Then
                Call M46V999.gpMenErro("FNP0074 - " & M46V999.gPrefixo & "Somat�ria dos bens n�o deve ser maior que o LMG total.")
            End If
        End If

    End If

    If pStcPedLoc.ObjPai.Cod_Produto >= 999 And (pStcPedLoc.ObjPai.Cod_Ramo = e_CodRamo_Empresarial Or pStcPedLoc.ObjPai.Cod_Ramo = e_CodRamo_Residencial) And pStcPedLoc.Cod_Produto = 1 Then
        If pStcPedLoc.Cod_DizerItem <> 4 And pStcPedLoc.Cod_DizerItem <> 10 Then
            For Each lObjstcBens In pStcPedLoc.ColecaoPedLocBem
                pValBens = pValBens + lObjstcBens.Val_Bem
            Next lObjstcBens
            If pValBens <> pValLMG Then
                Call M46V999.gpMenErro("FNP0074 - " & M46V999.gPrefixo & "O somat�rio dos valores da rela��o de bens deve ser igual ao LMG da cobertura " & pCodCobert & ".")
            End If
        End If
    End If

    mfConsTotBens = True

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function

ControleErro:
    mfConsTotBens = False
    GoTo SAIDA

FechaObj:
    Set lObjstcBens = Nothing
    Return
End Function

Private Sub mpVerIteNiv_2(pobjstcItem As stcA46V708B, tagDynaTrace As String)
'Procedure: consiste item n�vel 2: consiste todos dados do item.

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mpVerIteNiv_2"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lobjA46V103A As clsA46V103A

    Set lobjA46V103A = New clsA46V103A
    Set lobjA46V103A.M46V103 = M46V103

    Call lobjA46V103A.gpVerIteNiv_2(pobjstcItem, pobjstcItem.ObjPai, pobjstcItem.ObjPai.Val_IS, True, objstcUsuario, tagDynaTrace)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Sub

FechaObj:
    Set lobjA46V103A = Nothing
    Return
End Sub

Private Sub mpValidaParcela(ByRef pStcPed As stcA46V702B, ByRef pColAlertas As Collection, ByRef pColErros As Collection)
    Dim lStcPedVenc As stcA46V713B
    Dim lStcParcela As stcParcelamento
    Dim lMaxParcela As Integer  'N�mero m�ximo de parcelas permitido
    Dim lMaxParcelaAux As Integer     'N�mero m�ximo de parcelas permitido
    Dim lMaxParcelaVista As Integer       'N�mero m�ximo de parcelas permitido a vista
    Dim lMaxParcela30Dia As Integer       'N�mero m�ximo de parcelas permitido 30 dias
    Dim lDataUltVencimento As Double          'Data do ultimo vencimento possivel
    Dim lQutdeParcelas As Integer     'Quantidade de parcelas
    Dim lstrIniVig As String
    Dim lstrTerVig As String
    Dim lMsgErr As String
    Dim lMaxParcPermitida As Byte
    Dim lArrayRetricaoCorr As Variant
    Dim i As Integer

    'Determina quantide m�xima de parcelas
    lstrIniVig = pStcPed.Dat_Ini_Vig
    lstrTerVig = pStcPed.Dat_Ter_Vig

    If pStcPed.gPagAntecipado = True Then
        lMaxParcPermitida = pStcPed.gQtdMaxParcAnt
    Else
        lMaxParcPermitida = pStcPed.gQtdMaxParcPos
    End If

    Call M46V998.mpObtemMaxParcelas(lMaxParcPermitida, lMaxParcelaVista, lMaxParcela30Dia, lstrIniVig, lstrTerVig, lDataUltVencimento)

    'Informa o maior valor de parcelas
    If lMaxParcelaVista > lMaxParcela30Dia Then
        lMaxParcela = lMaxParcelaVista
    Else
        lMaxParcela = lMaxParcela30Dia
    End If

    'Guarda na variavel o numero de parcelas informado
    lQutdeParcelas = pStcPed.Qtd_Parc

    'caso seja proposta limita o numero maximo para o tipo
    'escolhido (antecipado ou 30dias)

    'Ficha 900665
    '    If pstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
    If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
        lMaxParcela = lMaxParcelaVista
    Else
        lMaxParcela = lMaxParcela30Dia
    End If
    'ENDOSSO
    'Ap�s o c�lculo � que ser� consistido se a postecipada � permitida ou n�o.
    'atrav�s da vari�vel lMaxParcelaAux.
    lMaxParcelaAux = lMaxParcela
    If lMaxParcela = 0 And pStcPed.ColecaoPedLoc(1).Tip_Emissao >= 300 Then
        lMaxParcela = 1
    End If
    '    End If

    'Plano Mercedes
    'pStcPed.Cod_Estip
    'If M46V111.gCod_Plano >= 2021 And M46V111.gCod_Plano <= 2031 Then
    If pStcPed.Cod_Estip >= 2021 And pStcPed.Cod_Estip <= 2031 Then
        lMaxParcela = 11
    End If
    If pStcPed.Cod_Estip >= 2326 And pStcPed.Cod_Estip <= 2349 Then
        lMaxParcela = 11
    End If
    If pStcPed.Cod_Estip >= 2350 And pStcPed.Cod_Estip <= 2373 Then
        lMaxParcela = 11
    End If
    ' Plano RVA
    If (pStcPed.Cod_Estip = 2303 And pStcPed.Cod_Ramo = 112) Or _
       (pStcPed.Cod_Estip = 2374 And pStcPed.Cod_Ramo = 113) Then
        lMaxParcela = 11
    End If

    If pStcPed.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        'ajusta o numero de parcelas caso retorne 0
        If lMaxParcela <= 0 Then
            lMaxParcela = 1
        End If
        'Se = Cota��o ent�o atualiza o campo no form
        If pStcPed.Qtd_Parc > lMaxParcela Then
            pStcPed.Qtd_Parc = lMaxParcela
        End If
    End If

    If (pStcPed.Cod_Tip_Doc = e_TipDoc_Proposta) And _
       (lQutdeParcelas > lMaxParcela) Then
        'Verifica se o numero informado � maior que o permitido
        If pStcPed.ColecaoPedVenc.Count > 0 And objstcUsuario.TipoUsuario = e_Tipo_Usu_Yasuda Then
            For Each lStcPedVenc In pStcPed.ColecaoPedVenc
                If lStcPedVenc.DAT_VENC > pStcPed.Dat_Ter_Vig Then
                    lMsgErr = "FGE0075 - Quantidade de parcelas inv�lida (Data da �ltima parcela n�o deve ultrapassar " & _
                              M46V999.gfForData(2, lStcPedVenc.DAT_VENC) & ". " & val(M46V999.gfDifDatas(lDataUltVencimento, M46V999.gfForData(1, Format(CDate(M46V999.gfForData(2, pStcPed.Dat_Ter_Vig)), "dd/mm/yyyy")))) & _
                              " dias antes do t�mino da vig�ncia  )." & vbCrLf & "N�mero m�ximo de parcelas permitido: " & lMaxParcela
                    Call M46V999.gpMenErro(lMsgErr)
                End If
            Next lStcPedVenc
            lMaxParcela = lQutdeParcelas
        End If
    End If

    If (pStcPed.Cod_Tip_Doc = e_TipDoc_Proposta) Then
        If pStcPed.Qtd_Parc > lMaxParcela Then
            'Verifica se o numero informado � maior que o permitido
            lMsgErr = "FGE0075 - Quantidade de parcelas inv�lida. N�mero m�ximo de parcelas permitido: " & lMaxParcela
            Call M46V999.gpMenErro(lMsgErr)
            If Not (pStcPed.ColecaoPedVenc.Count > 0 And objstcUsuario.TipoUsuario = e_Tipo_Usu_Yasuda) Then
                If lDataUltVencimento = 0 Then
                    lMsgErr = "FGE0075 - Quantidade de parcelas inv�lida (Data da �ltima parcela deve ser inferior a data de t�rmino da vig�ncia)." & vbCrLf & "N�mero m�ximo de parcelas permitido: " & lMaxParcela
                    Call M46V999.gpMenErro(lMsgErr)
                Else
                    If val(M46V999.gfDifDatas(lDataUltVencimento, M46V999.gfForData(1, Format(CDate(M46V999.gfForData(2, pStcPed.Dat_Ter_Vig)), "dd/mm/yyyy")))) > 0 Then
                        lMsgErr = "FGE0075 - Quantidade de parcelas inv�lida (Data da �ltima parcela n�o deve ultrapassar " & _
                                  M46V999.gfForData(2, lDataUltVencimento) & ". " & val(M46V999.gfDifDatas(lDataUltVencimento, M46V999.gfForData(1, Format(CDate(M46V999.gfForData(2, pStcPed.Dat_Ter_Vig)), "dd/mm/yyyy")))) & _
                                  " dias antes do t�mino da vig�ncia  )." & vbCrLf & "N�mero m�ximo de parcelas permitido: " & lMaxParcela
                        Call M46V999.gpMenErro(lMsgErr)
                    Else
                        lMsgErr = "FGE0075 - Quantidade de parcelas inv�lida (Data da �ltima parcela deve ser inferior a data de t�rmino da vig�ncia)." & vbCrLf & "N�mero m�ximo de parcelas permitido: " & lMaxParcela
                        Call M46V999.gpMenErro(lMsgErr)
                    End If
                End If
            End If
        Else
            'Verifica se o numero informado � um parcelamento V�lido
            For Each lStcParcela In pStcPed.ColecaoParcelamento
                If pStcPed.Qtd_Parc = lStcParcela.Num_Parcela Then
                    If lStcParcela.EhParcelaValida = False Then
                        lMsgErr = "FGE0075 - Quantidade de parcelas inv�lida. Parcelamento permitido apenas para " & pStcPed.ParcelasValidas
                        Call M46V999.gpMenErro(lMsgErr)
                    Else
                        'Valida pagamento antecipado
                        If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 And lStcParcela.Val_Primeira_Vista = 0 Then
                            lMsgErr = "FGE0075 - Quantidade de parcelas(" & lStcParcela.Num_Parcela & ") inv�lida para pagamento Antecipado. Parcelamento permitido apenas para " & pStcPed.ParcelasValidas
                            Call M46V999.gpMenErro(lMsgErr)
                        Else
                            'Valida pagamento postecipada

                            'Pagamento postecipado excluido com exce��o de Planos e Segmenta��o
                            
                            'Item 1 Nova tarifa 1086 / 1087
                            If pStcPed.Cod_Produto >= 1086 Then
                                If pStcPed.Cod_Ramo = e_CodRamo_Empresarial Then
                                    lArrayRetricaoCorr = Split("1675;17048;723911;860301", ";")
                                Else
                                    lArrayRetricaoCorr = Split("860301", ";")
                                End If
                                
                                For i = 0 To UBound(lArrayRetricaoCorr)
                                    If M46V999.gCod_Corretor = lArrayRetricaoCorr(i) Then
                                        GoTo NaoValida
                                    End If
                                Next
                                
                            End If
                            
                            If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 2 And _
                               lStcParcela.Val_Primeira_30dias = 0 And _
                               M46V111.gCod_plano < 1000 And _
                               M46V999.gCodProduto >= 1072 And _
                              ((pStcPed.Cod_Dizer = 0 And M46V999.gCodProduto < 1080) Or _
                               (pStcPed.Cod_Dizer >= 0 And pStcPed.ColecaoPedLoc(1).Tip_Emissao <> 200 And pStcPed.objstcCtrlEmis.cod_origem <> 2 And _
                                M46V999.gCodProduto >= 1080 And pStcPed.Cod_Ramo <> e_CodRamo_Condominio)) Then

                                If M46V999.gCodProduto < 1074 Then
                                    lMsgErr = "FGE0075 - Pagamento Postecipado n�o permitido."
                                    Call M46V999.gpMenErro(lMsgErr)
                                Else
                                    If ((M46V999.gCod_Corretor <> 1601) Or (M46V999.gCod_Corretor = 1601 And M46V999.gCod_Ramo <> e_CodRamo_Residencial)) And _
                                       ((M46V999.gCod_Corretor <> 383111) Or (M46V999.gCod_Corretor = 383111 And M46V999.gCod_Ramo = e_CodRamo_Condominio)) Then
                                        lMsgErr = "FGE0075 - Pagamento Postecipado n�o permitido."
                                        Call M46V999.gpMenErro(lMsgErr)
                                    End If
                                End If
                            End If

                            If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 2 And lStcParcela.Val_Primeira_30dias = 0 Then
                                lMsgErr = "FGE0075 - Quantidade de parcelas(" & lStcParcela.Num_Parcela & ") inv�lida para pagamento Postecipado. Parcelamento permitido apenas para " & pStcPed.ParcelasValidas & ". Observado o valor m�nimo da parcela."
                                Call M46V999.gpMenErro(lMsgErr)
                            End If
                        End If
                    End If
NaoValida:
                    Exit For
                End If
            Next
        End If
    End If

SAIDA:
    GoSub FechaObj
    Exit Sub

FechaObj:
    Set lStcPedVenc = Nothing
    Set lStcParcela = Nothing
    Return
End Sub

Private Function mfPosicionaJuros(ByRef pStcPed As stcA46V702B, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfPosicionaJuros"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    
    Dim lJurosAnt As Single
    Dim lJurosPos As Single
    Dim lMensagem As String
    Dim lI As Integer
    Dim lI2 As Integer
    Dim lobjstcJuros As stcA46V711B
    Dim lobjPedJuros As clsA46V711A

    Set pStcPed.ColecaoPedJuros = New Collection
    Set lobjPedJuros = New clsA46V711A
    Set lobjPedJuros.M46V999 = M46V999

    Set lobjstcJuros = New stcA46V711B
    lobjstcJuros.NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
    If Not lobjPedJuros.mfExcluir(lobjstcJuros, objstcUsuario.PropTabela, "", tagDynaTrace) Then
        Call M46V999.gpMenErro(lobjPedJuros.UltimoErro)
        GoTo SAIDA
    End If
    For lI2 = 1 To M46V999.gQtdMaxComp
        For lI = 1 To 11
            '1001224 - Condi��es Diferenciada - Seguratec
            If Not M46V111.gfBuscaJurosParm(pStcPed.ColecaoPedLoc(1), lI, 1, lI2, lJurosAnt, lMensagem) Then
                Call M46V999.gpMenErro(lMensagem)
            End If
            If Not M46V111.gfBuscaJurosParm(pStcPed.ColecaoPedLoc(1), lI, 2, lI2, lJurosPos, lMensagem) Then
                Call M46V999.gpMenErro(lMensagem)
            End If
            Set lobjstcJuros = New stcA46V711B
            With lobjstcJuros
                .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
                .Num_Parcela = lI
                .Tip_Compensacao = lI2
                .Val_Ant = CDbl(lJurosAnt)
                .Val_Pos = CDbl(lJurosPos)
            End With
            If Not lobjPedJuros.mfIncluir(lobjstcJuros, objstcUsuario.PropTabela, "", tagDynaTrace) Then
                Call M46V999.gpMenErro(lobjPedJuros.UltimoErro)
                GoTo SAIDA
            End If
            pStcPed.ColecaoPedJuros.Add lobjstcJuros, lobjstcJuros.Class_ID
        Next lI
    Next lI2

    mfPosicionaJuros = True

SAIDA:
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

FechaObj:
    Set lobjstcJuros = Nothing
    Set lobjPedJuros = Nothing
    Return
End Function

Private Sub mpAddAlerta(ByVal pIDConsist As E_IDConsist, _
                        ByVal pMsgErr As String, _
                        ByVal pTipoConsistencia As E_TipConsist, _
                        ByRef pColAlertas As Collection, _
                        Optional ByVal pNumItem As Long = 0)
    Dim lConsist As stcConsistPreCalc

    Set lConsist = New stcConsistPreCalc
    With lConsist
        .NumItem = pNumItem
        .IDConsistencia = pIDConsist
        .Mensagem = pMsgErr
        .TipoConsistencia = pTipoConsistencia
    End With
    pColAlertas.Add lConsist, lConsist.Class_ID

FechaObj:
    Set lConsist = Nothing

End Sub


Public Function gfExcluirItem(ByVal pNosNumero As String, ByVal pNumItem As Long, tagDynaTrace As String) As Boolean
'Procedure: exclui documento selecionada (grava N no campo Cod_Ativo).

    Dim lTotalExcluir As Long           'Total de registros a serem excluidos
    Dim lRegExcluir As Variant      'Lista com os nosso_numeros para excluir.
    Dim lCount As Long    'Utilizado no for ... loop
    Dim lNosNumero As String      'Nosso n�mero.
    Dim lNumItem As Long

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfExcluirItem"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro: gfExcluirItem = False

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        Exit Function
    End If

    If M46V999.CDblx(pNosNumero) = 0 Or pNumItem = 0 Then
        UltimoErro = "FNP0014 - Documento n�o selecionado."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        Exit Function
    End If

    '2. Posiciona controles.
    lNosNumero = pNosNumero
    lNumItem = pNumItem
    Call M46V999.gpExcluiRegistros(lNosNumero, pNumItem, objstcUsuario.PropTabela, "", UltimoErro)
    Call M46V999.gpGraLog(0, "AGE0030 - Item " & lNumItem & " exclu�do com sucesso.")

    gfExcluirItem = True
    
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    
    Exit Function

ControleErro:
    UltimoErro = "[ERR] - gfExcluirItem ERRO - " & Err & " " & Err.Description
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
End Function

Public Function gfRetornaNovoItem(ByRef pStcPed As stcA46V702B, _
                                  ByVal pNovoNumItem As Long, _
                                  pCodDizer As Integer, tagDynaTrace As String) As stcA46V708B

    Dim lobjstcPedLoc As stcA46V708B
    Dim lobjPedLoc As clsA46V708A
    Dim lTipEmissao As e_TipEmissao
    Dim lobj101A As clsA46V101A
    '4. Cria registro na tabela P0042200.

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfRetornaNovoItem"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    Set gfRetornaNovoItem = Nothing

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    Call mpAtribuiClasseModulo(pStcPed, tagDynaTrace)

    Set lobjstcPedLoc = New stcA46V708B
    Set lobjstcPedLoc.M46V999 = M46V999

    '3. Grava registros de cota��o/proposta para renova��o Yasuda.
    'Ajustado para tratar todos os tipos de emiss�o
    lTipEmissao = pStcPed.ColecaoPedLoc(1).Tip_Emissao
'    Select Case pCodDizer
'    Case 0
'        lTipEmissao = e_TipEmissao_SeguroNovo
'    Case Else
'        lTipEmissao = e_TipEmissao_Endosso
'    End Select

    If pCodDizer = 1 Then
        Set lobj101A = New clsA46V101A
        Set lobj101A.M46V999 = M46V999
        pNovoNumItem = lobj101A.mfRetornaProxNumItem(pStcPed)
    End If
    
    If pStcPed.LMI_Unico = 1 Then
        pStcPed.Per_Desc_LMI = 0
        Call mfZerarDescontoLMI(pStcPed)
    End If

    If pCodDizer = 1 Or pCodDizer = 0 Then

        If M46V999.gfGraItem(M46V999.mConexaoSQL, pStcPed.NOSSO_NUMERO, pStcPed.Num_Ped, pNovoNumItem, UltimoErro, _
                             lTipEmissao, lobjstcPedLoc, , pStcPed.Num_Apol_An, , , pStcPed.Cod_Ramo, pStcPed.getProduto.CodProduto, , pStcPed) = False Then
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
    Else

        If M46V001.gfGraEndossoYasuda(pStcPed.Num_Apol_An, pNovoNumItem, pStcPed.NOSSO_NUMERO, _
                                      0, pCodDizer, UltimoErro, False, pStcPed, tagDynaTrace) = False Then
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
        Set lobjPedLoc = New clsA46V708A
        Set lobjPedLoc.M46V999 = M46V999
        Set lobjstcPedLoc = New stcA46V708B
        lobjstcPedLoc.NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
        lobjstcPedLoc.Num_Item = pNovoNumItem
        If Not lobjPedLoc.mfCarregarColecao(lobjstcPedLoc, objstcUsuario.PropTabela, "", True) Then
            UltimoErro = lobjPedLoc.UltimoErro
            GoTo ControleErro
        End If
        If lobjPedLoc.Colecao.Count > 0 Then
            Set lobjstcPedLoc = lobjPedLoc.Colecao.Item(1)
        End If

    End If

    Set lobjstcPedLoc.ObjPai = pStcPed
    pStcPed.ColecaoPedLoc.Add lobjstcPedLoc, lobjstcPedLoc.Class_ID
    
    Set gfRetornaNovoItem = lobjstcPedLoc
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    UltimoErro = "[ERR] - gfRetornaNovoItem ERRO - " & Err & " " & Err.Description
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lobjstcPedLoc = Nothing
    Set lobjPedLoc = Nothing
    Set lobj101A = Nothing
    Return
End Function

Private Function mfVerDados(ByRef pobjstcPedido As stcA46V702B) As Boolean
'Fun��o: verifica se os dados j� est�o calculados.
'        As bases de dados e os recordsets j� ficam abertos.
    Dim lIteNaoDigitado() As Integer    'Itens n�o digitados.
    Dim lOrdem As Byte    'Ordem dos itens.
    Dim lQuaNaoDigitado As Byte     'Quantidade de itens n�o digitados.
    Dim lobjstcItem As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)

    mfVerDados = False

    '1. Verifica se o pr�mio est� calculado,
    If pobjstcPedido.Cod_Sit_Ped <> e_SitPedido_ComCalculo And pobjstcPedido.Cod_Sit_Ped <> e_SitPedido_Transmitido Then
        UltimoErro = "AGE0022 - Pr�mio n�o calculado."
        GoTo SAIDA
    End If

    If pobjstcPedido.Qtd_Parc = 0 And pobjstcPedido.Cod_Reg < 106 Then
        UltimoErro = "FTE0012 - " & M46V999.gPrefixo & "Quantidade de parcelas n�o informado."
        GoTo SAIDA
    End If
    '2. Obt�m registros de itens.
    If pobjstcPedido.ColecaoPedLoc.Count = 0 Then
        UltimoErro = "FTE0012 - " & M46V999.gPrefixo & "N�o localizado registro de item."
        GoTo SAIDA
    End If


    '4. Verifica se tem item n�o digitado.
    lQuaNaoDigitado = 0
    For Each lobjstcItem In pobjstcPedido.ColecaoPedLoc
        If lobjstcItem.Dat_Liber = 0 Then
            lQuaNaoDigitado = lQuaNaoDigitado + 1
            ReDim Preserve lIteNaoDigitado(lQuaNaoDigitado)
            lIteNaoDigitado(lQuaNaoDigitado) = lobjstcItem.Num_Item
        End If
    Next lobjstcItem
    If lQuaNaoDigitado <> 0 Then
        Select Case lQuaNaoDigitado
        Case 1
            UltimoErro = "FGE0019 - " & M46V999.gPrefixo & "Item " & lIteNaoDigitado(1) & " n�o digitado."
        Case Else
            UltimoErro = "FGE0020 - " & M46V999.gPrefixo & "Itens n�o digitados: "
            lOrdem = 0
            Do
                lOrdem = lOrdem + 1
                lQuaNaoDigitado = lQuaNaoDigitado - 1
                Select Case lQuaNaoDigitado
                Case 0
                    UltimoErro = UltimoErro & Format$(lIteNaoDigitado(lOrdem), "0000") & "."
                    Exit Do
                Case 1
                    UltimoErro = UltimoErro & Format$(lIteNaoDigitado(lOrdem), "0000") & " e "
                Case Else
                    UltimoErro = UltimoErro & Format$(lIteNaoDigitado(lOrdem), "0000") & ", "
                End Select
            Loop
        End Select
        GoTo SAIDA
    End If
    '5. Verifica se produto est� zerado
    If pobjstcPedido.Cod_Produto = 0 Then
        UltimoErro = UltimoErro & "Produto n�o est� selecionado corretamente."
        GoTo SAIDA
    End If

    mfVerDados = True

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

FechaObj:
    Set lobjstcItem = Nothing
    Return
End Function

Public Function gfRetornaBoleto(ByVal pNossoNumero As String, _
                                ByVal pReUtilizaBoleto As Boolean, tagDynaTrace As String) As stcA99V009B  'teste ok
    Dim objstcPed As stcA46V702B
    Dim objPedido As clsA46V702A
    Dim lobjstcDadosBoleto As stcA28V720B
    Dim lobjDadosBoleto As clsA28V720A
    Dim lobjstcBoleto As stcA99V009B
    Dim lobjBoleto As clsA99V009A
    Dim lMensagem As String
    Dim lretorno As Byte
    Dim lobjCtrlEmis As clsA46V712A
    Dim lobjstcPedCorr As stcA46V705B

    Dim lKeyDynaTrace As Long
    Dim GerouBoleto As Boolean
    
    '*Luiz Alberto - 08/10/2015
    Dim lDiasVencto     As Integer          'Dias para vencimento
    Dim pMensagem       As String
    Dim IFlagDtVencto   As Integer
    Dim strData         As String
    
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfRetornaBoleto"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    GerouBoleto = False

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    Set objstcPed = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)

    If objstcPed Is Nothing Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Pedido n�o localizado", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    'Verifico o boleto se � o mesmo ou se � outro
    If objstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 And (objstcPed.Num_Cta_Bas_Cor = 0 Or objstcPed.objstcCtrlEmis.Cod_Bco_Boleto > 0) And _
       (objstcPed.objstcCtrlEmis.Num_Bloq = 0 Or Left(objstcPed.objstcCtrlEmis.Num_Bloq, 1) = 5) And _
       (objstcUsuario.TipoUsuario = "2" Or objstcUsuario.TipoUsuario = "3") And _
       objstcPed.Cod_Reg <> 120 And objstcPed.Cod_Reg <> 106 And objstcPed.Cod_Reg <> 107 And objstcPed.Cod_Reg <> 129 Then
        On Error GoTo ControleErro
        lretorno = 0
Retorno:
        Set lobjDadosBoleto = New clsA28V720A
        Set lobjDadosBoleto.M46V999 = M46V999

        Set lobjstcDadosBoleto = New stcA28V720B
        Set lobjstcDadosBoleto.M46V999 = M46V999

        lobjstcDadosBoleto.NOSSO_NUMERO = objstcPed.NOSSO_NUMERO

        If pReUtilizaBoleto = True Then
            If Not lobjDadosBoleto.gfCarregarBoletoDisp(lobjstcDadosBoleto.NOSSO_NUMERO, (lretorno = 1), objstcPed.Nosso_Numero_Original) Then
                UltimoErro = lobjDadosBoleto.UltimoErro
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
                GoTo SAIDA
            End If
        Else
            If Not lobjDadosBoleto.gfCarregarBoletoDisp(lobjstcDadosBoleto.NOSSO_NUMERO, (lretorno = 1)) Then
                UltimoErro = lobjDadosBoleto.UltimoErro
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
                GoTo SAIDA
            End If
            
            GerouBoleto = True
        End If

        If lobjDadosBoleto.Colecao.Count = 0 Then
            'Erro
        Else
            Set lobjstcDadosBoleto = lobjDadosBoleto.Colecao.Item(1)
            If lobjstcDadosBoleto.DAT_DIGIT = 0 Then
                Set objstcPed.objstcBoleto = lobjstcDadosBoleto
                'Ok
            Else
                If pReUtilizaBoleto = True Then
                    'Utilizo o mesmo
                    Set objstcPed.objstcBoleto = lobjstcDadosBoleto
                Else
                    'Gero outro boleto
                    If GerouBoleto = False Then
                    lretorno = 1
                    GoTo Retorno
                End If
            End If
        End If
        End If
        'Salvo o n�mero do boleto na tabela de comum.
        objstcPed.objstcCtrlEmis.Num_Bloq = objstcPed.objstcBoleto.NUM_BOLETO
        objstcPed.objstcCtrlEmis.DIG_BLOQUETO = objstcPed.objstcBoleto.gDigitoBloq

        Set lobjCtrlEmis = New clsA46V712A
        Set lobjCtrlEmis.M46V999 = M46V999

        If Not lobjCtrlEmis.mfAlterar(objstcPed.objstcCtrlEmis, objstcUsuario.PropTabela, "") Then
            lMensagem = lobjCtrlEmis.UltimoErro
            UltimoErro = lMensagem
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If


    If objstcPed.Val_PriParcela > 0 Then
        Set lobjstcBoleto = New stcA99V009B
        Set lobjstcBoleto.M46V999 = M46V999

        Set lobjstcPedCorr = objstcPed.ColecaoPedCorr.Item(1)
        With lobjstcBoleto
            .Corretor = lobjstcPedCorr.gNom_Corretor
            .LocalImagensBarra = App.Path
            .numeroBanco = IIf(objstcPed.objstcCtrlEmis.Cod_Bco_Boleto <> 0, objstcPed.objstcCtrlEmis.Cod_Bco_Boleto, objstcPed.Cod_Bco_Cobr)
            Select Case .numeroBanco
            Case 237, 341, 409, 356, 399, 33     '09029 boleto Santander
                'OK
            Case Else
                .numeroBanco = 409
            End Select
            '.NumeroBanco = 341
            Select Case .numeroBanco
            Case 237
                .LocalImagensBanco = M46V999.gAppPath & "STBS57_237.BMP"
            Case 341
                .LocalImagensBanco = M46V999.gAppPath & "STBS55_341.BMP"
            Case 409
                .LocalImagensBanco = M46V999.gAppPath & "STBS75_409.BMP"
            Case 347
                .LocalImagensBanco = M46V999.gAppPath & "STBS59_347.BMP"
            Case 356
                .LocalImagensBanco = M46V999.gAppPath & "STBS75_356.BMP"
            Case 399
                .LocalImagensBanco = M46V999.gAppPath & "STBS75_399.BMP"
            Case 33
                .LocalImagensBanco = M46V999.gAppPath & "STBS75_033.BMP"
            End Select
            .NomeSacado = objstcPed.Nom_Segurado
            .NossoNumero = objstcPed.objstcCtrlEmis.Num_Bloq
            .DigYasNossoNumero = objstcPed.objstcCtrlEmis.DIG_BLOQUETO
            .NumeroDoc = objstcPed.objstcCtrlEmis.Num_Bloq & "-" & objstcPed.objstcCtrlEmis.DIG_BLOQUETO
            .ramo = objstcPed.gDesc_Ramo
            .valorDocumento = Round(objstcPed.Val_PriParcela, 2)
                
            '********************************************************************************
            '*Luiz Alberto - 08/10/2015                                                     *
            '*Vai procurar na tabela de Ajustes pontuais se tem um vencimento diferenciado  *
            '********************************************************************************
            If M46V103.mfAjustesPontuais(e_Tip_Valor_DiasParaVencimentoBoleto, TpoValor, lDiasVencto, pMensagem, Nothing, 0, tagDynaTrace) = False Then
                Call M46V999.gpGraLog(0, pMensagem)
                Exit Function
            End If
                
            '*Luiz Alberto - 16/10/2015
            IFlagDtVencto = 0
            
            If objstcPed.Dat_Ini_Vig < M46V999.gHoje Then
                If Trim(objstcPed.Dat_Transmissao) = "" Then
                    '*Verifica Corretor (Vila Velha) e Produto
                    If objstcPed.objstcCtrlEmis.Cod_Corr = 18901 And objstcPed.Cod_Produto = 20 Then
                        IFlagDtVencto = 1
                    Else
                        '*Quantidade de dias para o vencimento do boleto.
                        .vencimento = Format(CDate(M46V999.gfForData(2, M46V999.gHoje)) + IIf(lDiasVencto = 0, 7, lDiasVencto), "yyyymmdd")
                    End If
                Else
                    '*Verifica Corretor (Vila Velha) e Produto
                    If objstcPed.objstcCtrlEmis.Cod_Corr = 18901 And objstcPed.Cod_Produto = 20 Then
                        IFlagDtVencto = 2
                    Else
                        '*Quantidade de dias para o vencimento do boleto.
                        .vencimento = Format(CDate(M46V999.gfForData(2, objstcPed.Dat_Transmissao)) + IIf(lDiasVencto = 0, 7, lDiasVencto), "yyyymmdd")
                    End If
                End If
            Else
                '*Verifica Corretor (Vila Velha), produto e se a data de transmiss�o esta preenchida.
                If Trim(objstcPed.Dat_Transmissao) <> "" And objstcPed.objstcCtrlEmis.Cod_Corr = 18901 And objstcPed.Cod_Produto = 20 Then
                    IFlagDtVencto = 3
                Else
                    '*Quantidade de dias para o vencimento do boleto.
                    .vencimento = Format(CDate(M46V999.gfForData(2, objstcPed.Dat_Ini_Vig)) + IIf(lDiasVencto = 0, 7, lDiasVencto), "yyyymmdd")
                End If
                
            End If
                  
            If IFlagDtVencto <> 0 Then
                If IFlagDtVencto = 1 Then
                   strData = M46V999.gHoje
                ElseIf IFlagDtVencto = 2 Then
                   strData = objstcPed.Dat_Transmissao
                ElseIf IFlagDtVencto = 3 Then
                   strData = objstcPed.Dat_Ini_Vig
            End If
    
                Select Case Day(CDate(M46V999.gfForData(2, IIf(Trim(objstcPed.Dat_Transmissao) = "", strData, objstcPed.Dat_Transmissao))))
                    '*Se a transmiss�o da proposta for do dia 1� ao dia 9, a data de vencimento ser� fixada no dia 10 deste mesmo m�s
                    Case Is <= 10
                        .vencimento = Year(CDate(M46V999.gfForData(2, strData))) & Month(CDate(M46V999.gfForData(2, strData))) & "10"
                    '*Se a transmiss�o da proposta for do dia 10 ao dia 20, a data de vencimento ser� a data + 10 dias corridos.
                    Case 11 To 20
                        .vencimento = Format(CDate(M46V999.gfForData(2, strData)) + 10, "yyyymmdd")
                    '*Se a transmiss�o da proposta for do dia 21 at� 30/31, a data de vencimento ser� fixada no dia 10 do m�s seguinte
                    Case Else
                        .vencimento = Year(DateAdd("m", 1, M46V999.gfForData(2, strData))) & Format(Month(DateAdd("m", 1, M46V999.gfForData(2, strData))), "00") & "10"
                End Select
            End If
            
            .Vigencia = M46V999.gfForData(2, objstcPed.Dat_Ini_Vig) & " a " & M46V999.gfForData(2, objstcPed.Dat_Ter_Vig)
            .MensagemPgto = objstcPed.Cod_Unid_Prod & " " & objstcPed.objstcCtrlEmis.Cod_Corr & " PARC 01/01"

            .CarregarDados = True
            Set .ConexaoBD = M46V999.mConexaoSQL
            .NomeTabela = "P0045300"
        End With

        Set lobjBoleto = New clsA99V009A
        If Not lobjBoleto.mfMontarBoleto(lobjstcBoleto) Then
            UltimoErro = lobjBoleto.UltimoErro
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If

        Set gfRetornaBoleto = lobjstcBoleto

        Set lobjstcDadosBoleto = New stcA28V720B
        Set lobjstcDadosBoleto.M46V999 = M46V999

        With lobjstcDadosBoleto
            .COD_BANCO = lobjstcBoleto.numeroBanco
            '.NUM_BOLETO = lobjstcBoleto.NossoNumero
            .NUM_BOLETO = objstcPed.objstcBoleto.NUM_BOLETO
            .NOSSO_NUMERO = objstcPed.NOSSO_NUMERO
            .COD_UNID = objstcPed.Cod_Unid_Prod
            .Cod_Ramo = objstcPed.Cod_Ramo
            .Nom_Segurado = Left(objstcPed.Nom_Segurado, 50)
            .VAL_PR = CSng(lobjstcBoleto.valorDocumento)
            .DAT_VENC = lobjstcBoleto.vencimento
            .Cod_Corr = objstcPed.objstcCtrlEmis.Cod_Corr
            .NUM_LINHA_DIGIT = Left(lobjstcBoleto.linhaDigitavel, 60)
            .Cod_User = M46V999.gNomUsuario6
            If lobjstcDadosBoleto.DAT_DIGIT = 0 Then
                .DAT_DIGIT = M46V999.gHoje
            Else
                .DAT_DIGIT = lobjstcDadosBoleto.DAT_DIGIT
            End If
            .HOR_DIGIT = Format(Time, "hhmmss")
            .STATUS = 2

        End With
        Set lobjDadosBoleto = New clsA28V720A
        Set lobjDadosBoleto.M46V999 = M46V999

        If Not lobjDadosBoleto.gfAlterar(lobjstcDadosBoleto, objstcUsuario.PropTabela) Then
            UltimoErro = lobjDadosBoleto.UltimoErro
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
    Else
        Set lobjstcBoleto = New stcA99V009B
        Set gfRetornaBoleto = lobjstcBoleto
    End If
    Else
            Set lobjstcBoleto = New stcA99V009B
            Set gfRetornaBoleto = lobjstcBoleto
    End If

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    If Not M46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set objstcPed = Nothing
    Set objPedido = Nothing
    Set lobjstcDadosBoleto = Nothing
    Set lobjDadosBoleto = Nothing
    Set lobjstcBoleto = Nothing
    Set lobjBoleto = Nothing
    Set lobjCtrlEmis = Nothing
    Set lobjstcPedCorr = Nothing
    Return
End Function

Private Function gfAtualizaTabCtrlPagAntec(objBoleto As stcA99V009B) As Boolean
'Column_name Type    Computed    Length  Prec    Scale
'NUM_BOLETO     numeric no  5   6       0
'NUM_PROP       varchar no  20
'NUM_PI         numeric no  9   10      0
'NUM_APOL       numeric no  9   10      0
'NUM_ENDO       numeric no  5   6       0
'COD_UNID       numeric no  5   4       0
'COD_RAMO       numeric no  5   3       0
'COD_BANCO      numeric no  5   4       0
'NOM_SEGURADO   varchar no  100
'VAL_PR         numeric no  9   15      2
'DAT_VENC       numeric no  5   8       0
'COD_CORR       numeric no  5   7       0
'NUM_LINHA_DIGIT varchar no 60
'COD_USER       varchar no  20
'DAT_DIGIT      numeric no  5   8       0
'HOR_DIGIT      numeric no  5   6       0
'DAT_COMPT      numeric no  5   8       0
'DAT_EMIS_APOL  numeric no  5   8       0
'TIPO_MOEDA     varchar no  10

    ' Ver se o boleto j� est� cadastrado
    
    ' Se n�o estiver, incluir
    
    ' Se estiver, atualizar
End Function

Public Function gfRetornaBoletoWS(ByVal pNossoNumero As String, _
    ByVal pReUtilizaBoleto As Boolean, tagDynaTrace As String, Num_PI_Proposta As String) As stcA99V009B  'teste ok
    Dim objstcPed As stcA46V702B
    Dim objPedido As clsA46V702A
    Dim lobjstcDadosBoleto As stcA28V720B
    Dim lobjDadosBoleto As clsA28V720A
    Dim lobjstcBoleto As stcA99V009B
    Dim lobjBoleto As clsA99V009A
    Dim lMensagem As String
    Dim lretorno As Byte
    Dim lobjCtrlEmis As clsA46V712A
    Dim lobjstcPedCorr As stcA46V705B
    
    Dim lKeyDynaTrace As Long
    Dim GerouBoleto As Boolean
    
'*Luiz Alberto - 08/10/2015
    Dim lDiasVencto     As Integer          'Dias para vencimento
    Dim pMensagem       As String
    Dim IFlagDtVencto   As Integer
    Dim strData         As String
    Dim WSClient        As Object
    Dim Retorno As String
    Dim arrRetorno
    Dim rst As clsYasRecordSet
    
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfRetornaBoletoWS"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro
    
    'Set WSClient = CreateObject("MSSOAP.SoapClient30")
    'gfLog_Emissao pNossoNumero, Num_PI_Proposta, "gfRetornaBoletoWS (" & M46V999.gUrlWebServiceBoleto & ") - Inicio", Enum_Tipo_LogEmissao_Inicio
    
    GerouBoleto = False
    
    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        gfLog_Emissao pNossoNumero, Num_PI_Proposta, "If Not mfPosicionaUsuario(tagDynaTrace) Then", Enum_Tipo_LogEmissao_Abortado
        GoTo SAIDA
    End If
    
    Set objstcPed = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)
    
    If objstcPed Is Nothing Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Pedido n�o localizado", eSeveridadeErro_Medio, Class_Name, methodName)
        gfLog_Emissao pNossoNumero, Num_PI_Proposta, "If objstcPed Is Nothing Then", Enum_Tipo_LogEmissao_Abortado
        GoTo SAIDA
    End If
    
    If objstcPed.Val_Pr_Antec = 0 Then
        Set lobjstcBoleto = New stcA99V009B
        Set gfRetornaBoletoWS = lobjstcBoleto
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
        gfLog_Emissao pNossoNumero, Num_PI_Proposta, "If objstcPed.Val_Pr_Antec = 0 Then", Enum_Tipo_LogEmissao_Abortado
        GoTo SAIDA
    End If
    
    ' Se for proposta recusada, n�o mostra boleto
    If objstcPed.objstcCtrlEmis.Cod_Iden_Stat_Acao = 80 Then
        Set lobjstcBoleto = New stcA99V009B
        Set gfRetornaBoletoWS = lobjstcBoleto
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
        gfLog_Emissao pNossoNumero, Num_PI_Proposta, "If objstcPed.objstcCtrlEmis.Cod_Iden_Stat_Acao = 80 Then", Enum_Tipo_LogEmissao_Abortado
        GoTo SAIDA
    End If
    
'    '2018-07-11: Juliano - removido
'    ' Se j� tem n�mero de boleto, n�o gerar outro boleto. Mostrar o que j� tem
'    If objstcPed.objstcCtrlEmis.Num_Bloq <> 0 Then
'        Set lobjstcBoleto = New stcA99V009B
'        Set lobjstcBoleto = gfRetornaBoleto(pNossoNumero, pReUtilizaBoleto, tagDynaTrace) ' Rotina antiga
'
'        If lobjstcBoleto.linhaDigitavel <> "" Then ' Se encontrou boleto velho, usar
'            Set gfRetornaBoletoWS = lobjstcBoleto
'            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
'            gfLog_Emissao pNossoNumero, Num_PI_Proposta, "If lobjstcBoleto.linhaDigitavel <> "" Then", Enum_Tipo_LogEmissao_Abortado
'            GoTo SAIDA
'        End If
'    End If
'    '2018-07-11: Fim

    
    If objstcPed.Val_PriParcela > 0 Then
        Set lobjstcBoleto = New stcA99V009B
        Set lobjstcBoleto.M46V999 = M46V999
        
        Set lobjstcPedCorr = objstcPed.ColecaoPedCorr.Item(1)
        With lobjstcBoleto
            .Corretor = lobjstcPedCorr.gNom_Corretor
            .LocalImagensBarra = App.Path
            .numeroBanco = IIf(objstcPed.objstcCtrlEmis.Cod_Bco_Boleto <> 0, objstcPed.objstcCtrlEmis.Cod_Bco_Boleto, objstcPed.Cod_Bco_Cobr)
            Select Case .numeroBanco
            Case 237, 341, 409, 356, 399, 33     '09029 boleto Santander
'OK
            Case Else
                .numeroBanco = 409
            End Select
'.NumeroBanco = 341
            Select Case .numeroBanco
            Case 237
                .LocalImagensBanco = M46V999.gAppPath & "STBS57_237.BMP"
            Case 341
                .LocalImagensBanco = M46V999.gAppPath & "STBS55_341.BMP"
            Case 409
                .LocalImagensBanco = M46V999.gAppPath & "STBS75_409.BMP"
            Case 347
                .LocalImagensBanco = M46V999.gAppPath & "STBS59_347.BMP"
            Case 356
                .LocalImagensBanco = M46V999.gAppPath & "STBS75_356.BMP"
            Case 399
                .LocalImagensBanco = M46V999.gAppPath & "STBS75_399.BMP"
            Case 33
                .LocalImagensBanco = M46V999.gAppPath & "STBS75_033.BMP"
            End Select
            .NomeSacado = objstcPed.Nom_Segurado
            .NossoNumero = objstcPed.objstcCtrlEmis.Num_Bloq
            .DigYasNossoNumero = objstcPed.objstcCtrlEmis.DIG_BLOQUETO
            .NumeroDoc = objstcPed.objstcCtrlEmis.Num_Bloq & "-" & objstcPed.objstcCtrlEmis.DIG_BLOQUETO
            .ramo = objstcPed.gDesc_Ramo
            .valorDocumento = Round(objstcPed.Val_PriParcela, 2)
            
'********************************************************************************
'*Luiz Alberto - 08/10/2015                                                     *
'*Vai procurar na tabela de Ajustes pontuais se tem um vencimento diferenciado  *
'********************************************************************************
            If M46V103.mfAjustesPontuais(e_Tip_Valor_DiasParaVencimentoBoleto, TpoValor, lDiasVencto, pMensagem, Nothing, 0, tagDynaTrace) = False Then
                Call M46V999.gpGraLog(0, pMensagem)
                Exit Function
            End If
            
'*Luiz Alberto - 16/10/2015
            IFlagDtVencto = 0
            
            If objstcPed.Dat_Ini_Vig < M46V999.gHoje Then
                If Trim(objstcPed.Dat_Transmissao) = "" Then
'*Verifica Corretor (Vila Velha) e Produto
                    If objstcPed.objstcCtrlEmis.Cod_Corr = 18901 And objstcPed.Cod_Produto = 20 Then
                        IFlagDtVencto = 1
                    Else
'*Quantidade de dias para o vencimento do boleto.
                        .vencimento = Format(CDate(M46V999.gfForData(2, M46V999.gHoje)) + IIf(lDiasVencto = 0, 7, lDiasVencto), "yyyymmdd")
                    End If
                Else
'*Verifica Corretor (Vila Velha) e Produto
                    If objstcPed.objstcCtrlEmis.Cod_Corr = 18901 And objstcPed.Cod_Produto = 20 Then
                        IFlagDtVencto = 2
                    Else
'*Quantidade de dias para o vencimento do boleto.
                        .vencimento = Format(CDate(M46V999.gfForData(2, objstcPed.Dat_Transmissao)) + IIf(lDiasVencto = 0, 7, lDiasVencto), "yyyymmdd")
                    End If
                End If
            Else
'*Verifica Corretor (Vila Velha), produto e se a data de transmiss�o esta preenchida.
                If Trim(objstcPed.Dat_Transmissao) <> "" And objstcPed.objstcCtrlEmis.Cod_Corr = 18901 And objstcPed.Cod_Produto = 20 Then
                    IFlagDtVencto = 3
                Else
'*Quantidade de dias para o vencimento do boleto.
                    .vencimento = Format(CDate(M46V999.gfForData(2, objstcPed.Dat_Ini_Vig)) + IIf(lDiasVencto = 0, 7, lDiasVencto), "yyyymmdd")
                End If
                
            End If
            
            If IFlagDtVencto <> 0 Then
                If IFlagDtVencto = 1 Then
                    strData = M46V999.gHoje
                ElseIf IFlagDtVencto = 2 Then
                    strData = objstcPed.Dat_Transmissao
                ElseIf IFlagDtVencto = 3 Then
                    strData = objstcPed.Dat_Ini_Vig
                End If
                
                Select Case Day(CDate(M46V999.gfForData(2, IIf(Trim(objstcPed.Dat_Transmissao) = "", strData, objstcPed.Dat_Transmissao))))
'*Se a transmiss�o da proposta for do dia 1� ao dia 9, a data de vencimento ser� fixada no dia 10 deste mesmo m�s
                Case Is <= 10
                    .vencimento = Year(CDate(M46V999.gfForData(2, strData))) & Month(CDate(M46V999.gfForData(2, strData))) & "10"
'*Se a transmiss�o da proposta for do dia 10 ao dia 20, a data de vencimento ser� a data + 10 dias corridos.
                Case 11 To 20
                    .vencimento = Format(CDate(M46V999.gfForData(2, strData)) + 10, "yyyymmdd")
'*Se a transmiss�o da proposta for do dia 21 at� 30/31, a data de vencimento ser� fixada no dia 10 do m�s seguinte
                Case Else
                    .vencimento = Year(DateAdd("m", 1, M46V999.gfForData(2, strData))) & Format(Month(DateAdd("m", 1, M46V999.gfForData(2, strData))), "00") & "10"
                End Select
            End If
            
            .Vigencia = M46V999.gfForData(2, objstcPed.Dat_Ini_Vig) & " a " & M46V999.gfForData(2, objstcPed.Dat_Ter_Vig)
            .MensagemPgto = objstcPed.Cod_Unid_Prod & " " & objstcPed.objstcCtrlEmis.Cod_Corr & " PARC 01/01"
            
            .CarregarDados = True
            Set .ConexaoBD = M46V999.mConexaoSQL
            .NomeTabela = "P0045300"
        End With
        
        'Set gfRetornaBoletoWS = lobjstcBoleto
        
        Set lobjstcDadosBoleto = New stcA28V720B
        Set lobjstcDadosBoleto.M46V999 = M46V999
        
        With lobjstcDadosBoleto
            .COD_BANCO = lobjstcBoleto.numeroBanco
            '.NUM_BOLETO = lobjstcBoleto.NossoNumero
            .NUM_BOLETO = objstcPed.objstcBoleto.NUM_BOLETO
            .NOSSO_NUMERO = objstcPed.NOSSO_NUMERO
            .COD_UNID = objstcPed.Cod_Unid_Prod
            .Cod_Ramo = objstcPed.Cod_Ramo
            .Nom_Segurado = Left(objstcPed.Nom_Segurado, 50)
            .VAL_PR = CSng(lobjstcBoleto.valorDocumento)
            .DAT_VENC = lobjstcBoleto.vencimento
            .Cod_Corr = objstcPed.objstcCtrlEmis.Cod_Corr
            .NUM_LINHA_DIGIT = Left(lobjstcBoleto.linhaDigitavel, 60)
            .Cod_User = M46V999.gNomUsuario6
            If lobjstcDadosBoleto.DAT_DIGIT = 0 Then
                .DAT_DIGIT = M46V999.gHoje
            Else
                .DAT_DIGIT = lobjstcDadosBoleto.DAT_DIGIT
            End If
            .HOR_DIGIT = Format(Time, "hhmmss")
            .STATUS = 2
            
        End With
        Set lobjDadosBoleto = New clsA28V720A
        Set lobjDadosBoleto.M46V999 = M46V999
        
        ' Chamar WS de Boleto no lugar da rotina local de Boleto
        Set WSClient = CreateObject("MSSOAP.SoapClient30")
        WSClient.ClientProperty("ServerHTTPRequest") = True
        Call WSClient.MSSoapInit(M46V999.gUrlWebServiceBoleto)

        
        Retorno = WSClient.RetornaBoletoSAP_VB( _
            lobjstcDadosBoleto.NOSSO_NUMERO, _
            lobjstcDadosBoleto.DAT_VENC, _
            lobjstcDadosBoleto.VAL_PR, _
            lobjstcDadosBoleto.Cod_Ramo, _
            lobjstcDadosBoleto.Nom_Segurado, _
            Num_PI_Proposta, _
            objstcPed.Num_Endo, _
            objstcPed.Nom_Logr & " " & objstcPed.Dsc_Compl_Logr, _
            objstcPed.Nom_Bairro, _
            objstcPed.Nom_Cid, _
            objstcPed.Sig_UF, _
            objstcPed.Num_CEP, _
            objstcPed.Num_Cgc_Cpf)
        arrRetorno = Split(Retorno, "###")
        Call M46V999.gpGraLog(0, "Retorno do WS de Boleto: " & Retorno)
        
        If arrRetorno(0) = "0" Or M46V999.gfPreenchido(arrRetorno(7)) Then
            Call M46V999.gpGraLog(1, "Erro ao gerar boleto: " & arrRetorno(7))
            UltimoErro = "Erro ao gerar boleto: " & arrRetorno(7)
            Set gfRetornaBoletoWS = Nothing
            GoTo ControleErro
        End If
        
        'Salvo o n�mero do boleto na tabela de comum.
        objstcPed.objstcCtrlEmis.Num_Bloq = arrRetorno(0)
        objstcPed.objstcCtrlEmis.DIG_BLOQUETO = arrRetorno(6)
    
        Set lobjCtrlEmis = New clsA46V712A
        Set lobjCtrlEmis.M46V999 = M46V999
    
        If Not lobjCtrlEmis.mfAlterar(objstcPed.objstcCtrlEmis, objstcUsuario.PropTabela, "") Then
            lMensagem = lobjCtrlEmis.UltimoErro
            UltimoErro = lMensagem
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
        
        ' Monta objeto boleto
        lobjstcBoleto.DigitoVerificador = arrRetorno(6)
        lobjstcBoleto.CodigoBarra = arrRetorno(1)
        lobjstcBoleto.LinhaWN = arrRetorno(5)
        lobjstcBoleto.LinhaWN_Fonte = arrRetorno(4)
        lobjstcBoleto.CodigoBarraBinario = arrRetorno(2)
        lobjstcBoleto.linhaDigitavel = arrRetorno(3)
        
        objstcPed.objstcBoleto.NUM_BOLETO = arrRetorno(0)
        objstcPed.objstcBoleto.NUM_LINHA_DIGIT = arrRetorno(3)
        
        ' Salvar / Atualizar o Boleto na Tab_Ctrl_Pagantec no SYAS_EMIS.
        lobjstcDadosBoleto.NUM_BOLETO = arrRetorno(0)
        lobjstcDadosBoleto.DigBoleto = arrRetorno(6)
        lobjstcDadosBoleto.NUM_LINHA_DIGIT = arrRetorno(3)
        lobjstcDadosBoleto.CPF_CNPJ_Segurado = objstcPed.Num_Cgc_Cpf
        
        Call lobjDadosBoleto.AtualizaBoletoRegistrado(lobjstcDadosBoleto)
    Else
        Set lobjstcBoleto = New stcA99V009B
    End If
    
    Set gfRetornaBoletoWS = lobjstcBoleto

Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
Err.Clear: On Error GoTo 0
GoSub FechaObj
Exit Function

ControleErro:
If Not M46V999.gfPreenchido(UltimoErro) Then
    UltimoErro = Err.Description
End If
Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
GoTo SAIDA

FechaObj:
Set objstcPed = Nothing
Set objPedido = Nothing
Set lobjstcDadosBoleto = Nothing
Set lobjDadosBoleto = Nothing
Set lobjstcBoleto = Nothing
Set lobjBoleto = Nothing
Set lobjCtrlEmis = Nothing
Set lobjstcPedCorr = Nothing
Return
End Function

Public Function gfTransmitirProposta(ByVal pNossoNumero As String, _
                                     ByVal pReUtilizaBoleto As Boolean, tagDynaTrace As String) As String
          Dim lobjstcPedido As stcA46V702B
          Dim lobjstcBoleto As stcA99V009B
          Dim lConexaoEAPDB As ADODB.Connection
          Dim lCnSyasPortal2008 As ADODB.Connection
          Dim lConexaoDiv As ADODB.Connection
          Dim lConexaoSyasEmis As ADODB.Connection
          Dim lAlertas As String
          Dim lRetornoEhErro As Boolean
          Dim lSelect As String
          Dim lrst As clsYasRecordSet
          Dim lobjstcPedLoc As stcA46V708B
          Dim lUpdate As String
          Dim lNum_pi As Double
          Dim lInsert As String

          Dim lKeyDynaTrace As Long
10        M46V999.tagDynaTrace = tagDynaTrace
          Dim methodName As String
20        methodName = "gfTransmitirProposta"
30        lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
          
40        Err.Clear: UltimoErro = ""

50        On Error GoTo ControleErro

60        If M46V999.gTipoTransmDLL = ETransmDLL_OFFLine Then
70            gfTransmitirProposta = mfTransmitirPropostaOFFLine(pNossoNumero, pReUtilizaBoleto, tagDynaTrace)
80            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Transmitir offline", eSeveridadeErro_Medio, Class_Name, methodName)
90            GoTo SAIDA
100       End If

          '- Abre as bases que ser�o utilizadas
110       If Not M46V999.gfAbrBasDados("EAPDB", M46V999.gTipEAPDB, M46V999.gEndEAPDB, lConexaoEAPDB, UltimoErro) Then
120           GoTo ControleErro
130       End If

          '- Abre as bases que ser�o utilizadas
140       If Not M46V999.gfAbrBasDados("SYASPORTAL", M46V999.gTipSyasPortal, M46V999.gEndSyasPortal, lCnSyasPortal2008, UltimoErro) Then
150           GoTo ControleErro
160       End If

          '- Abre as bases que ser�o utilizadas
170       If Not M46V999.gfAbrBasDados("RamosDiversos", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, lConexaoDiv, UltimoErro) Then
180           GoTo ControleErro
190       End If

          '- Abre as bases que ser�o utilizadas
          If Not M46V999.gfAbrBasDados("Syas_Emis", M46V999.gTipSYASEMIS, M46V999.gEndSYASEMIS, lConexaoSyasEmis, UltimoErro) Then
              GoTo ControleErro
          End If

200       Set lobjstcPedido = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)

210       If mfVerDados(lobjstcPedido) = False Then
220           Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "mfVerDados veio false", eSeveridadeErro_Medio, Class_Name, methodName)
230           GoTo SAIDA
240       End If
          
          '11.46 Item 2.Nova regra para vig�ncia futura
          '11.48 Item 8 Nova regra para vig�ncia futura - Seguro Novo / Renova��o Cong�nere
          Dim lDat_Ini_Vig_MAX As Date

250       If lobjstcPedido.Cod_Produto >= 1144 And _
             lobjstcPedido.Cod_Ramo = 113 And (lobjstcPedido.Cod_Estip = 10 Or lobjstcPedido.Cod_Estip = 30) Then
               
      '        If lobjstcPedido.ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_SeguroNovo Or _
      '            lobjstcPedido.ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_RenCongenere Then
      '            lDat_Ini_Vig_MAX = DateAdd("d", 4, Date)
      '        Else
      '            lDat_Ini_Vig_MAX = DateAdd("d", 15, Date)
      '        End If

260           lDat_Ini_Vig_MAX = DateAdd("d", 30, Date)
              
270           If lDat_Ini_Vig_MAX < CDate(M46V999.gfForData(2, lobjstcPedido.Dat_Ini_Vig)) Then
280               UltimoErro = "FGE0150 - Inicio de Vig�ncia inv�lida (superior a " & Format(lDat_Ini_Vig_MAX, "dd/MM/YYYY") & ")" & UltimoErro
290               GoTo ControleErro
300           End If
310       End If

          'Cartao Credito
          If lobjstcPedido.Cod_Bco_Cobr = 999 And lobjstcPedido.Cod_Agen_Bas_Cor = 9999 And lobjstcPedido.Num_Cta_Bas_Cor = 9999999 Then
            lSelect = ""
            lSelect = "SELECT * FROM Tab_Controle_Cartao WHERE NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO) & ""
                 
            If Not M46V999.gfObtRegistro(lConexaoDiv, lSelect, lrst, UltimoErro) Then
                GoTo ControleErro
            End If
            
            lInsert = ""
            
            If Not lrst.EOF Then
                lInsert = "Insert into SYAS_EMIS..TAB_DADOS_CARTAO VALUES ( "
                lInsert = lInsert & M46V999.gfFormatarCamposSQL(lrst("Num_PI")) & ","
                lInsert = lInsert & M46V999.gfFormatarCamposSQL(lobjstcPedido.Cod_Ramo) & ","
                lInsert = lInsert & M46V999.gfFormatarCamposSQL(lrst("TransactionId")) & ","
                lInsert = lInsert & M46V999.gfFormatarCamposSQL(lrst("AuthCode")) & ","
                lInsert = lInsert & "FORMAT(GETDATE(), 'yyyy-MM-dd'),"
                lInsert = lInsert & M46V999.gfFormatarCamposSQL(lrst("TransactionDate")) & ","
                lInsert = lInsert & M46V999.gfFormatarCamposSQL(lrst("TransactionStatus")) & ","
                lInsert = lInsert & M46V999.gfFormatarCamposSQL(lrst("TransactionType")) & ","
                lInsert = lInsert & M46V999.gfFormatarCamposSQL(lrst("OrderNumber")) & ","
                lInsert = lInsert & M46V999.gfFormatarCamposSQL(lrst("OrderTotal")) & ","
                lInsert = lInsert & M46V999.gfFormatarCamposSQL(lrst("Error")) & ","
                lInsert = lInsert & "CONVERT(XML, '" & lrst("xml_send") & "'),"
                lInsert = lInsert & "CONVERT(XML, '" & lrst("xml_receive") & "'),"
                lInsert = lInsert & M46V999.gfFormatarCamposSQL(lrst("num_tempo_total")) & ","
                lInsert = lInsert & "GETDATE() )"
                    
                If M46V999.gfExeSQL(lConexaoSyasEmis, lInsert, UltimoErro) <> 0 Then
                    UltimoErro = " Erro ao salvar dados da transa��o do cart�o de cr�dito na TAB_DADOS_CARTAO do SYAS_EMIS. Processo de transmiss�o interrompido. Favor tente novamente. Se o problema persistir, entre em contato com o administrador do sistema. " & UltimoErro
                    GoTo ControleErro
                End If
             End If
          End If
           
ExecutaTransmissao:
320       lSelect = "SELECT Num_PI FROM dbo.Tab_Vinculo_Proposta_PI "
330       lSelect = lSelect & "WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO)
340       If Not M46V999.gfObtRegistro(lConexaoEAPDB, lSelect, lrst, UltimoErro) Then
350           GoTo ControleErro
360       End If
          
370       If lrst.EOF Then
              'Busco o n�mero do PI (agora antes de gerar o boleto)
380           lSelect = " exec [dbo].proc_GetProxPI 2"
390           If Not M46V999.gfObtRegistro(lConexaoEAPDB, lSelect, lrst, UltimoErro) Then
400               UltimoErro = " Erro ao vincular com PI " & UltimoErro
410               GoTo ControleErro
420           End If
430           If Not lrst.EOF Then
440               lNum_pi = lrst("Num_PI")
                  
450               lSelect = " Insert into dbo.Tab_Vinculo_Proposta_PI (Nosso_Numero, Num_PI, Data , Hora, Usuario) values ( "
460               lSelect = lSelect & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO) & ","
470               lSelect = lSelect & M46V999.gfFormatarCamposSQL(lNum_pi) & ","
480               lSelect = lSelect & " convert(varchar,getdate(),112) ,"
490               lSelect = lSelect & " replace(left(convert(varchar,getdate(),114),8),':','') ,"
500               lSelect = lSelect & M46V999.gfFormatarCamposSQL(M46V999.gNomUsuario6) & ")"
                  
510               If M46V999.gfExeSQL(lConexaoEAPDB, lSelect, UltimoErro) <> 0 Then
520                   UltimoErro = " Erro ao salvar vinculo do PI. Processo de transmiss�o interrompido. " & UltimoErro
530                   GoTo ControleErro
540               End If
550           Else
560               Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Procedure proc_GetProxPI n�o retornou registros. Processo de transmiss�o interrompido.", 0, False)
570               UltimoErro = " Erro ao vincular com PI - Procedure proc_GetProxPI n�o retornou registros. Processo de transmiss�o interrompido."
580               GoTo ControleErro
590           End If
              
600       Else
610           lNum_pi = lrst("Num_PI")
620       End If
          
          '11.42
630       gfLog_Emissao pNossoNumero, lNum_pi, "Inicio", Enum_Tipo_LogEmissao_Inicio

          'CR28705 - REENTRADA COM DEBITO
          'se for reentrada com pagamento antecipado e forma de pagamento debito, retirar o boleto gerado na anterior.
640       If lobjstcPedido.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 And lobjstcPedido.objstcCtrlEmis.Num_Bloq > 0 And _
             lobjstcPedido.ehReentrada And lobjstcPedido.gTipPagamento = e_FormaPagamento_Debito Then
              'ATUALIZA O CAMPO RESPONSAVEL PELO BOLETO POIS A FORMA DE PAGAMENTO � DEBITO (N�O TEM BOLETO)
650           lSelect = " UPDATE tab_ctrl_emis SET num_bloq = 0 WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO)

660           If M46V999.gfExeSQL(lCnSyasPortal2008, lSelect, UltimoErro) <> 0 Then
670               UltimoErro = "Erro ao retirar o vinculo com o boleto quando pagamento for d�bito em conta - tabela tab_ctrl_emis"
680               GoTo ControleErro
690           End If
              
700           lobjstcPedido.objstcCtrlEmis.Num_Bloq = 0
710       End If

          'Busco o n�mero do PI
720       lRetornoEhErro = True

730       If Not M46V999.gfTransmissaoInterna(1, lobjstcPedido, UltimoErro, lCnSyasPortal2008, lRetornoEhErro) Then
740           If lRetornoEhErro = True Then
750               GoTo ControleErro
760           Else
770               GoTo RetornaProtocolo
780           End If
790       End If

800       If M46V999.gfExeSQL(lConexaoDiv, "Update ramosdiversos.dbo.ftp_tab_ped set tip_transmissao = 1 where nosso_numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO) & " and cod_unicidade =  " & lobjstcPedido.Cod_Unicidade, UltimoErro) <> 0 Then
810           GoTo ControleErro
820           End If
          
          '#BOLETO#
830       MensagemAlerta = lAlertas
840       If lobjstcPedido.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 And lobjstcPedido.objstcCtrlEmis.Num_Bloq = 0 And lobjstcPedido.Cod_Reg <> 120 And _
              lobjstcPedido.Cod_Reg <> 106 And lobjstcPedido.Cod_Reg <> 107 And lobjstcPedido.Cod_Reg <> 129 Then
              
              'Set lobjstcBoleto = gfRetornaBoleto(lobjstcPedido.Nosso_Numero, pReUtilizaBoleto, tagDynaTrace)
850           Set lobjstcBoleto = gfRetornaBoletoWS(lobjstcPedido.NOSSO_NUMERO, pReUtilizaBoleto, tagDynaTrace, CStr(lNum_pi))

860           If Not lobjstcBoleto Is Nothing Then
                  'Recarrega o Pedido, Agora com o Boleto gerado
870               Set lobjstcPedido = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)
880           Else
890               UltimoErro = "Falha ao carregar boleto:" & UltimoErro
900               GoTo ControleErro
910           End If
              
920           If lobjstcPedido.Val_Pr_Antec > 0 Then
930               If lobjstcPedido.objstcCtrlEmis.Num_Bloq = 0 Then
940                   UltimoErro = "Erro ao gerar o boleto! Favor tentar a transmiss�o novamente"
950                   GoTo ControleErro
960               End If

      '            'Verifica se o valor do boleto � diferente do valor da primeira parcela
      '            If lobjstcBoleto.valorDocumento <> Round(lobjstcPedido.Val_PriParcela, 2) Then
      '                UltimoErro = "O valor do boleto emitido diferente do valor calculado! Ser� necess�rio realizar uma copia da proposta."
      '                GoTo ControleErro
      '            End If
970           End If
              
              'Salva Numero e Digito do Boleto na FTP
980           lUpdate = " Update FTP_Tab_Ctrl_Emis set Num_Bloq = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.objstcCtrlEmis.Num_Bloq)
990           lUpdate = lUpdate & ", DIG_BLOQUETO = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.objstcCtrlEmis.DIG_BLOQUETO)
1000          lUpdate = lUpdate & " where nosso_numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO)
1010          If M46V999.gfExeSQL(lConexaoDiv, lUpdate, UltimoErro) <> 0 Then
1020              GoTo ControleErro
1030          End If
              
1040      End If
          '#BOLETO - FIM#
          
          'Atualiza o valor antecipado antes da execu��o da startfluxo.
1050      If lobjstcPedido.objstcCtrlEmis.Num_Bloq <> 0 Then
1060          lUpdate = " Update SyasPortal.dbo.Tab_Ped set val_pr_antec = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.Val_Pr_Antec)
1070          lUpdate = lUpdate & " where nosso_numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO)
1080          If M46V999.gfExeSQL(lCnSyasPortal2008, lUpdate, UltimoErro) <> 0 Then
1090              GoTo ControleErro
1100          End If
              
1110          lUpdate = " Update FTP_Tab_Ped set val_pr_antec = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.Val_Pr_Antec)
1120          lUpdate = lUpdate & " where nosso_numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO)
1130          If M46V999.gfExeSQL(lConexaoDiv, lUpdate, UltimoErro) <> 0 Then
1140              GoTo ControleErro
1150          End If
              
1160      Else
1170          lUpdate = " Update SyasPortal.dbo.Tab_Ped set val_pr_antec = " & 0
1180          lUpdate = lUpdate & " where nosso_numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO)
1190          If M46V999.gfExeSQL(lCnSyasPortal2008, lUpdate, UltimoErro) <> 0 Then
1200              GoTo ControleErro
1210          End If
              
1220          lUpdate = " Update FTP_Tab_Ped set val_pr_antec = " & 0
1230          lUpdate = lUpdate & " where nosso_numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO)
1240          If M46V999.gfExeSQL(lConexaoDiv, lUpdate, UltimoErro) <> 0 Then
1250              GoTo ControleErro
1260          End If
1270      End If

          '================ Inicia o Fluxo ================
1280      lSelect = " exec ramosdiversos.dbo.proc_StartFluxoEmissaoTransmissaoMulti " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO) & "," & lNum_pi
          'chamada trocada pois n�o estava lendo o retorno de erro da proc.
      '    If M46V999.gfExeSQL(lConexaoDiv, lSelect, UltimoErro) <> 0 Then
      '        GoTo ControleErro
      '    End If
1290      If Not M46V999.gfObtRegistro(lConexaoDiv, lSelect, lrst, UltimoErro) Then
1300          GoTo ControleErro
1310      End If
          
1320      If lrst!Protocolo = -1 Then
1330          UltimoErro = "Transmiss�o n�o executada: Par�metros: Nosso_numero - " & lobjstcPedido.NOSSO_NUMERO & ", Num_PI - " & lNum_pi & " - " & lrst!msgerro
1340          GoTo ControleErro
1350      End If
          '================ ================ ================
1360      lobjstcPedido.Num_PI_Emissao = lNum_pi
          
          '2. Inicia processo.
          '4. Obt�m registros de protocolo referente ao nosso n�mero.
          'a) Abre base de dados.
          Dim lobjTrans As clsA46V726A

          Dim lobjTransm As clsA46V718A    'Objeto da classe clsA46V718A (TAB_PED_TRANSM)
          Dim lobjstcTransm As stcA46V718B    'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM)

1370      Set lobjTransm = New clsA46V718A            'Objeto da classe clsA46V718A (TAB_PED_TRANSM)
1380      Set lobjTransm.M46V999 = M46V999

1390      Set lobjstcTransm = New stcA46V718B         'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM)

1400      Set lobjstcPedLoc = lobjstcPedido.ColecaoPedLoc.Item(1)

1410      lobjstcTransm.NOSSO_NUMERO = lobjstcPedido.NOSSO_NUMERO
1420      If Not lobjTransm.mfCarregarColecao(lobjstcTransm, objstcUsuario.PropTabela, "") Then
1430          UltimoErro = lobjTransm.UltimoErro
1440          GoTo ControleErro
1450      End If
          
1460      If lobjTransm.Colecao.Count = 0 Then
1470          lobjstcTransm.Cod_Ramo = lobjstcPedido.Cod_Ramo
1480          lobjstcTransm.Cod_Tip_Logr_Risco = lobjstcPedLoc.Cod_Tip_Logr_Risco
1490          lobjstcTransm.Dat_Hor_Envio = M46V999.gfForData(3)
1500          lobjstcTransm.Dat_Hor_Retorno = M46V999.gfForData(3)
1510          lobjstcTransm.Dat_Ini_Vig = lobjstcPedido.Dat_Ini_Vig
1520          lobjstcTransm.Dat_Ter_Vig = lobjstcPedido.Dat_Ter_Vig
1530          lobjstcTransm.Dsc_Compl_Risco = lobjstcPedLoc.Dsc_Compl_Risco
1540          lobjstcTransm.Ind_Proposta = 1
1550          lobjstcTransm.Nom_Arq_Envio = ""
1560          lobjstcTransm.Nom_Bairro_Risco = lobjstcPedLoc.Nom_Bairro_Risco
1570          lobjstcTransm.Nom_Cid_Risco = lobjstcPedLoc.Nom_Cid_Risco
1580          lobjstcTransm.Nom_Logr_Risco = lobjstcPedLoc.Nom_Logr_Risco
1590          lobjstcTransm.Nom_Segurado = Mid(Trim$(lobjstcPedido.Nom_Segurado), 1, 50)
1600          lobjstcTransm.Num_Cep_Risco = lobjstcPedLoc.Num_Cep_Risco
1610          lobjstcTransm.Num_Item = lobjstcPedLoc.Num_Item
1620          lobjstcTransm.Num_Ped = lobjstcPedido.Num_Ped
1630          lobjstcTransm.NUM_PI = lNum_pi
1640          lobjstcTransm.Num_Risco = lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco
1650          lobjstcTransm.Sig_UF_Risco = lobjstcPedLoc.Sig_UF_Risco
1660          lobjstcTransm.Tip_Emissao = lobjstcPedLoc.Tip_Emissao
1670          Set lobjTransm.ConexaoBD = lCnSyasPortal2008
                  
1680          If Not lobjTransm.mfIncluir(lobjstcTransm, objstcUsuario.PropTabela, "") Then
1690              UltimoErro = lobjTransm.UltimoErro
1700              GoTo ControleErro
1710          End If
          
              'Atualiza o Status de transmiss�o.
              ' retirado da Dll e adicionado na procedure proc_StartFluxoEmissaoTransmissaoMulti
          '    lUpdate = " Update SyasPortal.dbo.Tab_Ped set Cod_Sit_Ped = 3 "
          '    lUpdate = lUpdate & " where nosso_numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO)
          '    If M46V999.gfExeSQL(lCnSyasPortal2008, lUpdate, UltimoErro) <> 0 Then
          '        GoTo ControleErro
          '    End If
          
              'Agenda inspe��o!
1720          For Each lobjstcPedLoc In lobjstcPedido.ColecaoPedLoc
1730              If lobjstcPedLoc.Insp_Obr = True Then
1740                  If lobjstcPedLoc.Cod_Ramo = 112 And (lobjstcPedLoc.Cod_Plano = 2218 Or lobjstcPedLoc.Cod_Plano = 2219 Or (lobjstcPedLoc.Cod_Plano >= 2228 And lobjstcPedLoc.Cod_Plano <= 2271)) Then
                          ' N�o agendar inspe��o
1750                  Else
1760                      If gfAgendarInspecao(lobjstcPedido.NOSSO_NUMERO, lobjstcPedLoc.Num_Item, lobjstcPedido.Num_Cgc_Cpf, lobjstcPedLoc.objstcPedInsp.Pess_Cont, lobjstcPedLoc.objstcPedInsp.Num_Tel, tagDynaTrace) = False Then
1770                          GoTo ControleErro
1780                      End If
1790                  End If
1800              End If
1810          Next

1820                      End If

RetornaProtocolo:
1830      gfTransmitirProposta = gfProtocoloTransmissao(pNossoNumero, "", tagDynaTrace) & "|||" & lobjstcPedido.Num_PI_Emissao

          '11.42
1840      gfLog_Emissao pNossoNumero, lNum_pi, "Fim", Enum_Tipo_LogEmissao_Fim

1850      Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
          

SAIDA:
1860      Err.Clear: On Error GoTo 0
1870      GoSub FechaObj
1880      Exit Function

ControleErro:
1890      If Err <> 0 Then
1900          UltimoErro = UltimoErro & " - " & "Ln: " & Erl & " Erro: " & Err & " " & Err.Description
1910      End If
          
          '11.42
1920      gfLog_Emissao pNossoNumero, lNum_pi, UltimoErro, Enum_Tipo_LogEmissao_Erro
          
          '###### 11.44 TRATA ERRO PARA O USUARIO
1930      If InStr(UltimoErro, "<") > 0 Then
1940          UltimoErro = Left(UltimoErro, InStr(UltimoErro, "<") - 1)
1950          If Len(UltimoErro) - (InStr(UltimoErro, "Mensagem de retorno:") + 20) > 0 Then
1960              UltimoErro = Right(UltimoErro, Len(UltimoErro) - (InStr(UltimoErro, "Mensagem de retorno:") + 20))
1970              UltimoErro = Replace(UltimoErro, "##1", Empty)
1980              UltimoErro = Replace(UltimoErro, "Input:", Empty)
1990          End If
2000      End If
          
2010      If Err > 0 Then
2020          UltimoErro = UltimoErro & Err.Description & "Ln: " & Erl
2030          End If
          '###### FIM 11.44 TRATA ERRO PARA O USUARIO
          
2040      Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
          
2050      GoTo SAIDA

FechaObj:
2060      Set lobjstcPedido = Nothing
2070      Set lobjstcBoleto = Nothing
2080      Set lConexaoEAPDB = Nothing
2090      Set lCnSyasPortal2008 = Nothing
2100      Set lConexaoDiv = Nothing
          Set lConexaoSyasEmis = Nothing
2110      Set lrst = Nothing
2120      Set lobjstcPedLoc = Nothing
2130      Return
End Function

Public Function gfReservarPI(ByVal pNossoNumero As String, tagDynaTrace As String) As String

    Dim lobjstcPedido As stcA46V702B
    Dim lConexaoEAPDB As ADODB.Connection
    Dim lCnSyasPortal2008 As ADODB.Connection
    Dim lConexaoDiv As ADODB.Connection
    Dim lDat_Ini_Vig_MAX As Date
    Dim lSelect As String
    Dim lrst As clsYasRecordSet
    Dim lNum_pi As Double

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfReservarPI"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""

    On Error GoTo ControleErro

    '- Abre as bases que ser�o utilizadas
    If Not M46V999.gfAbrBasDados("EAPDB", M46V999.gTipEAPDB, M46V999.gEndEAPDB, lConexaoEAPDB, UltimoErro) Then
        GoTo ControleErro
    End If

    '- Abre as bases que ser�o utilizadas
    If Not M46V999.gfAbrBasDados("SYASPORTAL", M46V999.gTipSyasPortal, M46V999.gEndSyasPortal, lCnSyasPortal2008, UltimoErro) Then
        GoTo ControleErro
    End If

    '- Abre as bases que ser�o utilizadas
    If Not M46V999.gfAbrBasDados("RamosDiversos", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, lConexaoDiv, UltimoErro) Then
        GoTo ControleErro
    End If

    Set lobjstcPedido = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)

    If mfVerDados(lobjstcPedido) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "mfVerDados veio false", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    '11.46 Item 2.Nova regra para vig�ncia futura
    '11.48 Item 8 Nova regra para vig�ncia futura - Seguro Novo / Renova��o Cong�nere
    If lobjstcPedido.Cod_Produto >= 1144 And _
       lobjstcPedido.Cod_Ramo = 113 And (lobjstcPedido.Cod_Estip = 10 Or lobjstcPedido.Cod_Estip = 30) Then
        
        lDat_Ini_Vig_MAX = DateAdd("d", 30, Date)
        
        If lDat_Ini_Vig_MAX < CDate(M46V999.gfForData(2, lobjstcPedido.Dat_Ini_Vig)) Then
            UltimoErro = "FGE0150 - Inicio de Vig�ncia inv�lida (superior a " & Format(lDat_Ini_Vig_MAX, "dd/MM/YYYY") & ")" & UltimoErro
            GoTo ControleErro
        End If
    End If

    lSelect = "SELECT Num_PI FROM dbo.Tab_Vinculo_Proposta_PI "
    lSelect = lSelect & "WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO)
    If Not M46V999.gfObtRegistro(lConexaoEAPDB, lSelect, lrst, UltimoErro) Then
        GoTo ControleErro
    End If
      
    If lrst.EOF Then
        'Gera o n�mero do PI
        lSelect = " exec [dbo].proc_GetProxPI 2"
        If Not M46V999.gfObtRegistro(lConexaoEAPDB, lSelect, lrst, UltimoErro) Then
            UltimoErro = " Erro ao vincular com PI - Reserva de PI " & UltimoErro
            GoTo ControleErro
        End If
        If Not lrst.EOF Then
            lNum_pi = lrst("Num_PI")
               
            lSelect = " Insert into dbo.Tab_Vinculo_Proposta_PI (Nosso_Numero, Num_PI, Data , Hora, Usuario) values ( "
            lSelect = lSelect & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO) & ","
            lSelect = lSelect & M46V999.gfFormatarCamposSQL(lNum_pi) & ","
            lSelect = lSelect & " convert(varchar,getdate(),112) ,"
            lSelect = lSelect & " replace(left(convert(varchar,getdate(),114),8),':','') ,"
            lSelect = lSelect & M46V999.gfFormatarCamposSQL(M46V999.gNomUsuario6) & ")"
             
            If M46V999.gfExeSQL(lConexaoEAPDB, lSelect, UltimoErro) <> 0 Then
                UltimoErro = " Erro ao salvar vinculo do PI. Processo de reserva do PI interrompido. " & UltimoErro
                GoTo ControleErro
            End If
        Else
            Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Procedure proc_GetProxPI n�o retornou registros. Processo de Reserva de PI interrompido.", 0, False)
            UltimoErro = " Erro ao vincular com PI - Procedure proc_GetProxPI n�o retornou registros. Processo de Reserva de PI interrompido."
            GoTo ControleErro
        End If
           
    Else
        lNum_pi = lrst("Num_PI")
    End If
       
    gfLog_Emissao pNossoNumero, lNum_pi, "ReservaPI", 5
    
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    
    gfReservarPI = lNum_pi

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    If Err <> 0 Then
        UltimoErro = UltimoErro & " - " & "Ln: " & Erl & " Erro: " & Err & " " & Err.Description
    End If
          
    gfLog_Emissao pNossoNumero, lNum_pi, UltimoErro, Enum_Tipo_LogEmissao_Erro

    '###### 11.44 TRATA ERRO PARA O USUARIO
    If InStr(UltimoErro, "<") > 0 Then
        UltimoErro = Left(UltimoErro, InStr(UltimoErro, "<") - 1)
        If Len(UltimoErro) - (InStr(UltimoErro, "Mensagem de retorno:") + 20) > 0 Then
            UltimoErro = Right(UltimoErro, Len(UltimoErro) - (InStr(UltimoErro, "Mensagem de retorno:") + 20))
            UltimoErro = Replace(UltimoErro, "##1", Empty)
            UltimoErro = Replace(UltimoErro, "Input:", Empty)
        End If
    End If

    If Err > 0 Then
        UltimoErro = UltimoErro & Err.Description & "Ln: " & Erl
        End If
    '###### FIM 11.44 TRATA ERRO PARA O USUARIO

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)

    GoTo SAIDA

FechaObj:
    Set lobjstcPedido = Nothing
    Set lConexaoEAPDB = Nothing
    Set lCnSyasPortal2008 = Nothing
    Set lConexaoDiv = Nothing
    Set lrst = Nothing
    
    Return
    
End Function

Private Function mfTransmitirPropostaOFFLine(ByVal pNossoNumero As String, _
                                             ByVal pReUtilizaBoleto As Boolean, tagDynaTrace As String) As String
    Dim lobjstcPedido As stcA46V702B
    Dim lobjstcBoleto As stcA99V009B
    Dim lCnSyasPortal2008 As ADODB.Connection
    Dim lConexaoDiv As ADODB.Connection
    Dim lAlertas As String
    Dim lRetornoEhErro As Boolean
    Dim lSelect As String
    Dim lrst As clsYasRecordSet
    Dim lobjstcPedLoc As stcA46V708B
    Dim lObjA46V101A As clsA46V101A
    Dim lUsuario As String
    Dim lColecao As Collection
    Dim lArqTransmissao As String

    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro
    Dim lUpdate As String

    '- Abre as bases que ser�o utilizadas
    If Not M46V999.gfAbrBasDados("SYASPORTAL", M46V999.gTipSyasPortal, M46V999.gEndSyasPortal, lCnSyasPortal2008, UltimoErro) Then
        GoTo ControleErro
    End If

    '- Abre as bases que ser�o utilizadas
    If Not M46V999.gfAbrBasDados("RamosDiversos", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, lConexaoDiv, UltimoErro) Then
        GoTo ControleErro
    End If

    Set lobjstcPedido = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)


    If mfVerDados(lobjstcPedido) = False Then
        GoTo SAIDA
    End If

ExecutaTransmissao:

    'Busco o n�mero do PI
    lSelect = " exec [dbo].proc_GetProxPI 2"
    If Not M46V999.gfObtRegistro(lCnSyasPortal2008, lSelect, lrst, UltimoErro) Then
        UltimoErro = " Erro ao vincular com PI" & UltimoErro
    End If
    lobjstcPedido.Num_PI_Emissao = lrst("num_pi")
    
    MensagemAlerta = lAlertas
    If lobjstcPedido.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 And lobjstcPedido.objstcCtrlEmis.Num_Bloq = 0 And lobjstcPedido.Cod_Reg <> 120 And _
      lobjstcPedido.Cod_Reg <> 106 And lobjstcPedido.Cod_Reg <> 107 And lobjstcPedido.Cod_Reg <> 129 Then
        'Set lobjstcBoleto = gfRetornaBoleto(lobjstcPedido.Nosso_Numero, pReUtilizaBoleto, tagDynaTrace)
        Set lobjstcBoleto = gfRetornaBoletoWS(lobjstcPedido.NOSSO_NUMERO, pReUtilizaBoleto, tagDynaTrace, lobjstcPedido.Num_PI_Emissao)
        If Not lobjstcBoleto Is Nothing Then
            'Recarrega o Pedido, Agora com o Boleto gerado
            Set lobjstcPedido = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)
        Else
            UltimoErro = "Falha ao carregar boleto:" & UltimoErro
            GoTo ControleErro
        End If
    End If

    'Gera o Arquivo
    lRetornoEhErro = True

    '2. Inicia processo.
    '4. Obt�m registros de protocolo referente ao nosso n�mero.
    'a) Abre base de dados.
    Dim lobjTrans As clsA46V726A

    Dim lobjTransm As clsA46V718A    'Objeto da classe clsA46V718A (TAB_PED_TRANSM)
    Dim lobjstcTransm As stcA46V718B    'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM)

    Set lobjTransm = New clsA46V718A            'Objeto da classe clsA46V718A (TAB_PED_TRANSM)
    Set lobjTransm.M46V999 = M46V999

    Set lobjstcTransm = New stcA46V718B         'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM)

    Set lobjstcPedLoc = lobjstcPedido.ColecaoPedLoc.Item(1)

    lobjstcTransm.Cod_Ramo = lobjstcPedido.Cod_Ramo
    lobjstcTransm.Cod_Tip_Logr_Risco = lobjstcPedLoc.Cod_Tip_Logr_Risco
    lobjstcTransm.Dat_Hor_Envio = M46V999.gfForData(3)
    lobjstcTransm.Dat_Hor_Retorno = M46V999.gfForData(3)
    lobjstcTransm.Dat_Ini_Vig = lobjstcPedido.Dat_Ini_Vig
    lobjstcTransm.Dat_Ter_Vig = lobjstcPedido.Dat_Ter_Vig
    lobjstcTransm.Dsc_Compl_Risco = lobjstcPedLoc.Dsc_Compl_Risco
    lobjstcTransm.Ind_Proposta = 1
    lobjstcTransm.Nom_Arq_Envio = ""
    lobjstcTransm.Nom_Bairro_Risco = lobjstcPedLoc.Nom_Bairro_Risco
    lobjstcTransm.Nom_Cid_Risco = lobjstcPedLoc.Nom_Cid_Risco
    lobjstcTransm.Nom_Logr_Risco = lobjstcPedLoc.Nom_Logr_Risco
    lobjstcTransm.Nom_Segurado = Mid(Trim$(lobjstcPedido.Nom_Segurado), 1, 50)
    lobjstcTransm.NOSSO_NUMERO = lobjstcPedido.NOSSO_NUMERO
    lobjstcTransm.Num_Cep_Risco = lobjstcPedLoc.Num_Cep_Risco
    lobjstcTransm.Num_Item = lobjstcPedLoc.Num_Item
    lobjstcTransm.Num_Ped = lobjstcPedido.Num_Ped
    lobjstcTransm.NUM_PI = lrst("num_pi")
    lobjstcTransm.Num_Risco = lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco
    lobjstcTransm.Sig_UF_Risco = lobjstcPedLoc.Sig_UF_Risco
    lobjstcTransm.Tip_Emissao = lobjstcPedLoc.Tip_Emissao
    Set lobjTransm.ConexaoBD = lCnSyasPortal2008
    If Not lobjTransm.mfIncluir(lobjstcTransm, objstcUsuario.PropTabela, "") Then
        UltimoErro = lobjTransm.UltimoErro
        GoTo ControleErro
    End If

    'Atualiza o Status de transmiss�o.
    lUpdate = " Update SyasPortal.dbo.Tab_Ped set Cod_Sit_Ped = 3 "
    lUpdate = lUpdate & " where nosso_numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.NOSSO_NUMERO)
    If M46V999.gfExeSQL(lCnSyasPortal2008, lUpdate, UltimoErro) <> 0 Then
        GoTo ControleErro
    End If

    'Agenda inspe��o!
    For Each lobjstcPedLoc In lobjstcPedido.ColecaoPedLoc
        If lobjstcPedLoc.Insp_Obr = True Then
            If mfAgendarInspecaoArquivoOFFLine(lobjstcPedido.NOSSO_NUMERO, lobjstcPedLoc.Num_Item, lobjstcPedido.Num_Cgc_Cpf, lobjstcPedLoc.objstcPedInsp.Pess_Cont, lobjstcPedLoc.objstcPedInsp.Num_Tel, tagDynaTrace) = False Then
                GoTo ControleErro
            End If

            lSelect = "Select NOSSO_NUMERO_OLD From Tab_Ped_Num_Copia WITH (NOLOCK)"
            lSelect = lSelect & " Where NOSSO_NUMERO_NEW = '" & lobjstcPedido.NOSSO_NUMERO & "' "
            If M46V999.gfObtRegistro(lConexaoDiv, lSelect, lrst, UltimoErro) = False Then
                Call M46V999.gpGraLog(2, UltimoErro)
                GoTo SAIDA
            End If


            'Atualiza o Inspecao
            lUpdate = " update ramosdiversos..tab_trans_insp set cod_status_i = 1 where  cod_status_i = 0 and nosso_numero = " & lrst("NOSSO_NUMERO_OLD")
            If M46V999.gfExeSQL(lConexaoDiv, lUpdate, UltimoErro) <> 0 Then
                GoTo ControleErro
            End If
        End If
    Next

    'Gera o Arquivo!
    Set lObjA46V101A = New clsA46V101A
    Set lObjA46V101A.M46V999 = M46V999

    lUsuario = mUsuarioTransmissao
    Set lColecao = New Collection
    lColecao.Add lobjstcPedido
    If lObjA46V101A.gfCriArqTransmissao(objstcUsuario.codCorretor, lUsuario, _
                                        lColecao, lArqTransmissao, UltimoErro) = False Then
        GoTo ControleErro
    End If


RetornaProtocolo:
    mfTransmitirPropostaOFFLine = gfProtocoloTransmissao(pNossoNumero, "", tagDynaTrace) & "|||" & lobjstcPedido.Num_PI_Emissao

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    UltimoErro = "[ERR] - gfProtocoloTransmissao ERRO - " & Err & " " & Err.Description & "-" & UltimoErro
    GoTo SAIDA

FechaObj:
    Set lobjstcPedido = Nothing
    Set lobjstcBoleto = Nothing
    Set lCnSyasPortal2008 = Nothing
    Set lrst = Nothing
    Set lobjstcPedLoc = Nothing
    Set lObjA46V101A = Nothing
    Set lColecao = Nothing
    Return
End Function


Private Function mfArrayToColecao(ByRef pArray As Variant) As Collection
    Dim lContX As Integer
    Dim lColTemp As New Collection

    For lContX = 1 To UBound(pArray)
        lColTemp.Add pArray(lContX)
    Next

    Set mfArrayToColecao = lColTemp

End Function
Public Function gfVerifReUsoBoleto(ByVal pNossoNumero As String, ByRef pMsgAlerta As String, tagDynaTrace As String) As E_VerifBoleto
    Dim lObjstcPed As stcA46V702B
    Dim lobjstcDadosBoleto As stcA28V720B
    Dim lobjDadosBoleto As clsA28V720A

    Dim lretorno As E_VerifBoleto

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfVerifReUsoBoleto"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "N�o posicionou usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    Set lObjstcPed = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)
    If lObjstcPed Is Nothing Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "N�o encontrou nosso n�mero", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    If lObjstcPed.Ind_Reentrada = 1 Then
        If lObjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 And lObjstcPed.objstcCtrlEmis.Num_Bloq = 0 And _
        (lObjstcPed.Num_Cta_Bas_Cor = 0 Or lObjstcPed.objstcCtrlEmis.Cod_Bco_Boleto > 0) And _
        lObjstcPed.Cod_Reg <> 120 And lObjstcPed.Cod_Reg <> 106 And lObjstcPed.Cod_Reg <> 107 And lObjstcPed.Cod_Reg <> 129 Then

            Set lobjDadosBoleto = New clsA28V720A
            Set lobjDadosBoleto.M46V999 = M46V999

            If Not lobjDadosBoleto.gfCarregarBoletoDisp(lObjstcPed.NOSSO_NUMERO, False, lObjstcPed.Nosso_Numero_Original, True) Then
                UltimoErro = lobjDadosBoleto.UltimoErro
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
                GoTo SAIDA
            End If

            If lobjDadosBoleto.Colecao.Count = 0 Then
                UltimoErro = "Boleto anterior n�o localizado"
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
                GoTo SAIDA
            Else
                Set lobjstcDadosBoleto = lobjDadosBoleto.Colecao(1)
            End If

            If lobjstcDadosBoleto.VAL_PR <> lObjstcPed.Val_PriParcela Then
                'Gero outro boleto
                pMsgAlerta = "AGE0163 - Um novo boleto ser� impresso. Favor descartar o boleto emitido anteriormente."
                lretorno = E_VerifBoleto_Alerta
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
                GoTo Retorno
            ElseIf lobjstcDadosBoleto.DAT_DIGIT < Format(Date - 5, "yyyymmdd") And _
                   lobjstcDadosBoleto.DAT_VENC < Format(Date + 7, "yyyymmdd") Then
                'Gero outro boleto
                If lobjstcDadosBoleto.NOSSO_NUMERO = lObjstcPed.Nosso_Numero_Original Then
                    lretorno = E_VerifBoleto_Questiona
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
                    GoTo Retorno
                Else
                    pMsgAlerta = "AGE0163 - Um novo boleto ser� impresso. Favor descartar o boleto emitido anteriormente."
                    lretorno = E_VerifBoleto_Alerta
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
                    GoTo Retorno
                End If
            Else
                'Gero outro boleto
                lretorno = E_VerifBoleto_Questiona
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
                GoTo Retorno
            End If
        Else
            lretorno = E_VerifBoleto_Nao
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
            GoTo Retorno
        End If
    Else
        lretorno = E_VerifBoleto_Nao
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
        GoTo Retorno
    End If

    If UltimoErro = "" Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Else
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    End If
Retorno:
    gfVerifReUsoBoleto = lretorno

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

FechaObj:
    Set lObjstcPed = Nothing
    Set lobjstcDadosBoleto = Nothing
    Set lobjDadosBoleto = Nothing
    Return
End Function

Public Function gfAgendarInspecao(pNossoNumero As String, pNumItem As Integer, pCPFCNPJ As String, pNomContato As String, pTelContato As String, tagDynaTrace As String) As Boolean
    Dim lObjstcPed As stcA46V702B
    Dim lobjstcPedLoc As stcA46V708B
    Dim lobjPedInsp As clsA46V717A
    Dim lObjPed As clsA46V702A
    Dim lobjPedLoc As clsA46V708A

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfAgendarInspecao"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    If M46V999.gTipoTransmDLL = ETransmDLL_OFFLine Then
        gfAgendarInspecao = mfAgendarInspecaoArquivoOFFLine(pNossoNumero, pNumItem, pCPFCNPJ, pNomContato, pTelContato, tagDynaTrace)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
        GoTo SAIDA
    End If

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", 0, Class_Name, methodName)
        GoTo SAIDA
    End If

    Set lObjstcPed = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)

    If lObjstcPed Is Nothing Then
        UltimoErro = "FNL0003 - Pedido n�o localizado com nosso n�mero " & pNossoNumero & "."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If


    For Each lobjstcPedLoc In lObjstcPed.ColecaoPedLoc
        If lobjstcPedLoc.Insp_Obr = True And lobjstcPedLoc.Insp_Agend = False And (lobjstcPedLoc.Num_Item = pNumItem Or pNumItem = 0) Then

            Set lobjPedInsp = New clsA46V717A
            Set lobjPedInsp.M46V999 = M46V999

            lobjstcPedLoc.objstcPedInsp.Num_Tel = pTelContato
            lobjstcPedLoc.objstcPedInsp.Pess_Cont = pNomContato
            If Not lobjPedInsp.mfAlterar(lobjstcPedLoc.objstcPedInsp, objstcUsuario.PropTabela, "") Then
                UltimoErro = lobjPedInsp.UltimoErro
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
                GoTo SAIDA
            End If

            'Efetuar a
            If Not M46V111.gfGravarInspecaoProduto(lObjstcPed, lobjstcPedLoc, UltimoErro, 1) Then
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
                GoTo SAIDA
            End If

            Set lobjPedLoc = New clsA46V708A
            Set lobjPedLoc.M46V999 = M46V999

            Set lobjPedLoc.ConexaoBD = lobjPedInsp.ConexaoBD
            If Not lobjPedLoc.mfAlterarAgendaInspecao(lobjstcPedLoc, objstcUsuario.PropTabela, "") Then
                UltimoErro = lobjPedLoc.UltimoErro
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
                GoTo SAIDA
            End If
        End If
    Next lobjstcPedLoc

    lObjstcPed.Num_Cgc_Cpf = pCPFCNPJ
    lObjstcPed.Cod_Sit_Ped = e_SitPedido_Transmitido
    Set lObjPed = New clsA46V702A
    Set lObjPed.M46V999 = M46V999

    If Not lObjPed.mfAlterar(lObjstcPed, objstcUsuario.PropTabela, "") Then
        UltimoErro = lObjPed.UltimoErro
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If
    gfAgendarInspecao = True

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    If Not M46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If
    
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    
    GoTo SAIDA

FechaObj:
    Set lObjstcPed = Nothing
    Set lobjstcPedLoc = Nothing
    Set lobjPedInsp = Nothing
    Set lObjPed = Nothing
    Set lobjPedLoc = Nothing
    Return
End Function


Private Function mfAgendarInspecaoArquivoOFFLine(pNossoNumero As String, _
                                                pNumItem As Integer, _
                                                pCPFCNPJ As String, _
                                                pNomContato As String, _
                                                pTelContato As String, _
                                                tagDynaTrace As String) As Boolean
    Dim lObjstcPed As stcA46V702B
    Dim lobjstcPedLoc As stcA46V708B
    Dim lobjPedInsp As clsA46V717A
    Dim lObjPed As clsA46V702A
    Dim lobjPedLoc As clsA46V708A

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfAgendarInspecaoArquivoOFFLine"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    If Not mfPosicionaUsuario(tagDynaTrace) Then GoTo SAIDA

    Set lObjstcPed = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)

    If lObjstcPed Is Nothing Then
        UltimoErro = "FNL0003 - Pedido n�o localizado com nosso n�mero " & pNossoNumero & "."
        GoTo SAIDA
    End If


    For Each lobjstcPedLoc In lObjstcPed.ColecaoPedLoc
        If lobjstcPedLoc.Insp_Obr = True And lobjstcPedLoc.Insp_Agend = False And (lobjstcPedLoc.Num_Item = pNumItem Or pNumItem = 0) Then

            Set lobjPedInsp = New clsA46V717A
            Set lobjPedInsp.M46V999 = M46V999

            lobjstcPedLoc.objstcPedInsp.Num_Tel = pTelContato
            lobjstcPedLoc.objstcPedInsp.Pess_Cont = pNomContato
            If Not lobjPedInsp.mfAlterar(lobjstcPedLoc.objstcPedInsp, objstcUsuario.PropTabela, "") Then
                UltimoErro = lobjPedInsp.UltimoErro
                GoTo SAIDA
            End If


            Set lobjPedLoc = New clsA46V708A
            Set lobjPedLoc.M46V999 = M46V999

            Set lobjPedLoc.ConexaoBD = lobjPedInsp.ConexaoBD
            If Not lobjPedLoc.mfAlterarAgendaInspecao(lobjstcPedLoc, objstcUsuario.PropTabela, "") Then
                UltimoErro = lobjPedLoc.UltimoErro
                GoTo SAIDA
            End If
            lobjstcPedLoc.Insp_Agend = True

            'Efetuar a
            If mfGeraLogCalcInspecao(lObjstcPed, lobjstcPedLoc) = False Then
                GoTo SAIDA
            End If

        End If
    Next lobjstcPedLoc

    lObjstcPed.Num_Cgc_Cpf = pCPFCNPJ
    lObjstcPed.Cod_Sit_Ped = e_SitPedido_Transmitido
    Set lObjPed = New clsA46V702A
    Set lObjPed.M46V999 = M46V999

    If Not lObjPed.mfAlterar(lObjstcPed, objstcUsuario.PropTabela, "") Then
        UltimoErro = lObjPed.UltimoErro
        GoTo SAIDA
    End If
    mfAgendarInspecaoArquivoOFFLine = True

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    If Not M46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lObjstcPed = Nothing
    Set lobjstcPedLoc = Nothing
    Set lobjPedInsp = Nothing
    Set lObjPed = Nothing
    Set lobjPedLoc = Nothing
    Return
End Function

Public Function gfSolicitaRecotacao(ByVal pNossoNumero As String, ByVal pNumSiscota As String, tagDynaTrace As String) As stcA46V702B
    Dim lNovoNossoNumero As String
    Dim lNovReferencia As String
    Dim lobjstcTransm As stcA28V721B
    Dim lobjTransm As clsA28V721A
    Dim lObjstcPed As stcA46V702B

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfSolicitaRecotacao"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro
    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    'Quando � solicitado uma recota��o ent�o precisa ser feita em cima de um outro nosso numero
    'por isso � feito a c�pia.
    Set lObjstcPed = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)

    Set lObjstcPed = gfCopiarProposta(pNossoNumero, 2, tagDynaTrace)
    Set lobjstcTransm = New stcA28V721B

    lobjstcTransm.NOSSO_NUMERO = lObjstcPed.NOSSO_NUMERO
    lobjstcTransm.CCotacaoOrig = pNumSiscota
    lobjstcTransm.STATUS = -1
    Set lobjTransm = New clsA28V721A
    Set lobjTransm.M46V999 = M46V999

    If Not lobjTransm.gfIncluir(lobjstcTransm, objstcUsuario.PropTabela) Then
        UltimoErro = lobjTransm.UltimoErro
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    Set gfSolicitaRecotacao = lObjstcPed

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function
ControleErro:
    If Not M46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If
    
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    
    GoTo SAIDA
FechaObj:
    Set lobjstcTransm = Nothing
    Set lobjTransm = Nothing
    Set lObjstcPed = Nothing
    Return
End Function

Public Function gfSolicitaRecotacaoClasse4(ByVal pNossoNumero As String, _
                                           ByVal pEmail As String, _
                                           ByVal pTelefone As String, _
                                           ByVal pObservacoes As String, _
                                           ByVal pNome As String, _
                                           ByVal pErrosClasse4 As String, _
                                           ByVal pClasse As String, _
                                           ByVal pObjstc753B As stcA46V753B, _
                                           ByVal pLogin As String, _
                                           tagDynaTrace As String) As stcA46V702B
    
    Dim lObjstcPed As stcA46V702B
    Dim lobjstcTransm As stcA28V721B
    Dim lobjTransm As clsA28V721A
    Dim lMensageErro As String
    Dim lObjstcSiscota As stcA46V753B
    Dim strNumCotacao As String
    Dim lstcItem As stcA46V708B

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfSolicitaRecotacaoClasse4"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    Set lObjstcPed = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)
    
    If lObjstcPed Is Nothing Then
        UltimoErro = "Cota��o n�o localizada " & pNossoNumero
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    'Atualiza campo ITEM_ORIGEM_C4 de 0 para 1 na solicita��o de cota��o
    For Each lstcItem In lObjstcPed.ColecaoPedLoc
        lstcItem.ITEM_ORIGEM_C4 = 1
    Next lstcItem

    If lObjstcPed Is Nothing Then
        UltimoErro = "Cota��o n�o localizada " & pNossoNumero
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    Set lObjstcPed = gfCopiarPropostaClasse4(pNossoNumero, lObjstcPed.Num_Ped, lObjstcPed.objstcCtrlEmis.Num_Cotac, 2, tagDynaTrace)
    If M46V999.gfPreenchido(UltimoErro) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    On Error GoTo ControleErro

    Set lobjstcTransm = New stcA28V721B
    Set lobjTransm = New clsA28V721A
    Set lobjTransm.M46V999 = M46V999
    
    With lobjstcTransm
        .CCotacao = ""
        .Dat_Cotacao = Left(M46V999.gfForData(3), 8)
        .Email = pEmail
        .Nome = pNome
        .Telefone = pTelefone
        .NOSSO_NUMERO = lObjstcPed.NOSSO_NUMERO
        .STATUS = 0
        .Usuario = objstcUsuario.login_user_rede
        .CCotacaoOrig = lObjstcPed.objstcCtrlEmis.Num_Cotac_Ant
    End With
    
    If Not lobjTransm.gfIncluir(lobjstcTransm, objstcUsuario.PropTabela) Then
        UltimoErro = lobjTransm.UltimoErro
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    lMensageErro = "** Erros classe 4: " & Replace(pErrosClasse4, vbCrLf, "%!*")
    lMensageErro = lMensageErro & " **  Observa��o do corretor:" & Replace(M46V999.gfRetirarCaracterEspecial(pObservacoes), vbCrLf, "%!*")
    
    If Not M46V999.gfTransmissaoInterna(2, lObjstcPed, UltimoErro) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    If Not mfIncluirSiscotaMulti(lObjstcPed.NOSSO_NUMERO, lobjstcTransm.CCotacaoOrig, lobjstcTransm, lMensageErro, lObjstcPed.Cod_Ramo, pClasse, UltimoErro, strNumCotacao) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    'Atualiza Status Siscota do nosso numero atual
    If Not gfAtualizarStatusSiscota(pObjstc753B, pLogin) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    'Inclui Siscota para novo nosso numero gerado no Repique
    Set lObjstcSiscota = New stcA46V753B
    
    Set lObjstcSiscota = pObjstc753B
    lObjstcSiscota.IdSiscota = 0
    lObjstcSiscota.NOSSO_NUMERO = lObjstcPed.NOSSO_NUMERO
    lObjstcSiscota.Num_Ped = lObjstcPed.Num_Ped
    lObjstcSiscota.Siscota = strNumCotacao
    lObjstcSiscota.STATUS = 2
    lObjstcSiscota.ColecaoObs(1).CodOrigem = 1
    
    If Not gfObsSISCOTA_Salvar(pObjstc753B, pLogin, lMensageErro) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    Set gfSolicitaRecotacaoClasse4 = lObjstcPed

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    
SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
ControleErro:
    If Not M46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If
    
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    
    GoTo SAIDA
    
FechaObj:
    Set lobjstcTransm = Nothing
    Set lobjTransm = Nothing
    Set lObjstcPed = Nothing
    Return
    
End Function

Public Function gfSolicitarClasse4(ByVal pNossoNumero As String, _
                                   ByVal pEmail As String, _
                                   ByVal pTelefone As String, _
                                   ByVal pObservacoes As String, _
                                   ByVal pNome As String, _
                                   ByVal pErrosClasse4 As String, _
                                   ByVal pClasse As String, _
                                   ByVal pObjstc753B As stcA46V753B, _
                                   ByVal pLogin As String, _
                                   tagDynaTrace As String) As Boolean
    Dim lNosNumero As String
    Dim lNovReferencia As String
    Dim lObjstcPed As stcA46V702B
    Dim lobjstcTransm As stcA28V721B
    Dim lobjTransm As clsA28V721A
    Dim lUsuario As Long
    Dim lCodCorretor As Long
    Dim lArqIni As String
    Dim lPropostas(1) As String         'Matriz com propostas selecionadas (neste caso, apenas 1).
    Dim lArqTransmissao As String
    Dim lMensageErro As String
    Dim lstcItem As stcA46V708B
    Dim lstcCobert As stcA46V704B
    Dim strNumCotacao As String

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfSolicitarClasse4"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    If M46V999.gTipoTransmDLL = ETransmDLL_OFFLine Then
        gfSolicitarClasse4 = mfSolicitarClasse4OFFLine(pNossoNumero, pEmail, pTelefone, pObservacoes, pNome, pErrosClasse4, tagDynaTrace)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Transmiss�o Offline", 0, Class_Name, methodName)
        GoTo SAIDA
    End If

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    Set lObjstcPed = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)

    If lObjstcPed Is Nothing Then
        UltimoErro = "Cota��o n�o localizada " & pNossoNumero
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    'Atualiza campo ITEM_ORIGEM_C4 de 0 para 1 na solicita��o de cota��o
    For Each lstcItem In lObjstcPed.ColecaoPedLoc
        lstcItem.ITEM_ORIGEM_C4 = 1
        For Each lstcCobert In lstcItem.ColecaoPedCobert
            lstcCobert.Tip_Origem_C4 = 1
        Next lstcCobert
    Next lstcItem
    
    'Quando um c�lculo � visualizado ent�o pode ser que o corretor tenha impresso a cota��o (dat_pri_calculo <>0)
    'Nestes casos n�o poderemos realizar o classe 4 com o mesmo nosso_numero  porque se n�o pode ser que a cota��o
    'seja declinada e ele apresente a primeira impress�o para ser emitida, desta forma ser� recusado indevidamente
    'na emiss�o.

    If lObjstcPed.Dat_Pri_Calculo <> 0 Then
        Set lObjstcPed = gfCopiarProposta(pNossoNumero, 2, tagDynaTrace)
        If M46V999.gfPreenchido(UltimoErro) Then
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
            GoTo SAIDA
        End If
        lNosNumero = lObjstcPed.NOSSO_NUMERO
    Else
        lNosNumero = pNossoNumero
    End If

    On Error GoTo ControleErro

    Set lobjstcTransm = New stcA28V721B
    Set lobjTransm = New clsA28V721A
    Set lobjTransm.M46V999 = M46V999

    lobjstcTransm.NOSSO_NUMERO = lNosNumero
    lobjstcTransm.STATUS = -1

    If Not lobjTransm.gfCarregarColecao(lobjstcTransm, objstcUsuario.PropTabela) Then
        UltimoErro = lobjTransm.UltimoErro
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    If lobjTransm.Colecao.Count > 0 Then
        Set lobjstcTransm = lobjTransm.Colecao.Item(1)
    End If

    With lobjstcTransm
        .CCotacao = ""
        .Dat_Cotacao = Left(M46V999.gfForData(3), 8)
        .Email = pEmail
        .Nome = pNome
        .Telefone = pTelefone
        .NOSSO_NUMERO = lNosNumero
        .STATUS = 0
        .Usuario = objstcUsuario.login_user_rede
    End With

    If lobjTransm.Colecao.Count > 0 Then
        If Not lobjTransm.gfAlterar(lobjstcTransm, objstcUsuario.PropTabela) Then
            UltimoErro = lobjTransm.UltimoErro
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
    Else
        If Not lobjTransm.gfIncluir(lobjstcTransm, objstcUsuario.PropTabela) Then
            UltimoErro = lobjTransm.UltimoErro
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If
    lPropostas(1) = lNosNumero
    If objstcUsuario.codCorretor <> 0 Then
        lCodCorretor = objstcUsuario.codCorretor
    Else
        lCodCorretor = 9999999
    End If


    lMensageErro = "** Erros classe 4: " & Replace(pErrosClasse4, vbCrLf, "%!*")
    lMensageErro = lMensageErro & " **  Observa��o do corretor:" & Replace(M46V999.gfRetirarCaracterEspecial(pObservacoes), vbCrLf, "%!*")

    'Copia as dados do SQL 2008 (portal) para o SQL 2000 (produ��o)
    'A c�pia � feita para as bases do produto (GED)
    If Not M46V999.gfTransmissaoInterna(2, lObjstcPed, UltimoErro) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    'Realizo ent�o o start do fluxo de cota��o (SQL 2000 - produ��o)
    If Not mfIncluirSiscotaMulti(lNosNumero, lobjstcTransm.CCotacaoOrig, lobjstcTransm, lMensageErro, lObjstcPed.Cod_Ramo, pClasse, UltimoErro, strNumCotacao) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    'Atualiza Status Siscota
    pObjstc753B.Siscota = strNumCotacao
    If Not gfAtualizarStatusSiscota(pObjstc753B, pLogin) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    gfSolicitarClasse4 = True

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function
ControleErro:
    If Not M46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA
FechaObj:
    Set lObjstcPed = Nothing
    Set lobjstcTransm = Nothing
    Set lobjTransm = Nothing
    Return
End Function

Public Function gfSolicitarClasse4RiscosNomeados(ByVal pNossoNumero As String, _
                                                 ByVal pEmail As String, _
                                                 ByVal pTelefone As String, _
                                                 ByVal pObservacoes As String, _
                                                 ByVal pNome As String, _
                                                 ByVal pErrosClasse4 As String, _
                                                 ByVal pClasse As String, _
                                                 ByVal pObjstc753B As stcA46V753B, _
                                                 ByVal pLogin As String, _
                                                 tagDynaTrace As String) As Boolean
    ''ByVal pObservacoes As String,
    Dim lNosNumero As String
    Dim lNovReferencia As String
    Dim lObjstcPed As stcA46V702B
    Dim lobjstcTransm As stcA28V721B
    Dim lobjTransm As clsA28V721A
    Dim lUsuario As Long
    Dim lCodCorretor As Long
    Dim lArqIni As String
    Dim lPropostas(1) As String         'Matriz com propostas selecionadas (neste caso, apenas 1).
    Dim lArqTransmissao As String
    Dim lMensageErro As String
    Dim lstcItem As stcA46V708B

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfSolicitarClasse4"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

''2017-12-04: Removido
'    If M46V999.gTipoTransmDLL = ETransmDLL_OFFLine Then
'        gfSolicitarClasse4RiscosNomeados = mfSolicitarClasse4OFFLine(pNossoNumero, pEmail, pTelefone, pObservacoes, pNome, pErrosClasse4, tagDynaTrace)
'        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Transmiss�o Offline", 0, Class_Name, methodName)
'        GoTo SAIDA
'    End If

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    Set lObjstcPed = gfProcuraNossoNumero(pNossoNumero, False, "SyasPortal", tagDynaTrace)

    If lObjstcPed Is Nothing Then
        UltimoErro = "Cota��o n�o localizada " & pNossoNumero
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    'Atualiza campo ITEM_ORIGEM_C4 de 0 para 1 na solicita��o de cota��o
    For Each lstcItem In lObjstcPed.ColecaoPedLoc
        lstcItem.ITEM_ORIGEM_C4 = 1
    Next lstcItem
    
    'Quando um c�lculo � visualizado ent�o pode ser que o corretor tenha impresso a cota��o (dat_pri_calculo <>0)
    'Nestes casos n�o poderemos realizar o classe 4 com o mesmo nosso_numero  porque se n�o pode ser que a cota��o
    'seja declinada e ele apresente a primeira impress�o para ser emitida, desta forma ser� recusado indevidamente
    'na emiss�o.

''2017-12-04: Removido
'    If lObjstcPed.Dat_Pri_Calculo <> 0 Then
'        Set lObjstcPed = gfCopiarProposta(pNossoNumero, 2, tagDynaTrace)
'        If M46V999.gfPreenchido(UltimoErro) Then
'            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
'            GoTo SAIDA
'        End If
'        lNosNumero = lObjstcPed.NOSSO_NUMERO
'    Else
'        lNosNumero = pNossoNumero
'    End If
    lNosNumero = lObjstcPed.NOSSO_NUMERO

    On Error GoTo ControleErro

     Set lobjstcTransm = New stcA28V721B
     Set lobjTransm = New clsA28V721A
     Set lobjTransm.M46V999 = M46V999

    lobjstcTransm.NOSSO_NUMERO = lNosNumero
    lobjstcTransm.STATUS = -1

    If Not lobjTransm.gfCarregarColecao(lobjstcTransm, objstcUsuario.PropTabela) Then
        UltimoErro = lobjTransm.UltimoErro
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    If lobjTransm.Colecao.Count > 0 Then
        Set lobjstcTransm = lobjTransm.Colecao.Item(1)
    End If


'' ---------------------------------------------------------------------
'' 2017-12-04: Parametros importantes, colocar a mesma regra na chamada
'' ---------------------------------------------------------------------
    With lobjstcTransm
        .CCotacao = ""
        .Dat_Cotacao = Left(M46V999.gfForData(3), 8)
        .Email = pEmail
        .Nome = pNome
        .Telefone = pTelefone
        .NOSSO_NUMERO = lNosNumero
        .STATUS = 0
        .Usuario = objstcUsuario.login_user_rede
    End With

    If lobjTransm.Colecao.Count > 0 Then
        If Not lobjTransm.gfAlterar(lobjstcTransm, objstcUsuario.PropTabela) Then
            UltimoErro = lobjTransm.UltimoErro
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
    Else
        If Not lobjTransm.gfIncluir(lobjstcTransm, objstcUsuario.PropTabela) Then
            UltimoErro = lobjTransm.UltimoErro
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If

    lPropostas(1) = lNosNumero
    If objstcUsuario.codCorretor <> 0 Then
        lCodCorretor = objstcUsuario.codCorretor
    Else
        lCodCorretor = 9999999
    End If

''2017-12-04: Removido
    lMensageErro = "** Erros classe 4: " & Replace(pErrosClasse4, vbCrLf, "%!*")
    lMensageErro = lMensageErro & " **  Observa��o do corretor:" & Replace(M46V999.gfRetirarCaracterEspecial(pObservacoes), vbCrLf, "%!*")

'''2017-12-04: Removido
'    'Copia as dados do SQL 2008 (portal) para o SQL 2000 (produ��o)
'    'A c�pia � feita para as bases do produto (GED)
'    If Not M46V999.gfTransmissaoInterna(2, lObjstcPed, UltimoErro) Then
'        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
'        GoTo SAIDA
'    End If

    'Realizo ent�o o start do fluxo de cota��o (SQL 2000 - produ��o)
    'If Not mfIncluirSiscotaMulti(lNosNumero, lobjstcTransm.CCotacaoOrig, lobjstcTransm, lMensageErro, lObjstcPed.Cod_Ramo, pClasse, UltimoErro) Then
    If Not mfIncluirSiscotaMultiRiscosNomeados(lNosNumero, lobjstcTransm.CCotacaoOrig, lobjstcTransm, lMensageErro, lObjstcPed.Cod_Ramo, pClasse, UltimoErro) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If
    
    'Atualiza Status Siscota
    If Not gfAtualizarStatusSiscota(pObjstc753B, pLogin) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    gfSolicitarClasse4RiscosNomeados = True

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function
ControleErro:
    If Not M46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA
FechaObj:
    Set lObjstcPed = Nothing
    Set lobjstcTransm = Nothing
    Set lobjTransm = Nothing
    Return
End Function

Private Function mfSolicitarClasse4OFFLine(ByVal pNossoNumero As String, _
                                           ByVal pEmail As String, _
                                           ByVal pTelefone As String, _
                                           ByVal pObservacoes As String, _
                                           ByVal pNome As String, _
                                           ByVal pErrosClasse4 As String, tagDynaTrace As String) As Boolean
    Dim lNosNumero As String
    Dim lNovReferencia As String
    Dim lObjstcPed As stcA46V702B
    Dim lobjstcTransm As stcA28V721B
    Dim lobjTransm As clsA28V721A
    Dim lUsuario As String
    Dim lCodCorretor As Long
    Dim lArqIni As String
    Dim lPropostas(1) As String         'Matriz com propostas selecionadas (neste caso, apenas 1).
    Dim lArqTransmissao As String
    Dim lMensageErro As String
    Dim clsA46V101A As clsA46V101A
    Dim lColecao As Collection

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfSolicitarClasse4OFFLine"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro
    If Not mfPosicionaUsuario(tagDynaTrace) Then GoTo SAIDA

    Set lObjstcPed = gfProcuraNossoNumero(pNossoNumero, False, "", tagDynaTrace)

    If lObjstcPed Is Nothing Then
        UltimoErro = "Cota��o n�o localizada " & pNossoNumero
        GoTo SAIDA
    End If
    'Quando um c�lculo � visualizado ent�o pode ser que o corretor tenha impresso a cota��o (dat_pri_calculo <>0)
    'Nestes casos n�o poderemos realizar o classe 4 com o mesmo nosso_numero  porque se n�o pode ser que a cota��o
    'seja declinada e ele apresente a primeira impress�o para ser emitida, desta forma ser� recusado indevidamente
    'na emiss�o.

    If lObjstcPed.Dat_Pri_Calculo <> 0 Then
        Set lObjstcPed = gfCopiarProposta(pNossoNumero, 2, tagDynaTrace)
        If M46V999.gfPreenchido(UltimoErro) Then GoTo SAIDA
        lNosNumero = lObjstcPed.NOSSO_NUMERO
    Else
        lNosNumero = pNossoNumero
    End If

    On Error GoTo ControleErro

    Set lobjstcTransm = New stcA28V721B
    Set lobjTransm = New clsA28V721A
    Set lobjTransm.M46V999 = M46V999

    lobjstcTransm.NOSSO_NUMERO = lNosNumero
    lobjstcTransm.STATUS = -1

    If Not lobjTransm.gfCarregarColecao(lobjstcTransm, objstcUsuario.PropTabela) Then
        UltimoErro = lobjTransm.UltimoErro
        GoTo SAIDA
    End If

    If lobjTransm.Colecao.Count > 0 Then
        Set lobjstcTransm = lobjTransm.Colecao.Item(1)
    End If

    With lobjstcTransm
        .CCotacao = ""
        .Dat_Cotacao = Left(M46V999.gfForData(3), 8)
        .Email = pEmail
        .Nome = pNome
        .Telefone = pTelefone
        .NOSSO_NUMERO = lNosNumero
        .STATUS = 0
        .Usuario = objstcUsuario.login_user_rede
    End With

    If lobjTransm.Colecao.Count > 0 Then
        If Not lobjTransm.gfAlterar(lobjstcTransm, objstcUsuario.PropTabela) Then
            UltimoErro = lobjTransm.UltimoErro
            GoTo SAIDA
        End If
    Else
        If Not lobjTransm.gfIncluir(lobjstcTransm, objstcUsuario.PropTabela) Then
            UltimoErro = lobjTransm.UltimoErro
            GoTo SAIDA
        End If
    End If
    lPropostas(1) = lNosNumero
    If objstcUsuario.codCorretor <> 0 Then
        lCodCorretor = objstcUsuario.codCorretor
    Else
        lCodCorretor = 9999999
    End If

    lMensageErro = "** Erros classe 4: " & Replace(pErrosClasse4, vbCrLf, "%!*")
    lMensageErro = lMensageErro & " **  Observa��o do corretor:" & Replace(M46V999.gfRetirarCaracterEspecial(pObservacoes), vbCrLf, "%!*")

    Set clsA46V101A = New clsA46V101A
    Set clsA46V101A.M46V999 = M46V999

    lUsuario = mUsuarioTransmissao

    Set lColecao = New Collection

    lColecao.Add lObjstcPed
    If clsA46V101A.gfCriArqCotacao(lCodCorretor, lUsuario, _
                                   lColecao, pEmail, pNome, pTelefone, lobjstcTransm.CCotacaoOrig, lArqTransmissao, UltimoErro, lMensageErro) = False Then
        GoTo SAIDA
    End If

    mfSolicitarClasse4OFFLine = True

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function
ControleErro:
    If Not M46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA
FechaObj:
    Set lObjstcPed = Nothing
    Set lobjstcTransm = Nothing
    Set lobjTransm = Nothing
    Set clsA46V101A = Nothing
    Set lColecao = Nothing
    Return
End Function

Private Function mfIncluirSiscotaMulti(pNossoNumero As String, pNum_SiscotaAnt As String, pobjstcTransm As stcA28V721B, pObservacao As String, pRamo As Integer, pClasse As String, pMensagem As String, Optional ByRef strNumCotacao As String) As Boolean
    Dim lExec As String
    Dim lCCotacao As String
    Dim lRegistro As String
    Dim lConexaoProd As ADODB.Connection
    Dim lrst As clsYasRecordSet
    Dim lSenha As String
    Dim lAux As Integer

    If Not M46V999.gfAbrBasDados("RAMOSDIVERSOS", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, lConexaoProd, pMensagem) Then
        GoTo SAIDA
    End If

    'Chama a procedure que ir� retornar o n�mero da cota��o.
    lExec = " exec siscota.dbo.proc_CotacaoMultiSyasNovo " & M46V999.gfFormatarCamposSQL(pNossoNumero) & "," & _
            val(pNum_SiscotaAnt) & " ," & M46V999.gfFormatarCamposSQL(pobjstcTransm.Nome) & "," & _
            M46V999.gfFormatarCamposSQL(pobjstcTransm.Email) & "," & _
            M46V999.gfFormatarCamposSQL(pobjstcTransm.Telefone) & ", '" & M46V999.gPathSiscota & "', " & _
            "'" & Replace(pObservacao, "%!*", vbCrLf) & "'"
    If Not M46V999.gfObtRegistro(lConexaoProd, lExec, lrst, pMensagem) Then
        GoTo SAIDA
    End If

    If Not lrst.EOF Then
        lCCotacao = lrst("Cotacao")
    End If

    'Produ��o
    lExec = "UPDATE RAMOSDIVERSOS.DBO.TAB_CTRL_EMIS SET TIP_COTAC = 2 , NUM_COTAC = '" & lCCotacao & "' WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pNossoNumero)
    If M46V999.gfExeSQL(lConexaoProd, lExec, pMensagem) <> 0 Then
        GoTo SAIDA
    End If

    'Produ��o
    lExec = "Insert into siscota.dbo.tab_syas_cotacao (CCOTACAO,NOSSO_NUMERO, " & _
            " DATA_INCLUSAO, COD_RAMO, cod_status, STATUS_ENVIO, " & _
            "COD_CORRETOR, M99V001.Cod_Usuario, NomArq, SENHA ) values ( " & _
            M46V999.gfFormatarCamposSQL(lCCotacao) & "," & _
            M46V999.gfFormatarCamposSQL(pNossoNumero) & "," & _
            M46V999.gfForData(3) & "," & pRamo & _
            ", 1, 0," & _
            objstcUsuario.codCorretor & "," & 0 & " ,'','')"

    If M46V999.gfExeSQL(lConexaoProd, lExec, pMensagem) <> 0 Then
        GoTo SAIDA
    End If

    'Atualizo o n�mero do siscota na base P0045400 (SQL 2008 - SyasPortal)
    If Not M46V999.gfAbrBasDados("P0045400", M46V999.gTipP0045400, M46V999.gEndP0045400, M46V999.mConexaoSQL, pMensagem) Then
        GoTo SAIDA
    End If
    
    'Projeto Classe 3 web
    lExec = "UPDATE RAMOSDIVERSOS.DBO.TAB_CTRL_EMIS_COMPL SET TIP_CLASSE = " & pClasse & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pNossoNumero)
    If M46V999.gfExeSQL(lConexaoProd, lExec, pMensagem) <> 0 Then
        GoTo SAIDA
    End If
    
    lExec = "UPDATE SYASPORTAL.DBO.TAB_CTRL_EMIS_COMPL SET TIP_CLASSE = " & pClasse & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pNossoNumero)
    If M46V999.gfExeSQL(lConexaoProd, lExec, pMensagem) <> 0 Then
        GoTo SAIDA
    End If

    lExec = "Update " & objstcUsuario.PropTabela & "P0045400 set ccotacao = " & M46V999.gfFormatarCamposSQL(lCCotacao) & " , status = 1 where nosso_numero = " & M46V999.gfFormatarCamposSQL(pNossoNumero)
    If M46V999.gfExeSQL(M46V999.mConexaoSQL, lExec, pMensagem) <> 0 Then
        GoTo SAIDA
    End If
    
    '' 2018-01-16
    '' Copiar arquivo do TEMP web para SISCOTA
    Call SiscotaGravarAnexo(pNossoNumero, lCCotacao)
    
    strNumCotacao = lCCotacao
    
    mfIncluirSiscotaMulti = True

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

FechaObj:
    Set lConexaoProd = Nothing
    Set lrst = Nothing
    Return
End Function

Private Sub SiscotaGravarAnexo(pNossoNumero As String, lCCotacao As String)
    Dim i As Integer, siscotaName As String
    Dim myFSO As New FileSystemObject, myFile
    Dim objAnexoA As New clsA46V752A
    Dim objAnexoB As New stcA46V752B
    
    On Error GoTo Fim
    
    Set objAnexoA.M46V999 = M46V999
    '' M46V999.gPathSiscota
    If objAnexoA.mfCarregarColecaoByNossoNumero(pNossoNumero) Then
        For Each objAnexoB In objAnexoA.Colecao
           If myFSO.FileExists(objAnexoB.PastaArquivo & "\" & objAnexoB.NomeArquivoFisico) Then
                Set myFile = myFSO.GetFile(objAnexoB.PastaArquivo & "\" & objAnexoB.NomeArquivoFisico)
                
                If Not myFile Is Nothing Then
                    siscotaName = "P" & lCCotacao & CStr(objAnexoB.Sequencia) & "." & SiscotaArqExt(objAnexoB.NomeArquivoFisico)
                    myFile.Copy (M46V999.gPathSiscota & siscotaName)
                End If
            End If
        Next objAnexoB
    End If

Fim:
    Set objAnexoA = Nothing
End Sub

Private Function SiscotaArqExt(NomeArq As String) As String
    Dim arr() As String ' array tipo arq
    Dim max As Integer
    Dim strExt As String
    strExt = "---"
    
    On Error GoTo Fim
    
    arr = Split(NomeArq, ".")
    max = UBound(arr)
    strExt = arr(max)
    
Fim:
    SiscotaArqExt = strExt
End Function

Private Function mfIncluirSiscotaMultiRiscosNomeados(pNossoNumero As String, pNum_SiscotaAnt As String, pobjstcTransm As stcA28V721B, pObservacao As String, pRamo As Integer, pClasse As String, pMensagem As String) As Boolean
    Dim lExec As String
    Dim lCCotacao As String
    Dim lRegistro As String
    Dim lConexaoProd As ADODB.Connection
    Dim lrst As clsYasRecordSet
    Dim lSenha As String
    Dim lAux As Integer
    Dim lCodRamoRiscosNomeados As Integer

    If Not M46V999.gfAbrBasDados("RAMOSDIVERSOS", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, lConexaoProd, pMensagem) Then
        GoTo SAIDA
    End If

    'Chama a procedure que ir� retornar o n�mero da cota��o.
    lExec = " exec siscota.dbo.proc_CotacaoMultiSyasRiscosNomeados " & M46V999.gfFormatarCamposSQL(pNossoNumero) & "," & _
            val(pNum_SiscotaAnt) & " ," & M46V999.gfFormatarCamposSQL(pobjstcTransm.Nome) & "," & _
            M46V999.gfFormatarCamposSQL(pobjstcTransm.Email) & "," & _
            M46V999.gfFormatarCamposSQL(pobjstcTransm.Telefone) & ", '" & M46V999.gPathSiscota & "', " & _
            "'" & Replace(pObservacao, "%!*", vbCrLf) & "'"
    If Not M46V999.gfObtRegistro(lConexaoProd, lExec, lrst, pMensagem) Then
        GoTo SAIDA
    End If

    If Not lrst.EOF Then
        lCCotacao = lrst("Cotacao")
    End If

    'Produ��o
    lExec = "UPDATE SYASPORTAL.DBO.TAB_CTRL_EMIS SET TIP_COTAC = 2 , NUM_COTAC = '" & lCCotacao & "' WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pNossoNumero)
    If M46V999.gfExeSQL(lConexaoProd, lExec, pMensagem) <> 0 Then
        GoTo SAIDA
    End If

    'Produ��o
    lCodRamoRiscosNomeados = 110
    lExec = "Insert into siscota.dbo.tab_syas_cotacao (CCOTACAO,NOSSO_NUMERO, " & _
            " DATA_INCLUSAO, COD_RAMO, cod_status, STATUS_ENVIO, " & _
            "COD_CORRETOR, M99V001.Cod_Usuario, NomArq, SENHA ) values ( " & _
            M46V999.gfFormatarCamposSQL(lCCotacao) & "," & _
            M46V999.gfFormatarCamposSQL(pNossoNumero) & "," & _
            M46V999.gfForData(3) & "," & lCodRamoRiscosNomeados & _
            ", 1, 0," & _
            objstcUsuario.codCorretor & "," & 0 & " ,'','')"

    If M46V999.gfExeSQL(lConexaoProd, lExec, pMensagem) <> 0 Then
        GoTo SAIDA
    End If
    
    'Atualizo o n�mero do siscota na base P0045400 (SQL 2008 - SyasPortal)
    If Not M46V999.gfAbrBasDados("P0045400", M46V999.gTipP0045400, M46V999.gEndP0045400, M46V999.mConexaoSQL, pMensagem) Then
        GoTo SAIDA
    End If
    
    'Projeto Classe 3 web
    lExec = "UPDATE RAMOSDIVERSOS.DBO.TAB_CTRL_EMIS_COMPL SET TIP_CLASSE = " & pClasse & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pNossoNumero)
    If M46V999.gfExeSQL(lConexaoProd, lExec, pMensagem) <> 0 Then
        GoTo SAIDA
    End If
    
    lExec = "UPDATE SYASPORTAL.DBO.TAB_CTRL_EMIS_COMPL SET TIP_CLASSE = " & pClasse & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pNossoNumero)
    If M46V999.gfExeSQL(lConexaoProd, lExec, pMensagem) <> 0 Then
        GoTo SAIDA
    End If

    lExec = "Update " & objstcUsuario.PropTabela & "P0045400 set ccotacao = " & M46V999.gfFormatarCamposSQL(lCCotacao) & " , status = 1 where nosso_numero = " & M46V999.gfFormatarCamposSQL(pNossoNumero)
    If M46V999.gfExeSQL(M46V999.mConexaoSQL, lExec, pMensagem) <> 0 Then
        GoTo SAIDA
    End If

    mfIncluirSiscotaMultiRiscosNomeados = True

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

FechaObj:
    Set lConexaoProd = Nothing
    Set lrst = Nothing
    Return
End Function

Private Sub mpVerNormas(ByRef pobjstcPedLoc As stcA46V708B, ByVal pCod_Aceit As Integer, ByRef pMensInsp As String)
    Dim lstcNorma As stcA46V708N
    Dim lstrItem As String

    Dim lretorno As String
    Dim lI As Integer
    Dim lCod As Integer

    Dim lMensagem As String

    pMensInsp = ""
    Call M46V999.gpGraLog(0, "VerNormas:" & M46V999.VerificarNormas)
    If Not M46V999.VerificarNormas Then
        GoTo SAIDA
    End If
    lstrItem = "(Item: " & Format(pobjstcPedLoc.Num_Item, "000") & ")"
    Call M46V999.gpGraLog(0, "VerNormas")
    For Each lstcNorma In pobjstcPedLoc.ColecaoNormas
        lCod = CInt(Left(CStr(lstcNorma.COD_STAT_IRREG), 2))
        Call M46V999.gpGraLog(0, lCod)
        'Se na tabela de parametriza��o retornou classe 1 despreza retornos 10 e 80
        'Call M46V999.gpGraLog(0, "Simony4 Item: " & pobjstcPedLoc.Num_Item & " Codigo " & lCod & " Aceita��o " & pCod_Aceit)
        If Not (pCod_Aceit = 1 And lCod <> 55) Then
            If lCod = 55 Then
                pMensInsp = lstcNorma.DSC_ERR_NORMA_EXT
            Else
                If pobjstcPedLoc.ObjPai.StatusCotSiscota = enum_StatusTransm_Aprovada Or pobjstcPedLoc.ObjPai.StatusCotSiscota = enum_StatusTransm_EnvioKit Then
                    If mfValidaCoberturaClasse4(pobjstcPedLoc) Then
                        Call M46V999.gpMenErro("Normas " & lstcNorma.COD_ERR_NORMA & " - " & lstcNorma.DSC_ERR_NORMA_EXT & lstrItem, lstcNorma.COD_CLASSE_NORMA)
                    End If
                Else
                    'If pobjstcPedLoc.ObjPai.StatusCotSiscota <> enum_StatusTransm_Aprovada And pobjstcPedLoc.ObjPai.StatusCotSiscota <> enum_StatusTransm_EnvioKit Then
                    Call M46V999.gpMenErro("Normas " & lstcNorma.COD_ERR_NORMA & " - " & lstcNorma.DSC_ERR_NORMA_EXT & lstrItem, lstcNorma.COD_CLASSE_NORMA)
                End If
            End If
        End If
        'Call M46V999.gpGraLog(0, "Simony5 Mensagem: " & pMensInsp)
    Next lstcNorma

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lstcNorma = Nothing
    Return
End Sub

Public Function mfValidaCoberturaClasse4(ByRef pobjstcPedLoc As stcA46V708B) As Boolean
    Dim lstcCobertura As stcA46V704B
    Dim lretorno As Boolean

    lretorno = False

    For Each lstcCobertura In pobjstcPedLoc.ColecaoPedCobert
        If lstcCobertura.Tip_Origem_C4 = 0 Then
            lretorno = True
            Exit For
        End If
    Next

    mfValidaCoberturaClasse4 = lretorno

End Function

Public Function gNewCollection() As Collection
    Set gNewCollection = New Collection
End Function

Private Function mfNormalizaPedido(ByRef pStcPed As stcA46V702B, ByRef pMsgErr, tagDynaTrace As String) As Boolean
    Dim lMsrErr As String
    Dim lContX As Integer
    Dim lContY As Integer
    Dim lBonus As Double
    Dim lobjPedLoc As clsA46V708A
    Dim lCountZ As Integer
    Dim lNumApol As Double
    Dim lDatMov As Long

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfNormalizaPedido"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear
    On Error GoTo TrataErro: mfNormalizaPedido = False

    If pStcPed Is Nothing Then
        lMsrErr = "Parametro de entrada stcA46V702B n�o instanciado!"
        GoTo TrataErro
    End If

    'Call mpImprimeObjeto(pStcPed, tagDynaTrace)

    'Variaveis Globais!!!
    M46V999.gCodProduto = pStcPed.Cod_Produto
    M46V999.gCodSequencia = pStcPed.Cod_Sequencia
    M46V999.gCod_Ramo = pStcPed.Cod_Ramo
    M46V111.gCod_plano = IIf(pStcPed.Cod_Produto < 1080, pStcPed.Cod_Estip, M46V999.gfPlanoBasico(pStcPed.Cod_Ramo, pStcPed.Cod_Estip))
    '============

    Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Inicio mpAtribuiClasseModulo", 0)
    Call mpAtribuiClasseModulo(pStcPed, tagDynaTrace)
    Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Fim mpAtribuiClasseModulo", 0)

    '- Pedido
    Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Inicio da Normaliza��o do pedido", 0)
    With pStcPed
        .Cod_Pr_Min_Endo = 1
        .Cod_Pr_Min_Parc = 1
        'If M46V999.gProPreMinParcela < M46V999.gProPreMinParcelaTAB Then
        '    .Cod_Pr_Min_Parc = 2
        'End If
        'If M46V999.gCod_Ramo = 113 Then
        'Ficha 901507
        '    If M46V111.gCod_Plano = 1 Then
        '        Select Case frmT46V102A.txtCorretor(1).Tag
        '            Case 8001, 8002, 8003, 8004, 8005, 8006, 8008, 8011, 8014, 8028
        '                .Cod_Pr_Min_Parc = 2
        '        End Select
        '    End If
        'End If
        If objstcUsuario.TipoAcesso = DIGITACAO Then
            .Cod_Pr_Min_Parc = 2
        End If
        '.Val_Pr_Antec = 0
        .DAT_DIGIT = Left(M46V999.gfForData(3), 8)
        .HOR_DIGIT = Mid(M46V999.gfForData(3), 9, 4)
        .Cod_User_Liber = Left$(objstcUsuario.login_user_rede, 6)
        .Qtd_Copia = 3
        If Mid(.Num_Cgc_Cpf, 1, 8) = "59291534" Then
            If .Cod_Ramo = 112 Or .Cod_Ramo = 150 Or .Cod_Ramo = 710 Then
                .Qtd_Copia = 4
            End If
        End If

        '11.14 Igreja Universal sem cobran�a de IOF, 11.18 modificado p todos CNPJs
        If InStr(1, .Num_Cgc_Cpf, "29744778") > 0 And M46V999.gCod_Corretor = 860301 Then
            Select Case .Cod_Reg
                Case 1
                    .Cod_Reg = 3
                Case 4
                    .Cod_Reg = 2
                Case 101
                    .Cod_Reg = 103
                Case 104
                    .Cod_Reg = 102
            End Select
        End If

        If .Cod_Reg <> 3 And .Cod_Reg <> 103 And .Cod_Reg <> 106 And .Cod_Reg <> 107 And .Cod_Reg <> 120 And .Cod_Reg <> 2 And .Cod_Reg <> 102 And .Cod_Reg <> 129 Then
            If pStcPed.ColecaoPedLoc(1).Tip_Emissao >= 300 Then
                If .Dat_Ini_Vig >= 20130101 Then
                    .Cod_Reg = 104
                Else
                    .Cod_Reg = 101
                End If
            Else
                If .Dat_Ini_Vig >= 20130101 Then
                    .Cod_Reg = 4
                Else
                    .Cod_Reg = 1
                End If
            End If
        End If

        If .Cod_Reg = 120 Or .Cod_Reg = 106 Or .Cod_Reg = 107 Or .Cod_Reg = 129 Then
            .Tip_Cust_Emis = 1
            .Val_Pr_Antec = 0
        Else
            If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 2 Then     '30 dias
                .Tip_Cust_Emis = 1
            ElseIf pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 And (pStcPed.Num_Cta_Bas_Cor = 0 Or pStcPed.objstcCtrlEmis.Cod_Bco_Boleto > 0) Then   'boleto
                .Tip_Cust_Emis = 5
            ElseIf pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 And pStcPed.Num_Cta_Bas_Cor <> 0 Then    'ant com d�bito
                .Tip_Cust_Emis = 3
            End If
        End If
        '        Else    'n�o
        '            If frmT46V102A.optPagAntecipado(2).Value = True Then  '30 dias
        '                .Tip_Cust_Emis = 0
        '            ElseIf frmT46V102A.optPagAntecipado(1).Value = True And frmT46V102A.optForPagamento(0).Value = True Then    'boleto
        '                .Tip_Cust_Emis = 4
        '            ElseIf frmT46V102A.optPagAntecipado(1).Value = True And frmT46V102A.optForPagamento(1).Value = True Then    'ant com debito
        '                .Tip_Cust_Emis = 2
        '            End If
        '        End If
        .NUM_PI = Right(Year(Date), 2) & Format(pStcPed.Num_Ped, "00000000")
        .LMI_Unico = IIf(.LMI_Unico = 1, .LMI_Unico, 2)
        .Val_IS = .ColecaoPedLoc(1).gISBasica
        .Desc_Plano = .getPlano.NomPlano
        '' 20170524: Juliano
        .Ind_Oculta_Web = .getPlano.IndOcultaWeb
        .Cod_Iden_Praz = Me.gfTipoPrazo(CDbl(.Dat_Ini_Vig), CDbl(.Dat_Ter_Vig))
        '.Cod_Bco_Cobr = IIf(.Cod_Bco_Cobr = 0, 399, .Cod_Bco_Cobr)

'        Boleto santander
        lDatMov = 0
        If .Cod_Tip_Doc = e_TipDoc_Proposta And _
           .objstcCtrlEmis.Cod_Bco_Boleto > 0 And _
           (.Cod_Bco_Cobr = 0 Or .Cod_Bco_Cobr = e_Banco_Santander Or .Cod_Bco_Cobr = e_Banco_HSBC) And _
           .objstcCtrlEmis.Cod_Bco_Boleto <> 33 And _
           .objstcCtrlEmis.Cod_Bco_Boleto <> 399 Then
            If pStcPed.ColecaoPedLoc(1).Tip_Emissao = 300 Then
                lNumApol = pStcPed.gStcPedidoApolice.Num_Apol
            Else
                lNumApol = 0
            End If
            
            If pStcPed.ColecaoPedLoc(1).Tip_Emissao <= 200 Then
                lDatMov = 0
            Else
                lDatMov = pStcPed.gStcPedidoApolice.Dat_Emis
            End If
            
            If M46V999.mfObtemBancoBoleto(pStcPed.ColecaoPedLoc(1).Tip_Emissao, _
                                          pStcPed.Num_Cgc_Cpf, _
                                          pStcPed.Cod_Produto, _
                                          pStcPed.Cod_Segurado, _
                                          lNumApol, _
                                          pStcPed.objstcCtrlEmis.Cod_Corr, _
                                          lDatMov) Then
                .Cod_Bco_Cobr = 33
            Else
                .Cod_Bco_Cobr = 399
            End If
            
            '2016-08-22: Alterado Itau p Santander
            If .objstcCtrlEmis.Cod_Corr = 15701 Then ' Mercedes
                .Cod_Bco_Cobr = 33  '341
            End If
            '2016-08-22
        End If

        'Endere�o de Cobran�a
        If .gEnd_Cob_Diferente = False Then
            .Num_CEP = .objstcPedEndereco.Num_CEP
            .Cod_Tip_Logr = .objstcPedEndereco.Cod_Tip_Logr
            .Nom_Logr = .objstcPedEndereco.Nom_Logr
            .Dsc_Compl_Logr = .objstcPedEndereco.Dsc_Compl_Logr
            .Nom_Bairro = .objstcPedEndereco.Nom_Bairro
            .Nom_Cid = .objstcPedEndereco.Nom_Cid
            .Sig_UF = .objstcPedEndereco.Sig_UF
            .Num_Fone = .objstcPedEndereco.Num_Fone
        End If
        .Qtd_Anexo = .ColecaoPedLoc.Count
    End With
    Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Fim da Normaliza��o do pedido", 0)

    '- Dependentes do Pedido
    'Objeto da classe clsA46V712A (TAB_CTRL_EMIS)
    Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Inicio da Normaliza��o dos dependentes do pedido", 0)
    With pStcPed.objstcCtrlEmis
        .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
        .NUM_PI = pStcPed.NUM_PI
        .COD_UNID = pStcPed.Cod_Unid_Prod
        .Cod_Ramo = pStcPed.Cod_Ramo
        .Dat_Ini_Vig = pStcPed.Dat_Ini_Vig
        .Dat_Ter_Vig = pStcPed.Dat_Ter_Vig
        .Cod_Corr = 0
        .Per_Corr = 0
        .Cod_Corr_2 = 0
        .Per_Corr_2 = 0
        .Cod_Corr_3 = 0
        .Per_Corr_3 = 0
        For lContX = 1 To pStcPed.ColecaoPedCorr.Count
            If pStcPed.ColecaoPedCorr(lContX).Num_Ordem_Corr = 1 Then
                .Cod_Corr = pStcPed.ColecaoPedCorr(lContX).Cod_Corr
                .Per_Corr = pStcPed.ColecaoPedCorr(lContX).Per_Corr
            ElseIf pStcPed.ColecaoPedCorr(lContX).Num_Ordem_Corr = 2 Then
                .Cod_Corr_2 = pStcPed.ColecaoPedCorr(lContX).Cod_Corr
                .Per_Corr_2 = pStcPed.ColecaoPedCorr(lContX).Per_Corr
            ElseIf pStcPed.ColecaoPedCorr(lContX).Num_Ordem_Corr = 3 Then
                .Cod_Corr_3 = pStcPed.ColecaoPedCorr(lContX).Cod_Corr
                .Per_Corr_3 = pStcPed.ColecaoPedCorr(lContX).Per_Corr
            End If
        Next
        .Nom_Seg = pStcPed.Nom_Segurado
        .Qtd_Anex = pStcPed.Qtd_Anexo
        .Num_Apol_Ant = pStcPed.Num_Apol_An
        .Per_Juros_Neg = pStcPed.Per_Juro_Frac
        .Cod_Envi_Doc = "X"
        .DAT_DIGIT = pStcPed.DAT_DIGIT
        .HOR_DIGIT = pStcPed.HOR_DIGIT
        .Cod_User_Digit = Me.objstcUsuario.login_user_rede
        If .Dat_Digit_MF = 0 Then .Dat_Digit_MF = M46V999.gHoje
        .Hor_Digit_MF = pStcPed.HOR_DIGIT & "00"
        .Val_Pr_Tot = Format$(pStcPed.gValorTotal, "standard")
        .Num_Prop = pStcPed.Num_Ped
        Select Case pStcPed.ColecaoPedLoc.Item(1).Tip_Emissao
        Case e_TipEmissao_SeguroNovo, e_TipEmissao_RenCongenere
            .Cod_Tip_Doc = 1
        Case e_TipEmissao_RenYasuda
            .Cod_Tip_Doc = 3
        Case Else
            .Cod_Tip_Doc = 2
        End Select
        .Val_Pr_Emis = pStcPed.VAL_PR
        .Val_Pr_Liq = pStcPed.VAL_PR
        .Per_Tot_Coss = 0
        For lContX = 1 To pStcPed.ColecaoPedCoss.Count
            .Per_Tot_Coss = .Per_Tot_Coss + pStcPed.ColecaoPedCoss(lContX).Per_Coss
        Next
        .Qtd_Parc = pStcPed.Qtd_Parc
        .Cod_Prod = pStcPed.Cod_Prod
        .Val_IS = pStcPed.Val_IS
        .Num_Cgc_Cpf = pStcPed.Num_Cgc_Cpf
        .Cod_Iden_Pessoa = pStcPed.Tip_Pessoa
        .Num_Cotac = IIf(M46V999.gfPreenchido(.Num_Cotac), .Num_Cotac, "0")
        .Num_Ref_Corr = IIf(M46V999.gfPreenchido(.Num_Ref_Corr), .Num_Ref_Corr, "0")

        .Ind_Agenciamento = IIf(pStcPed.ColecaoAgenciador.Count > 0, 1, 0)
        .Ind_Prolabore = IIf(pStcPed.ColecaoProlabore.Count > 0, 1, 0)
        
        'Dados Maritima
        .cod_origem = .cod_origem
        .Num_Apol_Mar = .Num_Apol_Mar
        .Ind_Cond_Constituido = .Ind_Cond_Constituido
    End With

    'Objeto da classe clsA46V707A (TAB_PED_ENDER)
    With pStcPed.objstcPedEndereco
        .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
        .Num_Ped = pStcPed.Num_Ped
        .Cod_Org = pStcPed.Cod_Org
        .Cod_Ramo = pStcPed.Cod_Ramo
        .NUM_PI = pStcPed.NUM_PI
    End With

    For lContX = 1 To pStcPed.ColecaoPedCorr.Count
        With pStcPed.ColecaoPedCorr(lContX)
            Set .ObjPai = pStcPed
            .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
            .Num_Ped = pStcPed.Num_Ped
            .Cod_Org = pStcPed.Cod_Org
            .Cod_Ramo = pStcPed.Cod_Ramo
            .NUM_PI = pStcPed.NUM_PI
        End With
    Next

    Dim lSeq As Integer
    lSeq = 0
    For lContX = 1 To pStcPed.ColecaoProlabore.Count
        With pStcPed.ColecaoProlabore(lContX)
            Set .ObjPai = pStcPed
            .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
            .NUM_PI = pStcPed.NUM_PI
            .Tip_Canal = 2
            lSeq = lSeq + 1
            .NUM_SEQ = lSeq
        End With
    Next

    lSeq = 0
    For lContX = 1 To pStcPed.ColecaoAgenciador.Count
        With pStcPed.ColecaoAgenciador(lContX)
            Set .ObjPai = pStcPed
            .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
            .NUM_PI = pStcPed.NUM_PI
            .Tip_Canal = 3
            lSeq = lSeq + 1
            .NUM_SEQ = lSeq
        End With
    Next

    For lContX = 1 To pStcPed.ColecaoTelefones.Count
        With pStcPed.ColecaoTelefones(lContX)
            Set .ObjPai = pStcPed
            .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
            .NUM_PI = pStcPed.NUM_PI
        End With
    Next


    For lContX = 1 To pStcPed.ColecaoPedCoss.Count
        With pStcPed.ColecaoPedCoss(lContX)
            Set .ObjPai = pStcPed
            .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
            .Num_Ped = pStcPed.Num_Ped
            .Cod_Org = pStcPed.Cod_Org
            .Cod_Ramo = pStcPed.Cod_Ramo
            .NUM_PI = pStcPed.NUM_PI
        End With
    Next

    For lContX = 1 To pStcPed.ColecaoPedJuros.Count
        Set pStcPed.ColecaoPedJuros(lContX).ObjPai = pStcPed
        pStcPed.ColecaoPedJuros(lContX).NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
    Next

    For lContX = 1 To pStcPed.ColecaoPedVenc.Count
        With pStcPed.ColecaoPedVenc(lContX)
            Set .ObjPai = pStcPed
            .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
            .Num_Ped = pStcPed.Num_Ped
            .Cod_Org = pStcPed.Cod_Org
            .Cod_Ramo = pStcPed.Cod_Ramo
        End With
    Next

    For lContX = 1 To pStcPed.ColecaoTextoEspec.Count
        With pStcPed.ColecaoTextoEspec(lContX)
            Set .ObjPai = pStcPed
            .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
            .Num_Ped = pStcPed.Num_Ped
            .Cod_Org = pStcPed.Cod_Org
            .Cod_Ramo = pStcPed.Cod_Ramo
            .NUM_PI = pStcPed.NUM_PI
            .Qtd_vias = 1
        End With
    Next
    
    For lContX = 1 To pStcPed.ColecaoServicoAss24h.Count
        With pStcPed.ColecaoServicoAss24h(lContX)
            Set .ObjPai = pStcPed
                .Cod_Ramo = pStcPed.Cod_Ramo
        End With
    Next
    Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Fim da Normaliza��o dependentes do pedido", 0)

    '=====Itens
    Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Inicio da Normaliza��o dos itens do pedido", 0)
    For lContX = 1 To pStcPed.ColecaoPedLoc.Count
        '- Item
        With pStcPed.ColecaoPedLoc(lContX)
            Set .ObjPai = pStcPed
            .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
            .Num_Ped = pStcPed.Num_Ped
            .Cod_Org = pStcPed.Cod_Org
            .Cod_Ramo = pStcPed.Cod_Ramo
            .Cod_User_Digit = pStcPed.Cod_User_Digit
            .DAT_DIGIT = pStcPed.DAT_DIGIT
            .HOR_DIGIT = pStcPed.HOR_DIGIT
            .Cod_User_Liber = pStcPed.Cod_User_Liber
            .Dat_Liber = pStcPed.Dat_Liber
            .Hor_Liber = pStcPed.Hor_Liber
            .Cod_Sit_item = e_SitItem_NaoCalc
            .Cod_Iden_Tarif = pStcPed.Cod_Produto
            .NUM_PI = pStcPed.NUM_PI
            .Cod_Produto = pStcPed.Cod_Estip
            .Cod_Plano = .Cod_Produto
            .Tab_Clasf = .Cod_Produto
            If pStcPed.objstcCtrlEmis.cod_origem = 2 Then ' renova��o maritima
                .Tip_Emissao = e_TipEmissao_RenCongenere
                .Num_Apol_Anterior = 0
                .Ren_Con_Seguradora = 572
                .Ren_Con_Apolice = pStcPed.objstcCtrlEmis.Num_Apol_Mar
                .Ren_Con_Item = .Num_Item
            End If
            
            If .Tip_Emissao = 200 And .BonusAlterado = False Then
                .Ind_Sinistro = .Apol_Ant_IND_SINISTRO
                If .Ind_Sinistro = 0 Then .Ind_Sinistro = 2
                lBonus = 0
                If .mBonus_Anterior = 0 Then
                    Set lobjPedLoc = New clsA46V708A
                    Set lobjPedLoc.M46V999 = M46V999
                    lBonus = lobjPedLoc.mfRetornaBonusApolice(pStcPed.ColecaoPedLoc(lContX), tagDynaTrace)
                Else
                    lBonus = .mBonus_Anterior
                End If
                If .Ind_Sinistro = e_SimNao_Sim Then
                    lBonus = lBonus - 5
                Else
                    lBonus = lBonus + 5
                End If
                If lBonus < 0 Then lBonus = 0

                If .Cod_Ramo = e_CodRamo_Empresarial Then
                    If pStcPed.Cod_Produto >= 942 Then
                        'Soma mais 5 da fidelidade.
                        lBonus = lBonus + 5
                    End If
                    'If lBonus > 30 Then lBonus = 30
                    If lBonus > 25 Then lBonus = 25
                ElseIf .Cod_Ramo = e_CodRamo_Residencial Then
                    If pStcPed.Cod_Produto >= 946 Then
                        'Soma mais 5 da fidelidade.
                        If pStcPed.Cod_Produto < 1126 Then
                                lBonus = lBonus + 5
                        End If
                        'If lBonus > 30 Then lBonus = 30
                        If lBonus > 25 Then lBonus = 25
                    Else
                        If lBonus > 25 Then lBonus = 25
                    End If
                Else
                    If lBonus > 25 Then lBonus = 25
                End If

                'piloto maritima
                If pStcPed.Cod_Produto >= 1080 Then
                    If lBonus > 25 Then lBonus = 25
                End If

                If .Cod_Plano >= 2001 And .Cod_Plano <= 2031 Then
                    lBonus = 0
                End If

                'Ficha 901673
                If (.Cod_Plano >= 2051 And .Cod_Plano <= 2100) Or (.Cod_Plano >= 2111 And .Cod_Plano <= 2120) Then
                    lBonus = 0
                End If
                .Per_Bonus = lBonus
            End If

            .Val_LMG_Lc = .gISTotal_LC
            
            If (.Cod_Plano >= 1000) Then
                .Val_Vr_Lc = .gISTotal_LC
            End If

            ' CR 1425 - Desconto de experi�ncia n�o deve ser aplicado em planos
            If .Cod_Plano >= 2001 _
            And .Cod_Plano <> 2300 _
            And pStcPed.getPlano.EhPlanoSegmentado = False _
            Then
                lBonus = 0
                .Per_Bonus = 0
            End If
            'CR1425 - Fim


        End With


        '- Inclui os dependentes do Item
        For lContY = 1 To pStcPed.ColecaoPedLoc(lContX).ColecaoPedCobert.Count
            With pStcPed.ColecaoPedLoc(lContX).ColecaoPedCobert(lContY)
                Set .ObjPai = pStcPed.ColecaoPedLoc(lContX)
                .Num_Item = pStcPed.ColecaoPedLoc(lContX).Num_Item
                .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).NOSSO_NUMERO
                .Cod_Org = pStcPed.ColecaoPedLoc(lContX).Cod_Org
                .Cod_Ramo = pStcPed.ColecaoPedLoc(lContX).Cod_Ramo
                .Num_Ped = pStcPed.ColecaoPedLoc(lContX).Num_Ped
                .Tab_Cobert = pStcPed.ColecaoPedLoc(lContX).Cod_Plano
                If .Cod_Ramo = 112 Then
                    .Per_Desc_Exp = lBonus
                End If
                
                If .Cod_Grupo_Cobertura = 5 Then
                    Dim ColPedLoc As stcA46V704B
                    Set ColPedLoc = pStcPed.ColecaoPedLoc(1).ColecaoPedCobert(1)
                    .Per_Franq = ColPedLoc.Per_Franq
                    .Val_Franq_Min = ColPedLoc.Val_Franq_Min
                    .Val_Fator_Franq = ColPedLoc.Val_Fator_Franq
                    .TextoFranquia = ColPedLoc.TextoFranquia
                End If
                
                For lCountZ = 1 To pStcPed.ColecaoPedLoc(lContX).ColecaoPedCobert(lContY).ColecaoNovaFranquia.Count
                    With pStcPed.ColecaoPedLoc(lContX).ColecaoPedCobert(lContY).ColecaoNovaFranquia(1)
                        Set .ObjPai = pStcPed.ColecaoPedLoc(lContX).ColecaoPedCobert(lContY)
                        .Num_Item = pStcPed.ColecaoPedLoc(lContX).ColecaoPedCobert(lContY).Num_Item
                        .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).ColecaoPedCobert(lContY).NOSSO_NUMERO
                        .Cod_Tarif = pStcPed.Cod_Produto
                        .Cod_Ramo = pStcPed.ColecaoPedLoc(lContX).ColecaoPedCobert(lContY).Cod_Ramo
                        .Num_Ped = pStcPed.ColecaoPedLoc(lContX).ColecaoPedCobert(lContY).Num_Ped
                        .Cod_Modalidade = pStcPed.ColecaoPedLoc(lContX).ColecaoPedCobert(lContY).Tab_Cobert
                    End With
                Next
                
            End With
        Next

        For lContY = 1 To pStcPed.ColecaoPedLoc(lContX).ColecaoPedClau.Count
            With pStcPed.ColecaoPedLoc(lContX).ColecaoPedClau(lContY)
                Set .ObjPai = pStcPed.ColecaoPedLoc(lContX)
                .Num_Item = pStcPed.ColecaoPedLoc(lContX).Num_Item
                .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).NOSSO_NUMERO
                .Cod_Org = pStcPed.ColecaoPedLoc(lContX).Cod_Org
                .Cod_Ramo = pStcPed.ColecaoPedLoc(lContX).Cod_Ramo
                .Num_Ped = pStcPed.ColecaoPedLoc(lContX).Num_Ped
            End With
        Next

        If M46V999.gfPreenchido(pStcPed.ColecaoPedLoc(lContX).objstcPedObs.Observacao) Then
            With pStcPed.ColecaoPedLoc(lContX).objstcPedObs
                .Num_Item = pStcPed.ColecaoPedLoc(lContX).Num_Item
                .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).NOSSO_NUMERO
            End With
        End If

        For lContY = 1 To pStcPed.ColecaoPedLoc(lContX).ColecaoPedLocBem.Count
            With pStcPed.ColecaoPedLoc(lContX).ColecaoPedLocBem(lContY)
                Set .ObjPai = pStcPed.ColecaoPedLoc(lContX)
                .Num_Item = pStcPed.ColecaoPedLoc(lContX).Num_Item
                .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).NOSSO_NUMERO
                .Cod_Org = pStcPed.ColecaoPedLoc(lContX).Cod_Org
                .Cod_Ramo = pStcPed.ColecaoPedLoc(lContX).Cod_Ramo
                .Num_Ped = pStcPed.ColecaoPedLoc(lContX).Num_Ped
                .Cod_Seq = lContY
            End With
        Next

        For lContY = 1 To pStcPed.ColecaoPedLoc(lContX).ColecaoPedLocBicicleta.Count
           With pStcPed.ColecaoPedLoc(lContX).ColecaoPedLocBicicleta(lContY)
               .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).NOSSO_NUMERO
               .Cod_Ramo = pStcPed.ColecaoPedLoc(lContX).Cod_Ramo
               .Cod_Org = pStcPed.ColecaoPedLoc(lContX).Cod_Org
               .Num_Ped = pStcPed.ColecaoPedLoc(lContX).Num_Ped
               .Num_Apol = pStcPed.ColecaoPedLoc(lContX).Num_Apol
               .Num_Item = pStcPed.ColecaoPedLoc(lContX).Num_Item
               .Num_Endo = pStcPed.ColecaoPedLoc(lContX).Num_Endo
               .NUM_PI = pStcPed.ColecaoPedLoc(lContX).NUM_PI
           End With
       Next

        'Resdencial Supremo
        For lContY = 1 To pStcPed.ColecaoPedLoc(lContX).ColecaoAllRisks.Count
           With pStcPed.ColecaoPedLoc(lContX).ColecaoAllRisks(lContY)
               .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).NOSSO_NUMERO
               .Cod_Ramo = pStcPed.ColecaoPedLoc(lContX).Cod_Ramo
               .Cod_Org = pStcPed.ColecaoPedLoc(lContX).Cod_Org
               .Num_Ped = pStcPed.ColecaoPedLoc(lContX).Num_Ped
               .Num_Apol = pStcPed.ColecaoPedLoc(lContX).Num_Apol
               .Num_Item = pStcPed.ColecaoPedLoc(lContX).Num_Item
               .Num_Endo = pStcPed.ColecaoPedLoc(lContX).Num_Endo
               .NUM_PI = pStcPed.ColecaoPedLoc(lContX).NUM_PI
           End With
       Next
       
        For lContY = 1 To pStcPed.ColecaoPedLoc(lContX).ColecaoObrasArte.Count
           With pStcPed.ColecaoPedLoc(lContX).ColecaoObrasArte(lContY)
               .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).NOSSO_NUMERO
               .Cod_Ramo = pStcPed.ColecaoPedLoc(lContX).Cod_Ramo
               .Cod_Org = pStcPed.ColecaoPedLoc(lContX).Cod_Org
               .Num_Ped = pStcPed.ColecaoPedLoc(lContX).Num_Ped
               .Num_Apol = pStcPed.ColecaoPedLoc(lContX).Num_Apol
               .Num_Item = pStcPed.ColecaoPedLoc(lContX).Num_Item
               .Num_Endo = pStcPed.ColecaoPedLoc(lContX).Num_Endo
               .NUM_PI = pStcPed.ColecaoPedLoc(lContX).NUM_PI
           End With
       Next

        For lContY = 1 To pStcPed.ColecaoPedLoc(lContX).ColecaoPedSeguro.Count
            With pStcPed.ColecaoPedLoc(lContX).ColecaoPedSeguro(lContY)
                Set .ObjPai = pStcPed.ColecaoPedLoc(lContX)
                .Num_Item = pStcPed.ColecaoPedLoc(lContX).Num_Item
                .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).NOSSO_NUMERO
                .Cod_Org = pStcPed.ColecaoPedLoc(lContX).Cod_Org
                .Cod_Ramo = pStcPed.ColecaoPedLoc(lContX).Cod_Ramo
                .Num_Ped = pStcPed.ColecaoPedLoc(lContX).Num_Ped
            End With
        Next

        With pStcPed.ColecaoPedLoc(lContX).objstcPedInsp
            .Num_Item = pStcPed.ColecaoPedLoc(lContX).Num_Item
            .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).NOSSO_NUMERO
            .Cod_Org = pStcPed.ColecaoPedLoc(lContX).Cod_Org
            .Cod_Ramo = pStcPed.ColecaoPedLoc(lContX).Cod_Ramo
            .Num_Ped = pStcPed.ColecaoPedLoc(lContX).Num_Ped
            .Dat_Cadastro = pStcPed.ColecaoPedLoc(lContX).DAT_DIGIT
            .NUM_PI = pStcPed.ColecaoPedLoc(lContX).NUM_PI
        End With

        For lContY = 1 To pStcPed.ColecaoPedLoc(lContX).ColecaoPedFator.Count
            With pStcPed.ColecaoPedLoc(lContX).ColecaoPedFator(lContY)
                Set .ObjPai = pStcPed.ColecaoPedLoc(lContX)
                .Num_Item = pStcPed.ColecaoPedLoc(lContX).Num_Item
                .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).NOSSO_NUMERO
                .Cod_Org = pStcPed.ColecaoPedLoc(lContX).Cod_Org
                .Cod_Ramo = pStcPed.ColecaoPedLoc(lContX).Cod_Ramo
                .Num_Ped = pStcPed.ColecaoPedLoc(lContX).Num_Ped
            End With
        Next lContY


        'ProjetoFormularios
        For lContY = 1 To pStcPed.ColecaoPedLoc(lContX).ColecaoBeneficiario.Count
            With pStcPed.ColecaoPedLoc(lContX).ColecaoBeneficiario(lContY)
                Set .ObjPai = pStcPed.ColecaoPedLoc(lContX)
                .Tip_Loc_Pessoa = e_TipLocPessoa_Beneficiario
                .Num_Item = pStcPed.ColecaoPedLoc(lContX).Num_Item
                .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).NOSSO_NUMERO
                .NUM_PI = pStcPed.ColecaoPedLoc(lContX).Num_Ped
            End With
        Next


        'ProjetoFormularios
        With pStcPed.ColecaoPedLoc(lContX).objstcPedInquilino
            Set .ObjPai = pStcPed.ColecaoPedLoc(lContX)
            .Tip_Loc_Pessoa = e_TipLocPessoa_Inquilino
            .Num_Item = pStcPed.ColecaoPedLoc(lContX).Num_Item
            .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).NOSSO_NUMERO
            .NUM_PI = pStcPed.ColecaoPedLoc(lContX).Num_Ped
        End With


        'ProjetoFormularios
        With pStcPed.ColecaoPedLoc(lContX).objstcPedProprietario
            Set .ObjPai = pStcPed.ColecaoPedLoc(lContX)
            .Tip_Loc_Pessoa = e_TipLocPessoa_Proprietario
            .Num_Item = pStcPed.ColecaoPedLoc(lContX).Num_Item
            .NOSSO_NUMERO = pStcPed.ColecaoPedLoc(lContX).NOSSO_NUMERO
            .NUM_PI = pStcPed.ColecaoPedLoc(lContX).Num_Ped
        End With
        
    Next lContX
    Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Fim da Normaliza��o dos itens do pedido", 0)

    '***************** Retorno com Sucesso
    mfNormalizaPedido = True
    '*****************

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear
    GoSub FechaObj
    Exit Function

TrataErro:
    mfNormalizaPedido = False
    If M46V999.gfPreenchido(lMsrErr) = False Then
        lMsrErr = Err & " - " & ERROR
    End If
    pMsgErr = lMsrErr
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMsrErr, 2, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lobjPedLoc = Nothing
    Set ColPedLoc = Nothing
    Return
End Function

Private Sub mpImprimeObjeto(ByRef pobjstcPed As stcA46V702B, tagDynaTrace As String)
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mpImprimeObjeto"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lobjstcPedLoc As stcA46V708B
    Dim lobjstcPedClau As stcA46V703B
    Dim lobjstcPedCobert As stcA46V704B
    Dim lobjstcPedCorr As stcA46V705B
    Dim lobjstcPedCoss As stcA46V706B
    Dim lobjstcPedJuros As stcA46V711B
    Dim lobjstcPedParc As stcA46V713B
    Dim lobjstcPedBem As stcA46V715B
    Dim lobjstcPedEspec As stcA46V726B
    Dim lobjstcFator As stcA46V720B
    Dim lobjstcPedOutroSeg As stcA46V716B

    Dim lstcNorma As stcA46V708N
    Static flagImpObjeto As Byte

    'Controle para n�o imprimir log
    'Para gerar o log criar ou mudar a chave do ini (CONFIGURACOES/GravarLogCalc) para 1
    If flagImpObjeto = 0 Then
        flagImpObjeto = val(M46V001.LerIni("CONFIGURACOES", "GravarLogCalc", M46V999.gAppPath & "SyasMult.INI"))
    End If
    If flagImpObjeto <> 1 Then
        flagImpObjeto = 2
        GoTo FechaObj
    End If

    Call M46V999.gpGraLog(0, vbCrLf & "Dim lobjstcPedLoc   As stcA46V708B" & vbCrLf & _
                             "Dim lObjstcPedClau  As stcA46V703B" & vbCrLf & _
                             "Dim lobjstcPedCobert  As stcA46V704B" & vbCrLf & _
                             "Dim lobjstcPedCorr  As stcA46V705B" & vbCrLf & _
                             "Dim lobjstcPedCoss  As stcA46V706B" & vbCrLf & _
                             "Dim lobjstcPedEnder  As stcA46V707B" & vbCrLf & _
                             "Dim lobjstcPedJuros  As stcA46V711B" & vbCrLf & _
                             "Dim lobjstcPedCtrEmis As stcA46V712B" & vbCrLf & _
                             "Dim lobjstcPedParc As stcA46V713B" & vbCrLf & _
                             "Dim lobjstcPedObs As stcA46V714B" & vbCrLf & _
                             "Dim lobjstcPedBem As stcA46V715B" & vbCrLf & _
                             "Dim lobjstcPedOutroSeg As stcA46V716B" & vbCrLf & _
                             "Dim lobjstcPedInsp As stcA46V717B" & vbCrLf & _
                             "Dim lobjstcFator    As stcA46V720B" & vbCrLf & _
                             "Dim lobjstcPedEspec As stcA46V726B" & vbCrLf & _
                             "Dim lstcNorma       As stcA46V708N" & vbCrLf & _
                             "dim lobjstcPed      as stcA46V702B" & vbCrLf)
    With pobjstcPed
        Call M46V999.gpGraLog(0, vbCrLf & Replace(vbCrLf & "Set lobjstcPed = New stcA46V702B " & vbCrLf & "   With lobjstcPed " & vbCrLf & _
                                                  ".Nosso_Numero = " & Chr(34) & .NOSSO_NUMERO & Chr(34) & vbCrLf & _
                                                  ".Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                  ".Cod_Org = " & .Cod_Org & vbCrLf & _
                                                  ".Num_Ped = " & .Num_Ped & vbCrLf & _
                                                  ".Num_Apol = " & .Num_Apol & vbCrLf & _
                                                  ".Num_Endo = " & .Num_Endo & vbCrLf & _
                                                  ".Num_Apol_An = " & .Num_Apol_An & vbCrLf & _
                                                  ".Cod_Reg = " & .Cod_Reg & vbCrLf & _
                                                  ".Cod_Bco_Cobr = " & .Cod_Bco_Cobr & vbCrLf & _
                                                  ".Cod_Agen_Angar = " & .Cod_Agen_Angar & vbCrLf & _
                                                  ".Cod_Agen_Indicr = " & .Cod_Agen_Indicr & vbCrLf & _
                                                  ".Num_Cta_Indicr = " & .Num_Cta_Indicr & vbCrLf & _
                                                  ".Cod_Unid_Prod = " & .Cod_Unid_Prod & vbCrLf & _
                                                  ".Cod_Prod = " & .Cod_Prod & vbCrLf & _
                                                  ".Cod_Agen_Bas_Cor = " & .Cod_Agen_Bas_Cor & vbCrLf & _
                                                  ".Cod_Dig_Agen_Bas_Cor = " & .Cod_Dig_Agen_Bas_Cor & vbCrLf & _
                                                  ".Num_Cta_Bas_Cor = " & .Num_Cta_Bas_Cor & vbCrLf & _
                                                  ".Cod_Segurado = " & .Cod_Segurado & vbCrLf & _
                                                  ".Dig_Cod_Segurado = " & Chr(34) & .Dig_Cod_Segurado & Chr(34) & vbCrLf & _
                                                  ".Nom_Segurado = " & Chr(34) & .Nom_Segurado & Chr(34) & vbCrLf & _
                                                  ".Dsc_Compl_Segurado = " & Chr(34) & .Dsc_Compl_Segurado & Chr(34) & vbCrLf & _
                                                  ".Num_Cgc_Cpf = " & .Num_Cgc_Cpf & vbCrLf & _
                                                  ".Cod_Tip_Logr = " & Chr(34) & .Cod_Tip_Logr & Chr(34) & vbCrLf & _
                                                  ".Nom_Logr = " & Chr(34) & .Nom_Logr & Chr(34) & vbCrLf, ",", "."))

        Call M46V999.gpGraLog(0, vbCrLf & Replace(".Dsc_Compl_Logr = " & Chr(34) & .Dsc_Compl_Logr & Chr(34) & vbCrLf & _
                                                  ".Nom_Bairro = " & Chr(34) & .Nom_Bairro & Chr(34) & vbCrLf & _
                                                  ".Nom_Cid = " & Chr(34) & .Nom_Cid & Chr(34) & vbCrLf & _
                                                  ".Sig_UF = " & Chr(34) & .Sig_UF & Chr(34) & vbCrLf & _
                                                  ".Num_CEP = " & .Num_CEP & vbCrLf & _
                                                  ".Num_Fone = " & .Num_Fone & vbCrLf & _
                                                  ".Cod_Estip = " & .Cod_Estip & vbCrLf & _
                                                  ".Val_Coef_Cms_Pad = " & .Val_Coef_Cms_Pad & vbCrLf & _
                                                  ".Dat_Ini_Vig = " & .Dat_Ini_Vig & vbCrLf & _
                                                  ".Dat_Ter_Vig = " & .Dat_Ter_Vig & vbCrLf & _
                                                  ".Dat_Emis = " & .Dat_Emis & vbCrLf & _
                                                  ".Cod_Moed = " & .Cod_Moed & vbCrLf & _
                                                  ".Per_Juro_Frac = " & .Per_Juro_Frac & vbCrLf & _
                                                  ".Qtd_Parc = " & .Qtd_Parc & vbCrLf & _
                                                  ".Qtd_Anexo = " & .Qtd_Anexo & vbCrLf & _
                                                  ".Cod_Iden_Praz = " & .Cod_Iden_Praz & vbCrLf & _
                                                  ".Cod_Pr_Min_Endo = " & .Cod_Pr_Min_Endo & vbCrLf & _
                                                  ".Cod_Pr_Min_Parc = " & .Cod_Pr_Min_Parc & vbCrLf & _
                                                  ".Val_Pr_Antec = " & .Val_Pr_Antec & vbCrLf & _
                                                  ".Per_Adic_Form_Pag = " & .Per_Adic_Form_Pag & vbCrLf & _
                                                  ".Per_Desc_Form_Pag = " & .Per_Desc_Form_Pag & vbCrLf & _
                                                  ".Val_IS = " & .Val_IS & vbCrLf & _
                                                  ".Val_IS_2 = " & .Val_IS_2 & vbCrLf & _
                                                  ".VAL_PR = " & .VAL_PR & vbCrLf & _
                                                  ".Val_Pr_Ref_Tarif = " & .Val_Pr_Ref_Tarif & vbCrLf, ",", "."))

        Call M46V999.gpGraLog(0, vbCrLf & Replace(".Val_Pr_Ref_Resg = " & .Val_Pr_Ref_Resg & vbCrLf & _
                                                  ".Dat_Ini_Viag = " & .Dat_Ini_Viag & vbCrLf & _
                                                  ".Cod_Tip_Liber = " & .Cod_Tip_Liber & vbCrLf & _
                                                  ".Cod_Iden_Alt_Corr = " & Chr(34) & .Cod_Iden_Alt_Corr & Chr(34) & vbCrLf & _
                                                  ".Cod_User_Digit = " & Chr(34) & .Cod_User_Digit & Chr(34) & vbCrLf & _
                                                  ".DAT_DIGIT = " & .DAT_DIGIT & vbCrLf & _
                                                  ".HOR_DIGIT = " & .HOR_DIGIT & vbCrLf & _
                                                  ".Cod_User_Liber = " & Chr(34) & .Cod_User_Liber & Chr(34) & vbCrLf & _
                                                  ".Dat_Liber = " & .Dat_Liber & vbCrLf & _
                                                  ".Hor_Liber = " & .Hor_Liber & vbCrLf & _
                                                  ".Cod_Sit_Ped = " & .Cod_Sit_Ped & vbCrLf & _
                                                  ".Cod_Tip_Doc = " & .Cod_Tip_Doc & vbCrLf & _
                                                  ".Cod_Dizer = " & .Cod_Dizer & vbCrLf & _
                                                  ".Qtd_Copia = " & .Qtd_Copia & vbCrLf & _
                                                  ".Val_Cust_Emis = " & .Val_Cust_Emis & vbCrLf & _
                                                  ".Val_Parc_Min = " & .Val_Parc_Min & vbCrLf & _
                                                  ".Val_Juro = " & .Val_Juro & vbCrLf & _
                                                  ".Val_IOF = " & .Val_IOF & vbCrLf & _
                                                  ".Num_Prop = " & .Num_Prop & vbCrLf & _
                                                  ".Cod_Posi_Atual = " & Chr(34) & .Cod_Posi_Atual & Chr(34) & vbCrLf & _
                                                  ".Dat_Mov_Ini = " & .Dat_Mov_Ini & vbCrLf & _
                                                  ".Dat_Mov_Ter = " & .Dat_Mov_Ter & vbCrLf & _
                                                  ".Tip_Cust_Emis = " & .Tip_Cust_Emis & vbCrLf & _
                                                  ".Val_Cambio_Emis = " & .Val_Cambio_Emis & vbCrLf & _
                                                  ".Iden = " & Chr(34) & .Iden & Chr(34) & vbCrLf, ",", "."))

        Call M46V999.gpGraLog(0, vbCrLf & Replace(".Num_PI = " & .NUM_PI & vbCrLf & _
                                                  ".Cod_Produto = " & .Cod_Produto & vbCrLf & _
                                                  ".Cod_Protocolo = " & .Cod_Protocolo & vbCrLf & _
                                                  ".Cod_Sequencia = " & .Cod_Sequencia & vbCrLf & _
                                                  ".Tip_Pessoa = " & .Tip_Pessoa & vbCrLf & _
                                                  ".Cod_Ramo_Ativ = " & .Cod_Ramo_Ativ & vbCrLf & _
                                                  ".Cod_Doc = " & .Cod_Doc & vbCrLf & _
                                                  ".Num_Doc = " & Chr(34) & .Num_Doc & Chr(34) & vbCrLf & _
                                                  ".Org_Exp = " & Chr(34) & .Org_Exp & Chr(34) & vbCrLf & _
                                                  ".Dat_Emis_Doc = " & .Dat_Emis_Doc & vbCrLf & _
                                                  ".Nom_Empresa = " & Chr(34) & .Nom_Empresa & Chr(34) & vbCrLf & _
                                                  ".Sexo = " & Chr(34) & .Sexo & Chr(34) & vbCrLf & _
                                                  ".Dat_Pri_Calculo = " & .Dat_Pri_Calculo & vbCrLf & _
                                                  ".Ult_Tarifa = " & .Ult_Tarifa & vbCrLf & _
                                                  ".Cod_Unicidade = " & .Cod_Unicidade & vbCrLf & _
                                                  ".Cod_Ver_Unici = " & .Cod_Ver_Unici & vbCrLf & _
                                                  ".Val_PriParcela = " & .Val_PriParcela & vbCrLf & _
                                                  ".val_DemParcelas = " & .val_DemParcelas & vbCrLf & _
                                                  ".Desc_Plano = " & Chr(34) & .Desc_Plano & Chr(34) & vbCrLf & _
                                                  ".Val_RendaFatu = " & .Val_RendaFatu & vbCrLf & _
                                                  ".LMI_Unico = " & .LMI_Unico & vbCrLf & _
                                                  ".Per_Desc_LMI = " & .Per_Desc_LMI & vbCrLf & _
                                                  ".Num_Cotac_LMI = " & Chr(34) & .Num_Cotac_LMI & Chr(34) & vbCrLf & _
                                                  ".Tip_Transmissao = " & .Tip_Transmissao & vbCrLf & _
                                                  ".Ind_CondGerais = " & .Ind_CondGerais & vbCrLf, ",", "."))

        Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .ColecaoPedCorr = New Collection " & vbCrLf, ",", "."))
        For Each lobjstcPedCorr In .ColecaoPedCorr
            With lobjstcPedCorr
                Call M46V999.gpGraLog(0, vbCrLf & Replace("set lobjstcPedCorr = new stcA46V705B " & vbCrLf & _
                                                          "   with lobjstcPedCorr " & vbCrLf & _
                                                          "        .Nosso_Numero = " & Chr(34) & .NOSSO_NUMERO & Chr(34) & vbCrLf & _
                                                          "        .Cod_Corr = " & .Cod_Corr & vbCrLf & _
                                                          "        .Num_Ped = " & .Num_Ped & vbCrLf & _
                                                          "        .Cod_Org = " & .Cod_Org & vbCrLf & _
                                                          "        .Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                          "        .Num_Apol = " & .Num_Apol & vbCrLf & _
                                                          "        .Num_Endo = " & .Num_Endo & vbCrLf & _
                                                          "        .Per_Corr = " & .Per_Corr & vbCrLf & _
                                                          "        .Num_Ordem_Corr = " & .Num_Ordem_Corr & vbCrLf & _
                                                          "        .Num_PI = " & .NUM_PI & vbCrLf & _
                                                          "   End With" & vbCrLf, ",", "."))
                Call M46V999.gpGraLog(0, vbCrLf & ".ColecaoPedCorr.add  lobjstcPedCorr, lobjstcPedCorr.class_id " & vbCrLf)
            End With
        Next lobjstcPedCorr

        Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .ColecaoPedCoss = New Collection  " & vbCrLf, ",", "."))
        For Each lobjstcPedCoss In .ColecaoPedCoss
            With lobjstcPedCoss
                Call M46V999.gpGraLog(0, vbCrLf & Replace("set lobjstcPedCoss = new  stcA46V706B " & vbCrLf & " With lobjstcPedCoss " & vbCrLf & _
                                                          "        .Nosso_Numero = " & Chr(34) & .NOSSO_NUMERO & Chr(34) & vbCrLf & _
                                                          "        .Cod_Coss = " & .Cod_Coss & vbCrLf & _
                                                          "        .Num_Ped = " & .Num_Ped & vbCrLf & _
                                                          "        .Cod_Org = " & .Cod_Org & vbCrLf & _
                                                          "        .Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                          "        .Num_Apol = " & .Num_Apol & vbCrLf & _
                                                          "        .Num_Endo = " & .Num_Endo & vbCrLf & _
                                                          "        .Per_Coss = " & .Per_Coss & vbCrLf & _
                                                          "        .Num_PI = " & .NUM_PI & vbCrLf & _
                                                          "        .Nom_Seguradora = " & .Nom_Seguradora & vbCrLf & _
                                                          "    End With " & vbCrLf, ",", "."))
                Call M46V999.gpGraLog(0, vbCrLf & "   .ColecaoPedCoss.add lobjstcPedCoss,lobjstcPedCoss.class_id" & vbCrLf)
            End With
        Next lobjstcPedCoss

        Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .ColecaoPedJuros = New Collection " & vbCrLf, ",", "."))
        For Each lobjstcPedJuros In .ColecaoPedJuros
            With lobjstcPedJuros
                Call M46V999.gpGraLog(0, vbCrLf & Replace("set lobjstcPedJuros  = new stcA46V711B " & vbCrLf & " with lobjstcPedJuros " & vbCrLf & _
                                                          ".Nosso_Numero = " & Chr(34) & .NOSSO_NUMERO & Chr(34) & vbCrLf & _
                                                          ".Num_Parcela = " & .Num_Parcela & vbCrLf & _
                                                          ".Val_Pos = " & .Val_Pos & vbCrLf & _
                                                          ".Val_Ant = " & .Val_Ant & vbCrLf & _
                                                          " end with " & vbCrLf, ",", "."))
                Call M46V999.gpGraLog(0, vbCrLf & " .ColecaoPedJuros.add lobjstcPedJuros, lobjstcPedJuros.class_id" & vbCrLf)
            End With
        Next lobjstcPedJuros

        Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .ColecaoPedVenc = New Collection" & vbCrLf, ",", "."))
        For Each lobjstcPedParc In .ColecaoPedVenc
            With lobjstcPedParc
                Call M46V999.gpGraLog(0, vbCrLf & Replace("set lobjstcPedParc = new stcA46V713B " & vbCrLf & " with lobjstcPedParc " & vbCrLf & _
                                                          ".Nosso_Numero = " & Chr(34) & .NOSSO_NUMERO & Chr(34) & vbCrLf & _
                                                          ".Cod_Org = " & .Cod_Org & vbCrLf & _
                                                          ".Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                          ".Num_Ped = " & .Num_Ped & vbCrLf & _
                                                          ".Num_Apol = " & .Num_Apol & vbCrLf & _
                                                          ".Num_Endo = " & .Num_Endo & vbCrLf & _
                                                          ".Num_Parc = " & .Num_Parc & vbCrLf & _
                                                          ".DAT_VENC = " & .DAT_VENC & vbCrLf & _
                                                          " End With " & vbCrLf, ",", "."))
                Call M46V999.gpGraLog(0, vbCrLf & ".ColecaoPedVenc.add lobjstcPedParc, lobjstcPedParc.class_id " & vbCrLf)
            End With
        Next lobjstcPedParc

        Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .objstcPedEndereco = New stcA46V707B" & vbCrLf, ",", "."))
        With .objstcPedEndereco
            Call M46V999.gpGraLog(0, vbCrLf & Replace("With .objstcPedEndereco " & vbCrLf & _
                                                      ".Nosso_Numero = " & Chr(34) & .NOSSO_NUMERO & Chr(34) & vbCrLf & _
                                                      ".Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                      ".Num_Apol = " & .Num_Apol & vbCrLf & _
                                                      ".Num_Endo = " & .Num_Endo & vbCrLf & _
                                                      ".Cod_Tip_Logr = " & Chr(34) & .Cod_Tip_Logr & Chr(34) & vbCrLf & _
                                                      ".Nom_Logr = " & Chr(34) & .Nom_Logr & Chr(34) & vbCrLf & _
                                                      ".Dsc_Compl_Logr = " & Chr(34) & .Dsc_Compl_Logr & Chr(34) & vbCrLf & _
                                                      ".Nom_Bairro = " & Chr(34) & .Nom_Bairro & Chr(34) & vbCrLf & _
                                                      ".Nom_Cid = " & Chr(34) & .Nom_Cid & Chr(34) & vbCrLf & _
                                                      ".Sig_UF = " & Chr(34) & .Sig_UF & Chr(34) & vbCrLf & _
                                                      ".Num_CEP = " & .Num_CEP & vbCrLf & _
                                                      ".Num_Fone = " & .Num_Fone & vbCrLf & _
                                                      ".Num_PI = " & .NUM_PI & vbCrLf & _
                                                      ".Cod_Reg_Cons = " & .Cod_Reg_Cons & vbCrLf & _
                                                      ".Num_Reg_Cons = " & Chr(34) & .Num_Reg_Cons & Chr(34) & vbCrLf & _
                                                      "End With " & vbCrLf, ",", "."))
        End With

        Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .objstcCtrlEmis = New stcA46V712B " & vbCrLf, ",", "."))
        With .objstcCtrlEmis
            Call M46V999.gpGraLog(0, vbCrLf & Replace("With .objstcCtrlEmis" & vbCrLf & _
                                                      ".Nosso_Numero = " & Chr(34) & .NOSSO_NUMERO & Chr(34) & vbCrLf & _
                                                      ".Num_PI = " & .NUM_PI & vbCrLf & _
                                                      ".Num_Bloq = " & .Num_Bloq & vbCrLf & _
                                                      ".COD_UNID = " & .COD_UNID & vbCrLf & _
                                                      ".Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                      ".Dat_Ini_Vig = " & .Dat_Ini_Vig & vbCrLf & _
                                                      ".Dat_Ter_Vig = " & .Dat_Ter_Vig & vbCrLf & _
                                                      ".Cod_Corr = " & .Cod_Corr & vbCrLf & _
                                                      ".Per_Corr = " & .Per_Corr & vbCrLf & _
                                                      ".Cod_Corr_2 = " & .Cod_Corr_2 & vbCrLf & _
                                                      ".Per_Corr_2 = " & .Per_Corr_2 & vbCrLf & _
                                                      ".Cod_Corr_3 = " & .Cod_Corr_3 & vbCrLf & _
                                                      ".Per_Corr_3 = " & .Per_Corr_3 & vbCrLf & _
                                                      ".Nom_Seg = " & Chr(34) & .Nom_Seg & Chr(34) & vbCrLf & _
                                                      ".Qtd_Anex = " & .Qtd_Anex & vbCrLf & _
                                                      ".Num_Apol_Ant = " & .Num_Apol_Ant & vbCrLf & _
                                                      ".Dsc_Obs_Prop = " & Chr(34) & .Dsc_Obs_Prop & Chr(34) & vbCrLf & _
                                                      ".Tip_Cotac = " & Chr(34) & .Tip_Cotac & Chr(34) & vbCrLf & _
                                                      ".Num_Cotac = " & Chr(34) & .Num_Cotac & Chr(34) & vbCrLf & _
                                                      ".Per_Juros_Neg = " & .Per_Juros_Neg & vbCrLf & _
                                                      ".Cod_Envi_Doc = " & Chr(34) & .Cod_Envi_Doc & Chr(34) & vbCrLf & _
                                                      ".DAT_DIGIT = " & .DAT_DIGIT & vbCrLf & _
                                                      ".HOR_DIGIT = " & .HOR_DIGIT & vbCrLf & _
".Cod_Iden_Stat_Acao = " & .Cod_Iden_Stat_Acao & vbCrLf, ",", "."))
            Call M46V999.gpGraLog(0, vbCrLf & Replace(".Cod_User_Digit = " & Chr(34) & .Cod_User_Digit & Chr(34) & vbCrLf & _
                                                      ".Dat_Digit_MF = " & .Dat_Digit_MF & vbCrLf & _
                                                      ".Hor_Digit_MF = " & .Hor_Digit_MF & vbCrLf & _
                                                      ".Cod_User_Digit_MF = " & Chr(34) & .Cod_User_Digit_MF & Chr(34) & vbCrLf & _
                                                      ".Val_Pr_Tot = " & .Val_Pr_Tot & vbCrLf & _
                                                      ".Cod_Iden_Classe = " & .Cod_Iden_Classe & vbCrLf & _
                                                      ".Cod_Iden_Cobr = " & .Cod_Iden_Cobr & vbCrLf & _
                                                      ".Num_Prop = " & .Num_Prop & vbCrLf & _
                                                      ".Dat_Canc_Apol = " & .Dat_Canc_Apol & vbCrLf & _
                                                      ".Cod_Tip_Doc = " & .Cod_Tip_Doc & vbCrLf & _
                                                      ".Dat_Rec_Cia = " & .Dat_Rec_Cia & vbCrLf & _
                                                      ".Num_Ref_Corr = " & Chr(34) & .Num_Ref_Corr & Chr(34) & vbCrLf & _
                                                      ".Dat_Emis = " & .Dat_Emis & vbCrLf & _
                                                      ".Num_Apol = " & .Num_Apol & vbCrLf & _
                                                      ".Num_Endo = " & .Num_Endo & vbCrLf & _
                                                      ".Val_Pr_Emis = " & .Val_Pr_Emis & vbCrLf & _
                                                      ".Val_Pr_Liq = " & .Val_Pr_Liq & vbCrLf & _
                                                      ".Per_Tot_Coss = " & .Per_Tot_Coss & vbCrLf & _
                                                      ".Qtd_Parc = " & .Qtd_Parc & vbCrLf & _
                                                      ".Cod_Iden_Stat_Resp = " & .Cod_Iden_Stat_Resp & vbCrLf & _
                                                      ".Cod_Iden_SAC = " & .Cod_Iden_SAC & vbCrLf & _
                                                      ".Cod_Prod = " & .Cod_Prod & vbCrLf & _
                                                      ".Val_IS = " & .Cod_Prod & vbCrLf & _
                                                      ".Per_Desc_Coml = " & .Per_Desc_Coml & vbCrLf & _
".Num_Cgc_Cpf = " & .Num_Cgc_Cpf & vbCrLf, ",", "."))
            Call M46V999.gpGraLog(0, vbCrLf & Replace(".Cod_Iden_Pessoa = " & .Cod_Iden_Pessoa & vbCrLf & _
                                                      ".Cod_Iden_Pag_Antec = " & .Cod_Iden_Pag_Antec & vbCrLf & _
                                                      ".DIG_BLOQUETO = " & .DIG_BLOQUETO & vbCrLf & _
                                                      ".Parcelas_Impressao_Avista = " & .Parcelas_Impressao_Avista & vbCrLf & _
                                                      ".Parcelas_Impressao_30dias = " & .Parcelas_Impressao_30dias & vbCrLf & _
                                                      ".Parcelas_Impressao_AvistaDebito = " & .Parcelas_Impressao_AvistaDebito & vbCrLf & _
                                                      ".Parcelas_Impressao_30diasDebito = " & .Parcelas_Impressao_30diasDebito & vbCrLf & _
                                                      ".Parcelas_Impressao_Cartao = " & .Parcelas_Impressao_Cartao & vbCrLf & _
                                                      ".Email_Contato = " & .Email_Contato & vbCrLf & _
                                                      ".Ind_Impr_Apol = " & .Ind_Impr_Apol & vbCrLf & _
                                                      ".Ind_Prolabore = " & .Ind_Prolabore & vbCrLf & _
                                                      ".Ind_Agenciamento = " & .Ind_Agenciamento & vbCrLf & _
                                                      ".Cod_Bco_Boleto = " & .Cod_Bco_Boleto & vbCrLf & _
                                                      ".cod_origem = " & .cod_origem & vbCrLf & _
                                                      ".Num_Apol_Mar = " & .Num_Apol_Mar & vbCrLf & _
                                                      "end with " & vbCrLf, ",", "."))
        End With

        Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .ColecaoTextoEspec = New Collection" & vbCrLf, ",", "."))
        For Each lobjstcPedEspec In .ColecaoTextoEspec
            With lobjstcPedEspec
                Call M46V999.gpGraLog(0, vbCrLf & Replace("set lobjstcPedEspec = new stcA46V726B " & vbCrLf & " With lobjstcPedEspec " & vbCrLf & _
                                                          ".Nosso_Numero = " & Chr(34) & .NOSSO_NUMERO & Chr(34) & vbCrLf & _
                                                          ".Cod_Espec = " & .Cod_Espec & vbCrLf & _
                                                          ".Cod_Org = " & .Cod_Org & vbCrLf & _
                                                          ".Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                          ".Num_Ped = " & .Num_Ped & vbCrLf & _
                                                          ".Num_Apol = " & .Num_Apol & vbCrLf & _
                                                          ".Num_Endo = " & .Num_Endo & vbCrLf & _
                                                          ".Num_pag = " & .Num_pag & vbCrLf & _
                                                          ".Qtd_vias = " & .Qtd_vias & vbCrLf & _
                                                          ".Espec_lin1 = " & Chr(34) & .Espec_lin1 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin2 = " & Chr(34) & .Espec_lin2 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin3 = " & Chr(34) & .Espec_lin3 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin4 = " & Chr(34) & .Espec_lin4 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin5 = " & Chr(34) & .Espec_lin5 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin6 = " & Chr(34) & .Espec_lin6 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin7 = " & Chr(34) & .Espec_lin7 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin8 = " & Chr(34) & .Espec_lin8 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin9 = " & Chr(34) & .Espec_lin9 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin10 = " & Chr(34) & .Espec_lin10 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin11 = " & Chr(34) & .Espec_lin11 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin12 = " & Chr(34) & .Espec_lin12 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin13 = " & Chr(34) & .Espec_lin13 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin14 = " & Chr(34) & .Espec_lin14 & Chr(34) & vbCrLf & _
".Espec_lin15 = " & Chr(34) & .Espec_lin15 & Chr(34), ",", "."))
                Call M46V999.gpGraLog(0, vbCrLf & Replace(".Espec_lin16 = " & Chr(34) & .Espec_lin16 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin17 = " & Chr(34) & .Espec_lin17 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin18 = " & Chr(34) & .Espec_lin18 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin19 = " & Chr(34) & .Espec_lin19 & Chr(34) & vbCrLf & _
                                                          ".Espec_lin20 = " & Chr(34) & .Espec_lin20 & Chr(34) & vbCrLf & _
                                                          ".Num_PI = " & .NUM_PI & vbCrLf & _
                                                          "End With " & vbCrLf, ",", "."))
                Call M46V999.gpGraLog(0, vbCrLf & "ColecaoTextoEspec.add lobjstcPedEspec,lobjstcPedEspec.class_id" & vbCrLf)
            End With
        Next lobjstcPedEspec


        Call M46V999.gpGraLog(0, vbCrLf & Replace(" Set .ColecaoPedLoc = New Collection " & vbCrLf & "With lobjstcPedLoc" & vbCrLf, ",", "."))
        For Each lobjstcPedLoc In .ColecaoPedLoc
            With lobjstcPedLoc
                Call M46V999.gpGraLog(0, vbCrLf & Replace(" .Nosso_Numero = " & Chr(34) & .NOSSO_NUMERO & Chr(34) & vbCrLf & _
                                                          ".Num_Item = " & .Num_Item & vbCrLf & _
                                                          ".Num_Ped = " & .Num_Ped & vbCrLf & _
                                                          ".Cod_Org = " & .Cod_Org & vbCrLf & _
                                                          ".Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                          ".Num_Apol = " & .Num_Apol & vbCrLf & _
                                                          ".Num_Endo = " & .Num_Endo & vbCrLf & _
                                                          ".Cod_Tip_Logr_Risco = " & Chr(34) & .Cod_Tip_Logr_Risco & Chr(34) & vbCrLf & _
                                                          ".Nom_Logr_Risco = " & Chr(34) & .Nom_Logr_Risco & Chr(34) & vbCrLf & _
                                                          ".Dsc_Compl_Risco = " & Chr(34) & .Dsc_Compl_Risco & Chr(34) & vbCrLf & _
                                                          ".Nom_Bairro_Risco = " & Chr(34) & .Nom_Bairro_Risco & Chr(34) & vbCrLf & _
                                                          ".Nom_Cid_Risco = " & Chr(34) & .Nom_Cid_Risco & Chr(34) & vbCrLf & _
                                                          ".Sig_UF_Risco = " & Chr(34) & .Sig_UF_Risco & Chr(34) & vbCrLf & _
                                                          ".Num_Cep_Risco = " & .Num_Cep_Risco & vbCrLf & _
                                                          ".Per_Desc_Apar = " & .Per_Desc_Apar & vbCrLf & _
                                                          ".Per_Desc_Progres = " & .Per_Desc_Progres & vbCrLf & _
                                                          ".Per_IS_Vr = " & .Per_IS_Vr & vbCrLf & _
                                                          ".Val_Coef_IS_Vr = " & .Val_Coef_IS_Vr & vbCrLf & _
                                                          ".Val_Vr = " & .Val_Vr & vbCrLf & _
                                                          ".Tab_Clasf = " & .Tab_Clasf & vbCrLf & _
                                                          ".Cod_Clasf = " & Chr(34) & .Cod_Clasf & Chr(34) & vbCrLf & _
                                                          ".Cod_Compl_Clasf = " & Chr(34) & .Cod_Compl_Clasf & Chr(34) & vbCrLf & _
                                                          ".Dsc_Clasf = " & Chr(34) & .Dsc_Clasf & Chr(34) & vbCrLf & _
                                                          ".Cod_Classe_Ocup = " & .Cod_Classe_Ocup & vbCrLf & _
                                                          ".Cod_Constr = " & .Cod_Constr & vbCrLf, ",", "."))

                Call M46V999.gpGraLog(0, vbCrLf & Replace(".Cod_User_Digit = " & Chr(34) & .Cod_User_Digit & Chr(34) & vbCrLf & _
                                                          ".DAT_DIGIT = " & .DAT_DIGIT & vbCrLf & _
                                                          ".HOR_DIGIT = " & .HOR_DIGIT & vbCrLf & _
                                                          ".Cod_User_Liber = " & Chr(34) & .Cod_User_Liber & Chr(34) & vbCrLf & _
                                                          ".Dat_Liber = " & .Dat_Liber & vbCrLf & _
                                                          ".Hor_Liber = " & .Hor_Liber & vbCrLf & _
                                                          ".Cod_Sit_item = " & .Cod_Sit_item & vbCrLf & _
                                                          ".Per_Bonus = " & .Per_Bonus & vbCrLf & _
                                                          ".Qtd_Empreg = " & .Qtd_Empreg & vbCrLf & _
                                                          ".Num_Insp = " & .Num_Insp & vbCrLf & _
                                                          ".Cod_Iden_Segur_01_V = " & Chr(34) & .Cod_Iden_Segur_01_V & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Segur_02_A = " & Chr(34) & .Cod_Iden_Segur_02_A & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Segur_03_P = " & Chr(34) & .Cod_Iden_Segur_03_P & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Segur_04_I = " & Chr(34) & .Cod_Iden_Segur_04_I & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Segur_05_G = " & Chr(34) & .Cod_Iden_Segur_05_G & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Segur_06 = " & Chr(34) & .Cod_Iden_Segur_06 & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Segur_07 = " & Chr(34) & .Cod_Iden_Segur_07 & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Segur_08 = " & Chr(34) & .Cod_Iden_Segur_08 & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Segur_09 = " & Chr(34) & .Cod_Iden_Segur_09 & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Segur_10 = " & Chr(34) & .Cod_Iden_Segur_10 & Chr(34) & vbCrLf & _
                                                          ".Per_Tot_Segur = " & .Per_Tot_Segur & vbCrLf & _
                                                          ".Cod_Iden_Protec_01_E = " & Chr(34) & .Cod_Iden_Protec_01_E & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Protec_02_H = " & Chr(34) & .Cod_Iden_Protec_02_H & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Protec_03_S = " & Chr(34) & .Cod_Iden_Protec_03_S & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Protec_04_D = " & Chr(34) & .Cod_Iden_Protec_04_D & Chr(34) & vbCrLf, ",", "."))

                Call M46V999.gpGraLog(0, vbCrLf & Replace(".Cod_Iden_Protec_05 = " & Chr(34) & .Cod_Iden_Protec_05 & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Protec_06 = " & Chr(34) & .Cod_Iden_Protec_06 & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Protec_07 = " & Chr(34) & .Cod_Iden_Protec_07 & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Protec_08 = " & Chr(34) & .Cod_Iden_Protec_08 & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Protec_09 = " & Chr(34) & .Cod_Iden_Protec_09 & Chr(34) & vbCrLf & _
                                                          ".Cod_Iden_Protec_10 = " & Chr(34) & .Cod_Iden_Protec_10 & Chr(34) & vbCrLf & _
                                                          ".Per_Tot_Protec = " & .Per_Tot_Protec & vbCrLf & _
                                                          ".Cod_Produto = " & .Cod_Produto & vbCrLf & _
                                                          ".Cod_Regiao = " & .Cod_Regiao & vbCrLf & _
                                                          ".Cod_Iden_Tarif = " & .Cod_Iden_Tarif & vbCrLf & _
                                                          ".Num_PI = " & .NUM_PI & vbCrLf & _
                                                          ".Num_Item_Ant = " & .Num_Item_Ant & vbCrLf & _
                                                          ".Tip_Emissao = " & .Tip_Emissao & vbCrLf & _
                                                          ".Num_Apol_Anterior = " & .Num_Apol_Anterior & vbCrLf & _
                                                          ".Num_Item_Anterior = " & .Num_Item_Anterior & vbCrLf & _
                                                          ".Ind_Sinistro = " & .Ind_Sinistro & vbCrLf & _
                                                          ".Insp_Obr = " & .Insp_Obr & vbCrLf & _
                                                          ".Insp_Agend = " & .Insp_Agend & vbCrLf & _
                                                          ".Dig_Bens = " & .Dig_Bens & vbCrLf & _
                                                          ".Cob_Bens = " & .Cob_Bens & vbCrLf & _
                                                          ".Tip_Atividade = " & .Tip_Atividade & vbCrLf & _
                                                          ".Ren_Con_Seguradora = " & .Ren_Con_Seguradora & vbCrLf & _
                                                          ".Ren_Con_Apolice = " & Chr(34) & .Ren_Con_Apolice & Chr(34) & vbCrLf & _
                                                          ".Ren_Con_Item = " & Chr(34) & .Ren_Con_Item & Chr(34) & vbCrLf & _
                                                          ".Tip_Cobert = " & .Tip_Cobert & vbCrLf, ",", "."))

                Call M46V999.gpGraLog(0, vbCrLf & Replace(".Ind_VerbaSep = " & .Ind_VerbaSep & vbCrLf, ",", "."))

                Call M46V999.gpGraLog(0, vbCrLf & Replace(".Cod_Plano = " & .Cod_Plano & vbCrLf & _
                                                          "Set .ColecaoPedCobert = New Collection " & vbCrLf, ",", "."))

                For Each lobjstcPedCobert In .ColecaoPedCobert
                    With lobjstcPedCobert

                        Call M46V999.gpGraLog(0, vbCrLf & Replace("set lobjstcPedCobert  = new stcA46V704B" & vbCrLf & "with lobjstcPedCobert  " & vbCrLf & _
                                                                  ".Nosso_Numero = " & Chr(34) & .NOSSO_NUMERO & Chr(34) & vbCrLf & _
                                                                  ".Cod_Cobert = " & .Cod_Cobert & vbCrLf & _
                                                                  ".Num_Ped = " & .Num_Ped & vbCrLf & _
                                                                  ".Cod_Org = " & .Cod_Org & vbCrLf & _
                                                                  ".Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                                  ".Num_Apol = " & .Num_Apol & vbCrLf & _
                                                                  ".Num_Endo = " & .Num_Endo & vbCrLf & _
                                                                  ".Num_Item = " & .Num_Item & vbCrLf & _
                                                                  ".Tab_Cobert = " & .Tab_Cobert & vbCrLf & _
                                                                  ".Val_IS = " & .Val_IS & vbCrLf & _
                                                                  ".Per_Taxa = " & .Per_Taxa & vbCrLf & _
                                                                  ".Per_Desc_Coml = " & .Per_Desc_Coml & vbCrLf & _
                                                                  ".Per_Desc_Exp = " & .Per_Desc_Exp & vbCrLf & _
                                                                  ".VAL_PR = " & .VAL_PR & vbCrLf & _
                                                                  ".Val_Pr_Anu = " & .Val_Pr_Anu & vbCrLf & _
                                                                  ".Val_Pr_Tarif = " & .Val_Pr_Tarif & vbCrLf & _
                                                                  ".Val_Pr_Tarif_Anu = " & .Val_Pr_Tarif_Anu & vbCrLf & _
                                                                  ".Val_Pr_Resg = " & .Val_Pr_Resg & vbCrLf & _
                                                                  ".Val_Pr_Resg_Anu = " & .Val_Pr_Resg_Anu & vbCrLf & _
                                                                  ".Per_Franq = " & .Per_Franq & vbCrLf & _
                                                                  ".Val_Franq_Min = " & .Val_Franq_Min & vbCrLf & _
                                                                  ".Val_Franq_Max = " & .Val_Franq_Max & vbCrLf & _
                                                                  ".Num_PI = " & .NUM_PI & vbCrLf & _
".Num_Meses_PI = " & .Num_Meses_PI & vbCrLf, ",", "."))
                        Call M46V999.gpGraLog(0, vbCrLf & Replace(".Formula = " & Chr(34) & .Formula & Chr(34) & vbCrLf & _
                                                                  ".Cod_Classe = " & .COD_CLASSE & vbCrLf & _
                                                                  ".Val_Fator_Franq = " & .Val_Fator_Franq & vbCrLf & _
                                                                  ".Alterar_Cobert = " & Chr(34) & .Alterar_Cobert & Chr(34) & vbCrLf & _
                                                                  ".Val_Risco_Cob = " & .Val_Risco_Cob & vbCrLf & _
                                                                  ".Nom_Cobertura = " & Chr(34) & .Nom_Cobertura & Chr(34) & vbCrLf & _
                                                                  ".Cod_Grupo_Cobertura = " & .Cod_Grupo_Cobertura & vbCrLf & _
                                                                  "End with " & vbCrLf, ",", "."))
                        Call M46V999.gpGraLog(0, vbCrLf & ".ColecaoPedCobert.add lobjstcPedCobert, lobjstcPedCobert.class_id" & vbCrLf)
                    End With
                Next lobjstcPedCobert

                Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .ColecaoPedLocBem = New Collection" & vbCrLf, ",", "."))
                For Each lobjstcPedBem In .ColecaoPedLocBem
                    With lobjstcPedBem
                        Call M46V999.gpGraLog(0, vbCrLf & Replace("set lobjstcPedBem = new stcA46V715B" & vbCrLf & "with lobjstcPedBem" & vbCrLf & _
                                                                  ".Nosso_Numero = " & .NOSSO_NUMERO & vbCrLf & _
                                                                  ".Num_Ped = " & .Num_Ped & vbCrLf & _
                                                                  ".Cod_Org = " & .Cod_Org & vbCrLf & _
                                                                  ".Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                                  ".Num_Apol = " & .Num_Apol & vbCrLf & _
                                                                  ".Num_Endo = " & .Num_Endo & vbCrLf & _
                                                                  ".Num_Item = " & .Num_Item & vbCrLf & _
                                                                  ".Dsc_Bem = " & .Dsc_Bem & vbCrLf & _
                                                                  ".Val_Bem = " & .Val_Bem & vbCrLf & _
                                                                  ".Cod_Bem = " & .Cod_Bem & vbCrLf & _
                                                                  ".Cod_Seq = " & .Cod_Seq & vbCrLf & _
                                                                  "end with" & vbCrLf, ",", "."))
                        Call M46V999.gpGraLog(0, vbCrLf & ".ColecaoPedLocBem.add lobjstcPedBem, lobjstcPedBem.class_id" & vbCrLf)
                    End With
                Next lobjstcPedBem
                Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .ColecaoPedSeguro = New Collection" & vbCrLf, ",", "."))
                For Each lobjstcPedOutroSeg In .ColecaoPedSeguro
                    With lobjstcPedOutroSeg
                        Call M46V999.gpGraLog(0, vbCrLf & Replace("set lobjstcPedOutroSeg = new stcA46V716B" & vbCrLf & "With lobjstcPedOutroSeg" & vbCrLf & _
                                                                  ".Nosso_Numero = " & .NOSSO_NUMERO & vbCrLf & _
                                                                  ".Num_Ped = " & .Num_Ped & vbCrLf & _
                                                                  ".Cod_Org = " & .Cod_Org & vbCrLf & _
                                                                  ".Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                                  ".Num_Apol = " & .Num_Apol & vbCrLf & _
                                                                  ".Num_Endo = " & .Num_Endo & vbCrLf & _
                                                                  ".Num_Item = " & .Num_Item & vbCrLf & _
                                                                  ".Dsc_Outro_Seguro = " & .Dsc_Outro_Seguro & vbCrLf & _
                                                                  ".Cod_Seguradora = " & .Cod_Seguradora & vbCrLf & _
                                                                  ".Num_Ordem = " & .Num_Ordem & vbCrLf & _
                                                                  ".Num_Apol_Ant = " & .Num_Apol_Ant & vbCrLf & _
                                                                  ".Num_Item_Ant = " & .Num_Item_Ant & vbCrLf & _
                                                                  ".Ind_Sinistro = " & .Ind_Sinistro & vbCrLf & _
                                                                  "End With" & vbCrLf, ",", "."))
                        Call M46V999.gpGraLog(0, vbCrLf & ".ColecaoPedSeguro.add lobjstcPedOutroSeg, lobjstcPedOutroSeg.class_id" & vbCrLf)
                    End With
                Next lobjstcPedOutroSeg

                Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .ColecaoPedFator = New Collection", ",", "."))
                For Each lobjstcFator In .ColecaoPedFator
                    With lobjstcFator
                        Call M46V999.gpGraLog(0, vbCrLf & Replace("set lobjstcFator    = new stcA46V720B" & vbCrLf & "With lobjstcFator" & vbCrLf & _
                                                                  ".Nosso_Numero = " & .NOSSO_NUMERO & vbCrLf & _
                                                                  ".Num_Ped = " & .Num_Ped & vbCrLf & _
                                                                  ".Cod_Org = " & .Cod_Org & vbCrLf & _
                                                                  ".Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                                  ".Num_Apol = " & .Num_Apol & vbCrLf & _
                                                                  ".Num_Endo = " & .Num_Endo & vbCrLf & _
                                                                  ".Num_Item = " & .Num_Item & vbCrLf & _
                                                                  ".Cod_Fator = " & .Cod_Fator & vbCrLf & _
                                                                  ".Dsc_Compl_Fator = " & .Dsc_Compl_Fator & vbCrLf & _
                                                                  ".Num_PI = " & .NUM_PI & vbCrLf & _
                                                                  "End With" & vbCrLf, ",", "."))
                        Call M46V999.gpGraLog(0, vbCrLf & ".ColecaoPedFator.add  lobjstcFator, lobjstcFator.class_id" & vbCrLf)
                    End With
                Next lobjstcFator

                With .objstcPedObs
                    Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .objstcPedObs = New stcA46V714B" & vbCrLf & "With .objstcPedObs" & vbCrLf & _
                                                              ".Nosso_Numero = " & .NOSSO_NUMERO & vbCrLf & _
                                                              ".Num_Item = " & .Num_Item & vbCrLf & _
                                                              ".Observacao = " & Chr(34) & .Observacao & Chr(34) & vbCrLf & _
                                                              "End With" & vbCrLf, ",", "."))
                End With
                With .objstcPedInsp
                    Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .objstcPedInsp = New stcA46V717B" & vbCrLf & "With .objstcPedInsp" & vbCrLf & _
                                                              ".Nosso_Numero = " & Chr(34) & .NOSSO_NUMERO & Chr(34) & vbCrLf & _
                                                              ".Num_Ped = " & .Num_Ped & vbCrLf & _
                                                              ".Cod_Org = " & .Cod_Org & vbCrLf & _
                                                              ".Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                              ".Num_Item = " & .Num_Item & vbCrLf & _
                                                              ".Num_Insp = " & .Num_Insp & vbCrLf & _
                                                              ".Dat_Cadastro = " & .Dat_Cadastro & vbCrLf & _
                                                              ".Pess_Cont = " & Chr(34) & .Pess_Cont & Chr(34) & vbCrLf & _
                                                              ".Num_Tel = " & Chr(34) & .Num_Tel & Chr(34) & vbCrLf & _
                                                              ".Num_Logr_Risco = " & .Num_Logr_Risco & vbCrLf & _
                                                              "End With" & vbCrLf, ",", "."))
                End With
                Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .ColecaoPedClau = New Collection" & vbCrLf, ",", "."))
                For Each lobjstcPedClau In .ColecaoPedClau
                    With lobjstcPedClau
                        Call M46V999.gpGraLog(0, vbCrLf & Replace("set lObjstcPedClau  = new stcA46V703B" & "With lObjstcPedClau" & vbCrLf & _
                                                                  ".Cod_Clau = " & .Cod_Clau & vbCrLf & _
                                                                  ".Cod_Org = " & .Cod_Org & vbCrLf & _
                                                                  ".Cod_Ramo = " & .Cod_Ramo & vbCrLf & _
                                                                  ".Nosso_Numero = " & .NOSSO_NUMERO & vbCrLf & _
                                                                  ".Num_Apol = " & .Num_Apol & vbCrLf & _
                                                                  ".Num_Endo = " & .Num_Endo & vbCrLf & _
                                                                  ".Num_Item = " & .Num_Item & vbCrLf & _
                                                                  ".Num_Ped = " & .Num_Ped & vbCrLf & _
                                                                  ".Num_PI = " & .NUM_PI & vbCrLf, ",", "."))
                        Call M46V999.gpGraLog(0, vbCrLf & "End With" & vbCrLf & ".ColecaoPedClau.add lObjstcPedClau, lObjstcPedClau.class_id " & vbCrLf)
                    End With
                Next lobjstcPedClau

                Call M46V999.gpGraLog(0, vbCrLf & Replace("Set .ColecaoNormas = New Collection" & vbCrLf, ",", "."))
                For Each lstcNorma In .ColecaoNormas
                    With lstcNorma
                        Call M46V999.gpGraLog(0, vbCrLf & Replace("set lstcNorma = new stcA46V708N" & vbCrLf & "With lstcNorma" & vbCrLf & _
                                                                  ".COD_ERR_NORMA = " & .COD_ERR_NORMA & vbCrLf & _
                                                                  ".COD_STAT_IRREG = " & .COD_STAT_IRREG & vbCrLf & _
                                                                  ".DSC_ERR_NORMA_EXT = " & Chr(34) & .DSC_ERR_NORMA_EXT & Chr(34) & vbCrLf & _
                                                                  ".DSC_ERR_NORMA_INT = " & Chr(34) & .DSC_ERR_NORMA_INT & Chr(34) & vbCrLf & _
                                                                  ".Num_Item = " & .Num_Item & vbCrLf & _
                                                                  "End With" & vbCrLf, ",", "."))
                        Call M46V999.gpGraLog(0, vbCrLf & ".ColecaoNormas.add lstcNorma,lstcNorma.class_id")
                    End With
                Next lstcNorma
            End With
            Call M46V999.gpGraLog(0, vbCrLf & "End With" & vbCrLf & ".ColecaoPedLoc.add lobjstcPedLoc ,lobjstcPedLoc.class_id" & vbCrLf)
        Next lobjstcPedLoc
    End With


FechaObj:
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    
    Set lobjstcPedLoc = Nothing
    Set lobjstcPedClau = Nothing
    Set lobjstcPedCobert = Nothing
    Set lobjstcPedCorr = Nothing
    Set lobjstcPedCoss = Nothing
    Set lobjstcPedJuros = Nothing
    Set lobjstcPedParc = Nothing
    Set lobjstcPedBem = Nothing
    Set lobjstcPedEspec = Nothing
    Set lobjstcFator = Nothing
    Set lobjstcPedOutroSeg = Nothing
    Set lstcNorma = Nothing
End Sub

Private Sub mpPosPrazo(ByVal pInicio As Double, ByVal pTermino As Double, ByRef pPrazo As String, ByRef pTipoPrazo As e_TipoPrazo)
'Procedure: posiciona prazo.
'********************Vari�veis padr�o*************************************
    Dim lCkPoint As String    'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lCkStr As String    'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
'lCkPoint. (posi��o das variaveis em ordem de declara��o)
'********************Vari�veis Declaradas para o processamento************
    Dim lAuxQuaMeses As Double    'Auxiliar para c�lculo.
    Dim lDatInicio As Double    'Data de in�cio.
    Dim lDatTermino As Double    'Data de t�rmino.
    Dim lDesPrazo As Boolean    'Indicador para desabilitar prazo.
    Dim lQuaDias As Double    'Quantidade de dias.
    Dim lQuaMeses As Double    'Quantidade de meses.


    On Error GoTo ControleErro
    Err = 0
    lCkPoint = ""
    lCkStr = ""

    '1. Se data inv�lida, encerra processo.
    lCkPoint = "1"
    lDesPrazo = False
    If pInicio = 0 Then
        lDesPrazo = True
    Else
        If pTermino = 0 Then
            lDesPrazo = True
        Else
            If M46V999.gfVerData(1, pInicio) = False Then
                lDesPrazo = True
            Else
                If M46V999.gfVerData(1, pTermino) = False Then
                    lDesPrazo = True
                Else
                    lDatInicio = pInicio
                    lDatTermino = pTermino
                    If lDatInicio > lDatTermino Then
                        lDesPrazo = True
                    End If
                End If
            End If
        End If
    End If
    lCkPoint = "2"
    If lDesPrazo = True Then        'Desabilita prazo.
        pPrazo = " "
        pTipoPrazo = e_TipoPrazo_ND
        Exit Sub
    End If
    lCkPoint = "3"
    '2. Calcula prazo de vig�ncia:
    'a) Dias.
    lQuaDias = M46V999.gfDifDatas(lDatInicio, lDatTermino)
    If lQuaDias = 366 And Mid(lDatInicio, 5, 4) = Mid(lDatTermino, 5, 4) Then
        lQuaDias = 365
    End If
    lCkPoint = "4"
    'b) Meses.
    lAuxQuaMeses = lQuaDias / 30
    lQuaMeses = Int(lAuxQuaMeses)
    lAuxQuaMeses = lAuxQuaMeses - lQuaMeses
    If lAuxQuaMeses > 0.5 Then
        lQuaMeses = lQuaMeses + 1
    End If
    lCkPoint = "5"
    '3. Posiciona informa��es.
    If lQuaDias <= 365 Then
        Select Case lQuaDias
        Case 0
            pPrazo = "0"
        Case 1
            pPrazo = "1 dia"
        Case 366
            pPrazo = "365 dias"
        Case Else
            pPrazo = lQuaDias & " dias"
        End Select

        'Alterado Sandra 19/08/2009 - Para ENDOSSO verificar aqui
        If lQuaDias < 365 Then
            pTipoPrazo = e_TipoPrazo_PrazoCurto
        Else
            pTipoPrazo = e_TipoPrazo_ProRata
        End If
    Else
        Select Case lQuaMeses
        Case 0
            pPrazo = "0"
        Case 1
            pPrazo = "1 m�s"
        Case Else
            pPrazo = lQuaMeses & " meses"
        End Select
        pTipoPrazo = e_TipoPrazo_PrazoLongo
    End If
    lCkPoint = "6"
    Exit Sub
ControleErro:
    Call M46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    'Obtem o inicio da mensagem de erro
    lCkStr = "mpPosPrazo ERRO - " & Err & " " & ERROR & _
             Chr(13) & "CheckPoint " & lCkPoint & " - "
    'Carrega a variavel lCkStr
    On Error Resume Next
    lCkStr = lCkStr & " | " & lAuxQuaMeses
    lCkStr = lCkStr & " | " & lDatInicio
    lCkStr = lCkStr & " | " & lDatTermino
    lCkStr = lCkStr & " | " & lDesPrazo
    lCkStr = lCkStr & " | " & lQuaDias
    lCkStr = lCkStr & " | " & lQuaMeses
    On Error GoTo 0
    Call M46V999.gpGraLog(0, lCkStr)
End Sub

Public Function gfDescPrazo(ByVal pDataIni As Double, ByVal pDataTer As Double) As String
    Dim lDescPrazo As String
    Dim lDummy As e_TipoPrazo

    Call mpPosPrazo(pDataIni, pDataTer, lDescPrazo, lDummy)
    'Retorno
    gfDescPrazo = lDescPrazo

End Function

Public Function gfTipoPrazo(ByVal pDataIni As Double, ByVal pDataTer As Double) As e_TipoPrazo
    Dim lTipoPrazo As e_TipoPrazo
    Dim lDummy As String

    Call mpPosPrazo(pDataIni, pDataTer, lDummy, lTipoPrazo)
    'Retorno
    gfTipoPrazo = lTipoPrazo

End Function

Private Function mfPosUniProProdutor(ByRef pStcPed As stcA46V702B, ByRef pMensagem As String) As Boolean
'Procedure: posiciona unidade produtiva e produtor.

'Par�metro de entrada...pCodCorretor...C�digo do corretor.
'Par�metro de sa�da.....pMensagem......Mensagem.

    Dim lCodProdutor As Integer    'C�digo do produtor.
    Dim lCodUniProdutiva As Integer    'C�digo da unidade produtiva.
    Dim lTotalRegistros As Integer    'Total de registros na tabela YTAB_UNID_CORR_PROD
    mfPosUniProProdutor = False

    If pStcPed.ColecaoPedCorr.Count = 0 Then
        Call M46V999.gpGraLog(1, " Registro padr�o recriado com sucesso. Por favor repita a opera��o.")
        Err.Raise 765, " Registro padr�o", "Registro padr�o recriado com sucesso. Por favor repita a opera��o."
    End If

    '2. Obt�m dados.
    If M46V999.gfObtUniProdutiva(pStcPed.ColecaoPedCorr(1).Cod_Corr, lCodUniProdutiva, lCodProdutor, pMensagem, lTotalRegistros) = False Then
        Exit Function
    End If

    '3. Posiciona informa��es.
    pStcPed.Cod_Unid_Prod = lCodUniProdutiva
    pStcPed.Cod_Prod = lCodProdutor

    mfPosUniProProdutor = True
End Function

Private Function mfCarregaNormas(ByRef pStcPed As stcA46V702B, ByRef pColNormas As Collection, tagDynaTrace As String) As Boolean
    Dim lstcNorma As stcA46V708N
    Dim i As Integer

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfCarregaNormas"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    On Error GoTo TrataErro

    If Not pColNormas Is Nothing Then
        For Each lstcNorma In pColNormas
            For i = 1 To pStcPed.ColecaoPedLoc.Count
                Call M46V999.gpGraLog(0, pStcPed.ColecaoPedLoc(i).Num_Item & "/" & lstcNorma.Num_Item)
                If pStcPed.ColecaoPedLoc(i).Num_Item = lstcNorma.Num_Item Then
                    Call M46V999.gpGraLog(0, "Norma: " & lstcNorma.COD_ERR_NORMA)
                    pStcPed.ColecaoPedLoc(i).ColecaoNormas.Add lstcNorma    ', lstcNorma.Class_ID  'Thais solicitou para n�o gravar mais o ID - p/ evitar erro quando retornar mais de uma vez o mesmo codigo de norma
                    Exit For
                End If
            Next i
        Next
    End If

    mfCarregaNormas = True

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    GoSub FechaObj
    Exit Function
TrataErro:
    UltimoErro = "mfCarregaNormas Erro:" & Err & " - " & ERROR
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA
FechaObj:
    Set lstcNorma = Nothing
    Return
End Function

Private Function mfSalvarTextoEspecPadrao(ByRef pStcPed As stcA46V702B, pMensagem As String, pGravaParametrizado As Boolean, tagDynaTrace As String) As Boolean
    Dim lobjTextoEspec As clsA46V726A
    Dim lobjstcTextoEspec As stcA46V726B
    Dim lobjstcCobert As stcA46V704B
    Dim lobjstcItem As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)Dim lNumPI                As Long
    Dim lBDP00Multi As ADODB.Connection    'Arquivo - P00Multi.
    Dim lrst As clsYasRecordSet
    Dim lrsTab_Ped_Espec As clsYasRecordSet
    Dim lTextoInc As String
    Dim lTotLin As Long
    Dim lSeq As Long
    Dim AuxArray As Variant
    Dim i As Integer
    Dim lNum_Pag As Integer
    Dim lLinha As Integer
    Dim lCorretor As Double
    Dim lCod_Clasf As String
    Dim lCod_Compl As String
    Dim lCod_Regiao As Double
    Dim lCod_Produto As Double
    Dim lMensagem As String
    Dim lRegiao As Integer
    Dim lAuxiliar As String
    Dim lAuxiliar1 As String
    Dim lSelect As String
    Dim lwColecaoTextoAtual As Collection
    Dim lwStcTextoEspec As stcA46V726B

    Dim lBeneficiariosTmp As String
    Dim blnTemCobertTodosRisc As Boolean
    Dim lobjstcCoss As stcA46V706B

    On Error GoTo ControleErro

    blnTemCobertTodosRisc = False
    'Guarda os Benefici�rios (Preenchidos na tela-Web)
    lBeneficiariosTmp = pStcPed.Beneficiarios

    Set lobjstcCoss = New stcA46V706B

    Set lobjTextoEspec = New clsA46V726A
    Set lobjTextoEspec.M46V999 = M46V999

    Set lobjstcTextoEspec = New stcA46V726B
    
    Dim clsA46V101A As clsA46V101A
    Set clsA46V101A = New clsA46V101A
    Set clsA46V101A.M46V999 = M46V999
    
    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(M46V999.gDataCorteSompo) Then
            If CLng(M46V999.gDataCorteSompo) <= M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If
    
    '    If pStcPed.ColecaoTextoEspec.Count > 0 Then
    lobjstcTextoEspec.NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
    'Vai excluir tudo pois essa rotina vai salvar al�m
    ' dos benefici�rios as demais cl�usulas
    ' lobjstcTextoEspec.Cod_Espec = 8500 e
    ' as franquias das coberturas 56, 57 e 60
    If Not lobjTextoEspec.mfExcluir(lobjstcTextoEspec, objstcUsuario.PropTabela, "", tagDynaTrace) Then
        pMensagem = lobjTextoEspec.UltimoErro
        GoTo SAIDA
    End If
    '    End If

    'Inicializa a Cole��o de textos de especifica��o
    Set lwColecaoTextoAtual = pStcPed.ColecaoTextoEspec
    Set pStcPed.ColecaoTextoEspec = New Collection

    lSelect = "Select max(Num_Pag) as Nro_Pag from " & objstcUsuario.PropTabela & "Tab_Ped_Espec WITH (NOLOCK) where NOSSO_NUMERO = '" & pStcPed.NOSSO_NUMERO & "'"
    If M46V999.gfObtRegistro(lobjTextoEspec.ConexaoBD, lSelect, lrsTab_Ped_Espec, pMensagem) Then
        If Not lrsTab_Ped_Espec.EOF Then
            lNum_Pag = IIf(IsNull(lrsTab_Ped_Espec("Nro_Pag")), 0, lrsTab_Ped_Espec("Nro_Pag"))
        End If
    End If

    lNum_Pag = lNum_Pag + 1

    'Inclus�o texto especifica��o susep inicio
    Set lobjstcTextoEspec = New stcA46V726B
    With lobjstcTextoEspec
        .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
        .Cod_Org = 0
        .Cod_Ramo = M46V999.gCod_Ramo
        .Num_Ped = pStcPed.Num_Ped
        .Num_Apol = 0
        .Num_Endo = 0
        .Cod_Espec = 6500
        .Num_pag = lNum_Pag
        .Qtd_vias = 3

        'segmentado
'            If pStcped.Cod_Ramo = e_CodRamo_Empresarial Then
'                If pStcped.Cod_Estip < 1000 Then
'                    .Espec_lin1 = "Nr Proc Susep - 15414.901792/2013-70 (para coberturas de"
'                    .Espec_lin2 = "Responsabilidade Civil)."
'                    .Espec_lin3 = "Nr Proc Susep - 15414.003932/2006-13 (para coberturas de Lucros Cessantes)."
'                Else
'                    .Espec_lin1 = "Nr Proc Susep - 15414.901793/2013-14 (para coberturas de"
'                    .Espec_lin2 = "Responsabilidade Civil)."
'                    .Espec_lin3 = "Nr Proc Susep - 15414.003932/2006-13 (para coberturas de Lucros Cessantes)."
'                End If
'            End If
        If pStcPed.Cod_Ramo = e_CodRamo_Empresarial And pStcPed.getPlano.EhPlanoSegmentado = False Then
            .Espec_lin1 = lNomeSeguradora & " Empresarial - Processo SUSEP N. 15414.004605/2004-17;"
            .Espec_lin2 = "(Cob.Resp.Civil - 15414.901792/2013-70);" '(Cob.Cap.Global - 15414.003932/2006-13)"
            .Espec_lin3 = ";(Cob.Lucros Cessantes - 15414.003932/2006-13)"
        ElseIf pStcPed.getPlano.EhPlanoSegmentado Then
            .Espec_lin1 = lNomeSeguradora & " Empresarial " & pStcPed.getPlano.NomPlano
            .Espec_lin2 = " - Processo SUSEP PRINCIPAL N. " & pStcPed.gCodSusep
            .Espec_lin3 = "SECUND�RIOS: RC- 15414.901792/2013-70;LC- 15414.003932/2006-13"
        End If

        If pStcPed.Cod_Ramo = e_CodRamo_Condominio Then
            .Espec_lin1 = "Nr Proc Susep - 15414.901800/2013-88 (para coberturas de"
            .Espec_lin2 = "Responsabilidade Civil)."
            .Espec_lin3 = "Nr Proc Susep - 15414.900709/2013-45 (para coberturas de Capital Global)."
        End If

        If pStcPed.Cod_Ramo = e_CodRamo_Residencial Then
            .Espec_lin1 = "Nr Proc Susep - 15414.901794/2013-69 (para coberturas de"
            .Espec_lin2 = "Responsabilidade Civil)."
            .Espec_lin3 = "Nr Proc Susep - 15414.900709/2013-45 (para coberturas de Capital Global)."
        End If

        .Espec_lin4 = "SUSEP - superintend�ncia de seguros privados - autarquia federal"
        .Espec_lin5 = "respons�vel pela fiscaliza��o, normatiza��o e controle dos mercados de"
        .Espec_lin6 = "seguro, previd�ncia complementar aberta, capitaliza��o, resseguro e"
        .Espec_lin7 = "CORRETAGEM DE SEGUROS."
        .Espec_lin8 = "telefone de atendimento ao p�blico (susep)-0800 021 8484 das 9:30hs �s 17:00hs."
        
        Dim lVolta As Integer
        Dim lPerLider As Double

        If M46V999.gPrograma <> "D046O003" Then
        
            lVolta = 1
            lPerLider = 0
            For Each lobjstcCoss In pStcPed.ColecaoPedCoss
                lPerLider = lPerLider + lobjstcCoss.Per_Coss
            
                If lVolta = 1 Then
                    .Espec_lin10 = "C�d. Registro: " & lobjstcCoss.Cod_Coss & _
                                  " CNPJ: " & M46V999.gfObterCnpjCoss(lobjstcCoss.Cod_Coss)
                    .Espec_lin11 = " Companhia: " & Trim(lobjstcCoss.Nom_Seguradora) & _
                                   " %Part: " & lobjstcCoss.Per_Coss & "%"
                End If
    
                If lVolta > 1 Then
                    .Espec_lin12 = "C�d. Registro: " & lobjstcCoss.Cod_Coss & _
                                  " CNPJ: " & M46V999.gfObterCnpjCoss(lobjstcCoss.Cod_Coss)
                    .Espec_lin13 = " Companhia: " & Trim(lobjstcCoss.Nom_Seguradora) & _
                                   " %Part: " & lobjstcCoss.Per_Coss & "%"
                ElseIf lVolta > 2 Then
                    .Espec_lin14 = "C�d. Registro: " & lobjstcCoss.Cod_Coss & _
                                  " CNPJ: " & M46V999.gfObterCnpjCoss(lobjstcCoss.Cod_Coss)
                    .Espec_lin15 = " Companhia: " & Trim(lobjstcCoss.Nom_Seguradora) & _
                                   " %Part: " & lobjstcCoss.Per_Coss & "%"
                ElseIf lVolta > 3 Then
                    .Espec_lin16 = "C�d. Registro: " & lobjstcCoss.Cod_Coss & _
                                  " CNPJ: " & M46V999.gfObterCnpjCoss(lobjstcCoss.Cod_Coss)
                    .Espec_lin17 = " Companhia: " & Trim(lobjstcCoss.Nom_Seguradora) & _
                                   " %Part: " & lobjstcCoss.Per_Coss & "%"
                ElseIf lVolta > 4 Then
                    .Espec_lin18 = "C�d. Registro: " & lobjstcCoss.Cod_Coss & _
                                  " CNPJ: " & M46V999.gfObterCnpjCoss(lobjstcCoss.Cod_Coss)
                    .Espec_lin19 = " Companhia: " & Trim(lobjstcCoss.Nom_Seguradora) & _
                                   " %Part: " & lobjstcCoss.Per_Coss & "%"
                End If
    
                lVolta = lVolta + 1
            Next
            
            If pStcPed.ColecaoPedCoss.Count > 0 Then
                .Espec_lin9 = "Cosseguro: 572 - 61.383.493/0001-80 " & lNomeSeguradora & " Seguros S/A (Lider) " & 100 - lPerLider & "%"
            End If
        End If

        .NUM_PI = M46V111.gNum_Pi
    End With

    If Not lobjTextoEspec.mfIncluir(lobjstcTextoEspec, objstcUsuario.PropTabela, "", tagDynaTrace) Then
        pMensagem = lobjTextoEspec.UltimoErro
        GoTo SAIDA
    End If

    lNum_Pag = lNum_Pag + 1
    pStcPed.ColecaoTextoEspec.Add lobjstcTextoEspec, lobjstcTextoEspec.Class_ID
    'fim
    
    '02-07-14
    'beneficiarios (web n�o tem texto de clausula beneficiaria)
    For Each lobjstcItem In pStcPed.ColecaoPedLoc
    
        If lobjstcItem.ColecaoBeneficiario.Count > 0 Then
            lSeq = 0
            lTotLin = 1
    
            Set lobjstcTextoEspec = New stcA46V726B
    
            With lobjstcTextoEspec
                .NOSSO_NUMERO = lobjstcItem.NOSSO_NUMERO
                .Cod_Org = 0
                .Cod_Ramo = M46V999.gCod_Ramo
                .Num_Ped = lobjstcItem.Num_Ped
                .Num_Apol = 0
                .Num_Endo = 0
                .Cod_Espec = 8500
                .Num_pag = pStcPed.ColecaoTextoEspec.Count + 1
                .Qtd_vias = 3
    
                Dim lobjstcPessoa As stcA46V735B
    
                For i = 1 To lobjstcItem.ColecaoBeneficiario.Count
                    Set lobjstcPessoa = lobjstcItem.ColecaoBeneficiario.Item(i)
                    Select Case i
                    Case 1
                        .Espec_lin11 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
                    Case 2
                        .Espec_lin12 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
                    Case 3
                        .Espec_lin13 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
                    Case 4
                        .Espec_lin14 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
                    Case 5
                        .Espec_lin15 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
                    Case 6
                        .Espec_lin16 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
                    Case 7
                        .Espec_lin17 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
                    Case 8
                        .Espec_lin18 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
                    Case 9
                        .Espec_lin19 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
                    Case 10
                        .Espec_lin20 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
                    End Select
                Next i
    
                .NUM_PI = M46V111.gNum_Pi
            End With
    
                If Not lobjTextoEspec.mfIncluir(lobjstcTextoEspec, objstcUsuario.PropTabela, "", "") Then
                pMensagem = lobjTextoEspec.UltimoErro
                GoTo SAIDA
            End If
    
    
            pStcPed.ColecaoTextoEspec.Add lobjstcTextoEspec, lobjstcTextoEspec.Class_ID
            lNum_Pag = pStcPed.ColecaoTextoEspec.Count + 1

                If lobjstcItem.Tip_Emissao = e_TipEmissao_Endosso Then
            If Not clsA46V101A.mfComparareCalcular(lobjstcItem, lobjstcItem.ObjPai, pStcPed.gStcPedidoApolice, lMensagem, objstcUsuario.PropTabela, tagDynaTrace) Then
                Call M46V999.gpMenErro(lMensagem)
                GoTo SAIDA
            End If
                End If

        End If
    Next lobjstcItem
    
    'Guarda os Benefici�rios (Preenchidos na tela-Web) ANTIGO
    pStcPed.Beneficiarios = lBeneficiariosTmp
    
    If Not M46V999.gfPreenchido(pStcPed.Beneficiarios) Then
        GoTo SalvarFrmT46V166A
    End If
    
    lSeq = 0
    lTotLin = 1

    If InStr(pStcPed.Beneficiarios, vbCrLf) <> 0 Then
        AuxArray = Split(pStcPed.Beneficiarios, vbCrLf)
        lTextoInc = AuxArray(lSeq)
        lTotLin = UBound(AuxArray)
    Else
        lSeq = 1
        lTextoInc = pStcPed.Beneficiarios
    End If

    Set lobjstcTextoEspec = New stcA46V726B

    With lobjstcTextoEspec
        .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
        .Cod_Org = 0
        .Cod_Ramo = pStcPed.Cod_Ramo
        .Num_Ped = pStcPed.Num_Ped
        .Num_Apol = 0
        .Num_Endo = 0
        .Cod_Espec = 8500
        .Num_pag = lNum_Pag
        .Qtd_vias = 3
        For i = 1 To lTotLin + 1
            If i = 1 Then .Espec_lin1 = lTextoInc
            If i = 2 Then .Espec_lin2 = lTextoInc
            If i = 3 Then .Espec_lin3 = lTextoInc
            If i = 4 Then .Espec_lin4 = lTextoInc
            If i = 5 Then .Espec_lin5 = lTextoInc
            If i = 6 Then .Espec_lin6 = lTextoInc
            If i = 7 Then .Espec_lin7 = lTextoInc
            If i = 8 Then .Espec_lin8 = lTextoInc
            If i = 9 Then .Espec_lin9 = lTextoInc
            If i = 10 Then .Espec_lin10 = lTextoInc
            If i = 11 Then .Espec_lin11 = lTextoInc
            If i = 12 Then .Espec_lin12 = lTextoInc
            If lSeq < lTotLin And lSeq + 1 < 12 Then
                lSeq = lSeq + 1
                lTextoInc = AuxArray(lSeq)
            Else
                Exit For
            End If
        Next
        .Espec_lin20 = "frmT46V155A"
        .NUM_PI = M46V111.gNum_Pi
    End With

    If Not lobjTextoEspec.mfIncluir(lobjstcTextoEspec, objstcUsuario.PropTabela, "", tagDynaTrace) Then
        pMensagem = lobjTextoEspec.UltimoErro
        GoTo SAIDA
    End If

    lNum_Pag = lNum_Pag + 1
    

SalvarFrmT46V166A:

    ''    '=======================================================================================================
    ''    'FRH - FICHA 901577 - INSERIR ESPECIFICA��O AUTOM. PARA SEGURADO CASAS BAHIA E CORRETORA MARINE.
    ''    '=======================================================================================================
    '    Dim textoEspec As String
    '    Dim dizerOK As Boolean
        Dim lbdTab_Ped_Dizer_Pad   As ADODB.Connection
        Dim lrsTab_Ped_Dizer_Pad As clsYasRecordSet
    '    Dim A As Long
    '
    '    If pStcPed.Cod_Produto >= 974 Then
    '
    '      If pStcPed.objstcCtrlEmis.Cod_Corr Like "288911" And _
           '         Left(pStcPed.Num_Cgc_Cpf, 8) = "59291534" Then
    '
    '         If Not M46V999.gfDisponivel("frmT46V166A") Then Load frmT46V166A
    '      Else
    '        For A = 0 To frmT46V166A.flxEspec.Rows - 1
    '          If frmT46V166A.flxEspec.TextMatrix(A, 1) = "8690" Then
    '            If A > 1 Then
    '              frmT46V166A.flxEspec.RemoveItem (A)
    '            Else
    '              frmT46V166A.flxEspec.TextMatrix(A, 0) = ""
    '              frmT46V166A.flxEspec.TextMatrix(A, 1) = ""
    '              frmT46V166A.flxEspec.TextMatrix(A, 2) = ""
    '            End If
    '          End If
    '        Next
    '
    '        If Not M46V999.gfDisponivel("frmT46V166A") Then GoTo SalvarFranquias
    '      End If
    '
    '
    '      With frmT46V166A
    '
    '          'VERIFICAR C�DIGO DE CORRETOR E CNPJ DO SEGURADO
    '          If pStcPed.objstcCtrlEmis.Cod_Corr Like "288911" And _
               '             Left(pStcPed.Num_Cgc_Cpf, 8) = "59291534" Then
    '
    '              'VERIFICAR EXISTENCIA DO DIZER NA FLEX
    '              dizerOK = False
    '              For A = 0 To .flxEspec.Rows - 1
    '                If .flxEspec.TextMatrix(A, 1) = "8690" Then dizerOK = True
    '              Next
    '
    '              If Not dizerOK Then
    '
    '                'DIZER NAO ENCONTRADO, ACESSAR TAB_PED_DIZER_PAD
    '                If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTab_Ped_Dizer_Pad, lMensagem) = False _
                     '                Then Call M46V999.gpGraLog(2, lMensagem): GoTo SAIDA
    '
    '                lSelect = "SELECT * FROM TAB_PED_DIZER_PAD WHERE COD_DIZER_PAD = 8690"
    '
    '                If M46V999.gfObtRegistro(lbdTab_Ped_Dizer_Pad, lSelect, lrsTab_Ped_Dizer_Pad, lMensagem) = False _
                     '                Then Call M46V999.gpFecha2(lbdTab_Ped_Dizer_Pad): Call M46V999.gpGraLog(2, lMensagem): Screen.MousePointer = vbDefault: GoTo SAIDA
    '
    '                gCarga = True
    '
    '                If lrsTab_Ped_Dizer_Pad.EOF Then
    '                    MsgBox "Mensagem Padr�o 8690 n�o encontrada!", vbInformation
    '                Else
    '                    Do While Not lrsTab_Ped_Dizer_Pad.EOF
    '                        textoEspec = textoEspec & lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo") & vbCrLf
    '                        lrsTab_Ped_Dizer_Pad.MoveNext
    '                    Loop
    '
    '                    textoEspec = Mid(textoEspec, 1, Len(textoEspec) - 2)
    '                    textoEspec = RTrim(textoEspec)
    '                End If
    '
    '                'ADICIONAR O DIZER NA FLEX
    '                .flxEspec.AddItem ""
    '                If .flxEspec.Rows - 1 > 0 Then
    '                  If .flxEspec.TextMatrix(.flxEspec.Rows - 1, 0) <> "" Then
    '                    .flxEspec.TextMatrix(.flxEspec.Rows - 1, 0) = CStr(CInt(.flxEspec.TextMatrix(.flxEspec.Rows - 2, 0)) + 1)
    '                  Else
    '                    .flxEspec.TextMatrix(.flxEspec.Rows - 1, 0) = "0"
    '                  End If
    '                End If
    '
    '                .flxEspec.TextMatrix(.flxEspec.Rows - 1, 1) = M46V999.CDblx(8690)
    '                .flxEspec.TextMatrix(.flxEspec.Rows - 1, 2) = textoEspec
    '
    '                Call M46V999.gpFecha2(lbdTab_Ped_Dizer_Pad)
    '                gCarga = False
    '
    '              End If
    '
    '          End If
    '      End With
    '
    '    End If
    '
    '    '=======================================================================================================
    '    'FIM DAS ALTERA��ES DA F901577
    '    '=======================================================================================================
    '
    '    With frmT46V166A
    '
    '      For lLinha = 1 To .flxEspec.Rows - 1
    '        lSeq = 0
    '        lTotLin = 1
    '
    '        If M46V999.CDblx(.flxEspec.TextMatrix(lLinha, 1)) <> 0 Then
    '             lCod_Espec = .flxEspec.TextMatrix(lLinha, 1)
    '             If InStr(.flxEspec.TextMatrix(lLinha, 2), vbCrLf) <> 0 Then
    '                AuxArray = Split(.flxEspec.TextMatrix(lLinha, 2), vbCrLf)
    '                lTextoInc = AuxArray(lSeq)
    '                lTotLin = UBound(AuxArray)
    '             Else
    '                 lSeq = 1
    '                 lTextoInc = .flxEspec.TextMatrix(lLinha, 2)
    '             End If
    '
    '             Set lobjstcTextoEspec = New stcA46V726B
    '
    '             With lobjstcTextoEspec
    '                 .Nosso_Numero = pStcPed.Nosso_Numero
    '                 .Cod_Org = 0
    '                 .Cod_Ramo = M46V999.gCod_Ramo
    '                 .Num_Ped = pStcPed.Num_Ped
    '                 .Num_Apol = 0
    '                 .Num_Endo = 0
    '                 .Cod_Espec = lCod_Espec
    '                 .Num_pag = lNum_Pag
    '                 .Qtd_vias = 3
    '                 For I = 1 To lTotLin + 1
    '                     If I = 1 Then .Espec_lin1 = lTextoInc
    '                     If I = 2 Then
    '                        .Espec_lin2 = lTextoInc
    '                        If InStr(1, .Espec_lin1, "** FRANQUIA **", vbTextCompare) > 0 Or _
                             '                           InStr(1, .Espec_lin2, "** FRANQUIA **", vbTextCompare) > 0 Then
    '                            GoTo Proximo
    '                        End If
    '                     End If
    '                     If I = 3 Then .Espec_lin3 = lTextoInc
    '                     If I = 4 Then .Espec_lin4 = lTextoInc
    '                     If I = 5 Then .Espec_lin5 = lTextoInc
    '                     If I = 6 Then .Espec_lin6 = lTextoInc
    '                     If I = 7 Then .Espec_lin7 = lTextoInc
    '                     If I = 8 Then .Espec_lin8 = lTextoInc
    '                     If I = 9 Then .Espec_lin9 = lTextoInc
    '                     If I = 10 Then .Espec_lin10 = lTextoInc
    '                     If I = 11 Then .Espec_lin11 = lTextoInc
    '                     If I = 12 Then .Espec_lin12 = lTextoInc
    '                     If I = 13 Then .Espec_lin13 = lTextoInc
    '                     If I = 14 Then .Espec_lin14 = lTextoInc
    '                     If I = 15 Then .Espec_lin15 = lTextoInc
    '                     If I = 16 Then .Espec_lin16 = lTextoInc
    '                     If I = 17 Then .Espec_lin17 = lTextoInc
    '                     If I = 18 Then .Espec_lin18 = lTextoInc
    '                     If I = 19 Then .Espec_lin19 = lTextoInc
    '                     If I = 20 Then .Espec_lin20 = lTextoInc
    '                     If lSeq < lTotLin And lSeq + 1 < 20 Then
    '                         lSeq = lSeq + 1
    '                         lTextoInc = AuxArray(lSeq)
    '                     Else
    '                         Exit For
    '                     End If
    '                 Next
    '                 .Num_PI = M46V111.gNum_Pi
    '             End With
    '
    '             If Not lobjTextoEspec.mfIncluir(lobjstcTextoEspec) Then
    '                 pMensagem = lobjTextoEspec.UltimoErro
    '                 GoTo SAIDA
    '             End If
    '
    '             pStcPed.ColecaoTextoEspec.Add lobjstcTextoEspec, lobjstcTextoEspec.Class_ID
    '
    '             lNum_Pag = lNum_Pag + 1
    '
    '         End If
    'Proximo:
    '      Next
    '    End With

SalvarFranquias:

    'Ficha 901383
    'And M46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then

    If pStcPed.Cod_Produto >= 962 And pGravaParametrizado = True Then
        Set lobjstcTextoEspec = Nothing
        Set lobjstcCobert = Nothing

        Set lobjstcTextoEspec = New stcA46V726B

        Set lobjstcCobert = New stcA46V704B
        Set lobjstcCobert.M46V999 = M46V999

        If pStcPed.ColecaoPedCorr.Count > 0 Then
            lCorretor = pStcPed.ColecaoPedCorr.Item(1).Cod_Corr
        Else
            lCorretor = 0
        End If


        For Each lobjstcItem In pStcPed.ColecaoPedLoc
            'Verifica beneficiarios
'            If lobjstcItem.ColecaoBeneficiario.Count > 0 And lobjstcItem.Cod_Iden_Tarif >= 1080 Then
'                With lobjstcTextoEspec
'                    .NOSSO_NUMERO = lobjstcItem.NOSSO_NUMERO
'                    .Cod_Org = 0
'                    .Cod_Ramo = lobjstcItem.Cod_Ramo
'                    .Num_Ped = lobjstcItem.Num_Ped
'                    .Num_Apol = 0
'                    .Num_Endo = 0
'                    .Cod_Espec = 8500
'                    .Num_pag = pStcped.ColecaoTextoEspec.Count + 1
'                    .Qtd_vias = 3
'
'                    Dim lobjstcPessoa As stcA46V735B
'
'                    For i = 1 To lobjstcItem.ColecaoBeneficiario.Count
'                        Set lobjstcPessoa = lobjstcItem.ColecaoBeneficiario.Item(i)
'                        Select Case i
'                        Case 1
'                            .Espec_lin11 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
'                        Case 2
'                            .Espec_lin12 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
'                        Case 3
'                            .Espec_lin13 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
'                        Case 4
'                            .Espec_lin14 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
'                        Case 5
'                            .Espec_lin15 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
'                        Case 6
'                            .Espec_lin16 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
'                        Case 7
'                            .Espec_lin17 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
'                        Case 8
'                            .Espec_lin18 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
'                        Case 9
'                            .Espec_lin19 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
'                        Case 10
'                            .Espec_lin20 = lobjstcPessoa.Nom_Pessoa & IIf(lobjstcPessoa.Tip_Pessoa = e_TipPessoa_Fisica, " CPF:", " CNPJ:") & lobjstcPessoa.Num_CPF_CNPJ
'                        End Select
'                    Next i
'
'                    .NUM_PI = lobjstcItem.NUM_PI
'                End With
'
'                If Not lobjTextoEspec.mfIncluir(lobjstcTextoEspec, objstcUsuario.PropTabela, "", tagDynaTrace) Then
'                    pMensagem = lobjTextoEspec.UltimoErro
'                    GoTo SAIDA
'                End If
'                pStcped.ColecaoTextoEspec.Add lobjstcTextoEspec, lobjstcTextoEspec.Class_ID
'
'                lNum_Pag = pStcped.ColecaoTextoEspec.Count + 1
'            End If
            

            If pStcPed.ColecaoPedLoc.Count > 0 Then
                lCod_Clasf = lobjstcItem.Cod_Clasf
                lCod_Compl = lobjstcItem.Cod_Compl_Clasf
                lCod_Regiao = lobjstcItem.gCod_RegiaoNova
                lCod_Produto = lobjstcItem.Cod_Produto
            End If

            If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, _
                                     lBDP00Multi, lMensagem) = False Then
                Call M46V999.gpGraLog(1, lMensagem)
                GoTo SAIDA
            End If

            'Ficha 1000284
            Dim lValFranquia As Double

            For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
                If lobjstcCobert.Cod_Cobert = 27 Then blnTemCobertTodosRisc = True
                lValFranquia = 0
                Call mpBuscaLimFranquia_Cobert(lobjstcCobert, lValFranquia)

                '1001135 - Corretor Volks - Franquia Especial Concessionaria
                If pStcPed.Cod_Produto >= 1002 Then
                    If (InStr(1, "00008001/00008002/00008003/00008004/00008005/00008006/00008008/00008011/00008014/00008028", Format(pStcPed.objstcCtrlEmis.Cod_Corr, "00000000"), vbTextCompare) > 0) And _
                       (val(lobjstcItem.Cod_Clasf) = 22 And val(lobjstcItem.Cod_Compl_Clasf) = 13) And _
                       (InStr(1, "0038/0064/0066", Format(lobjstcCobert.Cod_Cobert, "0000"), vbTextCompare) > 0) Then
                        GoTo PulaGravacao
                    Else
                        If lValFranquia > 0 And lValFranquia <> lobjstcCobert.Val_Franq_Min Then
                            GoTo PulaGravacao
                        End If
                    End If
                Else
                    If lValFranquia > 0 And lValFranquia <> lobjstcCobert.Val_Franq_Min Then
                        GoTo PulaGravacao
                    End If
                End If

                'Ficha 1000655
                lRegiao = 0

                'Vai procurar na tabela de Ajustes pontuais se tem uma regi�o definida com franquia diferenciada
                If mfFranquiaRegiao(e_Tip_Valor_Regiao, TpoValor, pStcPed.Cod_Ramo, pStcPed.Cod_Estip, lobjstcItem.Cod_Clasf, lobjstcItem.Cod_Compl_Clasf, lobjstcCobert.Cod_Cobert, pStcPed.Cod_Produto, lobjstcItem.Cod_Constr, lobjstcItem.Num_Cep_Risco, lRegiao, lMensagem, tagDynaTrace) = False Then
                    GoTo SAIDA
                End If


                lAuxiliar = ";" & pStcPed.Cod_Estip & "/" & lRegiao & "/" & Trim(lobjstcItem.Cod_Clasf) & "/" & Trim(lobjstcItem.Cod_Compl_Clasf) & "/" & lCorretor & "/" & lobjstcCobert.Cod_Cobert & ";"
                If lAuxiliar1 Like "*" & lAuxiliar & "*" Then
                    GoTo PulaGravacao
                End If

                lAuxiliar1 = lAuxiliar1 & lAuxiliar

                'Consulta por Corretor + Classifica��o
                lSelect = " Select * from TAB_DESC_FRANQUIA WITH (NOLOCK) Where"
                lSelect = lSelect & " Cod_Ramo = " & pStcPed.Cod_Ramo
                lSelect = lSelect & " and Cod_Produto = " & pStcPed.Cod_Estip
                lSelect = lSelect & "  and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
                lSelect = lSelect & "  and Compl_Clasf = '" & Trim(lobjstcItem.Cod_Clasf) & Trim(lobjstcItem.Cod_Compl_Clasf) & "'"    '1000871 - Inclus�o de Org�o Categoria Conselho
                lSelect = lSelect & "  and (Cod_Corretor = " & lCorretor & " or Cod_Corretor = 0)"
                lSelect = lSelect & "  and (Cod_Cobert = " & lobjstcCobert.Cod_Cobert & ")"
                lSelect = lSelect & " and Cod_Tarf_Ini <= " & pStcPed.Cod_Produto
                lSelect = lSelect & " and Cod_Tarf_Final >= " & pStcPed.Cod_Produto
                lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"


                If M46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
                    Call M46V999.gpGraLog(1, lMensagem)
                    GoTo SAIDA
                End If

                If lrst.EOF Then

                    lAuxiliar = ";" & pStcPed.Cod_Estip & "/" & lRegiao & "/" & lCorretor & "/" & lobjstcCobert.Cod_Cobert & ";"
                    If lAuxiliar1 Like "*" & lAuxiliar & "*" Then
                        GoTo PulaGravacao
                    End If

                    lAuxiliar1 = lAuxiliar1 & lAuxiliar

                    'Consulta por Corretor + todas as classifica��es
                    lSelect = " Select * from TAB_DESC_FRANQUIA WITH (NOLOCK) Where"
                    lSelect = lSelect & " Cod_Ramo = " & pStcPed.Cod_Ramo
                    lSelect = lSelect & " and Cod_Produto = " & pStcPed.Cod_Estip
                    lSelect = lSelect & "  and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
                    lSelect = lSelect & "  and (Compl_Clasf ='0')"
                    lSelect = lSelect & "  and (Cod_Corretor = " & lCorretor & ")"
                    lSelect = lSelect & "  and (Cod_Cobert = " & lobjstcCobert.Cod_Cobert & ")"
                    lSelect = lSelect & " and Cod_Tarf_Ini <= " & pStcPed.Cod_Produto
                    lSelect = lSelect & " and Cod_Tarf_Final >= " & pStcPed.Cod_Produto
                    lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"

                    If M46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
                        Call M46V999.gpGraLog(1, lMensagem)
                        GoTo SAIDA
                    End If
                End If

                If lrst.EOF Then

                    lAuxiliar = ";" & pStcPed.Cod_Estip & "/" & lRegiao & "/" & "/" & lobjstcCobert.Cod_Cobert & ";"
                    If lAuxiliar1 Like "*" & lAuxiliar & "*" Then
                        GoTo PulaGravacao
                    End If

                    lAuxiliar1 = lAuxiliar1 & lAuxiliar

                    'Consulta por todos os corretores + todas as classifica��es
                    lSelect = " Select * from TAB_DESC_FRANQUIA WITH (NOLOCK) Where"
                    lSelect = lSelect & " Cod_Ramo = " & pStcPed.Cod_Ramo
                    lSelect = lSelect & " and Cod_Produto = " & pStcPed.Cod_Estip
                    lSelect = lSelect & " and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
                    lSelect = lSelect & " and (Compl_Clasf ='0')"
                    lSelect = lSelect & " and (Cod_Corretor = 0)"
                    lSelect = lSelect & " and (Cod_Cobert = " & lobjstcCobert.Cod_Cobert & ")"
                    lSelect = lSelect & " and Cod_Tarf_Ini <= " & pStcPed.Cod_Produto
                    lSelect = lSelect & " and Cod_Tarf_Final >= " & pStcPed.Cod_Produto
                    lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"

                    If M46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
                        Call M46V999.gpGraLog(1, lMensagem)
                        GoTo SAIDA
                    End If
                End If

                If lrst.EOF Then GoTo PulaGravacao

                Do While Not lrst.EOF
                    If lrst("Qtde_linha") = 1 Then
                        Set lobjstcTextoEspec = New stcA46V726B
                    End If
                    With lobjstcTextoEspec
                        .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
                        .Cod_Org = 0
                        .Cod_Ramo = M46V999.gCod_Ramo
                        .Num_Ped = pStcPed.Num_Ped
                        .Num_Apol = 0
                        .Num_Endo = 0
                        .Cod_Espec = 1
                        .Num_pag = lNum_Pag
                        .Qtd_vias = 3
                        For i = 1 To 20
                            If i = lrst("Qtde_Linha") Then
                                If i = 1 Then .Espec_lin1 = lrst("Descricao")
                                If i = 2 Then .Espec_lin2 = lrst("Descricao")
                                If i = 3 Then .Espec_lin3 = lrst("Descricao")
                                If i = 4 Then .Espec_lin4 = lrst("Descricao")
                                If i = 5 Then .Espec_lin5 = lrst("Descricao")
                                If i = 6 Then .Espec_lin6 = lrst("Descricao")
                                If i = 7 Then .Espec_lin7 = lrst("Descricao")
                                If i = 8 Then .Espec_lin8 = lrst("Descricao")
                                If i = 9 Then .Espec_lin9 = lrst("Descricao")
                                If i = 10 Then .Espec_lin10 = lrst("Descricao")
                                If i = 11 Then .Espec_lin11 = lrst("Descricao")
                                If i = 12 Then .Espec_lin12 = lrst("Descricao")
                                If i = 13 Then .Espec_lin13 = lrst("Descricao")
                                If i = 14 Then .Espec_lin14 = lrst("Descricao")
                                If i = 15 Then .Espec_lin15 = lrst("Descricao")
                                If i = 16 Then .Espec_lin16 = lrst("Descricao")
                                If i = 17 Then .Espec_lin17 = lrst("Descricao")
                                If i = 18 Then .Espec_lin18 = lrst("Descricao")
                                If i = 19 Then .Espec_lin19 = lrst("Descricao")
                                If i = 20 Then .Espec_lin20 = lrst("Descricao")
                            End If
                        Next
                        .NUM_PI = M46V111.gNum_Pi
                    End With
                    lrst.MoveNext
                Loop

                If Not lobjTextoEspec.mfIncluir(lobjstcTextoEspec, objstcUsuario.PropTabela, "", tagDynaTrace) Then
                    pMensagem = lobjTextoEspec.UltimoErro
                    GoTo SAIDA
                End If

                pStcPed.ColecaoTextoEspec.Add lobjstcTextoEspec, lobjstcTextoEspec.Class_ID

                lNum_Pag = lNum_Pag + 1

                If lNum_Pag = 100 Then
                    mfSalvarTextoEspecPadrao = True
                    GoTo SAIDA
                End If
PulaGravacao:
            Next lobjstcCobert
        Next lobjstcItem
    End If


    '1101105 - Texto Especifica��o LMI
'    If pStcPed.LMI_Unico = 1 Then
'        If Not M46V103.mfAdicionaTextoEspecificacao(8725, lobjTextoEspec, pStcPed, objstcUsuario.PropTabela, pMensagem, tagDynaTrace) Then
'            Call M46V999.gpGraLog(1, pMensagem)
'            GoTo ControleErro
'        End If
'    End If

    If pStcPed.LMI_Unico = 1 Then
        Set lobjstcTextoEspec = New stcA46V726B
        With lobjstcTextoEspec
            .NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
            .Cod_Org = 0
            .Cod_Ramo = M46V999.gCod_Ramo
            .Num_Ped = pStcPed.Num_Ped
            .Num_Apol = 0
            .Num_Endo = 0
            .Cod_Espec = 8725   'cadastrado no main-frame
            .Num_pag = lNum_Pag
            .Qtd_vias = 3
            '1200029 - Clausula 315 e Texto LMI
            'abre base de dados
            If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTab_Ped_Dizer_Pad, lMensagem) = False Then
                pMensagem = lMensagem
                GoTo ControleErro
            End If
            'consulta mensagem
            lSelect = "SELECT * FROM TAB_PED_DIZER_PAD WHERE COD_DIZER_PAD = " & .Cod_Espec
            lSelect = lSelect & " order by COD_DIZER_PAD , num_seq "
            If M46V999.gfObtRegistro(lbdTab_Ped_Dizer_Pad, lSelect, lrsTab_Ped_Dizer_Pad, lMensagem) = False Then
                pMensagem = lMensagem
                Call M46V999.gpFecha2(lbdTab_Ped_Dizer_Pad)
                GoTo ControleErro
            End If
            If lrsTab_Ped_Dizer_Pad.RecordCount = 0 Then pMensagem = "Mensagem Padr�o " & .Cod_Espec & " n�o encontrada!": GoTo ControleErro
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
            Call M46V999.gpFecha2(lbdTab_Ped_Dizer_Pad)
            .NUM_PI = M46V111.gNum_Pi
        End With
        If Not pStcPed.FindTextoEspec(lobjstcTextoEspec) Then
            If Not lobjTextoEspec.mfIncluir(lobjstcTextoEspec, objstcUsuario.PropTabela, "", "") Then
                pMensagem = lobjTextoEspec.UltimoErro
                GoTo SAIDA
            End If
            lNum_Pag = lNum_Pag + 1
            pStcPed.ColecaoTextoEspec.Add lobjstcTextoEspec, lobjstcTextoEspec.Class_ID
        End If
    End If

    '1200429 - Tarifa 1036 e 1037 Maio 2012 - Item 9
    If (pStcPed.Cod_Produto >= 1036) And _
       (pStcPed.Cod_Ramo = e_CodRamo_Residencial Or _
        pStcPed.Cod_Ramo = e_CodRamo_Empresarial Or _
        pStcPed.Cod_Ramo = e_CodRamo_Condominio) Then
        If Not M46V103.mfAdicionaTextoEspecificacao(8726, lobjTextoEspec, pStcPed, objstcUsuario.PropTabela, pMensagem, tagDynaTrace) Then
            Call M46V999.gpGraLog(1, pMensagem)
            GoTo ControleErro
        End If
    End If

    '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 9
    If (pStcPed.Cod_Ramo = e_CodRamo_Residencial And blnTemCobertTodosRisc = True) Then
        If (pStcPed.Cod_Produto >= 1038 And pStcPed.Cod_Produto <= 1041) Or _
           (pStcPed.Cod_Produto >= 1048) Then    '1201062 - Tarifa 1048 e 1049 item 1
            If Not M46V103.mfAdicionaTextoEspecificacao(8727, lobjTextoEspec, pStcPed, objstcUsuario.PropTabela, pMensagem, tagDynaTrace) Then
                Call M46V999.gpGraLog(1, pMensagem)
                GoTo ControleErro
            End If
        End If
    End If

    '1200806 - Tarifa 1044 e 1045
    If (pStcPed.Cod_Produto >= 1044) And _
       pStcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        For Each lobjstcItem In pStcPed.ColecaoPedLoc
            For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
                If lobjstcCobert.Cod_Cobert = 82 Then
                    If Not M46V103.mfAdicionaTextoEspecificacao(8728, lobjTextoEspec, pStcPed, objstcUsuario.PropTabela, pMensagem, tagDynaTrace) Then
                        Call M46V999.gpGraLog(1, pMensagem)
                        GoTo ControleErro
                    End If
                End If
                If lobjstcCobert.Cod_Cobert = 160 Then
                    If Not M46V103.mfAdicionaTextoEspecificacao(8729, lobjTextoEspec, pStcPed, objstcUsuario.PropTabela, pMensagem, tagDynaTrace) Then
                        Call M46V999.gpGraLog(1, pMensagem)
                        GoTo ControleErro
                    End If
                End If
            Next
        Next
    End If

    If pStcPed.Cod_Bco_Cobr = e_Banco_DoBrasil Then
        If Not M46V103.mfAdicionaTextoEspecificacao(6504, lobjTextoEspec, pStcPed, objstcUsuario.PropTabela, pMensagem, tagDynaTrace) Then
            Call M46V999.gpGraLog(1, pMensagem)
            GoTo ControleErro
        End If
    End If

    'Ajuste inclus�o texto circular 491
    If Not M46V103.mfAdicionaTextoEspecificacao(6110, lobjTextoEspec, pStcPed, objstcUsuario.PropTabela, pMensagem, tagDynaTrace) Then
        Call M46V999.gpGraLog(1, pMensagem)
        GoTo ControleErro
    End If

    If pStcPed.objstcCtrlEmis.Ind_Cond_Constituido = 1 And pStcPed.Cod_Produto >= 1100 Then
        If Not M46V103.mfAdicionaTextoEspecificacao(8730, lobjTextoEspec, pStcPed, objstcUsuario.PropTabela, pMensagem, tagDynaTrace) Then
            Call M46V999.gpGraLog(1, pMensagem)
            GoTo ControleErro
        End If
    End If

    'Grava os textos que estavam na cole��o original e nao foram gravados nessa rotina
    For Each lwStcTextoEspec In lwColecaoTextoAtual
        If Not pStcPed.FindTextoEspec(lwStcTextoEspec) Then
            If lwStcTextoEspec.Cod_Espec <> 6504 And lwStcTextoEspec.Cod_Espec <> 6110 And lwStcTextoEspec.Cod_Espec <> 8730 Then
                'lNum_Pag = lNum_Pag + 1
                'lwStcTextoEspec.Num_pag = lNum_Pag
                lwStcTextoEspec.Num_pag = lobjTextoEspec.Colecao.Count + 1
            If Not lobjTextoEspec.mfIncluir(lwStcTextoEspec, objstcUsuario.PropTabela, "", tagDynaTrace) Then
                pMensagem = lobjTextoEspec.UltimoErro
                GoTo SAIDA
            End If
            pStcPed.ColecaoTextoEspec.Add lwStcTextoEspec, lwStcTextoEspec.Class_ID
        End If
        End If
    Next

    mfSalvarTextoEspecPadrao = True
SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    Set lobjTextoEspec = Nothing
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call M46V999.gpGraLog(0, "mfSalvarTextoEspecifica��o ERRO - " & Err & " " & Err.Description & _
                             Chr(13) & pMensagem)

    GoTo SAIDA
FechaObj:
    Set lobjTextoEspec = Nothing
    Set lobjstcTextoEspec = Nothing
    Set lobjstcCobert = Nothing
    Set lobjstcItem = Nothing
    Set lBDP00Multi = Nothing
    Set lrst = Nothing
    Set lrsTab_Ped_Espec = Nothing
    Set lwColecaoTextoAtual = Nothing
    Set lwStcTextoEspec = Nothing
    Set lobjstcCoss = Nothing
    Set clsA46V101A = Nothing
    
    Return
End Function

Private Function mfSalvarNovosTextoCoberturasDaniel(ByRef pStcPed As stcA46V702B, _
                                                    pMensagem As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfSalvarNovosTextoCoberturasDaniel"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lobjTexto As clsA46V739A
    Dim lobjstcTexto As stcA46V739B
    Dim lobjstcCobert As stcA46V704B
    Dim lobjstcItem As stcA46V708B
    Dim ColecaoTextosTemp As Collection
    Dim lTotLin As Long
    Dim i As Integer
    Dim lMensagem As String

    On Error GoTo ControleErro

    Set lobjTexto = New clsA46V739A
    Set lobjTexto.M46V999 = M46V999
    
    Set lobjstcTexto = New stcA46V739B
    Set ColecaoTextosTemp = New Collection
    
    lobjstcTexto.NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
    If Not lobjTexto.mfExcluir(lobjstcTexto, objstcUsuario.PropTabela, "") Then
        pMensagem = lobjTexto.UltimoErro
        GoTo SAIDA
    End If
    
    'Monta a lista de coberturas e faz a chamada p/ a classe M46V999
    Dim strCoberturas As String
    Dim strClassificacoes As String
    Dim pSaida() As Variant

    For Each lobjstcItem In pStcPed.ColecaoPedLoc
        For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
            strCoberturas = strCoberturas & lobjstcCobert.Cod_Cobert & ","
        Next lobjstcCobert
    Next lobjstcItem

    'Luiz Alberto - 20/10/2015
    If Not M46V999.gfObterTextosCobertura(pStcPed.Cod_Ramo, M46V111.gCod_plano, strCoberturas, strClassificacoes, strClassificacoes, lobjstcItem.Tip_Cobert, pSaida, pMensagem) Then
        GoTo SAIDA
    End If
    
    lTotLin = UBound(pSaida)

    For Each lobjstcItem In pStcPed.ColecaoPedLoc
        For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
            Set lobjstcTexto = New stcA46V739B
            lobjstcTexto.NOSSO_NUMERO = pStcPed.NOSSO_NUMERO
            lobjstcTexto.Cod_Org = 0
            lobjstcTexto.Cod_Ramo = M46V999.gCod_Ramo
            lobjstcTexto.Num_Ped = pStcPed.Num_Ped
            lobjstcTexto.Num_Apol = pStcPed.Num_Apol
            lobjstcTexto.Num_Endo = pStcPed.Num_Endo
            lobjstcTexto.NUM_PI = pStcPed.NUM_PI
        
            lobjstcTexto.Num_Item = lobjstcItem.Num_Item
            lobjstcTexto.Cod_Cobert = lobjstcCobert.Cod_Cobert
            lobjstcTexto.Tipo_Texto = 1
            For i = 1 To lTotLin
                If (pSaida(i, 6) = lobjstcCobert.Cod_Cobert) Then
                    lobjstcTexto.Cod_Texto = pSaida(i, 10)
                    lobjstcTexto.Desc_Texto = pSaida(i, 11)
                    Exit For
                End If
            Next i
            
            If (lobjstcTexto.Cod_Texto <> 0) Then
                If Not lobjTexto.mfIncluir(lobjstcTexto, objstcUsuario.PropTabela, "", tagDynaTrace) Then
                    pMensagem = lobjTexto.UltimoErro
                    GoTo SAIDA
                End If
                ColecaoTextosTemp.Add lobjstcTexto, lobjstcTexto.Class_ID
            End If
        Next lobjstcCobert
        Set lobjstcItem.ColecaoTextos = ColecaoTextosTemp
    Next lobjstcItem

    mfSalvarNovosTextoCoberturasDaniel = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    
SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    Set lobjTexto = Nothing
    Set lobjstcTexto = Nothing
    
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call M46V999.gpGraLog(0, "mfSalvarNovosTextoCoberturas  ERRO - " & Err & " " & Err.Description & _
                             Chr(13) & pMensagem)

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
    GoTo SAIDA
FechaObj:
    Set lobjTexto = Nothing
    Set lobjstcTexto = Nothing
    Set lobjstcCobert = Nothing
    Set lobjstcItem = Nothing
    Return
End Function

    

Private Function mfFranquiaRegiao(ByVal pCod_Tip_Valor As Integer, ByVal pTip_Dado As TipoDados, ByVal pCod_Ramo As Integer, ByVal pCod_Prod As Integer, ByRef pCod_Clasf As String, ByRef pCod_Compl_Clasf As String, ByVal pCod_Cobert As Integer, ByVal pCod_Tarifa As Long, ByVal pCod_Constr As Integer, ByVal pCepRisco As Long, ByRef ValorAjustavel As Variant, ByRef pMensagem As String, tagDynaTrace As String) As Boolean
    Dim lbdTbm_Parameter As ADODB.Connection      'Arquivo: P00Tarifa.
    Dim lrsTbm_Parameter As clsYasRecordSet       'Registro : Tbm_Parameter.
    Dim lSelect As String

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfFranquiaRegiao"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    On Error GoTo ControleErros

    pMensagem = ""
    mfFranquiaRegiao = False

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Parameter, pMensagem) = False Then
        GoTo SAIDA
    End If

    lSelect = " Select Cod_Cobert, Cod_Tip_Chave, Chave_Ini, Chave_Ter, Valor from Tbm_Parameter WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & pCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & pCod_Prod
    lSelect = lSelect & " and Cod_Corr in (0, " & M46V999.gCod_Corretor & ")"
    lSelect = lSelect & " and ((rtrim(ltrim(Cod_Clasf)) = '" & Trim(pCod_Clasf) & "'"
    lSelect = lSelect & " and Cod_Compl_Clasf = '" & pCod_Compl_Clasf & "') or rtrim(ltrim(Cod_Clasf)) = '0' )"
    lSelect = lSelect & " and Cod_Cobert in (0," & pCod_Cobert & ")"
    lSelect = lSelect & " and Cod_Tarif = " & pCod_Tarifa
    lSelect = lSelect & " and Cod_Tip_Valor = " & pCod_Tip_Valor
    lSelect = lSelect & " and Cod_Constr in (0," & pCod_Constr & ")"
    lSelect = lSelect & " order by Cod_Cobert, Cod_Prioridade, Cod_Constr desc "

    If M46V999.gfObtRegistro(lbdTbm_Parameter, lSelect, lrsTbm_Parameter, pMensagem) = False Then
        GoTo SAIDA
    End If

    Do While Not lrsTbm_Parameter.EOF
        If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 3 Then
            If M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ini"))) <= pCepRisco And M46V999.CDblx(M46V999.Check(lrsTbm_Parameter("Chave_Ter"))) >= pCepRisco Then
                ValorAjustavel = M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        lrsTbm_Parameter.MoveNext
    Loop

    mfFranquiaRegiao = True
    Call M46V999.gpFecha2(lbdTbm_Parameter)

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    GoSub FechaObj
    Exit Function
ControleErros:
    mfFranquiaRegiao = False
    pMensagem = "mfFranquiaRegiao ERRO - " & Err & " " & Err.Description & _
                vbCrLf & " SQL: " & lSelect
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, eSeveridadeErro_Medio, Class_Name, methodName)
    Call M46V999.gpGraLog(0, pMensagem)
    GoTo SAIDA
FechaObj:
    Call M46V999.gpFecha2(lbdTbm_Parameter)
    Set lbdTbm_Parameter = Nothing
    Set lrsTbm_Parameter = Nothing
    Return
End Function

Private Sub mpBuscaLimFranquia_Cobert(ByRef pCodCobert As stcA46V704B, ByRef ValFraqMin As Double)
    Dim lbdTbm_Fat_Franquia As ADODB.Connection         'Arquivo: P00Multi.
    Dim lMensagem As String         'Mensagem.
    Dim lrsTbm_Fat_Franquia As clsYasRecordSet          'Registro : Tbm_Fator_Franquia.
    Dim lCodSequencia As Integer
    Dim lCod_Constr As Integer
    'Dim lTip_Cobert         As Integer
    Dim lSelect As String

    'Ficha 1000305
    If M46V999.gCodProduto >= 1080 Then
            Select Case pCodCobert.ObjPai.Cod_Constr
            Case 1
                lCod_Constr = e_CodContr_NvTarifa_Superior
            Case 2
                lCod_Constr = e_CodContr_NvTarifa_Solida
            Case 3
                lCod_Constr = e_CodContr_NvTarifa_Mista
            Case 4
                lCod_Constr = e_CodContr_NvTarifa_Inferior
            End Select
    Else
    If pCodCobert.Cod_Ramo = e_CodRamo_Empresarial Then
        Select Case pCodCobert.ObjPai.Cod_Constr
        Case 1
            lCod_Constr = e_CodContr_Superior
        Case 2
            lCod_Constr = e_CodContr_Inferior
        Case 3
            lCod_Constr = e_CodContr_Mista
        Case Else
            lCod_Constr = e_CodContr_Superior
        End Select
    Else
        If pCodCobert.ObjPai.Cod_Constr = e_CodContr_Inferior Or pCodCobert.ObjPai.Cod_Constr = e_CodContr_Mista Then
            lCod_Constr = e_CodContr_Inferior
        Else
            lCod_Constr = e_CodContr_Superior
        End If
    End If
    End If

    ''Ficha 1000305
    ''== Tipo Cobertura
    '    If pCodCobert.Cod_Ramo = e_CodRamo_Empresarial Then
    '        Select Case pCodCobert.ObjPai.Cod_Constr
    '            Case 0
    '                lTip_Cobert = e_TipCobert_PredioeConteudo
    '            Case 1
    '                lTip_Cobert = e_TipCobert_Predio
    '            Case 2
    '                lTip_Cobert = e_TipCobert_Conteudo
    '            Case Else
    '                lTip_Cobert = e_TipCobert_PredioeConteudo
    '        End Select
    '    Else
    '        lTip_Cobert = 0
    '    End If
    '    M46V999.gTipCobertura = lTip_Cobert

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Fat_Franquia, lMensagem) = False Then
        Call M46V999.gpGraLog(2, lMensagem)
        GoTo SAIDA
    End If

    lCodSequencia = pCodCobert.ObjPai.ObjPai.Cod_Sequencia

NovaSequencia1:

    '1 - Procura Registro de exce��o com a chave completa
    'Procura o valor do Val_Franq_Min e Perc_Franq
    lSelect = " Select Val_Franq_Min as Min from Tbm_Fator_Franquia WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & pCodCobert.Cod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & pCodCobert.ObjPai.Cod_Plano  '   frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
    lSelect = lSelect & " and Cod_Tip_Reg = 'E' "
    If pCodCobert.ObjPai.ObjPai.Cod_Produto >= 940 Then
        lSelect = lSelect & " and ((Cod_Clasf = '" & pCodCobert.ObjPai.Cod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & pCodCobert.ObjPai.Cod_Compl_Clasf & "')"
        lSelect = lSelect & " or (Cod_Clasf = '0'))"
    Else
        lSelect = lSelect & " and Cod_Clasf = '" & pCodCobert.ObjPai.Cod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & pCodCobert.ObjPai.Cod_Compl_Clasf & "'"
    End If
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert.Cod_Cobert
    lSelect = lSelect & " and Cod_Construcao in (0 , " & lCod_Constr & ")"
    lSelect = lSelect & " and Cod_Tarif = " & pCodCobert.ObjPai.Cod_Iden_Tarif
    lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia

    If pCodCobert.ObjPai.ObjPai.Cod_Produto >= 940 Then
        lSelect = lSelect & " AND ((CEP_INICIO <= " & pCodCobert.ObjPai.Num_Cep_Risco
        lSelect = lSelect & " AND CEP_FIM >= " & pCodCobert.ObjPai.Num_Cep_Risco & " ) "
        lSelect = lSelect & " OR (CEP_INICIO =  0 ))"
        lSelect = lSelect & " order by Prioridade desc "
    End If

    If M46V999.gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Fat_Franquia)
        Call M46V999.gpGraLog(2, lMensagem)
        GoTo SAIDA
    End If

    '2 - Se n�o encontrar a chave completa vai procurar o padr�o
    If lrsTbm_Fat_Franquia.EOF Then
        If lCodSequencia > 0 Then
            lCodSequencia = lCodSequencia - 1

            GoTo NovaSequencia1
        Else
            GoTo ProcuraDefault
        End If
    End If

    If lrsTbm_Fat_Franquia(0) = 0 Then
        GoTo ProcuraDefault
    End If

    ValFraqMin = lrsTbm_Fat_Franquia("Min")

    GoTo FimBusca

ProcuraDefault:



    '3 - Procura o registro somente com Ramo, Modalidade, Cobertura e TipoAtividade
    'Procura o valor do Val_Franq_Min, Perc_Franq

    lCodSequencia = pCodCobert.ObjPai.ObjPai.Cod_Sequencia

NovaSequencia2:

    lSelect = " Select Val_Franq_Min as Min from Tbm_Fator_Franquia WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & pCodCobert.Cod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & pCodCobert.ObjPai.Cod_Plano
    lSelect = lSelect & " and Cod_Tip_Reg = 'D' "
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert.Cod_Cobert
    lSelect = lSelect & " and Tip_Atividade in (3," & pCodCobert.ObjPai.gTip_Atividade & ")"
    lSelect = lSelect & " and Cod_Tarif = " & pCodCobert.ObjPai.Cod_Iden_Tarif
    lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia

    If M46V999.gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
        Call M46V999.gpGraLog(2, lMensagem)
        GoTo SAIDA
    End If

    If lrsTbm_Fat_Franquia.EOF Then
        If lCodSequencia > 0 Then
            lCodSequencia = lCodSequencia - 1

            GoTo NovaSequencia2
        Else
            GoTo FimBusca
        End If
    End If

    If lrsTbm_Fat_Franquia("Min") = 0 Then
        GoTo FimBusca
    End If

    ValFraqMin = lrsTbm_Fat_Franquia("Min")
FimBusca:
    Call M46V999.gpFecha2(lbdTbm_Fat_Franquia)
    '1001135 - Corretor Volks - Franquia Especial Concessionaria
    If pCodCobert.ObjPai.ObjPai.Cod_Produto >= 1002 Then
        If (InStr(1, "00008001/00008002/00008003/00008004/00008005/00008006/00008008/00008011/00008014/00008028", Format(pCodCobert.ObjPai.ObjPai.objstcCtrlEmis.Cod_Corr, "00000000"), vbTextCompare) > 0) And _
           (val(pCodCobert.ObjPai.Cod_Clasf) = 22 And val(pCodCobert.ObjPai.Cod_Compl_Clasf) = 13) And _
           (InStr(1, "0038/0064/0066", Format(pCodCobert.Cod_Cobert, "0000"), vbTextCompare) > 0) Then
            ValFraqMin = 2000
        End If
    End If

SAIDA:
    GoSub FechaObj
    Screen.MousePointer = vbDefault
    Exit Sub
FechaObj:
    Call M46V999.gpFecha2(lbdTbm_Fat_Franquia)
    Set lbdTbm_Fat_Franquia = Nothing
    Set lrsTbm_Fat_Franquia = Nothing
    Return
End Sub

Private Function PerguntaQuestionario(pColAlertas As Collection, pIDQuestionario As String, Optional pNumItem As Long = 0) As stcConsistPreCalc
    Dim intLoop As Integer
    Dim lstcConsist As stcConsistPreCalc
    If Not pColAlertas Is Nothing Then
        For intLoop = 1 To pColAlertas.Count
            Set lstcConsist = pColAlertas(intLoop)
            If lstcConsist.Class_ID = Format(pNumItem, "000") & "-" & Format(pIDQuestionario, "000") Then
                Set PerguntaQuestionario = lstcConsist
                Exit For
            End If
        Next
    End If
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lstcConsist = Nothing
    Return
End Function

Private Sub mfRemoveQuestionario(ByVal IDQuest As E_IDConsist, ByRef pColAlertas As Collection, Optional ByVal pNumItem As Long = 0)
    Dim lstrIDQuest As String
    Dim lstcConsist As stcConsistPreCalc
    Dim i As Integer

    Err.Clear
    On Error GoTo SAIDA
    lstrIDQuest = Format(pNumItem, "000") & "-" & Format(IDQuest, "000")

    If Not pColAlertas Is Nothing Then
        For i = 1 To pColAlertas.Count
            Set lstcConsist = pColAlertas(i)
            If lstrIDQuest = lstcConsist.Class_ID Then
                Call pColAlertas.Remove(i)
                Exit For
            End If
        Next
    End If

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lstcConsist = Nothing
    Return
End Sub

Private Function mfGeraLogCalcInspecao(pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B, Optional pTransmissao As Boolean = False, Optional pDiretorio As String = "") As Boolean
    Dim lobjstcCobert As stcA46V704B    'Objeto da classe de coberturas

'Procedure utilizada para gerar um log contendo informa��es, para serem enviadas para yasuda
'em que o servidor da yasuda armazena estas informa��es na tabela [RAMOSDIVERSOS].[dbo].[Tab_Trans_Insp]

'Parametros de entrada com base na tabela
'*******  LAYOUT DA TABELA NO SERVIDOR YASUDA
'    [VERSAO_ARQUIVO] [varchar] (2)
'    [NOSSO_NUMERO] [varchar] (20)
'    [NUM_PED] [varchar] (6)
'    [NUM_ITEM] [varchar] (5)
'    [NUM_APOL_ANT] [varchar] (10)
'    [NOM_SEGURADO] [varchar] (60)
'    [NUM_TEL] [varchar] (14)
'    [NUM_CGC_CPF] [varchar] (14)
'    [COD_RAMO] [varchar] (3)
'    [COD_ORG] [varchar] (2)
'    [DAT_INI_VIG] [varchar] (8)
'    [DAT_TER_VIG] [varchar] (8)
'    [COD_TIP_LOGR_RISCO] [varchar] (3)
'    [NOM_LOGR_RISCO] [varchar] (50)
'    [DSC_COMPL_RISCO] [varchar] (20)
'    [NOM_BAIRRO_RISCO] [varchar] (30)
'    [NOM_CID_RISCO] [varchar] (30)
'    [SIG_UF_RISCO] [varchar] (2)
'    [NUM_CEP_RISCO] [varchar] (8)
'    [NUM_RISCO] [varchar] (8)
'    [DAT_TRANS] [varchar] (8)
'    [COD_IDEN_PESSOA] [varchar] (1)
'    [COD_UNID_PROD] [varchar] (4)
'    [COD_PROD] [varchar] (4)
'    [COD_CORR] [varchar] (7)
'    [COD_CLASF] [varchar] (4)
'    [COD_COMPL_CLASF] [varchar] (2)
'    [PESS_CONT] [varchar] (40)
'    [COD_PRODUTO] [varchar] (4)
'    [FLAG_TRANS] [varchar] (1)
'    [COD_COBERT] [varchar] (4)
'    [VAL_IS_COBERT] [varchar] (20)

'******
    Dim lCaminho As String    'Caminho onde os dados s�o salvos
    Dim lEndArq055 As String    'Endereco do arquivo LOG
    Dim lNumArq055 As String    'Numero ref. para abertura do arquivo obtido com "FreeFile"
    Dim lLCoduser As String    'Codigo do usuario de transmiss�o Numerico
    Dim lCodUsuario As String    'Codigo do usuario de transmiss�o ( formatado 6 digitos)
    Dim lMsg As String    'Mensagem
    Dim lCont As Long
    Dim lCod_Seq As Long    'Sequ�ncia da transmiss�o
    Dim lImprime As String

    Dim lCod_Cobert(250) As Integer
    Dim lIS_Cobert(250) As Long

    Dim lSql As String

    Dim lrst As clsYasRecordSet
    Dim lbdRamosDiversos As ADODB.Connection
    Dim lCgcCpf055 As String    'CGC/CPF formatado p/ arquivo .055 (suporta CNPJ alfanum�rico).


    Err = 0

    mfGeraLogCalcInspecao = False

    On Error GoTo TrataErro

    '1- Posiciona o caminho dos dados
    If pTransmissao = True Then
        lCaminho = pDiretorio
        '2-Verifica se existencia do log ref. ao mesmo calculo
        lEndArq055 = lCaminho & "*" & pobjstcPed.NOSSO_NUMERO & "_" & pobjstcPedLoc.Num_Item & ".055"
    Else
        If M46V999.gfPreenchido(M46V999.gEndSYASUpDt) = True Then
            'Caminho ods arquivos de atualiza��o do SYAS
            'Informado no arquivo SYASPORD.INI
            lCaminho = M46V999.gEndSYASUpDt
        Else
            'Endere�o das bases de dados tarifarias
            lCaminho = M46V999.gEndServidor
        End If
        If M46V999.gPrograma = "D046O003" Then
            lCaminho = M46V999.gAppPath
        End If
        '2-Verifica se existencia do log ref. ao mesmo calculo
        If Not M46V999.gfPreenchido(Dir(lCaminho & "\REMESSA", vbDirectory)) Then
            On Error Resume Next
            MkDir (lCaminho & "\REMESSA")
            If Err <> 0 Then
                Call M46V999.gpGraLog(2, "FTE0015 - Erro ao criar pasta " & lCaminho & "\REMESSA" & _
                                         " (pasta de remessa para arquivos de inspe��o - " & Err & " - " & _
                                         ERROR & ")." & vbLf & vbLf & "O sistema ser� encerrado.")
                GoTo SAIDA
            End If
            On Error GoTo TrataErro
        End If

        lEndArq055 = lCaminho & "REMESSA\" & "*" & pobjstcPed.NOSSO_NUMERO & "_" & pobjstcPedLoc.Num_Item & ".055"
    End If



    Call M46V999.gpGraLog(0, "Verifica se existencia do log ref. ao mesmo calculo " & lEndArq055)

    If M46V999.gfPreenchido(Dir(lEndArq055)) = True Then
        'exclui o ultimo arquivo se existir
        Kill lEndArq055

        Call M46V999.gpGraLog(0, "exclui o ultimo arquivo se existir " & lEndArq055)

    End If


    '3-Verificar se foi agendada inspe��o
    If pobjstcPedLoc.Insp_Agend = False Then GoTo SAIDA

    '4-Carrega as coberturas

    lCont = 0
    For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
        lCont = lCont + 1
        If lCont > 250 Then Exit For
        lCod_Cobert(lCont) = lobjstcCobert.Cod_Cobert
        lIS_Cobert(lCont) = Format(lobjstcCobert.Val_IS, "standard")
    Next lobjstcCobert

    '5-Obtem o codigo do usuario de transmiss�o
    lCodUsuario = mUsuarioTransmissao

    '5-Inicia a cria��o do arquivo

    lNumArq055 = FreeFile
    If pTransmissao = True Then
        lEndArq055 = lCaminho & lCodUsuario & "_" & pobjstcPed.NOSSO_NUMERO & "_" & pobjstcPedLoc.Num_Item & ".055"
    Else
        lEndArq055 = lCaminho & "REMESSA\" & lCodUsuario & "_" & pobjstcPed.NOSSO_NUMERO & "_" & pobjstcPedLoc.Num_Item & ".055"
    End If

    Call M46V999.gpGraLog(0, "Inicia a cria��o do arquivo " & lEndArq055)

    Open lEndArq055 For Output Access Write As #lNumArq055

    If IsNumeric(pobjstcPed.Num_Cgc_Cpf) Then
        lCgcCpf055 = Format(pobjstcPed.Num_Cgc_Cpf, "00000000000000")
    Else
        lCgcCpf055 = Right$(String(14, "0") & Trim(pobjstcPed.Num_Cgc_Cpf), 14)
    End If

    lImprime = "V2;" & pobjstcPedLoc.NOSSO_NUMERO & ";" & Format(pobjstcPedLoc.Num_Ped, "000000") & _
               ";" & Format(pobjstcPedLoc.Num_Item, "00000") & ";" & Format(pobjstcPedLoc.Num_Apol_Anterior, "0000000000") & _
               ";" & pobjstcPed.Nom_Segurado & String(60 - Len(Left(pobjstcPed.Nom_Segurado, 60)), " ") & _
               ";" & pobjstcPedLoc.objstcPedInsp.Num_Tel & String(14 - Len(Left(pobjstcPedLoc.objstcPedInsp.Num_Tel, 14)), " ") & _
               ";" & lCgcCpf055 & ";" & Format(pobjstcPed.Cod_Ramo, "000") & ";" & Format(pobjstcPed.Cod_Org, "00") & _
               ";" & pobjstcPed.Dat_Ini_Vig & ";" & pobjstcPed.Dat_Ter_Vig & ";" & pobjstcPedLoc.Cod_Tip_Logr_Risco & String(3 - Len(Left(pobjstcPedLoc.Cod_Tip_Logr_Risco, 3)), " ") & _
               ";" & pobjstcPedLoc.Nom_Logr_Risco & String(50 - Len(Left(pobjstcPedLoc.Nom_Logr_Risco, 50)), " ") & ";" & pobjstcPedLoc.Dsc_Compl_Risco & String(20 - Len(Left(pobjstcPedLoc.Dsc_Compl_Risco, 20)), " ") & _
               ";" & pobjstcPedLoc.Nom_Bairro_Risco & String(30 - Len(Left(pobjstcPedLoc.Nom_Bairro_Risco, 30)), " ") & _
               ";" & pobjstcPedLoc.Nom_Cid_Risco & String(30 - Len(Left(pobjstcPedLoc.Nom_Cid_Risco, 30)), " ") & ";" & pobjstcPedLoc.Sig_UF_Risco & ";" & Format(pobjstcPedLoc.Num_Cep_Risco, "00000000") & _
               ";" & Format(pobjstcPedLoc.objstcPedInsp.Num_Logr_Risco, "00000000") & ";" & Format$(Now, "yyyymmdd") & _
               ";" & pobjstcPed.Tip_Pessoa & ";" & Format(pobjstcPed.Cod_Unid_Prod, "0000") & _
               ";" & Format(pobjstcPed.Cod_Prod, "0000") & ";" & Format(pobjstcPed.objstcCtrlEmis.Cod_Corr, "0000000") & _
               ";" & pobjstcPedLoc.Cod_Clasf & String(4 - Len(Left(pobjstcPedLoc.Cod_Clasf, 4)), " ") & ";" & pobjstcPedLoc.Cod_Compl_Clasf & _
               ";" & pobjstcPedLoc.objstcPedInsp.Pess_Cont & String(40 - Len(Left(pobjstcPedLoc.objstcPedInsp.Pess_Cont, 40)), " ") & _
               ";" & Format(pobjstcPedLoc.Cod_Produto, "0000") & ";0"


    lCont = 0
    For lCont = 1 To 250
        If lCod_Cobert(lCont) = 0 Then Exit For
        lImprime = lImprime & ";" & Format(lCod_Cobert(lCont), "0000") & ";" & Format(lIS_Cobert(lCont), "00000000000000000.00")
    Next lCont

    Print #lNumArq055, lImprime
    Close #lNumArq055

    Call M46V999.gpGraLog(0, "Final da cria��o do arquivo " & lEndArq055)

    mfGeraLogCalcInspecao = True

SAIDA:
    GoSub FechaObj
    Exit Function
TrataErro:
    UltimoErro = "mfGeraLogCalcInspecao ERRO " & Err & " " & ERROR$
    GoTo SAIDA
FechaObj:
    Set lobjstcCobert = Nothing
    Set lrst = Nothing
    Set lbdRamosDiversos = Nothing
    Return
End Function

Private Function mfBuscaUsuarioTransmissao(ByRef pCodUsuTransmissao As Long, ByRef pMensagem As String, tagDynaTrace As String) _
        As Boolean
'Fun��o: busca usu�rio de transmiss�o.

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfBuscaUsuarioTransmissao"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

'Par�metros de sa�da...pCodUsuTransmissao...C�digo do usu�rio.
'                      pMensagem............Mensagem.

    Dim lCodUsuario As Long    'C�digo do usu�rio.
    Dim lEndSyasEdi_EXE As String    'Endere�o do arquivo SYASEDI.EXE.
    Dim lNumSyasEdi_EXE As Integer    'Utilizado com FreeFile.
    Dim lRegistro As String    'Registro.

    mfBuscaUsuarioTransmissao = False

    '1. Verifica arquivo SYASEDI.EXE.
    lEndSyasEdi_EXE = M46V999.gAppPath & "SYASEDI.EXE"
    If Not M46V999.gfPreenchido(Dir(lEndSyasEdi_EXE)) Then
        'Procuro pelo arquivo antigo
        lEndSyasEdi_EXE = M46V999.gAppPath & "Eapscript.INI"
        If Not M46V999.gfPreenchido(Dir(lEndSyasEdi_EXE)) Then
            pMensagem = "FTExxxx - Erro ao obter c�digo do usu�rio de transmiss�o (identifica��o n�o " & _
                        "localizada)." & vbLf & vbLf & "Reinicie o processo de c�lculo."
            Exit Function
        End If
    End If
    lNumSyasEdi_EXE = FreeFile
    On Error Resume Next
    Open lEndSyasEdi_EXE For Input Access Read As #lNumSyasEdi_EXE
    If Err <> 0 Then
        pMensagem = "FTExxxx - Erro ao obter c�digo do usu�rio de transmiss�o (" & Err & " - " & _
                    ERROR & ")." & vbLf & vbLf & "Reinicie o processo de c�lculo."
        Exit Function
    End If
    On Error GoTo 0

    '2. Obt�m c�digo do usu�rio.
    Do While Not EOF(lNumSyasEdi_EXE)
        Line Input #lNumSyasEdi_EXE, lRegistro
        If Mid(lRegistro, 1, 7) = "sender=" Then
            lCodUsuario = val(Mid(lRegistro, 9, 6))
            Exit Do
        End If
    Loop
    Close #lNumSyasEdi_EXE

Retorno:
    pCodUsuTransmissao = lCodUsuario
    mfBuscaUsuarioTransmissao = True

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

End Function

Private Property Get mUsuarioTransmissao() As String
    Dim lLCoduser As Long
    Static lsRetorno As String

    If M46V999.gfPreenchido(lsRetorno) = False Then
        If mfBuscaUsuarioTransmissao(lLCoduser, UltimoErro, M46V999.tagDynaTrace) = True Then
            lsRetorno = Format(lLCoduser, "000000")
        End If
    End If
    mUsuarioTransmissao = lsRetorno

End Property

Private Sub mpAtribuiClasseModulo(ByRef pStcPed As stcA46V702B, tagDynaTrace As String)
    Dim lContX As Integer
    Dim lContY As Integer

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mpAtribuiClasseModulo"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    On Error GoTo TrataErro
    
    '========***Importante passa os dados das Classes internas p/ o Objeto que veio de fora deste escopo...
    With pStcPed
        Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Atribuindo M46V999 ao pedido...", 0, False)
        Set .M46V999 = M46V999
        
        For lContX = 1 To .ColecaoPedLoc.Count
            Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Atribuindo ObjPai e M46V999 ao local...", 0, False)
            
            Set .ColecaoPedLoc(lContX).ObjPai = pStcPed
            Set .ColecaoPedLoc(lContX).M46V999 = M46V999
            
            For lContY = 1 To .ColecaoPedLoc(lContX).ColecaoPedClau.Count
                Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Atribuindo ObjPai e M46V999 a Clausula...", 0, False)
                Set .ColecaoPedLoc(lContX).ColecaoPedClau(lContY).ObjPai = .ColecaoPedLoc(lContX)
                Set .ColecaoPedLoc(lContX).ColecaoPedClau(lContY).M46V999 = M46V999
            Next lContY
            For lContY = 1 To .ColecaoPedLoc(lContX).ColecaoPedCobert.Count
                Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Atribuindo ObjPai e M46V999 a Cobertura...", 0, False)
                Set .ColecaoPedLoc(lContX).ColecaoPedCobert(lContY).ObjPai = .ColecaoPedLoc(lContX)
                Set .ColecaoPedLoc(lContX).ColecaoPedCobert(lContY).M46V999 = M46V999
            Next lContY
            For lContY = 1 To .ColecaoPedLoc(lContX).ColecaoPedLocBem.Count
                Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Atribuindo ObjPai e M46V999 ao Bem...", 0, False)
                Set .ColecaoPedLoc(lContX).ColecaoPedLocBem(lContY).ObjPai = .ColecaoPedLoc(lContX)
                Set .ColecaoPedLoc(lContX).ColecaoPedLocBem(lContY).M46V999 = M46V999
            Next lContY
        Next lContX

        For lContX = 1 To .ColecaoPedCoss.Count
            Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Atribuindo ObjPai e M46V999 ao Cosseguro...", 0, False)
            Set .ColecaoPedCoss(lContX).ObjPai = pStcPed
            Set .ColecaoPedCoss(lContX).M46V999 = M46V999
        Next lContX

        For lContX = 1 To .ColecaoPedCorr.Count
            Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Atribuindo ObjPai e M46V999 ao Corretor...", 0, False)
            Set .ColecaoPedCorr(lContX).ObjPai = pStcPed
            Set .ColecaoPedCorr(lContX).M46V999 = M46V999
        Next lContX

        For lContX = 1 To .ColecaoPedVenc.Count
            Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Atribuindo ObjPai e M46V999 ao Vencimento...", 0, False)
            Set .ColecaoPedVenc(lContX).ObjPai = pStcPed
            Set .ColecaoPedVenc(lContX).M46V999 = M46V999
        Next lContX

        Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Atribuindo ObjPai e M46V999 ao Endere�o...", 0, False)
        Set .objstcPedEndereco.ObjPai = pStcPed
        Set .objstcPedEndereco.M46V999 = M46V999
    End With

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Sub
    
TrataErro:
    UltimoErro = Err.Description
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function gfProEndosso(ByVal pNumRenApolice As String, ByVal pNumRenItem As Integer, _
                             ByVal pItemRenovado As Boolean, ByVal pCodDizer As Integer, _
                             ByRef pMensagensRetorno As String, tagDynaTrace As String) As stcA46V702B
'Fun��o: processa op��o de renova��o de seguro Yasuda.

    Dim lIndSinistro As e_SimNao
    Dim lNossoNumero As String
    Dim lStcPed As stcA46V702B
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfProEndosso"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Err.Clear: UltimoErro = ""
    On Error GoTo ControleErro

    Set gfProEndosso = Nothing

    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(M46V999.gDataCorteSompo) Then
            If CLng(M46V999.gDataCorteSompo) <= M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If

    If Not mfPosicionaUsuario(tagDynaTrace) Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro ao posicionar usu�rio", eSeveridadeErro_Medio, Class_Name, methodName)
        GoTo SAIDA
    End If

    '1. Consiste n�mero de ap�lice anterior.
    If Not M46V999.gfPreenchido(pNumRenApolice) Then
        pMensagensRetorno = "FGE0069 - N�mero da ap�lice inv�lido (Endosso " & lNomeSeguradora & " - n�o preenchido)."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagensRetorno, 0, Class_Name, methodName)
        GoTo SAIDA
    End If
    If Not IsNumeric(pNumRenApolice) Then
        pMensagensRetorno = "FGE0069 - N�mero da ap�lice inv�lido (Endosso " & lNomeSeguradora & " - n�o num�rico)."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagensRetorno, 0, Class_Name, methodName)
        GoTo SAIDA
    End If
    If Len(Trim(pNumRenApolice)) < 10 Then
        pNumRenApolice = Trim$(pNumRenApolice)
        pNumRenApolice = String(10 - Len(pNumRenApolice), "0") & pNumRenApolice
    End If


    '2. Consiste n�mero do item da ap�lice.
    If pCodDizer <> 1 Then
        If Not M46V999.gfPreenchido(pNumRenItem) Then
            pMensagensRetorno = "FGE0070 - N�mero do item da ap�lice inv�lido (Endosso " & lNomeSeguradora & " - n�o preenchido)."
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagensRetorno, 0, Class_Name, methodName)
            GoTo SAIDA
        End If
        If Not IsNumeric(pNumRenItem) Then
            pMensagensRetorno = "FGE0070 - N�mero do item da ap�lice inv�lido (Endosso " & lNomeSeguradora & " - n�o num�rico)."
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagensRetorno, 0, Class_Name, methodName)
            GoTo SAIDA
        End If
        If Len(Trim(pNumRenItem)) < 4 Then
            pNumRenItem = Trim$(pNumRenItem)
            pNumRenItem = String(4 - Len(pNumRenItem), "0") & pNumRenItem
        End If
        If (pCodDizer = 12 Or pCodDizer = 4) And pNumRenItem = 0 Then
            pMensagensRetorno = "FGE0070 - N�mero do item da ap�lice deve ser informado."
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagensRetorno, 0, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If

    If pCodDizer = 0 Then
        UltimoErro = "FGE0070 - Cod. Dizer n�o informado."
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 0, Class_Name, methodName)
        GoTo SAIDA
    End If

    'lIndSinistro = pIndSinistro
    lIndSinistro = 0
    '3. Grava registros de cota��o/proposta para renova��o Yasuda.
    Dim objVazio As stcA46V702B
    'Set objVazio = New stcA46V702B
    If M46V001.gfGraEndossoYasuda(pNumRenApolice, pNumRenItem, lNossoNumero, _
                                  lIndSinistro, pCodDizer, pMensagensRetorno, True, objVazio, tagDynaTrace) = False Then
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagensRetorno, 0, Class_Name, methodName)
        GoTo SAIDA
    End If

    Set lStcPed = gfProcuraNossoNumero(lNossoNumero, False, "", tagDynaTrace)

    Set gfProEndosso = lStcPed
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    UltimoErro = "[ERR] - gfProEndosso ERRO - " & Err & " " & Err.Description & UltimoErro
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, eSeveridadeErro_Medio, Class_Name, methodName)
    GoTo SAIDA
FechaObj:
    Set lStcPed = Nothing
    Return
End Function

Private Sub gfLog_Emissao(ByVal pNossoNumero As String, ByVal pNum_pi As Double, ByVal pMsg As String, ByVal pTipo As Enum_Tipo_LogEmissao)
    Dim lCnSyasPortal2008 As ADODB.Connection
    Dim lSelect As String
    Dim lErro As String
    Dim lServidor As String

    On Error GoTo Fim
    
    If Not M46V999.gfAbrBasDados("SYASPORTAL", M46V999.gTipSyasPortal, M46V999.gEndSyasPortal, lCnSyasPortal2008, UltimoErro) Then
        GoTo Fim
    End If
    
    If UBound(GetIpAddrTable) = 1 Then
        lServidor = GetIpAddrTable(1)
    Else
        lServidor = "0.0.0.0"
    End If
    
    pMsg = Replace(pMsg, "'", "''")
    lSelect = "Insert into Log_Emissao (NUM_PI, Nosso_Numero, Servidor, Msg, Tipo_Msg) "
    lSelect = lSelect & "SELECT " & pNum_pi & ", '" & pNossoNumero & "', '" & lServidor & "', '" & pMsg & "', " & CStr(pTipo)

    M46V999.gfExeSQL lCnSyasPortal2008, lSelect, lErro

Fim:
    Err.Clear
    Set lCnSyasPortal2008 = Nothing
End Sub

' Returns an array with the local IP addresses (as strings).
' Author: Christian d'Heureuse, www.source-code.biz
Public Function GetIpAddrTable()
   Dim Buf(0 To 511) As Byte
   Dim BufSize As Long: BufSize = UBound(Buf) + 1
   Dim rc As Long
   rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
   If rc <> 0 Then Err.Raise vbObjectError, , "GetIpAddrTable failed with return value " & rc
   Dim NrOfEntries As Integer: NrOfEntries = Buf(1) * 256 + Buf(0)
   If NrOfEntries = 0 Then GetIpAddrTable = Array(): Exit Function
   ReDim IpAddrs(0 To NrOfEntries - 1) As String
   Dim i As Integer
   For i = 0 To NrOfEntries - 1
      Dim j As Integer, S As String: S = ""
      For j = 0 To 3: S = S & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j): Next
      IpAddrs(i) = S
      Next
   GetIpAddrTable = IpAddrs
   End Function


'=================================================

Public Property Get M46V999() As clsM46V999
    If mClsM46V999 Is Nothing Then
        Set mClsM46V999 = New clsM46V999
        
        ' Habilita o log de inicio, para o caso de dar erro antes da leitura do INI
        mClsM46V999.gEnableFileLog = True
        
        Set mClsM46V999.M46O001 = M46O001
        Set mClsM46V999.M46V111 = M46V111
    
        Set mClsM46V999.DynaTrace = New clsMDynaTrace
        Call mClsM46V999.DynaTrace.Iniciar("VB-D46O003A-1080")
    End If

Retorno:
    Set M46V999 = mClsM46V999
End Property

Public Property Get M46O001() As clsM46O001
    If mClsM46O001 Is Nothing Then
        Set mClsM46O001 = New clsM46O001
    End If

Retorno:
    Set M46O001 = mClsM46O001
End Property

Public Property Get M46V103() As clsM46V103
    If mClsM46V103 Is Nothing Then
        Set mClsM46V103 = New clsM46V103
        Set mClsM46V103.M46V111 = M46V111
    End If

Retorno:
    Set M46V103 = mClsM46V103
End Property

Public Property Get M46V111() As clsM46V111
    If mClsM46V111 Is Nothing Then
        Set mClsM46V111 = New clsM46V111
        Set mClsM46V111.M46V999 = M46V999
        Set mClsM46V111.M99V001 = M99V001
        Set mClsM46V111.M46V103 = M46V103
    End If

Retorno:
    Set M46V111 = mClsM46V111
End Property

Public Property Get M46V998() As clsM46V998
    If mClsM46V998 Is Nothing Then
        Set mClsM46V998 = New clsM46V998
        Set mClsM46V998.M46V999 = M46V999
    End If

Retorno:
    Set M46V998 = mClsM46V998
End Property

Public Property Get M99V001() As clsM99V001
    If mClsM99V001 Is Nothing Then
        Set mClsM99V001 = New clsM99V001
        Set mClsM99V001.M46V999 = M46V999
    End If

Retorno:
    Set M99V001 = mClsM99V001
End Property

Public Property Get M46V001() As clsM46V001
    If mClsM46V001 Is Nothing Then
        Set mClsM46V001 = New clsM46V001
        Set mClsM46V001.M46V111 = M46V111
    End If

Retorno:
    Set M46V001 = mClsM46V001
End Property



'' 2017-11-30: SISCOTA - Consultar por Nosso Numero
Public Function gfCarregarSISCOTA(ByVal strNossoNumero As String, ByRef pMensagensRetorno As String) As stcA46V753B
    Dim objSis As clsA46V753A
    Dim pobjstc As New stcA46V753B
        
    On Error GoTo Erro
    
    pMensagensRetorno = ""
    'gfCarregarSISCOTA = New stcA46V753B
    
    Set objSis = New clsA46V753A
    Set objSis.M46V999 = Me.M46V999
    pobjstc.NOSSO_NUMERO = strNossoNumero
    
    If (objSis.mfCarregarMinimo(pobjstc, pMensagensRetorno)) Then
        Set gfCarregarSISCOTA = pobjstc
        GoTo Fim
    End If

Erro:
    pMensagensRetorno = "(gfObsSISCOTA) Erro: " & Err.Description
    UltimoErro = "[ERR] - gfObsSISCOTA " & Err & " " & Err.Description & "|" & UltimoErro
    
Fim:
    Set objSis = Nothing

End Function

'' 2017-10-20: Obs_SISCOTA - Consultar
Public Function gfObsSISCOTA(ByRef pobjstc As stcA46V753B, ByRef pMensagensRetorno As String) As Boolean
    Dim objObsSis As clsA46V753A
    
    On Error GoTo Erro
    
    pMensagensRetorno = ""
    gfObsSISCOTA = False
    
    Set objObsSis = New clsA46V753A
    Set objObsSis.M46V999 = Me.M46V999
    
    If (objObsSis.mfCarregarObs(pobjstc, pMensagensRetorno)) Then
         gfObsSISCOTA = True
         GoTo Fim
    End If

Erro:
    gfObsSISCOTA = False
    pMensagensRetorno = "(gfObsSISCOTA) Erro: " & Err.Description
    UltimoErro = "[ERR] - gfObsSISCOTA " & Err & " " & Err.Description & "|" & UltimoErro
    
Fim:
    Set objObsSis = Nothing

End Function


'' 2017-10-20: Obs_SISCOTA - Salvar
Public Function gfObsSISCOTA_Salvar(ByVal pobjstc As stcA46V753B, ByVal pLogin As String, _
                                    ByRef pMensagensRetorno As String) As Boolean
 
    Dim objObsSis As clsA46V753A
    
    On Error GoTo Erro
    
    pMensagensRetorno = ""
    gfObsSISCOTA_Salvar = False
    
    Set objObsSis = New clsA46V753A
    Set objObsSis.M46V999 = Me.M46V999
    
    If (objObsSis.mfSalvar(pobjstc, pLogin, pMensagensRetorno)) Then
         gfObsSISCOTA_Salvar = True
         GoTo Fim
    End If

Erro:
    gfObsSISCOTA_Salvar = False
    pMensagensRetorno = "(gfObsSISCOTA_Salvar) Erro: " & Err.Description
    UltimoErro = "[ERR] - gfObsSISCOTA_Salvar " & Err & " " & Err.Description & "|" & UltimoErro
    
Fim:
    Set objObsSis = Nothing

End Function


'' 2017-10-20: Obs_SISCOTA - Salvar
Public Function gfObsSISCOTA_Excluir(ByVal pobjstc As stcA46V753B, ByVal pLogin As String, _
                                     ByRef pMensagensRetorno As String) As Boolean
 
    Dim objObsSiscota As clsA46V750A
    Dim objObsSiscotaLog As clsA46V751A
    Dim objObsSiscotaAnexo As clsA46V752A
    
    On Error GoTo Erro
    
    pMensagensRetorno = ""
    gfObsSISCOTA_Excluir = False
    
    Set objObsSiscota = New clsA46V750A
    Set objObsSiscota.M46V999 = Me.M46V999
    
    If (objObsSiscota.mfExcluir(pobjstc.ColecaoObs(1), pLogin, pMensagensRetorno)) Then
         gfObsSISCOTA_Excluir = True
         GoTo Fim
    End If

Erro:
    gfObsSISCOTA_Excluir = False
    pMensagensRetorno = "(gfObsSISCOTA_Salvar) Erro: " & Err.Description
    UltimoErro = "[ERR] - gfObsSISCOTA_Salvar " & Err & " " & Err.Description & "|" & UltimoErro
Fim:
    Set objObsSiscota = Nothing

End Function

Public Function gfAtualizarStatusSiscota(ByVal pobjstc As stcA46V753B, _
                                         ByVal pLogin As String) As Boolean
 
    Dim objSiscota As clsA46V753A
    
    On Error GoTo Erro
    
    gfAtualizarStatusSiscota = False
    
    Set objSiscota = New clsA46V753A
    Set objSiscota.M46V999 = Me.M46V999
    
    If objSiscota.mfAtualizaStatusSiscota(pobjstc, pLogin) Then
         gfAtualizarStatusSiscota = True
         GoTo Fim
    End If
    
Erro:
    gfAtualizarStatusSiscota = False
    UltimoErro = "[ERR] - gfAtualizarStatusSiscota " & Err & " " & Err.Description & "|" & UltimoErro
    Set objSiscota = Nothing
    
Fim:
    Set objSiscota = Nothing

End Function

