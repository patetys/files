VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsM46V999"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long

' ======================================================================================================================================================
' Class Name: clsM46V999
' Sistema: SyasMult
' Analista: Marcio Oliveira
' Objetivo: Substitui o antigo m�dulo bas Adaptado para ser utilizado como classe gerenciavel.
' ======================================================================================================================================================

'MODULO COM FUN��ES E PROCEDURES COMUNS DO SYAS.

'AO ALTERAR ESTE M�DULO, COMPILAR E CRIAR EXECUT�VEL DE TODOS OS M�DULOS QUE O REFERENCIAM.

Public gCorrecaoMonetariaIPC As Boolean
Public gQuaErros As Long
Public gNumSusep As String
Public gCod_Corretor As Long  'Cod do Corretor.
Public gCodProduto As Integer    'C�digo do TARIFA
Public gCodSegurado As Long    'Codigo do Segurado (Estipulante)
Public gCodProtocolo As Long  'C�digo do protocolo.
Public gCodSequencia As Long  'C�digo da sequ�ncia.
Public gQuaErrosNormas As Integer
Public gQuaErrosDemais As Integer
Public EhTeste As Boolean
Public gFinalizarPrograma As Boolean
Public gFormCancelado As Boolean
Public gDesabilitarForms As Boolean
Public gDesabFormsCotacao As Boolean
Public gPendenteWorkflow As Boolean
Public gAbrirWorkflow As Boolean
Public gDesabFormsTransmi As Boolean
Public gChamaSelecionarCoberturas As Boolean
Public gNum_Apolice As Double         'NUMERO DA APOLICE
'
'
Public gCod_Ramo As e_CodRamo    'C�digo do ramo.

Public MAPI_SERVER_EXCHANGE As String
Public MAPI_Profile As String
Public MAPI_password As String
Public MAPI_CONTA_EXCHANGE As String


Public gProCusApolice As Double     'Produto: custo de ap�lice.
Public gProCusEndosso As Double     'Produto: custo de endosso.
Public gProPreMinParcela As Double        'Produto: pr�mio m�nimo por parcela. (no Auto � pr�mio m�nimo do item)
Public gProPreMinParcelaEndosso As Double    'Produto: pr�mio m�nimo por parcela. (no Auto � pr�mio m�nimo do item)
Public gProPremioMinimo As Double       'Produto: pr�mio m�nimo do item. (no Auto � pr�mio m�nimo da parcela)
Public gProPreMinEndosso As Double
Public gProDatUltParcela As Long          'Produto: data da �ltima parcela (parcelamento)


Public gProPreMinParcelaTAB As Double
Public gProPremioMinimoTAB As Double
Public gProQuaMaxParcelasTAB As Byte
Public gNovaTransmissaoMulti As Boolean

Public gTipUsuario As e_Tipo_Usu    'Tipo do usu�rio.
Public gBaseLocal As Boolean
Public gTipP00APOL As String  'Tipo da base de dados: P00APOL.
Public gEndP00APOL As String  'Endere�o da base de dados: P00APOL.
Public gCodSeguradoYas As String

Public gTipP00MPED As String  'Tipo da base de dados: P00MPED.
Public gEndP00MPED As String  'Endere�o da base de dados: P00MPED.
Public gTipP0044700 As String   'Tipo da base de dados: P0044700.
Public gEndP0044700 As String   'Endere�o da base de dados: P0044700.
Public gPrefixo As String
Public gDiretorio As String
Public PerfilAcesso As e_PerfilAcesso    'M99V001.Perfil de Acesso no sistema para os usu�rios YASUDA
'1 - Gestor (Digita��o com acesso Master)
'2 - Usu�rio Digita��o (IST)
'3 - Gestor (C�lculo)
'4 - Usu�rio para C�lculo (vers�o corretor para usu�rios YASUDA)
'5 - Usu�rio Digita��o (YASUDA)
Public gTipP00MULTI As String   'Tipo da base de dados: P00Multi.
Public gEndP00MULTI As String   'Endere�o da base de dados: P00MULTI.
Public gAppPath As String    'App.Path considerando sempre uma \ no final.
Public gForHoje As String    'Data do dia formatada.
Public gHojAno As Integer    'Data do dia: ano.
Public gHojDia As Integer    'Data do dia: dia.
Public gHojMes As Integer    'Data do dia: m�s.
Public gHoje As Double    'Data do dia.

Public gNomArqLog As String    'Nome do arquivo de log.
Public gNomUsuario As String  'Nome do usu�rio.
Public gTipoHomologacao As String       'Verifica se � vers�o de homologa��o ou n�o
Public gNomUsuario6 As String   'Nome do usu�rio (6 posi��es).
Public gNomUsuMainFrame As String       'Nome do usu�rio no MainFrame

Public gVerDataMaq As Byte    '0 - N�o ver inconsist�ncia de datas
'1 - Ver inconsist�ncias de data
Public gTrvSistema As Boolean    'Travar Sistema

Public gPwdGed As String    'Par�metro PWS para base de dados: SISCAD.
Public gSerGed As String    'Par�metro SERVER para base de dados: SISCAD.
Public gTipGed As String    'Tipo da base de dados: SISCAD.
Public gUidGed As String    'Par�metro UID para base de dados: SISCAD.
Public gEndGed As String    'Endere�o da base de dados: SISCAD.

Public gPwdDIVERSOS As String   'Par�metro PWS para base de dados: P00MPED.
Public gSerDIVERSOS As String   'Par�metro SERVER para base de dados: P00MPED.
Public gTipDIVERSOS As String   'Tipo da base de dados: P00MPED.
Public gUidDIVERSOS As String   'Par�metro UID para base de dados: P00MPED.
Public gEndDIVERSOS As String   'Endere�o da base de dados: P00MPED.

Public gPwdYAS_ND As String    'Par�metro PWS para base de dados: P00MPED.
Public gSerYAS_ND As String    'Par�metro SERVER para base de dados: P00MPED.
Public gTipYAS_ND As String    'Tipo da base de dados: P00MPED.
Public gUidYAS_ND As String    'Par�metro UID para base de dados: P00MPED.
Public gEndYAS_ND As String    'Endere�o da base de dados: P00MPED.
Public mConexaoSQL As ADODB.Connection    'Conex�o com SQL Server.

Public gTipP0043000 As String   'Tipo da base de dados: P0043000.
Public gTipP0045200 As String   'Tipo da base de dados: P0045200.
Public gTipP0045300 As String   'Tipo da base de dados: P0045300.
Public gTipP0045400 As String   'Tipo da base de dados: P0045400.

Public gPwdSYASEMIS As String   'Par�metro PWS para base de dados: SYASEMIS.
Public gSerSYASEMIS As String   'Par�metro SERVER para base de dados: SYASEMIS.
Public gTipSYASEMIS As String   'Tipo da base de dados: SYASEMIS.
Public gUidSYASEMIS As String   'Par�metro UID para base de dados: SYASEMIS.
Public gEndSYASEMIS As String   'Endere�o da base de dados: SYASEMIS.

Public gEndP0043000 As String   'Endere�o da base de dados: P0043000.
Public gEndP0045200 As String   'Endere�o da base de dados: P0045200.
Public gEndP0045300 As String   'Endere�o da base de dados: P0045300.
Public gEndP0045400 As String   'Endere�o da base de dados: P0045400.


Public gPwdEAPDB As String    'Par�metro PWS para base de dados: EAPDB.
Public gSerEAPDB As String    'Par�metro SERVER para base de dados: EAPDB.
Public gTipEAPDB As String    'Tipo da base de dados: EAPDB.
Public gUidEAPDB As String    'Par�metro UID para base de dados: EAPDB.
Public gEndEAPDB As String    'Endere�o da base de dados: EAPDB.

Public gPwdSyasPortal As String     'Par�metro PWS para base de dados: SyasPortal
Public gSerSyasPortal As String     'Par�metro SERVER para base de dados: SyasPortal
Public gTipSyasPortal As String     'Tipo da base de dados: SyasPortal
Public gUidSyasPortal As String     'Par�metro UID para base de dados: SyasPortal
Public gEndSyasPortal As String     'Endere�o da base de dados: SyasPortal

Public gAgendarCot As Boolean    'Flag para carregar cota��es
Public gTipoAcesso As e_Acesso

Public gEndSYASUpDt As String   'Endere�o da localiza��o do s arquivos de atualiza��o do SYAS


Public gEndRemessa As String  'Endere�o da pasta de remessa (arquivos a serem transmitidos).
Public gEndServidor As String   'Endere�o do servidor.

'[FRH] - GLB_sinal_concess � uma global que � true se as propr. [.Cod_Clasf] e [.Cod_Compl_Clasf] do Item (pobjstcitem-mfcalpreitem) s�o relativos a "022-13-Concession�ria".
Public GLB_sinal_concess As Boolean
Public gerros As Variant    'Erros.
Public gAchouShopping As Boolean     'Indica se achou o fator "RISCOS EM SHOPPING" ou n�o
Public gNovaPropostaLMI As Boolean
Public VerificarNormas As Boolean
'Ficha 1000305
Public gTipCobertura As Integer    'Tipo de cobertura dos bens (pr�dio/conteudo, pr�dio , conte�do)
Public gEplano As Boolean
Public gEplanoSimplificado As Boolean
Public gPrograma As String    'Nome do programa em execu��o.

Public gEndExeSiscad As String    'Endere�o da base de dados: SISCAD.
Public gEndExeP47V001 As String     'Endere�o do execut�vel P47V001.

Public gEndConfCot1 As String   'Endere�o do formul�rio MULTCOT1.CXY
Public gEndImagCot1 As String   'Endere�o do formul�rio MULTCOT1.WMF
Public gEndConfCot2 As String   'Endere�o do formul�rio MULTCOT2.CXY
Public gEndImagCot2 As String   'Endere�o do formul�rio MULTCOT2.WMF
Public gEndConfCot3 As String   'Endere�o do formul�rio MULTCOT3.CXY
Public gEndImagCot3 As String   'Endere�o do formul�rio MULTCOT3.WMF
Public gEndConfCotB As String   'Endere�o do formul�rio MULTCOTB.CXY
Public gEndImagCotB As String   'Endere�o do formul�rio MULTCOTB.WMF
Public gEndConfCotA As String   'Endere�o do formul�rio MULTCOTA.CXY
Public gEndImagCotA As String   'Endere�o do formul�rio MULTCOTA.WMF
Public gEndConfCotC As String   'Endere�o do formul�rio MULTCOTC.CXY
Public gEndImagCotC As String   'Endere�o do formul�rio MULTCOTC.WMF

Public gEndImagDeb As String  'Endere�o do formul�rio L28V009A.CXY
Public gEndConfDeb As String  'Endere�o do formul�rio L28V009A.WMF
Public gEndConfPro1 As String   'Endere�o do formul�rio MULTPRO1.CXY
Public gEndImagPro1 As String   'Endere�o do formul�rio MULTPRO1.WMF
Public gEndConfPro2 As String   'Endere�o do formul�rio MULTPRO2.CXY
Public gEndImagPro2 As String   'Endere�o do formul�rio MULTPRO2.WMF
Public gEndConfPro3 As String   'Endere�o do formul�rio MULTPRO3.CXY
Public gEndImagPro3 As String   'Endere�o do formul�rio MULTPRO3.WMF
Public gEndConfPro4 As String   'Endere�o do formul�rio MULTPRO4.CXY
Public gEndImagPro4 As String   'Endere�o do formul�rio MULTPRO4.WMF
Public gEndConfProB As String   'Endere�o do formul�rio MULTPROB.CXY
Public gEndImagProB As String   'Endere�o do formul�rio MULTPROB.WMF
Public gEndConfProC As String   'Endere�o do formul�rio MULTPROC.CXY
Public gEndImagProC As String   'Endere�o do formul�rio MULTPROC.WMF
Public gEndConfProD As String   'Endere�o do formul�rio MULTPROD.CXY
Public gEndImagProD As String   'Endere�o do formul�rio MULTPROD.WMF


'Formularios Antigos
Public gEndCot1_v0 As String  'Endere�o do formul�rio MULTCOT1
Public gEndCot2_v0 As String  'Endere�o do formul�rio MULTCOT2
Public gEndCot3_v0 As String  'Endere�o do formul�rio MULTCOT3
Public gEndPro1_v0 As String  'Endere�o do formul�rio MULTPRO1
Public gEndPro2_v0 As String  'Endere�o do formul�rio MULTPRO2
Public gEndPro3_v0 As String  'Endere�o do formul�rio MULTPRO3
Public gEndPro4_v0 As String  'Endere�o do formul�rio MULTPRO4
Public gEndProB_v0 As String  'Endere�o do formul�rio MULTPROB

'## smrosa
'Formul�rio Roubo
Public gEndConfRob1 As String   'Endere�o do formul�rio ROUBOCOT1.CXY
Public gEndImagRob1 As String   'Endere�o do formul�rio ROUBOCOT1.WMF
Public gEndConfRob2 As String   'Endere�o do formul�rio ROUBOCOT2.CXY
Public gEndImagRob2 As String   'Endere�o do formul�rio ROUBOCOT2.WMF
Public gEndConfRob3 As String   'Endere�o do formul�rio ROUBOCOT3.CXY
Public gEndImagRob3 As String   'Endere�o do formul�rio ROUBOCOT3.WMF
Public gEndConfRob4 As String   'Endere�o do formul�rio ROUBOCOT4.CXY
Public gEndImagRob4 As String   'Endere�o do formul�rio ROUBOCOT4.WMF
Public gEndConfRob5 As String   'Endere�o do formul�rio ROUBOCOT5.CXY
Public gEndImagRob5 As String   'Endere�o do formul�rio ROUBOCOT5.WMF

Public gEndImagCotLMI As String     'Endere�o do formul�rio COT_LMI.WMF
Public gEndImagRelLMI As String     'Endere�o do formul�rio REL_LMI.WMF
Public gEndImagESPLMI As String     'Endere�o do formul�rio COT_LMI.WMF
Public gEndImagESPRelLMI As String     'Endere�o do formul�rio REL_LMI.WMF

Public gEfetuouCalculo As Boolean
Public gPodeEmitir As Boolean    'True se o d�gito de ag�ncia e conta estiverem OK
Public gPathArquivos As String    'Local onde est�o os arquivos de layout.
Public gPathImagePI As String   'Local das Imagens do PI
Public gImprFrente As String
Public gImprFrenteVerso As String

Public gUsuMulti As Boolean    'Indica se o usu�rio � somente do sistema Multi ou � do Syas
'tamb�m: TRUE = ambos sistemas e FALSE = apenas MULTISSEGUROS.
Public gDir_Ist As String    'Path para gravar arquivos de conferencia da IST

Public gQtdMaxComp As Byte    'Quantidade Maxima de tipos de compensacao, iniciando com 2 (boleto e debito)

Public gUrlWebServiceMotorCalc As String    ' URL do WebService de Motor de Calculo
Public gWebServiceEnabled As Boolean    ' Flag que diz se o sistema deve ou n�o usar o WebService de Calculo

Public gDataCorteSompo As String ' Data de corte (apos essa data, o nome da seguradora muda para Sompo)

Public gUrlWebServiceNormas As String    ' URL do WebService de Normas
Public gWebServiceNormasEnabled As Boolean    ' Flag que diz se o sistema deve ou n�o usar o WebService de Normas
Public gWebServiceNormasTarifa As Long    ' Flag que diz se o sistema deve ou n�o usar o WebService de Normas

Public gUrlWebServiceBoleto As String    ' URL do WebService de Normas
Public gWebServiceBoletoEnabled As Boolean    ' Flag que diz se o sistema deve ou n�o usar o WebService de Normas

Public gUrlRelatorio As String
Public gUsuarioRelatorio As String

Public gEnableFileLog As Boolean
Public sUniqueInstanceCode As Integer

Private mClsMDynaTrace As clsMDynaTrace
Public tagDynaTrace As String

Public Enum e_Acesso
    DIGITACAO = 1
    CALCULO = 2
End Enum

Public Enum e_Tipo_Usu
    e_Tipo_Usu_Yasuda = 1
    e_Tipo_Usu_Corretor = 2
    e_Tipo_Usu_CorretorYasuda = 3
End Enum

Public Enum e_SimNao
    e_SimNao_Sim = 1
    e_SimNao_Nao = 2
End Enum

Public Enum TipoDados
    tpoNumerico = 1
    TpoData = 2
    TpoHora = 3
    TpoValor = 4
    tpoString = 5
    TpoBoolean = 6
    TpoNumInteiro = 7
    TpoNumDecimal = 8
End Enum
Enum e_CodDizer
    e_CodDizerAlteracao = 12
    e_CodDizerInclusao = 1
    e_CodDizerexclusao = 4
    e_CodDizerCancelamento = 10
End Enum

Public Enum e_PerfilAcesso
    e_Perfil_DIG_Gestor = 1
    e_Perfil_DIG_IST = 2
    e_Perfil_CALC_Gestor = 3
    e_Perfil_CALC = 4
    e_Perfil_DIG_Yas = 5
    e_Perfil_Operacao = 6
    e_Perfil_Triagem = 7
End Enum



Public Enum e_TipoSistema
    e_TipoSistema_ND = 0
    e_TipoSistema_Serasa = 1
    e_TipoSistema_Nortix = 2
    e_TipoSistema_BoaVista = 3
End Enum

Public Enum e_TipoLayout
    e_TipoLayout_ND = 0
    e_TipoLayout_Entrada = 1
    e_TipoLayout_Saida = 2
End Enum

Public Enum e_Delimitador
    e_Delimitador_ND = 0
    e_Delimitador_XML = 1
    e_Delimitador_Fixo = 2
    e_Delimitador_PontVirg = 3
End Enum


Public Enum e_TipoChave
    e_TipoChave_ND = 0
    e_TipoChave_Sistema = 1
    e_TipoChave_Cnpj = 2
    e_TipoChave_Placa = 3
    e_TipoChave_Chassi = 4
    e_TipoChave_Renavam = 5
End Enum

Public Enum ETransmDLL
    ETransmDLL_ONLine = 0
    ETransmDLL_OFFLine = 1
End Enum


Public Enum e_CodPlanoBasico
    e_CodPlano_Empresarial_Simplificado = 10
    e_CodPlano_Empresarial_Super = 11
    e_CodPlanoBasico_Residencial = 10
    e_CodPlano_Condominio_Tradicional = 10
    e_CodPlano_Condominio_Amplo = 11
End Enum

Public Enum e_OrigemChamada
    e_Origem_Kit = 1
    e_Origem_Web = 2
    e_Origem_FTP = 3    '2016-07-11: Usado no Imobiliario/KitMar
End Enum

Public gTipoTransmDLL As ETransmDLL

Public gProibidoMsgBox As Boolean    'Indica que um MsgBox N�o pode ser exibido nunca (Exemplo Servi�os e execu��es em modo oculto)

Private mClsM46V111 As clsM46V111
Private mClsM46O001 As clsM46O001
Private mClsM46V103 As clsM46V103

Public gDesconsiderou As Integer
Public gPermitePostecipada As Boolean
Public gPermitePrInformado As Boolean
Public gRelUsuPostecipada As String
Public gRelUsuPrInformado As String
Public gRst_Tbm_De_Para As ADODB.Recordset
Public gRelUsuBookKit As String
Public gTarifaVigente As String
Public gTarifaTeste As String
Public gBufferSql As Boolean
Public gScriptSql As String
Public gRst_Tab_Cobert As ADODB.Recordset
Public gCod_Segmentacao As Integer
Public gParcMaxCorr As Integer
Public gPathScript As String
Public gOrigemChamada As e_OrigemChamada
Public gValidadeCotCorr As Integer
Public gDiasVencBoleto As Integer
Public gRelPilotoSupremo As String
Public gResidencialSupremo As String
Public gResidencialSupremoTodos As String
Public gPathSiscota As String
'' 2016-08-11 Removido
'Public gRelUsuImpressaoFtp As String            ''2016-07-12
'Public gPermiteImpressaoFtp As Boolean

Public gStcA46V750BrptSelected As stcA46V750Brpt
Public gStcA46V750BrptCollection As Collection

Public gRASelected As stcA46V750Brpt

Public Enum e_CodRamoSusep
    e_CodRamoSusep_Empresarial = 18
    e_CodRamoSusep_Residencial = 14
    e_CodRamoSusep_Condominio = 16
End Enum

Private m_gProQuaMaxParcelas As Byte           'Produto: quantidade m�xima de parcelas.

Property Get gProQuaMaxParcelas() As Byte
    gProQuaMaxParcelas = m_gProQuaMaxParcelas
End Property

Property Let gProQuaMaxParcelas(oValue As Byte)
    m_gProQuaMaxParcelas = oValue
End Property


Function gfBuscaPerIndenitarioDefault(ByVal pCodCobert As Integer, ByVal pCodModalidade As Integer) As Integer
    Dim lBanco As ADODB.Connection              'Arquivo: P00Multi.
    Dim lMensagem As String                     'Mensagem.
    Dim lrsTbm_Coef_PI As clsYasRecordSet       'Registro : Tbm_Coef_PI.
    Dim lSelect As String

    gfBuscaPerIndenitarioDefault = 0

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lBanco, lMensagem) = False Then
        Call gpGraLog(2, lMensagem)
        Exit Function
    End If
    
    '1 - Procura pela Cobertura na Tabela de Coeficiente
    'Se encontrar � permitido Cadastrar
    lSelect = " Select *  from Tbm_Coef_PI "
    lSelect = lSelect & " Where Cod_Ramo = " & gCod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & " and Cod_Tarif = " & gCodProduto
    lSelect = lSelect & " and Cod_Sequencia <= " & gCodSequencia
    lSelect = lSelect & " order by Cod_Sequencia desc "
    
    If gfObtRegistro(lBanco, lSelect, lrsTbm_Coef_PI, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If
    
    If Not lrsTbm_Coef_PI.EOF Then
        '2 - Procura pela Cobertura na Tabela de PI Default
        lSelect = " Select *  from Tbm_Per_Ind_Padrao "
        lSelect = lSelect & " Where Cod_Ramo = " & gCod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & pCodModalidade
        lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
        lSelect = lSelect & " and Cod_Tarif = " & gCodProduto
        lSelect = lSelect & " and Cod_Sequencia <= " & gCodSequencia
        lSelect = lSelect & " order by Cod_Sequencia desc "
    
        If gfObtRegistro(lBanco, lSelect, lrsTbm_Coef_PI, lMensagem) = False Then
            Call gpGraLog(1, lMensagem)
            GoTo SAIDA
        End If
    
        '3 - Se encontrar permite cadastro de Per�odo Indenit�rio
        If Not lrsTbm_Coef_PI.EOF Then
            gfBuscaPerIndenitarioDefault = lrsTbm_Coef_PI("Per_Indenit_Def")
        Else
            gfBuscaPerIndenitarioDefault = 12
        End If
    End If

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lrsTbm_Coef_PI = Nothing
    Return
End Function

Function gfBuscaTextoPiIndenitario(pCodCobert, ByVal pNum_Pi_Meses) As String
    Dim lBanco As ADODB.Connection    'Arquivo: P00Multi.
    Dim lMensagem As String
    Dim lrst As clsYasRecordSet
    Dim lSelect As String
    Dim lLimite As Integer

    gfBuscaTextoPiIndenitario = ""

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lBanco, lMensagem) = False Then
        Call gpGraLog(2, lMensagem)
        Exit Function
    End If
    lSelect = " Select * from Tbm_Fator_Franquia_PI WITH (NOLOCK) Where"
    lSelect = lSelect & " Cod_Ramo = " & gCod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & " and Cod_Tarif = " & gCodProduto
    lSelect = lSelect & " and Cod_Sequencia <= " & gCodSequencia

    If gfObtRegistro(lBanco, lSelect, lrst, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    If Not lrst.EOF Then
        If pNum_Pi_Meses > lrst("Limite_PI") Then
            gfBuscaTextoPiIndenitario = lrst("TextoMaior")
        Else
            gfBuscaTextoPiIndenitario = lrst("TextoMenor")
        End If
    End If

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lrst = Nothing
    Return
End Function


Public Property Get Class_Name() As String
    Class_Name = "M46V999"
End Property

Function gfRetornaFatores() As Boolean

End Function

Public Sub gpDesconsiderou()
    If gDesconsiderou = 0 Then
        If MsgBox("Deseja desconsiderar TODAS as consist�ncias?", vbQuestion + vbYesNo) = vbYes Then
            gDesconsiderou = 1
        Else
            gDesconsiderou = 2
        End If
    End If
End Sub

Public Sub gpInicializar_gErros()
    ReDim gerros(0) As String
    gQuaErros = 0
    gQuaErrosNormas = 0
    gQuaErrosDemais = 0
End Sub

Public Sub gpCenForm(ByVal pForm As Object)
'Procedure: centraliza form.

    pForm.Left = (Screen.Width - pForm.Width) / 2
    pForm.Top = (Screen.Height - pForm.Height) / 2
End Sub
Public Function gfBuscaSusep(ByVal pCodCorretor As String, ByRef pSusep As String, _
                             ByRef pMensagem As String) As Boolean
'Fun��o: obt�m nome do corretor.

'Par�metro de entrada...pCodCorretor...C�digo do corretor.
'Par�metros de sa�da....pNomCorretor...Nome do corretor.
'                       pMensagem......Mensagem.

    Dim bdP0044700 As ADODB.Connection    'Arquivo - P0044700.
    Dim lSelect As String       'Comando SELECT.
    Dim rsCORR As clsYasRecordSet    'YTAB_CORR: tabela.

    gfBuscaSusep = False

    '1. Verifica preenchimento.
    If Not gfPreenchido(pCodCorretor) Then
        gfBuscaSusep = ""
        gfBuscaSusep = True
        Exit Function
    End If
    If val(pCodCorretor) = 0 Then
        gfBuscaSusep = ""
        gfBuscaSusep = True
        Exit Function
    End If

    '2. Obt�m nome do corretor.
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    lSelect = "SELECT NUM_SUSEP FROM YTAB_CORR WHERE COD_CORR = " & pCodCorretor
    If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
        Call Me.gpFecha2(bdP0044700)
        Exit Function
    End If
    If rsCORR.EOF = True Then
        Call Me.gpFecha2(bdP0044700)
        pMensagem = "FGE0044 - " & gPrefixo & "Corretor inv�lido (c�digo de corretor n�o localizado " & _
                    "- " & pCodCorretor & ")."
        Exit Function
    End If
    pSusep = Trim$(rsCORR!NUM_SUSEP)
    Call Me.gpFecha2(bdP0044700)

    gfBuscaSusep = True
End Function

Public Function Check(campo As Object) As Variant
    On Error GoTo TrataCheck
    Dim lYasField As clsYasField

    Set lYasField = campo
    If IsNull(lYasField) Then
        If lYasField.Type_ = 5 Or lYasField.Type_ = 200 Or lYasField.Type_ = 201 Or lYasField.Type_ = 202 Or lYasField.Type_ = 203 Or lYasField.Type_ = 129 Then
            Check = ""
        ElseIf lYasField.Type_ = 7 Then
            Check = ""
        ElseIf lYasField.Type_ = 2 Then
            Check = 0
        ElseIf lYasField.Type_ = 135 Then
        ElseIf lYasField.Type_ = 3 Or lYasField.Type_ = 131 Then
            Check = 0
        Else
            MsgBox "Verificar tipo Campo: " & lYasField.Name & " Tipo: " & lYasField.Type_
            Check = 0
        End If
    Else
        Check = lYasField.Value
    End If
SAIDA:
    GoSub FechaObj
    Exit Function
TrataCheck:
    If Err = 3021 Then Resume Next
    MsgBox "Problemas com o campo " & lYasField.Name
    GoTo SAIDA
FechaObj:
    Set lYasField = Nothing
    Return
End Function

Public Function gfVerEtiqueta(ByVal pCorretor As Long, ByVal pCodReg As Integer, ByVal pCodPessoa As Integer, ByVal pTipEnv As String, ByRef pMensagem As String, Optional ByVal tipoEnvio As String) As Boolean

    Dim bdP00MULTI As ADODB.Connection
    Dim rsTbsCorretor As clsYasRecordSet    'TBS_CORRETOR: tabela.
    Dim lSelect As String

    If gTipUsuario = e_Tipo_Usu_Corretor Or _
       gTipUsuario = e_Tipo_Usu_CorretorYasuda Or _
       gBaseLocal = True Or _
       PerfilAcesso = e_Perfil_CALC Or _
       PerfilAcesso = e_Perfil_CALC_Gestor Then
        gfVerEtiqueta = False
        GoTo SAIDA
    End If

    If (pCorretor = 15701 Or pCorretor = 8101) And _
       Not (pCodReg = 106 Or pCodReg = 107) Then
        If pCodPessoa = 1 And pTipEnv <> "C" Then
            If pTipEnv = "X" Then pMensagem = "O": GoTo SAIDA    '[FRH]Ficha:0801295 teste
            pMensagem = "Envio do documento deve ser para o 'Corretor'."
            gfVerEtiqueta = True
        ElseIf pCodPessoa = 2 And pTipEnv <> "S" Then
            If pTipEnv = "X" Then pMensagem = "E": GoTo SAIDA    '[FRH]Ficha:0801295
            pMensagem = "Envio do documento deve ser para o 'Segurado'."
            gfVerEtiqueta = True
        Else
            gfVerEtiqueta = False
        End If
        GoTo SAIDA
    End If

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdP00MULTI, pMensagem) = False Then
        GoTo SAIDA
    End If

    lSelect = "SELECT * FROM ged.dbo.TBS_CORRETOR WITH (NOLOCK)  WHERE Cod_Corretor = " & pCorretor

    If gfObtRegistro(bdP00MULTI, lSelect, rsTbsCorretor, pMensagem) = False Then
        Call Me.gpFecha2(bdP00MULTI)
        GoTo SAIDA
    End If

    If rsTbsCorretor.EOF Then
        pMensagem = "Corretor n�o cadastrado na TBS_CORRETOR!!"
        gfVerEtiqueta = True
        Call Me.gpFecha2(bdP00MULTI)
        GoTo SAIDA
    End If

    If rsTbsCorretor("Cod_Etiqueta") = "5" Or _
       rsTbsCorretor("Cod_Etiqueta") = "6" Or _
       rsTbsCorretor("Cod_Etiqueta") = "7" Or _
       rsTbsCorretor("Cod_Etiqueta") = "9" Or _
       rsTbsCorretor("Cod_Etiqueta") = "A" Then

        gfVerEtiqueta = False

        If pTipEnv = "S" Then
            pMensagem = "Envio do documento deve ser para o 'Corretor'."
            gfVerEtiqueta = True
            GoTo SAIDA
        ElseIf pTipEnv = "X" Then

            '[FRH]Ficha:0801295 [Conforme cadastro - "O"]
            pMensagem = "O"

            Call Me.gpFecha2(bdP00MULTI)
            GoTo SAIDA
        End If
    Else
        If pTipEnv = "X" Then

            '[FRH]Ficha:0801295 [Conforme cadastro - "E"]
            pMensagem = "E"

            Call Me.gpFecha2(bdP00MULTI)
            GoTo SAIDA
        End If
    End If

    Call Me.gpFecha2(bdP00MULTI)

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set bdP00MULTI = Nothing
    Set rsTbsCorretor = Nothing
    Return
End Function

Public Function gfCalCodUnicidade_V1(ByRef pobjstcPedido As stcA46V702B, ByRef pMensagem As String)
'Fun��o: Calcula c�digo de verifica��o da unicidade.

'Par�metro de sa�da...pMensagem.........Mensagem.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lCampo1() As String    'Campos utilizados para dividir a seq��ncia em campos de
'10 posi��es cada.
    Dim lCampo2() As Double    'Campos com totais das somas das posi��es de lCampo1.
    Dim lCodVerificacao As String       'C�digo de verifica��o.
    Dim lI As Double    'Utilizado com For...Next.
    Dim lPosInicial As Double   'Posi��o inicial.
    Dim lResto As Double    'Resto da divis�o.
    Dim lResultado As Double  'Resultado da divis�o.
    Dim lSequencia As String  'Seq��ncia para c�lculo do c�digo de unicidade.
    Dim lSequencia2 As String   'Seq��ncia para c�lculo do c�digo de unicidade.
    Dim lTamanho As Double    'Tamanho.
    Dim lTotal As Double    'Total.
    Dim lRamAtividade As String     'Pessoa jur�dica: ramo de atividade.
    Dim lTipDocumento As String     'Pessoa f�sica: tipo do documento.
    Dim lNumDocumento As String     'Pessoa f�sica: n�mero do documento.
    Dim lOrgEmiDocumento As String        'Pessoa f�sica: �rg�o emissor do documento.
    Dim lDatEmiDocumento As String        'Pessoa f�sica: data de emiss�o do documento.

    Dim lobjstcCoss As stcA46V706B    'Objeto da classe de estrutura stcA46V706B (TAB_PED_COSS).
    Dim lobjstcJuros As stcA46V711B    'Objeto da classe de estrutura stcA46V711B (TAB_PED_JUROS).
    Dim lObjstcVenc As stcA46V713B    'Objeto da classe de estrutura stcA46V713B (TAB_PED_VENC).
    Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcCobert As stcA46V704B

    gfCalCodUnicidade_V1 = False

    'Crio a sequencia
    lSequencia = ""
    lSequencia2 = ""
    '-*-*-*-*-*
    Select Case pobjstcPedido.Tip_Pessoa
    Case e_TipPessoa_Juridica       'Pessoa jur�dica.
        'a) Ramo de atividade - 2 posi��es num�ricas.
        If IsNumeric(pobjstcPedido.Cod_Ramo_Ativ) Then
            lRamAtividade = Format(pobjstcPedido.Cod_Ramo_Ativ, "00")
        Else
            lRamAtividade = "00"
        End If
    Case e_TipPessoa_Fisica       'Pessoa f�sica.
        'a) Registro de identifica��o do segurado.
        lTipDocumento = "00"
        lNumDocumento = String(20, " ")
        lOrgEmiDocumento = String(20, " ")
        lDatEmiDocumento = String(8, "0")

        'b) Tipo do documento - 2 posi��es num�ricas.
        If IsNumeric(pobjstcPedido.Cod_Doc) Then
            lTipDocumento = pobjstcPedido.Cod_Doc
        End If
        'c) N�mero do documento - 20 posi��es.
        lNumDocumento = pobjstcPedido.Num_Doc
        lNumDocumento = lNumDocumento & String(20 - Len(Left(pobjstcPedido.Num_Doc, 20)), " ")

        'd) �rg�o emissor do documento - 20 posi��es.
        lOrgEmiDocumento = pobjstcPedido.Org_Exp
        lOrgEmiDocumento = lOrgEmiDocumento & String(20 - Len(Left(pobjstcPedido.Org_Exp, 20)), " ")

        'e) Data de emiss�o do documento - 8 posi��es num�ricas.
        If IsNumeric(pobjstcPedido.Dat_Emis_Doc) Then
            lDatEmiDocumento = pobjstcPedido.Dat_Emis_Doc
        End If
    End Select

    '3. Inicia seq��ncia para c�lculo do c�digo de unicidade com dados comuns.
    '===001) Nosso n�mero......................................................(20 posi��es - total = 20).
    lSequencia = pobjstcPedido.nosso_numero
    '===002) Refer�ncia.........................................................(6 posi��es - total = 26).
    'lSequencia = lSequencia & Format$(pobjstcPedido.Num_Referencia, "000000")
    '===003) C�digo do corretor.................................................(7 posi��es - total = 33).
    lSequencia = lSequencia & Format$(pobjstcPedido.objstcCtrlEmis.Cod_Corr, "0000000")
    '===004) C�digo do produto (tarifa).........................................(4 posi��es - total = 37).
    lSequencia = lSequencia & Format$(pobjstcPedido.Cod_Produto, "0000")
    '===005) C�digo da unidade de neg�cios......................................(5 posi��es - total = 42).
    lSequencia = lSequencia & Format$(pobjstcPedido.Cod_Unid_Prod, "00000")
    '===006) C�digo do produtor.................................................(5 posi��es - total = 47).
    lSequencia = lSequencia & Format$(pobjstcPedido.Cod_Prod, "00000")
    '===007) Nome do proponente (segurado)....................................(60 posi��es - total = 107).
    lAuxiliar = UCase$(pobjstcPedido.Nom_Segurado)
    If Len(lAuxiliar) > 60 Then
        lAuxiliar = Left$(lAuxiliar, 60)
    Else
        lAuxiliar = lAuxiliar & String(60 - Len(lAuxiliar), " ")
    End If
    lSequencia = lSequencia & lAuxiliar

    '===008) C�digo do segurado................................................(7 posi��es - total = 114).
    lAuxiliar = String(7, "0")
    If gfPreenchido(pobjstcPedido.Cod_Segurado) Then
        If IsNumeric(pobjstcPedido.Cod_Segurado) Then
            lAuxiliar = Format$(val(pobjstcPedido.Cod_Segurado), "0000000")
        End If
    End If
    lSequencia = lSequencia & lAuxiliar
    '===009) Tipo de pessoa.....................................................(1 posi��o - total = 115).
    '===010) Identifica��o....................................................(61 posi��es - total = 176).
    Select Case pobjstcPedido.Tip_Pessoa
    Case e_TipPessoa_Juridica
        'a) Tipo de pessoa.
        lAuxiliar = "J"
        'b) CNPJ - 14 posi��es (12 alfanum�ricas + 2 d�gitos verificadores num�ricos).
        If IsNumeric(pobjstcPedido.Num_Cgc_Cpf) Then
            lAuxiliar = lAuxiliar & Format$(pobjstcPedido.Num_Cgc_Cpf, "00000000000000")
        Else
            lAuxiliar = lAuxiliar & Right$(String(14, "0") & Trim(pobjstcPedido.Num_Cgc_Cpf), 14)
        End If
        'c) Ramo de atividade - 2 posi��es num�ricas.
        lAuxiliar = lAuxiliar & lRamAtividade
        'd) Completa com espa�os.
        lAuxiliar = lAuxiliar & String(45, " ")
    Case e_TipPessoa_Fisica
        'a) Tipo de pessoa.
        lAuxiliar = "F"
        'b) CPF - 11 posi��es num�ricas.
        lAuxiliar = lAuxiliar & Format$(pobjstcPedido.Num_Cgc_Cpf, "00000000000")
        'c) Tipo do documento - 2 posi��es num�ricas.
        lAuxiliar = lAuxiliar & lTipDocumento
        'd) N�mero do documento - 20 posi��es.
        lAuxiliar = lAuxiliar & lNumDocumento
        'e) �rg�o emissor do documento - 20 posi��es.
        lAuxiliar = lAuxiliar & lOrgEmiDocumento
        'f) Data de emiss�o do documento - 8 posi��es num�ricas (aaaammdd).
        lAuxiliar = lAuxiliar & lDatEmiDocumento
    End Select
    lSequencia = lSequencia & lAuxiliar

    '===011) Endere�o do segurado - logradouro................................(73 posi��es - total = 249).
    'O Endere�o n�o faz mais parte do c�digo de unicidade (solicita��o Shiba)

    '**************************************************************************************
    '*    lAuxiliar = ""
    '*    If gfPreenchido(pObjstcPedido.Cod_Tip_Logr) Then
    '*        lAuxiliar = Trim$(pObjstcPedido.Cod_Tip_Logr) & " "
    '*    End If
    '*    If gfPreenchido(pObjstcPedido.Nom_Logr) Then
    '*        lAuxiliar = lAuxiliar & Trim$(pObjstcPedido.Nom_Logr) & " "
    '*    End If
    '*    If gfPreenchido(pObjstcPedido.Dsc_Compl_Logr) Then
    '*        lAuxiliar = lAuxiliar & Trim$(pObjstcPedido.Dsc_Compl_Logr)
    '*    End If
    '*    If Len(lAuxiliar) > 73 Then
    '*        lAuxiliar = Left$(lAuxiliar, 73)
    '*    Else
    '*        lAuxiliar = lAuxiliar & String(73 - Len(lAuxiliar), " ")
    '*    End If

    '*    lSequencia = lSequencia & lAuxiliar
    '===012) Endere�o do segurado - bairro....................................(30 posi��es - total = 279).
    '*    If gfPreenchido(pObjstcPedido.Nom_Bairro) Then
    '*        lAuxiliar = pObjstcPedido.Nom_Bairro
    '*        If Len(lAuxiliar) < 30 Then
    '*            lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
    '*        End If
    '*    Else
    '*        lAuxiliar = String(30, " ")
    '*    End If
    '*    lSequencia = lSequencia & lAuxiliar
    '===013) Endere�o do segurado - cidade....................................(30 posi��es - total = 309).
    '*    If gfPreenchido(pObjstcPedido.Nom_Cid) Then
    '*        lAuxiliar = pObjstcPedido.Nom_Cid
    '*        If Len(lAuxiliar) < 30 Then
    '*            lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
    '*        End If
    '*    Else
    '*        lAuxiliar = String(30, " ")
    '*    End If
    '*    lSequencia = lSequencia & lAuxiliar
    '===014) Endere�o do segurado - unidade da federa��o.......................(2 posi��es - total = 311).
    '*    If gfPreenchido(pObjstcPedido.Sig_UF) Then
    '*        lAuxiliar = pObjstcPedido.Sig_UF
    '*        If Len(lAuxiliar) < 2 Then
    '*            lAuxiliar = lAuxiliar & String(2 - Len(lAuxiliar), " ")
    '*        End If
    '*    Else
    '*        lAuxiliar = String(2, " ")
    '*    End If
    '*    lSequencia = lSequencia & lAuxiliar
    '===015) Endere�o do segurado - CEP........................................(8 posi��es - total = 319).
    '*    If gfPreenchido(pObjstcPedido.Num_Cep) Then
    '*        lAuxiliar = pObjstcPedido.Num_Cep
    '*        If Len(lAuxiliar) < 8 Then
    '*            lAuxiliar = String(8 - Len(lAuxiliar), "0") & lAuxiliar
    '*        End If
    '*    Else
    '*        lAuxiliar = String(8, "0")
    '*    End If
    '*    lSequencia = lSequencia & lAuxiliar
    '===016) Endere�o do segurado - telefones.................................(30 posi��es - total = 349).
    '*    If gfPreenchido(pObjstcPedido.Num_Fone) Then
    '*        lAuxiliar = pObjstcPedido.Num_Fone
    '*        If Len(lAuxiliar) < 30 Then
    '*            lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
    '*        End If
    '*    Else
    '*        lAuxiliar = String(30, " ")
    '*    End If
    '*    lSequencia = lSequencia & lAuxiliar
    '===017) Endere�o de cobran�a - logradouro................................(73 posi��es - total = 422).
    '*    lAuxiliar = ""
    '*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Cod_Tip_Logr) Then
    '*        lAuxiliar = Trim$(pObjstcPedido.objstcPedEndereco.Cod_Tip_Logr) & " "
    '*    End If
    '*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Nom_Logr) Then
    '*        lAuxiliar = lAuxiliar & Trim$(pObjstcPedido.objstcPedEndereco.Nom_Logr) & " "
    '*    End If
    '*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Dsc_Compl_Logr) Then
    '*        lAuxiliar = lAuxiliar & Trim$(pObjstcPedido.objstcPedEndereco.Dsc_Compl_Logr)
    '*    End If
    '*    If Len(lAuxiliar) > 73 Then
    '*        lAuxiliar = Left$(lAuxiliar, 73)
    '*    Else
    '*        lAuxiliar = lAuxiliar & String(73 - Len(lAuxiliar), " ")
    '*    End If
    '*    lSequencia = lSequencia & lAuxiliar
    '*    '===018) Endere�o de cobran�a - bairro....................................(30 posi��es - total = 452).
    '*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Nom_Bairro) Then
    '*        lAuxiliar = pObjstcPedido.objstcPedEndereco.Nom_Bairro
    '*        If Len(lAuxiliar) < 30 Then
    '*            lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
    '*        End If
    '*    Else
    '*        lAuxiliar = String(30, " ")
    '*    End If
    '*    lSequencia = lSequencia & lAuxiliar
    '===019) Endere�o de cobran�a - cidade....................................(30 posi��es - total = 482).
    '*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Nom_Cid) Then
    '*        lAuxiliar = pObjstcPedido.objstcPedEndereco.Nom_Cid
    '*        If Len(lAuxiliar) < 30 Then
    '*            lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
    '*        End If
    '*    Else
    '*        lAuxiliar = String(30, " ")
    '*    End If
    '*    lSequencia = lSequencia & lAuxiliar
    '===020) Endere�o de cobran�a - unidade da federa��o.......................(2 posi��es - total = 484).
    '*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Sig_UF) Then
    '*        lAuxiliar = pObjstcPedido.objstcPedEndereco.Sig_UF
    '*        If Len(lAuxiliar) < 2 Then
    '*            lAuxiliar = lAuxiliar & String(2 - Len(lAuxiliar), " ")
    '*        End If
    '*    Else
    '*        lAuxiliar = String(2, " ")
    '*    End If
    '*    lSequencia = lSequencia & lAuxiliar
    '===021) Endere�o de cobran�a - CEP........................................(8 posi��es - total = 492).
    '*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Num_Cep) Then
    '*        lAuxiliar = pObjstcPedido.objstcPedEndereco.Num_Cep
    '*        If Len(lAuxiliar) < 8 Then
    '*            lAuxiliar = String(8 - Len(lAuxiliar), "0") & lAuxiliar
    '*        End If
    '*    Else
    '*        lAuxiliar = String(8, "0")
    '*    End If
    '*    lSequencia = lSequencia & lAuxiliar
    '===022) Endere�o de cobran�a - telefones.................................(30 posi��es - total = 522).
    '*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Num_Fone) Then
    '*        lAuxiliar = pObjstcPedido.objstcPedEndereco.Num_Fone
    '*        If Len(lAuxiliar) < 30 Then
    '*            lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
    '*        End If
    '*    Else
    '*        lAuxiliar = String(30, " ")
    '*    End If
    '*    lSequencia = lSequencia & lAuxiliar
    '***************************************************************************

    '===023) In�cio de vig�ncia................................................(8 posi��es - total = 530).
    lSequencia = lSequencia & Format$(pobjstcPedido.Dat_Ini_Vig, "00000000")
    '===024) T�rmino de vig�ncia...............................................(8 posi��es - total = 538).
    lSequencia = lSequencia & Format$(pobjstcPedido.Dat_Ter_Vig, "00000000")
    '===025) Corretagem.......................................................(33 posi��es - total = 571).
    lSequencia = lSequencia & Format$(pobjstcPedido.objstcCtrlEmis.Cod_Corr, "0000000")
    lAuxiliar = Format$(pobjstcPedido.objstcCtrlEmis.Per_Corr, "00.00")
    lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
    lSequencia = lSequencia & Format$(pobjstcPedido.objstcCtrlEmis.Cod_Corr_2, "0000000")
    lAuxiliar = Format$(pobjstcPedido.objstcCtrlEmis.Per_Corr_2, "00.00")
    lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
    lSequencia = lSequencia & Format$(pobjstcPedido.objstcCtrlEmis.Cod_Corr_3, "0000000")
    lAuxiliar = Format$(pobjstcPedido.objstcCtrlEmis.Per_Corr_3, "00.00")
    lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)

    '===026) Cosseguros.......................................................(50 posi��es - total = 621).
    For Each lobjstcCoss In pobjstcPedido.ColecaoPedCoss
        lSequencia = lSequencia & Format$(lobjstcCoss.Cod_Coss, "000000")
        lAuxiliar = Format$(lobjstcCoss.Per_Coss, "00.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
    Next lobjstcCoss

    '===027) Juros ao m�s......................................................(4 posi��es - total = 625).
    lAuxiliar = "00,00"
    For Each lobjstcJuros In pobjstcPedido.ColecaoPedJuros
        If pobjstcPedido.Qtd_Parc = lobjstcJuros.Num_Parcela And pobjstcPedido.gTipPagamento = lobjstcJuros.Tip_Compensacao Then
            If pobjstcPedido.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
                lAuxiliar = Format$(lobjstcJuros.Val_Ant, "00.00")
            Else
                lAuxiliar = Format$(lobjstcJuros.Val_Pos, "00.00")
            End If
            Exit For
        End If
    Next
    lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)

    '===028) Custo da ap�lice.................................................(10 posi��es - total = 635).
    lAuxiliar = Format$(pobjstcPedido.Val_Cust_Emis, "00000000.00")
    lSequencia = lSequencia & Left$(lAuxiliar, 8) & Right$(lAuxiliar, 2)
    '===029) Valor do IOF.....................................................(10 posi��es - total = 645).
    lAuxiliar = Format$(pobjstcPedido.Val_IOF, "00000000.00")
    lSequencia = lSequencia & Left$(lAuxiliar, 8) & Right$(lAuxiliar, 2)
    '===030) N�mero de parcelas................................................(2 posi��es - total = 647).
    lAuxiliar = Format$(pobjstcPedido.Qtd_Parc, "00")
    lSequencia = lSequencia & lAuxiliar & lAuxiliar
    '===031) Indicador de primeira parcela paga antecipada......................(1 posi��o - total = 648).
    lAuxiliar = Format$(pobjstcPedido.objstcCtrlEmis.Cod_Iden_Pag_Antec, "0")
    lSequencia = lSequencia & lAuxiliar

    '===032) Valor da primeira parcela........................................(10 posi��es - total = 658).
    'If pobjstcPedido.Val_Parcela1_VM = 0 Then
    '    lAuxiliar = Format$(pobjstcPedido.Val_Parcela1_VD, "00000000.00")
    'Else
    '    lAuxiliar = Format$(pobjstcPedido.Val_Parcela1_VM, "00000000.00")
    'End If
    'lSequencia = lSequencia & Left$(lAuxiliar, 8) & Right$(lAuxiliar, 2)

    '===033) Valor das demais parcelas........................................(10 posi��es - total = 668).
    'If pobjstcPedido.Val_Parcelas_VM = 0 Then
    '    lAuxiliar = Format$(pobjstcPedido.Val_Parcelas_VD, "00000000.00")
    'Else
    '    lAuxiliar = Format$(pobjstcPedido.Val_Parcelas_VM, "00000000.00")
    'End If
    'lSequencia = lSequencia & Left$(lAuxiliar, 8) & Right$(lAuxiliar, 2)

    '===034) Pagamento em bloqueto - n�mero do bloqueto.......................(10 posi��es - total = 678).
    lSequencia = lSequencia & Format$(pobjstcPedido.objstcCtrlEmis.Num_Bloq, "0000000000")
    '===035) Pagamento atrav�s de carn�.......................................(11 posi��es - total = 689).
    'a) C�digo do banco.
    lSequencia = lSequencia & Format$(pobjstcPedido.Cod_Bco_Cobr, "000")
    'b) Vencimento da primeira parcela.
    'lSequencia = lSequencia & Format$(pobjstcPedido.Ven_Parcela1_VM, "00000000")
    '===036) Pagamento atrav�s de d�bito em conta.............................(18 posi��es - total = 707).

    'a) C�digo da ag�ncia.
    lSequencia = lSequencia & Format$(pobjstcPedido.Cod_Agen_Bas_Cor, "0000")
    'b) C�digo da conta corrente.
    lSequencia = lSequencia & Format$(pobjstcPedido.Num_Cta_Bas_Cor, "000000000")
    ''c) Dia para d�bito.
    'lSequencia = lSequencia & Format$(pobjstcPedido.Deb_Dia, "00")
    For Each lObjstcVenc In pobjstcPedido.ColecaoPedVenc
        lSequencia = lSequencia & Format$(lObjstcVenc.Num_Parc, "00")
        lSequencia = lSequencia & Format$(lObjstcVenc.DAT_VENC, "00000000")
    Next lObjstcVenc

    '===037) Juros de mora.....................................................(4 posi��es - total = 711).
    lSequencia = lSequencia & "0450"

    'a) Dados dos itens.
    If pobjstcPedido.ColecaoPedLoc.Count = 0 Then
        pMensagem = "FTE0056 - " & gPrefixo & _
                    "N�o localizado registro de item (c�lculo do c�digo de unicidade)."
        GoTo SAIDA
    End If

    '4. Continua seq��ncia para c�lculo do c�digo de unicidade com dados dos itens.
    For Each lobjstcPedLoc In pobjstcPedido.ColecaoPedLoc
        'a) Obt�m valores dos itens.
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Num_Item, "00000")
        lAuxiliar = ""
        If gfPreenchido(lobjstcPedLoc.Cod_Tip_Logr_Risco) Then
            lAuxiliar = Trim$(lobjstcPedLoc.Cod_Tip_Logr_Risco) & " "
        End If
        '== Risco - Logradouro.
        If gfPreenchido(lobjstcPedLoc.Nom_Logr_Risco) Then
            lAuxiliar = lAuxiliar & Trim$(lobjstcPedLoc.Nom_Logr_Risco) & " "
        End If
        '== Risco - Complemento.
        If gfPreenchido(lobjstcPedLoc.Dsc_Compl_Risco) Then
            lAuxiliar = lAuxiliar & Trim$(lobjstcPedLoc.Dsc_Compl_Risco) & " "
        End If
        If val(lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco) <> 0 Then
            lAuxiliar = lAuxiliar & Trim$(lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco) & " "
        End If
        If Len(lAuxiliar) > 73 Then
            lAuxiliar = Left$(lAuxiliar, 73)
        Else
            lAuxiliar = lAuxiliar & String(73 - Len(lAuxiliar), " ")
        End If
        lSequencia = lSequencia & lAuxiliar
        '== Risco - Bairro.
        If gfPreenchido(lobjstcPedLoc.Nom_Bairro_Risco) Then
            lAuxiliar = lobjstcPedLoc.Nom_Bairro_Risco
            If Len(lAuxiliar) < 30 Then
                lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
            End If
        Else
            lAuxiliar = String(30, " ")
        End If
        lSequencia = lSequencia & lAuxiliar
        '== Risco - Cidade.
        If gfPreenchido(lobjstcPedLoc.Nom_Cid_Risco) Then
            lAuxiliar = lobjstcPedLoc.Nom_Cid_Risco
            If Len(lAuxiliar) < 30 Then
                lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
            End If
        Else
            lAuxiliar = String(30, " ")
        End If
        lSequencia = lSequencia & lAuxiliar
        '== Risco - UF.
        If gfPreenchido(lobjstcPedLoc.Sig_UF_Risco) Then
            lAuxiliar = lobjstcPedLoc.Sig_UF_Risco
            If Len(lAuxiliar) < 2 Then
                lAuxiliar = lAuxiliar & String(2 - Len(lAuxiliar), " ")
            End If
        Else
            lAuxiliar = String(2, " ")
        End If
        lSequencia = lSequencia & lAuxiliar
        '== Risco - CEP.
        If gfPreenchido(lobjstcPedLoc.Num_Cep_Risco) Then
            lAuxiliar = lobjstcPedLoc.Num_Cep_Risco
            If Len(lAuxiliar) < 8 Then
                lAuxiliar = String(8 - Len(lAuxiliar), "0") & lAuxiliar
            End If
        Else
            lAuxiliar = String(8, "0")
        End If
        lSequencia = lSequencia & lAuxiliar
        '== Desconto progressivo.
        lAuxiliar = Format$(lobjstcPedLoc.Per_Desc_Progres, "00.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
        '== IS B�sica
        lAuxiliar = Format$(lobjstcPedLoc.Per_IS_Vr, "000.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 3) & Right$(lAuxiliar, 2)
        '= Coeficiente IS/VR
        lAuxiliar = Format$(lobjstcPedLoc.Val_Coef_IS_Vr, "000.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 3) & Right$(lAuxiliar, 2)
        '== Valor de risco
        lAuxiliar = Format$(lobjstcPedLoc.Val_Vr, "000000000000.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
        '== Classifica��o
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Tab_Clasf, "0000")
        lSequencia = lSequencia & lobjstcPedLoc.Cod_Clasf
        lSequencia = lSequencia & lobjstcPedLoc.Cod_Compl_Clasf
        '== Constru��o
        lSequencia = lSequencia & lobjstcPedLoc.Cod_Constr
        lAuxiliar = Format$(lobjstcPedLoc.Per_Bonus, "00.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
        '==Quantidade de empregados.
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Qtd_Empreg, "0000")

        lAuxiliar = Format$(lobjstcPedLoc.Per_Tot_Segur, "00.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
        '== Percentual total de prote��o.
        lAuxiliar = Format$(lobjstcPedLoc.Per_Tot_Protec, "00.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
        '== Desconto de aparelho.
        lAuxiliar = Format$(lobjstcPedLoc.Per_Desc_Apar, "00.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
        '== produto da tabela de regi�o.
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Cod_Produto, "0000")
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Cod_Regiao, "0000")
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Cod_Iden_Tarif, "0000")
        '== Tipo de emiss�o
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Tip_Emissao, "00")
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Num_Apol_Anterior, "0000000000")
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Num_Item_Anterior, "00000")
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Ind_Sinistro, "0")
        lSequencia2 = lSequencia2 & lSequencia
        lSequencia = ""
        For Each lobjstcCobert In lobjstcPedLoc.ColecaoPedCobert
            If lobjstcCobert.Num_Item = lobjstcPedLoc.Num_Item Then
                lSequencia = lSequencia & Format$(lobjstcCobert.Cod_Cobert, "0000")
                lAuxiliar = Format$(lobjstcCobert.Per_Desc_Coml, "00.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Per_Desc_Exp, "00.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Per_Franq, "00.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Per_Taxa, "00.00000")
                lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 5)
                lAuxiliar = Format$(lobjstcCobert.Val_Franq_Max, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lSequencia = lSequencia & Format$(lobjstcCobert.Tab_Cobert, "0000")
                lAuxiliar = Format$(lobjstcCobert.Val_Franq_Min, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Val_IS, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.VAL_PR, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Val_Pr_Anu, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Val_Pr_Resg, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Val_Pr_Resg_Anu, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Val_Pr_Tarif, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Val_Pr_Tarif_Anu, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                'TAB_PED_COBERT_COMPL
                lSequencia = lSequencia & Format$(lobjstcCobert.Num_Meses_PI, "00")
                lAuxiliar = Format$(lobjstcCobert.Val_Risco_Cob, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lSequencia = lSequencia & Format$(lobjstcCobert.COD_CLASSE, "0000")
                lAuxiliar = Format$(lobjstcCobert.Val_Fator_Franq, "000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 3) & Right$(lAuxiliar, 2)
            End If
        Next

        'Ficha - 1000556 - Inclus�o de dados de Beneficiarios (e demais) na transmiss�o
        Dim lobjstcTextoEspec As stcA46V726B    'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).

        For Each lobjstcTextoEspec In pobjstcPedido.ColecaoTextoEspec
            If gfPreenchido(lobjstcTextoEspec.Espec_lin1) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin1)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin2) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin2)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin3) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin3)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin4) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin4)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin5) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin5)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin6) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin6)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin7) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin7)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin8) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin8)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin9) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin9)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin10) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin10)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin11) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin11)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin12) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin12)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin13) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin13)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin14) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin14)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin15) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin15)
        Next

        '== Insp_Obr
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Insp_Obr, "0") * -1
        '== Insp_Agend
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Insp_Agend, "0") * -1
        '== Dig_Bens
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Dig_Bens, "0") * -1
        '== Cob_Bens
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Cob_Bens, "0") * -1

        '== Seguradora de renova��o
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Ren_Con_Seguradora, "00000")
        '== Apolice de renova��o
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Ren_Con_Apolice, "000000000000000")
        '== Item de renova��o
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Ren_Con_Item, "00000")

        '== Tip_Cobert
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Tip_Cobert, "0")

    Next
    lSequencia2 = lSequencia2 & lSequencia
    lSequencia = lSequencia2
    '-*-*-*-*-*

    '6. Preencher o campo de seq��ncia para que o tamanho fique em m�ltiplo de 10.
    lTamanho = Len(lSequencia)
    lResto = lTamanho Mod 10
    If lResto <> 0 Then
        lSequencia = lSequencia & String(10 - lResto, " ")
    End If

    '7. Dividir o campo de seq��ncia em campos de 10 posi��es cada.
    lTamanho = Len(lSequencia)
    lResultado = lTamanho / 10
    ReDim lCampo1(lResultado)
    For lI = 1 To lResultado
        lPosInicial = ((lI - 1) * 10) + 1
        lCampo1(lI) = Mid$(lSequencia, lPosInicial, 10)
    Next lI

    '8. Para cada campo de 10 posi��es, transforma cada caracter num valor ASC e soma estes valores.
    ReDim lCampo2(lResultado)
    For lI = 1 To lResultado
        lCampo2(lI) = Asc(Mid(lCampo1(lI), 1, 1)) + Asc(Mid(lCampo1(lI), 2, 1)) + _
                      Asc(Mid(lCampo1(lI), 3, 1)) + Asc(Mid(lCampo1(lI), 4, 1)) + _
                      Asc(Mid(lCampo1(lI), 5, 1)) + Asc(Mid(lCampo1(lI), 6, 1)) + _
                      Asc(Mid(lCampo1(lI), 7, 1)) + Asc(Mid(lCampo1(lI), 8, 1)) + _
                      Asc(Mid(lCampo1(lI), 9, 1)) + Asc(Mid(lCampo1(lI), 10, 1))
    Next lI

    '9. Somar todos os resultados obtidos.
    lTotal = 0
    For lI = 1 To lResultado
        lTotal = lTotal + lCampo2(lI)
    Next lI

    '10. Resultado: c�digo de verifica��o.
    lCodVerificacao = Right$(Format$(lTotal, "000000"), 6)

    '11. Atualiza registro de dados comuns.
    pobjstcPedido.Cod_Unicidade = lCodVerificacao
    pobjstcPedido.Cod_Ver_Unici = 1   'Primeira vers�o da unicidade

    gfCalCodUnicidade_V1 = True

SAIDA:
    GoSub FechaObj
    Exit Function

FechaObj:
    Set lobjstcCoss = Nothing
    Set lobjstcJuros = Nothing
    Set lObjstcVenc = Nothing
    Set lobjstcPedLoc = Nothing
    Set lobjstcCobert = Nothing
    Return
End Function

Public Sub mpNovaBuscaDescricaoFranquia(ByVal pobjstcPed, ByVal pobjstcItem, ByVal pTamArray As Integer, ByVal pCodCobert As Integer, ByRef pTextos() As String)
    Dim lCorretor As Double
    Dim lCod_Clasf As String
    Dim lCod_Compl As String
    Dim lCod_Regiao As Double
    Dim lCod_Produto As Double
    Dim lObjstcCorr As stcA46V705B    'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).
    Dim lValFranquia As Double
    Dim lTextos() As String
    Dim lBDP00Multi As ADODB.Connection
    Dim lMensagem As String
    Dim lobjstcCobert As stcA46V704B    'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
    Dim lRegiao As Integer
    Dim lAuxiliar As String
    Dim lAuxiliar1 As String
    Dim lSelect As String
    Dim lrst As clsYasRecordSet
    Dim lTextoFranquia As Boolean
    Dim lI As Integer
    Dim lTem_VideObs As String

    'CR9974: ERRO: Atualiza��o de modelo de texto no endosso.
    Dim lTarifa_Apolice As Integer

    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lCorretor = lObjstcCorr.Cod_Corr
    Else
        lCorretor = 0
    End If

    ReDim lTextos(pTamArray)

    lCod_Clasf = pobjstcItem.Cod_Clasf
    lCod_Compl = pobjstcItem.Cod_Compl_Clasf
    lCod_Regiao = pobjstcItem.gCod_RegiaoNova
    lCod_Produto = pobjstcItem.Cod_Produto

    'CR9974: ERRO: Atualiza��o de modelo de texto no endosso.
    gNum_Apolice = pobjstcPed.Num_Apol_An
    lTarifa_Apolice = CInt(gTarifa_Apolice)


    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, _
                     lBDP00Multi, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        Exit Sub
    End If

    'Ficha 1000284
    'Dim lValFranquia As Double

    For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert

        If pCodCobert <> 0 Then
            If pCodCobert <> lobjstcCobert.Cod_Cobert Then
                GoTo DesprezaTexto
            End If
        End If

        'lValFranquia = 0
        'lTem_VideObs = ""

        '???
        'Call mpProcuraValorDefault(pobjstcItem, pobjstcPed, lobjstcCobert.Cod_Cobert, lValFranquia, lTem_VideObs)

        '1001135 - Corretor Volks - Franquia Especial Concessionaria


        'If gCodProduto >= 1002 Then
        '    If (InStr(1, "00008001/00008002/00008003/00008004/00008005/00008006/00008008/00008011/00008014/00008028", pobjstcPed.objstcCtrlEmis.Cod_Corr, vbTextCompare) > 0) And _
             '       (Val(lCod_Clasf) = 22 And Val(lCod_Compl) = 13) And _
             '       (InStr(1, "0038/0064/0066", Format(lobjstcCobert.Cod_Cobert, "0000"), vbTextCompare) > 0) Then
        '        GoTo DesprezaTexto
        '    Else
        '        If lTem_VideObs <> "S" And lValFranquia <> lobjstcCobert.Val_Franq_Min Then
        '            GoTo DesprezaTexto
        '        End If
        '    End If
        'Else
        '    If lValFranquia > 0 And lValFranquia <> lobjstcCobert.Val_Franq_Min Then
        '        GoTo DesprezaTexto
        '    End If
        'End If

        lRegiao = 0

        'Vai procurar na tabela de Ajustes pontuais se tem uma regi�o definida com franquia diferenciada
        If gfFranquiaRegiao(e_Tip_Valor_Regiao, TpoValor, pobjstcPed.Cod_Ramo, pobjstcPed.Cod_Estip, lCod_Clasf, lCod_Compl, lobjstcCobert.Cod_Cobert, pobjstcPed.Cod_Produto, pobjstcItem.Cod_Constr, pobjstcItem.Num_Cep_Risco, lRegiao, lMensagem) = False Then
            Exit Sub
        End If

        lAuxiliar = ";" & pobjstcPed.Cod_Estip & "/" & lRegiao & "/" & Trim(lCod_Clasf) & "/" & Trim(lCod_Compl) & "/" & lCorretor & "/" & lobjstcCobert.Cod_Cobert & ";"
        If lAuxiliar1 Like "*" & lAuxiliar & "*" Then
            GoTo DesprezaTexto
        End If

        lAuxiliar1 = lAuxiliar1 & lAuxiliar

        'Consulta por Corretor + Classifica��o
        If pobjstcPed.Cod_Dizer = 12 Then    ' CR9974 - SE FOR ENDOSSO DE ALTERA��O BUSCAR AS FRANQUIAS DA TARIFA QUE A APOLICE FOI EMITIDA
            lSelect = " Select * from TAB_DESC_FRANQUIA WITH (NOLOCK) Where"
            lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and Cod_Produto = " & pobjstcPed.Cod_Estip
            lSelect = lSelect & "  and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
            lSelect = lSelect & "  and Compl_Clasf = '" & Trim(lCod_Clasf) & Trim(lCod_Compl) & "'"   '1000871 - Inclus�o de Org�o Categoria Conselho
            lSelect = lSelect & "  and (Cod_Corretor = " & lCorretor & " or Cod_Corretor = 0)"
            lSelect = lSelect & "  and (Cod_Cobert = " & lobjstcCobert.Cod_Cobert & ")"
            lSelect = lSelect & " and Cod_Tarf_Ini <= " & lTarifa_Apolice
            lSelect = lSelect & " and Cod_Tarf_Final >= " & lTarifa_Apolice
            lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"
        Else
            lSelect = " Select * from TAB_DESC_FRANQUIA WITH (NOLOCK) Where"
            lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and Cod_Produto = " & pobjstcPed.Cod_Estip
            lSelect = lSelect & "  and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
            lSelect = lSelect & "  and Compl_Clasf = '" & Trim(lCod_Clasf) & Trim(lCod_Compl) & "'"   '1000871 - Inclus�o de Org�o Categoria Conselho
            lSelect = lSelect & "  and (Cod_Corretor = " & lCorretor & " or Cod_Corretor = 0)"
            lSelect = lSelect & "  and (Cod_Cobert = " & lobjstcCobert.Cod_Cobert & ")"
            lSelect = lSelect & " and Cod_Tarf_Ini <= " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and Cod_Tarf_Final >= " & pobjstcPed.Cod_Produto
            lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"
        End If


        If gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
            Call gpGraLog(1, lMensagem)
            Exit Sub
        End If

        If lrst.EOF Then
            'Consulta por Corretor + todas as classifica��es
            If pobjstcPed.Cod_Dizer = 12 Then
                lSelect = " Select * from TAB_DESC_FRANQUIA WITH (NOLOCK) Where"
                lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
                lSelect = lSelect & " and Cod_Produto = " & pobjstcPed.Cod_Estip
                lSelect = lSelect & "  and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
                lSelect = lSelect & "  and (Compl_Clasf ='0')"
                lSelect = lSelect & "  and (Cod_Corretor = " & lCorretor & ")"
                lSelect = lSelect & "  and (Cod_Cobert = " & lobjstcCobert.Cod_Cobert & ")"
                lSelect = lSelect & " and Cod_Tarf_Ini <= " & lTarifa_Apolice
                lSelect = lSelect & " and Cod_Tarf_Final >= " & lTarifa_Apolice
                lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"
            Else
                lSelect = " Select * from TAB_DESC_FRANQUIA WITH (NOLOCK) Where"
                lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
                lSelect = lSelect & " and Cod_Produto = " & pobjstcPed.Cod_Estip
                lSelect = lSelect & "  and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
                lSelect = lSelect & "  and (Compl_Clasf ='0')"
                lSelect = lSelect & "  and (Cod_Corretor = " & lCorretor & ")"
                lSelect = lSelect & "  and (Cod_Cobert = " & lobjstcCobert.Cod_Cobert & ")"
                lSelect = lSelect & " and Cod_Tarf_Ini <= " & pobjstcPed.Cod_Produto
                lSelect = lSelect & " and Cod_Tarf_Final >= " & pobjstcPed.Cod_Produto
                lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"
            End If

            If gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
                Call gpGraLog(1, lMensagem)
                Exit Sub
            End If
        End If

        If lrst.EOF Then
            'Consulta por todos os corretores + todas as classifica��es
            If pobjstcPed.Cod_Dizer = 12 Then
                lSelect = " Select * from TAB_DESC_FRANQUIA WITH (NOLOCK) Where"
                lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
                lSelect = lSelect & " and Cod_Produto = " & pobjstcPed.Cod_Estip
                lSelect = lSelect & " and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
                lSelect = lSelect & " and (Compl_Clasf ='0')"
                lSelect = lSelect & " and (Cod_Corretor = 0)"
                lSelect = lSelect & " and (Cod_Cobert = " & lobjstcCobert.Cod_Cobert & ")"
                lSelect = lSelect & " and Cod_Tarf_Ini <= " & lTarifa_Apolice
                lSelect = lSelect & " and Cod_Tarf_Final >= " & lTarifa_Apolice
                lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"
            Else
                lSelect = " Select * from TAB_DESC_FRANQUIA WITH (NOLOCK) Where"
                lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
                lSelect = lSelect & " and Cod_Produto = " & pobjstcPed.Cod_Estip
                lSelect = lSelect & " and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
                lSelect = lSelect & " and (Compl_Clasf ='0')"
                lSelect = lSelect & " and (Cod_Corretor = 0)"
                lSelect = lSelect & " and (Cod_Cobert = " & lobjstcCobert.Cod_Cobert & ")"
                lSelect = lSelect & " and Cod_Tarf_Ini <= " & pobjstcPed.Cod_Produto
                lSelect = lSelect & " and Cod_Tarf_Final >= " & pobjstcPed.Cod_Produto
                lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"
            End If

            If gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
                Call gpGraLog(1, lMensagem)
                Exit Sub
            End If
        End If

        If lrst.EOF Then GoTo DesprezaTexto

        'ADICIONADO PARA S� MOSTRAR A DESCRICAO QUANDO TIVER "*".
        If Left(lobjstcCobert.gTem_VideObs, 1) <> "*" Then GoTo DesprezaTexto

        Do While Not lrst.EOF

            Select Case lobjstcCobert.Cod_Cobert
            Case gfBuscarDePara(39, 1), gfBuscarDePara(40, 1), gfBuscarDePara(41, 1), gfBuscarDePara(42, 1), gfBuscarDePara(43, 1), gfBuscarDePara(44, 1), _
                gfBuscarDePara(47, 1), gfBuscarDePara(50, 1)
                If pobjstcPed.Cod_Estip < 1020 Then
                    lTextoFranquia = True
                End If
            End Select
            lTextos(lI) = lrst("Descricao")
            lI = lI + 1
            If lI > pTamArray Then ReDim Preserve lTextos(lI)

Pula_Texto:
            lrst.MoveNext
        Loop
DesprezaTexto:
    Next

    If lTextoFranquia = True Then
        lTextos(lI) = "Franquia: 7 (sete) dias"
        lI = lI + 1
    End If

    Call gpFecha2(lBDP00Multi)

FinalTextos:
    pTextos = lTextos
    Set lobjstcCobert = Nothing
    Set lObjstcCorr = Nothing
    Exit Sub
ControleErro:
    Call gpGraLog(1, Err.Description)
End Sub












Public Function gfFranquiaRegiao(ByVal pCod_Tip_Valor As Integer, ByVal pTip_Dado As TipoDados, ByVal pCod_Ramo As Integer, ByVal pCod_Prod As Integer, ByRef pCod_Clasf As String, ByRef pCod_Compl_Clasf As String, ByVal pCod_Cobert As Integer, ByVal pCod_Tarifa As Long, ByVal pCod_Constr As Integer, ByVal pCepRisco As Long, ByRef ValorAjustavel As Variant, ByRef pMensagem As String) As Boolean
    Dim lbdTbm_Parameter As ADODB.Connection      'Arquivo: P00Tarifa.
    Dim lrsTbm_Parameter As clsYasRecordSet       'Registro : Tbm_Parameter.
    Dim lSelect As String
    On Error GoTo ControleErros

    pMensagem = ""
    gfFranquiaRegiao = False

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTbm_Parameter, pMensagem) = False Then
        Exit Function
    End If

    lSelect = " Select Cod_Cobert, Cod_Tip_Chave, Chave_Ini, Chave_Ter, Valor from Tbm_Parameter "
    lSelect = lSelect & " Where Cod_Ramo = " & pCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & pCod_Prod
    lSelect = lSelect & " and Cod_Corr in (0, " & gCod_Corretor & ")"
    lSelect = lSelect & " and ((rtrim(ltrim(Cod_Clasf)) = '" & Trim(pCod_Clasf) & "'"
    lSelect = lSelect & " and Cod_Compl_Clasf = '" & pCod_Compl_Clasf & "') or rtrim(ltrim(Cod_Clasf)) = '0' )"
    lSelect = lSelect & " and Cod_Cobert in (0," & pCod_Cobert & ")"
    lSelect = lSelect & " and Cod_Tarif = " & pCod_Tarifa
    lSelect = lSelect & " and Cod_Tip_Valor = " & pCod_Tip_Valor
    lSelect = lSelect & " and Cod_Constr in (0," & pCod_Constr & ")"
    lSelect = lSelect & " order by Cod_Cobert, Cod_Prioridade, Cod_Constr desc "

    If gfObtRegistro(lbdTbm_Parameter, lSelect, lrsTbm_Parameter, pMensagem) = False Then
        Call gpFecha2(lbdTbm_Parameter)
        Exit Function
    End If

    Do While Not lrsTbm_Parameter.EOF
        If CDblx(Check(lrsTbm_Parameter("Cod_Tip_Chave"))) = 3 Then
            If CDblx(Check(lrsTbm_Parameter("Chave_Ini"))) <= pCepRisco And CDblx(Check(lrsTbm_Parameter("Chave_Ter"))) >= pCepRisco Then
                ValorAjustavel = gfTratarNulo(lrsTbm_Parameter("Valor"), pTip_Dado)
            End If
        End If
        lrsTbm_Parameter.MoveNext
    Loop
    gfFranquiaRegiao = True
    Call gpFecha2(lbdTbm_Parameter)

    Exit Function
ControleErros:
    gfFranquiaRegiao = False
    pMensagem = "mfFranquiaRegiao ERRO - " & Err & " " & Err.Description & _
                vbCrLf & " SQL: " & lSelect
    Call gpGraLog(0, pMensagem)

End Function

Sub mpProcuraValorDefault(ByVal pobjstcPedLoc As stcA46V708B, ByVal pobjstcPed As stcA46V702B, ByVal pCod_Cobert As Integer, ByRef pValFranquia As Double, ByRef pTem_VideObs As String)
    Dim Dummy1 As Double
    Dim Dummy4 As Integer
    Dim Dummy5 As String

    Call mpNovoBuscaLimFranquia(pobjstcPedLoc.Cod_Ramo, pobjstcPed.Cod_Estip, pobjstcPed.Cod_Produto, _
                                pCod_Cobert, pobjstcPedLoc.Cod_Clasf, pobjstcPedLoc.Cod_Compl_Clasf, _
                                pobjstcPedLoc.Cod_Constr, pobjstcPedLoc.Num_Cep_Risco, pobjstcPed.Cod_Sequencia, pobjstcPed.objstcCtrlEmis.Cod_Corr, _
                                pValFranquia, Dummy1, Dummy1, Dummy1, Dummy5, Dummy5, Dummy4)


End Sub

Public Sub mpNovoBuscaLimFranquia(ByVal pCodRamo As Integer, ByVal pCodModalidade As Integer, ByVal pCodTarifa As Integer, _
                                  ByVal pCodCobert As Integer, ByVal pCod_Clasf As String, ByVal pCod_Compl_Clasf As String, _
                                  ByVal pCod_Constr As Integer, ByVal pCEP As Long, ByVal pCodSequencia As Integer, ByVal pCodCorretor As Long, _
                                  ByRef ValFraqMin As Double, ByRef Perc_Franquia As Double, ByRef Fator_Ajuste As Double, _
                                  ByRef ValFranqMax As Double, ByRef TextoFranquia As String, ByRef Dsc_Cober_Franq As String, ByVal pNumPiMeses As Integer)

    Dim lbdTbm_Fat_Franquia As ADODB.Connection         'Arquivo: P00Multi.
    Dim lrsTbm_Fat_Franquia As clsYasRecordSet          'Registro : Tbm_Fator_Franquia.
    Dim rstTabClas As clsYasRecordSet
    Dim lMensagem As String         'Mensagem.
    Dim lCodSequencia As Integer
    Dim lSelect As String
    Dim lTip_Atividade As Integer
    Dim Tem_VideObs As String

    On Error GoTo ControleErro

    ValFraqMin = 0
    Perc_Franquia = 0
    Fator_Ajuste = 0
    ValFranqMax = 0
    Dsc_Cober_Franq = ""
    TextoFranquia = ""

    lCodSequencia = pCodSequencia

NovaSequencia1:

    '1 - Procura Registro de exce��o com a chave completa
    'Procura o valor do Val_Franq_Min e Perc_Franq

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTbm_Fat_Franquia, lMensagem) = False Then
        Call gpGraLog(2, lMensagem)
        Exit Sub
    End If


    lSelect = " Select Val_Franq_Min, Fator_Ajuste, Perc_Franquia, VAL_FRANQ_MAX, Tem_VideObs,  Dsc_Cober_Franq  "
    lSelect = lSelect & " from Tbm_Fator_Franquia "
    lSelect = lSelect & " Where Cod_Ramo = " & pCodRamo
    lSelect = lSelect & " and Cod_Modalidade = " & pCodModalidade
    lSelect = lSelect & " and Cod_Tip_Reg = 'E' "
    If pCodTarifa >= 940 Then
        lSelect = lSelect & " and ((Cod_Clasf = '" & pCod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & pCod_Compl_Clasf & "')"
        lSelect = lSelect & " or (Cod_Clasf = '0'))"
    Else
        lSelect = lSelect & " and Cod_Clasf = '" & pCod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & pCod_Compl_Clasf & "'"
    End If
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & " and Cod_Construcao in (0 , " & pCod_Constr & ")"
    lSelect = lSelect & " and Cod_Tarif = " & pCodTarifa
    lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia

    If pCodTarifa >= 940 Then
        lSelect = lSelect & " AND ((CEP_INICIO <= " & pCEP
        lSelect = lSelect & " AND CEP_FIM >= " & pCEP & " ) "
        lSelect = lSelect & " OR (CEP_INICIO =  0 ))"
        lSelect = lSelect & " order by Prioridade desc "
    End If

    If gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
        Call gpFecha2(lbdTbm_Fat_Franquia)
        Call gpGraLog(2, " mpNovoBuscaLimFranquia " & lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
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

    If lrsTbm_Fat_Franquia("Val_Franq_Min") = 0 And lrsTbm_Fat_Franquia("Perc_Franquia") = 0 And Check(lrsTbm_Fat_Franquia("Tem_VideObs")) = "" Then
        GoTo ProcuraDefault
    End If

    GoTo PopulaCampos

ProcuraDefault:

    lCodSequencia = pCodSequencia
    lTip_Atividade = 0

    '3. Obt�m produto sem t�rmino de vig�ncia ou com t�rmino de vig�ncia superior � data do dia.
    If Trim(pCod_Clasf) <> "" And Trim(pCod_Clasf) <> "0" Then
        lSelect = " Select distinct dsc_clasf, COD_CLASF, " & _
                  "COD_COMPL_CLASF , tab_Clasf, Tip_Atividade  from Tbm_Classificacao " & _
                " Where cod_Ramo  = " & pCodRamo & _
                " and  tab_clasf = " & pCodModalidade & _
                " and cod_Clasf = '" & pCod_Clasf & "'" & _
                " and Cod_Compl_Clasf = '" & pCod_Compl_Clasf & "'" & _
                " and Cod_Tarif = " & pCodTarifa & _
                " and Cod_Sequencia <= " & lCodSequencia


        If gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, rstTabClas, lMensagem) = False Then
            Call gpFecha2(lbdTbm_Fat_Franquia)
            Call gpGraLog(1, " mpNovoBuscaLimFranquia " & lMensagem)
            Exit Sub
        End If
        If Not rstTabClas.EOF Then
            lTip_Atividade = rstTabClas("Tip_Atividade")
        End If

    End If

    '3 - Procura o registro somente com Ramo, Modalidade, Cobertura e TipoAtividade

NovaSequencia2:


    lSelect = " Select Val_Franq_Min, Fator_Ajuste, Perc_Franquia, VAL_FRANQ_MAX, Tem_VideObs,  Dsc_Cober_Franq   "
    lSelect = lSelect & " from Tbm_Fator_Franquia "
    lSelect = lSelect & " Where Cod_Ramo = " & pCodRamo
    lSelect = lSelect & " and Cod_Modalidade = " & pCodModalidade
    lSelect = lSelect & " and Cod_Tip_Reg = 'D' "
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & " and Tip_Atividade in (3," & lTip_Atividade & ")"
    lSelect = lSelect & " and Cod_Tarif = " & pCodTarifa
    lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia

    If gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
        Call gpFecha2(lbdTbm_Fat_Franquia)
        Call gpGraLog(2, " mpNovoBuscaLimFranquia " & lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If lrsTbm_Fat_Franquia.EOF Then
        If lCodSequencia > 0 Then
            lCodSequencia = lCodSequencia - 1
            GoTo NovaSequencia2
        Else
            GoTo FimBusca
        End If
    End If

PopulaCampos:

    ValFraqMin = Check(lrsTbm_Fat_Franquia("Val_Franq_Min"))
    Perc_Franquia = Check(lrsTbm_Fat_Franquia("Perc_Franquia"))
    Fator_Ajuste = Check(lrsTbm_Fat_Franquia("Fator_Ajuste"))
    ValFranqMax = CDblx(Check(lrsTbm_Fat_Franquia("VAL_FRANQ_MAX")))
    Tem_VideObs = Check(lrsTbm_Fat_Franquia("Tem_VideObs"))
    Dsc_Cober_Franq = Check(lrsTbm_Fat_Franquia("Dsc_Cober_Franq"))


FimBusca:
    TextoFranquia = gfTextoFranquia(ValFraqMin, Perc_Franquia, Tem_VideObs, pCodCobert, pNumPiMeses)

    Call gpFecha2(lbdTbm_Fat_Franquia)

    '1001135 - Corretor Volks - Franquia Especial Concessionaria
    If pCodTarifa >= 1002 Then
        If (InStr(1, "00008001/00008002/00008003/00008004/00008005/00008006/00008008/00008011/00008014/00008028", Format(pCodCorretor, "00000000"), vbTextCompare) > 0) And _
           (val(pCod_Clasf) = 22 And val(pCod_Compl_Clasf) = 13) And _
           (InStr(1, "0038/0064/0066", Format(pCodCobert, "0000"), vbTextCompare) > 0) Then
            Perc_Franquia = 15
            ValFraqMin = 2000
        End If
    End If
    Exit Sub
ControleErro:
    Call gpGraLog(2, " mpNovoBuscaLimFranquia " & Err & " - " & ERROR)

End Sub

Public Function gfTextoFranquia(ByVal ValFraqMin As Double, ByVal Perc_Franquia As Double, ByVal Tem_VideObs As String, ByVal pCodCobert As Integer, _
                                ByVal pNum_Pi_Meses As Integer, Optional pValIs As Double, Optional pobjstcPedLoc As stcA46V708B) As String
    Dim TextoFranquia As String
    Dim lrsFranqNovoLim As clsYasRecordSet
    '* 10% dos prej. Indeniz, limitado ao m�nimo de R$ 10.000,00

    TextoFranquia = ""

    If gCodProduto >= 1080 Then
        If Not pobjstcPedLoc Is Nothing Then
            Call mfFranqNovoLim(pCodCobert, pValIs, pobjstcPedLoc.Cod_Clasf, pobjstcPedLoc.Cod_Compl_Clasf, _
                                pobjstcPedLoc.Cod_Constr, pobjstcPedLoc.Num_Cep_Risco, pobjstcPedLoc.Tip_Cobert, pNum_Pi_Meses, lrsFranqNovoLim)
            If lrsFranqNovoLim.RecordCount > 0 Then
                If Me.gfPreenchido(Me.gfTratarNulo(lrsFranqNovoLim!Texto, tpoString)) Then
                    TextoFranquia = TextoFranquia & mfMontaTexofranquia(Perc_Franquia, IIf(pNum_Pi_Meses = 0, ValFraqMin, 0), IIf(pNum_Pi_Meses > 0, ValFraqMin, 0), Me.gfTratarNulo(lrsFranqNovoLim!Texto, tpoString), Me.gfTratarNulo(lrsFranqNovoLim!Compl, tpoString), Me.gfTratarNulo(lrsFranqNovoLim!Tem_VideObs, tpoString))
            End If
        End If
    End If
    
        If TextoFranquia = Empty Or TextoFranquia = "*" Then
        If Perc_Franquia > 0 Then
            If ValFraqMin > 0 Then
                TextoFranquia = TextoFranquia & Perc_Franquia & "% dos prej. Indeniz, limitado ao m�nimo de R$ " & Format(ValFraqMin, "###,##0.00")
            Else
                TextoFranquia = TextoFranquia & Perc_Franquia & "% dos prejuizos "
            End If
        Else
                If pNum_Pi_Meses > 0 And ValFraqMin > 0 Then
                TextoFranquia = ValFraqMin & " dia(s)"
            Else
            If ValFraqMin > 0 Then
                'TextoFranquia = TextoFranquia & "Limitado ao m�nino de R$ " & Format(ValFraqMin, "###,##0.00")
                TextoFranquia = TextoFranquia & "R$ " & Format(ValFraqMin, "###,##0.00")
            Else
                If pNum_Pi_Meses > 0 Then
                    TextoFranquia = TextoFranquia & gfBuscaTextoPiIndenitario(pCodCobert, ByVal pNum_Pi_Meses)
                End If
            End If
        End If
        End If
    End If
    
    If Tem_VideObs = "*" Then
        If Left(TextoFranquia, 1) <> "*" Then
            TextoFranquia = "*" + TextoFranquia
        End If
    Else
        If Left(TextoFranquia, 1) = "*" Then
            TextoFranquia = Right(TextoFranquia, Len(TextoFranquia) - 1)
        End If
    End If
    
    Else
        If Tem_VideObs = "*" Then
            TextoFranquia = "*"
        End If
        
        If Perc_Franquia > 0 Then
            If ValFraqMin > 0 Then
                TextoFranquia = TextoFranquia & Perc_Franquia & "% dos prej. Indeniz, limitado ao m�nimo de R$ " & Format(ValFraqMin, "###,##0.00")
            Else
                TextoFranquia = TextoFranquia & Perc_Franquia & "% dos prejuizos "
            End If
        Else
            If ValFraqMin > 0 Then
                TextoFranquia = TextoFranquia & "Limitado ao m�nino de R$ " & Format(ValFraqMin, "###,##0.00")
            Else
                If pNum_Pi_Meses > 0 Then
                    TextoFranquia = TextoFranquia & gfBuscaTextoPiIndenitario(pCodCobert, ByVal pNum_Pi_Meses)
                End If
            End If
        End If
    End If
    
    gfTextoFranquia = gfRetirarCaracterEspecial(TextoFranquia, True)
End Function




Public Function gfTratarVazio(pCampo As String) As String
'Fun��o: trata campo Vazio.
    If IsNull(pCampo) Then
        gfTratarVazio = " "
    ElseIf Len(pCampo) = 0 Then
        gfTratarVazio = " "
    Else
        gfTratarVazio = pCampo
    End If
End Function

Public Function gfVerTransmissao() As Boolean
'Fun��o: verifica se corretor pode realizar transmiss�o de propostas.

    Dim lCodCorretor1 As String    'C�digo do corretor no arquivo P2842400.EXE.
    Dim lCodCorretor2 As String    'C�digo do corretor no arquivo P366CORR.EXE.
    Dim lEncontrou As Boolean    'Indica que encontrou o registro de controle.
    Dim lEndP2842400 As String    'Endere�o do arquivo P2842400.EXE.
    Dim lEndP366CORR As String    'Endere�o do arquivo P366CORR.EXE.
    Dim lIndTransmissao As String   'Indicador de transmiss�o: S ou N.
    Dim lNumP2842400 As Integer    'Utilizado com FreeFile para o arquivo P2842400.EXE.
    Dim lNumP366CORR As Integer    'Utilizado com FreeFile para o arquivo P366CORR.EXE.
    Dim lRegistro As String    'Registro.

'*********************************************************
'*****DESATIVADO    TODOS OS CORRETORES PODEM TRANSMITIR PROPOSTAS
'*****MULTISSEGUROS

    gfVerTransmissao = True

    Exit Function
    '*********************************************************
    '*********************************************************
    '*********************************************************
    '*********************************************************
    '    gfVerTransmissao = False
    '
    '    '1. Obt�m c�digo do corretor.
    '    lEndP2842400 = gAppPath & "P2842400.EXE"
    '    lNumP2842400 = FreeFile
    '    On Error Resume Next
    '    Open lEndP2842400 For Input Access Read As lNumP2842400
    '    If Err = 0 Then     'Encontrou arquivo P2842400.EXE.
    '        Line Input #lNumP2842400, lRegistro
    '        lCodCorretor1 = Format(lRegistro, "00000000")
    '    Else
    '        lCodCorretor1 = "00000000"
    '    End If
    '    On Error GoTo 0
    '    Close lNumP2842400
    '
    '    '2. Verifica preenchimento.
    '    If Not gfPreenchido(lCodCorretor1) Then
    '        Call gpGraLog(1, "FGE0030 - Para transmiss�o de propostas para a Yasuda via Internet, " & _
             '                         "solicite o termo de ""ADES�O AO SISTEMA DE TRANSMISS�O DE PROPOSTAS DE " & _
             '                         "MULTISEGUROS"" com o produtor (c�digo do corretor n�o preenchido).")
    '        Exit Function
    '    End If
    '    lCodCorretor1 = Trim$(lCodCorretor1)
    '    If Not IsNumeric(lCodCorretor1) Then
    '        Call gpGraLog(1, "FGE0076 - Para transmiss�o de propostas para a Yasuda via Internet, " & _
             '                         "solicite o termo de ""ADES�O AO SISTEMA DE TRANSMISS�O DE PROPOSTAS DE " & _
             '                         "MULTISEGUROS"" com o produtor (c�digo do corretor n�o num�rico).")
    '        Exit Function
    '    End If
    '    If Val(lCodCorretor1) = 0 Then
    '        'Call gpGraLog(1, "FGE0077 - Para transmiss�o de propostas para a Yasuda via Internet, " & _
             '        '                 "solicite o termo de ""ADES�O AO SISTEMA DE TRANSMISS�O DE PROPOSTAS DE " & _
             '        '                 "MULTISEGUROS"" com o produtor (c�digo do corretor zerado).")
    '        Exit Function
    '    End If
    '    Select Case Len(lCodCorretor1)
    '        Case Is < 7
    '            lCodCorretor1 = String(7 - Len(lCodCorretor1), "0") & lCodCorretor1
    '        Case Is > 7
    '            lCodCorretor1 = Right$(lCodCorretor1, 7)
    '    End Select
    '
    '    '3. Abre aquivo de controle de transmiss�o (tempor�rio).
    '    lEndP366CORR = gAppPath & "P366CORR.EXE"
    '    If Not gfPreenchido(Dir(lEndP366CORR)) Then
    '        Call gpGraLog(1, "FTE0061 - Arquivo n�o localizado: controle do termo de ades�o.")
    '        Exit Function
    '    End If
    '    lNumP366CORR = FreeFile
    '    On Error Resume Next
    '    Open lEndP366CORR For Input Access Read As lNumP366CORR
    '    If Err <> 0 Then
    '        Call gpGraLog(1, "FTE0062 - Erro ao abrir arquivo de controle do termo de ades�o (" & Err & _
             '                         " - " & Error & ").")
    '        On Error GoTo 0
    '        Exit Function
    '    End If
    '    On Error GoTo 0
    '
    '    '4. Trata registros do arquivo de inicializa��o.
    '    lEncontrou = False
    '    Do While Not EOF(lNumP366CORR)
    '        Line Input #lNumP366CORR, lRegistro
    '        If gfPreenchido(lRegistro) Then
    '            lRegistro = Trim$(lRegistro)
    '            If UCase(lRegistro) = "[GRUPOS]" Then
    '                Exit Do
    '            End If
    '            If UCase(lRegistro) = "[SOMENTE MULTI]" Then
    '                Exit Do
    '            End If
    '            Select Case Len(lRegistro)
    '                Case 8
    '                    lCodCorretor2 = Left$(lRegistro, 7)
    '                    lIndTransmissao = "N"
    '                Case 9
    '                    lCodCorretor2 = Left$(lRegistro, 7)
    '                    lIndTransmissao = Mid$(lRegistro, 9, 1)
    '                Case Else
    '                    lCodCorretor2 = "0000000"
    '                    lIndTransmissao = "N"
    '            End Select
    '            If lCodCorretor1 = lCodCorretor2 Then
    '                lEncontrou = True
    '                Exit Do
    '            End If
    '        End If
    '    Loop
    '    Close lNumP366CORR
    '    If lEncontrou = False Then
    '        Call gpGraLog(1, "FGE0078 - Para transmiss�o de propostas para a Yasuda via Internet, " & _
             '                         "solicite o termo de ""ADES�O AO SISTEMA DE TRANSMISS�O DE PROPOSTAS DE " & _
             '                         "MULTISEGUROS"" com o produtor (c�digo do corretor inexistente no controle de " & _
             '                         "transmiss�o).")
    '        Exit Function
    '    End If
    '    If lIndTransmissao <> "S" Then
    '        Call gpGraLog(1, "FGE0079 - Para transmiss�o de propostas para a Yasuda via Internet, " & _
             '                         "solicite o termo de ""ADES�O AO SISTEMA DE TRANSMISS�O DE PROPOSTAS DE " & _
             '                         "MULTISEGUROS"" com o produtor (corretor sem autoriza��o).")
    '        Exit Function
    '    End If

    gfVerTransmissao = True
End Function

Private Sub mpCarregaCorretor(ByRef pCod_Corretor As Long)
'Fun��o: Carrega c�digo do corretor.

'Par�metro de sa�da...pCod_Corretor...C�digo do corretor.

    Dim lArquivo As String    'Endere�o do arquivo.
    Dim lNumArquivo As Integer  'Utilizado com FreeFile.
    Dim lRegistro As String    'Registro.

'1. Inicializa vari�vel.
    pCod_Corretor = 0

    '2. Obt�m dados no arquivo P2842400.EXE.
    lArquivo = gAppPath & "P2842400.EXE"
    lNumArquivo = FreeFile
    On Error Resume Next
    Open lArquivo For Input As #lNumArquivo
    If Err = 0 Then
        Line Input #lNumArquivo, lRegistro
        If gfPreenchido(lRegistro) Then
            If IsNumeric(lRegistro) Then
                If Len(lRegistro) >= 8 Then
                    pCod_Corretor = CLng(Mid$(lRegistro, 2, 7))
                Else
                    lRegistro = Format(lRegistro, "00000000")
                    pCod_Corretor = CLng(Mid$(lRegistro, 2, 7))
                End If
            End If
        End If
    End If
    On Error GoTo 0
    Close #lNumArquivo
End Sub
Public Sub gpForCNPJ(ByVal pOpcao As Byte, ByVal pEntrada As String, ByRef pSaida As String)
'Procedure: formata campo de CNPJ.

'Par�metros de entrada...pOpcao.....1 = coloca ".", "/" e "-"; 2 = tira ".", "/" e "-".
'                        pEntrada...Campo de entrada.
'Par�metro de sa�da......pSaida.....Capo de sa�da (n�o pode ser o mesmo campo de entrada).

    Dim lAuxiliar1 As String  'Auxiliar 1.
    Dim lAuxiliar2 As String  'Auxiliar 2.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lPosicao As Byte  'Utilizado com InStr.
    Dim lTamanho As Byte  'Tamanho.

    Select Case pOpcao
    Case 1      'Coloca ".", "/" e "-".
        'a) Inicializa campo de sa�da.
        pSaida = ""
        'b) Se campo de entrada n�o preenchido, nada a fazer.
        If Not gfPreenchido(pEntrada) Then
            Exit Sub
        End If
        'c) Se campo de entrada com 18 posi��es, nada a fazer.
        lAuxiliar1 = Trim(pEntrada)
        lTamanho = Len(lAuxiliar1)
        If lTamanho >= 18 Then
            Exit Sub
        End If
        'd) Prepara campo auxiliar.
        lPosicao = InStr(lAuxiliar1, "-")
        If lPosicao <> 0 Then
            Select Case Len(lAuxiliar1)
            Case lPosicao + 1
                lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & "0" & Right$(lAuxiliar1, 1)
            Case lPosicao + 2
                lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & Right$(lAuxiliar1, 2)
            Case Else
                Exit Sub
            End Select
        End If
        lPosicao = InStr(lAuxiliar1, "/")
        If lPosicao <> 0 Then
            Select Case Len(lAuxiliar1)
            Case lPosicao + 3
                lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & "000" & Right$(lAuxiliar1, 3)
            Case lPosicao + 4
                lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & "00" & Right$(lAuxiliar1, 4)
            Case lPosicao + 5
                lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & "0" & Right$(lAuxiliar1, 5)
            Case lPosicao + 6
                lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & Right$(lAuxiliar1, 6)
            Case Else
                Exit Sub
            End Select
        End If
        lTamanho = Len(lAuxiliar1)
        lAuxiliar2 = ""
        For lI = 1 To lTamanho
            Select Case Mid$(lAuxiliar1, lI, 1)
            Case ".", ","
                'Nada a fazer.
            Case Else
                lAuxiliar2 = lAuxiliar2 & Mid$(lAuxiliar1, lI, 1)
            End Select
        Next lI
        If Len(lAuxiliar2) < 14 Then
            lAuxiliar2 = String(14 - Len(lAuxiliar2), "0") & lAuxiliar2
        End If
        'e) Valida CNPJ - 12 primeiras posi��es alfanum�ricas, 2 �ltimas num�ricas (d�gitos verificadores).
        If Not gfCNPJValido(lAuxiliar2) Then
            pSaida = ""
            Exit Sub
        End If
        lAuxiliar2 = UCase(lAuxiliar2)
        'f) Formata CNPJ - XX.XXX.XXX/XXXX-XX.
        pSaida = Mid$(lAuxiliar2, 1, 2) & "." & _
                 Mid$(lAuxiliar2, 3, 3) & "." & _
                 Mid$(lAuxiliar2, 6, 3) & "/" & _
                 Mid$(lAuxiliar2, 9, 4) & "-" & _
                 Mid$(lAuxiliar2, 13, 2)
    Case 2      'Retira ".", "/" e "-"
        'a) Inicializa campo de sa�da.
        pSaida = "00000000000000"
        'b) Se campo de entrada diferente de 18 posi��es, nada a fazer.
        lAuxiliar1 = Trim(pEntrada)
        lTamanho = Len(lAuxiliar1)
        If lTamanho <> 18 Then
            Exit Sub
        End If
        'c) Se campo de entrada fora do formato XX.XXX.XXX/XXXX-XX, nada a fazer.
        If Mid$(lAuxiliar1, 3, 1) <> "." Or _
           Mid$(lAuxiliar1, 7, 1) <> "." Or _
           Mid$(lAuxiliar1, 11, 1) <> "/" Or _
           Mid$(lAuxiliar1, 16, 1) <> "-" Then
            Exit Sub
        End If
        'd) Prepara campo auxiliar.
        lAuxiliar2 = Mid$(lAuxiliar1, 1, 2) & _
                     Mid$(lAuxiliar1, 4, 3) & _
                     Mid$(lAuxiliar1, 8, 3) & _
                     Mid$(lAuxiliar1, 12, 4) & _
                     Mid$(lAuxiliar1, 17, 2)
        If Not gfCNPJValido(lAuxiliar2) Then
            Exit Sub
        End If
        pSaida = UCase(lAuxiliar2)
    End Select
End Sub

Public Function gfCNPJValido(ByVal pCNPJ As String) As Boolean
'Fun��o: valida formato de CNPJ (14 posi��es - 12 alfanum�ricas + 2 d�gitos verificadores num�ricos).
'Aceita tanto o CNPJ num�rico tradicional quanto o novo CNPJ alfanum�rico da Receita Federal.

    Dim lI          As Byte     'Utilizado com For...Next.
    Dim lCaracter   As String   'Caracter sendo analisado.

    gfCNPJValido = False
    If Len(pCNPJ) <> 14 Then
        Exit Function
    End If
    For lI = 1 To 12
        lCaracter = UCase(Mid$(pCNPJ, lI, 1))
        If Not ((lCaracter >= "0" And lCaracter <= "9") Or (lCaracter >= "A" And lCaracter <= "Z")) Then
            Exit Function
        End If
    Next lI
    For lI = 13 To 14
        lCaracter = Mid$(pCNPJ, lI, 1)
        If lCaracter < "0" Or lCaracter > "9" Then
            Exit Function
        End If
    Next lI
    gfCNPJValido = True
End Function

Public Sub gpForCPF(ByVal pOpcao As Byte, ByVal pEntrada As String, ByRef pSaida As String)
'Procedure: formata campo de CPF.

'Par�metros de entrada...pOpcao.....1 = coloca "." e "-"; 2 = tira "." e "-".
'                        pEntrada...Campo de entrada.
'Par�metro de sa�da......pSaida.....Capo de sa�da (n�o pode ser o mesmo campo de entrada).

    Dim lAuxiliar1 As String  'Auxiliar 1.
    Dim lAuxiliar2 As String  'Auxiliar 2.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lPosicao As Byte  'Utilizado com InStr.
    Dim lTamanho As Byte  'Tamanho.

    Select Case pOpcao
    Case 1      'Coloca "." e "-"
        'a) Inicializa campo de sa�da.
        pSaida = ""
        'b) Se campo de entrada n�o preenchido, nada a fazer.
        If Not gfPreenchido(pEntrada) Then
            Exit Sub
        End If
        'c) Se campo de entrada com 14 posi��es, nada a fazer.
        lAuxiliar1 = Trim(pEntrada)
        lTamanho = Len(lAuxiliar1)
        If lTamanho >= 14 Then
            pSaida = lAuxiliar1
            Exit Sub
        End If
        'd) Prepara campo auxiliar.
        lPosicao = InStr(lAuxiliar1, "-")
        If lPosicao <> 0 Then
            If Len(lAuxiliar1) = lPosicao + 1 Then
                lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & "0" & Right$(lAuxiliar1, 1)
            ElseIf Len(lAuxiliar1) = lPosicao + 2 Then
                lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & Right$(lAuxiliar1, 2)
            Else
                pSaida = lAuxiliar1
                Exit Sub
            End If
        End If
        lTamanho = Len(lAuxiliar1)
        lAuxiliar2 = ""
        For lI = 1 To lTamanho
            Select Case Mid$(lAuxiliar1, lI, 1)
            Case ".", ","
                'Nada a fazer.
            Case Else
                lAuxiliar2 = lAuxiliar2 & Mid$(lAuxiliar1, lI, 1)
            End Select
        Next lI
        If Not IsNumeric(lAuxiliar1) Then
            pSaida = lAuxiliar1
            Exit Sub
        End If
        If Len(lAuxiliar2) < 11 Then
            lAuxiliar2 = String(11 - Len(lAuxiliar2), "0") & lAuxiliar2
        End If
        'e) Formata CPF - XXX.XXX.XXX-XX.
        pSaida = Mid$(lAuxiliar2, 1, 3) & "." & _
                 Mid$(lAuxiliar2, 4, 3) & "." & _
                 Mid$(lAuxiliar2, 7, 3) & "-" & _
                 Mid$(lAuxiliar2, 10, 2)
    Case 2      'Retira "." e "-"
        'a) Inicializa campo de sa�da.
        pSaida = "00000000000"
        'b) Se campo de entrada diferente de 14 posi��es, nada a fazer.
        lAuxiliar1 = Trim(pEntrada)
        lTamanho = Len(lAuxiliar1)
        If lTamanho <> 14 Then
            Exit Sub
        End If
        'c) Se campo de entrada fora do formato XXX.XXX.XXX-XX, nada a fazer.
        If Mid$(lAuxiliar1, 4, 1) <> "." Or _
           Mid$(lAuxiliar1, 8, 1) <> "." Or _
           Mid$(lAuxiliar1, 12, 1) <> "-" Then
            Exit Sub
        End If
        'd) Prepara campo auxiliar.
        lAuxiliar2 = Mid$(lAuxiliar1, 1, 3) & _
                     Mid$(lAuxiliar1, 5, 3) & _
                     Mid$(lAuxiliar1, 9, 3) & _
                     Mid$(lAuxiliar1, 13, 2)
        If Not IsNumeric(lAuxiliar2) Then
            Exit Sub
        End If
        pSaida = lAuxiliar2
    End Select
End Sub

Public Function gfNomUniProdutiva(ByVal pCodUniNegocios As Integer, ByRef pNomUniNegocios As String, _
                                  ByRef pMensagem As String) As Boolean
'Procedure: obt�m nome da unidade de neg�cios.

'Par�metro de entrada...pCodUniNegocios...C�digo da unidade de neg�cios.
'Par�metros de sa�da....pNomUniNegocios...Nome da unidade de neg�cios.
'                       pMensagem.........Mensagem.

    Dim bdP0044700 As ADODB.Connection    'Arquivo - P0044700.
    Dim rsDEPT As clsYasRecordSet    'Registro - YTAB_DEPT.
    Dim lSelect As String

    gfNomUniProdutiva = False

    '1. Inicializa controle.
    pNomUniNegocios = " "

    '2. Verifica se c�digo est� preenchido.
    If IsNull(pCodUniNegocios) Then
        pMensagem = "FNP0008 - Unidade de neg�cios inv�lida (c�digo n�o preenchido)."
        GoTo SAIDA
    End If

    '3. Busca no arquivo.
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        GoTo SAIDA
    End If
    lSelect = "SELECT NOM_DEPT FROM YTAB_DEPT WHERE NUM_DEPT = " & pCodUniNegocios
    If gfObtRegistro(bdP0044700, lSelect, rsDEPT, pMensagem) = False Then
        GoTo SAIDA
    End If
    If rsDEPT.EOF = True Then
        Call Me.gpFecha2(bdP0044700)
        pMensagem = "FNL0010 - C�digo da unidade de neg�cios n�o localizado (" & pCodUniNegocios & ")."
        GoTo SAIDA
    End If

    '4. Posiciona nome da unidade produtiva.
    pNomUniNegocios = rsDEPT!NOM_DEPT

    gfNomUniProdutiva = True

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Call Me.gpFecha2(bdP0044700)
    Set bdP0044700 = Nothing
    Set rsDEPT = Nothing
    Return
End Function
Public Function gfNomProdutor(ByVal pCodProdutor As Integer, ByRef pNomProdutor As String, _
                              ByRef pMensagem As String) As Boolean
'Fun��o: obt�m nome do produtor.

'Par�metro de entrada...pCodProdutor...C�digo do produtor.
'Par�metros de sa�da....pNomProdutor...Nome do produtor.
'                       pMensagem......Mensagem.

    Dim bdP0044700 As ADODB.Connection    'Arquivo - P0044700.
    Dim rsPRODUTOR As clsYasRecordSet    'Registro - YTAB_PRODUTOR.
    Dim lSelect As String

    gfNomProdutor = False

    '1. Inicializa controle.
    pNomProdutor = " "

    '2. Verifica se c�digo est� prenchido.
    If Not gfPreenchido(pCodProdutor) Then
        pMensagem = "FNP0012 - " & gPrefixo & "Produtor inv�lido (c�digo do produtor n�o preenchido)."
        GoTo SAIDA
    End If

    '3. Obt�m nome do produtor.
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        GoTo SAIDA
    End If
    lSelect = "SELECT NOM_PROD FROM YTAB_PRODUTOR WHERE COD_PROD = " & pCodProdutor
    If gfObtRegistro(bdP0044700, lSelect, rsPRODUTOR, pMensagem) = False Then
        GoTo SAIDA
    End If
    If rsPRODUTOR.EOF = True Then
        pNomProdutor = "PRODUTOR " & pCodProdutor & "."
    Else
        pNomProdutor = rsPRODUTOR!NOM_PROD
    End If

    gfNomProdutor = True

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Call Me.gpFecha2(bdP0044700)
    Set bdP0044700 = Nothing
    Set rsPRODUTOR = Nothing
    Return
End Function

Public Function gfBusProduto(ByVal pVigInicio As Double, ByRef pCodProduto As Integer, _
                             ByRef pCodProtocolo As Long, ByRef pCodSequencia As Long, _
                             ByRef pDescProduto As String, ByRef pMensagem As String, _
                             Optional pCodRamo As Integer, Optional pEndosso As Boolean = False) As Boolean
'Procedure: busca produto de acordo com in�cio de vig�ncia.

'Par�metro de entrada...pVigInicio......In�cio de vig�ncia da cota��o/proposta.
'Par�metros de sa�da....pCodProduto.....C�digo do produto.
'                       pCodProtocolo...C�digo do protocolo.
'                       pCodSequencia...C�digo da seq��ncia.
'                       pMensagem.......Mensagem.

    Dim BDP00Tarifa As ADODB.Connection    'Arquivo - P00Tarifa.
    Dim lAuxAno As Integer  'Auxiliar: ano.
    Dim lAuxiliar As String     'Auxiliar.
    Dim lAuxMes As Integer  'Auxiliar: m�s.
    Dim lDatLimite As Double      'Data limite.
    Dim lAuxData As Double    'Data auxiliar
    Dim lAuxPreData As String       'Auxiliar.
    Dim lDias As Integer    'Dias.
    Dim lDifDias As Integer   'Dif dias.
    Dim rsTabProd As clsYasRecordSet    'TAB_PRODUTO: tabela.
    Dim rsChavesConfig As clsYasRecordSet     'Registro - Tbm_Configuracao_App.
    Dim lSelect As String
    Dim lMensagem As String
    
    'Lucas Garcia - Liberar Tarifa Um Dia Antes Da Passagem.
    Dim pProdFinalizado As Integer

    gfBusProduto = False

    If Trim(gTarifaVigente) <> "" And pCodProduto = 0 Then
        pCodProduto = CInt(gTarifaVigente)
    End If

    '2. Abre arquivo.
    If gfAbrBasDados("P00MULTI", "1", gEndP00MULTI, BDP00Tarifa, pMensagem) = False Then
        GoTo SAIDA
    End If

    'Lucas Garcia - Liberar Tarifa Um Dia Antes Da Passagem.
    pProdFinalizado = 0

    Dim lServidor As String
    If UBound(GetIpAddrTable) = 1 Then
        lServidor = GetIpAddrTable(1)
    Else
        lServidor = "0.0.0.0"
    End If
    
    lSelect = "EXEC proc_RetornaChaveSimples 'TesteTarifa'"
    If M99V001.M46V999.gfObtRegistro(BDP00Tarifa, lSelect, rsChavesConfig, lMensagem) = False Then
        Call M99V001.M46V999.gpFecha2(BDP00Tarifa)
        Call M99V001.M46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        GoTo SAIDA
    End If
       
    If Not rsChavesConfig.EOF Then
       
        Do While Not rsChavesConfig.EOF
            
            If lServidor = rsChavesConfig.Fields("Value") Then
                M99V001.M46V999.gTarifaTeste = 9
            End If
        
            rsChavesConfig.MoveNext
        Loop
        
    End If
    
    If Trim(gTarifaTeste) <> "" Then
        pProdFinalizado = 9
    End If

    '1. Inicializa vari�veis.
    If Not pCodProduto > 0 Then
        pCodProduto = 0
    End If
    pCodProtocolo = 0
    pCodSequencia = 0


    '3. Obt�m produto sem t�rmino de vig�ncia ou com t�rmino de vig�ncia superior � data do dia.
    'lSelect = "SELECT * FROM TAB_PRODUTO WHERE Tip_Prot = 1 AND Dat_Vig_Ini <= " & pVigInicio

    'Tip_prot  = 1  Tarifa Corretor
    'Tip_prot  = 2  Tarifa Interna
    If pCodProduto > 0 Then
        lSelect = "SELECT * FROM TAB_PRODUTO WITH (NOLOCK) WHERE cod_produto = " & pCodProduto
    Else
        lSelect = "SELECT * FROM TAB_PRODUTO WITH (NOLOCK) WHERE Dat_Vig_Ini <= " & pVigInicio
    End If

    'Lucas Garcia - Liberar Tarifa Um Dia Antes Da Passagem.
    If pProdFinalizado > 0 Then
        lSelect = "SELECT * FROM TAB_PRODUTO WITH (NOLOCK) WHERE "
        lSelect = lSelect & " PROD_FINALIZADO = 9" '9 = EmTeste
    Else
        'Lucas Garcia - Liberar Tarifa Um Dia Antes Da Passagem.
        lSelect = lSelect & " AND PROD_FINALIZADO <> 9" '9 = EmTeste
    End If

    If gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda Or Me.M46V111.gCod_plano > 1000 Then
        lSelect = lSelect & " and Tip_Prot = 1 "
    End If
    If pCodRamo <> 0 Then
        lSelect = lSelect & " and Cod_ramo = " & pCodRamo
    End If
    
   
    lSelect = lSelect & " ORDER BY Cod_Produto DESC , Cod_Sequencia desc"
       
    If gfObtRegistro(BDP00Tarifa, lSelect, rsTabProd, pMensagem) = False Then
        GoTo SAIDA
    End If
    
    If rsTabProd.EOF And pProdFinalizado = 9 Then
        pMensagem = "F20082018 - Nenhuma tarifa aberta para teste."
        Exit Function
    End If

    If rsTabProd.EOF Then
        lSelect = "SELECT * FROM TAB_PRODUTO WITH (NOLOCK) WHERE Dat_Vig_Ini <= " & pVigInicio
        If pCodRamo <> 0 Then
            lSelect = lSelect & " and Cod_ramo = " & pCodRamo
        End If
        
        'Lucas Garcia - Liberar Tarifa Um Dia Antes Da Passagem.
        lSelect = lSelect & " AND PROD_FINALIZADO <> 9" '9 = EmTeste
        
        lSelect = lSelect & " ORDER BY Cod_Produto DESC"
        If gfObtRegistro(BDP00Tarifa, lSelect, rsTabProd, pMensagem) = False Then
            GoTo SAIDA
        End If
        If rsTabProd.EOF Then
            gfBusProduto = True
            GoTo SAIDA
        End If
    End If

    '4. Verifica produtos dispon�veis.
    If pVigInicio > rsTabProd!Dat_Vig_Ter Then
        lAuxiliar = Format$(rsTabProd!Dat_Vig_Ter, "00000000")
        lAuxMes = CInt(Mid$(lAuxiliar, 5, 2)) + 1
        lAuxAno = CInt(Mid$(lAuxiliar, 1, 4))
        If lAuxMes = 13 Then
            lAuxMes = 1
            lAuxAno = lAuxAno + 1
        End If
        lDias = -1
        On Error Resume Next
        lDias = rsTabProd!Dias_Expiracao
        If lDias = -1 Then
            lDias = 15
        End If
        On Error GoTo 0
        lAuxPreData = CDate(Mid(rsTabProd!Dat_Vig_Ter, 7, 2) & "/" & Mid(rsTabProd!Dat_Vig_Ter, 5, 2) & "/" & Mid(rsTabProd!Dat_Vig_Ter, 1, 4)) + lDias
        lAuxData = Format(lAuxPreData, "YYYYMMDD")

        If Not (pVigInicio >= rsTabProd!Dat_Vig_Ini And pVigInicio <= lAuxData) Then
            'Documento dentro do limite do produto (at� x dias ap�s t�rmino da
            'vig�ncia) - sem problema.
            '?         gfBusProduto = True
            '?         GoTo SAIDA
        End If
    End If

    '5. Posiciona dados de sa�da.
    pCodProduto = Me.Check(rsTabProd!Cod_Produto)
    pCodProtocolo = Me.Check(rsTabProd!Cod_Protocolo)
    pCodSequencia = Me.Check(rsTabProd!Cod_Sequencia)
    pDescProduto = Me.Check(rsTabProd("Nom_Prot"))
    gProCusEndosso = Me.Check(rsTabProd!Val_Cus_Emis_Endo)
    gProCusApolice = Me.Check(rsTabProd!Val_Cus_Emis_Apol)
    gProPreMinParcela = Me.Check(rsTabProd!Par_Min_Parcela)
    gProPreMinParcelaTAB = gProPreMinParcela
    If pEndosso Then
        gProPreMinEndosso = Me.Check(rsTabProd!Pre_Min_Endosso)
        gProPremioMinimoTAB = Me.Check(rsTabProd!Pre_Min_Endosso)
        gProPreMinParcelaEndosso = Me.Check(rsTabProd!Par_Min_Endosso)
    Else
        gProPremioMinimo = Me.Check(rsTabProd!Pre_Min_Parcela)
        gProPremioMinimoTAB = Me.Check(rsTabProd!Pre_Min_Parcela)
        gProPreMinParcela = Me.Check(rsTabProd!Par_Min_Parcela)
        gProPreMinParcelaTAB = gProPreMinParcela
    End If

    gfBusProduto = True

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Call Me.gpFecha2(BDP00Tarifa)
    Set rsTabProd = Nothing
    Set BDP00Tarifa = Nothing
    Return
End Function

Public Function gfObtCorr_Pro_Agenc(ByVal pTipCorretor As String, ByVal pOpcao As Byte, ByVal pCorretor As String, _
                                    ByRef pSaida As Variant, ByRef pMensagem As String) As Boolean
'Fun��o: obt�m dados de corretores.

'Par�metros de entrada...pTipCorretor P = Prolabori, A = Agenciamento
'                        pOpcao......1 = C�digo do corretor e 2 = Nome do corretor.
'                        pCorretor...C�digo ou nome do corretor, de acordo com pOpcao.
'Par�metro de sa�da......pSaida......Array contendo resultados.
'                                    (x, 1) C�digo do corretor.
'                                    (x, 2) Nome do corretor.
'                                    (x, 3) Nome abreviado do corretor.
'                                    (x, 4) Num. Susep

    Dim BDP00Tarifa As ADODB.Connection    'Arquivo - P00Tarifa.
    Dim lI As Double  'Contador.
    Dim lSelect As String       'Comando SELECT.
    Dim rsCORR As clsYasRecordSet    'Tabela: YTAB_CORR.


'Os Agenciadores s�o identificados pelo c�digo Susep 999999999 ou 888888888.
'Os Pro-labores s�o identificados pelo in�cio do c�digo do corretor = 05


    gfObtCorr_Pro_Agenc = False

    '1. Verifica par�metros.
    'a) Op��o.
    If Not gfPreenchido(pTipCorretor) Then
        Call gpGraLog(2, "FTE0067 - Op��o inv�lida (obter tipo corretor - op��o n�o preenchida)." & vbLf & _
                         vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If

    If Not gfPreenchido(pOpcao) Then
        Call gpGraLog(2, "FTE0067 - Op��o inv�lida (obter corretor - op��o n�o preenchida)." & vbLf & _
                         vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If

    If Not IsNumeric(pOpcao) Then
        Call gpGraLog(2, "FTE0067 - Op��o inv�lida (obter corretor - op��o n�o num�rica)." & vbLf & _
                         vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If
    If pOpcao <> 1 And pOpcao <> 2 Then
        Call gpGraLog(2, "FTE0067 - Op��o inv�lida (obter corretor - op��o diferente de 1 e 2)." & _
                         vbLf & vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If
    'b) Corretor.
    If Not gfPreenchido(pCorretor) Then
        Call gpGraLog(2, "FTE0068 - Erro ao obter corretor (c�digo do corretor n�o preenchido)." & _
                         vbLf & vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If

    '2. Busca informa��o.

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, BDP00Tarifa, pMensagem) = False Then
        Exit Function
    End If


    Select Case pOpcao
    Case 1      'C�digo do corretor.
        If Not IsNumeric(pCorretor) Then
            Call gpGraLog(2, "FTE0068 - Erro ao obter corretor (c�digo do corretor n�o num�rico)." & _
                             vbLf & vbLf & "O sistema ser� encerrado.")
            Exit Function    'End   ***MOSILVA Verificar
        End If
        lSelect = "SELECT CODCORR, NOMCORR, NOMABRVCORR , numsusep FROM DR00802C321 "
        lSelect = lSelect & " WHERE CODCORR = " & pCorretor & " AND CODATIVO = 1"
        If pTipCorretor = "P" Then
            lSelect = lSelect & " AND CODCORR >= 500000  "
        Else
            lSelect = lSelect & " AND NUMSUSEP IN ('99999999' , '88888888') "
        End If
        If gfObtRegistro(BDP00Tarifa, lSelect, rsCORR, pMensagem) = False Then
            Exit Function
        End If
        If rsCORR.RecordCount = 0 Then
            Call Me.gpFecha2(BDP00Tarifa)
            pMensagem = "FGE0003 - " & gPrefixo & "Corretor inv�lido (c�digo de corretor n�o " & _
                        "localizado - " & pCorretor & ")."
            Exit Function
        End If
        ReDim pSaida(1, 4)
        pSaida(1, 1) = rsCORR!codCorr
        pSaida(1, 2) = rsCORR!NOMCORR
        pSaida(1, 3) = rsCORR!NOMABRVCORR
        pSaida(1, 4) = rsCORR!NUMSUSEP
        Call Me.gpFecha2(BDP00Tarifa)
        gfObtCorr_Pro_Agenc = True
        Exit Function
    Case 2      'Nome do corretor.

        lSelect = "SELECT CODCORR, NOMCORR, NOMABRVCORR , numsusep FROM DR00802C321 "
        lSelect = lSelect & "WHERE ((NOMCORR LIKE '%" & pCorretor & "%' "
        lSelect = lSelect & "OR     NOMABRVCORR LIKE '%" & pCorretor & "%') "
        lSelect = lSelect & "OR (NOMCORR LIKE '*" & pCorretor & "*' "
        lSelect = lSelect & "OR     NOMABRVCORR LIKE '*" & pCorretor & "*')) "
        lSelect = lSelect & "AND   CODATIVO = 1"

        If pTipCorretor = "P" Then
            lSelect = lSelect & " AND CODCORR >= 500000  "
        Else
            lSelect = lSelect & " AND NUMSUSEP IN ('99999999' , '88888888') "
        End If



        If gfObtRegistro(BDP00Tarifa, lSelect, rsCORR, pMensagem) = False Then
            Exit Function
        End If
        If rsCORR.RecordCount = 0 Then
            Call Me.gpFecha2(BDP00Tarifa)
            pMensagem = "FGE0003 - Corretor inv�lido (corretor n�o localizado com nome ou parte " & _
                        "do nome = " & pCorretor & ")."
            Exit Function
        End If
        rsCORR.MoveLast
        ReDim Preserve pSaida(rsCORR.RecordCount, 4)
        lI = 0
        rsCORR.MoveFirst
        While Not rsCORR.EOF
            lI = lI + 1
            pSaida(lI, 1) = rsCORR!codCorr
            pSaida(lI, 2) = rsCORR!NOMCORR
            pSaida(lI, 3) = rsCORR!NOMABRVCORR
            pSaida(lI, 4) = rsCORR!NUMSUSEP
            rsCORR.MoveNext
        Wend
        Call Me.gpFecha2(BDP00Tarifa)
        gfObtCorr_Pro_Agenc = True
        Exit Function
    End Select
End Function

Public Function gfDifDatas(ByVal pInicio As Long, ByVal pTermino As Long) As Single
'Fun��o #014: calcula diferen�a entre datas.

'Par�metros de entrada...pInicio....data de in�cio.
'                        pTermino...data de t�rmino.

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "gfDifDatas"
    lKeyDynaTrace = DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

    Dim lAuxIniDias As String    'Auxiliar: valores para tabela para c�lculo da diferen�a de dias.
    Dim lAuxTerDias As String    'Auxiliar: valores para tabela para c�lculo da diferen�a de dias.
    Dim lDia As Single             '�ndice auxiliar: dia.
    Dim lI As Integer             '�ndice.
    Dim lIniDia As Single             'Dia do in�cio de vig�ncia.
    Dim lIniMes As Single             'M�s do in�cio de vig�ncia.
    Dim lMes As Single             '�ndice auxiliar: m�s.
    Dim lQuaIniDias As Single    'Quantidade de dias (in�cio de vig�ncia).
    Dim lQuaTerDias As Single    'Quantidade de dias (t�rmino de vig�ncia).
    Dim lTabIniDias(31, 12) As String   'Tabela para c�lculo de dias (data de in�cio).
    Dim lTabTerDias(31, 12) As String   'Tabela para c�lculo de dias (data de t�rmino).
    Dim lTerDia As Single             'Dia do t�rmino de vig�ncia.
    Dim lTerMes As Single             'M�s do t�rmino de vig�ncia.

    If M46V111.gCod_plano = 20 Then 'Imobili�rio
        gfDifDatas = CSng(DateDiff("d", gfForData(2, pInicio), gfForData(2, pTermino)))
        Exit Function
    End If

    Call DynaTrace.gravaLog(lKeyDynaTrace, "Inicializa valores auxiliares para c�lculo da diferen�a de dias.", 0)
'1. Inicializa valores auxiliares para c�lculo da diferen�a de dias.
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

    Call DynaTrace.gravaLog(lKeyDynaTrace, "Posiciona dados da tabela de data de in�cio.", 0)

    '2. Posiciona dados da tabela de data de in�cio.
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
    Call DynaTrace.gravaLog(lKeyDynaTrace, "Posiciona dados da tabela de data de t�rmino.", 0)

    '3. Posiciona dados da tabela de data de t�rmino.
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
    Call DynaTrace.gravaLog(lKeyDynaTrace, "Apropria informa��o de dia e m�s de in�cio e t�rmino.", 0)

    '4. Apropria informa��o de dia e m�s de in�cio e t�rmino.
    lIniDia = val(Mid(pInicio, 7, 2))
    lIniMes = val(Mid(pInicio, 5, 2))
    lTerDia = val(Mid(pTermino, 7, 2))
    lTerMes = val(Mid(pTermino, 5, 2))
    Call DynaTrace.gravaLog(lKeyDynaTrace, "Posiciona quantidade de dias.", 0)

    '5. Posiciona quantidade de dias.
    lQuaIniDias = lTabIniDias(lIniDia, lIniMes)
    If val(Mid(pInicio, 1, 4)) <> val(Mid(pTermino, 1, 4)) Then
        lQuaTerDias = lTabTerDias(lTerDia, lTerMes)
    Else
        lQuaTerDias = lTabIniDias(lTerDia, lTerMes)
    End If
    Call DynaTrace.gravaLog(lKeyDynaTrace, "Calcula diferen�a de dias.", 0)

    '6. Calcula diferen�a de dias.
    gfDifDatas = lQuaTerDias - lQuaIniDias
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

End Function
Public Function gfCalNosNumero(ByRef pNosNumero As String, ByRef pMensagem As String) As Boolean
'Fun��o: gera nosso n�mero a partir de usu�rio, ano, dia juliano e n�mero de segundos decorridos no
'        dia.

'Par�metros de sa�da...pNosNumero...Nosso n�mero.
'                      pMensagem....Mensagem.

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "gfCalNosNumero"
    lKeyDynaTrace = DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

    On Error GoTo TrataErro
    
'Formato do nosso n�mero:
'   AAADDDSSSSSUUUUUUUUD...AAA........ano.
'                          DDD........dia juliano (dias decorridos no ano).
'                          SSSSS......segundos decorridos no dia.
'                          UUUUUUUU...usu�rio.
'                          D..........d�gito verificador (pela Prova 11).

    Dim clsA99V001 As clsA99V001  'Calcula d�gitos.
    Dim lAuxAno As String   'Ano no formato string.
    Dim lAuxDia As String   'Dia no formato string.
    Dim lAuxSegundo As String       'Segundo no formato string.
    Dim lCodUsuario As String       'C�digo do usu�rio.
    Dim lDia As Integer    'Dia.
    Dim lDigito As Integer  'D�gito.
    Dim lMesBissexto As String        'Controle de dias do ano bissexto.
    Dim lMesNormal As String      'Controle de dias do ano normal.
    Dim lNosNumero As String      'Nosso n�mero.
    Dim lSegundo As Double    'Segundo.

    gfCalNosNumero = False

    '1. Posiciona dados para nosso n�mero.
    'a) Ano.
    lAuxAno = Right$(Format$(gHojAno, "0000"), 3)
    'b) Dia juliano.
    lMesNormal = "000031059090120150181212243273303334365"
    lMesBissexto = "000031060091121151182213244274304335366"
    lDia = gHojDia
    Select Case gHojAno Mod 4
    Case 0      'Ano bissexto.
        lDia = lDia + val(Mid(lMesBissexto, gHojMes * 3 - 2, 3))
    Case Else   'Ano N�O bissexto.
        lDia = lDia + val(Mid(lMesNormal, gHojMes * 3 - 2, 3))
    End Select
    lAuxDia = Right("000" & Trim(Str(lDia)), 3)
    'c) Segundos decorridos do dia.
    lSegundo = (Hour(Now) * 3600) + (Minute(Now) * 60) + Second(Now)
    lAuxSegundo = Right(Format$(lSegundo, "00000"), 5)
    'd) C�digo do usu�rio.
    'lCodUsuario = gNomUsuario
    'Simony

    Randomize
    lCodUsuario = Int((99999999 - 1 + 1) * Rnd + 1)

    If IsNumeric(lCodUsuario) Then
        lCodUsuario = Trim$(lCodUsuario)
        If Len(lCodUsuario) < 8 Then
            lCodUsuario = String(8 - Len(lCodUsuario), "0") & lCodUsuario
        End If
    Else
        If gCod_Corretor <> 0 Then
            lCodUsuario = Format(gCod_Corretor, "00000000")
        Else
            lCodUsuario = "00000000"
        End If
    End If
    'e) Monta nosso n�mero sem d�gito.
    lNosNumero = lAuxAno & lAuxDia & lAuxSegundo & lCodUsuario

    '2. Calcula d�gito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = Me

    If clsA99V001.DIG_MODULO11(lNosNumero, "20191817161514131211100908070605040302", lDigito, _
                               pMensagem) = False Then
        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "DIG_MODULO11 retornou false", 2, Class_Name, methodName)
        Exit Function
    End If
    Set clsA99V001 = Nothing

    '3. Completa nosso n�mero.
    pNosNumero = lNosNumero & Format$(lDigito, "0")
    gfCalNosNumero = True

    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function

TrataErro:
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, Err.Description, 2, Class_Name, methodName)
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function gfDisponivel(ByVal pNome As String) As Boolean
'Fun��o: indica se form est� dispon�vel.

    Dim lI As Byte    'Utilizado com For...Next.

    gfDisponivel = False
    For lI = 0 To (Forms.Count - 1)
        If UCase(Forms(lI).Name) = UCase(pNome) Then
            gfDisponivel = True
            Exit For
        End If
    Next lI
End Function
Public Function gfExecutaTransmissao() As Boolean
'Fun��o: Aciona componente EDIWare.
    ChDir gAppPath
    Shell gAppPath & "client.bat", vbMinimizedNoFocus
End Function
Public Function gfForAno(pData As Integer) As String
    If Left$(pData, 2) > 50 Then    'De 1951 a 1999.
        gfForAno = "19" & Format$(pData, "00")
    Else                            'De 2000 a 2050.
        gfForAno = "20" & Format$(pData, "00")
    End If
End Function
Public Function gfForAspas(ByVal pCampo As String) As String
'Fun��o: formata campo texto com aspas para ser utilizado no comando SQL.

    Dim lAuxiliar As String     'Utilizado com InStr.
    Dim lPosicao As Integer   'Utilizado com InStr.
    Dim lTemAspa As Boolean   'Indicador de que tem aspa.

'1. Se campo n�o preenchido, retorna campo nulo.
    If Not gfPreenchido(pCampo) Then
        gfForAspas = ""
        Exit Function
    End If

    '2. Posiciona controle.
    lTemAspa = True

    '3. Verifica se campo tem aspa.
    lPosicao = InStr(pCampo, "'")

    '4. Verifica se campo tem "|".
    If lPosicao = 0 Then
        gfForAspas = pCampo
        lTemAspa = False
        lPosicao = InStr(pCampo, "|")
        If lPosicao = 0 Then
            gfForAspas = pCampo
            Exit Function
        End If
    End If

    '5. Se campo tem aspa, duplica campo de aspa para comando SQL.
    lAuxiliar = pCampo
    If lTemAspa = True Then
        Do
            If lPosicao = Len(lAuxiliar) Then
                gfForAspas = lAuxiliar & "'"
                Exit Function
            End If
            lAuxiliar = Left$(lAuxiliar, lPosicao) & "'" & Mid$(lAuxiliar, lPosicao + 1)
            lPosicao = InStr(lPosicao + 2, lAuxiliar, "'")
            If lPosicao = 0 Then
                lPosicao = InStr(pCampo, "|")
                If lPosicao = 0 Then
                    gfForAspas = lAuxiliar
                    Exit Function
                Else
                    Exit Do
                End If
            End If
        Loop
    End If

    '6. Retira "|" do campo.
    Do
        If lPosicao = Len(lAuxiliar) Then
            gfForAspas = Left$(lAuxiliar, lPosicao - 1)
            Exit Function
        End If
        lAuxiliar = Left$(lAuxiliar, lPosicao - 1) & Mid$(lAuxiliar, lPosicao + 1)
        lPosicao = InStr(lPosicao + 1, lAuxiliar, "|")
        If lPosicao = 0 Then
            gfForAspas = lAuxiliar
            Exit Function
        End If
    Loop
End Function
Public Function gfForData(ByVal pOpcao As Byte, Optional ByVal pData, Optional ByVal pTipArquivo) _
       As String
'Fun��o: formata data.

'Formato dos campos de data (num�rico duplo): aaaammdd.

'Par�metros...pOpcao.........01 = formata campo de "dd/mm/aaaa" para "aaaammdd" para comandos SQL.
'                            02 = formata campo de "aaaammdd" para "dd/mm/aaaa".
'                            03 = fornece data atual no formato aaaammddhhmmss.
'                            04 = formata campo de "aaaammddhhmmss" para "dd/mm/aaaa hh:mm:ss".
'                            05 = formata campo de "aaaammdd" para formato conveniente
'                                 ('mm/dd/aaaa' ou #mm/dd/aaaa#).
'                            06 = formata campo de "dd/mm/aaaa" para formato conveniente
'                                 ('mm/dd/aaaa' ou #mm/dd/aaaa#).
'                            07 = formata campo de "aammdd" para "dd/mm/aaaa".
'                            08 = formata campo de ano para "aaaa".
'                            09 = trata campo de data na digita��o (utilizado nos eventos LostFocus).
'                            10 = formata campo de "aammdd" para "aaaammdd".
'                            11 = fornece data atual no formato aaaammdd.
'                            12 = formata campo de "aaaammdd" para "dd/mm/aaaa" se n�o for zero
'             pData..........Campo de data a ser tratado (opcional).
'             pTipArquivo...Tipo do arquivo (opcional - utilizado para op��o 5).

    Dim lAno As String    'Ano.
    Dim lAuxData As String    'Auxiliar: data.
    Dim lDia As String    'Dia.
    Dim lMes As String    'M�s.
    Dim lPosicao As Integer    'Utilizado com InStr.

'1. Verifica op��o.
    If Not gfPreenchido(pOpcao) Then
        Call gpGraLog(2, "FTE0064 - Op��o inv�lida (formata data - op��o n�o preenchida)." & vbLf & _
                         vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If
    If Not IsNumeric(pOpcao) Then
        Call gpGraLog(2, "FTE0064 - Op��o inv�lida (formata data - op��o n�o num�rica)." & vbLf & vbLf & _
                         "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If
    Select Case pOpcao
    Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
        'Nada a fazer - op��es v�lidas.
    Case Else
        Call gpGraLog(2, "FTE0064 - Op��o inv�lida (formata data - op��o desconhecida: " & pOpcao & _
                         ")." & vbLf & vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End Select

    '2. Processa formata��o de data.
    Select Case pOpcao
    Case 1
        If gfPreenchido(pData) Then
            If Len(pData) = 10 Then
                gfForData = Mid$(pData, 7, 4) & Mid$(pData, 4, 2) & Mid$(pData, 1, 2)
            Else
                gfForData = "00000000"
            End If
        Else
            gfForData = "00000000"
        End If
    Case 2
        If gfPreenchido(pData) Then
            lAuxData = Format$(pData, "00000000")
        Else
            lAuxData = "00000000"
        End If
        gfForData = Mid$(lAuxData, 7, 2) & "/" & Mid$(lAuxData, 5, 2) & "/" & Mid$(lAuxData, 1, 4)
    Case 3
        gfForData = Format$(Now, "yyyymmddhhnnss")
    Case 4
        If gfPreenchido(pData) Then
            lAuxData = Format$(pData, "00000000000000")
        Else
            lAuxData = "00000000000000"
        End If
        gfForData = Mid$(lAuxData, 7, 2) & "/" & _
                    Mid$(lAuxData, 5, 2) & "/" & _
                    Mid$(lAuxData, 1, 4) & " " & _
                    Mid$(lAuxData, 9, 2) & ":" & _
                    Mid$(lAuxData, 11, 2) & ":" & _
                    Mid$(lAuxData, 13, 4)
    Case 5
        lAuxData = Format$(pData, "00000000")
        Select Case pTipArquivo
        Case "1"        'Access.
            gfForData = "#" & Mid$(lAuxData, 5, 2) & "/" & Mid$(lAuxData, 7, 2) & "/" & _
                        Mid$(lAuxData, 1, 4) & "#"
        Case "2"        'SQL-Server
            gfForData = "'" & Mid$(lAuxData, 5, 2) & "/" & Mid$(lAuxData, 7, 2) & "/" & _
                        Mid$(lAuxData, 1, 4) & "'"
        End Select
    Case 6
        Select Case pTipArquivo
        Case "1"        'Access.
            gfForData = "#" & Mid$(pData, 4, 2) & "/" & Mid$(pData, 1, 2) & "/" & _
                        Mid$(pData, 7, 4) & "#"
        Case "2"        'SQL-Server
            gfForData = "'" & Mid$(pData, 4, 2) & "/" & Mid$(pData, 1, 2) & "/" & _
                        Mid$(pData, 7, 4) & "'"
        End Select
    Case 7
        If gfPreenchido(pData) Then
            lAuxData = Format$(pData, "000000")
        Else
            lAuxData = "000000"
        End If
        If Left$(lAuxData, 2) > 50 Then    'De 1951 a 1999.
            gfForData = Mid$(lAuxData, 5, 2) & "/" & _
                        Mid$(lAuxData, 3, 2) & "/19" & _
                        Left$(lAuxData, 2)
        Else                            'De 2000 a 2050.
            gfForData = Mid$(lAuxData, 5, 2) & "/" & _
                        Mid$(lAuxData, 3, 2) & "/20" & _
                        Left$(lAuxData, 2)
        End If
    Case 8
        If gfPreenchido(pData) Then
            lAuxData = Format$(pData, "0000")
        Else
            lAuxData = "0000"
        End If
        If Left$(lAuxData, 2) = "00" Then
            If Right$(lAuxData, 2) > 50 Then    'De 1951 a 1999.
                gfForData = "19" & Right$(lAuxData, 2)
            Else                            'De 2000 a 2050.
                gfForData = "20" & Right$(lAuxData, 2)
            End If
        Else
            gfForData = lAuxData
        End If
    Case 9
        If IsMissing(pData) Then
            gfForData = ""
        Else
            If Not gfPreenchido(pData) Then
                gfForData = ""
            Else
                gfForData = lAuxData
                If IsNumeric(pData) Then
                    Select Case Len(Trim(pData))
                    Case 6      'Prov�vel digita��o: ddmmaa.
                        lDia = Mid$(pData, 1, 2)
                        lMes = Mid$(pData, 3, 2)
                        lAno = Mid$(pData, 5, 2)
                        If val(lAno) > 50 Then  'De 1951 a 1999.
                            lAno = "19" & lAno
                        Else                    'De 2000 a 2050.
                            lAno = "20" & lAno
                        End If
                        gfForData = lDia & "/" & lMes & "/" & lAno
                    Case 8      'Prov�vel digita��o: ddmmaaaa.
                        gfForData = Mid$(pData, 1, 2) & "/" & _
                                    Mid$(pData, 3, 2) & "/" & _
                                    Mid$(pData, 5, 4)
                    Case Else
                        gfForData = pData
                    End Select
                Else
                    lAuxData = Trim$(pData)
                    'Dia.
                    lPosicao = InStr(lAuxData, "/")
                    Select Case lPosicao
                    Case 2
                        lDia = "0" & Mid$(lAuxData, 1, 1)
                    Case 3
                        lDia = Mid$(lAuxData, 1, 2)
                    Case Else
                        Exit Function
                    End Select
                    If Not IsNumeric(lDia) Then
                        Exit Function
                    End If
                    If Len(lAuxData) = lPosicao Then
                        Exit Function
                    End If
                    lAuxData = Mid$(lAuxData, lPosicao + 1)
                    'M�s.
                    lPosicao = InStr(lAuxData, "/")
                    Select Case lPosicao
                    Case 2
                        lMes = "0" & Mid$(lAuxData, 1, 1)
                    Case 3
                        lMes = Mid$(lAuxData, 1, 2)
                    Case Else
                        Exit Function
                    End Select
                    If Not IsNumeric(lMes) Then
                        Exit Function
                    End If
                    If Len(lAuxData) = lPosicao Then
                        Exit Function
                    End If
                    'Ano.
                    lAno = Mid$(lAuxData, lPosicao + 1)
                    If Not IsNumeric(lAno) Then
                        Exit Function
                    End If
                    If Len(lAno) = 1 Then
                        lAno = "0" & lAno
                    End If
                    If Len(lAno) = 2 Then
                        If val(lAno) > 50 Then  'De 1951 a 1999.
                            lAno = "19" & lAno
                        Else                    'De 2000 a 2050.
                            lAno = "20" & lAno
                        End If
                    End If
                    If Len(lAno) = 3 Then
                        lAno = "0" & lAno
                    End If
                    If Len(lAno) > 4 Then
                        lAno = Left$(lAno, 4)
                    End If
                    gfForData = lDia & "/" & lMes & "/" & lAno
                End If
                If Not IsDate(gfForData) Then
                    gfForData = ""
                End If
            End If
        End If
    Case 10
        If gfPreenchido(pData) Then
            lAuxData = Format$(pData, "000000")
            If Left$(lAuxData, 2) > 50 Then    'De 1951 a 1999.
                gfForData = "19" & lAuxData
            Else                            'De 2000 a 2050.
                gfForData = "20" & lAuxData
            End If
        Else
            lAuxData = "00000000"
        End If
    Case 11
        gfForData = Format$(Now, "yyyymmdd")
    Case 12
        gfForData = ""
        If gfPreenchido(pData) And CDblx(pData) <> 0 Then
            lAuxData = Format$(pData, "00000000")
            gfForData = Mid$(lAuxData, 7, 2) & "/" & Mid$(lAuxData, 5, 2) & "/" & Mid$(lAuxData, 1, 4)
        End If

    End Select
End Function
Public Function gfForCurrency(pValor As Variant, Optional pPreencheZero As Boolean = True) As String
    Dim Aux As String
    On Error GoTo MostraErro

    If Not IsNumeric(pValor) Then
        If pPreencheZero Then
            gfForCurrency = "0,00"
        End If
    ElseIf pValor = 0 Then
        If pPreencheZero Then
            gfForCurrency = "0,00"
        End If
    Else
        Aux = Replace(pValor, ".", "")
        gfForCurrency = Format(Aux, "##,###,###,###,###,##0.00")
    End If

    Exit Function
MostraErro:
    gpGraLog 1, "UTI001- Erro no gfForCurrency " & Err.Number & " - " & Err.Description

End Function


Public Function CDblx(ByRef Valor As Variant) As Double
    If IsNumeric(Valor) Then
        Valor = Valor
    Else
        Valor = 0
    End If
    CDblx = Valor
End Function


Public Function gfForValor(ByVal pOpcao As Byte, ByVal pCampo As String, Optional pPadrao As String) _
       As String
'Fun��o: formata valor.

'Par�metros...pOpcao....1 = troca v�rgula por ponto e tira ponto, para ser utilizado no comando SQL.
'                       2 = formata no padr�o standard e coloca brancos � esquerda para ser utilizado
'                           com os formul�rios.
'                       3 = coloca brancos � esquerda para ser utilizado com os formul�rios.
'             pCampo....campo a ser formatado.
'             pPadrao...para op��o 1, informa padr�o a ser utilizado caso campo n�o esteja preenchido.
'                       para op��o 2 e 3, informa tamanho do campo (se omitido, assume 15 posi��es).

    Dim lAuxiliar As String    'Auxiliar.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lTamanho As Byte  'Tamanho.
    Dim lTexto As String    'Auxiliar para tirar ponto.

'1. Verifica op��o.
    If Not gfPreenchido(pOpcao) Then
        Call gpGraLog(2, "FTE0065 - Op��o inv�lida (formata valor - op��o n�o preenchida)." & vbLf & _
                         vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If
    If Not IsNumeric(pOpcao) Then
        Call gpGraLog(2, "FTE0065 - Op��o inv�lida (formata valor - op��o n�o num�rica)." & vbLf & _
                         vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If
    If pOpcao <> 1 And pOpcao <> 2 And pOpcao <> 3 Then
        Call gpGraLog(2, "FTE0065 - Op��o inv�lida (formata valor - op��o diferente de 1, 2 e 3: " & _
                         pOpcao & ")." & vbLf & vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If

    '2. Processa formata��o de data.
    Select Case pOpcao
    Case 1
        If gfPreenchido(pCampo) Then
            If IsNumeric(pCampo) Then
                lAuxiliar = Trim$(pCampo)
                lTexto = ""
                For lI = 1 To Len(lAuxiliar)
                    Select Case Mid(lAuxiliar, lI, 1)
                    Case ","
                        lTexto = lTexto & "."
                    Case "."
                        'lTexto = lTexto
                    Case Else
                        lTexto = lTexto & Mid(lAuxiliar, lI, 1)
                    End Select
                Next lI
                If lTexto = "" Then
                    lTexto = "0"
                End If
                gfForValor = lTexto
                Exit Function
            End If
        End If
        If IsMissing(pPadrao) Then
            gfForValor = "0"
        Else
            If gfPreenchido(pPadrao) Then
                gfForValor = pPadrao
            Else
                gfForValor = "0"
            End If
        End If
    Case 2
        If gfPreenchido(pCampo) Then
            If IsNumeric(pCampo) Then
                lAuxiliar = Trim$(Format$(pCampo, "standard"))
            Else
                lAuxiliar = "0,00"
            End If
        Else
            lAuxiliar = "0,00"
        End If
        If IsMissing(pPadrao) Then
            lTamanho = 15
        Else
            If Not IsNumeric(pPadrao) Then
                lTamanho = 15
            Else
                lTamanho = pPadrao
            End If
        End If
        If Len(lAuxiliar) > lTamanho Then
            gfForValor = lAuxiliar
        Else
            gfForValor = Space(lTamanho - Len(lAuxiliar)) & lAuxiliar
        End If
    Case 3
        If IsMissing(pPadrao) Then
            lTamanho = 5
        Else
            If Not IsNumeric(pPadrao) Then
                lTamanho = 5
            Else
                lTamanho = pPadrao
            End If
        End If
        If gfPreenchido(pCampo) Then
            lAuxiliar = Trim$(pCampo)
            If Len(lAuxiliar) > lTamanho Then
                gfForValor = lAuxiliar
            Else
                gfForValor = Space(lTamanho - Len(lAuxiliar)) & lAuxiliar
            End If
        Else
            gfForValor = Space(lTamanho)
        End If
    End Select
End Function
Public Function gfGraItem(ByRef pBasDados As ADODB.Connection, _
                          ByVal pNosNumero As String, _
                          ByVal pNumPedido As String, _
                          ByVal pNumItem As String, _
                          ByRef pMensagem As String, _
                          pTipEmissao As e_TipEmissao, _
                          Optional ByRef pobjstcPedLoc As stcA46V708B, _
                          Optional ByVal pConCodSeguradora As Integer, _
                          Optional ByVal pConApolice As String, _
                          Optional ByVal pConItem As String, _
                          Optional ByVal pIndSinistro As Byte, _
                          Optional pCod_Ramo As Integer, _
                          Optional pCodProduto As Integer, _
                          Optional pCodPlano As Long, _
                          Optional pStcPed As stcA46V702B) As Boolean    '1000347 - Inclus�o de informa��es da cong�nere em renova��es
'Fun��o: grava registro do item - TAB_PED_LOC.
'Par�metros de entrada...pBasDados....Base de dados.
'                        pNosNumero...Nosso n�mero.
'                        pNumPediro...N�mero do pedido.
'                        pNumItem.....N�mero do item.
'Par�metro de sa�da......pMensagem....Mensagem.

    Dim methodName As String
    Dim lKeyDynaTrace As Long
    methodName = "gfGraItem"
    lKeyDynaTrace = DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    On Error GoTo TrataErro
    
    Dim lobjPedLoc As clsA46V708A    'Objeto da classe clsA46V708A (TAB_PED_LOC)
    Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)

    Dim lobjCobert As clsA46V704A    'Objeto da classe clsA46V704A (TAB_PED_COBERT)
    Dim lObjCobert_Temp As clsA46V704A
    Dim lobjstcCobert As stcA46V704B    'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)

    Dim lstcPedFator As stcA46V720B
    Dim lObjPedFator As clsA46V720A

    Dim lstcPedClau As stcA46V703B
    Dim lobjPedClau As clsA46V703A

    Dim lPropTabela As String
    Dim lNomeTabImport As String
    Dim lInd As Integer

    Dim lobjstcCobertFraq As stcA46V738B
    Dim lobjCobertFraq As clsA46V738A
    
    Set lobjstcCobertFraq = New stcA46V738B
    Set lobjstcCobertFraq.M46V999 = Me

    Set lobjCobertFraq = New clsA46V738A
    Set lobjCobertFraq.M46V999 = Me

    Dim lobjTextos As clsA46V739A
    Dim lobjstcTextos As stcA46V739B
    
    Set lobjTextos = New clsA46V739A
    Set lobjTextos.M46V999 = Me
    
    Set lobjstcTextos = New stcA46V739B
    Set lobjstcTextos.M46V999 = Me

    gfGraItem = False

    Call gfObtSYASMULTINI(lPropTabela)
    lNomeTabImport = ""


    If Not IsNumeric(pCod_Ramo) Then pCod_Ramo = 0

    '4. Instancio os objetos do Tab_Ped_Loc
    Set lobjPedLoc = New clsA46V708A
    Set lobjPedLoc.M46V999 = Me
    
    Set lobjstcPedLoc = New stcA46V708B
    Set lobjstcPedLoc.M46V999 = Me
    
    If Not pobjstcPedLoc Is Nothing Then
        Set lobjstcPedLoc.ObjPai = pobjstcPedLoc.ObjPai
        lobjstcPedLoc.NUM_PI = pobjstcPedLoc.NUM_PI
    End If

    Set lobjPedLoc.ConexaoBD = pBasDados

    '5. Verifica se tem registro padr�o do item.
    lobjstcPedLoc.nosso_numero = "PADRAO" & gNomUsuario6 & pCod_Ramo
    If Not lobjPedLoc.mfCarregarColecao(lobjstcPedLoc, lPropTabela, "", True) Then
        pMensagem = lobjPedLoc.UltimoErro
        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        GoTo SAIDA
    End If

    '6. Se n�o tem registro padr�o - ent�o inclui.
    If lobjPedLoc.Colecao.Count = 0 Then
        Call mfVerRegPadrao(pMensagem)
    Else
        Set lobjstcPedLoc = lobjPedLoc.Colecao.Item(1)
    End If

    '2. Carrego as informa��es
    With lobjstcPedLoc
        .nosso_numero = pNosNumero
        .Num_Ped = pNumPedido
        .Num_Item = pNumItem
        .Cod_Ramo = pCod_Ramo
        .Tip_Emissao = pTipEmissao
        '1000347 - Inclus�o de informa��es da cong�nere em renova��es
        If pTipEmissao = e_TipEmissao_RenCongenere Then
            .Ren_Con_Seguradora = pConCodSeguradora
            .Ren_Con_Apolice = pConApolice
            .Ren_Con_Item = pConItem
            .Ind_Sinistro = pIndSinistro
        End If
        If pTipEmissao = e_TipEmissao_Endosso Then
            .Cod_DizerItem = e_CodDizerInclusao
            .Num_Apol_Anterior = pConApolice
        End If
        If pCodPlano <> 0 Then
            .Cod_Plano = pCodPlano
            .Cod_Produto = pCodPlano
        Else
            .Cod_Plano = gfPlanoBasico(.Cod_Ramo)
            .Cod_Produto = gfPlanoBasico(.Cod_Ramo)
        End If
        If Not pStcPed Is Nothing Then
            .NUM_PI = pStcPed.NUM_PI
        End If
        ''.ITEM_ORIGEM_C4 = 0
        If App.EXEName = "P46V101" Then
            .ITEM_ORIGEM_C4 = 2
        End If
    End With

    '3. Chamo a rotina de incluir
    If Not lobjPedLoc.mfIncluir(lobjstcPedLoc, lPropTabela, lNomeTabImport, tagDynaTrace) Then
        pMensagem = lobjPedLoc.UltimoErro
        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        GoTo SAIDA
    End If


    'Inspe��o
    Dim lobjInsp As clsA46V717A
    Set lobjInsp = New clsA46V717A
    Set lobjInsp.M46V999 = Me

    Set lobjInsp.ConexaoBD = pBasDados
    lobjstcPedLoc.objstcPedInsp.nosso_numero = pNosNumero
    lobjstcPedLoc.objstcPedInsp.Num_Ped = pNumPedido
    lobjstcPedLoc.objstcPedInsp.Num_Item = pNumItem
    lobjstcPedLoc.objstcPedInsp.Cod_Ramo = pCod_Ramo
    lobjstcPedLoc.objstcPedInsp.Cod_Org = lobjstcPedLoc.Cod_Org
    lobjstcPedLoc.objstcPedInsp.NUM_PI = lobjstcPedLoc.NUM_PI
    If Not lobjInsp.mfIncluir(lobjstcPedLoc.objstcPedInsp, lPropTabela, lNomeTabImport, tagDynaTrace) Then
        pMensagem = lobjInsp.UltimoErro
        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        GoTo SAIDA
    End If

    '5. Verifica se tem registro padr�o
    Set lobjCobert = New clsA46V704A
    Set lobjCobert.M46V999 = Me

    Set lobjstcCobert = New stcA46V704B
    Set lobjstcCobert.M46V999 = Me

    Set lobjCobert.ConexaoBD = pBasDados

    '6. Se n�o tem registro padr�o - ent�o inclui.
    If lobjstcCobert.ColecaoNovaFranquia.Count = 0 Then
        Call mfVerRegPadrao(pMensagem)
    End If

    '6. Se n�o tem registro padr�o - ent�o inclui.
    If lobjstcPedLoc.ColecaoPedCobert.Count = 0 Then
        Set lobjstcCobert = New stcA46V704B
        Set lobjstcCobert.M46V999 = Me

        With lobjstcCobert
            .nosso_numero = pNosNumero
            .Num_Ped = pNumPedido
            .Num_Item = pNumItem
            .Cod_Ramo = pCod_Ramo
            .Cod_Org = lobjstcPedLoc.Cod_Org
            .Cod_Cobert = 1
            If pCodPlano <> 0 Then
                .Tab_Cobert = pCodPlano
            Else
                .Tab_Cobert = 1
            End If
            Set .ObjPai = lobjstcPedLoc
        End With
        If Not lobjCobert.mfIncluir(lobjstcCobert, lPropTabela, lNomeTabImport, False, False, tagDynaTrace) Then
            pMensagem = lobjCobert.UltimoErro
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        Set lobjstcPedLoc.ColecaoPedCobert = New Collection
        If Not lobjCobert.Colecao Is Nothing Then  'inclu�do juntamente com Thais - 21/10/2011
            Set lobjstcPedLoc.ColecaoPedCobert = lobjCobert.Colecao
        End If
    Else
        For Each lobjstcCobert In lobjstcPedLoc.ColecaoPedCobert
            With lobjstcCobert
                .nosso_numero = pNosNumero
                .Num_Ped = pNumPedido
                .Num_Item = pNumItem
                .Cod_Ramo = pCod_Ramo
                .Cod_Org = lobjstcPedLoc.Cod_Org
                Set .ObjPai = lobjstcPedLoc
            End With
            If Not lobjCobert.mfIncluir(lobjstcCobert, lPropTabela, lNomeTabImport, False, False, tagDynaTrace) Then
                pMensagem = lobjCobert.UltimoErro
                Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If
            
            'Multiplo Franquia
            For Each lobjstcCobertFraq In lobjstcCobert.ColecaoNovaFranquia
                With lobjstcCobertFraq
                    .nosso_numero = pNosNumero
                    .Num_Ped = pNumPedido
                    .Num_Item = pNumItem
                    .Cod_Ramo = pCod_Ramo
                    .Cod_Tarif = pCodProduto
                    Set .ObjPai = lobjstcCobert
                End With
                If Not lobjCobertFraq.mfIncluir(lobjstcCobertFraq, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                    pMensagem = lobjCobert.UltimoErro
                    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If
            Next lobjstcCobertFraq
        Next lobjstcCobert
    End If

    'Fatores de Prote��o
    Set lObjPedFator = New clsA46V720A
    Set lObjPedFator.M46V999 = Me

    Set lObjPedFator.ConexaoBD = pBasDados
    For Each lstcPedFator In lobjstcPedLoc.ColecaoPedFator
        lstcPedFator.nosso_numero = pNosNumero
        lstcPedFator.Num_Ped = pNumPedido
        lstcPedFator.Num_Item = pNumItem
        lstcPedFator.Cod_Ramo = pCod_Ramo
        lstcPedFator.Cod_Org = lobjstcPedLoc.Cod_Org
        Set lstcPedFator.ObjPai = lobjstcPedLoc
        If Not lObjPedFator.mfIncluir(lstcPedFator, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lObjPedFator.UltimoErro
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    Next

    Set lobjstcPedLoc.ObjPai = pStcPed

    '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 8
    If (lobjstcPedLoc.Cod_Ramo = e_CodRamo_Residencial And lobjstcPedLoc.Cod_Plano = 1) Then
        'If lObjstcPedLoc.gTemProtecaoObrigatoria = True Then  '1300102 - 1201062 - Tarifa 1048 e 1049 item 1
        If lobjstcPedLoc.gLiberaDescontoEspecial = True Then
            Set lstcPedFator = New stcA46V720B
            With lstcPedFator
                .nosso_numero = pNosNumero
                .Num_Ped = pNumPedido
                .Num_Item = pNumItem
                .Cod_Ramo = pCod_Ramo
                .Cod_Org = lobjstcPedLoc.Cod_Org
                .Cod_Fator = 7
                .Desc_Fator = "DESCONTO ESPECIAL"
                Set .ObjPai = lobjstcPedLoc
            End With
            If Not lObjPedFator.mfIncluir(lstcPedFator, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lObjPedFator.UltimoErro
                Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If
        End If
    End If

    'Ajuste Pontual Fator de Prote��o  Simony
    Dim FatoresObrigatorios As String
    Dim RetFatores() As String    'Recebe itens da variavel lInfoConexao
    'Dim e_Tip_Valor_FatorProtecao As String
    'e_Tip_Valor_FatorProtecao = 22
    If M46V103.mfAjustesPontuais(e_Tip_Valor_FatorProtecao, TpoValor, FatoresObrigatorios, pMensagem, lobjstcPedLoc, 0, tagDynaTrace) = False Then
        RetFatores = Split(FatoresObrigatorios, ";")
        For lInd = 0 To UBound(RetFatores())
            If RetFatores(lInd) <> "" Then
                Set lstcPedFator = New stcA46V720B
                With lstcPedFator
                    .nosso_numero = pNosNumero
                    .Num_Ped = pNumPedido
                    .Num_Item = pNumItem
                    .Cod_Ramo = pCod_Ramo
                    .Cod_Org = lobjstcPedLoc.Cod_Org
                    .Cod_Fator = Mid(RetFatores(lInd), 1, InStr(RetFatores(lInd), "@") - 1)
                    .Desc_Fator = Mid(RetFatores(lInd), InStr(RetFatores(lInd), "@") + 1)
                    Set .ObjPai = lobjstcPedLoc
                End With
                If Not lObjPedFator.mfIncluir(lstcPedFator, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                    pMensagem = lObjPedFator.UltimoErro
                    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If
            End If
        Next
    End If


    Dim lstcPedLocBem As stcA46V715B
    Dim lobjPedLocBem As clsA46V715A

    Set lobjPedLocBem = New clsA46V715A
    Set lobjPedLocBem.M46V999 = Me

    Set lobjPedLocBem.ConexaoBD = pBasDados
    For Each lstcPedLocBem In lobjstcPedLoc.ColecaoPedLocBem
        lstcPedLocBem.nosso_numero = pNosNumero
        lstcPedLocBem.Num_Ped = pNumPedido
        lstcPedLocBem.Num_Item = pNumItem
        lstcPedLocBem.Cod_Ramo = pCod_Ramo
        lstcPedLocBem.Cod_Org = lobjstcPedLoc.Cod_Org
        Set lstcPedLocBem.ObjPai = lobjstcPedLoc
        If Not lobjPedLocBem.mfIncluir(lstcPedLocBem, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedLocBem.UltimoErro
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    Next

    If pTipEmissao = e_TipEmissao_Endosso Then
        'Verifico se a apolice anterior tem
        Set lstcPedClau = New stcA46V703B
        Set lobjPedClau = New clsA46V703A
        Set lobjPedClau.M46V999 = Me

        lstcPedClau.Cod_Clau = 400
        lstcPedClau.Cod_Ramo = pCod_Ramo
        lstcPedClau.Num_Apol = pConApolice
        Set lstcPedClau.ObjPai = lobjstcPedLoc
        If Not lobjPedClau.mfApoliceComNovaClausulaAss24hr(lstcPedClau, "GED.DBO.", lNomeTabImport) Then
            GoTo PularClausula
        End If
        Set lstcPedClau = New stcA46V703B
        Set lobjPedClau = New clsA46V703A
        Set lobjPedClau.M46V999 = Me
        lstcPedClau.nosso_numero = pNosNumero
        lstcPedClau.Num_Ped = pNumPedido
        lstcPedClau.Cod_Clau = 400
        lstcPedClau.Num_Item = pNumItem
        lstcPedClau.Cod_Ramo = pCod_Ramo
        lstcPedClau.Cod_Org = lobjstcPedLoc.Cod_Org
        Set lstcPedClau.ObjPai = lobjstcPedLoc
        If Not lobjPedClau.mfIncluir(lstcPedClau, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedClau.UltimoErro
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If
PularClausula:



    '--Beneficiarios
    Dim lstcPedBenefiario As stcA46V735B
    Dim lobjPedBenefiario As clsA46V735A

    Set lobjPedBenefiario = New clsA46V735A
    Set lobjPedBenefiario.M46V999 = Me

    Set lobjPedBenefiario.ConexaoBD = pBasDados
    For Each lstcPedBenefiario In lobjstcPedLoc.ColecaoBeneficiario
        lstcPedBenefiario.nosso_numero = pNosNumero
        lstcPedBenefiario.Num_Item = pNumItem
        lstcPedBenefiario.NUM_PI = lobjstcPedLoc.NUM_PI
        lstcPedBenefiario.Tip_Loc_Pessoa = e_TipLocPessoa_Beneficiario
        Set lstcPedBenefiario.ObjPai = lobjstcPedLoc
        If Not lobjPedBenefiario.mfIncluir(lstcPedBenefiario, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedBenefiario.UltimoErro
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    Next



    '--Inquilino
    If gfPreenchido(lobjstcPedLoc.objstcPedInquilino.Num_CPF_CNPJ) And Trim(lobjstcPedLoc.objstcPedInquilino.Num_CPF_CNPJ) <> "0" Then
        Dim lobjstcInquilino As stcA46V735B
        Dim lobjPedInquilino As clsA46V735A

        Set lobjPedInquilino = New clsA46V735A
        Set lobjPedInquilino.M46V999 = Me

        Set lobjPedInquilino.ConexaoBD = pBasDados
        Set lobjstcInquilino = lobjstcPedLoc.objstcPedInquilino
        lobjstcInquilino.nosso_numero = pNosNumero
        lobjstcInquilino.Num_Item = pNumItem
        lobjstcInquilino.NUM_PI = lobjstcPedLoc.NUM_PI
        lobjstcInquilino.Tip_Loc_Pessoa = e_TipLocPessoa_Inquilino
        Set lobjstcInquilino.ObjPai = lobjstcPedLoc
        If Not lobjPedInquilino.mfIncluir(lobjstcInquilino, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedInquilino.UltimoErro
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If


    '--Proprietario
    If gfPreenchido(lobjstcPedLoc.objstcPedProprietario.Num_CPF_CNPJ) And Trim(lobjstcPedLoc.objstcPedProprietario.Num_CPF_CNPJ) <> "0" Then
        Dim lobjstcProprietario As stcA46V735B
        Dim lobjPedProprietario As clsA46V735A

        Set lobjPedProprietario = New clsA46V735A
        Set lobjPedProprietario.M46V999 = Me

        Set lobjPedProprietario.ConexaoBD = pBasDados
        Set lobjstcProprietario = lobjstcPedLoc.objstcPedProprietario
        lobjstcProprietario.nosso_numero = pNosNumero
        lobjstcProprietario.Num_Item = pNumItem
        lobjstcProprietario.NUM_PI = lobjstcPedLoc.NUM_PI
        lobjstcProprietario.Tip_Loc_Pessoa = e_TipLocPessoa_Proprietario
        Set lobjstcProprietario.ObjPai = lobjstcPedLoc
        If Not lobjPedProprietario.mfIncluir(lobjstcProprietario, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedProprietario.UltimoErro
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If
    
    For Each lobjstcTextos In lobjstcPedLoc.ColecaoTextos
        lobjstcTextos.nosso_numero = pNosNumero
        lobjstcTextos.Num_Ped = pNumPedido
        lobjstcTextos.Num_Item = pNumItem
        lobjstcTextos.Cod_Ramo = pCod_Ramo
        lobjstcTextos.Cod_Org = lobjstcPedLoc.Cod_Org
        Set lobjstcTextos.ObjPai = lobjstcPedLoc
        If Not lobjTextos.mfIncluir(lobjstcTextos, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjTextos.UltimoErro
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    Next
    
    Set pobjstcPedLoc = lobjstcPedLoc

    gfGraItem = True

    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    GoSub FechaObj
    Exit Function

FechaObj:
    '4. Destruo os objetos
    Set lobjPedLoc = Nothing
    Set lobjstcPedLoc = Nothing
    Set lobjstcCobert = Nothing
    Set lobjCobert = Nothing
    Set lstcPedFator = Nothing
    Set lObjPedFator = Nothing
    Set lobjInsp = Nothing
    Set lobjTextos = Nothing
    Set lobjstcTextos = Nothing
    Set lObjCobert_Temp = Nothing
    Set lstcPedClau = Nothing
    Set lobjPedClau = Nothing
    Set lobjstcCobertFraq = Nothing
    Set lobjCobertFraq = Nothing
    Set lobjTextos = Nothing
    Set lobjstcTextos = Nothing
    
    Return
    
TrataErro:
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
    gfGraItem = False
    GoTo SAIDA
    
End Function
Public Function gfNomCorretor(ByVal pCodCorretor As String, ByRef pNomCorretor As String, _
                              ByRef pMensagem As String) As Boolean
'Fun��o: obt�m nome do corretor.

'Par�metro de entrada...pCodCorretor...C�digo do corretor.
'Par�metros de sa�da....pNomCorretor...Nome do corretor.
'                       pMensagem......Mensagem.

    Dim bdP0044700 As ADODB.Connection    'Arquivo - P0044700.
    Dim lSelect As String       'Comando SELECT.
    Dim rsCORR As clsYasRecordSet    'YTAB_CORR: tabela.

    gfNomCorretor = False

    '1. Verifica preenchimento.
    If Not gfPreenchido(pCodCorretor) Then
        pNomCorretor = "C�digo do corretor n�o preenchido."
        gfNomCorretor = True
        Exit Function
    End If
    If val(pCodCorretor) = 0 Then
        pNomCorretor = "C�digo do corretor zerado."
        gfNomCorretor = True
        Exit Function
    End If

    '2. Obt�m nome do corretor.
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    lSelect = "SELECT NOM_ABRV_CORR FROM YTAB_CORR WHERE COD_CORR = " & pCodCorretor
    If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
        Call Me.gpFecha2(bdP0044700)
        Exit Function
    End If
    If rsCORR.EOF = True Then
        Call Me.gpFecha2(bdP0044700)
        pMensagem = "FGE0003 - " & gPrefixo & "Corretor inv�lido (c�digo de corretor n�o localizado " & _
                    "- " & pCodCorretor & ")."
        Exit Function
    End If
    pNomCorretor = Trim$(rsCORR!NOM_ABRV_CORR)
    Call Me.gpFecha2(bdP0044700)

    gfNomCorretor = True
End Function
Public Function gfNomSeguradora(ByRef pCodSeguradora As String, ByRef pNomSeguradora As String, _
                                ByRef pMensagem As String, Optional ByRef pNomAbrev As String) As Boolean
'Fun��o: obt�m nome da seguradora

'Par�metro de entrada...pCodSeguradora...C�digo da seguradora.
'Par�metros de sa�da....pNomSeguradora...Nome da seguradora.
'                       pMensagem........Mensagem.

    Dim bdP0044700 As ADODB.Connection    'Arquivo - P0044700.
    Dim lSelect As String       'Comando SELECT.
    Dim rsP0044700 As clsYasRecordSet    'P0044700: tabela.

    gfNomSeguradora = False

    '1. Verifica preenchimento.
    If Not gfPreenchido(pCodSeguradora) Then
        pMensagem = "FTE0066 - " & gPrefixo & _
                    "Erro ao obter nome da seguradora (c�digo da seguradora n�o preenchido)."
        Exit Function
    End If
    If Not IsNumeric(pCodSeguradora) Then
        pMensagem = "FTE0066 - " & gPrefixo & _
                    "Erro ao obter nome da seguradora (c�digo da seguradora n�o num�rico)."
        Exit Function
    End If
    If val(pCodSeguradora) = 0 Then
        pMensagem = "FTE0066 - " & gPrefixo & _
                    "Erro ao obter nome da seguradora (c�digo da seguradora zerado)."
        Exit Function
    End If

    '2. Obt�m nome da seguradora.
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    lSelect = "SELECT Seguradora , aSeguradora FROM YTAB_SEGURADORA WHERE CSeguradora = '" & Trim(pCodSeguradora) & "'"
    If gfObtRegistro(bdP0044700, lSelect, rsP0044700, pMensagem) = False Then
        Call Me.gpFecha2(bdP0044700)
        Exit Function
    End If
    If rsP0044700.EOF = True Then
        pMensagem = "FTE0066 - " & gPrefixo & _
                    "Erro ao obter nome da seguradora (c�digo da seguradora n�o localizado - " & _
                    pCodSeguradora & ")."
        Exit Function
    End If
    pNomSeguradora = rsP0044700!Seguradora
    pNomAbrev = rsP0044700!aSeguradora
    Call Me.gpFecha2(bdP0044700)

    gfNomSeguradora = True
End Function
Public Function gfVerCorrPermitido(ByVal pCorrInstalacao As Long, ByVal pCorrSelecionado As Long, pRegistro As String) As Boolean
    Dim lCorrInstacao As Long    'C�digo do corretor - instala��o.
    Dim lCorrSelecionado As Long  'C�digo do corretor - selecionado.
    Dim lEncArquivo As Boolean    'Identifica se a verifica��o foi feita atrav�s do arquivo, caso
'contr�rio dever� ser buscado do c�digo
    Dim lEncontrou As Boolean    'Identifica se foi encontrato o [grupos] no arquivo.
    Dim lEndP366CORR As String    'Endere�o do arquivo P366CORR.EXE.
    Dim lNumP366CORR As Integer    'Utilizado com FreeFile para o arquivo P366CORR.EXE.
    Dim lRegistro As String    'Registro.

'A compara��o no arquivo � feita com o n�mero do corretor sem os dois �ltimos d�gitos.
'Os dois �ltimos d�gitos variam mas dentro de um mesmo grupo de corretor.
'No arquivo P366CORR.EXE n�o tem o n�mero completo do corretor, est� sem os dois �ltimos d�gitos.
    lCorrInstacao = Left(Format(pCorrInstalacao, "0000000"), 5)
    lCorrSelecionado = Left(Format(pCorrSelecionado, "0000000"), 5)
    lEndP366CORR = gAppPath & "P366CORR.EXE"
    If Not gfPreenchido(Dir(lEndP366CORR)) Then
        lEncArquivo = False
    Else
        lEncArquivo = True
    End If
    If lEncArquivo = True Then
        lNumP366CORR = FreeFile
        On Error Resume Next
        Open lEndP366CORR For Input Access Read As lNumP366CORR
        If Err <> 0 Then
            lEncArquivo = False
        End If
        If lEncArquivo = True Then
            On Error GoTo 0
            '4. Trata registros do arquivo de inicializa��o.
            lEncontrou = False
            lEncArquivo = False
            Do While Not EOF(lNumP366CORR)
                Line Input #lNumP366CORR, lRegistro
                If gfPreenchido(lRegistro) Then
                    lRegistro = Trim$(lRegistro)
                    If UCase(lRegistro) = "[GRUPOS]" Then
                        lEncontrou = True
                    End If
                    If UCase(lRegistro) = "[SOMENTE MULTI]" Then
                        Exit Do
                    End If
                    If lEncontrou = True Then
                        If InStr(1, lRegistro, ";" & lCorrInstacao & ";") > 0 Then
                            pRegistro = lRegistro
                            If InStr(1, lRegistro, ";" & lCorrSelecionado & ";") > 0 Then
                                gfVerCorrPermitido = True
                            Else
                                gfVerCorrPermitido = False
                            End If
                            lEncArquivo = True
                            Exit Do
                        End If
                        If lCorrInstacao <> lCorrSelecionado Then
                            lEncArquivo = False
                            gfVerCorrPermitido = False
                        Else
                            lEncArquivo = True
                            gfVerCorrPermitido = True
                            Exit Do
                        End If
                    End If
                End If
            Loop
            Close lNumP366CORR
        End If
    End If
    If lEncArquivo = False Then
        Select Case lCorrInstacao
        Case 809, 2387
            Select Case lCorrSelecionado
            Case 809, 2387
                'Sele��o OK
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 376, 3813
            Select Case lCorrSelecionado
            Case 376, 3813
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2234, 904
            Select Case lCorrSelecionado
            Case 2234, 904
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 183, 4148
            Select Case lCorrSelecionado
            Case 183, 4148
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 77, 4246, 4708, 4709, 4710
            Select Case lCorrSelecionado
            Case 77, 4246, 4708, 4709, 4710
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 4618, 1102
            Select Case lCorrSelecionado
            Case 4618, 1102
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 843, 1172, 845, 377
            Select Case lCorrSelecionado
            Case 843, 1172, 845, 377
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 6449, 1749
            Select Case lCorrSelecionado
            Case 6449, 1749
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 1889, 495
            Select Case lCorrSelecionado
            Case 1889, 495
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 6106, 4978, 4800, 4792, 4846, 2769, 3223
            Select Case lCorrSelecionado
            Case 6106, 4978, 4800, 4792, 4846, 2769, 3223
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3272, 34, 3335, 393, 3070, 73
            Select Case lCorrSelecionado
            Case 3272, 34, 3335, 393, 3070, 73
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2577, 2101
            Select Case lCorrSelecionado
            Case 2577, 2101
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2422, 6145
            Select Case lCorrSelecionado
            Case 2422, 6145
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 172, 3407
            Select Case lCorrSelecionado
            Case 172, 3407
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 855, 1170, 1171, 2582, 6659
            Select Case lCorrSelecionado
            Case 855, 1170, 1171, 2582, 6659
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 60, 3301, 3302
            Select Case lCorrSelecionado
            Case 60, 3301, 3302
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 1184, 709
            Select Case lCorrSelecionado
            Case 1184, 709
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 389, 4252
            Select Case lCorrSelecionado
            Case 389, 4252
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2505, 4162
            Select Case lCorrSelecionado
            Case 2505, 4162
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 96, 534, 1454
            Select Case lCorrSelecionado
            Case 96, 534, 1454
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 822, 406
            Select Case lCorrSelecionado
            Case 822, 406
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 530, 2171
            Select Case lCorrSelecionado
            Case 530, 2171
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 6369, 2994, 6544
            Select Case lCorrSelecionado
            Case 6369, 2994, 6544
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3668, 1496
            Select Case lCorrSelecionado
            Case 3668, 1496
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 4675, 6446
            Select Case lCorrSelecionado
            Case 4675, 6446
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 1491, 3821
            Select Case lCorrSelecionado
            Case 1491, 3821
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3272, 34, 4726, 4431, 3438, 3302, 73, 3456, 4723, 3459, 3070, 3392
            Select Case lCorrSelecionado
            Case 3272, 34, 4726, 4431, 3438, 3302, 73, 3456, 4723, 3459, 3070, 3392
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2360, 2888
            Select Case lCorrSelecionado
            Case 2360, 2888
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 4565, 4669
            Select Case lCorrSelecionado
            Case 4565, 4669
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 748, 1069
            Select Case lCorrSelecionado
            Case 748, 1069
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 1155, 2501, 1255, 6205
            Select Case lCorrSelecionado
            Case 1155, 2501, 1255, 6205
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3494, 3189, 3583, 3509, 356, 2956
            Select Case lCorrSelecionado
            Case 3494, 3189, 3583, 3509, 356, 2956
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3075, 3367
            Select Case lCorrSelecionado
            Case 3075, 3367
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 1863, 6238, 4621, 3311, 6237, 6244, 6242, 6248, 6231, 1860, 6191, 6243, 4622, 6658
            Select Case lCorrSelecionado
            Case 1863, 6238, 4621, 3311, 6237, 6244, 6242, 6248, 6231, 1860, 6191, 6243, 4622, 6658
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 322, 6361, 263, 2311, 3865, 4866, 6439
            Select Case lCorrSelecionado
            Case 322, 6361, 263, 2311, 3865, 4866, 6439
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 4267, 6678
            Select Case lCorrSelecionado
            Case 4267, 6678
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 214, 6125
            Select Case lCorrSelecionado
            Case 214, 6125
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2841, 2842
            Select Case lCorrSelecionado
            Case 2841, 2842
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 6593, 1465
            Select Case lCorrSelecionado
            Case 6593, 1465
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2453, 4485
            Select Case lCorrSelecionado
            Case 2453, 4485
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 4503, 4353, 3416
            Select Case lCorrSelecionado
            Case 4503, 4353, 3416
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3111, 1614
            Select Case lCorrSelecionado
            Case 3111, 1614
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2355, 2344
            Select Case lCorrSelecionado
            Case 2355, 2344
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 1116, 1901, 1321, 1890
            Select Case lCorrSelecionado
            Case 1116, 1901, 1321, 1890
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2451, 1148
            Select Case lCorrSelecionado
            Case 2451, 1148
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3043, 6515
            Select Case lCorrSelecionado
            Case 3043, 6515
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2142, 133
            Select Case lCorrSelecionado
            Case 2142, 133
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3775, 3098, 1807, 1769, 3082, 6621, 768, 4213, 3081
            Select Case lCorrSelecionado
            Case 3775, 3098, 1807, 1769, 3082, 6621, 768, 4213, 3081
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 6435, 1688
            Select Case lCorrSelecionado
            Case 6435, 1688
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3379, 6625
            Select Case lCorrSelecionado
            Case 3379, 6625
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2145, 871
            Select Case lCorrSelecionado
            Case 2145, 871
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2369, 162, 469, 483
            Select Case lCorrSelecionado
            Case 2369, 162, 469, 483
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2912, 2911
            Select Case lCorrSelecionado
            Case 2912, 2911
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2479, 2889, 4352
            Select Case lCorrSelecionado
            Case 2479, 2889, 4352
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3610, 6592, 4561, 3642, 3271, 4245, 3742, 1640, 1923, 2859, 3667
            Select Case lCorrSelecionado
            Case 3610, 6592, 4561, 3642, 3271, 4245, 3742, 1640, 1923, 2859, 3667
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3544, 3543, 3547, 3541, 3542, 3545, 3546, 3548, 2740
            Select Case lCorrSelecionado
            Case 3544, 3543, 3547, 3541, 3542, 3545, 3546, 3548, 2740
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2734, 4886
            Select Case lCorrSelecionado
            Case 2734, 4886
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3457, 3458, 3459, 3456
            Select Case lCorrSelecionado
            Case 3457, 3458, 3459, 3456
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 4573, 6008
            Select Case lCorrSelecionado
            Case 4573, 6008
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 1060, 2702
            Select Case lCorrSelecionado
            Case 1060, 2702
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 58, 6487, 4037
            Select Case lCorrSelecionado
            Case 58, 6487, 4037
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3982, 3844, 3119, 6080, 3149, 6378
            Select Case lCorrSelecionado
            Case 3982, 3844, 3119, 6080, 3149, 6378
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2909, 2182
            Select Case lCorrSelecionado
            Case 2909, 2182
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 62, 6616
            Select Case lCorrSelecionado
            Case 62, 6616
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 2621, 179
            Select Case lCorrSelecionado
            Case 2621, 179
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3438, 3392
            Select Case lCorrSelecionado
            Case 3438, 3392
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3319, 6131
            Select Case lCorrSelecionado
            Case 3319, 6131
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 827, 337
            Select Case lCorrSelecionado
            Case 827, 337
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 4481, 2147, 4589
            Select Case lCorrSelecionado
            Case 4481, 2147, 4589
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 80, 4086
            Select Case lCorrSelecionado
            Case 80, 4086
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 6240, 6198, 6195, 3397, 6210, 6256, 6298, 6219, 6209, 6442, 6420, 6172, 6297, 6301, 6193, 6226, 6413, 6218, 1895, 6228, 1864, 6419, 6588, 6208, 6224, 1435, 6261, 6175, 6596, 6184
            Select Case lCorrSelecionado
            Case 6240, 6198, 6195, 3397, 6210, 6256, 6298, 6219, 6209, 6442, 6420, 6172, 6297, 6301, 6193, 6226, 6413, 6218, 1895, 6228, 1864, 6419, 6588, 6208, 6224, 1435, 6261, 6175, 6596, 6184
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 6128, 4700, 6217, 1771, 4661, 1872, 6372
            Select Case lCorrSelecionado
            Case 6128, 4700, 6217, 1771, 4661, 1872, 6372
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 1825, 4118, 4389, 6611, 6404, 4724, 4391, 1705, 4775, 6293, 4393, 4117, 4619, 4394, _
             6606, 4463, 6294, 4556, 6580, 4396, 4576, 4916, 4398, 4457, 4482, 6603, 4458, 4399, _
             6291, 1893, 4735, 6609, 4693, 1707, 6394, 6636, 1871, 1761, 6009, 4538, 6655, 6098, _
             6406, 6351, 1828, 4405, 6255, 4640, 4407, 4408, 4702, 6352, 6643, 4483, 4410, 1823, _
             1715, 6440, 1827, 4413, 4429, 4414, 4416, 1791, 6232, 6299, 6485, 4587, 6608, 4419, _
             6626, 4448, 6418, 6019, 6182, 4518, 4690, 4633, 4818, 6190, 6513, 1911
            Select Case lCorrSelecionado
            Case 1825, 4118, 4389, 6611, 6404, 4724, 4391, 1705, 4775, 6293, 4393, 4117, _
                 4619, 4394, 6606, 4463, 6294, 4556, 6580, 4396, 4576, 4916, 4398, 4457, _
                 4482, 6603, 4458, 4399, 6291, 1893, 4735, 6609, 4693, 1707, 6394, 6636, _
                 1871, 1761, 6009, 4538, 6655, 6098, 6406, 6351, 1828, 4405, 6255, 4640, _
                 4407, 4408, 4702, 6352, 6643, 4483, 4410, 1823, 1715, 6440, 1827, 4413, _
                 4429, 4414, 4416, 1791, 6232, 6299, 6485, 4587, 6608, 4419, 6626, 4448, _
                 6418, 6019, 6182, 4518, 4690, 4633, 4818, 6190, 6513, 1911
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 4966, 4875, 4900, 4830, 4898
            Select Case lCorrSelecionado
            Case 4966, 4875, 4900, 4830, 4898
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case 3611, 3708, 1448, 1468, 1464, 1460, 3485, 3486, 2747, 3100, 3612, 3484, 3511, 3482, _
             3483, 3585, 1449, 3502, 3892, 3531, 1474, 1484, 3569, 1451, 3117, 1485, 3481, 3636, _
             3581, 3638, 3480, 3568, 3732, 3637, 3487, 2585, 3627, 3385, 3493, 3665, 3273, 3654, _
             3499
            Select Case lCorrSelecionado
            Case 3611, 3708, 1448, 1468, 1464, 1460, 3485, 3486, 2747, 3100, 3612, 3484, _
                 3511, 3482, 3483, 3585, 1449, 3502, 3892, 3531, 1474, 1484, 3569, 1451, _
                 3117, 1485, 3481, 3636, 3581, 3638, 3480, 3568, 3732, 3637, 3487, 2585, _
                 3627, 3385, 3493, 3665, 3273, 3654, 3499
                gfVerCorrPermitido = True
            Case Else
                'Erro
                gfVerCorrPermitido = False
            End Select
        Case Else
            If lCorrInstacao <> lCorrSelecionado Then
                gfVerCorrPermitido = False
            Else
                gfVerCorrPermitido = True
            End If
        End Select
    End If
End Function

Public Function gfObtCorretor(ByVal pOpcao As Byte, ByVal pCorretor As String, _
                              ByRef pSaida As Variant, ByRef pMensagem As String) As Boolean
'Fun��o: obt�m dados de corretores.

'Par�metros de entrada...pOpcao......1 = C�digo do corretor e 2 = Nome do corretor.
'                        pCorretor...C�digo ou nome do corretor, de acordo com pOpcao.
'Par�metro de sa�da......pSaida......Array contendo resultados.
'                                    (x, 1) C�digo do corretor.
'                                    (x, 2) Nome do corretor.
'                                    (x, 3) Nome abreviado do corretor.
'                                    (x, 4) Num. Susep
    Dim bdP0044700 As ADODB.Connection    'Arquivo - P0044700.
    Dim lI As Double  'Contador.
    Dim lSelect As String       'Comando SELECT.
    Dim rsCORR As clsYasRecordSet    'Tabela: YTAB_CORR.

    gfObtCorretor = False

    '1. Verifica par�metros.
    'a) Op��o.
    If Not gfPreenchido(pOpcao) Then
        Call gpGraLog(2, "FTE0067 - Op��o inv�lida (obter corretor - op��o n�o preenchida)." & vbLf & _
                         vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If
    If Not IsNumeric(pOpcao) Then
        Call gpGraLog(2, "FTE0067 - Op��o inv�lida (obter corretor - op��o n�o num�rica)." & vbLf & _
                         vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If
        'Tarifa 11.72 - Cocorretagem
    If pOpcao <> 1 And pOpcao <> 2 And pOpcao <> 3 Then
        Call gpGraLog(2, "FTE0067 - Op��o inv�lida (obter corretor - op��o diferente de 1 e 2)." & _
                         vbLf & vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If
    'b) Corretor.
    If Not gfPreenchido(pCorretor) Then
        Call gpGraLog(2, "FTE0068 - Erro ao obter corretor (c�digo do corretor n�o preenchido)." & _
                         vbLf & vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If

    '2. Busca informa��o.
    Select Case pOpcao
    Case 1      'C�digo do corretor.
        If Not IsNumeric(pCorretor) Then
            Call gpGraLog(2, "FTE0068 - Erro ao obter corretor (c�digo do corretor n�o num�rico)." & _
                             vbLf & vbLf & "O sistema ser� encerrado.")
            Exit Function    'End   ***MOSILVA Verificar
        End If
        lSelect = "SELECT COD_CORR, NOM_CORR, NOM_ABRV_CORR , num_susep FROM YTAB_CORR WHERE COD_CORR = " & _
                  pCorretor & " AND COD_ATIVO = 1"
        If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
            Exit Function
        End If
        If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
            Exit Function
        End If
        If rsCORR.RecordCount = 0 Then
            Call Me.gpFecha2(bdP0044700)
            pMensagem = "FGE0003 - " & gPrefixo & "Corretor inv�lido (c�digo de corretor n�o " & _
                        "localizado - " & pCorretor & ")."
            Exit Function
        End If
        ReDim pSaida(1, 4)
        pSaida(1, 1) = rsCORR!Cod_Corr
        pSaida(1, 2) = rsCORR!NOM_CORR
        pSaida(1, 3) = rsCORR!NOM_ABRV_CORR
        pSaida(1, 4) = rsCORR!NUM_SUSEP
        Call Me.gpFecha2(bdP0044700)
        gfObtCorretor = True
        Exit Function
    Case 2      'Nome do corretor.
        lSelect = "SELECT COD_CORR, NOM_CORR, NOM_ABRV_CORR , Num_Susep FROM YTAB_CORR " & _
                  "WHERE (NOM_CORR LIKE '%" & pCorretor & "%' " & _
                  "OR     NOM_ABRV_CORR LIKE '%" & pCorretor & "%') " & _
                  "AND   COD_ATIVO = 1"
        If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
            Exit Function
        End If
        If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
            Exit Function
        End If
        If rsCORR.RecordCount = 0 Then
            Call Me.gpFecha2(bdP0044700)
            pMensagem = "FGE0003 - Corretor inv�lido (corretor n�o localizado com nome ou parte " & _
                        "do nome = " & pCorretor & ")."
            Exit Function
        End If
        rsCORR.MoveLast
        ReDim Preserve pSaida(rsCORR.RecordCount, 4)
        lI = 0
        rsCORR.MoveFirst
        While Not rsCORR.EOF
            lI = lI + 1
            pSaida(lI, 1) = rsCORR!Cod_Corr
            pSaida(lI, 2) = rsCORR!NOM_CORR
            pSaida(lI, 3) = rsCORR!NOM_ABRV_CORR
            pSaida(lI, 4) = rsCORR!NUM_SUSEP
            rsCORR.MoveNext
        Wend
        Call Me.gpFecha2(bdP0044700)
        gfObtCorretor = True
        Exit Function
    Case 3
        'Tarifa 11.72 - Cocorretagem
        Dim intCodPlano As Integer
        intCodPlano = M46V111.gCod_plano
           
        lSelect = lSelect & "   SELECT                              " & vbCr
        lSelect = lSelect & "        COD_PLANO                      " & vbCr
        lSelect = lSelect & "       ,COD_CORRETOR                   " & vbCr
        lSelect = lSelect & "       ,B.NOM_CORR                     " & vbCr
        lSelect = lSelect & "       ,B.NOM_ABRV_CORR                " & vbCr
        lSelect = lSelect & "       ,B.NUM_SUSEP                    " & vbCr
        lSelect = lSelect & "       ,COD_RAMO                       " & vbCr
        lSelect = lSelect & "       ,IND_LIDER                      " & vbCr
        lSelect = lSelect & "       ,PERC_CORR                      " & vbCr
        lSelect = lSelect & "   FROM                                " & vbCr
        lSelect = lSelect & "       TBM_PLANO_CORRETOR A            " & vbCr
        lSelect = lSelect & "       INNER JOIN YTAB_CORR B          " & vbCr
        lSelect = lSelect & "       ON A.COD_CORRETOR = B.COD_CORR  " & vbCr
        lSelect = lSelect & "   WHERE                               " & vbCr
        lSelect = lSelect & "       A.COD_PLANO = " & CStr(intCodPlano) & vbCr
        lSelect = lSelect & "       AND A.PERC_CORR > 0             " & vbCr
        lSelect = lSelect & "    order by   PERC_CORR desc          " & vbCr
   
        If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
            Exit Function
        End If
        
        If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
            Exit Function
        End If
        
        If rsCORR.RecordCount = 0 Then
            Call Me.gpFecha2(bdP0044700)
            pMensagem = "FGE0003 - Corretor inv�lido (corretor n�o localizado com nome ou parte " & _
                        "do nome = " & pCorretor & ")."
            Exit Function
        End If
        rsCORR.MoveLast
        ReDim pSaida(2, 8)
        lI = 0
        rsCORR.MoveFirst
        While Not rsCORR.EOF
            lI = lI + 1
            pSaida(lI, 1) = rsCORR!Cod_Corretor
            pSaida(lI, 2) = rsCORR!NOM_CORR
            pSaida(lI, 3) = rsCORR!NOM_ABRV_CORR
            pSaida(lI, 4) = rsCORR!NUM_SUSEP
            pSaida(lI, 5) = rsCORR!Cod_Plano
            pSaida(lI, 6) = rsCORR!Cod_Ramo
            pSaida(lI, 7) = rsCORR!IND_LIDER
            pSaida(lI, 8) = rsCORR!PERC_CORR
                
            rsCORR.MoveNext
        Wend
        Call Me.gpFecha2(bdP0044700)
        gfObtCorretor = True
        Exit Function
    End Select
    
End Function
Public Function gfObtBuscaParceiroProlaborista(ByVal pTipoParceiro As Integer, _
                                               ByVal pCodigoCorretor As String, _
                                               ByRef pSaida As Variant, _
                                               ByRef pMensagem As String) As Boolean
'Fun��o: obt�m dados de corretores.

'Par�metros de entrada...pTipoParceiro......2 = C�digo do Para Buscar o Parceido do Corretor.
'                        pCodigoCorretor...C�digo do corretor.
'Par�metro de sa�da......pSaida......Array contendo resultados.
'                                    (x, 1) C�digo do corretor.
'                                    (x, 2) Nome do corretor.
    Dim bdYas_ND As ADODB.Connection
    Dim lI As Double  'Contador.
    Dim lSelect As String       'Comando SELECT.
    Dim rsCORR As clsYasRecordSet    'Procedure: proc_BuscaParceiro.

    gfObtBuscaParceiroProlaborista = False

    '1. Verifica par�metros.
    'TipoParceiro.
    If Not gfPreenchido(pTipoParceiro) Then
        pMensagem = "FTE0067 - Op��o inv�lida (Tipo Parceiro - op��o n�o preenchida)."
        Exit Function
    End If

    If Not IsNumeric(pTipoParceiro) Then
        pMensagem = "FTE0067 - Op��o inv�lida (Tipo Parceiro - op��o n�o num�rica)."
        Exit Function
    End If

    'CodigoCorretor.
    If Not gfPreenchido(pCodigoCorretor) Then
        pMensagem = "FTE0068 - Erro ao obter corretor (c�digo do corretor n�o preenchido)."
        Exit Function
    End If

    '2. Busca informa��o.
    lSelect = "proc_BuscaParceiro " & pTipoParceiro & ", " & pCodigoCorretor
    If gfAbrBasDados("YAS_ND", gTipYAS_ND, gEndYAS_ND, bdYas_ND, pMensagem) = False Then
        Exit Function
    End If

    If gfObtRegistro(bdYas_ND, lSelect, rsCORR, pMensagem) = False Then
        Exit Function
    End If

    If rsCORR.RecordCount = 0 Then
        Call Me.gpFecha2(bdYas_ND)
        pMensagem = "FGE0003 - Corretor inv�lido (Nenhum prolaborista localizado para o corretor " & _
                    pCodigoCorretor & ")."
        Exit Function
    End If

    rsCORR.MoveLast

    ReDim pSaida(rsCORR.RecordCount, 2)
    lI = 0

    rsCORR.MoveFirst

    While Not rsCORR.EOF
        lI = lI + 1
        pSaida(lI, 1) = rsCORR!Cod_Corr
        pSaida(lI, 2) = rsCORR!NOM_CORR
        rsCORR.MoveNext
    Wend
    Call Me.gpFecha2(bdYas_ND)
    gfObtBuscaParceiroProlaborista = True
End Function

Public Function gfObtLimiteMaximoProlaboreCorretor(ByVal pTipoConsulta As Integer, _
                                                   ByVal pCodigoCorretor As String, _
                                                   ByRef pSaida As Variant, _
                                                   ByRef pMensagem As String) As Boolean
'Fun��o: obt�m dados de Limite Maximo do Percentual de Comiss�o do Prolaborista.

'Par�metros de entrada...pCodigoCorretor...C�digo do corretor.
'Par�metro de sa�da......pSaida......Array contendo resultados.
'                                    (x, 1) Cod_Ramo
'                                    (x, 2) Cod_Modalidade
'                                    (x, 3) Cod_Corr
'                                    (x, 4) Cod_Tip_Limite
'                                    (x, 5) Val_Limite
'                                    (x, 6) Cod_Tarif
'                                    (x, 7) Cod_Sequencia
'                                    (x, 8) Cod_Tipo_Reg

    Dim bdDiversos As ADODB.Connection
    Dim lI As Double  'Contador.
    Dim lSelect As String       'Comando SELECT.
    Dim rsLimites As clsYasRecordSet    'Tabela: Tbm_Limite.

    gfObtLimiteMaximoProlaboreCorretor = False

    If pTipoConsulta = 1 Then

        '1. Verifica par�metros.
        'CodigoCorretor.
        If Not gfPreenchido(pCodigoCorretor) Then
            pMensagem = "FTE0068 - Erro ao obter corretor (c�digo do corretor n�o preenchido)."
            Exit Function
        End If

        '2. Busca informa��o.
        lSelect = "SELECT    Cod_Ramo, " & _
                  "Cod_Modalidade, " & _
                  "Cod_Corr, " & _
                  "Cod_Tip_Limite, " & _
                  "Val_Limite, " & _
                  "Cod_Tarif, " & _
                  "Cod_Sequencia, " & _
                  "Cod_Tipo_Reg " & _
                  "FROM   Tbm_Limite WITH (NOLOCK) " & _
                  "WHERE  Cod_Ramo = " & gCod_Ramo & " " & _
                  "AND    Cod_Modalidade = " & M46V111.gCod_plano & " " & _
                  "AND    Cod_Tipo_Reg = 'E' " & _
                  "AND    Cod_Tip_Limite = '4' " & _
                  "AND    Cod_Corr = " & pCodigoCorretor & " " & _
                  "AND    Cod_Tarif = " & gCodProduto
    Else
        lSelect = "SELECT  Cod_Ramo, " & _
                  "Cod_Modalidade, " & _
                  "Cod_Corr, " & _
                  "Cod_Tip_Limite, " & _
                  "Val_Limite, " & _
                  "Cod_Tarif, " & _
                  "Cod_Sequencia, " & _
                  "Cod_Tipo_Reg " & _
                  "FROM   Tbm_Limite WITH (NOLOCK) " & _
                  "WHERE  Cod_Ramo = " & gCod_Ramo & " " & _
                  "AND    Cod_Modalidade = " & M46V111.gCod_plano & " " & _
                  "AND    Cod_Tipo_Reg = 'D' " & _
                  "AND    Cod_Tip_Limite = '4' " & _
                  "AND    Cod_Tarif = " & gCodProduto
    End If

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdDiversos, pMensagem) = False Then
        Exit Function
    End If

    If gfObtRegistro(bdDiversos, lSelect, rsLimites, pMensagem) = False Then
        Exit Function
    End If

    If rsLimites.RecordCount = 0 Then
        Call Me.gpFecha2(bdDiversos)

        If pTipoConsulta = 2 Then
            pMensagem = "FGE0003 - Limite Prolabore Inv�lido (Nenhum limite de prolabore localizado)."
        End If
        ReDim Preserve pSaida(0, 0)
        Exit Function
    End If

    rsLimites.MoveLast

    ReDim Preserve pSaida(rsLimites.RecordCount, 8)
    lI = 0

    rsLimites.MoveFirst

    While Not rsLimites.EOF
        lI = lI + 1
        pSaida(lI, 1) = rsLimites!Cod_Ramo
        pSaida(lI, 2) = rsLimites!Cod_Modalidade
        pSaida(lI, 3) = rsLimites!Cod_Corr
        pSaida(lI, 4) = rsLimites!Cod_Tip_Limite
        pSaida(lI, 5) = rsLimites!Val_Limite
        pSaida(lI, 6) = rsLimites!Cod_Tarif
        pSaida(lI, 7) = rsLimites!Cod_Sequencia
        pSaida(lI, 8) = rsLimites!Cod_Tipo_Reg
        rsLimites.MoveNext
    Wend

    Call Me.gpFecha2(bdDiversos)
    gfObtLimiteMaximoProlaboreCorretor = True
End Function


Public Function gfObtLimiteMaximoProlaboreCorretor2(ByVal pCodigoCorretor As String, _
                                                    ByRef pLimite As Double, _
                                                    ByRef pMensagem As String) As Boolean
'Fun��o: obt�m dados de Limite Maximo do Percentual de Comiss�o do Prolaborista.

'Par�metros de entrada...pCodigoCorretor...C�digo do corretor.
'Par�metro de sa�da......pLimite......

    Dim bdDiversos As ADODB.Connection
    Dim lI As Double  'Contador.
    Dim lSelect As String       'Comando SELECT.
    Dim rsLimites As clsYasRecordSet    'Tabela: Tbm_Limite.

    On Error GoTo ControleErro

    gfObtLimiteMaximoProlaboreCorretor2 = False

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdDiversos, pMensagem) = False Then
        Exit Function
    End If


    '1. Verifica par�metros.
    'CodigoCorretor.
    If Not gfPreenchido(pCodigoCorretor) Then
        'pMensagem = "FTE0068 - Erro ao obter corretor (c�digo do corretor n�o preenchido)."
        'Exit Function
        'Busca Limite Default
        GoTo LimiteDefault
    End If

    '2. Busca informa��o.
    lSelect = "SELECT    Cod_Ramo, " & _
              "Cod_Modalidade, " & _
              "Cod_Corr, " & _
              "Cod_Tip_Limite, " & _
              "Val_Limite, " & _
              "Cod_Tarif, " & _
              "Cod_Sequencia, " & _
              "Cod_Tipo_Reg " & _
              "FROM   Tbm_Limite WITH (NOLOCK) " & _
              "WHERE  Cod_Ramo = " & gCod_Ramo & " " & _
              "AND    Cod_Modalidade = " & M46V111.gCod_plano & " " & _
              "AND    Cod_Tipo_Reg = 'E' " & _
              "AND    Cod_Tip_Limite = '4' " & _
              "AND    Cod_Corr = " & pCodigoCorretor & " " & _
              "AND    Cod_Tarif = " & gCodProduto


    If gfObtRegistro(bdDiversos, lSelect, rsLimites, pMensagem) = False Then
        Exit Function
    End If

    If rsLimites.RecordCount = 0 Then
        GoTo LimiteDefault
    End If

    pLimite = rsLimites!Val_Limite
    GoTo Fim


LimiteDefault:

    lSelect = "SELECT  Cod_Ramo, " & _
              "Cod_Modalidade, " & _
              "Cod_Corr, " & _
              "Cod_Tip_Limite, " & _
              "Val_Limite, " & _
              "Cod_Tarif, " & _
              "Cod_Sequencia, " & _
              "Cod_Tipo_Reg " & _
              "FROM   Tbm_Limite WITH (NOLOCK) " & _
              "WHERE  Cod_Ramo = " & gCod_Ramo & " " & _
              "AND    Cod_Modalidade = " & M46V111.gCod_plano & " " & _
              "AND    Cod_Tipo_Reg = 'D' " & _
              "AND    Cod_Tip_Limite = '4' " & _
              "AND    Cod_Tarif = " & gCodProduto


    If gfObtRegistro(bdDiversos, lSelect, rsLimites, pMensagem) = False Then
        Exit Function
    End If

    If rsLimites.RecordCount = 0 Then
        Call Me.gpFecha2(bdDiversos)
        pMensagem = "FGE0003 - Limite Prolabore Inv�lido (Nenhum limite de prolabore localizado)."
        Exit Function
    End If


    pLimite = rsLimites!Val_Limite

    Call Me.gpFecha2(bdDiversos)

Fim:
    gfObtLimiteMaximoProlaboreCorretor2 = True

    Exit Function
ControleErro:
    pMensagem = "gfObtLimiteMaximoProlaboreCorretor2 ERRO - " & Err & " " & Err.Description
    Call gpGraLog(0, pMensagem)

End Function

Public Function gfObtBuscaBeneficiariosAntigos(ByVal pNum_Apol_Ant As Double, _
                                               ByVal pNum_Item As Integer, _
                                               ByRef pSaida As Variant, _
                                               ByRef pMensagem As String) As Boolean
'Fun��o: obt�m dados de corretores.

'Par�metros de entrada...pNum_Apol_Ant...... = Numero da apolice que ser� anterior a renova��o.
'                        pNum_Item.......... = Numero do item que est� sendo renovado.
'Par�metro de sa�da......pSaida............. = Array contendo resultado.
'                                              (x, 1) Texto Espec.
    Dim bdP00MULTI As ADODB.Connection
    Dim lI As Double                    'Contador.
    Dim lSelect As String               'Comando SELECT.
    Dim rsTexto As clsYasRecordSet      'Procedure: proc_BuscarBeneficiariosAntigos.

    gfObtBuscaBeneficiariosAntigos = False

    '1. Verifica par�metros.
    If Not gfPreenchido(pNum_Apol_Ant) Then
        pMensagem = "FTE9999 - Op��o inv�lida (Numero da ap�lice - op��o n�o preenchida)."
        Exit Function
    End If

    If Not gfPreenchido(pNum_Item) Then
        pMensagem = "FTE9999 - Op��o inv�lida (Numero do item - op��o n�o preenchida)."
        Exit Function
    End If

    'If Not IsNumeric(pNum_Apol_Ant) Then
    '    pMensagem = "FTE0067 - Op��o inv�lida (Numero do Item - op��o n�o preenchida)."
    '    Exit Function
    'End If

    '2. Busca informa��o.
    lSelect = "proc_BuscarBeneficiariosAntigos " & pNum_Apol_Ant & ", " & pNum_Item
    '    If gM46V999.gfAbrBasDados("SYASPORTAL", gM46V999.gTipSyasPortal, gM46V999.gEndSyasPortal, bdSyasPortal, pMensagem) = False Then
    '        Exit Function
    '    End If

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdP00MULTI, pMensagem) = False Then
        GoTo SAIDA
    End If

    If gfObtRegistro(bdP00MULTI, lSelect, rsTexto, pMensagem) = False Then
        Exit Function
    End If

    If rsTexto.RecordCount = 0 Then
        Call Me.gpFecha2(bdP00MULTI)
        Exit Function
    End If

    rsTexto.MoveLast

    ReDim pSaida(rsTexto.RecordCount, 1)
    lI = 0

    While Not rsTexto.EOF
        lI = lI + 1
        pSaida(lI, 1) = rsTexto(0)
        rsTexto.MoveNext
    Wend
    Call Me.gpFecha2(bdP00MULTI)
    gfObtBuscaBeneficiariosAntigos = True

SAIDA:
    Call Me.gpFecha2(bdP00MULTI)
End Function

Public Function gfObtBuscaNovaClausulaBeneficiaria(ByVal pTipoPessoa As Long, _
                                                   ByVal pTipoBem As Integer, _
                                                   ByRef pSaida As Variant, _
                                                   ByRef pMensagem As String) As Boolean
'Fun��o: obt�m dados de corretores.

'Par�metros de entrada...pNum_Apol_Ant...... = Numero da apolice que ser� anterior a renova��o.
'                        pNum_Item.......... = Numero do item que est� sendo renovado.
'Par�metro de sa�da......pSaida............. = Array contendo resultado.
'                                              (x, 1) Texto Espec.
    Dim bdP00MULTI As ADODB.Connection
    Dim lI As Double                    'Contador.
    Dim lSelect As String               'Comando SELECT.
    Dim rsTexto As clsYasRecordSet      'Procedure: proc_BuscarBeneficiariosAntigos.

    gfObtBuscaNovaClausulaBeneficiaria = False

    '1. Verifica par�metros.
    If Not gfPreenchido(pTipoPessoa) Then
        pMensagem = "FTE9999 - Op��o inv�lida (Tipo de Pessoa - op��o n�o preenchida)."
        Exit Function
    End If

    If Not gfPreenchido(pTipoBem) Then
        pMensagem = "FTE9999 - Op��o inv�lida (Tipo do Bem compreendido - op��o n�o preenchida)."
        Exit Function
    End If

    'If Not IsNumeric(pNum_Apol_Ant) Then
    '    pMensagem = "FTE0067 - Op��o inv�lida (Numero do Item - op��o n�o preenchida)."
    '    Exit Function
    'End If

    '2. Busca informa��o.
    lSelect = "proc_BuscarNovaClausulaBeneficiaria " & pTipoPessoa & ", " & pTipoBem
    
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdP00MULTI, pMensagem) = False Then
        GoTo SAIDA
    End If

    If gfObtRegistro(bdP00MULTI, lSelect, rsTexto, pMensagem) = False Then
        Call Me.gpFecha2(bdP00MULTI)
        GoTo SAIDA
    End If

    If rsTexto.RecordCount = 0 Then
        Call Me.gpFecha2(bdP00MULTI)
        pMensagem = "(Nenhum benefici�rio antigo localizado."
        Exit Function
    End If

    rsTexto.MoveLast

    ReDim pSaida(rsTexto.RecordCount, 1)
    lI = 0

    rsTexto.MoveFirst

    While Not rsTexto.EOF
        lI = lI + 1
        pSaida(lI, 1) = rsTexto(0)
        rsTexto.MoveNext
    Wend
    
    gfObtBuscaNovaClausulaBeneficiaria = True

SAIDA:
    Call Me.gpFecha2(bdP00MULTI)
    
End Function


Public Function gfObtLimiteMaximoAgravo(ByVal pCodigoCorretor As String, _
                                        ByRef pLimite As Double, _
                                        ByRef pMensagem As String) As Boolean
'Fun��o: obt�m dados de Limite Maximo do Percentual de Comiss�o do Prolaborista.

'Par�metros de entrada...pCodigoCorretor...C�digo do corretor.
'Par�metro de sa�da......pLimite......

    Dim bdDiversos As ADODB.Connection
    Dim lI As Double  'Contador.
    Dim lSelect As String       'Comando SELECT.
    Dim rsLimites As clsYasRecordSet    'Tabela: Tbm_Limite.

    On Error GoTo ControleErro

    gfObtLimiteMaximoAgravo = False

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdDiversos, pMensagem) = False Then
        Exit Function
    End If


    '1. Verifica par�metros.
    'CodigoCorretor.
    If Not gfPreenchido(pCodigoCorretor) Then
        GoTo LimiteDefault
    End If

    '2. Busca informa��o.
    lSelect = "SELECT    Cod_Ramo, " & _
              "Cod_Modalidade, " & _
              "Cod_Corr, " & _
              "Cod_Tip_Limite, " & _
              "Val_Limite, " & _
              "Cod_Tarif, " & _
              "Cod_Sequencia, " & _
              "Cod_Tipo_Reg " & _
              "FROM   Tbm_Limite WITH (NOLOCK) " & _
              "WHERE  Cod_Ramo = " & gCod_Ramo & " " & _
              "AND    Cod_Modalidade = " & M46V111.gCod_plano & " " & _
              "AND    Cod_Tipo_Reg = 'E' " & _
              "AND    Cod_Tip_Limite = '5' " & _
              "AND    Cod_Corr = " & pCodigoCorretor & " " & _
              "AND    Cod_Tarif = " & gCodProduto


    If gfObtRegistro(bdDiversos, lSelect, rsLimites, pMensagem) = False Then
        Exit Function
    End If

    If rsLimites.RecordCount = 0 Then
        GoTo LimiteDefault
    End If

    pLimite = rsLimites!Val_Limite
    GoTo Fim


LimiteDefault:
    'rodrigoS era Cod_Tip_Limite = '5'
    lSelect = "SELECT  Cod_Ramo, " & _
              "Cod_Modalidade, " & _
              "Cod_Corr, " & _
              "Cod_Tip_Limite, " & _
              "Val_Limite, " & _
              "Cod_Tarif, " & _
              "Cod_Sequencia, " & _
              "Cod_Tipo_Reg " & _
              "FROM   Tbm_Limite WITH (NOLOCK) " & _
              "WHERE  Cod_Ramo = " & gCod_Ramo & " " & _
              "AND    Cod_Modalidade = " & M46V111.gCod_plano & " " & _
              "AND    Cod_Tipo_Reg = 'D' " & _
              "AND    Cod_Tip_Limite = '5' " & _
              "AND    Cod_Tarif = " & gCodProduto


    If gfObtRegistro(bdDiversos, lSelect, rsLimites, pMensagem) = False Then
        Exit Function
    End If

    If rsLimites.RecordCount = 0 Then
        Call Me.gpFecha2(bdDiversos)
        pMensagem = "FGR0001 - Limite Agravo Inv�lido (Nenhum limite de Agravo localizado)."
        Exit Function
    End If


    pLimite = rsLimites!Val_Limite

    Call Me.gpFecha2(bdDiversos)

Fim:
    gfObtLimiteMaximoAgravo = True

    Exit Function
ControleErro:
    pMensagem = "gfObtLimiteMaximoAgravo ERRO - " & Err & " " & Err.Description
    Call gpGraLog(0, pMensagem)

End Function




Public Function gfObtUniProdutiva(ByVal pCodCorretor As String, ByRef pCodUniNegocios As Integer, _
                                  ByRef pCodProdutor As Integer, ByRef pMensagem As String, _
                                  Optional ByRef pTotalRegistros As Integer) As Boolean
'Fun��o: obt�m unidade de neg�cios e produtor do corretor.

'Par�metro de entrada...pCodCorretor......C�digo do corretor.
'Par�metros de sa�da....pCodUniNegocios...C�digo da unidade de neg�cios.
'                       pCodProdutor......C�digo do produtor.
'                       pMensagem.........Mensagem.

    Dim bdP0044700 As ADODB.Connection    'Arquivo - P0044700.
    Dim lI As Double    'Contador.
    Dim lNomCorretor As String        'Nome do corretor.
    Dim lSelect As String   'Comando SELECT.
    Dim rsUNIDCORRPROD As clsYasRecordSet    'Registro - YTAB_UNID_CORR_PROD.

    gfObtUniProdutiva = False

    '1. Inicializa sa�da.
    pCodUniNegocios = 0
    pCodProdutor = 0

    '2. Verifica par�metro.
    If Not gfPreenchido(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade de neg�cios e produtor (c�digo do corretor n�o " & _
                    "preenchido)."
        Exit Function
    End If

    '3. Busca informa��o.
    If Not IsNumeric(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade neg�cios e produtor (c�digo do corretor n�o " & _
                    "num�rico)."
        Exit Function
    End If

    lSelect = "SELECT A.COD_UNID, A.COD_PROD FROM YTAB_UNID_CORR_PROD A "
    lSelect = lSelect & " INNER JOIN YTAB_CORR B ON A.COD_CORR = B.COD_CORR "
    lSelect = lSelect & " WHERE A.COD_CORR = " & pCodCorretor
    'lSelect = lSelect & " AND   A.COD_UNID IN(1751,211,231,251,271,1691,2531,2671,3731,3791,5551,5771,6631,6651,6711)" '[FRH]-F:900052
    lSelect = lSelect & " ORDER BY A.COD_CORR, A.COD_UNID"
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    If gfObtRegistro(bdP0044700, lSelect, rsUNIDCORRPROD, pMensagem) = False Then
        Exit Function
    End If
    If rsUNIDCORRPROD.RecordCount = 0 Then
        Call Me.gpFecha2(bdP0044700)
        If gfNomCorretor(pCodCorretor, lNomCorretor, pMensagem) = False Then
            Exit Function
        End If
        pMensagem = "FTE0069 - Erro ao obter unidade de neg�cios e produtor (corretor " & pCodCorretor & _
                  " - " & lNomCorretor & " sem informa��o de unidade de neg�cios e produtor)."
        Exit Function
    Else
        pTotalRegistros = rsUNIDCORRPROD.RecordCount
    End If

    '4. Encerra processo.
    pCodUniNegocios = rsUNIDCORRPROD!COD_UNID
    pCodProdutor = rsUNIDCORRPROD!Cod_Prod
    Call Me.gpFecha2(bdP0044700)

    gfObtUniProdutiva = True
End Function
Public Function gfObtProdutor(ByVal pCodCorretor As String, ByRef pCodUniNegocios As Integer, _
                              ByRef pCodProdutor As Integer, ByRef pMensagem As String, _
                              Optional ByRef pTotalRegistros As Integer) As Boolean
'Fun��o: obt�m unidade de neg�cios e produtor do corretor.

'Par�metro de entrada...pCodCorretor......C�digo do corretor.
'Par�metros de sa�da....pCodUniNegocios...C�digo da unidade de neg�cios.
'                       pCodProdutor......C�digo do produtor.
'                       pMensagem.........Mensagem.

    Dim bdP0044700 As ADODB.Connection    'Arquivo - P0044700.
    Dim lI As Double    'Contador.
    Dim lNomCorretor As String        'Nome do corretor.
    Dim lSelect As String   'Comando SELECT.
    Dim rsUNIDCORRPROD As clsYasRecordSet    'Registro - YTAB_UNID_CORR_PROD.

    gfObtProdutor = False

    '1. Inicializa sa�da.
    pCodProdutor = 0

    '2. Verifica par�metro.
    If Not gfPreenchido(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade de neg�cios e produtor (c�digo do corretor n�o " & _
                    "preenchido)."
        Exit Function
    End If

    '3. Busca informa��o.
    If Not IsNumeric(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade neg�cios e produtor (c�digo do corretor n�o " & _
                    "num�rico)."
        Exit Function
    End If
    lSelect = "SELECT COD_PROD FROM YTAB_UNID_CORR_PROD WHERE " & _
            " COD_CORR = " & pCodCorretor & _
            " and cod_unid = " & pCodUniNegocios
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    If gfObtRegistro(bdP0044700, lSelect, rsUNIDCORRPROD, pMensagem) = False Then
        Exit Function
    End If
    If rsUNIDCORRPROD.RecordCount = 0 Then
        Call Me.gpFecha2(bdP0044700)
        If gfNomCorretor(pCodCorretor, lNomCorretor, pMensagem) = False Then
            Exit Function
        End If
        pMensagem = "FTE0069 - Erro ao obter unidade de neg�cios e produtor (corretor " & pCodCorretor & _
                  " - " & lNomCorretor & " sem informa��o de unidade de neg�cios e produtor)."
        Exit Function
    Else
        pTotalRegistros = rsUNIDCORRPROD.RecordCount
    End If

    '4. Encerra processo.
    pCodProdutor = rsUNIDCORRPROD!Cod_Prod
    Call Me.gpFecha2(bdP0044700)

    gfObtProdutor = True
End Function

Public Function gfObtUnidCorr(ByVal pCodCorretor As String, ByRef pCodUniNegocios As Integer, ByRef pMensagem As String, _
                              ByRef pTotalRegistros As Integer) As Boolean
'Fun��o: obt�m unidade de neg�cios e produtor do corretor.

'Par�metro de entrada...pCodCorretor......C�digo do corretor.
'Par�metros de sa�da....pCodUniNegocios...C�digo da unidade de neg�cios.
'                       pCodProdutor......C�digo do produtor.
'                       pMensagem.........Mensagem.

    Dim bdP0044700 As ADODB.Connection    'Arquivo - P0044700.
    Dim lI As Double    'Contador.
    Dim lNomCorretor As String        'Nome do corretor.
    Dim lSelect As String   'Comando SELECT.
    Dim rsUNIDCORRPROD As clsYasRecordSet    'Registro - YTAB_UNID_CORR_PROD.

    gfObtUnidCorr = False

    '2. Verifica par�metro.
    If Not gfPreenchido(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade de neg�cios e produtor (c�digo do corretor n�o " & _
                    "preenchido)."
        Exit Function
    End If

    '3. Busca informa��o.
    If Not IsNumeric(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade neg�cios e produtor (c�digo do corretor n�o " & _
                    "num�rico)."
        Exit Function
    End If
    lSelect = "SELECT COD_PROD FROM YTAB_UNID_CORR_PROD WHERE " & _
            " COD_CORR = " & pCodCorretor & _
            " and cod_unid = " & pCodUniNegocios
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    If gfObtRegistro(bdP0044700, lSelect, rsUNIDCORRPROD, pMensagem) = False Then
        Exit Function
    End If
    If rsUNIDCORRPROD.RecordCount = 0 Then
        pTotalRegistros = 0
    Else
        pTotalRegistros = rsUNIDCORRPROD.RecordCount
    End If

    '4. Encerra processo.
    Call Me.gpFecha2(bdP0044700)

    gfObtUnidCorr = True
End Function


Public Function gfRepBasDados_DAO(ByVal pArquivo As String) As Boolean
'Fun��o: Repara arquivo.

    gfRepBasDados_DAO = False

    On Error Resume Next
    DBEngine.RepairDatabase pArquivo
    If Err <> 0 Then
        Call gpGraLog(2, "FTE0070 - Erro ao reparar arquivo " & pArquivo & ": " & Err & " - " & ERROR & _
                         ".")
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    gfRepBasDados_DAO = True
End Function
Public Function gfRetornarListIndex(ByVal pCombo As Object, ByVal pItemData As String) As Long
'Fun��o:

    Dim lCont As Long

    gfRetornarListIndex = -1
    For lCont = 0 To pCombo.ListCount - 1
        If pCombo.ItemData(lCont) = pItemData Then
            gfRetornarListIndex = lCont
            Exit Function
        End If
    Next
End Function
Public Function gfVerData(ByVal pOpcao As Byte, ByVal pData As String) As Boolean
'Fun��o: verifica data.

'Par�metros de entrada...pOpcao...Op��o 1 = data no formato aaaammdd.
'                                 Op��o 2 = data no formato dd/mm/aaaa.
'                        pData....Data no formato indicado pela op��o.

    Dim lAno As Integer    'Ano.
    Dim lAuxiliar As String    'Auxiliar.
    Dim lDia As Integer    'Dia.
    Dim lMes As Integer    'M�s.
    Dim lResto As Integer    'Resto.

    gfVerData = False

    '1. Verifica op��o.
    If Not gfPreenchido(pOpcao) Then
        Call gpGraLog(2, "FTE0071 - Op��o inv�lida (verifica data - op��o n�o preenchida)." & vbLf & _
                         vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If
    If Not IsNumeric(pOpcao) Then
        Call gpGraLog(2, "FTE0071 - Op��o inv�lida (verifica data - op��o n�o num�rica)." & vbLf & _
                         vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If
    If pOpcao <> 1 And pOpcao <> 2 Then
        Call gpGraLog(2, "FTE0071 - Op��o inv�lida (verifica data - diferente de 1 e 2: " & pOpcao & _
                         ")." & vbLf & vbLf & "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If

    '2. Separa dia, m�s e ano.
    Select Case pOpcao
    Case 1      'Data no formato aaaammdd.
        lAuxiliar = Format$(pData, "00000000")
        lDia = Mid$(lAuxiliar, 7, 2)
        lMes = Mid$(lAuxiliar, 5, 2)
        lAno = Mid$(lAuxiliar, 1, 4)
    Case 2
        lDia = Mid$(pData, 1, 2)
        lMes = Mid$(pData, 4, 2)
        lAno = Mid$(pData, 7, 4)
    End Select

    '3. Verifica m�s.
    If lMes < 1 Or lMes > 12 Then
        Exit Function
    End If

    '4. Verifica dia.
    If lDia < 1 Then
        Exit Function
    End If
    Select Case lMes
    Case 1, 3, 5, 7, 8, 10, 12
        If lDia > 31 Then
            Exit Function
        End If
    Case 4, 6, 9, 11
        If lDia > 30 Then
            Exit Function
        End If
    Case 2
        Select Case (lAno Mod 4)
        Case 0      'Ano bissexto.
            If lDia > 29 Then
                Exit Function
            End If
        Case Else   'Ano N�O bissexto.
            If lDia > 28 Then
                Exit Function
            End If
        End Select
    End Select

    gfVerData = True
End Function


Public Sub gpCalDifDatas(ByVal pOpcao As Byte, ByVal pData1 As Double, ByVal pData2 As Double, _
                         ByRef pSaida As Double)
'Procedure: calcula diferen�a entre datas.

'Par�metros de entrada...pOpcao...Op��o 1 = quantidade de dias.
'                                 Op��o 2 = quantidade de meses (prevista - n�o codificada).
'                                 Op��o 3 = quantidade de anos (idade).
'                        pData1...Data 1 no formato aaaammdd.
'                        pData2...Data 2 no formato aaaammdd.
'Par�metro de sa�da......pSaida...Quantidade calculada.

    Dim lAno1 As Integer    'Ano 1: menor data.
    Dim lAno2 As Integer    'Ano 2: maior data.
    Dim lAuxiliar As String    'Auxiliar.
    Dim lCalcula As Boolean    'Indica que � para calcular.
    Dim lDia1 As Integer    'Dia 1: menor data.
    Dim lDia2 As Integer    'Dia 2: maior data.
    Dim lMes1 As Integer    'M�s 1: menor data.
    Dim lMes2 As Integer    'M�s 2: maior data.

'1. Inicializa sa�da.
    pSaida = 0

    '2. Verifica op��o.
    If Not gfPreenchido(pOpcao) Then
        Call gpGraLog(2, "FTE0072 - Op��o inv�lida (calcula diferen�a entre datas - op��o n�o " & _
                         "preenchida)." & vbLf & vbLf & "O sistema ser� encerrado.")
        Exit Sub    'End   ***MOSILVA Verificar
    End If
    If Not IsNumeric(pOpcao) Then
        Call gpGraLog(2, "FTE0072 - Op��o inv�lida (calcula diferen�a entre datas - op��o n�o " & _
                         "num�rica)." & vbLf & vbLf & "O sistema ser� encerrado.")
        Exit Sub    'End   ***MOSILVA Verificar
    End If
    If pOpcao <> 1 And pOpcao <> 3 Then
        Call gpGraLog(2, "FTE0072 - Op��o inv�lida (calcula diferen�a entre datas - op��o diferente " & _
                         "de 1, 2 e 3)." & vbLf & vbLf & "O sistema ser� encerrado.")
        Exit Sub    'End   ***MOSILVA Verificar
    End If

    '3. Verifica datas.
    If gfVerData(1, pData1) = False Then
        Call gpGraLog(2, "FTE0073 - Erro ao calcular diferen�a entre datas (primeira data inv�lida - " & _
                         Me.gfForData(2, pData1) & ").")
        Exit Sub
    End If
    If gfVerData(1, pData2) = False Then
        Call gpGraLog(2, "FTE0073 - Erro ao calcular diferen�a entre datas (segunda data inv�lida - " & _
                         Me.gfForData(2, pData2) & ").")
        Exit Sub
    End If

    '4. Se datas iguais, quantidade calculada igual a 0.
    If pData1 = pData2 Then
        Exit Sub
    End If

    '5. Prepara campos.
    If pData1 < pData2 Then
        lAuxiliar = Format$(pData1, "00000000")    'Menor data.
        lDia1 = CInt(Mid$(lAuxiliar, 7, 2))
        lMes1 = CInt(Mid$(lAuxiliar, 5, 2))
        lAno1 = CInt(Mid$(lAuxiliar, 1, 4))
        lAuxiliar = Format$(pData2, "00000000")    'Maior data.
        lDia2 = CInt(Mid$(lAuxiliar, 7, 2))
        lMes2 = CInt(Mid$(lAuxiliar, 5, 2))
        lAno2 = CInt(Mid$(lAuxiliar, 1, 4))
    Else
        lAuxiliar = Format$(pData2, "00000000")    'Menor data.
        lDia1 = CInt(Mid$(lAuxiliar, 7, 2))
        lMes1 = CInt(Mid$(lAuxiliar, 5, 2))
        lAno1 = CInt(Mid$(lAuxiliar, 1, 4))
        lAuxiliar = Format$(pData1, "00000000")    'Maior data.
        lDia2 = CInt(Mid$(lAuxiliar, 7, 2))
        lMes2 = CInt(Mid$(lAuxiliar, 5, 2))
        lAno2 = CInt(Mid$(lAuxiliar, 1, 4))
    End If

    '4. Calcula.
    Select Case pOpcao
    Case 1      'Diferen�a de dias.
        lCalcula = False
        Do      'Controle de sa�da dentro do loop.
            If lAno1 = lAno2 Then
                If lMes1 = lMes2 Then
                    pSaida = pSaida + (lDia2 - lDia1)
                    Exit Do
                Else    'lMes1 < lMes2.
                    lCalcula = True
                End If
            Else    'lAno1 < lAno2.
                If lMes2 = 1 Then
                    pSaida = pSaida + lDia2
                    lAno2 = lAno2 - 1
                    lMes2 = 12
                    lDia2 = 31
                Else
                    If lMes2 = 12 And lDia2 = 31 Then
                        Select Case lAno2 Mod 4
                        Case 0      'Ano bissexto.
                            pSaida = pSaida + 366
                        Case Else   'Ano n�o bissexto.
                            pSaida = pSaida + 365
                        End Select
                        lAno2 = lAno2 - 1
                    Else
                        lCalcula = True
                    End If
                End If
            End If
            If lCalcula = True Then
                lCalcula = False
                pSaida = pSaida + lDia2
                lMes2 = lMes2 - 1
                Select Case lMes2
                Case 1, 3, 5, 7, 8, 10, 12
                    lDia2 = 31
                Case 4, 6, 9, 11
                    lDia2 = 30
                Case 2
                    Select Case lAno2 Mod 4
                    Case 0      'Ano bissexto.
                        lDia2 = 29
                    Case Else   'Ano N�O bissexto.
                        lDia2 = 28
                    End Select
                End Select
            End If
        Loop
    Case 2      'Diferen�a de meses.
        MsgBox "Codificar procedure gpCalDifDatas - op��o 2."
    Case 3      'Diferen�a de anos (idade).
        'a) Calcula diferen�a entre anos.
        pSaida = lAno2 - lAno1
        'b) Verifica m�s.
        If lMes2 > lMes1 Then
            Exit Sub
        End If
        If lMes2 < lMes1 Then
            pSaida = pSaida - 1
            Exit Sub
        End If
        'c) Verifica dia.
        If lDia2 < lDia1 Then
            pSaida = pSaida - 1
        End If
        Exit Sub
    End Select
End Sub
Public Sub gpPreencherCombo(ByRef pComboListBox As Object, prst As clsYasRecordSet, _
                            Optional pCampoDescricao As String, Optional pCampoID As String, _
                            Optional pValorDefault As String, Optional pPosicaoDefault As Integer = -1, _
                            Optional pForm As Object = Nothing)
'Procedure: preencher um listbox ou combobox a partir de uma cole��o de objetos de estrutura.

'Par�metros...pComboListBox...combobox ou listbox que ser� preenchido.
'             pColecao........cole��o de objetos de estrutura que implementam a classe stcGenerica.
'             pValorDefault...item que ser� posicionado o combo ou listbox.

    Dim bErroValor As Boolean         'Controle de erro no valor.
    Dim bFoco As Boolean    'Foco.

    On Error GoTo ControleErro

    bFoco = False
    If Not pForm Is Nothing Then
        If Not pForm.ActiveControl Is Nothing Then
            If UCase(TypeName(pForm.ActiveControl)) = "COMBOBOX" Then
                bFoco = (pForm.ActiveControl.Name = pComboListBox.Name And pForm.ActiveControl.index = pComboListBox.index)
            End If
        End If
    End If

    bErroValor = False
    If Not bFoco Then pComboListBox.Visible = False
    pComboListBox.Clear
    If Not (prst.BOF And prst.EOF) Then
        Do While Not prst.EOF
            pComboListBox.AddItem prst("" & pCampoDescricao & "")
            pComboListBox.ItemData(pComboListBox.NewIndex) = prst("" & pCampoID & "")
            prst.MoveNext
        Loop

        If pValorDefault <> "" Then
            bErroValor = True
            pComboListBox.Text = pValorDefault
        Else
            pComboListBox.ListIndex = pPosicaoDefault
        End If
    End If
    If Not bFoco Then pComboListBox.Visible = True
    Exit Sub

ControleErro:
    If bErroValor Then pComboListBox.ListIndex = pPosicaoDefault
    pComboListBox.Visible = True
End Sub

'[FRH]-> Movidas para serem utilizadas tamb�m no Reg.Padr�o [mnuEdiExcRegPad_click]
'[gpExcluiRegistros][antes no BAS111]
'[mfVerRegPadrao][antes no 001A]

Public Sub gpExcluiRegistros(ByVal pNosso_Numero As String, _
                             ByVal pNum_Item As Long, _
                             ByVal pPropTabela As String, _
                             ByVal pNomeTabImport As String, _
                             ByRef pMensagem As String)
    Dim lobjPedido As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED)
    Dim lobjstcPedido As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED)
    Dim lobjItem As clsA46V708A    'Objeto da classe clsA46V708A (TAB_PED_LOC)
    Dim lobjstcItem As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)

    Dim methodName As String
    Dim lKeyDynaTrace As Long
    methodName = "gpExcluiRegistros"
    lKeyDynaTrace = DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    On Error GoTo ControleErro

    If pNum_Item = 0 Then
        Set lobjPedido = New clsA46V702A
        Set lobjPedido.M46V999 = Me

        Set lobjstcPedido = New stcA46V702B
        Set lobjstcPedido.M46V999 = Me

        lobjstcPedido.nosso_numero = pNosso_Numero
        lobjstcPedido.NUM_PI = M46V111.gNum_Pi
        If Not lobjPedido.mfExcluir(lobjstcPedido, pPropTabela, pNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedido.UltimoErro
            Set lobjPedido = Nothing
            Set lobjstcPedido = Nothing
            GoTo SAIDA
        End If
        Set lobjPedido = Nothing
        Set lobjstcPedido = Nothing
    Else
        Set lobjItem = New clsA46V708A
        Set lobjItem.M46V999 = Me

        Set lobjstcItem = New stcA46V708B
        Set lobjstcItem.M46V999 = Me

        lobjstcItem.nosso_numero = pNosso_Numero
        lobjstcItem.Num_Item = pNum_Item
        lobjstcItem.NUM_PI = M46V111.gNum_Pi
        
        If Not lobjItem.mfExcluir(lobjstcItem, pPropTabela, pNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjItem.UltimoErro
            Set lobjItem = Nothing
            Set lobjstcItem = Nothing
            GoTo SAIDA
        End If
        Set lobjItem = Nothing
        Set lobjstcItem = Nothing
    End If

    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    GoSub FechaObj
    Exit Sub

ControleErro:
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, Err.Description, 2, Class_Name, methodName)
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, pMensagem = "gpExcluiRegistros ERRO - " & Err & " " & Err.Description)
    GoTo SAIDA
FechaObj:
    Set lobjPedido = Nothing
    Set lobjstcPedido = Nothing
    Set lobjItem = Nothing
    Set lobjstcItem = Nothing
    Return
End Sub

Public Function mfVerRegPadrao(ByRef pMensagem As String) As Boolean

    If gPrograma <> "D046O003" And (InStr(UCase(gNomUsuario), "CYBER07") > 0 Or InStr(UCase(Command), "DEV") > 0) Then
        mfVerRegPadrao = True
        Exit Function
    End If
    
    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfVerRegPadrao"
    lKeyDynaTrace = DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    On Error GoTo ControleErro
    'Fun��o: salva padr�o do registro padr�o.

    'Par�metro de sa�da...pMensagem...Mensagem.

    Dim objPed As clsA46V702A    'Objeto da classe clsA46V702A       (Tab_Ped)
    Dim objPedLoc As clsA46V708A    'Objeto da classe clsA46V708A       (Tab_Ped_Loc)
    'Dim objJuros        As clsA46V711A      'Objeto da classe clsA46V711A       (Tab_Ped_Juros)
    Dim lDatAlteracao As Double         'Data da �ltima atualiza��o.
    Dim objstcPed As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (Tab_Ped)
    Dim objstcPedLoc As stcA46V708B   'Objeto da classe de estrutura stcA46V708B (Tab_Ped_Loc)
    'Dim objstcJuros     As stcA46V711B      'Objeto da classe de estrutura stcA46V711B (Tab_Ped_Juros)
    '    Dim lMensagem       As String

    Dim lobjstcCobert As stcA46V704B
    Dim lobjCobert As clsA46V704A

    Dim pCodProduto As Integer
    Dim pProtocolo As Long
    Dim pSequencia As Long
    Dim pDescrProd As String

    Dim lPropTabela As String
    Dim lbdP00MPED As ADODB.Connection
    Dim lrst As clsYasRecordSet
    Dim lSql As String

    Dim lobjstcCobertFraq As stcA46V738B
    Dim lobjCobertFraq As clsA46V738A


    'Call gpGraLog(0, "vai abrir o ini")
    Call gfObtSYASMULTINI(lPropTabela)
    'Call gpGraLog(0, "fechou o ini ")

    mfVerRegPadrao = False

    'Call gpGraLog(0, "1. Posiciona controle.")
    '1. Posiciona controle.
    lDatAlteracao = Format(Date, "yyyymmdd")

    'Call gpGraLog(0, "2. Instancio os objetos Tab_Ped")
    '2. Instancio os objetos Tab_Ped
    Set objPed = New clsA46V702A
    Set objPed.M46V999 = Me


    Set objstcPed = New stcA46V702B
    Set objstcPed.M46V999 = Me

    'Call gpGraLog(0, "3. Verifica se tem registro padr�o de dados comuns.")
    '3. Verifica se tem registro padr�o de dados comuns.
    objstcPed.nosso_numero = "PADRAO" & gNomUsuario6
    If Not objPed.mfCarregarColecao(objstcPed, lPropTabela, "", True) Then
        pMensagem = objPed.UltimoErro
        Call gpGraLog(1, pMensagem)
        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        GoTo SAIDA
    End If

    'Call gpGraLog(0, "Se cole��o estiver vazia ent�o � inclu�do o registro padr�o.")
    'Se cole��o estiver vazia ent�o � inclu�do o registro padr�o.
    If objPed.Colecao.Count = 0 Then
        'N�o tem registro padr�o - ent�o inclui.
        With objstcPed
            'Carrego as insforma��es no STC
            .nosso_numero = "PADRAO" & gNomUsuario6
            .Cod_Tip_Doc = e_TipDoc_Cotacao
            .Cod_Sit_Ped = e_SitPedido_SemCalculo
            .DAT_DIGIT = lDatAlteracao
            .Cod_User_Digit = gNomUsuario6
            .Tip_Cust_Emis = 1
            .Cod_Estip = gfPlanoBasico(gCod_Ramo)
        End With
        'Chamo a fun��o de incluir o registro na TAb_PEd
        If Not objPed.mfIncluir(objstcPed, lPropTabela, "", tagDynaTrace) Then
            pMensagem = objPed.UltimoErro
            Call gpGraLog(1, pMensagem)
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If

    'Call gpGraLog(0, "4. Instancio os objetos do Tab_Ped_Loc")
    '4. Instancio os objetos do Tab_Ped_Loc
    Set objPedLoc = New clsA46V708A
    Set objPedLoc.M46V999 = Me

    Set objstcPedLoc = New stcA46V708B
    Set objstcPedLoc.M46V999 = Me

    Set objPedLoc.ConexaoBD = objPed.ConexaoBD

    'Call gpGraLog(0, "5. Verifica se tem registro padr�o do item.")
    '5. Verifica se tem registro padr�o do item.
    objstcPedLoc.nosso_numero = "PADRAO" & gNomUsuario6
    If Not objPedLoc.mfCarregarColecao(objstcPedLoc, lPropTabela, "", True) Then
        pMensagem = objPedLoc.UltimoErro
        Call gpGraLog(1, pMensagem)
        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        GoTo SAIDA
    End If

    'Call gpGraLog(0, "6. Se n�o tem registro padr�o - ent�o inclui.")
    '6. Se n�o tem registro padr�o - ent�o inclui.
    If objPedLoc.Colecao.Count = 0 Then
        With objstcPedLoc
            'Carrego as informa��es no stc
            .nosso_numero = "PADRAO" & gNomUsuario6
            .Num_Item = 1
            .Cod_Sit_item = e_SitItem_NaoCalc
            .Tip_Emissao = e_TipEmissao_SeguroNovo
            .DAT_DIGIT = lDatAlteracao
            .Cod_User_Digit = gNomUsuario6
            .Cod_Plano = 1
        End With
        'Chamo a fun��o para incluir o registro na TAB_PED_LOC
        If Not objPedLoc.mfIncluir(objstcPedLoc, lPropTabela, "", tagDynaTrace) Then
            pMensagem = objPedLoc.UltimoErro
            Call gpGraLog(1, pMensagem)
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
    End If


    '>---[FRH]--->
    '[Criar um registro padr�o para cada um dos ramos de acordo com tabela]
    '[Acessar a tbm_ramo]


    Dim lSql_ramos As String
    Dim rs_ramos As clsYasRecordSet
    Dim bdP00MPED As ADODB.Connection

    Dim objPed_TMP As clsA46V702A
    Dim objstcPed_TMP As stcA46V702B

    Dim objPedLoc_TMP As clsA46V708A
    Dim objstcPedLoc_tmp As stcA46V708B

    Dim objCorr_TMP As clsA46V705A
    Dim objstcCorr_TMP As stcA46V705B

    If gfAbrBasDados("P00MPED", gTipP00MPED, gEndP00MPED, bdP00MPED, pMensagem) Then

        lSql_ramos = "SELECT * FROM TBM_RAMO WITH (NOLOCK)"
        If gPrograma = "D046O003" Or gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
            lSql_ramos = lSql_ramos & " where cod_ramo in (112,113,114)"
        Else
            lSql_ramos = lSql_ramos & " where cod_ramo in (112,113,114,710,510,150)"
        End If

        If gfObtRegistro(bdP00MPED, lSql_ramos, rs_ramos, pMensagem) Then
            Do While Not rs_ramos.EOF

                'Call gpGraLog(0, "[COMECA ramo] " & rs_ramos("cod_ramo"))
                '[TAB_PED]->
                '*****************************************************************************************************************

                Set objPed_TMP = New clsA46V702A
                'Call gpGraLog(0, "[fez 1 new  ")
                Set objPed_TMP.M46V999 = Me

                Set objstcPed_TMP = New stcA46V702B
                'Call gpGraLog(0, "[fez 2 new  ")
                Set objstcPed_TMP.M46V999 = Me

                '[Verificar se existe registro padr�o de dados comuns]

                objstcPed_TMP.nosso_numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                'Call gpGraLog(0, "[carrega tab_ped ] " & objstcPed_TMP.NOSSO_NUMERO)
                If Not objPed_TMP.mfCarregarColecao(objstcPed_TMP, lPropTabela, "", True) Then
                    pMensagem = objPed_TMP.UltimoErro:
                    Call gpGraLog(1, pMensagem)
                    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If
                'Call gpGraLog(0, "[retornou count ] " & objPed_TMP.Colecao.Count)
                '[Verificar se a cole��o � vazia, ent�o incluir registro padr�o]
                If objPed_TMP.Colecao.Count = 0 Then
                    'Call gpGraLog(0, "COLECAO VAZIA ")
                    With objstcPed_TMP
                        '[Carregar informa��es no STC]
                        .Cod_Ramo = gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                        .Cod_Tip_Doc = e_TipDoc_Cotacao
                        .Cod_Sit_Ped = e_SitPedido_SemCalculo
                        .DAT_DIGIT = lDatAlteracao
                        .Cod_User_Digit = gNomUsuario6

                        '[Per�odo de vig�ncia]
                        .Dat_Ini_Vig = gHoje
                        If Right(.Dat_Ini_Vig, 4) = "0229" Then
                            .Dat_Ini_Vig = gHoje - 1
                        End If
                        .Dat_Ter_Vig = .Dat_Ini_Vig + 10000

                        '[Produto]
                        '   If Me.gfBusProduto(gHoje, pCodProduto, pProtocolo, pSequencia, pDescrProd, pmensagem, 0) = False Then
                        '       Call gpGraLog(1, pmensagem)
                        '  End If
                        .Cod_Produto = 0
                        .Cod_Reg = 1
                        .Cod_Estip = gfPlanoBasico(objstcPed_TMP.Cod_Ramo)
                        .Tip_Cust_Emis = 1
                    End With
                    '[Incluir registro na TAB_PED]
                    'Call gpGraLog(0, "Incluir registro na TAB_PED")
                    If Not objPed_TMP.mfIncluir(objstcPed_TMP, lPropTabela, "", tagDynaTrace) Then
                        pMensagem = objPed_TMP.UltimoErro:
                        Call gpGraLog(0, pMensagem)
                        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                        GoTo SAIDA
                    End If
                End If

                '*****************************************************************************************************************
                '<-[TAB_PED]
                'Call gpGraLog(0, "COLECAO VAZIA TAB_PED_CORR ")
                '[TAB_PED_CORR]->
                '*****************************************************************************************************************

                Set objCorr_TMP = New clsA46V705A
                'Call gpGraLog(0, "[fez 3 new  ")
                Set objCorr_TMP.M46V999 = Me

                Set objstcCorr_TMP = New stcA46V705B
                'Call gpGraLog(0, "[fez 4 new  ")
                Set objstcCorr_TMP.M46V999 = Me


                '[Verificar se existe registro padr�o de dados comuns]
                objstcCorr_TMP.nosso_numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                'Call gpGraLog(0, "[colecao tab_ped_corr   " & objstcCorr_TMP.NOSSO_NUMERO)
                If Not objCorr_TMP.mfCarregarColecao(objstcCorr_TMP, lPropTabela, "") Then
                    pMensagem = objCorr_TMP.UltimoErro:
                    Call gpGraLog(1, pMensagem):
                    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If

                '[Verificar se a cole��o � vazia, ent�o incluir registro padr�o]
                'Call gpGraLog(0, "[retornou cout " & objCorr_TMP.Colecao.Count)
                If objCorr_TMP.Colecao.Count = 0 Then

                    With objstcCorr_TMP
                        .nosso_numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                        .Num_Ped = 0
                        .Num_Ordem_Corr = 1
                        If gCod_Corretor <> 0 Then
                            .Cod_Corr = gCod_Corretor
                        End If
                    End With
                    'Call gpGraLog(0, "[Incluir registro na TAB_PED_corr 2 ] ")
                    If Not objCorr_TMP.mfIncluir(objstcCorr_TMP, lPropTabela, "", tagDynaTrace) Then
                        pMensagem = objCorr_TMP.UltimoErro:
                        Call gpGraLog(1, pMensagem)
                        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                        GoTo SAIDA
                    End If
                End If

                '*****************************************************************************************************************
                '<-[TAB_PED_CORR]
                'Call gpGraLog(0, "[Incluir registro na TAB_PED_LOC] ")
                '[TAB_PED_LOC]->
                '*****************************************************************************************************************

                Set objPedLoc_TMP = New clsA46V708A
                Set objPedLoc_TMP.M46V999 = Me

                Set objstcPedLoc_tmp = New stcA46V708B
                Set objstcPedLoc_tmp.M46V999 = Me

                Set objPedLoc_TMP.ConexaoBD = objPed_TMP.ConexaoBD

                '[Verificar se tem registro padr�o de dados comuns]
                objstcPedLoc_tmp.nosso_numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                'Call gpGraLog(0, "[vai carregar tab_ped_loc " & objstcPedLoc_tmp.NOSSO_NUMERO)
                If Not objPedLoc_TMP.mfCarregarColecao(objstcPedLoc_tmp, lPropTabela, "", True) Then
                    pMensagem = objPedLoc_TMP.UltimoErro
                    Call gpGraLog(1, pMensagem)
                    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If

                'Call gpGraLog(0, "[carregpou colecao de tab_ped_loc ] ")

                '[Verificar se a cole��o � vazia, ent�o incluir registro padr�o]
                If objPedLoc_TMP.Colecao.Count = 0 Then
                    With objstcPedLoc_tmp
                        '[Carregar informa��es no STC]
                        .Cod_Ramo = gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                        .Num_Item = 1
                        .Cod_Sit_item = e_SitItem_NaoCalc
                        .Tip_Emissao = e_TipEmissao_SeguroNovo
                        .DAT_DIGIT = lDatAlteracao
                        .Cod_User_Digit = gNomUsuario6
                    End With
                    '[Incluir registro na TAB_PED_LOC]
                    'Call gpGraLog(0, "incluir tab_ped_loc 2")
                    If Not objPedLoc_TMP.mfIncluir(objstcPedLoc_tmp, lPropTabela, "", tagDynaTrace) Then
                        pMensagem = objPedLoc_TMP.UltimoErro:
                        Call gpGraLog(1, pMensagem)
                        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                        GoTo SAIDA
                    End If
                End If

                '5. Verifica se tem registro padr�o .
                Set lobjstcCobert = New stcA46V704B
                Set lobjstcCobert.M46V999 = Me

                Set lobjCobert = New clsA46V704A
                Set lobjCobert.M46V999 = Me
                'Call gpGraLog(0, "[Incluir registro na TAB_PED_COBERT ] ")

                lobjstcCobert.nosso_numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                'Call gpGraLog(0, "carregar tab_ped_COBERT " & lobjstcCobert.NOSSO_NUMERO)
                Set lobjstcCobert.ObjPai = objstcPedLoc
                Set lobjCobert.ConexaoBD = objPed_TMP.ConexaoBD
                If Not lobjCobert.mfCarregarColecao(lobjstcCobert, lPropTabela, "") Then
                    pMensagem = lobjCobert.UltimoErro
                    Call gpGraLog(1, pMensagem)
                    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If

                '6. Se n�o tem registro padr�o - ent�o inclui.
                If lobjCobert.Colecao.Count = 0 Then
                    Set lobjstcCobert = New stcA46V704B
                    Set lobjstcCobert.M46V999 = Me

                    Set lobjCobert = New clsA46V704A
                    Set lobjCobert.M46V999 = Me

                    With lobjstcCobert
                        .nosso_numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                        .Num_Item = 1
                        .Cod_Ramo = gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                        .Cod_Cobert = 1
                        .Tab_Cobert = 1
                    End With
                    Set lobjCobert.ConexaoBD = objPed.ConexaoBD
                    'Call gpGraLog(0, "incluir tab_ped_COBERT " & lobjstcCobert.NOSSO_NUMERO)
                    If Not lobjCobert.mfIncluir(lobjstcCobert, lPropTabela, "", False, False, tagDynaTrace) Then
                        'Call gpGraLog(0, "ERRO NO INSERT DA TAB_PED_COBERT ] ")
                        pMensagem = lobjCobert.UltimoErro
                        Call gpGraLog(1, pMensagem)
                        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                        GoTo SAIDA
                    End If
                End If

                '*****************************************************************************************************************
                
               '5. Verifica se tem registro padr�o . TAB_PED_COBERT_FRANQUIA
                Set lobjstcCobertFraq = New stcA46V738B
                Set lobjstcCobertFraq.M46V999 = Me

                Set lobjCobertFraq = New clsA46V738A
                Set lobjCobertFraq.M46V999 = Me

                lobjstcCobertFraq.nosso_numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                Set lobjstcCobertFraq.ObjPai = lobjstcCobert
                Set lobjCobertFraq.ConexaoBD = lobjCobert.ConexaoBD
                If Not lobjCobertFraq.mfCarregarColecao(lobjstcCobertFraq, lPropTabela, "") Then
                    pMensagem = lobjCobert.UltimoErro
                    Call gpGraLog(1, pMensagem)
                    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If

                '6. Se n�o tem registro padr�o - ent�o inclui.
                If lobjCobertFraq.Colecao.Count = 0 Then

                    With lobjstcCobertFraq
                        .nosso_numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                        .Num_Item = 1
                        .Cod_Ramo = gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                        .Cod_Cobert = 1
                        .Cod_Modalidade = 1
                    End With
                    Set lobjCobertFraq.ConexaoBD = lobjCobert.ConexaoBD
                    If Not lobjCobertFraq.mfIncluir(lobjstcCobertFraq, lPropTabela, "", tagDynaTrace) Then
                        pMensagem = lobjCobert.UltimoErro
                        Call gpGraLog(1, pMensagem)
                        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                        GoTo SAIDA
                    End If
                End If

                rs_ramos.MoveNext
            Loop
            Set rs_ramos = Nothing
        Else
            Call gpGraLog(0, "FTE00010" & pMensagem)
        End If

        'deleta indice Tab_Ped_Cobert_Compl_IND03 - Estava dando erro quando o erro j� estava deletado
        If gPrograma <> "D046O003" And gTipUsuario = e_Tipo_Usu_Corretor And gBaseLocal = True Then
            'Call gpGraLog(0, "[ENTROU NO DROP ] ")
            On Error Resume Next
            Call gfExeSQL(bdP00MPED, "Drop Index Tab_Ped_Cobert_Compl_IND03 on Tab_Ped_Cobert_Compl", pMensagem)
            'Call gpGraLog(0, "retornou a mensagem  " & pMensagem)
            pMensagem = ""
            Err = 0
            On Error GoTo ControleErro
        End If

        'Call gpGraLog(0, "[VAI FECHAR ] ")
        Call Me.gpFecha2(bdP00MPED)
    Else
        Call gpGraLog(1, pMensagem)
    End If

    'Call gpGraLog(0, "[fim loop ramos ] ")
    '<---[FRH]---<

    If gTipUsuario = e_Tipo_Usu_Corretor And gPrograma = "P46V001" Then
        If Not gfAbrBasDados("P00MPED", gTipP00MPED, gEndP00MPED, lbdP00MPED, pMensagem) Then
            Call gpGraLog(0, pMensagem)
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            Exit Function
        End If
        Dim lObservacao As String
        lSql = " Select * from TabConv WITH (NOLOCK)"
        'Call gpGraLog(0, "[SQL  ] " & lSql)
        If Not gfObtRegistro(lbdP00MPED, lSql, lrst, pMensagem) Then
            'Call gpGraLog(0, "RETORNOU " & lSql)
            pMensagem = ""

            lSql = " update tab_ped  set cod_ramo_ativ = 1 where cod_ramo_ativ = 2"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 2 where cod_ramo_ativ = 3"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 3 where cod_ramo_ativ = 4"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 4 where cod_ramo_ativ = 5"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 100 where cod_ramo_ativ = 8"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 8 where cod_ramo_ativ = 9"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 9 where cod_ramo_ativ = 100"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 100 where cod_ramo_ativ = 10"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 101 where cod_ramo_ativ = 11"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 11 where cod_ramo_ativ = 100"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 100 where cod_ramo_ativ = 12"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 12 where cod_ramo_ativ = 101"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 101 where cod_ramo_ativ = 15"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 15 where cod_ramo_ativ = 100"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 100 where cod_ramo_ativ = 16"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 16 where cod_ramo_ativ = 13"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 102 where cod_ramo_ativ = 17"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 17 where cod_ramo_ativ = 14"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 103 where cod_ramo_ativ = 19"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 101"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 101 where cod_ramo_ativ = 20"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 20 where cod_ramo_ativ = 100"
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 100 where cod_ramo_ativ = 21"
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 21 where cod_ramo_ativ = 102"
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 102 where cod_ramo_ativ = 22"
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 18"
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 23"
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 103"
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 24"
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 25"
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 26"
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)

            lSql = " Create table TabConv (data float null);"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)

            lSql = " Insert into TabConv values (" & Format(Date, "yyyymmdd") & ")"
            'Call gpGraLog(0, "[SQL  ] " & lSql)
            Call gfExeSQL(lbdP00MPED, lSql, pMensagem)
        End If
    End If

    mfVerRegPadrao = True
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    'Call gpGraLog(0, "[SAIU  ] ")
    GoSub FechaObj
    Exit Function

FechaObj:
    '7. Destruo os objetos
    Set lobjstcCobert = Nothing
    Set lobjCobert = Nothing
    Set objPed = Nothing
    Set objPedLoc = Nothing
    'Set objJuros = Nothing
    Set objstcPed = Nothing
    Set objstcPedLoc = Nothing
    'Set objstcJuros = Nothing
    Set rs_ramos = Nothing
    Set lbdP00MPED = Nothing
    Set lrst = Nothing
    Set bdP00MPED = Nothing
    Set objPed_TMP = Nothing
    Set objstcPed_TMP = Nothing
    Set objPedLoc_TMP = Nothing
    Set objstcPedLoc_tmp = Nothing
    Set objCorr_TMP = Nothing
    Set objstcCorr_TMP = Nothing
    Set lobjstcCobertFraq = Nothing
    Set lobjCobertFraq = Nothing
    
    Return
    Exit Function
ControleErro:
    pMensagem = Err.Description
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
    Call gpGraLog(1, " mfVerRegPadrao " & pMensagem)
    GoTo SAIDA
End Function


'Fun��o para verificar grupo de corretor na base de dados
'Ficha 1000153 - Automatiza��o do agrupamento de corretores no Syas ap�s o cadastro realizado pelo help-desk.
Public Function gfVerCorrPermitido_TBS_Corr_Grupo(ByVal pCorrInstalacao As Long, _
                                                  ByVal pCorrSelecionado As Long, _
                                                  ByVal pCodRamo As Integer, _
                                                  ByRef pMensagem As String) As Boolean
    Dim bdP0044700 As ADODB.Connection    'Arquivo - P0044700.
    Dim lCnGed As ADODB.Connection    'Arquivo - P0044700.
    Dim lSelect As String       'Comando SELECT.
    Dim rsCORR As clsYasRecordSet    'Tabela: YTAB_CORR.
    Dim lCorrInstacao As String
    Dim lCorrSelecionado As String


    Err.Clear
    pMensagem = ""
    On Error GoTo TrataErro: gfVerCorrPermitido_TBS_Corr_Grupo = False

    lCorrInstacao = Format(pCorrInstalacao, "0000000")
    lCorrSelecionado = Format(pCorrSelecionado, "0000000")

    '1. Verifica par�metros.
    'a) Op��o.
    If Not gfPreenchido(pCorrInstalacao) Then
        pMensagem = "FTE0067 - Op��o inv�lida (obter corretor - op��o n�o preenchida)."
        GoTo SAIDA
    End If
    'b) Corretor.
    If Not gfPreenchido(pCorrSelecionado) Then
        pMensagem = "FTE0068 - Erro ao obter corretor (c�digo do corretor n�o preenchido)."
        GoTo SAIDA
    End If
    'C) Ramo
    If pCodRamo = 0 Then
        pMensagem = "FTE0068 - Erro ao obter corretor (c�digo do ramo n�o preenchido)."
        GoTo SAIDA
    End If

    If Not (lCorrInstacao = lCorrSelecionado) Then
        If gPrograma = "D046O003" Then
            If gfAbrBasDados("GED", gTipP0044700, gEndP0044700, lCnGed, pMensagem) = False Then
                GoTo SAIDA
            End If
            lSelect = "SELECT CodCorr FROM "
            lSelect = lSelect & " (SELECT CodCorr FROM GED.dbo.Func_BuscaCorrGrupo('" & lCorrInstacao & "', 112)) as Grupo"
            lSelect = lSelect & " WHERE CodCorr = " & val(lCorrSelecionado)

            If gfObtRegistro(lCnGed, lSelect, rsCORR, pMensagem) = True Then
                If rsCORR.EOF = True Then
                    pMensagem = "FGE0017 - C�digo corretor informado (" & lCorrSelecionado & ") n�o permitido para este usu�rio(" & gNomUsuario & " - Corr:" & lCorrInstacao & ")"
                    GoTo SAIDA
                End If
            Else
                GoTo SAIDA
            End If
        Else
            If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
                GoTo SAIDA
            End If
            'instala��o  = C�digo aliado  e informado = c�d l�der
            lSelect = " SELECT * FROM TBS_Corr_Grupo WITH (NOLOCK) " & _
                      "WHERE CodAliado = '" & lCorrInstacao & _
                      "' AND CodCorr = '" & lCorrSelecionado & _
                      "' AND CodRamo = 112 "
            If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
                GoTo SAIDA
            End If
            If rsCORR.RecordCount = 0 Then
                'instala��o  = c�d l�der  e  informado = C�digo aliado
                lSelect = " SELECT * FROM TBS_Corr_Grupo WITH (NOLOCK) " & _
                          "WHERE CodCorr = '" & lCorrInstacao & _
                          "' AND CodAliado = '" & lCorrSelecionado & _
                          "' AND CodRamo = 112 "
                If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
                    GoTo SAIDA
                End If
            Else
                gfVerCorrPermitido_TBS_Corr_Grupo = True
                GoTo SAIDA
            End If
            If rsCORR.RecordCount = 0 Then
                'instala��o =  c�digo aliado e informado = c�digo aliado s� que ambos s�o aliados do mesmo l�der.
                lSelect = "SELECT A.CodCorr" & _
                        "  FROM TBS_Corr_Grupo AS A " & _
                        " INNER JOIN TBS_Corr_Grupo AS B " & _
                        "    ON A.CodCorr = B.CodCorr " & _
                        " WHERE (((A.CodAliado)= '" & lCorrInstacao & "' ) " & _
                        "   AND ((B.CodAliado)='" & lCorrSelecionado & "') " & _
                        "   AND ((A.CodRamo)=112)) "
                If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If rsCORR.RecordCount = 0 Then
                    Call Me.gpFecha2(bdP0044700)
                    pMensagem = "FGE0003 - " & gPrefixo & "Corretor inv�lido (c�digo de corretor n�o " & _
                                "localizado - " & lCorrInstacao & ")."
                    GoTo SAIDA
                Else
                    gfVerCorrPermitido_TBS_Corr_Grupo = True
                    GoTo SAIDA
                End If
            Else
                gfVerCorrPermitido_TBS_Corr_Grupo = True
                GoTo SAIDA
            End If
        End If
    End If

    '=================== OK Permitido!!!
    gfVerCorrPermitido_TBS_Corr_Grupo = True
    '===================
SAIDA:
    If gfVerCorrPermitido_TBS_Corr_Grupo = False Then
        Call gpGraLog(2, pMensagem)
    End If
    GoSub FechaObj
    Err.Clear
    Exit Function

TrataErro:
    If gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "gfVerCorrPermitido_TBS_Corr_Grupo Erro: " & pMensagem
    GoTo SAIDA
FechaObj:
    Call Me.gpFecha2(lCnGed)
    Set lCnGed = Nothing
    Set rsCORR = Nothing
    Set bdP0044700 = Nothing
    Return
End Function

Public Function gfCalNumPedido(ByRef pNumPedido As String, _
                               ByRef pMensagem As String, _
                               Optional ByRef pBasDados As ADODB.Connection) As Boolean
'Fun��o: calcula n�mero do pedido.

'Par�metro de entrada...pBasDados....Base de dados.
'Par�metros de sa�da....pNumPedido...N�mero do pedido.
'                       pMensagem....Mensagem.

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "gfCalNumPedido"
    lKeyDynaTrace = DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lAuxiliar As Long       'Auxiliar.
    Dim lobjUltPed As clsA46V701A    'Objeto da classe clsA46V701A (TAB_PED_ULT)
    Dim lobjstcUltPed As stcA46V701B    'Objeto da classe de estrutura stcA46V701B (TAB_PED_ULT)
    Dim lPropTabela As String

    On Error GoTo ControleErro: gfCalNumPedido = False

    Call gfObtSYASMULTINI(lPropTabela)

    Set lobjUltPed = New clsA46V701A
    Set lobjUltPed.M46V999 = Me

    Set lobjstcUltPed = New stcA46V701B

    Set lobjUltPed.ConexaoBD = pBasDados

    If gPrograma = "D046O003" Then
        If Not lobjUltPed.mfCarregarColecaoSyasPortal(gNomUsuario6, lPropTabela) Then
            pMensagem = lobjUltPed.UltimoErro
            GoTo ControleErro
        End If
        If lobjUltPed.Colecao.Count = 0 Then
            pMensagem = "N�o foi possivel gerar o numero do pedido."
            GoTo ControleErro
        Else
            Set lobjstcUltPed = lobjUltPed.Colecao.Item(1)
        End If

        'Exclui o pedido anteriormente gravado com este numero
        Dim lClsPed As clsA46V702A
        Dim lStcPed As stcA46V702B

        Set lClsPed = New clsA46V702A
        Set lClsPed.M46V999 = Me

        Set lStcPed = New stcA46V702B
        Set lStcPed.M46V999 = Me

        lStcPed.Num_Ped = lobjstcUltPed.Ult_Num_Ped
        If lClsPed.mfCarregarColecao(lStcPed, lPropTabela, "") = True Then
            For Each lStcPed In lClsPed.Colecao
                If lClsPed.mfExcluir(lStcPed, lPropTabela, "", tagDynaTrace) = False Then
                    pMensagem = lClsPed.UltimoErro
                    GoTo ControleErro
                End If
            Next lStcPed
        Else
            pMensagem = lClsPed.UltimoErro
            GoTo ControleErro
        End If

    Else
        '1. Busca o �ltimo n�mero do pedido
        If Not lobjUltPed.mfCarregarColecao(lPropTabela) Then
            pMensagem = lobjUltPed.UltimoErro
            GoTo ControleErro
        End If

        '2. Calcula n�mero do pedido e atualiza tabela.
        If lobjUltPed.Colecao.Count = 0 Then
            lobjstcUltPed.Ult_Num_Ped = 1
            lobjstcUltPed.Cod_User_Alt = gNomUsuario6
            lobjstcUltPed.Dat_Alt = Me.gfForData(11)
            If Not lobjUltPed.mfIncluir(lobjstcUltPed, lPropTabela) Then
                pMensagem = lobjUltPed.UltimoErro
                GoTo ControleErro
            End If
        Else
            Set lobjstcUltPed = lobjUltPed.Colecao.Item(1)
            If lobjstcUltPed.Ult_Num_Ped = 0 Then
                lobjstcUltPed.Ult_Num_Ped = 1
            Else
                lobjstcUltPed.Ult_Num_Ped = lobjstcUltPed.Ult_Num_Ped + 1
                If lobjstcUltPed.Ult_Num_Ped > 999999 Then
                    lobjstcUltPed.Ult_Num_Ped = 1
                End If
            End If
            lobjstcUltPed.Cod_User_Alt = gNomUsuario6
            lobjstcUltPed.Dat_Alt = Me.gfForData(11)
            If Not lobjUltPed.mfAlterar(lobjstcUltPed, lPropTabela) Then
                pMensagem = lobjUltPed.UltimoErro
                GoTo ControleErro
            End If
        End If
    End If

    pNumPedido = Format$(lobjstcUltPed.Ult_Num_Ped, "000000")

    gfCalNumPedido = True
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function

ControleErro:
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfCalNumPedido ERRO - " & Err & " " & ERROR & pMensagem)
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfCalNumPedido ERRO - " & Err & " " & ERROR & pMensagem, 2, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lClsPed = Nothing
    Set lStcPed = Nothing
    Set lobjUltPed = Nothing
    Set lobjstcUltPed = Nothing
    Return
End Function


Public Function gfCopProposta(ByVal pAntNosNumero As String, _
                              ByVal pNovNosNumero As String, _
                              ByVal pNovNumPedido As String, _
                              ByRef pMensagem As String, _
                              Optional ByVal pNovPI As Long, _
                              Optional pDatRecCia As Long, _
                              Optional pCodRamo As Integer, _
                              Optional ByRef pCalcLMI As Integer, _
                              Optional pNaoRepProduto As Boolean, _
                              Optional pReentrada As Integer, _
                              Optional pCodProduto As Integer = 0) As Boolean
'Fun��o: monta comando INSERT gen�rico para dados de proposta.

'Par�metros de entrada...pAntNumItem....Anterior: n�mero do item.
'                        pmfNovNumItem....Novo: n�mero do item.
'Par�metro de sa�da......pMensagem......Mensagem.

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "gfCopProposta"
    lKeyDynaTrace = DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim bdBasDados As ADODB.Connection    'Base de dados.
    Dim flTabela As clsYasField   'Campos da tabela.
    Dim rsTabela As clsYasRecordSet    'Registro da tabela.
    Dim lCount As Long        'Utilizado no for...next.
    Dim lIdeTabela As String          'Nome da tabela.
    Dim lContinuar As Boolean         'Indicador se continua o procedimento ou n�o.
    Dim lSelect As String
    Dim lInsert As String
    Dim lPropTabela As String
    Dim lTipEmissao As Integer
    Dim lCodTarifaCopia As Integer
    Dim lCodRamo As Integer
    Dim lCodClasf As String
    Dim lCodClasfCompl As String
    Dim lCarregarDePara As Boolean
    
    On Error GoTo ControleErro

    gfCopProposta = False

    Call gfObtSYASMULTINI(lPropTabela)

    '07/01/2010 - Ficha: 0901900 - C�pia de cota��o/proposta c/ �ltima tarifa -- inicio
    Dim lCodProduto As Integer      'C�digo do produto.
    Dim lCodProtocolo As Long           'C�digo do protocolo.
    Dim lCodSequencia As Long           'C�digo da seq��ncia.
    Dim lDesTarifa As String      'Descri��o da tarifa.
    Dim lMensagem As String     'Mensagem.
    
    Dim Per_Max_Comis_Corr As Double
    Dim Dat_Ini_Vig_Orig As Long
    Dim Dat_Ter_Vig_Orig As Long

    Call gTipEmissao(pAntNosNumero, lTipEmissao, lCodTarifaCopia, lPropTabela)

    If lCodTarifaCopia < TarifaPrecificacaoMaritima And (lTipEmissao = e_TipEmissao_Endosso Or pReentrada = 1) Then
        lCodProduto = lCodTarifaCopia
    End If

    If gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        'Ficha 901718
        If pNaoRepProduto = False Then    'Reposicionar na �ltima tarifa
            If Me.gfBusProduto(gHoje, lCodProduto, lCodProtocolo, lCodSequencia, lDesTarifa, lMensagem, pCodRamo) = _
               False Then
                Call gpGraLog(1, lMensagem)
                Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If
            '07/01/2010 - Ficha: 0901900 - C�pia de cota��o/proposta c/ �ltima tarifa -- fim
        End If
    End If

    If lCodProduto = 0 Then
    lCodProduto = pCodProduto
    End If

    '1. Obt�m registros da tabela.
    If gfAbrBasDados("P00MPED", gTipP00MPED, gEndP00MPED, bdBasDados, _
                     pMensagem) = False Then
        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        GoTo SAIDA
    End If

    For lCount = 1 To 33
        lContinuar = True
        Select Case lCount
        Case 1
            lIdeTabela = "TAB_PED"
        Case 2
            lIdeTabela = "TAB_PED_CORR"
        Case 3
            lIdeTabela = "TAB_PED_COSS"
        Case 4
            lIdeTabela = "TAB_PED_JUROS"
        Case 5
            lIdeTabela = "TAB_PED_VENC"
        Case 6
            lIdeTabela = "TAB_PED_ENDER"
        Case 7
            lIdeTabela = "TAB_CTRL_EMIS"
        Case 8
            lIdeTabela = "TAB_PED_LOC"
        Case 9
            lIdeTabela = "TAB_PED_CLAU"
        Case 10
            lIdeTabela = "TAB_PED_COBERT"
        Case 11
            lIdeTabela = "TAB_PED_LOC_BEM"
        Case 12
            lIdeTabela = "TAB_PED_LOC_SEGURO"
        Case 13
            lIdeTabela = "TAB_PED_OBS"
        Case 14
            lIdeTabela = "TAB_PED_INSP"
        Case 15
            lIdeTabela = "TAB_PED_AUTO"
        Case 16
            lIdeTabela = "TAB_PED_COBERT_COMPL"
        Case 17
            lIdeTabela = "TAB_PED_FAT_RISCO"
        Case 18
            lIdeTabela = "TAB_PED_DIZER"
        Case 19
            lIdeTabela = "TAB_PED_ESPEC"
        Case 20
            lIdeTabela = "TAB_PED_ESPEC2"
        Case 21
            lIdeTabela = "Tab_Ped_Loc_LMI"
        Case 22
            lIdeTabela = "Tab_Canal_Emiss"
        Case 23
            lIdeTabela = "TAB_CTRL_EMIS_COMPL"
        Case 24
            lIdeTabela = "TAB_PED_LOC_PESSOAS"
        Case 25
            lIdeTabela = "TAB_PED_TELEFONES"
        Case 26
            lIdeTabela = "TAB_PED_LOC_COMPL"
        Case 27
            lIdeTabela = "TAB_PED_COBERT_FRANQUIA"
        Case 28
            lIdeTabela = "TAB_PED_TEXTOS"
        Case 29
            lIdeTabela = "TAB_PED_INSP_COMPL"
        Case 30
            lIdeTabela = "TAB_PED_LOC_BICICLETA"
        Case 31
            lIdeTabela = "TAB_PED_LIBERACAO_SERASA"
        Case 32
            lIdeTabela = "TAB_PED_ALL_RISKS"
        Case 33
            lIdeTabela = "TAB_PED_OBRAS_ARTE"
        End Select

        If gEhRetornoSiscota(pAntNosNumero) And lIdeTabela = "TAB_PED_LIBERACAO_SERASA" Then
            GoTo ProximaTabela
        End If

        lSelect = "SELECT * FROM " & lPropTabela & lIdeTabela & " WITH (NOLOCK) WHERE Nosso_Numero = '" & pAntNosNumero & "'"
        If gfObtRegistro(bdBasDados, lSelect, rsTabela, pMensagem) = False Then
            Call Me.gpFecha2(bdBasDados)
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        If rsTabela.EOF = True Then
            lContinuar = False
            If lIdeTabela = "TAB_PED" Then
                pMensagem = "FNL0003 - " & gPrefixo & "Pedido n�o localizado (" & pAntNosNumero & _
                          " - c�pia da cota��o/proposta.)."
                Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If
        End If
        If lContinuar = True Then
            '2. Grava nova cota��o/proposta.
            While Not rsTabela.EOF
                lCodRamo = 0
                lCodClasf = 0
                lCodClasfCompl = 0

                If lIdeTabela = "TAB_PED" Or lIdeTabela = "TAB_PED_COBERT" Or lIdeTabela = "TAB_PED_COBERT_COMPL" Then
                    lCodRamo = rsTabela!Cod_Ramo
                    If lIdeTabela = "TAB_PED" Then
                        M46V111.gCod_plano = rsTabela!Cod_Estip
                        lCarregarDePara = False
                        
                        lCarregarDePara = lCodProduto >= 1080 And gRst_Tbm_De_Para.State <> adStateOpen
                        If lCodProduto >= 1080 And Not lCarregarDePara Then
                            lCarregarDePara = gRst_Tbm_De_Para.RecordCount = 0
                        End If
                    
                        If lCarregarDePara Then
                            gCodProduto = lCodProduto
                            gCod_Ramo = lCodRamo
                            M46V111.gCod_plano = gfPlanoBasico(lCodRamo, rsTabela!Cod_Estip)
                            Call mfCarregarDePara
                        End If
                        
                    End If
                End If
                
                If lIdeTabela = "TAB_PED_LOC" And lCodProduto >= 1080 Then
                    lCodClasf = rsTabela!Cod_Clasf
                    lCodClasfCompl = rsTabela!Cod_Compl_Clasf
                End If
                
                If lIdeTabela = "TAB_CTRL_EMIS_COMPL" Then
                    Per_Max_Comis_Corr = rsTabela!Per_Max_Comis_Corr
                    Dat_Ini_Vig_Orig = IIf(IsNull(rsTabela!Dat_Ini_Vig_Orig), 0, rsTabela!Dat_Ini_Vig_Orig)
                    Dat_Ter_Vig_Orig = IIf(IsNull(rsTabela!Dat_Ter_Vig_Orig), 0, rsTabela!Dat_Ter_Vig_Orig)
                End If
                
                lInsert = "INSERT INTO " & lPropTabela & lIdeTabela & " ("
                For Each flTabela In rsTabela.Fields
                    If InStr("PED_LOC_PESSOA_ID,TELEFONE_ID, ID_FRANQUIA_AVANCADA ", UCase(flTabela.Name)) = 0 Then
                        If Not ((flTabela.Name = "VAL_PR_ORIG" Or UCase(flTabela.Name) = "PREMIOINFORMADO") And gEhRetornoSiscota(pAntNosNumero) = 1) Then
                            lInsert = lInsert & flTabela.Name & ", "
                        End If
                    End If
                Next flTabela
                lInsert = Left$(lInsert, Len(lInsert) - 2) & ") VALUES ("
                For Each flTabela In rsTabela.Fields
                    'CR Inspecao
                    'Debug.Assert UCase(flTabela.Name) <> "IND_STATUSINSPECAO"
                    If Me.gOrigemChamada = e_Origem_Web And (UCase(flTabela.Name) = "IND_STATUSINSPECAO" Or UCase(flTabela.Name) = "NUM_INSP") Then
                        lInsert = lInsert & "0, "
                        GoTo Proximo
                    End If
                
                    Select Case UCase(flTabela.Name)
                    Case "PED_LOC_PESSOA_ID", "TELEFONE_ID", "ID_FRANQUIA_AVANCADA"
                    Case "NOSSO_NUMERO"
                        lInsert = lInsert & "'" & pNovNosNumero & "', "                         'Nosso n�mero
                    Case "NUM_PED"
                        lInsert = lInsert & pNovNumPedido & ", "                                'N�mero do pedido.
                    Case "SIT_DOCUMENTO", "COD_SIT_PED"
                        lInsert = lInsert & "1, "       'DOCUMENTO N�O CALCULADO.               'Sit_Documento.
                    Case "DAT_LIBER"
                        lInsert = lInsert & Left(Me.gfForData(3), 8) & ", "                        'Dat_Alt.
                    Case "DAT_PRI_CALCULO", "IND_PROMOCIONAL", "IND_TEXTO_REL_BENS", "IND_ENVIO_ALCADA"
                        lInsert = lInsert & "0, "                                               'Dat_Pri_Calculo.
                    Case "VAL_PR_ORIG", "PREMIOINFORMADO"
                        If gEhRetornoSiscota(pAntNosNumero) Then
                            lInsert = lInsert & "0, "
                        Else
                            lInsert = lInsert & gfForValor(1, flTabela.Value) & ", "
                        End If
                    Case "DAT_DIGIT"
                        lInsert = lInsert & Left(Me.gfForData(3), 8) & ", "                        'Dat_Digit
                    Case "COD_USER_LIBER"
                        lInsert = lInsert & "'COPIA',"                                  'Cod_User_Alt.
                    Case "NUM_BLOQ"
                        If pReentrada = 1 Then
                            lInsert = lInsert & flTabela.Value & ", "
                        Else
                            lInsert = lInsert & "0,"                                        'NUM_BLOQ
                        End If
                    Case "PER_MAX_COMIS_CORR"
                                                lInsert = lInsert & Per_Max_Comis_Corr & ","
                    Case "Dat_Ini_Vig_Orig"
                                                lInsert = lInsert & Dat_Ini_Vig_Orig & ","
                    Case "Dat_Ter_Vig_Orig"
                                                lInsert = lInsert & Dat_Ter_Vig_Orig & ","
                    Case "NUM_PI"
                        If pNovPI = 0 Then
                            lInsert = lInsert & Right(Year(Date), 2) & Format(pNovNumPedido, "00000000") & ", "
                        Else
                            lInsert = lInsert & pNovPI & ", "
                        End If
                    Case "DAT_REC_CIA"
                        If pDatRecCia = 0 Then
                            lInsert = lInsert & flTabela.Value & ", "
                        Else
                            lInsert = lInsert & pDatRecCia & ", "
                        End If
                    Case "NUM_APOL", "NUM_ENDO", "DAT_EMIS", "COD_IDEN_STAT_ACAO", "BONUSALTERADO"
                        lInsert = lInsert & "0, "
                    Case "COD_PRODUTO"
                        If lIdeTabela = "TAB_PED" Then
                            If lCodProduto <> 0 Then
                                lInsert = lInsert & lCodProduto & ", "
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Else
                            lInsert = lInsert & flTabela.Value & ", "
                        End If
                    Case "COD_IDEN_TARIF"
                        If lCodProduto <> 0 Then
                            lInsert = lInsert & lCodProduto & ", "
                        Else
                            lInsert = lInsert & flTabela.Value & ", "
                        End If
                    Case "COD_SEQUENCIA"
                        If lCodSequencia <> 0 Then
                            lInsert = lInsert & lCodSequencia & ", "
                        Else
                            lInsert = lInsert & flTabela.Value & ", "
                        End If
                    Case "LMI_UNICO"
                        If IsNull(flTabela.Value) Then
                            lInsert = lInsert & "0, "
                        Else
                            lInsert = lInsert & flTabela.Value & ", "
                        End If
                        pCalcLMI = gfTratarNulo(flTabela.Value, tpoNumerico)
                    Case "NUM_COTAC"    '1100090 - Inspecao Desabilitada
                        lInsert = lInsert & "'000000000000',"                               'NUM_COTAC.
                    
                    Case "SUPRIMIR_BATENTE"
                        If flTabela.Value Then
                            lInsert = lInsert & "1, "
                        Else
                            lInsert = lInsert & "0, "
                        End If
                    'Sustentacao 13/07/2017 - CRQ - CRQ000000607435
                    Case "PER_DESC_COML"
                        If gEhRetornoSiscota(pAntNosNumero) Then
                            lInsert = lInsert & "0, "
                        Else
                            lInsert = lInsert & Replace(flTabela.Value, ",", ".") & ", "
                        End If
                        
                    Case "VAL_VERBACOM"
                        lInsert = lInsert & "0, "
                    
                    Case "VAL_VERBACOMPROP"
                        lInsert = lInsert & "0, "
                                        
                    Case Else
                        Select Case flTabela.Type_
                        Case adBigInt, adBinary, adInteger, adLongVarBinary, adSmallInt, _
                             adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, _
                             adUnsignedTinyInt                          'Num�rico - inteiro.
                            If IsNull(flTabela.Value) Then
                                lInsert = lInsert & "0, "
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case adCurrency, adSingle, adDouble, adNumeric             'Num�rico - com v�rgula.
                            If IsNull(flTabela.Value) Then
                                lInsert = lInsert & "0, "
                            Else
                                If (flTabela.Name = "COD_ESTIP" Or flTabela.Name = "TAB_COBERT") And lCodProduto >= 1080 And M46V111.gCod_plano < 2000 Then
                                    lInsert = lInsert & gfForValor(1, gfPlanoBasico(lCodRamo, flTabela.Value)) & ", "
                                Else
                                    If flTabela.Name = "COD_COBERT" And lCodProduto >= 1080 And M46V111.gCod_plano < 2000 Then
                                        lInsert = lInsert & gfForValor(1, gfBuscarDePara(flTabela.Value, 1)) & ", "
                                    Else
                                        If flTabela.Name = "TAB_COBERT" And lCodProduto >= 1080 And M46V111.gCod_plano < 2000 Then
                                            lInsert = lInsert & gfForValor(1, gfBuscarDePara(flTabela.Value, 3)) & ", "
                                        Else
                                            lInsert = lInsert & gfForValor(1, flTabela.Value) & ", "
                                        End If
                                    End If
                                End If
                            End If
                        Case adChar, adVarChar, adLongVarChar                          'Texto.
                            If gfPreenchido(flTabela.Value) Then
                                If flTabela.Name = "COD_CLASF" And lCodProduto >= 1080 And M46V111.gCod_plano < 2000 Then
                                    lInsert = lInsert & "'" & Me.gfForAspas(gfBuscarDePara(flTabela.Value, 2, lCodClasfCompl)) & "', "
                                Else
                                    If flTabela.Name = "COD_COMPL_CLASF" And lCodProduto >= 1080 And M46V111.gCod_plano < 2000 Then
                                        lInsert = lInsert & "'" & Me.gfForAspas(gfBuscarDePara(lCodClasf, 2, lCodClasfCompl, True)) & "', "
                                    Else
                                lInsert = lInsert & "'" & Me.gfForAspas(flTabela.Value) & "', "
                                    End If
                                End If
                            Else
                                lInsert = lInsert & "' ', "
                            End If
                        
                        Case adBoolean                                  'Sim/N�o.
                            If IsNull(flTabela.Value) Then
                                lInsert = lInsert & "TRUE, "
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case Else
                            Call Me.gpFecha2(bdBasDados)
                            pMensagem = "FTE0008 - Tabela " & lIdeTabela & " com tipo do campo " & _
                                        flTabela.Name & " diferente dos previstos."
                            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                            GoTo SAIDA
                        End Select
                    End Select
Proximo:
                Next flTabela
                If Mid(lInsert, Len(lInsert), 1) = "," Then
                    lInsert = Left(lInsert, Len(lInsert) - 1) & ")"
                Else
                    lInsert = Left(lInsert, Len(lInsert) - 2) & ")"
                End If
                If gfExeSQL(bdBasDados, lInsert, pMensagem) <> 0 Then
                    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If
                rsTabela.MoveNext
            Wend
        End If
ProximaTabela:
    Next lCount

    gfCopProposta = True
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

    '' SiscotaEnd - Copiar Observacoes
    Call gfCopPropostaSiscotaEnd(pAntNosNumero, pNovNosNumero)

SAIDA:
    GoSub FechaObj
    Exit Function

ControleErro:
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfCopProposta ERRO - " & Err & " " & Err.Description)
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfCopProposta ERRO - " & Err & " " & Err.Description, 2, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Call Me.gpFecha2(bdBasDados)
    Set bdBasDados = Nothing
    Set flTabela = Nothing
    Set rsTabela = Nothing
    Return

End Function

Public Function gfCopPropostaSiscotaEnd(ByVal pAntNosNumero As String, ByVal pNovNosNumero As String) As Boolean
    Dim lConn As ADODB.Connection
    Dim lRsQuery As clsYasRecordSet
    Dim lQuery As String
    Dim pMensagem As String

    gfCopPropostaSiscotaEnd = False

    On Error GoTo Erro

    If gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lConn, pMensagem) = False Then
        GoTo Erro
    End If

    lQuery = "exec RamosDiversos..proc_CopiarSiscotaObservacao " & gfFormatarCamposSQL(pAntNosNumero) & ", " & gfFormatarCamposSQL(pNovNosNumero)

    If gfExeSQL(lConn, lQuery, pMensagem) = False Then
        GoTo Erro
    End If

    gfCopPropostaSiscotaEnd = True
Fim:
    Set lRsQuery = Nothing
    Call Me.gpFecha2(lConn)
    Exit Function

Erro:
    If Err.Number > 0 Then
        pMensagem = pMensagem & "ERR " & Err.Number & " DESC " & Err.Description
    End If
    GoSub Fim
End Function

Public Function gfCopPropostaClasse4(ByVal pNossoNumeroAnt As String, _
                                     ByVal pNumPedAnt As String, _
                                     ByVal pNumCotacAnt As String, _
                                     ByVal pNovNosNumero As String, _
                                     ByVal pNovNumPedido As String, _
                                     ByRef pMensagem As String, _
                                     Optional ByVal pNovPI As Long, _
                                     Optional pDatRecCia As Long, _
                                     Optional pCodRamo As Integer, _
                                     Optional ByRef pCalcLMI As Integer, _
                                     Optional pNaoRepProduto As Boolean, _
                                     Optional pReentrada As Integer, _
                                     Optional pCodProduto As Integer = 0) As Boolean
                                      
    'Fun��o: monta comando INSERT gen�rico para dados de proposta.
    
    'Par�metros de entrada: pNossoNumeroAnt = Nosso N�mero da Cota��o anterior a nova gerada.
                           'pNumPedAnt      = Refer�ncia da Cota��o anterior a nova gerada.
                           'pNumCotacAnt    = N�mero de SISCOTA gerada na cota��o anterior a nova gerada.
                           'pNovNosNumero   = Nosso N�mero da nova cota��o gerada
                           'pNovNumPedido   = Refer�ncia da nova cota��o gerada
    
    'Par�metro de sa�da: pMensagem = Mensagem de retorno caso ocorra erro.

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "gfCopPropostaClasse4"
    lKeyDynaTrace = DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim bdBasDados As ADODB.Connection    'Base de dados.
    Dim flTabela As clsYasField           'Campos da tabela.
    Dim rsTabela As clsYasRecordSet       'Registro da tabela.
    Dim lCount As Long                    'Utilizado no for...next.
    Dim lIdeTabela As String              'Nome da tabela.
    Dim lContinuar As Boolean             'Indicador se continua o procedimento ou n�o.
    Dim lSelect As String                 'String de instru��o SQL para selecionar dados no banco
    Dim lInsert As String                 'String de instru��o SQL para inserir dados no banco
    Dim lPropTabela As String             '
    Dim lTipEmissao As Integer            'Indica o tipo de emiss�o
    Dim lCodTarifaCopia As Integer
    Dim lCodProduto As Integer            'C�digo do produto.
    Dim lCodProtocolo As Long             'C�digo do protocolo.
    Dim lCodSequencia As Long             'C�digo da seq��ncia.
    Dim lDesTarifa As String              'Descri��o da tarifa.
    Dim lMensagem As String               'Mensagem
    Dim lCodRamo As Integer
    Dim lCarregarDePara As Boolean
    Dim lCodClasf As String
    Dim lCodClasfCompl As String
    
    On Error GoTo ControleErro

    gfCopPropostaClasse4 = False

    Call gfObtSYASMULTINI(lPropTabela)
    
    Call gTipEmissao(pNossoNumeroAnt, lTipEmissao, lCodTarifaCopia, lPropTabela)

    If lCodTarifaCopia < TarifaPrecificacaoMaritima And (lTipEmissao = e_TipEmissao_Endosso Or pReentrada = 1) Then
        lCodProduto = lCodTarifaCopia
    End If
    
    'Reposicionar na �ltima tarifa
    If gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        If pNaoRepProduto = False Then
            If Me.gfBusProduto(gHoje, lCodProduto, lCodProtocolo, lCodSequencia, lDesTarifa, lMensagem, pCodRamo) = False Then
                Call gpGraLog(1, lMensagem)
                Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If
        End If
    End If

    If lCodProduto = 0 Then
        lCodProduto = pCodProduto
    End If
    
    '1. Obt�m registros da tabela.
    If gfAbrBasDados("P00MPED", gTipP00MPED, gEndP00MPED, bdBasDados, pMensagem) = False Then
        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        GoTo SAIDA
    End If

    For lCount = 1 To 33
        lContinuar = True
        Select Case lCount
        Case 1
            lIdeTabela = "TAB_PED"
        Case 2
            lIdeTabela = "TAB_PED_CORR"
        Case 3
            lIdeTabela = "TAB_PED_COSS"
        Case 4
            lIdeTabela = "TAB_PED_JUROS"
        Case 5
            lIdeTabela = "TAB_PED_VENC"
        Case 6
            lIdeTabela = "TAB_PED_ENDER"
        Case 7
            lIdeTabela = "TAB_CTRL_EMIS"
        Case 8
            lIdeTabela = "TAB_PED_LOC"
        Case 9
            lIdeTabela = "TAB_PED_CLAU"
        Case 10
            lIdeTabela = "TAB_PED_COBERT"
        Case 11
            lIdeTabela = "TAB_PED_LOC_BEM"
        Case 12
            lIdeTabela = "TAB_PED_LOC_SEGURO"
        Case 13
            lIdeTabela = "TAB_PED_OBS"
        Case 14
            lIdeTabela = "TAB_PED_INSP"
        Case 15
            lIdeTabela = "TAB_PED_AUTO"
        Case 16
            lIdeTabela = "TAB_PED_COBERT_COMPL"
        Case 17
            lIdeTabela = "TAB_PED_FAT_RISCO"
        Case 18
            lIdeTabela = "TAB_PED_DIZER"
        Case 19
            lIdeTabela = "TAB_PED_ESPEC"
        Case 20
            lIdeTabela = "TAB_PED_ESPEC2"
        Case 21
            lIdeTabela = "Tab_Ped_Loc_LMI"
        Case 22
            lIdeTabela = "Tab_Canal_Emiss"
        Case 23
            lIdeTabela = "TAB_CTRL_EMIS_COMPL"
        Case 24
            lIdeTabela = "TAB_PED_LOC_PESSOAS"
        Case 25
            lIdeTabela = "TAB_PED_TELEFONES"
        Case 26
            lIdeTabela = "TAB_PED_LOC_COMPL"
        Case 27
            lIdeTabela = "TAB_PED_COBERT_FRANQUIA"
        Case 28
            lIdeTabela = "TAB_PED_TEXTOS"
        Case 29
            lIdeTabela = "TAB_PED_INSP_COMPL"
        Case 30
            lIdeTabela = "TAB_PED_LOC_BICICLETA"
        Case 31
            lIdeTabela = "TAB_PED_LIBERACAO_SERASA"
        Case 32
            lIdeTabela = "TAB_PED_ALL_RISKS"
        Case 33
            lIdeTabela = "TAB_PED_OBRAS_ARTE"
        End Select

        lSelect = "SELECT * FROM " & lPropTabela & lIdeTabela & " WITH (NOLOCK) WHERE Nosso_Numero = '" & pNossoNumeroAnt & "'"
        If gfObtRegistro(bdBasDados, lSelect, rsTabela, pMensagem) = False Then
            Call Me.gpFecha2(bdBasDados)
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        
        If rsTabela.EOF = True Then
            lContinuar = False
            If lIdeTabela = "TAB_PED" Then
                pMensagem = "FNL0003 - " & gPrefixo & "Pedido n�o localizado (" & pNossoNumeroAnt & _
                          " - c�pia da cota��o/proposta classe 4.)."
                Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                GoTo SAIDA
            End If
        End If
        
        If lContinuar = True Then
            '2. Grava nova cota��o/proposta.
            While Not rsTabela.EOF
                lCodRamo = 0
                lCodClasf = 0
                lCodClasfCompl = 0

                If lIdeTabela = "TAB_PED" Or lIdeTabela = "TAB_PED_COBERT" Or lIdeTabela = "TAB_PED_COBERT_COMPL" Then
                    lCodRamo = rsTabela!Cod_Ramo
                    If lIdeTabela = "TAB_PED" Then
                        M46V111.gCod_plano = rsTabela!Cod_Estip
                        lCarregarDePara = False
                        
                        lCarregarDePara = lCodProduto >= 1080 And gRst_Tbm_De_Para.State <> adStateOpen
                        If lCodProduto >= 1080 And Not lCarregarDePara Then
                            lCarregarDePara = gRst_Tbm_De_Para.RecordCount = 0
                        End If
                    
                        If lCarregarDePara Then
                            gCodProduto = lCodProduto
                            gCod_Ramo = lCodRamo
                            M46V111.gCod_plano = gfPlanoBasico(lCodRamo, rsTabela!Cod_Estip)
                            Call mfCarregarDePara
                        End If
                        
                    End If
                End If
                
                If lIdeTabela = "TAB_PED_LOC" And lCodProduto >= 1080 Then
                    lCodClasf = rsTabela!Cod_Clasf
                    lCodClasfCompl = rsTabela!Cod_Compl_Clasf
                End If
                
                lInsert = "INSERT INTO " & lPropTabela & lIdeTabela & " VALUES ("
                
                For Each flTabela In rsTabela.Fields
                    If Me.gOrigemChamada = e_Origem_Web And (UCase(flTabela.Name) = "IND_STATUSINSPECAO" Or UCase(flTabela.Name) = "NUM_INSP") Then
                        lInsert = lInsert & "0, "
                        GoTo Proximo
                    End If
                    
                    Select Case UCase(flTabela.Name)
                        Case "PED_LOC_PESSOA_ID", "TELEFONE_ID", "ID_FRANQUIA_AVANCADA"
                        Case "NOSSO_NUMERO"
                            lInsert = lInsert & "'" & pNovNosNumero & "', "                                         'Nosso n�mero
                        Case "NUM_PED"
                            lInsert = lInsert & pNovNumPedido & ", "                                                'N�mero do pedido.
                        Case "SIT_DOCUMENTO", "COD_SIT_PED"
                            lInsert = lInsert & "1, "                                                               'Sit_Documento.
                        Case "NUM_BLOQ"
                            If pReentrada = 1 Then
                                lInsert = lInsert & flTabela.Value & ", "
                            Else
                                lInsert = lInsert & "0,"                                                            'NUM_BLOQ
                            End If
                        Case "DAT_PRI_CALCULO", "IND_PROMOCIONAL", "IND_TEXTO_REL_BENS", "IND_BLOQ_WEB", "IND_ENVIO_ALCADA"
                            lInsert = lInsert & "0, "                                               'Dat_Pri_Calculo.
                        Case "VAL_PR_ORIG", "PREMIOINFORMADO"
                            lInsert = lInsert & gfForValor(1, flTabela.Value) & ", "
                        Case "DAT_DIGIT"
                            lInsert = lInsert & Left(Me.gfForData(3), 8) & ", "                                     'Dat_Digit
                        Case "COD_USER_LIBER"
                            lInsert = lInsert & "'COPIA',"                                                          'Cod_User_Alt.
                        Case "NUM_PI"
                            If pNovPI = 0 Then
                                lInsert = lInsert & Right(Year(Date), 2) & Format(pNovNumPedido, "00000000") & ", "
                            Else
                                lInsert = lInsert & pNovPI & ", "
                            End If
                        Case "NUM_APOL", "NUM_ENDO", "DAT_EMIS", "COD_IDEN_STAT_ACAO"
                            lInsert = lInsert & "0, "
                        Case "BONUSALTERADO"
                            If flTabela.Value = True Then
                                lInsert = lInsert & "1, "
                            Else
                                lInsert = lInsert & "0, "
                            End If
                        Case "DAT_REC_CIA"
                            If pDatRecCia = 0 Then
                                lInsert = lInsert & flTabela.Value & ", "
                            Else
                                lInsert = lInsert & pDatRecCia & ", "
                            End If
                        Case "COD_PRODUTO"
                            If lIdeTabela = "TAB_PED" Then
                                If lCodProduto <> 0 Then
                                    lInsert = lInsert & lCodProduto & ", "
                                Else
                                    lInsert = lInsert & flTabela.Value & ", "
                                End If
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case "COD_IDEN_TARIF"
                            If lCodProduto <> 0 Then
                                lInsert = lInsert & lCodProduto & ", "
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case "COD_SEQUENCIA"
                            If lCodSequencia <> 0 Then
                                lInsert = lInsert & lCodSequencia & ", "
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case "LMI_UNICO"
                            If IsNull(flTabela.Value) Then
                                lInsert = lInsert & "0, "
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                            pCalcLMI = gfTratarNulo(flTabela.Value, tpoNumerico)
                        Case "NUM_COTAC"
                            lInsert = lInsert & "'000000000000',"                                                   'NUM_COTAC.
                        Case "SUPRIMIR_BATENTE"
                            If flTabela.Value Then
                                lInsert = lInsert & "1, "
                            Else
                                lInsert = lInsert & "0, "
                            End If
                        Case "NUM_COTAC_ANT"
                            lInsert = lInsert & "'" & pNumCotacAnt & "', "                                          'NUM_COTAC_ANT.
                        Case "NUM_PED_ANT"
                            lInsert = lInsert & "'" & pNumPedAnt & "', "                                            'NUM_PED_ANT.
                        Case "TIP_ORIGEM_C4"
                            lInsert = lInsert & "0, "
                        Case "ITEM_ORIGEM_C4"
                            lInsert = lInsert & "1, "
                        Case "RA_C4"
                            lInsert = lInsert & "0, "
                        Case "RA_Comissao"
                            lInsert = lInsert & "0, "
                        Case "Per_Comis_Corr_RA"
                            lInsert = lInsert & "0, "
                        Case "VAL_IS_RA"
                            lInsert = lInsert & "0, "
                    Case Else
                        Select Case flTabela.Type_
                            Case adBigInt, adBinary, adInteger, adLongVarBinary, adSmallInt, _
                                 adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, _
                                 adUnsignedTinyInt                          'Num�rico - inteiro.
                                If IsNull(flTabela.Value) Then
                                    lInsert = lInsert & "0, "
                                Else
                                    lInsert = lInsert & flTabela.Value & ", "
                                End If
                            Case adCurrency, adSingle, adDouble, adNumeric             'Num�rico - com v�rgula.
                                If IsNull(flTabela.Value) Then
                                    lInsert = lInsert & "0, "
                                Else
                                    If (flTabela.Name = "COD_ESTIP" Or flTabela.Name = "TAB_COBERT") And lCodProduto >= 1080 And M46V111.gCod_plano < 2000 Then
                                        lInsert = lInsert & gfForValor(1, gfPlanoBasico(lCodRamo, flTabela.Value)) & ", "
                                    Else
                                        If flTabela.Name = "COD_COBERT" And lCodProduto >= 1080 And M46V111.gCod_plano < 2000 Then
                                            lInsert = lInsert & gfForValor(1, gfBuscarDePara(flTabela.Value, 1)) & ", "
                                        Else
                                            If flTabela.Name = "TAB_COBERT" And lCodProduto >= 1080 And M46V111.gCod_plano < 2000 Then
                                                lInsert = lInsert & gfForValor(1, gfBuscarDePara(flTabela.Value, 3)) & ", "
                                            Else
                                                lInsert = lInsert & gfForValor(1, flTabela.Value) & ", "
                                            End If
                                        End If
                                    End If
                                End If
                            Case adChar, adVarChar, adLongVarChar                          'Texto.
                                If gfPreenchido(flTabela.Value) Then
                                    If flTabela.Name = "COD_CLASF" And lCodProduto >= 1080 And M46V111.gCod_plano < 2000 Then
                                        lInsert = lInsert & "'" & Me.gfForAspas(gfBuscarDePara(flTabela.Value, 2, lCodClasfCompl)) & "', "
                                    Else
                                        If flTabela.Name = "COD_COMPL_CLASF" And lCodProduto >= 1080 And M46V111.gCod_plano < 2000 Then
                                            lInsert = lInsert & "'" & Me.gfForAspas(gfBuscarDePara(lCodClasf, 2, lCodClasfCompl, True)) & "', "
                                        Else
                                    lInsert = lInsert & "'" & Me.gfForAspas(flTabela.Value) & "', "
                                        End If
                                    End If
                                Else
                                    lInsert = lInsert & "' ', "
                                End If
                            
                            Case adBoolean                                  'Sim/N�o.
                                If IsNull(flTabela.Value) Then
                                    lInsert = lInsert & "TRUE, "
                                Else
                                    lInsert = lInsert & flTabela.Value & ", "
                                End If
                            Case Else
                                Call Me.gpFecha2(bdBasDados)
                                pMensagem = "FTE0008 - Tabela " & lIdeTabela & " com tipo do campo " & _
                                            flTabela.Name & " diferente dos previstos."
                                Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                                GoTo SAIDA
                        End Select
                    End Select
Proximo:
                Next flTabela
                If Mid(lInsert, Len(lInsert), 1) = "," Then
                    lInsert = Left(lInsert, Len(lInsert) - 1) & ")"
                Else
                    lInsert = Left(lInsert, Len(lInsert) - 2) & ")"
                End If
                If gfExeSQL(bdBasDados, lInsert, pMensagem) <> 0 Then
                    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If
                rsTabela.MoveNext
            Wend
        End If
        
ProximaTabela:

    Next lCount

    gfCopPropostaClasse4 = True
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    GoSub FechaObj
    Exit Function

ControleErro:
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfCopPropostaClasse4 ERRO - " & Err & " " & Err.Description)
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfCopPropostaClasse4 ERRO - " & Err & " " & Err.Description, 2, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Call Me.gpFecha2(bdBasDados)
    Set bdBasDados = Nothing
    Set flTabela = Nothing
    Set rsTabela = Nothing
    Return

End Function

Public Sub gpGraLog(ByVal pOpcao As Byte, ByVal pMensagem As String, Optional pMsgBoxSyle As VbMsgBoxStyle = vbApplicationModal)
'Procedure: grava arquivo log.

'Par�metros de entrada...pOpcao......0 = grava log.
'                                    1 = grava log e emite mensagem.
'                                    2 = grava log e emite mensagem com frase inicial
'                                        "Comunique Central de Atendimento".
'                        pMensagem...Mensagem.

    Dim lMensagem As String    'Mensagem.
    Dim lNumLog As Integer             'N�mero do arquivo log.
    Dim lPosicao As Integer    'Utilizado com InStr.
    Dim lMsgErr As String             'Mensagem de erro dentro desta procedure
    Dim lCkPoint As String    'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lArqAberto As Boolean

    Static lEndLog As String    'Endere�o do arquivo log.
    Static lNomComputador As String    'Nome do computador.
    Static lNomPrograma As String    'Nome do programa.
    Static lNomUsuario As String    'Nome do usu�rio.
    Static lNumMensagem As Integer    'N�mero de mensagem.
    Static lPriVez As Boolean    'Indicador de primeira vez (false).

    Err = 0
    
    If (gProibidoMsgBox = True Or pOpcao = 0) And gPrograma = "D046O003" And gEnableFileLog = False Then
        Exit Sub
    End If
    On Error GoTo TRATAMENTO_ERRO

    '1. Verifica op��o.
    lCkPoint = "1"
    If Not gfPreenchido(pOpcao) Then
        MsgBox "FTE0006 - Op��o inv�lida (grava log - n�o preenchida)."
        GoTo SAIDA
    End If
    lCkPoint = "2"
    If Not IsNumeric(pOpcao) Then
        MsgBox "FTE0006 - Op��o inv�lida (grava log - n�o num�rica)."
        GoTo SAIDA
    End If
    lCkPoint = "3"
    If pOpcao <> 0 And pOpcao <> 1 And pOpcao <> 2 Then
        MsgBox pMensagem
        MsgBox "FTE0006 - Op��o inv�lida (grava log - diferente de 0, 1 e 2 - " & pOpcao & ")."
        GoTo SAIDA
    End If
    lCkPoint = "4"
    If gVerDataMaq = 1 Then
        lCkPoint = "5"
        If gfIncDatas = False Then
            'Travar o sistema
            gTrvSistema = True
        End If
    End If
    lCkPoint = "6    "
    '2. Ao executar esta procedure pela primeira vez, obter dados utilizados na grava��o das mensagens.
    If lPriVez = False Then
        lCkPoint = "7"
        lPriVez = True
        'a) Verifica arquivo log.
        lCkPoint = "8"
        If gfPreenchido(gNomArqLog) = False Then
            gNomArqLog = "SYASMULT_" & Format(Date, "DD_MM_YYYY") & ".LOG"
        End If
        lCkPoint = "9"
        lEndLog = gAppPath & gNomArqLog
        'b) Obt�m nome do usu�rio.
        lCkPoint = "10"
        If gfPreenchido(gNomUsuario) Then
            lCkPoint = "11"
            lNomUsuario = gNomUsuario
            While Len(lNomUsuario) < 20
                lCkPoint = "12"
                lNomUsuario = lNomUsuario & " "
            Wend
            lCkPoint = "13"
        Else
            lCkPoint = "14"
            lNomUsuario = String(20, "X")
        End If
        lCkPoint = "15"
        'c) Obt�m nome do computador.
        lNomComputador = M99V001.GetLoginComputerName()
        lCkPoint = "16"
        If Not gfPreenchido(lNomComputador) Then
            lNomComputador = "********"
        End If
        lCkPoint = "17"
        While Len(lNomComputador) < 16
            lCkPoint = "18"
            lNomComputador = lNomComputador & " "
        Wend
        lCkPoint = "19"
        'd) Obt�m nome do programa.
        lNomPrograma = Trim$(App.EXEName)
        lCkPoint = "20"
        While Len(lNomPrograma) < 12
            lCkPoint = "21"
            lNomPrograma = lNomPrograma & " "
        Wend
        lCkPoint = "22"
    End If
    lCkPoint = "23"

    '3. Consiste par�metro de entrada (Mensagem).
    lCkPoint = "24"
    If Not gfPreenchido(pMensagem) Then
        lCkPoint = "25"
        pMensagem = "FTE0045 - Mensagem n�o preenchida."
    End If
    lCkPoint = "26"
    '4. Ajusta a mensagem
    lMensagem = pMensagem
    Do
        lCkPoint = "27"
        lPosicao = InStr(lMensagem, vbLf)
        lCkPoint = "28"
        If lPosicao = 0 Then
            Exit Do
        End If
        lCkPoint = "29"
        lMensagem = Left$(lMensagem, lPosicao - 1) & " " & Mid$(lMensagem, lPosicao + 1)
    Loop

    '3. Abre arquivo log.
    lCkPoint = "30"
    lNumLog = FreeFile
    lCkPoint = "31"
    lArqAberto = False
    lCkPoint = "32"
    Open lEndLog For Append Shared As #lNumLog
    lCkPoint = "33"
    lArqAberto = True

    'x. Posiciona n�mero da mensagem.
    lCkPoint = "34"
    If lNumMensagem = 9999 Then
        lCkPoint = "35"
        lNumMensagem = 1
    Else
        lCkPoint = "36"
        lNumMensagem = lNumMensagem + 1
    End If
    lCkPoint = "37"

    'Print #lNumLog, lMensagem
    'x. Grava mensagem.
    Print #lNumLog, gForHoje & " | " & Format$(Now, "hh:nn:ss") & " | " & lNomUsuario & " | " & _
                    lNomComputador & " | " & lNomPrograma & " | " & Format$(sUniqueInstanceCode, "0000") & " | " & _
                    Format$(lNumMensagem, "0000") & _
                  " | " & lMensagem
    lCkPoint = "38"
FecharArquivo:
    lCkPoint = "36"
    Close #lNumLog
    lCkPoint = "40"
    lArqAberto = False      'marca que o arquivo n�o est� mais aberto
    lCkPoint = "41"

    '6. Emite mensagem ao usu�rio.
    lCkPoint = "42"
    If gProibidoMsgBox = True Then
        pOpcao = 0
    End If
    Select Case pOpcao
    Case 0
        lCkPoint = "43"
        'Nada a fazer.
    Case 1
        lCkPoint = "44"
        MsgBox pMensagem, pMsgBoxSyle
    Case 2
        lCkPoint = "45"
        MsgBox "Comunique Central de Atendimento. " & vbLf & vbLf & pMensagem & vbLf & vbLf & "Data/Hora da ocorr�ncia : " & gForHoje & " " & Format$(Now, "Hh:Nn:Ss"), pMsgBoxSyle
    End Select
    lCkPoint = "46"

SAIDA:
    If lArqAberto = True Then
        'Fecha o arquivo caso o c�digo tenha sido desviado para o erro
        'ap�s ter aberto o arquivo.
        Close #lNumLog
        lArqAberto = False
    End If
    Err = 0
    On Error GoTo 0
    Exit Sub

TRATAMENTO_ERRO:
    If gfPreenchido(lMsgErr) = False Then
        lMsgErr = Err & " - " & ERROR
    End If
    lMsgErr = "gpGraLog Erro " & lMsgErr
    lMsgErr = lMsgErr & vbCrLf & "CheckPoint " & lCkPoint

    lMensagem = "Comunique Central de Atendimento. " & vbLf & vbLf & lMensagem & vbLf & vbLf & "Data/Hora da ocorr�ncia : " & gForHoje & " " & Format$(Now, "Hh:Nn:Ss")
    lMsgErr = lMsgErr & vbCrLf & vbCrLf & gForHoje & " | " & Format$(Now, "hh:nn:ss") & " | " & lNomUsuario & " | " & _
              lNomComputador & " | " & lNomPrograma & " | " & Format$(lNumMensagem, "0000") & _
            " | " & lMensagem
    If pOpcao <> 0 Then
        MsgBox lMsgErr
    End If

    GoTo SAIDA
  
End Sub
Public Function gfPreenchido(ByVal pCampo) As Boolean
'Fun��o: verifica se campo preenchido.

    If IsNull(pCampo) Then
        gfPreenchido = False
    Else
        If Len(Trim(pCampo)) = 0 Then
            gfPreenchido = False
        Else
            gfPreenchido = True
        End If
    End If
End Function

Public Function gfAbrBasDados(ByVal pIndBasDados As String, ByVal pTipBasDados As String, _
                              ByVal pEndBasDados As String, ByRef pBasDados As ADODB.Connection, _
                              ByRef pMensagem As String) As Boolean
'Fun��o: abre base de dados.

'Par�metros de entrada...pIndBasDados...Indicador da base de dados.
'                        pTipBasDados...Tipo da base de dados.
'                        pEndBasDados...Endere�o da base de dados.
'Par�metros de sa�da.....pBasDados......Base de dados.
'                        pMensagem......Mensagem.

    Dim lSqlServer As Boolean           'Indica se a base dever� ser aberta com SQL SERVER.

    gfAbrBasDados = False

    '1. Verifica indicador do arquivo.
    If Not gfPreenchido(pIndBasDados) Then
        pMensagem = "[ERR] - FTE0028 - Erro ao abrir base de dados (indicador n�o preenchido)."
        Exit Function
    End If

    lSqlServer = False
    If (UCase(pIndBasDados) = "SYAS_EMIS" And PerfilAcesso = 3) Or gPrograma = "D046O003" Then
        lSqlServer = True
    ElseIf UCase(pIndBasDados) = "SYAS_EMIS" And UCase(App.EXEName) Like "P46V600*" Then
        lSqlServer = True
    Else
        If gBaseLocal Then
            Select Case UCase$(pIndBasDados)
            Case "P00MPED"
                'If gTipoAcesso = DIGITACAO Or PerfilAcesso = e_Perfil_Operacao Then
                '    lSqlServer = True
                'Else
                lSqlServer = False
                'End If
            Case "P00APOL", "P00MPED", "P00MULTI", "P0044700", "P0043000", _
                 "P00MULTI2", "P0045200", "P0045300", "P0045400"
                lSqlServer = False
            Case "GED", "YAS_ND"
                lSqlServer = True
            Case Else
                pMensagem = "[ERR] - FTE0029 - Erro ao abrir base de dados (indicador desconhecido)."
                Exit Function
            End Select
        Else
            Select Case UCase$(pIndBasDados)
            Case "P00MULTI2"
                lSqlServer = False
            Case "P00APOL", "P00MPED", "GED", "P00MULTI", "RAMOSDIVERSOS", "P0043000", "P0044700", "P0045400", "EAPDB", "SYASPORTAL", "YAS_ND"
                lSqlServer = True
            Case Else
                pMensagem = "[ERR] - FTE0029 - Erro ao abrir base de dados (indicador desconhecido)."
                Exit Function
            End Select
        End If
    End If
    If lSqlServer = True Then
        If UCase$(pIndBasDados) = "GED" Then
            gfAbrBasDados = gfAbrBasSQL("GED", "2", gSerGed, gEndGed, gUidGed, gPwdGed, pBasDados, pMensagem)
        ElseIf UCase$(pIndBasDados) = "EAPDB" Then
            gfAbrBasDados = gfAbrBasSQL("EAPDB", "2", gSerEAPDB, gEndEAPDB, gUidEAPDB, gPwdEAPDB, pBasDados, pMensagem)
        ElseIf UCase$(pIndBasDados) = "SYASPORTAL" Then
            gfAbrBasDados = gfAbrBasSQL("SYASPORTAL", "2", gSerSyasPortal, gEndSyasPortal, gUidSyasPortal, gPwdSyasPortal, pBasDados, pMensagem)
'        ElseIf UCase$(pIndBasDados) = "P0045400" Then
'            gfAbrBasDados = gfAbrBasSQL("P0045400", "2", gSerSYASEMIS, gEndSYASEMIS, gUidSYASEMIS, gPwdSYASEMIS, pBasDados, pMensagem)
        ElseIf UCase$(pIndBasDados) = "YAS_ND" Then
            gfAbrBasDados = gfAbrBasSQL("YAS_ND", "2", gSerYAS_ND, gEndYAS_ND, gUidYAS_ND, gPwdYAS_ND, pBasDados, pMensagem)
        ElseIf UCase$(pIndBasDados) = "SYAS_EMIS" Then
            gfAbrBasDados = gfAbrBasSQL("SYAS_EMIS", "2", gSerSYASEMIS, gEndSYASEMIS, gUidSYASEMIS, gPwdSYASEMIS, pBasDados, pMensagem)
        Else
            gfAbrBasDados = gfAbrBasSQL("DBDIVERSOS", "2", gSerDIVERSOS, gEndDIVERSOS, gUidDIVERSOS, _
                                        gPwdDIVERSOS, pBasDados, pMensagem)
        End If
    Else
        'pMensagem = "[ERR] - FTE0129 - Tentou abrir uma base de dados Access antiga."
        'Exit Function
        gfAbrBasDados = gfAbrBasDadosAccess(pIndBasDados, pTipBasDados, pEndBasDados, pBasDados, _
                                            pMensagem)
    End If
    On Error GoTo 0
End Function


Public Function gfAbrBasDadosAccess(ByVal pIndBasDados As String, ByVal pTipBasDados As String, _
                                    ByVal pEndBasDados As String, ByRef pBasDados As ADODB.Connection, _
                                    ByRef pMensagem As String) As Boolean
'Fun��o: abre base de dados Access.

'Par�metros de entrada...pIndBasDados...Indicador da base de dados.
'                        pTipBasDados...Tipo da base de dados.
'                        pEndBasDados...Endere�o da base de dados.
'Par�metros de sa�da.....pBasDados......Base de dados.
'                        pMensagem......Mensagem.

    Dim lAbriu As Boolean    'Abriu a base de dados?
    Dim lComando As String    'Comando.
    Dim lDriver(6) As String    'Drivers.
    Dim lErrDriver(6) As String    'Grava mensagem de erro obtida em cada driver
    Dim lI As Byte             'Utilizado com For...Next.
    Dim lIniCatalog As String    'Par�metro INITIAL CATALOG.
    Dim lPwd As String    'Par�metro PWD.
    Dim lServer As String    'Par�metro SERVER.
    Dim lTemErro As Boolean    'Tem erro?
    Dim lUid As String    'Par�metro UID.
    Dim lMsgErro As String    'Mensagem de erro

    gfAbrBasDadosAccess = False

    If pIndBasDados = "P00MULTI2" Then
        pIndBasDados = "P00MULTI"
        pTipBasDados = 1
    End If


    '1. Verifica indicador do arquivo.
    If Not gfPreenchido(pIndBasDados) Then
        pMensagem = "[ERR] - FTE0028 - Erro ao abrir base de dados (indicador n�o preenchido)."
        Exit Function
    End If
    Select Case UCase$(pIndBasDados)
    Case "P00APOL", "P00MPED", "P00MULTI", "P0044700", "P0043000", _
         "P0045200", "P0045300", "P0045400"
        'Nada a fazer - arquivos v�lidos.
    Case Else
        pMensagem = "[ERR] - FTE0029 - Erro ao abrir base de dados (indicador desconhecido - " & _
                    pIndBasDados & ")."
        Exit Function
    End Select

    '2. Verifica tipo do arquivo.
    If Not gfPreenchido(pTipBasDados) Then
        pMensagem = "[ERR] - FTE0030 - Erro ao abrir base de dados (" & pIndBasDados & _
                  " - tipo da base de dados n�o preenchido)."
        Exit Function
    End If
    If pTipBasDados <> "1" Then
        pMensagem = "[ERR] - FTE0031 - Erro ao abrir base de dados (" & pIndBasDados & _
                  " - tipo da base de dados desconhecido: " & pTipBasDados & ")."
        Exit Function
    End If

    '3. Verifica endere�o do arquivo.
    If Not gfPreenchido(pEndBasDados) Then
        pMensagem = "[ERR] - FTE0032 - Erro ao abrir base de dados (" & pIndBasDados & _
                  " - endere�o da base de dados n�o preenchido)."
        Exit Function
    End If
    If Not gfPreenchido(Dir(pEndBasDados)) Then
        pMensagem = "[ERR] - FTE0033 - Erro ao abrir base de dados (" & pIndBasDados & _
                  " - endere�o da base de dados n�o localizado: " & pEndBasDados & ")."
        Exit Function
    End If

    '4. Abre arquivo.
    lDriver(1) = "Microsoft Access Driver (*.mdb)"
    lDriver(2) = "Microsoft Access Driver(*.mdb)"
    lDriver(3) = "Driver do Microsoft Access (*.mdb)"
    lDriver(4) = "Driver do Microsoft Access(*.mdb)"
    lDriver(5) = "Driver para o Microsoft Access (*.mdb)"
    lDriver(6) = "Driver para o Microsoft Access(*.mdb)"
    lTemErro = False
    lAbriu = False
    lMsgErro = ""
    For lI = 1 To 6
        'a) Posiciona conex�o.
        On Error Resume Next
        Set pBasDados = New ADODB.Connection
        If Err <> 0 Then
            pMensagem = "[ERR] - FTE0034 - Erro ao abrir base de dados (" & pIndBasDados & _
                      " - comando: SET New ADODB.Connection - Erro " & Err & " - " & ERROR & ")."
            lTemErro = True
            Exit For
        End If
        'b) Posiciona propriedade CursorLocation.
        pBasDados.CursorLocation = adUseClient
        If Err <> 0 Then
            pMensagem = "[ERR] - FTE0035 - Erro ao abrir base de dados (" & pIndBasDados & _
                      " - comando: CursorLocation = adUseClient - Erro " & Err & " - " & ERROR & ")."
            lTemErro = True
            Exit For
        End If
        'c) Abre arquivo.
        lComando = "Driver={" & lDriver(lI) & "};Dbq=" & pEndBasDados
        ''Call gpGraLog(0, "AGE0000 - gfAbrBasDadosAccess: lComando = " & lComando & " pIndBasDados=" & _
         pIndBasDados)

        Select Case pIndBasDados
        Case "P00APOL", "P00MPED", "P00MULTI"
            pBasDados.Open lComando, , LCase(pIndBasDados) & "46"
        Case "P0044700", "P0043000", "P0045200", "P0045300", "P0045400"
            pBasDados.Open lComando, , LCase(pIndBasDados) & "21"
        End Select
        If Err = 0 Then
            lAbriu = True
            Exit For
        Else
            lMsgErro = " - Erro " & Err & " - " & ERROR
            'o n�mero da msg � o mesmo, por isso estou consistindo pelo nome.
            If InStr(1, UCase(lMsgErro), "CORROMPIDO") > 0 Or InStr(1, UCase(lMsgErro), "DANIFICADO") > 0 Then
                lMsgErro = " - Arquivo corrompido ou danificado."
                Exit For
            ElseIf InStr(1, UCase(lMsgErro), "N�O � UM CAMINHO V�LIDO") > 0 Or InStr(1, UCase(lMsgErro), "ESCRITO CORRETAMENTE") > 0 Then
                lMsgErro = " - Problemas de comunica��o, verifique comunica��o com a rede."
                Exit For
            ElseIf InStr(1, UCase(lMsgErro), "MODO EXCLUSIVO") > 0 Or InStr(1, UCase(lMsgErro), "PERMISS�O") > 0 Then
                If Len(pEndBasDados) > 12 Then
                    lMsgErro = " - Privil�gios insuficientes no diret�rio " & Left(pEndBasDados, Len(pEndBasDados) - 12)
                Else
                    lMsgErro = " - Privil�gios insuficientes no diret�rio " & gAppPath
                End If
                Exit For
            End If
            'Guarda a mensagem de erro n�o prevista
            lErrDriver(lI) = lMsgErro
        End If
        On Error GoTo 0
    Next lI
    If lTemErro Then
        On Error GoTo 0
        Exit Function
    End If
    If lAbriu = False Then
        pMensagem = "[ERR] - FTE0036 - Erro ao abrir base de dados (" & pIndBasDados & " - comando: " & _
                    lComando & IIf(lMsgErro = "", " - Erro " & Err & " - " & ERROR, lMsgErro) & ")."
        On Error GoTo 0
        For lI = 1 To 6
            'Concatena mensagens de erro diferentes
            If gfPreenchido(lErrDriver(lI)) = True Then
                If lErrDriver(lI) <> lMsgErro Then
                    pMensagem = pMensagem & Chr(13) & "- " & lI & lErrDriver(lI)
                End If
            End If
        Next lI
        Exit Function
    End If
    On Error GoTo 0

    gfAbrBasDadosAccess = True
End Function

Public Function gfObtRegistro(ByVal pNomArquivo As ADODB.Connection, ByVal pComando As String, _
                              ByRef pRegistro As clsYasRecordSet, ByRef pMensagem As String) As Boolean
'Fun��o: obt�m registro.

'Par�metros de entrada...pNomArquivo...Nome do arquivo.
'                        pComando .....Comando SELECT.
'Par�metros de sa�da.....pRegistro.....Registro (YasRecordSet).
'                        pMensagem.....Mensagem.

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "gfObtRegistro"
    lKeyDynaTrace = DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    'pNomArquivo.CommandTimeout = 30
    'Call gpGraLog(0, "667 Inicio: " & Format(Time, "hh:mm:ss") & " SQL: " & pComando)
    
    gfObtRegistro = False
    Dim lTimeTer As Date
    Dim lTimeInic As Date
    Dim lRegistro As ADODB.Recordset

    lTimeInic = Time

    If pComando Like "*EXEC *" Then
        '    Call gpGraLog(0, pComando)
    End If
    '1. Posiciona controles.
    On Error Resume Next
    Call DynaTrace.gravaLog(lKeyDynaTrace, "AGE0000 - gfObtRegistro: pComando = " & pComando, 0)
    '    Call gpGraLog(0, "AGE0000 - gfObtRegistro: pComando = " & pComando)
    Set lRegistro = New ADODB.Recordset
    Set pRegistro = New clsYasRecordSet

    lRegistro.CursorLocation = adUseClient

    '2. Ajuste o comando.
    If InStr(1, UCase(pNomArquivo.ConnectionString), "SYAS_EMIS") > 0 Then
        If InStr(1, UCase(pComando), "GED") > 0 And InStr(1, UCase(pComando), "SYAS_EMIS") > 0 Then
            'N�o altera o comando.
'        ElseIf InStr(1, UCase(pComando), "P0045400") > 0 Then
'            'N�o altera o comando.
        Else
            'Somente se n�o estiver fazendo "JOIN" entre tabelas da base GED e SYAS_EMIS atrav�s de
            'conex�o com a base SYAS_EMIS.
            pComando = Replace(pComando, "'", "!;$")
            If Len(pComando) > 5000 Then
                pComando = " exec Yproc_ExecDML2 '" & Mid(pComando, 1, 7500) & "','" & _
                           Mid(pComando, 7501) & "'"
            Else
                pComando = " exec Yproc_ExecDML2 '" & Mid(pComando, 1, 10) & "','" & _
                           Mid(pComando, 11) & "'"
            End If
        End If
    Else
        If InStr(1, UCase(pNomArquivo), "EAPDB") > 0 Then
            pComando = " exec EAPDB_ExecDML '" & Replace(pComando, "'", "!;$") & "'"
        End If
    End If
    If (Not pNomArquivo.State = adStateOpen) And (gPrograma = "D046O003") Then
        Call pNomArquivo.Open
    End If

    'Call gpGraLog(0, "667 Fim: " & Format(Time, "hh:mm:ss"))

    'Controle de DEADLOCKED
    Dim intTentativa As Integer
    intTentativa = 0
    
    Do
        If InStr(UCase(Err.Description), "DEADLOCKED") Then
            Call gpGraLog(0, "DeadLocked tentativa: " & intTentativa & " Comando: " & pComando)
            Call DynaTrace.gravaLog(lKeyDynaTrace, "DeadLocked tentativa: " & intTentativa & " Comando: " & pComando, 1)
        End If
        
        Err.Clear
            '3. Abre o registro.
            
            Call DynaTrace.gravaLog(lKeyDynaTrace, pComando, 0)
            
            lRegistro.Open pComando, pNomArquivo, adOpenForwardOnly, adLockReadOnly
            'Set lRegistro = pNomArquivo.Execute(pComando)
            intTentativa = intTentativa + 1
            
    Loop While intTentativa <= 5 And InStr(UCase(Err.Description), "DEADLOCKED") > 0

    If Err <> 0 Then
        pMensagem = "[ERR] - FTE0039 - Erro ao obter registro (Erro = " & Err & " - " & Err.Description & " - comando: " & _
                    pComando & ")."
        Call gpGraLog(0, pMensagem)
        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        Exit Function
    End If

    lTimeTer = Time
    '    Call gpGraLog(0, DateDiff("s", lTimeInic, lTimeTer) & " - " & pComando)

    If InStr(pNomArquivo, "UID=") = 0 And InStr(pNomArquivo, "Initial Catalog=") = 0 Then
        Set lRegistro.ActiveConnection = Nothing
    End If
    Select Case Err
    Case 0      'Sem erro.
        Call pRegistro.LoadRecordSet(lRegistro)
        gfObtRegistro = True
        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Case 3159   'Marcador inv�lido - arquivo dados corrompido.
        pMensagem = "[ERR] - FTE0039 - Erro ao obter registro (arquivo corrompido - comando: " & pComando & _
                    ")."
        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
    Case Else
        pMensagem = "[ERR] - FTE0039 - Erro ao obter registro (Erro = " & Err & " - " & ERROR & _
                  " - comando: " & pComando & ")."
        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
    End Select
    On Error GoTo 0
End Function

Public Function gfWhereAnd(pSql As String) As String
'Fun��o: Preenche com WHERE ou AND.

    If InStr(1, UCase(pSql), "WHERE") = 0 Then
        gfWhereAnd = pSql & " WHERE "
    Else
        gfWhereAnd = pSql & " AND "
    End If
End Function

Public Function gfFormatarCamposSQL(pAtributo As Variant, Optional SemDesformatar As Boolean = False) As String
'Fun��o: formata campos SQL.

    Select Case VarType(pAtributo)
    Case vbEmpty, vbNull:
        gfFormatarCamposSQL = "Null"
    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte:
        gfFormatarCamposSQL = Trim(Str(pAtributo))
    Case vbString:
        If gfPreenchido(pAtributo) = False Then
            gfFormatarCamposSQL = "' '"
        Else
            If SemDesformatar Then
                gfFormatarCamposSQL = "'" & Replace(CStr(pAtributo), "'", " ") & "'"
            Else
                gfFormatarCamposSQL = "'" & gfRetirarCaracterEspecial(CStr(pAtributo), True) & "'"
            End If
        End If
    Case vbDate:
        If val(pAtributo) = 0 Then
            gfFormatarCamposSQL = "Null"
        Else
            gfFormatarCamposSQL = "'" & Format(pAtributo, "mm-dd-yyyy") & " " & _
                                  Format(pAtributo, "hh:mm:ss") & "'"
        End If
    Case vbBoolean:
        gfFormatarCamposSQL = IIf(pAtributo = True, "1", "0")
    End Select
End Function

Public Function gfTratarNulo(pAtributo As Variant, pTipodeDado As TipoDados) As Variant
'Fun��o: trata campo nulo.
    Dim blnVazio As Boolean
    
    On Error GoTo Erro
    
    If IsObject(pAtributo) Then
        blnVazio = pAtributo Is Nothing Or IsNull(pAtributo)
    Else
        blnVazio = IsNull(pAtributo)
    End If
    
    If blnVazio Then
        Select Case pTipodeDado
        Case TpoBoolean
            gfTratarNulo = 0
        Case TpoData
            gfTratarNulo = "00:00:00"
        Case tpoNumerico
            gfTratarNulo = 0
        Case TpoHora
            gfTratarNulo = "00:00"
        Case tpoString
            gfTratarNulo = " "
        Case TpoValor
            gfTratarNulo = 0
        End Select
    Else
        gfTratarNulo = Trim(pAtributo)
    End If
    
    Exit Function
Erro:
    Call Me.gpGraLog(1, "gfTratarNulo: " & Err.Number & "-" & Err.Description)
End Function

Public Function gfExeSQL(ByVal pBasDados As ADODB.Connection, ByVal pComando As String, _
                         ByRef pMensagem As String, Optional ByRef pRegAfetados As Long) As Double
'Fun��o: executa comando SQL.

'Par�metros de entrada...pBasDados......base de dados.
'                        pComando ......comando SQL a ser executado.
'Par�metros de sa�da.....pMensagem......mensagem.
'                        pRegAfetados...quantidade de registros afetados pelo comando SQL (OPCIONAL).

'Resultado da fun��o.....C�digo de retorno do comando SQL: 0 = comando OK.

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    
    methodName = "gfExeSQL"
    pMensagem = Empty
    lKeyDynaTrace = DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    'pBasDados.CommandTimeout = 30
    'Call gpGraLog(0, "gfExeSQL -> Inicio: " & Format(Time, "hh:mm:ss") & " SQL: " & pComando)
    
    gfExeSQL = 999
    Dim lTimeInic As Date
    Dim lTimeTer As Date
    '''Call gpGraLog(0, pComando)
    lTimeInic = Time
    If InStr(1, UCase(pBasDados), "SYAS_EMIS") > 0 Then
        If Not App.EXEName Like "P46V600*" And Not App.EXEName Like "D46O003*" Then
            pComando = " exec Yproc_ExecDML '" & Replace(pComando, "'", "!;$") & "'"
        End If
    Else
        If InStr(1, UCase(pBasDados), "EAPDB") > 0 Then
            pComando = " exec EAPDB_ExecDML '" & Replace(pComando, "'", "!;$") & "'"
        End If
    End If

    If (Not pBasDados.State = adStateOpen) And (gPrograma = "D046O003") Then
        Call pBasDados.Open
    End If
    '    Call gpGraLog(0, pComando)
    On Error Resume Next
    
    'Controle de DEADLOCKED
    Dim intTentativa As Integer
    intTentativa = 0
    
    Do
        If InStr(UCase(Err.Description), "DEADLOCKED") Then
            Call gpGraLog(0, "DeadLocked tentativa: " & intTentativa & " Comando: " & pComando)
            Call DynaTrace.gravaLog(lKeyDynaTrace, "DeadLocked tentativa: " & intTentativa & " Comando: " & pComando, 1)
        End If
        
        Err.Clear
            
        'N�o utilizar o buffer por enquanto
        gBufferSql = False
        gScriptSql = Empty
        
        If gBufferSql = False Then
ExecutaSql:
            If Len(gScriptSql) > 0 Then
                If Not gScriptSql Like pComando Then
                    Call DynaTrace.gravaLog(lKeyDynaTrace, "�ltima Query incluida no Buffer: " & pComando, 0)
                    gScriptSql = gScriptSql & pComando & ";"
                End If
                
                Call DynaTrace.gravaLog(lKeyDynaTrace, "Flush executado: " & gScriptSql, 0)
                
                If IsMissing(pRegAfetados) Then
                    pBasDados.Execute gScriptSql
                Else
                    pBasDados.Execute gScriptSql, pRegAfetados
                End If
                
                gScriptSql = ""
            Else
                Call DynaTrace.gravaLog(lKeyDynaTrace, pComando, 0)
                
                If IsMissing(pRegAfetados) Then
                    pBasDados.Execute pComando
                Else
                    pBasDados.Execute pComando, pRegAfetados
                End If
            End If
        Else
            If Len(gScriptSql) + Len(pComando) >= 10000 Then
                GoTo ExecutaSql
            Else
                Call DynaTrace.gravaLog(lKeyDynaTrace, "Query incluida no Buffer: " & pComando, 0)
            
                gScriptSql = gScriptSql & pComando & ";"
            End If
        End If
            
        intTentativa = intTentativa + 1
    Loop While intTentativa <= 5 And InStr(UCase(Err.Description), "DEADLOCKED") > 0
    
    'Call gpGraLog(0, "666 Fim: " & Format(Time, "hh:mm:ss"))
    lTimeTer = Time

    If Err <> 0 Then
        If (InStr(1, UCase(Err.Description), "REGISTRO EXCLU�DO") <> 0 And InStr(1, UCase(pComando), "DELETE") <> 0) Then
            gfExeSQL = 0
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
            Exit Function
        Else
            pMensagem = "[ERR] - FTE0038 - Erro ao executar comando SQL: " & Err & " - " & ERROR & " - " & pComando
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
        End If
    Else
        Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    End If
    
    Dim Erro As Long
    Dim lErro As Double
    Dim lSqlErro As ADODB.ERROR
    
    If Trim(pMensagem) <> "" Then
        lErro = 1 'Err.Number
    End If
    For Each lSqlErro In pBasDados.Errors
        If lSqlErro.NativeError <> 0 Then
                lErro = 1
                pMensagem = pMensagem & " SQL Mensagem: " & pComando & " msg: " & lSqlErro.NativeError & " - " & lSqlErro.Description
        End If
    Next
    
    gfExeSQL = lErro      'Este comando precisa ser o �ltimo antes do On Error GoTo 0.
    On Error GoTo 0
End Function

Public Sub gpFecha2(ByRef pBasDados As ADODB.Connection)
'Procedure: fecha base de dados.
'Par�metro de entrada...pBasDados...Base de dados.

'Par�metro de entrada...pBasDados...Vari�vel Database.
    If Not pBasDados Is Nothing Then
        If (InStr(1, UCase(pBasDados.ConnectionString), ".MDB") > 0 Or _
            InStr(1, UCase(pBasDados.Properties("Extended Properties")), ".MDB") > 0) Or _
            gPrograma = "D046O003" Then
            If pBasDados.State = 1 Then
                pBasDados.Close
            End If
            If (InStr(1, UCase(pBasDados.ConnectionString), ".MDB") > 0 Or _
                InStr(1, UCase(pBasDados.Properties("Extended Properties")), ".MDB") > 0) Then
                Set pBasDados = Nothing
            End If
        End If
    End If

End Sub
Public Function gfIncDatas() As Boolean
'Fun��o:

    Dim lDatArqDestrava As String    'Data do arquivo que destrava o sistema.
    Dim lDatArquivo As String    'Data que est� no arquivo de destravamento
    Dim lDatMaquina As String    'Data da m�quina.
    Dim lDatRegistro As String    'Data que est� salva no registro.
    Dim lDatUltArquivo As String    'Data do �ltimo arquivo destravado.
    Dim lEndArqDestrava As String    'Endere�o do arquivo que destrava o sistema.
    Dim lIndDestravado As Byte    'Indicador de destravado:
'0 = travado.
'1 = destravado.
'3 = ?
    Dim lMensagem As String    'Mensagem.
    Dim lNumDestravado As Byte    'Utilizado no FreeFile.
    Dim lRegistro As Variant    'Registro - linha que est� no arquivo.
    Dim lStaData As String    'Status da data.
    Dim lUltArqDestravado As String    'Endere�o do �ltimo arquivo destravado.

'1. Mensagem de sistema bloqueado.
    lMensagem = "Comunique Central de Atendimento. " & vbLf & vbLf & _
                "FGE0022 - O sistema est� bloqueado pois houve altera��o na data/hora da m�quina." & _
                vbLf & vbLf & "Para desbloquear o sistema, acerte a data/hora da m�quina e entre em " & _
                "contato com a Central de Atendimento."

    '2. Obt�m registro da �ltima data de log.
    lDatRegistro = GetSetting("SYAS", "UltimaData", "LOG")
    lDatRegistro = Replace(lDatRegistro, "-", "/")
    lDatRegistro = Replace(lDatRegistro, ".", "/")
    lStaData = GetSetting("SYAS", "StatusData", "LOG")
    lDatMaquina = Format(Date, "dd/mm/yyyy")
    lUltArqDestravado = GetSetting("SYAS", "ArquivoDestravar", "LOG")

    '3. Se n�o tem registro salvo, grava a data atual no registro.
    If Not gfPreenchido(lDatRegistro) Then
        Call SaveSetting("SYAS", "UltimaData", "LOG", lDatMaquina)
    End If

    '4. Verifica se pode continuar o processo.
    If lStaData = "TRAVADO" Then
        lIndDestravado = 0
        '� necess�rio destravar o sistema - verifica se existe arquivo que destrava o sistema.
        lEndArqDestrava = Dir(gAppPath & "*.dtr")
        lDatArquivo = "01/01/1001 00:00:00"
        Do While lEndArqDestrava <> ""
            'Nome do arquivo: dd - dia atual                             --> Exemplo: 17.
            '                 mm - m�s atual                             --> Exemplo: 01.
            '                 yyyy - dois primeiros n�mero do ano atual  --> Exemplo: 2004.
            '                 hh - horas                                 --> Exemplo: 16.
            '                 mm - minutos                               --> Exemplo: 26.
            '                 ss - dois �ltimo n�mero do ano atual       --> Exemplo: 04.
            '                 resultado ddmmyyyyhhmmss.dtr               --> Exemplo: 17012004162604.dtr.
            lDatArqDestrava = Mid(lEndArqDestrava, 1, 2) & "/" & _
                              Mid(lEndArqDestrava, 3, 2) & "/" & _
                              Mid(lEndArqDestrava, 5, 4) & " " & _
                              Mid(lEndArqDestrava, 9, 2) & ":" & _
                              Mid(lEndArqDestrava, 11, 2) & ":" & _
                              Mid(lEndArqDestrava, 13, 2)
            If Len(lEndArqDestrava) = 18 Then
                If lEndArqDestrava <> lUltArqDestravado Then
                    If gfPreenchido(lUltArqDestravado) Then
                        lDatUltArquivo = Mid(lUltArqDestravado, 1, 2) & "/" & _
                                         Mid(lUltArqDestravado, 3, 2) & "/" & _
                                         Mid(lUltArqDestravado, 5, 4) & " " & _
                                         Mid(lUltArqDestravado, 9, 2) & ":" & _
                                         Mid(lUltArqDestravado, 11, 2) & ":" & _
                                         Mid(lUltArqDestravado, 13, 2)
                        If CVDate(lDatArqDestrava) < CVDate(lDatUltArquivo) Then
                            lIndDestravado = 3
                        End If
                    End If
                    If CVDate(lDatArqDestrava) > CDate(lDatArquivo) And lIndDestravado = 0 Then
                        lDatArquivo = lDatArqDestrava
                        lNumDestravado = FreeFile
                        Open lEndArqDestrava For Input As #lNumDestravado
                        While Not EOF(lNumDestravado)
                            'a) Obt�m registro de entrada.
                            Line Input #lNumDestravado, lRegistro
                            If gfCripEncript(lRegistro, 0) <> lEndArqDestrava Then
                                If gfCripEncript(lRegistro, 0) <> lEndArqDestrava & "T" Then
                                    lIndDestravado = 0
                                    Call SaveSetting("SYAS", "StatusData", "LOG", "TRAVADO")
                                Else
                                    Call SaveSetting("SYAS", "StatusData", "LOG", "DESTRAVADO")
                                    Call SaveSetting("SYAS", "UltimaData", "LOG", _
                                                     Format(Date, "dd/mm/yyyy"))
                                    lIndDestravado = 1
                                End If
                            Else
                                If lDatMaquina < lDatRegistro Then
                                    'Se a data do computador n�o foi ajustada, finalizar o sistema.
                                    lIndDestravado = 0
                                Else
                                    'Se a data do computador for igual � data do registro, destravar.
                                    Call SaveSetting("SYAS", "StatusData", "LOG", "DESTRAVADO")
                                    Call SaveSetting("SYAS", "ArquivoDestravar", "LOG", lEndArqDestrava)
                                    lIndDestravado = 1
                                End If
                            End If
                        Wend
                        Close #lNumDestravado
                        Kill (gAppPath & "*.dtr")
                    Else
                        Kill (gAppPath & "*.dtr")
                    End If
                End If
            End If
            lEndArqDestrava = Dir   'Obt�m pr�xima entrada.
        Loop
        If lIndDestravado = 0 Or lIndDestravado = 3 Then
            MsgBox lMensagem
            gfIncDatas = False
            Exit Function
        End If
    Else
        If CDate(lDatMaquina) < CDate(lDatRegistro) Then
            'Finalizar o sistema.
            Call SaveSetting("SYAS", "StatusData", "LOG", "TRAVADO")
            If Dir(gAppPath & "*.dtr") <> "" Then
                Kill (gAppPath & "*.dtr")
            End If
            MsgBox lMensagem
            gfIncDatas = False
            Exit Function
        Else
            If lDatMaquina <> lDatRegistro Then
                'Registra a �ltima data.
                Call SaveSetting("SYAS", "UltimaData", "LOG", lDatMaquina)
            End If
        End If
    End If
    gfIncDatas = True
End Function

Public Function gfAbrBasSQL(ByVal pIdeBasDados As String, ByVal pTipBasDados As String, _
                            ByVal pSerBasDados As String, ByVal pEndBasDados As String, _
                            ByVal pUIDBasDados As String, ByVal pPWDBasDados As String, _
                            ByRef pBasDados As ADODB.Connection, ByRef pMensagem As String) _
                            As Boolean
'Fun��o: abre base de dados.

'Par�metros de entrada...pIdeBasDados...Identificador da base de dados.
'                        pTipBasDados...Tipo da base de dados.
'                        pEndBasDados...Endere�o da base de dados.
'Par�metros de sa�da.....pBasDados......Base de dados.
'                        pMensagem......Mensagem.

    Dim lCnStr As String    'Linha para comando OPEN.
    Dim lDriver As String
    Dim lCN As ADODB.Connection

    gfAbrBasSQL = False

    'a) Verifica identifica��o da base de dados.
    If Not gfPreenchido(pIdeBasDados) Then
        pMensagem = "[ERR] - FTE0051 - Erro ao abrir base de dados (identifica��o da base de dados n�o " & _
                    "preenchida)."
        Exit Function
    End If
    Select Case UCase$(pIdeBasDados)
    Case "DBDIVERSOS", "GED", "EAPDB", "P0044800", "SYASPORTAL", "YAS_ND", "SYAS_EMIS"
        'Nada a fazer - bases de dados v�lidas.
    Case Else
        pMensagem = "[ERR] - FTE0052 - Erro ao abrir base de dados (identifica��o da base de dados " & _
                    "desconhecida: " & pIdeBasDados & ")."
        Exit Function
    End Select
    'b) Verifica tipo da base de dados.
    If Not gfPreenchido(pTipBasDados) Then
        pMensagem = "[ERR] - FTE0053 - Erro ao abrir base de dados (tipo da base de dados " & pIdeBasDados & _
                  " n�o preenchido)."
        Exit Function
    End If
    If pTipBasDados <> "2" Then
        pMensagem = "[ERR] - FTE0052 - Erro ao abrir base de dados (tipo da base de dados " & pIdeBasDados & _
                  " desconhecido: " & pTipBasDados & ")."
        Exit Function
    End If
    'c) Verifica endere�o da base de dados.
    If Not gfPreenchido(pEndBasDados) Then
        pMensagem = "[ERR] - FTE0055 - Erro ao abrir base de dados (endere�o da base de dados " & _
                    pEndBasDados & " n�o preenchido)."
        Exit Function
    End If

    'd) Posiciona conex�o.
    On Error Resume Next
    Set lCN = New ADODB.Connection
    If Err <> 0 Then
        pMensagem = "[ERR] - FTE0030 - Erro ao abrir base de dados (" & pIdeBasDados & _
                  " - comando SET New ADODB.Connection - erro " & Err & " - " & ERROR & ")."
        On Error GoTo 0
        Exit Function
    End If

    Select Case LCase(pSerBasDados)
    Case "192.168.1.81", "SPX21181PSLUSQL"
        lDriver = "{SQL Server Native Client 10.0}"
    Case Else
'        lDriver = "{SQL Server}"
        lDriver = "SQLOLEDB.1"
    End Select

    If pIdeBasDados = "YAS_ND" Then
        lCnStr = "Provider=" & lDriver & _
                 ";Data Source=" & pSerBasDados & _
                 ";User ID=" & pUIDBasDados & _
                 ";Password=" & pPWDBasDados & _
                 ";Initial Catalog=" & pEndBasDados & _
                ";Persist Security Info=True"
    Else
'        lCnStr = "driver=" & lDriver & _
'                 ";server=" & pSerBasDados & _
'                 ";uid=" & LCase(pUIDBasDados) & _
'                 ";pwd=" & LCase(pPWDBasDados) & _
'                 ";database=" & pEndBasDados

        lCnStr = "Provider=" & lDriver & _
                         ";Data Source=" & pSerBasDados & _
                         ";User ID=" & LCase(pUIDBasDados) & _
                         ";Password=" & LCase(pPWDBasDados) & _
                         ";Initial Catalog=" & pEndBasDados & _
                        ";Persist Security Info=True"
    End If

    lCN.ConnectionString = lCnStr

    lCN.ConnectionTimeout = 900
    lCN.CommandTimeout = 900

    If Err <> 0 Then
        pMensagem = "[ERR] - FTE0035 - Erro ao abrir base de dados (" & pIdeBasDados & _
                  " - Comando CursorLocation = adUseClient - erro: " & Err & " - " & ERROR & ")."
        On Error GoTo 0
        Exit Function
    End If
    ''Call gpGraLog(0, "AGE0000 - gfAbrBasSQL : ConnectionString=" & lCN.ConnectionString)
    'e) Abre base de dados.

    If mfConexaoAberta(pIdeBasDados, pSerBasDados, pEndBasDados, pUIDBasDados, pPWDBasDados, mConexaoSQL) = False Then
        lCN.Open
        Call Me.gpFecha2(mConexaoSQL)
        Set mConexaoSQL = Nothing
        Set mConexaoSQL = New ADODB.Connection
        Set mConexaoSQL = lCN
    Else
        pMensagem = ""
    End If

    If Err <> 0 Then
        'pMensagem = "FTE0001 - Erro na base de dados (" & pIdeBasDados & _
         '        " - comando Open " & lCnStr & " - erro: " & Err & " - " & Error & ")."
        pMensagem = "FTE0001 - Erro na base de dados. - erro: " & Err & " - " & ERROR & ")."
        GoTo SAIDA
    End If

Retorno:
    Set pBasDados = mConexaoSQL

    gfAbrBasSQL = True

SAIDA:
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lCN = Nothing
    Return

End Function

Public Function gfRetirarCaracterEspecial(pTexto As String, Optional pUcase As Boolean = False, Optional pRetiraCrLf As Boolean = True, Optional pRetiraDemais As Boolean = False) As String
    Dim i As Long    'Utilizado no for...loop
    Dim Letra As String    'Letra
    Dim lretorno As String    'Retorno

    gfRetirarCaracterEspecial = ""

    For i = 1 To Len(pTexto)
        Letra = Mid(pTexto, i, 1)
        Select Case Asc(Letra)
        Case 34, 39, 94, 96, 126, 168, 180   ' ", ' , ^, `, ~, �, �,
            Letra = " "
        Case 13, 10, 59  'Enter, ;
            If pRetiraCrLf Then
                Letra = " "
            End If
        Case 9  'TAB
            Letra = " "
        Case 192 To 198    ' �, �, �, �, �, �, �
            Letra = "A"
        Case 224 To 229    ' �, �, �, �, �, �, �
            Letra = "a"
        Case 199        ' �
            Letra = "C"
        Case 231        ' �
            Letra = "c"
        Case 200 To 203    ' �, �, �, �
            Letra = "E"
        Case 232 To 235    ' �, �, �, �
            Letra = "e"
        Case 204 To 207    ' �, �, �, �
            Letra = "I"
        Case 236 To 239    ' �, �, �, �
            Letra = "i"
        Case 209        ' �
            Letra = "N"
        Case 241        ' �
            Letra = "n"
        Case 210 To 214    ' �, �, �, �, �
            Letra = "O"
        Case 242 To 246    ' �, �, �, �, �
            Letra = "o"
        Case 217 To 220    ' �, �, �, �
            Letra = "U"
        Case 249 To 252    ' �, �, �, �
            Letra = "u"
        Case 33, 35, 36, 37, 38, 58, 60, 61, 62, 63, 64    ' !, #, $, %, &, :, <, =, >, ?, @
            If pRetiraDemais Then
                Letra = " "
            End If
        Case 91, 92, 93, 95, 123, 124, 125, 126    ' [, \, ], _, {, |, }, ~
            If pRetiraDemais Then
                Letra = " "
            End If
        Case 40 To 47    ' (, ), *, +, -, ., /
            If pRetiraDemais Then
                Letra = " "
            End If
        End Select
        lretorno = lretorno & Letra
    Next i
'    If pUcase = True Then
'        'Mudo o texto para mai�sculo.
'        gfRetirarCaracterEspecial = UCase(lretorno)
'    Else
        gfRetirarCaracterEspecial = lretorno
'    End If
End Function

Function gfCripEncript(ByVal strng, force%)
'C�digo copiado da internet para encriptar e decriptar

    Dim chK As String
    Dim look As String
    Dim addin As String
    Dim eNcryptflag As String
    Dim passUP As String
    Dim passMAx As String
    Dim looPer As Variant
    Dim tOChange As String
    Dim password As String

    'Set error capture routine
    On Local Error GoTo ErrorHandler

    password = "Password"

    'Is there a strng$ to work with?
    If Len(strng) = 0 Then Error 31100

    'Check if file is encrypted and not forcing
    If force% = 0 Then

        'Check for encryption ID tag
        chK = Left$(strng, 4) + Right$(strng, 4)
        If chK = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then

            'Remove ID tag
            strng = Mid$(strng, 5, Len(strng) - 8)

            'String was encrypted so filter out CHR$(1) flags
            look = 1
            Do
                look = InStr(look, strng, Chr$(1))
                If look = 0 Then
                    Exit Do
                Else
                    addin = Chr$(Asc(Mid$(strng, look + 1)) - 1)
                    strng = Left$(strng, look - 1) + addin + Mid$(strng, look + 2)
                End If
                look = look + 1
            Loop

            'Since it is encrypted we want to decrypt it
            eNcryptflag = False
        Else
            'Tag not found so flag to encrypt string
            eNcryptflag = True
        End If
    Else
        'force% flag set, ecrypt string regardless of tag
        eNcryptflag = True
    End If

    'Set up variables
    passUP = 1
    passMAx = Len(password)

    'Tack on leading characters to prevent repetative recognition
    password = Chr$(Asc(Left$(password, 1)) Xor passMAx) + password
    password = Chr$(Asc(Mid$(password, 1, 1)) Xor Asc(Mid$(password, 2, 1))) + password
    password = password + Chr$(Asc(Right$(password, 1)) Xor passMAx)
    password = password + Chr$(Asc(Right$(password, 2)) Xor Asc(Right$(password, 1)))

    'If Encrypting add password check tag now so it is encrypted with string
    If eNcryptflag = True Then
        strng = Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") + strng
    End If

    'Loop until scanned though the whole string
    For looPer = 1 To Len(strng)
        DoEvents
        'Alter character code
        tOChange = Asc(Mid$(strng, looPer, 1)) Xor Asc(Mid$(password, passUP, 1))

        'Insert altered character code
        Mid$(strng, looPer, 1) = Chr$(tOChange)

        'Scroll through password string one character at a time
        passUP = passUP + 1
        If passUP > passMAx + 4 Then passUP = 1
    Next looPer

    'If encrypting we need to filter out all bad character codes (0, 10, 13, 26)
    If eNcryptflag = True Then
        'First get rid of all CHR$(1) since that is what we use for our flag
        look = 1
        Do
            look = InStr(look, strng, Chr$(1))
            If look > 0 Then
                strng = Left$(strng, look - 1) + Chr$(1) + Chr$(2) + Mid$(strng, look + 1)
                look = look + 1
            End If
        Loop While look > 0

        'Check for CHR$(0)
        Do
            look = InStr(strng, Chr$(0))
            If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(1) + Mid$(strng, look + 1)
        Loop While look > 0

        'Check for CHR$(10)
        Do
            look = InStr(strng, Chr$(10))
            If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(11) + Mid$(strng, look + 1)
        Loop While look > 0

        'Check for CHR$(13)
        Do
            look = InStr(strng, Chr$(13))
            If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(14) + Mid$(strng, look + 1)
        Loop While look > 0

        'Check for CHR$(26)
        Do
            look = InStr(strng, Chr$(26))
            If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(27) + Mid$(strng, look + 1)
        Loop While look > 0

        'Tack on encryted tag
        strng = Chr$(1) + "KT" + Chr$(1) + strng + Chr$(1) + "KT" + Chr$(1)

    Else

        'We decrypted so ensure password used was the correct one
        If Left$(strng, 9) <> Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") Then
            'Password bad cause error
            Error 31100
        Else
            'Password good, remove password check tag
            strng = Mid$(strng, 10)
        End If
    End If

    'Set function equal to modified string
    gfCripEncript = strng

    'Were out of here
    Exit Function

ErrorHandler:
    'We had an error!  Were out of here
    Exit Function
End Function

Public Sub gpMenErro(ByVal pMensagem As String, Optional ByVal pClasseNorma As Integer = 0)
'Procedure: posiciona mensagem de erro.
'Par�metro de entrada.........pMensagem...Mensagem.

'[FRH] - F:900097 [FGE0155]
'[FRH]-GLB_sinal_concess-> true se as propr.Cod_Clasf e Cod_Compl_Clasf do pobjstcitem-mfcalpreitem s�o relativos a "022-13-Concession�ria".
    If UCase(pMensagem) Like "*CODIGO DO RAMO NAO CONFERE COM A APOLICE*" And gPrograma <> "D046O003" And (InStr(UCase(gNomUsuario), "CYBER") > 0 Or InStr(UCase(Command), "DEV") > 0) Then
        Exit Sub
    End If
    
    If GLB_sinal_concess And ( _
       pMensagem Like "*0058*0022*" Or _
       pMensagem Like "*0062*0022*" Or _
       (UCase(pMensagem) Like "FGE0155*COBERTURA*86*N�O*PODE*SER*CONTRATADA*SEM*COBERTURA*28*" And _
        gCodProduto >= 956) _
      ) Then
        Exit Sub
    End If
    If gCod_Ramo = e_CodRamo_Roubos And (pMensagem Like "*52*") Then
        Exit Sub
    End If
    If pClasseNorma = 0 Then
        If (UCase(pMensagem) Like "*NORMA*" Or UCase(pMensagem) Like "*FGE0116*") And _
        Not UCase(pMensagem) Like "*SEM ACEITA*" Then
                gQuaErrosNormas = gQuaErrosNormas + 1
        Else
                gQuaErrosDemais = gQuaErrosDemais + 1
        End If
    ElseIf pClasseNorma = 3 Then
        gQuaErrosDemais = gQuaErrosDemais + 1
    Else
        gQuaErrosNormas = gQuaErrosNormas + 1
    End If

    'Rotina de captura de erros original
    If TypeName(gerros) = "Empty" Then
        Call gpInicializar_gErros
    End If
    
    gQuaErros = gQuaErros + 1
    ReDim Preserve gerros(gQuaErros)
    gerros(gQuaErros) = pMensagem
    '[FRH]

End Sub

Public Function gfTrunca(ByVal pCampo As String, ByVal pQuaCasas As Byte) As Single
'Fun��o #031: trunca campo.

    Dim lPosicao As Byte  'Utilizado com InStr.

    lPosicao = InStr(CDec(pCampo), ",")
    If lPosicao = 0 Then
        lPosicao = InStr(CDec(pCampo), ".")
    End If
    If lPosicao = 0 Then
        gfTrunca = CDec(pCampo)
    Else
        If Not IsNumeric(Mid(Trim(CDec(pCampo)), lPosicao + pQuaCasas, 1)) Then
            gfTrunca = CDec(pCampo)
        Else
            gfTrunca = (Mid(CDec(pCampo), 1, lPosicao - 1) & _
                        Mid(CDec(pCampo), lPosicao + 1, pQuaCasas)) / 10 ^ pQuaCasas
        End If
    End If
End Function

Public Sub gpInicia(ByRef pPropTabela As String)
'Procedure: obt�m informa��es inicias do sistema:
'           - Arquivo de inicializa��o: tipo do usu�rio, endere�os das bases de dados (endere�o do
'                                       servidor), endere�os dos arquivos imagem e configura��o de
'                                       impress�o.
'           - Nome do usu�rio (ou c�digo do usu�rio no caso de corretor).

    Dim lAuxiliar As String    'Auxiliar.
    Dim lEndSyasEdi_EXE As String    'Endere�o do arquivo SyasEdi.EXE.
    Dim lEndSyaslibSGS As String    'Endere�o do arquivo SYASLIB.SGS.
    Dim lLiberado As Boolean    'Indica que o Segsul est� liberado.
    Dim lLimite As Double             'Data limite para o arquivo de libera��o do Segsul.
    Dim lNumSyasEdi_EXE As Integer    'Utilizado com FreeFile para o arquivo SYASEDI.EXE.
    Dim lNumSyaslibSGS As Integer    'Utilizado com FreeFile para o arquivo SYASLIB.SGS.
    Dim lPosicao As Integer    'Utilizado com InStr.
    Dim lRegSyasEdi_EXE As String    'Registro do arquivo SYASEDI.EXE.
    Dim lRegSyaslibSGS As String    'Registro do arquivo SYASLIB.SGS.

'1. Posiciona data de hoje.
    gHoje = CDbl(Format$(Now, "yyyymmdd"))
    lAuxiliar = Format$(gHoje, "00000000")
    gHojDia = CInt(Mid$(lAuxiliar, 7, 2))
    gHojMes = CInt(Mid$(lAuxiliar, 5, 2))
    gHojAno = CInt(Mid$(lAuxiliar, 1, 4))
    gForHoje = Format$(Now, "dd/mm/yyyy")

    '2. Posiciona controle dos endere�os dos arquivos (gAppPath).
    If Right$(App.Path, 1) = "\" Then
        gAppPath = App.Path
    Else
        gAppPath = App.Path & "\"
    End If

    '3. Obt�m dados do arquivo de inicializa��o.
    Call gfObtSYASMULTINI(pPropTabela)

    If gPrograma = "YASUDA.EXE" Then
        'Sai da procedure para processar o arquivo YASUDA.EXE processar as atualiza��es.
        Exit Sub
    End If

    '4. Verifica informa��es.
    'a) Tipo de usu�rio.
    If Not gfPreenchido(gTipUsuario) Then
        Call gpGraLog(2, "FTE0042 - Erro no arquivo de inicializa��o (tipo do usu�rio n�o preenchido).")
        Exit Sub    'End   ***MOSILVA Verificar
    End If
    If Not IsNumeric(gTipUsuario) Then
        Call gpGraLog(2, "FTE0043 - Erro no arquivo de inicializa��o (tipo do usu�rio n�o num�rico).")
        Exit Sub    'End   ***MOSILVA Verificar
    End If
    If gTipUsuario <> e_Tipo_Usu_Corretor And gTipUsuario <> e_Tipo_Usu_Yasuda And gTipUsuario <> e_Tipo_Usu_CorretorYasuda Then
        Call gpGraLog(2, "FTE0044 - Erro no arquivo de inicializa��o (tipo do usu�rio desconhecido).")
        Exit Sub    'End   ***MOSILVA Verificar
    End If
    'b) Nome do usu�rio.
    'For�a Usuario
    Select Case gTipUsuario
    Case 1        'Yasuda.
        'For�ar usuario
        gNomUsuario = Trim$(M99V001.GetLoginUserName)
        If Not gfPreenchido(gNomUsuario) Then
            Call gpGraLog(2, "FTE0045 - Identifica��o do usu�rio Yasuda n�o localizada.")
            Exit Sub    'End   ***MOSILVA Verificar
        End If
        If Len(gNomUsuario) > 6 Then
            gNomUsuario6 = Left$(gNomUsuario, 6)
        Else
            gNomUsuario6 = gNomUsuario
        End If
    Case 2        'Corretor.
        lEndSyasEdi_EXE = gAppPath & "SYASEDI.EXE"
        If Not gfPreenchido(Dir(lEndSyasEdi_EXE)) Then
            lEndSyasEdi_EXE = gAppPath & "EAPSCRIPT.INI"
        End If
        If Not gfPreenchido(Dir(lEndSyasEdi_EXE)) Then
            gNomUsuario = "XXXXXXXX"
        Else
            lNumSyasEdi_EXE = FreeFile
            On Error Resume Next
            Open lEndSyasEdi_EXE For Input Access Read As #lNumSyasEdi_EXE
            If Err <> 0 Then
                Call gpGraLog(2, "FTE0026 - Erro ao abrir arquivo " & lEndSyasEdi_EXE & ": " & _
                                 Err & " - " & ERROR & vbLf & vbLf & "O sistema ser� encerrado.")
                On Error GoTo 0
                Exit Sub    'End   ***MOSILVA Verificar
            End If
            On Error GoTo 0
            Do While Not EOF(lNumSyasEdi_EXE)
                Line Input #lNumSyasEdi_EXE, lRegSyasEdi_EXE
                lRegSyasEdi_EXE = UCase$(lRegSyasEdi_EXE)
                lPosicao = InStr(lRegSyasEdi_EXE, "SENDER=""")
                If lPosicao <> 0 Then
                    gNomUsuario = Mid(lRegSyasEdi_EXE, lPosicao + 8, 6)
                    Exit Do
                End If
            Loop
            Close #lNumSyasEdi_EXE
            If Not gfPreenchido(gNomUsuario) Then
                gNomUsuario = "XXXXXXXX"
            End If
        End If
        gNomUsuario6 = Left$(gNomUsuario, 6)
    Case 3
        gNomUsuario = Trim$(M99V001.GetLoginUserName)
        If Not gfPreenchido(gNomUsuario) Then
            Call gpGraLog(2, "FTE0045 - Identifica��o do usu�rio Yasuda n�o localizada.")
            Exit Sub    'End   ***MOSILVA Verificar
        End If
        If Len(gNomUsuario) > 6 Then
            gNomUsuario6 = Left$(gNomUsuario, 6)
        Else
            gNomUsuario6 = gNomUsuario
        End If

        'gCodSeguradoYas = "0777702"
        'gCod_Corretor = 1601

    End Select

    gQtdMaxComp = 4

End Sub


Public Sub gfObtSYASMULTINI(ByRef pPropTabela As String)
'Procedure: obt�m dados do arquivo SYASMULT.INI.

    Dim lDado As String    'Dado da linha de registro.
    Dim lDesconhecido As Boolean    'Tipo desconhecido.
    Dim lEndBasDados() As String  'Bases de dados: endere�os.
    Dim lEndereco As String    'Endereco.
    Dim lEndSYASMULTINI As String   'Endere�o do arquivo SYASMULT.INI.
    Dim lI As Integer             'Utilizado com For...Next.
    Dim lImaImpressao() As String   'Arquivos de impress�o: endere�o do arquivo de imagem.
    Dim lNomBasDados() As String  'Bases de dados: nomes.
    Dim lNumSYASMULTINI As Integer  'Utilizado com FreeFile para o arquivo SYASMULT.INI.
    Dim lPosicao As Integer    'Utilizado com InStr.
    Dim lQuaBasDados As Byte  'Bases de dados: quantidade.
    Dim lRegSYASMULTINI As String   'Registro do arquivo SYASMULT.INI.
    Dim lTamanho As Integer    'Tamanho.
    Dim lTipBasDados() As String  'Bases de dados: tipos.
    Dim lTipo As String    'Tipo de linha do arquivo de inicializa��o.
    Dim lTitulo As String    'Linha de t�tulo.
    Dim lEndSiscad As String
    Dim lEndP47V001 As String

    Static lsVerificado As Boolean
    Static lsPropTabela As String

    ' Habilita o log de inicio, para o caso de dar erro antes da leitura do INI
    gEnableFileLog = True
    
    If lsVerificado = True Then
        GoTo Retorno
    Else
        lsVerificado = True
    End If

    '1. Abre arquivo SYASMULT.INI.
    lEndSiscad = ""
    lEndP47V001 = ""

    lEndSYASMULTINI = gAppPath & "SYASMULT.INI"
    If Not gfPreenchido(Dir(lEndSYASMULTINI)) Then
        Call gpGraLog(2, "FTE0057 - Arquivo n�o localizado: " & lEndSYASMULTINI & _
                       " (obter dados de inicializa��o do sistema)." & vbLf & vbLf & _
                         "O sistema ser� encerrado.")
        GoTo SAIDA    'End   ***MOSILVA Verificar
    End If
    lNumSYASMULTINI = FreeFile
    On Error Resume Next
    Open lEndSYASMULTINI For Input Access Read As lNumSYASMULTINI
    If Err <> 0 Then
        Call gpGraLog(2, "FTE0048 - Erro ao abrir arquivo " & lEndSYASMULTINI & ": " & Err & " - " & _
                         ERROR & vbLf & vbLf & "O sistema ser� encerrado.")
        On Error GoTo 0
        GoTo SAIDA    'End   ***MOSILVA Verificar
    End If
    On Error GoTo 0

    '2. Trata registros do arquivo de inicializa��o.
    lDesconhecido = False
    Do While Not EOF(lNumSYASMULTINI)
        Line Input #lNumSYASMULTINI, lRegSYASMULTINI
        If gfPreenchido(lRegSYASMULTINI) Then
            Select Case Left$(lRegSYASMULTINI, 1)
            Case ";"    'Linha de observa��o.
                'Nada a processar - linha de observa��o.
            Case "["    'Linha de t�tulo.
                lTitulo = Trim$(UCase$(lRegSYASMULTINI))
            Case Else   'Linha de tipo.
                lRegSYASMULTINI = Trim$(lRegSYASMULTINI)
                lPosicao = InStr(lRegSYASMULTINI, "=")
                lTamanho = Len(lRegSYASMULTINI)
                If lPosicao = 0 Then
                    Call gpGraLog(2, "FTE0049 - Erro no arquivo de inicializa��o (sinal de igual " & _
                                     "n�o localizado no registro)." & vbLf & vbLf & _
                                     lRegSYASMULTINI & vbLf & vbLf & "O sistema ser� encerrado.")
                    GoTo SAIDA    'End   ***MOSILVA Verificar
                End If
                If lPosicao = 1 Then
                    Call gpGraLog(2, "FTE0049 - Erro no arquivo de inicializa��o (sinal de igual " & _
                                     "na primeira posi��o do registro)." & vbLf & vbLf & _
                                     lRegSYASMULTINI & vbLf & vbLf & "O sistema ser� encerrado.")
                    GoTo SAIDA    'End   ***MOSILVA Verificar
                End If
                If lPosicao = lTamanho Then
                    Call gpGraLog(2, "FTE0049 - Erro no arquivo de inicializa��o (sinal de igual " & _
                                     "na �ltima posi��o do registro)." & vbLf & vbLf & _
                                     lRegSYASMULTINI & vbLf & vbLf & "O sistema ser� encerrado.")
                    GoTo SAIDA    'End   ***MOSILVA Verificar
                End If
                lTipo = Trim$(UCase$(Left$(lRegSYASMULTINI, lPosicao - 1)))
                lDado = Trim$(Mid$(lRegSYASMULTINI, lPosicao + 1))
                Select Case lTitulo
                Case "[AUXILIO]"
                Case "[DBSYASEMIS]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipSYASEMIS = lDado
                    Case "SERVER"
                        gSerSYASEMIS = lDado
                    Case "DATABASE"
                        gEndSYASEMIS = lDado
                    Case "UID"
                        gUidSYASEMIS = lDado
                    Case "PWD"
                        gPwdSYASEMIS = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[CORTENOMESEGURADORA]"
                    Select Case lTipo
                    Case "DATASOMPO"
                        gDataCorteSompo = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[WSMOTORCALC]"
                    Select Case lTipo
                    Case "URL"
                        gUrlWebServiceMotorCalc = lDado
                    Case "ENABLED"
                        If lDado = "true" Then
                            gWebServiceEnabled = True
                        Else
                            gWebServiceEnabled = False
                        End If
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[WSNORMAS]"
                    Select Case lTipo
                    Case "URL"
                        gUrlWebServiceNormas = lDado
                    Case "ENABLED"
                        If UCase(lDado) = "TRUE" Then
                            gWebServiceNormasEnabled = True
                        Else
                            gWebServiceNormasEnabled = False
                        End If
                    Case "TARIFAINICIO"
                        If IsNumeric(lDado) = True Then
                            gWebServiceNormasTarifa = CLng(lDado)
                        Else
                            gWebServiceNormasTarifa = 0
                        End If
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[WSBOLETO]"
                    Select Case lTipo
                    Case "URL"
                        gUrlWebServiceBoleto = lDado
                    Case "ENABLED"
                        If UCase(lDado) = "TRUE" Then
                            gWebServiceBoletoEnabled = True
                        Else
                            gWebServiceBoletoEnabled = False
                        End If
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[WSBOLETOMULTISOMPO]"
                    Select Case lTipo
                    Case "URL"
                        gUrlWebServiceBoleto = lDado
                    Case "ENABLED"
                        If UCase(lDado) = "TRUE" Then
                            gWebServiceBoletoEnabled = True
                        Else
                            gWebServiceBoletoEnabled = False
                        End If
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[RELATORIO]"
                    Select Case lTipo
                    Case "URL"
                        gUrlRelatorio = lDado
                    Case "USUARIO"
                        gUsuarioRelatorio = lDado
                    End Select
                Case "[LOG]"
                    Select Case lTipo
                    Case "ENABLED"
                        If lDado = "true" Then
                            gEnableFileLog = True
                        Else
                            gEnableFileLog = False
                        End If
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTCOT1]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfCot1 = lDado
                    Case "IMAGEM"
                        gEndImagCot1 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTCOT2]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfCot2 = lDado
                    Case "IMAGEM"
                        gEndImagCot2 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTCOT3]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfCot3 = lDado
                    Case "IMAGEM"
                        gEndImagCot3 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTCOTB]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfCotB = lDado
                    Case "IMAGEM"
                        gEndImagCotB = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTPRO1]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfPro1 = lDado
                    Case "IMAGEM"
                        gEndImagPro1 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTPRO2]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfPro2 = lDado
                    Case "IMAGEM"
                        gEndImagPro2 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTPRO3]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfPro3 = lDado
                    Case "IMAGEM"
                        gEndImagPro3 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTPRO4]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfPro4 = lDado
                    Case "IMAGEM"
                        gEndImagPro4 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTPROB]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfProB = lDado
                    Case "IMAGEM"
                        gEndImagProB = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTPROC]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfProC = lDado
                    Case "IMAGEM"
                        gEndImagProC = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTPROD]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfProD = lDado
                    Case "IMAGEM"
                        gEndImagProD = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTCOTA]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfCotA = lDado
                    Case "IMAGEM"
                        gEndImagCotA = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTCOTC]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfCotC = lDado
                    Case "IMAGEM"
                        gEndImagCotC = lDado
                    Case Else
                        lDesconhecido = True
                    End Select


                    '*********************************************************
                    ' FORMULARIOS ANTIGOS
                    '*********************************************************
                Case "[MULTCOT1_V0]"
                    Select Case lTipo
                    Case "NOME"
                        gEndCot1_v0 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTCOT2_V0]"
                    Select Case lTipo
                    Case "NOME"
                        gEndCot2_v0 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTCOT3_V0]"
                    Select Case lTipo
                    Case "NOME"
                        gEndCot3_v0 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTPRO1_V0]"
                    Select Case lTipo
                    Case "NOME"
                        gEndPro1_v0 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTPRO2_V0]"
                    Select Case lTipo
                    Case "NOME"
                        gEndPro2_v0 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTPRO3_V0]"
                    Select Case lTipo
                    Case "NOME"
                        gEndPro3_v0 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTPRO4_V0]"
                    Select Case lTipo
                    Case "NOME"
                        gEndPro4_v0 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[MULTPROB_V0]"
                    Select Case lTipo
                    Case "NOME"
                        gEndProB_v0 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                    '*********************************************************
                    ' FORMULARIOS ROUBO
                    '*********************************************************
                Case "[ROUBOCOT1]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfRob1 = lDado
                    Case "IMAGEM"
                        gEndImagRob1 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[ROUBOCOT2]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfRob2 = lDado
                    Case "IMAGEM"
                        gEndImagRob2 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[ROUBOCOT3]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfRob3 = lDado
                    Case "IMAGEM"
                        gEndImagRob3 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[ROUBOCOT4]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfRob4 = lDado
                    Case "IMAGEM"
                        gEndImagRob4 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[ROUBOCOT5]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfRob5 = lDado
                    Case "IMAGEM"
                        gEndImagRob5 = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[COT_LMI]"
                    Select Case lTipo
                    Case "CONFIG"
                        'gEndConfRob5 = lDado
                    Case "IMAGEM"
                        gEndImagCotLMI = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[REL_LMI]"
                    Select Case lTipo
                    Case "CONFIG"
                    Case "IMAGEM"
                        gEndImagRelLMI = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                    '*********************************************************
                Case "[P00MPED]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipP00MPED = lDado
                    Case "DATABASE"
                        gEndP00MPED = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                    gTipP00MULTI = gTipP00MPED
                    gEndP00MULTI = gAppPath & "P00MULTI.MDB"
                Case "[P00APOL]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipP00APOL = lDado
                    Case "DATABASE"
                        gEndP00APOL = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[REMESSA]"
                    Select Case lTipo
                    Case "ENDERECO"
                        gEndRemessa = lDado
                    End Select
                Case "[RETORNO]"
                Case "[USO EM REDE]"
                Case "[USO VERSAO DOS]"
                Case "[USUARIO]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipUsuario = lDado
                    Case "TIPOHOMOLOG"
                        gTipoHomologacao = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[ATUALIZACAO]"
                    gEndSYASUpDt = lDado
                    If gfPreenchido(gEndSYASUpDt) Then
                        lTamanho = Len(gEndSYASUpDt)
                        If lTamanho = 5 Then
                            gEndSYASUpDt = gAppPath
                        Else
                            gEndSYASUpDt = Left$(gEndSYASUpDt, lTamanho - 5)
                        End If
                    End If
                Case "[PROPRIETARIO TABELA]"
                    lsPropTabela = lDado
                Case "[DBDIVERSOS]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipDIVERSOS = lDado
                    Case "SERVER"
                        gSerDIVERSOS = lDado
                    Case "DATABASE"
                        gEndDIVERSOS = lDado
                    Case "UID"
                        gUidDIVERSOS = lDado
                    Case "PWD"
                        gPwdDIVERSOS = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[DBYAS_ND]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipYAS_ND = lDado
                    Case "SERVER"
                        gSerYAS_ND = lDado
                    Case "DATABASE"
                        gEndYAS_ND = lDado
                    Case "UID"
                        gUidYAS_ND = lDado
                    Case "PWD"
                        gPwdYAS_ND = lDado
                    Case Else
                        lDesconhecido = True
                    End Select

                Case "[DBGED]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipGed = lDado
                    Case "SERVER"
                        gSerGed = lDado
                    Case "DATABASE"
                        gEndGed = lDado
                    Case "UID"
                        gUidGed = lDado
                    Case "PWD"
                        gPwdGed = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[DBEAPDB]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipEAPDB = lDado
                    Case "SERVER"
                        gSerEAPDB = lDado
                    Case "DATABASE"
                        gEndEAPDB = lDado
                    Case "UID"
                        gUidEAPDB = lDado
                    Case "PWD"
                        gPwdEAPDB = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[DBSYASPORTAL]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipSyasPortal = lDado
                    Case "SERVER"
                        gSerSyasPortal = lDado
                    Case "DATABASE"
                        gEndSyasPortal = lDado
                    Case "UID"
                        gUidSyasPortal = lDado
                    Case "PWD"
                        gPwdSyasPortal = lDado
                    Case Else
                        lDesconhecido = True
                    End Select
                Case "[PATH_ARQUIVOS]"
                    gPathArquivos = lDado
                Case "[IMPRESSORAFRENTE]"
                    gImprFrente = lDado
                Case "[IMPRESSORAFRENTEVERSO]"
                    gImprFrenteVerso = lDado
                Case "[SISCAD]"
                    lEndSiscad = lDado
                Case "[P47V001]"
                    lEndP47V001 = lDado
                Case "[DIR_IST]"
                    gDir_Ist = lDado
                Case "[CONFIG]"
                    Select Case lTipo
                    Case "MAPI_SERVER_EXCHANGE"
                        MAPI_SERVER_EXCHANGE = lDado
                    Case "MAPI_CONTA_EXCHANGE"
                        MAPI_CONTA_EXCHANGE = lDado
                    Case "MAPI_Profile"
                        MAPI_Profile = lDado
                    End Select
                Case "[CAMINHOPI]"
                    gPathImagePI = lDado
                Case "[DIRETORIO]"
                    gDiretorio = lDado
                Case "[LIBERA_PGTO_POSTECIPADO]"
                    gRelUsuPostecipada = lDado
                Case "[BOOK_KIT]"
                    gRelUsuBookKit = lDado
                Case "[TARIFAVIGENTE]"
                    gTarifaVigente = lDado
                Case "[LIBERA_PREMIO_VALOR_INFORMADO]"
                    gRelUsuPrInformado = lDado
                Case "[LIBERA_PILOTO_RES_SUPREMO]" ''Apenas corretores presentes no syasmulti.ini
                    gRelPilotoSupremo = lDado
                Case "[RESIDENCIAL_SUPREMO]" ''Corretores cadastrados no gerenciador
                    gResidencialSupremo = lDado
                Case "[RESIDENCIAL_SUPREMO_TODOS]" ''Todos os corretores
                    gResidencialSupremoTodos = lDado
                Case "[PATH_SISCOTA]" ''Path dos anexos do siscota
                    gPathSiscota = lDado
                ''2016-08-11: Removido
                'Case "[LIBERA_IMPRESSAO_FTP]"       ''2016-07-12
                '    gRelUsuImpressaoFtp = lDado
                End Select
                If lDesconhecido = True Then
                    Call gpGraLog(2, "FTE0049 - Erro no arquivo de inicializa��o (tipo " & _
                                     "desconhecido para t�tulo " & lTitulo & ")." & vbLf & vbLf & _
                                     lRegSYASMULTINI & vbLf & vbLf & "O sistema ser� encerrado.")
                    GoTo SAIDA    'End   ***MOSILVA Verificar
                End If
            End Select
        End If
    Loop
    Close lNumSYASMULTINI

    gBaseLocal = True
    If gTipUsuario = e_Tipo_Usu_Corretor Or gTipP00MPED = 1 Then   'corretor
        gBaseLocal = True
        'ElseIf gTipP00MPED = 2 And PerfilAcesso <> e_Perfil_DIG_IST Then
    Else
        gBaseLocal = False
    End If


    If gTipUsuario <> e_Tipo_Usu_CorretorYasuda Then
        lsPropTabela = ""
    End If

    If gTipUsuario = e_Tipo_Usu_CorretorYasuda And gfPreenchido(lsPropTabela) = False Then
        Call gpGraLog(2, "FTE0049 - Erro no arquivo de inicializa��o (Propriet�rio " & _
                         "das tabelas n�o informado)." & vbLf & vbLf & _
                         lRegSYASMULTINI & vbLf & vbLf & "O sistema ser� encerrado.")
    End If

    '4. Finaliza a procedure para iniciar procedimento de atualiza��o do SYAS com os arquivos
    '   SYASxxnn.003. Processo iniciado apenas se o m�dulo estiver sendo executado pelo programa
    '   YASUDA.EXE.
    If gfPreenchido(gPrograma) Then
        If gPrograma = "YASUDA.EXE" Then
            'Encerra procedure para processar atualiza��es.
            GoTo SAIDA
        End If
    End If

    If gEndConfCot1 = "MULTCOT1.CXY" Then
        gEndConfCot1 = gAppPath & gEndConfCot1
    End If
    If gEndImagCot1 = "MULTCOT1.WMF" Then
        gEndImagCot1 = gAppPath & gEndImagCot1
    End If
    If gEndConfCotA = "MULTCOTA.CXY" Then
        gEndConfCotA = gAppPath & gEndConfCotA
    End If
    If gEndImagCotA = "MULTCOTA.WMF" Then
        gEndImagCotA = gAppPath & gEndImagCotA
    End If
    If gEndConfPro1 = "MULTPROP1.CXY" Then
        gEndConfPro1 = gAppPath & gEndConfPro1
    End If
    If gEndImagPro1 = "MULTPROP1.WMF" Then
        gEndImagPro1 = gAppPath & gEndImagPro1
    End If

    If gEndConfPro2 = "MULTPROP2.CXY" Then
        gEndConfPro2 = gAppPath & gEndConfPro2
    End If
    If gEndImagPro2 = "MULTPROP2.WMF" Then
        gEndImagPro2 = gAppPath & gEndImagPro2
    End If

    If gfPreenchido(gEndConfRob1) = False Then
        gEndConfRob1 = Replace(gEndConfCot1, "MULTCOT1.CXY", "") & "ROUBOCOT1.CXY"
        gEndConfRob2 = Replace(gEndConfCot1, "MULTCOT1.CXY", "") & "ROUBOCOT2.CXY"
        gEndConfRob3 = Replace(gEndConfCot1, "MULTCOT1.CXY", "") & "ROUBOCOT3.CXY"
        gEndConfRob4 = Replace(gEndConfCot1, "MULTCOT1.CXY", "") & "ROUBOCOT4.CXY"
        gEndConfRob5 = Replace(gEndConfCot1, "MULTCOT1.CXY", "") & "ROUBOCOT5.CXY"
    End If
    If gfPreenchido(gEndImagRob1) = False Then
        gEndImagRob1 = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "ROUBOCOT1.WMF"
        gEndImagRob2 = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "ROUBOCOT2.WMF"
        gEndImagRob3 = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "ROUBOCOT3.WMF"
        gEndImagRob4 = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "ROUBOCOT4.WMF"
        gEndImagRob5 = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "ROUBOCOT5.WMF"
    End If


    If gfPreenchido(gEndConfCotC) = False Then
        gEndConfCotC = Replace(gEndConfCotA, "MULTCOTA.CXY", "") & "MULTCOTC.CXY"
        gEndImagCotC = Replace(gEndImagCotA, "MULTCOTA.WMF", "") & "MULTCOTC.WMF"
    End If

    If gfPreenchido(gEndImagCotLMI) = False Then
        'gEndConfCotLMI = Replace(gEndConfCot1, "MULTCOT1.CXY", "") & "COT_LMI.CXY"
        gEndImagCotLMI = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "COT_LMI.WMF"
        gEndImagRelLMI = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "REL_LMI.WMF"
    End If

    If gfPreenchido(gEndImagESPLMI) = False Then
        gEndImagESPLMI = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "MULTLMIE.WMF"
        gEndImagESPRelLMI = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "MULTLMIR.WMF"
    End If

    '5. Verifica bases de dados.
    'a) Posiciona quantidade de bases de dados.
    lQuaBasDados = 2    '<<<<<<====== Verificar quantidade de bases de dados a verificar.
    'b) Posiciona nomes, tipos e endere�os das bases de dados.
    ReDim lNomBasDados(lQuaBasDados)
    ReDim lEndBasDados(lQuaBasDados)
    ReDim lTipBasDados(lQuaBasDados)
    lNomBasDados(1) = "P00MPED"
    lEndBasDados(1) = gEndP00MPED
    lTipBasDados(1) = gTipP00MPED      'Observar vari�vel lQuaBasDados.
    lNomBasDados(2) = "P00APOL"
    lEndBasDados(2) = gEndP00APOL
    lTipBasDados(2) = gTipP00APOL      'Observar vari�vel lQuaBasDados.
    'b) Processa verifica��o.
    If gUsuMulti = False Then
        For lI = 1 To lQuaBasDados
            Do
                '===1) Verifica tipo da base de dados.
                If Not gfPreenchido(lTipBasDados(lI)) Then
                    Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(lI) & _
                                     ": tipo da base de dados n�o preenchido." & vbLf & vbLf & _
                                     "O sistema ser� encerrado.")

                    GoTo SAIDA    'End   ***MOSILVA Verificar
                End If
                If lTipBasDados(lI) <> "1" And lTipBasDados(lI) <> "2" Then
                    Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(lI) & _
                                     ": tipo da base de dados diferente de 1 - " & lTipBasDados(lI) & "." & _
                                     vbLf & vbLf & "O sistema ser� encerrado.")
                    GoTo SAIDA    'End   ***MOSILVA Verificar
                End If
                '===2) Verifica endere�o do arquivo.
                If gfPreenchido(lEndBasDados(lI)) Then
                    Select Case UCase(lEndBasDados(lI))
                    Case "P00MULTI.MDB"
                        lEndereco = gAppPath & lEndBasDados(lI)
                        gEndP00MULTI = lEndereco
                    Case "P00MPED.MDB"
                        lEndereco = gAppPath & lEndBasDados(lI)
                        gEndP00MPED = lEndereco
                    Case "P00APOL.MDB"
                        lEndereco = gAppPath & lEndBasDados(lI)
                        gEndP00APOL = lEndereco
                    Case Else
                        lEndereco = lEndBasDados(lI)
                    End Select
                Else
                    lEndereco = gAppPath & lNomBasDados(lI) & ".MDB"
                End If
                If UCase(lNomBasDados(lI)) = "P0044000" And gTipUsuario = 1 Then
                    Exit Do
                End If
                If gBaseLocal Then
                    If Not gfPreenchido(Dir(lEndereco)) Then
                        Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(lI) & _
                                         ": base de dados n�o localizada." & vbLf & vbLf & lEndereco & _
                                         vbLf & vbLf & "O sistema ser� encerrado.")
                        GoTo SAIDA    'End   ***MOSILVA Verificar
                    End If
                End If
                'lPosicao = InStr(lEndereco, lNomBasDados(lI) & ".MDB")
                'If lPosicao = 0 Then
                '    Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(lI) & _
                     '                     ": nome da base de dados n�o consta no endere�o." & vbLf & vbLf & _
                     '                     lEndereco & vbLf & vbLf & "O sistema ser� encerrado.")
                '    End
                'End If
                Exit Do
            Loop
        Next lI
    End If

    '5. Verifica pasta remessa.
    If Not gfPreenchido(gEndRemessa) Then
        gEndRemessa = gAppPath & "REMESSA"
    End If
    If Right$(gEndRemessa, 1) = "\" Then
        gEndRemessa = Left$(gEndRemessa, Len(gEndRemessa) - 1)
    End If
    If Not gfPreenchido(Dir(gEndRemessa, vbDirectory)) Then
        On Error Resume Next
        MkDir (gEndRemessa)
        If Err <> 0 Then
            Call gpGraLog(0, "FTE0015 - Erro ao criar pasta " & gEndRemessa & _
                           " (pasta de remessa obter dados de inicializa��o do sistema).")
        End If
    End If
    gEndRemessa = gEndRemessa & "\"

    Call mfObtSYASPRODINI

    If lEndSiscad <> "" Then
        gEndExeSiscad = lEndSiscad
    End If

    If lEndP47V001 <> "" Then
        gEndExeP47V001 = lEndP47V001
    End If

Retorno:
    pPropTabela = lsPropTabela
SAIDA:
    Exit Sub

End Sub

Private Function mfObtSYASPRODINI() As Boolean
'Fun��o: obt�m dados do arquivo SYASPROD.INI.

    Dim lDado As String    'Dado da linha de registro.
    Dim lEndSyasprodINI As String   'Endere�o do arquivo SYASPROD.INI.
    Dim lNumSyasprodINI As Integer  'Utilizado com FreeFile para o arquivo SYASPROD.INI.
    Dim lPosicao As Integer    'Utilizado com InStr.
    Dim lRegSyasprodINI As String   'Registro do arquivo SYASPROD.INI.
    Dim lTamanho As Integer    'Tamanho.
    Dim lTipo As String    'Tipo de linha do arquivo de inicializa��o.
    Dim lTitulo As String    'Linha de t�tulo.
    Dim lEndBasDados() As String  'Bases de dados: endere�os.
    Dim lEndereco As String    'Endereco.
    Dim lI As Integer             'Utilizado com For...Next.
    Dim lNomBasDados() As String  'Bases de dados: nomes.
    Dim lQuaBasDados As Byte  'Bases de dados: quantidade.
    Dim lTipBasDados() As String  'Bases de dados: tipos.

'1. Abre arquivo SYASPROD.INI.
    lEndSyasprodINI = gAppPath & "SYASPROD.INI"
    If Not gfPreenchido(Dir(lEndSyasprodINI)) Then
        Call gpGraLog(2, "FTE0025 - Arquivo n�o localizado: " & lEndSyasprodINI & _
                       " (obter dados de inicializa��o do sistema)." & vbLf & vbLf & _
                         "O sistema ser� encerrado.")
        Exit Function    'End   ***MOSILVA Verificar
    End If
    lNumSyasprodINI = FreeFile
    On Error Resume Next
    Open lEndSyasprodINI For Input Access Read As lNumSyasprodINI
    If Err <> 0 Then
        Call gpGraLog(2, "FTE0026 - Erro ao abrir arquivo " & lEndSyasprodINI & ": " & Err & " - " & _
                         ERROR & vbLf & vbLf & "O sistema ser� encerrado.")
        On Error GoTo 0
        Exit Function    'End   ***MOSILVA Verificar
    End If
    On Error GoTo 0

    '2. Trata registros do arquivo de inicializa��o.
    Do While Not EOF(lNumSyasprodINI)
        Line Input #lNumSyasprodINI, lRegSyasprodINI
        If gfPreenchido(lRegSyasprodINI) Then
            Select Case Left$(lRegSyasprodINI, 1)
            Case ";"    'Linha de observa��o.
                'Nada a processar - linha de observa��o.
            Case "["    'Linha de t�tulo.
                lTitulo = Trim$(UCase$(lRegSyasprodINI))
            Case Else   'Linha de tipo.
                lRegSyasprodINI = Trim$(lRegSyasprodINI)
                lPosicao = InStr(lRegSyasprodINI, "=")
                lTamanho = Len(lRegSyasprodINI)
                If lPosicao = 0 Then
                    Call gpGraLog(2, "FTE0046 - Erro no arquivo de inicializa��o (sinal de igual " & _
                                     "n�o localizado no registro)." & vbLf & vbLf & _
                                     lRegSyasprodINI & vbLf & vbLf & "O sistema ser� encerrado.")
                    Exit Function    'End   ***MOSILVA Verificar
                End If
                If lPosicao = 1 Then
                    Call gpGraLog(2, "FTE0046 - Erro no arquivo de inicializa��o (sinal de igual " & _
                                     "na primeira posi��o do registro)." & vbLf & vbLf & _
                                     lRegSyasprodINI & vbLf & vbLf & "O sistema ser� encerrado.")
                    Exit Function    'End   ***MOSILVA Verificar
                End If
                If lPosicao = lTamanho Then
                    Call gpGraLog(2, "FTE0046 - Erro no arquivo de inicializa��o (sinal de igual " & _
                                     "na �ltima posi��o do registro)." & vbLf & vbLf & _
                                     lRegSyasprodINI & vbLf & vbLf & "O sistema ser� encerrado.")
                    Exit Function    'End   ***MOSILVA Verificar
                End If
                lTipo = Trim$(UCase$(Left$(lRegSyasprodINI, lPosicao - 1)))
                lDado = Trim$(Mid$(lRegSyasprodINI, lPosicao + 1))
                Select Case lTitulo
                Case "[L28V009A]"
                    Select Case lTipo
                    Case "CONFIG"
                        gEndConfDeb = lDado
                    Case "IMAGEM"
                        gEndImagDeb = lDado
                    End Select
                Case "[P0043700]"
                    Select Case lTipo
                    Case "DATABASE"
                        If gfPreenchido(lDado) Then
                            lTamanho = Len(lDado)
                            If lTamanho = 12 Then
                                gEndServidor = gAppPath
                            Else
                                gEndServidor = Left$(lDado, lTamanho - 12)
                            End If
                        End If
                    End Select
                Case "[SISCAD]"
                    gEndExeSiscad = lDado
                Case "[P47V001]"
                    gEndExeP47V001 = lDado
                Case "[P0043000]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipP0043000 = lDado
                    Case "DATABASE"
                        gEndP0043000 = lDado
                        If UCase(gEndP0043000) = "P0043000.MDB" Then
                            gEndP0043000 = gAppPath & "P0043000.MDB"
                        End If
                    End Select
                Case "[P0044700]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipP0044700 = lDado
                    Case "DATABASE"
                        gEndP0044700 = lDado
                        If UCase(gEndP0044700) = "P0044700.MDB" Then
                            gEndP0044700 = gAppPath & "P0044700.MDB"
                        End If
                    End Select
                Case "[P0045200]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipP0045200 = lDado
                    Case "DATABASE"
                        gEndP0045200 = lDado
                        If UCase(gEndP0045200) = "P0045200.MDB" Then
                            gEndP0045200 = gAppPath & "P0045200.MDB"
                        End If
                    End Select
                Case "[P0045300]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipP0045300 = lDado
                    Case "DATABASE"
                        gEndP0045300 = lDado
                        If UCase(gEndP0045300) = "P0045300.MDB" Then
                            gEndP0045300 = gAppPath & "P0045300.MDB"
                        End If
                    End Select
                Case "[P0045400]"
                    Select Case lTipo
                    Case "TIPO"
                        gTipP0045400 = lDado
                    Case "DATABASE"
                        gEndP0045400 = lDado
                        If UCase(gEndP0045400) = "P0045400.MDB" Then
                            gEndP0045400 = gAppPath & "P0045400.MDB"
                        End If
                    End Select
                Case "[ATUALIZACAO]"
                    gEndSYASUpDt = lDado

                    If gfPreenchido(gEndSYASUpDt) Then
                        lTamanho = Len(gEndSYASUpDt)
                        If lTamanho = 5 Then
                            gEndSYASUpDt = gAppPath
                        Else
                            gEndSYASUpDt = Left$(gEndSYASUpDt, lTamanho - 5)
                        End If
                    End If
                Case "[TIPOTRANSMISSAO]"
                    If Trim(UCase(lDado)) = "OFFLINE" Then
                        gTipoTransmDLL = ETransmDLL_OFFLine
                    Else
                        gTipoTransmDLL = ETransmDLL_ONLine
                    End If
                End Select
            End Select
        End If
    Loop
    Close lNumSyasprodINI
    If gPrograma = "D046O003" Then
        mfObtSYASPRODINI = True
        Exit Function
    End If
    If gfPreenchido(gEndP0044700) Then
        lTamanho = Len(gEndP0044700)
        If lTamanho = 12 Then
            lEndereco = ""
        Else
            lEndereco = Left$(gEndP0044700, lTamanho - 12)
        End If
    Else
        lEndereco = gAppPath
    End If

    If Not gfPreenchido(gEndP0045200) Then
        gEndP0045200 = lEndereco & "P0045200.MDB"
        gTipP0045200 = "1"
    End If
    If Not gfPreenchido(gEndP0045300) Then
        gEndP0045300 = lEndereco & "P0045300.MDB"
        gTipP0045300 = "1"
    End If

    If Not gfPreenchido(gEndP0045400) Then
        gEndP0045400 = lEndereco & "P0045400.MDB"
        gTipP0045400 = "1"
    End If

    '5. Verifica bases de dados.
    'a) Posiciona quantidade de bases de dados.
    lQuaBasDados = 2    '<<<<<<====== Verificar quantidade de bases de dados a verificar.
    'b) Posiciona nomes, tipos e endere�os das bases de dados.
    ReDim lNomBasDados(lQuaBasDados)
    ReDim lEndBasDados(lQuaBasDados)
    ReDim lTipBasDados(lQuaBasDados)
    lNomBasDados(1) = "P0044700"
    lEndBasDados(1) = gEndP0044700
    lTipBasDados(1) = gTipP0044700  'Observar vari�vel lQuaBasDados.
    lNomBasDados(2) = "P0043000"
    lEndBasDados(2) = gEndP0043000
    lTipBasDados(2) = gTipP0043000      'Observar vari�vel lQuaBasDados.



    'lNomBasDados(3) = "P0045200"
    'lEndBasDados(3) = gEndP0045200
    'lTipBasDados(3) = gTipP0045200      'Observar vari�vel lQuaBasDados.
    'lNomBasDados(4) = "P0045300"
    'lEndBasDados(4) = gEndP0045300
    'lTipBasDados(4) = gTipP0045300      'Observar vari�vel lQuaBasDados.


    'b) Processa verifica��o.
    If gUsuMulti = False Then
        For lI = 1 To lQuaBasDados
            Do
                '===1) Verifica tipo da base de dados.
                If Not gfPreenchido(lTipBasDados(lI)) Then
                    Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(lI) & _
                                     ": tipo da base de dados n�o preenchido." & vbLf & vbLf & _
                                     "O sistema ser� encerrado.")

                    Exit Function    'End   ***MOSILVA Verificar
                End If
                If lTipBasDados(lI) <> "1" And lTipBasDados(lI) <> "2" Then
                    Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(lI) & _
                                     ": tipo da base de dados diferente de 1 - " & lTipBasDados(lI) & "." & _
                                     vbLf & vbLf & "O sistema ser� encerrado.")
                    Exit Function    'End   ***MOSILVA Verificar
                End If
                '===2) Verifica endere�o do arquivo.
                If gfPreenchido(lEndBasDados(lI)) Then
                    Select Case UCase(lEndBasDados(lI))
                    Case "P0044700.MDB"
                        lEndereco = gAppPath & lEndBasDados(lI)
                        gEndP0044700 = lEndereco
                    Case "P0043000.MDB"
                        lEndereco = gAppPath & lEndBasDados(lI)
                        gEndP0043000 = lEndereco
                    Case "P0045200.MDB"
                        lEndereco = gAppPath & lEndBasDados(lI)
                        gEndP0045200 = lEndereco
                    Case "P0045300.MDB"
                        lEndereco = gAppPath & lEndBasDados(lI)
                        gEndP0045300 = lEndereco
                    Case "P0045400.MDB"
                        lEndereco = gAppPath & lEndBasDados(lI)
                        gEndP0045400 = lEndereco
                    Case Else
                        lEndereco = lEndBasDados(lI)
                    End Select
                Else
                    lEndereco = gAppPath & lNomBasDados(lI) & ".MDB"
                End If
                If UCase(lNomBasDados(lI)) = "P0044000" And gTipUsuario = 1 Then
                    Exit Do
                End If
                If InStr(UCase(lEndereco), ".MDB") = 0 Then
                    If Not gfPreenchido(Dir(lEndereco)) Then
                        Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(lI) & _
                                         ": base de dados n�o localizada." & vbLf & vbLf & lEndereco & _
                                         vbLf & vbLf & "O sistema ser� encerrado.")
                        Exit Function    'End   ***MOSILVA Verificar
                    End If
                End If
                'lPosicao = InStr(lEndereco, lNomBasDados(lI) & ".MDB")
                'If lPosicao = 0 Then
                '    Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(lI) & _
                     '                     ": nome da base de dados n�o consta no endere�o." & vbLf & vbLf & _
                     '                     lEndereco & vbLf & vbLf & "O sistema ser� encerrado.")
                '    End
                'End If
                Exit Do
            Loop
        Next lI
    End If

End Function
Public Function gfTransmissaoInterna(ByVal pTipoTransmissao As Integer, _
                                     ByRef pobjstcPed As stcA46V702B, _
                                     ByRef pMensagem As String, _
                                     Optional pCnSQL2008 As ADODB.Connection = Nothing, _
                                     Optional pRetornoEhErro As Boolean = True) As Boolean

    Dim lObjstcPed As stcA46V702B
    Dim lobjPedido As clsA46V702A
    Dim lobjstcPedClau As stcA46V703B
    Dim lobjPedClau As clsA46V703A
    Dim lobjstcPedCob As stcA46V704B
    Dim lobjPedCob As clsA46V704A
    Dim lobjstcPedCor As stcA46V705B
    Dim lobjPedCor As clsA46V705A
    Dim lobjstcPedCos As stcA46V706B
    Dim lobjPedCos As clsA46V706A

    Dim lobjPedEnder As clsA46V707A
    Dim lobjstcPedLoc As stcA46V708B
    Dim lobjPedLoc As clsA46V708A
    Dim lobjstcPedJuros As stcA46V711B
    Dim lobjPedJuros As clsA46V711A

    Dim lobjCtrlEmis As clsA46V712A
    Dim lObjstcPedVenc As stcA46V713B
    Dim lObjPedVenc As clsA46V713A

    Dim lobjPedObs As clsA46V714A
    Dim lobjstcPedLocBem As stcA46V715B
    Dim lobjPedLocBem As clsA46V715A
    Dim lobjstcPedLocSeg As stcA46V716B
    Dim lobjPedLocSeg As clsA46V716A
    Dim lobjPedInsp As clsA46V717A
    Dim lobjstcPedFat As stcA46V720B
    Dim lobjPedFat As clsA46V720A
    Dim lobjstcPedEspec As stcA46V726B
    Dim lobjPedEspec As clsA46V726A

    Dim lobjstcPedLMI As stcA46V732B
    Dim lobjPedPedLMI As clsA46V732A


    Dim lobjTrans As clsA46V718A
    Dim lobjstcTrans As stcA46V718B

    Dim lobjstcPedDizer As stcA46V722B
    Dim lobjPedPedDizer As clsA46V722A

    Dim lobjstcPedCanal As stcA46V733B
    Dim lobjPedCanal As clsA46V733A

    Dim lobjPedPessoa As New clsA46V735A
    Dim lobjstcPedPessoa As New stcA46V735B

    Dim lobjPedTelefones As New clsA46V736A
    Dim lobjstcPedTelefones As New stcA46V736B

    Dim lobjPedFranquia As New clsA46V738A
    Dim lobjstcPedFranquia As New stcA46V738B

    Dim lobjTexto As clsA46V739A
    Dim lobjstcTexto As stcA46V739B

    Dim lobjPedLocBicicleta As clsA46V744A
    Dim lobjstcPedLocBicicleta As stcA46V744B

    Dim objLiberacaoSerasa As clsA46V746A

    Dim lobjPedAllRisks As clsA46V747A
    Dim lobjstcPedAllRisks As stcA46V747B
    
    Dim lobjPedObrasArte As clsA46V748A
    Dim lobjstcPedObrasArte As stcA46V748B

    Dim lConexaoSQL As ADODB.Connection
    Dim lExec As String
    Dim lPropTabela As String
    Dim lNomeTabImport As String

    On Error GoTo ControleErro

    lPropTabela = "RamosDiversos.dbo."
    If pTipoTransmissao = 1 Then    'MOSILVA: Questionar Thais!!
        lNomeTabImport = "FTP_"
    Else
        lNomeTabImport = ""
    End If

    'Abre a base SQL2000 de Produ��o (LPX em caso de teste Abre MIX)
    If Not gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lConexaoSQL, pMensagem) Then
        GoTo ControleErro
    End If

    Set lobjPedido = New clsA46V702A: Set lobjPedido.ConexaoBD = lConexaoSQL
    Set lobjPedido.M46V999 = Me

    Set lObjstcPed = New stcA46V702B
    Set lObjstcPed.M46V999 = Me

    lObjstcPed.nosso_numero = pobjstcPed.nosso_numero
    lObjstcPed.Cod_Unicidade = -1

    If Not lobjPedido.mfCarregarColecao(lObjstcPed, lPropTabela, lNomeTabImport, False) Then
        pMensagem = lobjPedido.UltimoErro
        GoTo ControleErro
    End If

    If lobjPedido.Colecao.Count <> 0 Then
        Set lObjstcPed = lobjPedido.Colecao.Item(1)
        M46V111.gNum_Pi = pobjstcPed.NUM_PI
        Call gpExcluiRegistros(pobjstcPed.nosso_numero, 0, lPropTabela, lNomeTabImport, pMensagem)
    End If

    Set lobjPedido = New clsA46V702A: Set lobjPedido.ConexaoBD = lConexaoSQL
    Set lobjPedido.M46V999 = Me

    'Ajusta os campos
    '''thais - ver santander
    
    If pobjstcPed.Num_Cta_Bas_Cor <> 9999999 And pobjstcPed.Cod_Agen_Bas_Cor <> 9999 And pobjstcPed.Cod_Bco_Cobr <> 999 Then
        pobjstcPed.Num_Cta_Bas_Cor = Left(Format(pobjstcPed.Num_Cta_Bas_Cor, "000000000"), 8)
    End If

    '
    If Not lobjPedido.mfIncluir(pobjstcPed, lPropTabela, lNomeTabImport, tagDynaTrace) Then
        pMensagem = lobjPedido.UltimoErro
        GoTo ControleErro
    End If

    Set lobjPedCor = New clsA46V705A
    Set lobjPedCor.M46V999 = Me

    Set lobjPedCor.ConexaoBD = lConexaoSQL
    For Each lobjstcPedCor In pobjstcPed.ColecaoPedCorr
        Set lobjstcPedCor.ObjPai = pobjstcPed
        If Not lobjPedCor.mfIncluir(lobjstcPedCor, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedCor.UltimoErro
            GoTo ControleErro
        End If
    Next lobjstcPedCor

    Set lobjPedCos = New clsA46V706A
    Set lobjPedCos.M46V999 = Me
    Set lobjPedCos.ConexaoBD = lConexaoSQL

    For Each lobjstcPedCos In pobjstcPed.ColecaoPedCoss
        Set lobjstcPedCos.ObjPai = pobjstcPed
        If Not lobjPedCos.mfIncluir(lobjstcPedCos, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedCos.UltimoErro
            GoTo ControleErro
        End If
    Next lobjstcPedCos

    Set lobjPedEnder = New clsA46V707A
    Set lobjPedEnder.M46V999 = Me

    Set lobjPedEnder.ConexaoBD = lConexaoSQL
    pobjstcPed.Cod_Unicidade = pobjstcPed.Cod_Unicidade
    Set pobjstcPed.objstcPedEndereco.ObjPai = pobjstcPed
    If Not lobjPedEnder.mfIncluir(pobjstcPed.objstcPedEndereco, lPropTabela, lNomeTabImport, tagDynaTrace) Then
        pMensagem = lobjPedEnder.UltimoErro
        GoTo ControleErro
    End If

    Set lobjPedPedLMI = New clsA46V732A
    Set lobjPedPedLMI.M46V999 = Me
    Set lobjPedPedLMI.ConexaoBD = lConexaoSQL

    For Each lobjstcPedLMI In pobjstcPed.ColecaoLMI
        Set lobjstcPedLMI.ObjPai = pobjstcPed
        If Not lobjPedPedLMI.mfIncluir(lobjstcPedLMI, lPropTabela, lNomeTabImport) Then
            pMensagem = lobjPedPedLMI.UltimoErro
            GoTo ControleErro
        End If
    Next lobjstcPedLMI


    Set lobjPedPedDizer = New clsA46V722A
    Set lobjPedPedDizer.M46V999 = Me
    Set lobjPedPedDizer.ConexaoBD = lConexaoSQL

    For Each lobjstcPedDizer In pobjstcPed.ColecaoTextoEndo
        Set lobjstcPedDizer.ObjPai = pobjstcPed
        If Not lobjPedPedDizer.mfIncluir(lobjstcPedDizer, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedPedDizer.UltimoErro
            GoTo ControleErro
        End If
    Next lobjstcPedDizer



    Set lobjCtrlEmis = New clsA46V712A
    Set lobjCtrlEmis.M46V999 = Me
    Set lobjCtrlEmis.ConexaoBD = lConexaoSQL

    Set pobjstcPed.objstcCtrlEmis.ObjPai = pobjstcPed
    If Not lobjCtrlEmis.mfIncluir(pobjstcPed.objstcCtrlEmis, lPropTabela, lNomeTabImport, tagDynaTrace) Then
        pMensagem = lobjCtrlEmis.UltimoErro
        GoTo ControleErro
    End If

    Set lobjPedJuros = New clsA46V711A
    Set lobjPedJuros.M46V999 = Me

    Set lobjPedJuros.ConexaoBD = lConexaoSQL
    For Each lobjstcPedJuros In pobjstcPed.ColecaoPedJuros
        Set lobjstcPedJuros.ObjPai = pobjstcPed
        If Not lobjPedJuros.mfIncluir(lobjstcPedJuros, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedJuros.UltimoErro
            GoTo ControleErro
        End If
    Next lobjstcPedJuros

    Set lObjPedVenc = New clsA46V713A
    Set lObjPedVenc.M46V999 = Me
    Set lObjPedVenc.ConexaoBD = lConexaoSQL

    For Each lObjstcPedVenc In pobjstcPed.ColecaoPedVenc
        Set lObjstcPedVenc.ObjPai = pobjstcPed
        If Not lObjPedVenc.mfIncluir(lObjstcPedVenc, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lObjPedVenc.UltimoErro
            GoTo ControleErro
        End If
    Next lObjstcPedVenc

    Set lobjPedEspec = New clsA46V726A
    Set lobjPedEspec.M46V999 = Me

    Set lobjPedEspec.ConexaoBD = lConexaoSQL
    For Each lobjstcPedEspec In pobjstcPed.ColecaoTextoEspec
        Set lobjstcPedEspec.ObjPai = pobjstcPed
        If Not lobjPedEspec.mfIncluir(lobjstcPedEspec, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedEspec.UltimoErro
            GoTo ControleErro
        End If
    Next lobjstcPedEspec

    Set lobjPedTelefones = New clsA46V736A
    Set lobjPedTelefones.M46V999 = Me
    Set lobjstcPedTelefones.M46V999 = Me

    Set lobjPedTelefones.ConexaoBD = lConexaoSQL
    For Each lobjstcPedTelefones In pobjstcPed.ColecaoTelefones
        Set lobjstcPedTelefones.ObjPai = pobjstcPed
        If Not lobjPedTelefones.mfIncluir(lobjstcPedTelefones, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedTelefones.UltimoErro
            GoTo ControleErro
        End If
    Next lobjstcPedTelefones

    Set objLiberacaoSerasa = New clsA46V746A
    Set objLiberacaoSerasa.M46V999 = Me
    Set objLiberacaoSerasa.ConexaoBD = lConexaoSQL
    
    If Not objLiberacaoSerasa.mfIncluir(pobjstcPed.objstcLiberacaoSerasa, lPropTabela, lNomeTabImport, tagDynaTrace) Then
        pMensagem = objLiberacaoSerasa.UltimoErro
        GoTo ControleErro
    End If

    Set lobjPedLoc = New clsA46V708A
    Set lobjPedLoc.M46V999 = Me

    Set lobjPedLoc.ConexaoBD = lConexaoSQL

    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        Set lobjstcPedLoc.ObjPai = pobjstcPed
        If Not lobjPedLoc.mfIncluir(lobjstcPedLoc, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedLoc.UltimoErro
            GoTo ControleErro
        End If

        Set lobjPedClau = New clsA46V703A
        Set lobjPedClau.M46V999 = Me

        Set lobjPedClau.ConexaoBD = lConexaoSQL

        For Each lobjstcPedClau In lobjstcPedLoc.ColecaoPedClau
            Set lobjstcPedClau.ObjPai = lobjstcPedLoc
            If Not lobjPedClau.mfIncluir(lobjstcPedClau, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjPedClau.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedClau

        Set lobjPedCob = New clsA46V704A
        Set lobjPedCob.M46V999 = Me

        Set lobjPedFranquia = New clsA46V738A
        Set lobjPedFranquia.M46V999 = Me
        Set lobjPedFranquia.ConexaoBD = lConexaoSQL

        Set lobjPedCob.ConexaoBD = lConexaoSQL
        For Each lobjstcPedCob In lobjstcPedLoc.ColecaoPedCobert
            Set lobjstcPedCob.ObjPai = lobjstcPedLoc
            If Not lobjPedCob.mfIncluir(lobjstcPedCob, lPropTabela, lNomeTabImport, False, False, tagDynaTrace) Then
                pMensagem = lobjPedCob.UltimoErro
                GoTo ControleErro
            End If
            
            'Projeto Multiplo Franquia
            For Each lobjstcPedFranquia In lobjstcPedCob.ColecaoNovaFranquia
                Set lobjstcPedFranquia.ObjPai = lobjstcPedCob
                If lobjPedFranquia.mfIncluir(lobjstcPedFranquia, lPropTabela, lNomeTabImport, tagDynaTrace) = False Then
                    pMensagem = lobjPedFranquia.UltimoErro
                    GoTo ControleErro
                End If
            Next
        Next lobjstcPedCob

        Set lobjPedObs = New clsA46V714A
        Set lobjPedObs.M46V999 = Me

        Set lobjPedObs.ConexaoBD = lConexaoSQL
        If gfPreenchido(lobjstcPedLoc.objstcPedObs.Observacao) Then
            Set lobjstcPedLoc.objstcPedObs.ObjPai = lobjstcPedLoc
            If Not lobjPedObs.mfIncluir(lobjstcPedLoc.objstcPedObs, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjPedObs.UltimoErro
                GoTo ControleErro
            End If
        End If

        Set lobjPedLocBem = New clsA46V715A
        Set lobjPedLocBem.M46V999 = Me

        Set lobjPedLocBem.ConexaoBD = lConexaoSQL
        For Each lobjstcPedLocBem In lobjstcPedLoc.ColecaoPedLocBem
            Set lobjstcPedLocBem.ObjPai = lobjstcPedLoc
            If Not lobjPedLocBem.mfIncluir(lobjstcPedLocBem, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjPedLocBem.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedLocBem

        Set lobjPedLocBicicleta = New clsA46V744A
        Set lobjPedLocBicicleta.M46V999 = Me

        Set lobjPedLocBicicleta.ConexaoBD = lConexaoSQL
        For Each lobjstcPedLocBicicleta In lobjstcPedLoc.ColecaoPedLocBicicleta
            Set lobjstcPedLocBicicleta.ObjPai = lobjstcPedLoc
            If Not lobjPedLocBicicleta.mfIncluir(lobjstcPedLocBicicleta, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjPedLocBicicleta.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedLocBicicleta

        Set lobjPedLocSeg = New clsA46V716A
        Set lobjPedLocSeg.M46V999 = Me

        Set lobjPedLocSeg.ConexaoBD = lConexaoSQL
        For Each lobjstcPedLocSeg In lobjstcPedLoc.ColecaoPedSeguro
            Set lobjstcPedLocSeg.ObjPai = lobjstcPedLoc
            If Not lobjPedLocSeg.mfIncluir(lobjstcPedLocSeg, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjPedLocSeg.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedLocSeg


        Set lobjPedInsp = New clsA46V717A
        Set lobjPedInsp.M46V999 = Me

        Set lobjPedInsp.ConexaoBD = lConexaoSQL
        Set lobjstcPedLoc.objstcPedInsp.ObjPai = lobjstcPedLoc
        If Not lobjPedInsp.mfIncluir(lobjstcPedLoc.objstcPedInsp, lPropTabela, lNomeTabImport, tagDynaTrace) Then
            pMensagem = lobjPedInsp.UltimoErro
            GoTo ControleErro
        End If

        Set lobjPedFat = New clsA46V720A
        Set lobjPedFat.M46V999 = Me

        Set lobjPedFat.ConexaoBD = lConexaoSQL
        For Each lobjstcPedFat In lobjstcPedLoc.ColecaoPedFator
            Set lobjstcPedFat.ObjPai = lobjstcPedLoc
            If Not lobjPedFat.mfIncluir(lobjstcPedFat, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjPedFat.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedFat

        Set lobjPedCanal = New clsA46V733A
        Set lobjPedCanal.M46V999 = Me

        Set lobjPedCanal.ConexaoBD = lConexaoSQL
        For Each lobjstcPedCanal In pobjstcPed.ColecaoProlabore
            Set lobjstcPedCanal.ObjPai = pobjstcPed
            If Not lobjPedCanal.mfIncluir(lobjstcPedCanal, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjPedCanal.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedCanal


        Set lobjPedCanal = New clsA46V733A
        Set lobjPedCanal.M46V999 = Me

        Set lobjPedCanal.ConexaoBD = lConexaoSQL
        For Each lobjstcPedCanal In pobjstcPed.ColecaoAgenciador
            Set lobjstcPedCanal.ObjPai = pobjstcPed
            If Not lobjPedCanal.mfIncluir(lobjstcPedCanal, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjPedCanal.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedCanal

        Set lobjPedPessoa = New clsA46V735A
        Set lobjPedPessoa.M46V999 = Me

        Set lobjPedPessoa.ConexaoBD = lConexaoSQL

        For Each lobjstcPedPessoa In lobjstcPedLoc.ColecaoBeneficiario
            Set lobjstcPedPessoa.ObjPai = lobjstcPedLoc
            lobjstcPedPessoa.Tip_Loc_Pessoa = e_TipLocPessoa_Beneficiario
            If Not lobjPedPessoa.mfIncluir(lobjstcPedPessoa, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjPedPessoa.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedPessoa


        Set lobjstcPedPessoa = lobjstcPedLoc.objstcPedInquilino
        If gfPreenchido(lobjstcPedPessoa.Num_CPF_CNPJ) And Trim(lobjstcPedPessoa.Num_CPF_CNPJ) <> "0" Then
            Set lobjstcPedPessoa.ObjPai = lobjstcPedLoc
            lobjstcPedPessoa.Tip_Loc_Pessoa = e_TipLocPessoa_Inquilino
            If Not lobjPedPessoa.mfIncluir(lobjstcPedPessoa, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjPedPessoa.UltimoErro
                GoTo ControleErro
            End If
        End If

        Set lobjstcPedPessoa = lobjstcPedLoc.objstcPedProprietario
        If gfPreenchido(lobjstcPedPessoa.Num_CPF_CNPJ) And Trim(lobjstcPedPessoa.Num_CPF_CNPJ) <> "0" Then
            Set lobjstcPedPessoa.ObjPai = lobjstcPedLoc
            lobjstcPedPessoa.Tip_Loc_Pessoa = e_TipLocPessoa_Proprietario
            If Not lobjPedPessoa.mfIncluir(lobjstcPedPessoa, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjPedPessoa.UltimoErro
                GoTo ControleErro
            End If
        End If


        'TAB_PED_TEXTOS
        Set lobjTexto = New clsA46V739A
        Set lobjstcTexto = New stcA46V739B
        Set lobjTexto.M46V999 = Me
        
        For Each lobjstcTexto In lobjstcPedLoc.ColecaoTextos
            Set lobjstcTexto.ObjPai = lobjstcPedLoc
            If Not lobjTexto.mfIncluir(lobjstcTexto, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjTexto.UltimoErro
                GoTo ControleErro
            End If
        Next

        'AllRisks
        Set lobjPedAllRisks = New clsA46V747A
        Set lobjPedAllRisks.M46V999 = Me

        Set lobjPedAllRisks.ConexaoBD = lConexaoSQL
        For Each lobjstcPedAllRisks In lobjstcPedLoc.ColecaoAllRisks
            Set lobjstcPedAllRisks.ObjPai = lobjstcPedLoc
            If Not lobjPedAllRisks.mfIncluir(lobjstcPedAllRisks, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjPedAllRisks.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedAllRisks
        
        'ObrasArte
        Set lobjPedObrasArte = New clsA46V748A
        Set lobjPedObrasArte.M46V999 = Me

        Set lobjPedObrasArte.ConexaoBD = lConexaoSQL
        For Each lobjstcPedObrasArte In lobjstcPedLoc.ColecaoObrasArte
            Set lobjstcPedObrasArte.ObjPai = lobjstcPedLoc
            If Not lobjPedObrasArte.mfIncluir(lobjstcPedObrasArte, lPropTabela, lNomeTabImport, tagDynaTrace) Then
                pMensagem = lobjPedObrasArte.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedObrasArte

    Next lobjstcPedLoc
    
    gfTransmissaoInterna = True

SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    gfTransmissaoInterna = False

    If Not gfPreenchido(pMensagem) Then
        pMensagem = Err.Description
    End If
    GoTo SAIDA

FechaObj:
    Set lObjstcPed = Nothing
    Set lobjPedido = Nothing
    Set lobjstcPedClau = Nothing
    Set lobjPedClau = Nothing
    Set lobjstcPedCob = Nothing
    Set lobjPedCob = Nothing
    Set lobjstcPedCor = Nothing
    Set lobjPedCor = Nothing
    Set lobjstcPedCos = Nothing
    Set lobjPedCos = Nothing
    Set lobjPedEnder = Nothing
    Set lobjstcPedLoc = Nothing
    Set lobjPedLoc = Nothing
    Set lobjstcPedJuros = Nothing
    Set lobjPedJuros = Nothing
    Set lobjCtrlEmis = Nothing
    Set lObjstcPedVenc = Nothing
    Set lObjPedVenc = Nothing
    Set lobjPedObs = Nothing
    Set lobjstcPedLocBem = Nothing
    Set lobjPedLocBem = Nothing
    Set lobjstcPedLocSeg = Nothing
    Set lobjPedLocSeg = Nothing
    Set lobjPedInsp = Nothing
    Set lobjstcPedFat = Nothing
    Set lobjPedFat = Nothing
    Set lobjstcPedEspec = Nothing
    Set lobjPedEspec = Nothing
    Set lobjTrans = Nothing
    Set lobjstcTrans = Nothing
    Set lobjstcPedCanal = Nothing
    Set lobjPedCanal = Nothing
    Set lobjstcPedTelefones = Nothing
    Set lobjPedTelefones = Nothing
    Set lobjstcPedLMI = Nothing
    Set lobjPedPedLMI = Nothing
    Set lobjstcPedDizer = Nothing
    Set lobjPedPedDizer = Nothing
    Set lobjPedPessoa = Nothing
    Set lobjstcPedPessoa = Nothing
    Set lobjPedFranquia = Nothing
    Set lobjstcPedFranquia = Nothing
    Set lobjTexto = Nothing
    Set lobjstcTexto = Nothing
    Set objLiberacaoSerasa = Nothing

    Set lConexaoSQL = Nothing
    Return
End Function

Private Sub Class_Initialize()
    Set mConexaoSQL = New ADODB.Connection
    Set mClsM46V111 = Nothing
    Set mClsM46O001 = Nothing
    Set mClsM46V103 = Nothing
    Set gRst_Tbm_De_Para = New ADODB.Recordset
    Set gRst_Tab_Cobert = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
    Call Me.gpFecha2(mConexaoSQL)
    Set mConexaoSQL = Nothing
    Set mClsM46V111 = Nothing
    Set mClsM46O001 = Nothing
    Set M46V103 = Nothing
    Set mClsMDynaTrace = Nothing
    Set gRst_Tbm_De_Para = Nothing
    Set gRst_Tab_Cobert = Nothing
End Sub

Public Property Get TarifaAntiga() As Long
    TarifaAntiga = 921
End Property

Public Property Get TarifaMigracaoMaritima() As Long
    TarifaMigracaoMaritima = 1076
End Property
Public Property Get TarifaPrecificacaoMaritima() As Long
    TarifaPrecificacaoMaritima = 1080
End Property
Public Property Get M99V001() As clsM99V001
    Set M99V001 = M46V111.M99V001
End Property

Public Property Set M46V111(pClsM46V111 As clsM46V111)
    Set mClsM46V111 = pClsM46V111
End Property

Public Property Get M46V111() As clsM46V111
    If mClsM46V111 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V111 n�o atribu�do para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V111 j�
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46V111 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46V111 = mClsM46V111
End Property

Public Property Set M46O001(pClsM46O001 As clsM46O001)
    Set mClsM46O001 = pClsM46O001
End Property

Public Property Get M46O001() As clsM46O001
    If mClsM46O001 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46O001 n�o atribu�do para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46O001 j�
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46O001 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46O001 = mClsM46O001
End Property

Private Function mfConexaoAberta(ByVal pIdeBasDados As String, ByVal pSerBasDados As String, _
                                 ByVal pEndBasDados As String, ByVal pUIDBasDados As String, _
                                 ByVal pPWDBasDados As String, ByRef pADOBasDados As ADODB.Connection) As Boolean

'Par�metros de entrada...pIdeBasDados...Identificador da base de dados.
'                        pEndBasDados...Endere�o da base de dados.
'Par�metros de sa�da.....pADOBasDados...Base de dados.

    Dim lretorno As Boolean    'Variavel l�gica, indica se o procedimento utilizar� ou n�o a ultima conex�o aberta
' TRUE = Abre nova conex�o ; FALSE = Mantem a ultima conex�o aberta
    Dim lFlag As Boolean    'VAriavel l�gica auxiliar
'Utilizada para verificar da solicita��o de abertura de Base
'Caso a verifica��o encontre divergencias.
'TRUE = encontrado divergencia ; FALSE = N�o encontrado Divergencia Continua a verifica��o
    Dim lInd As Integer    'Contador.
    Dim lInd2 As Integer    'Contador auxiliar
    Dim lInfoConexao() As String  'informa��o sobre a conex�o.
    Dim lSubInfoCn() As String    'Recebe itens da variavel lInfoConexao

    pIdeBasDados = UCase(pIdeBasDados)
    pSerBasDados = UCase(pSerBasDados)
    pEndBasDados = UCase(pEndBasDados)
    pUIDBasDados = UCase(pUIDBasDados)
    pPWDBasDados = UCase(pPWDBasDados)

    '1. Verifica se a conex�o foi a ultima a ser aberta
    If pADOBasDados Is Nothing Then
        'Atribui true para a variavel lRetorno
        'Para abrir nova conex�o
        lretorno = False
    ElseIf pADOBasDados.State <> adStateOpen Then
        lretorno = False
    Else
        'separa as informa��es da ultima conex�o aberta em uma variavel matriz
        lInfoConexao = Split(UCase(pADOBasDados), ";")

        'Inicia a variavel para abrir nova conex�o
        lretorno = False

        'Inicia a variavel para prosseguir com a verifica��o
        lFlag = False

        For lInd = LBound(lInfoConexao()) To UBound(lInfoConexao())

            If UBound(lInfoConexao()) < 0 Then
                'Se a String estiver vazia
                'encerra a verifica��o, e o procedimento abrir� nova conex�o
                Exit For
            End If

            lSubInfoCn = Split(lInfoConexao(lInd), "=")

            For lInd2 = LBound(lSubInfoCn()) To 0

                If UBound(lSubInfoCn()) < 0 Then
                    Exit For
                End If

                Select Case lSubInfoCn(lInd2)
                Case "SERVER"
                    If lInfoConexao(lInd) = "SERVER=" & UCase(pSerBasDados) Then
                        'servidores de base de dados s�o iguais
                        'Atribui false para a Variavel lRetorno.
                        'Para manter a ultima conex�o aberta
                        lretorno = True
                    Else
                        'Encontrado Divergencia
                        'Atribui TRUE para a Variavel lFlag
                        'Para encerrar a verifica��o
                        lFlag = True
                        'Atribui TRUE para abrir nova conex�o
                        lretorno = False
                    End If
                Case "UID"
                    If lInfoConexao(lInd) = "UID=" & UCase(pUIDBasDados) Then
                        'Usuarios s�o iguais
                        lretorno = True
                    Else
                        lFlag = True
                        lretorno = False
                    End If
                Case "PWD"
                    If lInfoConexao(lInd) = "PWD=" & UCase(pPWDBasDados) Then
                        'Senhas de acesso ao Banco s�o iguias
                        lretorno = True
                    Else
                        lFlag = True
                        lretorno = False
                    End If
                Case "DATABASE"
                    If Replace(lInfoConexao(lInd), """", "") = "DATABASE=" & UCase(pEndBasDados) Then
                        'Bancos de Dados s�o iguais
                        lretorno = True
                    Else
                        lFlag = True
                        lretorno = False
                    End If
                End Select

            Next lInd2

            If lFlag = True Then
                'Encontrado divergencia
                'lRetorno � igual a False, o procedimento abrir� nova conex�o
                'Encerra a verifica��o
                Exit For
            End If

        Next lInd
    End If

Retorno:
    mfConexaoAberta = lretorno

End Function

Public Function mfEnviaProtocolo(ByVal pCodUser As String, _
                                 ByVal pFileName As String, _
                                 ByRef pMensagem As String) As Boolean

    Dim lIDUser As Integer
    Dim lCopyOrigem As String
    Dim lPasteDestino As String
    Dim lPathKill As String
    Dim lMensagem As String
    Dim lExec As String
    Dim lRsEAPDB As ADODB.Recordset
    Dim lCN_EAPDB As ADODB.Connection

    Err = 0
    On Error GoTo mfEnviaProtocolo_ERRO
    pMensagem = ""

    '1- verifica a existencia diretorio e o arquivo de origem
    'lCopyOrigem = gAppPath & "Remessa\" & pFileName
    lCopyOrigem = gAppPath & pFileName

    If gfPreenchido(Dir(gAppPath & "Remessa", vbDirectory)) = True Then
        If gfPreenchido(Dir(lCopyOrigem)) = False Then
            pMensagem = lCopyOrigem & Chr(13) & " arquivo n�o localizado"
            GoTo mfEnviaProtocolo_ERRO
        End If
    Else
        pMensagem = gAppPath & "APOLICES" & Chr(13) & " Caminho inexistente"
        GoTo mfEnviaProtocolo_ERRO
    End If

    '2- verifica a existencia diretorio destino

    '''    lPasteDestino = gEndPastas
    lPasteDestino = gAppPath

    lPasteDestino = lPasteDestino & pCodUser
    lPathKill = lPasteDestino & "\" & Mid(pFileName, 1, 9) & ".009.009"

    If gfPreenchido(Dir(lPasteDestino, vbDirectory)) = False Then
        'Verifica se o usuario j� possui pasta de arquivos
        'caso contr�rio cria a pasta
        MkDir lPasteDestino
    End If

    If gfPreenchido(Dir(lPasteDestino, vbDirectory)) = False Then
        pMensagem = lPasteDestino & Chr(13) & "N�o foi possivel acessar o destino."
        GoTo mfEnviaProtocolo_ERRO
    End If

    'Verifica tambem a existencia da pasta temp
    '********** Sem esta pasta a trasmiss�o de protocolo ao usuario
    '********** gera erro e encera o servi�o EDIServer
    If gfPreenchido(Dir(lPasteDestino & "\temp", vbDirectory)) = False Then
        'Verifica se o usuario j� possui pasta de arquivos
        'caso contr�rio cria a pasta
        MkDir lPasteDestino & "\temp"
    End If

    If gfPreenchido(Dir(lPasteDestino & "\temp", vbDirectory)) = False Then
        pMensagem = lPasteDestino & Chr(13) & "N�o foi possivel acessar o destino."
        GoTo mfEnviaProtocolo_ERRO
    End If

    '3- Apaga o arquivo destino caso j� exista no diretorio destino

    'Concatena o nome do arquivo da String
    lPasteDestino = lPasteDestino & "\" & Mid(pFileName, 1, Len(pFileName) - 4) & ".009.009"

    If gfPreenchido(Dir(lPathKill)) = True Then
        Kill lPathKill
    End If

    '4- Copia o arquivo de ap�lices para a pasta do usuario
    FileCopy lCopyOrigem, lPasteDestino

    '5- verifica a existencia do arquivo no destino
    If gfPreenchido(Dir(lPasteDestino)) = False Then
        'Arquivo n�o foi copiado
        pMensagem = lPasteDestino & Chr(13) & "N�o foi possivel copiar o arquivo."
        GoTo mfEnviaProtocolo_ERRO
    End If

    '6- Abre a base de dados

    If gfAbrBasSQL(gEndGed, gTipGed, gSerGed, gEndGed, gUidGed, _
                   gPwdGed, lCN_EAPDB, lMensagem) = False Then

        '    If gM46V999.gfAbrBasSQL("EAPDB", gTipBasEAPDB, gSerBasEAPDB, gEndBasEAPDB, gUIDBasEAPDB, _
             '                   gPWDBasEAPDB, lCN_EAPDB, lMensagem) = False Then
        Call gpGraLog(0, "gpMarcaRecebido - " & lMensagem)
        GoTo mfEnviaProtocolo_ERRO
    Else
        'Inclui registro para receber este arquivo
        lExec = "[EAPDB].[dbo].[Proc_EAPDB_IncluiEnvio] "
        lExec = lExec & "'YASUDA', "
        lExec = lExec & "'" & Mid(pFileName, 1, Len(pFileName) - 4) & ".009', "
        lExec = lExec & "'" & pCodUser & "'"

        If gfObtRegistroSIES(lCN_EAPDB, lExec, lRsEAPDB, lMensagem) = False Then
            If Err = 0 Then
                Err = 1
            End If
            GoTo mfEnviaProtocolo_ERRO
        Else
            lCN_EAPDB.Close
        End If
    End If

    '7- Exclui o arquivo de origem
    If gfPreenchido(Dir(lCopyOrigem)) = True Then
        Kill lCopyOrigem
    End If

    '8 Finalizado OK
    mfEnviaProtocolo = True

    'Saida da Fun��o
mfEnviaProtocolo_SAIDA:
    Exit Function

    'Tratamento de erro
mfEnviaProtocolo_ERRO:
    mfEnviaProtocolo = False
    pMensagem = "mfEnviaProtocolo - Erro " & Err & " " & ERROR$ & pMensagem

    GoTo mfEnviaProtocolo_SAIDA
End Function

Public Function gfObtRegistroSIES(ByVal pNomArquivo As ADODB.Connection, ByVal pComando As String, _
                                  ByRef pRegistro As ADODB.Recordset, ByRef pMensagem As String) As Boolean
'Fun��o: obt�m registro.

'Par�metros de entrada...pNomArquivo...Nome do arquivo.
'                        pComando .....Comando SELECT.
'Par�metros de sa�da.....pRegistro.....Registro (recordset).
'                        pMensagem.....Mensagem.
    Dim lInicio As Double
    gfObtRegistroSIES = False
Inicio:

    On Error Resume Next

    Set pRegistro = New ADODB.Recordset

    pRegistro.CursorLocation = adUseClient
    'Abre o Recordset

    '    If gGraLog = "2" Then
    '        Call gpGraLog(0, "AGE0000 - gfObtRegistroSIES : pComando=" & pComando)
    '    End If

    If InStr(1, UCase(pNomArquivo), "SYAS_EMIS") > 0 Then
        If InStr(1, UCase(pComando), "GED") > 0 And _
           InStr(1, UCase(pComando), "SYAS_EMIS") > 0 Then

            'N�o altera o comando

        Else
            'Somente se n�o estiver fazendo "JOIN" entre tabelas da base GED e SYAS_EMIS
            'atraves de conex�o com a base SYAS_EMIS
            pComando = Replace(pComando, "'", "!;$")
            If Len(pComando) > 5000 Then
                pComando = " exec [SYAS_EMIS].[SYAS].[Yproc_ExecDML2] '" & Mid(pComando, 1, 7500) & "','" & Mid(pComando, 7501) & "'"
            Else
                pComando = " exec [SYAS_EMIS].[SYAS].[Yproc_ExecDML2] '" & Mid(pComando, 1, 10) & "','" & Mid(pComando, 11) & "'"
            End If
        End If
    Else
        If InStr(1, UCase(pNomArquivo), "EAPDB") > 0 Then
            pComando = " exec EAPDB_ExecDML '" & Replace(pComando, "'", "!;$") & "'"
        End If
    End If
    lInicio = Timer
    If UCase(pNomArquivo) Like "*MIX*" Then
        Call gpGraLog(0, pComando)
    End If
    pRegistro.Open pComando, pNomArquivo, adOpenForwardOnly, adLockReadOnly
    'Set pRegistro = pNomArquivo.Execute(pComando)
    If Err <> 0 Then
        If UCase(Err.Description) Like "FALHA DE V�NCULO DE COMUNICA��O" Or _
           UCase(Err.Description) Like "COMMUNICATION LINK FAILURE" Then
            Call gpGraLog(0, Err.Description & " / tentativa de se reconectar")
            If UCase(App.EXEName) = "P28V200" Then
                If InStr(1, UCase(pNomArquivo), "EAPDB") > 0 Then
                    Set pNomArquivo = Nothing

                    If gfAbrBasSQL(gEndSYASEMIS, gTipSYASEMIS, gSerSYASEMIS, gEndSYASEMIS, gUidSYASEMIS, _
                                   gPwdSYASEMIS, pNomArquivo, pMensagem) = False Then
                        Call gpGraLog(2, pMensagem)
                        Call gpGraLog(0, pMensagem & " / erro na tentativa de se reconectar")
                        Exit Function
                    End If
                    GoTo Inicio
                Else
                    If UCase(pNomArquivo) Like "*MIX*" Then
                        If gfAbrBasSQL("DBSYASEMIS", "2", "SPX20112CSLUMIX", "SYAS_EMIS", "workflow", _
                                       "yasworkflow", pNomArquivo, pMensagem) = False Then
                            Call gpGraLog(2, pMensagem)
                            Exit Function
                        End If
                    Else
                        Set pNomArquivo = Nothing
                        If gfAbrBasSQL("DBSYASEMIS", "2", "LPX20102PSLUSQL", "SYAS_EMIS", "workflow", _
                                       "yasworkflow", pNomArquivo, pMensagem) = False Then
                            Call gpGraLog(2, pMensagem)
                            Exit Function
                        End If
                    End If
                    GoTo Inicio
                End If
            End If
            Call gpGraLog(2, "Comunica��o com o banco de dados foi perdida. Verifique se o acesso ao servidor " & Mid(pNomArquivo, InStr(1, pNomArquivo, "SERVER=") + 7, 15) & " e reinicie o procedimento.")
            Select Case UCase(App.EXEName)
            Case "P28V100", "P28V200", "P28V101", "P28V001"
                'End
                Exit Function
            End Select
        End If
        pMensagem = "FTE0044 - Erro ao obter registro (Erro = " & Err & " - " & ERROR & _
                  " - comando: " & pComando & ")."
        Exit Function
    End If
    If InStr(pNomArquivo, "UID=") = 0 And InStr(pNomArquivo, "Initial Catalog=") = 0 Then
        pRegistro.ActiveConnection = Nothing
    End If
    Select Case Err
    Case 0      'Sem erro.
        gfObtRegistroSIES = True
    Case 3159   'Marcador inv�lido - arquivo dados corrompido.
        pMensagem = "FTE0044 - Erro ao obter registro (arquivo corrompido - comando: " & pComando & _
                    ")."
    Case Else
        pMensagem = "FTE0044 - Erro ao obter registro (Erro = " & Err & " - " & ERROR & _
                  " - comando: " & pComando & ")."
    End Select
    On Error GoTo 0
End Function
Public Property Get M46V103() As clsM46V103
    Set M46V103 = M46V111.M46V103
End Property

Public Property Set M46V103(pClsM46V103 As clsM46V103)
    Set mClsM46V103 = pClsM46V103
End Property

Public Property Get gTarifa_Apolice() As String
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstTarifa As clsYasRecordSet
    Dim lSelect As String
    Dim lMensagem As String

    'Rotina para obter a tarifa que a apolice foi emitida
    gTarifa_Apolice = "0"

    If gNum_Apolice = 0 Then
        GoTo SAIDA
    End If

    '1. Abre arquivo.
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    '3. Obt�m a Tarifa da apolice
    lSelect = "select cod_produto from Tab_Ped where NUM_APOL = " & gNum_Apolice & " And Num_Endo = 0 "

    If gfObtRegistro(BDP00Tarifa, lSelect, rstTarifa, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    Call gpFecha2(BDP00Tarifa)

    If Not rstTarifa.EOF Then
        gTarifa_Apolice = rstTarifa("cod_produto")
    End If

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set BDP00Tarifa = Nothing
    Set rstTarifa = Nothing
    Return

End Property

Public Sub updAtualizaStatusInspecao(NUM_PI As String, ByRef pMensInsp As String, ByRef lObjstcPed As stcA46V702B)
    Dim lretorno As String
    Dim lI As Integer
    Dim lCod As Integer
    Dim lAuxiliar As String
    Dim lBDP00Multi As ADODB.Connection
    Dim lRstErro As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String
    Dim lrsPadrao As clsYasRecordSet
    Dim NossoNumero As String
    Dim lAchou As Boolean
   
    
    'piloto
    lAchou = False

    pMensInsp = ""

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
        Call gpGraLog(2, lMensagem)
        Exit Sub
    End If

    ' Tentar encontrar a Inspe��o a partir do PI e Triagem
    lSelect = "SELECT TTI.Num_Cgc_Cpf " & vbCrLf
    lSelect = lSelect & "FROM        Tab_Trans_Insp   TTI " & vbCrLf
    lSelect = lSelect & "INNER JOIN  ged..Tab_Triagem                TT " & vbCrLf
    lSelect = lSelect & "    ON      TTI.NOSSO_NUMERO                = TT.num_siscota " & vbCrLf
    lSelect = lSelect & "        AND TT.num_pi                       = '" & NUM_PI & "' AND TTI.Cod_Status_I = '0' " & vbCrLf

    If gfObtRegistro(lBDP00Multi, lSelect, lrsPadrao, lMensagem) Then
        If Not lrsPadrao.EOF Then
            lAchou = True
        
            lSelect = "UPDATE      TTI " & vbCrLf
            lSelect = lSelect & "SET         Cod_Status_I                    = " & IIf(lrsPadrao!Num_Cgc_Cpf = "0" And lObjstcPed.Num_Cgc_Cpf = "0", "'0'", "'1'") & vbCrLf    'so atualiza o status se retornar CNPJ ou CPF.
            lSelect = lSelect & "            ,NUM_CGC_CPF                    = '" & IIf(Trim(lObjstcPed.Num_Cgc_Cpf) = "0", lrsPadrao!Num_Cgc_Cpf, lObjstcPed.Num_Cgc_Cpf) & "'" & vbCrLf
            lSelect = lSelect & "FROM        Tab_Trans_Insp   TTI " & vbCrLf
            lSelect = lSelect & "INNER JOIN  ged..Tab_Triagem                TT " & vbCrLf
            lSelect = lSelect & "    ON      TTI.NOSSO_NUMERO                = TT.num_siscota " & vbCrLf
            lSelect = lSelect & "        AND TT.num_pi                       = '" & NUM_PI & "' AND TTI.Cod_Status_I = '0' " & vbCrLf

            If gfExeSQL(lBDP00Multi, lSelect, lMensagem) <> 0 Then
                Call gpGraLog(2, lMensagem)
                Exit Sub
            End If
        Else
            ' N�o encontrou, tentar ent�o obter o nosso n�mero a partir da TAB_PED
            lSelect = "SELECT      TPNC.NOSSO_NUMERO_OLD,TP.Num_Cgc_Cpf " & vbCrLf
            lSelect = lSelect & "FROM        Tab_Ped                 TP WITH (NOLOCK) " & vbCrLf
            lSelect = lSelect & "INNER JOIN  Tab_Ped_Num_Copia       TPNC WITH (NOLOCK) " & vbCrLf
            lSelect = lSelect & "    ON      TPNC.NOSSO_NUMERO_NEW   = TP.NOSSO_NUMERO " & vbCrLf
            lSelect = lSelect & "INNER JOIN  Tab_Trans_Insp          TTI WITH (NOLOCK) " & vbCrLf
            lSelect = lSelect & "    ON      TTI.NOSSO_NUMERO        = TPNC.NOSSO_NUMERO_OLD " & vbCrLf
            lSelect = lSelect & "WHERE       TP.NUM_PI       = '" & NUM_PI & "' "

            If gfObtRegistro(lBDP00Multi, lSelect, lrsPadrao, lMensagem) Then
                If Not lrsPadrao.EOF Then
                    lAchou = True
                    NossoNumero = lrsPadrao("NOSSO_NUMERO_OLD")
                    lSelect = "UPDATE Tab_Trans_Insp " & vbCrLf
                    lSelect = lSelect & "SET Cod_Status_I = " & IIf(lrsPadrao!Num_Cgc_Cpf = "0" And lObjstcPed.Num_Cgc_Cpf = "0", "'0'", "'1'") & vbCrLf
                    lSelect = lSelect & ",NUM_CGC_CPF = '" & IIf(Trim(lObjstcPed.Num_Cgc_Cpf) = "0", lrsPadrao!Num_Cgc_Cpf, lObjstcPed.Num_Cgc_Cpf) & "'" & vbCrLf
                    lSelect = lSelect & "WHERE NOSSO_NUMERO = '" & NossoNumero & "' AND Cod_Status_I = '0'  " & vbCrLf

                    If gfExeSQL(lBDP00Multi, lSelect, lMensagem) <> 0 Then
                        Call gpGraLog(2, lMensagem)
                        Exit Sub
                    End If

                End If
            End If
        End If
    End If

    'piloto
    ' N�O ESQUE�A PARA O KIT - V� em Project > Properties > Aba "Make" e preencha
    ' o campo "Conditional Compilation Arguments" com "APPTarget = 1"
    #If APPTarget = 1 Then
        If Not lAchou And lObjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
            Call mfAgendarInspecaoProduto(lObjstcPed, lMensagem, 1)
        End If
    #End If

    Call gpFecha2(lBDP00Multi)

End Sub

Public Sub forcaStatusInspecao(NUM_PI As String)
    Dim lretorno As String
    Dim lI As Integer
    Dim lCod As Integer
    Dim lAuxiliar As String
    Dim lBDP00Multi As ADODB.Connection
    Dim lRstErro As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String
    Dim lrsPadrao As clsYasRecordSet
    Dim NossoNumero As String
    Dim lAchou As Boolean
    
    'piloto
    lAchou = False

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
        Call gpGraLog(2, lMensagem)
        Exit Sub
    End If
    
    lSelect = "UPDATE          TTI" & vbCrLf
    lSelect = lSelect & "Set Cod_Status_I = 1" & vbCrLf
    lSelect = lSelect & "FROM            RamosDiversos..Tab_Trans_Insp       TTI" & vbCrLf
    lSelect = lSelect & "INNER JOIN      RamosDiversos..Tab_Ped              TP" & vbCrLf
    lSelect = lSelect & "    ON          TP.NOSSO_NUMERO                     = TTI.NOSSO_NUMERO" & vbCrLf
    lSelect = lSelect & "            AND TP.NUM_PI                           = '" & NUM_PI & "'" & vbCrLf
    lSelect = lSelect & "WHERE           TTI.Cod_Status_I                    = 0" & vbCrLf

    If gfExeSQL(lBDP00Multi, lSelect, lMensagem) <> 0 Then
        Call gpGraLog(2, lMensagem)
        Exit Sub
    End If
    
    lSelect = "Update TTI" & vbCrLf
    lSelect = lSelect & "Set Cod_Status_I = 1" & vbCrLf
    lSelect = lSelect & "FROM            RamosDiversos..Tab_Trans_Insp       TTI" & vbCrLf
    lSelect = lSelect & "INNER JOIN      RamosDiversos..Tab_Ped_Num_Copia    TPNC" & vbCrLf
    lSelect = lSelect & "    ON          TPNC.NOSSO_NUMERO_OLD               = TTI.NOSSO_NUMERO" & vbCrLf
    lSelect = lSelect & "INNER JOIN      RamosDiversos..Tab_Ped              TP" & vbCrLf
    lSelect = lSelect & "    ON          TP.NOSSO_NUMERO                     = TPNC.NOSSO_NUMERO_NEW" & vbCrLf
    lSelect = lSelect & "            AND TP.NUM_PI                           = '" & NUM_PI & "'" & vbCrLf
    lSelect = lSelect & "Where TTI.Cod_Status_I = 0" & vbCrLf
    
    If gfExeSQL(lBDP00Multi, lSelect, lMensagem) <> 0 Then
        Call gpGraLog(2, lMensagem)
        Exit Sub
    End If
    
    Call gpFecha2(lBDP00Multi)

End Sub

Public Sub mpNovoBuscaLimFranquiaAvancada(ByVal pCodRamo As Integer, ByVal pCodModalidade As Integer, ByVal pCodTarifa As Integer, _
                                            ByVal pCodCobert As Integer, ByVal pCod_Clasf As String, ByVal pCod_Compl_Clasf As String, _
                                            ByVal pCod_Constr As Integer, ByVal pCEP As Long, ByVal pCodSequencia As Integer, ByVal pCodCorretor As Long, _
                                            ByRef ValFraqMin As Double, ByRef Perc_Franquia As Double, ByRef Fator_Ajuste As Double, _
                                            ByRef ValFranqMax As Double, ByRef TextoFranquia As String, ByRef Dsc_Cober_Franq As String, _
                                            ByVal pNumPiMeses As Integer, ByRef pMuliploFraq As Integer, ByRef pDescMultFraq As Double, _
                                            ByVal pValIs As Double)

    Dim lbdTbm_Franquia_Avanc As ADODB.Connection         'Arquivo: P00Multi.
    Dim lrsTbm_Franquia_Avanc As clsYasRecordSet          'Registro : Tbm_Franquia_Avancada.
    Dim lrsFranquia As clsYasRecordSet
    Dim rstTabClas As clsYasRecordSet
    Dim lMensagem As String         'Mensagem.
    Dim lCodSequencia As Integer
    Dim lSelect As String
    Dim lTip_Atividade As Integer
    Dim Tem_VideObs As String

    On Error GoTo ControleErro

    ValFraqMin = 0
    Perc_Franquia = 0
    Fator_Ajuste = 0
    ValFranqMax = 0
    Dsc_Cober_Franq = ""
    TextoFranquia = ""

    lCodSequencia = pCodSequencia

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTbm_Franquia_Avanc, lMensagem) = False Then
        Call gpGraLog(2, lMensagem)
        Exit Sub
    End If

    lSelect = " Select Val_Min_Franq, Fator_Franq, Per_Franq, Periodo_Dias, Texto_Franquia, Texto_Observacao  "
    lSelect = lSelect & " from Tbm_FranquiaS_Avancada "
    lSelect = lSelect & " Where Cod_Ramo = " & pCodRamo
    lSelect = lSelect & " and Cod_Modalidade = " & pCodModalidade
    'comentar por enquanto
    'lSelect = lSelect & " and Tip_Cobert = " & pTip_Cobert
    lSelect = lSelect & " and ((Cod_Clasf = '" & pCod_Clasf & "'"
    lSelect = lSelect & " and Cod_Compl_Clasf = '" & pCod_Compl_Clasf & "')"
    lSelect = lSelect & " or (Cod_Clasf = '0' OR Cod_Clasf IS NULL))"
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & " and Cod_Tarif = " & pCodTarifa
    lSelect = lSelect & " AND ((CONVERT(NUMERIC(8),CEP_Ini) <= " & pCEP
    lSelect = lSelect & " AND CONVERT(NUMERIC(8),CEP_ter) >= " & pCEP & " ) "
    lSelect = lSelect & " OR (CONVERT(NUMERIC(8),CEP_Ini) =  0 OR CEP_Ini IS NULL))"
    lSelect = lSelect & " AND (COD_CORR = pCodCorretor OR COD_CORR IS NULL)"
    lSelect = lSelect & " order by Val_Min_Franq desc, Per_Franq "


    If gfObtRegistro(lbdTbm_Franquia_Avanc, lSelect, lrsTbm_Franquia_Avanc, lMensagem) = False Then
        Call gpFecha2(lbdTbm_Franquia_Avanc)
        Call gpGraLog(2, " mpNovoBuscaLimFranquia " & lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If lrsTbm_Franquia_Avanc.RecordCount > 0 Then
        ValFraqMin = lrsTbm_Franquia_Avanc("Val_Min_Franq")
        Perc_Franquia = lrsTbm_Franquia_Avanc("Per_Franq")
        Fator_Ajuste = IIf(IsNull(lrsTbm_Franquia_Avanc("Fator_Franq")), 1, lrsTbm_Franquia_Avanc("Fator_Franq"))
        Tem_VideObs = Check(lrsTbm_Franquia_Avanc("Texto_Observacao"))
        TextoFranquia = Check(lrsTbm_Franquia_Avanc("Texto_Franquia"))
        pMuliploFraq = 1
        pDescMultFraq = 0
    End If
    
FimBusca:

    Call gpFecha2(lbdTbm_Franquia_Avanc)

    '1001135 - Corretor Volks - Franquia Especial Concessionaria
    If pCodTarifa >= 1002 Then
        If (InStr(1, "00008001/00008002/00008003/00008004/00008005/00008006/00008008/00008011/00008014/00008028", Format(pCodCorretor, "00000000"), vbTextCompare) > 0) And _
           (val(pCod_Clasf) = 22 And val(pCod_Compl_Clasf) = 13) And _
           (InStr(1, "0038/0064/0066", Format(pCodCobert, "0000"), vbTextCompare) > 0) Then
            Perc_Franquia = 15
            ValFraqMin = 2000
        End If
    End If
    Exit Sub
ControleErro:
    Call gpGraLog(2, " mpNovoBuscaLimFranquia " & Err & " - " & ERROR)
End Sub

Public Function gfPlanoBasico(ByVal pRamo As Integer, Optional pPlano As Integer) As Integer
'Retorna os c�digos dos planos basico do ramo especificado no parametro
    If pPlano <> 10 And pPlano <> 11 And pPlano < 2000 Then
        Select Case pRamo
            Case e_CodRamo_Empresarial
                If pPlano = 2 Then
                    gfPlanoBasico = e_CodPlano_Empresarial_Super
                Else
                    gfPlanoBasico = e_CodPlano_Empresarial_Simplificado
                End If
    
            Case e_CodRamo_Condominio
                If pPlano = 2 Then
                    gfPlanoBasico = e_CodPlano_Condominio_Amplo
                    Else
                    gfPlanoBasico = e_CodPlano_Condominio_Tradicional
                    End If
                
            Case e_CodRamo_Residencial
                If pPlano = 30 Then
                    gfPlanoBasico = pPlano
                Else
                gfPlanoBasico = e_CodPlanoBasico_Residencial
                End If
                
            Case Else
                gfPlanoBasico = 1
        End Select
    Else
        gfPlanoBasico = pPlano
    End If
End Function

Sub mfCarregarUnidAtivas(ByRef Rst_Unid_Ativa As clsYasRecordSet)
    Dim lSel_Unid_Ativa As String
    Dim lP00Unid_Ativa As ADODB.Connection
    Dim lMensagem As String
            
    On Local Error GoTo ERROR
    
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lP00Unid_Ativa, lMensagem) = False Then
        GoTo ERROR
    End If
    
    lSel_Unid_Ativa = "EXEC dbo.proc_ConsultarUnidAtiva "

    If gfObtRegistro(lP00Unid_Ativa, lSel_Unid_Ativa, Rst_Unid_Ativa, lMensagem) = False Then
        Call gpFecha2(lP00Unid_Ativa)
        GoTo ERROR
    End If

SAIDA:
    Call gpFecha2(lP00Unid_Ativa)
    Set lP00Unid_Ativa = Nothing
    Exit Sub
ERROR:
    Call gpGraLog(2, lMensagem)
    GoTo SAIDA

End Sub

Public Sub mfFranqNovoLim(pCodCobert As Integer, pValIs As Double, pCodClasf As String, pCodComplClasf As String, pCodConstr As Integer, _
                            pNumCep As String, pTipCobert As Integer, pNumMesesPi As Integer, ByRef lrsFranqNovoLim As clsYasRecordSet)
    Dim lbdFranqNovoLim As ADODB.Connection
    Dim lMensagem As String
    Dim lSelect As String
    Dim lretorno As Integer
    Dim lCod_Tarifa As Integer

    lretorno = False
                
    lCod_Tarifa = gCodProduto
    If lCod_Tarifa >= 1124 And gCod_Ramo = e_CodRamo_Empresarial And mClsM46V111.gCod_Dizer And M46V111.gCod_plano < 1000 And gTarifa_Apolice > 0 And gTarifa_Apolice < 1124 Then
        lCod_Tarifa = gTarifa_Apolice
    End If
                
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdFranqNovoLim, lMensagem) = False Then
        Call gpGraLog(2, lMensagem)
        Exit Sub
    End If
    
    lSelect = "EXEC proc_BuscaLimFranquiaNovo2 "
    lSelect = lSelect & gCod_Ramo & ","
    lSelect = lSelect & M46V111.gCod_plano & ","
    lSelect = lSelect & lCod_Tarifa & ","
    lSelect = lSelect & pCodCobert & ","
    lSelect = lSelect & gfFormatarCamposSQL(pValIs) & ","
    lSelect = lSelect & gfFormatarCamposSQL(pCodClasf) & ","
    lSelect = lSelect & gfFormatarCamposSQL(pCodComplClasf) & ","
    lSelect = lSelect & pCodConstr & ","
    lSelect = lSelect & gfFormatarCamposSQL(pNumCep) & ","
    lSelect = lSelect & 0 & ","
    lSelect = lSelect & gCod_Corretor & ","
    lSelect = lSelect & pTipCobert & ","
    lSelect = lSelect & "0, "
    lSelect = lSelect & pNumMesesPi
    
    If Not gfObtRegistro(lbdFranqNovoLim, lSelect, lrsFranqNovoLim, lMensagem) Then
        Call gpGraLog(1, lMensagem)
        Exit Sub
    End If
    
SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Call gpFecha2(lbdFranqNovoLim)
    Return
End Sub

Public Function mfMontaTexofranquia(pPerFranq As Double, pValMinFranq As Double, pDiasFranqMin As Integer, pTextoPadrao As String, pTextoCmpl As String, Tem_VideObs As String) As String
    Dim lretorno As String
 
    'Refaz texto
    If pPerFranq > 0 And pValMinFranq > 0 And pDiasFranqMin = 0 Then
        pTextoPadrao = UCase("% dos prej. Indeniz, limitado ao m�nimo de R$ ")
    Else
        If pPerFranq > 0 Then
            pTextoPadrao = CStr(pPerFranq) & UCase("% dos prej. Indeniz�veis")
        Else
            If pValMinFranq > 0 Then
                pTextoPadrao = "R$"
            End If
        End If
    End If
    
    If pValMinFranq > 0 Then
        If pPerFranq > 0 Then
            lretorno = pPerFranq & pTextoPadrao & " " & Format(pValMinFranq, "Standard")
        Else
            lretorno = pTextoPadrao & " " & Format(pValMinFranq, "Standard")
        End If
    End If
    
    If pDiasFranqMin > 0 Then
        If Trim(pTextoPadrao) <> Empty Then
        lretorno = pDiasFranqMin & " " & pTextoPadrao
        Else
            lretorno = pDiasFranqMin & " dia(s)"
        End If
    End If
    
    If pValMinFranq > 0 Or pDiasFranqMin > 0 Then
        lretorno = lretorno & " " & pTextoCmpl
    End If

        If Tem_VideObs = "*" Then
                lretorno = "*" & lretorno
        End If
    
SAIDA:
    If lretorno = Empty Then
        lretorno = pTextoPadrao
    End If
    mfMontaTexofranquia = UCase(lretorno)
End Function

Public Function gfObterTextosCobertura(ByVal pCodigoRamo As Integer, _
                                       ByVal pCodigoModalidade As Integer, _
                                       ByVal pCodigoCobertura As String, _
                                       ByVal pCodigoClassificacao As String, _
                                       ByVal pCodigoComplClassificacao As String, _
                                       ByVal pTipoCobertura As Integer, _
                                       ByRef pSaida() As Variant, _
                                       ByRef pMensagem As String) As Boolean
'Fun��o: obt�m dados de Textos das Coberturas.
    
'Par�metros de entrada...pCodigoRamo...C�digo do Ramo.
'Par�metros de entrada...pCodigoModalidade...C�digo do Plano.
'Par�metros de entrada...pTipoTexto...Tipo do Texto.
'Par�metros de entrada...pCodigoCobertura...C�digo da Cobertura.
'Par�metros de entrada...pCodigoClassificacao...C�digo da Classificacao.
'Par�metros de entrada...pCodigoComplClassificacao...C�digo Complementar da Classificacao.
'Par�metros de entrada...pCodigoTarifa...C�digo da Tarifa.
'Par�metros de entrada...pTipoCobertura...Tipo da Cobertura.


    Dim bdDiversos As ADODB.Connection
    Dim lI As Double                    'Contador.
    Dim lSelect As String               'Comando SELECT.
    Dim rsTextos As clsYasRecordSet     'Tabela: Tbm_Textos.

    On Error GoTo TrataErro

    gfObterTextosCobertura = False

    '1. Verifica par�metros.
    'Cod_Ramo.
    If Not gfPreenchido(pCodigoRamo) Then
        pMensagem = "FTE0068 - Erro ao obter ramo (c�digo do ramo n�o preenchido)."
        Exit Function
    End If
    
    'pCodigoModalidade
    If Not gfPreenchido(pCodigoModalidade) Then
        pMensagem = "FTE0068 - Erro ao obter plano (plano n�o preenchido)."
        Exit Function
    End If

    'Tipo_Texto
'    If Not gfPreenchido(pTipoTexto) Then
'        pMensagem = "FTE0068 - Erro ao obter plano (plano n�o preenchido)."
    '    Exit Function
    'End If

    '2. Busca informa��o.
    lSelect = "EXEC sp_Selecionar_Textos " & pCodigoRamo & ", " & pCodigoModalidade & ", '" & _
                                             pCodigoCobertura & "', '" & _
                                             pCodigoClassificacao & "', '" & _
                                             pCodigoComplClassificacao & "', " & _
                                             gCodProduto & ", " & gCod_Corretor & ", " & _
                                             pTipoCobertura                 'Luiz Alberto - 20/10/2015

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdDiversos, pMensagem) = False Then
        Exit Function
    End If

    If gfObtRegistro(bdDiversos, lSelect, rsTextos, pMensagem) = False Then
        Exit Function
    End If

    If rsTextos.RecordCount = 0 Then
        Call Me.gpFecha2(bdDiversos)
        ReDim pSaida(1, 5)
        gfObterTextosCobertura = True
        Exit Function
    End If

    rsTextos.MoveLast

    ReDim pSaida(rsTextos.RecordCount, 5)
    lI = 0

    rsTextos.MoveFirst

    While Not rsTextos.EOF
        lI = lI + 1
        pSaida(lI, 1) = rsTextos!Tipo_Texto
        pSaida(lI, 2) = rsTextos!Cod_Cobertura
        pSaida(lI, 3) = rsTextos!Cod_Texto
        pSaida(lI, 4) = rsTextos!Desc_Texto
        pSaida(lI, 5) = CInt("" & rsTextos!NUM_SEQ)

        rsTextos.MoveNext
    Wend

    Call Me.gpFecha2(bdDiversos)
    
    gfObterTextosCobertura = True
    
SAIDA:
    Set bdDiversos = Nothing
    Set rsTextos = Nothing
    
    Exit Function
    
TrataErro:
    gfObterTextosCobertura = False
    If gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " gfObterTextosCobertura - " & ERROR
    End If
    GoTo SAIDA
    
End Function

Public Property Set DynaTrace(obj As clsMDynaTrace)
    Set mClsMDynaTrace = obj
End Property

Public Property Get DynaTrace() As clsMDynaTrace
    If mClsMDynaTrace Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto clsMDynaTrace n�o atribu�do para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsMDynaTrace j�
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsMDynaTrace ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set DynaTrace = mClsMDynaTrace
End Property

Public Function mfSalvarNovosTextoCoberturas(ByRef pStcPed As stcA46V702B, pMensagem As String, pPropTabela As String) As Boolean
    
    '8000042754 - Resposta dada a quest�o parametrizada para o plano.
    Dim pSaidaPergunta() As Variant
    Dim pSaidaResposta() As Variant
    
    Dim lTotLinResposta As Long
    Dim intCodResposta As Integer
    
    Dim CodigoTextoLivreIncluido As String 'JIRA-25945 - Op��o "Sem Deprecia��o" no Condom�nio
    
    Dim lobjTexto As clsA46V739A
    Dim lobjstcTexto As stcA46V739B
    Dim lobjstcCobert As stcA46V704B
    Dim lobjstcItem As stcA46V708B
    
    Dim ColecaoTextosTemp As Collection
    Dim ColecaoItens      As Collection
    
    Dim lTotLin As Long
    Dim i As Integer
    Dim lMensagem As String

    On Error GoTo ControleErro

    Set lobjTexto = New clsA46V739A
    Set lobjTexto.M46V999 = Me
    
    Set lobjstcTexto = New stcA46V739B
    Set ColecaoTextosTemp = New Collection
    
    lobjstcTexto.nosso_numero = pStcPed.nosso_numero
    If Not lobjTexto.mfExcluir(lobjstcTexto, pPropTabela, "") Then
        pMensagem = lobjTexto.UltimoErro
        GoTo SAIDA
    End If
    
    'Monta a lista de coberturas e faz a chamada p/ a classe M46V999
    Dim strCoberturas As String
    Dim strClassificacoes As String
    Dim pSaida() As Variant

    Set ColecaoItens = New Collection

    For Each lobjstcItem In pStcPed.ColecaoPedLoc
        strCoberturas = ""
        For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
            strCoberturas = strCoberturas & lobjstcCobert.Cod_Cobert & ","
        Next lobjstcCobert
        
        'JIRA-25945 - Op��o "Sem Deprecia��o" no Condom�nio
        CodigoTextoLivreIncluido = "" 'vari�vel para identificar se esse texto livre j� foi inclu�do no item Local risco
        
        'Luiz Alberto (lobjstcItem.Tip_Cobert)  - 20/10/2015
        If Not gfObterTextosCobertura(pStcPed.Cod_Ramo, M46V111.gCod_plano, strCoberturas, lobjstcItem.Cod_Clasf, lobjstcItem.Cod_Compl_Clasf, lobjstcItem.Tip_Cobert, pSaida(), pMensagem) Then
            GoTo SAIDA
        End If
        
        lTotLin = UBound(pSaida)
        If pSaida(1, 1) = Empty Then
            '8000043007, Sistema apresentando clausula indevida
            Set lobjstcItem.ColecaoTextos = lobjTexto.Colecao
            GoTo Proximo
        End If
        
        For i = 1 To lTotLin
            Set lobjstcTexto = New stcA46V739B
            lobjstcTexto.nosso_numero = pStcPed.nosso_numero
            lobjstcTexto.Cod_Org = 0
            lobjstcTexto.Cod_Ramo = gCod_Ramo
            lobjstcTexto.Num_Ped = pStcPed.Num_Ped
            lobjstcTexto.Num_Apol = pStcPed.Num_Apol
            lobjstcTexto.Num_Endo = pStcPed.Num_Endo
            lobjstcTexto.NUM_PI = pStcPed.NUM_PI
        
            lobjstcTexto.Num_Item = lobjstcItem.Num_Item
            lobjstcTexto.Tipo_Texto = pSaida(i, 1)
            
            
            lobjstcTexto.Cod_Cobert = pSaida(i, 2)
            lobjstcTexto.Cod_Texto = pSaida(i, 3)
            lobjstcTexto.Desc_Texto = pSaida(i, 4)
            lobjstcTexto.NUM_SEQ = pSaida(i, 5)
            'lobjstcTexto.Tip_Cobert = pSaida(i, 6)
            
            'Refatorar
            'Inicio - 8000042754 - Resposta dada a quest�o parametrizada para o plano.
            
            'Verificar se existem parametrizacoes por question�rio.
            If Not gfObterPerguntaResposta(pStcPed.Cod_Ramo, M46V111.gCod_plano, gCodProduto, CInt(pSaida(i, 3)), pSaidaPergunta(), pMensagem, pPropTabela) Then
                GoTo SAIDA
            End If

            lTotLinResposta = UBound(pSaidaPergunta)

            If Not pSaidaPergunta(0, 1) = Empty Then

                'Obter resposta da pergunta encontrada.
                If Not gfObterRespostaPergunta(pStcPed.nosso_numero, lobjstcItem.Num_Item, pSaidaResposta(), pMensagem, pPropTabela) Then
                    GoTo SAIDA
                End If
                
                If Not IsNull(pSaidaResposta(0, 16)) Then
                
                    intCodResposta = CInt(pSaidaResposta(0, 16))
                        
                    Dim v As Integer
                    Dim d As Boolean
                    d = False
                    
                    For v = 0 To lTotLinResposta
                        If intCodResposta = CInt(pSaidaPergunta(v, 5)) Then
                            d = True
                        End If
                    Next
                    
                End If
                
                If d Then
                
                    If (lobjstcTexto.Cod_Texto <> 0) Then
                        If Not lobjTexto.mfIncluir(lobjstcTexto, pPropTabela, "", tagDynaTrace) Then
                            pMensagem = lobjTexto.UltimoErro
                            GoTo SAIDA
                        End If
                         
                    End If
                                       
                End If

            End If
            
            If pSaidaPergunta(0, 1) = Empty Then
            
                If (lobjstcTexto.Cod_Texto <> 0) Then
                    If Not lobjTexto.mfIncluir(lobjstcTexto, pPropTabela, "", tagDynaTrace) Then
                        pMensagem = lobjTexto.UltimoErro
                        GoTo SAIDA
                    End If
                     
                End If
                                   
            End If

            'Fim - 8000042754 - Resposta dada a quest�o parametrizada para o plano.

            Set lobjstcItem.ColecaoTextos = lobjTexto.Colecao
        Next
Proximo:
        ColecaoItens.Add lobjstcItem, lobjstcItem.Class_ID
    Next lobjstcItem
    Set pStcPed.ColecaoPedLoc = ColecaoItens
    mfSalvarNovosTextoCoberturas = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    Set lobjTexto = Nothing
    Set lobjstcTexto = Nothing
    
    If gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    
    Call gpGraLog(0, "mfSalvarNovosTextoCoberturas  ERRO - " & Err & " " & Err.Description & Chr(13) & pMensagem)

    GoTo SAIDA
FechaObj:
    Set lobjTexto = Nothing
    Set lobjstcTexto = Nothing
    Set lobjstcCobert = Nothing
    Set lobjstcItem = Nothing
    Return
End Function

'JIRA-25945 - Op��o "Sem Deprecia��o" no Condom�nio
Public Function mfFormatarConsultaTextoLivreSQL(ByRef pobjstc As stcA46V739B, ByVal pPropTabela As String) As clsYasRecordSet
                           
    On Error GoTo TRATA_ERRO
    
    Dim db As ADODB.Connection
    Dim rsTexto As clsYasRecordSet
    
    Dim encontrou As Boolean
        
    Dim mensagemErro As String

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, db, mensagemErro) = False Then
        Exit Function
    End If

    Set rsTexto = New clsYasRecordSet
    
    Dim lSelect As String 'Query.

    lSelect = "SELECT Tab_Ped_Textos.NOSSO_NUMERO, "
    lSelect = lSelect & "Tab_Ped_Textos.Cod_Org, "
    lSelect = lSelect & "Tab_Ped_Textos.Cod_Ramo, "
    lSelect = lSelect & "Tab_Ped_Textos.Num_Ped, "
    lSelect = lSelect & "Tab_Ped_Textos.Num_Apol, "
    lSelect = lSelect & "Tab_Ped_Textos.Num_Endo, "
    lSelect = lSelect & "Tab_Ped_Textos.Num_PI, "
    lSelect = lSelect & "Tab_Ped_Textos.Num_Item, "
    lSelect = lSelect & "Tab_Ped_Textos.Cod_Cobert, "
    lSelect = lSelect & "Tab_Ped_Textos.Tipo_Texto, "
    lSelect = lSelect & "Tab_Ped_Textos.Cod_Texto, "
    lSelect = lSelect & "Tab_Ped_Textos.Desc_Texto, "
    lSelect = lSelect & "Tab_Ped_Textos.Num_Seq"
    
    lSelect = lSelect & " From " & pPropTabela & "Tab_Ped_Textos WITH (NOLOCK) "
    
    If M99V001.M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lSelect = M99V001.M46V999.gfWhereAnd(lSelect) & " Tab_Ped_Textos.Nosso_Numero= " & M99V001.M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If

    If pobjstc.Num_Item > 0 Then
        lSelect = M99V001.M46V999.gfWhereAnd(lSelect) & " Tab_Ped_Textos.Num_Item = " & M99V001.M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If
    
    If pobjstc.Cod_Cobert > 0 Then
        lSelect = M99V001.M46V999.gfWhereAnd(lSelect) & " Tab_Ped_Textos.Cod_Cobert = " & M99V001.M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert)
    End If
     
    If pobjstc.NUM_SEQ > 0 Then
        lSelect = M99V001.M46V999.gfWhereAnd(lSelect) & " Tab_Ped_Textos.NUM_SEQ = " & M99V001.M46V999.gfFormatarCamposSQL(pobjstc.NUM_SEQ)
    End If
    
    If pobjstc.Cod_Texto > 0 Then
        lSelect = M99V001.M46V999.gfWhereAnd(lSelect) & " Tab_Ped_Textos.Cod_Texto = " & M99V001.M46V999.gfFormatarCamposSQL(pobjstc.Cod_Texto)
    End If
       
    lSelect = lSelect & " order by Tab_Ped_Textos.nosso_numero , Tab_Ped_Textos.Num_Item, Tab_Ped_Textos.Cod_Cobert, Tab_Ped_Textos.NUM_SEQ"
          
    If gfObtRegistro(db, lSelect, rsTexto, mensagemErro) = False Then
        Exit Function
    End If
    
    Set mfFormatarConsultaTextoLivreSQL = rsTexto
    
FIM_FUNCTION:
    On Error Resume Next
    Set rsTexto = Nothing
    db.Close
    Set db = Nothing
    On Error GoTo 0
    Exit Function
TRATA_ERRO:
        
    mensagemErro = " Function BuscaTextoLivre (Error) - (" & Err.Number & ")-" & Err.Description
    
    GoTo FIM_FUNCTION

End Function

'8000042754 - Resposta dada a quest�o parametrizada para o plano.
Public Function gfObterPerguntaResposta(ByVal pCodigoRamo As Integer, _
                                       ByVal pCodigoModalidade As Integer, _
                                       ByVal pCodigoTarifa As String, _
                                       ByVal pIdTexto As Integer, _
                                       ByRef pSaida() As Variant, _
                                       ByRef pMensagem As String, _
                                       ByVal pPropTabela As String) As Boolean
                                       
                                      
    Dim bdDiversos As ADODB.Connection
    Dim lI As Double                                'Contador.
    Dim sqlResposta As String                       'Comando SELECT.
    Dim rsPerguntaResposta As clsYasRecordSet       'Tabela: Tbm_PerguntaResposta.

    On Error GoTo TrataErro

    gfObterPerguntaResposta = False

    '1. Verifica par�metros.
    'pCodigoRamo.
    If Not gfPreenchido(pCodigoRamo) Then
        pMensagem = "FTE0068 - Erro ao obter ramo (c�digo do ramo n�o preenchido)."
        Exit Function
    End If
    
    'pCodigoModalidade
    If Not gfPreenchido(pCodigoModalidade) Then
        pMensagem = "FTE0068 - Erro ao obter plano (plano n�o preenchido)."
        Exit Function
    End If

    'pCodigoTarifa
    If Not gfPreenchido(pCodigoTarifa) Then
        pMensagem = "FTE0068 - Erro ao obter tarifa (tarifa n�o preenchida)."
        Exit Function
    End If

    '2. Busca informa��o.
    sqlResposta = Empty
    
    sqlResposta = sqlResposta & "   SELECT                                      " & vbCrLf
    sqlResposta = sqlResposta & "        COD_RAMO                               " & vbCrLf
    sqlResposta = sqlResposta & "       ,COD_MODALIDADE                         " & vbCrLf
    sqlResposta = sqlResposta & "       ,COD_TARIFA                             " & vbCrLf
    sqlResposta = sqlResposta & "       ,COD_PERGUNTA                           " & vbCrLf
    sqlResposta = sqlResposta & "       ,COD_RESPOSTA                           " & vbCrLf
    sqlResposta = sqlResposta & "       ,COD_TEXTO                              " & vbCrLf
    sqlResposta = sqlResposta & "   FROM                                        " & vbCrLf
    sqlResposta = sqlResposta & "       TBM_PERGUNTARESPOSTA  WITH(NOLOCK)      " & vbCrLf
    sqlResposta = sqlResposta & "   WHERE                                       " & vbCrLf
    sqlResposta = sqlResposta & "       COD_RAMO            =   " & CStr(pCodigoRamo) & vbCrLf
    sqlResposta = sqlResposta & "       AND COD_MODALIDADE  =   " & CStr(pCodigoModalidade) & vbCrLf
    sqlResposta = sqlResposta & "       AND COD_TARIFA      =   " & CStr(pCodigoTarifa) & vbCrLf
    sqlResposta = sqlResposta & "       AND COD_TEXTO       =   " & CStr(pIdTexto) & vbCrLf

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdDiversos, pMensagem) = False Then
        Exit Function
    End If

    If gfObtRegistro(bdDiversos, sqlResposta, rsPerguntaResposta, pMensagem) = False Then
        Exit Function
    End If

    If rsPerguntaResposta.RecordCount = 0 Then
        Call Me.gpFecha2(bdDiversos)
        ReDim pSaida(1, 6)
        gfObterPerguntaResposta = True
        Exit Function
    End If

    rsPerguntaResposta.MoveLast

    ReDim pSaida(rsPerguntaResposta.RecordCount - 1, 6)
    lI = 0

    rsPerguntaResposta.MoveFirst

    While Not rsPerguntaResposta.EOF
        
        pSaida(lI, 1) = rsPerguntaResposta!Cod_Ramo
        pSaida(lI, 2) = rsPerguntaResposta!Cod_Modalidade
        pSaida(lI, 3) = rsPerguntaResposta!Cod_Tarifa
        pSaida(lI, 4) = rsPerguntaResposta!COD_PERGUNTA
        pSaida(lI, 5) = rsPerguntaResposta!COD_RESPOSTA
        pSaida(lI, 6) = rsPerguntaResposta!Cod_Texto
        
        lI = lI + 1
        
        rsPerguntaResposta.MoveNext
    Wend

    Call Me.gpFecha2(bdDiversos)
    
    gfObterPerguntaResposta = True
    
SAIDA:
    Set bdDiversos = Nothing
    Set rsPerguntaResposta = Nothing
    
    Exit Function
    
TrataErro:
    gfObterPerguntaResposta = False
    If gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " gfObterPerguntaResposta - " & ERROR
    End If
    GoTo SAIDA
    
End Function
'8000042754 - Resposta dada a quest�o parametrizada para o plano.
Public Function gfObterRespostaPergunta(ByVal pNossoNumero As String, _
                                       ByVal pNumeroItem As Integer, _
                                       ByRef pSaida() As Variant, _
                                       ByRef pMensagem As String, _
                                       ByVal pPropTabela As String) As Boolean
                                       
                                      
    Dim bdDiversos As ADODB.Connection
    Dim lI As Double                                'Contador.
    Dim sqlResposta As String                       'Comando SELECT.
    Dim rsResposta As clsYasRecordSet               'Tabela: TAB_PED_LOC_COMPL.

    On Error GoTo TrataErro

    gfObterRespostaPergunta = False

    '1. Verifica par�metros.
    'pNossoNumero.
    If Not gfPreenchido(pNossoNumero) Then
        pMensagem = "FTE0068 - Erro ao obter nosso n�mero (nosso n�mero n�o preenchido)."
        Exit Function
    End If
    
    'pNumeroItem
    If Not gfPreenchido(pNumeroItem) Then
        pMensagem = "FTE0068 - Erro ao obter n�mero do item (n�mero do item n�o preenchido)."
        Exit Function
    End If

    '2. Busca informa��o.
    sqlResposta = Empty
    
    sqlResposta = sqlResposta & "       SELECT                                                  " & vbCrLf
    sqlResposta = sqlResposta & "            NOSSO_NUMERO               ,QTD_VIDAS              " & vbCrLf
    sqlResposta = sqlResposta & "           ,COD_ORG                    ,VAL_LMG_VIDA           " & vbCrLf
    sqlResposta = sqlResposta & "           ,COD_RAMO                   ,IND_AUX_FUNERAL        " & vbCrLf
    sqlResposta = sqlResposta & "           ,NUM_PED                    ,VAL_LMG_FUNERAL        " & vbCrLf
    sqlResposta = sqlResposta & "           ,NUM_PI                     ,IND_INC_CONJUGE        " & vbCrLf
    sqlResposta = sqlResposta & "           ,NUM_APOL                   ,VAL_AUX_CESTA          " & vbCrLf
    sqlResposta = sqlResposta & "           ,NUM_ENDO                   ,POSSUI_ELEVADOR        " & vbCrLf
    sqlResposta = sqlResposta & "           ,NUM_ITEM                   ,BONUSALTERADO          " & vbCrLf
    sqlResposta = sqlResposta & "           ,IND_VERBASEP               ,FLG_ENVIARCERTIFICADO  " & vbCrLf
    sqlResposta = sqlResposta & "           ,VAL_VR_LC                  ,TXT_OBSERVACOES        " & vbCrLf
    sqlResposta = sqlResposta & "           ,VAL_LMG_LC                 ,IND_STATUSINSPECAO     " & vbCrLf
    sqlResposta = sqlResposta & "           ,PER_IS_VR_LC               ,VAL_VR_ORIG            " & vbCrLf
    sqlResposta = sqlResposta & "           ,COEF_IS_VR_LC              ,VAL_BAS_ORIG           " & vbCrLf
    sqlResposta = sqlResposta & "           ,TIP_PLANO_VIDA             ,VAL_VR_LC_ORIG         " & vbCrLf
    sqlResposta = sqlResposta & "           ,COD_PLANO_COBERTURA        ,ITEM_ORIGEM_C4         " & vbCrLf
    sqlResposta = sqlResposta & "           ,COD_RESPOSTA                                       " & vbCrLf
    sqlResposta = sqlResposta & "       FROM                                                    " & vbCrLf
    sqlResposta = sqlResposta & pPropTabela & "TAB_PED_LOC_COMPL  WITH(NOLOCK)                  " & vbCrLf
    sqlResposta = sqlResposta & "       WHERE                                                   " & vbCrLf
    sqlResposta = sqlResposta & "               NOSSO_NUMERO = '" & CStr(pNossoNumero) & "'" & vbCrLf
    sqlResposta = sqlResposta & "           AND NUM_ITEM = " & CStr(pNumeroItem) & vbCrLf

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdDiversos, pMensagem) = False Then
        Exit Function
    End If

    If gfObtRegistro(bdDiversos, sqlResposta, rsResposta, pMensagem) = False Then
        Exit Function
    End If

    If rsResposta.RecordCount = 0 Then
        Call Me.gpFecha2(bdDiversos)
        ReDim pSaida(1, 6)
        gfObterRespostaPergunta = True
        Exit Function
    End If

    rsResposta.MoveLast

    ReDim pSaida(rsResposta.RecordCount - 1, 31)
    lI = 0

    rsResposta.MoveFirst

    While Not rsResposta.EOF
        
        pSaida(lI, 1) = rsResposta!nosso_numero
        pSaida(lI, 2) = rsResposta!Cod_Org
        pSaida(lI, 3) = rsResposta!Cod_Ramo
        pSaida(lI, 4) = rsResposta!Num_Ped
        pSaida(lI, 5) = rsResposta!NUM_PI
        pSaida(lI, 6) = rsResposta!Num_Apol
        pSaida(lI, 7) = rsResposta!Num_Endo
        pSaida(lI, 8) = rsResposta!Num_Item
        pSaida(lI, 9) = rsResposta!Ind_VerbaSep
        pSaida(lI, 10) = rsResposta!Val_Vr_Lc
        pSaida(lI, 11) = rsResposta!Val_LMG_Lc
        pSaida(lI, 12) = rsResposta!Per_IS_Vr_Lc
        pSaida(lI, 13) = rsResposta!Coef_IS_Vr_Lc
        pSaida(lI, 14) = rsResposta!Tip_Plano_Vida
        pSaida(lI, 15) = rsResposta!Cod_Plano_Cobertura
        pSaida(lI, 16) = rsResposta!COD_RESPOSTA
        pSaida(lI, 17) = rsResposta!Qtd_Vidas
        pSaida(lI, 18) = rsResposta!Val_LMG_Vida
        pSaida(lI, 19) = rsResposta!Ind_Aux_Funeral
        pSaida(lI, 20) = rsResposta!Val_LMG_Funeral
        pSaida(lI, 21) = rsResposta!Ind_Inc_Conjuge
        pSaida(lI, 22) = rsResposta!Val_Aux_Cesta
        pSaida(lI, 23) = rsResposta!Possui_Elevador
        pSaida(lI, 24) = rsResposta!BonusAlterado
        pSaida(lI, 25) = rsResposta!Flg_EnviarCertificado
        pSaida(lI, 26) = rsResposta!Txt_Observacoes
        pSaida(lI, 27) = rsResposta!Ind_StatusInspecao
        pSaida(lI, 28) = rsResposta!Val_Vr_Orig
        pSaida(lI, 29) = rsResposta!Val_Bas_Orig
        pSaida(lI, 30) = rsResposta!Val_Vr_Lc_Orig
        pSaida(lI, 31) = rsResposta!ITEM_ORIGEM_C4
        
        lI = lI + 1
        
        rsResposta.MoveNext
    Wend

    Call Me.gpFecha2(bdDiversos)
    
    gfObterRespostaPergunta = True
    
SAIDA:
    Set bdDiversos = Nothing
    Set rsResposta = Nothing
    
    Exit Function
    
TrataErro:
    gfObterRespostaPergunta = False
    If gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " gfObterRespostaPergunta - " & ERROR
    End If
    GoTo SAIDA
    
End Function

Public Function mfSalvarPIAutoProgresso(ByRef pStcPed As stcA46V702B, pMensagem As String, pPropTabela As String) As Boolean
    Dim lMensagem As String
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstProgresso As clsYasRecordSet
    Dim lSelect As String

    On Error GoTo ControleErro

    '1. Abre arquivo.
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    '3. Obt�m a Tip_emissao e Tarifa
    lSelect = "SELECT * FROM Tab_Ped_Integracao_Auto_Multi WHERE NUM_PI_RE = '" & pStcPed.NUM_PI & "'"

    If gfObtRegistro(BDP00Tarifa, lSelect, rstProgresso, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If
    
    If rstProgresso.EOF Then
        lSelect = "INSERT INTO Tab_Ped_Integracao_Auto_Multi (Num_PI_Auto, NUM_PI_RE, Cod_Tarif_Auto, Per_Juro_Parc, Qtd_Parc) VALUES (" & pStcPed.Num_PI_Auto_Progresso & ", " & pStcPed.NUM_PI & ", " & pStcPed.Cod_Tarif_Auto & ", " & pStcPed.Per_Juro_Progresso & ", " & pStcPed.Qtd_Parc_Progresso & ")"
        If gfExeSQL(BDP00Tarifa, lSelect, lMensagem) <> 0 Then
            Call gpGraLog(1, lMensagem)
            GoTo SAIDA
        End If
    End If

    Call gpFecha2(BDP00Tarifa)



    mfSalvarPIAutoProgresso = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    If gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gpGraLog(0, "mfSalvarPIAutoProgresso  ERRO - " & Err & " " & Err.Description & _
                             Chr(13) & pMensagem)

    GoTo SAIDA
FechaObj:
    Set BDP00Tarifa = Nothing
    Set rstProgresso = Nothing
    Return
End Function

Public Function mfObterPIAutoProgresso(ByRef pStcPed As stcA46V702B, pMensagem As String, pPropTabela As String) As Boolean
    Dim lMensagem As String
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstProgresso As clsYasRecordSet
    Dim lSelect As String

    On Error GoTo ControleErro

    '1. Abre arquivo.
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    '3. Obt�m a Tip_emissao e Tarifa
    lSelect = "SELECT * FROM Tab_Ped_Integracao_Auto_Multi WHERE NUM_PI_RE = '" & pStcPed.NUM_PI & "'"

    If gfObtRegistro(BDP00Tarifa, lSelect, rstProgresso, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If
    
    If Not rstProgresso.EOF Then
        pStcPed.Num_PI_Auto_Progresso = rstProgresso("NUM_PI_AUTO")
        pStcPed.Cod_Tarif_Auto = rstProgresso("Cod_Tarif_Auto")
        pStcPed.Per_Juro_Progresso = rstProgresso("Per_Juro_Parc")
        pStcPed.Qtd_Parc_Progresso = rstProgresso("Qtd_Parc")
        
        pStcPed.Qtd_Parc = pStcPed.Qtd_Parc_Progresso
        pStcPed.Per_Juro_Frac = pStcPed.Per_Juro_Progresso
    Else
        pStcPed.Num_PI_Auto_Progresso = 0
        pStcPed.Cod_Tarif_Auto = 0
        pStcPed.Per_Juro_Progresso = 0
        pStcPed.Qtd_Parc_Progresso = 0
    End If

    Call gpFecha2(BDP00Tarifa)



    mfObterPIAutoProgresso = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    If gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call gpGraLog(0, "mfObterPIAutoProgresso  ERRO - " & Err & " " & Err.Description & _
                             Chr(13) & pMensagem)

    GoTo SAIDA
FechaObj:
    Set BDP00Tarifa = Nothing
    Set rstProgresso = Nothing
    Return
End Function

Public Function gTipEmissao(ByVal lNossoNumeroAnt As String, lTipEmissao As Integer, lCodTarif As Integer, lPropTabela As String)
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstTarifa As clsYasRecordSet
    Dim lSelect As String
    Dim lMensagem As String

    '1. Abre arquivo.
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    '3. Obt�m a Tip_emissao e Tarifa
    lSelect = "select COD_IDEN_TARIF ,Tip_emissao from " & lPropTabela & "Tab_Ped_loc where nosso_numero = '" & lNossoNumeroAnt & "'"

    If gfObtRegistro(BDP00Tarifa, lSelect, rstTarifa, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    Call gpFecha2(BDP00Tarifa)

    If Not rstTarifa.EOF Then
        lTipEmissao = rstTarifa("Tip_emissao")
        lCodTarif = rstTarifa("COD_IDEN_TARIF")
    End If

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set BDP00Tarifa = Nothing
    Set rstTarifa = Nothing
    Return
End Function

Sub mfCarregarDePara()
    Dim lSel_De_Para As String
    Dim lP00Multi_Tplano As ADODB.Connection
    Dim lMensagem As String
    Dim Rst_Tab_Ped As ADODB.Recordset
        
    Set Rst_Tab_Ped = New ADODB.Recordset
    
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lP00Multi_Tplano, lMensagem) = False Then Call gpGraLog(2, lMensagem): Exit Sub
    
    lSel_De_Para = "EXEC dbo.proc_Buscar_De_Para "
    lSel_De_Para = lSel_De_Para & gCod_Ramo & ","
    lSel_De_Para = lSel_De_Para & M46V111.gCod_plano & ","
    lSel_De_Para = lSel_De_Para & gCodProduto & ","
    lSel_De_Para = lSel_De_Para & "NULL,"
    lSel_De_Para = lSel_De_Para & "NULL,"
    lSel_De_Para = lSel_De_Para & "NULL"

    On Error Resume Next
    Err.Clear
    
    If gRst_Tbm_De_Para.State = adStateOpen Then
        Call gpGraLog(0, "mfCarregarDePara -> Ser� fechado SQL Source= " & gRst_Tbm_De_Para.Source)
        gRst_Tbm_De_Para.Close
    End If
    
    gRst_Tbm_De_Para.CursorLocation = adUseClient
    Call gRst_Tbm_De_Para.Open(lSel_De_Para, lP00Multi_Tplano)    'Este recordSet esta usando a propriedade Filter (nao implementado no YasRecordSet)
    Set gRst_Tbm_De_Para.ActiveConnection = Nothing
    If Err <> 0 Then
        Call gpGraLog(2, Err.Number & " - " & Err.Description): Exit Sub
    End If

    Set lP00Multi_Tplano = Nothing
End Sub

Public Function gfBuscarDePara(ByVal pRetorno As String, ByRef pTipoDePara As Integer, Optional ByVal pCompl As String, Optional pRetornaComplemento As Boolean) As String
    'Fun��o reponsavel por trazer o De Para correspondente ao codigo recebido por par�metro.
        'FILTRO DE/PARA
        'TIPO_DE_PARA =  1 - Coberturas
        '                2 - Classifica��es
        '                3 - Planos
        If gCodProduto >= 1080 Then
            If gRst_Tbm_De_Para.State = 0 Then
                If pCompl <> "" Then
                    gfBuscarDePara = pCompl
                Else
                    gfBuscarDePara = pRetorno
                End If
                Exit Function
            End If
            If gRst_Tbm_De_Para.RecordCount > 0 Then
                gRst_Tbm_De_Para.MoveFirst
                If pCompl = "" Then
                    gRst_Tbm_De_Para.Filter = "DE = '" & pRetorno & "' AND TIPO_DE_PARA = " & pTipoDePara
                    If gRst_Tbm_De_Para.RecordCount > 0 Then
                        gfBuscarDePara = gRst_Tbm_De_Para!Para
                    Else
                        gfBuscarDePara = pRetorno
                    End If
                Else
                    gRst_Tbm_De_Para.Filter = "DE = '" & pRetorno & "' AND DE_COMPL = '" & pCompl & "' AND TIPO_DE_PARA = " & pTipoDePara
                    
                    If gRst_Tbm_De_Para.RecordCount > 0 Then
                        If pRetornaComplemento = False Then
                            gfBuscarDePara = gRst_Tbm_De_Para!Para
                        Else
                            gfBuscarDePara = gRst_Tbm_De_Para!Para_Compl
                        End If
                    Else
                        If pRetornaComplemento Then
                            gfBuscarDePara = pCompl
                        Else
                            gfBuscarDePara = pRetorno
                        End If
                    End If
                                        
                End If

                gRst_Tbm_De_Para.Filter = adFilterNone
            Else
                If pCompl <> "" Then
                    gfBuscarDePara = pCompl
                Else
                    gfBuscarDePara = pRetorno
                End If

            End If
        Else
            If pCompl <> "" Then
                gfBuscarDePara = pCompl
            Else
                gfBuscarDePara = pRetorno
            End If
        End If

End Function

Public Function gfObtPesquisaCadin(ByVal pCnpjCPF As Double, ByVal pTipoPessoa As e_TipPessoa, ByRef pCodigoCadin As String, ByRef pMensagem As String) As Boolean
'Fun��o: obt�m dados de pesquisa do cadin
                                    
    Dim bdYas_ND As ADODB.Connection
    Dim lI As Double  'Contador.
    Dim lSelect As String       'Comando SELECT.
    Dim rsCadin As clsYasRecordSet    'Procedure: PROC_PESQUISA_CADIN.

    gfObtPesquisaCadin = False
    pCodigoCadin = ""

    '2. Busca informa��o.
    lSelect = "EXEC YAS.PROC_PESQUISA_CADIN " & pCnpjCPF & "," & pTipoPessoa
    If gfAbrBasDados("YAS_ND", gTipYAS_ND, gEndYAS_ND, bdYas_ND, pMensagem) = False Then
        GoTo Fim
    End If

    If gfObtRegistro(bdYas_ND, lSelect, rsCadin, pMensagem) = False Then
        GoTo Fim
    End If

    If rsCadin.RecordCount > 0 Then
        pCodigoCadin = CStr(rsCadin!COD_IRREG & " - " & rsCadin!SUB_COD_IRREG)
        GoTo Fim
    End If

Fim:
    Call Me.gpFecha2(bdYas_ND)
    gfObtPesquisaCadin = True
End Function

Public Function gfObterCodSegmentado(lCod_Clasf As String, lCod_Compl_Clasf As String) As Integer
    Dim lCN_RamosDiversos As ADODB.Connection
    Dim lrs As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "gfObterCodSegmentado"
    lKeyDynaTrace = DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    On Error GoTo TrataErro

    gfObterCodSegmentado = 0
    
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lCN_RamosDiversos, lMensagem) = False Then GoTo TrataErro

    lSelect = "EXEC PROC_VERIFICA_SEGMENTADO '" & lCod_Clasf & "','"
    lSelect = lSelect & lCod_Compl_Clasf & "',"
    lSelect = lSelect & gCod_Ramo & ","
    lSelect = lSelect & gCodProduto
    
    If gfObtRegistro(lCN_RamosDiversos, lSelect, lrs, lMensagem) = False Then GoTo TrataErro

    If lrs.RecordCount > 0 Then
       gfObterCodSegmentado = lrs!Cod_Segmentacao
    End If
    
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    Set lCN_RamosDiversos = Nothing
    Set lrs = Nothing
    
    Exit Function
TrataErro:
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
    Call gpGraLog(2, lMensagem)
    GoTo SAIDA
End Function

Public Function gfObterLimComissMax(lCod_Ramo As Integer, ByVal lVal_Lmi As Double) As Double

    On Error GoTo TrataErro

    gfObterLimComissMax = 0
    
    If gCodProduto >= 1082 Then
        Select Case lCod_Ramo
            Case 112
                If lVal_Lmi > CDbl(20000000) Then
                    gfObterLimComissMax = 30
                End If
            Case 114
                If lVal_Lmi > CDbl(20000000) Then
                    gfObterLimComissMax = 40
                End If
        End Select
    End If
    
SAIDA:
    Exit Function
TrataErro:
    Call gpGraLog(2, "Origem Erro: gfObterLimComissMax - " & Err.Description)
    GoTo SAIDA
End Function

Public Function BuscaDadosAdicionais(ByRef pbdTAB_Cobert As ADODB.Connection, ByRef pMensagem As String) As Boolean
    Dim lSelect As String
    BuscaDadosAdicionais = False

    On Error GoTo TrataErro

'    If gPrograma = "D046O003" Then
'        M46V111.gCod_Regiao = 0
'        M46V111.gCod_plano = Me.ObjPai.Cod_Estip
'        gCodProduto = Me.ObjPai.Cod_Produto
'    End If

    lSelect = " Select distinct COD_TARIF_INI, cod_Ramo, Cod_grupo_cobert, DSC_Cobert, Cod_Cobert, Per_Lim_Max_Cobert, COD_LIM_COBERT, TAB_COBERT  , PER_FRANQ , VAL_FRANQ_MIN, Val_Prem_Min, Val_Lim_Max_Cobert, Val_Lim_Min_Cobert from TAB_COBERT WITH (NOLOCK)"
    lSelect = lSelect & " Where cod_Ramo = " & gCod_Ramo
    lSelect = lSelect & " and tab_cobert = " & M46V111.gCod_plano
    lSelect = lSelect & " and COD_TARIF_INI = " & gCodProduto
    lSelect = lSelect & " and COD_TARIF_TER = " & gCodProduto

    If gRst_Tab_Cobert.State = adStateOpen Then
        If gRst_Tab_Cobert.Source = lSelect Then
            GoTo Fim
        Else
            gRst_Tab_Cobert.Close
        End If
    End If

    gRst_Tab_Cobert.CursorLocation = adUseClient
    Call gRst_Tab_Cobert.Open(lSelect, pbdTAB_Cobert)    'Este recordSet esta usando a propriedade Filter (nao implementado no YasRecordSet)
    Set gRst_Tab_Cobert.ActiveConnection = Nothing

'    If M46V999.gfObtRegistro(pbdTAB_Cobert, lSelect, prsTab_Cobert, pMensagem) = False Then
'        Call M46V999.gpFecha2(pbdTAB_Cobert)
'        Call M46V999.gpGraLog(2, pMensagem)
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If

Fim:
    BuscaDadosAdicionais = True
    Exit Function

TrataErro:
        Call gpFecha2(pbdTAB_Cobert)
        Call gpGraLog(2, Err.Description)
        Screen.MousePointer = vbDefault
        Exit Function

End Function

Public Function gfObterCnpjCoss(lCodCoss As Integer) As String
    Dim lCN_Yas_ND As ADODB.Connection
    Dim lrs As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String

    On Error GoTo TrataErro
    
    If gfAbrBasDados("YAS_ND", gTipYAS_ND, gEndYAS_ND, lCN_Yas_ND, lMensagem) = False Then GoTo TrataErro

    lSelect = "EXEC PROC_CONSULTA_DADOS_COSS " & lCodCoss
    
    If gfObtRegistro(lCN_Yas_ND, lSelect, lrs, lMensagem) = False Then GoTo TrataErro

    If lrs.RecordCount > 0 Then
        If lrs(0).Value Like "ERRO*" Then
            lMensagem = lrs(0).Value
            GoTo TrataErro
        Else
            gfObterCnpjCoss = lrs!NUM_CNPJ
        End If
    End If
SAIDA:
    Set lCN_Yas_ND = Nothing
    Set lrs = Nothing
    
    Exit Function
TrataErro:
    Call gpGraLog(2, lMensagem)
    GoTo SAIDA
End Function

Public Function mfObtemBancoBoleto(ByVal pTip_Emissao As Long, _
                            ByVal pCPFCNPJ As String, _
                            ByVal pCodProduto As Long, _
                            ByVal pCodSegurado As Long, _
                            ByVal pNumApol As Long, _
                            ByVal pCorCorr As Long, _
                            ByVal pDatMov As Long) As Boolean
Dim lrs         As clsYasRecordSet
Dim lSelect     As String
Dim lretorno    As Boolean
Dim lCN         As ADODB.Connection
Dim UltimoErro As String

On Error GoTo TrataErro:    lretorno = False: UltimoErro = ""

    lSelect = " BEGIN"
    lSelect = lSelect & " DECLARE @return_value int,"
    lSelect = lSelect & "         @COD_RETORNO decimal(4, 0),"
    lSelect = lSelect & "         @MSG_RETORNO varchar(80),"
    lSelect = lSelect & "         @STA_SEGREGACAO decimal(4, 0)"
    lSelect = lSelect & " EXEC    @return_value = [YAS].[PROC_VALIDA_REGRAS_SEGREGACAO]"
    lSelect = lSelect & "         @DATA_MOVIMENTO = '" & IIf(pDatMov = 0, Format(Now, "yyyymmdd"), pDatMov) & "',"
    lSelect = lSelect & "         @COD_RAMO = " & gCod_Ramo & ","
    lSelect = lSelect & "         @NUM_CPFCNPJ_SEG = " & gfFormatarCamposSQL(pCPFCNPJ) & ","
    lSelect = lSelect & "         @COD_PRODUTO = " & pCodProduto & "1,"
    lSelect = lSelect & "         @NUM_CLIENTE_ESP = " & pCodSegurado & ","
    lSelect = lSelect & "         @NUM_APOLICE = " & pNumApol & ","
    lSelect = lSelect & "         @COD_CORRETOR = " & pCorCorr & ","
    lSelect = lSelect & "         @COD_RETORNO = @COD_RETORNO OUTPUT,"
    lSelect = lSelect & "         @MSG_RETORNO = @MSG_RETORNO OUTPUT,"
    lSelect = lSelect & "         @STA_SEGREGACAO = @STA_SEGREGACAO OUTPUT"
    lSelect = lSelect & " SELECT @STA_SEGREGACAO as N'STA_SEGREGACAO'"
    lSelect = lSelect & " END;"
        
    If gfAbrBasDados("YAS_ND", gTipYAS_ND, gEndYAS_ND, lCN, UltimoErro) = False Then GoTo TrataErro

    If gfObtRegistro(lCN, lSelect, lrs, UltimoErro) = False Then GoTo TrataErro
    
    lretorno = IIf(gfTratarNulo(lrs("STA_SEGREGACAO"), tpoNumerico) = 1, True, False)
    
    mfObtemBancoBoleto = lretorno
SAIDA:
    Set lCN = Nothing
    Set lrs = Nothing
    
    Exit Function
TrataErro:
    If Not UltimoErro = "" Then
        UltimoErro = Err & " - " & ERROR
    End If
    GoTo SAIDA
End Function

Public Function gObtemBoletoSegregacao(ByVal pTip_Emissao As Long, _
                                        ByVal pCPFCNPJ As String, _
                                        ByVal pCodProduto As Long, _
                                        ByVal pCodSegurado As Long, _
                                        ByVal pNumApol As Long, _
                                        ByVal pCorCorr As Long, _
                                        ByVal pDatMov As Long, _
                                        ByVal pTipDoc As e_TipDoc) As Integer

    If Not gfPreenchido(pCPFCNPJ) Or Trim(pCPFCNPJ) = "0" Or pTipDoc <> e_TipDoc_Proposta And gPrograma = "D046O003" Then
        gObtemBoletoSegregacao = 0
        Exit Function
    End If
                                        
    If pTip_Emissao < 300 Then
        pNumApol = 0
    End If
    
    If mfObtemBancoBoleto(pTip_Emissao, pCPFCNPJ, pCodProduto, pCodSegurado, pNumApol, pCorCorr, pDatMov) Then
        gObtemBoletoSegregacao = 33 'Santander
    '2016-08-22: Alterar Itau p Santander
    ElseIf pCorCorr = 15701 Then 'Corretor Mercedes
        gObtemBoletoSegregacao = 33 'Santander, old: 341 - Itau
    '2016-08-22
    Else
        gObtemBoletoSegregacao = 399 'Hsbc
    End If

End Function

Public Function gfCopDev(ByVal pAntNosNumero As String, ByRef pMensagem As String) As String
'Fun��o: monta comando INSERT gen�rico para dados de proposta.

'Par�metros de entrada...pAntNumItem....Anterior: n�mero do item.
'                        pmfNovNumItem....Novo: n�mero do item.
'Par�metro de sa�da......pMensagem......Mensagem.

    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "gfCopDev"
    lKeyDynaTrace = DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim bdBasDados As ADODB.Connection    'Base de dados.
    Dim flTabela As clsYasField   'Campos da tabela.
    Dim rsTabela As clsYasRecordSet    'Registro da tabela.
    Dim lCount As Long        'Utilizado no for...next.
    Dim lIdeTabela As String          'Nome da tabela.
    Dim lContinuar As Boolean         'Indicador se continua o procedimento ou n�o.
    Dim lSelect As String
    Dim lInsert As String
    Dim lPropTabela As String
    Dim lTipEmissao As Integer
    Dim lCodTarifaCopia As Integer
    Dim lCodRamo As Integer
    Dim lCodClasf As String
    Dim lCodClasfCompl As String
    Dim lCarregarDePara As Boolean
    Dim lMsgErr As String
    Dim lInsertTodo As String
    
    On Error GoTo ControleErro

    gfCopDev = Empty
    lInsertTodo = Empty

    Call gfObtSYASMULTINI(lPropTabela)

    '07/01/2010 - Ficha: 0901900 - C�pia de cota��o/proposta c/ �ltima tarifa -- inicio
    Dim lCodProduto As Integer      'C�digo do produto.
    Dim lCodProtocolo As Long           'C�digo do protocolo.
    Dim lCodSequencia As Long           'C�digo da seq��ncia.
    Dim lDesTarifa As String      'Descri��o da tarifa.

    '1. Obt�m registros da tabela.
    If UCase(Left(pAntNosNumero, 1)) = UCase("w") Then
        pAntNosNumero = Right(pAntNosNumero, Len(pAntNosNumero) - 1)
        Set bdBasDados = New ADODB.Connection
        bdBasDados.Open "Provider=SQLOLEDB.1;Password=yasworkflow;Persist Security Info=True;User ID=workflow;Initial Catalog=SYASPORTAL;Data Source=" & gSerSyasPortal & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=CYBER07;Use Encryption for Data=False;Tag with column collation when possible=False", "workflow", "yasworkflow"
        
    Else
        If gfAbrBasDados("P00MPED", gTipP00MPED, gEndP00MPED, bdBasDados, pMensagem) = False Then
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        
        If UCase(Left(pAntNosNumero, 1)) = UCase("f") Then
            pAntNosNumero = Right(pAntNosNumero, Len(pAntNosNumero) - 1)
            lPropTabela = "FTP_"
        End If
    End If
    
    If Len(pAntNosNumero) > 10 Then
        lSelect = "SELECT Num_Apol_An FROM " & lPropTabela & "Tab_Ped WITH (NOLOCK) WHERE Nosso_Numero = '" & pAntNosNumero & "'"
        If gfObtRegistro(bdBasDados, lSelect, rsTabela, pMensagem) = False Then
            Call Me.gpFecha2(bdBasDados)
            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        
        If val("" & rsTabela!Num_Apol_An) > 0 Then
            lInsertTodo = "/* INICIO APOLICE */" & vbCrLf
            lInsertTodo = lInsertTodo & gfCopDev(val("" & rsTabela!Num_Apol_An), pMensagem)
            lInsertTodo = lInsertTodo & vbCrLf & "/* FIM APOLICE */" & vbCrLf
        End If
    End If
    
    For lCount = 1 To 32
        lContinuar = True
        Select Case lCount
        Case 1
            lIdeTabela = "TAB_PED"
        Case 2
            lIdeTabela = "TAB_PED_CORR"
        Case 3
            lIdeTabela = "TAB_PED_COSS"
        Case 4
            lIdeTabela = "TAB_PED_JUROS"
        Case 5
            lIdeTabela = "TAB_PED_VENC"
        Case 6
            lIdeTabela = "TAB_PED_ENDER"
        Case 7
            lIdeTabela = "TAB_CTRL_EMIS"
        Case 8
            lIdeTabela = "TAB_PED_LOC"
        Case 9
            lIdeTabela = "TAB_PED_CLAU"
        Case 10
            lIdeTabela = "TAB_PED_COBERT"
        Case 11
            lIdeTabela = "TAB_PED_LOC_BEM"
        Case 12
            lIdeTabela = "TAB_PED_LOC_SEGURO"
        Case 13
            lIdeTabela = "TAB_PED_OBS"
        Case 14
            lIdeTabela = "TAB_PED_INSP"
        Case 15
            lIdeTabela = "TAB_PED_AUTO"
        Case 16
            lIdeTabela = "TAB_PED_COBERT_COMPL"
        Case 17
            lIdeTabela = "TAB_PED_FAT_RISCO"
        Case 18
            lIdeTabela = "TAB_PED_DIZER"
        Case 19
            lIdeTabela = "TAB_PED_ESPEC"
        Case 20
            lIdeTabela = "TAB_PED_ESPEC2"
        Case 21
            lIdeTabela = "Tab_Ped_Loc_LMI"
        Case 22
            lIdeTabela = "Tab_Canal_Emiss"
        Case 23
            lIdeTabela = "TAB_CTRL_EMIS_COMPL"
        Case 24
            lIdeTabela = "TAB_PED_LOC_PESSOAS"
        Case 25
            lIdeTabela = "TAB_PED_TELEFONES"
        Case 26
            lIdeTabela = "TAB_PED_LOC_COMPL"
        Case 27
            lIdeTabela = "TAB_PED_COBERT_FRANQUIA"
        Case 28
            lIdeTabela = "TAB_PED_TEXTOS"
        Case 29
            lIdeTabela = "TAB_PED_LOC_BICICLETA"
        Case 30
            lIdeTabela = "TAB_PED_LIBERACAO_SERASA"
        Case 31
            lIdeTabela = "TAB_PED_ALL_RISKS"
        Case 32
            lIdeTabela = "TAB_PED_OBRAS_ARTE"
        End Select

        
        If Len(pAntNosNumero) > 10 Then
            lSelect = "SELECT * FROM " & lPropTabela & lIdeTabela & " WITH (NOLOCK) "
            lSelect = lSelect & " WHERE Nosso_Numero = '" & pAntNosNumero & "'"
        Else
            lIdeTabela = "GED.." & lIdeTabela & "_Atual"
            lSelect = "SELECT * FROM " & lPropTabela & lIdeTabela & " WITH (NOLOCK) "
            lSelect = lSelect & " WHERE Num_Apol = " & pAntNosNumero
        End If
        
        pMensagem = Empty
        If gfObtRegistro(bdBasDados, lSelect, rsTabela, pMensagem) = False Then
            lContinuar = False
        End If
        
        If lContinuar = True Then
            '2. Grava nova cota��o/proposta.
            While Not rsTabela.EOF
                lCodRamo = 0
                lCodClasf = 0
                lCodClasfCompl = 0

                If lIdeTabela = "TAB_PED" Or lIdeTabela = "TAB_PED_COBERT" Or lIdeTabela = "TAB_PED_COBERT_COMPL" Then
                    lCodRamo = rsTabela!Cod_Ramo
                    If lIdeTabela = "TAB_PED" Then
                        M46V111.gCod_plano = rsTabela!Cod_Estip
                        lCarregarDePara = False
                        
                        lCarregarDePara = lCodProduto >= 1080 And gRst_Tbm_De_Para.State <> adStateOpen
                        If lCodProduto >= 1080 And Not lCarregarDePara Then
                            lCarregarDePara = gRst_Tbm_De_Para.RecordCount = 0
                        End If
                    
                        If lCarregarDePara Then
                            gCodProduto = lCodProduto
                            gCod_Ramo = lCodRamo
                            M46V111.gCod_plano = gfPlanoBasico(lCodRamo, rsTabela!Cod_Estip)
                            Call mfCarregarDePara
                        End If
                    End If
                End If
                
                If lIdeTabela = "TAB_PED_LOC" And lCodProduto >= 1080 Then
                    lCodClasf = rsTabela!Cod_Clasf
                    lCodClasfCompl = rsTabela!Cod_Compl_Clasf
                End If
                
                lInsert = "INSERT INTO " & lPropTabela & lIdeTabela & " ("
                For Each flTabela In rsTabela.Fields
                    If InStr("PED_LOC_PESSOA_ID,TELEFONE_ID, ID_FRANQUIA_AVANCADA", UCase(flTabela.Name)) = 0 Then
                        lInsert = lInsert & flTabela.Name & ", "
                    End If
                Next flTabela
                lInsert = Left$(lInsert, Len(lInsert) - 2) & ") VALUES ("
                For Each flTabela In rsTabela.Fields
                    Select Case UCase(flTabela.Name)
                    Case "PED_LOC_PESSOA_ID", "TELEFONE_ID", "ID_FRANQUIA_AVANCADA"
                    Case "NOSSO_NUMERO" 'Nosso n�mero
                        If InStr(lIdeTabela, "Atual") = 0 Then
                            lInsert = lInsert & "'" & pAntNosNumero & "', "
                        Else
                            lInsert = lInsert & flTabela.Value & ", "
                        End If
                    Case "NUM_PED"
                        lInsert = lInsert & flTabela.Value & ", "                                'N�mero do pedido.
                    Case "SIT_DOCUMENTO", "COD_SIT_PED"
                        lInsert = lInsert & "1, "       'DOCUMENTO N�O CALCULADO.               'Sit_Documento.
                    Case "DAT_LIBER"
                        lInsert = lInsert & Left(Me.gfForData(3), 8) & ", "                        'Dat_Alt.
                    Case "DAT_PRI_CALCULO"
                        lInsert = lInsert & "0, "                                               'Dat_Pri_Calculo.
                    Case "DAT_DIGIT"
                        lInsert = lInsert & Left(Me.gfForData(3), 8) & ", "                        'Dat_Digit
                    Case "COD_USER_LIBER"
                        lInsert = lInsert & "'COPIA',"                                  'Cod_User_Alt.
                    Case "NUM_BLOQ"
                        lInsert = lInsert & flTabela.Value & ", "
                    Case "NUM_PI"
                        lInsert = lInsert & flTabela.Value & ", "
                    Case "DAT_REC_CIA"
                        lInsert = lInsert & flTabela.Value & ", "
                    Case "NUM_APOL", "NUM_ENDO", "DAT_EMIS", "COD_IDEN_STAT_ACAO", "BONUSALTERADO"
                        If InStr(lIdeTabela, "Atual") = 0 Then
                            lInsert = lInsert & "0, "
                        Else
                            lInsert = lInsert & flTabela.Value & ", "
                        End If
                    Case "COD_PRODUTO"
                        If lIdeTabela = "TAB_PED" Then
                            lInsert = lInsert & flTabela.Value & ", "
                        Else
                            lInsert = lInsert & flTabela.Value & ", "
                        End If
                    Case "COD_IDEN_TARIF"
                        lInsert = lInsert & flTabela.Value & ", "
                    Case "COD_SEQUENCIA"
                        lInsert = lInsert & flTabela.Value & ", "
                    Case "LMI_UNICO"
                        If IsNull(flTabela.Value) Then
                            lInsert = lInsert & "0, "
                        Else
                            lInsert = lInsert & flTabela.Value & ", "
                        End If
                        'pCalcLMI = gfTratarNulo(flTabela.Value, tpoNumerico)
                    Case "NUM_COTAC"    '1100090 - Inspecao Desabilitada
                        lInsert = lInsert & "'000000000000',"                               'NUM_COTAC.
                    Case Else
                        Select Case flTabela.Type_
                        Case adBigInt, adBinary, adInteger, adLongVarBinary, adSmallInt, _
                             adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, _
                             adUnsignedTinyInt                          'Num�rico - inteiro.
                            If IsNull(flTabela.Value) Then
                                lInsert = lInsert & "0, "
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case adCurrency, adSingle, adDouble, adNumeric             'Num�rico - com v�rgula.
                            If IsNull(flTabela.Value) Then
                                lInsert = lInsert & "0, "
                            Else
                                lInsert = lInsert & gfForValor(1, flTabela.Value) & ", "
                            End If
                        Case adChar, adVarChar, adLongVarChar, adVarWChar                          'Texto.
                            If gfPreenchido(flTabela.Value) Then
                                lInsert = lInsert & "'" & Me.gfForAspas(flTabela.Value) & "', "
                            Else
                                lInsert = lInsert & "' ', "
                            End If
                        Case adBoolean                                  'Sim/N�o.
                            If IsNull(flTabela.Value) Then
                                lInsert = lInsert & "1, "
                            Else
                                lInsert = lInsert & IIf(flTabela.Value, "1", "0") & ", "
                            End If
                        
                        Case adDBTimeStamp
                            lInsert = lInsert & "'" & Format$(flTabela.Value, "YYYYMMDD hh:mm:ss") & "', "
                            
                        Case Else
                            Call Me.gpFecha2(bdBasDados)
                            pMensagem = "FTE0008 - Tabela " & lIdeTabela & " com tipo do campo " & _
                                        flTabela.Name & " diferente dos previstos."
                            Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
                            GoTo SAIDA
                        End Select
                    End Select
                Next flTabela
                If Mid(lInsert, Len(lInsert), 1) = "," Then
                    lInsert = Left(lInsert, Len(lInsert) - 1) & ")"
                Else
                    lInsert = Left(lInsert, Len(lInsert) - 2) & ")"
                End If
                
                lInsertTodo = lInsertTodo & vbCrLf & lInsert
                
                If lMsgErr <> "" Then
                   GoTo ControleErro
                End If

                rsTabela.MoveNext
            Wend
        End If
    Next lCount
    
    gfCopDev = lInsertTodo
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    If lMsgErr <> "" Then
       Call gpGraLog(0, lMsgErr)
       GoTo SAIDA
    End If
    
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfCopDev ERRO - " & Err & " " & Err.Description)
    Call DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfCopProposta ERRO - " & Err & " " & Err.Description, 2, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Call Me.gpFecha2(bdBasDados)
    Set bdBasDados = Nothing
    Set flTabela = Nothing
    Set rsTabela = Nothing
    Return

End Function

Public Sub gpGraScriptDev(ByVal pMensagem As String, ByRef lMsgErr As String)
    Dim fso As FileSystemObject
    Dim file As file
    Dim Texto As TextStream
    Dim lEndLog As String    'Endere�o do arquivo log.
    Dim i As Integer
    Dim lNomeArquivo As String

    On Error GoTo TrataErro

    Err = 0
    lMsgErr = ""
    lNomeArquivo = "SYASMULT_COPIA_COTACAO"
    lEndLog = gAppPath & lNomeArquivo & ".TXT"
    Set fso = New FileSystemObject
    
    'Apaga as copias anteriores
    On Error Resume Next
    fso.DeleteFile (gAppPath & lNomeArquivo & "*")
    
    On Error GoTo TrataErro
    
    i = 0
    
    Do While fso.FileExists(lEndLog)
        i = i + 1
        lEndLog = gAppPath & lNomeArquivo
        lEndLog = lEndLog & "(" & i & ").TXT"
    Loop
    
    Set Texto = fso.CreateTextFile(lEndLog, True, True)
    Texto.Write pMensagem
    Texto.Close
    
    gPathScript = lEndLog
    Exit Sub
    
TrataErro:
    lMsgErr = "[ERR] gpGraScriptDev: " & Err.Number & " - " & Err.Description
End Sub

Public Function gfDisponibilizaCotacaoPortal(ByVal pNossoNumero As String, _
                                        pDesconto As Double, _
                                        ByRef pMensagem As String, pConexao As ADODB.Connection, _
                                        ByRef pNumPed As Long) As Boolean
    Dim lSql As String
    Dim lrs As clsYasRecordSet
    Dim lrst As clsYasRecordSet
    
    lSql = "exec syasportal.dbo.proc_CalNumPedido " & gNomUsuario
    If gfObtRegistro(pConexao, lSql, lrst, pMensagem) = False Then Exit Function

    lSql = "Exec syasportal.dbo.PROC_DisponibilizaSiscotaWeb " & gfFormatarCamposSQL(pNossoNumero) & ","
    lSql = lSql & gfFormatarCamposSQL(pDesconto) & ","
    lSql = lSql & lrst!Ult_Num_Ped
    
    
    If gfObtRegistro(pConexao, lSql, lrs, pMensagem) = False Then Exit Function
    
    If lrs!Erro = "" Then
        pNumPed = lrs!Nro_Pedido
    Else
        pMensagem = lrs!Erro
        Exit Function
    End If
            
    gfDisponibilizaCotacaoPortal = True

End Function

Public Function gfVerificaParcelaAntec(ByRef pStcPed As stcA46V702B) As Boolean
    gfVerificaParcelaAntec = False
    
    If pStcPed.Cod_Reg = 106 Or pStcPed.Cod_Reg = 107 Or pStcPed.Cod_Reg = 120 Or pStcPed.Cod_Reg = 129 Then
        Exit Function
    End If
    
    If pStcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
        If pStcPed.objstcCtrlEmis.Cod_Bco_Boleto > 0 And pStcPed.Cod_Bco_Cobr > 0 Then '1 carne com demais em debito
            gfVerificaParcelaAntec = True
        ElseIf pStcPed.gTipPagamento = e_FormaPagamento_Carne Then 'carne
            gfVerificaParcelaAntec = True
        ElseIf pStcPed.objstcCtrlEmis.Cod_Bco_Boleto = 0 And pStcPed.Cod_Bco_Cobr > 0 Then 'Debito
            gfVerificaParcelaAntec = False
        End If
    End If

End Function

Public Function gfObtemIPC(ByVal pMensagem As String) As Double
    Dim lSql As String
    Dim lrst As clsYasRecordSet
    
    lSql = "SELECT TOP 1 Taxa FROM RamosDiversos..Tbm_TaxaAtualizacaoMonetaria ORDER BY AnoMes DESC"
    If gfObtRegistro(mConexaoSQL, lSql, lrst, pMensagem) = False Then Exit Function
        
    If lrst.RecordCount > 0 Then
        gfObtemIPC = lrst!Taxa
    Else
        gfObtemIPC = 0
        pMensagem = "Erro ao obter Indice IPC."
        Exit Function
    End If

End Function

Public Function gObtemValFranqProRata(pDatIniVig As Long, pDatTerVig As Long, pValFranqMin As Double, pTipFranqDias As Boolean) As Double

    Dim lqtdDias As Integer
    
    lqtdDias = gfDifDatas(pDatIniVig, pDatTerVig)
    If lqtdDias < 365 Then
        gObtemValFranqProRata = lqtdDias * Round(pValFranqMin / 365, 5)
        
        If pTipFranqDias Then ' se a franquia for em dias converte para inteiro o pro-rata
            gObtemValFranqProRata = CInt(pTipFranqDias)
        End If
    Else
        gObtemValFranqProRata = pValFranqMin
    End If
                
SAIDA:

    Exit Function

End Function

Public Property Get gDataEmis_Apol() As Long
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstTarifa As clsYasRecordSet
    Dim lSelect As String
    Dim lMensagem As String

    'Rotina para obter a data de emissao da apolice
    gDataEmis_Apol = "0"

    If gNum_Apolice = 0 Then
        GoTo SAIDA
    End If

    '1. Abre arquivo.
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    '3. Obt�m a Tarifa da apolice
    lSelect = "select dat_Emis from Tab_Ped where NUM_APOL = " & gNum_Apolice & " And Num_Endo = 0 "

    If gfObtRegistro(BDP00Tarifa, lSelect, rstTarifa, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    Call gpFecha2(BDP00Tarifa)

    If Not rstTarifa.EOF Then
        gDataEmis_Apol = rstTarifa("dat_Emis")
    End If

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set BDP00Tarifa = Nothing
    Set rstTarifa = Nothing
    Return

End Property
'*Luiz Alberto - 18/01/2016
'**************************
Public Function gfValidaInspecao(ByRef Num_Item As Double, _
                                 ByRef Num_Insp As Double, _
                                 ByRef num_cnpj_cpf As String, _
                                 ByRef Num_CEP As Double, _
                                 ByRef Num_Logr As Double, _
                                 ByRef pMensagem As String, _
                                 pConexao As ADODB.Connection) As Boolean
    Dim lSql As String
    Dim lrs As clsYasRecordSet
    Dim lrst As clsYasRecordSet

    lSql = "exec ramosdiversos.dbo.Proc_Valida_Inspecao "
    lSql = lSql & "'" & Trim(num_cnpj_cpf) & "'"
    lSql = lSql & ", " & Num_CEP
    lSql = lSql & ", " & Num_Logr
    lSql = lSql & ", " & Num_Insp

    If gfObtRegistro(pConexao, lSql, lrst, pMensagem) = False Then Exit Function
    
    If lrst.RecordCount > 0 Then
        gfValidaInspecao = True
        pMensagem = " N�mero de inspe��o: " & Num_Insp & vbCrLf
        If num_cnpj_cpf <> lrst.Fields("num_cnpj_cpf").Value Then
            pMensagem = pMensagem & " 0 N�mero de CNPJ/CPF: " & num_cnpj_cpf & " diverge do encontrado na base: " & lrst.Fields("num_cnpj_cpf").Value & vbCrLf
            gfValidaInspecao = False
        End If
        If Num_CEP <> lrst.Fields("Num_CEP").Value Then
            pMensagem = pMensagem & " O CEP: " & Format(Num_CEP, "00000-000") & " diverge do encontrado na base: " & Format(lrst.Fields("Num_CEP").Value, "00000-000") & vbCrLf
            gfValidaInspecao = False
        End If
        If Num_Logr <> lrst.Fields("Num_Logr").Value Then
            pMensagem = pMensagem & " O N�mero do Logradouro: " & Num_Logr & " diverge do encontrado na base: " & lrst.Fields("Num_Logr").Value & vbCrLf
            gfValidaInspecao = False
        End If
    Else
        'pMensagem = "O n�mero da inspe��o informada para o item " & Format$(Num_Item, "00000") & " n�o coincide com o local de risco e ou segurado."
        pMensagem = "Inspe��o do item " & Format$(Num_Item, "00000") & " n�o localizada."
        gfValidaInspecao = False
    End If

End Function

Public Sub GetDigitoConta(ByRef pobjstcPed As stcA46V702B, DebConta As String)
    Dim lAuxiliar As String
    Dim lDigito As String
    
    Dim clsA99V001 As clsA99V001
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = Me
    
    Select Case pobjstcPed.Cod_Bco_Cobr
    Case 237
        Call clsA99V001.DIG_BRADESCO(Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000"), lDigito, lAuxiliar)
    Case 341
        Call clsA99V001.DIG_ITAU(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "00000"), lDigito, lAuxiliar)
    Case 347
        If pobjstcPed.gAgenciaDig >= 1000 Then
            Call clsA99V001.DIG_SUDAMERIS_NOVO(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000"), lDigito, lAuxiliar)
        Else
            Call clsA99V001.DIG_SUDAMERIS(Format(pobjstcPed.gAgenciaDig, "000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "000000000"), lDigito, lAuxiliar)
        End If
    Case 409
        Call clsA99V001.DIG_UNIBANCO(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "000000"), lDigito, lAuxiliar)
    Case 479
        Call clsA99V001.DIG_BOSTON(Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000"), lDigito, lAuxiliar)
    Case 356
        Call clsA99V001.DIG_REAL(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000"), lDigito, lAuxiliar)
    Case 399
        Call clsA99V001.DIG_HSBC(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000000"), lDigito, lAuxiliar)
    Case 47
        Call clsA99V001.DIG_BANESE(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000000"), lDigito, lAuxiliar)
    Case 33
        Call clsA99V001.DIG_SANTANDER(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000000"), lDigito, lAuxiliar)
    Case 756
        Call clsA99V001.DIG_BANCOOB(Format(pobjstcPed.Cod_Agen_Bas_Cor, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "000000000"), lDigito, lAuxiliar)
    'Novo Conv�nio banco BRB
    Case 70
        Call clsA99V001.DIG_BANCOBRB(Format(pobjstcPed.Cod_Agen_Bas_Cor, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "000000000"), lDigito, lAuxiliar)
    End Select
        
    If pobjstcPed.Cod_Bco_Cobr = 399 Then
        DebConta = Left(Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000000"), 9) & "-" & Right(Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000000"), 1) & lDigito
    Else
        DebConta = pobjstcPed.Num_Cta_Bas_Cor & "-" & lDigito
    End If
End Sub

''2016-07-20 - Juliano
Public Function gfFormatarCNPJ(ByVal tipPessoa As Integer, ByVal strCNPJ As String) As String
'Fun��o: formata CNPJ 2
'e_TipPessoa_Juridica = 1
'e_TipPessoa_Fisica = 2
    Dim strAux As String, strResult As String
    
    strAux = Replace(Trim(strCNPJ), ".", "")
    strAux = Replace(strAux, "/", "")
    strAux = Replace(strAux, "-", "")
    
    If (tipPessoa = 1) Then
        strAux = Right("00000000000000" & strAux, 14)
        strResult = Left(strAux, 2) & "." & Mid(strAux, 3, 3) & "." & Mid(strAux, 6, 3) & "/" & Mid(strAux, 9, 4) & "-" & Right(strAux, 2)
    Else
        strAux = Right("00000000000000" & strAux, 11)
        strResult = Left(strAux, 3) & "." & Mid(strAux, 4, 3) & "." & Mid(strAux, 7, 3) & "-" & Right(strAux, 2)
    End If
    
    gfFormatarCNPJ = strResult
End Function

Public Function Log_Execucao(ByVal pNum_pi As Double, ByVal pUsuario As String, ByVal pDt_Inicio As Date, _
                             ByVal pHr_Inicio As Date, ByVal pNome_Procedimento As String, ByVal pNivel As Integer, _
                             Optional ByVal pMsgErro As String) As Long
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstTarifa As clsYasRecordSet
    Dim lSelect As String
    Dim lMensagem As String

    '1. Abre arquivo.
    If gfAbrBasDados("SYASPORTAL", gSerSyasPortal, gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If
    
    Dim dt_Fim As Date
    Dim hr_Fim As Date
    
    hr_Fim = Time
    dt_Fim = Date

    '3. Obt�m a Tarifa da apolice
    lSelect = "INSERT INTO Log_Execucao (Num_PI, Usuario, Servidor, Dt_Inicio, Dt_Termino, Nome_Procedimento, Msg_Erro) "
    lSelect = lSelect & " Values ( " & pNum_pi & ", "
    lSelect = lSelect & "'" & pUsuario & "', "
    lSelect = lSelect & "'" & M99V001.GetLoginComputerName & "', "
    lSelect = lSelect & "'" & Format(pDt_Inicio, "yyyy/MM/dd") & " " & Format(pHr_Inicio, "hh:mm:ss") & "', "
    lSelect = lSelect & "'" & Format(dt_Fim, "yyyy/MM/dd") & " " & Format(hr_Fim, "hh:mm:ss") & "', "
    lSelect = lSelect & "'" & pMsgErro & "', "
    lSelect = lSelect & "'" & pNome_Procedimento & "') "
    
    If gfExeSQL(BDP00Tarifa, lSelect, lMensagem) <> 0 Then
        Call gpGraLog(0, "Log_Execucao: " & lMensagem)
        GoTo SAIDA
    End If

    Call gpFecha2(BDP00Tarifa)

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set BDP00Tarifa = Nothing
    Set rstTarifa = Nothing
    Return
End Function




Function gfBuscaNossoNumeroUnicidade(ByVal pTabela As String, ByVal pPI As String, ByVal pNossoNumero As String, _
                                     ByRef rCodUnicidade As Long, ByRef rMensagem As String) As Boolean
    Dim lBanco As ADODB.Connection
    Dim lMensagem As String
    Dim lrsBusca As clsYasRecordSet
    Dim lSelect As String

    rCodUnicidade = 0
    gfBuscaNossoNumeroUnicidade = False
    
    If pPI = "" And pNossoNumero = "" Then
        rMensagem = "Erro: PI ou NossoNumero devem ser informados"
        GoTo SAIDA
    End If

    If gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lBanco, lMensagem) = False Then
        rMensagem = lMensagem
        Call gpGraLog(2, lMensagem)
        GoTo SAIDA
    End If
    
    '' Buscar CodUnicidade pelo PI ou NN
    '' Ordenacao poderia ser na aplicacao, mas o codUnicidade nao eh sequencial
    lSelect = " Select top 1 * from " & pTabela & "TAB_PED (nolock) Where cod_ramo in (112, 113, 114) "
    
    If pPI <> "" Then
        lSelect = lSelect & " and NUM_PI = " & pPI
    Else
        If pNossoNumero <> "" Then
            lSelect = lSelect & " and NOSSO_NUMERO = '" & pNossoNumero & "' "
        End If
    End If
    
    lSelect = lSelect & " order by dat_digit desc, cod_unicidade desc "
    
    If gfObtRegistro(lBanco, lSelect, lrsBusca, lMensagem) = False Then
        rMensagem = lMensagem
        Call gpGraLog(1, lMensagem)
        GoTo SAIDA
    Else
        gfBuscaNossoNumeroUnicidade = True
    End If
    
    If Not lrsBusca.EOF Then
        rCodUnicidade = lrsBusca("Cod_Unicidade")
    End If

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lrsBusca = Nothing
    Return
End Function

Public Function gEhRetornoSiscota(ByVal pNosso_Numero As String) As Boolean
    Dim lobjstcA28V721B As stcA28V721B
    Dim lobjA28V721A As clsA28V721A
    
    gEhRetornoSiscota = False
    
    If UCase(App.EXEName) Like "D46O003*" Then
        Set lobjstcA28V721B = New stcA28V721B
        Set lobjA28V721A = New clsA28V721A
        Set lobjA28V721A.M46V999 = Me
        
        lobjstcA28V721B.nosso_numero = pNosso_Numero
        If Not lobjA28V721A.gfCarregarColecao(lobjstcA28V721B, "SYASPORTAL.DBO.") Then
            Call gpGraLog(1, lobjA28V721A.UltimoErro)
            Exit Function
        Else
            If lobjA28V721A.Colecao.Count = 1 Then
                Set lobjstcA28V721B = lobjA28V721A.Colecao(1)
                gEhRetornoSiscota = lobjstcA28V721B.STATUS = enum_StatusTransm_EnvioKit Or lobjstcA28V721B.STATUS = enum_StatusTransm_Aprovada
            End If
        End If
    End If

SAIDA:
    Set lobjstcA28V721B = Nothing
    Set lobjA28V721A = Nothing
End Function

Public Function gfTarifaAtual_Impar(ByRef pMensagem As String) As Integer
    Dim lBanco As ADODB.Connection
    Dim lrsBusca As clsYasRecordSet
    Dim lSelect As String
    
    On Error GoTo Erro

    If gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lBanco, pMensagem) = False Then
        GoTo Erro
    End If
    
    lSelect = "SELECT MAX(Cod_Produto) Cod_Produto FROM tab_produto"
    
    If gfObtRegistro(lBanco, lSelect, lrsBusca, pMensagem) = False Then
        GoTo Erro
    End If
    
    If Not lrsBusca.EOF Then
        gfTarifaAtual_Impar = lrsBusca("Cod_Produto")
    End If

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lrsBusca = Nothing
    Return
Erro:
    If Err.Number > 0 Then
        pMensagem = pMensagem & "ERR " & Err.Number & " DESC " & Err.Description
    End If
    GoSub SAIDA
End Function

Public Function Is_BlackList(ByVal pCod_Tarifa As Integer, ByVal pCod_Ramo As e_CodRamo, ByVal pTip_Pessoa As e_TipPessoa, ByVal pCNPJ_CPF As String, _
                             ByRef pMensagem As String) As Boolean
    Dim lBanco As ADODB.Connection
    Dim lrsBusca As clsYasRecordSet
    Dim lSelect As String
    
    On Error GoTo Erro
    
    Is_BlackList = False
    
    If gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lBanco, pMensagem) = False Then
        GoTo Erro
    End If
    
    pCNPJ_CPF = Trim(pCNPJ_CPF)
    If pTip_Pessoa = e_TipPessoa_Fisica Then
        pCNPJ_CPF = Right(String$(11, "0") + pCNPJ_CPF, 11)
    Else
        pCNPJ_CPF = Right(String$(14, "0") + pCNPJ_CPF, 14)
    End If
    
    lSelect = "SELECT NUM_CNPJ_CPF "
    lSelect = lSelect & "FROM TBM_Black_List "
    lSelect = gfWhereAnd(lSelect) & " Cod_Tarif = " & gfFormatarCamposSQL(pCod_Tarifa)
    lSelect = gfWhereAnd(lSelect) & " Cod_Ramo = " & gfFormatarCamposSQL(pCod_Ramo)
    lSelect = gfWhereAnd(lSelect) & " Tip_Pessoa = " & gfFormatarCamposSQL(pTip_Pessoa)
    lSelect = gfWhereAnd(lSelect) & " NUM_CNPJ_CPF = " & gfFormatarCamposSQL(pCNPJ_CPF)
    
    If gfObtRegistro(lBanco, lSelect, lrsBusca, pMensagem) = False Then
        GoTo Erro
    End If
    
    If Not lrsBusca.EOF Then
        Is_BlackList = True
    End If

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lrsBusca = Nothing
    Return
Erro:
    If Err.Number > 0 Then
        pMensagem = pMensagem & "ERR " & Err.Number & " DESC " & Err.Description
    End If
    GoSub SAIDA
End Function

Public Function mfRetorna_ValIS_VerbasSeparadas(pobjStcCobert As stcA46V704B) As Double
    'Criado para somar o valor da cobertura obrigatoria em caso de verbas separadas
    Dim lobjstcPedLoc As stcA46V708B
    Dim lobjstcPedCobert  As stcA46V704B
   
    Set lobjstcPedLoc = pobjStcCobert.ObjPai
    
    mfRetorna_ValIS_VerbasSeparadas = pobjStcCobert.Val_IS
    If lobjstcPedLoc.Ind_VerbaSep And pobjStcCobert.Cod_Grupo_Cobertura = 1 Then
        For Each lobjstcPedCobert In lobjstcPedLoc.ColecaoPedCobert
            If lobjstcPedCobert.Cod_Grupo_Cobertura = 5 Then
                mfRetorna_ValIS_VerbasSeparadas = mfRetorna_ValIS_VerbasSeparadas + lobjstcPedCobert.Val_IS
                Exit For
            End If
        Next
    End If

    Set lobjstcPedLoc = Nothing
    Set lobjstcPedCobert = Nothing

End Function

Public Function ContemResidencialSupremo(Optional ByVal rsP00APOL As clsYasRecordSet, Optional Item As stcA46V708B) As Boolean
    Dim ColObjPedLoc As Collection
    Dim i As Integer
    Dim Coberturas As stcA46V704B
    
    ContemResidencialSupremo = False
    
    If Not rsP00APOL Is Nothing Then
        Do Until rsP00APOL.EOF
            If rsP00APOL!Cod_Ramo <> 113 Then
                Exit Function
            End If
            If rsP00APOL!Val_Vr >= 2000000 Then
                rsP00APOL!Cod_Estip = 30
                rsP00APOL!Tab_Clasf = 30
                Exit Function
            End If
            
            rsP00APOL.MoveNext
        Loop
    End If
    
    If Not Item Is Nothing Then
        For Each Coberturas In Item.ColecaoPedCobert
            If Coberturas.Cod_Ramo <> 113 Then
                Exit Function
            End If
            If ((Coberturas.Val_IS >= 2000000) Or (mfRetorna_ValIS_VerbasSeparadas(Coberturas) >= 2000000)) And Coberturas.Tab_Cobert = 30 Then
                ContemResidencialSupremo = True
                Exit Function
            End If
        Next Coberturas
    End If
    
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set ColObjPedLoc = Nothing
    Set Coberturas = Nothing
    Return
End Function

'' 2017-01-16
Public Function GetBancoByID(ByVal codBanco As Integer, _
                             ByRef pcodigo As Integer, ByRef pNome As String, ByRef pNomeAbreviado As String, ByRef pIsCarne As Boolean, ByRef pIsDebito As Boolean, ByRef pIsRestituicao As Boolean, _
                             ByRef pMensagem As String) As Boolean
    Dim lConn As ADODB.Connection
    Dim lRsQuery As clsYasRecordSet
    Dim lSelect As String
    
    On Error GoTo Erro
    
    GetBancoByID = False
    
    If gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lConn, pMensagem) = False Then
        GoTo Erro
    End If
    
    lSelect = "select top 1 Codigo, Nome, NomeAbreviado, Carne, Debito, Restituicao from Tbm_Bancos (nolock) where Codigo=" & gfFormatarCamposSQL(codBanco)
    
    If gfObtRegistro(lConn, lSelect, lRsQuery, pMensagem) = False Then
        GoTo Erro
    End If
    
    If lRsQuery.RecordCount <= 0 Then
        pMensagem = pMensagem + "0 registros. Banco nao encontrado."
    End If
    
    lRsQuery.MoveFirst
    
    While Not lRsQuery.EOF
        GetBancoByID = True
        
        pcodigo = lRsQuery!codigo
        pNome = lRsQuery!Nome
        pNomeAbreviado = lRsQuery!nomeAbreviado
        pIsCarne = lRsQuery!Carne
        pIsDebito = lRsQuery!Debito
        pIsRestituicao = lRsQuery!Restituicao
        
        lRsQuery.MoveNext
    Wend

Fim:
    Set lRsQuery = Nothing
    Call Me.gpFecha2(lConn)
    Exit Function
    
Erro:
    If Err.Number > 0 Then
        pMensagem = pMensagem & "ERR " & Err.Number & " DESC " & Err.Description
    End If
    GoSub Fim
End Function

Public Function LiberaPlanoResidencialSupremo(ByVal CodPlano As Integer, ByVal TipoEmissao As Integer) As Boolean
    
    If gResidencialSupremoTodos = "true" Then
        LiberaPlanoResidencialSupremo = True
    ElseIf gResidencialSupremo = "true" Then
        If gCodProduto >= 1140 Then
            LiberaPlanoResidencialSupremo = PermitirResidencialSupremo(gCod_Ramo, CodPlano, gCod_Corretor, gCodProduto, TipoEmissao, "")
        End If
    ElseIf gRelPilotoSupremo <> "" Then
    Dim RelLiberaSupremo() As String
    Dim lInd As Integer
    
    RelLiberaSupremo = Split(gRelPilotoSupremo, ";")
    LiberaPlanoResidencialSupremo = False
    
    If gCodProduto >= 1140 Then
        If UBound(RelLiberaSupremo()) = -1 Then
            LiberaPlanoResidencialSupremo = True
        End If
    
        For lInd = 0 To UBound(RelLiberaSupremo())
            If RelLiberaSupremo(lInd) <> "" Then
                If UCase(RelLiberaSupremo(lInd)) = UCase(gCod_Corretor) Then
                    LiberaPlanoResidencialSupremo = True
                    Exit For
                End If
            End If
        Next
    End If
    Else
        LiberaPlanoResidencialSupremo = False
    End If

SAIDA:
    Exit Function
End Function

Public Function PermitirResidencialSupremo(ByVal codRamo As Integer, _
                                           ByVal CodPlano As Integer, _
                                           ByVal codCorretor As Long, _
                                           ByVal CodTarifa As Integer, _
                                           ByVal TipoEmissao As Integer, _
                                           ByRef pMensagem As String) As Boolean
                                                                                     
    Dim lConn As ADODB.Connection
    Dim lRsQuery As clsYasRecordSet
    Dim lSelect As String
    
    On Error GoTo Erro
    
    PermitirResidencialSupremo = False
    
    If gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lConn, pMensagem) = False Then
        GoTo Erro
    End If
    
    lSelect = "EXEC proc_PermitirCorretorSupremo " & gfFormatarCamposSQL(codRamo) & "," & gfFormatarCamposSQL(CodPlano) & "," & gfFormatarCamposSQL(codCorretor) & "," & gfFormatarCamposSQL(CodTarifa) & "," & gfFormatarCamposSQL(TipoEmissao)
    
    If gfObtRegistro(lConn, lSelect, lRsQuery, pMensagem) = False Then
        GoTo Erro
    End If
    
    If lRsQuery.RecordCount <= 0 Then
        pMensagem = pMensagem + "Nenhum registro encontrado."
    Else
        lRsQuery.MoveFirst
        
        While Not lRsQuery.EOF
            If lRsQuery!PERMITIR = 1 Then
                PermitirResidencialSupremo = True
            End If
            lRsQuery.MoveNext
        Wend
    End If

Fim:
    Set lRsQuery = Nothing
    Call Me.gpFecha2(lConn)
    Exit Function
    
Erro:
    If Err.Number > 0 Then
        pMensagem = pMensagem & "ERR " & Err.Number & " DESC " & Err.Description
    End If
    GoSub Fim
End Function

Public Function gfRetornaBoleto(ByVal pBoleto As String, _
                                        ByVal pCodRamo As Integer, _
                                        ByVal pCodProduto As Integer, _
                                        ByVal pCodCorr As Double, _
                                        ByRef pMensagem As String) As Boolean
'Fun��o: Verifica se o boleto existe na tab_ctrl_pagantec

'Par�metros de entrada...pCodigoCorretor...C�digo do corretor.
'Par�metro de sa�da......pLimite......

    Dim bdDiversos As ADODB.Connection
    Dim lI As Double  'Contador.
    Dim lSelect As String       'Comando SELECT.
    Dim rsPagantec As clsYasRecordSet    'Tabela: tab_ctrl_pagantec.
    Dim rsConsist As clsYasRecordSet    'Tabela: tab_parm_consist.
    Dim lRamoSusep As Integer
    Dim lRamoBol As Integer
    Dim lCodCorrBol As Double
    Dim lDataParam As String

    On Error GoTo ControleErro

    gfRetornaBoleto = False

    If gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, bdDiversos, pMensagem) = False Then
        GoTo ControleErro
    End If
    
    lSelect = "SELECT Cod_Tipo " & _
              "FROM   GED..Tab_Parm_Consist With (nolock) " & _
              "WHERE  Cod_Tipo = 1 and Cod_Ramo = " & pCodRamo & " " & _
              "AND    Cod_Produto = " & pCodProduto
              
    If gfObtRegistro(bdDiversos, lSelect, rsConsist, pMensagem) = False Then
        GoTo ControleErro
    End If
    If rsConsist.RecordCount > 0 Then
        gfRetornaBoleto = True
        GoTo Fim
    End If
    
    lDataParam = gfForData(1, DateAdd("YYYY", -1, Date))
    
    '1. Busca informa��o.
    lSelect = "SELECT COD_RAMO, " & _
              "NOM_SEGURADO, " & _
              "COD_CORR, " & _
              "VAL_PR " & _
              "FROM   SYAS_EMIS..TAB_CTRL_PAGANTEC WITH (NOLOCK) " & _
              "WHERE  NUM_BOLETO = " & pBoleto & " " & _
              "AND DAT_DIGIT > " & lDataParam & " " & _
              "ORDER BY DAT_DIGIT ASC"

    If gfObtRegistro(bdDiversos, lSelect, rsPagantec, pMensagem) = False Then
        GoTo ControleErro
    End If
    
    If rsPagantec.EOF Then
        pMensagem = "Boleto n�o localizado."
        GoTo Fim
    End If

    Do While Not rsPagantec.EOF
        lRamoBol = rsPagantec!Cod_Ramo
        lCodCorrBol = rsPagantec!Cod_Corr
        
        Select Case lRamoBol
            Case e_CodRamoSusep_Empresarial
                lRamoBol = 112
            Case e_CodRamoSusep_Residencial
                lRamoBol = 113
            Case e_CodRamoSusep_Condominio
                lRamoBol = 114
        End Select
            
        If lRamoBol = pCodRamo And (lCodCorrBol = 0 Or lCodCorrBol = pCodCorr) Then
            gfRetornaBoleto = True
            pMensagem = ""
            GoTo Fim
        Else
            If lCodCorrBol = 0 Then
                lCodCorrBol = pCodCorr
            End If
            
            pMensagem = "Boleto com Ramo: " & lRamoBol & vbCrLf & _
            "Corretor: " & lCodCorrBol & vbCrLf & _
            "Valor: R$: " & rsPagantec!VAL_PR & vbCrLf & _
            "Divergente da proposta."
        End If
        rsPagantec.MoveNext
    Loop

Fim:
    Set rsConsist = Nothing
    Set rsPagantec = Nothing
    Call Me.gpFecha2(bdDiversos)
    Exit Function
    
ControleErro:
    pMensagem = "gfRetornaBoleto ERRO - " & Err & " " & Err.Description
    Call gpGraLog(0, pMensagem)
End Function

Public Function gfVersao() As String
    gfVersao = App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Function GetIndSinistroSimNao(ByVal NUM_PI As Double, ByVal Num_Apol As Double, ByVal Num_Item As Integer, _
                                     ByRef IndSinistro As Integer, ByRef IndSinistroOrig As Integer) As Boolean
    Dim lConn As ADODB.Connection
    Dim lRsQuery As clsYasRecordSet
    Dim lSelect As String
    Dim pMensagem As String
    
    GetIndSinistroSimNao = False
    IndSinistro = 0
    IndSinistroOrig = 0
    
    On Error GoTo Erro
    
    If gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lConn, pMensagem) = False Then
        GoTo Erro
    End If
    
    lSelect = "exec Ged..proc_Log_SinEmissao_Consultar " & gfFormatarCamposSQL(NUM_PI) & ", " & gfFormatarCamposSQL(Num_Apol) & ", " & gfFormatarCamposSQL(Num_Item)
    
    If gfObtRegistro(lConn, lSelect, lRsQuery, pMensagem) = False Then
        GoTo Erro
    End If
    
    If lRsQuery.RecordCount <= 0 Then
        pMensagem = pMensagem + "0 registros. Banco nao encontrado."
        GoTo Erro
    End If
    
    lRsQuery.MoveFirst
    
    While Not lRsQuery.EOF
        IndSinistro = lRsQuery!Ind_Sinistro
        IndSinistroOrig = lRsQuery!Ind_Sinistro_ORIG
        
        lRsQuery.MoveNext
    Wend

    GetIndSinistroSimNao = True
Fim:
    Set lRsQuery = Nothing
    Call Me.gpFecha2(lConn)
    Exit Function
    
Erro:
    If Err.Number > 0 Then
        pMensagem = pMensagem & "ERR " & Err.Number & " DESC " & Err.Description
    End If
    GoSub Fim
End Function

Public Function SetIndSinistroSimNao(ByVal NUM_PI As Double, ByVal Num_Apol As Double, ByVal Num_Item As Integer, ByVal IndSinistro As Integer, ByVal Cod_User As String) As Boolean
    Dim lConn As ADODB.Connection
    Dim lRsQuery As clsYasRecordSet
    Dim lQuery As String
    Dim pMensagem As String
    
    SetIndSinistroSimNao = False
    
    On Error GoTo Erro
    
    If gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lConn, pMensagem) = False Then
        GoTo Erro
    End If
    
    lQuery = "exec Ged..proc_Log_SinEmissao_Inserir " & gfFormatarCamposSQL(NUM_PI) & ", " & gfFormatarCamposSQL(Num_Apol) & ", " & gfFormatarCamposSQL(Num_Item) & ", " & gfFormatarCamposSQL(IndSinistro) & ", " & gfFormatarCamposSQL(Cod_User)
    
    If gfExeSQL(lConn, lQuery, pMensagem) = False Then
        GoTo Erro
    End If
    
    SetIndSinistroSimNao = True
Fim:
    Set lRsQuery = Nothing
    Call Me.gpFecha2(lConn)
    Exit Function
    
Erro:
    If Err.Number > 0 Then
        pMensagem = pMensagem & "ERR " & Err.Number & " DESC " & Err.Description
    End If
    GoSub Fim
End Function


Public Function CheckSiscotaRA(ByVal NossoNumero As String) As Integer
    Dim lConn As ADODB.Connection
    Dim lRsQuery As clsYasRecordSet
    Dim lQuery As String
    Dim pMensagem As String
    Dim qtdReg As Long
    
    On Error GoTo Erro
    
    If gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lConn, pMensagem) = False Then
        GoTo Erro
    End If
    
    lQuery = "update RamosDiversos..Tab_ped_Siscota set status=2 where nosso_numero = '" & NossoNumero & "' and status not in (2,3)"
    Call gfExeSQL(lConn, lQuery, pMensagem, qtdReg)
    
    If pMensagem <> "" Then GoTo Erro

    
    lQuery = "Select top 10 * from RamosDiversos..Tab_ped_Siscota where nosso_numero = '" & NossoNumero & "'"

    If gfObtRegistro(lConn, lQuery, lRsQuery, pMensagem) = False Then
        GoTo Erro
    End If

    If lRsQuery.RecordCount <= 0 Then
        pMensagem = pMensagem + "0 registros, erro ao acessar o banco."
        GoTo Erro
    End If

    qtdReg = lRsQuery.RecordCount

Fim:
    CheckSiscotaRA = qtdReg
    
    Set lRsQuery = Nothing
    Call Me.gpFecha2(lConn)
    Exit Function
    
Erro:
    If Err.Number > 0 Then
        pMensagem = pMensagem & "ERR " & Err.Number & " DESC " & Err.Description
    End If
    
    qtdReg = -1
    
    GoSub Fim
End Function
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

'Repique OnLine
Public Function GetRepique(ByVal NossoNumero As String) As Boolean
    Dim lConn As ADODB.Connection
    Dim lRsQuery As clsYasRecordSet
    Dim lSelect As String
    Dim pMensagem As String
    
    GetRepique = False
    
    On Error GoTo Erro
    
    If gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lConn, pMensagem) = False Then
        GoTo Erro
    End If
    
    lSelect = "select top 1 Id_Repique, Dat_Repique, Cod_User, Nosso_Numero, Status from Tab_ControleRepique (nolock) where Nosso_Numero = '" & NossoNumero & "'  Order by Id_Repique desc"
    
    If gfObtRegistro(lConn, lSelect, lRsQuery, pMensagem) = False Then
        GoTo Erro
    End If
    
    If lRsQuery.RecordCount <= 0 Then
        GoTo Erro
    End If
    
    lRsQuery.MoveFirst
    
    While Not lRsQuery.EOF
        If (lRsQuery!STATUS) = 1 Then
            GetRepique = True
        End If
        lRsQuery.MoveNext
    Wend

Fim:
    Set lRsQuery = Nothing
    Call Me.gpFecha2(lConn)
    Exit Function
    
Erro:
    If Err.Number > 0 Then
        pMensagem = pMensagem & "ERR " & Err.Number & " DESC " & Err.Description
    End If
    GoSub Fim
    
End Function

'******************************************************************************************************************
'#20180813 - Jira: 25946
'Metodo: expObterDados
'Finalidade: Encapsular Consulta
'******************************************************************************************************************
Public Function expObterDados(BancoDados As String, Consulta As String, ByRef prst As clsYasRecordSet, ByRef pMensagem As String)

    Dim lBDSiscota As ADODB.Connection
    Dim IsPortabilidade As Boolean
    Dim pAlerta As String
     
    'TODO: Verificar VB6 com form
    If M46V111.M46V999.gfAbrBasDados(BancoDados, M46V111.M46V999.gTipGed, M46V111.M46V999.gEndGed, lBDSiscota, pMensagem) = False Then
        Call M46V111.M46V999.gpGraLog(2, pMensagem)
        Exit Function
    End If
            
   If M46V111.M46V999.gfObtRegistro(lBDSiscota, Consulta, prst, pMensagem) = False Then
        Call M46V111.M46V999.gpFecha2(lBDSiscota)
        Call M46V111.M46V999.gpGraLog(2, pMensagem)
        Exit Function
    End If
    'If prst.EOF Then
    '    Call gM46V999.gpGraLog(1, "AGE0023 - SISCOTA n�o informado ou inv�lido.")
    '    Call gM46V999.gpFecha2(lBDSiscota)
    '    Exit Function
    'End If
    Call M46V111.M46V999.gpFecha2(lBDSiscota)
End Function

'******************************************************************************************************
'#20180813 - Jira: 25946
'M�todo: expObterDados
'Finalidade: Encapsular Comando
'******************************************************************************************************

Public Function expExeSQL(BancoDados As String, Comando As String, ByRef pMensagem As String)

    Dim lBDSiscota As ADODB.Connection
    Dim IsPortabilidade As Boolean
    Dim pAlerta As String

 
    If M99V001.M46V999.gfAbrBasDados(BancoDados, M99V001.M46V999.gTipGed, M99V001.M46V999.gEndGed, lBDSiscota, pMensagem) = False Then
        Call M99V001.M46V999.gpGraLog(2, pMensagem)
        Exit Function
    End If
            
    If M99V001.M46V999.gfExeSQL(lBDSiscota, Comando, pMensagem) <> 0 Then
        Call M99V001.M46V999.gpGraLog(2, pMensagem)
        Exit Function
    End If
    Call M99V001.M46V999.gpFecha2(lBDSiscota)
End Function

