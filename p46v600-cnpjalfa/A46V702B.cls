VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V702B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================'
'                                                                                                                                                       '
' Class Name: stcA46V702B --> Tabela Tab_Ped                                                                                                            '
' Data/Hora: 31/08/2005 - 11:51:10                                                                                                                      '
' Sistema: Multi                                                                                                                                        '
' Thais Juren Gomes                                                                                                                           '
'                                                                                                                                                       '
' ======================================================================================================================================================'

Public nosso_numero As String       'Nosso N�mero.
Public Cod_Ramo As e_CodRamo    'C�digo do Ramo.
Public Cod_Org As Integer    'C�digo do Org�o.
Public Num_Ped As Long    'N�mero do pedido.
Public Num_Apol As Double   'Sem uso - N�mero da ap�lice.
Public Num_Endo As Long     'Sem uso - N�mero do endosso.
Public Num_Apol_An As Double      'N�mero da ap�lice anterior.
Public Cod_Reg As Integer    'C�digo do Registro.
Public Cod_Bco_Cobr As e_Banco      'Banco de cobran�a. (d�bito ou carn�)
Public Cod_Agen_Angar As Integer
Public Cod_Agen_Indicr As Integer
Public Num_Cta_Indicr As Long
Public Cod_Unid_Prod As Integer       'C�digo da Unidade.
Public Cod_Prod As Integer  'C�digo do produtor.
Public Cod_Agen_Bas_Cor As Integer          'Ag�ncia de d�bito em conta corrente.
Public Cod_Dig_Agen_Bas_Cor As Integer              'D�gito da ag�ncia.
Public Num_Cta_Bas_Cor As Long            'N�mero da conta para d�bito em conta corrente.
Public mNum_Cta_Bas_Cor_Compl As Long                   'N�mero da conta para d�bito em conta corrente.
Public Cod_Segurado As Long         'C�digo do segurado.
Public Dig_Cod_Segurado As String          'D�gito do C�digo do segurado. [FRH.08]
Public Nom_Segurado As String       'Nome do segurado.
Public Dsc_Compl_Segurado As String             'Complemento do nome.
Public Num_Cgc_Cpf As String      'CPF-CNPJ do segurado.
Public Cod_Tip_Logr As String       'Tipo do logradouro.
Public Nom_Logr As String   'Nome do logradouro.
Public Dsc_Compl_Logr As String         'Complemento do logradouro.
Public Nom_Bairro As String     'Nome do bairro.
Public Nom_Cid As String  'Cidade.
Public Sig_UF As String    'UF.
Public Num_CEP As Long    'N�mero do CEP.
Public Num_Fone As Long     'Telefone.
Public Cod_Estip As Double    'C�digo do estipulante.
Public Val_Coef_Cms_Pad As Double           'Coeficiente da comiss�o padr�o.
Public Dat_Ini_Vig As Long        'In�cio da vig�ncia.
Public Dat_Ter_Vig As Long        'T�rmino da vig�ncia.
Public Dat_Emis As Long     'Data de emiss�o.
Public Cod_Moed As Integer  'C�digo da moeda.
Public Per_Juro_Frac As Double        'Percentual de juros.
Private Qtd_Parc_ As Integer  'Quantidade de parcelas.
Public Qtd_Anexo As Integer   'Quantidade de anexos.
Public Cod_Iden_Praz As e_TipoPrazo   'Tipo de prazo.
Public Cod_Pr_Min_Endo As Byte
Public Cod_Pr_Min_Parc As Byte
Public Val_Pr_Antec As Double
Public Per_Adic_Form_Pag As Double
Public Per_Desc_Form_Pag As Double
Public Val_IS As Double    'Valor de IS.
Public Val_IS_2 As Double   'Valor de Risco.
Public VAL_PR As Double    'Pr�mio
Public VAL_PR_Informado As Double    'Pr�mio Informado
Public Val_Pr_Ref_Tarif As Double
Public Val_Pr_Ref_Resg As Double
Public Dat_Ini_Viag As Long
Public Cod_Tip_Liber As Byte
Public Cod_Iden_Alt_Corr As String
Public Cod_User_Digit As String         'Usu�rio de digita��o.
Public DAT_DIGIT As Long      'Data da digita��o.
Public HOR_DIGIT As Integer   'Hora da digita��o.
Public Cod_User_Liber As String         'Usu�rio de libera��o.
Public Dat_Liber As Long      'Data da libera��o.
Public Hor_Liber As Integer   'Hora da libera��o.
Public Cod_Sit_Ped As e_SitPedido    'Situa��o do pedido.
Public Cod_Tip_Doc As e_TipDoc    'Tipo de documento.
Public Cod_Dizer As Integer   'C�digo do dizer.
Public Qtd_Copia As Integer   'Quantidade de c�pias.
Public Val_Cust_Emis As Double        'Valor do custo de emiss�o.
Public Val_Parc_Min As Double       'Valor da parcela m�nima
Public Val_Juro As Double   'Juros.
Public Val_IOF As Double  'IOF
Public Num_Prop As Long
Public Cod_Posi_Atual As String
Public Dat_Mov_Ini As Long
Public Dat_Mov_Ter As Long
Public Tip_Cust_Emis As Byte
Public Val_Cambio_Emis As Double
Public Iden As String
'Public NUM_PI As Long
Public NUM_PI As Double
Public NossoNumeroIn As String
Public Cod_Produto As Integer     'C�digo do produto.
Public Cod_Protocolo As Integer
Public Cod_Sequencia As Integer
Public Tip_Pessoa As e_TipPessoa    '1 - Jur�dica 2 - F�sica
Public Cod_Ramo_Ativ As Integer       'C�digo do ramo de atividade.
Public Cod_Doc As Integer    'C�digo do tipo de documento.
Public Num_Doc As String  'N�mero do documento.
Public Org_Exp As String  '�rg�o expedidor.
Public Dat_Emis_Doc As Long         'Data da emiss�o do documento (yyyymmdd).
Public Nom_Empresa As String      'Nome da empresa onde trabalha.
Public Sexo As String    'Sexo do segurado: M -masculino , F - Feminino.
Public Dat_Pri_Calculo As Long            'Data do primeiro c�lculo.
Public Ult_Tarifa As Integer    '�ltima tarifa v�lida.
Public Cod_Unicidade As Long          'C�digo de unicidade.
Public Cod_Ver_Unici As Integer       'C�digo da Vers�o da unicidade.
Public Val_PriParcela As Double
Public val_DemParcelas As Double
Public Desc_Plano As String     'Descri��o do plano.
Public Val_RendaFatu As Double
Public LMI_Unico As Integer   'Se � LMI = 1 = Sim / 2 = N�o
Public Per_Desc_LMI As Double
Public Num_Cotac_LMI As String
Public Tip_Transmissao As eTipTransmissao
Public Ind_CondGerais As Integer
Public mAuxPriVenc As String
Public mDigitApolice As Long
'Public Num_PI_Emissao As Long
Public Num_PI_Emissao As Double
Public Sit_Emissao As Long
Public Ind_Reentrada As Integer
Public Dig_PI_Emissao As String
Public Nosso_Numero_Original As String
Public Cod_Segmentacao As Integer
Public Dat_Transmissao As String

'Public Num_PI_Auto_Progresso As Long
Public Num_PI_Auto_Progresso As Double
Public Per_Juro_Progresso As Double
Public Qtd_Parc_Progresso As Integer
Public Cod_Tarif_Auto As Long

Public Ind_Oculta_Web As Integer    '' 20170524: Indica se oculta plano na web
' =====================================================================================================================================================

'Corretor
Public Ind_CotacaoSolic As Boolean
Public SenhaCotacao As String
Public StatusCotSiscota As enum_StatusTransm
Public SiscMaxDesc As Double

'Interno
Public mStatusDivSiscota As enum_StatusTransm
Public mCotacaoSiscota As String
Public mTransmissao As Boolean
Public mValPrTransmitido As Double
Public mValPrAntecTransmitido As Double
Public mCotacaoInterna As Boolean    '1000913 - Permitir alterar dados cotacao interna

Public Enum eTipTransmissao     'Tipo de Transmissao
    eTipTransmissao_ND = 0
    eTipTransmissao_Corretor = 1
    eTipTransmissao_Automatic = 2
End Enum

Public Enum e_TipDoc            'Enum de tipo de documento.
    e_TipDoc_Cotacao = 1
    e_TipDoc_Proposta = 2
End Enum

Public Enum e_SitPedido         'Enum de situa��o do pedido.
    e_SitPedido_SemCalculo = 1
    e_SitPedido_ComCalculo = 2
    e_SitPedido_Transmitido = 3
    e_SitPedido_Nenhum = -1
End Enum

Public Enum e_SitTransmissao
    e_SitTransmissao_Nenhum = 0
    e_SitTransmissao_Transmitido = 1
    e_SitTransmissao_Cancelado = 2
    e_SitTransmissao_Emitido = 3
    e_SitTransmissao_Recusado = 4
End Enum

Public Enum e_TipoPrazo         'Enum de tipo de prazo.
    e_TipoPrazo_ND = 0
    e_TipoPrazo_ProRata = 1
    e_TipoPrazo_PrazoCurto = 2
    e_TipoPrazo_PrazoLongo = 3
End Enum

Public Enum e_CodRamo           'Enum de c�digo do ramo.
    e_CodRamo_NaoDefinido = 0
    e_CodRamo_Empresarial = 112
    e_CodRamo_Residencial = 113
    e_CodRamo_Condominio = 114
    e_CodRamo_Incendio = 110
    e_CodRamo_Roubos = 150
    e_CodRamo_Tumultos = 170
    e_CodRamo_Transp_internacional = 211
    e_CodRamo_Transp_nacional = 212
    e_CodRamo_Automoveis = 310
    e_CodRamo_Cascos = 330
    e_CodRamo_Riscos_de_Petroleo = 340
    e_CodRamo_Aeronauticos = 350
    e_CodRamo_LucrosCessantes = 410
    e_CodRamo_Fianca_Locataria = 460
    e_CodRamo_Resp_Civil = 510
    e_CodRamo_Rc_Transportador = 540
    e_CodRamo_Rcf_dc = 550
    e_CodRamo_Rc_Armador = 560
    e_CodRamo_D_P_E_M = 570
    e_CodRamo_Penhor_rural = 620
    e_CodRamo_Riscos_de_engenharia = 670
    e_CodRamo_Riscos_diversos = 710
    e_CodRamo_Garantia_obrigacoes = 750
    e_CodRamo_Acidentes_pessoais = 810
    e_CodRamo_Ap_bilhete = 820
    e_CodRamo_GarantiaEstendida = 950
    e_CodRamo_Vida_em_grupo = 930
    e_CodRamo_Vida_prestamista = 932
    e_CodRamo_Vida_individual = 910
End Enum
Public Enum enum_StatusTransm
    enum_StatusTransm_SolRecotacao = -1
    enum_StatusTransm_Inicio = 0
    enum_StatusTransm_Solicitada = 1
    enum_StatusTransm_Aprovada = 2
    enum_StatusTransm_Declinada = 3
    enum_StatusTransm_Cancelada = 4
    enum_StatusTransm_EnvioKit = 5
End Enum

Public Enum e_TipPessoa         'Enum de tipo de pessoa.
    e_TipPessoa_Juridica = 1
    e_TipPessoa_Fisica = 2
End Enum
Public ColecaoPedLoc As Collection    'Cole��o de itens.
Public ColecaoPedCorr As Collection     'Cole��o de corretores.
Public ColecaoMovAgen As Collection     'Cole��o de TAB AGEN.
Public ColecaoPedAgregados As Collection     'Cole��o de Ped Agregados.
Public ColecaoPedCoss As Collection     'Cole��o de cosseguros.
Public ColecaoPedJuros As Collection      'Cole��o de juros.
Public ColecaoPedVenc As Collection     'Cole��o de datas de vencimento.
Public ColecaoParcelamento As Collection          'Cole��o de Parcelamento
Public objstcPedEndereco As stcA46V707B       'Objeto de endere�o.
Public objstcPedAuto As stcA46V719B   'Objeto de endere�o.
Public objstcCtrlEmis As stcA46V712B    'Objeto de controle de emiss�o.
Public objstcMovAgen As stcA46V743B    'Objeto de movimento de ag�ncia.
Public objstcPedAgregados As stcA46v741B    'Objeto de movimento de Ped Agregados.
Public objstcEspec As stcA46V721B    'Objeto de Especifica��o
Public ColecaoTextoEndo As Collection       'Objeto de Texto para Endosso.
Public ColecaoTextoEspec As Collection        'Objeto de Texto para Especifica��o Padr�o.
Public objstcP0045400 As stcA28V721B    'Solicita��o de cota��o.
Public objstcLiberacaoSerasa As stcA46V746B
Public objstcBoleto As stcA28V720B  'Boleto.
Public ColecaoLMI As Collection    'Capa - LMI
Private mCod_Unidade As Long          'C�digo da unidade.
Private mNom_Unidade As String        'Nome da unidade.
Private mCod_Produtor As Long           'C�digo do produtor.
Private mNom_Produtor As String         'Nome do produtor.
Private mNumSusep As String

Public Enum e_IdeEmitida
    e_IdeEmitida_Nao = 0
    e_IdeEmitida_Emitir = 1
    e_IdeEmitida_Suspender = 2
    e_IdeEmitida_EmitirProd = 3
    e_IdeEmitida_SuspenderProd = 4
    e_IdeEmitida_Reemissao = 5
    e_IdeEmitida_ProcessadoMainframe = 10
End Enum

Public Enum e_FormaPagamento
    e_FormaPagamento_ND = 0
    e_FormaPagamento_Carne = 1
    e_FormaPagamento_Debito = 2
    e_FormaPagamento_Credito = 3
End Enum

Private lTransmitidaeEmitida As Integer

'1101015 - Projeto 11010 - Renova��o Expressa
Private lApol_Ant_QTD_PARC As Integer
Private lApol_Ant_PER_JURO_FRAC As Long
Private lApol_Ant_PER_CORR As Long

Public Enum e_Banco
    e_Banco_ND = 0
    e_Banco_DoBrasil = 1
    e_Banco_Bradesco = 237
    e_Banco_Itau = 341
    e_Banco_Sudameris = 347
    e_Banco_Unibanco = 409
    e_Banco_Boston = 479
    e_Banco_Real = 356
    e_Banco_HSBC = 399
    e_Banco_Santander = 33
    e_Banco_Banese = 47
    e_Banco_Safra = 422
    e_Banco_Bancoob = 756
End Enum
Private mStcPedidoApolice As stcA46V702B    '***********NUNCA MUDAR PARA PUBLIC!!!
Private mClsM46V999 As clsM46V999

Public ColecaoProlabore As Collection      'Cole��o de Prolabore
Public ColecaoAgenciador As Collection       'Cole��o de Agenciamento.
Public ColecaoTelefones As Collection
Public ColecaoServicoAss24h As Collection

Public CalBatente As Boolean     'Define se o calculo de batente vai ser realizado

Private gDescrBancoRestituicao As String
Private gDescrBanco As String
Public TransactionId As String     'Define se o calculo de batente vai ser realizado


' ======================================================================================================================================================

Public Property Get Class_Name() As String
    Class_Name = "A46V702B"
End Property

Public Property Get Class_ID() As String
    If Tip_Transmissao = eTipTransmissao_ND Then
        Class_ID = CStr(nosso_numero)
    Else
        'Variavel Tip_Transmissao.. preenchida somente quando esta trabalhando com as tabelas "FTP_"
        Class_ID = CStr(nosso_numero) & CStr(Cod_Unicidade)
    End If
End Property
Public Property Get Class_ID_PI() As String
    Class_ID_PI = Format(NUM_PI, "0000000000")
End Property
Public Property Get gTransmitidaeEmitida() As Integer
    Dim Aux As Integer
    If lTransmitidaeEmitida > 0 Then
        gTransmitidaeEmitida = lTransmitidaeEmitida
    Else
        Aux = gCod_Acao
        gTransmitidaeEmitida = lTransmitidaeEmitida
    End If
End Property

Private Sub Class_Initialize()
    mNumSusep = ""
    Set ColecaoPedLoc = New Collection
    Set ColecaoPedCorr = New Collection
    Set ColecaoProlabore = New Collection
    Set ColecaoAgenciador = New Collection
    Set ColecaoTelefones = New Collection
    Set ColecaoPedCoss = New Collection
    Set ColecaoPedJuros = New Collection
    Set ColecaoPedVenc = New Collection
    Set objstcPedEndereco = New stcA46V707B 'Nothing
    Set objstcPedAuto = New stcA46V719B
    Set objstcCtrlEmis = New stcA46V712B
    Set objstcEspec = New stcA46V721B
    Set ColecaoTextoEndo = New Collection
    Set ColecaoTextoEspec = New Collection
    Set ColecaoParcelamento = New Collection
    Set ColecaoLMI = New Collection
    Set ColecaoServicoAss24h = New Collection
    Set mStcPedidoApolice = Nothing
    Set mClsM46V999 = Nothing
End Sub


Private Sub Class_Terminate()
    Set ColecaoPedLoc = Nothing
    Set ColecaoPedCorr = Nothing
    Set ColecaoProlabore = Nothing
    Set ColecaoAgenciador = Nothing
    Set ColecaoTelefones = Nothing
    Set ColecaoPedCoss = Nothing
    Set ColecaoPedJuros = Nothing
    Set ColecaoPedVenc = Nothing
    Set objstcPedEndereco = Nothing
    Set objstcPedAuto = Nothing
    Set objstcCtrlEmis = Nothing
    Set objstcEspec = Nothing
    Set ColecaoTextoEndo = Nothing
    Set ColecaoTextoEspec = Nothing
    Set ColecaoParcelamento = Nothing

    Set objstcP0045400 = Nothing
    Set objstcLiberacaoSerasa = Nothing
    Set objstcBoleto = Nothing
    Set mStcPedidoApolice = Nothing
    Set mClsM46V999 = Nothing
    Set ColecaoLMI = Nothing
    Set ColecaoServicoAss24h = Nothing
End Sub

Public Property Get gNom_UnidadeProd() As String
    Dim lMensagem As String

    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        If Me.Cod_Unid_Prod > 1000 Then
            If Me.Cod_Unid_Prod <> mCod_Unidade Or Trim(mNom_Unidade) = "" Then
                mCod_Unidade = Me.Cod_Unid_Prod
                If M46V999.gfNomUniProdutiva(Me.Cod_Unid_Prod, mNom_Unidade, lMensagem) = False Then
                    Call M46V999.gpGraLog(1, lMensagem)
                    Exit Property
                End If
            End If
        Else
            If Me.Cod_Unid_Prod > 0 Then
                mNom_Unidade = "MATRIZ"
            Else
                mNom_Unidade = ""
            End If
        End If
    Else
        If Me.Cod_Unid_Prod <> mCod_Unidade Or Trim(mNom_Unidade) = "" Then
            mCod_Unidade = Me.Cod_Unid_Prod
            If M46V999.gfNomUniProdutiva(Me.Cod_Unid_Prod, mNom_Unidade, lMensagem) = False Then
                Call M46V999.gpGraLog(1, lMensagem)
                Exit Property
            End If
        End If
    End If
    gNom_UnidadeProd = mNom_Unidade

End Property

Public Property Get gNom_Produtor() As String
    Dim lMensagem As String

    If Me.Cod_Prod <> mCod_Produtor Or Trim(mNom_Produtor) = "" Then
        mCod_Produtor = Me.Cod_Prod
        If M46V999.gfNomProdutor(mCod_Produtor, mNom_Produtor, lMensagem) = False Then
            Call M46V999.gpGraLog(1, lMensagem)
            Exit Property
        End If
    End If
    gNom_Produtor = mNom_Produtor

End Property


Public Property Get gCNPJ_CPF_Formatado() As String
    Dim lCgcCpf As String   'CGC/CPF sem formatação.

    lCgcCpf = Trim(Me.Num_Cgc_Cpf)
    If M46V999.gfPreenchido(lCgcCpf) And lCgcCpf <> "0" Then
        Select Case Len(lCgcCpf)
        Case Is > 11
            If IsNumeric(lCgcCpf) Then
                Call M46V999.gpForCNPJ(1, Format$(lCgcCpf, "00000000000000"), gCNPJ_CPF_Formatado)
            Else
                Call M46V999.gpForCNPJ(1, Right$(String(14, "0") & lCgcCpf, 14), gCNPJ_CPF_Formatado)
            End If
        Case Else
            If Me.Tip_Pessoa = e_TipPessoa_Juridica Then
                If IsNumeric(lCgcCpf) Then
                    Call M46V999.gpForCNPJ(1, Format$(lCgcCpf, "00000000000000"), gCNPJ_CPF_Formatado)
                Else
                    Call M46V999.gpForCNPJ(1, Right$(String(14, "0") & lCgcCpf, 14), gCNPJ_CPF_Formatado)
                End If
            Else
                Call M46V999.gpForCPF(1, Format$(lCgcCpf, "00000000000"), gCNPJ_CPF_Formatado)
            End If
        End Select
    Else
        gCNPJ_CPF_Formatado = " "
    End If
End Property

Public Property Get gDesc_Ramo() As String
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstTbm_Ramo As clsYasRecordSet
    Dim lSelect As String
    Dim lMensagem As String

    'Rotina Alterada para atender a nova Precifica��o
    gDesc_Ramo = "N�o definido"

    If Cod_Ramo = e_CodRamo_NaoDefinido Then
        GoTo SAIDA
    End If

    '1. Abre arquivo.
    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
        Call M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    '3. Obt�m c�digo da Ramo da Tbm_Ramo
    lSelect = "Select Dsc_Ramo from Tbm_Ramo where Cod_Ramo = " & Cod_Ramo

    If M46V999.gfObtRegistro(BDP00Tarifa, lSelect, rstTbm_Ramo, lMensagem) = False Then
        Call M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    If Not rstTbm_Ramo.EOF Then
        gDesc_Ramo = rstTbm_Ramo("Dsc_Ramo")
    End If

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set BDP00Tarifa = Nothing
    Set rstTbm_Ramo = Nothing
    Return

End Property


Public Property Get gEnd_Cob_Diferente() As Boolean

    If (Me.Num_CEP <> 0) Or _
       (M46V999.gfPreenchido(Me.Cod_Tip_Logr)) Or _
       (M46V999.gfPreenchido(Me.Nom_Logr)) Or _
       (M46V999.gfPreenchido(Me.Dsc_Compl_Logr)) Or _
       (M46V999.gfPreenchido(Me.Nom_Bairro)) Or _
       (M46V999.gfPreenchido(Me.Nom_Cid)) Or _
       (M46V999.gfPreenchido(Me.Sig_UF)) Or _
       (Me.Num_Fone <> 0) Then

        If Me.objstcPedEndereco.Num_CEP = Me.Num_CEP And Me.objstcPedEndereco.Dsc_Compl_Logr = Me.Dsc_Compl_Logr Then
            gEnd_Cob_Diferente = False
        Else
            gEnd_Cob_Diferente = True
        End If
    Else
        gEnd_Cob_Diferente = False
    End If
End Property
Public Property Get gDesc_SitPedido() As String
    Select Case Me.Cod_Sit_Ped
    Case e_SitPedido_ComCalculo
        gDesc_SitPedido = "Calculado"
    Case e_SitPedido_SemCalculo
        gDesc_SitPedido = "Sem c�lculo"
    Case e_SitPedido_Transmitido
        gDesc_SitPedido = "Transmitido"
    Case Else
        gDesc_SitPedido = "N�o definido"
    End Select
End Property


Public Property Get gPropostaTransmitida() As Integer
    gPropostaTransmitida = 2
    If Me.Cod_Sit_Ped = e_SitPedido_Transmitido Then
        gPropostaTransmitida = 1
    End If
End Property
Public Property Get gPropostaExpirada() As Integer
    Dim lValidaData As Integer
    gPropostaExpirada = 2
    lValidaData = 15
    
    If Me.Dat_Pri_Calculo <> 0 Then
        If ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_RenYasuda Or ColecaoPedLoc(1).Ren_Con_Seguradora = 572 Then 'Renova��o Yasuda e Maritima
           lValidaData = 30
        End If
        
        If M46V999.gValidadeCotCorr > 0 Then
            lValidaData = M46V999.gValidadeCotCorr
        End If
        
        '11.46 - Empresarial item 4
        'Tranferencia da regra de 917581 - Triunfo para 939571 - BR Insurance - 1174
        If Me.Cod_Ramo = e_CodRamo_Condominio And M46V999.gCod_Corretor = 939571 And (ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_RenCongenere Or ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_SeguroNovo) Then
            lValidaData = 30
        End If
        
        If M46V999.gfDifDatas(Me.Dat_Pri_Calculo, M46V999.gfForData(1, Format(Date, "dd/mm/yyyy"))) > lValidaData Then
            If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
                gPropostaExpirada = 1
            End If
        End If
    End If
    
End Property

Public Property Get gCod_Doc_Segur() As String
    If Me.Tip_Pessoa = e_TipPessoa_Fisica Then
        Select Case Cod_Doc
        Case 1
            gCod_Doc_Segur = "RG / Carteira de identidade"
        Case 2
            gCod_Doc_Segur = "CNH / Carteira Nacional de Habilita��o"
        Case 3
            gCod_Doc_Segur = "Carteira de trabalho"
        Case 4
            gCod_Doc_Segur = "Certid�o de casamento"
        Case 5
            gCod_Doc_Segur = "Certid�o de nascimento"
        Case 6
            gCod_Doc_Segur = "Certificado de reservista"
        Case 7
            gCod_Doc_Segur = "Entidades de Classe"
        Case 8
            gCod_Doc_Segur = "RNE / Identidade estrangeira"
        Case 9
            gCod_Doc_Segur = "Passaporte"
        Case 10
            gCod_Doc_Segur = "PIS"
        Case 12
            gCod_Doc_Segur = "T�tulo de eleitor"
        Case Else
            gCod_Doc_Segur = "Documento desconhecido"
        End Select


    Else
        gCod_Doc_Segur = " "
    End If
End Property


Public Property Get gCod_Acao() As Integer
    Dim objTransm As clsA46V718A    'Objeto da classe clsA46V718A (TAB_PED_TRANSM).
    Dim objstcTransm As stcA46V718B   'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM).

    Dim lPropTabela As String

    gCod_Acao = 0
    lTransmitidaeEmitida = 2
    If gPropostaTransmitida = 1 Then
        Call M46V999.gfObtSYASMULTINI(lPropTabela)
        Set objTransm = New clsA46V718A
        Set objTransm.M46V999 = M46V999

        Set objstcTransm = New stcA46V718B
        objstcTransm.nosso_numero = Me.nosso_numero
        If Not objTransm.mfCarregarColecao(objstcTransm, lPropTabela, "") Then
            'UltimoErro = objTransm.UltimoErro
            Call M46V999.gpGraLog(0, objTransm.UltimoErro)
            GoTo SAIDA
        End If
        If objTransm.Colecao.Count > 0 Then
            Set objstcTransm = objTransm.Colecao.Item(1)
            gCod_Acao = val(objstcTransm.Cod_Acao)
            If val(objstcTransm.Cod_Acao) = 80 Or val(objstcTransm.Cod_Acao) = 81 Then
                If Trim(objstcTransm.Dat_Emissao) <> "" Then
                If M46V999.gfForData(1, CDate(M46V999.gfForData(2, objstcTransm.Dat_Emissao)) + 7) >= M46V999.gHoje Then
                    lTransmitidaeEmitida = 1
                    End If
                End If
            End If
        End If
    End If

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set objTransm = Nothing
    Set objstcTransm = Nothing
    Return
End Property

Public Property Get gDesc_RamoAtiv() As String
    If Tip_Pessoa = e_TipPessoa_Juridica Then
        Select Case Cod_Ramo_Ativ
        Case 1
            gDesc_RamoAtiv = "Agricultura, pecu�ria, produ��o florestal, pesca e aquicultura"
        Case 2
            gDesc_RamoAtiv = "Ind�stria extrativa"
        Case 3
            gDesc_RamoAtiv = "Ind�stria de transforma��o"
        Case 4
            gDesc_RamoAtiv = "Eletrecidade e g�s"
        Case 5
            gDesc_RamoAtiv = "�gua, esgoto, atividades de gest�o de res�duos e descontamina��o"
        Case 6
            gDesc_RamoAtiv = "Constru��o"
        Case 7
            gDesc_RamoAtiv = "Com�rcio; Repara��o de ve�culos automotores e motocicletas"
        Case 8
            gDesc_RamoAtiv = "Transporte, armazenagem ou comunica��es"
        Case 9
            gDesc_RamoAtiv = "Alojamento e alimenta��o"
        Case 10
            gDesc_RamoAtiv = "Informa��o e comunica��o"
        Case 11
            gDesc_RamoAtiv = "Atividades financeiras, de seguros e servi�os relacionados"
        Case 12
            gDesc_RamoAtiv = "Atividades imobili�rias"
        Case 13
            gDesc_RamoAtiv = "Atividades profissionais, cient�ficas e t�cnicas"
        Case 14
            gDesc_RamoAtiv = "Atividades administrativas e servi�os complementares"
        Case 15
            gDesc_RamoAtiv = "Administra��o p�blica, defesa e seguridade social"
        Case 16
            gDesc_RamoAtiv = "Educa��o"
        Case 17
            gDesc_RamoAtiv = "Sa�de humana e servi�os sociais"
        Case 18
            gDesc_RamoAtiv = "Artes, cultura, esporte e recrea��o"
        Case 19
            gDesc_RamoAtiv = "Outras atividades de servi�os"
        Case 20
            gDesc_RamoAtiv = "Servi�os dom�sticos"
        Case 21
            gDesc_RamoAtiv = "Organismos internacionais e outras institui��es extraterritoriais"
        End Select
    Else
        gDesc_RamoAtiv = " "
    End If

End Property
Public Property Get gValorTotal() As Double
    Dim lretorno As Double
    'If Me.Qtd_Parc = 1 Then
    '    lretorno = M46V999.gfTrunca(gValPrLiqTotal + Val_Cust_Emis + Val_IOF, 2)
    'Else
    '    lretorno = gValPrLiqTotal + Val_Cust_Emis + Val_IOF
    'End If
    lretorno = gValPrLiqTotal + Val_Cust_Emis + Val_IOF
    gValorTotal = lretorno
End Property

Public Property Get gValPrLiqTotal() As Double
    Dim lretorno As Double
    Dim lobjstcLoc As stcA46V708B
    Dim lPremioMinimo As Double

    For Each lobjstcLoc In ColecaoPedLoc
        lretorno = lretorno + lobjstcLoc.ValPrTotCober
    Next lobjstcLoc
    'Ajute retirar premio minimo do imobiliario
    If M46V999.gPermitePrInformado = False And M46V111.gCod_plano <> 20 Then
            If Me.Cod_Reg > 100 Then
                lPremioMinimo = Me.getProduto.ProPreMinEndosso
            Else
                lPremioMinimo = Me.getProduto.ProPremioMinimo
            End If
    Else
        lPremioMinimo = lretorno
    End If
    
    If M46V111.gCod_plano = 20 Then ' Respeitar o valor da tab_ped
        lretorno = Me.VAL_PR
    End If
    
    If lretorno > 0 Then
        If lretorno < lPremioMinimo Then
            lretorno = lPremioMinimo
        End If
    ElseIf lretorno < 0 Then
        If lretorno > (lPremioMinimo * -1) And Cod_Produto <= 1076 Then
            lretorno = 0
        End If
    End If

    gValPrLiqTotal = Round(lretorno, 2)

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lobjstcLoc = Nothing
    Return
End Property

Public Property Get gCEP_Formatado() As String
    Dim lAux As String
    If Num_CEP <> 0 Then
        lAux = Format(Num_CEP, "00000000")
        gCEP_Formatado = Mid$(lAux, 1, 5) & "-" & Mid$(lAux, 6, 3)
    Else
        gCEP_Formatado = " "
    End If
End Property

Public Property Get gCustoParcelado() As Boolean
    Select Case Tip_Cust_Emis
    Case 0, 2, 4
        gCustoParcelado = False
    Case Else
        gCustoParcelado = True
    End Select
End Property

Public Property Let gDesc_Banco(val As String)
    gDescrBanco = val
End Property

Public Property Get gDesc_Banco() As String
    If (gDescrBanco = "") Then
        Dim codigo As Integer, Nome As String, nomeAbreviado As String, isCarne As Boolean, isDebito As Boolean, isRestituicao As Boolean, pMensagem As String
        
        If (M46V999.GetBancoByID(Cod_Bco_Cobr, codigo, Nome, nomeAbreviado, isCarne, isDebito, isRestituicao, pMensagem) = False) Then
            gDescrBanco = "..."
        Else
            gDescrBanco = nomeAbreviado
        End If
    End If
    
    gDesc_Banco = gDescrBanco

''20170120: Busca no banco
'    Select Case Cod_Bco_Cobr
'    Case e_Banco_DoBrasil
'        gDesc_Banco = "BANCO DO BRASIL"
'    Case e_Banco_Bradesco
'        gDesc_Banco = "BRADESCO SA"
'    Case e_Banco_Itau
'        gDesc_Banco = "ITAU"
'    Case e_Banco_Sudameris
'        gDesc_Banco = "SUDAMERIS BRASIL"
'    Case e_Banco_Unibanco
'        gDesc_Banco = "UNIAO DE BANCOS"
'    Case e_Banco_Boston
'        gDesc_Banco = "BOSTON"
'    Case e_Banco_Real
'        gDesc_Banco = "BANCO REAL"
'    Case e_Banco_HSBC
'        gDesc_Banco = "HSBC"
'    Case e_Banco_Santander
'        gDesc_Banco = "Santander"
'    Case e_Banco_Banese
'        gDesc_Banco = "Banese"
'    Case e_Banco_Safra
'        gDesc_Banco = "Safra"
'    Case e_Banco_Bancoob
'        gDesc_Banco = "Bancoob"
'    End Select
End Property

Public Property Get gAgenciaDig() As Long
    If Cod_Bco_Cobr = e_Banco_Bradesco Or Cod_Bco_Cobr = e_Banco_DoBrasil Then
        gAgenciaDig = Cod_Agen_Bas_Cor & Cod_Dig_Agen_Bas_Cor
    Else
        gAgenciaDig = Cod_Agen_Bas_Cor
    End If
End Property

Public Property Get gCodSusep() As String
    Dim bdMulti As ADODB.Connection
    Dim rstMulti As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String

    If Not M46V999.gfPreenchido(mNumSusep) Then
        If Not M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, bdMulti, lMensagem) Then
            Call M46V999.gpGraLog(1, lMensagem)
            GoTo SAIDA
        End If

        lSelect = " Select * from tab_cod_produto WITH (NOLOCK) where cod_ramo = " & Me.Cod_Ramo
        lSelect = lSelect & " AND COD_PRODUTO = " & Cod_Estip

        If Not M46V999.gfObtRegistro(bdMulti, lSelect, rstMulti, lMensagem) Then
            Call M46V999.gpGraLog(1, lMensagem)
            GoTo SAIDA
        End If


        If Not rstMulti.EOF Then
            mNumSusep = rstMulti("Num_Proc_Susep")
        End If
    End If
    gCodSusep = mNumSusep

SAIDA:
    GoSub FechaObj
    Exit Property

FechaObj:
    Set bdMulti = Nothing
    Set rstMulti = Nothing
    Return
End Property

Public Property Get gSegmentado() As Boolean
    Dim bdMulti As ADODB.Connection
    Dim rstMulti As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String
    
    gSegmentado = False
    
    If Not M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, bdMulti, lMensagem) Then
        Call M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    lSelect = " Select Segmentado from tab_plano WITH (NOLOCK) where cod_ramo = " & Me.Cod_Ramo
    lSelect = lSelect & " AND COD_PRODUTO = " & Cod_Estip

    If Not M46V999.gfObtRegistro(bdMulti, lSelect, rstMulti, lMensagem) Then
        Call M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    If Not rstMulti.EOF Then
        gSegmentado = rstMulti("Segmentado")
    End If

SAIDA:
    GoSub FechaObj
    Exit Property

FechaObj:
    Set bdMulti = Nothing
    Set rstMulti = Nothing
    Return
End Property
Public Property Get gLetraRamoAtiv() As String

    Select Case Cod_Ramo_Ativ
    Case 1
        gLetraRamoAtiv = "A"
    Case 2
        gLetraRamoAtiv = "B"
    Case 3
        gLetraRamoAtiv = "C"
    Case 4
        gLetraRamoAtiv = "D"
    Case 5
        gLetraRamoAtiv = "E"
    Case 6
        gLetraRamoAtiv = "F"
    Case 7
        gLetraRamoAtiv = "G"
    Case 8
        gLetraRamoAtiv = "H"
    Case 9
        gLetraRamoAtiv = "I"
    Case 10
        gLetraRamoAtiv = "J"
    Case 11
        gLetraRamoAtiv = "K"
    Case 12
        gLetraRamoAtiv = "L"
    Case 13
        gLetraRamoAtiv = "M"
    Case 14
        gLetraRamoAtiv = "N"
    Case 15
        gLetraRamoAtiv = "O"
    Case 16
        gLetraRamoAtiv = "P"
    Case 17
        gLetraRamoAtiv = "Q"
        'Ficha 1000365 - Inclus�o Ramos de Atividade
    Case 18
        gLetraRamoAtiv = "R"
    Case 19
        gLetraRamoAtiv = "S"
    Case 20
        gLetraRamoAtiv = "T"
    Case 21
        gLetraRamoAtiv = "U"
    Case 22
        gLetraRamoAtiv = "V"
    Case 23
        gLetraRamoAtiv = "W"
    Case 24
        gLetraRamoAtiv = "X"
    Case 25
        gLetraRamoAtiv = "Y"
    Case 26
        gLetraRamoAtiv = "Z"
    End Select
End Property

Public Property Get gCodigoRamoAtiv(pLetra As String) As Integer

    Select Case UCase(Trim(pLetra))
    Case "A"
        gCodigoRamoAtiv = 1
    Case "B"
        gCodigoRamoAtiv = 2
    Case "C"
        gCodigoRamoAtiv = 3
    Case "D"
        gCodigoRamoAtiv = 4
    Case "E"
        gCodigoRamoAtiv = 5
    Case "F"
        gCodigoRamoAtiv = 6
    Case "G"
        gCodigoRamoAtiv = 7
    Case "H"
        gCodigoRamoAtiv = 8
    Case "I"
        gCodigoRamoAtiv = 9
    Case "J"
        gCodigoRamoAtiv = 10
    Case "K"
        gCodigoRamoAtiv = 11
    Case "L"
        gCodigoRamoAtiv = 12
    Case "M"
        gCodigoRamoAtiv = 13
    Case "N"
        gCodigoRamoAtiv = 14
    Case "O"
        gCodigoRamoAtiv = 15
    Case "P"
        gCodigoRamoAtiv = 16
    Case "Q"
        gCodigoRamoAtiv = 17
        'Ficha 1000365 - Inclus�o Ramos de Atividade
    Case "R"
        gCodigoRamoAtiv = 18
    Case "S"
        gCodigoRamoAtiv = 19
    Case "T"
        gCodigoRamoAtiv = 20
    Case "U"
        gCodigoRamoAtiv = 21
    Case "V"
        gCodigoRamoAtiv = 22
    Case "W"
        gCodigoRamoAtiv = 23
    Case "X"
        gCodigoRamoAtiv = 24
    Case "Y"
        gCodigoRamoAtiv = 25
    Case "Z"
        gCodigoRamoAtiv = 26
    End Select
End Property

'Ficha 1000837 - Triagem Informar Beneficiario nos Itens
Public Function FindTextoEspec(pobjstc As stcA46V726B) As Boolean
    Dim mObjstc As stcA46V726B
    On Local Error GoTo ControleErros
    FindTextoEspec = False

    For Each mObjstc In ColecaoTextoEspec
        If mObjstc.Cod_Espec = pobjstc.Cod_Espec Then
            FindTextoEspec = True
            Exit For
        End If
    Next

SAIDA:
    GoSub FechaObj
    Exit Function
ControleErros:
    FindTextoEspec = False
    Call M46V999.gpGraLog(0, "stcA46V702B.FindTextoEspec ERRO - " & Err & " " & Err.Description)
    GoTo SAIDA
FechaObj:
    Set mObjstc = Nothing
    Return
End Function

Public Property Get ValTotalComissao() As Double
    Dim lretorno As Double
    Dim lstcPedCorr As stcA46V705B

    For Each lstcPedCorr In ColecaoPedCorr
        lretorno = lretorno + lstcPedCorr.Per_Corr
    Next lstcPedCorr

    ValTotalComissao = lretorno

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lstcPedCorr = Nothing
    Return
End Property

Public Property Get ValTotalProlabore() As Double
    Dim lretorno As Double
    Dim lstcPedProlabore As stcA46V733B

    For Each lstcPedProlabore In ColecaoProlabore
        If lstcPedProlabore.Tip_Canal = "1" Then
            lretorno = lretorno + lstcPedProlabore.Perc_Canal
        End If
    Next lstcPedProlabore

    ValTotalProlabore = lretorno

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lstcPedProlabore = Nothing
    Return
End Property
Public Property Get ValTotalAgenciamento() As Double
    Dim lretorno As Double
    Dim lstcPedAgenciamento As stcA46V733B

    For Each lstcPedAgenciamento In ColecaoAgenciador
        If lstcPedAgenciamento.Tip_Canal = "3" Then
            lretorno = lretorno + lstcPedAgenciamento.Perc_Canal
        End If
    Next lstcPedAgenciamento

    ValTotalAgenciamento = lretorno

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    ValTotalAgenciamento = 0
    Return
End Property

Public Property Get gDescTarifa() As String

    gDescTarifa = " "
    
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstTarifa As clsYasRecordSet
    Dim lSelect As String
    Dim lMensagem As String

    'Rotina Alterada para atender a nova Precifica��o

    '1. Abre arquivo.
    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
        Call M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    '3. Obt�m a descri��o da tarifa
    lSelect = "Exec proc_RetornarDscTarifa " & Cod_Produto
    lSelect = lSelect & ", " & Cod_Ramo

    If M46V999.gfObtRegistro(BDP00Tarifa, lSelect, rstTarifa, lMensagem) = False Then
        Call M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    If Not rstTarifa.EOF Then
        If Cod_Produto >= 1080 Then
            gDescTarifa = rstTarifa("Nom_Prot")
        Else
            gDescTarifa = rstTarifa("Cod_Produto")
        End If
    End If

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set BDP00Tarifa = Nothing
    Set rstTarifa = Nothing
    Return
End Property

Public Property Get gPagAntecipado() As Boolean
    Dim lretorno As Boolean

    If Me.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
        If Me.objstcCtrlEmis.Num_Bloq <> 0 Then
            lretorno = True        'Sim - Antecipado com bloqueto.
        ElseIf (Me.objstcCtrlEmis.Num_Bloq = 0 And Me.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 And Me.Cod_Agen_Bas_Cor <> 0) Then
            lretorno = True        'Sim - Antecipado com D�bito em conta corrente
        Else
            lretorno = True        'Sim - Bloqueto.
        End If
    Else
        lretorno = False
    End If

    gPagAntecipado = lretorno

End Property

Public Property Get gTemBoleto() As Boolean
    Dim lretorno As Boolean

    If Me.Cod_Reg = 106 Or Me.Cod_Reg = 107 Or Me.Cod_Reg = 120 Or Me.Cod_Reg = 129 Then
        lretorno = False
    ElseIf gTipPagamento = e_FormaPagamento_Carne Then
        lretorno = gPagAntecipado
    Else
        lretorno = False
    End If

    gTemBoleto = lretorno

End Property

Public Function getRpt() As stcA46V702rpt
    Static lrpt As stcA46V702rpt

    If lrpt Is Nothing Then
        Set lrpt = New stcA46V702rpt
        Set lrpt.stcOrigem = Me
    End If

    Set getRpt = lrpt

End Function

Public Function gfPosicaoColecao(pItem As Integer) As Integer
    Dim lPosicao As Integer
    Dim lstcItem As stcA46V708B
    Dim lEhUltimo As Boolean

    lEhUltimo = True
    For Each lstcItem In Me.ColecaoPedLoc
        lPosicao = lPosicao + 1
        If pItem <= lstcItem.Num_Item Then
            lEhUltimo = False
            Exit For
        End If
    Next lstcItem
    If lEhUltimo = True Then
        lPosicao = lPosicao + 1
    End If

    gfPosicaoColecao = lPosicao

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lstcItem = Nothing
    Return

End Function

Public Property Get gTipPagamento() As e_FormaPagamento
    Dim lretorno As e_FormaPagamento
    If Cod_Agen_Bas_Cor <> 0 Then
        If Cod_Agen_Bas_Cor = 9999 And Num_Cta_Bas_Cor = 9999999 And Cod_Bco_Cobr = 999 Then
            lretorno = 4
        Else
            lretorno = e_FormaPagamento_Debito
        End If
    Else
        lretorno = e_FormaPagamento_Carne
    End If
    gTipPagamento = lretorno
End Property

Public Property Get gTemAjusVenc() As Boolean
    gTemAjusVenc = (Me.ColecaoPedVenc.Count > 0)
End Property

Public Property Get gTemPlanoComercial() As Boolean
    Dim lretorno As Boolean
    Dim lstcItem As stcA46V708B

    lretorno = False
    For Each lstcItem In Me.ColecaoPedLoc
        'Ficha 1000308 - Permitir itens com planos diferentes no produto Roubo (c�lculo coletivo)
        'Roubo n�o permite c�lculo de itens (Riscos Comerciais e/ou Industriais) com os demais
        If lstcItem.Cod_Ramo = e_CodRamo_Roubos Then
            If lstcItem.Tab_Clasf = 2 Then lretorno = True
        End If
    Next lstcItem

    gTemPlanoComercial = lretorno

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lstcItem = Nothing
    Return
End Property

Public Property Get gTemPlanoNaoComercial() As Boolean
    Dim lretorno As Boolean
    Dim lstcItem As stcA46V708B

    lretorno = False
    For Each lstcItem In Me.ColecaoPedLoc
        'Ficha 1000308 - Permitir itens com planos diferentes no produto Roubo (c�lculo coletivo)
        'Roubo n�o permite c�lculo de itens (Riscos Comerciais e/ou Industriais) com os demais
        If lstcItem.Cod_Ramo = e_CodRamo_Roubos Then
            If lstcItem.Tab_Clasf <> 2 Then lretorno = True
        End If
    Next lstcItem

    gTemPlanoNaoComercial = lretorno

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lstcItem = Nothing
    Return
End Property

Public Property Get gValLMG() As Double
    Dim lretorno As Double
    Dim lstcItem As stcA46V708B
    Dim lobjstcCobert As stcA46V704B    'Objeto da classe (stc) de coberturas.

    For Each lstcItem In Me.ColecaoPedLoc
        For Each lobjstcCobert In lstcItem.ColecaoPedCobert
            If lobjstcCobert.Cod_Cobert = 1 Then
                lretorno = lobjstcCobert.Val_IS
                Exit For
            End If
        Next lobjstcCobert
    Next lstcItem

    gValLMG = lretorno

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lstcItem = Nothing
    Set lobjstcCobert = Nothing
    Return
End Property

Public Function getProduto() As stcProduto
    Static lstcProduto As stcProduto

    If lstcProduto Is Nothing Then
        Set lstcProduto = New stcProduto
        Set lstcProduto.M46V999 = Me.M46V999
        Set lstcProduto.stcOrigem = Me

    End If

    Set getProduto = lstcProduto

End Function

Public Function getPlano() As stcPlano
    Static lstcPlano As stcPlano

    If lstcPlano Is Nothing Then
        Set lstcPlano = New stcPlano
        Set lstcPlano.stcOrigem = Me
    End If

    Set getPlano = lstcPlano

End Function

Public Sub gpAtualizaDatPriCal()
    Dim lCnSyasPortal As ADODB.Connection
    Dim lMsg As String
    Dim lUpdate As String
    Dim lPropTabela As String

    If Me.Dat_Pri_Calculo = 0 Then
        Call M46V999.gfObtSYASMULTINI(lPropTabela)
        If M46V999.gfAbrBasDados("SYASPORTAL", M46V999.gTipSyasPortal, M46V999.gEndSyasPortal, lCnSyasPortal, lMsg) = False Then
            GoTo SAIDA
        End If

        Me.Dat_Pri_Calculo = val(M46V999.gfForData(1, Format(Now, "DD/MM/YYYY")))

        lUpdate = lUpdate & " Update " & lPropTabela & "Tab_Ped"
        lUpdate = lUpdate & " Set Dat_Pri_Calculo = " & Me.Dat_Pri_Calculo
        lUpdate = lUpdate & " WHERE nosso_numero = " & M46V999.gfFormatarCamposSQL(Me.nosso_numero)
        lUpdate = lUpdate & " AND (Dat_Pri_Calculo = 0 OR Dat_Pri_Calculo IS NULL)"

        If M46V999.gfExeSQL(lCnSyasPortal, lUpdate, lMsg) <> 0 Then
            GoTo SAIDA
        End If
    End If

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lCnSyasPortal = Nothing
    Return
End Sub

Public Property Get Beneficiarios() As String
    Dim lStcTextoEspec As stcA46V726B
    Dim lretorno As String


    For Each lStcTextoEspec In Me.ColecaoTextoEspec
        With lStcTextoEspec
            If .Cod_Espec = 8500 And UCase(.Espec_lin20) = "FRMT46V155A" Then
                lretorno = .Espec_lin1
                lretorno = lretorno & IIf(M46V999.gfPreenchido(.Espec_lin2), vbCrLf & .Espec_lin2, "")
                lretorno = lretorno & IIf(M46V999.gfPreenchido(.Espec_lin3), vbCrLf & .Espec_lin3, "")
                lretorno = lretorno & IIf(M46V999.gfPreenchido(.Espec_lin4), vbCrLf & .Espec_lin4, "")
                lretorno = lretorno & IIf(M46V999.gfPreenchido(.Espec_lin5), vbCrLf & .Espec_lin5, "")
                lretorno = lretorno & IIf(M46V999.gfPreenchido(.Espec_lin6), vbCrLf & .Espec_lin6, "")
                lretorno = lretorno & IIf(M46V999.gfPreenchido(.Espec_lin7), vbCrLf & .Espec_lin7, "")
                lretorno = lretorno & IIf(M46V999.gfPreenchido(.Espec_lin8), vbCrLf & .Espec_lin8, "")
                lretorno = lretorno & IIf(M46V999.gfPreenchido(.Espec_lin9), vbCrLf & .Espec_lin9, "")
                lretorno = lretorno & IIf(M46V999.gfPreenchido(.Espec_lin10), vbCrLf & .Espec_lin10, "")
                lretorno = lretorno & IIf(M46V999.gfPreenchido(.Espec_lin11), vbCrLf & .Espec_lin11, "")
                lretorno = lretorno & IIf(M46V999.gfPreenchido(.Espec_lin12), .Espec_lin12, "")
                Exit For
            End If
        End With
    Next

    Beneficiarios = lretorno

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lStcTextoEspec = Nothing
    Return

End Property

Public Property Let Beneficiarios(ByVal strBeneficiarios As String)
    Dim lSeq As Byte
    Dim lTotLin As Byte
    Dim i As Byte
    Dim AuxArray() As String
    Dim lStcTextoEspec As stcA46V726B
    Dim lTextoInc As String
    Dim wAuxArray() As String
    Dim lTextoAux As String

    If Trim(strBeneficiarios) = "" Then Exit Property

    wAuxArray = Split(strBeneficiarios, vbCrLf)
    For i = 0 To UBound(wAuxArray)
        'ReDim Preserve AuxArray(UBound(AuxArray) + 1)
Voltar:
        If Len(wAuxArray(i)) > 80 Then
            lTextoAux = lTextoAux & Left(wAuxArray(i), 80) & vbCrLf
            wAuxArray(i) = Mid(wAuxArray(i), 81)
            GoTo Voltar
        Else
            lTextoAux = lTextoAux & Left(wAuxArray(i), 80) & vbCrLf
        End If
    Next

    strBeneficiarios = lTextoAux
    wAuxArray = Split(strBeneficiarios, vbCrLf)
    lSeq = 0
    lTotLin = 1

    If InStr(strBeneficiarios, vbCrLf) <> 0 Then
        AuxArray = Split(strBeneficiarios, vbCrLf)
        lTextoInc = AuxArray(lSeq)
        lTotLin = UBound(AuxArray)
    Else
        lSeq = 1
        lTextoInc = strBeneficiarios
    End If

    Set lStcTextoEspec = New stcA46V726B

    With lStcTextoEspec
        .nosso_numero = Me.nosso_numero
        .Cod_Org = 0
        .Cod_Ramo = Me.Cod_Ramo
        .Num_Ped = Me.Num_Ped
        .Num_Apol = 0
        .Num_Endo = 0
        .Cod_Espec = 8500
        .Num_pag = Me.ColecaoTextoEspec.Count + 1    '1
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
        .NUM_PI = Me.NUM_PI
    End With
    Call mpRemoveBeneficiario(lStcTextoEspec)
    Me.ColecaoTextoEspec.Add lStcTextoEspec, lStcTextoEspec.Class_ID

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lStcTextoEspec = Nothing
    Return
End Property

Private Sub mpRemoveBeneficiario(ByRef pStcTextoEspec As stcA46V726B)

    On Error GoTo SAIDA

    If Me.ColecaoTextoEspec.Count > 0 Then
        Call Me.ColecaoTextoEspec.Remove(pStcTextoEspec.Class_ID)
    End If

SAIDA:
    Exit Sub

End Sub

Public Property Get CodSitTransmissao() As e_SitTransmissao
    Dim lStatus As e_SitTransmissao

    lStatus = e_SitTransmissao_Nenhum

    'J� transmitido.
    If Me.Cod_Sit_Ped = e_SitPedido_Transmitido Then
        lStatus = e_SitTransmissao_Transmitido
        Select Case Me.gCod_Acao
        Case 80   'Recusada
            lStatus = e_SitTransmissao_Recusado
        Case 81   'Cancelada
            lStatus = e_SitTransmissao_Cancelado
        Case 90   'Emitida
            lStatus = e_SitTransmissao_Emitido
        End Select
    End If

    CodSitTransmissao = lStatus

SAIDA:
    Exit Property

End Property

Public Property Get gQtdMaxParcAnt() As Byte
    Dim lretorno As Byte
    Dim stcParcela As stcParcelamento
    Dim TipoPgto As Integer
    
    TipoPgto = 0
    
    If (objstcCtrlEmis.Cod_Bco_Boleto > 0 And Cod_Bco_Cobr > 0) Or (Cod_Agen_Bas_Cor <> 0 And objstcCtrlEmis.Cod_Bco_Boleto = 0 And Cod_Bco_Cobr > 0) Then
        TipoPgto = 2 'debito
    Else
        TipoPgto = 1 'carne
    End If

    For Each stcParcela In Me.ColecaoParcelamento
        If stcParcela.TipPagamento = TipoPgto Then
        If Not (stcParcela.Val_Primeira_Vista = 0 And stcParcela.Val_Demais_Vista = 0) Then
            lretorno = stcParcela.Num_Parcela
        End If
        End If
    Next
    gQtdMaxParcAnt = lretorno

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set stcParcela = Nothing
    Return
End Property

Public Property Get gQtdMaxParcPos() As Byte
    Dim lretorno As Byte
    Dim stcParcela As stcParcelamento
    Dim TipoPgto As Integer
    
    TipoPgto = 0
    
    If (objstcCtrlEmis.Cod_Bco_Boleto > 0 And Cod_Bco_Cobr > 0) Or (Cod_Agen_Bas_Cor <> 0 And objstcCtrlEmis.Cod_Bco_Boleto = 0 And Cod_Bco_Cobr > 0) Then
        TipoPgto = 2 'debito
    Else
        TipoPgto = 1 'carne
    End If

    For Each stcParcela In Me.ColecaoParcelamento
        If stcParcela.TipPagamento = TipoPgto Then
        If Not (stcParcela.Val_Primeira_30dias = 0 And stcParcela.Val_Demais_30dias = 0) Then
            lretorno = stcParcela.Num_Parcela
            End If
        End If
    Next
    gQtdMaxParcPos = lretorno

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set stcParcela = Nothing
    Return
End Property

Public Property Get ehReentrada() As Boolean
    ehReentrada = M46V999.gfPreenchido(Nosso_Numero_Original)
End Property

Public Property Get BoletoGerado() As Boolean
    Dim lretorno As Boolean
    lretorno = False
    On Error GoTo SAIDA
    If gTemBoleto = True And Me.objstcCtrlEmis.Num_Bloq <> 0 Then
        If val(Me.objstcBoleto.gobjBoletoCompl.valorDocumento) = val(Me.Val_PriParcela) Then
            lretorno = True
        End If
    End If

    'Ajuste forma de pagamento - Carn� + pagto � 30 dias n�o gera boleto
    If Me.Cod_Sit_Ped = e_SitPedido_Transmitido And Me.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 Then
        If lretorno = False Then lretorno = (Me.objstcCtrlEmis.Cod_Bco_Boleto <> 0 And Me.Cod_Bco_Cobr <> 0)
    End If
    BoletoGerado = lretorno

SAIDA:
    Exit Property
End Property

Public Property Get ValidarValorBoletoGerado() As Boolean
    Dim lretorno As Boolean
    lretorno = False
    On Error GoTo SAIDA
    If gTemBoleto = True And Me.objstcCtrlEmis.Num_Bloq <> 0 Then
        If val(Me.objstcBoleto.gobjBoletoCompl.valorDocumento) = val(Me.Val_PriParcela) Then
            lretorno = True
        End If
    End If

    ValidarValorBoletoGerado = lretorno

SAIDA:
    Exit Property
End Property

'1101015 - Projeto 11010 - Renova��o Expressa
Public Property Get Apol_Ant_PER_JURO_FRAC() As Long
    If lApol_Ant_PER_JURO_FRAC = 0 Then
        Call BuscaDadosApoliceRenovacao
    End If
    Apol_Ant_PER_JURO_FRAC = lApol_Ant_PER_JURO_FRAC
End Property
'1101015 - Projeto 11010 - Renova��o Expressa
Public Property Get Apol_Ant_QTD_PARC() As Integer
    If lApol_Ant_QTD_PARC = 0 Then
        Call BuscaDadosApoliceRenovacao
    End If
    Apol_Ant_QTD_PARC = lApol_Ant_QTD_PARC
End Property
'1101015 - Projeto 11010 - Renova��o Expressa
Public Property Get Apol_Ant_PER_CORR() As Long
    If lApol_Ant_PER_CORR = 0 Then
        Call BuscaDadosApoliceRenovacao
    End If
    Apol_Ant_PER_CORR = lApol_Ant_PER_CORR
End Property

Public Property Get Qtd_Parc() As Integer
    Qtd_Parc = Qtd_Parc_
End Property

Public Property Let Qtd_Parc(ByVal Value As Integer)
    Qtd_Parc_ = Value
End Property


'1101015 - Projeto 11010 - Renova��o Expressa
Private Sub BuscaDadosApoliceRenovacao()
    Dim lRdsApol As clsYasRecordSet
    Dim bdP00APOL As ADODB.Connection
    Dim lSelect As String
    Dim lNomeTab As String
    Dim pMensagem As String

    If M46V999.gPrograma = "D046O003" Then
        lNomeTab = "Syasportal.dbo.TAB_APOL"
    Else
        'If M46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
        'Else
        '    lNomeTab = "GED.DBO.TAB_PED"
        'End If
    End If

    lSelect = " SELECT " & lNomeTab & ".QTD_PARC, " & lNomeTab & ".PER_JURO_FRAC, Sum(" & lNomeTab & "_CORR.PER_CORR) AS SomaDePER_CORR "
    lSelect = lSelect & " FROM " & lNomeTab & " WITH (NOLOCK) "
    lSelect = lSelect & " INNER JOIN " & lNomeTab & "_CORR WITH (NOLOCK) "
    lSelect = lSelect & " ON " & lNomeTab & ".NUM_APOL = " & lNomeTab & "_CORR.NUM_APOL "
    lSelect = lSelect & " WHERE (((" & lNomeTab & ".Num_Apol) = " & Me.Num_Apol_An & ") "
    lSelect = lSelect & " AND ((" & lNomeTab & ".Cod_Org) = " & Me.Cod_Org & ") "
    lSelect = lSelect & " AND ((" & lNomeTab & ".Cod_Ramo) = " & Me.Cod_Ramo & ")) "
    lSelect = lSelect & " GROUP BY " & lNomeTab & ".QTD_PARC, " & lNomeTab & ".PER_JURO_FRAC "

    'b) Abre arquivo.
    If M46V999.gfAbrBasDados("P00APOL", M46V999.gTipP00APOL, M46V999.gEndP00APOL, bdP00APOL, pMensagem) = False Then
        GoTo SAIDA
    End If
    'c) Obt�m registro.
    If M46V999.gfObtRegistro(bdP00APOL, lSelect, lRdsApol, pMensagem) = False Then
        GoTo SAIDA
    End If

    If Not lRdsApol.EOF Then
        lApol_Ant_QTD_PARC = M46V999.gfTratarNulo(lRdsApol("QTD_PARC"), tpoNumerico)
        lApol_Ant_PER_JURO_FRAC = M46V999.gfTratarNulo(lRdsApol("PER_JURO_FRAC"), tpoNumerico)
        lApol_Ant_PER_CORR = M46V999.gfTratarNulo(lRdsApol("SomaDePER_CORR"), tpoNumerico)
    End If

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Call M46V999.gpFecha2(bdP00APOL)
    Set bdP00APOL = Nothing
    Return
End Sub

Public Property Get ParcelasValidas() As String
    Dim lretorno As String
    Dim lNumParcValida As String
    Dim lParcela As stcParcelamento
    Dim AuxParcela As Integer

    AuxParcela = 0
    For Each lParcela In Me.ColecaoParcelamento
        lretorno = lretorno & IIf(lretorno = "" Or lNumParcValida = "", "", ", ")
        lretorno = lretorno & lNumParcValida
        lNumParcValida = ""
        
        'ADICIONADO TRATAMENTO PARA N�O EXIBIR MENSAGEM DUPLICADA
        If lParcela.EhParcelaValida = True And lParcela.Num_Parcela <> AuxParcela Then
            AuxParcela = lParcela.Num_Parcela
            lNumParcValida = Format(lParcela.Num_Parcela, "00") & IIf(lParcela.Val_Primeira_Vista <> 0, IIf(lParcela.Num_Parcela = 1, "(A Vista)", "(1� A Vista)"), "") & IIf(lParcela.Val_Primeira_30dias <> 0, "(30 Dias)", "")
        End If
    Next
    'Posiciona a ultima!
    lretorno = lretorno & IIf(lretorno = "" Or lNumParcValida = "", "", " ou ")
    lretorno = lretorno & lNumParcValida

Retorno:
    ParcelasValidas = lretorno
SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lParcela = Nothing
    Return
End Property

'Fichas 1200168 e 1200169 Inclus�o dos campos de listagem Renda e Faixa de Faturamento Pessoa f�sica e jur�dica
Public Property Get gDesc_Renda_Faturamento() As String
    If Tip_Pessoa = e_TipPessoa_Juridica Then
        Select Case Val_RendaFatu
        Case 0
            gDesc_Renda_Faturamento = "N�o desejo informar"
        Case 1
            gDesc_Renda_Faturamento = "At� R$ 500 mil"
        Case 2
            gDesc_Renda_Faturamento = "De R$ 500 mil a 1 milh�o"
        Case 3
            gDesc_Renda_Faturamento = "De R$ 1 milh�o a 3 milh�es"
        Case 4
            gDesc_Renda_Faturamento = "De R$ 3 milh�es a 6 milh�es"
        Case 5
            gDesc_Renda_Faturamento = "De 6 milh�es a 15 milh�es"
        Case 6
            gDesc_Renda_Faturamento = "De 15 milh�es a 40 milh�es"
        Case 7
            gDesc_Renda_Faturamento = "De R$ 40 milh�es a 80 milh�es"
        Case 8
            gDesc_Renda_Faturamento = "De R$ 80 milh�es a 150 milh�es"
        Case 9
            gDesc_Renda_Faturamento = "Acima de 150 milh�es"
        End Select
    Else
        Select Case Val_RendaFatu
        Case 0
            gDesc_Renda_Faturamento = "N�o desejo informar"
        Case 1
            gDesc_Renda_Faturamento = "Menor ou igual a � sal�rio m�nimo"
        Case 2
            gDesc_Renda_Faturamento = "Maior que � sal�rio m�nimo at�1 sal�rio m�nimo"
        Case 3
            gDesc_Renda_Faturamento = "Maior que 1 sal�rio m�nimo at�2 sal�rios m�nimos"
        Case 4
            gDesc_Renda_Faturamento = "Maior que 2 sal�rios m�nimos at� 3 sal�rios m�nimos"
        Case 5
            gDesc_Renda_Faturamento = "Maior que 3 sal�rios m�nimos at� 5 sal�rios m�nimos"
        Case 6
            gDesc_Renda_Faturamento = "Maior que 5 sal�rios m�nimos at� 10 sal�rios m�nimos"
        Case 7
            gDesc_Renda_Faturamento = "Maior que 10 sal�rios m�nimos at� 20 sal�rios m�nimos"
        Case 8
            gDesc_Renda_Faturamento = "Maior que 20 sal�rios m�nimos"
        End Select
    End If
End Property

'Inclu�do 27/03/2012 - solicita��o da Thais - Email
Public Property Get gCodProfissao() As Long
    gCodProfissao = val(Nom_Empresa)
End Property

'Inclu�do 27/03/2012 - solicita��o da Thais - Email
Public Property Get gDescProfissao() As String
    Dim lrst As clsYasRecordSet
    Dim bdGed As ADODB.Connection
    Dim lSelect As String
    Dim pMensagem As String

    If IsNumeric(Nom_Empresa) And Len(Nom_Empresa) = 6 And M46V999.gPrograma = "D046O003" Then
        'b) Abre arquivo.
        If M46V999.gfAbrBasDados("GED", M46V999.gTipGed, M46V999.gEndGed, bdGed, pMensagem) = False Then
            GoTo SAIDA
        End If
        lSelect = " Select * from P0044800.dbo.T0044811 WITH (NOLOCK) where COD_OCUPACAO = " & Left(Nom_Empresa, 4) & " and  SUB_CODIGO_OCUP = " & Right(Nom_Empresa, 2)
        If Not M46V999.gfObtRegistro(bdGed, lSelect, lrst, pMensagem) Then
            Exit Sub
        End If
        If Not lrst.EOF Then
            gDescProfissao = lrst("OCUPACAO")
        End If
    Else
        gDescProfissao = Nom_Empresa
    End If
SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Call M46V999.gpFecha2(bdGed)
    Return
End Property

Public Property Get gObrRegistro() As Boolean
    Dim lItem As stcA46V708B
    Dim lretorno As Boolean
    For Each lItem In Me.ColecaoPedLoc
        lretorno = lItem.gTemRegProfissoes
        If lretorno = True Then
            Exit For
        End If
    Next lItem
    Set lItem = Nothing
    gObrRegistro = lretorno

End Property

Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
End Property

Public Property Get M46V999() As clsM46V999
    If mClsM46V999 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V999 n�o atribu�do para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V999 j�
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46V999 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46V999 = mClsM46V999
End Property

Public Property Get M46V111() As clsM46V111
    Set M46V111 = M46V999.M46V111
End Property

Public Property Get M46O001() As clsM46O001
    Set M46O001 = M46V999.M46O001
End Property

Public Property Get gApoliceQuitada() As Boolean
    Dim bdGed As ADODB.Connection
    Dim lSql As String
    Dim lrst As clsYasRecordSet
    Dim lMensagem As String

    If Me.Cod_Dizer = 10 Then
        'b) Abre arquivo.
        If M46V999.gfAbrBasDados("YAS_ND", "1", "", bdGed, lMensagem) = False Then
            Call M46V999.gpGraLog(0, lMensagem)
            Exit Property
        End If
        lSql = "SELECT count(*) as total FROM YAS.DT1102M WHERE APOLICE=" & Me.Num_Apol_An
        lSql = lSql & " and CODCOBR <> 1"

        If Not M46V999.gfObtRegistro(bdGed, lSql, lrst, lMensagem) Then
            Call M46V999.gpGraLog(0, lMensagem)
            Exit Property
        End If
        If lrst("TOTAL") = 0 Then
            gApoliceQuitada = True
        End If
    End If
End Property

Public Property Get gStcPedidoApolice() As stcA46V702B
    Dim lObj46V101A As clsA46V101A
    Static lsNum_Apol_An As Double

    If Me.Num_Apol_An <> 0 And (lsNum_Apol_An <> Me.Num_Apol_An) Then
        lsNum_Apol_An = Me.Num_Apol_An
        Set lObj46V101A = New clsA46V101A
        Set lObj46V101A.M46V999 = M46V999
        Call lObj46V101A.mfCarregarDadosOriginais(Me, mStcPedidoApolice)
    End If

    'Retorno
    Set gStcPedidoApolice = mStcPedidoApolice
SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lObj46V101A = Nothing
    Return
End Property

Public Sub addItemPedidoApolice(ByVal pobjPedLoc As stcA46V708B)
    Dim lObj46V101A As clsA46V101A
   
    Set lObj46V101A = New clsA46V101A
    Set lObj46V101A.M46V999 = M46V999
    Call lObj46V101A.mfCarregarItemOriginal(pobjPedLoc, mStcPedidoApolice)
    
SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lObj46V101A = Nothing
    Return
End Sub

Public Function getstcJuros(ByVal pNParcela As Integer, ByVal pTip_Compensacao As e_FormaPagamento) As stcA46V711B
    Dim Retorno As stcA46V711B
    Dim lobjstcJuros As stcA46V711B

    For Each lobjstcJuros In Me.ColecaoPedJuros
        If lobjstcJuros.Num_Parcela = pNParcela And lobjstcJuros.Tip_Compensacao = pTip_Compensacao Then
            Set Retorno = lobjstcJuros
            Exit For
        End If
    Next

    Set getstcJuros = Retorno
End Function

Public Property Let gDesc_BancoRestituicao(val As String)
    gDescrBancoRestituicao = val
End Property

Public Property Get gDesc_BancoRestituicao() As String
    gDesc_BancoRestituicao = gDescrBancoRestituicao
    'Select Case objstcCtrlEmis.Cod_Banco_Cre
    'Case e_Banco_DoBrasil
    '    gDesc_BancoRestituicao = "BANCO DO BRASIL"
    'Case e_Banco_Bradesco
    '    gDesc_BancoRestituicao = "BRADESCO SA"
    'Case e_Banco_Itau
    '    gDesc_BancoRestituicao = "ITAU"
    'Case e_Banco_Sudameris
    '    gDesc_BancoRestituicao = "SUDAMERIS BRASIL"
    'Case e_Banco_Unibanco
    '    gDesc_BancoRestituicao = "UNIAO DE BANCOS"
    'Case e_Banco_Boston
    '    gDesc_BancoRestituicao = "BOSTON"
    'Case e_Banco_Real
    '    gDesc_BancoRestituicao = "BANCO REAL"
    'Case e_Banco_HSBC
    '    gDesc_BancoRestituicao = "HSBC"
    'Case e_Banco_Santander
    '    gDesc_BancoRestituicao = "Santander"
    'Case e_Banco_Banese
    '    gDesc_BancoRestituicao = "Banese"
    'Case e_Banco_Safra
    '    gDesc_BancoRestituicao = "Safra"
    'Case e_Banco_Bancoob
    '    gDesc_BancoRestituicao = "Bancoob"
    'End Select
End Property

Public Property Get gColecaoPedLoc_ByNum_Item(ByVal pIndice As Integer) As stcA46V708B
    Dim lobjTabPedLoc As stcA46V708B
    
    Set lobjTabPedLoc = Nothing
    For Each lobjTabPedLoc In ColecaoPedLoc
        If lobjTabPedLoc.Num_Item = pIndice Then
            Set gColecaoPedLoc_ByNum_Item = lobjTabPedLoc
            Exit For
        End If
    Next
    
End Property
