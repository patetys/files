VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V712B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================================================================================
' Class Name: stcA46V712B -- Tab_Ctrl_Emis
' Data/Hora: 19/09/2005 - 10:20:46
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================
Public nosso_numero As String           'PK               'Nosso n�mero
Public NUM_PI As Double   'NUMERO DO PROCESSO INTERNO
Public Num_Bloq As Long     'NUMERO DO BLOQUETO SEM DIGITO VERIFICADOR
Public COD_UNID As Long     'CODIGO DA UNIDADE PRODUTORA
Public Cod_Ramo As Long     'CODIGO DO RAMO
Public Dat_Ini_Vig As Long        'DATA DO INICIO DA VIGENCIA
Public Dat_Ter_Vig As Long        'DATA DO TERMINO DA VIGENCIA
Public Cod_Corr As Long     'CODIGO DO CORRETOR
Public Per_Corr As Double   'PERCENTUAL DE COMISSAO DO CORRETOR 1
Public Cod_Corr_2 As Long       'CODIGO DO CORRETOR 2
Public Per_Corr_2 As Double     'PERCENTUAL DE COMISSAO DO CORRETOR 2
Public Cod_Corr_3 As Long       'CODIGO DO CORRETOR 3
Public Per_Corr_3 As Double     'PERCENTUAL DE COMISSAO DO CORRETOR 3
Public Nom_Seg As String  'NOME DO SEGURADO
Public Qtd_Anex As Long     'QUANTIDADE DE ANEXOS
Public Num_Apol_Ant As Double         'CODIGO DO CORRETOR??
Public Dsc_Obs_Prop As String       'OBSERVACAO
Public Tip_Cotac As String    'TIPO DE COTACAO
Public Num_Cotac As String    'NUMERO DA COTACAO
Public Per_Juros_Neg As Double        'PERC.JUROS NEGOCIAVEL
Public Cod_Envi_Doc As String       'CODIGO DE ENVIO DE ETIQ
Public DAT_DIGIT As Long      'DATA DA DIGITACAO
Public HOR_DIGIT As Long      'HORA DA DIGITACAO
Public Cod_Iden_Stat_Acao As Long               'CODIGO DO STATUS DA ACAO
Public Cod_User_Digit As String         'CODIGO DO USUARIO
Public Dat_Digit_MF As Long         'DATA DA DIGITALIZACAO
Public Hor_Digit_MF As Long         'HORA DA DIGITALIZACO
Public Cod_User_Digit_MF As String            'CODIGO DO USUARIO DA DIGITALIZACAO
Public Val_Pr_Tot As Double     'VALOR DO PREMIO
Public Cod_Iden_Classe As Long            'CLASSE DA NORMA
Public Cod_Iden_Cobr As Long          'CODIGO DE CANCELAMENTO DO COBRANCA
Public Num_Prop As Long     'NUMERO DA PROPOSTA
Public Dat_Canc_Apol As Long          'DATA DE CANCELAMENTO DO DOCUMENTO
Public Cod_Tip_Doc As Long        'TIPO DE DOCUMENTO:APOL/ENDO/RENOV
Public Dat_Rec_Cia As Long        'DATA DE RECEBIMENTO NA COMPANHIA
Public Num_Ref_Corr As String       'NUMERO DE REFERENICA DO CORRETOR
Public Dat_Emis As Long     'DATA DE EMIS DO DOCUMENTO
Public Num_Apol As Double     'NUMERO DA APOLICE
Public Num_Endo As Long     'NUMERO DO ENDOSSO
Public Val_Pr_Emis As Double      'VALOR DO PREMIO EMITIDO
Public Val_Pr_Liq As Double     'VALOR DO PREMIO LIQUIDO
Public Per_Tot_Coss As Double       'PERC.TOTAL DE COSSEGURO
Public Qtd_Parc As Long     'QUANTIDADE DE PARCELAS
Public Cod_Iden_Stat_Resp As Long               'STATUS DA RESPOSTA
Public Cod_Iden_SAC As Long         'CODIGO IDENTIFICADOR DE ENVIO SAC
Public Cod_Prod As Long     'CODIGO DO PRODUTOR
Public Val_IS As Double    'VALOR DA IS
Public Per_Desc_Coml As Double        'PERCENTUAL DO DESCONTO COMERCIAL
Public Num_Cgc_Cpf As String      'NUMERO DO CGC/CPF
Public Cod_Iden_Pessoa As e_TipPessoa     'Tipo de pessoa: 1=PJ , 2=PF
Public Cod_Iden_Pag_Antec As Byte    'e_SimNao           'PAGAMENTO ANTECIPADO 1=SIM 2=NAO ***MOSILVA VERIFICAR!!!
Public DIG_BLOQUETO As Byte
Private mDigito As String
Private mBloquetoDigito As Long


Public Parcelas_Impressao_Avista As String
Public Parcelas_Impressao_30dias As String
Public Parcelas_Impressao_AvistaDebito As String
Public Parcelas_Impressao_30diasDebito As String
Public Parcelas_Impressao_Cartao As String
Public Email_Contato As String
Public Ind_Impr_Apol As Byte
Public Ind_Prolabore As Byte
Public Ind_Agenciamento As Byte
Public Cod_Bco_Boleto As Integer
Public Per_Desc_Agrav As Double


Public Cod_Bco_Dev As Integer
Public Cod_Agen_Dev As Long
Public Dig_Agen_Dev As Integer

Public cod_origem As Integer
Public Num_Apol_Mar As Double
Public Num_Logr_Segurado As String
Public Dsc_Compl_Segurado As String
Public Num_Logr_Cobranca As String
Public Dsc_Compl_Cobranca As String
Public Estado_Civil As Byte
Public DAT_Nascimento As Long

Public Cod_Banco_Cre As Integer
Public Cod_Agenc_Cre As Long
Public Dig_Agenc_Cre As Integer
Public Cod_Conta_Cre As Double
Public Cpf_Cnpj_Cre As String
Public Tip_Res_Cre As Integer
Public Cod_DIG_Conta As String

Public Ind_Cond_Constituido As Integer
Public Indice_IPC As Double
Public Ind_Promocional As Integer
Public Ind_Texto_Rel_Bens As Integer
Public Suprimir_Batente As Boolean
Public Ind_Bloq_Web As Integer
Public Ind_Envio_Alcada As Integer
Public Val_VerbaCom As Double
Public per_corr_max As Double

Public Per_Max_Comis_Corr As Integer
Public Dat_Ini_Vig_Orig As Long
Public Dat_Ter_Vig_Orig As Long
Public Tip_Classe As Integer
Public Val_Pr_C4 As Double
Public Num_Ped_Ant As Double
Public Num_Cotac_Ant As String
Public RA_Comissao As Integer
Public Per_Comis_Corr_RA As Double
Public Ind_Org_Publico As Boolean

' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B
' ======================================================================================================================================================
Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero)
End Property

Public Property Set ObjPai(pobjstc As stcA46V702B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property

Public Property Get NomeProdutor() As String
    If Not ObjStcPai Is Nothing Then
        NomeProdutor = ObjStcPai.gNom_Produtor
    End If
End Property

Public Property Get NomeUnidadeProdutora() As String
    If Not ObjStcPai Is Nothing Then
        NomeUnidadeProdutora = ObjStcPai.gNom_UnidadeProd
    End If
End Property

Public Property Get AgenciaDig_Cre() As String
    Dim lretorno As String
    Dim clsDigito As clsA99V001
    Dim lDigito As String
    Dim lMensagem As String
    
'    If Cod_Banco_Cre <> 0 And Cod_Agenc_Cre <> 0 Then
'
'        Select Case Cod_Banco_Cre
'
'            Case e_Banco_Bradesco
'                lretorno = Cod_Agenc_Cre & Dig_Agenc_Cre
'                lretorno = Left(Format(lretorno, "00000"), 4) & "-" & Right(lretorno, 1)
'
'            Case e_Banco_DoBrasil
'                lretorno = Cod_Agenc_Cre & Dig_Agenc_Cre
'                If Right(lretorno, 1) = "0" Then
'                    Set clsDigito = New clsA99V001
'                    Set clsDigito.M46V999 = ObjStcPai.M46V999
'
'                    Call clsDigito.DIG_BANCOBRASIL(Format(lretorno, "00000000"), lDigito, lMensagem)
'                    If lDigito = "X" Then lretorno = Left(Format(lretorno, "00000"), 4) & "X"
'                End If
'                lretorno = Left(Format(lretorno, "00000"), 4) & "-" & Right(lretorno, 1)
'
'            Case Else
                lretorno = Cod_Agenc_Cre
'        End Select
'    End If
    
    lretorno = IIf(lretorno = "", " ", lretorno)
    AgenciaDig_Cre = lretorno
End Property

Public Property Get gAgenciaDig() As Long
    If Cod_Banco_Cre = e_Banco_Bradesco Or Cod_Banco_Cre = e_Banco_DoBrasil Then
        gAgenciaDig = Cod_Agenc_Cre & Dig_Agenc_Cre
    Else
        gAgenciaDig = Cod_Agenc_Cre
    End If
End Property

