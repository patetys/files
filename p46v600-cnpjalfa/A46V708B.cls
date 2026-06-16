VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V708B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V708B --> Tabela Tab_Ped_Loc
' Data/Hora: 31/08/2005 - 11:54:41
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================

Public nosso_numero As String           'Nosso número.
Public Num_Item As Long         'PK                 'Número do item.
Public Num_Ped As Long         'PK                 'Número do pedido.
Public Cod_Org As Integer     'Código do órgão.
Public Cod_Ramo As Integer      'Código do ramo.   Simony
Public Num_Apol As Double         'Número da apólice.
Public Num_Endo As Long         'Número do endosso.
Public Cod_Tip_Logr_Risco As String                 'Tipo de logradouro - Risco.
Public Nom_Logr_Risco As String             'Nome do logradouro - Risco.
Public Dsc_Compl_Risco As String              'Complemento do logradouro - Risco.
Public Nom_Bairro_Risco As String               'Bairro - Risco.
Public Nom_Cid_Risco As String            'Cidade - Risco.
Public Sig_UF_Risco As String           'UF - Risco.
Public Num_Cep_Risco As Long              'CEP - Risco.
Public Per_Desc_Apar As Double            'Percentual de desconto do aparelho.
Public Per_Desc_Progres As Double               'Percentual de desconto progressivo.
Public Per_IS_Vr As Double        'Percentual da IS do valor de risco.
Public Val_Coef_IS_Vr As Double             'Valor do coeficiente da IS do valor de risco.
Public Val_Coef_IS_Vr_LC As Double                'Valor do coeficiente da IS do valor de risco (Lucros Cessantes).
Public Val_Vr As Double     'Valor de risco.
Public Tab_Clasf As Integer       'Tab classificação.
Public Cod_Clasf As String        'Código da classificação.
Public Cod_Compl_Clasf As String              'Código complementar classificação.
Public Dsc_Clasf As String        'Descrição da classificação.
Public Cod_Classe_Ocup As Integer
'Public Cod_Constr As e_CodContr     'Código da construção.
Public Cod_Constr As Integer     'Código da construção.
Public Cod_User_Digit As String             'Usuário de digitação.
Public DAT_DIGIT As Long          'Data da digitação (yyyymmdd).
Public HOR_DIGIT As Integer       'Hora da digitação (hhmm).
Public Cod_User_Liber As String             'Usuário de liberação.
Public Dat_Liber As Long          'Data da liberação (yyyymmdd).
Public Hor_Liber As Integer       'Hora da liberação (hhmm).
Public Cod_Sit_item As e_SitItem        'Situação do item.
Public Per_Bonus As Double        'Percentual de bonus
Public Qtd_Empreg As Integer        'Quantidade de empregados.
Public Num_Insp As Long         'Número da inspeção
Public Cod_Iden_Segur_01_V As String                  'Vigilância exclusiva e permanente do imóvel entendendo-se ainda como tal, porteiros fixos.: S ou N
Public Cod_Iden_Segur_02_A As String                  'Alarmes: S ou N
Public Cod_Iden_Segur_03_P As String                  'Porteiros eletrônicos: S ou N
Public Cod_Iden_Segur_04_I As String                  'Inexistência de terreno baldio,  imóvel desocupado   ou em construção na vizinhança do risco.: S ou N
Public Cod_Iden_Segur_05_G As String                  'Grades metálicas de proteção em todas as janelas e portas de vidros que dão acesso ao interior do imóvel desde que também possuam alarme.: S ou N
Public Cod_Iden_Segur_06 As String
Public Cod_Iden_Segur_07 As String
Public Cod_Iden_Segur_08 As String
Public Cod_Iden_Segur_09 As String
Public Cod_Iden_Segur_10 As String
Public Per_Tot_Segur As Double
Public Cod_Iden_Protec_01_E As String                   'Indicador de Extintor: S ou N.
Public Cod_Iden_Protec_02_H As String                   'Indicador de Hidrantes: S ou N.
Public Cod_Iden_Protec_03_S As String                   'Indicador de Spreinklers: S ou N.
Public Cod_Iden_Protec_04_D As String                   'Indicador de Detecção e alarmes: S ou N.
Public Cod_Iden_Protec_05 As String
Public Cod_Iden_Protec_06 As String
Public Cod_Iden_Protec_07 As String
Public Cod_Iden_Protec_08 As String
Public Cod_Iden_Protec_09 As String    'Será utilizada para protecionais de Incendio
Public Cod_Iden_Protec_10 As String    'Será utilizada para protecionais de Roubo
Public Per_Tot_Protec As Double
Public Cod_Produto As Long            'Plano
Public Cod_Regiao As Integer        'Código da região.
Public Cod_Iden_Tarif As Long
Public NUM_PI As Double     'Número do PI
Public Num_Item_Ant As Long             'Número do item anterior.
Public Tip_Emissao As e_TipEmissao    'Tipo de emissão.
Public Num_Apol_Anterior As Double                  'Número da Apólice anterior.
Public Num_Item_Anterior As Double                  'Número do item anterior.
Public Ind_Sinistro As Byte         'e_SimNao                 'Indicador de sinistro. ***MOSILVA Verificar!!!
Public Insp_Obr As Boolean      'Obrigatório Inspeção?
Public Insp_Agend As Boolean        'Inspeção agendada?
Public Dig_Bens As Boolean      'Bens digitados?
Public Cob_Bens As Boolean      'Coberturas consiste bens: Residencial (27, 29, 30, 31) Empresarial (94)
Public mBonus_Anterior As Long                'Bonus da apólice.
Public Tip_Atividade As Integer           'Tipo de Atividade (1 - Comercial / 2 - Industrial )

'1000347 - Inclusão de informações da congênere em renovações
Public Ren_Con_Seguradora As Integer                'Renovação de congênere: código da seguradora.
Public Ren_Con_Apolice As String              'Renovação de congênere: número da apólice de renovação.
Public Ren_Con_Item As String           'Renovação de congênere: número do item da apólice de renovação.

'*#20180829'    *Finalidade: Apura Desconto Comercial Congenere
Public DescComlCongenereIncorporado As Integer

'Ficha 1000305
Public Tip_Cobert As e_TipCobert    'Código do tipo de cobertura

'Ficha 1200134
Public Ind_Sin_Inspec As Integer         'Flag de Inspeção Obrigatório por causa de Sinistro
Public Cod_DizerItem As Integer

'8000042754 - Resposta dada a questão parametrizada para o plano.
Public RespostaPergunta As Integer          'Resposta dada a questão parametrizada para o plano.

Private mCepAnt As Double
Private mCodRegiao As Long
Public Enum e_TipEmissao
    e_TipEmissao_SeguroNovo = 100
    e_TipEmissao_RenCongenere = 101
    e_TipEmissao_RenYasuda = 200
    e_TipEmissao_Endosso = 300
End Enum

Public Cod_Plano As Integer
' ======================================================================================================================================================
Public ColecaoPedCobert As Collection
Public ColecaoPedLocBem As Collection
Public ColecaoPedLocBicicleta As Collection
Public ColecaoPedSeguro As Collection
Public ColecaoPedFator As Collection
Public objstcPedObs As stcA46V714B
Public objstcPedInsp As stcA46V717B
Public ColecaoPedClau As Collection
'Public objstcPedClau       As stcA46V703B
Public ColecaoLMI As Collection

Public ColecaoBeneficiario As Collection
Public ColecaoRegraBatente As Collection
Public ColecaoTelefones As Collection
Public ColecaoTextos As Collection
Public ColecaoAllRisks As Collection
Public ColecaoObrasArte As Collection
Public objstcPedInquilino As stcA46V735B
Public objstcPedProprietario As stcA46V735B

Private mCod_Regiao As Long
Private mDesc_Regiao As String

''2017-07-28: Juliano - Ind_SinistroSimNao
Public Ind_Sinistro_ORIG As Byte


Public Enum e_SitItem
    e_SitItem_NaoCalc = 1
    e_SitItem_Calculado = 2
End Enum

Public Enum e_CodContr
    e_CodContr_Superior = 1
    e_CodContr_Inferior = 2
    e_CodContr_Mista = 3
End Enum

Public Enum e_CodContr_NovaTarifa
    e_CodContr_NvTarifa_Superior = 1
    e_CodContr_NvTarifa_Solida = 2
    e_CodContr_NvTarifa_Mista = 3
    e_CodContr_NvTarifa_Inferior = 4
End Enum


Public Enum e_TipCobert
    e_TipCobert_PredioeConteudo = 0
    e_TipCobert_Predio = 1
    e_TipCobert_Conteudo = 2
End Enum

Public Enum e_TipPlanoVida
    e_TipPlanoVida_NaoContratado = 0
    e_TipPlanoVida_CapitalGlobal = 1
End Enum

' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B
Public ColecaoNormas As Collection

'1101015 - Projeto 11010 - Renovação Expressa
Private lApol_Ant_PER_DESC_COML As Double
Private lApol_Ant_IND_SINISTRO As Integer

Private mStcItemApolice As stcA46V708B    '****Nunca mudar para Public!!!!

Private mClsM46V999 As clsM46V999
Private mClsM46O001 As clsM46O001

Private mLiberaDescontoEspecial As Boolean

Public Enum e_TipLocPessoa
    e_TipLocPessoa_Beneficiario = 1
    e_TipLocPessoa_Inquilino = 2
    e_TipLocPessoa_Proprietario = 3
End Enum


Public Ind_VerbaSep As Integer
Public Val_Vr_Lc As Double
Public Val_LMG_Lc As Double
Public Per_IS_Vr_Lc As Double
Public Coef_IS_Vr_Lc As Double

Public Tip_Plano_Vida As Integer
Public Cod_Plano_Cobertura As Integer
Public Qtd_Vidas As Long
Public Val_LMG_Vida As Double
Public Ind_Aux_Funeral As Integer
Public Val_LMG_Funeral As Double
Public Ind_Inc_Conjuge As Integer
Public Val_Aux_Cesta As Double
Public ClausulaNaoDepreciacao As Integer 'JIRA-25945 - Opção "Sem Depreciação" no Condomínio
Public Possui_Elevador As Integer
Public BonusAlterado As Boolean
Public Txt_Observacoes As String
Public Flg_EnviarCertificado  As Integer
Public Ind_StatusInspecao As Integer
Public Val_Vr_Orig As Double
Public Val_Bas_Orig As Double
Public Val_Vr_Lc_Orig As Double
Public ITEM_ORIGEM_C4 As Integer

Public Enum e_StatusInspecao
    e_StatusInspecao_ConfNormas = 0
    e_StatusInspecao_Dispensada = 1
    e_StatusInspecao_obrigatoria = 2
    e_StatusInspecao_programada = 3
End Enum
Public Val_Ass24h As Double
Public ValCustoInspecao As Double
Public CupomDesconto As String
Public Property Get gISTotal_LC() As Double
    Dim lretorno As Double
    Dim i As Integer
    lretorno = 0
    If Me.ColecaoPedCobert.Count > 0 Then
        For i = 1 To Me.ColecaoPedCobert.Count
            If Me.ColecaoPedCobert(i).gCod_Grupo_Cobert = 3 Then
                lretorno = Me.ColecaoPedCobert(i).Val_IS + lretorno
            End If
        Next
    End If
    gISTotal_LC = lretorno
    Me.Val_LMG_Lc = gISTotal_LC
    If Me.Val_Vr_Lc > 0 And gISTotal_LC > 0 Then
        On Error Resume Next
        Me.Per_IS_Vr_Lc = gISTotal_LC / Me.Val_Vr_Lc
    End If
End Property

Public Property Get Class_Name() As String
    Class_Name = "A46V708B"
End Property

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Num_Item, "00000"))
End Property

Public Property Get gChaveClasf() As String
    gChaveClasf = Me.Cod_Clasf & String(4 - Len(Me.Cod_Clasf), " ") & Format(Me.Cod_Compl_Clasf, "00") & Format(Me.Tab_Clasf, "0000")
End Property

Public Property Get gDesc_Regiao() As String
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstTab_Regiao As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String

    'Rotina Alterada para atender a nova Precificação

    If M46V999.gCodProduto > M46V999.TarifaAntiga Then
        mDesc_Regiao = ""
        mCod_Regiao = 0
        Exit Sub
    End If

    If mCod_Regiao <> Me.Cod_Regiao Or mDesc_Regiao = "" Or Cod_Produto = 0 Then
        If Me.Cod_Regiao <> 0 Then
            '1. Abre arquivo.
            If M46V999.gfAbrBasDados("P00MULTI2", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
                Call M46V999.gpGraLog(1, lMensagem)
                Exit Property
            End If

            '3. Obtém produto sem término de vigência ou com término de vigência superior à data do dia.
            lSelect = " SELECT TAB_REGIAO.DSC_REGIAO, TAB_REGIAO.COD_REGIAO,TAB_MULTI_REGIAO.COD_PRODUTO "
            lSelect = lSelect & " FROM TAB_MULTI_REGIAO WITH (NOLOCK) INNER JOIN TAB_REGIAO WITH (NOLOCK) ON "
            lSelect = lSelect & " TAB_MULTI_REGIAO.COD_REGIAO = TAB_REGIAO.COD_REGIAO"
            lSelect = lSelect & " Where cod_Ramo = " & Cod_Ramo
            lSelect = lSelect & " and DAT_TER_VIG = 99999999 and TAB_MULTI_REGIAO.COD_REGIAO = " & Me.Cod_Regiao
            lSelect = lSelect & " order by TAB_REGIAO.cod_regiao"

            'lSelect = "SELECT * FROM Tab_Faixa_Regiao WHERE " & _
             " Cod_Faixa = " & mCod_Regiao & _
             " and Cod_Ramo = " & Cod_Ramo

            If M46V999.gfObtRegistro(BDP00Tarifa, lSelect, rstTab_Regiao, lMensagem) = False Then
                Call M46V999.gpGraLog(1, lMensagem)
                GoTo SAIDA
            End If

            If rstTab_Regiao.EOF Then
                mDesc_Regiao = ""
                mCod_Regiao = Me.Cod_Regiao
            Else
                mDesc_Regiao = rstTab_Regiao("DSC_REGIAO")
                Me.Cod_Produto = rstTab_Regiao("Cod_Produto")
                'mDesc_Regiao = rstTab_Regiao("Dsc_Faixa")
                'mCod_Regiao = Me.Cod_Regiao
            End If
        End If
    End If
    gDesc_Regiao = mDesc_Regiao

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set BDP00Tarifa = Nothing
    Set rstTab_Regiao = Nothing
    Return
End Property

Private Sub Class_Initialize()
    Set objstcPedObs = New stcA46V714B
    Set objstcPedInsp = New stcA46V717B
    Set ColecaoPedClau = New Collection
    Set ColecaoPedCobert = New Collection
    Set ColecaoPedFator = New Collection
    Set ColecaoPedLocBem = New Collection
    Set ColecaoPedLocBicicleta = New Collection
    Set ColecaoPedSeguro = New Collection
    Set ColecaoLMI = New Collection
    Set ColecaoNormas = New Collection
    Set mStcItemApolice = Nothing
    Set mClsM46V999 = Nothing

    Set ColecaoBeneficiario = New Collection
    Set ColecaoRegraBatente = New Collection
    Set ColecaoTelefones = New Collection
    Set ColecaoTextos = New Collection
    Set ColecaoAllRisks = New Collection
    Set ColecaoObrasArte = New Collection
    Set objstcPedInquilino = New stcA46V735B
    Set objstcPedProprietario = New stcA46V735B
End Sub

Private Sub Class_Terminate()
    Set objstcPedObs = Nothing
    Set objstcPedInsp = Nothing
    'Set objstcPedClau = Nothing
    Set ColecaoPedCobert = Nothing
    Set ColecaoPedClau = Nothing
    Set ColecaoPedFator = Nothing
    Set ColecaoPedLocBem = Nothing
    Set ColecaoPedLocBicicleta = Nothing
    Set ColecaoPedSeguro = Nothing
    Set ColecaoLMI = Nothing
    Set ColecaoNormas = Nothing
    Set mStcItemApolice = Nothing
    Set mClsM46V999 = Nothing

    Set ColecaoBeneficiario = New Collection
    Set ColecaoRegraBatente = Nothing
    Set ColecaoAllRisks = Nothing
    Set ColecaoObrasArte = Nothing
    Set objstcPedInquilino = Nothing
    Set objstcPedProprietario = Nothing

End Sub
Public Property Get gDescSegur_1() As Double
    If Cod_Iden_Segur_01_V = "S" Then
        gDescSegur_1 = 8
    Else
        gDescSegur_1 = 0
    End If
End Property
Public Property Get gDescSegur_2() As Double
    If Cod_Iden_Segur_02_A = "S" Then
        gDescSegur_2 = 2
    Else
        gDescSegur_2 = 0
    End If
End Property
Public Property Get gDescSegur_3() As Double
    If Cod_Iden_Segur_03_P = "S" Then
        gDescSegur_3 = 2
    Else
        gDescSegur_3 = 0
    End If
End Property
Public Property Get gDescSegur_4() As Double
    If Cod_Iden_Segur_04_I = "S" Then
        gDescSegur_4 = 2
    Else
        gDescSegur_4 = 0
    End If
End Property
Public Property Get gDescSegur_5() As Double
    If Cod_Iden_Segur_05_G = "S" Then
        gDescSegur_5 = 2
    Else
        gDescSegur_5 = 0
    End If
End Property
Public Property Get gPer_Tot_Segur() As Double
    gPer_Tot_Segur = gDescSegur_1 + gDescSegur_2 + gDescSegur_3 + gDescSegur_4 + gDescSegur_5
End Property
Public Property Get gDescProt_1() As Double
    If Cod_Iden_Protec_01_E = "S" Then
        gDescProt_1 = 5
    Else
        gDescProt_1 = 0
    End If
End Property
Public Property Get gDescProt_2() As Double
    If Cod_Iden_Protec_02_H = "S" Then
        gDescProt_2 = 15
    Else
        gDescProt_2 = 0
    End If
End Property
Public Property Get gDescProt_3() As Double
    If Cod_Iden_Protec_03_S = "S" Then
        gDescProt_3 = 50
    Else
        gDescProt_3 = 0
    End If
End Property
Public Property Get gDescProt_4() As Double
    If Cod_Iden_Protec_04_D = "S" Then
        gDescProt_4 = 10
    Else
        gDescProt_4 = 0
    End If
End Property
Public Property Get gPer_Tot_Protec() As Double
    Dim lTotal As Double
    lTotal = gDescProt_1 + gDescProt_2 + gDescProt_3 + gDescProt_4
    gPer_Tot_Protec = IIf(lTotal > 50, 50, lTotal)
End Property

Public Property Get gCopiaItem(pNum_Item As Long) As stcA46V708B
    Dim lobjstcItem As stcA46V708B
    Dim lobjstcClau As stcA46V703B
    Dim lobjstcCobert As stcA46V704B
    Dim lobjstcBem As stcA46V715B
    Dim lobjstcBicicleta As stcA46V744B
    Dim lobjstcSeguro As stcA46V716B
    Dim lobjstcFator As stcA46V720B
    Dim lobjstcLMI As stcA46V728B
    Dim lbojstcEspec As stcA46V721B

    Dim lobjstcPessoa As stcA46V735B

    Set lobjstcItem = New stcA46V708B
    Set lobjstcItem.M46V999 = M46V999
    Set lobjstcItem.ObjPai = Me.ObjPai

    With lobjstcItem
        .nosso_numero = Me.nosso_numero
        .Num_Item = pNum_Item
        .Num_Ped = Num_Ped
        .Cod_Org = Cod_Org
        .Cod_Ramo = Cod_Ramo
        .Num_Apol = Num_Apol
        .Num_Endo = Num_Endo
        .Cod_Tip_Logr_Risco = Cod_Tip_Logr_Risco
        .Nom_Logr_Risco = Nom_Logr_Risco
        .Dsc_Compl_Risco = Dsc_Compl_Risco
        .Nom_Bairro_Risco = Nom_Bairro_Risco
        .Nom_Cid_Risco = Nom_Cid_Risco
        .Sig_UF_Risco = Sig_UF_Risco
        .Num_Cep_Risco = Num_Cep_Risco
        .Per_Desc_Apar = Per_Desc_Apar
        .Per_Desc_Progres = Per_Desc_Progres
        .Per_IS_Vr = Per_IS_Vr
        .Val_Coef_IS_Vr = Val_Coef_IS_Vr
        .Val_Coef_IS_Vr_LC = Val_Coef_IS_Vr_LC
        .Val_Vr = Val_Vr
        .Tab_Clasf = Tab_Clasf
        .Cod_Clasf = Cod_Clasf
        .Cod_Compl_Clasf = Cod_Compl_Clasf
        .Dsc_Clasf = Dsc_Clasf
        .Cod_Classe_Ocup = Cod_Classe_Ocup
        .Cod_Constr = Cod_Constr
        .Cod_User_Digit = Cod_User_Digit
        .DAT_DIGIT = DAT_DIGIT
        .HOR_DIGIT = HOR_DIGIT
        .Cod_User_Liber = Cod_User_Liber
        .Dat_Liber = Dat_Liber
        .Hor_Liber = Hor_Liber
        .Cod_Sit_item = Cod_Sit_item
        .Per_Bonus = Per_Bonus
        .Qtd_Empreg = Qtd_Empreg
        'Luiz Alberto em 06/11/2015 - Seta "0" para cópia iniciar com opção de inspeção: Conforme Normas
        .Num_Insp = 0 'Num_Insp
        .Cod_Iden_Segur_01_V = Cod_Iden_Segur_01_V
        .Cod_Iden_Segur_02_A = Cod_Iden_Segur_02_A
        .Cod_Iden_Segur_03_P = Cod_Iden_Segur_03_P
        .Cod_Iden_Segur_04_I = Cod_Iden_Segur_04_I
        .Cod_Iden_Segur_05_G = Cod_Iden_Segur_05_G
        .Cod_Iden_Segur_06 = Cod_Iden_Segur_06
        .Cod_Iden_Segur_07 = Cod_Iden_Segur_07
        .Cod_Iden_Segur_08 = Cod_Iden_Segur_08
        .Cod_Iden_Segur_09 = Cod_Iden_Segur_09
        .Cod_Iden_Segur_10 = Cod_Iden_Segur_10
        .Per_Tot_Segur = Per_Tot_Segur
        .Cod_Iden_Protec_01_E = Cod_Iden_Protec_01_E
        .Cod_Iden_Protec_02_H = Cod_Iden_Protec_02_H
        .Cod_Iden_Protec_03_S = Cod_Iden_Protec_03_S
        .Cod_Iden_Protec_04_D = Cod_Iden_Protec_04_D
        .Cod_Iden_Protec_05 = Cod_Iden_Protec_05
        .Cod_Iden_Protec_06 = Cod_Iden_Protec_06
        .Cod_Iden_Protec_07 = Cod_Iden_Protec_07
        .Cod_Iden_Protec_08 = Cod_Iden_Protec_08
        .Cod_Iden_Protec_09 = Cod_Iden_Protec_09
        .Cod_Iden_Protec_10 = Cod_Iden_Protec_10
        .Per_Tot_Protec = Per_Tot_Protec
        .Cod_Produto = Cod_Produto
        .Cod_Regiao = Cod_Regiao
        .Cod_Iden_Tarif = Cod_Iden_Tarif
        .NUM_PI = NUM_PI
        .Num_Item_Ant = Num_Item_Ant
        .Tip_Emissao = Tip_Emissao
        .Num_Apol_Anterior = Num_Apol_Anterior
        .Num_Item_Anterior = Num_Item_Anterior
        .Ind_Sinistro = Ind_Sinistro
        .Possui_Elevador = Possui_Elevador
        .ClausulaNaoDepreciacao = ClausulaNaoDepreciacao 'JIRA-25945 - Opção "Sem Depreciação" no Condomínio
        .BonusAlterado = BonusAlterado

        'Luiz Alberto em 06/11/2015 - Seta "0" para cópia iniciar com opção de inspeção: Conforme Normas
        .Ind_StatusInspecao = 0 'Ind_StatusInspecao
        
        '1000347 - Inclusão de informações da congênere em renovações
        .Ren_Con_Seguradora = Ren_Con_Seguradora
        .Ren_Con_Apolice = Ren_Con_Apolice
        .Ren_Con_Item = Ren_Con_Item

        '1200134 - Inspeção obrigatório por causa de sinistro
        .Ind_Sin_Inspec = Ind_Sin_Inspec
        .Cod_DizerItem = Cod_DizerItem

        .Tip_Cobert = Tip_Cobert
        .Ind_VerbaSep = Ind_VerbaSep
        .Val_Vr_Lc = Val_Vr_Lc
        .Val_LMG_Lc = Val_LMG_Lc
        .Per_IS_Vr_Lc = Per_IS_Vr_Lc
        .Coef_IS_Vr_Lc = Coef_IS_Vr_Lc

        .Tip_Plano_Vida = Tip_Plano_Vida
        .Cod_Plano_Cobertura = Cod_Plano_Cobertura
        .Qtd_Vidas = Qtd_Vidas
        .Val_LMG_Vida = Val_LMG_Vida
        .Ind_Aux_Funeral = Ind_Aux_Funeral
        .Val_LMG_Funeral = Val_LMG_Funeral
        .Ind_Inc_Conjuge = Ind_Inc_Conjuge
        .Val_Aux_Cesta = Val_Aux_Cesta

        .Possui_Elevador = Possui_Elevador
        .ClausulaNaoDepreciacao = ClausulaNaoDepreciacao 'JIRA-25945 - Opção "Sem Depreciação" no Condomínio
        .BonusAlterado = BonusAlterado
        .Txt_Observacoes = Txt_Observacoes
        .Flg_EnviarCertificado = Flg_EnviarCertificado
        .Ind_StatusInspecao = Ind_StatusInspecao
        .Val_Vr_Orig = Val_Vr_Orig
        .Val_Bas_Orig = Val_Bas_Orig
        .Val_Vr_Lc_Orig = Val_Vr_Lc_Orig
        .ITEM_ORIGEM_C4 = ITEM_ORIGEM_C4

        'Set .objstcPedClau = Me.objstcPedClau.gCopiaClau(pNum_Item)
        Set .ColecaoPedClau = New Collection
        For Each lobjstcClau In Me.ColecaoPedClau
            Set lobjstcClau = lobjstcClau.gCopiaClau(pNum_Item)
            Set lobjstcClau.ObjPai = Me
            .ColecaoPedClau.Add lobjstcClau, lobjstcClau.Class_ID
        Next lobjstcClau

        Set .ColecaoPedCobert = New Collection
        For Each lobjstcCobert In Me.ColecaoPedCobert
            Set lobjstcCobert = lobjstcCobert.gCopiaCobert(pNum_Item)
            Set lobjstcCobert.ObjPai = Me
            .ColecaoPedCobert.Add lobjstcCobert, lobjstcCobert.Class_ID
        Next lobjstcCobert

        Set .ColecaoPedFator = New Collection
        For Each lobjstcFator In Me.ColecaoPedFator
            Set lobjstcFator = lobjstcFator.gCopiaFatores(pNum_Item)
            Set lobjstcFator.ObjPai = Me
            .ColecaoPedFator.Add lobjstcFator, lobjstcFator.Class_ID
        Next lobjstcFator


        Set .ColecaoPedLocBem = New Collection
        For Each lobjstcBem In Me.ColecaoPedLocBem
            Set lobjstcBem = lobjstcBem.gCopiaBem(pNum_Item)
            Set lobjstcBem.ObjPai = Me
            .ColecaoPedLocBem.Add lobjstcBem
        Next lobjstcBem
        
'        Copia de item não deve copiar lista de bicicletas
'        Set .ColecaoPedLocBicicleta = New Collection
'        For Each lobjstcBicicleta In Me.ColecaoPedLocBicicleta
'            Set lobjstcBicicleta = lobjstcBicicleta.gCopiaBicicleta(pNum_Item)
'            Set lobjstcBicicleta.ObjPai = Me
'            .ColecaoPedLocBicicleta.Add lobjstcBicicleta
'        Next lobjstcBicicleta

        Set .ColecaoPedSeguro = New Collection
        For Each lobjstcSeguro In Me.ColecaoPedSeguro
            Set lobjstcSeguro = lobjstcSeguro.gCopiaSeguro(pNum_Item)
            Set lobjstcSeguro.ObjPai = Me
            .ColecaoPedSeguro.Add lobjstcSeguro, lobjstcSeguro.Class_ID
        Next lobjstcSeguro

        Set .objstcPedObs = objstcPedObs.gCopiaObs(pNum_Item)

        Set .objstcPedInsp = objstcPedInsp.gCopiaInsp(pNum_Item)

        Set .ColecaoLMI = New Collection
        For Each lobjstcLMI In Me.ColecaoLMI
            Set lobjstcLMI = lobjstcLMI.gCopiaLMI(pNum_Item)
            .ColecaoLMI.Add lobjstcLMI, lobjstcLMI.Class_ID
        Next lobjstcLMI

'        Beneficiario, Inquilino e proprietario não são mais carregados
'        Set .ColecaoBeneficiario = New Collection
'        For Each lobjstcPessoa In Me.ColecaoBeneficiario
'            Set lobjstcPessoa = lobjstcPessoa.gCopiaPessoa(pNum_Item)
'            Set lobjstcPessoa.ObjPai = Me
'            .ColecaoBeneficiario.Add lobjstcPessoa
'        Next lobjstcPessoa
'
'        Set .objstcPedInquilino = objstcPedInquilino.gCopiaPessoa(pNum_Item)
'        Set .objstcPedProprietario = objstcPedProprietario.gCopiaPessoa(pNum_Item)


    End With
    Set gCopiaItem = lobjstcItem

SAIDA:
    GoSub FechaObj
    Exit Property

FechaObj:
    Set lobjstcItem = Nothing
    Set lobjstcClau = Nothing
    Set lobjstcCobert = Nothing
    Set lobjstcBem = Nothing
    Set lobjstcBicicleta = Nothing
    Set lobjstcSeguro = Nothing
    Set lobjstcFator = Nothing
    Set lobjstcLMI = Nothing
    Set lbojstcEspec = Nothing
    Set lobjstcPessoa = Nothing
    Return
End Property
Public Property Get gCEP_Formatado() As String
    Dim lAux As String
    If Num_Cep_Risco <> 0 Then
        lAux = Format(Num_Cep_Risco, "00000000")
        gCEP_Formatado = Mid$(lAux, 1, 5) & "-" & Mid$(lAux, 6, 3)
    Else
        gCEP_Formatado = " "
    End If

End Property

Public Property Get gDesc_TipEmissao() As String
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

    Select Case Tip_Emissao
    Case e_TipEmissao_RenCongenere
        gDesc_TipEmissao = "Renovação de Congênere"
    Case e_TipEmissao_RenYasuda
        gDesc_TipEmissao = "Renovação " & lNomeSeguradora
    Case e_TipEmissao_SeguroNovo
        gDesc_TipEmissao = "Seguro Novo"
    Case e_TipEmissao_Endosso
        gDesc_TipEmissao = "Endosso"
    Case Else
        gDesc_TipEmissao = "não informado"
    End Select
End Property


Public Property Get gNom_Class() As String
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstTabClas As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String

    If Trim(Dsc_Clasf) = "" And Trim(Cod_Clasf) <> "" And Trim(Cod_Clasf) <> "0" Then
        '1. Abre arquivo.
        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
            Call M46V999.gpGraLog(1, lMensagem)
            GoTo SAIDA
        End If
        If Me.Cod_Iden_Tarif > M46V999.TarifaAntiga Then
            lSelect = " Select distinct dsc_clasf, COD_CLASF, " & _
                      "COD_COMPL_CLASF , tab_Clasf from Tbm_Classificacao WITH (NOLOCK) " & _
                    " Where cod_Ramo  = " & Me.Cod_Ramo & _
                    " and Cod_Tarif = " & Me.Cod_Iden_Tarif & _
                    " and cod_Clasf = '" & Me.Cod_Clasf & "'" & _
                    " and Cod_Compl_Clasf = '" & Me.Cod_Compl_Clasf & "'"
        Else
            '3. Obtém produto sem término de vigência ou com término de vigência superior à data do dia.
            lSelect = " Select distinct dsc_clasf, COD_CLASF, " & _
                      "COD_COMPL_CLASF , tab_Clasf from TAB_CLASF WITH (NOLOCK) " & _
                    " Where cod_Ramo  = " & Me.Cod_Ramo & _
                    " and cod_Clasf = '" & Me.Cod_Clasf & "'" & _
                    " and Cod_Compl_Clasf = '" & Me.Cod_Compl_Clasf & "'"
        End If

        If Cod_Plano <> 0 Then
            Select Case Cod_Plano
            Case 1020, 1030, 1040, 1050, 1060
                lSelect = lSelect & " and  tab_clasf = " & Cod_Plano
            Case Else
                If Cod_Produto <> 0 Then
                    lSelect = lSelect & " and  tab_clasf = " & Cod_Produto
                Else
                    lSelect = lSelect & " and  tab_clasf = " & Cod_Plano
                End If
            End Select
        End If

        If M46V999.gfObtRegistro(BDP00Tarifa, lSelect, rstTabClas, lMensagem) = False Then
            Call M46V999.gpGraLog(1, lMensagem)
            GoTo SAIDA
        End If

        If Not rstTabClas.EOF Then
            Dsc_Clasf = rstTabClas("Dsc_Clasf")
        End If
    End If
    gNom_Class = Dsc_Clasf

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set BDP00Tarifa = Nothing
    Set rstTabClas = Nothing
    Return
End Property


Public Property Get gCod_RegiaoForm() As String
    Dim BDP00Tarifa As ADODB.Connection
    Dim rst As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String
    If mCepAnt <> Num_Cep_Risco Or mCod_Regiao = 0 Then
        mCepAnt = Num_Cep_Risco
        '1. Abre arquivo.
        If M46V999.gfAbrBasDados("P00MULTI2", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
            Call M46V999.gpGraLog(1, lMensagem)
            Exit Property
        End If

        '3. Obtém produto sem término de vigência ou com término de vigência superior à data do dia.
        lSelect = "SELECT * FROM TAB_CEP_REGIAO WITH (NOLOCK) WHERE CEP_INI <= " & Num_Cep_Risco & _
                " and CEP_FIM >= " & Num_Cep_Risco

        If M46V999.gfObtRegistro(BDP00Tarifa, lSelect, rst, lMensagem) = False Then
            Call M46V999.gpFecha2(BDP00Tarifa)
            Call M46V999.gpGraLog(1, lMensagem)
            Exit Property
        End If

        If Not rst.EOF Then
            mCod_Regiao = rst("Cod_Regiao")
        End If
    End If
    gCod_RegiaoForm = mCod_Regiao
End Property

Public Property Get gCod_RegiaoNova() As String
    Dim BDP00Tarifa As ADODB.Connection
    Dim rst As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String

    If mCepAnt <> Num_Cep_Risco Or mCod_Regiao = 0 Then
        mCepAnt = Num_Cep_Risco
        '1. Abre arquivo.
        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
            Call M46V999.gpGraLog(1, lMensagem)
            Exit Property
        End If

        '3. Obtém produto sem término de vigência ou com término de vigência superior à data do dia.
        lSelect = "SELECT * FROM tbm_faixa_regiao WITH (NOLOCK) WHERE CEP_INI <= " & Num_Cep_Risco & _
                " and CEP_FIM >= " & Num_Cep_Risco & " and cod_tarif =  " & M46V999.gCodProduto & _
                " and cod_modalidade = " & Me.Cod_Produto & " and cod_ramo = " & Me.Cod_Ramo

        If M46V999.gfObtRegistro(BDP00Tarifa, lSelect, rst, lMensagem) = False Then
            Call M46V999.gpFecha2(BDP00Tarifa)
            Call M46V999.gpGraLog(1, lMensagem)
            Exit Property
        End If

        If Not rst.EOF Then
            mCod_Regiao = rst("Cod_Reg")
        End If
    End If
    gCod_RegiaoNova = mCod_Regiao
End Property

Public Property Set ObjPai(pobjstc As stcA46V702B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get ObjPai() As stcA46V702B
    Set ObjPai = ObjStcPai
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property

Public Property Get ValPrTotCober() As Double
    Dim lretorno As Double
    Dim lobjstcCob As stcA46V704B

    For Each lobjstcCob In ColecaoPedCobert
        lretorno = lretorno + lobjstcCob.VAL_PR
    Next lobjstcCob
    If Me.ObjPai Is Nothing Then
    Else
        'Ajute retirar premio minimo do imobiliario
        If Me.ObjPai.ColecaoPedLoc.Count = 1 And M46V111.gCod_plano <> 20 Then
            If Me.Tip_Emissao = e_TipEmissao_Endosso Then
                If lretorno > 0 Then
                    If lretorno < M46V999.gProPreMinEndosso And M46V999.gPermitePrInformado = False Then
                        If Cod_Iden_Tarif >= M46V999.TarifaPrecificacaoMaritima Then
                            If lretorno > CDbl(8) Then
                                lretorno = M46V999.gProPreMinEndosso
                            Else
                                lretorno = 0
                            End If
                        Else
                            lretorno = 0
                        End If
                    End If
                ElseIf lretorno < 0 Then
                    If lretorno > (M46V999.gProPreMinEndosso * -1) And Cod_Iden_Tarif <= 1076 Then
                        lretorno = 0
                    End If
                End If
            Else
                If lretorno > 0 Then
                    If lretorno < M46V999.gProPremioMinimo Then
                        lretorno = M46V999.gProPremioMinimo
                    End If
                ElseIf lretorno < 0 Then
                    If lretorno > (M46V999.gProPremioMinimo * -1) Then
                        lretorno = 0
                    End If
                End If
            End If
        End If
    End If
    Set lobjstcCob = Nothing
    ValPrTotCober = lretorno
End Property

Public Property Get perDescComlUltCob() As Double
    Dim lretorno As Double
    Dim lobjstcCob As stcA46V704B

    If ColecaoPedCobert.Count > 0 Then
        'Set lobjstcCob = ColecaoPedCobert(ColecaoPedCobert.Count)
        Set lobjstcCob = ColecaoPedCobert(1)
        lretorno = lobjstcCob.Per_Desc_Coml
    End If
    Set lobjstcCob = Nothing
    perDescComlUltCob = lretorno

End Property

Public Property Get gDescPlano() As String
'PáFazer
    gDescPlano = "ToDo"
End Property

Public Property Get gTemRegProfissoes() As Boolean
'Ficha 1000681
    Static lsRetorno As Boolean
    Static lsCod_Ramo As Integer
    Static lsCod_Plano As Integer
    Static lsCod_Clasf As String
    Static lsCod_Compl_Clasf As String
    Dim lSelect As String
    Dim lbdTAB_RegCons As ADODB.Connection
    Dim lrsTAB_RegCons As clsYasRecordSet    'Registro - TBM_REG_PROFISSOES.
    Dim lMensagem As String

    If (lsCod_Ramo <> Me.Cod_Ramo) And _
       (lsCod_Plano <> Me.Cod_Plano) And _
       (lsCod_Clasf <> Me.Cod_Clasf) And _
       (lsCod_Compl_Clasf <> Me.Cod_Compl_Clasf) Then

        lsCod_Ramo = Me.Cod_Ramo
        lsCod_Plano = Me.Cod_Plano
        lsCod_Clasf = Me.Cod_Clasf
        lsCod_Compl_Clasf = Me.Cod_Compl_Clasf

        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTAB_RegCons, lMensagem) = False Then
            Call M46V999.gpGraLog(0, lMensagem)
            GoTo Retorno
        End If

        lSelect = "SELECT DISTINCT a.Cod_Reg, a.Dsc_Reg FROM TBM_REG_PROFISSOES A WITH (NOLOCK), TBM_CLASF_PROFISSOES B WITH (NOLOCK) "
        lSelect = lSelect & " WHERE a.Cod_Reg = b.Cod_Reg"
        lSelect = lSelect & " and b.Cod_Ramo = " & lsCod_Ramo
        lSelect = lSelect & " and b.Cod_Modalidade = " & lsCod_Plano
        lSelect = lSelect & " and (b.Cod_Clasf = '" & Trim(lsCod_Clasf) & "'"   '1000871 - Inclusão de Orgão Categoria Conselho
        lSelect = lSelect & " OR b.Cod_Clasf = '" & val(lsCod_Clasf) & "')"   '1000871 - Inclusão de Orgão Categoria Conselho
        lSelect = lSelect & " and b.Cod_Compl_Clasf = '" & Trim(lsCod_Compl_Clasf) & "'"   '1000871 - Inclusão de Orgão Categoria Conselho
        lSelect = lSelect & " order by a.Cod_Reg"

        If M46V999.gfObtRegistro(lbdTAB_RegCons, lSelect, lrsTAB_RegCons, lMensagem) = False Then
            Call M46V999.gpFecha2(lbdTAB_RegCons)
            Call M46V999.gpGraLog(0, lMensagem)
            GoTo Retorno
        End If

        If Not lrsTAB_RegCons.EOF Then
            lsRetorno = True
        Else
            lsRetorno = False
        End If
        Call M46V999.gpFecha2(lbdTAB_RegCons)
    End If

Retorno:
    gTemRegProfissoes = lsRetorno

End Property

Public Function getRpt() As stcA46V708rpt
    Static lrpt As stcA46V708rpt

    If lrpt Is Nothing Then
        Set lrpt = New stcA46V708rpt
        Set lrpt.stcOrigem = Me
    End If

    Set getRpt = lrpt

End Function

Public Property Get gNormaInspecao() As Boolean
    Dim lstcNorma As stcA46V708N
    Dim lPrefix As Integer
    Dim lretorno As Boolean

    lretorno = False
    For Each lstcNorma In ColecaoNormas
        lPrefix = CInt(Left(CStr(lstcNorma.COD_STAT_IRREG), 2))
        If lPrefix = 55 Then
            lretorno = True
            Exit For
        End If
    Next
    Set lstcNorma = Nothing
    gNormaInspecao = lretorno

End Property

Public Property Get gISBasica() As Double
    Dim lretorno As Double
    Dim i As Integer
    lretorno = 0
    If Me.ColecaoPedCobert.Count > 0 Then
        For i = 1 To Me.ColecaoPedCobert.Count
            If ColecaoPedCobert(i).Cod_Cobert = 1 Then
                lretorno = lretorno + Me.ColecaoPedCobert(i).Val_IS
            End If
            If ColecaoPedCobert(i).gCod_Grupo_Cobert = 5 Then
                lretorno = lretorno + Me.ColecaoPedCobert(i).Val_IS
            End If
        Next
    End If
    gISBasica = lretorno
End Property


Public Property Get gTip_Atividade() As Integer
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstTabClas As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String

    Static lsCodRamo As Integer
    Static lsCodClasf As String
    Static lsCodComplClasf As String
    Static lsCodProduto As Integer
    Static lsCodSequencia As Integer
    Static lsCodPlano As Double

    Static lsRetorno As Integer

    If (lsCodRamo <> Me.Cod_Ramo) Or _
       (lsCodClasf <> Me.Cod_Clasf) Or _
       (lsCodComplClasf <> Me.Cod_Compl_Clasf) Or _
       (lsCodProduto <> Me.ObjPai.Cod_Produto) Or _
       (lsCodSequencia <> Me.ObjPai.Cod_Sequencia) Or _
       (lsCodPlano <> Me.ObjPai.Cod_Estip) Then

        '==== Salva os parametros nas variaveis staticas p/ otimizar desempenho!
        lsCodRamo = Me.Cod_Ramo
        lsCodClasf = Me.Cod_Clasf
        lsCodComplClasf = Me.Cod_Compl_Clasf
        lsCodProduto = Me.ObjPai.Cod_Produto
        lsCodSequencia = Me.ObjPai.Cod_Sequencia
        lsCodPlano = Me.ObjPai.Cod_Estip
        lsRetorno = 0
        '====

        If Trim(Me.Cod_Clasf) <> "" And Trim(Me.Cod_Clasf) <> "0" Then
            '1. Abre arquivo.
            If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
                Call M46V999.gpGraLog(1, lMensagem)
                Exit Property
            End If

            '3. Obtém produto sem término de vigência ou com término de vigência superior à data do dia.
            lSelect = " Select distinct dsc_clasf, COD_CLASF, " & _
                      "COD_COMPL_CLASF , tab_Clasf, Tip_Atividade  from Tbm_Classificacao WITH (NOLOCK) " & _
                    " Where cod_Ramo  = " & Me.Cod_Ramo & _
                    " and cod_Clasf = '" & Me.Cod_Clasf & "'" & _
                    " and Cod_Compl_Clasf = '" & Me.Cod_Compl_Clasf & "'" & _
                    " and Cod_Tarif = " & Me.ObjPai.Cod_Produto & _
                    " and Cod_Sequencia <= " & Me.ObjPai.Cod_Sequencia
            lSelect = lSelect & " and  tab_clasf = " & Me.ObjPai.Cod_Estip

            If M46V999.gfObtRegistro(BDP00Tarifa, lSelect, rstTabClas, lMensagem) = False Then
                Call M46V999.gpFecha2(BDP00Tarifa)
                Call M46V999.gpGraLog(1, lMensagem)
            Else
                If Not rstTabClas.EOF Then
                    lsRetorno = M46V999.gfTratarNulo(rstTabClas("Tip_Atividade"), tpoNumerico)
                End If
            End If
        End If
    End If
    gTip_Atividade = lsRetorno

End Property

Public Property Get gTemProtecaoObrigatoria() As String
    Dim BDP00Tarifa As ADODB.Connection
    Dim lrsTbm_Parameter As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String
    Dim lobjstcFator As stcA46V720B

    Static lsCodRamo2 As Integer
    Static lsCodClasf2 As String
    Static lsCodComplClasf2 As String
    Static lsCodProduto2 As Integer
    Static lsCodSequencia2 As Integer
    Static lsCodPlano2 As Double

    Static lsRetorno2 As String

    'e_Tip_Valor_FatorProtecao = 22               'Carrega automaticamente um Fator de Protecao
    'e_Cod_Tip_Chave_Fator_Ajuste = 10            'Desconto / Fator de Proteção


    If ((lsCodRamo2 <> Me.Cod_Ramo) Or _
        (lsCodClasf2 <> Me.Cod_Clasf) Or _
        (lsCodComplClasf2 <> Me.Cod_Compl_Clasf) Or _
        (lsCodProduto2 <> Me.ObjPai.Cod_Produto) Or _
        (lsCodSequencia2 <> Me.ObjPai.Cod_Sequencia) Or _
        (lsCodPlano2 <> Me.ObjPai.Cod_Estip)) Or lsRetorno2 = "" Then

        '==== Salva os parametros nas variaveis staticas p/ otimizar desempenho!
        lsCodRamo2 = Me.Cod_Ramo
        lsCodClasf2 = Me.Cod_Clasf
        lsCodComplClasf2 = Me.Cod_Compl_Clasf
        lsCodProduto2 = Me.ObjPai.Cod_Produto
        lsCodSequencia2 = Me.ObjPai.Cod_Sequencia
        lsCodPlano2 = Me.ObjPai.Cod_Estip
        lsRetorno2 = ""
        '====

        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
            Call M46V999.gpGraLog(1, lMensagem)
            Exit Property
        End If


        lSelect = " Select Cod_Cobert, Cod_Tip_Chave, Chave_Ini, Chave_Ter, Valor from Tbm_Parameter WITH (NOLOCK) "
        lSelect = lSelect & " Where Cod_Ramo = " & Me.Cod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & Me.ObjPai.Cod_Estip
        lSelect = lSelect & " and Cod_Corr in (0, " & Me.ObjPai.objstcCtrlEmis.Cod_Corr & ")"
        lSelect = lSelect & " and ((rtrim(ltrim(Cod_Clasf)) = '" & Me.Cod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & Me.Cod_Compl_Clasf & "') or rtrim(ltrim(Cod_Clasf)) = '0' )"
        lSelect = lSelect & " and Cod_Cobert = 0 "
        lSelect = lSelect & " and Cod_Tarif = " & Me.ObjPai.Cod_Produto
        lSelect = lSelect & " and Cod_Sequencia <= " & Me.ObjPai.Cod_Sequencia
        lSelect = lSelect & " and Cod_Tip_Valor = " & 22
        lSelect = lSelect & " and Cod_Tip_Chave = " & 10
        lSelect = lSelect & " and Cod_Constr in (0," & Me.Cod_Constr & ")"
        lSelect = lSelect & " order by Cod_Cobert, Cod_Prioridade, Cod_Constr desc "

        If M46V999.gfObtRegistro(BDP00Tarifa, lSelect, lrsTbm_Parameter, lMensagem) = False Then
            Call M46V999.gpFecha2(BDP00Tarifa)
            Call M46V999.gpGraLog(1, lMensagem)
            Exit Function
        End If

        '1201062 - Tarifa 1048 e 1049 - Item 1
        mLiberaDescontoEspecial = False
        Do While Not lrsTbm_Parameter.EOF
            lsRetorno2 = lsRetorno2 & M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), 4) & "@" & M46V999.gfTratarNulo(lrsTbm_Parameter("Chave_Ini"), 5) & ";"
            '1201062 - Tarifa 1048 e 1049 - Item 1
            If Me.Cod_Ramo = e_CodRamo_Residencial And M46V999.gfTratarNulo(lrsTbm_Parameter("Valor"), 4) = 7 Then mLiberaDescontoEspecial = True
            lrsTbm_Parameter.MoveNext
        Loop

        For Each lobjstcFator In ColecaoPedFator
            If lobjstcFator.Cod_Ramo = e_CodRamo_Residencial And _
               lobjstcFator.Cod_Fator = 7 And M46V999.gTipoAcesso = DIGITACAO Then
                mLiberaDescontoEspecial = True
            End If
        Next lobjstcFator

    End If
    'Retirado por Simony - o sistema está preparado para preenchido e não preenchido.
    'If lsRetorno2 = "" Then lsRetorno2 = False
    gTemProtecaoObrigatoria = lsRetorno2

End Property


'1101015 - Projeto 11010 - Renovação Expressa
Public Property Get Apol_Ant_PER_DESC_COML() As Double
    Dim lRdsApol As clsYasRecordSet
    Dim bdP00APOL As ADODB.Connection
    Dim lSelect As String
    Dim lNomeTab As String
    Dim pMensagem As String

    If lApol_Ant_PER_DESC_COML > 0 Then GoTo Fim

    If M46V999.gPrograma = "D046O003" Then
        lNomeTab = "Syasportal.dbo.TAB_APOL"
    Else
        'If M46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
        'Else
        '    lNomeTab = "GED.DBO.TAB_PED"
        'End If
    End If

    lSelect = " SELECT " & lNomeTab & "_COBERT.PER_DESC_COML"
    lSelect = lSelect & " FROM " & lNomeTab & " WITH (NOLOCK)"
    lSelect = lSelect & " INNER JOIN " & lNomeTab & "_COBERT  WITH (NOLOCK)"
    lSelect = lSelect & " ON (" & lNomeTab & ".COD_RAMO = " & lNomeTab & "_COBERT.COD_RAMO) "
    lSelect = lSelect & " AND (" & lNomeTab & ".COD_ORG = " & lNomeTab & "_COBERT.COD_ORG) "
    lSelect = lSelect & " AND (" & lNomeTab & ".NUM_APOL = " & lNomeTab & "_COBERT.NUM_APOL) "
    lSelect = lSelect & " WHERE (((" & lNomeTab & ".NUM_APOL)=" & Me.ObjPai.Num_Apol_An & ") "
    lSelect = lSelect & " AND ((" & lNomeTab & ".COD_RAMO)=" & Me.Cod_Ramo & ") "
    lSelect = lSelect & " AND ((" & lNomeTab & ".COD_ORG)=" & Me.Cod_Org & ") "
    lSelect = lSelect & " AND ((" & lNomeTab & "_COBERT.NUM_ITEM)=" & Me.Num_Item & ")) "

    'b) Abre arquivo.
    If M46V999.gfAbrBasDados("P00APOL", M46V999.gTipP00APOL, M46V999.gEndP00APOL, bdP00APOL, pMensagem) = False Then
        GoTo SAIDA
    End If
    'c) Obtém registro.
    If M46V999.gfObtRegistro(bdP00APOL, lSelect, lRdsApol, pMensagem) = False Then
        Call M46V999.gpFecha2(bdP00APOL)
        GoTo SAIDA
    End If

    'Retorna o maior desconto
    Do Until lRdsApol.EOF
        If M46V999.gfTratarNulo(lRdsApol("PER_DESC_COML"), tpoNumerico) > lApol_Ant_PER_DESC_COML Then
            lApol_Ant_PER_DESC_COML = M46V999.gfTratarNulo(lRdsApol("PER_DESC_COML"), tpoNumerico)
        End If
        lRdsApol.MoveNext
    Loop


Fim:
    Apol_Ant_PER_DESC_COML = lApol_Ant_PER_DESC_COML

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Call M46V999.gpFecha2(bdP00APOL)
    Set bdP00APOL = Nothing
    Return
End Property

'1101015 - Projeto 11010 - Renovação Expressa
Public Property Get Apol_Ant_IND_SINISTRO() As Double
    Dim lRdsApol As clsYasRecordSet
    Dim bdP00APOL As ADODB.Connection
    Dim lSelect As String
    Dim lNomeTab As String
    Dim pMensagem As String
    Dim lSinistro As Long

    'If lApol_Ant_IND_SINISTRO > 0 Then GoTo Fim

    If M46V999.gPrograma = "D046O003" Then
        lNomeTab = "Syasportal.dbo.TAB_APOL"
    Else
        'If M46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
        'Else
        '    lNomeTab = "GED.DBO.TAB_PED"
        'End If
    End If
'   Comentado pois a informação correta de sinistro agora esta sendo carregada na TAB_APOL_LOC
'    If lNomeTab = "TAB_APOL" And M46V999.gBaseLocal = False Then
'        'lSelect = " select count(*) as IND_SINISTRO from ramosdiversos.dbo.tb_aviso_sin_diversos where num_apol = " & Me.Num_Apol_Anterior
'        'lSelect = lSelect & " and num_item = " & Me.Num_Item_Anterior
'        lSelect = "  select count(a.num_apol) as IND_SINISTRO  from RamosDiversos..tb_aviso_sin_diversos AS a WITH (NOLOCK)"
'        lSelect = lSelect & " inner join Ged..TB_Valores_Socket as V WITH (NOLOCK)"
'        lSelect = lSelect & " on a.num_apol = v.num_apol and  a.num_item  = v.num_item"
'        lSelect = lSelect & " where tip_pagto in (211,212,213) and dat_contab > 0 and  iden_estor = 0 and "
'        lSelect = lSelect & " V.num_apol = " & Me.Num_Apol_Anterior & " And v.Num_Item = " & Me.Num_Item_Anterior
'    Else
'''2017-07-31: Juliano - SinistroSimNao
''        lSelect = " SELECT IND_SINISTRO "
''        lSelect = lSelect & " FROM " & lNomeTab & "_LOC  WITH (NOLOCK)"
''        lSelect = lSelect & " WHERE NUM_APOL = " & Me.Num_Apol_Anterior
''        lSelect = lSelect & " AND   NUM_ITEM = " & Me.Num_Item_Anterior
''        lSelect = lSelect & " AND   COD_ORG = " & Me.Cod_Org
''        lSelect = lSelect & " AND   COD_RAMO = " & Me.Cod_Ramo
    lSelect = "exec Ged..proc_Log_SinEmissao_Consultar @Num_PI=" & Me.NUM_PI & ", @Num_Apol=" & Me.Num_Apol_Anterior & ", @Num_Item=" & Me.Num_Item_Anterior & ", @Cod_Org=" & Me.Cod_Org & ", @Cod_Ramo=" & Me.Cod_Ramo
'    End If

    'b) Abre arquivo.
    If M46V999.gfAbrBasDados("P00APOL", M46V999.gTipP00APOL, M46V999.gEndP00APOL, bdP00APOL, pMensagem) = False Then
        GoTo SAIDA
    End If
    'c) Obtém registro.
    If M46V999.gfObtRegistro(bdP00APOL, lSelect, lRdsApol, pMensagem) = False Then
        Call M46V999.gpFecha2(bdP00APOL)
        GoTo SAIDA
    End If

    If Not lRdsApol.EOF Then
        lApol_Ant_IND_SINISTRO = M46V999.gfTratarNulo(lRdsApol("IND_SINISTRO"), tpoNumerico)
        If lApol_Ant_IND_SINISTRO > 1 Then lApol_Ant_IND_SINISTRO = 1
        Ind_Sinistro_ORIG = M46V999.gfTratarNulo(lRdsApol("IND_SINISTRO_ORIG"), tpoNumerico)
    End If

    Call M46V999.gpFecha2(bdP00APOL)

Fim:
    Apol_Ant_IND_SINISTRO = lApol_Ant_IND_SINISTRO

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Call M46V999.gpFecha2(bdP00APOL)
    Set bdP00APOL = Nothing
    Return
End Property


Public Function getPlanoItem() As stcPlanoItem
    Static lstcPlanoItem As stcPlanoItem

    If lstcPlanoItem Is Nothing Then
        Set lstcPlanoItem = New stcPlanoItem
        Set lstcPlanoItem.stcOrigem = Me
    End If

    Set getPlanoItem = lstcPlanoItem

End Function


Public Property Get FindFatorProtecao(pobjstc As stcA46V720B) As Boolean
    Dim lobjstc As stcA46V720B

    On Local Error GoTo ControleErros
    FindFatorProtecao = False

    For Each lobjstc In ColecaoPedFator
        If lobjstc.Cod_Fator = pobjstc.Cod_Fator _
           And lobjstc.Cod_Org = pobjstc.Cod_Org _
           And lobjstc.Cod_Ramo = pobjstc.Cod_Ramo _
           And lobjstc.Cod_Unicidade = pobjstc.Cod_Unicidade Then
            FindFatorProtecao = True
            Exit For
        End If
    Next

SAIDA:
    GoSub FechaObj
    Exit Property
ControleErros:
    FindFatorProtecao = False
    'Call gpGraLog(0, "stcA46V708B.FindFatorProtecao ERRO - " & Err & " " & Err.Description)
    GoTo SAIDA
FechaObj:
    Set lobjstc = Nothing
    Return
End Property

Public Function gCodClauAss24h() As Long
    Dim lstcPedClau As stcA46V703B
    Dim lsRetorno As Long

    'CR93893 Comentado por problema ao recalcular franquia padão ao mudar classificação da cotação
    If M46V999.M46V111.gCod_plano <> Me.Cod_Plano And Me.Cod_Plano > 0 And M46V999.M46V111.gCod_plano = 0 Then
        M46V999.M46V111.gCod_plano = Me.Cod_Plano
    End If
    
    For Each lstcPedClau In Me.ColecaoPedClau
        Set lstcPedClau.M46V999 = mClsM46V999
        
        If lstcPedClau.gEhAss24h Then
            lsRetorno = lstcPedClau.Cod_Clau
        End If
    Next lstcPedClau

    gCodClauAss24h = lsRetorno

End Function

Private Property Get gStcAss24hs() As stcA46V734B
    Dim lColAss24H As Collection
    Dim lStcAss24h As stcA46V734B
    Dim lCodClauAss24h As Long

    Set lStcAss24h = New stcA46V734B
    Set lStcAss24h.M46V999 = M46V999

    lCodClauAss24h = gCodClauAss24h
    If lCodClauAss24h > 0 Then
        Set lColAss24H = M46V111.gfCarregaAss24h(Me.ObjPai.Cod_Produto, _
                                                 Me.ObjPai.Cod_Sequencia, _
                                                 Me.ObjPai.Cod_Ramo, _
                                                 Me.Cod_Plano)
        For Each lStcAss24h In lColAss24H
            If lStcAss24h.Cod_Clau = lCodClauAss24h Then
                Exit For
            End If
        Next lStcAss24h

        If lStcAss24h Is Nothing Then
            Set lStcAss24h = New stcA46V734B
        End If
    End If

    Dim mPer_ComlAss24hrs As Double
    If Me.ObjPai.Cod_Estip = 20 Then
        mPer_ComlAss24hrs = 0
        If M46V999.M46V103.mfAjustesPontuais(e_Tip_Valor_ComissaoAss24hrs, tpoNumerico, mPer_ComlAss24hrs, "", Me, 0, "") = False Then
            Call M46V999.gpGraLog(0, "")
            GoTo SAIDA
        End If
        
        lStcAss24h.Premio = lStcAss24h.Premio + ((lStcAss24h.Premio / 100) * mPer_ComlAss24hrs)
    End If

    Set gStcAss24hs = lStcAss24h

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lStcAss24h = Nothing
    Return
End Property

Public Property Get gVal_PR_Ass24h() As Double
    Dim lQtdias As Integer
    
    lQtdias = M46V999.gfDifDatas(ObjPai.Dat_Ini_Vig, ObjPai.Dat_Ter_Vig)
    If lQtdias < 365 And ObjPai.Cod_Estip = 20 Then
        gVal_PR_Ass24h = lQtdias * Round(gStcAss24hs.Premio / 365, 2)
    Else
        gVal_PR_Ass24h = gStcAss24hs.Premio
    End If
    'Guarda o valor de ass24hr
    Val_Ass24h = gVal_PR_Ass24h
End Property

Public Property Get gDescricao_Ass24h() As String
    gDescricao_Ass24h = gStcAss24hs.Descricao
End Property

Public Property Get gStcItemApolice() As stcA46V708B
    Dim lStcItemApolice As stcA46V708B
    Dim lStcPedidoApolice As stcA46V702B

    If mStcItemApolice Is Nothing Then
        Set lStcPedidoApolice = Me.ObjPai.gStcPedidoApolice
        If Not lStcPedidoApolice Is Nothing Then
            For Each lStcItemApolice In lStcPedidoApolice.ColecaoPedLoc
                If Me.Num_Item = lStcItemApolice.Num_Item Then
                    Set mStcItemApolice = lStcItemApolice
                End If
            Next lStcItemApolice
        End If
    End If
    If mStcItemApolice Is Nothing Then Set mStcItemApolice = New stcA46V708B

    'Retorno
    Set mStcItemApolice.M46V999 = M46V999
    Set gStcItemApolice = mStcItemApolice
SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lStcPedidoApolice = Nothing
    Set lStcItemApolice = Nothing
    Return
End Property

Public Property Get gPodeOptarAss24h() As Boolean
    Dim lretorno As Boolean

    lretorno = True

    '*** Regras de restrição
    If (Me.Tip_Emissao = e_TipEmissao_Endosso And Me.Cod_DizerItem = e_CodDizerAlteracao And _
        Me.gStcItemApolice.gCodClauAss24h = 0) Then
        'Nao possui Ass24h na apólice anterior.
        lretorno = False
    ElseIf (Me.Cod_Ramo = e_CodRamo_Residencial And val(Me.Cod_Clasf) = val(M46V999.gfBuscarDePara(1, 2))) And _
           (val(Me.Cod_Compl_Clasf) = val(M46V999.gfBuscarDePara(1, 2, 23)) Or val(Me.Cod_Compl_Clasf) = val(M46V999.gfBuscarDePara(1, 2, 23))) Then
        'Não contemplarão imóveis desocupados. Classificações: 001-23(Apart. Desocupado) / 001-13(Casa Desocupada).
        lretorno = False
    End If
    gPodeOptarAss24h = lretorno

End Property

Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
End Property

Public Property Get M46V999() As clsM46V999
    If mClsM46V999 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V999 não atribuído para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V999 já
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46V999 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46V999 = mClsM46V999
End Property
Public Property Get M46V111() As clsM46V111
    Set M46V111 = M46V999.M46V111
End Property

'1201062 - Tarifa 1048 e 1049 - Item 1
Public Property Get gLiberaDescontoEspecial() As Boolean
    Dim FatoresObrigatorios As String
    FatoresObrigatorios = gTemProtecaoObrigatoria
    gLiberaDescontoEspecial = mLiberaDescontoEspecial
End Property

Public Function FindCobertura(lCodCobert As Integer) As Boolean
    Dim mObjstc As stcA46V704B
    On Local Error GoTo ControleErros
    FindCobertura = False

    For Each mObjstc In ColecaoPedCobert
        If mObjstc.Cod_Cobert = lCodCobert Then
            FindCobertura = True
            Exit For
        End If
    Next

SAIDA:
    GoSub FechaObj
    Exit Function
ControleErros:
    FindCobertura = False
    Call M46V999.gpGraLog(0, "stcA46V708B.FindCobertura ERRO - " & Err & " " & Err.Description)
    GoTo SAIDA
FechaObj:
    Set mObjstc = Nothing
    Return
End Function
