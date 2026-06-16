VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V704B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: stcA46V704B --> Tabela Tab_Ped_Cobert
' Data/Hora: 31/08/2005 - 11:52:48
' Sistema: Multi
' Thais Juren Gomes

' ======================================================================================================================================================


' ======================================================================================================================================================
Public nosso_numero As String          'Nosso número.
Public Cod_Cobert As Integer           'Código da cobertura.
Public Num_Ped As Long                 'Número do pedido.
Public Cod_Org As Integer              'Código do órgão.
Public Cod_Ramo As e_CodRamo           'Código do ramo.
Public Num_Apol As Double                'Sem uso - número da apólice.
Public Num_Endo As Long                'Sem uso - número do endosso.
Public Num_Item As Long                'Número do item.
Public Tab_Cobert As Integer
Public Val_IS As Double                'Valor da importância segurada.
Public Per_Taxa As Double              'Percentual da taxa.
Public Per_Desc_Coml As Double         'Percentual de desconto comercial.
Public Per_Desc_Exp As Double          'Percentual de desconto de experiência.
Public Val_Pr_Anu As Double            'Prêmio anual.
Public Val_Pr_Tarif As Double
Public Val_Pr_Tarif_Anu As Double
Public Val_Pr_Resg As Double
Public Val_Pr_Resg_Anu As Double
Public Per_Franq As Double              'Percentual da franquia.
Public Val_Franq_Min As Double          'Valor da franquia mínima.
Public Val_Franq_Max As Double          'Valor da franquia máxima.
'Public NUM_PI As Long                   'Número do PI.
Public NUM_PI As Double                   'Número do PI.
Public Num_Meses_PI As Integer          'Número de Meses do Período Indenitário.
Public Formula As String                'Formula de Cálculo.
Public COD_CLASSE As Integer            'Código da Classe
Public Val_Fator_Franq As Double        'Fator de Ajuste da Franquia
Public Alterar_Cobert As String         'Permite Alterar Cobertura ou Taxa já alterada não recalcular
Public Val_Risco_Cob As Double          'Valor em Risco da Cobertura
Public Nom_Cobertura As String          'Nome da cobertura.
Public Cod_Grupo_Cobertura As Integer   'Código do Grupo das Coberturas .
'  1 - Basica
'  2 - Acessórias
'  3 - Adicionais
'  4 - Vida
'  5 - Basica Verba Separadas
Public Val_Prem_Min As Double               'Premio Minimo
Public AlterouFranquia As String
Public AjusteEndosso As String
Public EndossoTipoRecalculo As Integer
Public EndossoValPremioApol As Double
Public TextoFranquia As String
Public VAL_PR_BATENTE As Double       'Valor do premio atual sem o calculo do batente(Só recebe valor quando a apólice sofre calculo de batente)
Public Ind_Agrup_Cobert As Boolean

Public Compl_Desc_Cobertura As String
Public Per_Desc_Franq As Double

Public ColecaoNovaFranquia As Collection
Public Val_Lim_Max_Cobert As Double
Public Val_Lim_Min_Cobert As Double
Public EhGrandesRiscos As Boolean
Public RealizaCalBatente As Boolean
Public Log_Batente As stcA46V745B
Public Per_Desc_Coml_Web As Double

Private PER_LIM_MAX_COBERT As Double
Private COD_LIM_COBERT As Integer
Private Nom_Franquia As String
Private Tem_VideObs As String

''2016-10-19~2018-02-28: VerbaCom
Public Val_VerbaComProp As Double

' ======================================================================================================================================================
Private ObjStcPai As stcA46V708B
' ======================================================================================================================================================

Private mClsM46V999 As clsM46V999
Private mRpt As stcA46V704rpt
Private VAL_PR_ As Double                'Prêmio.
Private PremioInformado_ As Double
Private Val_Pr_Orig_ As Double

Private Val_IS_Orig_ As Double
Private Per_Franq_Orig_ As Double

Private Tip_Origem_C4_ As Integer
Private RA_C4_ As Integer
Private Val_IS_RA_ As Double

Public Property Get VAL_PR() As Double
    VAL_PR = VAL_PR_
End Property

Public Property Let VAL_PR(ByVal vNewValue As Double)
    'Debug.Assert Cod_Cobert <> 1103
    VAL_PR_ = vNewValue
End Property

Public Property Get PremioInformado() As Double
    PremioInformado = PremioInformado_
End Property

Public Property Let PremioInformado(ByVal vNewValue As Double)
    'Debug.Assert Cod_Cobert <> 1103
    PremioInformado_ = vNewValue
End Property

Public Property Let Val_Pr_Orig(ByVal vNewValue As Double)
    'Debug.Assert Cod_Cobert <> 1103
    Val_Pr_Orig_ = vNewValue
End Property

Public Property Get Val_Pr_Orig() As Double
    Val_Pr_Orig = Val_Pr_Orig_
End Property

Sub AtualizaLimFranquia(ByRef pPerFranqDefault As Double, ByRef pValFraqMinDefault As Double, ByRef pFatorFranqDefault As Double, ByRef pDescricaoFranquia As String, ByRef pValFraqMaxDefault As Double, ByRef pTem_VideObs As String, pNum_Meses_PI As Integer)

    Dim Dummy1 As String
    Dim Dummy2 As String

    '====Parametros
    Static lsCod_Sequencia As Integer
    Static lsCod_Ramo As e_CodRamo
    Static lsCodPlano As Double
    Static lsCodProduto As Integer
    Static lsCodClasf As String
    Static lsCodComplClasf As String
    Static lsCod_Cobert As Integer
    Static lsCod_Constr As Integer
    Static lsNum_Cep_Risco As Long
    Static lsTip_Atividade As Integer
    Static lsCod_Corr As Long
    '====Retorno
    Static lsPerFranqDefault As Double
    Static lsValFraqMinDefault As Double
    Static lsFatorFranqDefault As Double
    Static lsDescricaoFranquia As String
    Static lsValFraqMaxDefault As Double
    Static lsTem_VideObs As String

    If (lsCod_Sequencia <> Me.ObjPai.ObjPai.Cod_Sequencia) Or _
       (lsCod_Ramo <> Me.Cod_Ramo) Or _
       (lsCodPlano <> Me.ObjPai.ObjPai.Cod_Estip) Or _
       (lsCodProduto <> Me.ObjPai.ObjPai.Cod_Produto) Or _
       (lsCodClasf <> Me.ObjPai.Cod_Clasf) Or _
       (lsCodComplClasf <> Me.ObjPai.Cod_Compl_Clasf) Or _
       (lsCod_Cobert <> Me.Cod_Cobert) Or _
       (lsCod_Constr <> Me.ObjPai.Cod_Constr) Or _
       (lsNum_Cep_Risco <> Me.ObjPai.Num_Cep_Risco) Or _
       (lsTip_Atividade <> Me.ObjPai.gTip_Atividade) Or _
       (lsCod_Corr <> Me.ObjPai.ObjPai.objstcCtrlEmis.Cod_Corr) Then

        '==== Salva os parametros nas variaveis staticas p/ otimizar desempenho!
        lsCod_Sequencia = Me.ObjPai.ObjPai.Cod_Sequencia
        lsCod_Ramo = Me.Cod_Ramo
        lsCodPlano = Me.ObjPai.ObjPai.Cod_Estip
        lsCodProduto = Me.ObjPai.ObjPai.Cod_Produto
        lsCodClasf = Me.ObjPai.Cod_Clasf
        lsCodComplClasf = Me.ObjPai.Cod_Compl_Clasf
        lsCod_Cobert = Me.Cod_Cobert
        lsCod_Constr = Me.ObjPai.Cod_Constr
        lsNum_Cep_Risco = Me.ObjPai.Num_Cep_Risco
        lsTip_Atividade = Me.ObjPai.gTip_Atividade
        lsCod_Corr = Me.ObjPai.ObjPai.objstcCtrlEmis.Cod_Corr
        lsPerFranqDefault = 0
        lsValFraqMinDefault = 0
        lsValFraqMaxDefault = 0
        lsFatorFranqDefault = 0
        lsDescricaoFranquia = ""
        lsTem_VideObs = ""
        '====

        Call M46V999.mpNovoBuscaLimFranquia(Me.Cod_Ramo, Me.ObjPai.ObjPai.Cod_Estip, Me.ObjPai.ObjPai.Cod_Produto, _
                                            Me.Cod_Cobert, Me.ObjPai.Cod_Clasf, Me.ObjPai.Cod_Compl_Clasf, _
                                            Me.ObjPai.Cod_Constr, Me.ObjPai.Num_Cep_Risco, Me.ObjPai.ObjPai.Cod_Sequencia, Me.ObjPai.ObjPai.objstcCtrlEmis.Cod_Corr, _
                                            lsValFraqMinDefault, lsPerFranqDefault, lsFatorFranqDefault, lsValFraqMaxDefault, lsTem_VideObs, lsDescricaoFranquia, pNum_Meses_PI)

    End If

    '========Retorno
    pPerFranqDefault = lsPerFranqDefault
    pValFraqMinDefault = lsValFraqMinDefault
    pFatorFranqDefault = lsFatorFranqDefault
    pDescricaoFranquia = lsDescricaoFranquia
    pValFraqMaxDefault = lsValFraqMaxDefault
    pTem_VideObs = lsTem_VideObs

End Sub

Public Property Get Class_Name() As String
    Class_Name = "A46V704B"
End Property

'Private Function BuscaDadosAdicionais(ByRef pbdTAB_Cobert As ADODB.Connection, ByRef prsTab_Cobert As clsYasRecordSet, ByRef pMensagem As String) As Boolean
'    Dim lSelect As String
'    BuscaDadosAdicionais = False
'
'    If M46V999.gPrograma = "D046O003" Then
'        M46V111.gCod_Regiao = 0
'        If Me.ObjPai.ObjPai.Cod_Produto >= 1080 And Me.ObjPai.ObjPai.Cod_Estip = 1 Then
'            M46V111.gCod_plano = M46V999.gfPlanoBasico(Me.ObjPai.ObjPai.Cod_Ramo)
'        Else
'        M46V111.gCod_plano = Me.ObjPai.ObjPai.Cod_Estip
'        End If
'        M46V999.gCodProduto = Me.ObjPai.ObjPai.Cod_Produto
'    End If
'
'    lSelect = " Select distinct Cod_grupo_cobert, DSC_Cobert, Cod_Cobert, Per_Lim_Max_Cobert, COD_LIM_COBERT, TAB_COBERT  , PER_FRANQ , VAL_FRANQ_MIN, Val_Prem_Min, Val_Lim_Max_Cobert, Val_Lim_Min_Cobert from TAB_COBERT WITH (NOLOCK)"
'    lSelect = lSelect & " Where cod_Ramo = " & Cod_Ramo
'    If Me.ObjPai.ObjPai.Cod_Produto >= 1080 And Me.ObjPai.ObjPai.Cod_Estip < 1000 Then
'    lSelect = lSelect & " AND Cod_Cobert = " & M46V999.gfBuscarDePara(Cod_Cobert, 1)
'    Else
'       lSelect = lSelect & " AND Cod_Cobert = " & Cod_Cobert
'    End If
'
'    'If Val(Right(frmT46V101A.txtRisRegiao.Tag, 4)) <> 0 Then
'    If M46V111.gCod_Regiao <> 0 Then
'        lSelect = lSelect & " and tab_cobert = " & M46V111.gCod_Regiao
'    Else
'        'lSelect = lSelect & " and tab_cobert = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
'        If M46V111.gCod_plano = 0 Then
'            lSelect = lSelect & " and tab_cobert = " & Tab_Cobert
'        Else
'            If Tab_Cobert = 1 And M46V999.gCodProduto >= 1080 And M46V999.gPrograma <> "D046O003" Then
'                lSelect = lSelect & " and tab_cobert = " & M46V999.gfPlanoBasico(Cod_Ramo)
'            Else
'            lSelect = lSelect & " and tab_cobert = " & M46V111.gCod_plano
'            End If
'        End If
'    End If
'    lSelect = lSelect & " and COD_TARIF_INI <= " & M46V999.gCodProduto
'    lSelect = lSelect & " and COD_TARIF_TER >= " & M46V999.gCodProduto
'
'    If M46V999.gfObtRegistro(pbdTAB_Cobert, lSelect, prsTab_Cobert, pMensagem) = False Then
'        Call M46V999.gpFecha2(pbdTAB_Cobert)
'        Call M46V999.gpGraLog(2, pMensagem)
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'
'    BuscaDadosAdicionais = True
'End Function

' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Num_Item, "00000")) & _
               CStr(Format(Cod_Cobert, "0000")) & _
               CStr(Format(Tab_Cobert, "00000"))
End Property

Public Property Get gNomeCobertura() As String
    Dim lbdTAB_Cobert As ADODB.Connection     'Arquivo: P00Tarifa.
    Dim lMensagem As String                   'Mensagem.
    Dim lrsTab_Cobert As clsYasRecordSet      'Registro : TAB_COBERT.
    Dim lIndBasDados As String
    
    On Error GoTo Erro:

    Static lsNomeCobertura As String

    If ObjPai.ObjPai.Cod_Ramo <> 0 And ((lsNomeCobertura <> Nom_Cobertura) Or (Not M46V999.gfPreenchido(Nom_Cobertura))) Then
        If Cod_Cobert <> 0 Then
            If UCase(App.EXEName) = "D46O003" Then
                lIndBasDados = "P00MPED"
            Else
                lIndBasDados = "P00MULTI"
                'lIndBasDados = "P00MULTI2"
                'lIndBasDados = "P00MPED"
            End If
            
            If Not ObjStcPai Is Nothing Then
                If M46V999.gfAbrBasDados(lIndBasDados, M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTAB_Cobert, lMensagem) = False Then
                    Call M46V999.gpGraLog(2, lMensagem)
                    Exit Sub
                End If
                
                M46V999.gCod_Ramo = ObjPai.ObjPai.Cod_Ramo
                M46V999.gCodProduto = ObjPai.ObjPai.Cod_Produto
                M46V999.M46V111.gCod_plano = ObjPai.Cod_Plano
                If M46V999.gRst_Tab_Cobert.State = adStateClosed Then
                    Call M46V999.BuscaDadosAdicionais(lbdTAB_Cobert, lMensagem)
                Else
                    If M46V999.gRst_Tab_Cobert.EOF Then
                        Call M46V999.BuscaDadosAdicionais(lbdTAB_Cobert, lMensagem)
                    Else
                        M46V999.gRst_Tab_Cobert.Filter = "COD_Ramo = " & M46V999.gCod_Ramo & " AND COD_TARIF_INI = " & ObjPai.ObjPai.Cod_Produto & " AND tab_cobert = " & M46V999.M46V111.gCod_plano
                        If M46V999.gRst_Tab_Cobert.EOF Then
                            Call M46V999.BuscaDadosAdicionais(lbdTAB_Cobert, lMensagem)
                        End If
                        M46V999.gRst_Tab_Cobert.Filter = adFilterNone
                    End If
                End If
                
                If Not M46V999.gRst_Tab_Cobert.EOF Then
                    M46V999.gRst_Tab_Cobert.Filter = "cod_cobert = " & Cod_Cobert
                    
                    If M46V999.gRst_Tab_Cobert.EOF Then
                        Call M46V999.gpGraLog(0, "stcA46V704B - gNomeCobertura - Cobertura não encontrada Num_Ped: " & Me.Num_Ped & " Item: " & Me.Num_Item & " Cobert: " & Cod_Cobert)
                        Call M46V999.gpGraLog(0, "Count: " & M46V999.gRst_Tab_Cobert.RecordCount & " SQL: " & M46V999.gRst_Tab_Cobert.Source)
                    Else
                        Nom_Cobertura = M46V999.gRst_Tab_Cobert("DSC_Cobert")
                        Cod_Grupo_Cobertura = M46V999.gRst_Tab_Cobert("Cod_grupo_cobert")
                        PER_LIM_MAX_COBERT = M46V999.gRst_Tab_Cobert("PER_LIM_MAX_COBERT")
                        COD_LIM_COBERT = M46V999.gRst_Tab_Cobert("COD_LIM_COBERT")
                        
                        If Not IsNull(M46V999.gRst_Tab_Cobert("Val_Prem_Min")) Then
                            Val_Prem_Min = M46V999.gRst_Tab_Cobert("Val_Prem_Min")
                        End If
                        
                        lsNomeCobertura = Nom_Cobertura    '!!! Guarda na variavel statica
                        
                        If Not IsNull(M46V999.gRst_Tab_Cobert("Val_Lim_Max_Cobert")) Then
                            Val_Lim_Max_Cobert = M46V999.gRst_Tab_Cobert("Val_Lim_Max_Cobert")
                        End If
                        
                        If Not IsNull(M46V999.gRst_Tab_Cobert("Val_Lim_Min_Cobert")) Then
                            Val_Lim_Min_Cobert = M46V999.gRst_Tab_Cobert("Val_Lim_Min_Cobert")
                        End If
                    End If
                End If

'                If BuscaDadosAdicionais(lbdTAB_Cobert, lrsTab_Cobert, lMensagem) Then
'                    If Not lrsTab_Cobert.EOF Then
'                        Nom_Cobertura = M46V999.Check(lrsTab_Cobert("DSC_Cobert"))
'                        Cod_Grupo_Cobertura = M46V999.Check(lrsTab_Cobert("Cod_grupo_cobert"))
'                        PER_LIM_MAX_COBERT = M46V999.Check(lrsTab_Cobert("PER_LIM_MAX_COBERT"))
'                        COD_LIM_COBERT = M46V999.Check(lrsTab_Cobert("COD_LIM_COBERT"))
'                        Val_Prem_Min = M46V999.Check(lrsTab_Cobert("Val_Prem_Min"))
'                        lsNomeCobertura = Nom_Cobertura    '!!! Guarda na variavel statica
'                        Val_Lim_Max_Cobert = M46V999.Check(lrsTab_Cobert("Val_Lim_Max_Cobert"))
'                        Val_Lim_Min_Cobert = M46V999.Check(lrsTab_Cobert("Val_Lim_Min_Cobert"))
'                    End If
'                End If

            End If
        End If
    End If

Fim:
    Set lbdTAB_Cobert = Nothing
    M46V999.gRst_Tab_Cobert.Filter = adFilterNone
    gNomeCobertura = lsNomeCobertura
    Exit Property

Erro:
    Call M46V999.gpGraLog(2, "stcA46V704B - gNomeCobertura - " & Err.Description)
    GoTo Fim:
End Property

Public Property Get gNomeFranquia() As String
    Dim lbdTAB_Cobert As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lMensagem As String                     'Mensagem.

    Dim Dummy1 As Double
    Dim Dummy2 As Double
    Dim Dummy3 As Integer
    Dim Dummy6 As String
    Dim lretorno As String

    Static lsNomeFranquia As String

    If (lsNomeFranquia <> Nom_Franquia) Or (Not M46V999.gfPreenchido(Nom_Franquia)) Then


        lretorno = ""
        Call AtualizaLimFranquia(Dummy1, Dummy2, Dummy2, lretorno, Dummy2, Dummy6, Dummy3)
        lsNomeFranquia = lretorno

        If Trim(lsNomeFranquia) = "" Then
            lsNomeFranquia = gNomeCobertura
        End If
        Nom_Franquia = lsNomeFranquia
    End If

    gNomeFranquia = lsNomeFranquia

End Property


Public Property Get gTem_VideObs() As String
    Dim lMensagem As String                   'Mensagem.

    Dim Dummy1 As Double
    Dim Dummy2 As String
    Dim Dummy3 As Integer
    Dim lretorno As String


    Static lsTem_VideObs As String

    If (lsTem_VideObs <> Tem_VideObs) Or (Not M46V999.gfPreenchido(Tem_VideObs)) Then

        lretorno = ""
        Call AtualizaLimFranquia(Dummy1, Dummy1, Dummy1, Dummy2, Dummy1, lretorno, Dummy3)
        lsTem_VideObs = lretorno
        Tem_VideObs = lretorno
    End If

    gTem_VideObs = lsTem_VideObs

End Property



Public Property Get gCod_Grupo_Cobert() As String
    Dim lstrDummy As String

    If Not ObjPai.ObjPai Is Nothing Then ' Adicionado esta condição pois os dados da Tab_ped são acessados nesta função
        lstrDummy = gNomeCobertura
        gCod_Grupo_Cobert = Cod_Grupo_Cobertura
    End If

End Property

Public Property Get gVal_Prem_Min() As String
    Dim lstrDummy As String

    lstrDummy = gNomeCobertura
    gVal_Prem_Min = Val_Prem_Min

End Property

Public Property Get gPerLimMaxCobert() As Double
    Dim lstrDummy As String

    lstrDummy = gNomeCobertura
    gPerLimMaxCobert = PER_LIM_MAX_COBERT

End Property

Public Property Get gCodLimCobert() As Integer
    Dim lstrDummy As String

    lstrDummy = gNomeCobertura
    gCodLimCobert = COD_LIM_COBERT

End Property

Public Property Get gPerFranqDefault() As Double
    Dim lDummy1 As Double
    Dim lDummy2 As Double
    Dim lDummy3 As String
    Dim lDummy4 As Integer
    Dim lretorno As Double

    'Call mpBuscaLimFranquia(lRetorno, lDummy1, lDummy2)
    If Me.ObjPai.Cod_DizerItem = 0 And M46V999.gCodProduto < 1080 Then
        Call AtualizaLimFranquia(lretorno, lDummy1, lDummy2, lDummy3, lDummy2, lDummy3, lDummy4)
        gPerFranqDefault = lretorno
    Else
        gPerFranqDefault = Per_Franq
    End If

End Sub

Public Property Get gValFranqMinDefault() As Double
    Dim lDummy1 As Double
    Dim lDummy2 As Integer
    Dim lretorno As Double
    Dim lDummy3 As String
    Dim lDummy4 As Double

    'Call mpBuscaLimFranquia(lDummy1, lRetorno, lDummy2)
    If Me.ObjPai.Cod_DizerItem = 0 And M46V999.gCodProduto < 1080 Then
        Call AtualizaLimFranquia(lDummy1, lretorno, lDummy1, lDummy3, lDummy4, lDummy3, lDummy2)
        gValFranqMinDefault = lretorno
    Else
        gValFranqMinDefault = Me.Val_Franq_Min
    End If

End Property
Public Property Get gFatorFranqDefault() As Double
    Dim lDummy1 As Double
    Dim lDummy2 As Integer
    Dim lretorno As Double
    Dim lDummy3 As String
    Dim lDummy4 As Double
    Dim lrsFranqNovoLim As clsYasRecordSet
    

    'Call mpBuscaLimFranquia(lDummy1, lDummy2, lRetorno)
    If Me.ObjPai.Cod_DizerItem = 0 Then
        If Me.ObjPai.Cod_Iden_Tarif < 1080 Then
        Call AtualizaLimFranquia(lDummy1, lDummy1, lretorno, lDummy3, lDummy4, lDummy3, lDummy2)
        gFatorFranqDefault = lretorno
    Else
            
        Call M46V999.mfFranqNovoLim(Cod_Cobert, Val_IS, ObjPai.Cod_Clasf, ObjPai.Cod_Compl_Clasf, _
                                        ObjPai.Cod_Constr, ObjPai.Num_Cep_Risco, ObjPai.Tip_Cobert, 0, lrsFranqNovoLim)

        If lrsFranqNovoLim.RecordCount > 0 Then
            If lrsFranqNovoLim!Fator_Ajuste <> Null Then
                    gFatorFranqDefault = lrsFranqNovoLim!Fator_Ajuste
            End If
        End If
        End If
    Else
        gFatorFranqDefault = Me.Val_Fator_Franq
    End If

    Set lrsFranqNovoLim = Nothing

End Property


Public Property Get gPrazo() As Integer
    Static lsRetorno As Integer
    Static lsFlagCalculado As Boolean

    If lsFlagCalculado = False Then
        lsFlagCalculado = True
        lsRetorno = 0
        If Num_Meses_PI <> 0 Then
            Call mpPrazoTabCoefPI(lsRetorno)
            
            '8000048888, Franquia Perda Pagamento de Aluguel
            If Cod_Cobert = 1117 And Cod_Ramo = 112 And ObjTabPed.Cod_Produto >= 1170 Then  '1117 - Perda/Pagamento de Aluguel
                lsRetorno = 0
                If Me.gValFranqMinDefault > 0 Then
                    lsRetorno = gValFranqMinDefault
                End If
            End If
            
        End If
    End If

Retorno:
    gPrazo = lsRetorno
End Property
Private Sub mpPrazoTabCoefPI(ByRef pPrazo As Integer)
    Dim lBDP00Multi As ADODB.Connection
    Dim lMensagem As String
    Dim lrst As clsYasRecordSet
    Dim lSelect As String

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, _
                             lBDP00Multi, lMensagem) = False Then
        Call M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If
    lSelect = " Select Prazo from Tbm_COEF_PI WITH (NOLOCK) Where"
    lSelect = lSelect & " Cod_Ramo = " & ObjTabPed.Cod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & Cod_Cobert
    lSelect = lSelect & " and Num_meses = " & Num_Meses_PI
    lSelect = lSelect & " and Cod_Tarif = " & ObjTabPed.Cod_Produto
    lSelect = lSelect & " and Cod_Sequencia <= " & ObjTabPed.Cod_Sequencia

    If M46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
        Call M46V999.gpGraLog(1, lMensagem)
        pPrazo = 0
        GoTo SAIDA
    End If
    pPrazo = 0
    If Not lrst.EOF Then
        pPrazo = IIf(IsNull(lrst("Prazo")), 0, lrst("Prazo"))
    End If

    'Supermercado SMART
    If ObjTabPed.Cod_Ramo = 112 And ObjTabPed.Cod_Produto >= 1042 And (ObjTabPed.Cod_Estip >= 2178 And ObjTabPed.Cod_Estip <= 2182) And Cod_Cobert = M46V999.gfBuscarDePara(39, 1) Then
        pPrazo = 7
    End If


SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Call M46V999.gpFecha2(lBDP00Multi)
    Set lrst = Nothing
    Set lBDP00Multi = Nothing
    Return
End Sub

Public Property Get gCopiaCobert(pNumItem As Long) As stcA46V704B
    Dim lobjstcCobert As stcA46V704B
    Dim lobjstcFranquia As stcA46V738B
    Dim lobjstcFranquiaNova As stcA46V738B

    Set lobjstcCobert = New stcA46V704B
    Set lobjstcCobert.M46V999 = M46V999

    With lobjstcCobert
        .nosso_numero = nosso_numero
        .Cod_Cobert = Cod_Cobert
        .Num_Ped = .Num_Ped
        .Cod_Org = Cod_Org
        .Cod_Ramo = Cod_Ramo
        .Num_Apol = Num_Apol
        .Num_Endo = Num_Endo
        .Num_Item = pNumItem
        .Tab_Cobert = Tab_Cobert
        .Val_IS = Val_IS
        .Per_Taxa = Per_Taxa
        .Per_Desc_Coml = Per_Desc_Coml
        .Per_Desc_Exp = Per_Desc_Exp
        .VAL_PR = VAL_PR
        .Val_Pr_Anu = Val_Pr_Anu
        .Val_Pr_Tarif = Val_Pr_Tarif
        .Val_Pr_Tarif_Anu = Val_Pr_Tarif_Anu
        .Val_Pr_Resg = Val_Pr_Resg
        .Val_Pr_Resg_Anu = Val_Pr_Resg_Anu
        .Per_Franq = Per_Franq
        .Val_Franq_Min = Val_Franq_Min
        .Val_Franq_Max = Val_Franq_Max
        .NUM_PI = NUM_PI
        .Num_Meses_PI = Num_Meses_PI
        .COD_CLASSE = COD_CLASSE
        .Formula = Formula
        .Val_Fator_Franq = Val_Fator_Franq
        .Alterar_Cobert = Alterar_Cobert
        .Val_Risco_Cob = Val_Risco_Cob
        .Nom_Cobertura = Nom_Cobertura
        .Cod_Grupo_Cobertura = Cod_Grupo_Cobertura
        .AjusteEndosso = AjusteEndosso
        .TextoFranquia = TextoFranquia
        .Val_Pr_Orig = Val_Pr_Orig
        .PremioInformado = PremioInformado
        .Compl_Desc_Cobertura = Compl_Desc_Cobertura
        .Ind_Agrup_Cobert = Ind_Agrup_Cobert
        .VAL_PR_BATENTE = VAL_PR_BATENTE
        .Per_Desc_Franq = Per_Desc_Franq
        .Val_Lim_Max_Cobert = Val_Lim_Max_Cobert
    End With
    
    'Copia Franquia Nova
    For Each lobjstcFranquia In ColecaoNovaFranquia
        
        Set lobjstcFranquiaNova = New stcA46V738B
        With lobjstcFranquiaNova
            .Cod_Ramo = lobjstcFranquia.Cod_Ramo
            .Cod_Modalidade = lobjstcFranquia.Cod_Modalidade
            .Cod_Tarif = lobjstcFranquia.Cod_Tarif
            .Cod_Cobert = lobjstcFranquia.Cod_Cobert
            .Cod_Evento = lobjstcFranquia.Cod_Evento
            .Num_Ped = lobjstcFranquia.Num_Ped
            .Num_Apol = lobjstcFranquia.Num_Apol
            .Num_Item = pNumItem
            .Num_Endosso = lobjstcFranquia.Num_Endosso
            .nosso_numero = lobjstcFranquia.nosso_numero
            .Val_Min_Franq = lobjstcFranquia.Val_Min_Franq
            .Per_Franq = lobjstcFranquiaNova.Per_Franq
            .Texto_Franquia = lobjstcFranquia.Texto_Franquia
            .Texto_Observacoes = lobjstcFranquia.Texto_Observacoes
            .Multiplo_Franquia = lobjstcFranquia.Multiplo_Franquia
        End With
        
        lobjstcCobert.ColecaoNovaFranquia.Add lobjstcFranquiaNova
    Next
    
    Set gCopiaCobert = lobjstcCobert

SAIDA:
    GoSub FechaObj
    Exit Property

FechaObj:
    Set lobjstcCobert = Nothing
    Set lobjstcFranquia = Nothing
    Set lobjstcFranquiaNova = Nothing
    Return
End Property

Public Property Set ObjPai(pobjstc As stcA46V708B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get ObjPai() As stcA46V708B
    If ObjStcPai Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto ObjStcPai(Tipo:stcA46V708B) não atribuído para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe stcA46V708B já
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (stcA46V708B) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set ObjPai = ObjStcPai
End Property

Private Property Get ObjTabPed() As stcA46V702B
    Set ObjTabPed = ObjStcPai.ObjPai
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property

Public Property Get Ind_Calc_LMG() As Boolean
    Dim lretorno As Boolean
    '====!!!! Deve ser igual a procedure SQL RamosDiversos.dbo.proc_ListarCobertura
    'Select Case Me.ObjPai.Cod_Plano
    Select Case Me.Tab_Cobert
    Case 1020 To 1040
        lretorno = True
    Case 1050
        Select Case Me.Cod_Cobert
        Case M46V999.gfBuscarDePara(28, 1), M46V999.gfBuscarDePara(38, 1)
            lretorno = False
        Case Else
            lretorno = True
        End Select
    Case 1060
        Select Case Me.Cod_Cobert
        Case M46V999.gfBuscarDePara(9, 1), M46V999.gfBuscarDePara(22, 1), M46V999.gfBuscarDePara(23, 1), M46V999.gfBuscarDePara(28, 1), _
        M46V999.gfBuscarDePara(30, 1), M46V999.gfBuscarDePara(38, 1)
            lretorno = False
        Case Else
            lretorno = True
        End Select
    Case Is > 2000
        If Me.ObjPai.ObjPai.getPlano.ehPlanoEngessado = True Then
            lretorno = True
        Else
            lretorno = False
        End If
    Case Else
        lretorno = False
    End Select

    Ind_Calc_LMG = lretorno
End Property

Private Sub mpBuscaLimFranquiaDesativado(ByRef pPerFranqDefault As Double, ByRef pValFraqMinDefault As Double, ByRef pFatorFranqDefault As Double)
'Desativado Simony
    Dim lbdTbm_Fat_Franquia As ADODB.Connection         'Arquivo: P00Multi.
    Dim lMensagem As String         'Mensagem.
    Dim lrsTbm_Fat_Franquia As clsYasRecordSet          'Registro : Tbm_Fator_Franquia.
    Dim lCodSequencia As Integer
    Dim lSelect As String

    '====Parametros
    Static lsCod_Sequencia As Integer
    Static lsCod_Ramo As e_CodRamo
    Static lsCodPlano As Double
    Static lsCodProduto As Integer
    Static lsCodClasf As String
    Static lsCodComplClasf As String
    Static lsCod_Cobert As Integer
    Static lsCod_Constr As Integer
    Static lsNum_Cep_Risco As Long
    Static lsTip_Atividade As Integer
    Static lsCod_Corr As Long
    '====Retorno
    Static lsPerFranqDefault As Double
    Static lsValFraqMinDefault As Double
    Static lsFatorFranqDefault As Double

    If (lsCod_Sequencia <> Me.ObjPai.ObjPai.Cod_Sequencia) Or _
       (lsCod_Ramo <> Me.Cod_Ramo) Or _
       (lsCodPlano <> Me.ObjPai.ObjPai.Cod_Estip) Or _
       (lsCodProduto <> Me.ObjPai.ObjPai.Cod_Produto) Or _
       (lsCodClasf <> Me.ObjPai.Cod_Clasf) Or _
       (lsCodComplClasf <> Me.ObjPai.Cod_Compl_Clasf) Or _
       (lsCod_Cobert <> Me.Cod_Cobert) Or _
       (lsCod_Constr <> Me.ObjPai.Cod_Constr) Or _
       (lsNum_Cep_Risco <> Me.ObjPai.Num_Cep_Risco) Or _
       (lsTip_Atividade <> Me.ObjPai.gTip_Atividade) Or _
       (lsCod_Corr <> Me.ObjPai.ObjPai.objstcCtrlEmis.Cod_Corr) Then

        '==== Salva os parametros nas variaveis staticas p/ otimizar desempenho!
        lsCod_Sequencia = Me.ObjPai.ObjPai.Cod_Sequencia
        lsCod_Ramo = Me.Cod_Ramo
        lsCodPlano = Me.ObjPai.ObjPai.Cod_Estip
        lsCodProduto = Me.ObjPai.ObjPai.Cod_Produto
        lsCodClasf = Me.ObjPai.Cod_Clasf
        lsCodComplClasf = Me.ObjPai.Cod_Compl_Clasf
        lsCod_Cobert = Me.Cod_Cobert
        lsCod_Constr = Me.ObjPai.Cod_Constr
        lsNum_Cep_Risco = Me.ObjPai.Num_Cep_Risco
        lsTip_Atividade = Me.ObjPai.gTip_Atividade
        lsCod_Corr = Me.ObjPai.ObjPai.objstcCtrlEmis.Cod_Corr
        lsPerFranqDefault = 0
        lsValFraqMinDefault = 0
        lsFatorFranqDefault = 0
        '====

        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Fat_Franquia, lMensagem) = False Then
            Call M46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If

        lCodSequencia = Me.ObjPai.ObjPai.Cod_Sequencia

NovaSequencia1:

        '1 - Procura Registro de exceção com a chave completa
        'Procura o valor do Val_Franq_Min e Perc_Franq
        lSelect = " Select Val_Franq_Min as Min , Perc_Franquia as Perc, Fator_Ajuste as Fator from Tbm_Fator_Franquia WITH (NOLOCK) "
        lSelect = lSelect & " Where Cod_Ramo = " & Me.Cod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & Me.ObjPai.ObjPai.Cod_Estip
        lSelect = lSelect & " and Cod_Tip_Reg = 'E' "
        If Me.ObjPai.ObjPai.Cod_Produto >= 940 Then
            lSelect = lSelect & " and ((Cod_Clasf = '" & Me.ObjPai.Cod_Clasf & "'"
            lSelect = lSelect & " and Cod_Compl_Clasf = '" & Me.ObjPai.Cod_Compl_Clasf & "')"
            lSelect = lSelect & " or (Cod_Clasf = '0'))"
        Else
            lSelect = lSelect & " and Cod_Clasf = '" & Me.ObjPai.Cod_Clasf & "'"
            lSelect = lSelect & " and Cod_Compl_Clasf = '" & Me.ObjPai.Cod_Compl_Clasf & "'"
        End If
        lSelect = lSelect & " and Cod_Cobert = " & Me.Cod_Cobert
        lSelect = lSelect & " and Cod_Construcao in (0 , " & Me.ObjPai.Cod_Constr & ")"
        lSelect = lSelect & " and Cod_Tarif = " & Me.ObjPai.ObjPai.Cod_Produto
        lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia

        If Me.ObjPai.ObjPai.Cod_Produto >= 940 Then
            lSelect = lSelect & " AND ((CEP_INICIO <= " & Me.ObjPai.Num_Cep_Risco
            lSelect = lSelect & " AND CEP_FIM >= " & Me.ObjPai.Num_Cep_Risco & " ) "
            lSelect = lSelect & " OR (CEP_INICIO =  0 ))"
            lSelect = lSelect & " order by Prioridade desc "
        End If

        If M46V999.gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
            Call M46V999.gpFecha2(lbdTbm_Fat_Franquia)
            Call M46V999.gpGraLog(2, lMensagem)
        Else
            '2 - Se não encontrar a chave completa vai procurar o padrão
            If lrsTbm_Fat_Franquia.EOF Then
                If lCodSequencia > 0 Then
                    lCodSequencia = lCodSequencia - 1

                    GoTo NovaSequencia1
                Else
                    GoTo ProcuraDefault
                End If
            End If
        End If

        If lrsTbm_Fat_Franquia("Min") = 0 Then
            GoTo ProcuraDefault
        End If

        lsPerFranqDefault = M46V999.gfTratarNulo(lrsTbm_Fat_Franquia("Perc"), tpoNumerico)
        lsValFraqMinDefault = M46V999.gfTratarNulo(lrsTbm_Fat_Franquia("Min"), tpoNumerico)
        lsFatorFranqDefault = M46V999.gfTratarNulo(lrsTbm_Fat_Franquia("Fator"), tpoNumerico)

        GoTo FimBusca

ProcuraDefault:



        '3 - Procura o registro somente com Ramo, Modalidade, Cobertura e TipoAtividade
        'Procura o valor do Val_Franq_Min, Perc_Franq

        lCodSequencia = Me.ObjPai.ObjPai.Cod_Sequencia

NovaSequencia2:

        lSelect = " Select Val_Franq_Min as Min , Perc_Franquia as Perc, Fator_Ajuste as Fator  from Tbm_Fator_Franquia WITH (NOLOCK) "
        lSelect = lSelect & " Where Cod_Ramo = " & Me.Cod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & Me.ObjPai.ObjPai.Cod_Estip
        lSelect = lSelect & " and Cod_Tip_Reg = 'D' "
        lSelect = lSelect & " and Cod_Cobert = " & Me.Cod_Cobert
        lSelect = lSelect & " and Tip_Atividade in (3," & Me.ObjPai.gTip_Atividade & ")"
        lSelect = lSelect & " and Cod_Tarif = " & Me.ObjPai.ObjPai.Cod_Produto
        lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia

        If M46V999.gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
            Call M46V999.gpFecha2(lbdTbm_Fat_Franquia)
            Call M46V999.gpGraLog(2, lMensagem)
        Else
            If lrsTbm_Fat_Franquia.EOF Then
                If lCodSequencia > 0 Then
                    lCodSequencia = lCodSequencia - 1

                    GoTo NovaSequencia2
                Else
                    GoTo FimBusca
                End If
            End If
        End If

        lsPerFranqDefault = M46V999.gfTratarNulo(lrsTbm_Fat_Franquia("Perc"), tpoNumerico)
        lsValFraqMinDefault = M46V999.gfTratarNulo(lrsTbm_Fat_Franquia("Min"), tpoNumerico)
        lsFatorFranqDefault = M46V999.gfTratarNulo(lrsTbm_Fat_Franquia("Fator"), tpoNumerico)

FimBusca:
        Call M46V999.gpFecha2(lbdTbm_Fat_Franquia)

        '1001135 - Corretor Volks - Franquia Especial Concessionaria
        If Me.ObjPai.ObjPai.Cod_Produto >= 1002 Then
            If (InStr(1, "00008001/00008002/00008003/00008004/00008005/00008006/00008008/00008011/00008014/00008028", Format(Me.ObjPai.ObjPai.objstcCtrlEmis.Cod_Corr, "00000000"), vbTextCompare) > 0) And _
               (val(Me.ObjPai.Cod_Clasf) = 22 And val(Me.ObjPai.Cod_Compl_Clasf) = 13) And _
               (InStr(1, "0038/0064/0066", Format(Me.Cod_Cobert, "0000"), vbTextCompare) > 0) Then
                lsPerFranqDefault = 15
                lsValFraqMinDefault = 2000
            End If
        End If
    End If

    '========Retorno
    pPerFranqDefault = lsPerFranqDefault
    pValFraqMinDefault = lsValFraqMinDefault
    pFatorFranqDefault = lsFatorFranqDefault

End Sub

Public Function getRpt() As stcA46V704rpt

    If mRpt Is Nothing Then
        Set mRpt = New stcA46V704rpt
        Set mRpt.stcOrigem = Me
    End If

    Set getRpt = mRpt

End Function

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

Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
    Set ObjStcPai = Nothing
    Set mRpt = Nothing
    Set ColecaoNovaFranquia = New Collection
End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
    Set ObjStcPai = Nothing
    Set mRpt = Nothing
    Set ColecaoNovaFranquia = Nothing
End Sub

Private Sub mpMultiploTabCobertFraq(ByRef pMultiplo As Integer)
    Dim lBDP00Multi As ADODB.Connection
    Dim lMensagem As String
    Dim lrst As clsYasRecordSet
    Dim lSelect As String

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, _
                             lBDP00Multi, lMensagem) = False Then
        Call M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If
    lSelect = " Select Multiplo_Franquia from Tab_Ped_Cobert_Franquia WITH (NOLOCK) Where"
    lSelect = lSelect & " Cod_Ramo = " & Cod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & Tab_Cobert
    lSelect = lSelect & " and Cod_Tarif = " & ObjTabPed.Cod_Produto
    lSelect = lSelect & " and Num_Item = " & Num_Item
    lSelect = lSelect & " and Cod_Cobert = " & Cod_Cobert
    lSelect = lSelect & " and Nosso_Numero = " & M46V999.gfFormatarCamposSQL(nosso_numero)

    If M46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
        Call M46V999.gpGraLog(1, lMensagem)
        pMultiplo = 0
        GoTo SAIDA
    End If
    pMultiplo = 0
    If Not lrst.EOF Then
        pMultiplo = IIf(IsNull(lrst("Multiplo_Franquia")), 0, lrst("Multiplo_Franquia"))
    End If

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Call M46V999.gpFecha2(lBDP00Multi)
    Set lrst = Nothing
    Set lBDP00Multi = Nothing
    Return
End Sub

Public Property Get gMultiploFraq() As Integer
    Static lsRetorno As Integer

    Call mpMultiploTabCobertFraq(lsRetorno)
    
Retorno:
    gMultiploFraq = lsRetorno
End Property

Public Property Get gVal_Lim_Max_Cobert() As Double
    Dim lstrDummy As String

    lstrDummy = gNomeCobertura
    gVal_Lim_Max_Cobert = Val_Lim_Max_Cobert

End Property

Public Property Get gVal_Lim_Min_Cobert() As Double
    Dim lstrDummy As String

    lstrDummy = gNomeCobertura
    gVal_Lim_Min_Cobert = Val_Lim_Min_Cobert

End Property


Public Property Get Val_IS_Orig() As Double
    Val_IS_Orig = Val_IS_Orig_
End Property

Public Property Let Val_IS_Orig(ByVal vNewValue As Double)
    Val_IS_Orig_ = vNewValue
End Property


Public Property Get Per_Franq_Orig() As Double
    Per_Franq_Orig = Per_Franq_Orig_
End Property

Public Property Let Per_Franq_Orig(ByVal vNewValue As Double)
    Per_Franq_Orig_ = vNewValue
End Property

Public Property Get Tip_Origem_C4() As Integer
    Tip_Origem_C4 = Tip_Origem_C4_
End Property

Public Property Let Tip_Origem_C4(ByVal vNewValue As Integer)
    Tip_Origem_C4_ = vNewValue
End Property

Public Property Get RA_C4() As Integer
    RA_C4 = RA_C4_
End Property

Public Property Let RA_C4(ByVal vNewValue As Integer)
    RA_C4_ = vNewValue
End Property

Public Property Get VAL_IS_RA() As Double
    VAL_IS_RA = Val_IS_RA_
End Property

Public Property Let VAL_IS_RA(ByVal vNewValue As Double)
    Val_IS_RA_ = vNewValue
End Property

