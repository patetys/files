VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V101A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================================================================================
' Class Name: clsA46V101A - Gera arquivos para transmiss�o
' Sistema: P46V600
' Analista: Simony Sales
' ======================================================================================================================================================
'C�digos da Tabelas
'TAB_PED              = 001
'TAB_CTRL_EMIS        = 002
'TAB_PED_CORR         = 003
'TAB_PED_DIZER        = 004
'TAB_PED_VENC         = 005
'TAB_PED_JUROS        = 006
'TAB_PED_COSS         = 007
'TAB_PED_ENDER        = 008
'TAB_PED_LOC          = 009
'TAB_PED_CLAU         = 010
'TAB_PED_INSP         = 011
'TAB_PED_OBS          = 012
'TAB_PED_LOC_SEGURO   = 013
'TAB_PED_LOC_BEM      = 014
'TAB_PED_FAT_RISCO    = 015
'TAB_PED_COBERT       = 016
'TAB_PED_COBERT_COMPL = 017
'TAB_PED_ESPEC        = 018
'TAB_PED_LMI          = 019
'===================================================

Private mRegistro As String    'Registro.

Public Enum E_TipoEndosso
    e_AumentodeLMG = 1
    e_AlteracaodeVR = 2
    e_ReducaodeLMG = 3
    e_AlteracaodePlano = 4
    e_Inclusaodecobertura = 5
    e_EXclusaodecobertura = 6
    e_AlteracaodeLocaldeRisco = 7
    e_InclusaodeLocaldeRisco = 8
    e_ExclusaodeLocaldeRisco = 9
    e_Retificacaodeenderecodecorrespondencia = 10
    e_Alteracaodeclassificacao = 11
    e_InclusaodeClausulaBeneficiaria = 12
    e_ExclusaodaClausulaBeneficiaria = 13
    e_Retificacaodeclausulabeneficiaria = 14
    e_Alteracaodeclausulabeneficiaria = 15
    e_InclusaodeRelacaodeObjetosBens = 16
    e_Exclusaodeobjetosnarelacaoexistente = 17
    e_InclusaodasClausulas201ou202MSE = 18
    e_InclusaodasClausulas301ate317MSE = 19
    e_Inclusaoclausulas301a304MSR = 20
    e_Transferenciadecorretagem = 21
    e_AlteracaodeFranquia = 22
    e_AlteracaodoPeriodoIndenitario = 23
    e_AlteracaodeCPFouCNPJ = 24
    e_Cancelamentodaapolice = 25
    e_Inclusaodeprotecional = 26
    e_AlteracaoClausulaAss24hr = 27
    e_AlteracaoDadosPessoais = 28
    
    e_InclusaoRelacaoBicicletas = 29
    e_AlteracaoRelacaoBicicletas = 30
    e_ExclusaoRelacaoBicicletas = 31
End Enum



Private mClsM46V999 As clsM46V999


'Public Enum E_TipoEndosso
'    e_AumentodeLMG = 1
'    e_AlteracaodeVR = 2
'    e_ReducaodeLMG = 3
'    e_AlteracaodePlano = 4
'    e_Inclusaodecobertura = 5
'    e_EXclusaodecobertura = 6
'    e_AlteracaodeLocaldeRisco = 7
'    e_InclusaodeLocaldeRisco = 8
'    e_ExclusaodeLocaldeRisco = 9
'    e_Retificacaodeenderecodecorrespondencia = 10
'    e_Alteracaodeclassificacao = 11
'    e_InclusaodeClausulaBeneficiaria = 12
'    e_ExclusaodaClausulaBeneficiaria = 13
'    e_Retificacaodeclausulabeneficiaria = 14
'    e_Alteracaodeclausulabeneficiaria = 15
'    e_InclusaodeRelacaodeObjetosBens = 16
'    e_Exclusaodeobjetosnarelacaoexistente = 17
'    e_InclusaodasClausulas201ou202MSE = 18
'    e_InclusaodasClausulas301ate317MSE = 19
'    e_Inclusaoclausulas301a304MSR = 20
'    e_Transferenciadecorretagem = 21
'    e_AlteracaodeFranquia = 22
'    e_AlteracaodoPeriodoIndenitario = 23
'    e_AlteracaodeCPFouCNPJ = 24
'    e_Cancelamentodaapolice = 25
'    e_Inclusaodeprotecional = 26
'End Enum
'
'
'Enum e_CodDizer
'    e_CodDizerAlteracao = 12
'    e_CodDizerInclusao = 1
'    e_CodDizerexclusao = 4
'    e_CodDizerCancelamento = 10
'End Enum


Public Function mfSalvarTextoEndosso(ByRef pobjstcPedido As stcA46V702B, pTextoEndosso As String, _
                                    pMensagem As String, pPropTabela As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfSalvarTextoEndosso"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lobjTextoEndo As clsA46V722A
    Dim lobjstcTextoEndo As stcA46V722B
    Dim lNumPI As Long
    Dim lTextoInc As String
    Dim lTotLin As Long
    Dim lSeq As Long
    Dim AuxArray As Variant

    On Error GoTo ControleErro
    Set lobjTextoEndo = New clsA46V722A
    Set lobjTextoEndo.M46V999 = M46V999

    Set lobjstcTextoEndo = New stcA46V722B

    If pobjstcPedido.ColecaoTextoEndo.Count > 0 Then
        lobjstcTextoEndo.nosso_numero = pobjstcPedido.nosso_numero
        If Not lobjTextoEndo.mfExcluir(lobjstcTextoEndo, pPropTabela, "", tagDynaTrace) Then
            pMensagem = lobjTextoEndo.UltimoErro
            GoTo Finalizar
        End If
    End If

    If Trim(pTextoEndosso) = "" Then
        mfSalvarTextoEndosso = True
        Exit Function
    End If

    lSeq = 0
    lTotLin = 0
    Set pobjstcPedido.ColecaoTextoEndo = New Collection

    If InStr(pTextoEndosso, vbCrLf) <> 0 Then
        AuxArray = Split(pTextoEndosso, vbCrLf)
        lTextoInc = AuxArray(lSeq)
        lTotLin = UBound(AuxArray)
    Else
        lTextoInc = pTextoEndosso
    End If


OutraLinha:

    Set lobjstcTextoEndo = New stcA46V722B

    With lobjstcTextoEndo
        .nosso_numero = pobjstcPedido.nosso_numero
        .Cod_Org = 0
        .Cod_Ramo = M46V999.gCod_Ramo
        .Num_Ped = pobjstcPedido.Num_Ped
        .Num_Apol = 0
        .Num_Endo = 0
        .NUM_SEQ = lSeq + 1
        .Dsc_Dizer_Endo = lTextoInc
    End With

    If Not lobjTextoEndo.mfIncluir(lobjstcTextoEndo, pPropTabela, "", tagDynaTrace) Then
        pMensagem = lobjTextoEndo.UltimoErro
        GoTo Finalizar
    End If

    pobjstcPedido.ColecaoTextoEndo.Add lobjstcTextoEndo, lobjstcTextoEndo.Class_ID

    If lSeq < lTotLin And lSeq + 1 < 20 Then
        lSeq = lSeq + 1
        lTextoInc = AuxArray(lSeq)
        GoTo OutraLinha
    End If

    mfSalvarTextoEndosso = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

Finalizar:
    Set lobjTextoEndo = Nothing
    Set lobjstcTextoEndo = Nothing
    Exit Function
ControleErro:
    Set lobjTextoEndo = Nothing
    Set lobjstcTextoEndo = Nothing
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err.Description
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "mfSalvarTextoEndosso ERRO - " & Err & " " & Err.Description & _
                             Chr(13) & pMensagem, 2, Class_Name, methodName)
    Call M46V999.gpGraLog(0, "mfSalvarTextoEndosso ERRO - " & Err & " " & Err.Description & _
                             Chr(13) & pMensagem)

End Function

Public Sub mfCarregarItemOriginal(ByRef pobjstcItemPedido As stcA46V708B, ByRef pobjstcPedidoApolice As stcA46V702B)
    Dim lNomeTab As String
    Dim lNumIteApolice As Integer
    Dim lNumApolice As String
    Dim lNossoNumero As String
    Dim lNumEndo As Long
    Dim lNumPI As Long

    Dim lSelect As String
    Dim lMensagem As String

    Dim bdP00APOL As ADODB.Connection
    Dim RstAux As clsYasRecordSet

    Dim lobjPedLoc As stcA46V708B
    Dim lobjStcClausula As stcA46V703B
    Dim lobjstcCobert As stcA46V704B
    Dim lobjStcFatRisco As stcA46V720B
    Dim lobjStcLocBem As stcA46V715B
    Dim lobjstcLocBicicleta As stcA46V744B
    Dim lObjstcInsp As stcA46V717B    'Objeto da classe clsA46V717A (TAB_PED_INSP)
    Dim lobjstcAllRisks As stcA46V747B
    Dim lobjstcObraArte As stcA46V748B
    Dim Item As stcA46V708B

    If pobjstcItemPedido.M46V999.gPrograma = "D046O003" Then
        lNomeTab = "Syasportal.dbo.TAB_APOL"
    Else
        'If pobjstcPedido.M46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
        'Else
        '    lNomeTab = "GED.DBO.TAB_PED"
        'End If
    End If

    'lNomeTab = "GED.DBO.TAB_PED"

    'b) Abre arquivo.
    If pobjstcItemPedido.M46V999.gfAbrBasDados("P00APOL", pobjstcItemPedido.M46V999.gTipP00APOL, pobjstcItemPedido.M46V999.gEndP00APOL, bdP00APOL, lMensagem) = False Then
        GoTo SAIDA
    End If
    lNumApolice = pobjstcItemPedido.ObjPai.Num_Apol_An
    lNumIteApolice = pobjstcItemPedido.Num_Item

        lSelect = "SELECT TOP 1  "
        lSelect = lSelect & " Loc.*, "
        lSelect = lSelect & " Loc.NUM_ENDO AS NUM_ENDO2 , DAT_EMIS"
        lSelect = lSelect & " FROM " & lNomeTab & "_LOC "    ' & "_Atual"
        lSelect = lSelect & " as Loc WITH (NOLOCK) INNER JOIN " & lNomeTab  ' & "_Atual"
        lSelect = lSelect & " as Ped WITH (NOLOCK) on Loc.NUM_APOL = Ped.NUM_APOL "
        'If (Not pobjstcPedido.M46V999.gBaseLocal) And (Not pobjstcPedido.M46V999.gPrograma = "D046O003") Then
        '    lSelect = lSelect & " AND Loc.NUM_ENDO = Ped.NUM_ENDO "
        'End If
        lSelect = lSelect & " Where Loc.Num_Apol = " & lNumApolice

    If pobjstcItemPedido.M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
            lSelect = lSelect & " AND Loc.NOM_LOGR_RISCO not like '%CONFORME ESPEC%'"
        End If
        lSelect = lSelect & " AND COD_DIZER NOT IN (4,10)          "
        lSelect = lSelect & " AND NUM_ITEM = " & lNumIteApolice
        lSelect = lSelect & " ORDER BY NUM_ITEM, Ped.DAT_EMIS desc"

    If pobjstcItemPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
        Call pobjstcItemPedido.M46V999.gpFecha2(bdP00APOL)
            GoTo SAIDA
        End If

        'd) encontrou registro de ap�lice.
        If Not RstAux.EOF Then
            Set lobjPedLoc = New stcA46V708B

            Call mpCarregarAtributosItem(lobjPedLoc, RstAux, True)

            lNumPI = RstAux("Num_Pi")
            lNumEndo = RstAux("Num_Endo")
            On Error Resume Next
            lNossoNumero = RstAux("Nosso_Numero")
            On Error GoTo 0
            
            lSelect = "SELECT c.* , cl.*, Num_Meses_Pi, Val_Risco_Cob "
            lSelect = lSelect & " FROM " & lNomeTab & "_COBERT as c WITH (NOLOCK) "    ' & "_Atual"
            lSelect = lSelect & " left join " & lNomeTab & "_COBERT_COMPL as cl WITH (NOLOCK)"
            lSelect = lSelect & " on c.num_pi = cl.num_pi"
            lSelect = lSelect & " and c.cod_cobert = cl.cod_cobert"
            lSelect = lSelect & " and c.num_item = cl.num_item"
            lSelect = lSelect & " Where c.Num_PI = " & lNumPI
            'CR18725 Tunning de query
            'lSelect = lSelect & " AND c.NUM_ITEM = " & lNumIteApolice
            lSelect = lSelect & " AND c.NUM_ITEM = " & lNumIteApolice & " OPTION (MAXDOP 1)"
        If pobjstcItemPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
            Call pobjstcItemPedido.M46V999.gpFecha2(bdP00APOL)
                GoTo SAIDA
            End If

            'CR93893 Problema de DUPLICIDADE !!!! Verificar com a Simony APOL 0011644361 ROGER
            Dim lErr As Long
            Dim lDescr As String
            
            Do While Not RstAux.EOF
                Set lobjstcCobert = New stcA46V704B
                Call mpCarregarAtributosCobert(lobjstcCobert, RstAux)
                On Error Resume Next
                lobjPedLoc.ColecaoPedCobert.Add lobjstcCobert, lobjstcCobert.Class_ID

                If Err.Number > 0 Then
                    If Err.Number = 457 Then
                                M46V999.gpGraLog 0, "mfCarregarDadosOriginais - " & Err.Number & ": " & Err.Description
                                Err.Clear
                    Else
                        lErr = Err.Number
                        lDescr = Err.Description
                        On Error GoTo 0
                        Err.Raise lErr, lDescr
                        End If
                End If
                On Error GoTo 0
                RstAux.MoveNext
            Loop

                        On Error GoTo 0
            lSelect = "SELECT * "
            lSelect = lSelect & " FROM " & lNomeTab & "_CLAU WITH (NOLOCK)"    ' & "_Atual"
            lSelect = lSelect & " Where Num_PI = " & lNumPI
            lSelect = lSelect & " AND NUM_ITEM = " & lNumIteApolice

        If pobjstcItemPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
            Call pobjstcItemPedido.M46V999.gpFecha2(bdP00APOL)
                GoTo SAIDA
            End If

            Do While Not RstAux.EOF
                Set lobjStcClausula = New stcA46V703B
                Call mpCarregarAtributosClau(lobjStcClausula, RstAux)
                lobjPedLoc.ColecaoPedClau.Add lobjStcClausula, lobjStcClausula.Class_ID
                RstAux.MoveNext
            Loop


            lSelect = "SELECT * "
            lSelect = lSelect & " FROM " & lNomeTab & "_INSP WITH (NOLOCK) "    ' & "_Atual"
            If lNossoNumero <> "" Then
                lSelect = lSelect & " Where Nosso_Numero = " & lNossoNumero
                lSelect = Replace(UCase(lSelect), UCase(" ged.dbo"), UCase(" ramosdiversos.dbo"))
            Else
                lSelect = lSelect & " Where Num_PI = " & lNumPI
            End If
            lSelect = lSelect & " AND NUM_ITEM = " & lNumIteApolice

        If pobjstcItemPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
            Call pobjstcItemPedido.M46V999.gpFecha2(bdP00APOL)
            Else
                If Not RstAux.EOF Then
                    Set lObjstcInsp = New stcA46V717B
                    Call mpCarregarAtributosInspecao(lObjstcInsp, RstAux)
                    Set lobjPedLoc.objstcPedInsp = lObjstcInsp
                End If
            End If


            lSelect = "SELECT * "
            lSelect = lSelect & " FROM " & lNomeTab & "_FAT_RISCO WITH (NOLOCK)"    ' & "_Atual"
            lSelect = lSelect & " Where Num_PI = " & lNumPI
            lSelect = lSelect & " AND NUM_ITEM = " & lNumIteApolice

            lSelect = Replace(UCase(lSelect), UCase(" ged.dbo"), UCase(" ramosdiversos.dbo"))

        If pobjstcItemPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
            Call pobjstcItemPedido.M46V999.gpFecha2(bdP00APOL)
                GoTo SAIDA
            End If

            Do While Not RstAux.EOF
                Set lobjStcFatRisco = New stcA46V720B
                Call mpCarregarAtributosFatRisco(lobjStcFatRisco, RstAux)
                lobjPedLoc.ColecaoPedFator.Add lobjStcFatRisco, lobjStcFatRisco.Class_ID
                RstAux.MoveNext
            Loop

            'ProjetoFormularios
            'Carregar cole��o de beneficiario.
            lSelect = "SELECT * "
            lSelect = lSelect & " FROM ged.dbo.Tab_Ped_Loc_Pessoas WITH (NOLOCK)"    ' & "_Atual"
            lSelect = lSelect & " Where Num_PI = " & lNumPI

        If pobjstcItemPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
            Call pobjstcItemPedido.M46V999.gpFecha2(bdP00APOL)
                GoTo SAIDA
            End If

            Dim lobjPessoa As clsA46V735A
            Dim lobjstcPessoa As stcA46V735B

            Set lobjPessoa = New clsA46V735A
            Set lobjPessoa.M46V999 = M46V999
            Set lobjstcPessoa = New stcA46V735B

            Do While Not RstAux.EOF
                Call mpCarregarAtributosLocPessoa(lobjstcPessoa, RstAux)
                lobjstcPessoa.nosso_numero = lobjPedLoc.nosso_numero
                lobjPessoa.Colecao.Add Item:=lobjstcPessoa
                RstAux.MoveNext
            Loop
            Set lobjPedLoc.ColecaoBeneficiario = lobjPessoa.Colecao

            'lSelect = "SELECT * "
            'lSelect = lSelect & " FROM " & lNomeTab & "_LOC_BEM" ' & "_Atual"
            'lSelect = lSelect & " Where Num_PI = " & lNumPI

            'lSelect = Replace(UCase(lSelect), UCase(" ged.dbo"), UCase(" ramosdiversos.dbo"))

            'If pobjstcPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
            '    Call pobjstcPedido.M46V999.gpFecha2(bdP00APOL)
            '    goto SAIDA
            'End If

            'Do While Not RstAux.EOF
            '    Set lobjStcLocBem = New stcA46V715B
            '    Call mpCarregarAtributosLocBem(lobjStcLocBem, RstAux)
            '    lObjPedLoc.ColecaoPedLocBem.Add lobjStcLocBem, lobjStcLocBem.Class_ID
            '    RstAux.MoveNext
            'Loop
            
            lSelect = "SELECT * "
            lSelect = lSelect & " FROM " & lNomeTab & "_LOC_BICICLETA" ' & "_Atual"
            lSelect = lSelect & " Where Num_PI = " & lNumPI & " AND Num_Item = " & lNumIteApolice

            lSelect = Replace(UCase(lSelect), UCase(" ged.dbo"), UCase(" ramosdiversos.dbo"))
            lSelect = Replace(UCase(lSelect), UCase("Tab_Apol"), UCase("Tab_Ped"))
            
            
        If pobjstcItemPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
            Call pobjstcItemPedido.M46V999.gpFecha2(bdP00APOL)
                GoTo SAIDA
            End If

            Do While Not RstAux.EOF
                Set lobjstcLocBicicleta = New stcA46V744B
                Call mpCarregarAtributosLocBicicleta(lobjstcLocBicicleta, RstAux)
                lobjPedLoc.ColecaoPedLocBicicleta.Add lobjstcLocBicicleta
                RstAux.MoveNext
            Loop

            For Each Item In pobjstcPedidoApolice.ColecaoPedLoc
                If Item.Num_Item = lobjPedLoc.Num_Item Then
                    GoTo SAIDA
                End If
            Next
        
            pobjstcPedidoApolice.ColecaoPedLoc.Add lobjPedLoc, lobjPedLoc.Class_ID
        
        'AllRisks
        lSelect = "SELECT * "
        lSelect = lSelect & " FROM " & lNomeTab & "_All_Risks" ' & "_Atual"
        lSelect = lSelect & " Where Num_PI = " & lNumPI & " AND Num_Item = " & lNumIteApolice

        lSelect = Replace(UCase(lSelect), UCase(" ged.dbo"), UCase(" ramosdiversos.dbo"))
        lSelect = Replace(UCase(lSelect), UCase("Tab_Apol"), UCase("Tab_Ped"))
            
            
        If pobjstcItemPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
            Call pobjstcItemPedido.M46V999.gpFecha2(bdP00APOL)
            GoTo SAIDA
        End If

        Do While Not RstAux.EOF
            Set lobjstcAllRisks = New stcA46V747B
            Call mpCarregarAtributosAllRisks(lobjstcAllRisks, RstAux)
            lobjPedLoc.ColecaoAllRisks.Add lobjstcAllRisks
            RstAux.MoveNext
        Loop
        
        For Each Item In pobjstcPedidoApolice.ColecaoPedLoc
            If Item.Num_Item = lobjPedLoc.Num_Item Then
                GoTo SAIDA
            End If
        Next
    
        pobjstcPedidoApolice.ColecaoPedLoc.Add lobjPedLoc, lobjPedLoc.Class_ID
        
        'Obras Arte
        lSelect = "SELECT * "
        lSelect = lSelect & " FROM " & lNomeTab & "_Obras_Arte" ' & "_Atual"
        lSelect = lSelect & " Where Num_PI = " & lNumPI & " AND Num_Item = " & lNumIteApolice

        lSelect = Replace(UCase(lSelect), UCase(" ged.dbo"), UCase(" ramosdiversos.dbo"))
        lSelect = Replace(UCase(lSelect), UCase("Tab_Apol"), UCase("Tab_Ped"))
            
            
        If pobjstcItemPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
            Call pobjstcItemPedido.M46V999.gpFecha2(bdP00APOL)
            GoTo SAIDA
        End If

        Do While Not RstAux.EOF
            Set lobjstcObraArte = New stcA46V748B
            Call mpCarregarAtributosObraArte(lobjstcObraArte, RstAux)
            lobjPedLoc.ColecaoObrasArte.Add lobjstcObraArte
            RstAux.MoveNext
        Loop
        
        For Each Item In pobjstcPedidoApolice.ColecaoPedLoc
            If Item.Num_Item = lobjPedLoc.Num_Item Then
                GoTo SAIDA
            End If
        Next
        
            pobjstcPedidoApolice.ColecaoPedLoc.Add lobjPedLoc, lobjPedLoc.Class_ID
        End If

SAIDA:
    GoSub FechaObj

FechaObj:
    Set RstAux = Nothing
    Set lobjPedLoc = Nothing
    Set lobjStcClausula = Nothing
    Set lobjstcCobert = Nothing
    Set lobjStcFatRisco = Nothing
    Set lobjStcLocBem = Nothing
    Set lObjstcInsp = Nothing
    Set lobjstcAllRisks = Nothing
    Set lobjstcObraArte = Nothing
End Sub

Public Sub mfCarregarDadosOriginais(ByRef pobjstcPedido As stcA46V702B, ByRef pobjstcPedidoApolice As stcA46V702B)
    Dim lNomeTab As String
    Dim lNumApolice As String
    Dim lNumPI As Long

    Dim lSelect As String
    Dim lMensagem As String

    Dim bdP00APOL As ADODB.Connection
    Dim RstAux As clsYasRecordSet
    

    Dim lObjpedLocProp As stcA46V708B
    Dim lobjstcEndereco As stcA46V707B      'Objeto da classe clsA46V707A (TAB_PED_ENDER)
    Dim lobjstcCorretor As stcA46V705B      'Objeto da classe clsA46V705A (TAB_PED_CORR)
    Dim lobjstcTextoEspec As stcA46V726B    'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).

    If pobjstcPedido.M46V999.gPrograma = "D046O003" Then
        lNomeTab = "Syasportal.dbo.TAB_APOL"
    Else
        'If pobjstcPedido.M46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
        'Else
        '    lNomeTab = "GED.DBO.TAB_PED"
        'End If
    End If

    'lNomeTab = "GED.DBO.TAB_PED"

    'b) Abre arquivo.
    If pobjstcPedido.M46V999.gfAbrBasDados("P00APOL", pobjstcPedido.M46V999.gTipP00APOL, pobjstcPedido.M46V999.gEndP00APOL, bdP00APOL, lMensagem) = False Then
        GoTo SAIDA
    End If

    Set pobjstcPedidoApolice = New stcA46V702B

    lNumApolice = pobjstcPedido.Num_Apol_An

    If lNumApolice = 0 Then
        GoTo SAIDA    'Verificar pois isso n�o pode ocorrer
    End If

    lSelect = "SELECT * "
    lSelect = lSelect & " FROM " & lNomeTab & " WITH (NOLOCK)" ' & "_Atual "
    lSelect = lSelect & " Where Num_Apol = " & lNumApolice
    lSelect = lSelect & " AND [COD_DIZER] NOT IN (4,10) "
    lSelect = lSelect & " ORDER BY DAT_EMIS desc"

    If pobjstcPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
        Call pobjstcPedido.M46V999.gpFecha2(bdP00APOL)
        GoTo SAIDA
    End If

    If RstAux.EOF Then
        'N�o achou snif ..
        Set RstAux = Nothing
        GoTo SAIDA
    End If
    Call mpCarregarAtributosPedido(pobjstcPedidoApolice, RstAux)

    'Carrega a Tarifa da apolice.
    pobjstcPedido.M46V999.gNum_Apolice = pobjstcPedidoApolice.Num_Apol
    pobjstcPedidoApolice.Cod_Produto = pobjstcPedido.M46V999.gTarifa_Apolice

    lNumPI = RstAux("Num_Pi")

    lSelect = "SELECT * "
    lSelect = lSelect & " FROM " & lNomeTab & "_CORR WITH (NOLOCK)"    ' & "_Atual"
    lSelect = lSelect & " Where Num_PI = " & lNumPI

    If pobjstcPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
        Call pobjstcPedido.M46V999.gpFecha2(bdP00APOL)
        GoTo SAIDA
    End If

    Do While Not RstAux.EOF
        Set lobjstcCorretor = New stcA46V705B
        Call mpCarregarAtributosCorr(lobjstcCorretor, RstAux)
        pobjstcPedidoApolice.ColecaoPedCorr.Add lobjstcCorretor, lobjstcCorretor.Class_ID
        RstAux.MoveNext
    Loop

    Set lobjstcTextoEspec = New stcA46V726B    'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).

    lSelect = "SELECT * "
    lSelect = lSelect & " FROM " & lNomeTab & "_ESPEC WITH (NOLOCK)"    ' & "_Atual"
    lSelect = lSelect & " Where Num_PI = " & lNumPI
    If pobjstcPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
        Call pobjstcPedido.M46V999.gpFecha2(bdP00APOL)
        GoTo SAIDA
    End If

    Do While Not RstAux.EOF
        Set lobjstcTextoEspec = New stcA46V726B
        Call mpCarregarAtributosEspec(lobjstcTextoEspec, RstAux, lNomeTab)
        pobjstcPedidoApolice.ColecaoTextoEspec.Add lobjstcTextoEspec, lobjstcTextoEspec.Class_ID
        RstAux.MoveNext
    Loop

    lSelect = "SELECT * "
    lSelect = lSelect & " FROM " & lNomeTab & "_ENDER WITH (NOLOCK)"    ' & "_Atual"
    lSelect = lSelect & " Where Num_PI = " & lNumPI
    If pobjstcPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
        Call pobjstcPedido.M46V999.gpFecha2(bdP00APOL)
        GoTo SAIDA
    End If

    Set lobjstcEndereco = New stcA46V707B
    If Not RstAux.EOF Then
        'Set lobjstcEndereco = New stcA46V707B
        Call mpCarregarAtributosPedEnder(lobjstcEndereco, RstAux, "")
        'Set pobjstcPedidoApolice.objstcPedEndereco = lobjstcEndereco
    End If
    Set pobjstcPedidoApolice.objstcPedEndereco = lobjstcEndereco

    For Each lObjpedLocProp In pobjstcPedido.ColecaoPedLoc
        Set lObjpedLocProp.ObjPai = pobjstcPedido
        Set lObjpedLocProp.M46V999 = pobjstcPedido.M46V999
        Call mfCarregarItemOriginal(lObjpedLocProp, pobjstcPedidoApolice)
    Next

SAIDA:
    GoSub FechaObj

FechaObj:
    Set RstAux = Nothing
    Set lObjpedLocProp = Nothing
    Set lobjstcEndereco = Nothing
    Set lobjstcCorretor = Nothing
    Set lobjstcTextoEspec = Nothing
End Sub


Sub mpCarregarAtributosInspecao(ByRef pobjstc As stcA46V717B, prst As clsYasRecordSet)

    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_Item"), tpoNumerico)
    'pObjStc.NOSSO_NUMERO = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Num_Insp = M46V999.gfTratarNulo(prst("Num_Insp"), tpoNumerico)
    pobjstc.Dat_Cadastro = M46V999.gfTratarNulo(prst("Dat_Cadastro"), tpoNumerico)
    pobjstc.Pess_Cont = M46V999.gfTratarNulo(prst("Pess_Cont"), tpoString)
    pobjstc.Num_Tel = M46V999.gfTratarNulo(prst("Num_Tel"), tpoString)
    pobjstc.Num_Logr_Risco = M46V999.gfTratarNulo(prst("Num_Logr_Risco"), tpoNumerico)

End Sub


Private Sub mpCarregarAtributosItem(ByRef pobjstc As stcA46V708B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)

    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_item"), tpoNumerico)
    pobjstc.Num_Item_Ant = pobjstc.Num_Item
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("NUM_PI"), tpoString)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Cod_Tip_Logr_Risco = M46V999.gfTratarNulo(prst("Cod_Tip_Logr_Risco"), tpoString)
    pobjstc.Nom_Logr_Risco = M46V999.gfTratarNulo(prst("Nom_Logr_Risco"), tpoString)
    pobjstc.Dsc_Compl_Risco = M46V999.gfTratarNulo(prst("Dsc_Compl_Risco"), tpoString)
    pobjstc.Nom_Bairro_Risco = M46V999.gfTratarNulo(prst("Nom_Bairro_Risco"), tpoString)
    pobjstc.Nom_Cid_Risco = M46V999.gfTratarNulo(prst("Nom_Cid_Risco"), tpoString)
    pobjstc.Sig_UF_Risco = M46V999.gfTratarNulo(prst("Sig_UF_Risco"), tpoString)
    pobjstc.Num_Cep_Risco = M46V999.gfTratarNulo(prst("Num_CEP_Risco"), tpoNumerico)
    pobjstc.Per_Desc_Apar = M46V999.gfTratarNulo(prst("Per_Desc_Apar"), tpoNumerico)
    pobjstc.Per_Desc_Progres = M46V999.gfTratarNulo(prst("Per_Desc_Progres"), tpoNumerico)
    pobjstc.Per_IS_Vr = M46V999.gfTratarNulo(prst("Per_IS_Vr"), tpoNumerico)
    pobjstc.Val_Coef_IS_Vr = M46V999.gfTratarNulo(prst("Val_Coef_IS_Vr"), tpoNumerico)
    pobjstc.Val_Vr = M46V999.gfTratarNulo(prst("Val_Vr"), tpoNumerico)
    pobjstc.Tab_Clasf = M46V999.gfTratarNulo(prst("Tab_Clasf"), tpoNumerico)
    pobjstc.Cod_Clasf = M46V999.gfTratarNulo(prst("Cod_Clasf"), tpoString)
    pobjstc.Cod_Compl_Clasf = M46V999.gfTratarNulo(prst("Cod_Compl_Clasf"), tpoString)
    pobjstc.Dsc_Clasf = M46V999.gfTratarNulo(prst("Dsc_Clasf"), tpoString)
    pobjstc.Cod_Classe_Ocup = M46V999.gfTratarNulo(prst("Cod_Classe_Ocup"), tpoNumerico)
    pobjstc.Cod_Constr = M46V999.gfTratarNulo(prst("Cod_Constr"), tpoNumerico)
    pobjstc.Cod_User_Digit = M46V999.gfTratarNulo(prst("Cod_User_Digit"), tpoString)
    pobjstc.DAT_DIGIT = M46V999.gfTratarNulo(prst("Dat_Digit"), tpoNumerico)
    pobjstc.HOR_DIGIT = M46V999.gfTratarNulo(prst("Hor_Digit"), tpoNumerico)
    pobjstc.Cod_User_Liber = M46V999.gfTratarNulo(prst("Cod_User_Liber"), tpoString)
    pobjstc.Dat_Liber = M46V999.gfTratarNulo(prst("Dat_Liber"), tpoNumerico)
    pobjstc.Hor_Liber = M46V999.gfTratarNulo(prst("Hor_Liber"), tpoNumerico)
    pobjstc.Cod_Sit_item = M46V999.gfTratarNulo(prst("Cod_Sit_item"), tpoNumerico)
    pobjstc.Per_Bonus = M46V999.gfTratarNulo(prst("Per_Bonus"), tpoNumerico)
    pobjstc.Qtd_Empreg = M46V999.gfTratarNulo(prst("Qtd_Empreg"), tpoNumerico)
    pobjstc.Num_Insp = M46V999.gfTratarNulo(prst("Num_INSP"), tpoNumerico)
    pobjstc.Cod_Iden_Segur_01_V = M46V999.gfTratarNulo(prst("Cod_Iden_Segur_01"), tpoString)
    pobjstc.Cod_Iden_Segur_02_A = M46V999.gfTratarNulo(prst("Cod_Iden_Segur_02"), tpoString)
    pobjstc.Cod_Iden_Segur_03_P = M46V999.gfTratarNulo(prst("Cod_Iden_Segur_03"), tpoString)
    pobjstc.Cod_Iden_Segur_04_I = M46V999.gfTratarNulo(prst("Cod_Iden_Segur_04"), tpoString)
    pobjstc.Cod_Iden_Segur_05_G = M46V999.gfTratarNulo(prst("Cod_Iden_Segur_05"), tpoString)
    pobjstc.Cod_Iden_Segur_06 = M46V999.gfTratarNulo(prst("Cod_Iden_Segur_06"), tpoString)
    pobjstc.Cod_Iden_Segur_07 = M46V999.gfTratarNulo(prst("Cod_Iden_Segur_07"), tpoString)
    pobjstc.Cod_Iden_Segur_08 = M46V999.gfTratarNulo(prst("Cod_Iden_Segur_08"), tpoString)
    pobjstc.Cod_Iden_Segur_09 = M46V999.gfTratarNulo(prst("Cod_Iden_Segur_09"), tpoString)
    pobjstc.Cod_Iden_Segur_10 = M46V999.gfTratarNulo(prst("Cod_Iden_Segur_10"), tpoString)
    pobjstc.Per_Tot_Segur = M46V999.gfTratarNulo(prst("Per_Tot_Segur"), tpoNumerico)
    pobjstc.Cod_Iden_Protec_01_E = M46V999.gfTratarNulo(prst("Cod_Iden_Protec_01"), tpoString)
    pobjstc.Cod_Iden_Protec_02_H = M46V999.gfTratarNulo(prst("Cod_Iden_Protec_02"), tpoString)
    pobjstc.Cod_Iden_Protec_03_S = M46V999.gfTratarNulo(prst("Cod_Iden_Protec_03"), tpoString)
    pobjstc.Cod_Iden_Protec_04_D = M46V999.gfTratarNulo(prst("Cod_Iden_Protec_04"), tpoString)
    pobjstc.Cod_Iden_Protec_05 = M46V999.gfTratarNulo(prst("Cod_Iden_Protec_05"), tpoString)
    pobjstc.Cod_Iden_Protec_06 = M46V999.gfTratarNulo(prst("Cod_Iden_Protec_06"), tpoString)
    pobjstc.Cod_Iden_Protec_07 = M46V999.gfTratarNulo(prst("Cod_Iden_Protec_07"), tpoString)
    pobjstc.Cod_Iden_Protec_08 = M46V999.gfTratarNulo(prst("Cod_Iden_Protec_08"), tpoString)
    pobjstc.Cod_Iden_Protec_09 = M46V999.gfTratarNulo(prst("Cod_Iden_Protec_09"), tpoString)
    pobjstc.Cod_Iden_Protec_10 = M46V999.gfTratarNulo(prst("Cod_Iden_Protec_10"), tpoString)
    pobjstc.Per_Tot_Protec = M46V999.gfTratarNulo(prst("Per_Tot_Protec"), tpoNumerico)
    pobjstc.Cod_Produto = M46V999.gfTratarNulo(prst("Cod_Produto"), tpoNumerico)
    pobjstc.Cod_Plano = pobjstc.Cod_Produto
    pobjstc.Cod_Regiao = M46V999.gfTratarNulo(prst("Cod_Regiao"), tpoNumerico)
    pobjstc.Cod_Iden_Tarif = M46V999.gfTratarNulo(prst("Cod_Iden_Tarif"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
    'pObjStc.Tip_Emissao = M46V999.gfTratarNulo(prst("Tip_Emissao"), tpoNumerico)
    'pObjStc.Num_Apol_Anterior = M46V999.gfTratarNulo(prst("Num_Apol_Anterior"), tpoNumerico)
    'pObjStc.Num_Item_Anterior = M46V999.gfTratarNulo(prst("Num_Item_Anterior"), tpoNumerico)
    'pObjStc.Ren_Con_Seguradora = M46V999.gfTratarNulo(prst("Ren_Con_Seguradora"), tpoNumerico)
    'pObjStc.Ren_Con_Apolice = M46V999.gfTratarNulo(prst("Ren_Con_Apolice"), tpoString)
    'pObjStc.Ren_Con_Item = M46V999.gfTratarNulo(prst("Ren_Con_Item"), tpoString)
    'pObjStc.Ind_Sinistro = M46V999.gfTratarNulo(prst("Ind_Sinistro"), tpoNumerico)
    'pObjStc.Insp_Obr = M46V999.gfTratarNulo(prst("Insp_Obr"), TpoBoolean)
    'pObjStc.Insp_Agend = M46V999.gfTratarNulo(prst("Insp_Agend"), TpoBoolean)
    'pObjStc.Dig_Bens = M46V999.gfTratarNulo(prst("Dig_Bens"), TpoBoolean)
    'pObjStc.Cob_Bens = M46V999.gfTratarNulo(prst("Cob_Bens"), TpoBoolean)
    'pObjStc.Tip_Cobert = M46V999.gfTratarNulo(prst("Tip_Cobert"), TpoNumInteiro)
    'pObjStc.Ind_Sin_Inspec = M46V999.gfTratarNulo(prst("Ind_Sin_Inspec"), TpoNumInteiro)

End Sub


Private Sub mpCarregarAtributosClau(ByRef pobjstc As stcA46V703B, prst As clsYasRecordSet)
    pobjstc.Cod_Clau = M46V999.gfTratarNulo(prst("Cod_Clau"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_item"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Descricao = M46V999.gfTratarNulo(prst("Descricao"), tpoString)
End Sub

Private Sub mpCarregarAtributosCobert(ByRef pobjstc As stcA46V704B, prst As clsYasRecordSet)

    pobjstc.Cod_Cobert = M46V999.gfTratarNulo(prst("Cod_Cobert"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("NOsso_Numero"), tpoString)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_item"), tpoNumerico)
    pobjstc.Tab_Cobert = M46V999.gfTratarNulo(prst("Tab_Cobert"), tpoNumerico)
    pobjstc.Val_IS = M46V999.gfTratarNulo(prst("Val_IS"), tpoNumerico)
    pobjstc.Per_Taxa = M46V999.gfTratarNulo(prst("Per_Taxa"), tpoNumerico)
    pobjstc.Per_Desc_Coml = M46V999.gfTratarNulo(prst("Per_Desc_Coml"), tpoNumerico)
    pobjstc.Per_Desc_Exp = M46V999.gfTratarNulo(prst("Per_Desc_Exp"), tpoNumerico)
    pobjstc.VAL_PR = M46V999.gfTratarNulo(prst("Val_PR"), tpoNumerico)
    pobjstc.Val_Pr_Anu = M46V999.gfTratarNulo(prst("Val_Pr_Anu"), tpoNumerico)
    pobjstc.Val_Pr_Tarif = M46V999.gfTratarNulo(prst("Val_Pr_Tarif"), tpoNumerico)
    pobjstc.Val_Pr_Tarif_Anu = M46V999.gfTratarNulo(prst("Val_Pr_Tarif_Anu"), tpoNumerico)
    pobjstc.Val_Pr_Resg = M46V999.gfTratarNulo(prst("Val_Pr_Resg"), tpoNumerico)
    pobjstc.Val_Pr_Resg_Anu = M46V999.gfTratarNulo(prst("Val_Pr_Resg_Anu"), tpoNumerico)
    pobjstc.Per_Franq = M46V999.gfTratarNulo(prst("Per_Franq"), tpoNumerico)
    pobjstc.Val_Franq_Min = M46V999.gfTratarNulo(prst("Val_Franq_Min"), tpoNumerico)
    pobjstc.Val_Franq_Max = M46V999.gfTratarNulo(prst("Val_Franq_Max"), tpoNumerico)
    pobjstc.Num_Meses_PI = M46V999.gfTratarNulo(prst("Num_Meses_PI"), tpoNumerico)
    pobjstc.COD_CLASSE = M46V999.gfTratarNulo(prst("Cod_Classe"), tpoNumerico)
    pobjstc.Formula = M46V999.gfTratarNulo(prst("Formula"), tpoString)
    pobjstc.Val_Fator_Franq = M46V999.gfTratarNulo(prst("Val_Fator_Franq"), TpoValor)
    pobjstc.Alterar_Cobert = M46V999.gfTratarNulo(prst("Alterar_Cobert"), tpoString)
    pobjstc.Val_Risco_Cob = M46V999.gfTratarNulo(prst("Val_Risco_Cob"), TpoValor)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
    pobjstc.AjusteEndosso = M46V999.gfTratarNulo(prst("AjusteEndosso"), tpoString)
    pobjstc.AlterouFranquia = M46V999.gfTratarNulo(prst("AlterouFranquia"), tpoString)
    pobjstc.TextoFranquia = M46V999.gfTratarNulo(prst("TextoFranquia"), tpoString)
    pobjstc.Val_Pr_Orig = M46V999.gfTratarNulo(prst("Val_Pr_Orig"), TpoValor)
End Sub

Private Sub mpCarregarAtributosLocBem(ByRef pobjstc As stcA46V715B, prst As clsYasRecordSet)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_Item"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Dsc_Bem = M46V999.gfTratarNulo(prst("Dsc_Bem"), tpoString)
    pobjstc.Val_Bem = M46V999.gfTratarNulo(prst("Val_Bem"), tpoNumerico)
    pobjstc.Cod_Bem = M46V999.gfTratarNulo(prst("Cod_Bem"), tpoNumerico)
    pobjstc.Cod_Seq = M46V999.gfTratarNulo(prst("Cod_Seq"), tpoNumerico)
End Sub

Private Sub mpCarregarAtributosLocBicicleta(ByRef pobjstc As stcA46V744B, prst As clsYasRecordSet)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_Item"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("NUM_PI"), tpoNumerico)
    pobjstc.Cod_Seq = M46V999.gfTratarNulo(prst("Cod_Seq"), tpoNumerico)
    pobjstc.Descricao = M46V999.gfTratarNulo(prst("Descricao"), tpoString)
    pobjstc.Valor = M46V999.gfTratarNulo(prst("Valor"), tpoNumerico)
End Sub

Private Sub mpCarregarAtributosAllRisks(ByRef pobjstc As stcA46V747B, prst As clsYasRecordSet)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_Item"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("NUM_PI"), tpoNumerico)
    pobjstc.Cod_Seq = M46V999.gfTratarNulo(prst("Cod_Seq"), tpoNumerico)
    pobjstc.Descricao = M46V999.gfTratarNulo(prst("Descricao"), tpoString)
    pobjstc.Valor = M46V999.gfTratarNulo(prst("Valor"), tpoNumerico)
End Sub

Private Sub mpCarregarAtributosObraArte(ByRef pobjstc As stcA46V748B, prst As clsYasRecordSet)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_Item"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("NUM_PI"), tpoNumerico)
    pobjstc.Cod_Seq = M46V999.gfTratarNulo(prst("Cod_Seq"), tpoNumerico)
    pobjstc.Descricao = M46V999.gfTratarNulo(prst("Descricao"), tpoString)
    pobjstc.Valor = M46V999.gfTratarNulo(prst("Valor"), tpoNumerico)
End Sub

Private Sub mpCarregarAtributosLocPessoa(ByRef pobjstc As stcA46V735B, prst As clsYasRecordSet)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_Item"), tpoNumerico)

    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_Item"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)

    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)

    pobjstc.Ped_Loc_Pessoa_ID = M46V999.gfTratarNulo(prst("Ped_Loc_Pessoa_ID"), tpoNumerico)
    pobjstc.Tip_Loc_Pessoa = M46V999.gfTratarNulo(prst("Tip_Loc_Pessoa"), tpoNumerico)
    pobjstc.Nom_Pessoa = M46V999.gfTratarNulo(prst("Nom_Pessoa"), tpoString)
    pobjstc.Tip_Pessoa = M46V999.gfTratarNulo(prst("Tip_Pessoa"), tpoNumerico)
    pobjstc.Num_CPF_CNPJ = M46V999.gfTratarNulo(prst("Num_CPF_CNPJ"), tpoString)

    pobjstc.Tip_Cobert = M46V999.gfTratarNulo(prst("Tip_Cobert"), tpoNumerico)
    pobjstc.Per_Participacao = M46V999.gfTratarNulo(prst("Per_Participacao"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_Item"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
End Sub

Private Sub mpCarregarAtributosFatRisco(ByRef pobjstc As stcA46V720B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("NOsso_Numero"), tpoString)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_item"), tpoNumerico)
    pobjstc.Cod_Fator = M46V999.gfTratarNulo(prst("Cod_Fator"), tpoNumerico)
    pobjstc.Dsc_Compl_Fator = M46V999.gfTratarNulo(prst("Dsc_Compl_Fator"), tpoString)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
End Sub


Private Sub mpCarregarAtributosEspec(ByRef pobjstc As stcA46V726B, ByRef prst As clsYasRecordSet, ByVal pNomeTabImport As String)
    Dim lTemp As Integer
    Dim lAlter As String

    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Cod_Espec = M46V999.gfTratarNulo(prst("Cod_Espec"), tpoNumerico)
    pobjstc.Espec_lin1 = M46V999.gfTratarNulo(prst("Espec_lin1"), tpoString)
    pobjstc.Espec_lin2 = M46V999.gfTratarNulo(prst("Espec_lin2"), tpoString)
    pobjstc.Espec_lin3 = M46V999.gfTratarNulo(prst("Espec_lin3"), tpoString)
    pobjstc.Espec_lin4 = M46V999.gfTratarNulo(prst("Espec_lin4"), tpoString)
    pobjstc.Espec_lin5 = M46V999.gfTratarNulo(prst("Espec_lin5"), tpoString)
    pobjstc.Espec_lin6 = M46V999.gfTratarNulo(prst("Espec_lin6"), tpoString)
    pobjstc.Espec_lin7 = M46V999.gfTratarNulo(prst("Espec_lin7"), tpoString)
    pobjstc.Espec_lin8 = M46V999.gfTratarNulo(prst("Espec_lin8"), tpoString)
    pobjstc.Espec_lin9 = M46V999.gfTratarNulo(prst("Espec_lin9"), tpoString)
    pobjstc.Espec_lin10 = M46V999.gfTratarNulo(prst("Espec_lin10"), tpoString)
    pobjstc.Espec_lin11 = M46V999.gfTratarNulo(prst("Espec_lin11"), tpoString)
    pobjstc.Espec_lin12 = M46V999.gfTratarNulo(prst("Espec_lin12"), tpoString)
    pobjstc.Espec_lin13 = M46V999.gfTratarNulo(prst("Espec_lin13"), tpoString)
    pobjstc.Espec_lin14 = M46V999.gfTratarNulo(prst("Espec_lin14"), tpoString)
    pobjstc.Espec_lin15 = M46V999.gfTratarNulo(prst("Espec_lin15"), tpoString)
    pobjstc.Espec_lin16 = M46V999.gfTratarNulo(prst("Espec_lin16"), tpoString)
    pobjstc.Espec_lin17 = M46V999.gfTratarNulo(prst("Espec_lin17"), tpoString)
    pobjstc.Espec_lin18 = M46V999.gfTratarNulo(prst("Espec_lin18"), tpoString)
    pobjstc.Espec_lin19 = M46V999.gfTratarNulo(prst("Espec_lin19"), tpoString)
    pobjstc.Espec_lin20 = M46V999.gfTratarNulo(prst("Espec_lin20"), tpoString)
    pobjstc.Num_pag = M46V999.gfTratarNulo(prst("Num_pag"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
End Sub

Private Sub mpCarregarAtributosPedEnder(ByRef pobjstc As stcA46V707B, ByRef prst As clsYasRecordSet, ByVal pNomeTabImport As String)
    Dim lTemp As Integer
    Dim lAlter As String
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Cod_Tip_Logr = M46V999.gfTratarNulo(prst("Cod_Tip_Logr"), tpoString)
    pobjstc.Nom_Logr = M46V999.gfTratarNulo(prst("Nom_Logr"), tpoString)
    pobjstc.Dsc_Compl_Logr = M46V999.gfTratarNulo(prst("Dsc_Compl_Logr"), tpoString)
    pobjstc.Nom_Bairro = M46V999.gfTratarNulo(prst("Nom_Bairro"), tpoString)
    pobjstc.Nom_Cid = M46V999.gfTratarNulo(prst("Nom_Cid"), tpoString)
    pobjstc.Sig_UF = M46V999.gfTratarNulo(prst("Sig_UF"), tpoString)
    pobjstc.Num_CEP = M46V999.gfTratarNulo(prst("Num_CEP"), tpoNumerico)
    pobjstc.Num_Fone = M46V999.gfTratarNulo(prst("Num_Fone"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
    'pObjStc.Cod_Reg_Cons = M46V999.gfTratarNulo(prst("Cod_Reg_Cons"), tpoNumerico)
    'pObjStc.Num_Reg_Cons = M46V999.gfTratarNulo(prst("Num_Reg_Cons"), tpoString)

End Sub

Private Sub mpCarregarAtributosCorr(ByRef pobjstc As stcA46V705B, prst As clsYasRecordSet)

    pobjstc.Cod_Corr = M46V999.gfTratarNulo(prst("Cod_Corr"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Num_PI"), tpoString)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Per_Corr = M46V999.gfTratarNulo(prst("Per_Corr"), tpoNumerico)
    pobjstc.Num_Ordem_Corr = M46V999.gfTratarNulo(prst("Num_Ordem_Corr"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)

End Sub

Private Sub mpCarregarAtributosPedido(ByRef pobjstc As stcA46V702B, ByRef prst As clsYasRecordSet)

    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("NUM_PI"), tpoString)
    'pObjStc.Cod_Unicidade = M46V999.gfTratarNulo(prst("Cod_Unicidade"), tpoNumerico)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Num_Apol_An = M46V999.gfTratarNulo(prst("Num_Apol_An"), tpoNumerico)
    pobjstc.Cod_Reg = M46V999.gfTratarNulo(prst("Cod_Reg"), tpoNumerico)
    pobjstc.Cod_Bco_Cobr = M46V999.gfTratarNulo(prst("Cod_Bco_Cobr"), tpoNumerico)
    pobjstc.Cod_Agen_Angar = M46V999.gfTratarNulo(prst("Cod_Agen_Angar"), tpoNumerico)
    pobjstc.Cod_Agen_Indicr = M46V999.gfTratarNulo(prst("Cod_Agen_Indicr"), tpoNumerico)
    pobjstc.Num_Cta_Indicr = M46V999.gfTratarNulo(prst("Num_Cta_Indicr"), tpoNumerico)
    pobjstc.Cod_Unid_Prod = M46V999.gfTratarNulo(prst("Cod_Unid_Prod"), tpoNumerico)
    pobjstc.Cod_Prod = M46V999.gfTratarNulo(prst("Cod_Prod"), tpoNumerico)
    pobjstc.Cod_Agen_Bas_Cor = M46V999.gfTratarNulo(prst("Cod_Agen_Bas_Cor"), tpoNumerico)
    pobjstc.Num_Cta_Bas_Cor = M46V999.gfTratarNulo(prst("Num_Cta_Bas_Cor"), tpoNumerico)
    pobjstc.Cod_Segurado = M46V999.gfTratarNulo(prst("Cod_Segurado"), tpoNumerico)
    pobjstc.Nom_Segurado = M46V999.gfTratarNulo(prst("Nom_Segurado"), tpoString)
    pobjstc.Dsc_Compl_Segurado = M46V999.gfTratarNulo(prst("Dsc_Compl_Segurado"), tpoString)
    pobjstc.Num_Cgc_Cpf = M46V999.gfTratarNulo(prst("Num_CGC_CPF"), tpoString)
    pobjstc.Cod_Tip_Logr = M46V999.gfTratarNulo(prst("Cod_Tip_Logr"), tpoString)
    pobjstc.Nom_Logr = M46V999.gfTratarNulo(prst("Nom_Logr"), tpoString)
    pobjstc.Dsc_Compl_Logr = M46V999.gfTratarNulo(prst("Dsc_Compl_Logr"), tpoString)
    pobjstc.Nom_Bairro = M46V999.gfTratarNulo(prst("Nom_Bairro"), tpoString)
    pobjstc.Nom_Cid = M46V999.gfTratarNulo(prst("Nom_Cid"), tpoString)
    pobjstc.Sig_UF = M46V999.gfTratarNulo(prst("Sig_UF"), tpoString)
    pobjstc.Num_CEP = M46V999.gfTratarNulo(prst("Num_CEP"), tpoNumerico)
    pobjstc.Num_Fone = M46V999.gfTratarNulo(prst("Num_Fone"), tpoNumerico)
    pobjstc.Cod_Estip = M46V999.gfTratarNulo(prst("Cod_Estip"), tpoNumerico)
    pobjstc.Val_Coef_Cms_Pad = M46V999.gfTratarNulo(prst("Val_Coef_Cms_Pad"), tpoNumerico)
    pobjstc.Dat_Ini_Vig = M46V999.gfTratarNulo(prst("Dat_Ini_Vig"), tpoNumerico)
    pobjstc.Dat_Ter_Vig = M46V999.gfTratarNulo(prst("Dat_Ter_Vig"), tpoNumerico)
    pobjstc.Dat_Emis = M46V999.gfTratarNulo(prst("Dat_Emis"), tpoNumerico)
    pobjstc.Cod_Moed = M46V999.gfTratarNulo(prst("Cod_Moed"), tpoNumerico)
    pobjstc.Per_Juro_Frac = M46V999.gfTratarNulo(prst("Per_Juro_Frac"), tpoNumerico)
    pobjstc.Qtd_Parc = M46V999.gfTratarNulo(prst("Qtd_Parc"), tpoNumerico)
    pobjstc.Qtd_Anexo = M46V999.gfTratarNulo(prst("Qtd_Anexo"), tpoNumerico)
    pobjstc.Cod_Iden_Praz = M46V999.gfTratarNulo(prst("Cod_Iden_Praz"), tpoNumerico)
    pobjstc.Cod_Pr_Min_Endo = M46V999.gfTratarNulo(prst("Cod_Pr_Min_Endo"), tpoNumerico)
    pobjstc.Cod_Pr_Min_Parc = M46V999.gfTratarNulo(prst("Cod_Pr_Min_Parc"), tpoNumerico)
    pobjstc.Val_Pr_Antec = M46V999.gfTratarNulo(prst("Val_Pr_Antec"), tpoNumerico)
    pobjstc.Per_Adic_Form_Pag = M46V999.gfTratarNulo(prst("Per_Adic_Form_Pag"), tpoNumerico)
    pobjstc.Per_Desc_Form_Pag = M46V999.gfTratarNulo(prst("Per_Desc_Form_Pag"), tpoNumerico)
    pobjstc.Val_IS = M46V999.gfTratarNulo(prst("Val_IS"), tpoNumerico)
    pobjstc.Val_IS_2 = M46V999.gfTratarNulo(prst("Val_IS_2"), tpoNumerico)
    pobjstc.VAL_PR = M46V999.gfTratarNulo(prst("Val_PR"), tpoNumerico)
    pobjstc.Val_Pr_Ref_Tarif = M46V999.gfTratarNulo(prst("Val_Pr_Ref_Tarif"), tpoNumerico)
    pobjstc.Val_Pr_Ref_Resg = M46V999.gfTratarNulo(prst("Val_Pr_Ref_Resg"), tpoNumerico)
    pobjstc.Dat_Ini_Viag = M46V999.gfTratarNulo(prst("Dat_Ini_Viag"), tpoNumerico)
    pobjstc.Cod_Tip_Liber = M46V999.gfTratarNulo(prst("Cod_Tip_Liber"), tpoNumerico)
    pobjstc.Cod_Iden_Alt_Corr = M46V999.gfTratarNulo(prst("Cod_Iden_Alt_Corr"), tpoString)
    pobjstc.Cod_User_Digit = M46V999.gfTratarNulo(prst("Cod_User_Digit"), tpoString)
    pobjstc.DAT_DIGIT = M46V999.gfTratarNulo(prst("Dat_Digit"), tpoNumerico)
    pobjstc.HOR_DIGIT = M46V999.gfTratarNulo(prst("Hor_Digit"), tpoNumerico)
    pobjstc.Cod_User_Liber = M46V999.gfTratarNulo(prst("Cod_User_Liber"), tpoString)
    pobjstc.Dat_Liber = M46V999.gfTratarNulo(prst("Dat_Liber"), tpoNumerico)
    pobjstc.Hor_Liber = M46V999.gfTratarNulo(prst("Hor_Liber"), tpoNumerico)
    pobjstc.Cod_Sit_Ped = M46V999.gfTratarNulo(prst("Cod_Sit_Ped"), tpoNumerico)
    pobjstc.Cod_Tip_Doc = M46V999.gfTratarNulo(prst("Cod_Tip_Doc"), tpoNumerico)
    pobjstc.Cod_Dizer = M46V999.gfTratarNulo(prst("Cod_Dizer"), tpoNumerico)
    pobjstc.Qtd_Copia = M46V999.gfTratarNulo(prst("Qtd_Copia"), tpoNumerico)
    pobjstc.Val_Cust_Emis = M46V999.gfTratarNulo(prst("Val_Cust_Emis"), tpoNumerico)
    pobjstc.Val_Juro = M46V999.gfTratarNulo(prst("Val_Juro"), tpoNumerico)
    pobjstc.Val_IOF = M46V999.gfTratarNulo(prst("Val_IOF"), tpoNumerico)
    pobjstc.Num_Prop = M46V999.gfTratarNulo(prst("Num_Prop"), tpoNumerico)
    pobjstc.Cod_Posi_Atual = M46V999.gfTratarNulo(prst("Cod_Posi_Atual"), tpoString)
    pobjstc.Dat_Mov_Ini = M46V999.gfTratarNulo(prst("Dat_Mov_Ini"), tpoNumerico)
    pobjstc.Dat_Mov_Ter = M46V999.gfTratarNulo(prst("Dat_Mov_Ter"), tpoNumerico)
    pobjstc.Tip_Cust_Emis = M46V999.gfTratarNulo(prst("Tip_Cust_Emis"), tpoNumerico)
    pobjstc.Val_Cambio_Emis = M46V999.gfTratarNulo(prst("Val_Cambio_Emis"), tpoNumerico)
    pobjstc.Iden = M46V999.gfTratarNulo(prst("Iden"), tpoString)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)

    'pObjStc.Cod_Produto = M46V999.gfTratarNulo(prst("Cod_Produto"), tpoNumerico)
    'pObjStc.Cod_Sequencia = M46V999.gfTratarNulo(prst("Cod_Sequencia"), tpoNumerico)
    'pObjStc.Cod_Protocolo = M46V999.gfTratarNulo(prst("Cod_Protocolo"), tpoNumerico)
    'pObjStc.Tip_Pessoa = M46V999.gfTratarNulo(prst("Tip_Pessoa"), tpoNumerico)
    'pObjStc.Cod_Ramo_Ativ = M46V999.gfTratarNulo(prst("Cod_Ramo_Ativ"), tpoNumerico)
    'pObjStc.Cod_Doc = M46V999.gfTratarNulo(prst("Cod_Doc"), tpoNumerico)
    'pObjStc.Num_Doc = M46V999.gfTratarNulo(prst("Num_Doc"), tpoString)
    'pObjStc.Org_Exp = M46V999.gfTratarNulo(prst("Org_Exp"), tpoString)
    'pObjStc.Dat_Emis_Doc = M46V999.gfTratarNulo(prst("Dat_Emis_Doc"), tpoNumerico)
    'pObjStc.Nom_Empresa = M46V999.gfTratarNulo(prst("Nom_Empresa"), tpoString)
    'pObjStc.Sexo = M46V999.gfTratarNulo(prst("Sexo"), tpoString)
    'pObjStc.Dat_Pri_Calculo = M46V999.gfTratarNulo(prst("Dat_Pri_Calculo"), tpoNumerico)
    'pObjStc.Ult_Tarifa = M46V999.gfTratarNulo(prst("Ult_Tarifa"), tpoNumerico)
    'pObjStc.Cod_Dig_Agen_Bas_Cor = M46V999.gfTratarNulo(prst("Cod_Dig_Agen_Bas_Cor"), tpoNumerico)
    'pObjStc.Cod_Ver_Unici = M46V999.gfTratarNulo(prst("Cod_Ver_Unici"), tpoNumerico)
    'pObjStc.Val_RendaFatu = M46V999.gfTratarNulo(prst("Val_RendaFatu"), TpoValor)
    'pObjStc.Per_Desc_LMI = M46V999.gfTratarNulo(prst("Per_Desc_LMI"), TpoValor)
    'pObjStc.LMI_Unico = M46V999.gfTratarNulo(prst("LMI_Unico"), TpoValor)
    'pObjStc.Num_Cotac_LMI = M46V999.gfTratarNulo(prst("Num_Cotac_LMI"), tpoString)
    'pObjStc.Ind_CondGerais = M46V999.gfTratarNulo(prst("Ind_CondGerais"), tpoNumerico)
    'pObjStc.Ind_Reentrada = M46V999.gfTratarNulo(prst("Ind_Reentrada"), tpoNumerico)
    'pObjStc.Nosso_Numero_Original = M46V999.gfTratarNulo(prst("Nosso_Numero_Original"), tpoNumerico)
    'If pObjStc.Ind_Reentrada <> 1 Then
    '    pObjStc.Nosso_Numero_Original = ""
    'End If
End Sub



Public Function mfComparareCalcular(ByRef pobjstcItem As stcA46V708B, _
                                    ByRef pobjstcPedido As stcA46V702B, _
                                    ByRef pobjstcPedidoApolice As stcA46V702B, _
                                    ByRef pMensagem As String, _
                                    ByRef pPropTabela As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfComparareCalcular"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim StringEndossos As String
    Dim ValorLMGApol As Double
    Dim lCodCobert As String
    Dim lAlterouAss24hr As Boolean

    Dim pobjstcItemApol As stcA46V708B
    Dim lobjCobert As clsA46V704A
    Dim lobjstcEndereco As stcA46V707B   'Objeto da classe clsA46V707A (TAB_PED_ENDER)
    Dim lobjstcCorretor As stcA46V705B   'Objeto da classe clsA46V705A (TAB_PED_CORR)
    Dim lobjStcClausula As stcA46V703B
    Dim lobjstcCobert As stcA46V704B
    Dim lobjStcFatRisco As stcA46V720B
    Dim lobjStcLocBem As stcA46V715B
    Dim lobjstcTextoEspec As stcA46V726B    'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).
    Dim lobjstcTextoEspecAnt As stcA46V726B    'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).
    Dim StringCobertApol As String
    Dim StringCobertProp As String


    Dim ArrayEndosso As Variant

    Dim CalculaTodos As Boolean
    Dim CalculaItem As Boolean

    Dim lAuxiliar As String

    Dim i As Integer
    Dim jApol As Integer
    Dim lArrayCalcular() As Variant

    Dim lCoePrazo As Single
    Dim ColecaoAux As Collection

    Dim lbdTab_Prazo As ADODB.Connection    'Arquivo: P00Tarifa.
    Dim lrsTab_Prazo As clsYasRecordSet    'Registro : Tab_Prazo.

    Dim lobjstcCobertAtual As stcA46V704B
    Dim lobjstcCobertAnterior As stcA46V704B

    Dim CalculaTodaProposta As Boolean
    Dim CalculaTodoItem As Boolean

    Dim lQtdiasFalta As Integer
    Dim lQtdiasUtilizados As Integer

    Dim lPremio As Double
    Dim lSelect As String

    Dim lobjstcClauAtual As stcA46V703B
    Dim lobjstcClauAnterior As stcA46V703B

    Dim lInclusaoBeneficiario As Boolean
    Dim lExclusaoBeneficiario As Boolean

    On Error GoTo ControleErro
    mfComparareCalcular = False

    'C�lculo Pontual
    '    e_AumentodeLMG = 1
    '    e_AlteracaodeVR = 2
    '    e_ReducaodeLMG = 3
    '    e_Inclusaodecobertura = 5
    '    e_Exclusaodecobertura = 6
    '    e_InclusaodeLocaldeRisco = 8


    'Calculo Devolu��o
    '    e_ExclusaodeLocaldeRisco = 9
    '    e_Cancelamentodaapolice = 25


    'Rec�lculo Total
    '    e_AlteracaodePlano = 4
    '    e_AlteracaodeLocaldeRisco = 7
    '    e_Alteracaodeclassificacao = 11
    '    e_InclusaodasClausulas201ou202MSE = 18
    '    e_AlteracaodeFranquia = 22
    '    e_AlteracaodoPeriodoIndenitario = 23

    'SemRecalculo8
    'e_Retificacaodeenderecodecorrespondencia = 10
    'e_AlteracaodeCPFouCNPJ = 24

    StringEndossos = ""
    '   TipoEndossoPrincipal = ""

    CalculaTodos = False
    CalculaItem = False

    If pobjstcItem.Cod_DizerItem = 1 Then
        Call gDescEndosso(e_InclusaodeLocaldeRisco, pobjstcItem.Num_Item, 0)
    End If
    If pobjstcItem.Cod_DizerItem = 4 Then
        Call gDescEndosso(e_ExclusaodeLocaldeRisco, pobjstcItem.Num_Item, 0)
    End If


    If pobjstcPedido.objstcPedEndereco.Num_CEP <> pobjstcPedidoApolice.objstcPedEndereco.Num_CEP Or _
       Trim(pobjstcPedido.objstcPedEndereco.Dsc_Compl_Logr) <> Trim(pobjstcPedidoApolice.objstcPedEndereco.Dsc_Compl_Logr) Or _
       pobjstcPedido.objstcPedEndereco.Nom_Logr <> pobjstcPedidoApolice.objstcPedEndereco.Nom_Logr Or _
       pobjstcPedido.objstcPedEndereco.Nom_Bairro <> pobjstcPedidoApolice.objstcPedEndereco.Nom_Bairro Or _
       pobjstcPedido.objstcPedEndereco.Nom_Cid <> pobjstcPedidoApolice.objstcPedEndereco.Nom_Cid Or _
       pobjstcPedido.objstcPedEndereco.Sig_UF <> pobjstcPedidoApolice.objstcPedEndereco.Sig_UF Then
        Call gDescEndosso(e_Retificacaodeenderecodecorrespondencia, pobjstcItem.Num_Item, 0)
    End If

    If pobjstcPedido.Num_Cgc_Cpf <> pobjstcPedidoApolice.Num_Cgc_Cpf Then
        Call gDescEndosso(e_AlteracaodeCPFouCNPJ, pobjstcItem.Num_Item, 0)
    End If
    
    'Inclus�o de descri��o de endosso para altera��es de dados pessoais
    'Aguardar informa��es para libera��o do ajuste abaixo
'    If pobjstcPedido.Tip_Pessoa = e_TipPessoa_Fisica Then
'        If pobjstcPedido.Cod_Doc <> pobjstcPedidoApolice.Cod_Doc Or _
'            pobjstcPedido.Num_Doc <> pobjstcPedidoApolice.Num_Doc Or _
'            pobjstcPedido.Org_Exp <> pobjstcPedidoApolice.Org_Exp Or _
'            pobjstcPedido.Dat_Emis_Doc <> pobjstcPedidoApolice.Dat_Emis_Doc Or _
'            pobjstcPedido.objstcCtrlEmis.DAT_Nascimento <> pobjstcPedidoApolice.objstcCtrlEmis.DAT_Nascimento Or _
'            pobjstcPedido.objstcCtrlEmis.Estado_Civil <> pobjstcPedidoApolice.objstcCtrlEmis.Estado_Civil Or _
'            Trim(pobjstcPedido.Sexo) <> Trim(pobjstcPedidoApolice.Sexo) Or _
'            pobjstcPedido.objstcPedEndereco.Cod_Reg_Cons <> pobjstcPedidoApolice.objstcPedEndereco.Cod_Reg_Cons Or _
'            Trim(pobjstcPedido.objstcPedEndereco.Num_Reg_Cons) <> Trim(pobjstcPedidoApolice.objstcPedEndereco.Num_Reg_Cons) Then
'            Call gDescEndosso(e_AlteracaoDadosPessoais, pobjstcItem.Num_Item, 0)
'        End If
'    Else
'        If pobjstcPedido.Cod_Ramo_Ativ <> pobjstcPedidoApolice.Cod_Ramo_Ativ Then
'            Call gDescEndosso(e_AlteracaoDadosPessoais, pobjstcItem.Num_Item, 0)
'        End If
'    End If
        

    For Each lobjstcCorretor In pobjstcPedidoApolice.ColecaoPedCorr
        If lobjstcCorretor.Num_Ordem_Corr = 1 Then
            If lobjstcCorretor.Cod_Corr <> pobjstcPedido.objstcCtrlEmis.Cod_Corr Then
                Call gDescEndosso(e_Transferenciadecorretagem, pobjstcItem.Num_Item, 0)
            End If
        ElseIf lobjstcCorretor.Num_Ordem_Corr = 2 Then
            If lobjstcCorretor.Cod_Corr <> pobjstcPedido.objstcCtrlEmis.Cod_Corr_2 Then
                Call gDescEndosso(e_Transferenciadecorretagem, pobjstcItem.Num_Item, 0)
            End If
        Else
            If lobjstcCorretor.Cod_Corr <> pobjstcPedido.objstcCtrlEmis.Cod_Corr_3 Then
                Call gDescEndosso(e_Transferenciadecorretagem, pobjstcItem.Num_Item, 0)
            End If
        End If
    Next lobjstcCorretor
    

'    'Verifica Inclus�o/Exclus�o de Benefici�rios
'    For Each pobjstcItemApol In pobjstcPedidoApolice.ColecaoPedLoc
'        If pobjstcItem.ColecaoBeneficiario.Count > pobjstcItemApol.ColecaoBeneficiario.Count Then
'            Call gDescEndosso(e_InclusaodeClausulaBeneficiaria, pobjstcItem.Num_Item, 0)
'        Else
'            If pobjstcItem.ColecaoBeneficiario.Count < pobjstcItemApol.ColecaoBeneficiario.Count Then
'                Call gDescEndosso(e_ExclusaodaClausulaBeneficiaria, pobjstcItem.Num_Item, 0)
'            End If
'        End If
'    Next
    

    'Verifica Inclus�o/Exclus�o de Clausula Benefici�ria
    lInclusaoBeneficiario = False
    lExclusaoBeneficiario = False
    For Each lobjstcTextoEspec In pobjstcPedido.ColecaoTextoEspec
        Select Case lobjstcTextoEspec.Cod_Espec
        Case 8500, 8509, 8501, 8502, 8512
            lInclusaoBeneficiario = True
        End Select
        For Each lobjstcTextoEspecAnt In pobjstcPedidoApolice.ColecaoTextoEspec
            If lobjstcTextoEspec.Cod_Espec = lobjstcTextoEspecAnt.Cod_Espec Then
                Select Case lobjstcTextoEspecAnt.Cod_Espec
                Case 8500, 8509, 8501, 8502, 8512
                    If lInclusaoBeneficiario = True Then
                        lInclusaoBeneficiario = False
                    End If
                End Select
            End If
        Next lobjstcTextoEspecAnt

        If lInclusaoBeneficiario = True Then
            Call gDescEndosso(e_InclusaodeClausulaBeneficiaria, pobjstcItem.Num_Item, 0)
            Exit For
        End If
    Next lobjstcTextoEspec


    For Each lobjstcTextoEspecAnt In pobjstcPedidoApolice.ColecaoTextoEspec
        Select Case lobjstcTextoEspecAnt.Cod_Espec
        Case 8500, 8509, 8501, 8502, 8512
            lExclusaoBeneficiario = True
        End Select
        For Each lobjstcTextoEspec In pobjstcPedido.ColecaoTextoEspec
            If lobjstcTextoEspec.Cod_Espec = lobjstcTextoEspecAnt.Cod_Espec Then
                Select Case lobjstcTextoEspec.Cod_Espec
                Case 8500, 8509, 8501, 8502, 8512
                    If lExclusaoBeneficiario = True Then
                        lExclusaoBeneficiario = False
                    End If
                End Select
            End If
        Next lobjstcTextoEspec

        If lExclusaoBeneficiario = True Then
            Call gDescEndosso(e_ExclusaodaClausulaBeneficiaria, pobjstcItem.Num_Item, 0)
            Exit For
        End If
    Next lobjstcTextoEspecAnt

    '02-07-14
    'Altera��o do texto da clausula beneficiaria
    For Each lobjstcTextoEspecAnt In pobjstcPedidoApolice.ColecaoTextoEspec

        If lobjstcTextoEspecAnt.Cod_Espec = 8500 Then

            For Each lobjstcTextoEspec In pobjstcPedido.ColecaoTextoEspec
                If lobjstcTextoEspec.Cod_Espec = 8500 Then
                    If lobjstcTextoEspecAnt.Num_pag = lobjstcTextoEspec.Num_pag Then

                        If lobjstcTextoEspecAnt.Espec_lin1 <> lobjstcTextoEspec.Espec_lin1 Or _
                           lobjstcTextoEspecAnt.Espec_lin2 <> lobjstcTextoEspec.Espec_lin2 Or _
                           lobjstcTextoEspecAnt.Espec_lin3 <> lobjstcTextoEspec.Espec_lin3 Or _
                           lobjstcTextoEspecAnt.Espec_lin4 <> lobjstcTextoEspec.Espec_lin4 Or _
                           lobjstcTextoEspecAnt.Espec_lin5 <> lobjstcTextoEspec.Espec_lin5 Or _
                           lobjstcTextoEspecAnt.Espec_lin6 <> lobjstcTextoEspec.Espec_lin6 Or _
                           lobjstcTextoEspecAnt.Espec_lin7 <> lobjstcTextoEspec.Espec_lin7 Or _
                           lobjstcTextoEspecAnt.Espec_lin8 <> lobjstcTextoEspec.Espec_lin8 Or _
                           lobjstcTextoEspecAnt.Espec_lin9 <> lobjstcTextoEspec.Espec_lin9 Or _
                           lobjstcTextoEspecAnt.Espec_lin10 <> lobjstcTextoEspec.Espec_lin10 Or _
                           lobjstcTextoEspecAnt.Espec_lin11 <> lobjstcTextoEspec.Espec_lin11 Or _
                           lobjstcTextoEspecAnt.Espec_lin12 <> lobjstcTextoEspec.Espec_lin12 Or _
                           lobjstcTextoEspecAnt.Espec_lin13 <> lobjstcTextoEspec.Espec_lin13 Or _
                           lobjstcTextoEspecAnt.Espec_lin14 <> lobjstcTextoEspec.Espec_lin14 Or _
                           lobjstcTextoEspecAnt.Espec_lin15 <> lobjstcTextoEspec.Espec_lin15 Or _
                           lobjstcTextoEspecAnt.Espec_lin16 <> lobjstcTextoEspec.Espec_lin16 Or _
                           lobjstcTextoEspecAnt.Espec_lin17 <> lobjstcTextoEspec.Espec_lin17 Or _
                           lobjstcTextoEspecAnt.Espec_lin18 <> lobjstcTextoEspec.Espec_lin18 Or _
                           lobjstcTextoEspecAnt.Espec_lin19 <> lobjstcTextoEspec.Espec_lin19 Or _
                           lobjstcTextoEspecAnt.Espec_lin20 <> lobjstcTextoEspec.Espec_lin20 Then

                            Call gDescEndosso(e_Alteracaodeclausulabeneficiaria, pobjstcItem.Num_Item, 0)
                            Exit For

                        End If
                    End If
                End If

            Next lobjstcTextoEspec

        End If

    Next lobjstcTextoEspecAnt

    lQtdiasFalta = M46V999.M46V103.mfDifDatas(pobjstcPedido.Dat_Ini_Vig, pobjstcPedido.Dat_Ter_Vig)
    lQtdiasUtilizados = M46V999.M46V103.mfDifDatas(pobjstcPedidoApolice.Dat_Ini_Vig, pobjstcPedido.Dat_Ini_Vig)

    Set pobjstcItemApol = New stcA46V708B

    CalculaItem = False

    If pobjstcItem.Cod_DizerItem = e_CodDizerInclusao Then
        Call gDescEndosso(e_InclusaodeLocaldeRisco, pobjstcItem.Num_Item, 0)
        CalculaItem = True
    End If

    If pobjstcItem.Cod_DizerItem = e_CodDizerexclusao Then
        Call gDescEndosso(e_ExclusaodeLocaldeRisco, pobjstcItem.Num_Item, 0)
        CalculaItem = True
    End If
    
    'Ao Alterar a vig�ncia de Pro Rata para Prazo Curto ou vice-versa o pr�mio deve ser alterado.
    lCoePrazo = M46V999.gfTrunca(lQtdiasUtilizados / 365, 5)
    
'    If pobjstcItem.Cod_DizerItem = e_CodDizerCancelamento Then
'        lSelect = "SELECT Val_Coef FROM TAB_PRAZO_CURTO WHERE Qtd_Dias_Ini <= " & lQtdiasUtilizados & _
'                " AND Qtd_Dias_Fim >= " & lQtdiasUtilizados
'
'        If pMensagem = "" Then
'            lSelect = lSelect & " and Posicao = 1"
'        End If
'
'        If M46V999.gfAbrBasDados("P00MULTI2", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTab_Prazo, pMensagem) = False Then
'            Call M46V999.gpGraLog(2, pMensagem)
'            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
'            Exit Function
'        End If
'
'        If M46V999.gfObtRegistro(lbdTab_Prazo, lSelect, lrsTab_Prazo, pMensagem) = False Then
'            If InStr(1, pMensagem, "-2147217904") = 0 Then
'                Call M46V999.gpGraLog(2, pMensagem)
'                lCoePrazo = 0
'            End If
'        Else
'            If lrsTab_Prazo.EOF = True Then
'                Call M46V999.gpGraLog(2, "FNL0019 - Coeficiente de prazo curto n�o localizado (quantidade " & _
'                                         "de dias de vig�ncia = " & lQtdiasUtilizados & ").")
'                lCoePrazo = 0
'            Else
'                lCoePrazo = lrsTab_Prazo!Val_Coef / 100
'            End If
'        End If
'
'    End If
    
    Dim AchouItem As Boolean
    AchouItem = False

    If pobjstcItem.Cod_DizerItem <> e_CodDizerInclusao Then
        For Each pobjstcItemApol In pobjstcPedidoApolice.ColecaoPedLoc
            lAlterouAss24hr = False
            If pobjstcItem.Num_Item = pobjstcItemApol.Num_Item Then

                For Each lobjstcClauAtual In pobjstcItem.ColecaoPedClau
                    For Each lobjstcClauAnterior In pobjstcItemApol.ColecaoPedClau
                        If lobjstcClauAnterior.Cod_Clau = lobjstcClauAtual.Cod_Clau Then
                            GoTo ProximaClau
                        End If
                    Next lobjstcClauAnterior
                    If lobjstcClauAtual.Cod_Clau >= 201 And lobjstcClauAtual.Cod_Clau <= 202 Then
                        Call gDescEndosso(e_InclusaodasClausulas201ou202MSE, pobjstcItem.Num_Item, lobjstcClauAtual.Cod_Clau)
                    ElseIf lobjstcClauAtual.Cod_Clau >= 301 And lobjstcClauAtual.Cod_Clau <= 317 And lobjstcClauAtual.Cod_Ramo = 112 Then
                        Call gDescEndosso(e_InclusaodasClausulas301ate317MSE, pobjstcItem.Num_Item, lobjstcClauAtual.Cod_Clau)
                    ElseIf lobjstcClauAtual.Cod_Clau >= 301 And lobjstcClauAtual.Cod_Clau <= 301 And lobjstcClauAtual.Cod_Ramo = 113 Then
                        Call gDescEndosso(e_Inclusaoclausulas301a304MSR, pobjstcItem.Num_Item, lobjstcClauAtual.Cod_Clau)
                    ElseIf lobjstcClauAtual.Cod_Clau >= 400 And lobjstcClauAtual.Cod_Clau <= 405 Then
                        Call gDescEndosso(e_AlteracaoClausulaAss24hr, pobjstcItem.Num_Item, lobjstcClauAtual.Cod_Clau)
                        lAlterouAss24hr = True
                    End If

ProximaClau:
                Next lobjstcClauAtual

                'Valida��o de altera��o na lista de bicicletas
                Dim lobjstcBicicleta As stcA46V744B
                Dim lobjstcBicicletaApol As stcA46V744B
                
                'Exclus�o
                If pobjstcItemApol.ColecaoPedLocBicicleta.Count > pobjstcItem.ColecaoPedLocBicicleta.Count Then
                    Call gDescEndosso(e_ExclusaoRelacaoBicicletas, pobjstcItem.Num_Item, 0)
                Else
                    'Inclus�o
                    If pobjstcItemApol.ColecaoPedLocBicicleta.Count < pobjstcItem.ColecaoPedLocBicicleta.Count Then
                        Call gDescEndosso(e_InclusaoRelacaoBicicletas, pobjstcItem.Num_Item, 0)
                    Else
                        'Altera��o
                        For Each lobjstcBicicleta In pobjstcItem.ColecaoPedLocBicicleta
                            For Each lobjstcBicicletaApol In pobjstcItemApol.ColecaoPedLocBicicleta
                                If lobjstcBicicleta.Cod_Seq = lobjstcBicicletaApol.Cod_Seq And _
                                   (lobjstcBicicleta.Descricao <> lobjstcBicicletaApol.Descricao Or _
                                    lobjstcBicicleta.Valor <> lobjstcBicicletaApol.Valor) Then
                                    Call gDescEndosso(e_AlteracaoRelacaoBicicletas, pobjstcItem.Num_Item, 0)
                                    Exit For
                                End If
                            Next
                        Next
                    End If
                End If
                
                AchouItem = True
                If pobjstcItemApol.Cod_Produto = 302 Then pobjstcItemApol.Cod_Produto = 1
                If pobjstcItem.Cod_Produto <> pobjstcItemApol.Cod_Produto Then
                    Call gDescEndosso(e_AlteracaodePlano, pobjstcItem.Num_Item, 0)
                    CalculaTodos = True
                End If

                If pobjstcItem.Num_Cep_Risco <> pobjstcItemApol.Num_Cep_Risco Or _
                   pobjstcItem.Nom_Logr_Risco <> pobjstcItemApol.Nom_Logr_Risco Or _
                   pobjstcItem.Nom_Bairro_Risco <> pobjstcItemApol.Nom_Bairro_Risco Then
                    Call gDescEndosso(e_AlteracaodeLocaldeRisco, pobjstcItem.Num_Item, 0)
                    CalculaItem = True
                Else
                    lAuxiliar = pobjstcItemApol.Dsc_Compl_Risco
                    If Left(lAuxiliar, 2) = "N." Then
                        lAuxiliar = Trim(Mid(lAuxiliar, InStr(1, lAuxiliar, "-") + 1))
                    End If
                    If pobjstcItem.Dsc_Compl_Risco <> lAuxiliar Then
                        Call gDescEndosso(e_AlteracaodeLocaldeRisco, pobjstcItem.Num_Item, 0)
                        CalculaItem = True
                    Else
                        If pobjstcItem.objstcPedInsp.Num_Logr_Risco <> pobjstcItemApol.objstcPedInsp.Num_Logr_Risco And pobjstcItemApol.objstcPedInsp.Num_Logr_Risco <> 0 Then
                            Call gDescEndosso(e_AlteracaodeLocaldeRisco, pobjstcItem.Num_Item, 0)
                            CalculaItem = True
                        End If
                    End If


                End If

                If Trim(pobjstcItem.Cod_Clasf) <> Trim(pobjstcItemApol.Cod_Clasf) Or _
                   pobjstcItem.Cod_Compl_Clasf <> pobjstcItemApol.Cod_Compl_Clasf Then
                    Call gDescEndosso(e_Alteracaodeclassificacao, pobjstcItem.Num_Item, 0)
                    CalculaItem = True
                End If



                'Compara Dados Cobertura da Apolice

                jApol = 1
                i = 1
                Set ColecaoAux = New Collection
                'Vai limpar o primeiro
                'If pobjstcItem.ColecaoPedCobert.Count > 0 Then pobjstcItem.ColecaoPedCobert(I).EndossoTipoRecalculo = "0"
                'EndossoTipoRecalculo = "0" - N�o sobreu altera��o
                'EndossoTipoRecalculo = "1" - Recalcula Alteracao
                'EndossoTipoRecalculo = "2" - inclusao
                'EndossoTipoRecalculo = "3" - Exclusao

                For Each lobjstcCobertAtual In pobjstcItem.ColecaoPedCobert
                    For Each lobjstcCobertAnterior In pobjstcItemApol.ColecaoPedCobert
                        If lobjstcCobertAtual.Cod_Cobert = lobjstcCobertAnterior.Cod_Cobert Then

                            lobjstcCobertAtual.EndossoTipoRecalculo = "0"

                            lobjstcCobertAtual.EndossoValPremioApol = lobjstcCobertAnterior.Val_Pr_Anu
                            If lobjstcCobertAtual.Cod_Cobert = 1 Then
                                If lAlterouAss24hr = True Then
                                    lobjstcCobertAtual.EndossoTipoRecalculo = "1"
                                End If
                                If Trim(pobjstcItem.Val_Vr) <> Trim(pobjstcItemApol.Val_Vr) Then
                                    Call gDescEndosso(e_AlteracaodeVR, pobjstcItem.Num_Item, lobjstcCobertAtual.Cod_Cobert)
                                    lobjstcCobertAtual.EndossoTipoRecalculo = "1"
                                End If
                            End If

                            If lobjstcCobertAtual.Val_IS > 0 And pobjstcItem.Cod_DizerItem <> 10 Then
                                If lobjstcCobertAnterior.Val_IS = 0 Then
                                    'Reintegra��o de IS
                                    lobjstcCobertAtual.EndossoTipoRecalculo = "2"
                                    lobjstcCobertAtual.EndossoValPremioApol = 0
                                    Call gDescEndosso(e_AumentodeLMG, pobjstcItem.Num_Item, lobjstcCobertAtual.Cod_Cobert)
                                Else
                                    If lobjstcCobertAtual.Val_IS <> lobjstcCobertAnterior.Val_IS Then
                                        If lobjstcCobertAtual.Val_IS > lobjstcCobertAnterior.Val_IS Then
                                            Call gDescEndosso(e_AumentodeLMG, pobjstcItem.Num_Item, lobjstcCobertAtual.Cod_Cobert)
                                        Else
                                            Call gDescEndosso(e_ReducaodeLMG, pobjstcItem.Num_Item, lobjstcCobertAtual.Cod_Cobert)
                                        End If
                                        lobjstcCobertAtual.EndossoTipoRecalculo = "1"
                                    End If
                                    If lobjstcCobertAtual.Num_Meses_PI <> lobjstcCobertAnterior.Num_Meses_PI Then
                                        Call gDescEndosso(e_AlteracaodoPeriodoIndenitario, pobjstcItem.Num_Item, lobjstcCobertAtual.Cod_Cobert)
                                        lobjstcCobertAtual.EndossoTipoRecalculo = "1"
                                    End If
                                    If M46V999.PerfilAcesso = e_Perfil_CALC And lobjstcCobertAtual.EndossoTipoRecalculo = "0" Then
                                        lobjstcCobertAtual.Val_Franq_Min = lobjstcCobertAnterior.Val_Franq_Min
                                        lobjstcCobertAtual.Per_Franq = lobjstcCobertAnterior.Per_Franq
                                    ElseIf lobjstcCobertAtual.Val_Franq_Min <> lobjstcCobertAnterior.Val_Franq_Min Or _
                                           lobjstcCobertAtual.Per_Franq <> lobjstcCobertAnterior.Per_Franq Then
                                        Call gDescEndosso(e_AlteracaodeFranquia, pobjstcItem.Num_Item, lobjstcCobertAtual.Cod_Cobert)
                                        lobjstcCobertAtual.EndossoTipoRecalculo = "1"
                                    End If
                                End If
                            ElseIf lobjstcCobertAnterior.Val_IS <> 0 Then
                                'Quando IS zerada ent�o a cobertura est� sendo exclu�da.
                                lobjstcCobertAtual.EndossoTipoRecalculo = 3
                                If pobjstcItem.Cod_DizerItem <> 10 Then
                                    lobjstcCobertAtual.Val_IS = 0
                                End If
                                lobjstcCobertAtual.Val_Pr_Anu = 0
                                lobjstcCobertAtual.EndossoValPremioApol = lobjstcCobertAnterior.VAL_PR
                                If pobjstcItem.Cod_DizerItem <> 10 Then
                                    Call gDescEndosso(e_EXclusaodecobertura, pobjstcItem.Num_Item, lobjstcCobertAtual.Cod_Cobert)
                                End If

                            End If
                            pobjstcItemApol.ColecaoPedCobert.Remove lobjstcCobertAnterior.Class_ID
                            GoTo Proxima
                        End If
                    Next lobjstcCobertAnterior
                    'N�o encontrou na ap�lice anterior ent�o � inclus�o de cobertura
                    lobjstcCobertAtual.EndossoTipoRecalculo = "2"
                    lobjstcCobertAtual.EndossoValPremioApol = 0
                    Call gDescEndosso(e_Inclusaodecobertura, pobjstcItem.Num_Item, lobjstcCobertAtual.Cod_Cobert)


Proxima:
                    ColecaoAux.Add lobjstcCobertAtual, lobjstcCobertAtual.Class_ID
                Next lobjstcCobertAtual

                For Each lobjstcCobertAnterior In pobjstcItemApol.ColecaoPedCobert
                    'Coberturas exclu�das
                    lobjstcCobertAnterior.nosso_numero = pobjstcItem.nosso_numero
                    lobjstcCobertAnterior.Num_Item = pobjstcItem.Num_Item
                    lobjstcCobertAnterior.NUM_PI = pobjstcItem.NUM_PI
                    lobjstcCobertAnterior.EndossoTipoRecalculo = "3"
                    lobjstcCobertAnterior.EndossoValPremioApol = lobjstcCobertAnterior.Val_Pr_Anu


                    Call gDescEndosso(e_EXclusaodecobertura, pobjstcItem.Num_Item, lobjstcCobertAnterior.Cod_Cobert)

                    Set lobjstcCobertAnterior.M46V999 = M46V999
                    Set lobjstcCobertAnterior.ObjPai = pobjstcItem
                    lobjstcCobertAnterior.Val_IS = 0
                    lobjstcCobertAnterior.Val_Pr_Anu = 0
                    'Ajuste calculo errado CRQ000000425129
                    lobjstcCobertAnterior.Num_Apol = 0
                    lobjstcCobertAnterior.Num_Endo = 0
                    lobjstcCobertAnterior.Num_Ped = pobjstcItem.Num_Ped


                    pobjstcItem.ColecaoPedCobert.Add lobjstcCobertAnterior, lobjstcCobertAnterior.Class_ID

                    Set lobjCobert = New clsA46V704A
                    Set lobjCobert.M46V999 = M46V999

                    If Not lobjCobert.mfIncluir(lobjstcCobertAnterior, pPropTabela, "", False, False, tagDynaTrace) Then
                        GoTo ControleErro
                    End If
                    ColecaoAux.Add lobjstcCobertAnterior
                Next lobjstcCobertAnterior


                Set pobjstcItem.ColecaoPedCobert = Nothing
                Set pobjstcItem.ColecaoPedCobert = ColecaoAux
                '                Do While True
                '                    If I <= pobjstcItem.ColecaoPedCobert.Count And jApol <= pobjstcItemApol.ColecaoPedCobert.Count Then
                '                        If pobjstcItem.ColecaoPedCobert(I).Cod_Cobert = pobjstcItemApol.ColecaoPedCobert(jApol).Cod_Cobert Then
                '
                '                            pobjstcItem.ColecaoPedCobert(I).EndossoValPremioApol = pobjstcItemApol.ColecaoPedCobert(jApol).Val_Pr_Anu
                '
                '                            If pobjstcItem.ColecaoPedCobert(I).Val_IS <> pobjstcItemApol.ColecaoPedCobert(jApol).Val_IS Then
                '                                pobjstcItem.ColecaoPedCobert(I).EndossoTipoRecalculo = "1"
                '                                If pobjstcItem.ColecaoPedCobert(I).Val_IS > pobjstcItemApol.ColecaoPedCobert(jApol).Val_IS Then
                '                                    StringEndossos = StringEndossos & e_AumentodeLMG & "@"
                '                                Else
                '                                    StringEndossos = StringEndossos & e_ReducaodeLMG & "@"
                '                                End If
                '                            End If
                '                            If pobjstcItem.ColecaoPedCobert(I).Val_Franq_Min <> pobjstcItemApol.ColecaoPedCobert(jApol).Val_Franq_Min Or _
                                             '                                pobjstcItem.ColecaoPedCobert(I).Per_Franq <> pobjstcItemApol.ColecaoPedCobert(jApol).Per_Franq Then
                '                                pobjstcItem.ColecaoPedCobert(I).EndossoTipoRecalculo = "1"
                '                                StringEndossos = StringEndossos & e_AlteracaodeFranquia & "@"
                '                            End If
                '
                '                            If pobjstcItem.ColecaoPedCobert(I).Val_Risco_Cob <> pobjstcItemApol.ColecaoPedCobert(jApol).Val_Risco_Cob Then
                '                               ' pobjstcItem.ColecaoPedCobert(i).EndossoTipoRecalculo = "0"
                '                               ' pobjstcItem.ColecaoPedCobert(i).EndossoValPremioApol = pobjstcItemApol.ColecaoPedCobert(jApol).Val_Pr_Anu
                '                            '    StringEndossos = StringEndossos & e_AlteracaodeVR & "@"
                '                            End If
                '                            I = I + 1
                '                            'Vai limpar o proximo
                '                            If I <= pobjstcItem.ColecaoPedCobert.Count Then pobjstcItem.ColecaoPedCobert(I).EndossoTipoRecalculo = "0"
                '                            jApol = jApol + 1
                '                        Else
                '                            If pobjstcItem.ColecaoPedCobert(I).Cod_Cobert < pobjstcItemApol.ColecaoPedCobert(jApol).Cod_Cobert Then
                '                                'Inclusao de Cobertura
                '                                pobjstcItem.ColecaoPedCobert(I).EndossoTipoRecalculo = "2"
                '                                pobjstcItem.ColecaoPedCobert(I).EndossoValPremioApol = 0
                '                                lPremio = (pobjstcItem.ColecaoPedCobert(I).Val_Pr_Anu) * lCoePrazo
                '                                StringEndossos = StringEndossos & e_Inclusaodecobertura & "@"
                '                                I = I + 1
                '                                'Vai limpar o proximo
                '                                If I <= pobjstcItem.ColecaoPedCobert.Count Then pobjstcItem.ColecaoPedCobert(I).EndossoTipoRecalculo = "0"
                '                            Else
                '                                'Exclusao de Cobertura
                '                                pobjstcItemApol.ColecaoPedCobert(jApol).EndossoTipoRecalculo = "3"
                '                                pobjstcItemApol.ColecaoPedCobert(jApol).EndossoValPremioApol = pobjstcItemApol.ColecaoPedCobert(jApol).Val_Pr_Anu
                '                                ColecaoAux.Add pobjstcItemApol.ColecaoPedCobert(jApol)
                '                                StringEndossos = StringEndossos & e_EXclusaodecobertura & "@"
                '                                jApol = jApol + 1
                '                            End If
                '                        End If
                '                    Else
                '                        If I <= pobjstcItem.ColecaoPedCobert.Count Then
                '                            'Inclusao de Cobertura
                '                            pobjstcItem.ColecaoPedCobert(I).EndossoTipoRecalculo = "2"
                '                            pobjstcItem.ColecaoPedCobert(I).EndossoValPremioApol = 0
                '                            I = I + 1
                '                            'Vai limpar o proximo
                '                            If I <= pobjstcItem.ColecaoPedCobert.Count Then pobjstcItem.ColecaoPedCobert(I).EndossoTipoRecalculo = "0"
                '                        Else
                '                            'Exclusao de Cobertura
                '                            pobjstcItemApol.ColecaoPedCobert(jApol).EndossoTipoRecalculo = "3"
                '                            pobjstcItemApol.ColecaoPedCobert(jApol).EndossoValPremioApol = pobjstcItemApol.ColecaoPedCobert(jApol).Val_Pr_Anu
                '                            ColecaoAux.Add pobjstcItemApol.ColecaoPedCobert(jApol)
                '                            jApol = jApol + 1
                '                        End If
                '                    End If
                '                    If I > pobjstcItem.ColecaoPedCobert.Count And jApol > pobjstcItemApol.ColecaoPedCobert.Count Then
                '                        Exit Do
                '                    End If
                '                Loop
                '
                '                If ColecaoAux.Count > 0 Then
                '                    For Each lobjstcCobert In ColecaoAux
                '                        Set lobjstcCobert.M46V999 = M46V999
                '                        Set lobjstcCobert.ObjPai = pobjstcItem
                '                        If lobjstcCobert.EndossoTipoRecalculo = 3 Then
                '                            lobjstcCobert.Val_IS = 0
                '                            lobjstcCobert.Val_Pr_Anu = 0
                '                        End If
                '                        lobjstcCobert.NOSSO_NUMERO = pobjstcItem.NOSSO_NUMERO
                '                        lobjstcCobert.Num_Item = pobjstcItem.Num_Item
                '                        lobjstcCobert.Num_PI = pobjstcItem.Num_PI
                '                        pobjstcItem.ColecaoPedCobert.Add lobjstcCobert, lobjstcCobert.Class_ID
                '                        Set ColecaoAux = New Collection
                '                        Set lobjCobert = New clsA46V704A
                '                        Set lobjCobert.M46V999 = M46V999
                '
                '                        If Not lobjCobert.mfIncluir(lobjstcCobert, pPropTabela, "") Then
                '                            GoTo ControleErro
                '                        End If
                '                    Next
                '                End If
                '                Exit For
            End If
        Next
    End If
    'Calcular

    If Not AchouItem And pobjstcItem.Cod_DizerItem = e_CodDizerAlteracao Then
        pobjstcItem.Cod_DizerItem = e_CodDizerInclusao
    End If



    Set ColecaoAux = New Collection
    Set lobjCobert = New clsA46V704A
    Set lobjCobert.M46V999 = M46V999


    For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert
        If pobjstcItem.Cod_DizerItem = e_CodDizerAlteracao Then
            If CalculaItem = False And CalculaTodos = False And lobjstcCobert.EndossoTipoRecalculo = 0 Then
                lPremio = 0
            ElseIf lobjstcCobert.EndossoValPremioApol = 0 Then  'Inclusao de cobertura
                lPremio = (lobjstcCobert.Val_Pr_Anu) * (lQtdiasFalta / 365)
            ElseIf lobjstcCobert.Val_Pr_Anu = 0 Then    'Exclusao de Cobertura
                lPremio = (lobjstcCobert.EndossoValPremioApol) * (lQtdiasFalta / 365) * -1
            Else
                lPremio = (Round(lobjstcCobert.Val_Pr_Anu, 2) - lobjstcCobert.EndossoValPremioApol) * (lQtdiasFalta / 365)
            End If
        End If
        If pobjstcItem.Cod_DizerItem = e_CodDizerInclusao Then
            lPremio = (lobjstcCobert.Val_Pr_Anu) * (lQtdiasFalta / 365)
        End If
        If pobjstcItem.Cod_DizerItem = e_CodDizerexclusao Then
            lPremio = (lobjstcCobert.EndossoValPremioApol) * (lQtdiasFalta / 365)
            lPremio = lPremio * -1
        End If
        If pobjstcItem.Cod_DizerItem = e_CodDizerCancelamento Then
            lPremio = (lobjstcCobert.EndossoValPremioApol * lCoePrazo) - lobjstcCobert.EndossoValPremioApol
        End If

        If lobjstcCobert.PremioInformado = True Then
            lobjstcCobert.Val_Pr_Orig = lPremio
            lPremio = lobjstcCobert.VAL_PR
        End If

        lobjstcCobert.VAL_PR = lPremio
        lobjstcCobert.Val_Pr_Tarif = lPremio
        lobjstcCobert.Val_Pr_Resg = lPremio

        If Not lobjCobert.mfAlterarValoresPrecificacao(lobjstcCobert, pPropTabela, "") Then
            pMensagem = lobjCobert.UltimoErro
        End If

ProximaCobertura:
        'Adiciono a cobertura em uma nova cole��o
        Set lobjstcCobert.ObjPai = pobjstcItem
        ColecaoAux.Add lobjstcCobert, lobjstcCobert.Class_ID

    Next lobjstcCobert
    ' CR3923 - Item 8 - Inicio
    ' Se endosso com restitui��o, dever� calcular na UPC
    If pobjstcItem.ValPrTotCober < 0 And pobjstcItem.ObjPai.Ind_CotacaoSolic = False Then
        If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
            'Call M46V999.gpMenErro("N�o foi poss�vel finalizar a cota��o de endosso. Favor clicar em SOLICITAR COTA��O e aguardar o retorno.")
            Call M46V999.gpMenErro("N�o foi poss�vel realizar a cota��o de endosso. Favor clicar em SOLICITAR COTA��O e aguardar o retorno.")
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "N�o foi poss�vel realizar a cota��o de endosso. Favor clicar em SOLICITAR COTA��O e aguardar o retorno.", 2, Class_Name, methodName)
            Exit Function
        End If
    End If
    ' CR3923 - Item 8 - Fim

    Set pobjstcItem.ColecaoPedCobert = Nothing
    Set pobjstcItem.ColecaoPedCobert = ColecaoAux


    'e_InclusaodeClausulaBeneficiaria = 12
    'e_ExclusaodaClausulaBeneficiaria = 13
    'e_Retificacaodeclausulabeneficiaria = 14
    'e_Alteracaodeclausulabeneficiaria = 15
    'e_InclusaodeRelacaodeObjetosBens = 16
    'e_Exclusaodeobjetosnarelacaoexistente = 17
    'e_InclusaodasClausulas301ate317MSE = 19
    'e_Inclusaoclausulas301a304MSR = 20
    'e_Transferenciadecorretagem = 21
    'e_Inclusaodeprotecional = 26

    mfComparareCalcular = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    pMensagem = Err & " " & Err.Description
    Call M46V999.gpGraLog(1, "mfComparareCalcular ERRO - " & pMensagem)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
    GoTo SAIDA
FechaObj:
    Set pobjstcItemApol = Nothing
    Set lobjCobert = Nothing
    Set lobjstcEndereco = Nothing
    Set lobjstcCorretor = Nothing
    Set lobjStcClausula = Nothing
    Set lobjstcCobert = Nothing
    Set lobjStcFatRisco = Nothing
    Set lobjStcLocBem = Nothing
    Set lobjstcTextoEspec = Nothing
    Set lobjstcCobertAtual = Nothing
    Set lobjstcCobertAnterior = Nothing
    Set lobjstcClauAtual = Nothing
    Set lobjstcClauAnterior = Nothing
    Set lrsTab_Prazo = Nothing
    Call M46V999.gpFecha2(lbdTab_Prazo)
    Set lbdTab_Prazo = Nothing
    Set ColecaoAux = Nothing
    Set lobjstcTextoEspecAnt = Nothing
    Return
End Function
Private Sub gDescEndosso(pAlt As E_TipoEndosso, pItem As Long, pCobertura As Long)

    Select Case pAlt
    Case e_AumentodeLMG
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Aumento de LMI da cobertura " & pCobertura & vbCrLf
    Case e_AlteracaodeVR
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Altera��o do Valor de Risco " & vbCrLf
    Case e_ReducaodeLMG
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Redu��o de LMI da coberta " & pCobertura & vbCrLf
    Case e_AlteracaodePlano
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Altera��o do Plano " & vbCrLf
    Case e_Inclusaodecobertura
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Inclusao da Cobertura " & pCobertura & vbCrLf
    Case e_EXclusaodecobertura
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Exclusao da Cobertura " & pCobertura & vbCrLf
    Case e_AlteracaodeLocaldeRisco
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Altera��o do Local de Risco " & vbCrLf
    Case e_InclusaodeLocaldeRisco
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Inclus�o do Local de Risco " & vbCrLf
    Case e_ExclusaodeLocaldeRisco
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Exclus�o do Local de Risco " & vbCrLf
    Case e_Retificacaodeenderecodecorrespondencia
        If InStr(1, M46V999.M46V111.gCod_DizerTexto, " - Retifica��o do Endere�o de Correspond�ncia ") = 0 Then
            M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & " - Retifica��o do Endere�o de Correspond�ncia " & vbCrLf
        End If
    Case e_Alteracaodeclassificacao
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Altera��o de Classifica��o  " & vbCrLf
    Case e_InclusaodeClausulaBeneficiaria
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Inclus�o de Cl�usula e/ou Benefici�rio " & vbCrLf
    Case e_ExclusaodaClausulaBeneficiaria
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Exclus�o de Cl�usula e/ou Benefici�rio  " & vbCrLf
    Case e_Retificacaodeclausulabeneficiaria
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Retifica��o de Cl�usula e/ou Benefici�rio  " & vbCrLf
    Case e_Alteracaodeclausulabeneficiaria
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Altera��o de Cl�usula e/ou Benefici�rio  " & vbCrLf
    Case e_InclusaodeRelacaodeObjetosBens
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Inclus�o de Objetos na Rela��o de Bens " & vbCrLf
    Case e_Exclusaodeobjetosnarelacaoexistente
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Exclus�o de Objetos na Rela��o de Bens " & vbCrLf
    Case e_InclusaodasClausulas201ou202MSE
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Inclus�o de Cl�usula (201 ou 202):" & pCobertura & vbCrLf
    Case e_InclusaodasClausulas301ate317MSE
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Inclus�o de Cl�usula (301 at� 317): " & pCobertura & vbCrLf
    Case e_Inclusaoclausulas301a304MSR
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Inclus�o de Cl�usula (301 a 304): " & pCobertura & vbCrLf
    Case e_Transferenciadecorretagem
        If InStr(1, M46V999.M46V111.gCod_DizerTexto, " - Transferencia de Corretagem  ") = 0 Then
            M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Transferencia de Corretagem  " & vbCrLf
        End If
    Case e_AlteracaodeFranquia
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Altera��o de Franquia da cobertura " & pCobertura & vbCrLf
    Case e_AlteracaodoPeriodoIndenitario
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Altera��o de Per�odo Indenit�rio da cobertura " & pCobertura & vbCrLf
    Case e_AlteracaodeCPFouCNPJ
        If InStr(1, M46V999.M46V111.gCod_DizerTexto, " - Altera��o de CPF ou CNPJ ") = 0 Then
            M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Altera��o de CPF ou CNPJ " & vbCrLf
        End If
    Case e_Cancelamentodaapolice
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Cancelamento da Ap�lice " & vbCrLf
    Case e_Inclusaodeprotecional
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Inclus�o de Protecional " & vbCrLf
    Case e_AlteracaoClausulaAss24hr
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Altera��o da cl�usula de assist�ncia 24hr." & vbCrLf
    Case e_AlteracaoDadosPessoais
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Altera��o de Dados Pessoais." & vbCrLf
    
    Case e_InclusaoRelacaoBicicletas
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Inclus�o na Rela��o de Bicicletas " & vbCrLf
    Case e_AlteracaoRelacaoBicicletas
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Altera��o na Rela��o de Bicicletas " & vbCrLf
    Case e_ExclusaoRelacaoBicicletas
        M46V999.M46V111.gCod_DizerTexto = M46V999.M46V111.gCod_DizerTexto & "Item: " & Format(pItem, "0000") & " - Exclus�o na Rela��o de Bicicletas " & vbCrLf
    End Select
End Sub
Public Function mfRetornaProxNumItem(ByRef pobjstcPedido As stcA46V702B) As Long

'    Dim lRegistro           As String
    Dim lNomeTab As String
    '    Dim lNumIteApolice      As Integer
    Dim lNumApolice As String
    '    Dim lNossoNumero        As String
    '    Dim lNumEndo            As Long
    '    Dim lNumPI              As Long

    Dim lSelect As String
    Dim lMensagem As String

    Dim bdP00APOL As ADODB.Connection
    Dim RstAux As clsYasRecordSet
    Dim lItemAux As Long
    Dim lItemAux2 As Long
    Dim lItemAux3 As Long

    If M46V999.gPrograma = "D046O003" Then
        lNomeTab = "SYASPORTAL.DBO."
    Else
        lNomeTab = "RAMOSDIVERSOS.DBO."
    End If


    'b) Abre arquivo.
    If pobjstcPedido.M46V999.gfAbrBasDados("P00APOL", pobjstcPedido.M46V999.gTipP00APOL, pobjstcPedido.M46V999.gEndP00APOL, bdP00APOL, lMensagem) = False Then
        Exit Function
    End If

    lNumApolice = pobjstcPedido.Num_Apol_An

    If lNumApolice = 0 Then
        Exit Function    'Verificar pois isso n�o pode ocorrer
    End If

    lSelect = "SELECT max(num_item) as proxnum "
    lSelect = lSelect & " FROM GED.DBO.TAB_PED_LOC WITH (NOLOCK)"
    lSelect = lSelect & " Where Num_Apol = " & lNumApolice

    If pobjstcPedido.M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then
        Call pobjstcPedido.M46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If

    If Not RstAux.EOF Then
        lItemAux = RstAux("proxnum") + 1
    End If


    lItemAux2 = 0
    lItemAux3 = 0
Retorno:
    'ver se j� foi emitido em outro c�lculo.
    'If CnDB Is Nothing Then GoTo ControleErros



    lSelect = "Select max(num_item) as max from " & lNomeTab & "TAB_PED_LOC a1 WITH (NOLOCK)"
    lSelect = lSelect & " inner join " & lNomeTab & "TAB_PED a2 WITH (NOLOCK) on a1.nosso_numero = a2.nosso_numero"
    lSelect = lSelect & " Where Num_Apol_An = " & M46V999.gfFormatarCamposSQL(pobjstcPedido.Num_Apol_An)
    If M46V999.gfPreenchido(pobjstcPedido.nosso_numero) Then
        lSelect = lSelect & " and a1.NOSSO_NUMERO  = " & M46V999.gfFormatarCamposSQL(pobjstcPedido.nosso_numero)
    End If
    lSelect = lSelect & " and Cod_DIzerItem = 1"


    If M46V999.gfObtRegistro(bdP00APOL, lSelect, RstAux, lMensagem) = False Then GoTo SAIDA
    If Not RstAux.EOF Then
        If IsNull(RstAux("max")) Then
            lItemAux2 = 1
        Else
            lItemAux2 = RstAux("max") + 1
        End If
    End If

    'Pego o maior.
    If lItemAux > lItemAux2 Then
        mfRetornaProxNumItem = lItemAux
    Else
        mfRetornaProxNumItem = lItemAux2
    End If

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Call M46V999.gpFecha2(bdP00APOL)
    Set bdP00APOL = Nothing
    Set RstAux = Nothing
    Return
End Function

Public Property Get Class_Name() As String
    Class_Name = "A46V101A"
End Property

Public Function gfCriArqCancelamentoEmissao(ByVal pCodCorretor As Long, pUsuario As String, _
                                            ByVal pPropostas As String, ByVal pNumPI As Long, _
                                            ByRef pMensagem As String) As Boolean
'Fun��o: cria arquivo de cancelamento de emiss�o


    Dim lCodUsuTransmissao As String          'C�digo do usu�rio de transmiss�o.
    Dim lDatHorEnvio As String    'Data e hora de envio.
    Dim lEndArqCancelamento As String           'Endere�o do arquivo de transmiss�o.
    Dim lEndArqSEQ As String    'Endere�o do arquivo    que armazena o numero de sequencia da transmissao.
    Dim lHora As String    'Hora de envio.
    Dim lNomArqCancelamento As String           'Nome do arquivo de transmiss�o.
    Dim lNumArqCancelamento As Integer          'Utilizado com FreeFile para arquivo de transmiss�o.
    Dim lSequencia As String  'Seq��ncia para o arquivo de transmiss�o.

    Err.Clear
    On Error GoTo TrataErro
    gfCriArqCancelamentoEmissao = False

    If M46V999.gTipUsuario <> "2" Then
        'S� gera arquivo de transmiss�o automatico para usu�rio tipo = 2
        gfCriArqCancelamentoEmissao = True
        GoTo SAIDA
    End If
    lCodUsuTransmissao = Format(pUsuario, "000000")

    '5. Gera arquivo para transmiss�o.
    'a) Verifica pasta remessa.
    If Right$(M46V999.gEndRemessa, 1) = "\" Then
        M46V999.gEndRemessa = Left$(M46V999.gEndRemessa, Len(M46V999.gEndRemessa) - 1)
    End If
    If Not M46V999.gfPreenchido(Dir(M46V999.gEndRemessa, vbDirectory)) Then
        On Error Resume Next
        MkDir (M46V999.gEndRemessa)
        If Err <> 0 Then
            Call M46V999.gpGraLog(2, "FTE0033 - Erro ao criar pasta " & M46V999.gEndRemessa & _
                                   " (pasta de remessa para arquivos de transmiss�o - " & Err & " - " & _
                                     ERROR & ")." & vbLf & vbLf & "O sistema ser� encerrado.")
            Exit Function
        End If
    End If
    M46V999.gEndRemessa = M46V999.gEndRemessa & "\"

    'b) Obt�m seq��ncia para arquivo de transmiss�o.
    'c) Cria arquivo de transmiss�o.
    lEndArqSEQ = M46V999.gAppPath & "SEQ012.LOG"
    Call mpObtSeq(lEndArqSEQ, lSequencia)
    lNomArqCancelamento = lCodUsuTransmissao & lSequencia & ".012"
    If M46V999.gfPreenchido(M46V999.gEndSYASUpDt) = True Then
        'Caminho dos arquivos de atualiza��o do SYAS
        'Informado no arquivo SYASPORD.INI
        lEndArqCancelamento = Trim(M46V999.gEndSYASUpDt) & lNomArqCancelamento
    Else
        'Endere�o das bases de dados tarifarias
        lEndArqCancelamento = M46V999.gAppPath & "Remessa\"
    End If
    lEndArqCancelamento = lEndArqCancelamento & lNomArqCancelamento
    If M46V999.gfPreenchido(Dir(lEndArqCancelamento)) Then
        On Error Resume Next
        Kill lEndArqCancelamento
        On Error GoTo TrataErro
    End If

    lNumArqCancelamento = FreeFile
    Open lEndArqCancelamento For Output As #lNumArqCancelamento
    If Err <> 0 Then
        pMensagem = "FTE0034 - Erro ao criar arquivo " & lEndArqCancelamento & _
                  " (arquivo de transmiss�o - " & Err & " - " & ERROR & ")."
        GoTo SAIDA
    End If
    'd) Grava primeiro registro.
    lHora = Format$(Now, "hh:nn:ss")
    lDatHorEnvio = Mid$(M46V999.gForHoje, 7, 4) & Mid$(M46V999.gForHoje, 4, 2) & Mid$(M46V999.gForHoje, 1, 2) & _
                   Mid$(lHora, 1, 2) & Mid$(lHora, 4, 2) & Mid$(lHora, 7, 2)
    Print #lNumArqCancelamento, "CORRETOR " & pCodCorretor & " USUARIO " & lCodUsuTransmissao & _
                              " DATA " & M46V999.gForHoje & " HORA " & lHora


    Print #lNumArqCancelamento, "CANCELAR" & pPropostas & pNumPI

    '7. Fecha arquivo - ser� utilizado no pr�ximo comando.
    Close #lNumArqCancelamento

    '8. Grava registro de controle no final do arquivo (�ltimo registro do arquivo que serve para
    '   verificar se arquivo enviado pelo corretor manteve sua integridade ao ser processado pelo
    '   servidor YASUDA).
    '   Formato do registro de controle: "PROPOSTA XXXXXXXXXXX", onde XXXXXXXXXXX � o valor da soma dos
    '   valores ASC de todos os caracteres das linhas anteriores, gravado de forma invertida.
    '   Por exemplo: Se soma dos valores ASC de todos os caracteres dos registros de um arquivo
    '   (excetuando-se �ltimo registro) for igual a 39.485, �ltimo registro ser�: PROPOSTA 58493.
    'a) Soma valores ASC do registro.

    If Not M46V999.gfPreenchido(Dir(lEndArqCancelamento)) Then
        pMensagem = "FTE0025 - Arquivo n�o localizado: " & lEndArqCancelamento & _
                  " (criar arquivo de transmiss�o)."
    Else
        Open lEndArqCancelamento For Input As #lNumArqCancelamento
        If Err <> 0 Then
            pMensagem = "FTE0026 - Erro ao abrir arquivo " & lEndArqCancelamento & ": " & Err & " - " & _
                        ERROR & ")."
            GoTo SAIDA
        End If
        Close #lNumArqCancelamento
        On Error GoTo 0
    End If

    gfCriArqCancelamentoEmissao = True

SAIDA:
    Err.Clear
    Exit Function

TrataErro:
    pMensagem = "gfCriArqCancelamentoEmissao Erro: " & Err & " - " & ERROR
    GoTo SAIDA
End Function
Public Function gfAtualizaClasse4Portal(ByVal pNossoNumero As String, _
                                        ByVal pStatus As Integer, pDesconto As Double, _
                                        ByRef pMensagem As String, pConexao As ADODB.Connection) As Boolean
    Dim lSql As String


    lSql = "Exec syasportal.dbo.proc_RecebeRetornoCotacao " & M46V999.gfFormatarCamposSQL(pNossoNumero) & "," & pStatus & "," & M46V999.gfFormatarCamposSQL(pDesconto)
    If M46V999.gfExeSQL(pConexao, lSql, pMensagem) <> 0 Then
        Exit Function
    End If

    gfAtualizaClasse4Portal = True

End Function

Public Function gfCriCotacaoAceita(ByVal pCodCorretor As Long, pCodUsuTransmissao As String, _
                                   pColecaoPedido As Collection, _
                                   pStatus As Integer, pCotacao As String, pMaxDesc As Double, pNomeArq As String, _
                                   pSenha As String, ByRef pMensagem As String) As Boolean

'Fun��o: cria arquivo de retorno da cota��o

    Dim lAppPath As String    'App.Path com \ no final.
    Dim lCodUsuTransmissao As String          'C�digo do usu�rio de transmiss�o.
    Dim lDatHorEnvio As String    'Data e hora de envio.
    Dim lEncontrou As Boolean    'Indica que encontrou c�digo do usu�rio de transmiss�o.
    Dim lEndArqTransmissao As String          'Endere�o do arquivo de transmiss�o.
    Dim lEndGeisedi2LOG As String       'Endere�o do arquivo de LOG EDIWare
'(GEISEDI2.LOG ou ERROR.LOG).
    Dim lNumSyasEdi_EXE As Integer       'Utilizado com FreeFile para SyasEdi.exe.
    Dim lEndEapscriptINI As String        'Endere�o do arquivo EAPSCRIPT.INI.
    Dim lHora As String    'Hora de envio.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lInvertido As String  'Total invertido.
    Dim lMensagem As String    'Mensagem.
    Dim lNumArqTransmissao As Integer         'Utilizado com FreeFile para arquivo de transmiss�o.
    Dim lNumEapscriptINI As Integer       'Utilizado com FreeFile para EAPSCRIPT.INI.
    Dim lRegistro As String    'Registro.
    Dim lSequencia As String  'Seq��ncia para o arquivo de transmiss�o.
    Dim lTotal As Double    'Total.
    Dim lobjstcPedido As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED)
    Dim lEndSEQ032LOG As String     'Endere�o do arquivo de LOG EDIWare
'(GEISEDI2.LOG ou ERROR.LOG).
    Dim lEndSyasEdi_EXE As String       'Endere�o do arquivo SyasEdi.EXE.

    On Error GoTo ControleErro
    gfCriCotacaoAceita = False
    'Call M46V999.gpGraLog(0, "AGE0001 - Gerando arquivo ")
    '1. Posiciona controle.
    If Right$(App.Path, 1) = "\" Then
        lAppPath = App.Path
    Else
        lAppPath = App.Path & "\"
    End If

    '4. Obt�m c�digo do usu�rio de transmiss�o.
    lCodUsuTransmissao = pCodUsuTransmissao


    '3. Gera arquivo para transmiss�o.
    'a) Verifica pasta remessa.
    If Right$(M46V999.gEndRemessa, 1) = "\" Then
        M46V999.gEndRemessa = Left$(M46V999.gEndRemessa, Len(M46V999.gEndRemessa) - 1)
    End If
    If Not M46V999.gfPreenchido(Dir(M46V999.gEndRemessa, vbDirectory)) Then
        On Error Resume Next
        MkDir (M46V999.gEndRemessa)
        If Err <> 0 Then
            Call M46V999.gpGraLog(2, "FTE0015 - Erro ao criar pasta " & M46V999.gEndRemessa & _
                                   " (pasta de remessa para arquivos de transmiss�o - " & Err & " - " & _
                                     ERROR & ")." & vbLf & vbLf & "O sistema ser� encerrado.")
            GoTo SAIDA
        End If
    End If
    M46V999.gEndRemessa = M46V999.gEndRemessa & "\"

    Set lobjstcPedido = pColecaoPedido(1)


    'c) Cria arquivo de retorno da cota��o

    '*** Remo��o da rotina abaixo �ra poder gravar mais de uma cota��o por arquivo - SIES - 11/10/2013
    '    lEndArqTransmissao = Trim(M46V999.gEndRemessa) & pNomeArq
    '    If M46V999.gfPreenchido(Dir(lEndArqTransmissao)) Then
    '        On Error Resume Next
    '        Kill lEndArqTransmissao
    '        On Error GoTo ControleErro
    '    End If

    lEndArqTransmissao = pNomeArq

    lNumArqTransmissao = FreeFile
    On Error Resume Next
    'Open lEndArqTransmissao For Output As #lNumArqTransmissao
    Open lEndArqTransmissao For Append Shared As #lNumArqTransmissao
    If Err <> 0 Then
        pMensagem = "FTE0016 - Erro ao criar arquivo " & lEndArqTransmissao & _
                  " (arquivo de transmiss�o - " & Err & " - " & ERROR & ")."
        On Error GoTo ControleErro
    End If

    'd) Grava primeiro registro.
    lHora = Format$(Now, "hh:nn:ss")
    lDatHorEnvio = Mid$(M46V999.gForHoje, 7, 4) & Mid$(M46V999.gForHoje, 4, 2) & Mid$(M46V999.gForHoje, 1, 2) & _
                   Mid$(lHora, 1, 2) & Mid$(lHora, 4, 2) & Mid$(lHora, 7, 2)
    Print #lNumArqTransmissao, "V0001 " & _
                               "CORRETOR " & pCodCorretor & " USUARIO " & pCodUsuTransmissao & _
                             " DATA " & M46V999.gForHoje & " HORA " & lHora

    '4. Processa os objetos da cole��o.
    For Each lobjstcPedido In pColecaoPedido
        'a) Obt�m nosso n�mero.

        If Not IsNumeric(lobjstcPedido.nosso_numero) Then
            pMensagem = "FTE0017 - Erro ao criar arquivo de transmiss�o (Nosso N�mero n�o num�rico)."
            Kill pNomeArq
            GoTo SAIDA
        End If
        If Len(CStr(lobjstcPedido.nosso_numero)) <> 20 Then
            pMensagem = "FTE0018 - Erro ao criar arquivo de transmiss�o ( Nosso n�mero com tamanho diferente de 20 posi��es)."
            Kill pNomeArq
            GoTo SAIDA
        End If

        'Primeira linha � com os dados de solicita��o da cota��o
        If mfTraCotacaoAceita(lNumArqTransmissao, lobjstcPedido.nosso_numero, pStatus, pCotacao, pSenha, pMaxDesc, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lEndArqTransmissao
            GoTo SAIDA
        End If

        If lobjstcPedido.Cod_Ramo = 710 Or lobjstcPedido.Cod_Ramo = 300 Or lobjstcPedido.Cod_Ramo = 620 Then     'Arquivo SIES n�o precisa enviar o detalhe
            GoTo FinalizaProcesso
        End If

        'Grava cabe�alho de nosso numero
        Print #lNumArqTransmissao, "100" & lobjstcPedido.nosso_numero
        'b) Grava registro de dados comuns.
        If mfTraPedido(lNumArqTransmissao, lobjstcPedido, pCodUsuTransmissao, _
                       pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'c) Grava registro de controle de emiss�o
        If mfTraCtrlEmis(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'd) Grava registros de corretores.
        If mfTraCorretores(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'e) Grava registros de cosseguros
        If mfTraCosseguro(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'f) Grava registros de endere�o.
        If mfTraEndereco(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'g) Grava registros de juros.
        If mfTraJuros(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'h) Grava registros de vencimento.
        If mfTraVenc(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'i) Grava registros de itens.
        If mfTraItens(pNomeArq, lDatHorEnvio, lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'j) Grava registros de Clausulas.
        If mfTraClausulas(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'k) Grava registros de Coberturas.
        If mfTraCoberturas(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'l) Grava registros de Seguros

        If mfTraSeguro(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'm) Grava registros de Bens.
        If mfTraBens(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'n) Grava registros de inspe��o

        If mfTraInspecao(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'o) Grava registros de observa��o
        If mfTraObservacao(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If


        'p) Grava registros de Dados Complementares das Coberturas.
        If mfTraCobertCompl(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'q) Grava registros de Texto para Endosso.
        If mfTraTextoEndosso(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'r) Grava registros de Dados Fatores de Risco.
        If mfTraFatoresProtecao(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        's) Grava registros de Benefici�rios.
        If mfTraBeneficiario(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        If mfTraLMI(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill pNomeArq
            GoTo SAIDA
        End If

        'Grava final de nosso n�mero de nosso numero
        Print #lNumArqTransmissao, "199" & lobjstcPedido.nosso_numero
    Next lobjstcPedido

FinalizaProcesso:

    '7. Fecha arquivo - ser� utilizado no pr�ximo comando.
    Close #lNumArqTransmissao
    'Call M46V999.gpGraLog(0, "AGE0001 - FIM")
    '8. Grava registro de controle no final do arquivo (�ltimo registro do arquivo que serve para
    '   verificar se arquivo enviado pelo corretor manteve sua integridade ao ser processado pelo
    '   servidor YASUDA).
    '   Formato do registro de controle: "PROPOSTA XXXXXXXXXXX", onde XXXXXXXXXXX � o valor da soma dos
    '   valores ASC de todos os caracteres das linhas anteriores, gravado de forma invertida.
    '   Por exemplo: Se soma dos valores ASC de todos os caracteres dos registros de um arquivo
    '   (excetuando-se �ltimo registro) for igual a 39.485, �ltimo registro ser�: PROPOSTA 58493.

    '*** Ajustes nas rotinas Open para cria��o de arquivo de retorno (.034 - Auto/Multi | .036 - SIES) - 14/10

    'a) Soma valores ASC do registro.
    lTotal = 0
    If Not M46V999.gfPreenchido(Dir(lEndArqTransmissao)) Then
        pMensagem = "FTE0019 - Arquivo n�o localizado: " & lEndArqTransmissao & _
                  " (criar arquivo de transmiss�o)."
    Else
        On Error Resume Next
        Open lEndArqTransmissao For Append Shared As #lNumArqTransmissao
        'Open lEndArqTransmissao For Input As #lNumArqTransmissao
        If Err <> 0 Then
            pMensagem = "FTE0020 - Erro ao abrir arquivo " & lEndArqTransmissao & ": " & Err & " - " & _
                        ERROR & ")."
            On Error GoTo ControleErro
        End If
        On Error GoTo ControleErro
    End If
    'Line Input #lNumArqTransmissao, lRegistro
    Do
        For lI = 1 To Len(Trim(lRegistro))
            lTotal = lTotal + Asc(Mid(lRegistro, lI, 1))
        Next lI
        If EOF(lNumArqTransmissao) Then
            Exit Do
        End If
        '        Line Input #lNumArqTransmissao, lRegistro
        Print #lNumArqTransmissao, ""
    Loop
    Close #lNumArqTransmissao

    'b) Inverte total e monta registro.
    lInvertido = ""
    For lI = Len(Trim(CStr(lTotal))) To 1 Step -1
        lInvertido = lInvertido & Mid(Trim(CStr(lTotal)), lI, 1)
    Next lI
    lRegistro = "COTACAO " & lInvertido
    'c) Grava registro.
    'Open lEndArqTransmissao For Append As #lNumArqTransmissao
    Open lEndArqTransmissao For Append Shared As #lNumArqTransmissao
    Print #lNumArqTransmissao, lRegistro
    Print #lNumArqTransmissao, ""
    Close #lNumArqTransmissao
    gfCriCotacaoAceita = True
SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    Set lobjstcPedido = Nothing
    If Not M46V999.gfPreenchido(pMensagem) Then
        pMensagem = Err.Description
    End If
    Call M46V999.gpGraLog(0, "rotina gfCriCotacaoAceita : " & pMensagem)
    GoTo SAIDA
FechaObj:
    Set lobjstcPedido = Nothing
    Return
End Function

Public Function gfCriArqCotacao(ByVal pCodCorretor As Long, pCodUsuTransmissao As String, _
                                pColecaoPedido As Collection, _
                                pEmail As String, pNome As String, pTelefone As String, _
                                pCCotacaoOrig As String, _
                                ByRef pEndArqTransmissao As String, _
                                ByRef pMensagem As String, ByVal pObservacao As String) As Boolean
'Fun��o: cria arquivo de transmiss�o.

'Par�metros de entrada...pCodCorretor.........C�digo do corretor.
'                        pPropostas...........Matriz com rela��o das propostas.
'Par�metros de sa�da.....pEndArqTransmissao...Endere�o do arquivo de transmiss�o.
'                        pMensagem............Mensagem.

    Dim lAppPath As String    'App.Path com \ no final.
    Dim lCodUsuTransmissao As String          'C�digo do usu�rio de transmiss�o.
    Dim lDatHorEnvio As String    'Data e hora de envio.
    Dim lEncontrou As Boolean    'Indica que encontrou c�digo do usu�rio de transmiss�o.
    Dim lEndArqTransmissao As String          'Endere�o do arquivo de transmiss�o.
    Dim lEndGeisedi2LOG As String       'Endere�o do arquivo de LOG EDIWare
'(GEISEDI2.LOG ou ERROR.LOG).
    Dim lNumSyasEdi_EXE As Integer       'Utilizado com FreeFile para SyasEdi.exe.
    Dim lEndEapscriptINI As String        'Endere�o do arquivo EAPSCRIPT.INI.
    Dim lHora As String    'Hora de envio.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lInvertido As String  'Total invertido.
    Dim lMensagem As String    'Mensagem.
    Dim lNomArqTransmissao As String          'Nome do arquivo de transmiss�o.
    Dim lNumArqTransmissao As Integer         'Utilizado com FreeFile para arquivo de transmiss�o.
    Dim lNumEapscriptINI As Integer       'Utilizado com FreeFile para EAPSCRIPT.INI.
    Dim lRegistro As String    'Registro.
    Dim lSequencia As String  'Seq��ncia para o arquivo de transmiss�o.
    Dim lTotal As Double    'Total.
    Dim lobjstcPedido As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED)
    Dim lEndSEQ032LOG As String     'Endere�o do arquivo de LOG EDIWare
'(GEISEDI2.LOG ou ERROR.LOG).
    Dim lEndSyasEdi_EXE As String       'Endere�o do arquivo SyasEdi.EXE.

    On Error GoTo ControleErro
    gfCriArqCotacao = False

    '1. Posiciona controle.
    If Right$(App.Path, 1) = "\" Then
        lAppPath = App.Path
    Else
        lAppPath = App.Path & "\"
    End If

    '4. Obt�m c�digo do usu�rio de transmiss�o.
    lCodUsuTransmissao = pCodUsuTransmissao


    '3. Gera arquivo para transmiss�o.
    'a) Verifica pasta remessa.
    If Right$(M46V999.gEndRemessa, 1) = "\" Then
        M46V999.gEndRemessa = Left$(M46V999.gEndRemessa, Len(M46V999.gEndRemessa) - 1)
    End If
    If Not M46V999.gfPreenchido(Dir(M46V999.gEndRemessa, vbDirectory)) Then
        On Error Resume Next
        MkDir (M46V999.gEndRemessa)
        If Err <> 0 Then
            Call M46V999.gpGraLog(2, "FTE0015 - Erro ao criar pasta " & M46V999.gEndRemessa & _
                                   " (pasta de remessa para arquivos de transmiss�o - " & Err & " - " & _
                                     ERROR & ")." & vbLf & vbLf & "O sistema ser� encerrado.")
            GoTo SAIDA
        End If
    End If
    M46V999.gEndRemessa = M46V999.gEndRemessa & "\"


    lEndSEQ032LOG = lAppPath & "SEQ032.LOG"

    Call mpObtSeq(lEndSEQ032LOG, lSequencia, 9999)

    Set lobjstcPedido = pColecaoPedido(1)
Retorno:
    'c) Cria arquivo de transmiss�o.
    lNomArqTransmissao = lCodUsuTransmissao & lSequencia & ".032"

    lEndArqTransmissao = Trim(M46V999.gEndRemessa) & lNomArqTransmissao
    If M46V999.gfPreenchido(Dir(lEndArqTransmissao)) Then
        If lSequencia = 9999 Then lSequencia = 0
        lSequencia = Format(CInt(lSequencia) + 1, "0000")
        GoTo Retorno
    End If
    lNumArqTransmissao = FreeFile
    On Error Resume Next
    Open lEndArqTransmissao For Output As #lNumArqTransmissao
    If Err <> 0 Then
        pMensagem = "FTE0016 - Erro ao criar arquivo " & lEndArqTransmissao & _
                  " (arquivo de transmiss�o - " & Err & " - " & ERROR & ")."
        On Error GoTo ControleErro
    End If

    'd) Grava primeiro registro.
    lHora = Format$(Now, "hh:nn:ss")
    lDatHorEnvio = Mid$(M46V999.gForHoje, 7, 4) & Mid$(M46V999.gForHoje, 4, 2) & Mid$(M46V999.gForHoje, 1, 2) & _
                   Mid$(lHora, 1, 2) & Mid$(lHora, 4, 2) & Mid$(lHora, 7, 2)
    Print #lNumArqTransmissao, "V0001 " & _
                               "CORRETOR " & pCodCorretor & " USUARIO " & pCodUsuTransmissao & _
                             " DATA " & M46V999.gForHoje & " HORA " & lHora

    '4. Processa os objetos da cole��o.
    For Each lobjstcPedido In pColecaoPedido
        'a) Obt�m nosso n�mero.

        If Not IsNumeric(lobjstcPedido.nosso_numero) Then
            pMensagem = "FTE0017 - Erro ao criar arquivo de transmiss�o (Nosso N�mero n�o num�rico)."
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If
        If Len(CStr(lobjstcPedido.nosso_numero)) <> 20 Then
            pMensagem = "FTE0018 - Erro ao criar arquivo de transmiss�o ( Nosso n�mero com tamanho diferente de 20 posi��es)."
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'Primeira linha � com os dados de solicita��o da cota��o
        If mfTraCotacao(lNumArqTransmissao, lobjstcPedido.nosso_numero, pNome, _
                        pTelefone, pEmail, pCCotacaoOrig, pObservacao, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lEndArqTransmissao
            GoTo SAIDA
        End If

        'Grava cabe�alho de nosso numero
        Print #lNumArqTransmissao, "100" & lobjstcPedido.nosso_numero

        'b) Grava registro de dados comuns.
        If mfTraPedido(lNumArqTransmissao, lobjstcPedido, pCodUsuTransmissao, _
                       pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'c) Grava registro de controle de emiss�o
        If mfTraCtrlEmis(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'd) Grava registros de corretores.
        If mfTraCorretores(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'e) Grava registros de cosseguros
        If mfTraCosseguro(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'f) Grava registros de endere�o.
        If mfTraEndereco(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'g) Grava registros de juros.
        If mfTraJuros(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'h) Grava registros de vencimento.
        If mfTraVenc(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'i) Grava registros de itens.
        If mfTraItens(lNomArqTransmissao, lDatHorEnvio, lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'j) Grava registros de Clausulas.
        If mfTraClausulas(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'k) Grava registros de Coberturas.
        If mfTraCoberturas(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'l) Grava registros de Seguros
        If mfTraSeguro(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'm) Grava registros de Bens.
        If mfTraBens(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'n) Grava registros de inspe��o
        If mfTraInspecao(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'o) Grava registros de observa��o
        If mfTraObservacao(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If


        'p) Grava registros de Dados Complementares das Coberturas.
        If mfTraCobertCompl(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'q) Grava registros de Texto para Endosso.
        If mfTraTextoEndosso(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'r) Grava registros de Dados Fatores de Risco.
        If mfTraFatoresProtecao(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        's) Grava registros de Benefici�rios.
        If mfTraBeneficiario(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'Grava final de nosso n�mero de nosso numero
        Print #lNumArqTransmissao, "199" & lobjstcPedido.nosso_numero
    Next lobjstcPedido

    '7. Fecha arquivo - ser� utilizado no pr�ximo comando.
    Close #lNumArqTransmissao

    '8. Grava registro de controle no final do arquivo (�ltimo registro do arquivo que serve para
    '   verificar se arquivo enviado pelo corretor manteve sua integridade ao ser processado pelo
    '   servidor YASUDA).
    '   Formato do registro de controle: "PROPOSTA XXXXXXXXXXX", onde XXXXXXXXXXX � o valor da soma dos
    '   valores ASC de todos os caracteres das linhas anteriores, gravado de forma invertida.
    '   Por exemplo: Se soma dos valores ASC de todos os caracteres dos registros de um arquivo
    '   (excetuando-se �ltimo registro) for igual a 39.485, �ltimo registro ser�: PROPOSTA 58493.
    'a) Soma valores ASC do registro.
    lTotal = 0
    If Not M46V999.gfPreenchido(Dir(lEndArqTransmissao)) Then
        pMensagem = "FTE0019 - Arquivo n�o localizado: " & lEndArqTransmissao & _
                  " (criar arquivo de transmiss�o)."
    Else
        On Error Resume Next
        Open lEndArqTransmissao For Input As #lNumArqTransmissao
        If Err <> 0 Then
            pMensagem = "FTE0020 - Erro ao abrir arquivo " & lEndArqTransmissao & ": " & Err & " - " & _
                        ERROR & ")."
            On Error GoTo ControleErro
        End If
        On Error GoTo ControleErro
    End If
    Line Input #lNumArqTransmissao, lRegistro
    Do
        For lI = 1 To Len(Trim(lRegistro))
            lTotal = lTotal + Asc(Mid(lRegistro, lI, 1))
        Next lI
        If EOF(lNumArqTransmissao) Then
            Exit Do
        End If
        Line Input #lNumArqTransmissao, lRegistro
    Loop
    Close #lNumArqTransmissao
    'b) Inverte total e monta registro.
    lInvertido = ""
    For lI = Len(Trim(CStr(lTotal))) To 1 Step -1
        lInvertido = lInvertido & Mid(Trim(CStr(lTotal)), lI, 1)
    Next lI
    lRegistro = "COTACAO " & lInvertido
    'c) Grava registro.
    Open lEndArqTransmissao For Append As #lNumArqTransmissao
    Print #lNumArqTransmissao, lRegistro
    Close #lNumArqTransmissao

    '9. Posiciona endere�o do arquivo de transmiss�o.
    pEndArqTransmissao = lEndArqTransmissao
    gfCriArqCotacao = True
SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    Set lobjstcPedido = Nothing
    If Not M46V999.gfPreenchido(pMensagem) Then
        pMensagem = Err.Description
    End If
    Call M46V999.gpGraLog(0, "rotina gfCriArqCotacao : " & pMensagem)
    GoTo SAIDA
FechaObj:
    Set lobjstcPedido = Nothing
    Return
End Function

Public Function gfCriArqTransmissao(ByVal pCodCorretor As Long, _
                                    pCodUsuTransmissao As String, _
                                    pColecaoPedido As Collection, _
                                    ByRef pEndArqTransmissao As String, _
                                    ByRef pMensagem As String, _
                                    Optional pAutomatico As Boolean = False, _
                                    Optional pCodUnicidade As Double = 0, _
                                    Optional pLogCalculo As Boolean = False) As Boolean
'Fun��o: cria arquivo de transmiss�o.

'Par�metros de entrada...pCodCorretor.........C�digo do corretor.
'                        pPropostas...........Matriz com rela��o das propostas.
'Par�metros de sa�da.....pEndArqTransmissao...Endere�o do arquivo de transmiss�o.
'                        pMensagem............Mensagem.

    Dim lAppPath As String    'App.Path com \ no final.
'Dim lCodUsuTransmissao  As String           'C�digo do usu�rio de transmiss�o.
    Dim lDatHorEnvio As String    'Data e hora de envio.
    Dim lEncontrou As Boolean    'Indica que encontrou c�digo do usu�rio de transmiss�o.
    Dim lEndArqTransmissao As String          'Endere�o do arquivo de transmiss�o.
    Dim lEndGeisedi2LOG As String       'Endere�o do arquivo de LOG EDIWare
'(GEISEDI2.LOG ou ERROR.LOG).
    Dim lEndEapscriptINI As String        'Endere�o do arquivo EAPSCRIPT.INI.
    Dim lHora As String    'Hora de envio.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lInvertido As String  'Total invertido.
    Dim lMensagem As String    'Mensagem.
    Dim lNomArqTransmissao As String          'Nome do arquivo de transmiss�o.
    Dim lNumArqTransmissao As Integer         'Utilizado com FreeFile para arquivo de transmiss�o.
    Dim lNumEapscriptINI As Integer       'Utilizado com FreeFile para EAPSCRIPT.INI.
    Dim lRegistro As String    'Registro.
    Dim lSequencia As String  'Seq��ncia para o arquivo de transmiss�o.
    Dim lTotal As Double    'Total.
    Dim lobjstcPedido As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED)
    Dim lExtensao As String
    Dim lEndArqSEQ As String

    gfCriArqTransmissao = False

    '1. Posiciona controle.
    If Right$(App.Path, 1) = "\" Then
        lAppPath = App.Path
    Else
        lAppPath = App.Path & "\"
    End If

    '2. Obt�m c�digo do usu�rio de transmiss�o.

    'a) Abre arquivo EAPSCRIPT.INI.
    'lEndEapscriptINI = lAppPath & "EAPSCRIPT.INI"
    'If Not M46V999.gfPreenchido(Dir(lEndEapscriptINI)) Then
    '    pMensagem = "FTE0013 - Arquivo n�o localizado: " & lEndEapscriptINI & _
         '                " (cria arquivo de transmiss�o - c�digo do usu�rio de transmiss�o)."
    '    Exit Function
    'End If
    'lNumEapscriptINI = FreeFile
    'On Error Resume Next
    'Open lEndEapscriptINI For Input Access Read As #lNumEapscriptINI
    'If Err <> 0 Then
    '    pMensagem = "FTE0014 - Erro ao criar arquivo de transmiss�o (c�digo do usu�rio de " & _
         '                "transmiss�o n�o localizado - " & Err & " - " & Error & ")."
    '    Exit Function
    'End If
    '
    ' On Error GoTo 0
    'b) Obt�m c�digo do usu�rio.
    '***************************************
    ' No Syas Multi foi criado o par�metro
    '***************************************
    'lEncontrou = False
    'Do While Not EOF(lNumEapscriptINI)
    '    Line Input #lNumEapscriptINI, lRegistro
    '   If UCase$(Mid(lRegistro, 1, 7)) = "SENDER=" Then
    '       lCodUsuTransmissao = Trim$(Mid$(lRegistro, 9, 6))
    '        lEncontrou = True
    '        Exit Do
    '    End If
    'Loop
    'Close #lNumEapscriptINI
    'c) N�o encontrou c�digo do usu�rio de transmiss�o.
    'If lEncontrou = False Then
    '    pMensagem = "FTE0014 - Erro ao criar arquivo de transmiss�o (c�digo do usu�rio de " & _
         '                "transmiss�o n�o localizado)."
    '    Exit Function
    'End If

    '3. Gera arquivo para transmiss�o.
    'a) Verifica pasta remessa.
    If Right$(M46V999.gEndRemessa, 1) = "\" Then
        M46V999.gEndRemessa = Left$(M46V999.gEndRemessa, Len(M46V999.gEndRemessa) - 1)
    End If
    If Not M46V999.gfPreenchido(Dir(M46V999.gEndRemessa, vbDirectory)) Then
        On Error Resume Next
        MkDir (M46V999.gEndRemessa)
        If Err <> 0 Then
            Call M46V999.gpGraLog(2, "FTE0015 - Erro ao criar pasta " & M46V999.gEndRemessa & _
                                   " (pasta de remessa para arquivos de transmiss�o - " & Err & " - " & _
                                     ERROR & ")." & vbLf & vbLf & "O sistema ser� encerrado.")
            GoTo SAIDA
        End If
    End If
    M46V999.gEndRemessa = M46V999.gEndRemessa & "\"

    Set lobjstcPedido = pColecaoPedido(1)

    If pAutomatico = False And pLogCalculo = False Then
        If M46V999.gPrograma <> "D046O003" Then
            lExtensao = ".M08"
            lSequencia = "_" & lobjstcPedido.nosso_numero
        Else
            lExtensao = ".015"
            lEndArqSEQ = M46V999.gAppPath & "SEQ015.LOG"
            Call mpObtSeq(lEndArqSEQ, lSequencia, 9999)
        End If

        'C�digo do Usu�rio + Sequencial -> Chave
        lNomArqTransmissao = pCodUsuTransmissao & lSequencia & lExtensao    'Transmissao

        lEndArqTransmissao = Trim(M46V999.gEndRemessa) & lNomArqTransmissao
    Else
        If pLogCalculo = True Then
            lNomArqTransmissao = lobjstcPedido.nosso_numero & ".046"    '-Log de Calculo
        Else
            lNomArqTransmissao = lobjstcPedido.nosso_numero & pCodUnicidade & ".091"    'Transmissao Oculta
        End If
        '- Posiciona o caminho dos dados
        If M46V999.gfPreenchido(M46V999.gEndSYASUpDt) = True Then
            'Caminho dos arquivos de atualiza��o do SYAS
            'Informado no arquivo SYASPORD.INI
            lEndArqTransmissao = Trim(M46V999.gEndSYASUpDt) & lNomArqTransmissao
        Else
            'Endere�o das bases de dados tarifarias
            lEndArqTransmissao = Trim(M46V999.gEndServidor) & lNomArqTransmissao
        End If
    End If

    If M46V999.gfPreenchido(Dir(lEndArqTransmissao)) Then
        On Error Resume Next
        Kill lEndArqTransmissao
        On Error GoTo 0
    End If
    lNumArqTransmissao = FreeFile
    On Error Resume Next
    Open lEndArqTransmissao For Output As #lNumArqTransmissao
    If Err <> 0 Then
        pMensagem = "FTE0016 - Erro ao criar arquivo " & lEndArqTransmissao & _
                  " (arquivo de transmiss�o - " & Err & " - " & ERROR & ")."
        On Error GoTo 0
        GoTo SAIDA
    End If

    'd) Grava primeiro registro.
    lHora = Format$(Now, "hh:nn:ss")
    lDatHorEnvio = Mid$(M46V999.gForHoje, 7, 4) & Mid$(M46V999.gForHoje, 4, 2) & Mid$(M46V999.gForHoje, 1, 2) & _
                   Mid$(lHora, 1, 2) & Mid$(lHora, 4, 2) & Mid$(lHora, 7, 2)
    Print #lNumArqTransmissao, "V0001 " & _
                               "CORRETOR " & pCodCorretor & " USUARIO " & pCodUsuTransmissao & _
                             " DATA " & M46V999.gForHoje & " HORA " & lHora

    '4. Processa os objetos da cole��o.
    For Each lobjstcPedido In pColecaoPedido
        'a) Obt�m nosso n�mero.

        If Not IsNumeric(lobjstcPedido.nosso_numero) Then
            pMensagem = "FTE0017 - Erro ao criar arquivo de transmiss�o (Nosso N�mero n�o num�rico)."
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If
        If Len(CStr(lobjstcPedido.nosso_numero)) <> 20 Then
            pMensagem = "FTE0018 - Erro ao criar arquivo de transmiss�o ( Nosso n�mero com tamanho diferente de 20 posi��es)."
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If
        'Grava cabe�alho de nosso numero
        Print #lNumArqTransmissao, "100" & lobjstcPedido.nosso_numero

        'b) Grava registro de dados comuns.
        If mfTraPedido(lNumArqTransmissao, lobjstcPedido, pCodUsuTransmissao, _
                       pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'c) Grava registro de controle de emiss�o
        If mfTraCtrlEmis(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'd) Grava registros de corretores.
        If mfTraCorretores(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'e) Grava registros de cosseguros
        If mfTraCosseguro(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'f) Grava registros de endere�o.
        If mfTraEndereco(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'g) Grava registros de juros.
        If mfTraJuros(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'h) Grava registros de vencimento.
        If mfTraVenc(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'i) Grava registros de itens.
        If mfTraItens(lNomArqTransmissao, lDatHorEnvio, lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'j) Grava registros de Clausulas.
        If mfTraClausulas(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'k) Grava registros de Coberturas.
        If mfTraCoberturas(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'l) Grava registros de Seguros
        If mfTraSeguro(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'm) Grava registros de Bens.
        If mfTraBens(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'n) Grava registros de inspe��o
        If mfTraInspecao(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'o) Grava registros de observa��o
        If mfTraObservacao(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If


        'p) Grava registros de Dados Complementares das Coberturas.
        If mfTraCobertCompl(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'q) Grava registros de Texto para Endosso.
        If mfTraTextoEndosso(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'r) Grava registros de Dados Fatores de Risco.
        If mfTraFatoresProtecao(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        's) Grava registros de Benefici�rios.
        If mfTraBeneficiario(lNumArqTransmissao, lobjstcPedido, pMensagem) = False Then
            Close #lNumArqTransmissao
            Kill lNomArqTransmissao
            GoTo SAIDA
        End If

        'Grava final de nosso n�mero de nosso numero
        Print #lNumArqTransmissao, "199" & lobjstcPedido.nosso_numero
    Next lobjstcPedido

    '7. Fecha arquivo - ser� utilizado no pr�ximo comando.
    Close #lNumArqTransmissao

    '8. Grava registro de controle no final do arquivo (�ltimo registro do arquivo que serve para
    '   verificar se arquivo enviado pelo corretor manteve sua integridade ao ser processado pelo
    '   servidor YASUDA).
    '   Formato do registro de controle: "PROPOSTA XXXXXXXXXXX", onde XXXXXXXXXXX � o valor da soma dos
    '   valores ASC de todos os caracteres das linhas anteriores, gravado de forma invertida.
    '   Por exemplo: Se soma dos valores ASC de todos os caracteres dos registros de um arquivo
    '   (excetuando-se �ltimo registro) for igual a 39.485, �ltimo registro ser�: PROPOSTA 58493.
    'a) Soma valores ASC do registro.
    lTotal = 0
    If Not M46V999.gfPreenchido(Dir(lEndArqTransmissao)) Then
        pMensagem = "FTE0019 - Arquivo n�o localizado: " & lEndArqTransmissao & _
                  " (criar arquivo de transmiss�o)."
    Else
        On Error Resume Next
        Open lEndArqTransmissao For Input As #lNumArqTransmissao
        If Err <> 0 Then
            pMensagem = "FTE0020 - Erro ao abrir arquivo " & lEndArqTransmissao & ": " & Err & " - " & _
                        ERROR & ")."
            On Error GoTo 0
            GoTo SAIDA
        End If
        On Error GoTo 0
    End If
    Line Input #lNumArqTransmissao, lRegistro
    Do
        For lI = 1 To Len(Trim(lRegistro))
            lTotal = lTotal + Asc(Mid(lRegistro, lI, 1))
        Next lI
        If EOF(lNumArqTransmissao) Then
            Exit Do
        End If
        Line Input #lNumArqTransmissao, lRegistro
    Loop
    Close #lNumArqTransmissao
    'b) Inverte total e monta registro.
    lInvertido = ""
    For lI = Len(Trim(CStr(lTotal))) To 1 Step -1
        lInvertido = lInvertido & Mid(Trim(CStr(lTotal)), lI, 1)
    Next lI
    lRegistro = "PROPOSTA " & lInvertido
    'c) Grava registro.
    Open lEndArqTransmissao For Append As #lNumArqTransmissao
    Print #lNumArqTransmissao, lRegistro
    Close #lNumArqTransmissao

    '9. Posiciona endere�o do arquivo de transmiss�o.
    pEndArqTransmissao = lEndArqTransmissao
    If lExtensao = ".015" Then
        Call mpAtualizaSeqLog(Replace(lExtensao, ".", ""), lSequencia)
    End If

    gfCriArqTransmissao = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcPedido = Nothing
    Return
End Function
Private Function mfPosCampo(ByVal pOpcao As Byte, ByVal pNomCampo As String, ByVal pCampo, _
                            ByVal pTamanho As Integer, ByVal pDecimais As Byte, ByRef pMensagem As String) _
                            As Boolean
'Fun��o: posiciona campo.

'Par�metros de entrada...pOpcao......1 = posiciona campo num�rico com tamanho e casas decimais
'                                        especificados (consiste e emite mensagem).
'                                    2 = posiciona campo num�rico com tamanho e casas decimais
'                                        especificados (preenche com zero se campo com erro).
'                                    3 = indicador (S ou N).
'                                    4 = campo alfanum�rico (preenche com branco se campo com erro).
'                                    5 = campo num�rico - preenche com zero se registro incompleto ou
'                                        campo n�o num�rico - considera 2 casas decimais.
'                                    6 = campo alfanum�rico com conte�do num�rico - consiste e
'                                        emite mensagem.
'                        pNomCampo...nome do campo.
'                        pCampo......campo a ser formatado e posicionado no registro.
'                        pTamanho....tamanho do campo no registro de sa�da.
'                        pDecimais...quantidade de casas decimais a serem consideradas.
'Par�metros de sa�da.....pMensagem...mensagem.

    Dim lAuxiliar As String    'Auxiliar.

    mfPosCampo = False

    '1. Consiste informa��o.
    If pTamanho = 0 Then
        pMensagem = "FTE0021 - Erro ao criar arquivo de transmiss�o (tamanho do campo " & pNomCampo & _
                  " inv�lido - zerado)."
        Exit Function
    End If

    '2. Formata campo a ser colocado no registro.
    Select Case pOpcao
    Case 1      'Posiciona campo num�rico (consiste e emite mensagem).
        If Not IsNumeric(pCampo) Then
            pMensagem = "FTE0022 - Erro ao criar arquivo de transmiss�o (campo " & pNomCampo & _
                      " n�o num�rico)."
            Exit Function
        End If
        Select Case pDecimais
        Case 0
            lAuxiliar = Format$(pCampo, String(pTamanho, "0"))
        Case 1
            lAuxiliar = Format$(pCampo * 10, String(pTamanho, "0"))
        Case 2
            lAuxiliar = Format$(pCampo * 100, String(pTamanho, "0"))
        Case 3
            lAuxiliar = Format$(pCampo * 1000, String(pTamanho, "0"))
        Case 4
            lAuxiliar = Format$(pCampo * 10000, String(pTamanho, "0"))
        Case 5
            lAuxiliar = Format$(pCampo * 100000, String(pTamanho, "0"))
        Case 6
            lAuxiliar = Format$(pCampo * 1000000, String(pTamanho, "0"))
        Case Else
            pMensagem = "FTE0023 - Erro ao criar arquivo de transmiss�o (quantidade de " & _
                        "casas decimais inv�lida ao posicionar campo " & pNomCampo & " - " & _
                        pDecimais & ")."
            Exit Function
        End Select
        lAuxiliar = Right(lAuxiliar, pTamanho)
    Case 2      'Posiciona campo num�rico (preenche com zero se campo com erro).
        If IsNull(pCampo) Then
            lAuxiliar = String(pTamanho, "0")
        Else
            If IsNumeric(pCampo) Then
                Select Case pDecimais
                Case 0
                    lAuxiliar = Format$(pCampo, String(pTamanho, "0"))
                Case 1
                    lAuxiliar = Format$(pCampo * 10, String(pTamanho, "0"))
                Case 2
                    lAuxiliar = Format$(pCampo * 100, String(pTamanho, "0"))
                Case 3
                    lAuxiliar = Format$(pCampo * 1000, String(pTamanho, "0"))
                Case 4
                    lAuxiliar = Format$(pCampo * 10000, String(pTamanho, "0"))
                Case 5
                    lAuxiliar = Format$(pCampo * 100000, String(pTamanho, "0"))
                Case 6
                    lAuxiliar = Format$(pCampo * 1000000, String(pTamanho, "0"))
                Case Else
                    pMensagem = "FTE0023 - Erro ao criar arquivo de transmiss�o (quantidade " & _
                                "de casas decimais inv�lida ao posicionar campo " & pNomCampo & _
                              " - " & pDecimais & ")."
                    Exit Function
                End Select
            Else
                lAuxiliar = String(pTamanho, "0")
            End If
        End If
        lAuxiliar = Right(lAuxiliar, pTamanho)
    Case 4      'Posiciona campo alfanum�rico (preenche com branco se campo com erro).
        If M46V999.gfPreenchido(pCampo) Then
            Select Case Len(pCampo)
            Case Is < pTamanho
                lAuxiliar = pCampo
                While Len(lAuxiliar) < pTamanho
                    lAuxiliar = lAuxiliar & " "
                Wend
            Case pTamanho
                lAuxiliar = pCampo
            Case Is > pTamanho
                lAuxiliar = Left$(pCampo, pTamanho)
            End Select
        Else
            lAuxiliar = Space(pTamanho)
        End If
    Case 6      'Posiciona campo alfanum�rico com conte�do num�rico (consiste e emite mensagem).
        If Not M46V999.gfPreenchido(pCampo) Then
            pMensagem = "FTE0024 - Erro ao criar arquivo de transmiss�o (campo " & pNomCampo & _
                      " n�o preenchido)."
            Exit Function
        End If
        If Not IsNumeric(pCampo) Then
            pMensagem = "FTE0025 - Erro ao criar arquivo de transmiss�o (campo " & pNomCampo & _
                      " n�o num�rico)."
            Exit Function
        End If
        lAuxiliar = Right$(String(pTamanho, "0") & Trim(pCampo), pTamanho)
    Case Else
        pMensagem = "FTE0026 - Op��o inv�lida (posiciona campo no registro de transmiss�o - " & _
                    "op��o desconhecida - " & pOpcao & ")."
        Exit Function
    End Select

    '3. Acrescenta campo ao registro.
    mRegistro = mRegistro & lAuxiliar

    mfPosCampo = True
End Function
Private Function mfTraPedido(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                             ByVal pCodUsuTransmissao As String, ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de dados comuns para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        ........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.


    mfTraPedido = False

    '1. Grava registro de dados comuns.
    'mRegistro = "TAB_PED             "      'Indicador."
    mRegistro = "001"                        'Indicador."

    With pobjstcPedido
        If mfPosCampo(4, "Nosso n�mero", .nosso_numero, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do org�o", .Cod_Org, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do ramo", .Cod_Ramo, 3, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "N�mero do pedido", .Num_Ped, 6, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "N�mero da ap�lice", .Num_Apol, 10, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "N�mero do endosso", .Num_Endo, 6, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "N�mero da ap�lice antiga", .Num_Apol_An, 10, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do registro", .Cod_Reg, 3, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do banco de cobran�a", .Cod_Bco_Cobr, 3, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo da ag�ncia angar", .Cod_Agen_Angar, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo da ag�ncia Indicr", .Cod_Agen_Indicr, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "N�mero da conta", .Num_Cta_Indicr, 7, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo da unidade de produto", .Cod_Unid_Prod, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do produtor", .Cod_Prod, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo da ag�ncia de d�bito", .Cod_Agen_Bas_Cor, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "N�mero da conta de d�bito", Left(Format(.Num_Cta_Bas_Cor, "00000000"), 7), 7, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do Segurado", .Cod_Segurado, 10, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Nome do segurado", .Nom_Segurado, 60, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Complemento do nome do segurado", .Dsc_Compl_Segurado, 60, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "N�mero do CGC-CPF do segurado", .Num_Cgc_Cpf, 14, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Tipo de logradouro", .Cod_Tip_Logr, 3, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Logradouro", .Nom_Logr, 50, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Complemento do logradouro", .Dsc_Compl_Logr, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Bairro", .Nom_Bairro, 30, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Cidade", .Nom_Cid, 30, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Sigla UF", .Sig_UF, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "CEP", .Num_CEP, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Telefone", .Num_Fone, 7, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo Estip", .Cod_Estip, 10, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Coeficiente de comiss�o Pad", .Val_Coef_Cms_Pad, 5, 4, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data de in�cio da vig�ncia", .Dat_Ini_Vig, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data de t�rmino da vig�ncia", .Dat_Ter_Vig, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data de emiss�o", .Dat_Emis, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo da moeda", .Cod_Moed, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Percentual de juros do fracionamento", .Per_Juro_Frac, 4, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Quantidade de parcelas", .Qtd_Parc, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Quantidade de anexos", .Qtd_Anexo, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo de identifica��o do tipo de prazo", .Cod_Iden_Praz, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do pr�mio m�nimo de endosso", .Cod_Pr_Min_Endo, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do pr�mio m�nimo de parcela", .Cod_Pr_Min_Parc, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Valor pr�mio antecipado", .Val_Pr_Antec, 15, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Percentual adicioanl na forma de pagamento", .Per_Adic_Form_Pag, 4, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Percentual de desconto na forma de pagamento", .Per_Desc_Form_Pag, 4, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Valor da IS", .Val_IS, 15, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Valor da IS 2", .Val_IS_2, 15, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Valor do pr�mio", .VAL_PR, 15, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Valor do pr�mio ref tarif", .Val_Pr_Ref_Tarif, 15, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Valro do pr�mio ref  resg", .Val_Pr_Ref_Resg, 15, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data de in�cio Viag", .Dat_Ini_Viag, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do tipo de libera��o", .Cod_Tip_Liber, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "C�digo identificador de altera��o de corretor", .Cod_Iden_Alt_Corr, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "C�digo do usu�rio de digita��o", .Cod_User_Digit, 6, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data da digita��o", .DAT_DIGIT, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Hora da digita��o", .HOR_DIGIT, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "C�digo do usu�rio de libera��o", .Cod_User_Liber, 6, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data da libera��o", .Dat_Liber, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Hora da libera��o", .Hor_Liber, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo da situa��o do pedido", .Cod_Sit_Ped, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do tipo de documento", .Cod_Tip_Doc, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do dizer", .Cod_Dizer, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Quantidade de c�pias", .Qtd_Copia, 3, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Valor do custo de emiss�o", .Val_Cust_Emis, 15, 2, pMensagem) = False Then
            Exit Function
        End If
        'If mfPosCampo(2, "Valor da parcela m�nima", .Val_Parc_Min, 15, 2, pMensagem) = False Then
        '    Exit Function
        'End If
        If mfPosCampo(2, "Valor de juros", .Val_Juro, 15, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Valor IOF", .Val_IOF, 15, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "N�mero da proposta", .Num_Prop, 6, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "C�digo da posi��o atual", .Cod_Posi_Atual, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data do in�cio do movimento", .Dat_Mov_Ini, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data do t�rmino do movimento", .Dat_Mov_Ter, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Tipo de custo de emiss�o", .Tip_Cust_Emis, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Valor do cambio na emiss�o", .Val_Cambio_Emis, 11, 5, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Iden", .Iden, 5, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "N�mero do PI", .NUM_PI, 10, 0, pMensagem) = False Then
            Exit Function
        End If
        'Corretor tem somente produto par
        If .Cod_Produto Mod 2 = 1 Then
            .Cod_Produto = .Cod_Produto - 1
        End If
        If mfPosCampo(2, "C�digo do produto", .Cod_Produto, 3, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do protocolo", .Cod_Protocolo, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo da sequencia", .Cod_Sequencia, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Tipo de pessoa", .Tip_Pessoa, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do ramo de atividade", .Cod_Ramo_Ativ, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do tipo de documento", .Cod_Doc, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "N�mero do documento", .Num_Doc, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "�rg�o expedidor", .Org_Exp, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data da emiss�o do documento (yyyymmdd)", .Dat_Emis_Doc, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Nome da empresa onde trabalha", .Nom_Empresa, 30, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Sexo do segurado: M -masculino , F - Feminino.", .Sexo, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data do Primeira C�lculo.", .Dat_Pri_Calculo, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "�ltima Tarifa.", .Ult_Tarifa, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Codigo de Unicidade.", .Cod_Unicidade, 6, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Codigo do Digito da Agencia BasCor.", .Cod_Dig_Agen_Bas_Cor, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Codigo da Vers�o da Unicidade.", .Cod_Ver_Unici, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Renda/Faturamento.", .Val_RendaFatu, 18, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Percentual de LMI", .Per_Desc_LMI, 18, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Se eh LMI", .LMI_Unico, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "N�mero da Cotacao de LMI", .Num_Cotac_LMI, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        '1001046 - Telefone do Segurado 8 Digitos
        If mfPosCampo(2, "Telefone", .Num_Fone, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        'Tarifa 10.00 - campo ULT_TARIFA alterado de 4 para 5
        If mfPosCampo(2, "�ltima Tarifa.", .Ult_Tarifa, 5, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Indicador de condi��es gerais", .Ind_CondGerais, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        'Tarifa 10.00 - campo Cod_Produto alterado de 3 para 4
        If mfPosCampo(2, "C�digo do produto", .Cod_Produto, 4, 0, pMensagem) = False Then
            Exit Function
        End If

        If mfPosCampo(2, "Indicador de reentrada", .Ind_Reentrada, 1, 0, pMensagem) = False Then
            Exit Function
        End If

        If mfPosCampo(2, "Nosso Numero original", .Nosso_Numero_Original, 20, 0, pMensagem) = False Then
            Exit Function
        End If

        If mfPosCampo(2, "Numero de PI Emissao", .Num_PI_Emissao, 10, 0, pMensagem) = False Then
            Exit Function
        End If
        '1200178 - Telefone 9 Digitos
        If mfPosCampo(2, "Telefone", .Num_Fone, 9, 0, pMensagem) = False Then
            Exit Function
        End If
        'Conta com 15 digitos
        If mfPosCampo(2, "N�mero da conta de d�bito - 15 NUMEROS", Left(Format(.Num_Cta_Bas_Cor, "00000000"), 15), 15, 0, pMensagem) = False Then
            Exit Function
        End If

        mRegistro = UCase$(mRegistro)
        Print #pNumArqTransmissao, mRegistro
    End With
    mfTraPedido = True
End Function

Private Function mfTraCtrlEmis(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                               ByRef pMensagem As String) As Boolean

'Fun��o: grava registro de controle de emiss�o

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.


'1. Grava registro de dados de controle de emiss�o
'mRegistro = "TAB_CTRL_EMIS       " 'Indicador."
    mRegistro = "002"                   'Indicador."

    With pobjstcPedido.objstcCtrlEmis
        If mfPosCampo(4, "Nosso n�mero", .nosso_numero, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "N�mero do PI", .NUM_PI, 10, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "N�mero do bloqueto sem d�gito verificador", .Num_Bloq, 7, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo da unidade produtiva", .COD_UNID, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do ramo", .Cod_Ramo, 3, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data de in�cio da vig�ncia", .Dat_Ini_Vig, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data de t�rmino da vig�ncia", .Dat_Ter_Vig, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do corretor 1", .Cod_Corr, 7, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Percentual do corretor 1", .Per_Corr, 4, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do corretor 2", .Cod_Corr_2, 7, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Percentual do corretor 2", .Per_Corr_2, 4, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do corretor 3", .Cod_Corr_3, 7, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Percentual do corretor 3", .Per_Corr_3, 4, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Nome do segurado", .Nom_Seg, 60, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Quantidade de anexos", .Qtd_Anex, 5, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "N�mero da ap�lice anterior", .Num_Apol_Ant, 10, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Observa��o", .Dsc_Obs_Prop, 42, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Tipo de cota��o", .Tip_Cotac, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "N�mero da cota��o", .Num_Cotac, 12, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Percentual de juros negociafo", .Per_Juros_Neg, 4, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "C�digo da etiqueta", .Cod_Envi_Doc, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data da digita��o", .DAT_DIGIT, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Hora da digita��o", .HOR_DIGIT, 6, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do status da a��o", .Cod_Iden_Stat_Acao, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "C�digo do usu�rio de digitaliza��o", .Cod_User_Digit, 12, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data da digita��o mainframe", .Dat_Digit_MF, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Hora da digita��o mainframe", .Hor_Digit_MF, 6, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "C�digo do usu�rio de digitaliza��o mainframe", .Cod_User_Digit_MF, 12, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Valor do pr�mio", .Val_Pr_Tot, 9, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Classe de norma", .Cod_Iden_Classe, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo de cancelamento do cobran�a", .Cod_Iden_Cobr, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Numero da proposta", .Num_Prop, 6, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data de cancelamento do documento", .Dat_Canc_Apol, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Tipo de documento", .Cod_Tip_Doc, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data de recebimento na companhia", .Dat_Rec_Cia, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Numero de refer�ncia do corretor", .Num_Ref_Corr, 15, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Data de emiss�o do documento", .Dat_Emis, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "N�mero da ap�lice", .Num_Apol, 10, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "N�mero do endosso", .Num_Endo, 6, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Valor do pr�mio emitido", .Val_Pr_Emis, 9, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Valor do pr�mio l�quido", .Val_Pr_Liq, 9, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Percentual total de cosseguro", .Per_Tot_Coss, 4, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Quantidade de parcelas", .Qtd_Parc, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Status da resposta", .Cod_Iden_Stat_Resp, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo identificador de envio ao SAC", .Cod_Iden_SAC, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "C�digo do produtor", .Cod_Prod, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Valor da IS", .Val_IS, 15, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Percentual do desconto comercial", .Per_Desc_Coml, 4, 2, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(4, "Numero do cgc/cpf", .Num_Cgc_Cpf, 14, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Tipo de pessoa", .Cod_Iden_Pessoa, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "Pagamento antecipado", .Cod_Iden_Pag_Antec, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        If mfPosCampo(2, "D�gito do bloqueto", .DIG_BLOQUETO, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        mRegistro = UCase$(mRegistro)
        Print #pNumArqTransmissao, mRegistro
        mRegistro = ""
    End With

    mfTraCtrlEmis = True

End Function
Private Function mfTraCorretores(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                                 ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de corretores para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lobjstcCorretor As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)

    mfTraCorretores = False

    '1. Grava registro de dados corretores.
    For Each lobjstcCorretor In pobjstcPedido.ColecaoPedCorr
        'mRegistro = "TAB_PED_CORR        "      'Indicador.
        mRegistro = "003"                        'Indicador.
        If mfPosCampo(4, "Nosso n�mero", lobjstcCorretor.nosso_numero, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do org�o", lobjstcCorretor.Cod_Org, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do ramo", lobjstcCorretor.Cod_Ramo, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do pedido", lobjstcCorretor.Num_Ped, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero da ap�lice", lobjstcCorretor.Num_Apol, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do endosso", lobjstcCorretor.Num_Endo, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do corretor", lobjstcCorretor.Cod_Corr, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Percentual do corretor", lobjstcCorretor.Per_Corr, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero da ordem do corretor", lobjstcCorretor.Num_Ordem_Corr, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do PI", lobjstcCorretor.NUM_PI, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        mRegistro = UCase$(mRegistro)
        Print #pNumArqTransmissao, mRegistro
        mRegistro = ""
    Next lobjstcCorretor
    mfTraCorretores = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcCorretor = Nothing
    Return
End Function

Private Function mfTraTextoEndosso(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                                   ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de corretores para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lobjstcTextoEndo As stcA46V722B       'Objeto da classe de estrutura stcA46V705B (TAB_PED_DIZER)

    mfTraTextoEndosso = False

    '1. Grava registro de dados corretores.
    For Each lobjstcTextoEndo In pobjstcPedido.ColecaoTextoEndo
        'mRegistro = "TAB_PED_DIZER       "      'Indicador.
        mRegistro = "004"                       'Indicador.
        If mfPosCampo(4, "Nosso n�mero", lobjstcTextoEndo.nosso_numero, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do org�o", lobjstcTextoEndo.Cod_Org, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do ramo", lobjstcTextoEndo.Cod_Ramo, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do pedido", lobjstcTextoEndo.Num_Ped, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero da ap�lice", lobjstcTextoEndo.Num_Apol, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do endosso", lobjstcTextoEndo.Num_Endo, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo da Linha", lobjstcTextoEndo.NUM_SEQ, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Conteudo da Linha", lobjstcTextoEndo.Dsc_Dizer_Endo, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        mRegistro = UCase$(mRegistro)
        Print #pNumArqTransmissao, mRegistro
        mRegistro = ""
    Next lobjstcTextoEndo

    mfTraTextoEndosso = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcTextoEndo = Nothing
    Return
End Function

Private Function mfTraVenc(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                           ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de vencimento para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lObjstcVenc As stcA46V713B      'Objeto da classe de estrutura stcA46V713B (TAB_PED_VENC)

    mfTraVenc = False

    For Each lObjstcVenc In pobjstcPedido.ColecaoPedVenc
        'mRegistro = "TAB_PED_VENC        "      'Indicador."
        mRegistro = "005"                        'Indicador."
        If mfPosCampo(4, "Nosso n�mero", lObjstcVenc.nosso_numero, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do org�o", lObjstcVenc.Cod_Org, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do ramo", lObjstcVenc.Cod_Ramo, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do pedido", lObjstcVenc.Num_Ped, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero da ap�lice", lObjstcVenc.Num_Apol, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do endosso", lObjstcVenc.Num_Endo, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero da parcela", lObjstcVenc.Num_Parc, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Data do vencimento", lObjstcVenc.DAT_VENC, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        mRegistro = UCase$(mRegistro)
        Print #pNumArqTransmissao, mRegistro
    Next lObjstcVenc
    mfTraVenc = True

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lObjstcVenc = Nothing
    Return
End Function

Private Function mfTraJuros(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                            ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de juros para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lobjstcJuros As stcA46V711B    'Objeto da classe de estrutura stcA46V711B (TAB_PED_JUROS)

    mfTraJuros = False

    For Each lobjstcJuros In pobjstcPedido.ColecaoPedJuros
        'mRegistro = "TAB_PED_JUROS       "      'Indicador."
        mRegistro = "006"                        'Indicador."
        If mfPosCampo(4, "Nosso n�mero", lobjstcJuros.nosso_numero, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero da parcela", lobjstcJuros.Num_Parcela, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Valor do juros para pagamento antecipado", lobjstcJuros.Val_Ant, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Valor do juros para pagamento postecipado", lobjstcJuros.Val_Pos, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Tipo Compensacao", lobjstcJuros.Tip_Compensacao, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        mRegistro = UCase$(mRegistro)
        Print #pNumArqTransmissao, mRegistro
    Next lobjstcJuros
    mfTraJuros = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcJuros = Nothing
    Return
End Function

Private Function mfTraCosseguro(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                                ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de cosseguros para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lobjstcCosseguro As stcA46V706B     'Objeto da classe de estrutura stcA46V706B (TAB_PED_COSS)
    mfTraCosseguro = False

    For Each lobjstcCosseguro In pobjstcPedido.ColecaoPedCoss
        'mRegistro = "TAB_PED_COSS        "      'Indicador."
        mRegistro = "007"                       'Indicador."
        If mfPosCampo(4, "Nosso n�mero", lobjstcCosseguro.nosso_numero, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do org�o", lobjstcCosseguro.Cod_Org, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do ramo", lobjstcCosseguro.Cod_Ramo, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do pedido", lobjstcCosseguro.Num_Ped, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero da ap�lice", lobjstcCosseguro.Num_Apol, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do endosso", lobjstcCosseguro.Num_Endo, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do Cosseguro", lobjstcCosseguro.Cod_Coss, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Percentual do cosseguro", lobjstcCosseguro.Per_Coss, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do PI", lobjstcCosseguro.NUM_PI, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        mRegistro = UCase$(mRegistro)
        Print #pNumArqTransmissao, mRegistro
    Next lobjstcCosseguro
    mfTraCosseguro = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcCosseguro = Nothing
    Return
End Function

Private Function mfTraEndereco(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                               ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de endere�o para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    mfTraEndereco = False

    '1. Grava registro
    'mRegistro = "TAB_PED_ENDER       "      'Indicador."
    mRegistro = "008"                        'Indicador."
    If mfPosCampo(4, "Nosso n�mero", pobjstcPedido.objstcPedEndereco.nosso_numero, 20, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(2, "C�digo do org�o", pobjstcPedido.objstcPedEndereco.Cod_Org, 2, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(2, "C�digo do ramo", pobjstcPedido.objstcPedEndereco.Cod_Ramo, 3, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(2, "N�mero do pedido", pobjstcPedido.objstcPedEndereco.Num_Ped, 6, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(2, "N�mero da ap�lice", pobjstcPedido.objstcPedEndereco.Num_Apol, 10, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(2, "N�mero do endosso", pobjstcPedido.objstcPedEndereco.Num_Endo, 6, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(4, "Tipo de logradouro", pobjstcPedido.objstcPedEndereco.Cod_Tip_Logr, 3, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(4, "Logradouro", pobjstcPedido.objstcPedEndereco.Nom_Logr, 50, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(4, "Complemento do logradouro", pobjstcPedido.objstcPedEndereco.Dsc_Compl_Logr, 20, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(4, "Bairro", pobjstcPedido.objstcPedEndereco.Nom_Bairro, 30, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(4, "Cidade", pobjstcPedido.objstcPedEndereco.Nom_Cid, 30, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(4, "Sigla da UF", pobjstcPedido.objstcPedEndereco.Sig_UF, 2, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(2, "CEP", pobjstcPedido.objstcPedEndereco.Num_CEP, 8, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(2, "N�mero do telefone", pobjstcPedido.objstcPedEndereco.Num_Fone, 7, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(2, "N�mero do PI", pobjstcPedido.objstcPedEndereco.NUM_PI, 10, 0, pMensagem) = False Then
        Exit Function
    End If
    'Ficha 1000307
    If mfPosCampo(2, "C�digo registro do conselho", pobjstcPedido.objstcPedEndereco.Cod_Reg_Cons, 2, 0, pMensagem) = False Then
        Exit Function
    End If
    If mfPosCampo(4, "N�mero registro no conselho", pobjstcPedido.objstcPedEndereco.Num_Reg_Cons, 10, 0, pMensagem) = False Then
        Exit Function
    End If

    '1200178 - Telefone 9 Digitos
    If mfPosCampo(2, "N�mero do telefone", pobjstcPedido.objstcPedEndereco.Num_Fone, 9, 0, pMensagem) = False Then
        Exit Function
    End If

    mRegistro = UCase$(mRegistro)
    Print #pNumArqTransmissao, mRegistro

    mfTraEndereco = True
End Function


Private Function mfTraItens(pNomeArquivo As String, pHoraEnvio As String, ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                            ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de itens para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lObjstcItens As stcA46V708B   'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjTrans As clsA46V718A    'Objeto da classe clsA46V718A (TAB_PED_TRANSM)
    Dim lobjstcTrans As stcA46V718B   'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM)
    mfTraItens = False

    Set lobjTrans = New clsA46V718A
    Set lobjTrans.M46V999 = M46V999

    Set lobjstcTrans = New stcA46V718B
    lobjstcTrans.nosso_numero = pobjstcPedido.nosso_numero
    If Not lobjTrans.mfExcluir(lobjstcTrans, "", "") Then
        pMensagem = lobjTrans.UltimoErro
        GoTo SAIDA
    End If

    For Each lObjstcItens In pobjstcPedido.ColecaoPedLoc
        'mRegistro = "TAB_PED_LOC         "      'Indicador."
        mRegistro = "009"                        'Indicador."
        If mfPosCampo(4, "Nosso n�mero", lObjstcItens.nosso_numero, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do org�o", lObjstcItens.Cod_Org, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do ramo", lObjstcItens.Cod_Ramo, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do pedido", lObjstcItens.Num_Ped, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero da ap�lice", lObjstcItens.Num_Apol, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do endosso", lObjstcItens.Num_Endo, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do endosso", lObjstcItens.Num_Item, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Tipo de logradouro - Risco", lObjstcItens.Cod_Tip_Logr_Risco, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Logradouro - Risco", lObjstcItens.Nom_Logr_Risco, 50, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Complemento do logradouro - Risco", lObjstcItens.Dsc_Compl_Risco, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Nome do bairro - Risco", lObjstcItens.Nom_Bairro_Risco, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Nome dad cidade - Risco", lObjstcItens.Nom_Cid_Risco, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Sigla UF - Risco", lObjstcItens.Sig_UF_Risco, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "CEP - Risco", lObjstcItens.Num_Cep_Risco, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Percentual desconto - Apar", lObjstcItens.Per_Desc_Apar, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Percentual desconto - Progres", lObjstcItens.Per_Desc_Progres, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Percentual IS - VR", lObjstcItens.Per_IS_Vr, 5, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Valor do coeficiente IS - VR", lObjstcItens.Val_Coef_IS_Vr, 5, 4, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Valor VR", lObjstcItens.Val_Vr, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Tab clasf", lObjstcItens.Tab_Clasf, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Codigo Clasf", lObjstcItens.Cod_Clasf, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Codigo do complemento Clasf", lObjstcItens.Cod_Compl_Clasf, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Descri��o Clasf", lObjstcItens.Dsc_Clasf, 60, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Codigo da classe de ocup", lObjstcItens.Cod_Classe_Ocup, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Codigo constru��o", lObjstcItens.Cod_Constr, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Usu�rio de digita��o", lObjstcItens.Cod_User_Digit, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Data da digita��o", lObjstcItens.DAT_DIGIT, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Hora da digita��o", lObjstcItens.HOR_DIGIT, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Usu�rio de libera��o", lObjstcItens.Cod_User_Liber, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Data da libera��o", lObjstcItens.Dat_Liber, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Hora da libera��o", lObjstcItens.Hor_Liber, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Situa��o do item", lObjstcItens.Cod_Sit_item, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Percentual de b�nus", lObjstcItens.Per_Bonus, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Quantidade de empreg", lObjstcItens.Qtd_Empreg, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero INSP", lObjstcItens.Num_Insp, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de seguran�a 1", lObjstcItens.Cod_Iden_Segur_01_V, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de seguran�a 2", lObjstcItens.Cod_Iden_Segur_02_A, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de seguran�a 3", lObjstcItens.Cod_Iden_Segur_03_P, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de seguran�a 4", lObjstcItens.Cod_Iden_Segur_04_I, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de seguran�a 5", lObjstcItens.Cod_Iden_Segur_05_G, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de seguran�a 6", lObjstcItens.Cod_Iden_Segur_06, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de seguran�a 7", lObjstcItens.Cod_Iden_Segur_07, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de seguran�a 8", lObjstcItens.Cod_Iden_Segur_08, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de seguran�a 9", lObjstcItens.Cod_Iden_Segur_09, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de seguran�a 10", lObjstcItens.Cod_Iden_Segur_10, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Percentual total de seguran�a", lObjstcItens.Per_Tot_Segur, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de prote��o 1", lObjstcItens.Cod_Iden_Protec_01_E, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de prote��o 2", lObjstcItens.Cod_Iden_Protec_02_H, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de prote��o 3", lObjstcItens.Cod_Iden_Protec_03_S, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de prote��o 4", lObjstcItens.Cod_Iden_Protec_04_D, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de prote��o 5", lObjstcItens.Cod_Iden_Protec_05, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de prote��o 6", lObjstcItens.Cod_Iden_Protec_06, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de prote��o 7", lObjstcItens.Cod_Iden_Protec_07, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de prote��o 8", lObjstcItens.Cod_Iden_Protec_08, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de prote��o 9", lObjstcItens.Cod_Iden_Protec_09, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Identifica��o de prote��o 10", lObjstcItens.Cod_Iden_Protec_10, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Percentual total de prote��o", lObjstcItens.Per_Tot_Protec, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do produto", lObjstcItens.Cod_Produto, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo da regi�o", lObjstcItens.Cod_Regiao, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If

        'Corretor tem somente produto par
        If lObjstcItens.Cod_Iden_Tarif Mod 2 = 1 Then
            lObjstcItens.Cod_Iden_Tarif = lObjstcItens.Cod_Iden_Tarif - 1
        End If

        If mfPosCampo(2, "C�digo de identifica��o da tarifa", lObjstcItens.Cod_Iden_Tarif, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do PI", lObjstcItens.NUM_PI, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Tipo de emiss�o", lObjstcItens.Tip_Emissao, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Ap�lice Anterior", lObjstcItens.Num_Apol_Anterior, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Item Anterior", lObjstcItens.Num_Item_Anterior, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Ind_Sinistro", lObjstcItens.Ind_Sinistro, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If

        '1000347 - Inclus�o de informa��es da cong�nere em renova��es
        If mfPosCampo(2, "c�digo da seguradora", lObjstcItens.Ren_Con_Seguradora, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If

        '1000347 - Inclus�o de informa��es da cong�nere em renova��es
        If mfPosCampo(4, "n�mero da ap�lice de renova��o", lObjstcItens.Ren_Con_Apolice, 15, 0, pMensagem) = False Then
            GoTo SAIDA
        End If

        '1000347 - Inclus�o de informa��es da cong�nere em renova��es
        If mfPosCampo(4, "n�mero do item da ap�lice de renova��o", lObjstcItens.Ren_Con_Item, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If

        'Ficha 1000305
        If mfPosCampo(2, "Tip_Cobert", lObjstcItens.Tip_Cobert, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If

        If mfPosCampo(2, "Insp_Obr ", lObjstcItens.Insp_Obr, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If

        If mfPosCampo(2, "Insp_Agend", lObjstcItens.Insp_Agend, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If

        If mfPosCampo(2, "Dig_Bens", lObjstcItens.Dig_Bens, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If

        If mfPosCampo(2, "Cob_Bens", lObjstcItens.Cob_Bens, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If

        If mfPosCampo(2, "Ind_Sin_Inspec", lObjstcItens.Ind_Sin_Inspec, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Cod_DizerItem", lObjstcItens.Cod_DizerItem, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If

        mRegistro = UCase$(mRegistro)
        Print #pNumArqTransmissao, mRegistro
        mRegistro = ""

        Set lobjstcTrans = New stcA46V718B
        With lobjstcTrans
            .nosso_numero = lObjstcItens.nosso_numero
            .Num_Ped = Format(pobjstcPedido.Num_Ped, "000000")
            .Num_Item = Format(lObjstcItens.Num_Item, "00000")
            .Nom_Arq_Envio = pNomeArquivo
            .Dat_Hor_Envio = pHoraEnvio
            .Nom_Arq_Retorno = ""
            .Dat_Hor_Retorno = ""
            .Ind_Proposta = 1
            .Nom_Segurado = pobjstcPedido.Nom_Segurado
            .Cod_Ramo = Format(pobjstcPedido.Cod_Ramo, "000")
            .Dat_Ini_Vig = pobjstcPedido.Dat_Ini_Vig
            .Dat_Ter_Vig = pobjstcPedido.Dat_Ter_Vig
            .Cod_Tip_Logr_Risco = lObjstcItens.Cod_Tip_Logr_Risco
            .Nom_Logr_Risco = lObjstcItens.Nom_Logr_Risco
            .Dsc_Compl_Risco = lObjstcItens.Dsc_Compl_Risco
            .Nom_Bairro_Risco = lObjstcItens.Nom_Bairro_Risco
            .Nom_Cid_Risco = lObjstcItens.Nom_Cid_Risco
            .Sig_UF_Risco = lObjstcItens.Sig_UF_Risco
            .Num_Cep_Risco = lObjstcItens.Num_Cep_Risco
            .Num_Risco = lObjstcItens.objstcPedInsp.Num_Logr_Risco
            .Tip_Emissao = Format(lObjstcItens.Tip_Emissao, "000")
        End With

        '********N�o atualiza a tabela de transmiss�o
        '**Esta informa��o � atualziada somente no retorno do protocolo.
        '**pelo programa P28V400.exe
        '        If Not lobjtrans.mfIncluir(lobjstcTrans) Then
        '            pMensagem = lobjtrans.UltimoErro
        '           goto SAIDA
        '        End If
    Next lObjstcItens
    mfTraItens = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lObjstcItens = Nothing
    Set lobjTrans = Nothing
    Set lobjstcTrans = Nothing
    Return
End Function
Private Function mfTraClausulas(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                                ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de cl�usulas para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lObjstcItens As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcPedClau As stcA46V703B    'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)

    mfTraClausulas = False

    '1. Grava registro de dados de cl�usula.
    For Each lObjstcItens In pobjstcPedido.ColecaoPedLoc
        For Each lobjstcPedClau In lObjstcItens.ColecaoPedClau
            'With lObjstcItens.objstcPedClau
            With lobjstcPedClau
                If M46V999.gfPreenchido(.nosso_numero) Then
                    'mRegistro = "TAB_PED_CLAU        "      'Indicador."
                    mRegistro = "010"                        'Indicador.
                    If mfPosCampo(4, "Nosso n�mero", .nosso_numero, 20, 0, pMensagem) = False Then
                        GoTo SAIDA
                    End If
                    If mfPosCampo(2, "C�digo do org�o", .Cod_Org, 2, 0, pMensagem) = False Then
                        GoTo SAIDA
                    End If
                    If mfPosCampo(2, "C�digo do ramo", .Cod_Ramo, 3, 0, pMensagem) = False Then
                        GoTo SAIDA
                    End If
                    If mfPosCampo(2, "N�mero do pedido", .Num_Ped, 6, 0, pMensagem) = False Then
                        GoTo SAIDA
                    End If
                    If mfPosCampo(2, "N�mero da ap�lice", .Num_Apol, 10, 0, pMensagem) = False Then
                        GoTo SAIDA
                    End If
                    If mfPosCampo(2, "N�mero do endosso", .Num_Endo, 6, 0, pMensagem) = False Then
                        GoTo SAIDA
                    End If
                    If mfPosCampo(2, "N�mero do Item", .Num_Item, 5, 0, pMensagem) = False Then
                        GoTo SAIDA
                    End If
                    If mfPosCampo(2, "C�digo da Cl�usula", .Cod_Clau, 5, 0, pMensagem) = False Then
                        GoTo SAIDA
                    End If
                    If mfPosCampo(2, "N�mero do PI", .NUM_PI, 10, 0, pMensagem) = False Then
                        GoTo SAIDA
                    End If
                    If mfPosCampo(4, "Descri��o", .Descricao, 255, 0, pMensagem) = False Then
                        GoTo SAIDA
                    End If
                    mRegistro = UCase$(mRegistro)
                    Print #pNumArqTransmissao, mRegistro
                End If
            End With
        Next lobjstcPedClau
    Next lObjstcItens

    mfTraClausulas = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcPedClau = Nothing
    Set lObjstcItens = Nothing
    Return
End Function

Private Function mfTraInspecao(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                               ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de inspe��o

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lObjstcItens As stcA46V708B             'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)

    mfTraInspecao = False

    '1. Grava registro de dados de inspe��o.
    For Each lObjstcItens In pobjstcPedido.ColecaoPedLoc
        With lObjstcItens.objstcPedInsp
            'mRegistro = "TAB_PED_INSP        "      'Indicador."
            mRegistro = "011"                        'Indicador."
            If mfPosCampo(4, "Nosso n�mero", .nosso_numero, 20, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "C�digo do org�o", .Cod_Org, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "C�digo do ramo", .Cod_Ramo, 3, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do pedido", .Num_Ped, 6, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do Item", .Num_Item, 5, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero da inspe��o", .Num_Insp, 7, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Data do cadastro", .Dat_Cadastro, 8, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Nome do contato", .Pess_Cont, 40, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Telefone de contato", .Num_Tel, 14, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do logradouro - risco", .Num_Logr_Risco, 8, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            mRegistro = UCase$(mRegistro)
            Print #pNumArqTransmissao, mRegistro
        End With
    Next lObjstcItens
    mfTraInspecao = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lObjstcItens = Nothing
    Return
End Function

Private Function mfTraObservacao(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                                 ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de observa��o

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lObjstcItens As stcA46V708B     'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)

    mfTraObservacao = False

    '1. Grava registro de dados de observa��o.
    For Each lObjstcItens In pobjstcPedido.ColecaoPedLoc
        With lObjstcItens.objstcPedObs
            'mRegistro = "TAB_PED_OBS         "      'Indicador."
            mRegistro = "012"                        'Indicador."
            If mfPosCampo(4, "Nosso n�mero", .nosso_numero, 20, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do Item", .Num_Item, 5, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Observa��o", .Observacao, 500, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            mRegistro = UCase$(mRegistro)
            Print #pNumArqTransmissao, mRegistro
        End With
    Next lObjstcItens
    mfTraObservacao = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lObjstcItens = Nothing
    Return
End Function

Private Function mfTraSeguro(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                             ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de seguros para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lObjstcItens As stcA46V708B         'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcSeguro As stcA46V716B        'Objeto da classe de estrutura stcA46V716B (TAB_PED_LOC_SEGURO)

    mfTraSeguro = False


    For Each lObjstcItens In pobjstcPedido.ColecaoPedLoc
        For Each lobjstcSeguro In lObjstcItens.ColecaoPedSeguro
            With lobjstcSeguro
                'mRegistro = "TAB_PED_LOC_SEGURO  "      'Indicador."
                mRegistro = "013"                        'Indicador."
                If mfPosCampo(4, "Nosso n�mero", .nosso_numero, 20, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "C�digo do org�o", .Cod_Org, 2, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "C�digo do ramo", .Cod_Ramo, 3, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do pedido", .Num_Ped, 6, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero da ap�lice", .Num_Apol, 10, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do endosso", .Num_Endo, 6, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do Item", .Num_Item, 5, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(4, "Descri��o do outro seguro", .Dsc_Outro_Seguro, 35, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "C�digo da seguradora", .Cod_Seguradora, 4, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero da ordem", .Num_Ordem, 4, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(4, "N�mero da ap�lice anterior", .Num_Apol_Ant, 20, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do item anterior", .Num_Item_Ant, 5, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Indicador de sinistro", .Ind_Sinistro, 1, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                mRegistro = UCase$(mRegistro)
                Print #pNumArqTransmissao, mRegistro
            End With
        Next lobjstcSeguro
    Next lObjstcItens
    mfTraSeguro = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcSeguro = Nothing
    Set lObjstcItens = Nothing
    Return
End Function

Private Function mfTraBens(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                           ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de bens para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lObjstcItens As stcA46V708B         'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lObjstcBens As stcA46V715B          'Objeto da classe de estrutura stcA46V715B (TAB_PED_LOC_BEM)

    mfTraBens = False


    For Each lObjstcItens In pobjstcPedido.ColecaoPedLoc
        For Each lObjstcBens In lObjstcItens.ColecaoPedLocBem
            With lObjstcBens
                'mRegistro = "TAB_PED_LOC_BEM     "      'Indicador."
                mRegistro = "014"                        'Indicador."

                If mfPosCampo(4, "Nosso n�mero", .nosso_numero, 20, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "C�digo do org�o", .Cod_Org, 2, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "C�digo do ramo", .Cod_Ramo, 3, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do pedido", .Num_Ped, 6, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero da ap�lice", .Num_Apol, 10, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do endosso", .Num_Endo, 6, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do Item", .Num_Item, 5, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(4, "Descri��o do bem segurado", .Dsc_Bem, 60, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Valor do bem declarado", .Val_Bem, 15, 2, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "C�digo do bem declarado", .Cod_Bem, 3, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "C�digo da Sequencia", .Cod_Seq, 3, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                mRegistro = UCase$(mRegistro)
                Print #pNumArqTransmissao, mRegistro
            End With
        Next lObjstcBens
    Next lObjstcItens
    mfTraBens = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lObjstcBens = Nothing
    Set lObjstcItens = Nothing
    Return
End Function

Private Function mfTraPessoa(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                             ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de pessoas para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lObjstcItens As stcA46V708B         'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcPessoa As stcA46V735B          'Objeto da classe de estrutura stcA46V735B

    mfTraPessoa = False


    For Each lObjstcItens In pobjstcPedido.ColecaoPedLoc
        For Each lobjstcPessoa In lObjstcItens.ColecaoBeneficiario
            With lobjstcPessoa
                'mRegistro = "TAB_PED_LOC_PESSOA  - BENEFICIARIO  "      'Indicador."
                mRegistro = "022"                        'Indicador."

                If mfPosCampo(4, "Nosso n�mero", .nosso_numero, 20, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero da ap�lice", .Num_Apol, 10, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do endosso", .Num_Endo, 6, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do Item", .Num_Item, 5, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do PI", .NUM_PI, 10, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If

                If mfPosCampo(2, "C�digo Unico ID Pessoa", .Ped_Loc_Pessoa_ID, 10, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Tipo de Pessoa - Inq, Prop, Bene", .Tip_Loc_Pessoa, 2, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(4, "Nome Pessoa", .Nom_Pessoa, 60, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Tipo Pessoa, Fisi, Juri, Instit", .Tip_Pessoa, 2, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(4, "N�mero do CPF - CNPJ", .Num_CPF_CNPJ, 14, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Tipo de Cobertura", .Tip_Cobert, 2, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Percentual de Participa��o - Bene", .Per_Participacao, 5, 2, pMensagem) = False Then
                    GoTo SAIDA
                End If


                mRegistro = UCase$(mRegistro)
                Print #pNumArqTransmissao, mRegistro
            End With
        Next lobjstcPessoa

        If M46V999.gfPreenchido(lObjstcItens.objstcPedInquilino.Num_CPF_CNPJ) And Trim(lObjstcItens.objstcPedInquilino.Num_CPF_CNPJ) <> "0" Then
            With lObjstcItens.objstcPedInquilino
                'mRegistro = "TAB_PED_LOC_PESSOA  - INQUILINO  "      'Indicador."
                mRegistro = "023"                        'Indicador."

                If mfPosCampo(4, "Nosso n�mero", .nosso_numero, 20, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero da ap�lice", .Num_Apol, 10, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do endosso", .Num_Endo, 6, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do Item", .Num_Item, 5, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do PI", .NUM_PI, 10, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If

                If mfPosCampo(2, "C�digo Unico ID Pessoa", .Ped_Loc_Pessoa_ID, 10, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Tipo de Pessoa - Inq, Prop, Bene", .Tip_Loc_Pessoa, 2, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(4, "Nome Pessoa", .Nom_Pessoa, 60, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Tipo Pessoa, Fisi, Juri, Instit", .Tip_Pessoa, 2, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(4, "N�mero do CPF - CNPJ", .Num_CPF_CNPJ, 14, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Tipo de Cobertura", .Tip_Cobert, 2, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Percentual de Participa��o - Bene", .Per_Participacao, 5, 2, pMensagem) = False Then
                    GoTo SAIDA
                End If


                mRegistro = UCase$(mRegistro)
                Print #pNumArqTransmissao, mRegistro
            End With
        End If

        If M46V999.gfPreenchido(lObjstcItens.objstcPedProprietario.Num_CPF_CNPJ) And Trim(lObjstcItens.objstcPedProprietario.Num_CPF_CNPJ) <> "0" Then
            With lObjstcItens.objstcPedProprietario
                'mRegistro = "TAB_PED_LOC_PESSOA  - PROPRIETARIO  "      'Indicador."
                mRegistro = "024"                        'Indicador."

                If mfPosCampo(4, "Nosso n�mero", .nosso_numero, 20, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero da ap�lice", .Num_Apol, 10, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do endosso", .Num_Endo, 6, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do Item", .Num_Item, 5, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do PI", .NUM_PI, 10, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If

                If mfPosCampo(2, "C�digo Unico ID Pessoa", .Ped_Loc_Pessoa_ID, 10, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Tipo de Pessoa - Inq, Prop, Bene", .Tip_Loc_Pessoa, 2, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(4, "Nome Pessoa", .Nom_Pessoa, 60, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Tipo Pessoa, Fisi, Juri, Instit", .Tip_Pessoa, 2, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(4, "N�mero do CPF - CNPJ", .Num_CPF_CNPJ, 14, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Tipo de Cobertura", .Tip_Cobert, 2, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "Percentual de Participa��o - Bene", .Per_Participacao, 5, 2, pMensagem) = False Then
                    GoTo SAIDA
                End If


                mRegistro = UCase$(mRegistro)
                Print #pNumArqTransmissao, mRegistro
            End With
        End If


    Next lObjstcItens





    mfTraPessoa = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcPessoa = Nothing

    Return
End Function


Private Function mfTraFatoresProtecao(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                                      ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de bens para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lObjstcItens As stcA46V708B         'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lObjstcFatores As stcA46V720B       'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO)

    mfTraFatoresProtecao = False

    For Each lObjstcItens In pobjstcPedido.ColecaoPedLoc
        For Each lObjstcFatores In lObjstcItens.ColecaoPedFator
            With lObjstcFatores
                'mRegistro = "TAB_PED_FAT_RISCO   "      'Indicador."
                mRegistro = "015"                        'Indicador."
                If mfPosCampo(4, "Nosso n�mero", .nosso_numero, 20, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "C�digo do org�o", .Cod_Org, 2, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "C�digo do ramo", .Cod_Ramo, 3, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do pedido", .Num_Ped, 6, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero da ap�lice", .Num_Apol, 10, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do endosso", .Num_Endo, 6, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do Item", .Num_Item, 5, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "C�digo do Fator", .Cod_Fator, 4, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(4, "Descri��o do Fator", .Dsc_Compl_Fator, 30, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If mfPosCampo(2, "N�mero do PI", .NUM_PI, 10, 0, pMensagem) = False Then
                    GoTo SAIDA
                End If
                mRegistro = UCase$(mRegistro)
                Print #pNumArqTransmissao, mRegistro
            End With
        Next lObjstcFatores
    Next lObjstcItens
    mfTraFatoresProtecao = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lObjstcFatores = Nothing
    Set lObjstcItens = Nothing
    Return
End Function


Private Function mfTraCoberturas(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                                 ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de coberturas para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lObjstcItens As stcA46V708B             'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcCobertura As stcA46V704B         'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)


    mfTraCoberturas = False


    For Each lObjstcItens In pobjstcPedido.ColecaoPedLoc
        For Each lobjstcCobertura In lObjstcItens.ColecaoPedCobert
            'mRegistro = "TAB_PED_COBERT      "      'Indicador."
            mRegistro = "016"                        'Indicador."
            If mfPosCampo(4, "Nosso n�mero", lobjstcCobertura.nosso_numero, 20, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "C�digo do org�o", lobjstcCobertura.Cod_Org, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "C�digo do ramo", lobjstcCobertura.Cod_Ramo, 3, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do pedido", lobjstcCobertura.Num_Ped, 6, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero da ap�lice", lobjstcCobertura.Num_Apol, 10, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do endosso", lobjstcCobertura.Num_Endo, 6, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do Item", lobjstcCobertura.Num_Item, 5, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Tab Cobertura", lobjstcCobertura.Tab_Cobert, 4, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "C�digo da Cobertura", lobjstcCobertura.Cod_Cobert, 4, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Valor da IS", lobjstcCobertura.Val_IS, 15, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Percentual da Taxa", lobjstcCobertura.Per_Taxa, 7, 5, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Percentual de desconto comercial", lobjstcCobertura.Per_Desc_Coml, 4, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Percentual de desconto Exp", lobjstcCobertura.Per_Desc_Exp, 4, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Valor do pr�mio", lobjstcCobertura.VAL_PR, 15, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Valor do pr�mio anual", lobjstcCobertura.Val_Pr_Anu, 15, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Valor do pr�mio Tarif", lobjstcCobertura.Val_Pr_Tarif, 15, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Valor do pr�mio Tarif anual", lobjstcCobertura.Val_Pr_Tarif_Anu, 15, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Valor do pr�mio Resg", lobjstcCobertura.Val_Pr_Resg, 15, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Valor do pr�mio Resg anual", lobjstcCobertura.Val_Pr_Resg_Anu, 15, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Percentual da franquia", lobjstcCobertura.Per_Franq, 4, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Valor m�nimo da franquia", lobjstcCobertura.Val_Franq_Min, 15, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Valor m�ximo da franquia", lobjstcCobertura.Val_Franq_Max, 15, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do PI", lobjstcCobertura.NUM_PI, 10, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            mRegistro = UCase$(mRegistro)
            Print #pNumArqTransmissao, mRegistro
        Next lobjstcCobertura
    Next lObjstcItens
    mfTraCoberturas = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcCobertura = Nothing
    Set lObjstcItens = Nothing
    Return
End Function
Private Function mfTraLMI(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                          ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de coberturas para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lobjstcLMI As stcA46V732B             'Objeto da classe de estrutura stcA46V732B (TAB_PED_LMI)


    mfTraLMI = False


    For Each lobjstcLMI In pobjstcPedido.ColecaoLMI
        mRegistro = "020"                        'Indicador."
        If mfPosCampo(4, "Nosso n�mero", lobjstcLMI.nosso_numero, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do org�o", lobjstcLMI.Cod_Org, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo do ramo", lobjstcLMI.Cod_Ramo, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do pedido", lobjstcLMI.Num_Ped, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero da ap�lice", lobjstcLMI.Num_Apol, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do endosso", lobjstcLMI.Num_Endo, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Tab Cobertura", lobjstcLMI.Tab_Cobert, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "C�digo da Cobertura", lobjstcLMI.Cod_Cobert, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Valor da IS", lobjstcLMI.Val_IS, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Percentual da franquia", lobjstcLMI.Per_Franq, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Valor m�nimo da franquia", lobjstcLMI.Val_Franq_Min, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "Valor m�ximo da franquia", lobjstcLMI.Val_Franq_Max, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero do PI", lobjstcLMI.NUM_PI, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(2, "N�mero de meses do PI", lobjstcLMI.Num_Meses_PI, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If mfPosCampo(4, "Texto Franquia", lobjstcLMI.TextoFranquia, 200, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        mRegistro = UCase$(mRegistro)
        Print #pNumArqTransmissao, mRegistro
    Next lobjstcLMI
    mfTraLMI = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcLMI = Nothing
    Return
End Function

Private Function mfTraCobertCompl(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                                  ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de coberturas para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lObjstcItens As stcA46V708B             'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcCobertura As stcA46V704B         'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)


    mfTraCobertCompl = False


    For Each lObjstcItens In pobjstcPedido.ColecaoPedLoc
        For Each lobjstcCobertura In lObjstcItens.ColecaoPedCobert
            'mRegistro = "TAB_PED_COBERT_COMPL"      'Indicador."
            mRegistro = "017"                        'Indicador."
            If mfPosCampo(4, "Nosso n�mero", lobjstcCobertura.nosso_numero, 20, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "C�digo do org�o", lobjstcCobertura.Cod_Org, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "C�digo do ramo", lobjstcCobertura.Cod_Ramo, 3, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do pedido", lobjstcCobertura.Num_Ped, 6, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero da ap�lice", lobjstcCobertura.Num_Apol, 10, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do endosso", lobjstcCobertura.Num_Endo, 6, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do Item", lobjstcCobertura.Num_Item, 5, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Tab Cobertura", lobjstcCobertura.Tab_Cobert, 4, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "C�digo da Cobertura", lobjstcCobertura.Cod_Cobert, 4, 0, pMensagem) = False Then
                GoTo SAIDA
            End If

            If mfPosCampo(2, "Num_Meses_Pi", lobjstcCobertura.Num_Meses_PI, 4, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            'If mfPosCampo(4, "Formula", lobjstcCobertura.Formula, 100, 0, pMensagem) = False Then
            '    goto SAIDA
            'End If
            If mfPosCampo(4, "Formula", " ", 100, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Cod_Classe", lobjstcCobertura.COD_CLASSE, 4, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            'este valor val_fator_franquia est� errado pois tem 2 casas decimais e o correto � 5
            If mfPosCampo(2, "Val_Fator_Franq", lobjstcCobertura.Val_Fator_Franq, 15, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Alterar_Cobert", lobjstcCobertura.Alterar_Cobert, 10, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Val_Risco_Cob", lobjstcCobertura.Val_Risco_Cob, 15, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do PI", lobjstcCobertura.NUM_PI, 10, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            'Esta sendo repetido o valor de franquia pois agora est� com a quantidade de casas decimais correta
            If mfPosCampo(2, "Val_Fator_Franq", lobjstcCobertura.Val_Fator_Franq, 15, 5, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "AjusteEndosso", lobjstcCobertura.AjusteEndosso, 5, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Alteroufranquia", lobjstcCobertura.AlterouFranquia, 3, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "TextoFranquia", lobjstcCobertura.TextoFranquia, 200, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Val_Pr_Orig", lobjstcCobertura.Val_Pr_Orig, 15, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "PremioInformado", lobjstcCobertura.PremioInformado, 3, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            'ProjetoFormularios
            If mfPosCampo(4, "Compl_Desc_Cobertura", lobjstcCobertura.Compl_Desc_Cobertura, 50, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            'Regra de Batente
            If mfPosCampo(2, "Val_Pr_Batente", lobjstcCobertura.VAL_PR_BATENTE, 15, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            'Multiplo Franquia
            If mfPosCampo(2, "Per_Desc_Franq", lobjstcCobertura.Per_Desc_Franq, 4, 2, pMensagem) = False Then
                GoTo SAIDA
            End If


            mRegistro = UCase$(mRegistro)
            Print #pNumArqTransmissao, mRegistro
        Next lobjstcCobertura
    Next lObjstcItens
    mfTraCobertCompl = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcCobertura = Nothing
    Set lObjstcItens = Nothing
    Return
End Function

Private Function mfTraBeneficiario(ByVal pNumArqTransmissao As Integer, ByVal pobjstcPedido As stcA46V702B, _
                                   ByRef pMensagem As String) As Boolean
'Fun��o: grava registro de benefici�rios

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'                        pObjstcPedido........objeto de pedido.
'                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lobjstcTextoEspec As stcA46V726B    'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC)
    mfTraBeneficiario = False

    '1. Grava registro de dados de benefici�rio.
    'mRegistro = "TAB_PED_ESPEC       "      'Indicador.


    For Each lobjstcTextoEspec In pobjstcPedido.ColecaoTextoEspec
        With lobjstcTextoEspec
            mRegistro = "018"                       'Indicador.
            If mfPosCampo(4, "Nosso n�mero", .nosso_numero, 20, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "C�digo do org�o", .Cod_Org, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "C�digo do ramo", .Cod_Ramo, 3, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do pedido", .Num_Ped, 6, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero da ap�lice", .Num_Apol, 10, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do endosso", .Num_Endo, 6, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero de p�ginas", .Num_pag, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Quantidade de vias", .Qtd_vias, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 1", .Espec_lin1, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 2", .Espec_lin2, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 3", .Espec_lin3, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 4", .Espec_lin4, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 5", .Espec_lin5, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 6", .Espec_lin6, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 7", .Espec_lin7, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 8", .Espec_lin8, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 9", .Espec_lin9, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 10", .Espec_lin10, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 11", .Espec_lin11, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 12", .Espec_lin12, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 13", .Espec_lin13, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 14", .Espec_lin14, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 15", .Espec_lin15, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 16", .Espec_lin16, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 17", .Espec_lin17, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 18", .Espec_lin18, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 19", .Espec_lin19, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(4, "Texto - Linha 20", .Espec_lin20, 80, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "N�mero do PI", .NUM_PI, 10, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            If mfPosCampo(2, "Cod Espec", .Cod_Espec, 4, 0, pMensagem) = False Then
                GoTo SAIDA
            End If

            mRegistro = UCase$(mRegistro)
            Print #pNumArqTransmissao, mRegistro
        End With
    Next lobjstcTextoEspec
    mfTraBeneficiario = True
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lobjstcTextoEspec = Nothing
    Return
    'AQUIX!!!
End Function

Private Sub mpObtSeq(ByVal pEndLOG As String, _
                     ByRef pSequencia As String, _
                     Optional pNumMaxSeq As Integer = 99)
'Procedure: obt�m �ltima seq��ncia do arquivo de transmiss�o.

'Par�metro de entrada...pEndLOG...Endere�o do arquivo LOG <SEQ010.log>

'Par�metro de sa�da.....pSequencia........�ltima seq��ncia do arquivo de transmiss�o.

    Dim lNumLog As Integer  'Utilizado com FreeFile para .LOG.
    Dim lPosicao As Integer    'Utilizado com InStr.
    Dim lRegistro As String    'Registro.
    Dim lSequencia As Integer    'Seq��ncia.

'1. Inicializa campos.
    lSequencia = 0
    pSequencia = "00"

    If M46V999.gfPreenchido(Dir(pEndLOG)) = True Then
        lNumLog = FreeFile
        On Error Resume Next
        Open pEndLOG For Input Access Read As lNumLog
        If Err <> 0 Then
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        Do While Not EOF(lNumLog)
            Line Input #lNumLog, lRegistro
            lSequencia = val(lRegistro)
            If lSequencia > 0 Then
                Exit Do
            End If
        Loop
        Close #lNumLog
    Else
        lSequencia = 0
    End If

    If lSequencia < pNumMaxSeq Then
        lSequencia = lSequencia + 1
    Else
        lSequencia = 0
    End If

    If pNumMaxSeq > 999 Then
        pSequencia = Format$(lSequencia, "0000")
    ElseIf pNumMaxSeq > 99 Then
        pSequencia = Format$(lSequencia, "000")
    Else
        pSequencia = Format$(lSequencia, "00")
    End If
End Sub

Private Sub mpObtSeq032(ByVal pEndGeisedi2LOG As String, ByRef pSequencia As String)
'Procedure: obt�m �ltima seq��ncia do arquivo de transmiss�o.

'Par�metro de entrada...pEndGeisedi2LOG...Endere�o do arquivo LOG do EDIWare
'                                         (GEISEDI2.LOG ou ERROR.LOG).
'Par�metro de sa�da.....pSequencia........�ltima seq��ncia do arquivo de transmiss�o.

    Dim lNumGeisedi2LOG As Integer  'Utilizado com FreeFile para GEISEDI2.LOG.
    Dim lPosicao As Integer    'Utilizado com InStr.
    Dim lRegistro As String    'Registro.
    Dim lSequencia As Integer    'Seq��ncia.

'1. Inicializa campos.
    lSequencia = 0
    pSequencia = "00"

    If M46V999.gfPreenchido(Dir(pEndGeisedi2LOG)) = True Then
        lNumGeisedi2LOG = FreeFile
        On Error Resume Next
        Open pEndGeisedi2LOG For Input Access Read As lNumGeisedi2LOG
        If Err <> 0 Then
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        Do While Not EOF(lNumGeisedi2LOG)
            Line Input #lNumGeisedi2LOG, lRegistro
            lSequencia = val(lRegistro)
            If lSequencia > 0 Then
                Exit Do
            End If
        Loop
        Close #lNumGeisedi2LOG
    Else
        lSequencia = 0
    End If

    If lSequencia < 99 Then
        lSequencia = lSequencia + 1
    Else
        lSequencia = 0
    End If

    pSequencia = Format$(lSequencia, "00")
End Sub
Private Function mfTraCotacao(ByVal pNumArqTransmissao As Integer, ByVal pNosNumero As String, _
                              pNome As String, pTelefone As String, pEmail As String, _
                              pCotacaoOrig As String, pObservacao As String, pMensagem As String) As Boolean

    mRegistro = "COTACAO "
    If mfPosCampo(6, "nosso n�mero", pNosNumero, 20, 0, pMensagem) = False Then
        Exit Function
    End If

    If mfPosCampo(4, "Nome do solicitante", pNome, 100, 0, pMensagem) = False Then
        Exit Function
    End If

    If mfPosCampo(4, "Email", pEmail, 50, 0, pMensagem) = False Then
        Exit Function
    End If

    If mfPosCampo(4, "Telefone", pTelefone, 50, 0, pMensagem) = False Then
        Exit Function
    End If

    If mfPosCampo(4, "Cotacao Original", pCotacaoOrig, 12, 0, pMensagem) = False Then
        Exit Function
    End If

    If mfPosCampo(6, "Data da solicitacao", Format(Date, "yyyymmdd"), 8, 0, pMensagem) = False Then
        Exit Function
    End If

    If mfPosCampo(4, "Observa��o ", pObservacao, Len(pObservacao), 0, pMensagem) = False Then
        Exit Function
    End If


    mRegistro = UCase$(mRegistro)

    Print #pNumArqTransmissao, mRegistro

    mfTraCotacao = True
End Function
Private Function mfTraCotacaoAceita(ByVal pNumArqTransmissao As Integer, ByVal pNosNumero As String, _
                                    pStatus As Integer, pCotacao As String, _
                                    pSenha As String, pMaxDesc As Double, pMensagem As String) As Boolean

    mRegistro = "COTACAO "
    If mfPosCampo(6, "nosso n�mero", pNosNumero, 20, 0, pMensagem) = False Then
        Exit Function
    End If

    If mfPosCampo(6, "Cotacao", pCotacao, 12, 0, pMensagem) = False Then
        Exit Function
    End If

    If mfPosCampo(6, "Status", pStatus, 2, 0, pMensagem) = False Then
        Exit Function
    End If

    If mfPosCampo(4, "Senha", pSenha, 6, 0, pMensagem) = False Then
        Exit Function
    End If

    If mfPosCampo(6, "Data do retorno", Format(Date, "yyyymmdd"), 8, 0, pMensagem) = False Then
        Exit Function
    End If
    Dim lValor As String
    lValor = pMaxDesc * 100
    If mfPosCampo(6, "Desconto m�ximo", lValor, 4, 0, pMensagem) = False Then
        Exit Function
    End If
    mRegistro = UCase$(mRegistro)

    Print #pNumArqTransmissao, mRegistro

    mfTraCotacaoAceita = True
End Function

Private Sub mpAtualizaSeqLog(ByVal pExtensao As String, _
                             ByVal pSequencia As Integer)

    Dim lEndSeqLog As String
    Dim lNumSeqLog As String
    Dim lMaxSeq As Integer
    Err = 0
    On Error GoTo gpAtualizaSeqLog_ERRO

    Select Case pExtensao
    Case "014"
        lEndSeqLog = M46V999.gAppPath & "SEQ014.LOG"
        lMaxSeq = 9999
    Case "015"
        lEndSeqLog = M46V999.gAppPath & "SEQ015.LOG"
        lMaxSeq = 9999
    Case "010"
        lEndSeqLog = M46V999.gAppPath & "SEQ010.LOG"
        lMaxSeq = 99
    Case "007"
        lEndSeqLog = M46V999.gAppPath & "SEQ007.LOG"
        lMaxSeq = 99
    Case "008"
        lEndSeqLog = M46V999.gAppPath & "SEQ008.LOG"
        lMaxSeq = 99
    Case "021"
        lEndSeqLog = M46V999.gAppPath & "SEQ021.LOG"
        lMaxSeq = 99
    Case "030"
        lEndSeqLog = M46V999.gAppPath & "SEQ030.LOG"
        lMaxSeq = 9999
    Case "032"
        lEndSeqLog = M46V999.gAppPath & "SEQ032.LOG"
        lMaxSeq = 9999
    Case "090"
        lEndSeqLog = M46V999.gAppPath & "SEQ090.LOG"
        lMaxSeq = 999
    Case "091"
        lEndSeqLog = M46V999.gAppPath & "SEQ091.LOG"
        lMaxSeq = 999
    Case "012"
        lEndSeqLog = M46V999.gAppPath & "SEQ012.LOG"
        lMaxSeq = 9999
    Case "046"
        lEndSeqLog = M46V999.gAppPath & "SEQ046.LOG"
        lMaxSeq = 999
    End Select

    lNumSeqLog = FreeFile
    Open lEndSeqLog For Output Access Write As #lNumSeqLog
    If pSequencia >= lMaxSeq Then
        Print #lNumSeqLog, 0
    Else
        Print #lNumSeqLog, pSequencia
    End If
    Close #lNumSeqLog
gpAtualizaSeqLog_SAIDA:
    Exit Sub

gpAtualizaSeqLog_ERRO:
    Call M46V999.gpGraLog(0, "gpAtualizaSeqLog ERRO " & Err & " " & ERROR$)
    On Error Resume Next
    Kill lEndSeqLog         'Exclui o arquivo pode estar danificado
    On Error GoTo 0
    GoTo gpAtualizaSeqLog_SAIDA
End Sub

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

Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
End Sub

Public Function mfCalcularProRata(ByRef pobjstcItem As stcA46V708B, _
                                    ByRef pobjstcPedido As stcA46V702B, _
                                    ByRef pobjstcPedidoApolice As stcA46V702B, _
                                    ByRef pMensagem As String, _
                                    ByRef pPropTabela As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfCalcularProRata"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lObjPed As clsA46V702A
    Dim lobjCobert As clsA46V704A
    Dim lobjstcCobert As stcA46V704B

    Dim lCoePrazo As Single
    Dim ColecaoAux As Collection

    Dim lQtdiasFalta As Integer
    Dim lQtdiasUtilizados As Integer
    Dim lQtdiasApolice As Integer

    Dim lPremio As Double
    Dim lPremioMinimoCobert As Double

    Set lObjPed = New clsA46V702A
    Set lObjPed.M46V999 = M46V999
    
    Set ColecaoAux = New Collection
    Set lobjCobert = New clsA46V704A
    Set lobjCobert.M46V999 = M46V999

    On Error GoTo ControleErro
    mfCalcularProRata = False

    'Calcular
    lPremioMinimoCobert = pobjstcPedido.getProduto.ProPreMinCobert

    'Verificar limite minimo cobertura
    For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert
        If lobjstcCobert.VAL_PR < lPremioMinimoCobert Then
            lobjstcCobert.VAL_PR = lPremioMinimoCobert
            lobjstcCobert.Val_Pr_Tarif = lPremioMinimoCobert
            lobjstcCobert.Val_Pr_Resg = lPremioMinimoCobert

            If Not lobjCobert.mfAlterarValoresPrecificacao(lobjstcCobert, pPropTabela, "") Then
                pMensagem = lobjCobert.UltimoErro
            End If
        End If
        
        'Adiciono a cobertura em uma nova cole��o
        Set lobjstcCobert.ObjPai = pobjstcItem
        ColecaoAux.Add lobjstcCobert, lobjstcCobert.Class_ID

    Next lobjstcCobert
    
    Set pobjstcItem.ColecaoPedCobert = Nothing
    Set pobjstcItem.ColecaoPedCobert = ColecaoAux

    mfCalcularProRata = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    pMensagem = Err & " " & Err.Description
    Call M46V999.gpGraLog(1, "mfCalcularProRata ERRO - " & pMensagem)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
    GoTo SAIDA
FechaObj:
    Set lobjCobert = Nothing
    Set lobjstcCobert = Nothing
    Set ColecaoAux = Nothing
    Set lObjPed = Nothing
    Return
End Function

