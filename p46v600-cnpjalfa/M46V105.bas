Attribute VB_Name = "basM46V105"
Option Explicit
' ======================================================================================================================================================
' Modulo: basM46V105 --> Classe do Form frmT46V101A
' Data/Hora: 01/06/2011
' Sistema: SyasMulti Web
' Simony Sales Andrade
' ======================================================================================================================================================

Public gRst_RegrasPlano As ADODB.Recordset



Sub LixoApagar()
'StringCobertApol = ""
'StringCobertProp = ""
'For Each lobjstcCobert In lObjPedLocApol.ColecaoPedCobert
'    StringCobertApol = StringCobertApol & lobjstcCobert.Cod_Cobert & "@"
'    If lobjstcCobert.Cod_Cobert = 1 Then
'        ValorLMGApol = lobjstcCobert.Val_IS
'    End If
'Next
'For Each lobjstcCobert In lObjPedLoc.ColecaoPedCobert
'    If lobjstcCobert.Cod_Cobert = 1 Then
'       If ValorLMGApol <> lobjstcCobert.Val_IS Then
'           If ValorLMGApol > lobjstcCobert.Val_IS Then
'               StringEndossos = StringEndossos & e_ReducaodeLMG & "@"
'               If TipoEndossoPrincipal <> "RT" Then TipoEndossoPrincipal = "CP"
'            Else
'               StringEndossos = StringEndossos & e_AumentodeLMG & "@"
'               If TipoEndossoPrincipal <> "RT" Then TipoEndossoPrincipal = "CP"
'           End If
'        End If
'   End If
''    StringCobertProp = StringCobertProp & lobjstcCobert.Cod_Cobert & "@"
'Next

'Comparar Coberturas

'lCoberturasApol = Split(StringCobertApol, "@")
'lCoberturasProp = Split(StringCobertProp, "@")

'If UBound(lCoberturasApol) > UBound(lCoberturasProp) Then
'    Tot = UBound(lCoberturasApol)
'Else
'    Tot = UBound(lCoberturasProp)
'End If

'Inclusao = False
'Exclusao = False
'For i = 0 To Tot - 1
'   If i < UBound(lCoberturasApol) And i < UBound(lCoberturasProp) Then
'       If lCoberturasApol(i) <> lCoberturasProp(i) Then
'            If lCoberturasApol(i) > lCoberturasProp(i) Then
'                Inclusao = True
'            Else
'                Exclusao = True
'            End If
'        End If
'    Else
'    End If
'Next
'If Inclusao Then
'    StringEndossos = StringEndossos & e_Inclusaodecobertura & "@"
'    If TipoEndossoPrincipal <> "RT" Then TipoEndossoPrincipal = "CP"
'End If
'If Exclusao Then
'    StringEndossos = StringEndossos & e_Exclusaodecobertura & "@"
'    If TipoEndossoPrincipal <> "RT" Then TipoEndossoPrincipal = "CP"
'End If

End Sub

'### PROJETO INSPECAO
Public Function mfAgendarInspecaoProduto(pobjstcPedido As stcA46V702B, pMensagem As String, Optional pConfirmada As Integer = 0) As Boolean
    Dim objPedLoc As clsA46V708A    'Objeto da classe clsA46V708A (TAB_PED_LOC)
    Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lMensagem As String

    Err.Clear
    mfAgendarInspecaoProduto = False
    On Error GoTo TrataErro

    Set objPedLoc = New clsA46V708A
    Set objPedLoc.M46V999 = gM46V999


    For Each lobjstcPedLoc In pobjstcPedido.ColecaoPedLoc
        If lobjstcPedLoc.Insp_Obr = True Then
            Call gM46V111.gfGravarInspecaoProduto(pobjstcPedido, lobjstcPedLoc, lMensagem, pConfirmada)
        End If
    Next

    mfAgendarInspecaoProduto = True

SAIDA:
    Set objPedLoc = Nothing
    Set lobjstcPedLoc = Nothing

    Exit Function
TrataErro:
    If gM46V999.gfPreenchido(lMensagem) = False Then
        lMensagem = Err & " - " & Error
    End If
    pMensagem = "mfAgendarInspecaoProduto Erro:" & lMensagem
    Call gM46V999.gpGraLog(2, lMensagem)
    GoTo SAIDA
End Function


'Public Function mfCalPreItem(pobjstcItem As stcA46V708B, pobjstcPedido As stcA46V702B, pobjPedido As clsA46V702A, pQtdInsp As Integer, ByRef pNaoPerguntar As Boolean) As Boolean
'    'Função: Calcula prêmios do item.N
'    'Parâmetros de entrada...
'    '                        pAtuStsItem.........Indica se o status do prêmio pode ou não ser atualizado.
'    'Função: calcula prêmios.
'    '********************Variáveis padrão*************************************
'    Dim lCkPoint        As String      'Variavel utilizada para determinar pontos onde o código já passou.
'    Dim lCkStr          As String       'Variavel utilizada para guardar a posição das variaveis até o ultimo
'                                        'lCkPoint. (posição das variaveis em ordem de declaração)
'    '********************Variáveis Declaradas para o processamento************
'    Dim lAguarde            As String           'Mensagem de aguarde.
'    Dim lbdMulti            As ADODB.Connection 'Conexão do Banco local das TAB_PED temporárias
'    Dim lChave              As Long             'Posição na coleção.
'    Dim lErro               As Variant          'Erros da DLL de cálculo de prêmio.
'    Dim lI                  As Byte             'Utilizado com For...Next.
'    Dim lMensagem           As String           'Mensagem.
'    Dim lMsgInsp            As String           'Mensagem de inspeção
'    Dim lPrefixo            As String           'Prefixo de mensagem.
'    Dim lRetorno            As Integer          'Retorno.
'    Dim lrst                As clsYasRecordSet  'Registro.
'    Dim lSelect             As String           'Select
'    Dim lUltData            As Double           'Data da última atualização.
'    Dim lobjstcCobert       As stcA46V704B      'Objeto da classe (stc) de coberturas.
'    Dim lobjstcCobert_Acess As stcA46V704B      'Objeto da classe (stc) de coberturas.
'    Dim lobjCobert          As clsA46V704A      'Objeto da classe de coberturas.
'    Dim lobjPedObs          As clsA46V714A      'Objeto da classe de observação.
'    Dim lObjPedLoc          As clsA46V708A      'Objeto da classe de itens.
'    Dim lBonus              As Long
'    Dim lColecaoTemp        As Collection
'    Dim AtualizaProgres     As Boolean
'    Dim lCod_Aceit          As Integer
'    Dim lDivEndereco        As Boolean
'    Dim lAux                As String
'
'    Dim objstcPedidoApolice  As stcA46V702B  'Classe do Tab_ped que vai carregar os dados do Endosso
'
'    Dim clsA46V101A         As clsA46V101A
'
'    On Error GoTo ControleErro
'
'    'Ficha:900097
'    Dim msg_disp As String
'    Dim msg_01 As String
'    Dim msg_02 As String
'    Dim CobCodNom As String
'    Dim sinal_cob20 As Boolean
'    Dim sinal_cob86 As Boolean
'    Dim sinal_cob64 As Boolean
'    Dim sinal_cob66 As Boolean
'    Dim mat_cobs
'    Dim mat_rub
'    Dim X As Long
'
'    sinal_cob86 = False
'    sinal_cob64 = False
'    sinal_cob66 = False
'    sinal_cob20 = False
'
'    mat_rub = Split("022-13", "-")             '<-- Rubrica a ser considerada: 22-13 Concessionária
'
'    'Ficha 1000268 - permitir a contratação da cobertura 22
'    'mat_cobs = Split("22,23,28,33,41,44", ",") '<-- Coberturas não permitidas
'
'    mat_cobs = Split("23,28,33,41,44", ",") '<-- Coberturas não permitidas
'
'    msg_01 = "Para cobertura XXX é necessário a cobertura YYY ou ZZZ."
'    msg_02 = "As coberturas XXX e YYY não podem ser contratadas simultâneamente."
'
'
'    'Se for a rubrica:
'    If Trim(pobjstcItem.Cod_Clasf) = Val(mat_rub(0)) And Trim(pobjstcItem.Cod_Compl_Clasf) = Val(mat_rub(1)) And gM46V999.gCodProduto >= 956 Then  '1000871 - Inclusão de Orgão Categoria Conselho
'
'       For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert
'         If lobjstcCobert.Cod_Cobert > 1 Then
'
'            'Conferir se a cob 20,64 ou 66 foram contratadas:
'            If lobjstcCobert.Cod_Cobert = 86 Then sinal_cob86 = True
'            If lobjstcCobert.Cod_Cobert = 20 Then sinal_cob20 = True
'            If lobjstcCobert.Cod_Cobert = 64 Then sinal_cob64 = True
'            If lobjstcCobert.Cod_Cobert = 66 Then sinal_cob66 = True
'
'            'Verificar todas as coberturas se:
'            For X = 0 To UBound(mat_cobs)
'                'Estão dentro da matriz de cobs não permitidas;
'                If StrComp(lobjstcCobert.Cod_Cobert, mat_cobs(X)) = 0 Then
'                    'If gM46V999.gTipoAcesso = DIGITACAO And chkLMI.Value = vbChecked Then Exit For
'                    If gM46V999.gTipoAcesso = DIGITACAO And pobjstcItem.Nom_Logr_Risco = "CONFORME ESPECIFICACAO EM ANEXO" Then Exit For
'                    gM46V999.gpMenErro ("A cobertura " & lobjstcCobert.Cod_Cobert & " não se aplica para a rubrica " & mat_rub(0) & "-" & mat_rub(1) & ".")
'                    Exit For
'                End If
'            Next X
'          End If
'        Next
'        'Se cob 86, necessário a 64 ou 66:
'        If sinal_cob86 And Not (sinal_cob64 Or sinal_cob66) Then
'            msg_disp = msg_01
'            gM46V999.gpMenErro (Replace(Replace(Replace(msg_disp, "XXX", "86"), "YYY", "64"), "ZZZ", "66"))
'        End If
'        'Se existem duas que não podem ser contratadas ao mesmo tempo;
'        If sinal_cob20 And sinal_cob64 Then
'            msg_disp = msg_02
'            gM46V999.gpMenErro (Replace(Replace(msg_disp, "XXX", "20"), "YYY", "64"))
'        End If
'        If sinal_cob20 And sinal_cob66 Then
'            msg_disp = msg_02
'            gM46V999.gpMenErro (Replace(Replace(msg_disp, "XXX", "20"), "YYY", "66"))
'        End If
''        If sinal_cob64 And sinal_cob66 Then
''            msg_disp = msg_02
''            gM46V999.gpMenErro (Replace(Replace(msg_disp, "XXX", "64"), "YYY", "66"))
''        End If
'    End If
'
'    'Ficha :0801779
'    Dim IS_valBas As Long
'    Dim str_sql As String
'    Dim cxn As ADODB.Connection
'    Dim obj_rs As clsYasRecordSet
'    Dim msg_erro As String
'
'    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, cxn, lMensagem) = False Then Call gM46V999.gpGraLog(2, lMensagem): Exit Function
'
'    For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert
'        If lobjstcCobert.Cod_Cobert = 1 Then
'            IS_valBas = lobjstcCobert.Val_IS
'        End If
'        If lobjstcCobert.Cod_Cobert > 1 Then
'
'            str_sql = "SELECT COD_LIM_COBERT,PER_LIM_MAX_COBERT FROM TAB_COBERT " & _
             '                      "WHERE COD_RAMO = " & gM46V999.gCod_Ramo & _
             '                      " AND COD_TARIF_INI <= " & gM46V999.gCodProduto & _
             '                      " AND COD_TARIF_TER >= " & gM46V999.gCodProduto & _
             '                      " AND TAB_COBERT = " & gM46V111.gCod_Plano & _
             '                      " AND COD_COBERT = " & lobjstcCobert.Cod_Cobert
'
'            If gM46V999.gfObtRegistro(cxn, str_sql, obj_rs, lMensagem) = False Then Call gM46V999.gpGraLog(1, lMensagem): Exit Function
'
'            If Not (obj_rs.BOF And obj_rs.EOF) Then
'                For Each lobjstcCobert_Acess In pobjstcItem.ColecaoPedCobert
'                    If Trim(Val(lobjstcCobert_Acess.Cod_Cobert)) = Trim(Val(obj_rs("COD_LIM_COBERT"))) And Val(lobjstcCobert_Acess.Cod_Cobert) > 1 Then
'                        If lobjstcCobert.Val_IS > Val(lobjstcCobert_Acess.Val_IS) Then
'                            msg_erro = " Consistência do Item " & CStr(lobjstcCobert.Num_Item) & " - COBERTURA " & CStr(lobjstcCobert.Cod_Cobert) & " COM IS SUPERIOR AO DA COBERTURA " & lobjstcCobert_Acess.Cod_Cobert & "."
'                            Call gM46V999.gpMenErro(msg_erro)
'                        End If
'                        Exit For
'                    End If
'                Next
'            End If
'        End If
'
'        '1001146 - Tarifa 1000 e 1001 Dezembro 2010
'        If gM46V999.gCodProduto > 999 Then
'            If (gM46V999.gCod_Ramo = e_CodRamo_Empresarial And lobjstcCobert.Cod_Cobert = 94) Or _
             '               (gM46V999.gCod_Ramo = e_CodRamo_Residencial And (lobjstcCobert.Cod_Cobert = 27 Or _
             '                                                      lobjstcCobert.Cod_Cobert = 29 Or _
             '                                                      lobjstcCobert.Cod_Cobert = 30 Or _
             '                                                      lobjstcCobert.Cod_Cobert = 31)) Then
'                pobjstcItem.Cob_Bens = True
'                pobjstcItem.Dig_Bens = (pobjstcItem.ColecaoPedLocBem.Count > 0)
'                'Comparar soma dos Bens com o LMG da cobertura 27 ou 29 ou 30 ou 31 ou 94
'                If pobjstcItem.Dig_Bens = True Then
'                    Call mfConsTotBens(Format(lobjstcCobert.Val_IS, "standard"), lobjstcCobert.Cod_Cobert, pobjstcItem)
'                End If
'            End If
'        End If
'
'    Next
'
'    'Se selecionada cobs. 58 ou 62 E classif. "Concessionária", necessaria cobert. 57.
''    'NOVA ADIÇÃO: Quando concessionária, impedir cobertura 22.
'
''Ficha 901419
''A cobertura 62 não está relacionada com a cobertura 57
''Ao contratar a 58 tem que contratar a 57 ou 124 ou 125
'
'    Dim lObjstcCobert_aux As stcA46V704B
'    Dim BOL_sinal_cob57 As Boolean
'    Dim BOL_sinal_cob124 As Boolean
'    Dim BOL_sinal_cob125 As Boolean
'    gM46V999.GLB_sinal_concess = False
'
'    If Val(pobjstcItem.Cod_Clasf) = 22 And Val(pobjstcItem.Cod_Compl_Clasf) = 13 And gM46V999.gCodProduto >= 936 Then
'        gM46V999.GLB_sinal_concess = True
'
'        For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert
'
'    'Ficha 1000268 - Quando concessionária, permitir cobertura 22 no Syas interno
'            If lobjstcCobert.Cod_Cobert = 22 Then
'               If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
'                    If MsgBox("Cobertura 22 - RC OPERAÇÕES não se aplica para a classificação 022-13 (Concessionária)." & vbCrLf & "Deseja desconsiderar esta consistência?", vbQuestion + vbYesNo) = vbNo Then
'                        GoTo LBL_GRAVAR
'                    End If
'                Else
'                    Call gM46V999.gpMenErro(" Consistência do Item " & CStr(lobjstcCobert.Num_Item) & " -PARA ESTA CLASSIFICAÇÃO, A COBERTURA 00" & CStr(lobjstcCobert.Cod_Cobert) & " NÃO É VÁLIDA.")
'               End If
'            End If
'
'            If lobjstcCobert.Cod_Cobert = 58 Then
'                BOL_sinal_cob57 = False
'                BOL_sinal_cob124 = False
'                BOL_sinal_cob125 = False
'
'                For Each lObjstcCobert_aux In pobjstcItem.ColecaoPedCobert
'                 If lObjstcCobert_aux.Cod_Cobert = 57 Then BOL_sinal_cob57 = True: Exit For
'                 If lObjstcCobert_aux.Cod_Cobert = 124 Then BOL_sinal_cob124 = True: Exit For
'                 If lObjstcCobert_aux.Cod_Cobert = 125 Then BOL_sinal_cob125 = True: Exit For
'                Next
'
'                If Not BOL_sinal_cob57 And Not BOL_sinal_cob124 And Not BOL_sinal_cob125 Then
'                 Call gM46V999.gpMenErro(" Consistência do Item " & CStr(lobjstcCobert.Num_Item) & " - PARA COBERTURA 00" & CStr(lobjstcCobert.Cod_Cobert) & ", FALTA A COBERTURA 0057 OU 0124 OU 0125.")
'                End If
'            End If
'
'        GoTo LBL_FIM:
'LBL_GRAVAR:
'        Call gM46V999.gpMenErro(" Consistência do Item " & CStr(lobjstcCobert.Num_Item) & " -PARA ESTA CLASSIFICAÇÃO, A COBERTURA 00" & CStr(lobjstcCobert.Cod_Cobert) & " NÃO É VÁLIDA.")
'LBL_FIM:
'      Next
'    End If
'
'
'    'Comparar soma dos Bens com o LMG - Syas Roubo para plano AllRisks
'    If gM46V999.gCod_Ramo = e_CodRamo_Roubos Then
'
'        'Ficha 1000308 - Permitir itens com planos diferentes no produto Roubo (cálculo coletivo)
'        Call mfConsTotBens(Format(pobjstcItem.Val_Vr, "standard"), 0, pobjstcItem)
'        If pobjstcItem.ColecaoPedLocBem.Count = 0 Then
'
'            If pobjstcItem.Cod_Plano = 2 Then
'
'                For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert
'                    If lobjstcCobert.Cod_Cobert = 5 Then
'                        'If Val(pobjstcItem.Num_Item) = Val(txtNumItem.Text) Then
'                        '    Screen.MousePointer = vbDefault
'                        '    Call gM46V999.gpGraLog(1, "Obrigatório incluir a relação de objetos com seus respectivos valores unitários.")
'                        '    frmT46V129A.Show vbModal
'                        '    Exit Function
'                        'Else
'                            Call gM46V999.gpMenErro(" Consistência do Item " & CStr(lobjstcCobert.Num_Item) & " Obrigatório incluir a relação de objetos com seus respectivos valores unitários.")
'                            Exit For
'                        'End If
'                    End If
'                Next lobjstcCobert
'            End If
'        End If
'
'
'    End If
'
'    '0800635
'    If pobjstcItem.Num_Cep_Risco = pobjstcPedido.Num_CEP And _
     '        pNaoPerguntar = False And pobjstcPedido.Cod_Tip_Doc = e_TipDoc_Proposta And _
     '        (gM46V999.PerfilAcesso = e_Perfil_CALC Or gM46V999.PerfilAcesso = e_Perfil_DIG_IST Or _
     '        gM46V999.PerfilAcesso = e_Perfil_DIG_Yas) Then
'        'Verifico se os demais campos são iguais
'        lMensagem = "Dados do(s) campo(s) (<Campo>) divergente entre Proponente e Local em Risco."
'        lDivEndereco = False
'        If pobjstcItem.Cod_Tip_Logr_Risco <> pobjstcPedido.Cod_Tip_Logr Then
'            lAux = "Tipo de logradouro;"
'            lDivEndereco = True
'        End If
'        If pobjstcItem.Nom_Logr_Risco <> pobjstcPedido.Nom_Logr Then
'            lAux = lAux & "Logradouro;"
'            lDivEndereco = True
'        End If
'        If pobjstcItem.Nom_Bairro_Risco <> pobjstcPedido.Nom_Bairro Then
'            lAux = lAux & "Bairro;"
'            lDivEndereco = True
'        End If
'        If Trim(Replace(pobjstcItem.objstcPedInsp.Num_Logr_Risco & " " & pobjstcItem.Dsc_Compl_Risco, ".", "")) <> Trim(pobjstcPedido.Dsc_Compl_Logr) Then
'            lAux = lAux & "Complemento;"
'            lDivEndereco = True
'        End If
'        If pobjstcItem.Sig_UF_Risco <> pobjstcPedido.Sig_UF Then
'            lAux = lAux & "UF;"
'            lDivEndereco = True
'        End If
'        If lDivEndereco = True Then
'            lMensagem = Replace(lMensagem, "<Campo>", lAux)
'            Call gM46V999.gpGraLog(0, lMensagem)
'            If MsgBox("Item " & pobjstcItem.Num_Item & " -" & lMensagem & " Deseja continuar desta forma?", vbQuestion + vbYesNo) = vbNo Then
'                Call gM46V999.gpMenErro(lMensagem)
'            Else
'                pNaoPerguntar = True
'            End If
'        End If
'    End If
'
'    mfCalPreItem = False
'    'Calcula
'    'a) Prepara controles de mensagens.
'    lAguarde = "Cálculo de prêmios. Item " & pobjstcItem.Num_Item & ". " & vbLf
'    lPrefixo = "Item " & pobjstcItem.Num_Item & " - "
'
'
'    'b) Calcula prêmios e salva registro com resultado do cálculo.
'    'panAguarde.Caption = lAguarde & "Efetuando cálculo de prêmios do item..."
'    'panAguarde.Refresh
'
'    Set lObjPedLoc = New clsA46V708A
'    Set lObjPedLoc.M46V999 = gM46V999
'
'    Set lObjPedLoc.ConexaoBD = pobjPedido.ConexaoBD
'
'
'    ''Resseguro facultativo
'    'Call gM46V999.gpInicializarIdentificacaoResseguro(objstcPedido)
'
'    Call mpConsisteItem(pobjstcItem.Num_Item, pobjstcPedido, pobjstcItem, lRetorno)
'    lCkPoint = "8"
'    'c) Consiste item - nível 2.
'
'    'SyasWeb
'    'panAguarde.Caption = lAguarde & "Consistindo dados do item após o cálculo... "
'    'panAguarde.Refresh
'    Call mpVerIteNiv_2(pobjstcItem, pobjstcPedido)
'    lCkPoint = "9"
'    'd) Consiste item - normas
'    'panAguarde.Caption = lAguarde & "Consistindo as normas..."
'    'panAguarde.Refresh
'
'    lCkPoint = "10"
'    'Verifica Regras de Aceitação Pontuais
'
'    If gM46V999.VerificarNormas Then
'        lCod_Aceit = 0
'        lCod_Aceit = gM46V103.mfClasseAceit(pobjstcItem, pobjstcPedido)
'        If lCod_Aceit = 3 Then
'            gM46V999.gpMenErro ("FGE0156 - Proposta sem aceitação. ")
'        ElseIf lCod_Aceit = 4 Then
'            gM46V999.gpMenErro ("FGE0157 - Para essa proposta solicitar cotação a Unidade de Produtos Diversos.")
'        End If
'        lCkPoint = "11"
'    End If
'
'    Call mpVerNormas(pobjstcItem.Num_Item, lRetorno, lCod_Aceit, pobjstcPedido, pobjstcItem, lMsgInsp)
'
'    If lMsgInsp = "" And pobjstcItem.Ind_Sin_Inspec > 0 Then
'         lMsgInsp = "OBRIGATORIO A REALIZACAO DE INSPECAO DO RISCO"
'    End If
'
'PulaNorma:
'
''############## PROJETO 10018 - INSPECAO ####################
'    pobjstcItem.Insp_Obr = False
'
'    If gM46V999.gfPreenchido(lMsgInsp) = True Then     'INSPEÇÃO OBRIGATÓRIA
'        pobjstcItem.Insp_Obr = True
'        If pobjstcPedido.Cod_Tip_Doc = e_TipDoc_Proposta Then
'            If gM46V999.gfPreenchido(pobjstcItem.objstcPedInsp.Num_Tel) = False Then
'                gM46V999.gpMenErro ("FGE0105 - Telefone de contato para inspeção inválido (não preenchido).")
'            Else
'                '1000477 - Telefone de Contato Formatado
'                '1 - valida se esta totalmente preenchido
'                If Len(pobjstcItem.objstcPedInsp.Num_Tel) < 13 Then
'                    gM46V999.gpMenErro ("FGE0105 - Telefone de contato para inspeção inválido (Favor informar o DDD e o número do telefone com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
'                Else
'                    '2 - valida formado (99)9999-9999
'                    If Mid(pobjstcItem.objstcPedInsp.Num_Tel, 1, 1) <> "(" Or Mid(pobjstcItem.objstcPedInsp.Num_Tel, 4, 1) <> ")" Or Mid(pobjstcItem.objstcPedInsp.Num_Tel, 9, 1) <> "-" Then
'                        gM46V999.gpMenErro ("FGE0105 - Telefone de contato para inspeção inválido (Favor informar o DDD e o número do telefone com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
'                    Else
'                        '3 - não permite todos os numeros do telefone iguais
'                        If InStr(1, "0000-0000 1111-1111 2222-2222 3333-3333 4444-4444 5555-5555 6666-6666 7777-7777 8888-8888 9999-9999", Mid(pobjstcItem.objstcPedInsp.Num_Tel, 5, Len(pobjstcItem.objstcPedInsp.Num_Tel)), vbTextCompare) > 0 Then
'                            gM46V999.gpMenErro ("FGE0105 - Telefone de contato para inspeção inválido (Favor informar o DDD e o número do telefone com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
'                        End If
'                    End If
'                End If
'            End If
'
'            If gM46V999.gfPreenchido(pobjstcItem.objstcPedInsp.Pess_Cont) = False Or Len(pobjstcItem.objstcPedInsp.Pess_Cont) <= 3 Then
'                gM46V999.gpMenErro ("FGE0106 - Contato para inspeção inválido (não preenchido).")
'            End If
'
'            If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
'                If pQtdInsp = 1 Then
'                    Call gM46V999.gpGraLog(1, "A aceitação deste seguro está condicionada a análise da inspeção do risco.")
'                End If
'            Else
'                Call gM46V999.gpGraLog(1, "A aceitação deste seguro está condicionada a análise da inspeção do risco - Item:" & pobjstcItem.Num_Item & ".")
'            End If
'            pobjstcItem.Insp_Agend = True
'            Call gM46V999.gpGraLog(0, lMensagem & " SIM.")
'
'        ElseIf pobjstcPedido.Cod_Tip_Doc = e_TipDoc_Cotacao Then
'                pobjstcItem.Insp_Agend = False
'                Call gM46V999.gpGraLog(0, lMensagem & " NÃO.")
'        End If
'
'        lRetorno = 0
'
'    End If
'
'    If pobjstcItem.ColecaoPedCobert.Count = 0 Then
'        gM46V999.gpMenErro ("AGE0047 - Nenhuma cobertura foi selecionada.")
'    End If
'
'    'Se houve erro nos tipos de produto selecionados, encerra processo.
'    If gM46V999.gQuaErros <> 0 Or lRetorno <> 0 Then
'        Exit Function
'    End If
'
'    If Not gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdMulti, lMensagem) Then
'        Call gM46V999.gpGraLog(1, lMensagem)
'        Exit Function
'    End If
'
'    AtualizaProgres = True
''    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Or _
 ''       (gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And _
 ''       (txtDescProgr.Enabled = False Or pobjstcItem.Per_Desc_Progres = 99)) Then
'    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Or _
     '       (gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And _
     '       (pobjstcItem.Per_Desc_Progres = 99)) Then
'        lSelect = "Select * from TAB_PED_LOC"
'        lSelect = lSelect & " where num_item = " & pobjstcItem.Num_Item
'        lSelect = lSelect & " and num_apol = 0"
'
'        'Busca os prêmios de coberturas
'        If Not gM46V999.gfObtRegistro(lbdMulti, lSelect, lrst, lMensagem) Then
'            Call gM46V999.gpGraLog(1, lMensagem)
'            Exit Function
'        End If
'        gCarga = True
'
'        'Ficha 901673 e 901772
'        'If cboPlano.ItemData(cboPlano.ListIndex) < 2000 Or _
         '        '   (cboPlano.ItemData(cboPlano.ListIndex) >= 2021 And cboPlano.ItemData(cboPlano.ListIndex) <= 2031) Or _
         '        '   (cboPlano.ItemData(cboPlano.ListIndex) >= 2061 And cboPlano.ItemData(cboPlano.ListIndex) <= 2100) Or _
         '        '   (cboPlano.ItemData(cboPlano.ListIndex) >= 2111 And cboPlano.ItemData(cboPlano.ListIndex) <= 2130) Then
'
'        If pobjstcPedido.Cod_Produto < 2000 Or _
         '           (pobjstcPedido.Cod_Produto >= 2021 And pobjstcPedido.Cod_Produto <= 2031) Or _
         '           (pobjstcPedido.Cod_Produto >= 2061 And pobjstcPedido.Cod_Produto <= 2100) Or _
         '           (pobjstcPedido.Cod_Produto >= 2111 And pobjstcPedido.Cod_Produto <= 2130) Then
'             pobjstcItem.Per_Desc_Progres = lrst("Per_Desc_Progres")
'        End If
'
'        gCarga = False
'
'        If Not lObjPedLoc.mfAlterar(pobjstcItem, "", "") Then
'            lMensagem = lObjPedLoc.UltimoErro
'            Call gM46V999.gpGraLog(1, lMensagem)
'            Exit Function
'        End If
'    Else
'        AtualizaProgres = False
'    End If
'
'    If gM46V999.gCodProduto <= gM46V999.TarifaAntiga Then
'        lSelect = "Select * from Tab_Ped_Cobert"
'        lSelect = lSelect & " where Tab_Ped_Cobert.num_item = " & pobjstcItem.Num_Item
'        lSelect = lSelect & " and Tab_Ped_Cobert.num_apol = 0"
'        lSelect = lSelect & " and Tab_Ped_Cobert.Val_IS <> 0"
'
'        'Busca os prêmios de coberturas
'        If Not gM46V999.gfObtRegistro(lbdMulti, lSelect, lrst, lMensagem) Then
'            Call gM46V999.gpGraLog(1, lMensagem)
'            Exit Function
'        End If
'
'TentadeNovo:
'        Set pobjstcItem.ColecaoPedCobert = New Collection
'        Set lobjstcCobert = New stcA46V704B
'        Set lobjstcCobert.M46V999 = gM46V999
'
'        Set lobjCobert = New clsA46V704A
'        Set lobjCobert.M46V999 = gM46V999
'
'        'Set lObjCobert.ConexaoBD = objPedido.ConexaoBD
'        lobjstcCobert.NOSSO_NUMERO = pobjstcPedido.NOSSO_NUMERO
'        lobjstcCobert.Num_Item = pobjstcItem.Num_Item
'        If Not lobjCobert.mfExcluir(lobjstcCobert, "", "", True) Then
'            Call gM46V999.gpGraLog(1, lobjCobert.UltimoErro)
'            Exit Function
'        End If
'        Do While Not lrst.EOF
'            Set lobjstcCobert = New stcA46V704B
'            Set lobjstcCobert.M46V999 = gM46V999
'
'            lobjstcCobert.Cod_Cobert = lrst("Cod_Cobert")
'            lobjstcCobert.Cod_Org = lrst("Cod_Org")
'            lobjstcCobert.Cod_Ramo = pobjstcPedido.Cod_Ramo
'            lobjstcCobert.NOSSO_NUMERO = pobjstcPedido.NOSSO_NUMERO
'            lobjstcCobert.Num_Item = pobjstcItem.Num_Item
'            lobjstcCobert.Num_Ped = pobjstcItem.Num_Ped
'            lobjstcCobert.Per_Desc_Coml = lrst("Per_Desc_Coml")
'            lobjstcCobert.Per_Desc_Exp = lrst("Per_Desc_Exp")
'            lobjstcCobert.Per_Franq = lrst("Per_Franq")
'            lobjstcCobert.Per_Taxa = lrst("Per_Taxa")
'            lobjstcCobert.Tab_Cobert = lrst("Tab_Cobert")
'            lobjstcCobert.Val_Franq_Max = lrst("Val_Franq_Max")
'            lobjstcCobert.Val_Franq_Min = lrst("Val_Franq_Min")
'            lobjstcCobert.Val_IS = lrst("Val_IS")
'            lobjstcCobert.VAL_PR = lrst("Val_Pr")
'            lobjstcCobert.Val_Pr_Anu = lrst("Val_Pr_Anu")
'            lobjstcCobert.Val_Pr_Resg = lrst("Val_Pr_Resg")
'            lobjstcCobert.Val_Pr_Resg_Anu = lrst("Val_Pr_Resg_Anu")
'            lobjstcCobert.Val_Pr_Tarif = lrst("Val_Pr_Tarif")
'            lobjstcCobert.Val_Pr_Tarif_Anu = lrst("Val_Pr_Tarif_Anu")
'            If Not lobjCobert.mfIncluir(lobjstcCobert, "", "", True) Then
'                Call gM46V999.gpGraLog(1, lobjCobert.UltimoErro)
'                Exit Function
'            End If
'            pobjstcItem.ColecaoPedCobert.Add lobjstcCobert, lobjstcCobert.Class_ID
'            lrst.MoveNext
'        Loop
'
'    End If
'    lCkPoint = "50"
'    'Antes de Carregar objeto vai calcular novamente
'    If gM46V999.gCodProduto > gM46V999.TarifaAntiga Then
'        If gM46V111.gTipoGrupMulti = e_GrupoRamo_Multi Then
'           If Not gM46V103.mfCalNewPrecif(pobjstcItem, AtualizaProgres, lMensagem, "", "") Then
'                Call gM46V999.gpMenErro(lMensagem)
'                Exit Function
'            End If
'        End If
'    End If
'
'    'SimonyEndo
'    Set clsA46V101A = New clsA46V101A
'    Set clsA46V101A.M46V999 = gM46V999
'
'    If objstcPedido.Cod_Dizer <> 0 Then
'        Call clsA46V101A.mfCarregarDadosOriginais(objstcPedido, objstcPedidoApolice)
'        If Not clsA46V101A.mfComparareCalcular(pobjstcItem, objstcPedido, objstcPedidoApolice, lMensagem) Then
'            Call gM46V999.gpMenErro(lMensagem)
'            Exit Function
'        End If
'
'    End If
'    Set clsA46V101A = Nothing
'    gCarga = True
'
'    'If cboPlano.ItemData(cboPlano.ListIndex) < 2000 Then
'    If pobjstcPedido.Cod_Produto < 2000 Then
'    'VER VER VER
'        'txtDescProgr.Text = pobjstcItem.Per_Desc_Progres
'    End If
'
'    gCarga = False
'
'
'    'Salva resultado do cálculo.
'    'a) Atualiza registro do item.
'    lUltData = CDbl(gM46V999.gfForData(3))
''Ficha 0901834 Correção de Status (Para Emissão Mainframe não pode ir 2)
''    pobjstcItem.Cod_Sit_item = e_SitItem_Calculado
'
'    'VER VER VER
'    'If pobjstcItem.Num_Item = Val(txtNumItem.Text) Then
'    '    gIteUltUsuario = gM46V999.gNomUsuario6
'    '    gIteUltData = lUltData
'    '    Call mpCarregarCobertGrid(pobjstcItem.ColecaoPedCobert)
'    'End If
'
'    pobjstcPedido.ColecaoPedLoc.Remove pobjstcItem.Class_ID
'
'    'Verifico se deve ser gravada a observação de inspeção.
'    If gM46V999.gfPreenchido(lMsgInsp) = True Then
'        If InStr(1, gM46V999.gfRetirarCaracterEspecial(pobjstcItem.objstcPedObs.Observacao, True), gM46V999.gfRetirarCaracterEspecial(lMsgInsp, True)) = 0 Then
'            If gM46V999.gfPreenchido(pobjstcItem.objstcPedObs.Observacao) Then
'                If gM46V999.gNovaPropostaLMI Then
'                    pobjstcItem.objstcPedObs.Observacao = pobjstcItem.objstcPedObs.Observacao & vbCrLf & "Item " & Format(pobjstcItem.Num_Item, "00000") & ": " & gM46V999.gfRetirarCaracterEspecial(lMsgInsp, True)
'                Else
'                    pobjstcItem.objstcPedObs.Observacao = gM46V999.gfRetirarCaracterEspecial(lMsgInsp, True) & vbCrLf & pobjstcItem.objstcPedObs.Observacao
'                End If
'
'            Else
'                If gM46V999.gNovaPropostaLMI Then
'                    pobjstcItem.objstcPedObs.Observacao = "Item " & Format(pobjstcItem.Num_Item, "00000") & ": " & gM46V999.gfRetirarCaracterEspecial(lMsgInsp, True)
'                Else
'                    pobjstcItem.objstcPedObs.Observacao = gM46V999.gfRetirarCaracterEspecial(lMsgInsp, True)
'                End If
'            End If
'            pobjstcItem.objstcPedObs.Observacao = Replace(pobjstcItem.objstcPedObs.Observacao, vbCrLf, "&%")
'            Set lobjPedObs = New clsA46V714A
'            Set lobjPedObs.M46V999 = gM46V999
'
'            Set lobjPedObs.ConexaoBD = pobjPedido.ConexaoBD
'            If Not lobjPedObs.mfExcluir(pobjstcItem.objstcPedObs, "", "") Then
'                lMensagem = lobjPedObs.UltimoErro
'                Call gM46V999.gpMenErro(lMensagem)
'                Exit Function
'            End If
'            With pobjstcItem.objstcPedObs
'                .NOSSO_NUMERO = pobjstcItem.NOSSO_NUMERO
'                .Num_Item = pobjstcItem.Num_Item
'            End With
'            If Not lobjPedObs.mfIncluir(pobjstcItem.objstcPedObs, "", "") Then
'                lMensagem = lobjPedObs.UltimoErro
'                Call gM46V999.gpMenErro(lMensagem)
'                Exit Function
'            End If
'        End If
'    End If
'
'    '############## PROJETO 10018 - INSPECAO ####################
'    If gM46V999.gfPreenchido(lMsgInsp) = True Then
'        If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
'            'Não vai agM46V999.gEndar inspeção qdo for CLASSE 4, o cálculo do Produto já vai ter gravado o agM46V999.gEndamento
'            'O corretor irá acessar pelo YasudaNet e confirmar o agM46V999.gEndamento
'            'If objstcPedido.Ind_CotacaoSolic = True And _
             '            '    objstcPedido.StatusCotSiscota = enum_StatusTransm_Aprovada Then
'            '        GoTo NaoAgendar
'            'End If
'            If gpGeraLogCalcInspecao(pobjstcPedido, pobjstcItem) = False Then
'                If pobjstcItem.Insp_Agend = True Then
'                    Call gM46V999.gpGraLog(1, "Não gerou arquivo de transmissão da inspeção " & pobjstcPedido.NOSSO_NUMERO)
'                End If
'            End If
'        End If
'    End If
'
'NaoAgendar:
'
'    'Verifico em que lugar da coleção será adicionado.
'    'É feito isso para que não seja exibido com os número fora de ordem
'    lChave = gpPosicaoColecao(pobjstcItem.Num_Item)
'    If pobjstcPedido.ColecaoPedLoc.Count > 0 And pobjstcPedido.ColecaoPedLoc.Count <> lChave - 1 Then
'        pobjstcPedido.ColecaoPedLoc.Add pobjstcItem, pobjstcItem.Class_ID, lChave
'    Else
'        pobjstcPedido.ColecaoPedLoc.Add pobjstcItem, pobjstcItem.Class_ID
'    End If
'
'    mfCalPreItem = True
'    Exit Function
'ControleErro:
'    Call gM46V999.gpGraLog(1, Err.Description)
'End Function




Public Sub mpVerNormas(pItem As Long, pCod_Retorno As Integer, pCod_Aceit As Integer, pobjstcPedido As stcA46V702B, pobjstcPedLoc As stcA46V708B, ByRef pMensInsp As String)
    Dim lretorno As String
    Dim lI As Integer
    Dim lCod As Integer
    Dim lAuxiliar As String
    Dim lAuxiliar2 As String
    Dim lBDP00Multi As ADODB.Connection
    Dim lRstErro As clsYasRecordSet
    Dim lMensagem As String
    Dim QtdErr As Integer
    Dim lListaErros As String
    Dim lSelect As String
    Dim lUpdate As String

    pMensInsp = ""
    lListaErros = ""
    QtdErr = 0
    If gM46V999.VerificarNormas = False And gM46V999.gTipoAcesso = DIGITACAO Then

        If gM46V999.PerfilAcesso = e_Perfil_CALC_Gestor And gM46V999.gfPreenchido(pobjstcPedido.objstcCtrlEmis.Num_Cotac) Then
            'Continua
        Else
            Exit Sub
        End If
    End If

    lretorno = gM46V111.gfExecDLL(pobjstcPedido, pobjstcPedLoc, e_TipProc_Normas, pItem, pCod_Retorno, gM46V999.gTipoHomologacao)

    If InStr(1, lretorno, "Existem erros de consistência!") <> 0 Then
        'Busca os erros de consistência.
        If gM46V999.VerificarNormas Then
            Call mpCapturaErros("Normas ")
        End If

    ElseIf gM46V999.gfPreenchido(lretorno) = True Then    'retornou com erros
        lI = InStr(1, lretorno, "--") + 2
        If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
        Do
            lCod = val(Mid(lretorno, lI, 3))
            'Se na tabela de parametrização retornou classe 1 despreza retornos 10 e 80
            If pCod_Aceit = 1 Then
                If lCod <> 55 Then
                    lI = InStr(lI, lretorno, "-") + 1
                    GoTo LerProximoErro
                End If
            End If
            lI = InStr(lI, lretorno, "-") + 1
            lAuxiliar = Mid(lretorno, lI, 11)

            If gM46V999.VerificarNormas = False Then
                'Gravo o erro da norma
                lListaErros = lListaErros & ";" & val(lAuxiliar)
            End If

            lSelect = " Select * from Tab_Dsc_Retorno  where Codigo = " & val(lAuxiliar)

            If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lRstErro, lMensagem) = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
                Exit Sub
            End If

            If Not lRstErro.EOF Then
                ' CR3923 - Item 1 - Início
                Dim mensagemInterna As String
                mensagemInterna = ""
                If Not lRstErro("descricao_interna") Is Nothing Then
                    mensagemInterna = Trim(lRstErro("descricao_interna"))
                End If

                If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And mensagemInterna <> "" Then
                    lAuxiliar = val(lAuxiliar) & " - " & Trim(lRstErro("descricao_interna"))
                Else
                    lAuxiliar = val(lAuxiliar) & " - " & Trim(lRstErro("Descricao"))
                End If
                ' CR3923 - Item 1 - Fim
            End If

            If gM46V999.gCod_Ramo = e_CodRamo_Residencial And gM46V999.gCodProduto >= 948 Then
                If pobjstcPedido.objstcCtrlEmis.Cod_Corr = 469999 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 680299 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 462513 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 340709 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 363409 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 683709 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 685309 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 455121 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 490799 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 653799 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 473721 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 654021 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 234621 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 653621 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 269821 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 663299 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 72221 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 37121 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 309801 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 1601 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 428601 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 18301 Or _
                   pobjstcPedido.objstcCtrlEmis.Cod_Corr = 383901 Then
                    lAuxiliar2 = Replace(lAuxiliar, "Sem aceitação", "Solicitar cotação")
                    lAuxiliar = lAuxiliar2
                ElseIf pobjstcPedido.objstcCtrlEmis.Cod_Corr = 425201 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 448199 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 33601 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 398301 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 611501 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 680301 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 208901 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 687201 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 32201 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 675801 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 684899 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 687101 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 26301 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 688901 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 678401 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 250501 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 382101 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 21401 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 393501 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 692201 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 691101 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 684903 Then
                    lAuxiliar2 = Replace(lAuxiliar, "Sem aceitação", "Solicitar cotação")
                    lAuxiliar = lAuxiliar2
                ElseIf pobjstcPedido.objstcCtrlEmis.Cod_Corr = 475599 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 466699 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 675199 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 689999 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 349605 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 689105 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 327499 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 364699 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 34946 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 660105 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 259605 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 675305 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 455405 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 328905 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 328999 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 333176 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 270205 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 661505 Or _
                       pobjstcPedido.objstcCtrlEmis.Cod_Corr = 17905 Then
                    lAuxiliar2 = Replace(lAuxiliar, "Sem aceitação", "Solicitar cotação")
                    lAuxiliar = lAuxiliar2
                End If
            End If

            If gM46V111.gCod_plano > 2000 And lAuxiliar Like "123341 -*" Or _
               gM46V111.gCod_plano = 2041 And lAuxiliar Like "1234413 -*" Then
                GoTo LerProximoErro
            End If
            If lCod = 55 Then
                pMensInsp = Trim(Mid(lAuxiliar, InStr(1, lAuxiliar, "-") + 1))
            Else
                'Ficha 801779
                'Ficha 901346
                If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And (gM46V999.gCod_Ramo = e_CodRamo_Residencial Or gM46V999.gCod_Ramo = e_CodRamo_Empresarial) And _
                   Trim(UCase(lAuxiliar)) Like UCase("*LMG*Superior ao limite*Sem aceitação*") Then
                    If gM46V999.gDesconsiderou <> 1 Then
                        If MsgBox(UCase(lAuxiliar) & vbCrLf & "Deseja desconsiderar esta consistência?", vbQuestion + vbYesNo) = vbYes Then
                            Call gM46V999.gpDesconsiderou
                            GoTo LerProximoErro
                        End If
                    ElseIf gM46V999.gDesconsiderou = 1 Then
                        GoTo LerProximoErro
                    End If
                End If
                If gM46V999.VerificarNormas Then
                    Call gM46V999.gpMenErro("Normas " & lAuxiliar)
                    QtdErr = QtdErr + 1
                End If
            End If
LerProximoErro:
            lI = InStr(lI, lretorno, "-") + 1
            If lI = 1 Then Exit Do
        Loop
        Call gM46V999.gpFecha2(lBDP00Multi)
    End If
    If QtdErr = 0 Then
        pCod_Retorno = 0
    End If
    DoEvents

    '###PROJETO SISCOTA
    'DESCOMENTAR LINHAS
    If (gM46V999.gfPreenchido(lListaErros)) And (gM46V999.PerfilAcesso = e_Perfil_CALC_Gestor) And (gM46V999.gBaseLocal = False) Then
        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
        lUpdate = " Update siscota.dbo.tab_cotacao_div set normas = '" & lListaErros & "' where ccotacao = '" & pobjstcPedido.objstcCtrlEmis.Num_Cotac & "'"
        If gM46V999.gfExeSQL(lBDP00Multi, lUpdate, lMensagem) <> 0 Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
    End If

End Sub

Private Sub mpVerIteNiv_2(pobjstcItem As stcA46V708B, pobjstcPedido As stcA46V702B)
'Procedure: consiste item nível 2: consiste todos dados do item.

    Dim lobjA46V103A As clsA46V103A
    Dim objTabUser As clsA46V725A
    Dim objstcTabUser As stcA46V725B

    Set lobjA46V103A = New clsA46V103A
    Set lobjA46V103A.M46V103 = gM46V103
    
    Set objTabUser = New clsA46V725A
    Set objTabUser.M46V999 = gM46V999

    Set objstcTabUser = New stcA46V725B
    objstcTabUser.login_user_rede = UCase(gM46V999.gNomUsuario)
    If Not objTabUser.mfCarregarColecao(objstcTabUser) Then
        Call gM46V999.gpGraLog(1, objTabUser.UltimoErro)
        Exit Sub
    End If

    'VER VER VER
    'Call lobjA46V103A.gpVerIteNiv_2(pobjstcItem, pobjstcPedido, Val(txtISBas.Text), panLocRisco.Enabled)
    Call lobjA46V103A.gpVerIteNiv_2(pobjstcItem, pobjstcPedido, pobjstcPedido.Val_IS, True, objstcTabUser, "")

End Sub

Private Sub mpConsisteItem(pItem As Long, pobjstcPedido As stcA46V702B, pobjstcPedLoc As stcA46V708B, ByRef pCod_Retorno As Integer)
    Dim lretorno As String
    Dim lbdP00MPED As ADODB.Connection    'Base de dados: P0042300.
    Dim lRstErro As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String

    'Não consiste Itens dos ramos diferente de Multi
    If gM46V111.gTipoGrupMulti <> e_GrupoRamo_Multi Then Exit Sub

    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, lbdP00MPED, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    lSelect = " select count(*) as cobert from Tab_Ped_Cobert where nosso_numero = '" & pobjstcPedido.NOSSO_NUMERO & "' And Num_Item = " & pItem
    If gM46V999.gfObtRegistro(lbdP00MPED, lSelect, lRstErro, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    If Not lRstErro.EOF Then
        'Call gM46V999.gpGraLog(2, "Antes Consistir *** Tab_Cobert ==> " & lRstErro("cobert") & " *** ITEM " & pItem)
    End If

    lretorno = gM46V111.gfExecDLL(pobjstcPedido, pobjstcPedLoc, e_TipProc_ConsItem, pItem, pCod_Retorno, gM46V999.gTipoHomologacao)

    lSelect = " select count(*) as cobert from Tab_Ped_Cobert where nosso_numero = '" & pobjstcPedido.NOSSO_NUMERO & "' And Num_Item = " & pItem
    If gM46V999.gfObtRegistro(lbdP00MPED, lSelect, lRstErro, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    If Not lRstErro.EOF Then
        'Call gM46V999.gpGraLog(2, "Depois de Consistir *** Tab_Cobert ==> " & lRstErro("cobert") & " *** ITEM " & pItem)
    End If


    If InStr(1, lretorno, "Existem erros de consistência!") <> 0 Then
        'Busca os erros de consistência.
        Call mpCapturaErros("Consistência do item " & pItem)
        pCod_Retorno = 0
    ElseIf gM46V999.gfPreenchido(lretorno) = True Then    'retornou sem erros
        Call gM46V999.gpMenErro("Item " & Format(pItem, "00000") & "  " & lretorno)
    End If

    DoEvents
End Sub

Public Sub mpCapturaErros(pPref As String)
    Dim lBDP00Multi As ADODB.Connection
    Dim lRstErro As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String
    Dim lDelete As String
    
    'Procedimento necessario apenas para normas COBOL
    If gM46V999.gCodProduto >= 1080 Then Exit Sub

    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    lSelect = " Select * from Tab_Erro"

    If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lRstErro, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    Do While Not lRstErro.EOF
        'Se o erro for de dependência de cobertura e for usuário de emissão então questionar
        'se libera o erro (temporário)
        If Trim(lRstErro("DSC_ERRO")) = "PARA COBERTURA 0032, FALTA A COBERTURA 0039" And gM46V999.gCodProduto >= 938 Then
            GoTo LBL_FIM
        End If

        If Trim(lRstErro("DSC_ERRO")) = "PARA COBERTURA 0024, A IS NAO PODE SER SUPERIOR A  20,00% DA COB. 0001" And gM46V999.gCodProduto >= 950 And gM46V999.gCod_Corretor = 15701 Then
            GoTo LBL_FIM
        End If

        If gM46V999.PerfilAcesso = e_Perfil_CALC_Gestor And gM46V111.gNum_Pi = 0 Then GoTo LBL_GRAVAR
        If gM46V999.PerfilAcesso <> e_Perfil_DIG_Gestor And gM46V999.PerfilAcesso <> e_Perfil_CALC_Gestor Then GoTo LBL_GRAVAR

        'gM46V999.GLB_sinal_concess é uma global que é true se as propr. [.Cod_Clasf] e [.Cod_Compl_Clasf] do Item (pobjstcitem-mfcalpreitem) são relativos a "022-13-Concessionária".
        If Not gM46V999.GLB_sinal_concess And (Trim(lRstErro("DSC_ERRO")) Like "COBERTURA * VALIDA SOMENTE PARA A CLASSIF. *" Or _
                                               Trim(lRstErro("DSC_ERRO")) Like "PARA COBERTURA * FALTA A COBERTURA *") Then
            If gM46V999.gDesconsiderou <> 1 Then
                If MsgBox(lRstErro("DSC_ERRO") & vbCrLf & "Deseja desconsiderar esta consistência?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    GoTo LBL_GRAVAR
                Else
                    Call gM46V999.gpDesconsiderou
                End If
            End If
        Else
            'Ficha 801779
            'Ficha 901346
            If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And (gM46V999.gCod_Ramo = e_CodRamo_Residencial Or gM46V999.gCod_Ramo = e_CodRamo_Empresarial) And _
               Trim(UCase(lRstErro("DSC_ERRO"))) Like "*PARA COBERTURA*A IS NAO PODE SER SUPERIOR A*" Then
                If gM46V999.gDesconsiderou <> 1 Then
                    If MsgBox(lRstErro("DSC_ERRO") & vbCrLf & "Deseja desconsiderar esta consistência?", vbQuestion + vbYesNo) = vbNo Then
                        GoTo LBL_GRAVAR
                    Else
                        Call gM46V999.gpDesconsiderou
                    End If
                End If
            Else
                GoTo LBL_GRAVAR
            End If
        End If

        GoTo LBL_FIM:
LBL_GRAVAR:
        Call gM46V999.gpMenErro(pPref & " - " & Trim(lRstErro("DSC_ERRO")))
LBL_FIM:
        lRstErro.MoveNext
    Loop

    lDelete = " delete from Tab_Erro"
    If gM46V999.gfExeSQL(lBDP00Multi, lDelete, lMensagem) <> 0 Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    Call gM46V999.gpFecha2(lBDP00Multi)
End Sub


Private Function mfConsTotBens(ByVal pValLMG As Double, ByVal pCodCobert As Long, pobjstcLoc As stcA46V708B)
    Dim pValBens As Double
    Dim lI As Integer    'Utilizado no for...loop.
    Dim lObjstcBens As stcA46V715B

    'Verificar se a soma dos Bens é igual ao LMG - para plano All Risks
    pValBens = 0

    If gM46V999.gCodProduto >= 947 And gM46V999.gCod_Ramo = e_CodRamo_Roubos Then


        For Each lObjstcBens In pobjstcLoc.ColecaoPedLocBem
            pValBens = pValBens + lObjstcBens.Val_Bem
        Next lObjstcBens

        'Ficha 1000308 - Permitir itens com planos diferentes no produto Roubo (cálculo coletivo)
        If pobjstcLoc.Cod_Plano = 1 Then
            If pValBens <> pValLMG Then
                Screen.MousePointer = vbDefault
                Call gM46V999.gpMenErro("FNP0074 - " & gM46V999.gPrefixo & "Somatória dos bens deve ser igual ao LMG.")
                Exit Function
            End If
        Else
            If pValBens > pValLMG Then
                Screen.MousePointer = vbDefault
                Call gM46V999.gpMenErro("FNP0074 - " & gM46V999.gPrefixo & "Somatória dos bens não deve ser maior que o LMG total.")
                Exit Function
            End If
        End If

    End If

    If gM46V999.gCodProduto >= 999 And (gM46V999.gCod_Ramo = e_CodRamo_Empresarial Or gM46V999.gCod_Ramo = e_CodRamo_Residencial) And gM46V111.gCod_plano = 1 Then

        For Each lObjstcBens In pobjstcLoc.ColecaoPedLocBem
            pValBens = pValBens + lObjstcBens.Val_Bem
        Next lObjstcBens

        If pValBens <> pValLMG Then
            Screen.MousePointer = vbDefault
            Call gM46V999.gpMenErro("FNP0074 - " & gM46V999.gPrefixo & "O somatório dos valores da relação de bens deve ser igual ao LMG da cobertura " & pCodCobert & ".")
            Exit Function
        End If

    End If

    mfConsTotBens = True
    Exit Function

ControleErro:
    mfConsTotBens = False

End Function


