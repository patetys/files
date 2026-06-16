VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V708A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V708A --> Tabela Tab_Ped_Loc
' Data/Hora: 31/08/2005 - 11:54:41
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================

Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V708B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V708A"
End Property

Public Function mfAlterarTodosItemNaoCalc(ByRef pobjstc As stcA46V708B, _
                                          ByVal pPropTabela As String, _
                                          ByVal pNomeTabImport As String) As Boolean

    Dim lUpdate As String

    On Error GoTo ControleErros

    If Not mfConectar Then GoTo ControleErros

    lUpdate = " Update " & pPropTabela & pNomeTabImport & "Tab_Ped_loc set Cod_Sit_item  = 1"
    lUpdate = lUpdate & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErros

    mfAlterarTodosItemNaoCalc = True

    Exit Function
ControleErros:
    mfAlterarTodosItemNaoCalc = False
    UltimoErro = "A46V708A.mfAlterarTodosItemNaoCalc ERRO - " & Err & " " & Err.Description & _
                 vbCrLf & " SQL: " & lUpdate
End Function

Public Function mfAlterarValorRisco(ByRef pobjstc As stcA46V708B, _
                                    ByVal pPropTabela As String, _
                                    ByVal pNomeTabImport As String) As Boolean

    Dim lUpdate As String

    On Error GoTo ControleErros

    If Not mfConectar Then GoTo ControleErros

    lUpdate = " Update " & pPropTabela & pNomeTabImport & "Tab_Ped_loc set val_vr = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Vr)
    lUpdate = lUpdate & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    lUpdate = lUpdate & " and  NUM_ITEM = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)

    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErros

    mfAlterarValorRisco = True

    Exit Function
ControleErros:
    mfAlterarValorRisco = False
    UltimoErro = "A46V708A.mfAlterarValorRisco  ERRO - " & Err & " " & Err.Description & _
                 vbCrLf & " SQL: " & lUpdate
End Function

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V708B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       Optional ByVal CodTab As Integer) As String

    Dim lInsert As String                'Query.

    If IsNumeric(pobjstc.nosso_numero) And pobjstc.ColecaoPedCobert.Count > 0 Then ' não seta valor quando for processo criação de registro padrão
        pobjstc.Val_LMG_Lc = pobjstc.gISTotal_LC
    End If

    If CodTab = 2 Then GoTo Tabela2

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc ( "
    lInsert = lInsert & " Nosso_Numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " Num_item,"
    lInsert = lInsert & " Num_Ped,"
    lInsert = lInsert & " Cod_Org,"
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " Num_Apol,"
    lInsert = lInsert & " Num_Endo,"
    lInsert = lInsert & " Cod_Tip_Logr_Risco,"
    lInsert = lInsert & " Nom_Logr_Risco,"
    lInsert = lInsert & " Dsc_Compl_Risco,"
    lInsert = lInsert & " Nom_Bairro_Risco,"
    lInsert = lInsert & " Nom_Cid_Risco,"
    lInsert = lInsert & " Sig_UF_Risco,"
    lInsert = lInsert & " Num_CEP_Risco,"
    lInsert = lInsert & " Per_Desc_Apar,"
    lInsert = lInsert & " Per_Desc_Progres,"
    lInsert = lInsert & " Per_IS_Vr,"
    lInsert = lInsert & " Val_Coef_IS_Vr,"
    lInsert = lInsert & " Val_Vr,"
    lInsert = lInsert & " Tab_Clasf,"
    lInsert = lInsert & " Cod_Clasf,"
    lInsert = lInsert & " Cod_Compl_Clasf,"
    lInsert = lInsert & " Dsc_Clasf,"
    lInsert = lInsert & " Cod_Classe_Ocup,"
    lInsert = lInsert & " Cod_Constr,"
    lInsert = lInsert & " Cod_User_Digit,"
    lInsert = lInsert & " Dat_Digit,"
    lInsert = lInsert & " Hor_Digit,"
    lInsert = lInsert & " Cod_User_Liber,"
    lInsert = lInsert & " Dat_Liber,"
    lInsert = lInsert & " Hor_Liber,"
    lInsert = lInsert & " Cod_Sit_item,"
    lInsert = lInsert & " Per_Bonus,"
    lInsert = lInsert & " Qtd_Empreg,"
    lInsert = lInsert & " Num_INSP,"
    lInsert = lInsert & " Cod_Iden_Segur_01,"
    lInsert = lInsert & " Cod_Iden_Segur_02,"
    lInsert = lInsert & " Cod_Iden_Segur_03,"
    lInsert = lInsert & " Cod_Iden_Segur_04,"
    lInsert = lInsert & " Cod_Iden_Segur_05,"
    lInsert = lInsert & " Cod_Iden_Segur_06,"
    lInsert = lInsert & " Cod_Iden_Segur_07,"
    lInsert = lInsert & " Cod_Iden_Segur_08,"
    lInsert = lInsert & " Cod_Iden_Segur_09,"
    lInsert = lInsert & " Cod_Iden_Segur_10,"
    lInsert = lInsert & " Per_Tot_Segur,"
    lInsert = lInsert & " Cod_Iden_Protec_01,"
    lInsert = lInsert & " Cod_Iden_Protec_02,"
    lInsert = lInsert & " Cod_Iden_Protec_03,"
    lInsert = lInsert & " Cod_Iden_Protec_04,"
    lInsert = lInsert & " Cod_Iden_Protec_05,"
    lInsert = lInsert & " Cod_Iden_Protec_06,"
    lInsert = lInsert & " Cod_Iden_Protec_07,"
    lInsert = lInsert & " Cod_Iden_Protec_08,"
    lInsert = lInsert & " Cod_Iden_Protec_09,"
    lInsert = lInsert & " Cod_Iden_Protec_10,"
    lInsert = lInsert & " Per_Tot_Protec,"
    lInsert = lInsert & " Cod_Produto,"
    lInsert = lInsert & " Cod_Regiao,"
    lInsert = lInsert & " Cod_Iden_Tarif,"
    lInsert = lInsert & " Num_PI,"
    lInsert = lInsert & " Tip_Emissao,"
    lInsert = lInsert & " Num_Apol_Anterior,"
    lInsert = lInsert & " Num_Item_Anterior,"

    '1000347 - Inclusão de informações da congênere em renovações
    lInsert = lInsert & " Ren_Con_Seguradora,"
    lInsert = lInsert & " Ren_Con_Apolice,"
    lInsert = lInsert & " Ren_Con_Item,"

    lInsert = lInsert & " Ind_Sinistro"
    '    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
    lInsert = lInsert & ",Insp_Obr"
    lInsert = lInsert & ",Insp_Agend"
    lInsert = lInsert & ",Dig_Bens"
    lInsert = lInsert & ",Cob_Bens"
    '    End If

    'Ficha 1000305
    lInsert = lInsert & ",Tip_Cobert"

    'Ficha 1200134
    lInsert = lInsert & ",Ind_Sin_Inspec"
    lInsert = lInsert & ",Cod_DizerItem "
       
    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Logr_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Logr_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Bairro_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Cid_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Sig_UF_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cep_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Apar) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Progres) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_IS_Vr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Coef_IS_Vr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Vr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tab_Clasf) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Clasf) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Compl_Clasf) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Clasf) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Classe_Ocup) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Constr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Digit) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.DAT_DIGIT) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.HOR_DIGIT) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Liber) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Liber) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Hor_Liber) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Sit_item) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Bonus) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Empreg) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Insp) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_01_V) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_02_A) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_03_P) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_04_I) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_05_G) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_06) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_07) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_08) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_09) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_10) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Tot_Segur) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_01_E) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_02_H) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_03_S) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_04_D) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_05) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_06) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_07) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_08) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_09) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_10) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Tot_Protec) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Produto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Regiao) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Tarif) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Emissao) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol_Anterior) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item_Anterior) & ","

    '1000347 - Inclusão de informações da congênere em renovações
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ren_Con_Seguradora) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ren_Con_Apolice) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ren_Con_Item) & ","

    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Sinistro)
    '    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
    lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Insp_Obr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Insp_Agend)
    lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Dig_Bens)
    lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Cob_Bens)
    '    End If

    'Ficha 1000305
    lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Cobert)

    'Ficha 1200134
    lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Sin_Inspec)
    lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Cod_DizerItem)
    
    lInsert = lInsert & ")"

    GoTo Fim

Tabela2:

    'ProjetoFormularios

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc_Compl ( "
    lInsert = lInsert & " Nosso_Numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " Num_item,"
    lInsert = lInsert & " Num_Ped,"
    lInsert = lInsert & " Num_PI,"
    lInsert = lInsert & " Cod_Org,"
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " Num_Apol,"
    lInsert = lInsert & " Num_Endo, "
    lInsert = lInsert & " Ind_VerbaSep, "
    lInsert = lInsert & " Val_Vr_Lc, "
    lInsert = lInsert & " Val_LMG_Lc, "
    lInsert = lInsert & " Per_IS_Vr_Lc, "
    lInsert = lInsert & " Coef_IS_Vr_Lc, "
    lInsert = lInsert & " Tip_Plano_Vida, "
    lInsert = lInsert & " Cod_Plano_Cobertura, "
    lInsert = lInsert & " Qtd_Vidas, "
    lInsert = lInsert & " Val_LMG_Vida, "
    lInsert = lInsert & " Ind_Aux_Funeral, "
    lInsert = lInsert & " Val_LMG_Funeral, "
    lInsert = lInsert & " Ind_Inc_Conjuge, "
    lInsert = lInsert & " Val_Aux_Cesta, "
    lInsert = lInsert & " Possui_Elevador, "
    lInsert = lInsert & " BonusAlterado, "
    lInsert = lInsert & " Flg_EnviarCertificado, "
    lInsert = lInsert & " Txt_Observacoes, "
    lInsert = lInsert & " Ind_StatusInspecao "

    If UCase(pPropTabela) <> "GED.DBO." And pNomeTabImport = "" Then
        lInsert = lInsert & " ,Val_Vr_Orig "
        lInsert = lInsert & " ,Val_Bas_Orig "
        lInsert = lInsert & " ,Val_Vr_Lc_Orig "
        lInsert = lInsert & " ,ITEM_ORIGEM_C4 "
    End If

    '8000042754 - Resposta dada a questão parametrizada para o plano.
    lInsert = lInsert & " ,CLAUSULANAODEPRECIACAO " 'JIRA-25945 - Opção "Sem Depreciação" no Condomínio
    lInsert = lInsert & ", COD_RESPOSTA "
    
'    *******************************************************************************************************************
'    *#20180829
'    *Finalidade: Apura Desconto Comercial Congenere
'     Inclusão do Campo c.DescComlCongenereIncorporado na Select
'    *******************************************************************************************************************
    lInsert = lInsert & ",DescComlCongenereIncorporado "

    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","

    'ProjetoFormularios
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ind_VerbaSep) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Vr_Lc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_LMG_Lc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_IS_Vr_Lc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Coef_IS_Vr_Lc) & ","

    'Capital Global
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Plano_Vida) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Plano_Cobertura) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Vidas) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_LMG_Vida) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Aux_Funeral) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_LMG_Funeral) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Inc_Conjuge) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Aux_Cesta) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Possui_Elevador) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.BonusAlterado) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Flg_EnviarCertificado) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(M46V999.gfRetirarCaracterEspecial(pobjstc.Txt_Observacoes)) & "," 'Luiz Alberto - 26/10/2015
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ind_StatusInspecao)
    
    If UCase(pPropTabela) <> "GED.DBO." And pNomeTabImport = "" Then
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Val_Vr_Orig)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Val_Bas_Orig)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Val_Vr_Lc_Orig)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.ITEM_ORIGEM_C4)
    End If
    
    '8000042754 - Resposta dada a questão parametrizada para o plano.
    lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.ClausulaNaoDepreciacao) 'JIRA-25945 - Opção "Sem Depreciação" no Condomínio
    lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.RespostaPergunta)
    
'    *******************************************************************************************************************
'    *#20180829
'    *Finalidade: Apura Desconto Comercial Congenere
'     Inclusão do Campo c.DescComlCongenereIncorporado na Select
'    *******************************************************************************************************************
    lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.DescComlCongenereIncorporado)
    
    lInsert = lInsert & ")"

Fim:

    mfFormatarInclusaoSQL = lInsert

End Function


Private Function mfFormatarAlteracaoSQL(ByRef pobjstc As stcA46V708B, _
                                        ByVal pPropTabela As String, _
                                        ByVal pNomeTabImport As String, _
                                        Optional ByVal CodTab As Integer) As String
    Dim lUpdate As String                'Query.

    pobjstc.Val_LMG_Lc = pobjstc.gISTotal_LC

    If CodTab = 2 Then GoTo Tabela2

    lUpdate = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc Set "
    lUpdate = lUpdate & " Num_item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lUpdate = lUpdate & " Num_Ped = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lUpdate = lUpdate & " Cod_Org = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lUpdate = lUpdate & " Cod_Ramo = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lUpdate = lUpdate & " Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lUpdate = lUpdate & " Num_Endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lUpdate = lUpdate & " Cod_Tip_Logr_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Logr_Risco) & ","
    lUpdate = lUpdate & " Nom_Logr_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Logr_Risco) & ","
    lUpdate = lUpdate & " Dsc_Compl_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Risco) & ","
    lUpdate = lUpdate & " Nom_Bairro_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Bairro_Risco) & ","
    lUpdate = lUpdate & " Nom_Cid_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Cid_Risco) & ","
    lUpdate = lUpdate & " Sig_UF_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Sig_UF_Risco) & ","
    lUpdate = lUpdate & " Num_CEP_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cep_Risco) & ","
    lUpdate = lUpdate & " Per_Desc_Apar = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Apar) & ","
    lUpdate = lUpdate & " Per_Desc_Progres = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Progres) & ","
    lUpdate = lUpdate & " Per_IS_Vr = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_IS_Vr) & ","
    lUpdate = lUpdate & " Val_Coef_IS_Vr = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Coef_IS_Vr) & ","
    lUpdate = lUpdate & " Val_Vr = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Vr) & ","
    lUpdate = lUpdate & " Tab_Clasf = " & M46V999.gfFormatarCamposSQL(pobjstc.Tab_Clasf) & ","
    lUpdate = lUpdate & " Cod_Clasf = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Clasf) & ","
    lUpdate = lUpdate & " Cod_Compl_Clasf = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Compl_Clasf) & ","
    lUpdate = lUpdate & " Dsc_Clasf = " & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Clasf) & ","
    lUpdate = lUpdate & " Cod_Classe_Ocup = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Classe_Ocup) & ","
    lUpdate = lUpdate & " Cod_Constr = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Constr) & ","
    lUpdate = lUpdate & " Cod_User_Digit = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Digit) & ","
    lUpdate = lUpdate & " Dat_Digit = " & M46V999.gfFormatarCamposSQL(pobjstc.DAT_DIGIT) & ","
    lUpdate = lUpdate & " Hor_Digit = " & M46V999.gfFormatarCamposSQL(pobjstc.HOR_DIGIT) & ","
    lUpdate = lUpdate & " Cod_User_Liber = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Liber) & ","
    lUpdate = lUpdate & " Dat_Liber = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Liber) & ","
    lUpdate = lUpdate & " Hor_Liber = " & M46V999.gfFormatarCamposSQL(pobjstc.Hor_Liber) & ","
    lUpdate = lUpdate & " Cod_Sit_item = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Sit_item) & ","
    lUpdate = lUpdate & " Per_Bonus = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Bonus) & ","
    lUpdate = lUpdate & " Qtd_Empreg = " & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Empreg) & ","
    lUpdate = lUpdate & " Num_INSP = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Insp) & ","
    lUpdate = lUpdate & " Cod_Iden_Segur_01 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_01_V) & ","
    lUpdate = lUpdate & " Cod_Iden_Segur_02 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_02_A) & ","
    lUpdate = lUpdate & " Cod_Iden_Segur_03 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_03_P) & ","
    lUpdate = lUpdate & " Cod_Iden_Segur_04 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_04_I) & ","
    lUpdate = lUpdate & " Cod_Iden_Segur_05 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_05_G) & ","
    lUpdate = lUpdate & " Cod_Iden_Segur_06 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_06) & ","
    lUpdate = lUpdate & " Cod_Iden_Segur_07 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_07) & ","
    lUpdate = lUpdate & " Cod_Iden_Segur_08 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_08) & ","
    lUpdate = lUpdate & " Cod_Iden_Segur_09 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_09) & ","
    lUpdate = lUpdate & " Cod_Iden_Segur_10 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Segur_10) & ","
    lUpdate = lUpdate & " Per_Tot_Segur = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Tot_Segur) & ","
    lUpdate = lUpdate & " Cod_Iden_Protec_01 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_01_E) & ","
    lUpdate = lUpdate & " Cod_Iden_Protec_02 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_02_H) & ","
    lUpdate = lUpdate & " Cod_Iden_Protec_03 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_03_S) & ","
    lUpdate = lUpdate & " Cod_Iden_Protec_04 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_04_D) & ","
    lUpdate = lUpdate & " Cod_Iden_Protec_05 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_05) & ","
    lUpdate = lUpdate & " Cod_Iden_Protec_06 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_06) & ","
    lUpdate = lUpdate & " Cod_Iden_Protec_07 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_07) & ","
    lUpdate = lUpdate & " Cod_Iden_Protec_08 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_08) & ","
    lUpdate = lUpdate & " Cod_Iden_Protec_09 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_09) & ","
    lUpdate = lUpdate & " Cod_Iden_Protec_10 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Protec_10) & ","
    lUpdate = lUpdate & " Per_Tot_Protec = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Tot_Protec) & ","
    lUpdate = lUpdate & " Cod_Produto = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Produto) & ","
    lUpdate = lUpdate & " Cod_Regiao = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Regiao) & ","
    lUpdate = lUpdate & " Cod_Iden_Tarif = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Tarif) & ","
    lUpdate = lUpdate & " Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lUpdate = lUpdate & " Tip_Emissao = " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Emissao) & ","
    lUpdate = lUpdate & " Num_Apol_Anterior = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol_Anterior) & ","
    lUpdate = lUpdate & " Num_Item_Anterior = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item_Anterior) & ","

    '1000347 - Inclusão de informações da congênere em renovações
    lUpdate = lUpdate & " Ren_Con_Seguradora = " & M46V999.gfFormatarCamposSQL(pobjstc.Ren_Con_Seguradora) & ","
    lUpdate = lUpdate & " Ren_Con_Apolice = " & M46V999.gfFormatarCamposSQL(pobjstc.Ren_Con_Apolice) & ","
    lUpdate = lUpdate & " Ren_Con_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Ren_Con_Item) & ","

    lUpdate = lUpdate & " Ind_Sinistro = " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Sinistro)
    '    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
    lUpdate = lUpdate & ", Insp_Obr = " & M46V999.gfFormatarCamposSQL(pobjstc.Insp_Obr) & ","
    lUpdate = lUpdate & " Insp_Agend = " & M46V999.gfFormatarCamposSQL(pobjstc.Insp_Agend)
    lUpdate = lUpdate & ", Dig_Bens = " & M46V999.gfFormatarCamposSQL(pobjstc.Dig_Bens) & ","
    lUpdate = lUpdate & " Cob_Bens = " & M46V999.gfFormatarCamposSQL(pobjstc.Cob_Bens)
    '    End If

    'Ficha 1000305
    lUpdate = lUpdate & ", Tip_Cobert = " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Cobert)

    'Ficha 1200134
    lUpdate = lUpdate & ", Ind_Sin_Inspec = " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Sin_Inspec)
    lUpdate = lUpdate & ", Cod_DizerItem = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_DizerItem)
    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    ' Estava gerando duplicidade quando o item a ser alterado tinha origem em uma copia
    '
    'If (pobjstc.Num_Item_Ant <> 0 And pobjstc.Num_Item <> pobjstc.Num_Item_Ant) Then
    '    'Se a alteração envolve tbm em alterar o número do item...
    '    lUpdate = lUpdate & " and Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item_Ant)
    'Else
        lUpdate = lUpdate & " and Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    'End If

    GoTo Fim

Tabela2:
    pobjstc.Val_LMG_Lc = pobjstc.gISTotal_LC

    lUpdate = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc_Compl  Set "
    lUpdate = lUpdate & " Num_item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lUpdate = lUpdate & " Num_Ped = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lUpdate = lUpdate & " Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lUpdate = lUpdate & " Cod_Org = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lUpdate = lUpdate & " Cod_Ramo = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lUpdate = lUpdate & " Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lUpdate = lUpdate & " Num_Endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","

    'ProjetoFormularios
    lUpdate = lUpdate & " Ind_VerbaSep = " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_VerbaSep) & ","
    lUpdate = lUpdate & " Val_Vr_Lc = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Vr_Lc) & ","
    lUpdate = lUpdate & " Val_LMG_Lc = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_LMG_Lc) & ","
    lUpdate = lUpdate & " Per_IS_Vr_Lc = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_IS_Vr_Lc) & ","
    lUpdate = lUpdate & " Coef_IS_Vr_Lc = " & M46V999.gfFormatarCamposSQL(pobjstc.Coef_IS_Vr_Lc) & ","

    'Capital Global.
    lUpdate = lUpdate & " Tip_Plano_Vida = " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Plano_Vida) & ","
    lUpdate = lUpdate & " Cod_Plano_Cobertura = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Plano_Cobertura) & ","
    lUpdate = lUpdate & " Qtd_Vidas = " & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Vidas) & ","
    lUpdate = lUpdate & " Val_LMG_Vida = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_LMG_Vida) & ","
    lUpdate = lUpdate & " Ind_Aux_Funeral = " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Aux_Funeral) & ","
    lUpdate = lUpdate & " Val_LMG_Funeral = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_LMG_Funeral) & ","
    lUpdate = lUpdate & " Ind_Inc_Conjuge = " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Inc_Conjuge) & ","
    lUpdate = lUpdate & " Val_Aux_Cesta = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Aux_Cesta) & ","
    lUpdate = lUpdate & " Possui_Elevador = " & M46V999.gfFormatarCamposSQL(pobjstc.Possui_Elevador) & ","
    lUpdate = lUpdate & " BonusAlterado = " & M46V999.gfFormatarCamposSQL(pobjstc.BonusAlterado) & ","
    lUpdate = lUpdate & " Flg_EnviarCertificado = " & M46V999.gfFormatarCamposSQL(pobjstc.Flg_EnviarCertificado) & ","
    lUpdate = lUpdate & " Txt_Observacoes = " & M46V999.gfFormatarCamposSQL(pobjstc.Txt_Observacoes) & "," 'Luiz Alberto - 26/10/2015
    lUpdate = lUpdate & " Ind_StatusInspecao = " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_StatusInspecao)

    If UCase(pPropTabela) <> "GED.DBO." And pNomeTabImport = "" Then
        lUpdate = lUpdate & " ,Val_Vr_Orig = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Vr_Orig)
        lUpdate = lUpdate & " ,Val_Bas_Orig = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Bas_Orig)
        lUpdate = lUpdate & " ,Val_Vr_Lc_Orig = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Vr_Lc_Orig)
        lUpdate = lUpdate & " ,ITEM_ORIGEM_C4 = " & M46V999.gfFormatarCamposSQL(pobjstc.ITEM_ORIGEM_C4)
        lUpdate = lUpdate & " ,CLAUSULANAODEPRECIACAO = " & M46V999.gfFormatarCamposSQL(pobjstc.ClausulaNaoDepreciacao) 'JIRA-25945 - Opção "Sem Depreciação" no Condomínio
        lUpdate = lUpdate & " ,DescComlCongenereIncorporado = " & M46V999.gfFormatarCamposSQL(pobjstc.DescComlCongenereIncorporado) '*Finalidade: Apura Desconto Comercial Congenere
    End If

    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    ' Estava gerando duplicidade quando o item a ser alterado tinha origem em uma copia
    '
    'If (pobjstc.Num_Item_Ant <> 0 And pobjstc.Num_Item <> pobjstc.Num_Item_Ant) Then
    '    lUpdate = lUpdate & " and Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item_Ant)
    'Else
        lUpdate = lUpdate & " and Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    'End If

Fim:

    mfFormatarAlteracaoSQL = lUpdate

End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V708B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       Optional ByVal CodTab As Integer) As String
    Dim lDelete As String                'Query.

    If CodTab = 2 Then GoTo Tabela2

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc WHERE "
    lDelete = lDelete & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Item > 0 Then
        lDelete = lDelete & " AND Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If

    GoTo Fim
Tabela2:

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc_Compl WHERE "
    lDelete = lDelete & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Item > 0 Then
        lDelete = lDelete & " AND Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If




Fim:
    mfFormatarExclusaoSQL = lDelete
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V708B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String                'Query.

    lSelect = "SELECT Loc.* "

    'ProjetoFormularios
    lSelect = lSelect & ", Ind_VerbaSep, Val_Vr_Lc, Val_LMG_Lc, Per_IS_Vr_Lc, Coef_IS_Vr_Lc "
    lSelect = lSelect & ",  Tip_Plano_Vida"
    lSelect = lSelect & ", Cod_Plano_Cobertura"
    lSelect = lSelect & ", Qtd_Vidas"
    lSelect = lSelect & ", Val_LMG_Vida"
    lSelect = lSelect & ", Ind_Aux_Funeral"
    lSelect = lSelect & ", Val_LMG_Funeral"
    lSelect = lSelect & ", Ind_Inc_Conjuge"
    lSelect = lSelect & ", Val_Aux_Cesta"
    lSelect = lSelect & ", Possui_Elevador"
    lSelect = lSelect & ", BonusAlterado "
    lSelect = lSelect & ", Flg_EnviarCertificado "
    lSelect = lSelect & ", Txt_Observacoes" 'Luiz Alberto - 26/10/2015
    lSelect = lSelect & ", Ind_StatusInspecao "
'    *******************************************************************************************************************
'    *#20180829
'    *Finalidade: Apura Desconto Comercial Congenere
'     Inclusão do Campo c.DescComlCongenereIncorporado na Select
'    *******************************************************************************************************************
    lSelect = lSelect & ", Compl.DescComlCongenereIncorporado "

    If UCase(pPropTabela) <> "GED.DBO." And pNomeTabImport = "" Then
        lSelect = lSelect & " ,Val_Vr_Orig "
        lSelect = lSelect & " ,Val_Bas_Orig "
        lSelect = lSelect & " ,Val_Vr_Lc_Orig "
        lSelect = lSelect & " ,ITEM_ORIGEM_C4 "
        lSelect = lSelect & " ,Compl.CLAUSULANAODEPRECIACAO " 'JIRA-25945 - Opção "Sem Depreciação" no Condomínio
    End If

    '8000042754 - Resposta dada a questão parametrizada para o plano.
    lSelect = lSelect & " ,COD_RESPOSTA "

    lSelect = lSelect & " FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc as Loc  with(nolock) "

    'ProjetoFormularios
    lSelect = lSelect & " Left join " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc_Compl as Compl "
    
    'lSelect = lSelect & " on  Loc.NOSSO_NUMERO = Compl.NOSSO_NUMERO  "
    If pPropTabela = "GED.DBO." Then
        lSelect = lSelect & " on  Loc.Num_PI = Compl.Num_PI  "
    Else
        lSelect = lSelect & " on  Loc.NOSSO_NUMERO = Compl.NOSSO_NUMERO  "
    End If
    
    lSelect = lSelect & " and Loc.Num_Item = Compl.Num_Item  "


    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lSelect = lSelect & " Where Loc.Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If
    If pobjstc.Num_Item <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Loc.Num_Item= " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Loc.Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Apol <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Loc.Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol)
        lSelect = M46V999.gfWhereAnd(lSelect) & " Loc.num_endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo)
    End If
    lSelect = lSelect & " order by Loc.num_item"
    mfFormatarConsultaSQL = lSelect
End Function
Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V708B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)
    Dim lAuxNumerico As Long
    Dim lAlter As String
    Dim indAux As Integer, indAuxAnt As Integer

    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_item"), tpoNumerico)
    pobjstc.Num_Item_Ant = pobjstc.Num_Item
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
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
    pobjstc.Val_Vr_Lc = M46V999.gfTratarNulo(prst("Val_Vr_Lc"), tpoNumerico)
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
    pobjstc.Tip_Emissao = M46V999.gfTratarNulo(prst("Tip_Emissao"), tpoNumerico)
    pobjstc.Num_Apol_Anterior = M46V999.gfTratarNulo(prst("Num_Apol_Anterior"), tpoNumerico)
    pobjstc.Num_Item_Anterior = M46V999.gfTratarNulo(prst("Num_Item_Anterior"), tpoNumerico)

    '8000042754 - Resposta dada a questão parametrizada para o plano.
    pobjstc.RespostaPergunta = M46V999.gfTratarNulo(prst("Cod_Resposta"), tpoNumerico)

'    *******************************************************************************************************************
'    *#20180829 - Finalidade: Apura Desconto Comercial Congenere
'    *******************************************************************************************************************
    pobjstc.DescComlCongenereIncorporado = M46V999.gfTratarNulo(prst("DescComlCongenereIncorporado"), tpoNumerico)
  
    '1000347 - Inclusão de informações da congênere em renovações
    On Error Resume Next
    pobjstc.Ren_Con_Seguradora = M46V999.gfTratarNulo(prst("Ren_Con_Seguradora"), tpoNumerico)
    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.PerfilAcesso = 1 Then
        If Err.Number = 3265 Then
            lAlter = " alter table Tab_Ped_Loc add Ren_Con_Seguradora int "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
    End If
    Err.Clear
    '1000347 - Inclusão de informações da congênere em renovações
    pobjstc.Ren_Con_Apolice = M46V999.gfTratarNulo(prst("Ren_Con_Apolice"), tpoString)
    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.PerfilAcesso = 1 Then
        If Err.Number = 3265 Then
            lAlter = " alter table Tab_Ped_Loc add Ren_Con_Apolice text(15) "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
    End If
    Err.Clear
    '1000347 - Inclusão de informações da congênere em renovações
    pobjstc.Ren_Con_Item = M46V999.gfTratarNulo(prst("Ren_Con_Item"), tpoString)
    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.PerfilAcesso = 1 Then
        If Err.Number = 3265 Then
            lAlter = " alter table Tab_Ped_Loc add Ren_Con_Item text(5) "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
    End If
    On Error GoTo 0

    '' 2017-07-28: Juliano - Ind_SinistroSimNao
    pobjstc.Ind_Sinistro = M46V999.gfTratarNulo(prst("Ind_Sinistro"), tpoNumerico)
    pobjstc.Ind_Sinistro_ORIG = pobjstc.Ind_Sinistro

    If pobjstc.NUM_PI > 0 Then
        If (M46V999.GetIndSinistroSimNao(pobjstc.NUM_PI, pobjstc.Num_Apol, pobjstc.Num_Item, indAux, indAuxAnt)) Then
            pobjstc.Ind_Sinistro = indAux
            pobjstc.Ind_Sinistro_ORIG = indAuxAnt
        End If
    End If
    
    ''' 2017-07-28: Juliano - Ind_SinistroSimNao, removido para aceitar o informado
    If pobjstc.Tip_Emissao = e_TipEmissao_RenYasuda Then
        pobjstc.Ind_Sinistro = pobjstc.Apol_Ant_IND_SINISTRO
    End If
    If pobjstc.Ind_Sinistro = 0 Then pobjstc.Ind_Sinistro = e_SimNao_Nao

    ' Essa regra já é executada no metodo mfGraRenItem
    '    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
    '        If pObjStc.Tip_Emissao = e_TipEmissao_RenYasuda Then
    '            pObjStc.mBonus_Anterior = mfRetornaBonusApolice(pObjStc)
    '            If pObjStc.Ind_Sinistro = e_SimNao_Sim Then
    '                lAuxNumerico = pObjStc.mBonus_Anterior - 5
    '            Else
    '                lAuxNumerico = pObjStc.mBonus_Anterior + 5
    '            End If
    '
    '            If lAuxNumerico < 0 Then lAuxNumerico = 0
    '            If M46V999.gCodProduto >= 940 Or pObjStc.Cod_Iden_Tarif >= 940 Then
    '                If pObjStc.Cod_Ramo = e_CodRamo_Empresarial Then
    '                    If M46V999.gCodProduto >= 942 Or pObjStc.Cod_Iden_Tarif >= 942 Then
    '                        'Soma mais 5 da fidelidade.
    '                        lAuxNumerico = lAuxNumerico + 5
    '                    End If
    '                    If lAuxNumerico > 30 Then lAuxNumerico = 30
    '                ElseIf pObjStc.Cod_Ramo = e_CodRamo_Residencial Then
    '                    If M46V999.gCodProduto >= 946 Or pObjStc.Cod_Iden_Tarif >= 946 Then
    '                        'Soma mais 5 da fidelidade.
    '                        lAuxNumerico = lAuxNumerico + 5
    '                        If lAuxNumerico > 30 Then lAuxNumerico = 30
    '                    Else
    '                        If lAuxNumerico > 25 Then lAuxNumerico = 25
    '                    End If
    '                Else
    '                    If lAuxNumerico > 25 Then lAuxNumerico = 25
    '                End If
    '            End If
    '
    '            If pObjStc.Per_Bonus <> lAuxNumerico Then
    '                pObjStc.Per_Bonus = lAuxNumerico
    '                pObjStc.Cod_Sit_item = e_SitItem_NaoCalc
    '            End If
    '        End If
    '    End If
    '************

    On Error Resume Next
    pobjstc.Insp_Obr = M46V999.gfTratarNulo(prst("Insp_Obr"), TpoBoolean)
    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.PerfilAcesso = 1 Then
        If Err.Number = 3265 Then
            lAlter = " alter table Tab_Ped_Loc add Insp_Obr int "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
    End If

    On Error Resume Next
    pobjstc.Insp_Agend = M46V999.gfTratarNulo(prst("Insp_Agend"), TpoBoolean)
    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.PerfilAcesso = 1 Then
        If Err.Number = 3265 Then
            lAlter = " alter table Tab_Ped_Loc add Insp_Agend int "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
    End If

    On Error Resume Next
    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.PerfilAcesso = 1 Then
        pobjstc.Dig_Bens = M46V999.gfTratarNulo(prst("Dig_Bens"), TpoBoolean)
        If Err.Number = 3265 Then
            lAlter = " alter table Tab_Ped_Loc add Dig_Bens int "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If

        On Error Resume Next
        pobjstc.Cob_Bens = M46V999.gfTratarNulo(prst("Cob_Bens"), TpoBoolean)
        If Err.Number = 3265 Then
            lAlter = " alter table Tab_Ped_Loc add Cob_Bens int "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
    End If

    'Ficha 1000305
    On Error Resume Next
    pobjstc.Tip_Cobert = M46V999.gfTratarNulo(prst("Tip_Cobert"), TpoNumInteiro)
    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.PerfilAcesso = 1 Then
        If Err.Number = 3265 Then
            lAlter = " alter table Tab_Ped_Loc add Tip_Cobert int "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
    End If

    'Ficha 1200134
    On Error Resume Next
    pobjstc.Ind_Sin_Inspec = M46V999.gfTratarNulo(prst("Ind_Sin_Inspec"), TpoNumInteiro)
    ''Call M46V999.gpGraLog(0, "Simony3 Item: " & pObjStc.Num_Item & " Total sinistro " & pObjStc.Ind_Sin_Inspec)

    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        If Err.Number = 91 Then    'No objeto clsYasRecordSet o erro é (91 - Object variable or With block variable not set)
            lAlter = " alter table Tab_Ped_Loc add Ind_Sin_Inspec int "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
    End If


    pobjstc.Cod_DizerItem = M46V999.gfTratarNulo(prst("Cod_DizerItem"), TpoNumInteiro)
    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        If Err.Number = 91 Then    'No objeto clsYasRecordSet o erro é (91 - Object variable or With block variable not set)
            lAlter = " alter table Tab_Ped_Loc add Cod_DizerItem int "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
    End If

    'ProjetoFormularios
    pobjstc.Ind_VerbaSep = M46V999.gfTratarNulo(prst("Ind_VerbaSep"), TpoNumInteiro)
    If M46V999.gBaseLocal = True Then
        If Err.Number = 91 Then    'No objeto clsYasRecordSet o erro é (91 - Object variable or With block variable not set)
            lAlter = " alter table Tab_Ped_Loc add Ind_VerbaSep int "
            lAlter = lAlter & " , Val_Vr_Lc double  "
            lAlter = lAlter & " , Val_LMG_Lc double  "
            lAlter = lAlter & " , Per_IS_Vr_Lc double  "
            lAlter = lAlter & " , Coef_IS_Vr_Lc double  "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
    End If
    pobjstc.Val_Vr_Lc = M46V999.gfTratarNulo(prst("Val_Vr_Lc"), TpoValor)
    pobjstc.Val_LMG_Lc = M46V999.gfTratarNulo(prst("Val_LMG_Lc"), TpoValor)
    pobjstc.Per_IS_Vr_Lc = M46V999.gfTratarNulo(prst("Per_IS_Vr_Lc"), TpoValor)
    pobjstc.Coef_IS_Vr_Lc = M46V999.gfTratarNulo(prst("Coef_IS_Vr_Lc"), TpoValor)

    'ProjetoFormularios
    'pobjstc.Ind_VerbaSep = M46V999.gfTratarNulo(prst("Ind_VerbaSep"), TpoNumInteiro)
    'If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
    '    If Err.Number = 91 Then 'No objeto clsYasRecordSet o erro é (91 - Object variable or With block variable not set)
    '        lAlter = " alter table Tab_Ped_Loc_Compl add Ind_VerbaSep int "
    '        If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
    '            Call M46V999.gpGraLog(2, UltimoErro)
    '            Exit Sub
    '        End If
    '    End If
    'End If

    'Capital Global
    pobjstc.Tip_Plano_Vida = M46V999.gfTratarNulo(prst("Tip_Plano_Vida"), TpoNumInteiro)
    If M46V999.gBaseLocal = True Then
        If Err.Number = 91 Then    'No objeto clsYasRecordSet o erro é (91 - Object variable or With block variable not set)
            lAlter = " alter table Tab_Ped_Loc_Compl add Tip_Plano_Vida int "
            lAlter = lAlter & " , Cod_Plano_Cobertura int  "
            lAlter = lAlter & " , Qtd_Vidas int  "
            lAlter = lAlter & " , Val_LMG_Vida double  "
            lAlter = lAlter & " , Ind_Aux_Funeral int  "
            lAlter = lAlter & " , Val_LMG_Funeral double  "
            lAlter = lAlter & " , Ind_Inc_Conjuge int  "
            lAlter = lAlter & " , Val_Aux_Cesta double  "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
    End If
    pobjstc.Cod_Plano_Cobertura = M46V999.gfTratarNulo(prst("Cod_Plano_Cobertura"), tpoNumerico)
    pobjstc.Qtd_Vidas = M46V999.gfTratarNulo(prst("Qtd_Vidas"), tpoNumerico)
    pobjstc.Val_LMG_Vida = M46V999.gfTratarNulo(prst("Val_LMG_Vida"), TpoValor)
    pobjstc.Ind_Aux_Funeral = M46V999.gfTratarNulo(prst("Ind_Aux_Funeral"), tpoNumerico)
    pobjstc.Val_LMG_Funeral = M46V999.gfTratarNulo(prst("Val_LMG_Funeral"), TpoValor)
    pobjstc.Ind_Inc_Conjuge = M46V999.gfTratarNulo(prst("Ind_Inc_Conjuge"), tpoNumerico)
    pobjstc.Val_Aux_Cesta = M46V999.gfTratarNulo(prst("Val_Aux_Cesta"), TpoValor)
    pobjstc.Possui_Elevador = M46V999.gfTratarNulo(prst("Possui_Elevador"), tpoNumerico)
    pobjstc.BonusAlterado = M46V999.gfTratarNulo(prst("BonusAlterado"), TpoBoolean)
    pobjstc.Flg_EnviarCertificado = M46V999.gfTratarNulo(prst("Flg_EnviarCertificado"), tpoNumerico)
    pobjstc.Txt_Observacoes = M46V999.gfTratarNulo(prst("Txt_Observacoes"), tpoString) 'Luiz Alberto - 26/10/2015
    pobjstc.Ind_StatusInspecao = M46V999.gfTratarNulo(prst("Ind_StatusInspecao"), tpoNumerico)
    
    On Error Resume Next
    pobjstc.Val_Vr_Orig = M46V999.gfTratarNulo(prst("Val_Vr_Orig"), tpoNumerico)
    pobjstc.Val_Bas_Orig = M46V999.gfTratarNulo(prst("Val_Bas_Orig"), tpoNumerico)
    pobjstc.Val_Vr_Lc_Orig = M46V999.gfTratarNulo(prst("Val_Vr_Lc_Orig"), tpoNumerico)
    pobjstc.ITEM_ORIGEM_C4 = M46V999.gfTratarNulo(prst("ITEM_ORIGEM_C4"), tpoNumerico)
    pobjstc.ClausulaNaoDepreciacao = M46V999.gfTratarNulo(prst("CLAUSULANAODEPRECIACAO"), tpoNumerico) 'JIRA-25945 - Opção "Sem Depreciação" no Condomínio
End Sub

Public Function mfRetornaBonusApolice(ByRef pobjstc As stcA46V708B, tagDynaTrace As String) As Long
    Dim lSql As String
    Dim lrst As clsYasRecordSet
    Dim lbdP00APOL As ADODB.Connection
    Dim lNomeTab As String
    Dim BonusRetorno1 As Long
    Dim BonusRetorno2 As Long

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfRetornaBonusApolice"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    If M46V999.gPrograma = "D046O003" Then
        lNomeTab = "Syasportal.dbo.TAB_APOL"
    Else
        'If M46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
        'Else
        '   lNomeTab = "GED.DBO.TAB_PED"
        'End If
    End If
    If M46V999.gfAbrBasDados("P00APOL", M46V999.gTipP00APOL, M46V999.gEndP00APOL, lbdP00APOL, UltimoErro) Then
        'If pObjStc.Cod_Ramo = e_CodRamo_Empresarial Then

        lSql = " Select top 1 PER_DESC_EXP from " & lNomeTab & "_cobert WITH (NOLOCK) where num_apol = " & pobjstc.Num_Apol_Anterior
        lSql = lSql & " and num_item = " & pobjstc.Num_Item_Anterior & " order by cod_cobert"
        If M46V999.gfObtRegistro(lbdP00APOL, lSql, lrst, UltimoErro) Then
            If Not lrst.EOF Then
                'mfRetornaBonusApolice = lrst("PER_DESC_EXP")
                BonusRetorno1 = lrst("PER_DESC_EXP")
            Else
                Call M46V999.gpGraLog(0, "Erro - O registro de Bonus em 'mfRetornaBonusApolice' para apólice: " & pobjstc.Num_Apol_Anterior & " não foi encontrado-" & UltimoErro)
                Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "Erro - O registro de Bonus em 'mfRetornaBonusApolice' para apólice: " & pobjstc.Num_Apol_Anterior & " não foi encontrado-" & UltimoErro, 2)
                'GoTo Fim: CR12460 RETORNO DO DESCONTO DE RENOVAÇÃO YASUDA INCORRETO
            End If
        End If

        'Else

        lSql = " Select Per_Bonus from " & lNomeTab & "_Loc WITH (NOLOCK) where num_apol = " & pobjstc.Num_Apol_Anterior
        lSql = lSql & " and num_item = " & pobjstc.Num_Item_Anterior
        If M46V999.gfObtRegistro(lbdP00APOL, lSql, lrst, UltimoErro) Then
            If Not lrst.EOF Then
                'mfRetornaBonusApolice = lrst("Per_Bonus")
                BonusRetorno2 = lrst("Per_Bonus")
            Else
                Call M46V999.gpGraLog(0, "O registro de Bonus em 'mfRetornaBonusApolice' para apólice: " & pobjstc.Num_Apol_Anterior & " não foi encontrado-" & UltimoErro)
                Call M46V999.DynaTrace.gravaLog(lKeyDynaTrace, "O registro de Bonus em 'mfRetornaBonusApolice' para apólice: " & pobjstc.Num_Apol_Anterior & " não foi encontrado-" & UltimoErro, 2)
                'GoTo Fim: CR12460 RETORNO DO DESCONTO DE RENOVAÇÃO YASUDA INCORRETO
            End If
        End If

        If BonusRetorno1 > BonusRetorno2 Then
            mfRetornaBonusApolice = BonusRetorno1
        Else
            mfRetornaBonusApolice = BonusRetorno2
        End If

        If mfRetornaBonusApolice > 30 Then
            mfRetornaBonusApolice = 30
        End If
        
        'piloto marítima
         If pobjstc.Cod_Iden_Tarif >= 1080 Then
            If mfRetornaBonusApolice > 25 Then
                mfRetornaBonusApolice = 25
            End If
        End If
         
        Call M46V999.gpFecha2(lbdP00APOL)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Else
        Call M46V999.gpGraLog(0, "Erro - Abertura do Banco P00APOL em 'mfRetornaBonusApolice' para apólice: " & pobjstc.Num_Apol_Anterior & "-" & UltimoErro)
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "Erro - Abertura do Banco P00APOL em 'mfRetornaBonusApolice' para apólice: " & pobjstc.Num_Apol_Anterior & "-" & UltimoErro, 2, Class_Name, methodName)
    End If
    Set lrst = Nothing


End Function
Public Function mfIncluir(ByRef pobjstc As stcA46V708B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean
    Dim lsSQl As String          'Query.
    Dim lMensagem As String
    Dim lAlter As String
    Dim lRetornar As Byte


    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfIncluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    On Local Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro
Retornar:
    lRetornar = 0

    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro

    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, 2)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro

    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfIncluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErro:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        If Err.Number = 3022 Then   'Já existe.
            UltimoErro = "FTE0027 - Item já existente (nosso número " & pobjstc.nosso_numero & _
                       " - item " & pobjstc.Num_Item & ")."
            Call M46V999.gpGraLog(0, "A46V708A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                     vbCrLf & " SQL: " & lsSQl)
        Else
            UltimoErro = Err.Description
            Call M46V999.gpGraLog(0, "A46V708A.mfIncluir ERRO - " & UltimoErro & _
                                     vbCrLf & " SQL: " & lsSQl)
        End If
        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
        Exit Function
    Else
        If UCase(UltimoErro) Like "*'DIG_BENS'*" Then
            lRetornar = 1
            On Error Resume Next
            If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
                lAlter = " alter table Tab_Ped_Loc add Dig_Bens int "
                If M46V999.gfExeSQL(ConexaoBD, lAlter, lMensagem) <> 0 Then
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                    Exit Function
                End If
                lAlter = " alter table Tab_Ped_Loc add Cob_Bens int "
                If M46V999.gfExeSQL(ConexaoBD, lAlter, lMensagem) <> 0 Then
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                    Exit Function
                End If
                If lRetornar = 1 Then
                    GoTo Retornar
                End If
            End If
        End If
        If UCase(UltimoErro) Like "*'INSP_OBR'*" Then
            lRetornar = 1
            On Error Resume Next
            If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
                lAlter = " alter table Tab_Ped_Loc add Insp_Obr int "
                If M46V999.gfExeSQL(ConexaoBD, lAlter, lMensagem) <> 0 Then
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                    Exit Function
                End If
                lAlter = " alter table Tab_Ped_Loc add Insp_Agend int "
                If M46V999.gfExeSQL(ConexaoBD, lAlter, lMensagem) <> 0 Then
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                    Exit Function
                End If
                If lRetornar = 1 Then
                    GoTo Retornar
                End If
            End If
        End If
        'Ficha 1000305
        If UCase(UltimoErro) Like "*'TIP_COBERT'*" Then
            lRetornar = 1
            On Error Resume Next
            If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
                lAlter = " alter table Tab_Ped_Loc add Tip_Cobert int "
                If M46V999.gfExeSQL(ConexaoBD, lAlter, lMensagem) <> 0 Then
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                    Exit Function
                End If
                If lRetornar = 1 Then
                    GoTo Retornar
                End If
            End If
        End If

        If UCase(UltimoErro) Like "*'REN_CON_SEGURADORA'*" Then
            lRetornar = 1
            lAlter = " alter table Tab_Ped_Loc add Ren_Con_Apolice text(15) "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, lMensagem) <> 0 Then
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                Exit Function
            End If
            lAlter = " alter table Tab_Ped_Loc add Ren_Con_Seguradora int "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, lMensagem) <> 0 Then
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                Exit Function
            End If
            lAlter = " alter table Tab_Ped_Loc add Ren_Con_Item text(5) "
            If M46V999.gfExeSQL(ConexaoBD, lAlter, lMensagem) <> 0 Then
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                Exit Function
            End If

            If lRetornar = 1 Then
                GoTo Retornar
            End If
        End If
        'Ficha 1200134
        If UCase(UltimoErro) Like "*'IND_SIN_INSPEC'*" Then
            lRetornar = 1
            On Error Resume Next
            If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
                lAlter = " alter table Tab_Ped_Loc add Ind_Sin_Inspec int "
                If M46V999.gfExeSQL(ConexaoBD, lAlter, lMensagem) <> 0 Then
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                    Exit Function
                End If
                If lRetornar = 1 Then
                    GoTo Retornar
                End If
            End If
        End If
        If UCase(UltimoErro) Like "*'COD_DIZERITEM'*" Then
            lRetornar = 1
            On Error Resume Next
            If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
                lAlter = " alter table Tab_Ped_Loc add Cod_DizerItem int "
                If M46V999.gfExeSQL(ConexaoBD, lAlter, lMensagem) <> 0 Then
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                    Exit Function
                End If
                If lRetornar = 1 Then
                    GoTo Retornar
                End If
            End If
        End If

        If UCase(UltimoErro) Like "*'IND_VERBASEP'*" Then
            lRetornar = 1
            On Error Resume Next
            If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
                lAlter = " alter table Tab_Ped_Loc add Ind_VerbaSep int "
                lAlter = lAlter & " , Val_Vr_Lc double  "
                lAlter = lAlter & " , Val_LMG_Lc double  "
                lAlter = lAlter & " , Per_IS_Vr_Lc double  "
                lAlter = lAlter & " , Coef_IS_Vr_Lc double  "

                If M46V999.gfExeSQL(ConexaoBD, lAlter, lMensagem) <> 0 Then
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                    Exit Function
                End If
                If lRetornar = 1 Then
                    GoTo Retornar
                End If
            End If
        End If
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)

End Function

Public Function mfAlterar(ByRef pobjstc As stcA46V708B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                          Optional ByVal pRemoverColecao As Boolean = True) As Boolean
    Dim lUpdate As String           'Query.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lUpdate = mfFormatarAlteracaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErro

    lUpdate = mfFormatarAlteracaoSQL(pobjstc, pPropTabela, pNomeTabImport, 2)
    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErro

    lUpdate = "Update " & pPropTabela & pNomeTabImport & "tab_ped_loc set Cod_Iden_Tarif = " & pobjstc.Cod_Iden_Tarif
    lUpdate = lUpdate & " where Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErro

    If Colecao.Count > 0 And pRemoverColecao = True Then
        Colecao.Remove (pobjstc.Class_ID)
    End If
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfAlterar = True
    Exit Function
ControleErro:
    mfAlterar = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V708A.mfAlterar ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lUpdate)
    Else
        Call M46V999.gpGraLog(0, "A46V708A.mfAlterar ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lUpdate)
    End If
End Function

Public Function mfExcluir(ByRef pobjstc As stcA46V708B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfExcluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lsSQl As String           'Query.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    If Not mfExcluirDependentes(pobjstc, pPropTabela, pNomeTabImport) Then GoTo ControleErro

    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro

    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, 2)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro

    If Colecao.Count > 0 Then
        Colecao.Remove (pobjstc.Class_ID)
    End If
    Set pobjstc = Nothing
    mfExcluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErro:
    mfExcluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V708A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V708A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function

Private Sub mfRecuperacaoRegistro(ByVal pobjstc As stcA46V708B)
    Dim lsSQl As String
    Dim lRegAfetados As Long
    
    On Error GoTo Erro:
    
    lsSQl = "INSERT INTO Tab_Ped_Loc "
    lsSQl = lsSQl + "SELECT"
    lsSQl = lsSQl + "   H.NOSSO_NUMERO, "
    lsSQl = lsSQl + "   H.COD_ORG, "
    lsSQl = lsSQl + "   H.COD_RAMO, "
    lsSQl = lsSQl + "   H.NUM_PED, "
    lsSQl = lsSQl + "   H.NUM_APOL, "
    lsSQl = lsSQl + "   H.NUM_ENDO, "
    lsSQl = lsSQl + "   H.NUM_ITEM, "
    lsSQl = lsSQl + "   H.COD_TIP_LOGR_RISCO, "
    lsSQl = lsSQl + "   H.NOM_LOGR_RISCO, "
    lsSQl = lsSQl + "   H.DSC_COMPL_RISCO, "
    lsSQl = lsSQl + "   H.NOM_BAIRRO_RISCO, "
    lsSQl = lsSQl + "   H.NOM_CID_RISCO, "
    lsSQl = lsSQl + "   H.SIG_UF_RISCO, "
    lsSQl = lsSQl + "   H.NUM_CEP_RISCO, "
    lsSQl = lsSQl + "   H.PER_DESC_APAR, "
    lsSQl = lsSQl + "   H.PER_DESC_PROGRES, "
    lsSQl = lsSQl + "   H.PER_IS_VR, "
    lsSQl = lsSQl + "   H.VAL_COEF_IS_VR, "
    lsSQl = lsSQl + "   H.VAL_VR, "
    lsSQl = lsSQl + "   H.TAB_CLASF, "
    lsSQl = lsSQl + "   H.COD_CLASF, "
    lsSQl = lsSQl + "   H.COD_COMPL_CLASF, "
    lsSQl = lsSQl + "   H.DSC_CLASF, "
    lsSQl = lsSQl + "   H.COD_CLASSE_OCUP, "
    lsSQl = lsSQl + "   H.COD_CONSTR, "
    lsSQl = lsSQl + "   H.COD_USER_DIGIT, "
    lsSQl = lsSQl + "   H.DAT_DIGIT, "
    lsSQl = lsSQl + "   H.HOR_DIGIT, "
    lsSQl = lsSQl + "   H.COD_USER_LIBER, "
    lsSQl = lsSQl + "   H.DAT_LIBER, "
    lsSQl = lsSQl + "   H.HOR_LIBER, "
    lsSQl = lsSQl + "   H.COD_SIT_ITEM, "
    lsSQl = lsSQl + "   H.PER_BONUS, "
    lsSQl = lsSQl + "   H.QTD_EMPREG, "
    lsSQl = lsSQl + "   H.NUM_INSP, "
    lsSQl = lsSQl + "   H.COD_IDEN_SEGUR_01, "
    lsSQl = lsSQl + "   H.COD_IDEN_SEGUR_02, "
    lsSQl = lsSQl + "   H.COD_IDEN_SEGUR_03, "
    lsSQl = lsSQl + "   H.COD_IDEN_SEGUR_04, "
    lsSQl = lsSQl + "   H.COD_IDEN_SEGUR_05, "
    lsSQl = lsSQl + "   H.COD_IDEN_SEGUR_06, "
    lsSQl = lsSQl + "   H.COD_IDEN_SEGUR_07, "
    lsSQl = lsSQl + "   H.COD_IDEN_SEGUR_08, "
    lsSQl = lsSQl + "   H.COD_IDEN_SEGUR_09, "
    lsSQl = lsSQl + "   H.COD_IDEN_SEGUR_10, "
    lsSQl = lsSQl + "   H.PER_TOT_SEGUR, "
    lsSQl = lsSQl + "   H.COD_IDEN_PROTEC_01, "
    lsSQl = lsSQl + "   H.COD_IDEN_PROTEC_02, "
    lsSQl = lsSQl + "   H.COD_IDEN_PROTEC_03, "
    lsSQl = lsSQl + "   H.COD_IDEN_PROTEC_04, "
    lsSQl = lsSQl + "   H.COD_IDEN_PROTEC_05, "
    lsSQl = lsSQl + "   H.COD_IDEN_PROTEC_06, "
    lsSQl = lsSQl + "   H.COD_IDEN_PROTEC_07, "
    lsSQl = lsSQl + "   H.COD_IDEN_PROTEC_08, "
    lsSQl = lsSQl + "   H.COD_IDEN_PROTEC_09, "
    lsSQl = lsSQl + "   H.COD_IDEN_PROTEC_10, "
    lsSQl = lsSQl + "   H.PER_TOT_PROTEC, "
    lsSQl = lsSQl + "   H.COD_PRODUTO, "
    lsSQl = lsSQl + "   H.COD_REGIAO, "
    lsSQl = lsSQl + "   H.COD_IDEN_TARIF, "
    lsSQl = lsSQl + "   H.NUM_PI, "
    lsSQl = lsSQl + "   H.TIP_EMISSAO, "
    lsSQl = lsSQl + "   H.NUM_APOL_ANTERIOR, "
    lsSQl = lsSQl + "   H.NUM_ITEM_ANTERIOR, "
    lsSQl = lsSQl + "   H.IND_SINISTRO, "
    lsSQl = lsSQl + "   H.REN_CON_SEGURADORA, "
    lsSQl = lsSQl + "   H.REN_CON_APOLICE, "
    lsSQl = lsSQl + "   H.REN_CON_ITEM, "
    lsSQl = lsSQl + "   H.TIP_COBERT, "
    lsSQl = lsSQl + "   H.Insp_Obr, "
    lsSQl = lsSQl + "   H.Insp_Agend, "
    lsSQl = lsSQl + "   H.Dig_Bens, "
    lsSQl = lsSQl + "   H.Cob_Bens, "
    lsSQl = lsSQl + "   H.Ind_Sin_Inspec, "
    lsSQl = lsSQl + "   H.Cod_DizerItem "
    lsSQl = lsSQl + "FROM   "
    lsSQl = lsSQl + "   Historico2..Tab_Ped_Loc_Alt H "
    lsSQl = lsSQl + "JOIN"
    lsSQl = lsSQl + "   ("
    lsSQl = lsSQl + "   SELECT  NOSSO_NUMERO, NUM_ITEM, MAX(Dat_Inc_Alt) Dat_Inc_Alt "
    lsSQl = lsSQl + "   FROM    Historico2..Tab_Ped_Loc_Alt "
    lsSQl = lsSQl + "   WHERE   NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    lsSQl = lsSQl + "   group by NOSSO_NUMERO, NUM_ITEM "
    lsSQl = lsSQl + "   ) H2 "
    lsSQl = lsSQl + "ON "
    lsSQl = lsSQl + "   H.NOSSO_NUMERO = H2.NOSSO_NUMERO AND "
    lsSQl = lsSQl + "   H.NUM_ITEM = H2.NUM_ITEM AND "
    lsSQl = lsSQl + "   H.Dat_Inc_Alt = H2.Dat_Inc_Alt "
    
    UltimoErro = Empty
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, _
                         UltimoErro, lRegAfetados) <> 0 Then GoTo Erro
                         
    If lRegAfetados > 0 Then
        Call M46V999.gpGraLog(0, lRegAfetados & " Registros Tab_Ped_Loc recuperados do historico  Nosso Numero: " & pobjstc.nosso_numero)
    Else
        Call M46V999.gpGraLog(0, "Nenhum Registros da Tab_Ped_Loc pode ser recuperado do historico Nosso Numero: " & pobjstc.nosso_numero)
    End If
                         
Fim:
    
Erro:
    Call M46V999.gpGraLog(0, "A46V708A.mfRecuperacaoRegistro ERRO - " & Err & " " & IIf(Err.Description = Empty, UltimoErro, Err.Description) & _
                         vbCrLf & " SQL: " & lsSQl)
End Sub

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V708B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String, _
                                  Optional ByVal pCarregarDependentes As Boolean = False) As Boolean
    Dim lobjstc As stcA46V708B          'Objeto da classe de estrutura stcA46V708B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    
    ' ****** Caso não encontre nenhum item, força a recuperação do Historico
    If lrst.EOF And pPropTabela = Empty And pNomeTabImport = Empty And Not (pobjstc.nosso_numero Like "PADRAO*") Then
        Call mfRecuperacaoRegistro(pobjstc)
        
        If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErro
    End If
    
    Do While Not lrst.EOF

        Set lobjstc = New stcA46V708B
        Set lobjstc.M46V999 = M46V999

        Call mpCarregarAtributos(lobjstc, lrst)
        'Impedir que o Bonus seja recalculado para impressões do Kit
        If lobjstc.M46V999.gOrigemChamada = e_Origem_Kit Then
            lobjstc.BonusAlterado = True
        End If
        
        Set lobjstc.ObjPai = pobjstc.ObjPai
        If pCarregarDependentes = True Then
            'Call M46V999.gpGraLog(0, "mfCarregarDependentes - Tab_Ped_loc - Apolice: " & lobjstc.Num_Apol & " Item: " & lobjstc.Num_Item)
            If Not mfCarregarDependentes(lobjstc, pPropTabela, pNomeTabImport) Then GoTo ControleErro
        End If
        
        If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
            If lobjstc.objstcPedObs.Observacao Like "*OBRIGATORIA A REALIZACAO PREVIA DE INSPECAO DO RISCO*" Or _
               lobjstc.objstcPedObs.Observacao Like "%ACEITACAO CONDICIONADA A REALIZACAO DA INSPECAO.%" Then
                lobjstc.Insp_Obr = True
            End If
        End If
        Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID
        lrst.MoveNext
    Loop
    mfCarregarColecao = True
SAIDA:
    GoSub FechaObj
    Exit Function

ControleErro:
    mfCarregarColecao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V708A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V708A.mfCarregarColecao ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    GoTo SAIDA
FechaObj:
    Set lobjstc = Nothing
    Set lrst = Nothing
    Return
End Function

Private Function mfCarregarDependentes(ByRef pobjstc As stcA46V708B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As Boolean
    Dim lobjCobertura As clsA46V704A    'Objeto da classe clsA46V704A (TAB_PED_COBERT)
    Dim lobjstcCobertura As stcA46V704B    'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
    Dim lobjClausula As clsA46V703A    'Objeto da classe clsA46V703A (TAB_PED_CLAU)
    Dim lobjstcPedClau As stcA46V703B    'Objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU)
    Dim lobjPedObservacao As clsA46V714A    'Objeto da classe clsA46V714A (TAB_PED_OBS)
    Dim lobjBem As clsA46V715A    'Objeto da classe clsA46V715A (TAB_PED_LOC_BEM)
    Dim lobjstcBem As stcA46V715B    'Objeto da classe de estrutura stcA46V715B (TAB_PED_LOC_BEM)
    Dim lobjLocBicicleta As clsA46V744A             'Objeto da classe clsA46V715A (TAB_PED_LOC_BICICLETA)
    Dim lobjstcLocBicicleta As stcA46V744B          'Objeto da classe de estrutura stcA46V715B (TAB_PED_LOC_BICICLETA)
    Dim lobjSeguro As clsA46V716A    'Objeto da classe clsA46V716A (TAB_PED_LOC_SEGURO)
    Dim lobjstcSeguro As stcA46V716B    'Objeto da classe de estrutura stcA46V716B (TAB_PED_LOC_SEGURO)
    Dim lobjInsp As clsA46V717A    'Objeto da classe clsA46V717A (TAB_PED_INSP)
    Dim lObjstcInsp As stcA46V717B    'Objeto da classe de estrutura stcA46V717B (TAB_PED_INSP)
    Dim lobjFator As clsA46V720A    'Objeto da classe clsA46V720A (Tab_Ped_Fat_Risco)
    Dim lobjstcFator As stcA46V720B    'Objeto da classe de estrutura stcA46V720B (Tab_Ped_Fat_Risco)
    Dim lobjLMI As clsA46V728A    'Objeto da classe clsA46V728A (Tab_Ped_Loc_LMI)
    Dim lobjstcLMI As stcA46V728B    'Objeto da classe de estrutura stcA46V728B (Tab_Ped_Loc_LMI)

    Dim lobjPessoa As clsA46V735A    'Objeto da classe clsA46V735A (TAB_PED_LOC_PESSOA)
    Dim lobjstcPessoa As stcA46V735B    'Objeto da classe de estrutura stcA46V735B (TAB_PED_LOC_PESSOA)
    Dim lobjNovaFranquia As clsA46V738A     'Objeto da classe clsA46V738A (TAB_PED_COBERT_FRANQUIA)
    Dim lobjstcNovaFranquia As stcA46V738B  'Objeto da classe de estrutura stcA46V738B (TAB_PED_COBERT_FRANQUIA)

    Dim lobjNovoTexto As clsA46V739A        'Objeto da classe clsA46V739A (TAB_PED_TEXTOS)
    Dim lobjstcNovoTexto As stcA46V739B  'Objeto da classe de estrutura stcA46V739B (TAB_PED_TEXTOS)

    Dim lobjAllRisks As clsA46V747A        'Objeto da classe clsA46V747A (Tab_Ped_All_Risks)
    Dim lobjstcAllRisks As stcA46V747B  'Objeto da classe de estrutura stcA46V747B (Tab_Ped_All_Risks)
    
    Dim lobjObraArte As clsA46V748A        'Objeto da classe clsA46V748A (tab_ped_obras_arte)
    Dim lobjstcObraArte As stcA46V748B  'Objeto da classe de estrutura clsA46V748A (tab_ped_obras_arte)


    On Error GoTo ControleErro
    mfCarregarDependentes = False

    'Carregar coleção de coberturas
    Set lobjCobertura = New clsA46V704A
    Set lobjCobertura.M46V999 = M46V999

    Set lobjstcCobertura = New stcA46V704B
    Set lobjstcCobertura.M46V999 = M46V999

    lobjstcCobertura.nosso_numero = pobjstc.nosso_numero
    lobjstcCobertura.Num_Apol = pobjstc.Num_Apol
    lobjstcCobertura.Num_Endo = pobjstc.Num_Endo
    lobjstcCobertura.Num_Item = pobjstc.Num_Item
    Set lobjstcCobertura.ObjPai = pobjstc
    Set lobjCobertura.ConexaoBD = Me.ConexaoBD
    If Not lobjCobertura.mfCarregarColecao(lobjstcCobertura, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjCobertura.UltimoErro
        GoTo ControleErro
    End If
    Set pobjstc.ColecaoPedCobert = lobjCobertura.Colecao

    Set Me.ConexaoBD = lobjCobertura.ConexaoBD
    Set lobjCobertura = Nothing
    Set lobjstcCobertura = Nothing

    'Carregar coleção de cláusulas
    Set lobjClausula = New clsA46V703A
    Set lobjClausula.M46V999 = M46V999

    'Set pObjStc.objstcPedClau = New stcA46V703B
    'Set pObjStc.objstcPedClau.M46V999 = M46V999

    'pObjStc.objstcPedClau.Nosso_Numero = pObjStc.Nosso_Numero
    'pObjStc.objstcPedClau.Num_Item = pObjStc.Num_Item
    Set lobjstcPedClau = New stcA46V703B
    Set lobjstcPedClau.M46V999 = M46V999

    lobjstcPedClau.nosso_numero = pobjstc.nosso_numero
    lobjstcPedClau.Num_Apol = pobjstc.Num_Apol
    lobjstcPedClau.Num_Endo = pobjstc.Num_Endo
    lobjstcPedClau.Num_Item = pobjstc.Num_Item
    Set lobjstcPedClau.ObjPai = pobjstc
    Set lobjClausula.ConexaoBD = Me.ConexaoBD
    'If Not lobjClausula.mfCarregarColecao(pObjStc.objstcPedClau) Then
    If Not lobjClausula.mfCarregarColecao(lobjstcPedClau, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjClausula.UltimoErro
        GoTo ControleErro
    End If
    'If lobjClausula.Colecao.Count > 0 Then
    '    Set pObjStc.objstcPedClau = lobjClausula.Colecao.Item(1)
    'Else
    '    Set pObjStc.objstcPedClau = New stcA46V703B
    '    Set pObjStc.objstcPedClau.M46V999 = M46V999
    'End If
    Set pobjstc.ColecaoPedClau = lobjClausula.Colecao
    Set Me.ConexaoBD = lobjClausula.ConexaoBD
    Set lobjClausula = Nothing

    'Carregar observação
    Set lobjPedObservacao = New clsA46V714A
    Set lobjPedObservacao.M46V999 = M46V999

    Set lobjPedObservacao.ConexaoBD = Me.ConexaoBD
    pobjstc.objstcPedObs.nosso_numero = pobjstc.nosso_numero
    pobjstc.objstcPedObs.Num_Item = pobjstc.Num_Item
    Set pobjstc.objstcPedObs.ObjPai = pobjstc
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjPedObservacao.mfCarregarColecao(pobjstc.objstcPedObs, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjPedObservacao.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjPedObservacao.ConexaoBD
    If lobjPedObservacao.Colecao.Count > 0 Then
        Set pobjstc.objstcPedObs = lobjPedObservacao.Colecao.Item(1)
    End If
    Set lobjPedObservacao = Nothing

    'Carregar coleção de bem.
    Set lobjBem = New clsA46V715A
    Set lobjBem.M46V999 = M46V999

    Set lobjstcBem = New stcA46V715B
    Set lobjstcBem.M46V999 = M46V999

    lobjstcBem.nosso_numero = pobjstc.nosso_numero
    lobjstcBem.Num_Item = pobjstc.Num_Item
    lobjstcBem.Num_Apol = pobjstc.Num_Apol
    lobjstcBem.Num_Endo = pobjstc.Num_Endo
    Set lobjstcBem.ObjPai = pobjstc
    Set lobjBem.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjBem.mfCarregarColecao(lobjstcBem, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjBem.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjBem.ConexaoBD
    Set pobjstc.ColecaoPedLocBem = lobjBem.Colecao
    Set lobjBem = Nothing
    Set lobjstcBem = Nothing

    'Carregar coleção de bem.
    Set lobjLocBicicleta = New clsA46V744A
    Set lobjLocBicicleta.M46V999 = M46V999

    Set lobjstcLocBicicleta = New stcA46V744B
    Set lobjstcLocBicicleta.M46V999 = M46V999

    lobjstcLocBicicleta.nosso_numero = pobjstc.nosso_numero
    lobjstcLocBicicleta.Num_Item = pobjstc.Num_Item
    lobjstcLocBicicleta.Num_Apol = pobjstc.Num_Apol
    lobjstcLocBicicleta.Num_Endo = pobjstc.Num_Endo
    Set lobjstcLocBicicleta.ObjPai = pobjstc
    Set lobjLocBicicleta.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjLocBicicleta.mfCarregarColecao(lobjstcLocBicicleta, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjLocBicicleta.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjLocBicicleta.ConexaoBD
    Set pobjstc.ColecaoPedLocBicicleta = lobjLocBicicleta.Colecao
    Set lobjLocBicicleta = Nothing
    Set lobjstcLocBicicleta = Nothing

    'Carregar coleção de All Risks.
    Set lobjAllRisks = New clsA46V747A
    Set lobjAllRisks.M46V999 = M46V999

    Set lobjstcAllRisks = New stcA46V747B
    Set lobjstcAllRisks.M46V999 = M46V999

    lobjstcAllRisks.nosso_numero = pobjstc.nosso_numero
    lobjstcAllRisks.Num_Item = pobjstc.Num_Item
    lobjstcAllRisks.Num_Apol = pobjstc.Num_Apol
    lobjstcAllRisks.Num_Endo = pobjstc.Num_Endo
    Set lobjstcAllRisks.ObjPai = pobjstc
    Set lobjAllRisks.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjAllRisks.mfCarregarColecao(lobjstcAllRisks, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjAllRisks.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjAllRisks.ConexaoBD
    Set pobjstc.ColecaoAllRisks = lobjAllRisks.Colecao
    Set lobjAllRisks = Nothing
    Set lobjstcAllRisks = Nothing
    
    'Carregar coleção de Obras Arte.
    Set lobjObraArte = New clsA46V748A
    Set lobjObraArte.M46V999 = M46V999

    Set lobjstcObraArte = New stcA46V748B
    Set lobjstcObraArte.M46V999 = M46V999

    lobjstcObraArte.nosso_numero = pobjstc.nosso_numero
    lobjstcObraArte.Num_Item = pobjstc.Num_Item
    lobjstcObraArte.Num_Apol = pobjstc.Num_Apol
    lobjstcObraArte.Num_Endo = pobjstc.Num_Endo
    Set lobjstcObraArte.ObjPai = pobjstc
    Set lobjObraArte.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjObraArte.mfCarregarColecao(lobjstcObraArte, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjObraArte.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjObraArte.ConexaoBD
    Set pobjstc.ColecaoObrasArte = lobjObraArte.Colecao
    Set lobjObraArte = Nothing
    Set lobjstcObraArte = Nothing

    Set lobjSeguro = New clsA46V716A
    Set lobjSeguro.M46V999 = M46V999

    Set lobjstcSeguro = New stcA46V716B
    lobjstcSeguro.nosso_numero = pobjstc.nosso_numero
    lobjstcSeguro.Num_Item = pobjstc.Num_Item
    Set lobjstcSeguro.ObjPai = pobjstc
    Set lobjSeguro.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjSeguro.mfCarregarColecao(lobjstcSeguro, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjSeguro.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set pobjstc.ColecaoPedSeguro = lobjSeguro.Colecao
    Set Me.ConexaoBD = lobjSeguro.ConexaoBD
    Set lobjSeguro = Nothing
    Set lobjstcSeguro = Nothing

    'Carrega o lobjeto de inspeção.
    Set lobjInsp = New clsA46V717A
    Set lobjInsp.M46V999 = M46V999

    Set lobjInsp.ConexaoBD = Me.ConexaoBD
    pobjstc.objstcPedInsp.nosso_numero = pobjstc.nosso_numero
    pobjstc.objstcPedInsp.Num_Item = pobjstc.Num_Item
    pobjstc.objstcPedInsp.NUM_PI = pobjstc.NUM_PI
    Set pobjstc.objstcPedInsp.ObjPai = pobjstc
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjInsp.mfCarregarColecao(pobjstc.objstcPedInsp, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjInsp.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjInsp.ConexaoBD
    If lobjInsp.Colecao.Count > 0 Then
        Set pobjstc.objstcPedInsp = lobjInsp.Colecao.Item(1)
    End If
    Set lobjInsp = Nothing

    'Carregar coleção de fatores de Risco
    Set lobjFator = New clsA46V720A
    Set lobjFator.M46V999 = M46V999

    Set lobjstcFator = New stcA46V720B
    lobjstcFator.nosso_numero = pobjstc.nosso_numero
    lobjstcFator.Num_Item = pobjstc.Num_Item
    lobjstcFator.Num_Apol = pobjstc.Num_Apol
    lobjstcFator.Num_Endo = pobjstc.Num_Endo
    Set lobjstcFator.ObjPai = pobjstc
    Set lobjFator.ConexaoBD = Me.ConexaoBD
    If Not lobjFator.mfCarregarColecao(lobjstcFator, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjFator.UltimoErro
        GoTo ControleErro
    End If
    Set pobjstc.ColecaoPedFator = lobjFator.Colecao
    Set Me.ConexaoBD = lobjFator.ConexaoBD
    Set lobjFator = Nothing
    Set lobjstcFator = Nothing

    'Carregar coleção de LMI
    If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda And M46V999.gTipP00MPED = 2 And UCase(pPropTabela) <> "GED.DBO." Then
        Set lobjLMI = New clsA46V728A
        Set lobjLMI.M46V999 = M46V999

        Set lobjstcLMI = New stcA46V728B
        lobjstcLMI.nosso_numero = pobjstc.nosso_numero
        lobjstcLMI.Num_Item = pobjstc.Num_Item
        Set lobjstcLMI.ObjPai = pobjstc
        Set lobjLMI.ConexaoBD = Me.ConexaoBD
        If Not lobjLMI.mfCarregarColecao(lobjstcLMI, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjLMI.UltimoErro
            GoTo ControleErro
        End If
        Set pobjstc.ColecaoLMI = lobjLMI.Colecao

        Set Me.ConexaoBD = lobjLMI.ConexaoBD
        Set lobjLMI = Nothing
        Set lobjstcLMI = Nothing
    End If


    'Carregar coleção de Inquilino.
    'ProjetoFormularios
    Set lobjPessoa = New clsA46V735A
    Set lobjPessoa.M46V999 = M46V999

    Set lobjstcPessoa = New stcA46V735B

    lobjstcPessoa.nosso_numero = pobjstc.nosso_numero
    lobjstcPessoa.Num_Item = pobjstc.Num_Item
    lobjstcPessoa.Num_Apol = pobjstc.Num_Apol
    lobjstcPessoa.Num_Endo = pobjstc.Num_Endo
    lobjstcPessoa.Tip_Loc_Pessoa = e_TipLocPessoa_Inquilino
    Set lobjstcPessoa.ObjPai = pobjstc

    Set lobjPessoa.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjPessoa.mfCarregarColecao(lobjstcPessoa, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjPessoa.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjPessoa.ConexaoBD
    If lobjPessoa.Colecao.Count > 0 Then
        Set pobjstc.objstcPedInquilino = lobjPessoa.Colecao(1)
    End If
    Set lobjPessoa = Nothing
    Set lobjstcPessoa = Nothing

    'ProjetoFormularios
    'Carregar coleção de proprietarios.
    Set lobjPessoa = New clsA46V735A
    Set lobjPessoa.M46V999 = M46V999

    Set lobjstcPessoa = New stcA46V735B

    lobjstcPessoa.nosso_numero = pobjstc.nosso_numero
    lobjstcPessoa.Num_Item = pobjstc.Num_Item
    lobjstcPessoa.Num_Apol = pobjstc.Num_Apol
    lobjstcPessoa.Num_Endo = pobjstc.Num_Endo
    lobjstcPessoa.Tip_Loc_Pessoa = e_TipLocPessoa_Proprietario
    Set lobjstcPessoa.ObjPai = pobjstc

    Set lobjPessoa.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjPessoa.mfCarregarColecao(lobjstcPessoa, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjPessoa.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjPessoa.ConexaoBD

    If lobjPessoa.Colecao.Count > 0 Then
        Set pobjstc.objstcPedProprietario = lobjPessoa.Colecao(1)
    End If
    Set lobjPessoa = Nothing
    Set lobjstcPessoa = Nothing
    '

    'ProjetoFormularios
    'Carregar coleção de beneficiario.

    Set lobjPessoa = New clsA46V735A
    Set lobjPessoa.M46V999 = M46V999

    Set lobjstcPessoa = New stcA46V735B

    lobjstcPessoa.nosso_numero = pobjstc.nosso_numero
    lobjstcPessoa.Num_Item = pobjstc.Num_Item
    lobjstcPessoa.Num_Apol = pobjstc.Num_Apol
    lobjstcPessoa.Num_Endo = pobjstc.Num_Endo
    lobjstcPessoa.Tip_Loc_Pessoa = e_TipLocPessoa_Beneficiario
    Set lobjstcPessoa.ObjPai = pobjstc

    Set lobjPessoa.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjPessoa.mfCarregarColecao(lobjstcPessoa, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjPessoa.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjPessoa.ConexaoBD
    Set pobjstc.ColecaoBeneficiario = lobjPessoa.Colecao
    Set lobjPessoa = Nothing
    Set lobjstcPessoa = Nothing

    'Textos
    'Carregar coleção de Textos.

    Set lobjNovoTexto = New clsA46V739A
    Set lobjNovoTexto.M46V999 = M46V999

    Set lobjstcNovoTexto = New stcA46V739B

    lobjstcNovoTexto.nosso_numero = pobjstc.nosso_numero
    lobjstcNovoTexto.Num_Item = pobjstc.Num_Item
    lobjstcNovoTexto.Num_Apol = pobjstc.Num_Apol
    lobjstcNovoTexto.Num_Endo = pobjstc.Num_Endo
    lobjstcNovoTexto.NUM_PI = pobjstc.NUM_PI
    Set lobjstcNovoTexto.ObjPai = pobjstc

    Set lobjNovoTexto.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjNovoTexto.mfCarregarColecao(lobjstcNovoTexto, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjNovoTexto.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjNovoTexto.ConexaoBD
    Set pobjstc.ColecaoTextos = lobjNovoTexto.Colecao
    Set lobjNovoTexto = Nothing
    Set lobjstcNovoTexto = Nothing

    mfCarregarDependentes = True

SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V708A.mfCarregarDependentes ERRO - " & Err & " " & Err.Description)
    Else
        Call M46V999.gpGraLog(0, "A46V708A.mfCarregarDependentes ERRO - " & UltimoErro)
    End If
    GoTo SAIDA
FechaObj:
    Set lobjCobertura = Nothing
    Set lobjstcCobertura = Nothing
    Set lobjClausula = Nothing
    Set lobjstcPedClau = Nothing
    Set lobjPedObservacao = Nothing
    Set lobjBem = Nothing
    Set lobjstcBem = Nothing
    Set lobjLocBicicleta = Nothing
    Set lobjstcLocBicicleta = Nothing
    Set lobjSeguro = Nothing
    Set lobjstcSeguro = Nothing
    Set lobjInsp = Nothing
    Set lObjstcInsp = Nothing
    Set lobjFator = Nothing
    Set lobjstcFator = Nothing
    Set lobjLMI = Nothing
    Set lobjstcLMI = Nothing
    Set lobjPessoa = Nothing
    Set lobjstcPessoa = Nothing
    Set lobjNovaFranquia = Nothing
    Set lobjstcNovaFranquia = Nothing
    Set lobjNovoTexto = Nothing
    Set lobjstcNovoTexto = Nothing

    Return
End Function

Private Function mfExcluirDependentes(ByRef pobjstc As stcA46V708B, _
                                      ByVal pPropTabela As String, _
                                      ByVal pNomeTabImport As String) As Boolean
    Dim lobjCobertura As clsA46V704A    'objeto da classe clsA46V704A (TAB_PED_COBERT)
    Dim lobjstcCobertura As stcA46V704B  'objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
    Dim lobjClausula As clsA46V703A    'objeto da classe clsA46V703A (TAB_PED_CLAU)
    Dim lobjStcClausula As stcA46V703B    'objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU)
    Dim lobjObservacao As clsA46V714A    'objeto da classe clsA46V714A (TAB_PED_OBS)
    Dim lobjBem As clsA46V715A              'objeto da classe clsA46V715A (TAB_PED_LOC_BEM)
    Dim lobjstcBem As stcA46V715B           'objeto da classe de estrutura stcA46V715B (TAB_PED_LOC_BEM)
    Dim lobjLocBicicleta As clsA46V744A     'objeto da classe clsA46V714A (TAB_PED_LOC_BICICLETA)
    Dim lobjstcLocBicicleta As stcA46V744B  'objeto da classe de estrutura stcA46V714B (TAB_PED_LOC_BICICLETA)
    
    Dim lobjSeguro As clsA46V716A    'objeto da classe clsA46V716A (TAB_PED_LOC_SEGURO)
    Dim lobjstcSeguro As stcA46V716B    'objeto da classe de estrutura stcA46V716B (TAB_PED_LOC_SEGURO)
    Dim lobjInsp As clsA46V717A    'objeto da classe clsA46V717A (TAB_PED_INSP)
    Dim lobjFator As clsA46V720A    'objeto da classe clsA46V720A (Tab_Ped_Fat_Risco)
    Dim lobjstcFator As stcA46V720B    'objeto da classe de estrutura stcA46V720B (Tab_Ped_Fat_Risco)
    Dim lobjLMI As clsA46V728A    'objeto da classe clsA46V728A (Tab_Ped_Loc_LMI)
    Dim lobjstcLMI As stcA46V728B    'objeto da classe de estrutura stcA46V728B (Tab_Ped_Loc_LMI)

    Dim lobjPessoa As clsA46V735A    'objeto da classe clsA46V728A (Tab_Ped_Loc_Pessoa)
    Dim lobjstcPessoa As stcA46V735B    'objeto da classe de estrutura stcA46V728B (Tab_Ped_Loc_Pessoa)

    Dim lobjTexto As clsA46V739A        'objeto da classe clsA46V739A (Tab_Ped_Textos)
    Dim lobjstcTexto As stcA46V739B    'objeto da classe de estrutura stcA46V728B (Tab_Ped_Textos)

    Dim lobjAllRisks As clsA46V747A        'Objeto da classe clsA46V747A (Tab_Ped_All_Risks)
    Dim lobjstcAllRisks As stcA46V747B  'Objeto da classe de estrutura stcA46V747B (Tab_Ped_All_Risks)
    
    Dim lobjObraArte As clsA46V748A        'Objeto da classe clsA46V748A (tab_ped_obras_arte)
    Dim lobjstcObraArte As stcA46V748B  'Objeto da classe de estrutura clsA46V748A (tab_ped_obras_arte)

    On Error GoTo ControleErro
    mfExcluirDependentes = False

    'Exclui coberturas
    Set lobjCobertura = New clsA46V704A
    Set lobjCobertura.M46V999 = M46V999

    Set lobjstcCobertura = New stcA46V704B
    Set lobjstcCobertura.M46V999 = M46V999

    lobjstcCobertura.nosso_numero = pobjstc.nosso_numero
    lobjstcCobertura.Num_Item = pobjstc.Num_Item
    Set lobjstcCobertura.ObjPai = pobjstc
    Set lobjCobertura.ConexaoBD = Me.ConexaoBD
    If Not lobjCobertura.mfExcluir(lobjstcCobertura, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjCobertura.UltimoErro
        GoTo ControleErro
    End If
    Set Me.ConexaoBD = lobjCobertura.ConexaoBD
    Set lobjCobertura = Nothing
    Set lobjstcCobertura = Nothing

    'Exclui cláusulas
    Set lobjClausula = New clsA46V703A
    Set lobjClausula.M46V999 = M46V999

    Set lobjStcClausula = New stcA46V703B
    Set lobjStcClausula.M46V999 = M46V999

    lobjStcClausula.nosso_numero = pobjstc.nosso_numero
    lobjStcClausula.Num_Item = pobjstc.Num_Item
    Set lobjStcClausula.ObjPai = pobjstc
    Set lobjClausula.ConexaoBD = Me.ConexaoBD
    If Not lobjClausula.mfExcluir(lobjStcClausula, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjClausula.UltimoErro
        GoTo ControleErro
    End If
    Set Me.ConexaoBD = lobjClausula.ConexaoBD
    Set lobjClausula = Nothing
    Set lobjStcClausula = Nothing

    'Exclui as observações CRQ000000425129
    If Not gInd_CotacaoSolic(pPropTabela, pobjstc.nosso_numero) Then
            Set lobjObservacao = New clsA46V714A
            Set lobjObservacao.M46V999 = M46V999
        
            Set lobjObservacao.ConexaoBD = Me.ConexaoBD
            pobjstc.objstcPedObs.Num_Item = pobjstc.Num_Item
            pobjstc.objstcPedObs.nosso_numero = pobjstc.nosso_numero
            Set pobjstc.objstcPedObs.ObjPai = pobjstc
            If Not lobjObservacao.mfExcluir(pobjstc.objstcPedObs, pPropTabela, pNomeTabImport) Then
                UltimoErro = lobjObservacao.UltimoErro
                GoTo ControleErro
            End If
            Set Me.ConexaoBD = lobjObservacao.ConexaoBD
            Set lobjObservacao = Nothing
    End If

    'Excluir bens
    Set lobjBem = New clsA46V715A
    Set lobjBem.M46V999 = M46V999

    Set lobjstcBem = New stcA46V715B
    Set lobjstcBem.M46V999 = M46V999

    lobjstcBem.nosso_numero = pobjstc.nosso_numero
    lobjstcBem.Num_Item = pobjstc.Num_Item
    Set lobjstcBem.ObjPai = pobjstc
    Set lobjBem.ConexaoBD = Me.ConexaoBD
    If Not lobjBem.mfExcluir(lobjstcBem, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjBem.UltimoErro
        GoTo ControleErro
    End If
    Set Me.ConexaoBD = lobjBem.ConexaoBD
    Set lobjBem = Nothing
    Set lobjstcBem = Nothing

    'Excluir bicicletas
    Set lobjLocBicicleta = New clsA46V744A
    Set lobjLocBicicleta.M46V999 = M46V999

    Set lobjstcLocBicicleta = New stcA46V744B
    Set lobjstcLocBicicleta.M46V999 = M46V999

    lobjstcLocBicicleta.nosso_numero = pobjstc.nosso_numero
    lobjstcLocBicicleta.Num_Item = pobjstc.Num_Item
    Set lobjstcLocBicicleta.ObjPai = pobjstc
    Set lobjLocBicicleta.ConexaoBD = Me.ConexaoBD
    If Not lobjLocBicicleta.mfExcluir(lobjstcLocBicicleta, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjLocBicicleta.UltimoErro
        GoTo ControleErro
    End If
    Set Me.ConexaoBD = lobjLocBicicleta.ConexaoBD
    Set lobjLocBicicleta = Nothing
    Set lobjstcLocBicicleta = Nothing

    Set lobjSeguro = New clsA46V716A
    Set lobjSeguro.M46V999 = M46V999

    Set lobjstcSeguro = New stcA46V716B
    lobjstcSeguro.nosso_numero = pobjstc.nosso_numero
    lobjstcSeguro.Num_Item = pobjstc.Num_Item
    Set lobjstcSeguro.ObjPai = pobjstc
    Set lobjSeguro.ConexaoBD = Me.ConexaoBD
    If Not lobjSeguro.mfExcluir(lobjstcSeguro, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjSeguro.UltimoErro
        GoTo ControleErro
    End If
    Set Me.ConexaoBD = lobjSeguro.ConexaoBD
    Set lobjSeguro = Nothing
    Set lobjstcSeguro = Nothing

    'Excluir inspeção
    Set lobjInsp = New clsA46V717A
    Set lobjInsp.M46V999 = M46V999

    Set lobjInsp.ConexaoBD = Me.ConexaoBD
    pobjstc.objstcPedInsp.nosso_numero = pobjstc.nosso_numero
    pobjstc.objstcPedInsp.Num_Item = pobjstc.Num_Item
    pobjstc.objstcPedInsp.NUM_PI = pobjstc.NUM_PI
    Set pobjstc.objstcPedInsp.ObjPai = pobjstc
    If Not lobjInsp.mfExcluir(pobjstc.objstcPedInsp, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjInsp.UltimoErro
        GoTo ControleErro
    End If
    Set Me.ConexaoBD = lobjInsp.ConexaoBD

    'Exclui Fatores de Proteção
    Set lobjFator = New clsA46V720A
    Set lobjFator.M46V999 = M46V999

    Set lobjstcFator = New stcA46V720B
    lobjstcFator.nosso_numero = pobjstc.nosso_numero
    lobjstcFator.Num_Item = pobjstc.Num_Item
    Set lobjstcFator.ObjPai = pobjstc
    Set lobjFator.ConexaoBD = Me.ConexaoBD
    If Not lobjFator.mfExcluir(lobjstcFator, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjFator.UltimoErro
        GoTo ControleErro
    End If
    Set Me.ConexaoBD = lobjFator.ConexaoBD
    Set lobjFator = Nothing
    Set lobjstcFator = Nothing

    'Exclui LMI
    If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda And M46V999.gTipP00MPED = 2 Then
        Set lobjLMI = New clsA46V728A
        Set lobjLMI.M46V999 = M46V999

        Set lobjstcLMI = New stcA46V728B
        lobjstcLMI.nosso_numero = pobjstc.nosso_numero
        lobjstcLMI.Num_Item = pobjstc.Num_Item
        Set lobjstcLMI.ObjPai = pobjstc
        Set lobjLMI.ConexaoBD = Me.ConexaoBD
        If Not lobjLMI.mfExcluir(lobjstcLMI, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjLMI.UltimoErro
            GoTo ControleErro
        End If
        Set Me.ConexaoBD = lobjLMI.ConexaoBD
        Set lobjLMI = Nothing
        Set lobjstcLMI = Nothing
    End If

    'Exclui Pessoas: Inquilino, Proprietario, Beneficiario.
    Set lobjPessoa = New clsA46V735A
    Set lobjPessoa.M46V999 = M46V999

    Set lobjstcPessoa = New stcA46V735B

    lobjstcPessoa.nosso_numero = pobjstc.nosso_numero
    lobjstcPessoa.Num_Item = pobjstc.Num_Item
    lobjstcPessoa.Num_Apol = pobjstc.Num_Apol
    lobjstcPessoa.Num_Endo = pobjstc.Num_Endo

    Set lobjstcPessoa.ObjPai = pobjstc

    Set lobjPessoa.ConexaoBD = Me.ConexaoBD
    If Not lobjPessoa.mfExcluir(lobjstcPessoa, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjPessoa.UltimoErro
        GoTo ControleErro
    End If

    Set Me.ConexaoBD = lobjPessoa.ConexaoBD
    If lobjPessoa.Colecao.Count > 0 Then
        Set pobjstc.objstcPedInquilino = lobjPessoa.Colecao(1)
    End If
    Set lobjPessoa = Nothing
    Set lobjstcPessoa = Nothing

    'Exclui Textos.
    Set lobjTexto = New clsA46V739A
    Set lobjTexto.M46V999 = M46V999

    Set lobjstcTexto = New stcA46V739B

    lobjstcTexto.nosso_numero = pobjstc.nosso_numero
    lobjstcTexto.Num_Item = pobjstc.Num_Item
    lobjstcTexto.Num_Apol = pobjstc.Num_Apol
    lobjstcTexto.Num_Endo = pobjstc.Num_Endo

    Set lobjstcTexto.ObjPai = pobjstc

    Set lobjTexto.ConexaoBD = Me.ConexaoBD
    If Not lobjTexto.mfExcluir(lobjstcTexto, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjTexto.UltimoErro
        GoTo ControleErro
    End If

    Set Me.ConexaoBD = lobjTexto.ConexaoBD
    If lobjTexto.Colecao.Count > 0 Then
        Set pobjstc.ColecaoTextos = lobjTexto
    End If
    Set lobjTexto = Nothing
    Set lobjstcTexto = Nothing

    'Exclui All Risks.
    Set lobjAllRisks = New clsA46V747A
    Set lobjAllRisks.M46V999 = M46V999

    Set lobjstcAllRisks = New stcA46V747B
    Set lobjstcAllRisks.M46V999 = M46V999

    lobjstcAllRisks.nosso_numero = pobjstc.nosso_numero
    lobjstcAllRisks.Num_Item = pobjstc.Num_Item
    lobjstcAllRisks.Num_Apol = pobjstc.Num_Apol
    lobjstcAllRisks.Num_Endo = pobjstc.Num_Endo
    Set lobjstcAllRisks.ObjPai = pobjstc
    Set lobjAllRisks.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjAllRisks.mfExcluir(lobjstcAllRisks, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjAllRisks.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjAllRisks.ConexaoBD
    Set lobjAllRisks = Nothing
    Set lobjstcAllRisks = Nothing
    
    'Exclui Obras Arte.
    Set lobjObraArte = New clsA46V748A
    Set lobjObraArte.M46V999 = M46V999

    Set lobjstcObraArte = New stcA46V748B
    Set lobjstcObraArte.M46V999 = M46V999

    lobjstcObraArte.nosso_numero = pobjstc.nosso_numero
    lobjstcObraArte.Num_Item = pobjstc.Num_Item
    lobjstcObraArte.Num_Apol = pobjstc.Num_Apol
    lobjstcObraArte.Num_Endo = pobjstc.Num_Endo
    Set lobjstcObraArte.ObjPai = pobjstc
    Set lobjObraArte.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjObraArte.mfExcluir(lobjstcObraArte, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjObraArte.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjObraArte.ConexaoBD
    Set lobjObraArte = Nothing
    Set lobjstcObraArte = Nothing
    

    mfExcluirDependentes = True
SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V708A.mfExcluirDependentes ERRO - " & Err & " " & Err.Description)
    Else
        Call M46V999.gpGraLog(0, "A46V708A.mfExcluirDependentes ERRO - " & UltimoErro)
    End If
    GoTo SAIDA
FechaObj:
    Set lobjCobertura = Nothing
    Set lobjstcCobertura = Nothing
    Set lobjClausula = Nothing
    Set lobjStcClausula = Nothing
    Set lobjObservacao = Nothing
    Set lobjBem = Nothing
    Set lobjstcBem = Nothing
    Set lobjLocBicicleta = Nothing
    Set lobjstcLocBicicleta = Nothing
    Set lobjSeguro = Nothing
    Set lobjstcSeguro = Nothing
    Set lobjInsp = Nothing
    Set lobjFator = Nothing
    Set lobjstcFator = Nothing
    Set lobjLMI = Nothing
    Set lobjstcLMI = Nothing
    Set lobjPessoa = Nothing
    Set lobjstcPessoa = Nothing
    Set lobjTexto = Nothing
    Set lobjstcTexto = Nothing
    Set lobjAllRisks = Nothing
    Set lobjstcAllRisks = Nothing
    Set lobjObraArte = Nothing
    Set lobjstcObraArte = Nothing
    Return
End Function

Private Sub Class_Initialize()
    Set Colecao = New Collection
    Set mClsM46V999 = Nothing
End Sub
Private Sub Class_Terminate()
    Set Colecao = Nothing
    Set mClsM46V999 = Nothing
End Sub

Private Function mfConectar() As Boolean
    mfConectar = False
    If ConexaoBD Is Nothing Then
        If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, ConexaoBD, _
                                 UltimoErro) = False Then Exit Function
    ElseIf ConexaoBD.State <> 1 Then
        If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, ConexaoBD, _
                                 UltimoErro) = False Then Exit Function
    End If
    mfConectar = True
End Function

'#####PROJETO INSPECAO
Public Function mfAlterarAgendaInspecao(ByRef pobjstc As stcA46V708B, _
                                        ByVal pPropTabela As String, _
                                        ByVal pNomeTabImport As String) As Boolean

    Dim lUpdate As String

    On Error GoTo ControleErros

    If Not mfConectar Then GoTo ControleErros

    lUpdate = " Update " & pPropTabela & pNomeTabImport & "Tab_Ped_loc set Insp_Agend  = 1"
    lUpdate = lUpdate & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    '1100008 - Inspeção por Lote
    lUpdate = M46V999.gfWhereAnd(lUpdate) & " NUM_ITEM = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErros

    mfAlterarAgendaInspecao = True

    Exit Function
ControleErros:
    mfAlterarAgendaInspecao = False
    UltimoErro = "A46V708A.mfAlterarAgendaInspecao ERRO - " & Err & " " & Err.Description & _
                 vbCrLf & " SQL: " & lUpdate
End Function

Private Function gInd_CotacaoSolic(ByVal pPropTabela As String, ByVal pNosso_Numero As String) As Boolean
    Dim lobjstcA28V721B As stcA28V721B
    Dim lobjA28V721A As clsA28V721A
    
    gInd_CotacaoSolic = False
    
    If UCase(pPropTabela) Like "SYASPORTAL*" Then
        Set lobjstcA28V721B = New stcA28V721B
        Set lobjA28V721A = New clsA28V721A
        Set lobjA28V721A.M46V999 = M46V999
        
        lobjstcA28V721B.nosso_numero = pNosso_Numero
        If Not lobjA28V721A.gfCarregarColecao(lobjstcA28V721B, pPropTabela) Then
            Call M46V999.gpGraLog(1, lobjA28V721A.UltimoErro)
            Exit Function
        Else
            If lobjA28V721A.Colecao.Count = 1 Then
                gInd_CotacaoSolic = True
            End If
        End If
    End If

SAIDA:
    Set lobjstcA28V721B = Nothing
    Set lobjA28V721A = Nothing
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
