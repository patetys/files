VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V702A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V702A --> Tabela Tab_Ped
' Data/Hora: 31/08/2005 - 11:51:10
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================

' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conex�o com o Banco de Dados
Public Colecao As Collection    ' Cole��o de objetos do tipo stcA46V702B
Public UltimoErro As String           ' Descri��o do �ltimo erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V702A"
End Property
Public Function mfRetornaInicioVigencia(ByRef pobjstc As stcA46V702B) As Long
    Dim lSql As String
    Dim lrst As clsYasRecordSet
    Dim lbdP00APOL As ADODB.Connection
    Dim lNomeTab As String

    If M46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
    Else
        lNomeTab = "GED.DBO.TAB_PED"
    End If

    If M46V999.gfAbrBasDados("P00APOL", M46V999.gTipP00APOL, M46V999.gEndP00APOL, lbdP00APOL, UltimoErro) Then
        lSql = " Select Dat_Ini_Vig from " & lNomeTab & " where num_apol = " & pobjstc.Num_Apol_An
        lSql = lSql & " order by Dat_Emis, dat_ter_vig desc "
        If M46V999.gfObtRegistro(lbdP00APOL, lSql, lrst, UltimoErro) Then
            If Not lrst.EOF Then
                mfRetornaInicioVigencia = lrst("Dat_Ini_Vig")
            Else
                mfRetornaInicioVigencia = 0
                Call M46V999.gpGraLog(0, "O registro de Bonus em 'mfRetornaBonusApolice' para ap�lice: " & pobjstc.Num_Apol_An & " n�o foi encontrado-" & UltimoErro)
            End If
        End If
        Call M46V999.gpFecha2(lbdP00APOL)
    End If
    Set lrst = Nothing
End Function
Public Function mfRetornaTerminoVigencia(ByRef pobjstc As stcA46V702B) As Long
    Dim lSql As String
    Dim lrst As clsYasRecordSet
    Dim lbdP00APOL As ADODB.Connection
    Dim lNomeTab As String

    If pobjstc.Num_Apol_An = 0 Then
        mfRetornaTerminoVigencia = 0
        Exit Function
    End If

'    If M46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
'    Else
'        lNomeTab = "GED.DBO.TAB_PED"
'    End If

    If M46V999.gfAbrBasDados("P00APOL", M46V999.gTipP00APOL, M46V999.gEndP00APOL, lbdP00APOL, UltimoErro) Then
        lSql = " Select Dat_Ter_Vig from " & lNomeTab & " (nolock) where num_apol = " & pobjstc.Num_Apol_An

        If M46V999.gfObtRegistro(lbdP00APOL, lSql, lrst, UltimoErro) Then
            If Not lrst.EOF Then
                mfRetornaTerminoVigencia = lrst("Dat_Ter_Vig")
            Else
                mfRetornaTerminoVigencia = 0
                Call M46V999.gpGraLog(0, "O registro de Bonus em 'mfRetornaBonusApolice' para ap�lice: " & pobjstc.Num_Apol_An & " n�o foi encontrado-" & UltimoErro)
            End If
        End If
        Call M46V999.gpFecha2(lbdP00APOL)
    End If
    Set lrst = Nothing
End Function

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V702B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String

    Dim lInsert As String                'Query.

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped ( "
    lInsert = lInsert & " Nosso_Numero,"
    lInsert = lInsert & " Cod_Unicidade,"
    lInsert = lInsert & " Num_Ped,"
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " Cod_Org,"
    lInsert = lInsert & " Num_Apol,"
    lInsert = lInsert & " Num_Endo,"
    lInsert = lInsert & " Num_Apol_An,"
    lInsert = lInsert & " Cod_Reg,"
    lInsert = lInsert & " Cod_Bco_Cobr,"
    lInsert = lInsert & " Cod_Agen_Angar,"
    lInsert = lInsert & " Cod_Agen_Indicr,"
    lInsert = lInsert & " Num_Cta_Indicr,"
    lInsert = lInsert & " Cod_Unid_Prod,"
    lInsert = lInsert & " Cod_Prod,"
    lInsert = lInsert & " Cod_Agen_Bas_Cor,"
    lInsert = lInsert & " Num_Cta_Bas_Cor,"
    lInsert = lInsert & " Cod_Segurado,"
    lInsert = lInsert & " Nom_Segurado,"
    lInsert = lInsert & " Dsc_Compl_Segurado,"
    lInsert = lInsert & " Num_CGC_CPF,"
    lInsert = lInsert & " Cod_Tip_Logr,"
    lInsert = lInsert & " Nom_Logr,"
    lInsert = lInsert & " Dsc_Compl_Logr,"
    lInsert = lInsert & " Nom_Bairro,"
    lInsert = lInsert & " Nom_Cid,"
    lInsert = lInsert & " Sig_UF,"
    lInsert = lInsert & " Num_CEP,"
    lInsert = lInsert & " Num_Fone,"
    lInsert = lInsert & " Cod_Estip,"
    lInsert = lInsert & " Val_Coef_Cms_Pad,"
    lInsert = lInsert & " Dat_Ini_Vig,"
    lInsert = lInsert & " Dat_Ter_Vig,"
    lInsert = lInsert & " Dat_Emis,"
    lInsert = lInsert & " Cod_Moed,"
    lInsert = lInsert & " Per_Juro_Frac,"
    lInsert = lInsert & " Qtd_Parc,"
    lInsert = lInsert & " Qtd_Anexo,"
    lInsert = lInsert & " Cod_Iden_Praz,"
    lInsert = lInsert & " Cod_Pr_Min_Endo,"
    lInsert = lInsert & " Cod_Pr_Min_Parc,"
    lInsert = lInsert & " Val_Pr_Antec,"
    lInsert = lInsert & " Per_Adic_Form_Pag,"
    lInsert = lInsert & " Per_Desc_Form_Pag,"
    lInsert = lInsert & " Val_IS,"
    lInsert = lInsert & " Val_IS_2,"
    lInsert = lInsert & " Val_PR,"
    lInsert = lInsert & " Val_Pr_Ref_Tarif,"
    lInsert = lInsert & " Val_Pr_Ref_Resg,"
    lInsert = lInsert & " Dat_Ini_Viag,"
    lInsert = lInsert & " Cod_Tip_Liber,"
    lInsert = lInsert & " Cod_Iden_Alt_Corr,"
    lInsert = lInsert & " Cod_User_Digit,"
    lInsert = lInsert & " Dat_Digit,"
    lInsert = lInsert & " Hor_Digit,"
    lInsert = lInsert & " Cod_User_Liber,"
    lInsert = lInsert & " Dat_Liber,"
    lInsert = lInsert & " Hor_Liber,"
    lInsert = lInsert & " Cod_Sit_Ped,"
    lInsert = lInsert & " Cod_Tip_Doc,"
    lInsert = lInsert & " Cod_Dizer,"
    lInsert = lInsert & " Qtd_Copia,"
    lInsert = lInsert & " Val_Cust_Emis,"
    lInsert = lInsert & " Val_Juro,"
    lInsert = lInsert & " Val_IOF,"
    lInsert = lInsert & " Num_Prop,"
    lInsert = lInsert & " Cod_Posi_Atual,"
    lInsert = lInsert & " Dat_Mov_Ini,"
    lInsert = lInsert & " Dat_Mov_Ter,"
    lInsert = lInsert & " Tip_Cust_Emis,"
    lInsert = lInsert & " Val_Cambio_Emis,"
    lInsert = lInsert & " Iden,"
    lInsert = lInsert & " Num_PI,"
    lInsert = lInsert & " Cod_Produto,"
    lInsert = lInsert & " Cod_Protocolo ,"
    lInsert = lInsert & " Cod_Sequencia,"
    lInsert = lInsert & " Tip_Pessoa,"
    lInsert = lInsert & " Cod_Ramo_Ativ,"
    lInsert = lInsert & " Cod_Doc,"
    lInsert = lInsert & " Num_Doc,"
    lInsert = lInsert & " Org_Exp,"
    lInsert = lInsert & " Dat_Emis_Doc,"
    lInsert = lInsert & " Nom_Empresa,"
    lInsert = lInsert & " Sexo,"
    lInsert = lInsert & " Dat_Pri_Calculo,"
    lInsert = lInsert & " Ult_Tarifa,"
    lInsert = lInsert & " Cod_Dig_Agen_Bas_Cor,"
    lInsert = lInsert & " Cod_Ver_Unici"
    'sandra - comentado para passar vers�o
    lInsert = lInsert & ",  Per_Desc_LMI,"
    lInsert = lInsert & " LMI_Unico,"
    lInsert = lInsert & " Num_Cotac_LMI"
    If (M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Or pNomeTabImport = "FTP_") And _
       (Not (UCase(pPropTabela) = "RAMOSDIVERSOS.DBO." And M46V999.gfPreenchido(pNomeTabImport) = False)) Then
        If (pNomeTabImport = "FTP_" Or UCase(pPropTabela) = "SYASPORTAL.DBO.") Then
            lInsert = lInsert & " , Val_RendaFatu "
        End If
    End If
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " , Tip_Transmissao "
    End If
    lInsert = lInsert & ", Ind_CondGerais"
    lInsert = lInsert & ", Ind_Reentrada"
    lInsert = lInsert & ", Nosso_Numero_Original"
    If M46V999.gPrograma = "D046O003" And Not UCase(pPropTabela) = "RAMOSDIVERSOS.DBO." Then
        lInsert = lInsert & ", Cod_Segmentacao"
    End If
    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol_An) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Reg) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Bco_Cobr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Agen_Angar) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Agen_Indicr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cta_Indicr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unid_Prod) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Prod) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Agen_Bas_Cor) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cta_Bas_Cor) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Segurado) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Segurado) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Segurado) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cgc_Cpf) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Logr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Logr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Logr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Bairro) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Cid) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Sig_UF) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_CEP) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Fone) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Estip) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Coef_Cms_Pad) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ini_Vig) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ter_Vig) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Emis) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Moed) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Juro_Frac) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Parc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Anexo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Praz) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Pr_Min_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Pr_Min_Parc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Antec) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Adic_Form_Pag) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Form_Pag) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS_2) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.VAL_PR) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Ref_Tarif) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Ref_Resg) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ini_Viag) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Liber) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Alt_Corr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Digit) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.DAT_DIGIT) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.HOR_DIGIT) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Liber) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Liber) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Hor_Liber) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Sit_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Doc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Dizer) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Copia) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Cust_Emis) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Juro) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_IOF) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Prop) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Posi_Atual) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Mov_Ini) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Mov_Ter) & ","
    If pobjstc.Tip_Cust_Emis = 0 Then
        lInsert = lInsert & " 1 ,"
    Else
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Cust_Emis) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Cambio_Emis) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Iden) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Produto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Protocolo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Sequencia) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Pessoa) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo_Ativ) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Doc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Doc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Org_Exp) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Emis_Doc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Empresa) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Sexo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Pri_Calculo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ult_Tarifa) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Dig_Agen_Bas_Cor) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ver_Unici)
    'sandra - comentado para passar vers�o
    lInsert = lInsert & ", " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_LMI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.LMI_Unico) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cotac_LMI)

    If (M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Or pNomeTabImport = "FTP_") And _
       (Not (UCase(pPropTabela) = "RAMOSDIVERSOS.DBO." And M46V999.gfPreenchido(pNomeTabImport) = False)) Then
       If (pNomeTabImport = "FTP_" Or UCase(pPropTabela) = "SYASPORTAL.DBO.") Then
            lInsert = lInsert & " , " & M46V999.gfFormatarCamposSQL(pobjstc.Val_RendaFatu)
        End If
    End If
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " , " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Transmissao)
    End If
    lInsert = lInsert & ", " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_CondGerais)
    lInsert = lInsert & ", " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Reentrada)
    lInsert = lInsert & ", " & M46V999.gfFormatarCamposSQL(pobjstc.Nosso_Numero_Original)
    If M46V999.gPrograma = "D046O003" And Not UCase(pPropTabela) = "RAMOSDIVERSOS.DBO." Then
        lInsert = lInsert & ", " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Segmentacao)
    End If
    lInsert = lInsert & ")"

    mfFormatarInclusaoSQL = lInsert

End Function


Private Function mfFormatarAlteracaoSQL(ByRef pobjstc As stcA46V702B, _
                                        ByVal pPropTabela As String, _
                                        ByVal pNomeTabImport As String, _
                                        Optional ByVal pDataUltAlt As Double, _
                                        Optional ByVal pUltUsuAlt As String) As String

    Dim lstrTMP As String                'Query.

    lstrTMP = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ped Set "
    lstrTMP = lstrTMP & " Cod_Ramo = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lstrTMP = lstrTMP & " Cod_Org = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lstrTMP = lstrTMP & " Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lstrTMP = lstrTMP & " Num_Endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lstrTMP = lstrTMP & " Num_Apol_An = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol_An) & ","
    lstrTMP = lstrTMP & " Cod_Reg = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Reg) & ","
    lstrTMP = lstrTMP & " Cod_Bco_Cobr = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Bco_Cobr) & ","
    lstrTMP = lstrTMP & " Cod_Agen_Angar = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Agen_Angar) & ","
    lstrTMP = lstrTMP & " Cod_Agen_Indicr = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Agen_Indicr) & ","
    lstrTMP = lstrTMP & " Num_Cta_Indicr = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cta_Indicr) & ","
    lstrTMP = lstrTMP & " Cod_Unid_Prod = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unid_Prod) & ","
    lstrTMP = lstrTMP & " Cod_Prod = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Prod) & ","
    lstrTMP = lstrTMP & " Cod_Agen_Bas_Cor = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Agen_Bas_Cor) & ","
    lstrTMP = lstrTMP & " Num_Cta_Bas_Cor = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cta_Bas_Cor) & ","
    lstrTMP = lstrTMP & " Cod_Segurado = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Segurado) & ","
    lstrTMP = lstrTMP & " Nom_Segurado = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Segurado) & ","
    lstrTMP = lstrTMP & " Dsc_Compl_Segurado = " & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Segurado) & ","
    lstrTMP = lstrTMP & " Num_CGC_CPF = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cgc_Cpf) & ","
    lstrTMP = lstrTMP & " Cod_Tip_Logr = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Logr) & ","
    lstrTMP = lstrTMP & " Nom_Logr = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Logr) & ","
    lstrTMP = lstrTMP & " Dsc_Compl_Logr = " & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Logr) & ","
    lstrTMP = lstrTMP & " Nom_Bairro = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Bairro) & ","
    lstrTMP = lstrTMP & " Nom_Cid = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Cid) & ","
    lstrTMP = lstrTMP & " Sig_UF = " & M46V999.gfFormatarCamposSQL(pobjstc.Sig_UF) & ","
    lstrTMP = lstrTMP & " Num_CEP = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_CEP) & ","
    lstrTMP = lstrTMP & " Num_Fone = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Fone) & ","
    lstrTMP = lstrTMP & " Cod_Estip = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Estip) & ","
    lstrTMP = lstrTMP & " Val_Coef_Cms_Pad = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Coef_Cms_Pad) & ","
    lstrTMP = lstrTMP & " Dat_Ini_Vig = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ini_Vig) & ","
    lstrTMP = lstrTMP & " Dat_Ter_Vig = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ter_Vig) & ","
    lstrTMP = lstrTMP & " Dat_Emis = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Emis) & ","
    lstrTMP = lstrTMP & " Cod_Moed = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Moed) & ","
    lstrTMP = lstrTMP & " Per_Juro_Frac = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Juro_Frac) & ","
    lstrTMP = lstrTMP & " Qtd_Parc = " & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Parc) & ","
    lstrTMP = lstrTMP & " Qtd_Anexo = " & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Anexo) & ","
    lstrTMP = lstrTMP & " Cod_Iden_Praz = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Praz) & ","
    lstrTMP = lstrTMP & " Cod_Pr_Min_Endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Pr_Min_Endo) & ","
    lstrTMP = lstrTMP & " Cod_Pr_Min_Parc = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Pr_Min_Parc) & ","
    lstrTMP = lstrTMP & " Val_Pr_Antec = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Antec) & ","
    lstrTMP = lstrTMP & " Per_Adic_Form_Pag = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Adic_Form_Pag) & ","
    lstrTMP = lstrTMP & " Per_Desc_Form_Pag = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Form_Pag) & ","
    lstrTMP = lstrTMP & " Val_IS = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS) & ","
    lstrTMP = lstrTMP & " Val_IS_2 = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS_2) & ","
    lstrTMP = lstrTMP & " Val_PR = " & M46V999.gfFormatarCamposSQL(pobjstc.VAL_PR) & ","
    lstrTMP = lstrTMP & " Val_Pr_Ref_Tarif = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Ref_Tarif) & ","
    lstrTMP = lstrTMP & " Val_Pr_Ref_Resg = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Ref_Resg) & ","
    lstrTMP = lstrTMP & " Dat_Ini_Viag = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ini_Viag) & ","
    lstrTMP = lstrTMP & " Cod_Tip_Liber = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Liber) & ","
    lstrTMP = lstrTMP & " Cod_Iden_Alt_Corr = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Alt_Corr) & ","
    lstrTMP = lstrTMP & " Cod_User_Digit = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Digit) & ","
    lstrTMP = lstrTMP & " Dat_Digit = " & M46V999.gfFormatarCamposSQL(pobjstc.DAT_DIGIT) & ","
    lstrTMP = lstrTMP & " Hor_Digit = " & M46V999.gfFormatarCamposSQL(pobjstc.HOR_DIGIT) & ","
    lstrTMP = lstrTMP & " Cod_User_Liber = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Liber) & ","
    lstrTMP = lstrTMP & " Dat_Liber = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Liber) & ","
    lstrTMP = lstrTMP & " Hor_Liber = " & M46V999.gfFormatarCamposSQL(pobjstc.Hor_Liber) & ","
    lstrTMP = lstrTMP & " Cod_Sit_Ped = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Sit_Ped) & ","
    lstrTMP = lstrTMP & " Cod_Tip_Doc = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Doc) & ","
    lstrTMP = lstrTMP & " Cod_Dizer = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Dizer) & ","
    lstrTMP = lstrTMP & " Qtd_Copia = " & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Copia) & ","
    lstrTMP = lstrTMP & " Val_Cust_Emis = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Cust_Emis) & ","
    lstrTMP = lstrTMP & " Val_Juro = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Juro) & ","
    lstrTMP = lstrTMP & " Val_IOF = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_IOF) & ","
    lstrTMP = lstrTMP & " Num_Prop = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Prop) & ","
    lstrTMP = lstrTMP & " Cod_Posi_Atual = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Posi_Atual) & ","
    lstrTMP = lstrTMP & " Dat_Mov_Ini = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Mov_Ini) & ","
    lstrTMP = lstrTMP & " Dat_Mov_Ter = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Mov_Ter) & ","
    lstrTMP = lstrTMP & " Tip_Cust_Emis = " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Cust_Emis) & ","
    lstrTMP = lstrTMP & " Val_Cambio_Emis = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Cambio_Emis) & ","
    lstrTMP = lstrTMP & " Iden = " & M46V999.gfFormatarCamposSQL(pobjstc.Iden) & ","
    lstrTMP = lstrTMP & " Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lstrTMP = lstrTMP & " Cod_Produto = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Produto) & ","
    lstrTMP = lstrTMP & " Cod_Sequencia = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Sequencia) & ","
    lstrTMP = lstrTMP & " Cod_Protocolo = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Protocolo) & ","
    lstrTMP = lstrTMP & " Tip_Pessoa= " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Pessoa) & ","
    lstrTMP = lstrTMP & " Cod_Ramo_Ativ= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo_Ativ) & ","
    lstrTMP = lstrTMP & " Cod_Doc= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Doc) & ","
    lstrTMP = lstrTMP & " Num_Doc= " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Doc) & ","
    lstrTMP = lstrTMP & " Org_Exp= " & M46V999.gfFormatarCamposSQL(pobjstc.Org_Exp) & ","
    lstrTMP = lstrTMP & " Dat_Emis_Doc= " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Emis_Doc) & ","
    lstrTMP = lstrTMP & " Nom_Empresa= " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Empresa) & ","
    lstrTMP = lstrTMP & " Sexo= " & M46V999.gfFormatarCamposSQL(pobjstc.Sexo) & ","
    lstrTMP = lstrTMP & " Dat_Pri_Calculo = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Pri_Calculo) & ","
    lstrTMP = lstrTMP & " Ult_Tarifa = " & M46V999.gfFormatarCamposSQL(pobjstc.Ult_Tarifa) & ","
    lstrTMP = lstrTMP & " Cod_Unicidade = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    lstrTMP = lstrTMP & " Cod_Dig_Agen_Bas_Cor = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Dig_Agen_Bas_Cor) & ","
    lstrTMP = lstrTMP & " Cod_Ver_Unici = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ver_Unici)
    lstrTMP = lstrTMP & ", Per_Desc_LMI = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_LMI) & ","
    lstrTMP = lstrTMP & " LMI_Unico = " & M46V999.gfFormatarCamposSQL(pobjstc.LMI_Unico) & ","
    lstrTMP = lstrTMP & " Num_Cotac_LMI = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cotac_LMI)

    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        lstrTMP = lstrTMP & " , Val_RendaFatu = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_RendaFatu)
    End If
    lstrTMP = lstrTMP & " , Ind_CondGerais = " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_CondGerais)

    lstrTMP = lstrTMP & " WHERE "
    lstrTMP = lstrTMP & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    lstrTMP = lstrTMP & " and Dat_Liber = " & pDataUltAlt
    lstrTMP = lstrTMP & " and Cod_User_Liber = " & M46V999.gfFormatarCamposSQL(pUltUsuAlt)

    mfFormatarAlteracaoSQL = lstrTMP
End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V702B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, tagDynaTrace As String) As String

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfFormatarExclusaoSQL"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lstrTMP As String                'Query.

    lstrTMP = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped WHERE "
    lstrTMP = lstrTMP & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarExclusaoSQL = lstrTMP
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V702B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lstrTMP As String                'Query.
    Dim lDataValida As Long

    lstrTMP = "SELECT * FROM " & pPropTabela & pNomeTabImport & "Tab_Ped with(nolock) "

    'Nosso N�mero
    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lstrTMP = lstrTMP & " where Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If
    If (pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_") And pobjstc.Cod_Unicidade > -1 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    'N�mero do pedido
    If pobjstc.Num_Ped <> 0 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Num_Ped= " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped)
    End If

    'N�mero do PI
    If pobjstc.NUM_PI <> 0 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)
    End If

    'Situa��o do pedido
    If pobjstc.Cod_Sit_Ped = e_SitPedido_Nenhum Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Cod_Sit_Ped <> 3 "
    ElseIf pobjstc.Cod_Sit_Ped <> 0 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Cod_Sit_Ped = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Sit_Ped)
    End If

    'Tipo de documento
    If pobjstc.Cod_Tip_Doc <> 0 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Cod_Tip_Doc = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Doc)
    End If
    'Nome do segurado.
    If M46V999.gfPreenchido(pobjstc.Nom_Segurado) Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Nom_Segurado like '%" & UCase(pobjstc.Nom_Segurado) & "%'"
    End If

    'In�cio da vig�ncia.
    If pobjstc.Dat_Ini_Vig <> 0 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Dat_Ini_Vig = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ini_Vig)
    End If

    'T�rmino da vig�ncia.
    If pobjstc.Dat_Ter_Vig <> 0 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Dat_Ter_Vig = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ter_Vig)
    End If

    'Produtor.
    If pobjstc.Cod_Prod <> 0 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Cod_Prod = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Prod)
    End If

    'Ramo
    If pobjstc.Cod_Ramo <> e_CodRamo_NaoDefinido Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Cod_Ramo = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo)
    End If


    If M46V999.gfPreenchido(pobjstc.NossoNumeroIn) Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Nosso_Numero in ( " & pobjstc.NossoNumeroIn & ")"
    End If


    If pobjstc.Cod_Segurado <> 0 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Cod_Segurado = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Segurado)
    End If

    'Ap�lice pendentes de emiss�o
    If pobjstc.Num_Apol = 9999999999# Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Num_Apol  = 0 "
        'If Not M46V999.gBaseLocal Then
        '    lstrTMP = lstrTMP & " and " & pNomeTabImport & "Tab_Ped.NUM_PI IN (SELECT NUM_PI "
        '    lstrTMP = lstrTMP & " From tab_ctrl_emis"
        '    lstrTMP = lstrTMP & " WHERE  cod_iden_stat_acao not in (70,80,81,90 )) "
        'End If
    ElseIf pobjstc.Num_Apol <> 0 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol)
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " num_endo= " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo)
    End If

    If M46V999.gAgendarCot = True Then
        lDataValida = M46V999.gfForData(1, Date - 35)
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Dat_Pri_Calculo = 0 or Dat_Pri_Calculo >= " & lDataValida
    End If

    lstrTMP = lstrTMP & " ORDER BY Num_Ped desc"

    mfFormatarConsultaSQL = lstrTMP
End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V702B, ByRef prst As clsYasRecordSet, ByVal pNomeTabImport As String)
    Dim lAlter As String
    Dim lTemp As String
    On Error Resume Next
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Cod_Unicidade = M46V999.gfTratarNulo(prst("Cod_Unicidade"), tpoNumerico)
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
    pobjstc.Cod_Produto = M46V999.gfTratarNulo(prst("Cod_Produto"), tpoNumerico)
    pobjstc.Cod_Sequencia = M46V999.gfTratarNulo(prst("Cod_Sequencia"), tpoNumerico)
    pobjstc.Cod_Protocolo = M46V999.gfTratarNulo(prst("Cod_Protocolo"), tpoNumerico)
    pobjstc.Tip_Pessoa = M46V999.gfTratarNulo(prst("Tip_Pessoa"), tpoNumerico)
    pobjstc.Cod_Ramo_Ativ = M46V999.gfTratarNulo(prst("Cod_Ramo_Ativ"), tpoNumerico)
    pobjstc.Cod_Doc = M46V999.gfTratarNulo(prst("Cod_Doc"), tpoNumerico)
    pobjstc.Num_Doc = M46V999.gfTratarNulo(prst("Num_Doc"), tpoString)
    pobjstc.Org_Exp = M46V999.gfTratarNulo(prst("Org_Exp"), tpoString)
    pobjstc.Dat_Emis_Doc = M46V999.gfTratarNulo(prst("Dat_Emis_Doc"), tpoNumerico)
    pobjstc.Nom_Empresa = M46V999.gfTratarNulo(prst("Nom_Empresa"), tpoString)
    pobjstc.Sexo = M46V999.gfTratarNulo(prst("Sexo"), tpoString)
    pobjstc.Dat_Pri_Calculo = M46V999.gfTratarNulo(prst("Dat_Pri_Calculo"), tpoNumerico)
    pobjstc.Ult_Tarifa = M46V999.gfTratarNulo(prst("Ult_Tarifa"), tpoNumerico)


    lTemp = IsNull(prst("Cod_Dig_Agen_Bas_Cor"))
    If Err.Number = 3265 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.gPrograma <> "D046O003" Then
        'Crio o campo Cod_Dig_Agen_Bas_Cor na tabela TAB_PED
        lAlter = " alter table " & pNomeTabImport & "tab_ped add COLUMN  Cod_Dig_Agen_Bas_Cor int;"
        If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(2, UltimoErro)
            Exit Sub
        End If
        lAlter = " Update " & pNomeTabImport & "tab_ped"
        lAlter = lAlter & " set Cod_dig_Agen_Bas_Cor = right(Cod_Agen_Bas_Cor,1),"
        lAlter = lAlter & " Cod_Agen_Bas_Cor = Left(Cod_Agen_Bas_Cor, Len(Cod_Agen_Bas_Cor) - 1)"
        lAlter = lAlter & " Where Cod_Agen_Bas_Cor <> 0 and Cod_Bco_Cobr = 237"
        If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(2, UltimoErro)
            Exit Sub
        End If
    Else
        pobjstc.Cod_Dig_Agen_Bas_Cor = M46V999.gfTratarNulo(prst("Cod_Dig_Agen_Bas_Cor"), tpoNumerico)
    End If

    On Error Resume Next
    lTemp = IsNull(prst("Cod_Ver_Unici"))
    If Err.Number = 3265 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.gPrograma <> "D046O003" Then
        lAlter = " alter table " & pNomeTabImport & "tab_ped add COLUMN  Cod_Ver_Unici int;"
        If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(2, UltimoErro)
            Exit Sub
        End If
        pobjstc.Cod_Ver_Unici = 0
    Else
        pobjstc.Cod_Ver_Unici = M46V999.gfTratarNulo(prst("Cod_Ver_Unici"), tpoNumerico)
    End If

    On Error Resume Next
    Err.Clear
    lTemp = IsNull(prst("Val_RendaFatu"))
    If Err.Number = 3265 Then
        If M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.gPrograma <> "D046O003" Then
            lAlter = " alter table " & pNomeTabImport & "tab_ped add COLUMN  Val_RendaFatu float;"
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
        pobjstc.Val_RendaFatu = 0
    Else
        pobjstc.Val_RendaFatu = M46V999.gfTratarNulo(prst("Val_RendaFatu"), TpoValor)
    End If

    'sandra - comentado para passar vers�o
    On Error Resume Next
    Err.Clear
    lTemp = IsNull(prst("LMI_Unico"))
    If Err.Number = 3265 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.gPrograma <> "D046O003" Then
        'Crio o campo Per_Desc_LMI na tabela TAB_PED
        lAlter = " Alter table " & pNomeTabImport & "tab_ped add COLUMN  Per_Desc_LMI  float;"
        If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(2, UltimoErro)
            Exit Sub
        End If
        'Crio o campo LMI_Unico na tabela TAB_PED
        lAlter = " Alter table " & pNomeTabImport & "tab_ped add COLUMN  LMI_Unico  int;"
        If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(2, UltimoErro)
            Exit Sub
        End If
    Else
        'sandra - comentado para passar vers�o
        pobjstc.Per_Desc_LMI = M46V999.gfTratarNulo(prst("Per_Desc_LMI"), TpoValor)
        pobjstc.LMI_Unico = M46V999.gfTratarNulo(prst("LMI_Unico"), TpoValor)
    End If

    'sandra - comentado para passar vers�o
    lTemp = IsNull(prst("Num_Cotac_LMI"))
    If Err.Number = 3265 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.gPrograma <> "D046O003" Then
        'Crio o campo Num_Cotac_LMI na tabela TAB_PED
        lAlter = " Alter table " & pNomeTabImport & "tab_ped add COLUMN  Num_Cotac_LMI  varchar(20);"
        If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(2, UltimoErro)
            Exit Sub
        End If
    Else
        'sandra - comentado para passar vers�o
        pobjstc.Num_Cotac_LMI = M46V999.gfTratarNulo(prst("Num_Cotac_LMI"), tpoString)
    End If
    Err.Clear
    lTemp = IsNull(prst("Ind_CondGerais"))
    If Err.Number = 3265 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.gPrograma <> "D046O003" Then
        'Crio o campo Num_Cotac_LMI na tabela TAB_PED
        lAlter = " Alter table " & pNomeTabImport & "tab_ped add COLUMN  Ind_CondGerais int;"
        If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(2, UltimoErro)
            Exit Sub
        End If
    Else
        'sandra - comentado para passar vers�o
        pobjstc.Ind_CondGerais = M46V999.gfTratarNulo(prst("Ind_CondGerais"), tpoNumerico)
    End If
    Err.Clear
    lTemp = IsNull(prst("Ind_Reentrada"))

    If Err.Number = 3265 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.gPrograma <> "D046O003" Then
        If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
            lAlter = " alter table Tab_Ped add COLUMN Ind_Reentrada numeric;"
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
    Else
        pobjstc.Ind_Reentrada = M46V999.gfTratarNulo(prst("Ind_Reentrada"), tpoNumerico)
    End If

    Err.Clear
    lTemp = IsNull(prst("Nosso_Numero_Original"))

    If Err.Number = 3265 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.gPrograma <> "D046O003" Then
        If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
            lAlter = " alter table Tab_Ped add COLUMN Nosso_Numero_Original varchar(20);"
            If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Sub
            End If
        End If
    Else
        pobjstc.Nosso_Numero_Original = M46V999.gfTratarNulo(prst("Nosso_Numero_Original"), tpoNumerico)
    End If
    If pobjstc.Ind_Reentrada <> 1 Then
        pobjstc.Nosso_Numero_Original = ""
    End If

    Err.Clear

    pobjstc.Cod_Segmentacao = M46V999.gfTratarNulo(prst("Cod_Segmentacao"), tpoNumerico)
    If Err.Number = 3265 Then
        pobjstc.Cod_Segmentacao = 0
    End If


End Sub

Public Function mfIncluir(ByRef pobjstc As stcA46V702B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfIncluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lsSQl As String           'Query.

    On Local Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro
Retorno:
    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfIncluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErro:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        If Err.Number = 3022 Then
            UltimoErro = "FTE0003 - Documento " & pobjstc.nosso_numero & " j� existente."
        Else
            UltimoErro = Err.Description
        End If
    Else
        If InStr(1, UCase(UltimoErro), "'COD_DIG_AGEN_BAS_COR'") > 0 Then
            lsSQl = " alter table " & pNomeTabImport & "tab_ped add COLUMN  Cod_Dig_Agen_Bas_Cor int;"
            If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
                Exit Function
            End If
            GoTo Retorno
        End If
        If InStr(1, UCase(UltimoErro), "'COD_VER_UNICI'") > 0 Then
            lsSQl = " alter table " & pNomeTabImport & "tab_ped add COLUMN  Cod_Ver_Unici int;"
            If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
                Exit Function
            End If
            GoTo Retorno
        End If
        If InStr(1, UCase(UltimoErro), "'VAL_RENDAFATU'") > 0 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
            lsSQl = " alter table " & pNomeTabImport & "tab_ped add COLUMN  Val_RendaFatu float;"
            If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
                Exit Function
            End If
            GoTo Retorno
        End If
        'sandra
        If InStr(1, UCase(UltimoErro), "'PER_DESC_LMI'") > 0 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
            lsSQl = " alter table " & pNomeTabImport & "tab_ped add COLUMN  Per_Desc_LMI float;"
            If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
                Exit Function
            End If
            GoTo Retorno
        End If
        If InStr(1, UCase(UltimoErro), "'LMI_UNICO'") > 0 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
            lsSQl = " alter table " & pNomeTabImport & "tab_ped add COLUMN  LMI_Unico int;"
            If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
                Exit Function
            End If
            GoTo Retorno
        End If
        If InStr(1, UCase(UltimoErro), "'NUM_COTAC_LMI'") > 0 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
            lsSQl = " alter table " & pNomeTabImport & "tab_ped add COLUMN  Num_Cotac_LMI varchar(20);"
            If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
                Exit Function
            End If
            GoTo Retorno
        End If

        If InStr(1, UCase(UltimoErro), "'IND_CONDGERAIS'") > 0 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
            lsSQl = " alter table " & pNomeTabImport & "tab_ped add COLUMN  Ind_CondGerais int;"
            If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
                Exit Function
            End If
            GoTo Retorno
        End If
        If InStr(1, UCase(UltimoErro), "'IND_REENTRADA'") > 0 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
            lsSQl = " alter table " & pNomeTabImport & "tab_ped add COLUMN  Ind_Reentrada numeric;"
            If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
                Exit Function
            End If
            GoTo Retorno
        End If
        If InStr(1, UCase(UltimoErro), "'NOSSO_NUMERO_ORIGINAL'") > 0 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
            lsSQl = " alter table " & pNomeTabImport & "tab_ped add COLUMN  Nosso_Numero_Original varchar(20);"
            If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
                Exit Function
            End If
            GoTo Retorno
        End If
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)

End Function
Public Function mfAlterar(ByRef pobjstc As stcA46V702B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                          Optional ByVal pDataUltAlt As Double, _
                          Optional ByVal pUltUsuAlt As String) As Boolean

    Dim lsSQl As String         'Query.
    Dim lRegAlt As Long             'Total de registros alterados.
    Dim lrsPed As clsYasRecordSet    'Registro : TAB_PED.
    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarAlteracaoSQL(pobjstc, pPropTabela, pNomeTabImport, pDataUltAlt, pUltUsuAlt)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro, lRegAlt) <> 0 Then GoTo ControleErro
    If M46V999.gfPreenchido(pUltUsuAlt) = True And lRegAlt = 0 Then
        'Se entrou nesta condi��o � pq o registro foi alterado por outro
        'usu�rio durante o processamento.
        lsSQl = "SELECT * FROM " & pPropTabela & pNomeTabImport & "TAB_PED WITH (NOLOCK) WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
        If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrsPed, UltimoErro) = False Then
            Exit Function
        End If
        If lrsPed.EOF = True Then
            UltimoErro = "FTE2012 - " & M46V999.gPrefixo & "Erro ao atualizar dados comuns do documento " & _
                         pobjstc.nosso_numero & " (registro exclu�do durante processo)."
        Else
            UltimoErro = "FTE2012 - " & M46V999.gPrefixo & "Erro ao atualizar dados comuns do documento " & _
                         pobjstc.nosso_numero & " (registro alterado durante processo. �ltima atualiza��o: " & _
                         "Usu�rio - " & lrsPed("Cod_User_Liber") & ")."
        End If
        Exit Function
    End If

    '[FRH]->
    Call mpLimparColecao
    'If Colecao.Count > 0 Then
    '    Colecao.Remove (pObjStc.Class_ID)
    'End If
    '<-[FRH]

    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfAlterar = True
    Set lrsPed = Nothing
    Exit Function
ControleErro:
    mfAlterar = False
    Set lrsPed = Nothing
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V702A.mfAlterar ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V702A.mfAlterar ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfAlterarReentrada(ByRef pobjstc As stcA46V702B, _
                                   ByVal pPropTabela As String, _
                                   ByVal pIndReentrada As Integer, _
                                   ByVal pNossoNumeroOriginal As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    Dim methodName As String
    methodName = "mfAlterarReentrada"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim sSql As String
    Dim lRegAlt As Long
    Dim lUpdate As String

    On Error GoTo ControleErros

    If Not mfConectar Then GoTo ControleErros

    sSql = "Update " & pPropTabela & "Tab_Ped set Ind_Reentrada = " & pIndReentrada
    sSql = sSql & " , Nosso_Numero_Original = " & M46V999.gfFormatarCamposSQL(pNossoNumeroOriginal)
    sSql = sSql & " where nosso_numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
Retornar:
    If M46V999.gfExeSQL(ConexaoBD, sSql, UltimoErro, lRegAlt) <> 0 Then GoTo ControleErros

    Call M46V999.gpFecha2(ConexaoBD)

    mfAlterarReentrada = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

    Exit Function
ControleErros:
    mfAlterarReentrada = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
    End If
    If UCase(UltimoErro) Like "*'IND_REENTRADA'*" Then
        On Error Resume Next
        If M46V999.gTipUsuario = "2" Then
            lUpdate = " alter table Tab_Ped add COLUMN Ind_Reentrada numeric;"
            If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
                Exit Function
            End If
            GoTo Retornar
        End If
    End If
    If UCase(UltimoErro) Like "*'NOSSO_NUMERO_ORIGINAL'*" Then
        On Error Resume Next
        If M46V999.gTipUsuario = "2" Then
            lUpdate = " alter table Tab_Ped add COLUMN Nosso_Numero_Original varchar(20);"
            If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
                Exit Function
            End If
            GoTo Retornar
        End If
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function

Public Function mfExcluir(ByRef pobjstc As stcA46V702B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    
    methodName = "mfExcluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lsSQl As String           'Query.

    On Error GoTo ControleErro
    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, tagDynaTrace)

    If Not mfConectar Then GoTo ControleErro
    'Exclui os registros dependentes tamb�m.
    If Not mfExcluirDependentes(pobjstc, pPropTabela, pNomeTabImport, tagDynaTrace) Then GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    M46V999.gBufferSql = False
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
        Call M46V999.gpGraLog(0, "A46V702A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V702A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function

Private Function mfExcluirDependentes(ByRef pobjstc As stcA46V702B, _
                                      ByVal pPropTabela As String, _
                                      ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfExcluirDependentes"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lobjEndereco As clsA46V707A    'Objeto da classe clsA46V707A (TAB_PED_ENDER)
    Dim lobjstcEndereco As stcA46V707B    'Objeto da classe de estrutura stcA46V707B (TAB_PED_ENDER)
    Dim lobjCorretor As clsA46V705A    'Objeto da classe clsA46V705A (TAB_PED_CORR)
    Dim lobjstcCorretor As stcA46V705B    'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    Dim lobjCosseguro As clsA46V706A    'Objeto da classe clsA46V706A (TAB_PED_COSS)
    Dim lobjstcCosseguro As stcA46V706B    'Objeto da classe de estrutura stcA46V706B (TAB_PED_COSS)
    Dim lobjJuros As clsA46V711A    'Objeto da classe clsA46V711A (TAB_PED_JUROS)
    Dim lobjstcJuros As stcA46V711B    'Objeto da classe de estrutura stcA46V711B (TAB_PED_JUROS)
    Dim lobjCtrlEmis As clsA46V712A    'Objeto da classe clsA46V712A (TAB_CTRL_EMIS)
    Dim lobjMovAgen As clsA46V743A    'Objeto da classe clsA46V743A (TAB_MOV_AGEN)
    Dim lobjPedAgregados As clsA46v741A    'Objeto da classe clsA46V741A (TAB_PED_AGREGADOS)
    Dim lobjstcCtrlEmis As stcA46V712B    'Objeto da classe de estrutura stcA46V712B (TAB_CTRL_EMIS)
    Dim lobjstcMovAgen As stcA46V743B    'Objeto da classe de estrutura stcA46V743B (TAB_MOV_AGEN)
    Dim lobjstcPedAgregados As stcA46v741B    'Objeto da classe de estrutura stcA46V741B (TAB_Ped_Agregados)
    Dim lobjItem As clsA46V708A    'Objeto da classe clsA46V708A (TAB_PED_LOC)
    Dim lobjstcItem As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjVenc As clsA46V713A    'Objeto da classe clsA46V713A (TAB_PED_VENC)
    Dim lObjstcVenc As stcA46V713B    'Objeto da classe de estrutura stcA46V713A (TAB_PED_VENC)
    Dim lobjAuto As clsA46V719A    'Objeto da classe clsA46V719A (TAB_PED_AUTO)
    Dim lobjstcAuto As stcA46V719B    'Objeto da classe de estrutura stcA46V719B (TAB_PED_AUTO)
    Dim lobjEspec As clsA46V721A    'Objeto da classe clsA46V721A (TAB_PED_ESPEC2)
    Dim lobjstcEspec As stcA46V721B    'Objeto da classe de estrutura stcA46V721B (TAB_PED_ESPEC2)
    Dim lobjTextoEndo As clsA46V722A    'Objeto da classe clsA46V722A (TAB_PED_DIZER)
    Dim lobjstcTextoEndo As stcA46V722B    'Objeto da classe de estrutura stcA46V722B (TAB_PED_DIZER)
    Dim lobjTextoEspec As clsA46V726A    'Objeto da classe clsA46V726A (TAB_PED_ESPEC)
    Dim lobjstcTextoEspec As stcA46V726B    'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC)
    Dim lobjLiberacaoSerasa As clsA46V746A
    Dim lobjstcLiberacaoSerasa As stcA46V746B

    Dim lobjProAgenc As clsA46V733A   'Objeto da classe clsA46V733A (Tab_Ped_ProAgenc)
    Dim lobjstcProAgenc As stcA46V733B      'Objeto da classe de estrutura stcA46V733B (Tab_Ped_ProAgenc)

    Dim lobjTelefones As clsA46V736A
    Dim lobjstcTelefones As stcA46V736B
    
    Err.Clear
    mfExcluirDependentes = False
    On Error GoTo ControleErro

    'Exclui endere�o
    Set lobjEndereco = New clsA46V707A
    Set lobjEndereco.M46V999 = M46V999

    Set lobjstcEndereco = New stcA46V707B
    Set lobjstcEndereco.M46V999 = M46V999

    Set lobjstcEndereco.ObjPai = pobjstc
    lobjstcEndereco.nosso_numero = pobjstc.nosso_numero
    Set lobjEndereco.ConexaoBD = Me.ConexaoBD
    If Not lobjEndereco.mfExcluir(lobjstcEndereco, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjEndereco.UltimoErro
        GoTo SAIDA
    End If

    'Exclui corretor
    Set lobjCorretor = New clsA46V705A
    Set lobjCorretor.M46V999 = M46V999

    Set lobjstcCorretor = New stcA46V705B
    Set lobjstcCorretor.M46V999 = M46V999

    Set lobjstcCorretor.ObjPai = pobjstc
    lobjstcCorretor.nosso_numero = pobjstc.nosso_numero
    Set lobjCorretor.ConexaoBD = Me.ConexaoBD
    If Not lobjCorretor.mfExcluir(lobjstcCorretor, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjCorretor.UltimoErro
        GoTo SAIDA
    End If

    'Exclui cosseguro
    Set lobjCosseguro = New clsA46V706A
    Set lobjCosseguro.M46V999 = M46V999

    Set lobjstcCosseguro = New stcA46V706B
    Set lobjstcCosseguro.M46V999 = M46V999

    Set lobjstcCosseguro.ObjPai = pobjstc
    lobjstcCosseguro.nosso_numero = pobjstc.nosso_numero
    Set lobjCosseguro.ConexaoBD = Me.ConexaoBD
    If Not lobjCosseguro.mfExcluir(lobjstcCosseguro, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjCosseguro.UltimoErro
        GoTo SAIDA
    End If

    'Exclui a cole��o de Juros
    Set lobjJuros = New clsA46V711A
    Set lobjJuros.M46V999 = M46V999

    Set lobjstcJuros = New stcA46V711B
    Set lobjstcJuros.ObjPai = pobjstc
    lobjstcJuros.nosso_numero = pobjstc.nosso_numero
    Set lobjJuros.ConexaoBD = Me.ConexaoBD
    If Not lobjJuros.mfExcluir(lobjstcJuros, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjJuros.UltimoErro
        GoTo SAIDA
    End If
    Set pobjstc.ColecaoPedJuros = lobjJuros.Colecao

    'Excluir controle de emiss�o
    Set lobjCtrlEmis = New clsA46V712A
    Set lobjCtrlEmis.M46V999 = M46V999

    Set lobjstcCtrlEmis = New stcA46V712B
    Set lobjstcCtrlEmis.ObjPai = pobjstc
    lobjstcCtrlEmis.nosso_numero = pobjstc.nosso_numero
    Set lobjCtrlEmis.ConexaoBD = Me.ConexaoBD
    If Not lobjCtrlEmis.mfExcluir(lobjstcCtrlEmis, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjCtrlEmis.UltimoErro
        GoTo SAIDA
    End If

    'Excluir Movimento de ag�ncia
    If M46V999.gPrograma <> "D046O003" Then
        Set lobjMovAgen = New clsA46V743A
        Set lobjMovAgen.M46V999 = M46V999

        Set lobjstcMovAgen = New stcA46V743B
        Set lobjstcMovAgen.ObjPai = pobjstc
        lobjstcMovAgen.nosso_numero = pobjstc.nosso_numero
        Set lobjMovAgen.ConexaoBD = Me.ConexaoBD
        If Not lobjMovAgen.mfExcluir(lobjstcMovAgen, pPropTabela, pNomeTabImport, tagDynaTrace) Then
                UltimoErro = lobjMovAgen.UltimoErro
                GoTo SAIDA
        End If
    End If

    'Excluir controle de Ped Agragados
    Set lobjPedAgregados = New clsA46v741A
    Set lobjPedAgregados.M46V999 = M46V999


    Set lobjstcPedAgregados = New stcA46v741B
    Set lobjstcPedAgregados.ObjPai = pobjstc
    lobjstcPedAgregados.nosso_numero = pobjstc.nosso_numero
    Set lobjPedAgregados.ConexaoBD = Me.ConexaoBD
    
    If M46V999.gOrigemChamada = e_Origem_Kit Then
    If Not lobjPedAgregados.mfExcluir(lobjstcPedAgregados, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjPedAgregados.UltimoErro
        GoTo SAIDA
    End If
    End If


    'Excluir os vencimentos
    Set lobjVenc = New clsA46V713A
    Set lobjVenc.M46V999 = M46V999

    Set lObjstcVenc = New stcA46V713B
    Set lObjstcVenc.M46V999 = M46V999

    Set lObjstcVenc.ObjPai = pobjstc
    lObjstcVenc.nosso_numero = pobjstc.nosso_numero
    Set lobjVenc.ConexaoBD = Me.ConexaoBD
    If Not lobjVenc.mfExcluir(lObjstcVenc, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjVenc.UltimoErro
        GoTo SAIDA
    End If

    'Exclui itens
    Set lobjItem = New clsA46V708A
    Set lobjItem.M46V999 = M46V999

    Set lobjstcItem = New stcA46V708B
    Set lobjstcItem.M46V999 = M46V999

    Set lobjstcItem.ObjPai = pobjstc
    lobjstcItem.nosso_numero = pobjstc.nosso_numero
    lobjstcItem.NUM_PI = pobjstc.NUM_PI
    Set lobjItem.ConexaoBD = Me.ConexaoBD
    If Not lobjItem.mfExcluir(lobjstcItem, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjItem.UltimoErro
        GoTo SAIDA
    End If

    'Exclui Dados Autom�vel
    Set lobjAuto = New clsA46V719A
    Set lobjAuto.M46V999 = M46V999

    Set lobjstcAuto = New stcA46V719B
    Set lobjstcAuto.ObjPai = pobjstc
    lobjstcAuto.nosso_numero = pobjstc.nosso_numero
    Set lobjAuto.ConexaoBD = Me.ConexaoBD
    If Not lobjAuto.mfExcluir(lobjstcAuto, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjAuto.UltimoErro
        GoTo SAIDA
    End If

    'Excluir Especifica��es
    Set lobjEspec = New clsA46V721A
    Set lobjEspec.M46V999 = M46V999

    Set lobjstcEspec = New stcA46V721B
    Set lobjstcEspec.ObjPai = pobjstc
    lobjstcEspec.nosso_numero = pobjstc.nosso_numero
    Set lobjEspec.ConexaoBD = Me.ConexaoBD
    If Not lobjEspec.mfExcluir(lobjstcEspec, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjEspec.UltimoErro
        GoTo SAIDA
    End If

    'Excluir Texto do Endosso
    Set lobjTextoEndo = New clsA46V722A
    Set lobjTextoEndo.M46V999 = M46V999

    Set lobjstcTextoEndo = New stcA46V722B
    Set lobjstcTextoEndo.ObjPai = pobjstc
    lobjstcTextoEndo.nosso_numero = pobjstc.nosso_numero
    Set lobjTextoEndo.ConexaoBD = Me.ConexaoBD
    If Not lobjTextoEndo.mfExcluir(lobjstcTextoEndo, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjTextoEndo.UltimoErro
        GoTo SAIDA
    End If

    'Excluir Texto do Especifica��o Padr�o
    Set lobjTextoEspec = New clsA46V726A
    Set lobjTextoEspec.M46V999 = M46V999

    Set lobjstcTextoEspec = New stcA46V726B
    Set lobjstcTextoEspec.ObjPai = pobjstc
    lobjstcTextoEspec.nosso_numero = pobjstc.nosso_numero
    Set lobjTextoEspec.ConexaoBD = Me.ConexaoBD
    If Not lobjTextoEspec.mfExcluir(lobjstcTextoEspec, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjTextoEspec.UltimoErro
        GoTo SAIDA
    End If

    'Excluir Libera��o SERASA
'    Set lobjLiberacaoSerasa = New clsA46V746A
'    Set lobjLiberacaoSerasa.M46V999 = M46V999
'
'    Set lobjstcLiberacaoSerasa = New stcA46V746B
'    lobjstcLiberacaoSerasa.Num_Pi = pobjstc.Num_Pi
'    Set lobjLiberacaoSerasa.ConexaoBD = Me.ConexaoBD
'    If Not lobjLiberacaoSerasa.mfExcluir(lobjstcLiberacaoSerasa, pPropTabela, pNomeTabImport, tagDynaTrace) Then
'        UltimoErro = lobjLiberacaoSerasa.UltimoErro
'        GoTo SAIDA
'    End If
    

    'Exclui Prolabore / Agenc
    Set lobjProAgenc = New clsA46V733A
    Set lobjProAgenc.M46V999 = M46V999

    Set lobjstcProAgenc = New stcA46V733B
    Set lobjstcProAgenc.M46V999 = M46V999

    Set lobjstcProAgenc.ObjPai = pobjstc
    lobjstcProAgenc.nosso_numero = pobjstc.nosso_numero
    Set lobjProAgenc.ConexaoBD = Me.ConexaoBD
    If Not lobjProAgenc.mfExcluir(lobjstcProAgenc, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjProAgenc.UltimoErro
        GoTo SAIDA
    End If

    'Exclui Telefones
    Set lobjTelefones = New clsA46V736A
    Set lobjTelefones.M46V999 = M46V999

    Set lobjstcTelefones = New stcA46V736B
    Set lobjstcTelefones.M46V999 = M46V999

    Set lobjstcTelefones.ObjPai = pobjstc
    lobjstcTelefones.nosso_numero = pobjstc.nosso_numero
    Set lobjTelefones.ConexaoBD = Me.ConexaoBD
    If Not lobjTelefones.mfExcluir(lobjstcTelefones, pPropTabela, pNomeTabImport, tagDynaTrace) Then
        UltimoErro = lobjTelefones.UltimoErro
        GoTo SAIDA
    End If


    '==================================
    'Fun��o eralizada com sucesso
    mfExcluirDependentes = True
    '==================================

SAIDA:
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Err.Clear
    GoSub FechaObj
    Exit Function

ControleErro:
    Call M46V999.gpGraLog(0, "A46V702A.mfExcluirDependentes ERRO - " & Err & " " & Err.Description)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "A46V702A.mfExcluirDependentes ERRO - " & Err & " " & Err.Description, 2, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set lobjEndereco = Nothing
    Set lobjstcEndereco = Nothing
    Set lobjCorretor = Nothing
    Set lobjstcCorretor = Nothing
    Set lobjCosseguro = Nothing
    Set lobjstcCosseguro = Nothing
    Set lobjJuros = Nothing
    Set lobjstcJuros = Nothing
    Set lobjCtrlEmis = Nothing
    Set lobjstcCtrlEmis = Nothing
    Set lobjVenc = Nothing
    Set lObjstcVenc = Nothing
    Set lobjItem = Nothing
    Set lobjstcItem = Nothing
    Set lobjAuto = Nothing
    Set lobjstcAuto = Nothing
    Set lobjEspec = Nothing
    Set lobjstcEspec = Nothing
    Set lobjTextoEndo = Nothing
    Set lobjstcTextoEndo = Nothing
    Set lobjTextoEspec = Nothing
    Set lobjstcTextoEspec = Nothing
    Set lobjLiberacaoSerasa = Nothing
    Set lobjProAgenc = Nothing
    Set lobjstcProAgenc = Nothing
    Set lobjstcTelefones = Nothing
    Set lobjTelefones = Nothing
    Return

End Function

Public Function mfCarregarColecaoInspecaoWeb(ByVal pCorretor As Long, _
                                             ByVal pPropTabela As String, _
                                             ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V702B          'Objeto da classe de estrutura stcA46V702B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lsSQl = " select distinct p.* From " & pPropTabela & "tab_ped_loc as l inner join " & pPropTabela & "tab_ped as p on l.nosso_numero = p.nosso_numero"
    lsSQl = lsSQl & " inner join " & pPropTabela & "tab_ped_corr as c on c.nosso_numero = p.nosso_numero"
    lsSQl = lsSQl & " WHERE  Cod_corr in (SELECT * FROM [Ged].[dbo].[Func_BuscaCorrGrupoAtivo] ('" & Format(pCorretor, "0000000") & "', 112))"
    lsSQl = lsSQl & " and Cod_Sit_Ped = 2 and Cod_Tip_Doc = 1 and Insp_Obr = 1 and Insp_Agend = 0 and p.DAT_DIGIT > " & Format(Date - 35, "yyyymmdd")

    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V702B
        Set lobjstc.M46V999 = M46V999

        Call mpCarregarAtributos(lobjstc, lrst, pNomeTabImport)
        If Not mfCarregarDependentes(lobjstc, pPropTabela, pNomeTabImport, True) Then GoTo ControleErro
        Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID
        lrst.MoveNext
    Loop
    mfCarregarColecaoInspecaoWeb = True

SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    mfCarregarColecaoInspecaoWeb = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V702A.mfCarregarColecaoInspecaoWeb ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V702A.mfCarregarColecaoInspecaoWeb ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    GoTo SAIDA
FechaObj:
    Set lobjstc = Nothing
    Set lrst = Nothing
    Return
End Function

Public Function mfInspecaoObrigatoria(ByVal pNossoNumero As String, _
                                      ByVal pPropTabela As String, _
                                      ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V702B          'Objeto da classe de estrutura stcA46V702B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro
    
    mfInspecaoObrigatoria = False
    
    lsSQl = " SELECT COUNT(*) Contador "
    lsSQl = lsSQl & " From " & pPropTabela & "..Tab_Ped_Loc"
    lsSQl = lsSQl & " WHERE Insp_Obr <> 0 and  nosso_numero = " & M46V999.gfFormatarCamposSQL(pNossoNumero)
    
    If Not M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrst, UltimoErro) Then
            If M46V999.gfPreenchido(UltimoErro) Then
                GoTo ControleErro
            End If
    End If
    
    If Not lrst.EOF Then
        mfInspecaoObrigatoria = (lrst!contador > 0)
    End If
    
SAIDA:
    GoSub FechaObj
    Exit Function
    
ControleErro:
    mfInspecaoObrigatoria = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V702A.mfInspecaoObrigatoria ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V702A.mfInspecaoObrigatoria ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    GoTo SAIDA
FechaObj:
    Set lobjstc = Nothing
    Set lrst = Nothing
    Return
End Function

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V702B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String, _
                                  Optional ByVal pCarregarDependentes As Boolean = False, _
                                  Optional ByVal pSomenteItem As Boolean = False, _
                                  Optional ByVal pChavePI As Boolean = False) As Boolean
    Dim lobjstc As stcA46V702B          'Objeto da classe de estrutura stcA46V702B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)

    If InStr(UCase(lsSQl), "WHERE") = 0 And M46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        UltimoErro = "FGE0004 - Favor Selecionar um par�metro de consulta "
        GoTo SAIDA
    End If

    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrst, UltimoErro) = False Then GoTo ControleErro

    If Not lrst Is Nothing Then
        If lrst.RecordCount > 1000 Then
            UltimoErro = "FGE0004 - Pesquisa resultou em " & lrst.RecordCount & _
                       " registros. - quantidade excedente.(Pesquisa de pedidos)"
            GoTo SAIDA
        End If
        If lrst.RecordCount > 0 Then
          If UCase(pPropTabela) <> "GED.DBO." Then
                If lrst!Cod_Produto >= 1080 And lrst!Cod_Estip < 2000 Then
                        M46V999.gCodProduto = lrst!Cod_Produto
                        M46V999.gCod_Ramo = lrst!Cod_Ramo
                        M46V999.M46V111.gCod_plano = M46V999.gfPlanoBasico(lrst!Cod_Ramo, lrst!Cod_Estip)
                    
                        If M46V999.gRst_Tbm_De_Para.State = 0 Then
                                Call M46V999.mfCarregarDePara
                        ElseIf M46V999.gRst_Tbm_De_Para.RecordCount = 0 Then
                                Call M46V999.mfCarregarDePara
                    End If
                End If
           End If
        End If
            
        ' Boleto
        Call mpLimparColecao
        Do While Not lrst.EOF
            Set lobjstc = New stcA46V702B
            Set lobjstc.M46V999 = M46V999
            Call mpCarregarAtributos(lobjstc, lrst, pNomeTabImport)
            If pCarregarDependentes = True Or pSomenteItem = True Then

                If Not mfCarregarDependentes(lobjstc, pPropTabela, pNomeTabImport, pSomenteItem) Then GoTo ControleErro
            End If
            If pChavePI = True Then
                
                Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID_PI
            Else
                
                Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID
            End If
            lrst.MoveNext
        Loop
    End If
    mfCarregarColecao = True
     

SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    mfCarregarColecao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V702A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V702A.mfCarregarColecao ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    GoTo SAIDA
FechaObj:
    Set lobjstc = Nothing
    Set lrst = Nothing
    Return
End Function

Private Function mfCarregarDependentes(ByRef pobjstc As stcA46V702B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       ByVal pSomenteItem As Boolean) As Boolean
    Dim lobjEndereco As clsA46V707A   'Objeto da classe clsA46V707A (TAB_PED_ENDER)
    Dim lobjstcEndereco As stcA46V707B      'Objeto da classe de estrutura stcA46V707B (TAB_PED_ENDER)
    Dim lobjCorretor As clsA46V705A   'Objeto da classe clsA46V705A (TAB_PED_CORR)
    Dim lobjstcCorretor As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    Dim lobjCosseguro As clsA46V706A    'Objeto da classe clsA46V706A (TAB_PED_COSS)
    Dim lobjstcCosseguro As stcA46V706B       'Objeto da classe de estrutura stcA46V706B (TAB_PED_COSS)
    Dim lobjJuros As clsA46V711A    'Objeto da classe clsA46V711A (TAB_PED_JUROS)
    Dim lobjstcJuros As stcA46V711B   'Objeto da classe de estrutura stcA46V711B (TAB_PED_JUROS)
    Dim lobjLoc As clsA46V708A    'Objeto da classe clsA46V708A (TAB_PED_LOC)
    Dim lobjstcLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjCtrlEmis As clsA46V712A   'Objeto da classe clsA46V712A (TAB_CTRL_EMIS)
    Dim lobjMovAgen As clsA46V743A   'Objeto da classe clsA46V743A (TAB_MOV_AGEN)
    Dim lobjPedAgregados As clsA46v741A   'Objeto da classe clsA46V743A (TAB_PED_AGREGADOS)
    Dim lobjstcCtrlEmis As stcA46V712B      'Objeto da classe de estrutura stcA46V712A (TAB_CTRL_EMIS)
    Dim lobjstcMovAgen As stcA46V743B      'Objeto da classe de estrutura stcA46V743A (TAB_MOV_AGEN)
    Dim lobjstcPedAgregados As stcA46v741B      'Objeto da classe de estrutura stcA46V743A (TAB_PED_AGREGADOS)
    Dim lobjVenc As clsA46V713A    'Objeto da classe clsA46V713A (TAB_PED_VENC)
    Dim lObjstcVenc As stcA46V713B  'Objeto da classe de estrutura stcA46V713A (TAB_PED_VENC)
    Dim lobjAuto As clsA46V719A    'Objeto da classe clsA46V719A (TAB_PED_AUTO)
    Dim lobjstcAuto As stcA46V719B  'Objeto da classe de estrutura stcA46V719B (TAB_PED_AUTO)
    Dim lobjEspec As clsA46V721A    'Objeto da classe clsA46V721A (TAB_PED_ESPEC2)
    Dim lobjstcEspec As stcA46V721B   'Objeto da classe de estrutura stcA46V721A (TAB_PED_ESPEC2)
    Dim lobjTextoEndo As clsA46V722A    'Objeto da classe clsA46V722A (TAB_PED_DIZER)
    Dim lobjstcTextoEndo As stcA46V722B       'Objeto da classe de estrutura stcA46V722A (TAB_PED_DIZER)
    Dim lobjTextoEspec As clsA46V726A     'Objeto da classe clsA46V726A (TAB_PED_ESPEC)
    Dim lobjstcTextoEspec As stcA46V726B        'Objeto da classe de estrutura stcA46V726A (TAB_PED_ESPEC)
    Dim lobjstcP0045400 As stcA28V721B
    Dim lobjLiberacaoSerasa As clsA46V746A
    Dim lobjstcLiberacaoSerasa As stcA46V746B
    Dim lobjP0045400 As clsA28V721A
    Dim lobjstcTrans As stcA46V718B
    Dim lobjTrans As clsA46V718A
    Dim lDigitoPI As String
    Dim clsA99V001 As clsA99V001
    Dim lSelect As String
    Dim lrst As clsYasRecordSet
    Dim lMensagem As String
    Dim lobjP0045200 As clsA28V720A
    Dim lobjstcPedLMI As stcA46V732B
    Dim lobjPedLMI As clsA46V732A

    Dim lobjProAgenc As clsA46V733A   'Objeto da classe clsA46V733A (Tab_Ped_ProAgenc)
    Dim lobjstcProAgenc As stcA46V733B      'Objeto da classe de estrutura stcA46V733B (Tab_Ped_ProAgenc)

    Dim lobjTelefones As clsA46V736A
    Dim lobjstcTelefones As stcA46V736B
    Dim lobjServicoAss24h As clsA46V749A
    Dim lobjstcServicoAss24h As stcA46V749B


    On Error GoTo ControleErro

    mfCarregarDependentes = False
    If pSomenteItem = False Then
        'Carregar o objeto de Endere�o
        Set lobjEndereco = New clsA46V707A
        Set lobjEndereco.M46V999 = M46V999

        Set lobjstcEndereco = New stcA46V707B
        Set lobjstcEndereco.M46V999 = M46V999

        lobjstcEndereco.nosso_numero = pobjstc.nosso_numero
        lobjstcEndereco.Num_Apol = pobjstc.Num_Apol
        lobjstcEndereco.Num_Endo = pobjstc.Num_Endo
        Set lobjstcEndereco.ObjPai = pobjstc
        Set lobjEndereco.ConexaoBD = Me.ConexaoBD
        If Not lobjEndereco.mfCarregarColecao(lobjstcEndereco, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjEndereco.UltimoErro
            GoTo ControleErro
        End If
        If lobjEndereco.Colecao.Count > 0 Then
            Set pobjstc.objstcPedEndereco = lobjEndereco.Colecao.Item(1)
        Else
            Set pobjstc.objstcPedEndereco = New stcA46V707B
            Set pobjstc.objstcPedEndereco.M46V999 = M46V999
        End If
        
        Set lobjEndereco = Nothing
        Set lobjstcEndereco = Nothing


        'Carregar a cole��o de corretores
        Set lobjCorretor = New clsA46V705A
        Set lobjCorretor.M46V999 = M46V999

        Set lobjstcCorretor = New stcA46V705B
        Set lobjstcCorretor.M46V999 = M46V999

        lobjstcCorretor.nosso_numero = pobjstc.nosso_numero
        lobjstcCorretor.Num_Apol = pobjstc.Num_Apol
        lobjstcCorretor.Num_Endo = pobjstc.Num_Endo
        Set lobjstcCorretor.ObjPai = pobjstc
        Set lobjCorretor.ConexaoBD = Me.ConexaoBD
        If Not lobjCorretor.mfCarregarColecao(lobjstcCorretor, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjCorretor.UltimoErro
            GoTo ControleErro
        End If
        Set pobjstc.ColecaoPedCorr = lobjCorretor.Colecao
        

        Set lobjCorretor = Nothing
        Set lobjstcCorretor = Nothing


        'Carregar a cole��o de cosseguro
        Set lobjCosseguro = New clsA46V706A
        Set lobjCosseguro.M46V999 = M46V999

        Set lobjstcCosseguro = New stcA46V706B
        Set lobjstcCosseguro.M46V999 = M46V999

        lobjstcCosseguro.nosso_numero = pobjstc.nosso_numero
        lobjstcCosseguro.Num_Apol = pobjstc.Num_Apol
        lobjstcCosseguro.Num_Endo = pobjstc.Num_Endo
        Set lobjstcCosseguro.ObjPai = pobjstc
        Set lobjCosseguro.ConexaoBD = Me.ConexaoBD
        If Not lobjCosseguro.mfCarregarColecao(lobjstcCosseguro, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjCosseguro.UltimoErro
            GoTo ControleErro
        End If
        Set pobjstc.ColecaoPedCoss = lobjCosseguro.Colecao
        Set lobjCosseguro = Nothing
        Set lobjstcCosseguro = Nothing
        

        'Carregar a cole��o de Juros
        Set lobjJuros = New clsA46V711A
        Set lobjJuros.M46V999 = M46V999

        Set lobjstcJuros = New stcA46V711B
        lobjstcJuros.nosso_numero = pobjstc.nosso_numero
        Set lobjstcJuros.ObjPai = pobjstc
        Set lobjJuros.ConexaoBD = Me.ConexaoBD
        If UCase(pPropTabela) <> "GED.DBO." Then
            If Not lobjJuros.mfCarregarColecao(lobjstcJuros, pPropTabela, pNomeTabImport) Then
                UltimoErro = lobjJuros.UltimoErro
                GoTo ControleErro
            End If
        End If
        Set pobjstc.ColecaoPedJuros = lobjJuros.Colecao
        Set lobjJuros = Nothing
        Set lobjstcJuros = Nothing
        
        
        'Carregar o controle de emiss�o
        Set lobjCtrlEmis = New clsA46V712A
        Set lobjCtrlEmis.M46V999 = M46V999

        Set lobjstcCtrlEmis = New stcA46V712B
        lobjstcCtrlEmis.nosso_numero = pobjstc.nosso_numero
        lobjstcCtrlEmis.Num_Apol = pobjstc.Num_Apol
        lobjstcCtrlEmis.Num_Endo = pobjstc.Num_Endo
        Set lobjstcCtrlEmis.ObjPai = pobjstc
        Set lobjCtrlEmis.ConexaoBD = Me.ConexaoBD
        If Not lobjCtrlEmis.mfCarregarColecao(lobjstcCtrlEmis, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjCtrlEmis.UltimoErro
            GoTo ControleErro
        End If
        If lobjCtrlEmis.Colecao.Count > 0 Then
            Set pobjstc.objstcCtrlEmis = lobjCtrlEmis.Colecao.Item(1)
        End If
        

        'Carregar a cole��o de vencimento
        Set lobjVenc = New clsA46V713A
        Set lobjVenc.M46V999 = M46V999

        Set lObjstcVenc = New stcA46V713B
        Set lObjstcVenc.M46V999 = M46V999

        lObjstcVenc.nosso_numero = pobjstc.nosso_numero
        Set lObjstcVenc.ObjPai = pobjstc
        Set lobjVenc.ConexaoBD = Me.ConexaoBD
        If UCase(pPropTabela) <> "GED.DBO." Then
            If Not lobjVenc.mfCarregarColecao(lObjstcVenc, pPropTabela, pNomeTabImport) Then
                UltimoErro = lobjVenc.UltimoErro
                GoTo ControleErro
            End If
        End If
        Set pobjstc.ColecaoPedVenc = lobjVenc.Colecao
        Set lobjVenc = Nothing
        Set lObjstcVenc = Nothing
        

        'Carregar o objeto de Auto
        Set lobjAuto = New clsA46V719A
        Set lobjAuto.M46V999 = M46V999

        Set lobjstcAuto = New stcA46V719B
        lobjstcAuto.nosso_numero = pobjstc.nosso_numero
        Set lobjstcAuto.ObjPai = pobjstc
        Set lobjAuto.ConexaoBD = Me.ConexaoBD
        If UCase(pPropTabela) <> "GED.DBO." Then
            If Not lobjAuto.mfCarregarColecao(lobjstcAuto, pPropTabela, pNomeTabImport) Then
                UltimoErro = lobjAuto.UltimoErro
                GoTo ControleErro
            End If
        End If
        If lobjAuto.Colecao.Count > 0 Then
            Set pobjstc.objstcPedAuto = lobjAuto.Colecao.Item(1)
        Else
            Set pobjstc.objstcPedAuto = New stcA46V719B
        End If
        Set lobjAuto = Nothing
        Set lobjstcAuto = Nothing
        
    End If

    'Carregar a cole��o de itens
    Set lobjLoc = New clsA46V708A
    Set lobjLoc.M46V999 = M46V999


    Set lobjstcLoc = New stcA46V708B
    Set lobjstcLoc.M46V999 = M46V999

    lobjstcLoc.nosso_numero = pobjstc.nosso_numero
    lobjstcLoc.Num_Apol = pobjstc.Num_Apol
    lobjstcLoc.Num_Endo = pobjstc.Num_Endo
    Set lobjstcLoc.ObjPai = pobjstc
    Set lobjLoc.ConexaoBD = Me.ConexaoBD
    If Not lobjLoc.mfCarregarColecao(lobjstcLoc, pPropTabela, pNomeTabImport, (pSomenteItem = False)) Then
        UltimoErro = lobjLoc.UltimoErro
        GoTo ControleErro
    End If
    Set pobjstc.ColecaoPedLoc = lobjLoc.Colecao
    Set lobjLoc = Nothing
    Set lobjstcLoc = Nothing
    

    'Carregar o tabela de especifica��o
    Set lobjEspec = New clsA46V721A
    Set lobjEspec.M46V999 = M46V999

    Set lobjstcEspec = New stcA46V721B
    lobjstcEspec.nosso_numero = pobjstc.nosso_numero
    Set lobjstcEspec.ObjPai = pobjstc
    Set lobjEspec.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjEspec.mfCarregarColecao(lobjstcEspec, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjEspec.UltimoErro
            GoTo ControleErro
        End If
    End If
    If lobjEspec.Colecao.Count > 0 Then
        Set pobjstc.objstcEspec = lobjEspec.Colecao.Item(1)
    End If
    
    
    'Carregar o Texto do Endosso
    Set lobjTextoEndo = New clsA46V722A
    Set lobjTextoEndo.M46V999 = M46V999

    Set lobjstcTextoEndo = New stcA46V722B
    lobjstcTextoEndo.nosso_numero = pobjstc.nosso_numero
    Set lobjstcTextoEndo.ObjPai = pobjstc
    Set lobjTextoEndo.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjTextoEndo.mfCarregarColecao(lobjstcTextoEndo, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjTextoEndo.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set pobjstc.ColecaoTextoEndo = lobjTextoEndo.Colecao
    Set lobjTextoEndo = Nothing
    Set lobjstcTextoEndo = Nothing
    

    'Carregar o Texto da Especifica��o Padr�o
    Set lobjTextoEspec = New clsA46V726A
    Set lobjTextoEspec.M46V999 = M46V999

    Set lobjstcTextoEspec = New stcA46V726B
    lobjstcTextoEspec.nosso_numero = pobjstc.nosso_numero
    Set lobjstcTextoEspec.ObjPai = pobjstc
    Set lobjTextoEspec.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjTextoEspec.mfCarregarColecao(lobjstcTextoEspec, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjTextoEspec.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set pobjstc.ColecaoTextoEspec = lobjTextoEspec.Colecao
    Set lobjTextoEspec = Nothing
    Set lobjstcTextoEspec = Nothing
    

    'Carregar Libera��o do Serasa
    Set lobjLiberacaoSerasa = New clsA46V746A
    Set lobjLiberacaoSerasa.M46V999 = M46V999

    Set lobjstcLiberacaoSerasa = New stcA46V746B
    lobjstcLiberacaoSerasa.nosso_numero = pobjstc.nosso_numero
    Set lobjLiberacaoSerasa.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjLiberacaoSerasa.mfCarregar(lobjstcLiberacaoSerasa, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjLiberacaoSerasa.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set pobjstc.objstcLiberacaoSerasa = lobjstcLiberacaoSerasa
    Set lobjLiberacaoSerasa = Nothing
    
    
    Set pobjstc.objstcP0045400 = New stcA28V721B
    pobjstc.Ind_CotacaoSolic = False
    pobjstc.mCotacaoSiscota = ""
    pobjstc.StatusCotSiscota = enum_StatusTransm_Inicio     '0
    pobjstc.mStatusDivSiscota = enum_StatusTransm_Inicio     '0


    'Classe 4
    If (M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) And UCase(pPropTabela) <> "RAMOSDIVERSOS.DBO." Then        'Se for corretor
        '***COTACAO
        If Dir(M46V999.gEndP0045400) <> "" Or M46V999.gPrograma = "D046O003" Then
            Set lobjstcP0045400 = New stcA28V721B
            Set lobjP0045400 = New clsA28V721A
            Set lobjP0045400.M46V999 = M46V999

            lobjstcP0045400.nosso_numero = pobjstc.nosso_numero
            If Not lobjP0045400.gfCarregarColecao(lobjstcP0045400, pPropTabela) Then
                Call M46V999.gpGraLog(1, lobjP0045400.UltimoErro)
                GoTo SAIDA
            Else
                If lobjP0045400.Colecao.Count = 1 Then
                    Set lobjstcP0045400 = lobjP0045400.Colecao.Item(1)
                    Set pobjstc.objstcP0045400 = lobjP0045400.Colecao.Item(1)
                    pobjstc.Ind_CotacaoSolic = True
                    pobjstc.StatusCotSiscota = lobjstcP0045400.STATUS
                    pobjstc.SiscMaxDesc = lobjstcP0045400.MaxDesc
                    pobjstc.SenhaCotacao = lobjstcP0045400.Senha
                    pobjstc.objstcCtrlEmis.Num_Cotac = lobjstcP0045400.CCotacao
                    pobjstc.objstcCtrlEmis.Tip_Cotac = 2
                End If
            End If
        End If
    Else
        If Not UCase(ConexaoBD.Properties("Password")) Like "*P00MPED46*" Then
            'Verifico
            lSelect = " Select * from siscota.dbo.tab_syas_cotacao WITH (NOLOCK) where " & _
                    " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
            If Not M46V999.gfObtRegistro(ConexaoBD, lSelect, lrst, UltimoErro) Then
                GoTo SAIDA
            End If
            If Not lrst.EOF Then
                pobjstc.mCotacaoSiscota = lrst("CCotacao")
                pobjstc.mStatusDivSiscota = lrst("cod_status")
                pobjstc.objstcCtrlEmis.Num_Cotac = pobjstc.mCotacaoSiscota
                pobjstc.objstcCtrlEmis.Tip_Cotac = 2
            Else
                'Verifico se tem na tabela de c�pia
                lSelect = " select s.* from ramosdiversos.dbo.tab_ped_num_copia as c"
                lSelect = lSelect & " inner join siscota.dbo.tab_syas_cotacao as s "
                lSelect = lSelect & " on c.nosso_numero_old = S.Nosso_Numero "
                lSelect = lSelect & " where nosso_numero_new = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)

                If Not M46V999.gfObtRegistro(ConexaoBD, lSelect, lrst, UltimoErro) Then
                    GoTo SAIDA
                End If
                If Not lrst.EOF Then
                    pobjstc.mCotacaoSiscota = lrst("CCotacao")
                    pobjstc.mStatusDivSiscota = lrst("cod_status")
                    pobjstc.objstcCtrlEmis.Num_Cotac = pobjstc.mCotacaoSiscota
                    pobjstc.objstcCtrlEmis.Tip_Cotac = 2
                End If
            End If
        End If
        If M46V999.gTipoAcesso = DIGITACAO And Not M46V999.gBaseLocal Then
            'Verifico se tem na tabela de c�pia
            lSelect = " select f.* from ramosdiversos.dbo.tab_ped_num_copia as c"
            lSelect = lSelect & " inner join ramosdiversos.dbo.ftp_tab_ped as f "
            lSelect = lSelect & " on c.nosso_numero_old = f.Nosso_Numero "
            lSelect = lSelect & " where nosso_numero_new = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)

            If Not M46V999.gfObtRegistro(ConexaoBD, lSelect, lrst, UltimoErro) Then
                GoTo SAIDA
            End If
            If Not lrst.EOF Then
                pobjstc.mTransmissao = True
                pobjstc.mValPrTransmitido = lrst("val_pr")
                pobjstc.mValPrAntecTransmitido = lrst("val_pr_antec")
                On Error Resume Next
                pobjstc.Val_RendaFatu = M46V999.gfTratarNulo(lrst("Val_RendaFatu"), TpoValor)
                On Error GoTo 0
            End If

            '1000913 - Permitir alterar dados cotacao interna
            'Verifico se tem na tabela de c�pia
            lSelect = " select f.* from ramosdiversos.dbo.tab_ped_num_copia as c"
            lSelect = lSelect & " inner join ramosdiversos.dbo.tab_ped as f "
            lSelect = lSelect & " on c.nosso_numero_old = f.Nosso_Numero "
            lSelect = lSelect & " where nosso_numero_new = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)

            If Not M46V999.gfObtRegistro(ConexaoBD, lSelect, lrst, UltimoErro) Then
                GoTo SAIDA
            End If
            pobjstc.mCotacaoInterna = (Not lrst.EOF)

        End If
    End If

    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor And pobjstc.Cod_Sit_Ped = e_SitPedido_Transmitido Then
        Set lobjstcTrans = New stcA46V718B
        Set lobjTrans = New clsA46V718A
        Set lobjTrans.M46V999 = M46V999

        lobjstcTrans.nosso_numero = pobjstc.nosso_numero
        If Not lobjTrans.mfCarregarColecao(lobjstcTrans, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjTrans.UltimoErro
            GoTo SAIDA
        End If
        If lobjTrans.Colecao.Count > 0 Then
            Set lobjstcTrans = lobjTrans.Colecao.Item(1)
            pobjstc.Num_PI_Emissao = lobjstcTrans.NUM_PI
            pobjstc.Dat_Transmissao = lobjstcTrans.Dat_Hor_Envio
            pobjstc.Sit_Emissao = val(lobjstcTrans.Cod_Acao)
            Set clsA99V001 = New clsA99V001
            Set clsA99V001.M46V999 = M46V999

            If clsA99V001.DIG_YASUDA(pobjstc.Num_PI_Emissao, lDigitoPI, lMensagem) = False Then
                GoTo SAIDA
            End If
            pobjstc.Dig_PI_Emissao = lDigitoPI
        End If
    End If

    Set pobjstc.objstcBoleto = New stcA28V720B
    Set pobjstc.objstcBoleto.M46V999 = M46V999

    If (M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) And _
        pobjstc.objstcCtrlEmis.Num_Bloq <> 0 And _
        (pobjstc.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 And pobjstc.Cod_Bco_Cobr <> 0) And _
         pobjstc.Cod_Reg <> 120 And pobjstc.Cod_Reg <> 129 And pobjstc.Cod_Reg <> 106 And pobjstc.Cod_Reg <> 107 Then
        
        Set lobjP0045200 = New clsA28V720A
        Set lobjP0045200.M46V999 = M46V999

       ' Rotina que reserva boleto na P0045200 foi desativada por causa do projeto "boletos registrados"
        If Not lobjP0045200.gfCarregarBoletoDisp(pobjstc.nosso_numero, False, pobjstc.Nosso_Numero_Original) Then
            UltimoErro = lobjP0045200.UltimoErro
            GoTo SAIDA
        End If
        If lobjP0045200.Colecao.Count > 0 Then
            Set pobjstc.objstcBoleto = lobjP0045200.Colecao.Item(1)
            pobjstc.objstcBoleto.DigBoleto = pobjstc.objstcCtrlEmis.DIG_BLOQUETO
            pobjstc.objstcBoleto.Vigencia = M46V999.gfForData(2, pobjstc.Dat_Ini_Vig) & " a " & M46V999.gfForData(2, pobjstc.Dat_Ter_Vig)
        End If
        
        
    End If

    'If pobjstc.LMI_Unico = 1 And M46V999.gTipUsuario = e_Tipo_Usu_Yasuda And M46V999.gBaseLocal = False Then
    If pobjstc.LMI_Unico = 1 Then
        Set lobjstcPedLMI = New stcA46V732B
        Set lobjstcPedLMI.M46V999 = M46V999

        Set lobjPedLMI = New clsA46V732A
        Set lobjPedLMI.M46V999 = M46V999

        lobjstcPedLMI.nosso_numero = pobjstc.nosso_numero
        lobjstcPedLMI.Num_Apol = pobjstc.Num_Apol
        lobjstcPedLMI.Num_Endo = pobjstc.Num_Endo
        Set lobjstcPedLMI.ObjPai = pobjstc
        If Not lobjPedLMI.mfCarregarColecao(lobjstcPedLMI, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjPedLMI.UltimoErro
            GoTo ControleErro
        End If
        Set pobjstc.ColecaoLMI = lobjPedLMI.Colecao
    End If
    


    'Carregar a cole��o de Prolabore / Agenciamento
    Set lobjProAgenc = New clsA46V733A
    Set lobjProAgenc.M46V999 = M46V999

    Set lobjstcProAgenc = New stcA46V733B
    Set lobjstcProAgenc.M46V999 = M46V999

    lobjstcProAgenc.nosso_numero = pobjstc.nosso_numero
    lobjstcProAgenc.NUM_PI = pobjstc.NUM_PI
    lobjstcProAgenc.Tip_Canal = 2
    Set lobjstcProAgenc.ObjPai = pobjstc
    Set lobjProAgenc.ConexaoBD = Me.ConexaoBD
    If Not lobjProAgenc.mfCarregarColecao(lobjstcProAgenc, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjProAgenc.UltimoErro
        GoTo ControleErro
    End If
    Set pobjstc.ColecaoProlabore = lobjProAgenc.Colecao
    Set lobjProAgenc = Nothing
    Set lobjstcProAgenc = Nothing

    Set lobjProAgenc = New clsA46V733A
    Set lobjProAgenc.M46V999 = M46V999

    Set lobjstcProAgenc = New stcA46V733B
    Set lobjstcProAgenc.M46V999 = M46V999

    Set lobjProAgenc.Colecao = New Collection
    lobjstcProAgenc.nosso_numero = pobjstc.nosso_numero
    lobjstcProAgenc.NUM_PI = pobjstc.NUM_PI
    lobjstcProAgenc.Tip_Canal = 3
    Set lobjstcProAgenc.ObjPai = pobjstc
    Set lobjProAgenc.ConexaoBD = Me.ConexaoBD
    If Not lobjProAgenc.mfCarregarColecao(lobjstcProAgenc, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjProAgenc.UltimoErro
        GoTo ControleErro
    End If
    Set pobjstc.ColecaoAgenciador = lobjProAgenc.Colecao
    Set lobjProAgenc = Nothing
    Set lobjstcProAgenc = Nothing
    

    'Carregar a cole��o de Telefones
    Set lobjTelefones = New clsA46V736A
    Set lobjTelefones.M46V999 = M46V999

    Set lobjstcTelefones = New stcA46V736B
    Set lobjstcTelefones.M46V999 = M46V999

    lobjstcTelefones.nosso_numero = pobjstc.nosso_numero
    lobjstcTelefones.NUM_PI = pobjstc.NUM_PI
    Set lobjstcTelefones.ObjPai = pobjstc
    Set lobjTelefones.ConexaoBD = Me.ConexaoBD
    If Not lobjTelefones.mfCarregarColecao(lobjstcTelefones, pPropTabela, pNomeTabImport) Then
        UltimoErro = lobjTelefones.UltimoErro
        GoTo ControleErro
    End If

    Set pobjstc.ColecaoTelefones = lobjTelefones.Colecao
    Set lobjTelefones = Nothing
    Set lobjstcTelefones = Nothing
    
    'Carregar a cole��o de Servico Ass24h
    Set lobjServicoAss24h = New clsA46V749A
    Set lobjServicoAss24h.M46V999 = M46V999

    M46V999.gCodProduto = pobjstc.Cod_Produto
    Set lobjstcServicoAss24h = New stcA46V749B
    Set lobjstcServicoAss24h.M46V999 = M46V999

    lobjstcServicoAss24h.Cod_Ramo = pobjstc.Cod_Ramo
    
    Set lobjstcServicoAss24h.ObjPai = pobjstc
    Set lobjServicoAss24h.ConexaoBD = Me.ConexaoBD
    
    If Not lobjServicoAss24h.mfCarregarColecao(lobjstcServicoAss24h) Then
        UltimoErro = lobjServicoAss24h.UltimoErro
        GoTo ControleErro
    End If

    Set pobjstc.ColecaoServicoAss24h = lobjServicoAss24h.Colecao
    Set lobjServicoAss24h = Nothing
    Set lobjstcServicoAss24h = Nothing
    
    mfCarregarDependentes = True

SAIDA:
    GoSub FechaObj
    Exit Function

ControleErro:
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V702A.mfCarregarDependentes ERRO - " & Err & " " & Err.Description)
    Else
        Call M46V999.gpGraLog(0, "A46V702A.mfCarregarDependentes ERRO - " & UltimoErro)
    End If
    GoTo SAIDA
FechaObj:
    Set lobjEndereco = Nothing
    Set lobjstcEndereco = Nothing
    Set lobjCorretor = Nothing
    Set lobjstcCorretor = Nothing
    Set lobjCosseguro = Nothing
    Set lobjstcCosseguro = Nothing
    Set lobjJuros = Nothing
    Set lobjstcJuros = Nothing
    Set lobjLoc = Nothing
    Set lobjstcLoc = Nothing
    Set lobjCtrlEmis = Nothing
    Set lobjstcCtrlEmis = Nothing
    Set lobjVenc = Nothing
    Set lObjstcVenc = Nothing
    Set lobjAuto = Nothing
    Set lobjstcAuto = Nothing
    Set lobjEspec = Nothing
    Set lobjstcEspec = Nothing
    Set lobjTextoEndo = Nothing
    Set lobjstcTextoEndo = Nothing
    Set lobjTextoEspec = Nothing
    Set lobjstcTextoEspec = Nothing
    Set lobjstcP0045400 = Nothing
    Set lobjLiberacaoSerasa = Nothing
    Set lobjP0045400 = Nothing
    Set lobjstcTrans = Nothing
    Set lobjTrans = Nothing
    Set clsA99V001 = Nothing
    Set lrst = Nothing
    Set lobjstcPedLMI = Nothing
    Set lobjPedLMI = Nothing
    Set lobjProAgenc = Nothing
    Set lobjstcProAgenc = Nothing

    Return
End Function

Private Sub Class_Initialize()
    Set Colecao = New Collection
    Set mClsM46V999 = Nothing
End Sub
Private Sub Class_Terminate()
    Set Colecao = Nothing
    Call M46V999.gpFecha2(ConexaoBD)
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
