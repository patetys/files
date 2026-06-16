VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V712A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V712A
' Data/Hora: 19/09/2005 - 10:20:46
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conex�o com o Banco de Dados
Public Colecao As Collection    ' Cole��o de objetos do tipo stcA46V712B
Public UltimoErro As String           ' Descri��o do �ltimo erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V712A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V712B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       ByVal TabAtual As Integer) As String
    Dim lInsert As String                'Query.

    If TabAtual = 2 Then GoTo IncluirCompl:

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ctrl_Emis ( "
    lInsert = lInsert & " Nosso_Numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " NUM_PI,"
    lInsert = lInsert & " NUM_BLOQ,"
    lInsert = lInsert & " COD_UNID,"
    lInsert = lInsert & " COD_RAMO,"
    lInsert = lInsert & " DAT_INI_VIG,"
    lInsert = lInsert & " DAT_TER_VIG,"
    lInsert = lInsert & " Cod_Corr,"
    lInsert = lInsert & " Per_Corr,"
    lInsert = lInsert & " Cod_Corr_2,"
    lInsert = lInsert & " Per_Corr_2,"
    lInsert = lInsert & " Cod_Corr_3,"
    lInsert = lInsert & " Per_Corr_3,"
    lInsert = lInsert & " Nom_Seg,"
    lInsert = lInsert & " Qtd_Anex,"
    lInsert = lInsert & " Num_Apol_Ant,"
    lInsert = lInsert & " Dsc_Obs_Prop,"
    lInsert = lInsert & " Tip_Cotac,"
    lInsert = lInsert & " Num_Cotac,"
    lInsert = lInsert & " Per_Juros_Neg,"
    lInsert = lInsert & " Cod_Envi_Doc,"
    lInsert = lInsert & " Dat_Digit,"
    lInsert = lInsert & " Hor_Digit,"
    lInsert = lInsert & " Cod_Iden_Stat_Acao,"
    lInsert = lInsert & " Cod_User_Digit,"
    lInsert = lInsert & " Dat_Digit_MF,"
    lInsert = lInsert & " Hor_Digit_MF,"
    lInsert = lInsert & " Cod_User_Digit_MF,"
    lInsert = lInsert & " Val_Pr_Tot,"
    lInsert = lInsert & " Cod_Iden_Classe,"
    lInsert = lInsert & " Cod_Iden_Cobr,"
    lInsert = lInsert & " Num_Prop,"
    lInsert = lInsert & " Dat_Canc_Apol,"
    lInsert = lInsert & " Cod_Tip_Doc,"
    lInsert = lInsert & " Dat_Rec_Cia,"
    lInsert = lInsert & " Num_Ref_Corr,"
    lInsert = lInsert & " Dat_Emis,"
    lInsert = lInsert & " Num_Apol,"
    lInsert = lInsert & " Num_Endo,"
    lInsert = lInsert & " Val_Pr_Emis,"
    lInsert = lInsert & " Val_Pr_Liq,"
    lInsert = lInsert & " Per_Tot_Coss,"
    lInsert = lInsert & " Qtd_Parc,"
    lInsert = lInsert & " Cod_Iden_Stat_Resp,"
    lInsert = lInsert & " Cod_Iden_SAC,"
    lInsert = lInsert & " Cod_Prod,"
    lInsert = lInsert & " Val_IS,"
    lInsert = lInsert & " Per_Desc_Coml,"
    lInsert = lInsert & " Num_Cgc_Cpf,"
    lInsert = lInsert & " Cod_Iden_Pessoa,"
    lInsert = lInsert & " Cod_Iden_Pag_Antec,"
    lInsert = lInsert & " Dig_Bloqueto"


    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Bloq) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.COD_UNID) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ini_Vig) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ter_Vig) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Corr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corr_2) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Corr_2) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corr_3) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Corr_3) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Seg) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Anex) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol_Ant) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Obs_Prop) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Cotac) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cotac) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Juros_Neg) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Envi_Doc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.DAT_DIGIT) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.HOR_DIGIT) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Stat_Acao) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Digit) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Digit_MF) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Hor_Digit_MF) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Digit_MF) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Tot) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Classe) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Cobr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Prop) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Canc_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Doc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Rec_Cia) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ref_Corr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Emis) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Emis) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Liq) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Tot_Coss) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Parc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Stat_Resp) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_SAC) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Prod) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Coml) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cgc_Cpf) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Pessoa) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Pag_Antec) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.DIG_BLOQUETO) & ")"

    mfFormatarInclusaoSQL = lInsert
    GoTo Fim

IncluirCompl:


    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ctrl_Emis_Compl ( "
    lInsert = lInsert & " Nosso_Numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " NUM_PI,"
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " Num_Prop,"
    lInsert = lInsert & " Num_Apol,"
    lInsert = lInsert & " Num_Endo,"
    lInsert = lInsert & " Parcelas_Impressao_Avista,"
    lInsert = lInsert & " Parcelas_Impressao_30dias,"
    If pNomeTabImport <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lInsert = lInsert & " Parcelas_Impressao_AvistaDebito,"
        lInsert = lInsert & " Parcelas_Impressao_30diasDebito,"
        lInsert = lInsert & " Parcelas_Impressao_Cartao,"
    End If
    lInsert = lInsert & " Email_Contato,"
    lInsert = lInsert & " Ind_Impr_Apol,"
    lInsert = lInsert & " Ind_Prolabore,"
    lInsert = lInsert & " Ind_Agenciamento,"
    lInsert = lInsert & " Cod_Bco_Boleto,"
    lInsert = lInsert & " DAT_DIGIT,"
    lInsert = lInsert & " HOR_DIGIT, "
    lInsert = lInsert & " Per_Desc_Agrav, "
    lInsert = lInsert & " cod_origem, "
    lInsert = lInsert & " Num_Apol_Mar, "
    lInsert = lInsert & " Num_Logr_Segurado, "
    lInsert = lInsert & " Dsc_Compl_Segurado, "
    lInsert = lInsert & " Num_Logr_Cobranca, "
    lInsert = lInsert & " Dsc_Compl_Cobranca, "
    lInsert = lInsert & " ESTADO_CIVIL, "
    lInsert = lInsert & " DAT_NASCIMENTO,"
    'conta restitui��o credito
    lInsert = lInsert & " Tip_Res_Cre,"
    lInsert = lInsert & " Cod_Banco_Cre, "
    lInsert = lInsert & " Cod_Agenc_Cre, "
    lInsert = lInsert & " Dig_Agenc_Cre, "
    lInsert = lInsert & " Cod_Conta_Cre, "
    lInsert = lInsert & " Cpf_Cnpj_Cre,"
    lInsert = lInsert & " Ind_Cond_Constituido,"
    lInsert = lInsert & " Indice_IPC,"
    lInsert = lInsert & " Val_VerbaCom,"
    lInsert = lInsert & " Cod_DIG_Conta"
    
    If UCase(pPropTabela) = "SYASPORTAL.DBO." Then
        lInsert = lInsert & ", Ind_Promocional"
        lInsert = lInsert & ", Ind_Texto_Rel_Bens"
        End If
    lInsert = lInsert & ", Suprimir_Batente"

    If UCase(pPropTabela) <> "GED.DBO." And pNomeTabImport = "" Then
        lInsert = lInsert & " ,Ind_Bloq_Web"
        lInsert = lInsert & " ,Ind_Envio_Alcada"
        lInsert = lInsert & " ,per_corr_max"
        lInsert = lInsert & " ,Per_Max_Comis_Corr"
        lInsert = lInsert & " ,Dat_Ini_Vig_Orig"
        lInsert = lInsert & " ,Dat_Ter_Vig_Orig"
        lInsert = lInsert & " ,Tip_Classe"
        lInsert = lInsert & " ,Val_Pr_C4"
        lInsert = lInsert & " ,Num_Ped_Ant"
        lInsert = lInsert & " ,Num_Cotac_Ant"
        lInsert = lInsert & " ,RA_Comissao"
        lInsert = lInsert & " ,Per_Comis_Corr_RA"
    End If
    
    lInsert = lInsert & " ,Ind_Org_Publico"

    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Prop) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_Avista) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_30dias) & ","
    If pNomeTabImport <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_AvistaDebito) & ","
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_30diasDebito) & ","
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_Cartao) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Email_Contato) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Impr_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Prolabore) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Agenciamento) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Bco_Boleto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.DAT_DIGIT) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.HOR_DIGIT) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Agrav) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.cod_origem) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol_Mar) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Logr_Segurado) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Segurado) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Logr_Cobranca) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Cobranca) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Estado_Civil) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.DAT_Nascimento) & ","
    'conta restitui��o credito
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Res_Cre) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Banco_Cre) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Agenc_Cre) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dig_Agenc_Cre) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Conta_Cre) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cpf_Cnpj_Cre) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Cond_Constituido) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Indice_IPC) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_VerbaCom) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(UCase(pobjstc.Cod_DIG_Conta))
    
    If UCase(pPropTabela) = "SYASPORTAL.DBO." Then
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Promocional)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Texto_Rel_Bens)
    End If
        
    lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Suprimir_Batente)
    
    If UCase(pPropTabela) <> "GED.DBO." And pNomeTabImport = "" Then
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Bloq_Web)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Envio_Alcada)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.per_corr_max)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Per_Max_Comis_Corr)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ini_Vig_Orig)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ter_Vig_Orig)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Classe)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_C4)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped_Ant)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cotac_Ant)
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.RA_Comissao)
        
        'JIRA-14958 8000048511,SISCOTA END CO-CORRETAGEM
        'lUpdate = lUpdate & ",Per_Comis_Corr_RA =" & M46V999.gfFormatarCamposSQL(pobjstc.Per_Comis_Corr_RA)
        pobjstc.Per_Comis_Corr_RA = pobjstc.Per_Corr + pobjstc.Per_Corr_2 + pobjstc.Per_Corr_3
        
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Per_Comis_Corr_RA)
    End If
    
    lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Org_Publico)
    
    lInsert = lInsert & ")"
    
    mfFormatarInclusaoSQL = lInsert

Fim:

End Function


Private Function mfFormatarAlteracaoSQL(ByRef pobjstc As stcA46V712B, _
                                        ByVal pPropTabela As String, _
                                        ByVal pNomeTabImport As String, _
                                        ByVal TabAtual As Integer) As String

    Dim lUpdate As String                'Query.

    If TabAtual = 2 Then GoTo AtualizarCompl:

    lUpdate = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ctrl_Emis Set "
    lUpdate = lUpdate & " NUM_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lUpdate = lUpdate & " NUM_BLOQ = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Bloq) & ","
    lUpdate = lUpdate & " COD_UNID = " & M46V999.gfFormatarCamposSQL(pobjstc.COD_UNID) & ","
    lUpdate = lUpdate & " COD_RAMO = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lUpdate = lUpdate & " DAT_INI_VIG = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ini_Vig) & ","
    lUpdate = lUpdate & " DAT_TER_VIG = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ter_Vig) & ","
    lUpdate = lUpdate & " Cod_Corr = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corr) & ","
    lUpdate = lUpdate & " Per_Corr = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Corr) & ","
    lUpdate = lUpdate & " Cod_Corr_2 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corr_2) & ","
    lUpdate = lUpdate & " Per_Corr_2 = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Corr_2) & ","
    lUpdate = lUpdate & " Cod_Corr_3 = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corr_3) & ","
    lUpdate = lUpdate & " Per_Corr_3 = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Corr_3) & ","
    lUpdate = lUpdate & " Nom_Seg = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Seg) & ","
    lUpdate = lUpdate & " Qtd_Anex = " & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Anex) & ","
    lUpdate = lUpdate & " Num_Apol_Ant = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol_Ant) & ","
    lUpdate = lUpdate & " Dsc_Obs_Prop = " & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Obs_Prop) & ","
    lUpdate = lUpdate & " Tip_Cotac = " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Cotac) & ","
    lUpdate = lUpdate & " Num_Cotac = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cotac) & ","
    lUpdate = lUpdate & " Per_Juros_Neg = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Juros_Neg) & ","
    lUpdate = lUpdate & " Cod_Envi_Doc = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Envi_Doc) & ","
    lUpdate = lUpdate & " Dat_Digit = " & M46V999.gfFormatarCamposSQL(pobjstc.DAT_DIGIT) & ","
    lUpdate = lUpdate & " Hor_Digit = " & M46V999.gfFormatarCamposSQL(pobjstc.HOR_DIGIT) & ","
    lUpdate = lUpdate & " Cod_Iden_Stat_Acao = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Stat_Acao) & ","
    lUpdate = lUpdate & " Cod_User_Digit = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Digit) & ","
    lUpdate = lUpdate & " Dat_Digit_MF = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Digit_MF) & ","
    lUpdate = lUpdate & " Hor_Digit_MF = " & M46V999.gfFormatarCamposSQL(pobjstc.Hor_Digit_MF) & ","
    lUpdate = lUpdate & " Cod_User_Digit_MF = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Digit_MF) & ","
    lUpdate = lUpdate & " Val_Pr_Tot = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Tot) & ","
    lUpdate = lUpdate & " Cod_Iden_Classe = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Classe) & ","
    lUpdate = lUpdate & " Cod_Iden_Cobr = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Cobr) & ","
    lUpdate = lUpdate & " Num_Prop = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Prop) & ","
    lUpdate = lUpdate & " Dat_Canc_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Canc_Apol) & ","
    lUpdate = lUpdate & " Cod_Tip_Doc = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Doc) & ","
    lUpdate = lUpdate & " Dat_Rec_Cia = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Rec_Cia) & ","
    lUpdate = lUpdate & " Num_Ref_Corr = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ref_Corr) & ","
    lUpdate = lUpdate & " Dat_Emis = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Emis) & ","
    lUpdate = lUpdate & " Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lUpdate = lUpdate & " Num_Endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lUpdate = lUpdate & " Val_Pr_Emis = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Emis) & ","
    lUpdate = lUpdate & " Val_Pr_Liq = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Liq) & ","
    lUpdate = lUpdate & " Per_Tot_Coss = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Tot_Coss) & ","
    lUpdate = lUpdate & " Qtd_Parc = " & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Parc) & ","
    lUpdate = lUpdate & " Cod_Iden_Stat_Resp = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Stat_Resp) & ","
    lUpdate = lUpdate & " Cod_Iden_SAC = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_SAC) & ","
    lUpdate = lUpdate & " Cod_Prod = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Prod) & ","
    lUpdate = lUpdate & " Val_IS = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS) & ","
    lUpdate = lUpdate & " Per_Desc_Coml = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Coml) & ","
    lUpdate = lUpdate & " Num_Cgc_Cpf = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cgc_Cpf) & ","
    lUpdate = lUpdate & " Cod_Iden_Pessoa = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Pessoa) & ","
    lUpdate = lUpdate & " Dig_Bloqueto= " & M46V999.gfFormatarCamposSQL(pobjstc.DIG_BLOQUETO) & ","
    lUpdate = lUpdate & " Cod_Iden_Pag_Antec = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Pag_Antec)

    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If


    mfFormatarAlteracaoSQL = lUpdate
    GoTo Fim

AtualizarCompl:

    'ProjetoFormularios

    lUpdate = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ctrl_Emis_Compl Set "
    lUpdate = lUpdate & " NUM_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lUpdate = lUpdate & " COD_RAMO = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lUpdate = lUpdate & " Num_Prop = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Prop) & ","
    lUpdate = lUpdate & " Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lUpdate = lUpdate & " Num_Endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","

    lUpdate = lUpdate & " Parcelas_Impressao_Avista = " & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_Avista) & ","
    lUpdate = lUpdate & " Parcelas_Impressao_30dias = " & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_30dias) & ","
    If pNomeTabImport <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lUpdate = lUpdate & " Parcelas_Impressao_AvistaDebito = " & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_AvistaDebito) & ","
        lUpdate = lUpdate & " Parcelas_Impressao_30diasDebito = " & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_30diasDebito) & ","
        lUpdate = lUpdate & " Parcelas_Impressao_Cartao = " & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_Cartao) & ","
    End If
    lUpdate = lUpdate & " Email_Contato = " & M46V999.gfFormatarCamposSQL(pobjstc.Email_Contato) & ","
    lUpdate = lUpdate & " Ind_Impr_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Impr_Apol) & ","
    lUpdate = lUpdate & " Ind_Prolabore = " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Prolabore) & ","
    lUpdate = lUpdate & " Ind_Agenciamento = " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Agenciamento) & ","
    lUpdate = lUpdate & " Cod_Bco_Boleto = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Bco_Boleto) & ","
    lUpdate = lUpdate & " Dat_Digit = " & M46V999.gfFormatarCamposSQL(pobjstc.DAT_DIGIT) & ","
    lUpdate = lUpdate & " Hor_Digit = " & M46V999.gfFormatarCamposSQL(pobjstc.HOR_DIGIT) & ","
    lUpdate = lUpdate & " Per_Desc_Agrav = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Agrav) & ","
    lUpdate = lUpdate & " cod_origem = " & M46V999.gfFormatarCamposSQL(pobjstc.cod_origem) & ","
    lUpdate = lUpdate & " Num_Apol_Mar = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol_Mar) & ","
    lUpdate = lUpdate & " Num_Logr_Segurado = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Logr_Segurado) & ","
    lUpdate = lUpdate & " Dsc_Compl_Segurado = " & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Segurado) & ","
    lUpdate = lUpdate & " Num_Logr_Cobranca = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Logr_Cobranca) & ","
    lUpdate = lUpdate & " Dsc_Compl_Cobranca = " & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Cobranca) & ","
    lUpdate = lUpdate & " Estado_Civil = " & M46V999.gfFormatarCamposSQL(pobjstc.Estado_Civil) & ","
    lUpdate = lUpdate & " DAT_Nascimento = " & M46V999.gfFormatarCamposSQL(pobjstc.DAT_Nascimento) & ","
    'credito restitui��o
    lUpdate = lUpdate & " Tip_Res_Cre = " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Res_Cre) & ","
    lUpdate = lUpdate & " Cod_Banco_Cre = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Banco_Cre) & ","
    lUpdate = lUpdate & " Cod_Agenc_Cre = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Agenc_Cre) & ","
    lUpdate = lUpdate & " Dig_Agenc_Cre = " & M46V999.gfFormatarCamposSQL(pobjstc.Dig_Agenc_Cre) & ","
    lUpdate = lUpdate & " Cod_Conta_Cre = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Conta_Cre) & ","
    lUpdate = lUpdate & " Cpf_Cnpj_Cre = " & M46V999.gfFormatarCamposSQL(pobjstc.Cpf_Cnpj_Cre) & ","
    lUpdate = lUpdate & " Ind_Cond_Constituido = " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Cond_Constituido) & ","
    lUpdate = lUpdate & " Indice_IPC = " & M46V999.gfFormatarCamposSQL(pobjstc.Indice_IPC) & ","
    lUpdate = lUpdate & " Val_VerbaCom = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_VerbaCom) & ","
    lUpdate = lUpdate & " Cod_DIG_Conta = " & M46V999.gfFormatarCamposSQL(UCase(pobjstc.Cod_DIG_Conta))
    lUpdate = lUpdate & " ,Suprimir_Batente = " & M46V999.gfFormatarCamposSQL(pobjstc.Suprimir_Batente)
    
    If pPropTabela <> "GED.DBO" And pNomeTabImport = "" Then
        lUpdate = lUpdate & ",Ind_Bloq_Web = " & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Bloq_Web)
        lUpdate = lUpdate & ",Ind_Envio_Alcada =" & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Envio_Alcada)
        lUpdate = lUpdate & ",per_corr_max =" & M46V999.gfFormatarCamposSQL(pobjstc.per_corr_max)
        lUpdate = lUpdate & ",Per_Max_Comis_Corr =" & M46V999.gfFormatarCamposSQL(pobjstc.Per_Max_Comis_Corr)
        lUpdate = lUpdate & ",Dat_Ini_Vig_Orig =" & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ini_Vig_Orig)
        lUpdate = lUpdate & ",Dat_Ter_Vig_Orig =" & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ter_Vig_Orig)
        lUpdate = lUpdate & ",Tip_Classe =" & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Classe)
        lUpdate = lUpdate & ",Val_Pr_C4 =" & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_C4)
        lUpdate = lUpdate & ",Num_Ped_Ant =" & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped_Ant)
        lUpdate = lUpdate & ",Num_Cotac_Ant =" & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cotac_Ant)
        lUpdate = lUpdate & ",RA_Comissao =" & M46V999.gfFormatarCamposSQL(pobjstc.RA_Comissao)
        
        'JIRA-14958 8000048511,SISCOTA END CO-CORRETAGEM
        'lUpdate = lUpdate & ",Per_Comis_Corr_RA =" & M46V999.gfFormatarCamposSQL(pobjstc.Per_Comis_Corr_RA)
        pobjstc.Per_Comis_Corr_RA = pobjstc.Per_Corr + pobjstc.Per_Corr_2 + pobjstc.Per_Corr_3
        
        lUpdate = lUpdate & ",Per_Comis_Corr_RA =" & M46V999.gfFormatarCamposSQL(pobjstc.Per_Comis_Corr_RA)
    End If
    
    lUpdate = lUpdate & ",Ind_Org_Publico =" & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Org_Publico)
    
    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarAlteracaoSQL = lUpdate

Fim:

End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V712B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       ByVal TabAtual As Integer) As String
    Dim lDelete As String                'Query.

    If TabAtual = 2 Then GoTo ExcluirCompl:

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ctrl_Emis WHERE "
    lDelete = lDelete & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarExclusaoSQL = lDelete

    GoTo Fim
ExcluirCompl:

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ctrl_Emis_Compl WHERE "
    lDelete = lDelete & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarExclusaoSQL = lDelete

Fim:
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V712B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String


    lSelect = "SELECT E.* "
    'ProjetoFormularios
    lSelect = lSelect & " , C.Parcelas_Impressao_Avista, C.Parcelas_Impressao_30dias, C.Email_Contato, C.Ind_Impr_Apol "
    lSelect = lSelect & " , C.Ind_Prolabore, C.Ind_Agenciamento, C.Cod_Bco_Boleto, c.Per_Desc_Agrav, c.cod_origem, c.Num_Apol_Mar, c.Num_Logr_Segurado, c.Dsc_Compl_Segurado, c.Num_Logr_Cobranca, c.Dsc_Compl_Cobranca, c.Estado_Civil, c.DAT_Nascimento "
    lSelect = lSelect & " , isnull(c.Tip_Res_Cre,0) Tip_Res_Cre, c.Cod_Banco_Cre, c.Cod_Agenc_Cre, c.Dig_Agenc_Cre, c.Cod_Conta_Cre, c.Cpf_Cnpj_Cre, c.Ind_Cond_Constituido, c.Indice_IPC, c.Val_VerbaCom, UPPER(c.Cod_DIG_Conta) Cod_DIG_Conta "
    If UCase(pPropTabela) Like "*SYASPORTAL*" Then
        lSelect = lSelect & ", c.Ind_Promocional, c.Ind_Texto_Rel_Bens"
        End If
        
    lSelect = lSelect & ", c.Suprimir_Batente"

    If pPropTabela <> "GED.DBO" And pNomeTabImport = "" Then
        lSelect = lSelect & ",Ind_Bloq_Web,Ind_Envio_Alcada,c.per_corr_max, c.Per_Max_Comis_Corr, c.Dat_Ini_Vig_Orig, c.Dat_Ter_Vig_Orig, c.Tip_Classe, c.Val_Pr_C4, c.Num_Ped_Ant, c.Num_Cotac_Ant, c.RA_Comissao, c.Per_Comis_Corr_RA "
    End If
    
    If pNomeTabImport <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lSelect = lSelect & ", C.Parcelas_Impressao_AvistaDebito "
        lSelect = lSelect & ", C.Parcelas_Impressao_30diasDebito "
        lSelect = lSelect & ", C.Parcelas_Impressao_Cartao "
    End If

        lSelect = lSelect & ",c.Ind_Org_Publico "

    lSelect = lSelect & " FROM "
    lSelect = lSelect & pPropTabela & pNomeTabImport & "Tab_Ctrl_Emis As E  with(nolock)"

    'ProjetoFormularios
    lSelect = lSelect & " Left Join " & pPropTabela & pNomeTabImport & "Tab_Ctrl_Emis_Compl as C with(nolock) "

    If pPropTabela = "GED.DBO." Then
        lSelect = lSelect & " on E.Num_PI  =  C.Num_PI "
    Else
        lSelect = lSelect & " on E.Nosso_Numero  =  C.Nosso_Numero "
    End If

    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lSelect = lSelect & " Where E.Nosso_numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " E.Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If val(pobjstc.NUM_PI) <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " E.Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)
    End If
    If pobjstc.Num_Apol <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " E.Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol)
        lSelect = M46V999.gfWhereAnd(lSelect) & " E.num_endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo)
    End If

    mfFormatarConsultaSQL = lSelect
End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V712B, prst As clsYasRecordSet)

    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("NUM_PI"), tpoNumerico)
    pobjstc.Num_Bloq = M46V999.gfTratarNulo(prst("NUM_BLOQ"), tpoNumerico)
    pobjstc.COD_UNID = M46V999.gfTratarNulo(prst("COD_UNID"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("COD_RAMO"), tpoNumerico)
    pobjstc.Dat_Ini_Vig = M46V999.gfTratarNulo(prst("DAT_INI_VIG"), tpoNumerico)
    pobjstc.Dat_Ter_Vig = M46V999.gfTratarNulo(prst("DAT_TER_VIG"), tpoNumerico)
    pobjstc.Cod_Corr = M46V999.gfTratarNulo(prst("Cod_Corr"), tpoNumerico)
    pobjstc.Per_Corr = M46V999.gfTratarNulo(prst("Per_Corr"), tpoNumerico)
    pobjstc.Cod_Corr_2 = M46V999.gfTratarNulo(prst("Cod_Corr_2"), tpoNumerico)
    pobjstc.Per_Corr_2 = M46V999.gfTratarNulo(prst("Per_Corr_2"), tpoNumerico)
    pobjstc.Cod_Corr_3 = M46V999.gfTratarNulo(prst("Cod_Corr_3"), tpoNumerico)
    pobjstc.Per_Corr_3 = M46V999.gfTratarNulo(prst("Per_Corr_3"), tpoNumerico)
    pobjstc.Nom_Seg = M46V999.gfTratarNulo(prst("Nom_Seg"), tpoString)
    pobjstc.Qtd_Anex = M46V999.gfTratarNulo(prst("Qtd_Anex"), tpoNumerico)
    pobjstc.Num_Apol_Ant = M46V999.gfTratarNulo(prst("Num_Apol_Ant"), tpoNumerico)
    pobjstc.Dsc_Obs_Prop = M46V999.gfTratarNulo(prst("Dsc_Obs_Prop"), tpoString)
    pobjstc.Tip_Cotac = M46V999.gfTratarNulo(prst("Tip_Cotac"), tpoString)
    pobjstc.Num_Cotac = M46V999.gfTratarNulo(prst("Num_Cotac"), tpoString)
    pobjstc.Per_Juros_Neg = M46V999.gfTratarNulo(prst("Per_Juros_Neg"), tpoNumerico)
    pobjstc.Cod_Envi_Doc = M46V999.gfTratarNulo(prst("Cod_Envi_Doc"), tpoNumerico)
    pobjstc.DAT_DIGIT = M46V999.gfTratarNulo(prst("Dat_Digit"), tpoNumerico)
    pobjstc.HOR_DIGIT = M46V999.gfTratarNulo(prst("Hor_Digit"), tpoNumerico)
    pobjstc.Cod_Iden_Stat_Acao = M46V999.gfTratarNulo(prst("Cod_Iden_Stat_Acao"), tpoNumerico)
    pobjstc.Cod_User_Digit = M46V999.gfTratarNulo(prst("Cod_User_Digit"), tpoString)
    pobjstc.Dat_Digit_MF = M46V999.gfTratarNulo(prst("Dat_Digit_MF"), tpoNumerico)
    pobjstc.Hor_Digit_MF = M46V999.gfTratarNulo(prst("Hor_Digit_MF"), tpoNumerico)
    pobjstc.Cod_User_Digit_MF = M46V999.gfTratarNulo(prst("Cod_User_Digit_MF"), tpoString)
    pobjstc.Val_Pr_Tot = M46V999.gfTratarNulo(prst("Val_Pr_Tot"), tpoNumerico)
    pobjstc.Cod_Iden_Classe = M46V999.gfTratarNulo(prst("Cod_Iden_Classe"), tpoNumerico)
    pobjstc.Cod_Iden_Cobr = M46V999.gfTratarNulo(prst("Cod_Iden_Cobr"), tpoNumerico)
    pobjstc.Num_Prop = M46V999.gfTratarNulo(prst("Num_Prop"), tpoNumerico)
    pobjstc.Dat_Canc_Apol = M46V999.gfTratarNulo(prst("Dat_Canc_Apol"), tpoNumerico)
    pobjstc.Cod_Tip_Doc = M46V999.gfTratarNulo(prst("Cod_Tip_Doc"), tpoNumerico)
    pobjstc.Dat_Rec_Cia = M46V999.gfTratarNulo(prst("Dat_Rec_Cia"), tpoNumerico)
    pobjstc.Num_Ref_Corr = M46V999.gfTratarNulo(prst("Num_Ref_Corr"), tpoString)
    pobjstc.Dat_Emis = M46V999.gfTratarNulo(prst("Dat_Emis"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Val_Pr_Emis = M46V999.gfTratarNulo(prst("Val_Pr_Emis"), tpoNumerico)
    pobjstc.Val_Pr_Liq = M46V999.gfTratarNulo(prst("Val_Pr_Liq"), tpoNumerico)
    pobjstc.Per_Tot_Coss = M46V999.gfTratarNulo(prst("Per_Tot_Coss"), tpoNumerico)
    pobjstc.Qtd_Parc = M46V999.gfTratarNulo(prst("Qtd_Parc"), tpoNumerico)
    pobjstc.Cod_Iden_Stat_Resp = M46V999.gfTratarNulo(prst("Cod_Iden_Stat_Resp"), tpoNumerico)
    pobjstc.Cod_Iden_SAC = M46V999.gfTratarNulo(prst("Cod_Iden_SAC"), tpoNumerico)
    pobjstc.Cod_Prod = M46V999.gfTratarNulo(prst("Cod_Prod"), tpoNumerico)
    pobjstc.Val_IS = M46V999.gfTratarNulo(prst("Val_IS"), tpoNumerico)
    pobjstc.Per_Desc_Coml = M46V999.gfTratarNulo(prst("Per_Desc_Coml"), tpoNumerico)
    pobjstc.Num_Cgc_Cpf = M46V999.gfTratarNulo(prst("Num_Cgc_Cpf"), tpoString)
    pobjstc.Cod_Iden_Pessoa = M46V999.gfTratarNulo(prst("Cod_Iden_Pessoa"), tpoNumerico)
    pobjstc.Cod_Iden_Pag_Antec = M46V999.gfTratarNulo(prst("Cod_Iden_Pag_Antec"), tpoNumerico)
    pobjstc.DIG_BLOQUETO = M46V999.gfTratarNulo(prst("Dig_Bloqueto"), tpoNumerico)

    On Error Resume Next
    pobjstc.Parcelas_Impressao_Avista = M46V999.gfTratarNulo(prst("Parcelas_Impressao_Avista"), tpoString)
    pobjstc.Parcelas_Impressao_30dias = M46V999.gfTratarNulo(prst("Parcelas_Impressao_30dias"), tpoString)
    pobjstc.Parcelas_Impressao_AvistaDebito = M46V999.gfTratarNulo(prst("Parcelas_Impressao_AvistaDebito"), tpoString)
    pobjstc.Parcelas_Impressao_30diasDebito = M46V999.gfTratarNulo(prst("Parcelas_Impressao_30diasDebito"), tpoString)
    pobjstc.Parcelas_Impressao_Cartao = M46V999.gfTratarNulo(prst("Parcelas_Impressao_Cartao"), tpoString)
    pobjstc.Email_Contato = M46V999.gfTratarNulo(prst("Email_Contato"), tpoString)
    pobjstc.Ind_Impr_Apol = M46V999.gfTratarNulo(prst("Ind_Impr_Apol"), tpoNumerico)
    pobjstc.Ind_Prolabore = M46V999.gfTratarNulo(prst("Ind_Prolabore"), tpoNumerico)
    pobjstc.Ind_Agenciamento = M46V999.gfTratarNulo(prst("Ind_Agenciamento"), tpoNumerico)
    pobjstc.Cod_Bco_Boleto = M46V999.gfTratarNulo(prst("Cod_Bco_Boleto"), tpoNumerico)
    pobjstc.Per_Desc_Agrav = M46V999.gfTratarNulo(prst("Per_Desc_Agrav"), tpoNumerico)
    pobjstc.cod_origem = M46V999.gfTratarNulo(prst("cod_origem"), tpoNumerico)
    pobjstc.Num_Apol_Mar = M46V999.gfTratarNulo(prst("Num_Apol_Mar"), tpoNumerico)
    pobjstc.Num_Logr_Segurado = M46V999.gfTratarNulo(prst("Num_Logr_Segurado"), tpoNumerico)
    pobjstc.Dsc_Compl_Segurado = M46V999.gfTratarNulo(prst("Dsc_Compl_Segurado"), tpoNumerico)
    pobjstc.Num_Logr_Cobranca = M46V999.gfTratarNulo(prst("Num_Logr_Cobranca"), tpoNumerico)
    pobjstc.Dsc_Compl_Cobranca = M46V999.gfTratarNulo(prst("Dsc_Compl_Cobranca"), tpoNumerico)
    pobjstc.Estado_Civil = M46V999.gfTratarNulo(prst("Estado_Civil"), tpoNumerico)
    pobjstc.DAT_Nascimento = M46V999.gfTratarNulo(prst("DAT_Nascimento"), tpoNumerico)
    'restitui��o credito
    pobjstc.Tip_Res_Cre = M46V999.gfTratarNulo(prst("Tip_Res_Cre"), tpoNumerico)
    pobjstc.Cod_Banco_Cre = M46V999.gfTratarNulo(prst("Cod_Banco_Cre"), tpoNumerico)
    pobjstc.Cod_Agenc_Cre = M46V999.gfTratarNulo(prst("Cod_Agenc_Cre"), tpoNumerico)
    pobjstc.Dig_Agenc_Cre = M46V999.gfTratarNulo(prst("Dig_Agenc_Cre"), tpoNumerico)
    pobjstc.Cod_Conta_Cre = M46V999.gfTratarNulo(prst("Cod_Conta_Cre"), tpoNumerico)
    pobjstc.Cpf_Cnpj_Cre = M46V999.gfTratarNulo(prst("Cpf_Cnpj_Cre"), tpoString)
    pobjstc.Ind_Cond_Constituido = M46V999.gfTratarNulo(prst("Ind_Cond_Constituido"), tpoNumerico)
    pobjstc.Indice_IPC = M46V999.gfTratarNulo(prst("Indice_IPC"), tpoNumerico)
    pobjstc.Val_VerbaCom = M46V999.gfTratarNulo(prst("Val_VerbaCom"), TpoNumDecimal)
    pobjstc.Cod_DIG_Conta = M46V999.gfTratarNulo(prst("Cod_DIG_Conta"), tpoString)
    pobjstc.Cod_DIG_Conta = UCase(pobjstc.Cod_DIG_Conta)
    pobjstc.Ind_Promocional = M46V999.gfTratarNulo(prst("Ind_Promocional"), tpoNumerico)
    pobjstc.Ind_Texto_Rel_Bens = M46V999.gfTratarNulo(prst("Ind_Texto_Rel_Bens"), tpoNumerico)
    pobjstc.Suprimir_Batente = False 'Valor Padr�o
    pobjstc.Suprimir_Batente = M46V999.gfTratarNulo(prst("Suprimir_Batente"), TpoBoolean)
    pobjstc.Ind_Bloq_Web = M46V999.gfTratarNulo(prst("Ind_Bloq_Web"), tpoNumerico)
    pobjstc.Ind_Envio_Alcada = M46V999.gfTratarNulo(prst("Ind_Envio_Alcada"), tpoNumerico)
    pobjstc.per_corr_max = M46V999.gfTratarNulo(prst("per_corr_max"), TpoNumDecimal)
    pobjstc.Per_Max_Comis_Corr = M46V999.gfTratarNulo(prst("Per_Max_Comis_Corr"), tpoNumerico)
    pobjstc.Dat_Ini_Vig_Orig = M46V999.gfTratarNulo(prst("Dat_Ini_Vig_Orig"), tpoNumerico)
    pobjstc.Dat_Ter_Vig_Orig = M46V999.gfTratarNulo(prst("Dat_Ter_Vig_Orig"), tpoNumerico)
    pobjstc.Tip_Classe = M46V999.gfTratarNulo(prst("Tip_Classe"), tpoNumerico)
    pobjstc.Val_Pr_C4 = M46V999.gfTratarNulo(prst("Val_Pr_C4"), tpoNumerico)
    pobjstc.Num_Ped_Ant = M46V999.gfTratarNulo(prst("Num_Ped_Ant"), tpoNumerico)
    pobjstc.Num_Cotac_Ant = M46V999.gfTratarNulo(prst("Num_Cotac_Ant"), tpoString)
    pobjstc.RA_Comissao = M46V999.gfTratarNulo(prst("RA_Comissao"), tpoString)
    pobjstc.Per_Comis_Corr_RA = M46V999.gfTratarNulo(prst("Per_Comis_Corr_RA"), tpoNumerico)
    pobjstc.Ind_Org_Publico = M46V999.gfTratarNulo(prst("Ind_Org_Publico"), TpoBoolean)
    
    Err = 0

End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V712B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean
    Dim lsSQl As String           'Query.

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfIncluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    On Local Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros

    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, 1)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros

    'ProjetoFormularios
    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, 2)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros

    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfIncluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErros:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V712A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V712A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function

Public Function mfAlterar(ByRef pobjstc As stcA46V712B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                          Optional pDataUltAlt As Double, _
                          Optional pUltUsuAlt As String) As Boolean
    Dim lsSQl As String           'Query.
    Dim lQtdRegAfetados As Long
    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros

    lsSQl = mfFormatarAlteracaoSQL(pobjstc, pPropTabela, pNomeTabImport, 1)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros

    'ProjetoFormularios
    lsSQl = mfFormatarAlteracaoSQL(pobjstc, pPropTabela, pNomeTabImport, 2)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro, lQtdRegAfetados) <> 0 Then GoTo ControleErros
        
    If lQtdRegAfetados = 0 Then
        lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, 2)
        If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
    End If

    If Colecao.Count > 0 Then
        Colecao.Remove (pobjstc.Class_ID)
    End If
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfAlterar = True
    Exit Function
ControleErros:
    mfAlterar = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V712A.mfAlterar ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V712A.mfAlterar ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfAlterarNum_PI(ByRef pobjstc As stcA46V712B, _
                                ByVal pPropTabela As String, _
                                ByVal pNomeTabImport As String) As Boolean
    Dim lUpdate As String           'Query.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros

    lUpdate = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ctrl_Emis Set "
    lUpdate = lUpdate & " num_pi = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)
    lUpdate = lUpdate & ", dat_rec_cia = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Rec_Cia)
    lUpdate = lUpdate & ", cod_iden_stat_acao = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Stat_Acao)
    lUpdate = lUpdate & ", Num_Ref_Corr = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ref_Corr)
    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErros
    If Colecao.Count > 0 Then
        Colecao.Remove (pobjstc.Class_ID)
    End If
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfAlterarNum_PI = True
    Exit Function
ControleErros:
    mfAlterarNum_PI = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V712A.mfAlterar ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lUpdate)
    Else
        Call M46V999.gpGraLog(0, "A46V712A.mfAlterar ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lUpdate)
    End If
End Function


Public Function mfAlterarDescontoComercial(ByRef pobjstc As stcA46V712B, _
                                           ByVal pPropTabela As String, _
                                           ByVal pNomeTabImport As String) As Boolean
    Dim lUpdate As String
    Dim lRegAlterados As Long

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros

    lUpdate = " Update " & pPropTabela & pNomeTabImport & "Tab_Ctrl_Emis set Per_Desc_Coml = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Coml)
    lUpdate = lUpdate & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErros

    mfAlterarDescontoComercial = True

    Exit Function
ControleErros:
    mfAlterarDescontoComercial = False
    UltimoErro = "A46V704A.mfAlterarDescontoComercial ERRO - " & Err & " " & Err.Description & _
                 vbCrLf & " SQL: " & lUpdate
End Function

Public Function mfAlterarStringParcelamento(ByRef pobjstc As stcA46V712B, _
                                            ByVal pPropTabela As String, _
                                            ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfAlterarStringParcelamento"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lUpdate As String
    Dim lRegAlterados As Long

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros

    lUpdate = " Update " & pPropTabela & pNomeTabImport & "Tab_Ctrl_Emis_Compl set "
    lUpdate = lUpdate & "  Parcelas_Impressao_Avista = " & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_Avista)
    lUpdate = lUpdate & ", Parcelas_Impressao_30dias = " & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_30dias)
    lUpdate = lUpdate & ", Parcelas_Impressao_AvistaDebito = " & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_AvistaDebito)
    lUpdate = lUpdate & ", Parcelas_Impressao_30diasDebito = " & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_30diasDebito)
    lUpdate = lUpdate & ", Parcelas_Impressao_Cartao = " & M46V999.gfFormatarCamposSQL(pobjstc.Parcelas_Impressao_Cartao)
    lUpdate = lUpdate & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErros

    mfAlterarStringParcelamento = True

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    
    Exit Function
ControleErros:
    mfAlterarStringParcelamento = False
    UltimoErro = "A46V712A.mfAlterarStringParcelamento ERRO - " & Err & " " & Err.Description & _
                 vbCrLf & " SQL: " & lUpdate
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)

End Function


Public Function mfExcluir(ByRef pobjstc As stcA46V712B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfExcluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lsSQl As String           'Query.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, 1)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros

    'ProjetoFormularios
    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, 2)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros

    If Colecao.Count > 0 Then
        Colecao.Remove (pobjstc.Class_ID)
    End If
    Set pobjstc = Nothing
    mfExcluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErros:
    mfExcluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V712A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V712A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V712B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V712B          'Objeto da classe de estrutura stcA46V712B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErros
    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V712B
        Call mpCarregarAtributos(lobjstc, lrst)
        Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID
        lrst.MoveNext
    Loop
    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecao = True
    Exit Function
ControleErros:
    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V712A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V712A.mfCarregarColecao ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Private Sub Class_Initialize()
    Set Colecao = New Collection
    Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
    Set Colecao = Nothing
    Set mClsM46V999 = Nothing
End Sub

Public Function mfConectar() As Boolean
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
