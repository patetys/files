VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V730A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46v730A
' Data/Hora: 20/11/2008 - 10:20:00
' Sistema: P46V101
' Analista: Sandra M. De Rosa
' ======================================================================================================================================================

Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46v730B
Public UltimoErro As String    ' Descrição do último erro gerado
'Private mNomeArquivoUpdate      As String
'Private mFile                   As FileSystemObject
'Private mArquivo                As TextStream

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V730A"
End Property

Private Function mfFormatarConsultaSQL(pobjstc As stcA46V730B) As String

    Dim strTMP As String

    strTMP = "Select * from Tab_Trans_Insp  with(nolock) where Dat_Trans >= " & pobjstc.Dat_Trans
    strTMP = strTMP & " and Dat_Trans <= " & pobjstc.Dat_Trans

    mfFormatarConsultaSQL = strTMP

End Function

Private Function mfFormatarInclusaoSQL(pobjstc As stcA46V730B) As String

    Dim strTMP As String

    strTMP = " INSERT INTO Tab_Trans_Insp ( "
    strTMP = strTMP & " NOSSO_NUMERO,"
    strTMP = strTMP & " NUM_PED,"
    strTMP = strTMP & " NUM_ITEM,"
    strTMP = strTMP & " NUM_APOL_ANT,"
    strTMP = strTMP & " NOM_SEGURADO,"
    strTMP = strTMP & " NUM_TEL,"
    strTMP = strTMP & " NUM_CGC_CPF,"
    strTMP = strTMP & " COD_RAMO,"
    strTMP = strTMP & " COD_ORG,"
    strTMP = strTMP & " DAT_INI_VIG,"
    strTMP = strTMP & " DAT_TER_VIG,"
    strTMP = strTMP & " COD_TIP_LOGR_RISCO,"
    strTMP = strTMP & " NOM_LOGR_RISCO,"
    strTMP = strTMP & " DSC_COMPL_RISCO,"
    strTMP = strTMP & " NOM_BAIRRO_RISCO,"
    strTMP = strTMP & " NOM_CID_RISCO,"
    strTMP = strTMP & " SIG_UF_RISCO,"
    strTMP = strTMP & " NUM_CEP_RISCO,"
    strTMP = strTMP & " NUM_RISCO,"
    strTMP = strTMP & " DAT_TRANS,"
    strTMP = strTMP & " COD_IDEN_PESSOA,"
    strTMP = strTMP & " COD_UNID_PROD,"
    strTMP = strTMP & " COD_PROD,"
    strTMP = strTMP & " COD_CORR,"
    strTMP = strTMP & " COD_COBERT_01,"
    strTMP = strTMP & " COD_COBERT_02,"
    strTMP = strTMP & " COD_COBERT_03,"
    strTMP = strTMP & " COD_COBERT_04,"
    strTMP = strTMP & " COD_COBERT_05,"
    strTMP = strTMP & " COD_COBERT_06,"
    strTMP = strTMP & " COD_COBERT_07,"
    strTMP = strTMP & " COD_COBERT_08,"
    strTMP = strTMP & " COD_COBERT_09,"
    strTMP = strTMP & " COD_COBERT_10,"
    strTMP = strTMP & " VAL_IS_COBERT_01,"
    strTMP = strTMP & " VAL_IS_COBERT_02,"
    strTMP = strTMP & " VAL_IS_COBERT_03,"
    strTMP = strTMP & " VAL_IS_COBERT_04,"
    strTMP = strTMP & " VAL_IS_COBERT_05,"
    strTMP = strTMP & " VAL_IS_COBERT_06,"
    strTMP = strTMP & " VAL_IS_COBERT_07,"
    strTMP = strTMP & " VAL_IS_COBERT_08,"
    strTMP = strTMP & " VAL_IS_COBERT_09,"
    strTMP = strTMP & " VAL_IS_COBERT_10,"
    strTMP = strTMP & " COD_CLASF,"
    strTMP = strTMP & " COD_COMPL_CLASF,"
    strTMP = strTMP & " PESS_CONT,"
    strTMP = strTMP & " COD_PRODUTO,"
    strTMP = strTMP & " FLAG_TRANS,"
    strTMP = strTMP & " NUM_COTAC"
    strTMP = strTMP & ", NUM_INSP,"
    strTMP = strTMP & " COD_STATUS_I,"
    strTMP = strTMP & " DAT_CONFIRM,"
    strTMP = strTMP & " DAT_FINAL,"
    strTMP = strTMP & " DAT_CANCEL"
    strTMP = strTMP & ") VALUES ("
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol_Ant) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Segurado) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Tel) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cgc_Cpf) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ini_Vig) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ter_Vig) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Logr_Risco) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Logr_Risco) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Risco) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Bairro_Risco) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Cid_Risco) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Sig_UF_Risco) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cep_Risco) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Risco) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Trans) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Pessoa) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unid_Prod) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Prod) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corr) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert_01) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert_02) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert_03) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert_04) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert_05) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert_06) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert_07) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert_08) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert_09) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert_10) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS_Cobert_01) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS_Cobert_02) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS_Cobert_03) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS_Cobert_04) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS_Cobert_05) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS_Cobert_06) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS_Cobert_07) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS_Cobert_08) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS_Cobert_09) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS_Cobert_10) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Clasf) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Compl_Clasf) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Pess_Cont) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Produto) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Flag_Trans) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cotac) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Insp) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Status_I) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Confirm) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Final) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Cancel) & ")"

    mfFormatarInclusaoSQL = strTMP

End Function

Private Function mfFormatarExclusaoSQL(pobjstc As stcA46V730B) As String

    Dim strTMP As String

    strTMP = " DELETE from Tab_Trans_Insp "
    strTMP = strTMP & " where Nosso_Numero = "
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    strTMP = strTMP & " and Num_Ped = "
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped)
    strTMP = strTMP & " and Num_Item = "
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    strTMP = strTMP & " AND FLAG_TRANS = '0'"

    mfFormatarExclusaoSQL = strTMP

End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Public Function gfIncluir(pobjstc As stcA46V730B) As Boolean

    Dim lConexaoSQL As ADODB.Connection
    Dim sSql As String
    Dim pMensagem As String
    Dim lobjstcTransCobInsp As stcA46V731B    'Objeto da classe de estrutura stcA46V731B (TAB_TRANS_INSP)
    Dim lobjclsTransCobInsp As clsA46V731A    'Objeto da classe de estrutura stcA46V731A (TAB_TRANS_INSP)
    Dim lRegAlt As Long

    On Local Error GoTo ControleErros

    sSql = mfFormatarExclusaoSQL(pobjstc)

    If M46V999.gfAbrBasDados("RAMOSDIVERSOS", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, lConexaoSQL, pMensagem) = False Then
        Exit Function
    End If

    If M46V999.gfExeSQL(lConexaoSQL, sSql, UltimoErro, lRegAlt) <> 0 Then GoTo ControleErros

'    If lRegAlt > 0 Then
        Set lobjstcTransCobInsp = New stcA46V731B
        Set lobjclsTransCobInsp = New clsA46V731A
        Set lobjclsTransCobInsp.M46V999 = M46V999

        lobjstcTransCobInsp.nosso_numero = pobjstc.nosso_numero
        lobjstcTransCobInsp.Num_Item = pobjstc.Num_Item
        If Not lobjclsTransCobInsp.gfExcluir(lobjstcTransCobInsp) Then
            UltimoErro = lobjclsTransCobInsp.UltimoErro
            Exit Function
        End If
'    End If

    sSql = mfFormatarInclusaoSQL(pobjstc)

    If M46V999.gfExeSQL(lConexaoSQL, sSql, UltimoErro) <> 0 Then GoTo ControleErros

    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID

    Set lobjstcTransCobInsp = Nothing
    Set lobjclsTransCobInsp = Nothing
    gfIncluir = True

    Exit Function

ControleErros:
    Set lobjstcTransCobInsp = Nothing
    Set lobjclsTransCobInsp = Nothing
    gfIncluir = False
    If UCase(UltimoErro) Like "*" & UCase("duplicate key") & "*" Then
        UltimoErro = "DUPLICADO"
        gfIncluir = True
    Else
        If Not M46V999.gfPreenchido(UltimoErro) Then
            UltimoErro = Err.Description
        End If
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
