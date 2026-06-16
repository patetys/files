VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46v741A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Inspeção – Equipamentos
' ======================================================================================================================================================
' Class Name: clsA46v741A
' Data/Hora: 26/05/2015 - 11:00:00
' Sistema: P46V600
' Analista: Carlos Araujo
' ======================================================================================================================================================

Public ConexaoBD                As ADODB.Connection  ' Conexão com o Banco de Dados
Public Colecao                  As Collection        ' Coleção de objetos do tipo stcA46v730B
Public UltimoErro               As String            ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46v741A"
End Property

Private Function mfFormatarConsultaSQL(pobjstc As stcA46v741B) As String

Dim strTMP As String
    
    strTMP = "Select * from Tab_Ped_Agregados where NUM_PI = " & pobjstc.NUM_PI
    strTMP = strTMP & " and TIP_AGREGADO = " & pobjstc.Tip_Agragado
    
    mfFormatarConsultaSQL = strTMP
    
End Function

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46v741B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       ByVal TabAtual As Integer) As String
    Dim lInsert As String                'Query.

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Agregados ( "
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " NOSSO_NUMERO,"
    lInsert = lInsert & " NUM_PI,"
    lInsert = lInsert & " COD_ORG,"
    lInsert = lInsert & " COD_RAMO,"
    lInsert = lInsert & " NUM_PED,"
    lInsert = lInsert & " NUM_APOL,"
    lInsert = lInsert & " NUM_ENDO,"
    lInsert = lInsert & " TIP_AGREGADO,"
    lInsert = lInsert & " NOM_SEGURADO,"
    lInsert = lInsert & " NUM_CGC_CPF,"
    lInsert = lInsert & " COD_TIP_LOGR,"
    lInsert = lInsert & " NOM_LOGR,"
    lInsert = lInsert & " DSC_COMPL_LOGR,"
    lInsert = lInsert & " NUM_LOGR,"
    lInsert = lInsert & " NOM_BAIRRO,"
    lInsert = lInsert & " NOM_CID,"
    lInsert = lInsert & " SIG_UF,"
    lInsert = lInsert & " NUM_CEP,"
    lInsert = lInsert & " Tipo_fone,"
    lInsert = lInsert & " Num_ddd_fone,"
    lInsert = lInsert & " Num_fone,"
    lInsert = lInsert & " Num_ramal,"
    lInsert = lInsert & " TIP_PESSOA,"
    lInsert = lInsert & " COD_RAMO_ATIV,"
    lInsert = lInsert & " COD_DOC,"
    lInsert = lInsert & " NUM_DOC,"
    lInsert = lInsert & " ORG_EXP,"
    lInsert = lInsert & " DAT_EMIS_DOC,"
    lInsert = lInsert & " DAT_NASC,"
    lInsert = lInsert & " NOM_EMPRESA,"
    lInsert = lInsert & " SEXO,"
    lInsert = lInsert & " Renda_Fatura"
    lInsert = lInsert & ") VALUES ("
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Agragado) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Segurado) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cgc_Cpf) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Logr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Logr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Logr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Logr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Bairro) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Cid) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Sig_UF) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_CEP) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tipo_Fone) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_DDD_Fone) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Fone) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ramal) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Pessoa) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo_Ativ) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Doc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Doc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Org_Exp) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Emis_Doc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.DAT_Nasc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Empresa) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Sexo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Renda_Fatura) & ")"

    mfFormatarInclusaoSQL = lInsert

End Function

Private Function mfFormatarAlteracaoSQL(ByRef pobjstc As stcA46v741B, _
                                        ByVal pPropTabela As String, _
                                        ByVal pNomeTabImport As String, _
                                        ByVal TabAtual As Integer) As String

    Dim lUpdate As String                'Query.

    lUpdate = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ped_Agregados Set "
    lUpdate = lUpdate & " NUM_PI                          = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lUpdate = lUpdate & " COD_ORG                         = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lUpdate = lUpdate & " COD_RAMO                        = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lUpdate = lUpdate & " NUM_PED                         = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lUpdate = lUpdate & " NUM_APOL                        = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lUpdate = lUpdate & " NUM_ENDO                        = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lUpdate = lUpdate & " TIP_AGREGADO                    = " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Agragado) & ","
    lUpdate = lUpdate & " NOM_SEGURADO                    = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Segurado) & ","
    lUpdate = lUpdate & " NUM_CGC_CPF                     = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cgc_Cpf) & ","
    lUpdate = lUpdate & " COD_TIP_LOGR                    = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Logr) & ","
    lUpdate = lUpdate & " NOM_LOGR                        = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Logr) & ","
    lUpdate = lUpdate & " DSC_COMPL_LOGR                  = " & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Logr) & ","
    lUpdate = lUpdate & " NUM_LOGR                        = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Logr) & ","
    lUpdate = lUpdate & " NOM_BAIRRO                      = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Bairro) & ","
    lUpdate = lUpdate & " NOM_CID                         = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Cid) & ","
    lUpdate = lUpdate & " SIG_UF                          = " & M46V999.gfFormatarCamposSQL(pobjstc.Sig_UF) & ","
    lUpdate = lUpdate & " NUM_CEP                         = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_CEP) & ","
    lUpdate = lUpdate & " Tipo_fone                       = " & M46V999.gfFormatarCamposSQL(pobjstc.Tipo_Fone) & ","
    lUpdate = lUpdate & " Num_ddd_fone                    = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_DDD_Fone) & ","
    lUpdate = lUpdate & " Num_fone                        = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Fone) & ","
    lUpdate = lUpdate & " Num_ramal                       = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ramal) & ","
    lUpdate = lUpdate & " TIP_PESSOA                      = " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Pessoa) & ","
    lUpdate = lUpdate & " COD_RAMO_ATIV                   = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo_Ativ) & ","
    lUpdate = lUpdate & " COD_DOC                         = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Doc) & ","
    lUpdate = lUpdate & " NUM_DOC                         = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Doc) & ","
    lUpdate = lUpdate & " ORG_EXP                         = " & M46V999.gfFormatarCamposSQL(pobjstc.Org_Exp) & ","
    lUpdate = lUpdate & " DAT_EMIS_DOC                    = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Emis_Doc) & ","
    lUpdate = lUpdate & " DAT_NASC                        = " & M46V999.gfFormatarCamposSQL(pobjstc.DAT_Nasc) & ","
    lUpdate = lUpdate & " NOM_EMPRESA                     = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Empresa) & ","
    lUpdate = lUpdate & " SEXO                            = " & M46V999.gfFormatarCamposSQL(pobjstc.Sexo) & ","
    lUpdate = lUpdate & " Renda_Fatura                    = " & M46V999.gfFormatarCamposSQL(pobjstc.Renda_Fatura) & ","
    
    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " NUM_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)
    lUpdate = lUpdate & " TIP_AGREGADO = " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Agragado)
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarAlteracaoSQL = lUpdate

End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46v741B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       ByVal TabAtual As Integer) As String
    Dim lDelete As String                'Query.

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Agregados WHERE "
    lDelete = lDelete & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarExclusaoSQL = lDelete

End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46v741B, _
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
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID

    mfIncluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErros:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V741A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V741A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function
Public Function mfAlterar(ByRef pobjstc As stcA46v741B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                          Optional pDataUltAlt As Double, _
                          Optional pUltUsuAlt As String) As Boolean
    Dim lsSQl As String           'Query.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarAlteracaoSQL(pobjstc, pPropTabela, pNomeTabImport, 1)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
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
        Call M46V999.gpGraLog(0, "A46V741A.mfAlterar ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V741A.mfAlterar ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function
Public Function mfExcluir(ByRef pobjstc As stcA46v741B, _
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
        Call M46V999.gpGraLog(0, "A46V741A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V741A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
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


