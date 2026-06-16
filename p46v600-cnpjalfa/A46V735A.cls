VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V735A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V735A -
' Data/Hora: 08/11/2013
' Sistema: Multi
' Banco: RamosDiversos
' Tabela: TAB_PED_LOC_PESSOAS
' Simony Sales
' ======================================================================================================================================================

Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V735B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V735A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V735B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lInsert As String                'Query.

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc_Pessoas ( "
    lInsert = lInsert & " Num_Item,"
    lInsert = lInsert & " Nosso_Numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If

    lInsert = lInsert & " Num_Apol,"
    lInsert = lInsert & " Num_Endo,"
    lInsert = lInsert & " Num_Pi,"

    '  lInsert = lInsert & " Ped_Loc_Pessoa_ID, "

    lInsert = lInsert & " Tip_Loc_Pessoa, "
    lInsert = lInsert & " Nom_Pessoa, "
    lInsert = lInsert & " Tip_Pessoa, "
    lInsert = lInsert & " Num_CPF_CNPJ, "
    lInsert = lInsert & " Tip_Cobert, "
    lInsert = lInsert & "  Per_Participacao, "
    lInsert = lInsert & "  Tip_Doc, "
    lInsert = lInsert & "  Num_Doc "

    lInsert = lInsert & ") VALUES ("

    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If

    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","

    '  lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ped_Loc_Pessoa_ID) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Loc_Pessoa) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Pessoa) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Pessoa) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_CPF_CNPJ) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Cobert) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Participacao) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Doc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Doc) & ")"

    mfFormatarInclusaoSQL = lInsert
End Function


Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V735B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lDelete As String                'Query.

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc_Pessoas WHERE "
    lDelete = lDelete & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Item > 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If
    If pobjstc.Tip_Loc_Pessoa > 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Tip_Loc_Pessoa = " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Loc_Pessoa)
    End If

    mfFormatarExclusaoSQL = lDelete
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V735B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String                'Query.
    lSelect = "SELECT * FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc_Pessoas with(nolock) "
    'lSelect = lSelect & " Where nosso_numero = " & M46V999.gfFormatarCamposSQL(pobjstc.NOSSO_NUMERO)
    If UCase(pPropTabela) <> "GED.DBO." Then
        lSelect = lSelect & " Where nosso_numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Apol <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol)
        lSelect = M46V999.gfWhereAnd(lSelect) & " num_endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo)
    End If
    If pobjstc.NUM_PI <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)
    End If
    lSelect = lSelect & " and Tip_Loc_Pessoa = " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Loc_Pessoa)
    lSelect = lSelect & " and Num_item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)


    mfFormatarConsultaSQL = lSelect

End Function
Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V735B, prst As clsYasRecordSet)

    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_Item"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)

    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)


    pobjstc.Ped_Loc_Pessoa_ID = M46V999.gfTratarNulo(prst("Ped_Loc_Pessoa_ID"), tpoNumerico)
    pobjstc.Tip_Loc_Pessoa = M46V999.gfTratarNulo(prst("Tip_Loc_Pessoa"), tpoNumerico)
    pobjstc.Nom_Pessoa = M46V999.gfTratarNulo(prst("Nom_Pessoa"), tpoString)
    pobjstc.Tip_Pessoa = M46V999.gfTratarNulo(prst("Tip_Pessoa"), tpoNumerico)
    pobjstc.Num_CPF_CNPJ = M46V999.gfTratarNulo(prst("Num_CPF_CNPJ"), tpoNumerico)

    pobjstc.Tip_Cobert = M46V999.gfTratarNulo(prst("Tip_Cobert"), tpoNumerico)
    pobjstc.Per_Participacao = M46V999.gfTratarNulo(prst("Per_Participacao"), tpoNumerico)
    pobjstc.Tip_Doc = M46V999.gfTratarNulo(prst("Tip_Doc"), tpoNumerico)
    pobjstc.Num_Doc = M46V999.gfTratarNulo(prst("Num_Doc"), tpoString)

End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V735B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfIncluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lsSQl As String           'Query.

    On Local Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
Retorno:
    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
    Colecao.Add Item:=pobjstc
    mfIncluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErros:
    mfIncluir = False
    If Not M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = UltimoErro & Err & " " & Err.Description
    End If

    Call M46V999.gpGraLog(0, "A46V735A.mfIncluir ERRO - " & UltimoErro & _
                             vbCrLf & " SQL: " & lsSQl)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)

End Function

Public Function mfExcluir(ByRef pobjstc As stcA46V735B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean
    Dim lsSQl As String           'Query.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
    Set pobjstc = Nothing
    mfExcluir = True
    Exit Function
ControleErros:
    mfExcluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V735A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V735A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V735B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V735B          'Objeto da classe de estrutura stcA46V735B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErros
    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V735B
        Set lobjstc.M46V999 = M46V999

        Call mpCarregarAtributos(lobjstc, lrst)
        Colecao.Add Item:=lobjstc
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
        Call M46V999.gpGraLog(0, "A46V735A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V735A.mfCarregarColecao ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

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
Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
End Property

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
