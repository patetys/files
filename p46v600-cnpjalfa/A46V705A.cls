VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V705A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: clsA46V705A --> Tabela Tab_Ped_Corr
' Data/Hora: 31/08/2005 - 11:53:12
' Sistema: Multi
' Thais Juren Gomes

' ======================================================================================================================================================


' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V705B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V705A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V705B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String

    Dim lInsert As String                'Query.

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Corr ( "
    lInsert = lInsert & " Nosso_Numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " Cod_Corr,"
    lInsert = lInsert & " Num_Ped,"
    lInsert = lInsert & " Cod_Org,"
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " Num_Apol,"
    lInsert = lInsert & " Num_Endo,"
    lInsert = lInsert & " Per_Corr,"
    lInsert = lInsert & " Num_Ordem_Corr,"
    lInsert = lInsert & " Num_PI"

    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Corr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ordem_Corr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ")"

    mfFormatarInclusaoSQL = lInsert
End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V705B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String

    Dim lDelete As String                'Query.

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Corr WHERE "
    lDelete = lDelete & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Cod_Corr <> 0 Then
        lDelete = lDelete & " And Cod_Corr= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corr)
    End If
    If pobjstc.Cod_Ramo <> 0 Then
        lDelete = lDelete & " And Cod_Ramo = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo)
    End If

    mfFormatarExclusaoSQL = lDelete
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V705B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String                'Query.

    lSelect = "SELECT * FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Corr with(nolock) "
    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lSelect = lSelect & " Where Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
        If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
            lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
        End If
        If pobjstc.Cod_Corr <> 0 Then
            lSelect = lSelect & " And Cod_Corr= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corr)
        End If
        If pobjstc.Cod_Ramo <> 0 Then
            lSelect = lSelect & " And Cod_Ramo = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo)
        End If
    Else
        If pobjstc.Cod_Corr <> 0 Then
            lSelect = lSelect & " Where Cod_Corr= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corr)
        End If
    End If
    If pobjstc.Num_Apol <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol)
        lSelect = M46V999.gfWhereAnd(lSelect) & " num_endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo)
    End If
    lSelect = lSelect & " Order By NUM_ORDEM_CORR"

    mfFormatarConsultaSQL = lSelect
End Function
Private Sub mpLimparColecao()

    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V705B, prst As clsYasRecordSet)

    pobjstc.Cod_Corr = M46V999.gfTratarNulo(prst("Cod_Corr"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Per_Corr = M46V999.gfTratarNulo(prst("Per_Corr"), tpoNumerico)
    pobjstc.Num_Ordem_Corr = M46V999.gfTratarNulo(prst("Num_Ordem_Corr"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)



    If M46V999.gSerDIVERSOS Like "*MIX*" And pobjstc.nosso_numero Like "*PADRAO*" Then
        pobjstc.Cod_Corr = M46V999.gCod_Corretor
    End If

End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V705B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean

    Dim lsSQl As String           'Query.

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfIncluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    On Local Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfIncluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErro:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V705A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V705A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function
Public Function mfExcluir(ByRef pobjstc As stcA46V705B, _
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

    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
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
        Call M46V999.gpGraLog(0, "A46V705A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V705A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function
Public Function mfCarregarColecao(ByRef pobjstc As stcA46V705B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String, _
                                  Optional ByRef pListaNosso_Numero As String) As Boolean
    Dim lobjstc As stcA46V705B      'Objeto da classe de estrutura stcA46V705B.
    Dim lrst As clsYasRecordSet    'Registro.
    Dim lsSQl As String         'SQL.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V705B
        Set lobjstc.M46V999 = M46V999

        Call mpCarregarAtributos(lobjstc, lrst)
        If M46V999.gfPreenchido(pListaNosso_Numero) Then pListaNosso_Numero = pListaNosso_Numero & ","
        pListaNosso_Numero = pListaNosso_Numero & "'" & lobjstc.nosso_numero & "'"
        Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID
        lrst.MoveNext
    Loop

    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecao = True
    Exit Function

ControleErro:
    mfCarregarColecao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V705A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V705A.mfCarregarColecao ERRO - " & UltimoErro & _
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

Public Function mfCopiarPadraoParaNovo(ByVal pNossoNumero As String, _
                                       ByVal pPropTabela As String, _
                                       ByRef pRetorno As Long) As Boolean
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros

    lsSQl = " INSERT INTO " & pPropTabela & "Tab_Ped_Corr ( "
    lsSQl = lsSQl & " Nosso_numero, "
    lsSQl = lsSQl & " COD_ORG, "
    lsSQl = lsSQl & " COD_RAMO, "
    lsSQl = lsSQl & " NUM_PED, "
    lsSQl = lsSQl & " NUM_APOL, "
    lsSQl = lsSQl & " NUM_ENDO, "
    lsSQl = lsSQl & " COD_CORR, "
    lsSQl = lsSQl & " PER_CORR, "
    lsSQl = lsSQl & " NUM_ORDEM_CORR, "
    lsSQl = lsSQl & " NUM_PI )"
    lsSQl = lsSQl & " SELECT "
    lsSQl = lsSQl & M46V999.gfFormatarCamposSQL(pNossoNumero) & ","
    lsSQl = lsSQl & " COD_ORG, "
    lsSQl = lsSQl & " COD_RAMO, "
    lsSQl = lsSQl & " NUM_PED, "
    lsSQl = lsSQl & " NUM_APOL, "
    lsSQl = lsSQl & " NUM_ENDO, "
    lsSQl = lsSQl & " COD_CORR, "
    'lsSQl = lsSQl & " PER_CORR, " 'Retirado conf. 701722
    lsSQl = lsSQl & " 0, "
    lsSQl = lsSQl & " NUM_ORDEM_CORR, "
    lsSQl = lsSQl & " NUM_PI "
    lsSQl = lsSQl & " FROM " & pPropTabela & "Tab_Ped_Corr WITH (NOLOCK)"
    lsSQl = lsSQl & " Where nosso_numero = 'PADRAO" & M46V999.gNomUsuario6 & "'"

    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro, pRetorno) <> 0 Then GoTo ControleErros

    mfCopiarPadraoParaNovo = True

    Exit Function
ControleErros:
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
    End If
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
