VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V733A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: clsA46V733A --> Tabela Tab_Canal_Emiss
' Sistema: Multi
' Simony Sales

' ======================================================================================================================================================


' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V733B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V733A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V733B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String

    Dim lInsert As String                'Query.

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Canal_Emiss ( "
    lInsert = lInsert & " Nosso_Numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " Num_Seq,"
    lInsert = lInsert & " Num_PI,"
    lInsert = lInsert & " Num_Apol,"
    lInsert = lInsert & " Num_Compl,"
    lInsert = lInsert & " Tip_Canal,"
    lInsert = lInsert & " Cod_Canal,"
    lInsert = lInsert & " Cod_Corretor,"
    lInsert = lInsert & " Perc_Canal,"
    lInsert = lInsert & " Qtd_Parc_Pag,"
    lInsert = lInsert & " Qtd_Parc_Apol,"
    lInsert = lInsert & " Cod_Form_Pag_Ag"


    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If

    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_SEQ) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Compl) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Canal) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Canal) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corretor) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Perc_Canal) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Parc_Pag) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_Parc_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Form_Pag_Ag) & ")"

    mfFormatarInclusaoSQL = lInsert
End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V733B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String

    Dim lDelete As String                'Query.

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Canal_Emiss WHERE "
    lDelete = lDelete & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Cod_Corretor <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Corretor= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corretor)
    End If
    If pobjstc.Cod_Canal <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Canal= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Canal)
    End If
    If pobjstc.Tip_Canal <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Tip_Canal= " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Canal)
    End If

    mfFormatarExclusaoSQL = lDelete
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V733B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String                'Query.

    lSelect = "SELECT C.* "
    lSelect = lSelect & " from " & pPropTabela & pNomeTabImport & "Tab_Canal_Emiss as C with(nolock) "
    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lSelect = lSelect & " Where Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
        If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
            lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
        End If
    End If

    If pobjstc.Cod_Corretor <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Corretor= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corretor)
    End If

    If pobjstc.Cod_Canal <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Canal= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Canal)
    End If

    If pobjstc.Tip_Canal <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Tip_Canal= " & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Canal)
    End If

    lSelect = lSelect & " Order By Num_seq"

    mfFormatarConsultaSQL = lSelect
End Function
Private Sub mpLimparColecao()

    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V733B, prst As clsYasRecordSet)


    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.NUM_SEQ = M46V999.gfTratarNulo(prst("Num_Seq"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Compl = M46V999.gfTratarNulo(prst("Num_Compl"), tpoNumerico)
    pobjstc.Tip_Canal = M46V999.gfTratarNulo(prst("Tip_Canal"), tpoNumerico)
    pobjstc.Cod_Canal = M46V999.gfTratarNulo(prst("Cod_Canal"), tpoNumerico)
    pobjstc.Cod_Corretor = M46V999.gfTratarNulo(prst("Cod_Corretor"), tpoNumerico)
    pobjstc.Perc_Canal = M46V999.gfTratarNulo(prst("Perc_Canal"), tpoNumerico)
    pobjstc.Qtd_Parc_Pag = M46V999.gfTratarNulo(prst("Qtd_Parc_Pag"), tpoNumerico)
    pobjstc.Qtd_Parc_Apol = M46V999.gfTratarNulo(prst("Qtd_Parc_Apol"), tpoNumerico)
    pobjstc.Cod_Form_Pag_Ag = M46V999.gfTratarNulo(prst("Cod_Form_Pag_Ag"), tpoNumerico)

End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V733B, _
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
        Call M46V999.gpGraLog(0, "A46V733A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V733A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function
Public Function mfExcluir(ByRef pobjstc As stcA46V733B, _
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
        Call M46V999.gpGraLog(0, "A46V733A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V733A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function
Public Function mfCarregarColecao(ByRef pobjstc As stcA46V733B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String, _
                                  Optional ByRef pListaNosso_Numero As String) As Boolean
    Dim lobjstc As stcA46V733B      'Objeto da classe de estrutura stcA46V733B.
    Dim lrst As clsYasRecordSet    'Registro.
    Dim lsSQl As String         'SQL.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V733B
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
    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V733A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V733A.mfCarregarColecao ERRO - " & UltimoErro & _
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

    lsSQl = " INSERT INTO " & pPropTabela & "Tab_Canal_Emiss ( "
    lsSQl = lsSQl & " Nosso_numero, "
    lsSQl = lsSQl & " Num_Seq,"
    lsSQl = lsSQl & " Num_PI,"
    lsSQl = lsSQl & " Num_Apol,"
    lsSQl = lsSQl & " Num_Compl,"
    lsSQl = lsSQl & " Tip_Canal,"
    lsSQl = lsSQl & " Cod_Canal,"
    lsSQl = lsSQl & " Cod_Corretor,"
    lsSQl = lsSQl & " Perc_Canal,"
    lsSQl = lsSQl & " Qtd_Parc_Pag,"
    lsSQl = lsSQl & " Qtd_Parc_Apol,"
    lsSQl = lsSQl & " Cod_Form_Pag_Ag"
    lsSQl = lsSQl & " ) SELECT "
    lsSQl = lsSQl & M46V999.gfFormatarCamposSQL(pNossoNumero) & ","
    lsSQl = lsSQl & " Num_Seq,"
    lsSQl = lsSQl & " Num_PI,"
    lsSQl = lsSQl & " 0, "
    lsSQl = lsSQl & " 0,"
    lsSQl = lsSQl & " Tip_Canal,"
    lsSQl = lsSQl & " Cod_Canal,"
    lsSQl = lsSQl & " Cod_Corretor,"
    lsSQl = lsSQl & " Perc_Canal,"
    lsSQl = lsSQl & " Qtd_Parc_Pag,"
    lsSQl = lsSQl & " Qtd_Parc_Apol,"
    lsSQl = lsSQl & " Cod_Form_Pag_Ag"
    lsSQl = lsSQl & " FROM " & pPropTabela & "Tab_Canal_Emiss WITH (NOLOCK)"
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
