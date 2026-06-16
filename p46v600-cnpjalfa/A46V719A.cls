VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V719A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: clsA46V719A --> Tabela Tab_Ped_Auto
' Data/Hora: 08/08/2006 - 15:29:14
' Sistema: Multi
' Simony Sales
' ======================================================================================================================================================


' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V719B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V719A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V719B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String

    Dim lInsert As String                'Query.

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Auto ( "
    lInsert = lInsert & " Nosso_Numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " Num_Ped,"
    lInsert = lInsert & " Cod_Org,"
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " Num_Apol,"
    lInsert = lInsert & " Num_Endo,"
    lInsert = lInsert & " Val_is_auto, "
    lInsert = lInsert & " Val_is_rc, "
    lInsert = lInsert & " Val_is_app, "
    lInsert = lInsert & " Val_pr_auto, "
    lInsert = lInsert & " Val_pr_rc, "
    lInsert = lInsert & " Val_pr_app, "
    lInsert = lInsert & " Num_PI"

    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Is_Auto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Is_Rc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Is_App) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Auto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Rc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_App) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ")"

    mfFormatarInclusaoSQL = lInsert
End Function


Private Function mfFormatarAlteracaoSQL(ByRef pobjstc As stcA46V719B, _
                                        ByVal pPropTabela As String, _
                                        ByVal pNomeTabImport As String) As String

    Dim lUpdate As String                'Query.

    lUpdate = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ped_Auto Set "
    lUpdate = lUpdate & " Num_Ped = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lUpdate = lUpdate & " Cod_Org = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lUpdate = lUpdate & " Cod_Ramo = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lUpdate = lUpdate & " Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lUpdate = lUpdate & " Num_Endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lUpdate = lUpdate & " Val_is_auto = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Is_Auto) & ","
    lUpdate = lUpdate & " Val_is_rc = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Is_Rc) & ","
    lUpdate = lUpdate & " Val_is_app = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Is_App) & ","
    lUpdate = lUpdate & " Val_pr_auto = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Auto) & ","
    lUpdate = lUpdate & " Val_pr_rc = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Rc) & ","
    lUpdate = lUpdate & " Val_pr_app  = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_App) & ","
    lUpdate = lUpdate & " Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)
    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " NOsso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarAlteracaoSQL = lUpdate
End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V719B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String

    Dim lDelete As String                'Query.

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Auto WHERE "
    lDelete = lDelete & " Nosso_NUmero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarExclusaoSQL = lDelete
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V719B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String                'Query.

    lSelect = "SELECT * FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Auto  with(nolock)"
    lSelect = lSelect & " Where Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarConsultaSQL = lSelect
End Function
Private Sub mpLimparColecao()

    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V719B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)

    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)

    pobjstc.Val_Is_Auto = M46V999.gfTratarNulo(prst("Val_Is_Auto"), TpoValor)
    pobjstc.Val_Is_Rc = M46V999.gfTratarNulo(prst("Val_Is_Rc"), TpoValor)
    pobjstc.Val_Is_App = M46V999.gfTratarNulo(prst("Val_Is_App"), TpoValor)
    pobjstc.Val_Pr_Auto = M46V999.gfTratarNulo(prst("Val_Pr_Auto"), TpoValor)
    pobjstc.Val_Pr_Rc = M46V999.gfTratarNulo(prst("Val_Pr_Rc"), TpoValor)
    pobjstc.Val_Pr_App = M46V999.gfTratarNulo(prst("Val_Pr_App"), TpoValor)

    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)

End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V719B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean

    Dim lsSQl As String           'Query.

    On Local Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfIncluir = True
    Exit Function
ControleErro:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V719A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V719A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function
Public Function mfAlterar(ByRef pobjstc As stcA46V719B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean

    Dim lsSQl As String           'Query.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarAlteracaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
    If Colecao.Count > 0 Then
        Colecao.Remove (pobjstc.Class_ID)
    End If
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfAlterar = True
    Exit Function
ControleErro:
    mfAlterar = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
    End If
End Function
Public Function mfExcluir(ByRef pobjstc As stcA46V719B, _
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
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function
Public Function mfCarregarColecao(ByRef pobjstc As stcA46V719B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V719B          'Objeto da classe de estrutura stcA46V719B.
    Dim lrst As clsYasRecordSet   'Registro.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    If M46V999.gfObtRegistro(ConexaoBD, mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport), _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V719B
        Call mpCarregarAtributos(lobjstc, lrst)
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
