VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V756A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V756A -
' Data/Hora: 07/02/2018
' Tabela: TAB_DADOS_CARTAO
' ======================================================================================================================================================

Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection            ' Coleção de objetos do tipo stcA46V756B
Public UltimoErro As String             ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V756A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V756B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lInsert As String 'Query.

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Dados_Cartao ( "
    lInsert = lInsert & " NUM_PI_EMIT,"
    lInsert = lInsert & " COD_RAMO_EMIT,"
    lInsert = lInsert & " TRANSACTIONID,"
    lInsert = lInsert & " AUTHCODE,"
    lInsert = lInsert & " DTMOV,"
    lInsert = lInsert & " TRANSACTIONDATE,"
    lInsert = lInsert & " TRANSACTIONSTATUS,"
    lInsert = lInsert & " TRANSACTIONTYPE,"
    lInsert = lInsert & " ORDERNUMBER,"
    lInsert = lInsert & " ORDERTOTAL,"
    lInsert = lInsert & " ERROR,"
    lInsert = lInsert & " XML_SEND,"
    lInsert = lInsert & " XML_RECEIVE,"
    lInsert = lInsert & " NUM_TEMPO_TOTAL,"
    lInsert = lInsert & " TIMESATAMPBW "

    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI_EMIT) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.COD_RAMO_EMIT) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.TransactionId) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.AUTHCODE) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.DTMOV) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.TRANSACTIONDATE) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.TRANSACTIONSTATUS) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.TRANSACTIONTYPE) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.ORDERNUMBER) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.ORDERTOTAL) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(Trim(pobjstc.ERROR)) & ","
    lInsert = lInsert & "CONVERT(XML, '" & pobjstc.XML_SEND & "'),"
    lInsert = lInsert & "CONVERT(XML, '" & pobjstc.XML_RECEIVE & "'),"
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_TEMPO_TOTAL) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.TIMESATAMPBW) & ")"

    mfFormatarInclusaoSQL = lInsert
    
End Function


Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V756B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lDelete As String 'Query.

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Dados_Cartao WHERE "
    lDelete = lDelete & " NUM_PI_EMIT = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI_EMIT)

    mfFormatarExclusaoSQL = lDelete
    
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V756B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String 'Query.
    lSelect = "SELECT * FROM " & pPropTabela & pNomeTabImport & "Tab_Dados_Cartao with(nolock) "
    lSelect = lSelect & " Where Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI_EMIT)
    
    mfFormatarConsultaSQL = lSelect

End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V756B, prst As clsYasRecordSet)

    pobjstc.NUM_PI_EMIT = M46V999.gfTratarNulo(prst("NUM_PI_EMIT"), TpoNumDecimal)
    pobjstc.COD_RAMO_EMIT = M46V999.gfTratarNulo(prst("COD_RAMO_EMIT"), TpoNumDecimal)
    pobjstc.TransactionId = M46V999.gfTratarNulo(prst("TRANSACTIONID"), tpoString)
    pobjstc.AUTHCODE = M46V999.gfTratarNulo(prst("AUTHCODE"), tpoString)
    pobjstc.DTMOV = M46V999.gfTratarNulo(prst("DTMOV"), TpoData)
    pobjstc.TRANSACTIONDATE = M46V999.gfTratarNulo(prst("TRANSACTIONDATE"), TpoData)
    pobjstc.TRANSACTIONSTATUS = M46V999.gfTratarNulo(prst("TRANSACTIONSTATUS"), TpoNumInteiro)
    pobjstc.TRANSACTIONTYPE = M46V999.gfTratarNulo(prst("TRANSACTIONTYPE"), TpoNumInteiro)
    pobjstc.ORDERNUMBER = M46V999.gfTratarNulo(prst("ORDERNUMBER"), tpoNumerico)
    pobjstc.ORDERTOTAL = M46V999.gfTratarNulo(prst("ORDERTOTAL"), TpoNumDecimal)
    pobjstc.ERROR = M46V999.gfTratarNulo(prst("ERROR"), tpoString)
    pobjstc.XML_SEND = M46V999.gfTratarNulo(prst("XML_SEND"), tpoString)
    pobjstc.XML_RECEIVE = M46V999.gfTratarNulo(prst("XML_RECEIVE"), tpoString)
    pobjstc.NUM_TEMPO_TOTAL = M46V999.gfTratarNulo(prst("NUM_TEMPO_TOTAL"), tpoNumerico)
    pobjstc.TIMESATAMPBW = M46V999.gfTratarNulo(prst("TIMESATAMPBW"), TpoData)

End Sub

Public Function mfIncluir(ByRef pobjstc As stcA46V756B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                          tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    Dim methodName As String
    Dim lsSQl As String 'Query
    
    M46V999.tagDynaTrace = tagDynaTrace
    methodName = "mfIncluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

    On Local Error GoTo ControleErros
    
    If Not mfConectar Then GoTo ControleErros

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

    Call M46V999.gpGraLog(0, "A46V756A.mfIncluir ERRO - " & UltimoErro & vbCrLf & " SQL: " & lsSQl)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)

End Function

Public Function mfExcluir(ByRef pobjstc As stcA46V756B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean
    Dim lsSQl As String 'Query.

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
        Call M46V999.gpGraLog(0, "A46V756A.mfExcluir ERRO - " & Err & " " & Err.Description & vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V756A.mfExcluir ERRO - " & UltimoErro & vbCrLf & " SQL: " & lsSQl)
    End If
    
End Function

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V756B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
                                  
    Dim lobjstc As stcA46V756B          'Objeto da classe de estrutura stcA46V735B.
    Dim lrst As clsYasRecordSet         'Registro.
    Dim lsSQl As String                 'SQL.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    
    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrst, UltimoErro) = False Then GoTo ControleErros
    
    Call mpLimparColecao
    
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V756B
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
        Call M46V999.gpGraLog(0, "A46V756A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V756A.mfCarregarColecao ERRO - " & UltimoErro & vbCrLf & " SQL: " & lsSQl)
    End If
    
End Function

Public Property Get M46V999() As clsM46V999
    
    If mClsM46V999 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V999 não atribuído para a classe (" & Me.Class_Name & ")"
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




