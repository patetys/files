VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V726A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V726A - Texto de Especificação
' Data/Hora: 05/06/2007 - 10:28:32
' Sistema: Multi
' Simony Sales
' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V72B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V726A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V726B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       Optional blnFormataTexto As Boolean = False) As String
    Dim lInsert As String                'Query.

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Espec ( "
    lInsert = lInsert & " Nosso_Numero ,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " Cod_Espec ,"
    lInsert = lInsert & " Cod_Org ,"
    lInsert = lInsert & " Cod_Ramo ,"
    lInsert = lInsert & " Num_Ped ,"
    lInsert = lInsert & " Num_Apol ,"
    lInsert = lInsert & " Num_Endo ,"
    lInsert = lInsert & " Num_pag ,"
    lInsert = lInsert & " Qtd_vias ,"
    lInsert = lInsert & " Espec_lin1 ,"
    lInsert = lInsert & " Espec_lin2 ,"
    lInsert = lInsert & " Espec_lin3 ,"
    lInsert = lInsert & " Espec_lin4 ,"
    lInsert = lInsert & " Espec_lin5 ,"
    lInsert = lInsert & " Espec_lin6 ,"
    lInsert = lInsert & " Espec_lin7 ,"
    lInsert = lInsert & " Espec_lin8 ,"
    lInsert = lInsert & " Espec_lin9 ,"
    lInsert = lInsert & " Espec_lin10 ,"
    lInsert = lInsert & " Espec_lin11 ,"
    lInsert = lInsert & " Espec_lin12 ,"
    lInsert = lInsert & " Espec_lin13 ,"
    lInsert = lInsert & " Espec_lin14 ,"
    lInsert = lInsert & " Espec_lin15 ,"
    lInsert = lInsert & " Espec_lin16 ,"
    lInsert = lInsert & " Espec_lin17 ,"
    lInsert = lInsert & " Espec_lin18 ,"
    lInsert = lInsert & " Espec_lin19 ,"
    lInsert = lInsert & " Espec_lin20 ,"
    lInsert = lInsert & " Num_pi "
    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Espec) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_pag) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Qtd_vias) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin1, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin2, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin3, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin4, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin5, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin6, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin7, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin8, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin9, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin10, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin11, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin12, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin13, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin14, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin15, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin16, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin17, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin18, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin19, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Espec_lin20, blnFormataTexto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ")"
    mfFormatarInclusaoSQL = lInsert
End Function


Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V726B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lDelete As String                'Query.

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Espec WHERE "
    lDelete = lDelete & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    If M46V999.CDblx(pobjstc.Cod_Espec) <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " cod_espec = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Espec)
    End If

    mfFormatarExclusaoSQL = lDelete
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V726B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String                'Query.
    lSelect = "SELECT * FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Espec  with(nolock) "
    lSelect = lSelect & " where Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If M46V999.CDblx(pobjstc.Cod_Espec) <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " cod_espec = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Espec)
    End If

    lSelect = lSelect & " order by Num_Pag "

    mfFormatarConsultaSQL = lSelect
End Function
Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V726B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Cod_Espec = M46V999.gfTratarNulo(prst("Cod_Espec"), tpoNumerico)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Num_pag = M46V999.gfTratarNulo(prst("Num_pag"), tpoNumerico)
    pobjstc.Qtd_vias = M46V999.gfTratarNulo(prst("Qtd_vias"), tpoNumerico)
    pobjstc.Espec_lin1 = M46V999.gfTratarNulo(prst("Espec_lin1"), tpoString)
    pobjstc.Espec_lin2 = M46V999.gfTratarNulo(prst("Espec_lin2"), tpoString)
    pobjstc.Espec_lin3 = M46V999.gfTratarNulo(prst("Espec_lin3"), tpoString)
    pobjstc.Espec_lin4 = M46V999.gfTratarNulo(prst("Espec_lin4"), tpoString)
    pobjstc.Espec_lin5 = M46V999.gfTratarNulo(prst("Espec_lin5"), tpoString)
    pobjstc.Espec_lin6 = M46V999.gfTratarNulo(prst("Espec_lin6"), tpoString)
    pobjstc.Espec_lin7 = M46V999.gfTratarNulo(prst("Espec_lin7"), tpoString)
    pobjstc.Espec_lin8 = M46V999.gfTratarNulo(prst("Espec_lin8"), tpoString)
    pobjstc.Espec_lin9 = M46V999.gfTratarNulo(prst("Espec_lin9"), tpoString)
    pobjstc.Espec_lin10 = M46V999.gfTratarNulo(prst("Espec_lin10"), tpoString)
    pobjstc.Espec_lin11 = M46V999.gfTratarNulo(prst("Espec_lin11"), tpoString)
    pobjstc.Espec_lin12 = M46V999.gfTratarNulo(prst("Espec_lin12"), tpoString)
    pobjstc.Espec_lin13 = M46V999.gfTratarNulo(prst("Espec_lin13"), tpoString)
    pobjstc.Espec_lin14 = M46V999.gfTratarNulo(prst("Espec_lin14"), tpoString)
    pobjstc.Espec_lin15 = M46V999.gfTratarNulo(prst("Espec_lin15"), tpoString)
    pobjstc.Espec_lin16 = M46V999.gfTratarNulo(prst("Espec_lin16"), tpoString)
    pobjstc.Espec_lin17 = M46V999.gfTratarNulo(prst("Espec_lin17"), tpoString)
    pobjstc.Espec_lin18 = M46V999.gfTratarNulo(prst("Espec_lin18"), tpoString)
    pobjstc.Espec_lin19 = M46V999.gfTratarNulo(prst("Espec_lin19"), tpoString)
    pobjstc.Espec_lin20 = M46V999.gfTratarNulo(prst("Espec_lin20"), tpoString)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_pi"), tpoNumerico)
End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V726B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, tagDynaTrace As String, Optional blnFormataTexto As Boolean = False) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfIncluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lsSQl As String           'Query.

    On Local Error GoTo ControleErros

    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, blnFormataTexto)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfIncluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErros:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V726A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V726A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function
Public Function mfExcluir(ByRef pobjstc As stcA46V726B, _
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
    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
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
        Call M46V999.gpGraLog(0, "A46V726A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V726A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function
Public Function mfCarregarColecao(ByRef pobjstc As stcA46V726B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V726B          'Objeto da classe de estrutura stcA46V726B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'Sql.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErros
    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V726B
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
        Call M46V999.gpGraLog(0, "A46V726A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V726A.mfCarregarColecao ERRO - " & UltimoErro & _
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
