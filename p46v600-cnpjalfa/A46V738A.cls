VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V738A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V738A -
' Data/Hora: 22/07/2014
' Sistema: Multi
' Banco: RamosDiversos
' Tabela: TAB_PED_COBERT_FRANQUIA
' Rodrigo Berçot
' ======================================================================================================================================================

Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V738B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V738A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V738B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lInsert As String                'Query.

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Franquia ( "
    lInsert = lInsert & " Nosso_Numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If

    lInsert = lInsert & " Cod_Ramo, "
    lInsert = lInsert & " Cod_Modalidade, "
    lInsert = lInsert & " Cod_Tarif, "
    lInsert = lInsert & " Cod_Cobert, "
    lInsert = lInsert & " Cod_Evento, "
    lInsert = lInsert & " Num_Ped, "
    lInsert = lInsert & " Num_Apol, "
    lInsert = lInsert & " Num_Endosso, "
    lInsert = lInsert & " Num_Item, "
    lInsert = lInsert & " Num_PI, "
    lInsert = lInsert & " Val_Min_Franq, "
    lInsert = lInsert & " Per_Franq, "
    lInsert = lInsert & " Texto_Franquia, "
    lInsert = lInsert & " Texto_Observacoes, "
    lInsert = lInsert & " Multiplo_Franquia "

    lInsert = lInsert & ") VALUES ("

    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If

    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Modalidade) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tarif) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Evento) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endosso) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Min_Franq) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Franq) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Texto_Franquia, True) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Texto_Observacoes, True) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Multiplo_Franquia) & ")"

    mfFormatarInclusaoSQL = lInsert
End Function


Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V738B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lDelete As String                'Query.

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Franquia WHERE "
    
    If pPropTabela <> "GED.DBO." Then
        lDelete = lDelete & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.NUM_PI > 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)
    End If
    
    If pobjstc.Num_Item > 0 Then
        lDelete = lDelete & " AND Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If

    mfFormatarExclusaoSQL = lDelete
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V738B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String                'Query.
    lSelect = "SELECT * FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Franquia with(nolock) "

    If UCase(pPropTabela) <> "GED.DBO." Then
        lSelect = lSelect & " Where nosso_numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If
    
    If M46V999.gfPreenchido(pobjstc.NUM_PI) Then
        lSelect = M46V999.gfWhereAnd(lSelect) & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Franquia.Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)
    End If

    If pobjstc.Cod_Cobert > 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Franquia.Cod_Cobert = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert)
    End If

    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Franquia.Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Item > 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Franquia.Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If
    If pobjstc.Num_Apol <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Franquia.Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol)
        lSelect = M46V999.gfWhereAnd(lSelect) & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Franquia.num_endosso = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endosso)
    End If
    lSelect = lSelect & " order by " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Franquia.num_item, " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Franquia.cod_cobert"
    mfFormatarConsultaSQL = lSelect

End Function
Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V738B, prst As clsYasRecordSet)

    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Cod_Modalidade = M46V999.gfTratarNulo(prst("Cod_Modalidade"), tpoNumerico)
    pobjstc.Cod_Tarif = M46V999.gfTratarNulo(prst("Cod_Tarif"), tpoNumerico)
    pobjstc.Cod_Cobert = M46V999.gfTratarNulo(prst("Cod_Cobert"), tpoNumerico)
    pobjstc.Cod_Evento = M46V999.gfTratarNulo(prst("Cod_Evento"), tpoNumerico)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_Item"), tpoNumerico)
    pobjstc.Num_Endosso = M46V999.gfTratarNulo(prst("Num_Endosso"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Val_Min_Franq = M46V999.gfTratarNulo(prst("Val_Min_Franq"), tpoNumerico)
    pobjstc.Per_Franq = M46V999.gfTratarNulo(prst("Per_Franq"), tpoNumerico)
    pobjstc.Texto_Franquia = M46V999.gfTratarNulo(prst("Texto_Franquia"), tpoString)
    pobjstc.Texto_Observacoes = M46V999.gfTratarNulo(prst("Texto_Observacoes"), tpoString)
    pobjstc.Multiplo_Franquia = M46V999.gfTratarNulo(prst("Multiplo_Franquia"), tpoNumerico)

End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V738B, _
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

    Call M46V999.gpGraLog(0, "A46V738A.mfIncluir ERRO - " & UltimoErro & _
                             vbCrLf & " SQL: " & lsSQl)

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)

End Function

Public Function mfExcluir(ByRef pobjstc As stcA46V738B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                          Optional pCod_Cobert As Integer = 0) As Boolean
    Dim lsSQl As String           'Query.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    
    If pCod_Cobert <> 0 Then
        lsSQl = lsSQl & " AND Cod_Cobert = " & M46V999.gfFormatarCamposSQL(pCod_Cobert)
    End If
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
    
    If pCod_Cobert = 0 Then
        Set pobjstc = Nothing
    End If
    mfExcluir = True
    Exit Function
ControleErros:
    mfExcluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V738A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V738A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V738B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V738B          'Objeto da classe de estrutura stcA46V738B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErros
    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V738B
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
        Call M46V999.gpGraLog(0, "A46V738A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V738A.mfCarregarColecao ERRO - " & UltimoErro & _
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




