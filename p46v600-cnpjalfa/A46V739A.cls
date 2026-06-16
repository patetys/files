VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V739A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: clsA46V739A --> Tabela Tab_Ped_Textos
' Data/Hora: 23/09/2014 - 16:23:00
' Sistema: Multi
' Daniel Del Bianco Costa

' ======================================================================================================================================================


' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection            ' Coleção de objetos do tipo stcA46V739B
Public UltimoErro As String             ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V739A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V739B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       Optional ByVal CodTab As Integer) As String
    Dim lInsert As String 'Query.

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos ( "
    lInsert = lInsert & " Nosso_numero,"
    'If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
    '    lInsert = lInsert & " Cod_Unicidade,"
    'End If
    lInsert = lInsert & " COD_ORG,"
    lInsert = lInsert & " COD_RAMO,"
    lInsert = lInsert & " NUM_PED,"
    lInsert = lInsert & " NUM_APOL,"
    lInsert = lInsert & " NUM_ENDO,"
    lInsert = lInsert & " NUM_PI,"
    lInsert = lInsert & " NUM_ITEM,"
    lInsert = lInsert & " COD_COBERT,"
    lInsert = lInsert & " TIPO_TEXTO,"
    lInsert = lInsert & " COD_TEXTO,"
    lInsert = lInsert & " DESC_TEXTO,"
    lInsert = lInsert & " NUM_SEQ"
    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    'If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
    '    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    'End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tipo_Texto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Texto) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Desc_Texto, True) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_SEQ) & ")"

    mfFormatarInclusaoSQL = lInsert
    Exit Function

End Function

Private Function mfFormatarExclusaoSQL(pobjstc As stcA46V739B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       Optional ByVal CodTab As Integer) As String
    Dim lDelete As String 'Query.

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos WHERE "
    lDelete = lDelete & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    'If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
    '    lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    'End If
    If pobjstc.Num_Item > 0 Then
        lDelete = lDelete & " AND Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If

    mfFormatarExclusaoSQL = lDelete
    Exit Function

End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V739B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String 'Query.

    lSelect = "SELECT DISTINCT " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.NOSSO_NUMERO, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Cod_Org, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Cod_Ramo, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Num_Ped, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Num_Apol, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Num_Endo, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Num_PI, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Num_Item, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Cod_Cobert, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Tipo_Texto, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Cod_Texto, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Desc_Texto, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Num_Seq"
    
    lSelect = lSelect & " From " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos WITH (NOLOCK) "
    
    If pobjstc.Cod_Cobert > 0 Then
        lSelect = lSelect & " Where  " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Cod_Cobert = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert)
    End If
    
    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If
    'If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
    '    lSelect = M46V999.gfWhereAnd(lSelect) & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    'End If
    
    If pobjstc.Num_Item > 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If
    
    If pobjstc.Num_Apol <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol)
        lSelect = M46V999.gfWhereAnd(lSelect) & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.num_endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo)
    End If
    
    lSelect = lSelect & " order by " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.num_item, " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.cod_cobert, " & pPropTabela & pNomeTabImport & "Tab_Ped_Textos.Num_Seq"
    
    mfFormatarConsultaSQL = lSelect

End Function

Private Sub mpLimparColecao()

    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V739B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)
    
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("NOsso_Numero"), tpoString)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_item"), tpoNumerico)
    pobjstc.Cod_Cobert = M46V999.gfTratarNulo(prst("Cod_Cobert"), tpoNumerico)
    pobjstc.Tipo_Texto = M46V999.gfTratarNulo(prst("Tipo_Texto"), tpoNumerico)
    pobjstc.Cod_Texto = M46V999.gfTratarNulo(prst("Cod_Texto"), tpoNumerico)
    pobjstc.Desc_Texto = M46V999.gfTratarNulo(prst("Desc_Texto"), tpoString)
    pobjstc.NUM_SEQ = M46V999.gfTratarNulo(prst("Num_Seq"), tpoNumerico)

End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V739B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfIncluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lsSQl As String 'Query.
    Dim i As Integer

    On Local Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro
  
    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, 2)

    'Adicionado controle pois estava incluindo duplicado na base de dados
    For i = 1 To Colecao.Count
        If Colecao(i).Class_ID = pobjstc.Class_ID Then
            GoTo Proximo
        End If
    Next
    
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro

    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID

Proximo:
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    mfIncluir = True
    Exit Function

ControleErro:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V739A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V739A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function

Public Function mfExcluir(ByRef pobjstc As stcA46V739B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean

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
    
    Exit Function
    
ControleErro:
    mfExcluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V739A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V73A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V739B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V739B    'Objeto da classe de estrutura stcA46V739B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    If M46V999.gfObtRegistro(ConexaoBD, mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport), _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V739B
        Set lobjstc.M46V999 = M46V999

        Call mpCarregarAtributos(lobjstc, lrst)
        Set lobjstc.ObjPai = pobjstc.ObjPai
                
        Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID
        lrst.MoveNext
    Loop
    mfCarregarColecao = True

SAIDA:
    GoSub FechaObj
    Exit Function

ControleErro:
    mfCarregarColecao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V739A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V739A.mfCarregarColecao ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    GoTo SAIDA
FechaObj:
    Set lobjstc = Nothing
    Set lrst = Nothing
    Return
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




