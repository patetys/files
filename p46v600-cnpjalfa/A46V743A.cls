VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V743A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V743A
' Data/Hora: 19/09/2005 - 10:20:46
' Sistema: Multi
' Carlos Araújo.
' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V743B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V743A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V743B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       ByVal TabAtual As Integer) As String
    Dim lInsert As String                'Query.

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "TAB_MOV_AGEN ( "
    lInsert = lInsert & " NOSSO_NUMERO,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " NUM_SEQ,"
    lInsert = lInsert & " NUM_PI,"
    lInsert = lInsert & " NUM_APOL_MOV,"
    lInsert = lInsert & " NUM_COMPL_MOV,"
    lInsert = lInsert & " REF_MOV,"
    lInsert = lInsert & " TIPO_MOV,"
    lInsert = lInsert & " VL_PREMIO_LIQ_MOV,"
    lInsert = lInsert & " QTD_ITEM_MOV,"
    lInsert = lInsert & " DAT_GRAV,"
    lInsert = lInsert & " COD_USER "

    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","        'crda erro compilação
    End If
    
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_SEQ) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_APOL_MOV) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_COMPL_MOV) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.REF_MOV) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.TIPO_MOV) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.VL_PREMIO_LIQ_MOV) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.QTD_ITEM_MOV) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.DAT_GRAV) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(M46V999.gNomUsuario) & ")"

    mfFormatarInclusaoSQL = lInsert

End Function


Private Function mfFormatarAlteracaoSQL(ByRef pobjstc As stcA46V743B, _
                                        ByVal pPropTabela As String, _
                                        ByVal pNomeTabImport As String, _
                                        ByVal TabAtual As Integer) As String

    Dim lUpdate As String                'Query.

    lUpdate = "UPDATE " & pPropTabela & pNomeTabImport & "TAB_MOV_AGEN Set "
    lUpdate = lUpdate & " NUM_SEQ = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_SEQ) & ","
    lUpdate = lUpdate & " NUM_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lUpdate = lUpdate & " NUM_APOL_MOV = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_APOL_MOV) & ","
    lUpdate = lUpdate & " NUM_COMPL_MOV = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_COMPL_MOV) & ","
    lUpdate = lUpdate & " REF_MOV = " & M46V999.gfFormatarCamposSQL(pobjstc.REF_MOV) & ","
    lUpdate = lUpdate & " TIPO_MOV = " & M46V999.gfFormatarCamposSQL(pobjstc.TIPO_MOV) & ","
    lUpdate = lUpdate & " VL_PREMIO_LIQ_MOV = " & M46V999.gfFormatarCamposSQL(pobjstc.VL_PREMIO_LIQ_MOV) & ","
    lUpdate = lUpdate & " QTD_ITEM_MOV = " & M46V999.gfFormatarCamposSQL(pobjstc.QTD_ITEM_MOV) & ","
    lUpdate = lUpdate & " DAT_GRAV = " & M46V999.gfFormatarCamposSQL(pobjstc.DAT_GRAV) & ","
    lUpdate = lUpdate & " Cod_User = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User) & ","
    
    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarAlteracaoSQL = lUpdate

End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V743B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       ByVal TabAtual As Integer) As String
    Dim lDelete As String                'Query.

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "TAB_MOV_AGEN WHERE "
    lDelete = lDelete & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarExclusaoSQL = lDelete

End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V743B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String


    lSelect = "SELECT * FROM " & pPropTabela & pNomeTabImport & "TAB_MOV_AGEN"
    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lSelect = lSelect & " Where Nosso_numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " E.Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If val(pobjstc.NUM_PI) <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " E.Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)
    End If
'    If pobjstc.Num_Apol <> 0 Then
'        lSelect = M46V999.gfWhereAnd(lSelect) & " E.Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol)
'        lSelect = M46V999.gfWhereAnd(lSelect) & " E.num_endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo)
'    End If

    mfFormatarConsultaSQL = lSelect
End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V743B, prst As clsYasRecordSet)

    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.NUM_SEQ = M46V999.gfTratarNulo(prst("NUM_SEQ"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("NUM_PI"), tpoNumerico)
    pobjstc.NUM_APOL_MOV = M46V999.gfTratarNulo(prst("NUM_APOL_MOV"), tpoNumerico)
    pobjstc.NUM_COMPL_MOV = M46V999.gfTratarNulo(prst("NUM_COMPL_MOV"), tpoNumerico)
    pobjstc.REF_MOV = M46V999.gfTratarNulo(prst("REF_MOV"), tpoNumerico)
    pobjstc.TIPO_MOV = M46V999.gfTratarNulo(prst("TIPO_MOV"), tpoString)
    pobjstc.VL_PREMIO_LIQ_MOV = M46V999.gfTratarNulo(prst("VL_PREMIO_LIQ_MOV"), tpoNumerico)
    pobjstc.QTD_ITEM_MOV = M46V999.gfTratarNulo(prst("QTD_ITEM_MOV"), tpoNumerico)
    pobjstc.DAT_GRAV = M46V999.gfTratarNulo(prst("DAT_GRAV"), tpoNumerico)
    pobjstc.Cod_User = M46V999.gfTratarNulo(prst("COD_USER"), tpoString)

    On Error Resume Next
'    pobjstc.Num_Apol_Mar = M46V999.gfTratarNulo(prst("Num_Apol_Mar"), tpoNumerico)

    Err = 0

End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V743B, _
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
        Call M46V999.gpGraLog(0, "A46V743A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V743A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function

Public Function mfAlterar(ByRef pobjstc As stcA46V743B, _
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
        Call M46V999.gpGraLog(0, "A46V743A.mfAlterar ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V743A.mfAlterar ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfAlterarNum_PI(ByRef pobjstc As stcA46V743B, _
                                ByVal pPropTabela As String, _
                                ByVal pNomeTabImport As String) As Boolean
    Dim lUpdate As String           'Query.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros

    lUpdate = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ctrl_Emis Set "
    lUpdate = lUpdate & " num_pi = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)
    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErros
    If Colecao.Count > 0 Then
        Colecao.Remove (pobjstc.Class_ID)
    End If
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfAlterarNum_PI = True
    Exit Function
ControleErros:
    mfAlterarNum_PI = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V743A.mfAlterar ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lUpdate)
    Else
        Call M46V999.gpGraLog(0, "A46V743A.mfAlterar ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lUpdate)
    End If
End Function


Public Function mfExcluir(ByRef pobjstc As stcA46V743B, _
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
        Call M46V999.gpGraLog(0, "A46V743A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V743A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V743B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V743B          'Objeto da classe de estrutura stcA46V743B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErros
    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V743B
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
        Call M46V999.gpGraLog(0, "A46V743A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V743A.mfCarregarColecao ERRO - " & UltimoErro & _
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
