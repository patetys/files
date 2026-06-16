VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V717A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V717A
' Data/Hora: 26/09/2005 - 17:03:11
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V717B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V717A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V717B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       Optional ByVal CodTab As Integer) As String
    Dim lInsert As String                'Query.

    If CodTab = 2 Then GoTo Tabela2

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Insp ( "
    lInsert = lInsert & " Num_Item,"
    lInsert = lInsert & " Nosso_Numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " Cod_Org,"
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " Num_Ped,"
    lInsert = lInsert & " Num_Insp,"
    lInsert = lInsert & " Dat_Cadastro,"
    lInsert = lInsert & " Pess_Cont,"
    lInsert = lInsert & " Num_Logr_Risco,"
    lInsert = lInsert & " Num_Tel"

    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Insp) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Cadastro) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Pess_Cont) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Logr_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Tel) & ")"

    GoTo Fim
    
Tabela2:
    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Insp_Compl ( "
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " Nosso_numero, "
    lInsert = lInsert & " Num_Item,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " Num_PI,"
    lInsert = lInsert & " Pess_Contato_2,"
    lInsert = lInsert & " Pess_Contato_3,"
    lInsert = lInsert & " Num_Tel_2,"
    lInsert = lInsert & " Num_Tel_3,"
    lInsert = lInsert & " Dsc_Email,"
    lInsert = lInsert & " Dsc_Ponto_Ref"

    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.PESS_CONTATO_2) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.PESS_CONTATO_3) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_TEL_2) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_TEL_3) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.DSC_EMAIL) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.DSC_PONTO_REF) & ")"

Fim:
    mfFormatarInclusaoSQL = lInsert
End Function


Private Function mfFormatarAlteracaoSQL(ByRef pobjstc As stcA46V717B, _
                                        ByVal pPropTabela As String, _
                                        ByVal pNomeTabImport As String, _
                                        Optional ByVal CodTab As Integer) As String
    Dim lUpdate As String                'Query.

    If CodTab = 2 Then GoTo Tabela2

    lUpdate = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ped_Insp Set "
    lUpdate = lUpdate & " Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lUpdate = lUpdate & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    lUpdate = lUpdate & " Cod_Org = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lUpdate = lUpdate & " Cod_Ramo = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lUpdate = lUpdate & " Num_Ped = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lUpdate = lUpdate & " Num_Insp = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Insp) & ","
    lUpdate = lUpdate & " Dat_Cadastro = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Cadastro) & ","
    lUpdate = lUpdate & " Pess_Cont = " & M46V999.gfFormatarCamposSQL(pobjstc.Pess_Cont) & ","
    lUpdate = lUpdate & " Num_Logr_Risco= " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Logr_Risco) & ","
    lUpdate = lUpdate & " Num_Tel = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Tel)

    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    lUpdate = lUpdate & " AND Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

GoTo Fim

Tabela2:
    lUpdate = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ped_Insp_Compl Set "
    lUpdate = lUpdate & " Cod_Ramo = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lUpdate = lUpdate & " Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lUpdate = lUpdate & " Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lUpdate = lUpdate & " Pess_Contato_2 = " & M46V999.gfFormatarCamposSQL(pobjstc.PESS_CONTATO_2) & ","
    lUpdate = lUpdate & " Pess_Contato_3 = " & M46V999.gfFormatarCamposSQL(pobjstc.PESS_CONTATO_3) & ","
    lUpdate = lUpdate & " Num_Tel_2 = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_TEL_2) & ","
    lUpdate = lUpdate & " Num_Tel_3 = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_TEL_3) & ","
    lUpdate = lUpdate & " Dsc_Email = " & M46V999.gfFormatarCamposSQL(pobjstc.DSC_EMAIL) & ","
    lUpdate = lUpdate & " Dsc_Ponto_Ref = " & M46V999.gfFormatarCamposSQL(pobjstc.DSC_PONTO_REF)

    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    lUpdate = lUpdate & " AND Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

Fim:
    mfFormatarAlteracaoSQL = lUpdate
End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V717B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       Optional ByVal CodTab As Integer) As String
    Dim lDelete As String                'Query.

    If CodTab = 2 Then GoTo Tabela2

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Insp WHERE "
    lDelete = lDelete & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Item > 0 Then
        lDelete = lDelete & " and Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If

GoTo Fim
    
Tabela2:
    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Insp_Compl WHERE "
    lDelete = lDelete & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Item > 0 Then
        lDelete = lDelete & " and Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If

Fim:
    mfFormatarExclusaoSQL = lDelete
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V717B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String                'Query.
    lSelect = "SELECT DISTINCT Insp.*,Compl.NUM_PI,Compl.PESS_CONTATO_2,Compl.PESS_CONTATO_3,Compl.NUM_TEL_2,Compl.NUM_TEL_3,Compl.DSC_EMAIL, "
    lSelect = lSelect & " Compl.DSC_PONTO_REF FROM "
    lSelect = lSelect & pPropTabela & pNomeTabImport & "Tab_Ped_Insp as Insp with(nolock) "
    lSelect = lSelect & " Left join " & pPropTabela & pNomeTabImport & "Tab_Ped_Insp_Compl as Compl with(nolock)  "
    lSelect = lSelect & " on Compl.Num_Item = Insp.Num_Item "
    lSelect = lSelect & " and Compl.NOSSO_NUMERO = Insp.NOSSO_NUMERO "
    
    lSelect = lSelect & " where Insp.Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Insp.Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Item > 0 Then
        lSelect = lSelect & " and Insp.Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If
    mfFormatarConsultaSQL = lSelect
End Function
Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V717B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)

    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_Item"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Num_Insp = M46V999.gfTratarNulo(prst("Num_Insp"), tpoNumerico)
    pobjstc.Dat_Cadastro = M46V999.gfTratarNulo(prst("Dat_Cadastro"), tpoNumerico)
    pobjstc.Pess_Cont = M46V999.gfTratarNulo(prst("Pess_Cont"), tpoString)
    pobjstc.Num_Tel = M46V999.gfTratarNulo(prst("Num_Tel"), tpoString)
    pobjstc.Num_Logr_Risco = M46V999.gfTratarNulo(prst("Num_Logr_Risco"), tpoNumerico)
    
    If M46V999.gfTratarNulo(prst("NUM_PI"), tpoNumerico) <> 0 Then
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("NUM_PI"), tpoNumerico)
    End If
    
    pobjstc.PESS_CONTATO_2 = M46V999.gfTratarNulo(prst("PESS_CONTATO_2"), tpoString)
    pobjstc.PESS_CONTATO_3 = M46V999.gfTratarNulo(prst("PESS_CONTATO_3"), tpoString)
    pobjstc.NUM_TEL_2 = M46V999.gfTratarNulo(prst("NUM_TEL_2"), tpoString)
    pobjstc.NUM_TEL_3 = M46V999.gfTratarNulo(prst("NUM_TEL_3"), tpoString)
    pobjstc.DSC_EMAIL = M46V999.gfTratarNulo(prst("DSC_EMAIL"), tpoString)
    pobjstc.DSC_PONTO_REF = M46V999.gfTratarNulo(prst("DSC_PONTO_REF"), tpoString)
End Sub

Public Function mfIncluir(ByRef pobjstc As stcA46V717B, _
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
    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros

    If pobjstc.NUM_PI <> 0 Then
        lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, 2)
        If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
    End If

    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfIncluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErros:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V717A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V717A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function
Public Function mfAlterar(ByRef pobjstc As stcA46V717B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean
    Dim lsSQl As String           'Query.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros

    lsSQl = mfFormatarAlteracaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
    
    lsSQl = mfFormatarAlteracaoSQL(pobjstc, pPropTabela, pNomeTabImport, 2)
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
        Call M46V999.gpGraLog(0, "A46V717A.mfAlterar ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V717A.mfAlterar ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function
Public Function mfExcluir(ByRef pobjstc As stcA46V717B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean
    Dim lsSQl As String           'Query.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
    
    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, 2)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
    If Colecao.Count > 0 Then
        On Error Resume Next
        Colecao.Remove (pobjstc.Class_ID)
        On Error GoTo 0
    End If
    Set pobjstc = Nothing
    mfExcluir = True
    Exit Function
ControleErros:
    mfExcluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V717A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V717A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function
Public Function mfCarregarColecao(ByRef pobjstc As stcA46V717B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V717B          'Objeto da classe de estrutura stcA46V717B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErros
    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V717B
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
        Call M46V999.gpGraLog(0, "A46V717A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V717A.mfCarregarColecao ERRO - " & UltimoErro & _
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
