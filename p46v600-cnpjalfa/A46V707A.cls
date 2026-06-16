VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V707A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: clsA46V707A --> Tabela Tab_Ped_Ender
' Data/Hora: 31/08/2005 - 11:54:14
' Sistema: Multi
' Thais Juren Gomes

' ======================================================================================================================================================


' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V707B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V707A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V707B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String

    Dim lInsert As String                'Query.
    Dim lAlter As String

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Ender ( "
    lInsert = lInsert & " Nosso_Numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " Num_Ped,"
    lInsert = lInsert & " Cod_Org,"
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " Num_Apol,"
    lInsert = lInsert & " Num_Endo,"
    lInsert = lInsert & " Cod_Tip_Logr,"
    lInsert = lInsert & " Nom_Logr,"
    lInsert = lInsert & " Dsc_Compl_Logr,"
    lInsert = lInsert & " Nom_Bairro,"
    lInsert = lInsert & " Nom_Cid,"
    lInsert = lInsert & " Sig_UF,"
    lInsert = lInsert & " Num_CEP,"
    lInsert = lInsert & " Num_Fone,"
    'Ficha 1000307
    lInsert = lInsert & " Num_PI,"
    lInsert = lInsert & " Cod_Reg_Cons,"
    lInsert = lInsert & " Num_Reg_Cons "
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
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Logr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Logr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Logr) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Bairro) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Cid) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Sig_UF) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_CEP) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Fone) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    'Ficha 1000307
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Reg_Cons) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Reg_Cons) & ")"

    mfFormatarInclusaoSQL = lInsert
End Function


Private Function mfFormatarAlteracaoSQL(ByRef pobjstc As stcA46V707B, _
                                        ByVal pPropTabela As String, _
                                        ByVal pNomeTabImport As String) As String

    Dim lUpdate As String                'Query.

    lUpdate = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ped_Ender Set "
    lUpdate = lUpdate & " Num_Ped = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lUpdate = lUpdate & " Cod_Org = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lUpdate = lUpdate & " Cod_Ramo = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lUpdate = lUpdate & " Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lUpdate = lUpdate & " Num_Endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lUpdate = lUpdate & " Cod_Tip_Logr = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Logr) & ","
    lUpdate = lUpdate & " Nom_Logr = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Logr) & ","
    lUpdate = lUpdate & " Dsc_Compl_Logr = " & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Logr) & ","
    lUpdate = lUpdate & " Nom_Bairro = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Bairro) & ","
    lUpdate = lUpdate & " Nom_Cid = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Cid) & ","
    lUpdate = lUpdate & " Sig_UF = " & M46V999.gfFormatarCamposSQL(pobjstc.Sig_UF) & ","
    lUpdate = lUpdate & " Num_CEP = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_CEP) & ","
    lUpdate = lUpdate & " Num_Fone = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Fone) & ","
    lUpdate = lUpdate & " Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    'Ficha 1000307
    lUpdate = lUpdate & " Cod_Reg_Cons = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Reg_Cons) & ","
    lUpdate = lUpdate & " Num_Reg_Cons = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Reg_Cons)

    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " NOsso_NUmero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarAlteracaoSQL = lUpdate
End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V707B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String

    Dim lDelete As String                'Query.

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Ender WHERE "
    lDelete = lDelete & " Nosso_NUmero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    mfFormatarExclusaoSQL = lDelete
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V707B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String                'Query.

    lSelect = "SELECT * FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Ender with(nolock)"
    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If
    If pobjstc.Num_Apol <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol)
        lSelect = M46V999.gfWhereAnd(lSelect) & " Num_Endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo)
    End If
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarConsultaSQL = lSelect
End Function
Private Sub mpLimparColecao()

    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V707B, ByRef prst As clsYasRecordSet, ByVal pNomeTabImport As String)
    Dim lTemp As Integer
    Dim lAlter As String
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Cod_Tip_Logr = M46V999.gfTratarNulo(prst("Cod_Tip_Logr"), tpoString)
    pobjstc.Nom_Logr = M46V999.gfTratarNulo(prst("Nom_Logr"), tpoString)
    pobjstc.Dsc_Compl_Logr = M46V999.gfTratarNulo(prst("Dsc_Compl_Logr"), tpoString)
    pobjstc.Nom_Bairro = M46V999.gfTratarNulo(prst("Nom_Bairro"), tpoString)
    pobjstc.Nom_Cid = M46V999.gfTratarNulo(prst("Nom_Cid"), tpoString)
    pobjstc.Sig_UF = M46V999.gfTratarNulo(prst("Sig_UF"), tpoString)
    pobjstc.Num_CEP = M46V999.gfTratarNulo(prst("Num_CEP"), tpoNumerico)
    pobjstc.Num_Fone = M46V999.gfTratarNulo(prst("Num_Fone"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)

    'Ficha 1000307
    On Error Resume Next
    lTemp = IsNull(prst("Cod_Reg_Cons"))
    If Err.Number = 3265 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.gPrograma <> "D046O003" Then
        'Crio o campo Cod_Reg_Cons na tabela TAB_PED_ENDER
        lAlter = " alter table " & pNomeTabImport & "tab_ped_ender add COLUMN  Cod_Reg_Cons int;"
        If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(2, UltimoErro)
            Exit Sub
        End If
    Else
        pobjstc.Cod_Reg_Cons = M46V999.gfTratarNulo(prst("Cod_Reg_Cons"), tpoNumerico)
    End If

    On Error Resume Next
    lTemp = IsNull(prst("Num_Reg_Cons"))
    If Err.Number = 3265 And M46V999.gTipUsuario = e_Tipo_Usu_Corretor And M46V999.gPrograma <> "D046O003" Then
        'Crio o campo Num_Reg_Cons na tabela TAB_PED_ENDER
        lAlter = " alter table " & pNomeTabImport & "tab_ped_ender add COLUMN  Num_Reg_Cons char(10);"
        If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(2, UltimoErro)
            Exit Sub
        End If
    Else
        pobjstc.Num_Reg_Cons = M46V999.gfTratarNulo(prst("Num_Reg_Cons"), tpoString)
    End If


End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V707B, _
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
Retorno:
    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfIncluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function

ControleErro:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        If Err.Number = 3022 Then
            UltimoErro = "FTE0003 - Documento " & pobjstc.nosso_numero & " já existente."
        Else
            UltimoErro = Err.Description
        End If
    Else
        If InStr(1, UCase(UltimoErro), "'COD_REG_CONS'") > 0 Then
            lsSQl = " alter table " & pNomeTabImport & "tab_ped_ender add COLUMN  Cod_Reg_Cons int;"
            If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
                Exit Function
            End If
            GoTo Retorno
        End If
        If InStr(1, UCase(UltimoErro), "'NUM_REG_CONS'") > 0 Then
            lsSQl = " alter table " & pNomeTabImport & "tab_ped_ender add COLUMN  Num_Reg_Cons char(10);"
            If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
                Exit Function
            End If
            GoTo Retorno
        End If
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)

End Function
Public Function mfAlterar(ByRef pobjstc As stcA46V707B, _
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
Public Function mfExcluir(ByRef pobjstc As stcA46V707B, _
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
Public Function mfCarregarColecao(ByRef pobjstc As stcA46V707B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V707B          'Objeto da classe de estrutura stcA46V707B.
    Dim lrst As clsYasRecordSet   'Registro.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    If M46V999.gfObtRegistro(ConexaoBD, mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport), _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V707B
        Set lobjstc.M46V999 = M46V999

        Call mpCarregarAtributos(lobjstc, lrst, pNomeTabImport)
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
