VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V701A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V701A --> Tabela Tab_ped_ult
' Data/Hora: 05/09/2005 - 16:47:03
' Sistema: Multi
' Thais Juren Gomes
' Objetivo:
' ======================================================================================================================================================

' Properties...
' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V701B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V701A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V701B, _
                                       ByVal pPropTabela As String) As String
    Dim lstrTMP As String                'Query.


    lstrTMP = " INSERT INTO " & pPropTabela & "Tab_ped_ult ( "
    lstrTMP = lstrTMP & " Ult_Num_Ped,"
    lstrTMP = lstrTMP & " Cod_User_Alt,"
    lstrTMP = lstrTMP & " Dat_Alt"

    lstrTMP = lstrTMP & ") VALUES ("
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Ult_Num_Ped) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Alt) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Alt) & ")"

    mfFormatarInclusaoSQL = lstrTMP
End Function


Private Function mfFormatarAlteracaoSQL(ByRef pobjstc As stcA46V701B, _
                                        ByVal pPropTabela As String) As String
    Dim lstrTMP As String                'Query.

    lstrTMP = "UPDATE " & pPropTabela & "Tab_ped_ult Set "
    lstrTMP = lstrTMP & " Ult_Num_Ped = " & M46V999.gfFormatarCamposSQL(pobjstc.Ult_Num_Ped) & ","
    lstrTMP = lstrTMP & " Cod_User_Alt = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Alt) & ","
    lstrTMP = lstrTMP & " Dat_Alt = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Alt)

    mfFormatarAlteracaoSQL = lstrTMP
End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V701B, _
                                       ByVal pPropTabela As String) As String
    Dim lstrTMP As String                'Query.

    lstrTMP = "DELETE FROM " & pPropTabela & "Tab_ped_ult WHERE "
    lstrTMP = lstrTMP & " Ult_Num_Ped = " & M46V999.gfFormatarCamposSQL(pobjstc.Ult_Num_Ped)

    mfFormatarExclusaoSQL = lstrTMP
End Function

Private Function mfFormatarConsultaSQL(ByVal pPropTabela As String) As String
    Dim lstrTMP As String                'Query.
    lstrTMP = "SELECT * FROM " & pPropTabela & "Tab_ped_ult with(nolock)"
    lstrTMP = lstrTMP & " WHERE Ult_Num_Ped < 100000"
    mfFormatarConsultaSQL = lstrTMP
End Function
Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V701B, prst As clsYasRecordSet)
    pobjstc.Ult_Num_Ped = M46V999.gfTratarNulo(prst("Ult_Num_Ped"), tpoNumerico)
    pobjstc.Cod_User_Alt = M46V999.gfTratarNulo(prst("Cod_User_Alt"), tpoString)
    pobjstc.Dat_Alt = M46V999.gfTratarNulo(prst("Dat_Alt"), tpoNumerico)
End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V701B, _
                          ByVal pPropTabela As String) As Boolean
    Dim lsSQl As String          'Query.

    On Local Error GoTo ControleErro
    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfIncluir = True
    Exit Function
ControleErro:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V701A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V701A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function
Public Function mfAlterar(ByRef pobjstc As stcA46V701B, _
                          ByVal pPropTabela As String) As Boolean
    Dim lsSQl As String          'Query.

    On Error GoTo ControleErro
    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarAlteracaoSQL(pobjstc, pPropTabela)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro

    mfAlterar = True
    Exit Function
ControleErro:
    mfAlterar = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V701A.mfAlterar ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V701A.mfAlterar ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function
Public Function mfCarregarColecao(ByVal pPropTabela As String) As Boolean
    Dim lobjstc As stcA46V701B      'Objeto da classe de estrutura stcA46V701B.
    Dim lrst As clsYasRecordSet    'Registro.
    Dim lSql As String        'SQL

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro
    lSql = mfFormatarConsultaSQL(pPropTabela)
    If M46V999.gfObtRegistro(ConexaoBD, lSql, _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V701B
        Call mpCarregarAtributos(lobjstc, lrst)
        Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID
        lrst.MoveNext
    Loop
    mfCarregarColecao = True
    Set lobjstc = Nothing
    Set lrst = Nothing
    Exit Function
ControleErro:
    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V701A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lSql)
    Else
        Call M46V999.gpGraLog(0, "A46V701A.mfCarregarColecao ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lSql)
    End If
End Function

Public Function mfCarregarColecaoSyasPortal(ByVal pCodUserAlt As String, _
                                            ByVal pPropTabela As String) As Boolean
    Dim lobjstc As stcA46V701B      'Objeto da classe de estrutura stcA46V701B.
    Dim lrst As clsYasRecordSet    'Registro.
    Dim lSql As String        'SQL

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro
    lSql = "exec " & pPropTabela & "proc_CalNumPedido " & M46V999.gfFormatarCamposSQL(pCodUserAlt)
    If M46V999.gfObtRegistro(ConexaoBD, lSql, _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V701B
        Call mpCarregarAtributos(lobjstc, lrst)
        Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID
        lrst.MoveNext
    Loop
    mfCarregarColecaoSyasPortal = True
    Set lobjstc = Nothing
    Set lrst = Nothing
    Exit Function
ControleErro:
    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecaoSyasPortal = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V701A.mfCarregarColecaoSyasPortal ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lSql)
    Else
        Call M46V999.gpGraLog(0, "A46V701A.mfCarregarColecaoSyasPortal ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lSql)
    End If
End Function

Private Sub Class_Initialize()
    Set Colecao = New Collection
    Set mClsM46V999 = Nothing
End Sub
Private Sub Class_Terminate()
    Set Colecao = Nothing
    Call M46V999.gpFecha2(ConexaoBD)
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
