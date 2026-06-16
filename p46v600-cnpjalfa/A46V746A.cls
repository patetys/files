VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V746A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V746A - Tab_Ped_Liberacao_Serasa
' Data/Hora: 06/09/2016
' Sistema: Multi
' Rodrigo P S Rogério
' ======================================================================================================================================================

Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public UltimoErro As String             ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V746A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V746B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lInsert As String 'Query.
    
    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Liberacao_Serasa ( "
    lInsert = lInsert & " NUM_PI,"
    lInsert = lInsert & " NOSSO_NUMERO,"
    lInsert = lInsert & " COD_USER_ALTER,"
    lInsert = lInsert & " DAT_ALTER"
    
    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User_Alt) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Alter) & ")"
    mfFormatarInclusaoSQL = lInsert
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V746B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String 'Query.
    lSelect = "SELECT * FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Liberacao_Serasa with(nolock)"
    lSelect = lSelect & " Where Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    
    mfFormatarConsultaSQL = lSelect
End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V746B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String 'Query.
    lSelect = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Liberacao_Serasa"
    lSelect = lSelect & " Where Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
   
    mfFormatarExclusaoSQL = lSelect
End Function

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V746B, prst As clsYasRecordSet)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoString)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Cod_User_Alt = M46V999.gfTratarNulo(prst("Cod_User_Alter"), tpoNumerico)
    pobjstc.Dat_Alter = M46V999.gfTratarNulo(prst("Dat_Alter"), tpoNumerico)
End Sub

Public Function mfCarregar(ByRef pobjstc As stcA46V746B, _
                           ByVal pPropTabela As String, _
                           ByVal pNomeTabImport As String) As Boolean
    Dim lrst As clsYasRecordSet         'Registro.
    Dim lsSQl As String                 'SQL.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErros
    
    If Not lrst.EOF Then
        Call mpCarregarAtributos(pobjstc, lrst)
    Else
        Set pobjstc = Nothing
    End If
    
    Set lrst = Nothing
    mfCarregar = True
    
    Exit Function
ControleErros:
    Set lrst = Nothing
    mfCarregar = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V746A.mfCarregar ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V746A.mfCarregar ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfIncluir(ByRef pobjstc As stcA46V746B, _
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
    If pobjstc Is Nothing Then
        mfIncluir = True
        Exit Function
    End If
    
Retorno:
    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
    
    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
    mfIncluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErros:
    mfIncluir = False
    Call M46V999.gpGraLog(0, "A46V746A.mfIncluir ERRO - " & UltimoErro & _
                             vbCrLf & " SQL: " & lsSQl)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function

Public Function mfExcluir(ByRef pobjstc As stcA46V746B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfExcluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lsSQl As String           'Query.

    On Local Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    If pobjstc Is Nothing Then
        mfExcluir = True
        Exit Function
    End If
    
Retorno:
    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
    mfExcluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErros:
    mfExcluir = False
    Call M46V999.gpGraLog(0, "A46V746A.mfExcluir ERRO - " & UltimoErro & _
                             vbCrLf & " SQL: " & lsSQl)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function

Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
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

