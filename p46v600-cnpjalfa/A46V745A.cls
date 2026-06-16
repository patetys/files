VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V745A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V745A - Log_Tab_Ped_Cobert_Batente
' Data/Hora: 16/04/2016
' Sistema: Multi
' Rodrigo P S Rogério
' ======================================================================================================================================================

Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public UltimoErro As String             ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V745A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V745B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lInsert As String 'Query.
    
    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Log_Tab_Ped_Cobert_Batente ( "
    lInsert = lInsert & " NOSSO_NUMERO,"
    lInsert = lInsert & " NUM_ITEM,"
    lInsert = lInsert & " COD_COBERT,"
    lInsert = lInsert & " Usuario,"
    lInsert = lInsert & " Status_Batente,"
    lInsert = lInsert & " Val_Pr_Base_Apol,"
    lInsert = lInsert & " Val_Pr_Pago_Apol,"
    lInsert = lInsert & " Formula_Apol,"
    lInsert = lInsert & " Per_Dif_Batente,"
    lInsert = lInsert & " Per_Dif_Renovacao,"
    lInsert = lInsert & " Val_Pr_Calc,"
    lInsert = lInsert & " Formula"
    
    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Usuario) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Status_Batente) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Base_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Pago_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Formula_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Dif_Batente) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Dif_Renovacao) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Calc) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Formula) & ")"
    mfFormatarInclusaoSQL = lInsert
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V745B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String 'Query.
    lSelect = "SELECT TOP 1 * FROM " & pPropTabela & pNomeTabImport & "Log_Tab_Ped_Cobert_Batente with(nolock)"
    lSelect = lSelect & " Where nosso_numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    lSelect = lSelect & " order by Dat_Calculo DESC"

    mfFormatarConsultaSQL = lSelect
End Function

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V745B, prst As clsYasRecordSet)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_Item"), tpoNumerico)
    pobjstc.Cod_Cobert = M46V999.gfTratarNulo(prst("Cod_Cobert"), tpoNumerico)
    pobjstc.Dat_Calculo = M46V999.gfTratarNulo(prst("Dat_Calculo"), TpoData)
    pobjstc.Usuario = M46V999.gfTratarNulo(prst("Usuario"), tpoString)
    pobjstc.Status_Batente = M46V999.gfTratarNulo(prst("Status_Batente"), tpoNumerico)
    pobjstc.Val_Pr_Base_Apol = M46V999.gfTratarNulo(prst("Val_Pr_Base_Apol"), TpoNumDecimal)
    pobjstc.Val_Pr_Pago_Apol = M46V999.gfTratarNulo(prst("Val_Pr_Pago_Apol"), TpoNumDecimal)
    pobjstc.Formula_Apol = M46V999.gfTratarNulo(prst("Formula_Apol"), tpoString)
    pobjstc.Per_Dif_Batente = M46V999.gfTratarNulo(prst("Per_Dif_Batente"), TpoNumDecimal)
    pobjstc.Per_Dif_Renovacao = M46V999.gfTratarNulo(prst("Per_Dif_Renovacao"), TpoNumDecimal)
    pobjstc.Val_Pr_Calc = M46V999.gfTratarNulo(prst("Val_Pr_Calc"), TpoNumDecimal)
    pobjstc.Formula = M46V999.gfTratarNulo(prst("Formula"), tpoNumerico)
End Sub

Public Function mfCarregar(ByRef pobjstc As stcA46V745B, _
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
        Call M46V999.gpGraLog(0, "A46V745A.mfCarregar ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V745A.mfCarregar ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfIncluir(ByRef pobjstc As stcA46V745B, _
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
    mfIncluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErros:
    mfIncluir = False
    Call M46V999.gpGraLog(0, "A46V745A.mfIncluir ERRO - " & UltimoErro & _
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
