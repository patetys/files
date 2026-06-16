VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V709A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V709A --Tab_Ped_Com
' Data/Hora: 05/09/2005 - 11:24:45
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V709B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V709A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V709B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lstrTMP As String                'Query.

    lstrTMP = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Com ( "
    lstrTMP = lstrTMP & " Cod_User,"
    lstrTMP = lstrTMP & " Nosso_Numero,"
    lstrTMP = lstrTMP & " Dat_Alt"

    lstrTMP = lstrTMP & ") VALUES ("
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Alt) & ")"

    mfFormatarInclusaoSQL = lstrTMP
End Function

Private Function mfFormatarExclusaoSQL(ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lstrTMP As String                'Query.

    lstrTMP = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Com WHERE "
    lstrTMP = lstrTMP & " Cod_User = " & M46V999.gfFormatarCamposSQL(M46V999.gNomUsuario6)

    mfFormatarExclusaoSQL = lstrTMP
End Function

Private Function mfFormatarConsultaSQL(ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lstrTMP As String                'Query.

    lstrTMP = "SELECT * FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Com with(nolock)"
    lstrTMP = lstrTMP & " Where cod_User = " & M46V999.gfFormatarCamposSQL(M46V999.gNomUsuario6)
    mfFormatarConsultaSQL = lstrTMP

End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V709B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)
    pobjstc.Cod_User = M46V999.gfTratarNulo(prst("Cod_User"), tpoString)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Dat_Alt = M46V999.gfTratarNulo(prst("Dat_Alt"), tpoNumerico)
End Sub

Public Function mfIncluir(ByRef pobjstc As stcA46V709B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean
    Dim lsSQl As String           'Query.

    On Local Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)

    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro

    mfIncluir = True
    Exit Function

ControleErro:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V709A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V709A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfExcluir(ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean
    Dim lsSQl As String           'Query.
    Dim lExecutar As Boolean

    On Local Error GoTo ControleErro
    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarExclusaoSQL(pPropTabela, pNomeTabImport)

    lExecutar = True
    If M46V999.EhTeste Then
        If MsgBox("Pode excluir tab_ped_com?", vbYesNo + vbDefaultButton1) <> vbYes Then
            lExecutar = False
        End If
    End If

    If lExecutar Then
        If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
    End If

    mfExcluir = True
    Exit Function
ControleErro:
    mfExcluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V709A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V709A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfCarregarColecao(ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V709B      'Objeto da classe de estrutura stcA46V709B.
    Dim lrst As clsYasRecordSet    'Registro.
    Dim lsSQl As String         'SQL.

    On Error GoTo ControleErro
    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarConsultaSQL(pPropTabela, pNomeTabImport)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V709B
        Call mpCarregarAtributos(lobjstc, lrst)
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
        Call M46V999.gpGraLog(0, "A46V709A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V709A.mfCarregarColecao ERRO - " & UltimoErro & _
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
