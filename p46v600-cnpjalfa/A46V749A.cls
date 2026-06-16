VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V749A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V749A - Tab_Ped_Servico_Ass24h
' Data/Hora: 08/11/2016
' Sistema: Multi
' Rodrigo Berçot
' ======================================================================================================================================================

Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection            ' Coleção de objetos do tipo stcA46V749B
Public UltimoErro As String             ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V749A"
End Property

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V749B) As String
    Dim lSelect As String 'Query.
    lSelect = "SELECT Cod_Ass24h, "
    lSelect = lSelect & " Dsc_Servico_Ass24h, "
    lSelect = lSelect & " Cod_Ramo, "
    lSelect = lSelect & " NomeAssistencia FROM Tab_Ped_Servico_Ass24h with(nolock)"
    lSelect = lSelect & " Inner Join Tab_Assistencia24h with(nolock)"
    lSelect = lSelect & " on Cod_Ass24h = CodigoClausula"
    lSelect = lSelect & " and Cod_ramo = CodigoRamo"
    lSelect = lSelect & " Where Cod_Ramo = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo)
    lSelect = lSelect & " and Cod_Tarif = " & M46V999.gCodProduto
        lSelect = lSelect & " Order by Cod_Ass24h,Cod_Sequencia "

    mfFormatarConsultaSQL = lSelect

End Function
Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V749B, prst As clsYasRecordSet)
    pobjstc.Cod_Ass24h = M46V999.gfTratarNulo(prst("Cod_Ass24h"), tpoNumerico)
    pobjstc.Dsc_Servico_Ass24h = M46V999.gfTratarNulo(prst("Dsc_Servico_Ass24h"), tpoString)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.NomeAssistencia = M46V999.gfTratarNulo(prst("NomeAssistencia"), tpoString)
End Sub

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V749B) As Boolean
    Dim lobjstc As stcA46V749B          'Objeto da classe de estrutura stcA46V747B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarConsultaSQL(pobjstc)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErros
    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V749B
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
        Call M46V999.gpGraLog(0, "A46V749A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V749A.mfCarregarColecao ERRO - " & UltimoErro & _
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
