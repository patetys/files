VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V751A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V751A - Tab_Ped_Obs_SISCOTA_Log
' Data/Hora: 19/10/2017
' Sistema: Multi
' Juliano Iwai
' ======================================================================================================================================================

Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection            ' Coleção de objetos do tipo stcA46V749B
Public UltimoErro As String             ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V751A"
End Property

Public Function mfFormatarConsultaSQL(ByVal IdObs As Long) As String
    Dim lSelect As String 'Query.
    lSelect = "SELECT IdLog, IdObs, Data, Tipo, Descricao FROM RamosDiversos..Tab_Ped_SISCOTA_Obs_Log(nolock)"
    lSelect = lSelect & " Where IdObs = " & M46V999.gfFormatarCamposSQL(IdObs)
    lSelect = lSelect & " Order by IdLog DESC "

    mfFormatarConsultaSQL = lSelect
End Function

Public Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V751B, ByVal pLogin As String) As String
    Dim lInsert As String 'Query.
    
    If pobjstc.IdObs <> "" Then
        lInsert = lInsert & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs_LOG "
        lInsert = lInsert & "(IdObs, Data, Tipo, Descricao) VALUES ("
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.IdObs) & ","
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Data) & ","
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tipo) & ","
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Descricao) & ")" & vbCrLf
    End If
End Function

Public Sub mpCarregarAtributos(ByRef pobjstc As stcA46V751B, prst As clsYasRecordSet)
    pobjstc.IdLog = M46V999.gfTratarNulo(prst("IdLog"), tpoNumerico)
    pobjstc.IdObs = M46V999.gfTratarNulo(prst("IdObs"), tpoNumerico)
    pobjstc.Data = M46V999.gfTratarNulo(prst("Data"), TpoData)
    pobjstc.Tipo = M46V999.gfTratarNulo(prst("Tipo"), tpoNumerico)
    pobjstc.Descricao = M46V999.gfTratarNulo(prst("Descricao"), tpoString)
End Sub

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V751B) As Boolean
    Dim lobjstc As stcA46V751B      'Objeto da classe de estrutura stcA46V751B.
    Dim lrst As clsYasRecordSet     'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErros
    
    lsSQl = "mfCarregarColecao: stcA46V751B"
    If Not mfConectar Then GoTo ControleErros
    
    lsSQl = mfFormatarConsultaSQL(pobjstc)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErros
    
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V751B
        
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
        Call M46V999.gpGraLog(0, "A46V751A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V751A.mfCarregarColecao ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfIncluirItem(ByRef pobjstc As stcA46V750B, _
                              ByVal pLogin As String, _
                              ByRef Mensagem As String) As Boolean
    
    Dim lsSQl As String           'Query.
    
    Mensagem = ""
    
    On Local Error GoTo ControleErros
    
    lsSQl = "mfIncluirItem"
    If Not mfConectar Then GoTo ControleErros

    lsSQl = mfFormatarInclusaoSQL(pobjstc, pLogin)
    
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
        GoTo ControleErros
    End If
    
    mfIncluirItem = True
    Exit Function

ControleErros:
    Mensagem = "A46V750A.mfIncluir ERRO - " & UltimoErro & vbCrLf & " SQL: " & lsSQl
    Call M46V999.gpGraLog(0, Mensagem)

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


