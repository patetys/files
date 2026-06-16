VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V734A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V734A
' Sistema: SyasMult
' Analista: Marcio Oliveira
' ======================================================================================================================================================
Private mCnDB As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V734B
Public UltimoErro As String      ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V734A"
End Property

Private Function mfFormatarInclusaoSQL(pobjstc As stcA46V734B) As String
    Dim lInsert As String

    lInsert = " INSERT INTO Tbm_Ass24h ( "
    lInsert = lInsert & " Cod_Tarif,"
    lInsert = lInsert & " Cod_Sequencia,"
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " Cod_Modalidade,"
    lInsert = lInsert & " Cod_Clau,"
    lInsert = lInsert & " Premio"
    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tarif) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Sequencia) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Modalidade) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Clau) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Premio) & ")"

    mfFormatarInclusaoSQL = lInsert

End Function

Private Function mfFormatarExclusaoSQL(pobjstc As stcA46V734B) As String

    Dim lDelete As String

    lDelete = " DELETE from Tbm_Ass24h "
    lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Tarif= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tarif)
    lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Sequencia= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Sequencia)
    lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Ramo= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo)
    lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Modalidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Modalidade)
    If pobjstc.Cod_Clau > 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Clau= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Clau)
    End If

    mfFormatarExclusaoSQL = lDelete

End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Public Function gfIncluir(pobjstc As stcA46V734B) As Boolean
    Dim lInsert As String

    On Local Error GoTo ControleErros

    If CnDB Is Nothing Then GoTo ControleErros

    lInsert = mfFormatarInclusaoSQL(pobjstc)

    If M46V999.gfExeSQL(CnDB, lInsert, UltimoErro) <> 0 Then GoTo ControleErros

    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID

    gfIncluir = True

    Exit Function

ControleErros:

    gfIncluir = False

    If Not M46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If

End Function
Public Function gfExcluir(pobjstc As stcA46V734B) As Boolean

    Dim lDelete As String

    On Local Error GoTo ControleErros

    lDelete = mfFormatarExclusaoSQL(pobjstc)

    If CnDB Is Nothing Then GoTo ControleErros

    If M46V999.gfExeSQL(CnDB, lDelete, UltimoErro) <> 0 Then GoTo ControleErros

    gfExcluir = True

    Exit Function

ControleErros:

    gfExcluir = False

    If Not M46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If

End Function
Public Function gfCarregarColecao(ByRef pobjstc As stcA46V734B) As Boolean
    Dim lobjstc As stcA46V734B          'Objeto da classe de estrutura stcA46V734B.
    Dim lrst As clsYasRecordSet   'Registro.

    On Error GoTo ControleErro

    If CnDB Is Nothing Then GoTo ControleErro

    If M46V999.gfObtRegistro(CnDB, mfFormatarConsultaSQL(pobjstc), _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V734B
        Set lobjstc.M46V999 = M46V999

        Call mpCarregarAtributos(lobjstc, lrst)
        Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID
        lrst.MoveNext
    Loop

    gfCarregarColecao = True
    Exit Function
ControleErro:
    gfCarregarColecao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
    End If
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V734B) As String
    Dim lSelect As String                'Query.

    lSelect = "SELECT * FROM Tbm_Ass24h  with(nolock) "
    lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Tarif= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tarif)
    lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Sequencia= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Sequencia)
    lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Ramo= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo)
    lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Modalidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Modalidade)
    If pobjstc.Cod_Clau > 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Clau= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Clau)
    End If
    lSelect = lSelect & " ORDER BY Cod_Ordem ASC, Cod_Clau ASC"
    mfFormatarConsultaSQL = lSelect
End Function


Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V734B, prst As clsYasRecordSet)
    pobjstc.Cod_Tarif = M46V999.gfTratarNulo(prst("Cod_Tarif"), tpoNumerico)
    pobjstc.Cod_Sequencia = M46V999.gfTratarNulo(prst("Cod_Sequencia"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Cod_Modalidade = M46V999.gfTratarNulo(prst("Cod_Modalidade"), tpoNumerico)
    pobjstc.Cod_Clau = M46V999.gfTratarNulo(prst("Cod_Clau"), tpoNumerico)
    pobjstc.Premio = M46V999.gfTratarNulo(prst("Premio"), tpoNumerico)

End Sub

Private Sub Class_Initialize()
    Set Colecao = New Collection
    Set mCnDB = Nothing
    Set mClsM46V999 = Nothing
End Sub
Private Sub Class_Terminate()
    Set Colecao = Nothing
    Call M46V999.gpFecha2(mCnDB)
    Set mCnDB = Nothing
    Set mClsM46V999 = Nothing
End Sub

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


Public Property Get CnDB() As ADODB.Connection

    If mCnDB Is Nothing Then
        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, mCnDB, _
                                 UltimoErro) = False Then Set mCnDB = Nothing
    End If

Retorno:
    Set CnDB = mCnDB
End Property



