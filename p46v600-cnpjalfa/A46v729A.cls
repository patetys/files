VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46v729A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46v729A
' Data/Hora: 28/03/2008 - 09:46:30
' Sistema: P46V101
' Analista: fatiitus
' ======================================================================================================================================================

Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA28V719B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V729A"
End Property

Private Function mfFormatarConsultaSQL(pobjstc As stcA46v729B) As String

    Dim strTMP As String

    strTMP = "Select * from Tab_Triagem_Multi with(nolock) where dat_triagem >= " & pobjstc.Dat_Triagem & "000000"
    strTMP = strTMP & " and dat_triagem <= " & pobjstc.Dat_Triagem & "999999"

    mfFormatarConsultaSQL = strTMP

End Function

Private Function mfFormatarInclusaoSQL(pobjstc As stcA46v729B) As String

    Dim strTMP As String

    strTMP = " INSERT INTO Tab_Triagem_Multi ( "
    strTMP = strTMP & " Num_PI,"
    strTMP = strTMP & " Nosso_Numero,"
    strTMP = strTMP & " Dat_Triagem,"
    strTMP = strTMP & " User_Triagem,"
    strTMP = strTMP & " Status"

    strTMP = strTMP & ") VALUES ("
    strTMP = strTMP & gM46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    strTMP = strTMP & gM46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    strTMP = strTMP & gM46V999.gfFormatarCamposSQL(pobjstc.Dat_Triagem) & ","
    strTMP = strTMP & gM46V999.gfFormatarCamposSQL(pobjstc.User_Triagem) & ","
    strTMP = strTMP & gM46V999.gfFormatarCamposSQL(pobjstc.STATUS) & ")"

    mfFormatarInclusaoSQL = strTMP

End Function

Private Function mfFormatarExclusaoSQL(pobjstc As stcA46v729B) As String

    Dim strTMP As String

    strTMP = " DELETE from Tab_Triagem_Multi "
    strTMP = strTMP & " where Num_PI = "
    strTMP = strTMP & gM46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)

    mfFormatarExclusaoSQL = strTMP

End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46v729B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)
    pobjstc.NUM_PI = gM46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
    pobjstc.nosso_numero = gM46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Dat_Triagem = gM46V999.gfTratarNulo(prst("Dat_Triagem"), tpoNumerico)
    pobjstc.User_Triagem = gM46V999.gfTratarNulo(prst("User_Triagem"), tpoString)
    pobjstc.STATUS = gM46V999.gfTratarNulo(prst("Status"), tpoNumerico)

End Sub
Public Function gfIncluir(pobjstc As stcA46v729B) As Boolean

    Dim mConexaoSQL As ADODB.Connection
    Dim sSql As String
    Dim pMensagem As String

    On Local Error GoTo ControleErros

    sSql = mfFormatarExclusaoSQL(pobjstc)

    If gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gM46V999.gTipDIVERSOS, gM46V999.gEndDIVERSOS, mConexaoSQL, pMensagem) = False Then
        Exit Function
    End If

    If gM46V999.gfExeSQL(mConexaoSQL, sSql, UltimoErro) <> 0 Then GoTo ControleErros

    sSql = mfFormatarInclusaoSQL(pobjstc)

    If gM46V999.gfExeSQL(mConexaoSQL, sSql, UltimoErro) <> 0 Then GoTo ControleErros

    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID

    gfIncluir = True

    Exit Function

ControleErros:

    gfIncluir = False

    If Not gM46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If

End Function
Public Function gfCarregarColecao(pobjstc As stcA46v729B) As Boolean

    Dim objstc As stcA46v729B
    Dim rst As clsYasRecordSet

    On Error GoTo ControleErros
    If Not M46V999.gfObtRegistro(ConexaoBD, mfFormatarConsultaSQL(pobjstc), rst, UltimoErro) Then
        GoTo ControleErros
    End If

    Call mpLimparColecao

    Do While Not rst.EOF
        Set objstc = New stcA46v729B
        Call mpCarregarAtributos(objstc, rst)
        Colecao.Add Item:=objstc, Key:=objstc.Class_ID
        rst.MoveNext
    Loop

    gfCarregarColecao = True

    Exit Function

ControleErros:

    gfCarregarColecao = False

    If Not gM46V999.gfPreenchido(UltimoErro) Then
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

