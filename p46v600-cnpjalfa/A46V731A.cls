VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V731A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46v731A
' Data/Hora: 16/06/2010 - 10:20:00
' Sistema: P46V600
' Analista: Sandra M. De Rosa
' ======================================================================================================================================================

Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46v731B
Public UltimoErro As String    ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V731A"
End Property


Private Function mfFormatarInclusaoSQL(pobjstc As stcA46V731B) As String

    Dim strTMP As String

    strTMP = " INSERT INTO Tab_Trans_Cobert_Insp ( "
    strTMP = strTMP & " NOSSO_NUMERO,"
    strTMP = strTMP & " NUM_PED,"
    strTMP = strTMP & " NUM_ITEM,"
    strTMP = strTMP & " NUM_COTAC,"
    strTMP = strTMP & " NUM_INSP,"
    strTMP = strTMP & " COD_COBERT,"
    strTMP = strTMP & " LMG_COBERT"
    strTMP = strTMP & ") VALUES ("
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cotac) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Insp) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Lmg_Cobert) & ")"

    mfFormatarInclusaoSQL = strTMP

End Function

Private Function mfFormatarExclusaoSQL(pobjstc As stcA46V731B) As String

    Dim strTMP As String

    strTMP = " DELETE from Tab_Trans_Cobert_Insp "
    strTMP = strTMP & " where Nosso_Numero = "
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    strTMP = strTMP & " and Num_Item = "
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)

    mfFormatarExclusaoSQL = strTMP

End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Public Function gfIncluir(pobjstc As stcA46V731B) As Boolean

    Dim lCnSQL As ADODB.Connection
    Dim sSql As String
    Dim pMensagem As String

    On Local Error GoTo ControleErros

    'sSql = mfFormatarExclusaoSQL(pObjstc)

    If M46V999.gfAbrBasDados("RAMOSDIVERSOS", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, lCnSQL, pMensagem) = False Then
        Exit Function
    End If

    'If M46V999.gfExeSQL(lCnSQL, sSql, UltimoErro) <> 0 Then GoTo ControleErros

    sSql = mfFormatarInclusaoSQL(pobjstc)

    If M46V999.gfExeSQL(lCnSQL, sSql, UltimoErro) <> 0 Then GoTo ControleErros

    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID

    gfIncluir = True

    Exit Function

ControleErros:

    gfIncluir = False

    If Not M46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If

End Function
Public Function gfExcluir(pobjstc As stcA46V731B) As Boolean

    Dim lCnSQL As ADODB.Connection
    Dim sSql As String
    Dim pMensagem As String

    On Local Error GoTo ControleErros

    sSql = mfFormatarExclusaoSQL(pobjstc)

    If M46V999.gfAbrBasDados("RAMOSDIVERSOS", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, lCnSQL, pMensagem) = False Then
        Exit Function
    End If

    If M46V999.gfExeSQL(lCnSQL, sSql, UltimoErro) <> 0 Then GoTo ControleErros

    gfExcluir = True

    Exit Function

ControleErros:

    gfExcluir = False

    If Not M46V999.gfPreenchido(UltimoErro) Then
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
