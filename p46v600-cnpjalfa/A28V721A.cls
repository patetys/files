VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA28V721A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA28V721A
' Sistema: P28V101
' Analista: tjgomes
' ======================================================================================================================================================

' ======================================================================================================================================================
Private mCnDB As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA28V721B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A28V721A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA28V721B, _
                                       ByVal pPropTabela As String) As String

    Dim strTMP As String

    strTMP = " INSERT INTO " & pPropTabela & "P0045400 ( "
    strTMP = strTMP & " CCotacao,"
    strTMP = strTMP & " Nosso_Numero,"
    strTMP = strTMP & " Dat_Cotacao,"
    strTMP = strTMP & " Email,"
    strTMP = strTMP & " Nome,"
    strTMP = strTMP & " Telefone,"
    strTMP = strTMP & " Status,"
    strTMP = strTMP & " Senha,"
    strTMP = strTMP & " CCotacaoOrig,"
    strTMP = strTMP & " MaxDesc,"
    strTMP = strTMP & " Usuario"

    strTMP = strTMP & ") VALUES ("
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.CCotacao) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Cotacao) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Email) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Nome) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Telefone) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.STATUS) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Senha) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.CCotacaoOrig) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.MaxDesc) & ","
    If M46V999.gfPreenchido(pPropTabela) Then
        strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Usuario) & ")"
    Else
        strTMP = strTMP & M46V999.gfFormatarCamposSQL(val(pobjstc.Usuario)) & ")"
    End If



    mfFormatarInclusaoSQL = strTMP
End Function


Private Function mfFormatarAlteracaoSQL(ByRef pobjstc As stcA28V721B, _
                                        ByVal pPropTabela As String) As String

    Dim strTMP As String

    strTMP = "UPDATE " & pPropTabela & "P0045400 Set "
    'strTMP = strTMP & " CCotacao = " & M46V999.gfFormatarCamposSQL(pobjstc.CCotacao) & ","
    'strTMP = strTMP & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.Nosso_Numero) & ","
    strTMP = strTMP & " Dat_Cotacao = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Cotacao) & ","
    strTMP = strTMP & " Email = " & M46V999.gfFormatarCamposSQL(pobjstc.Email) & ","
    strTMP = strTMP & " Nome = " & M46V999.gfFormatarCamposSQL(pobjstc.Nome) & ","
    strTMP = strTMP & " Telefone = " & M46V999.gfFormatarCamposSQL(pobjstc.Telefone) & ","
    strTMP = strTMP & " Status = " & M46V999.gfFormatarCamposSQL(pobjstc.STATUS) & ","
    strTMP = strTMP & " CCotacaoOrig = " & M46V999.gfFormatarCamposSQL(pobjstc.CCotacaoOrig) & ","
    If M46V999.gfPreenchido(pPropTabela) Then
        strTMP = strTMP & " Usuario = " & M46V999.gfFormatarCamposSQL(pobjstc.Usuario) & ","
    Else
        strTMP = strTMP & " Usuario = " & M46V999.gfFormatarCamposSQL(val(pobjstc.Usuario)) & ","
    End If
    strTMP = strTMP & " Senha = " & M46V999.gfFormatarCamposSQL(pobjstc.Senha)



    strTMP = strTMP & " WHERE "
    strTMP = strTMP & " CCotacao = " & M46V999.gfFormatarCamposSQL(pobjstc.CCotacao)
    strTMP = strTMP & " AND Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)

    mfFormatarAlteracaoSQL = strTMP
End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA28V721B, _
                                       ByVal pPropTabela As String) As String

    Dim strTMP As String

    strTMP = "DELETE FROM " & pPropTabela & "P0045400 WHERE "
    strTMP = strTMP & " CCotacao = " & M46V999.gfFormatarCamposSQL(pobjstc.CCotacao)
    strTMP = strTMP & " AND Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    strTMP = strTMP & " AND status = " & M46V999.gfFormatarCamposSQL(pobjstc.STATUS)

    mfFormatarExclusaoSQL = strTMP
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA28V721B, _
                                       ByVal pPropTabela As String) As String
    Dim strTMP As String

    strTMP = "SELECT * FROM " & pPropTabela & "P0045400 WITH (NOLOCK)"
    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        strTMP = strTMP & " WHERE Nosso_Numero=" & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
        If pobjstc.STATUS = -1 Then
            strTMP = strTMP & " and status = -1"
        Else
            strTMP = strTMP & " and status <> -1"
        End If
    End If
    mfFormatarConsultaSQL = strTMP
End Function
Private Sub mpLimparColecao()

    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Private Sub mpCarregarAtributos(ByRef pobjstc As stcA28V721B, prst As clsYasRecordSet)

    pobjstc.CCotacao = M46V999.gfTratarNulo(prst("CCotacao"), tpoString)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Dat_Cotacao = M46V999.gfTratarNulo(prst("Dat_Cotacao"), tpoNumerico)
    pobjstc.Email = M46V999.gfTratarNulo(prst("Email"), tpoString)
    pobjstc.Nome = M46V999.gfTratarNulo(prst("Nome"), tpoString)
    pobjstc.Telefone = M46V999.gfTratarNulo(prst("Telefone"), tpoString)
    pobjstc.STATUS = M46V999.gfTratarNulo(prst("Status"), tpoNumerico)
    pobjstc.Senha = M46V999.gfTratarNulo(prst("Senha"), tpoNumerico)
    pobjstc.Usuario = M46V999.gfTratarNulo(prst("Usuario"), tpoNumerico)
    pobjstc.CCotacaoOrig = M46V999.gfTratarNulo(prst("CCotacaoOrig"), tpoString)
    pobjstc.MaxDesc = M46V999.gfTratarNulo(prst("MaxDesc"), tpoNumerico)


End Sub
Public Function gfIncluir(ByRef pobjstc As stcA28V721B, _
                          ByVal pPropTabela As String) As Boolean

    Dim sSql As String

    On Local Error GoTo ControleErros
    Dim rst As clsYasRecordSet

    If CnDB Is Nothing Then GoTo ControleErros

    If M46V999.gfExeSQL(CnDB, mfFormatarInclusaoSQL(pobjstc, pPropTabela), UltimoErro) <> 0 Then
        If UCase(UltimoErro) Like "*DUPLIC*" Then
            gfIncluir = gfAlterar(pobjstc, pPropTabela)
            Exit Function
        Else
            GoTo ControleErros
        End If
    End If
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    gfIncluir = True
    Exit Function
ControleErros:
    gfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
    End If
End Function
Public Function gfAlterar(ByRef pobjstc As stcA28V721B, _
                          ByVal pPropTabela As String) As Boolean

    Dim sSql As String

    On Error GoTo ControleErros
    If CnDB Is Nothing Then GoTo ControleErros

    If M46V999.gfExeSQL(CnDB, mfFormatarAlteracaoSQL(pobjstc, pPropTabela), UltimoErro) <> 0 Then
        GoTo ControleErros
    End If
    If Colecao.Count > 0 Then
        Colecao.Remove (pobjstc.Class_ID)
    End If
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    gfAlterar = True
    Exit Function
ControleErros:
    gfAlterar = False
    UltimoErro = Err.Description
End Function
Public Function gfExcluir(ByRef pobjstc As stcA28V721B, _
                          ByVal pPropTabela As String) As Boolean

    Dim sSql As String

    On Error GoTo ControleErros
    If CnDB Is Nothing Then GoTo ControleErros

    If M46V999.gfExeSQL(CnDB, mfFormatarExclusaoSQL(pobjstc, pPropTabela), UltimoErro) <> 0 Then
        GoTo ControleErros
    End If

    gfExcluir = True
    Exit Function
ControleErros:
    gfExcluir = False
    UltimoErro = Err.Description
End Function
Public Function gfCarregarColecao(ByRef pobjstc As stcA28V721B, _
                                  ByVal pPropTabela As String) As Boolean

    Dim objstc As stcA28V721B
    Dim rst As clsYasRecordSet

    On Error GoTo ControleErros
    If CnDB Is Nothing Then GoTo ControleErros

    If Not M46V999.gfObtRegistro(CnDB, mfFormatarConsultaSQL(pobjstc, pPropTabela), rst, UltimoErro) Then
        GoTo ControleErros
    End If
    Call mpLimparColecao
    Do While Not rst.EOF
        Set objstc = New stcA28V721B
        Call mpCarregarAtributos(objstc, rst)
        Colecao.Add Item:=objstc, Key:=objstc.Class_ID
        rst.MoveNext
    Loop

    gfCarregarColecao = True
    Exit Function
ControleErros:
    gfCarregarColecao = False
    If Not M46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If
End Function
Private Sub Class_Initialize()
    Set Colecao = New Collection
    Set mClsM46V999 = Nothing
    Set mCnDB = Nothing
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
        If M46V999.gfAbrBasDados("P0045400", M46V999.gTipP0045400, M46V999.gEndP0045400, mCnDB, _
                                 UltimoErro) = False Then Set mCnDB = Nothing
    End If

Retorno:
    Set CnDB = mCnDB
End Property

