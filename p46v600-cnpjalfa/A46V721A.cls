VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V721A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: clsA46V721A --> Tabela Tab_Ped_Espec2
' Data/Hora: 15/09/2005 - 09:52:48
' Sistema: Multi
' Simony Sales
' ======================================================================================================================================================


' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V721B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V721A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V721B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lInsert As String                'Query.

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Espec2 ( "
    lInsert = lInsert & " Nosso_numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " Num_Ped,"
    lInsert = lInsert & " Cod_Org,"
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " Num_Apol,"
    lInsert = lInsert & " Num_Endo,"
    lInsert = lInsert & " Dsc_LocalExposicao,"
    lInsert = lInsert & " Dsc_Objetos,"
    lInsert = lInsert & " Dsc_RCobertos,"
    lInsert = lInsert & " Dsc_Caracteristica,"
    lInsert = lInsert & " Dsc_Perimetro,"
    lInsert = lInsert & " Dsc_Especificacao,"
    lInsert = lInsert & " Dsc_Coberturas,"
    lInsert = lInsert & " Dsc_Franquia,"
    lInsert = lInsert & " Nom_Contratante,"
    lInsert = lInsert & " Dsc_Clausulas,"
    lInsert = lInsert & " Dsc_Condicao,"
    lInsert = lInsert & " Dsc_OutrasInf,"
    lInsert = lInsert & " Dsc_Observacao,"
    lInsert = lInsert & " Dat_Obra,"
    lInsert = lInsert & " Dat_Manut,"
    lInsert = lInsert & " Num_PI,"
    lInsert = lInsert & " Ind_Impressao"
    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NOSSO_NUMERO) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_LocalExposicao) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Objetos) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_RCobertos) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Caracteristica) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Perimetro) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Especificacao) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Coberturas) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Franquia) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Contratante) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Clausulas) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Condicao) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_OutrasInf) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Observacao) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Obra) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Manut) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Impressao) & ")"


    mfFormatarInclusaoSQL = lInsert
End Function


Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V721B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String

    Dim lDelete As String                'Query.

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Espec2 WHERE "
    lDelete = lDelete & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.NOSSO_NUMERO)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarExclusaoSQL = lDelete
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V721B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String                'Query.
    lSelect = "SELECT * FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Espec2  with(nolock)"
    If M46V999.gfPreenchido(pobjstc.NOSSO_NUMERO) Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.NOSSO_NUMERO)
    End If
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    mfFormatarConsultaSQL = lSelect
End Function

Private Function mfFormatarConsultaImpressaoSQL(ByVal pPropTabela As String, _
                                                ByVal pNomeTabImport As String) As String
    Dim lstrTMP As String                'Query.

    lstrTMP = " SELECT esp.* FROM "
    lstrTMP = lstrTMP & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Espec2 as esp inner join "
    lstrTMP = lstrTMP & " " & pPropTabela & pNomeTabImport & "tab_ctrl_emis as ctrl on "
    lstrTMP = lstrTMP & " esp.Num_PI = ctrl.Num_PI "
    lstrTMP = lstrTMP & " inner join siscota.siscota.tab_cotacao as cota "
    lstrTMP = lstrTMP & " on ccotacao = ctrl.num_cotac "
    lstrTMP = lstrTMP & " and especificacao = 1 "
    lstrTMP = lstrTMP & " Where esp.Num_Apol <> 0 "
    lstrTMP = lstrTMP & " and Ind_Impressao <> 1 "
    lstrTMP = lstrTMP & " and cod_iden_stat_acao = 90 "

    mfFormatarConsultaImpressaoSQL = lstrTMP
End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V721B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)

    pobjstc.NOSSO_NUMERO = M46V999.gfTratarNulo(prst("NOsso_Numero"), tpoString)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Dsc_Objetos = M46V999.gfTratarNulo(prst("Dsc_Objetos"), tpoString)
    pobjstc.Dsc_LocalExposicao = M46V999.gfTratarNulo(prst("Dsc_LocalExposicao"), tpoString)
    pobjstc.Dsc_RCobertos = M46V999.gfTratarNulo(prst("Dsc_RCobertos"), tpoString)
    pobjstc.Dsc_Caracteristica = M46V999.gfTratarNulo(prst("Dsc_Caracteristica"), tpoString)
    pobjstc.Dsc_Perimetro = M46V999.gfTratarNulo(prst("Dsc_Perimetro"), tpoString)
    pobjstc.Dsc_Especificacao = M46V999.gfTratarNulo(prst("Dsc_Especificacao"), tpoString)
    pobjstc.Dsc_Coberturas = M46V999.gfTratarNulo(prst("Dsc_Coberturas"), tpoString)
    pobjstc.Dsc_Franquia = M46V999.gfTratarNulo(prst("Dsc_Franquia"), tpoString)
    pobjstc.Nom_Contratante = M46V999.gfTratarNulo(prst("Nom_Contratante"), tpoString)
    pobjstc.Dsc_Clausulas = M46V999.gfTratarNulo(prst("Dsc_Clausulas"), tpoString)
    pobjstc.Dsc_Condicao = M46V999.gfTratarNulo(prst("Dsc_Condicao"), tpoString)
    pobjstc.Dsc_OutrasInf = M46V999.gfTratarNulo(prst("Dsc_OutrasInf"), tpoString)
    pobjstc.Dsc_Observacao = M46V999.gfTratarNulo(prst("Dsc_Observacao"), tpoString)
    pobjstc.Dat_Obra = M46V999.gfTratarNulo(prst("Dat_Obra"), tpoNumerico)
    pobjstc.Dat_Manut = M46V999.gfTratarNulo(prst("Dat_Manut"), tpoNumerico)
    pobjstc.Num_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
    On Error Resume Next
    pobjstc.Ind_Impressao = M46V999.gfTratarNulo(prst("Ind_Impressao"), tpoNumerico)
    If Err.Number = 3265 Then
        Dim lAlter As String
        lAlter = " alter table Tab_Ped_Espec2 add Ind_Impressao int "
        If M46V999.gfExeSQL(ConexaoBD, lAlter, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(2, UltimoErro)
            Exit Sub
        End If
    End If
End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V721B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean

    Dim lSelect As String           'Query.
    Dim lbdTab_Espec As ADODB.Connection    'Arquivo: P00Tarifa.
    Dim lMensagem As String     'Mensagem.
    Dim lrsTab_Espec As ADODB.Recordset    'Registro : Tbm_Espec_Padrao.

    On Local Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    '******************************************
    'Em razão dos campos RichText será utilizado outro métode de inclusão
    'para permitir os caracteres especiais
    '******************************************
    'lSelect = mfFormatarInclusaoSQL(pObjstc)
    'If M46V999.gfExeSQL(ConexaoBD, lSelect, UltimoErro) <> 0 Then GoTo ControleErro


    lSelect = " Select *  from " & pPropTabela & pNomeTabImport & "Tab_Ped_Espec2 "
    lSelect = lSelect & " Where Nosso_Numero  = " & M46V999.gfFormatarCamposSQL(pobjstc.NOSSO_NUMERO)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    Set lrsTab_Espec = New ADODB.Recordset    'Este recordSet usa uma funcao nao implementada no YasRecordSet
    lrsTab_Espec.CursorType = adOpenKeyset
    lrsTab_Espec.LockType = adLockOptimistic

    lrsTab_Espec.Open lSelect, ConexaoBD

    lrsTab_Espec.AddNew

    lrsTab_Espec("Nosso_numero") = pobjstc.NOSSO_NUMERO
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lrsTab_Espec("Cod_Unicidade") = pobjstc.Cod_Unicidade
    End If
    lrsTab_Espec("Num_Ped") = pobjstc.Num_Ped
    lrsTab_Espec("Cod_Org") = pobjstc.Cod_Org
    lrsTab_Espec("Cod_Ramo") = pobjstc.Cod_Ramo
    lrsTab_Espec("Num_Apol") = pobjstc.Num_Apol
    lrsTab_Espec("Num_Endo") = pobjstc.Num_Endo
    lrsTab_Espec("Dsc_LocalExposicao") = M46V999.gfTratarVazio(pobjstc.Dsc_LocalExposicao)
    lrsTab_Espec("Dsc_Objetos") = M46V999.gfTratarVazio(pobjstc.Dsc_Objetos)
    lrsTab_Espec("Dsc_RCobertos") = M46V999.gfTratarVazio(pobjstc.Dsc_RCobertos)
    lrsTab_Espec("Dsc_Caracteristica") = M46V999.gfTratarVazio(pobjstc.Dsc_Caracteristica)
    lrsTab_Espec("Dsc_Perimetro") = M46V999.gfTratarVazio(pobjstc.Dsc_Perimetro)
    lrsTab_Espec("Dsc_Especificacao") = M46V999.gfTratarVazio(pobjstc.Dsc_Especificacao)
    lrsTab_Espec("Dsc_Coberturas") = M46V999.gfTratarVazio(pobjstc.Dsc_Coberturas)
    lrsTab_Espec("Dsc_Franquia") = M46V999.gfTratarVazio(pobjstc.Dsc_Franquia)
    lrsTab_Espec("Nom_Contratante") = M46V999.gfTratarVazio(pobjstc.Nom_Contratante)
    lrsTab_Espec("Dsc_Clausulas") = M46V999.gfTratarVazio(pobjstc.Dsc_Clausulas)
    lrsTab_Espec("Dsc_Condicao") = M46V999.gfTratarVazio(pobjstc.Dsc_Condicao)
    lrsTab_Espec("Dsc_OutrasInf") = M46V999.gfTratarVazio(pobjstc.Dsc_OutrasInf)
    lrsTab_Espec("Dsc_Observacao") = M46V999.gfTratarVazio(pobjstc.Dsc_Observacao)
    lrsTab_Espec("Dat_Obra") = pobjstc.Dat_Obra
    lrsTab_Espec("Dat_Manut") = pobjstc.Dat_Manut
    lrsTab_Espec("Num_PI") = pobjstc.Num_PI
    lrsTab_Espec("Ind_Impressao") = 0

    lrsTab_Espec.Update

    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfIncluir = True
    Exit Function

ControleErro:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V721A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lSelect)
    Else
        Call M46V999.gpGraLog(0, "A46V721A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lSelect)
    End If
End Function
Public Function mfExcluir(ByRef pobjstc As stcA46V721B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, tagDynaTrace As String) As Boolean

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfExcluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lsSQl As String           'Query.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
    If Colecao.Count > 0 Then
        Colecao.Remove (pobjstc.Class_ID)
    End If
    mfExcluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErro:
    mfExcluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V721A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V721A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V721B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V721B          'Objeto da classe de estrutura stcA46V721B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    If M46V999.gfObtRegistro(ConexaoBD, mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport), _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V721B
        Call mpCarregarAtributos(lobjstc, lrst)
        Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID
        lrst.MoveNext
    Loop
    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecao = True
    Exit Function
ControleErro:
    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V721A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V721A.mfCarregarColecao ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfCarregarColecaoImpressao(ByVal pPropTabela As String, _
                                           ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V721B          'Objeto da classe de estrutura stcA46V721B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    If M46V999.gfObtRegistro(ConexaoBD, mfFormatarConsultaImpressaoSQL(pPropTabela, pNomeTabImport), _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V721B
        Call mpCarregarAtributos(lobjstc, lrst)
        Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID
        lrst.MoveNext
    Loop

    mfCarregarColecaoImpressao = True
    Exit Function
ControleErro:
    mfCarregarColecaoImpressao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Number & " " & Err.Description
    End If

    Call M46V999.gpGraLog(0, "A46V721A.mfCarregarColecaoImpressao ERRO - " & UltimoErro & _
                             vbCrLf & " SQL: " & lsSQl)
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
