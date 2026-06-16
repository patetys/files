VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V737A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V737A --> Tabela Tab_Ped_Loc_Adic
' Data/Hora: 03/01/2014 - 14:00:00
' Sistema: Multi
' Marcio Oliveira da Silva
' ======================================================================================================================================================

Public ConexaoBD   As ADODB.Connection  ' Conexão com o Banco de Dados
Public Colecao     As Collection        ' Coleção de objetos do tipo stcA46V737B
Public UltimoErro  As String            ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V737A"
End Property

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V737B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
Dim lInsert As String                'Query.
 
    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc_Adic ( "
    lInsert = lInsert & " Nosso_Numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " Num_item,"
    lInsert = lInsert & " Num_Seq,"
    lInsert = lInsert & " Cod_Tip_Logr_Risco,"
    lInsert = lInsert & " Nom_Logr_Risco,"
    lInsert = lInsert & " Dsc_Compl_Risco,"
    lInsert = lInsert & " Nom_Bairro_Risco,"
    lInsert = lInsert & " Nom_Cid_Risco,"
    lInsert = lInsert & " Sig_UF_Risco,"
    lInsert = lInsert & " Num_CEP_Risco"
  
    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_SEQ) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Logr_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Logr_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Bairro_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Cid_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Sig_UF_Risco) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cep_Risco) & ")"
    
    mfFormatarInclusaoSQL = lInsert
End Function


Private Function mfFormatarAlteracaoSQL(ByRef pobjstc As stcA46V737B, _
                                        ByVal pPropTabela As String, _
                                        ByVal pNomeTabImport As String) As String
Dim lUpdate As String                'Query.

    lUpdate = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc_Adic Set "
    lUpdate = lUpdate & " Num_item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lUpdate = lUpdate & " Num_Seq = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_SEQ) & ","
    lUpdate = lUpdate & " Cod_Tip_Logr_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Logr_Risco) & ","
    lUpdate = lUpdate & " Nom_Logr_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Logr_Risco) & ","
    lUpdate = lUpdate & " Dsc_Compl_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Risco) & ","
    lUpdate = lUpdate & " Nom_Bairro_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Bairro_Risco) & ","
    lUpdate = lUpdate & " Nom_Cid_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Cid_Risco) & ","
    lUpdate = lUpdate & " Sig_UF_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Sig_UF_Risco) & ","
    lUpdate = lUpdate & " Num_CEP_Risco = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cep_Risco)
    
    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    lUpdate = lUpdate & " and Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    lUpdate = lUpdate & " and Num_Seq = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_SEQ)

    mfFormatarAlteracaoSQL = lUpdate
End Function

Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V737B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
Dim lDelete As String                'Query.
 
    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc_Adic WHERE "
    lDelete = lDelete & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Item > 0 Then
        lDelete = lDelete & " AND Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If
    If pobjstc.NUM_SEQ > 0 Then
        lDelete = lDelete & " AND Num_Seq = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_SEQ)
    End If
    mfFormatarExclusaoSQL = lDelete
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V737B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
Dim lSelect As String                'Query.
 
    lSelect = "SELECT * FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Loc_Adic"
    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lSelect = lSelect & " Where Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If
    If pobjstc.Num_Item <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Num_Item= " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.NUM_SEQ <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Num_Seq= " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_SEQ)
    End If
    lSelect = lSelect & " order by num_item, Num_Seq"
    mfFormatarConsultaSQL = lSelect
End Function
Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V737B, prst As clsYasRecordSet)

    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_item"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.NUM_SEQ = M46V999.gfTratarNulo(prst("Num_Seq"), tpoNumerico)
    pobjstc.Cod_Tip_Logr_Risco = M46V999.gfTratarNulo(prst("Cod_Tip_Logr_Risco"), tpoString)
    pobjstc.Nom_Logr_Risco = M46V999.gfTratarNulo(prst("Nom_Logr_Risco"), tpoString)
    pobjstc.Dsc_Compl_Risco = M46V999.gfTratarNulo(prst("Dsc_Compl_Risco"), tpoString)
    pobjstc.Nom_Bairro_Risco = M46V999.gfTratarNulo(prst("Nom_Bairro_Risco"), tpoString)
    pobjstc.Nom_Cid_Risco = M46V999.gfTratarNulo(prst("Nom_Cid_Risco"), tpoString)
    pobjstc.Sig_UF_Risco = M46V999.gfTratarNulo(prst("Sig_UF_Risco"), tpoString)
    pobjstc.Num_Cep_Risco = M46V999.gfTratarNulo(prst("Num_CEP_Risco"), tpoNumerico)
    
End Sub

Public Function gfIncluir(ByRef pobjstc As stcA46V737B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean
    Dim lsSQl       As String                'Query.
    Dim lMensagem   As String
    Dim lAlter      As String
    Dim lRetornar   As Byte
    
 
    On Local Error GoTo ControleErro
    
    If Not mfConectar Then GoTo ControleErro
Retornar:
    lRetornar = 0
    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
    
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    gfIncluir = True
    Exit Function
    
ControleErro:
    gfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        If Err.Number = 3022 Then   'Já existe.
            UltimoErro = "FTE0027 - Item já existente (nosso número " & pobjstc.nosso_numero & _
                        " - item " & pobjstc.Num_Item & ")."
            Call M46V999.gpGraLog(0, "A46V737A.gfIncluir ERRO - " & Err & " " & Err.Description & _
            vbCrLf & " SQL: " & lsSQl)
        Else
            UltimoErro = Err.Description
            Call M46V999.gpGraLog(0, "A46V737A.gfIncluir ERRO - " & UltimoErro & _
            vbCrLf & " SQL: " & lsSQl)
        End If
    End If
    
End Function

Public Function gfAlterar(ByRef pobjstc As stcA46V737B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                          Optional ByVal pRemoverColecao As Boolean = True) As Boolean
Dim lUpdate      As String                'Query.
 
    On Error GoTo ControleErro
    
    If Not mfConectar Then GoTo ControleErro
    
    lUpdate = mfFormatarAlteracaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    
    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErro
        
    If Colecao.Count > 0 And pRemoverColecao = True Then
        Colecao.Remove (pobjstc.Class_ID)
    End If
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    gfAlterar = True
    Exit Function
ControleErro:
    gfAlterar = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V737A.gfAlterar ERRO - " & Err & " " & Err.Description & _
        vbCrLf & " SQL: " & lUpdate)
    Else
        Call M46V999.gpGraLog(0, "A46V737A.gfAlterar ERRO - " & UltimoErro & _
        vbCrLf & " SQL: " & lUpdate)
    End If
End Function

Public Function gfExcluir(ByRef pobjstc As stcA46V737B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean
    Dim lsSQl      As String                'Query.

    On Error GoTo ControleErro
    
    If Not mfConectar Then GoTo ControleErro
    
    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
    If Colecao.Count > 0 Then
        Colecao.Remove (pobjstc.Class_ID)
    End If
    Set pobjstc = Nothing
    gfExcluir = True
    Exit Function
ControleErro:
    gfExcluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V737A.gfExcluir ERRO - " & Err & " " & Err.Description & _
        vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V737A.gfExcluir ERRO - " & UltimoErro & _
        vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V737B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc     As stcA46V737B              'Objeto da classe de estrutura stcA46V737B.
    Dim lrst        As clsYasRecordSet          'Registro.
    Dim lsSQl       As String                   'SQL.
    
    On Error GoTo ControleErro
    
    If Not mfConectar Then GoTo ControleErro
    
    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
        lrst, UltimoErro) = False Then GoTo ControleErro
        
    Call mpLimparColecao
    Do While Not lrst.EOF
        
        Set lobjstc = New stcA46V737B
        
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
        Call M46V999.gpGraLog(0, "A46V737A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
        vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V737A.mfCarregarColecao ERRO - " & UltimoErro & _
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
