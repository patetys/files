VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V732A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: clsA46V732A --> Tabela Tab_Ped_LMI
' Data/Hora: 13/11/2012 - 11:52:48
' Sistema: Multi
' Thais Juren Gomes

' ======================================================================================================================================================


' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V732B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V732A"
End Property


Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V732B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lInsert As String                'Query.


    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_LMI ( "
    lInsert = lInsert & " Nosso_numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " COD_ORG, "
    lInsert = lInsert & " COD_RAMO, "
    lInsert = lInsert & " NUM_PED, "
    lInsert = lInsert & " NUM_APOL, "
    lInsert = lInsert & " NUM_ENDO, "
    lInsert = lInsert & " TAB_COBERT, "
    lInsert = lInsert & " COD_COBERT, "
    lInsert = lInsert & " VAL_IS, "
    lInsert = lInsert & " PER_FRANQ, "
    lInsert = lInsert & " VAL_FRANQ_MIN, "
    lInsert = lInsert & " VAL_FRANQ_MAX, "
    lInsert = lInsert & " NUM_PI, "
    lInsert = lInsert & " NUM_MESES_PI"
    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tab_Cobert) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Franq) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Franq_Min) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Franq_Max) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Meses_PI) & ")"

    mfFormatarInclusaoSQL = lInsert

End Function


Private Function mfFormatarExclusaoSQL(pobjstc As stcA46V732B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lDelete As String                'Query.


    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_LMI WHERE "
    lDelete = lDelete & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If

    mfFormatarExclusaoSQL = lDelete

End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V732B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String                'Query.

    lSelect = "SELECT * from " & pPropTabela & pNomeTabImport & "Tab_Ped_LMI  with(nolock) "
    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Apol <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol)
        lSelect = M46V999.gfWhereAnd(lSelect) & " num_endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo)
    End If


    mfFormatarConsultaSQL = lSelect
End Function
Private Sub mpLimparColecao()

    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V732B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)

    pobjstc.Cod_Cobert = M46V999.gfTratarNulo(prst("Cod_Cobert"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("NOsso_Numero"), tpoString)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Tab_Cobert = M46V999.gfTratarNulo(prst("Tab_Cobert"), tpoNumerico)
    pobjstc.Val_IS = M46V999.gfTratarNulo(prst("Val_IS"), tpoNumerico)
    pobjstc.Per_Franq = M46V999.gfTratarNulo(prst("Per_Franq"), tpoNumerico)
    pobjstc.Val_Franq_Min = M46V999.gfTratarNulo(prst("Val_Franq_Min"), tpoNumerico)
    pobjstc.Val_Franq_Max = M46V999.gfTratarNulo(prst("Val_Franq_Max"), tpoNumerico)
    pobjstc.Num_Meses_PI = M46V999.gfTratarNulo(prst("Num_Meses_PI"), tpoNumerico)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)

End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V732B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean

    Dim lsSQl As String           'Query.

    On Local Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro


    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    mfIncluir = True
    Exit Function
ControleErro:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V732A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V732A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function
Public Function mfExcluir(ByRef pobjstc As stcA46V732B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean

    Dim lsSQl As String           'Query.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro


    Set Colecao = New Collection

    Set pobjstc = Nothing
    mfExcluir = True
    Exit Function
ControleErro:
    mfExcluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V732A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V732A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function
Public Function mfCarregarColecao(ByRef pobjstc As stcA46V732B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V732B          'Objeto da classe de estrutura stcA46V732B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    If M46V999.gfObtRegistro(ConexaoBD, mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport), _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V732B
        Set lobjstc.M46V999 = M46V999

        Call mpCarregarAtributos(lobjstc, lrst)
        Set lobjstc.ObjPai = pobjstc.ObjPai
        Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID
        lrst.MoveNext
    Loop
    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecao = True

SAIDA:
    GoSub FechaObj
    Exit Function

ControleErro:
    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V732A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V732A.mfCarregarColecao ERRO - " & UltimoErro & _
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

Private Function mfRetorna_ValIS(pobjStcCobert As stcA46V704B) As Double
    'Criado para somar o valor da cobertura obrigatoria em caso de verbas separadas
    Dim lobjstcPedLoc As stcA46V708B
    Dim lobjstcPedCobert  As stcA46V704B
   
    Set lobjstcPedLoc = pobjStcCobert.ObjPai
    
    mfRetorna_ValIS = pobjStcCobert.Val_IS
    If lobjstcPedLoc.Ind_VerbaSep And pobjStcCobert.Cod_Grupo_Cobertura = 1 Then
        For Each lobjstcPedCobert In lobjstcPedLoc.ColecaoPedCobert
            If lobjstcPedCobert.Cod_Grupo_Cobertura = 5 Then
                mfRetorna_ValIS = mfRetorna_ValIS + lobjstcPedCobert.Val_IS
                Exit For
            End If
        Next
    End If

    Set lobjstcPedLoc = Nothing
    Set lobjstcPedCobert = Nothing

End Function

Public Function mfGerarDadosLMI(lobjstcCobertPed As stcA46V702B, ByVal pPropTabela As String, _
                                ByVal pNomeTabImport As String) As Boolean

    Dim lobjstcPedLoc As stcA46V708B
    Dim lAchou As Boolean
    Dim lobjstcPedCobert As stcA46V704B
    Dim lI As Long
    Dim lQtdCobert As Long
    Dim vCobert() As String
    Dim vArrayItens() As String
    Dim vArrayCoberts() As String
    Dim lobjstcPedLMI As stcA46V732B
    Dim lTabCobert As Long
    Dim lSelect As String
    Dim lNumItem As String
    Dim RstAux As clsYasRecordSet
    Dim lColecaoCobertura As Collection
    Dim lobjstcCobert As stcA46V704B
    Dim lEndosso As Boolean
    
    lEndosso = False
    lNumItem = ""
    For Each lobjstcPedLoc In lobjstcCobertPed.ColecaoPedLoc

        If lobjstcPedLoc.Cod_DizerItem <> 0 Then lEndosso = True

        lNumItem = lNumItem & lobjstcPedLoc.Num_Item & ","

        For Each lobjstcPedCobert In lobjstcPedLoc.ColecaoPedCobert
            If lobjstcPedCobert.Cod_Grupo_Cobertura <> 5 Then
            lAchou = False
            lTabCobert = lobjstcPedCobert.Tab_Cobert
            For lI = 1 To lQtdCobert
                If vCobert(0, lI) = lobjstcPedCobert.Cod_Cobert Then
                    lAchou = True
                    Exit For
                End If
            Next
            If Not lAchou Then
                lQtdCobert = lQtdCobert + 1
                ReDim Preserve vCobert(5, lQtdCobert)
                vCobert(0, lQtdCobert) = lobjstcPedCobert.Cod_Cobert
                'vCobert(1, lQtdCobert) = lobjstcPedCobert.gNomeCobertura
                vCobert(1, lQtdCobert) = lobjstcPedCobert.gNomeFranquia
                vCobert(2, lQtdCobert) = mfRetorna_ValIS(lobjstcPedCobert)
                vCobert(3, lQtdCobert) = lobjstcPedCobert.Val_Franq_Min
                vCobert(4, lQtdCobert) = lobjstcPedCobert.Num_Meses_PI
                vCobert(5, lQtdCobert) = lobjstcPedCobert.Per_Franq
            Else
                If vCobert(2, lI) < lobjstcPedCobert.Val_IS Then
                    vCobert(2, lI) = lobjstcPedCobert.Val_IS
                End If
                If vCobert(3, lI) < lobjstcPedCobert.Val_Franq_Min Then
                    vCobert(3, lI) = lobjstcPedCobert.Val_Franq_Min
                End If
                If vCobert(4, lI) < lobjstcPedCobert.Num_Meses_PI Then
                    vCobert(4, lI) = lobjstcPedCobert.Num_Meses_PI
                End If
                If vCobert(5, lI) < lobjstcPedCobert.Per_Franq Then
                    vCobert(5, lI) = lobjstcPedCobert.Per_Franq
                End If
            End If
            End If
        Next
    Next

    If lEndosso Then
        lSelect = "SELECT c.* , Num_Meses_Pi, Val_Risco_Cob "
        lSelect = lSelect & " FROM TAB_APOL_COBERT as c "    ' & "_Atual"
        lSelect = lSelect & " left join TAB_APOL_COBERT_COMPL as cl"
        lSelect = lSelect & " on c.num_pi = cl.num_pi"
        lSelect = lSelect & " and c.cod_cobert = cl.cod_cobert"
        lSelect = lSelect & " and c.num_item = cl.num_item"
        lSelect = lSelect & " Where c.NUM_APOL = " & lobjstcCobertPed.Num_Apol_An
        lSelect = lSelect & " and c.num_item not in ( " & Left(lNumItem, Len(lNumItem) - 1) & ")"
        lSelect = lSelect & " order by c.num_item , c.cod_cobert"


        If Not mfConectar Then GoTo ControleErro
        If M46V999.gfObtRegistro(ConexaoBD, lSelect, RstAux, UltimoErro) = False Then
            Call M46V999.gpFecha2(ConexaoBD)
            GoTo ControleErro
        End If
        Set lColecaoCobertura = New Collection
        Do While Not RstAux.EOF
            Set lobjstcCobert = New stcA46V704B
            lobjstcCobert.Cod_Cobert = M46V999.gfTratarNulo(RstAux("Cod_Cobert"), tpoNumerico)
            lobjstcCobert.Cod_Ramo = M46V999.gfTratarNulo(RstAux("Cod_Ramo"), tpoNumerico)
            lobjstcCobert.Num_Item = M46V999.gfTratarNulo(RstAux("Num_item"), tpoNumerico)
            lobjstcCobert.Tab_Cobert = M46V999.gfTratarNulo(RstAux("Tab_Cobert"), tpoNumerico)
            lobjstcCobert.Val_IS = M46V999.gfTratarNulo(RstAux("Val_IS"), tpoNumerico)
            lobjstcCobert.Per_Taxa = M46V999.gfTratarNulo(RstAux("Per_Taxa"), tpoNumerico)
            lobjstcCobert.Per_Franq = M46V999.gfTratarNulo(RstAux("Per_Franq"), tpoNumerico)
            lobjstcCobert.Val_Franq_Min = M46V999.gfTratarNulo(RstAux("Val_Franq_Min"), tpoNumerico)
            lobjstcCobert.Val_Franq_Max = M46V999.gfTratarNulo(RstAux("Val_Franq_Max"), tpoNumerico)
            lobjstcCobert.Num_Meses_PI = M46V999.gfTratarNulo(RstAux("Num_Meses_PI"), tpoNumerico)
            lobjstcCobert.Val_Risco_Cob = M46V999.gfTratarNulo(RstAux("Val_Risco_Cob"), TpoValor)
            lobjstcCobert.NUM_PI = M46V999.gfTratarNulo(RstAux("Num_PI"), tpoNumerico)
            lColecaoCobertura.Add lobjstcCobert, lobjstcCobert.Class_ID
            RstAux.MoveNext
        Loop


        For Each lobjstcPedCobert In lColecaoCobertura
            If lobjstcPedCobert.Cod_Grupo_Cobertura <> 5 Then
            lAchou = False
            lTabCobert = lobjstcPedCobert.Tab_Cobert
            For lI = 1 To lQtdCobert
                If vCobert(0, lI) = lobjstcPedCobert.Cod_Cobert Then
                    lAchou = True
                    Exit For
                End If
            Next
            
            Set lobjstcPedCobert.M46V999 = Me.M46V999
            
            If Not lAchou Then
                lQtdCobert = lQtdCobert + 1
                ReDim Preserve vCobert(5, lQtdCobert)
                vCobert(0, lQtdCobert) = lobjstcPedCobert.Cod_Cobert
                
                'Linha abaixo comentada pq o ObjPai não está carregado e o NomeCobertura não é usado
                'vCobert(1, lQtdCobert) = lobjstcPedCobert.gNomeCobertura
                
                'vCobert(1, lQtdCobert) = lobjstcPedCobert.gNomeFranquia
                vCobert(2, lQtdCobert) = lobjstcPedCobert.Val_IS
                'vCobert(2, lQtdCobert) = mfRetorna_ValIS(lobjstcPedCobert)
                vCobert(3, lQtdCobert) = lobjstcPedCobert.Val_Franq_Min
                vCobert(4, lQtdCobert) = lobjstcPedCobert.Num_Meses_PI
                vCobert(5, lQtdCobert) = lobjstcPedCobert.Per_Franq
            Else
                If vCobert(2, lI) < lobjstcPedCobert.Val_IS Then
                    vCobert(2, lI) = lobjstcPedCobert.Val_IS
                End If
                If vCobert(3, lI) < lobjstcPedCobert.Val_Franq_Min Then
                    vCobert(3, lI) = lobjstcPedCobert.Val_Franq_Min
                End If
                If vCobert(4, lI) < lobjstcPedCobert.Num_Meses_PI Then
                    vCobert(4, lI) = lobjstcPedCobert.Num_Meses_PI
                End If
                If vCobert(5, lI) < lobjstcPedCobert.Per_Franq Then
                    vCobert(5, lI) = lobjstcPedCobert.Per_Franq
                End If
            End If
                End If
        Next
    End If

    Set lobjstcPedLMI = New stcA46V732B
    lobjstcPedLMI.nosso_numero = lobjstcCobertPed.nosso_numero
    If Not mfExcluir(lobjstcPedLMI, pPropTabela, pNomeTabImport) Then
        Exit Function
    End If
    Set lobjstcCobertPed.ColecaoLMI = New Collection
    For lI = 1 To lQtdCobert
        Set lobjstcPedLMI = New stcA46V732B
        With lobjstcPedLMI
            .Cod_Cobert = vCobert(0, lI)
            .Cod_Ramo = lobjstcCobertPed.Cod_Ramo
            .nosso_numero = lobjstcCobertPed.nosso_numero
            .Num_Meses_PI = vCobert(4, lI)
            .Num_Ped = lobjstcCobertPed.Num_Ped
            .NUM_PI = lobjstcCobertPed.NUM_PI
            .Per_Franq = vCobert(5, lI)
            .Tab_Cobert = lTabCobert
            .Val_Franq_Max = 0
            .Val_Franq_Min = vCobert(3, lI)
            .Val_IS = vCobert(2, lI)
        End With

        If Not mfIncluir(lobjstcPedLMI, pPropTabela, pNomeTabImport) Then
            Exit Function
        End If

        lobjstcCobertPed.ColecaoLMI.Add lobjstcPedLMI
    Next

    Set lobjstcPedCobert = Nothing
    Set lobjstcPedLoc = Nothing
    Set lobjstcPedLMI = Nothing
    Set lobjstcCobert = Nothing
    Set lColecaoCobertura = Nothing
    mfGerarDadosLMI = True

    Exit Function
ControleErro:
    UltimoErro = Err.Description
End Function

