VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V718A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V718A
' Data/Hora: 04/10/2005 - 13:42:14
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V718B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V718A"
End Property

Private Function mfFormatarUpDateTransmitidoSQL(ByRef pobjstc As stcA46V718B, _
                                                ByVal pPropTabela As String, _
                                                ByVal pNomeTabImport As String) As String
    Dim lstrTMP As String                'Query.

    lstrTMP = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ped Set "
    lstrTMP = lstrTMP & " Cod_Sit_Ped = 3"  'Trasmitido (e_SitPedido_Transmitido)
    lstrTMP = lstrTMP & " WHERE "
    lstrTMP = lstrTMP & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    lstrTMP = lstrTMP & " AND Cod_Sit_Ped <> 3 "    'evita acesso desnecessário ao registro

    mfFormatarUpDateTransmitidoSQL = lstrTMP
End Function
Private Function mfFormatarUpDateIndCancel(ByRef pobjstc As stcA46V718B, _
                                           ByVal pPropTabela As String, _
                                           ByVal pNomeTabImport As String) As String
    Dim lstrTMP As String                'Query.

    lstrTMP = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ped_Transm Set "
    lstrTMP = lstrTMP & " Ind_Cancel = " & pobjstc.Ind_Cancel
    lstrTMP = lstrTMP & ", Cod_Acao = 80"
    lstrTMP = lstrTMP & " WHERE "
    lstrTMP = lstrTMP & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)

    mfFormatarUpDateIndCancel = lstrTMP
End Function

Private Function mfFormatarUpdate(ByRef pobjstc As stcA46V718B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As String
    Dim lstrTMP As String                'Query.

    lstrTMP = "UPDATE " & pPropTabela & pNomeTabImport & "Tab_Ped_transm Set "
    lstrTMP = lstrTMP & " Num_PI = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)
    lstrTMP = lstrTMP & ", Cod_Acao = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Acao)
    lstrTMP = lstrTMP & ", Dat_Emissao = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Emissao)
    lstrTMP = lstrTMP & ", Dat_Ult_Inf = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ult_Inf)
    lstrTMP = lstrTMP & " WHERE "
    lstrTMP = lstrTMP & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    mfFormatarUpdate = lstrTMP
End Function

Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V718B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lstrTMP As String                'Query.

    lstrTMP = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Transm ( "
    lstrTMP = lstrTMP & " Nosso_Numero,"
    lstrTMP = lstrTMP & " Num_Ped,"
    lstrTMP = lstrTMP & " Num_Item,"
    lstrTMP = lstrTMP & " Nom_Arq_Envio,"
    lstrTMP = lstrTMP & " Dat_Hor_Envio,"
    lstrTMP = lstrTMP & " Nom_Arq_Retorno,"
    lstrTMP = lstrTMP & " Dat_Hor_Retorno,"
    lstrTMP = lstrTMP & " Ind_Proposta,"
    lstrTMP = lstrTMP & " Nom_Segurado,"
    lstrTMP = lstrTMP & " Cod_Ramo,"
    lstrTMP = lstrTMP & " Dat_Ini_Vig,"
    lstrTMP = lstrTMP & " Dat_Ter_Vig,"
    lstrTMP = lstrTMP & " Cod_Tip_Logr_Risco,"
    lstrTMP = lstrTMP & " Nom_Logr_Risco,"
    lstrTMP = lstrTMP & " Dsc_Compl_Risco,"
    lstrTMP = lstrTMP & " Nom_Bairro_Risco,"
    lstrTMP = lstrTMP & " Nom_Cid_Risco,"
    lstrTMP = lstrTMP & " Sig_UF_Risco,"
    lstrTMP = lstrTMP & " Num_CEP_Risco,"
    lstrTMP = lstrTMP & " Num_Risco,"
    lstrTMP = lstrTMP & " Tip_Emissao,"
    lstrTMP = lstrTMP & " Num_PI"

    lstrTMP = lstrTMP & ") VALUES ("
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Arq_Envio) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Hor_Envio) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Arq_Retorno) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Hor_Retorno) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Ind_Proposta) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Segurado) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ini_Vig) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ter_Vig) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Tip_Logr_Risco) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Logr_Risco) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dsc_Compl_Risco) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Bairro_Risco) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Cid_Risco) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Sig_UF_Risco) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cep_Risco) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Risco) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Emissao) & ","
    lstrTMP = lstrTMP & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ")"

    mfFormatarInclusaoSQL = lstrTMP
End Function


Private Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V718B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lstrTMP As String                'Query.
    lstrTMP = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Transm WHERE "
    lstrTMP = lstrTMP & " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)

    mfFormatarExclusaoSQL = lstrTMP
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V718B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lstrTMP As String                'Query.
    lstrTMP = "SELECT * FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Transm with(nolock) "
    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Nosso_NUmero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If
    If val(pobjstc.Num_Ped) <> 0 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Num_Ped = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped)
    End If
    If val(pobjstc.Dat_Ini_Vig) <> 0 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Dat_Ini_Vig = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ini_Vig)
    End If
    If val(pobjstc.Dat_Ter_Vig) <> 0 Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Dat_Ter_Vig = " & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ter_Vig)
    End If
    If M46V999.gfPreenchido(pobjstc.Nom_Segurado) Then
        lstrTMP = M46V999.gfWhereAnd(lstrTMP) & " Nom_Segurado like '%" & pobjstc.Nom_Segurado & "%'"
    End If
    lstrTMP = lstrTMP & " order by Num_Ped, Num_Item"
    mfFormatarConsultaSQL = lstrTMP
End Function
Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V718B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)

    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoString)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_Item"), tpoString)
    pobjstc.Nom_Arq_Envio = M46V999.gfTratarNulo(prst("Nom_Arq_Envio"), tpoString)
    pobjstc.Dat_Hor_Envio = M46V999.gfTratarNulo(prst("Dat_Hor_Envio"), tpoString)
    pobjstc.Nom_Arq_Retorno = M46V999.gfTratarNulo(prst("Nom_Arq_Retorno"), tpoString)
    pobjstc.Dat_Hor_Retorno = M46V999.gfTratarNulo(prst("Dat_Hor_Retorno"), tpoString)
    pobjstc.Ind_Proposta = M46V999.gfTratarNulo(prst("Ind_Proposta"), tpoString)
    pobjstc.Nom_Segurado = M46V999.gfTratarNulo(prst("Nom_Segurado"), tpoString)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoString)
    pobjstc.Dat_Ini_Vig = M46V999.gfTratarNulo(prst("Dat_Ini_Vig"), tpoString)
    pobjstc.Dat_Ter_Vig = M46V999.gfTratarNulo(prst("Dat_Ter_Vig"), tpoString)
    pobjstc.Cod_Tip_Logr_Risco = M46V999.gfTratarNulo(prst("Cod_Tip_Logr_Risco"), tpoString)
    pobjstc.Nom_Logr_Risco = M46V999.gfTratarNulo(prst("Nom_Logr_Risco"), tpoString)
    pobjstc.Dsc_Compl_Risco = M46V999.gfTratarNulo(prst("Dsc_Compl_Risco"), tpoString)
    pobjstc.Nom_Bairro_Risco = M46V999.gfTratarNulo(prst("Nom_Bairro_Risco"), tpoString)
    pobjstc.Nom_Cid_Risco = M46V999.gfTratarNulo(prst("Nom_Cid_Risco"), tpoString)
    pobjstc.Sig_UF_Risco = M46V999.gfTratarNulo(prst("Sig_UF_Risco"), tpoString)
    pobjstc.Num_Cep_Risco = M46V999.gfTratarNulo(prst("Num_CEP_Risco"), tpoString)
    pobjstc.Num_Risco = M46V999.gfTratarNulo(prst("Num_Risco"), tpoString)
    pobjstc.Tip_Emissao = M46V999.gfTratarNulo(prst("Tip_Emissao"), tpoString)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
    pobjstc.Cod_Acao = M46V999.gfTratarNulo(prst("Cod_Acao"), tpoString)
    pobjstc.Dat_Ult_Inf = M46V999.gfTratarNulo(prst("Dat_Ult_Inf"), tpoString)
    pobjstc.Ind_Cancel = M46V999.gfTratarNulo(prst("Ind_Cancel"), tpoString)
    pobjstc.Dat_Emissao = M46V999.gfTratarNulo(prst("Dat_Emissao"), tpoString)

End Sub
Public Function mfIncluir(ByRef pobjstc As stcA46V718B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean

    Dim lsSQl As String           'Query.

    On Local Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros

Retorno:
    lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
        If InStr(1, UCase(UltimoErro), "'NUM_PI'") > 0 Then
            lsSQl = " alter table " & pNomeTabImport & "Tab_Ped_Transm add COLUMN  Num_PI int;"
            If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
                Call M46V999.gpGraLog(2, UltimoErro)
                Exit Function
            End If
            GoTo Retorno
        End If
        GoTo ControleErros
    End If
    Colecao.Add Item:=pobjstc

    'Retirado da Dll e adicionado na procedure proc_StartFluxoEmissaoTransmissaoMulti
'    If pobjstc.Ind_Proposta = 1 Then
'        'Proposta transmitida OK.'
'        'Atualiza COD_SIT_PED na tabela TAB_PED
'        lsSQl = mfFormatarUpDateTransmitidoSQL(pobjstc, pPropTabela, pNomeTabImport)
'        If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
'    End If

    mfIncluir = True
    Exit Function
ControleErros:
    mfIncluir = False

    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V718A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V718A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfUpdateSitTransm(ByRef pobjstc As stcA46V718B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lsSQl As String           'Query.

    On Local Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros

    lsSQl = mfFormatarUpdate(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
        GoTo ControleErros
    End If

    'Colecao.Add Item:=pObjStc

    mfUpdateSitTransm = True
    Exit Function
ControleErros:
    mfUpdateSitTransm = False

    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfUpdateCancel(ByRef pobjstc As stcA46V718B, _
                               ByVal pPropTabela As String, _
                               ByVal pNomeTabImport As String) As Boolean
    Dim lsSQl As String           'Query.

    On Local Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros

    lsSQl = mfFormatarUpDateIndCancel(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
        GoTo ControleErros
    End If

    'Colecao.Add Item:=pObjStc

    mfUpdateCancel = True
    Exit Function
ControleErros:
    mfUpdateCancel = False

    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfExcluir(ByRef pobjstc As stcA46V718B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String) As Boolean
    Dim lsSQl As String           'Query.

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErros
    If Colecao.Count > 0 Then
        Colecao.Remove (pobjstc.Class_ID)
    End If
    Set pobjstc = Nothing
    mfExcluir = True
    Exit Function
ControleErros:
    mfExcluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V718A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V718A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V718B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V718B          'Objeto da classe de estrutura stcA46V718B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.
    Dim lCmdDDL As String

    On Error GoTo ControleErros
    If Not mfConectar Then GoTo ControleErros
    lsSQl = mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport)
Retorno:
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, _
                             lrst, UltimoErro) = False Then GoTo ControleErros
    Call mpLimparColecao
    If lrst.RecordCount > 1000 Then
        UltimoErro = "FGE0023 - Pesquisa resultou em " & lrst.RecordCount & _
                   " registros. Restrinja critérios de pesquisa. (protocolos)"
        Exit Function
    End If
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V718B
        Call mpCarregarAtributos(lobjstc, lrst)

        Colecao.Add Item:=lobjstc
        lrst.MoveNext
    Loop
    mfCarregarColecao = True
    Set lobjstc = Nothing
    Set lrst = Nothing
    Exit Function
ControleErros:
    Set lobjstc = Nothing
    Set lrst = Nothing
    If Err.Number = 3265 Then
        'Cria o campos na tabela
        On Error Resume Next
        lCmdDDL = "ALTER TABLE TAB_PED_TRANSM ADD COLUMN Num_PI INT;"
        If M46V999.gfExeSQL(ConexaoBD, lCmdDDL, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(0, UltimoErro)
            'Exit Function
        End If
        lCmdDDL = "ALTER TABLE TAB_PED_TRANSM ADD COLUMN Cod_Acao INT;"
        If M46V999.gfExeSQL(ConexaoBD, lCmdDDL, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(0, UltimoErro)
            'Exit Function
        End If
        lCmdDDL = "ALTER TABLE TAB_PED_TRANSM ADD COLUMN Dat_Emissao INT;"
        If M46V999.gfExeSQL(ConexaoBD, lCmdDDL, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(0, UltimoErro)
            'Exit Function
        End If
        lCmdDDL = "ALTER TABLE TAB_PED_TRANSM ADD COLUMN Dat_Ult_Inf NUMERIC;"
        If M46V999.gfExeSQL(ConexaoBD, lCmdDDL, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(0, UltimoErro)
            'Exit Function
        End If

        lCmdDDL = "ALTER TABLE TAB_PED_TRANSM ADD COLUMN Ind_Cancel NUMERIC;"
        If M46V999.gfExeSQL(ConexaoBD, lCmdDDL, UltimoErro) <> 0 Then
            Call M46V999.gpGraLog(0, UltimoErro)
            'Exit Function
        End If

        Err.Number = 0
        GoTo Retorno
    End If
    mfCarregarColecao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V718A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V718A.mfCarregarColecao ERRO - " & UltimoErro & _
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
