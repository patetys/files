VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTabConsDadosExternos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Colecao     As Collection        ' Coleção de objetos do tipo stcTabConsDadosExternos
Public UltimoErro  As String            ' Descrição do último erro gerado

Private CnDB       As ADODB.Connection

Sub AbreBancoCnDB()
    Dim lMsgErr As String
    Set CnDB = New ADODB.Connection
    'If gM46V999.gfAbrBasSQL("P0044800", 2, gM46V999.gSerSYASEMIS, "P0044800", "workflow", "yasworkflow", CnDB, lMsgErr) = False Then
    If Not gM46V999.gfAbrBasSQL("P0044800", gM46V999.gTipGed, gM46V999.gSerGed, "Ged", gM46V999.gUidGed, gM46V999.gPwdGed, CnDB, lMsgErr) Then
        UltimoErro = "[ERR] - " & lMsgErr
    End If
End Sub

Private Function mfFormatarInclusaoSQL(pobjstc As stcTabConsDadosExternos) As String
    Dim lInsert As String                'Query.

    lInsert = " INSERT INTO P0044800.dbo.Tab_ConsDadosExternos ( "
    lInsert = lInsert & " Num_PI,"
    lInsert = lInsert & " Num_Item,"
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " TipoSistema,"
    lInsert = lInsert & " Tip_Chave1,"
    lInsert = lInsert & " Chave1,"
    lInsert = lInsert & " Tip_Chave2,"
    lInsert = lInsert & " Chave2,"
    lInsert = lInsert & " Tip_Chave3,"
    lInsert = lInsert & " Chave3,"
    lInsert = lInsert & " Tip_Chave4,"
    lInsert = lInsert & " Chave4,"
    lInsert = lInsert & " Dat_Processamento,"
    lInsert = lInsert & " Hor_Processamento"
    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.TipoSistema) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.Tip_Chave1) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.Chave1) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.Tip_Chave2) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.Chave2) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.Tip_Chave3) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.Chave3) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.Tip_Chave4) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.Chave4) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.Dat_Processamento) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(pobjstc.Hor_Processamento) & "); "

    mfFormatarInclusaoSQL = lInsert
End Function

Private Function mfFormatarAlteracaoSQL(pobjstc As stcTabConsDadosExternos) As String

    Dim lUpdate As String                'Query.

    lUpdate = "UPDATE P0044800.dbo.Tab_ConsDadosExternos Set "
    lUpdate = lUpdate & " Cod_Ramo = " & gM46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lUpdate = lUpdate & " Tip_Chave1 = " & gM46V999.gfFormatarCamposSQL(pobjstc.Tip_Chave1) & ","
    lUpdate = lUpdate & " Chave1 = " & gM46V999.gfFormatarCamposSQL(pobjstc.Chave1) & ","
    lUpdate = lUpdate & " Tip_Chave2 = " & gM46V999.gfFormatarCamposSQL(pobjstc.Tip_Chave2) & ","
    lUpdate = lUpdate & " Chave2 = " & gM46V999.gfFormatarCamposSQL(pobjstc.Chave2) & ","
    lUpdate = lUpdate & " Tip_Chave3 = " & gM46V999.gfFormatarCamposSQL(pobjstc.Tip_Chave3) & ","
    lUpdate = lUpdate & " Chave3 = " & gM46V999.gfFormatarCamposSQL(pobjstc.Chave3) & ","
    lUpdate = lUpdate & " Tip_Chave4 = " & gM46V999.gfFormatarCamposSQL(pobjstc.Tip_Chave4) & ","
    lUpdate = lUpdate & " Chave4 = " & gM46V999.gfFormatarCamposSQL(pobjstc.Chave4) & ","
    lUpdate = lUpdate & " Dat_Processamento = " & gM46V999.gfFormatarCamposSQL(pobjstc.Dat_Processamento) & ","
    lUpdate = lUpdate & " Hor_Processamento = " & gM46V999.gfFormatarCamposSQL(pobjstc.Hor_Processamento)
    lUpdate = lUpdate & " WHERE "
    lUpdate = lUpdate & " TipoSistema = " & gM46V999.gfFormatarCamposSQL(pobjstc.TipoSistema) & " AND "
    lUpdate = lUpdate & " Num_PI = " & gM46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & " AND "
    lUpdate = lUpdate & " Num_Item = " & gM46V999.gfFormatarCamposSQL(pobjstc.Num_Item)

    mfFormatarAlteracaoSQL = lUpdate
End Function

Private Function mfFormatarExclusaoSQL(pobjstc As stcTabConsDadosExternos) As String
    Dim lDelete As String                'Query.

    lDelete = "DELETE FROM P0044800.dbo.Tab_ConsDadosExternos WHERE "
    lDelete = lDelete & " TipoSistema = " & gM46V999.gfFormatarCamposSQL(pobjstc.TipoSistema) & " AND "
    lDelete = lDelete & " Num_PI = " & gM46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)

    mfFormatarExclusaoSQL = lDelete
End Function

Private Function mfFormatarConsultaSQL(pobjstc As stcTabConsDadosExternos) As String
    Dim lSelect As String

    lSelect = "SELECT * FROM P0044800.dbo.Tab_ConsDadosExternos WHERE "
    If pobjstc.Dat_Processamento = 999999 Then
        lSelect = lSelect & " Dat_Processamento = 0 "
        If pobjstc.TipoSistema <> 0 Then
            lSelect = lSelect & " and TipoSistema = " & gM46V999.gfFormatarCamposSQL(pobjstc.TipoSistema)
        End If
    Else
        lSelect = lSelect & " TipoSistema = " & gM46V999.gfFormatarCamposSQL(pobjstc.TipoSistema) & " AND "
        lSelect = lSelect & " Num_PI = " & gM46V999.gfFormatarCamposSQL(pobjstc.NUM_PI)
        If pobjstc.Num_Item <> 0 Then
            lSelect = lSelect & " and  Num_Item = " & gM46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
        End If
    End If
    lSelect = lSelect & " order by Num_PI desc , Num_Item , TipoSistema, Chave1"
    mfFormatarConsultaSQL = lSelect
End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcTabConsDadosExternos, prst As clsYasRecordSet)
    pobjstc.NUM_PI = gM46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
    pobjstc.Num_Item = gM46V999.gfTratarNulo(prst("Num_Item"), tpoNumerico)
    pobjstc.Cod_Ramo = gM46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.TipoSistema = gM46V999.gfTratarNulo(prst("TipoSistema"), tpoNumerico)
    pobjstc.Tip_Chave1 = gM46V999.gfTratarNulo(prst("Tip_Chave1"), tpoNumerico)
    pobjstc.Chave1 = gM46V999.gfTratarNulo(prst("Chave1"), tpoString)
    pobjstc.Tip_Chave2 = gM46V999.gfTratarNulo(prst("Tip_Chave2"), tpoNumerico)
    pobjstc.Chave2 = gM46V999.gfTratarNulo(prst("Chave2"), tpoString)
    pobjstc.Tip_Chave3 = gM46V999.gfTratarNulo(prst("Tip_Chave3"), tpoNumerico)
    pobjstc.Chave3 = gM46V999.gfTratarNulo(prst("Chave3"), tpoString)
    pobjstc.Tip_Chave4 = gM46V999.gfTratarNulo(prst("Tip_Chave4"), tpoNumerico)
    pobjstc.Chave4 = gM46V999.gfTratarNulo(prst("Chave4"), tpoString)
    pobjstc.Dat_Processamento = gM46V999.gfTratarNulo(prst("Dat_Processamento"), tpoNumerico)
    pobjstc.Hor_Processamento = gM46V999.gfTratarNulo(prst("Hor_Processamento"), tpoNumerico)
End Sub

Public Function gfIncluir(pobjstc As stcTabConsDadosExternos) As Boolean
    Dim lInsert As String  'Query.

    On Local Error GoTo ControleErros

    lInsert = mfFormatarInclusaoSQL(pobjstc)

    If gM46V999.gfExeSQL(CnDB, lInsert, UltimoErro) <> 0 Then GoTo ControleErros

    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    gfIncluir = True

    Exit Function

ControleErros:

    gfIncluir = False

    If Not gM46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If

End Function

Public Function gfAlterar(pobjstc As stcTabConsDadosExternos) As Boolean
    Dim lUpdate As String  'Query.

    On Error GoTo ControleErros
    lUpdate = mfFormatarAlteracaoSQL(pobjstc)
    If CnDB.State = 0 Then
        AbreBancoCnDB
    End If
    If gM46V999.gfExeSQL(CnDB, lUpdate, UltimoErro) <> 0 Then GoTo ControleErros
    If Colecao.Count > 0 Then
        On Error Resume Next
        Colecao.Remove (pobjstc.Class_ID)
        On Error GoTo ControleErros
    End If
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    gfAlterar = True
    Exit Function
ControleErros:
    gfAlterar = False
    If Not gM46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If
End Function

Public Function gfExcluir(pobjstc As stcTabConsDadosExternos) As Boolean
    Dim lDelete As String  'Query.

    On Error GoTo ControleErros
    lDelete = mfFormatarExclusaoSQL(pobjstc)
    If gM46V999.gfExeSQL(CnDB, lDelete, UltimoErro) <> 0 Then GoTo ControleErros
    If Colecao.Count > 0 Then
        Colecao.Remove (pobjstc.Class_ID)
    End If
    gfExcluir = True
    Exit Function
ControleErros:
    gfExcluir = False
    If Not gM46V999.gfPreenchido(UltimoErro) Then
        UltimoErro = Err.Description
    End If
End Function

Public Function gfCarregarColecao(pobjstc As stcTabConsDadosExternos) As Boolean
    Dim objstc  As stcTabConsDadosExternos      'objeto da classe de estrutura da tabela de embarcação.
    Dim rst     As clsYasRecordSet  'Registro.

    On Error GoTo ControleErros
    If Not gM46V999.gfObtRegistro(CnDB, mfFormatarConsultaSQL(pobjstc), _
                                  rst, UltimoErro) Then GoTo ControleErros

    Call mpLimparColecao
    Do While Not rst.EOF
        Set objstc = New stcTabConsDadosExternos
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
    AbreBancoCnDB
End Sub

Private Sub Class_Terminate()
    Set Colecao = Nothing
    On Error Resume Next
    CnDB.Close
    Set CnDB = Nothing
End Sub


