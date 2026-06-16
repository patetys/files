VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA28V720A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA28V720A
' Analista: tjgomes
' Properties...
' ======================================================================================================================================================
Private mCnDB As ADODB.Connection    ' Conexão com o Banco de Dados
Private mCnDBApol As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA28V720B
Public UltimoErro As String      ' Descrição do último erro gerado
Public cpfcnpjsacado As String

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A28V720A"
End Property


Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA28V720B, _
                                       ByVal pPropTabela As String) As String

    Dim strTMP As String

    strTMP = " INSERT INTO " & pPropTabela & "P0045200 ( "
    strTMP = strTMP & " NUM_BOLETO,"
    strTMP = strTMP & " NOSSO_NUMERO,"
    strTMP = strTMP & " COD_UNID,"
    strTMP = strTMP & " COD_RAMO,"
    strTMP = strTMP & " COD_BANCO,"
    strTMP = strTMP & " NOM_SEGURADO,"
    strTMP = strTMP & " VAL_PR,"
    strTMP = strTMP & " DAT_VENC,"
    strTMP = strTMP & " COD_CORR,"
    strTMP = strTMP & " NUM_LINHA_DIGIT,"
    strTMP = strTMP & " COD_USER,"
    strTMP = strTMP & " DAT_DIGIT,"
    strTMP = strTMP & " HOR_DIGIT,"
    strTMP = strTMP & " STATUS"

    strTMP = strTMP & ") VALUES ("
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.NUM_BOLETO) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.COD_UNID) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.COD_BANCO) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Segurado) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.VAL_PR) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.DAT_VENC) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corr) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.NUM_LINHA_DIGIT) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.DAT_DIGIT) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.HOR_DIGIT) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.STATUS) & ")"

    mfFormatarInclusaoSQL = strTMP
End Function


Private Function mfFormatarAlteracaoSQL(ByRef pobjstc As stcA28V720B, _
                                        ByVal pPropTabela As String) As String
    Dim strTMP As String

    strTMP = "UPDATE " & pPropTabela & "P0045200 Set "
    strTMP = strTMP & " NUM_BOLETO = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_BOLETO) & ","
    strTMP = strTMP & " NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    strTMP = strTMP & " COD_UNID = " & M46V999.gfFormatarCamposSQL(pobjstc.COD_UNID) & ","
    strTMP = strTMP & " COD_RAMO = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    strTMP = strTMP & " COD_BANCO = " & M46V999.gfFormatarCamposSQL(pobjstc.COD_BANCO) & ","
    strTMP = strTMP & " NOM_SEGURADO = " & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Segurado) & ","
    strTMP = strTMP & " VAL_PR = " & M46V999.gfFormatarCamposSQL(pobjstc.VAL_PR) & ","
    strTMP = strTMP & " DAT_VENC = " & M46V999.gfFormatarCamposSQL(pobjstc.DAT_VENC) & ","
    strTMP = strTMP & " COD_CORR = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Corr) & ","
    strTMP = strTMP & " NUM_LINHA_DIGIT = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_LINHA_DIGIT) & ","
    strTMP = strTMP & " COD_USER = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_User) & ","
    strTMP = strTMP & " DAT_DIGIT = " & M46V999.gfFormatarCamposSQL(pobjstc.DAT_DIGIT) & ","
    strTMP = strTMP & " HOR_DIGIT = " & M46V999.gfFormatarCamposSQL(pobjstc.HOR_DIGIT) & ","
    strTMP = strTMP & " STATUS = " & M46V999.gfFormatarCamposSQL(pobjstc.STATUS)

    strTMP = strTMP & " WHERE "
    strTMP = strTMP & " NUM_BOLETO = " & M46V999.gfFormatarCamposSQL(pobjstc.NUM_BOLETO)

    mfFormatarAlteracaoSQL = strTMP
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA28V720B, _
                                       ByVal pPropTabela As String) As String
    Dim strTMP As String
    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        strTMP = "SELECT top 1 * FROM " & pPropTabela & "P0045200 where nosso_numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
        strTMP = strTMP & " order by DAT_DIGIT  desc, HOR_DIGIT  desc"
    Else
        strTMP = "SELECT top 1 * FROM " & pPropTabela & "P0045200 where status = 0 "
    End If
    mfFormatarConsultaSQL = strTMP
End Function
Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Private Sub mpCarregarAtributos(ByRef pobjstc As stcA28V720B, prst As clsYasRecordSet)
    pobjstc.NUM_BOLETO = M46V999.gfTratarNulo(prst("NUM_BOLETO"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("NOSSO_NUMERO"), tpoString)
    pobjstc.COD_UNID = M46V999.gfTratarNulo(prst("COD_UNID"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("COD_RAMO"), tpoNumerico)
    pobjstc.COD_BANCO = M46V999.gfTratarNulo(prst("COD_BANCO"), tpoNumerico)
    pobjstc.Nom_Segurado = M46V999.gfTratarNulo(prst("NOM_SEGURADO"), tpoString)
    pobjstc.CPF_CNPJ_Segurado = fCpfCnpjSacado(pobjstc.nosso_numero)
    pobjstc.VAL_PR = M46V999.gfTratarNulo(prst("VAL_PR"), TpoValor)
    pobjstc.DAT_VENC = M46V999.gfTratarNulo(prst("DAT_VENC"), tpoNumerico)
    pobjstc.Cod_Corr = M46V999.gfTratarNulo(prst("COD_CORR"), tpoNumerico)
    pobjstc.NUM_LINHA_DIGIT = M46V999.gfTratarNulo(prst("NUM_LINHA_DIGIT"), tpoString)
    pobjstc.Cod_User = M46V999.gfTratarNulo(prst("COD_USER"), tpoNumerico)
    pobjstc.DAT_DIGIT = M46V999.gfTratarNulo(prst("DAT_DIGIT"), tpoNumerico)
    pobjstc.HOR_DIGIT = M46V999.gfTratarNulo(prst("HOR_DIGIT"), tpoNumerico)
    pobjstc.STATUS = M46V999.gfTratarNulo(prst("STATUS"), tpoNumerico)
End Sub

Public Function gfAlterar(ByRef pobjstc As stcA28V720B, _
                          ByVal pPropTabela As String) As Boolean
    Dim sSql As String
    Dim lRegAlt As Long
    On Error GoTo ControleErros

    If CnDB Is Nothing Then GoTo ControleErros

    sSql = mfFormatarAlteracaoSQL(pobjstc, pPropTabela)

    If M46V999.gfExeSQL(CnDB, sSql, UltimoErro, lRegAlt) <> 0 Then GoTo ControleErros
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    gfAlterar = True
    Exit Function
ControleErros:
    gfAlterar = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
    End If
End Function

Public Function AtualizaBoletoRegistrado(ByRef pobjstc As stcA28V720B) As Boolean
    AtualizaBoletoRegistrado = False
    
    Dim rst As clsYasRecordSet
    Dim strSQL As String
    Dim iOperacao As Integer ' 0 = Inserir, 1 = Atualizar, 2 = Gerar erro
    Dim dataCorte As Long
    Dim lPropTabela As String
    Dim lRegAlt As Long
    
    On Error GoTo ControleErros

    Call M46V999.gfObtSYASMULTINI(lPropTabela)
    'If lPropTabela = "" Then lPropTabela = pPropTabela
    If CnDB Is Nothing Then GoTo ControleErros
    
    ' Procurar registro na P0045200 para saber se existe, o status e a data de vencimento
    strSQL = "SELECT * FROM " & lPropTabela & "P0045200 WHERE NUM_BOLETO = " & Right(pobjstc.NUM_BOLETO, 7)
    If Not M46V999.gfObtRegistro(CnDB, strSQL, rst, UltimoErro) Then GoTo ControleErros
    
    If rst.EOF Then
        iOperacao = 0 ' Inserir
    Else
        dataCorte = CLng(Format(Date - 60, "yyyymmdd"))
        If rst("STATUS") = 0 Then
            iOperacao = 1 'Boleto está livre - Atualizar
        ElseIf rst("DAT_VENC") > dataCorte Then
            iOperacao = 1 'Boleto existe mas é velho - Atualizar
        Else
            iOperacao = 2 'Boleto ainda é novo, não pode atualizar
        End If
    End If
    
    ' Realizar a operação
    If iOperacao = 0 Then ' Inserir
        strSQL = mfFormatarInclusaoSQL(pobjstc, lPropTabela)
        If M46V999.gfExeSQL(CnDB, strSQL, UltimoErro, lRegAlt) <> 0 Then GoTo ControleErros
    ElseIf iOperacao = 1 Then ' Atualizar
        strSQL = mfFormatarAlteracaoSQL(pobjstc, lPropTabela)
        If M46V999.gfExeSQL(CnDB, strSQL, UltimoErro, lRegAlt) <> 0 Then GoTo ControleErros
    Else ' Erro
        Call M46V999.gpGraLog(1, "Boleto " & Right(pobjstc.NUM_BOLETO, 7) & " já existe na tabela P0045200 com vencimento mais próximo que 60 dias.")
    End If
    
    AtualizaBoletoRegistrado = True
    GoTo SAIDA:
    
    
ControleErros:
    AtualizaBoletoRegistrado = False
    If M46V999.gfPreenchido(UltimoErro) = False And Err.Number > 0 Then
        UltimoErro = Err.Description
    End If

SAIDA:
    Set rst = Nothing
    Exit Function
    
    
End Function


Public Function gfCarregarBoletoDisp(ByVal pNossoNumero As String, pNovo As Boolean, _
                                     Optional ByVal pNossoNumeroOriginal As String, _
                                     Optional ByVal pSohVerificacao As Boolean = False, Optional pPropTabela As String) As Boolean
    Dim objstc As stcA28V720B
    Dim rst As clsYasRecordSet
    Dim lCount As Integer
    Dim strTMP As String
    Dim lRegAlt As Long
    Dim lUtilizouOriginal As Integer
    Dim lBoletoNovo As Integer
    Dim lPropTabela As String

    On Error GoTo ControleErros

    Call M46V999.gfObtSYASMULTINI(lPropTabela)
    If lPropTabela = "" Then lPropTabela = pPropTabela

    ''Call M46V999.gpGraLog(0, "Busca boleto disp : " & M46V999.gEndP0045200)
    If CnDB Is Nothing Then GoTo ControleErros

    lUtilizouOriginal = 0
    lBoletoNovo = 0
    Set objstc = New stcA28V720B
    Set objstc.M46V999 = M46V999

    lCount = 0
    If pNovo = False Then
        objstc.nosso_numero = pNossoNumero
    Else
        lCount = 1
    End If
    ''Call M46V999.gpGraLog(0, "Procura por " & lCount & "-" & objstc.Nosso_Numero)
Retorna:
    If Not M46V999.gfObtRegistro(CnDB, mfFormatarConsultaSQL(objstc, lPropTabela), rst, UltimoErro) Then GoTo ControleErros
    Call mpLimparColecao
    If rst.EOF Then
        If lCount = 0 Then
            lCount = 1
            If M46V999.gfPreenchido(pNossoNumeroOriginal) And pNossoNumeroOriginal <> "0" And objstc.nosso_numero <> pNossoNumeroOriginal Then
                objstc.nosso_numero = pNossoNumeroOriginal
                lUtilizouOriginal = 1
            Else
                objstc.nosso_numero = ""
            End If
            GoTo Retorna
        ElseIf lCount = 1 And M46V999.gfPreenchido(objstc.nosso_numero) Then
            objstc.nosso_numero = ""
            lCount = 2
            GoTo Retorna
        Else
            'Não existem boletos disponíveis.
            'Call M46V999.gpGraLog(0, "Não tem boleto")
            UltimoErro = "Não existem boletos disponíveis, por favor conectar com a Yasuda para baixar o lote de boletos."
            gfCarregarBoletoDisp = False
            GoTo SAIDA
        End If
    Else
        Set objstc = New stcA28V720B
        Set objstc.M46V999 = M46V999

        Call mpCarregarAtributos(objstc, rst)
        'Call M46V999.gpGraLog(0, "Novo - " & objstc.NUM_BOLETO)
        If objstc.nosso_numero <> pNossoNumeroOriginal Then
            lUtilizouOriginal = 0
        End If
        Colecao.Add Item:=objstc, Key:=objstc.Class_ID
        rst.MoveNext
    End If
    objstc.nosso_numero = pNossoNumero
    If (lCount = 1 Or lCount = 2) Then
        'Call M46V999.gpGraLog(0, "Altera registro boleto")
        If pSohVerificacao = True Then
            If lUtilizouOriginal = 1 Then
                objstc.nosso_numero = pNossoNumeroOriginal
                GoTo Fim
            End If
            'objstc.Nosso_Numero = ""
            'GoTo Retorna
        Else
            strTMP = "UPDATE " & lPropTabela & "P0045200 Set "
            strTMP = strTMP & " NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(objstc.nosso_numero)
            strTMP = strTMP & " , STATUS = 1"
            strTMP = strTMP & " Where NUM_BOLETO = " & M46V999.gfFormatarCamposSQL(objstc.NUM_BOLETO)
            strTMP = strTMP & " and STATUS = 0"

            If M46V999.gfExeSQL(CnDB, strTMP, UltimoErro, lRegAlt) <> 0 Then GoTo ControleErros
            If lRegAlt = 0 Then
                If lUtilizouOriginal = 1 Then
                    objstc.nosso_numero = pNossoNumeroOriginal
                    GoTo Fim
                End If
                objstc.nosso_numero = ""
                GoTo Retorna
            End If
        End If
    End If
Fim:

    gfCarregarBoletoDisp = True
    GoTo SAIDA:

ControleErros:
    gfCarregarBoletoDisp = False
    If M46V999.gfPreenchido(UltimoErro) = False And Err.Number > 0 Then
        UltimoErro = Err.Description
    End If

SAIDA:
    Set objstc = Nothing
    Set rst = Nothing
    Exit Function

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
        M46V999.gBaseLocal = True
        If M46V999.gfAbrBasDados("P0045200", M46V999.gTipP0045200, M46V999.gEndP0045200, mCnDB, _
                                 UltimoErro) = False Then Set mCnDB = Nothing
        M46V999.gBaseLocal = False
    End If

Retorno:
    Set CnDB = mCnDB
End Property

Public Property Get CnDBApol() As ADODB.Connection

    If mCnDBApol Is Nothing Then
        If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, mCnDBApol, _
                                 UltimoErro) = False Then Set mCnDBApol = Nothing
    End If

Retorno:
    Set CnDBApol = mCnDBApol
End Property

Private Function fCpfCnpjSacado(NossoNumero As String) As String
    Dim sSql As String
    Dim rst As clsYasRecordSet
    Dim lPropTabela As String

    Call M46V999.gfObtSYASMULTINI(lPropTabela)

    If cpfcnpjsacado <> "" Then
        fCpfCnpjSacado = cpfcnpjsacado
    Else
        fCpfCnpjSacado = "***"
        On Error GoTo ControleErros

        sSql = "SELECT NUM_CGC_CPF FROM " & lPropTabela & "TAB_PED WHERE NOSSO_NUMERO like '" & NossoNumero & "' "

        If M46V999.gfObtRegistro(CnDB, sSql, rst, UltimoErro) Then
            If Not rst.EOF Then
                cpfcnpjsacado = rst("NUM_CGC_CPF")
                fCpfCnpjSacado = cpfcnpjsacado
            End If
        End If
    End If
SAIDA:

    GoSub FechaObj
    Exit Function

ControleErros:
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
    End If

FechaObj:
    Set rst = Nothing
    Return
End Function

