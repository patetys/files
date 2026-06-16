VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V725A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V725A
' Sistema: P46V001
' Simony Sales
' ======================================================================================================================================================
Public bdGed As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V725B
Public UltimoErro As String           ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V725A"
End Property


Private Function mfFormatarConsultaSQL(pobjstc As stcA46V725B) As String
    Dim sSql As String

    sSql = "SELECT login_user_rede ,nom_user ,cod_user_msp ,ender_ip, "
    '###PROJETO SISCOTA
    'COMENTAR a LINHA
    '    sSql = sSql & "  prior_multi,prior_m_taxa,prior_m_dcoml,prior_m_normas , cod_depto, prior_m_franquia "

    ''###PROJETO SISCOTA
    ''DESCOMENTAR LINHAS
    sSql = sSql & "  prior_multi,prior_m_taxa,prior_m_dcoml,prior_m_normas , cod_depto, prior_m_franquia, prior_m_siscota "

    sSql = sSql & "  FROM ged.dbo.tab_user with(nolock) "
    If M46V999.gfPreenchido(pobjstc.login_user_rede) Then
        sSql = sSql & "  WHERE login_user_rede= " & M46V999.gfFormatarCamposSQL(pobjstc.login_user_rede)
    End If
    If M46V999.gfPreenchido(pobjstc.ender_ip) And pobjstc.prior_multi = 2 Then
        sSql = M46V999.gfWhereAnd(sSql) & "  ender_ip= " & M46V999.gfFormatarCamposSQL(pobjstc.ender_ip)
    End If

    mfFormatarConsultaSQL = sSql

End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub

Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V725B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)

    pobjstc.login_user_rede = M46V999.gfTratarNulo(prst("login_user_rede"), tpoString)
    pobjstc.nom_user = M46V999.gfTratarNulo(prst("nom_user"), tpoString)
    pobjstc.cod_user_msp = M46V999.gfTratarNulo(prst("cod_user_msp"), tpoString)
    pobjstc.ender_ip = M46V999.gfTratarNulo(prst("ender_ip"), tpoString)
    pobjstc.prior_multi = M46V999.gfTratarNulo(prst("prior_multi"), tpoNumerico)
    pobjstc.prior_m_taxa = M46V999.gfTratarNulo(prst("prior_m_taxa"), tpoNumerico)
    pobjstc.prior_m_dcoml = M46V999.gfTratarNulo(prst("prior_m_dcoml"), tpoNumerico)
    pobjstc.prior_m_normas = M46V999.gfTratarNulo(prst("prior_m_normas"), tpoNumerico)
    pobjstc.cod_Depto = M46V999.gfTratarNulo(prst("cod_depto"), tpoNumerico)
    pobjstc.prior_m_franquia = M46V999.gfTratarNulo(prst("prior_m_franquia"), tpoNumerico)
    pobjstc.prior_m_siscota = M46V999.gfTratarNulo(prst("prior_m_siscota"), tpoNumerico)

    M46V999.PerfilAcesso = pobjstc.prior_multi

    Select Case M46V999.PerfilAcesso
    Case 1, 2, 3, 4, 5, 6, 7
    Case Else
        UltimoErro = "AGE0134 - O sistema de cotação SYAS MULTI está desativado para as Unidades de Negócios e Filiais."
        Call M46V999.gpGraLog(2, UltimoErro)
        M46V999.gFinalizarPrograma = True
        Exit Sub
    End Select

    Select Case M46V999.PerfilAcesso
    Case e_Perfil_DIG_Gestor, e_Perfil_DIG_IST, e_Perfil_DIG_Yas, e_Perfil_Triagem
        If M46V999.gEndExeSiscad = "" Then
            UltimoErro = "FTE0046 - Erro no arquivo de inicialização ( Endereço do sistema SISCAD não localizado. O sistema será encerrado.)"
            Call M46V999.gpGraLog(2, UltimoErro)
            M46V999.gFinalizarPrograma = True
            Exit Sub
        End If
        M46V999.gTipoAcesso = DIGITACAO
    Case Else
        M46V999.gTipoAcesso = CALCULO
    End Select

    pobjstc.TipoAcesso = M46V999.gTipoAcesso
    'If M46V999.gTipoAcesso = DIGITACAO Then
    '    'Digitação não pode alter a taxa
    '    pobjstc.prior_m_taxa = 0
    'End If

    M46V999.gNomUsuMainFrame = pobjstc.cod_user_msp

    M46O001.PermiteAlterarTaxaFinal = IIf(pobjstc.prior_m_taxa = 1, True, False)
    M46O001.PermiteQQDescontoComercial = IIf(pobjstc.prior_m_dcoml = 1, True, False)
    M46O001.PermiteDigitarSemNormas = IIf(pobjstc.prior_m_normas = 1, True, False)
    M46O001.PermiteAlterarFranquia = IIf(pobjstc.prior_m_franquia = 1, True, False)
    M46O001.PermiteSemSiscota = IIf(pobjstc.prior_m_siscota = 1, True, False)


    If M46V999.gTipoAcesso = DIGITACAO Then
        M46O001.PermiteAlterarFranquia = True
    End If

    If M46O001.PermiteAlterarFranquia = True Then
        M46O001.gAlterarFranquia = True
    Else
        M46O001.gAlterarFranquia = False
    End If

    If M46O001.PermiteSemSiscota = True Then
        M46O001.gPermiteSemSiscota = True
    Else
        M46O001.gPermiteSemSiscota = False
    End If

End Sub

Public Function mfCarregarColecao(pobjstc As stcA46V725B) As Boolean
    Dim objstc As stcA46V725B
    Dim rst As clsYasRecordSet

    On Error GoTo ControleErros

    M46V999.gTipoAcesso = DIGITACAO
    M46O001.PermiteAlterarTaxaFinal = False
    M46O001.PermiteQQDescontoComercial = False
    M46O001.PermiteDigitarSemNormas = False
    M46O001.PermiteAlterarFranquia = False

    If M46V999.gfAbrBasDados("GED", M46V999.gTipGed, M46V999.gEndGed, bdGed, UltimoErro) = False Then GoTo ControleErros

    M46V999.gTipoAcesso = CALCULO

    Call mpLimparColecao

    If M46V999.gfObtRegistro(bdGed, mfFormatarConsultaSQL(pobjstc), rst, UltimoErro) = False Then GoTo ControleErros
    
    If rst.EOF And M46V999.gOrigemChamada = e_Origem_Kit Then
        UltimoErro = "AGE0134 - Usuário " & pobjstc.login_user_rede & " não possui permissões suficientes para acessar o sistema!"
        mfCarregarColecao = False
        GoTo Fim:
    End If
    
    Do While Not rst.EOF
        Set objstc = New stcA46V725B
        Call mpCarregarAtributos(objstc, rst)
        Colecao.Add Item:=objstc, Key:=objstc.Class_ID
        rst.MoveNext
    Loop
    
    mfCarregarColecao = True
Fim:
    Set objstc = Nothing
    Set rst = Nothing
    Call M46V999.gpFecha2(bdGed)
    Exit Function
ControleErros:
    Set objstc = Nothing
    Set rst = Nothing
    mfCarregarColecao = False
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

Private Function mfFormatarAlteracaoSQL(pobjstc As stcA46V725B) As String
    Dim lstrTMP As String                'Query.

    lstrTMP = "UPDATE tab_user  Set "
    lstrTMP = lstrTMP & " prior_multi = " & M46V999.gfFormatarCamposSQL(pobjstc.prior_multi) & ","
    lstrTMP = lstrTMP & " prior_m_taxa = " & M46V999.gfFormatarCamposSQL(pobjstc.prior_m_taxa) & ","
    lstrTMP = lstrTMP & " prior_m_dcoml= " & M46V999.gfFormatarCamposSQL(pobjstc.prior_m_dcoml) & ","
    lstrTMP = lstrTMP & " prior_m_normas= " & M46V999.gfFormatarCamposSQL(pobjstc.prior_m_normas) & ","
    '###PROJETO SISCOTA
    'COMENTAR a LINHA
    '    lstrTMP = lstrTMP & " prior_m_franquia= " & M46V999.gfFormatarCamposSQL(pObjStc.prior_m_franquia)

    ''###PROJETO SISCOTA
    ''DESCOMENTAR LINHAS
    lstrTMP = lstrTMP & " prior_m_franquia= " & M46V999.gfFormatarCamposSQL(pobjstc.prior_m_franquia) & ","
    lstrTMP = lstrTMP & " prior_m_siscota= " & M46V999.gfFormatarCamposSQL(pobjstc.prior_m_siscota)

    lstrTMP = lstrTMP & "  WHERE login_user_rede= " & M46V999.gfFormatarCamposSQL(pobjstc.login_user_rede)

    mfFormatarAlteracaoSQL = lstrTMP
End Function

Public Function mfAlterar(pobjstc As stcA46V725B) As Boolean
    Dim lsSQl As String          'Query.

    On Error GoTo ControleErro

    If M46V999.gfAbrBasDados("GED", M46V999.gTipGed, M46V999.gEndGed, bdGed, _
                             UltimoErro) = False Then GoTo ControleErro

    lsSQl = mfFormatarAlteracaoSQL(pobjstc)
    If M46V999.gfExeSQL(bdGed, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro

    M46V999.PerfilAcesso = pobjstc.prior_multi
    M46O001.PermiteAlterarTaxaFinal = IIf(pobjstc.prior_m_taxa = 1, True, False)
    M46O001.PermiteQQDescontoComercial = IIf(pobjstc.prior_m_dcoml = 1, True, False)
    M46O001.PermiteDigitarSemNormas = IIf(pobjstc.prior_m_normas = 1, True, False)
    M46O001.PermiteAlterarFranquia = IIf(pobjstc.prior_m_franquia = 1, True, False)
    '    '###PROJETO SISCOTA
    '    'DESCOMENTAR LINHAS
    M46O001.PermiteSemSiscota = IIf(pobjstc.prior_m_siscota = 1, True, False)

    If M46O001.PermiteAlterarFranquia = True Then
        M46O001.gAlterarFranquia = True
    Else
        M46O001.gAlterarFranquia = False
    End If

    '    '###PROJETO SISCOTA
    '    'DESCOMENTAR LINHAS
    If M46O001.PermiteSemSiscota = True Then
        M46O001.gPermiteSemSiscota = True
    Else
        M46O001.gPermiteSemSiscota = False
    End If

    mfAlterar = True
    Exit Function
ControleErro:
    mfAlterar = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V701A.mfAlterar ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V701A.mfAlterar ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
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

Public Property Get M46O001() As clsM46O001
    Set M46O001 = M46V999.M46O001
End Property
