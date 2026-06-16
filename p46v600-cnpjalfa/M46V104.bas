Attribute VB_Name = "basM46V104"
Option Explicit

'Public objSession As MAPI.Session
Public Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
                                            (ByVal hwndParent As Long, ByVal fRequest As Long, _
                                             ByVal lpszDriver As String, ByVal lpszAttributes As String) _
                                             As Long

Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128


Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, _
                                                        ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal _
                                                          hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal _
                                                                  hpvSource&, ByVal cbCopy&)


Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type




Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Public Sub gpVerificaSerasa(ByVal pNumPI As Long, ByVal pCodRamo As Integer, ByVal pNumCnpjCPF As Double, ByRef pMensagem As String)
    Dim oObjClsDados As New clsTabConsDadosExternos
    Dim oObjStrDados As New stcTabConsDadosExternos

    Err.Clear
    On Error GoTo TrataErro

    Set oObjStrDados = New stcTabConsDadosExternos
    Set oObjClsDados = New clsTabConsDadosExternos

    oObjStrDados.NUM_PI = pNumPI
    oObjStrDados.Cod_Ramo = pCodRamo
    oObjStrDados.TipoSistema = e_TipoSistema_Serasa
    oObjStrDados.Tip_Chave1 = e_TipoChave_Cnpj
    oObjStrDados.Chave1 = pNumCnpjCPF


    If Not oObjClsDados.gfCarregarColecao(oObjStrDados) Then
        pMensagem = oObjClsDados.UltimoErro
        Exit Sub
    End If

    If oObjClsDados.Colecao.Count = 0 Then
        If Not oObjClsDados.gfIncluir(oObjStrDados) Then
            Call gM46V999.gpGraLog(1, oObjClsDados.UltimoErro)
        End If
    End If

    Set oObjStrDados = Nothing
    Set oObjClsDados = Nothing

    Exit Sub
TrataErro:
    gM46V999.gpMenErro ("gpVerificaSerasa Erro: " & Err & "-" & Error)

End Sub


Public Sub gpVerificaBoaVista(ByVal pNumPI As Long, ByVal pCodRamo As Integer, ByVal pNumCnpjCPF As Double, ByRef pMensagem As String)
    Dim oObjClsDados As New clsTabConsDadosExternos
    Dim oObjStrDados As New stcTabConsDadosExternos

    Err.Clear
    On Error GoTo TrataErro


    Set oObjStrDados = New stcTabConsDadosExternos
    Set oObjClsDados = New clsTabConsDadosExternos


    oObjStrDados.NUM_PI = pNumPI
    oObjStrDados.Cod_Ramo = pCodRamo
    oObjStrDados.TipoSistema = e_TipoSistema_BoaVista
    oObjStrDados.Tip_Chave1 = e_TipoChave_Cnpj
    oObjStrDados.Chave1 = pNumCnpjCPF


    If Not oObjClsDados.gfCarregarColecao(oObjStrDados) Then
        pMensagem = oObjClsDados.UltimoErro
        Exit Sub
    End If

    If oObjClsDados.Colecao.Count = 0 Then
        If Not oObjClsDados.gfIncluir(oObjStrDados) Then
            Call gM46V999.gpGraLog(1, oObjClsDados.UltimoErro)
        End If
    End If

    Set oObjStrDados = Nothing
    Set oObjClsDados = Nothing

    Exit Sub
TrataErro:
    gM46V999.gpMenErro ("gpVerificaBoaVista Erro: " & Err & "-" & Error)

End Sub

Public Function gfIsMetaframe() As Boolean
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    Dim ret As Boolean

'        gfIsMetaframe = False
'        Exit Function

    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        gfIsMetaframe = False
        Exit Function
    End If

    SocketsInitialize

    ret = False

    If gethostname(hostname, 256) = SOCKET_ERROR Then
        MsgBox "Windows Sockets error " & Str(WSAGetLastError())
        Exit Function
    Else
        hostname = Trim$(hostname)
    End If

    hostent_addr = gethostbyname(hostname)

    If hostent_addr = 0 Then
        MsgBox "Winsock.dll is not responding."
        Exit Function
    End If

    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4

    '       MsgBox hostname
    
    'get all of the IP address if machine is  multi-homed
    
    Do
        ReDim temp_ip_address(1 To host.hLength)
        RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

        If temp_ip_address(3) = 1 Then
            ret = True
            Exit Do
    End If
    
        For i = 1 To host.hLength
            ip_address = ip_address & temp_ip_address(i) & "."
        Next
        ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

        '           MsgBox ip_address

        ip_address = ""
        host.hAddrList = host.hAddrList + LenB(host.hAddrList)
        RtlMoveMemory hostip_addr, host.hAddrList, 4
    Loop While (hostip_addr <> 0)

    SocketsCleanup

    gfIsMetaframe = ret
End Function

'Public Function gfIsMetaframe() As Boolean
'    Dim lArquivo As Scripting.FileSystemObject
'
'    gfIsMetaframe = False
'
'    Set lArquivo = New Scripting.FileSystemObject
'    If lArquivo.DriveExists("L:") Then
'        gfIsMetaframe = True
'    End If
'
'End Function

Function lobyte(ByVal wParam As Integer)

    lobyte = wParam And &HFF&

End Function


Sub SocketsCleanup()
    Dim lReturn As Long

    lReturn = WSACleanup()

    If lReturn <> 0 Then
        MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
        End
    End If
End Sub

Sub SocketsInitialize()
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String

    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

    If iReturn <> 0 Then
        MsgBox "Winsock.dll is not responding."
        End
    End If

    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
                                                    WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then

        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported by winsock.dll "
        MsgBox sMsg
        End
    End If

    'iMaxSockets is not used in winsock 2. So the following check is only
    'necessary for winsock 1. If winsock 2 is requested,
    'the following check can be skipped.

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox sMsg
        End
    End If

End Sub


Function hibyte(ByVal wParam As Integer)

    hibyte = wParam \ &H100 And &HFF&

End Function

Public Sub gp_mnuArqPesquisa()

'Procedure : Controle de chamada do programa 101.

'    Call gM46V999.gpGraLog(0, "AGE0004 - Opção selecionada: arquivo - pesquisa cotações, propostas, apólices " & _
     '                     "e protocolo.")
    If gM46V999.gfDisponivel("frmT46V101A") = True Then
        frmT46V101A.SetFocus
        Exit Sub
    End If

    If gM46V999.PerfilAcesso <> e_Perfil_CALC And gM46V999.PerfilAcesso <> e_Perfil_CALC_Gestor Then         'Yasuda
        frmT46V012A.Show 1
        If gM46V999.gfDisponivel("frmT46V012A") = True Then
            Unload frmT46V012A
            Call gpP46v100Main
        End If
    Else                        'Corretor ou Cotação
        frmT46V003A.Show 1
        If gM46V999.gfDisponivel("frmT46V003A") = True Then
            Unload frmT46V003A
            Call gpP46v100Main
        End If
    End If
End Sub


'Public Function open_conect_mapi() As Boolean
'
'    Dim Usuario As Variant
'    Dim lOpenConn As Boolean
'
'    On Error Resume Next
'
'    If objSession Is Nothing Then
'        lOpenConn = True
'    Else
'        Usuario = objSession.CurrentUser
'        If Usuario = "Unknown" Or Err.Number <> 0 Then
'            lOpenConn = True
'        End If
'        Usuario = ""
'    End If
'
'    On Error GoTo TrataErro
'
'    If lOpenConn = True Then
'
'        On Error Resume Next
'
'        If Not (objSession Is Nothing) Then
'            objSession.Logoff
'            Set objSession = Nothing
'        End If
'
'        Err.Clear
'        Set objSession = New MAPI.Session
'        If objSession Is Nothing Or Err.Number <> 0 Then
'            Call gM46V999.gpGraLog(0, "[ERR] - Erro ao criar objeto MAPI.Session. Descrição: " & Err.Description)
'            Exit Function
'        End If
'
'        Err.Clear
'
'        objSession.Logon , "", False, True, , True, MAPI_SERVER_EXCHANGE & Chr(10) & MAPI_CONTA_EXCHANGE
'
'        If Err <> 0 Then
'            Call gM46V999.gpGraLog(0, "[ERR] - Não Foi possivel efetuar logon com a MAPI.Logon. Descrição:  " & Err.Description)
'            Exit Function
'        End If
'
'        Err.Clear
'        Usuario = objSession.CurrentUser
'        If Usuario = "Unknown" Then
'            Call gM46V999.gpGraLog(0, "[ERR] - Não Foi possivel efetuar logon com a MAPI.Logon. Descrição:  " & Err.Description)
'            Exit Function
'        End If
'
'        On Error GoTo TrataErro
'
'        Call gM46V999.gpGraLog(0, "[MSG] - Conexao MAPI Inicializada. Usuário: " & Trim(objSession.CurrentUser))
'
'    End If
'
'    open_conect_mapi = True
'    Exit Function
'
'TrataErro:
'
'    Dim ErrDescription
'    ErrDescription = Err.Description
'
'    open_conect_mapi = False
'    Call gM46V999.gpGraLog(0, "[ERR] -  Erro ao efetuar conexção com Exchange... Descrição: " & ErrDescription)
'    On Error GoTo 0: Err.Clear
'
'End Function


'Public Function Close_conect_Mapi() As Boolean
'
'    On Error GoTo TrataErro
'
'    If Not objSession Is Nothing Then
'        objSession.Logoff
'        Set objSession = Nothing
'    End If
'
'    Close_conect_Mapi = True
'
'Exit Function
'TrataErro:
'
'   Call gM46V999.gpGraLog(0, "[ERR] - Falha ao efetuar logoff de MAPI Descrição: " & Err.Description)
'
'End Function
'
Public Sub gpP46v100Main()
    Dim lbdP00MPED As ADODB.Connection    'Base de dados: P0042300.
    Dim lDatModulo As Double      'Data do módulo.
    Dim lDatVersao As String      'Data da versão.
    Dim lLargura As Integer   'Largura.
    Dim lMensagem As String     'Mensagem.
    Dim lPosicao As Integer   'Utilizado com InStr.
    Dim lRegistro As String     'Registro.
    Dim objTabUser As clsA46V725A    'Objeto da classe clsA28V713A (Tab_User)
    Dim objstcTabUser As stcA46V725B    'Objeto da classe de estrutura stcA28V713B (Tab_User)

    Screen.MousePointer = vbHourglass

    '3. Emite mensagem de entrada (precisa ser depois de obter dados do arquivo de inicialização).
    Call gM46V999.gpGraLog(0, "AGE0001 - gpP46v100Main Início do processamento do programa " & App.EXEName & ".")

    '5. Posiciona variável gM46V999.gVerDataMaq (indica se a função que verifica se houve ajuste na máquina
    '   deve ser processada: 0 = não verifica e 1 = verifica).
    If App.EXEName <> "P46V101" Then
        gM46V999.gVerDataMaq = 1
    Else
        gM46V999.gVerDataMaq = 0
    End If

    'Carrego o perfil do usuário de digitação
    gM46V999.VerificarNormas = True

    '6. Carrega perfil do usuário de digitação.
    If gM46V999.gTipUsuario = 1 Then
        Set objTabUser = New clsA46V725A
        Set objTabUser.M46V999 = gM46V999

        Set objstcTabUser = New stcA46V725B
        objstcTabUser.login_user_rede = UCase(gM46V999.gNomUsuario)
        If Not objTabUser.mfCarregarColecao(objstcTabUser) Then
            lMensagem = objTabUser.UltimoErro
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
        If gM46V999.gFinalizarPrograma = True Then End
    Else
        gM46O001.PermiteAlterarTaxaFinal = False
        gM46O001.PermiteQQDescontoComercial = False
        gM46O001.PermiteDigitarSemNormas = False
        gM46O001.PermiteAlterarFranquia = False
    End If

    Call gM46V111.gpObterRegComunicacao(True, "", "")

    '9. Exibe rodapé.
    'a) Obtém data da versão do programa.
    lDatVersao = "20/09/2007"

    '    gNomeTabImport = ""

    '10. Posiciona endereço do arquivo SYASPROD.HLP.
    App.HelpFile = gM46V999.gAppPath & "SYASMULT.HLP"

    '12. Registro do Microfocus
    lMensagem = ""
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        If gfRegistroMicrofocus(lMensagem) = False Then
            Call gM46V999.gpGraLog(0, lMensagem)
        End If
    End If

    '13. Emite janela de proposta.
    'Fecha antes de abrir

    If UCase(frmT46V001A.panCorretor.Caption) <> "YASUDA" Then
        If InStr(1, frmT46V001A.panCorretor.Caption, "Corretor") > 0 Then
            gCodCorrInst = Val(Mid(frmT46V001A.panCorretor.Caption, InStr(1, frmT46V001A.panCorretor.Caption, "Corretor") + 8))
        End If
    End If

    frmT46V101A.Show 0, frmT46V001A
    Screen.MousePointer = vbDefault
End Sub

Public Function RegisterODBCDatabase(dsn As String, driver As String, silent As Integer, attributes As String) As Boolean

    Dim ret As Integer
    Dim temp As String
    Dim spot As Integer

    While InStr(attributes, Chr(13))      'Replace Carriage returns
        spot = InStr(attributes, Chr(13))    'with nulls.
        Mid(attributes, spot, 1) = Chr(0)
    Wend
    attributes = attributes & Chr(0) & Chr(0)    'End of attribute section.
    temp = "DSN=" & dsn & Chr(0) & attributes

    ret = SQLConfigDataSource(0, 4, driver, temp)

    'ret is equal to 1 on success and 0 if there is an error.
    If ret <> 1 Then
        'MsgBox "SQLConfigDataSource call failed - " & ret
        RegisterODBCDatabase = False
    Else
        'MsgBox " SQLConfigDataSource call succeeded!"
        RegisterODBCDatabase = True
    End If

End Function

'Primeira função testada, que cria o ODBC nas fontes de dados de usuário
'não sendo utilizada.
'Public Function RegisterODBCDatabase1(dsn As String, driver As String, silent As Integer, attributes As String) As Boolean

'   Dim ret As Integer
'   On Error GoTo errorhandler

'   RegisterODBCDatabase = False

'   DBEngine.RegisterDatabase dsn, driver, silent, attributes

'   RegisterODBCDatabase = True
'   Exit Function

'errorhandler:
'   If Err = 3146 Or Err = 429 Then    ' ODBC Call Failed.
'      Dim temp As String
'      Dim spot As Integer
'      MsgErr = Err & " - " & Error
'      While InStr(attributes, Chr(13))     ' Replace Carriage returns
'         spot = InStr(attributes, Chr(13)) ' with nulls.
'         Mid(attributes, spot, 1) = Chr(0)
'      Wend
'      attributes = attributes & Chr(0) & Chr(0) ' End of attribute section.
'      temp = "DSN=" & dsn & Chr(0) & attributes

'      ret = SQLConfigDataSource(0, ODBC_ADD_DSN, driver, temp)

' ret is equal to 1 on success and 0 if there is an error.
'      If ret <> 1 Then
'  MsgBox "SQLConfigDataSource call failed - " & ret
'         MsgErr = "SQLConfigDataSource call failed - " & ret & " - " & MsgErr
'      Else
'  MsgBox " SQLConfigDataSource call succeeded!"
'         MsgErr = ""
'         RegisterODBCDatabase = True
'      End If
'   Else
'         MsgErr = Err & " - " & Error
'   End If
'   Exit Function

'End Function




