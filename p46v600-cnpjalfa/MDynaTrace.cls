VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMDynaTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function SetEnvironmentVariable Lib "kernel32.dll" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32.dll" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nsize As Long) As String

Private Declare Function VB_simpleTest Lib "dtAgentWrapper.dll" () As Long
Private Declare Function VB_dTInitialize Lib "dtAgentWrapper.dll" (ByVal inManagerRef As Long, ByVal outManagerRef As String) As Long
Private Declare Function VB_dTUnInitialize Lib "dtAgentWrapper.dll" () As Long
Private Declare Function VB_dTEnter Lib "dtAgentWrapper.dll" (ByVal method_id As Long, ByVal serial_no As Long) As Long
Private Declare Function VB_dTExit Lib "dtAgentWrapper.dll" (ByVal method_id As Long, ByVal serial_no As Long) As Long
Private Declare Function VB_dTGetMethodID Lib "dtAgentWrapper.dll" (ByVal method_name As String, ByVal sourcefile_name As String, ByVal source_line_no As Long, ByVal api As String, ByVal is_artificial As String) As Long
Private Declare Function VB_dTGetSerialNo Lib "dtAgentWrapper.dll" (ByVal method_id As Long, ByVal entry_point As Long) As Long
Private Declare Function VB_dTCaptureString Lib "dtAgentWrapper.dll" (ByVal serial_no As Long, ByVal argument As String) As Long
Private Declare Function VB_dTCustomLog Lib "dtAgentWrapper.dll" (ByVal method_id As Long, ByVal logmessage As String, ByVal level As Long, ByVal loggername As String) As Long
Private Declare Function VB_dTExceptionThrow Lib "dtAgentWrapper.dll" (ByVal method_id As Long, ByVal message As String, ByVal classname As String) As Long
Private Declare Function VB_dTEndServerPurepath Lib "dtAgentWrapper.dll" () As Long
Private Declare Function VB_dTSetTagFromString Lib "dtAgentWrapper.dll" (ByVal tag As String) As Long
Private Declare Function VB_dTStartServerPurepath Lib "dtAgentWrapper.dll" () As Long
Private Declare Function VB_dTExitException Lib "dtAgentWrapper.dll" (ByVal method_id As Long, ByVal serial_no As Long, ByVal message As String) As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Enum EP_DT   'Entry Point DynaTrace
    epMetodoPublico = 1
    epMetodoInterno = 0
End Enum

Public Enum eSeveridadeErro
    eSeveridadeAlerta = 0
    eSeveridadeErro_Baixo = 1
    eSeveridadeErro_Medio = 2
    eSeveridadeErro_Alto = 3
End Enum

Private dt_method_id_publico    As Long
Private dt_serial_no_publico    As Long
Private nomeAgente              As String
Private AtivarInstrumentacao    As Boolean

Public Type metodoFilho
    Key             As Long
    method_id       As Long
    serial_no       As Long
    nomeMetodo      As String
    nomeObjeto      As String
    Key_Pai         As Long
End Type

Private colMetodoFilho  As Collection

Private Sub Class_Initialize()
 'Inicializa Variaveis
    Call inicializarVariaveis
End Sub

Private Sub Class_Terminate()
Dim rv As Long
    'rv = VB_dTUnInitialize
    Call inicializarVariaveis
End Sub

Public Sub Iniciar(ByVal pNomeAgente As String)
'DynaTrace variaveis
    Dim rv                 As Long
    Dim argC                As Long
    Dim argV                As Long
'====================
    rv = SetEnvironmentVariable("DT_AGENTNAME", pNomeAgente)
    nomeAgente = pNomeAgente
    rv = SetEnvironmentVariable("DT_AGENTACTIVE", agenteAtivo) 'Liga ou desliga o DinaTrace
    rv = SetEnvironmentVariable("DT_WAIT", timeOut) 'TimeOut do DynaTarce em Segundos
    rv = VB_dTInitialize(argC, argV)
    AtivarInstrumentacao = CBool(agenteAtivo)
    
End Sub

Public Function iniciarMetodoPublico(ByVal tag As String, ByVal nomeMetodo As String, ByVal nomeObjeto As String) As Long
Dim rv                  As Long
Dim lKeyMetodoPublico   As Long
Dim lMsgAlertaInterno   As String
        
    If tag = "" Then
        AtivarInstrumentacao = False
        'Call Err.Raise(44777, "DynaTrace", "Método " & nomeMetodo & " enviou a tag vazia. Verificar.")
    End If
    
    If AtivarInstrumentacao = False Then
        Exit Function
    End If
    
    If nomeAgente = "" Then
        Call Err.Raise(11120, "DynaTrace", "Nome do agente não especificado. Isso indica que foi invocado sem o metodo Iniciar. O método invocado foi " & nomeMetodo & ".  Verificar.")
    End If
        
    If Atual_method_id_Publico = 0 Then
        '===================================
        'DynaTrace - Metodo Publico
        '===================================
        '1- TAG
        rv = VB_dTSetTagFromString(tag)
        '2- PurePath
        rv = VB_dTStartServerPurepath
        '3- Atribui MethodID publico Atual
        dt_method_id_publico = VB_dTGetMethodID(nomeMetodo, nomeObjeto, 0, nomeObjeto, "TestIsArtificial")
        '4- Atribui Serial Number publico Atual
        dt_serial_no_publico = VB_dTGetSerialNo(dt_method_id_publico, epMetodoPublico)
        '5- ENTER (Inicaliza o Metodo Publico)
        rv = VB_dTEnter(Atual_method_id_Publico, Atual_serial_no_Publico)
        '===================================
        lKeyMetodoPublico = 0
        lMsgAlertaInterno = "Metodo publico " & nomeMetodo & " (Tag:" & tag & ")" ' & PegarInformacoesProcesso
        rv = VB_dTCustomLog(dt_method_id_publico, lMsgAlertaInterno, CLng(eSeveridadeAlerta), "")
    
        '===================================
        'Arvore interna de metotodos
        '===================================
        Dim lMetodoFilho    As metodoFilho
        With lMetodoFilho
            .Key = CStr(CLng(MetodoFilho_LastKey) + 1)
            .Key_Pai = 0
            .method_id = dt_method_id_publico
            .serial_no = dt_serial_no_publico
            .nomeMetodo = nomeMetodo
            .nomeObjeto = nomeObjeto
        End With
        Call mpAdicMetodoFilho(lMetodoFilho)
        '===================================
        lKeyMetodoPublico = lMetodoFilho.Key
    Else
        'Metodo publico iniciado internamente
        lKeyMetodoPublico = iniciarMetodoInterno(nomeMetodo, nomeObjeto)
        
        'lMsgAlertaInterno = "Metodo publico chamado internamente. Tag Ignorada, aceito como metodo interno . (Tag:" & tag & ")"
        'rv = VB_dTCustomLog(colMetodoFilho.Item(MetodoFilho_LastKey).method_id, lMsgAlertaInterno, CLng(eSeveridadeErro_Baixo), "")
    End If
    
    iniciarMetodoPublico = lKeyMetodoPublico
    
End Function

Public Sub finalizarMetodoPublico(ByVal pKeyMetodoPublico As Long, ByVal execSucesso As Boolean, Optional pMsgErr As String = "", Optional severidadeErro As eSeveridadeErro = eSeveridadeErro_Baixo, Optional nomeObjetoFechado As String = "", Optional nomeMetodoFechado As String = "")
Dim rv                 As Long
 
    If AtivarInstrumentacao = False Then
        Exit Sub
    End If
    
    If pKeyMetodoPublico <> 1 Then
    'Metodo publico iniciado internamente
        Call finalizarMetodoInterno(pKeyMetodoPublico, execSucesso, pMsgErr, severidadeErro, nomeObjetoFechado, nomeMetodoFechado)
    Else
        '===================================
        'Arvore interna de metotodos
        '===================================
        Call mpRemoveMetodoFilho(pKeyMetodoPublico, nomeObjetoFechado, nomeMetodoFechado)
        '===================================
        'rv = VB_dTCustomLog(Atual_method_id_Publico, PegarInformacoesProcesso, 0, "")
        If execSucesso = True Then
            rv = VB_dTCustomLog(Atual_method_id_Publico, "Metodo publico" & IIf(nomeMetodoFechado <> "", " " & nomeMetodoFechado, "") & " finalizado com sucesso.", 0, nomeAgente)
            rv = VB_dTExit(Atual_method_id_Publico, Atual_serial_no_Publico)
        Else
            rv = VB_dTCustomLog(Atual_method_id_Publico, "Metodo publico finalizado com falha: " & pMsgErr, severidadeErro, nomeAgente)
            rv = VB_dTExitException(Atual_method_id_Publico, Atual_serial_no_Publico, pMsgErr)
        End If
        rv = VB_dTEndServerPurepath
        dt_method_id_publico = 0
    End If
 
End Sub

Private Function iniciarMetodoInterno(ByVal nomeMetodo As String, ByVal nomeObjeto As String) As Long
'DynaTrace variaveis
        Dim rv                      As Long
        Dim dt_method_id_Interno    As Long
        Dim dt_serial_no_Interno    As Long
'===================================

'===================================
'DynaTrace - Metodo Privado
'===================================
'1- Atribui MethodID privado Atual
    dt_method_id_Interno = VB_dTGetMethodID(nomeMetodo, nomeObjeto, 0, nomeObjeto, "TestIsArtificial")
'2- Atribui Serial Number publico Atual
    dt_serial_no_Interno = VB_dTGetSerialNo(dt_method_id_Interno, epMetodoInterno)
'3- ENTER (Inicaliza o Metodo Publico)
    rv = VB_dTEnter(dt_method_id_Interno, dt_serial_no_Interno)
    rv = VB_dTCustomLog(dt_method_id_Interno, "Entrou no metodo " & nomeMetodo & " da classe " & nomeObjeto, 0, nomeAgente)
'===================================

'===================================
'Arvore interna de metotodos
'===================================
    Dim lMetodoFilho    As metodoFilho
    With lMetodoFilho
        .Key = CStr(CLng(MetodoFilho_LastKey) + 1)
        .Key_Pai = MetodoFilho_LastKey
        .method_id = dt_method_id_Interno
        .serial_no = dt_serial_no_Interno
        .nomeMetodo = nomeMetodo
        .nomeObjeto = nomeObjeto
    End With
    Call mpAdicMetodoFilho(lMetodoFilho)
'===================================
    iniciarMetodoInterno = lMetodoFilho.Key
End Function

Private Sub finalizarMetodoInterno(ByVal pKeyMetodoInterno As Long, ByVal execSucesso As Boolean, Optional pMsgErr As String = "", Optional severidadeErro As eSeveridadeErro = eSeveridadeErro_Baixo, Optional nomeObjetoFechado As String = "", Optional nomeMetodoFechado As String = "")
Dim rv                  As Long
Dim lMetodoInterno      As metodoFilho
Dim lMsgFalhaInterna    As String

    If ExiteFilho(pKeyMetodoInterno) = True Then
        '===================================
        'Arvore interna de metotodos
        '===================================
        lMetodoInterno = colMetodoFilho.Item(CStr(pKeyMetodoInterno))
        Call mpRemoveMetodoFilho(pKeyMetodoInterno, nomeObjetoFechado, nomeMetodoFechado)
        '===================================
        'rv = VB_dTCustomLog(lMetodoInterno.method_id, PegarInformacoesProcesso, 0, "")
        If execSucesso = True Then
            rv = VB_dTCustomLog(lMetodoInterno.method_id, "Saiu do metodo " & lMetodoInterno.nomeMetodo & " da classe " & lMetodoInterno.nomeObjeto, 0, nomeAgente)
            rv = VB_dTExit(lMetodoInterno.method_id, lMetodoInterno.serial_no)
        Else
            rv = VB_dTCustomLog(lMetodoInterno.method_id, "Saiu com falha do metodo " & lMetodoInterno.nomeMetodo & " da classe " & lMetodoInterno.nomeObjeto & ": " & pMsgErr, 0, nomeAgente)
            'rv = VB_dTCustomLog(lMetodoInterno.method_id, pMsgErr, severidadeErro, nomeAgente)
            rv = VB_dTExitException(lMetodoInterno.method_id, lMetodoInterno.serial_no, pMsgErr)
        End If
    Else
        lMsgFalhaInterna = "Nao encontrado metodo na pilha para finalizar. (KeyInterno:" & pKeyMetodoInterno & ")"
        rv = VB_dTCustomLog(Atual_method_id_Publico, lMsgFalhaInterna, CLng(eSeveridadeErro_Baixo), "")
    End If

End Sub

Public Sub gravaLog(ByVal pKeyMetodo As Long, ByVal pMsg As String, ByRef pSeveridade As eSeveridadeErro, Optional pIncluirMemoria As Boolean = False)
Dim rv                  As Long
Dim lMethod_id          As Long

    If AtivarInstrumentacao = False Then
        Exit Sub
    End If
    
    If pKeyMetodo = 0 Then
        lMethod_id = Atual_method_id_Publico
    Else
        If ExiteFilho(pKeyMetodo) = True Then
            lMethod_id = colMetodoFilho.Item(CStr(pKeyMetodo)).method_id
        Else
            lMethod_id = Atual_method_id_Publico
        End If
    End If
    
    If pIncluirMemoria = True Then
        pMsg = pMsg & " | " & PegarInformacoesProcesso
    End If
    
    rv = VB_dTCustomLog(lMethod_id, pMsg, CLng(pSeveridade), "")
    
End Sub

Public Property Get PilhaMetodos() As Collection
    Set PilhaMetodos = colMetodoFilho
End Property

Private Property Get Atual_method_id_Publico() As Long
    Atual_method_id_Publico = dt_method_id_publico
End Property

Private Property Get Atual_serial_no_Publico() As Long
    Atual_serial_no_Publico = dt_serial_no_publico
End Property

Private Property Get agenteAtivo() As String
'Liga ou desliga o DinaTrace
'Obter do .INI
    agenteAtivo = "false"
End Property

Private Property Get timeOut() As String
'TimeOut do DynaTarce em Segundos
'Obter do .INI
    timeOut = "5"
End Property

Private Sub mpAdicMetodoFilho(ByRef pMetodoFilho As metodoFilho)
On Error Resume Next
    colMetodoFilho.Add pMetodoFilho, CStr(pMetodoFilho.Key)
On Error GoTo 0
End Sub

Private Sub mpRemoveMetodoFilho(ByVal pKey As Long, Optional nomeObjetoFechado As String = "", Optional nomeMetodoFechado As String = "")
Dim i                   As Long
Dim lMsgLogCorrompido   As String
Dim lMetodoFilho        As metodoFilho
Dim rv                  As Long
Dim lLastKey            As Long
 
    lLastKey = CLng(MetodoFilho_LastKey)
 
    If pKey <> lLastKey Then
    'Método sendo encerrado com filhos abertos....
        lMsgLogCorrompido = "<Descricao = 'Metodos filhos nao encerrados'>"
        For i = 0 To colMetodoFilho.Count
            If ExiteFilho(i) Then
                lMetodoFilho = colMetodoFilho.Item(CStr(i))
                With lMetodoFilho
                    lMsgLogCorrompido = lMsgLogCorrompido & "<Item"
                    lMsgLogCorrompido = lMsgLogCorrompido & " key = '" & .Key & "'"
                    lMsgLogCorrompido = lMsgLogCorrompido & " method_id = '" & .method_id & "'"
                    lMsgLogCorrompido = lMsgLogCorrompido & " serial_no = '" & .serial_no & "'"
                    lMsgLogCorrompido = lMsgLogCorrompido & " nomeMetodo = '" & .nomeMetodo & "'"
                    lMsgLogCorrompido = lMsgLogCorrompido & " nomeObjeto = '" & .nomeObjeto & "'"
                    lMsgLogCorrompido = lMsgLogCorrompido & " Key_Pai = '" & .Key_Pai & "'"
                    lMsgLogCorrompido = lMsgLogCorrompido & "/>" & vbCrLf
                    colMetodoFilho.Remove CStr(.Key)
                End With
            End If
        Next i
        lMsgLogCorrompido = lMsgLogCorrompido & "</Descricao>"
        If pKey > 0 And ExiteFilho(pKey) Then
            rv = VB_dTCustomLog(colMetodoFilho.Item(CStr(pKey)).method_id, lMsgLogCorrompido, CLng(eSeveridadeErro_Medio), "")
        Else
            rv = VB_dTCustomLog(Atual_method_id_Publico, lMsgLogCorrompido, CLng(eSeveridadeErro_Medio), "")
        End If
    End If
   
On Error Resume Next
    If nomeMetodoFechado <> "" Then
        lMetodoFilho = colMetodoFilho.Item(CStr(pKey))
        If (lMetodoFilho.nomeMetodo <> nomeMetodoFechado Or lMetodoFilho.nomeObjeto <> nomeObjetoFechado) Then
            lMsgLogCorrompido = "Erro na instrumentação! Método " & nomeMetodoFechado & " está sendo finalizado mas o método finalizado na verdade será " & lMetodoFilho.nomeMetodo
            If pKey > 0 And ExiteFilho(pKey) Then
                rv = VB_dTCustomLog(colMetodoFilho.Item(CStr(pKey)).method_id, lMsgLogCorrompido, CLng(eSeveridadeErro_Medio), "")
            Else
                rv = VB_dTCustomLog(Atual_method_id_Publico, lMsgLogCorrompido, CLng(eSeveridadeErro_Medio), "")
            End If
        End If
    End If
    If pKey > 0 Then
        colMetodoFilho.Remove CStr(pKey)
    End If
On Error GoTo 0
End Sub

Private Property Get MetodoFilho_LastKey() As String
Dim lretorno    As String
    
    If colMetodoFilho.Count > 0 Then
        lretorno = CStr(colMetodoFilho.Item(colMetodoFilho.Count).Key)
    Else
        lretorno = "0"
    End If

    MetodoFilho_LastKey = lretorno

End Property
 
Private Function ExiteFilho(ByVal pKey As Long) As Boolean
Dim lretorno    As Boolean
Dim Dummy As metodoFilho
On Error GoTo TrataErro
 
    Dummy = colMetodoFilho.Item(CStr(pKey))
    lretorno = True
SAIDA:
    Err.Clear
    ExiteFilho = lretorno
    Exit Function
TrataErro:
    lretorno = False
    GoTo SAIDA
End Function

Private Sub inicializarVariaveis()
    dt_method_id_publico = 0
    dt_serial_no_publico = 0
    Set colMetodoFilho = New Collection
End Sub

Private Function GetProcessMemory(ByVal PID As String) As String
    Dim Process As Object
    Dim dMemory As Double
    For Each Process In GetObject("winmgmts:").ExecQuery("Select WorkingSetSize from Win32_Process Where ProcessId = " & PID & "")
        dMemory = Process.WorkingSetSize
    Next
    If dMemory > 0 Then
        GetProcessMemory = ResizeKb(dMemory)
    Else
        GetProcessMemory = "0 Bytes"
    End If
End Function
 
Private Function ResizeKb(ByVal B As Double) As String
    Dim bSize(8) As String, i As Integer
    bSize(0) = "Bytes"
    bSize(1) = "KB" 'Kilobytes
    bSize(2) = "MB" 'Megabytes
    bSize(3) = "GB" 'Gigabytes
    bSize(4) = "TB" 'Terabytes
    bSize(5) = "PB" 'Petabytes
    bSize(6) = "EB" 'Exabytes
    bSize(7) = "ZB" 'Zettabytes
    bSize(8) = "YB" 'Yottabytes
    For i = UBound(bSize) To 0 Step -1
        If B >= (1024 ^ i) Then
            ResizeKb = ThreeNonZeroDigits(B / (1024 ^ _
                i)) & " " & bSize(i)
            Exit For
        End If
    Next
End Function
 
Private Function ThreeNonZeroDigits(ByVal Value As Double) As Double
    If Value >= 100 Then
        ThreeNonZeroDigits = FormatNumber(Value)
    ElseIf Value >= 10 Then
        ThreeNonZeroDigits = FormatNumber(Value, 1)
    Else
        ThreeNonZeroDigits = FormatNumber(Value, 2)
    End If
End Function

Public Function PegarInformacoesProcesso() As String
    Dim PID As Long
    Dim Memoria As String
    
    PID = GetCurrentProcessId
    Memoria = GetProcessMemory(PID)
    
    PegarInformacoesProcesso = "ProcessID " & PID & " - " & Memoria
End Function
