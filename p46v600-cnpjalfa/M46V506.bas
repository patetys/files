Attribute VB_Name = "basM46V506"
Option Explicit
Public Cronometros(10) As Double
Public Crono_Stop(10) As Double
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
                                      (ByVal nDrive As String) As Long
Private Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" _
                                              (ByVal lpszName As String, ByVal bForce As Long) As Long
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" _
                                            (lpNetResource As NETRESOURCE, _
                                             ByVal lpPassword As String, _
                                             ByVal lpUserName As String, _
                                             ByVal dwFlags As Long) As Long
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" _
                                          (lpError As Long, ByVal lpErrorBuf As String, _
                                           ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, _
                                           ByVal nNameBufSize As Long) As Long

Public Const WN_SUCCESS = &H0                 ' Function was successful.
Public Const WN_NOT_SUPPORTED = &H1           ' Function was not supported.
Public Const WN_OUT_OF_MEMORY = &HB           ' System was out of memory.
Public Const WN_NET_ERROR = &H2               ' An error occurred on the network.
Public Const WN_BAD_POINTER = &H4             ' Pointer was invalid.
Public Const WN_BAD_NETNAME = &H32            ' Network resource name was invalid.
Public Const WN_BAD_LOCALNAME = &H33          ' Local device name was invalid.
Public Const WN_BAD_PASSWORD = &H6            ' Password was invalid.
Public Const WN_ACCESS_DENIED = &H7           ' A security violation occurred.
Public Const WN_ALREADY_CONNECTED = &H34      ' Local device was already connected to a remote resource.

Private Declare Function WaitForSingleObject Lib "kernel32" _
                                             (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" _
                                        (ByVal lpApplicationName As Long, _
                                         ByVal lpCommandLine As String, _
                                         ByVal lpProcessAttributes As Long, _
                                         ByVal lpThreadAttributes As Long, _
                                         ByVal bInheritHandles As Long, _
                                         ByVal dwCreationFlags As Long, _
                                         ByVal lpEnvironment As Long, _
                                         ByVal lpCurrentDirectory As Long, _
                                         lpStartupInfo As STARTUPINFO, _
                                         lpProcessInformation As _
                                         PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
                                            (ByVal lpBuffer As String, ByVal nsize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
                                           (ByVal lpBuffer As String, ByVal nsize As Long) As Long

Public Const Ligado = -1
Public Const Desligado = 0
Public Function PegaParam(pOrdem As Long) As String
'Procedure: pega parâmetro do comando utilizado para acionar o programa.

'Parâmetro de entrada...pOrdem...Ordem do parâmetro.

    Dim lComando As String    'Comando utilizado para acionar o programa.
    Dim lFaz As Boolean    'Indica que concluiu o processo.
    Dim lI As Long             'Índice.
    Dim lParametro(15) As String  'Parâmetro.
    Dim lPosicao As Long    'Utilizado com InStr.

    lComando = Command()
    lI = 1
    lFaz = True
    Do
        lPosicao = InStr(lI, lComando, " ")
        If lPosicao <> 0 Then
            lParametro(lI) = Mid(lComando, 1, lPosicao - 1)
            lComando = Mid(lComando, lPosicao + 1)
            lI = lI + 1
        Else
            lParametro(lI) = lComando
            lComando = ""
            lFaz = False
        End If
    Loop Until (Not lFaz)
    If pOrdem <= lI Then
        PegaParam = lParametro(pOrdem)
    Else
        PegaParam = ""
    End If
End Function
Public Function sPegaCampo(ByRef pRegistro As String, Optional pDelimitador As String = ";") As String
'Função: pega campo.

'Parâmetro de entrada/saída...pRegistro...Registro.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lPosicao As Long  'Utilizado com InStr.

    lAuxiliar = pRegistro
    lPosicao = InStr(lAuxiliar, pDelimitador)
    If lPosicao <> 0 Then
        sPegaCampo = LTrim(RTrim(Mid(lAuxiliar, 1, lPosicao - 1)))
        lAuxiliar = Mid(lAuxiliar, lPosicao + 1)
    Else
        sPegaCampo = LTrim(RTrim(lAuxiliar))
        lAuxiliar = ""
    End If
    pRegistro = lAuxiliar
End Function
'Sub DELAY(PauseTime As Long)
'    Dim lInicio     'Início.
'    Dim lTermino    'Término.
'
'    lInicio = Timer     'Set lInicio time.
'    Do While Timer < (lInicio + (PauseTime / 10))
'        DoEvents        'Yield to other processes.
'    Loop
'    lTermino = Timer    'Set end time.
'End Sub
'Sub Mouse(pTipo As Long)
'    If pTipo = Ligado Then
'        Screen.MousePointer = vbHourglass
'    Else
'        Screen.MousePointer = vbNormal
'    End If
'End Sub
'Sub ZeraCrono(QualCronometro As Integer)
'    Cronometros(QualCronometro) = Timer
'    Crono_Stop(QualCronometro) = 0
'End Sub
'Function Crono(QualCronometro As Integer) As Double
'    If Crono_Stop(QualCronometro) = 0 Then
'        Crono = Timer - Cronometros(QualCronometro)
'    Else
'        Crono = Crono_Stop(QualCronometro) - Cronometros(QualCronometro)
'    End If
'End Function
'Sub CronoStop(QualCronometro As Integer)
'    Crono_Stop(QualCronometro) = Timer
'End Sub
'Sub CenterForm(formulario As Form, Optional dl = 0, Optional dt = 0)
'    formulario.Left = (Screen.Width - formulario.Width) / 2 + dl
'    formulario.Top = (Screen.Height - formulario.Height) / 2 + dt
'End Sub
Public Function Confirma(ByVal pMensagem As String, Optional pOpcao As String = "S") As Boolean
'Função: emite mensagem de confirmação.

    Select Case pOpcao
    Case "S"
        If MsgBox(pMensagem, vbYesNo, "Confirmação") = vbYes Then
            Confirma = True
            Call gM46V999.gpGraLog(0, pMensagem & " SIM.")
        Else
            Confirma = False
            Call gM46V999.gpGraLog(0, pMensagem & " NÃO.")
        End If
    Case Else
        If MsgBox(pMensagem, vbYesNo + vbDefaultButton2, "Confirmação") = vbYes Then
            Confirma = True
            Call gM46V999.gpGraLog(0, pMensagem & " SIM.")
        Else
            Confirma = False
            Call gM46V999.gpGraLog(0, pMensagem & " NÃO.")
        End If
    End Select
End Function
Public Sub ExecCmd(pComando As String, Optional pEspera As Integer = 0)
'Procedure: executa comando.

'Parâmetros de entrada...pComando...Linha de comando.
'                        pEspera....Indicador de espera.

    Dim lProc As PROCESS_INFORMATION
    Dim lretorno As Long
    Dim lStart As STARTUPINFO

    '1. Inicializa a estrutura STARTUPINFO.
    lStart.cb = Len(lStart)

    '2. Inicia a execução da aplicação.
    lretorno = CreateProcessA(0&, pComando$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, lStart, lProc)
    If pEspera = 1 Then
        'Espera a aplicação encerrar.
        lretorno = WaitForSingleObject(lProc.hProcess, INFINITE)
        lretorno = CloseHandle(lProc.hProcess)
    End If
End Sub
'Function Limpa_String_Nao_Pode2(ByRef sNome As String, ByRef sOqueNaoPode As String, Optional PoeBranco As Boolean = False) As String
'    Dim lAuxiliar   As String   'Auxiliar.
'    Dim lPosicao    As Long     'Utilizado com InStr.
'
'    lAuxiliar = sNome
'    lPosicao = InStr(lAuxiliar, sOqueNaoPode)
'    Do While lPosicao <> 0
'        lAuxiliar = Mid(lAuxiliar, 1, lPosicao - 1)
'        If PoeBranco Then
'            lAuxiliar = lAuxiliar & " "
'        End If
'        lAuxiliar = lAuxiliar & Mid(sNome, lPosicao + Len(sOqueNaoPode))
'        lPosicao = InStr(lAuxiliar, sOqueNaoPode)
'        sNome = lAuxiliar
'    Loop
'    Limpa_String_Nao_Pode2 = lAuxiliar
'End Function
'Function Troca_String(ByRef sString As String, ByRef sOqueTrocar As String, sTrocarPor As String) As String
'    Dim lAuxiliar   As String   'Auxiliar.
'    Dim lAuxiliar0  As String   'Auxiliar 0.
'    Dim lPosicao    As Long     'Utilizado com InStr.
'
'    lAuxiliar = sString
'    lPosicao = InStr(lAuxiliar, sOqueTrocar)
'    Do While lPosicao <> 0
'        lAuxiliar0 = lAuxiliar
'        lAuxiliar = Mid(lAuxiliar, 1, lPosicao - 1)
'        lAuxiliar = lAuxiliar & sTrocarPor & Mid(lAuxiliar0, lPosicao + Len(sOqueTrocar))
'        lPosicao = InStr(lAuxiliar, sOqueTrocar)
'    Loop
'    Troca_String = lAuxiliar
'End Function
'Function Limpa_String_Pode(ByRef sNome As String, ByRef sOquePode As String, Optional PoeBranco As Boolean = False) As String
'    Dim lAuxiliar   As String   'Auxiliar.
'    Dim lPosicao    As Long     'Utilizado com InStr.
'
'    If IsNull(sNome) Then
'        lAuxiliar = ""
'    End If
'    lAuxiliar = ""
'    For lPosicao = 1 To Len(sNome)
'        If InStr(sOquePode, Mid(sNome, lPosicao, 1)) <> 0 Then
'            lAuxiliar = lAuxiliar & Mid(sNome, lPosicao, 1)
'        Else
'            If PoeBranco Then
'                lAuxiliar = lAuxiliar & " "
'            End If
'        End If
'    Next
'    Limpa_String_Pode = lAuxiliar
'End Function
'Function Limpa_String_Nao_Pode(ByRef sNome As String, ByRef sOqueNaoPode As String, Optional PoeBranco As Boolean = False) As String
'    Dim lAuxiliar   As String   'Auxiliar.
'    Dim lPosicao    As Long     'Utilizado com InStr.
'
'    lAuxiliar = ""
'    For lPosicao = 1 To Len(sNome)
'        If InStr(sOqueNaoPode, Mid(sNome, lPosicao, 1)) = 0 Then
'            lAuxiliar = lAuxiliar & Mid(sNome, lPosicao, 1)
'        Else
'            If PoeBranco Then
'                lAuxiliar = lAuxiliar & " "
'            End If
'        End If
'    Next
'    Limpa_String_Nao_Pode = lAuxiliar
'End Function
Public Function GetWindowsPath() As String
    Dim lAuxiliar As String    'Auxiliar.
    Dim lI As Long    'Índice.
    Dim lLong As Long    'Auxiliar.

    lAuxiliar = Space(255)
    lLong = 254
    GetWindowsPath = ""
    GetWindowsDirectory lAuxiliar, lLong
    For lI = 1 To 255
        If Asc(Mid(lAuxiliar, lI, 1)) <> 0 Then
            GetWindowsPath = GetWindowsPath & Mid(lAuxiliar, lI, 1)
        Else
            Exit For
        End If
    Next lI
    GetWindowsPath = GetWindowsPath & "\"
End Function
Public Function GetWindowsSystemPath() As String
    Dim lAuxiliar As String    'Auxiliar.
    Dim lI As Long    'Índice.
    Dim lLong As Long    'Auxiliar.

    lAuxiliar = Space(255)
    lLong = 254
    GetWindowsSystemPath = ""
    GetSystemDirectory lAuxiliar, lLong
    For lI = 1 To 255
        If Asc(Mid(lAuxiliar, lI, 1)) <> 0 Then
            GetWindowsSystemPath = GetWindowsSystemPath & Mid(lAuxiliar, lI, 1)
        Else
            Exit For
        End If
    Next
    GetWindowsSystemPath = GetWindowsSystemPath & "\"
End Function
Public Function Registra_DLL(pComponente As String) As Boolean
    Dim lEndREGSVR32 As String    'Endereço do programa REGSVR32.EXE.

    lEndREGSVR32 = GetWindowsSystemPath() & "REGSVR32.EXE"

    If Not gM46V999.gfPreenchido(Dir(lEndREGSVR32)) Then
        Registra_DLL = False
        Exit Function
    End If
    Registra_DLL = True
    ExecCmd lEndREGSVR32 & " " & pComponente & " /s ", 1
End Function
Public Function DesRegistra_DLL(pComponente As String) As Boolean
    Dim lEndREGSVR32 As String    'Endereço do programa REGSVR32.EXE.

    lEndREGSVR32 = GetWindowsSystemPath() & "REGSVR32.EXE"
    If Not gM46V999.gfPreenchido(Dir(lEndREGSVR32)) Then
        DesRegistra_DLL = False
        Exit Function
    End If
    DesRegistra_DLL = True
    ExecCmd lEndREGSVR32 & " " & pComponente & " /u /s ", 1
End Function
