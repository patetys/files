Attribute VB_Name = "basM47V701"
Option Explicit
Public gCarga               As Boolean          'Indicador de carga de registro.
Public gGraLog              As String           'Identificador de gravação de log
Public gRst_RegrasPlano         As ADODB.Recordset
Public gAltComum            As Boolean          'Indica que houve alteração nos dados de comum.
Public gNomCompleto         As String           'Nome do usuário.
Public gNomComputador       As String           'Nome do computador.
Public gNomUsuario8         As String           'Nome do usuário (8 caracteres).
Public gPWD                 As String           'Senha SQL Server.
Public gServer              As String           'Servidor do SQL Server.
Public gTipo_bd             As String           'Tipo de banco de dados.
Public gUID                 As String           'UID do banco de dados.
Public gDataBase            As String           'Base de dados do transporte.
Public gConexaoSql          As ADODB.Connection
Public gTemErro_T47V025A    As Boolean          'Indica que tem erro na montagem do form T47V025A.
Public gTemErro_T47V026A    As Boolean          'Indica que tem erro na montagem do form T47V026A.
Public gImpressoraEmissao   As String
Public gMetaframe           As Boolean
Public gIndCosseguro        As Byte             'Índice do cosseguro.
Public gCod_RamoCorr         As e_CodRamoTip   'Códigos do ramo de Automóvel para Auto (310), RC (530)  e APP (812). Controle de corretores
Public Enum e_CodRamoTip
    e_CodRamoAuto = 310
    e_CodRamoRC = 530
    e_CodRamoapp = 812
End Enum

Public objSession As MAPI.Session
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

Private mConexaoSQL         As ADODB.Connection 'Conexão.

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, _
ByVal nsize As Long, ByVal lpFileName As String) As Long


Private Declare Function WritePrivateProfileString& Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal filename$)

Public gPathArquivos    As String       'Local onde estão os arquivos de layout.
Public gImprFrente      As String
Public gImprFrenteVerso As String
'Public gPrograma        As String       'Nome do programa em execução.
Public gProTabela       As String       'Proprietário da tabela.

Public gUsuMulti        As Boolean      'Indica se o usuário é somente do sistema Multi ou é do Syas
                                        'também: TRUE = ambos sistemas e FALSE = apenas MULTISSEGUROS.
Public gVerDataMaq      As Byte         '0 - Não ver inconsistência de datas



Public gProChamado      As String
Public gNomCorretor     As String
Public gCodCorretor     As Long
Private mNomForm        As String
Private mEndArq         As String
Private lListaArq       As String
Private lTotPagina      As Integer

Public gCod_Registro As Integer

Public gNomCorrProlabAgenc     As String
Public gCodCorrProlabAgenc     As Long

Public Function CDblx(ByRef Valor As Variant) As Double
    If IsNumeric(Valor) Then
        Valor = Valor
    Else
        Valor = 0
    End If
    CDblx = Valor
End Function

Public Function gfCarregaPerIOF(pDataIniVig As Long, pMensagem As String) As Boolean
Dim bdP00MULTI As ADODB.Connection
Dim lSelect     As String
Dim lrst        As clsYasRecordSet


    If Not gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, bdP00MULTI, pMensagem) Then
        Exit Function
    End If

    lSelect = " Select * from TBM_IOF where Dat_Ini_Vig <= " & pDataIniVig
    lSelect = lSelect & " and Dat_ter_Vig > " & pDataIniVig
    lSelect = lSelect & " and cod_ramo = " & gM46V999.gCod_Ramo
    
    If Not gM46V999.gfObtRegistro(bdP00MULTI, lSelect, lrst, pMensagem) Then
        gM46V111.gProJurIOF = 7
        gfCarregaPerIOF = True
        Exit Function
    End If
    
    If Not lrst.EOF Then
        gM46V111.gProJurIOF = lrst("Per_IOF")
    Else
        gM46V111.gProJurIOF = 7
    End If

    gfCarregaPerIOF = True
    
End Function

Public Sub gpHabSalvar()
    'Procedure: habilita opções salvar e desfazer.

    If gCarga = False Then
        With frmT47V001A
            
            .mnuArqImprimir.Enabled = True
            .mnuArqVisualizar.Enabled = True

            .mnuArqSalvar.Enabled = True
            .barFerramentas.Buttons(4).Enabled = True    'Salvar.
            
            .barFerramentas.Buttons(6).Enabled = True   'Imprimir.
            .barFerramentas.Buttons(7).Enabled = True   'Visualizar.
            
        End With
    End If
End Sub

Sub Main()
Dim lMensagem       As String       'Mensagem.
Dim lAuxiliar       As String       'Auxiliar.
Dim lEndSyasprodINI   As String

    Screen.MousePointer = vbHourglass
    'Sleep (100)
    
    ChDir App.Path
    
    'Verifica se o Programa já esta ativo
    If App.PrevInstance = True Then
        Call gM46V999.gpGraLog(2, "O Sistema de Cálculo já esta ativo.")
        End
    End If
    
    '2. Obtém informações iniciais do sistema: tipo do usuário, endereços das bases de dados, endereços
    '   dos arquivos de impressão, indicador de uso do auxílio automático, nome ou código do usuário.
    Call gM46V999.gpInicia("")

    '3. Emite mensagem de entrada (precisa ser depois de obter dados do arquivo de inicialização).
    Call gM46V999.gpGraLog(0, "AGE0001 - Início do processamento do programa " & App.EXEName & ".")

    '4. Se data do sistema incoerente com data do módulo executável, encerra processo.
'    If Dir(gM46V999.gAppPath & "P47V001.EXE") <> "" Then
'        lDatModulo = Format$(FileDateTime(gM46V999.gAppPath & "P47V001.EXE"), "yyyymmdd")
'        If lDatModulo > gM46V999.gHoje Then
'            Call gM46V999.gpGraLog(2, "FTE2008 - Data do microcomputador (" & gM46V999.gForHoje & _
'                             ") menor que data do programa (" & gM46V999.gfForData(2, lDatModulo) & ")." & vbLf & _
'                             vbLf & "Programa será encerrado.")
'            End     'Encerra programa.
'        End If
'    End If
    '5. Posiciona variável gVerDataMaq (indica se a função que verifica se houve ajuste na máquina
    '   deve ser processada: 0 = não verifica e 1 = verifica).
    If App.EXEName <> "P47V001" Then
        gM46V999.gVerDataMaq = 1
    Else
        gM46V999.gVerDataMaq = 0
    End If
    gM46V999.gTrvSistema = False
      
    'Posiciona controle dos endereços dos arquivos (gM46V999.gAppPath).
    gM46V999.gForHoje = Format$(Now, "dd/mm/yyyy")
    
    
    '2- Obtem informações do arquivo de inicialização
    'If mfObtemDadosINI(lMensagem) = False Then
    '    Call gM46V999.gpGraLog(2, lMensagem)
    '    End
    'End If
    gNomComputador = Trim$(gM99V001.GetLoginComputerName)
    gM46V999.gBaseLocal = True
    gM46V999.gTipP00MULTI = "1"
    'Simony
    lEndSyasprodINI = gM46V999.gAppPath & "SYASPROD.INI"
    
    
    
    
    frmT47V001A.mfVerificaVersao
    
    Call gM46V999.gpGraLog(0, "Endereço " & gM46V999.gEndP00MULTI & ", Tipo " & gM46V999.gTipoHomologacao)
    '10. Posiciona endereço do arquivo SYASPROD.HLP.
    App.HelpFile = gM46V999.gAppPath & "SYASMULT.HLP"

    '11. Posiciona controle.
    gM46V111.gSaida = False
    gVerDataMaq = 0
    If gM46V999.gTrvSistema = True Then
        End
    End If

    
'
'    gM46V999.gEndP00MULTI = LerIni("P00MULTI", "DATABASE", lEndSyasprodINI)
'
'    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, gConexaoSql, lMensagem) = False Then
'        Call gM46V999.gpGraLog(2, lMensagem)
'        Exit Sub
'    End If

    Screen.MousePointer = vbDefault
    frmT47V001A.Show
    
        
End Sub

Sub CriaTabelaAnexo()
Dim CmdSql     As String
Dim UltimoErro As String

             
    CmdSql = "CREATE TABLE  Tab_Anexos  (" & _
            " Cod_Nosso_Contrato  text (20) Not NULL ," & _
            " Num_Item  integer Not NULL ," & _
            " Nome_Arq  text (250)  NULL ," & _
            " Novo_Arq  text (40)  NULL  ," & _
            " Dat_Inclusao  double NULL ," & _
            " Dir_Arq  text (255)  NULL  ," & _
            " Tam_arq  text (10)  NULL , " & _
            "CONSTRAINT PKTab_Anexos PRIMARY KEY (" & _
                  " Cod_Nosso_Contrato , Num_Item )) ;"

    UltimoErro = ""
    Call gM46V999.gfExeSQL(gConexaoSql, CmdSql, UltimoErro)
    
    

End Sub

Public Sub gpPausa2(ByVal lSegundos)
Dim i
Dim lMiliSegundos   As Double
    
    lMiliSegundos = lSegundos * 1000

    i = DoEvents
    Sleep (lMiliSegundos)
    i = DoEvents
    
    Exit Sub
    
End Sub


Private Function mfObtemDadosINI(ByRef pMensagem As String) As Boolean
   'Função: Busca os parametros de inicialização
Dim lEndArq_INI As String
On Error GoTo mfObtemDadosINI_ERRO

    mfObtemDadosINI = False

    '1. Verifica arquivo .INI.
    lEndArq_INI = gM46V999.gAppPath & "P12V001.INI"
    If Not gM46V999.gfPreenchido(Dir(lEndArq_INI)) Then
        pMensagem = "Não localizado arquivo de inicialização"
        Exit Function
    End If

    gTipo_bd = LerIni("PRINCIPAL", "TIPO", lEndArq_INI)
    gPWD = LerIni("PRINCIPAL", "PWD", lEndArq_INI)
    gServer = LerIni("PRINCIPAL", "SERVER", lEndArq_INI)
    gUID = LerIni("PRINCIPAL", "UID", lEndArq_INI)
    gDataBase = LerIni("PRINCIPAL", "DATABASE", lEndArq_INI)
    
    
    mfObtemDadosINI = True
mfObtemDadosINI_SAIDA:
    Exit Function

mfObtemDadosINI_ERRO:
    If gM46V999.gfPreenchido(pMensagem) = True Then
        pMensagem = " mfObtemDadosINI - ERRO " & pMensagem
    Else
        pMensagem = " mfObtemDadosINI - ERRO " & Err & " - " & Error$
    End If
    GoTo mfObtemDadosINI_SAIDA
    
    Resume Next
    
End Function


Function GravarINI(Sessao As String, NomeChave As String, Valor As String, NomedoArquivo As String)
Dim rc As Long
   rc = WritePrivateProfileString(Sessao, NomeChave, Valor, NomedoArquivo)
End Function
Function LerIni(Secao As String, NomeChave As String, NomedoArquivo As String) As String
  'NomedoArquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'NomeChave=nome do que se encontra antes do sinal de igual

  Dim RetCode As Long
  Dim RetString As String
       
  LerIni = ""
  RetString = String$(255, 0)
  RetCode = GetPrivateProfileString(Secao, NomeChave, "", RetString, Len(RetString), NomedoArquivo)
  LerIni = Left(RetString, RetCode)
    
End Function

Public Function gfTextBox_GotFocus(ByVal pTextBox As TextBox, _
                                   ByVal E_SelGotFocus As E_SelGotFocus, _
                                   ByRef pMsgErr As String) As Boolean

Err = 0
pMsgErr = Empty
On Error GoTo Func_Erro
    gfTextBox_GotFocus = False

    With pTextBox
        .Text = Trim$(.Text)
        Select Case E_SelGotFocus
            Case E_SelInicio
                .SelStart = 0
                .SelLength = 0
            Case E_SelFim
                .SelStart = Len(.Text)
                .SelLength = Len(.Text)
            Case E_SelTudo
                .SelStart = 0
                .SelLength = Len(.Text)
        End Select
    End With

Ok:
    gfTextBox_GotFocus = True

Func_Saida:
    Err = 0
    On Error GoTo 0
    Exit Function
    
Func_Erro:
    On Error Resume Next
        pMsgErr = "gfTextBox_GotFocus Erro " & Err & " - " & Error & Chr(10) & Chr(13)
        pMsgErr = pMsgErr & "Objeto : " & pTextBox.Name
    Err = 0
    On Error GoTo 0
    GoTo Func_Saida
End Function

Public Function gfTextBox_KeyPress(ByVal pTextBox As TextBox, _
                                   ByVal pTipoDado As TipoDados, _
                                   ByVal pTamanho As Double, _
                                   KeyAscii As Integer, _
                                   ByRef pMsgErr As String, _
                                   Optional pTextMultLine As Boolean = False) As Boolean
Dim lContX      As Integer
Dim lNumAux     As Integer
Dim lNumBarras  As Byte         'Numero de barras em campo data
Dim lValTexto   As String
Dim lTamanhoInt As Integer
Dim lTamanhoDec As Integer
Dim lPosVirgula As Integer

Err = 0
pMsgErr = Empty
On Error GoTo Func_Erro
    gfTextBox_KeyPress = False
       
'Valida a entrada no campo
    Select Case pTipoDado
        Case TpoData
'Tipo Data
            If KeyAscii = vbKeyReturn Then
                KeyAscii = 0
                SendKeys "{TAB}"
            End If
            
            Select Case KeyAscii
                Case 3
                '"Control"+"C" (Copiar)
                    Clipboard.Clear
                    If pTextBox.SelLength > 0 Then
                        If pTextBox.SelStart = 0 Then
                            Clipboard.SetText pTextBox.Text
                        Else
                            Clipboard.SetText Mid(pTextBox.Text, pTextBox.SelStart)
                        End If
                    End If
                Case 47 To 57 'Numeros e barra "/"
                'Caracteres válidos
                    If (pTextBox.SelLength = Len(pTextBox.Text) And pTextBox.SelLength > 0) Then  'Não esta substituindo o conteudo atual
                        If (KeyAscii = 47) And (pTextBox.SelLength = Len(pTextBox.Text)) Then
                            'Esta substituindo o texto 1º caracter não pode ser barra
                            KeyAscii = 8 'Apaga tudo
                        End If
                    Else
                        lValTexto = pTextBox.Text
                        If KeyAscii = 47 Then
                            If gM46V999.gfPreenchido(lValTexto) = False Then
                            'barra não pobe ser o 1º caractar
                                KeyAscii = 0
                            Else
                            'Confere as barras
                                lNumAux = 0
                                lNumBarras = 0
                                Do
                                    lNumAux = InStr((lNumAux + 1), lValTexto, "/")
                                    If lNumAux = 0 Then
                                        'A barra 1ª esta sendo inserida
                                        lNumAux = Len(lValTexto) + 1
                                        If lNumAux > 1 And lNumAux < 4 Then
                                            lNumBarras = 1
                                        Else
                                        'Barra possivel somente nas posições 2 e 3
                                            KeyAscii = 0
                                        End If
                                        Exit Do 'Sai do loop de validação
                                    Else
                                        lNumAux = InStr((lNumAux + 1), lValTexto, "/")
                                        If lNumAux = 0 Then
                                            'A barra 2ª esta sendo inserida
                                            lNumAux = Len(lValTexto) + 1
                                            If lNumAux > 3 And lNumAux < 7 Then
                                                If Not ((lNumAux = 4) And (Mid(lValTexto, 3, 1) = "/")) Then
                                                    lNumBarras = 2
                                                Else
                                                'As barras não podem aparecer em sequencia
                                                    KeyAscii = 0 'Sai do loop de validação
                                                End If
                                            Else
                                            'Barra possivel somente nas posições 2 e 3
                                                KeyAscii = 0
                                            End If
                                        Else
                                        'Já existem 2 barras não insere mais nenhuma
                                            KeyAscii = 0
                                        End If
                                        Exit Do 'Sai do loop de validação
                                    End If
                                Loop
                            End If
                        Else
                            'Apenas conta o total de barras
                            lNumBarras = 0
                            lNumAux = 0
                            Do
                                lNumAux = InStr((lNumAux + 1), lValTexto, "/")
                                If lNumAux <> 0 Then
                                    lNumBarras = lNumBarras + 1
                                Else
                                    Exit Do
                                End If
                            Loop
                        End If
                
                        If lNumBarras = 2 Then
                        'Pode no máximo adicionar 4 digitos após a segunda barra
                            lNumAux = InStr(4, pTextBox.Text, "/")
                            lValTexto = Mid(pTextBox.Text, lNumAux + 1)
                            If lNumAux > 0 And Len(lValTexto) > 3 Then
                                KeyAscii = 0
                            End If
                        Else
                            lValTexto = Replace(lValTexto, "/", "")
                            lContX = pTextBox.SelStart
                            If Len(lValTexto) = pTamanho And lContX <> 0 Then
                                KeyAscii = 0
                            End If
                        End If
                        
                        
                    End If
                Case 8
                    'Ok! BackSpace
                Case Else
                    KeyAscii = 0
            End Select

        Case TpoNumInteiro
'Tipo Numero Inteiro
            If KeyAscii = vbKeyReturn Then
                KeyAscii = 0
                SendKeys "{TAB}"
            End If
            
            Select Case KeyAscii
                Case 3
                '"Control"+"C" (Copiar)
                    Clipboard.Clear
                    If pTextBox.SelLength > 0 Then
                        If pTextBox.SelStart = 0 Then
                            Clipboard.SetText pTextBox.Text
                        Else
                            Clipboard.SetText Mid(pTextBox.Text, pTextBox.SelStart)
                        End If
                    End If
                Case 48 To 57 'Numeros
                'Caracteres válidos
                    lValTexto = pTextBox.Text
                    lContX = pTextBox.SelStart
                    If Len(lValTexto) >= Val(pTamanho) And lContX <> 0 Then
                        KeyAscii = 0
                    End If
               Case 8
                    'Ok! BackSpace
                Case 22 '"Control"+"V" (Colar)
                    With Clipboard
                        If .GetFormat(vbCFText) = False Then
                            'Permite apenas texto
                            KeyAscii = 0
                            .Clear
                        Else
                            lValTexto = .GetText
                            If IsNumeric(lValTexto) = False Then
                                KeyAscii = 0
                                .Clear
                            Else
                                If Len(lValTexto) > pTamanho Then
                                    'Trunca para o tamanho máximo permitido
                                    .SetText Mid(lValTexto, 1, pTamanho)
                                End If
                                'limita no tamanho máximo
                                If (Len(pTextBox.Text) + Len(lValTexto)) > Val(pTamanho) Then
                                    KeyAscii = 0
                                    If Len(pTextBox.Text) > Val(pTamanho) Then
                                        'trunca
                                        pTextBox.Text = Mid(pTextBox.Text, 1, pTamanho)
                                    End If
                                End If
                            End If
                        End If
                    End With
                Case Else
                    KeyAscii = 0
            End Select
            
        Case tpoString
'Tipo STRING
            'Define o tamanho máximo do campo
            If pTextBox.MaxLength <> pTamanho Then
                pTextBox.MaxLength = pTamanho
            End If
            Select Case KeyAscii
                Case 3
                '"Control"+"C" (Copiar)
                    Clipboard.Clear
                    If pTextBox.SelLength > 0 Then
                        If pTextBox.SelStart = 0 Then
                            Clipboard.SetText pTextBox.Text
                        Else
                            Clipboard.SetText Mid(pTextBox.Text, pTextBox.SelStart)
                        End If
                    End If
                Case 8
                    'Ok! BackSpace
                Case 22
                    With Clipboard
                        If .GetFormat(vbCFText) = False Then
                            'Permite apenas texto
                            KeyAscii = 0
                            .Clear
                        Else
                            lValTexto = .GetText
                            If Len(lValTexto) > pTamanho Then
                                'Trunca para o tamanho máximo permitido
                                .SetText Mid(lValTexto, 1, pTamanho)
                            End If
                            'limita no tamanho máximo
                            If (Len(pTextBox.Text) + Len(lValTexto)) > Val(pTamanho) Then
                                KeyAscii = 0
                                If Len(pTextBox.Text) > Val(pTamanho) Then
                                    'trunca
                                    pTextBox.Text = Mid(pTextBox.Text, 1, pTamanho)
                                End If
                            End If
                        End If
                    End With
                Case Else
                    If KeyAscii = vbKeyReturn And pTextMultLine = False Then
                        KeyAscii = 0
                        SendKeys "{TAB}"
                    End If
                    'limita no tamanho máximo
                    If Len(pTextBox.Text) >= Val(pTamanho) Then
                        KeyAscii = 0
                        If Len(pTextBox.Text) > Val(pTamanho) Then
                            'trunca
                            pTextBox.Text = Mid(pTextBox.Text, 1, pTamanho)
                        End If
                    End If
            End Select
        Case TpoNumDecimal
'Tipo Numero Decimal
            If KeyAscii = vbKeyReturn Then
                KeyAscii = 0
                SendKeys "{TAB}"
            End If
            
            'Obtem os tamanhos das partes inteiras e decimais
            lTamanhoInt = Val(pTamanho)
            If Val(pTamanho) < pTamanho Then
                lTamanhoDec = ((pTamanho - Val(pTamanho)) * CDbl(10))
            Else
                lTamanhoDec = 0
            End If
            
            lValTexto = pTextBox.Text
            Select Case KeyAscii
                Case 3
                '"Control"+"C" (Copiar)
                    Clipboard.Clear
                    If pTextBox.SelLength > 0 Then
                        If pTextBox.SelStart = 0 Then
                            Clipboard.SetText pTextBox.Text
                        Else
                            Clipboard.SetText Mid(pTextBox.Text, pTextBox.SelStart)
                        End If
                    End If
                Case 44 'vírgula
                    If InStr(1, lValTexto, ",") > 0 Then
                    'permite apenas 1 virgula
                        KeyAscii = 0
                    Else
                        If lTamanhoDec = 0 Then
                            'Não informado casas decimais, Não permite virgula
                            KeyAscii = 0
                        Else
                            'Virgula somente até o ultimo caracter inteiro
                            'Ok .. case abaixo não permite incluir além do permitido
                        End If
                    End If
                Case 48 To 57   'Numeros
                    lContX = pTextBox.SelStart
                    lPosVirgula = InStr(1, lValTexto, ",")
                    If lPosVirgula = 0 Then
                    'Parte inteira
                        If Len(lValTexto) = lTamanhoInt And lContX <> 0 Then
                            KeyAscii = 0
                        End If
                    Else
                        If lContX < lPosVirgula Then
                        'Parte inteira
                            If Len(Mid(lValTexto, 1, lPosVirgula - 1)) = lTamanhoInt Then
                                KeyAscii = 0
                            End If
                        Else
                        'parte decimal
                            If Len(Mid(lValTexto, lPosVirgula + 1)) = lTamanhoDec Then
                                KeyAscii = 0
                            End If
                        End If
                    End If
                Case 8
                    'Ok! BackSpace
                Case Else
                    KeyAscii = 0
            End Select
        Case Else
            pMsgErr = "gfTextBox_KeyPress Erro: 'Tipo de dado informado desconhecido (" & pTipoDado & ")'" & Chr(10) & Chr(13)
            pMsgErr = pMsgErr & "Objeto : " & pTextBox.Name
            GoTo Func_Saida
    End Select
    
Ok:
    gfTextBox_KeyPress = True

Func_Saida:
    Err = 0
    On Error GoTo 0
    Exit Function
    
Func_Erro:
    On Error Resume Next
        pMsgErr = "gfTextBox_KeyPress Erro " & Err & " - " & Error & Chr(10) & Chr(13)
        pMsgErr = pMsgErr & "Objeto : " & pTextBox.Name
    Err = 0
    On Error GoTo 0
    GoTo Func_Saida
End Function

Public Function gfCarCboBancos(ByVal pOpcao As Byte, ByVal pCombo As ComboBox, _
                               ByRef pMensagem As String, Optional pCorretor As Long) As Boolean
    'Função: carrega combo de bancos.

    'Parâmetro de entrada...pOpcao......1 = Todos os bancos.
    '                                   2 = Bancos para carnê...............237, 341, 347, 409, 479.
    '                                   3 = Bancos para débito em conta.....237, 341, 347, 409, 479.
    '                                   4 = Bancos de indicador de seguro...215, 341, 479, 347.
    '                       pCombo......Combo de bancos.
    'Parâmetro de saída.....pMensagem...Mensagem.

    Dim lBanco      As String           'Linha de banco.
    
    gfCarCboBancos = False

    '1. Verifica opção.
    If Not gM46V999.gfPreenchido(pOpcao) Then
        pMensagem = "FTE0059 - Opção inválida (carrega relação de bancos - opção não preenchida)."
        Exit Function
    End If
    If Not IsNumeric(pOpcao) Then
        pMensagem = "FTE0059 - Opção inválida (carrega relação de bancos - opção não numérica)."
        Exit Function
    End If
    If pOpcao <> 1 And pOpcao <> 2 And pOpcao <> 3 And pOpcao <> 4 Then
        pMensagem = "FTE0059 - Opção inválida (carrega relação de bancos - opção diferente de 1, 2, " & _
                    "3 e 4: " & pOpcao & ")."
        Exit Function
    End If

    '2. Limpa combo.
    pCombo.Clear
    
    '4. Carrega combo.
    Select Case pOpcao
        Case 1      'Todos os bancos.
        Case 2      'Bancos para carnê.
            
            'pCombo.AddItem "237 - BRADESCO SA"
            'pCombo.ItemData(pCombo.NewIndex) = 237
            'pCombo.AddItem "341 - ITAU"
            'pCombo.ItemData(pCombo.NewIndex) = 341
            ''pCombo.AddItem "347 - SUDAMERIS BRASIL"
            ''pCombo.ItemData(pCombo.NewIndex) = 347
            'pCombo.AddItem "409 - UNIBANCO"
            'pCombo.ItemData(pCombo.NewIndex) = 409
            
            ''09029
            'pCombo.AddItem "356 - REAL"
            'pCombo.ItemData(pCombo.NewIndex) = 356
            pCombo.AddItem "033 - SANTANDER"
            pCombo.ItemData(pCombo.NewIndex) = 33
        Case 3      'Bancos para débito em conta
            pCombo.AddItem "001 - BANCO DO BRASIL"
            pCombo.ItemData(pCombo.NewIndex) = 1
            pCombo.AddItem "237 - BRADESCO SA"
            pCombo.ItemData(pCombo.NewIndex) = 237
            pCombo.AddItem "341 - ITAU"
            pCombo.ItemData(pCombo.NewIndex) = 341
            'pCombo.AddItem "347 - SUDAMERIS BRASIL"
            'pCombo.ItemData(pCombo.NewIndex) = 347
            'pCombo.AddItem "409 - UNIAO DE BANCOS"
            'pCombo.ItemData(pCombo.NewIndex) = 409
            'pCombo.AddItem "479 - BOSTON"
            'pCombo.ItemData(pCombo.NewIndex) = 479
            '09029
            pCombo.AddItem "356 - REAL"
            pCombo.ItemData(pCombo.NewIndex) = 356
            
            pCombo.AddItem "399 - HSBC"
            pCombo.ItemData(pCombo.NewIndex) = 399
            pCombo.AddItem "047 - BANESE"
            pCombo.ItemData(pCombo.NewIndex) = 47
            
        Case 4      'Bancos de indicador de seguro.
            pCombo.AddItem "033 - SANTANDER"
            pCombo.ItemData(pCombo.NewIndex) = 33
            'pCombo.AddItem "341 - ITAU"
            'pCombo.ItemData(pCombo.NewIndex) = 341
            'pCombo.AddItem "347 - SUDAMERIS BRASIL"
            'pCombo.ItemData(pCombo.NewIndex) = 347
            'pCombo.AddItem "409 - UNIAO DE BANCOS"
            'pCombo.ItemData(pCombo.NewIndex) = 409
            '09029
            'pCombo.AddItem "356 - REAL"
            'pCombo.ItemData(pCombo.NewIndex) = 356
    End Select

    gfCarCboBancos = True
End Function
Public Function gfVerAgencia(ByVal pOpcao As Byte, ByVal pCodBanco As String, _
                             ByVal pCodAgencia As String, ByRef pSaida As String, _
                             ByRef pMensagem As String) As Boolean
    'Função: verifica e formata campo de agência de acordo com banco informado.

    'Parâmetros de entrada...pOpcao........1 = Conta para débito.
    '                                      2 = Conta do indicador de seguro.
    '                        pCodBanco.....Código do banco.
    '                        pCodAgencia...Código de agência.
    'Parâmetros de saída.....pSaida........Saída - código de agência.
    '                        pMensagem.....Mensagem.

    Dim lAuxiliar   As String   'Auxiliar.
    Dim lPosicao    As Byte     'Utilizado com InStr.
    Dim lTemErro    As Boolean  'Tem erro?

    gfVerAgencia = False

    '1. Verifica opção.
    If IsNull(pOpcao) Then
        pMensagem = "FTE0060 - Opção inválida (verifica código de agência - opção não preenchida)."
        Exit Function
    End If
    If Not IsNumeric(pOpcao) Then
        pMensagem = "FTE0060 - Opção inválida (verifica código de agência - opção não numérica)."
        Exit Function
    End If
    If pOpcao <> 1 And pOpcao <> 2 Then
        pMensagem = "FTE0060 - Opção inválida (verifica código de agência - opção diferente de 1 e " & _
                    "2: " & pOpcao & ")."
        Exit Function
    End If

    '2. Verifica código da agência.
    If Not gM46V999.gfPreenchido(pCodAgencia) Then
        Select Case pOpcao
            Case 1
                pMensagem = "FGE0029 - " & gM46V999.gPrefixo & _
                            "Código de agência inválido (conta do indicador do seguro - não preenchido)."
            Case 2
                pMensagem = "FGE0029 - " & gM46V999.gPrefixo & _
                            "Código de agência inválido (conta para débito - não preenchido)."
        End Select
        pSaida = ""
        Exit Function
    End If

    '3. Verifica preenchimento do código da agência.
    lAuxiliar = Trim$(pCodAgencia)
    If Not gM46V999.gfPreenchido(pCodBanco) Then
        pSaida = lAuxiliar
        gfVerAgencia = True
        Exit Function
    End If
    lTemErro = False
    Select Case pCodBanco
        Case "237"      'Bradesco.
            lPosicao = InStr(lAuxiliar, "-")
            If Len(lAuxiliar) <> 6 Or lPosicao <> 5 Then
                Select Case lPosicao
                    Case 0
                        'Nada a fazer.
                    Case 1
                        If Len(lAuxiliar) = 1 Then
                            lAuxiliar = ""
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 2)
                        End If
                    Case Else
                        If Len(lAuxiliar) = lPosicao Then
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1)
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1) & _
                                        Mid$(lAuxiliar, lPosicao + 1)
                        End If
                End Select
                If Len(lAuxiliar) < 5 Then
                    lAuxiliar = String(5 - Len(lAuxiliar), "0") & lAuxiliar
                End If
                lAuxiliar = Mid$(lAuxiliar, 1, 4) & "-" & Mid$(lAuxiliar, 5, 1)
            End If
            If Not IsNumeric(Mid(lAuxiliar, 1, 4)) Or Not IsNumeric(Mid(lAuxiliar, 6, 1)) Then
                lTemErro = True
            End If
        Case "409"      'Unibanco.
            If Not IsNumeric(lAuxiliar) Then
                lTemErro = True
            Else
                If Len(lAuxiliar) < 4 Then
                    lAuxiliar = String(4 - Len(lAuxiliar), "0") & lAuxiliar
                End If
            End If
        Case "341"      'Itaú.
            If Not IsNumeric(lAuxiliar) Then
                lTemErro = True
            Else
                If Len(lAuxiliar) < 4 Then
                    lAuxiliar = String(4 - Len(lAuxiliar), "0") & lAuxiliar
                End If
            End If
        Case "347"      'Sudameris.
            If Not IsNumeric(lAuxiliar) Then
                lTemErro = True
            Else
                If Len(lAuxiliar) < 3 Then
                    lAuxiliar = String(3 - Len(lAuxiliar), "0") & lAuxiliar
                End If
            End If
        Case "479"      'Boston.
            If Not IsNumeric(lAuxiliar) Then
                lTemErro = True
            Else
                If Len(lAuxiliar) < 3 Then
                    lAuxiliar = String(3 - Len(lAuxiliar), "0") & lAuxiliar
                End If
            End If
        Case "356"      'Real
            If Not IsNumeric(lAuxiliar) Then
                lTemErro = True
            Else
                If Len(lAuxiliar) < 3 Then
                    lAuxiliar = String(3 - Len(lAuxiliar), "0") & lAuxiliar
                End If
            End If
        Case Else       'Outros bancos.
            If Not IsNumeric(lAuxiliar) Then
                lTemErro = True
            End If
    End Select
    If lTemErro = True Then
        Select Case pOpcao
            Case 1
                pMensagem = "FGE0029 - " & gM46V999.gPrefixo & _
                            "Código de agência inválido (não numérico - conta do indicador do seguro)."
            Case 2
                pMensagem = "FGE0029 - " & gM46V999.gPrefixo & _
                            "Código de agência inválido (não numérico - conta para débito)."
        End Select
        Exit Function
    End If
    pSaida = lAuxiliar

    gfVerAgencia = True
End Function

Public Function gfVerConta(ByVal pOpcao As Byte, ByVal pCodBanco As String, _
                           ByVal pNumConta As String, ByRef pSaida As String, _
                           ByRef pMensagem As String) As Boolean
    'Função: verifica e formata campo de conta de acordo com banco informado.

    'Parâmetros de entrada...pOpcao........1 = Conta para débito.
    '                                      2 = Conta do indicador de seguro.
    '                        pCodBanco.....Código do banco.
    '                        pNumConta.....Número da conta corrente.
    'Parâmetros de saída.....pSaida........Saída - número da conta corrente.
    '                        pMensagem.....Mensagem.

    Dim lAuxiliar   As String       'Auxiliar.
    Dim lPosicao    As Byte         'Utilizado com InStr.

    gfVerConta = False

    '1. Verifica opção.
    If IsNull(pOpcao) Then
        pMensagem = "FTE0058 - Opção inválida (verifica conta corrente - opção não preenchida)."
        Exit Function
    End If
    If Not IsNumeric(pOpcao) Then
        pMensagem = "FTE0058 - Opção inválida (verifica conta corrente - opção não numérica)."
        Exit Function
    End If
    If pOpcao <> 1 And pOpcao <> 2 Then
        pMensagem = "FTE0058 - Opção inválida (verifica conta corrente - opção diferente de 1 e 2: " & _
                    pOpcao & ")."
        Exit Function
    End If

    '2. Verifica se número da conta corrente está preenchido.
    If Not gM46V999.gfPreenchido(pNumConta) Then
        Select Case pOpcao
            Case 1
                pMensagem = "FNP2002 - " & gM46V999.gPrefixo & _
                            "Número da conta corrente não preenchido (conta para débito)."
            Case 2
                pMensagem = "FNP2002 - " & gM46V999.gPrefixo & _
                            "Número da conta corrente não preenchido (conta do indicador do seguro)."
        End Select
        pSaida = ""
        Exit Function
    End If

    '3. Verifica preenchimento do número da conta corrente.
    lAuxiliar = Trim$(pNumConta)
    If Not gM46V999.gfPreenchido(pCodBanco) Then
        pSaida = lAuxiliar
        gfVerConta = True
        Exit Function
    End If
    Select Case pCodBanco
        Case "237"      'Bradesco.
            lPosicao = InStr(lAuxiliar, "-")
            If Len(lAuxiliar) <> 9 Or lPosicao <> 8 Then
                Select Case lPosicao
                    Case 0
                        'Nada a fazer.
                    Case 1
                        If Len(lAuxiliar) = 1 Then
                            lAuxiliar = ""
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 2)
                        End If
                    Case Else
                        If Len(lAuxiliar) = lPosicao Then
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1)
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1) & _
                                        Mid$(lAuxiliar, lPosicao + 1)
                        End If
                End Select
                If Len(lAuxiliar) < 8 Then
                    lAuxiliar = String(8 - Len(lAuxiliar), "0") & lAuxiliar
                End If
                lAuxiliar = Mid$(lAuxiliar, 1, 7) & "-" & Mid$(lAuxiliar, 8, 1)
            End If
            If Not IsNumeric(Mid(lAuxiliar, 1, 7)) Or Not IsNumeric(Mid(lAuxiliar, 9, 1)) Then
                Select Case pOpcao
                    Case 1
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta para débito. " & _
                                    "Número de conta corrente inválido (Bradesco - formato NNNNNNN-N)."
                    Case 2
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta do indicador do seguro. " & _
                                    "Número de conta corrente inválido (Bradesco - formato NNNNNNN-N)."
                End Select
                pSaida = ""
                Exit Function
            End If
        Case "409", "001"  'Unibanco 'Projeto 11080 - Debito automatico - Banco do Brasil
            lPosicao = InStr(lAuxiliar, "-")
            If Len(lAuxiliar) <> 8 Or lPosicao <> 7 Then
                Select Case lPosicao
                    Case 0
                        'Nada a fazer.
                    Case 1
                        If Len(lAuxiliar) = 1 Then
                            lAuxiliar = ""
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 2)
                        End If
                    Case Else
                        If Len(lAuxiliar) = lPosicao Then
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1)
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1) & _
                                        Mid$(lAuxiliar, lPosicao + 1)
                        End If
                End Select
                If Len(lAuxiliar) < 7 Then
                    lAuxiliar = String(7 - Len(lAuxiliar), "0") & lAuxiliar
                End If
                lAuxiliar = Mid$(lAuxiliar, 1, 6) & "-" & Mid$(lAuxiliar, 7, 1)
            End If
            If Not IsNumeric(Mid(lAuxiliar, 1, 6)) Or Not IsNumeric(Mid(lAuxiliar, 8, 1)) Then
                Select Case pOpcao
                    Case 1
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta para débito. " & _
                                    "Número de conta corrente inválido (Unibanco - formato NNNNNN-N)."
                    Case 2
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta do indicador do seguro. " & _
                                    "Número de conta corrente inválido (Unibanco - formato NNNNNN-N)."
                End Select
                pSaida = ""
                Exit Function
            End If
        Case "341"      'Itaú.
            lPosicao = InStr(lAuxiliar, "-")
            If Len(lAuxiliar) <> 7 Or lPosicao <> 6 Then
                Select Case lPosicao
                    Case 0
                        'Nada a fazer.
                    Case 1
                        If Len(lAuxiliar) = 1 Then
                            lAuxiliar = ""
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 2)
                        End If
                    Case Else
                        If Len(lAuxiliar) = lPosicao Then
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1)
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1) & _
                                        Mid$(lAuxiliar, lPosicao + 1)
                        End If
                End Select
                If Len(lAuxiliar) < 6 Then
                    lAuxiliar = String(6 - Len(lAuxiliar), "0") & lAuxiliar
                End If
                lAuxiliar = Mid$(lAuxiliar, 1, 5) & "-" & Mid$(lAuxiliar, 6, 1)
            End If
            If Not IsNumeric(Mid(lAuxiliar, 1, 5)) Or Not IsNumeric(Mid(lAuxiliar, 7, 1)) Then
                Select Case pOpcao
                    Case 1
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta para débito. " & _
                                    "Número de conta corrente inválido (Banco Itaú - formato NNNNN-N)."
                    Case 2
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta do indicador do seguro. " & _
                                    "Número de conta corrente inválido (Banco Itaú - formato NNNNN-N)."
                End Select
                Exit Function
            End If
        Case "347"      'Sudameris.
            lPosicao = InStr(lAuxiliar, "-")
            If Len(lAuxiliar) <> 11 Or lPosicao <> 10 Then
                Select Case lPosicao
                    Case 0
                        'Nada a fazer.
                    Case 1
                        If Len(lAuxiliar) = 1 Then
                            lAuxiliar = ""
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 2)
                        End If
                    Case Else
                        If Len(lAuxiliar) = lPosicao Then
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1)
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1) & _
                                        Mid$(lAuxiliar, lPosicao + 1)
                        End If
                End Select
                If Len(lAuxiliar) < 10 Then
                    lAuxiliar = String(10 - Len(lAuxiliar), "0") & lAuxiliar
                End If
                lAuxiliar = Mid$(lAuxiliar, 1, 9) & "-" & Mid$(lAuxiliar, 10, 1)
            End If
            If Not IsNumeric(Mid(lAuxiliar, 1, 9)) Or Not IsNumeric(Mid(lAuxiliar, 11, 1)) Then
                Select Case pOpcao
                    Case 1
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta para débito. " & _
                                    "Número de conta corrente inválido (Sudameris - formato NNNNNNNNN-N)."
                    Case 2
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta do indicador do seguro. " & _
                                    "Número de conta corrente inválido (Sudameris - formato NNNNNNNNN-N)."
                End Select
                pSaida = ""
                Exit Function
            End If
        Case "356"      'Real
            lPosicao = InStr(lAuxiliar, "-")
            If Len(lAuxiliar) <> 11 Or lPosicao <> 10 Then
                Select Case lPosicao
                    Case 0
                        'Nada a fazer.
                    Case 1
                        If Len(lAuxiliar) = 1 Then
                            lAuxiliar = ""
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 2)
                        End If
                    Case Else
                        If Len(lAuxiliar) = lPosicao Then
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1)
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1) & _
                                        Mid$(lAuxiliar, lPosicao + 1)
                        End If
                End Select
                If Len(lAuxiliar) < 10 Then
                    lAuxiliar = String(10 - Len(lAuxiliar), "0") & lAuxiliar
                End If
                lAuxiliar = Mid$(lAuxiliar, 1, 9) & "-" & Mid$(lAuxiliar, 10, 1)
            End If
            If Not IsNumeric(Mid(lAuxiliar, 1, 9)) Or Not IsNumeric(Mid(lAuxiliar, 11, 1)) Then
                Select Case pOpcao
                    Case 1
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta para débito. " & _
                                    "Número de conta corrente inválido (Real - formato NNNNNNNNN-N)."
                    Case 2
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta do indicador do seguro. " & _
                                    "Número de conta corrente inválido (Real - formato NNNNNNNNN-N)."
                End Select
                pSaida = ""
                Exit Function
            End If
        Case "479"      'Boston.
            lPosicao = InStr(lAuxiliar, "-")
            If Len(lAuxiliar) <> 9 Or lPosicao <> 8 Then
                Select Case lPosicao
                    Case 0
                        'Nada a fazer.
                    Case 1
                        If Len(lAuxiliar) = 1 Then
                            lAuxiliar = ""
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 2)
                        End If
                    Case Else
                        If Len(lAuxiliar) = lPosicao Then
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1)
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1) & _
                                        Mid$(lAuxiliar, lPosicao + 1)
                        End If
                End Select
                If Len(lAuxiliar) < 8 Then
                    lAuxiliar = String(8 - Len(lAuxiliar), "0") & lAuxiliar
                End If
                lAuxiliar = Mid$(lAuxiliar, 1, 7) & "-" & Mid$(lAuxiliar, 8, 1)
            End If
            If Not IsNumeric(Mid(lAuxiliar, 1, 7)) Or Not IsNumeric(Mid(lAuxiliar, 9, 1)) Then
                Select Case pOpcao
                    Case 1
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta para débito. " & _
                                    "Número de conta corrente inválido (Bank Boston - formato NNNNNNN-N)."
                    Case 2
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta do indicador do seguro. " & _
                                    "Número de conta corrente inválido (Bank Boston - formato NNNNNNN-N)."
                End Select
                pSaida = ""
                Exit Function
            End If
        Case "399"      'HSBC
            lPosicao = InStr(lAuxiliar, "-")
            If Len(lAuxiliar) <> 12 Then
                Select Case lPosicao
                    Case 0
                        'Nada a fazer.
                    Case 1
                        If Len(lAuxiliar) = 1 Then
                            lAuxiliar = ""
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 2)
                        End If
                    Case Else
                        If Len(lAuxiliar) = lPosicao Then
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1)
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1) & _
                                        Mid$(lAuxiliar, lPosicao + 1)
                        End If
                End Select
                If Len(lAuxiliar) < 11 Then
                    lAuxiliar = String(11 - Len(lAuxiliar), "0") & lAuxiliar
                End If
                lAuxiliar = Mid$(lAuxiliar, 1, 9) & "-" & Mid$(lAuxiliar, 10, 2)
            End If
            If Not IsNumeric(Mid(lAuxiliar, 1, 7)) Or Not IsNumeric(Mid(lAuxiliar, 9, 1)) Then
                Select Case pOpcao
                    Case 1
                        If frmT47V010A.Caption = "Crédito em conta" Then
                            pMensagem = "FGE2076 - " & gM46V999.gPrefixo & "Conta para crédito. " & _
                                    "Número de conta corrente inválido (HSBC - formato NNNNNN-NN)."
                        Else
                            pMensagem = "FGE2076 - " & gM46V999.gPrefixo & "Conta para débito. " & _
                                    "Número de conta corrente inválido (HSBC - formato NNNNNN-NN)."
                        End If
                    Case 2
                        pMensagem = "FGE2076 - " & gM46V999.gPrefixo & "Conta do indicador do seguro. " & _
                                    "Número de conta corrente inválido (HSBC - formato NNNNNN-NN)."
                End Select
                pSaida = ""
                Exit Function
            End If
        Case Else       'Outros bancos.
            lPosicao = InStr(lAuxiliar, "-")
            If Len(lAuxiliar) <> 9 Or lPosicao <> 8 Then
                Select Case lPosicao
                    Case 0
                        'Nada a fazer.
                    Case 1
                        If Len(lAuxiliar) = 1 Then
                            lAuxiliar = ""
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 2)
                        End If
                    Case Else
                        If Len(lAuxiliar) = lPosicao Then
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1)
                        Else
                            lAuxiliar = Mid$(lAuxiliar, 1, lPosicao - 1) & _
                                        Mid$(lAuxiliar, lPosicao + 1)
                        End If
                End Select
                If Len(lAuxiliar) < 8 Then
                    lAuxiliar = String(8 - Len(lAuxiliar), "0") & lAuxiliar
                End If
                lAuxiliar = Mid$(lAuxiliar, 1, 7) & "-" & Mid$(lAuxiliar, 8, 1)
            End If
            If Not IsNumeric(Mid(lAuxiliar, 1, 7)) Or Not IsNumeric(Mid(lAuxiliar, 9, 1)) Then
                Select Case pOpcao
                    Case 1
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta para débito. " & _
                                    "Número de conta corrente inválido (formato NNNNNNN-N)."
                    Case 2
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta do indicador do seguro. " & _
                                    "Número de conta corrente inválido (formato NNNNNNN-N)."
                End Select
                pSaida = ""
                Exit Function
            End If
    End Select
    pSaida = lAuxiliar

    gfVerConta = True
    
End Function

Public Function gfExportarParaBaseTemp(pNosso_Numero As String, pNum_Ped As Long, _
                    pNumBloq As Long, pMensagem As String) As Boolean
                    
Dim bdBasDados  As ADODB.Connection 'Base de dados.
Dim bdTemp      As ADODB.Connection 'Base temporária.
Dim flTabela    As clsYasField      'Campos da tabela.
Dim rsTabela    As clsYasRecordSet  'Registro da tabela.
Dim lCount      As Long             'Utilizado no for...next.
Dim lIdeTabela  As String           'Nome da tabela.
Dim lInsert     As String           'Insert
Dim lContinuar  As Boolean          'Indicador se continua o procedimento ou não.
Dim lBanco      As Integer          'Banco
Dim lSelect     As String
Dim lDelete     As String

    gfExportarParaBaseTemp = False
    On Error GoTo ControleErro
    
    '1. Obtém registros da tabela.
    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, bdTemp, _
            pMensagem) = False Then Exit Function
    
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, _
            pMensagem) = False Then Exit Function
    
    For lCount = 1 To 11
        lContinuar = True
        Select Case lCount
            Case 1
                lIdeTabela = "TAB_PED"
            Case 2
                lIdeTabela = "TAB_PED_CORR"
            Case 3
                lIdeTabela = "TAB_PED_COSS"
            Case 4
                lIdeTabela = "TAB_PED_ENDER"
            Case 5
                lIdeTabela = "TAB_CTRL_EMIS"
            Case 6
                lIdeTabela = "TAB_PED_LOC"
            Case 7
                lIdeTabela = "TAB_PED_CLAU"
            Case 8
                lIdeTabela = "TAB_PED_COBERT"
            Case 9
                lIdeTabela = "TAB_PED_INSP"
            Case 10
                lIdeTabela = "TAB_PED_VENC"
            Case 11
                lIdeTabela = "TAB_PED_AUTO"
        End Select
        'Excluo tudo o que está na base temporária.
        If lCount = 1 Then
            lDelete = " DELETE FROM  Tab_Erro"
            If gM46V999.gfExeSQL(bdTemp, lDelete, pMensagem) <> 0 Then
                Call gM46V999.gpFecha2(bdTemp)
                Exit Function
            End If
        End If
        lDelete = " DELETE FROM  " & lIdeTabela
        If gM46V999.gfExeSQL(bdTemp, lDelete, pMensagem) <> 0 Then
            Call gM46V999.gpFecha2(bdTemp)
            Exit Function
        End If
        If lIdeTabela <> "TAB_PED_COBERT" Then
            lSelect = "SELECT * FROM " & lIdeTabela & " WHERE Nosso_Numero = '" & pNosso_Numero & "'"
        Else
            lSelect = "SELECT Tab_Ped_Cobert.*, "
            lSelect = lSelect & " Tab_Ped_Cobert_Compl.Alterar_Cobert"
            lSelect = lSelect & " From "
            lSelect = lSelect & "Tab_Ped_Cobert as Tab_Ped_Cobert INNER JOIN     " & "Tab_Ped_Cobert_Compl as Tab_Ped_Cobert_Compl ON"
            lSelect = lSelect & " (Tab_Ped_Cobert.COD_COBERT   = Tab_Ped_Cobert_Compl.COD_COBERT) AND"
            lSelect = lSelect & " (Tab_Ped_Cobert.NUM_ITEM     = Tab_Ped_Cobert_Compl.NUM_ITEM) AND"
            lSelect = lSelect & " (Tab_Ped_Cobert.NOSSO_NUMERO = Tab_Ped_Cobert_Compl.NOSSO_NUMERO)"
            lSelect = lSelect & " Where Tab_Ped_Cobert.Nosso_Numero = '" & pNosso_Numero & "'"
        End If
        If gM46V999.gfObtRegistro(bdBasDados, lSelect, rsTabela, pMensagem) = False Then
            Call gM46V999.gpFecha2(bdTemp)
            Call gM46V999.gpFecha2(bdBasDados)
            Exit Function
        End If
        If rsTabela.EOF = True Then
            lContinuar = False
            If lIdeTabela = "TAB_PED" Then
                pMensagem = "FNL0003 - " & gM46V999.gPrefixo & "Pedido não localizado (" & pNosso_Numero & _
                                 " - exportar para base temp.)."
                Exit Function
            End If
        End If
                
        If lContinuar = True Then
            '2. Grava nova cotação/proposta.
            While Not rsTabela.EOF
                lInsert = "INSERT INTO " & lIdeTabela & " ("
                For Each flTabela In rsTabela.Fields
                    Select Case UCase(flTabela.Name)
                        'Campos que não serão exportados.
'Ficha 1000307
                        Case "NOSSO_NUMERO", "COD_UNICIDADE", "ULT_TARIFA", _
                             "DAT_PRI_CALCULO", "SEXO", "NOM_EMPRESA", _
                             "DAT_EMIS_DOC", "ORG_EXP", "NUM_DOC", _
                             "COD_DOC", "COD_RAMO_ATIV", "TIP_PESSOA", _
                             "COD_SEQUENCIA", "COD_PROTOCOLO", "TIP_EMISSAO", _
                             "DESCRICAO", "NUM_APOL_ANTERIOR", "NUM_ITEM_ANTERIOR", _
                             "IND_SINISTRO", "DIG_BLOQUETO", "COD_DIG_AGEN_BAS_COR", _
                             "ALTERAR_COBERT", "COD_VER_UNICI", "DIG_COD_SEGURADO", _
                             "INSP_OBR", "INSP_AGEND", "DIG_BENS", "COB_BENS", _
                             "VAL_RENDAFATU", "PER_DESC_LMI", "LMI_UNICO", "NUM_COTAC_LMI", _
                             "COD_REG_CONS", "NUM_REG_CONS", "TIP_COBERT", _
                             "REN_CON_SEGURADORA", "REN_CON_APOLICE", "REN_CON_ITEM", _
                             "IND_CONDGERAIS", "IND_REENTRADA", "NOSSO_NUMERO_ORIGINAL", _
                             "IND_SIN_INSPEC", "COD_DIZERITEM"
                        Case "COD_PRODUTO"
                            If lIdeTabela = "TAB_PED_LOC" Then
                                lInsert = lInsert & flTabela.Name & ", "
                            End If
                        Case "NUM_PI"
                            If lIdeTabela = "TAB_CTRL_EMIS" Then
                                lInsert = lInsert & flTabela.Name & ", "
                            End If
                        Case Else
                            lInsert = lInsert & flTabela.Name & ", "
                    End Select
                Next flTabela
                lInsert = Left$(lInsert, Len(lInsert) - 2) & ") VALUES ("
                For Each flTabela In rsTabela.Fields
                    If UCase(flTabela.Name) = "VAL_PR_ANTEC" Then
                        lInsert = lInsert
                    End If
                    Select Case UCase(flTabela.Name)
                        'Campos que não serão exportados.
                        'Ficha 1000307
                        Case "NOSSO_NUMERO", "COD_UNICIDADE", "ULT_TARIFA", _
                             "DAT_PRI_CALCULO", "SEXO", "NOM_EMPRESA", _
                             "DAT_EMIS_DOC", "ORG_EXP", "NUM_DOC", _
                             "COD_DOC", "COD_RAMO_ATIV", "TIP_PESSOA", _
                             "COD_SEQUENCIA", "COD_PROTOCOLO", "TIP_EMISSAO", _
                             "DESCRICAO", "NUM_APOL_ANTERIOR", "NUM_ITEM_ANTERIOR", _
                             "IND_SINISTRO", "DIG_BLOQUETO", "COD_DIG_AGEN_BAS_COR", _
                             "ALTERAR_COBERT", "COD_VER_UNICI", "DIG_COD_SEGURADO", _
                             "INSP_OBR", "INSP_AGEND", "DIG_BENS", "COB_BENS", _
                             "VAL_RENDAFATU", "PER_DESC_LMI", "LMI_UNICO", "NUM_COTAC_LMI", _
                             "COD_REG_CONS", "NUM_REG_CONS", "TIP_COBERT", _
                             "REN_CON_SEGURADORA", "REN_CON_APOLICE", "REN_CON_ITEM", _
                             "IND_CONDGERAIS", "IND_REENTRADA", "NOSSO_NUMERO_ORIGINAL", _
                             "IND_SIN_INSPEC", "COD_DIZERITEM"

                        '[FRH]Ficha:0801295 BYPASS PontSyas
                        'Novas opções "Tipo de Emissão=Conforme cadastro": "O" e "E" não estão
                        'contempladas na PontSyas, auto-ajusta COD_ENVI_DOC da base-temp para "S".
                        Case "COD_ENVI_DOC"
                            lInsert = lInsert & "'S', "
                            
                        Case "COD_PRODUTO"
                            If lIdeTabela = "TAB_PED_LOC" Then
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case "NUM_PI"
                            If lIdeTabela = "TAB_CTRL_EMIS" Then
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case "COD_TIP_DOC" ' AQUI
                            If gM46V111.gCod_Dizer = 0 Then
                                lInsert = lInsert & "0,"
                            Else
                                lInsert = lInsert & "3,"
                            End If
                        Case "PER_TAXA"
                            If lIdeTabela = "TAB_PED_COBERT" Then
                                If gM46V999.gfPreenchido(rsTabela("Alterar_Cobert")) Then
                                    lInsert = lInsert & gM46V999.gfForValor(1, flTabela.Value) & ", "
                                Else
                                    lInsert = lInsert & "0,"
                                End If
                            Else
                                lInsert = lInsert & "0,"
                            End If
                        '"VAL_PR"  SIMONY
                        Case "VAL_PR_ANU", "VAL_PR_RESG", "VAL_PR_RESG_ANU", "VAL_PR_TARIF", "VAL_PR_TARIF_ANU"
                            lInsert = lInsert & "0,"
                        Case "COD_BCO_COBR"
                            lInsert = lInsert & flTabela.Value & ", "
                            lBanco = flTabela.Value
                        Case "NUM_CTA_BAS_COR"
                            If Len(CStr(flTabela.Value)) <> 1 Then
                                Select Case lBanco
                                    Case 237, 341
                                        lInsert = lInsert & flTabela.Value & ", "
                                    Case 347, 356
                                        lInsert = lInsert & Left(Format(flTabela.Value, "0000000000"), 7) & Right(flTabela.Value, 1) & ", "
                                    Case Else
                                        lInsert = lInsert & Left(flTabela.Value, Len(CStr(flTabela.Value)) - 1) & ", "
                                End Select
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case "COD_AGEN_BAS_COR"
                            If Len(CStr(flTabela.Value)) <> 1 Then
                                Select Case lBanco
                                    Case 341, 479, 347, 237, 356
                                        lInsert = lInsert & flTabela.Value & ", "
                                    Case Else
                                        lInsert = lInsert & Left(flTabela.Value, Len(CStr(flTabela.Value)) - 1) & ", "
                                End Select
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case "NUM_FONE"
                            lInsert = lInsert & Right(flTabela.Value, 7) & ", "
                        Case "NUM_BLOQ"
                            lInsert = lInsert & pNumBloq & ", "
                        Case Else
                            Select Case flTabela.Type_
                                Case adBigInt, adBinary, adInteger, adLongVarBinary, adSmallInt, _
                                     adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, _
                                     adUnsignedTinyInt                          'Numérico - inteiro.
                                    If IsNull(flTabela.Value) Then
                                        lInsert = lInsert & "0, "
                                    Else
                                        lInsert = lInsert & gM46V999.gfForValor(1, flTabela.Value) & ", "
                                    End If
                                Case adCurrency, adSingle, adDouble, adNumeric             'Numérico - com vírgula.
                                    If IsNull(flTabela.Value) Then
                                        lInsert = lInsert & "0, "
                                    Else
                                        lInsert = lInsert & gM46V999.gfForValor(1, flTabela.Value) & ", "
                                    End If
                                Case adChar, adVarChar, adLongVarChar                           'Texto.
                                    If gM46V999.gfPreenchido(flTabela.Value) Then
                                        lInsert = lInsert & "'" & gM46V999.gfForAspas(flTabela.Value) & "', "
                                    Else
                                        lInsert = lInsert & "' ', "
                                    End If
                                Case adBoolean                                  'Sim/Não.
                                    If IsNull(flTabela.Value) Then
                                        lInsert = lInsert & "TRUE, "
                                    Else
                                        lInsert = lInsert & flTabela.Value & ", "
                                    End If
                                Case Else
                                    Call gM46V999.gpFecha2(bdBasDados)
                                    pMensagem = "FTE0008 - Tabela " & lIdeTabela & " com tipo do campo " & _
                                                flTabela.Name & " diferente dos previstos."
                                    Exit Function
                            End Select
                    End Select
                Next flTabela
                If Mid(lInsert, Len(lInsert), 1) = "," Then
                    lInsert = Left(lInsert, Len(lInsert) - 1) & ")"
                Else
                    lInsert = Left(lInsert, Len(lInsert) - 2) & ")"
                End If
                If gM46V999.gfExeSQL(bdTemp, lInsert, pMensagem) <> 0 Then
                    Call gM46V999.gpFecha2(bdBasDados)
                    Exit Function
                End If
                rsTabela.MoveNext
            Wend
        End If
    Next lCount
    Call gM46V999.gpFecha2(bdBasDados)
    Call gM46V999.gpFecha2(bdTemp)
    gfExportarParaBaseTemp = True
    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "gfExportarParaBaseTemp ERRO - " & Err & " " & Err.Description)
End Function


Public Sub gpMensagem(ByVal pMensagem As String)
    'Procedure: posiciona mensagem no FlexGrid.

    'Parâmetro de entrada...pMensagem...Mensagem.
    
    ''1. Se o form T47V017A ainda não estiver disponível, abrir e deixar a grade limpa.
    If gM46V999.gfDisponivel("frmT47V017A") = False Then
        Load frmT47V017A
        frmT47V017A.flxMensagens.Rows = 0
    End If

    ''2. Posiciona mensagem na grade.
    With frmT47V017A
        .lblAuxMensagem.Width = .flxMensagens.ColWidth(0) - 300
        .lblAuxMensagem.Caption = Trim$(pMensagem)
        .flxMensagens.Rows = .flxMensagens.Rows + 1
        .flxMensagens.Row = .flxMensagens.Rows - 1
        .flxMensagens.RowHeight(.flxMensagens.Row) = .lblAuxMensagem.Height
        .flxMensagens.Text = .lblAuxMensagem.Caption
    End With

    '3. Grava no arquivo log.
    Call gM46V999.gpGraLog(0, pMensagem)
End Sub

Private Function mpVerCalculoParalelo() As Boolean
Dim lDiretorio  As String
Dim lOutroArq   As String
Dim lArquivo    As Integer
Dim lNomeArquivo    As String
Dim lNomeComputador As String
    On Error GoTo ControleErro
    lDiretorio = Replace(UCase(gM46V999.gEndP00MULTI), "P00MULTI.MDB", "")
    lNomeComputador = gM99V001.GetLoginComputerName()
    lNomeArquivo = lNomeComputador & ".PROCESSANDO"
Retorno:
    lOutroArq = Dir(lDiretorio & "*.PROCESSANDO")
    If gM46V999.gfPreenchido(lOutroArq) Then
        If FileDateTime(lDiretorio & lOutroArq) < (DateAdd("s", -20, Now)) Then
            Kill lDiretorio & lOutroArq
            GoTo Retorno
        Else
            'If lNomeMaq <> Left(lOutroArq, InStr(1, lOutroArq, ".") - 1) Then
                'Call  gM46V999.gpMenErro("Existe outra estação efetuando um cálculo simultâneo.(" & lNomeComputador & "," & Left(lOutroArq, InStr(1, lOutroArq, ".") - 1) & ")")
                mpVerCalculoParalelo = True
            'End If
        End If
    Else
        'crio o arquivo
        lArquivo = FreeFile
        Open lDiretorio & lNomeArquivo For Output As #lArquivo
            Print #lArquivo, "Processando " & Now
         Close #lArquivo
    End If
    
    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(0, "mpVerCalculoParalelo - " & Err.Description)
End Function
Private Sub mpDeletarArqCalculo()
Dim lDiretorio  As String
Dim lNomeArquivo    As String
Dim lNomeComputador As String

    On Error Resume Next
    lDiretorio = Replace(UCase(gM46V999.gEndP00MULTI), "P00MULTI.MDB", "")
    lNomeComputador = gM99V001.GetLoginComputerName()
    lNomeArquivo = lNomeComputador & ".PROCESSANDO"
    Kill lDiretorio & lNomeArquivo
    Err.Clear
    On Error GoTo 0
End Sub
Public Sub gpLimPremios()
    'Procedure: limpa informações de prêmios já calculados.

    '1. Se é carga do form, encerra processo.
    If gCarga = True Then
        Exit Sub
    End If

    Call gpHabSalvar

    Screen.MousePointer = vbDefault
End Sub
Public Function gfCalcData(pData As Long, pDias As Integer) As Long
Dim lAuxPreData As Date
    If pData <> 0 Then
    
        If Right(pData, 4) = "0229" Then
            pData = Left(pData, 4) & "0228"
        End If
        
        If Mid(pData, 7, 2) & "/" & Mid(pData, 5, 2) & "/" & Mid(pData, 1, 4) <> Format(Mid(pData, 7, 2) & "/" & Mid(pData, 5, 2) & "/" & Mid(pData, 1, 4), "dd/mm/yyyy") Then
            lAuxPreData = CDate(Mid(pData, 5, 2) & "/" & Mid(pData, 7, 2) & "/" & Mid(pData, 1, 4))
        Else
            lAuxPreData = CDate(Mid(pData, 7, 2) & "/" & Mid(pData, 5, 2) & "/" & Mid(pData, 1, 4))
        End If
    
        gfCalcData = Format(CDate(lAuxPreData) + pDias, "yyyymmdd")
    End If
    
End Function

Public Sub gpPos025_026()
    'Procedure: posiciona forms T46V123A e T47V026A na posição original.

    Dim lAltura         As Integer      'Altura original.
    Dim lLargura        As Integer      'Largura original.
    Dim lMarEsquerda    As Integer      'Margem esquerda.
    Dim lMarSuperior    As Integer      'Margem superior.

    '1. Calcula margens.
    lAltura = 5000
    lLargura = 3950
    With frmT47V001A
        lMarEsquerda = .Left + .Width - lLargura
        lMarSuperior = .Top + .panAguarde.Top + 500
    End With

    '2. Form T47V025A.
    If gM46V999.gfDisponivel("frmT47V025A") Then
        With frmT47V025A
            .Move lMarEsquerda, lMarSuperior, lLargura, lAltura
            .tabPremios.Tab = 0
        End With
    End If

    '2. Form T47V026A.
    If gM46V999.gfDisponivel("frmT47V026A") Then
        With frmT47V026A
            .Move lMarEsquerda, lMarSuperior, lLargura, lAltura
        End With
    End If
    
End Sub

Public Function gfCalFracionamento(ByVal pQuaParcelas As Byte, ByVal pIniVigencia As String, _
                                   ByVal pTerVigencia As String, _
                                   ByVal pCusEmissao As Single, ByVal pCusParcelado As String, _
                                   ByVal pPreLiquido As Single, pJuros As Variant, _
                                   ByVal pTip_Emissao As Long, ByVal pCod_Iden_Tarf As Long, _
                                   ByRef pParcela As Variant, ByRef pMensagem As String, _
                                   ByVal pProJurIOF As Double, Optional ByVal pCodReg As Integer, _
                                   Optional ByRef pIOF As Single, Optional ByVal pParcMin As Double)
    
    'Função: calcula fracionamento.

    'Parâmetros de entrada...pQuaParcelas.....quantidade de parcelas.
    '                        pIniVigencia.....início de vigência.
    '                        pCusEmissao......custo de emissão.
    '                        pCusParcelado....indicador de custo parcelado: S = Sim ou N = Não.
    '                        pVD_PreLiquido...valor determinado - valor do prêmio líquido.
    '                        pVM_PreLiquido...valor de mercado - valor do prêmio líquido.
    '                        pJurAntxx........juros para pagamento antecipado - xx parcelas.
    '                        pJurPosxx........juros para pagamento postecipado - xx parcelas.
    '                        pCod_Registro.....Código do Registro
    '                        pParcMin.........valor da parcela mínima
    'Parâmetro de saída......pParcela.........matriz com resultados das parcelas.
    '                        pMensagem........mensagem.
    '                        pErro............lista de erros.
    '                        pVD-IOF..........Valor do IOF para valor determinado (opcional).
    '                        pVM-IOF..........Valor do IOF para valor de mercado (opcional).
    Dim lI1             As Byte         'Utilizado com For...Next.
    Dim lI2             As Byte         'Utilizado com For...Next.
    Dim lValorMin       As Double       'Valor mínimo da parcela ou endosso.
    Dim lPremioMin      As Double       'Valor mínimo do prêmio de parcela ou endosso.
    Dim lParcVista      As Integer      'Quantidade de parcelas máxima possível em caso de pagamento antecipado.
    Dim lParcs30dia     As Integer      'Quandidade de parcelas máxima possível em caso de paramento não antecipado.
    Dim lPreTotal       As Double       'Prêmio total.
    Dim lDivCust        As Integer      'Número pelo qual será dividido o custo
    Dim lValor          As Double       'Prêmio Liquido.
    Dim lIOF            As Double       'Valor IOF.
    Dim lPer_Juros_1    As Double       'Juros da 1ª parcela.
    Dim lPer_Juros_2    As Double       'Juros da 2ª parcela.
    Dim lPer_Juros_3    As Double       'Juros da 3ª parcela.
    Dim lPer_Juros_4    As Double       'Juros da 4ª parcela.
    Dim lPer_Juros_5    As Double       'Juros da 5ª parcela.
    Dim lPer_Juros_6    As Double       'Juros da 6ª parcela.
    Dim lPer_Juros_7    As Double       'Juros da 7ª parcela.
    Dim lPer_Juros_8    As Double       'Juros da 8ª parcela.
    Dim lPer_Juros_9    As Double       'Juros da 9ª parcela.
    Dim lPer_Juros_10   As Double       'Juros da 10ª parcela.
    Dim lPer_Juros_11   As Double       'Juros da 11ª parcela.
    Dim lPer_Juros_12   As Double       'Juros da 12ª parcela.
    Dim lPer_Juros_0_1  As Double       'Juros da 1ª parcela antecipada.
    Dim lPer_Juros_0_2  As Double       'Juros da 2ª parcela antecipada.
    Dim lPer_Juros_0_3  As Double       'Juros da 3ª parcela antecipada.
    Dim lPer_Juros_0_4  As Double       'Juros da 4ª parcela antecipada.
    Dim lPer_Juros_0_5  As Double       'Juros da 5ª parcela antecipada.
    Dim lPer_Juros_0_6  As Double       'Juros da 6ª parcela antecipada.
    Dim lPer_Juros_0_7  As Double       'Juros da 7ª parcela antecipada.
    Dim lPer_Juros_0_8  As Double       'Juros da 8ª parcela antecipada.
    Dim lPer_Juros_0_9  As Double       'Juros da 9ª parcela antecipada.
    Dim lPer_Juros_0_10 As Double       'Juros da 10ª parcela antecipada.
    Dim lPer_Juros_0_11 As Double       'Juros da 11ª parcela antecipada.
    Dim lPer_Juros_0_12 As Double       'Juros da 12ª parcela antecipada.
    Dim lobjstcJuros    As stcA46V711B
    Dim lDatTerVig      As Double
    Dim lUltParcela     As Double
    Dim lTemAjusteVenc  As Boolean
    
    gfCalFracionamento = False
  
    
    lPer_Juros_1 = pJuros
    lPer_Juros_2 = pJuros
    lPer_Juros_3 = pJuros
    lPer_Juros_4 = pJuros
    lPer_Juros_5 = pJuros
    lPer_Juros_6 = pJuros
    lPer_Juros_7 = pJuros
    lPer_Juros_8 = pJuros
    lPer_Juros_9 = pJuros
    lPer_Juros_10 = pJuros
    lPer_Juros_11 = pJuros
    lPer_Juros_12 = pJuros
    lPer_Juros_0_1 = pJuros
    lPer_Juros_0_2 = pJuros
    lPer_Juros_0_3 = pJuros
    lPer_Juros_0_4 = pJuros
    lPer_Juros_0_5 = pJuros
    lPer_Juros_0_6 = pJuros
    lPer_Juros_0_7 = pJuros
    lPer_Juros_0_8 = pJuros
    lPer_Juros_0_9 = pJuros
    lPer_Juros_0_10 = pJuros
    lPer_Juros_0_11 = pJuros
    lPer_Juros_0_12 = pJuros


    'Incluido Simony
    If pCodReg = 2 Or _
       pCodReg = 3 Or _
       pCodReg = 102 Or _
       pCodReg = 103 Or _
       pCodReg = 202 Or _
       pCodReg = 203 Or _
       pQuaParcelas = 0 Then
        pProJurIOF = 0
        pIOF = 0
    End If
    'Fim
    
    '2. Prepara planilha de resultados.
    ReDim pParcela(12, 4)
    For lI1 = 1 To 12
        For lI2 = 1 To 4
            pParcela(lI1, lI2) = 0
        Next lI2
    Next lI1

    If Left(frmT47V016A.cboCodReg.Text, 3) = "020" Or Left(frmT47V016A.cboCodReg.Text, 3) = "120" Or Left(frmT47V016A.cboCodReg.Text, 3) = "121" Or Left(frmT47V016A.cboCodReg.Text, 3) = "122" Then
        gM46V999.gProPremioMinimo = 0
        gM46V999.gProCusApolice = 0
    End If


    '3. Calcula fracionamento e posiciona resultados na planilha.
    lValor = pPreLiquido
    If lValor <> 0 Then
                
        
        lValorMin = pParcMin
        lPremioMin = gM46V999.gProPremioMinimo
        If lValor >= 0 Then
            If lValor < lPremioMin Then
                lValor = 0
                lPreTotal = 0
                pCusEmissao = 0
                lIOF = 0
            End If
        ElseIf lValor > (lPremioMin * -1) Then
                lValor = 0
                lPreTotal = 0
                pCusEmissao = 0
                lIOF = 0
        End If
        
        ' CR96910
        If lValor < 0 Then
            pProJurIOF = 0
        End If
        
        lIOF = Round((lValor + pCusEmissao) * (pProJurIOF / 100), 2)
        lPreTotal = lValor + pCusEmissao + lIOF
        
        '2. Calcula valores das parcelas.
        'c) Posiciona resultados.
        '===>>> Parcelas.
        If pCusParcelado = "S" Then
            'a) Considerando custo de emissão parcelado.
            pParcela(1, 1) = gM46V999.gfTrunca((lValor + pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 2, lPer_Juros_2), 2)) + _
                                       (pCusEmissao / 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 2) = pParcela(2, 1)
            pParcela(3, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 3, lPer_Juros_3), 2)) + _
                                       (pCusEmissao / 3)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 2) = pParcela(3, 1)
            pParcela(4, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 4, lPer_Juros_4), 2)) + _
                                       (pCusEmissao / 4)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 2) = pParcela(4, 1)
            pParcela(5, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 5, lPer_Juros_5), 2)) + _
                                       (pCusEmissao / 5)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 2) = pParcela(5, 1)
            pParcela(6, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 6, lPer_Juros_6), 2)) + _
                                       (pCusEmissao / 6)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 2) = pParcela(6, 1)
            pParcela(7, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 7, lPer_Juros_7), 2)) + _
                                       (pCusEmissao / 7)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 2) = pParcela(7, 1)
            pParcela(8, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 8, lPer_Juros_8), 2)) + _
                                       (pCusEmissao / 8)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 2) = pParcela(8, 1)
            pParcela(9, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 9, lPer_Juros_9), 2)) + _
                                       (pCusEmissao / 9)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 2) = pParcela(9, 1)
            pParcela(10, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 10, lPer_Juros_10), 2)) + _
                                        (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 2) = pParcela(10, 1)
            pParcela(11, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 11, lPer_Juros_11), 2)) + _
                                        (pCusEmissao / 11)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 2) = pParcela(11, 1)
            pParcela(12, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 12, lPer_Juros_12), 2)) + _
                                        (pCusEmissao / 12)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 2) = pParcela(12, 1)
            
            
            pParcela(1, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 1, lPer_Juros_0_1), 2)) + _
                                       (pCusEmissao / 1)) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 2, lPer_Juros_0_2), 2)) + _
                                       (pCusEmissao / 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 4) = pParcela(2, 3)
            pParcela(3, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 3, lPer_Juros_0_3), 2)) + _
                                       (pCusEmissao / 3)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 4) = pParcela(3, 3)
            pParcela(4, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 4, lPer_Juros_0_4), 2)) + _
                                       (pCusEmissao / 4)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 4) = pParcela(4, 3)
            pParcela(5, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 5, lPer_Juros_0_5), 2)) + _
                                       (pCusEmissao / 5)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 4) = pParcela(5, 3)
            pParcela(6, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 6, lPer_Juros_0_6), 2)) + _
                                       (pCusEmissao / 6)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 4) = pParcela(6, 3)
            pParcela(7, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 7, lPer_Juros_0_7), 2)) + _
                                       (pCusEmissao / 7)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 4) = pParcela(7, 3)
            pParcela(8, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 8, lPer_Juros_0_8), 2)) + _
                                       (pCusEmissao / 8)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 4) = pParcela(8, 3)
            pParcela(9, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 9, lPer_Juros_0_9), 2)) + _
                                       (pCusEmissao / 9)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 4) = pParcela(9, 3)
            pParcela(10, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 10, lPer_Juros_0_10), 2)) + _
                                        (pCusEmissao / 10)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 4) = pParcela(10, 3)
            
            pParcela(11, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 11, lPer_Juros_0_11), 2)) + _
                                        (pCusEmissao / 11)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 4) = pParcela(11, 3)
            pParcela(12, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 12, lPer_Juros_0_12), 2)) + _
                                        (pCusEmissao / 12)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 4) = pParcela(12, 3)
        Else
            'b) Considerando custo de emissão NÃO parcelado.
            pParcela(1, 1) = gM46V999.gfTrunca((lValor + pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 2, lPer_Juros_2), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 2) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(1, 2, lPer_Juros_2), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 3, lPer_Juros_3), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 2) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(1, 3, lPer_Juros_3), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 4, lPer_Juros_4), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 2) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(1, 4, lPer_Juros_4), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 5, lPer_Juros_5), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 2) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(1, 5, lPer_Juros_5), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 6, lPer_Juros_6), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 2) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(1, 6, lPer_Juros_6), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 7, lPer_Juros_7), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 2) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(1, 7, lPer_Juros_7), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 8, lPer_Juros_8), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 2) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(1, 8, lPer_Juros_8), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 9, lPer_Juros_9), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 2) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(1, 9, lPer_Juros_9), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 10, lPer_Juros_10), 2)) + _
                                         pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 2) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                               mfCoeJuros(1, 10, lPer_Juros_10), 2)) * (1 + pProJurIOF / 100), 2)
            
            pParcela(11, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 11, lPer_Juros_11), 2)) + _
                                         pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 2) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                               mfCoeJuros(1, 11, lPer_Juros_11), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 1) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(1, 12, lPer_Juros_12), 2)) + _
                                         pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 2) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                               mfCoeJuros(1, 12, lPer_Juros_12), 2)) * (1 + pProJurIOF / 100), 2)
                                               
            pParcela(1, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 1, lPer_Juros_0_1), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 2, lPer_Juros_0_2), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(2, 4) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(0, 2, lPer_Juros_0_2), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 3, lPer_Juros_0_3), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(3, 4) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(0, 3, lPer_Juros_0_3), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 4, lPer_Juros_0_4), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(4, 4) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(0, 4, lPer_Juros_0_4), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 5, lPer_Juros_0_5), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(5, 4) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(0, 5, lPer_Juros_0_5), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 6, lPer_Juros_0_6), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(6, 4) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(0, 6, lPer_Juros_0_6), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 7, lPer_Juros_0_7), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(7, 4) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(0, 7, lPer_Juros_0_7), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 8, lPer_Juros_0_8), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(8, 4) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(0, 8, lPer_Juros_0_8), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 9, lPer_Juros_0_9), 2)) + _
                                        pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(9, 4) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(0, 9, lPer_Juros_0_9), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 10, lPer_Juros_0_10), 2)) + _
                                         pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(10, 4) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(0, 10, lPer_Juros_0_10), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 11, lPer_Juros_0_11), 2)) + _
                                         pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(11, 4) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(0, 11, lPer_Juros_0_11), 2)) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 3) = gM46V999.gfTrunca(((gM46V999.gfTrunca(lValor * mfCoeJuros(0, 12, lPer_Juros_0_12), 2)) + _
                                         pCusEmissao) * (1 + pProJurIOF / 100), 2)
            pParcela(12, 4) = gM46V999.gfTrunca((gM46V999.gfTrunca(lValor * _
                                              mfCoeJuros(0, 12, lPer_Juros_0_12), 2)) * (1 + pProJurIOF / 100), 2)
        End If
        '===>>> IOF.
        If Not IsMissing(pIOF) Then
            pIOF = lIOF
        End If
    
    
        'forma de pagamento de acordo com o plano se estiver cadastrado
        If Not gRst_RegrasPlano Is Nothing Then
            If gRst_RegrasPlano.RecordCount > 0 Then gRst_RegrasPlano.MoveFirst
            lI2 = 1
            For lI1 = 20 To 40 Step 2
                'Parcela à Vista
                gRst_RegrasPlano.Filter = "Cod_Regra = " & lI1
                If gRst_RegrasPlano.RecordCount = 0 Then
                    pParcela(lI2, 1) = 0
                    pParcela(lI2, 2) = 0
                End If
                gRst_RegrasPlano.Filter = ""
            
                'Parcela à 30 Dias
                gRst_RegrasPlano.Filter = "Cod_Regra = " & lI1 + 1
                If gRst_RegrasPlano.RecordCount = 0 Then
                    pParcela(lI2, 3) = 0
                    pParcela(lI2, 4) = 0
                End If
                gRst_RegrasPlano.Filter = ""
                lI2 = lI2 + 1
            Next lI1
        End If
    
        '===>>> IOF.
        If gM46V999.gfDisponivel("frmT47V001A") Then
            If frmT47V001A.objstcPedido.ColecaoPedVenc.Count > 0 Then
                lTemAjusteVenc = True
            End If
        End If
        If (lParcVista = 0 And lParcs30dia = 0) Or (gM46V999.gTipoAcesso = DIGITACAO And lTemAjusteVenc = True) Then
            lParcVista = 12
            lParcs30dia = 12
        End If
        If pPreLiquido > 0 Then
            If pParcela(1, 1) < lValorMin Then
                If pParcela(1, 1) < lPremioMin Then
                    pParcela(1, 1) = 0
                End If
            End If
            For lI1 = 2 To 12
                If pParcela(lI1, 2) < lValorMin Or lParcVista < lI1 Then
                    pParcela(lI1, 2) = 0
                    pParcela(lI1, 1) = 0
                End If
                If pParcela(lI1, 4) < lValorMin Or lParcs30dia < lI1 Then
                    pParcela(lI1, 4) = 0
                    pParcela(lI1, 3) = 0
                End If
            Next lI1
        Else
            pParcela(1, 1) = pPreLiquido
            For lI1 = 2 To 12
                pParcela(lI1, 2) = 0
                pParcela(lI1, 1) = 0
                pParcela(lI1, 4) = 0
                pParcela(lI1, 3) = 0
            Next lI1
        End If
    End If
    gfCalFracionamento = True
End Function

Private Function mfCoeJuros(ByVal pTipPagamento As Byte, ByVal pQuaParcelas As Single, _
                            ByVal pJuros As Single) As Single
    'Função #007: calcula coeficiente de juros.

    'Parâmetros de entrada...pTipPagamento...0 = pagamento postecipado e 1 = pagamento antecipado.
    '                        pQuaParcelas....quantidade de parcelas.
    '                        pJuros..........juros.

    Dim lDesPagamento   As String   'Descrição do pagamento.

    If pJuros = 0 Then
        If pQuaParcelas = 0 Then
            mfCoeJuros = 1
        Else
            mfCoeJuros = 1 / pQuaParcelas
        End If
    Else
        Select Case pTipPagamento
            Case 0  'Pagamento postecipado.
                mfCoeJuros = gM46V999.gfTrunca(((pJuros / 100) / (1 - (1 / (1 + (pJuros / 100)) ^ _
                                        pQuaParcelas))), 4)
            Case 1  'Pagamento antecipado.
                mfCoeJuros = gM46V999.gfTrunca(((pJuros / 100) / (1 - (1 / (1 + (pJuros / 100)) ^ _
                                        pQuaParcelas))) * (1 / (1 + (pJuros / 100))), 4)
        End Select
    End If

    If pTipPagamento = 0 Then
        lDesPagamento = "0 (pagamento postecipado)"
    Else
        lDesPagamento = "1 (pagamento antecipado)"
    End If
    
End Function


Public Function gfImpDocumento(ByVal pOpcao As Byte, ByVal pFormulario As String, _
                                ByVal pImagem As String, ByVal pobjstcPed As stcA46V702B, _
                                pobjstcPedLoc As stcA46V708B, pPagDe As Integer, _
                                pPagAte As Integer, pNomeArqIST As String) As Boolean
    'Função: imprime documento.

    'Parâmetros de entrada...pOpcao...........0 = visualizar impressão.
    '                                         1 = imprimir.
    '                                         2 = criar arquivo imagem.
    '                        pFormulario......Nome do formulário.
    '                        pImagem..........Nome para o arquivo imagem a ser criado.
    '                        lObjP0042200.....Registros de item.
    Dim lVerifOverf As Boolean
    Dim lI          As Integer
    Dim lArqDados   As String           'Endereço do arquivo de dados.
    Dim lEndConfig  As String           'Endereço do arquivo de configuração.
    Dim lEndImagem  As String           'Endereço do arquivo de imagem.
    Dim lNumArquivo As Integer          'Número do arquivo.
    Dim lRefConfig  As String           'Número de referência do arquivo de configuração.
    Dim lRefImagem  As String           'Número de referência do arquivo de imagem.
    Dim lResultado  As Boolean          'Resultado das funções de criação de dados.
    Dim lPossuiObs  As Boolean          'Possui observação?
    Dim lobjstcLoc  As stcA46V708B      'Objeto de itens.
    Dim lArrayArq
    Dim lCount      As Integer
    Dim lPos        As Integer
    Dim lArqIST     As Boolean
    Dim lRegistro  As String
    Dim lNumArqIST As Integer
        
Retorno:

    lArqIST = False
    If pOpcao = 1 And (gM46V999.PerfilAcesso = e_Perfil_DIG_IST Or gM46V999.PerfilAcesso = e_Perfil_DIG_Yas) And Trim(pNomeArqIST) <> "" Then
        lArqIST = True
    End If
    
    gfImpDocumento = False
    mNomForm = pFormulario
    '1. carrego o endereço do formulário.
    Select Case UCase$(pFormulario)
        Case "MULTCOT1"
            lEndConfig = gM46V999.gEndConfCot1
            lEndImagem = gM46V999.gEndImagCot1
        Case "MULTCOT2"
            lEndConfig = gM46V999.gEndConfCot2
            lEndImagem = gM46V999.gEndImagCot2
        Case "MULTCOT3"
            lEndConfig = gM46V999.gEndConfCot3
            lEndImagem = gM46V999.gEndImagCot3
        Case "MULTPRO1"
            lEndConfig = gM46V999.gEndConfPro1
            lEndImagem = gM46V999.gEndImagPro1
        Case "MULTPRO2"
            lEndConfig = gM46V999.gEndConfPro2
            lEndImagem = gM46V999.gEndImagPro2
        Case "MULTPRO3"
            lEndConfig = gM46V999.gEndConfPro3
            lEndImagem = gM46V999.gEndImagPro3
        Case "MULTPRO4"
            lEndConfig = gM46V999.gEndConfPro4
            lEndImagem = gM46V999.gEndImagPro4
        Case "L28V009A"
            lEndConfig = gM46V999.gEndConfDeb
            lEndImagem = gM46V999.gEndImagDeb
        Case "MULTPROB"
            lEndConfig = gM46V999.gEndConfProB
            lEndImagem = gM46V999.gEndImagProB
        Case "MULTPROC"
            lEndConfig = gM46V999.gEndConfProC
            lEndImagem = gM46V999.gEndImagProC
        Case "MULTPROD"
            lEndConfig = gM46V999.gEndConfProD
            lEndImagem = gM46V999.gEndImagProD
        Case "MULTCOTA"
            lEndConfig = gM46V999.gEndConfCotA
            lEndImagem = gM46V999.gEndImagCotA
        Case "MULTCOTC"
            lEndConfig = gM46V999.gEndConfCotC
            lEndImagem = gM46V999.gEndImagCotC
        Case "ROUBOCOT1"
            lEndConfig = gM46V999.gEndConfRob1
            lEndImagem = gM46V999.gEndImagRob1
        Case "ROUBOCOT2"
            lEndConfig = gM46V999.gEndConfRob2
            lEndImagem = gM46V999.gEndImagRob2
        Case "ROUBOCOT3"
            lEndConfig = gM46V999.gEndConfRob3
            lEndImagem = gM46V999.gEndImagRob3
        Case "ROUBOCOT4"
            lEndConfig = gM46V999.gEndConfRob4
            lEndImagem = gM46V999.gEndImagRob4
        Case "ROUBOCOT5"
            lEndConfig = gM46V999.gEndConfRob5
            lEndImagem = gM46V999.gEndImagRob5
        Case "COT_LMI"
            lEndImagem = gM46V999.gEndImagCotLMI
        Case "REL_LMI"
            lEndImagem = gM46V999.gEndImagRelLMI
        Case "MULTLMIR"
            lEndImagem = gM46V999.gEndImagESPRelLMI
        Case "MULTLMIE"
            lEndImagem = gM46V999.gEndImagESPLMI
        Case Else
            Call gM46V999.gpGraLog(1, "FTE0080 - Código de formulário desconhecido: " & pFormulario & ".")
            Exit Function
    End Select
    
    '2. Verifica arquivo de configuração.
    If Not gM46V999.gfPreenchido(Dir(lEndConfig)) Then
        Call gM46V999.gpGraLog(1, "FTE0081 - Arquivo não localizado: " & lEndConfig & ".")
        Exit Function
    End If

    '3. Verifica arquivo de imagem.
    If Not gM46V999.gfPreenchido(Dir(lEndImagem)) Then
        Call gM46V999.gpGraLog(1, "FTE0081 - Arquivo não localizado: " & lEndImagem & ".")
        Exit Function
    End If
    '5. Informa endereços dos arquivos de impressão.
    Call gM46V999.gpGraLog(0, "AGE0064 - " & gM46V999.gPrefixo & "Arquivos utilizados: " & lEndConfig & " e " & _
                     lEndImagem & ".")

    '6. Cria arquivo de dados no drive C.
    If gfIsMetaframe() Then
        lArqDados = "L:\SYAS.DDD"
        mEndArq = lArqDados
        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & lArqDados)
    Else
        lArqDados = gM46V999.gAppPath & "SYAS.DDD"
        mEndArq = lArqDados
    End If
    
    lNumArquivo = FreeFile
    
    On Error Resume Next
    Open lArqDados For Output Access Write Lock Read Write As #lNumArquivo
    Select Case Err
        Case 0
            'Não houve erro.
        Case 70
            Call gM46V999.gpGraLog(1, "FTE0083 - " & gM46V999.gPrefixo & _
                             "Erro ao imprimir documento (usuário não esperou concluir emissão do " & _
                             "documento - sair do sistema e reiniciar processo).")
            Exit Function
        Case Else
            Call gM46V999.gpGraLog(1, "FTE0084 - " & gM46V999.gPrefixo & _
                             "Erro ao imprimir documento (criação do arquivo de dados - " & Err & _
                             " - " & Error & ").")
            Exit Function
    End Select
    
    lListaArq = lArqDados & "|" & lEndConfig & "|" & lEndImagem & ";"

    On Error GoTo 0
    '7. Grava registros no arquivo de dados.
    Select Case pFormulario
    End Select
    
    Close #lNumArquivo
    
    If lVerifOverf = False Or lI = 0 Then
        lArrayArq = Split(lListaArq, ";")
        lPos = InStr(1, lArrayArq(0), "|")
        lArqDados = Left(lArrayArq(0), lPos - 1)
        lEndConfig = Mid(lArrayArq(0), lPos + 1, InStr(lPos + 1, lArrayArq(0), "|") - 1 - lPos)
        lPos = InStr(lPos + 1, lArrayArq(0), "|")
        lEndImagem = Mid(lArrayArq(0), lPos + 1)
        lCount = 1
    Else
        lPos = InStr(1, lArrayArq(lI), "|")
        lArqDados = Left(lArrayArq(lI), lPos - 1)
        lEndConfig = Mid(lArrayArq(lI), lPos + 1, InStr(lPos + 1, lArrayArq(lI), "|") - 1 - lPos)
        lPos = InStr(lPos + 1, lArrayArq(lI), "|")
        lEndImagem = Mid(lArrayArq(lI), lPos + 1)
        Call gM46V999.gpGraLog(0, lI & "  -  " & lArqDados & "  -  " & lEndConfig & "  -  " & lEndImagem)
        pImagem = Replace(pImagem, ".WMF", lI & ".WMF")
    End If

ProximoForm:

    If lResultado = False Then
        Kill lArqDados
        If pOpcao = 2 And pFormulario = "MULTCOTA" Then
            gfImpDocumento = True
        End If
        Exit Function
    End If
    
    'Gerar arquivo para IST
    If lArqIST Then
        lNumArquivo = FreeFile
        Open lArqDados For Input Access Read As #lNumArquivo
        lNumArqIST = FreeFile
        If Trim(Dir(pNomeArqIST)) <> "" Then
            Open pNomeArqIST For Append Shared As #lNumArqIST
        Else
            Open pNomeArqIST For Output As #lNumArqIST
        End If
        On Error GoTo 0
        Do While Not EOF(lNumArquivo)
            Line Input #lNumArquivo, lRegistro
            Print #lNumArqIST, lRegistro
        Loop
        Close #lNumArquivo
        Close #lNumArqIST
        Kill lArqDados
        'GoTo ProxArquivo
        gfImpDocumento = True
        Exit Function
    End If
        
    If pFormulario = "COT_LMI" Or pFormulario = "REL_LMI" Or pFormulario = "MULTLMIE" Or pFormulario = "MULTLMIR" Then
        'Antes de imprimir verifico se tem a variavel de quantidade de páginas para trocar
        Dim lNomArqAux  As String
        lNumArquivo = FreeFile
        Open lArqDados For Input Access Read As #lNumArquivo
        lNumArqIST = FreeFile
        lNomArqAux = Replace(lArqDados, ".DDD", ".DD1")
        If Dir(lNomArqAux) <> "" Then
            On Error Resume Next
            Kill lNomArqAux
            On Error GoTo 0
        End If
        If Dir(lNomArqAux) <> "" Then
            lNomArqAux = Replace(lArqDados, ".DDD", ".DD2")
        End If
        If Dir(lNomArqAux) <> "" Then
            On Error Resume Next
            Kill lNomArqAux
            On Error GoTo 0
        End If
        If Dir(lNomArqAux) <> "" Then
            lNomArqAux = Replace(lArqDados, ".DDD", ".DD3")
        End If
            
        Open lNomArqAux For Append Shared As #lNumArqIST
        On Error GoTo 0
        Do While Not EOF(lNumArquivo)
            Line Input #lNumArquivo, lRegistro
            lRegistro = Replace(lRegistro, "@@TOTALPAG@@", lTotPagina)
            Print #lNumArqIST, lRegistro
        Loop
        
        Close #lNumArquivo
        Close #lNumArqIST
        lArqDados = lNomArqAux
    End If

    '8. Processo para imprimir.

    On Error Resume Next
    Call clearDados
    If Err <> 0 Then
        Call gM46V999.gpGraLog(1, "FTE0085 - " & gM46V999.gPrefixo & _
                         "Erro ao imprimir documento (comando CLEARDADOS - " & Err & " - " & Error & ").")
        Close #lNumArquivo
        Kill lArqDados
        On Error GoTo 0
        Exit Function
    End If
    Call setImage(lEndImagem, 0, 0, 0, 0)                       'Arquivo imagem.
    If Err <> 0 Then
        Call gM46V999.gpGraLog(1, "FTE0086 - " & gM46V999.gPrefixo & _
                         "Erro ao imprimir documento (comando SETIMAGE - " & Err & " - " & Error & ").")
        Close #lNumArquivo
        Kill lArqDados
        On Error GoTo 0
        Exit Function
    End If
    If pOpcao = 2 Then      'Criar arquivo imagem.
        Call setArqImage(pImagem)                               'Arquivo imagem a ser gerado.
        If Err <> 0 Then
            Call gM46V999.gpGraLog(1, "FTE0087 - " & gM46V999.gPrefixo & _
                             "Erro ao imprimir documento (comando SETARQIMAGE(""C:\SYAS.XXX"") - " & _
                             Err & " - " & Error & ").")
            Close #lNumArquivo
            Kill lArqDados
            On Error GoTo 0
            Exit Function
        End If
    Else
        Call setArqImage("")                                    'Arquivo imagem a ser gerado.
        If Err <> 0 Then
            Call gM46V999.gpGraLog(1, "FTE0088 - " & gM46V999.gPrefixo & _
                             "Erro ao imprimir documento (comando SETARQIMAGE("""") - " & Err & " - " & _
                             Error & ").")
            Close #lNumArquivo
            Kill lArqDados
            On Error GoTo 0
            Exit Function
        End If
    End If
    
    Call setArqConfig(lEndConfig, "1")                    'Arquivo de configuração.
    If Err <> 0 Then
        Call gM46V999.gpGraLog(1, "FTE0089 - " & gM46V999.gPrefixo & _
                         "Erro ao imprimir documento (comando SETARQCONFIG - " & Err & " - " & Error & _
                         ").")
        Close #lNumArquivo
        Kill lArqDados
        On Error GoTo 0
        Exit Function
    End If
    Call setArqDados(lArqDados)                                 'Arquivo de dados.
    If Err <> 0 Then
        Call gM46V999.gpGraLog(1, "FTE0090 - " & gM46V999.gPrefixo & _
                         "Erro ao imprimir documento (comando SETARQDADOS - " & Err & " - " & Error & _
                         ").")
        Close #lNumArquivo
        Kill lArqDados
        On Error GoTo 0
        Exit Function
    End If
    DoEvents
    Err = 0
    Call PrintReld(CStr(pOpcao))        '<<<<<<<<<<========== Comando que apresentou erro - Dr. Watson...
    If Err <> 0 Then
        'para o form "COT_LMI"esta dando esse erro 6
        If Err.Number = 6 And lVerifOverf = False Then
            lVerifOverf = True
            Err = 0
            Close #lNumArquivo
            Kill lArqDados
            GoTo Retorno
        End If
        
        Call gM46V999.gpGraLog(1, "FTE0091 - " & gM46V999.gPrefixo & _
                         "Erro ao imprimir documento (comando PRINTRELD - " & Err & " - " & Error & ").")
        Close #lNumArquivo
        Kill lArqDados
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    '9. Exclui arquivo de dados.
    On Error Resume Next
    Kill lArqDados
    Select Case Err
        Case 0
            gfImpDocumento = True
        Case 70     'Erro 70 - Permission denied.
            'Este erro acontece quando o usuário não espera concluir a emissão do documento.
            Call gM46V999.gpGraLog(1, "FTE0083 - " & gM46V999.gPrefixo & _
                             "Erro ao imprimir documento (usuário não esperou concluir emissão do " & _
                             "documento - sair do sistema e reiniciar processo).")
        Case Else
            Call gM46V999.gpGraLog(1, "FTE0092 - " & gM46V999.gPrefixo & _
                             "Erro ao imprimir documento (exclusão do arquivo de dados - " & Err & _
                             " - " & Error & ").")
    End Select
    Close #lNumArquivo
    On Error GoTo 0

ProxArquivo:
    For lI = lCount To UBound(lArrayArq) - 1
        
        lPos = InStr(1, lArrayArq(lI), "|")
        lArqDados = Left(lArrayArq(lI), lPos - 1)
        lEndConfig = Mid(lArrayArq(lI), lPos + 1, InStr(lPos + 1, lArrayArq(lI), "|") - 1 - lPos)
        lPos = InStr(lPos + 1, lArrayArq(lI), "|")
        lEndImagem = Mid(lArrayArq(lI), lPos + 1)
        pImagem = Replace(pImagem, ".WMF", lI & ".WMF")
        lCount = lCount + 1
        If gM46V999.gfPreenchido(lEndConfig) And Dir(lEndConfig) <> "" Then
            GoTo ProximoForm
        Else
            Exit For
        End If
    Next lI
    gfImpDocumento = True
End Function

Public Function gfIsMetaframe() As Boolean
   Dim hostname As String * 256
   Dim hostent_addr As Long
   Dim host As HOSTENT
   Dim hostip_addr As Long
   Dim temp_ip_address() As Byte
   Dim i As Integer
   Dim ip_address As String
   Dim ret As Boolean
      
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

Sub SocketsCleanup()
   Dim lReturn As Long

       lReturn = WSACleanup()

       If lReturn <> 0 Then
           MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
           End
       End If
End Sub

    
Function lobyte(ByVal wParam As Integer)

       lobyte = wParam And &HFF&

   End Function


   Function hibyte(ByVal wParam As Integer)

       hibyte = wParam \ &H100 And &HFF&

   End Function

