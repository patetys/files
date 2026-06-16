Attribute VB_Name = "basM46V401"
Option Explicit
Public gDiretorio           As String
Public objPedido            As clsA46V702A  'Objeto da classe clsA46V702A  (TAB_PED)
Public objstcPedido         As stcA46V702B  'Objeto da classe de estrutura stcA46V702B (TAB_PED)
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" _
                         (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Sub gpPausa(ByVal lSegundos)

Dim i
    Dim lMiliSegundos   As Double
    
        lMiliSegundos = lSegundos * 1000

    i = DoEvents
    Sleep (lMiliSegundos)
    i = DoEvents
    
    Exit Sub
    
End Sub

Sub Main()
    Dim lbdP00MPED      As ADODB.Connection 'Base de dados: P0042300.
    Dim lCodCorretor    As String           'Código do corretor.
    Dim lCodUsuario     As String  'Long             'Código do usuário.
    Dim lDatModulo      As Double           'Data do módulo.
    Dim lDatVersao      As String           'Data da versão.
    Dim lEndP2842400    As String           'Endereço do arquivo P2842400.EXE.
    Dim lLargura        As Integer          'Largura.
    Dim lMensagem       As String           'Mensagem.
    Dim lNumP2842400    As Integer          'Utilizado com FreeFile para o arquivo P2842400.EXE.
    Dim lPosicao        As Integer          'Utilizado com InStr.
    Dim lRegistro       As String           'Registro.
    

    Screen.MousePointer = vbHourglass

    '1. Verifica se form já está ativo.
    If App.PrevInstance = True Then
        End
    End If

    '2. Obtém informações iniciais do sistema: tipo do usuário, endereços das bases de dados, endereços
    '   dos arquivos de impressão, indicador de uso do auxílio automático, nome ou código do usuário.
    Call gM46V999.gpInicia("")

    '3. Emite mensagem de entrada (precisa ser depois de obter dados do arquivo de inicialização).
    Call gM46V999.gpGraLog(0, "AGE0001 - Início do processamento do programa " & App.EXEName & ".")

    '4. procesa os arquivos
    Call mpCarregaArquivosRecebidos

    
    'Finaliza o Programa
    End

End Sub
Private Sub mpCarregaArquivosRecebidos()
Dim clsA46V601          As clsA46V601A      'Classe: cria arquivo de transmissão.
Dim lEndArqTransmissao  As String
Dim lMensagem           As String
Dim lPathArqs           As String
Dim lArquivo            As String
Dim lCkPoint            As String

    
Err = 0
On Error GoTo TRATAMENTO_ERRO

    lArquivo = Dir(gM46V999.gAppPath & "*.034")
    Do While gM46V999.gfPreenchido(lArquivo) = True
        lEndArqTransmissao = gM46V999.gAppPath & lArquivo
        Set clsA46V601 = New clsA46V601A
        Set clsA46V601.M46V999 = gM46V999
        If Not clsA46V601.gfRecArqCotacaoAceitaRec(lEndArqTransmissao, lMensagem) Then
            Call gM46V999.gpGraLog(0, lMensagem)
        End If
        Kill gM46V999.gAppPath & lArquivo
        lArquivo = Dir(gM46V999.gAppPath & "*.034")
    Loop
    Set clsA46V601 = Nothing

SAIDA:
    Err = 0
    On Error GoTo 0
    Exit Sub
TRATAMENTO_ERRO:
    If gM46V999.gfPreenchido(lMensagem) = False Then
        lMensagem = Err & " - " & Error
    End If
    lMensagem = "mpCarregaArquivosRecebidos Erro: " & lMensagem & vbCrLf
    lMensagem = lMensagem & "CheckPoint: " & lCkPoint
    Call gM46V999.gpGraLog(0, lMensagem)
    GoTo SAIDA
    
End Sub

Private Function mfMoveFile(ByVal pOrigem As String, ByVal pDestino As String) As Boolean
Dim lMsgErr         As String
Err = 0
On Error GoTo TratamentoErro
    mfMoveFile = False
    
'1- Efetua a cópia do arquivo
    MoveFile pOrigem, pDestino

'2- Verifica a existência do arquivo no destino copiado.
    If gM46V999.gfPreenchido(Dir(pDestino)) = False Then
        lMsgErr = "O arquivo não foi movido para o caminho: " & pDestino
        GoTo TratamentoErro
    End If
    
    mfMoveFile = True
SAIDA:
    Err = 0
    On Error GoTo 0
    Exit Function
    
TratamentoErro:
    If gM46V999.gfPreenchido(lMsgErr) = False Then
        lMsgErr = Err & " - " & Error
    End If
    lMsgErr = "mfMoveFile Erro: " & lMsgErr
    Call gM46V999.gpGraLog(0, lMsgErr)
    GoTo SAIDA

End Function

Private Sub mpGravaErroTransmissao(ByVal pMensagem As String, ByVal pNomArqLog As String)
    'Procedure: grava arquivo log.

    '                        pMensagem...Mensagem.
Dim lEndLog             As String
Dim lPosicao            As Integer  'Utilizado com InStr.
Dim lNumLog             As Integer
Dim lMensagem           As String   'Mensagem.
Dim lMsgErr             As String
Dim lconta              As Integer
Static lNomComputador   As String   'Nome do computador.
Static lNomPrograma     As String   'Nome do programa.
Static lNomUsuario      As String   'Nome do usuário.
Static lPriVez          As Boolean  'Indicador de primeira vez (false).


Err = 0
On Error GoTo TRATAMENTO_ERRO

    lEndLog = gM46V999.gAppPath & pNomArqLog
    
    '2. Ao executar esta procedure pela primeira vez, obter dados utilizados na gravação das mensagens.
    If lPriVez = False Then
        lPriVez = True

        'b) Obtém nome do usuário.
        If gM46V999.gfPreenchido(gM46V999.gNomUsuario) Then
            lNomUsuario = gM46V999.gNomUsuario
            While Len(lNomUsuario) < 20
                lNomUsuario = lNomUsuario & " "
            Wend
        Else
            lNomUsuario = String(20, "X")
        End If
        'c) Obtém nome do computador.
        lNomComputador = gM99V001.GetLoginComputerName()
        If Not gM46V999.gfPreenchido(lNomComputador) Then
            lNomComputador = "********"
        End If
        While Len(lNomComputador) < 16
            lNomComputador = lNomComputador & " "
        Wend
        'd) Obtém nome do programa.
        lNomPrograma = Trim$(App.EXEName)
        While Len(lNomPrograma) < 12
            lNomPrograma = lNomPrograma & " "
        Wend
    End If

    '3. Abre arquivo log.
    lNumLog = FreeFile
    Open lEndLog For Output Access Write As #lNumLog
        If Err <> 0 Then
            GoTo TRATAMENTO_ERRO
        End If
        
        '6. Grava mensagem.
        lMensagem = pMensagem
        Do
            lPosicao = InStr(lMensagem, vbLf)
            If lPosicao = 0 Then
                Exit Do
            End If
            lMensagem = Left$(lMensagem, lPosicao - 1) & " " & Mid$(lMensagem, lPosicao + 1)
        Loop
        Print #lNumLog, gM46V999.gForHoje & " | " & Format$(Now, "hh:nn:ss") & " | " & lNomUsuario & " | " & _
                        lNomComputador & " | " & lNomPrograma & " | " & lMensagem
                    
    Close #lNumLog

SAIDA:
    Err = 0
    On Error GoTo 0
    Exit Sub

TRATAMENTO_ERRO:
    If gM46V999.gfPreenchido(lMsgErr) = False Then
        lMsgErr = Err & " - " & Error
    End If
    lMsgErr = "mpGravaErroTransmissao Erro: " & lMsgErr
    Call gM46V999.gpGraLog(0, lMsgErr)
    GoTo SAIDA
    
End Sub




