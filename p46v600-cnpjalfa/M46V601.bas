Attribute VB_Name = "basM46V601"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public gProJurIOF       As Double           'Produto: juros de IOF.
Public gQuaErros        As Byte             'Quantidade de erros.
Public objPedido            As clsA46V702A  'Objeto da classe clsA46V702A  (TAB_PED)
Public objstcPedido         As stcA46V702B  'Objeto da classe de estrutura stcA46V702B (TAB_PED)
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" _
                         (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long


Public gLiberaDescontoEspecial  As Integer
Public gCod_plano   As Long
Public gCod_Regiao  As Long
Public EhTeste      As Boolean
Private Enum e_TipoSistema
    e_TipoSistema_ND = 0
    e_TipoSistema_Serasa = 1
    e_TipoSistema_Nortix = 2
End Enum

Private Enum e_TipoChave
    e_TipoChave_ND = 0
    e_TipoChave_Sistema = 1
    e_TipoChave_Cnpj = 2
    e_TipoChave_Placa = 3
    e_TipoChave_Chassi = 4
    e_TipoChave_Renavam = 5
End Enum

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
    
    Randomize
    gM46V999.sUniqueInstanceCode = CLng(Rnd * 9999) + 1
     
    If gM46V999.gfPreenchido(gM46V999.gDiretorio) = False Then
        gM46V999.gDiretorio = "X:\mailboxes\"
    End If

    If UCase(gM46V999.gNomUsuario) = "JMIWAI" Then
        'gM46V999.gDiretorio = gM46V999.gAppPath & "mailboxes\" 'qa2
        gM46V999.gDiretorio = "C:\Soft\ediserver\mailboxes\"
        gM46V999.gAppPath = "C:\Soft\EDIWare\P46V600\"
        EhTeste = True
    End If
    
    If UCase(gM46V999.gNomUsuario) = "MOSILVA" _
    Or UCase(gM46V999.gNomUsuario) = "TOKAMOTO" _
    Or UCase(gM46V999.gNomUsuario) = "ANDREL" _
    Or UCase(gM46V999.gNomUsuario) = "CYBER18" _
    Or UCase(gM46V999.gNomUsuario) = "DARIBEIRO" _
    Or UCase(gM46V999.gNomUsuario) = "SONDA01" Then
        EhTeste = True
        gM46V999.gDiretorio = "C:\Mailboxes\"
        gM46V999.gAppPath = "C:\Mailboxes\"
    End If
    'procesa os arquivos
    'Call gM46V999.gpGraLog(0, "AGE0001 - Processa arquivos de transmissão inspeção.")
    Call mpCarregaArquivosInspecao
    
    'Call gM46V999.gpGraLog(0, "AGE0001 - Processa arquivos de transmissao.")
    Call mpCarregaArquivosRecebidos
    
    'Call gM46V999.gpGraLog(0, "AGE0001 - Processa arquivos de cotação.")
    Call mpCarregaArquivosCotacao
    
    'Call gM46V999.gpGraLog(0, "AGE0001 - Gera arquivos de retorno de cotação.")
    Call mpEnviaRetornoCotacao
    
    'Call gM46V999.gpGraLog(0, "AGE0001 - Verificação propostas SIES.")
    Call mpVerificaSies

    'Finaliza o Programa
    End

End Sub
Private Sub mpCarregaArquivosCotacao()
Dim clsA46V601          As clsA46V601A      'Classe: cria arquivo de transmissão.
Dim lEndArqTransmissao  As String
Dim lMensagem           As String
Dim objstcPedido        As stcA46V702B
Dim lPathArqs           As String
Dim lArquivo            As String
Dim lCkPoint            As String
Dim lPreArq             As String

    
Err = 0
On Error GoTo TRATAMENTO_ERRO

    lPathArqs = gM46V999.gDiretorio & "Yasuda\"
    
       
    lCkPoint = "0"
    lArquivo = Dir(lPathArqs & "*.032")
    lCkPoint = "1"
    Do While gM46V999.gfPreenchido(lArquivo) = True
        lCkPoint = "2"
        If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & lArquivo)) = True Then
            lCkPoint = "3"
            Kill gM46V999.gAppPath & lArquivo
        End If
        lCkPoint = "4"
        If mfMoveFile(lPathArqs & lArquivo, gM46V999.gAppPath & lArquivo) = True Then
            lCkPoint = "5"
            If Len(lArquivo) = 18 Then
                lPreArq = Mid(lArquivo, 1, 10)
            Else
                '*** Ajustes para deixar o nome do arquivo com tamanho dinâmico - SIES - taflorencio - 27/11/2013
                lPreArq = Mid(lArquivo, 1, Len(lArquivo) - 4)
            End If
            'Copia o protocolo ref. ao mesmo arquivo. (apenas para armazenamento)
            If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & lPreArq & ".031")) = True Then
                lCkPoint = "6"
                Kill gM46V999.gAppPath & lPreArq & ".031"
            End If
            lCkPoint = "7"
            'Call mfMoveFile(lPathArqs & Mid(lArquivo, 1, 8) & ".031", gM46V999.gAppPath & Mid(lArquivo, 1, 8) & ".031")
            lCkPoint = "8"
            'Posiciona o nome do arquivo para processamento
            lEndArqTransmissao = gM46V999.gAppPath & lArquivo
            'Call gM46V999.gpGraLog(0, "AGE0001 - Lendo arquiv " & lEndArqTransmissao)
            lCkPoint = "9"
            Set clsA46V601 = New clsA46V601A
            Set clsA46V601.M46V999 = gM46V999
            
            lCkPoint = "10"
            If Not clsA46V601.gfRecArqCotacao(lEndArqTransmissao, lMensagem, gM46V999.gAppPath & lPreArq & ".031", objstcPedido) Then
                lCkPoint = "11"
                Call gM46V999.gpGraLog(0, lMensagem)
                                
                'marca como ***Processado nÃO OK!
                lCkPoint = "12"
                If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & lPreArq & ".N32")) = True Then
                    lCkPoint = "13"
                    Kill gM46V999.gAppPath & lPreArq & ".N32"
                End If
                lCkPoint = "14"
                Name gM46V999.gAppPath & lArquivo As gM46V999.gAppPath & lPreArq & ".N32"
                'Grava o arquivo um arquivo de log contendo a mensagem de erro
                lCkPoint = "16"
                If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & lPreArq & ".ERR")) = True Then
                    lCkPoint = "17"
                    Kill gM46V999.gAppPath & lPreArq & ".ERR"
                End If
                On Error Resume Next
                Shell "NET SEND TJGOMES P46V600 - Não conseguiu ler arquivo de cotação"
                Shell "NET SEND TJGOMES P46V600 - Não conseguiu ler arquivo de cotação"
                FileCopy gM46V999.gAppPath & lPreArq & ".N32", "x:\" & lPreArq & ".N32"
                Err.Clear
                On Error GoTo TRATAMENTO_ERRO
                lCkPoint = "18"
                'Call mpGravaErroTransmissao(lMensagem, Mid(lArquivo, 1, 8) & ".ERR")
            Else
                'marca como ***Processado OK!
                lCkPoint = "19"
                If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & lPreArq & ".Y32")) = True Then
                    lCkPoint = "20"
                    Kill gM46V999.gAppPath & lPreArq & ".Y32"
                End If
                lCkPoint = "21"
                Name gM46V999.gAppPath & lArquivo As gM46V999.gAppPath & lPreArq & ".Y32"
            End If
        End If
        
        'busca o próximo
        lCkPoint = "22"
        lArquivo = Dir(lPathArqs & "*.032")
    Loop
    
    Set clsA46V601 = Nothing

SAIDA:
    Err = 0
    On Error GoTo 0
    Exit Sub
TRATAMENTO_ERRO:
    If gM46V999.gfPreenchido(lMensagem) = False Then
        lMensagem = Err & " - " & ERROR
    End If
    lMensagem = "mpCarregaArquivosCotacao Erro: " & lMensagem & vbCrLf
    lMensagem = lMensagem & "CheckPoint: " & lCkPoint
    Call gM46V999.gpGraLog(0, lMensagem)
    GoTo SAIDA
    
End Sub
Private Sub mpCarregaArquivosRecebidos()
Dim clsA46V601          As clsA46V601A      'Classe: cria arquivo de transmissão.
Dim lEndArqTransmissao  As String
Dim lMensagem           As String
Dim objstcPedido        As stcA46V702B
Dim lPathArqs           As String
Dim lArquivo            As String
Dim lCkPoint            As String
Dim lExtensao           As String
Dim lExtOK              As String
Dim lExtNOK             As String
Dim lConexaoEAPDB       As ADODB.Connection

'*** CRIAÇÃO DE ROTINA PARA PEGAR O TIPO DE COTAÇÃO - PROJETO SIES - 17/09/2013 - TAFLORENCIO ***
    
Err = 0
On Error GoTo TRATAMENTO_ERRO

    lPathArqs = gM46V999.gDiretorio & "multi\"
    
    lCkPoint = "0"
    lExtensao = "008": lExtOK = ".Y08": lExtNOK = ".N08"       '-- Transmissão Do Corretor
    lArquivo = Dir(lPathArqs & "*." & lExtensao)
    If gM46V999.gfPreenchido(lArquivo) = False And lExtensao = "008" Then
        lExtensao = "091": lExtOK = ".Y91": lExtNOK = ".N91"   '-- Transmissão Automatica
        lArquivo = Dir(lPathArqs & "*." & lExtensao)
    End If
    If gM46V999.gfPreenchido(lArquivo) = False And lExtensao = "091" Then
        lExtensao = "046": lExtOK = ".Y46": lExtNOK = ".N46"   '-- Log de Calculo
        lPathArqs = gM46V999.gDiretorio & "yasuda\"
        lArquivo = Dir(lPathArqs & "*." & lExtensao)
    End If
    If gM46V999.gfPreenchido(lArquivo) = False And lExtensao = "046" Then
        lExtensao = "015": lExtOK = ".Y15": lExtNOK = ".N15"   '-- Log de Calculo
        lPathArqs = gM46V999.gDiretorio & "yasuda\"
        lArquivo = Dir(lPathArqs & "*." & lExtensao)
    End If

    lCkPoint = "1"
    Do While gM46V999.gfPreenchido(lArquivo) = True
        lCkPoint = "2"
        If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & lArquivo)) = True Then
            lCkPoint = "3"
            Kill gM46V999.gAppPath & lArquivo
        End If
        lCkPoint = "4"
        If mfMoveFile(lPathArqs & lArquivo, gM46V999.gAppPath & lArquivo) = True Then
            lCkPoint = "5"
            If lExtensao = "008" Then
            'Copia o protocolo ref. ao mesmo arquivo. (apenas para armazenamento)
                If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & ".009")) = True Then
                    lCkPoint = "6"
                    Kill gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & ".009"
                End If
                lCkPoint = "7"
                Call mfMoveFile(lPathArqs & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & ".009", gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & ".009")
            End If
            If lExtensao = "015" Then
            'Copia o protocolo ref. ao mesmo arquivo. (apenas para armazenamento)
                If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & ".Y15")) = True Then
                    lCkPoint = "6"
                    Kill gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & ".Y15"
                End If
                lCkPoint = "7"
                Call mfMoveFile(lPathArqs & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & ".Y15", gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & ".Y15")
            End If
            lCkPoint = "8"
            'Posiciona o nome do arquivo para processamento
            lEndArqTransmissao = gM46V999.gAppPath & lArquivo
            'Call gM46V999.gpGraLog(0, "AGE0001 - Lendo arquivo " & lEndArqTransmissao)
            lCkPoint = "9"
            Set clsA46V601 = New clsA46V601A
            Set clsA46V601.M46V999 = gM46V999
            
            lCkPoint = "10"
            Set objstcPedido = New stcA46V702B
            Set objstcPedido.M46V999 = gM46V999
            
            If lExtensao = "008" Or lExtensao = "015" Then
                objstcPedido.Tip_Transmissao = eTipTransmissao_Corretor
            ElseIf lExtensao = "091" Then
                objstcPedido.Tip_Transmissao = eTipTransmissao_Automatic
            End If

'            '20180702: Removido
'            If UCase(gM46V999.gSerDIVERSOS) Like "*MIX*" Or EhTeste = True Then
                If Not gM46V999.gfAbrBasSQL("EAPDB", gM46V999.gTipEAPDB, gM46V999.gSerEAPDB, gM46V999.gEndEAPDB, gM46V999.gUidEAPDB, gM46V999.gPwdEAPDB, lConexaoEAPDB, lMensagem) Then
                    GoTo TRATAMENTO_ERRO
                End If
'            Else
'                If Not gM46V999.gfAbrBasSQL("EAPDB", "2", "SP310110PSLUWEB", "EAPDB", "yasuser", "user00.msdeweb", lConexaoEAPDB, lMensagem) Then
'                    GoTo TRATAMENTO_ERRO
'                End If
'            End If
            Set clsA46V601.mConexaoEAPDB = lConexaoEAPDB
            If Not clsA46V601.gfRecArqTransmissao(lEndArqTransmissao, lMensagem, objstcPedido) Then
                lCkPoint = "11"
                Call gM46V999.gpGraLog(0, lMensagem)
                'marca como ***Processado nÃO OK!
                lCkPoint = "12"
                If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & lExtNOK)) = True Then
                    lCkPoint = "13"
                    Kill gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & lExtNOK
                End If
                lCkPoint = "14"
                Name gM46V999.gAppPath & lArquivo As gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & lExtNOK
                'Grava um arquivo de log contendo a mensagem de erro
                lCkPoint = "16"
                If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & ".ERR")) = True Then
                    lCkPoint = "17"
                    Kill gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & ".ERR"
                End If
                lCkPoint = "18"
                Call mpGravaErroTransmissao(lMensagem, Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & ".ERR")
            Else
                'marca como ***Processado OK!
                lCkPoint = "19"
                If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & lExtOK)) = True Then
                    lCkPoint = "20"
                    Kill gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & lExtOK
                End If
                lCkPoint = "21"
                Name gM46V999.gAppPath & lArquivo As gM46V999.gAppPath & Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & lExtOK
              
                '*** Inclusão da rotina para geração de protocolo de recebimento de arquivo SIES - 16/09/2013 Projeto RD - SIES - taflorencio ***
                If objstcPedido.objstcCtrlEmis.Tip_Cotac = "4" Or objstcPedido.objstcCtrlEmis.Tip_Cotac = "5" Then
                    lArquivo = Mid(lArquivo, 1, (InStr(1, lArquivo, ".") - 1)) & ".009"
                    If gM46V999.mfEnviaProtocolo(Mid(lArquivo, 1, 6), lArquivo, lMensagem) = False Then
                        Call gM46V999.gpGraLog(0, "P46V600 - ERRO ao enviar protocolo" & lMensagem)
                        Shell "net send simony " & App.EXEName & " - ERRO ao enviar protocolo" & lMensagem
                        Shell "net send taflorencio " & App.EXEName & " - ERRO ao enviar protocolo" & lMensagem
                    End If
                End If
                lCkPoint = "22"
            End If
        End If
        
        'busca o próximo
        lCkPoint = "23"
        lArquivo = Dir(lPathArqs & "*." & lExtensao)
        If gM46V999.gfPreenchido(lArquivo) = False And lExtensao = "008" Then
            lExtensao = "091": lExtOK = ".Y91": lExtNOK = ".N91"   '-- Transmissão Automatica
            lArquivo = Dir(lPathArqs & "*." & lExtensao)
        End If
        If gM46V999.gfPreenchido(lArquivo) = False And lExtensao = "091" Then
            lExtensao = "046": lExtOK = ".Y46": lExtNOK = ".N46"   '-- Log de Calculo
            lPathArqs = gM46V999.gDiretorio & "yasuda\"
            lArquivo = Dir(lPathArqs & "*." & lExtensao)
        End If
        If gM46V999.gfPreenchido(lArquivo) = False And lExtensao = "046" Then
            lExtensao = "015": lExtOK = ".Y15": lExtNOK = ".N15"   '-- Log de Calculo
            lPathArqs = gM46V999.gDiretorio & "yasuda\"
            lArquivo = Dir(lPathArqs & "*." & lExtensao)
        End If
    Loop
    
    Set clsA46V601 = Nothing

SAIDA:
    Err = 0
    On Error GoTo 0
    Exit Sub
TRATAMENTO_ERRO:
    If gM46V999.gfPreenchido(lMensagem) = False Then
        lMensagem = Err & " - " & ERROR
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
        lMsgErr = Err & " - " & ERROR
    End If
    lMsgErr = "mfMoveFile Erro: " & lMsgErr
    'Call gM46V999.gpGraLog(0, lMsgErr)
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
Dim lConta              As Integer
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
        lMsgErr = Err & " - " & ERROR
    End If
    lMsgErr = "mpGravaErroTransmissao Erro: " & lMsgErr
    Call gM46V999.gpGraLog(0, lMsgErr)
    
    For lConta = 1 To 2
        Shell "NET SEND TJGOMES P46V600 - " & lMsgErr & lMensagem
        Shell "NET SEND VAGNERSK P46V600 - " & lMsgErr & lMensagem
    Next lConta
    
    GoTo SAIDA
    
End Sub

Private Sub mpEnviaRetornoCotacao()
Dim lMensagem       As String
Dim lComando        As String
Dim lrst            As clsYasRecordSet
Dim lNossoNumero    As String
Dim lColecao        As Collection
Dim lCodCorretor    As String
Dim lCodUsuario     As String
Dim lCotacao        As String
Dim lStatus         As Integer
Dim lSenha          As String
Dim lBusca          As String
Dim lMaxDesc        As Double
Dim clsA46V101A     As clsA46V101A
Dim lDir            As String
Dim bdSYASEMIS      As ADODB.Connection
Dim lObjPed         As clsA46V702A
Dim lObjstcPed      As stcA46V702B
Dim lEndArqTransmissao  As String

    If gM46V999.gfAbrBasSQL(gM46V999.gEndGed, gM46V999.gTipGed, gM46V999.gSerGed, gM46V999.gEndGed, gM46V999.gUidGed, _
       gM46V999.gPwdGed, bdSYASEMIS, lMensagem) = False Then
       If Not UCase(gM46V999.gSerSYASEMIS) Like "*MIX*" Then
            Shell "net send tjgomes Servidor PSLUEDI com falha na conexão SQL.  Verificar com urgência."
            Shell "net send hideo Servidor PSLUEDI com falha na conexão SQL.  Verificar com urgência."
            Shell "net send newton Servidor PSLUEDI com falha na conexão SQL.  Verificar com urgência."
            Shell "net send oscar Servidor PSLUEDI com falha na conexão SQL.  Verificar com urgência."
            Shell "net send jayme Servidor PSLUEDI com falha na conexão SQL.  Verificar com urgência."
        End If
        Call gM46V999.gpGraLog(2, lMensagem & vbLf & vbLf & "O sistema será encerrado.")
        End
    End If
    
    lComando = " Select s.*, c.vdesconto, cod_ramo from siscota.dbo.tab_syas_cotacao as s"
    lComando = lComando & " inner join siscota.siscota.tab_cotacao as c"
    lComando = lComando & " on s.ccotacao = c.ccotacao"
    lComando = lComando & " where cod_ramo not in (310,112,113,114,0)"
    lComando = lComando & " and STATUS_ENVIO = 1 "
    
    If Not gM46V999.gfObtRegistro(bdSYASEMIS, lComando, lrst, lMensagem) Then
        Call gM46V999.gpGraLog(0, lMensagem)
        End
    End If
    
    Do While Not lrst.EOF
        If lrst("COD_RAMO") = 710 Or lrst("COD_RAMO") = 510 Or lrst("COD_RAMO") = 670 Or lrst("COD_RAMO") = 300 Or lrst("COD_RAMO") = 620 Then
            lEndArqTransmissao = gM46V999.gEndRemessa & lrst("NomArq") & ".036"
        Else
            lEndArqTransmissao = gM46V999.gEndRemessa & lrst("NomArq") & ".034"
        End If
        
        If gM46V999.gfPreenchido(Dir(lEndArqTransmissao)) Then
            On Error Resume Next
            Kill lEndArqTransmissao
        End If
    lrst.MoveNext
    Loop
    
    lrst.MoveFirst
    
    Do While Not lrst.EOF
        
        '*** Inclusão da rotina abaixo para deração de mais de uma cotação por arquivo protocolo - SIES - taflorencio - 11/10/2013
        If lrst("COD_RAMO") = 710 Or lrst("COD_RAMO") = 510 Or lrst("COD_RAMO") = 670 Or lrst("COD_RAMO") = 300 Or lrst("COD_RAMO") = 620 Then
            lEndArqTransmissao = gM46V999.gEndRemessa & lrst("NomArq") & ".036"
        Else
            lEndArqTransmissao = gM46V999.gEndRemessa & lrst("NomArq") & ".034"
        End If
        
        'Call gM46V999.gpGraLog(0, "AGE0001 - Verificando o registro " & lrst("Nosso_Numero"))
        lNossoNumero = lrst("Nosso_Numero")
        lCodCorretor = lrst("COD_CORRETOR")
        lCodUsuario = lrst("COD_Usuario")
        lStatus = lrst("Cod_Status")
        lCotacao = lrst("CCotacao")
        lSenha = lrst("Senha")
        lMaxDesc = lrst("VDesconto")
        
        Set lObjPed = New clsA46V702A
        Set lObjPed.M46V999 = gM46V999
        Set lObjstcPed = New stcA46V702B
        Set lObjstcPed.M46V999 = gM46V999
        
        lObjstcPed.NOSSO_NUMERO = lNossoNumero
        If Not lObjPed.mfCarregarColecao(lObjstcPed, "", "ramosdiversos.dbo.", True) Then
            Call gM46V999.gpGraLog(0, lObjPed.UltimoErro)
            End
        End If
        If lObjPed.Colecao.Count > 0 Then
            Set lColecao = lObjPed.Colecao
            'Call gM46V999.gpGraLog(0, "AGE0001 - Gerando arquivo " & lrst("NomArq") & ".034")
            Set clsA46V101A = New clsA46V101A
            Set clsA46V101A.M46V999 = gM46V999
            If lrst("COD_Usuario") = "0" Then
                If clsA46V101A.gfAtualizaClasse4Portal(lNossoNumero, lStatus, lMaxDesc, lMensagem, bdSYASEMIS) = False Then
                    Exit Sub
                End If
            Else
                If clsA46V101A.gfCriCotacaoAceita(lCodCorretor, lCodUsuario, lColecao, _
                         lStatus, lCotacao, lMaxDesc, lEndArqTransmissao, lSenha, lMensagem) = False Then
                    Exit Sub
                End If
            End If
            'Call gM46V999.gpGraLog(0, "AGE0001 - Arquivo " & lrst("NomArq") & ".034 gerado.")
            Set clsA46V101A = Nothing
            
            lComando = "Update siscota.dbo.tab_syas_cotacao set status_envio = 2 where "
            lComando = lComando & " Nosso_Numero = " & gM46V999.gfFormatarCamposSQL(lrst("Nosso_Numero"))
            If gM46V999.gfExeSQL(bdSYASEMIS, lComando, lMensagem) <> 0 Then
                Call gM46V999.gpGraLog(0, lMensagem)
                End
            End If
            'Call gM46V999.gpGraLog(0, "AGE0001 - Proximo")
        Else
            'Call gM46V999.gpGraLog(0, "AGE0001 - Não localizado registros : " & lNossoNumero)
        End If
        lrst.MoveNext
    Loop
    'Call gM46V999.gpGraLog(0, "AGE0001 - acabou")
    Exit Sub
    
End Sub


Private Sub mpCarregaArquivosInspecao()
Dim clsA46V601          As clsA46V601A      'Classe: cria arquivo de transmissão.
Dim lEndArqTransmissao  As String
Dim lMensagem           As String
Dim objstcPedido        As stcA46V702B
Dim lPathArqs           As String
Dim lArquivo            As String
Dim lCkPoint            As String

    
Err = 0
On Error GoTo TRATAMENTO_ERRO

    lPathArqs = gM46V999.gDiretorio & "multi\"
    
    
    lCkPoint = "0"
    
    lArquivo = Dir(lPathArqs & "*.055")
    
    lCkPoint = "1"
    Do While gM46V999.gfPreenchido(lArquivo) = True
        lCkPoint = "2"
        If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & lArquivo)) = True Then
            lCkPoint = "3"
            Kill gM46V999.gAppPath & lArquivo
        End If
        lCkPoint = "4"
        If mfMoveFile(lPathArqs & lArquivo, gM46V999.gAppPath & lArquivo) = True Then
            lCkPoint = "5"
            'Copia o protocolo ref. ao mesmo arquivo. (apenas para armazenamento)
             lCkPoint = "8"
            'Posiciona o nome do arquivo para processamento
            lEndArqTransmissao = gM46V999.gAppPath & lArquivo
            'Call gM46V999.gpGraLog(0, "AGE0001 - Lendo arquivo " & lEndArqTransmissao)
            lCkPoint = "9"
            Set clsA46V601 = New clsA46V601A
            Set clsA46V601.M46V999 = gM46V999
            
            lCkPoint = "10"
            If Not clsA46V601.gfRecArqTransInsp(lEndArqTransmissao, lMensagem, objstcPedido) Then
                lCkPoint = "11"
                Call gM46V999.gpGraLog(0, lMensagem)
                'marca como ***Processado nÃO OK!
                lCkPoint = "12"
                If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & Mid(lArquivo, 1, 29) & ".N55")) = True Then
                    lCkPoint = "13"
                    Kill gM46V999.gAppPath & Mid(lArquivo, 1, 29) & ".N55"
                End If
                lCkPoint = "14"
                Name gM46V999.gAppPath & lArquivo As gM46V999.gAppPath & Mid(lArquivo, 1, 29) & ".N55"
                'Grava o arquivo um arquivo de log contendo a mensagem de erro
                lCkPoint = "16"
                If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & Mid(lArquivo, 1, 29) & ".ERR")) = True Then
                    lCkPoint = "17"
                    Kill gM46V999.gAppPath & Mid(lArquivo, 1, 29) & ".ERR"
                End If
                lCkPoint = "18"
                Call mpGravaErroTransmissao(lMensagem, Mid(lArquivo, 1, 29) & ".ERR")
            Else
                'marca como ***Processado OK!
                lCkPoint = "19"
                If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & Mid(lArquivo, 1, 29) & ".Y55")) = True Then
                    lCkPoint = "20"
                    Kill gM46V999.gAppPath & Mid(lArquivo, 1, 29) & ".Y55"
                End If
                lCkPoint = "21"
                Name gM46V999.gAppPath & lArquivo As gM46V999.gAppPath & Mid(lArquivo, 1, 29) & ".Y55"
            End If
        End If
        
        'busca o próximo
        lCkPoint = "22"
        lArquivo = Dir(lPathArqs & "*.055")
        
    Loop
    
    Set clsA46V601 = Nothing

SAIDA:
    Err = 0
    On Error GoTo 0
    Exit Sub
TRATAMENTO_ERRO:
    If gM46V999.gfPreenchido(lMensagem) = False Then
        lMensagem = Err & " - " & ERROR
    End If
    lMensagem = "mpCarregaArquivosInspecao Erro: " & lMensagem & vbCrLf
    lMensagem = lMensagem & "CheckPoint: " & lCkPoint
    Call gM46V999.gpGraLog(0, lMensagem)
    GoTo SAIDA
    
End Sub

Private Sub mpVerificaSies()
Dim lrs                 As ADODB.Recordset
Dim lRsProc             As ADODB.Recordset
Dim lConexao            As ADODB.Connection
Dim clsA46V101A         As clsA46V101A
Dim lSelect             As String
Dim lSelectProc         As String
Dim lData               As Long
Dim lDataSist           As Date
Dim pMensagem           As String
Dim lPi                 As String

Err = 0
On Error GoTo Func_Erro
    
    pMensagem = ""
    'Execução após expediente
    lDataSist = Now()
    If DatePart("h", lDataSist) = 18 Then
        Set clsA46V101A = New clsA46V101A
        Set clsA46V101A.M46V999 = gM46V999
        
        If Not gM46V999.gfAbrBasSQL("EAPDB", "2", "LPX20102PSLUSQL", "EAPDB", "workflow", "yasworkflow", lConexao, pMensagem) Then
            GoTo Func_Erro
        End If

        lData = Format(DateAdd("D", -3, Date), "YYYYMMDD")
        lSelect = "select TOP 50 nosso_numero, num_pi from ramosdiversos..ftp_tab_ctrl_emis as e (nolock)"
        lSelect = lSelect & " left join RamosDiversos..Tab_Ped_Num_Copia as c (nolock) on e.nosso_numero = c.NOSSO_NUMERO_OLD"
        lSelect = lSelect & " where c.NOSSO_NUMERO_NEW is null and cod_ramo in (300,620,710) and num_pi <> 0 and tip_cotac in (4,5) and dat_digit >= " & lData
        lSelect = lSelect & " and num_pi not in "
        lSelect = lSelect & " (select num_pi from ged..tab_ctrl_emis (nolock) where cod_ramo in (300,620,710) and num_pi <> 0 and tip_cotac in (4,5) and dat_digit >= " & lData & ")"
        
        Set lrs = New ADODB.Recordset

        lrs.Open lSelect, lConexao
        While Not lrs.EOF
            Set lRsProc = New ADODB.Recordset
            If gM46V999.gfPreenchido(lrs("nosso_numero")) = False Or gM46V999.gfPreenchido(lrs("Num_PI")) = False Then
                lrs.MoveNext
            End If
            
            lSelectProc = " exec ramosdiversos.dbo.proc_StartFluxoEmissaoTransmissaoMulti '" & lrs("nosso_numero") & "'," & lrs("Num_PI")
            lRsProc.Open lSelectProc, lConexao
            lPi = lPi & lrs("Num_PI") & ";"
            lrs.MoveNext
        Wend
        Call gM46V999.gpGraLog(0, "mpVerificaSies - Processou: " & lPi)
    End If

SAIDA:
    Set lrs = Nothing
    Set lRsProc = Nothing
    Err = 0
    On Error GoTo 0
    Exit Sub
  
Func_Erro:
    If gM46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = "mpVerificaSies Error1: " & Err & " | " & ERROR
    Else
        pMensagem = pMensagem & "mpVerificaSies Error2: " & pMensagem & " | proc: " & lSelectProc
    End If
    Call gM46V999.gpGraLog(0, pMensagem)
    GoTo SAIDA
End Sub

