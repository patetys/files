Attribute VB_Name = "Bas_ADO"
Option Explicit
'*********
' Modulo genérico para manipulações com base de dados Access ( .MBD)
' e tambem bases de dados SQL Server - Microsoft
' Ex.:  - Abertura de Banco de dados
'       - Fechamanto do Banco de dados
'       - Consultas, inclusão , exclusão atualização de registros etc...
'
'
'********Para utilização deste módulo adicione as Referencias Abaixo:
'        - Microsoft ActiveX Data Objects 2.8 Lybrary
'        - Microsoft DAO 2.5/3.51 Compatibility Library
'*********
Public gSelect      As String
Public gInsert      As String
Public gUpdate      As String
Public gDelete      As String

Public Enum gCMD_DML
    gEnumDML_SELECT = 0
    gEnumDML_INSERT = 1
    gEnumDML_UPDATE = 2
    gEnumDML_DELETE = 3
End Enum

Public Enum TipoDados
   tpoNumerico = 1
   TpoData = 2
   TpoHora = 3
   TpoValor = 4
   tpoString = 5
   TpoBoolean = 6
   TpoNumInteiro = 7
   TpoNumDecimal = 8
End Enum

Public Enum E_StatusAtInat
    StatusAtInat_ND = 0
    StatusAtInat_Ativo = 1
    StatusAtInat_Inativo = 2
End Enum
Public Function gfComBasDados_DAO(ByVal pEndBasDados As String) As Boolean
    'Função: Compactar base de dados.

    'Parâmetro de entrada...pEndBasDados...Endereço da base de dados.

    Dim lCaracter       As String   'Caracter identificador do endereço.
    Dim lNomBasDados    As String   'Nome da base de dados.
    Dim lNovBasDados    As String   'Nova base de dados.
    Dim lPosicao1       As Byte     'Posição 1 - utilizado com InStr.
    Dim lPosicao2       As Byte     'Posição 2 - utilizado com InStr.
    Dim lSegBasDados    As String   'Nome de segurança da base de dados.

    gfComBasDados_DAO = False

    '1. Posiciona endereço da base de dados.
    lCaracter = "/"
    lPosicao1 = InStr(pEndBasDados, lCaracter)
    If lPosicao1 = 0 Then
        lCaracter = "\"
        lPosicao1 = InStr(pEndBasDados, lCaracter)
    End If
    Do
        lPosicao2 = InStr(lPosicao1 + 1, pEndBasDados, lCaracter)
        If lPosicao2 = 0 Then
            Exit Do
        End If
        lPosicao1 = lPosicao2
    Loop
    lNovBasDados = Left$(pEndBasDados, lPosicao1) & "BD1.MDB"
    lNomBasDados = Mid(pEndBasDados, lPosicao1 + 1, 8)

    '2. Se já existe arquivo BD1.MDB, excluí-lo.
    If gfPreenchido(Dir(lNovBasDados)) Then
        Kill lNovBasDados
    End If

    '3. Compacta arquivo.
    DBEngine.CompactDatabase pEndBasDados, lNovBasDados, , , ";PWD=" & LCase(lNomBasDados) & "_mrc"

    '4. Renomeia arquivo atual para segurança (salva arquivo para efeito de segurança).
    lSegBasDados = Left$(pEndBasDados, lPosicao1) & "XXXXXXXX.MDB"
    Name pEndBasDados As lSegBasDados

    '5. Renomeia arquivo compactado.
    Name lNovBasDados As pEndBasDados

    '6. Exclui arquivo antigo.
    Kill lSegBasDados

    gfComBasDados_DAO = True
End Function

Public Function gfAbrBasDadosAccess(ByVal pIndBasDados As String, _
                                    ByVal pEndBasDados As String, _
                                    ByRef pADOBasDados As ADODB.Connection, _
                                    ByRef pMensagem As String) As Boolean
    'Função: abre arquivo.

    'Parâmetros de entrada...pIndBasDados...Indicador da base de dados.
    '                        pEndBasDados...Endereço da base de dados.
    'Parâmetros de saída.....pADOBasDados...Base de dados.
    '                        pMensagem......Mensagem.

    Dim lAbriu      As Boolean  'Abriu a base de dados?
    Dim lComando    As String   'Comando.
    Dim lDriver(6)  As String   'Drivers.
    Dim lI          As Byte     'Utilizado com For...Next.
    Dim lIniCatalog As String   'Parâmetro INITIAL CATALOG.
    Dim lPwd        As String   'Parâmetro PWD.
    Dim lServer     As String   'Parâmetro SERVER.
    Dim lTemErro    As Boolean  'Tem erro?
    Dim lUid        As String   'Parâmetro UID.
    Dim lMsgErro    As String   'Mensagem de erro

    gfAbrBasDadosAccess = False

    '3. Verifica endereço do arquivo.
    If Not gfPreenchido(pEndBasDados) Then
        pMensagem = "FTE0001 - Erro ao abrir base de dados (" & _
                    " - endereço da base de dados não preenchido)."
        Exit Function
    End If
    
    If Not gfPreenchido(Dir(pEndBasDados)) Then
        pMensagem = "FTE0001 - Erro ao abrir base de dados (" & _
                    " - endereço da base de dados não localizado: " & pEndBasDados & ")."
        Exit Function
    End If

    '4. Abre arquivo.
    lDriver(1) = "Microsoft Access Driver (*.mdb)"
    lDriver(2) = "Microsoft Access Driver(*.mdb)"
    lDriver(3) = "Driver do Microsoft Access (*.mdb)"
    lDriver(4) = "Driver do Microsoft Access(*.mdb)"
    lDriver(5) = "Driver para o Microsoft Access (*.mdb)"
    lDriver(6) = "Driver para o Microsoft Access(*.mdb)"
    lTemErro = False
    lAbriu = False
    lMsgErro = ""
    
    For lI = 1 To 6
        'a) Posiciona conexão.
        On Error Resume Next
        Set pADOBasDados = New ADODB.Connection
        If Err <> 0 Then
            pMensagem = "FTE0001 - Erro ao abrir base de dados (" & _
                        " - comando: SET New ADODB.Connection - Erro " & Err & " - " & Error & ")."
            lTemErro = True
            Exit For
        End If
        
        'b) Posiciona propriedade CursorLocation.
        pADOBasDados.CursorLocation = adUseClient
        If Err <> 0 Then
            pMensagem = "FTE0001 - Erro ao abrir base de dados (" & _
                        " - comando: CursorLocation = adUseClient - Erro " & Err & " - " & Error & ")."
            lTemErro = True
            Exit For
        End If
        
        'c) Abre arquivo.
        lComando = "Driver={" & lDriver(lI) & "};Dbq=" & pEndBasDados
        
        If gfPreenchido(pIndBasDados) = True Then
            'Base de dados Acces com senha generica ( <nome do arquivo> & 21)
            If LCase(pIndBasDados) = "p0043000" Then
                pADOBasDados.Open lComando, , LCase(pIndBasDados) & "21"
            Else
                pADOBasDados.Open lComando, , LCase(pIndBasDados) & "_mrc"
            End If
        Else
            'Base de dados access sem senha
            pADOBasDados.Open lComando
        End If
        
        If Err = 0 Then
            lAbriu = True
            Exit For
        Else
        
            lMsgErro = " - Erro " & Err & " - " & Error
            'o número da msg é o mesmo, por isso estou consistindo pelo nome.
            If InStr(1, UCase(lMsgErro), "CORROMPIDO") > 0 Or InStr(1, UCase(lMsgErro), "DANIFICADO") > 0 Then
                lMsgErro = " - Arquivo corrompido ou danificado."
                Exit For
            ElseIf InStr(1, UCase(lMsgErro), "NÃO É UM CAMINHO VÁLIDO") > 0 Or InStr(1, UCase(lMsgErro), "ESCRITO CORRETAMENTE") > 0 Then
                lMsgErro = " - Problemas de comunicação, verifique comunicação com a rede."
                Exit For
            End If
            
        End If
        On Error GoTo 0
        
    Next lI
    
    If lTemErro Then
        On Error GoTo 0
        Exit Function
    End If
    
    If lAbriu = False Then
        pMensagem = "FTE0001 - Erro ao abrir base de dados (" & pIndBasDados & " - comando: " & _
                    lComando & IIf(lMsgErro = "", " - Erro " & Err & " - " & Error, lMsgErro) & ")."
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    gfAbrBasDadosAccess = True
    
End Function

Public Function gfExeSQL(ByVal pBasDados As ADODB.Connection, ByVal pComando As String, _
                         ByRef pMensagem As String, Optional ByRef pRegAfetados As Long) As Double
    'Função: executa comando SQL.

    'Parâmetros de entrada...pBasDados......base de dados.
    '                        pComando ......comando SQL a ser executado.
    'Parâmetros de saída.....pMensagem......mensagem.
    '                        pRegAfetados...quantidade de registros afetados pelo comando SQL (OPCIONAL).

    'Resultado da função.....Código de retorno do comando SQL: 0 = comando OK.

    gfExeSQL = 999
    
    On Error Resume Next
    
    If IsMissing(pRegAfetados) Then
        pBasDados.Execute pComando
    Else
        pBasDados.Execute pComando, pRegAfetados
    End If
    
    If Err <> 0 Then
        pMensagem = "FTE0027 - Erro ao executar comando SQL: " & Err & " - " & Error & " - " & pComando
    Else
        pMensagem = ""
    End If
    
    gfExeSQL = Err      'Este comando precisa ser o último antes do On Error GoTo 0.
    
    On Error GoTo 0
    
End Function

Public Function gfObtRegistro(ByVal pNomArquivo As ADODB.Connection, ByVal pComando As String, _
                              ByRef pRegistro As ADODB.Recordset, ByRef pMensagem As String) As Boolean
    'Função: obtém registro.

    'Parâmetros de entrada...pNomArquivo...Nome do arquivo.
    '                        pComando .....Comando SELECT.
    'Parâmetros de saída.....pRegistro.....Registro (recordset).
    '                        pMensagem.....Mensagem.

    gfObtRegistro = False

    On Error Resume Next

    Set pRegistro = New ADODB.Recordset

    pRegistro.CursorLocation = adUseClient
    'Abre o Recordset
    pRegistro.Open pComando, pNomArquivo, adOpenForwardOnly, adLockReadOnly

    'Set pRegistro = pNomArquivo.Execute(pComando)
    If Err <> 0 Then
        pMensagem = "FTE0044 - Erro ao obter registro (Erro = " & Err & " - " & Error & _
                        " - comando: " & pComando & ")."
        Exit Function
    End If
    
    If InStr(pNomArquivo, "UID=") = 0 And InStr(pNomArquivo, "Initial Catalog=") = 0 Then
        Set pRegistro.ActiveConnection = Nothing
    End If
    
    Select Case Err
        Case 0      'Sem erro.
            gfObtRegistro = True
        Case 3159   'Marcador inválido - arquivo dados corrompido.
            pMensagem = "FTE0044 - Erro ao obter registro (arquivo corrompido - comando: " & pComando & _
                        ")."
        Case Else
            pMensagem = "FTE0044 - Erro ao obter registro (Erro = " & Err & " - " & Error & _
                        " - comando: " & pComando & ")."
    End Select
    
    On Error GoTo 0
End Function

Public Sub gpFecha_RS(ByRef pTabela As ADODB.Recordset)
    'Procedure: fecha tabela.

    'Parâmetro de entrada...pTabela...Variável Recordset.
    If Not pTabela Is Nothing Then
        pTabela.Close
        Set pTabela = Nothing
    End If
    
End Sub

Public Sub gpFecha_CN(ByRef pBasDados As ADODB.Connection)
    'Procedure: fecha arquivo.

    'Parâmetro de entrada...pBasDados...Variável Database.
    If Not pBasDados Is Nothing Then
        If InStr(1, pBasDados.ConnectionString, ".mdb") > 0 Then
            If pBasDados.State = 1 Then
                pBasDados.Close
            End If
            Set pBasDados = Nothing
        End If
    End If
End Sub
Public Function gfAbrBasSQL(ByVal pIdeBasDados As String, _
                            ByVal pTipBasDados As String, _
                             ByVal pSerBasDados As String, _
                             ByVal pEndBasDados As String, _
                             ByVal pUIDBasDados As String, _
                             ByVal pPWDBasDados As String, _
                             ByRef pADOBasDados As ADODB.Connection, _
                             ByRef pMensagem As String) _
                             As Boolean
    'Função: abre base de dados.

    'Parâmetros de entrada...pIdeBasDados...Identificador da base de dados.
    '                        pEndBasDados...Endereço da base de dados.
    'Parâmetros de saída.....pADOBasDados...Base de dados.
    '                        pMensagem......Mensagem.

    Dim lFlag           As Boolean  'Variavel lógica, indica se o procedimento utilizará ou não a ultima conexão aberta
                                    ' TRUE = Abre nova conexão ; FALSE = Mantem a ultima conexão aberta
    Dim lFlag2          As Boolean  'VAriavel lógica auxiliar
                                    'Utilizada para verificar da solicitação de abertura de Base
                                    'Caso a verificação encontre divergencias.
                                    'TRUE = encontrado divergencia ; FALSE = Não encontrado Divergencia Continua a verificação
    Dim lInd            As Integer  'Contador.
    Dim lInd2           As Integer  'Contador auxiliar
    Dim lIndMax         As Integer  'Maior índice.
    Dim lInfoConexao()  As String   'informação sobre a conexão.
    Dim lSubInfoCn()    As String   'Recebe itens da variavel lInfoConexao
    Dim lIniCatalog     As String   'Parâmetro INITIAL CATALOG.
    Dim lMensagem       As String   'Mensagem de log.
    Dim lOpen           As String   'Linha para comando OPEN.
    
    Static ls_ConexaoSQL  As ADODB.Connection
    
    '1. Verifica se a conexão foi a ultima a ser aberta
    If ls_ConexaoSQL Is Nothing Then
        'ls_ConexaoSQL = Nothing
        'Atribui true para a variavel lFlag
        'Para abrir nova conexão
        lFlag = True
    Else
        'separa as informações da ultima conexão aberta em uma variavel matriz
        lInfoConexao = Split(ls_ConexaoSQL, ";")
        
        'Inicia a variavel para abrir nova conexão
        lFlag = True
        
        'Inicia a variavel para prosseguir com a verificação
        lFlag2 = False
        
        For lInd = LBound(lInfoConexao()) To UBound(lInfoConexao())
            
            If UBound(lInfoConexao()) < 0 Then
                'Se a String estiver vazia
                'encerra a verificação, e o procedimento abrirá nova conexão
                Exit For
            End If
                        
            lSubInfoCn = Split(lInfoConexao(lInd), "=")
                
            For lInd2 = LBound(lSubInfoCn()) To 0
                
                If UBound(lSubInfoCn()) < 0 Then
                    Exit For
                End If
                
                Select Case lSubInfoCn(lInd2)
                    Case "SERVER"
                        If lInfoConexao(lInd) = "SERVER=" & pSerBasDados Then
                           'servidores de base de dados são iguais
                           'Atribui false para a Variavel lFlag.
                           'Para manter a ultima conexão aberta
                           lFlag = False
                        Else
                            'Encontrado Divergencia
                            'Atribui TRUE para a Variavel lFlag2
                            'Para encerrar a verificação
                            lFlag2 = True
                            'Atribui TRUE para abrir nova conexão
                            lFlag = True
                        End If
                    Case "UID"
                        If lInfoConexao(lInd) = "UID=" & pUIDBasDados Then
                            'Usuarios são iguais
                            lFlag = False
                        Else
                            lFlag2 = True
                            lFlag = True
                        End If
                    Case "PWD"
                        If lInfoConexao(lInd) = "PWD=" & pPWDBasDados Then
                            'Senhas de acesso ao Banco são iguias
                            lFlag = False
                        Else
                            lFlag2 = True
                            lFlag = True
                        End If
                    Case "DATABASE"
                        If Replace(lInfoConexao(lInd), """", "") = "DATABASE=" & pEndBasDados Then
                            'Bancos de Dados são iguais
                            lFlag = False
                        Else
                            lFlag2 = True
                            lFlag = True
                        End If
                End Select
                
            Next lInd2
            
            If lFlag2 = True Then
                'Encontrado divergencia
                'lFlag é igual a TRUE, o procedimento abrirá nova conexão
                'Encerra a verificação
                Exit For
            End If
                            
        Next lInd
        
    End If

    gfAbrBasSQL = False

    If lFlag = True Then
        '2. Verifica identificação da base de dados.
        If Not gfPreenchido(pIdeBasDados) Then
            pMensagem = "FTE0001 - Erro ao abrir base de dados (identificação da base de dados não " & _
                        "preenchida)."
            Exit Function
        End If

        '4. Verifica endereço da base de dados.
        If Not gfPreenchido(pEndBasDados) Then
            pMensagem = "FTE0001 - Erro ao abrir base de dados (endereço da base de dados " & pEndBasDados & _
                        " não preenchido)."
            Exit Function
        End If

        '5. Abre base de dados.
        'a) Posiciona conexão.
        On Error Resume Next
        Set ls_ConexaoSQL = New ADODB.Connection
        If Err <> 0 Then
            pMensagem = "FTE0001 - Erro ao base de dados (" & pIdeBasDados & _
                        " - comando SET New ADODB.Connection - erro " & Err & " - " & Error & ")."
            On Error GoTo 0
            Exit Function
        End If

        ls_ConexaoSQL.ConnectionString = "driver={SQL Server};server=" & pSerBasDados & _
                                    ";uid=" & pUIDBasDados & ";pwd=" & pPWDBasDados & _
                                    ";database=" & pEndBasDados

        ls_ConexaoSQL.ConnectionTimeout = 60
        ls_ConexaoSQL.CommandTimeout = 60
        'b) Posiciona propriedade CursorLocation.
        'pADOBasDados.CursorLocation = adUseClient
        If Err <> 0 Then
            pMensagem = "FTE0001 - Erro ao base de dados (" & pIdeBasDados & _
                        " - Comando CursorLocation = adUseClient - erro: " & Err & " - " & Error & ")."
            On Error GoTo 0
            Exit Function
        End If
        
        'c) Abre base de dados.
        ls_ConexaoSQL.Open
        If Err <> 0 Then
            pMensagem = "FTE0001 - Erro ao base de dados (" & pIdeBasDados & _
                        " - comando Open " & lOpen & " - erro: " & Err & " - " & Error & ")."
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    Set pADOBasDados = ls_ConexaoSQL

    gfAbrBasSQL = True
End Function

Public Function mfFormatarCamposSQL(pAtributo As Variant, Optional SemDesformatar As Boolean = False) As String

    Select Case VarType(pAtributo)
        Case vbEmpty, vbNull:
            mfFormatarCamposSQL = "Null"
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte:
            mfFormatarCamposSQL = Trim(Str(pAtributo))
        Case vbString:
            If gfPreenchido(pAtributo) = False Then
                mfFormatarCamposSQL = "' '"
            Else
                If SemDesformatar Then
                    mfFormatarCamposSQL = "'" & Replace(CStr(pAtributo), "'", " ") & "'"
                Else
                    mfFormatarCamposSQL = "'" & UCase(gfRetirarCaracterEspecial(CStr(pAtributo))) & "'"
                End If
            End If
        Case vbDate:
            If Val(pAtributo) = 0 Then
                mfFormatarCamposSQL = "Null"
            Else
                mfFormatarCamposSQL = "'" & Format(pAtributo, "yyyy-mm-dd") & " " & Format(pAtributo, "hh:mm:ss") & "'"
            End If
        Case vbBoolean:
            mfFormatarCamposSQL = IIf(pAtributo = True, "1", "0")
    End Select

End Function

Public Function gfWhereAnd(pSql As String) As String
    If InStr(1, UCase(pSql), "WHERE") = 0 Then
        gfWhereAnd = pSql & " WHERE "
    Else
        gfWhereAnd = pSql & " AND "
    End If
End Function
Public Function gfAbrBasDados(ByVal pIndBasDados As String, ByVal pTipBasDados As String, _
                              ByVal pEndBasDados As String, ByRef pADOBasDados As ADODB.Connection, _
                              ByRef pMensagem As String) As Boolean
    'Função: abre arquivo.

    'Parâmetros de entrada...pIndBasDados...Indicador da base de dados.
    '                        pTipBasDados...Tipo da base de dados.
    '                        pEndBasDados...Endereço da base de dados.
    'Parâmetros de saída.....pADOBasDados...Base de dados.
    '                        pMensagem......Mensagem.
    Dim lSqlServer As Boolean
    gfAbrBasDados = False
    
    '1. Verifica indicador do arquivo.
    
End Function

Public Sub gpFecha1(ByRef pTabela As ADODB.Recordset)
    'Procedure: fecha tabela.

    'Parâmetro de entrada...pTabela...Variável Recordset.
    If Not pTabela Is Nothing Then
        If pTabela.State = 1 Then
            pTabela.Close
        End If
    End If
    Set pTabela = Nothing
End Sub
Public Sub gpFecha3(ByRef pTabela As ADODB.Recordset, ByRef pBasDados As ADODB.Connection)
    'Procedure: fecha tabela e arquivo.

    'Parâmetros de entrada...pTabela.....Recordset da tabela.
    '                        pBasDados...Variável Database.

    '1. Fecha tabela.
    If pTabela.State = 1 Then
        pTabela.Close
    End If
    Set pTabela = Nothing

    '2. Fecha arquivo.
    If InStr(1, UCase(pBasDados.ConnectionString), ".MDB") > 0 Then
        If pBasDados.State = 1 Then
            pBasDados.Close
        End If
        Set pBasDados = Nothing
    End If
End Sub


Public Sub gpFecha2(ByRef pBasDados As ADODB.Connection)
    'Procedure: fecha arquivo.

    'Parâmetro de entrada...pBasDados...Variável Database.
    If Not pBasDados Is Nothing Then
        If InStr(1, UCase(pBasDados.ConnectionString), ".MDB") > 0 Then
            If pBasDados.State = 1 Then
                pBasDados.Close
            End If
            Set pBasDados = Nothing
        End If
    End If
End Sub

