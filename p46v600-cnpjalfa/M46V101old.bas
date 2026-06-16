Attribute VB_Name = "basM46V101"
Option Explicit
Public gAltT46V112A         As Boolean          'Indica que houve alteração ao acessar form T46V112A.
Public gAltComum            As Boolean          'Indica que houve alteração nos dados de comum.
Public gAltDadosCobert      As Boolean          'Indica que houve alteração nos principais dados e é necessário retornar os valores das franquias
Public gCarga               As Boolean          'Indicador de carga de registro.
Public gCodCorrInst         As Long             'Código do corretor de instalação.
Public gCodCorretor         As Long             'Código do corretor.
Public gIndCosseguro        As Byte             'Índice do cosseguro.
Public gIteUltData          As Double           'Item - última atualização: data.
Public gIteUltUsuario       As String           'Item - última atualização: usuário.
Public gNomCorretor         As String           'Nome do corretor.
Public gProChamado          As String           'Programa chamado.
Public gProChamador         As String           'Programa chamador.
Public gTemErro_T46V123A    As Boolean          'Indica que tem erro na montagem do form T46V123A.
Public gTemErro_T46V124A    As Boolean          'Indica que tem erro na montagem do form T46V124A.
Public gUsuarioFilial       As Boolean          'Usuario True = Usuario filial
Public gCarregaCusto        As Boolean          'True = carregar custo default
Private mNumArqRetorno      As Integer          'Utilizado com FreeFile para arquivo de retorno.(ipsocx.dll)

Public AlterarCobertura    As e_SimNao
Public IncluirCobertura    As e_SimNao
Public AlterarIs           As e_SimNao

'Public gNumPI              As Long

Public gCod_RamoCorr         As e_CodRamoTip   'Códigos do ramo de Automóvel para Auto (310), RC (530)  e APP (812). Controle de corretores

Public Enum e_CodRamoTip
    e_CodRamoAuto = 310
    e_CodRamoRC = 530
    e_CodRamoapp = 812
End Enum
    
Public Enum e_Tipo_Espec
    e_Tipo_Espec_Objetos = 1
    e_Tipo_Espec_RCobertos = 2
    e_Tipo_Espec_Caracteristica = 3
    e_Tipo_Espec_Perimetro = 4
    e_Tipo_Espec_Especificacao = 5
    e_Tipo_Espec_Coberturas = 6
    e_Tipo_Espec_Franquia = 7
    e_Tipo_Espec_Clausulas = 8
    e_Tipo_Espec_Condicao = 9
    e_Tipo_Espec_OutrasInf = 10
    e_Tipo_Espec_Observacao = 11
    e_Tipo_Espec_PrazoObra = 12
    e_Tipo_Espec_PrazoManut = 13
    e_Tipo_Espec_Contratante = 14
    e_Tipo_Espec_LocalExpos = 15
    e_Tipo_Espec_Nenhuma = 99
End Enum

'Ficha 1000307
Public gObrRegistro    As Boolean      'Obrigatório informar o código do registro no conselho e o número

'Proposta Eletrônica
Public mAntCNPJ        As String   'Anterior: CPF/CNPJ
Public mAntSeg         As String   'Anterior: Segurado
Public mAtuCNPJ        As String   'Atual: CPF/CNPJ

Public Function gTip_Atividade() As String
Dim BDP00Tarifa       As ADODB.Connection
Dim rstTabClas        As clsYasRecordSet
Dim lMensagem         As String
Dim lCod_Clasf        As String
Dim lCod_Compl_Clasf  As String
Dim lSelect           As String
    
    lCod_Clasf = Left(frmT46V101A.txtClassificacao.Tag, 4)
    lCod_Compl_Clasf = Mid(frmT46V101A.txtClassificacao.Tag, 5, 2)
    
    gTip_Atividade = 0
    
    If Trim(lCod_Clasf) <> "" And Trim(lCod_Clasf) <> "0" Then
        '1. Abre arquivo.
        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Function
        End If

        '3. Obtém produto sem término de vigência ou com término de vigência superior à data do dia.
        lSelect = " Select distinct dsc_clasf, COD_CLASF, " & _
                "COD_COMPL_CLASF , tab_Clasf, Tip_Atividade  from Tbm_Classificacao " & _
                " Where cod_Ramo  = " & gM46V999.gCod_Ramo & _
                " and cod_Clasf = '" & lCod_Clasf & "'" & _
                " and Cod_Compl_Clasf = '" & lCod_Compl_Clasf & "'" & _
                " and Cod_Tarif = " & gM46V999.gCodProduto & _
                " and Cod_Sequencia <= " & gM46V999.gCodSequencia
        lSelect = lSelect & " and  tab_clasf = " & gM46V111.gCod_Plano
        
        If gM46V999.gfObtRegistro(BDP00Tarifa, lSelect, rstTabClas, lMensagem) = False Then
            Call gM46V999.gpFecha2(BDP00Tarifa)
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Function
        End If

        If Not rstTabClas.EOF Then
            gTip_Atividade = rstTabClas("Tip_Atividade")
        End If
    End If
    
End Function

Public Function gfExisteParmJuros(ByVal pCorretor As Long, _
                                    ByRef pExiste As Boolean, _
                                    pMensagem As String) As Boolean
                                    
Dim bdP00MULTI  As ADODB.Connection
Dim lrst        As clsYasRecordSet
Dim lSelect     As String

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, bdP00MULTI, _
            pMensagem) = False Then Exit Function
    
    lSelect = "Select count(*) as total from TAB_PARM_DIVERSOS"
    lSelect = lSelect & " Where Cod_Produto = " & gM46V999.gCodProduto
    lSelect = lSelect & " and Cod_Protocolo = " & gM46V999.gCodProtocolo
    lSelect = lSelect & " and Cod_Sequencia <= " & gM46V999.gCodSequencia
    lSelect = lSelect & " and cod_ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " and TIPO_VALOR = 7"
    lSelect = lSelect & " and CHAVE_1 = '" & pCorretor & "'"
    
    If gM46V999.gfObtRegistro(bdP00MULTI, lSelect, lrst, _
            pMensagem) = False Then Exit Function
    
    If lrst("Total") = 0 Then
        pExiste = False
    Else
        pExiste = True
    End If
    
    gfExisteParmJuros = True
End Function

Public Function gpPosicaoColecao(pItem As Long) As Long
Dim lCount As Long

    For lCount = 1 To frmT46V125A.flxItens.Rows - 1
        If pItem = frmT46V125A.flxItens.TextMatrix(lCount, 0) Then
            gpPosicaoColecao = lCount
            Exit Function
        End If
    Next
    'Se houve alteração de número e o número for menor que algum dos itens...
    For lCount = 1 To frmT46V125A.flxItens.Rows - 1
        If pItem < frmT46V125A.flxItens.TextMatrix(lCount, 0) Then
            gpPosicaoColecao = lCount
            Exit Function
        End If
    Next
    'Se houve alteração de número e o número for maior que os demais...
    For lCount = 1 To frmT46V125A.flxItens.Rows - 1
        If pItem > frmT46V125A.flxItens.TextMatrix(lCount, 0) Then
            gpPosicaoColecao = lCount
        End If
    Next
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

Public Sub Main()
    Dim lbdP00MPED      As ADODB.Connection 'Base de dados: P0042300.
    Dim lCodCorretor    As String           'Código do corretor.
    Dim lCodUsuario     As String  'Long    'Código do usuário.
    Dim lDatModulo      As Double           'Data do módulo.
    Dim lDatVersao      As String           'Data da versão.
    Dim lEndP2842400    As String           'Endereço do arquivo P2842400.EXE.
    Dim lLargura        As Integer          'Largura.
    Dim lMensagem       As String           'Mensagem.
    Dim lNumP2842400    As Integer          'Utilizado com FreeFile para o arquivo P2842400.EXE.
    Dim lPosicao        As Integer          'Utilizado com InStr.
    Dim lRegistro       As String           'Registro.
    Dim objTabUser      As clsA46V725A      'Objeto da classe clsA28V713A (Tab_User)
    Dim objstcTabUser   As stcA46V725B      'Objeto da classe de estrutura stcA28V713B (Tab_User)

    Screen.MousePointer = vbHourglass

    '1. Verifica se form já está ativo.
    If App.PrevInstance = True Then
        Call gM46V999.gpGraLog(2, "AGE0003 - O programa de Cálculo SYAS - MULTISSEGUROS já está ativo neste " & _
                         "microcomputador." & vbLf & vbLf & "Utilize o programa que já está ativo.")
        End
    End If
    
    '2. Obtém informações iniciais do sistema: tipo do usuário, endereços das bases de dados, endereços
    '   dos arquivos de impressão, indicador de uso do auxílio automático, nome ou código do usuário.
    Call gM46V999.gpInicia("")

    'gM46V999.EhTeste = True
    
    If gM46V999.EhTeste Then
        MsgBox "Versão de Teste"
    End If
    
    '3. Emite mensagem de entrada (precisa ser depois de obter dados do arquivo de inicialização).
    Call gM46V999.gpGraLog(0, "AGE0001 - Início do processamento do programa " & App.EXEName & ".")

    '4. Se data do sistema incoerente com data do módulo executável, encerra processo.
    If Dir(gM46V999.gAppPath & "P46V100.EXE") <> "" Then
        lDatModulo = Format$(FileDateTime(gM46V999.gAppPath & "P46V100.EXE"), "yyyymmdd")
        If lDatModulo > gM46V999.gHoje Then
            Call gM46V999.gpGraLog(2, "FTE2008 - Data do microcomputador (" & gM46V999.gForHoje & _
                             ") menor que data do programa (" & gM46V999.gfForData(2, lDatModulo) & ")." & vbLf & _
                             vbLf & "Programa será encerrado.")
            End     'Encerra programa.
        End If
    End If
    '5. Posiciona variável gM46V999.gVerDataMaq (indica se a função que verifica se houve ajuste na máquina
    '   deve ser processada: 0 = não verifica e 1 = verifica).
    If App.EXEName <> "P46V101" Then
        gM46V999.gVerDataMaq = 1
    Else
        gM46V999.gVerDataMaq = 0
    End If
    gM46V999.gTrvSistema = False

    'Carrego o perfil do usuário de digitação
    gM46V999.PerfilAcesso = e_Perfil_CALC
    gM46V999.gTipoAcesso = CALCULO
    gM46V999.VerificarNormas = True

    '6. Carrega perfil do usuário de digitação.
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
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
    'lDatVersao = "01/02/2008"
    lDatVersao = Format(FileDateTime(gM46V999.gAppPath & App.EXEName & ".EXE"), "dd/mm/yyyy") 'Data da versão: dd/mm/aaaa.
    'b) Obtém código do corretor.
    Select Case gM46V999.gTipUsuario
        Case 1      'Yasuda.
            lCodCorretor = "YASUDA"
        Case 2      'Corretor.
            lEndP2842400 = gM46V999.gAppPath & "P2842400.EXE"
            If gM46V999.gfPreenchido(Dir(lEndP2842400)) = False Then
                lEndP2842400 = gM46V999.gAppPath & "P2842401.EXE"
            End If
            'Obs.: (a) Se existe o arquivo P2842400.EXE, obtém o código do corretor do registro.
            '      (b) Se não existe o arquivo P2842400.EXE, solicitar ao usuário.
            If gM46V999.gfPreenchido(Dir(lEndP2842400)) Then
                lNumP2842400 = FreeFile
                Open lEndP2842400 For Input Access Read As lNumP2842400
                Line Input #lNumP2842400, lRegistro
                lCodCorretor = Format(lRegistro, "0000000")
                gM46V999.gCod_Corretor = Val(lCodCorretor)
                Close lNumP2842400
            Else
                Call gM46V999.gpGraLog(1, "FGE9003 - " & gM46V999.gPrefixo & _
                                 "Código do corretor inválido (não localizado)." & vbLf & vbLf & _
                                 "O sistema será encerrado.")
                End
            End If
            If gM46V999.gNomUsuario = "XXXXXXXX" Then
                If gfBuscaUsuarioTransmissao(lCodUsuario, lMensagem) = False Then
                    Call gM46V999.gpGraLog(1, lMensagem & vbLf & vbLf & "O sistema será encerrado.")
                    End
                End If
                gM46V999.gNomUsuario = lCodUsuario
                gM46V999.gNomUsuario6 = Left$(gM46V999.gNomUsuario, 6)
            End If
    End Select
    'c) Monta rodapé.
    Load frmT46V100A
    
    With frmT46V100A
        .Show
        .panRodape.Visible = True
        .panForm.Caption = "T46V001A"
        lPosicao = InStr(UCase(gM46V999.gEndP00MPED), "P00MPED.MDB")
        Select Case lPosicao
            Case 0
                .panBasDados.Caption = "Não encontrado"
            Case 1
                .panBasDados.Caption = gM46V999.gAppPath
            Case Else
                .panBasDados.Caption = Left$(gM46V999.gEndP00MPED, lPosicao - 1)
        End Select
        .panVersao.Caption = "Versão " & lDatVersao
        If UCase(lCodCorretor) = "YASUDA" Then
            .panCorretor.Caption = lCodCorretor
            gCodCorrInst = 0
        Else
            .panCorretor.Caption = "Corretor " & lCodCorretor
            gCodCorrInst = Val(lCodCorretor)
        End If
        .panUsuario.Caption = "Usuário " & gM46V999.gNomUsuario6
        .panData.Caption = gM46V999.gForHoje
        lLargura = .panForm.Width + .panBasDados.Width + .panVersao.Width + .panCorretor.Width + _
                   .panUsuario.Width + .panData.Width
        If .panRodape.Width > lLargura Then
            .panBasDados.Width = .panRodape.Width + .panBasDados.Width - lLargura - 120
        End If
        .panForm.Left = 60
        .panBasDados.Left = .panForm.Left + .panForm.Width
        .panVersao.Left = .panBasDados.Left + .panBasDados.Width
        .panCorretor.Left = .panVersao.Left + .panVersao.Width
        .panUsuario.Left = .panCorretor.Left + .panCorretor.Width
        .panData.Left = .panUsuario.Left + .panUsuario.Width
        .Refresh
    End With
    
'    gNomeTabImport = ""
    
    frmT46V100A.mfVerificaVersao
    
    Call gM46V999.gpGraLog(0, "Endereço " & gM46V999.gEndP00MULTI & ", Tipo " & gM46V999.gTipoHomologacao)
    '10. Posiciona endereço do arquivo SYASPROD.HLP.
    App.HelpFile = gM46V999.gAppPath & "SYASMULT.HLP"

    '11. Posiciona controle.
    gM46V111.gSaida = False
    gM46V999.gVerDataMaq = 0
    If gM46V999.gTrvSistema = True Then
        End
    End If

    
    lMensagem = ""
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
    
        '12. Registro do Microfocus
        If gfRegistroMicrofocus(lMensagem) = False Then
            Call gM46V999.gpGraLog(0, lMensagem)
        End If
        
        '[FRH] - Ficha.900530
        Dim Y As Long
        Dim X As Long
        Dim ct_A As Long
        Dim lQtdRegs As Integer
        Dim pMensagem As String
        Dim lSql As String
        
        If gM46V999.gfAbrBasDadosAccess("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, lbdP00MPED, pMensagem) = False Then Exit Sub
                
            lSql = ""
            lSql = "SELECT ID FROM PRC900530"
            On Error Resume Next
            lbdP00MPED.Execute lSql, lQtdRegs
                       
            If lQtdRegs = 0 Then
            
                lSql = ""
                lSql = "CREATE TABLE PRC900530(ID INT)"
                lbdP00MPED.Execute lSql
                
                lSql = ""
                lSql = "INSERT INTO PRC900530(ID) VALUES(1)"
                lbdP00MPED.Execute lSql
                
                ct_A = 1
                Do While ct_A <= 9
                
                    If ct_A = 1 Then X = 2: Y = 8
                    If ct_A = 2 Then X = 3: Y = 2
                    If ct_A = 3 Then X = 4: Y = 5
                    If ct_A = 4 Then X = 5: Y = 4
                    If ct_A = 5 Then X = 6: Y = 9
                    If ct_A = 6 Then X = 7: Y = 3
                    If ct_A = 7 Then X = 8: Y = 6
                    If ct_A = 8 Then X = 9: Y = 10
                    If ct_A = 9 Then X = 10: Y = 12
                    
                    lSql = ""
                    lSql = "UPDATE TAB_PED SET COD_DOC=" & X & " WHERE COD_DOC=" & Y
                    lbdP00MPED.Execute lSql
                    
                    ct_A = ct_A + 1
                Loop
            End If
    End If
        
    '13. Emite janela de proposta.
    frmT46V101A.Show vbModeless, frmT46V100A
    Call gM46V999.gpFecha2(lbdP00MPED)

    Screen.MousePointer = vbDefault
End Sub



Public Function mfFileCopy(ByVal pOrigem As String, _
                            ByVal pDestino As String, _
                            Optional pShowMsg As Boolean = True) As Boolean

Dim lMsgErr             As String
Dim lUltComando         As String
Dim lCkPoint            As String
Dim lMsgSN              As String
Dim I                   As Integer
Inicio:
Err = 0
lMsgErr = ""
'On Error GoTo TRATAMENTO_ERRO
    mfFileCopy = False
    I = 1
    
    On Error Resume Next
    Err = 0
TentaNovamente:
    Call FileCopy(pOrigem, pDestino)
    If Err <> 0 Then
        If I > 3 Then GoTo TRATAMENTO_ERRO
        I = I + 1
        DoEvents
        Call gM46V999.gpGraLog(0, "Deu erro na copia, vai tentar novamente. Tentativa: " & I & " Arquivo: " & pOrigem)
        Err = 0
        GoTo TentaNovamente
    End If
    
    mfFileCopy = True
SAIDA:
    Err = 0
    On Error GoTo 0
    Exit Function

TRATAMENTO_ERRO:
    lMsgSN = "O sistema não conseguiu copiar o arquivo." & vbCrLf
    lMsgSN = lMsgSN & "Origem  : " & pOrigem & vbCrLf
    lMsgSN = lMsgSN & "Destino : " & pDestino & vbCrLf & vbCrLf
    lMsgSN = lMsgSN & "Possíveis causas para o problema:" & vbCrLf
    
    Select Case Err
        Case 52 'Bad file name or number
            lMsgSN = lMsgSN & "O caminho (origem ou destino) está especificado incorretamente ou não existe."
        Case 53 'File not found
            lMsgSN = lMsgSN & "O arquivo de origem não foi localizado."
        Case 70 'Permission denied
            lMsgSN = lMsgSN & "Acesso negado. Privilégios insuficientes no diretório ou o arquivo está em uso."
        Case 75 'Path/File access error
            lMsgSN = lMsgSN & "O arquivo tem atributo 'somente leitura', o caminho não esta especificado corretamente " & vbCrLf
            lMsgSN = lMsgSN & "ou o arquivo está em uso."
        Case 76 'Path not found
            lMsgSN = lMsgSN & "O caminho não foi localizado, está indisponível ou não existe."
        Case Else
            'error não previstos
            If gM46V999.gfPreenchido(lMsgErr) = False Then
                lMsgErr = "Erro não previsto" & Err & " - " & Error
            End If
            lMsgErr = lMsgErr & vbCrLf & vbCrLf & "Informarção complementar: " & vbCrLf & lUltComando
            lMsgErr = lMsgErr & vbCrLf & "AppPath:" & gM46V999.gAppPath
            
            lMsgSN = lMsgSN & lMsgErr
    End Select
    Err.Clear
        
    If pShowMsg = True Then
        If MsgBox(lMsgSN, vbYesNo + vbInformation) = vbYes Then
            If mfFileCopy(pOrigem, pDestino, pShowMsg) = True Then
                mfFileCopy = True
            End If
        Else
        End If
    Else
    End If
    
    GoTo SAIDA
    
End Function

Function mfVerificaStatAcao(ByVal pNumPI As String, ByRef pAcao As Integer, ByRef pMensagem As String) As Boolean
    Dim lbdBasDados          As ADODB.Connection 'Base de dados.
    Dim lrsTabela            As clsYasRecordSet  'Registro da tabela.
    Dim lSql                 As String
    Dim lI                   As Integer
    Dim lNum_pi As Long
    Dim lNosNumero As String
    
    On Error GoTo mfVerificaStatAcao_Err

    mfVerificaStatAcao = False
    pAcao = 0
    
    If Val(pNumPI) = 0 Then
        pMensagem = "Número do PI não pode ser zero!"
        Exit Function
    End If
    
    If gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, lbdBasDados, pMensagem) = False Then Exit Function
    
    lSql = "Select  cod_iden_stat_acao  "
    lSql = lSql & " from tab_ctrl_emis where num_pi = " & pNumPI

    If gM46V999.gfObtRegistro(lbdBasDados, lSql, lrsTabela, pMensagem) = False Then
        Call gM46V999.gpFecha2(lbdBasDados)
        Exit Function
    End If

    If lrsTabela.EOF Then
        mfVerificaStatAcao = True
        GoTo Finalizar
    End If
    
    pAcao = lrsTabela("cod_iden_stat_acao")
    mfVerificaStatAcao = True
    
Finalizar:
    Set lrsTabela = Nothing
    Set lbdBasDados = Nothing
    
    Exit Function
    
mfVerificaStatAcao_Err:
    MsgBox "Erro no módulo de  Verificação do CodIdenStatAcao "
End Function

Public Function gfCarCboBancos(ByVal pOpcao As Byte, ByVal pCombo As ComboBox, _
                               ByRef pMensagem As String, Optional pCodCorretor As Long) As Boolean
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
            Select Case pCodCorretor
'                Case 8001, 8002, 8003, 8004, 8005, 8006, 8008, 8011, 8014, 8028 'Volkswagem
'                    pCombo.AddItem "237 - BRADESCO SA"
'                    pCombo.ItemData(pCombo.NewIndex) = 237
                Case 15701  'Mercedes
                    pCombo.AddItem "341 - ITAU"
                    pCombo.ItemData(pCombo.NewIndex) = 341
                Case Else
                    pCombo.AddItem "399 - HSBC"
                    pCombo.ItemData(pCombo.NewIndex) = 399
            End Select
            'pCombo.AddItem "341 - ITAU"
            'pCombo.ItemData(pCombo.NewIndex) = 341
            'pCombo.AddItem "347 - SUDAMERIS BRASIL"
            'pCombo.ItemData(pCombo.NewIndex) = 347
            'pCombo.AddItem "409 - UNIBANCO"
            'pCombo.ItemData(pCombo.NewIndex) = 409
            '09029
            'pCombo.AddItem "356 - REAL"
            'pCombo.ItemData(pCombo.NewIndex) = 356
            
        Case 3      'Bancos para débito em conta
            
            'Projeto 11080 - Debito automatico - Banco do Brasil
            pCombo.AddItem "001 - BANCO DO BRASIL"
            pCombo.ItemData(pCombo.NewIndex) = 1
            
            pCombo.AddItem "237 - BRADESCO SA"
            pCombo.ItemData(pCombo.NewIndex) = 237
            pCombo.AddItem "341 - ITAU"
            pCombo.ItemData(pCombo.NewIndex) = 341
            'pCombo.AddItem "347 - SUDAMERIS BRASIL"
            'pCombo.ItemData(pCombo.NewIndex) = 347
            '1001220 - Tarifa 1002 e 1003 Janeiro 2011  (desabilitado Unibanco - para débito em conta)
'            pCombo.AddItem "409 - UNIAO DE BANCOS"
'            pCombo.ItemData(pCombo.NewIndex) = 409
            'pCombo.AddItem "479 - BOSTON"
            'pCombo.ItemData(pCombo.NewIndex) = 479
            '09029
            'pCombo.AddItem "356 - REAL"
            'pCombo.ItemData(pCombo.NewIndex) = 356
            
            pCombo.AddItem "399 - HSBC"
            pCombo.ItemData(pCombo.NewIndex) = 399
            
            
            pCombo.AddItem "033 - SANTANDER"
            pCombo.ItemData(pCombo.NewIndex) = 33
            
            pCombo.AddItem "047 - BANESE"
            pCombo.ItemData(pCombo.NewIndex) = 47
        Case 4      'Bancos de indicador de seguro.
            pCombo.AddItem "215 "
            pCombo.ItemData(pCombo.NewIndex) = 215
            pCombo.AddItem "341 - ITAU"
            pCombo.ItemData(pCombo.NewIndex) = 341
            'pCombo.AddItem "347 - SUDAMERIS BRASIL"
            'pCombo.ItemData(pCombo.NewIndex) = 347
            '1001220 - Tarifa 1002 e 1003 Janeiro 2011  (desabilitado Unibanco)
'            pCombo.AddItem "409 - UNIAO DE BANCOS"
'            pCombo.ItemData(pCombo.NewIndex) = 409
            '09029
            pCombo.AddItem "356 - REAL"
            pCombo.ItemData(pCombo.NewIndex) = 356
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
        Case "001" 'Projeto 11080 - Debito automatico - Banco do Brasil
'            If frmT28V110A.Caption <> "Crédito em conta" Then
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
                If Not IsNumeric(Mid(lAuxiliar, 1, 4)) Then
                    lTemErro = True
                End If
                If Not IsNumeric(Mid(lAuxiliar, 6, 1)) Then
                    If UCase(Mid(lAuxiliar, 6, 1)) <> "X" Then
                        lTemErro = True
                    End If
                End If
'            Else
'                If Len(lAuxiliar) < 4 Then
'                    lAuxiliar = String(4 - Len(lAuxiliar), "0") & lAuxiliar
'                End If
'                If Not IsNumeric(Mid(lAuxiliar, 1, 4)) Then
'                    lTemErro = True
'                End If
'            End If
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
        Case "33", "033"     'SAntander
            lPosicao = InStr(lAuxiliar, "-")
            If Len(lAuxiliar) <> 10 Or lPosicao <> 9 Then
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
                If Len(lAuxiliar) < 9 Then
                    lAuxiliar = String(9 - Len(lAuxiliar), "0") & lAuxiliar
                End If
                lAuxiliar = Mid$(lAuxiliar, 1, 8) & "-" & Mid$(lAuxiliar, 9, 1)
            End If
            If Not IsNumeric(Mid(lAuxiliar, 1, 8)) Or Not IsNumeric(Mid(lAuxiliar, 10, 1)) Then
                Select Case pOpcao
                    Case 1
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta para débito. " & _
                                    "Número de conta corrente inválido (Santander - formato NNNNNNNN-N)."
                    Case 2
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta do indicador do seguro. " & _
                                    "Número de conta corrente inválido (Santander - formato NNNNNNNN-N)."
                End Select
                pSaida = ""
                Exit Function
            End If
        Case "47", "047"     'Banese
            lPosicao = InStr(lAuxiliar, "-")
            If Len(lAuxiliar) <> 10 Or lPosicao <> 9 Then
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
                If Len(lAuxiliar) < 9 Then
                    lAuxiliar = String(9 - Len(lAuxiliar), "0") & lAuxiliar
                End If
                lAuxiliar = Mid$(lAuxiliar, 1, 8) & "-" & Mid$(lAuxiliar, 9, 1)
            End If
            If Not IsNumeric(Mid(lAuxiliar, 1, 8)) Or Not IsNumeric(Mid(lAuxiliar, 10, 1)) Then
                Select Case pOpcao
                    Case 1
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta para débito. " & _
                                    "Número de conta corrente inválido (Banese - formato NNNNNNNN-N)."
                    Case 2
                        pMensagem = "FGE0025 - " & gM46V999.gPrefixo & "Conta do indicador do seguro. " & _
                                    "Número de conta corrente inválido (Banese - formato NNNNNNNN-N)."
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
                        If frmT46V110A.Caption = "Crédito em conta" Then
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

Public Sub gpHabSalvar()
    'Procedure: habilita opções salvar e desfazer.

    If Not gM46V999.gfDisponivel("frmT46V101A") Then Exit Sub    'Ficha 1000837 - Triagem Informar Beneficiario nos Itens

    If gCarga = False Then
        With frmT46V101A
        
            If .objstcPedido.objstcCtrlEmis.Cod_Iden_Stat_Acao = 0 Then
            .mnuEdiDesfazer.Enabled = True
            .barFerramentas.Buttons(12).Enabled = True 'Desfazer.
            End If
            
            
            .mnuArqImprimir.Enabled = False
            .mnuArqVisualizar.Enabled = False
            
            .mnuArqSalvar.Enabled = True
            .barFerramentas.Buttons(4).Enabled = True    'Salvar.
            
            .barFerramentas.Buttons(6).Enabled = False   'Imprimir.
            .barFerramentas.Buttons(7).Enabled = False   'Visualizar.
            .flxCobertura.Rows = 1
        End With
    End If
End Sub

Public Sub gpLimPremios()
    'Procedure: limpa informações de prêmios já calculados.

    '1. Se é carga do form, encerra processo.
    If gCarga = True Then
        Exit Sub
    End If

    '2. Fecha janelas de prêmios do item.
    If gM46V999.gfDisponivel("frmT46V123A") Then
        Unload frmT46V123A
    End If
    If gM46V999.gfDisponivel("frmT46V124A") Then
        Unload frmT46V124A
    End If

    '3. Habilita botão de comando e menu de salvar item.
    If gProChamado <> "T46V107A" Then
        Call gpHabSalvar
    End If
    Screen.MousePointer = vbDefault
End Sub

Public Sub gpMensagem(ByVal pMensagem As String)
    'Procedure: posiciona mensagem no FlexGrid.

    'Parâmetro de entrada...pMensagem...Mensagem.
    
    ''1. Se o form T46V136A ainda não estiver disponível, abrir e deixar a grade limpa.
    If gM46V999.gfDisponivel("frmT46V136A") = False Then
        Load frmT46V136A
        frmT46V136A.flxMensagens.Rows = 0
    End If

    ''2. Posiciona mensagem na grade.
    With frmT46V136A
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
Public Sub gpPos123_124()
    'Procedure: posiciona forms T46V123A e T46V124A na posição original.

    Dim lAltura         As Integer      'Altura original.
    Dim lLargura        As Integer      'Largura original.
    Dim lMarEsquerda    As Integer      'Margem esquerda.
    Dim lMarSuperior    As Integer      'Margem superior.

    '1. Calcula margens.
    lAltura = 5000
    lLargura = 3950
    With frmT46V101A
        lMarEsquerda = .Left + .Width - lLargura
        lMarSuperior = .Top + .panMulti(7).Top + 500
    End With

    '2. Form T46V123A.
    If gM46V999.gfDisponivel("frmT46V123A") Then
        With frmT46V123A
            .Move lMarEsquerda, lMarSuperior, lLargura, lAltura
            .tabPremios.Tab = 0
        End With
    End If

    '3. Form T46V124A.
    If gM46V999.gfDisponivel("frmT46V124A") Then
        With frmT46V124A
            .Move lMarEsquerda, lMarSuperior, lLargura, lAltura
            .tabResultado.Tab = 0
        End With
    End If
End Sub

Public Sub gpPosJuros(ByRef pAntJuros As Variant, ByRef pPosJuros As Variant, pCusParcelado As String)
    'Procedure: posiciona juros e indicador de custo de emissão parcelado.

    'Parâmetros de saída...pAntJuros.......Juros das parcelas com pagamento antecipado.
    '                      pPosJuros.......Juros das parcelas com pagamento postecipado.
    '                      pCusParcelado...Indicador de custo de emissão parcelado.

    Dim lAntJuros(12)   As Single       'Juros para parcelas com pagamento antecipado.
    Dim lI              As Byte         'Utilizado com For...Next.
    Dim lPosJuros(12)   As Single       'Juros para parcelas com pagamento postecipado.


    '1. Posiciona juros.
    For lI = 1 To gM46V999.gProQuaMaxParcelas '<- originalmente 10
        lAntJuros(lI) = 0
        lPosJuros(lI) = 0
    Next lI
    
    With frmT46V142A
        If gM46V999.gfPreenchido(.txtAntJuro(1).Text) Then
            lAntJuros(1) = CSng(.txtAntJuro(1).Text)
        End If
        If gM46V999.gfPreenchido(.txtAntJuro(2).Text) Then
            lAntJuros(2) = CSng(.txtAntJuro(2).Text)
        End If
        If gM46V999.gfPreenchido(.txtAntJuro(3).Text) Then
            lAntJuros(3) = CSng(.txtAntJuro(3).Text)
        End If
        If gM46V999.gfPreenchido(.txtAntJuro(4).Text) Then
            lAntJuros(4) = CSng(.txtAntJuro(4).Text)
        End If
        If gM46V999.gfPreenchido(.txtAntJuro(5).Text) Then
            lAntJuros(5) = CSng(.txtAntJuro(5).Text)
        End If
        If gM46V999.gfPreenchido(.txtAntJuro(6).Text) Then
            lAntJuros(6) = CSng(.txtAntJuro(6).Text)
        End If
        If gM46V999.gfPreenchido(.txtAntJuro(7).Text) Then
            lAntJuros(7) = CSng(.txtAntJuro(7).Text)
        End If
        If gM46V999.gfPreenchido(.txtAntJuro(8).Text) Then
            lAntJuros(8) = CSng(.txtAntJuro(8).Text)
        End If
        If gM46V999.gfPreenchido(.txtAntJuro(9).Text) Then
            lAntJuros(9) = CSng(.txtAntJuro(9).Text)
        End If
        If gM46V999.gfPreenchido(.txtAntJuro(10).Text) Then
            lAntJuros(10) = CSng(.txtAntJuro(10).Text)
        End If
        If gM46V999.gfPreenchido(.txtAntJuro(11).Text) Then
            lAntJuros(11) = CSng(.txtAntJuro(11).Text)
        End If
        'If gM46V999.gfPreenchido(.txtAntJuro(12).Text) Then
            'lAntJuros(12) = CSng(.txtAntJuro(12).Text)
            lAntJuros(12) = 0
        'End If
        If gM46V999.gfPreenchido(.txtPosJuro(1).Text) Then
            lPosJuros(1) = CSng(.txtPosJuro(1).Text)
        End If
        If gM46V999.gfPreenchido(.txtPosJuro(2).Text) Then
            lPosJuros(2) = CSng(.txtPosJuro(2).Text)
        End If
        If gM46V999.gfPreenchido(.txtPosJuro(3).Text) Then
            lPosJuros(3) = CSng(.txtPosJuro(3).Text)
        End If
        If gM46V999.gfPreenchido(.txtPosJuro(4).Text) Then
            lPosJuros(4) = CSng(.txtPosJuro(4).Text)
        End If
        If gM46V999.gfPreenchido(.txtPosJuro(5).Text) Then
            lPosJuros(5) = CSng(.txtPosJuro(5).Text)
        End If
        If gM46V999.gfPreenchido(.txtPosJuro(6).Text) Then
            lPosJuros(6) = CSng(.txtPosJuro(6).Text)
        End If
        If gM46V999.gfPreenchido(.txtPosJuro(7).Text) Then
            lPosJuros(7) = CSng(.txtPosJuro(7).Text)
        End If
        If gM46V999.gfPreenchido(.txtPosJuro(8).Text) Then
            lPosJuros(8) = CSng(.txtPosJuro(8).Text)
        End If
        If gM46V999.gfPreenchido(.txtPosJuro(9).Text) Then
            lPosJuros(9) = CSng(.txtPosJuro(9).Text)
        End If
        If gM46V999.gfPreenchido(.txtPosJuro(10).Text) Then
            lPosJuros(10) = CSng(.txtPosJuro(10).Text)
        End If
        If gM46V999.gfPreenchido(.txtPosJuro(11).Text) Then
            lPosJuros(11) = CSng(.txtPosJuro(11).Text)
        End If
        'If gM46V999.gfPreenchido(.txtPosJuro(12).Text) Then
            'lPosJuros(12) = CSng(.txtPosJuro(12).Text)
            lPosJuros(12) = 0
        'End If
    End With
    pAntJuros = Array(lAntJuros(1), lAntJuros(2), lAntJuros(3), lAntJuros(4), lAntJuros(5), _
                      lAntJuros(6), lAntJuros(7), lAntJuros(8), lAntJuros(9), lAntJuros(10), _
                      lAntJuros(11), lAntJuros(12))
    pPosJuros = Array(lPosJuros(1), lPosJuros(2), lPosJuros(3), lPosJuros(4), lPosJuros(5), _
                      lPosJuros(6), lPosJuros(7), lPosJuros(8), lPosJuros(9), lPosJuros(10), _
                      lPosJuros(11), lPosJuros(12))

    '2. Posiciona indicador de custo de emissão parcelado.
    If gM46V999.gfDisponivel("frmT46V120A") Then
        pCusParcelado = IIf(frmT46V120A.optCusEmissao(0) = True, "S", "N")
    Else
        pCusParcelado = "S"
    End If
End Sub


Public Sub gpVerificarVencimento(ByVal pQuantParcelas As Integer, ByVal pTipoPagamento As Integer, _
                                ByVal pPriVenc As Long, ByVal pFormaPagamento As Integer, _
                                ByVal pDataIniVig As Long, ByRef pDataVenc As Variant, _
                                ByRef pVencIrregular As Boolean)
        'Procedure : Verifica data de vencimento das parcelas.
        
        'Parâmetro de entrada...pQuantParcelas..........As Integer
        '                       pTipoPagamento..........As Integer
        '                       pPriVenc................As Long
        '                       pFormaPagamento.........As integer
        '                       pDataIniVig.............As Long
        'Parâmetro de saída.....pDataVenc...............As Variant
        '                       pVencIrregular..........As Boolean

Dim lI              As Integer
Dim lCount          As Integer
Dim lSomaDia        As Integer
Dim lDatVenc        As Long
Dim lDatAux         As Long
Dim lAuxiliar       As Long
Dim lArrayData      As Variant
Dim lDataIrreg      As Boolean

    If pTipoPagamento = 1 Then            'Antecipado
        lSomaDia = 25
    Else
        lSomaDia = 8
    End If
    lDataIrreg = False
    'calculo a primeira data de vencimento
    If Date > CDate(gM46V999.gfForData(2, pDataIniVig)) Then
        If pPriVenc = 0 Then
            lDatVenc = gM46V999.gfForData(1, Format(CDate(Date + lSomaDia), "dd/mm/yyyy"))
            lAuxiliar = 1
            Do While lAuxiliar = 1
                If Not IsDate(gM46V999.gfForData(2, lDatVenc)) Then
                    lDatVenc = lDatVenc - 1
                    lDataIrreg = True
                Else
                     lAuxiliar = 0
                End If
            Loop
        Else
            lDatVenc = pPriVenc
        End If
    Else
        If pPriVenc = 0 Then
            lDatVenc = gM46V999.gfForData(1, Format(CDate(Date + lSomaDia), "dd/mm/yyyy"))
            lAuxiliar = 1
            Do While lAuxiliar = 1
                If Not IsDate(gM46V999.gfForData(2, lDatVenc)) Then
                    lDatVenc = lDatVenc - 1
                    lDataIrreg = True
                Else
                     lAuxiliar = 0
                End If
            Loop
        Else
            lDatVenc = pPriVenc
        End If
    End If
    
    'Verifico todas as demais datas
    ReDim lArrayData(pQuantParcelas)
    'Grava a primeira data de vencimento
    If pTipoPagamento = 2 Then 'Não é antecipado
        lArrayData(0) = lDatVenc
        lCount = 1
    Else 'Antecipado

        lArrayData(0) = gM46V999.gfForData(1, Format(Date + 8, "dd/mm/yyyy"))
        If pQuantParcelas > 0 Then
            lArrayData(1) = lDatVenc
        End If
        lCount = 2
    End If
    pPriVenc = Right(lDatVenc, 2)
    For lI = lCount To pQuantParcelas - 1
        lDatVenc = Left(lDatVenc, 6) & Format(pPriVenc, "00")
        If Mid(lDatVenc, 5, 2) = 12 Then
            lDatVenc = Mid(lDatVenc, 1, 4) + 1 & "01" & Mid(lDatVenc, 7, 2)
        Else
            lDatVenc = Mid(lDatVenc, 1, 6) + 1 & Mid(lDatVenc, 7, 2)
        End If
        lAuxiliar = 1
        Do While lAuxiliar = 1
            If Not IsDate(gM46V999.gfForData(2, lDatVenc)) Then
                lDatVenc = lDatVenc - 1
                lDataIrreg = True
            Else
                 lAuxiliar = 0
            End If
        Loop
        lArrayData(lI) = lDatVenc
        
    Next
    pVencIrregular = lDataIrreg
    pDataVenc = lArrayData
End Sub

Private Function mfCalcData(pData As Long, pDias As Integer) As Long
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
    
        mfCalcData = Format(CDate(lAuxPreData) + pDias, "yyyymmdd")
    End If
    
End Function

Public Function gfExportarParaBaseTemp(pNosso_Numero As String, pNum_Ped As Long, _
                    pNumBloq As Long, pMensagem As String) As Boolean
                    
Dim bdBasDados  As ADODB.Connection 'Base de dados.
Dim bdTemp      As ADODB.Connection 'Base temporária.
Dim flTabela    As clsYasField       'Campos da tabela.
Dim rsTabela    As clsYasRecordSet  'Registro da tabela.
Dim lCount      As Long             'Utilizado no for...next.
Dim lIdeTabela  As String           'Nome da tabela.
Dim lInsert     As String           'Insert
Dim lContinuar  As Boolean          'Indicador se continua o procedimento ou não.
Dim lBanco      As Integer          'Banco
Dim lPropTabela As String
Dim lSelect     As String
Dim lDelete     As String

    gfExportarParaBaseTemp = False
    On Error GoTo ControleErro
    Call gM46V999.gpGraLog(0, "Multi:" & gM46V999.gEndP00MULTI)
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
                Call gM46V999.gpGraLog(0, "Multi 1:" & pMensagem)
                Call gM46V999.gpFecha2(bdTemp)
                Exit Function
            End If
        End If
        lDelete = " DELETE FROM  " & lIdeTabela
        If gM46V999.gfExeSQL(bdTemp, lDelete, pMensagem) <> 0 Then
            Call gM46V999.gpGraLog(0, "Multi2 :" & pMensagem)
            Call gM46V999.gpFecha2(bdTemp)
            Exit Function
        End If
        If lIdeTabela <> "TAB_PED_COBERT" Then
            lSelect = "SELECT * FROM " & lPropTabela & lIdeTabela & " WHERE Nosso_Numero = '" & pNosso_Numero & "'"
        Else
            lSelect = "SELECT Tab_Ped_Cobert.*, "
            lSelect = lSelect & " Tab_Ped_Cobert_Compl.Alterar_Cobert"
            lSelect = lSelect & " From "
            lSelect = lSelect & lPropTabela & "Tab_Ped_Cobert as Tab_Ped_Cobert INNER JOIN     " & lPropTabela & "Tab_Ped_Cobert_Compl as Tab_Ped_Cobert_Compl ON"
            lSelect = lSelect & " (Tab_Ped_Cobert.COD_COBERT   = Tab_Ped_Cobert_Compl.COD_COBERT) AND"
            lSelect = lSelect & " (Tab_Ped_Cobert.NUM_ITEM     = Tab_Ped_Cobert_Compl.NUM_ITEM) AND"
            lSelect = lSelect & " (Tab_Ped_Cobert.NOSSO_NUMERO = Tab_Ped_Cobert_Compl.NOSSO_NUMERO)"
            lSelect = lSelect & " Where Tab_Ped_Cobert.Nosso_Numero = '" & pNosso_Numero & "'"
        End If

        
        If gM46V999.gfObtRegistro(bdBasDados, lSelect, rsTabela, pMensagem) = False Then
            Call gM46V999.gpGraLog(0, "Multi 3 :" & pMensagem)
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
                             "IND_SIN_INSPEC", "COD_DIZERITEM", "AJUSTEENDOSSO"
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
                             "IND_SIN_INSPEC", "COD_DIZERITEM", "AJUSTEENDOSSO"

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
                    Call gM46V999.gpGraLog(0, "Multi 4:" & pMensagem)
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
Public Function gfExportarDadosApolice(pNum_Apol As String, _
                    pMensagem As String) As Boolean
                    
Dim bdBasDados  As ADODB.Connection 'Base de dados.
Dim bdTemp      As ADODB.Connection 'Base temporária.
Dim flTabela    As clsYasField      'Campos da tabela.
Dim rsTabela    As clsYasRecordSet  'Registro da tabela.
Dim lCount      As Long             'Utilizado no for...next.
Dim lIdeTabela  As String           'Nome da tabela.
Dim lParTabela  As String           'Nome da tabela de destino.
Dim lInsert     As String           'Insert
Dim lContinuar  As Boolean          'Indicador se continua o procedimento ou não.
Dim lBanco      As Integer          'Banco
Dim lSelect     As String

    gfExportarDadosApolice = False
    On Error GoTo ControleErro
    
    '1. Obtém registros da tabela.
    'If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gAppPath & "P00Multi.mdb", bdTemp, _
    '        pMensagem) = False Then Exit Function
    
    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, bdTemp, _
            pMensagem) = False Then Exit Function
    
    If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, bdBasDados, _
            pMensagem) = False Then Exit Function
    
    For lCount = 1 To 7
        lContinuar = True
        Select Case lCount
            Case 1
                If gM46V999.gBaseLocal = True Then
                    lIdeTabela = "TAB_APOL"
                Else
                    lIdeTabela = "GED.DBO.TAB_PED"
                End If
                lParTabela = "TAB_PED"
            Case 2
                If gM46V999.gBaseLocal = True Then
                    lIdeTabela = "TAB_APOL_CORR"
                Else
                    lIdeTabela = "GED.DBO.TAB_PED_CORR"
                End If
                lParTabela = "TAB_PED_CORR"
            Case 3
                If gM46V999.gBaseLocal = True Then
                    lIdeTabela = "TAB_APOL_COSS"
                Else
                    lIdeTabela = "GED.DBO.TAB_PED_COSS"
                End If
                lParTabela = "TAB_PED_COSS"
            Case 4
                If gM46V999.gBaseLocal = True Then
                    lIdeTabela = "TAB_APOL_ENDER"
                Else
                    lIdeTabela = "GED.DBO.TAB_PED_ENDER"
                End If
                lParTabela = "TAB_PED_ENDER"
            Case 5
                If gM46V999.gBaseLocal = True Then
                    lIdeTabela = "TAB_APOL_LOC"
                Else
                    lIdeTabela = "GED.DBO.TAB_PED_LOC"
                End If
                lParTabela = "TAB_PED_LOC"
            Case 6
                If gM46V999.gBaseLocal = True Then
                    lIdeTabela = "TAB_APOL_CLAU"
                Else
                    lIdeTabela = "GED.DBO.TAB_PED_CLAU"
                End If
                lParTabela = "TAB_PED_CLAU"
            Case 7
                If gM46V999.gBaseLocal = True Then
                    lIdeTabela = "TAB_APOL_COBERT"
                Else
                    lIdeTabela = "GED.DBO.TAB_PED_COBERT"
                End If
                lParTabela = "TAB_PED_COBERT"
        End Select
        lSelect = "SELECT * FROM " & lIdeTabela & " WHERE Num_Apol = " & pNum_Apol
        
        If gM46V999.gfObtRegistro(bdBasDados, lSelect, rsTabela, pMensagem) = False Then
            Call gM46V999.gpFecha2(bdTemp)
            Call gM46V999.gpFecha2(bdBasDados)
            Exit Function
        End If
        If rsTabela.EOF = True Then
            lContinuar = False
            If UCase(lIdeTabela) = "TAB_APOL" Or UCase(lIdeTabela) = "TAB_PED" Then
                pMensagem = "FNL0003 - " & gM46V999.gPrefixo & "Apólice não localizada (" & pNum_Apol & _
                                 " - exportando para base de cálculo.)."
                Exit Function
            End If
        End If
        If lContinuar = True Then
            '2. Grava nova cotação/proposta.
            While Not rsTabela.EOF
                lInsert = "INSERT INTO " & lParTabela & " ("
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
                             "PER_DESC_LMI", "LMI_UNICO", "NUM_COTAC_LMI", _
                             "COD_REG_CONS", "NUM_REG_CONS", "IND_CONDGERAIS", "IND_REENTRADA", _
                             "NOSSO_NUMERO_ORIGINAL", "IND_SIN_INSPEC", "AJUSTEENDOSSO"
                        Case "NUM_PI"
                        Case "COD_PRODUTO"
                            If lIdeTabela = "GED.DBO.TAB_PED_LOC" Or lIdeTabela = "TAB_APOL_LOC" Then
                                lInsert = lInsert & flTabela.Name & ", "
                            End If
                        Case Else
                            'Campos que não serão exportados.
                            lInsert = lInsert & flTabela.Name & ", "
                    End Select
                Next flTabela
                lInsert = Left$(lInsert, Len(lInsert) - 2) & ") VALUES ("
                For Each flTabela In rsTabela.Fields
                    Select Case UCase(flTabela.Name)
                    'Ficha 1000307
                        Case "NOSSO_NUMERO", "COD_UNICIDADE", "ULT_TARIFA", _
                             "DAT_PRI_CALCULO", "SEXO", "NOM_EMPRESA", _
                             "DAT_EMIS_DOC", "ORG_EXP", "NUM_DOC", _
                             "COD_DOC", "COD_RAMO_ATIV", "TIP_PESSOA", _
                             "COD_SEQUENCIA", "COD_PROTOCOLO", "TIP_EMISSAO", _
                             "DESCRICAO", "NUM_APOL_ANTERIOR", "NUM_ITEM_ANTERIOR", _
                             "IND_SINISTRO", "DIG_BLOQUETO", "COD_DIG_AGEN_BAS_COR", _
                             "ALTERAR_COBERT", "COD_VER_UNICI", "DIG_COD_SEGURADO", _
                             "PER_DESC_LMI", "LMI_UNICO", "NUM_COTAC_LMI", _
                             "COD_REG_CONS", "NUM_REG_CONS", "IND_CONDGERAIS", "IND_REENTRADA", _
                             "NOSSO_NUMERO_ORIGINAL", "IND_SIN_INSPEC", "COD_DIZERITEM", "AJUSTEENDOSSO"
                        Case "COD_PRODUTO"
                            If lIdeTabela = "GED.DBO.TAB_PED_LOC" Or lIdeTabela = "TAB_APOL_LOC" Then
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case "NUM_PI"
                        Case "NUM_APOL_AN"
                            lInsert = lInsert & "0, "
                        'Case "NUM_ENDO"
                        '    lInsert = lInsert & "0, "
                        Case Else
                            'Campos que não serão exportados.
                            Select Case flTabela.Type_
                                Case adBigInt, adBinary, adInteger, adLongVarBinary, adSmallInt, _
                                     adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, _
                                     adUnsignedTinyInt                          'Numérico - inteiro.
                                    If IsNull(flTabela.Value) Then
                                        lInsert = lInsert & "0, "
                                    Else
                                        lInsert = lInsert & flTabela.Value & ", "
                                    End If
                                Case adCurrency, adSingle, adDouble, adNumeric             'Numérico - com vírgula.
                                    If IsNull(flTabela.Value) Then
                                        lInsert = lInsert & "0, "
                                    Else
                                        lInsert = lInsert & gM46V999.gfForValor(1, flTabela.Value) & ", "
                                    End If
                                Case adChar, adVarChar, adLongVarChar                            'Texto.
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
    gfExportarDadosApolice = True
    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "gfExportarDadosApolice ERRO - " & Err & " " & Err.Description)
End Function






