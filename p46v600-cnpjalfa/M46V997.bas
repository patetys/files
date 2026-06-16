Attribute VB_Name = "basM46V997"
''M�DULO COM FUN��ES E PROCEDURES COMUNS DO SYAS.

'AO ALTERAR ESTE M�DULO, COMPILAR E CRIAR EXECUT�VEL DE TODOS OS M�DULOS QUE O REFERENCIAM.

Option Explicit
Public Enum E_SelGotFocus
    E_SelInicio = 1
    E_SelFim = 2
    E_SelTudo = 3
End Enum

Public gPerfilEspecial As Long        '1 - Especial
'2 - N�o especial

Public Function gfPreencheBranco(pTexto As String, pTamanho As Integer)
    If pTamanho > Len(pTexto) Then
        gfPreencheBranco = Space(pTamanho - Len(pTexto)) & pTexto
    Else
        gfPreencheBranco = pTexto
    End If
End Function

'**********************
Public Function gfComBasDados_DAO(ByVal pEndBasDados As String) As Boolean
'Fun��o: Compactar base de dados.

'Par�metro de entrada...pEndBasDados...Endere�o da base de dados.

    Dim lCaracter As String    'Caracter identificador do endere�o.
    Dim lNomBasDados As String    'Nome da base de dados.
    Dim lNovBasDados As String    'Nova base de dados.
    Dim lPosicao1 As Byte    'Posi��o 1 - utilizado com InStr.
    Dim lPosicao2 As Byte    'Posi��o 2 - utilizado com InStr.
    Dim lSegBasDados As String    'Nome de seguran�a da base de dados.

    gfComBasDados_DAO = False

    '1. Posiciona endere�o da base de dados.
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

    '2. Se j� existe arquivo BD1.MDB, exclu�-lo.
    If gM46V999.gfPreenchido(Dir(lNovBasDados)) Then
        Kill lNovBasDados
    End If

    '3. Compacta arquivo.
    DBEngine.CompactDatabase pEndBasDados, lNovBasDados, , , ";PWD=" & LCase(lNomBasDados) & "46"

    '4. Renomeia arquivo atual para seguran�a (salva arquivo para efeito de seguran�a).
    lSegBasDados = Left$(pEndBasDados, lPosicao1) & "XXXXXXXX.MDB"
    Name pEndBasDados As lSegBasDados

    '5. Renomeia arquivo compactado.
    Name lNovBasDados As pEndBasDados

    '6. Exclui arquivo antigo.
    Kill lSegBasDados

    gfComBasDados_DAO = True
End Function
Public Function gfBuscaUsuarioTransmissao(ByRef pCodUsuTransmissao As String, ByRef pMensagem As String) _
       As Boolean
'Fun��o: busca usu�rio de transmiss�o.

'Par�metros de sa�da...pCodUsuTransmissao...C�digo do usu�rio.
'                      pMensagem............Mensagem.

    Dim lCodUsuario As Long    'C�digo do usu�rio.
    Dim lEndSyasEdi_EXE As String    'Endere�o do arquivo SYASEDI.EXE.
    Dim lNumSyasEdi_EXE As Integer    'Utilizado com FreeFile.
    Dim lRegistro As String    'Registro.

    gfBuscaUsuarioTransmissao = False

    '1. Verifica arquivo SYASEDI.EXE.
    lEndSyasEdi_EXE = gM46V999.gAppPath & "SYASEDI.EXE"
    If Not gM46V999.gfPreenchido(Dir(lEndSyasEdi_EXE)) Then
        'Procuro pelo arquivo antigo
        lEndSyasEdi_EXE = gM46V999.gAppPath & "Eapscript.INI"
        If Not gM46V999.gfPreenchido(Dir(lEndSyasEdi_EXE)) Then
            pMensagem = "FTExxxx - Erro ao obter c�digo do usu�rio de transmiss�o (identifica��o n�o " & _
                        "localizada)." & vbLf & vbLf & "Reinicie o processo de c�lculo."
            Exit Function
        End If
    End If
    lNumSyasEdi_EXE = FreeFile
    On Error Resume Next
    Open lEndSyasEdi_EXE For Input Access Read As #lNumSyasEdi_EXE
    If Err <> 0 Then
        pMensagem = "FTExxxx - Erro ao obter c�digo do usu�rio de transmiss�o (" & Err & " - " & _
                    Error & ")." & vbLf & vbLf & "Reinicie o processo de c�lculo."
        Exit Function
    End If
    On Error GoTo 0

    '2. Obt�m c�digo do usu�rio.
    Do While Not EOF(lNumSyasEdi_EXE)
        Line Input #lNumSyasEdi_EXE, lRegistro
        If Mid(lRegistro, 1, 7) = "sender=" Then
            lCodUsuario = Val(Mid(lRegistro, 9, 6))
        End If
    Loop
    Close #lNumSyasEdi_EXE
    pCodUsuTransmissao = lCodUsuario

    gfBuscaUsuarioTransmissao = True
End Function

Public Function gfRepBasDados_DAO(ByVal pArquivo As String) As Boolean
'Fun��o: Repara arquivo.

    gfRepBasDados_DAO = False

    On Error Resume Next
    DBEngine.RepairDatabase pArquivo
    If Err <> 0 Then
        Call gM46V999.gpGraLog(2, "FTE0024 - Erro ao reparar arquivo " & pArquivo & ": " & Err & " - " & Error & _
                                  ".")
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    gfRepBasDados_DAO = True
End Function

Public Function ConsisteKeyNum(Tecla As Integer, Optional sExcecao As String) As Integer
    If (Tecla < 48 Or Tecla > 57) And Tecla <> 9 And Tecla <> 8 And Tecla <> 13 And Tecla <> 27 And Tecla <> 22 Then
        If InStr(sExcecao, Chr(Tecla)) <> 0 And Len(Trim(sExcecao)) <> 0 Then
            ConsisteKeyNum = Tecla
        Else
            ConsisteKeyNum = 0
        End If
    Else
        ConsisteKeyNum = Tecla
    End If
    If Tecla = vbKeyReturn Then
        ConsisteKeyNum = 0
        SendKeys "{TAB}"
    End If
End Function

Public Function gpGeraLogCalcInspecao(pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B, Optional pTransmissao As Boolean = False, Optional pDiretorio As String = "") As Boolean
    Dim lobjstcCobert As stcA46V704B    'Objeto da classe de coberturas
    Dim lCgcCpf055 As String

'Procedure utilizada para gerar um log contendo informa��es, para serem enviadas para yasuda
'em que o servidor da yasuda armazena estas informa��es na tabela [RAMOSDIVERSOS].[dbo].[Tab_Trans_Insp]

'Parametros de entrada com base na tabela
'*******  LAYOUT DA TABELA NO SERVIDOR YASUDA
'    [VERSAO_ARQUIVO] [varchar] (2)
'    [NOSSO_NUMERO] [varchar] (20)
'    [NUM_PED] [varchar] (6)
'    [NUM_ITEM] [varchar] (5)
'    [NUM_APOL_ANT] [varchar] (10)
'    [NOM_SEGURADO] [varchar] (60)
'    [NUM_TEL] [varchar] (14)
'    [NUM_CGC_CPF] [varchar] (14)
'    [COD_RAMO] [varchar] (3)
'    [COD_ORG] [varchar] (2)
'    [DAT_INI_VIG] [varchar] (8)
'    [DAT_TER_VIG] [varchar] (8)
'    [COD_TIP_LOGR_RISCO] [varchar] (3)
'    [NOM_LOGR_RISCO] [varchar] (50)
'    [DSC_COMPL_RISCO] [varchar] (20)
'    [NOM_BAIRRO_RISCO] [varchar] (30)
'    [NOM_CID_RISCO] [varchar] (30)
'    [SIG_UF_RISCO] [varchar] (2)
'    [NUM_CEP_RISCO] [varchar] (8)
'    [NUM_RISCO] [varchar] (8)
'    [DAT_TRANS] [varchar] (8)
'    [COD_IDEN_PESSOA] [varchar] (1)
'    [COD_UNID_PROD] [varchar] (4)
'    [COD_PROD] [varchar] (4)
'    [COD_CORR] [varchar] (7)
'    [COD_CLASF] [varchar] (4)
'    [COD_COMPL_CLASF] [varchar] (2)
'    [PESS_CONT] [varchar] (40)
'    [COD_PRODUTO] [varchar] (4)
'    [FLAG_TRANS] [varchar] (1)
'    [COD_COBERT] [varchar] (4)
'    [VAL_IS_COBERT] [varchar] (20)

'******
    Dim lCaminho As String    'Caminho onde os dados s�o salvos
    Dim lEndArq055 As String    'Endereco do arquivo LOG
    Dim lNumArq055 As String    'Numero ref. para abertura do arquivo obtido com "FreeFile"
    Dim lLCoduser As String    'Codigo do usuario de transmiss�o Numerico
    Dim lCodUsuario As String    'Codigo do usuario de transmiss�o ( formatado 6 digitos)
    Dim lMsg As String    'Mensagem
    Dim lCont As Long
    Dim lCod_Seq As Long    'Sequ�ncia da transmiss�o
    Dim lImprime As String

    Dim lCod_Cobert(250) As Integer
    Dim lIS_Cobert(250) As Long

    Dim lSql As String

    Err = 0

    gpGeraLogCalcInspecao = False

    On Error GoTo gpGeraLogCalcInspecao_Erro

    '1- Posiciona o caminho dos dados
    If pTransmissao = True Then
        lCaminho = pDiretorio
        '2-Verifica se existencia do log ref. ao mesmo calculo
        lEndArq055 = lCaminho & "*" & pobjstcPed.NOSSO_NUMERO & "_" & pobjstcPedLoc.Num_Item & ".055"
    Else
        If gM46V999.gfPreenchido(gM46V999.gEndSYASUpDt) = True Then
            'Caminho ods arquivos de atualiza��o do SYAS
            'Informado no arquivo SYASPORD.INI
            lCaminho = gM46V999.gEndSYASUpDt
        Else
            'Endere�o das bases de dados tarifarias
            lCaminho = gM46V999.gEndServidor
        End If
        '2-Verifica se existencia do log ref. ao mesmo calculo
        If Not gM46V999.gfPreenchido(Dir(lCaminho & "\REMESSA", vbDirectory)) Then
            On Error Resume Next
            MkDir (lCaminho & "\REMESSA")
            If Err <> 0 Then
                Call gM46V999.gpGraLog(2, "FTE0015 - Erro ao criar pasta " & lCaminho & "\REMESSA" & _
                                        " (pasta de remessa para arquivos de inspe��o - " & Err & " - " & _
                                          Error & ")." & vbLf & vbLf & "O sistema ser� encerrado.")
                Exit Function
            End If
            On Error GoTo gpGeraLogCalcInspecao_Erro
        End If

        lEndArq055 = lCaminho & "REMESSA\" & "*" & pobjstcPed.NOSSO_NUMERO & "_" & pobjstcPedLoc.Num_Item & ".055"
    End If



    Call gM46V999.gpGraLog(0, "Verifica se existencia do log ref. ao mesmo calculo " & lEndArq055)

    If gM46V999.gfPreenchido(Dir(lEndArq055)) = True Then
        'exclui o ultimo arquivo se existir
        Kill lEndArq055

        Call gM46V999.gpGraLog(0, "exclui o ultimo arquivo se existir " & lEndArq055)

    End If


    '3-Verificar se foi agM46V999.gEndada inspe��o
    If pobjstcPedLoc.Insp_Agend = False Then Exit Function

    '4-Carrega as coberturas

    lCont = 0
    For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
        lCont = lCont + 1
        If lCont > 250 Then Exit For
        lCod_Cobert(lCont) = lobjstcCobert.Cod_Cobert
        lIS_Cobert(lCont) = Format(lobjstcCobert.Val_IS, "standard")
    Next lobjstcCobert

    '5-Obtem o codigo do usuario de transmiss�o
    If pTransmissao = True Then
        lLCoduser = Mid(pobjstcPed.NOSSO_NUMERO, 14, 6)
    Else
        If gfBuscaUsuarioTransmissao(lLCoduser, lMsg) = False Then
            'Sai do procedimento
            'N�o obteve codigo de usu�rio
            Call gM46V999.gpGraLog(0, lMsg)
            Exit Function
        End If
    End If
    lCodUsuario = Format(lLCoduser, "000000")

    '5-Inicia a cria��o do arquivo

    lNumArq055 = FreeFile
    If pTransmissao = True Then
        lEndArq055 = lCaminho & lCodUsuario & "_" & pobjstcPed.NOSSO_NUMERO & "_" & pobjstcPedLoc.Num_Item & ".055"
    Else
        lEndArq055 = lCaminho & "REMESSA\" & lCodUsuario & "_" & pobjstcPed.NOSSO_NUMERO & "_" & pobjstcPedLoc.Num_Item & ".055"
    End If

    Call gM46V999.gpGraLog(0, "Inicia a cria��o do arquivo " & lEndArq055)

    Open lEndArq055 For Output Access Write As #lNumArq055

    If IsNumeric(pobjstcPed.Num_Cgc_Cpf) Then
        lCgcCpf055 = Format(pobjstcPed.Num_Cgc_Cpf, "00000000000000")
    Else
        lCgcCpf055 = Right$(String(14, "0") & Trim(pobjstcPed.Num_Cgc_Cpf), 14)
    End If

    lImprime = "V2;" & pobjstcPedLoc.NOSSO_NUMERO & ";" & Format(pobjstcPedLoc.Num_Ped, "000000") & _
               ";" & Format(pobjstcPedLoc.Num_Item, "00000") & ";" & Format(pobjstcPedLoc.Num_Apol_Anterior, "0000000000") & _
               ";" & pobjstcPed.Nom_Segurado & String(60 - Len(Left(pobjstcPed.Nom_Segurado, 60)), " ") & _
               ";" & pobjstcPedLoc.objstcPedInsp.Num_Tel & String(14 - Len(Left(pobjstcPedLoc.objstcPedInsp.Num_Tel, 14)), " ") & _
               ";" & lCgcCpf055 & ";" & Format(pobjstcPed.Cod_Ramo, "000") & ";" & Format(pobjstcPed.Cod_Org, "00") & _
               ";" & pobjstcPed.Dat_Ini_Vig & ";" & pobjstcPed.Dat_Ter_Vig & ";" & pobjstcPedLoc.Cod_Tip_Logr_Risco & String(3 - Len(Left(pobjstcPedLoc.Cod_Tip_Logr_Risco, 3)), " ") & _
               ";" & pobjstcPedLoc.Nom_Logr_Risco & String(50 - Len(Left(pobjstcPedLoc.Nom_Logr_Risco, 50)), " ") & ";" & pobjstcPedLoc.Dsc_Compl_Risco & String(20 - Len(Left(pobjstcPedLoc.Dsc_Compl_Risco, 20)), " ") & _
               ";" & pobjstcPedLoc.Nom_Bairro_Risco & String(30 - Len(Left(pobjstcPedLoc.Nom_Bairro_Risco, 30)), " ") & _
               ";" & pobjstcPedLoc.Nom_Cid_Risco & String(30 - Len(Left(pobjstcPedLoc.Nom_Cid_Risco, 30)), " ") & ";" & pobjstcPedLoc.Sig_UF_Risco & ";" & Format(pobjstcPedLoc.Num_Cep_Risco, "00000000") & _
               ";" & Format(pobjstcPedLoc.objstcPedInsp.Num_Logr_Risco, "00000000") & ";" & Format$(Now, "yyyymmdd") & _
               ";" & pobjstcPed.Tip_Pessoa & ";" & Format(pobjstcPed.Cod_Unid_Prod, "0000") & _
               ";" & Format(pobjstcPed.Cod_Prod, "0000") & ";" & Format(pobjstcPed.objstcCtrlEmis.Cod_Corr, "0000000") & _
               ";" & pobjstcPedLoc.Cod_Clasf & String(4 - Len(Left(pobjstcPedLoc.Cod_Clasf, 4)), " ") & ";" & pobjstcPedLoc.Cod_Compl_Clasf & _
               ";" & pobjstcPedLoc.objstcPedInsp.Pess_Cont & String(40 - Len(Left(pobjstcPedLoc.objstcPedInsp.Pess_Cont, 40)), " ") & _
               ";" & Format(pobjstcPedLoc.Cod_Produto, "0000") & ";0"


    lCont = 0
    For lCont = 1 To 250
        If lCod_Cobert(lCont) = 0 Then Exit For
        lImprime = lImprime & ";" & Format(lCod_Cobert(lCont), "0000") & ";" & Format(lIS_Cobert(lCont), "00000000000000000.00")
    Next lCont

    Print #lNumArq055, lImprime
    Close #lNumArq055

    Call gM46V999.gpGraLog(0, "Final da cria��o do arquivo " & lEndArq055)

    gpGeraLogCalcInspecao = True

gpGeraLogCalcInspecao_Saida:
    On Error GoTo 0
    Exit Function
gpGeraLogCalcInspecao_Erro:
    Call gM46V999.gpGraLog(0, "gpGeraLogCalcInspecao ERRO " & Err & " " & Error$)
    GoTo gpGeraLogCalcInspecao_Saida
End Function
Public Function gfCompactar(ByVal pEndBasDados As String) As Boolean
'Fun��o: compactar base de dados.

'Par�metro de entrada...pEndBasDados...Endere�o da base de dados.

    gfCompactar = False

    '1. Verifica endere�o da base de dados.
    If Not gM46V999.gfPreenchido(pEndBasDados) Then
        Call gM46V999.gpGraLog(2, "FTE0063 - Endere�o da base de dados n�o preenchido (processo de compactar " & _
                                  "base de dados).")
        End
    End If

    '2. Inicia processo para reparar e compactar base de dados.
    Call gM46V999.gpGraLog(0, "AGE0006 - Compacta base de dados " & pEndBasDados & _
                            " (in�cio - data da �ltima atualiza��o: " & FileDateTime(pEndBasDados) & _
                            " - tamanho: " & FileLen(pEndBasDados) & ").")

    '3. Repara base de dados.
    If gfRepBasDados_DAO(pEndBasDados) = False Then
        Exit Function
    End If

    '4. Compacta base de dados.
    If gfComBasDados_DAO(pEndBasDados) = False Then
        Exit Function
    End If

    '5. Encerra processo.
    Call gM46V999.gpGraLog(0, "AGE0006 - Compacta base de dados " & pEndBasDados & _
                            " (t�rmino - data da �ltima atualiza��o: " & FileDateTime(pEndBasDados) & _
                            " - tamanho: " & FileLen(pEndBasDados) & ").")

    gfCompactar = True
End Function

Function DesabilitaControles(FrmDisable As Form, Optional ByVal TagPermite As String, Optional pStatAcao As Integer)
    Dim i As Integer

    '[FRH.07]->
    Dim lEstadoEmissao As Boolean
    lEstadoEmissao = True

    If UCase(FrmDisable.Name) = "FRMT46V143B" Then
        Select Case pStatAcao
        Case 70, 80, 81, 90
            lEstadoEmissao = False
        End Select
    End If
    '<-[FRH]

    For i = 0 To FrmDisable.Controls.Count - 1
        If (TypeOf FrmDisable.Controls(i) Is CheckBox) Or _
           (TypeOf FrmDisable.Controls(i) Is SSCheck) Or _
           (TypeOf FrmDisable.Controls(i) Is ComboBox) Or _
           (TypeOf FrmDisable.Controls(i) Is OptionButton) Or _
           (TypeOf FrmDisable.Controls(i) Is MSFlexGrid) Then

            FrmDisable.Controls(i).Enabled = False

            '[FRH.07]->'[� Status 70,80,81 ou 90 ent�o desabilite os SSCommand e troque suas figuras (cinza)]
        ElseIf TypeOf FrmDisable.Controls(i) Is SSCommand And Not lEstadoEmissao Then

            FrmDisable.Controls(i).Enabled = False
            FrmDisable.Controls(i).Picture = FrmDisable.cmdExtra(2).Picture

        ElseIf TypeOf FrmDisable.Controls(i) Is SSCommand And lEstadoEmissao Then

            FrmDisable.Controls(i).Enabled = False

            '<-[FRH]
        ElseIf TypeOf FrmDisable.Controls(i) Is TextBox Then
            ''  FrmDisable.Controls(I).Locked = True
            FrmDisable.Controls(i).Enabled = False
        End If
    Next

    FrmDisable.Tag = "Desabilitado"
End Function

Public Function gfRegistroMicrofocus(ByRef pMensagem As String) As Boolean
'Inclui o registro e a licen�a de uso do Runtime da Microfocus

    Dim lMensagem As String
    Dim lRootPath As String
    Dim lStrLicenca As String
    Dim lRegChave As String
    Dim lRegValor As String
    Dim lRegDadosVal As String
    Dim lConta As Byte
    Dim lHoje As Double
    Dim lDatLicen As Double
    Dim lPathRegOk As Boolean
    Dim lretorno As String
    Dim lHKey As Long

    Err = 0
    gfRegistroMicrofocus = False
    On Error GoTo Func_Erro

    '1- Obtem o caminho de instala��o
    If Right$(App.Path, 1) = "\" Then
        lRootPath = App.Path
    Else
        lRootPath = App.Path & "\"
    End If

    '-Consiste se vai executar o Step1
    If gM46V999.gfPreenchido(Dir(lRootPath & "Micro Focus\Application Server\ASLMPCSILENT.EXE")) = True And _
       gM46V999.gfPreenchido(Dir(lRootPath & "Micro Focus\Application Server\aslmpclocate.EXE")) = True Then
        If gM46V999.gfPreenchido(Dir(lRootPath & "\mfaslmf\mfasdb")) = False Then
            'N�o existindo base de licen�a, inicia a cria��o de uma
            GoTo Step1
        Else
            'Verifica se a licen�a esta registrada no App.Path
            lPathRegOk = True

            lRegChave = "SOFTWARE\Micro Focus\ASLMF"
            lRegValor = "ASLMFDIR"
            'Verifica se o caminho registrado para a licenca
            lRegDadosVal = gfRegistry_Read(lRegChave, lRegValor, _
                                           REG_SZ, HKEY_LOCAL_MACHINE)
            If Left(UCase(lRegDadosVal), Len(lRootPath & "mfaslmf")) <> UCase(lRootPath & "mfaslmf") Then
                lPathRegOk = False
                GoTo Step1
            End If

            lRegChave = "SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0\Environment"
            lRegValor = "PATH"
            'Verifica se o caminho registrado para o ambiente Cobol
            lRegDadosVal = gfRegistry_Read(lRegChave, lRegValor, _
                                           REG_EXPAND_SZ, HKEY_LOCAL_MACHINE)
            If Left(UCase(lRegDadosVal), Len("""" & lRootPath & "Micro Focus")) <> UCase("""" & lRootPath & "Micro Focus") Then
                lPathRegOk = False
                GoTo Step1
            End If

            lRegChave = "SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0\Setup"
            lRegValor = "RootDir"
            'Verifica se o caminho registrado para a pasta dos arquivos Microfocus
            lRegDadosVal = gfRegistry_Read(lRegChave, lRegValor, _
                                           REG_SZ, HKEY_LOCAL_MACHINE)
            If Left(UCase(lRegDadosVal), Len(lRootPath & "Micro Focus")) <> UCase(lRootPath & "Micro Focus") Then
                lPathRegOk = False
                GoTo Step1
            End If

            If lPathRegOk = False Then
                'Caminho incorreto, recria licen�a.
                GoTo Step1
            Else
                'Caminho OK!
                'Verifica se o registro foi  criado o mais de 2 dias
                lHoje = CDbl(Format(Now, "YYYYMMDD"))
                lDatLicen = CDbl(Format(FileDateTime(lRootPath & "\mfaslmf\mfasdb"), "YYYYMMDD"))
                If (lHoje - lDatLicen) > 1 Then
                    'Recria a licen�a
                    GoTo Step1
                Else
                    'N�o altera a licen�a neste momento
                    GoTo Step2
                End If
            End If
        End If    'If gM46V999.gfPreenchido(Dir(lRootPath & "\mfaslmf\mfasdb")) = False Then
    Else
        'N�o possui arquivos do Microfocus na marquina
        pMensagem = "N�o localizados arquivos de instala��o Microfocus" & Chr(13)
        pMensagem = pMensagem & "(" & lRootPath & "Micro Focus\Application Server\aslmpclocate.EXE" & ")"
        GoTo Func_Saida
    End If

Step1:
    '2- Cria a base de licen�as

    'Remove o registro Microfocus
    '************CUIDADO COM O PARAMETRO DE ENTRADA DESTA FUN��O!!!!!
    If gfRegistry_DeleteKey("SOFTWARE", "Micro Focus", lMensagem, _
                            HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na exclus�o de registro" & Chr(13)
        pMensagem = pMensagem & "SOFTWARE" & "\" & "Micro Focus"
        pMensagem = pMensagem & Chr(13) & lMensagem
        Call gM46V999.gpGraLog(0, pMensagem)
        'Continua o procedimento
        'Caso atualiza com sucesso, este erro pode ser desconsiderado
    Else
        'Cria as pastas da Microfocus uma a uma
        'incluindo permiss�o de acesso
        If gfRegistry_Create("SOFTWARE\Micro Focus", False, lHKey, lMensagem, HKEY_LOCAL_MACHINE) = False Then
            Call gM46V999.gpGraLog(0, pMensagem)
        End If
        If gfRegistry_Create("SOFTWARE\Micro Focus\ASLMF", False, lHKey, lMensagem, HKEY_LOCAL_MACHINE) = False Then
            Call gM46V999.gpGraLog(0, pMensagem)
        End If
        If gfRegistry_Create("SOFTWARE\Micro Focus\NetExpress", False, lHKey, lMensagem, HKEY_LOCAL_MACHINE) = False Then
            Call gM46V999.gpGraLog(0, pMensagem)
        End If
        If gfRegistry_Create("SOFTWARE\Micro Focus\NetExpress\4.0", False, lHKey, lMensagem, HKEY_LOCAL_MACHINE) = False Then
            Call gM46V999.gpGraLog(0, pMensagem)
        End If
        If gfRegistry_Create("SOFTWARE\Micro Focus\NetExpress\4.0\COBOL", False, lHKey, lMensagem, HKEY_LOCAL_MACHINE) = False Then
            Call gM46V999.gpGraLog(0, pMensagem)
        End If
        If gfRegistry_Create("SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0", False, lHKey, lMensagem, HKEY_LOCAL_MACHINE) = False Then
            Call gM46V999.gpGraLog(0, pMensagem)
        End If
        If gfRegistry_Create("SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0\Environment", False, lHKey, lMensagem, HKEY_LOCAL_MACHINE) = False Then
            Call gM46V999.gpGraLog(0, pMensagem)
        End If
        If gfRegistry_Create("SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0\Setup", False, lHKey, lMensagem, HKEY_LOCAL_MACHINE) = False Then
            Call gM46V999.gpGraLog(0, pMensagem)
        End If
    End If

    Call ChDir(lRootPath)
    If gM46V999.gfPreenchido(Dir(lRootPath & "\mfaslmf\mfasdb")) = True Then
        'Exclui a base de dados de licen�as p/ recriar.
        Kill lRootPath & "\mfaslmf\mfasdb"
    End If

    lStrLicenca = "ASLMPCSILENT 2 600000071753 01280 10000 0A8AD 4D709 3D9C"
    For lConta = 1 To 3
        Call ChDir(lRootPath)
        Call gM46V111.gpPausa(0.5)
        Call ExecCmd(lRootPath & "Micro Focus\Application Server\" & "aslmpclocate -q " & lRootPath & "mfaslmf", 1)
        Call gM46V111.gpPausa(1)
        Call ExecCmd(lRootPath & "Micro Focus\Application Server\" & lStrLicenca, 1)
        Call gM46V111.gpPausa(1)
        If gM46V999.gfPreenchido(Dir(lRootPath & "\mfaslmf\mfasdb")) = True Then
            Exit For
        End If
    Next lConta

    '3- Registra o NetExprees no windows Registry.
    lRegChave = "SOFTWARE\Micro Focus\ASLMF"
    lRegValor = "ASLMFDIR"
    lRegDadosVal = lRootPath & "mfaslmf"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '----
    lRegChave = "SOFTWARE\Micro Focus\ASLMF"
    lRegValor = "status"
    lRegDadosVal = "0"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '---
    lRegChave = "SOFTWARE\Micro Focus\NetExpress"
    lRegValor = "Version"
    lRegDadosVal = "4.0"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '---
    lRegChave = "SOFTWARE\Micro Focus\NetExpress"
    lRegValor = "DefaultVersion"
    lRegDadosVal = "4.0"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '---
    lRegChave = "SOFTWARE\Micro Focus\NetExpress\4.0\COBOL"
    lRegValor = "Version"
    lRegDadosVal = "4.0"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '---
    lRegChave = "SOFTWARE\Micro Focus\NetExpress\4.0\COBOL"
    lRegValor = "selector"
    lRegDadosVal = "S"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '---
    '***EXPAND"
    lRegChave = "SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0\Environment"
    lRegValor = "COBDIR"
    lRegDadosVal = """" & lRootPath & "Micro Focus\Application Server""" & ";%COBDIR%"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_EXPAND_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '---
    '***EXPAND"
    lRegChave = "SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0\Environment"
    lRegValor = "PATH"
    lRegDadosVal = """" & lRootPath & "Micro Focus\Application Server""" & ";" & lRootPath & "Micro Focus\Application Server;%PATH%"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_EXPAND_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '---
    '***EXPAND"
    lRegChave = "SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0\Environment"
    lRegValor = "COBHNF"
    lRegDadosVal = """" & lRootPath & "Micro Focus\Application Server""" & ";%COBHNF%"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_EXPAND_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '---
    lRegChave = "SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0\Setup"
    lRegValor = "UserName"
    lRegDadosVal = "SYAS"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '--
    lRegChave = "SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0\Setup"
    lRegValor = "UserCompany"
    lRegDadosVal = "YASUDA"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '--
    lRegChave = "SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0\Setup"
    lRegValor = "RootDir"
    lRegDadosVal = lRootPath & "Micro Focus\Application Server"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '--
    lRegChave = "SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0\Setup"
    lRegValor = "UninstallKey"
    lRegDadosVal = "Micro Focus Application Server for Net Express 4.0"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '--
    lRegChave = "SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0\Setup"
    lRegValor = "Version"
    lRegDadosVal = "4.0"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If
    '--
    lRegChave = "SOFTWARE\Micro Focus\NetExpress\4.0\COBOL\4.0\Setup"
    lRegValor = "Application Server for Net Express"
    lRegDadosVal = "4.0"
    If gfRegistry_Write(lRegChave, lRegValor, _
                        lRegDadosVal, REG_SZ, _
                        lMensagem, _
                        HKEY_LOCAL_MACHINE) = False Then
        pMensagem = "Falha na inclus�o de registro" & Chr(13)
        pMensagem = pMensagem & lRegChave & Chr(13)
        pMensagem = pMensagem & lRegValor & Chr(13)
        pMensagem = pMensagem & lRegDadosVal & Chr(13)
        pMensagem = pMensagem & lMensagem
        GoTo Func_Saida
    End If

    '---------------
Step2:
    'Inclui a variavel de ambiente COBSW = -F
    lRegChave = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
    lRegValor = "COBSW"
    'Verifica se j� possui a variavel de ambiente
    lRegDadosVal = gfRegistry_Read(lRegChave, lRegValor, _
                                   REG_SZ, HKEY_LOCAL_MACHINE)

    If UCase(lRegDadosVal) <> "-F" Then
        lRegDadosVal = "-F"
        If gfRegistry_Write(lRegChave, lRegValor, _
                            lRegDadosVal, REG_SZ, _
                            lMensagem, _
                            HKEY_LOCAL_MACHINE) = False Then
            pMensagem = "Falha na inclus�o de registro" & Chr(13)
            pMensagem = pMensagem & lRegChave & Chr(13)
            pMensagem = pMensagem & lRegValor & Chr(13)
            pMensagem = pMensagem & lRegDadosVal & Chr(13)
            pMensagem = pMensagem & lMensagem
            GoTo Func_Saida
        End If
    End If

    gfRegistroMicrofocus = True
Func_Saida:
    Exit Function

Func_Erro:
    If gM46V999.gfPreenchido(pMensagem) = True Then
        pMensagem = "gfRegistroMicrofocus - Erro " & pMensagem & Chr(13)
        pMensagem = pMensagem & "(" & lRootPath & ")"
    Else
        pMensagem = "gfRegistroMicrofocus - Erro " & Err & " - " & Error
    End If
    GoTo Func_Saida

End Function

Function HabilitaControles(FrmDisable As Form, Optional ByVal TagPermite As String)
    Dim i As Integer

    For i = 0 To FrmDisable.Controls.Count - 1
        If (TypeOf FrmDisable.Controls(i) Is CheckBox) Or _
           (TypeOf FrmDisable.Controls(i) Is SSCheck) Or _
           (TypeOf FrmDisable.Controls(i) Is OptionButton) Or _
           (TypeOf FrmDisable.Controls(i) Is ComboBox) Or _
           (TypeOf FrmDisable.Controls(i) Is MSFlexGrid) Then
            FrmDisable.Controls(i).Enabled = True
        ElseIf TypeOf FrmDisable.Controls(i) Is TextBox Then
            FrmDisable.Controls(i).Locked = False
        End If
    Next
    FrmDisable.Tag = ""
End Function

Public Function gfExecTransmissao(pTipoTransmissao As enum_TipTransmissao, pEsperar As Boolean) As Boolean
'Fun��o: executa transmiss�o.

    Dim lParam As String

    ChDir gM46V999.gAppPath
    If pTipoTransmissao = enum_TipTransmissao_Conexao Then
        lParam = "/CONEXAO"
    ElseIf pTipoTransmissao = enum_TipTransmissao_Transmissao Then
        lParam = "/TRANSMISSAO"
    ElseIf pTipoTransmissao = enum_TipTransmissao_Cotacao Then
        lParam = "/COTACAO"
    ElseIf pTipoTransmissao = enum_TipTransmissao_TransmHide Then
        lParam = "/TRANSMISSHIDE"
    End If
    'If gM46V999.gfPreenchido(Dir(gM46V999.gAppPath & "CLIENT001.BAT")) = False Then
    '    'Utiliza processo antigo caso n�o tenha os programas preparados
    '    Shell gM46V999.gAppPath & "client.bat", vbMinimizedNoFocus
    '    Exit Function
    'End If
    If pEsperar = True Then
        Call ExecCmd(gM46V999.gAppPath & "yasuda.exe " & lParam, 1)
    Else
        Shell gM46V999.gAppPath & "yasuda.exe " & lParam, vbNormalFocus
    End If
End Function

Public Function gfShell(ByVal pPrograma As String, Optional pParsYAS As Boolean, _
                        Optional pFocus As VbAppWinStyle = vbNormalFocus) As Boolean
'Procedure: aciona programa.

'Par�metro de entrada...pPrograma...Endere�o do programa.

    Dim lModulo As String   'Endere�o do m�dulo.
    Dim lShell As Variant    'Utilizado com comando SHELL.

    gfShell = False

    '1. Posiciona endere�o do m�dulo.
    lModulo = gM46V999.gAppPath & pPrograma & ".EXE"

    '2. Verifica se m�dulo est� presente.
    If Not gM46V999.gfPreenchido(Dir(lModulo)) Then
        Call gM46V999.gpGraLog(2, "FTE0001 - Programa n�o localizado: " & lModulo & "." & vbLf & vbLf & _
                                  "O sistema ser� encerrado.")
        End
    End If

    If pParsYAS = True Then
        lModulo = lModulo & " /YAS"
    End If

    '3. Aciona programa.
    On Error Resume Next
    lShell = Shell(lModulo, pFocus)
    If Err = 0 Then
        Call gM46V999.gpGraLog(0, "AGE0005 - Chamado programa " & pPrograma & ".")
    Else
        Call gM46V999.gpGraLog(2, "FTE0002 - Erro ao chamar programa " & pPrograma & " (" & Err & " - " & _
                                  Error & ")." & vbLf & vbLf & "O sistema ser� encerrado.")
        On Error GoTo 0
        End
    End If
    On Error GoTo 0

    gfShell = True
End Function

Public Sub gpBuscaCodigoCorretor(ByRef pCodCorretor As String, Optional ByRef pPriVez As Boolean)
'Procedure: Busca c�digo do corretor.

    Dim lEndP2842400 As String    'Endere�o do arquivo P2842400.EXE.
    Dim lEndP2842401 As String    'Endere�o do arquivo P2842401.EXE.
    Dim lNumP2842400 As Integer    'Utilizado com FreeFile para o arquivo P2842400.EXE.
    Dim lRegistro As String    'Registro.
    Dim lCodUsuario As String    ', Long     'C�digo do corretor.
    Dim lMensagem As String    'Mensagem.

'1. Obt�m c�digo do corretor.
    Select Case gM46V999.gTipUsuario
    Case 1      'Yasuda.
        pCodCorretor = "YASUDA"
        pPriVez = True
        If UCase(gM46V999.gAppPath) = "S:\SYASPROD\DIVERSOS\" Then
            lEndP2842400 = gM46V999.gAppPath & "P2842400.EXE"
            If gM46V999.gfPreenchido(Dir(lEndP2842400)) Then
                lNumP2842400 = FreeFile
                Open lEndP2842400 For Input Access Read As lNumP2842400
                Line Input #lNumP2842400, lRegistro
                pCodCorretor = Format(lRegistro, "0000000")
                Close lNumP2842400
                If UCase(pCodCorretor) <> "YASUDA" Then
                    gM46V999.gTipUsuario = e_Tipo_Usu_Corretor  '1001224 - Condi��es Diferenciada - Seguratec
                End If
            Else
                Call gM46V999.gpGraLog(1, "FGE0003 - " & gM46V999.gPrefixo & _
                                          "C�digo do corretor inv�lido (n�o localizado)." & vbLf & vbLf & _
                                          "O sistema ser� encerrado.")
                End
            End If
        End If
    Case 2      'Corretor.
        lEndP2842400 = gM46V999.gAppPath & "P2842400.EXE"
        If gM46V999.gfPreenchido(Dir(lEndP2842400)) = False Then
            lEndP2842400 = gM46V999.gAppPath & "P2842401.EXE"
        End If
        lEndP2842401 = gM46V999.gAppPath & "P2842401.EXE"
        If gM46V999.gfPreenchido(Dir(lEndP2842401)) Then
            'Se existe o arquivo P2842401.EXE, isto indica que � a primeira vez que se executa o
            'sistema. Posiciona indicador e altera o nome do arquivo para P2842400.EXE.
            pPriVez = True
            On Error Resume Next
            Kill lEndP2842401
            On Error GoTo 0
        Else
            pPriVez = False
        End If
        'Obs.: (a) Se existe o arquivo P2842400.EXE, obt�m o c�digo do corretor do registro.
        '      (b) Se n�o existe o arquivo P2842400.EXE, solicitar ao usu�rio.
        If gM46V999.gfPreenchido(Dir(lEndP2842400)) Then
            lNumP2842400 = FreeFile
            Open lEndP2842400 For Input Access Read As lNumP2842400
            Line Input #lNumP2842400, lRegistro
            pCodCorretor = Format(lRegistro, "0000000")
            Close lNumP2842400
        Else
            Call gM46V999.gpGraLog(1, "FGE0003 - " & gM46V999.gPrefixo & _
                                      "C�digo do corretor inv�lido (n�o localizado)." & vbLf & vbLf & _
                                      "O sistema ser� encerrado.")
            End
        End If
        If gM46V999.gNomUsuario = "XXXXXXXX" Then
            If gfBuscaUsuarioTransmissao(lCodUsuario, lMensagem) = False Then
                Call gM46V999.gpGraLog(1, lMensagem & vbLf & vbLf & "O sistema ser� encerrado.")
                End
            End If
            gM46V999.gNomUsuario = lCodUsuario
            gM46V999.gNomUsuario6 = Left$(gM46V999.gNomUsuario, 6)
        End If
    Case 3
        pCodCorretor = gM46V999.gCod_Corretor
    End Select
End Sub

