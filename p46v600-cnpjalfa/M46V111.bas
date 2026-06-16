Attribute VB_Name = "basM46V111"
Option Explicit
Public gRegPadrao       As Boolean          'Indicador de registro padrão.
Public gComNosNumero    As String       'Registro de comunicação - nosso número.
Public gNumPedido       As String       'Número do pedido.
Public gNum_Pi          As Long
Public gSaida           As Boolean      'Indica que trata-se de saída.
Public gProJurIOF       As Double           'Produto: juros de IOF.
Public gCod_Plano       As Integer
'Public gQuaErros        As Byte             'Quantidade de erros.
Public gCod_Dizer       As Integer
Public gModalidade      As String       'Nome da Modalidade
Public gTipoGrupMulti   As e_GrupoRamo  'Define Ramos que fazem parte de MultiSeguros (112,113,114)
Public gCod_Regiao      As Integer      'Código da Regiao
Public gDesc_Coml           As Double       'Desconto Comercial para os Planinhos
Public gDesc_Prog           As Double       'Desconto Progressivo para os Planinhos
Public gDesc_Bonus          As Double       'Desconto Bônus para os Planinhos

Declare Sub PontSyas Lib "D046O001.dll" (Reg_Parm As ty_Parm)
'Utilizado para executar pausa no codigo do programa
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type ty_Parm
    Cod_User    As String * 6
    Cod_Ramo As Integer
    Num_PI  As Long
    Num_Ped  As Long
    Num_Item As Long
    Num_Apol_Ant  As Long
    Num_Item_Ant  As Long
    Tip_Proc As Long
    flag_debug As Integer
    Cod_Dizer As Integer
    Versao_Corr As Integer
    Cod_Retorno As Integer
    Cod_Err_Nor_1 As Long
    Mens_Err_1 As String * 80
    Cod_Err_Nor_2 As Long
    Mens_Err_2 As String * 80
    Cod_Err_Nor_3 As Long
    Mens_Err_3 As String * 80
    Cod_Err_Nor_4 As Long
    Mens_Err_4 As String * 80
    Cod_Err_Nor_5 As Long
    Mens_Err_5 As String * 80
    Cod_Err_Nor_6 As Long
    Mens_Err_6 As String * 80
    Cod_Err_Nor_7 As Long
    Mens_Err_7 As String * 80
    Cod_Err_Nor_8 As Long
    Mens_Err_8 As String * 80
    Cod_Err_Nor_9 As Long
    Mens_Err_9 As String * 80
    Cod_Err_Nor_10 As Long
    Mens_Err_10 As String * 80
    Cod_Err_Nor_11 As Long
    Mens_Err_11 As String * 80
    Cod_Err_Nor_12 As Long
    Mens_Err_12 As String * 80
    Cod_Err_Nor_13 As Long
    Mens_Err_13 As String * 80
    Cod_Err_Nor_14 As Long
    Mens_Err_14 As String * 80
    Cod_Err_Nor_15 As Long
    Mens_Err_15 As String * 80
    Cod_Err_Nor_16 As Long
    Mens_Err_16 As String * 80
    Cod_Err_Nor_17 As Long
    Mens_Err_17 As String * 80
    Cod_Err_Nor_18 As Long
    Mens_Err_18 As String * 80
    Cod_Err_Nor_19 As Long
    Mens_Err_19 As String * 80
    Cod_Err_Nor_20 As Long
    Mens_Err_20 As String * 80
 End Type

' Declare API functions.
'Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
'   (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Enum e_TipProc
    e_TipProc_ConsComum = 1
    e_TipProc_ConsItem = 2
    e_TipProc_Normas = 3
End Enum


Public Enum e_GrupoRamo
    e_GrupoRamo_NDefinido = 0
    e_GrupoRamo_Multi = 1
    e_GrupoRamo_570_820 = 2
    e_GrupoRamo_Transp_Diversos = 3
    e_GrupoRamo_Auto = 4
    e_GrupoRamo_Acidentes_pessoais = 5
    e_GrupoRamo_Generico = 6
End Enum


Public Enum enum_TipTransmissao
    enum_TipTransmissao_Transmissao = 1
    enum_TipTransmissao_Conexao = 2
    enum_TipTransmissao_Cotacao = 3
    enum_TipTransmissao_TransmHide = 4
End Enum

Public Declare Function WaitForSingleObject Lib "kernel32" _
                                            (ByVal hHandle As Long, _
                                             ByVal dwMilliseconds As Long) As Long

Public Declare Function CreateProcessA Lib "kernel32" _
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

Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessID As Long
        dwThreadID As Long
End Type

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

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&

Private Const INFINITE = -1&

'Utilizado para executar pausa no codigo do programa
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Public Enum e_ColFlxSelecionadas
    ColCodCobert = 0
    ColNomCobert = 1
    ColValLMG = 2
    ColLimiteMaximo = 3
    ColValRisco = 4
    ColTabCobert = 5
    ColPerFranq = 6
    ColValFranq = 7
    ColDescComl = 8
    ColDescExp = 9
    ColCobObrigat = 10
    ColPerLimite = 11
    ColPermitePInd = 12
    ColPInd = 13
    ColCoefFranq = 14
    ColTaxaFinal = 15
    ColTaxaAlterada = 16
    ColValFranMin = 17
    ColCobertPerLimite = 18
    ColValISAte = 19
End Enum


Public Const gFonteEspecificacao = "Arial"


Public Function gfVerSegSiscota(pobjstcPed As stcA46V702B) As Boolean
Dim lSelect As String
Dim lBDSiscota  As ADODB.Connection
Dim lrst    As ADODB.Recordset
Dim lMensagem   As String
Dim lCodSegurado    As Long
Dim lCpfCnpj        As String

    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = "2" Then
        If gfPreenchido(pobjstcPed.objstcCtrlEmis.Num_Cotac) Then
            '3. Abre arquivo.
            If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lBDSiscota, lMensagem) = False Then
                Call gpGraLog(2, lMensagem)
                Exit Function
            End If
        
            lSelect = "select * from siscota.siscota.tab_cotacao where ccotacao= " & mfFormatarCamposSQL(pobjstcPed.objstcCtrlEmis.Num_Cotac)
                      
            If gfObtRegistro(lBDSiscota, lSelect, lrst, lMensagem) = False Then
                Call gpFecha2(lBDSiscota)
                Call gpGraLog(2, lMensagem)
                Exit Function
            End If
            
            If Not lrst.EOF Then
                lCodSegurado = lrst("CSegurado")
                lCpfCnpj = lrst("Numdoctoseg")
                If lCodSegurado = 0 Then
                    If pobjstcPed.Tip_Pessoa = e_TipPessoa_Fisica Then
                        lSelect = "select * from siscota.siscota.tab_segurado where CPF = '" & Format(pobjstcPed.Num_Cgc_Cpf, "00000000000") & "'"
                    Else
                        lSelect = "select * from siscota.siscota.tab_segurado where CGC = '" & Format(pobjstcPed.Num_Cgc_Cpf, "00000000000000") & "'"
                    End If
                    
                    If gfObtRegistro(lBDSiscota, lSelect, lrst, lMensagem) = False Then
                        Call gpFecha2(lBDSiscota)
                        Call gpGraLog(2, lMensagem)
                        Exit Function
                    End If
                    If lrst.EOF Then
                        lSelect = "select max(csegurado) as codigo from siscota.siscota.tab_segurado"
                        
                        If gfObtRegistro(lBDSiscota, lSelect, lrst, lMensagem) = False Then
                            Call gpFecha2(lBDSiscota)
                            Call gpGraLog(2, lMensagem)
                            Exit Function
                        End If
                        
                        lCodSegurado = lrst("Codigo") + 1
                    
                        lSelect = " INSERT INTO [SISCOTA].[SISCOTA].[TAB_SEGURADO]([CSEGURADO], " & _
                                " [SEGURADO], [ASEGURADO], [XPESSOA], [CGC], [CPF], [CUSER], [DSYS], " & _
                                " [DATUAL], [SREG], [ABREV_SISCOTA], [HSYS], [CUSER_EXC], [DEXC], [HEXC]) VALUES( "
                                                            
                        lSelect = lSelect & lCodSegurado & "," & _
                                mfFormatarCamposSQL(Left(pobjstcPed.Nom_Segurado, 70)) & "," & _
                                mfFormatarCamposSQL(Left(pobjstcPed.Nom_Segurado, 15)) & ","
                        If pobjstcPed.Tip_Pessoa = e_TipPessoa_Fisica Then
                            lSelect = lSelect & "'F',''," & pobjstcPed.Num_Cgc_Cpf & ","
                        Else
                            lSelect = lSelect & "'J'," & pobjstcPed.Num_Cgc_Cpf & ",'',"
                        End If
                        lSelect = lSelect & "'SYAS',"
                        lSelect = lSelect & Left(gfForData(3), 8) & ","
                        lSelect = lSelect & "0,'A','M',"
                        lSelect = lSelect & Right(gfForData(3), 6) & ", NULL, NULL, NULL )"
                        
                        If gfExeSQL(lBDSiscota, lSelect, lMensagem) <> 0 Then
                            Call gpGraLog(1, lMensagem)
                            Exit Function
                        End If
                    Else
                        lCodSegurado = lrst("CSegurado")
                    End If
                End If
                    
                lSelect = " Update siscota.siscota.tab_cotacao set csegurado = " & lCodSegurado
                lSelect = lSelect & " , Numdoctoseg = " & mfFormatarCamposSQL(pobjstcPed.Num_Cgc_Cpf)
                lSelect = lSelect & " where ccotacao = " & mfFormatarCamposSQL(pobjstcPed.objstcCtrlEmis.Num_Cotac)
                
                If gfExeSQL(lBDSiscota, lSelect, lMensagem) <> 0 Then
                    Call gpGraLog(1, lMensagem)
                    Exit Function
                End If
            End If
        End If
    End If

End Function
Public Function gfExecDLL(ByRef pobjstcPed As stcA46V702B, ByRef pobjstcPedLoc As stcA46V708B, _
                        ByVal pTipoProc As e_TipProc, ByVal pItem As Long, _
                        ByRef pCod_Retorno As Integer, pTipHomologacao As String) As String
    ' 1 - para consistir comum
    ' 2 - para consistir o item
    ' 3 - para normas
Dim lParm       As ty_Parm
Dim lRetorno    As String
Dim lNomeArq    As String
Dim lMensagem As String


    'vou criar um arquivo somente para validar se o usuário é único no processamento.
    If mpVerCalculoParalelo() Then
        Call gpGraLog(1, lMensagem)
        'Exit Function
    End If


    lParm.Cod_Ramo = pobjstcPed.Cod_Ramo
    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        lParm.Cod_User = "SYAS98"
    Else
        lParm.Cod_User = "SYAS99"
    End If
    lParm.Num_PI = gNum_Pi
    lParm.Num_Ped = pobjstcPed.Num_Ped
    lParm.Num_Item = pItem
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Or pobjstcPedLoc.Tip_Emissao = e_TipEmissao_Endosso Then
'Acerto
        'lParm.Num_Apol_Ant = objstcPedido.Num_Apol_An
        lParm.Num_Apol_Ant = pobjstcPedLoc.Num_Apol_Anterior
        lParm.Num_Item_Ant = Val(pItem)
    Else
        lParm.Num_Apol_Ant = 0
        lParm.Num_Item_Ant = 0
    End If
    lParm.flag_debug = 1
    lParm.flag_debug = 0
    lParm.Tip_Proc = pTipoProc
    'se pTipoProc = 3, = Vernormas - Passar Tarifa no número do PI
    If pTipoProc = 3 Then
        lParm.Num_PI = pobjstcPed.Cod_Produto
    End If
    lParm.Cod_Dizer = pobjstcPed.Cod_Dizer
    If pTipHomologacao = "1" Then
        lParm.Versao_Corr = 2   'versão homologação
    Else
        lParm.Versao_Corr = 1
    End If
    
    
    'VbParm.
    ''GoTo Thais
    Call PontSyas(lParm)
    
'*  ret-cod               10 - Erro OPEN OUTPUT VSAM              *
'*                        11 - Erro Delete Vsam                   *
'*                        12 - Erro Write TMENU                   *
'*                        13 - Erro Write T01O002A                *
'*                        14 - Erro no Open I-O                   *
'*                        15 - Quantidade de Registros inválida   *
'*                        16 - Erro na abertura do banco          *
'*                        17 - Erro no diretorio corrente         *
'*                        18 - Erro no Join do nome do arquivo    *
'*                        19 - Erro no Read VSAM                  *
'*                        40 - Problemas nos dados de entrada     *
'*                        50 - Erros na tela P01O002              *
'*                        52 - Erros na tela P01O003              *
'*                        53 - Erros na tela P01O004              *
'*                        54 - Erros na tela P01O062              *
'*                        55 - Aceitação depende de inspeção      *
'*                        80 - Proposta nao aceita                *

    lRetorno = lParm.Cod_Retorno & " -- " & Trim(lParm.Cod_Err_Nor_1) & " - " & Trim(lParm.Mens_Err_1) & _
                " - " & Trim(lParm.Cod_Err_Nor_2) & " - " & Trim(lParm.Mens_Err_2) & _
                " - " & Trim(lParm.Cod_Err_Nor_3) & " - " & Trim(lParm.Mens_Err_3) & _
                " - " & Trim(lParm.Cod_Err_Nor_4) & " - " & Trim(lParm.Mens_Err_4) & _
                " - " & Trim(lParm.Cod_Err_Nor_5) & " - " & Trim(lParm.Mens_Err_5) & _
                " - " & Trim(lParm.Cod_Err_Nor_6) & " - " & Trim(lParm.Mens_Err_6) & _
                " - " & Trim(lParm.Cod_Err_Nor_7) & " - " & Trim(lParm.Mens_Err_7) & _
                " - " & Trim(lParm.Cod_Err_Nor_8) & " - " & Trim(lParm.Mens_Err_8) & _
                " - " & Trim(lParm.Cod_Err_Nor_9) & " - " & Trim(lParm.Mens_Err_9) & _
                " - " & Trim(lParm.Cod_Err_Nor_10) & " - " & Trim(lParm.Mens_Err_10) & _
                " - " & Trim(lParm.Cod_Err_Nor_11) & " - " & Trim(lParm.Mens_Err_11) & _
                " - " & Trim(lParm.Cod_Err_Nor_12) & " - " & Trim(lParm.Mens_Err_12) & _
                " - " & Trim(lParm.Cod_Err_Nor_13) & " - " & Trim(lParm.Mens_Err_13) & _
                " - " & Trim(lParm.Cod_Err_Nor_14) & " - " & Trim(lParm.Mens_Err_14) & _
                " - " & Trim(lParm.Cod_Err_Nor_15) & " - " & Trim(lParm.Mens_Err_15) & _
                " - " & Trim(lParm.Cod_Err_Nor_16) & " - " & Trim(lParm.Mens_Err_16) & _
                " - " & Trim(lParm.Cod_Err_Nor_17) & " - " & Trim(lParm.Mens_Err_17) & _
                " - " & Trim(lParm.Cod_Err_Nor_18) & " - " & Trim(lParm.Mens_Err_19) & _
                " - " & Trim(lParm.Cod_Err_Nor_20) & " - " & Trim(lParm.Mens_Err_20)
    lRetorno = Trim(Replace(lRetorno, "- 0 -", ""))

    If lRetorno = "0 -" Then
Thais:
        gfExecDLL = ""
    Else
        gfExecDLL = lRetorno
    End If
    pCod_Retorno = lParm.Cod_Retorno
    
    Call mpDeletarArqCalculo
End Function
Private Sub mpDeletarArqCalculo()
Dim lDiretorio  As String
Dim lNomeArquivo    As String
Dim lNomeComputador As String

    On Error Resume Next
    lDiretorio = Replace(UCase(gEndP00MULTI), "P00MULTI.MDB", "")
    lNomeComputador = GetLoginComputerName()
    lNomeArquivo = lNomeComputador & ".PROCESSANDO"
    Kill lDiretorio & lNomeArquivo
    Err.Clear
    On Error GoTo 0
End Sub

Private Function mpVerCalculoParalelo() As Boolean
Dim lDiretorio  As String
Dim lOutroArq   As String
Dim lArquivo    As Integer
Dim lNomeArquivo    As String
Dim lNomeComputador As String
    On Error GoTo ControleErro
    lDiretorio = Replace(UCase(gEndP00MULTI), "P00MULTI.MDB", "")
    lNomeComputador = GetLoginComputerName()
    lNomeArquivo = lNomeComputador & ".PROCESSANDO"
Retorno:
    lOutroArq = Dir(lDiretorio & "*.PROCESSANDO")
    If gfPreenchido(lOutroArq) Then
        If FileDateTime(lDiretorio & lOutroArq) < (DateAdd("s", -20, Now)) Then
            Kill lDiretorio & lOutroArq
            GoTo Retorno
        Else
            'If lNomeMaq <> Left(lOutroArq, InStr(1, lOutroArq, ".") - 1) Then
                'Call gpMenErro("Existe outra estação efetuando um cálculo simultâneo.(" & lNomeComputador & "," & Left(lOutroArq, InStr(1, lOutroArq, ".") - 1) & ")")
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
    Call gpGraLog(0, "mpVerCalculoParalelo - " & Err.Description)
End Function

Public Sub gpObterRegComunicacao(ByVal pExibeMensagem As Boolean, _
                                 ByVal pPropTabela As String, _
                                 ByVal pNomeTabImport As String)

'********************Variáveis padrão*************************************
Dim lCkPoint        As String      'Variavel utilizada para determinar pontos onde o código já passou.
Dim lCkStr          As String       'Variavel utilizada para guardar a posição das variaveis até o ultimo
                                    'lCkPoint. (posição das variaveis em ordem de declaração)
'********************Variáveis Declaradas para o processamento************
Dim lObjPedCom      As clsA46V709A       'Objeto da classe clsA46V709A (PedCom)
Dim lObjstcPedCom   As stcA46V709B       'Objeto da classe de estrutura stcA46V709B (PedCom)
    
    Err = 0
    lCkPoint = ""
    lCkStr = ""
    'Procedure: Obtem o registro de comunicação.
    
    'Instancio os objetos.
On Error GoTo ControleErro
    lCkPoint = "1"
    Set lObjPedCom = New clsA46V709A
    lCkPoint = "2"
    'Busco o registro de comunicação.
    If Not lObjPedCom.mfCarregarColecao(pPropTabela, pNomeTabImport) Then
        Call gpGraLog(1, lObjPedCom.UltimoErro)
        GoTo SAIDA
    End If
    lCkPoint = "3"
    If lObjPedCom.Colecao.Count = 0 Then
        If pExibeMensagem = False Then
            GoTo SAIDA
        Else
            Call gpGraLog(2, "FTE2003 - Registro de comunicação não localizado.")
            GoTo SAIDA
        End If
    End If
    lCkPoint = "4"
    Set lObjstcPedCom = lObjPedCom.Colecao.Item(1)
    lCkPoint = "5"
    gRegPadrao = False
    If Not gfPreenchido(lObjstcPedCom.NOSSO_NUMERO) Then
        Call gpGraLog(2, "FTE2003 - Registro de comunicação inválido (nosso número não " & _
                         "preenchido)." & vbLf & vbLf & "Programa será encerrado.")
        GoTo SAIDA
    End If
    lCkPoint = "6"
    gComNosNumero = Trim$(lObjstcPedCom.NOSSO_NUMERO)
    Select Case Len(gComNosNumero)
        Case Is <= 6
            Call gpGraLog(2, "FTE2003 - Registro de comunicação inválido (nosso número com tamanho " & _
                             "de " & Len(gComNosNumero) & " posições)." & vbLf & vbLf & _
                             "Programa será encerrado.")
            GoTo SAIDA
        Case 20
            If Not IsNumeric(gComNosNumero) Then
                Call gpGraLog(2, "FTE2003 - Registro de comunicação inválido (nosso número não numérico)." & _
                                 vbLf & vbLf & "Programa será encerrado.")
                GoTo SAIDA
            End If
            If CDbl(gComNosNumero) = 0 Then
                Call gpGraLog(2, "FTE2003 - Registro de comunicação inválido (nosso número zerado)." & _
                                 vbLf & vbLf & "Programa será encerrado.")
                GoTo SAIDA
            End If
        Case Else
            If Left$(gComNosNumero, 6) <> "PADRAO" Then
                Call gpGraLog(2, "FTE2003 - Registro de comunicação inválido (nosso número não e " & _
                                 "diferente de PADRÃO)." & vbLf & vbLf & "Programa será encerrado.")
                GoTo SAIDA
            End If
            gRegPadrao = True
    End Select
    If gPerfilAcesso = e_Perfil_Triagem Then
        gNum_Pi = lObjstcPedCom.Dat_Alt
    End If
    
    lCkPoint = "7"
    'Excluo o mesmo registro da base de dados.
    If Not lObjPedCom.mfExcluir(pPropTabela, pNomeTabImport) Then
        Call gpGraLog(2, lObjPedCom.UltimoErro)
        GoTo SAIDA
    End If
    lCkPoint = "8"
    
SAIDA:
    GoSub FechaObj
    Exit Sub
ControleErro:
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gpObterRegComunicacao ERRO - " & Err & " " & Err.Description & _
                Chr(13) & "CheckPoint " & lCkPoint)
    GoTo SAIDA
FechaObj:
    Set lObjPedCom = Nothing
    Set lObjstcPedCom = Nothing
    Return
End Sub


Public Function gfGraRegComunicacao(ByVal pNosNumero As String, _
                                    ByRef pMensagem As String, _
                                    Optional pBasDados As ADODB.Connection) As Boolean
    'Função: grava registro de comunicação entre P46V001 e P46V100.

    'Parâmetros de entrada...pBasDados........Base de dados.
    '                        pNosNumero...Nosso número.
    'Parâmetro de saída......pMensagem....Mensagem.
    
Dim lObjPedCom      As clsA46V709A      'Objeto da classe clsA46V709A (TAB_PED_COM)
Dim lObjstcPedCom   As stcA46V709B      'Objeto da classe de estrutura stcA46V709B (TAB_PED_COM)
    
Dim lPropTabela     As String
Dim lNomeTabImport  As String

On Error GoTo ControleErro
    gfGraRegComunicacao = False
        
    Call gfObtSYASMULTINI(lPropTabela)
    lNomeTabImport = ""
        
    '1. Consiste parâmetros.
    'a) Nosso número.
    If Not gfPreenchido(pNosNumero) Then
        pMensagem = "FTE0074 - Erro ao criar registro de comunicação (nosso número não preenchido)."
        GoTo SAIDA
    End If
    pNosNumero = UCase$(Trim$(pNosNumero))
    If IsNumeric(pNosNumero) Then
        If Val(pNosNumero) = 0 Then
            pMensagem = "FTE0074 - Erro ao criar registro de comunicação (nosso número zerado)."
            GoTo SAIDA
        End If
        If Len(pNosNumero) > 20 Then
            pMensagem = "FTE0074 - Erro ao criar registro de comunicação (nosso número com mais de " & _
                        "de 20 posições)."
            GoTo SAIDA
        End If
        pNosNumero = String(20 - Len(pNosNumero), "0") & pNosNumero
    Else
        If Len(pNosNumero) <> 6 Then
            pMensagem = "FTE0074 - Erro ao criar registro de comunicação (nosso número não numérico " & _
                        "e tamanho diferente de 6 posições)."
            GoTo SAIDA
        End If
        pMensagem = "FTE0074 - Erro ao criar registro de comunicação (nosso número não numérico."
    End If

    '2. Instancio as classes
    Set lObjPedCom = New clsA46V709A
    Set lObjstcPedCom = New stcA46V709B
    Set lObjPedCom.ConexaoBD = pBasDados
    
    If Not lObjPedCom.mfExcluir(lPropTabela, lNomeTabImport) Then
        pMensagem = lObjPedCom.UltimoErro
        GoTo ControleErro
    End If
    
    Set lObjstcPedCom = New stcA46V709B
    With lObjstcPedCom
        If pNosNumero = "PADRAO" Then
            .NOSSO_NUMERO = pNosNumero & gNomUsuario6
        Else
            .NOSSO_NUMERO = pNosNumero
        End If
        .Cod_User = gNomUsuario6
        .Dat_Alt = gfForData(3)
    End With
    If Not lObjPedCom.mfIncluir(lObjstcPedCom, lPropTabela, lNomeTabImport) Then
        pMensagem = lObjPedCom.UltimoErro
        GoTo ControleErro
    End If

    gfGraRegComunicacao = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfGraRegComunicacao ERRO - " & Err & " " & Err.Description & " " & pMensagem)
    GoTo SAIDA
FechaObj:
    Set lObjPedCom = Nothing
    Set lObjstcPedCom = Nothing
    Return
End Function

Public Sub gpPausa(ByVal lSegundos)

Dim I
    Dim lMiliSegundos   As Double
    
        lMiliSegundos = lSegundos * 1000

    I = DoEvents
    Sleep (lMiliSegundos)
    I = DoEvents
    
    Exit Sub
    
End Sub


Public Sub gpPausa_wf(ByVal lSegundos)
'Pausa a execução do codigo durante o total de segundos
'recebido no parametro (Libera o uso da CPU durante a pausa)
    
Dim lMiliSegundos   As Double
Dim lDoEvents
    
    lMiliSegundos = lSegundos * 1000

    lDoEvents = DoEvents
    Sleep (lMiliSegundos)
    lDoEvents = DoEvents
    
    Exit Sub
    
End Sub

Public Sub gpObtModalidade(ByVal pCod_Plano As Integer, ByVal pCodRamo As Long)
    'Procedure: obtém produto e informações dependentes de produto.

Dim lbdTAB_Plano    As ADODB.Connection 'Base de dados - TAB_PROD.
Dim lI              As Byte             'Utilizado com For...Next.
Dim lMensagem       As String           'Mensagem.
Dim lPosicao        As Integer          'Utilizado com InStr.
Dim lrsTAB_Plano    As ADODB.Recordset  'Registro - TAB_PROD.
Dim lCarga          As Boolean
Dim lLocal          As Integer
Dim lTemp           As String
Dim lSelect         As String

    '1. Inicializa variáveis globais - dados do produto.
    gModalidade = " "

    '2. Se produto zerado, encerra processo.
    If pCod_Plano = 0 Or pCodRamo = 0 Then
        GoTo SAIDA
    End If

    '3. Abre arquivo.
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTAB_Plano, lMensagem) = False Then
        Call gpGraLog(2, lMensagem)
        GoTo SAIDA
    End If

    '4. Obtém produto informado.
    lSelect = "SELECT * FROM TAB_Plano WHERE Cod_Ramo = " & pCodRamo
    lSelect = lSelect & " and Cod_Produto = " & pCod_Plano
    lSelect = lSelect & " and Dat_Vig_Ini <= " & gfForData(1, Date)
    lSelect = lSelect & " order by Cod_Produto"
              
    If gfObtRegistro(lbdTAB_Plano, lSelect, lrsTAB_Plano, lMensagem) = False Then
        Call gpGraLog(2, lMensagem)
        GoTo SAIDA
    End If
    
    If lrsTAB_Plano.EOF = True Then
        Call gpGraLog(1, "FNL0002 - " & gPrefixo & "Produto não localizado (" & _
                         Format$(pCod_Plano, "000") & "/" & Format$(pCodRamo, "000") & ").")
        GoTo SAIDA
    End If

    lTemp = lrsTAB_Plano("Nom_Prot")
    lLocal = InStr(1, lrsTAB_Plano("Nom_Prot"), "- ")
    If lLocal > 0 Then
        lLocal = InStr(lLocal + 1, lrsTAB_Plano("Nom_Prot"), "- ")
        If lLocal > 0 Then
            lTemp = Mid(lrsTAB_Plano("Nom_Prot"), 1, lLocal - 1)
        End If
    End If
                
    '5. Carrega variáveis globais.
    gModalidade = lTemp
    
SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Call gpFecha3(lrsTAB_Plano, lbdTAB_Plano)
    Set lrsTAB_Plano = Nothing
    Set lbdTAB_Plano = Nothing
    Return
End Sub

'*******************************************
' gpExcluiRegistros agora contida no BAS 999
'*******************************************

'###PROJETO SISCOTA
'DESCOMENTAR AS LINHAS
Public Function gfVerStatusSiscota(pobjstcPed As stcA46V702B, pStatusSiscota As Integer, pRamoSiscota As Long) As Boolean
Dim lSelect             As String
Dim lBDSiscota          As ADODB.Connection
Dim lrst                As ADODB.Recordset
Dim lMensagem           As String

    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = "2" Then
        If gfPreenchido(pobjstcPed.objstcCtrlEmis.Num_Cotac) Then
            '3. Abre arquivo.
            If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lBDSiscota, lMensagem) = False Then
                Call gpGraLog(2, lMensagem)
                Exit Function
            End If
        
            lSelect = "select * from siscota.siscota.tab_cotacao where ccotacao= " & mfFormatarCamposSQL(pobjstcPed.objstcCtrlEmis.Num_Cotac)
                      
            If gfObtRegistro(lBDSiscota, lSelect, lrst, lMensagem) = False Then
                Call gpFecha2(lBDSiscota)
                Call gpGraLog(2, lMensagem)
                Exit Function
            End If
            
            If Not lrst.EOF Then
                pRamoSiscota = lrst("CRamo")
                pStatusSiscota = lrst("SCotacao")
            End If
        End If
    End If

End Function

Public Function gfBuscaComissaoParm(pCorretor As Long, pUnidade As Integer, _
                                pComissao As Single, pValLMG As Long, pClassif As String, _
                                pMensagem As String, pobjstcPed As stcA46V702B, _
                                pobjstcPedLoc As stcA46V708B) As Boolean '1001224 - Condições Diferenciada - Seguratec

Dim bdP00MULTI          As ADODB.Connection
Dim lrst                As ADODB.Recordset
Dim lSelect             As String
'1001224 - Condições Diferenciada - Seguratec
Dim lobjstcPedCobert    As stcA46V704B
Dim lRetorno            As String

    '1001224 - Condições Diferenciada - Seguratec
    If gCodProduto >= 1000 Then
        If pobjstcPedLoc.ColecaoPedCobert.Count > 0 Then
            Set lobjstcPedCobert = pobjstcPedLoc.ColecaoPedCobert.Item(1)
        Else
            Set lobjstcPedCobert = Nothing
        End If
        If Not gfVerParametro(11, pobjstcPed, pobjstcPedLoc, lobjstcPedCobert, lRetorno, pMensagem) Then
            Exit Function
        End If
        pComissao = Val(lRetorno)
    Else
        If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdP00MULTI, _
                pMensagem) = False Then Exit Function
        
        lSelect = "Select * from TAB_PARM_DIVERSOS"
        lSelect = lSelect & " Where Cod_Produto = " & gCodProduto
        lSelect = lSelect & " and Cod_Protocolo = " & gCodProtocolo
        lSelect = lSelect & " and Cod_Sequencia <= " & gCodSequencia
        If gCodProduto >= 998 Then
            lSelect = lSelect & " and TIPO_VALOR = 11"
        Else
            lSelect = lSelect & " and TIPO_VALOR = 10"
        End If
        lSelect = lSelect & " and cod_ramo = " & gCod_Ramo
        lSelect = lSelect & " order by NIVEL_PRIORIDADE"
        
        If gfObtRegistro(bdP00MULTI, lSelect, lrst, _
                pMensagem) = False Then Exit Function
        
        Do While Not lrst.EOF
            Select Case lrst("TIPO_CHAVE")
                Case 0  'Tarifa básica
                    pComissao = lrst("RETORNO")
                Case 1  'Unidade
                    If Val(lrst("CHAVE_1")) = pUnidade Then
                        pComissao = lrst("RETORNO")
                    End If
                Case 2  'Corretor
                    If Val(lrst("CHAVE_1")) = pCorretor Then
                        pComissao = lrst("RETORNO")
                    End If
                Case 11  'Comissão
                    If Val(lrst("CHAVE_1")) = pCorretor Then
                        If Val(lrst("CHAVE_2")) = Val(pClassif) Then
                            If Val(lrst("CHAVE_3")) = 1 Then
                               pComissao = lrst("RETORNO")
                               Exit Do
                            End If
                        End If
                    End If
    
            End Select
            lrst.MoveNext
        Loop
        Call gpFecha3(lrst, bdP00MULTI)
    End If
    gfBuscaComissaoParm = True
End Function

Public Function gfBuscaDescontoParm(pCorretor As Long, pUnidade As Integer, _
                                pDesconto As Double, pValLMG As Long, pClassif As String, _
                                pSegurado As Long, pMensagem As String) As Boolean

Dim bdP00MULTI  As ADODB.Connection
Dim lrst        As ADODB.Recordset
Dim lSelect     As String

    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdP00MULTI, _
            pMensagem) = False Then Exit Function
    
    lSelect = "Select * from TAB_PARM_DIVERSOS"
    lSelect = lSelect & " Where Cod_Produto = " & gCodProduto
    lSelect = lSelect & " and Cod_Protocolo = " & gCodProtocolo
    lSelect = lSelect & " and Cod_Sequencia <= " & gCodSequencia
    lSelect = lSelect & " and TIPO_VALOR = 10"
    lSelect = lSelect & " and cod_ramo = " & gCod_Ramo
    lSelect = lSelect & " order by NIVEL_PRIORIDADE"
    
    If gfObtRegistro(bdP00MULTI, lSelect, lrst, _
            pMensagem) = False Then Exit Function
    
    Do While Not lrst.EOF
        Select Case lrst("TIPO_CHAVE")
            Case 10  'Desconto Comercial
                If Val(lrst("CHAVE_1")) = pCorretor Then
                    If Val(lrst("CHAVE_2")) = Val(pClassif) Then
                        If Val(lrst("CHAVE_3")) = 1 Then
                            If Val(lrst("CHAVE_4")) < pValLMG Then
                                pDesconto = lrst("RETORNO")
                                Exit Do
                            End If
                        End If
                    End If
                End If
        'Ficha 1000408
            Case 13 'Desconto Comercial / Segurado
                If Val(lrst("CHAVE_1")) = pCorretor Then
                    If Val(lrst("CHAVE_2")) = pSegurado Then
                        pDesconto = lrst("RETORNO")
                        Exit Do
                    End If
                End If
        'Ficha 1000851 - Tarifa Setembro 2010
            Case 14  'Desconto Comercial / Corretor / Classificação
                If Val(lrst("CHAVE_1")) = pCorretor Then
                    If Val(lrst("CHAVE_2")) = Val(pClassif) Then
                        pDesconto = lrst("RETORNO")
                        Exit Do
                    End If
                End If
        End Select
        lrst.MoveNext
    Loop
    Call gpFecha3(lrst, bdP00MULTI)
    gfBuscaDescontoParm = True
End Function

Public Function gfGravarInspecaoProduto(pobjstcPedido As stcA46V702B, pobjstcPedLoc As stcA46V708B, pMensagem As String, Optional pConfirmada As Integer = 0) As Boolean
Dim bdBasDados              As ADODB.Connection 'Base de dados.
Dim lobjstcTransInsp        As stcA46V730B 'Objeto da classe de estrutura stcA46V730B (TAB_TRANS_INSP)
Dim lobjclsTransInsp        As clsA46V730A 'Objeto da classe de estrutura stcA46V730A (TAB_TRANS_INSP)
Dim lobjstcTransCobInsp     As stcA46V731B 'Objeto da classe de estrutura stcA46V731B (TAB_TRANS_INSP)
Dim lobjclsTransCobInsp     As clsA46V731A 'Objeto da classe de estrutura stcA46V731A (TAB_TRANS_INSP)
Dim lobjstcPedCobert        As stcA46V704B 'Objeto da classe de coberturas stcA46V704B (TAB_COBERT)
Dim lMensagem               As String

Err.Clear
On Error GoTo TrataErro: gfGravarInspecaoProduto = False
    '1. Grava registro de solicitação de agendamento da inspeção (Produto).
                
    Set lobjstcTransInsp = New stcA46V730B
    Set lobjstcTransCobInsp = New stcA46V731B
        
    With lobjstcTransInsp
        .Num_Item = pobjstcPedLoc.Num_Item
        .Num_Ped = pobjstcPedido.Num_Ped
        .Num_Apol_Ant = pobjstcPedLoc.Num_Apol_Anterior
        .Nom_Segurado = pobjstcPedido.Nom_Segurado
        .Num_Tel = pobjstcPedLoc.objstcPedInsp.Num_Tel
        .Num_Cgc_Cpf = pobjstcPedido.Num_Cgc_Cpf
        .Cod_Ramo = pobjstcPedido.Cod_Ramo
        .Cod_Org = pobjstcPedido.Cod_Org
        .Dat_Ini_Vig = pobjstcPedido.Dat_Ini_Vig
        .Dat_Ter_Vig = pobjstcPedido.Dat_Ter_Vig
        .Cod_Tip_Logr_Risco = pobjstcPedLoc.Cod_Tip_Logr_Risco
        .Nom_Logr_Risco = pobjstcPedLoc.Nom_Logr_Risco
        .Dsc_Compl_Risco = pobjstcPedLoc.Dsc_Compl_Risco
        .Nom_Bairro_Risco = pobjstcPedLoc.Nom_Bairro_Risco
        .Nom_Cid_Risco = pobjstcPedLoc.Nom_Cid_Risco
        .Sig_UF_Risco = pobjstcPedLoc.Sig_UF_Risco
        .Num_Cep_Risco = pobjstcPedLoc.Num_Cep_Risco
        .Num_Risco = pobjstcPedLoc.objstcPedInsp.Num_Logr_Risco
        .Dat_Trans = Format(gForHoje, "YYYYMMDD")
        .Cod_Iden_Pessoa = pobjstcPedido.Tip_Pessoa
        .Cod_Unid_Prod = pobjstcPedido.Cod_Unid_Prod
        .Cod_Prod = pobjstcPedido.Cod_Prod
        .Cod_Corr = pobjstcPedido.objstcCtrlEmis.Cod_Corr
        .Cod_Clasf = pobjstcPedLoc.Cod_Clasf
        .Cod_Compl_Clasf = pobjstcPedLoc.Cod_Compl_Clasf
        .Pess_Cont = pobjstcPedLoc.objstcPedInsp.Pess_Cont
        .Cod_Produto = pobjstcPedido.Cod_Produto
        .Flag_Trans = 0
        .Num_Cotac = ""
        .Cod_Cobert_01 = 0
        .Val_IS_Cobert_01 = 0
        .Cod_Cobert_02 = 0
        .Val_IS_Cobert_02 = 0
        .Cod_Cobert_03 = 0
        .Val_IS_Cobert_03 = 0
        .Cod_Cobert_04 = 0
        .Val_IS_Cobert_04 = 0
        .Cod_Cobert_05 = 0
        .Val_IS_Cobert_05 = 0
        .Cod_Cobert_06 = 0
        .Val_IS_Cobert_06 = 0
        .Cod_Cobert_07 = 0
        .Val_IS_Cobert_07 = 0
        .Cod_Cobert_08 = 0
        .Val_IS_Cobert_08 = 0
        .Cod_Cobert_09 = 0
        .Val_IS_Cobert_09 = 0
        .Cod_Cobert_10 = 0
        .Val_IS_Cobert_10 = 0
        
Fim:
        .Num_Insp = 0
        .Dat_Final = 0
        .Dat_Cancel = 0
        .Cod_Status_I = pConfirmada
        .Num_Cotac = pobjstcPedido.objstcCtrlEmis.Num_Cotac
        .Dat_Confirm = Format(gForHoje, "YYYYMMDD")
        
    End With
            
    '2. Salvar tabela informações de Inspeção
    Set lobjclsTransInsp = New clsA46V730A
    lobjstcTransInsp.NOSSO_NUMERO = pobjstcPedido.NOSSO_NUMERO
    lobjstcTransCobInsp.Num_Ped = pobjstcPedido.Num_Ped
    lobjstcTransCobInsp.Num_Item = pobjstcPedLoc.Num_Item
    lobjstcTransCobInsp.Num_Cotac = pobjstcPedido.objstcCtrlEmis.Num_Cotac
    
    Set lobjclsTransInsp.ConexaoBD = bdBasDados
    If Not lobjclsTransInsp.gfIncluir(lobjstcTransInsp) Then
        lMensagem = lobjclsTransInsp.UltimoErro
        GoTo TrataErro
    End If
    
    If lobjclsTransInsp.UltimoErro = "DUPLICADO" Then
        gfGravarInspecaoProduto = True
        GoTo TrataErro
    End If
    
    '3. Salvar informações das coberturas
        
    For Each lobjstcPedCobert In pobjstcPedLoc.ColecaoPedCobert
        lobjstcTransCobInsp.Cod_Cobert = lobjstcPedCobert.Cod_Cobert
        lobjstcTransCobInsp.Lmg_Cobert = lobjstcPedCobert.Val_IS

        Set lobjclsTransCobInsp = New clsA46V731A
        lobjstcTransCobInsp.NOSSO_NUMERO = lobjstcPedCobert.NOSSO_NUMERO
        
        Set lobjclsTransCobInsp.ConexaoBD = bdBasDados
        If Not lobjclsTransCobInsp.gfIncluir(lobjstcTransCobInsp) Then
            lMensagem = lobjclsTransCobInsp.UltimoErro
            GoTo TrataErro
        End If
    Next
    
    gfGravarInspecaoProduto = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If gfPreenchido(lMensagem) = False Then
        lMensagem = Err & " - " & Error
    End If
    pMensagem = "gfGravarInspecaoProduto Erro:" & lMensagem
    GoTo SAIDA
    
FechaObj:
    Set bdBasDados = Nothing
    Set lobjstcTransInsp = Nothing
    Set lobjclsTransInsp = Nothing
    Set lobjstcTransCobInsp = Nothing
    Set lobjclsTransCobInsp = Nothing
    Set lobjstcPedCobert = Nothing
    Return
End Function

Public Function gfBuscaJurosParm(ByVal pobjstcItem As stcA46V708B, _
                                 ByVal pQtdParcela As Integer, ByVal pAntec As Integer, _
                                 ByRef pJuros As Single, _
                                 ByRef pMensagem As String) As Boolean '1001224 - Condições Diferenciada - Seguratec


Dim bdP00MULTI  As ADODB.Connection
Dim lrst        As ADODB.Recordset
Dim lSelect     As String

    Select Case pQtdParcela
        Case 1
            If pAntec = 1 Then
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt01
            Else
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos01
            End If
        Case 2
            If pAntec = 1 Then
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt02
            Else
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos02
            End If
        Case 3
            If pAntec = 1 Then
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt03
            Else
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos03
            End If
        Case 4
            If pAntec = 1 Then
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt04
            Else
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos04
            End If
        Case 5
            If pAntec = 1 Then
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt05
            Else
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos05
            End If
        Case 6
            If pAntec = 1 Then
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt06
            Else
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos06
            End If
        Case 7
            If pAntec = 1 Then
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt07
            Else
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos07
            End If
        Case 8
            If pAntec = 1 Then
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt08
            Else
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos08
            End If
        Case 9
            If pAntec = 1 Then
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt09
            Else
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos09
            End If
        Case 10
            If pAntec = 1 Then
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt10
            Else
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos10
            End If
        Case 11
            If pAntec = 1 Then
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt11
            Else
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos11
            End If
        Case 12
            If pAntec = 1 Then
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt12
            Else
                pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos12
            End If
    End Select
        
    'Para planos o Juros é Zero
    If pobjstcItem.ObjPai.getPlano.EhPlaninho Then
        pJuros = 0
    End If
        
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdP00MULTI, _
            pMensagem) = False Then GoTo SAIDA
    
    lSelect = "Select * from TAB_PARM_DIVERSOS"
    lSelect = lSelect & " Where Cod_Produto = " & pobjstcItem.ObjPai.Cod_Produto
    lSelect = lSelect & " and Cod_Protocolo = " & pobjstcItem.ObjPai.Cod_Protocolo
    lSelect = lSelect & " and Cod_Sequencia <= " & pobjstcItem.ObjPai.Cod_Sequencia
    lSelect = lSelect & " and cod_ramo = " & pobjstcItem.ObjPai.Cod_Ramo
    lSelect = lSelect & " and TIPO_VALOR = 7"
    lSelect = lSelect & " order by NIVEL_PRIORIDADE"
    If gfObtRegistro(bdP00MULTI, lSelect, lrst, _
            pMensagem) = False Then GoTo SAIDA
    
    Do While Not lrst.EOF
        Select Case lrst("TIPO_CHAVE")
            Case 0  'Tarifa básica
                pJuros = lrst("RETORNO")
            Case 1  'Unidade
                If Val(lrst("CHAVE_1")) = pobjstcItem.ObjPai.Cod_Unid_Prod Then   '1001224 - Condições Diferenciada - Seguratec
                    pJuros = lrst("RETORNO")
                End If
            Case 2  'Corretor
                If Val(lrst("CHAVE_1")) = pobjstcItem.ObjPai.objstcCtrlEmis.Cod_Corr Then   '1001224 - Condições Diferenciada - Seguratec
                    pJuros = lrst("RETORNO")
                End If
            Case 3  'Corretor/Parcela/Pagamento Antecipado?
                If Val(lrst("CHAVE_1")) = pobjstcItem.ObjPai.objstcCtrlEmis.Cod_Corr Then   '1001224 - Condições Diferenciada - Seguratec
                    If pQtdParcela <= Val(lrst("CHAVE_2")) Then
                        If Val(lrst("CHAVE_3")) = pAntec Then
                            pJuros = lrst("RETORNO")
                        End If
                    End If
                End If
            Case 8  'Modalidade/Corretor/Parcela/Pagamento Antecipado?
                If Val(lrst("CHAVE_1")) = pobjstcItem.Cod_Plano Then
                    If Val(lrst("CHAVE_2")) = pobjstcItem.ObjPai.objstcCtrlEmis.Cod_Corr Then   '1001224 - Condições Diferenciada - Seguratec
                        If pQtdParcela <= Val(lrst("CHAVE_3")) Then
                            If Val(lrst("CHAVE_4")) = pAntec Then
                                pJuros = lrst("RETORNO")
                            End If
                        End If
                    End If
                End If
            Case 9  'Modalidade/Parcela/Pagamento Antecipado?
                If Val(lrst("CHAVE_1")) = pobjstcItem.Cod_Plano Then
                    If pQtdParcela <= Val(lrst("CHAVE_2")) Then
                        If Val(lrst("CHAVE_3")) = pAntec Then
                            pJuros = lrst("RETORNO")
                        End If
                    End If
                End If
            Case 12 'Corretor/Rubrica/Parcela/Pagamento Antecipado?
                If Val(lrst("CHAVE_1")) = pobjstcItem.ObjPai.objstcCtrlEmis.Cod_Corr Then   '1001224 - Condições Diferenciada - Seguratec
                    If Val(lrst("CHAVE_2")) = Val(pobjstcItem.Cod_Clasf) & Val(pobjstcItem.Cod_Compl_Clasf) Then
                        If pQtdParcela <= Val(lrst("CHAVE_3")) Then
                            If Val(lrst("CHAVE_4")) = pAntec Then
                                pJuros = lrst("RETORNO")
                            End If
                        End If
                    End If
                End If
            '1001224 - Condições Diferenciada - Seguratec
            Case 17 'Corretor / Cod Segurado / Numero parcela/ Pagto antecipado?
                If Val(lrst("CHAVE_1")) = pobjstcItem.ObjPai.objstcCtrlEmis.Cod_Corr Then
                    If Val(lrst("CHAVE_2")) = pobjstcItem.ObjPai.Cod_Segurado Then
                        If pQtdParcela <= Val(lrst("CHAVE_3")) Then
                            If Val(lrst("CHAVE_4")) = pAntec Then
                                pJuros = lrst("RETORNO")
                            End If
                        End If
                    End If
                End If
            '1100237 - Tarifa 1008 e 1009 Abril 2011
            Case 18 'Cod Classificacao / Numero parcela/ Pagto antecipado?
                If Trim(lrst("CHAVE_1")) = Val(pobjstcItem.Cod_Clasf) Then
                    If pQtdParcela <= Val(lrst("CHAVE_2")) Then
                        If Val(lrst("CHAVE_3")) = pAntec Then
                            pJuros = lrst("RETORNO")
                        End If
                    End If
                End If
            '1100864 - Tarifa 1018 e 1019 Setembro 2011
            Case 20 'Classificacao + Complemento/ Numero parcela/ Pagto antecipado?
                If Trim(lrst("CHAVE_1")) = Val(pobjstcItem.Cod_Clasf) & Val(pobjstcItem.Cod_Compl_Clasf) Then
                    If pQtdParcela <= Val(lrst("CHAVE_2")) Then
                        If Val(lrst("CHAVE_3")) = pAntec Then
                            pJuros = lrst("RETORNO")
                        End If
                    End If
                End If
            '1200429 - Tarifa 1036 e 1037 Maio 2012 - Item 5
            Case 21 'Classificacao / Regiao / Numero parcela / Pagamento antecipado?
                If Trim(lrst("CHAVE_1")) = Trim(pobjstcItem.Cod_Clasf) Then
                    If Val(lrst("CHAVE_2")) = pobjstcItem.gCod_RegiaoNova Then
                        If pQtdParcela <= Val(lrst("CHAVE_3")) Then
                            If Val(lrst("CHAVE_4")) = pAntec Then
                                pJuros = lrst("RETORNO")
                            End If
                        End If
                    End If
                End If
        End Select
        lrst.MoveNext
    Loop
    Call gpFecha3(lrst, bdP00MULTI)
    
    '1101015 - Projeto 11010 - Renovação Expressa
    If (pobjstcItem.Tip_Emissao = e_TipEmissao_RenYasuda) Then
        If (pobjstcItem.Cod_Iden_Tarif >= 1026) And (pobjstcItem.Apol_Ant_IND_SINISTRO = 0) Then
            If (pQtdParcela <= pobjstcItem.ObjPai.Apol_Ant_QTD_PARC) And (pobjstcItem.ObjPai.Apol_Ant_PER_JURO_FRAC < pJuros) Then
                pJuros = pobjstcItem.ObjPai.Apol_Ant_PER_JURO_FRAC
            End If
        End If
    End If
     
Fim:
    gfBuscaJurosParm = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set bdP00MULTI = Nothing
    Set lrst = Nothing
    Return
    
End Function



