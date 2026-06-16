VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsM46V111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================================================================================
' Class Name: clsM46V111
' Sistema: SyasMult
' Analista: Marcio Oliveira
' Objetivo: Substitui o antigo m�dulo bas Adaptado para ser utilizado como classe gerenciavel.
' ======================================================================================================================================================
Public gRegPadrao As Boolean    'Indicador de registro padr�o.
Public gComNosNumero As String    'Registro de comunica��o - nosso n�mero.
Public gNumPedido As String    'N�mero do pedido.
'Public gNum_Pi As Long
Public gNum_Pi As Double
Public gSaida As Boolean    'Indica que trata-se de sa�da.
Public gProJurIOF As Double     'Produto: juros de IOF.

Private privCod_plano As Integer


'Public M46V999.gQuaErros        As Byte             'Quantidade de erros.
Public gCod_Dizer As Integer
Public gCod_DizerTexto As String

Public gModalidade As String  'Nome da Modalidade
Public gTipoGrupMulti As e_GrupoRamo    'Define Ramos que fazem parte de MultiSeguros (112,113,114)
Public gCod_Regiao As Integer    'C�digo da Regiao
Public gDesc_Coml As Double    'Desconto Comercial para os Planinhos
Public gDesc_Prog As Double    'Desconto Progressivo para os Planinhos
Public gDesc_Bonus As Double    'Desconto B�nus para os Planinhos

Private Declare Sub PontSyas Lib "D046O001.dll" (Reg_Parm As ty_Parm)
'Utilizado para executar pausa no codigo do programa
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type ty_Parm
    Cod_User As String * 6
    Cod_Ramo As Integer
'    NUM_PI As Long
    NUM_PI As Double
    Num_Ped As Long
    Num_Item As Long
    Num_Apol_Ant As Long 'N�o mudar para Double sen�o dar� erro nas tarifas < 1080 rSilva
    Num_Item_Ant As Long
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

Private Declare Function WaitForSingleObject Lib "kernel32" _
                                             (ByVal hHandle As Long, _
                                              ByVal dwMilliseconds As Long) As Long

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

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadId As Long
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
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long


Public Enum e_ColFlxSelecionadas
    ColCodCobert = 0
    ColNomCobert = 1
    ColValLMG = 2
    ColLimiteMaximo = 3
    ColValRisco = 4
    ColTabCobert = 5
    ColPerFranq = 6
    ColTipFranq = 7
    ColValFranq = 8
    ColDescComl = 9
    ColDescExp = 10
    ColCobObrigat = 11
    ColPerLimite = 12
    ColPermitePInd = 13
    ColMultiploFranquia = 14
    ColDesconto = 15
    ColTaxaFinal = 16
    ColTaxaAlterada = 17
    ColValFranMin = 18
    ColCobertPerLimite = 19
    ColValISAte = 20
    ColAjusteEndosso = 21
    ColAlterouFranquia = 22
    ColDescFranquia = 23
    ColTextoFranquia = 24
    ColCodGrupoCobert = 25
    ColCoefFranq = 26
    ColPInd = 27
    ColValISmin = 28
End Enum

Private mClsM46V999 As clsM46V999
Private mClsM99V001 As clsM99V001
Private mClsM46V103 As clsM46V103

Public Property Get Class_Name() As String
    Class_Name = "M46V111"
End Property

Public Property Get gFonteEspecificacao() As String
    gFonteEspecificacao = "Arial"
End Property

Public Property Let gCod_plano(pCod_Plano As Integer)
    privCod_plano = pCod_Plano
End Property

Public Property Get gCod_plano() As Integer
    gCod_plano = privCod_plano
End Property


Public Function gfVerSegSiscota(pobjstcPed As stcA46V702B) As Boolean
    Dim lSelect As String
    Dim lBDSiscota As ADODB.Connection
    Dim lrst As clsYasRecordSet
    Dim lMensagem As String
    Dim lCodSegurado As Long
    Dim lCpfCnpj As String

    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = "2" Then
        If M46V999.gfPreenchido(pobjstcPed.objstcCtrlEmis.Num_Cotac) Then
            '3. Abre arquivo.
            If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lBDSiscota, lMensagem) = False Then
                Call M46V999.gpGraLog(2, lMensagem)
                Exit Function
            End If

            lSelect = "select * from siscota.siscota.tab_cotacao WITH (NOLOCK) where ccotacao= " & M46V999.gfFormatarCamposSQL(pobjstcPed.objstcCtrlEmis.Num_Cotac)

            If M46V999.gfObtRegistro(lBDSiscota, lSelect, lrst, lMensagem) = False Then
                Call M46V999.gpFecha2(lBDSiscota)
                Call M46V999.gpGraLog(2, lMensagem)
                Exit Function
            End If

            If Not lrst.EOF Then
                lCodSegurado = lrst("CSegurado")
                lCpfCnpj = lrst("Numdoctoseg")
                If lCodSegurado = 0 Then
                    If pobjstcPed.Tip_Pessoa = e_TipPessoa_Fisica Then
                        lSelect = "select * from siscota.siscota.tab_segurado WITH (NOLOCK) where CPF = '" & Format(pobjstcPed.Num_Cgc_Cpf, "00000000000") & "'"
                    Else
                        If IsNumeric(pobjstcPed.Num_Cgc_Cpf) Then
                            lSelect = "select * from siscota.siscota.tab_segurado WITH (NOLOCK) where CGC = '" & Format(pobjstcPed.Num_Cgc_Cpf, "00000000000000") & "'"
                        Else
                            lSelect = "select * from siscota.siscota.tab_segurado WITH (NOLOCK) where CGC = '" & Right$(String(14, "0") & Trim(pobjstcPed.Num_Cgc_Cpf), 14) & "'"
                        End If
                    End If

                    If M46V999.gfObtRegistro(lBDSiscota, lSelect, lrst, lMensagem) = False Then
                        Call M46V999.gpFecha2(lBDSiscota)
                        Call M46V999.gpGraLog(2, lMensagem)
                        Exit Function
                    End If
                    If lrst.EOF Then
                        lSelect = "select max(csegurado) as codigo from siscota.siscota.tab_segurado WITH (NOLOCK)"

                        If M46V999.gfObtRegistro(lBDSiscota, lSelect, lrst, lMensagem) = False Then
                            Call M46V999.gpFecha2(lBDSiscota)
                            Call M46V999.gpGraLog(2, lMensagem)
                            Exit Function
                        End If

                        lCodSegurado = lrst("Codigo") + 1

                        lSelect = " INSERT INTO [SISCOTA].[SISCOTA].[TAB_SEGURADO]([CSEGURADO], " & _
                                " [SEGURADO], [ASEGURADO], [XPESSOA], [CGC], [CPF], [CUSER], [DSYS], " & _
                                " [DATUAL], [SREG], [ABREV_SISCOTA], [HSYS], [CUSER_EXC], [DEXC], [HEXC]) VALUES( "

                        lSelect = lSelect & lCodSegurado & "," & _
                                  M46V999.gfFormatarCamposSQL(Left(pobjstcPed.Nom_Segurado, 70)) & "," & _
                                  M46V999.gfFormatarCamposSQL(Left(pobjstcPed.Nom_Segurado, 15)) & ","
                        If pobjstcPed.Tip_Pessoa = e_TipPessoa_Fisica Then
                            lSelect = lSelect & "'F',''," & pobjstcPed.Num_Cgc_Cpf & ","
                        Else
                            lSelect = lSelect & "'J'," & pobjstcPed.Num_Cgc_Cpf & ",'',"
                        End If
                        lSelect = lSelect & "'SYAS',"
                        lSelect = lSelect & Left(M46V999.gfForData(3), 8) & ","
                        lSelect = lSelect & "0,'A','M',"
                        lSelect = lSelect & Right(M46V999.gfForData(3), 6) & ", NULL, NULL, NULL )"

                        If M46V999.gfExeSQL(lBDSiscota, lSelect, lMensagem) <> 0 Then
                            Call M46V999.gpGraLog(1, lMensagem)
                            Exit Function
                        End If
                    Else
                        lCodSegurado = lrst("CSegurado")
                    End If
                End If

                lSelect = " Update siscota.siscota.tab_cotacao set csegurado = " & lCodSegurado
                lSelect = lSelect & " , Numdoctoseg = " & M46V999.gfFormatarCamposSQL(pobjstcPed.Num_Cgc_Cpf)
                lSelect = lSelect & " where ccotacao = " & M46V999.gfFormatarCamposSQL(pobjstcPed.objstcCtrlEmis.Num_Cotac)

                If M46V999.gfExeSQL(lBDSiscota, lSelect, lMensagem) <> 0 Then
                    Call M46V999.gpGraLog(1, lMensagem)
                    Exit Function
                End If
            End If
        End If
    End If

    Set lBDSiscota = Nothing
    Set lrst = Nothing
End Function
Public Function gfExecDLL(ByRef pobjstcPed As stcA46V702B, ByRef pobjstcPedLoc As stcA46V708B, _
                          ByVal pTipoProc As e_TipProc, ByVal pItem As Long, _
                          ByRef pCod_Retorno As Integer, pTipHomologacao As String) As String
' 1 - para consistir comum
' 2 - para consistir o item
' 3 - para normas
    Dim lParm As ty_Parm
    Dim lretorno As String
    Dim lNomeArq As String
    Dim lMensagem As String

    On Error GoTo VerErro
    
    If M46V999.gCodProduto >= 1080 Then
        gfExecDLL = Empty
        Exit Function
    End If
    
    'vou criar um arquivo somente para validar se o usu�rio � �nico no processamento.
    If mpVerCalculoParalelo() Then
        Call M46V999.gpGraLog(1, lMensagem)
        'Exit Function
    End If

    lParm.Cod_Ramo = pobjstcPed.Cod_Ramo
    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        lParm.Cod_User = "SYAS98"
    Else
        lParm.Cod_User = "SYAS99"
    End If
    lParm.NUM_PI = gNum_Pi

    lParm.Num_Ped = pobjstcPed.Num_Ped

    lParm.Num_Item = pItem
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Or pobjstcPedLoc.Tip_Emissao = e_TipEmissao_Endosso Then
        'Acerto
        'lParm.Num_Apol_Ant = objstcPedido.Num_Apol_An
        lParm.Num_Apol_Ant = pobjstcPedLoc.Num_Apol_Anterior
        lParm.Num_Item_Ant = val(pItem)
    Else
        lParm.Num_Apol_Ant = 0
        lParm.Num_Item_Ant = 0
    End If
    lParm.flag_debug = 1
    lParm.flag_debug = 0
    lParm.Tip_Proc = pTipoProc
    'se pTipoProc = 3, = Vernormas - Passar Tarifa no n�mero do PI
    If pTipoProc = 3 Then
        lParm.NUM_PI = pobjstcPed.Cod_Produto
    End If

    'lParm.Cod_Dizer = pobjstcPed.Cod_Dizer
    lParm.Cod_Dizer = pobjstcPedLoc.Cod_DizerItem

    If lParm.Cod_Dizer = e_CodDizerInclusao Then
        lParm.Num_Item_Ant = 0
    End If

    If pTipHomologacao = "1" Then
        lParm.Versao_Corr = 2   'vers�o homologa��o
    Else
        lParm.Versao_Corr = 1
    End If

    'VbParm.
    ''  GoTo Thais

    Dim Tentativa As Integer
    Tentativa = 0

TentadeNovo:

    Call PontSyas(lParm)

    If Trim(lParm.Mens_Err_1) = "N�o � poss�vel abrir mais tabelas." Then
        If Tentativa = 0 Then
            Tentativa = 1
            GoTo TentadeNovo
        End If
    End If


    '*  ret-cod               10 - Erro OPEN OUTPUT VSAM              *
    '*                        11 - Erro Delete Vsam                   *
    '*                        12 - Erro Write TMENU                   *
    '*                        13 - Erro Write T01O002A                *
    '*                        14 - Erro no Open I-O                   *
    '*                        15 - Quantidade de Registros inv�lida   *
    '*                        16 - Erro na abertura do banco          *
    '*                        17 - Erro no diretorio corrente         *
    '*                        18 - Erro no Join do nome do arquivo    *
    '*                        19 - Erro no Read VSAM                  *
    '*                        40 - Problemas nos dados de entrada     *
    '*                        50 - Erros na tela P01O002              *
    '*                        52 - Erros na tela P01O003              *
    '*                        53 - Erros na tela P01O004              *
    '*                        54 - Erros na tela P01O062              *
    '*                        55 - Aceita��o depende de inspe��o      *
    '*                        80 - Proposta nao aceita                *

    lretorno = lParm.Cod_Retorno & " -- " & Trim(lParm.Cod_Err_Nor_1) & " - " & Trim(lParm.Mens_Err_1) & _
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
    lretorno = Trim(Replace(lretorno, "- 0 -", ""))
    lretorno = Trim(Replace(lretorno, Chr(0), ""))

    If Trim(Replace(lretorno, Chr(32), "")) = "0-" Then
Thais:
        gfExecDLL = ""
    Else
        gfExecDLL = lretorno
    End If
    pCod_Retorno = lParm.Cod_Retorno

    Call mpDeletarArqCalculo
    Exit Function

VerErro:
    MsgBox "ERRO - gfExecDLL " & Err & ERROR

End Function
Private Sub mpDeletarArqCalculo()
    Dim lDiretorio As String
    Dim lNomeArquivo As String
    Dim lNomeComputador As String

    On Error Resume Next
    lDiretorio = Replace(UCase(M46V999.gEndP00MULTI), "P00MULTI.MDB", "")
    lNomeComputador = M99V001.GetLoginComputerName()
    lNomeArquivo = lNomeComputador & ".PROCESSANDO"
    Kill lDiretorio & lNomeArquivo
    Err.Clear
    On Error GoTo 0
End Sub

Private Function mpVerCalculoParalelo() As Boolean
    Dim lDiretorio As String
    Dim lOutroArq As String
    Dim lArquivo As Integer
    Dim lNomeArquivo As String
    Dim lNomeComputador As String
    On Error GoTo ControleErro
    lDiretorio = Replace(UCase(M46V999.gEndP00MULTI), "P00MULTI.MDB", "")
    lNomeComputador = M99V001.GetLoginComputerName()
    lNomeArquivo = lNomeComputador & ".PROCESSANDO"
Retorno:
    lOutroArq = Dir(lDiretorio & "*.PROCESSANDO")
    If M46V999.gfPreenchido(lOutroArq) Then
        If FileDateTime(lDiretorio & lOutroArq) < (DateAdd("s", -20, Now)) Then
            Kill lDiretorio & lOutroArq
            GoTo Retorno
        Else
            'If lNomeMaq <> Left(lOutroArq, InStr(1, lOutroArq, ".") - 1) Then
            'Call M46V999.gpMenErro("Existe outra esta��o efetuando um c�lculo simult�neo.(" & lNomeComputador & "," & Left(lOutroArq, InStr(1, lOutroArq, ".") - 1) & ")")
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
    Call M46V999.gpGraLog(0, "mpVerCalculoParalelo - " & Err.Description)
End Function

Public Sub gpObterRegComunicacao(ByVal pExibeMensagem As Boolean, _
                                 ByVal pPropTabela As String, _
                                 ByVal pNomeTabImport As String)

'********************Vari�veis padr�o*************************************
    Dim lCkPoint As String    'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lCkStr As String    'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
'lCkPoint. (posi��o das variaveis em ordem de declara��o)
'********************Vari�veis Declaradas para o processamento************
    Dim lObjPedCom As clsA46V709A  'Objeto da classe clsA46V709A (PedCom)
    Dim lObjstcPedCom As stcA46V709B     'Objeto da classe de estrutura stcA46V709B (PedCom)

    Err = 0
    lCkPoint = ""
    lCkStr = ""
    'Procedure: Obtem o registro de comunica��o.

    'Instancio os objetos.
    On Error GoTo ControleErro
    lCkPoint = "1"
    Set lObjPedCom = New clsA46V709A
    Set lObjPedCom.M46V999 = M46V999

    lCkPoint = "2"
    'Busco o registro de comunica��o.
    If Not lObjPedCom.mfCarregarColecao(pPropTabela, pNomeTabImport) Then
        Call M46V999.gpGraLog(1, lObjPedCom.UltimoErro)
        GoTo SAIDA
    End If
    lCkPoint = "3"
    If lObjPedCom.Colecao.Count = 0 Then
        If pExibeMensagem = False Then
            GoTo SAIDA
        Else
            Call M46V999.gpGraLog(2, "FTE2003 - Registro de comunica��o n�o localizado.")
            GoTo SAIDA
        End If
    End If
    lCkPoint = "4"
    Set lObjstcPedCom = lObjPedCom.Colecao.Item(1)
    lCkPoint = "5"
    gRegPadrao = False
    If Not M46V999.gfPreenchido(lObjstcPedCom.nosso_numero) Then
        Call M46V999.gpGraLog(2, "FTE2003 - Registro de comunica��o inv�lido (nosso n�mero n�o " & _
                                 "preenchido)." & vbLf & vbLf & "Programa ser� encerrado.")
        GoTo SAIDA
    End If
    lCkPoint = "6"
    gComNosNumero = Trim$(lObjstcPedCom.nosso_numero)
    Select Case Len(gComNosNumero)
    Case Is <= 6
        Call M46V999.gpGraLog(2, "FTE2003 - Registro de comunica��o inv�lido (nosso n�mero com tamanho " & _
                                 "de " & Len(gComNosNumero) & " posi��es)." & vbLf & vbLf & _
                                 "Programa ser� encerrado.")
        GoTo SAIDA
    Case 20
        If Not IsNumeric(gComNosNumero) Then
            Call M46V999.gpGraLog(2, "FTE2003 - Registro de comunica��o inv�lido (nosso n�mero n�o num�rico)." & _
                                     vbLf & vbLf & "Programa ser� encerrado.")
            GoTo SAIDA
        End If
        If CDbl(gComNosNumero) = 0 Then
            Call M46V999.gpGraLog(2, "FTE2003 - Registro de comunica��o inv�lido (nosso n�mero zerado)." & _
                                     vbLf & vbLf & "Programa ser� encerrado.")
            GoTo SAIDA
        End If
    Case Else
        If Left$(gComNosNumero, 6) <> "PADRAO" Then
            Call M46V999.gpGraLog(2, "FTE2003 - Registro de comunica��o inv�lido (nosso n�mero n�o e " & _
                                     "diferente de PADR�O)." & vbLf & vbLf & "Programa ser� encerrado.")
            GoTo SAIDA
        End If
        gRegPadrao = True
    End Select
    If M46V999.PerfilAcesso = e_Perfil_Triagem Then
        gNum_Pi = lObjstcPedCom.Dat_Alt
    End If

    lCkPoint = "7"
    'Excluo o mesmo registro da base de dados.
    If Not lObjPedCom.mfExcluir(pPropTabela, pNomeTabImport) Then
        Call M46V999.gpGraLog(2, lObjPedCom.UltimoErro)
        GoTo SAIDA
    End If
    lCkPoint = "8"

SAIDA:
    GoSub FechaObj
    Exit Sub
ControleErro:
    Call M46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call M46V999.gpGraLog(0, "gpObterRegComunicacao ERRO - " & Err & " " & Err.Description & _
                             Chr(13) & "M46V999.CheckPoint " & lCkPoint)
    GoTo SAIDA
FechaObj:
    Set lObjPedCom = Nothing
    Set lObjstcPedCom = Nothing
    Return
End Sub


Public Function gfGraRegComunicacao(ByVal pNosNumero As String, _
                                    ByRef pMensagem As String, _
                                    Optional pBasDados As ADODB.Connection) As Boolean
'Fun��o: grava registro de comunica��o entre P46V001 e P46V100.

'Par�metros de entrada...pBasDados........Base de dados.
'                        pNosNumero...Nosso n�mero.
'Par�metro de sa�da......pMensagem....Mensagem.

    Dim lObjPedCom As clsA46V709A    'Objeto da classe clsA46V709A (TAB_PED_COM)
    Dim lObjstcPedCom As stcA46V709B    'Objeto da classe de estrutura stcA46V709B (TAB_PED_COM)

    Dim lPropTabela As String
    Dim lNomeTabImport As String

    On Error GoTo ControleErro
    gfGraRegComunicacao = False

    Call M46V999.gfObtSYASMULTINI(lPropTabela)
    lNomeTabImport = ""

    '1. Consiste par�metros.
    'a) Nosso n�mero.
    If Not M46V999.gfPreenchido(pNosNumero) Then
        pMensagem = "FTE0074 - Erro ao criar registro de comunica��o (nosso n�mero n�o preenchido)."
        GoTo SAIDA
    End If
    pNosNumero = UCase$(Trim$(pNosNumero))
    If IsNumeric(pNosNumero) Then
        If val(pNosNumero) = 0 Then
            pMensagem = "FTE0074 - Erro ao criar registro de comunica��o (nosso n�mero zerado)."
            GoTo SAIDA
        End If
        If Len(pNosNumero) > 20 Then
            pMensagem = "FTE0074 - Erro ao criar registro de comunica��o (nosso n�mero com mais de " & _
                        "de 20 posi��es)."
            GoTo SAIDA
        End If
        pNosNumero = String(20 - Len(pNosNumero), "0") & pNosNumero
    Else
        If Len(pNosNumero) <> 6 Then
            pMensagem = "FTE0074 - Erro ao criar registro de comunica��o (nosso n�mero n�o num�rico " & _
                        "e tamanho diferente de 6 posi��es)."
            GoTo SAIDA
        End If
        pMensagem = "FTE0074 - Erro ao criar registro de comunica��o (nosso n�mero n�o num�rico."
    End If

    '2. Instancio as classes
    Set lObjPedCom = New clsA46V709A
    Set lObjPedCom.M46V999 = M46V999

    Set lObjstcPedCom = New stcA46V709B
    Set lObjPedCom.ConexaoBD = pBasDados

    If Not lObjPedCom.mfExcluir(lPropTabela, lNomeTabImport) Then
        pMensagem = lObjPedCom.UltimoErro
        GoTo ControleErro
    End If

    Set lObjstcPedCom = New stcA46V709B
    With lObjstcPedCom
        If pNosNumero = "PADRAO" Then
            .nosso_numero = pNosNumero & M46V999.gNomUsuario6
        Else
            .nosso_numero = pNosNumero
        End If
        .Cod_User = M46V999.gNomUsuario6
        .Dat_Alt = M46V999.gfForData(3)
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
    Call M46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call M46V999.gpGraLog(0, "gfGraRegComunicacao ERRO - " & Err & " " & Err.Description & " " & pMensagem)
    GoTo SAIDA
FechaObj:
    Set lObjPedCom = Nothing
    Set lObjstcPedCom = Nothing
    Return
End Function

Public Sub gpPausa(ByVal lSegundos)

    Dim i
    Dim lMiliSegundos As Double

    lMiliSegundos = lSegundos * 1000

    i = DoEvents
    Sleep (lMiliSegundos)
    i = DoEvents

    Exit Sub

End Sub


Public Sub gpPausa_wf(ByVal lSegundos)
'Pausa a execu��o do codigo durante o total de segundos
'recebido no parametro (Libera o uso da CPU durante a pausa)

    Dim lMiliSegundos As Double
    Dim lDoEvents

    lMiliSegundos = lSegundos * 1000

    lDoEvents = DoEvents
    Sleep (lMiliSegundos)
    lDoEvents = DoEvents

    Exit Sub

End Sub

Public Sub gpObtModalidade(ByVal pCod_Plano As Integer, ByVal pCodRamo As Long)
'Procedure: obt�m produto e informa��es dependentes de produto.

    Dim lbdTAB_Plano As ADODB.Connection    'Base de dados - TAB_PROD.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lMensagem As String     'Mensagem.
    Dim lPosicao As Integer   'Utilizado com InStr.
    Dim lrsTAB_Plano As clsYasRecordSet    'Registro - TAB_PROD.
    Dim lCarga As Boolean
    Dim lLocal As Integer
    Dim lTemp As String
    Dim lSelect As String

    '1. Inicializa vari�veis globais - dados do produto.
    gModalidade = " "

    '2. Se produto zerado, encerra processo.
    If pCod_Plano = 0 Or pCodRamo = 0 Then
        GoTo SAIDA
    End If

    '3. Abre arquivo.
    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTAB_Plano, lMensagem) = False Then
        Call M46V999.gpGraLog(2, lMensagem)
        GoTo SAIDA
    End If

    '4. Obt�m produto informado.
    lSelect = "SELECT * FROM TAB_Plano WITH (NOLOCK) WHERE Cod_Ramo = " & pCodRamo
    lSelect = lSelect & " and Cod_Produto = " & pCod_Plano
    lSelect = lSelect & " and Dat_Vig_Ini <= " & M46V999.gfForData(1, Date)
    lSelect = lSelect & " order by Cod_Produto"

    If M46V999.gfObtRegistro(lbdTAB_Plano, lSelect, lrsTAB_Plano, lMensagem) = False Then
        Call M46V999.gpGraLog(2, lMensagem)
        GoTo SAIDA
    End If

    If lrsTAB_Plano.EOF = True Then
        Call M46V999.gpGraLog(1, "FNL0002 - " & M46V999.gPrefixo & "Produto n�o localizado (" & _
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

    '5. Carrega vari�veis globais.
    gModalidade = lTemp

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Call M46V999.gpFecha2(lbdTAB_Plano)
    Set lrsTAB_Plano = Nothing
    Set lbdTAB_Plano = Nothing
    Return
End Sub

Public Function gfCarregaAss24h(pCodTarifa As Integer, _
                                pCodSequencia As Integer, _
                                pCodRamo As e_CodRamo, _
                                pCodModalidade As Integer) As Collection
    Dim lObjAss24h As clsA46V734A
    Dim lStcAss24h As stcA46V734B
    Static lsRetorno As Collection
    Static lsCodTarifa As Integer
    Static lsCodSequencia As Integer
    Static lsCodRamo As e_CodRamo
    Static lsCodModalidade As Integer

    If pCodTarifa <> lsCodTarifa Or _
       pCodSequencia <> lsCodSequencia Or _
       pCodRamo <> lsCodRamo Or _
       pCodModalidade <> lsCodModalidade Then

        lsCodTarifa = pCodTarifa
        lsCodSequencia = pCodSequencia
        lsCodRamo = pCodRamo
        lsCodModalidade = pCodModalidade

        Set lObjAss24h = New clsA46V734A
        Set lObjAss24h.M46V999 = M46V999

        Set lStcAss24h = New stcA46V734B
        Set lStcAss24h.M46V999 = M46V999

        With lStcAss24h
            .Cod_Tarif = pCodTarifa
            .Cod_Sequencia = pCodSequencia
            .Cod_Ramo = pCodRamo
            .Cod_Modalidade = pCodModalidade
        End With

        Set lsRetorno = New Collection
        If lObjAss24h.gfCarregarColecao(lStcAss24h) = True Then
            Set lsRetorno = lObjAss24h.Colecao
        Else
            Call M46V999.gpGraLog(2, lObjAss24h.UltimoErro)
        End If
    End If
    'Retorno
    Set gfCarregaAss24h = lsRetorno
SAIDA:
    GoSub FechaObj
    Exit Function

FechaObj:
    Set lObjAss24h = Nothing
    Set lStcAss24h = Nothing
    Return
End Function

'*******************************************
' M46V999.gpExcluiRegistros agora contida no BAS 999
'*******************************************

'###PROJETO SISCOTA
'DESCOMENTAR AS LINHAS
Public Function gfVerStatusSiscota(pobjstcPed As stcA46V702B, pStatusSiscota As Integer, pRamoSiscota As Long) As Boolean
    Dim lSelect As String
    Dim lBDSiscota As ADODB.Connection
    Dim lrst As clsYasRecordSet
    Dim lMensagem As String

    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = "2" Then
        If M46V999.gfPreenchido(pobjstcPed.objstcCtrlEmis.Num_Cotac) Then
            '3. Abre arquivo.
            If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lBDSiscota, lMensagem) = False Then
                Call M46V999.gpGraLog(2, lMensagem)
                Exit Function
            End If

            lSelect = "select * from siscota.siscota.tab_cotacao WITH (NOLOCK) where ccotacao= " & M46V999.gfFormatarCamposSQL(pobjstcPed.objstcCtrlEmis.Num_Cotac)

            If M46V999.gfObtRegistro(lBDSiscota, lSelect, lrst, lMensagem) = False Then
                Call M46V999.gpFecha2(lBDSiscota)
                Call M46V999.gpGraLog(2, lMensagem)
                Exit Function
            End If

            If Not lrst.EOF Then
                pRamoSiscota = lrst("CRamo")
                pStatusSiscota = lrst("SCotacao")
            End If
        End If
    End If
    Set lBDSiscota = Nothing
    Set lrst = Nothing
End Function

Public Function gfBuscaComissaoParm(pCorretor As Long, pUnidade As Integer, _
                                    pComissao As Single, pValLMG As Long, pClassif As String, _
                                    pMensagem As String, pobjstcPed As stcA46V702B, _
                                    pobjstcPedLoc As stcA46V708B) As Boolean    '1001224 - Condi��es Diferenciada - Seguratec

    Dim bdP00MULTI As ADODB.Connection
    Dim lrst As clsYasRecordSet
    Dim lSelect As String
    '1001224 - Condi��es Diferenciada - Seguratec
    Dim lobjstcPedCobert As stcA46V704B
    Dim lretorno As String

    '1001224 - Condi��es Diferenciada - Seguratec
    If M46V999.gCodProduto >= 1000 Then
        If pobjstcPedLoc.ColecaoPedCobert.Count > 0 Then
            Set lobjstcPedCobert = pobjstcPedLoc.ColecaoPedCobert.Item(1)
        Else
            Set lobjstcPedCobert = Nothing
        End If
        If Not M46V103.gfVerParametro(11, pobjstcPed, pobjstcPedLoc, lobjstcPedCobert, lretorno, pMensagem) Then
            Exit Function
        End If
        pComissao = val(lretorno)
    Else
        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, bdP00MULTI, _
                                 pMensagem) = False Then Exit Function

        lSelect = "Select * from TAB_PARM_DIVERSOS WITH (NOLOCK)"
        lSelect = lSelect & " Where Cod_Produto = " & M46V999.gCodProduto
        lSelect = lSelect & " and Cod_Protocolo = " & M46V999.gCodProtocolo
        lSelect = lSelect & " and Cod_Sequencia <= " & M46V999.gCodSequencia
        If M46V999.gCodProduto >= 998 Then
            lSelect = lSelect & " and TIPO_VALOR = 11"
        Else
            lSelect = lSelect & " and TIPO_VALOR = 10"
        End If
        lSelect = lSelect & " and cod_ramo = " & M46V999.gCod_Ramo
        lSelect = lSelect & " order by NIVEL_PRIORIDADE"

        If M46V999.gfObtRegistro(bdP00MULTI, lSelect, lrst, _
                                 pMensagem) = False Then Exit Function

        Do While Not lrst.EOF
            Select Case lrst("TIPO_CHAVE")
            Case 0  'Tarifa b�sica
                pComissao = lrst("RETORNO")
            Case 1  'Unidade
                If val(lrst("CHAVE_1")) = pUnidade Then
                    pComissao = lrst("RETORNO")
                End If
            Case 2  'Corretor
                If val(lrst("CHAVE_1")) = pCorretor Then
                    pComissao = lrst("RETORNO")
                End If
            Case 11  'Comiss�o
                If val(lrst("CHAVE_1")) = pCorretor Then
                    If val(lrst("CHAVE_2")) = val(pClassif) Then
                        If val(lrst("CHAVE_3")) = 1 Then
                            pComissao = lrst("RETORNO")
                            Exit Do
                        End If
                    End If
                End If

            End Select
            lrst.MoveNext
        Loop
        Call M46V999.gpFecha2(bdP00MULTI)
    End If
    Set lobjstcPedCobert = Nothing
    Set lrst = Nothing
    Set bdP00MULTI = Nothing
    gfBuscaComissaoParm = True
End Function

Public Function gfBuscaDescontoParm(pCorretor As Long, pUnidade As Integer, _
                                    pDesconto As Double, pValLMG As Long, pClassif As String, _
                                    pSegurado As Long, pMensagem As String) As Boolean

    Dim bdP00MULTI As ADODB.Connection
    Dim lrst As clsYasRecordSet
    Dim lSelect As String

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, bdP00MULTI, _
                             pMensagem) = False Then Exit Function

    lSelect = "Select * from TAB_PARM_DIVERSOS WITH (NOLOCK)"
    lSelect = lSelect & " Where Cod_Produto = " & M46V999.gCodProduto
    lSelect = lSelect & " and Cod_Protocolo = " & M46V999.gCodProtocolo
    lSelect = lSelect & " and Cod_Sequencia <= " & M46V999.gCodSequencia
    lSelect = lSelect & " and TIPO_VALOR = 10"
    lSelect = lSelect & " and cod_ramo = " & M46V999.gCod_Ramo
    lSelect = lSelect & " order by NIVEL_PRIORIDADE"

    If M46V999.gfObtRegistro(bdP00MULTI, lSelect, lrst, _
                             pMensagem) = False Then Exit Function

    Do While Not lrst.EOF
        Select Case lrst("TIPO_CHAVE")
        Case 10  'Desconto Comercial
            If val(lrst("CHAVE_1")) = pCorretor Then
                If val(lrst("CHAVE_2")) = val(pClassif) Then
                    If val(lrst("CHAVE_3")) = 1 Then
                        If val(lrst("CHAVE_4")) < pValLMG Then
                            pDesconto = lrst("RETORNO")
                            Exit Do
                        End If
                    End If
                End If
            End If
            'Ficha 1000408
        Case 13    'Desconto Comercial / Segurado
            If val(lrst("CHAVE_1")) = pCorretor Then
                If val(lrst("CHAVE_2")) = pSegurado Then
                    pDesconto = lrst("RETORNO")
                    Exit Do
                End If
            End If
            'Ficha 1000851 - Tarifa Setembro 2010
        Case 14  'Desconto Comercial / Corretor / Classifica��o
            If val(lrst("CHAVE_1")) = pCorretor Then
                If val(lrst("CHAVE_2")) = val(pClassif) Then
                    pDesconto = lrst("RETORNO")
                    Exit Do
                End If
            End If
        End Select
        lrst.MoveNext
    Loop
    Call M46V999.gpFecha2(bdP00MULTI)
    Set lrst = Nothing
    gfBuscaDescontoParm = True
End Function

Public Function gfGravarInspecaoProduto(pobjstcPedido As stcA46V702B, pobjstcPedLoc As stcA46V708B, pMensagem As String, Optional pConfirmada As Integer = 0) As Boolean
    Dim bdBasDados As ADODB.Connection    'Base de dados.
    Dim lobjstcTransInsp As stcA46V730B    'Objeto da classe de estrutura stcA46V730B (TAB_TRANS_INSP)
    Dim lobjclsTransInsp As clsA46V730A    'Objeto da classe de estrutura stcA46V730A (TAB_TRANS_INSP)
    Dim lobjstcTransCobInsp As stcA46V731B    'Objeto da classe de estrutura stcA46V731B (TAB_TRANS_INSP)
    Dim lobjclsTransCobInsp As clsA46V731A    'Objeto da classe de estrutura stcA46V731A (TAB_TRANS_INSP)
    Dim lobjstcPedCobert As stcA46V704B    'Objeto da classe de coberturas stcA46V704B (TAB_COBERT)
    Dim lrsPadrao As clsYasRecordSet
    
    Dim lMensagem As String
    Dim lSelect As String
    Dim lNossoNumero As String

    Err.Clear
    On Error GoTo TrataErro: gfGravarInspecaoProduto = False
    '1. Grava registro de solicita��o de agendamento da inspe��o (Produto).

    Set lobjstcTransInsp = New stcA46V730B
    Set lobjstcTransCobInsp = New stcA46V731B
    lNossoNumero = ""

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
        .Dat_Trans = Format(M46V999.gForHoje, "YYYYMMDD")
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
        .Cod_Cobert_01 = pobjstcPedido.objstcCtrlEmis.cod_origem
        .Val_IS_Cobert_01 = pobjstcPedido.objstcCtrlEmis.Num_Apol_Mar
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
        .Dat_Confirm = Format(M46V999.gForHoje, "YYYYMMDD")

    End With

    '2. Salvar tabela informa��es de Inspe��o
    Set lobjclsTransInsp = New clsA46V730A
    Set lobjclsTransInsp.M46V999 = M46V999

    lobjstcTransInsp.nosso_numero = pobjstcPedido.nosso_numero
    
    'piloto
    'Pegar o nosso n�mero da tab_triagem
    
    If Not M46V999.gfAbrBasDados("RAMOSDIVERSOS", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, bdBasDados, lMensagem) Then
        Call M46V999.gpMenErro(lMensagem)
        GoTo TrataErro
    End If
    
    If M46V999.gPrograma <> "D046O003" Then
    
        If pobjstcPedido.Cod_Tip_Doc = e_TipDoc_Proposta Then
            lSelect = "SELECT  num_siscota "
            lSelect = lSelect & " FROM      ged..Tab_Triagem WITH (NOLOCK)      where "
            lSelect = lSelect & "  num_pi  = " & pobjstcPedLoc.NUM_PI
        
        
            If M46V999.gfObtRegistro(bdBasDados, lSelect, lrsPadrao, lMensagem) Then
                If Not lrsPadrao.EOF Then
                    If Trim(lrsPadrao("num_siscota")) <> "" Then
                        lobjstcTransInsp.nosso_numero = lrsPadrao("num_siscota")
                        GoTo GravarTransInsp
                    Else
                        GoTo BuscaTabPed
                    End If
                End If
            Else
                lMensagem = "Erro ao gravar inspe��o."
                Call M46V999.gpMenErro(lMensagem)
                GoTo TrataErro
            End If

        
BuscaTabPed:
        ' N�o encontrou, tentar ent�o obter o nosso n�mero a partir da TAB_PED
          lSelect = "SELECT   NOSSO_NUMERO_OLD  "
          lSelect = lSelect & "FROM       Tab_Ped_Num_Copia WITH (NOLOCK)     where  "
          lSelect = lSelect & "NOSSO_NUMERO_NEW   =  '" & pobjstcPedido.nosso_numero & "'"
          If M46V999.gfObtRegistro(bdBasDados, lSelect, lrsPadrao, lMensagem) Then
              If Not lrsPadrao.EOF Then
                  lobjstcTransInsp.nosso_numero = lrsPadrao("NOSSO_NUMERO_OLD")
              Else
                  If pobjstcPedido.objstcCtrlEmis.Num_Cotac <> "" And val(pobjstcPedido.objstcCtrlEmis.Num_Cotac) <> "0" Then
                      lobjstcTransInsp.nosso_numero = pobjstcPedido.objstcCtrlEmis.Num_Cotac
                  Else
                      lobjstcTransInsp.nosso_numero = pobjstcPedido.nosso_numero
                  End If
              End If
          Else
              Call M46V999.gpGraLog(0, lMensagem)
              Call M46V999.gpMenErro(lMensagem)
              GoTo TrataErro
          End If
        End If
    End If

GravarTransInsp:
    lobjstcTransCobInsp.Num_Ped = pobjstcPedido.Num_Ped
    lobjstcTransCobInsp.Num_Item = pobjstcPedLoc.Num_Item
    lobjstcTransCobInsp.Num_Cotac = pobjstcPedido.objstcCtrlEmis.Num_Cotac
    lNossoNumero = lobjstcTransInsp.nosso_numero

    Set lobjclsTransInsp.ConexaoBD = bdBasDados
    If Not lobjclsTransInsp.gfIncluir(lobjstcTransInsp) Then
        lMensagem = lobjclsTransInsp.UltimoErro
        GoTo TrataErro
    End If

    If lobjclsTransInsp.UltimoErro = "DUPLICADO" Then
        gfGravarInspecaoProduto = True
        GoTo TrataErro
    End If

    '3. Salvar informa��es das coberturas

    For Each lobjstcPedCobert In pobjstcPedLoc.ColecaoPedCobert
        lobjstcTransCobInsp.Cod_Cobert = lobjstcPedCobert.Cod_Cobert
        lobjstcTransCobInsp.Lmg_Cobert = lobjstcPedCobert.Val_IS

        Set lobjclsTransCobInsp = New clsA46V731A
        Set lobjclsTransCobInsp.M46V999 = M46V999

        lobjstcTransCobInsp.nosso_numero = lNossoNumero

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
    If M46V999.gfPreenchido(lMensagem) = False Then
        lMensagem = Err & " - " & ERROR
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
    Set lrsPadrao = Nothing
    Return
End Function

Public Function gfBuscaJurosParm(ByVal pobjstcItem As stcA46V708B, _
                                 ByVal pQtdParcela As Integer, ByVal pAntec As Integer, ByVal pTip_Compensacao As e_FormaPagamento, _
                                 ByRef pJuros As Single, _
                                 ByRef pMensagem As String) As Boolean    '1001224 - Condi��es Diferenciada - Seguratec


    Dim bdP00MULTI As ADODB.Connection
    Dim lrst As clsYasRecordSet
    Dim lSelect As String
    Dim TipPgtoCarne As Integer
    Dim lPagtoAntecipado As Integer

    If pobjstcItem.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        TipPgtoCarne = 0
    Else
    If pobjstcItem.ObjPai.objstcCtrlEmis.Cod_Bco_Boleto > 0 And pobjstcItem.ObjPai.Cod_Bco_Cobr > 0 And pTip_Compensacao = 2 Then '1 carne com demais em debito
        TipPgtoCarne = 1
    ElseIf pTip_Compensacao = 1 And pobjstcItem.ObjPai.gTipPagamento = e_FormaPagamento_Carne Then 'carne
        TipPgtoCarne = 1
    ElseIf pobjstcItem.ObjPai.objstcCtrlEmis.Cod_Bco_Boleto = 0 And pobjstcItem.ObjPai.Cod_Bco_Cobr > 0 And pTip_Compensacao = 2 Then 'Debito
        TipPgtoCarne = 2
    End If
    End If
    'tbm_juros

    Select Case pQtdParcela
    Case 1
        If pAntec = 1 Then
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt01(pTip_Compensacao)
        Else
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos01(pTip_Compensacao)
        End If
    Case 2
        If pAntec = 1 Then
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt02(pTip_Compensacao)
        Else
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos02(pTip_Compensacao)
        End If
    Case 3
        If pAntec = 1 Then
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt03(pTip_Compensacao)
        Else
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos03(pTip_Compensacao)
        End If
    Case 4
        If pAntec = 1 Then
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt04(pTip_Compensacao)
        Else
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos04(pTip_Compensacao)
        End If
    Case 5
        If pAntec = 1 Then
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt05(pTip_Compensacao)
        Else
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos05(pTip_Compensacao)
        End If
    Case 6
        If pAntec = 1 Then
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt06(pTip_Compensacao)
        Else
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos06(pTip_Compensacao)
        End If
    Case 7
        If pAntec = 1 Then
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt07(pTip_Compensacao)
        Else
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos07(pTip_Compensacao)
        End If
    Case 8
        If pAntec = 1 Then
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt08(pTip_Compensacao)
        Else
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos08(pTip_Compensacao)
        End If
    Case 9
        If pAntec = 1 Then
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt09(pTip_Compensacao)
        Else
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos09(pTip_Compensacao)
        End If
    Case 10
        If pAntec = 1 Then
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt10(pTip_Compensacao)
        Else
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos10(pTip_Compensacao)
        End If
    Case 11
        If pAntec = 1 Then
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt11(pTip_Compensacao)
        Else
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos11(pTip_Compensacao)
        End If
    Case 12
        If pAntec = 1 Then
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinAnt12(pTip_Compensacao)
        Else
            pJuros = pobjstcItem.ObjPai.getProduto.ProJurMinPos12(pTip_Compensacao)
        End If
    End Select

    'Para planos o Juros � Zero
    '11.24 Imobiliario n�o pode zerar os juros
    If pobjstcItem.ObjPai.getPlano.EhPlaninho And pobjstcItem.ObjPai.getPlano.EhPlanoSegmentado = False And Not pobjstcItem.ObjPai.getPlano.ehPlanoImobiliaria Then
        pJuros = 0
    End If

    'Tarifa 11.32
    If pobjstcItem.ObjPai.Cod_Produto >= 1132 And (pobjstcItem.Cod_Plano = "3055" Or pobjstcItem.Cod_Plano = "3057" Or pobjstcItem.Cod_Plano = "3058" Or _
       pobjstcItem.Cod_Plano = "3059") Then
        GoTo Fim:
    End If
    
        'Retirar regras do Segmentando antigo
    GoTo Fim
    
    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, bdP00MULTI, _
                             pMensagem) = False Then GoTo SAIDA

    lSelect = "Select * from TAB_PARM_DIVERSOS WITH (NOLOCK)"
    lSelect = lSelect & " Where Cod_Produto = " & pobjstcItem.ObjPai.Cod_Produto
    lSelect = lSelect & " and Cod_Protocolo = " & pobjstcItem.ObjPai.Cod_Protocolo
    lSelect = lSelect & " and Cod_Sequencia <= " & pobjstcItem.ObjPai.Cod_Sequencia
    lSelect = lSelect & " and cod_ramo = " & pobjstcItem.ObjPai.Cod_Ramo
    lSelect = lSelect & " and TIPO_VALOR = 7"
    lSelect = lSelect & " order by NIVEL_PRIORIDADE"
    If M46V999.gfObtRegistro(bdP00MULTI, lSelect, lrst, _
                             pMensagem) = False Then GoTo SAIDA

    Do While Not lrst.EOF
        Select Case lrst("TIPO_CHAVE")
        Case 0  'Tarifa b�sica
            pJuros = lrst("RETORNO")
        Case 1  'Unidade
            If val(lrst("CHAVE_1")) = pobjstcItem.ObjPai.Cod_Unid_Prod Then   '1001224 - Condi��es Diferenciada - Seguratec
                pJuros = lrst("RETORNO")
            End If
        Case 2  'Corretor
            If val(lrst("CHAVE_1")) = pobjstcItem.ObjPai.objstcCtrlEmis.Cod_Corr Then   '1001224 - Condi��es Diferenciada - Seguratec
                pJuros = lrst("RETORNO")
            End If
        Case 3  'Corretor/Parcela/Pagamento Antecipado?
            If val(lrst("CHAVE_1")) = pobjstcItem.ObjPai.objstcCtrlEmis.Cod_Corr Then   '1001224 - Condi��es Diferenciada - Seguratec
                If pQtdParcela <= val(lrst("CHAVE_2")) Then
                    If val(lrst("CHAVE_3")) = pAntec Then
                        pJuros = lrst("RETORNO")
                    End If
                End If
            End If
        Case 8  'Modalidade/Corretor/Parcela/Pagamento Antecipado?
            If val(lrst("CHAVE_1")) = pobjstcItem.Cod_Plano Then
                If val(lrst("CHAVE_2")) = pobjstcItem.ObjPai.objstcCtrlEmis.Cod_Corr Then   '1001224 - Condi��es Diferenciada - Seguratec
                    If pQtdParcela <= val(lrst("CHAVE_3")) Then
                        If val(lrst("CHAVE_4")) = pAntec Then
                            pJuros = lrst("RETORNO")
                        End If
                    End If
                End If
            End If
        Case 9  'Modalidade/Parcela/Pagamento Antecipado?
            If val(lrst("CHAVE_1")) = pobjstcItem.Cod_Plano Then
                If pQtdParcela <= val(lrst("CHAVE_2")) Then
                    If val(lrst("CHAVE_3")) = pAntec Then
                        pJuros = lrst("RETORNO")
                    End If
                End If
            End If
        Case 12    'Corretor/Rubrica/Parcela/Pagamento Antecipado?
            If val(lrst("CHAVE_1")) = pobjstcItem.ObjPai.objstcCtrlEmis.Cod_Corr Then   '1001224 - Condi��es Diferenciada - Seguratec
                If val(lrst("CHAVE_2")) = val(pobjstcItem.Cod_Clasf) & val(pobjstcItem.Cod_Compl_Clasf) Then
                    If pQtdParcela <= val(lrst("CHAVE_3")) Then
                        If val(lrst("CHAVE_4")) = pAntec Then
                            pJuros = lrst("RETORNO")
                        End If
                    End If
                End If
            End If
            '1001224 - Condi��es Diferenciada - Seguratec
        Case 17    'Corretor / Cod Segurado / Numero parcela/ Pagto antecipado?
            If val(lrst("CHAVE_1")) = pobjstcItem.ObjPai.objstcCtrlEmis.Cod_Corr Then
                If val(lrst("CHAVE_2")) = pobjstcItem.ObjPai.Cod_Segurado Then
                    If pQtdParcela <= val(lrst("CHAVE_3")) Then
                        If val(lrst("CHAVE_4")) = pAntec Then
                            pJuros = lrst("RETORNO")
                        End If
                    End If
                End If
            End If
            '1100237 - Tarifa 1008 e 1009 Abril 2011
        Case 18    'Cod Classificacao / Numero parcela/ Pagto antecipado?
            If Trim(lrst("CHAVE_1")) = val(pobjstcItem.Cod_Clasf) Then
                If pQtdParcela <= val(lrst("CHAVE_2")) Then
                    If val(lrst("CHAVE_3")) = pAntec Then
                        pJuros = lrst("RETORNO")
                    End If
                End If
            End If
            '1100864 - Tarifa 1018 e 1019 Setembro 2011
        Case 20    'Classificacao + Complemento/ Numero parcela/ Pagto antecipado?
            If Trim(lrst("CHAVE_1")) = val(pobjstcItem.Cod_Clasf) & val(pobjstcItem.Cod_Compl_Clasf) Then
                If pQtdParcela <= val(lrst("CHAVE_2")) Then
                    If val(lrst("CHAVE_3")) = pAntec Then
                        pJuros = lrst("RETORNO")
                    End If
                End If
            End If
            '1200429 - Tarifa 1036 e 1037 Maio 2012 - Item 5
        Case 21    'Classificacao / Regiao / Numero parcela / Pagamento antecipado?
            If Trim(lrst("CHAVE_1")) = Trim(pobjstcItem.Cod_Clasf) Then
                If val(lrst("CHAVE_2")) = pobjstcItem.gCod_RegiaoNova Then
                    If pQtdParcela <= val(lrst("CHAVE_3")) Then
                        If val(lrst("CHAVE_4")) = pAntec Then
                            pJuros = lrst("RETORNO")
                        End If
                    End If
                End If
            End If
        Case 22    'Segmentacao/ Numero parcela / Pagamento antecipado / Carn�?
            If CInt(Trim(lrst("CHAVE_1"))) = M46V999.gCod_Segmentacao Then
                If pQtdParcela <= val(lrst("CHAVE_2")) Then
                    If val(lrst("CHAVE_3")) = pAntec Then
                        If val(lrst("CHAVE_4")) = TipPgtoCarne Or TipPgtoCarne = 0 Then
                            pJuros = lrst("RETORNO")
                            Exit Do
                        End If
                    End If

                End If
            End If
        End Select
        lrst.MoveNext
    Loop
    Call M46V999.gpFecha2(bdP00MULTI)

'   Comentado - Nova regra produto yasuda maritima
'    '1101015 - Projeto 11010 - Renova��o Expressa
'    If (pobjstcItem.Tip_Emissao = e_TipEmissao_RenYasuda) Then
'        If (pobjstcItem.Cod_Iden_Tarif >= 1026) And (pobjstcItem.Apol_Ant_IND_SINISTRO = 0) Then
'            If (pQtdParcela <= pobjstcItem.ObjPai.Apol_Ant_QTD_PARC) And (pobjstcItem.ObjPai.Apol_Ant_PER_JURO_FRAC < pJuros) Then
'                pJuros = pobjstcItem.ObjPai.Apol_Ant_PER_JURO_FRAC
'            End If
'        End If
'    End If

    'Cart�o de Cr�dito - Para cart�o de cr�dito juros � 0
    

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

Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
End Property

Public Property Get M46V999() As clsM46V999
    If mClsM46V999 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V999 n�o atribu�do para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V999 j�
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46V999 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46V999 = mClsM46V999
End Property

Public Property Set M99V001(pClsM99V001 As clsM99V001)
    Set mClsM99V001 = pClsM99V001
End Property

Public Property Get M99V001() As clsM99V001
    If mClsM99V001 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M99V001 n�o atribu�do para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM99V001 j�
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM99V001 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M99V001 = mClsM99V001
End Property

Public Property Set M46V103(pClsM46V103 As clsM46V103)
    Set mClsM46V103 = pClsM46V103
End Property

Public Property Get M46V103() As clsM46V103
    If mClsM46V103 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V103 n�o atribu�do para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V103 j�
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46V103 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46V103 = mClsM46V103
End Property

Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
    Set mClsM99V001 = Nothing
    Set mClsM46V103 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
    Set mClsM99V001 = Nothing
    Set mClsM46V103 = Nothing
End Sub
