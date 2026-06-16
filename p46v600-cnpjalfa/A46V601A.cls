VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V601A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================================================================================
' Class Name: clsA46V601A - Carga de Arquivos Recebidos
' Sistema: P46V600
' Analista: Simony Sales
' ======================================================================================================================================================
'C�digos da Tabelas
'TAB_PED              = 001
'TAB_CTRL_EMIS        = 002
'Tab_Ctrl_Emis_Compl  = 002
'TAB_PED_CORR         = 003
'TAB_PED_DIZER        = 004
'TAB_PED_VENC         = 005
'TAB_PED_JUROS        = 006
'TAB_PED_COSS         = 007
'TAB_PED_ENDER        = 008
'TAB_PED_LOC          = 009
'TAB_PED_CLAU         = 010
'TAB_PED_INSP         = 011
'TAB_PED_OBS          = 012
'TAB_PED_LOC_SEGURO   = 013
'TAB_PED_LOC_BEM      = 014
'TAB_PED_FAT_RISCO    = 015
'TAB_PED_COBERT       = 016
'TAB_PED_COBERT_COMPL = 017
'TAB_PED_ESPEC        = 018
'TAB_MOV_AGEN         = 022
'TAB_PED_AGREGADOS    = 025
'===================================================
Private mRegistro   As String   'Registro.
Private mNomeContato        As String
Private mTelContato         As String
Private mEmailContato       As String
Private mCCotacaoOrig       As String
Private mObservacao         As String

Private mCotacao            As String
Private mStatus             As Integer
Private mMaxDesc            As Double
Private mData               As Long
Private mSenha              As String
Public mConexaoEAPDB        As ADODB.Connection
Dim mNossoTempor            As String   'Campo de Teste para nao duplicate
Dim Dummy                   As Boolean
Dim bdBasDados              As ADODB.Connection 'Base de dados.
Dim bdBasDadosSYAS_EMIS     As ADODB.Connection 'Base de dados.
Private mCodCorretor        As String
Private mUsuario            As String
Private mNomArqEnvio        As String           'Nome do arquivo de envio.

Private mClsM46V999         As clsM46V999

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Property Get Class_Name() As String
    Class_Name = "A46V601A"
End Property


Public Function LerIni(Secao As String, NomeChave As String, NomedoArquivo As String) As String
    Dim RetCode As Long
    Dim RetString As String

    LerIni = ""
    RetString = String$(255, 0)
    RetCode = GetPrivateProfileString(Secao, NomeChave, "", RetString, Len(RetString), NomedoArquivo)
    LerIni = Left(RetString, RetCode)

End Function


Public Function gfRecArqCotacao(ByRef pNomArqTransmissao As String, _
                                    ByRef pMensagem As String, _
                                    ByRef pNomArqRetorno As String, _
                                    Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: l� arquivos de cota��o

Dim lAppPath            As String           'App.Path com \ no final.
Dim lRegistro           As String           'Registro.
Dim lSequencia          As String           'Seq��ncia para o arquivo de Cota��o.
Dim lTotal              As Double           'Total.
Dim lobjstcPedido       As stcA46V702B      'Objeto da classe de estrutura stcA46V702B (TAB_PED)
Dim lNumArqTransmissao  As Integer          'Utilizado com FreeFile para arquivo de cota��o.
Dim lCodRegistro        As String
Dim lNossoNumero        As String
Dim lPriVez             As Boolean
Dim lArqAberto          As Boolean
Dim lBuscaCotacao       As Boolean
Dim lPosicao            As Integer
Dim lNomeTabImport      As String
Dim lBeginPendente      As Boolean
Dim lExecProc           As Boolean

Err = 0
On Error GoTo Func_Erro
    gfRecArqCotacao = False
    lNomeTabImport = ""
    pMensagem = ""
    lPriVez = True
       
'1. Posiciona controle.
     If Right$(App.Path, 1) = "\" Then
        lAppPath = App.Path
    Else
        lAppPath = App.Path & "\"
    End If

    If Not M46V999.gfPreenchido(Dir(pNomArqTransmissao)) Then
        pMensagem = "FTE0019 - Arquivo n�o localizado: " & pNomArqTransmissao & _
                    " (grava arquivo de cota��o)."
        GoTo SAIDA
    End If

    mNomArqEnvio = pNomArqTransmissao
    lTotal = 0
    'Abre banco de Dados
    If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, bdBasDados, _
            pMensagem) = False Then
            GoTo SAIDA
    End If
        
    'Abre o arquivo
    lNumArqTransmissao = FreeFile
    lArqAberto = False
    lBuscaCotacao = False
    Open pNomArqTransmissao For Input As #lNumArqTransmissao
        lArqAberto = True
        'L� a primeira linha do arquivo
        Line Input #lNumArqTransmissao, mRegistro
        Do
            If Mid(mRegistro, 1, 5) = "V0001" Then
                '*** Vari�vel inserida para processar mais de uma cota��o por arquivo (pegar mais de uma linha cota��o no if "COTACAO " - abaixo) - SIES - taflorencio 10/10/2013
                lBuscaCotacao = False
                mCodCorretor = val(Trim(Mid(mRegistro, 15, 7)))
                lPosicao = InStr(mRegistro, "USUARIO")
                If lPosicao > 0 Then
                    mUsuario = Mid(mRegistro, lPosicao + 8, 6)
                Else
                    mUsuario = ""
                End If
                GoTo NextRecord
            End If
            If Mid(mRegistro, 1, 8) = "COTACAO " Then
                If lBuscaCotacao = False Then
                    'Primeiro item de Cotacao
                    mNomeContato = Trim(Mid(mRegistro, 29, 100))
                    mEmailContato = Trim(Mid(mRegistro, 129, 50))
                    mTelContato = Trim(Mid(mRegistro, 179, 50))
                    mCCotacaoOrig = Trim(Mid(mRegistro, 229, 12))
                    mObservacao = Trim(Mid(mRegistro, 249))
                    lBuscaCotacao = True
                Else
                    If lNossoNumero <> "" Then
                        pMensagem = pMensagem & "Erro no Registro FINAL: " & lTotal & "-" & lRegistro
                    End If
                End If
                GoTo NextRecord
            End If
            
            If Not IsNumeric(Mid(mRegistro, 1, 3)) Then
                pMensagem = pMensagem & "Erro no Registro: " & lTotal & "-" & lRegistro
                GoTo NextRecord
            End If
            If Mid(mRegistro, 1, 3) = "100" Then
                lNossoNumero = Mid(mRegistro, 4, 20)
                'mNossoTempor = Format(lNossoNumero * Rnd, "00000000000000000000")
                mNossoTempor = lNossoNumero
                GoTo NextRecord
            End If
            If Mid(mRegistro, 1, 3) = "199" Then
                If lNossoNumero <> Mid(mRegistro, 4, 20) Then
                   pMensagem = pMensagem & "Footer do Registro: " & lTotal & "-" & lRegistro
                End If
                lNossoNumero = ""
                '*** Cria��o de Vari�vel lExecProc para processar a rotina mfIncluirSiscotaMulti para arquivos .032 com mais de uma cota��o - SIES - taflorencio 10/10/2013
                lExecProc = True
                GoTo NextRecord
            End If
            If Not IsNumeric(lNossoNumero) Then
                pMensagem = "FTE0017 - Erro ao criar arquivo de cota��o (Nosso N�mero n�o num�rico)."
                Exit Do
            End If
            lCodRegistro = Mid(mRegistro, 1, 3)
                        
            Select Case lCodRegistro
                Case "001"
                    If lPriVez Then
                        If M46V999.gfExeSQL(bdBasDados, "Begin Transaction ", pMensagem) <> 0 Then
                            Call M46V999.gpGraLog(0, pMensagem)
                            Call M46V999.gpFecha2(bdBasDados)
                            GoTo FecharArquivo
                        End If
                        lPriVez = False
                    End If
                    'b) Grava registro de dados comuns.
                    If mfRecPedido(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "002"
                    'c) Grava registro de controle de emiss�o
                    If mfRecCtrlEmis(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "003"
                    'd) Grava registros de corretores.
                    If mfRecCorretores(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "007"
                    'e) Grava registros de cosseguros
                    If mfRecCosseguro(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "008"
                    'f) Grava registros de endere�o.
                    If mfRecEndereco(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "006"
                    'g) Grava registros de juros.
                    If mfRecJuros(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "005"
                    'h) Grava registros de vencimento.
                    If mfRecVenc(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "009"
                    'i) Grava registros de itens.
                    If mfRecItens(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "010"
                    'j) Grava registros de Clausulas.
                    If mfRecClausulas(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "016"
                    'k) Grava registros de Coberturas.
                    If mfRecCoberturas(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "013"
                    'l) Grava registros de Seguros
                    If mfRecSeguro(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "014"
                    'm) Grava registros de Bens.
                    If mfRecBens(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "011"
                    'n) Grava registros de inspe��o
                    If mfRecInspecao(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "012"
                    'o) Grava registros de observa��o
                    If mfRecObservacao(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "017"
                    'p) Grava registros de Dados Complementares das Coberturas.
                    If mfRecCobertCompl(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "004"
                    'q) Grava registros de Texto para Endosso.
                    If mfRecTextoEndosso(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "015"
                    'r) Grava registros de Dados Fatores de Risco.
                    If mfRecFatoresProtecao(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "018"
                    's) Grava registros de Beneficiarios.
                    If mfRecBeneficiario(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "019"
                    's) Grava registros de Beneficiarios.
                    If mfRecPendenciasDiversas(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "026"
                    'j) Grava registros de Pessoas.
                    If mfRecPessoas(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
            
            End Select
            
NextRecord:
            '*** Condi��o inserida dentro do loop para processar a rotina mfIncluirSiscotaMulti para mais de uma cota��o por arquivo - SIES - taflorencio 10/10/2013
            If lExecProc = True Then
            
                If M46V999.gfPreenchido(pMensagem) = True Then
                    GoTo Func_Erro
                End If
            
                '8. Grava registro de controle no final do arquivo (�ltimo registro do arquivo que serve para
                '   verificar se arquivo enviado pelo corretor manteve sua integridade ao ser processado pelo
                '   servidor YASUDA).
                '   Formato do registro de controle: "PROPOSTA XXXXXXXXXXX", onde XXXXXXXXXXX � o valor da soma dos
                '   valores ASC de todos os caracteres das linhas anteriores, gravado de forma invertida.
                '   Por exemplo: Se soma dos valores ASC de todos os caracteres dos registros de um arquivo
                '   (excetuando-se �ltimo registro) for igual a 39.485, �ltimo registro ser�: PROPOSTA 58493.
                'a) Soma valores ASC do registro.
                lTotal = 0
            
                If True = False Then
                    If Not IsMissing(pStcPedido) Then
                        Set pStcPedido = lobjstcPedido
                    Else
                        pMensagem = "Gravar=Falso e pstcPedido n�o informado."
                    End If
                End If
            
                If Not mfIncluirSiscotaMulti(pNomArqRetorno, mNossoTempor, pMensagem) Then
                    lPriVez = True  'Para n�o fazer rollback.
                    GoTo Func_Erro
                End If
            lExecProc = False
            End If
            
            If EOF(lNumArqTransmissao) Then
                Exit Do
            End If
            Line Input #lNumArqTransmissao, mRegistro
            
        Loop
    
FecharArquivo:
    '7. Fecha arquivo - ser� utilizado no pr�ximo comando.
    Close #lNumArqTransmissao
    lArqAberto = False      'marca que o arquivo n�o est� mais aberto
    
'*** Rotina original abaixo foi comentada e transferida para dentro do loop acima ficando apenas o trecho
'*** abaixo que efetua commit no bco. de dados - SIES - taflorencio - 10/10/2013

'''''''''    If M46V999.gfPreenchido(pMensagem) = True Then
'''''''''        GoTo Func_Erro
'''''''''    End If
'''''''''    If Not lPriVez And True Then
'''''''''        Dummy = M46V999.gfExeSQL(bdBasDados, "Commit Transaction ", pMensagem)
'''''''''    End If
'''''''''    If M46V999.gfPreenchido(pMensagem) = True Then
'''''''''        GoTo Func_Erro
'''''''''    End If
'''''''''
'''''''''    '8. Grava registro de controle no final do arquivo (�ltimo registro do arquivo que serve para
'''''''''    '   verificar se arquivo enviado pelo corretor manteve sua integridade ao ser processado pelo
'''''''''    '   servidor YASUDA).
'''''''''    '   Formato do registro de controle: "PROPOSTA XXXXXXXXXXX", onde XXXXXXXXXXX � o valor da soma dos
'''''''''    '   valores ASC de todos os caracteres das linhas anteriores, gravado de forma invertida.
'''''''''    '   Por exemplo: Se soma dos valores ASC de todos os caracteres dos registros de um arquivo
'''''''''    '   (excetuando-se �ltimo registro) for igual a 39.485, �ltimo registro ser�: PROPOSTA 58493.
'''''''''    'a) Soma valores ASC do registro.
'''''''''    lTotal = 0
'''''''''
'''''''''    If True = False Then
'''''''''        If Not IsMissing(pStcPedido) Then
'''''''''            Set pStcPedido = lobjstcPedido
'''''''''        Else
'''''''''            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
'''''''''        End If
'''''''''    End If
'''''''''
'''''''''    If Not mfIncluirSiscotaMulti(pNomArqRetorno, mNossoTempor, pMensagem) Then
'''''''''        lPriVez = True  'Para n�o fazer rollback.
'''''''''        GoTo Func_Erro
'''''''''    End If
    
    If Not lPriVez And True Then
        Dummy = M46V999.gfExeSQL(bdBasDados, "Commit Transaction ", pMensagem)
    End If
    If M46V999.gfPreenchido(pMensagem) = True Then
        GoTo Func_Erro
    End If
    
    gfRecArqCotacao = True
    
SAIDA:
    If lArqAberto = True Then
        'Fecha o arquivo caso o c�digo tenha sido desviado para o erro
        'ap�s ter aberto o arquivo.
        Close #lNumArqTransmissao
        lArqAberto = False
    End If
    Err = 0
    On Error GoTo 0
    Exit Function
    
Func_Erro:
    If Not lPriVez And True And lBeginPendente Then
        Dummy = M46V999.gfExeSQL(bdBasDados, "Rollback Transaction ", pMensagem)
    End If
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "gfRecArqCotacao Erro " & pMensagem
    
    GoTo SAIDA
    
End Function
Private Function mfIncluirSiscotaMulti(pNomArqRetorno As String, pNosso_Numero As String, pMensagem As String) As Boolean
Dim lExec       As String
Dim lCCotacao   As String
Dim lRegistro   As String
Dim lrst        As clsYasRecordSet
Dim lSenha      As String
Dim lAux        As Integer
Dim lNumArqRetorno      As Integer

    'Chama a procedure que ir� retornar o n�mero da cota��o.
    
    lExec = " exec siscota.dbo.proc_CotacaoMultiSyas " & M46V999.gfFormatarCamposSQL(pNosso_Numero) & "," & _
            val(mCCotacaoOrig) & " ," & M46V999.gfFormatarCamposSQL(mNomeContato) & "," & _
            M46V999.gfFormatarCamposSQL(mEmailContato) & "," & _
            M46V999.gfFormatarCamposSQL(mTelContato) & ",'" & Replace(mObservacao, "%!*", vbCrLf) & "'"
            
    If Not M46V999.gfObtRegistro(bdBasDados, lExec, lrst, pMensagem) Then
        Exit Function
    End If
    
    If Not lrst.EOF Then
        lCCotacao = lrst("Cotacao")
    End If
    
    lNumArqRetorno = FreeFile
    Open pNomArqRetorno For Append Shared As #lNumArqRetorno
    lRegistro = "STATUS01"
    lRegistro = lRegistro & pNosso_Numero
    lRegistro = lRegistro & lCCotacao
    Print #lNumArqRetorno, lRegistro
    Close #lNumArqRetorno
    
    lAux = Second(Now)
Retorno:
    If lAux > 25 Then
        lAux = CInt(lAux / 2)
        GoTo Retorno
    End If
    lSenha = Chr(Int(lAux * Rnd) + 65)
    
    lAux = Second(Now)
retorno2:
    If lAux > 25 Then
        lAux = CInt(lAux / 2)
        GoTo retorno2
    End If
    lSenha = lSenha & Chr(Int(lAux * Rnd) + 65)
    
    lAux = Second(Now)
retorno3:
    If lAux > 25 Then
        lAux = CInt(lAux / 2)
        GoTo retorno3
    End If
    lSenha = lSenha & Chr(Int(lAux * Rnd) + 65)
    
    lAux = Second(Now) + Minute(Now) * Hour(Now)
retorno4:
    If lAux > 999 Then
        lAux = CInt(lAux / 2)
        GoTo retorno4
    End If
    lSenha = lSenha & Format(CInt(lAux * Rnd) + 100, "000") '
    
    
    Dim lPreArq As String
    Dim lI      As Long
    Dim lLocal  As Integer
    
    lLocal = InStr(1, mNomArqEnvio, "\")
    Do While lLocal <> 0
        lI = lLocal
        lLocal = InStr(lLocal + 1, mNomArqEnvio, "\")
        
    Loop
    
    
    lExec = "Insert into siscota.dbo.tab_syas_cotacao (CCOTACAO,NOSSO_NUMERO, " & _
            " DATA_INCLUSAO, COD_RAMO, cod_status, STATUS_ENVIO, " & _
            "COD_CORRETOR, COD_USUARIO, NomArq, SENHA ) values ( " & _
            M46V999.gfFormatarCamposSQL(lCCotacao) & "," & _
            M46V999.gfFormatarCamposSQL(pNosso_Numero) & "," & _
            M46V999.gfForData(3) & "," & _
            M46V999.gCod_Ramo & ", 1, 0," & _
            mCodCorretor & "," & mUsuario & " , " & Replace(M46V999.gfFormatarCamposSQL(Mid(mNomArqEnvio, lI + 1)), ".032", "") & ",'" & lSenha & "')"
            
    If M46V999.gfExeSQL(bdBasDados, lExec, pMensagem) <> 0 Then
        Exit Function
    End If
    
    mfIncluirSiscotaMulti = True
End Function

Public Function gfRecArqTransmissao(ByRef pNomArqTransmissao As String, _
                                    ByRef pMensagem As String, _
                                    Optional pStcPedido As stcA46V702B) As Boolean
'**********************
'Fun��o compartilhada com o projeto P24V007.exe (Servidor SYAS-EDI executado como servi�o)
'        N�O PODE EMITIR MENSAGEM!!!! PROIBIDO MSGBOX!!!!
'
'**********************
    'Fun��o: l� arquivos de transmiss�o.

Dim lAppPath            As String           'App.Path com \ no final.
Dim lRegistro           As String           'Registro.
Dim lSequencia          As String           'Seq��ncia para o arquivo de transmiss�o.
Dim lTotal              As Double           'Total.
Dim lobjstcPedido       As stcA46V702B      'Objeto da classe de estrutura stcA46V702B (TAB_PED)
Dim lobjstcPedLoc       As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
Dim lNumArqTransmissao  As Integer          'Utilizado com FreeFile para arquivo de transmiss�o.
Dim lCodRegistro        As String
Dim lNossoNumero        As String
Dim lFlagGravar         As Boolean
Dim lPriVez             As Boolean
Dim lArqAberto          As Boolean
Dim lCkPoint            As String
Dim lTipTransmissao     As Integer
Dim lrst                As clsYasRecordSet
Dim lRst2               As clsYasRecordSet
Dim lContador           As Integer
Dim lTarifaCorte        As Long
Dim lNomeTabImport      As String
Dim lSelect             As String
Dim lInsert             As String
Dim lNumPITransmissao   As Double
Dim lVincular           As Boolean
Dim lBeginPendente      As Boolean
Dim lNumPITeste         As String
Dim flagInspecaoEquipamento As Boolean
Dim lCountErro          As Integer
Dim lCount              As Integer
Dim pMsg                As String
Dim lNumeroLinhaProcessando As Long

Err = 0
lNumeroLinhaProcessando = 0

On Error GoTo Func_Erro
    gfRecArqTransmissao = False
    pMensagem = ""
    lPriVez = True
    lCount = 0
    lCountErro = 0
    
'1. Verifica se esta sendo executado pelo SYAS-EDI
    lNomeTabImport = ""
    If UCase(App.EXEName) = "P24V007" Then
        lFlagGravar = False
    Else
        lFlagGravar = True
        If Right$(pNomArqTransmissao, 3) = "046" Then
            lNomeTabImport = "LOG_"
        Else
            lNomeTabImport = "FTP_"
        End If
    End If
    
'2. Posiciona controle.
     If Right$(App.Path, 1) = "\" Then
        lAppPath = App.Path
    Else
        lAppPath = App.Path & "\"
    End If

    lCkPoint = "01"
    If Not M46V999.gfPreenchido(Dir(pNomArqTransmissao)) Then
        pMensagem = "FTE0019 - Arquivo n�o localizado: " & pNomArqTransmissao & _
                    " (grava arquivo de transmiss�o)."
        GoTo SAIDA
    End If

    lTotal = 0
    'Abre banco de Dados
    lCkPoint = "02"
    If lFlagGravar = True Then
        If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, bdBasDados, _
                pMensagem) = False Then
                GoTo SAIDA
        End If
    End If
        
    'Abre o arquivo
    lNumArqTransmissao = FreeFile
    lArqAberto = False
    lCkPoint = "03"
    lNumPITransmissao = 0
    lNumPITeste = ""
    lVincular = False
    lBeginPendente = False
    
    Open pNomArqTransmissao For Input As #lNumArqTransmissao
        lArqAberto = True
        'L� a primeira linha do arquivo
        lCkPoint = "04"
        Line Input #lNumArqTransmissao, mRegistro
        Do
            lNumeroLinhaProcessando = lNumeroLinhaProcessando + 1
            flagInspecaoEquipamento = False
            lCkPoint = "05"
            If Len(Trim(mRegistro)) = 0 Then
                GoTo NextRecord
            End If
            
            If Mid(mRegistro, 1, 5) = "V0001" Then
                GoTo NextRecord
            End If
            If Mid(mRegistro, 1, 8) = "PROPOSTA" Then
                If lNossoNumero <> "" Then
                    pMensagem = pMensagem & "Erro no Registro FINAL: " & lTotal & "-" & lRegistro
                End If
                GoTo NextRecord
            End If
            
            If Not IsNumeric(Mid(mRegistro, 1, 3)) Then
                pMensagem = pMensagem & "Erro no Registro: " & lTotal & "-" & lRegistro
                GoTo NextRecord
            End If
            If Mid(mRegistro, 1, 3) = "100" Then
                lNossoNumero = Mid(mRegistro, 4, 20)
                'mNossoTempor = Format(lNossoNumero * Rnd, "00000000000000000000")
                mNossoTempor = lNossoNumero
                lVincular = True
                GoTo NextRecord
            End If
            If Mid(mRegistro, 1, 3) = "199" Then
                If lNossoNumero <> Mid(mRegistro, 4, 20) Then
                   pMensagem = pMensagem & "Footer do Registro: " & lTotal & "-" & lRegistro
                End If
                'Ampsoft  Gravar na tab_vinculo Simony - SIES
                Set pStcPedido = lobjstcPedido
                GoSub GravaVinculo
                
                lNossoNumero = ""
                GoTo NextRecord
            End If
            ' Inspe��o � Equipamentos
            If Mid(mRegistro, 1, 3) = "023" Then
                lNossoNumero = Mid(mRegistro, 4, 20)
            End If

            If Not IsNumeric(lNossoNumero) Then
                pMensagem = "FTE0017 - Erro ao criar arquivo de transmiss�o (Nosso N�mero n�o num�rico)."
                Exit Do
            End If
            lCodRegistro = Mid(mRegistro, 1, 3)
            'On Error GoTo 0
            lCkPoint = "06"
            Select Case lCodRegistro
                Case "001"
                    If lPriVez = False Then
                        Set lobjstcPedido = New stcA46V702B
                        Set lobjstcPedido.M46V999 = M46V999
                        
                        lobjstcPedido.Tip_Transmissao = lTipTransmissao
                    Else
                        If lFlagGravar Then
                            lCkPoint = "07"
                            If M46V999.gfExeSQL(bdBasDados, "Begin Transaction ", pMensagem) <> 0 Then
                                Call M46V999.gpFecha2(bdBasDados)
                                GoTo FecharArquivo
                            End If
                            lBeginPendente = True
                        End If
                        lPriVez = False
                        If Not IsMissing(pStcPedido) Then
                            Set lobjstcPedido = pStcPedido
                            If Not pStcPedido Is Nothing Then
                                lTipTransmissao = pStcPedido.Tip_Transmissao
                            End If
                        End If
                    End If
                    'b) Grava registro de dados comuns.
                    lCkPoint = "08"
                    lNumPITransmissao = val(Mid(mRegistro, 912, 10))
                    lNumPITeste = lNumPITeste & lNumPITransmissao & " ; "
                    lCount = lCount + 1
                    If mfRecPedido(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "002"
                    'c) Grava registro de controle de emiss�o
                    lCkPoint = "09"
                    If mfRecCtrlEmis(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "003"
                    'd) Grava registros de corretores.
                    lCkPoint = "10"
                    If mfRecCorretores(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "007"
                    'e) Grava registros de cosseguros
                    lCkPoint = "11"
                    If mfRecCosseguro(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "008"
                    'f) Grava registros de endere�o.
                    lCkPoint = "12"
                    If mfRecEndereco(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "006"
                    'g) Grava registros de juros.
                    lCkPoint = "13"
                    If mfRecJuros(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "005"
                    'h) Grava registros de vencimento.
                    lCkPoint = "14"
                    If mfRecVenc(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "009"
                    'i) Grava registros de itens.
                    lCkPoint = "15"
                    If mfRecItens(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "010"
                    'j) Grava registros de Clausulas.
                    lCkPoint = "16"
                    If mfRecClausulas(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "016"
                    'k) Grava registros de Coberturas.
                    lCkPoint = "17"
                    If mfRecCoberturas(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "013"
                    'l) Grava registros de Seguros
                    lCkPoint = "18"
                    If mfRecSeguro(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "014"
                    'm) Grava registros de Bens.
                    lCkPoint = "19"
                    If mfRecBens(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "011"
                    'n) Grava registros de inspe��o
                    lCkPoint = "20"
                    If mfRecInspecao(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "012"
                    'o) Grava registros de observa��o
                    lCkPoint = "21"
                    If mfRecObservacao(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "017"
                    'p) Grava registros de Dados Complementares das Coberturas.
                    lCkPoint = "22"
                    If mfRecCobertCompl(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "004"
                    'q) Grava registros de Texto para Endosso.
                    lCkPoint = "23"
                    If mfRecTextoEndosso(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "015"
                    'r) Grava registros de Dados Fatores de Risco.
                    lCkPoint = "24"
                    If mfRecFatoresProtecao(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "018"
                    's) Grava registros de Beneficiarios.
                    lCkPoint = "25"
                    If mfRecBeneficiario(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "019"
                    lCkPoint = "26"
                    If mfRecPendenciasDiversas(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "020"
                    't) Grava registros de Beneficiarios.
                    lCkPoint = "27"
                    If mfRecLMI(lNossoNumero, pMensagem, "", lNomeTabImport, True, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "021"
                    'u) Grava registros de Pro-labore.
                    lCkPoint = "28"
                    If mfRecProLabore(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "022"
                    'v) Grava registros de Pro-labore.
                    lCkPoint = "29"
                    If mfRecMovAgen(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "023"
                    lCkPoint = "30"
                    'Grava registros de Inspe��o de Equipamentos
                    If mfRecTransInspecaoEquipamentos(lNossoNumero, pMensagem, lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                    flagInspecaoEquipamento = True
                Case "024"
                    lCkPoint = "31"
                    'x) Grava registros de telefone segurado.
                    If mfRecPedTelefones(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "025"
                    lCkPoint = "32"
                    'x) Grava registros de Agregados.
                    If mfRecPedAgregados(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "026"
                    lCkPoint = "33"
                    'x) Grava registros de Agregados.
                    If mfRecPessoas(lNossoNumero, pMensagem, "", lNomeTabImport, lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "027"
                    lCkPoint = "52"
                    'z) Grava registros de Cart�o de Cr�dito.
                    If mfRecPgtoCartaoCredito(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
            End Select
            
NextRecord:
            lCkPoint = "34"
            If EOF(lNumArqTransmissao) Then
                Exit Do
            End If
            lCkPoint = "35"
            Line Input #lNumArqTransmissao, mRegistro
        Loop
    
FecharArquivo:
    '7. Fecha arquivo - ser� utilizado no pr�ximo comando.
    Close #lNumArqTransmissao
    lArqAberto = False      'marca que o arquivo n�o est� mais aberto
    
    'Se a quantidade de PI's for igual a quantidade de PI's n�o procesasdos na proc proc_StartFluxoEmissaoTransmissaoMulti
    'Gerar arquivo com erro (.N08).
    lCkPoint = "36"
    If lCount = lCountErro Then
        GoTo Func_Erro
    End If
    If M46V999.gfPreenchido(pMensagem) = True Then
        GoTo Func_Erro
    End If
        lCkPoint = "37"
    If Not lPriVez And lFlagGravar And lBeginPendente Then
        lCkPoint = "38"
        Dummy = M46V999.gfExeSQL(bdBasDados, "Commit Transaction ", pMsg)
        lBeginPendente = False
    End If
            
    Call gM46V999.gpGraLog(0, "Arquivo: " & pNomArqTransmissao & " - PIs gravados: " & lNumPITeste)
        lCkPoint = "39"
    If M46V999.gfPreenchido(pMensagem) = True Then
        GoTo Func_Erro
    End If

    '8. Grava registro de controle no final do arquivo (�ltimo registro do arquivo que serve para
    '   verificar se arquivo enviado pelo corretor manteve sua integridade ao ser processado pelo
    '   servidor YASUDA).
    '   Formato do registro de controle: "PROPOSTA XXXXXXXXXXX", onde XXXXXXXXXXX � o valor da soma dos
    '   valores ASC de todos os caracteres das linhas anteriores, gravado de forma invertida.
    '   Por exemplo: Se soma dos valores ASC de todos os caracteres dos registros de um arquivo
    '   (excetuando-se �ltimo registro) for igual a 39.485, �ltimo registro ser�: PROPOSTA 58493.
    'a) Soma valores ASC do registro.
    
'Simony
    Set pStcPedido = lobjstcPedido
    GoSub GravaVinculo


    gfRecArqTransmissao = True
    
SAIDA:
    If lArqAberto = True Then
        'Fecha o arquivo caso o c�digo tenha sido desviado para o erro
        'ap�s ter aberto o arquivo.
        Close #lNumArqTransmissao
        lArqAberto = False
    End If
    Err = 0
    On Error GoTo 0
    Exit Function
    
Func_Erro:
    If Not lPriVez And lFlagGravar And lBeginPendente Then
        Dummy = M46V999.gfExeSQL(bdBasDados, "Rollback Transaction ", pMsg)
    End If
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = "gfRecArqTransmissao Erro " & Err & " - " & ERROR & " - " & pMsg
    End If
    pMensagem = pMensagem & " || Linha " & lNumeroLinhaProcessando & " || lCount=" & lCount & ", lCountErro=" & lCountErro & " || gfRecArqTransmissao.CkPoint(" & lCkPoint & ") - " & pMsg
    GoTo SAIDA
    
GravaVinculo:
    lCkPoint = "39"
    lTotal = 0
    lTarifaCorte = 1010
    If pStcPedido.Cod_Ramo = 810 Or pStcPedido.Cod_Ramo = 930 Or pStcPedido.Cod_Ramo = 910 Or pStcPedido.Cod_Ramo = 932 Or pStcPedido.Cod_Ramo = 112 Or pStcPedido.Cod_Ramo = 113 Then
        lTarifaCorte = 0
    End If
    If Not lVincular Then Return
    lCkPoint = "40"
    'Ajuste para n�o executar rotina de atualiza��o de PI quando for emiss�o SIES e Vida - GroupLife
    If lFlagGravar = False And lobjstcPedido.objstcCtrlEmis.NUM_PI <> 0 Then GoTo FimVinculo
    lCkPoint = "41"
    If lFlagGravar = False Then
        lCkPoint = "42"
        If Not IsMissing(pStcPedido) Then
            Set pStcPedido = lobjstcPedido
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
        End If
        pStcPedido.NUM_PI = "0"
        If pStcPedido.Cod_Produto >= lTarifaCorte Then
            lCkPoint = "43"
            lSelect = " select * from Tab_Vinculo_Proposta_PI where nosso_numero = " & M46V999.gfFormatarCamposSQL(pStcPedido.nosso_numero)
            If Not M46V999.gfObtRegistro(mConexaoEAPDB, lSelect, lrst, pMensagem) Then
                GoTo Func_Erro
            End If
            lContador = 0
            If lrst.EOF Then
                lCkPoint = "44"
                lSelect = " exec [dbo].proc_GetProxPI 2"
                If Not M46V999.gfObtRegistro(mConexaoEAPDB, lSelect, lrst, pMensagem) Then
                    GoTo Func_Erro
                End If
                If Not lrst.EOF Then
                    lInsert = " Insert into dbo.Tab_Vinculo_Proposta_PI (Nosso_Numero, Num_PI, Data , Hora, Usuario) values ( "
                    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pStcPedido.nosso_numero) & ","
                    lInsert = lInsert & M46V999.gfFormatarCamposSQL(lrst("Num_PI")) & ","
                    lInsert = lInsert & " convert(varchar,getdate(),112) ,"
                    lInsert = lInsert & " replace(left(convert(varchar,getdate(),114),8),':','') ,"
                    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pStcPedido.Cod_User_Liber) & ")"
                    
                    pStcPedido.NUM_PI = lrst("num_pi")
                    If M46V999.gfExeSQL(mConexaoEAPDB, lInsert, pMensagem) <> 0 Then
                        If UCase(pMensagem) Like "*VIOLATION OF PRIMARY KEY*" Then
                            pMensagem = "Proposta j� vinculada em outro processo de emiss�o"
                            'GoTo Func_Erro
                            GoTo FimVinculo
                        End If
                    End If
                End If
            Else
                pStcPedido.NUM_PI = lrst("num_pi")
            End If
        End If
        
    Else
        lCkPoint = "45"
        If lNumPITransmissao <> 0 Then
            If lBeginPendente Then
                Dummy = M46V999.gfExeSQL(bdBasDados, "Commit Transaction ", pMsg)
                lBeginPendente = False
            End If
            lPriVez = True
            lInsert = " Insert into dbo.Tab_Vinculo_Proposta_PI (Nosso_Numero, Num_PI, Data , Hora, Usuario) values ( "
            lInsert = lInsert & M46V999.gfFormatarCamposSQL(pStcPedido.nosso_numero) & ","
            lInsert = lInsert & M46V999.gfFormatarCamposSQL(lNumPITransmissao) & ","
            lInsert = lInsert & " convert(varchar,getdate(),112) ,"
            lInsert = lInsert & " replace(left(convert(varchar,getdate(),114),8),':','') ,"
            lInsert = lInsert & M46V999.gfFormatarCamposSQL(pStcPedido.Cod_User_Liber) & ")"
            

            If M46V999.gfExeSQL(mConexaoEAPDB, lInsert, pMensagem) <> 0 Then
                If UCase(pMensagem) Like "*VIOLATION OF PRIMARY KEY*" Then
                    pMensagem = "Proposta j� vinculada em outro processo de emiss�o"
                    'GoTo FimVinculo
                    'GoTo Func_Erro
                End If
            End If
        End If
        lCkPoint = "46"
        For Each lobjstcPedLoc In lobjstcPedido.ColecaoPedLoc
            If lobjstcPedLoc.Insp_Obr = True Then
                'Verifico se a inspe��o j� est� gravada
                Dim lSql    As String
                lSql = "Select * from ramosdiversos.dbo.tab_trans_Insp where nosso_numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.nosso_numero)
                If Not M46V999.gfObtRegistro(bdBasDados, lSql, lrst, pMensagem) Then
                    GoTo Func_Erro
                End If
                If lrst.EOF Then
                    Call gpGeraLogCalcInspecao(lobjstcPedido, lobjstcPedLoc, True, M46V999.gDiretorio & "Multi\")
                End If
            End If
        Next
        If (Not UCase(pNomArqTransmissao) Like "*.091.091") And (Not Right$(pNomArqTransmissao, 3) = "046") Then
            If lobjstcPedido.Cod_Produto >= lTarifaCorte Then
                lCkPoint = "47"
                lSelect = " select * from Tab_Vinculo_Proposta_PI where nosso_numero = " & M46V999.gfFormatarCamposSQL(lobjstcPedido.nosso_numero)
                If Not M46V999.gfObtRegistro(mConexaoEAPDB, lSelect, lrst, pMensagem) Then
                    GoTo Func_Erro
                End If

                If Not lrst.EOF Then
                    
                    Dim gTipYAS_ND As String
                    Dim gEndYAS_ND As String
                    Dim bdRamosDiversos As ADODB.Connection
                    Dim rsCORR As clsYasRecordSet
                    Dim numeroPiEmitido As String
                    numeroPiEmitido = 0
                   
                    If UBound(Filter(Split(LerIni("BP", "CODRAMO", M46V999.gAppPath & "BP.INI"), ";"), lobjstcPedido.Cod_Ramo)) = 0 Then
                  
                        Dim contadorStartFluxo As Integer
                  
                        contadorStartFluxo = 0
                  
                        If LerIni("BP", "LIGADO", M46V999.gAppPath & "BP.INI") = "1" And IsInArray(lobjstcPedido.Cod_Reg, Split(LerIni("BP", "CODREG", M46V999.gAppPath & "BP.INI"), ";")) = True Then
                      
                            lSelect = "exec P0044800.DBO.PROC_CONSULTA_PI_BP_EMITIDO " & lobjstcPedido.Num_Apol_An & " , " & M46V999.gfFormatarCamposSQL(lobjstcPedido.Num_Cgc_Cpf)
    
                            If M46V999.gfAbrBasDados("RAMOSDIVERSOS", gTipYAS_ND, gEndYAS_ND, bdRamosDiversos, pMensagem) = False Then
                               Exit Function
                            End If
                            
                            If M46V999.gfObtRegistro(bdRamosDiversos, lSelect, rsCORR, pMensagem) = False Then
                               Exit Function
                            End If
                            
                            numeroPiEmitido = IIf(rsCORR("NUM_PI_EMIT") Is Nothing, 0, rsCORR("NUM_PI_EMIT"))
                        End If
                        
                        If (numeroPiEmitido = 0) Then
                         
                            lSelect = " exec ramosdiversos.dbo.Proc_startfluxoemissaoCorporativo " & M46V999.gfFormatarCamposSQL(lrst("Nosso_numero")) & "," & lrst("Num_PI")
                         Else
                            lSelect = "exec P0044800.DBO.PROC_COMPLEMENTA_DADOS_BP " & lobjstcPedido.Num_Apol_An & " , " & M46V999.gfFormatarCamposSQL(lobjstcPedido.Num_Cgc_Cpf) & " , " & lobjstcPedido.NUM_PI & " , " & numeroPiEmitido
                            
                            If M46V999.gfAbrBasDados("RAMOSDIVERSOS", gTipYAS_ND, gEndYAS_ND, bdRamosDiversos, pMensagem) = False Then
                               Exit Function
                            End If
                            
                            If M46V999.gfObtRegistro(bdRamosDiversos, lSelect, rsCORR, pMensagem) = False Then
                               Exit Function
                            End If
                         
                            lSelect = " exec ramosdiversos.dbo.Proc_startfluxoemissaoCorporativo " & M46V999.gfFormatarCamposSQL(lrst("Nosso_numero")) & "," & lrst("Num_PI") & ", 1"
                         End If
                  
                         Do Until gfObtRegistroStartFluxo(lSelect, pMensagem) = True Or contadorStartFluxo = 5
                  
                            Call gM46V999.gpGraLog(0, "Arquivo: " & pNomArqTransmissao & " - erro no PI: " & lNumPITransmissao & " - " & pMensagem)
                  
                            pMensagem = ""
                  
                            Sleep (2000)
                  
                            contadorStartFluxo = contadorStartFluxo + 1 'contador
                         Loop
                   
                    Else
                  
                        pMensagem = ""
                        lCkPoint = "48"
                        lSelect = " exec ramosdiversos.dbo.proc_StartFluxoEmissaoTransmissaoMulti " & M46V999.gfFormatarCamposSQL(lrst("Nosso_numero")) & "," & lrst("Num_PI")
                        lCkPoint = "49"
                        If gfObtRegistroStartFluxo(lSelect, pMensagem) = False Then
                            lCkPoint = "50"
                            Call gM46V999.gpGraLog(0, "Arquivo: " & pNomArqTransmissao & " - proc: " & lSelect & " - " & pMensagem)
                            lCountErro = lCountErro + 1
                        Else
                            lCkPoint = "51"
                                'Call gM46V999.gpGraLog(0, "PI OK = " & lNumPITransmissao & " - " & pMensagem)
                            pMensagem = ""
                        End If
                    End If
                End If
            End If
        End If
    End If
FimVinculo:
    lVincular = False

Return
End Function

Public Function IsInArray(FindValue As Variant, arrSearch As Variant) As Boolean
    If Not IsArray(arrSearch) Then Exit Function
    If Not IsNumeric(FindValue) Then FindValue = UCase(FindValue)
    IsInArray = InStr(1, vbNullChar & Join(arrSearch, vbNullChar) & vbNullChar, _
    vbNullChar & FindValue & vbNullChar) > 0
Exit Function
End Function

Public Function gfObtRegistroStartFluxo(ByVal pComando As String, ByRef pMensagem As String) As Boolean
Dim lContador           As Integer
Dim lrs                 As ADODB.Recordset
Dim lSit                As Boolean
Dim lI                  As Integer
Dim lCkPoint            As Integer
Dim Erro                As Long
Dim lErro               As Double
Dim lSqlErro            As ADODB.ERROR
Dim lMensagem           As String
Dim lMensagemErr        As String

    Err = 0
    On Error GoTo Func_Erro
    lCkPoint = 0
    lContador = 0
    pMensagem = ""
    lMensagem = ""
    gfObtRegistroStartFluxo = True
    lSit = False
    
    Set lrs = New ADODB.Recordset
    lrs.Open pComando, bdBasDados
    lCkPoint = 1

    Do While Not lrs Is Nothing
        lCkPoint = 2
        While Not lrs.EOF
            For lI = 0 To lrs.Fields.Count - 1
                pMensagem = pMensagem & lrs.Fields(lI).Value & " - "
            Next lI
            If lrs.Fields(0).Name = "Protocolo" Then
               If lrs("Protocolo") = -1 Then
                  lCkPoint = 3
                  lSit = False
                  GoTo Func_Erro
                Else
                  lCkPoint = 4
                  lSit = True
               End If
            End If
            lrs.MoveNext
        Wend
        Set lrs = lrs.NextRecordset
    Loop
    
   If lSit = False Then
       GoTo Func_Erro
   End If

SAIDA:
    Err = 0
    On Error GoTo 0
    Exit Function
    
Func_Erro:
    gfObtRegistroStartFluxo = False

    If Trim(pMensagem) <> "" Then
        lErro = 1 'Err.Number
        pMensagem = pMensagem & Err & " | " & ERROR
    End If
    For Each lSqlErro In bdBasDados.Errors
        lErro = 1
        lMensagem = lMensagem & " | " & lSqlErro.Description
    Next
    pMensagem = pMensagem & "| " & lMensagem

    GoTo SAIDA
End Function
Public Function gfRecArqCotacaoAceitaRec(ByRef pNomArqTransmissao As String, _
                                    ByRef pMensagem As String) As Boolean
    'Fun��o: l� arquivos de retorno da cota��o

Dim lAppPath            As String           'App.Path com \ no final.
Dim lRegistro           As String           'Registro.
Dim lSequencia          As String           'Seq��ncia para o arquivo de transmiss�o.
Dim lTotal              As Double           'Total.
Dim lobjstcPedido       As stcA46V702B      'Objeto da classe de estrutura stcA46V702B (TAB_PED)
Dim lNumArqTransmissao  As Integer          'Utilizado com FreeFile para arquivo de transmiss�o.
Dim lCodRegistro        As String
Dim lNossoNumero        As String
Dim lFlagGravar         As Boolean
Dim lPriVez             As Boolean
Dim lArqAberto          As Boolean
Dim lobjPedido          As clsA46V702A

Err = 0
On Error GoTo Func_Erro
    gfRecArqCotacaoAceitaRec = False
    pMensagem = ""
    lPriVez = True
    
'2. Posiciona controle.
     If Right$(App.Path, 1) = "\" Then
        lAppPath = App.Path
    Else
        lAppPath = App.Path & "\"
    End If

    If Not M46V999.gfPreenchido(Dir(pNomArqTransmissao)) Then
        pMensagem = "FTE0019 - Arquivo n�o localizado: " & pNomArqTransmissao & _
                    " (grava arquivo de transmiss�o)."
        GoTo SAIDA
    End If

    lTotal = 0
    'Abre banco de Dados
    If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, bdBasDados, _
            pMensagem) = False Then
            GoTo SAIDA
    End If
        
    'Abre o arquivo
    lNumArqTransmissao = FreeFile
    lArqAberto = False
    mSenha = ""
    mCotacao = ""
    mData = 0
        
    
    Open pNomArqTransmissao For Input As #lNumArqTransmissao
        lArqAberto = True
        'L� a primeira linha do arquivo
        Line Input #lNumArqTransmissao, mRegistro
        Do
            If Mid(mRegistro, 1, 5) = "V0001" Then
                GoTo NextRecord
            End If
            If Mid(mRegistro, 1, 8) = "COTACAO " Then
                If lNossoNumero <> "" Then
                    pMensagem = pMensagem & "Erro no Registro FINAL: " & lTotal & "-" & lRegistro
                End If
                If M46V999.gfPreenchido(mCotacao) Then
                    Exit Do
                End If
                mCotacao = Mid(mRegistro, 29, 12) 'Cotacao
                mStatus = Mid(mRegistro, 41, 2) 'Status
                mSenha = Mid(mRegistro, 43, 6)  'Senha
                mData = Mid(mRegistro, 49, 8)   'Data
                If Len(Trim(Mid(mRegistro, 57, 4))) = 4 Then
                    mMaxDesc = Mid(mRegistro, 57, 4) / 100
                Else
                    mMaxDesc = Mid(mRegistro, 57, 4)
                End If
                If mStatus = 3 Then GoTo FecharArquivo
                
                Set lobjPedido = New clsA46V702A
                Set lobjPedido.M46V999 = M46V999
                Set lobjstcPedido = New stcA46V702B
                Set lobjstcPedido.M46V999 = M46V999
                
                lobjstcPedido.nosso_numero = Mid(mRegistro, 9, 20)
                If Not lobjPedido.mfCarregarColecao(lobjstcPedido, "", "") Then
                    Exit Do
                End If
                If lobjPedido.Colecao.Count > 0 Then
                    Set lobjstcPedido = lobjPedido.Colecao.Item(1)
                    If lobjstcPedido.Cod_Sit_Ped = e_SitPedido_Transmitido Then GoTo FecharArquivo
                End If
                Set lobjstcPedido = Nothing
                
                
                
                GoTo NextRecord
            End If
            
            If Not IsNumeric(Mid(mRegistro, 1, 3)) Then
                pMensagem = pMensagem & "Erro no Registro: " & lTotal & "-" & lRegistro
                GoTo NextRecord
            End If
            If Mid(mRegistro, 1, 3) = "100" Then
                lNossoNumero = Mid(mRegistro, 4, 20)
                'mNossoTempor = Format(lNossoNumero * Rnd, "00000000000000000000")
                mNossoTempor = lNossoNumero
                GoTo NextRecord
            End If
            If Mid(mRegistro, 1, 3) = "199" Then
                If lNossoNumero <> Mid(mRegistro, 4, 20) Then
                   pMensagem = pMensagem & "Footer do Registro: " & lTotal & "-" & lRegistro
                End If
                lNossoNumero = ""
                GoTo NextRecord
            End If
            If Not IsNumeric(lNossoNumero) Then
                pMensagem = "FTE0017 - Erro ao criar arquivo de transmiss�o (Nosso N�mero n�o num�rico)."
                Exit Do
            End If
            lCodRegistro = Mid(mRegistro, 1, 3)
            lFlagGravar = True
            Select Case lCodRegistro
                Case "001"
                    'b) Grava registro de dados comuns.
                    If mfRecPedido(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "002"
                    'c) Grava registro de controle de emiss�o
                    If mfRecCtrlEmis(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "003"
                    'd) Grava registros de corretores.
                    If mfRecCorretores(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "007"
                    'e) Grava registros de cosseguros
                    If mfRecCosseguro(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "008"
                    'f) Grava registros de endere�o.
                    If mfRecEndereco(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "006"
                    'g) Grava registros de juros.
                    If mfRecJuros(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "005"
                    'h) Grava registros de vencimento.
                    If mfRecVenc(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "009"
                    'i) Grava registros de itens.
                    If mfRecItens(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "010"
                    'j) Grava registros de Clausulas.
                    If mfRecClausulas(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "016"
                    'k) Grava registros de Coberturas.
                    If mfRecCoberturas(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "013"
                    'l) Grava registros de Seguros
                    If mfRecSeguro(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "014"
                    'm) Grava registros de Bens.
                    If mfRecBens(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "011"
                    'n) Grava registros de inspe��o
                    If mfRecInspecao(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "012"
                    'o) Grava registros de observa��o
                    If mfRecObservacao(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "017"
                    'p) Grava registros de Dados Complementares das Coberturas.
                    If mfRecCobertCompl(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "004"
                    'q) Grava registros de Texto para Endosso.
                    If mfRecTextoEndosso(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "015"
                    'r) Grava registros de Dados Fatores de Risco.
                    If mfRecFatoresProtecao(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "018"
                    's) Grava registros de Beneficiarios.
                    If mfRecBeneficiario(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                Case "019"
                    's) Grava registros de Beneficiarios.
                    If mfRecPendenciasDiversas(lNossoNumero, pMensagem, "", "", lFlagGravar, lobjstcPedido) = False Then
                        Exit Do
                    End If
                    
            End Select
NextRecord:
            If EOF(lNumArqTransmissao) Then
                Exit Do
            End If
            Line Input #lNumArqTransmissao, mRegistro
        Loop
    
FecharArquivo:
    '7. Fecha arquivo - ser� utilizado no pr�ximo comando.
    Close #lNumArqTransmissao
    lArqAberto = False      'marca que o arquivo n�o est� mais aberto
    
    If M46V999.gfPreenchido(pMensagem) = True Then
        GoTo Func_Erro
    End If
    If M46V999.gfPreenchido(pMensagem) = True Then
        GoTo Func_Erro
    End If
    If Not mfGraP0045400 Then GoTo Func_Erro
    gfRecArqCotacaoAceitaRec = True
    
SAIDA:
    If lArqAberto = True Then
        'Fecha o arquivo caso o c�digo tenha sido desviado para o erro
        'ap�s ter aberto o arquivo.
        Close #lNumArqTransmissao
        lArqAberto = False
    End If
    Err = 0
    On Error GoTo 0
    Exit Function
    
Func_Erro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    
    pMensagem = "gfRecArqCotacaoAceitaRec Erro " & pMensagem
    
    GoTo SAIDA
    
End Function
Private Function mfGraP0045400() As Boolean
    
    Dim bdConexao   As ADODB.Connection
    Dim lComando    As String
    Dim lrst        As clsYasRecordSet
    Dim lMensagem   As String
    
    
    mfGraP0045400 = False
            

    If M46V999.gfAbrBasDados("P0045400", M46V999.gTipP0045400, M46V999.gEndP0045400, bdConexao, lMensagem) = False Then
        Exit Function
    End If
            
    lComando = " alter table p0045400 alter column MaxDesc double "
    If M46V999.gfExeSQL(bdConexao, lComando, lMensagem) <> 0 Then
        Call M46V999.gpGraLog(0, lMensagem)
        'Exit Function
    End If
    
    lComando = " Select * from P0045400 where nosso_numero = " & M46V999.gfFormatarCamposSQL(mNossoTempor)
    
    If Not M46V999.gfObtRegistro(bdConexao, lComando, lrst, lMensagem) Then
        Call M46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If
    If lrst.EOF Then
        lComando = "Insert into P0045400 (CCotacao , Nosso_Numero , Dat_Cotacao , "
        lComando = lComando & " Email, Nome, Telefone, Status, Usuario, MaxDesc,Senha)"
        lComando = lComando & " Values ("
        lComando = lComando & M46V999.gfFormatarCamposSQL(mCotacao) & ","
        lComando = lComando & M46V999.gfFormatarCamposSQL(mNossoTempor) & ","
        lComando = lComando & M46V999.gfFormatarCamposSQL(mData) & ","
        lComando = lComando & "' ',' ',' ',"
        lComando = lComando & M46V999.gfFormatarCamposSQL(mStatus) & ","
        lComando = lComando & "0,"
        lComando = lComando & M46V999.gfFormatarCamposSQL(mMaxDesc) & ","
        lComando = lComando & M46V999.gfFormatarCamposSQL(mSenha) & ")"
    Else
        lComando = "Update P0045400 set "
        lComando = lComando & " CCotacao = " & M46V999.gfFormatarCamposSQL(mCotacao)
        lComando = lComando & " ,Status = " & M46V999.gfFormatarCamposSQL(mStatus)
        lComando = lComando & " ,Senha = " & M46V999.gfFormatarCamposSQL(mSenha)
        lComando = lComando & " ,MaxDesc= " & M46V999.gfFormatarCamposSQL(mMaxDesc)
        lComando = lComando & " Where nosso_numero = " & M46V999.gfFormatarCamposSQL(mNossoTempor)
    End If
    
    If M46V999.gfExeSQL(bdConexao, lComando, lMensagem) <> 0 Then
        Call M46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If
    mfGraP0045400 = True
    
End Function

Private Function mfPosCampo(ByVal pOpcao As Byte, ByVal pNomCampo As String, ByVal pCampo, _
                            ByVal pTamanho As Integer, ByVal pDecimais As Byte, ByRef pMensagem As String) _
                            As Boolean
    'Fun��o: posiciona campo.

    'Par�metros de entrada...pOpcao......1 = posiciona campo num�rico com tamanho e casas decimais
    '                                        especificados (consiste e emite mensagem).
    '                                    2 = posiciona campo num�rico com tamanho e casas decimais
    '                                        especificados (preenche com zero se campo com erro).
    '                                    3 = indicador (S ou N).
    '                                    4 = campo alfanum�rico (preenche com branco se campo com erro).
    '                                    5 = campo num�rico - preenche com zero se registro incompleto ou
    '                                        campo n�o num�rico - considera 2 casas decimais.
    '                                    6 = campo alfanum�rico com conte�do num�rico - consiste e
    '                                        emite mensagem.
    '                        pNomCampo...nome do campo.
    '                        pCampo......campo a ser formatado e posicionado no registro.
    '                        pTamanho....tamanho do campo no registro de sa�da.
    '                        pDecimais...quantidade de casas decimais a serem consideradas.
    'Par�metros de sa�da.....pMensagem...mensagem.

    Dim lAuxiliar   As String   'Auxiliar.

    mfPosCampo = False

    '1. Consiste informa��o.
    If pTamanho = 0 Then
        pMensagem = "FTE0021 - Erro ao criar arquivo de transmiss�o (tamanho do campo " & pNomCampo & _
                    " inv�lido - zerado)."
        Exit Function
    End If

    '2. Formata campo a ser colocado no registro.
    Select Case pOpcao
        Case 1      'Posiciona campo num�rico (consiste e emite mensagem).
            If Not IsNumeric(pCampo) Then
                pMensagem = "FTE0022 - Erro ao criar arquivo de transmiss�o (campo " & pNomCampo & _
                            " n�o num�rico)."
                Exit Function
            End If
            Select Case pDecimais
                Case 0
                    lAuxiliar = Format$(pCampo, String(pTamanho, "0"))
                Case 1
                    lAuxiliar = Format$(pCampo * 10, String(pTamanho, "0"))
                Case 2
                    lAuxiliar = Format$(pCampo * 100, String(pTamanho, "0"))
                Case 3
                    lAuxiliar = Format$(pCampo * 1000, String(pTamanho, "0"))
                Case 4
                    lAuxiliar = Format$(pCampo * 10000, String(pTamanho, "0"))
                Case 5
                    lAuxiliar = Format$(pCampo * 100000, String(pTamanho, "0"))
                Case 6
                    lAuxiliar = Format$(pCampo * 1000000, String(pTamanho, "0"))
                Case Else
                    pMensagem = "FTE0023 - Erro ao criar arquivo de transmiss�o (quantidade de " & _
                                "casas decimais inv�lida ao posicionar campo " & pNomCampo & " - " & _
                                pDecimais & ")."
                    Exit Function
            End Select
            lAuxiliar = Right(lAuxiliar, pTamanho)
        Case 2      'Posiciona campo num�rico (preenche com zero se campo com erro).
            If IsNull(pCampo) Then
                lAuxiliar = String(pTamanho, "0")
            Else
                If IsNumeric(pCampo) Then
                    Select Case pDecimais
                        Case 0
                            lAuxiliar = Format$(pCampo, String(pTamanho, "0"))
                        Case 1
                            lAuxiliar = Format$(pCampo * 10, String(pTamanho, "0"))
                        Case 2
                            lAuxiliar = Format$(pCampo * 100, String(pTamanho, "0"))
                        Case 3
                            lAuxiliar = Format$(pCampo * 1000, String(pTamanho, "0"))
                        Case 4
                            lAuxiliar = Format$(pCampo * 10000, String(pTamanho, "0"))
                        Case 5
                            lAuxiliar = Format$(pCampo * 100000, String(pTamanho, "0"))
                        Case 6
                            lAuxiliar = Format$(pCampo * 1000000, String(pTamanho, "0"))
                        Case Else
                            pMensagem = "FTE0023 - Erro ao criar arquivo de transmiss�o (quantidade " & _
                                        "de casas decimais inv�lida ao posicionar campo " & pNomCampo & _
                                        " - " & pDecimais & ")."
                            Exit Function
                    End Select
                    lAuxiliar = Right(lAuxiliar, pTamanho)
                Else
                    lAuxiliar = String(pTamanho, "0")
                End If
            End If
        Case 4      'Posiciona campo alfanum�rico (preenche com branco se campo com erro).
            If M46V999.gfPreenchido(pCampo) Then
                Select Case Len(pCampo)
                    Case Is < pTamanho
                        lAuxiliar = pCampo
                        While Len(lAuxiliar) < pTamanho
                            lAuxiliar = lAuxiliar & " "
                        Wend
                    Case pTamanho
                        lAuxiliar = pCampo
                    Case Is > pTamanho
                        lAuxiliar = Left$(pCampo, pTamanho)
                End Select
            Else
                lAuxiliar = Space(pTamanho)
            End If
        Case 6      'Posiciona campo alfanum�rico com conte�do num�rico (consiste e emite mensagem).
            If Not M46V999.gfPreenchido(pCampo) Then
                pMensagem = "FTE0024 - Erro ao criar arquivo de transmiss�o (campo " & pNomCampo & _
                            " n�o preenchido)."
                Exit Function
            End If
            If Not IsNumeric(pCampo) Then
                pMensagem = "FTE0025 - Erro ao criar arquivo de transmiss�o (campo " & pNomCampo & _
                            " n�o num�rico)."
                Exit Function
            End If
            lAuxiliar = Right$(String(pTamanho, "0") & Trim(pCampo), pTamanho)
        Case Else
            pMensagem = "FTE0026 - Op��o inv�lida (posiciona campo no registro de transmiss�o - " & _
                        "op��o desconhecida - " & pOpcao & ")."
            Exit Function
    End Select

    '3. Acrescenta campo ao registro.
    mRegistro = mRegistro & lAuxiliar

    mfPosCampo = True
End Function

Private Function mfRetCampo(ByVal pOpcao As Byte, ByVal pNomCampo As String, ByRef pCampo As Variant, _
                            ByRef pIndex As Integer, ByVal pTamanho As Integer, ByVal pDecimais As Byte, ByRef pMensagem As String) _
                            As Boolean
    'Fun��o: Retorna o valor do campo na String Campo.

    'Par�metros de entrada...pOpcao......1 = posiciona campo num�rico com tamanho e casas decimais
    '                                        especificados (consiste e emite mensagem).
    '                                    2 = posiciona campo num�rico com tamanho e casas decimais
    '                                        especificados (preenche com zero se campo com erro).
    '                                    4 = campo alfanum�rico (preenche com branco se campo com erro).
    '                                    6 = campo alfanum�rico com conte�do num�rico - consiste e
    '                                        emite mensagem.
    '                        pNomCampo...nome do campo.
    '                        pCampo......campo a ser formatado e posicionado no registro.
    '                        pTamanho....tamanho do campo no registro de sa�da.
    '                        pDecimais...quantidade de casas decimais a serem consideradas.
    'Par�metros de sa�da.....pMensagem...mensagem.

    Dim lAuxiliar   As String   'Auxiliar.
    Dim lValorAux   As Double   'Auxiliar.

    mfRetCampo = False

    '1. Consiste informa��o.
    If pTamanho = 0 Then
        pMensagem = "FTE0021 - Erro ao retornar o valor do arquivo de transmiss�o (tamanho do campo " & pNomCampo & _
                    " inv�lido - zerado)."
        Exit Function
    End If

    lAuxiliar = Mid(mRegistro, pIndex, pTamanho)
    lValorAux = 0
    
    '2. Formata campo a ser colocado no registro.
    Select Case pOpcao
        Case 1, 2       'Posiciona campo num�rico (consiste e emite mensagem).
            If Not IsNumeric(lAuxiliar) Then
                pMensagem = "FTE0022 - Erro ao recuperar arquivo de transmiss�o (campo " & pNomCampo & _
                            " n�o num�rico). Valor do campo: """ & lAuxiliar & """. Posi��o: " & pIndex & ", Tamanho: " & pTamanho & "."
                Exit Function
            End If
            Select Case pDecimais
                Case 0
                    lValorAux = CDblx(lAuxiliar)
                Case 1
                    lValorAux = CDblx(lAuxiliar) / 10
                Case 2
                    lValorAux = CDblx(lAuxiliar) / 100
                Case 3
                    lValorAux = CDblx(lAuxiliar) / 1000
                Case 4
                    lValorAux = CDblx(lAuxiliar) / 10000
                Case 5
                    lValorAux = CDblx(lAuxiliar) / 100000
                Case 6
                    lValorAux = CDblx(lAuxiliar) / 1000000
                Case Else
                    pMensagem = "FTE0023 - Erro ao recuperar dados do arquivo de transmiss�o (quantidade de " & _
                                "casas decimais inv�lida ao posicionar campo " & pNomCampo & " - " & _
                                pDecimais & "). Valor do campo: """ & lAuxiliar & """. Posi��o: " & pIndex & ", Tamanho: " & pTamanho & "."
                    Exit Function
            End Select
            pCampo = lValorAux
            lAuxiliar = ""
        Case 4      'Posiciona campo alfanum�rico (preenche com branco se campo com erro).
            pCampo = lAuxiliar
        Case 6      'Posiciona campo alfanum�rico com conte�do num�rico (consiste e emite mensagem).
            pCampo = Right$(String(pTamanho, "0") & Trim(lAuxiliar), pTamanho)
        Case Else
            pMensagem = "FTE0026 - Op��o inv�lida (posiciona campo no registro de transmiss�o - " & _
                        "op��o desconhecida - " & pOpcao & ").  Valor do campo: """ & lAuxiliar & """. Posi��o: " & pIndex & ", Tamanho: " & pTamanho & "."
            Exit Function
    End Select

    '3. Acrescenta valores do tamanho ao contador.
    pIndex = pIndex + pTamanho

    mfRetCampo = True
    
End Function


Private Function CDblx(ByRef Valor As Variant) As Double
    If IsNumeric(Valor) Then
        Valor = Valor
    Else
        Valor = 0
    End If
    CDblx = Valor
End Function

Private Function mfRecPedido(ByVal pNossoNumero As String, _
                             ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                             Optional pGravar As Boolean = True, _
                             Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de dados comuns para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    '                        ........objeto de pedido.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
Err.Clear
mfRecPedido = False
On Error GoTo TrataErro

    '1. Grava registro de dados comuns.
   
    Dim lobjPedido            As clsA46V702A  'Objeto da classe clsA46V702A  (TAB_PED)
    Dim lobjstcPedido         As stcA46V702B  'Objeto da classe de estrutura stcA46V702B (TAB_PED)
    Dim lobjstcPedidoDel      As stcA46V702B  'Objeto da classe de estrutura stcA46V702B (TAB_PED)
    Dim lIndice               As Integer
    Dim lAux                  As Variant
  
    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "001" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED"
        GoTo SAIDA
    End If
            
    If Not IsMissing(pStcPedido) Then
        Set lobjstcPedido = pStcPedido
    End If
    
    If lobjstcPedido Is Nothing Then
       Set lobjstcPedido = New stcA46V702B
       Set lobjstcPedido.M46V999 = M46V999
    End If
            
        
    lIndice = lIndice + 3
    With lobjstcPedido
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Pedido n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        M46V999.gCod_Ramo = .Cod_Ramo
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "N�mero da ap�lice antiga", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol_An = lAux
        If mfRetCampo(2, "C�digo do registro", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Reg = lAux
        If mfRetCampo(2, "C�digo do banco de cobran�a", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Bco_Cobr = lAux
        If mfRetCampo(2, "C�digo da ag�ncia angar", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Agen_Angar = lAux
        If mfRetCampo(2, "C�digo da ag�ncia Indicr", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Agen_Indicr = lAux
        If mfRetCampo(2, "N�mero da conta", lAux, lIndice, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Cta_Indicr = lAux
        If mfRetCampo(2, "C�digo da unidade de produto", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Unid_Prod = lAux
        If mfRetCampo(2, "C�digo do produtor", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Prod = lAux
        If mfRetCampo(2, "C�digo da ag�ncia de d�bito", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Agen_Bas_Cor = lAux
        If mfRetCampo(2, "N�mero da conta de d�bito", lAux, lIndice, 15, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If lAux = "999999999999999" Then
            lAux = 9999999
        End If
        .Num_Cta_Bas_Cor = lAux
        If mfRetCampo(2, "C�digo do Segurado", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Segurado = lAux
        If mfRetCampo(4, "Nome do segurado", lAux, lIndice, 60, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Segurado = lAux
        If mfRetCampo(4, "Complemento do nome do segurado", lAux, lIndice, 60, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dsc_Compl_Segurado = lAux
        If mfRetCampo(6, "N�mero do CGC-CPF do segurado", lAux, lIndice, 14, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Cgc_Cpf = lAux
        If mfRetCampo(4, "Tipo de logradouro", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Tip_Logr = lAux
        If mfRetCampo(4, "Logradouro", lAux, lIndice, 50, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Logr = lAux
        If mfRetCampo(4, "Complemento do logradouro", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dsc_Compl_Logr = lAux
        If mfRetCampo(4, "Bairro", lAux, lIndice, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Bairro = lAux
        If mfRetCampo(4, "Cidade", lAux, lIndice, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Cid = lAux
        If mfRetCampo(4, "Sigla UF", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Sig_UF = lAux
        If mfRetCampo(2, "CEP", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_CEP = lAux
        If mfRetCampo(2, "Telefone", lAux, lIndice, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Fone = val(lAux)
        If mfRetCampo(2, "C�digo Estip", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Estip = lAux
        If mfRetCampo(2, "Coeficiente de comiss�o Pad", lAux, lIndice, 5, 4, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Coef_Cms_Pad = lAux
        If mfRetCampo(2, "Data de in�cio da vig�ncia", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Ini_Vig = lAux
        If mfRetCampo(2, "Data de t�rmino da vig�ncia", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Ter_Vig = lAux
        If mfRetCampo(2, "Data de emiss�o", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Emis = lAux
        If mfRetCampo(2, "C�digo da moeda", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Moed = lAux
        If mfRetCampo(2, "Percentual de juros do fracionamento", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Juro_Frac = lAux
        If mfRetCampo(2, "Quantidade de parcelas", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Qtd_Parc = lAux
        If mfRetCampo(2, "Quantidade de anexos", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Qtd_Anexo = lAux
        If mfRetCampo(2, "C�digo de identifica��o do tipo de prazo", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Praz = lAux
        If mfRetCampo(2, "C�digo do pr�mio m�nimo de endosso", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Pr_Min_Endo = lAux
        If mfRetCampo(2, "C�digo do pr�mio m�nimo de parcela", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Pr_Min_Parc = lAux
        If mfRetCampo(2, "Valor pr�mio antecipado", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Pr_Antec = lAux
        If mfRetCampo(2, "Percentual adicioanl na forma de pagamento", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Adic_Form_Pag = lAux
        If mfRetCampo(2, "Percentual de desconto na forma de pagamento", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Desc_Form_Pag = lAux
        If mfRetCampo(2, "Valor da IS", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_IS = lAux
        If mfRetCampo(2, "Valor da IS 2", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_IS_2 = lAux
        If mfRetCampo(2, "Valor do pr�mio", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .VAL_PR = lAux
        If mfRetCampo(2, "Valor do pr�mio ref tarif", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Pr_Ref_Tarif = lAux
        If mfRetCampo(2, "Valro do pr�mio ref  resg", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Pr_Ref_Resg = lAux
        If mfRetCampo(2, "Data de in�cio Viag", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Ini_Viag = lAux
        If mfRetCampo(2, "C�digo do tipo de libera��o", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Tip_Liber = lAux
        If mfRetCampo(4, "C�digo identificador de altera��o de corretor", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Alt_Corr = lAux
        If mfRetCampo(4, "C�digo do usu�rio de digita��o", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_User_Digit = lAux
        If mfRetCampo(2, "Data da digita��o", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .DAT_DIGIT = lAux
        If mfRetCampo(2, "Hora da digita��o", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .HOR_DIGIT = lAux
        If mfRetCampo(4, "C�digo do usu�rio de libera��o", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_User_Liber = lAux
        If mfRetCampo(2, "Data da libera��o", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Liber = lAux
        If mfRetCampo(2, "Hora da libera��o", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Hor_Liber = lAux
        If mfRetCampo(2, "C�digo da situa��o do pedido", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Sit_Ped = lAux
        If mfRetCampo(2, "C�digo do tipo de documento", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Tip_Doc = lAux
        If mfRetCampo(2, "C�digo do dizer", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Dizer = lAux
        If mfRetCampo(2, "Quantidade de c�pias", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Qtd_Copia = lAux
        If mfRetCampo(2, "Valor do custo de emiss�o", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Cust_Emis = lAux
        If mfRetCampo(2, "Valor de juros", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Juro = lAux
        If mfRetCampo(2, "Valor IOF", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_IOF = lAux
        If mfRetCampo(2, "N�mero da proposta", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Prop = lAux
        If mfRetCampo(4, "C�digo da posi��o atual", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Posi_Atual = lAux
        If mfRetCampo(2, "Data do in�cio do movimento", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Mov_Ini = lAux
        If mfRetCampo(2, "Data do t�rmino do movimento", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Mov_Ter = lAux
        If mfRetCampo(2, "Tipo de custo de emiss�o", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tip_Cust_Emis = lAux
        If mfRetCampo(2, "Valor do cambio na emiss�o", lAux, lIndice, 11, 5, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Cambio_Emis = lAux
        If mfRetCampo(4, "Iden", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Iden = lAux
        If mfRetCampo(2, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
        If mfRetCampo(2, "C�digo do produto", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Produto = lAux
        If mfRetCampo(2, "C�digo do protocolo", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Protocolo = lAux
        If mfRetCampo(2, "C�digo da sequencia", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Sequencia = lAux
        If mfRetCampo(2, "Tipo de pessoa", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tip_Pessoa = lAux
        If mfRetCampo(2, "C�digo do ramo de atividade", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo_Ativ = lAux
        If mfRetCampo(2, "C�digo do tipo de documento", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Doc = lAux
        If mfRetCampo(4, "N�mero do documento", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Doc = lAux
        If mfRetCampo(4, "�rg�o expedidor", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Org_Exp = lAux
        If mfRetCampo(2, "Data da emiss�o do documento (yyyymmdd)", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Emis_Doc = lAux
        If mfRetCampo(4, "Nome da empresa onde trabalha", lAux, lIndice, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Empresa = lAux
        If mfRetCampo(4, "Sexo do segurado: M -masculino , F - Feminino.", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Sexo = lAux
        If mfRetCampo(2, "Data do Primeira C�lculo.", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Pri_Calculo = lAux
        If mfRetCampo(2, "�ltima Tarifa.", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Ult_Tarifa = lAux
        If mfRetCampo(2, "Codigo de Unicidade.", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Unicidade = lAux
        If mfRetCampo(2, "Codigo do Digito da Agencia BasCor.", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Dig_Agen_Bas_Cor = lAux
        If mfRetCampo(2, "Codigo da Vers�o da Unicidade.", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ver_Unici = lAux

        If mfRetCampo(2, "Renda/Faturamento.", lAux, lIndice, 18, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_RendaFatu = lAux
        If mfRetCampo(2, "Percentual de LMI.", lAux, lIndice, 18, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Desc_LMI = lAux
        If mfRetCampo(2, "Se eh LMI.", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .LMI_Unico = lAux
        If mfRetCampo(4, "N�mero da Cotacao de LMI.", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Cotac_LMI = lAux

        
        '1001046 - Telefone do Segurado 8 Digitos
        If Len(mRegistro) >= 872 Then
            If mfRetCampo(2, "Telefone", lAux, lIndice, 8, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Num_Fone = lAux
        'Tarifa 10.00 - campo ULT_TARIFA alterado de 4 para 5
            If mfRetCampo(2, "�ltima Tarifa.", lAux, lIndice, 5, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Ult_Tarifa = lAux
        End If

        'Projeto Condi��es Gerais
        If Len(mRegistro) >= 878 Then
            If mfRetCampo(2, "Indicador de condi��es gerais", lAux, lIndice, 1, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Ind_CondGerais = lAux
        End If
       
        'Tarifa 10.00 - campo Cod_Produto alterado de 3 para 4
        If Len(mRegistro) >= 879 Then
            If mfRetCampo(2, "C�digo do produto", lAux, lIndice, 4, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Cod_Produto = lAux
        End If
        
        
        If Len(mRegistro) > 882 Then
            If mfRetCampo(2, "Ind_Reentrada", lAux, lIndice, 1, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Ind_Reentrada = lAux
        End If
        
        If Len(mRegistro) > 883 Then
            If mfRetCampo(4, "Nosso_Numero_Original", lAux, lIndice, 20, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Nosso_Numero_Original = lAux
        End If

        If mfRetCampo(2, "Numero de PI Emissao", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If

        '1200178 - Telefone 9 Digitos
        If Len(mRegistro) > 913 Then
            If mfRetCampo(2, "Telefone", lAux, lIndice, 9, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Num_Fone = lAux
        End If
        
'        .Val_RendaFatu = 0
 '       If Mid(mRegistro, lIndice, 18) <> "" Then
  '          If mfRetCampo(2, "Valor da renda/Faturamento.", lAux, lIndice, 18, 2, pMensagem) = False Then
   '             GoTo SAIDA
    '        End If
     '       .Val_RendaFatu = lAux
      '  End If
        
    End With
    
    Set lobjPedido = New clsA46V702A
    Set lobjPedido.M46V999 = M46V999
    
    'Temporario
    lobjstcPedido.nosso_numero = mNossoTempor
    Set lobjstcPedidoDel = lobjstcPedido
    
    If pGravar = True Then
        Set lobjPedido.ConexaoBD = bdBasDados
        Dummy = lobjPedido.mfExcluir(lobjstcPedidoDel, pPropTabela, pNomeTabImport, Empty)
        If Not lobjPedido.mfIncluir(lobjstcPedido, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjPedido.UltimoErro
            GoTo SAIDA
        End If
        If Not IsMissing(pStcPedido) Then
            Set pStcPedido = lobjstcPedido
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            Set pStcPedido = lobjstcPedido
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
    
'================================
'******Fun��o realizada com sucesso
    mfRecPedido = True
'================================

    ''Dummy = mpVerificaSerasa(lobjstcPedido.Num_PI, 0, lobjstcPedido.Cod_Ramo, lobjstcPedido.Num_Cgc_Cpf, pMensagem)

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecPedido Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjPedido = Nothing
    Set lobjstcPedido = Nothing
    Return

End Function

Private Function mfRecCtrlEmis(ByVal pNossoNumero As String, _
                               ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                               Optional pGravar As Boolean = True, _
                               Optional pStcPedido As stcA46V702B) As Boolean
                               
    'Fun��o: grava registro de controle de emiss�o

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
    
    '1. Grava registro de dados de controle de emiss�o
        
Dim lobjCtrlEmis            As clsA46V712A      'Objeto da classe clsA46V712A (TAB_CTRL_EMIS)
Dim lobjstcCtrlEmis         As stcA46V712B      'Objeto da classe de estrutura stcA46V712B (TAB_CTRL_EMIS)
Dim lIndice                 As Integer
Dim lAux                    As Variant
  
Err.Clear
mfRecCtrlEmis = False
On Error GoTo TrataErro

    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "002" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_CTRL_EMIS"
        GoTo SAIDA
    End If
            
    Set lobjstcCtrlEmis = New stcA46V712B
        
    lIndice = lIndice + 3
        
    With lobjstcCtrlEmis
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Controle de Emiss�o n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
        If mfRetCampo(2, "N�mero do bloqueto sem d�gito verificador", lAux, lIndice, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Bloq = lAux
        If mfRetCampo(2, "C�digo da unidade produtiva", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .COD_UNID = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "Data de in�cio da vig�ncia", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Ini_Vig = lAux
        If mfRetCampo(2, "Data de t�rmino da vig�ncia", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Ter_Vig = lAux
        If mfRetCampo(2, "C�digo do corretor 1", lAux, lIndice, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Corr = lAux
        If mfRetCampo(2, "Percentual do corretor 1", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Corr = lAux
        If mfRetCampo(2, "C�digo do corretor 2", lAux, lIndice, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Corr_2 = lAux
        If mfRetCampo(2, "Percentual do corretor 2", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Corr_2 = lAux
        If mfRetCampo(2, "C�digo do corretor 3", lAux, lIndice, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Corr_3 = lAux
        If mfRetCampo(2, "Percentual do corretor 3", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Corr_3 = lAux
        If mfRetCampo(4, "Nome do segurado", lAux, lIndice, 60, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Seg = lAux
        If mfRetCampo(2, "Quantidade de anexos", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Qtd_Anex = lAux
        If mfRetCampo(2, "N�mero da ap�lice anterior", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol_Ant = lAux
        If mfRetCampo(4, "Observa��o", lAux, lIndice, 42, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dsc_Obs_Prop = lAux
        If mfRetCampo(4, "Tipo de cota��o", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tip_Cotac = lAux
        If mfRetCampo(4, "N�mero da cota��o", lAux, lIndice, 12, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Cotac = lAux
        If mfRetCampo(2, "Percentual de juros negociafo", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Juros_Neg = lAux
        If mfRetCampo(4, "C�digo da etiqueta", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Envi_Doc = lAux
        If mfRetCampo(2, "Data da digita��o", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .DAT_DIGIT = lAux
        If mfRetCampo(2, "Hora da digita��o", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .HOR_DIGIT = lAux
        If mfRetCampo(2, "C�digo do status da a��o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Stat_Acao = lAux
        If mfRetCampo(4, "C�digo do usu�rio de digitaliza��o", lAux, lIndice, 12, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_User_Digit = lAux
        If mfRetCampo(2, "Data da digita��o mainframe", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Digit_MF = lAux
        If mfRetCampo(2, "Hora da digita��o mainframe", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Hor_Digit_MF = lAux
        If mfRetCampo(4, "C�digo do usu�rio de digitaliza��o mainframe", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_User_Digit_MF = lAux
        If mfRetCampo(2, "Valor do pr�mio", lAux, lIndice, 9, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Pr_Tot = lAux
        If mfRetCampo(2, "Classe de norma", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Classe = lAux
        If mfRetCampo(2, "C�digo de cancelamento do cobran�a", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Cobr = lAux
        If mfRetCampo(2, "Numero da proposta", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Prop = lAux
        If mfRetCampo(2, "Data de cancelamento do documento", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Canc_Apol = lAux
        If mfRetCampo(2, "Tipo de documento", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Tip_Doc = lAux
        If mfRetCampo(2, "Data de recebimento na companhia", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Rec_Cia = lAux
        If mfRetCampo(2, "Numero de refer�ncia do corretor", lAux, lIndice, 15, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ref_Corr = lAux
        If mfRetCampo(2, "Data de emiss�o do documento", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Emis = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "Valor do pr�mio emitido", lAux, lIndice, 9, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Pr_Emis = lAux
        If mfRetCampo(2, "Valor do pr�mio l�quido", lAux, lIndice, 9, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Pr_Liq = lAux
        If mfRetCampo(2, "Percentual total de cosseguro", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Tot_Coss = lAux
        If mfRetCampo(2, "Quantidade de parcelas", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Qtd_Parc = lAux
        If mfRetCampo(2, "Status da resposta", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Stat_Resp = lAux
        If mfRetCampo(2, "C�digo identificador de envio ao SAC", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_SAC = lAux
        If mfRetCampo(2, "C�digo do produtor", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Prod = lAux
        If mfRetCampo(2, "Valor da IS", lAux, lIndice, 15, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_IS = lAux
        If mfRetCampo(2, "Percentual do desconto comercial", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Desc_Coml = lAux
        If mfRetCampo(6, "Numero do cgc/cpf", lAux, lIndice, 14, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Cgc_Cpf = lAux
        If mfRetCampo(2, "Tipo de pessoa", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Pessoa = lAux
        If mfRetCampo(2, "Pagamento antecipado", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Pag_Antec = lAux
        If mfRetCampo(2, "D�gito do bloqueto", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .DIG_BLOQUETO = lAux

        'TAB_CTRL_EMIS_COMPL
        If mfRetCampo(4, "Email", lAux, lIndice, 150, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Email_Contato = lAux
        If mfRetCampo(2, "Indicador prolabore", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Ind_Prolabore = lAux
        If mfRetCampo(2, "Indicador agenciamento", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Ind_Agenciamento = lAux
        If mfRetCampo(4, "Parcelas a vista", lAux, lIndice, 100, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Parcelas_Impressao_Avista = lAux
        If mfRetCampo(4, "Parcelas a 30 dias", lAux, lIndice, 100, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Parcelas_Impressao_30dias = lAux
        If mfRetCampo(2, "Apolice Impressa ou n�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Ind_Impr_Apol = lAux
        If mfRetCampo(2, "C�digo do Banco do Boleto (Quando demais d�bito)", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Bco_Boleto = lAux
        If mfRetCampo(2, "Agravo do corretor", lAux, lIndice, 5, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Desc_Agrav = lAux
        If mfRetCampo(2, "Origem da Renova��o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .cod_origem = lAux
        If mfRetCampo(2, "Numero da Ap�lice de Renova��o Mar�tima", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol_Mar = lAux
        If mfRetCampo(4, "Numero de logradouro do segurado", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Logr_Segurado = lAux
        If mfRetCampo(4, "Complemento do Logradouro do Segurado", lAux, lIndice, 13, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dsc_Compl_Segurado = lAux
        If mfRetCampo(4, "Numero do Logradouro de Cobran�a", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Logr_Cobranca = lAux
        If mfRetCampo(4, "Complemento do Logradouro de Cobran�a", lAux, lIndice, 13, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dsc_Compl_Cobranca = lAux
        If mfRetCampo(2, "Estado civil", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Estado_Civil = lAux
        If mfRetCampo(2, "Data de nascimento", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .DAT_Nascimento = lAux
        
        If (.Cod_Ramo = 300 Or .Cod_Ramo = 620 Or .Cod_Ramo = 710) And IsNumeric(Mid(mRegistro, lIndice, 3)) = True Then
            If mfRetCampo(2, "C�digo Banco cr�dito", lAux, lIndice, 3, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Cod_Banco_Cre = lAux
            If mfRetCampo(2, "C�digo Ag�ncia cr�dito", lAux, lIndice, 4, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Cod_Agenc_Cre = lAux
            If mfRetCampo(2, "C�digo Conta cr�dito", lAux, lIndice, 15, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Cod_Conta_Cre = lAux
            If mfRetCampo(6, "CPF/CNPJ cr�dito", lAux, lIndice, 14, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Cpf_Cnpj_Cre = lAux
            If mfRetCampo(4, "D�gito Ag�ncia cr�dito", lAux, lIndice, 1, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Dig_Agenc_Cre = lAux
            If mfRetCampo(2, "�ndice condom�nio", lAux, lIndice, 1, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Ind_Cond_Constituido = lAux
            If mfRetCampo(2, "�ndice IPC", lAux, lIndice, 1, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Indice_IPC = lAux
            If mfRetCampo(2, "Suprimir batente", lAux, lIndice, 1, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Suprimir_Batente = lAux
            If mfRetCampo(2, "Valor verba comercial", lAux, lIndice, 8, 2, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Val_VerbaCom = lAux
            If mfRetCampo(2, "Tipo restitui��o cr�dito", lAux, lIndice, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Tip_Res_Cre = lAux
            If mfRetCampo(4, "D�gito da Conta para restitui��o", lAux, lIndice, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Cod_DIG_Conta = lAux

        End If
    End With

    Set lobjCtrlEmis = New clsA46V712A
    Set lobjCtrlEmis.M46V999 = M46V999

    'Temporario
    lobjstcCtrlEmis.nosso_numero = mNossoTempor
    
    If pGravar = True Then
        Set lobjCtrlEmis.ConexaoBD = bdBasDados
        Set lobjstcCtrlEmis.ObjPai = pStcPedido
        If Not lobjCtrlEmis.mfIncluir(lobjstcCtrlEmis, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjCtrlEmis.UltimoErro
            GoTo SAIDA
        End If
        If Not IsMissing(pStcPedido) Then
            Set pStcPedido.objstcCtrlEmis = lobjstcCtrlEmis
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            Set pStcPedido.objstcCtrlEmis = lobjstcCtrlEmis
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecCtrlEmis = True
'================================
    
SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecPedido Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjCtrlEmis = Nothing
    Set lobjstcCtrlEmis = Nothing
    Return
    
End Function
Private Function mfRecCorretores(ByVal pNossoNumero As String, _
                                 ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                                 Optional pGravar As Boolean = True, _
                                 Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de corretores para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
   
    '1. Grava registro de dados corretores.
Dim lobjCorretor            As clsA46V705A      'Objeto da classe clsA46V705A (TAB_PED_CORR)
Dim lobjstcCorretor         As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
Dim lIndice                 As Integer
Dim lAux                    As Variant
  
Err.Clear
mfRecCorretores = False
On Error GoTo TrataErro

    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "003" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_CORR"
        GoTo SAIDA
    End If
            
    Set lobjstcCorretor = New stcA46V705B
    Set lobjstcCorretor.M46V999 = M46V999
    
    lIndice = lIndice + 3
        
    With lobjstcCorretor
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Corretor n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "C�digo do corretor", lAux, lIndice, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Corr = lAux
        If mfRetCampo(2, "Percentual do corretor", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Corr = lAux
        If mfRetCampo(2, "N�mero da ordem do corretor", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ordem_Corr = lAux
        If mfRetCampo(2, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
    End With

    Set lobjCorretor = New clsA46V705A
    Set lobjCorretor.M46V999 = M46V999

    'Temporario
    
    lobjstcCorretor.nosso_numero = mNossoTempor
    
    If pGravar = True Then
        Set lobjCorretor.ConexaoBD = bdBasDados
        Set lobjstcCorretor.ObjPai = pStcPedido
        If Not lobjCorretor.mfIncluir(lobjstcCorretor, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjCorretor.UltimoErro
            GoTo SAIDA
        End If
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoPedCorr.Add Item:=lobjstcCorretor, Key:=lobjstcCorretor.Class_ID
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoPedCorr.Add Item:=lobjstcCorretor, Key:=lobjstcCorretor.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
mfRecCorretores = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecPedido Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjCorretor = Nothing
    Set lobjstcCorretor = Nothing
    Return
    
End Function

Private Function mfRecTextoEndosso(ByVal pNossoNumero As String, _
                                   ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                                   Optional pGravar As Boolean = True, _
                                   Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de corretores para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
Dim lobjTextoEndo           As clsA46V722A      'Objeto da classe clsA46V722A (TAB_PED_DIZER)
Dim lobjstcTextoEndo        As stcA46V722B      'Objeto da classe de estrutura stcA46V722B (TAB_PED_DIZER)
Dim lIndice                 As Integer
Dim lAux                    As Variant
    
Err.Clear
mfRecTextoEndosso = False
On Error GoTo TrataErro
  
    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "004" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_DIZER"
        GoTo SAIDA
    End If
            
    Set lobjstcTextoEndo = New stcA46V722B
    
    lIndice = lIndice + 3
        
    With lobjstcTextoEndo
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Texto de Endosso n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "C�digo da Linha", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_SEQ = lAux
        If mfRetCampo(4, "Conteudo da Linha", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dsc_Dizer_Endo = lAux
    End With
    
    Set lobjTextoEndo = New clsA46V722A
    Set lobjTextoEndo.M46V999 = M46V999

    lobjstcTextoEndo.nosso_numero = mNossoTempor

    If pGravar = True Then
        Set lobjTextoEndo.ConexaoBD = bdBasDados
        Set lobjstcTextoEndo.ObjPai = pStcPedido
        If Not lobjTextoEndo.mfIncluir(lobjstcTextoEndo, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjTextoEndo.UltimoErro
            GoTo SAIDA
        End If
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoTextoEndo.Add Item:=lobjstcTextoEndo, Key:=lobjstcTextoEndo.Class_ID
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoTextoEndo.Add Item:=lobjstcTextoEndo, Key:=lobjstcTextoEndo.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecTextoEndosso = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecTextoEndosso Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjTextoEndo = Nothing
    Set lobjstcTextoEndo = Nothing
    Return
    
End Function

Private Function mfRecVenc(ByVal pNossoNumero As String, _
                           ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                           Optional pGravar As Boolean = True, _
                           Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de vencimento para transmiss�o.

    'Par�metros de entrada...pNossoNumero.........Nosso N�mero do registro principal.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
Dim lObjstcVenc      As stcA46V713B      'Objeto da classe de estrutura stcA46V713B (TAB_PED_VENC)
Dim lobjVenc         As clsA46V713A      'Objeto da classe clsA46V713A (TAB_PED_VENC)
Dim lIndice          As Integer
Dim lAux             As Variant

Err.Clear
mfRecVenc = False
On Error GoTo TrataErro

    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "005" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_VENC"
        GoTo SAIDA
    End If
            
    Set lObjstcVenc = New stcA46V713B
    Set lObjstcVenc.M46V999 = M46V999

    lIndice = lIndice + 3
    With lObjstcVenc
    
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Vencimento n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "N�mero da parcela", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Parc = lAux
        If mfRetCampo(2, "Data do vencimento", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .DAT_VENC = lAux
    End With
        
    Set lobjVenc = New clsA46V713A
    Set lobjVenc.M46V999 = M46V999

    'Temporario
    lObjstcVenc.nosso_numero = mNossoTempor
    
    If pGravar = True Then
        Set lobjVenc.ConexaoBD = bdBasDados
        Set lObjstcVenc.ObjPai = pStcPedido
        If Not lobjVenc.mfIncluir(lObjstcVenc, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjVenc.UltimoErro
            GoTo SAIDA
        End If
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoPedVenc.Add Item:=lObjstcVenc, Key:=lObjstcVenc.Class_ID
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoPedVenc.Add Item:=lObjstcVenc, Key:=lObjstcVenc.Class_ID
        Else
            pMensagem = "mfRecVenc Erro: Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecVenc = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecVenc Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjVenc = Nothing
    Set lObjstcVenc = Nothing
    Return
    
End Function

Private Function mfRecJuros(ByVal pNossoNumero As String, _
                            ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                            Optional pGravar As Boolean = True, _
                            Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de juros para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
Dim lobjJuros               As clsA46V711A      'Objeto da classe clsA46V711A (TAB_PED_JUROS)
Dim lobjstcJuros            As stcA46V711B      'Objeto da classe de estrutura stcA46V711B (TAB_PED_JUROS)
        
Dim lIndice                 As Integer
Dim lAux                    As Variant
  
Err.Clear
mfRecJuros = False
On Error GoTo TrataErro

    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "006" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_JUROS"
        GoTo SAIDA
    End If
            
    Set lobjstcJuros = New stcA46V711B
    
    lIndice = lIndice + 3
        
    With lobjstcJuros
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Juros n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "N�mero da parcela", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Parcela = lAux
        If mfRetCampo(2, "Valor do juros para pagamento antecipado", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Ant = lAux
        If mfRetCampo(2, "Valor do juros para pagamento postecipado", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Pos = lAux
    End With
    
    Set lobjJuros = New clsA46V711A
    Set lobjJuros.M46V999 = M46V999

    lobjstcJuros.nosso_numero = mNossoTempor
    
    If pGravar = True Then
        Set lobjJuros.ConexaoBD = bdBasDados
        Set lobjstcJuros.ObjPai = pStcPedido
        If Not lobjJuros.mfIncluir(lobjstcJuros, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjJuros.UltimoErro
            GoTo SAIDA
        End If
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoPedJuros.Add Item:=lobjstcJuros, Key:=lobjstcJuros.Class_ID
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoPedJuros.Add Item:=lobjstcJuros, Key:=lobjstcJuros.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecJuros = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecJuros Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjJuros = Nothing
    Set lobjstcJuros = Nothing
    Return

End Function

Private Function mfRecCosseguro(ByVal pNossoNumero As String, _
                                ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                                Optional pGravar As Boolean = True, _
                                Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de cosseguros para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    '                        pObjstcPedido........objeto de pedido.
    '                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
Dim lobjstcCosseguro        As stcA46V706B     'Objeto da classe de estrutura stcA46V706B (TAB_PED_COSS)
Dim lobjCosseguro           As clsA46V706A     'Objeto da classe clsA46V706B (TAB_PED_COSS)
Dim lIndice                 As Integer
Dim lAux                    As Variant
        
Err.Clear
mfRecCosseguro = False
On Error GoTo TrataErro
    
    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "007" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_COSS"
        GoTo SAIDA
    End If
            
    Set lobjstcCosseguro = New stcA46V706B
    Set lobjstcCosseguro.M46V999 = M46V999
    
    lIndice = lIndice + 3
        
    With lobjstcCosseguro
    
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Cosseguro n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "C�digo do Cosseguro", lAux, lIndice, 7, 0, pMensagem) = False Then ' simony  trocar para 3
            GoTo SAIDA
        End If
        .Cod_Coss = Mid(Trim(lAux), 1, 3)
        If mfRetCampo(2, "Percentual do cosseguro", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Coss = lAux
        If mfRetCampo(2, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
    End With

    Set lobjCosseguro = New clsA46V706A
    Set lobjCosseguro.M46V999 = M46V999

    lobjstcCosseguro.nosso_numero = mNossoTempor

    If pGravar = True Then
        Set lobjCosseguro.ConexaoBD = bdBasDados
        Set lobjstcCosseguro.ObjPai = pStcPedido
        If Not lobjCosseguro.mfIncluir(lobjstcCosseguro, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjCosseguro.UltimoErro
            GoTo SAIDA
        End If
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoPedCoss.Add Item:=lobjstcCosseguro, Key:=lobjstcCosseguro.Class_ID
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoPedCoss.Add Item:=lobjstcCosseguro, Key:=lobjstcCosseguro.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
   
'================================
'******Fun��o realizada com sucesso
    mfRecCosseguro = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecCosseguro Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjCosseguro = Nothing
    Set lobjstcCosseguro = Nothing
    Return

End Function

Private Function mfRecEndereco(ByVal pNossoNumero As String, _
                               ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                               Optional pGravar As Boolean = True, _
                               Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de endere�o para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
Dim lobjEndereco            As clsA46V707A      'Objeto da classe clsA46V707A (TAB_PED_ENDER)
Dim lobjstcEndereco         As stcA46V707B      'Objeto da classe de estrutura stcA46V707B (TAB_PED_ENDER)

Dim lIndice                 As Integer
Dim lAux                    As Variant
    
Err.Clear
mfRecEndereco = False
On Error GoTo TrataErro

    '1. Grava registro
    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "008" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_ENDER"
        GoTo SAIDA
    End If
            
    Set lobjstcEndereco = New stcA46V707B
    Set lobjstcEndereco.M46V999 = M46V999

    lIndice = lIndice + 3
        
    With lobjstcEndereco
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Endere�o n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(4, "Tipo de logradouro", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Tip_Logr = lAux
        If mfRetCampo(4, "Logradouro", lAux, lIndice, 50, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Logr = lAux
        If mfRetCampo(4, "Complemento do logradouro", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dsc_Compl_Logr = lAux
        If mfRetCampo(4, "Bairro", lAux, lIndice, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Bairro = lAux
        If mfRetCampo(4, "Cidade", lAux, lIndice, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Cid = lAux
        If mfRetCampo(4, "Sigla da UF", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Sig_UF = lAux
        If mfRetCampo(2, "CEP", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_CEP = lAux
        If mfRetCampo(2, "N�mero do telefone", lAux, lIndice, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Fone = lAux
        If mfRetCampo(2, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
'Ficha 1000307
        If Len(mRegistro) >= 222 Then
            If mfRetCampo(2, "C�digo do Registro do conselho", lAux, lIndice, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Cod_Reg_Cons = lAux
            If mfRetCampo(4, "N�mero do Registro do conselho", lAux, lIndice, 10, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Num_Reg_Cons = lAux
        End If
        
        '1001046 - Telefone do Segurado 8 Digitos
        If Len(mRegistro) >= 230 Then
            If mfRetCampo(2, "N�mero do telefone", lAux, lIndice, 9, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Num_Fone = lAux
        End If
        
    End With

    Set lobjEndereco = New clsA46V707A
    Set lobjEndereco.M46V999 = M46V999

    lobjstcEndereco.nosso_numero = mNossoTempor

    If pGravar = True Then
        Set lobjEndereco.ConexaoBD = bdBasDados
        Set lobjstcEndereco.ObjPai = pStcPedido
        If Not lobjEndereco.mfIncluir(lobjstcEndereco, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjEndereco.UltimoErro
            GoTo SAIDA
        End If
        If Not IsMissing(pStcPedido) Then
            Set pStcPedido.objstcPedEndereco = lobjstcEndereco
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            Set pStcPedido.objstcPedEndereco = lobjstcEndereco
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecEndereco = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecEndereco Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjEndereco = Nothing
    Set lobjstcEndereco = Nothing
    Return
End Function


Private Function mfRecItens(ByVal pNossoNumero As String, _
                            ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                            Optional pGravar As Boolean = True, _
                            Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de itens para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    '                        pObjstcPedido........objeto de pedido.
    '                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
Dim lobjItem                As clsA46V708A      'Objeto da classe clsA46V708A (TAB_PED_LOC)
Dim lobjstcItem             As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
Dim lIndice                 As Integer
Dim lAux                    As Variant
    
Err.Clear
mfRecItens = False
On Error GoTo TrataErro
    '1. Grava registro de dados de inspe��o.
    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "009" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_LOC"
        GoTo SAIDA
    End If
            
    Set lobjstcItem = New stcA46V708B
    Set lobjstcItem.M46V999 = M46V999

    lIndice = lIndice + 3
        
    With lobjstcItem
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro da Locais de Risco n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "N�mero do Item", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Item = lAux
        If mfRetCampo(4, "Tipo de logradouro - Risco", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Tip_Logr_Risco = lAux
        If mfRetCampo(4, "Logradouro - Risco", lAux, lIndice, 50, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Logr_Risco = lAux
        If mfRetCampo(4, "Complemento do logradouro - Risco", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dsc_Compl_Risco = lAux
        If mfRetCampo(4, "Nome do bairro - Risco", lAux, lIndice, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Bairro_Risco = lAux
        If mfRetCampo(4, "Nome dad cidade - Risco", lAux, lIndice, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Cid_Risco = lAux
        If mfRetCampo(4, "Sigla UF - Risco", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Sig_UF_Risco = lAux
        If mfRetCampo(2, "CEP - Risco", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Cep_Risco = lAux
        If mfRetCampo(2, "Percentual desconto - Apar", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Desc_Apar = lAux
        If mfRetCampo(2, "Percentual desconto - Progres", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Desc_Progres = lAux
        If mfRetCampo(2, "Percentual IS - VR", lAux, lIndice, 5, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_IS_Vr = lAux
        If mfRetCampo(2, "Valor do coeficiente IS - VR", lAux, lIndice, 5, 4, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Coef_IS_Vr = lAux
        If mfRetCampo(2, "Valor VR", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Vr = lAux
        If mfRetCampo(2, "Tab clasf", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tab_Clasf = lAux
        If mfRetCampo(4, "Codigo Clasf", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Clasf = lAux
        If mfRetCampo(4, "Codigo do complemento Clasf", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Compl_Clasf = lAux
        If mfRetCampo(4, "Descri��o Clasf", lAux, lIndice, 60, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dsc_Clasf = lAux
        If mfRetCampo(2, "Codigo da classe de ocup", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Classe_Ocup = lAux
        If mfRetCampo(2, "Codigo constru��o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Constr = lAux
        If mfRetCampo(4, "Usu�rio de digita��o", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_User_Digit = lAux
        If mfRetCampo(2, "Data da digita��o", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .DAT_DIGIT = lAux
        If mfRetCampo(2, "Hora da digita��o", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .HOR_DIGIT = lAux
        If mfRetCampo(4, "Usu�rio de libera��o", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_User_Liber = lAux
        If mfRetCampo(2, "Data da libera��o", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Liber = lAux
        If mfRetCampo(2, "Hora da libera��o", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Hor_Liber = lAux
        If mfRetCampo(2, "Situa��o do item", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Sit_item = lAux
        If mfRetCampo(2, "Percentual de b�nus", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Bonus = lAux
        If mfRetCampo(2, "Quantidade de empreg", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Qtd_Empreg = lAux
        If mfRetCampo(2, "N�mero INSP", lAux, lIndice, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Insp = lAux
        If mfRetCampo(4, "Identifica��o de seguran�a 1", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Segur_01_V = lAux
        If mfRetCampo(4, "Identifica��o de seguran�a 2", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Segur_02_A = lAux
        If mfRetCampo(4, "Identifica��o de seguran�a 3", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Segur_03_P = lAux
        If mfRetCampo(4, "Identifica��o de seguran�a 4", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Segur_04_I = lAux
        If mfRetCampo(4, "Identifica��o de seguran�a 5", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Segur_05_G = lAux
        If mfRetCampo(4, "Identifica��o de seguran�a 6", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Segur_06 = lAux
        If mfRetCampo(4, "Identifica��o de seguran�a 7", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Segur_07 = lAux
        If mfRetCampo(4, "Identifica��o de seguran�a 8", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Segur_08 = lAux
        If mfRetCampo(4, "Identifica��o de seguran�a 9", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Segur_09 = lAux
        If mfRetCampo(4, "Identifica��o de seguran�a 10", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Segur_10 = lAux
        If mfRetCampo(2, "Percentual total de seguran�a", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Tot_Segur = lAux
        If mfRetCampo(4, "Identifica��o de prote��o 1", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Protec_01_E = lAux
        If mfRetCampo(4, "Identifica��o de prote��o 2", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Protec_02_H = lAux
        If mfRetCampo(4, "Identifica��o de prote��o 3", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Protec_03_S = lAux
        If mfRetCampo(4, "Identifica��o de prote��o 4", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Protec_04_D = lAux
        If mfRetCampo(4, "Identifica��o de prote��o 5", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Protec_05 = lAux
        If mfRetCampo(4, "Identifica��o de prote��o 6", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Protec_06 = lAux
        If mfRetCampo(4, "Identifica��o de prote��o 7", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Protec_07 = lAux
        If mfRetCampo(4, "Identifica��o de prote��o 8", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Protec_08 = lAux
        If mfRetCampo(4, "Identifica��o de prote��o 9", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Protec_09 = lAux
        If mfRetCampo(4, "Identifica��o de prote��o 10", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Protec_10 = lAux
        If mfRetCampo(2, "Percentual total de prote��o", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Tot_Protec = lAux
        If mfRetCampo(2, "C�digo do produto", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Produto = lAux
        If mfRetCampo(2, "C�digo da regi�o", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Regiao = lAux
        If mfRetCampo(2, "C�digo de identifica��o da tarifa", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Iden_Tarif = lAux
        If mfRetCampo(2, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
        If mfRetCampo(2, "Tipo de emiss�o", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tip_Emissao = lAux
        If mfRetCampo(2, "Num_Apol_Anterior", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol_Anterior = lAux
        If mfRetCampo(2, "Item Anterior", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Item_Anterior = lAux
        If mfRetCampo(2, "Ind_Sinistro", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Ind_Sinistro = lAux
        
        '1000347 - Inclus�o de informa��es da cong�nere em renova��es
        If Len(mRegistro) > 432 Then
        
            If mfRetCampo(2, "c�digo da seguradora", lAux, lIndice, 5, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Ren_Con_Seguradora = lAux
        
            If mfRetCampo(4, "n�mero da ap�lice de renova��o", lAux, lIndice, 15, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Ren_Con_Apolice = lAux
        
            If mfRetCampo(4, "n�mero do item da ap�lice de renova��o", lAux, lIndice, 5, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Ren_Con_Item = lAux
            
            '1000305 - inclus�o de tipo de cobertura
            If mfRetCampo(2, "Tipo de Cobertura", lAux, lIndice, 1, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Tip_Cobert = lAux

        End If
        
        
        If Len(mRegistro) > 461 Then
            If mfRetCampo(2, "Inspe��o Obrigat�ria", lAux, lIndice, 1, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Insp_Obr = lAux
            
            If mfRetCampo(2, "Inspe��o Agend", lAux, lIndice, 1, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Insp_Agend = lAux
        
            If mfRetCampo(2, "Digita��o de Bens", lAux, lIndice, 1, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Dig_Bens = lAux
        
            If mfRetCampo(2, "Cobertura de Bens", lAux, lIndice, 1, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Cob_Bens = lAux
        End If
        
        If (.Cod_Ramo = 300 Or .Cod_Ramo = 620 Or .Cod_Ramo = 710) And Len(mRegistro) = 470 Then
            If mfRetCampo(2, "Indicador de Sinistro Inspe��o", lAux, lIndice, 4, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Ind_Sin_Inspec = lAux
            
            If mfRetCampo(2, "Codigo do Dizer do Item", lAux, lIndice, 4, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Cod_DizerItem = lAux
        End If
        
        ' Daniel Ribeiro - Imobili�rio - 29/10/2015
        If Len(mRegistro) > 596 Then
            If mfRetCampo(2, "Indicador de Sinistro Inspe��o", lAux, lIndice, 4, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Ind_Sin_Inspec = lAux
            
            If mfRetCampo(2, "Codigo do Dizer do Item", lAux, lIndice, 4, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Cod_DizerItem = lAux
        
            If mfRetCampo(2, "Indicador de verbas separadas", lAux, lIndice, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Ind_VerbaSep = lAux
        
            If mfRetCampo(2, "Valor em risco lucros cessantes", lAux, lIndice, 15, 5, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Val_Vr_Lc = lAux
        
            If mfRetCampo(2, "Valor LMG lucros cessantes", lAux, lIndice, 15, 5, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Val_LMG_Lc = lAux
        
            If mfRetCampo(2, "Percentual IS VR LC", lAux, lIndice, 15, 5, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Per_IS_Vr_Lc = lAux
        
            If mfRetCampo(2, "Coeficiente IS VR LC", lAux, lIndice, 15, 5, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Coef_IS_Vr_Lc = lAux
        
            If mfRetCampo(2, "Tipo de plano de vida", lAux, lIndice, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Tip_Plano_Vida = lAux
        
            If mfRetCampo(2, "C�digo do plano de cobertura", lAux, lIndice, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Cod_Plano_Cobertura = lAux
        
            If mfRetCampo(2, "Quantidade de vidas", lAux, lIndice, 6, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Qtd_Vidas = lAux
        
            If mfRetCampo(2, "Valor LMG Vida", lAux, lIndice, 15, 5, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Val_LMG_Vida = lAux
        
            If mfRetCampo(2, "Indicador de auxilio funeral", lAux, lIndice, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Ind_Aux_Funeral = lAux
        
            If mfRetCampo(2, "Valor LMG Funeral", lAux, lIndice, 15, 5, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Val_LMG_Funeral = lAux
        
            If mfRetCampo(2, "Indicador de inclus�o de conjuge", lAux, lIndice, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Ind_Inc_Conjuge = lAux
        
            If mfRetCampo(2, "Valor Aux�lio Cesta", lAux, lIndice, 15, 5, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Val_Aux_Cesta = lAux
        
            If mfRetCampo(2, "Indicador de envio de certificado", lAux, lIndice, 1, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Flg_EnviarCertificado = lAux
        
            If mfRetCampo(4, "Observa��es do item", lAux, lIndice, 254, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Txt_Observacoes = lAux
        
        
        End If
        
    End With

    Set lobjItem = New clsA46V708A
    Set lobjItem.M46V999 = M46V999

    'Temporario
    lobjstcItem.nosso_numero = mNossoTempor

    If pGravar = True Then
    
        If lobjstcItem.Cod_Ramo = e_CodRamo_Resp_Civil And lobjstcItem.Tab_Clasf > 0 Then
            'Local de risco adicional de um item j� gravado.
            If mfRecLocAdic(lobjstcItem, pMensagem, pPropTabela, pNomeTabImport) = False Then
                pMensagem = lobjItem.UltimoErro
                GoTo SAIDA
            End If
            
        Else
            Set lobjItem.ConexaoBD = bdBasDados
            Set lobjstcItem.ObjPai = pStcPedido
            If Not lobjItem.mfIncluir(lobjstcItem, pPropTabela, pNomeTabImport, Empty) Then
                pMensagem = lobjItem.UltimoErro
                GoTo SAIDA
            End If
            If Not IsMissing(pStcPedido) Then
                pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
            End If
        End If
        
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecItens = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecItens Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjItem = Nothing
    Set lobjstcItem = Nothing
    Return
    
End Function

Private Function mfRecLocAdic(ByRef pobjstcItem As stcA46V708B, _
                              ByRef pMensagem As String, _
                              ByVal pPropTabela As String, _
                              ByVal pNomeTabImport As String) As Boolean
Dim lClsLocAdic As clsA46V737A
Dim lStcLocAdic As stcA46V737B

Err.Clear
mfRecLocAdic = False
On Error GoTo TrataErro

    Set lClsLocAdic = New clsA46V737A
    Set lClsLocAdic.M46V999 = M46V999
    Set lStcLocAdic = New stcA46V737B

    With lStcLocAdic
        Set .ObjPai = pobjstcItem
        .nosso_numero = pobjstcItem.nosso_numero
        .Num_Item = pobjstcItem.Num_Item
        .NUM_SEQ = pobjstcItem.Tab_Clasf
        .Cod_Tip_Logr_Risco = pobjstcItem.Cod_Tip_Logr_Risco
        .Nom_Logr_Risco = pobjstcItem.Nom_Logr_Risco
        .Dsc_Compl_Risco = pobjstcItem.Dsc_Compl_Risco
        .Nom_Bairro_Risco = pobjstcItem.Nom_Bairro_Risco
        .Nom_Cid_Risco = pobjstcItem.Nom_Cid_Risco
        .Sig_UF_Risco = pobjstcItem.Sig_UF_Risco
        .Num_Cep_Risco = pobjstcItem.Num_Cep_Risco
    End With
    
    If Not lClsLocAdic.gfIncluir(lStcLocAdic, pPropTabela, pNomeTabImport) Then
        pMensagem = lClsLocAdic.UltimoErro
        GoTo SAIDA
    End If
    
'================================
'******Fun��o realizada com sucesso
    mfRecLocAdic = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecLocAdic Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lClsLocAdic = Nothing
    Set lStcLocAdic = Nothing
    Return
    
End Function

Private Function mfRecClausulas(ByVal pNossoNumero As String, _
                                ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                                Optional pGravar As Boolean = True, _
                                Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de cl�usulas para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
    '1. Grava registro de dados de cl�usula.
Dim lobjClausula         As clsA46V703A      'objeto da classe clsA46V703A (TAB_PED_CLAU)
Dim lobjStcClausula      As stcA46V703B      'objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU)
Dim lobjstcItem          As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
Dim lIndice                 As Integer
Dim lAux                    As Variant
    
Err.Clear
mfRecClausulas = False
On Error GoTo TrataErro

    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "010" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_CLAU"
        GoTo SAIDA
    End If
            
    Set lobjStcClausula = New stcA46V703B
    Set lobjStcClausula.M46V999 = M46V999
    
    lIndice = lIndice + 3
        
    With lobjStcClausula
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Clausulas n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "N�mero do Item", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Item = lAux
        If mfRetCampo(2, "C�digo da Cl�usula", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Clau = lAux
        If mfRetCampo(2, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
        If mfRetCampo(4, "Descri��o", lAux, lIndice, 255, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Descricao = lAux
    End With

    Set lobjClausula = New clsA46V703A
    Set lobjClausula.M46V999 = M46V999

    'Tempor�rio
    lobjStcClausula.nosso_numero = mNossoTempor

    If pGravar = True Then
        Set lobjClausula.ConexaoBD = bdBasDados
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjStcClausula.nosso_numero & _
                                                            Format(lobjStcClausula.Num_Item, "00000")))
            Set lobjStcClausula.ObjPai = lobjstcItem
            lobjstcItem.ColecaoPedClau.Add Item:=lobjStcClausula, Key:=lobjStcClausula.Class_ID
            
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        End If
        If Not lobjClausula.mfIncluir(lobjStcClausula, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjClausula.UltimoErro
            GoTo SAIDA
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjStcClausula.nosso_numero & _
                                                            Format(lobjStcClausula.Num_Item, "00000")))
            Set lobjStcClausula.ObjPai = lobjstcItem
            lobjstcItem.ColecaoPedClau.Add Item:=lobjStcClausula, Key:=lobjStcClausula.Class_ID
            
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecClausulas = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecClausulas Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjClausula = Nothing
    Set lobjStcClausula = Nothing
    Return

End Function

Private Function mfRecPessoas(ByVal pNossoNumero As String, _
                                ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                                Optional pGravar As Boolean = True, _
                                Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de pessoas para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
    '1. Grava registro de dados de pessoas.
Dim lobjPessoa           As clsA46V735A      'objeto da classe clsA46V735A (TAB_PED_LOC_PESSOAS)
Dim lobjstcPessoa      As stcA46V735B      'objeto da classe de estrutura stcA46V735B (TAB_PED_LOC_PESSOAS)
Dim lobjstcItem          As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
Dim lIndice                 As Integer
Dim lAux                    As Variant
    
Err.Clear
mfRecPessoas = False
On Error GoTo TrataErro

    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "026" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_LOC_PESSOAS"
        GoTo SAIDA
    End If
            
    Set lobjstcPessoa = New stcA46V735B
    Set lobjstcPessoa.M46V999 = M46V999
    
    lIndice = lIndice + 3
        
    With lobjstcPessoa
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Pessoas n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        '.Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        '.Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        '.Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "N�mero do Item", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Item = lAux
        If mfRetCampo(2, "Tipo de local pessoa", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tip_Loc_Pessoa = lAux
        If mfRetCampo(2, "Tipo de pessoa", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tip_Pessoa = lAux
        If mfRetCampo(6, "N�mero do CPF ou CNPJ", lAux, lIndice, 14, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_CPF_CNPJ = lAux
        If mfRetCampo(2, "Tipo de cobertura", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tip_Cobert = lAux
        If mfRetCampo(2, "Percentual de participa��o", lAux, lIndice, 5, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Participacao = lAux
        If mfRetCampo(4, "Nome da pessoa", lAux, lIndice, 60, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Pessoa = lAux
    End With

    Set lobjPessoa = New clsA46V735A
    Set lobjPessoa.M46V999 = M46V999

    'Tempor�rio
    lobjstcPessoa.nosso_numero = mNossoTempor

    If pGravar = True Then
        Set lobjPessoa.ConexaoBD = bdBasDados
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcPessoa.nosso_numero & _
                                                            Format(lobjstcPessoa.Num_Item, "00000")))
            Set lobjstcPessoa.ObjPai = lobjstcItem
            lobjstcItem.ColecaoBeneficiario.Add Item:=lobjstcPessoa ', Key:=lobjstcPessoa.Class_ID
            
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        End If
        If Not lobjPessoa.mfIncluir(lobjstcPessoa, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjPessoa.UltimoErro
            GoTo SAIDA
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcPessoa.nosso_numero & _
                                                            Format(lobjstcPessoa.Num_Item, "00000")))
            Set lobjstcPessoa.ObjPai = lobjstcItem
            lobjstcItem.ColecaoBeneficiario.Add Item:=lobjstcPessoa, Key:=lobjstcPessoa.Class_ID
            
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecPessoas = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecPessoas Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjPessoa = Nothing
    Set lobjstcPessoa = Nothing
    Return

End Function

Private Function mfRecPgtoCartaoCredito(ByVal pNossoNumero As String, _
                                        ByRef pMensagem As String, ByVal pPropTabela As String, _
                                        ByVal pNomeTabImport As String, _
                                        Optional pGravar As Boolean = True, _
                                        Optional pStcPedido As stcA46V702B) As Boolean
    
'Fun��o: grava registro de cart�o de cr�dito para transmiss�o.

'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
'Par�metros de sa�da.....pMensagem............Mensagem.

'1. Grava registro de dados de Cartao de Cr�dito.
Dim lobjCartao           As clsA46V756A      'objeto da classe clsA46V735A (TAB_DADOS_CARTAO)
Dim lobjstcCartao        As stcA46V756B      'objeto da classe de estrutura stcA46V756B (TAB_DADOS_CARTAO)
Dim lIndice                 As Integer
Dim lAux                    As Variant
    
Err.Clear
mfRecPgtoCartaoCredito = False
On Error GoTo TrataErro

    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "027" Then
        pMensagem = "Registro n�o faz parte da cole��o. [SYAS_EMIS].[dbo].[TAB_DADOS_CARTAO]"
        GoTo SAIDA
    End If
            
    Set lobjstcCartao = New stcA46V756B
    Set lobjstcCartao.M46V999 = M46V999
    
    lIndice = lIndice + 3
        
    With lobjstcCartao
        If mfRetCampo(4, "Nosso N�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        
        If mfRetCampo(6, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI_EMIT = lAux
        
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .COD_RAMO_EMIT = lAux
        
        If mfRetCampo(4, "TransactionID", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .TransactionId = lAux
        
        If mfRetCampo(4, "Authcode", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .AUTHCODE = lAux
        
        If mfRetCampo(4, "DtMov", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        lAux = Mid(lAux, 1, 4) & "-" & Mid(lAux, 5, 2) & "-" & Mid(lAux, 7, 2)
        .DTMOV = lAux
        
        If mfRetCampo(4, "TransactionDate", lAux, lIndice, 24, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        lAux = Trim(lAux)
        lAux = Mid(lAux, 1, 4) & "-" & Mid(lAux, 5, 2) & "-" & Mid(lAux, 7, 2) & " " & Mid(lAux, 9, 2) & ":" & Mid(lAux, 11, 2) & ":" & Mid(lAux, 13, 2) & "." & Mid(lAux, 15, 3)
        .TRANSACTIONDATE = lAux
        
        If mfRetCampo(2, "TransactionStatus", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .TRANSACTIONSTATUS = lAux
        
        If mfRetCampo(2, "TransacationType", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .TRANSACTIONTYPE = lAux
        
        If mfRetCampo(2, "OrderNumber", lAux, lIndice, 18, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .ORDERNUMBER = lAux
        
        If mfRetCampo(2, "OrderTotal", lAux, lIndice, 18, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .ORDERTOTAL = lAux
        
        If mfRetCampo(4, "Error", lAux, lIndice, 100, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .ERROR = lAux
        
        If mfRetCampo(4, "XML_send", lAux, lIndice, 400, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        lAux = Trim(lAux)
        .XML_SEND = lAux
        
        If mfRetCampo(4, "XML_receive", lAux, lIndice, 2300, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        lAux = Trim(lAux)
        If InStr(1, lAux, "<?xml version=""1.0"" encoding=""utf-16""?>") = 1 Then
            lAux = Replace(lAux, "<?xml version=""1.0"" encoding=""utf-16""?>", "")
        ElseIf InStr(1, lAux, "<?xml version=""1.0"" encoding=""utf-8""?>") = 1 Then
            lAux = Replace(lAux, "<?xml version=""1.0"" encoding=""utf-8""?>", "")
        End If
        .XML_RECEIVE = lAux
        
        If mfRetCampo(2, "Num_tempo_total", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_TEMPO_TOTAL = lAux
        
        If mfRetCampo(4, "Timesatampbw", lAux, lIndice, 24, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        lAux = Trim(lAux)
        lAux = Mid(lAux, 1, 4) & "-" & Mid(lAux, 5, 2) & "-" & Mid(lAux, 7, 2) & " " & Mid(lAux, 9, 2) & ":" & Mid(lAux, 11, 2) & ":" & Mid(lAux, 13, 2) & "." & Mid(lAux, 15, 3)
        .TIMESATAMPBW = lAux
        
    End With

    Set lobjCartao = New clsA46V756A
    Set lobjCartao.M46V999 = M46V999

    If pGravar = True Then
    
        'Abre banco de Dados SYAS_EMIS
        If M46V999.gfAbrBasDados("SYAS_EMIS", M46V999.gTipSYASEMIS, M46V999.gEndSYASEMIS, bdBasDadosSYAS_EMIS, pMensagem) = False Then
            GoTo SAIDA
        End If
        
        Set lobjCartao.ConexaoBD = bdBasDadosSYAS_EMIS
        If Not lobjCartao.mfIncluir(lobjstcCartao, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjCartao.UltimoErro
            GoTo SAIDA
        End If
        
    End If
    
'================================
'******Fun��o realizada com sucesso
    mfRecPgtoCartaoCredito = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecPgtoCartaoCredito Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjCartao = Nothing
    Set lobjstcCartao = Nothing
    Return

End Function

Private Function mfRecInspecao(ByVal pNossoNumero As String, _
                               ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                               Optional pGravar As Boolean = True, _
                               Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de inspe��o

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
Dim lobjInsp             As clsA46V717A      'Objeto da classe clsA46V717A (TAB_PED_INSP)
Dim lObjstcInsp          As stcA46V717B      'Objeto da classe de estrutura stcA46V717B (TAB_PED_INSP)
Dim lobjstcItem          As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
Dim lIndice                 As Integer
Dim lAux                    As Variant

Err.Clear
mfRecInspecao = False
On Error GoTo TrataErro

    '1. Grava registro de dados de inspe��o.
    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "011" Then
        pMensagem = "Registro n�o faz parte da cole��o. Tab_Ped_Fat_Risco"
        GoTo SAIDA
    End If
            
    Set lObjstcInsp = New stcA46V717B
    
    lIndice = lIndice + 3
        
    With lObjstcInsp
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Complemento de Inspe��o n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero do Item", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Item = lAux
        If mfRetCampo(2, "N�mero da inspe��o", lAux, lIndice, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Insp = lAux
        If mfRetCampo(2, "Data do cadastro", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Cadastro = lAux
        If mfRetCampo(4, "Nome do contato", lAux, lIndice, 40, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Pess_Cont = lAux
        If mfRetCampo(4, "Telefone de contato", lAux, lIndice, 14, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Tel = lAux
        If mfRetCampo(2, "N�mero do logradouro - risco", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Logr_Risco = lAux
    End With
    
    Set lobjInsp = New clsA46V717A
    Set lobjInsp.M46V999 = M46V999

    lObjstcInsp.nosso_numero = mNossoTempor
    
    If pGravar = True Then
        Set lobjInsp.ConexaoBD = bdBasDados
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lObjstcInsp.nosso_numero & _
                                                            Format(lObjstcInsp.Num_Item, "00000")))
            Set lObjstcInsp.ObjPai = lobjstcItem
            Set lobjstcItem.objstcPedInsp = lObjstcInsp
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        End If
        If Not lobjInsp.mfIncluir(lObjstcInsp, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjInsp.UltimoErro
            GoTo SAIDA
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lObjstcInsp.nosso_numero & _
                                                            Format(lObjstcInsp.Num_Item, "00000")))
            Set lObjstcInsp.ObjPai = lobjstcItem
            Set lobjstcItem.objstcPedInsp = lObjstcInsp
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecInspecao = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecInspecao Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjInsp = Nothing
    Set lObjstcInsp = Nothing
    Return

End Function

Private Function mfRecObservacao(ByVal pNossoNumero As String, _
                                 ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                                 Optional pGravar As Boolean = True, _
                                 Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de observa��o

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    '                        pObjstcPedido........objeto de pedido.
    '                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
Dim lobjstcPedObs        As stcA46V714B 'Objeto da classe de estrutura stcA46V714B (TAB_PED_OBS)
Dim lobjPedObservacao    As clsA46V714A 'Objeto da classe clsA46V714A (TAB_PED_OBS)
Dim lobjstcItem          As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
Dim lIndice                 As Integer
Dim lAux                    As Variant
    
Err.Clear
mfRecObservacao = False
On Error GoTo TrataErro

    '1. Grava registro de dados de observa��o.
    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "012" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_OBS"
        GoTo SAIDA
    End If
            
    Set lobjstcPedObs = New stcA46V714B
    
    lIndice = lIndice + 3
    
    '1. Grava registro de dados de observa��o.
        
    With lobjstcPedObs
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Observacao n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "N�mero do Item", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Item = lAux
        If mfRetCampo(4, "Observa��o", lAux, lIndice, 500, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Observacao = lAux
    End With
    Set lobjPedObservacao = New clsA46V714A
    Set lobjPedObservacao.M46V999 = M46V999
    
    'Temporario
    lobjstcPedObs.nosso_numero = mNossoTempor
    
    If pGravar = True Then
        Set lobjPedObservacao.ConexaoBD = bdBasDados
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcPedObs.nosso_numero & _
                                                            Format(lobjstcPedObs.Num_Item, "00000")))
                                                                
            Set lobjstcPedObs.ObjPai = lobjstcItem
            Set lobjstcItem.objstcPedObs = lobjstcPedObs
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        End If
        If Not lobjPedObservacao.mfIncluir(lobjstcPedObs, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjPedObservacao.UltimoErro
            GoTo SAIDA
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcPedObs.nosso_numero & _
                                                            Format(lobjstcPedObs.Num_Item, "00000")))
                                                                
            Set lobjstcPedObs.ObjPai = lobjstcItem
            Set lobjstcItem.objstcPedObs = lobjstcPedObs
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecObservacao = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecObservacao Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjPedObservacao = Nothing
    Set lobjstcPedObs = Nothing
    Return
End Function

Private Function mfRecSeguro(ByVal pNossoNumero As String, _
                             ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                             Optional pGravar As Boolean = True, _
                             Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de seguros para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    '                        pObjstcPedido........objeto de pedido.
    '                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
Dim lobjPedLocSeg       As clsA46V716A          'Objeto da classe clsA46V716A (TAB_PED_LOC_SEGURO).
Dim lobjstcPedLocSeg    As stcA46V716B          'Objeto da classe de estrutura stcA46V716B (TAB_PED_LOC_SEGURO).
Dim lobjstcItem          As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
Dim lIndice                 As Integer
Dim lAux                    As Variant

Err.Clear
mfRecSeguro = False
On Error GoTo TrataErro
  
    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "013" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_LOC_SEGURO"
        GoTo SAIDA
    End If
            
    Set lobjstcPedLocSeg = New stcA46V716B
    
    lIndice = lIndice + 3
        
    With lobjstcPedLocSeg
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Outros Seguros n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "N�mero do Item", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Item = lAux
        If mfRetCampo(4, "Descri��o do outro seguro", lAux, lIndice, 35, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dsc_Outro_Seguro = lAux
        If mfRetCampo(2, "C�digo da seguradora", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Seguradora = lAux
        If mfRetCampo(2, "N�mero da ordem", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ordem = lAux
        If mfRetCampo(4, "N�mero da ap�lice anterior", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol_Ant = lAux
        If mfRetCampo(2, "N�mero do item anterior", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Item_Ant = lAux
        If mfRetCampo(2, "Indicador de sinistro", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Ind_Sinistro = lAux
    End With

    Set lobjPedLocSeg = New clsA46V716A
    Set lobjPedLocSeg.M46V999 = M46V999

    lobjstcPedLocSeg.nosso_numero = mNossoTempor
    
    If pGravar = True Then
        Set lobjPedLocSeg.ConexaoBD = bdBasDados
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcPedLocSeg.nosso_numero & _
                                                            Format(lobjstcPedLocSeg.Num_Item, "00000")))
                                                                
            Set lobjstcPedLocSeg.ObjPai = lobjstcItem
            lobjstcItem.ColecaoPedSeguro.Add Item:=lobjstcPedLocSeg, Key:=lobjstcPedLocSeg.Class_ID
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        End If
        If Not lobjPedLocSeg.mfIncluir(lobjstcPedLocSeg, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjPedLocSeg.UltimoErro
            GoTo SAIDA
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcPedLocSeg.nosso_numero & _
                                                            Format(lobjstcPedLocSeg.Num_Item, "00000")))
            Set lobjstcPedLocSeg.ObjPai = lobjstcItem
            lobjstcItem.ColecaoPedSeguro.Add Item:=lobjstcPedLocSeg, Key:=lobjstcPedLocSeg.Class_ID
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecSeguro = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecSeguro Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjPedLocSeg = Nothing
    Set lobjstcPedLocSeg = Nothing
    Return

End Function

Private Function mfRecBens(ByVal pNossoNumero As String, _
                           ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                           Optional pGravar As Boolean = True, _
                           Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de bens para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    '                        pObjstcPedido........objeto de pedido.
    '                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
    Dim lobjPedBem       As clsA46V715A         'Objeto da classe clsA46V715A (TAB_PED_LOC_BEM).
    Dim lobjstcPedBem    As stcA46V715B         'Objeto da classe de estrutura stcA46V715B (TAB_PED_LOC_BEM).
    Dim lobjstcItem      As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lIndice                 As Integer
    Dim lAux                    As Variant
Err.Clear
mfRecBens = False
On Error GoTo TrataErro

    '1. Grava registro de dados de inspe��o.
    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "014" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_LOC_BEM"
        GoTo SAIDA
    End If
            
    Set lobjstcPedBem = New stcA46V715B
    Set lobjstcPedBem.M46V999 = M46V999

    lIndice = lIndice + 3
        
    With lobjstcPedBem
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Bens Segurados n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "N�mero do Item", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Item = lAux
        If mfRetCampo(4, "Descri��o do bem segurado", lAux, lIndice, 60, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dsc_Bem = lAux
        If mfRetCampo(2, "Valor do bem declarado", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Bem = lAux
        If mfRetCampo(2, "C�digo do bem ", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Bem = lAux
        If mfRetCampo(2, "C�digo da Sequencia ", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Seq = lAux
    End With
    
    Set lobjPedBem = New clsA46V715A
    Set lobjPedBem.M46V999 = M46V999

    'Temporario
    lobjstcPedBem.nosso_numero = mNossoTempor

    If pGravar = True Then
        Set lobjPedBem.ConexaoBD = bdBasDados
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcPedBem.nosso_numero & _
                                                            Format(lobjstcPedBem.Num_Item, "00000")))
                                                                
            Set lobjstcPedBem.ObjPai = lobjstcItem
            lobjstcItem.ColecaoPedLocBem.Add Item:=lobjstcPedBem, Key:=lobjstcPedBem.Class_ID
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        End If
        If Not lobjPedBem.mfIncluir(lobjstcPedBem, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjPedBem.UltimoErro
            GoTo SAIDA
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcPedBem.nosso_numero & _
                                                            Format(lobjstcPedBem.Num_Item, "00000")))
            Set lobjstcPedBem.ObjPai = lobjstcItem
            lobjstcItem.ColecaoPedLocBem.Add Item:=lobjstcPedBem, Key:=lobjstcPedBem.Class_ID
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecBens = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecBens Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjPedBem = Nothing
    Set lobjstcPedBem = Nothing
    Return

End Function

Private Function mfRecFatoresProtecao(ByVal pNossoNumero As String, _
                                      ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                                      Optional pGravar As Boolean = True, _
                                      Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de bens para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    '                        pObjstcPedido........objeto de pedido.
    '                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.

Dim lobjFator            As clsA46V720A      'Objeto da classe clsA46V720A (Tab_Ped_Fat_Risco)
Dim lobjstcFator         As stcA46V720B      'Objeto da classe de estrutura stcA46V720B (Tab_Ped_Fat_Risco)
Dim lobjstcItem      As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
Dim lIndice                 As Integer
Dim lAux                    As Variant
    
Err.Clear
mfRecFatoresProtecao = False
On Error GoTo TrataErro
    
    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "015" Then
        pMensagem = "Registro n�o faz parte da cole��o. Tab_Ped_Fat_Risco"
        GoTo SAIDA
    End If
            
    Set lobjstcFator = New stcA46V720B
    
    lIndice = lIndice + 3
        
    With lobjstcFator
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Fatores de Prote��o n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "N�mero do Item", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Item = lAux
        If mfRetCampo(2, "C�digo do Fator", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Fator = lAux
        If mfRetCampo(4, "Descri��o do Fator", lAux, lIndice, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dsc_Compl_Fator = lAux
        If mfRetCampo(2, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
    End With
    
    Set lobjFator = New clsA46V720A
    Set lobjFator.M46V999 = M46V999

    lobjstcFator.nosso_numero = mNossoTempor
    
    If pGravar = True Then
        Set lobjFator.ConexaoBD = bdBasDados
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcFator.nosso_numero & _
                                                            Format(lobjstcFator.Num_Item, "00000")))
                                                                
            Set lobjstcFator.ObjPai = lobjstcItem
            lobjstcItem.ColecaoPedFator.Add Item:=lobjstcFator, Key:=lobjstcFator.Class_ID
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        End If
        If Not lobjFator.mfIncluir(lobjstcFator, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjFator.UltimoErro
            GoTo SAIDA
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcFator.nosso_numero & _
                                                            Format(lobjstcFator.Num_Item, "00000")))
            Set lobjstcFator.ObjPai = lobjstcItem
            lobjstcItem.ColecaoPedFator.Add Item:=lobjstcFator, Key:=lobjstcFator.Class_ID
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecFatoresProtecao = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecFatoresProtecao Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjFator = Nothing
    Set lobjstcFator = Nothing
    Return

End Function


Private Function mfRecCoberturas(ByVal pNossoNumero As String, _
                                 ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                                 Optional pGravar As Boolean = True, _
                                 Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de coberturas para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
    Dim lobjCobertura        As clsA46V704A      'Objeto da classe clsA46V704A (TAB_PED_COBERT)
    Dim lobjstcCobertura     As stcA46V704B      'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
    Dim lobjstcItem          As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lIndice                 As Integer
    Dim lAux                    As Variant
    
Err.Clear
mfRecCoberturas = False
On Error GoTo TrataErro

    '1. Grava registro de coberturas.

    lIndice = 1
    If Mid(mRegistro, lIndice, 3) <> "016" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_COBERT"
        GoTo SAIDA
    End If
            
    Set lobjstcCobertura = New stcA46V704B
    Set lobjstcCobertura.M46V999 = M46V999
    
    lIndice = lIndice + 3
        
    With lobjstcCobertura
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Coberturas n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "N�mero do Item", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Item = lAux
        If mfRetCampo(2, "Tab Cobertura", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tab_Cobert = lAux
        If mfRetCampo(2, "C�digo da Cobertura", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Cobert = lAux
        If mfRetCampo(2, "Valor da IS", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_IS = lAux
        If mfRetCampo(2, "Percentual da Taxa", lAux, lIndice, 7, 5, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Taxa = lAux
        If mfRetCampo(2, "Percentual de desconto comercial", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Desc_Coml = lAux
        If mfRetCampo(2, "Percentual de desconto Exp", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Desc_Exp = lAux
        If mfRetCampo(2, "Valor do pr�mio", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .VAL_PR = lAux
        If mfRetCampo(2, "Valor do pr�mio anual", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Pr_Anu = lAux
        If mfRetCampo(2, "Valor do pr�mio Tarif", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Pr_Tarif = lAux
        If mfRetCampo(2, "Valor do pr�mio Tarif anual", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Pr_Tarif_Anu = lAux
        If mfRetCampo(2, "Valor do pr�mio Resg", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Pr_Resg = lAux
        If mfRetCampo(2, "Valor do pr�mio Resg anual", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Pr_Resg_Anu = lAux
        If mfRetCampo(2, "Percentual da franquia", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Franq = lAux
        If mfRetCampo(2, "Valor m�nimo da franquia", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Franq_Min = lAux
        If mfRetCampo(2, "Valor m�ximo da franquia", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Franq_Max = lAux
        If mfRetCampo(2, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
        
        ' Daniel Ribeiro - Imobili�rio - Inclus�o de Periodo Indenit�rio
        If Len(mRegistro) > 227 Then
            If mfRetCampo(2, "N�mero de meses Per�odo Indenit�rio", lAux, lIndice, 2, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Num_Meses_PI = lAux
            If .Tab_Cobert = 20 And (.Cod_Ramo = e_CodRamo_Empresarial Or .Cod_Ramo = e_CodRamo_Residencial) Then  ' � imobili�rio
                .Val_Pr_Tarif = lAux ' Gambi para mandar o numero de meses PI na TAB_PED_COBERT, j� que a TAB_PED_COBERT_COMPL n�o � salva pelo Imobili�rio
            End If
        End If
        
        
        
    End With
    
    Set lobjCobertura = New clsA46V704A
    Set lobjCobertura.M46V999 = M46V999

    lobjstcCobertura.nosso_numero = mNossoTempor

    If pGravar = True Then
        Set lobjCobertura.ConexaoBD = bdBasDados
        '*** Inclus�o da condi��o abaixo para casos com o Ramo 510, onde o Item � 99999 - 11/09/2013 Projeto RD - SIES - taflorencio ***
        If Format(lobjstcCobertura.Num_Item, "00000") <> 99999 Then
            If Not IsMissing(pStcPedido) Then
                Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcCobertura.nosso_numero & _
                                                                Format(lobjstcCobertura.Num_Item, "00000")))
                Set lobjstcCobertura.ObjPai = lobjstcItem
                lobjstcItem.ColecaoPedCobert.Add Item:=lobjstcCobertura, Key:=lobjstcCobertura.Class_ID
                
                pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
                pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
            End If
        End If
        If Not lobjCobertura.mfIncluir(lobjstcCobertura, pPropTabela, pNomeTabImport, True, False, Empty) Then
            pMensagem = lobjCobertura.UltimoErro
            GoTo SAIDA
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcCobertura.nosso_numero & _
                                                            Format(lobjstcCobertura.Num_Item, "00000")))
            Set lobjstcCobertura.ObjPai = lobjstcItem
            lobjstcItem.ColecaoPedCobert.Add Item:=lobjstcCobertura, Key:=lobjstcCobertura.Class_ID
            
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecCoberturas = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecCoberturas Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjCobertura = Nothing
    Set lobjstcCobertura = Nothing
    Return
End Function

Private Function mfRecCobertCompl(ByVal pNossoNumero As String, _
                                  ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                                  Optional pGravar As Boolean = True, _
                                  Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de coberturas para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    '                        pObjstcPedido........objeto de pedido.
    '                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
Dim lobjCobertura        As clsA46V704A      'Objeto da classe clsA46V704A (TAB_PED_COBERT)
Dim lobjstcCobertura     As stcA46V704B      'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
Dim lobjstcCobertura1    As stcA46V704B      'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
Dim lobjstcItem          As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
Dim lIndice                 As Integer
Dim lAux                    As Variant
    
Err.Clear
mfRecCobertCompl = False
On Error GoTo TrataErro
          
    '1. Grava registro de dados complementares das coberturas.
    
    lIndice = 1
    If Mid(mRegistro, lIndice, 3) <> "017" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_COBERT_COMPL"
        GoTo SAIDA
    End If
            
    Set lobjstcCobertura = New stcA46V704B
    Set lobjstcCobertura.M46V999 = M46V999
    
    lIndice = lIndice + 3
        
    With lobjstcCobertura
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Complemento de Coberturas n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "N�mero do Item", lAux, lIndice, 5, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Item = lAux
        If mfRetCampo(2, "Tab Cobertura", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tab_Cobert = lAux
        If mfRetCampo(2, "C�digo da Cobertura", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Cobert = lAux
        If mfRetCampo(2, "Num_Meses_Pi", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Meses_PI = lAux
        If mfRetCampo(4, "Formula", lAux, lIndice, 100, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Formula = lAux
        If mfRetCampo(2, "Cod_Classe", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .COD_CLASSE = lAux
        If mfRetCampo(2, "Val_Fator_Franq", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Fator_Franq = lAux
        If mfRetCampo(4, "Alterar_Cobert", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Alterar_Cobert = lAux
        If mfRetCampo(2, "Val_Risco_Cob", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Risco_Cob = lAux
        If mfRetCampo(2, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
        If Len(mRegistro) > 236 Then
            If mfRetCampo(2, "Val_Fator_Franq", lAux, lIndice, 15, 5, pMensagem) = False Then
                GoTo SAIDA
            End If
            .Val_Fator_Franq = lAux
        End If
        If Len(mRegistro) > 251 Then
            If mfRetCampo(4, "AjusteEndosso", lAux, lIndice, 5, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .AjusteEndosso = lAux
            If mfRetCampo(4, "Alteroufranquia", lAux, lIndice, 3, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .AlterouFranquia = lAux
            If mfRetCampo(4, "TextoFranquia", lAux, lIndice, 200, 0, pMensagem) = False Then
                GoTo SAIDA
            End If
            .TextoFranquia = lAux
        End If
    End With
    
    Set lobjCobertura = New clsA46V704A
    Set lobjCobertura.M46V999 = M46V999

    lobjstcCobertura.nosso_numero = mNossoTempor

    If pGravar = True Then
        Set lobjCobertura.ConexaoBD = bdBasDados
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcCobertura.nosso_numero & _
                                                            Format(lobjstcCobertura.Num_Item, "00000")))
            
            Set lobjstcCobertura1 = lobjstcItem.ColecaoPedCobert.Item(CStr(lobjstcCobertura.nosso_numero & _
                                                                      Format(lobjstcCobertura.Num_Item, "00000") & _
                                                                      Format(lobjstcCobertura.Cod_Cobert, "0000") & _
                                                                      Format(lobjstcCobertura.Tab_Cobert, "00000")))
            
            With lobjstcCobertura1
                .Num_Meses_PI = lobjstcCobertura.Num_Meses_PI
                .Formula = lobjstcCobertura.Formula
                .COD_CLASSE = lobjstcCobertura.COD_CLASSE
                .Val_Fator_Franq = lobjstcCobertura.Val_Fator_Franq
                .Alterar_Cobert = lobjstcCobertura.Alterar_Cobert
                .Val_Risco_Cob = lobjstcCobertura.Val_Risco_Cob
                .AjusteEndosso = lobjstcCobertura.Alterar_Cobert
                .AlterouFranquia = lobjstcCobertura.Alterar_Cobert
                .TextoFranquia = lobjstcCobertura.Alterar_Cobert
            End With
            Set lobjstcCobertura.ObjPai = lobjstcItem
            Set lobjstcCobertura1.ObjPai = lobjstcItem
            lobjstcItem.ColecaoPedCobert.Remove lobjstcCobertura.Class_ID
            lobjstcItem.ColecaoPedCobert.Add Item:=lobjstcCobertura1, Key:=lobjstcCobertura1.Class_ID
            
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        End If
        If Not lobjCobertura.mfIncluir(lobjstcCobertura, pPropTabela, pNomeTabImport, False, True, Empty) Then
            pMensagem = lobjCobertura.UltimoErro
            GoTo SAIDA
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            Set lobjstcItem = pStcPedido.ColecaoPedLoc.Item(CStr(lobjstcCobertura.nosso_numero & _
                                                            Format(lobjstcCobertura.Num_Item, "00000")))
                                                                
            Set lobjstcCobertura1 = lobjstcItem.ColecaoPedCobert.Item(CStr(lobjstcCobertura.nosso_numero & _
                                                                      Format(lobjstcCobertura.Num_Item, "00000") & _
                                                                      Format(lobjstcCobertura.Cod_Cobert, "0000") & _
                                                                      Format(lobjstcCobertura.Tab_Cobert, "00000")))
            
            With lobjstcCobertura1
                .Num_Meses_PI = lobjstcCobertura.Num_Meses_PI
                .Formula = lobjstcCobertura.Formula
                .COD_CLASSE = lobjstcCobertura.COD_CLASSE
                .Val_Fator_Franq = lobjstcCobertura.Val_Fator_Franq
                .Alterar_Cobert = lobjstcCobertura.Alterar_Cobert
                .Val_Risco_Cob = lobjstcCobertura.Val_Risco_Cob
                .AjusteEndosso = lobjstcCobertura.Alterar_Cobert
                .AlterouFranquia = lobjstcCobertura.Alterar_Cobert
                .TextoFranquia = lobjstcCobertura.Alterar_Cobert
            End With
            Set lobjstcCobertura.ObjPai = lobjstcItem
            Set lobjstcCobertura1.ObjPai = lobjstcItem
            lobjstcItem.ColecaoPedCobert.Remove lobjstcCobertura.Class_ID
            lobjstcItem.ColecaoPedCobert.Add Item:=lobjstcCobertura1, Key:=lobjstcCobertura1.Class_ID
            
            pStcPedido.ColecaoPedLoc.Remove lobjstcItem.Class_ID
            pStcPedido.ColecaoPedLoc.Add Item:=lobjstcItem, Key:=lobjstcItem.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecCobertCompl = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecCobertCompl Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjCobertura = Nothing
    Set lobjstcCobertura = Nothing
    Return
    
End Function

Private Function mfRecBeneficiario(ByVal pNossoNumero As String, _
                                 ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                                 Optional pGravar As Boolean = True, _
                                 Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de observa��o

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    '                        pObjstcPedido........objeto de pedido.
    '                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
Dim lobjstcPedEspec        As stcA46V726B 'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC)
Dim lobjclsPedEspec        As clsA46V726A 'Objeto da classe de estrutura stcA46V726A (TAB_PED_ESPEC)
Dim lIndice                 As Integer
Dim lAux                    As Variant

Err.Clear
mfRecBeneficiario = False
On Error GoTo TrataErro
    '1. Grava registro de benefici�rios.

    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "018" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_ESPEC"
        GoTo SAIDA
    End If
    
    Set lobjstcPedEspec = New stcA46V726B
    
    lIndice = lIndice + 3
    
    '1. Grava registro de benefici�rios.
        
    With lobjstcPedEspec
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Observacao n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        

        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        
        
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        
        If mfRetCampo(2, "N�mero de p�ginas", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_pag = lAux
        
        If mfRetCampo(2, "Quantidade de vias", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Qtd_vias = lAux
        
        If mfRetCampo(4, "Texto - Linha 1", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin1 = lAux
        
        If mfRetCampo(4, "Texto - Linha 2", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin2 = lAux
        
        If mfRetCampo(4, "Texto - Linha 3", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin3 = lAux
        
        If mfRetCampo(4, "Texto - Linha 4", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin4 = lAux
        
        If mfRetCampo(4, "Texto - Linha 5", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin5 = lAux
        
        If mfRetCampo(4, "Texto - Linha 6", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin6 = lAux
        
        If mfRetCampo(4, "Texto - Linha 7", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin7 = lAux
        
        If mfRetCampo(4, "Texto - Linha 8", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin8 = lAux
        
        If mfRetCampo(4, "Texto - Linha 9", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin9 = lAux
        
        If mfRetCampo(4, "Texto - Linha 10", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin10 = lAux
        
        If mfRetCampo(4, "Texto - Linha 11", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin11 = lAux
        
        If mfRetCampo(4, "Texto - Linha 12", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin12 = lAux
        
        If mfRetCampo(4, "Texto - Linha 13", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin13 = lAux
        
        If mfRetCampo(4, "Texto - Linha 14", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin14 = lAux
        
        If mfRetCampo(4, "Texto - Linha 15", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin15 = lAux
        
        If mfRetCampo(4, "Texto - Linha 16", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin16 = lAux
        
        If mfRetCampo(4, "Texto - Linha 17", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin17 = lAux
        
        If mfRetCampo(4, "Texto - Linha 18", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin18 = lAux
        
        If mfRetCampo(4, "Texto - Linha 19", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin19 = lAux
        
        If mfRetCampo(4, "Texto - Linha 20", lAux, lIndice, 80, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Espec_lin20 = lAux
        
        If mfRetCampo(2, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
        
        If Mid(mRegistro, lIndice, 4) = "" Then
            .Cod_Espec = 0
        Else
            If mfRetCampo(2, "C�digo da Especifica��o", lAux, lIndice, 4, 0, pMensagem) = False Then
                    GoTo SAIDA
            End If
            .Cod_Espec = lAux
        End If
    End With
    
        
    Set lobjclsPedEspec = New clsA46V726A
    Set lobjclsPedEspec.M46V999 = M46V999
    
    'Temporario
    lobjstcPedEspec.nosso_numero = mNossoTempor
    
    If pGravar = True Then
        Set lobjclsPedEspec.ConexaoBD = bdBasDados
        Set lobjstcPedEspec.ObjPai = pStcPedido
        If Not lobjclsPedEspec.mfIncluir(lobjstcPedEspec, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjclsPedEspec.UltimoErro
            GoTo SAIDA
        End If
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoTextoEspec.Add Item:=lobjstcPedEspec, Key:=lobjstcPedEspec.Class_ID
        End If
    Else
     'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoTextoEspec.Add Item:=lobjstcPedEspec, Key:=lobjstcPedEspec.Class_ID
        Else
            pMensagem = "mfRecPedido Erro: Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecBeneficiario = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecPedido Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjstcPedEspec = Nothing
    Set lobjstcPedEspec = Nothing
    Return

End Function
Private Function mfRecLMI(ByVal pNossoNumero As String, _
                                 ByRef pMensagem As String, ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                                 Optional pGravar As Boolean = True, _
                                 Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de coberturas para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
    Dim lobjLMI                 As clsA46V732A      'Objeto da classe clsA46V732A (TAB_PED_LMI)
    Dim lobjstcLMI              As stcA46V732B      'Objeto da classe de estrutura stcA46V732B (TAB_PED_LMI)
    Dim lIndice                 As Integer
    Dim lAux                    As Variant
    
Err.Clear
mfRecLMI = False
On Error GoTo TrataErro

    '1. Grava registro de coberturas.

    lIndice = 1
    If Mid(mRegistro, lIndice, 3) <> "020" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_LMI"
        GoTo SAIDA
    End If
            
    Set lobjstcLMI = New stcA46V732B
    Set lobjstcLMI.M46V999 = M46V999
    
    lIndice = lIndice + 3
        
    With lobjstcLMI
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Coberturas n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "C�digo do org�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "C�digo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "N�mero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "N�mero da ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "N�mero do endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "Tab Cobertura", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tab_Cobert = lAux
        If mfRetCampo(2, "C�digo da Cobertura", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Cobert = lAux
        If mfRetCampo(2, "Valor da IS", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_IS = lAux
        If mfRetCampo(2, "Percentual da franquia", lAux, lIndice, 4, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Per_Franq = lAux
        If mfRetCampo(2, "Valor m�nimo da franquia", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Franq_Min = lAux
        If mfRetCampo(2, "Valor m�ximo da franquia", lAux, lIndice, 15, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Val_Franq_Max = lAux
        If mfRetCampo(2, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
        If mfRetCampo(2, "N�mero de meses do PI", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Meses_PI = lAux
    End With
    
    Set lobjLMI = New clsA46V732A
    Set lobjLMI.M46V999 = M46V999

    lobjstcLMI.nosso_numero = mNossoTempor

    If pGravar = True Then
        Set lobjLMI.ConexaoBD = bdBasDados
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoLMI.Remove lobjstcLMI.Class_ID
            pStcPedido.ColecaoLMI.Add Item:=lobjstcLMI, Key:=lobjstcLMI.Class_ID
        End If
        If Not lobjLMI.mfIncluir(lobjstcLMI, pPropTabela, pNomeTabImport) Then
            pMensagem = lobjLMI.UltimoErro
            GoTo SAIDA
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            pStcPedido.ColecaoLMI.Remove lobjstcLMI.Class_ID
            pStcPedido.ColecaoLMI.Add Item:=lobjstcLMI, Key:=lobjstcLMI.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecLMI = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecLMI Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjLMI = Nothing
    Set lobjstcLMI = Nothing
    Return
End Function

Private Function mfRecPendenciasDiversas(ByVal pNossoNumero As String, _
                                 ByRef pMensagem As String, ByVal pPropTabela As String, _
                                 ByVal pNomeTabImport As String, _
                                 Optional pGravar As Boolean = True, _
                                 Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de Pend�ncias Diversos - Num primeiro momento ser� utilizada somente para o Sies mas pode ser utilizada de v�rias formas

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    '                        pObjstcPedido........objeto de pedido.
    '                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
Dim lTipoInformacao     As String
Dim lInsert             As String
Dim lBDP00Multi         As ADODB.Connection
Dim lIndice             As Integer
Dim lAux                    As Variant

Err.Clear
mfRecPendenciasDiversas = False
On Error GoTo TrataErro

    lIndice = 1
            
    If Mid(mRegistro, lIndice, 3) <> "019" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_INF_ADIC"
        GoTo SAIDA
    End If
    
    lIndice = lIndice + 3
    
    '1. Verifica que tipo de Informa��o �
        

    If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
        GoTo SAIDA
    End If
    If lAux <> pNossoNumero Then
        pMensagem = "Registro de Observacao n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
        GoTo SAIDA
    End If
        

    If mfRetCampo(4, "Tipo de Informacao Adicional ", lAux, lIndice, 2, 0, pMensagem) = False Then
        GoTo SAIDA
    End If
    lTipoInformacao = lAux
    If lTipoInformacao = "01" Then    'Gerar Pend�ncias de Verifica��o de Arquivos no SIES
        If mfRetCampo(4, "Pendencias", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        If lAux = "SIM" Then
            If pGravar = True Then
            
                If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lBDP00Multi, pMensagem) = False Then
                    Call M46V999.gpGraLog(0, pMensagem)
                    GoTo SAIDA
                End If
            
                lInsert = " INSERT INTO ged.dbo.tab_Prop_Mensagem(num_pi, item_pi, cod_mensagem, cod_tip_mensagem, motivo, dat_digit, dat_liber, hor_liber) "
                lInsert = lInsert & " VALUES(" & pStcPedido.NUM_PI & "," & 0 & ",9199,2,''," & Format(Date, "yyyymmdd") & ",0,0)"
                If M46V999.gfExeSQL(lBDP00Multi, lInsert, pMensagem) <> 0 Then
                    Call M46V999.gpGraLog(0, pMensagem)
                    GoTo SAIDA
                End If
            End If
        End If
    End If
    
'================================
'******Fun��o realizada com sucesso
    mfRecPendenciasDiversas = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecPendenciasDiversas Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lBDP00Multi = Nothing
    Return

End Function



'#####INSPECAO
Public Function gfRecArqTransInsp(ByRef pNomArqTransmissao As String, _
                                    ByRef pMensagem As String, _
                                    Optional pstcTransInsp As stcA46V730B) As Boolean
'**********************
'Fun��o compartilhada com o projeto P24V007.exe (Servidor SYAS-EDI executado como servi�o)
'        N�O PODE EMITIR MENSAGEM!!!! PROIBIDO MSGBOX!!!!
'
'**********************
    'Fun��o: l� arquivos de transmiss�o.

Dim lAppPath            As String           'App.Path com \ no final.
Dim lRegistro           As String           'Registro.
Dim lSequencia          As String           'Seq��ncia para o arquivo de transmiss�o.
Dim lTotal              As Double           'Total.
Dim lobjstcTransInsp    As stcA46V730B      'Objeto da classe de estrutura stcA46V702B (TAB_TRANS_INSP)
Dim lobjstcEquipInsp    As stcA46v740B      'Objeto da classe de estrutura stcA46V707B (TAB_PED_INSP)
Dim lNumArqTransmissao  As Integer          'Utilizado com FreeFile para arquivo de transmiss�o.
Dim lCodRegistro        As String
Dim lNossoNumero        As String
Dim lFlagGravar         As Boolean
Dim lPriVez             As Boolean
Dim lArqAberto          As Boolean
Dim lVersao             As String           'Utilizado para verificar se o arquivo transmitido foi gerado com a implementa��o do Projeto Al�ada
Dim lNomeTabImport      As String
Dim lBeginPendente      As Boolean
Dim lIdenRegistro       As String           'Utilizado para verificar � o novo registro de Inspe��o - Equipamentos
Err = 0
On Error GoTo Func_Erro
    gfRecArqTransInsp = False
    pMensagem = ""
    lPriVez = True
    
'1. Verifica se esta sendo executado pelo SYAS-EDI
    lNomeTabImport = ""
    If UCase(App.EXEName) = "P24V007" Then
        lFlagGravar = False
    Else
        lFlagGravar = True
        lNomeTabImport = "FTP_"
    End If
    
'2. Posiciona controle.
     If Right$(App.Path, 1) = "\" Then
        lAppPath = App.Path
    Else
        lAppPath = App.Path & "\"
    End If

    If Not M46V999.gfPreenchido(Dir(pNomArqTransmissao)) Then
        pMensagem = "FTE0019 - Arquivo n�o localizado: " & pNomArqTransmissao & _
                    " (grava arquivo de transmiss�o)."
        GoTo SAIDA
    End If

    lTotal = 0
    'Abre banco de Dados
    If lFlagGravar = True Then
        If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, bdBasDados, _
                pMensagem) = False Then
                GoTo SAIDA
        End If
    End If
        
    'Abre o arquivo
    lNumArqTransmissao = FreeFile
    lArqAberto = False
    Open pNomArqTransmissao For Input As #lNumArqTransmissao
        lArqAberto = True
        'L� a primeira linha do arquivo
        Line Input #lNumArqTransmissao, mRegistro
        Do

'##PROJETO INSPECAO
            lVersao = Mid(mRegistro, 1, 2)
            If lVersao = "V2" Then
                lNossoNumero = Mid(mRegistro, 4, 20)
            Else
                ' Inspe��o � Equipamentos
                lIdenRegistro = Mid(mRegistro, 1, 3)
                If lIdenRegistro = "023" Then
                    lNossoNumero = Mid(mRegistro, 4, 20)
                Else
                lNossoNumero = Mid(mRegistro, 1, 20)
            End If
            End If
            
            If Not IsNumeric(lNossoNumero) Then
                pMensagem = "FTE0017 - Erro ao criar arquivo de transmiss�o (Nosso N�mero n�o num�rico)."
                Exit Do
            End If
            
            If lVersao = "V2" Then
                'Grava registros de Transmiss�o Inspe��o. (NOVO)
                If mfRecTransInspecaoV2(lNossoNumero, pMensagem, lFlagGravar, lobjstcTransInsp) = False Then
                   Exit Do
                End If
            Else
               'Grava registros de Transmiss�o Inspe��o. (ANTIGO)
                If mfRecTransInspecao(lNossoNumero, pMensagem, lFlagGravar, lobjstcTransInsp) = False Then
                   Exit Do
                End If
            End If
            
NextRecord:
            If EOF(lNumArqTransmissao) Then
                Exit Do
            End If
            Line Input #lNumArqTransmissao, mRegistro
        Loop
    
FecharArquivo:
    '7. Fecha arquivo - ser� utilizado no pr�ximo comando.
    Close #lNumArqTransmissao
    lArqAberto = False      'marca que o arquivo n�o est� mais aberto
    
    If M46V999.gfPreenchido(pMensagem) = True Then
        GoTo Func_Erro
    End If
    If Not lPriVez And lFlagGravar And lBeginPendente Then
        Dummy = M46V999.gfExeSQL(bdBasDados, "Commit Transaction ", pMensagem)
    End If
    If M46V999.gfPreenchido(pMensagem) = True Then
        GoTo Func_Erro
    End If

    '8. Grava registro de controle no final do arquivo (�ltimo registro do arquivo que serve para
    '   verificar se arquivo enviado pelo corretor manteve sua integridade ao ser processado pelo
    '   servidor YASUDA).
    '   Formato do registro de controle: "PROPOSTA XXXXXXXXXXX", onde XXXXXXXXXXX � o valor da soma dos
    '   valores ASC de todos os caracteres das linhas anteriores, gravado de forma invertida.
    '   Por exemplo: Se soma dos valores ASC de todos os caracteres dos registros de um arquivo
    '   (excetuando-se �ltimo registro) for igual a 39.485, �ltimo registro ser�: PROPOSTA 58493.
    'a) Soma valores ASC do registro.
    lTotal = 0
    
    If lFlagGravar = False Then
        If Not IsMissing(pstcTransInsp) Then
            Set pstcTransInsp = lobjstcTransInsp
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
        End If
    End If

    gfRecArqTransInsp = True
    
SAIDA:
    If lArqAberto = True Then
        'Fecha o arquivo caso o c�digo tenha sido desviado para o erro
        'ap�s ter aberto o arquivo.
        Close #lNumArqTransmissao
        lArqAberto = False
    End If
    Err = 0
    On Error GoTo 0
    Exit Function
    
Func_Erro:
    If Not lPriVez And lFlagGravar And lBeginPendente Then
        Dummy = M46V999.gfExeSQL(bdBasDados, "Rollback Transaction ", pMensagem)
    End If
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "gfRecArqTransInsp Erro " & pMensagem
    
    GoTo SAIDA
    
End Function

Private Function mfRecTransInspecao(ByVal pNossoNumero As String, _
                                 ByRef pMensagem As String, _
                                 Optional pGravar As Boolean = True, _
                                 Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de transmiss�o inspe��o

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    '                        pObjstcPedido........objeto de pedido.
    '                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
    Dim lobjstcTransInsp       As stcA46V730B 'Objeto da classe de estrutura stcA46V726B (TAB_TRANS_INSP)
    Dim lobjclsTransInsp       As clsA46V730A 'Objeto da classe de estrutura stcA46V726A (TAB_TRANS_INSP)
         
    mfRecTransInspecao = False

    '1. Grava registro de transmiss�o inspe��o.
    Dim lIndice                 As Integer
    Dim lAux                    As Variant
    Dim lSelect                 As String
    Dim lrst                    As clsYasRecordSet
    Dim bdBasSiscota            As ADODB.Connection 'Base de dados.
  
    lIndice = lIndice + 1
                
    Set lobjstcTransInsp = New stcA46V730B
           
    With lobjstcTransInsp
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .nosso_numero = lAux
        mNossoTempor = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Observacao n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            Exit Function
        End If
        
        lIndice = lIndice + 1
        If mfRetCampo(2, "N�mero do Pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Ped = lAux
                 
        lIndice = lIndice + 1
        If mfRetCampo(2, "N�mero do Item", lAux, lIndice, 5, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Item = lAux
       
        lIndice = lIndice + 1
        If mfRetCampo(2, "N�mero da Ap�lice anterior", lAux, lIndice, 10, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Apol_Ant = lAux
   
        lIndice = lIndice + 1
        If mfRetCampo(4, "Nome Segurado", lAux, lIndice, 60, 0, pMensagem) = False Then
            Exit Function
        End If
        .Nom_Segurado = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "N�mero do Telefone", lAux, lIndice, 14, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Tel = lAux

        lIndice = lIndice + 1
        If mfRetCampo(6, "N�mero do CGC/CPF", lAux, lIndice, 14, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Cgc_Cpf = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo do Ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Ramo = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo do �rg�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Org = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "Data de In�cio Vig�ncia", lAux, lIndice, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        .Dat_Ini_Vig = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "Data de T�rmino Vig�ncia", lAux, lIndice, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        .Dat_Ter_Vig = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Tipo de Logradouro", lAux, lIndice, 3, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Tip_Logr_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Nome do Logradouro", lAux, lIndice, 50, 0, pMensagem) = False Then
            Exit Function
        End If
        .Nom_Logr_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Descri��o do Logradouro", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Dsc_Compl_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Nome do Bairro", lAux, lIndice, 30, 0, pMensagem) = False Then
            Exit Function
        End If
        .Nom_Bairro_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Nome da Cidade", lAux, lIndice, 30, 0, pMensagem) = False Then
            Exit Function
        End If
        .Nom_Cid_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Sigla do Estado", lAux, lIndice, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        .Sig_UF_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "N�mero do CEP", lAux, lIndice, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Cep_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "N�mero do Risco", lAux, lIndice, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "Data de Transmiss�o", lAux, lIndice, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        .Dat_Trans = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "Tipo de Pessoa", lAux, lIndice, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Iden_Pessoa = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo da Unidade Produtiva", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Unid_Prod = lAux
        
        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo do Produtor", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Prod = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo do Corretor", lAux, lIndice, 7, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Corr = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo da Cobertura 01", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_01 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo da Cobertura 02", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_02 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo da Cobertura 03", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_03 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo da Cobertura 04", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_04 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo da Cobertura 05", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_05 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo da Cobertura 06", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_06 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo da Cobertura 07", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_07 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo da Cobertura 08", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_08 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo da Cobertura 09", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_09 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo da Cobertura 10", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_10 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 01", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_01 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 02", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_02 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 03", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_03 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 04", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_04 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 05", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_05 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 06", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_06 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 07", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_07 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 08", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_08 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 09", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_09 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 10", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_10 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "C�digo da Classifica��o", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Clasf = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Complemento da Classifica��o", lAux, lIndice, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Compl_Clasf = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Contato", lAux, lIndice, 40, 0, pMensagem) = False Then
            Exit Function
        End If
        .Pess_Cont = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo do Produto", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Produto = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "Flag Transmiss�o", lAux, lIndice, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        .Flag_Trans = lAux
        .Num_Cotac = ""
        
        'Ficha 1000157 - Simony
        Set bdBasSiscota = bdBasDados
        If UCase(bdBasSiscota) Like "*MIX*" Then
            'Rotina s� para Teste - Abre banco de Dados para consulta do siscota de produ��o
            M46V999.gSerDIVERSOS = "LPX20102PSLUSQL"
            If M46V999.gfAbrBasDados("RAMOSDIVERSOS", "2", M46V999.gEndDIVERSOS, bdBasSiscota, pMensagem) = False Then
                GoTo SAIDA
            End If
            M46V999.gSerDIVERSOS = "SPX20112CSLUMIX"
        End If
        
        lSelect = " Select CCOTACAO from siscota.dbo.tab_syas_cotacao where " & _
            " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(mNossoTempor)
        If Not M46V999.gfObtRegistro(bdBasSiscota, lSelect, lrst, pMensagem) Then
            Exit Function
        End If
        If Not lrst.EOF Then
           .Num_Cotac = M46V999.Check(lrst("CCotacao"))
        End If
    End With
            
    Set lobjclsTransInsp = New clsA46V730A
    Set lobjclsTransInsp.M46V999 = M46V999
    lobjstcTransInsp.nosso_numero = mNossoTempor
    
    If pGravar = True Then
        Set lobjclsTransInsp.ConexaoBD = bdBasDados
        If Not lobjclsTransInsp.gfIncluir(lobjstcTransInsp) Then
            pMensagem = lobjclsTransInsp.UltimoErro
            Exit Function
        End If
    End If

    mfRecTransInspecao = True
SAIDA:
    Set lobjclsTransInsp = Nothing
    Set lobjstcTransInsp = Nothing

End Function

'## PROJETO INSPECAO
Private Function mfRecTransInspecaoV2(ByVal pNossoNumero As String, _
                                 ByRef pMensagem As String, _
                                 Optional pGravar As Boolean = True, _
                                 Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de transmiss�o inspe��o

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    '                        pObjstcPedido........objeto de pedido.
    '                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
    Dim lobjstcTransInsp       As stcA46V730B 'Objeto da classe de estrutura stcA46V730B (TAB_TRANS_INSP)
    Dim lobjclsTransInsp       As clsA46V730A 'Objeto da classe de estrutura stcA46V730A (TAB_TRANS_INSP)
    Dim lobjstcTransCobInsp    As stcA46V731B 'Objeto da classe de estrutura stcA46V731B (TAB_TRANS_INSP)
    Dim lobjclsTransCobInsp    As clsA46V731A 'Objeto da classe de estrutura stcA46V731A (TAB_TRANS_INSP)
     
    mfRecTransInspecaoV2 = False

    '1. Grava registro de transmiss�o inspe��o.
    Dim lIndice                 As Integer
    Dim lAux                    As Variant
    Dim lSelect                 As String
    Dim lrst                    As clsYasRecordSet
    Dim bdBasSiscota            As ADODB.Connection 'Base de dados.
    Dim lGravaIndice            As Integer
  
    lIndice = 3 'V2;
    lIndice = lIndice + 1
                
    Set lobjstcTransInsp = New stcA46V730B
    Set lobjstcTransCobInsp = New stcA46V731B
        
    With lobjstcTransInsp
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .nosso_numero = lAux
        lobjstcTransCobInsp.nosso_numero = lAux
        
        mNossoTempor = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Observacao n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            Exit Function
        End If
        
        lIndice = lIndice + 1
        If mfRetCampo(2, "N�mero do Pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Ped = lAux
        lobjstcTransCobInsp.Num_Ped = lAux
        
        lIndice = lIndice + 1
        If mfRetCampo(2, "N�mero do Item", lAux, lIndice, 5, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Item = lAux
        lobjstcTransCobInsp.Num_Item = lAux
        
        lIndice = lIndice + 1
        If mfRetCampo(2, "N�mero da Ap�lice anterior", lAux, lIndice, 10, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Apol_Ant = lAux
   
        lIndice = lIndice + 1
        If mfRetCampo(4, "Nome Segurado", lAux, lIndice, 60, 0, pMensagem) = False Then
            Exit Function
        End If
        .Nom_Segurado = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "N�mero do Telefone", lAux, lIndice, 14, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Tel = lAux

        lIndice = lIndice + 1
        If mfRetCampo(6, "N�mero do CGC/CPF", lAux, lIndice, 14, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Cgc_Cpf = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo do Ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Ramo = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo do �rg�o", lAux, lIndice, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Org = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "Data de In�cio Vig�ncia", lAux, lIndice, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        .Dat_Ini_Vig = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "Data de T�rmino Vig�ncia", lAux, lIndice, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        .Dat_Ter_Vig = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Tipo de Logradouro", lAux, lIndice, 3, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Tip_Logr_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Nome do Logradouro", lAux, lIndice, 50, 0, pMensagem) = False Then
            Exit Function
        End If
        .Nom_Logr_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Descri��o do Logradouro", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Dsc_Compl_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Nome do Bairro", lAux, lIndice, 30, 0, pMensagem) = False Then
            Exit Function
        End If
        .Nom_Bairro_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Nome da Cidade", lAux, lIndice, 30, 0, pMensagem) = False Then
            Exit Function
        End If
        .Nom_Cid_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Sigla do Estado", lAux, lIndice, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        .Sig_UF_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "N�mero do CEP", lAux, lIndice, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Cep_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "N�mero do Risco", lAux, lIndice, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Risco = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "Data de Transmiss�o", lAux, lIndice, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        .Dat_Trans = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "Tipo de Pessoa", lAux, lIndice, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Iden_Pessoa = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo da Unidade Produtiva", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Unid_Prod = lAux
        
        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo do Produtor", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Prod = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo do Corretor", lAux, lIndice, 7, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Corr = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "C�digo da Classifica��o", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Clasf = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Complemento da Classifica��o", lAux, lIndice, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Compl_Clasf = lAux

        lIndice = lIndice + 1
        If mfRetCampo(4, "Contato", lAux, lIndice, 40, 0, pMensagem) = False Then
            Exit Function
        End If
        .Pess_Cont = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "C�digo do Produto", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Produto = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "Flag Transmiss�o", lAux, lIndice, 1, 0, pMensagem) = False Then
            Exit Function
        End If
        .Flag_Trans = lAux
        .Num_Cotac = ""
        lobjstcTransCobInsp.Num_Cotac = ""
        
        'Ficha 1000157 - Simony
        Set bdBasSiscota = bdBasDados
        If UCase(bdBasSiscota) Like "*MIX*" Then
            'Rotina s� para Teste - Abre banco de Dados para consulta do siscota de produ��o
            M46V999.gSerDIVERSOS = "LPX20102PSLUSQL"
            If M46V999.gfAbrBasDados("RAMOSDIVERSOS", "2", M46V999.gEndDIVERSOS, bdBasSiscota, pMensagem) = False Then
                GoTo SAIDA
            End If
            M46V999.gSerDIVERSOS = "SPX20112CSLUMIX"
        End If

        lSelect = " Select CCOTACAO from siscota.dbo.tab_syas_cotacao where " & _
            " Nosso_Numero = " & M46V999.gfFormatarCamposSQL(mNossoTempor)
        If Not M46V999.gfObtRegistro(bdBasSiscota, lSelect, lrst, pMensagem) Then
            Exit Function
        End If
        If Not lrst.EOF Then
           .Num_Cotac = M46V999.Check(lrst("CCotacao"))
           lobjstcTransCobInsp.Num_Cotac = M46V999.Check(lrst("CCotacao"))
        End If
        
        lGravaIndice = lIndice
        
        lIndice = lIndice + 1
        If Len(mRegistro) < lIndice + 4 Then
            GoTo Fim
        End If
        If mfRetCampo(2, "C�digo da Cobertura 01", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_01 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 01", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_01 = lAux
        
        lIndice = lIndice + 1
        If Len(mRegistro) < lIndice + 4 Then
            GoTo Fim
        End If
        If mfRetCampo(2, "C�digo da Cobertura 02", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_02 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 02", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_02 = lAux
        
        lIndice = lIndice + 1
        If Len(mRegistro) < lIndice + 4 Then
            GoTo Fim
        End If
        If mfRetCampo(2, "C�digo da Cobertura 03", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_03 = lAux
        
        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 03", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_03 = lAux


        lIndice = lIndice + 1
        If Len(mRegistro) < lIndice + 4 Then
            GoTo Fim
        End If
        If mfRetCampo(2, "C�digo da Cobertura 04", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_04 = lAux
        
        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 04", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_04 = lAux
        

        lIndice = lIndice + 1
        If Len(mRegistro) < lIndice + 4 Then
            GoTo Fim
        End If
        If mfRetCampo(2, "C�digo da Cobertura 05", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_05 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 05", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_05 = lAux
        
        lIndice = lIndice + 1
        If Len(mRegistro) < lIndice + 4 Then
            GoTo Fim
        End If
        If mfRetCampo(2, "C�digo da Cobertura 06", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_06 = lAux
        
        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 06", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_06 = lAux
        

        lIndice = lIndice + 1
        If Len(mRegistro) < lIndice + 4 Then
            GoTo Fim
        End If
        If mfRetCampo(2, "C�digo da Cobertura 07", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_07 = lAux


        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 07", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_07 = lAux
        
        lIndice = lIndice + 1
        If Len(mRegistro) < lIndice + 4 Then
            GoTo Fim
        End If
        If mfRetCampo(2, "C�digo da Cobertura 08", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_08 = lAux
        
        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 08", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_08 = lAux
        

        lIndice = lIndice + 1
        If Len(mRegistro) < lIndice + 4 Then
            GoTo Fim
        End If
        If mfRetCampo(2, "C�digo da Cobertura 09", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_09 = lAux
        
        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 09", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_09 = lAux
        

        lIndice = lIndice + 1
        If Len(mRegistro) < lIndice + 4 Then
            GoTo Fim
        End If
        If mfRetCampo(2, "C�digo da Cobertura 10", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .Cod_Cobert_10 = lAux

        lIndice = lIndice + 1
        If mfRetCampo(2, "IS da Cobertura 10", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .Val_IS_Cobert_10 = lAux
        
Fim:
        .Num_Insp = 0
        .Dat_Final = 0
        .Dat_Cancel = 0
        .Cod_Status_I = 1
        .Dat_Confirm = Format(M46V999.gForHoje, "YYYYMMDD")
        
    End With
            
    Set lobjclsTransInsp = New clsA46V730A
    Set lobjclsTransInsp.M46V999 = M46V999
    lobjstcTransInsp.nosso_numero = mNossoTempor
    
    If pGravar = True Then
        Set lobjclsTransInsp.ConexaoBD = bdBasDados
        If Not lobjclsTransInsp.gfIncluir(lobjstcTransInsp) Then
            pMensagem = lobjclsTransInsp.UltimoErro
            Exit Function
        End If
    End If
    
    If lobjclsTransInsp.UltimoErro = "DUPLICADO" Then
        mfRecTransInspecaoV2 = True
        GoTo SAIDA
    End If
    
    Dim lCob As Long
    
    lIndice = lGravaIndice
    lCob = 0
    For lCob = 1 To 250
        lIndice = lIndice + 1
        Call mfRetCampo(2, "C�digo da Cobertura", lAux, lIndice, 4, 0, pMensagem)
        lobjstcTransCobInsp.Cod_Cobert = lAux
        If lAux = 0 Then
            pMensagem = ""
            Exit For
        End If
        lIndice = lIndice + 1
        Call mfRetCampo(2, "IS da Cobertura", lAux, lIndice, 20, 0, pMensagem)
        lobjstcTransCobInsp.Lmg_Cobert = lAux
        lAux = 0
        
        Set lobjclsTransCobInsp = New clsA46V731A
        Set lobjclsTransCobInsp.M46V999 = M46V999
        lobjstcTransCobInsp.nosso_numero = mNossoTempor
        
        If pGravar = True Then
            Set lobjclsTransCobInsp.ConexaoBD = bdBasDados
            If Not lobjclsTransCobInsp.gfIncluir(lobjstcTransCobInsp) Then
                pMensagem = lobjclsTransCobInsp.UltimoErro
                Exit Function
            End If
        End If

        Set lobjclsTransCobInsp = Nothing
    Next lCob
   
    mfRecTransInspecaoV2 = True
SAIDA:
    Set lobjclsTransInsp = Nothing
    Set lobjstcTransInsp = Nothing
    Set lobjstcTransCobInsp = Nothing
End Function

Private Function mfRecProLabore(ByVal pNossoNumero As String, _
                                ByRef pMensagem As String, ByVal pPropTabela As String, _
                                ByVal pNomeTabImport As String, _
                                Optional pGravar As Boolean = True, _
                                Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de coberturas para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
    Dim lobjProLabore                 As clsA46V733A      'Objeto da classe clsA46V733A (TAB_PED_ProLabore)
    Dim lobjstcProlabore              As stcA46V733B      'Objeto da classe de estrutura stcA46V733B (TAB_PED_ProLabore)
    Dim lIndice                 As Integer
    Dim lAux                    As Variant
    
Err.Clear
mfRecProLabore = False
On Error GoTo TrataErro

    '1. Grava registro de coberturas.

    lIndice = 1
    If Mid(mRegistro, lIndice, 3) <> "021" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_CANAL_EMISS"
        GoTo SAIDA
    End If
            
    Set lobjstcProlabore = New stcA46V733B
    'Set lobjstcProLabore.ObjPai = pStcPedido
    Set lobjstcProlabore.M46V999 = M46V999
    
    lIndice = lIndice + 3
        
    With lobjstcProlabore
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Pro-labore n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "Sequ�ncia", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_SEQ = lAux
        If mfRetCampo(2, "Num PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
        If mfRetCampo(2, "Apolice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "Endosso/Fatura", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Compl = lAux
        If mfRetCampo(2, "Tipo de Canal", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tip_Canal = lAux
        If mfRetCampo(2, "Pro-laborista", lAux, lIndice, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Canal = lAux
        If mfRetCampo(2, "Corretor", lAux, lIndice, 7, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Corretor = lAux
        If mfRetCampo(2, "Perc. Comissao", lAux, lIndice, 6, 3, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Perc_Canal = lAux
        If mfRetCampo(2, "Qtd. Parc. a pagar", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Qtd_Parc_Pag = lAux
        If mfRetCampo(2, "Qtd. Parc. do Contrato", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Qtd_Parc_Apol = lAux
        If mfRetCampo(2, "Forma de pagamento", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Form_Pag_Ag = lAux
        
    End With
    
    Set lobjProLabore = New clsA46V733A
    Set lobjProLabore.M46V999 = M46V999

    lobjstcProlabore.nosso_numero = mNossoTempor

    If pGravar = True Then
        Set lobjProLabore.ConexaoBD = bdBasDados
        Set lobjstcProlabore.ObjPai = pStcPedido
        If Not lobjProLabore.mfIncluir(lobjstcProlabore, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjProLabore.UltimoErro
            GoTo SAIDA
        End If
        If Not IsMissing(pStcPedido) Then
            'pStcPedido.ColecaoProlabore.Remove lobjstcProlabore.Class_ID
            pStcPedido.ColecaoProlabore.Add Item:=lobjstcProlabore, Key:=lobjstcProlabore.Class_ID
        End If
    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            'pStcPedido.ColecaoProlabore.Remove lobjstcProlabore.Class_ID
            pStcPedido.ColecaoProlabore.Add Item:=lobjstcProlabore, Key:=lobjstcProlabore.Class_ID
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecProLabore = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecProLabore Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjProLabore = Nothing
    Set lobjstcProlabore = Nothing
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

Private Function mfRecMovAgen(ByVal pNossoNumero As String, _
                                ByRef pMensagem As String, ByVal pPropTabela As String, _
                                ByVal pNomeTabImport As String, _
                                Optional pGravar As Boolean = True, _
                                Optional pStcPedido As stcA46V702B) As Boolean
        'Fun��o: grava registro de coberturas para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lobjMovAgen                 As clsA46V743A      'Objeto da classe clsA46V743A (TAB_MOV_Agen)
    Dim lobjstcMovAgen              As stcA46V743B      'Objeto da classe de estrutura stcA46V743B (TAB_MOV_Agen)
    Dim lIndice                 As Integer
    Dim lAux                    As Variant

Err.Clear
mfRecMovAgen = False
On Error GoTo TrataErro

    '1. Grava registro de coberturas.

    lIndice = 1
    If Mid(mRegistro, lIndice, 3) <> "022" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_MOV_AGEN"
        GoTo SAIDA
    End If
            
    Set lobjstcMovAgen = New stcA46V743B

    lIndice = lIndice + 3
        
    With lobjstcMovAgen
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Mov Agen n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "Sequ�ncia", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_SEQ = lAux
        If mfRetCampo(2, "Num PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
        If mfRetCampo(2, "Apolice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_APOL_MOV = lAux
        If mfRetCampo(2, "Endosso/Fatura", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_COMPL_MOV = lAux
        If mfRetCampo(2, "Referencia do Movimento", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .REF_MOV = lAux
        If mfRetCampo(4, "Tipo de movimento", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .TIPO_MOV = lAux
        If mfRetCampo(2, "Valor do Premio L�quido", lAux, lIndice, 11, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .VL_PREMIO_LIQ_MOV = lAux
        If mfRetCampo(2, "Qtd. de Itens", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .QTD_ITEM_MOV = lAux
        If mfRetCampo(4, "Data da grava��o", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .DAT_GRAV = lAux
'        If mfRetCampo(2, "Usu�rio da Grava��o", lAux, lIndice, 10, 0, pMensagem) = False Then
'            GoTo SAIDA
'        End If
        .Cod_User = lAux
       
    End With
    
    Set lobjMovAgen = New clsA46V743A
    Set lobjMovAgen.M46V999 = M46V999

    lobjstcMovAgen.nosso_numero = mNossoTempor

    If pGravar = True Then
        Set lobjMovAgen.ConexaoBD = bdBasDados
        Set lobjstcMovAgen.ObjPai = pStcPedido
        If Not lobjMovAgen.mfIncluir(lobjstcMovAgen, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjMovAgen.UltimoErro
            GoTo SAIDA
        End If
        If Not IsMissing(pStcPedido) Then
            'pStcPedido.ColecaoMovAgen.Remove lobjstcMovAgen.Class_ID
            Set pStcPedido.objstcMovAgen = lobjstcMovAgen
        End If

    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedido) Then
            Set pStcPedido.objstcMovAgen = lobjstcMovAgen
        Else
            pMensagem = "Gravar=Falso e pstcPedido n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecMovAgen = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecMovAgen Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjMovAgen = Nothing
    Set lobjstcMovAgen = Nothing
    Return
End Function

Private Function mfRecPedTelefones(ByVal pNossoNumero As String, _
                                ByRef pMensagem As String, ByVal pPropTabela As String, _
                                ByVal pNomeTabImport As String, _
                                Optional pGravar As Boolean = True, _
                                Optional pStcPedido As stcA46V702B) As Boolean
    'Fun��o: grava registro de coberturas para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lobjPedTelefones As New clsA46V736A
    Dim lobjstcPedTelefones As New stcA46V736B
    Dim lIndice As Integer
    Dim lAux As Variant

    Err.Clear
    mfRecPedTelefones = False
    On Error GoTo TrataErro

    '1. Grava registro de coberturas.

    lIndice = 1
    If Mid(mRegistro, lIndice, 3) <> "024" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_TELEFONES"
        GoTo SAIDA
    End If

    Set lobjstcPedTelefones = New stcA46V736B
    Set lobjstcPedTelefones.M46V999 = M46V999

    lIndice = lIndice + 3

    With lobjstcPedTelefones
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Telefones n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "N�mero de PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
        If mfRetCampo(2, "Numero de ap�lice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "Numero de endosso", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "Tipo de telefone", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tip_Telefone = lAux
        If mfRetCampo(2, "DDD do telefone", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .DDD_Telefone = lAux
        If mfRetCampo(2, "N�mero do telefone", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Numero_Telefone = lAux
        If mfRetCampo(2, "Ramal do telefone", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Ramal_Telefone = lAux
        If mfRetCampo(4, "Observa��o do telefone", lAux, lIndice, 50, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Obs_Telefone = lAux
        If mfRetCampo(2, "Telefone principal", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Ind_Principal = lAux

    End With

    Set lobjPedTelefones = New clsA46V736A
    Set lobjPedTelefones.M46V999 = M46V999

    lobjstcPedTelefones.nosso_numero = mNossoTempor

    If pGravar = True Then
        Set lobjPedTelefones.ConexaoBD = bdBasDados
        Set lobjstcPedTelefones.ObjPai = pStcPedido
        If Not lobjPedTelefones.mfIncluir(lobjstcPedTelefones, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjPedTelefones.UltimoErro
            GoTo SAIDA
        End If
        'If Not IsMissing(pStcPedido) Then
        '    pStcPedido.ColecaoTelefones.Add Item:=lobjstcPedTelefones, Key:=lobjstcPedTelefones.Class_ID
        'End If
    'Else
        'Retorna no parametro de entrada
        'If Not IsMissing(pStcPedido) Then
        '    pStcPedido.ColecaoTelefones.Add Item:=lobjstcPedTelefones, Key:=lobjstcPedTelefones.Class_ID
        'Else
        '    pMensagem = "Gravar=Falso e pstcPedido n�o informado."
        '    GoTo SAIDA
        'End If
    End If
    '================================
    '******Fun��o realizada com sucesso
    mfRecPedTelefones = True
    '================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function

TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecPedTelefones Erro: " & pMensagem
    GoTo SAIDA

FechaObj:
    Set lobjPedTelefones = Nothing
    Set lobjstcPedTelefones = Nothing
    Return
End Function
        
' Inspe��o � Equipamentos
Private Function mfRecTransInspecaoEquipamentos(ByVal pNossoNumero As String, _
                                 ByRef pMensagem As String, _
                                 ByVal pNomeTabImport As String, _
                                 Optional pGravar As Boolean = True, _
                                 Optional pStcPedido As stcA46V702B) As Boolean
     'Fun��o: grava registro de transmiss�o inspe��o de equipamentos

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    '                        pObjstcPedido........objeto de pedido.
    '                        pCodUsuTransmissao...C�digo do usu�rio de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.
    
    Dim lobjstcEquipInsp       As stcA46v740B 'Objeto da classe de estrutura stcA46V726B (TAB_TRANS_INSP)
    Dim lobjclsEquipInsp       As clsA46v740A 'Objeto da classe de estrutura stcA46V726A (TAB_TRANS_INSP)
         
    mfRecTransInspecaoEquipamentos = False

    '1. Grava registro de transmiss�o inspe��o.
    Dim lIndice                 As Integer
    Dim lAux                    As Variant
    Dim lSelect                 As String
    Dim lrst                    As clsYasRecordSet
    Dim UltimoErro              As String            ' Descri��o do �ltimo erro gerado
    
    lIndice = 4     ' 022
                
    Set lobjstcEquipInsp = New stcA46v740B
           
    With lobjstcEquipInsp
        If mfRetCampo(4, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .nosso_numero = lAux
        mNossoTempor = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Observacao n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            Exit Function
        End If
        
        If mfRetCampo(2, "N�mero do Item", lAux, lIndice, 5, 0, pMensagem) = False Then
            Exit Function
        End If
        .Num_Item = lAux
       
        If mfRetCampo(2, "N�mero do equipamento", lAux, lIndice, 5, 0, pMensagem) = False Then
            Exit Function
        End If
        .NUM_EQUIP = lAux
   
        If mfRetCampo(4, "Descri��o/nome do equipamento", lAux, lIndice, 60, 0, pMensagem) = False Then
            Exit Function
        End If
        .NOME_EQUIP = lAux

        If mfRetCampo(2, "Tipo do equipamento", lAux, lIndice, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        .TIPO_EQUIP = lAux

        If mfRetCampo(2, "Classe do equipamento", lAux, lIndice, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        .CLASSE_EQUIP = lAux
        
        If mfRetCampo(2, "Ano de fabrica��o do equipamento", lAux, lIndice, 4, 0, pMensagem) = False Then
            Exit Function
        End If
        .ANO_FABRICACAO_EQUIP = lAux
        
        If mfRetCampo(4, "Marca do equipamento", lAux, lIndice, 40, 0, pMensagem) = False Then
            Exit Function
        End If
        .MARCA_EQUIP = lAux
        
        If mfRetCampo(4, "Modelo do equipamento", lAux, lIndice, 100, 0, pMensagem) = False Then
            Exit Function
        End If
        .MODELO_EQUIP = lAux

        If mfRetCampo(4, "Nro. s�rie do equipamento", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .NUM_SERIE_EQUIP = lAux

        If mfRetCampo(4, "chassi do equipamento", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .CHASSI_EQUIP = lAux

        If mfRetCampo(4, "Nome do contato para a inspe��o 1", lAux, lIndice, 60, 0, pMensagem) = False Then
            Exit Function
        End If
        .NOME_CONTATO1 = lAux

        If mfRetCampo(2, "Telefone do contato para a inspe��o 1", lAux, lIndice, 11, 0, pMensagem) = False Then
            Exit Function
        End If
        .TEL_CONTATO1 = lAux

        If mfRetCampo(4, "Email de contato 1", lAux, lIndice, 60, 0, pMensagem) = False Then
            Exit Function
        End If
        .EMAIL1 = lAux

        If mfRetCampo(4, "Nome do contato para a inspe��o 2", lAux, lIndice, 60, 0, pMensagem) = False Then
            Exit Function
        End If
        .NOME_CONTATO2 = lAux

        If mfRetCampo(2, "Telefone do contato para a inspe��o 2", lAux, lIndice, 11, 0, pMensagem) = False Then
            Exit Function
        End If
        .TEL_CONTATO2 = lAux

        If mfRetCampo(4, "Email de contato 2", lAux, lIndice, 60, 0, pMensagem) = False Then
            Exit Function
        End If
        .EMAIL2 = lAux

        If mfRetCampo(4, "Nome do contato para a inspe��o 3", lAux, lIndice, 60, 0, pMensagem) = False Then
            Exit Function
        End If
        .NOME_CONTATO3 = lAux

        If mfRetCampo(2, "Telefone do contato para a inspe��o 3", lAux, lIndice, 11, 0, pMensagem) = False Then
            Exit Function
        End If
        .TEL_CONTATO3 = lAux

        If mfRetCampo(4, "Email de contato 3", lAux, lIndice, 60, 0, pMensagem) = False Then
            Exit Function
        End If
        .EMAIL3 = lAux

        If mfRetCampo(4, "Observa��es", lAux, lIndice, 250, 0, pMensagem) = False Then
            Exit Function
        End If
        .Observacoes = lAux

        If mfRetCampo(4, "Logradouro para a inspe��o", lAux, lIndice, 50, 0, pMensagem) = False Then
            Exit Function
        End If
        .LOGRADOURO = lAux

        If mfRetCampo(4, "Complemento do logradouro para a inspe��o", lAux, lIndice, 20, 0, pMensagem) = False Then
            Exit Function
        End If
        .LOGR_COMPL = lAux

        If mfRetCampo(2, "N�mero do logradouro para a inspe��o", lAux, lIndice, 6, 0, pMensagem) = False Then
            Exit Function
        End If
        .LOGR_NUMERO = lAux

        If mfRetCampo(4, "Bairro do logradouro para a inspe��o", lAux, lIndice, 30, 0, pMensagem) = False Then
            Exit Function
        End If
        .BAIRRO = lAux

        If mfRetCampo(4, "Cidade do logradouro para a inspe��o", lAux, lIndice, 30, 0, pMensagem) = False Then
            Exit Function
        End If
        .CIDADE = lAux

        If mfRetCampo(4, "UF do logradouro para a inspe��o", lAux, lIndice, 2, 0, pMensagem) = False Then
            Exit Function
        End If
        .UF = lAux

        If mfRetCampo(2, "CEP do logradouro para a inspe��o", lAux, lIndice, 8, 0, pMensagem) = False Then
            Exit Function
        End If
        .CEP = lAux
        
        If mfRetCampo(2, "N�mero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            Exit Function
        End If
        .NUM_PI = lAux
        
        ' Valores fornecidos pelo registro de dados comuns (Tipo "001")
        .Cod_Ramo = pStcPedido.Cod_Ramo
        .Num_Ped = pStcPedido.Num_Ped
        .Cod_Org = pStcPedido.Cod_Org
        .Num_Apol_Ant = pStcPedido.Num_Apol_An
        .Nom_Segurado = pStcPedido.Nom_Segurado
        .Num_Tel = pStcPedido.Num_Fone
        .Num_Cgc_Cpf = pStcPedido.Num_Cgc_Cpf
        .Dat_Ini_Vig = pStcPedido.Dat_Ini_Vig
        .Dat_Ter_Vig = pStcPedido.Dat_Ter_Vig
        .Dat_Trans = pStcPedido.Dat_Emis
        .Cod_Iden_Pessoa = pStcPedido.Tip_Pessoa
        .Cod_Unid_Prod = pStcPedido.Cod_Unid_Prod
        .Cod_Prod = pStcPedido.Cod_Prod
        
        .Num_Insp = 0
        .Dat_Final = 0
        .Dat_Cancel = 0
        .Cod_Status_I = 1
        .Dat_Confirm = Format(M46V999.gForHoje, "YYYYMMDD")
    
    End With
            
    Set lobjclsEquipInsp = New clsA46v740A
    Set lobjclsEquipInsp.M46V999 = M46V999
    lobjstcEquipInsp.nosso_numero = mNossoTempor
    
    If pGravar = True Then
        Set lobjclsEquipInsp.ConexaoBD = bdBasDados
        If Not lobjclsEquipInsp.gfIncluir(lobjstcEquipInsp) Then
            pMensagem = lobjclsEquipInsp.UltimoErro
            Exit Function
        End If
    End If

    mfRecTransInspecaoEquipamentos = True
SAIDA:
    Set lobjclsEquipInsp = Nothing
    Set lobjstcEquipInsp = Nothing
    Exit Function

ControleErro:
    Call M46V999.gpGraLog(0, "A46V601A.mfRecTransInspecaoEquipamentos ERRO - " & UltimoErro & _
        vbCrLf & " SQL: " & lSelect)

End Function

Private Function mfRecPedAgregados(ByVal pNossoNumero As String, _
                                ByRef pMensagem As String, ByVal pPropTabela As String, _
                                ByVal pNomeTabImport As String, _
                                Optional pGravar As Boolean = True, _
                                Optional pStcPedAgregados As stcA46V702B) As Boolean
        'Fun��o: grava registro de coberturas para transmiss�o.

    'Par�metros de entrada...pNumArqTransmissao...N�mero do arquivo de transmiss�o.
    'Par�metros de sa�da.....pMensagem............Mensagem.

    Dim lobjPedAgregados        As clsA46v741A      'Objeto da classe clsA46V741A (TAB_PED_AGREGADOS)
    Dim lobjstcPedAgregados     As stcA46v741B      'Objeto da classe de estrutura stcA46V741B (TAB_PED_AGREGADOS)
    Dim lIndice                 As Integer
    Dim lAux                    As Variant

Err.Clear
mfRecPedAgregados = False
On Error GoTo TrataErro

    '1. Grava registro de coberturas.

    lIndice = 1
    If Mid(mRegistro, lIndice, 3) <> "025" Then
        pMensagem = "Registro n�o faz parte da cole��o. TAB_PED_AGREGADOS"
        GoTo SAIDA
    End If

    Set lobjstcPedAgregados = New stcA46v741B

    lIndice = lIndice + 3

    With lobjstcPedAgregados
        If mfRetCampo(6, "Nosso n�mero", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .nosso_numero = lAux
        If .nosso_numero <> pNossoNumero Then
            pMensagem = "Registro de Ped Agregados n�o pertence ao Registro Principal! - Nosso Numero " & pNossoNumero & " Reg:" & mRegistro
            GoTo SAIDA
        End If
        If mfRetCampo(2, "Numero do PI", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .NUM_PI = lAux
        If mfRetCampo(2, "Codigo do orgao", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Org = lAux
        If mfRetCampo(2, "Codigo do ramo", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo = lAux
        If mfRetCampo(2, "Numero do pedido", lAux, lIndice, 6, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ped = lAux
        If mfRetCampo(2, "Numero da apolice", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Apol = lAux
        If mfRetCampo(2, "Numero do endosso", lAux, lIndice, 6, 2, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Endo = lAux
        If mfRetCampo(2, "Tipo de Agredago", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tip_Agragado = lAux
        If mfRetCampo(4, "Nome do segurado", lAux, lIndice, 60, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Segurado = lAux
        If mfRetCampo(6, "Numero do CNPJ- CPF", lAux, lIndice, 14, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Cgc_Cpf = lAux
        If mfRetCampo(4, "Tipo de logradouro", lAux, lIndice, 3, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Tip_Logr = lAux
        If mfRetCampo(4, "Logradouro", lAux, lIndice, 50, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Logr = lAux
        If mfRetCampo(4, "Complemento do logradouro", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dsc_Compl_Logr = lAux
        If mfRetCampo(4, "N�mero do Logradouro", lAux, lIndice, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Logr = lAux
        If mfRetCampo(4, "Bairro", lAux, lIndice, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Bairro = lAux
        If mfRetCampo(4, "Cidade", lAux, lIndice, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Cid = lAux
        If mfRetCampo(4, "Sigla UF", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Sig_UF = lAux
        If mfRetCampo(2, "CEP", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_CEP = lAux
        If mfRetCampo(2, "Tipo_fone", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tipo_Fone = lAux
        If mfRetCampo(2, "Num_ddd_fone", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_DDD_Fone = lAux
        If mfRetCampo(2, "Num_fone", lAux, lIndice, 10, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Fone = lAux
        If mfRetCampo(2, "Num_ramal", lAux, lIndice, 4, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Ramal = lAux
        If mfRetCampo(2, "Tipo de pessoa", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Tip_Pessoa = lAux
        If mfRetCampo(2, "Codigo do ramo de atividade", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Ramo_Ativ = lAux
        If mfRetCampo(2, "Codigo do tipo de documento", lAux, lIndice, 2, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Cod_Doc = lAux
        If mfRetCampo(2, "Numero do documento", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Num_Doc = lAux
        If mfRetCampo(4, "Orgao expedidor", lAux, lIndice, 20, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Org_Exp = lAux
        If mfRetCampo(2, "Data da emissaoo do documento (yyyymmdd)", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Dat_Emis_Doc = lAux
        If mfRetCampo(2, "Data de Nascimento (yyyymmdd)", lAux, lIndice, 8, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .DAT_Nasc = lAux
        If mfRetCampo(4, "PROFISSAO", lAux, lIndice, 30, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Nom_Empresa = lAux
        If mfRetCampo(4, "Sexo do segurado: M -masculino , F - Feminino.", lAux, lIndice, 1, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Sexo = lAux
        If mfRetCampo(2, "Renda/Faturamento.", lAux, lIndice, 18, 0, pMensagem) = False Then
            GoTo SAIDA
        End If
        .Renda_Fatura = lAux
'        If mfRetCampo(4, "Renda/Faturamento.", lAux, lIndice, 18, 0, pMensagem) = False Then
'            GoTo SAIDA
'        End If
'        .DAT_GRAV = lAux
'        If mfRetCampo(2, "Usu�rio da Grava��o", lAux, lIndice, 10, 0, pMensagem) = False Then
'            GoTo SAIDA
'        End If
'        .Cod_User = lAux

    End With

    Set lobjPedAgregados = New clsA46v741A
    Set lobjPedAgregados.M46V999 = M46V999

    lobjstcPedAgregados.nosso_numero = mNossoTempor

    If pGravar = True Then
        Set lobjPedAgregados.ConexaoBD = bdBasDados
        Set lobjstcPedAgregados.ObjPai = pStcPedAgregados
        If Not lobjPedAgregados.mfIncluir(lobjstcPedAgregados, pPropTabela, pNomeTabImport, Empty) Then
            pMensagem = lobjPedAgregados.UltimoErro
            GoTo SAIDA
        End If
        If Not IsMissing(pStcPedAgregados) Then
            'pStcPedAgregados.ColecaoPedAgregados.Remove lobjstcPedAgregados.Class_ID
            Set pStcPedAgregados.objstcPedAgregados = lobjstcPedAgregados
        End If

    Else
    'Retorna no parametro de entrada
        If Not IsMissing(pStcPedAgregados) Then
            Set pStcPedAgregados.objstcPedAgregados = lobjstcPedAgregados
        Else
            pMensagem = "Gravar=Falso e pstcPedAgregados n�o informado."
            GoTo SAIDA
        End If
    End If
'================================
'******Fun��o realizada com sucesso
    mfRecPedAgregados = True
'================================

SAIDA:
    Err.Clear
    On Error GoTo 0
    GoSub FechaObj
    Exit Function

TrataErro:
    If M46V999.gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & ERROR
    End If
    pMensagem = "mfRecPedAgregados Erro: " & pMensagem
    GoTo SAIDA:

FechaObj:
    Set lobjPedAgregados = Nothing
    Set lobjstcPedAgregados = Nothing
    Return
End Function






