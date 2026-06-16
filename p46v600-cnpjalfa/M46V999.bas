Attribute VB_Name = "basM46V999"
'MODULO COM FUNÇÕES E PROCEDURES COMUNS DO SYAS.

'AO ALTERAR ESTE MÓDULO, COMPILAR E CRIAR EXECUTÁVEL DE TODOS OS MÓDULOS QUE O REFERENCIAM.

Option Explicit
Public gQuaErros            As Long
Public gNumSusep            As String
Public gCod_Corretor        As Long         'Cod do Corretor.
Public gCodProduto          As Integer      'Código do TARIFA
Public gCodSegurado         As Long         'Codigo do Segurado (Estipulante)
Public gCodProtocolo        As Long             'Código do protocolo.
Public gCodSequencia        As Long         'Código da sequência.
Public gQuaErrosNormas      As Byte
Public gQuaErrosDemais      As Byte
Public EhTeste             As Boolean
Public gFinalizarPrograma  As Boolean
Public gFormCancelado      As Boolean
Public gDesabilitarForms   As Boolean
Public gDesabFormsCotacao  As Boolean
Public gPendenteWorkflow   As Boolean
Public gAbrirWorkflow      As Boolean
Public gDesabFormsTransmi  As Boolean
'
'
Public gCod_Ramo        As e_CodRamo        'Código do ramo.
Public Const TarifaAntiga = 921
Public MAPI_SERVER_EXCHANGE As String
Public MAPI_Profile As String
Public MAPI_password As String
Public MAPI_CONTA_EXCHANGE As String


Public gProCusApolice       As Double           'Produto: custo de apólice.
Public gProCusEndosso       As Double           'Produto: custo de endosso.
Public gProPreMinParcela    As Double           'Produto: prêmio mínimo por parcela. (no Auto é prêmio mínimo do item)
Public gProPremioMinimo     As Double           'Produto: prêmio mínimo do item. (no Auto é prêmio mínimo da parcela)
Public gProQuaMaxParcelas   As Byte             'Produto: quantidade máxima de parcelas.
Public gProDatUltParcela    As Long             'Produto: data da última parcela (parcelamento)


Public gProPreMinParcelaTAB As Double
Public gProPremioMinimoTAB  As Double
Public gProQuaMaxParcelasTAB  As Byte
Public gNovaTransmissaoMulti  As Boolean

Public gTipUsuario      As e_Tipo_Usu   'Tipo do usuário.
Public gBaseLocal       As Boolean
Public gTipP00APOL      As String       'Tipo da base de dados: P00APOL.
Public gEndP00APOL      As String       'Endereço da base de dados: P00APOL.
Public gCodSeguradoYas  As String

Public gTipP00MPED      As String       'Tipo da base de dados: P00MPED.
Public gEndP00MPED      As String       'Endereço da base de dados: P00MPED.
Public gTipP0044700     As String       'Tipo da base de dados: P0044700.
Public gEndP0044700     As String       'Endereço da base de dados: P0044700.
Public gPrefixo         As String
Public gPerfilAcesso    As e_PerfilAcesso 'Perfil de Acesso no sistema para os usuários YASUDA
                                            '1 - Gestor (Digitação com acesso Master)
                                            '2 - Usuário Digitação (IST)
                                            '3 - Gestor (Cálculo)
                                            '4 - Usuário para Cálculo (versão corretor para usuários YASUDA)
                                            '5 - Usuário Digitação (YASUDA)
Public gTipP00MULTI     As String       'Tipo da base de dados: P00Multi.
Public gEndP00MULTI     As String       'Endereço da base de dados: P00MULTI.
Public gAppPath         As String       'App.Path considerando sempre uma \ no final.
Public gForHoje         As String       'Data do dia formatada.
Public gHojAno          As Integer      'Data do dia: ano.
Public gHojDia          As Integer      'Data do dia: dia.
Public gHojMes          As Integer      'Data do dia: mês.
Public gHoje            As Double       'Data do dia.

Public gNomArqLog       As String       'Nome do arquivo de log.
Public gNomUsuario      As String       'Nome do usuário.
Public gTipoHomologacao As String       'Verifica se é versão de homologação ou não
Public gNomUsuario6     As String       'Nome do usuário (6 posições).
Public gNomUsuMainFrame As String       'Nome do usuário no MainFrame

Public gVerDataMaq      As Byte         '0 - Não ver inconsistência de datas
                                        '1 - Ver inconsistências de data
Public gTrvSistema      As Boolean      'Travar Sistema

Public gPwdGed          As String       'Parâmetro PWS para base de dados: SISCAD.
Public gSerGed          As String       'Parâmetro SERVER para base de dados: SISCAD.
Public gTipGed          As String       'Tipo da base de dados: SISCAD.
Public gUidGed          As String       'Parâmetro UID para base de dados: SISCAD.
Public gEndGed          As String       'Endereço da base de dados: SISCAD.

Public gPwdDIVERSOS     As String       'Parâmetro PWS para base de dados: P00MPED.
Public gSerDIVERSOS     As String       'Parâmetro SERVER para base de dados: P00MPED.
Public gTipDIVERSOS     As String       'Tipo da base de dados: P00MPED.
Public gUidDIVERSOS     As String       'Parâmetro UID para base de dados: P00MPED.
Public gEndDIVERSOS     As String       'Endereço da base de dados: P00MPED.

Public mConexaoSQL      As ADODB.Connection 'Conexão com SQL Server.

Public gTipP0043000     As String       'Tipo da base de dados: P0043000.
Public gTipP0045200     As String       'Tipo da base de dados: P0045200.
Public gTipP0045300     As String       'Tipo da base de dados: P0045300.
Public gTipP0045400     As String       'Tipo da base de dados: P0045400.

Public gPwdSYASEMIS     As String       'Parâmetro PWS para base de dados: SYASEMIS.
Public gSerSYASEMIS     As String       'Parâmetro SERVER para base de dados: SYASEMIS.
Public gTipSYASEMIS     As String       'Tipo da base de dados: SYASEMIS.
Public gUidSYASEMIS     As String       'Parâmetro UID para base de dados: SYASEMIS.
Public gEndSYASEMIS     As String       'Endereço da base de dados: SYASEMIS.

Public gEndP0043000     As String       'Endereço da base de dados: P0043000.
Public gEndP0045200     As String       'Endereço da base de dados: P0045200.
Public gEndP0045300     As String       'Endereço da base de dados: P0045300.
Public gEndP0045400     As String       'Endereço da base de dados: P0045400.


Public gPwdEAPDB        As String       'Parâmetro PWS para base de dados: EAPDB.
Public gSerEAPDB        As String       'Parâmetro SERVER para base de dados: EAPDB.
Public gTipEAPDB        As String       'Tipo da base de dados: EAPDB.
Public gUidEAPDB        As String       'Parâmetro UID para base de dados: EAPDB.
Public gEndEAPDB        As String       'Endereço da base de dados: EAPDB.

Public gPwdSyasPortal   As String       'Parâmetro PWS para base de dados: SyasPortal
Public gSerSyasPortal   As String       'Parâmetro SERVER para base de dados: SyasPortal
Public gTipSyasPortal   As String       'Tipo da base de dados: SyasPortal
Public gUidSyasPortal   As String       'Parâmetro UID para base de dados: SyasPortal
Public gEndSyasPortal   As String       'Endereço da base de dados: SyasPortal

Public gAgendarCot      As Boolean   'Flag para carregar cotações
Public gUpdate          As String       'Campo genérico para montar comandos UPDATE.
Public gTipoAcesso      As e_Acesso

Public gEndSYASUpDt     As String       'Endereço da localização do s arquivos de atualização do SYAS
Public gCreate          As String       'Campo genérico para montar comandos Create.
Public gDelete          As String       'Campo genérico para montar comandos DELETE.
Public gEndRemessa      As String       'Endereço da pasta de remessa (arquivos a serem transmitidos).
Public gEndServidor     As String       'Endereço do servidor.

'[FRH] - GLB_sinal_concess é uma global que é true se as propr. [.Cod_Clasf] e [.Cod_Compl_Clasf] do Item (pobjstcitem-mfcalpreitem) são relativos a "022-13-Concessionária".
Public GLB_sinal_concess As Boolean
Public gErros()             As String           'Erros.
Public gAchouShopping  As Boolean      'Indica se achou o fator "RISCOS EM SHOPPING" ou não
Public gNovaPropostaLMI As Boolean
Public VerificarNormas              As Boolean
'Ficha 1000305
Public gTipCobertura            As Integer     'Tipo de cobertura dos bens (prédio/conteudo, prédio , conteúdo)
Public gEplano                  As Boolean
Public gPrograma        As String       'Nome do programa em execução.

Public gEndExeSiscad    As String       'Endereço da base de dados: SISCAD.
Public gEndExeP47V001   As String       'Endereço do executável P47V001.

Public gEndConfCot1     As String       'Endereço do formulário MULTCOT1.CXY
Public gEndImagCot1     As String       'Endereço do formulário MULTCOT1.WMF
Public gEndConfCot2     As String       'Endereço do formulário MULTCOT2.CXY
Public gEndImagCot2     As String       'Endereço do formulário MULTCOT2.WMF
Public gEndConfCot3     As String       'Endereço do formulário MULTCOT3.CXY
Public gEndImagCot3     As String       'Endereço do formulário MULTCOT3.WMF
Public gEndConfCotB     As String       'Endereço do formulário MULTCOTB.CXY
Public gEndImagCotB     As String       'Endereço do formulário MULTCOTB.WMF
Public gEndConfCotA     As String       'Endereço do formulário MULTCOTA.CXY
Public gEndImagCotA     As String       'Endereço do formulário MULTCOTA.WMF
Public gEndConfCotC     As String       'Endereço do formulário MULTCOTC.CXY
Public gEndImagCotC     As String       'Endereço do formulário MULTCOTC.WMF

Public gEndImagDeb      As String       'Endereço do formulário L28V009A.CXY
Public gEndConfDeb      As String       'Endereço do formulário L28V009A.WMF
Public gEndConfPro1     As String       'Endereço do formulário MULTPRO1.CXY
Public gEndImagPro1     As String       'Endereço do formulário MULTPRO1.WMF
Public gEndConfPro2     As String       'Endereço do formulário MULTPRO2.CXY
Public gEndImagPro2     As String       'Endereço do formulário MULTPRO2.WMF
Public gEndConfPro3     As String       'Endereço do formulário MULTPRO3.CXY
Public gEndImagPro3     As String       'Endereço do formulário MULTPRO3.WMF
Public gEndConfPro4     As String       'Endereço do formulário MULTPRO4.CXY
Public gEndImagPro4     As String       'Endereço do formulário MULTPRO4.WMF
Public gEndConfProB     As String       'Endereço do formulário MULTPROB.CXY
Public gEndImagProB     As String       'Endereço do formulário MULTPROB.WMF
Public gEndConfProC     As String       'Endereço do formulário MULTPROC.CXY
Public gEndImagProC     As String       'Endereço do formulário MULTPROC.WMF
Public gEndConfProD     As String       'Endereço do formulário MULTPROD.CXY
Public gEndImagProD     As String       'Endereço do formulário MULTPROD.WMF


'Formularios Antigos
Public gEndCot1_v0      As String       'Endereço do formulário MULTCOT1
Public gEndCot2_v0      As String       'Endereço do formulário MULTCOT2
Public gEndCot3_v0      As String       'Endereço do formulário MULTCOT3
Public gEndPro1_v0      As String       'Endereço do formulário MULTPRO1
Public gEndPro2_v0      As String       'Endereço do formulário MULTPRO2
Public gEndPro3_v0      As String       'Endereço do formulário MULTPRO3
Public gEndPro4_v0      As String       'Endereço do formulário MULTPRO4
Public gEndProB_v0      As String       'Endereço do formulário MULTPROB

'## smrosa
'Formulário Roubo
Public gEndConfRob1     As String       'Endereço do formulário ROUBOCOT1.CXY
Public gEndImagRob1     As String       'Endereço do formulário ROUBOCOT1.WMF
Public gEndConfRob2     As String       'Endereço do formulário ROUBOCOT2.CXY
Public gEndImagRob2     As String       'Endereço do formulário ROUBOCOT2.WMF
Public gEndConfRob3     As String       'Endereço do formulário ROUBOCOT3.CXY
Public gEndImagRob3     As String       'Endereço do formulário ROUBOCOT3.WMF
Public gEndConfRob4     As String       'Endereço do formulário ROUBOCOT4.CXY
Public gEndImagRob4     As String       'Endereço do formulário ROUBOCOT4.WMF
Public gEndConfRob5     As String       'Endereço do formulário ROUBOCOT5.CXY
Public gEndImagRob5     As String       'Endereço do formulário ROUBOCOT5.WMF

Public gEndImagCotLMI   As String       'Endereço do formulário COT_LMI.WMF
Public gEndImagRelLMI   As String       'Endereço do formulário REL_LMI.WMF
Public gEndImagESPLMI   As String       'Endereço do formulário COT_LMI.WMF
Public gEndImagESPRelLMI   As String       'Endereço do formulário REL_LMI.WMF

Public gEfetuouCalculo  As Boolean
Public gPodeEmitir      As Boolean      'True se o dígito de agência e conta estiverem OK
Public gPathArquivos    As String       'Local onde estão os arquivos de layout.
Public gPathImagePI     As String       'Local das Imagens do PI
Public gImprFrente      As String
Public gImprFrenteVerso As String

Public gUsuMulti        As Boolean      'Indica se o usuário é somente do sistema Multi ou é do Syas
                                        'também: TRUE = ambos sistemas e FALSE = apenas MULTISSEGUROS.
Public gDir_Ist         As String    'Path para gravar arquivos de conferencia da IST

Public Enum e_Acesso
    DIGITACAO = 1
    CALCULO = 2
End Enum

Public Enum e_Tipo_Usu
    e_Tipo_Usu_Yasuda = 1
    e_Tipo_Usu_Corretor = 2
    e_Tipo_Usu_CorretorYasuda = 3
End Enum

Public Enum e_SimNao
    e_SimNao_Sim = 1
    e_SimNao_Nao = 2
End Enum

Public Enum TipoDados
   tpoNumerico = 1
   TpoData = 2
   TpoHora = 3
   TpoValor = 4
   TpoString = 5
   TpoBoolean = 6
   TpoNumInteiro = 7
   TpoNumDecimal = 8
End Enum

Public Enum e_PerfilAcesso
    e_Perfil_DIG_Gestor = 1
    e_Perfil_DIG_IST = 2
    e_Perfil_CALC_Gestor = 3
    e_Perfil_CALC = 4
    e_Perfil_DIG_Yas = 5
    e_Perfil_Operacao = 6
    e_Perfil_Triagem = 7
End Enum



Public Enum e_TipoSistema
    e_TipoSistema_ND = 0
    e_TipoSistema_Serasa = 1
    e_TipoSistema_Nortix = 2
End Enum

Public Enum e_TipoLayout
    e_TipoLayout_ND = 0
    e_TipoLayout_Entrada = 1
    e_TipoLayout_Saida = 2
End Enum

Public Enum e_Delimitador
    e_Delimitador_ND = 0
    e_Delimitador_XML = 1
    e_Delimitador_Fixo = 2
    e_Delimitador_PontVirg = 3
End Enum


Public Enum e_TipoChave
    e_TipoChave_ND = 0
    e_TipoChave_Sistema = 1
    e_TipoChave_Cnpj = 2
    e_TipoChave_Placa = 3
    e_TipoChave_Chassi = 4
    e_TipoChave_Renavam = 5
End Enum

Public Enum ETransmDLL
    ETransmDLL_ONLine = 0
    ETransmDLL_OFFLine = 1
End Enum

Public gTipoTransmDLL   As ETransmDLL

Public gProibidoMsgBox      As Boolean   'Indica que um MsgBox Não pode ser exibido nunca (Exemplo Serviços e execuções em modo oculto)

Public Sub gpCenForm(ByVal pForm As Form)
    'Procedure: centraliza form.

    pForm.Left = (Screen.Width - pForm.Width) / 2
    pForm.Top = (Screen.Height - pForm.Height) / 2
End Sub
Public Function gfBuscaSusep(ByVal pCodCorretor As String, ByRef pSusep As String, _
                              ByRef pMensagem As String) As Boolean
    'Função: obtém nome do corretor.

    'Parâmetro de entrada...pCodCorretor...Código do corretor.
    'Parâmetros de saída....pNomCorretor...Nome do corretor.
    '                       pMensagem......Mensagem.

    Dim bdP0044700  As ADODB.Connection 'Arquivo - P0044700.
    Dim lSelect     As String           'Comando SELECT.
    Dim rsCORR      As ADODB.Recordset  'YTAB_CORR: tabela.

    gfBuscaSusep = False

    '1. Verifica preenchimento.
    If Not gfPreenchido(pCodCorretor) Then
        gfBuscaSusep = ""
        gfBuscaSusep = True
        Exit Function
    End If
    If Val(pCodCorretor) = 0 Then
        gfBuscaSusep = ""
        gfBuscaSusep = True
        Exit Function
    End If

    '2. Obtém nome do corretor.
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    lSelect = "SELECT NUM_SUSEP FROM YTAB_CORR WHERE COD_CORR = " & pCodCorretor
    If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
        Call gpFecha2(bdP0044700)
        Exit Function
    End If
    If rsCORR.EOF = True Then
        Call gpFecha3(rsCORR, bdP0044700)
        pMensagem = "FGE0044 - " & gPrefixo & "Corretor inválido (código de corretor não localizado " & _
                    "- " & pCodCorretor & ")."
        Exit Function
    End If
    pSusep = Trim$(rsCORR!NUM_SUSEP)
    Call gpFecha3(rsCORR, bdP0044700)

    gfBuscaSusep = True
End Function

Public Function Check(campo As ADODB.Field) As Variant
On Error GoTo TrataCheck
    If IsNull(campo) Then
        If campo.Type = 5 Or campo.Type = 200 Or campo.Type = 201 Or campo.Type = 202 Or campo.Type = 203 Or campo.Type = 129 Then
            Check = ""
        ElseIf campo.Type = 7 Then
            Check = ""
        ElseIf campo.Type = 2 Then
            Check = 0
        ElseIf campo.Type = 135 Then
        ElseIf campo.Type = 3 Or campo.Type = 131 Then
            Check = 0
        Else
            MsgBox "Verificar tipo Campo: " & campo.Name & " Tipo: " & campo.Type
            Check = 0
        End If
    Else
        Check = campo.Value
    End If
    Exit Function
TrataCheck:
    If Err = 3021 Then Resume Next
    MsgBox "Problemas com o campo " & campo.Name
    Exit Function
End Function

Public Function gfVerEtiqueta(ByVal pCorretor As Long, ByVal pCodReg As Integer, ByVal pCodPessoa As Integer, ByVal pTipEnv As String, ByRef pMensagem As String, Optional ByVal tipoEnvio As String) As Boolean

Dim bdP00MULTI          As ADODB.Connection
Dim rsTbsCorretor       As ADODB.Recordset  'TBS_CORRETOR: tabela.
Dim lSelect             As String

    If gTipUsuario = e_Tipo_Usu_Corretor Or _
        gTipUsuario = e_Tipo_Usu_CorretorYasuda Or _
       gBaseLocal = True Or _
       gPerfilAcesso = e_Perfil_CALC Or _
       gPerfilAcesso = e_Perfil_CALC_Gestor Then
        gfVerEtiqueta = False
        GoTo SAIDA
    End If
    
    If (pCorretor = 15701 Or pCorretor = 8101) And _
        Not (pCodReg = 106 Or pCodReg = 107) Then
        If pCodPessoa = 1 And pTipEnv <> "C" Then
            If pTipEnv = "X" Then pMensagem = "O": GoTo SAIDA '[FRH]Ficha:0801295 teste
            pMensagem = "Envio do documento deve ser para o 'Corretor'."
            gfVerEtiqueta = True
        ElseIf pCodPessoa = 2 And pTipEnv <> "S" Then
            If pTipEnv = "X" Then pMensagem = "E": GoTo SAIDA '[FRH]Ficha:0801295
            pMensagem = "Envio do documento deve ser para o 'Segurado'."
            gfVerEtiqueta = True
        Else
            gfVerEtiqueta = False
        End If
        GoTo SAIDA
    End If
    
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdP00MULTI, pMensagem) = False Then
        GoTo SAIDA
    End If

    lSelect = "SELECT * FROM ged.dbo.TBS_CORRETOR  WHERE Cod_Corretor = " & pCorretor
              
    If gfObtRegistro(bdP00MULTI, lSelect, rsTbsCorretor, pMensagem) = False Then
        Call gpFecha2(bdP00MULTI)
        GoTo SAIDA
    End If
    
    If rsTbsCorretor.EOF Then
        pMensagem = "Corretor não cadastrado na TBS_CORRETOR!!"
        gfVerEtiqueta = True
        Call gpFecha2(bdP00MULTI)
        GoTo SAIDA
    End If

    If rsTbsCorretor("Cod_Etiqueta") = "5" Or _
        rsTbsCorretor("Cod_Etiqueta") = "6" Or _
        rsTbsCorretor("Cod_Etiqueta") = "7" Or _
        rsTbsCorretor("Cod_Etiqueta") = "9" Or _
        rsTbsCorretor("Cod_Etiqueta") = "A" Then

        gfVerEtiqueta = False

        If pTipEnv = "S" Then
            pMensagem = "Envio do documento deve ser para o 'Corretor'."
            gfVerEtiqueta = True
            GoTo SAIDA
        ElseIf pTipEnv = "X" Then
        
            '[FRH]Ficha:0801295 [Conforme cadastro - "O"]
            pMensagem = "O"
            
            Call gpFecha2(bdP00MULTI)
            GoTo SAIDA
        End If
    Else
        If pTipEnv = "X" Then
        
            '[FRH]Ficha:0801295 [Conforme cadastro - "E"]
            pMensagem = "E"
            
            Call gpFecha2(bdP00MULTI)
            GoTo SAIDA
        End If
    End If
    
    Call gpFecha2(bdP00MULTI)

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set bdP00MULTI = Nothing
    Set rsTbsCorretor = Nothing
    Return
End Function

Public Function mfCalCodUnicidade_V1(ByRef pobjstcPedido As stcA46V702B, ByRef pMensagem As String)
    'Função: Calcula código de verificação da unicidade.

    'Parâmetro de saída...pMensagem.........Mensagem.

    Dim lAuxiliar           As String           'Auxiliar.
    Dim lCampo1()           As String           'Campos utilizados para dividir a seqüência em campos de
                                                '10 posições cada.
    Dim lCampo2()           As Double           'Campos com totais das somas das posições de lCampo1.
    Dim lCodVerificacao     As String           'Código de verificação.
    Dim li                  As Double           'Utilizado com For...Next.
    Dim lPosInicial         As Double           'Posição inicial.
    Dim lResto              As Double           'Resto da divisão.
    Dim lResultado          As Double           'Resultado da divisão.
    Dim lSequencia          As String           'Seqüência para cálculo do código de unicidade.
    Dim lSequencia2         As String           'Seqüência para cálculo do código de unicidade.
    Dim lTamanho            As Double           'Tamanho.
    Dim lTotal              As Double           'Total.
    Dim lRamAtividade       As String           'Pessoa jurídica: ramo de atividade.
    Dim lTipDocumento       As String           'Pessoa física: tipo do documento.
    Dim lNumDocumento       As String           'Pessoa física: número do documento.
    Dim lOrgEmiDocumento    As String           'Pessoa física: órgão emissor do documento.
    Dim lDatEmiDocumento    As String           'Pessoa física: data de emissão do documento.
    
    Dim lobjstcCoss         As stcA46V706B      'Objeto da classe de estrutura stcA46V706B (TAB_PED_COSS).
    Dim lobjstcJuros        As stcA46V711B      'Objeto da classe de estrutura stcA46V711B (TAB_PED_JUROS).
    Dim lObjstcVenc         As stcA46V713B      'Objeto da classe de estrutura stcA46V713B (TAB_PED_VENC).
    Dim lobjstcPedLoc       As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcCobert       As stcA46V704B
    
    mfCalCodUnicidade_V1 = False
    
    'Crio a sequencia
    lSequencia = ""
    lSequencia2 = ""
'-*-*-*-*-*
    Select Case pobjstcPedido.Tip_Pessoa
        Case e_TipPessoa_Juridica       'Pessoa jurídica.
            'a) Ramo de atividade - 2 posições numéricas.
            If IsNumeric(pobjstcPedido.Cod_Ramo_Ativ) Then
                lRamAtividade = Format(pobjstcPedido.Cod_Ramo_Ativ, "00")
            Else
                lRamAtividade = "00"
            End If
        Case e_TipPessoa_Fisica       'Pessoa física.
            'a) Registro de identificação do segurado.
            lTipDocumento = "00"
            lNumDocumento = String(20, " ")
            lOrgEmiDocumento = String(20, " ")
            lDatEmiDocumento = String(8, "0")

            'b) Tipo do documento - 2 posições numéricas.
            If IsNumeric(pobjstcPedido.Cod_Doc) Then
                lTipDocumento = pobjstcPedido.Cod_Doc
            End If
            'c) Número do documento - 20 posições.
            lNumDocumento = pobjstcPedido.Num_Doc
            lNumDocumento = lNumDocumento & String(20 - Len(Left(pobjstcPedido.Num_Doc, 20)), " ")
            
            'd) Órgão emissor do documento - 20 posições.
            lOrgEmiDocumento = pobjstcPedido.Org_Exp
            lOrgEmiDocumento = lOrgEmiDocumento & String(20 - Len(Left(pobjstcPedido.Org_Exp, 20)), " ")

            'e) Data de emissão do documento - 8 posições numéricas.
            If IsNumeric(pobjstcPedido.Dat_Emis_Doc) Then
                lDatEmiDocumento = pobjstcPedido.Dat_Emis_Doc
            End If
    End Select
    
    '3. Inicia seqüência para cálculo do código de unicidade com dados comuns.
    '===001) Nosso número......................................................(20 posições - total = 20).
    lSequencia = pobjstcPedido.Nosso_Numero
    '===002) Referência.........................................................(6 posições - total = 26).
    'lSequencia = lSequencia & Format$(pobjstcPedido.Num_Referencia, "000000")
    '===003) Código do corretor.................................................(7 posições - total = 33).
    lSequencia = lSequencia & Format$(pobjstcPedido.objstcCtrlEmis.Cod_Corr, "0000000")
    '===004) Código do produto (tarifa).........................................(4 posições - total = 37).
    lSequencia = lSequencia & Format$(pobjstcPedido.Cod_Produto, "0000")
    '===005) Código da unidade de negócios......................................(5 posições - total = 42).
    lSequencia = lSequencia & Format$(pobjstcPedido.Cod_Unid_Prod, "00000")
    '===006) Código do produtor.................................................(5 posições - total = 47).
    lSequencia = lSequencia & Format$(pobjstcPedido.Cod_Prod, "00000")
    '===007) Nome do proponente (segurado)....................................(60 posições - total = 107).
    lAuxiliar = UCase$(pobjstcPedido.Nom_Segurado)
    If Len(lAuxiliar) > 60 Then
        lAuxiliar = Left$(lAuxiliar, 60)
    Else
        lAuxiliar = lAuxiliar & String(60 - Len(lAuxiliar), " ")
    End If
    lSequencia = lSequencia & lAuxiliar
    
    '===008) Código do segurado................................................(7 posições - total = 114).
    lAuxiliar = String(7, "0")
    If gfPreenchido(pobjstcPedido.Cod_Segurado) Then
        If IsNumeric(pobjstcPedido.Cod_Segurado) Then
            lAuxiliar = Format$(Val(pobjstcPedido.Cod_Segurado), "0000000")
        End If
    End If
    lSequencia = lSequencia & lAuxiliar
    '===009) Tipo de pessoa.....................................................(1 posição - total = 115).
    '===010) Identificação....................................................(61 posições - total = 176).
    Select Case pobjstcPedido.Tip_Pessoa
        Case e_TipPessoa_Juridica
            'a) Tipo de pessoa.
            lAuxiliar = "J"
            'b) CNPJ - 14 posições numéricas.
            lAuxiliar = lAuxiliar & Format$(pobjstcPedido.Num_Cgc_Cpf, "00000000000000")
            'c) Ramo de atividade - 2 posições numéricas.
            lAuxiliar = lAuxiliar & lRamAtividade
            'd) Completa com espaços.
            lAuxiliar = lAuxiliar & String(45, " ")
        Case e_TipPessoa_Fisica
            'a) Tipo de pessoa.
            lAuxiliar = "F"
            'b) CPF - 11 posições numéricas.
            lAuxiliar = lAuxiliar & Format$(pobjstcPedido.Num_Cgc_Cpf, "00000000000")
            'c) Tipo do documento - 2 posições numéricas.
            lAuxiliar = lAuxiliar & lTipDocumento
            'd) Número do documento - 20 posições.
            lAuxiliar = lAuxiliar & lNumDocumento
            'e) Órgão emissor do documento - 20 posições.
            lAuxiliar = lAuxiliar & lOrgEmiDocumento
            'f) Data de emissão do documento - 8 posições numéricas (aaaammdd).
            lAuxiliar = lAuxiliar & lDatEmiDocumento
    End Select
    lSequencia = lSequencia & lAuxiliar
    
    '===011) Endereço do segurado - logradouro................................(73 posições - total = 249).
    'O Endereço não faz mais parte do código de unicidade (solicitação Shiba)
    
'**************************************************************************************
'*    lAuxiliar = ""
'*    If gfPreenchido(pObjstcPedido.Cod_Tip_Logr) Then
'*        lAuxiliar = Trim$(pObjstcPedido.Cod_Tip_Logr) & " "
'*    End If
'*    If gfPreenchido(pObjstcPedido.Nom_Logr) Then
'*        lAuxiliar = lAuxiliar & Trim$(pObjstcPedido.Nom_Logr) & " "
'*    End If
'*    If gfPreenchido(pObjstcPedido.Dsc_Compl_Logr) Then
'*        lAuxiliar = lAuxiliar & Trim$(pObjstcPedido.Dsc_Compl_Logr)
'*    End If
'*    If Len(lAuxiliar) > 73 Then
'*        lAuxiliar = Left$(lAuxiliar, 73)
'*    Else
'*        lAuxiliar = lAuxiliar & String(73 - Len(lAuxiliar), " ")
'*    End If
    
'*    lSequencia = lSequencia & lAuxiliar
    '===012) Endereço do segurado - bairro....................................(30 posições - total = 279).
'*    If gfPreenchido(pObjstcPedido.Nom_Bairro) Then
'*        lAuxiliar = pObjstcPedido.Nom_Bairro
'*        If Len(lAuxiliar) < 30 Then
'*            lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
'*        End If
'*    Else
'*        lAuxiliar = String(30, " ")
'*    End If
'*    lSequencia = lSequencia & lAuxiliar
    '===013) Endereço do segurado - cidade....................................(30 posições - total = 309).
'*    If gfPreenchido(pObjstcPedido.Nom_Cid) Then
'*        lAuxiliar = pObjstcPedido.Nom_Cid
'*        If Len(lAuxiliar) < 30 Then
'*            lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
'*        End If
'*    Else
'*        lAuxiliar = String(30, " ")
'*    End If
'*    lSequencia = lSequencia & lAuxiliar
    '===014) Endereço do segurado - unidade da federação.......................(2 posições - total = 311).
'*    If gfPreenchido(pObjstcPedido.Sig_UF) Then
'*        lAuxiliar = pObjstcPedido.Sig_UF
'*        If Len(lAuxiliar) < 2 Then
'*            lAuxiliar = lAuxiliar & String(2 - Len(lAuxiliar), " ")
'*        End If
'*    Else
'*        lAuxiliar = String(2, " ")
'*    End If
'*    lSequencia = lSequencia & lAuxiliar
    '===015) Endereço do segurado - CEP........................................(8 posições - total = 319).
'*    If gfPreenchido(pObjstcPedido.Num_Cep) Then
'*        lAuxiliar = pObjstcPedido.Num_Cep
'*        If Len(lAuxiliar) < 8 Then
'*            lAuxiliar = String(8 - Len(lAuxiliar), "0") & lAuxiliar
'*        End If
'*    Else
'*        lAuxiliar = String(8, "0")
'*    End If
'*    lSequencia = lSequencia & lAuxiliar
    '===016) Endereço do segurado - telefones.................................(30 posições - total = 349).
'*    If gfPreenchido(pObjstcPedido.Num_Fone) Then
'*        lAuxiliar = pObjstcPedido.Num_Fone
'*        If Len(lAuxiliar) < 30 Then
'*            lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
'*        End If
'*    Else
'*        lAuxiliar = String(30, " ")
'*    End If
'*    lSequencia = lSequencia & lAuxiliar
    '===017) Endereço de cobrança - logradouro................................(73 posições - total = 422).
'*    lAuxiliar = ""
'*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Cod_Tip_Logr) Then
'*        lAuxiliar = Trim$(pObjstcPedido.objstcPedEndereco.Cod_Tip_Logr) & " "
'*    End If
'*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Nom_Logr) Then
'*        lAuxiliar = lAuxiliar & Trim$(pObjstcPedido.objstcPedEndereco.Nom_Logr) & " "
'*    End If
'*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Dsc_Compl_Logr) Then
'*        lAuxiliar = lAuxiliar & Trim$(pObjstcPedido.objstcPedEndereco.Dsc_Compl_Logr)
'*    End If
'*    If Len(lAuxiliar) > 73 Then
'*        lAuxiliar = Left$(lAuxiliar, 73)
'*    Else
'*        lAuxiliar = lAuxiliar & String(73 - Len(lAuxiliar), " ")
'*    End If
'*    lSequencia = lSequencia & lAuxiliar
'*    '===018) Endereço de cobrança - bairro....................................(30 posições - total = 452).
'*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Nom_Bairro) Then
'*        lAuxiliar = pObjstcPedido.objstcPedEndereco.Nom_Bairro
'*        If Len(lAuxiliar) < 30 Then
'*            lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
'*        End If
'*    Else
'*        lAuxiliar = String(30, " ")
'*    End If
'*    lSequencia = lSequencia & lAuxiliar
    '===019) Endereço de cobrança - cidade....................................(30 posições - total = 482).
'*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Nom_Cid) Then
'*        lAuxiliar = pObjstcPedido.objstcPedEndereco.Nom_Cid
'*        If Len(lAuxiliar) < 30 Then
'*            lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
'*        End If
'*    Else
'*        lAuxiliar = String(30, " ")
'*    End If
'*    lSequencia = lSequencia & lAuxiliar
    '===020) Endereço de cobrança - unidade da federação.......................(2 posições - total = 484).
'*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Sig_UF) Then
'*        lAuxiliar = pObjstcPedido.objstcPedEndereco.Sig_UF
'*        If Len(lAuxiliar) < 2 Then
'*            lAuxiliar = lAuxiliar & String(2 - Len(lAuxiliar), " ")
'*        End If
'*    Else
'*        lAuxiliar = String(2, " ")
'*    End If
'*    lSequencia = lSequencia & lAuxiliar
    '===021) Endereço de cobrança - CEP........................................(8 posições - total = 492).
'*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Num_Cep) Then
'*        lAuxiliar = pObjstcPedido.objstcPedEndereco.Num_Cep
'*        If Len(lAuxiliar) < 8 Then
'*            lAuxiliar = String(8 - Len(lAuxiliar), "0") & lAuxiliar
'*        End If
'*    Else
'*        lAuxiliar = String(8, "0")
'*    End If
'*    lSequencia = lSequencia & lAuxiliar
    '===022) Endereço de cobrança - telefones.................................(30 posições - total = 522).
'*    If gfPreenchido(pObjstcPedido.objstcPedEndereco.Num_Fone) Then
'*        lAuxiliar = pObjstcPedido.objstcPedEndereco.Num_Fone
'*        If Len(lAuxiliar) < 30 Then
'*            lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
'*        End If
'*    Else
'*        lAuxiliar = String(30, " ")
'*    End If
'*    lSequencia = lSequencia & lAuxiliar
'***************************************************************************

    '===023) Início de vigência................................................(8 posições - total = 530).
    lSequencia = lSequencia & Format$(pobjstcPedido.Dat_Ini_Vig, "00000000")
    '===024) Término de vigência...............................................(8 posições - total = 538).
    lSequencia = lSequencia & Format$(pobjstcPedido.Dat_Ter_Vig, "00000000")
    '===025) Corretagem.......................................................(33 posições - total = 571).
    lSequencia = lSequencia & Format$(pobjstcPedido.objstcCtrlEmis.Cod_Corr, "0000000")
    lAuxiliar = Format$(pobjstcPedido.objstcCtrlEmis.Per_Corr, "00.00")
    lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
    lSequencia = lSequencia & Format$(pobjstcPedido.objstcCtrlEmis.Cod_Corr_2, "0000000")
    lAuxiliar = Format$(pobjstcPedido.objstcCtrlEmis.Per_Corr_2, "00.00")
    lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
    lSequencia = lSequencia & Format$(pobjstcPedido.objstcCtrlEmis.Cod_Corr_3, "0000000")
    lAuxiliar = Format$(pobjstcPedido.objstcCtrlEmis.Per_Corr_3, "00.00")
    lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
    
    '===026) Cosseguros.......................................................(50 posições - total = 621).
    For Each lobjstcCoss In pobjstcPedido.ColecaoPedCoss
        lSequencia = lSequencia & Format$(lobjstcCoss.Cod_Coss, "000000")
        lAuxiliar = Format$(lobjstcCoss.Per_Coss, "00.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
    Next lobjstcCoss
    
    '===027) Juros ao mês......................................................(4 posições - total = 625).
    lAuxiliar = "00,00"
    For Each lobjstcJuros In pobjstcPedido.ColecaoPedJuros
        If pobjstcPedido.Qtd_Parc = lobjstcJuros.Num_Parcela Then
            If pobjstcPedido.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
                lAuxiliar = Format$(lobjstcJuros.Val_Ant, "00.00")
            Else
                lAuxiliar = Format$(lobjstcJuros.Val_Pos, "00.00")
            End If
            Exit For
        End If
    Next
    lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
    
    '===028) Custo da apólice.................................................(10 posições - total = 635).
    lAuxiliar = Format$(pobjstcPedido.Val_Cust_Emis, "00000000.00")
    lSequencia = lSequencia & Left$(lAuxiliar, 8) & Right$(lAuxiliar, 2)
    '===029) Valor do IOF.....................................................(10 posições - total = 645).
    lAuxiliar = Format$(pobjstcPedido.Val_IOF, "00000000.00")
    lSequencia = lSequencia & Left$(lAuxiliar, 8) & Right$(lAuxiliar, 2)
    '===030) Número de parcelas................................................(2 posições - total = 647).
    lAuxiliar = Format$(pobjstcPedido.Qtd_Parc, "00")
    lSequencia = lSequencia & lAuxiliar & lAuxiliar
    '===031) Indicador de primeira parcela paga antecipada......................(1 posição - total = 648).
    lAuxiliar = Format$(pobjstcPedido.objstcCtrlEmis.Cod_Iden_Pag_Antec, "0")
    lSequencia = lSequencia & lAuxiliar
    
    '===032) Valor da primeira parcela........................................(10 posições - total = 658).
    'If pobjstcPedido.Val_Parcela1_VM = 0 Then
    '    lAuxiliar = Format$(pobjstcPedido.Val_Parcela1_VD, "00000000.00")
    'Else
    '    lAuxiliar = Format$(pobjstcPedido.Val_Parcela1_VM, "00000000.00")
    'End If
    'lSequencia = lSequencia & Left$(lAuxiliar, 8) & Right$(lAuxiliar, 2)
    
    '===033) Valor das demais parcelas........................................(10 posições - total = 668).
    'If pobjstcPedido.Val_Parcelas_VM = 0 Then
    '    lAuxiliar = Format$(pobjstcPedido.Val_Parcelas_VD, "00000000.00")
    'Else
    '    lAuxiliar = Format$(pobjstcPedido.Val_Parcelas_VM, "00000000.00")
    'End If
    'lSequencia = lSequencia & Left$(lAuxiliar, 8) & Right$(lAuxiliar, 2)
    
    '===034) Pagamento em bloqueto - número do bloqueto.......................(10 posições - total = 678).
    lSequencia = lSequencia & Format$(pobjstcPedido.objstcCtrlEmis.Num_Bloq, "0000000000")
    '===035) Pagamento através de carnê.......................................(11 posições - total = 689).
    'a) Código do banco.
    lSequencia = lSequencia & Format$(pobjstcPedido.Cod_Bco_Cobr, "000")
    'b) Vencimento da primeira parcela.
    'lSequencia = lSequencia & Format$(pobjstcPedido.Ven_Parcela1_VM, "00000000")
    '===036) Pagamento através de débito em conta.............................(18 posições - total = 707).
    
    'a) Código da agência.
    lSequencia = lSequencia & Format$(pobjstcPedido.Cod_Agen_Bas_Cor, "0000")
    'b) Código da conta corrente.
    lSequencia = lSequencia & Format$(pobjstcPedido.Num_Cta_Bas_Cor, "000000000")
    ''c) Dia para débito.
    'lSequencia = lSequencia & Format$(pobjstcPedido.Deb_Dia, "00")
    For Each lObjstcVenc In pobjstcPedido.ColecaoPedVenc
        lSequencia = lSequencia & Format$(lObjstcVenc.Num_Parc, "00")
        lSequencia = lSequencia & Format$(lObjstcVenc.DAT_VENC, "00000000")
    Next lObjstcVenc
    
    '===037) Juros de mora.....................................................(4 posições - total = 711).
    lSequencia = lSequencia & "0450"

    'a) Dados dos itens.
    If pobjstcPedido.ColecaoPedLoc.Count = 0 Then
        pMensagem = "FTE0056 - " & gPrefixo & _
                    "Não localizado registro de item (cálculo do código de unicidade)."
        GoTo SAIDA
    End If
    
    '4. Continua seqüência para cálculo do código de unicidade com dados dos itens.
    For Each lobjstcPedLoc In pobjstcPedido.ColecaoPedLoc
        'a) Obtém valores dos itens.
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Num_Item, "00000")
        lAuxiliar = ""
        If gfPreenchido(lobjstcPedLoc.Cod_Tip_Logr_Risco) Then
            lAuxiliar = Trim$(lobjstcPedLoc.Cod_Tip_Logr_Risco) & " "
        End If
        '== Risco - Logradouro.
        If gfPreenchido(lobjstcPedLoc.Nom_Logr_Risco) Then
            lAuxiliar = lAuxiliar & Trim$(lobjstcPedLoc.Nom_Logr_Risco) & " "
        End If
        '== Risco - Complemento.
        If gfPreenchido(lobjstcPedLoc.Dsc_Compl_Risco) Then
            lAuxiliar = lAuxiliar & Trim$(lobjstcPedLoc.Dsc_Compl_Risco) & " "
        End If
        If Val(lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco) <> 0 Then
            lAuxiliar = lAuxiliar & Trim$(lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco) & " "
        End If
        If Len(lAuxiliar) > 73 Then
            lAuxiliar = Left$(lAuxiliar, 73)
        Else
            lAuxiliar = lAuxiliar & String(73 - Len(lAuxiliar), " ")
        End If
        lSequencia = lSequencia & lAuxiliar
        '== Risco - Bairro.
        If gfPreenchido(lobjstcPedLoc.Nom_Bairro_Risco) Then
            lAuxiliar = lobjstcPedLoc.Nom_Bairro_Risco
            If Len(lAuxiliar) < 30 Then
                lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
            End If
        Else
            lAuxiliar = String(30, " ")
        End If
        lSequencia = lSequencia & lAuxiliar
        '== Risco - Cidade.
        If gfPreenchido(lobjstcPedLoc.Nom_Cid_Risco) Then
            lAuxiliar = lobjstcPedLoc.Nom_Cid_Risco
            If Len(lAuxiliar) < 30 Then
                lAuxiliar = lAuxiliar & String(30 - Len(lAuxiliar), " ")
            End If
        Else
            lAuxiliar = String(30, " ")
        End If
        lSequencia = lSequencia & lAuxiliar
        '== Risco - UF.
        If gfPreenchido(lobjstcPedLoc.Sig_UF_Risco) Then
            lAuxiliar = lobjstcPedLoc.Sig_UF_Risco
            If Len(lAuxiliar) < 2 Then
                lAuxiliar = lAuxiliar & String(2 - Len(lAuxiliar), " ")
            End If
        Else
            lAuxiliar = String(2, " ")
        End If
        lSequencia = lSequencia & lAuxiliar
        '== Risco - CEP.
        If gfPreenchido(lobjstcPedLoc.Num_Cep_Risco) Then
            lAuxiliar = lobjstcPedLoc.Num_Cep_Risco
            If Len(lAuxiliar) < 8 Then
                lAuxiliar = String(8 - Len(lAuxiliar), "0") & lAuxiliar
            End If
        Else
            lAuxiliar = String(8, "0")
        End If
        lSequencia = lSequencia & lAuxiliar
        '== Desconto progressivo.
        lAuxiliar = Format$(lobjstcPedLoc.Per_Desc_Progres, "00.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
        '== IS Básica
        lAuxiliar = Format$(lobjstcPedLoc.Per_IS_Vr, "000.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 3) & Right$(lAuxiliar, 2)
        '= Coeficiente IS/VR
        lAuxiliar = Format$(lobjstcPedLoc.Val_Coef_IS_Vr, "000.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 3) & Right$(lAuxiliar, 2)
        '== Valor de risco
        lAuxiliar = Format$(lobjstcPedLoc.Val_Vr, "000000000000.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
        '== Classificação
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Tab_Clasf, "0000")
        lSequencia = lSequencia & lobjstcPedLoc.Cod_Clasf
        lSequencia = lSequencia & lobjstcPedLoc.Cod_Compl_Clasf
        '== Construção
        lSequencia = lSequencia & lobjstcPedLoc.Cod_Constr
        lAuxiliar = Format$(lobjstcPedLoc.Per_Bonus, "00.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
        '==Quantidade de empregados.
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Qtd_Empreg, "0000")
        
        lAuxiliar = Format$(lobjstcPedLoc.Per_Tot_Segur, "00.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
        '== Percentual total de proteção.
        lAuxiliar = Format$(lobjstcPedLoc.Per_Tot_Protec, "00.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
        '== Desconto de aparelho.
        lAuxiliar = Format$(lobjstcPedLoc.Per_Desc_Apar, "00.00")
        lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
        '== produto da tabela de região.
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Cod_Produto, "0000")
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Cod_Regiao, "0000")
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Cod_Iden_Tarif, "0000")
        '== Tipo de emissão
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Tip_Emissao, "00")
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Num_Apol_Anterior, "0000000000")
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Num_Item_Anterior, "00000")
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Ind_Sinistro, "0")
        lSequencia2 = lSequencia2 & lSequencia
        lSequencia = ""
        For Each lobjstcCobert In lobjstcPedLoc.ColecaoPedCobert
            If lobjstcCobert.Num_Item = lobjstcPedLoc.Num_Item Then
                lSequencia = lSequencia & Format$(lobjstcCobert.Cod_Cobert, "0000")
                lAuxiliar = Format$(lobjstcCobert.Per_Desc_Coml, "00.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Per_Desc_Exp, "00.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Per_Franq, "00.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Per_Taxa, "00.00000")
                lSequencia = lSequencia & Left$(lAuxiliar, 2) & Right$(lAuxiliar, 5)
                lAuxiliar = Format$(lobjstcCobert.Val_Franq_Max, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lSequencia = lSequencia & Format$(lobjstcCobert.Tab_Cobert, "0000")
                lAuxiliar = Format$(lobjstcCobert.Val_Franq_Min, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Val_IS, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.VAL_PR, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Val_Pr_Anu, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Val_Pr_Resg, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Val_Pr_Resg_Anu, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Val_Pr_Tarif, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lAuxiliar = Format$(lobjstcCobert.Val_Pr_Tarif_Anu, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
             'TAB_PED_COBERT_COMPL
                lSequencia = lSequencia & Format$(lobjstcCobert.Num_Meses_PI, "00")
                lAuxiliar = Format$(lobjstcCobert.Val_Risco_Cob, "000000000000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 12) & Right$(lAuxiliar, 2)
                lSequencia = lSequencia & Format$(lobjstcCobert.Cod_Classe, "0000")
                lAuxiliar = Format$(lobjstcCobert.Val_Fator_Franq, "000.00")
                lSequencia = lSequencia & Left$(lAuxiliar, 3) & Right$(lAuxiliar, 2)
            End If
        Next
        
        'Ficha - 1000556 - Inclusão de dados de Beneficiarios (e demais) na transmissão
        Dim lobjstcTextoEspec   As stcA46V726B      'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).
        
        For Each lobjstcTextoEspec In pobjstcPedido.ColecaoTextoEspec
            If gfPreenchido(lobjstcTextoEspec.Espec_lin1) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin1)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin2) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin2)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin3) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin3)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin4) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin4)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin5) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin5)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin6) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin6)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin7) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin7)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin8) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin8)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin9) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin9)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin10) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin10)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin11) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin11)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin12) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin12)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin13) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin13)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin14) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin14)
            If gfPreenchido(lobjstcTextoEspec.Espec_lin15) Then lSequencia = lSequencia & Trim$(lobjstcTextoEspec.Espec_lin15)
        Next
        
        '== Insp_Obr
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Insp_Obr, "0") * -1
        '== Insp_Agend
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Insp_Agend, "0") * -1
        '== Dig_Bens
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Dig_Bens, "0") * -1
        '== Cob_Bens
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Cob_Bens, "0") * -1

        '== Seguradora de renovação
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Ren_Con_Seguradora, "00000")
        '== Apolice de renovação
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Ren_Con_Apolice, "000000000000000")
        '== Item de renovação
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Ren_Con_Item, "00000")

        '== Tip_Cobert
        lSequencia = lSequencia & Format$(lobjstcPedLoc.Tip_Cobert, "0")
        
    Next
    lSequencia2 = lSequencia2 & lSequencia
    lSequencia = lSequencia2
'-*-*-*-*-*
    
    '6. Preencher o campo de seqüência para que o tamanho fique em múltiplo de 10.
    lTamanho = Len(lSequencia)
    lResto = lTamanho Mod 10
    If lResto <> 0 Then
        lSequencia = lSequencia & String(10 - lResto, " ")
    End If

    '7. Dividir o campo de seqüência em campos de 10 posições cada.
    lTamanho = Len(lSequencia)
    lResultado = lTamanho / 10
    ReDim lCampo1(lResultado)
    For li = 1 To lResultado
        lPosInicial = ((li - 1) * 10) + 1
        lCampo1(li) = Mid$(lSequencia, lPosInicial, 10)
    Next li

    '8. Para cada campo de 10 posições, transforma cada caracter num valor ASC e soma estes valores.
    ReDim lCampo2(lResultado)
    For li = 1 To lResultado
        lCampo2(li) = Asc(Mid(lCampo1(li), 1, 1)) + Asc(Mid(lCampo1(li), 2, 1)) + _
                      Asc(Mid(lCampo1(li), 3, 1)) + Asc(Mid(lCampo1(li), 4, 1)) + _
                      Asc(Mid(lCampo1(li), 5, 1)) + Asc(Mid(lCampo1(li), 6, 1)) + _
                      Asc(Mid(lCampo1(li), 7, 1)) + Asc(Mid(lCampo1(li), 8, 1)) + _
                      Asc(Mid(lCampo1(li), 9, 1)) + Asc(Mid(lCampo1(li), 10, 1))
    Next li

    '9. Somar todos os resultados obtidos.
    lTotal = 0
    For li = 1 To lResultado
        lTotal = lTotal + lCampo2(li)
    Next li

    '10. Resultado: código de verificação.
    lCodVerificacao = Right$(Format$(lTotal, "000000"), 6)
    
    '11. Atualiza registro de dados comuns.
    pobjstcPedido.Cod_Unicidade = lCodVerificacao
    pobjstcPedido.Cod_Ver_Unici = 1   'Primeira versão da unicidade
    
    mfCalCodUnicidade_V1 = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
    
FechaObj:
    Set lobjstcCoss = Nothing
    Set lobjstcJuros = Nothing
    Set lObjstcVenc = Nothing
    Set lobjstcPedLoc = Nothing
    Set lobjstcCobert = Nothing
    Return
End Function

Public Function gfTratarVazio(pCampo As String) As String
    'Função: trata campo Vazio.
    If IsNull(pCampo) Then
        gfTratarVazio = " "
    ElseIf Len(pCampo) = 0 Then
        gfTratarVazio = " "
    Else
        gfTratarVazio = pCampo
    End If
End Function

Public Function gfVerTransmissao() As Boolean
    'Função: verifica se corretor pode realizar transmissão de propostas.

    Dim lCodCorretor1       As String       'Código do corretor no arquivo P2842400.EXE.
    Dim lCodCorretor2       As String       'Código do corretor no arquivo P366CORR.EXE.
    Dim lEncontrou          As Boolean      'Indica que encontrou o registro de controle.
    Dim lEndP2842400        As String       'Endereço do arquivo P2842400.EXE.
    Dim lEndP366CORR        As String       'Endereço do arquivo P366CORR.EXE.
    Dim lIndTransmissao     As String       'Indicador de transmissão: S ou N.
    Dim lNumP2842400        As Integer      'Utilizado com FreeFile para o arquivo P2842400.EXE.
    Dim lNumP366CORR        As Integer      'Utilizado com FreeFile para o arquivo P366CORR.EXE.
    Dim lRegistro           As String       'Registro.

'*********************************************************
'*****DESATIVADO    TODOS OS CORRETORES PODEM TRANSMITIR PROPOSTAS
'*****MULTISSEGUROS

    gfVerTransmissao = True

    Exit Function
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'    gfVerTransmissao = False
'
'    '1. Obtém código do corretor.
'    lEndP2842400 = gAppPath & "P2842400.EXE"
'    lNumP2842400 = FreeFile
'    On Error Resume Next
'    Open lEndP2842400 For Input Access Read As lNumP2842400
'    If Err = 0 Then     'Encontrou arquivo P2842400.EXE.
'        Line Input #lNumP2842400, lRegistro
'        lCodCorretor1 = Format(lRegistro, "00000000")
'    Else
'        lCodCorretor1 = "00000000"
'    End If
'    On Error GoTo 0
'    Close lNumP2842400
'
'    '2. Verifica preenchimento.
'    If Not gfPreenchido(lCodCorretor1) Then
'        Call gpGraLog(1, "FGE0030 - Para transmissão de propostas para a Yasuda via Internet, " & _
'                         "solicite o termo de ""ADESÃO AO SISTEMA DE TRANSMISSÃO DE PROPOSTAS DE " & _
'                         "MULTISEGUROS"" com o produtor (código do corretor não preenchido).")
'        Exit Function
'    End If
'    lCodCorretor1 = Trim$(lCodCorretor1)
'    If Not IsNumeric(lCodCorretor1) Then
'        Call gpGraLog(1, "FGE0076 - Para transmissão de propostas para a Yasuda via Internet, " & _
'                         "solicite o termo de ""ADESÃO AO SISTEMA DE TRANSMISSÃO DE PROPOSTAS DE " & _
'                         "MULTISEGUROS"" com o produtor (código do corretor não numérico).")
'        Exit Function
'    End If
'    If Val(lCodCorretor1) = 0 Then
'        'Call gpGraLog(1, "FGE0077 - Para transmissão de propostas para a Yasuda via Internet, " & _
'        '                 "solicite o termo de ""ADESÃO AO SISTEMA DE TRANSMISSÃO DE PROPOSTAS DE " & _
'        '                 "MULTISEGUROS"" com o produtor (código do corretor zerado).")
'        Exit Function
'    End If
'    Select Case Len(lCodCorretor1)
'        Case Is < 7
'            lCodCorretor1 = String(7 - Len(lCodCorretor1), "0") & lCodCorretor1
'        Case Is > 7
'            lCodCorretor1 = Right$(lCodCorretor1, 7)
'    End Select
'
'    '3. Abre aquivo de controle de transmissão (temporário).
'    lEndP366CORR = gAppPath & "P366CORR.EXE"
'    If Not gfPreenchido(Dir(lEndP366CORR)) Then
'        Call gpGraLog(1, "FTE0061 - Arquivo não localizado: controle do termo de adesão.")
'        Exit Function
'    End If
'    lNumP366CORR = FreeFile
'    On Error Resume Next
'    Open lEndP366CORR For Input Access Read As lNumP366CORR
'    If Err <> 0 Then
'        Call gpGraLog(1, "FTE0062 - Erro ao abrir arquivo de controle do termo de adesão (" & Err & _
'                         " - " & Error & ").")
'        On Error GoTo 0
'        Exit Function
'    End If
'    On Error GoTo 0
'
'    '4. Trata registros do arquivo de inicialização.
'    lEncontrou = False
'    Do While Not EOF(lNumP366CORR)
'        Line Input #lNumP366CORR, lRegistro
'        If gfPreenchido(lRegistro) Then
'            lRegistro = Trim$(lRegistro)
'            If UCase(lRegistro) = "[GRUPOS]" Then
'                Exit Do
'            End If
'            If UCase(lRegistro) = "[SOMENTE MULTI]" Then
'                Exit Do
'            End If
'            Select Case Len(lRegistro)
'                Case 8
'                    lCodCorretor2 = Left$(lRegistro, 7)
'                    lIndTransmissao = "N"
'                Case 9
'                    lCodCorretor2 = Left$(lRegistro, 7)
'                    lIndTransmissao = Mid$(lRegistro, 9, 1)
'                Case Else
'                    lCodCorretor2 = "0000000"
'                    lIndTransmissao = "N"
'            End Select
'            If lCodCorretor1 = lCodCorretor2 Then
'                lEncontrou = True
'                Exit Do
'            End If
'        End If
'    Loop
'    Close lNumP366CORR
'    If lEncontrou = False Then
'        Call gpGraLog(1, "FGE0078 - Para transmissão de propostas para a Yasuda via Internet, " & _
'                         "solicite o termo de ""ADESÃO AO SISTEMA DE TRANSMISSÃO DE PROPOSTAS DE " & _
'                         "MULTISEGUROS"" com o produtor (código do corretor inexistente no controle de " & _
'                         "transmissão).")
'        Exit Function
'    End If
'    If lIndTransmissao <> "S" Then
'        Call gpGraLog(1, "FGE0079 - Para transmissão de propostas para a Yasuda via Internet, " & _
'                         "solicite o termo de ""ADESÃO AO SISTEMA DE TRANSMISSÃO DE PROPOSTAS DE " & _
'                         "MULTISEGUROS"" com o produtor (corretor sem autorização).")
'        Exit Function
'    End If

    gfVerTransmissao = True
End Function

Private Sub mpCarregaCorretor(ByRef pCod_Corretor As Long)
    'Função: Carrega código do corretor.

    'Parâmetro de saída...pCod_Corretor...Código do corretor.

    Dim lArquivo    As String   'Endereço do arquivo.
    Dim lNumArquivo As Integer  'Utilizado com FreeFile.
    Dim lRegistro   As String   'Registro.

    '1. Inicializa variável.
    pCod_Corretor = 0

    '2. Obtém dados no arquivo P2842400.EXE.
    lArquivo = gAppPath & "P2842400.EXE"
    lNumArquivo = FreeFile
    On Error Resume Next
    Open lArquivo For Input As #lNumArquivo
    If Err = 0 Then
        Line Input #lNumArquivo, lRegistro
        If gfPreenchido(lRegistro) Then
            If IsNumeric(lRegistro) Then
                If Len(lRegistro) >= 8 Then
                    pCod_Corretor = CLng(Mid$(lRegistro, 2, 7))
                Else
                    lRegistro = Format(lRegistro, "00000000")
                    pCod_Corretor = CLng(Mid$(lRegistro, 2, 7))
                End If
            End If
        End If
    End If
    On Error GoTo 0
    Close #lNumArquivo
End Sub
Public Sub gpForCNPJ(ByVal pOpcao As Byte, ByVal pEntrada As String, ByRef pSaida As String)
    'Procedure: formata campo de CNPJ.

    'Parâmetros de entrada...pOpcao.....1 = coloca ".", "/" e "-"; 2 = tira ".", "/" e "-".
    '                        pEntrada...Campo de entrada.
    'Parâmetro de saída......pSaida.....Capo de saída (não pode ser o mesmo campo de entrada).

    Dim lAuxiliar1  As String   'Auxiliar 1.
    Dim lAuxiliar2  As String   'Auxiliar 2.
    Dim li          As Byte     'Utilizado com For...Next.
    Dim lPosicao    As Byte     'Utilizado com InStr.
    Dim lTamanho    As Byte     'Tamanho.

    Select Case pOpcao
        Case 1      'Coloca ".", "/" e "-".
            'a) Inicializa campo de saída.
            pSaida = ""
            'b) Se campo de entrada não preenchido, nada a fazer.
            If Not gfPreenchido(pEntrada) Then
                Exit Sub
            End If
            'c) Se campo de entrada com 18 posições, nada a fazer.
            lAuxiliar1 = Trim(pEntrada)
            lTamanho = Len(lAuxiliar1)
            If lTamanho >= 18 Then
                Exit Sub
            End If
            'd) Prepara campo auxiliar.
            lPosicao = InStr(lAuxiliar1, "-")
            If lPosicao <> 0 Then
                Select Case Len(lAuxiliar1)
                    Case lPosicao + 1
                        lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & "0" & Right$(lAuxiliar1, 1)
                    Case lPosicao + 2
                        lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & Right$(lAuxiliar1, 2)
                    Case Else
                        Exit Sub
                End Select
            End If
            lPosicao = InStr(lAuxiliar1, "/")
            If lPosicao <> 0 Then
                Select Case Len(lAuxiliar1)
                    Case lPosicao + 3
                        lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & "000" & Right$(lAuxiliar1, 3)
                    Case lPosicao + 4
                        lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & "00" & Right$(lAuxiliar1, 4)
                    Case lPosicao + 5
                        lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & "0" & Right$(lAuxiliar1, 5)
                    Case lPosicao + 6
                        lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & Right$(lAuxiliar1, 6)
                    Case Else
                        Exit Sub
                End Select
            End If
            lTamanho = Len(lAuxiliar1)
            lAuxiliar2 = ""
            For li = 1 To lTamanho
                Select Case Mid$(lAuxiliar1, li, 1)
                    Case ".", ","
                        'Nada a fazer.
                    Case Else
                        lAuxiliar2 = lAuxiliar2 & Mid$(lAuxiliar1, li, 1)
                End Select
            Next li
            If Not IsNumeric(lAuxiliar1) Then
                Exit Sub
            End If
            If Len(lAuxiliar2) < 14 Then
                lAuxiliar2 = String(14 - Len(lAuxiliar2), "0") & lAuxiliar2
            End If
            'e) Formata CNPJ - XX.XXX.XXX/XXXX-XX.
            pSaida = Mid$(lAuxiliar2, 1, 2) & "." & _
                     Mid$(lAuxiliar2, 3, 3) & "." & _
                     Mid$(lAuxiliar2, 6, 3) & "/" & _
                     Mid$(lAuxiliar2, 9, 4) & "-" & _
                     Mid$(lAuxiliar2, 13, 2)
        Case 2      'Retira ".", "/" e "-"
            'a) Inicializa campo de saída.
            pSaida = "00000000000000"
            'b) Se campo de entrada diferente de 18 posições, nada a fazer.
            lAuxiliar1 = Trim(pEntrada)
            lTamanho = Len(lAuxiliar1)
            If lTamanho <> 18 Then
                Exit Sub
            End If
            'c) Se campo de entrada fora do formato XX.XXX.XXX/XXXX-XX, nada a fazer.
            If Mid$(lAuxiliar1, 3, 1) <> "." Or _
               Mid$(lAuxiliar1, 7, 1) <> "." Or _
               Mid$(lAuxiliar1, 11, 1) <> "/" Or _
               Mid$(lAuxiliar1, 16, 1) <> "-" Then
                Exit Sub
            End If
            'd) Prepara campo auxiliar.
            lAuxiliar2 = Mid$(lAuxiliar1, 1, 2) & _
                         Mid$(lAuxiliar1, 4, 3) & _
                         Mid$(lAuxiliar1, 8, 3) & _
                         Mid$(lAuxiliar1, 12, 4) & _
                         Mid$(lAuxiliar1, 17, 2)
            If Not IsNumeric(lAuxiliar2) Then
                Exit Sub
            End If
            pSaida = lAuxiliar2
    End Select
End Sub

Public Sub gpForCPF(ByVal pOpcao As Byte, ByVal pEntrada As String, ByRef pSaida As String)
    'Procedure: formata campo de CPF.

    'Parâmetros de entrada...pOpcao.....1 = coloca "." e "-"; 2 = tira "." e "-".
    '                        pEntrada...Campo de entrada.
    'Parâmetro de saída......pSaida.....Capo de saída (não pode ser o mesmo campo de entrada).

    Dim lAuxiliar1  As String   'Auxiliar 1.
    Dim lAuxiliar2  As String   'Auxiliar 2.
    Dim li          As Byte     'Utilizado com For...Next.
    Dim lPosicao    As Byte     'Utilizado com InStr.
    Dim lTamanho    As Byte     'Tamanho.

    Select Case pOpcao
        Case 1      'Coloca "." e "-"
            'a) Inicializa campo de saída.
            pSaida = ""
            'b) Se campo de entrada não preenchido, nada a fazer.
            If Not gfPreenchido(pEntrada) Then
                Exit Sub
            End If
            'c) Se campo de entrada com 14 posições, nada a fazer.
            lAuxiliar1 = Trim(pEntrada)
            lTamanho = Len(lAuxiliar1)
            If lTamanho >= 14 Then
                pSaida = lAuxiliar1
                Exit Sub
            End If
            'd) Prepara campo auxiliar.
            lPosicao = InStr(lAuxiliar1, "-")
            If lPosicao <> 0 Then
                If Len(lAuxiliar1) = lPosicao + 1 Then
                    lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & "0" & Right$(lAuxiliar1, 1)
                ElseIf Len(lAuxiliar1) = lPosicao + 2 Then
                    lAuxiliar1 = Left$(lAuxiliar1, lPosicao - 1) & Right$(lAuxiliar1, 2)
                Else
                    pSaida = lAuxiliar1
                    Exit Sub
                End If
            End If
            lTamanho = Len(lAuxiliar1)
            lAuxiliar2 = ""
            For li = 1 To lTamanho
                Select Case Mid$(lAuxiliar1, li, 1)
                    Case ".", ","
                        'Nada a fazer.
                    Case Else
                        lAuxiliar2 = lAuxiliar2 & Mid$(lAuxiliar1, li, 1)
                End Select
            Next li
            If Not IsNumeric(lAuxiliar1) Then
                pSaida = lAuxiliar1
                Exit Sub
            End If
            If Len(lAuxiliar2) < 11 Then
                lAuxiliar2 = String(11 - Len(lAuxiliar2), "0") & lAuxiliar2
            End If
            'e) Formata CPF - XXX.XXX.XXX-XX.
            pSaida = Mid$(lAuxiliar2, 1, 3) & "." & _
                     Mid$(lAuxiliar2, 4, 3) & "." & _
                     Mid$(lAuxiliar2, 7, 3) & "-" & _
                     Mid$(lAuxiliar2, 10, 2)
        Case 2      'Retira "." e "-"
            'a) Inicializa campo de saída.
            pSaida = "00000000000"
            'b) Se campo de entrada diferente de 14 posições, nada a fazer.
            lAuxiliar1 = Trim(pEntrada)
            lTamanho = Len(lAuxiliar1)
            If lTamanho <> 14 Then
                Exit Sub
            End If
            'c) Se campo de entrada fora do formato XXX.XXX.XXX-XX, nada a fazer.
            If Mid$(lAuxiliar1, 4, 1) <> "." Or _
               Mid$(lAuxiliar1, 8, 1) <> "." Or _
               Mid$(lAuxiliar1, 12, 1) <> "-" Then
                Exit Sub
            End If
            'd) Prepara campo auxiliar.
            lAuxiliar2 = Mid$(lAuxiliar1, 1, 3) & _
                         Mid$(lAuxiliar1, 5, 3) & _
                         Mid$(lAuxiliar1, 9, 3) & _
                         Mid$(lAuxiliar1, 13, 2)
            If Not IsNumeric(lAuxiliar2) Then
                Exit Sub
            End If
            pSaida = lAuxiliar2
    End Select
End Sub

Public Function gfNomUniProdutiva(ByVal pCodUniNegocios As Integer, ByRef pNomUniNegocios As String, _
                                  ByRef pMensagem As String) As Boolean
    'Procedure: obtém nome da unidade de negócios.

    'Parâmetro de entrada...pCodUniNegocios...Código da unidade de negócios.
    'Parâmetros de saída....pNomUniNegocios...Nome da unidade de negócios.
    '                       pMensagem.........Mensagem.

Dim bdP0044700  As ADODB.Connection 'Arquivo - P0044700.
Dim rsDEPT      As ADODB.Recordset  'Registro - YTAB_DEPT.
Dim lSelect     As String

    gfNomUniProdutiva = False

    '1. Inicializa controle.
    pNomUniNegocios = " "

    '2. Verifica se código está preenchido.
    If IsNull(pCodUniNegocios) Then
        pMensagem = "FNP0008 - Unidade de negócios inválida (código não preenchido)."
        GoTo SAIDA
    End If

    '3. Busca no arquivo.
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        GoTo SAIDA
    End If
    lSelect = "SELECT NOM_DEPT FROM YTAB_DEPT WHERE NUM_DEPT = " & pCodUniNegocios
    If gfObtRegistro(bdP0044700, lSelect, rsDEPT, pMensagem) = False Then
        GoTo SAIDA
    End If
    If rsDEPT.EOF = True Then
        Call gpFecha3(rsDEPT, bdP0044700)
        pMensagem = "FNL0010 - Código da unidade de negócios não localizado (" & pCodUniNegocios & ")."
        GoTo SAIDA
    End If

    '4. Posiciona nome da unidade produtiva.
    pNomUniNegocios = rsDEPT!NOM_DEPT

    gfNomUniProdutiva = True
   
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Call gpFecha3(rsDEPT, bdP0044700)
    Set bdP0044700 = Nothing
    Set rsDEPT = Nothing
    Return
End Function
Public Function gfNomProdutor(ByVal pCodProdutor As Integer, ByRef pNomProdutor As String, _
                              ByRef pMensagem As String) As Boolean
    'Função: obtém nome do produtor.

    'Parâmetro de entrada...pCodProdutor...Código do produtor.
    'Parâmetros de saída....pNomProdutor...Nome do produtor.
    '                       pMensagem......Mensagem.

Dim bdP0044700  As ADODB.Connection 'Arquivo - P0044700.
Dim rsPRODUTOR  As ADODB.Recordset  'Registro - YTAB_PRODUTOR.
Dim lSelect     As String

    gfNomProdutor = False

    '1. Inicializa controle.
    pNomProdutor = " "

    '2. Verifica se código está prenchido.
    If Not gfPreenchido(pCodProdutor) Then
        pMensagem = "FNP0012 - " & gPrefixo & "Produtor inválido (código do produtor não preenchido)."
        GoTo SAIDA
    End If

    '3. Obtém nome do produtor.
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        GoTo SAIDA
    End If
    lSelect = "SELECT NOM_PROD FROM YTAB_PRODUTOR WHERE COD_PROD = " & pCodProdutor
    If gfObtRegistro(bdP0044700, lSelect, rsPRODUTOR, pMensagem) = False Then
        GoTo SAIDA
    End If
    If rsPRODUTOR.EOF = True Then
        pNomProdutor = "PRODUTOR " & pCodProdutor & "."
    Else
        pNomProdutor = rsPRODUTOR!NOM_PROD
    End If

    gfNomProdutor = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Call gpFecha3(rsPRODUTOR, bdP0044700)
    Set bdP0044700 = Nothing
    Set rsPRODUTOR = Nothing
    Return
End Function

Public Function gfBusProduto(ByVal pVigInicio As Double, ByRef pCodProduto As Integer, _
                             ByRef pCodProtocolo As Long, ByRef pCodSequencia As Long, _
                             ByRef pDescProduto As String, ByRef pMensagem As String, _
                             Optional pCodRamo As Integer, Optional pEndosso As Boolean = False) As Boolean
    'Procedure: busca produto de acordo com início de vigência.

    'Parâmetro de entrada...pVigInicio......Início de vigência da cotação/proposta.
    'Parâmetros de saída....pCodProduto.....Código do produto.
    '                       pCodProtocolo...Código do protocolo.
    '                       pCodSequencia...Código da seqüência.
    '                       pMensagem.......Mensagem.

Dim BDP00Tarifa     As ADODB.Connection 'Arquivo - P00Tarifa.
Dim lAuxAno         As Integer          'Auxiliar: ano.
Dim lAuxiliar       As String           'Auxiliar.
Dim lAuxMes         As Integer          'Auxiliar: mês.
Dim lDatLimite      As Double           'Data limite.
Dim lAuxData        As Double           'Data auxiliar
Dim lAuxPreData     As String           'Auxiliar.
Dim lDias           As Integer          'Dias.
Dim lDifDias        As Integer          'Dif dias.
Dim rsTabProd       As ADODB.Recordset  'TAB_PRODUTO: tabela.
Dim lSelect         As String

    gfBusProduto = False

    '1. Inicializa variáveis.
    pCodProduto = 0
    pCodProtocolo = 0
    pCodSequencia = 0

    '2. Abre arquivo.
    If gfAbrBasDados("P00MULTI2", "1", gEndP00MULTI, BDP00Tarifa, pMensagem) = False Then
        GoTo SAIDA
    End If

    '3. Obtém produto sem término de vigência ou com término de vigência superior à data do dia.
    'lSelect = "SELECT * FROM TAB_PRODUTO WHERE Tip_Prot = 1 AND Dat_Vig_Ini <= " & pVigInicio
    
    'Tip_prot  = 1  Tarifa Corretor
    'Tip_prot  = 2  Tarifa Interna
    
    lSelect = "SELECT * FROM TAB_PRODUTO WHERE Dat_Vig_Ini <= " & pVigInicio
    If gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        lSelect = lSelect & " and Tip_Prot = 1 "
    End If
    If pCodRamo <> 0 Then
        lSelect = lSelect & " and Cod_ramo = " & pCodRamo
    End If
    lSelect = lSelect & " ORDER BY Cod_Produto DESC , Cod_Sequencia desc"
    If gfObtRegistro(BDP00Tarifa, lSelect, rsTabProd, pMensagem) = False Then
        GoTo SAIDA
    End If
    
    If rsTabProd.EOF Then
        lSelect = "SELECT * FROM TAB_PRODUTO WHERE Dat_Vig_Ini <= " & pVigInicio
        If pCodRamo <> 0 Then
            lSelect = lSelect & " and Cod_ramo = " & pCodRamo
        End If
        lSelect = lSelect & " ORDER BY Cod_Produto DESC"
        If gfObtRegistro(BDP00Tarifa, lSelect, rsTabProd, pMensagem) = False Then
            GoTo SAIDA
        End If
        If rsTabProd.EOF Then
            gfBusProduto = True
            GoTo SAIDA
        End If
    End If
    
    '4. Verifica produtos disponíveis.
    If pVigInicio > rsTabProd!Dat_Vig_Ter Then
        lAuxiliar = Format$(rsTabProd!Dat_Vig_Ter, "00000000")
        lAuxMes = CInt(Mid$(lAuxiliar, 5, 2)) + 1
        lAuxAno = CInt(Mid$(lAuxiliar, 1, 4))
        If lAuxMes = 13 Then
            lAuxMes = 1
            lAuxAno = lAuxAno + 1
        End If
        lDias = -1
        On Error Resume Next
        lDias = rsTabProd!Dias_Expiracao
        If lDias = -1 Then
            lDias = 15
        End If
        On Error GoTo 0
        lAuxPreData = CDate(Mid(rsTabProd!Dat_Vig_Ter, 7, 2) & "/" & Mid(rsTabProd!Dat_Vig_Ter, 5, 2) & "/" & Mid(rsTabProd!Dat_Vig_Ter, 1, 4)) + lDias
        lAuxData = Format(lAuxPreData, "YYYYMMDD")
        
        If Not (pVigInicio >= rsTabProd!Dat_Vig_Ini And pVigInicio <= lAuxData) Then
            'Documento dentro do limite do produto (até x dias após término da
            'vigência) - sem problema.
   '?         gfBusProduto = True
   '?         GoTo SAIDA
        End If
    End If

    '5. Posiciona dados de saída.
    pCodProduto = Check(rsTabProd!Cod_Produto)
    pCodProtocolo = Check(rsTabProd!Cod_Protocolo)
    pCodSequencia = Check(rsTabProd!Cod_Sequencia)
    pDescProduto = Check(rsTabProd("Nom_Prot"))
    gProCusEndosso = Check(rsTabProd!Val_Cus_Emis_Endo)
    gProCusApolice = Check(rsTabProd!Val_Cus_Emis_Apol)
    gProPreMinParcela = Check(rsTabProd!Par_Min_Parcela)
    gProPreMinParcelaTAB = gProPreMinParcela
    If pEndosso Then
        gProPremioMinimo = Check(rsTabProd!Pre_Min_Endosso)
        gProPremioMinimoTAB = Check(rsTabProd!Pre_Min_Endosso)
    Else
        gProPremioMinimo = Check(rsTabProd!Pre_Min_Parcela)
        gProPremioMinimoTAB = Check(rsTabProd!Pre_Min_Parcela)
    End If

    gfBusProduto = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Call gpFecha3(rsTabProd, BDP00Tarifa)
    Set rsTabProd = Nothing
    Set BDP00Tarifa = Nothing
    Return
End Function


Public Function gfDifDatas(ByVal pInicio As Long, ByVal pTermino As Long) As Single
    'Função #014: calcula diferença entre datas.

    'Parâmetros de entrada...pInicio....data de início.
    '                        pTermino...data de término.

    Dim lAuxIniDias         As String   'Auxiliar: valores para tabela para cálculo da diferença de dias.
    Dim lAuxTerDias         As String   'Auxiliar: valores para tabela para cálculo da diferença de dias.
    Dim lDia                As Single   'Índice auxiliar: dia.
    Dim li                  As Integer  'Índice.
    Dim lIniDia             As Single   'Dia do início de vigência.
    Dim lIniMes             As Single   'Mês do início de vigência.
    Dim lMes                As Single   'Índice auxiliar: mês.
    Dim lQuaIniDias         As Single   'Quantidade de dias (início de vigência).
    Dim lQuaTerDias         As Single   'Quantidade de dias (término de vigência).
    Dim lTabIniDias(31, 12) As String   'Tabela para cálculo de dias (data de início).
    Dim lTabTerDias(31, 12) As String   'Tabela para cálculo de dias (data de término).
    Dim lTerDia             As Single   'Dia do término de vigência.
    Dim lTerMes             As Single   'Mês do término de vigência.

    '1. Inicializa valores auxiliares para cálculo da diferença de dias.
    lAuxIniDias = "001032060091121152182213244274305335" & _
                  "002033061092122153183214245275306336" & _
                  "003034062093123154184215246276307337" & _
                  "004035063094124155185216247277308338" & _
                  "005036064095125156186217248278309339" & _
                  "006037065096126157187218249279310340" & _
                  "007038066097127158188219250280311341" & _
                  "008039067098128159189220251281312342" & _
                  "009040068099129160190221252282313343" & _
                  "010041069100130161191222253283314344" & _
                  "011042070101131162192223254284315345" & _
                  "012043071102132163193224255285316346" & _
                  "013044072103133164194225256286317347" & _
                  "014045073104134165195226257287318348" & _
                  "015046074105135166196227258288319349" & _
                  "016047075106136167197228259289320350" & _
                  "017048076107137168198229260290321351" & _
                  "018049077108138169199230261291322352" & _
                  "019050078109139170200231262292323353" & _
                  "020051079110140171201232263293324354" & _
                  "021052080111141172202233264294325355" & _
                  "022053081112142173203234265295326356" & _
                  "023054082113143174204235266296327357" & _
                  "024055083114144175205236267297328358" & _
                  "025056084115145176206237268298329359"
    lAuxIniDias = lAuxIniDias & _
                  "026057085116146177207238269299330360" & _
                  "027058086117147178208239270300331361" & _
                  "028059087118148179209240271301332362" & _
                  "029000088119149180210241272302333363" & _
                  "030000089120150181211242273303334364" & _
                  "031000090000151000212243000304000365"

    lAuxTerDias = "366397425456486517547578609639670700" & _
                  "367398426457487518548579610640671701" & _
                  "368399427458488519549580611641672702" & _
                  "369400428459489520550581612642673703" & _
                  "370401429460490521551582613643674704" & _
                  "371402430461491522552583614644675705" & _
                  "372403431462492523553584615645676706" & _
                  "373404432463493524554585616646677707" & _
                  "374405433464494525555586617647678708" & _
                  "375406434465495526556587618648679709" & _
                  "376407435466496527557588619649680710" & _
                  "377408436467497528558589620650681711" & _
                  "378409437468498529559590621651682712" & _
                  "379410438469499530560591622652683713" & _
                  "380411439470500531561592623653684714" & _
                  "381412440471501532562593624654685715" & _
                  "382413441472502533563594625655686716" & _
                  "383414442473503534564595626656687717" & _
                  "384415443474504535565596627657688718" & _
                  "385416444475505536566597628658689719" & _
                  "386417445476506537567598629659690720" & _
                  "387418446477507538568599630660691721" & _
                  "388419447478508539569600631661692722" & _
                  "389420448479509540570601632662693723" & _
                  "390421449480510541571602633663694724"
    lAuxTerDias = lAuxTerDias & _
                  "391422450481511542572603634664695725" & _
                  "392423451482512543573604635665696726" & _
                  "393424452483513544574605636666697727" & _
                  "394000453484514545575606637667698728" & _
                  "395000454485515546576607638668699729" & _
                  "396000455000516000577608000669000730"

    '2. Posiciona dados da tabela de data de início.
    li = 1
    lDia = 1
    Do Until lDia > 31
        lMes = 1
        Do Until lMes > 12
            lTabIniDias(lDia, lMes) = Val(Mid(lAuxIniDias, li, 3))
            lMes = lMes + 1
            li = li + 3
        Loop
        lDia = lDia + 1
    Loop

    '3. Posiciona dados da tabela de data de término.
    lDia = 1
    li = 1
    Do Until lDia > 31
        lMes = 1
        Do Until lMes > 12
            lTabTerDias(lDia, lMes) = Val(Mid(lAuxTerDias, li, 3))
            lMes = lMes + 1
            li = li + 3
        Loop
        lDia = lDia + 1
    Loop

    '4. Apropria informação de dia e mês de início e término.
    lIniDia = Val(Mid(pInicio, 7, 2))
    lIniMes = Val(Mid(pInicio, 5, 2))
    lTerDia = Val(Mid(pTermino, 7, 2))
    lTerMes = Val(Mid(pTermino, 5, 2))

    '5. Posiciona quantidade de dias.
    lQuaIniDias = lTabIniDias(lIniDia, lIniMes)
    If Val(Mid(pInicio, 1, 4)) <> Val(Mid(pTermino, 1, 4)) Then
        lQuaTerDias = lTabTerDias(lTerDia, lTerMes)
    Else
        lQuaTerDias = lTabIniDias(lTerDia, lTerMes)
    End If

    '6. Calcula diferença de dias.
    gfDifDatas = lQuaTerDias - lQuaIniDias
End Function
Public Function gfCalNosNumero(ByRef pNosNumero As String, ByRef pMensagem As String) As Boolean
    'Função: gera nosso número a partir de usuário, ano, dia juliano e número de segundos decorridos no
    '        dia.

    'Parâmetros de saída...pNosNumero...Nosso número.
    '                      pMensagem....Mensagem.

    'Formato do nosso número:
    '   AAADDDSSSSSUUUUUUUUD...AAA........ano.
    '                          DDD........dia juliano (dias decorridos no ano).
    '                          SSSSS......segundos decorridos no dia.
    '                          UUUUUUUU...usuário.
    '                          D..........dígito verificador (pela Prova 11).

    Dim clsA99V001      As clsA99V001       'Calcula dígitos.
    Dim lAuxAno         As String           'Ano no formato string.
    Dim lAuxDia         As String           'Dia no formato string.
    Dim lAuxSegundo     As String           'Segundo no formato string.
    Dim lCodUsuario     As String           'Código do usuário.
    Dim lDia            As Integer          'Dia.
    Dim lDigito         As Integer          'Dígito.
    Dim lMesBissexto    As String           'Controle de dias do ano bissexto.
    Dim lMesNormal      As String           'Controle de dias do ano normal.
    Dim lNosNumero      As String           'Nosso número.
    Dim lSegundo        As Double           'Segundo.

    gfCalNosNumero = False

    '1. Posiciona dados para nosso número.
    'a) Ano.
    lAuxAno = Right$(Format$(gHojAno, "0000"), 3)
    'b) Dia juliano.
    lMesNormal = "000031059090120150181212243273303334365"
    lMesBissexto = "000031060091121151182213244274304335366"
    lDia = gHojDia
    Select Case gHojAno Mod 4
        Case 0      'Ano bissexto.
            lDia = lDia + Val(Mid(lMesBissexto, gHojMes * 3 - 2, 3))
        Case Else   'Ano NÃO bissexto.
            lDia = lDia + Val(Mid(lMesNormal, gHojMes * 3 - 2, 3))
    End Select
    lAuxDia = Right("000" & Trim(Str(lDia)), 3)
    'c) Segundos decorridos do dia.
    lSegundo = (Hour(Now) * 3600) + (Minute(Now) * 60) + Second(Now)
    lAuxSegundo = Right(Format$(lSegundo, "00000"), 5)
    'd) Código do usuário.
    lCodUsuario = gNomUsuario
    If IsNumeric(lCodUsuario) Then
        lCodUsuario = Trim$(lCodUsuario)
        If Len(lCodUsuario) < 8 Then
            lCodUsuario = String(8 - Len(lCodUsuario), "0") & lCodUsuario
        End If
    Else
        If gCod_Corretor <> 0 Then
            lCodUsuario = Format(gCod_Corretor, "00000000")
        Else
            lCodUsuario = "00000000"
        End If
    End If
    'e) Monta nosso número sem dígito.
    lNosNumero = lAuxAno & lAuxDia & lAuxSegundo & lCodUsuario

    '2. Calcula dígito.
    Set clsA99V001 = New clsA99V001
    If clsA99V001.DIG_MODULO11(lNosNumero, "20191817161514131211100908070605040302", lDigito, _
       pMensagem) = False Then
        Exit Function
    End If
    Set clsA99V001 = Nothing

    '3. Completa nosso número.
    pNosNumero = lNosNumero & Format$(lDigito, "0")
    gfCalNosNumero = True
End Function

Public Function gfDisponivel(ByVal pNome As String) As Boolean
    'Função: indica se form está disponível.

    Dim li  As Byte     'Utilizado com For...Next.

    gfDisponivel = False
    For li = 0 To (Forms.Count - 1)
        If UCase(Forms(li).Name) = UCase(pNome) Then
            gfDisponivel = True
            Exit For
        End If
    Next li
End Function
Public Function gfExecutaTransmissao() As Boolean
    'Função: Aciona componente EDIWare.
    ChDir gAppPath
    Shell gAppPath & "client.bat", vbMinimizedNoFocus
End Function
Public Function gfForAno(pData As Integer) As String
    If Left$(pData, 2) > 50 Then    'De 1951 a 1999.
        gfForAno = "19" & Format$(pData, "00")
    Else                            'De 2000 a 2050.
        gfForAno = "20" & Format$(pData, "00")
    End If
End Function
Public Function gfForAspas(ByVal pCampo As String) As String
    'Função: formata campo texto com aspas para ser utilizado no comando SQL.

    Dim lAuxiliar   As String       'Utilizado com InStr.
    Dim lPosicao    As Integer      'Utilizado com InStr.
    Dim lTemAspa    As Boolean      'Indicador de que tem aspa.
    
    '1. Se campo não preenchido, retorna campo nulo.
    If Not gfPreenchido(pCampo) Then
        gfForAspas = ""
        Exit Function
    End If

    '2. Posiciona controle.
    lTemAspa = True

    '3. Verifica se campo tem aspa.
    lPosicao = InStr(pCampo, "'")
    
    '4. Verifica se campo tem "|".
    If lPosicao = 0 Then
        gfForAspas = pCampo
        lTemAspa = False
        lPosicao = InStr(pCampo, "|")
        If lPosicao = 0 Then
            gfForAspas = pCampo
            Exit Function
        End If
    End If

    '5. Se campo tem aspa, duplica campo de aspa para comando SQL.
    lAuxiliar = pCampo
    If lTemAspa = True Then
        Do
            If lPosicao = Len(lAuxiliar) Then
                gfForAspas = lAuxiliar & "'"
                Exit Function
            End If
            lAuxiliar = Left$(lAuxiliar, lPosicao) & "'" & Mid$(lAuxiliar, lPosicao + 1)
            lPosicao = InStr(lPosicao + 2, lAuxiliar, "'")
            If lPosicao = 0 Then
                lPosicao = InStr(pCampo, "|")
                If lPosicao = 0 Then
                    gfForAspas = lAuxiliar
                    Exit Function
                Else
                    Exit Do
                End If
            End If
        Loop
    End If

    '6. Retira "|" do campo.
    Do
        If lPosicao = Len(lAuxiliar) Then
            gfForAspas = Left$(lAuxiliar, lPosicao - 1)
            Exit Function
        End If
        lAuxiliar = Left$(lAuxiliar, lPosicao - 1) & Mid$(lAuxiliar, lPosicao + 1)
        lPosicao = InStr(lPosicao + 1, lAuxiliar, "|")
        If lPosicao = 0 Then
            gfForAspas = lAuxiliar
            Exit Function
        End If
    Loop
End Function
Public Function gfForData(ByVal pOpcao As Byte, Optional ByVal pData, Optional ByVal pTipArquivo) _
                          As String
    'Função: formata data.

    'Formato dos campos de data (numérico duplo): aaaammdd.

    'Parâmetros...pOpcao.........01 = formata campo de "dd/mm/aaaa" para "aaaammdd" para comandos SQL.
    '                            02 = formata campo de "aaaammdd" para "dd/mm/aaaa".
    '                            03 = fornece data atual no formato aaaammddhhmmss.
    '                            04 = formata campo de "aaaammddhhmmss" para "dd/mm/aaaa hh:mm:ss".
    '                            05 = formata campo de "aaaammdd" para formato conveniente
    '                                 ('mm/dd/aaaa' ou #mm/dd/aaaa#).
    '                            06 = formata campo de "dd/mm/aaaa" para formato conveniente
    '                                 ('mm/dd/aaaa' ou #mm/dd/aaaa#).
    '                            07 = formata campo de "aammdd" para "dd/mm/aaaa".
    '                            08 = formata campo de ano para "aaaa".
    '                            09 = trata campo de data na digitação (utilizado nos eventos LostFocus).
    '                            10 = formata campo de "aammdd" para "aaaammdd".
    '                            11 = fornece data atual no formato aaaammdd.
    '                            12 = formata campo de "aaaammdd" para "dd/mm/aaaa" se não for zero
    '             pData..........Campo de data a ser tratado (opcional).
    '             pTipArquivo...Tipo do arquivo (opcional - utilizado para opção 5).

    Dim lAno        As String   'Ano.
    Dim lAuxData    As String   'Auxiliar: data.
    Dim lDia        As String   'Dia.
    Dim lMes        As String   'Mês.
    Dim lPosicao    As Integer  'Utilizado com InStr.

    '1. Verifica opção.
    If Not gfPreenchido(pOpcao) Then
        Call gpGraLog(2, "FTE0064 - Opção inválida (formata data - opção não preenchida)." & vbLf & _
                         vbLf & "O sistema será encerrado.")
        Exit Function 'End   ***MOSILVA Verificar
    End If
    If Not IsNumeric(pOpcao) Then
        Call gpGraLog(2, "FTE0064 - Opção inválida (formata data - opção não numérica)." & vbLf & vbLf & _
                         "O sistema será encerrado.")
        Exit Function 'End   ***MOSILVA Verificar
    End If
    Select Case pOpcao
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
            'Nada a fazer - opções válidas.
        Case Else
            Call gpGraLog(2, "FTE0064 - Opção inválida (formata data - opção desconhecida: " & pOpcao & _
                             ")." & vbLf & vbLf & "O sistema será encerrado.")
            Exit Function 'End   ***MOSILVA Verificar
    End Select

    '2. Processa formatação de data.
    Select Case pOpcao
        Case 1
            If gfPreenchido(pData) Then
                If Len(pData) = 10 Then
                    gfForData = Mid$(pData, 7, 4) & Mid$(pData, 4, 2) & Mid$(pData, 1, 2)
                Else
                    gfForData = "00000000"
                End If
            Else
                gfForData = "00000000"
            End If
        Case 2
            If gfPreenchido(pData) Then
                lAuxData = Format$(pData, "00000000")
            Else
                lAuxData = "00000000"
            End If
            gfForData = Mid$(lAuxData, 7, 2) & "/" & Mid$(lAuxData, 5, 2) & "/" & Mid$(lAuxData, 1, 4)
        Case 3
            gfForData = Format$(Now, "yyyymmddhhnnss")
        Case 4
            If gfPreenchido(pData) Then
                lAuxData = Format$(pData, "00000000000000")
            Else
                lAuxData = "00000000000000"
            End If
            gfForData = Mid$(lAuxData, 7, 2) & "/" & _
                        Mid$(lAuxData, 5, 2) & "/" & _
                        Mid$(lAuxData, 1, 4) & " " & _
                        Mid$(lAuxData, 9, 2) & ":" & _
                        Mid$(lAuxData, 11, 2) & ":" & _
                        Mid$(lAuxData, 13, 4)
        Case 5
            lAuxData = Format$(pData, "00000000")
            Select Case pTipArquivo
                Case "1"        'Access.
                    gfForData = "#" & Mid$(lAuxData, 5, 2) & "/" & Mid$(lAuxData, 7, 2) & "/" & _
                                Mid$(lAuxData, 1, 4) & "#"
                Case "2"        'SQL-Server
                    gfForData = "'" & Mid$(lAuxData, 5, 2) & "/" & Mid$(lAuxData, 7, 2) & "/" & _
                                Mid$(lAuxData, 1, 4) & "'"
            End Select
        Case 6
            Select Case pTipArquivo
                Case "1"        'Access.
                    gfForData = "#" & Mid$(pData, 4, 2) & "/" & Mid$(pData, 1, 2) & "/" & _
                                Mid$(pData, 7, 4) & "#"
                Case "2"        'SQL-Server
                    gfForData = "'" & Mid$(pData, 4, 2) & "/" & Mid$(pData, 1, 2) & "/" & _
                                Mid$(pData, 7, 4) & "'"
            End Select
        Case 7
            If gfPreenchido(pData) Then
                lAuxData = Format$(pData, "000000")
            Else
                lAuxData = "000000"
            End If
            If Left$(lAuxData, 2) > 50 Then 'De 1951 a 1999.
                gfForData = Mid$(lAuxData, 5, 2) & "/" & _
                            Mid$(lAuxData, 3, 2) & "/19" & _
                            Left$(lAuxData, 2)
            Else                            'De 2000 a 2050.
                gfForData = Mid$(lAuxData, 5, 2) & "/" & _
                            Mid$(lAuxData, 3, 2) & "/20" & _
                            Left$(lAuxData, 2)
            End If
        Case 8
            If gfPreenchido(pData) Then
                lAuxData = Format$(pData, "0000")
            Else
                lAuxData = "0000"
            End If
            If Left$(lAuxData, 2) = "00" Then
                If Right$(lAuxData, 2) > 50 Then 'De 1951 a 1999.
                    gfForData = "19" & Right$(lAuxData, 2)
                Else                            'De 2000 a 2050.
                    gfForData = "20" & Right$(lAuxData, 2)
                End If
            Else
                gfForData = lAuxData
            End If
        Case 9
            If IsMissing(pData) Then
                gfForData = ""
            Else
                If Not gfPreenchido(pData) Then
                    gfForData = ""
                Else
                    gfForData = lAuxData
                    If IsNumeric(pData) Then
                        Select Case Len(Trim(pData))
                            Case 6      'Provável digitação: ddmmaa.
                                lDia = Mid$(pData, 1, 2)
                                lMes = Mid$(pData, 3, 2)
                                lAno = Mid$(pData, 5, 2)
                                If Val(lAno) > 50 Then  'De 1951 a 1999.
                                    lAno = "19" & lAno
                                Else                    'De 2000 a 2050.
                                    lAno = "20" & lAno
                                End If
                                gfForData = lDia & "/" & lMes & "/" & lAno
                            Case 8      'Provável digitação: ddmmaaaa.
                                gfForData = Mid$(pData, 1, 2) & "/" & _
                                            Mid$(pData, 3, 2) & "/" & _
                                            Mid$(pData, 5, 4)
                            Case Else
                                gfForData = pData
                        End Select
                    Else
                        lAuxData = Trim$(pData)
                        'Dia.
                        lPosicao = InStr(lAuxData, "/")
                        Select Case lPosicao
                            Case 2
                                lDia = "0" & Mid$(lAuxData, 1, 1)
                            Case 3
                                lDia = Mid$(lAuxData, 1, 2)
                            Case Else
                                Exit Function
                        End Select
                        If Not IsNumeric(lDia) Then
                            Exit Function
                        End If
                        If Len(lAuxData) = lPosicao Then
                            Exit Function
                        End If
                        lAuxData = Mid$(lAuxData, lPosicao + 1)
                        'Mês.
                        lPosicao = InStr(lAuxData, "/")
                        Select Case lPosicao
                            Case 2
                                lMes = "0" & Mid$(lAuxData, 1, 1)
                            Case 3
                                lMes = Mid$(lAuxData, 1, 2)
                            Case Else
                                Exit Function
                        End Select
                        If Not IsNumeric(lMes) Then
                            Exit Function
                        End If
                        If Len(lAuxData) = lPosicao Then
                            Exit Function
                        End If
                        'Ano.
                        lAno = Mid$(lAuxData, lPosicao + 1)
                        If Not IsNumeric(lAno) Then
                            Exit Function
                        End If
                        If Len(lAno) = 1 Then
                            lAno = "0" & lAno
                        End If
                        If Len(lAno) = 2 Then
                            If Val(lAno) > 50 Then  'De 1951 a 1999.
                                lAno = "19" & lAno
                            Else                    'De 2000 a 2050.
                                lAno = "20" & lAno
                            End If
                        End If
                        If Len(lAno) = 3 Then
                            lAno = "0" & lAno
                        End If
                        If Len(lAno) > 4 Then
                            lAno = Left$(lAno, 4)
                        End If
                        gfForData = lDia & "/" & lMes & "/" & lAno
                    End If
                    If Not IsDate(gfForData) Then
                        gfForData = ""
                    End If
                End If
            End If
        Case 10
            If gfPreenchido(pData) Then
                lAuxData = Format$(pData, "000000")
                If Left$(lAuxData, 2) > 50 Then 'De 1951 a 1999.
                    gfForData = "19" & lAuxData
                Else                            'De 2000 a 2050.
                    gfForData = "20" & lAuxData
                End If
            Else
                lAuxData = "00000000"
            End If
        Case 11
            gfForData = Format$(Now, "yyyymmdd")
        Case 12
            gfForData = ""
            If gfPreenchido(pData) And CDblx(pData) <> 0 Then
                lAuxData = Format$(pData, "00000000")
                gfForData = Mid$(lAuxData, 7, 2) & "/" & Mid$(lAuxData, 5, 2) & "/" & Mid$(lAuxData, 1, 4)
            End If
            
    End Select
End Function
Public Function gfForCurrency(pValor As Variant, Optional pPreencheZero As Boolean = True) As String
Dim Aux As String
On Error GoTo MostraErro

    If Not IsNumeric(pValor) Then
        If pPreencheZero Then
            gfForCurrency = "0,00"
        End If
    ElseIf pValor = 0 Then
        If pPreencheZero Then
            gfForCurrency = "0,00"
        End If
    Else
        Aux = Replace(pValor, ".", "")
        gfForCurrency = Format(Aux, "##,###,###,###,###,##0.00")
    End If
    
    Exit Function
MostraErro:
gpGraLog 1, "UTI001- Erro no gfForCurrency " & Err.Number & " - " & Err.Description
    
End Function


Public Function CDblx(ByRef Valor As Variant) As Double
    If IsNumeric(Valor) Then
        Valor = Valor
    Else
        Valor = 0
    End If
    CDblx = Valor
End Function


Public Function gfForValor(ByVal pOpcao As Byte, ByVal pCampo As String, Optional pPadrao As String) _
                           As String
    'Função: formata valor.

    'Parâmetros...pOpcao....1 = troca vírgula por ponto e tira ponto, para ser utilizado no comando SQL.
    '                       2 = formata no padrão standard e coloca brancos à esquerda para ser utilizado
    '                           com os formulários.
    '                       3 = coloca brancos à esquerda para ser utilizado com os formulários.
    '             pCampo....campo a ser formatado.
    '             pPadrao...para opção 1, informa padrão a ser utilizado caso campo não esteja preenchido.
    '                       para opção 2 e 3, informa tamanho do campo (se omitido, assume 15 posições).

    Dim lAuxiliar   As String   'Auxiliar.
    Dim li          As Byte     'Utilizado com For...Next.
    Dim lTamanho    As Byte     'Tamanho.
    Dim lTexto      As String   'Auxiliar para tirar ponto.

    '1. Verifica opção.
    If Not gfPreenchido(pOpcao) Then
        Call gpGraLog(2, "FTE0065 - Opção inválida (formata valor - opção não preenchida)." & vbLf & _
                         vbLf & "O sistema será encerrado.")
        Exit Function 'End   ***MOSILVA Verificar
    End If
    If Not IsNumeric(pOpcao) Then
        Call gpGraLog(2, "FTE0065 - Opção inválida (formata valor - opção não numérica)." & vbLf & _
                         vbLf & "O sistema será encerrado.")
        Exit Function 'End   ***MOSILVA Verificar
    End If
    If pOpcao <> 1 And pOpcao <> 2 And pOpcao <> 3 Then
        Call gpGraLog(2, "FTE0065 - Opção inválida (formata valor - opção diferente de 1, 2 e 3: " & _
                         pOpcao & ")." & vbLf & vbLf & "O sistema será encerrado.")
        Exit Function 'End   ***MOSILVA Verificar
    End If

    '2. Processa formatação de data.
    Select Case pOpcao
        Case 1
            If gfPreenchido(pCampo) Then
                If IsNumeric(pCampo) Then
                    lAuxiliar = Trim$(pCampo)
                    lTexto = ""
                    For li = 1 To Len(lAuxiliar)
                        Select Case Mid(lAuxiliar, li, 1)
                            Case ","
                                lTexto = lTexto & "."
                            Case "."
                                'lTexto = lTexto
                            Case Else
                                lTexto = lTexto & Mid(lAuxiliar, li, 1)
                        End Select
                    Next li
                    If lTexto = "" Then
                        lTexto = "0"
                    End If
                    gfForValor = lTexto
                    Exit Function
                End If
            End If
            If IsMissing(pPadrao) Then
                gfForValor = "0"
            Else
                If gfPreenchido(pPadrao) Then
                    gfForValor = pPadrao
                Else
                    gfForValor = "0"
                End If
            End If
        Case 2
            If gfPreenchido(pCampo) Then
                If IsNumeric(pCampo) Then
                    lAuxiliar = Trim$(Format$(pCampo, "standard"))
                Else
                    lAuxiliar = "0,00"
                End If
            Else
                lAuxiliar = "0,00"
            End If
            If IsMissing(pPadrao) Then
                lTamanho = 15
            Else
                If Not IsNumeric(pPadrao) Then
                    lTamanho = 15
                Else
                    lTamanho = pPadrao
                End If
            End If
            If Len(lAuxiliar) > lTamanho Then
                gfForValor = lAuxiliar
            Else
                gfForValor = Space(lTamanho - Len(lAuxiliar)) & lAuxiliar
            End If
        Case 3
            If IsMissing(pPadrao) Then
                lTamanho = 5
            Else
                If Not IsNumeric(pPadrao) Then
                    lTamanho = 5
                Else
                    lTamanho = pPadrao
                End If
            End If
            If gfPreenchido(pCampo) Then
                lAuxiliar = Trim$(pCampo)
                If Len(lAuxiliar) > lTamanho Then
                    gfForValor = lAuxiliar
                Else
                    gfForValor = Space(lTamanho - Len(lAuxiliar)) & lAuxiliar
                End If
            Else
                gfForValor = Space(lTamanho)
            End If
    End Select
End Function
Public Function gfGraItem(ByRef pBasDados As ADODB.Connection, _
                          ByVal pNosNumero As String, _
                          ByVal pNumPedido As String, _
                          ByVal pNumItem As String, _
                          ByRef pMensagem As String, _
                          pTipEmissao As e_TipEmissao, _
                          Optional ByRef pobjstcPedLoc As stcA46V708B, _
                          Optional ByVal pConCodSeguradora As Integer, _
                          Optional ByVal pConApolice As String, _
                          Optional ByVal pConItem As String, _
                          Optional ByVal pIndSinistro As Byte, _
                          Optional pCod_Ramo As Integer, _
                          Optional pCodProduto As Integer) As Boolean   '1000347 - Inclusão de informações da congênere em renovações
    'Função: grava registro do item - TAB_PED_LOC.
    'Parâmetros de entrada...pBasDados....Base de dados.
    '                        pNosNumero...Nosso número.
    '                        pNumPediro...Número do pedido.
    '                        pNumItem.....Número do item.
    'Parâmetro de saída......pMensagem....Mensagem.

Dim lObjPedLoc          As clsA46V708A      'Objeto da classe clsA46V708A (TAB_PED_LOC)
Dim lobjstcPedLoc       As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)

Dim lobjCobert          As clsA46V704A      'Objeto da classe clsA46V704A (TAB_PED_COBERT)
Dim lObjCobert_Temp     As clsA46V704A
Dim lobjstcCobert       As stcA46V704B      'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)

Dim lstcPedFator        As stcA46V720B
Dim lObjPedFator        As clsA46V720A

Dim lPropTabela         As String
Dim lNomeTabImport      As String
    
    gfGraItem = False

    Call gfObtSYASMULTINI(lPropTabela)
    lNomeTabImport = ""

    '1. Instancio as classes
    Set lObjPedLoc = New clsA46V708A
    Set lobjstcPedLoc = New stcA46V708B
    Set lObjPedLoc.ConexaoBD = pBasDados
    
    If Not IsNumeric(pCod_Ramo) Then pCod_Ramo = 0

'4. Instancio os objetos do Tab_Ped_Loc
    Set lObjPedLoc = New clsA46V708A
    Set lobjstcPedLoc = New stcA46V708B
    Set lObjPedLoc.ConexaoBD = pBasDados
    
    '5. Verifica se tem registro padrão do item.
    lobjstcPedLoc.Nosso_Numero = "PADRAO" & gNomUsuario6 & pCod_Ramo
    If Not lObjPedLoc.mfCarregarColecao(lobjstcPedLoc, lPropTabela, "", True) Then
        pMensagem = lObjPedLoc.UltimoErro
        GoTo SAIDA
    End If
    
    '6. Se não tem registro padrão - então inclui.
    If lObjPedLoc.Colecao.Count = 0 Then
        Call mfVerRegPadrao(pMensagem)
    Else
        Set lobjstcPedLoc = lObjPedLoc.Colecao.Item(1)
    End If
    
    '2. Carrego as informações
    With lobjstcPedLoc
        .Nosso_Numero = pNosNumero
        .Num_Ped = pNumPedido
        .Num_Item = pNumItem
        .Cod_Ramo = pCod_Ramo
        .Tip_Emissao = pTipEmissao
        '1000347 - Inclusão de informações da congênere em renovações
        If pTipEmissao = e_TipEmissao_RenCongenere Then
            .Ren_Con_Seguradora = pConCodSeguradora
            .Ren_Con_Apolice = pConApolice
            .Ren_Con_Item = pConItem
            .Ind_Sinistro = pIndSinistro
        End If
        .Cod_Plano = 1
        .Cod_Produto = 1
    End With

    '3. Chamo a rotina de incluir
    If Not lObjPedLoc.mfIncluir(lobjstcPedLoc, lPropTabela, lNomeTabImport) Then
        pMensagem = lObjPedLoc.UltimoErro
        GoTo SAIDA
    End If

    
    'Inspeção
    Dim lobjInsp As clsA46V717A
    Set lobjInsp = New clsA46V717A
    Set lobjInsp.ConexaoBD = pBasDados
    lobjstcPedLoc.objstcPedInsp.Nosso_Numero = pNosNumero
    lobjstcPedLoc.objstcPedInsp.Num_Ped = pNumPedido
    lobjstcPedLoc.objstcPedInsp.Num_Item = pNumItem
    lobjstcPedLoc.objstcPedInsp.Cod_Ramo = pCod_Ramo
    lobjstcPedLoc.objstcPedInsp.Cod_Org = lobjstcPedLoc.Cod_Org
    If Not lobjInsp.mfIncluir(lobjstcPedLoc.objstcPedInsp, lPropTabela, lNomeTabImport) Then
        pMensagem = lobjInsp.UltimoErro
        GoTo SAIDA
    End If

    '5. Verifica se tem registro padrão
    Set lobjCobert = New clsA46V704A
    Set lobjstcCobert = New stcA46V704B
    Set lobjCobert.ConexaoBD = pBasDados

    '6. Se não tem registro padrão - então inclui.
    If lobjstcPedLoc.ColecaoPedCobert.Count = 0 Then
        Set lobjstcCobert = New stcA46V704B
        With lobjstcCobert
            .Nosso_Numero = pNosNumero
            .Num_Ped = pNumPedido
            .Num_Item = pNumItem
            .Cod_Ramo = pCod_Ramo
            .Cod_Org = lobjstcPedLoc.Cod_Org
            .Cod_Cobert = 1
            .Tab_Cobert = 1
            Set .ObjPai = lobjstcPedLoc
        End With
        If Not lobjCobert.mfIncluir(lobjstcCobert, lPropTabela, lNomeTabImport) Then
            pMensagem = lobjCobert.UltimoErro
            GoTo SAIDA
        End If
        Set lobjstcPedLoc.ColecaoPedCobert = New Collection
        If Not lobjCobert.Colecao Is Nothing Then  'incluído juntamente com Thais - 21/10/2011
            Set lobjstcPedLoc.ColecaoPedCobert = lobjCobert.Colecao
        End If
    Else
        For Each lobjstcCobert In lobjstcPedLoc.ColecaoPedCobert
            With lobjstcCobert
                .Nosso_Numero = pNosNumero
                .Num_Ped = pNumPedido
                .Num_Item = pNumItem
                .Cod_Ramo = pCod_Ramo
                .Cod_Org = lobjstcPedLoc.Cod_Org
                Set .ObjPai = lobjstcPedLoc
            End With
            If Not lobjCobert.mfIncluir(lobjstcCobert, lPropTabela, lNomeTabImport) Then
                pMensagem = lobjCobert.UltimoErro
                GoTo SAIDA
            End If
        Next lobjstcCobert
    End If
    
    'Fatores de Proteção
    Set lObjPedFator = New clsA46V720A
    Set lObjPedFator.ConexaoBD = pBasDados
    For Each lstcPedFator In lobjstcPedLoc.ColecaoPedFator
        lstcPedFator.Nosso_Numero = pNosNumero
        lstcPedFator.Num_Ped = pNumPedido
        lstcPedFator.Num_Item = pNumItem
        lstcPedFator.Cod_Ramo = pCod_Ramo
        lstcPedFator.Cod_Org = lobjstcPedLoc.Cod_Org
        Set lstcPedFator.ObjPai = lobjstcPedLoc
        If Not lObjPedFator.mfIncluir(lstcPedFator, lPropTabela, lNomeTabImport) Then
            pMensagem = lObjPedFator.UltimoErro
            GoTo SAIDA
        End If
    Next
    
    If lobjstcPedLoc.Cod_Ramo = e_CodRamo_Residencial And _
       ((pCodProduto >= 1022 And pCodProduto <= 1033) Or (pCodProduto >= 1038)) Then  '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 8
        Set lstcPedFator = New stcA46V720B
        With lstcPedFator
            .Nosso_Numero = pNosNumero
            .Num_Ped = pNumPedido
            .Num_Item = pNumItem
            .Cod_Ramo = pCod_Ramo
            .Cod_Org = lobjstcPedLoc.Cod_Org
            .Cod_Fator = 7
            .Desc_Fator = "DESCONTO ESPECIAL"
            Set .ObjPai = lobjstcPedLoc
        End With
        If Not lObjPedFator.mfIncluir(lstcPedFator, lPropTabela, lNomeTabImport) Then
            pMensagem = lObjPedFator.UltimoErro
            GoTo SAIDA
        End If
    End If
    
    Dim lstcPedLocBem As stcA46V715B
    Dim lobjPedLocBem As clsA46V715A
    Set lobjPedLocBem = New clsA46V715A
    Set lobjPedLocBem.ConexaoBD = pBasDados
    For Each lstcPedLocBem In lobjstcPedLoc.ColecaoPedLocBem
        lstcPedLocBem.Nosso_Numero = pNosNumero
        lstcPedLocBem.Num_Ped = pNumPedido
        lstcPedLocBem.Num_Item = pNumItem
        lstcPedLocBem.Cod_Ramo = pCod_Ramo
        lstcPedLocBem.Cod_Org = lobjstcPedLoc.Cod_Org
        Set lstcPedLocBem.ObjPai = lobjstcPedLoc
        If Not lobjPedLocBem.mfIncluir(lstcPedLocBem, lPropTabela, lNomeTabImport) Then
            pMensagem = lobjPedLocBem.UltimoErro
            GoTo SAIDA
        End If
    Next
    
    Set pobjstcPedLoc = lobjstcPedLoc
    
    gfGraItem = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
    
FechaObj:
    '4. Destruo os objetos
    Set lObjPedLoc = Nothing
    Set lobjstcPedLoc = Nothing
    Set lobjstcCobert = Nothing
    Set lobjCobert = Nothing
    Set lstcPedFator = Nothing
    Set lObjPedFator = Nothing
    Set lobjInsp = Nothing
    Return
End Function
Public Function gfNomCorretor(ByVal pCodCorretor As String, ByRef pNomCorretor As String, _
                              ByRef pMensagem As String) As Boolean
    'Função: obtém nome do corretor.

    'Parâmetro de entrada...pCodCorretor...Código do corretor.
    'Parâmetros de saída....pNomCorretor...Nome do corretor.
    '                       pMensagem......Mensagem.

    Dim bdP0044700  As ADODB.Connection 'Arquivo - P0044700.
    Dim lSelect     As String           'Comando SELECT.
    Dim rsCORR      As ADODB.Recordset  'YTAB_CORR: tabela.

    gfNomCorretor = False

    '1. Verifica preenchimento.
    If Not gfPreenchido(pCodCorretor) Then
        pNomCorretor = "Código do corretor não preenchido."
        gfNomCorretor = True
        Exit Function
    End If
    If Val(pCodCorretor) = 0 Then
        pNomCorretor = "Código do corretor zerado."
        gfNomCorretor = True
        Exit Function
    End If

    '2. Obtém nome do corretor.
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    lSelect = "SELECT NOM_ABRV_CORR FROM YTAB_CORR WHERE COD_CORR = " & pCodCorretor
    If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
        Call gpFecha2(bdP0044700)
        Exit Function
    End If
    If rsCORR.EOF = True Then
        Call gpFecha3(rsCORR, bdP0044700)
        pMensagem = "FGE0003 - " & gPrefixo & "Corretor inválido (código de corretor não localizado " & _
                    "- " & pCodCorretor & ")."
        Exit Function
    End If
    pNomCorretor = Trim$(rsCORR!NOM_ABRV_CORR)
    Call gpFecha3(rsCORR, bdP0044700)

    gfNomCorretor = True
End Function
Public Function gfNomSeguradora(ByRef pCodSeguradora As String, ByRef pNomSeguradora As String, _
                                ByRef pMensagem As String, Optional ByRef pNomAbrev As String) As Boolean
    'Função: obtém nome da seguradora

    'Parâmetro de entrada...pCodSeguradora...Código da seguradora.
    'Parâmetros de saída....pNomSeguradora...Nome da seguradora.
    '                       pMensagem........Mensagem.

    Dim bdP0044700  As ADODB.Connection 'Arquivo - P0044700.
    Dim lSelect     As String           'Comando SELECT.
    Dim rsP0044700  As ADODB.Recordset  'P0044700: tabela.

    gfNomSeguradora = False

    '1. Verifica preenchimento.
    If Not gfPreenchido(pCodSeguradora) Then
        pMensagem = "FTE0066 - " & gPrefixo & _
                    "Erro ao obter nome da seguradora (código da seguradora não preenchido)."
        Exit Function
    End If
    If Not IsNumeric(pCodSeguradora) Then
        pMensagem = "FTE0066 - " & gPrefixo & _
                    "Erro ao obter nome da seguradora (código da seguradora não numérico)."
        Exit Function
    End If
    If Val(pCodSeguradora) = 0 Then
        pMensagem = "FTE0066 - " & gPrefixo & _
                    "Erro ao obter nome da seguradora (código da seguradora zerado)."
        Exit Function
    End If

    '2. Obtém nome da seguradora.
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    lSelect = "SELECT Seguradora , aSeguradora FROM YTAB_SEGURADORA WHERE CSeguradora = '" & Trim(pCodSeguradora) & "'"
    If gfObtRegistro(bdP0044700, lSelect, rsP0044700, pMensagem) = False Then
        Call gpFecha2(bdP0044700)
        Exit Function
    End If
    If rsP0044700.EOF = True Then
        pMensagem = "FTE0066 - " & gPrefixo & _
                    "Erro ao obter nome da seguradora (código da seguradora não localizado - " & _
                    pCodSeguradora & ")."
        Exit Function
    End If
    pNomSeguradora = rsP0044700!Seguradora
    pNomAbrev = rsP0044700!aSeguradora
    Call gpFecha3(rsP0044700, bdP0044700)

    gfNomSeguradora = True
End Function
Public Function gfVerCorrPermitido(ByVal pCorrInstalacao As Long, ByVal pCorrSelecionado As Long, pRegistro As String) As Boolean
    Dim lCorrInstacao       As Long     'Código do corretor - instalação.
    Dim lCorrSelecionado    As Long     'Código do corretor - selecionado.
    Dim lEncArquivo         As Boolean  'Identifica se a verificação foi feita através do arquivo, caso
                                        'contrário deverá ser buscado do código
    Dim lEncontrou          As Boolean  'Identifica se foi encontrato o [grupos] no arquivo.
    Dim lEndP366CORR        As String   'Endereço do arquivo P366CORR.EXE.
    Dim lNumP366CORR        As Integer  'Utilizado com FreeFile para o arquivo P366CORR.EXE.
    Dim lRegistro           As String   'Registro.

    'A comparação no arquivo é feita com o número do corretor sem os dois últimos dígitos.
    'Os dois últimos dígitos variam mas dentro de um mesmo grupo de corretor.
    'No arquivo P366CORR.EXE não tem o número completo do corretor, está sem os dois últimos dígitos.
    lCorrInstacao = Left(Format(pCorrInstalacao, "0000000"), 5)
    lCorrSelecionado = Left(Format(pCorrSelecionado, "0000000"), 5)
    lEndP366CORR = gAppPath & "P366CORR.EXE"
    If Not gfPreenchido(Dir(lEndP366CORR)) Then
        lEncArquivo = False
    Else
        lEncArquivo = True
    End If
    If lEncArquivo = True Then
        lNumP366CORR = FreeFile
        On Error Resume Next
        Open lEndP366CORR For Input Access Read As lNumP366CORR
        If Err <> 0 Then
            lEncArquivo = False
        End If
        If lEncArquivo = True Then
            On Error GoTo 0
            '4. Trata registros do arquivo de inicialização.
            lEncontrou = False
            lEncArquivo = False
            Do While Not EOF(lNumP366CORR)
                Line Input #lNumP366CORR, lRegistro
                If gfPreenchido(lRegistro) Then
                    lRegistro = Trim$(lRegistro)
                    If UCase(lRegistro) = "[GRUPOS]" Then
                        lEncontrou = True
                    End If
                    If UCase(lRegistro) = "[SOMENTE MULTI]" Then
                        Exit Do
                    End If
                    If lEncontrou = True Then
                        If InStr(1, lRegistro, ";" & lCorrInstacao & ";") > 0 Then
                            pRegistro = lRegistro
                            If InStr(1, lRegistro, ";" & lCorrSelecionado & ";") > 0 Then
                                gfVerCorrPermitido = True
                            Else
                                gfVerCorrPermitido = False
                            End If
                            lEncArquivo = True
                            Exit Do
                        End If
                        If lCorrInstacao <> lCorrSelecionado Then
                            lEncArquivo = False
                            gfVerCorrPermitido = False
                        Else
                            lEncArquivo = True
                            gfVerCorrPermitido = True
                            Exit Do
                        End If
                    End If
                End If
            Loop
            Close lNumP366CORR
        End If
    End If
    If lEncArquivo = False Then
        Select Case lCorrInstacao
                Case 809, 2387
                    Select Case lCorrSelecionado
                        Case 809, 2387
                            'Seleção OK
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 376, 3813
                    Select Case lCorrSelecionado
                        Case 376, 3813
                            gfVerCorrPermitido = True
                        Case Else
                                'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2234, 904
                    Select Case lCorrSelecionado
                        Case 2234, 904
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 183, 4148
                    Select Case lCorrSelecionado
                        Case 183, 4148
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 77, 4246, 4708, 4709, 4710
                    Select Case lCorrSelecionado
                        Case 77, 4246, 4708, 4709, 4710
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4618, 1102
                    Select Case lCorrSelecionado
                        Case 4618, 1102
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 843, 1172, 845, 377
                    Select Case lCorrSelecionado
                        Case 843, 1172, 845, 377
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6449, 1749
                    Select Case lCorrSelecionado
                        Case 6449, 1749
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1889, 495
                    Select Case lCorrSelecionado
                        Case 1889, 495
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6106, 4978, 4800, 4792, 4846, 2769, 3223
                    Select Case lCorrSelecionado
                        Case 6106, 4978, 4800, 4792, 4846, 2769, 3223
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3272, 34, 3335, 393, 3070, 73
                    Select Case lCorrSelecionado
                        Case 3272, 34, 3335, 393, 3070, 73
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2577, 2101
                    Select Case lCorrSelecionado
                        Case 2577, 2101
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2422, 6145
                    Select Case lCorrSelecionado
                        Case 2422, 6145
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 172, 3407
                    Select Case lCorrSelecionado
                        Case 172, 3407
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 855, 1170, 1171, 2582, 6659
                    Select Case lCorrSelecionado
                        Case 855, 1170, 1171, 2582, 6659
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 60, 3301, 3302
                    Select Case lCorrSelecionado
                        Case 60, 3301, 3302
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1184, 709
                    Select Case lCorrSelecionado
                        Case 1184, 709
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 389, 4252
                    Select Case lCorrSelecionado
                        Case 389, 4252
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2505, 4162
                    Select Case lCorrSelecionado
                        Case 2505, 4162
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 96, 534, 1454
                    Select Case lCorrSelecionado
                        Case 96, 534, 1454
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 822, 406
                    Select Case lCorrSelecionado
                        Case 822, 406
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 530, 2171
                    Select Case lCorrSelecionado
                        Case 530, 2171
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6369, 2994, 6544
                    Select Case lCorrSelecionado
                        Case 6369, 2994, 6544
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3668, 1496
                    Select Case lCorrSelecionado
                        Case 3668, 1496
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4675, 6446
                    Select Case lCorrSelecionado
                        Case 4675, 6446
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1491, 3821
                    Select Case lCorrSelecionado
                        Case 1491, 3821
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3272, 34, 4726, 4431, 3438, 3302, 73, 3456, 4723, 3459, 3070, 3392
                    Select Case lCorrSelecionado
                        Case 3272, 34, 4726, 4431, 3438, 3302, 73, 3456, 4723, 3459, 3070, 3392
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2360, 2888
                    Select Case lCorrSelecionado
                        Case 2360, 2888
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4565, 4669
                    Select Case lCorrSelecionado
                        Case 4565, 4669
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 748, 1069
                    Select Case lCorrSelecionado
                        Case 748, 1069
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1155, 2501, 1255, 6205
                    Select Case lCorrSelecionado
                        Case 1155, 2501, 1255, 6205
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3494, 3189, 3583, 3509, 356, 2956
                    Select Case lCorrSelecionado
                        Case 3494, 3189, 3583, 3509, 356, 2956
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3075, 3367
                    Select Case lCorrSelecionado
                        Case 3075, 3367
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1863, 6238, 4621, 3311, 6237, 6244, 6242, 6248, 6231, 1860, 6191, 6243, 4622, 6658
                    Select Case lCorrSelecionado
                        Case 1863, 6238, 4621, 3311, 6237, 6244, 6242, 6248, 6231, 1860, 6191, 6243, 4622, 6658
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 322, 6361, 263, 2311, 3865, 4866, 6439
                    Select Case lCorrSelecionado
                        Case 322, 6361, 263, 2311, 3865, 4866, 6439
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4267, 6678
                    Select Case lCorrSelecionado
                        Case 4267, 6678
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 214, 6125
                    Select Case lCorrSelecionado
                        Case 214, 6125
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2841, 2842
                    Select Case lCorrSelecionado
                        Case 2841, 2842
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6593, 1465
                    Select Case lCorrSelecionado
                        Case 6593, 1465
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2453, 4485
                    Select Case lCorrSelecionado
                        Case 2453, 4485
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4503, 4353, 3416
                    Select Case lCorrSelecionado
                        Case 4503, 4353, 3416
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3111, 1614
                    Select Case lCorrSelecionado
                        Case 3111, 1614
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2355, 2344
                    Select Case lCorrSelecionado
                        Case 2355, 2344
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1116, 1901, 1321, 1890
                    Select Case lCorrSelecionado
                        Case 1116, 1901, 1321, 1890
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2451, 1148
                    Select Case lCorrSelecionado
                        Case 2451, 1148
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3043, 6515
                    Select Case lCorrSelecionado
                        Case 3043, 6515
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2142, 133
                    Select Case lCorrSelecionado
                        Case 2142, 133
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3775, 3098, 1807, 1769, 3082, 6621, 768, 4213, 3081
                    Select Case lCorrSelecionado
                        Case 3775, 3098, 1807, 1769, 3082, 6621, 768, 4213, 3081
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6435, 1688
                    Select Case lCorrSelecionado
                        Case 6435, 1688
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3379, 6625
                    Select Case lCorrSelecionado
                        Case 3379, 6625
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2145, 871
                    Select Case lCorrSelecionado
                        Case 2145, 871
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2369, 162, 469, 483
                    Select Case lCorrSelecionado
                        Case 2369, 162, 469, 483
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2912, 2911
                    Select Case lCorrSelecionado
                        Case 2912, 2911
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2479, 2889, 4352
                    Select Case lCorrSelecionado
                        Case 2479, 2889, 4352
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3610, 6592, 4561, 3642, 3271, 4245, 3742, 1640, 1923, 2859, 3667
                    Select Case lCorrSelecionado
                        Case 3610, 6592, 4561, 3642, 3271, 4245, 3742, 1640, 1923, 2859, 3667
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3544, 3543, 3547, 3541, 3542, 3545, 3546, 3548, 2740
                    Select Case lCorrSelecionado
                        Case 3544, 3543, 3547, 3541, 3542, 3545, 3546, 3548, 2740
                                gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2734, 4886
                    Select Case lCorrSelecionado
                        Case 2734, 4886
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3457, 3458, 3459, 3456
                    Select Case lCorrSelecionado
                        Case 3457, 3458, 3459, 3456
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4573, 6008
                    Select Case lCorrSelecionado
                        Case 4573, 6008
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1060, 2702
                    Select Case lCorrSelecionado
                        Case 1060, 2702
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 58, 6487, 4037
                    Select Case lCorrSelecionado
                        Case 58, 6487, 4037
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3982, 3844, 3119, 6080, 3149, 6378
                    Select Case lCorrSelecionado
                        Case 3982, 3844, 3119, 6080, 3149, 6378
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2909, 2182
                    Select Case lCorrSelecionado
                        Case 2909, 2182
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 62, 6616
                    Select Case lCorrSelecionado
                        Case 62, 6616
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 2621, 179
                    Select Case lCorrSelecionado
                        Case 2621, 179
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3438, 3392
                    Select Case lCorrSelecionado
                        Case 3438, 3392
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3319, 6131
                    Select Case lCorrSelecionado
                        Case 3319, 6131
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 827, 337
                    Select Case lCorrSelecionado
                        Case 827, 337
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4481, 2147, 4589
                    Select Case lCorrSelecionado
                        Case 4481, 2147, 4589
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 80, 4086
                    Select Case lCorrSelecionado
                        Case 80, 4086
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6240, 6198, 6195, 3397, 6210, 6256, 6298, 6219, 6209, 6442, 6420, 6172, 6297, 6301, 6193, 6226, 6413, 6218, 1895, 6228, 1864, 6419, 6588, 6208, 6224, 1435, 6261, 6175, 6596, 6184
                    Select Case lCorrSelecionado
                        Case 6240, 6198, 6195, 3397, 6210, 6256, 6298, 6219, 6209, 6442, 6420, 6172, 6297, 6301, 6193, 6226, 6413, 6218, 1895, 6228, 1864, 6419, 6588, 6208, 6224, 1435, 6261, 6175, 6596, 6184
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 6128, 4700, 6217, 1771, 4661, 1872, 6372
                    Select Case lCorrSelecionado
                        Case 6128, 4700, 6217, 1771, 4661, 1872, 6372
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 1825, 4118, 4389, 6611, 6404, 4724, 4391, 1705, 4775, 6293, 4393, 4117, 4619, 4394, _
                     6606, 4463, 6294, 4556, 6580, 4396, 4576, 4916, 4398, 4457, 4482, 6603, 4458, 4399, _
                     6291, 1893, 4735, 6609, 4693, 1707, 6394, 6636, 1871, 1761, 6009, 4538, 6655, 6098, _
                     6406, 6351, 1828, 4405, 6255, 4640, 4407, 4408, 4702, 6352, 6643, 4483, 4410, 1823, _
                     1715, 6440, 1827, 4413, 4429, 4414, 4416, 1791, 6232, 6299, 6485, 4587, 6608, 4419, _
                     6626, 4448, 6418, 6019, 6182, 4518, 4690, 4633, 4818, 6190, 6513, 1911
                    Select Case lCorrSelecionado
                        Case 1825, 4118, 4389, 6611, 6404, 4724, 4391, 1705, 4775, 6293, 4393, 4117, _
                             4619, 4394, 6606, 4463, 6294, 4556, 6580, 4396, 4576, 4916, 4398, 4457, _
                             4482, 6603, 4458, 4399, 6291, 1893, 4735, 6609, 4693, 1707, 6394, 6636, _
                             1871, 1761, 6009, 4538, 6655, 6098, 6406, 6351, 1828, 4405, 6255, 4640, _
                             4407, 4408, 4702, 6352, 6643, 4483, 4410, 1823, 1715, 6440, 1827, 4413, _
                             4429, 4414, 4416, 1791, 6232, 6299, 6485, 4587, 6608, 4419, 6626, 4448, _
                             6418, 6019, 6182, 4518, 4690, 4633, 4818, 6190, 6513, 1911
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 4966, 4875, 4900, 4830, 4898
                    Select Case lCorrSelecionado
                        Case 4966, 4875, 4900, 4830, 4898
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case 3611, 3708, 1448, 1468, 1464, 1460, 3485, 3486, 2747, 3100, 3612, 3484, 3511, 3482, _
                     3483, 3585, 1449, 3502, 3892, 3531, 1474, 1484, 3569, 1451, 3117, 1485, 3481, 3636, _
                     3581, 3638, 3480, 3568, 3732, 3637, 3487, 2585, 3627, 3385, 3493, 3665, 3273, 3654, _
                     3499
                    Select Case lCorrSelecionado
                        Case 3611, 3708, 1448, 1468, 1464, 1460, 3485, 3486, 2747, 3100, 3612, 3484, _
                             3511, 3482, 3483, 3585, 1449, 3502, 3892, 3531, 1474, 1484, 3569, 1451, _
                             3117, 1485, 3481, 3636, 3581, 3638, 3480, 3568, 3732, 3637, 3487, 2585, _
                             3627, 3385, 3493, 3665, 3273, 3654, 3499
                            gfVerCorrPermitido = True
                        Case Else
                            'Erro
                            gfVerCorrPermitido = False
                    End Select
                Case Else
                    If lCorrInstacao <> lCorrSelecionado Then
                        gfVerCorrPermitido = False
                    Else
                        gfVerCorrPermitido = True
                    End If
        End Select
    End If
End Function

Public Function gfObtCorretor(ByVal pOpcao As Byte, ByVal pCorretor As String, _
                              ByRef pSaida As Variant, ByRef pMensagem As String) As Boolean
    'Função: obtém dados de corretores.

    'Parâmetros de entrada...pOpcao......1 = Código do corretor e 2 = Nome do corretor.
    '                        pCorretor...Código ou nome do corretor, de acordo com pOpcao.
    'Parâmetro de saída......pSaida......Array contendo resultados.
    '                                    (x, 1) Código do corretor.
    '                                    (x, 2) Nome do corretor.
    '                                    (x, 3) Nome abreviado do corretor.
    '                                    (x, 4) Num. Susep
    Dim bdP0044700  As ADODB.Connection 'Arquivo - P0044700.
    Dim li          As Double           'Contador.
    Dim lSelect     As String           'Comando SELECT.
    Dim rsCORR      As ADODB.Recordset  'Tabela: YTAB_CORR.

    gfObtCorretor = False

    '1. Verifica parâmetros.
    'a) Opção.
    If Not gfPreenchido(pOpcao) Then
        Call gpGraLog(2, "FTE0067 - Opção inválida (obter corretor - opção não preenchida)." & vbLf & _
                         vbLf & "O sistema será encerrado.")
        Exit Function 'End   ***MOSILVA Verificar
    End If
    If Not IsNumeric(pOpcao) Then
        Call gpGraLog(2, "FTE0067 - Opção inválida (obter corretor - opção não numérica)." & vbLf & _
                         vbLf & "O sistema será encerrado.")
        Exit Function 'End   ***MOSILVA Verificar
    End If
    If pOpcao <> 1 And pOpcao <> 2 Then
        Call gpGraLog(2, "FTE0067 - Opção inválida (obter corretor - opção diferente de 1 e 2)." & _
                         vbLf & vbLf & "O sistema será encerrado.")
        Exit Function 'End   ***MOSILVA Verificar
    End If
    'b) Corretor.
    If Not gfPreenchido(pCorretor) Then
        Call gpGraLog(2, "FTE0068 - Erro ao obter corretor (código do corretor não preenchido)." & _
                         vbLf & vbLf & "O sistema será encerrado.")
        Exit Function 'End   ***MOSILVA Verificar
    End If

    '2. Busca informação.
    Select Case pOpcao
        Case 1      'Código do corretor.
            If Not IsNumeric(pCorretor) Then
                Call gpGraLog(2, "FTE0068 - Erro ao obter corretor (código do corretor não numérico)." & _
                                 vbLf & vbLf & "O sistema será encerrado.")
                Exit Function 'End   ***MOSILVA Verificar
            End If
            lSelect = "SELECT COD_CORR, NOM_CORR, NOM_ABRV_CORR , num_susep FROM YTAB_CORR WHERE COD_CORR = " & _
                      pCorretor & " AND COD_ATIVO = 1"
            If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
                Exit Function
            End If
            If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
                Exit Function
            End If
            If rsCORR.RecordCount = 0 Then
                Call gpFecha3(rsCORR, bdP0044700)
                pMensagem = "FGE0003 - " & gPrefixo & "Corretor inválido (código de corretor não " & _
                            "localizado - " & pCorretor & ")."
                Exit Function
            End If
            ReDim pSaida(1, 4)
            pSaida(1, 1) = rsCORR!Cod_Corr
            pSaida(1, 2) = rsCORR!NOM_CORR
            pSaida(1, 3) = rsCORR!NOM_ABRV_CORR
            pSaida(1, 4) = rsCORR!NUM_SUSEP
            Call gpFecha3(rsCORR, bdP0044700)
            gfObtCorretor = True
            Exit Function
        Case 2      'Nome do corretor.
            lSelect = "SELECT COD_CORR, NOM_CORR, NOM_ABRV_CORR , Num_Susep FROM YTAB_CORR " & _
                      "WHERE (NOM_CORR LIKE '%" & pCorretor & "%' " & _
                      "OR     NOM_ABRV_CORR LIKE '%" & pCorretor & "%') " & _
                      "AND   COD_ATIVO = 1"
            If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
                Exit Function
            End If
            If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
                Exit Function
            End If
            If rsCORR.RecordCount = 0 Then
                Call gpFecha3(rsCORR, bdP0044700)
                pMensagem = "FGE0003 - Corretor inválido (corretor não localizado com nome ou parte " & _
                            "do nome = " & pCorretor & ")."
                Exit Function
            End If
            rsCORR.MoveLast
            ReDim Preserve pSaida(rsCORR.RecordCount, 4)
            li = 0
            rsCORR.MoveFirst
            While Not rsCORR.EOF
                li = li + 1
                pSaida(li, 1) = rsCORR!Cod_Corr
                pSaida(li, 2) = rsCORR!NOM_CORR
                pSaida(li, 3) = rsCORR!NOM_ABRV_CORR
                pSaida(li, 4) = rsCORR!NUM_SUSEP
                rsCORR.MoveNext
            Wend
            Call gpFecha3(rsCORR, bdP0044700)
            gfObtCorretor = True
            Exit Function
    End Select
End Function
Public Function gfObtUniProdutiva(ByVal pCodCorretor As String, ByRef pCodUniNegocios As Integer, _
                                  ByRef pCodProdutor As Integer, ByRef pMensagem As String, _
                                  Optional ByRef pTotalRegistros As Integer) As Boolean
    'Função: obtém unidade de negócios e produtor do corretor.

    'Parâmetro de entrada...pCodCorretor......Código do corretor.
    'Parâmetros de saída....pCodUniNegocios...Código da unidade de negócios.
    '                       pCodProdutor......Código do produtor.
    '                       pMensagem.........Mensagem.

    Dim bdP0044700      As ADODB.Connection 'Arquivo - P0044700.
    Dim li              As Double           'Contador.
    Dim lNomCorretor    As String           'Nome do corretor.
    Dim lSelect         As String           'Comando SELECT.
    Dim rsUNIDCORRPROD  As ADODB.Recordset  'Registro - YTAB_UNID_CORR_PROD.

    gfObtUniProdutiva = False

    '1. Inicializa saída.
    pCodUniNegocios = 0
    pCodProdutor = 0

    '2. Verifica parâmetro.
    If Not gfPreenchido(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade de negócios e produtor (código do corretor não " & _
                    "preenchido)."
        Exit Function
    End If

    '3. Busca informação.
    If Not IsNumeric(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade negócios e produtor (código do corretor não " & _
                    "numérico)."
        Exit Function
    End If
   
    lSelect = "SELECT A.COD_UNID, A.COD_PROD FROM YTAB_UNID_CORR_PROD A "
    lSelect = lSelect & " INNER JOIN YTAB_CORR B ON A.COD_CORR = B.COD_CORR "
    lSelect = lSelect & " WHERE A.COD_CORR = " & pCodCorretor
    'lSelect = lSelect & " AND   A.COD_UNID IN(1751,211,231,251,271,1691,2531,2671,3731,3791,5551,5771,6631,6651,6711)" '[FRH]-F:900052
    lSelect = lSelect & " ORDER BY A.COD_CORR, A.COD_UNID"
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    If gfObtRegistro(bdP0044700, lSelect, rsUNIDCORRPROD, pMensagem) = False Then
        Exit Function
    End If
    If rsUNIDCORRPROD.RecordCount = 0 Then
        Call gpFecha3(rsUNIDCORRPROD, bdP0044700)
        If gfNomCorretor(pCodCorretor, lNomCorretor, pMensagem) = False Then
            Exit Function
        End If
        pMensagem = "FTE0069 - Erro ao obter unidade de negócios e produtor (corretor " & pCodCorretor & _
                    " - " & lNomCorretor & " sem informação de unidade de negócios e produtor)."
        Exit Function
    Else
        pTotalRegistros = rsUNIDCORRPROD.RecordCount
    End If

    '4. Encerra processo.
    pCodUniNegocios = rsUNIDCORRPROD!COD_UNID
    pCodProdutor = rsUNIDCORRPROD!Cod_Prod
    Call gpFecha3(rsUNIDCORRPROD, bdP0044700)

    gfObtUniProdutiva = True
End Function
Public Function gfObtProdutor(ByVal pCodCorretor As String, ByRef pCodUniNegocios As Integer, _
                                  ByRef pCodProdutor As Integer, ByRef pMensagem As String, _
                                  Optional ByRef pTotalRegistros As Integer) As Boolean
    'Função: obtém unidade de negócios e produtor do corretor.

    'Parâmetro de entrada...pCodCorretor......Código do corretor.
    'Parâmetros de saída....pCodUniNegocios...Código da unidade de negócios.
    '                       pCodProdutor......Código do produtor.
    '                       pMensagem.........Mensagem.

    Dim bdP0044700      As ADODB.Connection 'Arquivo - P0044700.
    Dim li              As Double           'Contador.
    Dim lNomCorretor    As String           'Nome do corretor.
    Dim lSelect         As String           'Comando SELECT.
    Dim rsUNIDCORRPROD  As ADODB.Recordset  'Registro - YTAB_UNID_CORR_PROD.

    gfObtProdutor = False

    '1. Inicializa saída.
    pCodProdutor = 0

    '2. Verifica parâmetro.
    If Not gfPreenchido(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade de negócios e produtor (código do corretor não " & _
                    "preenchido)."
        Exit Function
    End If

    '3. Busca informação.
    If Not IsNumeric(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade negócios e produtor (código do corretor não " & _
                    "numérico)."
        Exit Function
    End If
    lSelect = "SELECT COD_PROD FROM YTAB_UNID_CORR_PROD WHERE " & _
                " COD_CORR = " & pCodCorretor & _
                " and cod_unid = " & pCodUniNegocios
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    If gfObtRegistro(bdP0044700, lSelect, rsUNIDCORRPROD, pMensagem) = False Then
        Exit Function
    End If
    If rsUNIDCORRPROD.RecordCount = 0 Then
        Call gpFecha3(rsUNIDCORRPROD, bdP0044700)
        If gfNomCorretor(pCodCorretor, lNomCorretor, pMensagem) = False Then
            Exit Function
        End If
        pMensagem = "FTE0069 - Erro ao obter unidade de negócios e produtor (corretor " & pCodCorretor & _
                    " - " & lNomCorretor & " sem informação de unidade de negócios e produtor)."
        Exit Function
    Else
        pTotalRegistros = rsUNIDCORRPROD.RecordCount
    End If

    '4. Encerra processo.
    pCodProdutor = rsUNIDCORRPROD!Cod_Prod
    Call gpFecha3(rsUNIDCORRPROD, bdP0044700)

    gfObtProdutor = True
End Function
Public Function gfObtUnidCorr(ByVal pCodCorretor As String, ByRef pCodUniNegocios As Integer, ByRef pMensagem As String, _
                                  ByRef pTotalRegistros As Integer) As Boolean
    'Função: obtém unidade de negócios e produtor do corretor.

    'Parâmetro de entrada...pCodCorretor......Código do corretor.
    'Parâmetros de saída....pCodUniNegocios...Código da unidade de negócios.
    '                       pCodProdutor......Código do produtor.
    '                       pMensagem.........Mensagem.

    Dim bdP0044700      As ADODB.Connection 'Arquivo - P0044700.
    Dim li              As Double           'Contador.
    Dim lNomCorretor    As String           'Nome do corretor.
    Dim lSelect         As String           'Comando SELECT.
    Dim rsUNIDCORRPROD  As ADODB.Recordset  'Registro - YTAB_UNID_CORR_PROD.

    gfObtUnidCorr = False

    '2. Verifica parâmetro.
    If Not gfPreenchido(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade de negócios e produtor (código do corretor não " & _
                    "preenchido)."
        Exit Function
    End If

    '3. Busca informação.
    If Not IsNumeric(pCodCorretor) Then
        pMensagem = "FTE0069 - Erro ao obter unidade negócios e produtor (código do corretor não " & _
                    "numérico)."
        Exit Function
    End If
    lSelect = "SELECT COD_PROD FROM YTAB_UNID_CORR_PROD WHERE " & _
                " COD_CORR = " & pCodCorretor & _
                " and cod_unid = " & pCodUniNegocios
    If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
        Exit Function
    End If
    If gfObtRegistro(bdP0044700, lSelect, rsUNIDCORRPROD, pMensagem) = False Then
        Exit Function
    End If
    If rsUNIDCORRPROD.RecordCount = 0 Then
        pTotalRegistros = 0
    Else
        pTotalRegistros = rsUNIDCORRPROD.RecordCount
    End If

    '4. Encerra processo.
    Call gpFecha3(rsUNIDCORRPROD, bdP0044700)

    gfObtUnidCorr = True
End Function

Public Function gfRepBasDados_DAO(ByVal pArquivo As String) As Boolean
    'Função: Repara arquivo.

    gfRepBasDados_DAO = False

    On Error Resume Next
    DBEngine.RepairDatabase pArquivo
    If Err <> 0 Then
        Call gpGraLog(2, "FTE0070 - Erro ao reparar arquivo " & pArquivo & ": " & Err & " - " & Error & _
                          ".")
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    gfRepBasDados_DAO = True
End Function
Public Function gfRetornarListIndex(ByVal pCombo As Object, ByVal pItemData As String) As Long
    'Função:

    Dim lCont   As Long

    gfRetornarListIndex = -1
    For lCont = 0 To pCombo.ListCount - 1
        If pCombo.ItemData(lCont) = pItemData Then
            gfRetornarListIndex = lCont
            Exit Function
        End If
    Next
End Function
Public Function gfVerData(ByVal pOpcao As Byte, ByVal pData As String) As Boolean
    'Função: verifica data.

    'Parâmetros de entrada...pOpcao...Opção 1 = data no formato aaaammdd.
    '                                 Opção 2 = data no formato dd/mm/aaaa.
    '                        pData....Data no formato indicado pela opção.

    Dim lAno        As Integer  'Ano.
    Dim lAuxiliar   As String   'Auxiliar.
    Dim lDia        As Integer  'Dia.
    Dim lMes        As Integer  'Mês.
    Dim lResto      As Integer  'Resto.

    gfVerData = False

    '1. Verifica opção.
    If Not gfPreenchido(pOpcao) Then
        Call gpGraLog(2, "FTE0071 - Opção inválida (verifica data - opção não preenchida)." & vbLf & _
                         vbLf & "O sistema será encerrado.")
        Exit Function 'End   ***MOSILVA Verificar
    End If
    If Not IsNumeric(pOpcao) Then
        Call gpGraLog(2, "FTE0071 - Opção inválida (verifica data - opção não numérica)." & vbLf & _
                         vbLf & "O sistema será encerrado.")
        Exit Function 'End   ***MOSILVA Verificar
    End If
    If pOpcao <> 1 And pOpcao <> 2 Then
        Call gpGraLog(2, "FTE0071 - Opção inválida (verifica data - diferente de 1 e 2: " & pOpcao & _
                         ")." & vbLf & vbLf & "O sistema será encerrado.")
        Exit Function 'End   ***MOSILVA Verificar
    End If

    '2. Separa dia, mês e ano.
    Select Case pOpcao
        Case 1      'Data no formato aaaammdd.
            lAuxiliar = Format$(pData, "00000000")
            lDia = Mid$(lAuxiliar, 7, 2)
            lMes = Mid$(lAuxiliar, 5, 2)
            lAno = Mid$(lAuxiliar, 1, 4)
        Case 2
            lDia = Mid$(pData, 1, 2)
            lMes = Mid$(pData, 4, 2)
            lAno = Mid$(pData, 7, 4)
    End Select

    '3. Verifica mês.
    If lMes < 1 Or lMes > 12 Then
        Exit Function
    End If

    '4. Verifica dia.
    If lDia < 1 Then
        Exit Function
    End If
    Select Case lMes
        Case 1, 3, 5, 7, 8, 10, 12
            If lDia > 31 Then
                Exit Function
            End If
        Case 4, 6, 9, 11
            If lDia > 30 Then
                Exit Function
            End If
        Case 2
            Select Case (lAno Mod 4)
                Case 0      'Ano bissexto.
                    If lDia > 29 Then
                        Exit Function
                    End If
                Case Else   'Ano NÃO bissexto.
                    If lDia > 28 Then
                        Exit Function
                    End If
            End Select
    End Select

    gfVerData = True
End Function


Public Sub gpCalDifDatas(ByVal pOpcao As Byte, ByVal pData1 As Double, ByVal pData2 As Double, _
                         ByRef pSaida As Double)
    'Procedure: calcula diferença entre datas.

    'Parâmetros de entrada...pOpcao...Opção 1 = quantidade de dias.
    '                                 Opção 2 = quantidade de meses (prevista - não codificada).
    '                                 Opção 3 = quantidade de anos (idade).
    '                        pData1...Data 1 no formato aaaammdd.
    '                        pData2...Data 2 no formato aaaammdd.
    'Parâmetro de saída......pSaida...Quantidade calculada.

    Dim lAno1       As Integer  'Ano 1: menor data.
    Dim lAno2       As Integer  'Ano 2: maior data.
    Dim lAuxiliar   As String   'Auxiliar.
    Dim lCalcula    As Boolean  'Indica que é para calcular.
    Dim lDia1       As Integer  'Dia 1: menor data.
    Dim lDia2       As Integer  'Dia 2: maior data.
    Dim lMes1       As Integer  'Mês 1: menor data.
    Dim lMes2       As Integer  'Mês 2: maior data.

    '1. Inicializa saída.
    pSaida = 0

    '2. Verifica opção.
    If Not gfPreenchido(pOpcao) Then
        Call gpGraLog(2, "FTE0072 - Opção inválida (calcula diferença entre datas - opção não " & _
                         "preenchida)." & vbLf & vbLf & "O sistema será encerrado.")
        Exit Sub 'End   ***MOSILVA Verificar
    End If
    If Not IsNumeric(pOpcao) Then
        Call gpGraLog(2, "FTE0072 - Opção inválida (calcula diferença entre datas - opção não " & _
                         "numérica)." & vbLf & vbLf & "O sistema será encerrado.")
        Exit Sub 'End   ***MOSILVA Verificar
    End If
    If pOpcao <> 1 And pOpcao <> 3 Then
        Call gpGraLog(2, "FTE0072 - Opção inválida (calcula diferença entre datas - opção diferente " & _
                         "de 1, 2 e 3)." & vbLf & vbLf & "O sistema será encerrado.")
        Exit Sub 'End   ***MOSILVA Verificar
    End If

    '3. Verifica datas.
    If gfVerData(1, pData1) = False Then
        Call gpGraLog(2, "FTE0073 - Erro ao calcular diferença entre datas (primeira data inválida - " & _
                         gfForData(2, pData1) & ").")
        Exit Sub
    End If
    If gfVerData(1, pData2) = False Then
        Call gpGraLog(2, "FTE0073 - Erro ao calcular diferença entre datas (segunda data inválida - " & _
                         gfForData(2, pData2) & ").")
        Exit Sub
    End If

    '4. Se datas iguais, quantidade calculada igual a 0.
    If pData1 = pData2 Then
        Exit Sub
    End If

    '5. Prepara campos.
    If pData1 < pData2 Then
        lAuxiliar = Format$(pData1, "00000000") 'Menor data.
        lDia1 = CInt(Mid$(lAuxiliar, 7, 2))
        lMes1 = CInt(Mid$(lAuxiliar, 5, 2))
        lAno1 = CInt(Mid$(lAuxiliar, 1, 4))
        lAuxiliar = Format$(pData2, "00000000") 'Maior data.
        lDia2 = CInt(Mid$(lAuxiliar, 7, 2))
        lMes2 = CInt(Mid$(lAuxiliar, 5, 2))
        lAno2 = CInt(Mid$(lAuxiliar, 1, 4))
    Else
        lAuxiliar = Format$(pData2, "00000000") 'Menor data.
        lDia1 = CInt(Mid$(lAuxiliar, 7, 2))
        lMes1 = CInt(Mid$(lAuxiliar, 5, 2))
        lAno1 = CInt(Mid$(lAuxiliar, 1, 4))
        lAuxiliar = Format$(pData1, "00000000") 'Maior data.
        lDia2 = CInt(Mid$(lAuxiliar, 7, 2))
        lMes2 = CInt(Mid$(lAuxiliar, 5, 2))
        lAno2 = CInt(Mid$(lAuxiliar, 1, 4))
    End If

    '4. Calcula.
    Select Case pOpcao
        Case 1      'Diferença de dias.
            lCalcula = False
            Do      'Controle de saída dentro do loop.
                If lAno1 = lAno2 Then
                    If lMes1 = lMes2 Then
                        pSaida = pSaida + (lDia2 - lDia1)
                        Exit Do
                    Else    'lMes1 < lMes2.
                        lCalcula = True
                    End If
                Else    'lAno1 < lAno2.
                    If lMes2 = 1 Then
                        pSaida = pSaida + lDia2
                        lAno2 = lAno2 - 1
                        lMes2 = 12
                        lDia2 = 31
                    Else
                        If lMes2 = 12 And lDia2 = 31 Then
                            Select Case lAno2 Mod 4
                                Case 0      'Ano bissexto.
                                    pSaida = pSaida + 366
                                Case Else   'Ano não bissexto.
                                    pSaida = pSaida + 365
                            End Select
                            lAno2 = lAno2 - 1
                        Else
                            lCalcula = True
                        End If
                    End If
                End If
                If lCalcula = True Then
                    lCalcula = False
                    pSaida = pSaida + lDia2
                    lMes2 = lMes2 - 1
                    Select Case lMes2
                        Case 1, 3, 5, 7, 8, 10, 12
                            lDia2 = 31
                        Case 4, 6, 9, 11
                            lDia2 = 30
                        Case 2
                            Select Case lAno2 Mod 4
                                Case 0      'Ano bissexto.
                                    lDia2 = 29
                                Case Else   'Ano NÃO bissexto.
                                    lDia2 = 28
                            End Select
                    End Select
                End If
            Loop
        Case 2      'Diferença de meses.
            MsgBox "Codificar procedure gpCalDifDatas - opção 2."
        Case 3      'Diferença de anos (idade).
            'a) Calcula diferença entre anos.
            pSaida = lAno2 - lAno1
            'b) Verifica mês.
            If lMes2 > lMes1 Then
                Exit Sub
            End If
            If lMes2 < lMes1 Then
                pSaida = pSaida - 1
                Exit Sub
            End If
            'c) Verifica dia.
            If lDia2 < lDia1 Then
                pSaida = pSaida - 1
            End If
            Exit Sub
    End Select
End Sub
Public Sub gpPreencherCombo(ByRef pComboListBox As ComboBox, prst As ADODB.Recordset, _
                            Optional pCampoDescricao As String, Optional pCampoID As String, _
                            Optional pValorDefault As String, Optional pPosicaoDefault As Integer = -1, _
                            Optional pForm As Form = Nothing)
    'Procedure: preencher um listbox ou combobox a partir de uma coleção de objetos de estrutura.

    'Parâmetros...pComboListBox...combobox ou listbox que será preenchido.
    '             pColecao........coleção de objetos de estrutura que implementam a classe stcGenerica.
    '             pValorDefault...item que será posicionado o combo ou listbox.

    Dim bErroValor  As Boolean          'Controle de erro no valor.
    Dim bFoco       As Boolean          'Foco.

    On Error GoTo ControleErro

    bFoco = False
    If Not pForm Is Nothing Then
        If Not pForm.ActiveControl Is Nothing Then
            If UCase(TypeName(pForm.ActiveControl)) = "COMBOBOX" Then
                bFoco = (pForm.ActiveControl.Name = pComboListBox.Name And pForm.ActiveControl.Index = pComboListBox.Index)
            End If
        End If
    End If

    bErroValor = False
    If Not bFoco Then pComboListBox.Visible = False
    pComboListBox.Clear
    If Not (prst.BOF And prst.EOF) Then
        Do While Not prst.EOF
            pComboListBox.AddItem prst("" & pCampoDescricao & "")
            pComboListBox.ItemData(pComboListBox.NewIndex) = prst("" & pCampoID & "")
            prst.MoveNext
        Loop
        
        If pValorDefault <> "" Then
            bErroValor = True
            pComboListBox.Text = pValorDefault
        Else
            pComboListBox.ListIndex = pPosicaoDefault
        End If
    End If
    If Not bFoco Then pComboListBox.Visible = True
    Exit Sub

ControleErro:
    If bErroValor Then pComboListBox.ListIndex = pPosicaoDefault
    pComboListBox.Visible = True
End Sub

'[FRH]-> Movidas para serem utilizadas também no Reg.Padrão [mnuEdiExcRegPad_click]
'[gpExcluiRegistros][antes no BAS111]
'[mfVerRegPadrao][antes no 001A]

Public Sub gpExcluiRegistros(ByVal pNosso_Numero As String, _
                            ByVal pNum_Item As Long, _
                            ByVal pPropTabela As String, _
                            ByVal pNomeTabImport As String, _
                            ByRef pMensagem As String)
Dim lobjPedido              As clsA46V702A      'Objeto da classe clsA46V702A (TAB_PED)
Dim lobjstcPedido           As stcA46V702B      'Objeto da classe de estrutura stcA46V702B (TAB_PED)
Dim lobjItem                As clsA46V708A      'Objeto da classe clsA46V708A (TAB_PED_LOC)
Dim lobjstcItem             As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)

    On Error GoTo ControleErro
  
        If pNum_Item = 0 Then
            Set lobjPedido = New clsA46V702A
            Set lobjstcPedido = New stcA46V702B
            lobjstcPedido.Nosso_Numero = pNosso_Numero
            If Not lobjPedido.mfExcluir(lobjstcPedido, pPropTabela, pNomeTabImport) Then
                pMensagem = lobjPedido.UltimoErro
                Set lobjPedido = Nothing
                Set lobjstcPedido = Nothing
                GoTo SAIDA
            End If
            Set lobjPedido = Nothing
            Set lobjstcPedido = Nothing
        Else
            Set lobjItem = New clsA46V708A
            Set lobjstcItem = New stcA46V708B
            lobjstcItem.Nosso_Numero = pNosso_Numero
            lobjstcItem.Num_Item = pNum_Item
            If Not lobjItem.mfExcluir(lobjstcItem, pPropTabela, pNomeTabImport) Then
                pMensagem = lobjItem.UltimoErro
                Set lobjItem = Nothing
                Set lobjstcItem = Nothing
                GoTo SAIDA
            End If
            Set lobjItem = Nothing
            Set lobjstcItem = Nothing
        End If
    
SAIDA:
    GoSub FechaObj
    Exit Sub
    
ControleErro:
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, pMensagem = "gpExcluiRegistros ERRO - " & Err & " " & Err.Description)
    GoTo SAIDA
FechaObj:
    Set lobjPedido = Nothing
    Set lobjstcPedido = Nothing
    Set lobjItem = Nothing
    Set lobjstcItem = Nothing
    Return
End Sub

Public Function mfVerRegPadrao(ByRef pMensagem As String) As Boolean
    'Função: salva padrão do registro padrão.

    'Parâmetro de saída...pMensagem...Mensagem.

    Dim objPed          As clsA46V702A      'Objeto da classe clsA46V702A       (Tab_Ped)
    Dim objPedLoc       As clsA46V708A      'Objeto da classe clsA46V708A       (Tab_Ped_Loc)
    Dim objJuros        As clsA46V711A      'Objeto da classe clsA46V711A       (Tab_Ped_Juros)
    Dim lDatAlteracao   As Double           'Data da última atualização.
    Dim objstcPed       As stcA46V702B      'Objeto da classe de estrutura stcA46V702B (Tab_Ped)
    Dim objstcPedLoc    As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (Tab_Ped_Loc)
    Dim objstcJuros     As stcA46V711B      'Objeto da classe de estrutura stcA46V711B (Tab_Ped_Juros)
    Dim lMensagem       As String
    
    Dim lobjstcCobert  As stcA46V704B
    Dim lobjCobert     As clsA46V704A
        
    Dim pCodProduto As Integer
    Dim pProtocolo As Long
    Dim pSequencia As Long
    Dim pDescrProd As String

    Dim lPropTabela As String
    Dim lbdP00MPED      As ADODB.Connection
    Dim lrst            As ADODB.Recordset
    Dim lSql            As String
    
    Call gfObtSYASMULTINI(lPropTabela)
    
    mfVerRegPadrao = False

    '1. Posiciona controle.
    lDatAlteracao = Format(Date, "yyyymmdd")

    '2. Instancio os objetos Tab_Ped
    Set objPed = New clsA46V702A
    Set objstcPed = New stcA46V702B
   
    '3. Verifica se tem registro padrão de dados comuns.
    objstcPed.Nosso_Numero = "PADRAO" & gNomUsuario6
    If Not objPed.mfCarregarColecao(objstcPed, lPropTabela, "", True) Then
        pMensagem = objPed.UltimoErro
        Call gpGraLog(1, pMensagem)
        GoTo SAIDA
    End If
    'Se coleção estiver vazia então é incluído o registro padrão.
    If objPed.Colecao.Count = 0 Then
        'Não tem registro padrão - então inclui.
        With objstcPed
            'Carrego as insformações no STC
            .Nosso_Numero = "PADRAO" & gNomUsuario6
            .Cod_Tip_Doc = e_TipDoc_Cotacao
            .Cod_Sit_Ped = e_SitPedido_SemCalculo
            .DAT_DIGIT = lDatAlteracao
            .Cod_User_Digit = gNomUsuario6
            .Tip_Cust_Emis = 1
            .Cod_Estip = 1
        End With
        'Chamo a função de incluir o registro na TAb_PEd
        If Not objPed.mfIncluir(objstcPed, lPropTabela, "") Then
            pMensagem = objPed.UltimoErro
            GoTo SAIDA
        End If
    End If

    '4. Instancio os objetos do Tab_Ped_Loc
    Set objPedLoc = New clsA46V708A
    Set objstcPedLoc = New stcA46V708B
    Set objPedLoc.ConexaoBD = objPed.ConexaoBD
    
    '5. Verifica se tem registro padrão do item.
    objstcPedLoc.Nosso_Numero = "PADRAO" & gNomUsuario6
    If Not objPedLoc.mfCarregarColecao(objstcPedLoc, lPropTabela, "", True) Then
        pMensagem = objPedLoc.UltimoErro
        Call gpGraLog(1, pMensagem)
        GoTo SAIDA
    End If
    
    '6. Se não tem registro padrão - então inclui.
    If objPedLoc.Colecao.Count = 0 Then
        With objstcPedLoc
            'Carrego as informações no stc
            .Nosso_Numero = "PADRAO" & gNomUsuario6
            .Num_Item = 1
            .Cod_Sit_item = e_SitItem_NaoCalc
            .Tip_Emissao = e_TipEmissao_SeguroNovo
            .DAT_DIGIT = lDatAlteracao
            .Cod_User_Digit = gNomUsuario6
            .Cod_Plano = 1
        End With
        'Chamo a função para incluir o registro na TAB_PED_LOC
        If Not objPedLoc.mfIncluir(objstcPedLoc, lPropTabela, "") Then
            pMensagem = objPedLoc.UltimoErro
            Call gpGraLog(1, pMensagem)
            GoTo SAIDA
        End If
    End If
    
                          
    '>---[FRH]--->
    '[Criar um registro padrão para cada um dos ramos de acordo com tabela]
    '[Acessar a tbm_ramo]
    
    Dim lSql_ramos As String
    Dim rs_ramos As ADODB.Recordset
    Dim bdP00MPED As ADODB.Connection
    
    Dim objPed_TMP As clsA46V702A
    Dim objstcPed_TMP As stcA46V702B
    
    Dim objPedLoc_TMP As clsA46V708A
    Dim objstcPedLoc_tmp As stcA46V708B
    
    Dim objCorr_TMP As clsA46V705A
    Dim objstcCorr_TMP As stcA46V705B

    If gfAbrBasDados("P00MPED", gTipP00MPED, gEndP00MPED, bdP00MPED, pMensagem) Then
    
        lSql_ramos = "SELECT * FROM TBM_RAMO"
        If gPrograma = "D046O003" Or gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
            lSql_ramos = lSql_ramos & " where cod_ramo in (112,113,114)"
        End If
        If gfObtRegistro(bdP00MPED, lSql_ramos, rs_ramos, pMensagem) Then
            Do While Not rs_ramos.EOF
                               
                '[TAB_PED]->
                '*****************************************************************************************************************
            
                Set objPed_TMP = New clsA46V702A
                Set objstcPed_TMP = New stcA46V702B
                
                '[Verificar se existe registro padrão de dados comuns]
                objstcPed_TMP.Nosso_Numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                If Not objPed_TMP.mfCarregarColecao(objstcPed_TMP, lPropTabela, "", True) Then
                    pMensagem = objPed_TMP.UltimoErro: Call gpGraLog(1, pMensagem): GoTo SAIDA
                End If
                
                '[Verificar se a coleção é vazia, então incluir registro padrão]
                If objPed_TMP.Colecao.Count = 0 Then
                    With objstcPed_TMP
                        '[Carregar informações no STC]
                        .Cod_Ramo = gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                        .Cod_Tip_Doc = e_TipDoc_Cotacao
                        .Cod_Sit_Ped = e_SitPedido_SemCalculo
                        .DAT_DIGIT = lDatAlteracao
                        .Cod_User_Digit = gNomUsuario6
                        
                        '[Período de vigência]
                        .Dat_Ini_Vig = gHoje
                        If Right(.Dat_Ini_Vig, 4) = "0229" Then
                            .Dat_Ini_Vig = gHoje - 1
                        End If
                        .Dat_Ter_Vig = .Dat_Ini_Vig + 10000
                        
                        '[Produto]
                     '   If gfBusProduto(gHoje, pCodProduto, pProtocolo, pSequencia, pDescrProd, lMensagem, 0) = False Then
                     '       Call gpGraLog(1, lMensagem)
                       '  End If
                        .Cod_Produto = 0
                        .Cod_Reg = 1
                        .Cod_Estip = 1
                        .Tip_Cust_Emis = 1
                    End With
                    '[Incluir registro na TAB_PED]
                    If Not objPed_TMP.mfIncluir(objstcPed_TMP, lPropTabela, "") Then
                        pMensagem = objPed_TMP.UltimoErro: GoTo SAIDA
                    End If
                End If

                '*****************************************************************************************************************
                '<-[TAB_PED]
                
                '[TAB_PED_CORR]->
                '*****************************************************************************************************************
                                
                Set objCorr_TMP = New clsA46V705A
                Set objstcCorr_TMP = New stcA46V705B
                
                '[Verificar se existe registro padrão de dados comuns]
                objstcCorr_TMP.Nosso_Numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                If Not objCorr_TMP.mfCarregarColecao(objstcCorr_TMP, lPropTabela, "") Then
                    pMensagem = objCorr_TMP.UltimoErro: Call gpGraLog(1, pMensagem): GoTo SAIDA
                End If
                
                '[Verificar se a coleção é vazia, então incluir registro padrão]
                If objCorr_TMP.Colecao.Count = 0 Then

                    With objstcCorr_TMP
                        .Nosso_Numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                        .Num_Ped = 0
                        .Num_Ordem_Corr = 1
                        If gCod_Corretor <> 0 Then
                            .Cod_Corr = gCod_Corretor
                        End If
                    End With
                
                    If Not objCorr_TMP.mfIncluir(objstcCorr_TMP, lPropTabela, "") Then
                        pMensagem = objCorr_TMP.UltimoErro: GoTo SAIDA
                    End If
                End If
                
                '*****************************************************************************************************************
                '<-[TAB_PED_CORR]
  
                '[TAB_PED_LOC]->
               '*****************************************************************************************************************
               
                Set objPedLoc_TMP = New clsA46V708A
                Set objstcPedLoc_tmp = New stcA46V708B
                Set objPedLoc_TMP.ConexaoBD = objPed_TMP.ConexaoBD
                
                '[Verificar se tem registro padrão de dados comuns]
                objstcPedLoc_tmp.Nosso_Numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                If Not objPedLoc_TMP.mfCarregarColecao(objstcPedLoc_tmp, lPropTabela, "", True) Then
                    pMensagem = objPedLoc_TMP.UltimoErro: Call gpGraLog(1, pMensagem): GoTo SAIDA
                End If
                
                '[Verificar se a coleção é vazia, então incluir registro padrão]
                If objPedLoc_TMP.Colecao.Count = 0 Then
                    With objstcPedLoc_tmp
                        '[Carregar informações no STC]
                        .Cod_Ramo = gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                        .Num_Item = 1
                        .Cod_Sit_item = e_SitItem_NaoCalc
                        .Tip_Emissao = e_TipEmissao_SeguroNovo
                        .DAT_DIGIT = lDatAlteracao
                        .Cod_User_Digit = gNomUsuario6
                    End With
                    '[Incluir registro na TAB_PED_LOC]
                    If Not objPedLoc_TMP.mfIncluir(objstcPedLoc_tmp, lPropTabela, "") Then
                        pMensagem = objPedLoc_TMP.UltimoErro: GoTo SAIDA
                    End If
                End If
    
                '5. Verifica se tem registro padrão .
                Set lobjstcCobert = New stcA46V704B
                Set lobjCobert = New clsA46V704A
                lobjstcCobert.Nosso_Numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                Set lobjCobert.ConexaoBD = objPed_TMP.ConexaoBD
                If Not lobjCobert.mfCarregarColecao(lobjstcCobert, lPropTabela, "") Then
                    pMensagem = lobjCobert.UltimoErro
                    GoTo SAIDA
                End If
                
                '6. Se não tem registro padrão - então inclui.
                If lobjCobert.Colecao.Count = 0 Then
                    Set lobjstcCobert = New stcA46V704B
                    Set lobjCobert = New clsA46V704A
                    With lobjstcCobert
                        .Nosso_Numero = "PADRAO" & gNomUsuario6 & gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                        .Num_Item = 1
                        .Cod_Ramo = gfTratarNulo(rs_ramos("cod_ramo"), tpoNumerico)
                        .Cod_Cobert = 1
                        .Tab_Cobert = 1
                    End With
                    Set lobjCobert.ConexaoBD = objPed.ConexaoBD
                    If Not lobjCobert.mfIncluir(lobjstcCobert, lPropTabela, "") Then
                        pMensagem = lobjCobert.UltimoErro
                        GoTo SAIDA
                    End If
                End If

                '*****************************************************************************************************************
                '<-[TAB_PED_LOC]
                
            rs_ramos.MoveNext
            Loop
            Set rs_ramos = Nothing
        End If
        
        'deleta indice Tab_Ped_Cobert_Compl_IND03
        If gPrograma <> "D046O003" And gTipUsuario = e_Tipo_Usu_Corretor And gBaseLocal = True Then
            Call gfExeSQL(bdP00MPED, "Drop Index Tab_Ped_Cobert_Compl_IND03 on Tab_Ped_Cobert_Compl", lMensagem)
        End If
        
    Call gpFecha2(bdP00MPED)
    End If
    '<---[FRH]---<
    
    If gTipUsuario = e_Tipo_Usu_Corretor And gPrograma = "P46V001" Then
        If Not gfAbrBasDados("P00MPED", gTipP00MPED, gEndP00MPED, lbdP00MPED, lMensagem) Then
            Call gpGraLog(0, lMensagem)
            Exit Function
        End If
        Dim lObservacao As String
        lSql = " Select * from TabConv"
        If Not gfObtRegistro(lbdP00MPED, lSql, lrst, lMensagem) Then
            lMensagem = ""
            
            lSql = " update tab_ped  set cod_ramo_ativ = 1 where cod_ramo_ativ = 2"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 2 where cod_ramo_ativ = 3"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 3 where cod_ramo_ativ = 4"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 4 where cod_ramo_ativ = 5"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 100 where cod_ramo_ativ = 8"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 8 where cod_ramo_ativ = 9"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 9 where cod_ramo_ativ = 100"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 100 where cod_ramo_ativ = 10"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 101 where cod_ramo_ativ = 11"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 11 where cod_ramo_ativ = 100"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 100 where cod_ramo_ativ = 12"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 12 where cod_ramo_ativ = 101"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 101 where cod_ramo_ativ = 15"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 15 where cod_ramo_ativ = 100"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 100 where cod_ramo_ativ = 16"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 16 where cod_ramo_ativ = 13"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 102 where cod_ramo_ativ = 17"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 17 where cod_ramo_ativ = 14"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 103 where cod_ramo_ativ = 19"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 101"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 101 where cod_ramo_ativ = 20"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 20 where cod_ramo_ativ = 100"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 100 where cod_ramo_ativ = 21"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 21 where cod_ramo_ativ = 102"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 102 where cod_ramo_ativ = 22"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 18"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 23"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 103"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 24"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 25"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            lSql = " update tab_ped  set cod_ramo_ativ = 19 where cod_ramo_ativ = 26"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
            
            lSql = " Create table TabConv (data float null);"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
        
            lSql = " Insert into TabConv values (" & Format(Date, "yyyymmdd") & ")"
            Call gfExeSQL(lbdP00MPED, lSql, lMensagem)
        End If
    End If
    
    mfVerRegPadrao = True
SAIDA:
    GoSub FechaObj
    Exit Function
    
FechaObj:
    '7. Destruo os objetos
    Set objPed = Nothing
    Set objPedLoc = Nothing
    Set objJuros = Nothing
    Set objstcPed = Nothing
    Set objstcPedLoc = Nothing
    Set objstcJuros = Nothing
    Set rs_ramos = Nothing
    Set bdP00MPED = Nothing
    Set objPed_TMP = Nothing
    Set objstcPed_TMP = Nothing
    Set objPedLoc_TMP = Nothing
    Set objstcPedLoc_tmp = Nothing
    Set objCorr_TMP = Nothing
    Set objstcCorr_TMP = Nothing
    Return
End Function


'Função para verificar grupo de corretor na base de dados
'Ficha 1000153 - Automatização do agrupamento de corretores no Syas após o cadastro realizado pelo help-desk.
Public Function gfVerCorrPermitido_TBS_Corr_Grupo(ByVal pCorrInstalacao As Long, _
                                                  ByVal pCorrSelecionado As Long, _
                                                  ByVal pCodRamo As Integer, _
                                                  ByRef pMensagem As String) As Boolean
    Dim bdP0044700  As ADODB.Connection 'Arquivo - P0044700.
    Dim lCnGed  As ADODB.Connection 'Arquivo - P0044700.
    Dim lSelect     As String           'Comando SELECT.
    Dim rsCORR      As ADODB.Recordset  'Tabela: YTAB_CORR.
    Dim lCorrInstacao As String
    Dim lCorrSelecionado As String


Err.Clear
pMensagem = ""
On Error GoTo TrataErro: gfVerCorrPermitido_TBS_Corr_Grupo = False

    lCorrInstacao = Format(pCorrInstalacao, "0000000")
    lCorrSelecionado = Format(pCorrSelecionado, "0000000")

    '1. Verifica parâmetros.
    'a) Opção.
    If Not gfPreenchido(pCorrInstalacao) Then
        pMensagem = "FTE0067 - Opção inválida (obter corretor - opção não preenchida)."
        GoTo SAIDA
    End If
    'b) Corretor.
    If Not gfPreenchido(pCorrSelecionado) Then
        pMensagem = "FTE0068 - Erro ao obter corretor (código do corretor não preenchido)."
        GoTo SAIDA
    End If
    'C) Ramo
    If pCodRamo = 0 Then
        pMensagem = "FTE0068 - Erro ao obter corretor (código do ramo não preenchido)."
        GoTo SAIDA
    End If
    
    If Not (lCorrInstacao = lCorrSelecionado) Then
        If gPrograma = "D046O003" Then
            If gfAbrBasDados("GED", gTipP0044700, gEndP0044700, lCnGed, pMensagem) = False Then
                GoTo SAIDA
            End If
            lSelect = "SELECT CodCorr FROM "
            lSelect = lSelect & " (SELECT CodCorr FROM GED.dbo.Func_BuscaCorrGrupo('" & lCorrInstacao & "', " & pCodRamo & ")) as Grupo"
            lSelect = lSelect & " WHERE CodCorr = " & Val(lCorrSelecionado)
            
            If gfObtRegistro(lCnGed, lSelect, rsCORR, pMensagem) = True Then
                If rsCORR.EOF = True Then
                    pMensagem = "FGE0017 - Código corretor informado (" & lCorrSelecionado & ") não permitido para este usuário(" & gNomUsuario & " - Corr:" & lCorrInstacao & ")"
                    GoTo SAIDA
                End If
            Else
                GoTo SAIDA
            End If
        Else
            If gfAbrBasDados("P0044700", gTipP0044700, gEndP0044700, bdP0044700, pMensagem) = False Then
                GoTo SAIDA
            End If
            'instalação  = Código aliado  e informado = cód líder
            lSelect = " SELECT * FROM TBS_Corr_Grupo " & _
                      "WHERE CodAliado = '" & lCorrInstacao & _
                      "' AND CodCorr = '" & lCorrSelecionado & _
                      "' AND CodRamo = 112 "
            If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
                GoTo SAIDA
            End If
            If rsCORR.RecordCount = 0 Then
                'instalação  = cód líder  e  informado = Código aliado
                lSelect = " SELECT * FROM TBS_Corr_Grupo " & _
                          "WHERE CodCorr = '" & lCorrInstacao & _
                          "' AND CodAliado = '" & lCorrSelecionado & _
                          "' AND CodRamo = 112 "
                If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
                    GoTo SAIDA
                End If
            Else
                gfVerCorrPermitido_TBS_Corr_Grupo = True
                GoTo SAIDA
            End If
            If rsCORR.RecordCount = 0 Then
                'instalação =  código aliado e informado = código aliado só que ambos são aliados do mesmo líder.
                lSelect = "SELECT A.CodCorr" & _
                          "  FROM TBS_Corr_Grupo AS A " & _
                          " INNER JOIN TBS_Corr_Grupo AS B " & _
                          "    ON A.CodCorr = B.CodCorr " & _
                          " WHERE (((A.CodAliado)= '" & lCorrInstacao & "' ) " & _
                          "   AND ((B.CodAliado)='" & lCorrSelecionado & "') " & _
                          "   AND ((A.CodRamo)=112)) "
                If gfObtRegistro(bdP0044700, lSelect, rsCORR, pMensagem) = False Then
                    GoTo SAIDA
                End If
                If rsCORR.RecordCount = 0 Then
                    Call gpFecha3(rsCORR, bdP0044700)
                    pMensagem = "FGE0003 - " & gPrefixo & "Corretor inválido (código de corretor não " & _
                                "localizado - " & lCorrInstacao & ")."
                    GoTo SAIDA
                Else
                    gfVerCorrPermitido_TBS_Corr_Grupo = True
                    GoTo SAIDA
                End If
            Else
                gfVerCorrPermitido_TBS_Corr_Grupo = True
                GoTo SAIDA
            End If
        End If
    End If
    
'=================== OK Permitido!!!
    gfVerCorrPermitido_TBS_Corr_Grupo = True
'===================
SAIDA:
    If gfVerCorrPermitido_TBS_Corr_Grupo = False Then
        Call gpGraLog(2, pMensagem)
    End If
    GoSub FechaObj
    Err.Clear
    Exit Function
    
TrataErro:
    If gfPreenchido(pMensagem) = False Then
        pMensagem = Err & " - " & Error
    End If
    pMensagem = "gfVerCorrPermitido_TBS_Corr_Grupo Erro: " & pMensagem
    GoTo SAIDA
FechaObj:
    Call gpFecha3(rsCORR, lCnGed)
    Set lCnGed = Nothing
    Set rsCORR = Nothing
    Set bdP0044700 = Nothing
    Return
End Function

Public Function gfCalNumPedido(ByRef pNumPedido As String, _
                               ByRef pMensagem As String, _
                               Optional ByRef pBasDados As ADODB.Connection) As Boolean
    'Função: calcula número do pedido.

    'Parâmetro de entrada...pBasDados....Base de dados.
    'Parâmetros de saída....pNumPedido...Número do pedido.
    '                       pMensagem....Mensagem.

Dim lAuxiliar       As Long             'Auxiliar.
Dim lobjUltPed      As clsA46V701A      'Objeto da classe clsA46V701A (TAB_PED_ULT)
Dim lobjstcUltPed   As stcA46V701B      'Objeto da classe de estrutura stcA46V701B (TAB_PED_ULT)
Dim lPropTabela     As String

On Error GoTo ControleErro: gfCalNumPedido = False

    Call gfObtSYASMULTINI(lPropTabela)

    Set lobjUltPed = New clsA46V701A
    Set lobjstcUltPed = New stcA46V701B
    
    Set lobjUltPed.ConexaoBD = pBasDados
    
    If gPrograma = "D046O003" Then
        If Not lobjUltPed.mfCarregarColecaoSyasPortal(gNomUsuario6, lPropTabela) Then
            pMensagem = lobjUltPed.UltimoErro
            GoTo ControleErro
        End If
        If lobjUltPed.Colecao.Count = 0 Then
            pMensagem = "Não foi possivel gerar o numero do pedido."
            GoTo ControleErro
        Else
            Set lobjstcUltPed = lobjUltPed.Colecao.Item(1)
        End If
        
        'Exclui o pedido anteriormente gravado com este numero
        Dim lClsPed As clsA46V702A
        Dim lStcPed As stcA46V702B
        Set lClsPed = New clsA46V702A
        Set lStcPed = New stcA46V702B
        
        lStcPed.Num_Ped = lobjstcUltPed.Ult_Num_Ped
        If lClsPed.mfCarregarColecao(lStcPed, lPropTabela, "") = True Then
            For Each lStcPed In lClsPed.Colecao
                If lClsPed.mfExcluir(lStcPed, lPropTabela, "") = False Then
                    pMensagem = lClsPed.UltimoErro
                    GoTo ControleErro
                End If
            Next lStcPed
        Else
            pMensagem = lClsPed.UltimoErro
            GoTo ControleErro
        End If
        
    Else
        '1. Busca o último número do pedido
        If Not lobjUltPed.mfCarregarColecao(lPropTabela) Then
            pMensagem = lobjUltPed.UltimoErro
            GoTo ControleErro
        End If
        
        '2. Calcula número do pedido e atualiza tabela.
        If lobjUltPed.Colecao.Count = 0 Then
            lobjstcUltPed.Ult_Num_Ped = 1
            lobjstcUltPed.Cod_User_Alt = gNomUsuario6
            lobjstcUltPed.Dat_Alt = gfForData(11)
            If Not lobjUltPed.mfIncluir(lobjstcUltPed, lPropTabela) Then
                pMensagem = lobjUltPed.UltimoErro
                GoTo ControleErro
            End If
        Else
            Set lobjstcUltPed = lobjUltPed.Colecao.Item(1)
            If lobjstcUltPed.Ult_Num_Ped = 0 Then
                lobjstcUltPed.Ult_Num_Ped = 1
            Else
                lobjstcUltPed.Ult_Num_Ped = lobjstcUltPed.Ult_Num_Ped + 1
                If lobjstcUltPed.Ult_Num_Ped > 999999 Then
                    lobjstcUltPed.Ult_Num_Ped = 1
                End If
            End If
            lobjstcUltPed.Cod_User_Alt = gNomUsuario6
            lobjstcUltPed.Dat_Alt = gfForData(11)
            If Not lobjUltPed.mfAlterar(lobjstcUltPed, lPropTabela) Then
                pMensagem = lobjUltPed.UltimoErro
                GoTo ControleErro
            End If
        End If
    End If
    
    pNumPedido = Format$(lobjstcUltPed.Ult_Num_Ped, "000000")

    gfCalNumPedido = True
    
SAIDA:
    Err.Clear: On Error GoTo 0
    GoSub FechaObj
    Exit Function
    
ControleErro:
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfCalNumPedido ERRO - " & Err & " " & Error & pMensagem)
    GoTo SAIDA
    
FechaObj:
    Set lClsPed = Nothing
    Set lStcPed = Nothing
    Set lobjUltPed = Nothing
    Set lobjstcUltPed = Nothing
    Return
End Function


Function gfCopProposta(ByVal pAntNosNumero As String, _
                       ByVal pNovNosNumero As String, _
                       ByVal pNovNumPedido As String, _
                       ByRef pMensagem As String, _
                       Optional ByVal pNovPI As Long, _
                       Optional pDatRecCia As Long, _
                       Optional pCodRamo As Integer, _
                       Optional ByRef pCalcLMI As Integer, _
                       Optional pNaoRepProduto As Boolean, _
                       Optional pReentrada As Integer) As Boolean
    'Função: monta comando INSERT genérico para dados de proposta.

    'Parâmetros de entrada...pAntNumItem....Anterior: número do item.
    '                        pmfNovNumItem....Novo: número do item.
    'Parâmetro de saída......pMensagem......Mensagem.

Dim bdBasDados  As ADODB.Connection 'Base de dados.
Dim flTabela    As ADODB.Field      'Campos da tabela.
Dim rsTabela    As ADODB.Recordset  'Registro da tabela.
Dim lCount      As Long             'Utilizado no for...next.
Dim lIdeTabela  As String           'Nome da tabela.
Dim lContinuar  As Boolean          'Indicador se continua o procedimento ou não.
Dim lSelect     As String
Dim lInsert     As String
Dim lPropTabela As String
On Error GoTo ControleErro
    
    gfCopProposta = False

    Call gfObtSYASMULTINI(lPropTabela)

    '07/01/2010 - Ficha: 0901900 - Cópia de cotação/proposta c/ última tarifa -- inicio
    Dim lCodProduto     As Integer          'Código do produto.
    Dim lCodProtocolo   As Long             'Código do protocolo.
    Dim lCodSequencia   As Long             'Código da seqüência.
    Dim lDesTarifa      As String           'Descrição da tarifa.
    Dim lMensagem       As String           'Mensagem.
    
    
    If gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
'Ficha 901718
        If pNaoRepProduto = False Then 'Reposicionar na última tarifa
            If gfBusProduto(gHoje, lCodProduto, lCodProtocolo, lCodSequencia, lDesTarifa, lMensagem, pCodRamo) = _
               False Then
                Call gpGraLog(1, lMensagem)
               GoTo SAIDA
            End If
            '07/01/2010 - Ficha: 0901900 - Cópia de cotação/proposta c/ última tarifa -- fim
        End If
    End If


    '1. Obtém registros da tabela.
    If gfAbrBasDados("P00MPED", gTipP00MPED, gEndP00MPED, bdBasDados, _
            pMensagem) = False Then GoTo SAIDA
    
    For lCount = 1 To 20
        lContinuar = True
        Select Case lCount
            Case 1
                lIdeTabela = "TAB_PED"
            Case 2
                lIdeTabela = "TAB_PED_CORR"
            Case 3
                lIdeTabela = "TAB_PED_COSS"
            Case 4
                lIdeTabela = "TAB_PED_JUROS"
            Case 5
                lIdeTabela = "TAB_PED_VENC"
            Case 6
                lIdeTabela = "TAB_PED_ENDER"
            Case 7
                lIdeTabela = "TAB_CTRL_EMIS"
            Case 8
                lIdeTabela = "TAB_PED_LOC"
            Case 9
                lIdeTabela = "TAB_PED_CLAU"
            Case 10
                lIdeTabela = "TAB_PED_COBERT"
            Case 11
                lIdeTabela = "TAB_PED_LOC_BEM"
            Case 12
                lIdeTabela = "TAB_PED_LOC_SEGURO"
            Case 13
                lIdeTabela = "TAB_PED_OBS"
            Case 14
                lIdeTabela = "TAB_PED_INSP"
            Case 15
                lIdeTabela = "TAB_PED_AUTO"
            Case 16
                lIdeTabela = "TAB_PED_COBERT_COMPL"
            Case 17
                lIdeTabela = "TAB_PED_FAT_RISCO"
            Case 18
                lIdeTabela = "TAB_PED_DIZER"
            Case 19
                lIdeTabela = "TAB_PED_ESPEC"
            Case 20
                lIdeTabela = "TAB_PED_ESPEC2"
        End Select
        
        lSelect = "SELECT * FROM " & lPropTabela & lIdeTabela & " WHERE Nosso_Numero = '" & pAntNosNumero & "'"
        If gfObtRegistro(bdBasDados, lSelect, rsTabela, pMensagem) = False Then
            Call gpFecha2(bdBasDados)
            GoTo SAIDA
        End If
        If rsTabela.EOF = True Then
            lContinuar = False
            If lIdeTabela = "TAB_PED" Then
                pMensagem = "FNL0003 - " & gPrefixo & "Pedido não localizado (" & pAntNosNumero & _
                                 " - cópia da cotação/proposta.)."
                GoTo SAIDA
            End If
        End If
        If lContinuar = True Then
            '2. Grava nova cotação/proposta.
            While Not rsTabela.EOF
                lInsert = "INSERT INTO " & lPropTabela & lIdeTabela & " ("
                For Each flTabela In rsTabela.Fields
                    lInsert = lInsert & flTabela.Name & ", "
                Next flTabela
                lInsert = Left$(lInsert, Len(lInsert) - 2) & ") VALUES ("
                For Each flTabela In rsTabela.Fields
                    Select Case UCase(flTabela.Name)
                        Case "NOSSO_NUMERO"
                            lInsert = lInsert & "'" & pNovNosNumero & "', "                         'Nosso número
                        Case "NUM_PED"
                            lInsert = lInsert & pNovNumPedido & ", "                                'Número do pedido.
                        Case "SIT_DOCUMENTO", "COD_SIT_PED"
                            lInsert = lInsert & "1, "       'DOCUMENTO NÃO CALCULADO.               'Sit_Documento.
                        Case "DAT_LIBER"
                            lInsert = lInsert & Left(gfForData(3), 8) & ", "                        'Dat_Alt.
                        Case "DAT_PRI_CALCULO"
                            lInsert = lInsert & "0, "                                               'Dat_Pri_Calculo.
                        Case "DAT_DIGIT"
                            lInsert = lInsert & Left(gfForData(3), 8) & ", "                        'Dat_Digit
                        Case "COD_USER_LIBER"
                            lInsert = lInsert & "'COPIA',"                                  'Cod_User_Alt.
                        Case "NUM_BLOQ"
                            If pReentrada = 1 Then
                                lInsert = lInsert & flTabela.Value & ", "
                            Else
                                lInsert = lInsert & "0,"                                        'NUM_BLOQ
                            End If
                        Case "NUM_PI"
                            If pNovPI = 0 Then
                                lInsert = lInsert & Right(Year(Date), 2) & Format(pNovNumPedido, "00000000") & ", "
                            Else
                                lInsert = lInsert & pNovPI & ", "
                            End If
                        Case "DAT_REC_CIA"
                            If pDatRecCia = 0 Then
                                lInsert = lInsert & flTabela.Value & ", "
                            Else
                                lInsert = lInsert & pDatRecCia & ", "
                            End If
                        Case "NUM_APOL", "NUM_ENDO", "DAT_EMIS", "COD_IDEN_STAT_ACAO"
                                lInsert = lInsert & "0, "
                        Case "COD_PRODUTO"
                            If lIdeTabela = "TAB_PED" Then
                                If lCodProduto <> 0 Then
                                    lInsert = lInsert & lCodProduto & ", "
                                Else
                                    lInsert = lInsert & flTabela.Value & ", "
                                End If
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case "COD_IDEN_TARIF"
                            If lCodProduto <> 0 Then
                                lInsert = lInsert & lCodProduto & ", "
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case "COD_SEQUENCIA"
                            If lCodSequencia <> 0 Then
                                lInsert = lInsert & lCodSequencia & ", "
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case "LMI_UNICO"
                            If IsNull(flTabela.Value) Then
                                lInsert = lInsert & "0, "
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                            pCalcLMI = gfTratarNulo(flTabela.Value, tpoNumerico)
                        Case "NUM_COTAC"    '1100090 - Inspecao Desabilitada
                            lInsert = lInsert & "'000000000000',"                               'NUM_COTAC.
                        Case Else
                            Select Case flTabela.Type
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
                                        lInsert = lInsert & gfForValor(1, flTabela.Value) & ", "
                                    End If
                                Case adChar, adVarChar, adLongVarChar                          'Texto.
                                    If gfPreenchido(flTabela.Value) Then
                                        lInsert = lInsert & "'" & gfForAspas(flTabela.Value) & "', "
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
                                    Call gpFecha3(rsTabela, bdBasDados)
                                    pMensagem = "FTE0008 - Tabela " & lIdeTabela & " com tipo do campo " & _
                                                flTabela.Name & " diferente dos previstos."
                                    GoTo SAIDA
                            End Select
                    End Select
                Next flTabela
                If Mid(lInsert, Len(lInsert), 1) = "," Then
                    lInsert = Left(lInsert, Len(lInsert) - 1) & ")"
                Else
                    lInsert = Left(lInsert, Len(lInsert) - 2) & ")"
                End If
                If gfExeSQL(bdBasDados, lInsert, pMensagem) <> 0 Then
                    GoTo SAIDA
                End If
                rsTabela.MoveNext
            Wend
        End If
        Call gpFecha1(rsTabela)
    Next lCount
    
    gfCopProposta = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
    
ControleErro:
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfCopProposta ERRO - " & Err & " " & Err.Description)
    GoTo SAIDA
    
FechaObj:
    Call gpFecha3(rsTabela, bdBasDados)
    Set bdBasDados = Nothing
    Set flTabela = Nothing
    Set rsTabela = Nothing
    Return
    
End Function
Public Sub gpGraLog(ByVal pOpcao As Byte, ByVal pMensagem As String, Optional pMsgBoxSyle As VbMsgBoxStyle = vbApplicationModal)
    'Procedure: grava arquivo log.

    'Parâmetros de entrada...pOpcao......0 = grava log.
    '                                    1 = grava log e emite mensagem.
    '                                    2 = grava log e emite mensagem com frase inicial
    '                                        "Comunique Central de Atendimento".
    '                        pMensagem...Mensagem.

Dim lMensagem           As String   'Mensagem.
Dim lNumLog             As Integer  'Número do arquivo log.
Dim lPosicao            As Integer  'Utilizado com InStr.
Dim lMsgErr             As String   'Mensagem de erro dentro desta procedure
Dim lCkPoint            As String    'Variavel utilizada para determinar pontos onde o código já passou.
Dim lArqAberto          As Boolean

Static lEndLog          As String   'Endereço do arquivo log.
Static lNomComputador   As String   'Nome do computador.
Static lNomPrograma     As String   'Nome do programa.
Static lNomUsuario      As String   'Nome do usuário.
Static lNumMensagem     As Integer  'Número de mensagem.
Static lPriVez          As Boolean  'Indicador de primeira vez (false).

Err = 0
On Error GoTo TRATAMENTO_ERRO
    
'1. Verifica opção.
    lCkPoint = "1"
    If Not gfPreenchido(pOpcao) Then
        MsgBox "FTE0006 - Opção inválida (grava log - não preenchida)."
        GoTo SAIDA
    End If
    lCkPoint = "2"
    If Not IsNumeric(pOpcao) Then
        MsgBox "FTE0006 - Opção inválida (grava log - não numérica)."
        GoTo SAIDA
    End If
    lCkPoint = "3"
    If pOpcao <> 0 And pOpcao <> 1 And pOpcao <> 2 Then
        MsgBox pMensagem
        MsgBox "FTE0006 - Opção inválida (grava log - diferente de 0, 1 e 2 - " & pOpcao & ")."
        GoTo SAIDA
    End If
    lCkPoint = "4"
    If gVerDataMaq = 1 Then
        lCkPoint = "5"
        If gfIncDatas = False Then
            'Travar o sistema
            gTrvSistema = True
        End If
    End If
    lCkPoint = "6    "
'2. Ao executar esta procedure pela primeira vez, obter dados utilizados na gravação das mensagens.
    If lPriVez = False Then
        lCkPoint = "7"
        lPriVez = True
        'a) Verifica arquivo log.
        lCkPoint = "8"
        If gfPreenchido(gNomArqLog) = False Then
            gNomArqLog = "SYASMULT.LOG"
        End If
        lCkPoint = "9"
        lEndLog = gAppPath & gNomArqLog
        'b) Obtém nome do usuário.
        lCkPoint = "10"
        If gfPreenchido(gNomUsuario) Then
            lCkPoint = "11"
            lNomUsuario = gNomUsuario
            While Len(lNomUsuario) < 20
                lCkPoint = "12"
                lNomUsuario = lNomUsuario & " "
            Wend
            lCkPoint = "13"
        Else
            lCkPoint = "14"
            lNomUsuario = String(20, "X")
        End If
        lCkPoint = "15"
        'c) Obtém nome do computador.
        lNomComputador = GetLoginComputerName()
        lCkPoint = "16"
        If Not gfPreenchido(lNomComputador) Then
            lNomComputador = "********"
        End If
        lCkPoint = "17"
        While Len(lNomComputador) < 16
            lCkPoint = "18"
            lNomComputador = lNomComputador & " "
        Wend
        lCkPoint = "19"
        'd) Obtém nome do programa.
        lNomPrograma = Trim$(App.EXEName)
        lCkPoint = "20"
        While Len(lNomPrograma) < 12
            lCkPoint = "21"
            lNomPrograma = lNomPrograma & " "
        Wend
        lCkPoint = "22"
    End If
    lCkPoint = "23"
    
'3. Consiste parâmetro de entrada (Mensagem).
    lCkPoint = "24"
    If Not gfPreenchido(pMensagem) Then
        lCkPoint = "25"
        pMensagem = "FTE0045 - Mensagem não preenchida."
    End If
    lCkPoint = "26"
'4. Ajusta a mensagem
    lMensagem = pMensagem
    Do
        lCkPoint = "27"
        lPosicao = InStr(lMensagem, vbLf)
        lCkPoint = "28"
        If lPosicao = 0 Then
            Exit Do
        End If
        lCkPoint = "29"
        lMensagem = Left$(lMensagem, lPosicao - 1) & " " & Mid$(lMensagem, lPosicao + 1)
    Loop
        
'3. Abre arquivo log.
    lCkPoint = "30"
    lNumLog = FreeFile
    lCkPoint = "31"
    lArqAberto = False
    lCkPoint = "32"
    Open lEndLog For Append Shared As #lNumLog
        lCkPoint = "33"
        lArqAberto = True
        
        'x. Posiciona número da mensagem.
        lCkPoint = "34"
        If lNumMensagem = 9999 Then
            lCkPoint = "35"
            lNumMensagem = 1
        Else
            lCkPoint = "36"
            lNumMensagem = lNumMensagem + 1
        End If
        lCkPoint = "37"
    
        'x. Grava mensagem.
        Print #lNumLog, gForHoje & " | " & Format$(Now, "hh:nn:ss") & " | " & lNomUsuario & " | " & _
                        lNomComputador & " | " & lNomPrograma & " | " & Format$(lNumMensagem, "0000") & _
                        " | " & lMensagem
        lCkPoint = "38"
FecharArquivo:
    lCkPoint = "36"
    Close #lNumLog
    lCkPoint = "40"
    lArqAberto = False      'marca que o arquivo não está mais aberto
    lCkPoint = "41"
    
'6. Emite mensagem ao usuário.
    lCkPoint = "42"
    If gProibidoMsgBox = True Then
        pOpcao = 0
    End If
    Select Case pOpcao
        Case 0
            lCkPoint = "43"
            'Nada a fazer.
        Case 1
            lCkPoint = "44"
            MsgBox pMensagem, pMsgBoxSyle
        Case 2
            lCkPoint = "45"
            MsgBox "Comunique Central de Atendimento. " & vbLf & vbLf & pMensagem & vbLf & vbLf & "Data/Hora da ocorrência : " & gForHoje & " " & Format$(Now, "Hh:Nn:Ss"), pMsgBoxSyle
    End Select
    lCkPoint = "46"
    
SAIDA:
    If lArqAberto = True Then
        'Fecha o arquivo caso o código tenha sido desviado para o erro
        'após ter aberto o arquivo.
        Close #lNumLog
        lArqAberto = False
    End If
    Err = 0
    On Error GoTo 0
    Exit Sub
    
TRATAMENTO_ERRO:
    If gfPreenchido(lMsgErr) = False Then
        lMsgErr = Err & " - " & Error
    End If
    lMsgErr = "gpGraLog Erro " & lMsgErr
    lMsgErr = lMsgErr & vbCrLf & "CheckPoint " & lCkPoint
    
    lMensagem = "Comunique Central de Atendimento. " & vbLf & vbLf & lMensagem & vbLf & vbLf & "Data/Hora da ocorrência : " & gForHoje & " " & Format$(Now, "Hh:Nn:Ss")
    lMsgErr = lMsgErr & vbCrLf & vbCrLf & gForHoje & " | " & Format$(Now, "hh:nn:ss") & " | " & lNomUsuario & " | " & _
                                            lNomComputador & " | " & lNomPrograma & " | " & Format$(lNumMensagem, "0000") & _
                                            " | " & lMensagem
    If pOpcao <> 0 Then
        MsgBox lMsgErr
    End If
    
    GoTo SAIDA
    
End Sub
Public Function gfPreenchido(ByVal pCampo) As Boolean
    'Função: verifica se campo preenchido.

    If IsNull(pCampo) Then
        gfPreenchido = False
    Else
        If Len(Trim(pCampo)) = 0 Then
            gfPreenchido = False
        Else
            gfPreenchido = True
        End If
    End If
End Function

Public Function gfAbrBasDados(ByVal pIndBasDados As String, ByVal pTipBasDados As String, _
                              ByVal pEndBasDados As String, ByRef pBasDados As ADODB.Connection, _
                              ByRef pMensagem As String) As Boolean
    'Função: abre base de dados.

    'Parâmetros de entrada...pIndBasDados...Indicador da base de dados.
    '                        pTipBasDados...Tipo da base de dados.
    '                        pEndBasDados...Endereço da base de dados.
    'Parâmetros de saída.....pBasDados......Base de dados.
    '                        pMensagem......Mensagem.

    Dim lSqlServer As Boolean           'Indica se a base deverá ser aberta com SQL SERVER.

    gfAbrBasDados = False
    
    '1. Verifica indicador do arquivo.
    If Not gfPreenchido(pIndBasDados) Then
        pMensagem = "[ERR] - FTE0028 - Erro ao abrir base de dados (indicador não preenchido)."
        Exit Function
    End If
    
    lSqlServer = False
    If (UCase(pIndBasDados) = "SYAS_EMIS" And gPerfilAcesso = 3) Or gPrograma = "D046O003" Then
        lSqlServer = True
    Else
        If gBaseLocal Then
            Select Case UCase$(pIndBasDados)
                Case "P00MPED"
                    'If gTipoAcesso = DIGITACAO Or gPerfilAcesso = e_Perfil_Operacao Then
                    '    lSqlServer = True
                    'Else
                        lSqlServer = False
                    'End If
                Case "P00APOL", "P00MPED", "P00MULTI", "P0044700", "P0043000", _
                     "P00MULTI2", "P0045200", "P0045300", "P0045400"
                    lSqlServer = False
                Case "GED"
                    lSqlServer = True
                Case Else
                    pMensagem = "[ERR] - FTE0029 - Erro ao abrir base de dados (indicador desconhecido)."
                    Exit Function
            End Select
        Else
            Select Case UCase$(pIndBasDados)
                Case "P00MULTI2"
                    lSqlServer = False
                Case "P00APOL", "P00MPED", "GED", "P00MULTI", "RAMOSDIVERSOS", "P0043000", "P0044700", "P0045400", "EAPDB", "SYASPORTAL"
                    lSqlServer = True
                Case Else
                    pMensagem = "[ERR] - FTE0029 - Erro ao abrir base de dados (indicador desconhecido)."
                    Exit Function
            End Select
        End If
    End If
    If lSqlServer = True Then
        If UCase$(pIndBasDados) = "GED" Then
            gfAbrBasDados = gfAbrBasSQL("GED", "2", gSerGed, gEndGed, gUidGed, gPwdGed, pBasDados, pMensagem)
        ElseIf UCase$(pIndBasDados) = "EAPDB" Then
            gfAbrBasDados = gfAbrBasSQL("EAPDB", "2", gSerEAPDB, gEndEAPDB, gUidEAPDB, gPwdEAPDB, pBasDados, pMensagem)
        ElseIf UCase$(pIndBasDados) = "SYASPORTAL" Then
            gfAbrBasDados = gfAbrBasSQL("SYASPORTAL", "2", gSerSyasPortal, gEndSyasPortal, gUidSyasPortal, gPwdSyasPortal, pBasDados, pMensagem)
        Else
            gfAbrBasDados = gfAbrBasSQL("DBDIVERSOS", "2", gSerDIVERSOS, gEndDIVERSOS, gUidDIVERSOS, _
                                    gPwdDIVERSOS, pBasDados, pMensagem)
        End If
    Else
        gfAbrBasDados = gfAbrBasDadosAccess(pIndBasDados, pTipBasDados, pEndBasDados, pBasDados, _
                                            pMensagem)
    End If
    On Error GoTo 0
End Function
Public Function gfAbrBasDadosAccess(ByVal pIndBasDados As String, ByVal pTipBasDados As String, _
                                    ByVal pEndBasDados As String, ByRef pBasDados As ADODB.Connection, _
                                    ByRef pMensagem As String) As Boolean
    'Função: abre base de dados Access.

    'Parâmetros de entrada...pIndBasDados...Indicador da base de dados.
    '                        pTipBasDados...Tipo da base de dados.
    '                        pEndBasDados...Endereço da base de dados.
    'Parâmetros de saída.....pBasDados......Base de dados.
    '                        pMensagem......Mensagem.

    Dim lAbriu          As Boolean  'Abriu a base de dados?
    Dim lComando        As String   'Comando.
    Dim lDriver(6)      As String   'Drivers.
    Dim lErrDriver(6)   As String   'Grava mensagem de erro obtida em cada driver
    Dim li              As Byte     'Utilizado com For...Next.
    Dim lIniCatalog     As String   'Parâmetro INITIAL CATALOG.
    Dim lPwd            As String   'Parâmetro PWD.
    Dim lServer         As String   'Parâmetro SERVER.
    Dim lTemErro        As Boolean  'Tem erro?
    Dim lUid            As String   'Parâmetro UID.
    Dim lMsgErro        As String   'Mensagem de erro

    gfAbrBasDadosAccess = False

    If pIndBasDados = "P00MULTI2" Then
        pIndBasDados = "P00MULTI"
        pTipBasDados = 1
    End If


    '1. Verifica indicador do arquivo.
    If Not gfPreenchido(pIndBasDados) Then
        pMensagem = "[ERR] - FTE0028 - Erro ao abrir base de dados (indicador não preenchido)."
        Exit Function
    End If
    Select Case UCase$(pIndBasDados)
        Case "P00APOL", "P00MPED", "P00MULTI", "P0044700", "P0043000", _
             "P0045200", "P0045300", "P0045400"
            'Nada a fazer - arquivos válidos.
        Case Else
            pMensagem = "[ERR] - FTE0029 - Erro ao abrir base de dados (indicador desconhecido - " & _
                        pIndBasDados & ")."
            Exit Function
    End Select

    '2. Verifica tipo do arquivo.
    If Not gfPreenchido(pTipBasDados) Then
        pMensagem = "[ERR] - FTE0030 - Erro ao abrir base de dados (" & pIndBasDados & _
                    " - tipo da base de dados não preenchido)."
        Exit Function
    End If
    If pTipBasDados <> "1" Then
        pMensagem = "[ERR] - FTE0031 - Erro ao abrir base de dados (" & pIndBasDados & _
                    " - tipo da base de dados desconhecido: " & pTipBasDados & ")."
        Exit Function
    End If

    '3. Verifica endereço do arquivo.
    If Not gfPreenchido(pEndBasDados) Then
        pMensagem = "[ERR] - FTE0032 - Erro ao abrir base de dados (" & pIndBasDados & _
                    " - endereço da base de dados não preenchido)."
        Exit Function
    End If
    If Not gfPreenchido(Dir(pEndBasDados)) Then
        pMensagem = "[ERR] - FTE0033 - Erro ao abrir base de dados (" & pIndBasDados & _
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
    For li = 1 To 6
        'a) Posiciona conexão.
        On Error Resume Next
        Set pBasDados = New ADODB.Connection
        If Err <> 0 Then
            pMensagem = "[ERR] - FTE0034 - Erro ao abrir base de dados (" & pIndBasDados & _
                        " - comando: SET New ADODB.Connection - Erro " & Err & " - " & Error & ")."
            lTemErro = True
            Exit For
        End If
        'b) Posiciona propriedade CursorLocation.
        pBasDados.CursorLocation = adUseClient
        If Err <> 0 Then
            pMensagem = "[ERR] - FTE0035 - Erro ao abrir base de dados (" & pIndBasDados & _
                        " - comando: CursorLocation = adUseClient - Erro " & Err & " - " & Error & ")."
            lTemErro = True
            Exit For
        End If
        'c) Abre arquivo.
        lComando = "Driver={" & lDriver(li) & "};Dbq=" & pEndBasDados
        ''Call gpGraLog(0, "AGE0000 - gfAbrBasDadosAccess: lComando = " & lComando & " pIndBasDados=" & _
                         pIndBasDados)
                         
        Select Case pIndBasDados
            Case "P00APOL", "P00MPED", "P00MULTI"
                pBasDados.Open lComando, , LCase(pIndBasDados) & "46"
            Case "P0044700", "P0043000", "P0045200", "P0045300", "P0045400"
                pBasDados.Open lComando, , LCase(pIndBasDados) & "21"
        End Select
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
            ElseIf InStr(1, UCase(lMsgErro), "MODO EXCLUSIVO") > 0 Or InStr(1, UCase(lMsgErro), "PERMISSÃO") > 0 Then
                If Len(pEndBasDados) > 12 Then
                    lMsgErro = " - Privilégios insuficientes no diretório " & Left(pEndBasDados, Len(pEndBasDados) - 12)
                Else
                    lMsgErro = " - Privilégios insuficientes no diretório " & gAppPath
                End If
                Exit For
            End If
            'Guarda a mensagem de erro não prevista
            lErrDriver(li) = lMsgErro
        End If
        On Error GoTo 0
    Next li
    If lTemErro Then
        On Error GoTo 0
        Exit Function
    End If
    If lAbriu = False Then
        pMensagem = "[ERR] - FTE0036 - Erro ao abrir base de dados (" & pIndBasDados & " - comando: " & _
                    lComando & IIf(lMsgErro = "", " - Erro " & Err & " - " & Error, lMsgErro) & ")."
        On Error GoTo 0
        For li = 1 To 6
            'Concatena mensagens de erro diferentes
            If gfPreenchido(lErrDriver(li)) = True Then
                If lErrDriver(li) <> lMsgErro Then
                    pMensagem = pMensagem & Chr(13) & "- " & li & lErrDriver(li)
                End If
            End If
        Next li
        Exit Function
    End If
    On Error GoTo 0

    gfAbrBasDadosAccess = True
End Function

Public Function gfObtRegistro(ByVal pNomArquivo As ADODB.Connection, ByVal pComando As String, _
                              ByRef pRegistro As ADODB.Recordset, ByRef pMensagem As String) As Boolean
    'Função: obtém registro.

    'Parâmetros de entrada...pNomArquivo...Nome do arquivo.
    '                        pComando .....Comando SELECT.
    'Parâmetros de saída.....pRegistro.....Registro (recordset).
    '                        pMensagem.....Mensagem.

    gfObtRegistro = False
    Dim lTimeTer    As Date
    Dim lTimeInic   As Date
    
    lTimeInic = Time

    '1. Posiciona controles.
    On Error Resume Next
    ''Call gpGraLog(0, "AGE0000 - gfObtRegistro: pComando = " & pComando)
    Set pRegistro = New ADODB.Recordset
    pRegistro.CursorLocation = adUseClient

    '2. Ajuste o comando.
    If InStr(1, UCase(pNomArquivo), "SYAS_EMIS") > 0 Then
        If InStr(1, UCase(pComando), "GED") > 0 And InStr(1, UCase(pComando), "SYAS_EMIS") > 0 Then
           'Não altera o comando.
        Else
            'Somente se não estiver fazendo "JOIN" entre tabelas da base GED e SYAS_EMIS através de
            'conexão com a base SYAS_EMIS.
            pComando = Replace(pComando, "'", "!;$")
            If Len(pComando) > 5000 Then
                pComando = " exec Yproc_ExecDML2 '" & Mid(pComando, 1, 7500) & "','" & _
                           Mid(pComando, 7501) & "'"
            Else
                pComando = " exec Yproc_ExecDML2 '" & Mid(pComando, 1, 10) & "','" & _
                           Mid(pComando, 11) & "'"
            End If
        End If
    Else
        If InStr(1, UCase(pNomArquivo), "EAPDB") > 0 Then
            pComando = " exec EAPDB_ExecDML '" & Replace(pComando, "'", "!;$") & "'"
        End If
    End If

    '3. Abre o registro.
    pRegistro.Open pComando, pNomArquivo, adOpenForwardOnly, adLockReadOnly
    'Set pRegistro = pNomArquivo.Execute(pComando)

    If Err <> 0 Then
        pMensagem = "[ERR] - FTE0039 - Erro ao obter registro (Erro = " & Err & " - " & Err.Description & " - comando: " & _
                    pComando & ")."
        Call gpGraLog(0, pMensagem)
        Exit Function
    End If
    
    lTimeTer = Time
'    Call gpGraLog(0, DateDiff("s", lTimeInic, lTimeTer) & " - " & pComando)
    
    If InStr(pNomArquivo, "UID=") = 0 And InStr(pNomArquivo, "Initial Catalog=") = 0 Then
        Set pRegistro.ActiveConnection = Nothing
    End If
    Select Case Err
        Case 0      'Sem erro.
            gfObtRegistro = True
        Case 3159   'Marcador inválido - arquivo dados corrompido.
            pMensagem = "[ERR] - FTE0039 - Erro ao obter registro (arquivo corrompido - comando: " & pComando & _
                        ")."
        Case Else
            pMensagem = "[ERR] - FTE0039 - Erro ao obter registro (Erro = " & Err & " - " & Error & _
                        " - comando: " & pComando & ")."
    End Select
    On Error GoTo 0
End Function

Public Function gfWhereAnd(pSql As String) As String
    'Função: Preenche com WHERE ou AND.

    If InStr(1, UCase(pSql), "WHERE") = 0 Then
        gfWhereAnd = pSql & " WHERE "
    Else
        gfWhereAnd = pSql & " AND "
    End If
End Function

Public Function mfFormatarCamposSQL(pAtributo As Variant, Optional SemDesformatar As Boolean = False) As String
    'Função: formata campos SQL.

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
                    mfFormatarCamposSQL = "'" & gfRetirarCaracterEspecial(CStr(pAtributo), True) & "'"
                End If
            End If
        Case vbDate:
            If Val(pAtributo) = 0 Then
                mfFormatarCamposSQL = "Null"
            Else
                mfFormatarCamposSQL = "'" & Format(pAtributo, "mm-dd-yyyy") & " " & _
                                            Format(pAtributo, "hh:mm:ss") & "'"
            End If
        Case vbBoolean:
            mfFormatarCamposSQL = IIf(pAtributo = True, "1", "0")
    End Select
End Function

Public Function gfTratarNulo(pAtributo As Variant, pTipodeDado As TipoDados) As Variant
    'Função: trata campo nulo.

    If IsNull(pAtributo) Then
        Select Case pTipodeDado
            Case TpoBoolean
                gfTratarNulo = 0
            Case TpoData
                gfTratarNulo = "00:00:00"
            Case tpoNumerico
                gfTratarNulo = 0
            Case TpoHora
                gfTratarNulo = "00:00"
            Case TpoString
                gfTratarNulo = " "
            Case TpoValor
                gfTratarNulo = 0
        End Select
    Else
        gfTratarNulo = Trim(pAtributo)
    End If
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
    Dim lTimeInic   As Date
    Dim lTimeTer   As Date
    
    lTimeInic = Time
    If InStr(1, UCase(pBasDados), "SYAS_EMIS") > 0 Then
        pComando = " exec Yproc_ExecDML '" & Replace(pComando, "'", "!;$") & "'"
    Else
        If InStr(1, UCase(pBasDados), "EAPDB") > 0 Then
            pComando = " exec EAPDB_ExecDML '" & Replace(pComando, "'", "!;$") & "'"
        End If
    End If
    'Call gpGraLog(0, "AGE0000 - gfExeSQL: pComando = " & pComando)
    On Error Resume Next
    If IsMissing(pRegAfetados) Then
        pBasDados.Execute pComando
    Else
        pBasDados.Execute pComando, pRegAfetados
    End If
    lTimeTer = Time
    'Call gpGraLog(0, DateDiff("s", lTimeInic, lTimeTer) & " - " & pComando)
    If Err <> 0 Then
        If InStr(1, UCase(Err.Description), "REGISTRO EXCLUÍDO") <> 0 And InStr(1, UCase(pComando), "DELETE") <> 0 Then
            gfExeSQL = 0
            Exit Function
        Else
            pMensagem = "[ERR] - FTE0038 - Erro ao executar comando SQL: " & Err & " - " & Error & " - " & pComando
        End If
    End If
    gfExeSQL = Err      'Este comando precisa ser o último antes do On Error GoTo 0.
    On Error GoTo 0
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
Public Sub gpFecha2(ByRef pBasDados As ADODB.Connection)
    'Procedure: fecha base de dados.
    'Parâmetro de entrada...pBasDados...Base de dados.

    If Not pBasDados Is Nothing Then
        If InStr(1, UCase(pBasDados.ConnectionString), ".MDB") > 0 Or _
           InStr(1, UCase(pBasDados.Properties("Extended Properties")), ".MDB") Then
            If pBasDados.State = 1 Then
                pBasDados.Close
            End If
            Set pBasDados = Nothing
        End If
    End If
    
End Sub
Public Sub gpFecha3(ByRef pTabela As ADODB.Recordset, ByRef pBasDados As ADODB.Connection)
    'Procedure: fecha tabela e base de dados.

    'Parâmetros de entrada...pTabela.....Recordset da tabela.
    '                        pBasDados...Base de dados.

    '1. Fecha tabela.
    If Not pTabela Is Nothing Then
        If pTabela.State = 1 Then
            pTabela.Close
        End If
    End If
    
    Set pTabela = Nothing

    '2. Fecha base de dadoss.
    If Not pBasDados Is Nothing Then
        If InStr(1, UCase(pBasDados.ConnectionString), ".MDB") > 0 Or _
           InStr(1, UCase(pBasDados.Properties("Extended Properties")), ".MDB") Then
            If pBasDados.State = 1 Then
                pBasDados.Close
            End If
            Set pBasDados = Nothing
        End If
    End If
End Sub

Public Function gfIncDatas() As Boolean
    'Função:

    Dim lDatArqDestrava     As String   'Data do arquivo que destrava o sistema.
    Dim lDatArquivo         As String   'Data que está no arquivo de destravamento
    Dim lDatMaquina         As String   'Data da máquina.
    Dim lDatRegistro        As String   'Data que está salva no registro.
    Dim lDatUltArquivo      As String   'Data do último arquivo destravado.
    Dim lEndArqDestrava     As String   'Endereço do arquivo que destrava o sistema.
    Dim lIndDestravado      As Byte     'Indicador de destravado:
                                        '0 = travado.
                                        '1 = destravado.
                                        '3 = ?
    Dim lMensagem           As String   'Mensagem.
    Dim lNumDestravado      As Byte     'Utilizado no FreeFile.
    Dim lRegistro           As Variant  'Registro - linha que está no arquivo.
    Dim lStaData            As String   'Status da data.
    Dim lUltArqDestravado   As String   'Endereço do último arquivo destravado.

    '1. Mensagem de sistema bloqueado.
    lMensagem = "Comunique Central de Atendimento. " & vbLf & vbLf & _
                "FGE0022 - O sistema está bloqueado pois houve alteração na data/hora da máquina." & _
                vbLf & vbLf & "Para desbloquear o sistema, acerte a data/hora da máquina e entre em " & _
                "contato com a Central de Atendimento."

    '2. Obtém registro da última data de log.
    lDatRegistro = GetSetting("SYAS", "UltimaData", "LOG")
    lDatRegistro = Replace(lDatRegistro, "-", "/")
    lDatRegistro = Replace(lDatRegistro, ".", "/")
    lStaData = GetSetting("SYAS", "StatusData", "LOG")
    lDatMaquina = Format(Date, "dd/mm/yyyy")
    lUltArqDestravado = GetSetting("SYAS", "ArquivoDestravar", "LOG")

    '3. Se não tem registro salvo, grava a data atual no registro.
    If Not gfPreenchido(lDatRegistro) Then
        Call SaveSetting("SYAS", "UltimaData", "LOG", lDatMaquina)
    End If

    '4. Verifica se pode continuar o processo.
    If lStaData = "TRAVADO" Then
        lIndDestravado = 0
        'É necessário destravar o sistema - verifica se existe arquivo que destrava o sistema.
        lEndArqDestrava = Dir(gAppPath & "*.dtr")
        lDatArquivo = "01/01/1001 00:00:00"
        Do While lEndArqDestrava <> ""
            'Nome do arquivo: dd - dia atual                             --> Exemplo: 17.
            '                 mm - mês atual                             --> Exemplo: 01.
            '                 yyyy - dois primeiros número do ano atual  --> Exemplo: 2004.
            '                 hh - horas                                 --> Exemplo: 16.
            '                 mm - minutos                               --> Exemplo: 26.
            '                 ss - dois último número do ano atual       --> Exemplo: 04.
            '                 resultado ddmmyyyyhhmmss.dtr               --> Exemplo: 17012004162604.dtr.
            lDatArqDestrava = Mid(lEndArqDestrava, 1, 2) & "/" & _
                              Mid(lEndArqDestrava, 3, 2) & "/" & _
                              Mid(lEndArqDestrava, 5, 4) & " " & _
                              Mid(lEndArqDestrava, 9, 2) & ":" & _
                              Mid(lEndArqDestrava, 11, 2) & ":" & _
                              Mid(lEndArqDestrava, 13, 2)
            If Len(lEndArqDestrava) = 18 Then
                If lEndArqDestrava <> lUltArqDestravado Then
                    If gfPreenchido(lUltArqDestravado) Then
                        lDatUltArquivo = Mid(lUltArqDestravado, 1, 2) & "/" & _
                                         Mid(lUltArqDestravado, 3, 2) & "/" & _
                                         Mid(lUltArqDestravado, 5, 4) & " " & _
                                         Mid(lUltArqDestravado, 9, 2) & ":" & _
                                         Mid(lUltArqDestravado, 11, 2) & ":" & _
                                         Mid(lUltArqDestravado, 13, 2)
                        If CVDate(lDatArqDestrava) < CVDate(lDatUltArquivo) Then
                            lIndDestravado = 3
                        End If
                    End If
                    If CVDate(lDatArqDestrava) > CDate(lDatArquivo) And lIndDestravado = 0 Then
                        lDatArquivo = lDatArqDestrava
                        lNumDestravado = FreeFile
                        Open lEndArqDestrava For Input As #lNumDestravado
                        While Not EOF(lNumDestravado)
                            'a) Obtém registro de entrada.
                            Line Input #lNumDestravado, lRegistro
                            If gfCripEncript(lRegistro, 0) <> lEndArqDestrava Then
                                If gfCripEncript(lRegistro, 0) <> lEndArqDestrava & "T" Then
                                    lIndDestravado = 0
                                    Call SaveSetting("SYAS", "StatusData", "LOG", "TRAVADO")
                                Else
                                    Call SaveSetting("SYAS", "StatusData", "LOG", "DESTRAVADO")
                                    Call SaveSetting("SYAS", "UltimaData", "LOG", _
                                                     Format(Date, "dd/mm/yyyy"))
                                    lIndDestravado = 1
                                End If
                            Else
                                If lDatMaquina < lDatRegistro Then
                                    'Se a data do computador não foi ajustada, finalizar o sistema.
                                    lIndDestravado = 0
                                Else
                                    'Se a data do computador for igual à data do registro, destravar.
                                    Call SaveSetting("SYAS", "StatusData", "LOG", "DESTRAVADO")
                                    Call SaveSetting("SYAS", "ArquivoDestravar", "LOG", lEndArqDestrava)
                                    lIndDestravado = 1
                                End If
                            End If
                        Wend
                        Close #lNumDestravado
                        Kill (gAppPath & "*.dtr")
                    Else
                        Kill (gAppPath & "*.dtr")
                    End If
                End If
            End If
            lEndArqDestrava = Dir   'Obtém próxima entrada.
        Loop
        If lIndDestravado = 0 Or lIndDestravado = 3 Then
            MsgBox lMensagem
            gfIncDatas = False
            Exit Function
        End If
    Else
        If CDate(lDatMaquina) < CDate(lDatRegistro) Then
            'Finalizar o sistema.
            Call SaveSetting("SYAS", "StatusData", "LOG", "TRAVADO")
            If Dir(gAppPath & "*.dtr") <> "" Then
                Kill (gAppPath & "*.dtr")
            End If
            MsgBox lMensagem
            gfIncDatas = False
            Exit Function
        Else
            If lDatMaquina <> lDatRegistro Then
                'Registra a última data.
                Call SaveSetting("SYAS", "UltimaData", "LOG", lDatMaquina)
            End If
        End If
    End If
    gfIncDatas = True
End Function

Public Function gfAbrBasSQL(ByVal pIdeBasDados As String, ByVal pTipBasDados As String, _
                             ByVal pSerBasDados As String, ByVal pEndBasDados As String, _
                             ByVal pUIDBasDados As String, ByVal pPWDBasDados As String, _
                             ByRef pBasDados As ADODB.Connection, ByRef pMensagem As String) _
                             As Boolean
    'Função: abre base de dados.

    'Parâmetros de entrada...pIdeBasDados...Identificador da base de dados.
    '                        pTipBasDados...Tipo da base de dados.
    '                        pEndBasDados...Endereço da base de dados.
    'Parâmetros de saída.....pBasDados......Base de dados.
    '                        pMensagem......Mensagem.

    Dim lFlag           As Boolean  'Indica se o procedimento utilizará ou não a ultima conexão aberta:
                                    'TRUE = abre nova conexão e FALSE = mantém última conexão aberta.
    Dim lFlag2          As Boolean  'Indica se encontrou divergência na abertura da base de dados.
                                    'TRUE = encontrou divergência e
                                    'FALSE = não encontrou divergência, continua a verificação.
    Dim lInd            As Integer  'Contador.
    Dim lInd2           As Integer  'Contador auxiliar.
    Dim lIndMax         As Integer  'Maior índice.
    Dim lInfoConexao()  As String   'Informação sobre conexão.
    Dim lIniCatalog     As String   'Parâmetro INITIAL CATALOG.
    Dim lMensagem       As String   'Mensagem.
    Dim lOpen           As String   'Linha para comando OPEN.
    Dim lSubInfoCn()    As String   'Recebe itens da variável lInfoConexao.

    '1. Verifica indicador de gravação de log.
    ''Call gpGraLog(0, "AGE0000 - gfAbrBasSQL: pSerBasDados = " & pSerBasDados & " pEndBasDados = " & _
    ''                 pEndBasDados & " pUIDBasDados = " & pUIDBasDados & " pPWDBasDados = " & pPWDBasDados)

    '2. Verifica se a conexão foi a última a ser aberta.
    If mConexaoSQL Is Nothing Then
        'Atribui TRUE para variável lFlag para abrir nova conexão.
        lFlag = True
    Else
        'Separa as informações da ultima conexão aberta em uma variável matriz.
        lInfoConexao = Split(mConexaoSQL, ";")
        lInd = 0
        lMensagem = ""
        lIndMax = UBound(lInfoConexao())
        Do Until lInd = lIndMax + 1
            lMensagem = lMensagem & "Ind(" & lInd & ")=" & lInfoConexao(lInd) & " "
            lInd = lInd + 1
        Loop
       '' Call gpGraLog(0, "AGE0000 - gfAbrBasSQL : lInfoConexao - " & lMensagem)
        'Inicia variável para abrir nova conexão.
        lFlag = True
        'Inicia variável para prosseguir com a verificação.
        lFlag2 = False
        For lInd = LBound(lInfoConexao()) To UBound(lInfoConexao())
            If UBound(lInfoConexao()) < 0 Then
                'Se o campo estiver vazio, encerra verificação e o procedimento abrirá nova conexão.
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
                            'Servidores de base de dados iguais:
                            'Atribui FALSE para variável lFlag para manter a última conexão aberta.
                            lFlag = False
                        Else
                            'Encontrada divergência:
                            'Atribui TRUE para variável lFlag2 para encerrar verificação.
                            lFlag2 = True
                            'Atribui TRUE para abrir nova conexão.
                            lFlag = True
                        End If
                    Case "UID"
                        If lInfoConexao(lInd) = "UID=" & pUIDBasDados Then
                            'Usuários são iguais.
                            lFlag = False
                        Else
                            lFlag2 = True
                            lFlag = True
                        End If
                    Case "PWD"
                        If lInfoConexao(lInd) = "PWD=" & pPWDBasDados Then
                            'Senhas de acesso à base de dados são iguais.
                            lFlag = False
                        Else
                            lFlag2 = True
                            lFlag = True
                        End If
                    Case "DATABASE"
                        If Replace(lInfoConexao(lInd), """", "") = "DATABASE=" & pEndBasDados Then
                            'Bases de dados são iguais.
                            lFlag = False
                        Else
                            lFlag2 = True
                            lFlag = True
                        End If
                End Select
            Next lInd2
            If lFlag2 = True Then
                'Encontrada divergência:
                'lFlag = TRUE - procedimento abrirá nova conexão.
                'Encerra a verificação.
                Exit For
            End If
        Next lInd
    End If
    gfAbrBasSQL = False

    '3. Abrir nova conexão.
    If lFlag = True Then
        'a) Verifica identificação da base de dados.
        If Not gfPreenchido(pIdeBasDados) Then
            pMensagem = "[ERR] - FTE0051 - Erro ao abrir base de dados (identificação da base de dados não " & _
                        "preenchida)."
            Exit Function
        End If
        Select Case UCase$(pIdeBasDados)
            Case "DBDIVERSOS", "GED", "EAPDB", "P0044800", "SYASPORTAL"
                'Nada a fazer - bases de dados válidas.
            Case Else
                pMensagem = "[ERR] - FTE0052 - Erro ao abrir base de dados (identificação da base de dados " & _
                            "desconhecida: " & pIdeBasDados & ")."
                Exit Function
        End Select
        'b) Verifica tipo da base de dados.
        If Not gfPreenchido(pTipBasDados) Then
            pMensagem = "[ERR] - FTE0053 - Erro ao abrir base de dados (tipo da base de dados " & pIdeBasDados & _
                        " não preenchido)."
            Exit Function
        End If
        If pTipBasDados <> "2" Then
            pMensagem = "[ERR] - FTE0052 - Erro ao abrir base de dados (tipo da base de dados " & pIdeBasDados & _
                        " desconhecido: " & pTipBasDados & ")."
            Exit Function
        End If
        'c) Verifica endereço da base de dados.
        If Not gfPreenchido(pEndBasDados) Then
            pMensagem = "[ERR] - FTE0055 - Erro ao abrir base de dados (endereço da base de dados " & _
                        pEndBasDados & " não preenchido)."
            Exit Function
        End If
        'd) Posiciona conexão.
        On Error Resume Next
        Set mConexaoSQL = New ADODB.Connection
        If Err <> 0 Then
            pMensagem = "[ERR] - FTE0030 - Erro ao abrir base de dados (" & pIdeBasDados & _
                        " - comando SET New ADODB.Connection - erro " & Err & " - " & Error & ")."
            On Error GoTo 0
            Exit Function
        End If
        mConexaoSQL.ConnectionString = "driver={SQL Server};server=" & pSerBasDados & _
                                       ";uid=" & pUIDBasDados & ";pwd=" & pPWDBasDados & _
                                       ";database=" & pEndBasDados
        mConexaoSQL.ConnectionTimeout = 180
        mConexaoSQL.CommandTimeout = 180
        If Err <> 0 Then
            pMensagem = "[ERR] - FTE0035 - Erro ao abrir base de dados (" & pIdeBasDados & _
                        " - Comando CursorLocation = adUseClient - erro: " & Err & " - " & Error & ")."
            On Error GoTo 0
            Exit Function
        End If
        ''Call gpGraLog(0, "AGE0000 - gfAbrBasSQL: ConnectionString = " & mConexaoSQL.ConnectionString)
        'e) Abre base de dados.
        mConexaoSQL.Open
        If Err <> 0 Then
            pMensagem = "[ERR] - FTE0056 - Erro ao abrir base de dados (" & pIdeBasDados & _
                        " - comando Open " & lOpen & " - erro: " & Err & " - " & Error & ")."
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If
    Set pBasDados = mConexaoSQL

    gfAbrBasSQL = True
End Function

Public Function gfRetirarCaracterEspecial(pTexto As String, Optional pUcase As Boolean = False, Optional pRetiraCrLf As Boolean = True, Optional pRetiraDemais As Boolean = False) As String
    Dim I           As Long         'Utilizado no for...loop
    Dim Letra       As String       'Letra
    Dim lRetorno    As String       'Retorno
    
    gfRetirarCaracterEspecial = ""

    For I = 1 To Len(pTexto)
        Letra = Mid(pTexto, I, 1)
        Select Case Asc(Letra)
        Case 34, 39, 94, 96, 126, 168, 180   ' ", ' , ^, `, ~, ¨, ´,
            Letra = " "
        Case 13, 10, 59  'Enter, ;
            If pRetiraCrLf Then
                Letra = " "
            End If
        Case 9  'TAB
            Letra = " "
        Case 192 To 198 ' À, Á, Â, Ã, Ä, Å, Æ
            Letra = "A"
        Case 224 To 229 ' à, á, â, ã, ä, å, æ
            Letra = "a"
        Case 199        ' Ç
            Letra = "C"
        Case 231        ' ç
            Letra = "c"
        Case 200 To 203 ' È, É, Ê, Ë
            Letra = "E"
        Case 232 To 235 ' è, é, ê, ë
            Letra = "e"
        Case 204 To 207 ' Ì, Í, Î, Ï
            Letra = "I"
        Case 236 To 239 ' ì, í, î, ï
            Letra = "i"
        Case 209        ' Ñ
            Letra = "N"
        Case 241        ' ñ
            Letra = "n"
        Case 210 To 214 ' Ò, Ó, Ô, Õ, Ö
            Letra = "O"
        Case 242 To 246 ' ò, ó, ô, õ, ö
            Letra = "o"
        Case 217 To 220 ' Ù, Ú, Û, Ü
            Letra = "U"
        Case 249 To 252 ' ù, ú, û, ü
            Letra = "u"
        Case 33, 35, 36, 37, 38, 58, 60, 61, 62, 63, 64 ' !, #, $, %, &, :, <, =, >, ?, @
            If pRetiraDemais Then
                Letra = " "
            End If
        Case 91, 92, 93, 95, 123, 124, 125, 126 ' [, \, ], _, {, |, }, ~
            If pRetiraDemais Then
                Letra = " "
            End If
        Case 40 To 47 ' (, ), *, +, -, ., /
            If pRetiraDemais Then
                Letra = " "
            End If
        End Select
        lRetorno = lRetorno & Letra
    Next I
    If pUcase = True Then
        'Mudo o texto para maiúsculo.
        gfRetirarCaracterEspecial = UCase(lRetorno)
    Else
        gfRetirarCaracterEspecial = lRetorno
    End If
End Function

Function gfCripEncript(ByVal strng, force%)
    'Código copiado da internet para encriptar e decriptar

    Dim chK             As String
    Dim look            As String
    Dim addin           As String
    Dim eNcryptflag     As String
    Dim passUP          As String
    Dim passMAx         As String
    Dim looPer          As Variant
    Dim tOChange        As String
    Dim password        As String

    'Set error capture routine
    On Local Error GoTo ErrorHandler

    password = "Password"

    'Is there a strng$ to work with?
    If Len(strng) = 0 Then Error 31100

    'Check if file is encrypted and not forcing
    If force% = 0 Then

        'Check for encryption ID tag
        chK = Left$(strng, 4) + Right$(strng, 4)
        If chK = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then

            'Remove ID tag
            strng = Mid$(strng, 5, Len(strng) - 8)

            'String was encrypted so filter out CHR$(1) flags
            look = 1
            Do
                look = InStr(look, strng, Chr$(1))
                If look = 0 Then
                    Exit Do
                Else
                    addin = Chr$(Asc(Mid$(strng, look + 1)) - 1)
                    strng = Left$(strng, look - 1) + addin + Mid$(strng, look + 2)
                End If
                look = look + 1
            Loop

            'Since it is encrypted we want to decrypt it
            eNcryptflag = False
        Else
            'Tag not found so flag to encrypt string
            eNcryptflag = True
        End If
    Else
        'force% flag set, ecrypt string regardless of tag
        eNcryptflag = True
    End If

    'Set up variables
    passUP = 1
    passMAx = Len(password)

    'Tack on leading characters to prevent repetative recognition
    password = Chr$(Asc(Left$(password, 1)) Xor passMAx) + password
    password = Chr$(Asc(Mid$(password, 1, 1)) Xor Asc(Mid$(password, 2, 1))) + password
    password = password + Chr$(Asc(Right$(password, 1)) Xor passMAx)
    password = password + Chr$(Asc(Right$(password, 2)) Xor Asc(Right$(password, 1)))

    'If Encrypting add password check tag now so it is encrypted with string
    If eNcryptflag = True Then
        strng = Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") + strng
    End If

    'Loop until scanned though the whole string
    For looPer = 1 To Len(strng)
        DoEvents
        'Alter character code
        tOChange = Asc(Mid$(strng, looPer, 1)) Xor Asc(Mid$(password, passUP, 1))

        'Insert altered character code
        Mid$(strng, looPer, 1) = Chr$(tOChange)

        'Scroll through password string one character at a time
        passUP = passUP + 1
        If passUP > passMAx + 4 Then passUP = 1
    Next looPer

    'If encrypting we need to filter out all bad character codes (0, 10, 13, 26)
    If eNcryptflag = True Then
        'First get rid of all CHR$(1) since that is what we use for our flag
        look = 1
        Do
            look = InStr(look, strng, Chr$(1))
            If look > 0 Then
                strng = Left$(strng, look - 1) + Chr$(1) + Chr$(2) + Mid$(strng, look + 1)
                look = look + 1
            End If
        Loop While look > 0

        'Check for CHR$(0)
        Do
            look = InStr(strng, Chr$(0))
            If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(1) + Mid$(strng, look + 1)
        Loop While look > 0

        'Check for CHR$(10)
        Do
            look = InStr(strng, Chr$(10))
            If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(11) + Mid$(strng, look + 1)
        Loop While look > 0

        'Check for CHR$(13)
        Do
            look = InStr(strng, Chr$(13))
            If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(14) + Mid$(strng, look + 1)
        Loop While look > 0

        'Check for CHR$(26)
        Do
            look = InStr(strng, Chr$(26))
            If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(27) + Mid$(strng, look + 1)
        Loop While look > 0

        'Tack on encryted tag
        strng = Chr$(1) + "KT" + Chr$(1) + strng + Chr$(1) + "KT" + Chr$(1)

    Else

        'We decrypted so ensure password used was the correct one
        If Left$(strng, 9) <> Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") Then
            'Password bad cause error
            Error 31100
        Else
            'Password good, remove password check tag
            strng = Mid$(strng, 10)
        End If
    End If

    'Set function equal to modified string
    gfCripEncript = strng

    'Were out of here
    Exit Function

ErrorHandler:
    'We had an error!  Were out of here
    Exit Function
End Function

Public Sub gpMenErro(ByVal pMensagem As String)
    'Procedure: posiciona mensagem de erro.
    'Parâmetro de entrada.........pMensagem...Mensagem.
                       
    '[FRH] - F:900097 [FGE0155]
    '[FRH]-GLB_sinal_concess-> true se as propr.Cod_Clasf e Cod_Compl_Clasf do pobjstcitem-mfcalpreitem são relativos a "022-13-Concessionária".
    If GLB_sinal_concess And ( _
        pMensagem Like "*0058*0022*" Or _
        pMensagem Like "*0062*0022*" Or _
            (UCase(pMensagem) Like "FGE0155*COBERTURA*86*NÃO*PODE*SER*CONTRATADA*SEM*COBERTURA*28*" And _
                gCodProduto >= 956) _
    ) Then
        Exit Sub
    End If
    If gCod_Ramo = e_CodRamo_Roubos And (pMensagem Like "*52*") Then
        Exit Sub
    End If
    If (UCase(pMensagem) Like "*NORMA*" Or UCase(pMensagem) Like "*FGE0116*") And _
            Not UCase(pMensagem) Like "*SEM ACEITA*" Then
        gQuaErrosNormas = gQuaErrosNormas + 1
    Else
        gQuaErrosDemais = gQuaErrosDemais + 1
    End If

    'Rotina de captura de erros original
    gQuaErros = gQuaErros + 1
    ReDim Preserve gErros(gQuaErros)
    gErros(gQuaErros) = pMensagem
    '[FRH]
    
End Sub

Public Function gfTrunca(ByVal pCampo As String, ByVal pQuaCasas As Byte) As Single
    'Função #031: trunca campo.

    Dim lPosicao    As Byte     'Utilizado com InStr.

    lPosicao = InStr(CDec(pCampo), ",")
    If lPosicao = 0 Then
        lPosicao = InStr(CDec(pCampo), ".")
    End If
    If lPosicao = 0 Then
        gfTrunca = CDec(pCampo)
    Else
        If Not IsNumeric(Mid(Trim(CDec(pCampo)), lPosicao + pQuaCasas, 1)) Then
            gfTrunca = CDec(pCampo)
        Else
            gfTrunca = (Mid(CDec(pCampo), 1, lPosicao - 1) & _
                        Mid(CDec(pCampo), lPosicao + 1, pQuaCasas)) / 10 ^ pQuaCasas
        End If
    End If
End Function

Public Sub gpInicia(ByRef pPropTabela As String)
    'Procedure: obtém informações inicias do sistema:
    '           - Arquivo de inicialização: tipo do usuário, endereços das bases de dados (endereço do
    '                                       servidor), endereços dos arquivos imagem e configuração de
    '                                       impressão.
    '           - Nome do usuário (ou código do usuário no caso de corretor).

    Dim lAuxiliar           As String   'Auxiliar.
    Dim lEndSyasEdi_EXE    As String    'Endereço do arquivo SyasEdi.EXE.
    Dim lEndSyaslibSGS      As String   'Endereço do arquivo SYASLIB.SGS.
    Dim lLiberado           As Boolean  'Indica que o Segsul está liberado.
    Dim lLimite             As Double   'Data limite para o arquivo de liberação do Segsul.
    Dim lNumSyasEdi_EXE    As Integer   'Utilizado com FreeFile para o arquivo SYASEDI.EXE.
    Dim lNumSyaslibSGS      As Integer  'Utilizado com FreeFile para o arquivo SYASLIB.SGS.
    Dim lPosicao            As Integer  'Utilizado com InStr.
    Dim lRegSyasEdi_EXE    As String    'Registro do arquivo SYASEDI.EXE.
    Dim lRegSyaslibSGS      As String   'Registro do arquivo SYASLIB.SGS.

    '1. Posiciona data de hoje.
    gHoje = CDbl(Format$(Now, "yyyymmdd"))
    lAuxiliar = Format$(gHoje, "00000000")
    gHojDia = CInt(Mid$(lAuxiliar, 7, 2))
    gHojMes = CInt(Mid$(lAuxiliar, 5, 2))
    gHojAno = CInt(Mid$(lAuxiliar, 1, 4))
    gForHoje = Format$(Now, "dd/mm/yyyy")

    '2. Posiciona controle dos endereços dos arquivos (gAppPath).
    If Right$(App.Path, 1) = "\" Then
        gAppPath = App.Path
    Else
        gAppPath = App.Path & "\"
    End If

    '3. Obtém dados do arquivo de inicialização.
    Call gfObtSYASMULTINI(pPropTabela)
    
    If gPrograma = "YASUDA.EXE" Then
        'Sai da procedure para processar o arquivo YASUDA.EXE processar as atualizações.
        Exit Sub
    End If

    '4. Verifica informações.
    'a) Tipo de usuário.
    If Not gfPreenchido(gTipUsuario) Then
        Call gpGraLog(2, "FTE0042 - Erro no arquivo de inicialização (tipo do usuário não preenchido).")
        Exit Sub 'End   ***MOSILVA Verificar
    End If
    If Not IsNumeric(gTipUsuario) Then
        Call gpGraLog(2, "FTE0043 - Erro no arquivo de inicialização (tipo do usuário não numérico).")
        Exit Sub 'End   ***MOSILVA Verificar
    End If
    If gTipUsuario <> e_Tipo_Usu_Corretor And gTipUsuario <> e_Tipo_Usu_Yasuda And gTipUsuario <> e_Tipo_Usu_CorretorYasuda Then
        Call gpGraLog(2, "FTE0044 - Erro no arquivo de inicialização (tipo do usuário desconhecido).")
        Exit Sub 'End   ***MOSILVA Verificar
    End If
    'b) Nome do usuário.
    Select Case gTipUsuario
        Case 1        'Yasuda.
            gNomUsuario = Trim$(GetLoginUserName)
            If Not gfPreenchido(gNomUsuario) Then
                Call gpGraLog(2, "FTE0045 - Identificação do usuário Yasuda não localizada.")
                Exit Sub 'End   ***MOSILVA Verificar
            End If
            If Len(gNomUsuario) > 6 Then
                gNomUsuario6 = Left$(gNomUsuario, 6)
            Else
                gNomUsuario6 = gNomUsuario
            End If
        Case 2        'Corretor.
            lEndSyasEdi_EXE = gAppPath & "SYASEDI.EXE"
            If Not gfPreenchido(Dir(lEndSyasEdi_EXE)) Then
                lEndSyasEdi_EXE = gAppPath & "EAPSCRIPT.INI"
            End If
            If Not gfPreenchido(Dir(lEndSyasEdi_EXE)) Then
                gNomUsuario = "XXXXXXXX"
            Else
                lNumSyasEdi_EXE = FreeFile
                On Error Resume Next
                Open lEndSyasEdi_EXE For Input Access Read As #lNumSyasEdi_EXE
                If Err <> 0 Then
                    Call gpGraLog(2, "FTE0026 - Erro ao abrir arquivo " & lEndSyasEdi_EXE & ": " & _
                                     Err & " - " & Error & vbLf & vbLf & "O sistema será encerrado.")
                    On Error GoTo 0
                    Exit Sub 'End   ***MOSILVA Verificar
                End If
                On Error GoTo 0
                Do While Not EOF(lNumSyasEdi_EXE)
                    Line Input #lNumSyasEdi_EXE, lRegSyasEdi_EXE
                    lRegSyasEdi_EXE = UCase$(lRegSyasEdi_EXE)
                    lPosicao = InStr(lRegSyasEdi_EXE, "SENDER=""")
                    If lPosicao <> 0 Then
                        gNomUsuario = Mid(lRegSyasEdi_EXE, lPosicao + 8, 6)
                        Exit Do
                    End If
                Loop
                Close #lNumSyasEdi_EXE
                If Not gfPreenchido(gNomUsuario) Then
                    gNomUsuario = "XXXXXXXX"
                End If
            End If
            gNomUsuario6 = Left$(gNomUsuario, 6)
        Case 3
            gNomUsuario = Trim$(GetLoginUserName)
            If Not gfPreenchido(gNomUsuario) Then
                Call gpGraLog(2, "FTE0045 - Identificação do usuário Yasuda não localizada.")
                Exit Sub 'End   ***MOSILVA Verificar
            End If
            If Len(gNomUsuario) > 6 Then
                gNomUsuario6 = Left$(gNomUsuario, 6)
            Else
                gNomUsuario6 = gNomUsuario
            End If
            
            'gCodSeguradoYas = "0777702"
            'gCod_Corretor = 1601
            
    End Select
End Sub


Public Sub gfObtSYASMULTINI(ByRef pPropTabela As String)
    'Procedure: obtém dados do arquivo SYASMULT.INI.

    Dim lDado           As String   'Dado da linha de registro.
    Dim lDesconhecido   As Boolean  'Tipo desconhecido.
    Dim lEndBasDados()  As String   'Bases de dados: endereços.
    Dim lEndereco       As String   'Endereco.
    Dim lEndSYASMULTINI As String   'Endereço do arquivo SYASMULT.INI.
    Dim li              As Integer  'Utilizado com For...Next.
    Dim lImaImpressao() As String   'Arquivos de impressão: endereço do arquivo de imagem.
    Dim lNomBasDados()  As String   'Bases de dados: nomes.
    Dim lNumSYASMULTINI As Integer  'Utilizado com FreeFile para o arquivo SYASMULT.INI.
    Dim lPosicao        As Integer  'Utilizado com InStr.
    Dim lQuaBasDados    As Byte     'Bases de dados: quantidade.
    Dim lRegSYASMULTINI As String   'Registro do arquivo SYASMULT.INI.
    Dim lTamanho        As Integer  'Tamanho.
    Dim lTipBasDados()  As String   'Bases de dados: tipos.
    Dim lTipo           As String   'Tipo de linha do arquivo de inicialização.
    Dim lTitulo         As String   'Linha de título.
    Dim lEndSiscad      As String
    Dim lEndP47V001     As String

Static lsVerificado     As Boolean
Static lsPropTabela     As String

    If lsVerificado = True Then
        GoTo Retorno
    Else
        lsVerificado = True
    End If
    
    '1. Abre arquivo SYASMULT.INI.
    lEndSiscad = ""
    lEndP47V001 = ""
    
    lEndSYASMULTINI = gAppPath & "SYASMULT.INI"
    If Not gfPreenchido(Dir(lEndSYASMULTINI)) Then
        Call gpGraLog(2, "FTE0057 - Arquivo não localizado: " & lEndSYASMULTINI & _
                         " (obter dados de inicialização do sistema)." & vbLf & vbLf & _
                         "O sistema será encerrado.")
        GoTo SAIDA 'End   ***MOSILVA Verificar
    End If
    lNumSYASMULTINI = FreeFile
    On Error Resume Next
    Open lEndSYASMULTINI For Input Access Read As lNumSYASMULTINI
    If Err <> 0 Then
        Call gpGraLog(2, "FTE0048 - Erro ao abrir arquivo " & lEndSYASMULTINI & ": " & Err & " - " & _
                         Error & vbLf & vbLf & "O sistema será encerrado.")
        On Error GoTo 0
        GoTo SAIDA 'End   ***MOSILVA Verificar
    End If
    On Error GoTo 0

    '2. Trata registros do arquivo de inicialização.
    lDesconhecido = False
    Do While Not EOF(lNumSYASMULTINI)
        Line Input #lNumSYASMULTINI, lRegSYASMULTINI
        If gfPreenchido(lRegSYASMULTINI) Then
            Select Case Left$(lRegSYASMULTINI, 1)
                Case ";"    'Linha de observação.
                    'Nada a processar - linha de observação.
                Case "["    'Linha de título.
                    lTitulo = Trim$(UCase$(lRegSYASMULTINI))
                Case Else   'Linha de tipo.
                    lRegSYASMULTINI = Trim$(lRegSYASMULTINI)
                    lPosicao = InStr(lRegSYASMULTINI, "=")
                    lTamanho = Len(lRegSYASMULTINI)
                    If lPosicao = 0 Then
                        Call gpGraLog(2, "FTE0049 - Erro no arquivo de inicialização (sinal de igual " & _
                                         "não localizado no registro)." & vbLf & vbLf & _
                                         lRegSYASMULTINI & vbLf & vbLf & "O sistema será encerrado.")
                        GoTo SAIDA 'End   ***MOSILVA Verificar
                    End If
                    If lPosicao = 1 Then
                        Call gpGraLog(2, "FTE0049 - Erro no arquivo de inicialização (sinal de igual " & _
                                         "na primeira posição do registro)." & vbLf & vbLf & _
                                         lRegSYASMULTINI & vbLf & vbLf & "O sistema será encerrado.")
                        GoTo SAIDA 'End   ***MOSILVA Verificar
                    End If
                    If lPosicao = lTamanho Then
                        Call gpGraLog(2, "FTE0049 - Erro no arquivo de inicialização (sinal de igual " & _
                                         "na última posição do registro)." & vbLf & vbLf & _
                                         lRegSYASMULTINI & vbLf & vbLf & "O sistema será encerrado.")
                        GoTo SAIDA 'End   ***MOSILVA Verificar
                    End If
                    lTipo = Trim$(UCase$(Left$(lRegSYASMULTINI, lPosicao - 1)))
                    lDado = Trim$(Mid$(lRegSYASMULTINI, lPosicao + 1))
                    Select Case lTitulo
                        Case "[AUXILIO]"
                        Case "[DBSYASEMIS]"
                            Select Case lTipo
                                Case "TIPO"
                                    gTipSYASEMIS = lDado
                                Case "SERVER"
                                    gSerSYASEMIS = lDado
                                Case "DATABASE"
                                    gEndSYASEMIS = lDado
                                Case "UID"
                                    gUidSYASEMIS = lDado
                                Case "PWD"
                                    gPwdSYASEMIS = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                        Case "[MULTCOT1]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfCot1 = lDado
                                Case "IMAGEM"
                                    gEndImagCot1 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                        Case "[MULTCOT2]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfCot2 = lDado
                                Case "IMAGEM"
                                    gEndImagCot2 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                        Case "[MULTCOT3]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfCot3 = lDado
                                Case "IMAGEM"
                                    gEndImagCot3 = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                        Case "[MULTCOTB]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfCotB = lDado
                                Case "IMAGEM"
                                    gEndImagCotB = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                        Case "[MULTPRO1]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfPro1 = lDado
                                Case "IMAGEM"
                                    gEndImagPro1 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                        Case "[MULTPRO2]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfPro2 = lDado
                                Case "IMAGEM"
                                    gEndImagPro2 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                        Case "[MULTPRO3]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfPro3 = lDado
                                Case "IMAGEM"
                                    gEndImagPro3 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                        Case "[MULTPRO4]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfPro4 = lDado
                                Case "IMAGEM"
                                    gEndImagPro4 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                        Case "[MULTPROB]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfProB = lDado
                                Case "IMAGEM"
                                    gEndImagProB = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                        Case "[MULTPROC]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfProC = lDado
                                Case "IMAGEM"
                                    gEndImagProC = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                        Case "[MULTPROD]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfProD = lDado
                                Case "IMAGEM"
                                    gEndImagProD = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                       Case "[MULTCOTA]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfCotA = lDado
                                Case "IMAGEM"
                                    gEndImagCotA = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                       Case "[MULTCOTC]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfCotC = lDado
                                Case "IMAGEM"
                                    gEndImagCotC = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select


        '*********************************************************
        ' FORMULARIOS ANTIGOS
        '*********************************************************
                       Case "[MULTCOT1_V0]"
                            Select Case lTipo
                                Case "NOME"
                                    gEndCot1_v0 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                       Case "[MULTCOT2_V0]"
                            Select Case lTipo
                                Case "NOME"
                                    gEndCot2_v0 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                       Case "[MULTCOT3_V0]"
                            Select Case lTipo
                                Case "NOME"
                                    gEndCot3_v0 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                       Case "[MULTPRO1_V0]"
                            Select Case lTipo
                                Case "NOME"
                                    gEndPro1_v0 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                       Case "[MULTPRO2_V0]"
                            Select Case lTipo
                                Case "NOME"
                                    gEndPro2_v0 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                       Case "[MULTPRO3_V0]"
                            Select Case lTipo
                                Case "NOME"
                                    gEndPro3_v0 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                       Case "[MULTPRO4_V0]"
                            Select Case lTipo
                                Case "NOME"
                                    gEndPro4_v0 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                       Case "[MULTPROB_V0]"
                            Select Case lTipo
                                Case "NOME"
                                    gEndProB_v0 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
          '*********************************************************
          ' FORMULARIOS ROUBO
          '*********************************************************
                        Case "[ROUBOCOT1]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfRob1 = lDado
                                Case "IMAGEM"
                                    gEndImagRob1 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                        Case "[ROUBOCOT2]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfRob2 = lDado
                                Case "IMAGEM"
                                    gEndImagRob2 = lDado
                                Case Else
                                    lDesconhecido = True
                           End Select
                        Case "[ROUBOCOT3]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfRob3 = lDado
                                Case "IMAGEM"
                                    gEndImagRob3 = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                        Case "[ROUBOCOT4]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfRob4 = lDado
                                Case "IMAGEM"
                                    gEndImagRob4 = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                        Case "[ROUBOCOT5]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfRob5 = lDado
                                Case "IMAGEM"
                                    gEndImagRob5 = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                        Case "[COT_LMI]"
                            Select Case lTipo
                                Case "CONFIG"
                                    'gEndConfRob5 = lDado
                                Case "IMAGEM"
                                    gEndImagCotLMI = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                        Case "[REL_LMI]"
                            Select Case lTipo
                                Case "CONFIG"
                                Case "IMAGEM"
                                    gEndImagRelLMI = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
        '*********************************************************
                        Case "[P00MPED]"
                            Select Case lTipo
                                Case "TIPO"
                                    gTipP00MPED = lDado
                                Case "DATABASE"
                                    gEndP00MPED = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                            gTipP00MULTI = gTipP00MPED
                            gEndP00MULTI = gAppPath & "P00MULTI.MDB"
                        Case "[P00APOL]"
                            Select Case lTipo
                                Case "TIPO"
                                    gTipP00APOL = lDado
                                Case "DATABASE"
                                    gEndP00APOL = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                        Case "[REMESSA]"
                            Select Case lTipo
                                Case "ENDERECO"
                                    gEndRemessa = lDado
                            End Select
                        Case "[RETORNO]"
                        Case "[USO EM REDE]"
                        Case "[USO VERSAO DOS]"
                        Case "[USUARIO]"
                            Select Case lTipo
                                Case "TIPO"
                                    gTipUsuario = lDado
                                Case "NOME"
                                    gNomUsuario = lDado
                                Case "TIPOHOMOLOG"
                                    gTipoHomologacao = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                        Case "[ATUALIZACAO]"
                            gEndSYASUpDt = lDado
                            If gfPreenchido(gEndSYASUpDt) Then
                                lTamanho = Len(gEndSYASUpDt)
                                If lTamanho = 5 Then
                                    gEndSYASUpDt = gAppPath
                                Else
                                    gEndSYASUpDt = Left$(gEndSYASUpDt, lTamanho - 5)
                                End If
                            End If
                        Case "[PROPRIETARIO TABELA]"
                            lsPropTabela = lDado
                        Case "[DBDIVERSOS]"
                            Select Case lTipo
                                Case "TIPO"
                                    gTipDIVERSOS = lDado
                                Case "SERVER"
                                    gSerDIVERSOS = lDado
                                Case "DATABASE"
                                    gEndDIVERSOS = lDado
                                Case "UID"
                                    gUidDIVERSOS = lDado
                                Case "PWD"
                                    gPwdDIVERSOS = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                        Case "[DBGED]"
                            Select Case lTipo
                                Case "TIPO"
                                    gTipGed = lDado
                                Case "SERVER"
                                    gSerGed = lDado
                                Case "DATABASE"
                                    gEndGed = lDado
                                Case "UID"
                                    gUidGed = lDado
                                Case "PWD"
                                    gPwdGed = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                        Case "[DBEAPDB]"
                            Select Case lTipo
                                Case "TIPO"
                                    gTipEAPDB = lDado
                                Case "SERVER"
                                    gSerEAPDB = lDado
                                Case "DATABASE"
                                    gEndEAPDB = lDado
                                Case "UID"
                                    gUidEAPDB = lDado
                                Case "PWD"
                                    gPwdEAPDB = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                        Case "[DBSYASPORTAL]"
                            Select Case lTipo
                                Case "TIPO"
                                    gTipSyasPortal = lDado
                                Case "SERVER"
                                    gSerSyasPortal = lDado
                                Case "DATABASE"
                                    gEndSyasPortal = lDado
                                Case "UID"
                                    gUidSyasPortal = lDado
                                Case "PWD"
                                    gPwdSyasPortal = lDado
                                Case Else
                                    lDesconhecido = True
                            End Select
                        Case "[PATH_ARQUIVOS]"
                            gPathArquivos = lDado
                        Case "[IMPRESSORAFRENTE]"
                            gImprFrente = lDado
                        Case "[IMPRESSORAFRENTEVERSO]"
                            gImprFrenteVerso = lDado
                        Case "[SISCAD]"
                            lEndSiscad = lDado
                        Case "[P47V001]"
                            lEndP47V001 = lDado
                        Case "[DIR_IST]"
                            gDir_Ist = lDado
                        Case "[CONFIG]"
                            Select Case lTipo
                                Case "MAPI_SERVER_EXCHANGE"
                                    MAPI_SERVER_EXCHANGE = lDado
                                Case "MAPI_CONTA_EXCHANGE"
                                    MAPI_CONTA_EXCHANGE = lDado
                                Case "MAPI_Profile"
                                   MAPI_Profile = lDado
                            End Select
                        Case "[CAMINHOPI]"
                            gPathImagePI = lDado
                    End Select
                    If lDesconhecido = True Then
                        Call gpGraLog(2, "FTE0049 - Erro no arquivo de inicialização (tipo " & _
                                         "desconhecido para título " & lTitulo & ")." & vbLf & vbLf & _
                                         lRegSYASMULTINI & vbLf & vbLf & "O sistema será encerrado.")
                        GoTo SAIDA 'End   ***MOSILVA Verificar
                    End If
            End Select
        End If
    Loop
    Close lNumSYASMULTINI
    
    gBaseLocal = True
    If gTipUsuario = e_Tipo_Usu_Corretor Or gTipP00MPED = 1 Then   'corretor
        gBaseLocal = True
    'ElseIf gTipP00MPED = 2 And gPerfilAcesso <> e_Perfil_DIG_IST Then
    Else
        gBaseLocal = False
    End If
    
    
    If gTipUsuario <> e_Tipo_Usu_CorretorYasuda Then
        lsPropTabela = ""
    End If
    
    If gTipUsuario = e_Tipo_Usu_CorretorYasuda And gfPreenchido(lsPropTabela) = False Then
        Call gpGraLog(2, "FTE0049 - Erro no arquivo de inicialização (Proprietário " & _
                                         "das tabelas não informado)." & vbLf & vbLf & _
                                         lRegSYASMULTINI & vbLf & vbLf & "O sistema será encerrado.")
    End If
    
    '4. Finaliza a procedure para iniciar procedimento de atualização do SYAS com os arquivos
    '   SYASxxnn.003. Processo iniciado apenas se o módulo estiver sendo executado pelo programa
    '   YASUDA.EXE.
    If gfPreenchido(gPrograma) Then
        If gPrograma = "YASUDA.EXE" Then
            'Encerra procedure para processar atualizações.
            GoTo SAIDA
        End If
    End If

    If gEndConfCot1 = "MULTCOT1.CXY" Then
        gEndConfCot1 = gAppPath & gEndConfCot1
    End If
    If gEndImagCot1 = "MULTCOT1.WMF" Then
        gEndImagCot1 = gAppPath & gEndImagCot1
    End If
    If gEndConfCotA = "MULTCOTA.CXY" Then
        gEndConfCotA = gAppPath & gEndConfCotA
    End If
    If gEndImagCotA = "MULTCOTA.WMF" Then
        gEndImagCotA = gAppPath & gEndImagCotA
    End If
    If gEndConfPro1 = "MULTPROP1.CXY" Then
        gEndConfPro1 = gAppPath & gEndConfPro1
    End If
    If gEndImagPro1 = "MULTPROP1.WMF" Then
        gEndImagPro1 = gAppPath & gEndImagPro1
    End If

    If gEndConfPro2 = "MULTPROP2.CXY" Then
        gEndConfPro2 = gAppPath & gEndConfPro2
    End If
    If gEndImagPro2 = "MULTPROP2.WMF" Then
        gEndImagPro2 = gAppPath & gEndImagPro2
    End If
    
    If gfPreenchido(gEndConfRob1) = False Then
        gEndConfRob1 = Replace(gEndConfCot1, "MULTCOT1.CXY", "") & "ROUBOCOT1.CXY"
        gEndConfRob2 = Replace(gEndConfCot1, "MULTCOT1.CXY", "") & "ROUBOCOT2.CXY"
        gEndConfRob3 = Replace(gEndConfCot1, "MULTCOT1.CXY", "") & "ROUBOCOT3.CXY"
        gEndConfRob4 = Replace(gEndConfCot1, "MULTCOT1.CXY", "") & "ROUBOCOT4.CXY"
        gEndConfRob5 = Replace(gEndConfCot1, "MULTCOT1.CXY", "") & "ROUBOCOT5.CXY"
    End If
    If gfPreenchido(gEndImagRob1) = False Then
        gEndImagRob1 = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "ROUBOCOT1.WMF"
        gEndImagRob2 = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "ROUBOCOT2.WMF"
        gEndImagRob3 = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "ROUBOCOT3.WMF"
        gEndImagRob4 = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "ROUBOCOT4.WMF"
        gEndImagRob5 = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "ROUBOCOT5.WMF"
    End If
   
    
    If gfPreenchido(gEndConfCotC) = False Then
        gEndConfCotC = Replace(gEndConfCotA, "MULTCOTA.CXY", "") & "MULTCOTC.CXY"
        gEndImagCotC = Replace(gEndImagCotA, "MULTCOTA.WMF", "") & "MULTCOTC.WMF"
    End If

    If gfPreenchido(gEndImagCotLMI) = False Then
        'gEndConfCotLMI = Replace(gEndConfCot1, "MULTCOT1.CXY", "") & "COT_LMI.CXY"
        gEndImagCotLMI = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "COT_LMI.WMF"
        gEndImagRelLMI = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "REL_LMI.WMF"
    End If

    If gfPreenchido(gEndImagESPLMI) = False Then
        gEndImagESPLMI = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "MULTLMIE.WMF"
        gEndImagESPRelLMI = Replace(gEndImagCot1, "MULTCOT1.WMF", "") & "MULTLMIR.WMF"
    End If
    
    '5. Verifica bases de dados.
    'a) Posiciona quantidade de bases de dados.
    lQuaBasDados = 2    '<<<<<<====== Verificar quantidade de bases de dados a verificar.
    'b) Posiciona nomes, tipos e endereços das bases de dados.
    ReDim lNomBasDados(lQuaBasDados)
    ReDim lEndBasDados(lQuaBasDados)
    ReDim lTipBasDados(lQuaBasDados)
    lNomBasDados(1) = "P00MPED"
    lEndBasDados(1) = gEndP00MPED
    lTipBasDados(1) = gTipP00MPED      'Observar variável lQuaBasDados.
    lNomBasDados(2) = "P00APOL"
    lEndBasDados(2) = gEndP00APOL
    lTipBasDados(2) = gTipP00APOL      'Observar variável lQuaBasDados.
    'b) Processa verificação.
    If gUsuMulti = False Then
        For li = 1 To lQuaBasDados
            Do
                '===1) Verifica tipo da base de dados.
                If Not gfPreenchido(lTipBasDados(li)) Then
                    Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(li) & _
                                     ": tipo da base de dados não preenchido." & vbLf & vbLf & _
                                     "O sistema será encerrado.")
                    
                    GoTo SAIDA 'End   ***MOSILVA Verificar
                End If
                If lTipBasDados(li) <> "1" And lTipBasDados(li) <> "2" Then
                    Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(li) & _
                                     ": tipo da base de dados diferente de 1 - " & lTipBasDados(li) & "." & _
                                     vbLf & vbLf & "O sistema será encerrado.")
                    GoTo SAIDA 'End   ***MOSILVA Verificar
                End If
                '===2) Verifica endereço do arquivo.
                If gfPreenchido(lEndBasDados(li)) Then
                    Select Case UCase(lEndBasDados(li))
                        Case "P00MULTI.MDB"
                            lEndereco = gAppPath & lEndBasDados(li)
                            gEndP00MULTI = lEndereco
                        Case "P00MPED.MDB"
                            lEndereco = gAppPath & lEndBasDados(li)
                            gEndP00MPED = lEndereco
                        Case "P00APOL.MDB"
                            lEndereco = gAppPath & lEndBasDados(li)
                            gEndP00APOL = lEndereco
                        Case Else
                            lEndereco = lEndBasDados(li)
                    End Select
                Else
                    lEndereco = gAppPath & lNomBasDados(li) & ".MDB"
                End If
                If UCase(lNomBasDados(li)) = "P0044000" And gTipUsuario = 1 Then
                    Exit Do
                End If
                If gBaseLocal Then
                    If Not gfPreenchido(Dir(lEndereco)) Then
                        Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(li) & _
                                         ": base de dados não localizada." & vbLf & vbLf & lEndereco & _
                                         vbLf & vbLf & "O sistema será encerrado.")
                        GoTo SAIDA 'End   ***MOSILVA Verificar
                    End If
                End If
                'lPosicao = InStr(lEndereco, lNomBasDados(lI) & ".MDB")
                'If lPosicao = 0 Then
                '    Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(lI) & _
                '                     ": nome da base de dados não consta no endereço." & vbLf & vbLf & _
                '                     lEndereco & vbLf & vbLf & "O sistema será encerrado.")
                '    End
                'End If
                Exit Do
            Loop
        Next li
    End If
    
    '5. Verifica pasta remessa.
    If Not gfPreenchido(gEndRemessa) Then
        gEndRemessa = gAppPath & "REMESSA"
    End If
    If Right$(gEndRemessa, 1) = "\" Then
        gEndRemessa = Left$(gEndRemessa, Len(gEndRemessa) - 1)
    End If
    If Not gfPreenchido(Dir(gEndRemessa, vbDirectory)) Then
        On Error Resume Next
        MkDir (gEndRemessa)
        If Err <> 0 Then
            Call gpGraLog(0, "FTE0015 - Erro ao criar pasta " & gEndRemessa & _
                             " (pasta de remessa obter dados de inicialização do sistema).")
        End If
    End If
    gEndRemessa = gEndRemessa & "\"
    
    Call mfObtSYASPRODINI
    
    If lEndSiscad <> "" Then
        gEndExeSiscad = lEndSiscad
    End If
    
    If lEndP47V001 <> "" Then
        gEndExeP47V001 = lEndP47V001
    End If
    
Retorno:
    pPropTabela = lsPropTabela
SAIDA:
    Exit Sub
    
End Sub

Private Function mfObtSYASPRODINI() As Boolean
    'Função: obtém dados do arquivo SYASPROD.INI.

    Dim lDado           As String   'Dado da linha de registro.
    Dim lEndSyasprodINI As String   'Endereço do arquivo SYASPROD.INI.
    Dim lNumSyasprodINI As Integer  'Utilizado com FreeFile para o arquivo SYASPROD.INI.
    Dim lPosicao        As Integer  'Utilizado com InStr.
    Dim lRegSyasprodINI As String   'Registro do arquivo SYASPROD.INI.
    Dim lTamanho        As Integer  'Tamanho.
    Dim lTipo           As String   'Tipo de linha do arquivo de inicialização.
    Dim lTitulo         As String   'Linha de título.
    Dim lEndBasDados()  As String   'Bases de dados: endereços.
    Dim lEndereco       As String   'Endereco.
    Dim li              As Integer  'Utilizado com For...Next.
    Dim lNomBasDados()  As String   'Bases de dados: nomes.
    Dim lQuaBasDados    As Byte     'Bases de dados: quantidade.
    Dim lTipBasDados()  As String   'Bases de dados: tipos.
    
    '1. Abre arquivo SYASPROD.INI.
    lEndSyasprodINI = gAppPath & "SYASPROD.INI"
    If Not gfPreenchido(Dir(lEndSyasprodINI)) Then
        Call gpGraLog(2, "FTE0025 - Arquivo não localizado: " & lEndSyasprodINI & _
                         " (obter dados de inicialização do sistema)." & vbLf & vbLf & _
                         "O sistema será encerrado.")
        Exit Function 'End   ***MOSILVA Verificar
    End If
    lNumSyasprodINI = FreeFile
    On Error Resume Next
    Open lEndSyasprodINI For Input Access Read As lNumSyasprodINI
    If Err <> 0 Then
        Call gpGraLog(2, "FTE0026 - Erro ao abrir arquivo " & lEndSyasprodINI & ": " & Err & " - " & _
                         Error & vbLf & vbLf & "O sistema será encerrado.")
        On Error GoTo 0
        Exit Function 'End   ***MOSILVA Verificar
    End If
    On Error GoTo 0

    '2. Trata registros do arquivo de inicialização.
    Do While Not EOF(lNumSyasprodINI)
        Line Input #lNumSyasprodINI, lRegSyasprodINI
        If gfPreenchido(lRegSyasprodINI) Then
            Select Case Left$(lRegSyasprodINI, 1)
                Case ";"    'Linha de observação.
                    'Nada a processar - linha de observação.
                Case "["    'Linha de título.
                    lTitulo = Trim$(UCase$(lRegSyasprodINI))
                Case Else   'Linha de tipo.
                    lRegSyasprodINI = Trim$(lRegSyasprodINI)
                    lPosicao = InStr(lRegSyasprodINI, "=")
                    lTamanho = Len(lRegSyasprodINI)
                    If lPosicao = 0 Then
                        Call gpGraLog(2, "FTE0046 - Erro no arquivo de inicialização (sinal de igual " & _
                                         "não localizado no registro)." & vbLf & vbLf & _
                                         lRegSyasprodINI & vbLf & vbLf & "O sistema será encerrado.")
                        Exit Function 'End   ***MOSILVA Verificar
                    End If
                    If lPosicao = 1 Then
                        Call gpGraLog(2, "FTE0046 - Erro no arquivo de inicialização (sinal de igual " & _
                                         "na primeira posição do registro)." & vbLf & vbLf & _
                                         lRegSyasprodINI & vbLf & vbLf & "O sistema será encerrado.")
                        Exit Function 'End   ***MOSILVA Verificar
                    End If
                    If lPosicao = lTamanho Then
                        Call gpGraLog(2, "FTE0046 - Erro no arquivo de inicialização (sinal de igual " & _
                                         "na última posição do registro)." & vbLf & vbLf & _
                                         lRegSyasprodINI & vbLf & vbLf & "O sistema será encerrado.")
                        Exit Function 'End   ***MOSILVA Verificar
                    End If
                    lTipo = Trim$(UCase$(Left$(lRegSyasprodINI, lPosicao - 1)))
                    lDado = Trim$(Mid$(lRegSyasprodINI, lPosicao + 1))
                    Select Case lTitulo
                        Case "[L28V009A]"
                            Select Case lTipo
                                Case "CONFIG"
                                    gEndConfDeb = lDado
                                Case "IMAGEM"
                                    gEndImagDeb = lDado
                            End Select
                        Case "[P0043700]"
                            Select Case lTipo
                                Case "DATABASE"
                                    If gfPreenchido(lDado) Then
                                        lTamanho = Len(lDado)
                                        If lTamanho = 12 Then
                                            gEndServidor = gAppPath
                                        Else
                                            gEndServidor = Left$(lDado, lTamanho - 12)
                                        End If
                                    End If
                            End Select
                        Case "[SISCAD]"
                            gEndExeSiscad = lDado
                        Case "[P47V001]"
                            gEndExeP47V001 = lDado
                        Case "[P0043000]"
                            Select Case lTipo
                                Case "TIPO"
                                    gTipP0043000 = lDado
                                Case "DATABASE"
                                    gEndP0043000 = lDado
                                    If UCase(gEndP0043000) = "P0043000.MDB" Then
                                        gEndP0043000 = gAppPath & "P0043000.MDB"
                                    End If
                            End Select
                        Case "[P0044700]"
                            Select Case lTipo
                                Case "TIPO"
                                    gTipP0044700 = lDado
                                Case "DATABASE"
                                    gEndP0044700 = lDado
                                    If UCase(gEndP0044700) = "P0044700.MDB" Then
                                        gEndP0044700 = gAppPath & "P0044700.MDB"
                                    End If
                            End Select
                        Case "[P0045200]"
                            Select Case lTipo
                                Case "TIPO"
                                    gTipP0045200 = lDado
                                Case "DATABASE"
                                    gEndP0045200 = lDado
                                    If UCase(gEndP0045200) = "P0045200.MDB" Then
                                        gEndP0045200 = gAppPath & "P0045200.MDB"
                                    End If
                            End Select
                        Case "[P0045300]"
                            Select Case lTipo
                                Case "TIPO"
                                    gTipP0045300 = lDado
                                Case "DATABASE"
                                    gEndP0045300 = lDado
                                    If UCase(gEndP0045300) = "P0045300.MDB" Then
                                        gEndP0045300 = gAppPath & "P0045300.MDB"
                                    End If
                            End Select
                        Case "[P0045400]"
                            Select Case lTipo
                                Case "TIPO"
                                    gTipP0045400 = lDado
                                Case "DATABASE"
                                    gEndP0045400 = lDado
                                    If UCase(gEndP0045400) = "P0045400.MDB" Then
                                        gEndP0045400 = gAppPath & "P0045400.MDB"
                                    End If
                            End Select
                        Case "[ATUALIZACAO]"
                            gEndSYASUpDt = lDado
                            
                            If gfPreenchido(gEndSYASUpDt) Then
                                lTamanho = Len(gEndSYASUpDt)
                                If lTamanho = 5 Then
                                    gEndSYASUpDt = gAppPath
                                Else
                                    gEndSYASUpDt = Left$(gEndSYASUpDt, lTamanho - 5)
                                End If
                            End If
                        Case "[TIPOTRANSMISSAO]"
                            If Trim(UCase(lDado)) = "OFFLINE" Then
                                gTipoTransmDLL = ETransmDLL_OFFLine
                            Else
                                gTipoTransmDLL = ETransmDLL_ONLine
                            End If
                    End Select
            End Select
        End If
    Loop
    Close lNumSyasprodINI
    If gPrograma = "D046O003" Then
        mfObtSYASPRODINI = True
        Exit Function
    End If
    If gfPreenchido(gEndP0044700) Then
        lTamanho = Len(gEndP0044700)
        If lTamanho = 12 Then
            lEndereco = ""
        Else
            lEndereco = Left$(gEndP0044700, lTamanho - 12)
        End If
    Else
        lEndereco = gAppPath
    End If
    
    If Not gfPreenchido(gEndP0045200) Then
        gEndP0045200 = lEndereco & "P0045200.MDB"
        gTipP0045200 = "1"
    End If
    If Not gfPreenchido(gEndP0045300) Then
        gEndP0045300 = lEndereco & "P0045300.MDB"
        gTipP0045300 = "1"
    End If
    
    If Not gfPreenchido(gEndP0045400) Then
        gEndP0045400 = lEndereco & "P0045400.MDB"
        gTipP0045400 = "1"
    End If
    
    '5. Verifica bases de dados.
    'a) Posiciona quantidade de bases de dados.
    lQuaBasDados = 2    '<<<<<<====== Verificar quantidade de bases de dados a verificar.
    'b) Posiciona nomes, tipos e endereços das bases de dados.
    ReDim lNomBasDados(lQuaBasDados)
    ReDim lEndBasDados(lQuaBasDados)
    ReDim lTipBasDados(lQuaBasDados)
    lNomBasDados(1) = "P0044700"
    lEndBasDados(1) = gEndP0044700
    lTipBasDados(1) = gTipP0044700  'Observar variável lQuaBasDados.
    lNomBasDados(2) = "P0043000"
    lEndBasDados(2) = gEndP0043000
    lTipBasDados(2) = gTipP0043000      'Observar variável lQuaBasDados.
    
    
    
    'lNomBasDados(3) = "P0045200"
    'lEndBasDados(3) = gEndP0045200
    'lTipBasDados(3) = gTipP0045200      'Observar variável lQuaBasDados.
    'lNomBasDados(4) = "P0045300"
    'lEndBasDados(4) = gEndP0045300
    'lTipBasDados(4) = gTipP0045300      'Observar variável lQuaBasDados.
    
    
    'b) Processa verificação.
    If gUsuMulti = False Then
        For li = 1 To lQuaBasDados
            Do
                '===1) Verifica tipo da base de dados.
                If Not gfPreenchido(lTipBasDados(li)) Then
                    Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(li) & _
                                     ": tipo da base de dados não preenchido." & vbLf & vbLf & _
                                     "O sistema será encerrado.")
                    
                    Exit Function 'End   ***MOSILVA Verificar
                End If
                If lTipBasDados(li) <> "1" And lTipBasDados(li) <> "2" Then
                    Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(li) & _
                                     ": tipo da base de dados diferente de 1 - " & lTipBasDados(li) & "." & _
                                     vbLf & vbLf & "O sistema será encerrado.")
                    Exit Function 'End   ***MOSILVA Verificar
                End If
                '===2) Verifica endereço do arquivo.
                If gfPreenchido(lEndBasDados(li)) Then
                    Select Case UCase(lEndBasDados(li))
                        Case "P0044700.MDB"
                            lEndereco = gAppPath & lEndBasDados(li)
                            gEndP0044700 = lEndereco
                        Case "P0043000.MDB"
                            lEndereco = gAppPath & lEndBasDados(li)
                            gEndP0043000 = lEndereco
                        Case "P0045200.MDB"
                            lEndereco = gAppPath & lEndBasDados(li)
                            gEndP0045200 = lEndereco
                        Case "P0045300.MDB"
                            lEndereco = gAppPath & lEndBasDados(li)
                            gEndP0045300 = lEndereco
                        Case "P0045400.MDB"
                            lEndereco = gAppPath & lEndBasDados(li)
                            gEndP0045400 = lEndereco
                        Case Else
                            lEndereco = lEndBasDados(li)
                    End Select
                Else
                    lEndereco = gAppPath & lNomBasDados(li) & ".MDB"
                End If
                If UCase(lNomBasDados(li)) = "P0044000" And gTipUsuario = 1 Then
                    Exit Do
                End If
                If Not gfPreenchido(Dir(lEndereco)) Then
                    Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(li) & _
                                     ": base de dados não localizada." & vbLf & vbLf & lEndereco & _
                                     vbLf & vbLf & "O sistema será encerrado.")
                    Exit Function 'End   ***MOSILVA Verificar
                End If
                'lPosicao = InStr(lEndereco, lNomBasDados(lI) & ".MDB")
                'If lPosicao = 0 Then
                '    Call gpGraLog(2, "FTE0050 - Erro ao verificar base de dados " & lNomBasDados(lI) & _
                '                     ": nome da base de dados não consta no endereço." & vbLf & vbLf & _
                '                     lEndereco & vbLf & vbLf & "O sistema será encerrado.")
                '    End
                'End If
                Exit Do
            Loop
        Next li
    End If
                                
End Function
Public Function gfTransmissaoInterna(ByVal pTipoTransmissao As Integer, _
                                     ByRef pobjstcPed As stcA46V702B, _
                                     ByRef pMensagem As String, _
                                     Optional pCnSQL2008 As ADODB.Connection = Nothing, _
                                     Optional pRetornoEhErro As Boolean = True) As Boolean
                                     
Dim lObjstcPed          As stcA46V702B
Dim lobjPedido          As clsA46V702A
Dim lObjstcPedClau      As stcA46V703B
Dim lobjPedClau         As clsA46V703A
Dim lobjstcPedCob       As stcA46V704B
Dim lobjPedCob          As clsA46V704A
Dim lobjstcPedCor       As stcA46V705B
Dim lobjPedCor          As clsA46V705A
Dim lobjstcPedCos       As stcA46V706B
Dim lobjPedCos          As clsA46V706A

Dim lobjPedEnder        As clsA46V707A
Dim lobjstcPedLoc       As stcA46V708B
Dim lObjPedLoc          As clsA46V708A
Dim lobjstcPedJuros     As stcA46V711B
Dim lobjPedJuros        As clsA46V711A

Dim lobjCtrlEmis        As clsA46V712A
Dim lObjstcPedVenc      As stcA46V713B
Dim lObjPedVenc         As clsA46V713A

Dim lobjPedObs          As clsA46V714A
Dim lobjstcPedLocBem    As stcA46V715B
Dim lobjPedLocBem       As clsA46V715A
Dim lobjstcPedLocSeg    As stcA46V716B
Dim lobjPedLocSeg       As clsA46V716A
Dim lobjPedInsp         As clsA46V717A
Dim lobjstcPedFat       As stcA46V720B
Dim lobjPedFat          As clsA46V720A
Dim lobjstcPedEspec     As stcA46V726B
Dim lobjPedEspec        As clsA46V726A


Dim lobjTrans           As clsA46V718A
Dim lobjstcTrans        As stcA46V718B
                                    
Dim lConexaoSQL         As ADODB.Connection
Dim lExec               As String
Dim lPropTabela         As String
Dim lNomeTabImport      As String

On Error GoTo ControleErro

    lPropTabela = "RamosDiversos.dbo."
    If pTipoTransmissao = 1 Then 'MOSILVA: Questionar Thais!!
        lNomeTabImport = "FTP_"
    Else
        lNomeTabImport = ""
    End If
    
    'Abre a base SQL2000 de Produção (LPX em caso de teste Abre MIX)
    If Not gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lConexaoSQL, pMensagem) Then
        GoTo ControleErro
    End If

    Set lobjPedido = New clsA46V702A: Set lobjPedido.ConexaoBD = lConexaoSQL
    Set lObjstcPed = New stcA46V702B
    lObjstcPed.Nosso_Numero = pobjstcPed.Nosso_Numero
    lObjstcPed.Cod_Unicidade = pobjstcPed.Cod_Unicidade

    If Not lobjPedido.mfCarregarColecao(lObjstcPed, lPropTabela, lNomeTabImport, False) Then
        pMensagem = lobjPedido.UltimoErro
        GoTo ControleErro
    End If

    If lobjPedido.Colecao.Count <> 0 Then
        Set lObjstcPed = lobjPedido.Colecao.Item(1)
        If Mid(lObjstcPed.Num_PI, 3, 1) = 2 Then
            pMensagem = "Proposta já em processo de emissão. Não será possível realizar a transmissão."
            pRetornoEhErro = False
            GoTo ControleErro
        End If
        Call gpExcluiRegistros(pobjstcPed.Nosso_Numero, 0, lPropTabela, lNomeTabImport, pMensagem)
    End If
    
    Set lobjPedido = New clsA46V702A: Set lobjPedido.ConexaoBD = lConexaoSQL
    
    'Ajusta os campos
    pobjstcPed.Num_Cta_Bas_Cor = Left(Format(pobjstcPed.Num_Cta_Bas_Cor, "00000000"), 7)
    
    '
    If Not lobjPedido.mfIncluir(pobjstcPed, lPropTabela, lNomeTabImport) Then
        pMensagem = lobjPedido.UltimoErro
        GoTo ControleErro
    End If
    
    Set lobjPedCor = New clsA46V705A
    Set lobjPedCor.ConexaoBD = lConexaoSQL
    For Each lobjstcPedCor In pobjstcPed.ColecaoPedCorr
        Set lobjstcPedCor.ObjPai = pobjstcPed
        If Not lobjPedCor.mfIncluir(lobjstcPedCor, lPropTabela, lNomeTabImport) Then
            pMensagem = lobjPedCor.UltimoErro
            GoTo ControleErro
        End If
    Next lobjstcPedCor
    
    Set lobjPedCos = New clsA46V706A
    Set lobjPedCos.ConexaoBD = lConexaoSQL
    For Each lobjstcPedCos In pobjstcPed.ColecaoPedCoss
        Set lobjstcPedCos.ObjPai = pobjstcPed
        If Not lobjPedCos.mfIncluir(lobjstcPedCos, lPropTabela, lNomeTabImport) Then
            pMensagem = lobjPedCos.UltimoErro
            GoTo ControleErro
        End If
    Next lobjstcPedCos

    Set lobjPedEnder = New clsA46V707A
    Set lobjPedEnder.ConexaoBD = lConexaoSQL
    pobjstcPed.Cod_Unicidade = pobjstcPed.Cod_Unicidade
    Set pobjstcPed.objstcPedEndereco.ObjPai = pobjstcPed
    If Not lobjPedEnder.mfIncluir(pobjstcPed.objstcPedEndereco, lPropTabela, lNomeTabImport) Then
        pMensagem = lobjPedEnder.UltimoErro
        GoTo ControleErro
    End If
    
    
    Set lobjCtrlEmis = New clsA46V712A
    Set lobjCtrlEmis.ConexaoBD = lConexaoSQL
    Set pobjstcPed.objstcCtrlEmis.ObjPai = pobjstcPed
    If Not lobjCtrlEmis.mfIncluir(pobjstcPed.objstcCtrlEmis, lPropTabela, lNomeTabImport) Then
        pMensagem = lobjCtrlEmis.UltimoErro
        GoTo ControleErro
    End If
    
    Set lobjPedJuros = New clsA46V711A
    Set lobjPedJuros.ConexaoBD = lConexaoSQL
    For Each lobjstcPedJuros In pobjstcPed.ColecaoPedJuros
        Set lobjstcPedJuros.ObjPai = pobjstcPed
        If Not lobjPedJuros.mfIncluir(lobjstcPedJuros, lPropTabela, lNomeTabImport) Then
            pMensagem = lobjPedJuros.UltimoErro
            GoTo ControleErro
        End If
    Next lobjstcPedJuros
    
    Set lObjPedVenc = New clsA46V713A
    Set lObjPedVenc.ConexaoBD = lConexaoSQL
    For Each lObjstcPedVenc In pobjstcPed.ColecaoPedVenc
        Set lObjstcPedVenc.ObjPai = pobjstcPed
        If Not lObjPedVenc.mfIncluir(lObjstcPedVenc, lPropTabela, lNomeTabImport) Then
            pMensagem = lObjPedVenc.UltimoErro
            GoTo ControleErro
        End If
    Next lObjstcPedVenc
    
    Set lobjPedEspec = New clsA46V726A
    Set lobjPedEspec.ConexaoBD = lConexaoSQL
    For Each lobjstcPedEspec In pobjstcPed.ColecaoTextoEspec
        Set lobjstcPedEspec.ObjPai = pobjstcPed
        If Not lobjPedEspec.mfIncluir(lobjstcPedEspec, lPropTabela, lNomeTabImport) Then
            pMensagem = lobjPedEspec.UltimoErro
            GoTo ControleErro
        End If
    Next lobjstcPedEspec
        

    Set lObjPedLoc = New clsA46V708A
    Set lObjPedLoc.ConexaoBD = lConexaoSQL
    
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        Set lobjstcPedLoc.ObjPai = pobjstcPed
        If Not lObjPedLoc.mfIncluir(lobjstcPedLoc, lPropTabela, lNomeTabImport) Then
            pMensagem = lObjPedLoc.UltimoErro
            GoTo ControleErro
        End If
            
        Set lobjPedClau = New clsA46V703A
        Set lobjPedClau.ConexaoBD = lConexaoSQL
        
        For Each lObjstcPedClau In lobjstcPedLoc.ColecaoPedClau
            Set lObjstcPedClau.ObjPai = lobjstcPedLoc
            If Not lobjPedClau.mfIncluir(lObjstcPedClau, lPropTabela, lNomeTabImport) Then
                pMensagem = lobjPedClau.UltimoErro
                GoTo ControleErro
            End If
        Next lObjstcPedClau
        
        Set lobjPedCob = New clsA46V704A
        Set lobjPedCob.ConexaoBD = lConexaoSQL
        For Each lobjstcPedCob In lobjstcPedLoc.ColecaoPedCobert
            Set lobjstcPedCob.ObjPai = lobjstcPedLoc
            If Not lobjPedCob.mfIncluir(lobjstcPedCob, lPropTabela, lNomeTabImport) Then
                pMensagem = lobjPedCob.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedCob
        
        Set lobjPedObs = New clsA46V714A
        Set lobjPedObs.ConexaoBD = lConexaoSQL
        If gfPreenchido(lobjstcPedLoc.objstcPedObs.Observacao) Then
            Set lobjstcPedLoc.objstcPedObs.ObjPai = lobjstcPedLoc
            If Not lobjPedObs.mfIncluir(lobjstcPedLoc.objstcPedObs, lPropTabela, lNomeTabImport) Then
                pMensagem = lobjPedObs.UltimoErro
                GoTo ControleErro
            End If
        End If
        
        Set lobjPedLocBem = New clsA46V715A
        Set lobjPedLocBem.ConexaoBD = lConexaoSQL
        For Each lobjstcPedLocBem In lobjstcPedLoc.ColecaoPedLocBem
            Set lobjstcPedLocBem.ObjPai = lobjstcPedLoc
            If Not lobjPedLocBem.mfIncluir(lobjstcPedLocBem, lPropTabela, lNomeTabImport) Then
                pMensagem = lobjPedLocBem.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedLocBem
        
        Set lobjPedLocSeg = New clsA46V716A
        Set lobjPedLocSeg.ConexaoBD = lConexaoSQL
        For Each lobjstcPedLocSeg In lobjstcPedLoc.ColecaoPedSeguro
            Set lobjstcPedLocSeg.ObjPai = lobjstcPedLoc
            If Not lobjPedLocSeg.mfIncluir(lobjstcPedLocSeg, lPropTabela, lNomeTabImport) Then
                pMensagem = lobjPedLocSeg.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedLocSeg
        

        Set lobjPedInsp = New clsA46V717A
        Set lobjPedInsp.ConexaoBD = lConexaoSQL
        Set lobjstcPedLoc.objstcPedInsp.ObjPai = lobjstcPedLoc
        If Not lobjPedInsp.mfIncluir(lobjstcPedLoc.objstcPedInsp, lPropTabela, lNomeTabImport) Then
            pMensagem = lobjPedInsp.UltimoErro
            GoTo ControleErro
        End If
        
        Set lobjPedFat = New clsA46V720A
        Set lobjPedFat.ConexaoBD = lConexaoSQL
        For Each lobjstcPedFat In lobjstcPedLoc.ColecaoPedFator
            Set lobjstcPedFat.ObjPai = lobjstcPedLoc
            If Not lobjPedFat.mfIncluir(lobjstcPedFat, lPropTabela, lNomeTabImport) Then
                pMensagem = lobjPedFat.UltimoErro
                GoTo ControleErro
            End If
        Next lobjstcPedFat
    Next lobjstcPedLoc

    gfTransmissaoInterna = True

SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    gfTransmissaoInterna = False
    
    If Not gfPreenchido(pMensagem) Then
        pMensagem = Err.Description
    End If
    GoTo SAIDA
    
FechaObj:
    Set lObjstcPed = Nothing
    Set lobjPedido = Nothing
    Set lObjstcPedClau = Nothing
    Set lobjPedClau = Nothing
    Set lobjstcPedCob = Nothing
    Set lobjPedCob = Nothing
    Set lobjstcPedCor = Nothing
    Set lobjPedCor = Nothing
    Set lobjstcPedCos = Nothing
    Set lobjPedCos = Nothing
    Set lobjPedEnder = Nothing
    Set lobjstcPedLoc = Nothing
    Set lObjPedLoc = Nothing
    Set lobjstcPedJuros = Nothing
    Set lobjPedJuros = Nothing
    Set lobjCtrlEmis = Nothing
    Set lObjstcPedVenc = Nothing
    Set lObjPedVenc = Nothing
    Set lobjPedObs = Nothing
    Set lobjstcPedLocBem = Nothing
    Set lobjPedLocBem = Nothing
    Set lobjstcPedLocSeg = Nothing
    Set lobjPedLocSeg = Nothing
    Set lobjPedInsp = Nothing
    Set lobjstcPedFat = Nothing
    Set lobjPedFat = Nothing
    Set lobjstcPedEspec = Nothing
    Set lobjPedEspec = Nothing
    Set lobjTrans = Nothing
    Set lobjstcTrans = Nothing
    Set lConexaoSQL = Nothing
    Return
End Function

