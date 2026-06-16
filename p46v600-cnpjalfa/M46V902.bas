Attribute VB_Name = "basM46V902"
Option Explicit
Dim lSelect As String
Public Function gfImpDocumento_Ver_0(ByVal pOpcao As Byte, ByVal pFormulario As String, _
                                     ByVal pImagem As String, ByVal pobjstcPed As stcA46V702B, _
                                     pobjstcPedLoc As stcA46V708B, pPagDe As Integer, _
                                     pPagAte As Integer) As Boolean
'Função: imprime documento.

'Parâmetros de entrada...pOpcao...........0 = visualizar impressão.
'                                         1 = imprimir.
'                                         2 = criar arquivo imagem.
'                        pFormulario......Nome do formulário.
'                        pImagem..........Nome para o arquivo imagem a ser criado.
'                        lObjP0042200.....Registros de item.
    Dim lVerifOverf As Boolean
Retorno:
    Dim lArqDados As String         'Endereço do arquivo de dados.
    Dim lEndConfig As String          'Endereço do arquivo de configuração.
    Dim lEndImagem As String          'Endereço do arquivo de imagem.
    Dim lNumArquivo As Integer          'Número do arquivo.
    Dim lRefConfig As String          'Número de referência do arquivo de configuração.
    Dim lRefImagem As String          'Número de referência do arquivo de imagem.
    Dim lResultado As Boolean         'Resultado das funções de criação de dados.
    Dim lPossuiObs As Boolean         'Possui observação?
    Dim lobjstcLoc As stcA46V708B     'Objeto de itens.
    Dim lNome As String     'Auxiliar

    gfImpDocumento_Ver_0 = False

    lEndConfig = ""
    lEndImagem = ""

    '1. carrego o endereço do formulário.
    Select Case UCase$(pFormulario)
    Case "MULTCOT1"
        lNome = gM46V999.gEndCot1_v0
    Case "MULTCOT2"
        lNome = gM46V999.gEndCot2_v0
    Case "MULTCOT3"
        lNome = gM46V999.gEndCot3_v0
    Case "MULTPRO1"
        lNome = gM46V999.gEndPro1_v0
    Case "MULTPRO2"
        lNome = gM46V999.gEndPro2_v0
    Case "MULTPRO3"
        lNome = gM46V999.gEndPro3_v0
    Case "MULTPRO4"
        lNome = gM46V999.gEndPro4_v0
    Case "L28V009A"
        lEndConfig = gM46V999.gEndConfDeb
        lEndImagem = gM46V999.gEndImagDeb
    Case "MULTPROB"
        lNome = gM46V999.gEndProB_v0
    Case Else
        Call gM46V999.gpGraLog(1, "FTE0080 - Código de formulário desconhecido: " & pFormulario & ".")
        Exit Function
    End Select

    If lEndConfig = "" Then
        lEndConfig = lNome & ".CXY"
        lEndImagem = lNome & ".WMF"
    End If


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
    '    If gM46V999.gTipUsuario = 1 Then
    '        lArqDados = "C:\SYAS.DDD"
    '    Else
    '        lArqDados = gM46V999.gAppPath & "SYAS.DDD"
    '    End If

    If gfIsMetaframe() Then
        lArqDados = "L:\SYAS.DDD"
        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & lArqDados)
    Else
        lArqDados = gM46V999.gAppPath & "SYAS.DDD"
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
                                " - " & Error & " - " & lArqDados & ").")
        Exit Function
    End Select
    On Error GoTo 0

    '7. Grava registros no arquivo de dados.
    Select Case pFormulario
    Case "MULTCOT1"                 'Cotação simples.
        lResultado = mfCriDados_MULTCOT1_V_0(lNumArquivo, pobjstcPed, pobjstcPedLoc)
    Case "MULTCOT2"                 'Cotação coletiva - capa.
        lResultado = mfCriDados_MULTCOT2_V_0(lNumArquivo, pobjstcPed, pobjstcPedLoc)
    Case "MULTCOT3"                 'Cotação coletiva - dados do item.
        lResultado = mfCriDados_MULTCOT3_V_0(lNumArquivo, pobjstcPed, pPagDe, pPagAte)
    Case "MULTPRO1"                 'Proposta simples - página 1.
        lResultado = mfCriDados_MULTPRO1_V_0(lNumArquivo, pobjstcPed, pobjstcPedLoc)
    Case "MULTPRO2"                 'Proposta simples - página 2.
        lResultado = mfCriDados_MULTPRO2_V_0(lNumArquivo, pobjstcPed, pobjstcPedLoc)
    Case "MULTPRO3"                 'Proposta coletiva - capa.
        lResultado = mfCriDados_MULTPRO3_V_0(lNumArquivo, pobjstcPed, pobjstcPedLoc, pPagAte)
    Case "MULTPRO4"                 'Proposta coletiva - dados do item.
        lResultado = mfCriDados_MULTPRO4_V_0(lNumArquivo, pobjstcPed, pPagDe, pPagAte)
    Case "MULTPROB"
        lResultado = mfCriDados_MULTPROB_V_0(lNumArquivo, pobjstcPed)
    Case "L28V009A"                 'Proposta coletiva - dados do item.
        lResultado = mfCriDados_Debito(lNumArquivo, pobjstcPed)
    Case Else
        Call gM46V999.gpGraLog(1, " FTE0082 - Erro ao imprimir documento (código de formulário desconhecido - " & pFormulario & ").")
        Close #lNumArquivo
        Kill lArqDados
        Exit Function
    End Select


    Close #lNumArquivo
    If lResultado = False Then
        Kill lArqDados
        Exit Function
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
    Call PrintReld(CStr(pOpcao))        '<<<<<<<<<<========== Comando que apresentou erro - Dr. Watson...
    If Err <> 0 Then
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
        gfImpDocumento_Ver_0 = True
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

End Function

Private Function mfCriDados_MULTCOT1_V_0(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                         pobjstcPedLoc As stcA46V708B) As Boolean
'Função: cria arquivo de dados - cotação empresarial simples.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.


    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lCodCobertura As Integer    'Código da cobertura.
    Dim lDemParcelas As String    'Demais parcelas.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lJuros As String    'Juros.
    Dim lOrdem As String    'Ordem.
    Dim lPriParcela As String   'Primeira parcela.
    Dim lTamanho As Integer    'Tamanho.
    Dim lValTotCober As Double    'Valor soma das coberturas.
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lobjstcCob As stcA46V704B    'Objeto da classe (stc) de cobertura.
    Dim lobjstcSeg As stcA46V716B    'Outra seguradora.
    Dim lobjstcJur As stcA46V711B    'Objeto da classe (stc) de juros.
    Dim lMensagem As String
    Dim lValorParc As Variant
    Dim lDescFranquias() As String
    Dim lDescImportante() As String
    mfCriDados_MULTCOT1_V_0 = False

    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado YASUDA RESIDENCIAL, YASUDA CONDOMÍNIO ou YASUDA EMPRESARIAL.
    lAuxiliar2 = "YASUDA " & UCase(pobjstcPed.gDesc_Ramo)
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '002. Emissão.
    Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

    '003. Validade.
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
        Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
    Else
        Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
    End If

    '004. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '005. Número do item .
    lAuxiliar1 = pobjstcPedLoc.Num_Item
    Print #pNumArqDados, Format$(lAuxiliar1, "0000")

    '006. Corretor (código e nome) .
    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
    lAuxiliar1 = lObjstcCorr.Cod_Corr
    lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '007. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '008. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '009. Proponente - nome do segurado .
    lAuxiliar2 = pobjstcPed.Nom_Segurado
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '010. Proponente - código do segurado - se houver .
    lAuxiliar1 = pobjstcPed.Cod_Segurado
    Print #pNumArqDados, Format$(lAuxiliar1, "0000000")

    '011. Proponente - tipo de pessoa .
    Select Case pobjstcPed.Tip_Pessoa
    Case e_TipPessoa_Juridica
        Print #pNumArqDados, "Jurídica"
        '012 e 013. Proponente - CNPJ ou CPF .
        lAuxiliar2 = "CNPJ.........:"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     "Ramo de atividade: " & pobjstcPed.gDesc_RamoAtiv
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Case e_TipPessoa_Fisica
        Print #pNumArqDados, "Física"
        '012 e 013. Proponente - CNPJ ou CPF .
        lAuxiliar2 = "CPF...........:"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     pobjstcPed.gCod_Doc_Segur & ": " & pobjstcPed.Num_Doc & "   " & _
                     "Órgão emissor: " & pobjstcPed.Org_Exp & "   " & _
                     "Emissão: " & gM46V999.gfForData(2, pobjstcPed.Dat_Emis_Doc)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End Select

    '014. Proponente - endereço .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Cod_Tip_Logr & " " & _
                 pobjstcPed.objstcPedEndereco.Nom_Logr & "," & _
                 pobjstcPed.objstcPedEndereco.Dsc_Compl_Logr
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '015. Proponente - bairro .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Nom_Bairro
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '016. Proponente - cidade .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Nom_Cid
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '017. Proponente - UF .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Sig_UF
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '018. Proponente - CEP .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.gCEP_Formatado
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '019. Proponente - Telefone .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Num_Fone
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '020. Local em risco - logradouro .
    lAuxiliar2 = pobjstcPedLoc.Cod_Tip_Logr_Risco & " " & pobjstcPedLoc.Nom_Logr_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '021. Local em risco - número .
    lAuxiliar2 = pobjstcPedLoc.objstcPedInsp.Num_Logr_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '022. Local em risco - complemento .
    lAuxiliar2 = pobjstcPedLoc.Dsc_Compl_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '023. Local em risco - bairro .
    lAuxiliar2 = pobjstcPedLoc.Nom_Bairro_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '024. Local em risco - cidade .
    lAuxiliar2 = pobjstcPedLoc.Nom_Cid_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '025. Local em risco - UF .
    lAuxiliar2 = pobjstcPedLoc.Sig_UF_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '026. Local em risco - CEP .
    lAuxiliar2 = pobjstcPedLoc.gCEP_Formatado
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '027. Dados para inspeção - nome da pessoa para contato .
    lAuxiliar2 = pobjstcPedLoc.objstcPedInsp.Pess_Cont
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '028. Dados para inspeção - telefone para contato .
    lAuxiliar2 = pobjstcPedLoc.objstcPedInsp.Num_Tel
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '029. Dados do seguro - classificação .
    lAuxiliar2 = pobjstcPedLoc.Cod_Clasf & "-" & pobjstcPedLoc.Cod_Compl_Clasf & "- " & pobjstcPedLoc.Dsc_Clasf
    lAuxiliar2 = lAuxiliar2 & String(15, " ") & " PLANO: " & pobjstcPed.Desc_Plano
    lAuxiliar2 = lAuxiliar2 & " - PRODUTO: " & Format(pobjstcPedLoc.Cod_Produto, "0000")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '030. Dados do seguro - construção .
    'Ficha 1000305
    If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        Select Case pobjstcPedLoc.Cod_Constr
        Case e_CodContr_Superior
            lAuxiliar2 = "Superior/sólida"
        Case e_CodContr_Inferior
            lAuxiliar2 = "Inferior"
        Case e_CodContr_Mista
            lAuxiliar2 = "Mista"
        Case Else
            lAuxiliar2 = "Superior/sólida"
        End Select
    Else
        Select Case pobjstcPedLoc.Cod_Constr
        Case e_CodContr_Superior
            lAuxiliar2 = "Superior/sólida"
        Case e_CodContr_Inferior
            lAuxiliar2 = "Inferior/mista"
        Case Else
            lAuxiliar2 = "Superior/sólida"
        End Select
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '031. Dados do seguro - região .
    lAuxiliar2 = pobjstcPedLoc.Cod_Regiao & " - " & Left(pobjstcPedLoc.gDesc_Regiao, 38)
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '032. Dados do seguro - vigência .
    lAuxiliar2 = "de " & gM46V999.gfForData(2, pobjstcPed.Dat_Ini_Vig) & _
               " a " & gM46V999.gfForData(2, pobjstcPed.Dat_Ter_Vig)
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '033. Dados do seguro - valor em risco .
    lAuxiliar2 = "R$ " & Format(pobjstcPedLoc.Val_Vr, "standard")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '034. Dados do seguro - tipo de emissão .
    lAuxiliar2 = pobjstcPedLoc.gDesc_TipEmissao
    Select Case pobjstcPedLoc.Tip_Emissao
    Case e_TipEmissao_RenYasuda Or e_TipEmissao_Endosso
        lAuxiliar2 = lAuxiliar2 & " " & "  Apólice: " & pobjstcPedLoc.Num_Apol_Anterior & "  Item: " & pobjstcPedLoc.Num_Item_Anterior & "  Sinistro: " & IIf(pobjstcPedLoc.Ind_Sinistro = e_SimNao_Sim, "Sim", "Não")
    End Select
    Print #pNumArqDados, lAuxiliar2

    If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        '035 e 036. Coberturas contratadas - título.
        lAuxiliar2 = "Desconto de"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = "experiência"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Else
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
    End If

    '037 a 168. Coberturas contratadas - linhas de 01 a 22.
    lCountLine = 0
    lValTotCober = 0
    For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
        lCountLine = lCountLine + 1
        'Código
        Print #pNumArqDados, Format$(lobjstcCob.Cod_Cobert, "0000")
        'Descrição
        Print #pNumArqDados, Left(lobjstcCob.gNomeCobertura, 48)

        'Valor da IS
        lAuxiliar1 = lobjstcCob.Val_IS
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        'd) Desconto comercial .
        lAuxiliar1 = lobjstcCob.Per_Desc_Coml
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 10 Then
            lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            'e) Desconto de experiência .
            lAuxiliar1 = lobjstcCob.Per_Desc_Exp
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 10 Then
                lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Else
            Print #pNumArqDados, " "
        End If

        'f) Prêmio líquido .
        lAuxiliar1 = lobjstcCob.VAL_PR
        lValTotCober = lValTotCober + lobjstcCob.VAL_PR
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Next lobjstcCob
    Do While lCountLine < 22
        lCountLine = lCountLine + 1
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
    Loop

    '169. Coberturas contratadas - soma.
    lAuxiliar1 = lValTotCober
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '170. Coberturas contratadas - prêmio líquido.
    lAuxiliar1 = pobjstcPed.VAL_PR
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '171. Coberturas contratadas - custo da apólice.
    lAuxiliar1 = pobjstcPed.Val_Cust_Emis
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '172. Coberturas contratadas - IOF.
    lAuxiliar1 = pobjstcPed.Val_IOF
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '173. Coberturas contratadas - prêmio total a vista.
    lAuxiliar1 = pobjstcPed.gValorTotal
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '174. Informação cifrada .

    '   B: 25 / A: 50 / P: 10 / CO: 99999 / C: 99999 Onde:
    '   B = bônus
    '   A = aparelho
    '   P = progressivo
    '   CO = coeficiente de comissão
    '   C = coeficiente LMG/VR

    'Ficha 900740 e 901021 - zerar bônus das renovações dos planinhos
    '2001 a 2010 - Correta
    '2011 a 2020 - GalCorr
    '2021 a 2031 - Mercedes
    '2032 a 2041 - National
    '2042 a 2049 - Manaus
    '2050 - Ballarin
    '2051 a 2060 - Marsh
    '2061 a 2070 - NSK
    '2071 a 2080 - Ajinomoto
    '2081 a 2090 - Nissin
    '2091 a 2100 - Panasonic
    '2101 - Fanel
    '2102 - JLA

    'Ficha 901673
    '2011 a 2020 - Embraer
    'Ficha 901772
    '2121 a 2130

    If (gM46V111.gCod_plano >= 2001 And gM46V111.gCod_plano <= 2049) Or (gM46V111.gCod_plano >= 2051 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
        lAuxiliar2 = "B:00"
    Else
        lAuxiliar2 = "B:" & Format(pobjstcPedLoc.Per_Bonus, "00")
    End If

    lAuxiliar2 = lAuxiliar2 & "/A:" & Format(pobjstcPedLoc.Per_Desc_Apar, "00")

    'Ficha 900740 e 901673 e 901772
    If (gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031) Or (gM46V111.gCod_plano >= 2061 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
        lAuxiliar2 = lAuxiliar2 & "/P:00"
    Else
        lAuxiliar2 = lAuxiliar2 & "/P:" & Format(pobjstcPedLoc.Per_Desc_Progres, "00")
    End If

    'lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(pObjstcPed.objstcCtrlEmis.Per_Corr + pObjstcPed.objstcCtrlEmis.Per_Corr_2 + pObjstcPed.objstcCtrlEmis.Per_Corr_3, "00.00")
    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
    lAuxiliar1 = lObjstcCorr.Per_Corr

    If pobjstcPed.ColecaoPedCorr.Count > 1 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(2)
        lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
        If pobjstcPed.ColecaoPedCorr.Count > 2 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(3)
            lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
        End If
    End If

    lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(lAuxiliar1 * 100, "0000")

    lAuxiliar2 = lAuxiliar2 & "/C:" & Format(Round(pobjstcPedLoc.Val_Coef_IS_Vr, 4) * 10000, "00000")
    lAuxiliar2 = lAuxiliar2 & "/PRBO19"

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '175 a 193. Segurança.
    '           O texto pode ser customizado de acordo com a necessidade.
    '           Temos 19 linhas disponíveis e cada linha cabe 30 caracteres (contando com o "(X)").
    '           O "X" deve ser colocado de acordo com a base de dados.
    'a) Linha 01.
    If pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Then
        Print #pNumArqDados, "SEGURANÇA"
        If pobjstcPedLoc.Cod_Iden_Segur_01_V = "S" Then
            lAuxiliar2 = "(X) "
        Else
            lAuxiliar2 = "( ) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Vigilância exclusiva e"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'b) Linha 02.
        lAuxiliar2 = "    permanente do imóvel"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'c) Linha 03.
        lAuxiliar2 = "    entendendo-se ainda"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'd) Linha 04.
        lAuxiliar2 = "    como tal, porteiros"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'e) Linha 05.
        lAuxiliar2 = "    fixos."
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'f) Linha 06.
        If pobjstcPedLoc.Cod_Iden_Segur_02_A = "S" Then
            lAuxiliar2 = "(X) "
        Else
            lAuxiliar2 = "( ) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Alarmes."
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'g) Linha 07.
        If pobjstcPedLoc.Cod_Iden_Segur_03_P = "S" Then
            lAuxiliar2 = "(X) "
        Else
            lAuxiliar2 = "( ) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Porteiros eletrônicos."
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'h) Linha 08.
        If pobjstcPedLoc.Cod_Iden_Segur_04_I = "S" Then
            lAuxiliar2 = "(X) "
        Else
            lAuxiliar2 = "( ) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Inexistência de terreno"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'i) Linha 09.
        lAuxiliar2 = "    baldio, imóvel desocupado"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'j) Linha 10.
        lAuxiliar2 = "    ou em construção na"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'k) Linha 11.
        lAuxiliar2 = "    vizinhança do risco."
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'l) Linha 12.
        If pobjstcPedLoc.Cod_Iden_Segur_05_G = "S" Then
            lAuxiliar2 = "(X) "
        Else
            lAuxiliar2 = "( ) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Grades metálicas de"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm) Linha 13.
        lAuxiliar2 = "    proteção em todas as"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'n) Linha 14.
        lAuxiliar2 = "    janelas e portas de vidro"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'o) Linha 15.
        lAuxiliar2 = "    que dão acesso ao"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'p) Linha 16.
        lAuxiliar2 = "    interior do imóvel desde"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'q) Linha 17.
        lAuxiliar2 = "    que também possuam"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'r) Linha 18.
        lAuxiliar2 = "    alarme."
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        's) Linha 19.
        lAuxiliar2 = " "
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Else
        Print #pNumArqDados, " "
        For lI = 1 To 19
            Print #pNumArqDados, String(30, "X")
        Next lI
    End If
    '194 a 198. Proteção.
    '           Para residencial, sairá em branco.
    If pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Or pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        '           O texto pode ser customizado de acordo com a necessidade.
        '           Temos 4 linhas disponíveis e cada linha cabe 39 caracteres (contando com o "(X)").
        '           O "X" deve ser colocado de acordo com a base de dados.
        'a) Título.
        lAuxiliar2 = "PROTEÇÃO"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'b) Linha 01.
        If pobjstcPedLoc.Cod_Iden_Protec_01_E = "S" Then
            lAuxiliar2 = "(X) "
        Else
            lAuxiliar2 = "( ) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Extintor"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'c) Linha 02.
        If pobjstcPedLoc.Cod_Iden_Protec_02_H = "S" Then
            lAuxiliar2 = "(X) "
        Else
            lAuxiliar2 = "( ) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Hidrantes"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'd) Linha 03.
        If pobjstcPedLoc.Cod_Iden_Protec_03_S = "S" Then
            lAuxiliar2 = "(X) "
        Else
            lAuxiliar2 = "( ) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Sprinkler (chuveiros automáticos)"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'e) Linha 04.
        If pobjstcPedLoc.Cod_Iden_Protec_04_D = "S" Then
            lAuxiliar2 = "(X) "
        Else
            lAuxiliar2 = "( ) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Detecção e alarme"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Else
        'a) Título.
        Print #pNumArqDados, " "
        For lI = 1 To 4
            Print #pNumArqDados, String(38, "X")
        Next lI
    End If

    'Calcula o fracionamento
    With pobjstcPed
        Call gM46V998.gfCalFracionamento(pobjstcPed, pobjstcPedLoc.Tip_Emissao, _
                                         lValorParc, lMensagem, gM46V111.gProJurIOF, .Cod_Reg, .Val_IOF, .Val_Parc_Min)
    End With
    '199 a 210. Formas de pagamento - primeira parcela a vista.

    'Ficha 900665
    For lI = 1 To gM46V999.gProQuaMaxParcelas    '<- originalmente 12
        If lI > pobjstcPed.Qtd_Parc Then  '<- originalmente 12
            Print #pNumArqDados, " "
        Else
            If lValorParc(lI, 1) <> 0 Then
                'a) Ordem da parcela.
                lOrdem = Format$(lI, "00")
                'b) Valor da primeira parcela.
                lAuxiliar1 = lValorParc(lI, 1)
                If lAuxiliar1 <> 0 Then
                    lPriParcela = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lPriParcela)
                    If lTamanho < 9 Then
                        lPriParcela = Space(9 - lTamanho) & lPriParcela
                    End If
                    'c) Valor das demais parcelas.
                    lAuxiliar1 = lValorParc(lI, 2)
                    lDemParcelas = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lDemParcelas)
                    If lTamanho < 10 Then
                        lDemParcelas = Space(10 - lTamanho) & lDemParcelas
                    End If
                    'd) Juros.
                    Set lobjstcJur = New stcA46V711B
                    On Error Resume Next
                    Set lobjstcJur = pobjstcPed.ColecaoPedJuros.item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
                    On Error GoTo 0
                    If gM46V999.gfPreenchido(lobjstcJur.NOSSO_NUMERO) Then
                        lAuxiliar1 = lobjstcJur.Val_Ant
                    Else
                        lAuxiliar1 = 0
                    End If
                    lJuros = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lJuros)
                    If lTamanho < 5 Then
                        lJuros = Space(5 - lTamanho) & lJuros & "%"
                    End If
                    lAuxiliar2 = lOrdem & "    " & lPriParcela & "  " & lDemParcelas & "  " & lJuros
                    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                Else
                    Print #pNumArqDados, " "
                End If
            Else
                Print #pNumArqDados, " "
            End If
        End If
    Next lI

    '211 a 222. Formas de pagamento - primeira parcela a 30 dias.

    'Ficha 900665
    For lI = 1 To gM46V999.gProQuaMaxParcelas  '<- originalmente 12
        If lI > pobjstcPed.Qtd_Parc Then  '<- originalmente 12
            Print #pNumArqDados, " "
        Else
            If lValorParc(lI, 3) <> 0 Then
                'a) Ordem da parcela.
                lOrdem = Format$(lI, "00")
                'b) Valor da primeira parcela.
                lAuxiliar1 = lValorParc(lI, 3)
                If lAuxiliar1 <> 0 Then
                    lPriParcela = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lPriParcela)
                    If lTamanho < 9 Then
                        lPriParcela = Space(9 - lTamanho) & lPriParcela
                    End If
                    'c) Valor das demais parcelas.
                    lAuxiliar1 = lValorParc(lI, 4)
                    lDemParcelas = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lDemParcelas)
                    If lTamanho < 10 Then
                        lDemParcelas = Space(10 - lTamanho) & lDemParcelas
                    End If
                    'd) Juros.
                    Set lobjstcJur = New stcA46V711B
                    On Error Resume Next
                    Set lobjstcJur = pobjstcPed.ColecaoPedJuros.item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
                    On Error GoTo 0
                    If gM46V999.gfPreenchido(lobjstcJur.NOSSO_NUMERO) Then
                        lAuxiliar1 = lobjstcJur.Val_Pos
                    Else
                        lAuxiliar1 = 0
                    End If
                    lJuros = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lJuros)
                    If lTamanho < 5 Then
                        lJuros = Space(5 - lTamanho) & lJuros & "%"
                    End If
                    lAuxiliar2 = lOrdem & "    " & lPriParcela & "  " & lDemParcelas & "  " & lJuros
                    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                Else
                    Print #pNumArqDados, " "
                End If
            Else
                Print #pNumArqDados, " "
            End If
        End If
    Next lI

    '223 a 230. Franquias.
    Call mpDescricaoFranquia_V_0(pobjstcPedLoc, pobjstcPed, lDescFranquias, 8)

    '    If lDescFranquias(0) <> "" Then
    '        'a) Linha 1.
    '        Print #pNumArqDados, "Por evento, serão deduzidas das indenizações a seguinte participação do " & _
             '                         "segurado:"
    '    Else
    '        Print #pNumArqDados, " "
    '    End If
    'Print #pNumArqDados, " "

    For lI = 0 To 7
        lAuxiliar2 = lDescFranquias(lI)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Next lI

    '231 a 240. Importante.
    Call mpTextoImportante(pobjstcPedLoc, lDescImportante, "MULTCOT1")
    lCountLine = 0
    For lI = 0 To UBound(lDescImportante) - 1
        lAuxiliar2 = lDescImportante(lCountLine)
        If gM46V999.gfPreenchido(lAuxiliar2) Then
            lCountLine = lCountLine + 1
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    Next lI
    Call mpTextoObs(pobjstcPedLoc, lDescImportante)
    For lI = 0 To UBound(lDescImportante) - 1
        lAuxiliar2 = lDescImportante(lI)
        If gM46V999.gfPreenchido(lAuxiliar2) Then
            lCountLine = lCountLine + 1
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    Next lI

    'Alterado em 23/03/2007 - Ficha: 0700328
    If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        lCountLine = lCountLine + 1
        lAuxiliar2 = "-Bens de Terceiro em Poder do Segurado: Para ser passível de cobertura, deverá ser destacado do Valor em Risco"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        lCountLine = lCountLine + 1
        lAuxiliar2 = " Total o Valor em Risco destes bens"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'Alterado em 24/04/2007 - Ficha: 0700328
        If (pobjstcPedLoc.Cod_Clasf = gM46V999.gfBuscarDePara(18, 2) And pobjstcPedLoc.Cod_Compl_Clasf = gM46V999.gfBuscarDePara(18, 2, 10)) Or _
           (pobjstcPedLoc.Cod_Clasf = gM46V999.gfBuscarDePara(18, 2) And pobjstcPedLoc.Cod_Compl_Clasf = gM46V999.gfBuscarDePara(18, 2, 11)) Or _
           (pobjstcPedLoc.Cod_Clasf = gM46V999.gfBuscarDePara(19, 2) And pobjstcPedLoc.Cod_Compl_Clasf = gM46V999.gfBuscarDePara(19, 2, 10)) Then
            lCountLine = lCountLine + 1
            lAuxiliar2 = " ******    RISCOS NOMEADOS - RAMO 96    ******"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    End If

    If gM46V999.gCodProduto >= 958 And pobjstcPedLoc.Cob_Bens = True And Not mfVerDigBens() Then
        lCountLine = 0
        For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
            If gM46V999.gCod_Ramo = e_CodRamo_Empresarial And lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(94, 1) Then
                lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
                lAuxiliar2 = "A relação poderá ser preenchida diretamente no SYAS ou ser entregue juntamente com a proposta de seguro."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
            End If
            If gM46V999.gCod_Ramo = e_CodRamo_Residencial And lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(27, 1) Or _
               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(29, 1) Or _
               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(30, 1) Or _
               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(31, 1) Then
                lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
                lAuxiliar2 = "A relação poderá ser preenchida diretamente no SYAS ou ser entregue juntamente com a proposta de seguro."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
            End If
        Next lobjstcCob
        Do While lCountLine < 4
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        Loop
        GoTo FimTexto
    End If

    Do While lCountLine < 10
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
    Loop

FimTexto:
    mfCriDados_MULTCOT1_V_0 = True
End Function

Private Function mfCriDados_MULTCOT2_V_0(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                         pobjstcPedLoc As stcA46V708B) As Boolean
'Função: cria arquivo de dados - cotação coletiva - capa.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.


    Dim lAuxiliar1 As Double    'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String    'Auxiliar - utilizado apenas como exemplo.
    Dim lDemParcelas As String    'Demais parcelas.
    Dim lI As Integer             'Utilizado com For...Next.
    Dim lJuros As String    'Juros.
    Dim lOrdem As String    'Ordem.
    Dim lPriParcela As String    'Primeira parcela.
    Dim lTamanho As Integer    'Tamanho.
    Dim lObjstcCorr As stcA46V705B
    Dim lobjstcJur As stcA46V711B
    Dim lobjstcCob As stcA46V704B    'Objeto da classe (stc) de cobertura.
    Dim lCountLine As Integer
    Dim lValorParc As Variant
    Dim lMensagem As String
    Dim lDescImportante() As String

    mfCriDados_MULTCOT2_V_0 = False

    '001. Título do documento.
    Print #pNumArqDados, "COTAÇÃO DE SEGURO"

    '002. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado YASUDA RESIDENCIAL, YASUDA CONDOMÍNIO ou YASUDA EMPRESARIAL.
    Print #pNumArqDados, "YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

    '003. Número da página.
    Print #pNumArqDados, "Página 1/" & pobjstcPed.ColecaoPedLoc.Count + 1

    '004. Emissão.
    Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

    '005. Validade.
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
        Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
    Else
        Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
    End If

    '006. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '007. Corretor (código e nome) .
    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
    lAuxiliar1 = lObjstcCorr.Cod_Corr
    lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '008. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '009. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '010. Proponente - nome do segurado .
    lAuxiliar2 = pobjstcPed.Nom_Segurado
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '011. Proponente - código do segurado - se houver .
    lAuxiliar1 = pobjstcPed.Cod_Segurado
    Print #pNumArqDados, Format$(lAuxiliar1, "0000000")

    '012. Proponente - tipo de pessoa .
    Select Case pobjstcPed.Tip_Pessoa
    Case e_TipPessoa_Fisica
        Print #pNumArqDados, "Física"
        '013 e 014. Proponente - CNPJ ou CPF .
        lAuxiliar2 = "CPF...........:"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     pobjstcPed.gCod_Doc_Segur & ": " & pobjstcPed.Num_Doc & "   " & _
                     "Órgão emissor: " & pobjstcPed.Org_Exp & "   " & _
                     "Emissão: " & gM46V999.gfForData(2, pobjstcPed.Dat_Emis_Doc)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Case e_TipPessoa_Juridica
        Print #pNumArqDados, "Jurídica"
        '013 e 014. Proponente - CNPJ ou CPF .
        lAuxiliar2 = "CNPJ.........:"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     "Ramo de atividade: " & pobjstcPed.gDesc_RamoAtiv
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Case Else
        Print #pNumArqDados, "Não informado"
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
    End Select

    '015. Proponente - endereço .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Cod_Tip_Logr & " " & _
                 pobjstcPed.objstcPedEndereco.Nom_Logr & "," & _
                 pobjstcPed.objstcPedEndereco.Dsc_Compl_Logr
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '016. Proponente - bairro .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Nom_Bairro
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '017. Proponente - cidade .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Nom_Cid
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '018. Proponente - UF .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Sig_UF
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '019. Proponente - CEP .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.gCEP_Formatado
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '020. Proponente - Telefone .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Num_Fone
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '021. Resumo: quantidade de itens.
    Print #pNumArqDados, pobjstcPed.ColecaoPedLoc.Count

    '022. Resumo: prêmio líquido.
    lAuxiliar1 = pobjstcPed.VAL_PR
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '023. Resumo: custo da apólice.
    lAuxiliar1 = pobjstcPed.Val_Cust_Emis
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '024. Resumo: IOF.
    lAuxiliar1 = pobjstcPed.Val_IOF
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '025. Resumo: prêmio total a vista.
    lAuxiliar1 = pobjstcPed.gValorTotal
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    'Chama o fracionamento
    'Calcula o fracionamento
    With pobjstcPed
        Call gM46V998.gfCalFracionamento(pobjstcPed, pobjstcPedLoc.Tip_Emissao, _
                                         lValorParc, lMensagem, gM46V111.gProJurIOF, .Cod_Reg, .Val_IOF, .Val_Parc_Min)
    End With
    '026 a 037. Formas de pagamento - primeira parcela a vista.

    'Ficha 900665
    For lI = 1 To gM46V999.gProQuaMaxParcelas  '<- originalmente 12
        If lI > pobjstcPed.Qtd_Parc Then  '<- originalmente 12
            Print #pNumArqDados, " "
        Else
            If lValorParc(lI, 1) <> 0 Then
                'a) Ordem da parcela.
                lOrdem = Format$(lI, "00")
                'b) Valor da primeira parcela.
                lAuxiliar1 = lValorParc(lI, 1)
                lPriParcela = Format$(lAuxiliar1, "standard")
                lTamanho = Len(lPriParcela)
                If lTamanho < 10 Then
                    lPriParcela = Space(10 - lTamanho) & lPriParcela
                End If
                'c) Valor das demais parcelas.
                lAuxiliar1 = lValorParc(lI, 2)
                lDemParcelas = Format$(lAuxiliar1, "standard")
                lTamanho = Len(lDemParcelas)
                If lTamanho < 10 Then
                    lDemParcelas = Space(10 - lTamanho) & lDemParcelas
                End If
                'd) Juros.
                Set lobjstcJur = New stcA46V711B
                On Error Resume Next
                Set lobjstcJur = pobjstcPed.ColecaoPedJuros.item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
                On Error GoTo 0
                If gM46V999.gfPreenchido(lobjstcJur.NOSSO_NUMERO) Then
                    lAuxiliar1 = lobjstcJur.Val_Ant
                Else
                    lAuxiliar1 = 0
                End If
                lJuros = Format$(lAuxiliar1, "standard")
                lTamanho = Len(lJuros)
                If lTamanho < 5 Then
                    lJuros = Space(5 - lTamanho) & lJuros & "%"
                End If
                lAuxiliar2 = lOrdem & "    " & lPriParcela & "  " & lDemParcelas & "  " & lJuros
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            Else
                Print #pNumArqDados, " "
            End If
        End If
    Next lI

    '038 a 049. Formas de pagamento - primeira parcela a 30 dias.

    'Ficha 900665
    For lI = 1 To gM46V999.gProQuaMaxParcelas  '<- originalmente 12
        If lI > pobjstcPed.Qtd_Parc Then  '<- originalmente 12
            Print #pNumArqDados, " "
        Else
            If lValorParc(lI, 3) <> 0 Then
                'a) Ordem da parcela.
                lOrdem = Format$(lI, "00")
                'b) Valor da primeira parcela.
                lAuxiliar1 = lValorParc(lI, 3)
                lPriParcela = Format$(lAuxiliar1, "standard")
                lTamanho = Len(lPriParcela)
                If lTamanho < 10 Then
                    lPriParcela = Space(10 - lTamanho) & lPriParcela
                End If
                'c) Valor das demais parcelas.
                lAuxiliar1 = lValorParc(lI, 4)
                lDemParcelas = Format$(lAuxiliar1, "standard")
                lTamanho = Len(lDemParcelas)
                If lTamanho < 10 Then
                    lDemParcelas = Space(10 - lTamanho) & lDemParcelas
                End If
                'd) Juros.
                Set lobjstcJur = New stcA46V711B
                On Error Resume Next
                Set lobjstcJur = pobjstcPed.ColecaoPedJuros.item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
                On Error GoTo 0
                If gM46V999.gfPreenchido(lobjstcJur.NOSSO_NUMERO) Then
                    lAuxiliar1 = lobjstcJur.Val_Pos
                Else
                    lAuxiliar1 = 0
                End If
                lJuros = Format$(lAuxiliar1, "standard")
                lTamanho = Len(lJuros)
                If lTamanho < 5 Then
                    lJuros = Space(5 - lTamanho) & lJuros & "%"
                End If
                lAuxiliar2 = lOrdem & "    " & lPriParcela & "  " & lDemParcelas & "  " & lJuros
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            Else
                Print #pNumArqDados, " "
            End If
        End If
    Next lI

    '050 a 059. Importante.
    Call mpTextoImportante(pobjstcPedLoc, lDescImportante, "MULTCOT2")
    lCountLine = 0
    For lI = 0 To UBound(lDescImportante) - 1
        lAuxiliar2 = lDescImportante(lCountLine)
        If gM46V999.gfPreenchido(lAuxiliar2) Then
            lCountLine = lCountLine + 1
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    Next lI

    'Alterado em 23/03/2007 - Ficha: 0700328
    If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        lCountLine = lCountLine + 1
        lAuxiliar2 = "-Bens de Terceiro em Poder do Segurado: Para ser passível de cobertura, deverá ser destacado do Valor em Risco"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        lCountLine = lCountLine + 1
        lAuxiliar2 = " Total o Valor em Risco destes bens"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'Alterado em 24/04/2007 - Ficha: 0700328
        If (pobjstcPedLoc.Cod_Clasf = gM46V999.gfBuscarDePara(18, 2) And pobjstcPedLoc.Cod_Compl_Clasf = gM46V999.gfBuscarDePara(18, 2, 10)) Or _
           (pobjstcPedLoc.Cod_Clasf = gM46V999.gfBuscarDePara(18, 2) And pobjstcPedLoc.Cod_Compl_Clasf = gM46V999.gfBuscarDePara(18, 2, 11)) Or _
           (pobjstcPedLoc.Cod_Clasf = gM46V999.gfBuscarDePara(19, 2) And pobjstcPedLoc.Cod_Compl_Clasf = gM46V999.gfBuscarDePara(19, 2, 10)) Then
            lCountLine = lCountLine + 1
            lAuxiliar2 = " ******    RISCOS NOMEADOS - RAMO 96    ******"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    End If

    If gM46V999.gCodProduto >= 958 And pobjstcPedLoc.Cob_Bens = True And Not mfVerDigBens() Then
        lCountLine = 0
        For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
            If gM46V999.gCod_Ramo = e_CodRamo_Empresarial And lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(94, 1) Then
                lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
                lAuxiliar2 = "A relação poderá ser preenchida diretamente no SYAS ou ser entregue juntamente com a proposta de seguro."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
            End If
            If gM46V999.gCod_Ramo = e_CodRamo_Residencial And lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(27, 1) Or _
               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(29, 1) Or _
               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(30, 1) Or _
               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(31, 1) Then
                lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
                lAuxiliar2 = "A relação poderá ser preenchida diretamente no SYAS ou ser entregue juntamente com a proposta de seguro."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
            End If
        Next lobjstcCob
        Do While lCountLine < 4
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        Loop
        GoTo FimTexto
    End If

    Do While lCountLine < 10
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
    Loop

FimTexto:
    mfCriDados_MULTCOT2_V_0 = True
End Function

Private Function mfCriDados_MULTCOT3_V_0(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                         pNumDe As Integer, _
                                         pNumAte As Integer) As Boolean
'Função: cria arquivo de dados - cotação coletiva - dados do item -

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lCodCobertura As Integer    'Código da cobertura.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lTamanho As Integer    'Tamanho.
    Dim lObjstcCorr As stcA46V705B
    Dim lobjstcCob As stcA46V704B    'Objeto da classe (stc) de cobertura.
    Dim lobjstcSeg As stcA46V716B    'Outra seguradora.
    Dim lCountLine As Integer
    Dim lValTotCober As Double    'Soma das coberturas.
    Dim lDescFranquias() As String
    Dim lobjstcPedLoc As stcA46V708B
    Dim lCount As Integer
    Dim lPrLiqItem As Double

    mfCriDados_MULTCOT3_V_0 = False
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        '001. Título do documento.
        Print #pNumArqDados, "ANEXO DA COTAÇÃO DE SEGURO"

        '002. Título do seguro (de acordo com a base de dados).
        '     Neste caso deve ser colocado YASUDA RESIDENCIAL, YASUDA CONDOMÍNIO ou YASUDA EMPRESARIAL.
        Print #pNumArqDados, pobjstcPed.gDesc_Ramo

        '003. Número da página.
        Print #pNumArqDados, "Página "; pNumDe & "/" & pNumAte
        pNumDe = pNumDe + 1

        '004. Emissão.
        Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

        '005. Validade.
        If lobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
            Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
        Else
            Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
        End If

        '006. Número do pedido .
        lAuxiliar1 = pobjstcPed.Num_Ped
        Print #pNumArqDados, Format$(lAuxiliar1, "000000")

        '007. Número do item .
        lAuxiliar1 = lobjstcPedLoc.Num_Item
        Print #pNumArqDados, Format$(lAuxiliar1, "0000")

        '008. Corretor (código e nome) .
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
        Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

        '009. Tarifa .
        lAuxiliar2 = pobjstcPed.Cod_Produto
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '010. Nosso número .
        lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '011. Local em risco - logradouro .
        lAuxiliar2 = lobjstcPedLoc.Cod_Tip_Logr_Risco & " " & lobjstcPedLoc.Nom_Logr_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '012. Local em risco - número .
        lAuxiliar2 = lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '013. Local em risco - complemento .
        lAuxiliar2 = lobjstcPedLoc.Dsc_Compl_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '014. Local em risco - bairro .
        lAuxiliar2 = lobjstcPedLoc.Nom_Bairro_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '015. Local em risco - cidade .
        lAuxiliar2 = lobjstcPedLoc.Nom_Cid_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '016. Local em risco - UF .
        lAuxiliar2 = lobjstcPedLoc.Sig_UF_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '017. Local em risco - CEP .
        lAuxiliar2 = lobjstcPedLoc.gCEP_Formatado
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '018. Dados para inspeção - nome da pessoa para contato .
        lAuxiliar2 = lobjstcPedLoc.objstcPedInsp.Pess_Cont
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '019. Dados para inspeção - telefone para contato .
        lAuxiliar2 = lobjstcPedLoc.objstcPedInsp.Num_Tel
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '020. Dados do seguro - classificação .
        lAuxiliar2 = lobjstcPedLoc.Cod_Clasf & "-" & lobjstcPedLoc.Cod_Compl_Clasf & "- " & lobjstcPedLoc.Dsc_Clasf
        lAuxiliar2 = lAuxiliar2 & String(15, " ") & " PLANO: " & pobjstcPed.Desc_Plano
        lAuxiliar2 = lAuxiliar2 & " - PRODUTO: " & Format(lobjstcPedLoc.Cod_Produto, "0000")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '021. Dados do seguro - construção .
        'Ficha 1000305
        If lobjstcPedLoc.Cod_Ramo = e_CodRamo_Empresarial Then
            Select Case lobjstcPedLoc.Cod_Constr
            Case e_CodContr_Superior
                lAuxiliar2 = "Superior/sólida"
            Case e_CodContr_Inferior
                lAuxiliar2 = "Inferior"
            Case e_CodContr_Mista
                lAuxiliar2 = "Mista"
            Case Else
                lAuxiliar2 = "Superior/sólida"
            End Select
        Else
            Select Case lobjstcPedLoc.Cod_Constr
            Case e_CodContr_Superior
                lAuxiliar2 = "Superior/sólida"
            Case e_CodContr_Inferior
                lAuxiliar2 = "Inferior/mista"
            Case Else
                lAuxiliar2 = "Superior/sólida"
            End Select
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '022. Dados do seguro - região .
        lAuxiliar2 = lobjstcPedLoc.Cod_Regiao & " - " & Left(lobjstcPedLoc.gDesc_Regiao, 38)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '023. Dados do seguro - vigência .
        lAuxiliar2 = "de " & gM46V999.gfForData(2, pobjstcPed.Dat_Ini_Vig) & _
                   " a " & gM46V999.gfForData(2, pobjstcPed.Dat_Ter_Vig)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '024. Dados do seguro - valor em risco .
        lAuxiliar2 = "R$ " & Format(lobjstcPedLoc.Val_Vr, "standard")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '025. Dados do seguro - tipo de emissão .
        lAuxiliar2 = lobjstcPedLoc.gDesc_TipEmissao
        Select Case lobjstcPedLoc.Tip_Emissao
        Case e_TipEmissao_RenYasuda Or e_TipEmissao_Endosso
            lAuxiliar2 = lAuxiliar2 & "  Apólice: " & lobjstcPedLoc.Num_Apol_Anterior & "  Item: " & lobjstcPedLoc.Num_Item_Anterior & "  Sinistro: " & IIf(lobjstcPedLoc.Ind_Sinistro = e_SimNao_Sim, "Sim", "Não")
        End Select
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '026 e 027. Coberturas contratadas - título.
        If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            '035 e 036. Coberturas contratadas - título.
            lAuxiliar2 = "Desconto de"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            lAuxiliar2 = "experiência"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Else
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
        End If


        '028 a 159. Coberturas contratadas - linhas de 01 a 22.
        lCountLine = 0
        lValTotCober = 0
        For Each lobjstcCob In lobjstcPedLoc.ColecaoPedCobert
            lCountLine = lCountLine + 1
            'Código
            Print #pNumArqDados, Format$(lobjstcCob.Cod_Cobert, "0000")
            'Descrição
            Print #pNumArqDados, Left(lobjstcCob.gNomeCobertura, 48)

            'Valor da IS
            lAuxiliar1 = lobjstcCob.Val_IS
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 14 Then
                lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            'd) Desconto comercial .
            lAuxiliar1 = lobjstcCob.Per_Desc_Coml
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 10 Then
                lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
                'e) Desconto de experiência .
                lAuxiliar1 = lobjstcCob.Per_Desc_Exp
                lAuxiliar2 = Format$(lAuxiliar1, "standard")
                lTamanho = Len(lAuxiliar2)
                If lTamanho < 10 Then
                    lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
                End If
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            Else
                Print #pNumArqDados, " "
            End If

            'f) Prêmio líquido .
            lAuxiliar1 = lobjstcCob.VAL_PR
            lValTotCober = lValTotCober + lobjstcCob.VAL_PR
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 14 Then
                lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Next lobjstcCob
        Do While lCountLine < 22
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
        Loop


        '160. Coberturas contratadas - soma.
        lAuxiliar1 = lValTotCober
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '161. Coberturas contratadas - prêmio líquido.
        If pobjstcPed.Cod_Dizer <> 0 Then
            If lValTotCober < gM46V999.gProPreMinParcelaEndosso Then
                lValTotCober = gM46V999.gProPreMinParcelaEndosso
            End If
        Else
            If lValTotCober < gM46V999.gProPreMinParcela Then
                lValTotCober = gM46V999.gProPreMinParcela
            End If
        End If

        lAuxiliar1 = lValTotCober
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '162. Informação cifrada .
        '   B: 25 / A: 50 / P: 10 / CO: 99999 / C: 99999  onde:
        '   B = bônus
        '   A = aparelho
        '   P = progressivo
        '   CO = coeficiente de comissão
        '   C = coeficiente LMG/VR

        'Ficha 900740 e 901021 - zerar bônus das renovações dos planinhos
        '2001 a 2010 - Correta
        '2011 a 2020 - GalCorr
        '2021 a 2031 - Mercedes
        '2032 a 2041 - National
        '2042 a 2049 - Manaus
        '2050 - Ballarin
        '2051 a 2060 - Marsh
        '2061 a 2070 - NSK
        '2071 0 2080 - Ajinomoto
        '2081 a 2090 - Nissin
        '2091 a 2100 - Panasonic
        '2101 - Fanel
        '2102 - JLA

        'Ficha 901673
        '2111 a 2120 - Embraer
        'Ficha 901772
        '2121 a 2130 - Bardella

        If (gM46V111.gCod_plano >= 2001 And gM46V111.gCod_plano <= 2049) Or (gM46V111.gCod_plano >= 2051 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
            lAuxiliar2 = "B:00"
        Else
            lAuxiliar2 = "B:" & Format(lobjstcPedLoc.Per_Bonus, "00")
        End If

        lAuxiliar2 = lAuxiliar2 & "/A:" & Format(lobjstcPedLoc.Per_Desc_Apar, "00")

        'Ficha 900740 e 901673 e 901772
        If (gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031) Or (gM46V111.gCod_plano >= 2061 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
            lAuxiliar2 = lAuxiliar2 & "/P:00"
        Else
            lAuxiliar2 = lAuxiliar2 & "/P:" & Format(lobjstcPedLoc.Per_Desc_Progres, "00")
        End If

        'lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(pObjstcPed.objstcCtrlEmis.Per_Corr + pObjstcPed.objstcCtrlEmis.Per_Corr_2 + pObjstcPed.objstcCtrlEmis.Per_Corr_3, "00.00")
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
        lAuxiliar1 = lObjstcCorr.Per_Corr

        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(2)
            lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(3)
                lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            End If
        End If

        lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(lAuxiliar1 * 100, "0000")

        lAuxiliar2 = lAuxiliar2 & "/C:" & Format(Round(lobjstcPedLoc.Val_Coef_IS_Vr, 4) * 10000, "00000")
        lAuxiliar2 = lAuxiliar2 & "/PRBO19"

        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '163 a 171. Segurança.
        '           O texto pode ser customizado de acordo com a necessidade.
        '           Temos 9 linhas disponíveis e cada linha cabe 70 caracteres (contando com o "(X)").
        '           O "X" deve ser colocado de acordo com a base de dados.
        'a) Linha 01.
        If pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Then
            Print #pNumArqDados, "SEGURANÇA"
            If lobjstcPedLoc.Cod_Iden_Segur_01_V = "S" Then
                lAuxiliar2 = "(X) "
            Else
                lAuxiliar2 = "( ) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Vigilância exclusiva e permanente do imóvel entendendo-se ainda"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            'b) Linha 02.
            lAuxiliar2 = "    como tal, porteiros fixos."
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            'c) Linha 03.
            If lobjstcPedLoc.Cod_Iden_Segur_02_A = "S" Then
                lAuxiliar2 = "(X) "
            Else
                lAuxiliar2 = "( ) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Alarmes."
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            'g) Linha 04.
            If lobjstcPedLoc.Cod_Iden_Segur_03_P = "S" Then
                lAuxiliar2 = "(X) "
            Else
                lAuxiliar2 = "( ) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Porteiros eletrônicos."
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            'h) Linha 05.
            If lobjstcPedLoc.Cod_Iden_Segur_04_I = "S" Then
                lAuxiliar2 = "(X) "
            Else
                lAuxiliar2 = "( ) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Inexistência de terreno baldio, imóvel desocupado ou em construção"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            'i) Linha 06.
            lAuxiliar2 = "     na vizinhança do risco."
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            'l) Linha 07.
            If lobjstcPedLoc.Cod_Iden_Segur_05_G = "S" Then
                lAuxiliar2 = "(X) "
            Else
                lAuxiliar2 = "( ) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Grades metálicas de proteção em todas as janelas e portas de vidro"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            'm) Linha 08.
            lAuxiliar2 = "    que dão acesso ao interior do imóvel desde que também possuam"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            'n) Linha 09.
            lAuxiliar2 = "    alarme."
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Else
            Print #pNumArqDados, " "
            For lI = 1 To 9
                Print #pNumArqDados, String(72, "X")
            Next lI
        End If

        '172 a 176. Proteção.
        If pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Or pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            '           O texto pode ser customizado de acordo com a necessidade.
            '           Temos 4 linhas disponíveis e cada linha cabe 39 caracteres (contando com o "(X)").
            '           O "X" deve ser colocado de acordo com a base de dados.
            'a) Título.
            lAuxiliar2 = "PROTEÇÃO"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'b) Linha 01.
            If lobjstcPedLoc.Cod_Iden_Protec_01_E = "S" Then
                lAuxiliar2 = "(X) "
            Else
                lAuxiliar2 = "( ) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Extintor"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'c) Linha 02.
            If lobjstcPedLoc.Cod_Iden_Protec_02_H = "S" Then
                lAuxiliar2 = "(X) "
            Else
                lAuxiliar2 = "( ) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Hidrantes"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'd) Linha 03.
            If lobjstcPedLoc.Cod_Iden_Protec_03_S = "S" Then
                lAuxiliar2 = "(X) "
            Else
                lAuxiliar2 = "( ) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Sprinkler (chuveiros automáticos)"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'e) Linha 04.
            If lobjstcPedLoc.Cod_Iden_Protec_04_D = "S" Then
                lAuxiliar2 = "(X) "
            Else
                lAuxiliar2 = "( ) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Detecção e alarme"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Else
            'a) Título.
            Print #pNumArqDados, " "
            For lI = 1 To 4
                Print #pNumArqDados, String(39, "X")
            Next lI
        End If


        '177 a 184. Franquias.
        Call mpDescricaoFranquia_V_0(lobjstcPedLoc, pobjstcPed, lDescFranquias, 8)
        If UBound(lDescFranquias) > 0 Then

            '            If lDescFranquias(0) <> "" Then
            '                'a) Linha 1.
            '                Print #pNumArqDados, "Por evento, serão deduzidas das indenizações a seguinte participação do " & _
                             '                                 "segurado:"
            '            Else
            '                Print #pNumArqDados, " "
            '            End If
            'Print #pNumArqDados, " "

            For lI = 0 To 7
                lAuxiliar2 = lDescFranquias(lI)
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            Next lI
        End If
    Next lobjstcPedLoc

    mfCriDados_MULTCOT3_V_0 = True
End Function

Private Function mfCriDados_MULTPRO1_V_0(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                         pobjstcPedLoc As stcA46V708B) As Boolean
'Função: cria arquivo de dados - proposta empresarial simples

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.


    Dim lAuxiliar1 As Double    'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String    'Auxiliar - utilizado apenas como exemplo.
    Dim lCodCobertura As Integer    'Código da cobertura.
    Dim lI As Integer             'Utilizado com For...Next.
    Dim lTamanho As Integer    'Tamanho.
    Dim lObjstcCorr As stcA46V705B
    Dim lobjstcCob As stcA46V704B
    Dim lobjstcJur As stcA46V711B
    Dim lobjstcSeg As stcA46V716B
    Dim lCountLine As Integer
    Dim lValTotCober As Double
    mfCriDados_MULTPRO1_V_0 = False

    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado PROPOSTA DE SEGURO YASUDA...
    '     RESIDENCIAL, CONDOMÍNIO ou EMPRESARIAL.
    Print #pNumArqDados, "PROPOSTA DE SEGURO YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

    '002. Código do Processo Susep.
    '     Observar que este número deve ser de acordo com o título do seguro:
    '     RESIDENCIAL, CONDOMÍNIO ou EMPRESARIAL.
    lAuxiliar2 = " "
    Select Case pobjstcPed.Cod_Ramo
    Case e_CodRamo_Condominio
        lAuxiliar2 = "Processo Susep 15414.004604/2004-72"
    Case e_CodRamo_Empresarial
        lAuxiliar2 = "Processo Susep 15414.004605/2004-17"
    Case e_CodRamo_Residencial
        lAuxiliar2 = "Processo Susep 15414.003659/2004-65"
    End Select

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '003. Emissão.
    Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

    '004. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '005. Número do item .
    lAuxiliar1 = pobjstcPedLoc.Num_Item
    Print #pNumArqDados, Format$(lAuxiliar1, "0000")

    '006. Corretor (código e nome) .
    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
    lAuxiliar1 = lObjstcCorr.Cod_Corr
    lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '007. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '008. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '009. Proponente - nome do segurado .
    lAuxiliar2 = pobjstcPed.Nom_Segurado
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '010. Proponente - código do segurado - se houver .
    lAuxiliar1 = pobjstcPed.Cod_Segurado
    Print #pNumArqDados, Format$(lAuxiliar1, "0000000")

    '011. Proponente - tipo de pessoa .
    Select Case pobjstcPed.Tip_Pessoa
    Case e_TipPessoa_Juridica
        Print #pNumArqDados, "Jurídica"
        '012 e 013. Proponente - CNPJ ou CPF .
        lAuxiliar2 = "CNPJ.........:"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     "Ramo de atividade: " & pobjstcPed.gDesc_RamoAtiv
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Case e_TipPessoa_Fisica
        Print #pNumArqDados, "Física"
        '012 e 013. Proponente - CNPJ ou CPF .
        lAuxiliar2 = "CPF...........:"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     pobjstcPed.gCod_Doc_Segur & ": " & pobjstcPed.Num_Doc & "   " & _
                     "Órgão emissor: " & pobjstcPed.Org_Exp & "   " & _
                     "Emissão: " & gM46V999.gfForData(2, pobjstcPed.Dat_Emis_Doc)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End Select

    '014. Proponente - endereço .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Cod_Tip_Logr & " " & _
                 pobjstcPed.objstcPedEndereco.Nom_Logr & "," & _
                 pobjstcPed.objstcPedEndereco.Dsc_Compl_Logr
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '015. Proponente - bairro .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Nom_Bairro
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '016. Proponente - cidade .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Nom_Cid
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '017. Proponente - UF .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Sig_UF
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '018. Proponente - CEP .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.gCEP_Formatado
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '019. Proponente - Telefone .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Num_Fone
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '020. Local em risco - logradouro .
    lAuxiliar2 = pobjstcPedLoc.Cod_Tip_Logr_Risco & " " & pobjstcPedLoc.Nom_Logr_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '021. Local em risco - número .
    lAuxiliar2 = pobjstcPedLoc.objstcPedInsp.Num_Logr_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '022. Local em risco - complemento .
    lAuxiliar2 = pobjstcPedLoc.Dsc_Compl_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '023. Local em risco - bairro .
    lAuxiliar2 = pobjstcPedLoc.Nom_Bairro_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '024. Local em risco - cidade .
    lAuxiliar2 = pobjstcPedLoc.Nom_Cid_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '025. Local em risco - UF .
    lAuxiliar2 = pobjstcPedLoc.Sig_UF_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '026. Local em risco - CEP .
    lAuxiliar2 = pobjstcPedLoc.gCEP_Formatado
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '027. Dados para inspeção - nome da pessoa para contato .
    lAuxiliar2 = pobjstcPedLoc.objstcPedInsp.Pess_Cont
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '028. Dados para inspeção - telefone para contato .
    lAuxiliar2 = pobjstcPedLoc.objstcPedInsp.Num_Tel
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '029. Dados do seguro - classificação .
    lAuxiliar2 = pobjstcPedLoc.Cod_Clasf & "-" & pobjstcPedLoc.Cod_Compl_Clasf & "- " & pobjstcPedLoc.Dsc_Clasf
    lAuxiliar2 = lAuxiliar2 & String(15, " ") & " PLANO: " & pobjstcPed.Desc_Plano
    lAuxiliar2 = lAuxiliar2 & " - PRODUTO: " & Format(pobjstcPedLoc.Cod_Produto, "0000")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '030. Dados do seguro - construção .
    'Ficha 1000305
    If pobjstcPedLoc.Cod_Ramo = e_CodRamo_Empresarial Then
        Select Case pobjstcPedLoc.Cod_Constr
        Case e_CodContr_Superior
            lAuxiliar2 = "Superior/sólida"
        Case e_CodContr_Inferior
            lAuxiliar2 = "Inferior"
        Case e_CodContr_Mista
            lAuxiliar2 = "Mista"
        Case Else
            lAuxiliar2 = "Superior/sólida"
        End Select
    Else
        Select Case pobjstcPedLoc.Cod_Constr
        Case e_CodContr_Superior
            lAuxiliar2 = "Superior/sólida"
        Case e_CodContr_Inferior
            lAuxiliar2 = "Inferior/mista"
        Case Else
            lAuxiliar2 = "Superior/sólida"
        End Select
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '031. Dados do seguro - região .
    lAuxiliar2 = pobjstcPedLoc.Cod_Regiao & " - " & Left(pobjstcPedLoc.gDesc_Regiao, 38)
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '032. Dados do seguro - vigência .
    lAuxiliar2 = "de " & gM46V999.gfForData(2, pobjstcPed.Dat_Ini_Vig) & _
               " a " & gM46V999.gfForData(2, pobjstcPed.Dat_Ter_Vig)
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '033. Dados do seguro - valor em risco .
    lAuxiliar2 = "R$ " & Format(pobjstcPedLoc.Val_Vr, "standard")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '034. Dados do seguro - tipo de emissão .
    lAuxiliar2 = pobjstcPedLoc.gDesc_TipEmissao
    Select Case pobjstcPedLoc.Tip_Emissao
    Case e_TipEmissao_RenYasuda Or e_TipEmissao_Endosso
        lAuxiliar2 = lAuxiliar2 & "  Apólice: " & pobjstcPedLoc.Num_Apol_Anterior & "  Item: " & pobjstcPedLoc.Num_Item_Anterior & "  Sinistro: " & IIf(pobjstcPedLoc.Ind_Sinistro = e_SimNao_Sim, "Sim", "Não")

    End Select
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '035 e 036. Coberturas contratadas - desconto - título.
    If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        lAuxiliar2 = "Desconto de"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = "experiência"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Else
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
    End If

    '037 a 168. Coberturas contratadas - linhas de 01 a 22.
    lCountLine = 0
    lValTotCober = 0
    For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
        lCountLine = lCountLine + 1
        'Código
        Print #pNumArqDados, Format$(lobjstcCob.Cod_Cobert, "0000")
        'Descrição
        Print #pNumArqDados, Left(lobjstcCob.gNomeCobertura, 48)

        'Valor da IS
        lAuxiliar1 = lobjstcCob.Val_IS
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        'd) Desconto comercial .
        lAuxiliar1 = lobjstcCob.Per_Desc_Coml
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 10 Then
            lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            'e) Desconto de experiência .
            lAuxiliar1 = lobjstcCob.Per_Desc_Exp
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 10 Then
                lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Else
            Print #pNumArqDados, " "
        End If

        'f) Prêmio líquido .
        lAuxiliar1 = lobjstcCob.VAL_PR
        lValTotCober = lValTotCober + lobjstcCob.VAL_PR
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Next lobjstcCob
    Do While lCountLine < 22
        lCountLine = lCountLine + 1
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
    Loop


    '169. Coberturas contratadas - soma.
    lAuxiliar1 = lValTotCober
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '170. Coberturas contratadas - prêmio líquido.
    lAuxiliar1 = pobjstcPed.VAL_PR
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '171. Coberturas contratadas - custo da apólice.
    lAuxiliar1 = pobjstcPed.Val_Cust_Emis
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


    '172. Coberturas contratadas - IOF.
    lAuxiliar1 = pobjstcPed.Val_IOF
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '173. Coberturas contratadas - prêmio total a vista.
    lAuxiliar1 = pobjstcPed.gValorTotal
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '174. Informação cifrada.
    '   B: 25 / A: 50 / P: 10 / CO: 99999 / C: 99999  onde:
    '   B = bônus
    '   A = aparelho
    '   P = progressivo
    '   CO = coeficiente de comissão
    '   C = coeficiente LMG/VR

    'Ficha 900740 e 901021 - zerar bônus das renovações dos planinhos
    '2001 a 2010 - Correta
    '2011 a 2020 - GalCorr
    '2021 a 2031 - Mercedes
    '2032 a 2041 - National
    '2042 a 2049 - Manaus
    '2050 - Ballarin
    '2051 a 2060 - Marsh
    '2061 a 2070 - NSK
    '2071 a 2080 - Ajinomoto
    '2081 a 2090 - Nissin
    '2091 a 2100 - Panasonic
    '2101 - Fanel
    '2102 - JLA

    'Ficha 901673
    '2111 a 2120 - Embraer
    'Ficha 901772
    '2121 a 2130 - Bardella

    If (gM46V111.gCod_plano >= 2001 And gM46V111.gCod_plano <= 2049) Or (gM46V111.gCod_plano >= 2051 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
        lAuxiliar2 = "B:00"
    Else
        lAuxiliar2 = "B:" & Format(pobjstcPedLoc.Per_Bonus, "00")
    End If

    lAuxiliar2 = lAuxiliar2 & "/A:" & Format(pobjstcPedLoc.Per_Desc_Apar, "00")

    'Ficha 900740 e 901673 e 901772
    If (gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031) Or (gM46V111.gCod_plano >= 2061 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
        lAuxiliar2 = lAuxiliar2 & "/P:00"
    Else
        lAuxiliar2 = lAuxiliar2 & "/P:" & Format(pobjstcPedLoc.Per_Desc_Progres, "00")
    End If

    'lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(pObjstcPed.objstcCtrlEmis.Per_Corr + pObjstcPed.objstcCtrlEmis.Per_Corr_2 + pObjstcPed.objstcCtrlEmis.Per_Corr_3, "00.00")
    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
    lAuxiliar1 = lObjstcCorr.Per_Corr

    If pobjstcPed.ColecaoPedCorr.Count > 1 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(2)
        lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
        If pobjstcPed.ColecaoPedCorr.Count > 2 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(3)
            lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
        End If
    End If

    lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(lAuxiliar1 * 100, "0000")

    lAuxiliar2 = lAuxiliar2 & "/C:" & Format(Round(pobjstcPedLoc.Val_Coef_IS_Vr, 4) * 10000, "00000")
    lAuxiliar2 = lAuxiliar2 & "/PRBO19"

    '###INSPECAO
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        If pobjstcPedLoc.Insp_Obr = True Or pobjstcPedLoc.Ind_Sin_Inspec > 0 Then
            If pobjstcPedLoc.Insp_Agend = True Then
                lAuxiliar2 = lAuxiliar2 & "/INSPS"
            Else
                lAuxiliar2 = lAuxiliar2 & "/INSPN"
            End If
        End If
    End If

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '175 a 184. Segurança.
    '           O texto pode ser customizado de acordo com a necessidade.
    '           Temos 10 linhas disponíveis e cada linha cabe 70 caracteres (contando com o "(X)").
    '           O "X" deve ser colocado de acordo com a base de dados.
    'a) Linha 01.
    If pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Then
        Print #pNumArqDados, "SEGURANÇA"
        If pobjstcPedLoc.Cod_Iden_Segur_01_V = "S" Then
            lAuxiliar2 = "(S) "
        Else
            lAuxiliar2 = "(N) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Vigilância exclusiva e permanente do imóvel entendendo-se ainda"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'b) Linha 02.
        lAuxiliar2 = "    como tal, porteiros fixos."
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'c) Linha 03.
        If pobjstcPedLoc.Cod_Iden_Segur_02_A = "S" Then
            lAuxiliar2 = "(S) "
        Else
            lAuxiliar2 = "(N) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Alarmes."
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'd) Linha 04.
        If pobjstcPedLoc.Cod_Iden_Segur_03_P = "S" Then
            lAuxiliar2 = "(S) "
        Else
            lAuxiliar2 = "(N) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Porteiros eletrônicos."
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'e) Linha 05.
        If pobjstcPedLoc.Cod_Iden_Segur_04_I = "S" Then
            lAuxiliar2 = "(S) "
        Else
            lAuxiliar2 = "(N) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Inexistência de terreno baldio, imóvel desocupado ou em construção"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'f) Linha 06.
        lAuxiliar2 = "    na vizinhança do risco."
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'g) Linha 07.
        If pobjstcPedLoc.Cod_Iden_Segur_05_G = "S" Then
            lAuxiliar2 = "(S) "
        Else
            lAuxiliar2 = "(N) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Grades metálicas de proteção em todas as janelas e portas de vidro"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'h) Linha 08.
        lAuxiliar2 = "    que dão acesso ao interior do imóvel desde que também possuam"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'i) Linha 09.
        lAuxiliar2 = "    alarme."
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'j) Linha 10.
        lAuxiliar2 = " "
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Else
        Print #pNumArqDados, " "
        For lI = 1 To 10
            Print #pNumArqDados, String(72, "X")
        Next lI
    End If

    '194 a 198. Proteção.
    '           Para residencial, sairá em branco.
    If pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Or pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        '           O texto pode ser customizado de acordo com a necessidade.
        '           Temos 4 linhas disponíveis e cada linha cabe 39 caracteres (contando com o "(X)").
        '           O "X" deve ser colocado de acordo com a base de dados.
        'a) Título.
        lAuxiliar2 = "PROTEÇÃO"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'b) Linha 01.
        If pobjstcPedLoc.Cod_Iden_Protec_01_E = "S" Then
            lAuxiliar2 = "(S) "
        Else
            lAuxiliar2 = "(N) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Extintor"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'c) Linha 02.
        If pobjstcPedLoc.Cod_Iden_Protec_02_H = "S" Then
            lAuxiliar2 = "(S) "
        Else
            lAuxiliar2 = "(N) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Hidrantes"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'd) Linha 03.
        If pobjstcPedLoc.Cod_Iden_Protec_03_S = "S" Then
            lAuxiliar2 = "(S) "
        Else
            lAuxiliar2 = "(N) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Sprinkler (chuveiros automáticos)"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'e) Linha 04.
        If pobjstcPedLoc.Cod_Iden_Protec_04_D = "S" Then
            lAuxiliar2 = "(S) "
        Else
            lAuxiliar2 = "(N) "
        End If
        lAuxiliar2 = lAuxiliar2 & "Detecção e alarme"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Else
        'a) Título.
        Print #pNumArqDados, " "
        For lI = 1 To 4
            Print #pNumArqDados, String(37, "X")
        Next lI
    End If


    '190 a 193. Observação 1.
    'a) Linha 1.
    lAuxiliar2 = "A aceitação do seguro estará sujeita à análise do risco. O registro deste plano na " & _
                 "SUSEP não implica,  por  parte"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'b) Linha 2.
    lAuxiliar2 = "da Autarquia, incentivo ou recomendação à sua comercialização. O segurado poderá " & _
                 "consultar a  situação  cadastral"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'c) Linha 3.
    lAuxiliar2 = "do seu corretor de seguros no site www.susep.gov.br, por meio do número de seu " & _
                 "registro na SUSEP, nome  completo, "
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'd) Linha 4.
    lAuxiliar2 = "CNPJ ou CPF."
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '194 a 205. Observação 2.
    'a) Linha 1.
    lAuxiliar2 = "Propõe(m) à YASUDA SEGUROS S.A. a realização do seguro, declarando conhecer " & _
                 "previamente e  aceitar  as  Condições"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'b) Linha 2.
    lAuxiliar2 = "Gerais do(s) anexo(s) que servem de base para esta proposta. Também presta(m)  as  " & _
                 "informações  solicitadas  pela"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'c) Linha 3.
    lAuxiliar2 = "YASUDA SEGUROS S.A., constantes do(s) anexo(s) a esta, para  fins  de  estudos  " & _
                 "da  concretização  do  respectivo"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'd) Linha 4.
    lAuxiliar2 = "contrato.  Assume(m)  o(s)  proponente(s),  nesta  oportunidade,  integral  " & _
                 "responsabilidade  pela  exatidão  das"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'e) Linha 5.
    lAuxiliar2 = "informações ora prestadas, autorizando a YASUDA SEGUROS S.A., caso aceite esta " & _
                 "proposta, a  emitir  a  respectiva"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'f) Linha 6.
    lAuxiliar2 = "apólice, cujo prêmio e eventuais despesas o(s) proponente(s) compromete(m)-se  a  " & _
                 "pagar,  tão  logo  lhe(s)  seja"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'g) Linha 7.
    lAuxiliar2 = "exigido. Em caso de recusa da proposta de seguro, os valores eventualmente " & _
                 "adiantados para  futuro  pagamento  de"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'h) Linha 8.
    lAuxiliar2 = "prêmio pelo(s) proponente(s) serão devolvidos integralmente, no prazo de 10 dias " & _
                 "corridos da data de sua  recusa."
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'i) Linha 9.
    lAuxiliar2 = "Ultrapassado este prazo, os valores sujeitam-se a atualização monetária pelo " & _
                 "IPC/FIPE, contados a partir da  data"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'j) Linha 10.
    lAuxiliar2 = "em que se tornarem exigíveis e juros moratórios de 0,5% (cinco décimos por cento) " & _
                 "ao mês, contados  a  partir  do"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'k) Linha 11.
    lAuxiliar2 = "primeiro dia posterior ao término do prazo fixado acima. Ambos calculados " & _
                 "pro-rata-die  até  a  data  da  efetiva"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'l) Linha 12.
    lAuxiliar2 = "restituição."
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    mfCriDados_MULTPRO1_V_0 = True
End Function

Private Function mfCriDados_MULTPRO2_V_0(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                         pobjstcPedLoc As stcA46V708B) As Boolean
'Função: cria arquivo de dados - proposta empresarial simples -

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.


    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lCodCobertura As Integer    'Código da cobertura.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lTamanho As Integer    'Tamanho.
    Dim lCountLine As Integer    'Contador de linha.
    Dim lobjstcBem As stcA46V715B    'Classe (stc) de Bem.
    Dim lObjstcCorr As stcA46V705B    'Classe (stc) de corretor.
    Dim lobjstcCoss As stcA46V706B    'Classe (stc) de cosseguro.
    Dim lobjstcOutros As stcA46V716B    'Classe (stc) de outros seguros.
    Dim lObjstcVenc As stcA46V713B
    Dim lobjstcCob As stcA46V704B    'Objeto da classe (stc) de cobertura.

    Dim lDescFranquias() As String
    Dim lDescImportante() As String
    Dim lValorParc As Variant
    Dim lMensagem As String
    Dim lblnAchou As Boolean
    Dim lobjstcClau As stcA46V703B  'Objeto da classe de estrutura stcA46V703B(TAB_PED_CLAU).


    mfCriDados_MULTPRO2_V_0 = False

    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado ANEXO DA PROPOSTA DE SEGURO YASUDA...
    '     RESIDENCIAL, CONDOMÍNIO ou EMPRESARIAL.
    Print #pNumArqDados, "ANEXO DA PROPOSTA DE SEGURO YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

    '002. Emissão.
    Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

    '003. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '004. Número do item .
    lAuxiliar1 = pobjstcPedLoc.Num_Item
    Print #pNumArqDados, Format$(lAuxiliar1, "0000")

    '005. Corretor (código e nome) .
    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
    lAuxiliar1 = lObjstcCorr.Cod_Corr
    lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '006. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '007. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '008 a 012. Declarações/Observações .
    lCountLine = 0
    If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Or pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Then
        lCountLine = lCountLine + 1
        lAuxiliar2 = "Valor em risco declarado: R$ " & Format(pobjstcPedLoc.Val_Vr, "standard")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End If

    'Alterado em 23/03/2007 - Ficha: 0700328
    If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        lCountLine = lCountLine + 1
        lAuxiliar2 = "-Bens de Terceiro em Poder do Segurado: Para ser passível de cobertura, deverá ser destacado do Valor em Risco"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        lCountLine = lCountLine + 1
        lAuxiliar2 = " Total o Valor em Risco destes bens"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'Alterado em 24/04/2007 - Ficha: 0700328
        If (pobjstcPedLoc.Cod_Clasf = gM46V999.gfBuscarDePara(18, 2) And pobjstcPedLoc.Cod_Compl_Clasf = gM46V999.gfBuscarDePara(18, 2, 10)) Or _
           (pobjstcPedLoc.Cod_Clasf = gM46V999.gfBuscarDePara(18, 2) And pobjstcPedLoc.Cod_Compl_Clasf = gM46V999.gfBuscarDePara(18, 2, 11)) Or _
           (pobjstcPedLoc.Cod_Clasf = gM46V999.gfBuscarDePara(19, 2) And pobjstcPedLoc.Cod_Compl_Clasf = gM46V999.gfBuscarDePara(19, 2, 10)) Then
            lCountLine = lCountLine + 1
            lAuxiliar2 = " ******    RISCOS NOMEADOS - RAMO 96    ******"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    End If

    lCountLine = lCountLine + 1
    lAuxiliar2 = "-Existem outros seguros cobrindo os mesmos bens? " & IIf(pobjstcPedLoc.ColecaoPedSeguro.Count > 0, "SIM", "NÃO")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    lblnAchou = False

    If pobjstcPed.Cod_Ramo <> e_CodRamo_Condominio Then
        lCountLine = lCountLine + 1
        lAuxiliar2 = "-Deseja a inclusão de Cláusula Beneficiária? "
        'If pobjstcPedLoc.objstcPedClau.Cod_Clau > 0 Then
        '    lAuxiliar2 = lAuxiliar2 & "SIM - " & Trim(Mid(pobjstcPedLoc.objstcPedClau.Descricao, 15, 60)) & " - " & Left(pobjstcPedLoc.objstcPedClau.Descricao, 14)
        'Else
        '    lAuxiliar2 = lAuxiliar2 & "NÃO"
        'End If
        Set lobjstcClau = New stcA46V703B
        Set lobjstcClau.M46V999 = gM46V999

        If pobjstcPedLoc.ColecaoPedClau.Count > 0 Then
            Set lobjstcClau = pobjstcPedLoc.ColecaoPedClau.item(1)
            If lobjstcClau.Cod_Clau > 0 Then
                lblnAchou = True
            End If
        End If
        If lblnAchou Then
            lAuxiliar2 = lAuxiliar2 & "SIM - " & Trim(Mid(lobjstcClau.Descricao, 15, 60)) & " - " & Left(lobjstcClau.Descricao, 14)
        Else
            lAuxiliar2 = lAuxiliar2 & "NÃO"
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End If

    If pobjstcPed.Cod_Ramo = e_CodRamo_Residencial Then
        lCountLine = lCountLine + 1
        'lAuxiliar2 = "Existem bens no interior da residência com valor unitário superior a R$ 5.000,00? "
        lAuxiliar2 = "-RELAÇÃO DE BENS SEGURADOS : Objetos com valores acima dos Limites de Indenização estabelecidos na Cláusula 4 das"
        Print #pNumArqDados, lAuxiliar2

        lCountLine = lCountLine + 1
        lAuxiliar2 = "Condições Gerais do produto Yasuda Residencial, no que tange a Cobertura de Roubo e Furto Com Vestígios, deverão "
        Print #pNumArqDados, lAuxiliar2

        lCountLine = lCountLine + 1
        lAuxiliar2 = "estar discriminados na proposta de seguro ( marca e modelo ) e estarão sujeitos a análise de aceitação por  parte"
        Print #pNumArqDados, lAuxiliar2

        lCountLine = lCountLine + 1
        lAuxiliar2 = "da Yasuda Seguros."
        Print #pNumArqDados, lAuxiliar2

        'If pobjstcPedLoc.ColecaoPedLocBem.Count > 0 Then
        '    lAuxiliar2 = lAuxiliar2 & "NÃO."
        '    For Each lobjstcBem In pobjstcPedLoc.ColecaoPedLocBem
        '        If lobjstcBem.Val_Bem > 5000 Then
        '            lAuxiliar2 = Replace(lAuxiliar2, "NÃO.", "SIM.")
        '            Exit For
        '        End If
        '    Next lobjstcBem
        'Else
        '    lAuxiliar2 = lAuxiliar2 & "NÃO."
        'End If
        'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        If pobjstcPedLoc.Cod_Ramo <> 112 Then
            lCountLine = lCountLine + 1
            lAuxiliar2 = "-Número de empregados na residência: " & pobjstcPedLoc.Qtd_Empreg
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    End If

    Do While (lCountLine < 7)
        lCountLine = lCountLine + 1
        Print #pNumArqDados, " "
    Loop

    '013. Número de parcelas .
    lAuxiliar1 = pobjstcPed.Qtd_Parc
    Print #pNumArqDados, Format$(lAuxiliar1, "00")

    '014. Parcelas antecipadas .
    lAuxiliar1 = IIf(pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim, "1", "0")
    Print #pNumArqDados, Format$(lAuxiliar1, "00")

    '015. Número do bloqueto.
    lAuxiliar2 = pobjstcPed.objstcCtrlEmis.Num_Bloq & "-" & pobjstcPed.objstcCtrlEmis.DIG_BLOQUETO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '016. Juros ao mês .
    lAuxiliar1 = pobjstcPed.objstcCtrlEmis.Per_Juros_Neg
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 5 Then
        lAuxiliar2 = Space(5 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2) & "%"

    '017. Juros de mora .
    If pobjstcPed.objstcCtrlEmis.Qtd_Parc = 1 And pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
        lAuxiliar1 = 0
    Else
        lAuxiliar1 = 4.5
    End If
    lAuxiliar1 = 4.5
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 5 Then
        lAuxiliar2 = Space(5 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2) & "%"

    'Calcula o fracionamento
    With pobjstcPed
        Call gM46V998.gfCalFracionamento(pobjstcPed, pobjstcPedLoc.Tip_Emissao, _
                                         lValorParc, lMensagem, gM46V111.gProJurIOF, .Cod_Reg, .Val_IOF, .Val_Parc_Min)
    End With
    If pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
        '018. Forma de pagamento: primeira parcela.
        Print #pNumArqDados, "R$" & Format(lValorParc(pobjstcPed.Qtd_Parc, 1), "standard")
        pobjstcPed.Val_PriParcela = lValorParc(pobjstcPed.Qtd_Parc, 1)

        '018. Forma de pagamento: demais parcelas.
        Print #pNumArqDados, "R$" & Format(lValorParc(pobjstcPed.Qtd_Parc, 2), "standard")
        pobjstcPed.val_DemParcelas = lValorParc(pobjstcPed.Qtd_Parc, 2)
    Else
        '019. Forma de pagamento: primeira parcela.
        Print #pNumArqDados, "R$" & Format(lValorParc(pobjstcPed.Qtd_Parc, 3), "standard")
        pobjstcPed.Val_PriParcela = lValorParc(pobjstcPed.Qtd_Parc, 3)

        '019. Forma de pagamento: demais parcelas.
        Print #pNumArqDados, "R$" & Format(lValorParc(pobjstcPed.Qtd_Parc, 4), "standard")
        pobjstcPed.val_DemParcelas = lValorParc(pobjstcPed.Qtd_Parc, 4)
    End If

    If pobjstcPed.Cod_Bco_Cobr <> 215 And pobjstcPed.Cod_Agen_Bas_Cor = 0 Then
        Print #pNumArqDados, "BANCO: " & pobjstcPed.Cod_Bco_Cobr & " - " & pobjstcPed.gDesc_Banco
    Else
        Print #pNumArqDados, " "
    End If

    '020 a 027. Franquias.
    'a) Linha 1.
    Call mpDescricaoFranquia_V_0(pobjstcPedLoc, pobjstcPed, lDescFranquias, 8)

    '    If lDescFranquias(0) <> "" Then
    '        'a) Linha 1.
    '        Print #pNumArqDados, "Por evento, serão deduzidas das indenizações a seguinte participação do " & _
             '                         "segurado:"
    '    Else
    '        Print #pNumArqDados, " "
    '    End If
    'Print #pNumArqDados, " "

    For lI = 0 To 7
        lAuxiliar2 = lDescFranquias(lI)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Next lI

    '028 a 037. Importante.
    Call mpTextoImportante(pobjstcPedLoc, lDescImportante, "MULTPRO2")
    lCountLine = 0
    For lI = 0 To UBound(lDescImportante) - 1
        lAuxiliar2 = lDescImportante(lCountLine)
        If gM46V999.gfPreenchido(lAuxiliar2) Then
            lCountLine = lCountLine + 1
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    Next lI
    If lCountLine < 10 Then
        Call mpTextoObs(pobjstcPedLoc, lDescImportante)
        For lI = 0 To UBound(lDescImportante) - 1
            lAuxiliar2 = lDescImportante(lI)
            If gM46V999.gfPreenchido(lAuxiliar2) And lCountLine < 10 Then
                lCountLine = lCountLine + 1
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
        Next lI
    End If

    If gM46V999.gCodProduto >= 958 And pobjstcPedLoc.Cob_Bens = True And Not mfVerDigBens() Then
        lCountLine = 0
        For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
            If gM46V999.gCod_Ramo = e_CodRamo_Empresarial And lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(94, 1) And pobjstcPedLoc.Dig_Bens = False Then
                lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores. Opção escolhida pelo corretor: "
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
                lAuxiliar2 = "Entrega da relação juntamente com a proposta de seguro."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
            End If
            If gM46V999.gCod_Ramo = e_CodRamo_Residencial And (lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(27, 1) Or lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(29, 1) Or _
                                                               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(30, 1) Or lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(31, 1)) And pobjstcPedLoc.Dig_Bens = False Then
                lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores. Opção escolhida pelo corretor: "
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
                lAuxiliar2 = "Entrega da relação juntamente com a proposta de seguro."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
            End If
        Next lobjstcCob
        Do While lCountLine < 4
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        Loop
        GoTo FimTexto
    End If

    Do While lCountLine < 10
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
    Loop

FimTexto:
    ' Cosseguro
    lAuxiliar2 = ""
    If pobjstcPed.ColecaoPedCoss.Count > 0 Then
        Print #pNumArqDados, "COSSEGURO:"
        lCountLine = 0
        For Each lobjstcCoss In pobjstcPed.ColecaoPedCoss
            lAuxiliar2 = lAuxiliar2 & " - ( " & lobjstcCoss.Cod_Coss & " - " & lobjstcCoss.Per_Coss & "% )"
        Next lobjstcCoss
        Print #pNumArqDados, lAuxiliar2
    Else
        For lI = 1 To 2
            Print #pNumArqDados, " "
        Next lI
    End If

    If pobjstcPedLoc.ColecaoPedSeguro.Count > 0 Then
        Print #pNumArqDados, "OUTROS SEGUROS:"
        lAuxiliar2 = ""
        lCountLine = 1
        For Each lobjstcOutros In pobjstcPedLoc.ColecaoPedSeguro
            If Not gM46V999.gfPreenchido(lAuxiliar2) Then
                lAuxiliar2 = " - (" & lobjstcOutros.Cod_Seguradora & " - Apolice: " & lobjstcOutros.Num_Apol_Ant & " Item: " & lobjstcOutros.Num_Item_Ant & ")"
            Else
                lCountLine = lCountLine + 1
                Print #pNumArqDados, lAuxiliar2 & "  - (" & lobjstcOutros.Cod_Seguradora & " - Apolice: " & lobjstcOutros.Num_Apol_Ant & " Item: " & lobjstcOutros.Num_Item_Ant & ")"
                lAuxiliar2 = ""
            End If
        Next lobjstcOutros
        If gM46V999.gfPreenchido(lAuxiliar2) Then
            lCountLine = lCountLine + 1
            Print #pNumArqDados, lAuxiliar2
        End If

        Do While lCountLine < 4
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "

        Loop
    Else
        lCountLine = 1
        Do While lCountLine < 5
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        Loop
    End If
    If pobjstcPed.ColecaoPedVenc.Count > 0 Then
        lAuxiliar2 = ""
        Print #pNumArqDados, "VENCIMENTO DAS PARCELAS"
        For Each lObjstcVenc In pobjstcPed.ColecaoPedVenc
            If lObjstcVenc.Num_Parc <= pobjstcPed.Qtd_Parc Then
                If lObjstcVenc.Num_Parc < 6 Then
                    lAuxiliar2 = lAuxiliar2 & "(" & lObjstcVenc.Num_Parc & "ª-" & gM46V999.gfForData(2, lObjstcVenc.DAT_VENC) & ") "
                Else
                    If lObjstcVenc.Num_Parc = 6 Then
                        Print #pNumArqDados, lAuxiliar2
                        lAuxiliar2 = ""
                    End If
                    lAuxiliar2 = lAuxiliar2 & "(" & lObjstcVenc.Num_Parc & "ª-" & gM46V999.gfForData(2, lObjstcVenc.DAT_VENC) & ") "
                End If
            End If
        Next lObjstcVenc
        If pobjstcPed.Qtd_Parc < 6 Then
            Print #pNumArqDados, lAuxiliar2
            Print #pNumArqDados, " "
        Else
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    Else
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
    End If

    '052. Unidade produtiva .
    lAuxiliar2 = pobjstcPed.Cod_Unid_Prod
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '053. Produtor .
    lAuxiliar2 = pobjstcPed.Cod_Prod
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '054. Corretor .
    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)

    lAuxiliar2 = Format(lObjstcCorr.Cod_Corr, "0000000")

    If pobjstcPed.ColecaoPedCorr.Count > 1 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(2)
        lAuxiliar2 = lAuxiliar2 & " " & Format(lObjstcCorr.Cod_Corr, "0000000")
        If pobjstcPed.ColecaoPedCorr.Count > 2 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(3)
            lAuxiliar2 = lAuxiliar2 & " " & Format(lObjstcCorr.Cod_Corr, "0000000")
        Else
            lAuxiliar2 = lAuxiliar2 & " 0000000 "
        End If
    Else
        lAuxiliar2 = lAuxiliar2 & " 0000000 0000000 "
    End If

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '055. Informação em código .

    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
    lAuxiliar2 = "C " & lObjstcCorr.Per_Corr * 100

    If pobjstcPed.ColecaoPedCorr.Count > 1 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(2)
        lAuxiliar2 = lAuxiliar2 & " " & lObjstcCorr.Per_Corr * 100
        If pobjstcPed.ColecaoPedCorr.Count > 2 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(3)
            lAuxiliar2 = lAuxiliar2 & " " & lObjstcCorr.Per_Corr * 100
        End If
    End If

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)



    '056. Dados para correspondência - endereço .
    lAuxiliar2 = pobjstcPed.Cod_Tip_Logr & " " & pobjstcPed.Nom_Logr & ", " & pobjstcPed.Dsc_Compl_Logr
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '057. Dados para correspondência - bairro .
    lAuxiliar2 = pobjstcPed.Nom_Bairro
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '058. Dados para correspondência - cidade .
    lAuxiliar2 = pobjstcPed.Nom_Cid
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '059. Dados para correspondência - UF .
    lAuxiliar2 = pobjstcPed.Sig_UF
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '060. Dados para correspondência - CEP .
    lAuxiliar2 = pobjstcPed.gCEP_Formatado
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '061. Local e data.
    lAuxiliar2 = "São Paulo, " & _
                 Format$(Now, "dd") & " de " & Format$(Now, "mmmm") & " de " & Format$(Now, "yyyy")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '062. Nome do corretor (assinatura)
    lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


    mfCriDados_MULTPRO2_V_0 = True
End Function
Private Function mfCriDados_MULTPRO3_V_0(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                         pobjstcPedLoc As stcA46V708B, pTotalPags As Integer) As Boolean
'Função: cria arquivo de dados - proposta coletiva - capa.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lDemParcelas As String    'Demais parcelas.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lJuros As String    'Juros.
    Dim lMensagem As String    'Mensagem.
    Dim lOrdem As String    'Ordem.
    Dim lPriParcela As String   'Primeira parcela.
    Dim lValorParc As Variant
    Dim lTamanho As Integer    'Tamanho.
    Dim lObjstcCorr As stcA46V705B
    Dim lObjstcVenc As stcA46V713B

    mfCriDados_MULTPRO3_V_0 = False

    '001. Título do documento.
    Print #pNumArqDados, "PROPOSTA DE SEGURO YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

    '002. Código do Processo Susep.
    '     Observar que este número deve ser de acordo com o título do seguro:
    '     RESIDENCIAL, CONDOMÍNIO ou EMPRESARIAL.
    lAuxiliar2 = " "
    Select Case pobjstcPed.Cod_Ramo
    Case e_CodRamo_Condominio
        lAuxiliar2 = "Processo Susep 15414.004604/2004-72"
    Case e_CodRamo_Empresarial
        lAuxiliar2 = "Processo Susep 15414.004605/2004-17"
    Case e_CodRamo_Residencial
        lAuxiliar2 = "Processo Susep 15414.003659/2004-65"
    End Select
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '003. Emissão.
    Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

    '004. Número da página.
    Print #pNumArqDados, "Página 1/" & pTotalPags

    '005. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '006. Corretor (código e nome) .
    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
    lAuxiliar1 = lObjstcCorr.Cod_Corr
    lAuxiliar2 = lObjstcCorr.gNom_Corretor

    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '007. Tarifa .
    Print #pNumArqDados, pobjstcPed.Cod_Produto

    '008. Nosso número .
    Print #pNumArqDados, pobjstcPed.NOSSO_NUMERO

    '009. Proponente - nome do segurado .
    Print #pNumArqDados, IIf(pobjstcPed.Nom_Segurado = "", " ", pobjstcPed.Nom_Segurado)

    '010. Proponente - código do segurado - se houver .
    lAuxiliar1 = pobjstcPed.Cod_Segurado
    Print #pNumArqDados, Format$(lAuxiliar1, "0000000")

    '011. Proponente - tipo de pessoa .
    Select Case pobjstcPed.Tip_Pessoa
    Case e_TipPessoa_Juridica
        Print #pNumArqDados, "Jurídica"
        '012 e 013. Proponente - CNPJ ou CPF .
        lAuxiliar2 = "CNPJ.........:"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     "Ramo de atividade: " & pobjstcPed.gDesc_RamoAtiv
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Case e_TipPessoa_Fisica
        Print #pNumArqDados, "Física"
        '012 e 013. Proponente - CNPJ ou CPF .
        lAuxiliar2 = "CPF...........:"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     pobjstcPed.gCod_Doc_Segur & ": " & pobjstcPed.Num_Doc & "   " & _
                     "Órgão emissor: " & pobjstcPed.Org_Exp & "   " & _
                     "Emissão: " & gM46V999.gfForData(2, pobjstcPed.Dat_Emis_Doc)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End Select

    '014. Proponente - endereço .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Cod_Tip_Logr & " " & _
                 pobjstcPed.objstcPedEndereco.Nom_Logr & "," & _
                 pobjstcPed.objstcPedEndereco.Dsc_Compl_Logr
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '015. Proponente - bairro .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Nom_Bairro
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '016. Proponente - cidade .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Nom_Cid
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '017. Proponente - UF .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Sig_UF
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '018. Proponente - CEP .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.gCEP_Formatado
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '019. Proponente - Telefone .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Num_Fone
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '020. Resumo: quantidade de itens.
    Print #pNumArqDados, pobjstcPed.ColecaoPedLoc.Count

    '021. Resumo: prêmio líquido.
    lAuxiliar1 = pobjstcPed.VAL_PR
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)



    '022. Resumo: custo da apólice.
    lAuxiliar1 = pobjstcPed.Val_Cust_Emis
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '023. Resumo: IOF.
    lAuxiliar1 = pobjstcPed.Val_IOF
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '024. Resumo: prêmio total a vista.
    lAuxiliar1 = pobjstcPed.gValorTotal
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '025. Forma de pagamento: número de parcelas.
    Print #pNumArqDados, pobjstcPed.Qtd_Parc

    '026. Forma de pagamento: parcelas antecipadas.
    Print #pNumArqDados, IIf(pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim, "1", "0")

    '027. Forma de pagamento: número do bloqueto.
    Print #pNumArqDados, pobjstcPed.objstcCtrlEmis.Num_Bloq & pobjstcPed.objstcCtrlEmis.DIG_BLOQUETO

    '028. Forma de pagamento: juros a. m.
    Print #pNumArqDados, Format(pobjstcPed.objstcCtrlEmis.Per_Juros_Neg, "standard") & "%"

    '029. Forma de pagamento: juros de mora.
    If pobjstcPed.objstcCtrlEmis.Qtd_Parc = 1 And pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
        Print #pNumArqDados, "0,00 %"
    Else
        Print #pNumArqDados, "4,50 %"
    End If

    'Calcula o fracionamento
    With pobjstcPed
        Call gM46V998.gfCalFracionamento(pobjstcPed, pobjstcPedLoc.Tip_Emissao, _
                                         lValorParc, lMensagem, gM46V111.gProJurIOF, .Cod_Reg, .Val_IOF, .Val_Parc_Min)
    End With



    If pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
        '030. Forma de pagamento: primeira parcela.
        Print #pNumArqDados, "R$" & Format(lValorParc(pobjstcPed.Qtd_Parc, 1), "standard")
        pobjstcPed.Val_PriParcela = lValorParc(pobjstcPed.Qtd_Parc, 1)

        '031. Forma de pagamento: demais parcelas.
        Print #pNumArqDados, "R$" & Format(lValorParc(pobjstcPed.Qtd_Parc, 2), "standard")
        pobjstcPed.val_DemParcelas = lValorParc(pobjstcPed.Qtd_Parc, 2)
    Else
        '030. Forma de pagamento: primeira parcela.
        Print #pNumArqDados, "R$" & Format(lValorParc(pobjstcPed.Qtd_Parc, 3), "standard")
        pobjstcPed.Val_PriParcela = lValorParc(pobjstcPed.Qtd_Parc, 3)

        '031. Forma de pagamento: demais parcelas.
        Print #pNumArqDados, "R$" & Format(lValorParc(pobjstcPed.Qtd_Parc, 4), "standard")
        pobjstcPed.val_DemParcelas = lValorParc(pobjstcPed.Qtd_Parc, 4)
    End If

    If pobjstcPed.Cod_Bco_Cobr <> 215 And pobjstcPed.Cod_Agen_Bas_Cor = 0 Then
        Print #pNumArqDados, "BANCO: " & pobjstcPed.Cod_Bco_Cobr & " - " & pobjstcPed.gDesc_Banco
    Else
        Print #pNumArqDados, " "
    End If

    '032 a 035. Observação 1.
    'a) Linha 1.
    lAuxiliar2 = "A aceitação do seguro estará sujeita à análise do risco. O registro deste plano na " & _
                 "SUSEP não implica,  por  parte"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'b) Linha 2.
    lAuxiliar2 = "da Autarquia, incentivo ou recomendação à sua comercialização. O segurado poderá " & _
                 "consultar a  situação  cadastral"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'c) Linha 3.
    lAuxiliar2 = "do seu corretor de seguros no site www.susep.gov.br, por meio do número de seu " & _
                 "registro na SUSEP, nome  completo, "
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'd) Linha 4.
    lAuxiliar2 = "CNPJ ou CPF."
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '036 a 047. Observação 2.
    'a) Linha 1.
    lAuxiliar2 = "Propõe(m) à YASUDA SEGUROS S.A. a realização do seguro, declarando conhecer " & _
                 "previamente e  aceitar  as  Condições"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'b) Linha 2.
    lAuxiliar2 = "Gerais do(s) anexo(s) que servem de base para esta proposta. Também presta(m)  as  " & _
                 "informações  solicitadas  pela"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'c) Linha 3.
    lAuxiliar2 = "YASUDA SEGUROS S.A., constantes do(s) anexo(s) a esta, para  fins  de  estudos  " & _
                 "da  concretização  do  respectivo"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'd) Linha 4.
    lAuxiliar2 = "contrato.  Assume(m)  o(s)  proponente(s),  nesta  oportunidade,  integral  " & _
                 "responsabilidade  pela  exatidão  das"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'e) Linha 5.
    lAuxiliar2 = "informações ora prestadas, autorizando a YASUDA SEGUROS S.A., caso aceite esta " & _
                 "proposta, a  emitir  a  respectiva"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'f) Linha 6.
    lAuxiliar2 = "apólice, cujo prêmio e eventuais despesas o(s) proponente(s) compromete(m)-se  a  " & _
                 "pagar,  tão  logo  lhe(s)  seja"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'g) Linha 7.
    lAuxiliar2 = "exigido. Em caso de recusa da proposta de seguro, os valores eventualmente " & _
                 "adiantados para  futuro  pagamento  de"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'h) Linha 8.
    lAuxiliar2 = "prêmio pelo(s) proponente(s) serão devolvidos integralmente, no prazo de 10 dias " & _
                 "corridos da data de sua  recusa."
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'i) Linha 9.
    lAuxiliar2 = "Ultrapassado este prazo, os valores sujeitam-se a atualização monetária pelo " & _
                 "IPC/FIPE, contados a partir da  data"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'j) Linha 10.
    lAuxiliar2 = "em que se tornarem exigíveis e juros moratórios de 0,5% (cinco décimos por cento) " & _
                 "ao mês, contados  a  partir  do"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'k) Linha 11.
    lAuxiliar2 = "primeiro dia posterior ao término do prazo fixado acima. Ambos calculados " & _
                 "pro-rata-die  até  a  data  da  efetiva"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'l) Linha 12.
    lAuxiliar2 = "restituição."
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '048. Local e data.
    lAuxiliar2 = "São Paulo, " & _
                 Format$(Now, "dd") & " de " & Format$(Now, "mmmm") & " de " & Format$(Now, "yyyy")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    Print #pNumArqDados, lObjstcCorr.gNom_Corretor

    '049. Unidade produtiva .
    Print #pNumArqDados, pobjstcPed.Cod_Unid_Prod

    '050. Produtor .
    Print #pNumArqDados, pobjstcPed.Cod_Prod

    '051. Corretor .

    Print #pNumArqDados, Format(lObjstcCorr.Cod_Corr, "0000000") & " 0000000 0000000"

    '052. Informação em código .
    Print #pNumArqDados, "C " & Format(lObjstcCorr.Per_Corr * 100, "0000")


    '053. Dados para correspondência - endereço .

    Print #pNumArqDados, pobjstcPed.Cod_Tip_Logr & " " & pobjstcPed.Nom_Logr & " " & pobjstcPed.Dsc_Compl_Logr

    '054. Dados para correspondência - bairro .
    lAuxiliar2 = " "
    Print #pNumArqDados, lAuxiliar2 & pobjstcPed.Nom_Bairro

    '055. Dados para correspondência - cidade .
    lAuxiliar2 = " "
    Print #pNumArqDados, lAuxiliar2 & pobjstcPed.Nom_Cid

    '056. Dados para correspondência - UF .
    lAuxiliar2 = " "
    Print #pNumArqDados, lAuxiliar2 & pobjstcPed.Sig_UF

    '057. Dados para correspondência - CEP .
    Print #pNumArqDados, pobjstcPed.gCEP_Formatado

    If pobjstcPed.ColecaoPedVenc.Count > 0 Then
        Print #pNumArqDados, "VENCIMENTO DAS PARCELAS"
        lAuxiliar2 = ""
        For Each lObjstcVenc In pobjstcPed.ColecaoPedVenc
            If lObjstcVenc.Num_Parc <= pobjstcPed.Qtd_Parc Then
                If lObjstcVenc.Num_Parc < 6 Then
                    lAuxiliar2 = lAuxiliar2 & "(" & lObjstcVenc.Num_Parc & "ª-" & gM46V999.gfForData(2, lObjstcVenc.DAT_VENC) & ")"
                Else
                    If lObjstcVenc.Num_Parc = 6 Then
                        Print #pNumArqDados, lAuxiliar2
                        lAuxiliar2 = ""
                    End If
                    lAuxiliar2 = lAuxiliar2 & "(" & lObjstcVenc.Num_Parc & "ª-" & gM46V999.gfForData(2, lObjstcVenc.DAT_VENC) & ")"
                End If
            End If
        Next lObjstcVenc
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        If pobjstcPed.Qtd_Parc < 6 Then
            Print #pNumArqDados, " "
        End If
    Else
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
    End If

    mfCriDados_MULTPRO3_V_0 = True
End Function
Private Function mfCriDados_MULTPRO4_V_0(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                         pPagAtual As Integer, pTotalPag As Integer) As Boolean
'Função: cria arquivo de dados - proposta coletiva - dados do item

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lCodCobertura As Integer    'Código da cobertura.
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lQuant As Integer    'Quantidade de itens.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lTamanho As Integer    'Tamanho.
    Dim lobjstcBem As stcA46V715B    'Classe (stc) de bens).
    Dim lobjstcCob As stcA46V704B    'Classe (stc) de coberturas.
    Dim lobjstcCoss As stcA46V706B    'Classe (stc) de cosseguro.
    Dim lObjstcCorr As stcA46V705B    'Classe (stc) de corretor.
    Dim lobjstcOutros As stcA46V716B    'Classe (stc) de Outros.
    Dim lobjstcSeg As stcA46V716B    'Classe (stc) de seguros.
    Dim lValTotCober As Double    'Soma dos valores de coberturas.
    Dim lobjstcPedLoc As stcA46V708B    'Classe (stc) de itens.

    Dim lDescFranquias() As String
    Dim lDescImportante() As String

    Dim lobjstcClau As stcA46V703B    'Objeto da classe de estrutura stcA46V703B(TAB_PED_CLAU).
    Dim lblnAchou As Boolean

    mfCriDados_MULTPRO4_V_0 = False
    lQuant = 1
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        '001 e 002. Título do documento.
        Print #pNumArqDados, "ANEXO DA PROPOSTA DE SEGURO YASUDA"
        Print #pNumArqDados, UCase(pobjstcPed.gDesc_Ramo)

        '003. Emissão.
        Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
        lQuant = lQuant + 1
        '004. Número da página.
        Print #pNumArqDados, "Página " & lQuant & "/" & pTotalPag

        '005. Número do pedido .
        Print #pNumArqDados, pobjstcPed.Num_Ped

        '006. Número do item .
        Print #pNumArqDados, lobjstcPedLoc.Num_Item

        '007. Corretor (código e nome) .
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
        Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2


        '008. Tarifa .
        Print #pNumArqDados, pobjstcPed.Cod_Produto

        '009. Nosso número .
        Print #pNumArqDados, pobjstcPed.NOSSO_NUMERO

        '010. Local em risco - logradouro .
        lAuxiliar2 = lobjstcPedLoc.Cod_Tip_Logr_Risco & " " & lobjstcPedLoc.Nom_Logr_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '011. Local em risco - número .
        lAuxiliar2 = " "
        Print #pNumArqDados, lAuxiliar2 & lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco

        '012. Local em risco - complemento .
        lAuxiliar2 = " "
        Print #pNumArqDados, lAuxiliar2 & lobjstcPedLoc.Dsc_Compl_Risco

        '013. Local em risco - bairro .
        lAuxiliar2 = " "
        Print #pNumArqDados, lAuxiliar2 & lobjstcPedLoc.Nom_Bairro_Risco

        '014. Local em risco - cidade .
        lAuxiliar2 = " "
        Print #pNumArqDados, lAuxiliar2 & lobjstcPedLoc.Nom_Cid_Risco

        '015. Local em risco - UF .
        lAuxiliar2 = " "
        Print #pNumArqDados, lAuxiliar2 & lobjstcPedLoc.Sig_UF_Risco

        '016. Local em risco - CEP .
        Print #pNumArqDados, lobjstcPedLoc.gCEP_Formatado

        '017. Dados para inspeção - nome da pessoa para contato .
        lAuxiliar2 = " "
        Print #pNumArqDados, lAuxiliar2 & lobjstcPedLoc.objstcPedInsp.Pess_Cont

        '018. Dados para inspeção - telefone para contato .
        lAuxiliar2 = " "
        Print #pNumArqDados, lAuxiliar2 & lobjstcPedLoc.objstcPedInsp.Num_Tel

        '019. Dados do seguro - classificação .
        lAuxiliar2 = lobjstcPedLoc.Cod_Clasf & "-" & lobjstcPedLoc.Cod_Compl_Clasf & "- " & lobjstcPedLoc.Dsc_Clasf
        lAuxiliar2 = lAuxiliar2 & String(15, " ") & " PLANO: " & pobjstcPed.Desc_Plano
        lAuxiliar2 = lAuxiliar2 & " - PRODUTO: " & Format(lobjstcPedLoc.Cod_Produto, "0000")
        Print #pNumArqDados, lAuxiliar2

        '020. Dados do seguro - construção .
        'Ficha 1000305
        If lobjstcPedLoc.Cod_Ramo = e_CodRamo_Empresarial Then
            Select Case lobjstcPedLoc.Cod_Constr
            Case e_CodContr_Superior
                lAuxiliar2 = "Superior/sólida"
            Case e_CodContr_Inferior
                lAuxiliar2 = "Inferior"
            Case e_CodContr_Mista
                lAuxiliar2 = "Mista"
            Case Else
                lAuxiliar2 = "Superior/sólida"
            End Select
        Else
            Select Case lobjstcPedLoc.Cod_Constr
            Case e_CodContr_Superior
                lAuxiliar2 = "Superior/sólida"
            Case e_CodContr_Inferior
                lAuxiliar2 = "Inferior/mista"
            Case Else
                lAuxiliar2 = "Superior/sólida"
            End Select
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


        '021. Dados do seguro - região .
        lAuxiliar2 = lobjstcPedLoc.Cod_Regiao & " - " & Left(lobjstcPedLoc.gDesc_Regiao, 38)

        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '022. Dados do seguro - vigência .
        lAuxiliar2 = "de " & gM46V999.gfForData(2, pobjstcPed.Dat_Ini_Vig) & _
                   " a " & gM46V999.gfForData(2, pobjstcPed.Dat_Ter_Vig)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


        '023. Dados do seguro - valor em risco .
        lAuxiliar2 = "R$ " & Format(lobjstcPedLoc.Val_Vr, "standard")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '024. Dados do seguro - tipo de emissão .
        lAuxiliar2 = lobjstcPedLoc.gDesc_TipEmissao
        Select Case lobjstcPedLoc.Tip_Emissao
        Case e_TipEmissao_RenYasuda Or e_TipEmissao_Endosso
            lAuxiliar2 = lAuxiliar2 & "  Apólice: " & lobjstcPedLoc.Num_Apol_Anterior & "  Item: " & lobjstcPedLoc.Num_Item_Anterior & "  Sinistro: " & IIf(lobjstcPedLoc.Ind_Sinistro = e_SimNao_Sim, "Sim", "Não")
        End Select
        Print #pNumArqDados, lAuxiliar2

        '025 e 026. Coberturas contratadas - desconto - título.
        If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            '035 e 036. Coberturas contratadas - título.
            lAuxiliar2 = "Desconto de"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            lAuxiliar2 = "experiência"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Else
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
        End If


        '027 a 158. Coberturas contratadas - linhas de 01 a 22.
        lCountLine = 0
        lValTotCober = 0
        For Each lobjstcCob In lobjstcPedLoc.ColecaoPedCobert
            lCountLine = lCountLine + 1
            'Código
            Print #pNumArqDados, Format$(lobjstcCob.Cod_Cobert, "0000")
            'Descrição
            Print #pNumArqDados, Left(lobjstcCob.gNomeCobertura, 48)

            'Valor da IS
            lAuxiliar1 = lobjstcCob.Val_IS
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 14 Then
                lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            'd) Desconto comercial .
            lAuxiliar1 = lobjstcCob.Per_Desc_Coml
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 10 Then
                lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
                'e) Desconto de experiência .
                lAuxiliar1 = lobjstcCob.Per_Desc_Exp
                lAuxiliar2 = Format$(lAuxiliar1, "standard")
                lTamanho = Len(lAuxiliar2)
                If lTamanho < 10 Then
                    lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
                End If
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            Else
                Print #pNumArqDados, " "
            End If

            'f) Prêmio líquido .
            lAuxiliar1 = lobjstcCob.VAL_PR
            lValTotCober = lValTotCober + lobjstcCob.VAL_PR
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 14 Then
                lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Next lobjstcCob
        Do While lCountLine < 22
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
        Loop

        '159. Coberturas contratadas - soma.
        lAuxiliar1 = lValTotCober
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '160. Coberturas contratadas - prêmio líquido.
        If lValTotCober < gM46V999.gProPreMinParcela Then
            lValTotCober = gM46V999.gProPreMinParcela
        End If

        lAuxiliar1 = lValTotCober
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '161. Informação cifrada.
        '   B: 25 / A: 50 / P: 10 / CO: 99999 / C: 99999  onde:
        '   B = bônus
        '   A = aparelho
        '   P = progressivo
        '   CO = coeficiente de comissão
        '   C = coeficiente LMG/VR

        'Ficha 900740 e 901021 - zerar bônus das renovações dos planinhos
        '2001 a 2010 - Correta
        '2011 a 2020 - GalCorr
        '2021 a 2031 - Mercedes
        '2032 a 2041 - National
        '2042 a 2049 - Manaus
        '2050 - Ballarin
        '2051 a 2060 - Marsh
        '2061 a 2070 - NSK
        '2071 a 2080 - Ajinomoto
        '2081 a 2090 - Nissin
        '2091 a 2100 - Panasonic
        '2101 - Fanel
        '2102 - JLA

        'Ficha 901673
        '2111 a 2120 - Embraer
        'Ficha 901772
        '2121 a 2130 - Bardella

        If (gM46V111.gCod_plano >= 2001 And gM46V111.gCod_plano <= 2049) Or (gM46V111.gCod_plano >= 2051 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
            lAuxiliar2 = "B:00"
        Else
            lAuxiliar2 = "B:" & Format(lobjstcPedLoc.Per_Bonus, "00")
        End If

        lAuxiliar2 = lAuxiliar2 & "/A:" & Format(lobjstcPedLoc.Per_Desc_Apar, "00")

        'Ficha 900740 e 901673 e 901772
        If (gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031) Or (gM46V111.gCod_plano >= 2061 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
            lAuxiliar2 = lAuxiliar2 & "/P:00"
        Else
            lAuxiliar2 = lAuxiliar2 & "/P:" & Format(lobjstcPedLoc.Per_Desc_Progres, "00")
        End If

        'lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(pObjstcPed.objstcCtrlEmis.Per_Corr + pObjstcPed.objstcCtrlEmis.Per_Corr_2 + pObjstcPed.objstcCtrlEmis.Per_Corr_3, "00.00")
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
        lAuxiliar1 = lObjstcCorr.Per_Corr

        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(2)
            lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(3)
                lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            End If
        End If

        lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(lAuxiliar1 * 100, "0000")

        lAuxiliar2 = lAuxiliar2 & "/C:" & Format(Round(lobjstcPedLoc.Val_Coef_IS_Vr, 4) * 10000, "00000")
        lAuxiliar2 = lAuxiliar2 & "/PRBO19"

        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '162 a 171. Segurança.
        '           O texto pode ser customizado de acordo com a necessidade.
        '           Temos 10 linhas disponíveis e cada linha cabe 70 caracteres (contando com o "(X)").
        '           O "X" deve ser colocado de acordo com a base de dados.
        If pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Then
            'a) Linha 01.
            Print #pNumArqDados, "SEGURANÇA"
            If lobjstcPedLoc.Cod_Iden_Segur_01_V = "S" Then
                lAuxiliar2 = "(S) "
            Else
                lAuxiliar2 = "(N) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Vigilância exclusiva e permanente do imóvel entendendo-se ainda"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'b) Linha 02.
            lAuxiliar2 = "    como tal, porteiros fixos."
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'c) Linha 03.
            If lobjstcPedLoc.Cod_Iden_Segur_02_A = "S" Then
                lAuxiliar2 = "(S) "
            Else
                lAuxiliar2 = "(N) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Alarmes."
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'd) Linha 04.
            If lobjstcPedLoc.Cod_Iden_Segur_03_P = "S" Then
                lAuxiliar2 = "(S) "
            Else
                lAuxiliar2 = "(N) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Porteiros eletrônicos."
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'e) Linha 05.
            If lobjstcPedLoc.Cod_Iden_Segur_04_I = "S" Then
                lAuxiliar2 = "(S) "
            Else
                lAuxiliar2 = "(N) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Inexistência de terreno baldio, imóvel desocupado ou em construção"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'f) Linha 06.
            lAuxiliar2 = "    na vizinhança do risco."
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'g) Linha 07.
            If lobjstcPedLoc.Cod_Iden_Segur_05_G = "S" Then
                lAuxiliar2 = "(S) "
            Else
                lAuxiliar2 = "(N) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Grades metálicas de proteção em todas as janelas e portas de vidro"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'h) Linha 08.
            lAuxiliar2 = "    que dão acesso ao interior do imóvel desde que também possuam"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'i) Linha 09.
            lAuxiliar2 = "    alarme."
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'j) Linha 10.
            lAuxiliar2 = " "
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Else
            Print #pNumArqDados, " "
            For lI = 1 To 10
                Print #pNumArqDados, String(71, "X")
            Next lI
        End If

        '194 a 198. Proteção.
        '           Para residencial, sairá em branco.
        If pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Or pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            '           O texto pode ser customizado de acordo com a necessidade.
            '           Temos 4 linhas disponíveis e cada linha cabe 39 caracteres (contando com o "(X)").
            '           O "X" deve ser colocado de acordo com a base de dados.
            'a) Título.
            lAuxiliar2 = "PROTEÇÃO"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'b) Linha 01.
            If lobjstcPedLoc.Cod_Iden_Protec_01_E = "S" Then
                lAuxiliar2 = "(S) "
            Else
                lAuxiliar2 = "(N) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Extintor"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'c) Linha 02.
            If lobjstcPedLoc.Cod_Iden_Protec_02_H = "S" Then
                lAuxiliar2 = "(S) "
            Else
                lAuxiliar2 = "(N) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Hidrantes"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'd) Linha 03.
            If lobjstcPedLoc.Cod_Iden_Protec_03_S = "S" Then
                lAuxiliar2 = "(S) "
            Else
                lAuxiliar2 = "(N) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Sprinkler (chuveiros automáticos)"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'e) Linha 04.
            If lobjstcPedLoc.Cod_Iden_Protec_04_D = "S" Then
                lAuxiliar2 = "(S) "
            Else
                lAuxiliar2 = "(N) "
            End If
            lAuxiliar2 = lAuxiliar2 & "Detecção e alarme"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Else
            Print #pNumArqDados, " "
            For lI = 1 To 4
                Print #pNumArqDados, String(40, "X")
            Next lI
        End If

        '177 a 181. Declarações/Observações .
        lCountLine = 0
        If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Or pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Then
            lCountLine = lCountLine + 1
            lAuxiliar2 = "Valor em risco declarado: R$ " & Format(lobjstcPedLoc.Val_Vr, "standard")
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If

        'Alterado em 23/03/2007 - Ficha: 0700328
        If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            lCountLine = lCountLine + 1
            lAuxiliar2 = "-Bens de Terceiro em Poder do Segurado: Para ser passível de cobertura, deverá ser destacado do Valor em Risco"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            lCountLine = lCountLine + 1
            lAuxiliar2 = " Total o Valor em Risco destes bens"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'Alterado em 24/04/2007 - Ficha: 0700328
        If (lobjstcPedLoc.Cod_Clasf = gM46V999.gfBuscarDePara(18, 2) And lobjstcPedLoc.Cod_Compl_Clasf = gM46V999.gfBuscarDePara(18, 2, 10)) Or _
           (lobjstcPedLoc.Cod_Clasf = gM46V999.gfBuscarDePara(18, 2) And lobjstcPedLoc.Cod_Compl_Clasf = gM46V999.gfBuscarDePara(18, 2, 11)) Or _
           (lobjstcPedLoc.Cod_Clasf = gM46V999.gfBuscarDePara(19, 2) And lobjstcPedLoc.Cod_Compl_Clasf = gM46V999.gfBuscarDePara(19, 2, 10)) Then
                lCountLine = lCountLine + 1
                lAuxiliar2 = " ******    RISCOS NOMEADOS - RAMO 96    ******"
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
        End If


        lCountLine = lCountLine + 1
        lAuxiliar2 = "Existem outros seguros cobrindo os mesmos bens? " & IIf(lobjstcPedLoc.ColecaoPedSeguro.Count > 0, "SIM", "NÃO")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        lblnAchou = False

        If pobjstcPed.Cod_Ramo <> e_CodRamo_Condominio Then
            lCountLine = lCountLine + 1
            lAuxiliar2 = "Deseja a inclusão de Cláusula Beneficiária? "
            'If lObjstcPedLoc.objstcPedClau.Cod_Clau > 0 Then
            '    lAuxiliar2 = lAuxiliar2 & "SIM - " & Trim(Mid(lObjstcPedLoc.objstcPedClau.Descricao, 15, 60)) & " - " & Left(lObjstcPedLoc.objstcPedClau.Descricao, 14)
            'Else
            '    lAuxiliar2 = lAuxiliar2 & "NÃO"
            'End If
            Set lobjstcClau = New stcA46V703B
            Set lobjstcClau.M46V999 = gM46V999

            If lobjstcPedLoc.ColecaoPedClau.Count > 0 Then
                Set lobjstcClau = lobjstcPedLoc.ColecaoPedClau.item(1)
                If lobjstcClau.Cod_Clau > 0 Then
                    lblnAchou = True
                End If
            End If
            If lblnAchou Then
                lAuxiliar2 = lAuxiliar2 & "SIM - " & Trim(Mid(lobjstcClau.Descricao, 15, 60)) & " - " & Left(lobjstcClau.Descricao, 14)
            Else
                lAuxiliar2 = lAuxiliar2 & "NÃO"
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
        If pobjstcPed.Cod_Ramo = e_CodRamo_Residencial Then
            lCountLine = lCountLine + 1
            lAuxiliar2 = "Existem bens no interior da residência com valor unitário superior a R$ 5.000,00? "
            If lobjstcPedLoc.ColecaoPedLocBem.Count > 0 Then
                lAuxiliar2 = lAuxiliar2 & "NÃO."
                For Each lobjstcBem In lobjstcPedLoc.ColecaoPedLocBem
                    If lobjstcBem.Val_Bem > 5000 Then
                        lAuxiliar2 = Replace(lAuxiliar2, "NÃO.", "SIM.")
                        Exit For
                    End If
                Next lobjstcBem
            Else
                lAuxiliar2 = lAuxiliar2 & "NÃO."
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            lCountLine = lCountLine + 1
            lAuxiliar2 = "Número de empregados na residência: " & lobjstcPedLoc.Qtd_Empreg
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        End If

        Do While (lCountLine < 5)
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "
        Loop


        '182 a 189. Franquias.
        Call mpDescricaoFranquia_V_0(lobjstcPedLoc, pobjstcPed, lDescFranquias, 8)

        'a) Linha 1.
        For lI = 0 To 7
            lAuxiliar2 = lDescFranquias(lI)
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Next lI

        '231 a 240. Importante. (10 linhas)
        Call mpTextoImportante(lobjstcPedLoc, lDescImportante, "MULTPRO4")
        lCountLine = 0
        For lI = 0 To UBound(lDescImportante) - 1
            lAuxiliar2 = lDescImportante(lCountLine)
            If gM46V999.gfPreenchido(lAuxiliar2) Then
                lCountLine = lCountLine + 1
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
        Next lI

        If gM46V999.gCodProduto >= 958 And lobjstcPedLoc.Cob_Bens = True And Not mfVerDigBens() Then
            lCountLine = 0
            For Each lobjstcCob In lobjstcPedLoc.ColecaoPedCobert
                If gM46V999.gCod_Ramo = e_CodRamo_Empresarial And lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(94, 1) And lobjstcPedLoc.Dig_Bens = False Then
                    lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores. Opção escolhida pelo corretor: "
                    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                    lCountLine = lCountLine + 1
                    lAuxiliar2 = "Entrega da relação juntamente com a proposta de seguro."
                    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                    lCountLine = lCountLine + 1
                End If
                If gM46V999.gCod_Ramo = e_CodRamo_Residencial And (lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(27, 1) Or lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(29, 1) Or _
                                                                   lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(30, 1) Or lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(31, 1)) And lobjstcPedLoc.Dig_Bens = False Then
                    lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores. Opção escolhida pelo corretor: "
                    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                    lCountLine = lCountLine + 1
                    lAuxiliar2 = "Entrega da relação juntamente com a proposta de seguro."
                    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                    lCountLine = lCountLine + 1
                End If
            Next lobjstcCob
        End If

        Call mpTextoObs(lobjstcPedLoc, lDescImportante)
        For lI = 0 To UBound(lDescImportante) - 1
            lAuxiliar2 = lDescImportante(lI)
            If gM46V999.gfPreenchido(lAuxiliar2) And lCountLine < 5 Then
                lCountLine = lCountLine + 1
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
        Next lI
        '046 a 051 . Cosseguro
        lAuxiliar2 = ""
        If pobjstcPed.ColecaoPedCoss.Count > 0 And lCountLine < 10 Then
            For Each lobjstcCoss In pobjstcPed.ColecaoPedCoss
                lAuxiliar2 = lAuxiliar2 & " - ( " & lobjstcCoss.Cod_Coss & " - " & lobjstcCoss.Per_Coss & "% )"
            Next lobjstcCoss
            lCountLine = lCountLine + 1
            Print #pNumArqDados, "COSSEGURO: " & lAuxiliar2
        End If

        If lobjstcPedLoc.ColecaoPedSeguro.Count > 0 And lCountLine < 10 Then
            lAuxiliar2 = ""
            For Each lobjstcOutros In lobjstcPedLoc.ColecaoPedSeguro
                If Not gM46V999.gfPreenchido(lAuxiliar2) Then
                    lAuxiliar2 = "OUTROS SEGUROS:(" & lobjstcOutros.Cod_Seguradora & "-Ap:" & lobjstcOutros.Num_Apol_Ant & "nº" & lobjstcOutros.Num_Item_Ant & ")"
                Else
                    lAuxiliar2 = lAuxiliar2 & "-(" & lobjstcOutros.Cod_Seguradora & "-Ap:" & lobjstcOutros.Num_Apol_Ant & "nº" & lobjstcOutros.Num_Item_Ant & ")"
                End If
            Next lobjstcOutros
            If gM46V999.gfPreenchido(lAuxiliar2) And lCountLine < 10 Then
                lCountLine = lCountLine + 1
                Print #pNumArqDados, lAuxiliar2
            End If
        End If
        Do While lCountLine < 10
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        Loop
    Next lobjstcPedLoc
    mfCriDados_MULTPRO4_V_0 = True
End Function
Private Function mfCriDados_MULTPROB_V_0(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B) As Boolean
'Função: cria arquivo de dados - proposta coletiva - dados do item

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lCodCobertura As Integer    'Código da cobertura.
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lQuant As Integer    'Quantidade de itens.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lTamanho As Integer    'Tamanho.
    Dim lobjstcBem As stcA46V715B    'Classe (stc) de bens).
    Dim lObjstcCorr As stcA46V705B    'Classe (stc) de corretor.
    Dim lobjstcPedLoc As stcA46V708B    'Classe (stc) de itens.

    mfCriDados_MULTPROB_V_0 = False
    lQuant = 0
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        If lobjstcPedLoc.ColecaoPedLocBem.Count > 0 Then
            '001 . Título do documento.
            Print #pNumArqDados, "ANEXO DA PROPOSTA DE SEGURO YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

            '002. Código do Processo Susep.
            '     Observar que este número deve ser de acordo com o título do seguro:
            '     RESIDENCIAL, CONDOMÍNIO ou EMPRESARIAL.
            lAuxiliar2 = " "
            Select Case pobjstcPed.Cod_Ramo
            Case e_CodRamo_Condominio
                lAuxiliar2 = "Processo Susep 15414.004604/2004-72"
            Case e_CodRamo_Empresarial
                lAuxiliar2 = "Processo Susep 15414.004605/2004-17"
            Case e_CodRamo_Residencial
                lAuxiliar2 = "Processo Susep 15414.003659/2004-65"
            End Select
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            '003. Emissão.
            Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

            lQuant = lQuant + 1
            '004. Número da página.
            Print #pNumArqDados, "Página " & lQuant

            '005. Número do pedido .
            Print #pNumArqDados, pobjstcPed.Num_Ped & "      Item: " & lobjstcPedLoc.Num_Item

            '007. Corretor (código e nome) .
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
            lAuxiliar1 = lObjstcCorr.Cod_Corr
            lAuxiliar2 = lObjstcCorr.gNom_Corretor
            Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

            '008. Tarifa .
            Print #pNumArqDados, pobjstcPed.Cod_Produto

            '009. Nosso número .
            Print #pNumArqDados, pobjstcPed.NOSSO_NUMERO

            Print #pNumArqDados, "Bens:"
            lCountLine = 1
            For Each lobjstcBem In lobjstcPedLoc.ColecaoPedLocBem
                lCountLine = lCountLine + 1
                lAuxiliar2 = Format(lCountLine, "000") & " - " & Left(lobjstcBem.Dsc_Bem, 100) & String(100 - Len(lobjstcBem.Dsc_Bem), ".") & "R$ " & Format(lobjstcBem.Val_Bem, "standard")
                Print #pNumArqDados, lAuxiliar2
            Next lobjstcBem

            lAuxiliar2 = " "
            lCountLine = lCountLine + 1
            Print #pNumArqDados, lAuxiliar2

            lAuxiliar2 = " "
            lCountLine = lCountLine + 1
            Print #pNumArqDados, lAuxiliar2

            lAuxiliar2 = " "
            lCountLine = lCountLine + 1
            Print #pNumArqDados, lAuxiliar2

            lAuxiliar2 = " "
            lCountLine = lCountLine + 1
            Print #pNumArqDados, lAuxiliar2

            'Local e data
            lAuxiliar2 = "São Paulo, " & _
                         Format$(Now, "dd") & " de " & Format$(Now, "mmmm") & " de " & Format$(Now, "yyyy")
            lCountLine = lCountLine + 1
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


            lAuxiliar2 = "__________________________________       _________________________________       _________________________________"
            lCountLine = lCountLine + 1
            Print #pNumArqDados, lAuxiliar2

            lAuxiliar2 = "          LOCAL E DATA                     ASSINATURA DO(A) PROPONENTE(A)           ASSINATURA DO(A) CORRETOR(A)  "
            lCountLine = lCountLine + 1
            Print #pNumArqDados, lAuxiliar2

            'Nome do corretor
            lCountLine = lCountLine + 1
            lAuxiliar2 = "                                                                                      " & lObjstcCorr.gNom_Corretor
            Print #pNumArqDados, lAuxiliar2

            Do While lCountLine < 75
                Print #pNumArqDados, " "
                lCountLine = lCountLine + 1
            Loop
        End If
    Next lobjstcPedLoc
    If lQuant > 0 Then
        mfCriDados_MULTPROB_V_0 = True
    End If
End Function

Private Sub mpTextoObs(pobjstcPedLoc As stcA46V708B, pTexto() As String)
    Dim lI As Integer
    Dim lObs As Variant
    Dim lCount As Integer
    Dim lAuxiliar As String
    Dim lPos As Long

    MsgBox "FUNCIONALIDADE ANTIGA - DESATIVADA"
    Exit Sub

    ReDim pTexto(10)
    lI = 0
    If gM46V999.gfPreenchido(pobjstcPedLoc.objstcPedObs.Observacao) Then
        pobjstcPedLoc.objstcPedObs.Observacao = pobjstcPedLoc.objstcPedObs.Observacao & "&%"
    End If
    lObs = Split(pobjstcPedLoc.objstcPedObs.Observacao, "&%")
    For lCount = 0 To UBound(lObs)
        lAuxiliar = lObs(lCount)
        If gM46V999.gfPreenchido(lAuxiliar) Then
            If Len(lAuxiliar) < 128 Then
                pTexto(lI) = lAuxiliar
                lI = lI + 1
                If lI > 10 Then Exit For
            Else
                lPos = 1
                Do
                    pTexto(lI) = Mid(lAuxiliar, lPos, 128)
                    lI = lI + 1
                    If Len(Mid(lAuxiliar, lPos, 128)) < 128 Then Exit Do
                    lPos = lPos + 128
                Loop
            End If
        End If
    Next lCount

End Sub

Private Sub mpTextoImportante(pobjstcPedLoc As stcA46V708B, pTexto() As String, pForm As String)
    Dim lAuxiliar As String
    Dim lobjstcCobert As stcA46V704B
    Dim lobjstcOutros As stcA46V716B
    Dim lObjstcPed As stcA46V702B
    Dim lI As Integer
    Dim lCod_Ramo As e_CodRamo

    MsgBox "FUNCIONALIDADE ANTIGA - DESATIVADA"

    ReDim pTexto(10)
    lCod_Ramo = pobjstcPedLoc.Cod_Ramo
    pTexto(lI) = "Cotação e L.M.G. para simples consulta, não implica na aceitação automática;"
    lI = lI + 1
    If lCod_Ramo = e_CodRamo_Residencial Then
        'pTexto(lI) = "Assistência 24 horas gratuita, restrita para apólices com no máximo 3 (três) itens."
        pTexto(lI) = "Assistência Domiciliar 24 Horas gratuita (exceto apólices com mais de 3 itens e residências eventuais)."
        lI = lI + 1
    ElseIf lCod_Ramo <> e_CodRamo_Roubos Then
        pTexto(lI) = "Assistência 24 horas gratuita, restrita para apólice individual."
        lI = lI + 1
    End If
'	 Tarifa 11.12
'    pTexto(lI) = "Na Razão Social não deverá constar a expressão PCPT, conforme informativo Yasuda  01/2004 de 13/04/2004."
'    lI = lI + 1
    For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
        Select Case lobjstcCobert.Cod_Cobert
        Case gM46V999.gfBuscarDePara(39, 1), gM46V999.gfBuscarDePara(40, 1), gM46V999.gfBuscarDePara(41, 1), _
            gM46V999.gfBuscarDePara(42, 1), gM46V999.gfBuscarDePara(43, 1), gM46V999.gfBuscarDePara(44, 1), _
            gM46V999.gfBuscarDePara(47, 1), gM46V999.gfBuscarDePara(50, 1)
            pTexto(lI) = "Período Indenitário: 12 meses. Para prazos inferiores, solicitar cotação à Unidade de produto"
            lI = lI + 1
            Exit For
        End Select
    Next lobjstcCobert

    '### INSPECAO
    If lObjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        pTexto(lI) = "A aceitação desta proposta está condicionada a análise da Yasuda Seguros até 15 dias de seu recebimento. Este prazo ficará "
        lI = lI + 1
        pTexto(lI) = "suspenso até a entrega de documentos complementares e/ou data da inspeção do risco."
        lI = lI + 1
    Else
        If pobjstcPedLoc.Insp_Obr = True Or pobjstcPedLoc.Ind_Sin_Inspec > 0 Then
            pTexto(lI) = "A aceitação deste seguro está condicionada a análise da inspeção do risco."
            lI = lI + 1
            pTexto(lI) = "A aceitação deste seguro está condicionada a análise da Yasuda Seguros."
            lI = lI + 1
        Else
            pTexto(lI) = "A aceitação deste seguro está condicionada a análise da Yasuda Seguros."
            lI = lI + 1
        End If
    End If

    'Ficha 901562
    If gM46V999.gCodProduto >= 958 And (pobjstcPedLoc.Cod_Plano = 2050 Or _
                                        pobjstcPedLoc.Cod_Plano = 2101 Or _
                                        pobjstcPedLoc.Cod_Plano = 2102) Then
        pTexto(lI) = "O presente seguro destina-se a cobrir exclusivamente o prédio (imóvel)."
        lI = lI + 1
    End If

    'Ficha 901378
    If pForm <> "MULTPRO2" Then
        pTexto(lI) = "O registro deste plano na SUSEP não implica, por parte da Autarquia, incentivo ou recomendação a sua comercialização."
        lI = lI + 1
        pTexto(lI) = "O segurado poderá consultar a situação cadastral de seu corretor de seguros, no site www.susep.gov.br, por meio de seu"
        lI = lI + 1
        pTexto(lI) = "registro na SUSEP, nome completo, CNPJ ou CPF."
        lI = lI + 1
    End If

    'Ficha 901346
    If Trim(pobjstcPedLoc.Cod_Clasf) = gM46V999.gfBuscarDePara("022", 2) And Trim(pobjstcPedLoc.Cod_Compl_Clasf) = gM46V999.gfBuscarDePara("022", 2, "13") Then
        For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
            Select Case lobjstcCobert.Cod_Cobert
            Case 56
                pTexto(lI) = "Cobertura 56 - Pátios - até 100 Km: Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros, a "
                lI = lI + 1
                pTexto(lI) = "Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 1.500,00/R$ 3.000,00 - "
                lI = lI + 1
                pTexto(lI) = "Nacionais/Importados."
                lI = lI + 1
            Case 57
                If lI > 7 Then Exit For
                pTexto(lI) = "Cobertura 57 - R C Concessionárias - até 100 Km: Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de "
                lI = lI + 1
                pTexto(lI) = "terceiros, a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 1.500,00/R$ 3.000,00 - "
                lI = lI + 1
                pTexto(lI) = "Nacionais/Importados."
                lI = lI + 1
            Case 122
                If lI > 7 Then Exit For
                pTexto(lI) = "Cobertura 122 - Pátios - até 200 Km: Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros, a "
                lI = lI + 1
                pTexto(lI) = "Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 1.500,00/R$ 3.000,00 - Nacionais/"
                lI = lI + 1
                pTexto(lI) = "Importados."
            Case 123
                If lI > 7 Then Exit For
                pTexto(lI) = "Cobertura 123 - Pátios - até 300 Km : Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros, a "
                lI = lI + 1
                pTexto(lI) = "Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 1.500,00/R$ 3.000,00 - Nacionais/"
                lI = lI + 1
                pTexto(lI) = "Importados."
            Case 124
                If lI > 7 Then Exit For
                pTexto(lI) = "Cobertura 124 - R C Concessionárias - até 200 Km:  Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de "
                lI = lI + 1
                pTexto(lI) = "terceiros, a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 2.000,00/R$ 3.000,00 - "
                lI = lI + 1
                pTexto(lI) = "Nacionais/Importados."
            Case 125
                If lI > 7 Then Exit For
                pTexto(lI) = "Cobertura 125 - R C Concessionárias - até 300 Km:  Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de "
                lI = lI + 1
                pTexto(lI) = "terceiros, a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 2.000,00/R$ 3.000,00 - "
                lI = lI + 1
                pTexto(lI) = "Nacionais/Importados."
            End Select
        Next lobjstcCobert
    End If

End Sub
Private Function mfCriDados_Debito(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B) As Boolean
'Função: cria arquivo de dados - autorização para débito em conta.

'Parâmetros de entrada...pNumArqDados.....Número do arquivo de dados.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lLocData As String    'Local/Data.

    mfCriDados_Debito = False

    '1. Prepara dados para imprimir.
    'a) Local e data.
    lLocData = "São Paulo, " & Format$(gM46V999.gHojDia, "00") & " de "
    Select Case gM46V999.gHojMes
    Case 1
        lLocData = lLocData & "janeiro"
    Case 2
        lLocData = lLocData & "fevereiro"
    Case 3
        lLocData = lLocData & "março"
    Case 4
        lLocData = lLocData & "abril"
    Case 5
        lLocData = lLocData & "maio"
    Case 6
        lLocData = lLocData & "junho"
    Case 7
        lLocData = lLocData & "julho"
    Case 8
        lLocData = lLocData & "agosto"
    Case 9
        lLocData = lLocData & "setembro"
    Case 10
        lLocData = lLocData & "outubro"
    Case 11
        lLocData = lLocData & "novembro"
    Case 12
        lLocData = lLocData & "dezembro"
    End Select
    lLocData = lLocData & " de " & Format$(gM46V999.gHojAno, "0000")

    '2. Dados principais.
    'NOME - Nome.
    If gM46V999.gfPreenchido(pobjstcPed.Nom_Segurado) Then
        lAuxiliar = pobjstcPed.Nom_Segurado
    Else
        lAuxiliar = " "
    End If
    Print #pNumArqDados, lAuxiliar
    'CNPJ - CNPJ/CPF.
    Print #pNumArqDados, pobjstcPed.gCNPJ_CPF_Formatado
    'PROP - Número do pedido/proposta.
    Print #pNumArqDados, pobjstcPed.NOSSO_NUMERO

    '3. Banco...
    'BSUD - Banco Sudameris.
    'BBOS - Bank Boston.
    'BBRA - Bradesco.
    'BITA - Banco Itaú.
    'BUNI - Unibanco.
    Select Case pobjstcPed.Cod_Bco_Cobr
    Case 347, 356    'Sudameris. 'Real
        Print #pNumArqDados, "X"
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
    Case 479    'Bank Boston.
        Print #pNumArqDados, " "
        Print #pNumArqDados, "X"
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
    Case 237    'Bradesco.
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, "X"
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
    Case 341    'Banco Itaú.
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, "X"
        Print #pNumArqDados, " "
    Case 409    'Unibanco.
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, "X"
    Case Else
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
    End Select
    'AGEN - Agência.
    Print #pNumArqDados, pobjstcPed.gAgenciaDig
    'CCOR - Conta corrente.
    Call mpForConta1(pobjstcPed.Cod_Bco_Cobr, pobjstcPed.Num_Cta_Bas_Cor, lAuxiliar)
    Print #pNumArqDados, lAuxiliar
    'LINH - Linha sobre as parcelas.
    If Val(pobjstcPed.objstcCtrlEmis.Num_Bloq & pobjstcPed.objstcCtrlEmis.DIG_BLOQUETO) <> 0 Then
        Print #pNumArqDados, Format(pobjstcPed.Qtd_Parc, "00") - 1 & " parcelas de R$ " & _
                             Format(pobjstcPed.val_DemParcelas, "Standard") & " cada."
    Else
        Print #pNumArqDados, "1ª Parcela de R$ " & Format(pobjstcPed.Val_PriParcela, "Standard") & " + " & _
                             Format(pobjstcPed.Qtd_Parc, "00") - 1 & " parcelas de R$ " & _
                             Format(pobjstcPed.val_DemParcelas, "Standard") & " cada."
    End If
    'DATA - Local e data.
    Print #pNumArqDados, lLocData

    mfCriDados_Debito = True
End Function

Private Sub mpForConta1(ByVal pCodBanco As Integer, ByVal pNumConta As Double, ByRef pSaida As String)
'Procedure: formata campo de conta corrente de numérico para caracter.

'Parâmetros de entrada...pBanco......Código do banco.
'                        pNumConta...Número da conta corrente.
'Parâmetro de saída......pSaida......Saída - número da conta corrente.

    Dim lAuxiliar As String    'Auxiliar.

    Select Case pCodBanco
    Case 237    'Bradesco.
        lAuxiliar = Format$(pNumConta, "00000000")
        pSaida = Mid$(lAuxiliar, 1, 7) & "-" & Mid$(lAuxiliar, 8, 1)
    Case 409    'Unibanco.
        lAuxiliar = Format$(pNumConta, "0000000")
        pSaida = Mid$(lAuxiliar, 1, 6) & "-" & Mid$(lAuxiliar, 7, 1)
    Case 341    'Itaú.
        lAuxiliar = Format$(pNumConta, "000000")
        pSaida = Mid$(lAuxiliar, 1, 5) & "-" & Mid$(lAuxiliar, 6, 1)
    Case 347    'Sudameris.
        lAuxiliar = Format$(pNumConta, "0000000000")
        pSaida = Mid$(lAuxiliar, 1, 9) & "-" & Mid$(lAuxiliar, 10, 1)
    Case 479    'Boston.
        lAuxiliar = Format$(pNumConta, "00000000")
        pSaida = Mid$(lAuxiliar, 1, 7) & "-" & Mid$(lAuxiliar, 8, 1)
    Case 356    'Real.
        lAuxiliar = Format$(pNumConta, "0000000000")
        pSaida = Mid$(lAuxiliar, 1, 9) & "-" & Mid$(lAuxiliar, 10, 1)
    Case Else       'Outros bancos.
        pSaida = Format$(pNumConta, "0000000000")
    End Select
End Sub

Private Sub mpDescricaoFranquia_V_0(pobjstcPedLoc As stcA46V708B, pobjstcPed As stcA46V702B, _
                                    pTextos() As String, pTamArray As Integer)
    Dim lBDP00Multi As ADODB.Connection
    Dim lColecaoTemp As Collection
    Dim lI As Integer
    Dim lTextoFranquia As Boolean
    Dim lMensagem As String
    Dim lobjstcCobert As stcA46V704B
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lrst As clsYasRecordSet
    Dim lTextos() As String
    Dim lCod_Ant As Integer
    Dim lCod_Linha As Integer
    Dim lCount As Integer
    Dim lCorretor As Double
    Dim lCobertura As Integer

    MsgBox "FUNCIONALIDADE ANTIGA - DESATIVADA"


    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, _
                              lBDP00Multi, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.item(1)
        lCorretor = lObjstcCorr.Cod_Corr
    Else
        lCorretor = 0
    End If

    'Ficha 901383
    If gM46V999.gCodProduto >= 962 Then

        'And gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then

        'Consulta por Corretor + Classificação
        lSelect = " Select * from TAB_DESC_FRANQUIA Where"
        lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
        lSelect = lSelect & " and (Cod_Produto = " & pobjstcPedLoc.Cod_Produto
        lSelect = lSelect & "  or (Cod_Produto = 0 and Cod_Regiao = " & pobjstcPedLoc.gCod_RegiaoForm & "))"
        lSelect = lSelect & "  and Compl_Clasf = '" & Trim(pobjstcPedLoc.Cod_Clasf) & Trim(pobjstcPedLoc.Cod_Compl_Clasf) & "'"
        lSelect = lSelect & "  and (Cod_Corretor = " & lCorretor & " or Cod_Corretor = 0)"
        lSelect = lSelect & " and Cod_Tarf_Ini <= " & pobjstcPed.Cod_Produto
        lSelect = lSelect & " and Cod_Tarf_Final >= " & pobjstcPed.Cod_Produto
        lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"

        If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            ReDim pTextos(6)
            Exit Sub
        End If

        lCobertura = 0
        Do While Not lrst.EOF
            For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
                If lobjstcCobert.Cod_Cobert = lrst("Cod_Cobert") Then
                    lCobertura = lCobertura + 1
                End If
            Next lobjstcCobert
            lrst.MoveNext
        Loop

        If lCobertura = 0 Then
            'Consulta por Corretor + todas as classificações
            lSelect = " Select * from TAB_DESC_FRANQUIA Where"
            lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and (Cod_Produto = " & pobjstcPedLoc.Cod_Produto
            lSelect = lSelect & "  or (Cod_Produto = 0 and Cod_Regiao = " & pobjstcPedLoc.gCod_RegiaoForm & "))"
            lSelect = lSelect & "  and (Compl_Clasf ='0')"
            lSelect = lSelect & "  and (Cod_Corretor = " & lCorretor & ")"
            lSelect = lSelect & " and Cod_Tarf_Ini <= " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and Cod_Tarf_Final >= " & pobjstcPed.Cod_Produto
            lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"

            If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
                Call gM46V999.gpGraLog(1, lMensagem)
                ReDim pTextos(6)
                Exit Sub
            End If

            lCobertura = 0
            Do While Not lrst.EOF
                For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
                    If lobjstcCobert.Cod_Cobert = lrst("Cod_Cobert") Then
                        lCobertura = lCobertura + 1
                    End If
                Next lobjstcCobert
                lrst.MoveNext
            Loop
        End If

        If lCobertura = 0 Then
            'Consulta por todos os corretores + todas as classificações
            lSelect = " Select * from TAB_DESC_FRANQUIA Where"
            lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and (Cod_Produto = " & pobjstcPedLoc.Cod_Produto
            lSelect = lSelect & "  or (Cod_Produto = 0 and Cod_Regiao = " & pobjstcPedLoc.gCod_RegiaoForm & "))"
            lSelect = lSelect & "  and (Compl_Clasf ='0')"
            lSelect = lSelect & "  and (Cod_Corretor = 0)"
            lSelect = lSelect & " and Cod_Tarf_Ini <= " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and Cod_Tarf_Final >= " & pobjstcPed.Cod_Produto
            lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"

            If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
                Call gM46V999.gpGraLog(1, lMensagem)
                ReDim pTextos(6)
                Exit Sub
            End If
        End If


        lI = 0
        lTextoFranquia = False
        ReDim lTextos(pTamArray)
        lCount = 1
        If lrst.EOF And lrst.BOF Then
            GoTo FinalTextos
        Else
            lrst.MoveFirst
        End If
        Do While Not lrst.EOF
            For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
                If lobjstcCobert.Cod_Cobert = lrst("Cod_Cobert") Then
                    Select Case lobjstcCobert.Cod_Cobert
                    Case gM46V999.gfBuscarDePara(39, 1), gM46V999.gfBuscarDePara(40, 1), gM46V999.gfBuscarDePara(41, 1), _
                        gM46V999.gfBuscarDePara(42, 1), gM46V999.gfBuscarDePara(43, 1), gM46V999.gfBuscarDePara(44, 1), _
                        gM46V999.gfBuscarDePara(47, 1), gM46V999.gfBuscarDePara(50, 1)
                        If pobjstcPed.Cod_Estip < 1020 Then
                            lTextoFranquia = True
                        End If
                    End Select
                    If lCod_Ant = lobjstcCobert.Cod_Cobert And _
                       lCod_Linha + 1 <> lrst("Qtde_Linha") Then
                        Exit For
                    End If
                    If lCod_Ant = lobjstcCobert.Cod_Cobert Then
                        lTextos(lI) = "   " & lrst("Descricao")
                    Else
                        lTextos(lI) = lCount & "- " & lrst("Descricao")
                        lCount = lCount + 1
                    End If
                    lCod_Ant = lrst("Cod_Cobert")
                    lCod_Linha = lrst("Qtde_Linha")
                    lI = lI + 1
                    If lI = pTamArray Then Exit Do
                    Exit For
                End If
            Next lobjstcCobert
            lrst.MoveNext
        Loop
        If lTextoFranquia = True Then
            lTextos(lI) = "Franquia: 7 (sete) dias"
            lI = lI + 1
        End If
        Call gM46V999.gpFecha2(lBDP00Multi)

    Else

        Dim lCorretorLiberado As Boolean

        lCorretorLiberado = False
        Select Case pobjstcPed.objstcCtrlEmis.Cod_Corr
        Case 49301, 8001, 8002, 8003, 8004, 8005, 8006, 8008, 8011, 8014, 8028, _
             3401, 345601, 486001, 489501, 646401, 5801, 648701, 648702, _
             648705, 648706, 648708, 358301, 219101, 15701, 493701, _
             6901, 686601, 403801, 8101
            lCorretorLiberado = True
        End Select

        lI = 0
        lTextoFranquia = False
        ReDim lTextos(pTamArray)
        lCount = 1

        If lCorretorLiberado = True Then
            For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
                If lobjstcCobert.Cod_Cobert = gM46V999.gfBuscarDePara(56, 1) Then
                    lTextos(lI) = "Franquia da cobertura de Pátios"
                    lI = lI + 1
                    lTextos(lI) = "R$ 750,00 - Moto até 400 cilindradas (por veículo)"
                    lI = lI + 1
                    lTextos(lI) = "R$ 1.500,00 - Auto/Moto acima de 400 cilindradas (nacionais por veículo)"
                    lI = lI + 1
                    lTextos(lI) = "R$ 3.000,00 - Auto/Moto acima de 400 cilindradas (importados por veículo)"
                    lI = lI + 1
                    Exit For
                End If
            Next
        End If

        If lI > 1 Then GoTo FinalTextos

        If gM46V999.gCodProduto >= 958 And Trim(pobjstcPedLoc.Cod_Clasf) = gM46V999.gfBuscarDePara("022", 2) And Trim(pobjstcPedLoc.Cod_Compl_Clasf) = gM46V999.gfBuscarDePara("022", 2, "13") Then
            For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
                If lobjstcCobert.Cod_Cobert = gM46V999.gfBuscarDePara(56, 1) And Not lCorretorLiberado Then
                    lTextos(lI) = "Fica entendido e acordado que, ao contrário do que possa constar no tópico 'Franquia' do frontispício desta apólice, a franquia"
                    lI = lI + 1
                    lTextos(lI) = "da cobertura 056 - Pátio deve ser lida e considerada conforme abaixo:"
                    lI = lI + 1
                    lTextos(lI) = ""
                    lI = lI + 1
                    lTextos(lI) = "Por evento, serão deduzidas das indenizações a seguinte participação do segurado:"
                    lI = lI + 1
                    lTextos(lI) = "R$ 1.500,00 - veículos nacionais (por veículo)"
                    lI = lI + 1
                    lTextos(lI) = "R$ 3.000,00 - veículos importados (por veículo)"
                    lI = lI + 1
                    Exit For
                End If
                If lobjstcCobert.Cod_Cobert = gM46V999.gfBuscarDePara(57, 1) Then
                    lTextos(lI) = "Fica entendido e acordado que, ao contrário do que possa constar no tópico 'Franquia' do frontispício desta apólice, a franquia"
                    lI = lI + 1
                    lTextos(lI) = "da cobertura 057 - Resp Civil Concessionárias deve ser lida e considerada conforme abaixo:"
                    lI = lI + 1
                    lTextos(lI) = ""
                    lI = lI + 1
                    lTextos(lI) = "Por evento, serão deduzidas das indenizações a seguinte participação do segurado:"
                    lI = lI + 1
                    lTextos(lI) = "R$ 2.000,00 - por veículo"
                    lI = lI + 1
                    Exit For
                End If
            Next
        End If

        If lI > 1 Then GoTo FinalTextos

        If gM46V999.gCodProduto >= 958 Then
            For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
                If lobjstcCobert.Cod_Cobert = gM46V999.gfBuscarDePara(60, 1) Then
                    lTextos(lI) = "Fica entendido e acordado que, ao contrário do que possa constar no tópico 'Franquia' do frontispício desta apólice, a franquia"
                    lI = lI + 1
                    lTextos(lI) = "da cobertura 060 - Resp Civil Garagista (Compreensiva) deve ser lida e considerada conforme abaixo:"
                    lI = lI + 1
                    lTextos(lI) = ""
                    lI = lI + 1
                    lTextos(lI) = "Por evento, serão deduzidas das indenizações a seguinte participação do segurado:"
                    lI = lI + 1
                    lTextos(lI) = "R$ 1.500,00 - veículos nacionais (por veículo)"
                    lI = lI + 1
                    lTextos(lI) = "R$ 3.000,00 - veículos importados (por veículo)"
                    lI = lI + 1
                    Exit For
                End If
            Next
        End If

    End If

FinalTextos:
    pTextos = lTextos

End Sub

Private Function gCoefComissao(pComissao As Single) As Single
    gCoefComissao = (1 - gM46O001.gProComBasica / 100) / (1 - pComissao / 100)
End Function
