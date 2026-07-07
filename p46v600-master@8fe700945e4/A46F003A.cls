VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46F003A"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ======================================================================================================================================================
' Class Name: clsA46F003 --> Classe do Form frmT46V003A
' Data/Hora: 01/06/2011
' Sistema: SyasMulti Web
' Simony Sales Andrade
' ======================================================================================================================================================


Public Enum Enum_TipDocumento
    Enum_TipDocumentoCot = 1        'Cotações e propostas.
    Enum_TipDocumentoApol = 2       'Apolice Vigentes
    Enum_TipDocumentoProt = 3       'Protocolos
End Enum

Public Enum Enum_SitDocumento
    Enum_SitDocumentoTodas = 0         'Todas
    Enum_SitDocumentoNaoTran = 1       'Não Transmitidas
    Enum_SitDocumentoTran = 2          'Transmitidas
    Enum_SitDocumentoInspec = 3        'Inspecao não solicitaas
End Enum

Private txtCorretor As String
Private txtProdutor As String
Private txtIniVigencia As String
Private txtTerVigencia As String
Private lTipDocumento As Enum_TipDocumento
Private txtNumPedido As String
Private txtSegurado As String
Private lCodRamo As String
Private lMensagem As String
Private lSitDocumento As Enum_SitDocumento

Private lRetRecordSet    As New ADODB.Recordset

Public UltimoErro As String
Public Function gfProtocoloTransmissao(pNossoNumero As String) As String    'TESTE OK
    Dim lDir            As String           'Utilizado com a função DIR.
    Dim lMensagem       As String           'Mensagem.
    Dim lNosNumero    As String           'Nosso numero
    Dim lNumArquivo     As Integer          'Utilizado com FreeFile.
    Dim lPriProposta    As Boolean          'Primeira proposta.
    Dim lRegistro       As String           'Registro.
    Dim lobjTransm      As clsA46V718A      'Objeto da classe clsA46V718A (TAB_PED_TRANSM)
    Dim lobjstcTransm   As stcA46V718B      'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM)
    Dim lCount          As Long             'Contador.
    Dim clsA99V001      As clsA99V001
    Dim lDigito         As String
    Dim mQuaLinhas      As Long
    Dim lProtocolo      As String
    Screen.MousePointer = vbHourglass
    On Error GoTo ControleErro
    '1. Centraliza form.

    '4. Obtém registros de protocolo referente ao nosso número.
    Set lobjTransm = New clsA46V718A
    Set lobjstcTransm = New stcA46V718B
    lobjstcTransm.Nosso_Numero = pNossoNumero
    
    If Not lobjTransm.mfCarregarColecao(lobjstcTransm) Then
        UltimoErro = lobjTransm.UltimoErro
        Exit Function
    End If
    If lobjTransm.Colecao.Count = 0 Then
        UltimoErro = "FNL0014 - Registro de protocolo de transmissão não localizado (pedido " & _
                         gNumPedido & " - nosso número " & pNossoNumero & ")."
        Set lobjTransm = Nothing
        Set lobjstcTransm = Nothing
        Exit Function
    End If
    
    'c) Obtém dados da proposta transmitida (detalhes - erro).
    '*************************Ver este ponto

    '5. Processa registro de protocolo.
    lPriProposta = True
    '6. Inicia processo.
    mQuaLinhas = 0
    Call mpPosLinha(lProtocolo, mQuaLinhas, "YASUDA - PROTOCOLO DE TRANSMISSÃO - " & Format$(Now, "dd/mm/yyyy hh:nn:ss") & ".")
    Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
    Call mpPosLinha(lProtocolo, mQuaLinhas, "Protocolo de transmissão - SYAS Multiseguros.")
    Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
    For Each lobjstcTransm In lobjTransm.Colecao
        'a) Exibe dados
        'Call mpPosLinha(lProtocolo, mQuaLinhas,"Transmissão da Yasuda: " & lobjstcTransm.Nom_Arq_Retorno & ".")
        If lNosNumero <> lobjstcTransm.Nosso_Numero Then
            lNosNumero = lobjstcTransm.Nosso_Numero
            If gfPreenchido(lobjstcTransm.Nosso_Numero) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Proposta recebida....." & lobjstcTransm.Nosso_Numero)
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Proposta recebida.....não informado")
            End If
            If gfPreenchido(lobjstcTransm.Num_Ped) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Pedido................" & lobjstcTransm.Num_Ped)
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Pedido................não informado")
            End If
            If gfPreenchido(lobjstcTransm.Cod_Ramo) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Ramo.................." & lobjstcTransm.Cod_Ramo)
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Ramo..................não informado")
            End If
            If gfPreenchido(lobjstcTransm.Nom_Segurado) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Segurado.............." & Trim$(lobjstcTransm.Nom_Segurado))
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Segurado..............não informado.")
            End If
            If gfPreenchido(lobjstcTransm.Dat_Ini_Vig) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Início de vigência...." & gfForData(2, lobjstcTransm.Dat_Ini_Vig))
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Início de vigência....não informado")
            End If
            If gfPreenchido(lobjstcTransm.Dat_Ter_Vig) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Término de vigência..." & gfForData(2, lobjstcTransm.Dat_Ter_Vig))
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Término de vigência...não informado")
            End If
            If gfPreenchido(lobjstcTransm.Nom_Arq_Envio) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Arquivo enviado......." & lobjstcTransm.Nom_Arq_Envio)
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Arquivo enviado.......não informado")
            End If
            If gfPreenchido(lobjstcTransm.Dat_Hor_Envio) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Data/Hora envio......." & gfForData(4, lobjstcTransm.Dat_Hor_Envio) & " horas.")
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Data/Hora envio.......não informado")
            End If
            If gfPreenchido(lobjstcTransm.Nom_Arq_Retorno) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Arquivo de retorno...." & lobjstcTransm.Nom_Arq_Retorno)
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Arquivo de retorno....não informado")
            End If
            If gfPreenchido(lobjstcTransm.Dat_Hor_Retorno) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Data/Hora retorno....." & gfForData(4, lobjstcTransm.Dat_Hor_Envio) & " horas.")
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Data/Hora retorno.....não informado")
            End If
            If Val(lobjstcTransm.Num_PI) <> 0 Then
                Set clsA99V001 = New clsA99V001
                If clsA99V001.DIG_YASUDA(lobjstcTransm.Num_PI, lDigito, lMensagem) = False Then
                    Exit Function
                End If
                Set clsA99V001 = Nothing
                If lDigito = "A" Then lDigito = "0"
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Protocolo Interno....." & lobjstcTransm.Num_PI & "-" & lDigito)
                Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Os anexos, quando necessários, poderão ser enviados das seguintes formas:")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "1) VIA FAX: Através  do  número  (0xx11) 3305-5664,  informando  o  número  do ")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "protocolo interno " & lobjstcTransm.Num_PI & lDigito & " no atendimento eletrônico ou;")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "2) VIA YasudaNet, 'ÁREA DO CORRETOR': Carregar o(s) documento(s) digitalizado(s)")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "através  dos  links  Seguro  Demais Ramos  -->  Emissão.  Filtrar pelo Número do")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Protocolo (PI).")
                Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
            End If
            Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
            Select Case lobjstcTransm.Ind_Proposta
                Case "0"
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "STATUS................Proposta não transmitida.")
                Case 1
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "STATUS...Proposta eletrônica recebida pela Yasuda em " & _
                                    gfForData(4, lobjstcTransm.Dat_Hor_Retorno) & " horas.")
                    
                Case 2
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "STATUS................Proposta apresentou erro.")
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "É necessário retransmitir esta proposta.")
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "Descrição do erro....." & lobjstcTransm.Tip_Emissao) 'Contém observação do erro
                Case Else
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "STATUS................Proposta apresentou erro.")
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "É necessário retransmitir esta proposta.")
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "Descrição do erro.....Indicador de proposta desconhecido - comunicar Central de Atendimento.")
            End Select
            Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Itens da proposta:")
         End If
        '***==== DADOS ITENS
        Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
        If gfPreenchido(lobjstcTransm.Num_Item) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Item.................." & lobjstcTransm.Num_Item)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Item..................não informado")
        End If
        Call mpPosLinha(lProtocolo, mQuaLinhas, "Logradouro............" & lobjstcTransm.Cod_Tip_Logr_Risco & _
                            " " & lobjstcTransm.Nom_Logr_Risco & _
                            ", " & lobjstcTransm.Num_Risco)
        
        If gfPreenchido(lobjstcTransm.Dsc_Compl_Risco) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Complemento..........." & lobjstcTransm.Dsc_Compl_Risco)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Complemento...........não informado")
        End If
        If gfPreenchido(lobjstcTransm.Num_Cep_Risco) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "CEP..................." & lobjstcTransm.Num_Cep_Risco)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "CEP...................não informado")
        End If
        If gfPreenchido(lobjstcTransm.Nom_Cid_Risco) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Cidade................" & lobjstcTransm.Nom_Cid_Risco)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Cidade................não informada")
        End If
        If gfPreenchido(lobjstcTransm.Sig_UF_Risco) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "UF...................." & lobjstcTransm.Sig_UF_Risco)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "UF....................não informada")
        End If
    Next lobjstcTransm
    Call mpPosLinha(lProtocolo, mQuaLinhas, String(80, "="))
    Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
    Call mpPosLinha(lProtocolo, mQuaLinhas, "FIM " & String(76, "="))
    
    gfProtocoloTransmissao = lProtocolo
    
    Set lobjstcTransm = Nothing
    Set lobjTransm = Nothing

    Exit Function
ControleErro:
    UltimoErro = UltimoErro & "gfProtocoloTransmissao ERRO - " & Err & " " & Err.Description
End Function

Private Sub mpPosLinha(ByRef rtbTexto As String, ByRef pQuaLinhas As Long, ByVal pMensagem As String)
    'Procedure: posiciona linha de mensagem.

    Dim lAuxiliar   As String   'Auxiliar.

    lAuxiliar = pMensagem
    If Len(lAuxiliar) < 81 Then
        If gfPreenchido(rtbTexto) Then
            rtbTexto = rtbTexto & vbLf & lAuxiliar
        Else
            rtbTexto = lAuxiliar
        End If
        pQuaLinhas = pQuaLinhas + 1
    Else
        If gfPreenchido(rtbTexto) Then
            rtbTexto = rtbTexto & vbLf & Left$(lAuxiliar, 77) & "..."
        Else
            rtbTexto = Left$(lAuxiliar, 77) & "..."
        End If
        pQuaLinhas = pQuaLinhas + 1
        lAuxiliar = Mid$(lAuxiliar, 78)
        Do
            If Len(lAuxiliar) < 77 Then
                rtbTexto = rtbTexto & vbLf & "..." & lAuxiliar
                pQuaLinhas = pQuaLinhas + 1
                Exit Do
            End If
            rtbTexto = rtbTexto & vbLf & "..." & Left$(lAuxiliar, 74) & "..."
            pQuaLinhas = pQuaLinhas + 1
            lAuxiliar = Mid$(lAuxiliar, 5)
        Loop
        rtbTexto = rtbTexto & vbLf & " "
        pQuaLinhas = pQuaLinhas + 1
    End If

End Sub

Public Sub mpPesquisaDocumento(pCorretor As String, pProdutor As String, pIniVigencia As String, _
                               pTerVigencia As String, pOptDocumento As Integer, pNumPedido As String, _
                               pSegurado As String, pCodRamo As Integer, pSitDocumento As Integer, ByRef pRstRetorno, ByRef pMensagem As String)
    Dim lTemNumPedido   As Boolean  'Indica que tem número do pedido para pesquisa.

    On Error GoTo ControleErro
        
    txtCorretor = IIf(pCorretor = "0", "", pCorretor)
    txtProdutor = IIf(pProdutor = "0", "", pProdutor)
    txtIniVigencia = IIf(pIniVigencia = "0", "", pIniVigencia)
    txtTerVigencia = IIf(pTerVigencia = "0", "", pTerVigencia)
    lTipDocumento = pOptDocumento
    txtNumPedido = IIf(pNumPedido = "0", "", pNumPedido)
    txtSegurado = pSegurado
    lCodRamo = pCodRamo
    lSitDocumento = pSitDocumento
    
    '2. Consiste critérios de pesquisa.
    'a) Tipo de documento.
    If lTipDocumento <> Enum_TipDocumentoCot And lTipDocumento <> Enum_TipDocumentoApol And lTipDocumento <> Enum_TipDocumentoProt Then
        lMensagem = "FNP0013 - Tipo de documento não selecionado."
        Call gpGraLog(0, lMensagem)
        GoTo Fim
    End If
    
    'b) Número do pedido preenchido.
    lTemNumPedido = False
    If gfPreenchido(txtNumPedido) Then
        If Not IsNumeric(txtNumPedido) Then
            lMensagem = "FGE0080 - Número do pedido não numérico."
            Call gpGraLog(0, lMensagem)
            GoTo Fim
        End If
        lTemNumPedido = True
    End If

        
    'c) Ramo
    If gTipUsuario = e_Tipo_Usu_Yasuda Then
        If gfPreenchido(txtSegurado) = False And gfPreenchido(txtNumPedido) = False Then
            If lCodRamo = 0 Then
                lMensagem = "FRA0001 - Ramo não selecionado."
                Call gpGraLog(0, lMensagem)
                GoTo Fim
            End If
        End If
    End If

    'd) Número do pedido não preenchido.
    If lTemNumPedido = False Then
        '=== Vigência.
        If gfPreenchido(txtIniVigencia) Then
            If gfVerData(2, txtIniVigencia) = False Then
                lMensagem = "FGE0018 - Vigência inválida (data de início de vigência inválida)."
                Call gpGraLog(0, lMensagem)
                GoTo Fim
            End If
        End If
        If gfPreenchido(txtTerVigencia) Then
            If gfVerData(2, txtTerVigencia) = False Then
                lMensagem = "FGE0081 - Vigência inválida (data de término de vigência inválida)."
                Call gpGraLog(0, lMensagem)
                GoTo Fim
            End If
        End If
        If gfPreenchido(txtIniVigencia) And gfPreenchido(txtTerVigencia) Then
            If gfForData(1, txtIniVigencia) > gfForData(1, txtTerVigencia) Then
                lMensagem = "FGE0082 - Vigência inválida (data de início de vigência " & _
                                 txtIniVigencia & ") maior que a data de término de vigência (" & _
                                 txtTerVigencia & ")."
                Call gpGraLog(0, lMensagem)
                GoTo Fim
            End If
        End If
    End If

    '3. Obtém dados.
    lMensagem = ""
    Set lRetRecordSet = New ADODB.Recordset
    Select Case lTipDocumento
        Case 1
            Call mpPesCotPropostas
        Case 2
            Call mpPesApolices
        Case 3
            Call mpPesProTransmissao
        Case Else
            lMensagem = "FNP0013 - Tipo de documento não selecionado."
            Call gpGraLog(0, lMensagem)
            GoTo Fim
    End Select
    If lMensagem = "" Then
        Set pRstRetorno = New ADODB.Recordset
        If lRetRecordSet.State = 1 Then
            lRetRecordSet.MoveFirst
            Set pRstRetorno = lRetRecordSet
        End If
    End If
    
Fim:

    pMensagem = lMensagem
    UltimoErro = lMensagem
    Exit Sub
ControleErro:
    lMensagem = lMensagem & Err & " " & Err.Description
    UltimoErro = lMensagem
    Call gpGraLog(0, "mpPesquisaDocumento ERRO - " & lMensagem)
    GoTo Fim
End Sub
Private Sub mpPesProTransmissao()
    'Procedure: pesquisa protocolo de transmissão.

    Dim lData           As String           'Campo auxiliar: data.
    Dim lDir            As String           'Utilizado com a função DIR.
    Dim lobjTransm      As clsA46V718A      'Objeto da classe clsA46V718A (TAB_PED_TRANSM)
    Dim lobjstcTransm   As stcA46V718B      'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM)

    On Error GoTo ControleErro
    
    '1. Instancio os objetos
    Set lobjTransm = New clsA46V718A
    Set lobjstcTransm = New stcA46V718B
    

    '2. Carrego os parâmetro de pesquisa.
    'Obs.: Se número do pedido está preenchido, utiliza apenas esta informação no SELECT.
    If gfPreenchido(txtNumPedido) Then
        lobjstcTransm.Num_Ped = Format$(txtNumPedido, "000000")
    Else
        If gfPreenchido(txtSegurado) Then
            lobjstcTransm.Nom_Segurado = gfForAspas(UCase(txtSegurado))
        End If
        'b) Vigência.
        If gfPreenchido(txtIniVigencia) Then
            lData = gfForData(1, txtIniVigencia)
            lobjstcTransm.Dat_Ini_Vig = lData
        End If
        If gfPreenchido(txtTerVigencia) Then
            lData = gfForData(1, txtTerVigencia)
            lobjstcTransm.Dat_Ter_Vig = lData
        End If
    End If

    '3. Obtém registros de protocolo.
    If Not lobjTransm.mfCarregarColecao(lobjstcTransm) Then
        UltimoErro = lobjTransm.UltimoErro
        Call gpGraLog(0, UltimoErro)
        Exit Sub
    End If

    '4. Se não encontrou registro, informa usuário.
    If lobjTransm.Colecao.Count = 0 Then
        lDir = Dir("*.Y01")
        If gfPreenchido(lDir) Then
            'Call mnuExiProAntigos_Click
            Exit Sub
        Else
            Call mpMenNaoEncontrado
            Exit Sub
        End If
    End If

    '5. Monta grade de resultados.
    
    lRetRecordSet.Fields.Append "Tipo", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Pedido", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Segurado", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Inicio_Vigencia", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Termino_Vigencia", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Tipo_logradouro", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Logradouro", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Nmmero_logradouro", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Complemento", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Bairro", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Cidade", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "UF", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "CEP", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Nosso_numero", adBSTR, , adFldIsNullable

    '6. Coloca registros na grade.
    
     For Each lobjstcTransm In lobjTransm.Colecao
        'a) Acrescenta linha na grade.
        lRetRecordSet.AddNew
        '===1) Tipo do documento.
        lRetRecordSet(0) = "Protocolo"
        '===2) Número do pedido.
        lRetRecordSet(1) = Format$(lobjstcTransm.Num_Ped, "000000")
        '===3) Nome do segurado.
        If gfPreenchido(lobjstcTransm.Nom_Segurado) Then
             lRetRecordSet(2) = lobjstcTransm.Nom_Segurado
         Else
             lRetRecordSet(2) = ""
         End If
         '===4) Início de vigência.
         If gfPreenchido(lobjstcTransm.Dat_Ini_Vig) Then
             lRetRecordSet(3) = gfForData(2, lobjstcTransm.Dat_Ini_Vig)
         Else
             lRetRecordSet(3) = ""
         End If
         '===5) Término de vigência.
         If gfPreenchido(lobjstcTransm.Dat_Ter_Vig) Then
             lRetRecordSet(4) = gfForData(2, lobjstcTransm.Dat_Ter_Vig)
         Else
             lRetRecordSet(4) = ""
         End If
         '===6) Tipo do logradouro
         If gfPreenchido(lobjstcTransm.Cod_Tip_Logr_Risco) Then
             lRetRecordSet(5) = lobjstcTransm.Cod_Tip_Logr_Risco
         Else
             lRetRecordSet(5) = ""
         End If
         '===7) logradouro.
         If gfPreenchido(lobjstcTransm.Nom_Logr_Risco) Then
             lRetRecordSet(6) = lobjstcTransm.Nom_Logr_Risco
         Else
             lRetRecordSet(6) = ""
         End If
         '===8) Número logradouro.
         If gfPreenchido(lobjstcTransm.Num_Risco) Then
             lRetRecordSet(7) = lobjstcTransm.Num_Risco
         Else
             lRetRecordSet(7) = ""
         End If
         '===9) Complemento.
         If gfPreenchido(lobjstcTransm.Dsc_Compl_Risco) Then
             lRetRecordSet(8) = lobjstcTransm.Dsc_Compl_Risco
         Else
             lRetRecordSet(8) = ""
         End If
         '===10) Complemento.
         If gfPreenchido(lobjstcTransm.Nom_Bairro_Risco) Then
             lRetRecordSet(9) = lobjstcTransm.Nom_Bairro_Risco
         Else
             lRetRecordSet(9) = ""
         End If
         '===11) Cidade.
         If gfPreenchido(lobjstcTransm.Nom_Cid_Risco) Then
             lRetRecordSet(10) = lobjstcTransm.Nom_Cid_Risco
         Else
             lRetRecordSet(10) = ""
         End If
         '===12) UF.
         If gfPreenchido(lobjstcTransm.Sig_UF_Risco) Then
             lRetRecordSet(11) = lobjstcTransm.Sig_UF_Risco
         Else
             lRetRecordSet(11) = ""
         End If
         '===13) CEP.
         If gfPreenchido(lobjstcTransm.Num_Cep_Risco) Then
             lRetRecordSet(12) = lobjstcTransm.Num_Cep_Risco
         Else
             lRetRecordSet(12) = ""
         End If
         '===14) Nosso número.
         lRetRecordSet(13) = lobjstcTransm.Nosso_Numero
         lRetRecordSet.Update
    Next lobjstcTransm


    '7. Fecha tabelas e bases de dados.
    Set lobjstcTransm = Nothing
    Set lobjTransm = Nothing
    Exit Sub
ControleErro:
    lMensagem = Err & " " & Err.Description
    Call gpGraLog(1, lMensagem, vbCritical)
    Call gpGraLog(0, "mpPesProTransmissao ERRO - " & Err & " " & Err.Description)
End Sub



Private Sub mpMenNaoEncontrado()

    lMensagem = "FNL0011 - Não localizado registro que satisfaça os critérios selecionados." & vbLf
    If gfPreenchido(txtNumPedido) Then
        lMensagem = lMensagem & vbLf & "Número do pedido = " & txtNumPedido & "."
    Else
        If lTipDocumento = Enum_TipDocumentoCot Then
            lMensagem = lMensagem & vbLf & "Tipo de documento = cotações e propostas"
            If lSitDocumento = Enum_SitDocumentoTodas Then
                lMensagem = lMensagem & "."
            Else
                If lSitDocumento = Enum_SitDocumentoNaoTran Then
                    lMensagem = lMensagem & " (apenas propostas não transmitidas)."
                Else
                    If lSitDocumento = Enum_SitDocumentoTran Then
                       lMensagem = lMensagem & " (apenas propostas transmitidas)."
                    Else
                       lMensagem = lMensagem & " (apenas cotações sem solicitação de agendamento)."
                    End If
                End If
            End If
        Else
            If lTipDocumento = Enum_TipDocumentoApol Then
                lMensagem = lMensagem & vbLf & "Tipo de documento = apólices vigentes."
            Else
                lMensagem = lMensagem & vbLf & "Tipo de documento = protocolo de transmissão."
            End If
        End If
        If gfPreenchido(txtSegurado) Then
            lMensagem = lMensagem & vbLf & "Nome do segurado que contenha """ & txtSegurado & """."
        End If
        If gfPreenchido(txtProdutor) Then
            lMensagem = lMensagem & vbLf & "Nome do produtor que contenha """ & txtProdutor & """."
        End If
        If gfPreenchido(txtCorretor) Then
            lMensagem = lMensagem & vbLf & "Corretor = " & txtCorretor & "."
        End If
        If gfPreenchido(txtIniVigencia) And gfPreenchido(txtTerVigencia) Then
            lMensagem = lMensagem & vbLf & "Vigência: de " & txtIniVigencia & " a " & _
                        txtTerVigencia & "."
        Else
            If gfPreenchido(txtIniVigencia) Then
                lMensagem = lMensagem & vbLf & "Vigência em " & txtIniVigencia & "."
            Else
                If gfPreenchido(txtTerVigencia) Then
                    lMensagem = lMensagem & vbLf & "Vigência em " & txtTerVigencia & "."
                End If
            End If
        End If
        
        If lCodRamo <> e_CodRamo_NaoDefinido Then
            lMensagem = lMensagem & vbLf & "Ramo =" & lCodRamo & "."
        End If
    End If
    Call gpGraLog(0, lMensagem)
End Sub


Private Sub mpPesApolices()
    'Procedure: pesquisa cotações e propostas.
    Dim lListaNossoNumero   As String           'Lista de nosso_numero.
    Dim lCount              As Double           'Contador.
    Dim lrst                As ADODB.Recordset  'Registro de apólice.
    Dim bdP00APOL           As ADODB.Connection 'Arquivo de apólice (Conexão).
    Dim lCpfCnpj            As String           'Retorno da função de formação de cpf e cnpj.
    Dim lCod_Corr           As String           'Codigo do corretor.
    Dim lRegistro           As String           'Registro.
    Dim lNomeTab     As String

    On Error GoTo ControleErro
    
    If gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
    Else
        lNomeTab = "GED.DBO.TAB_PED"
    End If
    
    '2. Abre base de dados
    If gfAbrBasDados("P00APOL", gTipP00APOL, gEndP00APOL, bdP00APOL, lMensagem) = False Then
        Call gpGraLog(0, lMensagem)
        Exit Sub
    End If
    
    '.3 Montagem do comando Select.
    lListaNossoNumero = ""
    If gfPreenchido(txtCorretor) Then
        gSelect = "SELECT * FROM " & lNomeTab & "_Corr"
        gSelect = gSelect & " Where Cod_Corr= " & txtCorretor
    
        If Not gfObtRegistro(bdP00APOL, gSelect, lrst, lMensagem) Then
            Call gpGraLog(1, lMensagem)
            Exit Sub
        End If
        lListaNossoNumero = ""
        Do While Not lrst.EOF
            If gfPreenchido(lListaNossoNumero) Then lListaNossoNumero = lListaNossoNumero & ","
            lListaNossoNumero = lListaNossoNumero & "'" & lrst("Nosso_Numero") & "'"
            lrst.MoveNext
        Loop
    End If
    
    Call gfVerCorrPermitido(gCod_Corretor, 0, lRegistro)
    'Busco os dados da apólice
    
    'Ficha 1000782 - Não permitir Renovação LMI para Corretores
    gSelect = "SELECT distinct " & lNomeTab & ".* , " & lNomeTab & "_LOC.NOM_LOGR_RISCO "
    gSelect = gSelect & " FROM (" & lNomeTab & " inner join " & lNomeTab & "_Corr"
    gSelect = gSelect & " on " & lNomeTab & ".Num_APol = " & lNomeTab & "_Corr.NUm_Apol and"
    'Ficha 1000782 - Não permitir Renovação LMI para Corretores
    gSelect = gSelect & " " & lNomeTab & ".Num_ENDO = " & lNomeTab & "_Corr.NUm_ENDO) "
    gSelect = gSelect & " inner join " & lNomeTab & "_LOC "
    gSelect = gSelect & " on " & lNomeTab & "_LOC.NUM_APOL = " & lNomeTab & ".NUM_APOL "
    If Not gBaseLocal Then
        gSelect = gSelect & " AND " & lNomeTab & "_LOC.NUM_ENDO = " & lNomeTab & ".NUM_ENDO "
    End If
    
    If Not gBaseLocal = True Then
        gSelect = gSelect & " and " & lNomeTab & ".NUM_PI IN (SELECT NUM_PI "
        gSelect = gSelect & " From [tab_ctrl_emis]"
        gSelect = gSelect & " WHERE  [cod_iden_stat_acao] = 90 ) "
    End If
            
    'Nome do segurado.
    If gfPreenchido(txtSegurado) Then
        gSelect = gfWhereAnd(gSelect) & " Nom_Segurado like '%" & UCase(txtSegurado) & "%'"
    End If
            
    'Início da vigência.
    If gfPreenchido(txtIniVigencia) Then
        gSelect = gfWhereAnd(gSelect) & " Dat_Ini_Vig >= " & gfForData(1, txtIniVigencia)
    End If
    
    'Término da vigência.
    If gfPreenchido(txtTerVigencia) Then
        gSelect = gfWhereAnd(gSelect) & " Dat_Ter_Vig <= " & gfForData(1, txtTerVigencia)
    End If

    'Produtor.
    If gfPreenchido(txtProdutor) Then
        gSelect = gfWhereAnd(gSelect) & " Cod_Prod = " & txtProdutor
    End If
    
    If lCodRamo <> e_CodRamo_NaoDefinido Then
        gSelect = gfWhereAnd(gSelect) & " " & lNomeTab & ".Cod_Ramo = " & lCodRamo
    End If

    If gfPreenchido(lListaNossoNumero) Then
        gSelect = gfWhereAnd(gSelect) & lNomeTab & ".Nosso_Numero in ( " & lListaNossoNumero & ")"
    End If
    
    'Ficha 1000782 - Não permitir Renovação LMI para Corretores
    If gTipUsuario = e_Tipo_Usu_Corretor Then
        gSelect = gfWhereAnd(gSelect) & lNomeTab & "_LOC.NOM_LOGR_RISCO not like '%CONFORME ESPEC%'"
    End If
    
    gSelect = gfWhereAnd(gSelect) & " Dat_Ter_Vig >= " & gfForData(1, Date - 30)
    gSelect = gfWhereAnd(gSelect) & " cod_dizer not in (4,10) "
    
    gSelect = gSelect & " ORDER BY " & lNomeTab & ".Num_Ped"

    '4. obtenho o registro.
    If Not gfObtRegistro(bdP00APOL, gSelect, lrst, lMensagem) Then
        Call gpGraLog(1, lMensagem)
        Exit Sub
    End If
    
    '5. Se não encontrou registro, informa usuário.
    If lrst.EOF Then
        Call mpMenNaoEncontrado
        Exit Sub
    End If
    
    '6. Monta grade de resultados.
    
    lRetRecordSet.Fields.Append "Excluir", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Ramo", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Apólice", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Segurado", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "CNPJ/CPF", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Início Vigência", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Término Vigência", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Nosso número", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Situação do documento", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Produto/Modalidade", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "CodRamo", adBSTR, , adFldIsNullable
    
    '7. Coloca registros na grade.
    Do While Not lrst.EOF
         'b) Coluna 0 - Excluir - checkBox para marcar itens a serem excluídos.
         If lrst("Cod_Ramo") = e_CodRamo_Condominio Then
             lRetRecordSet(1) = "Condomínio"
         ElseIf lrst("Cod_Ramo") = e_CodRamo_Empresarial Then
             lRetRecordSet(1) = "Empresarial"
         ElseIf lrst("Cod_Ramo") = e_CodRamo_Residencial Then
             lRetRecordSet(1) = "Residencial"
         Else
             lRetRecordSet(1) = "Não definido"
         End If
         lRetRecordSet(2) = lrst("Num_Apol")
         lRetRecordSet(3) = lrst("Nom_Segurado")
         Select Case Len(CStr(lrst("Num_Cgc_Cpf")))
             Case Is > 11
                 Call gpForCNPJ(1, Format$(lrst("Num_Cgc_Cpf"), "00000000000000"), lCpfCnpj)
             Case Else
                 Call gpForCPF(1, Format$(lrst("Num_Cgc_Cpf"), "00000000000"), lCpfCnpj)
         End Select
         lRetRecordSet(4) = lCpfCnpj
         lRetRecordSet(5) = gfForData(2, lrst("Dat_Ini_Vig"))
         lRetRecordSet(6) = gfForData(2, lrst("Dat_Ter_Vig"))
         lRetRecordSet(7) = 0
         lRetRecordSet(8) = 0
         lRetRecordSet(9) = lrst("Cod_Estip")
         lRetRecordSet(10) = lrst("Cod_Ramo")
         lRetRecordSet.Update
         lCount = lCount + 1
         lrst.MoveNext
    Loop

    Exit Sub
ControleErro:
    lMensagem = Err & " " & Err.Description
    Call gpGraLog(1, lMensagem, vbCritical)
    Call gpGraLog(0, "mpPesApolice ERRO - " & Err & " " & Err.Description)
    Exit Sub
Resume Next
End Sub


Private Sub mpPesCotPropostas()
    'Procedure: pesquisa cotações e propostas.
    
    Dim lListaNossoNumero   As String           'Lista de nosso_numero.
    Dim lCount              As Double           'Contador.
    Dim lI                  As Double           'Auxiliar.
    Dim lColuna             As Integer          'Coluna do grid.
    Dim lObjPed             As clsA46V702A      'Objeto da classe clsA46V702A (TAB_PED).
    Dim lObjstcPed          As stcA46V702B      'Objeto da classe de estrutura stcA46V702B (TAB_PED).
    Dim lObjPedCorr         As clsA46V705A      'Objeto da classe clsA46V705A (TAB_PED_CORR).
    Dim lObjstcPedCorr      As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).
    Dim lobjstcPedLoc       As stcA46V708B
    
    On Error GoTo ControleErro
   
    '2. Busco os registro com o corretor informado  (se informado).
    lListaNossoNumero = ""
    If gfPreenchido(txtCorretor) Then
        Set lObjPedCorr = New clsA46V705A
        Set lObjstcPedCorr = New stcA46V705B
        lObjstcPedCorr.Cod_Corr = txtCorretor
        If Not lObjPedCorr.mfCarregarColecao(lObjstcPedCorr, lListaNossoNumero) Then
            lMensagem = lObjPedCorr.UltimoErro
            Call gpGraLog(0, lObjPedCorr.UltimoErro)
            Exit Sub
        End If
    ElseIf gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        Set lObjPedCorr = New clsA46V705A
        Set lObjstcPedCorr = New stcA46V705B
        lObjstcPedCorr.Cod_Corr = gCod_Corretor
        If Not lObjPedCorr.mfCarregarColecao(lObjstcPedCorr, lListaNossoNumero) Then
            lMensagem = lObjPedCorr.UltimoErro
            Call gpGraLog(0, lObjPedCorr.UltimoErro)
            Exit Sub
        End If
    End If
    
    '3. Instancio os objetos
    Set lObjPed = New clsA46V702A
    Set lObjstcPed = New stcA46V702B
        
    If lCodRamo <> e_CodRamo_NaoDefinido Then
        lObjstcPed.Cod_Ramo = lCodRamo
    End If
    
    'Número do pedido.
    If gfPreenchido(txtNumPedido) Then
        lObjstcPed.Num_Ped = txtNumPedido
    End If
    'Situação do documento.
    If lSitDocumento = Enum_SitDocumentoNaoTran Then       'Propostas não transmitidas.
        lObjstcPed.Cod_Sit_Ped = e_SitPedido_Nenhum
    Else
        If lSitDocumento = Enum_SitDocumentoTran Then        'Propostas transmitidas.
            lObjstcPed.Cod_Sit_Ped = e_SitPedido_Transmitido
        End If
    End If
    'Nome do segurado.
    If gfPreenchido(txtSegurado) Then
        lObjstcPed.Nom_Segurado = txtSegurado
    End If
    'Vigência.
    If gfPreenchido(txtIniVigencia) And txtIniVigencia <> "0" Then
        lObjstcPed.Dat_Ini_Vig = gfForData(1, txtIniVigencia)
    End If
    If gfPreenchido(txtTerVigencia) And txtTerVigencia <> "0" Then
        lObjstcPed.Dat_Ter_Vig = gfForData(1, txtTerVigencia)
    End If
    'Produtor.
    If gfPreenchido(txtProdutor) Then
        lObjstcPed.Cod_Prod = txtProdutor
    End If
    lObjstcPed.Num_Apol = 9999999999#
    lObjstcPed.NossoNumeroIn = lListaNossoNumero
    
    If gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        lObjstcPed.Cod_Segurado = gCodSeguradoYas
    End If
    
    '4. Busco os dados com os parâmetros de pesquisa informados.
    '###PROJETO INSPECAO - incluído o TRUE
    If Not lObjPed.mfCarregarColecao(lObjstcPed, False) Then
        lMensagem = lObjPed.UltimoErro
        Call gpGraLog(0, lObjPed.UltimoErro)
        Exit Sub
    End If
    
    '5. Se não encontrou registro, informa usuário.
    If lObjPed.Colecao.Count = 0 Then
        Call mpMenNaoEncontrado
        GoTo Finalizar
    End If
    
    '6. Monta grade de resultados.
    '6. Monta Recordsert  de resultados.
    
    Set lRetRecordSet = New ADODB.Recordset
    
    lRetRecordSet.Fields.Append "Excluir", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Ramo", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Pedido", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Segurado", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "CNPJ_CPF", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Inicio_Vigencia", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Termino_Vigencia", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Nosso_número", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Situacao_documento", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Expiracao", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "CodRamo", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Cod_Sit_Ped", adBSTR, , adFldIsNullable
    
    
    lRetRecordSet.CursorLocation = adUseClient
    lRetRecordSet.CursorType = adOpenDynamic
    lRetRecordSet.LockType = adLockOptimistic
    lRetRecordSet.Open
    
    '7. Coloca registros na grade.
    If lSitDocumento = Enum_SitDocumentoInspec Then
        For Each lObjstcPed In lObjPed.Colecao
             For Each lobjstcPedLoc In lObjstcPed.ColecaoPedLoc
                 If lobjstcPedLoc.Insp_Obr = True And lobjstcPedLoc.Insp_Agend = False And lObjstcPed.DAT_DIGIT > 0 Then
                   If DateDiff("D", gfForData(2, lObjstcPed.DAT_DIGIT), Date) <= 35 Then
                      lRetRecordSet.AddNew
                     'b) Coluna 0 - Excluir - checkBox para marcar itens a serem excluídos.
                     lRetRecordSet(1) = lObjstcPed.gDesc_Ramo
                     lRetRecordSet(2) = lObjstcPed.Num_Ped
                     lRetRecordSet(3) = lObjstcPed.Nom_Segurado
                     lRetRecordSet(4) = lObjstcPed.gCNPJ_CPF_Formatado
                     lRetRecordSet(5) = gfForData(2, lObjstcPed.Dat_Ini_Vig)
                     lRetRecordSet(6) = gfForData(2, lObjstcPed.Dat_Ter_Vig)
                     lRetRecordSet(7) = lObjstcPed.Nosso_Numero
                     lRetRecordSet(8) = lObjstcPed.Cod_Sit_Ped
                     lRetRecordSet(10) = lObjstcPed.Cod_Ramo
    '                'Verifico se está expirada.
                     lRetRecordSet(9) = "0"
                    If lObjstcPed.Dat_Pri_Calculo <> 0 And (gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
                        If (gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gfForData(1, Format(Date, "dd/mm/yyyy"))) > 15 And lObjstcPed.Cod_Produto >= 978) Or _
                        (gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gfForData(1, Format(Date, "dd/mm/yyyy"))) > 30 And lObjstcPed.Cod_Produto < 978) Then
                            'Cotação expirada
                            lRetRecordSet(9) = "1"
                        End If
                    End If
                    lRetRecordSet(11) = lObjstcPed.Cod_Sit_Ped
                    lRetRecordSet.Update
                  End If
                End If
            Next lobjstcPedLoc
        Next lObjstcPed
    Else
        For Each lObjstcPed In lObjPed.Colecao
            lRetRecordSet.AddNew
            'b) Coluna 0 - Excluir - checkBox para marcar itens a serem excluídos.
            'Set flxResultado.CellPicture = picUnchecked.Picture
            lRetRecordSet(1) = lObjstcPed.gDesc_Ramo
            lRetRecordSet(2) = lObjstcPed.Num_Ped
            lRetRecordSet(3) = lObjstcPed.Nom_Segurado
            lRetRecordSet(4) = lObjstcPed.gCNPJ_CPF_Formatado
            lRetRecordSet(5) = gfForData(2, lObjstcPed.Dat_Ini_Vig)
            lRetRecordSet(6) = gfForData(2, lObjstcPed.Dat_Ter_Vig)
            lRetRecordSet(7) = lObjstcPed.Nosso_Numero
            lRetRecordSet(8) = lObjstcPed.Cod_Sit_Ped
            lRetRecordSet(10) = lObjstcPed.Cod_Ramo
            'Verifico se está expirada.
            lRetRecordSet(9) = "0"
            If lObjstcPed.Dat_Pri_Calculo <> 0 And (gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
                If (gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gfForData(1, Format(Date, "dd/mm/yyyy"))) > 15 And lObjstcPed.Cod_Produto >= 978) Or _
                (gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gfForData(1, Format(Date, "dd/mm/yyyy"))) > 30 And lObjstcPed.Cod_Produto < 978) Then
                    'Cotação expirada
                    lRetRecordSet(9) = "1"
                End If
            End If
            lRetRecordSet(11) = lObjstcPed.Cod_Sit_Ped
            lRetRecordSet.Update
        Next lObjstcPed
    End If

   
    '9. Destruo os objetos.
    Set lObjPed = Nothing
    Set lObjstcPed = Nothing
    Set lObjPedCorr = Nothing
    Set lObjstcPedCorr = Nothing
    
    Exit Sub
Finalizar:
    Screen.MousePointer = vbDefault
    Set lObjPed = Nothing
    Set lObjstcPed = Nothing
    Set lObjPedCorr = Nothing
    Set lObjstcPedCorr = Nothing
    Exit Sub
ControleErro:
    lMensagem = Err & " " & Err.Description
    Call gpGraLog(0, "mpPesCotProposta ERRO - " & lMensagem, vbCritical)
    Exit Sub
Resume Next
End Sub


Private Sub Class_Initialize()
    Call gpInicia
End Sub
