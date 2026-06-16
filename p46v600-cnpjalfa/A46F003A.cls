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
    Enum_TipDocumentoCot = 1        'Cota��es e propostas.
    Enum_TipDocumentoApol = 2       'Apolice Vigentes
    Enum_TipDocumentoProt = 3       'Protocolos
End Enum

Public Enum Enum_SitDocumento
    Enum_SitDocumentoTodas = 0         'Todas
    Enum_SitDocumentoNaoTran = 1       'N�o Transmitidas
    Enum_SitDocumentoTran = 2          'Transmitidas
    Enum_SitDocumentoInspec = 3        'Inspecao n�o solicitaas
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
    Dim lDir            As String           'Utilizado com a fun��o DIR.
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

    '4. Obt�m registros de protocolo referente ao nosso n�mero.
    Set lobjTransm = New clsA46V718A
    Set lobjstcTransm = New stcA46V718B
    lobjstcTransm.Nosso_Numero = pNossoNumero
    
    If Not lobjTransm.mfCarregarColecao(lobjstcTransm) Then
        UltimoErro = lobjTransm.UltimoErro
        Exit Function
    End If
    If lobjTransm.Colecao.Count = 0 Then
        UltimoErro = "FNL0014 - Registro de protocolo de transmiss�o n�o localizado (pedido " & _
                         gNumPedido & " - nosso n�mero " & pNossoNumero & ")."
        Set lobjTransm = Nothing
        Set lobjstcTransm = Nothing
        Exit Function
    End If
    
    'c) Obt�m dados da proposta transmitida (detalhes - erro).
    '*************************Ver este ponto

    '5. Processa registro de protocolo.
    lPriProposta = True
    '6. Inicia processo.
    mQuaLinhas = 0
    Call mpPosLinha(lProtocolo, mQuaLinhas, "YASUDA - PROTOCOLO DE TRANSMISS�O - " & Format$(Now, "dd/mm/yyyy hh:nn:ss") & ".")
    Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
    Call mpPosLinha(lProtocolo, mQuaLinhas, "Protocolo de transmiss�o - SYAS Multiseguros.")
    Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
    For Each lobjstcTransm In lobjTransm.Colecao
        'a) Exibe dados
        'Call mpPosLinha(lProtocolo, mQuaLinhas,"Transmiss�o da Yasuda: " & lobjstcTransm.Nom_Arq_Retorno & ".")
        If lNosNumero <> lobjstcTransm.Nosso_Numero Then
            lNosNumero = lobjstcTransm.Nosso_Numero
            If gfPreenchido(lobjstcTransm.Nosso_Numero) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Proposta recebida....." & lobjstcTransm.Nosso_Numero)
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Proposta recebida.....n�o informado")
            End If
            If gfPreenchido(lobjstcTransm.Num_Ped) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Pedido................" & lobjstcTransm.Num_Ped)
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Pedido................n�o informado")
            End If
            If gfPreenchido(lobjstcTransm.Cod_Ramo) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Ramo.................." & lobjstcTransm.Cod_Ramo)
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Ramo..................n�o informado")
            End If
            If gfPreenchido(lobjstcTransm.Nom_Segurado) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Segurado.............." & Trim$(lobjstcTransm.Nom_Segurado))
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Segurado..............n�o informado.")
            End If
            If gfPreenchido(lobjstcTransm.Dat_Ini_Vig) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "In�cio de vig�ncia...." & gfForData(2, lobjstcTransm.Dat_Ini_Vig))
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "In�cio de vig�ncia....n�o informado")
            End If
            If gfPreenchido(lobjstcTransm.Dat_Ter_Vig) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "T�rmino de vig�ncia..." & gfForData(2, lobjstcTransm.Dat_Ter_Vig))
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "T�rmino de vig�ncia...n�o informado")
            End If
            If gfPreenchido(lobjstcTransm.Nom_Arq_Envio) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Arquivo enviado......." & lobjstcTransm.Nom_Arq_Envio)
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Arquivo enviado.......n�o informado")
            End If
            If gfPreenchido(lobjstcTransm.Dat_Hor_Envio) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Data/Hora envio......." & gfForData(4, lobjstcTransm.Dat_Hor_Envio) & " horas.")
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Data/Hora envio.......n�o informado")
            End If
            If gfPreenchido(lobjstcTransm.Nom_Arq_Retorno) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Arquivo de retorno...." & lobjstcTransm.Nom_Arq_Retorno)
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Arquivo de retorno....n�o informado")
            End If
            If gfPreenchido(lobjstcTransm.Dat_Hor_Retorno) Then
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Data/Hora retorno....." & gfForData(4, lobjstcTransm.Dat_Hor_Envio) & " horas.")
            Else
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Data/Hora retorno.....n�o informado")
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
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Os anexos, quando necess�rios, poder�o ser enviados das seguintes formas:")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "1) VIA FAX: Atrav�s  do  n�mero  (0xx11) 3305-5664,  informando  o  n�mero  do ")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "protocolo interno " & lobjstcTransm.Num_PI & lDigito & " no atendimento eletr�nico ou;")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "2) VIA YasudaNet, '�REA DO CORRETOR': Carregar o(s) documento(s) digitalizado(s)")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "atrav�s  dos  links  Seguro  Demais Ramos  -->  Emiss�o.  Filtrar pelo N�mero do")
                Call mpPosLinha(lProtocolo, mQuaLinhas, "Protocolo (PI).")
                Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
            End If
            Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
            Select Case lobjstcTransm.Ind_Proposta
                Case "0"
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "STATUS................Proposta n�o transmitida.")
                Case 1
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "STATUS...Proposta eletr�nica recebida pela Yasuda em " & _
                                    gfForData(4, lobjstcTransm.Dat_Hor_Retorno) & " horas.")
                    
                Case 2
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "STATUS................Proposta apresentou erro.")
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "� necess�rio retransmitir esta proposta.")
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "Descri��o do erro....." & lobjstcTransm.Tip_Emissao) 'Cont�m observa��o do erro
                Case Else
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "STATUS................Proposta apresentou erro.")
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "� necess�rio retransmitir esta proposta.")
                    Call mpPosLinha(lProtocolo, mQuaLinhas, "Descri��o do erro.....Indicador de proposta desconhecido - comunicar Central de Atendimento.")
            End Select
            Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Itens da proposta:")
         End If
        '***==== DADOS ITENS
        Call mpPosLinha(lProtocolo, mQuaLinhas, " ")
        If gfPreenchido(lobjstcTransm.Num_Item) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Item.................." & lobjstcTransm.Num_Item)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Item..................n�o informado")
        End If
        Call mpPosLinha(lProtocolo, mQuaLinhas, "Logradouro............" & lobjstcTransm.Cod_Tip_Logr_Risco & _
                            " " & lobjstcTransm.Nom_Logr_Risco & _
                            ", " & lobjstcTransm.Num_Risco)
        
        If gfPreenchido(lobjstcTransm.Dsc_Compl_Risco) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Complemento..........." & lobjstcTransm.Dsc_Compl_Risco)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Complemento...........n�o informado")
        End If
        If gfPreenchido(lobjstcTransm.Num_Cep_Risco) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "CEP..................." & lobjstcTransm.Num_Cep_Risco)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "CEP...................n�o informado")
        End If
        If gfPreenchido(lobjstcTransm.Nom_Cid_Risco) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Cidade................" & lobjstcTransm.Nom_Cid_Risco)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "Cidade................n�o informada")
        End If
        If gfPreenchido(lobjstcTransm.Sig_UF_Risco) Then
            Call mpPosLinha(lProtocolo, mQuaLinhas, "UF...................." & lobjstcTransm.Sig_UF_Risco)
        Else
            Call mpPosLinha(lProtocolo, mQuaLinhas, "UF....................n�o informada")
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
    Dim lTemNumPedido   As Boolean  'Indica que tem n�mero do pedido para pesquisa.

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
    
    '2. Consiste crit�rios de pesquisa.
    'a) Tipo de documento.
    If lTipDocumento <> Enum_TipDocumentoCot And lTipDocumento <> Enum_TipDocumentoApol And lTipDocumento <> Enum_TipDocumentoProt Then
        lMensagem = "FNP0013 - Tipo de documento n�o selecionado."
        Call gpGraLog(0, lMensagem)
        GoTo Fim
    End If
    
    'b) N�mero do pedido preenchido.
    lTemNumPedido = False
    If gfPreenchido(txtNumPedido) Then
        If Not IsNumeric(txtNumPedido) Then
            lMensagem = "FGE0080 - N�mero do pedido n�o num�rico."
            Call gpGraLog(0, lMensagem)
            GoTo Fim
        End If
        lTemNumPedido = True
    End If

        
    'c) Ramo
    If gTipUsuario = e_Tipo_Usu_Yasuda Then
        If gfPreenchido(txtSegurado) = False And gfPreenchido(txtNumPedido) = False Then
            If lCodRamo = 0 Then
                lMensagem = "FRA0001 - Ramo n�o selecionado."
                Call gpGraLog(0, lMensagem)
                GoTo Fim
            End If
        End If
    End If

    'd) N�mero do pedido n�o preenchido.
    If lTemNumPedido = False Then
        '=== Vig�ncia.
        If gfPreenchido(txtIniVigencia) Then
            If gfVerData(2, txtIniVigencia) = False Then
                lMensagem = "FGE0018 - Vig�ncia inv�lida (data de in�cio de vig�ncia inv�lida)."
                Call gpGraLog(0, lMensagem)
                GoTo Fim
            End If
        End If
        If gfPreenchido(txtTerVigencia) Then
            If gfVerData(2, txtTerVigencia) = False Then
                lMensagem = "FGE0081 - Vig�ncia inv�lida (data de t�rmino de vig�ncia inv�lida)."
                Call gpGraLog(0, lMensagem)
                GoTo Fim
            End If
        End If
        If gfPreenchido(txtIniVigencia) And gfPreenchido(txtTerVigencia) Then
            If gfForData(1, txtIniVigencia) > gfForData(1, txtTerVigencia) Then
                lMensagem = "FGE0082 - Vig�ncia inv�lida (data de in�cio de vig�ncia " & _
                                 txtIniVigencia & ") maior que a data de t�rmino de vig�ncia (" & _
                                 txtTerVigencia & ")."
                Call gpGraLog(0, lMensagem)
                GoTo Fim
            End If
        End If
    End If

    '3. Obt�m dados.
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
            lMensagem = "FNP0013 - Tipo de documento n�o selecionado."
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
    'Procedure: pesquisa protocolo de transmiss�o.

    Dim lData           As String           'Campo auxiliar: data.
    Dim lDir            As String           'Utilizado com a fun��o DIR.
    Dim lobjTransm      As clsA46V718A      'Objeto da classe clsA46V718A (TAB_PED_TRANSM)
    Dim lobjstcTransm   As stcA46V718B      'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM)

    On Error GoTo ControleErro
    
    '1. Instancio os objetos
    Set lobjTransm = New clsA46V718A
    Set lobjstcTransm = New stcA46V718B
    

    '2. Carrego os par�metro de pesquisa.
    'Obs.: Se n�mero do pedido est� preenchido, utiliza apenas esta informa��o no SELECT.
    If gfPreenchido(txtNumPedido) Then
        lobjstcTransm.Num_Ped = Format$(txtNumPedido, "000000")
    Else
        If gfPreenchido(txtSegurado) Then
            lobjstcTransm.Nom_Segurado = gfForAspas(UCase(txtSegurado))
        End If
        'b) Vig�ncia.
        If gfPreenchido(txtIniVigencia) Then
            lData = gfForData(1, txtIniVigencia)
            lobjstcTransm.Dat_Ini_Vig = lData
        End If
        If gfPreenchido(txtTerVigencia) Then
            lData = gfForData(1, txtTerVigencia)
            lobjstcTransm.Dat_Ter_Vig = lData
        End If
    End If

    '3. Obt�m registros de protocolo.
    If Not lobjTransm.mfCarregarColecao(lobjstcTransm) Then
        UltimoErro = lobjTransm.UltimoErro
        Call gpGraLog(0, UltimoErro)
        Exit Sub
    End If

    '4. Se n�o encontrou registro, informa usu�rio.
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
        '===2) N�mero do pedido.
        lRetRecordSet(1) = Format$(lobjstcTransm.Num_Ped, "000000")
        '===3) Nome do segurado.
        If gfPreenchido(lobjstcTransm.Nom_Segurado) Then
             lRetRecordSet(2) = lobjstcTransm.Nom_Segurado
         Else
             lRetRecordSet(2) = ""
         End If
         '===4) In�cio de vig�ncia.
         If gfPreenchido(lobjstcTransm.Dat_Ini_Vig) Then
             lRetRecordSet(3) = gfForData(2, lobjstcTransm.Dat_Ini_Vig)
         Else
             lRetRecordSet(3) = ""
         End If
         '===5) T�rmino de vig�ncia.
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
         '===8) N�mero logradouro.
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
         '===14) Nosso n�mero.
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

    lMensagem = "FNL0011 - N�o localizado registro que satisfa�a os crit�rios selecionados." & vbLf
    If gfPreenchido(txtNumPedido) Then
        lMensagem = lMensagem & vbLf & "N�mero do pedido = " & txtNumPedido & "."
    Else
        If lTipDocumento = Enum_TipDocumentoCot Then
            lMensagem = lMensagem & vbLf & "Tipo de documento = cota��es e propostas"
            If lSitDocumento = Enum_SitDocumentoTodas Then
                lMensagem = lMensagem & "."
            Else
                If lSitDocumento = Enum_SitDocumentoNaoTran Then
                    lMensagem = lMensagem & " (apenas propostas n�o transmitidas)."
                Else
                    If lSitDocumento = Enum_SitDocumentoTran Then
                       lMensagem = lMensagem & " (apenas propostas transmitidas)."
                    Else
                       lMensagem = lMensagem & " (apenas cota��es sem solicita��o de agendamento)."
                    End If
                End If
            End If
        Else
            If lTipDocumento = Enum_TipDocumentoApol Then
                lMensagem = lMensagem & vbLf & "Tipo de documento = ap�lices vigentes."
            Else
                lMensagem = lMensagem & vbLf & "Tipo de documento = protocolo de transmiss�o."
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
            lMensagem = lMensagem & vbLf & "Vig�ncia: de " & txtIniVigencia & " a " & _
                        txtTerVigencia & "."
        Else
            If gfPreenchido(txtIniVigencia) Then
                lMensagem = lMensagem & vbLf & "Vig�ncia em " & txtIniVigencia & "."
            Else
                If gfPreenchido(txtTerVigencia) Then
                    lMensagem = lMensagem & vbLf & "Vig�ncia em " & txtTerVigencia & "."
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
    'Procedure: pesquisa cota��es e propostas.
    Dim lListaNossoNumero   As String           'Lista de nosso_numero.
    Dim lCount              As Double           'Contador.
    Dim lrst                As ADODB.Recordset  'Registro de ap�lice.
    Dim bdP00APOL           As ADODB.Connection 'Arquivo de ap�lice (Conex�o).
    Dim lCpfCnpj            As String           'Retorno da fun��o de forma��o de cpf e cnpj.
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
    'Busco os dados da ap�lice
    
    'Ficha 1000782 - N�o permitir Renova��o LMI para Corretores
    gSelect = "SELECT distinct " & lNomeTab & ".* , " & lNomeTab & "_LOC.NOM_LOGR_RISCO "
    gSelect = gSelect & " FROM (" & lNomeTab & " inner join " & lNomeTab & "_Corr"
    gSelect = gSelect & " on " & lNomeTab & ".Num_APol = " & lNomeTab & "_Corr.NUm_Apol and"
    'Ficha 1000782 - N�o permitir Renova��o LMI para Corretores
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
            
    'In�cio da vig�ncia.
    If gfPreenchido(txtIniVigencia) Then
        gSelect = gfWhereAnd(gSelect) & " Dat_Ini_Vig >= " & gfForData(1, txtIniVigencia)
    End If
    
    'T�rmino da vig�ncia.
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
    
    'Ficha 1000782 - N�o permitir Renova��o LMI para Corretores
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
    
    '5. Se n�o encontrou registro, informa usu�rio.
    If lrst.EOF Then
        Call mpMenNaoEncontrado
        Exit Sub
    End If
    
    '6. Monta grade de resultados.
    
    lRetRecordSet.Fields.Append "Excluir", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Ramo", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Ap�lice", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Segurado", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "CNPJ/CPF", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "In�cio Vig�ncia", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "T�rmino Vig�ncia", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Nosso n�mero", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Situa��o do documento", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "Produto/Modalidade", adBSTR, , adFldIsNullable
    lRetRecordSet.Fields.Append "CodRamo", adBSTR, , adFldIsNullable
    
    '7. Coloca registros na grade.
    Do While Not lrst.EOF
         'b) Coluna 0 - Excluir - checkBox para marcar itens a serem exclu�dos.
         If lrst("Cod_Ramo") = e_CodRamo_Condominio Then
             lRetRecordSet(1) = "Condom�nio"
         ElseIf lrst("Cod_Ramo") = e_CodRamo_Empresarial Then
             lRetRecordSet(1) = "Empresarial"
         ElseIf lrst("Cod_Ramo") = e_CodRamo_Residencial Then
             lRetRecordSet(1) = "Residencial"
         Else
             lRetRecordSet(1) = "N�o definido"
         End If
         lRetRecordSet(2) = lrst("Num_Apol")
         lRetRecordSet(3) = lrst("Nom_Segurado")
         Select Case Len(CStr(lrst("Num_Cgc_Cpf")))
             Case Is > 11
                 If IsNumeric(lrst("Num_Cgc_Cpf")) Then
                     Call gpForCNPJ(1, Format$(lrst("Num_Cgc_Cpf"), "00000000000000"), lCpfCnpj)
                 Else
                     Call gpForCNPJ(1, Right$(String(14, "0") & Trim(lrst("Num_Cgc_Cpf")), 14), lCpfCnpj)
                 End If
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
    'Procedure: pesquisa cota��es e propostas.
    
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
    
    'N�mero do pedido.
    If gfPreenchido(txtNumPedido) Then
        lObjstcPed.Num_Ped = txtNumPedido
    End If
    'Situa��o do documento.
    If lSitDocumento = Enum_SitDocumentoNaoTran Then       'Propostas n�o transmitidas.
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
    'Vig�ncia.
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
    
    '4. Busco os dados com os par�metros de pesquisa informados.
    '###PROJETO INSPECAO - inclu�do o TRUE
    If Not lObjPed.mfCarregarColecao(lObjstcPed, False) Then
        lMensagem = lObjPed.UltimoErro
        Call gpGraLog(0, lObjPed.UltimoErro)
        Exit Sub
    End If
    
    '5. Se n�o encontrou registro, informa usu�rio.
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
    lRetRecordSet.Fields.Append "Nosso_n�mero", adBSTR, , adFldIsNullable
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
                     'b) Coluna 0 - Excluir - checkBox para marcar itens a serem exclu�dos.
                     lRetRecordSet(1) = lObjstcPed.gDesc_Ramo
                     lRetRecordSet(2) = lObjstcPed.Num_Ped
                     lRetRecordSet(3) = lObjstcPed.Nom_Segurado
                     lRetRecordSet(4) = lObjstcPed.gCNPJ_CPF_Formatado
                     lRetRecordSet(5) = gfForData(2, lObjstcPed.Dat_Ini_Vig)
                     lRetRecordSet(6) = gfForData(2, lObjstcPed.Dat_Ter_Vig)
                     lRetRecordSet(7) = lObjstcPed.Nosso_Numero
                     lRetRecordSet(8) = lObjstcPed.Cod_Sit_Ped
                     lRetRecordSet(10) = lObjstcPed.Cod_Ramo
    '                'Verifico se est� expirada.
                     lRetRecordSet(9) = "0"
                    If lObjstcPed.Dat_Pri_Calculo <> 0 And (gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
                        If (gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gfForData(1, Format(Date, "dd/mm/yyyy"))) > 15 And lObjstcPed.Cod_Produto >= 978) Or _
                        (gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gfForData(1, Format(Date, "dd/mm/yyyy"))) > 30 And lObjstcPed.Cod_Produto < 978) Then
                            'Cota��o expirada
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
            'b) Coluna 0 - Excluir - checkBox para marcar itens a serem exclu�dos.
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
            'Verifico se est� expirada.
            lRetRecordSet(9) = "0"
            If lObjstcPed.Dat_Pri_Calculo <> 0 And (gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
                If (gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gfForData(1, Format(Date, "dd/mm/yyyy"))) > 15 And lObjstcPed.Cod_Produto >= 978) Or _
                (gfDifDatas(lObjstcPed.Dat_Pri_Calculo, gfForData(1, Format(Date, "dd/mm/yyyy"))) > 30 And lObjstcPed.Cod_Produto < 978) Then
                    'Cota��o expirada
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
